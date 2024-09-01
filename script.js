import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';

// Função para extrair dados da tabela
function extractTableData(table) {
  const tableData = [];
  for (const row of table.querySelectorAll('tbody tr')) {
    const rowData = [];
    for (const cell of row.querySelectorAll('td')) {
      rowData.push(cell.innerHTML.trim());
    }
    tableData.push(rowData);
  }
  return tableData;
}

// Função para extrair cabeçalhos da tabela
function extractTableHeaders(table) {
  const headers = [];
  for (const headerCell of table.querySelectorAll('thead th')) {
    headers.push(headerCell.textContent.trim());
  }
  return headers;
}

// Função para gerar o arquivo Excel
function exportToExcel() {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([]);

  const tables = document.querySelectorAll('.table.table-responsive.table-striped.table-bordered.table-sm');

  let currentRow = 0;

  // Adiciona o título principal à planilha
  const mainTitle = document.querySelector('h2').textContent.trim();
  const subtitle = document.querySelector('h3').textContent.trim();
  XLSX.utils.sheet_add_aoa(ws, [[mainTitle]], { origin: { r: currentRow, c: 0 } });
  currentRow++;
  XLSX.utils.sheet_add_aoa(ws, [[subtitle]], { origin: { r: currentRow, c: 0 } });
  currentRow += 2; // Pula uma linha extra

  tables.forEach((table) => {
    const tableData = extractTableData(table);
    const tableHeaders = extractTableHeaders(table);

    // Adiciona os cabeçalhos da tabela
    XLSX.utils.sheet_add_aoa(ws, [tableHeaders], { origin: { r: currentRow, c: 0 } });
    currentRow++;

    // Adiciona os dados da tabela
    tableData.forEach(rowData => {
      XLSX.utils.sheet_add_aoa(ws, [rowData], { origin: { r: currentRow, c: 0 } });
      currentRow++;
    });

    currentRow += 2; // Adiciona duas linhas em branco entre as tabelas
  });

  XLSX.utils.book_append_sheet(wb, ws, 'RelatorioCompleto'); 

  const excelFile = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([excelFile], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  saveAs(blob, 'RelatorioCompleto.xlsx');
}

// Aguarda o carregamento do DOM
document.addEventListener('DOMContentLoaded', function () {
  const printButton = document.querySelector('.float-right.btn.btn-outline-primary');

  printButton.addEventListener('click', function (event) {
    event.preventDefault();
    $('#formatoModal').modal('show');
  });

  $('#btn-excel').click(function () {
    exportToExcel();
    $('#formatoModal').modal('hide');
  });

  $('#btn-pdf').click(function () {
    exportToPDF();
    $('#formatoModal').modal('hide');
  });

  // Função para gerar o arquivo PDF
  function exportToPDF() {
    const doc = new jsPDF();
    let startY = 30; 

    const mainTitle = document.querySelector('h2').textContent.trim();
    const subtitle = document.querySelector('h3').textContent.trim();
    doc.setFontSize(16);
    doc.text(mainTitle, 10, 10);
    doc.setFontSize(14);
    doc.text(subtitle, 10, 18);

    /*
    const logoImg = new Image();
    logoImg.src = 'logo.png';
    logoImg.onload = function() {
      const logoWidth = 50; 
      const logoHeight = (logoImg.height * logoWidth) / logoImg.width;
      const logoX = doc.internal.pageSize.getWidth() - logoWidth - 10;
      const logoY = 10; 
      doc.addImage(logoImg, 'PNG', logoX, logoY, logoWidth, logoHeight);
    */ 

    addTablesToPDF(doc, startY);

    doc.save('RelatorioCompleto.pdf');
  }

  function addTablesToPDF(doc, startY) {
    const tables = document.querySelectorAll('.table.table-responsive.table-striped.table-bordered.table-sm');

    tables.forEach((table) => {
      const tableData = extractTableData(table);
      const tableHeaders = extractTableHeaders(table);

      startY += 5; 

      autoTable(doc, {
        head: [tableHeaders],
        body: tableData,
        startY: startY,
        styles: { halign: 'right' },
        didDrawPage: function (data) {
          if (data.pageNumber > 1) {
            doc.setFontSize(10);
            doc.text("HOSPITAL REGIONAL DA COSTA LESTE MAGID THOMÉ", 10, 10);
            doc.text("RL06 - CUSTO SERVIÇOS AUXILIARES", 10, 25);
          }
        },
        // Correção para o erro "Cannot read properties of undefined (reading 'includes')"
        didParseCell: function (data) {
          if (data.cell.raw && data.cell.raw.includes('<strong>')) { 
            data.cell.styles.fontStyle = 'bold';
            data.cell.text = data.cell.raw.replace(/<strong>/g, '').replace(/<\/strong>/g, '');
          }
        } 
      });

      startY = doc.lastAutoTable.finalY; 
    });

    doc.setFontSize(9);
    doc.setTextColor(0, 102, 0);
    doc.text("Sistemas integrados www.cjpnet.com.br", 10, doc.internal.pageSize.getHeight() - 10);
  }
});