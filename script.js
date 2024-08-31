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
  const wb = XLSX.utils.book_new(); // Cria uma nova pasta de trabalho Excel
  const ws = XLSX.utils.aoa_to_sheet([]); // Cria uma planilha vazia

  // Seleciona todas as tabelas que você quer incluir no Excel
  const tables = document.querySelectorAll('.table.table-responsive.table-striped.table-bordered.table-sm');

  let currentRow = 0; 

  tables.forEach((table, index) => {
    const tableData = extractTableData(table);
    const tableHeaders = extractTableHeaders(table);

    XLSX.utils.sheet_add_aoa(ws, [[`Tabela ${index + 1}`]], { origin: { r: currentRow, c: 0 } });
    currentRow++;

    XLSX.utils.sheet_add_aoa(ws, [tableHeaders], { origin: { r: currentRow, c: 0 } });
    currentRow++;

    tableData.forEach(rowData => {
      XLSX.utils.sheet_add_aoa(ws, [rowData], { origin: { r: currentRow, c: 0 } });
      currentRow++;
    });

    currentRow++; 
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
    let startY = 20;

    doc.setFontSize(16);
    doc.text("HOSPITAL REGIONAL DA COSTA LESTE MAGID THOMÉ", 10, 10);
    doc.setFontSize(14);
    doc.text("RL06 - CUSTO SERVIÇOS AUXILIARES", 10, 16);

    const tables = document.querySelectorAll('.table.table-responsive.table-striped.table-bordered.table-sm');

    tables.forEach((table, index) => {
      const tableData = extractTableData(table);
      const tableHeaders = extractTableHeaders(table);

      doc.setFontSize(12);
      doc.text(`Tabela ${index + 1}`, 10, startY);

      autoTable(doc, {
        head: [tableHeaders],
        body: tableData,
        startY: startY + 10,
        styles: { halign: 'right' },
      });

      startY = doc.lastAutoTable.finalY + 10;
    });

    doc.setFontSize(9);
    doc.setTextColor(0, 102, 0);
    doc.text("Sistemas integrados www.cjpnet.com.br", 10, doc.internal.pageSize.getHeight() - 10);

    doc.save('RelatorioCompleto.pdf');
  }
});