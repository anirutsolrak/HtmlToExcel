import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import { extractTableData, extractTableHeaders } from './utils.js'; // Importa as funções

// Função para gerar o arquivo PDF
export function exportToPDF() {
  const doc = new jsPDF();
  let startY = 30; // Posição inicial Y para a primeira tabela

  // Adiciona o título principal ao PDF
  const mainTitle = document.querySelector('h2').textContent.trim();
  const subtitle = document.querySelector('h3').textContent.trim();
  doc.setFontSize(16);
  doc.text(mainTitle, 10, 10);
  doc.setFontSize(14);
  doc.text(subtitle, 10, 18);

  // Adiciona as tabelas ao PDF usando autoTable
  addTablesToPDF(doc, startY);

  doc.save('RelatorioCompleto.pdf');
}

// Função para adicionar todas as tabelas ao PDF
function addTablesToPDF(doc, startY) {
  const tables = document.querySelectorAll('.table.table-responsive.table-striped.table-bordered.table-sm');

  tables.forEach((table) => {
    const tableData = extractTableData(table);
    const tableHeaders = extractTableHeaders(table);

    startY += 10;

    autoTable(doc, {
      head: [tableHeaders],
      body: tableData,
      startY: startY,
      styles: { halign: 'right' },
      didDrawPage: function (data) {
        if (data.pageNumber > 1) {
          doc.setFontSize(10);
          doc.text("HOSPITAL REGIONAL DA COSTA LESTE MAGID THOMÉ", 10, 10);
          doc.text("RL06 - CUSTO SERVIÇOS AUXILIARES", 10, 15);
        }
      },
      didParseCell: function (data) {
        if (data.cell.raw && data.cell.raw.includes('<strong>'))  {
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

  // Adiciona a numeração de páginas ao PDF
  const pageCount = doc.internal.getNumberOfPages();
  for (let i = 1; i <= pageCount; i++) {
    doc.setPage(i);
    const pageWidth = doc.internal.pageSize.width;
    const pageHeight = doc.internal.pageSize.height;
    doc.setFontSize(8); 
    doc.text(`Página ${i} de ${pageCount}`, pageWidth - 10, pageHeight - 5, { align: 'right' });
  }
}
