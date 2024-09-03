import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import { extractTableData, extractTableHeaders } from './utils.js';

export function exportToPDF() {
  const doc = new jsPDF({
    orientation: 'landscape'
  });

  const newLineHeightFactor = 0.5;
  doc.setLineHeightFactor(newLineHeightFactor);

  // Seleciona todas as tabelas que você quer incluir no PDF
  const tables = document.querySelectorAll('.table.table-responsive.table-striped.table-bordered.table-sm');

  let finalY = 30; // Posição inicial Y

  tables.forEach((table, index) => {
    let titleElement = table.previousElementSibling; // Assume que o título é o elemento anterior à tabela

    // Verifica se o elemento anterior é um título (h2 ou h3)
    while (titleElement && (titleElement.tagName !== 'H3' && titleElement.tagName !== 'H2')) {
      titleElement = titleElement.previousElementSibling;
    }

    if (titleElement) {
      let titleHeight = doc.getTextDimensions(titleElement.textContent).h;
      let titleX = doc.internal.pageSize.width / 2;
      let titleY = finalY + titleHeight + 5;
      doc.setFontSize(10);
      doc.text(titleElement.textContent, titleX, titleY, {
        align: 'center',
      });
      finalY = titleY; // Atualiza finalY para incluir a altura do título
    }

    doc.autoTable({
      html: table,
      startY: finalY + 10,
      theme: 'grid',
      headStyles: {
        fillColor: [200, 230, 255]
      },
      bodyStyles: {
        fillColor: false
      },
      alternateRowStyles: {
        fillColor: false
      },
      pageBreak: 'auto',
      rowPageBreak: 'noWrap',
      overFlow: "linebreak",
      showHead: 'everyPage',
      margin: {
        top: 30,
        right: 10,
        bottom: 5,
        left: 15
      },
      styles: {
        fontSize: 7,
        fontStyle: 'normal',
        fontFamily: 'Calibri, sans-serif'
      },
      columnStyles: {
        0: {
          halign: 'left'
        },
        1: {},
        2: {},
        3: {},
        4: {},
        5: {},
        6: {},
        7: {},
        8: {},
        9: {},
        10: {},
        11: {},
        12: {},
        13: {}
      },

      didParseCell: function (data) {
        if (data.row.index === 0 && data.section === 'body') {
          data.cell.styles.fillColor = [200, 230, 255];
          data.cell.styles.textColor = [0, 0, 0];
        }
        if (data.section === 'body' && data.column.index === 0) {
          data.cell.styles.halign = 'left';
        }
        if (data.section === 'body' && data.column.index !== 0) {
          data.cell.styles.halign = 'center';
        }
      }
    });

    finalY = doc.lastAutoTable.finalY || finalY;
  });

  addDynamicTextAndPageNumbers(doc);
  doc.save(document.querySelector('h3').textContent + '.pdf');
}

function addDynamicTextAndPageNumbers(doc) {
  const pageCount = doc.internal.getNumberOfPages();
  const dynamicTexts = [
    document.querySelector('h2').textContent.trim() + " - Início do Projeto: " + document.querySelectorAll('.col-sm-3 p')[1]?.textContent,
    document.querySelector('h3').textContent.trim() + " " + document.querySelector('.col-sm-6 p')?.textContent,
  ];

  for (let i = 1; i <= pageCount; i++) {
    doc.setPage(i);
    let yPos = 10;
    dynamicTexts.forEach(text => {
      doc.text(text, 155, yPos, {
        align: 'center',
        fontSize: 10
      });
      yPos += 7;
    });

    const pageWidth = doc.internal.pageSize.width;
    const pageHeight = doc.internal.pageSize.height;
    const str = 'Página ' + i + ' de ' + pageCount;
    doc.setFontSize(6);
    doc.text(str, pageWidth - 10, pageHeight - 3, {
      align: 'right'
    });
  }
}
