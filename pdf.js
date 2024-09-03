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

 while (titleElement && (titleElement.tagName !== 'H3' && titleElement.tagName !== 'H2')) {
      titleElement = titleElement.previousElementSibling;
    }

    // Define um título padrão se o título não for encontrado
    let titleText = 'Tabela ' + (index + 1); // Título padrão
    if (titleElement) {
      titleText = titleElement.textContent.trim(); 
    }

    let titleHeight = doc.getTextDimensions(titleText).h;
    let titleX = doc.internal.pageSize.width / 2;
    let titleY = finalY + titleHeight + 5;
    doc.setFontSize(10);
    doc.text(titleText, titleX, titleY, {
      align: 'center',
    });
    finalY = titleY; // Atualiza finalY para incluir a altura do título
    }

  doc.autoTable({
  html: table,
  startY: finalY + 10,
  theme: 'grid',
  headStyles: {
    fillColor: [200, 230, 255], 
    textColor: [0, 0, 0]
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

  for (let i = 1; i <= pageCount; i++) {
    doc.setPage(i);
    const pageWidth = doc.internal.pageSize.width;
    const pageHeight = doc.internal.pageSize.height;
    const str = 'Página ' + i + ' de ' + pageCount;
    doc.setFontSize(6);
    doc.text(str, pageWidth - 10, pageHeight - 3, {
      align: 'right'
    });
  }
}
