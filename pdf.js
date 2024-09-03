import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import { extractTableData, extractTableHeaders } from './utils.js';

export function exportToPDF() {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({
    orientation: 'landscape'
  });

  const newLineHeightFactor = 0.5;
  doc.setLineHeightFactor(newLineHeightFactor);

  // Obtém o título principal (h2) e o subtítulo (h3) diretamente do HTML
  const mainTitle = document.querySelector('.col-sm-9 h2').textContent.trim();
  const subtitle = document.querySelector('.col-sm-9 h3').textContent.trim();

  // Seleciona todas as tabelas que você quer incluir no PDF
  const tables = document.querySelectorAll('.table.table-responsive.table-striped.table-bordered.table-sm');

  let finalY = 30; // Posição inicial Y

  // Carrega a imagem da logo
  var logoImg = new Image();
  logoImg.src = './logo.png'; // Caminho da imagem na pasta public

  tables.forEach((table, index) => {
    // Adiciona o título da tabela ao PDF
    let titleHeight = doc.getTextDimensions(mainTitle).h; 
    let titleX = doc.internal.pageSize.width / 2;
    let titleY = finalY + titleHeight + 5;
    doc.setFontSize(16);
    doc.text(mainTitle, titleX, titleY, { align: 'center' });
    doc.setFontSize(14);
    doc.text(subtitle, titleX, titleY + 8, { align: 'center' });

    // Calcula a posição da logo à direita do título
    const logoWidth = 30; 
    const logoHeight = 10; 
    const logoX = titleX + doc.getTextDimensions(mainTitle).w / 2 + 5;
    const logoY = titleY - titleHeight / 2 - logoHeight / 2; 

    // Adiciona a logo ao PDF
    doc.addImage(logoImg, 'PNG', 150, 10, 50, 15);

    finalY = titleY + 10; 

    // Adiciona a tabela ao PDF
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
