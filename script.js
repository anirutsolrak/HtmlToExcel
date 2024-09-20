// Certifique-se de que as bibliotecas estejam incluídas no seu HTML antes deste script:
// <script src="https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js"></script>
// <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
// <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.5.28/jspdf.plugin.autotable.min.js"></script>

// Funções auxiliares para extrair dados da tabela
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

function extractTableHeaders(table) {
  const headers = [];
  for (const headerCell of table.querySelectorAll('thead th')) {
    headers.push(headerCell.textContent.trim());
  }
  return headers;
}

// ... (outras funções) ...

// Função para exportar para Excel (CORRIGIDA)
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
  currentRow += 2;

  tables.forEach((table, tableIndex) => {
    const tableData = extractTableData(table);

    // Obtém o elemento <thead> da tabela
    const tableHeader = table.querySelector('thead');

    // Adiciona o cabeçalho da tabela ao Excel
    const headerRows = tableHeader.querySelectorAll('tr');
    headerRows.forEach(headerRow => {
      const rowData = [];
      const headerCells = headerRow.querySelectorAll('th');
      let colIndex = 0;
      headerCells.forEach(headerCell => {
        const colspan = parseInt(headerCell.getAttribute('colspan')) || 1;
        rowData[colIndex] = headerCell.textContent.trim();
        colIndex += colspan;
      });
      XLSX.utils.sheet_add_aoa(ws, [rowData], { origin: { r: currentRow, c: 0 } });
      currentRow++;
    });

    // Adiciona os dados da tabela com formatação de negrito
    tableData.forEach(rowData => {
      const formattedRowData = rowData.map((cellData, cellIndex) => {
        if (cellData.includes('<strong>')) {
          const cellValue = cellData.replace(/<strong>/g, '').replace(/<\/strong>/g, '');
          return { v: cellValue, s: { font: { bold: true } } };
        } else {
          return cellData;
        }
      });
      XLSX.utils.sheet_add_aoa(ws, [formattedRowData], { origin: { r: currentRow, c: 0 } });
      currentRow++;
    });

    currentRow += 2;
  });

  XLSX.utils.book_append_sheet(wb, ws, 'RelatorioCompleto');

  const excelFile = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([excelFile], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  saveAs(blob, 'RelatorioCompleto.xlsx');
}

// Função para exportar para PDF (CORRIGIDA)
function exportToPDF() {
  // Obtém a orientação da página dos botões checkbox
  const retratoCheckbox = document.getElementById('retratoPDF');
  const paisagemCheckbox = document.getElementById('paisagemPDF');

  // Define a orientação com base nos checkboxes
  let orientation = 'portrait'; // Valor padrão
  if (retratoCheckbox.checked) {
    orientation = 'portrait';
  } else if (paisagemCheckbox.checked) {
    orientation = 'landscape';
  }

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({
    orientation: orientation
  });

  const newLineHeightFactor = 0.5;
  doc.setLineHeightFactor(newLineHeightFactor);

  // Obtém o título principal (h2) e o subtítulo (h3) diretamente do HTML
  const mainTitle = document.querySelector('.col-sm-9 h2').textContent.trim();
  const subtitle = document.querySelector('.col-sm-9 h3').textContent.trim();

  // Define a largura máxima para o título
  const maxTitleWidth = 150;

  // Calcula a altura do título
  let titleHeight = doc.getTextDimensions(mainTitle, { maxWidth: maxTitleWidth }).h;

  // Posição X do título (centralizado)
  let titleX = doc.internal.pageSize.width / 2;
  if (orientation === 'portrait') {
    titleX = 85;
  }

  // Posição Y do título (com margem superior)
  let titleY = 20 + titleHeight;

  // Carrega a imagem da logo (ajuste o caminho se necessário)
  const logoImg = 'data:image/png;base64,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';


  const logoWidth = 30;
  const logoHeight = 10;
  const logoX = titleX + doc.getTextDimensions(mainTitle, { maxWidth: maxTitleWidth }).w / 2 + 5;
  const logoY = titleY - titleHeight / 2 - logoHeight / 2;

  // Seleciona todas as tabelas de dados
  const dataTables = document.querySelectorAll('.table.table-responsive.table-striped.table-bordered.table-sm');

  // Renderiza o título, subtítulo e logo na primeira página
  doc.setFontSize(16);
  doc.text(mainTitle, titleX, titleY, { align: 'center', maxWidth: maxTitleWidth });
  doc.setFontSize(14);
  doc.text(subtitle, titleX, titleY + 8, { align: 'center' });
  doc.addImage(logoImg, 'PNG', logoX, logoY, logoWidth, logoHeight);

  // Define a posição inicial Y para a primeira tabela (considerando o título e a logo)
  let startY = titleY + 30;

  dataTables.forEach((table) => {
    const tableData = extractTableData(table);
    const tableHeaders = extractTableHeaders(table);

    // Obtém o elemento <thead> da tabela
    const tableHeader = table.querySelector('thead');

    // Cria uma nova tabela HTML com o cabeçalho e os dados
   const tableHTML = `<table>${tableHeader.outerHTML}<tbody>${table.querySelector('tbody').innerHTML}</tbody></table>`;

  // Converte a string tableHTML em um elemento HTML
  const parser = new DOMParser();
  const tableElement = parser.parseFromString(tableHTML, 'text/html').querySelector('table'); 

  doc.autoTable({
    html: tableElement, // Usa o elemento HTML da tabela
    startY: startY,
    theme: 'grid',
    useHTML: true, // Ativa a renderização de HTML
    pageBreak: 'avoid',
    headStyles: {
      fillColor: [200, 230, 255],
      textColor: [0, 0, 0],
      halign: 'center', // Centraliza os cabeçalhos
       lineWidth: 1, 
      lineColor: [0, 0, 0] // Preto
    },
    bodyStyles: {
      fillColor: false,
    },
    alternateRowStyles: {
      fillColor: false,
    },
    rowPageBreak: 'noWrap',
    overFlow: 'linebreak',
    showHead: 'everyPage',
    styles: {
      fontSize: 7,
      fontStyle: 'normal',
      fontFamily: 'Calibri, sans-serif',
    },
    didDrawPage: function (data) {
      if (data.pageNumber > 1) {
        doc.setFontSize(10);
        doc.text("HOSPITAL REGIONAL DA COSTA LESTE MAGID THOMÉ", 10, 10);
        doc.text("RL06 - CUSTO SERVIÇOS AUXILIARES", 10, 18);
      }
    },
    didParseCell: function (data) {
      if (data.row.index === 0 && data.section === 'body') {
        data.cell.styles.textColor = [0, 0, 0];
      }
      if (data.section === 'body' && data.column.index === 0) {
        data.cell.styles.halign = 'left';
      }
      if (data.section === 'body' && data.column.index !== 0) {
        data.cell.styles.halign = 'center';
      }
    },
    didDrawHeader: function (data) {
      let titleHeight = doc.getTextDimensions(mainTitle, { maxWidth: maxTitleWidth }).h;
      let titleX = doc.internal.pageSize.width / 2;
      let titleY = data.settings.startY - titleHeight - 10;

      doc.setFontSize(16);
      doc.text(mainTitle, titleX, titleY, { align: 'center', maxWidth: maxTitleWidth });
      doc.setFontSize(14);
      doc.text(subtitle, titleX, titleY + 8, { align: 'center' });
      doc.addImage(logoImg, 'PNG', logoX, logoY, logoWidth, logoHeight);
    }
  });

  startY = doc.lastAutoTable.finalY + 10;
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


// Manipuladores de eventos para os botões e checkboxes
document.addEventListener('DOMContentLoaded', function() {
  const printButton = document.querySelector('.float-right.btn.btn-outline-primary');
  const fecharButton = document.getElementById('btn-fechar'); 
  const retratoCheckbox = document.getElementById('retratoPDF');
  const paisagemCheckbox = document.getElementById('paisagemPDF');

  printButton.addEventListener('click', function(event) {
    event.preventDefault();
    $('#formatoModal').modal('show');
  });

  $('#btn-excel').click(function() {
    exportToExcel();
    $('#formatoModal').modal('hide');
  });

  $('#btn-pdf').click(function() {
    exportToPDF();
    $('#formatoModal').modal('hide');
  });

  // Manipulador de evento para o botão "Fechar"
  fecharButton.addEventListener('click', function() {
    $('#formatoModal').modal('hide');
  });

  // Lógica para alternar entre os checkboxes 
  retratoCheckbox.addEventListener('change', function() {
    paisagemCheckbox.checked = !this.checked;
  });

  paisagemCheckbox.addEventListener('change', function() {
    retratoCheckbox.checked = !this.checked;
  });
});
