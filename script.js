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

// Função para exportar para Excel
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
    const tableHeaders = extractTableHeaders(table);

    // Define as colunas a serem usadas, dependendo do tipo da tabela
    let columnsToUse = [0, 1, 2, 3]; // Colunas padrão para tabelas principais
    if (table.classList.contains('det-table')) {
      columnsToUse = [0, 1, 2]; // Colunas para tabelas de detalhes
    }

    // Adiciona os cabeçalhos da tabela
    XLSX.utils.sheet_add_aoa(ws, [tableHeaders.slice(0, columnsToUse.length)], { origin: { r: currentRow, c: 0 } });
    currentRow++;

    // Adiciona os dados da tabela com formatação de negrito
    tableData.forEach(rowData => {
      const formattedRowData = rowData.map((cellData, cellIndex) => {
        if (columnsToUse.includes(cellIndex) && cellData.includes('<strong>')) {
          const cellValue = cellData.replace(/<strong>/g, '').replace(/<\/strong>/g, '');
          return { v: cellValue, s: { font: { bold: true } } };
        } else if (columnsToUse.includes(cellIndex)) {
          return cellData;
        }
      }).filter(Boolean);
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

// Função para exportar para PDF
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

  // Carrega a imagem da logo em Base64 (substitua pelo seu código Base64)
  const logoImgData = 'data:image/png;base64,'; // Substitua pelo seu código Base64

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
  doc.addImage(logoImgData, 'PNG', logoX, logoY, logoWidth, logoHeight);

  // Define a posição inicial Y para a primeira tabela (considerando o título e a logo)
  let startY = titleY + 30;

  dataTables.forEach((table) => {
    const tableData = extractTableData(table);
    const tableHeaders = extractTableHeaders(table);

    // Obtém o número de colunas dinamicamente
    const numColumns = table.querySelectorAll('thead tr:first-child th').length;
    const columnsToUse = Array.from({ length: numColumns }, (_, i) => i);

    doc.autoTable({
      head: [tableHeaders],
      body: tableData,
      startY: startY,
      theme: 'grid',
      useHTML: true,
      pageBreak: 'avoid',
      headStyles: {
        fillColor: [200, 230, 255],
        textColor: [0, 0, 0],
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
      columns: columnsToUse.map(index => ({ dataKey: index })),
      didDrawPage: function (data) {
        if (data.pageNumber > 1) {
          doc.setFontSize(10);
          doc.text("HOSPITAL REGIONAL DA COSTA LESTE MAGID THOMÉ", 10, 10);
          doc.text("RL06 - CUSTO SERVIÇOS AUXILIARES", 10, 18);
        }
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
        if (data.cell.raw && data.cell.raw.includes('<strong>')) {
          data.cell.styles.fontStyle = 'bold';
          data.cell.text = data.cell.raw.replace(/<strong>/g, '').replace(/<\/strong>/g, '');
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
        doc.addImage(logoImgData, 'PNG', logoX, logoY, logoWidth, logoHeight);
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

// Manipuladores de eventos para os botões
document.addEventListener('DOMContentLoaded', function() {
  const printButton = document.querySelector('.float-right.btn.btn-outline-primary');

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
});
