//import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import { extractTableData, extractTableHeaders } from './utils.js';

export function exportToExcel() {
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
