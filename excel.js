import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import { extractTableData, extractTableHeaders } from './utils.js';

// Função para gerar o arquivo Excel
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
  currentRow += 2; // Pula uma linha extra

  tables.forEach((table) => {
    const tableData = extractTableData(table);
    const tableHeaders = extractTableHeaders(table);

    // Adiciona os cabeçalhos da tabela
    XLSX.utils.sheet_add_aoa(ws, [tableHeaders], { origin: { r: currentRow, c: 0 } });
    currentRow++;

    // Adiciona os dados da tabela com formatação de negrito
    tableData.forEach(rowData => {
      const formattedRowData = rowData.map(cellData => {
        if (cellData.includes('<strong>')) {
          // Remove a tag <strong> e aplica negrito
          const cellValue = cellData.replace(/<strong>/g, '').replace(/<\/strong>/g, '');
          return { v: cellValue, s: { font: { bold: true } } }; // Aplica negrito
        } else {
          return cellData;
        }
      });
      XLSX.utils.sheet_add_aoa(ws, [formattedRowData], { origin: { r: currentRow, c: 0 } });
      currentRow++;
    });

    currentRow += 2; // Adiciona duas linhas em branco entre as tabelas
  });

  XLSX.utils.book_append_sheet(wb, ws, 'RelatorioCompleto'); // Renomeia a planilha

  const excelFile = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([excelFile], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  saveAs(blob, 'RelatorioCompleto.xlsx');
}