// Carregando as bibliotecas necessárias (use CDN ou baixe os arquivos)
// REMOVA os CDNs do script.js
// <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
// <script src="https://cdn.jsdelivr.net/npm/file-saver@2.0.5/dist/FileSaver.min.js"></script>

import { saveAs } from 'file-saver'; // Importa a função saveAs do file-saver

// Função para gerar o arquivo Excel
function exportToExcel() {
  // Selecionando a tabela HTML
  const table = document.querySelector('.pri-table'); 

  // Convertendo a tabela HTML para um objeto Workbook do xlsx
  const wb = XLSX.utils.table_to_book(table);

  // Definindo o nome do arquivo
  const fileName = 'RelatorioCustos.xlsx';

  // Criando um Blob do arquivo Excel
  const excelFile = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([excelFile], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

  // Salvando o arquivo usando FileSaver.js
  saveAs(blob, fileName); // Use saveAs para salvar o arquivo
}

// Adicionando o evento de clique ao botão "Imprimir"
document.addEventListener('DOMContentLoaded', function() { // Espera o DOM carregar
  const printButton = document.querySelector('.float-right.btn.btn-outline-primary'); 

  // Desativa o comportamento padrão do botão "Imprimir"
  printButton.addEventListener('click', function(event) {
    event.preventDefault(); // Impede o comportamento padrão do botão
    exportToExcel(); // Chama a função para gerar o arquivo Excel
  });
});