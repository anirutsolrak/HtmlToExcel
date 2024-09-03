import * as excel from './excel.js'; 
import * as pdf from './pdf.js'; 

// Aguarda o carregamento do DOM
document.addEventListener('DOMContentLoaded', function () {
  const printButton = document.querySelector('.float-right.btn.btn-outline-primary');

  $('#btn-fechar').click(function() {
  $('#formatoModal').modal('hide');
});


  printButton.addEventListener('click', function (event) {
    event.preventDefault();
    $('#formatoModal').modal('show');
  });

  $('#btn-excel').click(function () {
    excel.exportToExcel(); // Chama a função exportToExcel do arquivo excel.js
    $('#formatoModal').modal('hide');
  });

  $('#btn-pdf').click(function () {
    pdf.exportToPDF(); // Chama a função exportToPDF do arquivo pdf.js
    $('#formatoModal').modal('hide');
  });
});
