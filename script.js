import { exportToExcel } from './excel.js'; 
import { gerarPDF } from './pdf.js'; 

document.addEventListener('DOMContentLoaded', function () {
  const printButton = document.querySelector('.float-right.btn.btn-outline-primary');

  printButton.addEventListener('click', function (event) {
    event.preventDefault();
    $('#formatoModal').modal('show');
  });

  $('#btn-fechar').click(function() {
    $('#formatoModal').modal('hide');
  });

  $('#btn-excel').click(function () {
    exportToExcel(); 
    $('#formatoModal').modal('hide');
  });

  $('#btn-pdf').click(function () {
    gerarPDF(); 
    $('#formatoModal').modal('hide');
  });
});
