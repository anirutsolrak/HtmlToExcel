import * as excel from './excel.js';
import * as pdf from './pdf.js';

document.addEventListener('DOMContentLoaded', function() {
  const printButton = document.querySelector('.float-right.btn.btn-outline-primary');

  printButton.addEventListener('click', function(event) {
    event.preventDefault();
    $('#formatoModal').modal('show');
  });

  $('#btn-excel').click(function() {
    excel.exportToExcel();
    $('#formatoModal').modal('hide');
  });

  $('#btn-pdf').click(function() {
    pdf.exportToPDF(); 
    $('#formatoModal').modal('hide');
  });

  // Adiciona o evento ao botão toggle para atualizar a orientação
  const orientationSwitch = document.getElementById('orientacaoPDF');
  orientationSwitch.addEventListener('change', function () {
    // Atualiza o texto do label para refletir a nova orientação
    const label = document.querySelector('.form-check-label');
    label.textContent = this.checked ? 'Paisagem' : 'Retrato';
  });
});