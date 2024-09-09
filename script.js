// script.js 
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
    exportToExcel(); // Chama a função exportToExcel (agora no escopo global)
    $('#formatoModal').modal('hide');
  });

  $('#btn-pdf').click(function () {
    gerarPDF(); // Chama a função gerarPDF (agora no escopo global)
    $('#formatoModal').modal('hide');
  });
});
