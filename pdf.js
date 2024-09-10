// pdf.js

// Certifique-se de que o botão "PDF" dentro do modal tenha um ID
const botaoGerarPDF = document.getElementById('btn-pdf');
botaoGerarPDF.addEventListener('click', () => {
  gerarPDF();
});

function gerarPDF() {
  // Acessando jsPDF no escopo global (window)
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({
    orientation: 'landscape',
  });

 // Define as margens personalizadas
  doc.margins = {
    top: 30,
    right: 10,
    bottom: 10,
    left: 10,
  };

  const logoImg = 'data:image/png;base64,...'; // Insira o código base64 da imagem do logo aqui
  const logoWidth = 30;
  const logoHeight = 10;

  // Capturar título dinamicamente
  const tituloPrincipal = document.querySelector('.col-sm-9 h2').textContent.trim();
  const subtitulo = document.querySelector('.col-sm-9 h3').textContent.trim();

  // Calcula a altura do título principal
  const maxTitleWidth = 350;
  let titleHeight = doc.getTextDimensions(tituloPrincipal, { maxWidth: maxTitleWidth }).h;

  // Posição X do título principal (centralizado com ajuste para a logo)
  let titleX = (doc.internal.pageSize.getWidth() - logoWidth) / 2 - 5;

  // Posição Y do título (com margem superior)
  let titleY = 20 + titleHeight;

  // Posição da logo (à direita do título)
  const logoX = titleX + doc.getTextDimensions(tituloPrincipal, { maxWidth: maxTitleWidth }).w / 2 + 10;
  const logoY = titleY - titleHeight / 2 - logoHeight / 2;

  // Renderiza o título, subtítulo e logo
  doc.setFontSize(16);
  doc.text(tituloPrincipal, titleX, titleY, { align: 'center', maxWidth: maxTitleWidth });
  doc.setFontSize(14);
  doc.text(subtitulo, titleX, titleY + 8, { align: 'center' });
  doc.addImage(logoImg, 'PNG', logoX, logoY, logoWidth, logoHeight);

  // Define o espaçamento inicial da primeira tabela
  let startY = titleY + 30; 

  // Obter todas as tabelas com a classe "pri-table"
  const tabelasPrincipais = document.querySelectorAll('.pri-table');

  tabelasPrincipais.forEach((tabela, index) => {
    // Obter o título da tabela
    const tituloTabela = tabela.previousElementSibling.textContent.trim();

    // Extrair dados da tabela principal
    const dadosTabelaPrincipal = extrairDadosTabela(tabela);

    // Adicionar tabela principal ao PDF
    adicionarTabelaAoPDF(doc, dadosTabelaPrincipal, tituloTabela, startY);

    // Atualiza startY para a próxima tabela 
    startY = doc.lastAutoTable.finalY + 10;

    // Verificar se há uma tabela de detalhes relacionada
    const tabelaDetalhes = tabela.nextElementSibling.classList.contains('det-table')
      ? tabela.nextElementSibling
      : null;

    if (tabelaDetalhes) {
      const dadosTabelaDetalhes = extrairDadosTabela(tabelaDetalhes);
      adicionarTabelaAoPDF(doc, dadosTabelaDetalhes, 'Detalhes', startY);

      // Atualiza startY após a tabela de detalhes
      startY = doc.lastAutoTable.finalY + 10;
    }
  });

  // Adiciona o rodapé e a numeração de páginas
  adicionarNumeracaoPaginas(doc);

  doc.save(`${tituloPrincipal}_${subtitulo}.pdf`);
}

// Função para extrair dados de uma tabela HTML
function extrairDadosTabela(tabela) {
  const dados = [];
  const cabecalho = [];
  const linhas = tabela.querySelectorAll('tbody tr');

  // Extrair dados do cabeçalho
  tabela.querySelectorAll('thead th').forEach(th => {
    cabecalho.push(th.textContent.trim()); 
  });
  dados.push(cabecalho); 

  // Extrair dados das linhas
  linhas.forEach(linha => {
    const linhaDados = [];
    linha.querySelectorAll('td').forEach(td => {
      linhaDados.push(td.textContent.trim()); 
    });
    dados.push(linhaDados);
  });

  return dados;
}

function adicionarCabecalho(doc, tituloPrincipal, subtitulo) {
  let yPos = 10;
  doc.text(tituloPrincipal, doc.internal.pageSize.width / 2, yPos, { align: 'center', fontSize: 10 });
  yPos += 7;
  doc.text(subtitulo, doc.internal.pageSize.width / 2, yPos, { align: 'center', fontSize: 10 });
}

function adicionarTabelaAoPDF(doc, dadosTabela, titulo, startY) {
  const titleHeight = doc.getTextDimensions(titulo).h;

  doc.setFontSize(10);
  doc.text(titulo, doc.internal.pageSize.width / 2, startY + titleHeight - 5, {
    align: 'center'
  });

  doc.autoTable({
    head: [dadosTabela[0]],
    body: dadosTabela.slice(1),
    startY: startY + titleHeight + 5,
    theme: 'grid',
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
    pageBreak: 'auto',
    rowPageBreak: 'noWrap',
    overFlow: 'linebreak',
    showHead: 'everyPage',
    styles: {
      fontSize: 7,
      fontStyle: 'normal',
      fontFamily: 'Calibri, sans-serif',
    },
    columnStyles: {
      0: {
        halign: 'left',
      },
      1: {
        halign: 'center'
      },
      2: {
        halign: 'center'
      },
      3: {
        halign: 'center'
      },
      4: {
        halign: 'center'
      },
      5: {
        halign: 'center'
      },
      6: {
        halign: 'center'
      },
      7: {
        halign: 'center'
      },
      8: {
        halign: 'center'
      },
      9: {
        halign: 'center'
      },
      10: {
        halign: 'center'
      },
      11: {
        halign: 'center'
      },
      12: {
        halign: 'center'
      },
      13: {
        halign: 'center'
      }
    },
    didParseCell: function (data) {
      if (data.row.index === 0 && data.section === 'body') {
        data.cell.styles.fillColor = [200, 230, 255];
        data.cell.styles.textColor = [0, 0, 0];
      }
      if (data.cell.raw && data.cell.raw.includes('<strong>')) {
        data.cell.styles.fontStyle = 'bold';
        data.cell.text = data.cell.raw.replace(/<strong>/g, '').replace(/<\/strong>/g, '');
      }
    }
  });
}

function adicionarNumeracaoPaginas(doc) {
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
export { gerarPDF };
