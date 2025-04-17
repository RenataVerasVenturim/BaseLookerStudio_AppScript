/*
  Direitos Autorais (c) 2025 Renata Veras Venturim

  Licença de Uso com Atribuição Obrigatória

  Este script foi desenvolvido por Renata Veras Venturim para automatizar a base de dados 
  que alimenta relatórios no Google Looker Studio, com objetivo de melhorar o acesso e 
  a transparência das informações no setor.

  Fica autorizada a utilização, cópia e redistribuição deste código, com ou sem modificações,
  desde que mantidos os créditos da autora de forma visível e permanente, tanto no código 
  quanto em qualquer sistema, dashboard ou relatório que utilize este script direta ou 
  indiretamente.

  É expressamente proibida a remoção ou ocultação do nome da autora, bem como a utilização 
  do código sem atribuição clara de autoria.

  Este software é fornecido "no estado em que se encontra", sem garantias de qualquer tipo, 
  expressas ou implícitas. O uso é de responsabilidade do usuário.

  Autoria e Desenvolvimento: Renata Veras Venturim
  Ano: 2025
*/

function gerarBaseParaLookerStudio() {
  const idPlanilha = '1mtWTEzH6ZU27OUAZvjdjHDpACt7MOgUo4ALa5FSU6Yw';
  const abaOrigem = 'Página1';
  const abaDestino = 'BaseLookerstudio';

  const planilha = SpreadsheetApp.openById(idPlanilha);
  const sheetOrigem = planilha.getSheetByName(abaOrigem);
  const ultimaLinha = sheetOrigem.getLastRow();

  const dados = sheetOrigem.getRange(5, 1, ultimaLinha - 4, 13).getValues();

  const novaBase = [];
  const cabecalho = ['E-MAIL DO VISUALIZADOR', ...sheetOrigem.getRange(5, 1, 1, 12).getValues()[0]];
  novaBase.push(cabecalho);

  for (let i = 0; i < dados.length; i++) {
    const linha = dados[i];
    const conteudoLinha = linha.slice(0, 12).map(valor => typeof valor === 'string' ? valor.trim() : valor);
    const emails = (linha[12] || '').split(',').map(e => e.trim()).filter(e => e !== '');

    for (let email of emails) {
      novaBase.push([email, ...conteudoLinha]);
    }
  }

  // Verifica se a aba destino existe
  let sheetDestino = planilha.getSheetByName(abaDestino);
  if (!sheetDestino) {
    sheetDestino = planilha.insertSheet(abaDestino);
  } else {
    sheetDestino.clearContents(); // Não apaga a aba — apenas limpa os dados
  }

  // Escreve os dados
  sheetDestino.getRange(1, 1, novaBase.length, novaBase[0].length).setValues(novaBase);
}
