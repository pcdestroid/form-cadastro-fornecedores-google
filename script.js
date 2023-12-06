//Google app Script

function Centralizar() {

  let sheet = SpreadsheetApp.openById('').getSheetByName("Respostas ao formulário 1");
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setHorizontalAlignment('center');

  let dados = SpreadsheetApp.openById('').getSheetByName("Dados");
  let fornecedores = SpreadsheetApp.openById('').getSheetByName("Fornecedores");

  let ul = sheet.getLastRow();
  let cnpj = sheet.getRange(ul, 4).getValue();
  let fornecedor = sheet.getRange(ul, 5).getValue();
  let tipo_pessoa = sheet.getRange(ul, 6).getValue();

  if (tipo_pessoa == "Pessoa Jurídica") {
    cnpj = formataCNPJ(cnpj) // Formata o CNPJ
    sheet.getRange(ul, 4).setValue(cnpj) // Inserir CNPJ formatado

    let validar_cnpj = validaCNPJ(cnpj)

    if (!validar_cnpj.valido) {
      sheet.getRange(ul, 4).setBackground('#FFFF00');
      sheet.getRange(ul, 4).setValue(cnpj + ' | ' + validar_cnpj.mensagem)
    }
  }

  let cidade = sheet.getRange(ul, 14).getValue();
  let telefone = sheet.getRange(ul, 16).getValue();
  let contato = sheet.getRange(ul, 19).getValue();
  let email = sheet.getRange(ul, 20).getValue();
  let endereco = sheet.getRange(ul, 10).getValue();
  let bairro = sheet.getRange(ul, 13).getValue();
  let uf = sheet.getRange(ul, 15).getValue();
  let cep = sheet.getRange(ul, 9).getValue();
  let faturamentoMinimo = sheet.getRange(ul, 28).getValue();
  let produtoServico = sheet.getRange(ul, 32).getValue();
  let localAtendimento = sheet.getRange(ul, 33).getValue();
  let cadastros = fornecedores.getRange(2, 1, fornecedores.getLastRow(), 1).getValues().toString().split(',');
  let loc = cadastros.indexOf(cnpj);

  let valores = fornecedores.getRange(2, 1, fornecedores.getLastRow(), 2).getValues();

  for (let i = 0; i < valores.length; i++) { valores[i][1] = valores[i][0] + ' | ' + valores[i][1] }

  dados.getRange(2, 1, fornecedores.getLastRow(), 2).setValues(valores);

  if (loc < 0) {
    let ulf = fornecedores.getLastRow() + 1;
    fornecedores.getRange(ulf, 1).setValue(cnpj);
    fornecedores.getRange(ulf, 2).setValue(fornecedor);
    fornecedores.getRange(ulf, 3).setValue(produtoServico);
    fornecedores.getRange(ulf, 4).setValue(cidade);
    fornecedores.getRange(ulf, 5).setValue(telefone);
    fornecedores.getRange(ulf, 6).setValue(contato);
    fornecedores.getRange(ulf, 7).setValue(email);
    fornecedores.getRange(ulf, 8).setValue(endereco);
    fornecedores.getRange(ulf, 9).setValue(bairro);
    fornecedores.getRange(ulf, 10).setValue(uf);
    fornecedores.getRange(ulf, 11).setValue(cep);
    fornecedores.getRange(ulf, 12).setValue('A');
    fornecedores.getRange(ulf, 13).setValue(faturamentoMinimo);
    fornecedores.getRange(ulf, 16).setValue('Verificar');
    fornecedores.getRange(ulf, 18).setValue(localAtendimento);
    fornecedores.getRange(ulf, 20).setValue(new Date()); //Data do Registro
    fornecedores.getRange(ulf, 21).setValue("Novo cadastro"); //Tipo de Registro
    sendMessage(`Novo cadastro - ${fornecedor}`)
  }
  else {
    fornecedores.getRange(loc + 2, 1).setValue(cnpj);
    fornecedores.getRange(loc + 2, 2).setValue(fornecedor);
    fornecedores.getRange(loc + 2, 3).setValue(produtoServico);
    fornecedores.getRange(loc + 2, 4).setValue(cidade);
    fornecedores.getRange(loc + 2, 5).setValue(telefone);
    fornecedores.getRange(loc + 2, 6).setValue(contato);
    fornecedores.getRange(loc + 2, 7).setValue(email);
    fornecedores.getRange(loc + 2, 8).setValue(endereco);
    fornecedores.getRange(loc + 2, 9).setValue(bairro);
    fornecedores.getRange(loc + 2, 10).setValue(uf);
    fornecedores.getRange(loc + 2, 11).setValue(cep);
    fornecedores.getRange(loc + 2, 12).setValue('A');
    fornecedores.getRange(loc + 2, 13).setValue(faturamentoMinimo);
    fornecedores.getRange(loc + 2, 16).setValue('Verificar');
    fornecedores.getRange(loc + 2, 18).setValue(localAtendimento);
    fornecedores.getRange(loc + 2, 20).setValue(new Date()); //Data do Registro
    fornecedores.getRange(loc + 2, 21).setValue("Atualização do cadastro"); //Tipo de Registro
    sendMessage(`Dados atualizados - ${fornecedor}`)
  }
};

function cadastro() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('AE2').activate();
  spreadsheet.insertSheet(1);
};

function cadastro1() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('2:2').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('AA2'));
  spreadsheet.getActiveRangeList().setBackground('#d9d9d9')
    .setBackground('#ffffff');
};

const GOOGLE_CHAT_WEBHOOK_LINK = "";

function sendMessage(text) {
  const payload = JSON.stringify({ text: text });
  const options = {
    method: 'POST',
    contentType: 'application/json',
    payload: payload,
  };
  UrlFetchApp.fetch(GOOGLE_CHAT_WEBHOOK_LINK, options);
}
