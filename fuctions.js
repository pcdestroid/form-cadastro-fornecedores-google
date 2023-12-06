function teste() {
  Logger.log(validaCNPJ().valido)
}

function formataCNPJ(x) {
  //Recebe um valor de CNPJ com ou sem caracteres e formata ele para "00.000.000/0001-00"

  // Remove caracteres não numéricos
  let valor = x;
  let cnpj = "";
  for (var i = 0; i < valor.length; i++) {
    if (!isNaN(parseInt(valor[i]))) {
      cnpj += valor[i];
    }
  }

  if (cnpj !== "") {
    // Adiciona a formatação desejada manualmente
    cnpj = cnpj.substring(0, 2) + '.' + cnpj.substring(2, 5) + '.' + cnpj.substring(5, 8) + '/' + cnpj.substring(8, 12) + '-' + cnpj.substring(12, 14);
  }
  return cnpj;
}

function validaCNPJ(x) {

  let valor = x;

  let retorno = { valido: true, mensagem: "ok" }

  // Remove caracteres não numéricos
  let cnpj = "";
  for (var i = 0; i < valor.length; i++) {
    if (!isNaN(parseInt(valor[i]))) {
      cnpj += valor[i];
    }
  }

  if (cnpj.length !== 14) {
    retorno = { valido: false, mensagem: "CNPJ do remetente deve conter 14 dígitos." }
    return retorno;
  }

  // Verifica se todos os dígitos são iguais
  if (todosDigitosIguais(cnpj)) {
    retorno = { valido: false, mensagem: "CNPJ do remetente inválido, todos os dígitos são iguais." }
    return retorno;
  }

  // Calcula o primeiro dígito verificador
  var tamanho = cnpj.length - 2;
  var numeros = cnpj.substring(0, tamanho);
  var digitos = cnpj.substring(tamanho);
  var soma = 0;
  var pos = tamanho - 7;

  for (var i = tamanho; i >= 1; i--) {
    soma += numeros.charAt(tamanho - i) * pos--;
    if (pos < 2) {
      pos = 9;
    }
  }

  var resultado = soma % 11 < 2 ? 0 : 11 - (soma % 11);

  if (resultado != digitos.charAt(0)) {
    retorno = { valido: false, mensagem: "CNPJ do remetente inválido, primeiro dígito verificador incorreto." }
    return retorno;
  }

  // Calcula o segundo dígito verificador
  tamanho = tamanho + 1;
  numeros = cnpj.substring(0, tamanho);
  digitos = cnpj.substring(tamanho);
  soma = 0;
  pos = tamanho - 7;

  for (var i = tamanho; i >= 1; i--) {
    soma += numeros.charAt(tamanho - i) * pos--;
    if (pos < 2) {
      pos = 9;
    }
  }

  resultado = soma % 11 < 2 ? 0 : 11 - (soma % 11);

  if (resultado != digitos.charAt(0)) {
    retorno = { valido: false, mensagem: "CNPJ do remetente inválido, segundo dígito verificador incorreto." }
    return retorno;
  }

  return true;
}

function todosDigitosIguais(cnpj) {
  for (var i = 1; i < cnpj.length; i++) {
    if (cnpj.charAt(i) !== cnpj.charAt(0)) {
      return false;
    }
  }
  return true;
}
