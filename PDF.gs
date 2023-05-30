function gerarPdf(){
  var paginaAtual = SpreadsheetApp.openById("1P1Ak83cmWLbaS5qvc4RE0SQdTd7JHfM0Td1n13hLsWg");
  var planilha = paginaAtual.getSheetByName("TABELAS HORAS");
  var url = paginaAtual.getUrl();
  url = url.replace(/edit$/,'');
  var token = ScriptApp.getOAuthToken();
  
  var exportOptions = {
    exportFormat: 'pdf',
    format: 'pdf',
    size: 'A4',
    portrait: true
  };
  
  var response = UrlFetchApp.fetch(url + 'export?exportFormat=pdf&format=pdf&size=A4&portrait=true&fitw=true&gridlines=false&printtitle=false&sheetnames=false&fzr=false&gid=' + planilha.getSheetId() + '&access_token=' + token, {
    muteHttpExceptions: true
  });
  
  var valorCelula = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B5").getValue();
  var nomeArquivo = 'Registro de Frequência' + ' - ' + valorCelula + '.pdf';
  var pdf = response.getBlob().setName(nomeArquivo);;
  
  if (pdf) {
    var folder = DriveApp.getFolderById('1d_iK8rfNF978KaSNnYyLY8PYPsxRdNUp');
    folder.createFile(pdf);
  }
  Browser.msgBox('PDF criado com sucesso!');
}
