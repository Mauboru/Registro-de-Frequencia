var pagina = SpreadsheetApp.openById("INSERIR O URL DO GOOGLE SHEET");
var bancoA = pagina.getSheetByName("INSERIR O NOME DA PAGINA DO URL");
var bancoP = pagina.getSheetByName("INSERIR O NOME DA PAGINA DO URL");

function timerBanco() {
  var valor = bancoA.getRange('A2').getValue();

  if(valor != ""){
    var linha = bancoA.getRange('A2:F2');

    bancoP.getRange('A2').activate();
    bancoP.insertRowsBefore(bancoP.getActiveRange().getRow(), 1);
    bancoP.getActiveRange().offset(0, 0, 1, bancoP.getActiveRange().getNumColumns()).activate();
    bancoA.getRange('A2:F2').copyTo(bancoP.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
    bancoA.deleteRow(2);

    timerBanco();    
  }else{
    return;
  }
}
