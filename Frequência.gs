//Variaveis para referenciar as paginas <TESTANDO O PUSH>
var pagina = SpreadsheetApp.openById("INSERIR O URL DO GOOGLE SHEETS");
var banco = pagina.getSheetByName("INSERIR O NOME DA PAGINA DO URL");
var cadastro = pagina.getSheetByName("INSERIR O NOME DA PAGINA DO URL");

//Função para poder enviar os dados
function enviarFrequencia(){
  //Verifica se esta logado antes de enviar os dados
  if(getLogin() == '') var aviso = Browser.msgBox('Aviso!','Você não está logado!',Browser.Buttons.OK);
  if(aviso == 'ok') return;

  var confirma = Browser.msgBox('ATENÇÃO!!','Você está logado como ' + getLogin() + ' ?', Browser.Buttons.YES_NO);

  if(confirma == 'yes'){
    var data = cadastro.getRange('C8').getValue();
    var falta = cadastro.getRange('H8').getValue();

    //Confere se os campos data e falta foram preenchidos para enviar os dados
    if(data == '' || falta == ''){
      Browser.msgBox('Aviso!', 'Preencha os dados com DATA e FALTA!',Browser.Buttons.OK);
      return;
    }else{
      //Pegando o valor em que fica armazenado o ID do usuario que logou
      var idBolsista = cadastro.getRange('J1').getValue();

      //Adiciona uma linha a tabela "BANCO DE DADOS" na parte de cima
      banco.getRange('B2').activate();
      banco.insertRowsBefore(banco.getActiveRange().getRow(), 1);
      banco.getActiveRange().offset(0, 0, 1, banco.getActiveRange().getNumColumns()).activate();

      //Seleciona e copia as infos da pagina CADASTRO e cola no BANCO DE DADOS menos o ID
      cadastro.getRange('D8:H8').copyTo(banco.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);

      //Faz a concatenação para filtro dinâmico dos nomes com datas
      banco.getRange('A2').activate();  
      banco.getRange('A2').setFormula('=CONCAT(INSERIR NOME DA PAGINA DO URL!K1;INSERIR NOME DA PAGINA DO URL!C8)');
      banco.getRange('INSERIR NOME DA PAGINA DO URL!A2').copyTo(banco.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

      //Retorna para página inicial e limpa o conteúdo digitado
      cadastro.getRange('C8:H8').activate();
      cadastro.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});

      //Encerra e mostra uma mensagem
      Browser.msgBox('Aviso', 'Processo Concluído com Sucesso!', Browser.Buttons.OK);
    }
  }else{
    logout();
  }
}
