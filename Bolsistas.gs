//Variaveis para referenciar as paginas
var pagina = SpreadsheetApp.openById("INSERIR O URL DO GOOGLE SHEET");
var cadastro = pagina.getSheetByName("INSERIR O NOME DA PAGINA DA URL");

var id = [];
var nome = [];
var email = [];

id.push("CRIAR UM ID");
nome.push("CRIAR NOME DO ID");
email.push("EMAIL");

function login() {
  var loginId = Browser.inputBox('Login', 'Escreva seu ID: ', Browser.Buttons.OK_CANCEL);
  var isReal = false;

  if(loginId == 'cancel') return;
  
  for(var i = 0; i < id.length; i++){
    if(loginId.toUpperCase() == id[i]){
      cadastro.getRange('K1').setValue(loginId);
      isReal = true;
      break;
    }
  }
  if(isReal == false) Browser.msgBox('Aviso!','O ID = ['+ loginId +'] não existe',Browser.Buttons.OK);
  if(isReal){
    cadastro.getRange('H5').setValue(getLogin());
    Browser.msgBox('Atenção!','Você está conectado como '+ getLogin() +'!',Browser.Buttons.OK);
  }
}

function logout(){
  cadastro.getRange('H5').setValue("");
  cadastro.getRange('K1').setValue("");
}

//Retorna o nome do bolsista
function getLogin(){
  var loginId = cadastro.getRange('K1').getValue();
  var loginNome = "";
  for(var i = 0; i < id.length; i++){
    if(loginId.toUpperCase() == id[i]){
      loginNome = nome[i];
      break;
    }
  }
  return loginNome
}
