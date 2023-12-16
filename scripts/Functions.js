var planilha = SpreadsheetApp.openById("1qzaTfSzugx70CG_CRSxYL35ygODMW5aubRAP_Q9Gri0");
var forms = 'https://forms.gle/83CsbZ3muSrAqNLc8';
var bdAuxiliar = planilha.getSheetByName("bdAuxiliar");
var bdPrincipal = planilha.getSheetByName("bdPrincipal");
var bdBolsistas = planilha.getSheetByName("bdBolsistas");
var bdBolsistasAux = planilha.getSheetByName("bdBolsistasAux");
var bdSkills = planilha.getSheetByName("bdSkills");
var registro = planilha.getSheetByName('Registros');
var tabelaMensal = planilha.getSheetByName("Tabela Mensal");
var datas = planilha.getSheetByName("Datas");
var dataBolsistas = bdBolsistas.getDataRange().getValues();
var username = registro.getRange('K1').getValue();
var password = registro.getRange('K2').getValue();
var name = getNamebyId(registro.getRange('K1').getValue());
var month = tabelaMensal.getRange('H4').getValue();
var year = tabelaMensal.getRange('I4').getValue();
var url = planilha.getUrl();
url = url.replace(/edit$/, '');
var token = ScriptApp.getOAuthToken();
var response = UrlFetchApp.fetch(url + 'export?exportFormat=pdf&format=pdf&size=A4&portrait=true&fitw=true&////////gridlines=false&printtitle=false&sheetnames=false&fzr=false&gid=' + tabelaMensal.getSheetId() + '&access_token=' + token, { muteHttpExceptions: true });

//Chama a janela em HTML, define tamanho da tela e título
function showHtml(file, widht, height, title) {
    var html = HtmlService.createHtmlOutputFromFile(file).setWidth(widht).setHeight(height)
    SpreadsheetApp.getUi().showModelessDialog(html, title);
}

//Verifica se o usuário logado esta certo
function checkUser(username, password) {
    for (var i = 1; i < dataBolsistas.length; i++) {
        if (dataBolsistas[i][0] == username && dataBolsistas[i][4] == password) {
            return true; // Credenciais válidas
        }
    }
    return false; // Credenciais inválidas
}

//Retorna o (ID) do usuário pelo (NOME)
function getIdbyName(name) {
    for (var i = 1; i < dataBolsistas.length; i++) {
        if (dataBolsistas[i][3] == name) {
            return dataBolsistas[i][0];
        }
    }
    return 'Não existe nenhum ID com esse nome!';
}

//Retorna o (NOME) do usuário pelo (ID)
function getNamebyId(id) {
    for (var i = 1; i < dataBolsistas.length; i++) {
        if (dataBolsistas[i][0] == id) {
            return dataBolsistas[i][3];
        }
    }
    return 'Não existe nenhum NOME com esse id!';
}

//Retorna o (EMAIL) do usuário pelo (ID)
function getEmailbyId(id) {
    for (var i = 1; i < dataBolsistas.length; i++) {
        if (dataBolsistas[i][0] == id) {
            return dataBolsistas[i][4];
        }
    }
    return 'Não existe nenhum EMAIL com esse id!';
}

//Redireciona para a página que desejar
function openLink(link) {
    var html = `
    <html>
      <script>
        function abrirLink() {
          window.open("${link}", "_blank");
          google.script.host.close();
        }
      </script>
      <body onload="abrirLink()"></body>
    </html>`;
    var output = HtmlService.createHtmlOutput(html);
    output.setWidth(1);
    output.setHeight(1);
    SpreadsheetApp.getUi().showModalDialog(output, ' ');
}

//Verifica se existe algum valor na bdAuxiliar
function checkBdAuxiliar() {
    var linha = bdAuxiliar.getRange('A2:F2').getValue();

    if (linha != "") {
        bdPrincipal.getRange('A2').activate();
        bdPrincipal.insertRowsBefore(bdPrincipal.getActiveRange().getRow(), 1);
        bdPrincipal.getActiveRange().offset(0, 0, 1, bdPrincipal.getActiveRange().getNumColumns()).activate();
        bdAuxiliar.getRange('A2:F2').copyTo(bdPrincipal.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
        bdAuxiliar.deleteRow(2);

        checkBdAuxiliar();
    } else {
        return;
    }
}

//Atualiza a coluna da bdSkills para poder utilizar o PROCV pelo (EMAIL + MÊS)
function updateFormSkills() {
    var nomeData = bdSkills.getRange('A2:A');
    nomeData.setFormula('=IFERROR(CONCAT(C2;VLOOKUP(D2;Datas!$A$2:$D$13;4;0));"")');
}