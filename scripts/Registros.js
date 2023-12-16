//Chama a tela de login
function callScreenLogin() {
    if (checkUser(username, password) == true) {
        return Browser.msgBox('Atenção!', 'Já tem algum bolsista logado!', Browser.Buttons.OK);
    }
    showHtml('TelaLogin.html', 700, 450, ' ');
}

//Verificação de usuario ao logar
function login(username, password) {
    for (var i = 1; i < dataBolsistas.length; i++) {
        if (dataBolsistas[i][0] == username && dataBolsistas[i][4] == password) {
            registro.getRange('K1').setValue(username);
            registro.getRange('K2').setValue(password);
            return true;
        }
    }
    return false;
}

//Sai da sessão logada
function logout() {
    if (username == "username" && password == "password") {
        Browser.msgBox('Atenção!', 'Não há ninguém logado!', Browser.Buttons.OK);
    } else {
        registro.getRange('K1').setValue('username');
        registro.getRange('K2').setValue('password');
    }
}

//Chama a tela de opções
function callScreenOptions() {
    showHtml('TelaOpcoes.html', 220, 200, 'Opções de Usuário');
}

//Envia o registro
function addRegister() {
    var date = registro.getRange('C9').getValue();
    var status = registro.getRange('H9').getValue();
    var hours = registro.getRange("E9").getValue() - registro.getRange("D9").getValue();
    var time = 3600000;
    var interval = registro.getRange('F9').getValue();

    //Verifica se esta logado antes de enviar os dados
    if (!checkUser(username, password)) {
        callScreenLogin();
        return;
    }

    //Confere se os campos data e falta foram preenchidos para enviar os dados
    if (date == '' || status == '') {
        Browser.msgBox('Aviso!', 'Preencha os dados com DATA e FALTA!', Browser.Buttons.OK);
        return;
    }

    //Verificando tempo trabalhado
    if (hours < (time * 6)) {
        if (interval > 0) {
            showHtml('TelaAviso.html', 500, 800, 'Aviso!');
            return;
        }
    }
    if (hours >= (time * 6) && hours < (time * 7)) {
        if (interval > 0) {
            showHtml('TelaAviso.html', 500, 500, 'Aviso!');
            return;
        }
    }

    //Adiciona uma linha a tabela "BANCO DE DADOS" na parte de cima
    bdAuxiliar.getRange('B2').activate();
    bdAuxiliar.insertRowsBefore(bdAuxiliar.getActiveRange().getRow(), 1);
    bdAuxiliar.getActiveRange().offset(0, 0, 1, bdAuxiliar.getActiveRange().getNumColumns()).activate();

    //Seleciona e copia as infos da pagina REGISTRO e cola no BANCO DE DADOS menos o ID
    registro.getRange('D9:H9').copyTo(bdAuxiliar.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

    //Faz a concatenação para filtro dinâmico dos nomes com datas
    bdAuxiliar.getRange('A2').activate();
    bdAuxiliar.getRange('A2').setFormula('=CONCAT(Registros!K1;Registros!C9)');
    bdAuxiliar.getRange('bdAuxiliar!A2').copyTo(bdAuxiliar.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

    //Retorna para página inicial e limpa o conteúdo digitado
    registro.getRange('C9:H9').activate();
    registro.getActiveRangeList().clear({ contentsOnly: true, skipFilteredRows: true });

    //Encerra e mostra uma mensagem
    Browser.msgBox('Aviso', 'Processo Concluído com Sucesso!', Browser.Buttons.OK);
}