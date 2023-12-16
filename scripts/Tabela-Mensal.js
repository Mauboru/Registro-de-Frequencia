//Envia o PDF para o DRIVE da Incubadora
function uploadRegister() {
    //Checando se está logado
    if (!checkUser(username, password)) {
        callScreenLogin();
        return;
    }
    //Requirindo senha novamente
    var acesso = Browser.inputBox('Acesso Restrito!', 'Digite a senha:', Browser.Buttons.OK_CANCEL);
    if (checkUser(username, acesso) == false) {
        Browser.msgBox('Acesso Negado!', 'A senha: ' + acesso + ' não é a correta', Browser.Buttons.OK);
        return;
    }

    //Link da pasta da incubadora
    var paste = DriveApp.getFolderById("1d_iK8rfNF978KaSNnYyLY8PYPsxRdNUp");
    var pasteYear = paste.getFoldersByName(year);
    var existYear = false;
    var nameArchive = name + ' - RF ' + '.pdf';
    var pdf = response.getBlob().setName(nameArchive);

    var exportOptions = {
        exportFormat: 'pdf',
        format: 'pdf',
        size: 'A4',
        portrait: false
    };

    // Procurando a pasta do ano
    while (pasteYear.hasNext()) {
        existYear = true;

        var pasteIfYear = pasteYear.next();
        var pasteMonth = pasteIfYear.getFoldersByName(month);
        var existMonth = pasteMonth.hasNext();

        if (!existMonth) {
            var pasteMonthAtt = pasteIfYear.createFolder(month); // Cria a pasta do mês se não existir
            pasteMonthAtt.createFile(pdf); // Cria o arquivo no mês correspondente
            Browser.msgBox('AVISO!', 'PDF criado com sucesso!', Browser.Buttons.OK);
            return; // Finaliza a função após criar o arquivo
        } else {
            var pasteMonthAtt = pasteMonth.next(); // Avança para a pasta do mês existente
            pasteMonthAtt.createFile(pdf); // Cria o arquivo no mês correspondente
            Browser.msgBox('AVISO!', 'PDF criado com sucesso!', Browser.Buttons.OK);
            return; // Finaliza a função após criar o arquivo
        }
    }

    // Se chegou até aqui, significa que não encontrou a pasta do ano, então cria uma nova
    if (existYear == false) {
        var newPasteYear = paste.createFolder(year);
        var newPasteMonth = newPasteYear.createFolder(month); // Cria a pasta do mês na pasta do novo ano

        if (pdf) {
            newPasteMonth.createFile(pdf); // Cria o arquivo no mês correspondente
            Browser.msgBox('AVISO!', 'PDF criado com sucesso!', Browser.Buttons.OK);
        }
    }
}

//Baixa o PDF no computador do usuario
function downloadRegister() {
    //Checando se está logado
    if (!checkUser(username, password)) {
        callScreenLogin();
        return;
    }
    //Requirindo senha novamente
    var acesso = Browser.inputBox('Acesso Restrito!', 'Digite a senha:', Browser.Buttons.OK_CANCEL);
    if (checkUser(username, acesso) == false) {
        Browser.msgBox('Acesso Negado!', 'A senha: ' + acesso + ' não é a correta', Browser.Buttons.OK);
        return;
    }

    var nameArchive = name + ' - RF ' + '.pdf';
    var pdf = response.getBlob().setName(nameArchive);

    var exportOptions = {
        exportFormat: 'pdf',
        format: 'pdf',
        size: 'A4',
        portrait: true
    };

    if (pdf) {
        var pdfUrl = "data:application/pdf;base64," + Utilities.base64Encode(pdf.getBytes());
        var htmlOutput = "<div style='text-align: center; padding-top: 40px;'>" +
            "<a href='" + pdfUrl + "' download='" + name + ' - ' + month + ".pdf'" +
            "onclick='google.script.host.close()'>Clique aqui para fazer o download!</a>" +
            "</div>";
        var ui = HtmlService.createHtmlOutput(htmlOutput);
        ui.setHeight(100);
        ui.setWidth(200); // Aumentei a largura para melhor exibir o texto
        SpreadsheetApp.getUi().showModalDialog(ui, ' ');
    } else {
        Browser.msgBox('Erro!', 'Não foi possível gerar o PDF', Browser.Buttons.OK);
    }
}

//Envia por EMAIL o PDF
function callScreenEmail() {
    //Checando se está logado
    if (!checkUser(username, password)) {
        callScreenLogin();
        return;
    }
    //Chamar tela de email em html
    showHtml('TelaEmail.html', 550, 350, ' ');
}


function sendRegisterByEmail(email, conteudo) {
    var pdf = response.getAs('application/pdf').setName(name + ' - ' + month + '.pdf');
    var subject = "Registro Mensal de " + month;

    var exportOptions = {
        exportFormat: 'pdf',
        format: 'pdf',
        size: 'A4',
        portrait: true
    };

    if (pdf) {
        var message = "Olá,\n\nSegue em anexo o PDF solicitado.\n\n" + conteudo + "\n\nAtenciosamente,\n" + name;
        GmailApp.sendEmail(email, subject, message, {
            attachments: [pdf]
        });
        Browser.msgBox('Sucesso!', 'Email enviado aos destinatários: ' + email, Browser.Buttons.OK);
    }
}