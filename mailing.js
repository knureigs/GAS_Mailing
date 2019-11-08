function sendEmails() {
    try {
        var googgleSpreadsheet = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1dAL3mooxa48Zs0n6p5zexoEsEMk4sE7Xcr0WXNOhpk0/edit'); // обрабатываемая гуглотаблица. 
        var sheet = googgleSpreadsheet.getSheetByName("Mailing"); // лист в таблице, в котором содержатся адреса. 
        var dataMailList = sheet.getRange("C21:C" + sheet.getLastRow()).getDisplayValues(); // данные столбца с адресами.
        var dataActiveList = sheet.getRange("D21:D" + sheet.getLastRow()).getDisplayValues(); // данные столбца с флагами активности.
        var mailList = [];
        var mailRegex = new RegExp("^[a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+$"); // небольшая проверка корректности почтового адреса.

        // формирование набора адресов, по которым будет нужно выполнить рассылку.
        for (var i = 0; i < dataMailList.length; i++) {
            var email = dataMailList[i][0];
            if(!mailRegex.test(email)){
                continue;
            }
            var isActive = dataActiveList[i][0];
            if(isActive=="FALSE"){
                continue; // если этот адрес не был отмечен к рассылке - он не нужен. 
            }
            console.log("email: " + email + ", isActive " + isActive);
            mailList.push(email);
        }
        var mailsAmount = mailList.length;
        if(mailsAmount == 0) {
            Browser.msgBox("Предупреждение", "В списке нет корректных (или отмеченных к рассылке) адресов электронной почты.", Browser.Buttons.OK);
            return;
        }

        // подготовка параметров, общих для всех писем.
        var mailSubject = sheet.getRange("B2").getDisplayValue();
        var mailBody = sheet.getRange("B4").getDisplayValue();
        var mailSign = sheet.getRange("B11").getDisplayValue();

        // собственно рассылка.
        var sendedMailCount = 0; // счетчик успешно отправленных писем.
        var sendedMailAdresses = []; // перечень адресов, по которым рассылка была успешной.
        for(var j = 0; j < mailsAmount; j++) {
            var to = mailList[j];         
            try{
                MailApp.sendEmail({
                    to: to,
                    subject: mailSubject,
                    htmlBody: mailBody + "<br><br>" + mailSign 
                });
                sendedMailCount++;
                sendedMailAdresses.push(to);
            } catch(errorSending) {
                Browser.msgBox("Ошибка", "Ошибка отправки письма по адресу " + to +". " + errorSending, Browser.Buttons.OK);
            }
        }

        var remainingEmails = MailApp.getRemainingDailyQuota();
        Browser.msgBox("Информация", "Успешно отправлено " + sendedMailCount + " почтовых сообщений, по адресам " + sendedMailAdresses.join(", ") + ".\\nОстаток квоты на сегодня: " + remainingEmails, Browser.Buttons.OK);
    } catch(error) {
        Browser.msgBox(error);
    }
}