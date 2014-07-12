function start() {
    if (Utilities.formatDate(new Date(), "GMT+4", "HH:mm") == "19:13") { //время, когда отправится письмо
        selectSheet();
    }
}

function selectSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    for (var sheetIndex in sheets) {
        var sheet = ss.getSheets()[sheetIndex];
        sendEmail(sheet.getName());
    }
}

function mailBody(sheetName) {
    var bodyHTML = '<table border="1">';
    var listSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName); //указывается лист
    var range = listSheet.getRange("A:C"); //выбираем диапазон - полностью 3 столбца
    var rangeData = range.getValues(); //массив значений ячеек
    var rangeHeight = range.getHeight(); //количество строк в диапазоне

    for (var i = 0; i < rangeHeight; i++) { //цикл по всем строкам    
        if (rangeData[i][0] != "" && rangeData[i][0].valueOf() < (new Date()).valueOf()) { // значение ячейки - строка с номером i и столбец с номером 0 - А1, А2...
            if (rangeData[i][1] == "") { // значение ячейки - строка с номером i и столбец с номером 0 - B1, B2...
                bodyHTML += '<tr><td> ' + Utilities.formatDate(new Date(rangeData[i][0]), "GMT+4", "dd.MM.yyyy") + ' </td><td> ' + rangeData[i][1] + ' </td><td> ' + rangeData[i][2] + ' </td></tr>';
            }
            if (rangeData[i][1] != "" && rangeData[i][1].valueOf() < rangeData[i][0].valueOf()) {
                bodyHTML += '<tr><td> ' + Utilities.formatDate(new Date(rangeData[i][0]), "GMT+4", "dd.MM.yyyy") + ' </td><td> ' + Utilities.formatDate(new Date(rangeData[i][1]), "GMT+4", "dd.MM.yyyy") + ' </td><td> ' + rangeData[i][2] + ' </td></tr>';
            }
        }
    }

    bodyHTML += '</table>';
    return (bodyHTML);
}

function sendEmail(sheetName) {
    MailApp.sendEmail({
        to: 'golerdg555@gmail.com', //email
        subject: "Документ: " + sheetName,
        htmlBody: mailBody(sheetName)
    });
}