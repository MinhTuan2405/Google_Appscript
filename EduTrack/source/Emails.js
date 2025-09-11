function emailCenter () {
  const html = HtmlService.createHtmlOutputFromFile('EmailCenter') // ref: EmailCenter.html
    .setTitle('Email Center');
  SpreadsheetApp.getUi().showSidebar(html);
}

function sendEmail () {

}