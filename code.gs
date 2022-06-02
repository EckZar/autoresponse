function onOpen(){
  SpreadsheetApp.getUi().createAddonMenu("__MENU__").addItem("отправить письма", "start").addToUi();
}

function start(){
  main.getRange(2, 2, main.getLastRow()-1, 15).getValues().forEach((row, index) => {row[14] ? "":sendMessage(row[7], row[0], row[1], index, row[13])});
}

function sendMessage(email, name, position, index, decision){
  switch(decision){
    case 'Отказ':
      declineMessage(email, name, position);
      main.getRange(index+2, 16).setValue(true);
      Logger.log(`Sending decline for ${email}`);
      break;
    case 'Приглашение':
      inviteMessage(email, name, position);
      main.getRange(index+2, 16).setValue(true);
      Logger.log(`Sending invite for ${email}`);
      break;
    case 'Рассмотрение':
      observeMessage(email, name, position);
      main.getRange(index+2, 16).setValue(true);
      Logger.log(`Sending observe for ${email}`);
      break;
    default:
      Logger.log(row);
  }
}

function declineMessage(email, name, position){
  
  let body =  `Здравствуйте, ${name}!` + newRow +
              `Большое спасибо за интерес, проявленный к вакансии "${position}". К сожалению,` + newRow +
              `в настоящий момент мы не готовы пригласить Вас на дальнейшее интервью. Мы` + newRow +
              `внимательно ознакомились с Вашим резюме, и, возможно, вернемся к Вашей` + newRow +
              `кандидатуре, когда у нас возникнет такая потребность.` + newRow +
              `С уважением,` + newRow +
              `Отдел персонала Самокат`;

  GmailApp.sendEmail(email, "ОТКАЗ", body);

}

function inviteMessage(email, name, position){

  let body =  `Здравствуйте, ${name}!` + newRow +
              `Благодарим Вас за отклик на вакансию "${position}". Ваша кандидатура показалась` + newRow +
              `нам очень интересной. Мы хотели бы пригласить Вас на интервью. Перезвоните,` + newRow +
              `пожалуйста, в рабочее время по телефону +7-966-774-62-17 - Ольга.` + newRow +
              `С уважением,` + newRow +
              `Отдел персонала Самокат`;

  GmailApp.sendEmail(email, "ПРИГЛАШЕНИЕ НА ТЕЛЕФОННОЕ ИНТЕРВЬЮ", body);

}

function observeMessage(email, name, position){

  let body =  `Здравствуйте, ${name}!` + newRow +
              `Компания «Самокат» рассмотрит Ваше резюме на вакансию «${position}» и позже` + newRow +
              `сообщит Вам о своем решении.` + newRow +
              `С уважением,` + newRow +
              `Отдел персонала Самокат`;

  GmailApp.sendEmail(email, "РАССМОТРЕНИЕ", body);

}

const main = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Ответы на форму (1)");
const newRow = String.fromCharCode(10);
