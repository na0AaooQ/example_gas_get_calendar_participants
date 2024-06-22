// Googleカレンダー予定の参加者(ゲスト)（オーナーを含む）を取得する
function　confirmStartProcessing(messageText) {
  var res = Browser.msgBox(messageText, Browser.Buttons.OK_CANCEL);

  // okボタンが押された時
  if　(res == "ok")　{
    return true;
  // cancelボタンが押された時の動作
  } else if(res == "cancel") {
    Browser.msgBox("処理を中止します。");
    return false;
  }

}

// Googleカレンダー予定の参加者のメールアドレス一覧を取得し、除外リストと組み合わせて、配信対象アドレス一覧を生成する
function srartMakeDistributionList() {

  var execFlg = confirmStartProcessing("Googleカレンダー予定の参加者のメールアドレス一覧を取得し、除外リストと組み合わせて、配信対象アドレス一覧を生成しますか？ (ok or cancel)");

  if ( execFlg == true ) {
    // Googleカレンダー予定の参加者のメールアドレス一覧を取得し、スプレッドシートに記録する
    getCalendarParticipant();

    // 配信除外リストを読み込んで、除外リストに書かれていないGoogleカレンダー予定参加者アドレスを配信リストに出力する
    makeDistributionlist();
  }

}

function startMakeDistributionList() {

  var execFlg = confirmStartProcessing("配信対象アドレス一覧を元に、リマインドメール配信を開始しますか？ (ok or cancel)");

  if ( execFlg == true ) {
    // 配信リストに出力されたメールアドレスに対して、リマインドメールを送信する
    sendDistribution();
  }

}

// Googleカレンダー予定の参加者のメールアドレス一覧を取得し、スプレッドシートに記録する
function getCalendarParticipant() {
  var sheetParticipant = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('参加者リスト');
  var lastRow = sheetParticipant.getLastRow();
  var lastCol = sheetParticipant.getLastColumn();

  var calendar = CalendarApp.getDefaultCalendar();
  let startDate = new Date('2024/06/06 21:00:00');
  let endDate = new Date('2024/06/06 22:00:00');
  var events = calendar.getEvents(startDate, endDate); //startDateからendDateまでの予定を取得
  var includeOwner = true;
  for (var i in events) {
    var event = events[i];
    //イベントIDで指定するときには下記のコード
    var eventId = events[i].getId();
    var event = calendar.getEventById(eventId);
    var guests = event.getGuestList(includeOwner);
    var i = lastRow + 1;
    for (var j in guests) {
      sheetParticipant.getRange(i, 1).setValue(guests[j].getEmail());
      i++;
    }
  }
}

// 配信除外リストを読み込んで、除外リストに書かれていないGoogleカレンダー予定参加者アドレスを配信リストに出力する
function makeDistributionlist() {
  var sheetParticipant = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('参加者リスト');  
  var sheetExclusion = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('除外リスト');
  var sheetDistribution = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('配信対象リスト');

  var addressParticipant = "";
  var addressExclusion = "";
  var k = 2; // シートの2行目から処理対象とする
  var flgExclusion = false;
  for ( i = 2; i <= sheetParticipant.getLastRow(); i++ ) {
    flgExclusion = false;
    addressParticipant = sheetParticipant.getRange(i, 1 , 1, 1).getDisplayValue();
    console.log("参加者リストのアドレス = [" + addressParticipant + "]");

    for ( j = 2; j <= sheetExclusion.getLastRow(); j++ ) {
      addressExclusion = sheetExclusion.getRange(j, 1 , 1, 1).getDisplayValue();

      console.log("除外リストのアドレス = [" + addressExclusion + "]");

      // 参加者アドレスが除外リストに登録されている場合、配信対象リストに追加しない
      if ( addressParticipant == addressExclusion ) {
        flgExclusion = true;
        break;
      }
      flgExclusion = false;
    }

    if ( flgExclusion == true ) {
      console.log("参加者リストのアドレス = [" + addressParticipant + "] は除外リストに登録されているアドレスなので、配信対象リストには追加しません。");
    } else {
      console.log("参加者アドレスが除外リストに登録されていないので、配信対象リストへ追加します。");
      sheetDistribution.getRange(k, 1, 1, 1).setValue(addressParticipant);
      k++;
    }

  }
}

// 配信リストに出力されたメールアドレスに対して、リマインドメールを送信する
function sendDistribution() {
  var sheetDistribution = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('配信対象リスト');
  var sheetMailBody = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('配信メール文面');

  var addressDistribution = "";
  for ( i = 2; i <= sheetDistribution.getLastRow(); i++ ) {
    addressDistribution = sheetDistribution.getRange(i, 1, 1, 1).getDisplayValue();
    console.log("配信対象アドレス = [" + addressDistribution + "]");
    sheetDistribution.getRange(i, 2).setValue(true);

    sendMailSubject = sheetMailBody.getRange(1, 2 , 1, 1).getDisplayValue();
    sendMailBody = addressDistribution + " 様" + "\n"
      + sheetMailBody.getRange(2, 2 , 1, 1).getDisplayValue();

    console.log(sendMailSubject);
    console.log(sendMailBody);

    if (addressDistribution.match(/.+@.+\..+/))　{
      sendMail(addressDistribution, sendMailSubject, sendMailBody);
    }
  }
}

function sendMail(toAddress, subject, body) {
  GmailApp.sendEmail(toAddress, subject, body);
}
