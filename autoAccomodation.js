function getGmailData() {
  const query = "ねっぱん 予約のお知らせ !is:starred !in:draft";
  const threads = GmailApp.search(query,0,1);
  Logger.log({threads: threads.length});
  if(!threads.length){
    return;
  }
  const thread = threads[0];
  const messages = thread.getMessages();
  const message = messages[0];
  //電話予約キャンセルの時は、返信したものにスターをつける
  if (messages.length == 2){
    messages[1].star();
  }
  
  if(!message.isStarred()){
    const text = messages[0].getPlainBody().replace(/\r\n|\r/g,"");
    const [_0,guestName] = text.match(/宿泊者氏名\s*?：(.*?)(?:\(|チェックイン)/);
    guestName = guestName.replace(/\s*/g, "");
    Logger.log(guestName);
    var [_1, price] = text.match(/合計金額\s*?：(.*?)円/);
    var [_2, paymentMethod] = text.match(/決済方法\s*?：(.*?)===/);
    const [_3, mailContent] = text.match(/予約サイト\s*?：(.*?)===/);
    const [_4, yearIn, monthIn, dateIn] = text.match(/チェックイン\s*?：(.*?)年(.*?)月(.*?)日/);
    const [_5, yearOut, monthOut, dateOut] = text.match(/チェックアウト\s*?：(.*?)年(.*?)月(.*?)日/);
    const [_6, roomType] = text.match(/部屋タイプ\s*?：(.*?)プラン/);
    const [_7, adultPeople, childPeople] = text.match(/大人\s*?：(.*?)名子供\s*?：(.*?)名/);
    const [_8, numberOfRooms] = text.match(/室数\s*?：(.*?)室/);
    if(/現地精算額/.test(text)){
      const [_9, onSitePaymentHW] = text.match(/現地精算額:(.*?)円/);
      const [_10, prepaymentHW] = text.match(/支払済金額:(.*?)円/);
      price = "現地" + onSitePaymentHW + "円 事前" + prepaymentHW;
      paymentMethod = 'HW';
      Logger.log('HWprice' + price);
    }
    //Expediaの事前決済は 0.87で割る
    if(mailContent.indexOf('Expedia') == 0 && paymentMethod == "事前決済"){
      price = String(Number(price.replace(",", "")) / 0.87.toLocaleString());
    }
    //Agodaの事前決済は 0.88で割る
    if(mailContent.indexOf('Agoda') == 0 && paymentMethod == "事前決済"){
      price = String(Number(price.replace(",", "")) / 0.88.toLocaleString());
    }
    const mailContent2 = mailContent.replace(/(予約番号|宿泊者氏名|チェック|部屋タイプ|プラン|室数|合計金額|大人||備考|決済方法|電話番号|受付)/g," \n$&");
    const mailContent3 = mailContent2.replace("Booking.com", "Booking .com").replace("Expedia(Expedia) "Expedia");
    // 月はGoogleの仕様に合わせてマイナス１   支払済金額:(.*?)円  + prepaymantHW + "円"
    var checkInDate = new Date(yearIn, monthIn - 1, dateIn);
    var checkOutDate = new Date(yearOut, monthOut - 1, dateOut);
    var numberOfNights = (checkOutDate - checkInDate) / 86400000;
    Logger.log('泊数' + numberOfNights);
    Logger.log('room' + roomType);
    Logger.log('チェックイン月' + monthIn + 'チェックイン日' + dateIn);
    sheetName = String(Number(monthIn)) + "月";
    Logger.log('シート名' + sheetName);
    Logger.log('支払い方法' + paymentMethod);
    var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    var row = 4;
    var nextRow = 3;
    var roomTypeNum = 0;
    if(roomType.indexOf('デラックス ファミリールーム') == 0 ||
       roomType.indexOf('Deluxe Double') == 0 ||
       roomType.indexOf('Deluxe Family Room') == 0 ||
       roomType.indexOf('Hotel Base Nara Deluxe Double Room') == 0 ||
       roomType.indexOf('Deluxe Room') == 0){
        roomTypeNum = 201;
        row = 19;
        nextRow = -15;
      }
    if(roomType.indexOf('エコノミー ツインルーム') == 0 ||
       roomType.indexOf('Economy Twin') == 0 ||
       roomType.indexOf('Hotel Base Nara Economy Twin Room') == 0){
        roomTypeNum = 204;
        row = 28;
        nextRow = -6;
      }
    if(roomType.indexOf('バジェット トリプルルーム') == 0 || roomType.indexOf('301') == 0 ||
       roomType.indexOf('Budget Triple') == 0 || roomType.indexOf('Deluxe Triple') == 0 ||
       roomType.indexOf('Hotel Base Budget Triple Room') == 0){
        roomTypeNum = 301;
        row = 34;
        nextRow = 71;
      }
    if(roomType.indexOf('バジェット ツインルーム') == 0 ||
       roomType.indexOf('Budget Twin') == 0 ||
       roomType.indexOf('Hotel Base Nara Budget Twin Room') == 0 ){
        roomTypeNum = 302;
        row = 37;
        nextRow = 3;
      }
    if(roomType.indexOf('ドミトリールーム') == 0 ||
       roomType.indexOf('Single Bed') == 0 || roomType.indexOf('Mixed Shared') == 0 ||
       roomType.indexOf('1 Person in Dormitory') == 0 ||
       roomType.indexOf('Hotel Base Nara Single Bed in Mixed Dormitory Room') == 0 ||
       roomType.indexOf('Mixed Dormitory') == 0 ){
        roomTypeNum = 306;
        row = 94;
        nextRow = -12;
      }        
    const col = Number(dateIn) + 1;
    Logger.log('列数:' + col + ' row:' + row + ' col:' + col);
    while(sheet.getRange(row, col).getValue()){
      row = row + nextRow;
      // エコノミーツインを2階から埋めていくために105を回避
      if(row == 16 && roomTypeNum == 204 && nextRow == -6){
        roomTypeNum = 201
        row = 25;
        nextRow = 3
      }
      // エコノミーツインが301になるのを回避
      if(row == 34 && roomTypeNum == 201 && nextRow == 3){
        roomTypeNum = 102
        row = 7;
        nextRow = 3
      }
      // バジェットツインルームが満室で306号室になるのを回避
      if(row >= 49 && roomTypeNum == 302 && nextRow == 3){
        row = 105;
        nextRow = 3;
      }
      // ドミトリーが305号室になるのを回避 この時 nextRow==-12
      if(row == 46 && roomTypeNum == 306 && nextRow == -12){
        row = 100;
      }
      // ドミトリーが303号室になるのを回避して上段へ
      if(row == 40 && roomTypeNum == 306 && nextRow == -12){
        row = 97;
        nextRow = -6;
      }
      // ドミトリーが満室でかつ305号室になるのを回避
      if(row <= 46 && roomTypeNum == 306 && nextRow == -6){
        row = 105;
        nextRow = 3;
      }
      // 枠からはみ出たらrow=105
      if(row <= 3 || row >= 103){
        row = 105;
        nextRow = 3;
        while(sheet.getRange(row, col).getValue()){
          row = row + nextRow;
        }
      }
    }
    
    var guestNumber = (childPeople == 0) ? '大' + adultPeople : '大' + adultPeople + '小' + childPeople
    sheet.getRange(row, col).setValue(guestName + ' ' + guestNumber);
    sheet.getRange(row+1, col).setValue(price+"円 "+paymentMethod);
    sheet.getRange(row+2, col).setValue(numberOfNights + "泊 " + mailContent3);
    const color="#a1c2f3";
    const color2 ="#F5DEB3";
    const color3 = "#e06666";
    if(paymentMethod == '現払い' || paymentMethod == '現地決済' || paymentMethod == 'HW'){
      sheet.getRange(row+1, col).setBackground(color);
    }
    if(numberOfNights >= 2){
      sheet.getRange(row+2, col).setBackground(color2);
    }
    if(numberOfRooms >= 2){
      sheet.getRange(row+2, col).setBackground(color3);
    }
    //楽天トラベル提携サイトの場合は要注意
    if(mailContent.indexOf('楽天トラベル(TYMS)') == 0){
      sheet.getRange(row+1, col).setBackground(color3);
    }
    //処理が済んだものにスターをつける
    message.star();
  }      
}


function getTodayData(){
  //今日の日付データを変数hidukeに格納
  const today = new Date(); 
  
  //年・月・日・曜日を取得する
  const month = today.getMonth()+1;
  const day = today.getDate();
  const sheetName = String(Number(month)) + "月";
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  const todaySheet = SpreadsheetApp.getActive().getSheetByName("今日");
  const rangeToClear = todaySheet.getRange(2, 2, 33, 3);
  rangeToClear.clearContent();
  rangeToClear.setBackground(null);
  todaySheet.getRange(1, 1).setValue(month + "月" + day + "日");
  for ( var i = 1;  i <= 33;  i++  ) {
    var data1 = sheet.getRange(i*3 + 1, day + 1).getValue();
    var color1 = sheet.getRange(i*3 + 1, day + 1).getBackground();
    todaySheet.getRange(i + 1, 2).setValue(data1);
    if (color1 == '#e6b8af'){
      todaySheet.getRange(i + 1, 2).setBackground(color1);
    }
    var data2 = sheet.getRange(i*3 + 2, day + 1).getValue();
    var color2 = sheet.getRange(i*3 + 2, day + 1).getBackground();
    todaySheet.getRange(i + 1, 3).setValue(data2);
    if (color2 == '#a1c2f3' || color2 == '#e6b8af'){
      todaySheet.getRange(i + 1, 3).setBackground(color2);
    }
    var data3 = sheet.getRange(i*3 + 3, day + 1).getValue();
    var color3 = sheet.getRange(i*3 + 3, day + 1).getBackground();
    todaySheet.getRange(i + 1, 4).setValue(data3);
    if (color3 == '#f5deb3' || color3 == '#d9ead3'){
      todaySheet.getRange(i + 1, 4).setBackground(color3);
    }
    
  }
  
  Logger.log('sheetName' + sheetName);
  Logger.log('data' + data1);
  Logger.log('day' + day);
}


function getYesterdayData(){
  //今日の日付データを変数hidukeに格納
  const today = new Date();
  
  //年・月・日・曜日を取得する
  const month = today.getMonth()+1;
  const day = today.getDate();
  const sheetName = String(Number(month)) + "月";
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  const todaySheet = SpreadsheetApp.getActive().getSheetByName("今日");
  const rangeToClear = todaySheet.getRange("E2:E34");
  rangeToClear.clearContent();
  rangeToClear.setBackground(null);
  for ( var i = 2; i <= 34 ; i++) {
    var dataYesterday = todaySheet.getRange(i, 2).getValue();    
    todaySheet.getRange(i, 5).setValue(dataYesterday);
  }  
}
