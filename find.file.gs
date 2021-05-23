function sendMail(e){
  var ss_id="xxxxx";//スプレッドシートのIDを取得
  var sh_name = "ALL_data"; //スプレッドシートのシート名を指定
  var sheet = SpreadsheetApp.openById(ss_id).getSheetByName(sh_name);
  var lastRow = sheet.getLastRow();
  //3. 指定するセルの範囲（lastRowから３０問の解答）を取得
  var range = sheet.getRange(lastRow, 3, 1, 30);
  //4. 値を取得する
  var value = range.getValues();
  var time = sheet.getRange(lastRow, 1).getValue();
  var name = sheet.getRange(lastRow, 2).getValue();
  var address = "t.naruhiro.1026@gmail.com";
  var address_to = sheet.getRange(lastRow, 33).getValue();
  // var address_from = "app.sandbox.naru@gmail.com";
  var type = {};
  type = value[0];
  // ３０問の質問の答えを集めて行きます。
  for(var i=0;i<30;i++) {
    sheet.getRange(2, i+40).setValue(type[i])
  };
  var image = createChart(sheet)
  var p = "";
  if(getValue(sheet, "AI3") == 1)
    {
      p+="Positive Language「肯定的な言葉」<br>"
      p+="思いを言葉にして表すこと。感謝、賞賛、励まし。勇気を与える言葉、優しい言葉、相手を尊重する言葉。<br>"
    };
    if(getValue(sheet, "AJ3") == 1){
      p+="Quality Time「充実した時間」<br>"
      p+="相手のために時間を作り、いっしょに過ごすこと。いっしょに楽しむ。目の前の相手に注意をそそぐこと。<br>"
    };
    if(getValue(sheet, "AK3") == 1){
      p+="Gift「贈り物」<br>"
      p+="相手を思っていること、考えていることを表現するプレゼント。品物の金額ではなくて、それが象徴する思いが重要。<br>"
    };
    if(getValue(sheet, "AL3") == 1){
      p+="Act of Service「サービス行為/手助け」<br>"
      p+="勉強や仕事を手伝う。猫の世話、花を活ける。料理や掃除など、ほとんどの家事や雑用はAct of Service。<br>"
    };
    if(getValue(sheet, "AM3") == 1){
      p+="Body Touch「身体的なタッチ/スキンシップ」<br>"
      p+="手をつなぐ、抱きしめる。性的なふれあいも含む。パートナーといっしょにソファでくっついて座るなどのふれあい。<br>"
    };
  //　以下メールの中身
  var body="<div align='center' style='margin: 0 auto, width: 80%;'>";
  body="氏名 : "+ name + "さん<br>";
  body+="結果 : <br>";
  body+="<img src='cid:inlineImg'><br>";
  body+="A Positive Language「肯定的な言葉」： "+getValue(sheet, "AI2") + "/12<br>";
  body+="B Quality Time「充実した時間」：      "+getValue(sheet, "AJ2") + "/12<br>";
  body+="C Gift「贈り物」：                   "+getValue(sheet, "AK2") + "/12<br>";
  body+="D Act of Service「サービス行為」：    "+getValue(sheet, "AL2") + "/12<br>";
  body+="E body Touch「身体的なタッチ」：       "+getValue(sheet, "AM2") + "/12<br>";
  body+="<br>";
  body+="あなたは<br>";
  body+="<b>"+ p + "</b>を多く求めます<br>";
  body+="<br>";
  body+="<a href='https://naru20181117.github.io/5LoveLanguage.html'>より詳しい説明はこちら</a><br>";
  body+="<br>";
  body+="<a href='https://docs.google.com/forms/d/e/1FAIpQLSdiMQs5493z7cEAIJysqDup4lXGbf-P4bfctavzoT1ECeraog/viewform'>もう一度回答したい場合はこちら！</a><br>";
  body+="<br>";
  body+="<a href='https://www.5lovelanguages.com'>参照</a><br>";
  body+="</div>";

  var options = {
    "name":'5_love_language_check',
    "htmlBody":body,
    "inlineImages":{ inlineImg:image }
  };

  //　メールを入力者へ送る
  MailApp.sendEmail(address_to,"FIVE LOVE LANGUAGE TEST 結果",body,options);

  //　メールを管理者へ送る
  MailApp.sendEmail(address,name+"さんの結果FIVE LOVE LANGUAGE TEST",body,options);
}

function createChart(mySheet) {
  // チャートの削除
  var chart=mySheet.getCharts()[0];
  mySheet.removeChart(chart);
  var range = mySheet.getRange("AI1:AM2");
  
  var chart = mySheet.newChart()
              .addRange(range)
              .setChartType(Charts.ChartType.COLUMN)
              .setOption('useFirstColumnAsDomain', false)
              .setNumHeaders(1)
              .setPosition(15,35,0,0)
              .setOption('title', '愛の言葉割合')
              .build();
              mySheet.insertChart(chart);
  var imageBlob = chart.getBlob().getAs('image/png').setName("chart_image.png");//グラフの画像を取得
  return imageBlob
}

function getValue(sheet, cell) {
  // シート上の値の取得
  return sheet.getRange(cell).getValue()
}

