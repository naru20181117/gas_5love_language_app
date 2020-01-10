function sendMail(e){
  var ss_id="xxxxx";//スプレッドシートのIDを取得
  var sh_name = "NEW.ver"; //スプレッドシートのシート名を指定
  var sheet = SpreadsheetApp.openById(ss_id).getSheetByName(sh_name);
  var lastRow = sheet.getLastRow();
  //3. 指定するセルの範囲（lastRowから３０問の解答）を取得
  var range = sheet.getRange(lastRow, 3, 1, 30);
  //4. 値を取得する
  var value = range.getValues();
  var time = sheet.getRange(lastRow, 1).getValue();
  var name = sheet.getRange(lastRow, 2).getValue();
  var address = "xxxxxx";
  var address_to = sheet.getRange(lastRow, 33).getValue();
  var type = {};
  type = value[0];
  // ３０問の質問の答えを集めて行きます。
  for(var i=0;i<30;i++)
  {
  sheet.getRange(2, i+35).setValue(type[i])
  };
  var A = sheet.getRange("AI4").getValue();
  var B = sheet.getRange("AJ4").getValue();
  var C = sheet.getRange("AK4").getValue();
  var D = sheet.getRange("AL4").getValue();t
  var E = sheet.getRange("AM4").getValue();
  var Anum = sheet.getRange("AI3").getValue();
  var Bnum = sheet.getRange("AJ3").getValue();
  var Cnum = sheet.getRange("AK3").getValue();
  var Dnum = sheet.getRange("AL3").getValue();
  var Enum = sheet.getRange("AM3").getValue();
  var p = "";
  var text = "を多く求めます";
  if(A == 1)
    {
      p+="Positive Language「肯定的な言葉」\n"
      p+="<思いを言葉にして表すこと。感謝、賞賛、励まし。勇気を与える言葉、優しい言葉、相手を尊重する言葉。>\n"
    };
    if(B == 1){
      p+="Quality Time「充実した時間」\n"
      p+="<相手のために時間を作り、いっしょに過ごすこと。いっしょに楽しむ。目の前の相手に注意をそそぐこと。>\n"
    };
    if(C == 1){
      p+="Gift「贈り物」\n"
      p+="<相手を思っていること、考えていることを表現するプレゼント。品物の金額ではなくて、それが象徴する思いが重要。>\n"
    };
    if(D == 1){
      p+="Act of Service「サービス行為/手助け」\n"
      p+="<勉強や仕事を手伝う。猫の世話、花を活ける。料理や掃除など、ほとんどの家事や雑用はAct of Service。>\n"
    };
    if(E == 1){
      p+="Body Touch「身体的なタッチ/スキンシップ」\n"
      p+="<手をつなぐ、抱きしめる。性的なふれあい。パートナーといっしょにソファでくっついて座るなど、性的でないふれあい。>\n"
    };
  //　以下メールの中身
  var body="氏名 : "+ name + "さん\n";
  body+="結果 : \n";
  body+="A Positive Language「肯定的な言葉」： "+Anum + "/12\n";
  body+="\n";
  body+="B Quality Time「充実した時間」：         "+Bnum + "/12\n";
  body+="\n";
  body+="C Gift「贈り物」：                              "+Cnum + "/12\n";
  body+="\n";
  body+="D Act of Service「サービス行為」：      "+Dnum + "/12\n";
  body+="\n";
  body+="E Body Touch「身体的なタッチ」：        "+Enum + "/12\n";
  body+="\n";
  body+="あなたは\n";
  body+= p + text +"\n";
  body+="\n";
  body+= "以下他解答の説明:\n";
  body+="https://naru20181117.github.io/5LoveLanguage.html\n";
  body+="\n";
  body+= "もう一度回答したい場合はこちら！\n";
  body+= "https://docs.google.com/forms/d/e/1FAIpQLSdiMQs5493z7cEAIJysqDup4lXGbf-P4bfctavzoT1ECeraog/viewform\n";
  body+="\n";
  body+="参照:" + "https://www.5lovelanguages.com/";
  //　メールを管理者へ送る
  MailApp.sendEmail(address_to,"FIVE LOVE LANGUAGE TEST 結果",body);
  //　メールを入力者へ送る
  MailApp.sendEmail(address,name + "さんの結果FIVE LOVE LANGUAGE TEST",body);
}

