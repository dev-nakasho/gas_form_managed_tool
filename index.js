var model = {
  // アクティブスプレッドシート 
  // activespreadsheet: {},
  activespreadsheet: SpreadsheetApp.getActiveSpreadsheet(),
  questionarray: [], // テスト
  answerarray: [], // テスト
};

function createForm() {
  // スプレッドシートオブジェクト
  model.activespreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // ２次元配列データ
  const overview = model.activespreadsheet.getSheetByName("検定概要").getDataRange().getValues();
  const formTitle = overview[0][1]; // タイトル
  const formDescription = overview[1][1]; // 概要

  // フォーム作成
  const form = FormApp.create(formTitle);
  // 説明を追加
  form.setDescription(formDescription);

  // ２次元配列データ：設問
  const examination = model.activespreadsheet.getSheetByName("検定データ").getDataRange().getValues();
  examination.shift();

  // 住所
  const residence = examination.map(record => record[0]).filter(value => value);

  // Eメールバリデーション
  // const emailValidation = FormApp.createTextValidation().requireTextIsEmail().build();

  // 記述式：氏名
  form.addTextItem()
    .setTitle("氏名")
    .setRequired(true);

  // 記述式：メールアドレス
  form.addTextItem()
    .setTitle("メールアドレス")
    .setRequired(true);

  // セレクトリスト：住所
  form.addListItem()
    .setTitle("住所")
    .setChoiceValues(residence)
    .setRequired(true);

  // タイトル
  const question_title = model.activespreadsheet.getSheetByName("検定データ").getDataRange().getValues();

  // 検定問題用変数
  // 問１
  const select_list_1 = examination.map(record => record[1]).filter(value => value);
  // 問２
  const select_list_2 = examination.map(record => record[2]).filter(value => value);
  // 問３
  const select_list_3 = examination.map(record => record[3]).filter(value => value);
  // 問４
  const select_list_4 = examination.map(record => record[4]).filter(value => value);
  // 問５
  const select_list_5 = examination.map(record => record[5]).filter(value => value);
  // 問６
  const select_list_6 = examination.map(record => record[6]).filter(value => value);
  // 問７
  const select_list_7 = examination.map(record => record[7]).filter(value => value);
  // 問８
  const select_list_8 = examination.map(record => record[8]).filter(value => value);
  // 問９
  const select_list_9 = examination.map(record => record[9]).filter(value => value);
  // 問１０
  const select_list_10 = examination.map(record => record[10]).filter(value => value);
  // 問１１
  const select_list_11 = examination.map(record => record[11]).filter(value => value);
  // 問１２
  const select_list_12 = examination.map(record => record[12]).filter(value => value);

  // for (let i = 1; i <= question_title[0].length; i++) {
  // }

  // 検定問題
  // ラジオボタン

  form.addMultipleChoiceItem()
    .setTitle(question_title[0][1].toString())
    .setChoiceValues(select_list_1)
    .setRequired(true);

  // ラジオボタン
  form.addMultipleChoiceItem()
    .setTitle(question_title[0][2].toString())
    .setChoiceValues(select_list_2)
    .setRequired(true);

  // ラジオボタン
  form.addMultipleChoiceItem()
    .setTitle(question_title[0][3].toString())
    .setChoiceValues(select_list_3)
    .setRequired(true);

  // ラジオボタン
  form.addMultipleChoiceItem()
    .setTitle(question_title[0][4].toString())
    .setChoiceValues(select_list_4)
    .setRequired(true);

  // ラジオボタン
  form.addMultipleChoiceItem()
    .setTitle(question_title[0][5].toString())
    .setChoiceValues(select_list_5)
    .setRequired(true);

  // ラジオボタン
  form.addMultipleChoiceItem()
    .setTitle(question_title[0][6].toString())
    .setChoiceValues(select_list_6)
    .setRequired(true);

  // ラジオボタン
  form.addMultipleChoiceItem()
    .setTitle(question_title[0][7].toString())
    .setChoiceValues(select_list_7)
    .setRequired(true);

  // ラジオボタン
  form.addMultipleChoiceItem()
    .setTitle(question_title[0][8].toString())
    .setChoiceValues(select_list_8)
    .setRequired(true);

  // ラジオボタン
  form.addMultipleChoiceItem()
    .setTitle(question_title[0][9].toString())
    .setChoiceValues(select_list_9)
    .setRequired(true);

  // ラジオボタン
  form.addMultipleChoiceItem()
    .setTitle(question_title[0][10].toString())
    .setChoiceValues(select_list_10)
    .setRequired(true);

  // ラジオボタン
  form.addMultipleChoiceItem()
    .setTitle(question_title[0][11].toString())
    .setChoiceValues(select_list_11)
    .setRequired(true);

  // ラジオボタン
  form.addMultipleChoiceItem()
    .setTitle(question_title[0][12].toString())
    .setChoiceValues(select_list_12)
    .setRequired(true);

  // フォームIDを取得する
  let formid = "";
  let targetfile = DriveApp.searchFiles('title contains "ひめみこ検定"');
  if (targetfile.hasNext()) {
    formid = targetfile.next().getId();
  } else {
    Logger.log("フォームIDを取得できませんでした。");
  }

  // トリガー設定
  ScriptApp.newTrigger("getFormAnswer")
    .forForm(formid)
    .onFormSubmit()
    .create();

}

function getFormAnswer (e) {

  // 回答データシートに回答データを集計する
  insertdata(e);

  // 答え合わせとメール送信
  checkanswer();

  // LINE通知
  sendmyline();
}

function insertdata (e) {
  // データ
  const answerdata = e.response.getItemResponses();
  // 対象シート
  const sheetobject = model.activespreadsheet.getSheetByName("回答データ");

  // A1セルが空白ならば、設問タイトルを挿入する
  if (sheetobject.getRange(1, 1).getValue() === "") {
    for (let j = 0; j < answerdata.length; j++) {
      let titledata = answerdata[j].getItem().getTitle();
      sheetobject.getRange(1, j + 1).setValue(titledata).setBackground("#00ff00");
      model.questionarray.push(titledata); // テスト
    }
  }

  let lastrow = sheetobject.getLastRow();

  for (let i = 0; i < answerdata.length; i++) {
    // 回答データを挿入する
    let setdata = answerdata[i].getResponse();
    sheetobject.getRange(lastrow + 1, i + 1).setValue(setdata);
    model.answerarray.push(setdata);
  }
}

function checkanswer () {
  let returnmsg = "こんにちは。\n" + model.answerarray[0] + " さん\n\n";
  let failcount = 0;

  for (let i = 3; i < model.answerarray.length; i++) {
    switch (model.answerarray[i]) {
      case "オオクノヒメミコ":
        break;
      case "名張夏見廃寺跡":
        break;
      case "岡山県":
        break;
      case "天武天皇":
        break;
      case "６首":
        break;
      case "大津皇子":
        break;
      case "斎王":
        break;
      case "伊勢":
        break;
      case "馬酔木":
        break;
      case "二上山":
        break;
      case "壬申の乱":
        break;
      case "薬師寺縁起":
        break;
      default:
        returnmsg += "第" + (i - 2) + "問 × \n";
        failcount++;
        break;
    }
  }

  let questioncount = model.answerarray.length - 3;
  let correctcount = questioncount - failcount;
  let correctrate = Math.floor(correctcount / questioncount * 100 * 10) / 10;

  if (failcount === 0) {
    returnmsg += "正答率: 100%\nおめでとうございます。全問正解です。"
  } else {
    returnmsg += "正答率: " + correctrate + "%\nまたの挑戦をお待ちしています。";
  }

  GmailApp.sendEmail(
    model.answerarray[1],
    "【ひめみこ運営事務局】ひめみこ検定(初級)　結果のご連絡",
    returnmsg
  );
}

function sendmyline () {
  const ACCESS_TOKEN = "7ypW4O0CAH5se2tWEYKQc8kmV9uGTI486mEhOPmA/cXRnmSETo9CssrzIg3kJTQLBZNVIzPV9gRboDlCov5ubBQ6lyTpHmRVI/YE0HyxuiPKDvqroZzvmbfk0zdjSJayaraZmh5dpt2C91CXWUuDdwdB04t89/1O/w1cDnyilFU=";
  const USER_ID = "U6474df320a2f549af9fc2603c75f7415";
  let url = "https://api.line.me/v2/bot/message/push";

  let textdata = {
    "to": USER_ID,
    "messages": [
      {
        "type": "text",
        "text": model.answerarray[0] + " さん から回答がありました。\n累計" + (model.activespreadsheet.getSheetByName("回答データ").getDataRange().getLastRow() - 1) + "件です。",
      }
    ],
  };

  let params = {
    "method": "post",
    "headers": {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + ACCESS_TOKEN,
    },
    "payload": JSON.stringify(textdata),
  };

  UrlFetchApp.fetch(url, params);
}

