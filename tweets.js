// ツイート内容の開始行番号
const baseRow = 7;

/**
 * スプレッドシート内のデータを取得してツイートを投稿する
 * */
function tweets(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const workSheet = ss.getSheetByName("tweets");

  let triggerType = "idMatched";
  let postRow = 0;
  if (e == undefined) {
    // ボタン実行の場合
    // トリガーIDが空白、日時が空白の行番号を取得する
    for (let i = 0; i < workSheet.getLastRow() - baseRow + 1; i++) {
      const date = workSheet.getRange(i + baseRow, 5).getValue();
      const time = workSheet.getRange(i + baseRow, 6).getValue();
      const triggerId = workSheet.getRange(i + baseRow, 7).getValue();
      if (date == "" && time == "" && triggerId == "") {
        postRow = i + baseRow;
        break
      }
    }
    if (postRow == 0) {
      console.log("ツイートする対象の行がありません。\n今すぐにツイートしたい場合は対象行の「予約日時」および「トリガーID」が空白になっていることを確認ください。")
      return;
    }

  } else {
    // トリガー自動実行の場合
    // トリガーIDが一致する行番号を取得する
    for (let i = 0; i < workSheet.getLastRow() - baseRow + 1; i++) {
      const triggerId = workSheet.getRange(i + baseRow, 7).getValue();
      if (triggerId == e.triggerUid) {
        postRow = i + baseRow;
        break
      }
    }
    if (postRow == 0) {
      console.log("トリガーIDに一致する行番号が存在しません。");
      // トリガーIDが空白、日時が空白の行番号を取得する
      for (let i = 0; i < workSheet.getLastRow() - baseRow + 1; i++) {
        const date = workSheet.getRange(i + baseRow, 5).getValue();
        const time = workSheet.getRange(i + baseRow, 6).getValue();
        const triggerId = workSheet.getRange(i + baseRow, 7).getValue();
        if (date == "" && time == "" && triggerId == "") {
          postRow = i + baseRow;
          break
        }
      }
      if (postRow == 0) {
        console.log("ツイートする対象の行がありません。")
        return;
      } else {
        triggerType = "definedDate"
      }
    }
  }
  // 文言、タグ、URLを取得
  const word = workSheet.getRange(postRow, 1).getValue();
  const tagValue = workSheet.getRange(postRow, 2).getValue();
  let tags = "";
  if (tagValue != "") {
    tags = decorateTag(tagValue);
  }
  const url = workSheet.getRange(postRow, 3).getValue();
  if (word == "" && tagValue == "" && url == "") {
    console.log("ツイートする内容がありません");
    return;
  }
  const tweet = word + "\n" + tags + "\n" + url

  const imageUrl = workSheet.getRange(postRow, 4).getValue();
  // const imageUrl = "";
  if (imageUrl == "") {
    postTweet(tweet);
  } else {
    postTweetWithImage(tweet, imageUrl);
  }
  // トリガー実行の場合、実行済のトリガーを削除する
  if (e != undefined && triggerType != 'definedDate') {
    deleteTrigger(e.triggerUid);
  }
  console.log("postRow：" + postRow + "getLastRow():" + workSheet.getLastRow());
  // baseRow以降の行で1列目が空白でない行の数を取得する
  const data = workSheet.getRange(baseRow, 1, workSheet.getLastRow() - baseRow+1, 1).getValues();
  let not_null_count = 0;
  while (data[not_null_count][0] !== ""){
    not_null_count++;
    console.log("not_null_count：" + not_null_count);
    if(not_null_count === data.length){
      break;
    }
  }
  console.log("not_null_count：" + not_null_count);
  // ツイート済みを最後の次の行にコピーする
  workSheet.getRange(postRow, 1, 1, 4).copyTo(workSheet.getRange(baseRow + not_null_count, 1, 1, 4));
  // ツイート済みを削除
  workSheet.deleteRow(postRow);
  // 空白行を1行追加
  workSheet.insertRows(workSheet.getMaxRows());
}

/**
 * 投稿する明日以降の日付と時刻を設定する
 * */
function settingPostDate(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const workSheet = ss.getSheetByName("tweets");

  const data = workSheet.getRange(baseRow, 1, workSheet.getLastRow() - baseRow+1, 1).getValues();
  console.log(data);
  let not_null_count = 0;
  while (data[not_null_count][0] !== ""){
    not_null_count++;
    console.log("not_null_count：" + not_null_count);
    if(not_null_count === data.length){
      break;
    }
  }

  const days = not_null_count; // N日後
  const today = new Date();
  const tomorrow = new Date(today.getTime() + (24 * 60 * 60 * 1000));
  const dateFormat = "yyyy-MM-dd";

  for (let i = 0; i < days; i++) {
    const date = new Date(tomorrow.getTime() + (i * 24 * 60 * 60 * 1000));
    const formattedDate = Utilities.formatDate(date, "JST", dateFormat);
    workSheet.getRange(baseRow + i, 5, 1, 2).setValues([[formattedDate, "21:00"]]);
  }
  //console.log(dateArray);

}

/**
 * タグにする
 * */
function decorateTag(words) {
  const words0 = "#" + words;
  const array = words0.split(",");
  return array.join(" #");
}
/**
 * ツイートを投稿する
 * */
function postTweet(tweet) {
  // ツイートするAPIリクエスト
  const endpointUri = "https://api.twitter.com/2/tweets";
  const payload = {
    "text": tweet
  };
  const options = {
    'method': 'post',
    'payload': JSON.stringify(payload),
    'contentType': 'application/json',
    'muteHttpExceptions': true
  };
  const postResult = makeRequest(endpointUri, options);
  console.log(postResult);
}

/**
 * ツイートを画像付きで投稿する
 * */
function postTweetWithImage(tweet, imageUrl) {
  const targetFolder = findOrCreateFolder("images");
  const imageFile = downloadImageToDrive(imageUrl, targetFolder);
  const fileId = imageFile.getId();
  const fileByApp = DriveApp.getFileById(fileId);
  // ファイルのデータをbase64形式に変換
  const base64Data = Utilities.base64Encode(fileByApp.getBlob().getBytes());

  // 画像アップロードのAPIを実行
  const uploadUrl = 'https://upload.twitter.com/1.1/media/upload.json'

  const uploadParam = {
    method: "post",
    payload: { media_data: base64Data } // base64形式なので media_data にデータを入れる
  }
  const uploadResult = makeRequest(uploadUrl, uploadParam);
  console.log(uploadResult);
  const mediaId = uploadResult.media_id_string;

  // アップロードしたファイルを添付して投稿
  const postUrl = "https://api.twitter.com/2/tweets";
  const payload = {
    "text": tweet,
    "media": {
      media_ids: [mediaId] // カンマ区切りで書く
    }
  };
  const postParam = {
    method: 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload)
  }

  const postResult = makeRequest(postUrl, postParam);
  console.log(postResult);
}


/**
 * Web上の画像をGoogleドライブ保存する
 * */
function downloadImageToDrive(imageUrl, saveFolder = undefined) {
  // imageUrl を指定してファイルを受け取る
  const imageFile = UrlFetchApp.fetch(imageUrl).getBlob();

  if (saveFolder == undefined) {
    console.log("Download: " + imageUrl + "\n => マイドライブ");
    const rootFolder = DriveApp.getRootFolder();
    // 新規ファイルとして保存
    return rootFolder.createFile(imageFile);
  } else {
    console.log("Download: " + imageUrl + "\n => " + saveFolder.getName());
    // 新規ファイルとして保存
    return saveFolder.createFile(imageFile);
  }
}

/**
 * Web上の画像をGoogleドライブ保存する
 * */
function findOrCreateFolder(folderName, parentFolderId = undefined) {
  // 指定 ID のフォルダを取得する
  // (ID はそのフォルダの URL https://drive.google.com/drive/folders/XXXX の XXXX の部分がフォルダ ID)
  let folder = ""
  if (parentFolderId == undefined) {
    folder = DriveApp.getRootFolder();
  } else {
    folder = DriveApp.getFolderById(FOLDER_ID);
  }
  const itr = folder.getFoldersByName(folderName);
  if (itr.hasNext()) {
    // フォルダが見つかった場合はそれを返す
    return itr.next();
  } else {
    // フォルダが見つからなかった場合は作成して返す
    const newFolder = folder.createFolder(folderName);
    newFolder.setName(folderName);
    return newFolder;
  }
}


/**
 * 特定日時にツイートするようにトリガーを設定する
 * */
function createTriggers() {
  //セルを取得
  const workSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const max_reservation = 20; // 予約できる件数は20件まで
  // 予約日付と時間が埋まっている箇所のトリガーを設定する
  let i = 0;
  while(i != workSheet.getLastRow() - baseRow + 1 || i != max_reservation) {
    const previousTriggered = workSheet.getRange(i + baseRow, 7).getValue()
    if (previousTriggered != "") {
      // トリガーID列が空白でない場合（すでにトリガーが設定されている場合）
      continue;
    }
    const date = workSheet.getRange(i + baseRow, 5).getValue();
    let time = workSheet.getRange(i + baseRow, 6).getDisplayValue();
    if (date != "") {
      if (time == "") {
        time = "00:00"
      }
      const timeArray = time.split(":");
      const hour = timeArray[0];
      const min = timeArray[1];
      date.setHours(hour);
      date.setMinutes(min);
      // 特定日時でトリガー登録
      const trigger = ScriptApp.newTrigger('tweets').timeBased().at(date).create();
      const triggerId = trigger.getUniqueId();
      workSheet.getRange(i + baseRow, 7).setValue(triggerId);
      console.log("トリガーを設定しました。トリガーID：" + triggerId + "\n" + date);
    }
    i++;
  }
}

/**
 * トリガーを含む予約設定を一括削除し、スプレッドシートからも削除する
 * */
function deleteAllTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const workSheet = ss.getSheetByName("tweets");
  const allTriggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < allTriggers.length; i++) {
    ScriptApp.deleteTrigger(allTriggers[i]);
    console.log("トリガーを削除しました。トリガーID：" + allTriggers[i]);
    workSheet.getRange(baseRow+i, 5, 1,3).clearContent();
  }
}

/**
 * トリガーを削除する
 * */
function deleteTrigger(triggerId) {
  const allTriggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < allTriggers.length; i++) {
    if (allTriggers[i].getUniqueId() == triggerId) {
      ScriptApp.deleteTrigger(allTriggers[i]);
      console.log("トリガーを削除しました。トリガーID：" + allTriggers[i]);
      break;
    }
  }
}