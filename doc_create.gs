
//スプレッドシートオープン時に実行される関数（メニュー追加）
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  //スプレッドシートのメニューにカスタムメニュー「自動処理」を作成（他の関数を実行するメニューを作成）
  var subMenus = [];
  subMenus.push({name: 'ドキュメント生成', functionName: 'createDocument'});
  ss.addMenu('ドキュメント生成実行', subMenus);
}

//担当者のセルを選択時に実行すると、ドキュメントが生成される関数
function createDocument() {
  //スプレッドシートを取得、シートを取得
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  
  //編集されたCell情報を取得
  var activeCell = sheet.getActiveCell();
  var activeCellValue = activeCell.getValue();
  var activeRow = activeCell.getRow();
  var activeColumn = activeCell.getColumn();
  
  //アクティブセル（編集されたセル）が担当者の列以外の場合はスクリプトを終了
  //※[列番号]は、A列なら1、Bなら2
  if (activeColumn != [2]) {
    Logger.log("担当者の列以外で実行されたため処理を終了しました。");
    Browser.msgBox("担当者の列以外で実行されたため処理を終了しました。");
    return;
  }
  //アクティブセルが空欄の場合はスクリプトを終了
  if (activeCellValue == "" || activeCellValue == "担当者") {
    Logger.log("担当者の列が空であるため、処理を終了しました。");
    Browser.msgBox("担当者の列が空であるため、処理を終了しました。");
    return;
  }
  //URLが既に記入されている場合はスクリプトを終了
  if(sheet.getRange(activeRow, [3]).getValue() != ""){
    Logger.log("担当者の列のURLが記入済みであるため、処理を終了しました。");
    Browser.msgBox("担当者の列のURLが記入済みであるため、処理を終了しました。");
    Logger.log(sheet.getRange(activeRow, [3]).getValue());
    return;
  }
  
  //ドキュメントのタイトル
  var docName = "原稿_"+ activeCellValue;
     
  //テンプレートファイル読み込み
  var template = DriveApp.getFileById("ここにドキュメントID");
  
  //テンプレートファイルをコピー
  var document = template.makeCopy(docName);
    
  //ドキュメントを格納するフォルダを取得（ブログ原稿フォルダ配下）
  var targetFolder = DriveApp.getFolderById("ここにフォルダID");
  
  //指定したフォルダに所属（移動）させる
  var docFile = DriveApp.getFileById(document.getId());
  targetFolder.addFile(docFile);
  Logger.log("ドキュメントID "+document.getId()+"を作成しました");
  Browser.msgBox("ドキュメントID "+document.getId()+"を作成しました");
  
  //作成したドキュメントURLを取得
  var documentUrl = document.getUrl();
  
  //シートにURLを書き込む
  sheet.getRange(activeRow, [3]).setValue(documentUrl);
  
  // スプレッドシート「Google アカウント管理」の読み込み
  var spreadSheetById = SpreadsheetApp.openById('ここにスプレッドシートID');
  var sheetByName = spreadSheetById.getSheetByName("シート1");
  
  // シートの中から選択社員を検索
  var textFinder = sheetByName.createTextFinder(activeCellValue);
  var ranges = textFinder.findAll();
  
  // 選択社員のGoogleアカウントが記入されていない場合
  if(ranges.length == 0){
    Logger.log(activeCellValue+"さんのGoogleアカウントは「Google アカウント管理」シートに記入されていません。権限を付与できません。");
    Browser.msgBox(activeCellValue+"さんのGoogleアカウントは「Google アカウント管理」シートに記入されていません。権限を付与できません。");
  }
  else{
    // 検索結果の2つ右隣の値をmailに格納
    var mail = ranges[0].offset(0,2).getValue();
    driveadd(mail, document.getId());  
    Logger.log(mail+"に編集権限を付与しました");
    Browser.msgBox(mail+"に編集権限を付与しました");
  
    driveadd_E_HOM(document.getId());
    Logger.log("ブログ管理メンバーに編集権限を付与しました");
    Browser.msgBox("ブログ管理メンバーに編集権限を付与しました");
  }
}

//特定のドキュメントに共有権限(編集)を付与する関数
function driveadd(pMail, pFileid){
  var file = DriveApp.getFileById(pFileid);
  file.addEditor(pMail);
}

//E-HOMのメンバーに共有権限（編集）を付与する関数
function driveadd_E_HOM(pFileid){
  var file = DriveApp.getFileById(pFileid);
  
  //ブログ管理メンバーのアドレス取得    
  file.addEditor("ここにアドレス");
  
}
