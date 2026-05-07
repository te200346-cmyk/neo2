/**
 * @license
 * このソフトウェアは、MITライセンスのもとで公開されています。
 * This software is released under the MIT License.
 * Copyright (c) 2024 Masaaki Maeta
 */

// 共通設定
const PARENT_FOLDER_NAME = "メディア保存フォルダ";
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const SUBFOLDER_SHEET_NAME = "シート1";
const SUBFOLDER_CELL = "B1";
const HISTORY_SHEET_NAME = "履歴";
const HISTORY_HEADERS = ["ファイル名", "保存日時", "フォルダパス", "ファイルリンク", "ファイル形式", "メディアタイプ"];

/**
 * Webアプリのエントリーポイント
 */
/**
 * Webアプリのエントリーポイント (動作確認用)
 */
function doGet(e) {
  return ContentService.createTextOutput('スーパー記録くん API 稼働中')
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * 外部からのAPIリクエストを処理 (POST)
 */
function doPost(e) {
  let response;
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('リクエストデータが空です。');
    }

    const requestData = JSON.parse(e.postData.contents);
    const functionName = requestData.function;
    const args = requestData.args || [];

    // 実行を許可する関数のリスト
    const allowedFunctions = [
      'saveAudioFile',
      'saveVideoFile',
      'savePhotoFile',
      'saveDrawingFile',
      'saveTextFile',
      'getModeSettings'
    ];

    if (allowedFunctions.indexOf(functionName) === -1) {
      throw new Error('未許可の関数呼び出しです: ' + functionName);
    }

    // 関数を明示的に呼び出す（最も確実な方法）
    let result;
    if (functionName === 'saveAudioFile') result = saveAudioFile.apply(null, args);
    else if (functionName === 'saveVideoFile') result = saveVideoFile.apply(null, args);
    else if (functionName === 'savePhotoFile') result = savePhotoFile.apply(null, args);
    else if (functionName === 'saveDrawingFile') result = saveDrawingFile.apply(null, args);
    else if (functionName === 'saveTextFile') result = saveTextFile.apply(null, args);
    else if (functionName === 'getModeSettings') result = getModeSettings.apply(null, args);
    else throw new Error('実行対象の関数が定義されていません: ' + functionName);

    response = result;

  } catch (error) {
    console.error(error);
    response = {
      success: false,
      message: 'APIエラー: ' + error.toString()
    };
  }

  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * フォルダを取得または作成
 */
function getOrCreateFolderIdByName(folderName, parentFolder = DriveApp.getRootFolder()) {
  if (!folderName || typeof folderName !== 'string' || folderName.trim() === '') {
    throw new Error("有効なフォルダ名が指定されていません。");
  }

  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next().getId();
  } else {
    try {
      const newFolder = parentFolder.createFolder(folderName);
      Logger.log(`フォルダ "${folderName}" (ID: ${newFolder.getId()}) を作成しました。`);
      return newFolder.getId();
    } catch (error) {
      Logger.log(`フォルダ "${folderName}" の作成に失敗しました: ${error.toString()}`);
      throw new Error(`フォルダ "${folderName}" の作成に失敗しました。`);
    }
  }
}

/**
 * スプレッドシートからサブフォルダ名を取得
 */
function getSubFolderNameFromSheet() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SUBFOLDER_SHEET_NAME);
    if (!sheet) {
      Logger.log(`シート "${SUBFOLDER_SHEET_NAME}" が見つかりません。`);
      return null;
    }

    const subFolderName = sheet.getRange(SUBFOLDER_CELL).getValue().toString().trim();
    if (!subFolderName) {
      Logger.log(`セル "${SUBFOLDER_CELL}" にサブフォルダ名が入力されていません。`);
      return null;
    }

    return subFolderName.replace(/[\\\/:\*\?"<>\|]/g, '_');
  } catch (e) {
    Logger.log(`スプレッドシートからのサブフォルダ名取得エラー: ${e.toString()}`);
    return null;
  }
}

/**
 * 履歴シートの作成と初期化
 */
function createHistorySheetIfNotExists() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(HISTORY_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(HISTORY_SHEET_NAME);
    sheet.appendRow(HISTORY_HEADERS);
    sheet.getRange(1, 1, 1, HISTORY_HEADERS.length).setFontWeight("bold");

    // 列幅の設定
    sheet.setColumnWidth(1, 250); // ファイル名
    sheet.setColumnWidth(2, 150); // 保存日時
    sheet.setColumnWidth(3, 250); // フォルダパス
    sheet.setColumnWidth(4, 300); // ファイルリンク
    sheet.setColumnWidth(5, 100); // ファイル形式
    sheet.setColumnWidth(6, 100); // メディアタイプ

    Logger.log(`履歴シート "${HISTORY_SHEET_NAME}" を作成しました。`);
  } else {
    // 既存シートの列確認と追加
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    // メディアタイプ列がない場合は追加
    if (!headers.includes("メディアタイプ")) {
      const newCol = headers.length + 1;
      sheet.getRange(1, newCol).setValue("メディアタイプ");
      sheet.getRange(1, newCol).setFontWeight("bold");
      sheet.setColumnWidth(newCol, 100);
    }

    // ファイル形式列がない場合は追加
    if (!headers.includes("ファイル形式")) {
      const newCol = headers.includes("メディアタイプ") ? headers.length : headers.length + 1;
      sheet.getRange(1, newCol).setValue("ファイル形式");
      sheet.getRange(1, newCol).setFontWeight("bold");
      sheet.setColumnWidth(newCol, 100);
    }
  }

  return sheet;
}

/**
 * 履歴シートに記録を追加
 */
function addRecordToHistorySheet(fileName, folderPathText, folderUrl, fileUrl, fileFormat, mediaType) {
  try {
    const sheet = createHistorySheetIfNotExists();
    const timestamp = new Date();
    const formattedTimestamp = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");

    // ハイパーリンクの作成
    const folderLinkFormula = `=HYPERLINK("${folderUrl}","${folderPathText}")`;
    const fileLinkFormula = `=HYPERLINK("${fileUrl}","${fileName}")`;

    sheet.appendRow([
      fileName,
      formattedTimestamp,
      folderLinkFormula,
      fileLinkFormula,
      fileFormat.toUpperCase(),
      mediaType
    ]);

    Logger.log(`履歴を記録しました: ${fileName}, ${mediaType}, ${fileFormat}`);
  } catch (e) {
    Logger.log(`履歴シートへの記録エラー: ${e.toString()}`);
  }
}

/**
 * 音声ファイル（MP3）を保存
 */
function saveAudioFile(audioDataUrl, baseFileName) {
  try {
    if (!audioDataUrl || typeof audioDataUrl !== 'string') {
      throw new Error("音声データ (Data URL) が無効です。");
    }
    if (!baseFileName || typeof baseFileName !== 'string' || baseFileName.trim() === '') {
      throw new Error("ファイル名が無効です。");
    }

    const parentFolderId = getOrCreateFolderIdByName(PARENT_FOLDER_NAME);
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const subFolderNameRaw = getSubFolderNameFromSheet();

    let targetFolder;
    let folderPathText;
    let targetFolderUrl;

    if (subFolderNameRaw) {
      const subFolderId = getOrCreateFolderIdByName(subFolderNameRaw, parentFolder);
      targetFolder = DriveApp.getFolderById(subFolderId);
      folderPathText = `${parentFolder.getName()} > ${targetFolder.getName()}`;
      targetFolderUrl = targetFolder.getUrl();
    } else {
      targetFolder = parentFolder;
      folderPathText = parentFolder.getName();
      targetFolderUrl = parentFolder.getUrl();
      Logger.log(`サブフォルダ名が取得できなかったため、親フォルダに保存します。`);
    }

    const parts = audioDataUrl.match(/^data:(.+?);base64,(.+)$/);
    if (!parts || parts.length !== 3) {
      throw new Error("無効なData URL形式です。");
    }

    const mimeType = parts[1];
    const base64Data = parts[2];

    const finalFileName = baseFileName.toLowerCase().endsWith('.mp3')
      ? baseFileName
      : `${baseFileName}.mp3`;

    const decodedData = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decodedData, mimeType, finalFileName);
    const file = targetFolder.createFile(blob);
    const fileUrl = file.getUrl();

    // 履歴シートに記録（音声タイプを明記）
    addRecordToHistorySheet(finalFileName, folderPathText, targetFolderUrl, fileUrl, "MP3", "音声");

    Logger.log(`音声ファイル "${finalFileName}" を保存しました。`);

    return {
      success: true,
      message: `音声ファイル "${finalFileName}" をフォルダ「${folderPathText}」に保存しました。`,
      fileId: file.getId(),
      fileName: finalFileName,
      fileUrl: fileUrl
    };
  } catch (error) {
    Logger.log(`saveAudioFileでエラーが発生しました: ${error.toString()}`);
    return {
      success: false,
      message: `音声ファイルの保存中にエラーが発生しました。`
    };
  }
}

/**
 * 動画ファイル（MP4/WebM）を保存
 */
function saveVideoFile(videoDataUrl, baseFileName, extension = 'webm', mimeType = 'video/webm') {
  try {
    if (!videoDataUrl || !baseFileName) {
      throw new Error("動画データまたはファイル名が無効です。");
    }

    // 拡張子の検証
    const validExtensions = ['mp4', 'webm'];
    if (!validExtensions.includes(extension.toLowerCase())) {
      extension = 'webm';
    }

    const parentFolderId = getOrCreateFolderIdByName(PARENT_FOLDER_NAME);
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const subFolderNameRaw = getSubFolderNameFromSheet();

    let targetFolder;
    let folderPathText;
    let targetFolderUrl;

    if (subFolderNameRaw) {
      const subFolderId = getOrCreateFolderIdByName(subFolderNameRaw, parentFolder);
      targetFolder = DriveApp.getFolderById(subFolderId);
      folderPathText = `${parentFolder.getName()} > ${targetFolder.getName()}`;
      targetFolderUrl = targetFolder.getUrl();
    } else {
      targetFolder = parentFolder;
      folderPathText = parentFolder.getName();
      targetFolderUrl = parentFolder.getUrl();
    }

    const parts = videoDataUrl.match(/^data:(.+?);base64,(.+)$/);
    if (!parts) throw new Error("無効なData URL形式です。");

    const detectedMimeType = parts[1];
    const base64Data = parts[2];

    const finalFileName = `${baseFileName}.${extension.toLowerCase()}`;
    const decodedData = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decodedData, mimeType || detectedMimeType, finalFileName);
    const file = targetFolder.createFile(blob);

    // 履歴に記録（動画タイプを明記）
    addRecordToHistorySheet(finalFileName, folderPathText, targetFolderUrl, file.getUrl(), extension, "動画");

    return {
      success: true,
      message: `動画ファイル "${finalFileName}" (${extension.toUpperCase()}形式) をフォルダ「${folderPathText}」に保存しました。`
    };
  } catch (error) {
    Logger.log(`saveVideoFileでエラーが発生しました: ${error.toString()}`);
    return {
      success: false,
      message: `動画ファイルの保存中にエラーが発生しました。`
    };
  }
}

/**
 * 写真ファイル（PNG/JPG）を保存
 */
function savePhotoFile(photoDataUrl, baseFileName, extension = 'jpg') {
  try {
    if (!photoDataUrl || !baseFileName) {
      throw new Error("写真データまたはファイル名が無効です。");
    }

    // 拡張子の検証
    const validExtensions = ['jpg', 'jpeg', 'png'];
    if (!validExtensions.includes(extension.toLowerCase())) {
      extension = 'jpg';
    }

    const parentFolderId = getOrCreateFolderIdByName(PARENT_FOLDER_NAME);
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const subFolderNameRaw = getSubFolderNameFromSheet();

    let targetFolder;
    let folderPathText;
    let targetFolderUrl;

    if (subFolderNameRaw) {
      const subFolderId = getOrCreateFolderIdByName(subFolderNameRaw, parentFolder);
      targetFolder = DriveApp.getFolderById(subFolderId);
      folderPathText = `${parentFolder.getName()} > ${targetFolder.getName()}`;
      targetFolderUrl = targetFolder.getUrl();
    } else {
      targetFolder = parentFolder;
      folderPathText = parentFolder.getName();
      targetFolderUrl = parentFolder.getUrl();
    }

    const parts = photoDataUrl.match(/^data:(.+?);base64,(.+)$/);
    if (!parts) throw new Error("無効なData URL形式です。");

    const detectedMimeType = parts[1];
    const base64Data = parts[2];

    const finalFileName = `${baseFileName}.${extension.toLowerCase()}`;
    const decodedData = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decodedData, detectedMimeType, finalFileName);
    const file = targetFolder.createFile(blob);

    // 履歴に記録（写真タイプを明記）
    const formatDisplay = extension.toUpperCase() === 'JPEG' ? 'JPG' : extension.toUpperCase();
    addRecordToHistorySheet(finalFileName, folderPathText, targetFolderUrl, file.getUrl(), formatDisplay, "写真");

    return {
      success: true,
      message: `写真ファイル "${finalFileName}" (${formatDisplay}形式) をフォルダ「${folderPathText}」に保存しました。`
    };
  } catch (error) {
    Logger.log(`savePhotoFileでエラーが発生しました: ${error.toString()}`);
    return {
      success: false,
      message: `写真ファイルの保存中にエラーが発生しました。`
    };
  }
}

/**
 * お絵かきファイル（PNG）を保存
 */
function saveDrawingFile(drawingDataUrl, baseFileName) {
  try {
    if (!drawingDataUrl || !baseFileName) {
      throw new Error("お絵かきデータまたはファイル名が無効です。");
    }

    const parentFolderId = getOrCreateFolderIdByName(PARENT_FOLDER_NAME);
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const subFolderNameRaw = getSubFolderNameFromSheet();

    let targetFolder;
    let folderPathText;
    let targetFolderUrl;

    if (subFolderNameRaw) {
      const subFolderId = getOrCreateFolderIdByName(subFolderNameRaw, parentFolder);
      targetFolder = DriveApp.getFolderById(subFolderId);
      folderPathText = `${parentFolder.getName()} > ${targetFolder.getName()}`;
      targetFolderUrl = targetFolder.getUrl();
    } else {
      targetFolder = parentFolder;
      folderPathText = parentFolder.getName();
      targetFolderUrl = parentFolder.getUrl();
    }

    const parts = drawingDataUrl.match(/^data:(image\/png);base64,(.+)$/);
    if (!parts) throw new Error("無効なData URL形式です。PNG形式である必要があります。");

    const mimeType = parts[1];
    const base64Data = parts[2];

    const finalFileName = `${baseFileName}.png`;
    const decodedData = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decodedData, mimeType, finalFileName);
    const file = targetFolder.createFile(blob);

    // 履歴に記録（お絵かきタイプを明記）
    addRecordToHistorySheet(finalFileName, folderPathText, targetFolderUrl, file.getUrl(), "PNG", "お絵かき");

    return {
      success: true,
      message: `お絵かきファイル "${finalFileName}" (PNG形式) をフォルダ「${folderPathText}」に保存しました。`
    };
  } catch (error) {
    Logger.log(`saveDrawingFileでエラーが発生しました: ${error.toString()}`);
    return {
      success: false,
      message: `お絵かきファイルの保存中にエラーが発生しました。`
    };
  }
}

/**
 * テキストファイル（TXT）を保存
 */
function saveTextFile(textData, baseFileName) {
  try {
    if (!textData || typeof textData !== 'string') {
      throw new Error("テキストデータが無効です。");
    }
    if (!baseFileName || typeof baseFileName !== 'string' || baseFileName.trim() === '') {
      throw new Error("ファイル名が無効です。");
    }

    const parentFolderId = getOrCreateFolderIdByName(PARENT_FOLDER_NAME);
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const subFolderNameRaw = getSubFolderNameFromSheet();

    let targetFolder;
    let folderPathText;
    let targetFolderUrl;

    if (subFolderNameRaw) {
      const subFolderId = getOrCreateFolderIdByName(subFolderNameRaw, parentFolder);
      targetFolder = DriveApp.getFolderById(subFolderId);
      folderPathText = `${parentFolder.getName()} > ${targetFolder.getName()}`;
      targetFolderUrl = targetFolder.getUrl();
    } else {
      targetFolder = parentFolder;
      folderPathText = parentFolder.getName();
      targetFolderUrl = parentFolder.getUrl();
    }

    const finalFileName = `${baseFileName}.txt`;
    const file = targetFolder.createFile(finalFileName, textData, MimeType.PLAIN_TEXT);
    const fileUrl = file.getUrl();

    // 履歴シートに記録
    addRecordToHistorySheet(finalFileName, folderPathText, targetFolderUrl, fileUrl, "TXT", "テキスト");

    return {
      success: true,
      message: `テキストファイル "${finalFileName}" をフォルダ「${folderPathText}」に保存しました。`
    };
  } catch (error) {
    Logger.log(`saveTextFileでエラーが発生しました: ${error.toString()}`);
    return {
      success: false,
      message: `テキストファイルの保存中にエラーが発生しました。`
    };
  }
}

/**
 * テスト用関数
 */
function test_UnifiedMediaApp() {
  try {
    createHistorySheetIfNotExists();
    Logger.log("履歴シートの確認/作成完了");

    const parentFolderId = getOrCreateFolderIdByName(PARENT_FOLDER_NAME);
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    Logger.log(`親フォルダ: "${parentFolder.getName()}" URL: ${parentFolder.getUrl()}`);

    const subFolderName = getSubFolderNameFromSheet();
    if (subFolderName) {
      Logger.log(`サブフォルダ名: "${subFolderName}"`);
    } else {
      Logger.log("サブフォルダ名が設定されていません");
    }

    // テスト用の履歴記録
    const testAudioFile = "test_audio.mp3";
    const testVideoFile = "test_video.mp4";
    const dummyUrl = "https://drive.google.com/file/d/dummy/view";

    addRecordToHistorySheet(testAudioFile, PARENT_FOLDER_NAME, parentFolder.getUrl(), dummyUrl, "MP3", "音声");
    addRecordToHistorySheet(testVideoFile, PARENT_FOLDER_NAME, parentFolder.getUrl(), dummyUrl, "MP4", "動画");

    Logger.log("テスト完了。履歴シートを確認してください。");
  } catch (e) {
    Logger.log(`テストエラー: ${e.toString()}`);
  }
}

/**
 * モード設定を取得
 */
function getModeSettings() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SUBFOLDER_SHEET_NAME); // "シート1"
    if (!sheet) {
      // Default to all visible if sheet not found
      return { audio: true, video: true, photo: true, drawing: true, text: true, error: `Sheet "${SUBFOLDER_SHEET_NAME}" not found.` };
    }

    const range = sheet.getRange("E1:E5");
    const values = range.getValues().flat(); // [E1, E2, E3, E4, E5]

    let audioVisible = values[0] == 1;
    let videoVisible = values[1] == 1;
    let photoVisible = values[2] == 1;
    let drawingVisible = values[3] == 1;
    let textVisible = values[4] == 1;

    // If all are hidden, force audio to be visible
    if (!audioVisible && !videoVisible && !photoVisible && !drawingVisible && !textVisible) {
      sheet.getRange("E1").setValue(1);
      audioVisible = true;
    }

    return {
      audio: audioVisible,
      video: videoVisible,
      photo: photoVisible,
      drawing: drawingVisible,
      text: textVisible
    };
  } catch (e) {
    Logger.log(`getModeSettings error: ${e.toString()}`);
    // On error, default to all visible
    return { audio: true, video: true, photo: true, drawing: true, text: true, error: e.toString() };
  }
}