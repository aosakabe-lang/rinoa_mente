/**
 * 株式会社リノア（REN）太陽光発電点検報告システム - GASバックエンド v4.0
 */

const SHEET_NAME = '点検データ';
const PCS_SHEET_NAME = 'PCS点検データ';

const HEADERS = [
  'ID', '点検日時', '担当者', '施主名', '現場住所', 'メール', '判定', 'コメント',
  '全体写真', 'パワコン写真', '異常写真', 'パネル写真', '金具写真', '脚部写真', '裏面写真',
  'ステータス', 'PDF_URL', '電圧区分', '発電所名', '電気主任技術者', '電力会社', '外気温度(℃)', '室内温度(℃)', '天気'
];

const PCS_HEADERS = [
  '点検ID', 'PCS番号', '型式', '製造番号', '端子増し締め', '瞬間発電量(kW)', '積算発電力量(kWh)',
  '絶縁抵抗値(MΩ)', '動作電圧(V)', '動作電流(A)', '不良備考メモ'
];

/**
 * 【必須】Webアプリを表示するための関数
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('REN | 太陽光発電設備 点検報告システム v4.0')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 点検報告データを保存するメイン処理
 */
function saveInspectionData(form) {
  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(30000);
    const id = 'REN-' + Utilities.getUuid().substring(0, 8).toUpperCase() + '-2026';
    
    // フォルダ取得（無い場合は自動作成）
    const folder = getOutputFolder_();

    // 写真のドライブ保存
    const photoMainUrl = saveImage_(form.photoMainData, id + '_01全体.jpg', folder);
    const photoPcsUrl = saveImage_(form.photoPcsData, id + '_02パワコン.jpg', folder);
    const photoSubUrl = saveImage_(form.photoSubData, id + '_03異常.jpg', folder);
    const photoPanelUrl = saveImage_(form.photoPanelData, id + '_04パネル.jpg', folder);
    const photoClampUrl = saveImage_(form.photoClampData, id + '_05金具.jpg', folder);
    const photoLegUrl = saveImage_(form.photoLegData, id + '_06脚部.jpg', folder);
    const photoBackUrl = saveImage_(form.photoBackData, id + '_07裏面.jpg', folder);

    // シートへデータ書き込み
    const sheet = getInspectionSheet_();
    sheet.appendRow([
      id, new Date(), form.worker, form.client, form.address, form.email || '', form.status,
      form.comment || '', photoMainUrl, photoPcsUrl, photoSubUrl, photoPanelUrl, photoClampUrl, photoLegUrl, photoBackUrl,
      '下書き', '', form.voltageType, form.plantName || '', form.chiefEngineer || '', form.powerCompany || '',
      form.outdoorTemp || '', form.indoorTemp || '', form.weather || ''
    ]);

    // PCSデータの書き込み
    if (form.pcsList && form.pcsList.length > 0) {
      const pcsSheet = getPcsSheet_();
      const pcsRows = form.pcsList.map(function (pcs) {
        return [
          id, pcs.pcsNo || '', pcs.model || '', pcs.serialNo || '', pcs.tightening || '済',
          pcs.instantPower || '', pcs.cumulativePower || '', pcs.insulationResistance || '100',
          pcs.operatingVoltage || '', pcs.operatingCurrent || '', pcs.note || ''
        ];
      });
      pcsSheet.getRange(pcsSheet.getLastRow() + 1, 1, pcsRows.length, PCS_HEADERS.length).setValues(pcsRows);
    }

    SpreadsheetApp.flush();
    return { success: true, message: '点検報告データを正常に保存しました。', id: id };
  } catch (error) {
    throw new Error('保存エラー: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * 【必須】出力先Googleドライブフォルダの取得・自動作成関数
 */
function getOutputFolder_() {
  const properties = PropertiesService.getScriptProperties();
  let folderId = properties.getProperty('OUTPUT_FOLDER_ID');
  
  if (!folderId) {
    const folder = DriveApp.createFolder('REN_太陽光点検報告書');
    folderId = folder.getId();
    properties.setProperty('OUTPUT_FOLDER_ID', folderId);
  }
  
  return DriveApp.getFolderById(folderId);
}

/**
 * 【必須】点検データシートの取得
 */
function getInspectionSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME) || forceResetSpreadsheet_();
}

/**
 * 【必須】PCSデータシートの取得
 */
function getPcsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(PCS_SHEET_NAME) || forceResetPcsSheet_();
}

/**
 * 【必須】シート初期設定関数（メイン）
 */
function forceResetSpreadsheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  headerRange.setValues([HEADERS])
    .setBackground('#0f172a').setFontColor('#38bdf8').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 35);
  sheet.setFrozenRows(1);
  return sheet;
}

/**
 * 【必須】シート初期設定関数（PCS）
 */
function forceResetPcsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(PCS_SHEET_NAME) || ss.insertSheet(PCS_SHEET_NAME);
  const headerRange = sheet.getRange(1, 1, 1, PCS_HEADERS.length);
  headerRange.setValues([PCS_HEADERS])
    .setBackground('#0f172a').setFontColor('#34d399').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 35);
  sheet.setFrozenRows(1);
  return sheet;
}

/**
 * 【必須】画像保存用のヘルパー関数
 */
function saveImage_(dataUrl, name, folder) {
  if (!dataUrl) return '';
  const match = String(dataUrl).match(/^data:(image\/(?:jpeg|png|webp));base64,([A-Za-z0-9+/=]+)$/);
  if (!match) return '';
  
  const bytes = Utilities.base64Decode(match[2]);
  if (bytes.length > 5 * 1024 * 1024) {
    throw new Error('画像サイズは1枚あたり5MB以下にしてください。');
  }
  
  const blob = Utilities.newBlob(bytes, match[1], name);
  const file = folder.createFile(blob);
  return file.getUrl();
}
