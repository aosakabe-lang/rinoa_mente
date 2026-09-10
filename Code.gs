/**
 * 株式会社リノア（REN）太陽光発電点検報告システム - GASバックエンド v3.2
 * Webアプリ完結型（現場入力 → 確認・修正 → 報告書生成）フル対応版
 */

const SHEET_NAME = '点検データ';
const PCS_SHEET_NAME = 'PCS点検データ';

const HEADERS = [
  'ID', '点検日時', '担当者', '施主名', '現場住所', 'メール', '判定', 'コメント',
  '全体写真', 'パワコン写真', '異常写真', '詳細数値', 'ステータス', 'PDF_URL',
  '発電所名', '電気主任技術者', '供給先(電力会社)', '供給先(支社)', '最大出力(kW)', '定格出力(kW)',
  '点検開始時刻', '点検終了時刻', '天気', '室内温度(℃)', '室内湿度(％)', '外気温度(℃)', '外気湿度(％)', '平均日射量(KW/m2)'
];

const PCS_HEADERS = [
  '点検ID', 'PCS番号', '設備ID', '型式', '製造番号', '瞬間発電量(kW)', '積算発電力量(kWh)',
  '抑制時間(分)', '直列数', '並列数', '開放電圧(V)', '動作電圧(V)', '動作電流(A)',
  '絶縁抵抗値(MΩ)', '交流電圧(V)', '接地抵抗値(Ω)', '自立運転電圧(V)'
];

const ALLOWED_STATUSES = ['〇 良好', '△ 経過観察', '× 要修繕'];
const PROCESS_STATUSES = ['下書き', '確定', 'PDF保存済み', 'メール確認', '送信完了', '送信エラー', '取り消し'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('REN | 太陽光発電設備 点検報告システム')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function setupOutputFolder() {
  const folder = DriveApp.createFolder('REN_太陽光点検報告書');
  PropertiesService.getScriptProperties().setProperty('OUTPUT_FOLDER_ID', folder.getId());
  return folder.getUrl();
}

function setupSystem() {
  const sheet = forceResetSpreadsheet_();
  const pcsSheet = forceResetPcsSheet_();
  const templateId = getOrCreateTemplateId_(true);
  return {
    message: '全28列シート・PCS点検データシートおよびPDFテンプレートを自動初期化しました。',
    sheetName: sheet.getName(),
    pcsSheetName: pcsSheet.getName(),
    templateId: templateId
  };
}

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

function getInspectionData() {
  const sheet = getInspectionSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];

  return values.slice(1).map(function (row, offset) {
    return {
      rowIndex: offset + 2,
      id: String(row[0] || ''),
      date: formatDate_(row[1], 'yyyy/MM/dd HH:mm'),
      worker: String(row[2] || ''),
      client: String(row[3] || ''),
      address: String(row[4] || ''),
      email: String(row[5] || ''),
      status: String(row[6] || ''),
      comment: String(row[7] || ''),
      photoMain: String(row[8] || ''),
      photoPcs: String(row[9] || ''),
      photoSub: String(row[10] || ''),
      details: String(row[11] || ''),
      processStatus: String(row[12] || '下書き'),
      pdfUrl: String(row[13] || ''),
      plantName: String(row[14] || ''),
      chiefEngineer: String(row[15] || ''),
      powerCompany: String(row[16] || ''),
      powerBranch: String(row[17] || ''),
      maxOutput: String(row[18] || ''),
      ratedOutput: String(row[19] || ''),
      startTime: String(row[20] || ''),
      endTime: String(row[21] || ''),
      weather: String(row[22] || ''),
      indoorTemp: String(row[23] || ''),
      indoorHumidity: String(row[24] || ''),
      outdoorTemp: String(row[25] || ''),
      outdoorHumidity: String(row[26] || ''),
      solarRadiation: String(row[27] || '')
    };
  }).filter(function (item) { return item.processStatus !== '取り消し'; }).reverse();
}

function getPcsDataById(inspectionId) {
  const id = String(inspectionId || '').trim();
  if (!id) return [];
  const sheet = getPcsSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];

  return values.slice(1)
    .filter(function (row) { return String(row[0]) === id; })
    .map(function (row) {
      return {
        inspectionId: String(row[0] || ''),
        pcsNo: String(row[1] || ''),
        equipmentId: String(row[2] || ''),
        model: String(row[3] || ''),
        serialNo: String(row[4] || ''),
        instantPower: String(row[5] || ''),
        cumulativePower: String(row[6] || ''),
        suppressionMinutes: String(row[7] || ''),
        seriesCount: String(row[8] || ''),
        parallelCount: String(row[9] || ''),
        openVoltage: String(row[10] || ''),
        operatingVoltage: String(row[11] || ''),
        operatingCurrent: String(row[12] || ''),
        insulationResistance: String(row[13] || ''),
        acVoltage: String(row[14] || ''),
        groundResistance: String(row[15] || ''),
        selfRunVoltage: String(row[16] || '')
      };
    });
}

/**
 * 新規点検データの保存（下書きとして追記）
 */
function saveInspectionData(form) {
  const lock = LockService.getDocumentLock();
  const savedImageUrls = [];
  try {
    const data = validateForm_(form);
    const id = 'REN-' + Utilities.getUuid().substring(0, 8).toUpperCase() + '-2026';
    const folder = getOutputFolder_();

    lock.waitLock(30000);

    const photoMainUrl = saveImage_(data.photoMainData, id + '_01全体.jpg', folder);
    const photoPcsUrl = saveImage_(data.photoPcsData, id + '_02パワコン.jpg', folder);
    const photoSubUrl = saveImage_(data.photoSubData, id + '_03異常.jpg', folder);
    [photoMainUrl, photoPcsUrl, photoSubUrl].forEach(function (url) { if (url) savedImageUrls.push(url); });

    const sheet = getInspectionSheet_();
    sheet.appendRow([
      id, new Date(), data.worker, data.client, data.address, data.email, data.status,
      data.comment, photoMainUrl, photoPcsUrl, photoSubUrl, data.details, '下書き', '',
      data.plantName, data.chiefEngineer, data.powerCompany, data.powerBranch,
      data.maxOutput, data.ratedOutput, data.startTime, data.endTime, data.weather,
      data.indoorTemp, data.indoorHumidity, data.outdoorTemp, data.outdoorHumidity, data.solarRadiation
    ]);

    if (data.pcsList && data.pcsList.length > 0) {
      const pcsSheet = getPcsSheet_();
      const pcsRows = data.pcsList.map(function (pcs, index) {
        return [
          id, pcs.pcsNo || ('PCS' + (index + 1)), pcs.equipmentId || '', pcs.model || '', pcs.serialNo || '',
          pcs.instantPower || '', pcs.cumulativePower || '', pcs.suppressionMinutes || '', pcs.seriesCount || '',
          pcs.parallelCount || '', pcs.openVoltage || '', pcs.operatingVoltage || '', pcs.operatingCurrent || '',
          pcs.insulationResistance || '', pcs.acVoltage || '', pcs.groundResistance || '', pcs.selfRunVoltage || ''
        ];
      });
      pcsSheet.getRange(pcsSheet.getLastRow() + 1, 1, pcsRows.length, PCS_HEADERS.length).setValues(pcsRows);
    }

    SpreadsheetApp.flush();
    return { success: true, message: '点検データを送信・保存しました（下書き保存完了）。', id: id };
  } catch (error) {
    Logger.log(JSON.stringify({ function: 'saveInspectionData', message: error.message, stack: error.stack }));
    savedImageUrls.forEach(function (url) {
      try { DriveApp.getFileById(extractDriveId_(url)).setTrashed(true); } catch (e) {}
    });
    throw new Error('点検データの保存に失敗しました: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * 事務・確認担当者による既存点検データの修正・上書き更新
 */
function updateInspectionData(form) {
  const lock = LockService.getDocumentLock();
  try {
    if (!form.id) throw new Error('更新対象のIDが指定されていません。');
    const data = validateForm_(form);
    const targetRowIndex = resolveRowIndex_(form.id);
    const sheet = getInspectionSheet_();
    const currentRow = getRow_(targetRowIndex);

    lock.waitLock(30000);
    const folder = getOutputFolder_();

    // 写真の差替えチェック（Base64形式で新画像が渡された場合のみ更新・差し替え）
    let photoMainUrl = currentRow[8];
    let photoPcsUrl = currentRow[9];
    let photoSubUrl = currentRow[10];

    if (data.photoMainData && data.photoMainData.startsWith('data:image')) {
      photoMainUrl = saveImage_(data.photoMainData, form.id + '_01全体_rev.jpg', folder);
    }
    if (data.photoPcsData && data.photoPcsData.startsWith('data:image')) {
      photoPcsUrl = saveImage_(data.photoPcsData, form.id + '_02パワコン_rev.jpg', folder);
    }
    if (data.photoSubData && data.photoSubData.startsWith('data:image')) {
      photoSubUrl = saveImage_(data.photoSubData, form.id + '_03異常_rev.jpg', folder);
    }

    // 該当行データの更新設定
    const updatedValues = [
      currentRow[0],  // A: ID
      currentRow[1],  // B: 点検日時
      data.worker,    // C: 担当者
      data.client,    // D: 施主名
      data.address,   // E: 現場住所
      data.email,     // F: メール
      data.status,    // G: 判定
      data.comment,   // H: コメント
      photoMainUrl,   // I: 全体写真
      photoPcsUrl,    // J: パワコン写真
      photoSubUrl,    // K: 異常写真
      data.details,   // L: 詳細数値
      currentRow[12], // M: ステータス（維持）
      currentRow[13], // N: PDF_URL
      data.plantName, data.chiefEngineer, data.powerCompany, data.powerBranch,
      data.maxOutput, data.ratedOutput, data.startTime, data.endTime, data.weather,
      data.indoorTemp, data.indoorHumidity, data.outdoorTemp, data.outdoorHumidity, data.solarRadiation
    ];

    sheet.getRange(targetRowIndex, 1, 1, updatedValues.length).setValues([updatedValues]);

    // PCS明細の更新（既存のPCSデータを削除してから再挿入）
    if (data.pcsList) {
      const pcsSheet = getPcsSheet_();
      const pcsValues = pcsSheet.getDataRange().getValues();
      for (let i = pcsValues.length - 1; i >= 1; i--) {
        if (String(pcsValues[i][0]) === form.id) {
          pcsSheet.deleteRow(i + 1);
        }
      }
      if (data.pcsList.length > 0) {
        const pcsRows = data.pcsList.map(function (pcs, index) {
          return [
            form.id, pcs.pcsNo || ('PCS' + (index + 1)), pcs.equipmentId || '', pcs.model || '', pcs.serialNo || '',
            pcs.instantPower || '', pcs.cumulativePower || '', pcs.suppressionMinutes || '', pcs.seriesCount || '',
            pcs.parallelCount || '', pcs.openVoltage || '', pcs.operatingVoltage || '', pcs.operatingCurrent || '',
            pcs.insulationResistance || '', pcs.acVoltage || '', pcs.groundResistance || '', pcs.selfRunVoltage || ''
          ];
        });
        pcsSheet.getRange(pcsSheet.getLastRow() + 1, 1, pcsRows.length, PCS_HEADERS.length).setValues(pcsRows);
      }
    }

    SpreadsheetApp.flush();
    return { success: true, message: '点検修正データを正常に上書き保存しました。' };
  } catch (error) {
    Logger.log(JSON.stringify({ function: 'updateInspectionData', message: error.message, stack: error.stack }));
    throw new Error('データの更新・修正に失敗しました: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

function createPdfReport_(row, retainPdf) {
  let copy = null;
  try {
    if (!row || row.length < HEADERS.length) throw new Error('データ行の項目数が不足しています。');

    const client = safeFilePart_(row[3]);
    const date = formatDate_(row[1], 'yyyyMMdd');
    const folder = getOutputFolder_();
    const templateId = getOrCreateTemplateId_(false);

    copy = DriveApp.getFileById(templateId).makeCopy('点検報告書_' + client + '様_' + date, folder);
    const doc = DocumentApp.openById(copy.getId());
    const body = doc.getBody();

    replaceText_(body, '{{施主名}}', row[3]);
    replaceText_(body, '{{現場住所}}', row[4] || '未登録');
    replaceText_(body, '{{点検日時}}', formatDate_(row[1], 'yyyy/MM/dd HH:mm'));
    replaceText_(body, '{{担当者}}', row[2]);
    replaceText_(body, '{{総合判定}}', row[6]);
    replaceText_(body, '{{見解コメント}}', row[7] || '特記事項なし');
    replaceText_(body, '{{詳細数値}}', row[11] || '特記事項なし');
    replaceText_(body, '{{発電所名}}', row[14] || '未登録');
    replaceText_(body, '{{電気主任技術者}}', row[15] || '未登録');
    replaceText_(body, '{{供給先}}', (row[16] || '') + ' ' + (row[17] || ''));
    replaceText_(body, '{{最大定格出力}}', (row[18] || '未登録') + 'kW / ' + (row[19] || '未登録') + 'kW');
    replaceText_(body, '{{点検時間}}', (row[20] || '') + '〜' + (row[21] || ''));
    replaceText_(body, '{{気象条件}}', '天気:' + (row[22] || '-') + ' 室内:' + (row[23] || '-') + '℃/' + (row[24] || '-') + '% 外気:' + (row[25] || '-') + '℃/' + (row[26] || '-') + '% 日射量:' + (row[27] || '-'));

    replacePhoto_(body, '{{現場全体写真}}', row[8]);
    replacePhoto_(body, '{{パワコン写真}}', row[9]);
    replacePhoto_(body, '{{異常箇所写真}}', row[10]);

    const pcsList = getPcsDataById(row[0]);
    if (pcsList.length > 0) {
      body.appendParagraph('\nPCS(パワコン)別 点検測定データ').setHeading(DocumentApp.ParagraphHeading.HEADING2);
      const tableRows = [['PCS番号', '設備ID', '型式', '製造番号', '瞬間発電量', '積算発電力量', '絶縁抵抗値']];
      pcsList.forEach(function (pcs) {
        tableRows.push([pcs.pcsNo, pcs.equipmentId, pcs.model, pcs.serialNo, pcs.instantPower, pcs.cumulativePower, pcs.insulationResistance]);
      });
      body.appendTable(tableRows);
    }

    doc.saveAndClose();
    SpreadsheetApp.flush();

    const pdfFile = folder.createFile(copy.getAs(MimeType.PDF)).setName('太陽光点検報告書_' + client + '様_' + date + '.pdf');
    copy.setTrashed(true);
    return pdfFile;
  } catch (error) {
    console.error(JSON.stringify({ function: 'createPdfReport_', recordId: row && row[0], message: error.message }));
    if (copy) copy.setTrashed(true);
    throw new Error('PDF生成に失敗しました: ' + error.message);
  }
}

function previewPdfReport(rowIndex) {
  const row = getRow_(rowIndex);
  if (!['確定', 'PDF保存済み', 'メール確認', '送信エラー'].includes(row[12])) {
    throw new Error('内容を確定してからPDFプレビューを実行してください。');
  }
  const pdf = createPdfReport_(row, false);
  try {
    return {
      success: true,
      name: pdf.getName(),
      base64: Utilities.base64Encode(pdf.getBlob().getBytes())
    };
  } finally {
    pdf.setTrashed(true);
  }
}

function confirmInspection(rowIndex) {
  return updateProcessStatus_(rowIndex, '確定', ['下書き']);
}

function savePdfReport(rowIndex) {
  const targetRowIndex = resolveRowIndex_(rowIndex);
  const sheet = getInspectionSheet_();
  const row = getRow_(targetRowIndex);
  if (!['確定', 'PDF保存済み', 'メール確認', '送信エラー'].includes(row[12])) {
    throw new Error('内容を確定してからPDFを保存してください。');
  }

  const pdf = createPdfReport_(row, true);
  sheet.getRange(targetRowIndex, 13).setValue('PDF保存済み');
  sheet.getRange(targetRowIndex, 14).setValue(pdf.getUrl());
  SpreadsheetApp.flush();
  return { success: true, url: pdf.getUrl(), name: pdf.getName() };
}

function confirmEmail(rowIndex) {
  return updateProcessStatus_(rowIndex, 'メール確認', ['PDF保存済み']);
}

function updateProcessStatus_(rowIndex, nextStatus, allowedCurrentStatuses) {
  if (!PROCESS_STATUSES.includes(nextStatus)) throw new Error('不正な処理ステータスです。');
  const targetRowIndex = resolveRowIndex_(rowIndex);
  const sheet = getInspectionSheet_();
  const row = getRow_(targetRowIndex);
  if (!allowedCurrentStatuses.includes(row[12])) {
    throw new Error('現在の状態（' + row[12] + '）から「' + nextStatus + '」へ変更できません。');
  }
  sheet.getRange(targetRowIndex, 13).setValue(nextStatus);
  SpreadsheetApp.flush();
  return { success: true, status: nextStatus };
}

function sendReportEmail(rowIndex) {
  const sheet = getInspectionSheet_();
  const targetRowIndex = resolveRowIndex_(rowIndex);
  const row = getRow_(targetRowIndex);
  
  if (row[12] === '送信完了') {
    throw new Error('この報告書はすでに送信済みです。');
  }
  
  const email = normalizeEmail_(row[5]);
  if (!email) throw new Error('施主様のメールアドレス形式が不正です。');

  if (row[12] !== 'メール確認') {
    throw new Error('メール内容を確認してから送信してください。');
  }
  const pdfId = extractDriveId_(row[13]);
  const pdf = DriveApp.getFileById(pdfId);
  const pdfUrl = pdf.getUrl();

  try {
    const client = String(row[3]);
    GmailApp.sendEmail(
      email,
      '【株式会社リノア】太陽光発電設備 定期点検報告書（' + client + '様）',
      client + ' 様\n\n株式会社リノア（REN）点検担当の ' + String(row[2]) + ' です。\n\n本日の太陽光発電設備 定期点検報告書を添付いたします。\n\n【点検総合判定】 ' + String(row[6]) + '\n【現場住所】 ' + String(row[4] || '記載なし') + '\n\nご確認のほどよろしくお願い申し上げます。',
      {
        attachments: [pdf.getBlob()],
        name: '株式会社リノア メンテナンス事業部'
      }
    );

    sheet.getRange(targetRowIndex, 13).setValue('送信完了');
    sheet.getRange(targetRowIndex, 14).setValue(pdfUrl);
    SpreadsheetApp.flush();

    return { success: true, message: client + '様へ点検報告書メールを正常送信しました。' };
  } catch (error) {
    sheet.getRange(targetRowIndex, 13).setValue('送信エラー');
    throw new Error('メール送信に失敗しました: ' + error.message);
  }
}

// 共通ヘルパー関数群
function getInspectionSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME) || forceResetSpreadsheet_();
}

function getPcsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(PCS_SHEET_NAME) || forceResetPcsSheet_();
}

function getOutputFolder_() {
  const id = PropertiesService.getScriptProperties().getProperty('OUTPUT_FOLDER_ID');
  if (!id) throw new Error('OUTPUT_FOLDER_ID 未設定。setupOutputFolder を実行してください。');
  return DriveApp.getFolderById(id);
}

function getOrCreateTemplateId_(forceCreate) {
  const properties = PropertiesService.getScriptProperties();
  const existing = properties.getProperty('TEMPLATE_DOC_ID');
  if (existing && !forceCreate) {
    try { DriveApp.getFileById(existing); return existing; } catch (e) {}
  }
  const doc = DocumentApp.create('【マスター】太陽光発電設備 定期点検報告書_v3');
  const body = doc.getBody();
  body.appendParagraph('太陽光発電設備 定期点検報告書').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendTable([
    ['お客様名', '{{施主名}} 様'], ['現場住所', '{{現場住所}}'], ['発電所名', '{{発電所名}}'],
    ['点検日時', '{{点検日時}}'], ['担当者名', '{{担当者}}'], ['総合判定', '{{総合判定}}'],
    ['気象条件', '{{気象条件}}']
  ]);
  body.appendParagraph('\n担当者見解').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('{{見解コメント}}');
  body.appendParagraph('\n現場点検写真').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('① 全体: {{現場全体写真}}');
  body.appendParagraph('② PCS: {{パワコン写真}}');
  body.appendParagraph('③ 異常: {{異常箇所写真}}');
  doc.saveAndClose();
  const file = DriveApp.getFileById(doc.getId());
  file.moveTo(getOutputFolder_());
  properties.setProperty('TEMPLATE_DOC_ID', doc.getId());
  return doc.getId();
}

function replacePhoto_(body, placeholder, url) {
  const found = body.findText(escapeRegex_(placeholder));
  if (!found || !url) { replaceText_(body, placeholder, '（写真なし）'); return; }
  try {
    const driveId = extractDriveId_(url);
    const blob = DriveApp.getFileById(driveId).getBlob();
    const text = found.getElement().asText();
    text.deleteText(found.getStartOffset(), found.getEndOffsetInclusive());
    const image = text.getParent().asParagraph().appendInlineImage(blob);
    if (image.getWidth() > 0) image.setWidth(300).setHeight(Math.round(image.getHeight() * 300 / image.getWidth()));
  } catch (e) {
    replaceText_(body, placeholder, '（画像エラー）');
  }
}

function replaceText_(body, placeholder, value) {
  body.replaceText(escapeRegex_(placeholder), escapeReplacement_(String(value)));
}

function escapeRegex_(value) { return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
function escapeReplacement_(value) { return value.replace(/\\/g, '\\\\').replace(/\$/g, '\\$'); }
function extractDriveId_(url) {
  const match = String(url).match(/[-\w]{25,}/);
  if (!match) throw new Error('Drive URLが無効です。');
  return match[0];
}
function safeFilePart_(value) { return String(value || '未設定').replace(/[\\/:*?"<>|]/g, '_').slice(0, 80); }
function formatDate_(value, pattern) {
  return value instanceof Date ? Utilities.formatDate(value, Session.getScriptTimeZone(), pattern) : String(value || '');
}
function getRow_(rowIndex) {
  const sheet = getInspectionSheet_();
  if (Number.isInteger(rowIndex) && rowIndex >= 2) return sheet.getRange(rowIndex, 1, 1, HEADERS.length).getValues()[0];
  const recordId = String(rowIndex || '').trim();
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) { if (String(values[i][0]) === recordId) return values[i]; }
  throw new Error('対象データが見つかりません: ' + recordId);
}
function resolveRowIndex_(rowIndex) {
  if (Number.isInteger(rowIndex) && rowIndex >= 2) return rowIndex;
  const recordId = String(rowIndex || '').trim();
  const sheet = getInspectionSheet_();
  const ids = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 1).getValues();
  for (let i = 0; i < ids.length; i++) { if (String(ids[i][0]) === recordId) return i + 2; }
  throw new Error('対象行が見つかりません: ' + recordId);
}
function normalizeEmail_(value) {
  const email = String(value || '').trim();
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) ? email : '';
}
function validateForm_(form) {
  if (!form) throw new Error('入力データが空です。');
  const status = String(form.status || '');
  if (!ALLOWED_STATUSES.includes(status)) throw new Error('総合判定の値が不正です。');
  const email = normalizeEmail_(form.email);
  if (!email) throw new Error('有効なメールアドレスを入力してください。');

  const pcsList = Array.isArray(form.pcsList) ? form.pcsList.slice(0, 30).map(function (pcs) {
    pcs = pcs || {};
    return {
      pcsNo: String(pcs.pcsNo || '').slice(0, 30),
      equipmentId: String(pcs.equipmentId || '').slice(0, 60),
      model: String(pcs.model || '').slice(0, 60),
      serialNo: String(pcs.serialNo || '').slice(0, 60),
      instantPower: String(pcs.instantPower || '').slice(0, 30),
      cumulativePower: String(pcs.cumulativePower || '').slice(0, 30),
      suppressionMinutes: String(pcs.suppressionMinutes || '').slice(0, 30),
      seriesCount: String(pcs.seriesCount || '').slice(0, 100),
      parallelCount: String(pcs.parallelCount || '').slice(0, 100),
      openVoltage: String(pcs.openVoltage || '').slice(0, 100),
      operatingVoltage: String(pcs.operatingVoltage || '').slice(0, 100),
      operatingCurrent: String(pcs.operatingCurrent || '').slice(0, 100),
      insulationResistance: String(pcs.insulationResistance || '').slice(0, 100),
      acVoltage: String(pcs.acVoltage || '').slice(0, 100),
      groundResistance: String(pcs.groundResistance || '').slice(0, 30),
      selfRunVoltage: String(pcs.selfRunVoltage || '').slice(0, 100)
    };
  }) : [];

  return {
    worker: String(form.worker || '').trim().slice(0, 100),
    client: String(form.client || '').trim().slice(0, 100),
    address: String(form.address || '').trim().slice(0, 200),
    email: email,
    status: status,
    comment: String(form.comment || '').slice(0, 3000),
    details: String(form.details || '').slice(0, 3000),
    photoMainData: form.photoMainData || '',
    photoPcsData: form.photoPcsData || '',
    photoSubData: form.photoSubData || '',
    plantName: String(form.plantName || '').trim().slice(0, 100),
    chiefEngineer: String(form.chiefEngineer || '').trim().slice(0, 100),
    powerCompany: String(form.powerCompany || '').trim().slice(0, 100),
    powerBranch: String(form.powerBranch || '').trim().slice(0, 100),
    maxOutput: String(form.maxOutput || '').trim().slice(0, 30),
    ratedOutput: String(form.ratedOutput || '').trim().slice(0, 30),
    startTime: String(form.startTime || '').trim().slice(0, 30),
    endTime: String(form.endTime || '').trim().slice(0, 30),
    weather: String(form.weather || '').trim().slice(0, 30),
    indoorTemp: String(form.indoorTemp || '').trim().slice(0, 30),
    indoorHumidity: String(form.indoorHumidity || '').trim().slice(0, 30),
    outdoorTemp: String(form.outdoorTemp || '').trim().slice(0, 30),
    outdoorHumidity: String(form.outdoorHumidity || '').trim().slice(0, 30),
    solarRadiation: String(form.solarRadiation || '').trim().slice(0, 30),
    pcsList: pcsList
  };
}

function saveImage_(dataUrl, name, folder) {
  if (!dataUrl) return '';
  const match = String(dataUrl).match(/^data:(image\/(?:jpeg|png|webp));base64,([A-Za-z0-9+/=]+)$/);
  if (!match) return '';
  const bytes = Utilities.base64Decode(match[2]);
  if (bytes.length > 5 * 1024 * 1024) throw new Error('画像サイズは1枚あたり5MB以下にしてください。');
  return folder.createFile(Utilities.newBlob(bytes, match[1], name)).getUrl();
}
