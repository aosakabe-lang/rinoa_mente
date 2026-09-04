/**
 * 株式会社リノア（REN）太陽光発電点検報告システム - GASバックエンド v2.0
 * 案1（A〜N全14列）完全対応・自動シートフォーマット・安全排他制御版
 */

const SHEET_NAME = '点検データ';

// 案1：拡張14列定義
const HEADERS = [
  'ID',          // A列: ユニークキー (REN-UUID)
  '点検日時',    // B列: タイムスタンプ
  '担当者',      // C列: エンジニア名
  '施主名',      // D列: 顧客名
  '現場住所',    // E列: 発電所所在地 (案1追加)
  'メール',      // F列: 施主様Email
  '判定',        // G列: 〇 良好 / △ 経過観察 / × 要修繕
  'コメント',    // H列: 総合所見
  '全体写真',    // I列: アレイ全体画像Drive URL
  'パワコン写真',// J列: PCS外観画像Drive URL (案1追加)
  '異常写真',    // K列: 特記事項画像Drive URL
  '詳細数値',    // L列: 電気特性データ
  'ステータス',  // M列: 未処理 / 送信完了 / 送信エラー / 取り消し
  'PDF_URL'      // N列: 生成PDFのDrive URL (案1追加)
];

const ALLOWED_STATUSES = ['〇 良好', '△ 経過観察', '× 要修繕'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('REN | 太陽光発電設備 点検報告システム')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * ドライブ出力フォルダの初期化 setup
 */
function setupOutputFolder() {
  const folder = DriveApp.createFolder('REN_太陽光点検報告書');
  PropertiesService.getScriptProperties().setProperty('OUTPUT_FOLDER_ID', folder.getId());
  return folder.getUrl();
}

/**
 * 【重要】スプレッドシート・テンプレートを完全上書き初期化（リセット）する管理用関数
 */
function setupSystem() {
  const sheet = forceResetSpreadsheet_();
  const templateId = getOrCreateTemplateId_(true); // テンプレート強制更新
  return {
    message: '全14列フォーマットでスプレッドシートおよびPDFテンプレートを自動上書き作成しました。',
    sheetName: sheet.getName(),
    templateId: templateId
  };
}

/**
 * シートを初期化・全14列ヘッダーでフォーマット装飾上書きする内部処理
 */
function forceResetSpreadsheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  
  // 1行目に全14列のヘッダーを強制設定
  const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  headerRange.setValues([HEADERS]);
  
  // デザイン装飾（ヘッダー背景・文字色・太字・フリーズ）
  headerRange
    .setBackground('#0f172a') // ダークネイビー
    .setFontColor('#38bdf8') // シアンテキスト
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  
  sheet.setRowHeight(1, 35);
  sheet.setFrozenRows(1);
  
  // 追加項目強調表示 (E列, J列, N列のヘッダー背景を変更)
  sheet.getRange(1, 5).setBackground('#1e293b').setFontColor('#fbbf24'); // E列: 現場住所
  sheet.getRange(1, 10).setBackground('#1e293b').setFontColor('#fbbf24'); // J列: パワコン写真
  sheet.getRange(1, 14).setBackground('#1e293b').setFontColor('#fbbf24'); // N列: PDF_URL

  return sheet;
}

/**
 * 点検一覧データ取得（14列対応）
 */
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
      address: String(row[4] || ''),       // E列: 現場住所
      email: String(row[5] || ''),
      status: String(row[6] || ''),
      comment: String(row[7] || ''),
      photoMain: String(row[8] || ''),
      photoPcs: String(row[9] || ''),      // J列: パワコン写真
      photoSub: String(row[10] || ''),
      details: String(row[11] || ''),
      processStatus: String(row[12] || '未処理'),
      pdfUrl: String(row[13] || '')         // N列: PDF_URL
    };
  }).filter(function (item) { 
    return item.processStatus !== '取り消し'; 
  }).reverse();
}

/**
 * 新規点検データの保存（LockServiceによる排他制御 & 14列追記）
 */
function saveInspectionData(form) {
  const data = validateForm_(form);
  const id = 'REN-' + Utilities.getUuid().substring(0, 8).toUpperCase() + '-2026';
  const folder = getOutputFolder_();

  // 画像ファイル保存処理
  const photoMainUrl = saveImage_(data.photoMainData, id + '_01全体.jpg', folder);
  const photoPcsUrl  = saveImage_(data.photoPcsData,  id + '_02パワコン.jpg', folder);
  const photoSubUrl  = saveImage_(data.photoSubData,  id + '_03異常.jpg', folder);

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000); // 同時書き込み競合待ち（最大30秒）

  try {
    const sheet = getInspectionSheet_();
    sheet.appendRow([
      id,                  // A: ID
      new Date(),          // B: 点検日時
      data.worker,         // C: 担当者
      data.client,         // D: 施主名
      data.address,        // E: 現場住所 (案1)
      data.email,          // F: メール
      data.status,         // G: 判定
      data.comment,        // H: コメント
      photoMainUrl,        // I: 全体写真
      photoPcsUrl,         // J: パワコン写真 (案1)
      photoSubUrl,         // K: 異常写真
      data.details,        // L: 詳細数値
      '未処理',            // M: ステータス
      ''                   // N: PDF_URL (PDF発行時更新)
    ]);
  } finally {
    lock.releaseLock();
  }

  return { success: true, message: '点検データ（14項目）をスプレッドシートへ正常保存しました。' };
}

/**
 * PDFレポート作成 ＆ スプレッドシート（N列）へのPDF URL記録
 */
function createPdfReport_(row, retainPdf) {
  const client = safeFilePart_(row[3]);
  const date = formatDate_(row[1], 'yyyyMMdd');
  const folder = getOutputFolder_();
  
  // テンプレート複製
  const templateId = getOrCreateTemplateId_(false);
  const copy = DriveApp.getFileById(templateId).makeCopy('点検報告書_' + client + '様_' + date, folder);
  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();

  // プレースホルダー置換（14列対応）
  replaceText_(body, '{{施主名}}', row[3]);
  replaceText_(body, '{{現場住所}}', row[4] || '未登録');
  replaceText_(body, '{{点検日時}}', formatDate_(row[1], 'yyyy/MM/dd HH:mm'));
  replaceText_(body, '{{担当者}}', row[2]);
  replaceText_(body, '{{総合判定}}', row[6]);
  replaceText_(body, '{{見解コメント}}', row[7] || '特記事項なし');
  replaceText_(body, '{{詳細数値}}', row[11] || '特記事項なし');

  // 画像埋め込み（全体・パワコン・異常）
  replacePhoto_(body, '{{現場全体写真}}', row[8]);
  replacePhoto_(body, '{{パワコン写真}}', row[9]);
  replacePhoto_(body, '{{異常箇所写真}}', row[10]);

  doc.saveAndClose();

  // PDF出力
  const pdfFile = folder.createFile(copy.getAs(MimeType.PDF)).setName('太陽光点検報告書_' + client + '様_' + date + '.pdf');
  copy.setTrashed(true); // 作業用Doc削除

  return pdfFile;
}

/**
 * PDF プレビュー取得処理
 */
function previewPdfReport(rowIndex) {
  const row = getRow_(rowIndex);
  const pdf = createPdfReport_(row, false);
  try {
    return { 
      success: true, 
      name: pdf.getName(), 
      base64: Utilities.base64Encode(pdf.getBlob().getBytes()) 
    };
  } finally {
    pdf.setTrashed(true); // プレビュー一時ファイル削除
  }
}

/**
 * 施主様宛メール自動送信処理
 */
function sendReportEmail(rowIndex) {
  const sheet = getInspectionSheet_();
  const row = getRow_(rowIndex);
  
  if (row[12] === '送信完了') {
    throw new Error('この報告書はすでに送信済みです。再送が必要な場合は管理者がステータスを変更してください。');
  }
  
  const email = normalizeEmail_(row[5]);
  if (!email) throw new Error('施主様のメールアドレス形式が不正です。');

  const pdf = createPdfReport_(row, true);
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

    // ステータス（M列: 13列目）と PDF_URL（N列: 14列目）を上書き更新
    sheet.getRange(rowIndex, 13).setValue('送信完了');
    sheet.getRange(rowIndex, 14).setValue(pdfUrl);

    return { success: true, message: client + '様へ点検報告書メールを正常送信しました。' };
  } catch (error) {
    sheet.getRange(rowIndex, 13).setValue('送信エラー');
    throw new Error('メール送信に失敗しました: ' + error.message);
  }
}

/**
 * Master PDF Template Creator (Supports all tags)
 */
function getOrCreateTemplateId_(forceCreate) {
  const properties = PropertiesService.getScriptProperties();
  const existing = properties.getProperty('TEMPLATE_DOC_ID');
  
  if (existing && !forceCreate) {
    try {
      DriveApp.getFileById(existing);
      return existing;
    } catch (ignore) {}
  }

  const doc = DocumentApp.create('【マスター】太陽光発電設備 定期点検報告書_v2');
  const body = doc.getBody();

  body.appendParagraph('太陽光発電設備 定期点検報告書')
      .setHeading(DocumentApp.ParagraphHeading.HEADING1)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      
  body.appendParagraph('発行元: 株式会社リノア（REN） メンテナンス事業部\n');

  // メタ情報テーブル
  body.appendTable([
    ['お客様名', '{{施主名}} 様'],
    ['現場住所', '{{現場住所}}'],
    ['点検日時', '{{点検日時}}'],
    ['担当者名', '{{担当者}}'],
    ['総合判定', '{{総合判定}}']
  ]);

  body.appendParagraph('\n担当者見解・総合コメント').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('{{見解コメント}}');

  body.appendParagraph('\n測定データ・詳細数値').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('{{詳細数値}}');

  body.appendParagraph('\n現場点検写真').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('① 全体設置状態: {{現場全体写真}}');
  body.appendParagraph('② パワコン外観: {{パワコン写真}}');
  body.appendParagraph('③ 異常特記事項: {{異常箇所写真}}');

  doc.saveAndClose();

  const file = DriveApp.getFileById(doc.getId());
  file.moveTo(getOutputFolder_());
  properties.setProperty('TEMPLATE_DOC_ID', doc.getId());
  
  return doc.getId();
}

/**
 * ドキュメント内画像置換ヘルパー
 */
function replacePhoto_(body, placeholder, url) {
  const found = body.findText(escapeRegex_(placeholder));
  if (!found || !url) { 
    replaceText_(body, placeholder, '（写真添付なし）'); 
    return; 
  }
  try {
    const driveId = extractDriveId_(url);
    const blob = DriveApp.getFileById(driveId).getBlob();
    const text = found.getElement().asText();
    text.deleteText(found.getStartOffset(), found.getEndOffsetInclusive());
    const image = text.getParent().asParagraph().appendInlineImage(blob);
    const width = image.getWidth();
    if (width > 0) {
      image.setWidth(300).setHeight(Math.round(image.getHeight() * 300 / width));
    }
  } catch (error) {
    replaceText_(body, placeholder, '（画像読み込みエラー）');
  }
}

function replaceText_(body, placeholder, value) {
  body.replaceText(escapeRegex_(placeholder), escapeReplacement_(String(value)));
}

function escapeRegex_(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function escapeReplacement_(value) {
  return value.replace(/\\/g, '\\\\').replace(/\$/g, '\\$');
}

function extractDriveId_(url) {
  const match = String(url).match(/[-\w]{25,}/);
  if (!match) throw new Error('画像Drive URLが無効です。');
  return match[0];
}

function safeFilePart_(value) {
  return String(value || '施主未設定').replace(/[\\/:*?"<>|]/g, '_').slice(0, 80);
}

function formatDate_(value, pattern) {
  return value instanceof Date ? Utilities.formatDate(value, Session.getScriptTimeZone(), pattern) : String(value || '');
}

function getInspectionSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = forceResetSpreadsheet_();
  }
  return sheet;
}

function getOutputFolder_() {
  const id = PropertiesService.getScriptProperties().getProperty('OUTPUT_FOLDER_ID');
  if (!id) throw new Error('OUTPUT_FOLDER_ID が未設定です。GASエディタで setupOutputFolder を一度実行してください。');
  return DriveApp.getFolderById(id);
}

function getRow_(rowIndex) {
  if (!Number.isInteger(rowIndex) || rowIndex < 2) throw new Error('不正な行番号です。');
  const sheet = getInspectionSheet_();
  const row = sheet.getRange(rowIndex, 1, 1, HEADERS.length).getValues()[0];
  if (!row[0]) throw new Error('対象の点検データが見つかりません。');
  return row;
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

  return {
    worker: String(form.worker || '').trim().slice(0, 100),
    client: String(form.client || '').trim().slice(0, 100),
    address: String(form.address || '').trim().slice(0, 200), // E列: 現場住所
    email: email,
    status: status,
    comment: String(form.comment || '').slice(0, 3000),
    details: String(form.details || '').slice(0, 3000),
    photoMainData: form.photoMainData || '',
    photoPcsData: form.photoPcsData || '',                  // J列: パワコン写真
    photoSubData: form.photoSubData || ''
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

// 連携テスト用コメント