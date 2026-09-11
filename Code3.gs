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

function saveInspectionData(form) {
  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(30000);
    const id = 'REN-' + Utilities.getUuid().substring(0, 8).toUpperCase() + '-2026';
    const folder = getOutputFolder_();

    // 写真のドライブ保存（新項目含む）
    const photoMainUrl = saveImage_(form.photoMainData, id + '_01全体.jpg', folder);
    const photoPcsUrl = saveImage_(form.photoPcsData, id + '_02パワコン.jpg', folder);
    const photoSubUrl = saveImage_(form.photoSubData, id + '_03異常.jpg', folder);
    const photoPanelUrl = saveImage_(form.photoPanelData, id + '_04パネル.jpg', folder);
    const photoClampUrl = saveImage_(form.photoClampData, id + '_05金具.jpg', folder);
    const photoLegUrl = saveImage_(form.photoLegData, id + '_06脚部.jpg', folder);
    const photoBackUrl = saveImage_(form.photoBackData, id + '_07裏面.jpg', folder);

    const sheet = getInspectionSheet_();
    sheet.appendRow([
      id, new Date(), form.worker, form.client, form.address, form.email || '', form.status,
      form.comment || '', photoMainUrl, photoPcsUrl, photoSubUrl, photoPanelUrl, photoClampUrl, photoLegUrl, photoBackUrl,
      '下書き', '', form.voltageType, form.plantName || '', form.chiefEngineer || '', form.powerCompany || '',
      form.outdoorTemp || '', form.indoorTemp || '', form.weather || ''
    ]);

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
