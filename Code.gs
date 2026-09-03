/**
 * PRESS MONITOR — クラウド連携バックエンド (Google Apps Script)
 *
 * 役割:
 *  1) iPadアプリから送られてくる日報1件ずつを、このスクリプトに紐づく
 *     Googleスプレッドシートへ自動保存する（管理者はいつでもシートを開いて閲覧可能）
 *  2) iPadアプリで「選択中の日報をメール送信」を押した際、選択された日報を
 *     Excel(.xlsx)添付ファイル付きメールとして指定の宛先へ自動送信する
 *
 * セットアップ手順は同梱の SETUP_GUIDE.md を参照してください。
 * このファイルは「コンテナバインドスクリプト」として、
 * Googleスプレッドシートの [拡張機能] > [Apps Script] から貼り付けて使う前提です。
 */

const SHEET_NAME = '日報データ';
const PART_MASTER_SHEET = '品番マスタ';
const OPERATOR_MASTER_SHEET = '作業員マスタ';

const HEADER_ROW = [
  '受信日時', 'レコードID', '日付', '設備名', '作業員', '作業人数', '部品番号(品番)',
  '製造開始時間', '製造終了時間', '設定SPM', '生産数(pcs)', '材料交換(回)', 'スクラップ交換(回)',
  '停止理由', '稼働時間(秒)', '金型交換時間(秒)', '計画停止時間(秒)', '異常停止時間(秒)', '性能稼働率(%)',
];

/* =========================================================
   エントリポイント
   ========================================================= */
function doPost(e) {
  let payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse({ status: 'error', message: '不正なリクエスト形式です' });
  }

  try {
    switch (payload.action) {
      case 'ping':
        return jsonResponse({ status: 'ok', message: 'pong' });
      case 'saveRecord':
        return handleSaveRecord(payload.record);
      case 'sendReport':
        return handleSendReport(payload.email, payload.equipment, payload.records);
      case 'getMasters':
        return handleGetMasters();
      case 'uploadPartMaster':
        return handleUploadPartMaster(payload.rows);
      case 'uploadOperatorMaster':
        return handleUploadOperatorMaster(payload.names);
      default:
        return jsonResponse({ status: 'error', message: '不明なアクションです: ' + payload.action });
    }
  } catch (err) {
    return jsonResponse({ status: 'error', message: String(err) });
  }
}

function doGet(e) {
  return jsonResponse({ status: 'ok', message: 'PRESS MONITOR backend is running.' });
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* =========================================================
   1) 日報1件をスプレッドシートへ保存
   ========================================================= */
function handleSaveRecord(record) {
  if (!record) return jsonResponse({ status: 'error', message: 'record がありません' });

  const sheet = getOrCreateSheet();
  const row = recordToRow(record);

  // 同一レコードID（iPad側で採番したid）が既にあれば上書き、なければ追記
  const idColIndex = 2; // B列 = レコードID
  const data = sheet.getDataRange().getValues();
  let targetRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idColIndex - 1]) === String(record.id)) { targetRow = i + 1; break; }
  }

  if (targetRow > 0) {
    sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }

  return jsonResponse({ status: 'ok', id: record.id });
}

function recordToRow(rec) {
  return [
    new Date(),
    rec.id,
    rec.date,
    rec.equipment || '',
    rec.operator || '未選択',
    rec.workerCount || 1,
    rec.partNo || '',
    rec.start || '',
    rec.end || '',
    rec.settingSPM || '',
    rec.count || 0,
    rec.materialCount || 0,
    rec.scrapCount || 0,
    rec.stopReason || 'なし',
    rec.runningSec || 0,
    rec.moldSec || 0,
    rec.breakSec || 0,
    rec.stopSec || 0,
    rec.perfRate || '',
  ];
}

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADER_ROW);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADER_ROW.length).setFontWeight('bold').setBackground('#1a2233').setFontColor('#ffffff');
  }
  return sheet;
}

/* =========================================================
   3) 品番マスタ・作業員マスタの取得
   管理者は「品番マスタ」「作業員マスタ」シートをExcelのように
   直接編集するだけで、iPad側に反映されます（列の並びは固定）。
   ========================================================= */
function handleGetMasters() {
  return jsonResponse({
    status: 'ok',
    partMasters: readPartMasterSheet(),
    operators: readOperatorMasterSheet(),
  });
}

function readPartMasterSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(PART_MASTER_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(PART_MASTER_SHEET);
    sheet.appendRow(['品番', 'SPM']);
    sheet.appendRow(['SP-200', 15]);
    sheet.appendRow(['SP-210', 18]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#1a2233').setFontColor('#ffffff');
  }
  const data = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    const part = String(data[i][0] || '').trim();
    const spm = parseFloat(data[i][1]);
    if (part && !isNaN(spm)) result.push({ part: part, spm: spm });
  }
  return result;
}

function readOperatorMasterSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(OPERATOR_MASTER_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(OPERATOR_MASTER_SHEET);
    sheet.appendRow(['作業員名']);
    sheet.appendRow(['山田']);
    sheet.appendRow(['佐藤']);
    sheet.appendRow(['鈴木']);
    sheet.appendRow(['田中']);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 1).setFontWeight('bold').setBackground('#1a2233').setFontColor('#ffffff');
  }
  const data = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][0] || '').trim();
    if (name) result.push(name);
  }
  return result;
}

/* =========================================================
   5) 管理者ページ(admin.html)からのマスタ上書き保存
   Excelファイルを解析した結果（品番+SPMの配列 / 作業員名の配列）を
   受け取り、品番マスタ・作業員マスタシートの中身を丸ごと置き換える。
   ========================================================= */
function handleUploadPartMaster(rows) {
  if (!Array.isArray(rows) || rows.length === 0) {
    return jsonResponse({ status: 'error', message: 'アップロードするデータがありません' });
  }
  const cleaned = rows
    .map(r => ({ part: String(r.part || '').trim(), spm: parseFloat(r.spm) }))
    .filter(r => r.part && !isNaN(r.spm));
  if (cleaned.length === 0) {
    return jsonResponse({ status: 'error', message: '有効な行が1件もありませんでした（品番またはSPMを確認してください）' });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(PART_MASTER_SHEET);
  if (sheet) { ss.deleteSheet(sheet); }
  sheet = ss.insertSheet(PART_MASTER_SHEET);
  sheet.appendRow(['品番', 'SPM']);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#1a2233').setFontColor('#ffffff');
  const values = cleaned.map(r => [r.part, r.spm]);
  sheet.getRange(2, 1, values.length, 2).setValues(values);

  return jsonResponse({ status: 'ok', count: cleaned.length });
}

function handleUploadOperatorMaster(names) {
  if (!Array.isArray(names) || names.length === 0) {
    return jsonResponse({ status: 'error', message: 'アップロードするデータがありません' });
  }
  const cleaned = [...new Set(names.map(n => String(n || '').trim()).filter(n => n))];
  if (cleaned.length === 0) {
    return jsonResponse({ status: 'error', message: '有効な作業員名が1件もありませんでした' });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(OPERATOR_MASTER_SHEET);
  if (sheet) { ss.deleteSheet(sheet); }
  sheet = ss.insertSheet(OPERATOR_MASTER_SHEET);
  sheet.appendRow(['作業員名']);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, 1).setFontWeight('bold').setBackground('#1a2233').setFontColor('#ffffff');
  const values = cleaned.map(n => [n]);
  sheet.getRange(2, 1, values.length, 1).setValues(values);

  return jsonResponse({ status: 'ok', count: cleaned.length });
}

/* =========================================================
   6) 選択した日報をExcel添付メールで自動送信
   ========================================================= */
function handleSendReport(email, equipment, records) {
  if (!email) return jsonResponse({ status: 'error', message: '送信先メールアドレスがありません' });
  if (!records || records.length === 0) return jsonResponse({ status: 'error', message: '送信対象データがありません' });

  const subject = `【${equipment || '設備'} 作業日報】${records[0].date || Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd')}`;
  const excelBlob = buildExcelBlob(records, equipment);

  const totalCount = records.reduce((s, r) => s + (r.count || 0), 0);
  const avgRate = (records.reduce((s, r) => s + parseFloat(r.perfRate || 0), 0) / records.length).toFixed(1);

  let body = `関係者各位\n\nお疲れ様です。${equipment || ''} の作業実績（全${records.length}件）を報告します。\n\n` +
    `----------------------------------------\n` +
    `総生産数: ${totalCount} pcs / 平均性能稼働率: ${avgRate}%\n` +
    `----------------------------------------\n`;

  records.forEach((rec, idx) => {
    body += `[${idx + 1}/${records.length}] 日付:${rec.date} | 部品:${rec.partNo} | 作業員:${rec.operator || '未選択'}\n` +
      `   時間:${rec.start}〜${rec.end} | 生産数:${rec.count}pcs | 材料交換:${rec.materialCount || 0}回 | ` +
      `スクラップ:${rec.scrapCount || 0}回 | 停止理由:${rec.stopReason || 'なし'} | 稼働率:${rec.perfRate}%\n`;
  });

  body += `----------------------------------------\n添付のExcelファイルに詳細データを記載しています。\n以上、よろしくお願いいたします。\n\n（本メールはPRESS MONITORアプリから自動送信されています）`;

  GmailApp.sendEmail(email, subject, body, { attachments: [excelBlob], name: 'PRESS MONITOR' });

  return jsonResponse({ status: 'ok', sent: records.length });
}

/**
 * 一時的なGoogleスプレッドシートを作成してExcel(.xlsx)形式のBlobとして書き出す。
 * 生成後は一時ファイルをゴミ箱へ移動して片付ける。
 */
function buildExcelBlob(records, equipment) {
  const currentEquip = equipment || '設備';
  const tempName = 'PressReport_tmp_' + new Date().getTime();
  const tempSs = SpreadsheetApp.create(tempName);
  const sheet = tempSs.getSheets()[0];
  sheet.setName('作業日報');

  const totalCount = records.reduce((s, r) => s + (r.count || 0), 0);
  const avgRate = (records.reduce((s, r) => s + parseFloat(r.perfRate || 0), 0) / records.length).toFixed(1);

  const header = ['日付', '設備名', '作業員', '作業人数', '部品番号(品番)', '製造開始時間', '製造終了時間',
    '生産数(pcs)', '材料交換(回)', 'スクラップ交換(回)', '停止理由', '実生産稼働時間', '金型交換時間',
    '計画停止時間', '異常停止時間', '性能稼働率(%)'];

  const rows = [
    [`【${currentEquip}】 プレス生産作業日報`],
    [`出力日時: ${Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')}`, `件数: ${records.length}件`, `総生産数: ${totalCount}pcs`, `平均性能稼働率: ${avgRate}%`],
    header,
  ];

  records.forEach(rec => {
    rows.push([
      rec.date, rec.equipment || currentEquip, rec.operator || '未選択', rec.workerCount || 1, rec.partNo,
      rec.start, rec.end, rec.count, rec.materialCount || 0, rec.scrapCount || 0, rec.stopReason || 'なし',
      secToHms(rec.runningSec), secToHms(rec.moldSec), secToHms(rec.breakSec), secToHms(rec.stopSec), parseFloat(rec.perfRate),
    ]);
  });

  sheet.getRange(1, 1, rows.length, header.length).setValues(
    rows.map(r => { const row = r.slice(); while (row.length < header.length) row.push(''); return row; })
  );
  sheet.getRange(3, 1, 1, header.length).setFontWeight('bold').setBackground('#1a2233').setFontColor('#ffffff');
  sheet.autoResizeColumns(1, header.length);
  SpreadsheetApp.flush();

  const fileId = tempSs.getId();
  const url = `https://docs.google.com/spreadsheets/d/${fileId}/export?format=xlsx`;
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token } });
  const blob = response.getBlob().setName(`プレス作業日報_${currentEquip}_${Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd')}.xlsx`);

  // 一時ファイルは不要なのでゴミ箱へ
  DriveApp.getFileById(fileId).setTrashed(true);

  return blob;
}

function secToHms(sec) {
  sec = sec || 0;
  const h = String(Math.floor(sec / 3600)).padStart(2, '0');
  const m = String(Math.floor((sec % 3600) / 60)).padStart(2, '0');
  const s = String(Math.floor(sec % 60)).padStart(2, '0');
  return `${h}:${m}:${s}`;
}
