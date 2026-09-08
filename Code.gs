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
const SEGMENT_SHEET_NAME = '作業員区間データ';
const PART_MASTER_SHEET = '品番マスタ';
const OPERATOR_MASTER_SHEET = '作業員マスタ';
const STOP_REASON_MASTER_SHEET = '停止理由マスタ';

const HEADER_ROW = [
  '受信日時', 'レコードID', '日付', '設備名', '作業員', '作業人数', '部品番号(品番)',
  '製造開始時間', '製造終了時間', '設定SPM', '生産数(pcs)', '材料交換(回)', 'スクラップ交換(回)',
  '停止理由', '稼働時間(秒)', '金型交換時間(秒)', '計画停止時間(秒)', '異常停止時間(秒)', '性能稼働率(%)',
  '停止理由カテゴリ',
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
      case 'uploadStopReasonMaster':
        return handleUploadStopReasonMaster(payload.reasons);
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

  writeSegmentRows(record);

  return jsonResponse({ status: 'ok', id: record.id });
}

/**
 * 「人数変更を記録」で確定した区間データを、別シート（作業員区間データ）へ書き込む。
 * 同じレコードIDの既存行があれば一旦削除してから書き直す（日報の編集・再送信に対応）。
 */
function writeSegmentRows(record) {
  if (!record.segments || record.segments.length === 0) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SEGMENT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SEGMENT_SHEET_NAME);
    sheet.appendRow(['受信日時', '元レコードID', '日付', '設備名', '部品番号', '区間開始', '区間終了', '作業員', '人数', '生産数(pcs)', '性能稼働率(%)']);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#1a2233').setFontColor('#ffffff');
  }

  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][1]) === String(record.id)) sheet.deleteRow(i + 1);
  }

  record.segments.forEach(seg => {
    sheet.appendRow([
      new Date(), record.id, record.date, record.equipment || '', seg.partNo || record.partNo || '',
      seg.start, seg.end, seg.operator || '未選択', seg.workerCount || 1, seg.count || 0, seg.perfRate || '',
    ]);
  });
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
    rec.stopReasonCategory || 'なし',
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
  } else {
    // 既存シートに新しい列（停止理由カテゴリ等）が無ければ、末尾に自動で追記する
    const currentLen = sheet.getLastColumn();
    if (currentLen < HEADER_ROW.length) {
      const missing = HEADER_ROW.slice(currentLen);
      const range = sheet.getRange(1, currentLen + 1, 1, missing.length);
      range.setValues([missing]);
      range.setFontWeight('bold').setBackground('#1a2233').setFontColor('#ffffff');
    }
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
    stopReasons: readStopReasonMasterSheet(),
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

function readStopReasonMasterSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(STOP_REASON_MASTER_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(STOP_REASON_MASTER_SHEET);
    sheet.appendRow(['停止理由カテゴリ']);
    ['なし', '段取り替え', '材料待ち', '品質確認・不良対応', '設備トラブル', '人員不足・応援待ち', '朝礼・清掃・保全', 'その他']
      .forEach(r => sheet.appendRow([r]));
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 1).setFontWeight('bold').setBackground('#1a2233').setFontColor('#ffffff');
  }
  const data = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    const reason = String(data[i][0] || '').trim();
    if (reason) result.push(reason);
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

function handleUploadStopReasonMaster(reasons) {
  if (!Array.isArray(reasons) || reasons.length === 0) {
    return jsonResponse({ status: 'error', message: 'アップロードするデータがありません' });
  }
  const cleaned = [...new Set(reasons.map(r => String(r || '').trim()).filter(r => r))];
  if (cleaned.length === 0) {
    return jsonResponse({ status: 'error', message: '有効な停止理由が1件もありませんでした' });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(STOP_REASON_MASTER_SHEET);
  if (sheet) { ss.deleteSheet(sheet); }
  sheet = ss.insertSheet(STOP_REASON_MASTER_SHEET);
  sheet.appendRow(['停止理由カテゴリ']);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, 1).setFontWeight('bold').setBackground('#1a2233').setFontColor('#ffffff');
  const values = cleaned.map(r => [r]);
  sheet.getRange(2, 1, values.length, 1).setValues(values);

  return jsonResponse({ status: 'ok', count: cleaned.length });
}
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
    '生産数(pcs)', '材料交換(回)', 'スクラップ交換(回)', '停止理由カテゴリ', '停止理由（詳細）', '実生産稼働時間', '金型交換時間',
    '計画停止時間', '異常停止時間', '性能稼働率(%)'];

  const rows = [
    [`【${currentEquip}】 プレス生産作業日報`],
    [`出力日時: ${Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')}`, `件数: ${records.length}件`, `総生産数: ${totalCount}pcs`, `平均性能稼働率: ${avgRate}%`],
    header,
  ];

  records.forEach(rec => {
    rows.push([
      rec.date, rec.equipment || currentEquip, rec.operator || '未選択', rec.workerCount || 1, rec.partNo,
      rec.start, rec.end, rec.count, rec.materialCount || 0, rec.scrapCount || 0,
      rec.stopReasonCategory || 'なし', rec.stopReasonDetail || '',
      secToHms(rec.runningSec), secToHms(rec.moldSec), secToHms(rec.breakSec), secToHms(rec.stopSec), parseFloat(rec.perfRate),
    ]);
  });

  sheet.getRange(1, 1, rows.length, header.length).setValues(
    rows.map(r => { const row = r.slice(); while (row.length < header.length) row.push(''); return row; })
  );
  sheet.getRange(3, 1, 1, header.length).setFontWeight('bold').setBackground('#1a2233').setFontColor('#ffffff');
  sheet.autoResizeColumns(1, header.length);

  // 区間内訳シート（人数変更の記録単位。未記録の日報は1区間＝シフト全体として出力される）
  const segSheet = tempSs.insertSheet('作業員区間データ');
  const segHeader = ['日付', '設備名', '部品番号', '区間開始', '区間終了', '作業員', '人数', '生産数(pcs)', '性能稼働率(%)'];
  const segRows = [segHeader];
  records.forEach(rec => {
    (rec.segments || []).forEach(seg => {
      segRows.push([
        rec.date, rec.equipment || currentEquip, seg.partNo || rec.partNo, seg.start, seg.end,
        seg.operator || '未選択', seg.workerCount || 1, seg.count || 0, parseFloat(seg.perfRate) || 0,
      ]);
    });
  });
  if (segRows.length > 1) {
    segSheet.getRange(1, 1, segRows.length, segHeader.length).setValues(segRows);
    segSheet.getRange(1, 1, 1, segHeader.length).setFontWeight('bold').setBackground('#1a2233').setFontColor('#ffffff');
    segSheet.autoResizeColumns(1, segHeader.length);
  }
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

/* =========================================================
   7) ダッシュボード生成（現場用・報告用）
   蓄積された「日報データ」「作業員区間データ」を集計し、
   2つのダッシュボードシートを作成・更新する。
   ライブ数式ではなく、実行時点の集計値を書き込むスナップショット方式。
   スプレッドシートのメニュー「PRESS MONITOR」→「ダッシュボードを更新」、
   または自動更新（毎日6時）で最新化される。
   ========================================================= */
const DASHBOARD_FIELD_SHEET = 'ダッシュボード_現場用';
const DASHBOARD_REPORT_SHEET = 'ダッシュボード_報告用';
const DASHBOARD_ALL_EQUIPMENT = '(全設備)';

/** スプレッドシートを開いたときに専用メニューを追加する（シンプルトリガー） */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('PRESS MONITOR')
    .addItem('ダッシュボードを更新', 'buildDashboards')
    .addSeparator()
    .addItem('自動更新を設定（毎日6時）', 'createDashboardTrigger')
    .addItem('自動更新を停止', 'removeDashboardTriggers')
    .addToUi();
}

/** 毎日決まった時刻にダッシュボードを自動更新するトリガーを設定する */
function createDashboardTrigger() {
  removeDashboardTriggers();
  ScriptApp.newTrigger('buildDashboards').timeBased().everyDays(1).atHour(6).create();
  SpreadsheetApp.getUi().alert('毎日6時に自動更新するよう設定しました。');
}

/** 自動更新トリガーを停止する */
function removeDashboardTriggers() {
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'buildDashboards') { ScriptApp.deleteTrigger(t); removed++; }
  });
  if (removed > 0) SpreadsheetApp.getUi().alert('自動更新を停止しました。');
}

/**
 * ダッシュボードの本体。メニュー・トリガー・手動実行のいずれからも呼び出せる。
 * Apps Scriptエディタで直接この関数を選んで実行することもできる。
 */
function buildDashboards() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportRows = readReportRows_(ss);
  const segmentRows = readSegmentRows_(ss);
  const equipmentList = uniqueValues_(reportRows.map(r => r.equipment)).sort();

  // 前回選んでいた設備フィルタを可能な範囲で引き継ぐ
  const fieldSheetExisting = ss.getSheetByName(DASHBOARD_FIELD_SHEET);
  const previousFilter = fieldSheetExisting ? String(fieldSheetExisting.getRange('B2').getValue() || DASHBOARD_ALL_EQUIPMENT) : DASHBOARD_ALL_EQUIPMENT;
  const currentFilter = (previousFilter === DASHBOARD_ALL_EQUIPMENT || equipmentList.indexOf(previousFilter) >= 0)
    ? previousFilter : DASHBOARD_ALL_EQUIPMENT;

  buildFieldDashboard_(ss, reportRows, segmentRows, equipmentList, currentFilter);
  buildReportDashboard_(ss, reportRows, equipmentList);
}

/* ---------- データ読み込み ---------- */
function readReportRows_(ss) {
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADER_ROW.length).getValues();
  return values.map(r => ({
    date: String(r[2] || ''), equipment: String(r[3] || ''), operator: String(r[4] || ''),
    workerCount: Number(r[5]) || 0, partNo: String(r[6] || ''), count: Number(r[10]) || 0,
    materialCount: Number(r[11]) || 0, scrapCount: Number(r[12]) || 0, stopReason: String(r[13] || ''),
    runningSec: Number(r[14]) || 0, moldSec: Number(r[15]) || 0, breakSec: Number(r[16]) || 0,
    stopSec: Number(r[17]) || 0, perfRate: Number(r[18]) || 0, stopReasonCategory: String(r[19] || 'なし'),
  })).filter(r => r.date);
}

function readSegmentRows_(ss) {
  const sheet = ss.getSheetByName(SEGMENT_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues();
  return values.map(r => ({
    date: String(r[2] || ''), equipment: String(r[3] || ''), partNo: String(r[4] || ''),
    start: String(r[5] || ''), end: String(r[6] || ''), operator: String(r[7] || ''),
    workerCount: Number(r[8]) || 0, count: Number(r[9]) || 0, perfRate: Number(r[10]) || 0,
  })).filter(r => r.date);
}

function uniqueValues_(arr) {
  return [...new Set(arr.filter(v => v))];
}

function filterByEquipment_(rows, equipment) {
  if (!equipment || equipment === DASHBOARD_ALL_EQUIPMENT) return rows;
  return rows.filter(r => r.equipment === equipment);
}

/* ---------- 集計 ---------- */
function aggregateAvgByKey_(rows, keyFn, valueFn) {
  const map = {};
  rows.forEach(r => {
    const k = keyFn(r);
    if (!k) return;
    if (!map[k]) map[k] = { sum: 0, n: 0 };
    map[k].sum += valueFn(r);
    map[k].n += 1;
  });
  return Object.keys(map).sort().map(k => [k, Math.round((map[k].sum / map[k].n) * 10) / 10, map[k].n]);
}

function aggregateStopReasons_(rows) {
  const map = {};
  rows.forEach(r => {
    const k = r.stopReasonCategory || 'なし';
    if (!map[k]) map[k] = { count: 0, sec: 0 };
    map[k].count += 1;
    map[k].sec += r.stopSec;
  });
  return Object.keys(map)
    .map(k => [k, map[k].count, Math.round(map[k].sec / 60)])
    .sort((a, b) => b[2] - a[2]);
}

function aggregateStatusTotals_(rows) {
  const totals = { running: 0, mold: 0, brk: 0, stop: 0 };
  rows.forEach(r => { totals.running += r.runningSec; totals.mold += r.moldSec; totals.brk += r.breakSec; totals.stop += r.stopSec; });
  return [
    ['稼働中', Math.round(totals.running / 60)],
    ['金型交換', Math.round(totals.mold / 60)],
    ['計画停止', Math.round(totals.brk / 60)],
    ['異常停止', Math.round(totals.stop / 60)],
  ];
}

/* ---------- シート構築ヘルパー ---------- */
function clearSheet_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (sheet) {
    sheet.getCharts().forEach(c => sheet.removeChart(c));
    sheet.clear();
  } else {
    sheet = ss.insertSheet(name);
  }
  sheet.setTabColor('#0f9d63');
  return sheet;
}

function writeTable_(sheet, startRow, startCol, headers, rows, title) {
  let r = startRow;
  if (title) {
    sheet.getRange(r, startCol, 1, 1).setValue(title).setFontWeight('bold').setFontSize(11).setFontColor('#1a2233');
    r += 1;
  }
  sheet.getRange(r, startCol, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#1a2233').setFontColor('#ffffff');
  r += 1;
  if (rows.length > 0) {
    sheet.getRange(r, startCol, rows.length, headers.length).setValues(rows);
  }
  return { headerRow: r - 1, firstDataRow: r, lastDataRow: r + Math.max(rows.length, 1) - 1 };
}

/* ---------- 現場用ダッシュボード ---------- */
function buildFieldDashboard_(ss, reportRows, segmentRows, equipmentList, currentFilter) {
  const sheet = clearSheet_(ss, DASHBOARD_FIELD_SHEET);
  sheet.getRange('A1:H1').merge().setValue('現場用ダッシュボード').setFontWeight('bold').setFontSize(16).setFontColor('#ffffff').setBackground('#1a2233');
  sheet.getRange('A2').setValue('設備で絞り込み：').setFontWeight('bold');
  const b2 = sheet.getRange('B2').setValue(currentFilter);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList([DASHBOARD_ALL_EQUIPMENT].concat(equipmentList), true)
    .setAllowInvalid(false)
    .build();
  b2.setDataValidation(rule);
  sheet.getRange('D2').setValue('最終更新: ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') +
    '　※設備を変更した場合は「PRESS MONITOR」メニューから再度更新してください').setFontColor('#9AA5B1').setFontSize(9);

  const filtered = filterByEquipment_(reportRows, currentFilter);
  const filteredSeg = filterByEquipment_(segmentRows, currentFilter);

  // ① 日別 性能稼働率の推移
  const byDate = aggregateAvgByKey_(filtered, r => r.date, r => r.perfRate);
  const t1 = writeTable_(sheet, 4, 1, ['日付', '平均性能稼働率(%)', '件数'], byDate, '① 日別 性能稼働率の推移');
  if (byDate.length > 0) {
    sheet.insertChart(sheet.newChart().setChartType(Charts.ChartType.LINE)
      .addRange(sheet.getRange(t1.headerRow, 1, byDate.length + 1, 2))
      .setPosition(4, 5, 0, 0)
      .setOption('title', '日別 性能稼働率の推移')
      .setOption('legend', { position: 'none' })
      .setOption('vAxis', { title: '%', minValue: 0 })
      .build());
  }

  // ② 品番別 平均性能稼働率
  const byPart = aggregateAvgByKey_(filtered, r => r.partNo, r => r.perfRate);
  const startRow2 = t1.lastDataRow + 3;
  const t2 = writeTable_(sheet, startRow2, 1, ['品番', '平均性能稼働率(%)', '件数'], byPart, '② 品番別 平均性能稼働率');
  if (byPart.length > 0) {
    sheet.insertChart(sheet.newChart().setChartType(Charts.ChartType.COLUMN)
      .addRange(sheet.getRange(t2.headerRow, 1, byPart.length + 1, 2))
      .setPosition(startRow2, 5, 0, 0)
      .setOption('title', '品番別 平均性能稼働率')
      .setOption('legend', { position: 'none' })
      .build());
  }

  // ③ 停止理由カテゴリ別（異常停止時間の多い順）
  const byReason = aggregateStopReasons_(filtered);
  const startRow3 = t2.lastDataRow + 3;
  const t3 = writeTable_(sheet, startRow3, 1, ['停止理由カテゴリ', '発生回数', '異常停止合計(分)'], byReason, '③ 停止理由カテゴリ別（異常停止時間の多い順）');
  if (byReason.length > 0) {
    sheet.insertChart(sheet.newChart().setChartType(Charts.ChartType.BAR)
      .addRange(sheet.getRange(t3.headerRow, 1, byReason.length + 1, 1))
      .addRange(sheet.getRange(t3.headerRow, 3, byReason.length + 1, 1))
      .setPosition(startRow3, 5, 0, 0)
      .setOption('title', '停止理由カテゴリ別 異常停止時間(分)')
      .setOption('legend', { position: 'none' })
      .build());
  }

  // ④ 稼働時間の内訳
  const statusTotals = aggregateStatusTotals_(filtered);
  const startRow4 = t3.lastDataRow + 3;
  const t4 = writeTable_(sheet, startRow4, 1, ['区分', '合計時間(分)'], statusTotals, '④ 稼働時間の内訳');
  sheet.insertChart(sheet.newChart().setChartType(Charts.ChartType.PIE)
    .addRange(sheet.getRange(t4.headerRow, 1, statusTotals.length + 1, 2))
    .setPosition(startRow4, 5, 0, 0)
    .setOption('title', '稼働時間の内訳（分）')
    .build());

  // ⑤ 作業人数別 平均性能稼働率（区間データより）
  const byWorkerCount = aggregateAvgByKey_(filteredSeg, r => String(r.workerCount) + '名', r => r.perfRate)
    .sort((a, b) => parseInt(a[0]) - parseInt(b[0]));
  const startRow5 = t4.lastDataRow + 3;
  const t5 = writeTable_(sheet, startRow5, 1, ['作業人数', '平均性能稼働率(%)', '区間数'], byWorkerCount,
    '⑤ 作業人数別 平均性能稼働率（「人数変更を記録」区間データより）');
  if (byWorkerCount.length > 0) {
    sheet.insertChart(sheet.newChart().setChartType(Charts.ChartType.COLUMN)
      .addRange(sheet.getRange(t5.headerRow, 1, byWorkerCount.length + 1, 2))
      .setPosition(startRow5, 5, 0, 0)
      .setOption('title', '作業人数別 平均性能稼働率')
      .setOption('legend', { position: 'none' })
      .setOption('vAxis', { title: '%', minValue: 0 })
      .build());
  }
  sheet.getRange(startRow5 + Math.max(byWorkerCount.length, 1) + 2, 1).setValue(
    '※区間数が少ないうちは参考値としてご覧ください。「人数変更を記録」を使わなかった日報は、シフト全体を1区間として集計しています。'
  ).setFontColor('#9AA5B1').setFontSize(9);

  sheet.setColumnWidth(1, 170);
  sheet.setColumnWidths(2, 3, 110);
}

/* ---------- 報告用ダッシュボード ---------- */
function buildReportDashboard_(ss, reportRows, equipmentList) {
  const sheet = clearSheet_(ss, DASHBOARD_REPORT_SHEET);
  sheet.getRange('A1:F1').merge().setValue('報告用ダッシュボード（サマリー）').setFontWeight('bold').setFontSize(16).setFontColor('#ffffff').setBackground('#1a2233');
  sheet.getRange('A2').setValue('最終更新: ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')).setFontColor('#9AA5B1').setFontSize(9);

  const totalCount = reportRows.reduce((s, r) => s + r.count, 0);
  const avgRate = reportRows.length ? Math.round((reportRows.reduce((s, r) => s + r.perfRate, 0) / reportRows.length) * 10) / 10 : 0;
  const summary = [
    ['対象期間の日報件数', reportRows.length],
    ['総生産数(pcs)', totalCount],
    ['平均性能稼働率(%)', avgRate],
    ['対象設備数', equipmentList.length],
  ];
  const t0 = writeTable_(sheet, 4, 1, ['指標', '値'], summary, 'サマリー');

  // 月次 性能稼働率の推移
  const byMonth = aggregateAvgByKey_(reportRows, r => r.date.slice(0, 7), r => r.perfRate);
  const startRow1 = t0.lastDataRow + 3;
  const t1 = writeTable_(sheet, startRow1, 1, ['年月', '平均性能稼働率(%)', '件数'], byMonth, '月次 性能稼働率の推移');
  if (byMonth.length > 0) {
    sheet.insertChart(sheet.newChart().setChartType(Charts.ChartType.LINE)
      .addRange(sheet.getRange(t1.headerRow, 1, byMonth.length + 1, 2))
      .setPosition(startRow1, 4, 0, 0)
      .setOption('title', '月次 性能稼働率の推移')
      .setOption('legend', { position: 'none' })
      .setOption('vAxis', { title: '%', minValue: 0 })
      .build());
  }

  // 設備別比較
  const byEquip = aggregateAvgByKey_(reportRows, r => r.equipment, r => r.perfRate);
  const startRow2 = t1.lastDataRow + 3;
  const t2 = writeTable_(sheet, startRow2, 1, ['設備名', '平均性能稼働率(%)', '件数'], byEquip, '設備別 比較');
  if (byEquip.length > 0) {
    sheet.insertChart(sheet.newChart().setChartType(Charts.ChartType.COLUMN)
      .addRange(sheet.getRange(t2.headerRow, 1, byEquip.length + 1, 2))
      .setPosition(startRow2, 4, 0, 0)
      .setOption('title', '設備別 平均性能稼働率')
      .setOption('legend', { position: 'none' })
      .build());
  }

  sheet.setColumnWidth(1, 170);
  sheet.setColumnWidths(2, 2, 110);
}
