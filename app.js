/* ==========================================================
   PRESS MONITOR — メインロジック（状態管理・画面制御）
   ========================================================== */

let settings = loadSettings();
let masters = loadMasters();

let state = {
  isShiftActive: false,
  status: 'IDLE',
  startTime: null,
  endTime: null,
  runningSeconds: 0,
  moldSeconds: 0,
  breakSeconds: 0,
  stopSeconds: 0,
  accumulatedCount: 0,
  settingSPM: 15.0,
  targetCount: CONFIG.INITIAL_TARGET_COUNT,
  partNo: '',
};

let historyLogs = [];
let checkedRecordIds = new Set();
let lastTickTime = null;

/* ---------- 要素参照 ---------- */
const rateEl = document.getElementById('perf-rate');
const clockEl = document.getElementById('main-clock');
const countEl = document.getElementById('total-count');
const statusChip = document.getElementById('status-chip');
const clockSub = document.getElementById('clock-sub');
const rateBox = document.getElementById('rate-box');
const timeStartLabel = document.getElementById('time-start-label');
const timeEndLabel = document.getElementById('time-end-label');
const timeMoldEl = document.getElementById('time-mold');
const timeBreakEl = document.getElementById('time-break');
const timeStopEl = document.getElementById('time-stop');
const runBtn = document.getElementById('btn-RUNNING');
const runIcon = document.getElementById('run-icon');
const runText = document.getElementById('run-text');
const btnStartShift = document.getElementById('btn-start-shift');
const btnEndShift = document.getElementById('btn-end-shift');
const partNoSelect = document.getElementById('part-no-select');
const spmSettingInput = document.getElementById('spm-setting');
const targetInput = document.getElementById('target-count');
const targetDisplayVal = document.getElementById('target-val');
const reportEmailInput = document.getElementById('report-email');
const equipmentSelect = document.getElementById('equipment-select');
const opSlotInputs = [
  document.getElementById('op-slot-0'),
  document.getElementById('op-slot-1'),
  document.getElementById('op-slot-2'),
  document.getElementById('op-slot-3'),
];
const operatorDatalist = document.getElementById('operator-master-list');

/* =========================================================
   初期化
   ========================================================= */
window.onload = function () {
  applySettingsToUI();
  renderEquipmentOptions();
  renderPartOptions();
  renderOperatorDatalist();
  renderMasterStatus();

  try {
    const saved = localStorage.getItem(STORAGE.HISTORY);
    if (saved) historyLogs = JSON.parse(saved);
  } catch (e) {}

  spmSettingInput.value = state.settingSPM.toFixed(1);
  targetInput.value = CONFIG.INITIAL_TARGET_COUNT;

  partNoSelect.addEventListener('change', applySelectedPart);

  spmSettingInput.addEventListener('input', (e) => {
    state.settingSPM = parseFloat(e.target.value) || 0.1;
    updateCalculations();
  });

  targetInput.addEventListener('input', (e) => {
    state.targetCount = parseInt(e.target.value) || 1;
    targetDisplayVal.innerText = state.targetCount.toLocaleString();
  });

  reportEmailInput.addEventListener('change', (e) => {
    settings.defaultEmail = e.target.value.trim();
    persistSettings();
  });

  equipmentSelect.addEventListener('change', (e) => {
    localStorage.setItem('pm_last_equipment_v1', e.target.value);
  });
  const savedEquip = localStorage.getItem('pm_last_equipment_v1');
  if (savedEquip && settings.equipmentList.includes(savedEquip)) equipmentSelect.value = savedEquip;

  // 作業員4窓の入力補助（予測変換の候補はrenderOperatorDatalistで供給）
  opSlotInputs.forEach(input => {
    input.addEventListener('change', () => input.value = input.value.trim());
  });

  lastTickTime = Date.now();
  setInterval(updateClock, 1000);
  setInterval(onTimerTick, 1000);
  renderHistoryTable();
  updateCloudStatusChip();
  updateSyncBadge();

  setInterval(() => retryPendingSync(false), CONFIG.SYNC_RETRY_INTERVAL_MS);
  setInterval(() => fetchMastersFromCloud(false), CONFIG.MASTER_AUTO_REFRESH_MS);
  window.addEventListener('online', () => { retryPendingSync(false); fetchMastersFromCloud(false); });

  // 起動時にバックグラウンドで最新マスタを取得（失敗してもキャッシュで動作継続）
  fetchMastersFromCloud(false);
};

/* =========================================================
   設定（ローカル永続化）
   ========================================================= */
function loadSettings() {
  try {
    const raw = localStorage.getItem(STORAGE.SETTINGS);
    if (raw) return Object.assign({}, DEFAULT_SETTINGS, JSON.parse(raw));
  } catch (e) {}
  return JSON.parse(JSON.stringify(DEFAULT_SETTINGS));
}

function persistSettings() {
  localStorage.setItem(STORAGE.SETTINGS, JSON.stringify(settings));
  updateCloudStatusChip();
}

function applySettingsToUI() {
  reportEmailInput.value = settings.defaultEmail || '';
}

/* ---- マスタデータ（品番・作業員）のローカルキャッシュ ---- */
function loadMasters() {
  try {
    const raw = localStorage.getItem(STORAGE.MASTERS);
    if (raw) return Object.assign({}, DEFAULT_MASTERS, JSON.parse(raw));
  } catch (e) {}
  return JSON.parse(JSON.stringify(DEFAULT_MASTERS));
}
function persistMasters() {
  localStorage.setItem(STORAGE.MASTERS, JSON.stringify(masters));
}

function renderPartOptions() {
  const list = masters.partMasters || [];
  if (list.length === 0) {
    partNoSelect.innerHTML = '<option value="">（品番マスタ未取得）</option>';
    partNoSelect.disabled = true;
    return;
  }
  partNoSelect.disabled = false;
  const current = partNoSelect.value;
  partNoSelect.innerHTML = list.map(p => `<option value="${escapeHtml(p.part)}">${escapeHtml(p.part)}（SPM ${p.spm}）</option>`).join('');
  const stillExists = list.some(p => p.part === current);
  partNoSelect.value = stillExists ? current : list[0].part;
  applySelectedPart();
}

function applySelectedPart() {
  const list = masters.partMasters || [];
  const selected = list.find(p => p.part === partNoSelect.value);
  if (selected) {
    state.partNo = selected.part;
    state.settingSPM = selected.spm;
    spmSettingInput.value = Number(selected.spm).toFixed(1);
    updateCalculations();
  }
}

function renderMasterStatus() {
  const el = document.getElementById('settings-master-status');
  if (!el) return;
  if (!masters.lastSyncedAt) {
    el.innerText = '未取得（クラウド未同期）';
    el.className = 'text-[10px] text-amber-400';
  } else {
    const d = new Date(masters.lastSyncedAt);
    el.innerText = `最終更新: ${d.toLocaleString('ja-JP')}（品番${masters.partMasters.length}件／作業員${masters.operators.length}名）`;
    el.className = 'text-[10px] text-slate-500';
  }
}

function renderEquipmentOptions() {
  const current = equipmentSelect.value;
  equipmentSelect.innerHTML = settings.equipmentList.map(eq => `<option value="${escapeHtml(eq)}">${escapeHtml(eq)}</option>`).join('');
  if (settings.equipmentList.includes(current)) equipmentSelect.value = current;
}

/* ---- 設定モーダル ---- */
function openSettingsModal() {
  document.getElementById('settings-gas-url').value = settings.gasUrl || '';
  document.getElementById('settings-default-email').value = settings.defaultEmail || '';
  document.getElementById('settings-gas-test-result').innerText = '';
  renderMasterStatus();
  renderSettingsEquipmentList();
  showModal('settings-modal');
}
function closeSettingsModal() { hideModal('settings-modal'); }

function renderSettingsEquipmentList() {
  const box = document.getElementById('settings-equipment-list');
  box.innerHTML = settings.equipmentList.map((eq, idx) => `
    <div class="flex gap-1.5 items-center">
      <input type="text" value="${escapeHtml(eq)}" data-idx="${idx}" onchange="updateEquipmentField(this)" class="field flex-1 p-1.5 text-xs">
      <button onclick="removeEquipmentRow(${idx})" class="text-rose-400 hover:text-rose-300 w-7 h-7 flex items-center justify-center"><i class="fa-solid fa-trash-can text-xs"></i></button>
    </div>`).join('') || '<p class="text-[10px] text-slate-500">設備が登録されていません</p>';
}
function updateEquipmentField(el) { settings.equipmentList[parseInt(el.dataset.idx)] = el.value.trim(); }
function addEquipmentRow() { settings.equipmentList.push('新規設備'); renderSettingsEquipmentList(); }
function removeEquipmentRow(idx) { settings.equipmentList.splice(idx, 1); renderSettingsEquipmentList(); }

function saveSettings() {
  settings.gasUrl = document.getElementById('settings-gas-url').value.trim();
  settings.defaultEmail = document.getElementById('settings-default-email').value.trim();
  settings.equipmentList = settings.equipmentList.filter(e => e);
  persistSettings();
  applySettingsToUI();
  renderEquipmentOptions();
  closeSettingsModal();
  showToast('設定を保存しました');
  retryPendingSync(false);
}

/* =========================================================
   作業員（クラウドマスタから選択）
   ========================================================= */
function renderOperatorDatalist() {
  const list = masters.operators || [];
  operatorDatalist.innerHTML = list.map(name => `<option value="${escapeHtml(name)}"></option>`).join('');
}
function getSelectedOperatorsText() {
  const names = opSlotInputs.map(i => i.value.trim()).filter(n => n);
  return names.length ? names.join('・') : '未選択';
}
function getSelectedOperatorsCount() {
  return opSlotInputs.filter(i => i.value.trim()).length;
}
function clearOperatorSlots() {
  opSlotInputs.forEach(i => i.value = '');
}

/* =========================================================
   タイマー / 計算
   ========================================================= */
function updateClock() {
  document.getElementById('live-time').innerText = new Date().toTimeString().split(' ')[0];
}

function onTimerTick() {
  const now = Date.now();
  if (!lastTickTime) lastTickTime = now;
  const deltaSec = (now - lastTickTime) / 1000;
  lastTickTime = now;
  if (!state.isShiftActive) return;

  if (state.status === 'RUNNING') {
    state.runningSeconds += deltaSec;
    state.accumulatedCount += (state.settingSPM / 60) * deltaSec;
  } else if (state.status === 'MOLD') {
    state.moldSeconds += deltaSec;
  } else if (state.status === 'BREAK') {
    state.breakSeconds += deltaSec;
  } else if (state.status === 'STOP') {
    state.stopSeconds += deltaSec;
  }

  updateCalculations();
  clockEl.innerText = formatTime(Math.floor(state.runningSeconds));
  countEl.innerText = Math.floor(state.accumulatedCount).toLocaleString();
  timeMoldEl.innerText = formatTime(Math.floor(state.moldSeconds));
  timeBreakEl.innerText = formatTime(Math.floor(state.breakSeconds));
  timeStopEl.innerText = formatTime(Math.floor(state.stopSeconds));
}

function updateCalculations() {
  if (state.settingSPM <= 0) return;
  const activeWorkSec = state.runningSeconds + state.moldSeconds + state.stopSeconds;
  if (activeWorkSec === 0) { rateEl.innerText = '100.0'; rateEl.style.color = '#10b981'; return; }
  const stdCT = 60 / state.settingSPM;
  let rate = ((stdCT * state.accumulatedCount) / activeWorkSec) * 100;
  rate = Math.min(rate, 200.0);
  rateEl.innerText = rate.toFixed(1);
  rateEl.style.color = rate >= CONFIG.PERFORMANCE_GOOD_LIMIT ? '#10b981' : (rate >= CONFIG.PERFORMANCE_NORMAL_LIMIT ? '#f2a93c' : '#f0556a');
}

/* =========================================================
   ステータス制御
   ※ btn-RUNNING の見た目切り替えは classList の個別操作のみで行う。
     className をまるごと書き換えると pointer-events-none が
     誤って残り、ボタンが反応しなくなる不具合の原因になるため禁止。
   ========================================================= */
function setStatus(newStatus) {
  document.querySelectorAll('.modebtn').forEach(b => b.classList.remove('is-on'));
  statusChip.className = 'status-chip';

  if (state.status === 'RUNNING' && newStatus === 'RUNNING') {
    state.status = 'IDLE';
    statusChip.innerText = '一時停止中';
    statusChip.classList.add('st-IDLE');
    clockSub.innerText = '一時停止中';
    runBtn.classList.remove('mb-pause');
    runBtn.classList.add('mb-running', 'is-on');
    runIcon.className = 'fa-solid fa-circle-play text-3xl sm:text-4xl';
    runText.innerText = '稼働再開';
    return;
  }

  if (state.status === newStatus) {
    state.status = 'IDLE';
    statusChip.innerText = '待機中';
    statusChip.classList.add('st-IDLE');
    clockSub.innerText = '停止中';
    resetRunButton('稼働再開');
  } else {
    state.status = newStatus;
    document.getElementById('btn-' + newStatus).classList.add('is-on');
    const labels = { RUNNING: '生産中', MOLD: '金型交換中', BREAK: '計画停止中', STOP: '異常停止中' };
    statusChip.innerText = labels[newStatus];
    statusChip.classList.add('st-' + newStatus);

    if (newStatus === 'RUNNING') {
      clockSub.innerText = '稼働中';
      runBtn.classList.remove('mb-running');
      runBtn.classList.add('mb-pause', 'is-on');
      runIcon.className = 'fa-solid fa-circle-pause text-3xl sm:text-4xl';
      runText.innerText = '一時停止';
    } else {
      clockSub.innerText = '非稼働（間接作業）';
      resetRunButton('稼働再開');
    }
  }
}

function resetRunButton(label = '稼働中') {
  runBtn.classList.remove('mb-pause', 'is-on');
  runBtn.classList.add('mb-running');
  runIcon.className = 'fa-solid fa-circle-play text-3xl sm:text-4xl';
  runText.innerText = label;
}

/* =========================================================
   製造開始 / 終了
   ========================================================= */
function startMfg() {
  if (state.isShiftActive) return;
  state.isShiftActive = true;
  state.startTime = new Date();
  lastTickTime = Date.now();

  timeStartLabel.innerText = state.startTime.toTimeString().split(' ')[0];
  timeEndLabel.innerText = '製造中…';
  btnStartShift.classList.add('opacity-50', 'pointer-events-none');
  btnEndShift.classList.remove('opacity-50', 'pointer-events-none');

  document.querySelectorAll('.modebtn').forEach(b => b.classList.remove('pointer-events-none'));

  setStatus('IDLE');
  clockSub.innerText = '稼働開始ボタンを押してください';
  showToast('製造を開始しました');
}

function openEndShiftModal() {
  if (!state.isShiftActive) return;
  document.getElementById('modal-end-time').innerText = new Date().toTimeString().split(' ')[0];
  document.getElementById('modal-actual-count').value = Math.floor(state.accumulatedCount);
  document.getElementById('modal-material-count').value = 0;
  document.getElementById('modal-scrap-count').value = 0;
  document.getElementById('modal-stop-reason').value = '';
  const workerText = getSelectedOperatorsText();
  const workerCount = getSelectedOperatorsCount();
  document.getElementById('modal-worker-preview').innerText = `${workerText}（${workerCount}名）`;
  showModal('end-shift-modal');
}
function closeEndShiftModal() { hideModal('end-shift-modal'); }

function confirmEndMfg() {
  state.endTime = new Date();
  const actualCount = parseInt(document.getElementById('modal-actual-count').value) || 0;
  const materialCount = parseInt(document.getElementById('modal-material-count').value) || 0;
  const scrapCount = parseInt(document.getElementById('modal-scrap-count').value) || 0;
  const stopReason = document.getElementById('modal-stop-reason').value.trim() || 'なし';

  closeEndShiftModal();
  saveCurrentToHistory(actualCount, materialCount, scrapCount, stopReason);
  executeResetSilently();
}

function executeResetSilently() {
  state.status = 'IDLE';
  state.isShiftActive = false;
  state.startTime = null;
  state.endTime = null;
  state.runningSeconds = 0;
  state.moldSeconds = 0;
  state.breakSeconds = 0;
  state.stopSeconds = 0;
  state.accumulatedCount = 0;
  clearOperatorSlots();

  document.querySelectorAll('.modebtn').forEach(b => { b.classList.remove('is-on'); b.classList.add('pointer-events-none'); });
  statusChip.className = 'status-chip st-IDLE';
  statusChip.innerText = '待機中';
  clockSub.innerText = '製造開始ボタンを押してください';
  timeStartLabel.innerText = '未記録';
  timeEndLabel.innerText = '未記録';
  clockEl.innerText = '00:00:00';
  countEl.innerText = '0';
  timeMoldEl.innerText = '00:00:00';
  timeBreakEl.innerText = '00:00:00';
  timeStopEl.innerText = '00:00:00';
  rateEl.innerText = '100.0';
  rateEl.style.color = '#10b981';
  btnStartShift.classList.remove('opacity-50', 'pointer-events-none');
  btnEndShift.classList.add('opacity-50', 'pointer-events-none');
  resetRunButton('稼働中');
}

function openResetModal() { showModal('reset-modal'); }
function closeResetModal() { hideModal('reset-modal'); }
function confirmReset() { closeResetModal(); executeResetSilently(); showToast('本日の稼働データをリセットしました'); }

/* =========================================================
   日報保存＋クラウド同期
   ========================================================= */
function saveCurrentToHistory(finalCount, materialCount, scrapCount, stopReason) {
  const today = new Date();
  const dateStr = today.getFullYear() + '/' + String(today.getMonth() + 1).padStart(2, '0') + '/' + String(today.getDate()).padStart(2, '0');
  const countValue = finalCount ?? Math.floor(state.accumulatedCount);
  const totalActiveSec = state.runningSeconds + state.moldSeconds + state.stopSeconds;

  let calculatedRate = '100.0';
  if (totalActiveSec > 0 && state.settingSPM > 0) {
    const stdCT = 60 / state.settingSPM;
    calculatedRate = Math.min(((stdCT * countValue) / totalActiveSec) * 100, 200.0).toFixed(1);
  }

  const record = {
    id: Date.now(),
    date: dateStr,
    equipment: equipmentSelect.value || settings.equipmentList[0] || '未設定',
    operator: getSelectedOperatorsText(),
    workerCount: getSelectedOperatorsCount(),
    partNo: state.partNo || partNoSelect.value || '未設定',
    start: state.startTime ? state.startTime.toTimeString().split(' ')[0] : '未記録',
    end: state.endTime ? state.endTime.toTimeString().split(' ')[0] : '未記録',
    settingSPM: state.settingSPM.toFixed(1),
    count: countValue,
    materialCount: materialCount || 0,
    scrapCount: scrapCount || 0,
    stopReason: stopReason || 'なし',
    runningSec: Math.floor(state.runningSeconds),
    moldSec: Math.floor(state.moldSeconds),
    breakSec: Math.floor(state.breakSeconds),
    stopSec: Math.floor(state.stopSeconds),
    perfRate: calculatedRate,
    syncStatus: 'pending',
  };

  historyLogs.unshift(record);
  checkedRecordIds.add(record.id);
  persistHistory();
  renderHistoryTable();
  updateSyncBadge();
  showToast(`日報データ（${record.partNo}：生産数 ${countValue}pcs）を保存しました`);

  syncRecordToCloud(record);
}

function persistHistory() {
  try { localStorage.setItem(STORAGE.HISTORY, JSON.stringify(historyLogs)); } catch (e) {}
}

/* =========================================================
   選択・履歴テーブル
   ========================================================= */
function toggleSelectAll(master) {
  if (master.checked) historyLogs.forEach(r => checkedRecordIds.add(r.id));
  else checkedRecordIds.clear();
  renderHistoryTable();
}
function toggleRecordCheck(id, checkbox) {
  if (checkbox.checked) checkedRecordIds.add(id); else checkedRecordIds.delete(id);
  updateMasterCheckboxState();
}
function updateMasterCheckboxState() {
  const master = document.getElementById('check-all');
  if (master && historyLogs.length > 0) master.checked = checkedRecordIds.size === historyLogs.length;
}

function syncStatusIcon(rec) {
  if (rec.syncStatus === 'synced') return '<i class="fa-solid fa-cloud-check sync-ok" title="クラウド同期済み"></i>';
  if (rec.syncStatus === 'error') return '<i class="fa-solid fa-triangle-exclamation sync-err" title="送信エラー"></i>';
  return '<i class="fa-solid fa-clock sync-pending" title="同期待ち"></i>';
}

function renderHistoryTable() {
  const tbody = document.getElementById('history-body');
  if (historyLogs.length === 0) {
    tbody.innerHTML = '<tr><td colspan="18" class="p-6 text-center text-slate-500 font-bold">保存された日報データはありません。</td></tr>';
    return;
  }
  tbody.innerHTML = historyLogs.map(rec => {
    const isChecked = checkedRecordIds.has(rec.id);
    return `
      <tr class="row-hover transition-all ${isChecked ? 'bg-white/5' : ''}">
        <td class="p-2 text-center" onclick="event.stopPropagation()"><input type="checkbox" value="${rec.id}" ${isChecked ? 'checked' : ''} onclick="toggleRecordCheck(${rec.id}, this)" class="w-4 h-4 rounded border-slate-700 bg-black text-emerald-500"></td>
        <td class="p-2 font-mono text-white">${rec.date}</td>
        <td class="p-2 font-mono text-amber-300 font-bold">${escapeHtml(rec.operator || '未選択')}</td>
        <td class="p-2 font-mono text-center text-amber-200 font-bold">${rec.workerCount || 1}名</td>
        <td class="p-2 font-mono text-emerald-400 font-bold">${escapeHtml(rec.partNo)}</td>
        <td class="p-2 font-mono text-slate-300">${rec.start}</td>
        <td class="p-2 font-mono text-slate-300">${rec.end}</td>
        <td class="p-2 font-mono font-bold text-white">${rec.count}</td>
        <td class="p-2 font-mono text-amber-300 font-bold">${rec.materialCount || 0}</td>
        <td class="p-2 font-mono text-purple-300 font-bold">${rec.scrapCount || 0}</td>
        <td class="p-2 text-slate-300 truncate max-w-[100px]">${escapeHtml(rec.stopReason || 'なし')}</td>
        <td class="p-2 font-mono">${formatTime(rec.runningSec)}</td>
        <td class="p-2 font-mono text-amber-400">${formatTime(rec.moldSec)}</td>
        <td class="p-2 font-mono text-sky-400">${formatTime(rec.breakSec)}</td>
        <td class="p-2 font-mono text-rose-500">${formatTime(rec.stopSec)}</td>
        <td class="p-2 font-mono font-bold text-emerald-400 text-right text-sm">${rec.perfRate}%</td>
        <td class="p-2 text-center text-sm">${syncStatusIcon(rec)}</td>
        <td class="p-2 text-center">
          <div class="flex justify-center gap-1.5" onclick="event.stopPropagation()">
            <button class="bg-amber-950/60 hover:bg-amber-900 text-amber-300 p-1.5 rounded active:scale-90" onclick="openEditRecordModal(${rec.id})" title="全項目修正"><i class="fa-solid fa-pen-to-square"></i></button>
            <button class="bg-rose-950/40 hover:bg-rose-900/60 text-rose-400 p-1.5 rounded active:scale-90" onclick="openDeleteModal(${rec.id})" title="削除"><i class="fa-solid fa-trash-can"></i></button>
          </div>
        </td>
      </tr>`;
  }).join('');
  updateMasterCheckboxState();
}

/* ---- 編集モーダル ---- */
function openEditRecordModal(id) {
  const rec = historyLogs.find(l => l.id === id);
  if (!rec) return;
  document.getElementById('edit-record-id').value = rec.id;
  document.getElementById('edit-date').value = rec.date || '';
  document.getElementById('edit-operator').value = rec.operator || '';
  document.getElementById('edit-partno').value = rec.partNo || '';
  document.getElementById('edit-count').value = rec.count || 0;
  document.getElementById('edit-material-count').value = rec.materialCount || 0;
  document.getElementById('edit-scrap-count').value = rec.scrapCount || 0;
  document.getElementById('edit-stop-reason').value = rec.stopReason || '';
  document.getElementById('edit-spm').value = rec.settingSPM || 15.0;
  document.getElementById('edit-start').value = rec.start || '';
  document.getElementById('edit-end').value = rec.end || '';
  document.getElementById('edit-running-min').value = Math.round(rec.runningSec / 60);
  document.getElementById('edit-mold-min').value = Math.round(rec.moldSec / 60);
  document.getElementById('edit-break-min').value = Math.round(rec.breakSec / 60);
  document.getElementById('edit-stop-min').value = Math.round(rec.stopSec / 60);
  showModal('edit-record-modal');
}
function closeEditRecordModal() { hideModal('edit-record-modal'); }

function saveEditedRecord() {
  const id = parseInt(document.getElementById('edit-record-id').value);
  const rec = historyLogs.find(l => l.id === id);
  if (!rec) return;

  rec.date = document.getElementById('edit-date').value.trim() || rec.date;
  rec.operator = document.getElementById('edit-operator').value.trim() || '未登録';
  rec.partNo = document.getElementById('edit-partno').value.trim() || '未設定';
  rec.count = parseInt(document.getElementById('edit-count').value) || 0;
  rec.materialCount = parseInt(document.getElementById('edit-material-count').value) || 0;
  rec.scrapCount = parseInt(document.getElementById('edit-scrap-count').value) || 0;
  rec.stopReason = document.getElementById('edit-stop-reason').value.trim() || 'なし';
  rec.settingSPM = (parseFloat(document.getElementById('edit-spm').value) || 0.1).toFixed(1);
  rec.start = document.getElementById('edit-start').value.trim() || rec.start;
  rec.end = document.getElementById('edit-end').value.trim() || rec.end;
  rec.runningSec = (parseInt(document.getElementById('edit-running-min').value) || 0) * 60;
  rec.moldSec = (parseInt(document.getElementById('edit-mold-min').value) || 0) * 60;
  rec.breakSec = (parseInt(document.getElementById('edit-break-min').value) || 0) * 60;
  rec.stopSec = (parseInt(document.getElementById('edit-stop-min').value) || 0) * 60;

  const totalActiveSec = rec.runningSec + rec.moldSec + rec.stopSec;
  if (totalActiveSec > 0 && parseFloat(rec.settingSPM) > 0) {
    const stdCT = 60 / parseFloat(rec.settingSPM);
    rec.perfRate = Math.min((stdCT * rec.count) / totalActiveSec * 100, 200.0).toFixed(1);
  } else {
    rec.perfRate = '100.0';
  }

  rec.syncStatus = 'pending'; // 修正後は再同期が必要
  persistHistory();
  renderHistoryTable();
  updateSyncBadge();
  closeEditRecordModal();
  showToast('日報データを修正しました。クラウドへ再同期します');
  syncRecordToCloud(rec);
}

/* ---- 削除 ---- */
function openDeleteModal(id) { document.getElementById('delete-record-id').value = id; showModal('delete-confirm-modal'); }
function closeDeleteModal() { hideModal('delete-confirm-modal'); }
function confirmDeleteHistory() {
  const id = parseInt(document.getElementById('delete-record-id').value);
  historyLogs = historyLogs.filter(l => l.id !== id);
  checkedRecordIds.delete(id);
  persistHistory();
  renderHistoryTable();
  updateSyncBadge();
  closeDeleteModal();
  showToast('実績データを削除しました（クラウド側は残ります）');
}

/* =========================================================
   Excel生成（ローカルバックアップ用）
   ========================================================= */
function generateExcelWorkbook(selectedLogs, currentEquip) {
  const sheetData = [
    [`【${currentEquip}】 プレス生産作業日報`],
    [
      '出力日時: ' + new Date().toLocaleString('ja-JP'),
      '選択件数: ' + selectedLogs.length + '件',
      '総生産数: ' + selectedLogs.reduce((s, r) => s + r.count, 0) + ' pcs',
      '平均性能稼働率: ' + (selectedLogs.reduce((s, r) => s + parseFloat(r.perfRate), 0) / (selectedLogs.length || 1)).toFixed(1) + '%',
    ],
    ['日付', '設備名', '作業員', '作業人数', '部品番号', '製造開始', '製造終了', '生産数(pcs)', '材料交換(回)', 'スクラップ交換(回)', '停止理由', '実生産稼働時間', '金型交換時間', '計画停止時間', '異常停止時間', '性能稼働率(%)'],
  ];
  selectedLogs.forEach(rec => {
    sheetData.push([
      rec.date, rec.equipment || currentEquip, rec.operator || '未選択', rec.workerCount || 1, rec.partNo,
      rec.start, rec.end, rec.count, rec.materialCount || 0, rec.scrapCount || 0, rec.stopReason || 'なし',
      formatTime(rec.runningSec), formatTime(rec.moldSec), formatTime(rec.breakSec), formatTime(rec.stopSec), parseFloat(rec.perfRate),
    ]);
  });
  const ws = XLSX.utils.aoa_to_sheet(sheetData);
  ws['!cols'] = [12,16,18,10,16,12,12,12,12,14,20,14,14,14,14,14].map(w => ({ wch: w }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '作業日報');
  return wb;
}

function exportExcelLocalBackup() {
  const selected = historyLogs.filter(l => checkedRecordIds.has(l.id));
  const target = selected.length > 0 ? selected : historyLogs;
  if (target.length === 0) { showToast('出力対象の日報データがありません', true); return; }
  const wb = generateExcelWorkbook(target, equipmentSelect.value);
  XLSX.writeFile(wb, `プレス作業日報_${new Date().toISOString().slice(0, 10)}.xlsx`);
  showToast(`ローカルにExcel保存しました（${target.length}件）`);
}

/* =========================================================
   共通UIヘルパー
   ========================================================= */
function switchView(name) {
  const dash = document.getElementById('view-dashboard');
  const hist = document.getElementById('view-history');
  const tabDash = document.getElementById('tab-dashboard');
  const tabHist = document.getElementById('tab-history');
  if (name === 'dashboard') {
    dash.classList.remove('hidden'); dash.classList.add('flex');
    hist.classList.add('hidden'); hist.classList.remove('flex');
    tabDash.className = 'px-3 py-1.5 text-[11px] font-bold rounded-lg transition-all text-white bg-panel2';
    tabHist.className = 'px-3 py-1.5 text-[11px] font-bold rounded-lg transition-all text-slate-400';
  } else {
    dash.classList.add('hidden'); dash.classList.remove('flex');
    hist.classList.remove('hidden'); hist.classList.add('flex');
    tabDash.className = 'px-3 py-1.5 text-[11px] font-bold rounded-lg transition-all text-slate-400';
    tabHist.className = 'px-3 py-1.5 text-[11px] font-bold rounded-lg transition-all text-white bg-panel2';
    renderHistoryTable();
    updateCloudStatusChip();
  }
}

function showModal(id) { document.getElementById(id).classList.remove('opacity-0', 'pointer-events-none'); }
function hideModal(id) { document.getElementById(id).classList.add('opacity-0', 'pointer-events-none'); }

function showToast(message, isError) {
  const toast = document.getElementById('toast');
  const icon = document.getElementById('toast-icon');
  document.getElementById('toast-text').innerText = message;
  toast.classList.remove('bg-emerald-600', 'border-emerald-400', 'bg-rose-600', 'border-rose-400');
  if (isError) { toast.classList.add('bg-rose-600', 'border-rose-400'); icon.className = 'fa-solid fa-circle-exclamation'; }
  else { toast.classList.add('bg-emerald-600', 'border-emerald-400'); icon.className = 'fa-solid fa-circle-check'; }
  toast.classList.remove('opacity-0', 'pointer-events-none');
  toast.classList.add('opacity-100');
  clearTimeout(window.__toastTimer);
  window.__toastTimer = setTimeout(() => {
    toast.classList.remove('opacity-100');
    toast.classList.add('opacity-0', 'pointer-events-none');
  }, 3200);
}

function formatTime(sec) {
  const h = String(Math.floor(sec / 3600)).padStart(2, '0');
  const m = String(Math.floor((sec % 3600) / 60)).padStart(2, '0');
  const s = String(Math.floor(sec % 60)).padStart(2, '0');
  return `${h}:${m}:${s}`;
}

function escapeHtml(str) {
  return String(str ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}
