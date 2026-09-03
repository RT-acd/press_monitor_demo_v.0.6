/* ==========================================================
   PRESS MONITOR — 設定値・定数
   このファイルは他のどのJSファイルよりも先に読み込まれます。
   ========================================================== */

const STORAGE = {
  HISTORY: 'pm_history_v1',
  SETTINGS: 'pm_settings_v1',
  MASTERS: 'pm_masters_v2', // 品番／作業員マスタのローカルキャッシュ（クラウドから取得）
};

const CONFIG = {
  INITIAL_TARGET_COUNT: 1200,
  PERFORMANCE_GOOD_LIMIT: 95,
  PERFORMANCE_NORMAL_LIMIT: 85,
  SYNC_RETRY_INTERVAL_MS: 30000,
  MASTER_AUTO_REFRESH_MS: 10 * 60 * 1000, // 10分ごとにマスタをバックグラウンド再取得
};

// アプリ固有の設定（クラウド接続情報など）。gasUrl未設定時はクラウド機能は使えません。
const DEFAULT_SETTINGS = {
  gasUrl: '',
  defaultEmail: 'factory-report@example.com',
  // 設備名リストは端末側で管理（品番・作業員のようにクラウド化していません）
  equipmentList: ['1軸100tプレス', '2軸200tプレス', '高速300tプレス', 'Aライン プレス機', 'Bライン プレス機'],
};

// 品番・作業員マスタの初期値（クラウド未接続時のフォールバック用）
const DEFAULT_MASTERS = {
  partMasters: [
    { part: 'SP-200', spm: 15.0 },
    { part: 'SP-210', spm: 18.0 },
  ],
  operators: ['山田', '佐藤', '鈴木', '田中'],
  lastSyncedAt: null,
};
