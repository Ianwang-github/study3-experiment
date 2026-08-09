/**
 * Study 3 v4（前置執行深度 × 方案符合度）資料收集後端
 *
 * 建置：新開一個 Google Sheet → 擴充功能 → Apps Script（務必從試算表內部建立，
 *       不可開獨立專案，否則 getActiveSpreadsheet() 會回傳 null）→ 貼上本檔覆蓋 Code.gs
 *       → SHARED_SECRET 改成與 gas-config.js 的 window.GAS_SECRET 完全一致
 *       → 部署 → 新增部署作業 → 網頁應用程式（執行身分：我／存取權：所有人）
 *       → 複製 /exec 網址貼進 gas-config.js
 *
 * ⚠️ 更新既有部署：部署 → 管理部署作業 → 鉛筆 → 版本選「新版本」→ 部署。
 *    只按 Ctrl+S 不會讓正式網址生效。
 * ⚠️ 改過欄位後，收案前務必把資料分頁整個清空（含表頭）讓程式重建。
 *
 * 資料結構：每位受試者 4 列（1 練習 + 3 正式）。練習列 is_practice=1，分析時排除。
 * 真值（哪幾條規格被違反）不存在本檔也不存在前端——judge_1..4 原樣記錄，
 * d′ 與 c 於離線分析時對照獨立評分者建立的真值表計算。
 */

const SHARED_SECRET = "smallbee"; // 已改成配合您 gas-config.js 裡設定的密鑰

const CONDS = ["A1", "A2", "A3"];          // 受試者間三組
const QUOTA_PER_CELL = 30;                  // N = 90
const ASSIGN_SHEET = "分派記錄";
const ASSIGN_HEADERS = ["assigned_at", "pid", "cond", "completed_at", "quota_status"];
const CELL_LIMIT = 45000;

const HEADERS = [
  // 識別
  "submitted_at", "pid", "A", "trial", "is_practice",
  "theme", "compliance", "exemplar", "stim_id", "stim_missing",
  // 設計與環境
  "latin", "trial_order", "think_sec", "version", "session_ts", "vw", "vh", "dpr",
  // 裝置與輸入方式：手指／觸控筆／滑鼠會影響手繪的費力程度與精度，
  // 須記錄以檢查效果是否跨輸入方式一致
  "max_touch", "coarse_pointer", "input_types", "input_primary",
  // 表徵形成度檢核（對照物是否真的形成）
  "ref_clear",
  // 背景共變項（每位受試者一次，四列重複）
  "age", "design_bg", "dse_1", "dse_2", "dse_3", "ai_use", "ai_att_1", "ai_att_2",
  // 規格逐條判定（→ 離線對真值計算 d′ 與 c）
  "judge_1", "judge_2", "judge_3", "judge_4",
  // 品質評價（CPSS：Resolution / Elaboration / Novelty）
  "qual_1", "qual_2", "qual_3",
  // 採納、修改必要性、擁有感
  "adopt_1", "need_1", "own_1",
  // 信任、情緒、負荷
  "trust_1", "trust_2", "trust_3", "trust_4", "tlx_1", "tlx_2", // SAM 已依前測疲勞回饋移除
  // 改進方向發想（固著代價指標；每項獨立欄位，編碼單位由受試者自行切分）
  "rev_json", "rev_count", "rev_chars", "rev_time_used", "rev_timeout",
  // 前置執行的產出證據
  "imp_1", "imp_2", "imp_3", "imp_4",           // A1
  "spec_text", "spec_len",                       // A2
  "stroke_count", "point_count", "ink", "bbox_w", "bbox_h", "clears", "undos", "strokes_json", // A3
  // 時間歷程（毫秒，相對該試驗起點）
  "t_brief", "t_first_stroke", "t_produce_done", "t_think_done",
  "t_stim_shown", "t_eval_start", "t_scales_done", "t_done",
  // 事後（每位受試者一次）
  "post_ac", "post_strategy", "post_guess",
  "trial_ts"
];

function getDataSheet(ss) {
  const props = PropertiesService.getScriptProperties();
  const saved = props.getProperty("DATA_SHEET_NAME");
  if (saved) { const sh = ss.getSheetByName(saved); if (sh) return sh; }
  const active = ss.getActiveSheet();
  props.setProperty("DATA_SHEET_NAME", active.getName());
  return active;
}

function getAssignSheet(ss) {
  let sh = ss.getSheetByName(ASSIGN_SHEET);
  if (!sh) { sh = ss.insertSheet(ASSIGN_SHEET); sh.appendRow(ASSIGN_HEADERS); }
  else if (sh.getLastRow() === 0) sh.appendRow(ASSIGN_HEADERS);
  return sh;
}

/** appendRow() 會靜默跳過陣列中的 undefined，導致其後欄位整列左移錯位。
 *  A1/A2/A3 三組各自有大量空白欄位，幾乎每列都會中招，故一律先轉字串。 */
function clip(v) {
  if (v === undefined || v === null) return "";
  const s = String(v);
  return s.length > CELL_LIMIT ? s.slice(0, CELL_LIMIT) + "…[TRUNCATED]" : s;
}

function doPost(e) {
  let payload;
  try { payload = JSON.parse(e.postData.contents); }
  catch (err) { return ContentService.createTextOutput("OK"); }

  const meta = payload.meta || {};
  if (SHARED_SECRET && meta.secret !== SHARED_SECRET) {
    return ContentService.createTextOutput("OK"); // 靜默拒絕，不洩漏驗證邏輯
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getDataSheet(ss);
  if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);

  const now = new Date().toISOString();
  const trials = payload.trials || [];

  trials.forEach(function (t) {
    const a = t.answers || {};
    const i = t.interaction || {};
    const imp = i.importance || {};

    const row = [
      now, meta.pid, meta.A, t.trial, t.practice ? 1 : 0,
      t.theme, t.compliance, t.exemplar, i.stim, i.stim_missing,
      meta.latin, meta.order, meta.think_sec, meta.version, meta.ts, meta.vw, meta.vh, meta.dpr,
      meta.max_touch, meta.coarse, i.input_types, i.input_primary,
      i.ref_clear,
      meta.age, meta.design_bg, meta.dse_1, meta.dse_2, meta.dse_3, meta.ai_use, meta.ai_att_1, meta.ai_att_2,
      a.judge_1, a.judge_2, a.judge_3, a.judge_4,
      a.qual_1, a.qual_2, a.qual_3,
      a.adopt_1, a.need_1, a.own_1,
      a.trust_1, a.trust_2, a.trust_3, a.trust_4, a.tlx_1, a.tlx_2,
      i.rev_ideas ? JSON.stringify(i.rev_ideas) : "", i.rev_count, i.rev_chars, i.rev_time_used, i.rev_timeout,
      imp.imp_1, imp.imp_2, imp.imp_3, imp.imp_4,
      i.spec_text, i.spec_len,
      i.stroke_count, i.point_count, i.ink, i.bbox_w, i.bbox_h, i.clears, i.undos,
      i.strokes ? JSON.stringify(i.strokes) : "",
      i.t_brief, i.t_first_stroke, i.t_produce_done, i.t_think_done,
      i.t_stim_shown, i.t_eval_start, i.t_scales_done, i.t_done,
      meta.ac, meta.strategy, meta.guess,
      t.t
    ].map(clip);

    sheet.appendRow(row);
  });

  if (meta.pid) { try { markDone(ss, meta.pid); } catch (err) {} }
  return ContentService.createTextOutput("OK");
}

function markDone(ss, pid) {
  const sh = ss.getSheetByName(ASSIGN_SHEET);
  if (!sh) return;
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (v[i][1] === pid && !v[i][3]) { sh.getRange(i + 1, 4).setValue(new Date().toISOString()); break; }
  }
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === "assign") return handleAssign(p);
  return ContentService.createTextOutput("OK - Study3 v4 backend online");
}

/**
 * 配額分派：永遠分給目前人數最少的組，並列最少時隨機挑一。
 * 性質：任何時點三組人數差 ≤ 1；每連續 3 次分派三組各得 1。
 * 等同持續進行的區塊隨機化，且無總量上限——收滿 N=90 後仍可續收，
 * 超額者標記 overflow，供後續彈性決定主分析樣本範圍。
 */
function handleAssign(p) {
  const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
  if (SHARED_SECRET && p.secret !== SHARED_SECRET) return out({ ok: false, error: "forbidden" });

  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (err) { return out({ ok: false, error: "busy" }); }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = getAssignSheet(ss);
    const v = sh.getDataRange().getValues();

    const counts = {};
    CONDS.forEach(c => counts[c] = 0);
    for (let i = 1; i < v.length; i++) if (counts.hasOwnProperty(v[i][2])) counts[v[i][2]]++;

    const min = Math.min.apply(null, CONDS.map(c => counts[c]));
    const pool = CONDS.filter(c => counts[c] === min);
    const cond = pool[Math.floor(Math.random() * pool.length)];

    const seq = v.length; // 含表頭，天然從 1 起跳且不重複
    const pid = "P" + String(seq).padStart(3, "0");
    const status = seq <= QUOTA_PER_CELL * CONDS.length ? "quota" : "overflow";

    sh.appendRow([new Date().toISOString(), pid, cond, "", status]);
    return out({ ok: true, pid: pid, cond: cond });
  } finally { lock.releaseLock(); }
}

/** 維護用：釋放已分派逾 40 分鐘仍未完成的名額（點了連結但中途放棄者） */
function releaseStale() {
  const MIN = 40;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(ASSIGN_SHEET);
  if (!sh) return;
  const v = sh.getDataRange().getValues();
  const cutoff = Date.now() - MIN * 60 * 1000;
  let n = 0;
  for (let i = v.length - 1; i >= 1; i--) {
    if (!v[i][3] && new Date(v[i][0]).getTime() < cutoff) { sh.deleteRow(i + 1); n++; }
  }
  Logger.log("釋放 " + n + " 筆逾時分派");
}

/** 進度檢視：各組已完成人數 */
function progress() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(ASSIGN_SHEET);
  if (!sh) { Logger.log("尚無分派記錄"); return; }
  const v = sh.getDataRange().getValues();
  const done = {}, all = {};
  CONDS.forEach(c => { done[c] = 0; all[c] = 0; });
  for (let i = 1; i < v.length; i++) {
    if (!all.hasOwnProperty(v[i][2])) continue;
    all[v[i][2]]++; if (v[i][3]) done[v[i][2]]++;
  }
  CONDS.forEach(c => Logger.log(c + "：已完成 " + done[c] + " / 已分派 " + all[c] + "（目標 " + QUOTA_PER_CELL + "）"));
}

/** 測試用：確認欄位對齊與寫入無誤 */
function testDoPost() {
  const fake = {
    meta: { pid: "P001", A: "A3", latin: 1, order: "chair-mug-bag", version: "s3-v4.0-fixation",
            think_sec: 90, vw: 390, vh: 844, dpr: 3, ts: new Date().toISOString(),
            age: 2, design_bg: 3, dse_1: 5, dse_2: 4, dse_3: 6, ai_use: 3, ai_att_1: 4, ai_att_2: 5,
            ac: 2, strategy: 3, guess: "想看先畫過會不會影響評價", secret: SHARED_SECRET },
    trials: [{
      trial: 1, practice: false, theme: "chair", compliance: "B3", exemplar: "e3", A: "A3",
      answers: { judge_1: 1, judge_2: 0, judge_3: 0, judge_4: 1,
                 qual_1: 3, qual_2: 4, qual_3: 5, adopt_1: 3, need_1: 6, own_1: 2,
                 trust_1: 4, trust_2: 3, trust_3: 4, trust_4: 3, tlx_1: 5, tlx_2: 4,
                 },
      interaction: { stim: "chair_B3_e3", input_types: "touch", input_primary: "touch", ref_clear: 5,
                     rev_ideas: ["椅背橫桿改成三根", "椅腳角度再往外撇", "座面前緣加大圓角"],
                     rev_count: 3, rev_chars: 24, rev_time_used: 78000, rev_timeout: 0,
                     stroke_count: 7, point_count: 412, ink: 1830,
                     bbox_w: 168, bbox_h: 204, clears: 1, undos: 2, strokes: [[[10,10],[20,22]]],
                     t_brief: 1200, t_first_stroke: 8400, t_produce_done: 62000,
                     t_stim_shown: 62500, t_eval_start: 71000, t_scales_done: 132000, t_done: 210000 },
      t: new Date().toISOString()
    }]
  };
  doPost({ postData: { contents: JSON.stringify(fake) } });
}
