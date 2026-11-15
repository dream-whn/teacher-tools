// 打字小達人 後端（Google Apps Script）
// 請從你的 Google 試算表開啟：上方選單「擴充功能」→「Apps Script」，再貼上本程式碼。
const PW = "老師自己設密碼";
const VERSION = "1.4";
const MASCOTS = ["🐣","🐰","🐱","🐶","🐼","🦊","🦁","🐯","🦄","🐲","🌟","🚀"];
const THEMES = ["pink","blue","green","purple","orange"];
const NAME_RE = /^[1-6][A-Z] \S.{0,19}$/;   // 例如「5A 陳大文」

function out(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function doGet(){ return out({ok:true}); }

// 取得資料用的試算表：優先用綁定的試算表；若程式是獨立專案，就自動在你的 Google Drive 建立一份並記住
function book_(){
  try { const a = SpreadsheetApp.getActiveSpreadsheet(); if (a) return a; } catch(e){}
  const P = PropertiesService.getScriptProperties();
  const id = P.getProperty("sheetId");
  if (id) { try { return SpreadsheetApp.openById(id); } catch(e){} }
  try {
    const b = SpreadsheetApp.create("打字小達人 資料");
    P.setProperty("sheetId", b.getId());
    return b;
  } catch(e){ return null; }
}
// 在 Apps Script 編輯器選這個函式按「執行」，再看「執行記錄」，就會顯示試算表的網址
function showSheetUrl(){ const b = book_(); Logger.log(b ? b.getUrl() : "找不到試算表，請確認已授權"); }
function sheet_(name, head, sample){
  const b = book_(); if (!b) return null;
  let sh = b.getSheetByName(name);
  if (!sh) {
    sh = b.insertSheet(name);
    sh.appendRow(head); sh.setFrozenRows(1);
    (sample || []).forEach(r => sh.appendRow(r));
  }
  return sh;
}

// 讀取試算表「題目」分頁（快取 60 秒，避免多人同時使用時超出額度）
function passages_(){
  const C = CacheService.getScriptCache(), hit = C.get("passages");
  if (hit) { try { return JSON.parse(hit); } catch(e){} }
  const sh = sheet_("題目", ["級別（1-10）","文字內容","模式（取代／加入）","啟用（Y／N）"],
                    [[1, "（範例）把這一列改成你自己的句子。", "加入", "N"]]);
  if (!sh) return null;
  const v = sh.getDataRange().getValues(), res = {};
  for (let i = 1; i < v.length; i++) {
    const lv = parseInt(v[i][0], 10);
    const t = String(v[i][1] || "").replace(/[\r\n]+/g, "").trim();
    if (!(lv >= 1 && lv <= 10) || !t || t.length > 300) continue;
    if (String(v[i][3]).trim().toUpperCase() === "N") continue;
    (res[lv] = res[lv] || []).push({t: t, r: String(v[i][2]).trim() === "取代"});
  }
  try { C.put("passages", JSON.stringify(res), 60); } catch(e){}
  return res;
}

function doPost(e){
  const d = JSON.parse(e.postData.contents);
  const P = PropertiesService.getScriptProperties();
  const teacher = d.pw === PW;

  if (d.action === "levels") {
    const v = P.getProperty("levels");
    const st = P.getProperty("settings");
    return out({levels: v ? JSON.parse(v) : null, passages: passages_(), settings: st ? JSON.parse(st) : null});
  }
  if (d.action === "saveLevels") {
    if (!teacher) return out({error:"pw"});
    P.setProperty("levels", JSON.stringify(d.levels));
    return out({ok:1});
  }
  if (d.action === "saveSettings") {              // 老師個人化：標題、吉祥物、主題色
    if (!teacher) return out({error:"pw"});
    const title = Array.from(String(d.title || "").trim()).slice(0, 20).join("");
    P.setProperty("settings", JSON.stringify({
      title: title,
      mascot: MASCOTS.indexOf(d.mascot) >= 0 ? d.mascot : "🐣",
      theme: THEMES.indexOf(d.theme) >= 0 ? d.theme : "pink"
    }));
    return out({ok:1});
  }
  if (d.action === "check") return out({ok: teacher, version: VERSION});
  if (d.action === "getp") {
    const v = P.getProperty("p:" + String(d.name).slice(0,30));
    return out({p: v ? JSON.parse(v) : null});
  }
  if (d.action === "setp") {
    const n = String(d.name);
    if (!NAME_RE.test(n)) return out({error:"name"});
    const old = P.getProperty("p:" + n);
    const o = old ? JSON.parse(old) : {unlocked:1};
    P.setProperty("p:" + n, JSON.stringify({unlocked: Math.min(11, Math.max(o.unlocked, +d.unlocked || 1)), t: Date.now()}));
    return out({ok:1});
  }
  if (d.action === "log") {                       // 每次練習記錄到「學生紀錄」分頁
    const n = String(d.name);
    if (!NAME_RE.test(n)) return out({error:"name"});
    const sh = sheet_("學生紀錄", ["時間","班別","姓名","級別","用時（秒）","速度（字/分）","正確率（%）","過關"]);
    if (!sh) return out({ok:0});
    const L = LockService.getScriptLock(); L.waitLock(10000);
    try {
      const sp = n.indexOf(" ");
      const c = (x, m) => Math.max(0, Math.min(m, +x || 0));
      const now = new Date(), cls = n.slice(0, sp), lv = c(d.level,10);
      // 班別、姓名欄先設成純文字，避免學生輸入以 = 開頭的名字被當成試算表公式
      const r1 = sh.getLastRow() + 1;
      sh.getRange(r1, 2, 1, 2).setNumberFormat("@");
      sh.getRange(r1, 1, 1, 8).setValues([[now, cls, n.slice(sp + 1), lv, c(d.sec,36000), c(d.cpm,2000), c(d.acc,100), d.pass ? "是" : "否"]]);
      // 錯字明細：[正確字, 學生打成的字（空白＝漏打）]，每次最多 60 筆
      const rows = [];
      (Array.isArray(d.errs) ? d.errs.slice(0, 60) : []).forEach(x => {
        if (Array.isArray(x) && typeof x[0] === "string" && typeof x[1] === "string" &&
            x[0].length >= 1 && x[0].length <= 2 && x[1].length <= 2) rows.push([now, cls, lv, x[0], x[1]]);
      });
      if (rows.length) {
        const es = sheet_("錯字紀錄", ["時間","班別","級別","正確字","打成（空白＝漏打）"]);
        if (es) {
          const r0 = es.getLastRow() + 1;
          es.getRange(r0, 4, rows.length, 2).setNumberFormat("@");
          es.getRange(r0, 1, rows.length, 5).setValues(rows);
        }
      }
    } finally { L.releaseLock(); }
    return out({ok:1});
  }
  if (d.action === "top") {                       // 光榮榜：速度前三名、級別前三名（快取 60 秒）
    const C = CacheService.getScriptCache(), hit = C.get("top");
    if (hit) { try { return out(JSON.parse(hit)); } catch(e){} }
    const best = {};
    const sh = book_() && book_().getSheetByName("學生紀錄");
    if (sh && sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues().forEach(r => {
        if (r[7] !== "是") return;                  // 只計算過關的練習
        const k = r[1] + " " + r[2], v = +r[5] || 0;
        if (!best[k] || v > best[k]) best[k] = v;
      });
    }
    const speed = Object.keys(best).map(k => ({n: k, v: best[k]})).sort((a, b) => b.v - a.v).slice(0, 3);
    const all = P.getProperties(), lv = [];
    for (const k in all) if (k.indexOf("p:") === 0) { const o = JSON.parse(all[k]); lv.push({n: k.slice(2), v: Math.min(10, o.unlocked - 1), t: o.t || 0}); }
    const level = lv.filter(x => x.v > 0).sort((a, b) => b.v - a.v || a.t - b.t).slice(0, 3).map(x => ({n: x.n, v: x.v}));
    const res = {speed: speed, level: level};
    try { C.put("top", JSON.stringify(res), 60); } catch(e){}
    return out(res);
  }
  if (d.action === "errStats") {                  // 老師：錯字統計（正確字、打成的字、次數）
    if (!teacher) return out({error:"pw"});
    const es = book_() && book_().getSheetByName("錯字紀錄");
    const cnt = {}; let total = 0;
    if (es && es.getLastRow() > 1) {
      es.getRange(2, 1, es.getLastRow() - 1, 5).getValues().forEach(r => {
        if (d.cls && r[1] !== d.cls) return;
        const k = String(r[3]) + "\u0001" + String(r[4]);
        cnt[k] = (cnt[k] || 0) + 1; total++;
      });
    }
    const pairs = Object.keys(cnt).map(k => { const q = k.split("\u0001"); return {e: q[0], t: q[1] || "", n: cnt[k]}; })
      .sort((a, b) => b.n - a.n).slice(0, 1500);
    return out({pairs: pairs, total: total});
  }
  if (d.action === "list") {
    if (!teacher) return out({error:"pw"});
    const a = P.getProperties(), r = [];
    for (const k in a) if (k.indexOf("p:") === 0) r.push(Object.assign({name: k.slice(2)}, JSON.parse(a[k])));
    return out({list: r});
  }
  if (d.action === "del") {
    if (!teacher) return out({error:"pw"});
    P.deleteProperty("p:" + d.name);
    return out({ok:1});
  }
  return out({error:"unknown"});
}
