// 打字小達人 後端（Google Apps Script）
// 請從你的 Google 試算表開啟：上方選單「擴充功能」→「Apps Script」，再貼上本程式碼。
const PW = "請改成你的老師密碼";
const VERSION = "1.1";
const NAME_RE = /^[1-6][A-Z] \S.{0,19}$/;   // 例如「5A 陳大文」

function out(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function doGet(){ return out({ok:true}); }

function book_(){ try { return SpreadsheetApp.getActiveSpreadsheet(); } catch(e){ return null; } }
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
    return out({levels: v ? JSON.parse(v) : null, passages: passages_()});
  }
  if (d.action === "saveLevels") {
    if (!teacher) return out({error:"pw"});
    P.setProperty("levels", JSON.stringify(d.levels));
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
      sh.appendRow([new Date(), n.slice(0, sp), n.slice(sp + 1), c(d.level,10), c(d.sec,36000), c(d.cpm,2000), c(d.acc,100), d.pass ? "是" : "否"]);
    } finally { L.releaseLock(); }
    return out({ok:1});
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
