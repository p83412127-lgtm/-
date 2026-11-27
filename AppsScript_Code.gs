/***** 神經系統練習／小考 後端（綁在「神經系統複習測驗_回覆統計」試算表）*****/
/* 分頁：題庫 / 班級名單 / 設定 / 練習作答_神經系統 / 作答狀態
   舊的 warm-up.html、gcs-practice.html 仍走下方的 legacy doPost / doGet（無 action 參數）。 */

var SH_BANK = "題庫";
var SH_ROSTER = "班級名單";
var SH_CONF = "設定";
var SH_ANS = "練習作答_神經系統";
var SH_STATE = "作答狀態";

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function now_() {
  return Utilities.formatDate(new Date(), ss_().getSpreadsheetTimeZone(), "yyyy/MM/dd HH:mm:ss");
}
function s_(x) { return (x === null || x === undefined) ? "" : String(x).trim(); }
function fmt_(x) {
  return (x instanceof Date)
    ? Utilities.formatDate(x, ss_().getSpreadsheetTimeZone(), "yyyy/MM/dd HH:mm:ss")
    : s_(x);
}

function conf_() {
  var sh = ss_().getSheetByName(SH_CONF), o = {};
  if (!sh || sh.getLastRow() < 2) return o;
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) { var k = s_(v[i][0]); if (k) o[k] = v[i][1]; }
  return o;
}

function mode_() { return s_(conf_()["模式"]) === "考試" ? "考試" : "練習"; }

/* 「適用班級」留空＝整份名單都能作答；填了就只比對該班（班級欄要一致）。 */
function roster_() {
  var sh = ss_().getSheetByName(SH_ROSTER), out = [];
  if (!sh || sh.getLastRow() < 2) return out;
  var only = s_(conf_()["適用班級"]);
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    var nm = s_(v[i][1]);
    if (!nm) continue;
    var cls = s_(v[i][2]);
    if (only && cls !== only) continue;
    out.push({ id: s_(v[i][0]), name: nm, cls: cls });
  }
  return out;
}

function bank_() {
  var sh = ss_().getSheetByName(SH_BANK), out = [];
  if (!sh || sh.getLastRow() < 2) return out;
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    var no = s_(v[i][0]);
    if (!no) continue;
    if (s_(v[i][6]) !== "啟用") continue;
    var opts = s_(v[i][2]).split(/[|｜]/).map(s_).filter(function (x) { return x !== ""; });
    var keys = s_(v[i][3]).toUpperCase().split(/[,，、]/).map(s_).filter(function (x) { return x !== ""; });
    var ans = [];
    for (var k = 0; k < keys.length; k++) {
      var idx = keys[k].charCodeAt(0) - 65;
      if (idx >= 0 && idx < opts.length) ans.push(opts[idx]);
    }
    if (!opts.length || !ans.length) continue;
    out.push({
      no: no, q: s_(v[i][1]), opts: opts, ans: ans, multi: ans.length > 1,
      why: s_(v[i][4]), cat: s_(v[i][5]), lvl: s_(v[i][7])
    });
  }
  return out;
}

function seedRand_(seed) {
  var h = 2166136261;
  for (var i = 0; i < seed.length; i++) { h = h ^ seed.charCodeAt(i); h = Math.imul(h, 16777619); }
  var s = (h >>> 0) || 1;
  return function () {
    s = s ^ (s << 13); s = s >>> 0;
    s = s ^ (s >>> 17);
    s = s ^ (s << 5); s = s >>> 0;
    return s / 4294967296;
  };
}
function shuffle_(arr, rnd) {
  var a = arr.slice();
  for (var i = a.length - 1; i > 0; i--) {
    var j = Math.floor(rnd() * (i + 1)), t = a[i]; a[i] = a[j]; a[j] = t;
  }
  return a;
}

function parseTime_(x) {
  if (x instanceof Date) return x;
  var t = s_(x);
  if (!t) return null;
  var d = new Date(t.replace(/-/g, "/"));
  return isNaN(d.getTime()) ? null : d;
}

function windowState_() {
  var c = conf_(), a = parseTime_(c["開放開始時間"]), b = parseTime_(c["開放結束時間"]), n = new Date();
  var tz = ss_().getSpreadsheetTimeZone();
  if (a && n < a) return { ok: false, why: "尚未開放作答（開放時間 " + Utilities.formatDate(a, tz, "yyyy/MM/dd HH:mm") + " 起）" };
  if (b && n > b) return { ok: false, why: "作答時間已結束（" + Utilities.formatDate(b, tz, "yyyy/MM/dd HH:mm") + " 截止）" };
  return { ok: true };
}

function matchRoster_(id, name) {
  var r = roster_();
  if (!r.length) return { ok: false, why: "「班級名單」分頁是空的，請先貼上名單" };
  for (var i = 0; i < r.length; i++) {
    if (r[i].name !== name) continue;
    if (r[i].id === "") return { ok: true, cls: r[i].cls, loose: true };
    if (r[i].id === id) return { ok: true, cls: r[i].cls, loose: false };
    return { ok: false, why: "學號與姓名對不上名單，請確認" };
  }
  return { ok: false, why: "查無此姓名，不在班級名單內" };
}

function answers_() {
  var sh = ss_().getSheetByName(SH_ANS), byId = {}, byName = {}, rows = [];
  if (!sh || sh.getLastRow() < 2) return { byId: byId, byName: byName, rows: rows };
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    var id = s_(v[i][1]), nm = s_(v[i][2]);
    if (!id && !nm) continue;
    var r = { time: fmt_(v[i][0]), id: id, name: nm, right: v[i][4], total: v[i][5], detail: s_(v[i][7]) };
    rows.push(r);
    if (id) byId[id] = r;
    if (nm) byName[nm] = r;
  }
  return { byId: byId, byName: byName, rows: rows };
}

/***** 作答狀態（考試模式用）*****/
function stateSheet_() {
  var ss = ss_(), sh = ss.getSheetByName(SH_STATE);
  if (!sh) {
    sh = ss.insertSheet(SH_STATE);
    sh.appendRow(["學號", "姓名", "狀態", "開始時間", "最後更新", "原因", "解鎖次數"]);
    sh.setFrozenRows(1);
  }
  return sh;
}
function states_() {
  var sh = stateSheet_(), map = {}, rows = [];
  if (sh.getLastRow() < 2) return { map: map, rows: rows, sh: sh };
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    var id = s_(v[i][0]);
    if (!id) continue;
    var r = {
      row: i + 1, id: id, name: s_(v[i][1]), st: s_(v[i][2]),
      start: fmt_(v[i][3]), upd: fmt_(v[i][4]), why: s_(v[i][5]), un: Number(v[i][6] || 0)
    };
    map[id] = r; rows.push(r);
  }
  return { map: map, rows: rows, sh: sh };
}
function setState_(id, name, st, why) {
  var S = states_(), r = S.map[id], t = now_();
  if (r) {
    S.sh.getRange(r.row, 2, 1, 5).setValues([[name || r.name, st, r.start || t, t, why || ""]]);
  } else {
    S.sh.appendRow([id, name, st, t, t, why || "", 0]);
  }
}

function pickFor_(seed) {
  var c = conf_(), b = bank_();
  var cats = s_(c["抽題類別"]).split(/[,，]/).map(s_).filter(function (x) { return x !== ""; });
  var pool = b.filter(function (x) { return !cats.length || cats.indexOf(x.cat) >= 0; });
  var n = parseInt(s_(c["抽題數"]), 10);
  if (!n || n < 1) n = 5;
  n = Math.min(n, pool.length);
  var sel = shuffle_(pool, seedRand_("Q|" + seed)).slice(0, n);
  return sel.map(function (x) {
    return {
      no: x.no, q: x.q, multi: x.multi, lvl: x.lvl, cat: x.cat,
      opts: shuffle_(x.opts, seedRand_("O|" + seed + "|" + x.no))
    };
  });
}

/***** API *****/
function tips_() {
  var t = s_(conf_()["考前提示"]);
  if (!t) return [];
  return t.split(/[｜|\n\r]+/).map(s_).filter(function (x) { return x !== ""; });
}

function apiMode_() {
  var c = conf_();
  return { ok: true, mode: mode_(), title: s_(c["測驗名稱"]) || "神經系統練習", tips: tips_() };
}

function apiQuiz_(p) {
  var id = s_(p.xh), name = s_(p.name);
  if (!id) return { ok: false, why: "請填學號" };
  if (!name) return { ok: false, why: "請填姓名" };
  var m = matchRoster_(id, name);
  if (!m.ok) return m;

  var a = answers_();
  if (a.byId[id] || a.byName[name]) {
    var r = a.byId[id] || a.byName[name];
    return { ok: true, done: true, time: r.time, why: "這個學號／姓名已經繳交過了" };
  }

  var w = windowState_();
  if (!w.ok) return w;

  var mo = mode_();
  if (mo === "考試") {
    var S = states_(), st = S.map[id];
    if (st && st.st === "已失效") {
      return { ok: false, dead: true, why: "本次作答已失效（" + (st.why || "離開頁面") + "）。請找老師解除鎖定後重新作答。" };
    }
    if (st && st.st === "進行中") {
      setState_(id, name, "已失效", "重新整理或重新開啟頁面");
      return { ok: false, dead: true, why: "考試模式下頁面只能開啟一次。偵測到重新開啟，本次作答已失效，請找老師解除鎖定。" };
    }
    setState_(id, name, "進行中", "");
  }

  var qs = pickFor_(id);
  if (!qs.length) return { ok: false, why: "題庫沒有符合條件的啟用題目" };
  return {
    ok: true, done: false, mode: mo, name: name, cls: m.cls, loose: !!m.loose,
    title: s_(conf_()["測驗名稱"]) || "神經系統練習", qs: qs
  };
}

function apiStatus_(p) {
  var id = s_(p.xh), name = s_(p.name), a = answers_();
  var r = (id && a.byId[id]) || (name && a.byName[name]) || null;
  if (r) return { ok: true, done: true, time: r.time };
  if (mode_() === "考試" && id) {
    var st = states_().map[id];
    if (st && st.st === "已失效") return { ok: true, done: false, dead: true, why: st.why };
  }
  return { ok: true, done: false };
}

function apiVoid_(p) {
  var id = s_(p.xh), name = s_(p.name), why = s_(p.why) || "離開頁面";
  if (!id) return { ok: false, why: "缺少學號" };
  if (mode_() !== "考試") return { ok: true, ignored: true, why: "練習模式不處理" };
  var a = answers_();
  if (a.byId[id] || a.byName[name]) return { ok: true, ignored: true, why: "已繳交，不處理" };
  var st = states_().map[id];
  if (st && st.st === "已失效") return { ok: true, already: true };
  setState_(id, name, "已失效", why);
  return { ok: true, dead: true };
}

function apiUnlock_(p) {
  var id = s_(p.xh);
  if (!id) return { ok: false, why: "請提供學號" };
  var S = states_(), r = S.map[id];
  if (!r) return { ok: false, why: "查無這個學號的作答狀態紀錄" };
  if (r.st === "已繳交") return { ok: false, why: "這個學號已經繳交完成，不需要解鎖" };
  S.sh.getRange(r.row, 3, 1, 4).setValues([["已解鎖", r.start, now_(), "老師解除鎖定"]]);
  S.sh.getRange(r.row, 7).setValue(r.un + 1);
  return { ok: true, why: "已解除鎖定：" + id + "（" + r.name + "）可以重新作答" };
}

function apiSubmit_(p) {
  var id = s_(p.xh), name = s_(p.name);
  if (!id || !name) return { ok: false, why: "學號與姓名都要填" };
  var m = matchRoster_(id, name);
  if (!m.ok) return m;
  var w = windowState_();
  if (!w.ok) return w;

  var items;
  try { items = JSON.parse(p.payload || "[]"); } catch (e) { return { ok: false, why: "作答內容格式錯誤" }; }
  if (!items.length) return { ok: false, why: "沒有收到作答內容" };

  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch (e) { return { ok: false, why: "系統忙碌，請再按一次送出" }; }
  try {
    var a = answers_();
    if (a.byId[id] || a.byName[name]) {
      var old = a.byId[id] || a.byName[name];
      return { ok: true, done: true, time: old.time, why: "這個學號／姓名已經繳交過了，不能重複送出" };
    }

    if (mode_() === "考試") {
      var st = states_().map[id];
      if (!st) return { ok: false, dead: true, why: "找不到你的作答階段紀錄，無法交卷。請找老師處理。" };
      if (st.st === "已失效") {
        return { ok: false, dead: true, why: "本次作答已失效（" + (st.why || "離開頁面") + "），無法交卷。請找老師解除鎖定。" };
      }
    }

    var b = bank_(), map = {};
    b.forEach(function (x) { map[x.no] = x; });

    var right = 0, detail = [], nos = [];
    items.forEach(function (it) {
      var q = map[s_(it.no)];
      if (!q) return;
      var mine = (it.a || []).map(s_).filter(function (x) { return x !== ""; });
      var ok = mine.length === q.ans.length && q.ans.every(function (x) { return mine.indexOf(x) >= 0; });
      if (ok) right++;
      nos.push(q.no);
      detail.push({ no: q.no, a: mine, ok: ok ? 1 : 0 });
    });
    if (!nos.length) return { ok: false, why: "作答的題號都不在題庫中" };

    var sh = ss_().getSheetByName(SH_ANS) || ss_().insertSheet(SH_ANS);
    if (sh.getLastRow() === 0) {
      sh.appendRow(["繳交時間", "學號", "姓名", "班級", "答對題數", "出題數", "題號順序", "作答明細JSON", "userAgent", "備註"]);
      sh.setFrozenRows(1);
    }
    var t = now_();
    sh.appendRow([t, id, name, m.cls, right, nos.length, nos.join(","), JSON.stringify(detail),
      s_(p.ua).slice(0, 300), m.loose ? "名單未填學號，僅比對姓名" : ""]);
    if (mode_() === "考試") setState_(id, name, "已繳交", "");
    return { ok: true, done: true, time: t };
  } finally {
    lock.releaseLock();
  }
}

function apiStats_() {
  var b = bank_(), map = {}, order = [];
  b.forEach(function (x) { map[x.no] = x; order.push(x.no); });
  var a = answers_(), agg = {};
  a.rows.forEach(function (r) {
    var d;
    try { d = JSON.parse(r.detail || "[]"); } catch (e) { d = []; }
    d.forEach(function (it) {
      var no = s_(it.no);
      if (!agg[no]) agg[no] = { n: 0, right: 0, counts: {} };
      agg[no].n++;
      if (it.ok) agg[no].right++;
      var pick = (it.a && it.a.length) ? it.a : ["未答"];
      pick.forEach(function (o) { agg[no].counts[o] = (agg[no].counts[o] || 0) + 1; });
    });
  });

  var perQ = [];
  order.forEach(function (no) {
    if (!agg[no]) return;
    var q = map[no], g = agg[no];
    perQ.push({
      no: no, q: q.q, opts: q.opts, ans: q.ans, why: q.why, cat: q.cat, multi: q.multi,
      n: g.n, right: g.right, rate: g.n ? Math.round(g.right / g.n * 1000) / 10 : 0, counts: g.counts
    });
  });

  var done = {}, allRight = 0, someWrong = 0;
  a.rows.forEach(function (r) {
    if (r.id) done[r.id] = 1;
    if (r.name) done[r.name] = 1;
    if (Number(r.right) === Number(r.total)) allRight++; else someWrong++;
  });

  var rost = roster_(), missing = [], idFilled = 0;
  rost.forEach(function (p) {
    if (p.id) idFilled++;
    if (done[p.name] || (p.id && done[p.id])) return;
    missing.push(p.name);
  });

  var S = states_();
  var open = S.rows.filter(function (r) { return r.st !== "已繳交"; }).map(function (r) {
    return { id: r.id, name: r.name, st: r.st, start: r.start, upd: r.upd, why: r.why, un: r.un };
  });

  var c = conf_(), tz = ss_().getSpreadsheetTimeZone();
  return {
    ok: true, n: a.rows.length, allRight: allRight, someWrong: someWrong,
    rosterN: rost.length, rosterIdFilled: idFilled, missing: missing, perQ: perQ,
    mode: mode_(), states: open,
    conf: {
      pick: s_(c["抽題數"]), cats: s_(c["抽題類別"]), cls: s_(c["適用班級"]),
      start: s_(c["開放開始時間"]) ? Utilities.formatDate(parseTime_(c["開放開始時間"]), tz, "yyyy/MM/dd HH:mm") : "",
      end: s_(c["開放結束時間"]) ? Utilities.formatDate(parseTime_(c["開放結束時間"]), tz, "yyyy/MM/dd HH:mm") : ""
    },
    updated: now_()
  };
}

function apiBank_() {
  return { ok: true, items: bank_() };
}

/***** 入口 *****/
function doGet(e) {
  var p = (e && e.parameter) || {};
  var act = s_(p.action);
  if (!act) return legacyStats_(s_(p.sheet));
  try {
    if (act === "mode") return json_(apiMode_());
    if (act === "quiz") return json_(apiQuiz_(p));
    if (act === "submit") return json_(apiSubmit_(p));
    if (act === "status") return json_(apiStatus_(p));
    if (act === "void") return json_(apiVoid_(p));
    if (act === "unlock") return json_(apiUnlock_(p));
    if (act === "stats") return json_(apiStats_());
    if (act === "bank") return json_(apiBank_());
    return json_({ ok: false, why: "未知的 action：" + act });
  } catch (err) {
    return json_({ ok: false, why: "後端錯誤：" + (err && err.message ? err.message : err) });
  }
}

/* 舊版相容：warm-up.html / gcs-practice.html 的送出與統計，不要刪 */
function doPost(e) {
  var d = JSON.parse(e.postData.contents);
  var name = d.sheet || "未命名";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0 && d.header) { sh.appendRow(d.header); sh.setFrozenRows(1); }
  sh.appendRow(d.row);
  return ContentService.createTextOutput("ok");
}

function legacyStats_(name) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  var out = { n: 0, cols: [] };
  if (sh && sh.getLastRow() >= 1) {
    var v = sh.getDataRange().getValues(), head = v[0];
    out.n = v.length - 1;
    for (var c = 0; c < head.length; c++) {
      var h = String(head[c]);
      if (h === "時間" || h === "姓名") continue;
      var counts = {};
      for (var r = 1; r < v.length; r++) {
        var k = String(v[r][c]); if (k === "") k = "未答";
        counts[k] = (counts[k] || 0) + 1;
      }
      out.cols.push({ h: h, counts: counts });
    }
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
