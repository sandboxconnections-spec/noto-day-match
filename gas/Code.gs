// Noto Day Match — 受け口（Google Apps Script のウェブアプリ）
// GitHub Pages のページから届いた回答を、スプレッドシートの「responses」に1人1行で書き込む。
// 「集計」シートは初回に自動で作る（#test の回答は集計から外す）。

var SPREADSHEET_ID = "1jhlc2UbIrCczlgbYzWo3WWjbCVu7i_7QsuRrAltGEMM"; // 「Noto Day Match 回答」（空なら、紐づいているスプレッドシートを使う）

var CARDS = [
  {id:"heal-near",    k:"静", theme:"癒し",         dist:"近い", title:"A Quiet Day by Nanao Bay"},
  {id:"heal-far",     k:"禅", theme:"癒し",         dist:"遠い", title:"Zen at Sojiji Soin"},
  {id:"culture-near", k:"漆", theme:"カルチャー",   dist:"近い", title:"Makers of Ipponsugi Street"},
  {id:"culture-far",  k:"匠", theme:"カルチャー",   dist:"遠い", title:"Lacquer Masters of Wajima"},
  {id:"active-near",  k:"海", theme:"アクティビティ", dist:"近い", title:"Paddle and Walk Noto Island"},
  {id:"local-near",   k:"縁", theme:"地元の人",     dist:"近い", title:"Stories of Ipponsugi"},
  {id:"food-near",    k:"鮨", theme:"食（寿司）",   dist:"近い", title:"Sushi, Three Ways"},
  {id:"food-far",     k:"酒", theme:"食（寿司）",   dist:"遠い", title:"Sake and Sushi of Oku-Noto"}
];
var CHOICES = ["love", "like", "no", "unsure"];
var STEPS = ["about", "howto", "swipe", "pick", "price", "words", "quake"];
var REGIONS = ["Europe", "North America", "Latin America", "Asia", "Oceania", "Middle East & Africa"];
var COMPANY = ["Solo", "Partner", "Friends", "Family"];
var SPEND = ["Under $50", "$50–150", "$150–400", "$400–1,000", "Over $1,000"];
var PRICE = ["Under $100", "$100–250", "$250–500", "$500–1,000", "Over $1,000"];
var WORDS = ["Unhurried", "Ma (the space between)", "Zen", "Decide nothing", "Hidden gem", "Master craftsman", "Hands-on", "Local life", "Private", "Slow travel", "Rebuilding", "Off the beaten path"];
var QUAKE = ["No", "A little", "Yes, a lot", "I didn't know about it"];
var TYPES = {heal:"Quiet Seeker", culture:"Culture Diver", active:"Active Explorer", local:"Local Connector", food:"Food Pilgrim"};
var UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/;

function headers_() {
  var h = ["received_at", "id", "created_at", "src", "test", "lang", "device", "region", "company", "max_spend",
           "pick", "pick_title", "why", "price", "quake", "type", "road_trip", "total_sec"];
  CARDS.forEach(function (c) { h.push(c.id, c.id + "_sec", c.id + "_pos"); });
  h.push("words_plus", "words_minus");
  STEPS.forEach(function (s) { h.push("step_" + s + "_sec"); });
  h.push("raw_json");
  return h;
}

// ---------- 受け取り ----------
function doPost(e) {
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    var body = e && e.postData ? e.postData.contents : "";
    if (!body || body.length > 30000) return out_({ok: false, error: "bad_size"});
    var p = JSON.parse(body);
    var err = check_(p);
    if (err) return out_({ok: false, error: err});
    lock.waitLock(20000); locked = true;
    var sh = sheet_();
    var found = findRow_(sh, p.id);
    if (found) return out_({ok: true, row: found, duplicate: true});
    sh.appendRow(toRow_(p));
    return out_({ok: true, row: sh.getLastRow()});
  } catch (x) {
    return out_({ok: false, error: String(x && x.message ? x.message : x)});
  } finally {
    if (locked) lock.releaseLock();
  }
}

// 引数なし＝動作確認。?id=<その回答のid> ＝その1行だけを返す（idを知っている人にしか読めない）
function doGet(e) {
  var id = e && e.parameter ? String(e.parameter.id || "") : "";
  var sh = sheet_();
  if (!id) return out_({ok: true, service: "noto-day-match", rows: Math.max(0, sh.getLastRow() - 1)});
  if (!UUID_RE.test(id)) return out_({ok: false, error: "bad_id"});
  var r = findRow_(sh, id);
  if (!r) return out_({ok: true, found: false});
  var h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var v = sh.getRange(r, 1, 1, h.length).getValues()[0];
  var o = {};
  h.forEach(function (k, i) { o[k] = v[i] instanceof Date ? Utilities.formatDate(v[i], "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss") : v[i]; });
  return out_({ok: true, found: true, row: r, data: o});
}

// 初回の許可と、シートの準備（エディタから1回実行してもよい）
function setup() {
  var sh = sheet_();
  return "ready: " + sh.getParent().getUrl();
}

// ---------- 中身 ----------
function check_(p) {
  if (!p || typeof p !== "object") return "not_object";
  if (!UUID_RE.test(String(p.id || ""))) return "bad_id";
  if (!Array.isArray(p.swipes) || p.swipes.length < 1 || p.swipes.length > 20) return "bad_swipes";
  var ids = CARDS.map(function (c) { return c.id; });
  for (var i = 0; i < p.swipes.length; i++) {
    var s = p.swipes[i];
    if (!s || ids.indexOf(s.id) < 0 || CHOICES.indexOf(s.v) < 0) return "bad_swipe";
    if (typeof s.ms !== "number" || s.ms < 0 || s.ms > 3600000) return "bad_ms";
  }
  if (p.pick && ids.indexOf(p.pick) < 0) return "bad_pick";
  var a = p.about || {};
  if (a.region && REGIONS.indexOf(a.region) < 0) return "bad_region";
  if (a.company && COMPANY.indexOf(a.company) < 0) return "bad_company";
  if (a.maxSpend && SPEND.indexOf(a.maxSpend) < 0) return "bad_spend";
  if (p.price && PRICE.indexOf(p.price) < 0) return "bad_price";
  if (p.quake && QUAKE.indexOf(p.quake) < 0) return "bad_quake";
  if (p.type && !TYPES[p.type]) return "bad_type";
  if (p.why && String(p.why).length > 300) return "bad_why";
  return "";
}

function toRow_(p) {
  var by = {};
  (p.swipes || []).forEach(function (s, i) { by[s.id] = {v: s.v, ms: s.ms, pos: s.pos || ((p.order || []).indexOf(s.id) + 1) || (i + 1)}; });
  var pick = CARDS.filter(function (c) { return c.id === p.pick; })[0];
  var a = p.about || {};
  var row = [
    new Date(), p.id, safe_(p.createdAt), safe_(p.src), p.src === "test" ? "test" : "", safe_(p.lang),
    p.mobile ? "mobile" : "desktop", a.region || "", a.company || "", a.maxSpend || "",
    pick ? pick.id : "", pick ? pick.title : "", safe_(p.why), p.price || "", p.quake || "", p.type || "",
    p.roadTrip === true ? "yes" : (p.roadTrip === false ? "no" : ""), sec_(p.durationMs)
  ];
  CARDS.forEach(function (c) {
    var s = by[c.id];
    row.push(s ? s.v : "", s ? sec_(s.ms) : "", s ? s.pos : "");
  });
  var plus = [], minus = [];
  Object.keys(p.words || {}).forEach(function (w) {
    if (WORDS.indexOf(w) < 0) return;
    if (p.words[w] === 1) plus.push(w); else if (p.words[w] === -1) minus.push(w);
  });
  row.push(plus.length ? "|" + plus.join("|") + "|" : "", minus.length ? "|" + minus.join("|") + "|" : "");
  STEPS.forEach(function (s) { row.push(p.steps && typeof p.steps[s] === "number" ? sec_(p.steps[s]) : ""); });
  row.push(safe_(JSON.stringify(p)));
  return row;
}

function sec_(ms) { return typeof ms === "number" ? Math.round(ms / 100) / 10 : ""; }

// セルに入れる文字は、先頭が = + - @ なら数式として動かないように ' を付ける
function safe_(v) {
  if (v === null || v === undefined) return "";
  var s = String(v).slice(0, 45000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function ss_() {
  return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_() {
  var ss = ss_();
  var sh = ss.getSheetByName("responses");
  var h = headers_();
  if (!sh) {
    sh = ss.insertSheet("responses", 0);
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight("bold");
    sh.setFrozenRows(1);
    sh.getRange("A:A").setNumberFormat("yyyy-mm-dd hh:mm:ss");
    var first = ss.getSheetByName("シート1") || ss.getSheetByName("Sheet1");
    if (first && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);
  }
  if (!ss.getSheetByName("集計")) buildSummary_(ss, h);
  return sh;
}

function findRow_(sh, id) {
  var last = sh.getLastRow();
  if (last < 2 || !id) return 0;
  var hit = sh.getRange(2, 2, last - 1, 1).createTextFinder(String(id)).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : 0;
}

function col_(n) {
  var s = "";
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// ---------- 集計シート ----------
function buildSummary_(ss, h) {
  var sm = ss.insertSheet("集計", 1);
  function R(name) { var c = col_(h.indexOf(name) + 1); return "responses!$" + c + "$2:$" + c; }
  var T = R("test"), ID = R("id");
  var rows = [];
  rows.push(["集計（リンク末尾に #test を付けた回答は除く）", "", "", "", "", "", "", "", "", "", ""]);
  rows.push(["回答数", '=COUNTIFS(' + ID + ',"<>",' + T + ',"<>test")', "", "", "", "", "", "", "", "", ""]);
  rows.push(["", "", "", "", "", "", "", "", "", "", ""]);
  rows.push(["カード", "中身", "距離", "答えた人", "Love it", "Like", "No", "Not sure", "Love＋Like", "迷った秒（中央値）", "「1日だけ」に選んだ人"]);
  var cardRow = {};
  CARDS.forEach(function (c, i) {
    var r = 5 + i; cardRow[c.id] = r;
    var C = R(c.id), S = R(c.id + "_sec"), P = R("pick");
    rows.push([
      c.k + " " + c.title, c.theme, c.dist,
      '=COUNTIFS(' + C + ',"<>",' + T + ',"<>test")',
      '=IFERROR(COUNTIFS(' + C + ',"love",' + T + ',"<>test")/$D' + r + ',"")',
      '=IFERROR(COUNTIFS(' + C + ',"like",' + T + ',"<>test")/$D' + r + ',"")',
      '=IFERROR(COUNTIFS(' + C + ',"no",' + T + ',"<>test")/$D' + r + ',"")',
      '=IFERROR(COUNTIFS(' + C + ',"unsure",' + T + ',"<>test")/$D' + r + ',"")',
      '=IFERROR(E' + r + '+F' + r + ',"")',
      '=IFERROR(MEDIAN(FILTER(' + S + ',' + S + '<>"",' + T + '<>"test")),"")',
      '=COUNTIFS(' + P + ',"' + c.id + '",' + T + ',"<>test")'
    ]);
  });
  sm.getRange(1, 1, rows.length, 11).setValues(rows);
  sm.getRange(5, 5, CARDS.length, 5).setNumberFormat("0%");
  sm.getRange(5, 10, CARDS.length, 1).setNumberFormat("0.0");

  var r0 = 5 + CARDS.length + 1;
  var pairs = [["癒し", "heal-near", "heal-far"], ["カルチャー", "culture-near", "culture-far"], ["食（寿司）", "food-near", "food-far"]];
  var block = [["近い版と遠い版（Love＋Like）", "近い", "遠い", "差（遠い−近い）"]];
  pairs.forEach(function (p, i) {
    var rr = r0 + 1 + i;
    block.push([p[0], "=I" + cardRow[p[1]], "=I" + cardRow[p[2]], '=IFERROR(C' + rr + '-B' + rr + ',"")']);
  });
  sm.getRange(r0, 1, block.length, 4).setValues(block);
  sm.getRange(r0 + 1, 2, pairs.length, 3).setNumberFormat("0%");

  var r = r0 + block.length + 1;
  function counts(title, colName, list, contains) {
    var b = [[title, "人数", "割合"]];
    list.forEach(function (v, i) {
      var crit = contains ? '"*|' + v.replace(/"/g, '""') + '|*"' : '"' + v.replace(/"/g, '""') + '"';
      b.push([v, '=COUNTIFS(' + R(colName) + ',' + crit + ',' + T + ',"<>test")', '=IFERROR(B' + (r + 1 + i) + '/$B$2,"")']);
    });
    sm.getRange(r, 1, b.length, 3).setValues(b);
    sm.getRange(r + 1, 3, list.length, 1).setNumberFormat("0%");
    sm.getRange(r, 1, 1, 3).setFontWeight("bold");
    r += b.length + 1;
  }
  counts("「いくらに見える？」（選んだ1日）", "price", PRICE, false);
  counts("地震のあとの心配", "quake", QUAKE, false);
  counts("旅のタイプ", "type", Object.keys(TYPES), false);
  counts("どこから", "region", REGIONS, false);
  counts("誰と旅する", "company", COMPANY, false);
  counts("1日の体験に払った最高額", "max_spend", SPEND, false);
  counts("言葉：○（惹かれる）", "words_plus", WORDS, true);
  counts("言葉：×（引っかかる）", "words_minus", WORDS, true);

  sm.getRange(1, 1).setFontWeight("bold");
  sm.getRange(4, 1, 1, 11).setFontWeight("bold");
  sm.getRange(r0, 1, 1, 4).setFontWeight("bold");
  sm.setColumnWidth(1, 300);
  sm.setFrozenRows(0);
}
