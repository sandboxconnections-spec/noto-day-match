// Noto Day Match — 受け口（Google Apps Script のウェブアプリ）
// GitHub Pages のページから届いた回答を、スプレッドシートの「responses」に1人1行で書き込む。
// 「集計」シートは初回に自動で作る（#test の回答は集計から外す）。
// 版が変わって列の並びが変わったときは、古いシートの名前に日時を付けて残し、新しいシートを作り直す。

var SPREADSHEET_ID = "1jhlc2UbIrCczlgbYzWo3WWjbCVu7i_7QsuRrAltGEMM"; // 「Noto Day Match 回答」（空なら、紐づいているスプレッドシートを使う）

var CARDS = [
  {id:"heal-near",    k:"静", theme:"癒し",           dist:"近い", title:"A Quiet Day by Nanao Bay"},
  {id:"heal-far",     k:"禅", theme:"癒し",           dist:"遠い", title:"Zen at Sojiji Soin"},
  {id:"culture-near", k:"漆", theme:"カルチャー",     dist:"近い", title:"Makers of Ipponsugi Street"},
  {id:"culture-far",  k:"匠", theme:"カルチャー",     dist:"遠い", title:"Lacquer Masters of Wajima"},
  {id:"active-near",  k:"海", theme:"アクティビティ", dist:"近い", title:"Paddle and Walk Noto Island"},
  {id:"local-near",   k:"縁", theme:"地元の人",       dist:"近い", title:"Stories of Ipponsugi"},
  {id:"food-near",    k:"鮨", theme:"食（寿司）",     dist:"近い", title:"Sushi, Three Ways"},
  {id:"food-far",     k:"酒", theme:"食（寿司）",     dist:"遠い", title:"Sake and Sushi of Oku-Noto"}
];
var PAIRS = [
  {theme:"heal",    ja:"癒し",       near:"heal-near",    far:"heal-far"},
  {theme:"culture", ja:"カルチャー", near:"culture-near", far:"culture-far"},
  {theme:"food",    ja:"食（寿司）", near:"food-near",    far:"food-far"}
];
var STEPS = ["about", "cards", "pairs", "reason", "pick", "price", "wplus", "wminus", "quake"];
var REGIONS = ["Europe", "North America", "Latin America", "Asia", "Oceania", "Middle East & Africa", "Japan"];
var COMPANY = ["Solo", "Partner", "Friends", "Family"];
var SPEND = ["Under $50", "$50–150", "$150–400", "$400–1,000", "Over $1,000"];
var PRICE = ["Under $100", "$100–250", "$250–500", "$500–1,000", "Over $1,000"];
var REASONS = ["Too long in the car", "I don't know these places", "Not my kind of day", "Looks expensive", "Something else"];
var WORDS = ["Unhurried", "Ma (the space between)", "Zen", "Decide nothing", "Hidden gem", "Master craftsman", "Hands-on", "Local life", "Private", "Slow travel", "Rebuilding", "Off the beaten path"];
var QUAKE = ["No", "A little", "Yes, a lot", "I didn't know about it"];
var TYPES = {heal:"Quiet Seeker", culture:"Culture Diver", active:"Active Explorer", local:"Local Connector", food:"Food Pilgrim"};
var TYPE_JA = {heal:"癒し", culture:"カルチャー", active:"アクティビティ", local:"地元の人", food:"食"};
var UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/;

function headers_() {
  var h = ["received_at", "id", "created_at", "src", "test", "lang", "device", "region", "company", "max_spend",
           "pick", "pick_title", "why", "price", "quake", "type", "road_trip", "total_sec"];
  CARDS.forEach(function (c) { h.push(c.id, c.id + "_sec", c.id + "_pos"); });
  PAIRS.forEach(function (p) { h.push("pair_" + p.theme); });
  h.push("far_reasons", "words_plus", "words_minus");
  STEPS.forEach(function (s) { h.push("step_" + s + "_sec"); });
  h.push("memo", "raw_json");
  h.push("ui_lang"); // v4で追加（末尾に足す＝既存の列の位置は変えない）
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
    if (p && p.kind === "memo") {
      if (!UUID_RE.test(String(p.id || ""))) return out_({ok: false, error: "bad_id"});
      if (typeof p.memo !== "string" || !p.memo.trim() || p.memo.length > 500) return out_({ok: false, error: "bad_memo"});
      lock.waitLock(20000); locked = true;
      var shm = sheet_();
      var rm = findRow_(shm, p.id);
      if (!rm) return out_({ok: false, error: "not_found"});
      shm.getRange(rm, headers_().indexOf("memo") + 1).setValue(safe_(p.memo.trim()));
      return out_({ok: true, row: rm, memo: true});
    }
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
  if (!id) return out_({ok: true, service: "noto-day-match", version: 3, rows: Math.max(0, sh.getLastRow() - 1)});
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
function inList_(list, v) { return list.indexOf(v) >= 0; }
function allIn_(list, arr) { return Array.isArray(arr) && arr.length <= list.length && arr.every(function (x) { return inList_(list, x); }); }

function check_(p) {
  if (!p || typeof p !== "object") return "not_object";
  if (!UUID_RE.test(String(p.id || ""))) return "bad_id";
  if (!Array.isArray(p.ratings) || p.ratings.length < 1 || p.ratings.length > 20) return "bad_ratings";
  var ids = CARDS.map(function (c) { return c.id; });
  for (var i = 0; i < p.ratings.length; i++) {
    var s = p.ratings[i];
    if (!s || ids.indexOf(s.id) < 0) return "bad_rating_id";
    if ([0, 1, 2, 3, 4, 5].indexOf(s.r) < 0) return "bad_rating";
    if (typeof s.ms !== "number" || s.ms < 0 || s.ms > 3600000) return "bad_ms";
  }
  var pairs = p.pairs || {};
  for (var k in pairs) { if (["heal", "culture", "food"].indexOf(k) < 0 || ["near", "far"].indexOf(pairs[k]) < 0) return "bad_pair"; }
  if (p.farReasons && !allIn_(REASONS, p.farReasons)) return "bad_reason";
  if (p.wordsPlus && !allIn_(WORDS, p.wordsPlus)) return "bad_words";
  if (p.wordsMinus && !allIn_(WORDS, p.wordsMinus)) return "bad_words";
  if (p.pick && ids.indexOf(p.pick) < 0) return "bad_pick";
  var a = p.about || {};
  if (a.region && !inList_(REGIONS, a.region)) return "bad_region";
  if (a.company && !inList_(COMPANY, a.company)) return "bad_company";
  if (a.maxSpend && !inList_(SPEND, a.maxSpend)) return "bad_spend";
  if (p.price && !inList_(PRICE, p.price)) return "bad_price";
  if (p.quake && !inList_(QUAKE, p.quake)) return "bad_quake";
  if (p.type && !TYPES[p.type]) return "bad_type";
  if (p.why && String(p.why).length > 300) return "bad_why";
  if (p.ui && ["en", "ja"].indexOf(p.ui) < 0) return "bad_ui";
  return "";
}

function bar_(arr) { return arr && arr.length ? "|" + arr.join("|") + "|" : ""; }

function toRow_(p) {
  var by = {};
  (p.ratings || []).forEach(function (s, i) { by[s.id] = {r: s.r, ms: s.ms, pos: s.pos || ((p.order || []).indexOf(s.id) + 1) || (i + 1)}; });
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
    row.push(s ? (s.r === 0 ? "?" : s.r) : "", s ? sec_(s.ms) : "", s ? s.pos : "");
  });
  PAIRS.forEach(function (pp) { row.push((p.pairs || {})[pp.theme] || ""); });
  row.push(bar_(p.farReasons), bar_(p.wordsPlus), bar_(p.wordsMinus));
  STEPS.forEach(function (s) { row.push(p.steps && typeof p.steps[s] === "number" ? sec_(p.steps[s]) : ""); });
  row.push("", safe_(JSON.stringify(p)));
  row.push(p.ui === "ja" ? "ja" : "en");
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
  if (ss.getSpreadsheetTimeZone() !== "Asia/Tokyo") ss.setSpreadsheetTimeZone("Asia/Tokyo"); // 新規シートは米国時間で作られるため
  var h = headers_();
  var tag = Utilities.formatDate(new Date(), "Asia/Tokyo", "MMdd_HHmm");
  var sh = ss.getSheetByName("responses");
  if (sh) {
    var cur = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0];
    var same = cur.join("\t") === h.join("\t");
    var grown = !same && cur.length < h.length && h.slice(0, cur.length).join("\t") === cur.join("\t");
    if (grown) {
      // 列が末尾に足されただけ＝見出しだけ足す（これまでの回答はそのまま）
      sh.getRange(1, cur.length + 1, 1, h.length - cur.length).setValues([h.slice(cur.length)]).setFontWeight("bold");
    } else if (!same) {
      // 列の並びが変わった＝古いシートは名前を変えて残し、作り直す
      sh.setName("responses_old_" + tag);
      ["集計", "集計（英語）", "集計（日本語）"].forEach(function (n) { var s = ss.getSheetByName(n); if (s) s.setName(n + "_old_" + tag); });
      sh = null;
    }
  }
  if (!sh) {
    sh = ss.insertSheet("responses", 0);
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight("bold");
    sh.setFrozenRows(1);
    sh.getRange("A:A").setNumberFormat("yyyy-mm-dd hh:mm:ss");
    var first = ss.getSheetByName("シート1") || ss.getSheetByName("Sheet1");
    if (first && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);
  }
  // v3までの「集計」は言語を分けていない＝名前を変えて残し、英語と日本語に分けて作り直す
  var legacy = ss.getSheetByName("集計");
  if (legacy) legacy.setName("集計_old_" + tag);
  if (!ss.getSheetByName("集計（英語）")) buildSummary_(ss, h, "集計（英語）", false);
  if (!ss.getSheetByName("集計（日本語）")) buildSummary_(ss, h, "集計（日本語）", true);
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
// isJa=false：英語の画面で答えた人（v3までの回答＝ui_lang が空の行も英語に数える）／isJa=true：日本語の画面で答えた人
function buildSummary_(ss, h, name, isJa) {
  var sm = ss.insertSheet(name, isJa ? 2 : 1);
  function R(n) { var c = col_(h.indexOf(n) + 1); return "responses!$" + c + "$2:$" + c; }
  var T = R("test"), ID = R("id"), L = R("ui_lang");
  var LC = "," + L + "," + (isJa ? '"ja"' : '"<>ja"');   // COUNTIFS / AVERAGEIFS に足す条件
  var LF = "," + L + (isJa ? '="ja"' : '<>"ja"');        // FILTER に足す条件
  var NT = T + ',"<>test"' + LC;                          // 「テストを除く・この言語」の共通条件
  var W = 12;
  function pad(row) { while (row.length < W) row.push(""); return row; }
  var rows = [];
  rows.push(pad(["集計（" + (isJa ? "日本語" : "英語") + "の画面で答えた人・リンク末尾に #test を付けた回答は除く）"]));
  rows.push(pad(["回答数", '=COUNTIFS(' + ID + ',"<>",' + NT + ')']));
  rows.push(pad([]));
  rows.push(pad(["カード", "中身", "距離", "答えた人", "平均（5点満点）", "予約する（4〜5）", "迷う（3）", "予約しない（1〜2）", "分からない（?）", "迷った秒（中央値）", "「1日だけ」に選んだ人", "表示順の平均"]));
  var cardRow = {};
  CARDS.forEach(function (c, i) {
    var r = 5 + i; cardRow[c.id] = r;
    var C = R(c.id), S = R(c.id + "_sec"), P = R("pick"), O = R(c.id + "_pos");
    rows.push([
      c.k + " " + c.title, c.theme, c.dist,
      '=COUNTIFS(' + C + ',"<>",' + NT + ')',
      '=IFERROR(AVERAGEIFS(' + C + ',' + NT + '),"")',
      '=IFERROR(COUNTIFS(' + C + ',">=4",' + NT + ')/$D' + r + ',"")',
      '=IFERROR(COUNTIFS(' + C + ',3,' + NT + ')/$D' + r + ',"")',
      '=IFERROR(COUNTIFS(' + C + ',"<=2",' + C + ',">=1",' + NT + ')/$D' + r + ',"")',
      '=IFERROR(COUNTIFS(' + C + ',"~?",' + NT + ')/$D' + r + ',"")',
      '=IFERROR(MEDIAN(FILTER(' + S + ',' + S + '<>"",' + T + '<>"test"' + LF + ')),"")',
      '=COUNTIFS(' + P + ',"' + c.id + '",' + NT + ')',
      '=IFERROR(AVERAGEIFS(' + O + ',' + NT + '),"")'
    ]);
  });
  sm.getRange(1, 1, rows.length, W).setValues(rows);
  sm.getRange(5, 5, CARDS.length, 1).setNumberFormat("0.0");
  sm.getRange(5, 6, CARDS.length, 4).setNumberFormat("0%");
  sm.getRange(5, 10, CARDS.length, 1).setNumberFormat("0.0");
  sm.getRange(5, 12, CARDS.length, 1).setNumberFormat("0.0");

  var r0 = 5 + CARDS.length + 1;
  var block = [["近い版と遠い版", "予約する（近い）", "予約する（遠い）", "差（遠い−近い）", "二択で遠い版を選んだ割合", "二択に答えた人"]];
  PAIRS.forEach(function (p, i) {
    var rr = r0 + 1 + i, PC = R("pair_" + p.theme);
    block.push([p.ja, "=F" + cardRow[p.near], "=F" + cardRow[p.far], '=IFERROR(C' + rr + '-B' + rr + ',"")',
      '=IFERROR(COUNTIFS(' + PC + ',"far",' + NT + ')/F' + rr + ',"")',
      '=COUNTIFS(' + PC + ',"<>",' + NT + ')']);
  });
  sm.getRange(r0, 1, block.length, 6).setValues(block);
  sm.getRange(r0 + 1, 2, PAIRS.length, 4).setNumberFormat("0%");
  sm.getRange(r0, 1, 1, 6).setFontWeight("bold");

  var r = r0 + block.length + 1;
  function counts(title, colName, items, contains) {
    var b = [[title, "人数", "割合"]];
    items.forEach(function (it, i) {
      var label = Array.isArray(it) ? it[0] : it, v = Array.isArray(it) ? it[1] : it;
      var q = String(v).replace(/"/g, '""').replace(/[~*?]/g, function (ch) { return "~" + ch; });
      var crit = contains ? '"*|' + q + '|*"' : '"' + q + '"';
      b.push([label, '=COUNTIFS(' + R(colName) + ',' + crit + ',' + NT + ')', '=IFERROR(B' + (r + 1 + i) + '/$B$2,"")']);
    });
    sm.getRange(r, 1, b.length, 3).setValues(b);
    sm.getRange(r + 1, 3, items.length, 1).setNumberFormat("0%");
    sm.getRange(r, 1, 1, 3).setFontWeight("bold");
    r += b.length + 1;
  }
  counts("遠い版で引っかかった理由（遠い版に3以下を付けた人）", "far_reasons", REASONS, true);
  counts("「いくらに見える？」（選んだ1日）", "price", PRICE, false);
  counts("地震のあとの心配", "quake", QUAKE, false);
  counts("旅のタイプ（「1日だけ」の中身）", "type", Object.keys(TYPES).map(function (k) { return [TYPE_JA[k] + "（" + TYPES[k] + "）", k]; }), false);
  counts("どこから", "region", REGIONS, false);
  counts("誰と旅する", "company", COMPANY, false);
  counts("1日の体験に払った最高額", "max_spend", SPEND, false);
  counts("言葉：惹かれる", "words_plus", WORDS, true);
  counts("言葉：引っかかる", "words_minus", WORDS, true);

  sm.getRange(1, 1).setFontWeight("bold");
  sm.getRange(4, 1, 1, W).setFontWeight("bold");
  sm.setColumnWidth(1, 320);
}
