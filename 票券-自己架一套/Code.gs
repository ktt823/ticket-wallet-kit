/**
 * ════════════════════════════════════════════════════════════════
 *  票券 — Google Apps Script 後端
 *  資料存在你自己的 Google Sheets，截圖存在你自己的雲端硬碟。
 *
 *  第一次使用：
 *   1. 在編輯器選 setup 這個函式，按執行，授權。
 *   2. 看「執行紀錄」，把印出來的 OWNER_KEY 抄下來。
 *   3. 指令碼屬性填入 GEMINI_KEY（Google AI Studio 免費申請）。
 *   4. 部署 → 新增部署作業 → 網頁應用程式
 *        執行身分：我
 *        誰可以存取：所有人      ← 分享連結要給沒有 Google 帳號的人點，所以必須是這個
 *   5. 把 /exec 結尾的網址和 OWNER_KEY 貼進 index.html 的「設定」。
 * ════════════════════════════════════════════════════════════════
 */

/** Gemini 模型。若哪天這個名稱被下架，去 https://ai.google.dev/gemini-api/docs/models 換一個 flash 系列的。 */
var MODEL = 'gemini-3.5-flash-lite';

/**
 * 這份後端程式的版本，跟網站（index.html、sw.js）用同一個號碼、一起加一。
 * App 會拿它跟自己比，網站和後端只更新了一邊時，會提醒還差哪一邊。
 */
var CODE_VERSION = 94;

var T_SHEET = 'Tickets';
var S_SHEET = 'Shares';
var T_HEAD = ['id','createdAt','merchant','title','value','code','codeType','expiry',
              'status','note','fileId','fingerprint','rawText','usedAt','usedBy','source','category',
              'usedVia','kind','balance','balanceUrl','balanceAt','pendingAt','pendingBy','usedTime'];
// v88 usedTime = 用掉的時間（HH:mm:ss，台灣時間）。票券變成「已使用」的那一刻記下來（updateTicket_），改回未使用就清掉；
//   usedAt 還是只有日期。v88 以前用掉的沒有時間（空白）；其中從分享連結標的，ticketView_ 會從 usedVia 推回來
// usedVia = 「分享token@標記時間」，只有從分享頁按「我用掉了」才會有，用來判斷對方能不能復原
// v69 餘額型票券（可以分次使用，用一次不代表用光）。新欄位都接在最後面，舊票券那幾格是空的、不影響：
//   kind = 'balance' 就是餘額型；空白 = 一般票券（用一次就沒了）
//   balance = 最近一次確認的剩餘金額（元，整數）；空白 = 還沒確認過
//   balanceUrl = 查餘額的網址（選填；只存在你的試算表，不會放進通知或分享頁）
//   balanceAt = 上次確認餘額的日期（YYYY-MM-DD）
//   pendingAt / pendingBy = 「用過了、還沒確認剩多少」：哪天、誰用的（你自己是「我」，朋友是他填的名字）；空白 = 沒有待確認
// category 空白 = 讓前端依券名自動判斷；填了值就以你填的為準
// source 空白 = 截圖是這個程式存進雲端硬碟的，刪票券時連檔案一起丟掉
// source 'drive' = 直接連到你原本就有的檔案，刪票券時不會動到那個檔案
// 'to' 是後來才加的，所以擺在最後面 —— 舊的分享紀錄那一格會是空的，不影響
var S_HEAD = ['token','createdAt','ticketIds','note','expiresAt','revoked','openCount','lastOpenedAt','to','keepIds'];
// v76 keepIds = 這條連結裡「分享出去、你自己這邊也留著」的票券 id（逗號隔開）：誰有空誰先用，不會從你的「可用」收起來、出示前也不會先問

var MAX_SHARE = 10;   // 一個連結最多幾張票券（截圖是直接嵌在頁面裡的，太多會很慢）

/**
 * 拿到分享連結的人按了「我用掉了」之後，幾分鐘內可以自己按「復原」。
 * 只能復原自己透過同一個連結標記的，你自己標記的他動不了。
 * 設成 0 = 不限時間（但要注意：對方用掉之後又按復原，你這邊就會以為還能用）。
 */
var SHARE_UNDO_MINUTES = 30;


/* ══════════════════ 安裝 ══════════════════ */

function setup() {
  var p = PropertiesService.getScriptProperties();
  var ss;

  if (P_('SHEET_ID')) {
    ss = SpreadsheetApp.openById(P_('SHEET_ID'));
  } else {
    ss = SpreadsheetApp.create('票券資料');
    p.setProperty('SHEET_ID', ss.getId());
  }
  ensureSheet_(ss, T_SHEET, T_HEAD);
  ensureSheet_(ss, S_SHEET, S_HEAD);
  ensureSheet_(ss, R_SHEET, R_HEAD);

  if (!P_('FOLDER_ID')) {
    p.setProperty('FOLDER_ID', DriveApp.createFolder('票券截圖').getId());
  }
  if (!P_('OWNER_KEY')) {
    p.setProperty('OWNER_KEY', randomToken_(24));
  }

  // 每天早上的到期提醒排程順便裝好（要設了 Discord 才會真的發出去）
  var remind;
  try {
    ensureReminderTrigger_();
    remind = '到期提醒：每天早上 ' + REMIND_HOUR + ' 點多（設了 Discord 才收得到）';
  } catch (e) {
    remind = '到期提醒排程沒裝好：' + e.message + '（之後執行 startReminder 就好）';
  }

  var lines = [
    '',
    '───────── 票券安裝完成 ─────────',
    '試算表：' + ss.getUrl(),
    '截圖資料夾：https://drive.google.com/drive/folders/' + P_('FOLDER_ID'),
    '',
    'OWNER_KEY（貼到前端「設定」裡）：',
    '    ' + P_('OWNER_KEY'),
    '',
    (P_('ANTHROPIC_KEY') || P_('GEMINI_KEY') || P_('VISION_KEY'))
      ? '辨識引擎：' + engineName_()
      : '⚠ 還沒設定辨識引擎。到「專案設定 → 指令碼屬性」新增下面其中一支金鑰：\n' +
        '    GEMINI_KEY     https://aistudio.google.com/apikey （有免費額度）\n' +
        '    ANTHROPIC_KEY  https://console.claude.com （付費，約每張截圖 NT$0.1）',
    '通知：' + ([
      P_('DISCORD_WEBHOOK') ? 'Discord' : '',
      (P_('LINE_TOKEN') && P_('LINE_USER_ID')) ? 'LINE' : ''
    ].filter(function (x) { return x; }).join('、') || '未設定（選用，見 README「通知」）') +
    '　執行 testNotify 可以試發一則',
    remind,
    '',
    '接著：部署 → 新增部署作業 → 網頁應用程式（執行身分：我／存取權：所有人）',
    '──────────────────────────────'
  ];
  Logger.log(lines.join('\n'));
}

function ensureSheet_(ss, name, head) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.getRange(1, 1, sh.getMaxRows(), head.length).setNumberFormat('@');  // 全部當文字，日期才不會被自動轉格式
  var first = sh.getRange(1, 1, 1, head.length).getValues()[0];
  if (first.join('') !== head.join('')) {
    sh.getRange(1, 1, 1, head.length).setValues([head]);
  }
  sh.setFrozenRows(1);
  var def = ss.getSheetByName('工作表1') || ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);
  return sh;
}


/* ══════════════════ 入口 ══════════════════ */

function doGet(e) {
  startRequest_();
  var p = (e && e.parameter) || {};
  try {
    // raw=1 是 _worker.js 用的：回 JSON 包著的 HTML，前面不會有 Apps Script 的沙箱外框
    // me=1 是你自己在 App 按「開啟」看的，不算進「對方打開過幾次」
    if (p.s && p.raw) return json_({ ok: true, data: { html: pageHtml_(p.s, p.me === '1') } });
    if (p.s) return htmlOut_(pageHtml_(p.s, p.me === '1'));
  } catch (err) {
    if (p.raw) return json_({ ok: false, error: String(err.message || err) });
    return htmlOut_(shell_('打不開', '<p class="msg">' + esc_(err.message) + '</p>'));
  }
  return htmlOut_(shell_('票券', '<p class="msg">票券後端運作中。這個網址要貼到 App 的「設定」裡。</p>'));
}

function doPost(e) {
  var t0 = startRequest_();
  var req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: '請求格式不對' });
  }

  var out;
  try {
    // 分享頁的「我用掉了」：只認分享 token，不需要 OWNER_KEY
    if (req.action === 'markUsed') {
      out = { ok: true, data: markUsedByShare_(req.token, req.id, req.by) };
    } else if (req.action === 'unmarkUsed') {
      out = { ok: true, data: unmarkUsedByShare_(req.token, req.id) };
    } else if (req.action === 'returnShare') {
      // 分享頁的「用不到，還給他」：一樣只認分享 token
      out = { ok: true, data: returnByShare_(req.token, req.id, req.by) };
    } else if (req.action === 'memberUse' || req.action === 'memberUndo' || req.action === 'memberImage') {
      // v75 親友頁：只認親友連結的碼，只能動他看得到的那幾張
      out = { ok: true, data: req.action === 'memberUse' ? memberUse_(req.token, req.id)
                            : req.action === 'memberUndo' ? memberUndo_(req.token, req.id)
                            : memberImage_(req.token, req.id) };
    } else if (req.action === 'deployPoll') {
      // 電腦上的部署小幫手每分鐘問一次；只回答「要不要跑」，不收任何指令
      out = { ok: true, data: deployPoll_() };
      if (!out.data.run) return json_(out);          // 沒事就不記執行紀錄，免得洗版
    } else if (req.action === 'relayCheck') {
      // _worker.js 轉送 Discord 之前來確認（只回對或不對，號碼只能用一次）
      return json_({ ok: true, data: relayCheck_(req) });
    } else if (req.action === 'deployNotice') {
      // 電腦上雙擊「一鍵部署」跑完，deploy.js 來報結果（內容由後端自己組，一分鐘最多一則）
      out = { ok: true, data: deployNotice_(req) };
    } else if (req.action === 'deployReport') {
      out = { ok: true, data: deployReport_(req.id, req.ok, req.text) };
    } else {
      var owner = P_('OWNER_KEY');
      if (!owner) throw new Error('後端還沒執行過 setup()');
      if (req.key !== owner) throw new Error('存取碼不對');
      out = { ok: true, data: dispatch_(req) };
    }
    logRequest_(req.action, t0, req.at);
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
    logRequest_(req.action, t0, req.at, out.error);
  }
  return json_(out);
}

/** 你自己在 App 裡改票券。狀態有變就順便通知自己（NOTIFY_SELF 可以關掉）。 */
function updateOwn_(id, patch, at) {
  var before, t;
  withLock_(function () {
    before = findTicket_(id);
    perf_('找票');
    if (!before) throw new Error('找不到這張票券');
    guardBalanceUse_(before, patch);        // v69：餘額型不能直接變「已使用」
    t = updateTicket_(id, patch, before);
    perf_('寫入');
  });
  var wasUsed = before.status === 'used';
  var isUsed = t.status === 'used';

  if (NOTIFY_SELF && wasUsed !== isUsed) {
    try {
      notify_({ kind: isUsed ? 'used' : 'undo', who: t.usedBy || '我', ticket: t, self: true, at: at });
    } catch (e) {
      Logger.log('自己標記的通知沒送出：' + e.message);   // 通知失敗不能影響票券本身
    }
    perf_('通知');
  }
  return { ticket: t };
}


/* ══════════════════ 餘額型票券（v69） ══════════════════
 *
 *  餘額型 = 可以分次使用，用一次不代表用光（例如 SOGO「好禮即享券（餘額型）」、商品卡）。
 *  標「一次抵用型」的券用一次就作廢，不在這裡。
 *  規矩：餘額型的券不能直接變「已使用」，要先確認剩餘金額、剩 0 才可以。
 *  這條規矩由後端把關，App 的每個按鈕、舊版 App、分享頁都一樣。
 */
var BAL_STRONG = /餘額型|額型|分次使用|可分次|多次使用|儲值|商品卡/;   // 標題寫了這些字就當餘額型（「額型」是怕辨識漏掉「餘」字；商品卡一律算餘額型）
var BAL_WEAK   = /電子禮券|禮物卡|禮品卡|點數卡/;                   // 通常也能分次，但標題沒寫，只列「疑似」，要你自己確認
var BAL_ONCE   = /一次抵用|一次性|單次|限用一次|僅限一次/;         // 寫了這些就不是餘額型
var BAL_LABEL  = /餘額|剩餘|可用金額|可用額度|目前金額|尚餘|還剩|剩下/;    // v70：讀到的餘額，證據那行字一定要有這類字（只寫「面額」的不算）
var BAL_COUNT  = /\d\s*(?:次|張|杯|個|份|顆|瓶|片|組|入)/;                  // v70：「剩餘 2 杯」「還剩 3 次」是數量不是金額，不當餘額（不然寄杯券會被誤標成餘額型）
var BAL_STALE_DAYS = 60;                                         // 確認過的餘額超過幾天沒再確認，晨間提醒會列出來

/** 'strong' = 標題寫了餘額型／分次使用／商品卡；'weak' = 電子禮券之類疑似；'' = 看不出來，或明寫一次抵用 */
function balanceGuess_(t) {
  t = t || {};
  // 改名規則改過名的，原名也一起看（備註會寫「已改名（原：…）」）
  var hay = [t.title, t.note, t.merchant, t.renamed && t.renamed.title].join(' ');
  if (BAL_ONCE.test(hay)) return '';
  if (BAL_STRONG.test(hay)) return 'strong';
  if (BAL_WEAK.test(hay)) return 'weak';
  return '';
}

/**
 * 匯入時：標題明寫餘額型的，自動標成餘額型（疑似的不動，留給你在「設定」裡確認）。
 * v70：票面上明寫著餘額的（辨識時讀到 balance，見 tidyDraft_），本身就是餘額型的證據，也自動標
 *      （除非明寫「一次抵用」之類）。標成餘額型又讀到餘額的，順便記下哪天確認的；不是餘額型的不收餘額。
 * 出任何問題都原樣放行。
 */
function flagBalance_(d) {
  try {
    if (!d) return d;
    if (!d.kind) {
      var once = BAL_ONCE.test([d.title, d.note, d.merchant, d.renamed && d.renamed.title].join(' '));
      if (balanceGuess_(d) === 'strong' || (d.balance && !once)) d.kind = 'balance';
    }
    if (d.kind !== 'balance') { d.balance = ''; d.balanceText = ''; }
    else if (d.balance && !d.balanceAt) d.balanceAt = today_();
  } catch (e) {}
  return d;
}

/** 金額只留數字（「NT$3,987」「1,200 元」「500.00」都收）；不是整數金額回空字串 */
function cleanAmount_(v) {
  var s = String(v == null ? '' : v).replace(/NT\$|\$|元|,|\s/gi, '').replace(/\.0+$/, '');
  return (/^\d+$/.test(s) && s.length <= 9) ? String(Number(s)) : '';
}

/**
 * 辨識（或讀截圖）讀到的餘額，要有「餘額／剩餘金額…」那行字當證據，而且證據裡真的有這個數字才收。
 * 只寫「面額」的、模型猜的、證據對不上的、寫的是數量（「剩餘 2 杯」「還剩 3 次」）的，一律當作沒讀到（寧可空著，讓你自己確認）。
 */
function balanceChecked_(balance, evidence) {
  var n = cleanAmount_(balance), t = String(evidence || '');
  return (n && BAL_LABEL.test(t) && !BAL_COUNT.test(t) && t.replace(/[,\s]/g, '').indexOf(n) >= 0) ? n : '';
}

/** 剩餘金額是不是「確認過是 0」（空白、不是數字都不算） */
function isZeroBalance_(v) {
  var s = String(v == null ? '' : v).replace(/[,\s]/g, '');
  return /^\d+(\.\d+)?$/.test(s) && Number(s) === 0;
}

/** 餘額型要變「已使用」，剩餘金額必須同時是 0；其他的修改、改回未使用都不管 */
function guardBalanceUse_(before, patch) {
  if (!patch || patch.status !== 'used') return;
  var kind = patch.kind !== undefined ? patch.kind : before.kind;
  if (kind !== 'balance') return;
  var bal = patch.balance !== undefined ? patch.balance : before.balance;
  if (isZeroBalance_(bal)) return;
  throw new Error('這張是餘額型，要先確認剩餘金額（剩 0 才會進「已使用」）');
}

/* ══════════════════ 用網址查餘額（v71：PayEasy 電子憑證） ══════════════════
 *
 *  PayEasy 賣的電子禮券（萬家福／樂家康、康達盛通…）有一個不用登入的網址
 *    https://www.payeasy.com.tw/Electronic/coupon?id=<12 碼>
 *  頁面的資料來自 https://www.payeasy.com.tw/Electronic/eticket/<id>（JSON：券名、面額、餘額、狀態、序號、條碼圖）。
 *  用「序號」把網址對到 App 裡的票券，再讀餘額。三個動作（都要存取碼）：
 *    payeasyPeek    只讀：每個網址對到 App 的哪一張（沒存任何東西）
 *    payeasyApply   你勾好之後才寫：把網址存進那張券並更新餘額，或用網站的資料新增一張
 *    balanceRefresh 一鍵更新：讀每張有網址的券；變少或一樣直接寫，剩 0 才進「已使用」，變多要你確認
 *  安全：只准讀 www.payeasy.com.tw（網址不對的一律不讀）；網址等於那張券的存取碼，只存在你的試算表，
 *  不放進通知和分享頁；序號對不上的券不寫（避免網址貼到別張券上）。
 */
var PE_URL_RE = /^https?:\/\/(?:www\.)?payeasy\.com\.tw\/Electronic\/coupon\?(?:[^#\s]*&)?id=([A-Za-z0-9]{6,40})(?:[&#]\S*)?$/i;
var PE_IMG_RE = /^https:\/\/www\.payeasy\.com\.tw\/Electronic\/image\/[0-9a-f]+\/[0-9a-f]+\/[A-Za-z0-9]+$/;
var PE_MAX = 40;                    // 一次最多處理幾個網址／幾張券
var PE_NEAR_DIFF = 3;               // 序號只差這幾位以內，列成「可能是這張」（要你自己勾；多半是辨識錯字）
var PE_UA = 'Mozilla/5.0 (Linux; Android 13) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Mobile Safari/537.36';

function peId_(url) {
  var m = PE_URL_RE.exec(String(url == null ? '' : url).trim());
  return m ? m[1] : '';
}
function peUrl_(id) { return 'https://www.payeasy.com.tw/Electronic/coupon?id=' + id; }
function digits_(s) { return String(s == null ? '' : s).replace(/\D/g, ''); }

/** 網站寫「2027/01/04止」→ 2027-01-04；「無期限」→ 空白 */
function peExpiry_(s) {
  var m = /(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})/.exec(String(s || ''));
  return m ? m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2) : '';
}

/** 兩串數字差幾位（長度不一樣回 99） */
function peDiff_(a, b) {
  if (a.length !== b.length) return 99;
  var n = 0;
  for (var i = 0; i < a.length; i++) if (a.charAt(i) !== b.charAt(i)) n++;
  return n;
}

function peOpts_(id) {
  return {
    method: 'get', muteHttpExceptions: true, followRedirects: true,
    headers: { 'User-Agent': PE_UA, 'Accept': 'application/json, text/plain, */*', 'Referer': peUrl_(id) }
  };
}

/** 把網站回的一包資料整理成我們要的幾欄；任何不對勁都回 {ok:false, why} */
function peRead_(code, text) {
  if (code === 401 || code === 403 || code === 429) return { ok: false, why: '網站不讓後端讀（HTTP ' + code + '）' };
  if (code === 404 || code === 500) return { ok: false, why: '網站找不到這張券（網址可能不對）' };
  if (code !== 200) return { ok: false, why: '網站回 ' + code };
  var j = parseJson_(text);
  if (!j || typeof j !== 'object') return { ok: false, why: '網站回的不是券的資料（可能被擋了）' };
  if (String(j.couponType || '') !== 'BALANCE') return { ok: false, why: '網站標示這張不是餘額型（' + String(j.couponType || '不明') + '），沒處理' };
  var serial = digits_(j.fullNum), bal = cleanAmount_(j.balance);
  if (serial.length < 8) return { ok: false, why: '網站沒有給序號' };
  if (bal === '') return { ok: false, why: '網站沒有給餘額' };
  return {
    ok: true,
    d: {
      name: String(j.name || ''), face: cleanAmount_(j.dvalue), balance: bal, serial: serial,
      status: String(j.status || ''), statusDesc: String(j.statusDesc || ''),
      expiry: peExpiry_(j.expiryDate),
      image: PE_IMG_RE.test(String(j.imagePath || '')) ? String(j.imagePath) : ''
    }
  };
}

/** 一次讀好幾個 id（同一個 id 只讀一次）；回 { id: peRead_ 的結果 } */
function peFetchMany_(ids) {
  var uniq = [], seen = {}, out = {};
  ids.forEach(function (id) { if (id && !seen[id]) { seen[id] = 1; uniq.push(id); } });
  var res = null;
  try {
    res = UrlFetchApp.fetchAll(uniq.map(function (id) {
      var o = peOpts_(id);
      o.url = 'https://www.payeasy.com.tw/Electronic/eticket/' + id;
      return o;
    }));
  } catch (e) { res = null; }                      // 一起讀失敗就一個一個讀，只有真的連不上的那個算失敗
  uniq.forEach(function (id, i) {
    try {
      var x = res ? res[i] : UrlFetchApp.fetch('https://www.payeasy.com.tw/Electronic/eticket/' + id, peOpts_(id));
      out[id] = peRead_(x.getResponseCode(), x.getContentText());
    } catch (e) {
      out[id] = { ok: false, why: '連不上 PayEasy：' + String((e && e.message) || e) };
    }
  });
  return out;
}

/** 條碼圖（網站給的 PNG）。只抓 www.payeasy.com.tw 的圖，抓不到就回 null（券照樣新增，只是沒有圖） */
function peImage_(url, id) {
  if (!PE_IMG_RE.test(String(url || ''))) return null;
  try {
    var r = UrlFetchApp.fetch(url, peOpts_(id));
    if (r.getResponseCode() !== 200) return null;
    var b = r.getBlob(), bytes = b.getBytes(), mime = String(b.getContentType() || '').split(';')[0];
    if (!/^image\//i.test(mime) || bytes.length < 100 || bytes.length > 2 * 1024 * 1024) return null;
    return { blob: b, mime: mime };
  } catch (e) { return null; }
}

/** 網站上的券名 → 商家＋券名（再套你的改名規則）。商家認不出來就放 PayEasy */
function peNames_(name) {
  var s = String(name || '').replace(/\s+/g, ' ').replace(/^【[^】]*】\s*/, '').replace(/[<＜]/g, '（').replace(/[>＞]/g, '）').trim();
  var merchant = '', title = s, m;
  if (/萬家福|樂家康|康達盛通|PROSPERITY/i.test(s)) {
    // 萬家福／樂家康的電子禮券在 PayEasy 上有好幾種寫法（康達盛通是同一家的公司名），一律歸在「萬家福/樂家康」，券名只留後面（跟手動整理過的名字一樣）
    merchant = '萬家福/樂家康';
    title = s.replace(/萬家福\/樂家康|PROSPERITY PLAZA\/Uni-PROSPERITY|康達盛通|Mia C['’]bon/gi, '').trim();
  } else if ((m = /^(.*?)(電子禮券|電子禮卡|禮券|禮物卡|商品卡)/.exec(s)) && m[1].trim()) {
    merchant = m[1].trim();
    title = s.slice(m[1].length).trim();
  }
  var d = { merchant: merchant || 'PayEasy', title: title || s || '電子禮券' };
  applyNameRules_(d);
  return d;
}

/** 票券一列 → App 看到的樣子（listTickets_ 也用這個；sh 是分享資訊，新券沒有） */
function ticketView_(t, sh) {
  sh = sh || {};
  return {
    id: t.id, createdAt: t.createdAt, merchant: t.merchant, title: t.title,
    value: t.value, code: t.code, codeType: t.codeType, expiry: t.expiry,
    status: t.status || 'active', note: t.note, fileId: t.fileId,
    usedAt: t.usedAt, usedBy: t.usedBy, category: t.category,
    usedTime: t.status === 'used' ? (t.usedTime || usedClockFromVia_(t.usedVia)) : '',
    kind: t.kind, balance: t.balance, balanceUrl: t.balanceUrl, balanceAt: t.balanceAt,
    pendingAt: t.pendingAt, pendingBy: t.pendingBy,
    sharedTo: sh.sharedTo || '', sharedAt: sh.sharedAt || '',
    shareToken: sh.shareToken || '', shareExpires: sh.shareExpires || '',
    shareKeep: !!(sh.shareToken && sh.keep),
    shareUrl: sh.shareToken ? shareUrl_(sh.shareToken) : '',
    // 連結被打開過幾次、最後哪天（對方還沒打開過 = 很可能忘了，該提醒）
    shareOpens: sh.shareToken ? sh.opens : 0, shareOpenedAt: sh.lastOpen || ''
  };
}

function peName_(t) { return ticketLabel_(t); }

/** 貼進來的網址整理成清單：去空白、去重複、認不出的留著報錯（不讀） */
function peList_(urls) {
  if (typeof urls === 'string') urls = urls.split(/[\s,]+/);
  var out = [], seen = {};
  (urls || []).forEach(function (u) {
    u = String(u == null ? '' : u).trim();
    if (!u) return;
    var id = peId_(u), key = id || u;
    if (seen[key]) return;
    seen[key] = 1;
    out.push({ input: u, id: id });
  });
  if (out.length > PE_MAX) throw new Error('一次最多 ' + PE_MAX + ' 個網址');
  return out;
}

/** 只讀：每個網址讀到什麼、對到 App 的哪一張 */
function payeasyPeek_(urls) {
  var list = peList_(urls);
  if (!list.length) throw new Error('沒有網址');
  var sites = peFetchMany_(list.map(function (x) { return x.id; }));
  var all = tickets_(), taken = {}, items = [];
  // 先配序號完全一樣的，再用剩下的券找「差幾位」的
  list.forEach(function (x) {
    var it = { url: x.id ? peUrl_(x.id) : '', input: x.input, id: x.id, ok: false, match: { kind: 'none' } };
    items.push(it);
    if (!x.id) { it.error = '這不是 PayEasy 的券網址（要長得像 https://www.payeasy.com.tw/Electronic/coupon?id=…）'; return; }
    var s = sites[x.id];
    if (!s.ok) { it.error = s.why; return; }
    it.ok = true; it.d = s.d;
    it.site = { name: s.d.name, face: s.d.face, balance: s.d.balance, status: s.d.status, statusDesc: s.d.statusDesc,
                expiry: s.d.expiry, serialTail: s.d.serial.slice(-4), hasImage: !!s.d.image };
    for (var i = 0; i < all.length; i++) {
      var t = all[i];
      if (taken[t.id] || digits_(t.code) !== s.d.serial) continue;
      taken[t.id] = 1;
      it.match = { kind: 'exact', id: t.id, title: peName_(t), balance: t.balance, status: t.status || 'active',
                   replaces: !!(t.balanceUrl && peId_(t.balanceUrl) !== x.id) };
      break;
    }
  });
  items.forEach(function (it) {
    if (!it.ok || it.match.kind !== 'none') return;
    var best = null, bestDiff = 99;
    all.forEach(function (t) {
      if (taken[t.id]) return;
      var diff = peDiff_(digits_(t.code), it.d.serial), viaUrl = peId_(t.balanceUrl) === it.id;
      if (viaUrl && diff > PE_NEAR_DIFF) diff = PE_NEAR_DIFF;       // 這張已經存著同一個網址：序號對不上，多半是辨識錯字
      if (diff <= PE_NEAR_DIFF && diff < bestDiff) { best = t; bestDiff = diff; }
    });
    if (best) {
      taken[best.id] = 1;
      it.match = { kind: 'near', id: best.id, title: peName_(best), balance: best.balance, status: best.status || 'active',
                   diff: bestDiff, viaUrl: peId_(best.balanceUrl) === it.id,
                   replaces: !!(best.balanceUrl && peId_(best.balanceUrl) !== it.id) };
    }
  });
  items.forEach(function (it) { delete it.d; });
  return { items: items };
}

/** 把網址存進這張券並更新餘額（序號要一樣；fixCode＝你確認過是這張，連序號一起校正成網站上的） */
function peLink_(ticketId, id, d, fixCode) {
  var cur = findTicket_(ticketId);
  if (!cur) return { ok: false, error: '找不到這張票券' };
  if (digits_(cur.code) !== d.serial && !fixCode) return { ok: false, error: '序號對不上，沒有存' };
  var patch = { kind: 'balance', balanceUrl: peUrl_(id), balance: d.balance, balanceAt: today_(), pendingAt: '', pendingBy: '' };
  if (digits_(cur.code) !== d.serial) patch.code = d.serial;
  var zero = isZeroBalance_(d.balance);
  if (zero && cur.status !== 'used') { patch.status = 'used'; patch.usedAt = today_(); patch.usedBy = '我'; }
  updateOwn_(ticketId, patch, Date.now());
  var r = { ok: true, action: zero && cur.status !== 'used' ? 'linked-used' : 'linked', id: ticketId, patch: patch,
            balance: d.balance, before: cur.balance };
  if (cur.status === 'used' && !zero) r.warn = 'App 標的是「已使用」，但網站寫還剩 $' + d.balance + '；要用的話把它改回未使用';
  return r;
}

/** 用網站的資料新增一張券（條碼圖當截圖）；App 裡已經有同序號的就不新增 */
function peCreate_(id, d) {
  var dupId = '';
  tickets_().forEach(function (t) { if (!dupId && digits_(t.code) === d.serial) dupId = t.id; });
  if (dupId) return { ok: false, error: 'App 裡已經有這張了（序號一樣），沒有重複新增', id: dupId };
  var nm = peNames_(d.name);
  var img = peImage_(d.image, id), fileId = '';
  if (img) {
    try {
      var folder = DriveApp.getFolderById(P_('FOLDER_ID'));
      var blob = img.blob.setName(today_() + '-payeasy-' + d.serial.slice(-4) + '.png');
      fileId = folder.createFile(blob).getId();
    } catch (e) { fileId = ''; }
  }
  var zero = isZeroBalance_(d.balance);
  var t = saveTicket_({
    merchant: nm.merchant, title: nm.title, value: d.face ? 'NT$' + d.face : '', code: d.serial, codeType: 'barcode',
    expiry: d.expiry, status: zero ? 'used' : 'active', note: nm.renamed ? renamedFlag_(nm) : '',
    fileId: fileId, rawText: '', source: '', category: '',
    usedAt: zero ? today_() : '', usedBy: zero ? '我' : '', usedTime: zero ? clock_() : '',
    kind: 'balance', balance: d.balance, balanceUrl: peUrl_(id), balanceAt: today_(), pendingAt: '', pendingBy: ''
  });
  return { ok: true, action: 'created', id: t.id, ticket: ticketView_(t), image: !!fileId };
}

/** 套用：items = [{url, ticketId, fixCode} 或 {url, create:true}]；每一項各自成功或失敗，網站的資料這時候重新讀一次 */
function payeasyApply_(items) {
  items = items || [];
  if (!items.length) throw new Error('沒有要套用的項目');
  if (items.length > PE_MAX) throw new Error('一次最多 ' + PE_MAX + ' 項');
  var sites = peFetchMany_(items.map(function (it) { return peId_(it && it.url); }));
  var results = items.map(function (it) {
    var id = peId_(it && it.url);
    var r;
    try {
      if (!id) r = { ok: false, error: '這不是 PayEasy 的券網址' };
      else if (!sites[id].ok) r = { ok: false, error: sites[id].why };
      else if (it.ticketId) r = peLink_(String(it.ticketId), id, sites[id].d, !!it.fixCode);
      else if (it.create) r = peCreate_(id, sites[id].d);
      else r = { ok: false, error: '沒說要配到哪張、也沒說要新增' };
    } catch (e) {
      r = { ok: false, error: String((e && e.message) || e) };
    }
    r.url = id ? peUrl_(id) : String((it && it.url) || '');
    return r;
  });
  return { results: results };
}

/**
 * 一鍵更新：讀每一張「餘額型、還沒用完、存了 PayEasy 網址」的券。
 *   req.ids   只更新這幾張（空白＝全部）　req.dry  只讀、不寫（App 的「從網址讀餘額」用）
 *   req.accept  你確認過「網站寫的比 App 記的多」的券 id（沒確認的不寫，只回報）
 * state：same 沒變／set 原本沒確認過、現在記下／down 變少／zero 用完了／up 變多（要確認）／mismatch 序號對不上／error 讀不到
 */
function balanceRefresh_(req) {
  req = req || {};
  var want = null, accept = {};
  if (req.ids && req.ids.length) { want = {}; req.ids.forEach(function (i) { want[String(i)] = 1; }); }
  (req.accept || []).forEach(function (i) { accept[String(i)] = 1; });
  var rows = tickets_().filter(function (t) {
    return t.kind === 'balance' && t.status !== 'used' && peId_(t.balanceUrl) && (!want || want[t.id]);
  });
  var left = Math.max(0, rows.length - PE_MAX);
  rows = rows.slice(0, PE_MAX);
  var sites = peFetchMany_(rows.map(function (t) { return peId_(t.balanceUrl); }));
  var today = today_();
  var results = rows.map(function (t) {
    var s = sites[peId_(t.balanceUrl)];
    var r = { id: t.id, name: peName_(t), before: t.balance, after: '' };
    if (!s.ok) { r.state = 'error'; r.why = s.why; return r; }
    var d = s.d;
    r.after = d.balance; r.face = d.face;
    if (digits_(t.code) !== d.serial) { r.state = 'mismatch'; r.why = '網址的序號跟這張券對不上，沒有更新'; return r; }
    var was = cleanAmount_(t.balance), now = Number(d.balance);
    var patch = { balance: d.balance, balanceAt: today, pendingAt: '', pendingBy: '' };
    if (was === '') r.state = 'set';
    else if (Number(was) === now) r.state = 'same';
    else if (now === 0) r.state = 'zero';
    else if (now < Number(was)) r.state = 'down';
    else r.state = 'up';
    if (r.state === 'zero') { patch.status = 'used'; patch.usedAt = today; patch.usedBy = '我'; }
    if (r.state === 'same' && t.balanceAt === today && !t.pendingAt) return r;     // 今天已經確認過，不用再寫
    if (req.dry || (r.state === 'up' && !accept[t.id])) return r;                  // 變多了：你沒確認就只回報，不寫
    try {
      updateOwn_(t.id, patch, Date.now());
      r.patch = patch; r.wrote = true;
    } catch (e) {
      r.state = 'error'; r.why = String((e && e.message) || e);
    }
    return r;
  });
  return { results: results, left: left };
}

function dispatch_(req) {
  switch (req.action) {
    case 'payeasyPeek':    return payeasyPeek_(req.urls);
    case 'payeasyApply':   return payeasyApply_(req.items);
    case 'balanceRefresh': return balanceRefresh_(req);
    case 'ping':   return { count: tickets_().length, engine: engineName_(), codeVersion: CODE_VERSION };
    case 'list':   return { tickets: listTickets_(), members: membersBrief_(), codeVersion: CODE_VERSION };
    case 'parse':  return parse_(req.image, req.mime, req.filename);
    case 'save':   return { ticket: saveTicket_(req.ticket) };
    case 'update': return updateOwn_(req.id, req.patch, req.at);
    case 'remove': return removeTicket_(req.id);
    case 'image':  return imageOf_(req.id);
    case 'share':  return makeShare_(req.ids, req.days, req.note, req.to, { token: req.token, later: !!req.later, instant: !!req.instant, force: !!req.force, until: req.until, move: !!req.move, keep: req.keep });
    case 'members':      return { members: listMembers_() };                 // v75 親友連結
    case 'memberSave':   return memberSave_(req.member);
    case 'memberRevoke': return memberSet_(req.id, { revoke: true });
    case 'memberRenew':  return memberSet_(req.id, { renew: true });
    case 'warm':   return { codeVersion: CODE_VERSION };     // v72：App 一進分享畫面就先叫醒後端（什麼都不讀）
    case 'shareNotify': return shareNotify_(req.token);
    case 'shares': return { shares: listShares_() };
    case 'revoke': return revokeShare_(req.token);
    case 'unshare': return unshareTicket_(req.token, req.id);
    case 'unshareMany': return unshareMany_(req.items);
    case 'shareKeep': return setShareKeep_(req.token, req.ids, !!req.keep);       // v86：分享出去之後再改「我也留著」
    case 'rulesInfo': return rulesInfo_();
    case 'rulesImport': return rulesImport_(req.text);
    case 'readBalance': return readBalance_(req.image, req.mime);
    case 'importNow': return importNow_();
    case 'importReport': return importReport_(req);
    case 'testNotify': return testNotifyNow_();
    case 'notifyLog': return notifyLog_();
    case 'notifyFlush': return notifyFlushNow_();
    case 'previewReminder': return sendReminder_(true);
    case 'fixReminder': return fixReminder_();
    case 'appUpdated': return appUpdated_(req);
    case 'requestDeploy': return requestDeploy_('App');
    case 'deployStatus': return deployStatus_();
    default:       throw new Error('不認得的動作：' + req.action);
  }
}


/* ══════════════════ 資料層 ══════════════════ */

/*
 * 同一次請求裡，指令碼屬性和試算表都只開一次。
 * 以前每讀一個設定、每找一次票都重新去問一次 Google，一次「標記已使用」要開三次試算表。
 * PROPS_ 只在網頁請求（doGet/doPost）時啟用；在編輯器跑的函式照舊每次現讀，改了設定馬上生效。
 */
var PROPS_ = null, SS_ = null, TSH_ = null, SSH_ = null;
function P_(k) {
  if (PROPS_) return PROPS_[k] || '';
  return PropertiesService.getScriptProperties().getProperty(k) || '';
}
function ss_()  { return SS_ || (SS_ = SpreadsheetApp.openById(P_('SHEET_ID'))); }
function tSh_() {
  if (!TSH_) {
    TSH_ = ss_().getSheetByName(T_SHEET);
    ensureTCols_(TSH_);
  }
  return TSH_;
}

/**
 * v69 起票券表多了幾欄。舊的試算表第一次被用到時，自動把欄補齊、標題列寫好
 * （每個版本只做一次，記在指令碼屬性 TCOLS）。讀寫票券表都照 T_HEAD.length 欄走，
 * 表的欄數比它少會出錯，所以所有讀寫票券表的地方都要先經過 tSh_()。
 */
function ensureTCols_(sh) {
  if (!sh) return;
  var want = T_HEAD.length;
  if (P_('TCOLS') === String(want)) return;
  withLock_(function () {
    var first;
    try {
      first = sh.getRange(1, 1, 1, want).getValues()[0];
    } catch (e) {                                   // 表的欄數比 T_HEAD 少：先補欄
      var have = sh.getMaxColumns();
      sh.insertColumnsAfter(have, want - have);
      first = sh.getRange(1, 1, 1, want).getValues()[0];
    }
    if (first.join('') !== T_HEAD.join('')) sh.getRange(1, 1, 1, want).setValues([T_HEAD]);
    sh.getRange(1, 1, sh.getMaxRows(), want).setNumberFormat('@');   // 全部當文字，日期、金額才不會被自動轉格式
    PropertiesService.getScriptProperties().setProperty('TCOLS', String(want));
  });
  if (PROPS_) PROPS_.TCOLS = String(want);
}
function sSh_() {
  if (!SSH_) {
    SSH_ = ss_().getSheetByName(S_SHEET);
    ensureSCols_(SSH_);
  }
  return SSH_;
}

/** v76 分享紀錄表多了 keepIds 一欄：舊的試算表第一次被用到時自動把標題補上（每個版本只做一次，記在指令碼屬性 SCOLS） */
function ensureSCols_(sh) {
  if (!sh) return;
  var want = S_HEAD.length;
  if (P_('SCOLS') === String(want)) return;
  withLock_(function () {
    var first;
    try {
      first = sh.getRange(1, 1, 1, want).getValues()[0];
    } catch (e) {
      var have = sh.getMaxColumns();
      sh.insertColumnsAfter(have, want - have);
      first = sh.getRange(1, 1, 1, want).getValues()[0];
    }
    if (first.join('') !== S_HEAD.join('')) sh.getRange(1, 1, 1, want).setValues([S_HEAD]);
    sh.getRange(1, 1, sh.getMaxRows(), want).setNumberFormat('@');
    PropertiesService.getScriptProperties().setProperty('SCOLS', String(want));
  });
  if (PROPS_) PROPS_.SCOLS = String(want);
}

/*
 * 寫入排隊：同一時間只讓一個請求寫試算表，寫完才輪下一個。
 *
 * 為什麼要：刪票券是整列刪掉、下面往上補；新增是寫在「最後一列的下一列」。
 * 兩個動作剛好同時（連續刪兩張、朋友按「我用掉了」時你剛好在刪、自動匯入寫入時你按儲存），
 * 後面那個照著舊的列號寫，就會刪錯、蓋錯別張。所以「找到那一列 → 寫進去」要整段包在鎖裡，
 * 而且一定在鎖裡面才找列號。通知、搬檔這些慢的事放在鎖外面，不拖住別人。
 */
var LOCK_WAIT_MS = 20000;
var LOCK_DEPTH_ = 0;
function withLock_(fn) {
  if (LOCK_DEPTH_ > 0) return fn();            // 已經在鎖裡面（外層函式拿過了），直接做
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(LOCK_WAIT_MS)) throw new Error('後端正忙，請稍後再試一次');
  LOCK_DEPTH_++;
  try {
    return fn();
  } finally {
    LOCK_DEPTH_--;
    try { SpreadsheetApp.flush(); } catch (e) {}   // 確定真的寫進去了，下一個拿到鎖的才讀得到
    lock.releaseLock();
  }
}

/* ── 計時：每個網頁請求在「執行項目」留一行，看得出時間花在哪 ── */
var PERF_ = null, LAST_DISCORD_ = null, LAST_LINE_ = null, FLUSHED_ = 0;
function perf_(name) {
  if (!PERF_) return;
  var now = Date.now();
  PERF_.laps.push(name + ' ' + ((now - PERF_.last) / 1000).toFixed(2) + 's');
  PERF_.last = now;
}
function startRequest_() {
  var t0 = Date.now();
  PERF_ = { t0: t0, last: t0, laps: [] };
  LAST_DISCORD_ = null; LAST_LINE_ = null; FLUSHED_ = 0;
  PROPS_ = PropertiesService.getScriptProperties().getProperties();
  perf_('讀設定');
  return t0;
}
function logRequest_(action, t0, at, err) {
  try {
    var parts = ['[' + (action || '?') + '] 後端共 ' + ((Date.now() - t0) / 1000).toFixed(2) + 's'];
    if (PERF_ && PERF_.laps.length) parts.push('（' + PERF_.laps.join('｜') + '）');
    var lag = at ? (t0 - Number(at)) / 1000 : NaN;
    if (lag > -30 && lag < 600) parts.push('App 按下→後端開始 ' + lag.toFixed(1) + 's');
    if (LAST_DISCORD_) {
      parts.push('Discord ' + LAST_DISCORD_.code + '，花 ' + LAST_DISCORD_.ms + 'ms' +
                 (LAST_DISCORD_.at ? '，Discord 伺服器 ' + LAST_DISCORD_.at + ' 收到' : ''));
    }
    if (err) parts.push('錯誤：' + err);
    console.log(parts.join(' '));
  } catch (e) {}
}

function readRows_(sh, head, idCol) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last - 1, head.length).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var o = { _row: i + 2 };
    for (var j = 0; j < head.length; j++) o[head[j]] = str_(values[i][j]);
    if (o[idCol]) out.push(o);
  }
  return out;
}

function str_(v) {
  if (v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v);
}

function tickets_()  { return readRows_(tSh_(), T_HEAD, 'id'); }
function shares_()   { return readRows_(sSh_(), S_HEAD, 'token'); }

/**
 * v86：已經分享出去的票券，事後改「我也留著」（分享當下忘了勾、或想改回只給對方）。
 * ids = 要改的幾張（只收這條連結裡真的有的）；keep = true 加進「我也留著」、false 拿掉。
 * 連結已經收回就不能改（那時票券已經回到自己手上了）。回傳 { keep: 這條連結現在「我也留著」的 id 清單, changed: 這次真的改了幾張 }。
 */
function setShareKeep_(token, ids, keep) {
  ids = (Array.isArray(ids) ? ids : []).map(String).filter(Boolean);
  if (!ids.length) throw new Error('沒有選任何票券');
  if (ids.length > MAX_SHARE) throw new Error('一次最多改 ' + MAX_SHARE + ' 張');
  return withLock_(function () {
    var s = shareRow_(String(token || ''));
    if (!s) throw new Error('找不到這個分享連結');
    if (s.revoked === 'y') throw new Error('這個連結已經收回了');
    var inLink = s.ticketIds ? s.ticketIds.split(',') : [];
    var cur = s.keepIds ? s.keepIds.split(',').filter(Boolean) : [];
    var changed = 0;
    ids.forEach(function (id) {
      if (inLink.indexOf(id) < 0) return;
      var at = cur.indexOf(id);
      if (keep && at < 0) { cur.push(id); changed++; }
      else if (!keep && at >= 0) { cur.splice(at, 1); changed++; }
    });
    if (changed) sSh_().getRange(s._row, S_HEAD.indexOf('keepIds') + 1).setNumberFormat('@').setValue(cur.join(','));
    return { keep: cur, changed: changed };
  });
}

/**
 * 哪些票券現在正分享在外面。
 * 以 Shares 表為唯一來源即時算出來，不在票券上另外存一份 ——
 * 這樣收回或連結過期之後，標記會自己消失，不會有對不起來的狀況。
 */
function activeShareMap_() {
  var today = today_();
  var map = {};
  shares_().forEach(function (s) {
    if (s.revoked === 'y') return;
    if (s.expiresAt && s.expiresAt < today) return;
    var ids = s.ticketIds ? s.ticketIds.split(',') : [];
    var keep = s.keepIds ? s.keepIds.split(',') : [];
    ids.forEach(function (id) {
      var prev = map[id];
      if (!prev || s.createdAt >= prev.sharedAt) {
        map[id] = {
          sharedTo: s.to || '', sharedAt: s.createdAt,
          shareToken: s.token, shareExpires: s.expiresAt,
          opens: +s.openCount || 0, lastOpen: s.lastOpenedAt || '',
          keep: keep.indexOf(id) >= 0        // v76：分享出去、自己這邊也留著
        };
      }
    });
  });
  return map;
}

function listTickets_() {
  var shared = activeShareMap_();
  return tickets_().map(function (t) { return ticketView_(t, shared[t.id]); });
}

/**
 * 用 id 找一張票券：只讀 id 那一欄，再讀命中的那一列。
 * 以前是整張表（含每張的辨識原文）讀進來再找，票券越多越慢。
 */
function findTicket_(id) {
  id = String(id || '');
  if (!id) return null;
  var sh = tSh_(), last = sh.getLastRow();
  if (last < 2) return null;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) !== id) continue;
    var row = i + 2;
    var v = sh.getRange(row, 1, 1, T_HEAD.length).getValues()[0];
    var o = { _row: row };
    for (var j = 0; j < T_HEAD.length; j++) o[T_HEAD[j]] = str_(v[j]);
    return o;
  }
  return null;
}

function rowFrom_(t, head) {
  return head.map(function (h) { return t[h] === undefined ? '' : String(t[h]); });
}

function saveTicket_(t) {
  if (!t) throw new Error('沒有票券內容');
  t.id = t.id || ('T' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6));
  t.createdAt = t.createdAt || today_();
  t.status = t.status || 'active';
  t.fingerprint = fingerprint_(t);
  withLock_(function () {
    var sh = tSh_();
    // 先把新的那一列設成文字格式再寫入：appendRow 會照「使用者輸入」解析，
    // 16 位數的序號會被當成數字，顯示成 7.3E+15，尾數就救不回來了
    var row = sh.getLastRow() + 1;
    sh.getRange(row, 1, 1, T_HEAD.length).setNumberFormat('@').setValues([rowFrom_(t, T_HEAD)]);
  });
  return t;
}

function updateTicket_(id, patch, found) {
  var fresh = LOCK_DEPTH_ > 0;          // 只有在鎖裡面找到的列號才可信，否則重新找一次
  return withLock_(function () {
    var t = (fresh && found) || findTicket_(id);
    if (!t) throw new Error('找不到這張票券');
    var merged = {};
    T_HEAD.forEach(function (h) { merged[h] = t[h]; });
    Object.keys(patch || {}).forEach(function (k) {
      if (T_HEAD.indexOf(k) >= 0) merged[k] = patch[k];
    });
    // 狀態被改了但沒說是從哪裡改的 → 就是你自己在 App 裡改的，分享連結的復原權就此失效
    if (patch && patch.status !== undefined && patch.usedVia === undefined) merged.usedVia = '';
    // v88 用掉的時間：變成已使用的那一刻記下來（App 有帶 usedTime 就用它——按下去的時間，離線補送時才準）；改回未使用就清掉；已經是已使用又沒帶新的就維持原樣
    if (patch && patch.status !== undefined) {
      if (merged.status === 'used') {
        if (validClock_(patch.usedTime)) merged.usedTime = String(patch.usedTime);
        else if (t.status !== 'used') merged.usedTime = clock_();
      } else {
        merged.usedTime = '';
      }
    }
    merged.fingerprint = fingerprint_(merged);
    tSh_().getRange(t._row, 1, 1, T_HEAD.length).setNumberFormat('@').setValues([rowFrom_(merged, T_HEAD)]);
    return merged;
  });
}

function removeTicket_(id) {
  var t = withLock_(function () {
    var x = findTicket_(id);
    if (!x) throw new Error('找不到這張票券');
    tSh_().deleteRow(x._row);
    return x;
  });
  // 只丟掉這個程式自己存的截圖；從你原本資料夾連過來的檔案絕對不動（放在鎖外面，不拖住別人）
  if (t.fileId && t.source !== 'drive') {
    try { DriveApp.getFileById(t.fileId).setTrashed(true); } catch (e) {}
  }
  return {};
}

/**
 * 'code' = 有抓到夠長的序號，兩張指紋一樣幾乎可以確定是同一張券。
 * 'weak' = 只能用商家＋券名＋到期日比，同一種券的不同序號也會長一樣，不可盡信。
 */
function fingerprintKind_(t) {
  return String(t.code || '').replace(/[^0-9A-Za-z]/g, '').length >= 6 ? 'code' : 'weak';
}

function fingerprint_(t) {
  var clean = String(t.code || '').replace(/[^0-9A-Za-z]/g, '').toUpperCase();
  var base = clean.length >= 6
    ? clean
    : [t.merchant || '', t.title || '', t.expiry || ''].join('|');
  if (!base.replace(/\|/g, '')) return '';
  var d = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, base, Utilities.Charset.UTF_8);
  return Utilities.base64Encode(d).replace(/[^A-Za-z0-9]/g, '').slice(0, 16);
}

function today_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/** v88 現在幾點幾分幾秒（HH:mm:ss，跟 today_() 同一個時區） */
function clock_(d) {
  return Utilities.formatDate(d || new Date(), Session.getScriptTimeZone(), 'HH:mm:ss');
}
function validClock_(s) {
  return /^([01]\d|2[0-3]):[0-5]\d:[0-5]\d$/.test(String(s == null ? '' : s));
}
/** v88 從 usedVia（「連結碼@2026-10-08T06:30:00.000Z」）推回用掉的時間；沒有或讀不懂就回空字串 */
function usedClockFromVia_(via) {
  var v = String(via || ''), i = v.indexOf('@');
  if (i < 0) return '';
  var d = new Date(v.slice(i + 1));
  return isNaN(d.getTime()) ? '' : clock_(d);
}

function randomToken_(n) {
  var abc = 'abcdefghijkmnpqrstuvwxyz23456789';
  var out = '';
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    String(Math.random()) + new Date().getTime() + Utilities.getUuid());
  for (var i = 0; i < n; i++) out += abc[Math.abs(bytes[i % bytes.length]) % abc.length];
  return out;
}


/* ══════════════════ 截圖 + 辨識 ══════════════════ */

function parse_(b64, mime, filename) {
  if (!b64) throw new Error('沒有收到圖片');
  mime = mime || 'image/jpeg';

  // 每一段花多久會記在「執行記錄」裡，匯入覺得慢的時候去那邊看是卡在哪
  var t0 = new Date().getTime(), tDrive = 0, tAi = 0, tDup = 0;

  var fileId = '';
  try {
    var folder = DriveApp.getFolderById(P_('FOLDER_ID'));
    var name = String(filename || 'ticket').replace(/[\\/:*?"<>|]/g, '_');
    if (!/\.(jpe?g|png|webp)$/i.test(name)) name += '.jpg';
    var blob = Utilities.newBlob(Utilities.base64Decode(b64), mime, today_() + '-' + name);
    fileId = folder.createFile(blob).getId();
  } catch (e) {
    throw new Error('存不進雲端硬碟：' + e.message);
  }
  tDrive = new Date().getTime();

  var draft;
  try {
    draft = parseTicket_(b64, mime);
  } catch (e) {
    // 辨識失敗就把剛存的圖丟垃圾桶，重試時才不會留下一堆重複的圖
    try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e2) {}
    throw e;
  }
  if (draft.renamed) draft.note = renamedFlag_(draft);     // v67：套了改名規則，備註記原名（App 確認畫面也看得到）
  draft.fingerprint = fingerprint_(draft);
  tAi = new Date().getTime();

  var dup = false, dupOf = null;
  if (draft.fingerprint) {
    var all = tickets_(), strong = fingerprintKind_(draft) === 'code';
    for (var i = 0; i < all.length; i++) {
      if (all[i].fingerprint && all[i].fingerprint === draft.fingerprint) {
        dup = true;
        // v76：對上了哪一張、用掉了沒（有好幾張對上就挑用掉的那張，匯入畫面才會提醒「這張已經用掉了」）
        if (!dupOf || (all[i].status === 'used' && dupOf.status !== 'used')) {
          dupOf = { id: all[i].id, status: all[i].status || 'active', usedAt: all[i].usedAt || '', usedBy: all[i].usedBy || '',
                    usedTime: all[i].status === 'used' ? (all[i].usedTime || usedClockFromVia_(all[i].usedVia)) : '',
                    merchant: all[i].merchant || '', title: all[i].title || '', strong: strong };
        }
      }
    }
  }
  tDup = new Date().getTime();

  Logger.log('匯入計時（秒）：雲端硬碟 ' + ((tDrive - t0) / 1000).toFixed(1) +
             '　辨識 ' + ((tAi - tDrive) / 1000).toFixed(1) +
             '　查重複 ' + ((tDup - tAi) / 1000).toFixed(1) +
             '　合計 ' + ((tDup - t0) / 1000).toFixed(1) +
             '　圖片 ' + Math.round(b64.length * 0.75 / 1024) + ' KB');

  return { draft: draft, fileId: fileId, duplicate: dup, dupOf: dupOf };
}

/**
 * 辨識引擎的順序：設了哪支金鑰就用哪個。
 *   ANTHROPIC_KEY → Claude（要付費，約每張 NT$0.1，複雜版面判斷較穩）
 *   GEMINI_KEY    → Gemini（有免費額度）
 *   VISION_KEY    → Vision API 取字 + 規則判斷（最陽春的備援）
 * 兩支都設的話，把 ENGINE 屬性設成 claude 或 gemini 可以指定要用哪個。
 */
function parseTicket_(b64, mime) {
  // v67：辨識完馬上套改名規則（App 匯入和雲端硬碟上傳區匯入都走這裡，查重複之前就改好，指紋才會跟已經改過名的票券對得上）
  // v69：標題明寫「餘額型」「分次使用」的，順便標成餘額型（改名之後也認得，原名一起看）
  return flagBalance_(applyNameRules_(parseTicketRaw_(b64, mime)));
}

function parseTicketRaw_(b64, mime) {
  var engine = P_('ENGINE').toLowerCase();
  if (engine === 'claude' && P_('ANTHROPIC_KEY')) return claudeParse_(b64, mime);
  if (engine === 'gemini' && P_('GEMINI_KEY'))    return geminiParse_(b64, mime);
  if (P_('ANTHROPIC_KEY')) return claudeParse_(b64, mime);
  if (P_('GEMINI_KEY'))    return geminiParse_(b64, mime);
  if (P_('VISION_KEY'))    return rulesParse_(visionOcr_(b64));
  throw new Error('還沒設定 ANTHROPIC_KEY 或 GEMINI_KEY，沒辦法辨識');
}

function engineName_() {
  var e = P_('ENGINE').toLowerCase();
  if (e === 'claude' && P_('ANTHROPIC_KEY')) return 'Claude';
  if (e === 'gemini' && P_('GEMINI_KEY'))    return 'Gemini';
  if (P_('ANTHROPIC_KEY')) return 'Claude';
  if (P_('GEMINI_KEY'))    return 'Gemini';
  if (P_('VISION_KEY'))    return 'Vision + 規則';
  return '（尚未設定）';
}

/* ══════════════════ 改名規則（v67） ══════════════════
 *
 * 「NameRules」分頁放「舊名字 → 新名字」，匯入票券時自動套用（App 匯入和雲端硬碟自動匯入都會）：
 *   fromMerchant | fromTitle | toMerchant | toTitle | note
 *   fromTitle 空白 ＝ 這家商家的任何券名（例如 7-ELEVEN → 7-11、遠東SOGO百貨 → SOGO）
 *   toMerchant／toTitle 空白 ＝ 那一項維持原樣
 *
 * 比對順序（越前面越有把握）：
 *   1. 一模一樣
 *   2. 同一家，券名只差空格、括號、商家名、「好禮即享券」這類固定字、「鷄／雞」這類異體字
 *   3. 被截斷的名字（結尾「...」）：拿前面那段去對
 *   4. 差一個字（辨識錯字，例如「喜客券 i禮鑽適用」「i禮讚適用」）；差的是數字、大中小、冰熱的不算
 * 都對不上的，只做通用整理：商家別名統一、去掉券名裡重複的商家名和固定字。
 * 分頁是空的就什麼都不做；指令碼屬性 NAME_RULES_OFF 設成 1 可以整個關掉。
 *
 * 種類維持原樣：改過名的票券，備註會寫「已改名（原：…）」，App 判斷種類時連備註一起看，
 * 所以改名不會讓種類跟著掉（例如「中薯」「豬滿蛋+冰美」沒有關鍵字，但備註裡的原名有）。
 */
var R_SHEET = 'NameRules';
var R_HEAD = ['fromMerchant', 'fromTitle', 'toMerchant', 'toTitle', 'note'];
var R_MAX = 2000;
var FIXED_WORDS_ = ['好禮即享券', '即享券', '兌換券'];    // 2026/10 整理名字時，每一張都拿掉了這幾個固定字（長的在前）
var SEMANTIC_CHARS_ = /[0-9大中小冰熱溫lms]/;            // 漏掉這些字的不算辨識錯字：大杯／中杯、冰／熱、數字都是不同的東西
var CONFUSABLE_ = ['鑽讚贊', '熱熟', '未末', '己已巳', '土士', '千干于', '日曰', '人入'];   // 長得像、辨識常認錯的字

/** 比對用的寫法：全形轉半形、小寫、異體字統一、空白壓成一個 */
function nameNorm_(s) {
  s = (s === undefined || s === null) ? '' : String(s);
  try { s = s.normalize('NFKC'); } catch (e) {}
  return s.toLowerCase().replace(/鷄/g, '雞').replace(/臺/g, '台').replace(/裏/g, '裡').replace(/麪/g, '麵').replace(/着/g, '著')
          .replace(/\s+/g, ' ').trim();
}

/** 去掉空白、括號、分隔符號，只剩內容 */
function stripSep_(s) {
  return String(s).replace(/[\s·・•:;_\-—–~,.。、\/\\|*+()\[\]【】《》「」『』<>"'`!?]/g, '');
}

/** 商家的比對鍵：別名統一之後，去掉符號（「Mr.Wish」「Mr. Wish」會一樣） */
function mkey_(merchant, rules) {
  var m = nameNorm_(merchant), c = rules.alias[m];
  return stripSep_(c ? nameNorm_(c) : m);
}

/** 這家商家在券名裡可能出現的寫法（比對寫法，長的在前） */
function merchantNames_(merchant, rules) {
  var m = nameNorm_(merchant), c = rules.alias[m], canon = c ? nameNorm_(c) : m;
  var set = {}, out = [];
  if (m) set[m] = 1;
  if (canon) set[canon] = 1;
  (rules.groups[canon] || []).forEach(function (n) { set[n] = 1; });
  Object.keys(set).forEach(function (n) { if (n.length >= 2) out.push(n); });
  out.sort(function (a, b) { return b.length - a.length; });
  return out;
}

/** 商家的各種原字樣（整理券名時要照原字樣去找），長的在前 */
function rawNames_(merchantRaw, merchantCanon, rules) {
  var out = [], seen = {};
  function add(n) {
    n = String(n || '').trim();
    var k = nameNorm_(n);
    if (n && !seen[k]) { seen[k] = 1; out.push(n); }
  }
  add(merchantRaw); add(merchantCanon);
  (rules.groupRaw[nameNorm_(merchantCanon)] || []).forEach(add);
  out.sort(function (a, b) { return b.length - a.length; });
  return out;
}

/** 券名的比對鍵：去掉商家名、開頭的括號標籤、截斷符號、符號、固定字 */
function titleKey_(title, names) {
  var t = nameNorm_(title), i, cut = false;
  for (i = 0; i < names.length; i++) t = t.split(names[i]).join(' ');
  if (/(\.{2,}|…)\s*$/.test(t)) { cut = true; t = t.replace(/(\.{2,}|…)\s*$/, ''); }
  t = t.replace(/^(?:\s*(?:\[[^\]]*\]|【[^】]*】))+/, '');
  t = stripSep_(t);
  for (i = 0; i < FIXED_WORDS_.length; i++) t = t.split(FIXED_WORDS_[i]).join('');
  return { key: t, cut: cut };
}

/** 把讀進來的規則整理成好比對的樣子 */
function prepareRules_(list) {
  var rules = { list: list || [], alias: {}, groups: {}, groupRaw: {}, items: [] };
  var gset = {};
  rules.list.forEach(function (r) {
    if (r.ft || !r.fm || !r.tm) return;                 // 只有商家的規則 ＝ 別名
    var f = nameNorm_(r.fm), c = nameNorm_(r.tm);
    if (f === c) return;
    rules.alias[f] = r.tm;
    gset[c] = gset[c] || { n: {}, raw: {} };
    gset[c].n[f] = 1; gset[c].n[c] = 1; gset[c].raw[r.fm] = 1; gset[c].raw[r.tm] = 1;
  });
  Object.keys(gset).forEach(function (c) { rules.groups[c] = Object.keys(gset[c].n); rules.groupRaw[c] = Object.keys(gset[c].raw); });
  rules.list.forEach(function (r) {
    if (!r.ft) return;
    var tk = titleKey_(r.ft, merchantNames_(r.fm, rules));
    rules.items.push({
      r: r, k1: nameNorm_(r.fm) + '|' + nameNorm_(r.ft), mkey: mkey_(r.fm, rules),
      key: tk.key, cut: tk.cut, digits: (tk.key.match(/\d+/g) || []).join(',')
    });
  });
  return rules;
}

/**
 * 兩段字「只差一個字」而且看得出是辨識錯字 → 1，否則 0。
 * 一般的換字不算（紅茶／綠茶、低脂／全脂、大薯／小薯都是不同的東西），只有這兩種算：
 *   ・多一個字或少一個字（漏字），而且那個字不是數字、大中小、冰熱
 *   ・換成長得像的字（CONFUSABLE_ 裡列的，例如 鑽／讚）
 */
function oneEdit_(a, b) {
  var la = a.length, lb = b.length;
  if (Math.abs(la - lb) > 1) return 0;
  var i = 0;
  while (i < la && i < lb && a.charAt(i) === b.charAt(i)) i++;
  var ca, cb, g;
  if (la === lb) {
    if (a.substr(i + 1) !== b.substr(i + 1)) return 0;
    ca = a.charAt(i); cb = b.charAt(i);
    for (g = 0; g < CONFUSABLE_.length; g++) {
      if (CONFUSABLE_[g].indexOf(ca) >= 0 && CONFUSABLE_[g].indexOf(cb) >= 0) return 1;
    }
    return 0;
  }
  if (la > lb) {
    if (a.substr(i + 1) !== b.substr(i)) return 0;
    return SEMANTIC_CHARS_.test(a.charAt(i)) ? 0 : 1;
  }
  if (b.substr(i + 1) !== a.substr(i)) return 0;
  return SEMANTIC_CHARS_.test(b.charAt(i)) ? 0 : 1;
}

/** 幾條規則都對上時：結果一樣才算數（互相矛盾就不動它）；一模一樣的以最後一條為準 */
function pickRule_(hit, tier) {
  if (!hit.length) return null;
  if (tier === 'exact') return { r: hit[hit.length - 1].r, tier: tier };
  var a = hit[0].r;
  for (var i = 1; i < hit.length; i++) {
    if (hit[i].r.tm !== a.tm || hit[i].r.tt !== a.tt) return null;
  }
  return { r: a, tier: tier };
}

function findRule_(merchant, title, rules) {
  var items = rules.items, i, hit = [];
  if (!items.length || !title) return null;
  var k1 = nameNorm_(merchant) + '|' + nameNorm_(title);
  for (i = 0; i < items.length; i++) if (items[i].k1 === k1) hit.push(items[i]);
  if (hit.length) return pickRule_(hit, 'exact');

  var mk = mkey_(merchant, rules), tk = titleKey_(title, merchantNames_(merchant, rules));
  if (tk.key.length < 2) return null;
  for (i = 0; i < items.length; i++) if (items[i].mkey === mk && items[i].key === tk.key) hit.push(items[i]);
  if (hit.length) return pickRule_(hit, 'norm');

  if (tk.key.length >= 6) {                               // 被截斷的名字：拿前面那段去對
    for (i = 0; i < items.length; i++) {
      var it = items[i];
      if (it.mkey !== mk) continue;
      if (tk.cut ? it.key.indexOf(tk.key) === 0 : (it.cut && it.key.length >= 6 && tk.key.indexOf(it.key) === 0)) hit.push(it);
    }
    if (hit.length) return pickRule_(hit, 'cut');
  }
  if (tk.key.length >= 8) {                               // 差一個字：辨識錯字
    var dg = (tk.key.match(/\d+/g) || []).join(',');
    for (i = 0; i < items.length; i++) {
      if (items[i].mkey === mk && items[i].digits === dg && oneEdit_(tk.key, items[i].key) === 1) hit.push(items[i]);
    }
    if (hit.length) return pickRule_(hit, 'near');
  }
  return null;
}

/** 沒有規則對上時的通用整理：去掉券名裡重複的商家名（含括號標籤）和固定字 */
function cleanTitle_(merchantRaw, merchantCanon, title, rules) {
  var t = String(title || '').trim();
  if (!t) return t;
  var raws = rawNames_(merchantRaw, merchantCanon, rules), i, j;
  var normNames = raws.map(nameNorm_).filter(function (n) { return n.length >= 2; });
  var step = t;
  // 1) 括號標籤裡有商家名的整個拿掉（[麥當勞·7月限定麥心意] 連活動名一起）
  step = step.replace(/\[[^\]]*\]|【[^】]*】/g, function (tag) {
    var inner = nameNorm_(tag);
    for (j = 0; j < normNames.length; j++) if (inner.indexOf(normNames[j]) >= 0) return ' ';
    return tag;
  });
  // 2) 開頭的商家名（後面接空格、符號，或商家名夠長不會誤砍）
  var again = true, guard = 0;
  while (again && guard++ < 4) {
    again = false;
    var trimmed = step.replace(/^[\s\-_:：·・]+/, '');
    for (i = 0; i < raws.length; i++) {
      var nm = raws[i];
      if (!nm || trimmed.length <= nm.length) continue;
      if (nameNorm_(trimmed.substr(0, nm.length)) !== nameNorm_(nm)) continue;
      var rest = trimmed.substr(nm.length);
      if (!/^[\s\-_:：·・\[【(（]/.test(rest) && stripSep_(nameNorm_(nm)).length < 3) continue;
      var rest2 = rest.replace(/^[\s\-_:：·・]+/, '');
      if (rest2.length < 2) continue;
      step = rest2; again = true; break;
    }
  }
  // 3) 固定字
  var fx = step;
  for (i = 0; i < FIXED_WORDS_.length; i++) fx = fx.split(FIXED_WORDS_[i]).join('');
  if (fx.replace(/[\s\-_:：·・()（）]/g, '').length >= 2) step = fx;
  step = step.replace(/\s+/g, ' ').replace(/^[\s\-_:：·・]+|[\s\-_:：·・]+$/g, '');
  return step || t;
}

/** 算出這個商家＋券名該改成什麼；不用改就回 null */
function renameName_(merchant, title, rules) {
  merchant = String(merchant || '').trim();
  title = String(title || '').trim();
  if (!merchant && !title) return null;
  var canonM = rules.alias[nameNorm_(merchant)] || merchant;
  var hit = findRule_(merchant, title, rules), outM, outT, tier;
  if (hit) {
    outM = hit.r.tm || canonM;
    outT = hit.r.tt || title;
    tier = hit.tier;
  } else {
    outM = canonM;
    outT = cleanTitle_(merchant, canonM, title, rules);
    tier = outT !== title ? 'clean' : 'alias';
  }
  if (outM === merchant && outT === title) return null;
  return { merchant: outM, title: outT, tier: tier };
}

var RULES_ = null;
/** 讀「改名規則」分頁（同一次請求只讀一次）。分頁不存在或空的 → 沒有規則 → 這個功能不動作 */
function loadRules_() {
  if (RULES_) return RULES_;
  var list = [];
  try {
    var sh = ss_().getSheetByName(R_SHEET), last = sh ? sh.getLastRow() : 0;
    if (last >= 2) {
      sh.getRange(2, 1, Math.min(last - 1, R_MAX), R_HEAD.length).getValues().forEach(function (v) {
        var fm = str_(v[0]).trim(), ft = str_(v[1]).trim(), tm = str_(v[2]).trim(), tt = str_(v[3]).trim();
        if ((fm || ft) && (tm || tt)) list.push({ fm: fm, ft: ft, tm: tm, tt: tt });
      });
    }
  } catch (e) { Logger.log('讀改名規則失敗：' + e.message); }
  RULES_ = prepareRules_(list);
  return RULES_;
}

/** 匯入時套用：辨識出來的 draft 直接改，原名記在 draft.renamed。出任何問題都原樣放行，不能擋住匯入 */
function applyNameRules_(d) {
  try {
    if (!d || P_('NAME_RULES_OFF') === '1') return d;
    var rules = loadRules_();
    if (!rules.list.length) return d;
    var om = String(d.merchant || '').trim(), ot = String(d.title || '').trim();
    var r = renameName_(om, ot, rules);
    if (!r) return d;
    d.renamed = { merchant: om, title: ot, tier: r.tier };
    d.merchant = r.merchant;
    d.title = r.title;
  } catch (e) { Logger.log('改名規則沒套用：' + e.message); }
  return d;
}

/** 寫進備註的那句（也讓 App 判斷種類時看得到原名） */
function renamedFlag_(d) {
  var r = d && d.renamed;
  if (!r) return '';
  return '已改名（原：' + (r.merchant && r.merchant !== d.merchant ? r.merchant + ' ' : '') + r.title + '）';
}

/** 設定畫面用：有幾條規則、規則分頁的網址 */
function rulesInfo_() {
  var sh = ss_().getSheetByName(R_SHEET), url = '';
  try { url = ss_().getUrl() + (sh ? '#gid=' + sh.getSheetId() : ''); } catch (e) {}
  return { count: loadRules_().list.length, url: url, off: P_('NAME_RULES_OFF') === '1' };
}

function csvLine_(ln) {
  var out = [], cur = '', q = false, i, c;
  for (i = 0; i < ln.length; i++) {
    c = ln.charAt(i);
    if (q) {
      if (c === '"') { if (ln.charAt(i + 1) === '"') { cur += '"'; i++; } else q = false; }
      else cur += c;
    } else if (c === '"') q = true;
    else if (c === ',') { out.push(cur); cur = ''; }
    else cur += c;
  }
  out.push(cur);
  return out;
}

/** 貼上一批規則（每行：舊商家、舊券名、新商家、新券名［、備註］，Tab 或逗號隔開）。舊名字一樣的更新、沒有的新增 */
function rulesImport_(text) {
  var cells = [];
  String(text || '').split(/\r?\n/).forEach(function (ln) {
    if (!ln.trim()) return;
    var c = (ln.indexOf('\t') >= 0 ? ln.split('\t') : csvLine_(ln)).map(function (x) { return String(x === undefined ? '' : x).trim(); });
    if (/^(舊商家|fromMerchant)$/i.test(c[0])) return;          // 標題列
    cells.push(c);
  });
  if (!cells.length) throw new Error('沒有讀到規則。每行一條：舊商家、舊券名、新商家、新券名，用 Tab 或逗號隔開');
  if (cells.length > R_MAX) throw new Error('一次最多 ' + R_MAX + ' 條');
  return withLock_(function () {
    var ss = ss_();
    var sh = ss.getSheetByName(R_SHEET) || ensureSheet_(ss, R_SHEET, R_HEAD);
    var last = sh.getLastRow();
    var cur = last >= 2 ? sh.getRange(2, 1, last - 1, R_HEAD.length).getValues().map(function (r) {
      return r.map(function (x) { return str_(x); });
    }) : [];
    var index = {};
    cur.forEach(function (r, i) { index[nameNorm_(r[0]) + '|' + nameNorm_(r[1])] = i; });
    var added = 0, updated = 0, same = 0, skipped = 0;
    cells.forEach(function (c) {
      var fm = c[0] || '', ft = c[1] || '', tm = c[2] || '', tt = c[3] || '', note = c[4] || '';
      if ((!fm && !ft) || (!tm && !tt) || fm.length > 200 || ft.length > 200 || tm.length > 200 || tt.length > 200) { skipped++; return; }
      if ((!tm || tm === fm) && (!tt || tt === ft)) { same++; return; }       // 新舊一樣，不用記
      var key = nameNorm_(fm) + '|' + nameNorm_(ft);
      if (key in index) {
        var row = cur[index[key]];
        if (row[2] === tm && row[3] === tt) { same++; return; }
        row[2] = tm; row[3] = tt; if (note) row[4] = note;
        updated++;
      } else {
        index[key] = cur.length;
        cur.push([fm, ft, tm, tt, note]);
        added++;
      }
    });
    if (cur.length > R_MAX) throw new Error('規則太多了（上限 ' + R_MAX + ' 條）');
    if (added || updated) sh.getRange(2, 1, cur.length, R_HEAD.length).setNumberFormat('@').setValues(cur);
    RULES_ = null;
    return { added: added, updated: updated, same: same, skipped: skipped, count: cur.length };
  });
}

/** 把模型回來的東西整理成固定形狀。 */
function tidyDraft_(d) {
  d = d || {};
  d.expiry = normDate_(d.expiry);
  ['merchant', 'title', 'value', 'code', 'codeType', 'rawText', 'confidence'].forEach(function (k) {
    d[k] = String(d[k] || '');
  });
  d.code = fixSci_(d.code, d.rawText);
  if (noExpiry_(d.rawText)) d.expiry = '';        // 「… ～ 無效期」這種其實是沒有到期日
  if (isSci_(d.code)) { d.code = ''; d.confidence = 'low'; }   // 救不回來就寧可空著，讓你自己補
  // v70：票面上明寫的餘額（要有證據那行字才收，見 balanceChecked_）
  d.balance = balanceChecked_(d.balance, d.balanceText);
  d.balanceText = d.balance ? String(d.balanceText).slice(0, 40) : '';
  return d;
}

/**
 * 票面上寫「無效期／無期限／長期有效」時，代表沒有到期日。
 * 這種券常見寫法是「兌換期間至：2026.09.24 ～ 無效期」，AI 很容易把前面的開始日當成到期日。
 */
function noExpiry_(rawText) {
  var t = String(rawText || '').replace(/\s+/g, '');
  return /(無效期|無期限|不限期|無使用期限|無到期日|永久有效|長期有效)/.test(t);
}

/** 「7.3E+15」這種被縮寫掉的序號。長序號一旦變成科學記號，尾數就全毀了。 */
function isSci_(v) {
  return /^\s*\d+(\.\d+)?\s*[eE]\s*[+-]?\s*\d+\s*$/.test(String(v || ''));
}

/** 序號變成科學記號時，回頭從 rawText 撈那串完整數字。撈不到就原樣退回。 */
function fixSci_(code, rawText) {
  code = String(code || '');
  if (!isSci_(code)) return code;
  var runs = String(rawText || '').match(/\d[\d\- ]{7,}\d/g) || [];
  var best = '';
  runs.forEach(function (r) {
    var digits = r.replace(/[^0-9]/g, '');
    if (digits.length > best.replace(/[^0-9]/g, '').length) best = r.trim();
  });
  return best || code;
}

var PROMPT = [
  '你在幫台灣的使用者整理票券截圖。看這張圖，把「一張票券」的資訊抽出來，只回 JSON。',
  '',
  'merchant：發行或可兌換的商家，例如「全家便利商店」「星巴克」「新光三越」。找不到就空字串。',
  'title：這張券能換什麼，盡量照抄畫面上的原文，例如「中杯拿鐵兌換券」「全館現金抵用券」。',
  'value：面額或數量，例如「NT$500」「1 杯」「2 張」。沒有就空字串。',
  'balance：這張券「現在還剩多少元」。只填畫面上明確寫著「餘額」「剩餘金額」「可用金額」「目前金額」「儲值餘額」的那個金額，只寫數字（不要 NT$、$、逗號、元）；',
  '  常在畫面右上角，或寫在面額旁邊。',
  '  · 只寫「面額」「票面金額」「儲值金額」、沒有寫餘額或剩餘的，balance 一律回空字串，不要拿面額當餘額。',
  '  · 有好幾個金額、分不出哪個是餘額，或看不清楚，也回空字串，不要猜。',
  'balanceText：你讀到餘額的那一行原文（例如「餘額：NT$3,987」），最多 30 字；balance 是空字串時也回空字串。',
  'code：兌換序號，或條碼／QR 下面印的那串數字，一個字一個字照抄，保留原本的分隔符號和開頭的 0。',
  '  · 一定要抄完整，長度再長也不能省略、不能四捨五入，',
  '    絕對不可以寫成科學記號（像 7.3E+15 這種就是錯的，要寫 7300002212340590）。',
  '  · 圖上同時有「序號」文字和條碼下面的數字時，以比較完整的那一串為準。',
  '  · 看不清楚就回空字串，不要猜或補數字。',
  'codeType：barcode、qrcode、text 三選一，看不出來就空字串。',
  'expiry：最後一天可以使用的日期，格式 YYYY-MM-DD。判斷規則：',
  '  · 「有效期限」「使用期限」「兌換期限」「使用至」「限…前使用」後面的日期才算；',
  '    「發行日」「購買日」「開立日期」「訂單日期」不是到期日。',
  '  · 民國年要換算成西元：115/03/31 → 2026-03-31，114.12.05 → 2025-12-05。',
  '    年份是三位數或小於 200 的，就是民國年，加 1911。',
  '  · 只寫到月份（例如「2026年10月」「115年10月」）就取該月最後一天。',
  '  · 寫成區間（2026/01/01-2026/03/31）取後面那個；前面那個是開始日，絕對不能當到期日。',
  '  · 區間的後半寫「無效期」「無期限」「不限期」「長期有效」「無使用期限」時，',
  '    代表這張沒有到期日，expiry 一律回空字串，不可以改用前面那個開始日。',
  '    例如「兌換期間至：2026.09.24 ～ 無效期」→ expiry 是空字串。',
  '  · 日期後面接「起」「開始」「發行」的也是開始日，不是到期日。',
  '  · 真的沒寫就回空字串，不要猜。',
  'rawText：只抄跟上面這些欄位有關的那幾行（商家、券名、期限、序號），最多 120 字 —— 不要整張票照抄，那會拖慢速度。',
  'confidence：high / medium / low，代表你對 expiry 和 merchant 的把握。',
  '',
  '只輸出 JSON，不要有其他說明文字。'
].join('\n');

/** Gemini 429 的錯誤裡會寫「"retryDelay": "23s"」，照它說的等（上限 20 秒）。 */
function retryDelay_(text) {
  var m = /"retryDelay"\s*:\s*"(\d+(?:\.\d+)?)s"/.exec(text || '');
  return m ? Math.min(Math.ceil(parseFloat(m[1]) * 1000) + 500, 20000) : 0;
}

function geminiParse_(b64, mime) {
  var body = {
    contents: [{
      role: 'user',
      parts: [
        { inlineData: { mimeType: mime, data: b64 } },
        { text: PROMPT }
      ]
    }],
    generationConfig: {
      temperature: 0,
      responseMimeType: 'application/json',
      responseSchema: {
        type: 'OBJECT',
        properties: {
          merchant:   { type: 'STRING' },
          title:      { type: 'STRING' },
          value:      { type: 'STRING' },
          code:       { type: 'STRING' },
          codeType:   { type: 'STRING' },
          expiry:     { type: 'STRING' },
          rawText:    { type: 'STRING' },
          balance:    { type: 'STRING' },
          balanceText: { type: 'STRING' },
          confidence: { type: 'STRING' }
        },
        required: ['merchant', 'title', 'expiry', 'confidence']
      }
    }
  };

  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + MODEL +
            ':generateContent?key=' + encodeURIComponent(P_('GEMINI_KEY'));

  // 連續匯入時 Gemini 常會回 429（每分鐘次數到頂）或 503（模型太忙）。
  // 這兩種等幾秒通常就好，所以自己重試，最多等約 50 秒，免得前端等到逾時。
  var code, text, waited = 0;
  for (var attempt = 1; ; attempt++) {
    var res = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify(body), muteHttpExceptions: true
    });
    code = res.getResponseCode();
    text = res.getContentText();
    if (code === 200) break;

    var retryable = (code === 429 || code === 500 || code === 502 || code === 503 || code === 504);
    if (code === 429 && /PerDay/i.test(text)) retryable = false;   // 今天的額度用完，等也沒用
    var wait = retryDelay_(text) || Math.min(4000 * attempt, 15000);
    if (!retryable || attempt >= 4 || waited + wait > 50000) break;
    Logger.log('Gemini 回 ' + code + '，第 ' + attempt + ' 次，等 ' + (wait / 1000) + ' 秒再試');
    Utilities.sleep(wait);
    waited += wait;
  }

  if (code !== 200) {
    if (code === 404) throw new Error('Gemini 說沒有「' + MODEL + '」這個模型，請到 Code.gs 最上面換一個目前還在的 flash 模型');
    if (code === 429 && /PerDay/i.test(text)) throw new Error('Gemini 今天的免費額度用完了，明天會自動恢復');
    if (code === 429) throw new Error('Gemini 每分鐘次數到頂，等一分鐘再按「重試」');
    if (code === 503 || code === 500) throw new Error('Gemini 現在太忙，等一下再按「重試」');
    throw new Error('Gemini 回 ' + code + '：' + text.replace(/\s+/g, ' ').slice(0, 200));
  }

  var j = JSON.parse(text);
  var out = j.candidates && j.candidates[0] && j.candidates[0].content &&
            j.candidates[0].content.parts && j.candidates[0].content.parts[0].text;
  if (!out) throw new Error('Gemini 沒有回傳內容');

  return tidyDraft_(JSON.parse(out));
}

/* ── Claude（Anthropic API）───────────────────────────────
   金鑰去 https://console.claude.com 開，跟 Claude 網頁版的訂閱是分開計費的。
   Haiku 便宜又夠用；想更準可以換成 sonnet 系列，價格大約是三倍。
   模型代號若失效，去 https://platform.claude.com/docs 查現行的 ID。 */
var CLAUDE_MODEL = 'claude-haiku-4-5';

function claudeParse_(b64, mime) {
  var body = {
    model: CLAUDE_MODEL,
    max_tokens: 1200,
    temperature: 0,
    system: '你只輸出一個 JSON 物件，不要有 markdown 圍欄、不要有任何說明文字。',
    messages: [
      {
        role: 'user',
        content: [
          { type: 'image', source: { type: 'base64', media_type: mime, data: b64 } },
          { type: 'text', text: PROMPT + '\n\nJSON 欄位固定為：merchant、title、value、code、codeType、expiry、rawText、balance、balanceText、confidence。' }
        ]
      },
      { role: 'assistant', content: '{' }   // 預填一個左括號，逼它直接接 JSON
    ]
  };

  var res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': P_('ANTHROPIC_KEY'),
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  var code = res.getResponseCode();
  var text = res.getContentText();
  if (code === 401) throw new Error('Anthropic 說金鑰不對，檢查 ANTHROPIC_KEY');
  if (code === 404) throw new Error('Anthropic 說沒有「' + CLAUDE_MODEL + '」這個模型，請到 Code.gs 換一個現行的 ID');
  if (code === 429) throw new Error('Anthropic 擋流量了，等一下再試');
  if (code !== 200) throw new Error('Anthropic 回 ' + code + '：' + text.slice(0, 200));

  var j = JSON.parse(text);
  var out = j.content && j.content[0] && j.content[0].text;
  if (!out) throw new Error('Claude 沒有回傳內容');

  out = '{' + out;                                  // 補回預填的括號
  out = out.replace(/```json|```/g, '').trim();     // 保險：萬一它還是包了圍欄
  var cut = out.lastIndexOf('}');
  if (cut > 0) out = out.slice(0, cut + 1);

  return tidyDraft_(JSON.parse(out));
}

/* ══════════════════ 從截圖讀餘額（v70） ══════════════════
 *
 *  ① 匯入時：辨識票券順便讀「票面上明寫的餘額」（PROMPT 的 balance／balanceText）。
 *     要有「餘額／剩餘金額…」那行字當證據才收（balanceChecked_）；只寫面額的不算。讀到的話自動標成餘額型。
 *  ② 餘額面板的「用截圖讀餘額」：新拍一張餘額頁面（例如 SOGO、商品卡的查餘額畫面）的截圖，
 *     這裡讀出目前剩餘金額——不存檔、不建票券，只把數字回給畫面，由你確認後才記下來。
 */
var BALANCE_PROMPT = [
  '這是一張可以分次使用的票券（儲值金額、商品卡、電子禮券、好禮即享券…）的截圖，或是查餘額頁面的截圖。',
  '找出「現在還剩多少元」，只回 JSON。',
  '',
  'balance：目前剩餘金額（元），只寫數字（不要 NT$、$、逗號、元）。只能填畫面上明確寫著「餘額」「剩餘金額」「可用金額」「目前金額」「儲值餘額」的那個金額；',
  '  常在畫面右上角，或寫在面額旁邊。',
  '  · 只有「面額」「票面金額」「儲值金額」、沒有寫餘額或剩餘的，balance 一律回空字串，不要拿面額當餘額。',
  '  · 有好幾個金額、分不出哪個是餘額，或看不清楚，也回空字串，不要猜。',
  'evidence：你讀到餘額的那一行原文（例如「餘額：NT$3,987」），最多 30 字；balance 是空字串時回空字串。',
  'face：畫面上寫的面額（元，只寫數字）；沒有就空字串。',
  'confidence：high / medium / low，代表你對 balance 的把握。',
  '',
  '只輸出 JSON，不要有其他說明文字。'
].join('\n');

/** 模型回來的東西整理成固定形狀；沒有證據的餘額一律當作沒讀到 */
function tidyBalanceRead_(r) {
  r = r || {};
  var ev = String(r.evidence || '').slice(0, 60);
  var bal = balanceChecked_(r.balance, ev);
  return { balance: bal, face: cleanAmount_(r.face), evidence: bal ? ev : '', confidence: String(r.confidence || '') };
}

function geminiBalance_(b64, mime) {
  var body = {
    contents: [{ role: 'user', parts: [{ inlineData: { mimeType: mime, data: b64 } }, { text: BALANCE_PROMPT }] }],
    generationConfig: {
      temperature: 0, responseMimeType: 'application/json',
      responseSchema: {
        type: 'OBJECT',
        properties: { balance: { type: 'STRING' }, evidence: { type: 'STRING' }, face: { type: 'STRING' }, confidence: { type: 'STRING' } },
        required: ['balance', 'evidence', 'face', 'confidence']
      }
    }
  };
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + MODEL +
            ':generateContent?key=' + encodeURIComponent(P_('GEMINI_KEY'));
  var code, text, waited = 0;
  for (var attempt = 1; ; attempt++) {
    var res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', payload: JSON.stringify(body), muteHttpExceptions: true });
    code = res.getResponseCode(); text = res.getContentText();
    if (code === 200) break;
    var retryable = (code === 429 || code === 500 || code === 502 || code === 503 || code === 504) && !(code === 429 && /PerDay/i.test(text));
    var wait = retryDelay_(text) || 3000 * attempt;
    if (!retryable || attempt >= 3 || waited + wait > 25000) break;
    Utilities.sleep(wait); waited += wait;
  }
  if (code !== 200) {
    if (code === 404) throw new Error('Gemini 說沒有「' + MODEL + '」這個模型，請到 Code.gs 最上面換一個目前還在的 flash 模型');
    if (code === 429 && /PerDay/i.test(text)) throw new Error('Gemini 今天的免費額度用完了，明天會自動恢復');
    if (code === 429) throw new Error('Gemini 每分鐘次數到頂，等一分鐘再試');
    if (code === 503 || code === 500) throw new Error('Gemini 現在太忙，等一下再試');
    throw new Error('Gemini 回 ' + code + '：' + String(text).replace(/\s+/g, ' ').slice(0, 200));
  }
  var j = JSON.parse(text);
  var out = j.candidates && j.candidates[0] && j.candidates[0].content &&
            j.candidates[0].content.parts && j.candidates[0].content.parts[0].text;
  if (!out) throw new Error('Gemini 沒有回傳內容');
  return JSON.parse(out);
}

function claudeBalance_(b64, mime) {
  var body = {
    model: CLAUDE_MODEL, max_tokens: 400, temperature: 0,
    system: '你只輸出一個 JSON 物件，不要有 markdown 圍欄、不要有任何說明文字。',
    messages: [
      { role: 'user', content: [
        { type: 'image', source: { type: 'base64', media_type: mime, data: b64 } },
        { type: 'text', text: BALANCE_PROMPT + '\n\nJSON 欄位固定為：balance、evidence、face、confidence。' }
      ] },
      { role: 'assistant', content: '{' }
    ]
  };
  var res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json',
    headers: { 'x-api-key': P_('ANTHROPIC_KEY'), 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify(body), muteHttpExceptions: true
  });
  var code = res.getResponseCode(), text = res.getContentText();
  if (code === 401) throw new Error('Anthropic 說金鑰不對，檢查 ANTHROPIC_KEY');
  if (code === 429) throw new Error('Anthropic 擋流量了，等一下再試');
  if (code !== 200) throw new Error('Anthropic 回 ' + code + '：' + text.slice(0, 200));
  var j = JSON.parse(text);
  var out = j.content && j.content[0] && j.content[0].text;
  if (!out) throw new Error('Claude 沒有回傳內容');
  out = ('{' + out).replace(/```json|```/g, '').trim();
  var cut = out.lastIndexOf('}');
  if (cut > 0) out = out.slice(0, cut + 1);
  return JSON.parse(out);
}

/** 看一張截圖，讀出目前剩餘金額（沒有證據就回空字串，不猜）。要用 Gemini 或 Claude，Vision 備援做不到 */
function readBalance_(b64, mime) {
  if (!b64) throw new Error('沒有收到圖片');
  mime = mime || 'image/jpeg';
  var engine = P_('ENGINE').toLowerCase(), r;
  if (engine === 'claude' && P_('ANTHROPIC_KEY')) r = claudeBalance_(b64, mime);
  else if (engine === 'gemini' && P_('GEMINI_KEY')) r = geminiBalance_(b64, mime);
  else if (P_('ANTHROPIC_KEY')) r = claudeBalance_(b64, mime);
  else if (P_('GEMINI_KEY')) r = geminiBalance_(b64, mime);
  else throw new Error('讀餘額要用 Gemini 或 Claude：還沒設定 GEMINI_KEY 或 ANTHROPIC_KEY');
  return tidyBalanceRead_(r);
}

/** 沒有 Gemini 時的備援：Vision API 取字 + 規則判斷日期。 */
function visionOcr_(b64) {
  var res = UrlFetchApp.fetch(
    'https://vision.googleapis.com/v1/images:annotate?key=' + encodeURIComponent(P_('VISION_KEY')), {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({ requests: [{
        image: { content: b64 },
        features: [{ type: 'TEXT_DETECTION' }],
        imageContext: { languageHints: ['zh-Hant', 'en'] }
      }] })
    });
  if (res.getResponseCode() !== 200) throw new Error('Vision 回 ' + res.getResponseCode());
  var j = JSON.parse(res.getContentText());
  var r = j.responses && j.responses[0];
  return (r && r.fullTextAnnotation && r.fullTextAnnotation.text) || '';
}

function rulesParse_(text) {
  var d = { merchant: '', title: '', value: '', code: '', codeType: '',
            expiry: '', rawText: String(text).slice(0, 600), confidence: 'low' };
  if (!text) return d;

  var labelled = text.match(
    /(?:有效期限|使用期限|兌換期限|優惠期限|有效期間|使用至|到期日)[^0-9]{0,10}((?:民國)?\d{2,4})\s*[\/\.\-年]\s*(\d{1,2})\s*[\/\.\-月]\s*(\d{1,2})/);
  var any = labelled || text.match(/((?:民國)?\d{2,4})\s*[\/\.\-年]\s*(\d{1,2})\s*[\/\.\-月]\s*(\d{1,2})/g);

  if (labelled) {
    d.expiry = ymd_(labelled[1], labelled[2], labelled[3]);
    d.confidence = 'medium';
  } else if (any && any.length) {
    var last = any[any.length - 1].match(/((?:民國)?\d{2,4})\s*[\/\.\-年]\s*(\d{1,2})\s*[\/\.\-月]\s*(\d{1,2})/);
    if (last) d.expiry = ymd_(last[1], last[2], last[3]);
  }

  var v = text.match(/NT\$\s?[\d,]+|\$\s?[\d,]+|\d+\s?(?:杯|張|份|客|籠|入|次)/);
  if (v) d.value = v[0].trim();

  var c = text.match(/[0-9]{4,}[-\s][0-9A-Z-\s]{4,}[0-9A-Z]|[0-9]{10,}/);
  if (c) { d.code = c[0].trim(); d.codeType = 'text'; }

  var lines = text.split(/\r?\n/).map(function (s) { return s.trim(); })
                  .filter(function (s) { return s.length >= 2 && s.length <= 30; });
  if (lines.length) { d.merchant = lines[0]; d.title = lines[1] || lines[0]; }

  return d;
}

/** 把各種寫法正規化成 YYYY-MM-DD（含民國年）。 */
function normDate_(s) {
  s = String(s || '').trim();
  if (!s) return '';
  var m = s.match(/^(\d{2,4})\D(\d{1,2})\D(\d{1,2})$/);
  if (m) return ymd_(m[1], m[2], m[3]);
  m = s.match(/^(\d{2,4})\D(\d{1,2})$/);          // 只到月份 → 當月最後一天
  if (m) {
    var y = toAD_(m[1]), mo = +m[2];
    return ymd_(String(y), String(mo), String(new Date(y, mo, 0).getDate()));
  }
  return '';
}

function toAD_(y) {
  y = parseInt(String(y).replace(/\D/g, ''), 10);
  if (isNaN(y)) return 0;
  return y < 1000 ? y + 1911 : y;   // 三位數以下視為民國年
}

function ymd_(y, m, d) {
  var Y = toAD_(y), M = parseInt(m, 10), D = parseInt(d, 10);
  if (!Y || !M || !D || M > 12 || D > 31) return '';
  return Y + '-' + ('0' + M).slice(-2) + '-' + ('0' + D).slice(-2);
}

function imageOf_(id) {
  var t = findTicket_(id);
  if (!t || !t.fileId) return { dataUrl: null };
  try {
    var b = DriveApp.getFileById(t.fileId).getBlob();
    return { dataUrl: 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes()) };
  } catch (e) {
    return { dataUrl: null };
  }
}


/* ══════════════════ 從雲端硬碟「票券上傳區」匯入 ══════════════════
 *
 *  上傳區是收件匣：處理過的截圖會自動搬走，資料夾裡只剩還沒處理的，所以永遠掃得很快。
 *    匯入成功 → 搬到 App 自己的「票券截圖」資料夾（FOLDER_ID），票券的圖跟著走，不會壞。
 *    序號重複 → 搬到上傳區裡的「重複」子資料夾，確認沒問題可以整個刪掉。
 *    辨識失敗 → 搬到上傳區裡的「辨識失敗」子資料夾；想重試就拖回上傳區。
 *
 *  用法：指令碼屬性 SOURCE_FOLDER_ID 填上傳區的 ID，執行 startAutoImport 打開每 15 分鐘的排程；
 *  等不及就在 App「設定 → 立即匯入上傳區」按一下，或在編輯器執行 importDriveFolder。
 *
 *  如果 SOURCE_FOLDER_ID 指的是你「不想被動到」的舊資料夾（只想一次性匯入），
 *  在指令碼屬性加 IMPORT_KEEP_FILES = 1，檔案就會原地不動。
 */

var IMPORT_BUDGET_MS = 4.5 * 60 * 1000;   // 留緩衝給收尾，別踩到 6 分鐘上限
var IMPORT_PAUSE_MS  = 800;               // 每張之間喘一下，免得打爆免費額度
var IMPORT_MAX_BYTES = 15 * 1024 * 1024;  // 超過這個大小的圖直接當失敗
var IMPORT_DUP_DIR   = '重複';
var IMPORT_FAIL_DIR  = '辨識失敗';

/*
 * 匯入一輪可能跑好幾分鐘，不能一直握著寫入鎖（不然你在 App 按「已使用」要等它跑完）。
 * 改成在指令碼屬性記「正在匯入、幾點開始的」，只有登記和清掉那一下需要鎖；
 * 超過 7 分鐘還沒清掉的記號，當作上一輪中途當掉了（Apps Script 單次最多 6 分鐘），直接接手。
 * 真正寫進試算表的那一下，才跟其他寫入一起排隊。
 */
var IMPORT_STALE_MS = 7 * 60 * 1000;
function beginImport_() {
  return withLock_(function () {
    var p = PropertiesService.getScriptProperties();
    var since = +p.getProperty('IMPORT_BUSY') || 0, now = new Date().getTime();
    if (since && now - since < IMPORT_STALE_MS) return 0;
    p.setProperty('IMPORT_BUSY', String(now));
    return now;
  });
}
function endImport_(mark) {
  if (!mark) return;
  try {
    withLock_(function () {
      var p = PropertiesService.getScriptProperties();
      if (p.getProperty('IMPORT_BUSY') === String(mark)) p.deleteProperty('IMPORT_BUSY');
    });
  } catch (e) { Logger.log('清不掉匯入記號（7 分鐘後會自己失效）：' + e.message); }
}

/** 排程、「立即匯入」、在編輯器直接執行都走這裡：同一時間只會有一輪在跑。 */
function importDriveFolder(budgetMs) {
  var mark = beginImport_();
  if (!mark) {
    Logger.log('另一輪匯入正在跑，這次跳過，等它跑完再試。');
    return { busy: true };
  }
  try {
    return importRun_(budgetMs);
  } finally {
    endImport_(mark);
  }
}

function importRun_(budgetMs) {
  var budget = (typeof budgetMs === 'number' && budgetMs > 0) ? budgetMs : IMPORT_BUDGET_MS;  // 排程直接呼叫時會傳事件物件進來
  var folderId = P_('SOURCE_FOLDER_ID');
  if (!folderId) {
    throw new Error('請先到「專案設定 → 指令碼屬性」新增 SOURCE_FOLDER_ID，' +
                    '值是資料夾網址 .../folders/ 後面那一串');
  }

  var p = PropertiesService.getScriptProperties();
  var t0 = new Date().getTime();
  var src = DriveApp.getFolderById(folderId);
  var keep = P_('IMPORT_KEEP_FILES') === '1' || folderId === P_('FOLDER_ID');

  // 先記住已經有的，重跑才不會重複匯入
  var known = {}, prints = {};
  tickets_().forEach(function (t) {
    if (t.fileId) known[t.fileId] = 1;
    if (t.fingerprint) {      // v76：記下對到的那張用掉了沒（好幾張就以用掉的為準）
      var pv = prints[t.fingerprint];
      if (!pv || (t.status === 'used' && !pv.used)) prints[t.fingerprint] = { used: t.status === 'used', usedAt: t.usedAt || '' };
    }
  });

  // 會搬檔的話，上傳區裡只剩沒處理的，每次從頭掃就好；原地不動才需要記進度
  var token = keep ? p.getProperty('IMPORT_CURSOR') : '';
  var it = token ? DriveApp.continueFileIterator(token) : src.getFiles();
  if (!keep) p.deleteProperty('IMPORT_CURSOR');

  var rows = [], items = [], moves = [];
  var skipped = 0, notImage = 0, maybeDup = 0, usedAgain = 0;
  var dupNames = [], failNames = [], stoppedBy = '';

  while (it.hasNext()) {
    if (new Date().getTime() - t0 > budget) { stoppedBy = '時間'; break; }

    var f = it.next();
    var mime = String(f.getMimeType() || '');
    if (mime.indexOf('image/') !== 0) { notImage++; continue; }
    if (known[f.getId()]) { skipped++; moves.push({ f: f, to: 'done' }); continue; }
    if (f.getSize() > IMPORT_MAX_BYTES) {
      failNames.push(f.getName() + '（檔案太大）');
      moves.push({ f: f, to: 'fail' });
      continue;
    }

    var draft;
    try {
      var img = fetchSmaller_(f, mime);
      draft = parseTicket_(img.b64, img.mime);
    } catch (e) {
      var msg = String(e.message || e);
      if (msg.indexOf('額度') >= 0 || msg.indexOf('429') >= 0 || msg.indexOf('流量') >= 0) {
        stoppedBy = '額度'; break;                      // 這張留在上傳區，下一輪再來
      }
      failNames.push(f.getName());
      Logger.log('辨識失敗：' + f.getName() + ' — ' + msg);
      moves.push({ f: f, to: 'fail' });
      continue;
    }

    draft.fingerprint = fingerprint_(draft);
    var flags = [];
    if (draft.confidence !== 'high') flags.push('待確認');
    if (draft.renamed) flags.push(renamedFlag_(draft));       // v67：套了改名規則

    var seen = draft.fingerprint ? prints[draft.fingerprint] : null;
    if (seen) {
      if (fingerprintKind_(draft) === 'code') {
        // 序號一樣 —— 幾乎確定是同一張，不收（v76：對到的那張已經用掉的，另外標出來、通知也特別寫）
        if (seen.used) usedAgain++;
        dupNames.push(ticketLabel_(draft) + '（' + f.getName() + '）' +
                      (seen.used ? '　⚠ 已使用過' + (seen.usedAt ? '（' + seen.usedAt.slice(5).replace('-', '/') + ' 用掉）' : '') : ''));
        moves.push({ f: f, to: 'dup' });
        continue;
      }
      // 只是商家＋券名＋到期日一樣，可能是同款券的不同組序號 —— 收進來讓你自己判斷
      flags.unshift(seen.used ? '可能重複（同款有一張用掉了）' : '可能重複');
      maybeDup++;
    }
    if (draft.fingerprint && !seen) prints[draft.fingerprint] = { used: false, usedAt: '' };

    draft.id = 'T' + new Date().getTime().toString(36) + Math.random().toString(36).slice(2, 6);
    draft.createdAt = Utilities.formatDate(f.getDateCreated(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    draft.status = 'active';
    draft.fileId = f.getId();
    draft.source = 'drive';
    draft.note = flags.join('・');

    rows.push(rowFrom_(draft, T_HEAD));
    items.push({ label: ticketLabel_(draft), value: draft.value || '', expiry: draft.expiry || '', flags: flags });
    known[f.getId()] = 1;
    moves.push({ f: f, to: 'done' });
    Utilities.sleep(IMPORT_PAUSE_MS);
  }

  // 先把票券寫進表，再搬檔 —— 萬一中途出錯，截圖還在上傳區，下一輪會補做
  if (rows.length) {
    withLock_(function () {
      var sh = tSh_();
      var start = sh.getLastRow() + 1;
      sh.getRange(start, 1, rows.length, T_HEAD.length).setNumberFormat('@').setValues(rows);
    });
  }

  var moveErr = 0;
  if (!keep && moves.length) {
    var dest = {
      done: P_('FOLDER_ID') ? DriveApp.getFolderById(P_('FOLDER_ID')) : subFolder_(src, '已匯入'),
      dup:  null, fail: null
    };
    moves.forEach(function (m) {
      try {
        if (!dest[m.to]) dest[m.to] = subFolder_(src, m.to === 'dup' ? IMPORT_DUP_DIR : IMPORT_FAIL_DIR);
        m.f.moveTo(dest[m.to]);
      } catch (e) { moveErr++; Logger.log('搬檔失敗：' + m.f.getName() + ' — ' + e.message); }
    });
  }

  var more = it.hasNext() || !!stoppedBy;
  if (keep) {
    if (it.hasNext()) p.setProperty('IMPORT_CURSOR', it.getContinuationToken());
    else p.deleteProperty('IMPORT_CURSOR');
  }

  var added = items.length, failed = failNames.length;
  Logger.log([
    '',
    '───────── 這一輪的結果 ─────────',
    '新增票券：' + added,
    items.map(function (x) {
      return '  ・' + x.label + (x.expiry ? '｜' + x.expiry + ' 到期' : '｜無到期日') +
             (x.flags.length ? '｜' + x.flags.join('、') : '');
    }).join('\n'),
    '序號重複沒收：' + dupNames.length + (dupNames.length ? '\n  ・' + dupNames.join('\n  ・') : ''),
    '辨識失敗：' + failed + (failed ? '\n  ・' + failNames.join('\n  ・') : ''),
    skipped ? '之前就匯入過：' + skipped : '',
    notImage ? '不是圖片（留在原地）：' + notImage : '',
    keep ? '檔案原地不動（IMPORT_KEEP_FILES）' : ('處理完的截圖已搬離上傳區' + (moveErr ? '，有 ' + moveErr + ' 個搬不動' : '')),
    '花了：' + Math.round((new Date().getTime() - t0) / 1000) + ' 秒',
    '',
    more
      ? (stoppedBy === '額度'
          ? '⏸ Gemini 額度暫時滿了，沒做完的還在上傳區，下一輪會接著做。'
          : '⏸ 快到單次執行的時間上限了，沒做完的還在上傳區，再執行一次接著跑。')
      : '✅ 上傳區清空了。',
    '────────────────────────────'
  ].filter(function (x) { return x !== ''; }).join('\n'));

  var r = { added: added, skipped: skipped, failed: failed, maybeDup: maybeDup, dup: dupNames.length, usedAgain: usedAgain,
            more: more, stoppedBy: stoppedBy, items: items, dupNames: dupNames, failNames: failNames };

  if (added || failed || dupNames.length) {
    if (!NOTIFY_AUTO_IMPORT) r.notified = { off: true };
    else {
      try { r.notified = notifyImport_(r); }
      catch (e) {
        Logger.log('匯入的通知沒送出：' + e.message);
        r.notified = { ok: false, why: '通知的程式出錯：' + e.message };
      }
    }
  }
  return r;
}

/** 在資料夾裡找（沒有就建）一個子資料夾 */
function subFolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

/** 匯入完成時，在 Discord 條列這一輪做了什麼。 */
function notifyImport_(r) {
  var MAX = 20;
  function list(arr, fmt) {
    var out = arr.slice(0, MAX).map(fmt);
    if (arr.length > MAX) out.push('…還有 ' + (arr.length - MAX) + ' 張');
    return out.join('\n');
  }
  var app = r.source === 'app';
  var notSaved = Array.isArray(r.notSaved) ? r.notSaved : [];
  var parts = [];
  if (r.added) {
    parts.push('**新增 ' + r.added + ' 張**');
    parts.push(list(r.items, function (x) {
      return '• ' + x.label + (x.value ? '（' + x.value + '）' : '') +
             '　' + (x.expiry ? x.expiry + ' 到期' : '無到期日') +
             (x.manual ? '　（手動新增）' : '') +
             (x.flags && x.flags.length ? '　⚠ ' + x.flags.join('、') : '');
    }));
  }
  if (r.dup) {
    parts.push('\n**序號重複，沒收 ' + r.dup + ' 張**（放在上傳區的「' + IMPORT_DUP_DIR + '」）');
    parts.push(list(r.dupNames, function (x) { return '• ' + x; }));
  }
  // v76：已經用掉的券又被匯入（上傳區：沒收；App：沒存，或存了但直接歸到已使用）
  if (r.usedAgain) {
    parts.push('\n⚠ **已經用掉的券又被匯入 ' + r.usedAgain + ' 張**' + (app ? '（App 裡存的會直接歸到「已使用」）' : '（序號跟「已使用」的那張一樣，沒有收進可用）'));
  }
  if (notSaved.length) {
    parts.push('\n**沒存 ' + notSaved.length + ' 張**');
    parts.push(list(notSaved, function (x) { return '• ' + x; }));
  }
  if (r.failed) {
    parts.push('\n**辨識失敗 ' + r.failed + ' 張**' + (app
      ? '（可以在 App 按「重試」，或用右上角 ＋ 自己填）'
      : '（放在上傳區的「' + IMPORT_FAIL_DIR + '」，拖回上傳區可重試）'));
    parts.push(list(r.failNames, function (x) { return '• ' + x; }));
  }
  if (r.error) parts.push('\n⚠ 存到一半出錯：' + r.error + '（上面列的是已經存進去的）');
  if (r.more) parts.push('\n⏸ 還有沒處理完的，' + (r.stoppedBy === '額度' ? 'Gemini 額度滿了，' : '') + '下一輪會接著做');

  var head = '票券匯入' + (app ? '（App）' : '') + '：' +
             [r.added ? '新增 ' + r.added : '', r.dup ? '重複 ' + r.dup : '', notSaved.length ? '沒存 ' + notSaved.length : '',
              r.usedAgain ? '已用過又匯入 ' + r.usedAgain : '', r.failed ? '失敗 ' + r.failed : ''].filter(String).join('、');
  var desc = parts.join('\n');
  if (desc.length > 4000) desc = desc.slice(0, 3990) + '\n…';

  LAST_DISCORD_ = null; LAST_LINE_ = null;
  postDiscord_({
    username: '票券',
    content: head,
    allowed_mentions: { parse: [] },
    embeds: [{ color: r.failed ? 0xC4372B : 0x6A35A8, description: desc, timestamp: new Date().toISOString() }]
  });
  notifyLine_(head);
  return notifyResult_();
}

/** App 裡「拍照或匯入」存完之後，也發一則匯入通知。不想收就改成 false。 */
var NOTIFY_APP_IMPORT = true;

/**
 * App 按「全部存入」之後送來的結果：存了哪幾張、哪幾張沒存（看起來匯過／沒勾）、哪幾張辨識失敗。
 * 跟上傳區的自動匯入用同一個格式發到 Discord。
 */
function importReport_(req) {
  if (!NOTIFY_APP_IMPORT) return { sent: false, notified: { off: true } };
  function clip(v, n) { return String(v == null ? '' : v).replace(/[\r\n\t]+/g, ' ').trim().slice(0, n); }
  function arr(a) { return Array.isArray(a) ? a.slice(0, 60) : []; }
  var items = arr(req.added).map(function (x) {
    x = x || {};
    return {
      label: ticketLabel_({ merchant: clip(x.merchant, 40), title: clip(x.title, 80) }),
      value: clip(x.value, 30),
      expiry: /^\d{4}-\d{2}-\d{2}$/.test(String(x.expiry || '')) ? String(x.expiry) : '',
      manual: !!x.manual,
      flags: arr(x.flags).map(function (f) { return clip(f, 20); }).filter(String)
    };
  });
  var skipped = arr(req.skipped).map(function (x) {
    x = x || {};
    return (clip(x.label, 80) || '一張票券') + (x.why ? '（' + clip(x.why, 20) + '）' : '');
  });
  var failNames = arr(req.failed).map(function (x) {
    x = x || {};
    return (clip(x.name, 60) || '一張截圖') + (x.error ? ' — ' + clip(x.error, 80) : '');
  });
  if (!items.length && !skipped.length && !failNames.length) return { sent: false };
  var usedAgain = Math.max(0, Math.min(60, Math.floor(+req.usedAgain) || 0));    // v76：App 說有幾張是「已經用掉的券又被匯入」
  var n = notifyImport_({ source: 'app', added: items.length, items: items, notSaved: skipped,
                          failed: failNames.length, failNames: failNames, dup: 0, dupNames: [], usedAgain: usedAgain,
                          more: false, error: clip(req.error, 120) });
  return { sent: !!(n && n.ok), notified: n };
}

/**
 * 打開自動匯入：每 15 分鐘掃一次雲端硬碟資料夾。
 * 在編輯器選這個函式按執行就好，只要執行一次。
 */
function startAutoImport() {
  if (!P_('SOURCE_FOLDER_ID')) {
    throw new Error('請先在指令碼屬性設好 SOURCE_FOLDER_ID（要自動匯入的雲端硬碟資料夾）');
  }
  stopAutoImport();
  ScriptApp.newTrigger('autoImport').timeBased().everyMinutes(15).create();
  Logger.log('自動匯入已經打開：每 15 分鐘掃一次。丟進資料夾的截圖會自己變成票券，完成後 Discord 會通知。');
}

/** 關掉自動匯入。 */
function stopAutoImport() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'autoImport' || t.getHandlerFunction() === 'importDriveFolder') {
      ScriptApp.deleteTrigger(t); n++;
    }
  });
  Logger.log(n ? ('已經關掉 ' + n + ' 個自動匯入的排程。') : '本來就沒有自動匯入的排程。');
}

/** 排程真正呼叫的函式：上一輪還在跑就跳過（importDriveFolder 自己會擋），沒東西可做就安靜結束。 */
function autoImport() {
  try { ensureReminderTrigger_(); } catch (e) { Logger.log('檢查到期提醒排程失敗：' + e.message); }
  flushQuietly_();                                  // 之前沒送到的通知，先補送
  try {
    importDriveFolder();
  } catch (e) {
    Logger.log('自動匯入出錯：' + (e && e.message ? e.message : e));
  }
}

/**
 * App 設定頁的「立即匯入上傳區」按鈕呼叫這裡。
 * 網頁請求不能等太久，所以一次只跑 70 秒；沒跑完的，再按一次或等排程接著做。
 * 排程剛好在跑的話會回 { busy: true }。
 */
var IMPORT_NOW_BUDGET_MS = 70 * 1000;
function importNow_() {
  if (!P_('SOURCE_FOLDER_ID')) throw new Error('後端還沒設定上傳區（指令碼屬性 SOURCE_FOLDER_ID）');
  return importDriveFolder(IMPORT_NOW_BUDGET_MS);
}

/** 清掉進度，下次 importDriveFolder 會從資料夾第一個檔案重來。 */
function resetDriveImport() {
  var p = PropertiesService.getScriptProperties();
  p.deleteProperty('IMPORT_CURSOR');
  p.deleteProperty('IMPORT_BUSY');     // 萬一上一輪當掉留下「正在匯入」的記號，也一起清掉
  Logger.log('進度已清除，下次會從頭掃一次（已匯入過的還是會自動跳過）。');
}

/**
 * 盡量拿一張縮小過的版本再送去辨識 —— 原圖動輒好幾 MB，
 * 直接丟給模型又慢又浪費。Drive 的縮圖端點拿不到就退回原圖。
 */
function fetchSmaller_(file, mime) {
  try {
    var res = UrlFetchApp.fetch(
      'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1280', {
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true, followRedirects: true
      });
    if (res.getResponseCode() === 200) {
      var blob = res.getBlob();
      var ct = String(blob.getContentType() || '');
      if (ct.indexOf('image/') === 0 && blob.getBytes().length > 8000) {
        return { b64: Utilities.base64Encode(blob.getBytes()), mime: ct };
      }
    }
  } catch (e) {}
  return { b64: Utilities.base64Encode(file.getBlob().getBytes()), mime: mime };
}


/* ══════════════════ 從手機遠端部署 ══════════════════
 *
 *  電腦上跑著「部署小幫手」（deploy-watcher.js，雙擊「安裝手機部署.bat」會設成開機自動在背景執行），
 *  它每分鐘問一次這裡有沒有人要求部署。要求的方式：
 *    App「設定 → 部署最新版」，或在編輯器執行 requestDeployNow。
 *  小幫手看到請求就在電腦上跑一鍵部署（固定只跑 deploy.js），跑完回報結果，Discord 也會通知。
 *  這裡傳不了任何指令到電腦上；沒有存取碼的人也發不出部署請求。
 */
function parseJson_(s) { try { return s ? JSON.parse(s) : null; } catch (e) { return null; } }

function requestDeploy_(by) {
  var p = PropertiesService.getScriptProperties();
  var cur = parseJson_(p.getProperty('DEPLOY_REQ'));
  var now = Date.now();
  if (cur && now - cur.at < 20 * 60000) {
    return { request: cur, seen: Number(p.getProperty('DEPLOY_SEEN') || 0), already: true };
  }
  var req = { id: randomToken_(10), at: now, by: by || 'App', state: 'pending' };
  p.setProperty('DEPLOY_REQ', JSON.stringify(req));
  return { request: req, seen: Number(p.getProperty('DEPLOY_SEEN') || 0) };
}

/** 在編輯器執行：請電腦上的小幫手部署一次 */
function requestDeployNow() {
  var r = requestDeploy_('Claude');
  var seen = r.seen ? Math.round((Date.now() - r.seen) / 60000) + ' 分鐘前' : '還沒回報過';
  Logger.log((r.already ? '已經有一個部署請求在排隊：' : '已送出部署請求：') + r.request.id +
             '。電腦小幫手上次回報：' + seen + '。');
}

function deployStatus_() {
  var p = PropertiesService.getScriptProperties();
  return {
    request: parseJson_(p.getProperty('DEPLOY_REQ')),
    last: parseJson_(p.getProperty('DEPLOY_LAST')),
    seen: Number(p.getProperty('DEPLOY_SEEN') || 0),
    now: Date.now()
  };
}

function deployPoll_() {
  var p = PropertiesService.getScriptProperties();
  p.setProperty('DEPLOY_SEEN', String(Date.now()));
  var req = parseJson_(p.getProperty('DEPLOY_REQ'));
  if (!req) return { run: false };
  if (req.state === 'running' && Date.now() - (req.startedAt || 0) < 15 * 60000) return { run: false };
  req.state = 'running';
  req.startedAt = Date.now();
  p.setProperty('DEPLOY_REQ', JSON.stringify(req));
  return { run: true, id: req.id };
}

function deployReport_(id, ok, text) {
  var p = PropertiesService.getScriptProperties();
  var req = parseJson_(p.getProperty('DEPLOY_REQ'));
  if (!req || req.id !== id) return { ignored: true };
  var body = String(text || '');
  var last = {
    id: id, ok: !!ok, at: Date.now(), by: req.by,
    site: (body.match(/網站版本：(v\d+)/) || [])[1] || '',
    backend: (body.match(/後端第 (\d+) 版/) || [])[1] || '',
    text: body.slice(-1500)
  };
  p.setProperty('DEPLOY_LAST', JSON.stringify(last));
  p.deleteProperty('DEPLOY_REQ');
  try { notifyDeploy_(last); } catch (e) { Logger.log('部署通知沒送出：' + e.message); }
  return { saved: true };
}

function notifyDeploy_(r) {
  var from = r.by === 'App' ? '從 App 按的' : r.by === 'Claude' ? 'Claude 幫你按的'
           : r.by === '電腦' ? '電腦' + (r.host ? ' ' + r.host : '') + ' 上按一鍵部署' : '';
  var head = r.head || (r.ok
    ? '🚀 部署完成' + (r.site ? '：網站 ' + r.site : '') + (r.backend ? '、後端第 ' + r.backend + ' 版' : '')
    : '⚠ 部署沒有完成');
  var lines = String(r.text || '').split('\n').map(function (x) { return x.trim(); })
    .filter(function (x) { return /^(網站版本|分享短網址)|✗|失敗|少了|找不到/.test(x); }).slice(0, 8);
  postDiscord_({
    username: '票券', allowed_mentions: { parse: [] },
    content: head + (from ? '（' + from + '）' : ''),
    embeds: [{
      color: r.ok ? 0x6A35A8 : 0xC4372B,
      description: (lines.join('\n') || (r.ok ? '' : '看電腦上的「部署紀錄.txt」')) +
                   (r.ok ? '\n手機 App：設定 → 最下面「強制更新到最新版」' : ''),
      timestamp: new Date().toISOString()
    }]
  });
}


/* ══════════════════ 到期提醒 ══════════════════
 *
 *  每天早上 9 點多，在 Discord 列出 10 天內到期、還沒用掉的票券。
 *  到期前 10 天開始，每天都會再提醒一次，直到用掉或過期；10 天內都沒有就不發。
 *
 *  排程會自己裝好：自動匯入每 15 分鐘跑一次時會順便檢查，沒有就補上。
 *  不想收：編輯器執行 stopReminder（之後也不會自己裝回來），要恢復就執行 startReminder。
 *  想馬上看看長怎樣：執行 previewReminder，或 App「設定 → 預覽到期提醒」。
 */
var REMIND_HOUR = 9;
var REMIND_WITHIN = 10;   // 到期前幾天開始每天提醒

/** 排程每天呼叫的 */
function dailyReminder() {
  flushQuietly_();
  try {
    var r = sendReminder_(false);
    // 有發的話 postDiscord_ 已經記了；沒東西發也記一筆，App 看得出「今天早上有跑」
    if (r && r.empty) logNotify_('排程', '⏰ 每日提醒', { ok: true, info: REMIND_WITHIN + ' 天內沒有要到期、還沒用的票券，今天不用發' });
    if (r && r.skipped === 'already') logNotify_('排程', '⏰ 每日提醒', { ok: true, info: '今天已經提醒過了' });
  } catch (e) {
    logNotify_('排程', '⏰ 每日提醒', { ok: false, why: '程式出錯：' + (e && e.message ? e.message : e) });
    throw e;
  }
}

/** 預覽：現在就把 10 天內到期、還沒用的列出來（不影響今天早上的正式提醒） */
function previewReminder() { sendReminder_(true); }

function startReminder() {
  PropertiesService.getScriptProperties().deleteProperty('REMINDER_OFF');
  removeReminderTriggers_();
  createReminderTrigger_();
  Logger.log('到期提醒已經打開：每天早上 ' + REMIND_HOUR + ' 點多，列出 ' + REMIND_WITHIN + ' 天內到期、還沒用的票券。');
}

function stopReminder() {
  PropertiesService.getScriptProperties().setProperty('REMINDER_OFF', '1');
  var n = removeReminderTriggers_();
  Logger.log(n ? '到期提醒已經關掉。' : '本來就沒有到期提醒的排程。');
}

function ensureReminderTrigger_() {
  if (P_('REMINDER_OFF') === '1') return;
  var has = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'dailyReminder'; });
  if (!has) {
    createReminderTrigger_();
    Logger.log('已自動裝好每天早上的到期提醒。');
  }
}

function createReminderTrigger_() {
  ScriptApp.newTrigger('dailyReminder').timeBased()
    .atHour(REMIND_HOUR).everyDays(1).inTimezone(Session.getScriptTimeZone()).create();
}

function removeReminderTriggers_() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'dailyReminder') { ScriptApp.deleteTrigger(t); n++; }
  });
  return n;
}

/** 兩個 yyyy-MM-dd 之間差幾個日曆天（b - a） */
function daysBetween_(a, b) {
  var x = /^(\d{4})-(\d{2})-(\d{2})$/.exec(a), y = /^(\d{4})-(\d{2})-(\d{2})$/.exec(b);
  if (!x || !y) return NaN;
  return Math.round((Date.UTC(+y[1], +y[2] - 1, +y[3]) - Date.UTC(+x[1], +x[2] - 1, +x[3])) / 86400000);
}

/** App 的網址（SHORT_BASE 去掉 /s/） */
function appUrl_() {
  var m = /^(https?:\/\/[^\/]+)/.exec(P_('SHORT_BASE') || '');
  return m ? m[1] + '/' : '';
}

/**
 * 找出 10 天內到期、還沒用的票券並送出。preview=true 時不記「今天發過了」。
 * 自己手上的照到期日分組；分享給朋友、對方還沒用的另外列一區，附上連結開過沒，
 * 方便決定要提醒他，還是收回來自己用。
 * 回傳有沒有真的發出一則。
 */
function sendReminder_(preview) {
  var tz = Session.getScriptTimeZone();
  var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var props = PropertiesService.getScriptProperties();
  if (!preview && props.getProperty('REMIND_LAST') === today) {
    Logger.log('今天已經提醒過了。');
    return { sent: false, skipped: 'already' };
  }

  var shared = activeShareMap_();
  var groups = {}, mine = 0, friends = [];
  var all = tickets_();
  all.forEach(function (t) {
    if (t.status === 'used' || !t.expiry) return;
    var d = daysBetween_(today, t.expiry);
    if (isNaN(d) || d < 0 || d > REMIND_WITHIN) return;
    var sh = shared[t.id];
    if (sh && !sh.keep) { friends.push({ t: t, d: d, sh: sh }); return; }
    if (sh) t._keepTo = sh.sharedTo || '朋友';     // v76：分享出去、你這邊也留著的，算你自己的（備註一下也給了誰）
    mine++;
    (groups[d] = groups[d] || []).push(t);
  });
  friends.sort(function (a, b) { return a.d - b.d; });

  // v69 餘額型：用過了、還沒確認剩多少的，每天提醒到你確認為止；
  // 確認過的餘額超過 BAL_STALE_DAYS 天沒再確認的也列出來。從來沒確認過的不吵你（不然剛開始一整排都是）
  var pend = [], stale = [];
  all.forEach(function (t) {
    if (t.kind !== 'balance' || t.status === 'used') return;
    if (t.expiry && t.expiry < today) return;                      // 已經過期的不算
    if (t.pendingAt) { pend.push(t); return; }
    if (t.balanceAt && !isZeroBalance_(t.balance)) {
      var age = daysBetween_(t.balanceAt, today);
      if (!isNaN(age) && age >= BAL_STALE_DAYS) stale.push({ t: t, age: age });
    }
  });
  pend.sort(function (a, b) { return String(a.pendingAt).localeCompare(String(b.pendingAt)); });
  stale.sort(function (a, b) { return b.age - a.age; });

  var days = Object.keys(groups).map(Number).sort(function (a, b) { return a - b; });
  if (!days.length && !friends.length && !pend.length && !stale.length) {
    Logger.log('今天沒有要提醒的票券：' + REMIND_WITHIN + ' 天內沒有要到期、還沒用的。');
    var said = false;
    if (preview) {
      said = postDiscord_({ username: '票券', allowed_mentions: { parse: [] },
        content: '⏰ 到期提醒（預覽）：' + REMIND_WITHIN + ' 天內沒有要到期的票券 👍' }, { noQueue: true });
    }
    return { sent: said, empty: true, why: preview && !said ? lastWhy_() : '' };
  }

  function when(d) { return d === 0 ? '今天' : d === 1 ? '明天' : d + ' 天後'; }
  function mmdd(iso) { return iso.slice(5).replace('-', '/'); }
  var MAX = 15, parts = [], summary = [];
  days.forEach(function (d) {
    var list = groups[d];
    summary.push(when(d) + ' ' + list.length + ' 張');
    parts.push('**' + when(d) + '到期（' + list.length + ' 張）**');
    list.slice(0, MAX).forEach(function (t) {
      parts.push('• ' + ticketLabel_(t) + (t.value ? '（' + t.value + '）' : '') + '　' + mmdd(t.expiry) +
                 (t._keepTo ? '　（也分享給 ' + t._keepTo + '）' : ''));
    });
    if (list.length > MAX) parts.push('…還有 ' + (list.length - MAX) + ' 張');
    parts.push('');
  });

  // 分享出去、對方還沒用的：寫出給了誰、哪天到期、連結開過沒
  if (friends.length) {
    summary.push('朋友手上 ' + friends.length + ' 張還沒用');
    parts.push('**朋友手上還沒用（' + friends.length + ' 張）**');
    friends.slice(0, MAX).forEach(function (x) {
      var t = x.t, sh = x.sh;
      parts.push('• ' + (sh.sharedTo || '朋友') + '：' + ticketLabel_(t) +
                 '　' + when(x.d) + '到期（' + mmdd(t.expiry) + '）' +
                 '　' + (sh.opens ? '連結開過 ' + sh.opens + ' 次' : '還沒打開過連結'));
    });
    if (friends.length > MAX) parts.push('…還有 ' + (friends.length - MAX) + ' 張');
    parts.push('在 App 點那張票券，可以「提醒他」或「收回這張」自己用。');
    parts.push('');
  }

  // v69 餘額型
  if (pend.length) {
    summary.push('待確認餘額 ' + pend.length + ' 張');
    parts.push('**待確認餘額（' + pend.length + ' 張）**');
    pend.slice(0, MAX).forEach(function (t) {
      parts.push('• ' + ticketLabel_(t) + '　' + (t.pendingBy || '有人') + ' ' + mmdd(t.pendingAt) + ' 用了' +
                 (t.balance ? '（上次確認剩 $' + t.balance + '）' : ''));
    });
    if (pend.length > MAX) parts.push('…還有 ' + (pend.length - MAX) + ' 張');
    parts.push('在 App 打開那張，填剩多少就算確認了。');
    parts.push('');
  }
  if (stale.length) {
    summary.push('餘額 ' + BAL_STALE_DAYS + ' 天沒確認 ' + stale.length + ' 張');
    parts.push('**很久沒確認餘額（' + stale.length + ' 張）**');
    stale.slice(0, MAX).forEach(function (x) {
      parts.push('• ' + ticketLabel_(x.t) + '　上次確認剩 $' + x.t.balance + '（' + x.age + ' 天前）');
    });
    if (stale.length > MAX) parts.push('…還有 ' + (stale.length - MAX) + ' 張');
    parts.push('');
  }

  var app = appUrl_();
  if (app) parts.push('[打開票券](' + app + ')' + (friends.length ? '　[看朋友手上的](' + app + '?f=shared)' : ''));

  var desc = parts.join('\n').trim();
  if (desc.length > 4000) desc = desc.slice(0, 3990) + '\n…';
  var head = '⏰ 到期提醒' + (preview ? '（預覽）' : '') + '：' + summary.join('、');
  var first = Math.min(days.length ? days[0] : 99, friends.length ? friends[0].d : 99);

  var ok = postDiscord_({
    username: '票券',
    content: head,
    allowed_mentions: { parse: [] },
    embeds: [{
      color: first <= 1 ? 0xC4372B : 0x6A35A8,
      description: desc,
      footer: { text: (mine + friends.length
                        ? REMIND_WITHIN + ' 天內到期、還沒用的共 ' + (mine + friends.length) + ' 張' +
                          (friends.length ? '（朋友手上 ' + friends.length + ' 張）' : '') +
                          '，每天早上提醒到用掉或過期為止'
                        : '餘額型的提醒，每天早上提醒到你確認為止') },
      timestamp: new Date().toISOString()
    }]
  }, { noQueue: !!preview });
  notifyLine_(head + '\n' + desc.replace(/\*\*/g, '').replace(/\[([^\]]+)\]\(([^)]+)\)/g, '$1 $2'));
  if (!preview) props.setProperty('REMIND_LAST', today);
  Logger.log(head);
  return { sent: ok, empty: false, count: mine + friends.length + pend.length + stale.length, why: ok ? '' : lastWhy_(),
           queued: !!(LAST_DISCORD_ && LAST_DISCORD_.queued) };
}


/* ══════════════════ 分享 ══════════════════ */

function webAppUrl_() {
  if (P_('WEBAPP_URL')) return P_('WEBAPP_URL');
  try { return ScriptApp.getService().getUrl(); } catch (e) { return ''; }
}

/** 自動匯入有新票券時，要不要在 Discord 報一聲。 */
var NOTIFY_AUTO_IMPORT = true;

/** 產生分享連結時，要不要也在 Discord 留一筆紀錄。不想收就改成 false。 */
var NOTIFY_SHARE = true;

/** 分享連結的 token 長度。短一點好唸，但太短容易被猜到，12 碼大概是 32^12 種組合。 */
var SHARE_TOKEN_LEN = 12;

/**
 * 分享連結長什麼樣子。
 * 指令碼屬性設了 SHORT_BASE（例如 https://你的網站.pages.dev/s/）就用短網址，
 * 那個網址靠前端資料夾裡的 _redirects 轉回這支 Apps Script。沒設就用原本的 /exec?s=token。
 */
/**
 * 分享頁要打哪個 API。設了 SHORT_BASE 就走自己網站的 /api（由 _worker.js 代轉），
 * 對方整趟都不會連到 script.google.com，也就不會撞上 Google 多帳號登入的問題。
 */
function shareApiBase_() {
  var short = P_('SHORT_BASE');
  var m = /^(https?:\/\/[^\/]+)/.exec(short || '');
  return m ? m[1] + '/api' : webAppUrl_();
}

function shareUrl_(token) {
  var short = P_('SHORT_BASE');
  if (short) return short.replace(/\/+$/, '') + '/' + token;
  var base = webAppUrl_();
  if (!base) throw new Error('抓不到網頁應用程式網址，請在指令碼屬性新增 WEBAPP_URL（你的 /exec 網址）');
  return base + '?s=' + token;
}

/*
 * v62：產生分享連結不再等 Discord。
 *   以前這個請求會一路等到通知送完才回應；Discord 慢（重試、改直接送）的時候超過 App 的 45 秒，
 *   App 就放棄等待，但連結其實已經建好了，畫面上卻沒有連結。
 *   現在：
 *   ・App 自己先想好連結碼（token）一起送來，建好立刻回應；回應掉了也能用這個碼去分享紀錄找回來。
 *   ・later=true 時不在這個請求裡送通知，App 顯示連結之後再打一次 shareNotify 補送。
 *   沒帶 token、沒帶 later 的舊版 App，行為跟以前一樣。
 */
var SHARE_TOKEN_RE_ = /^[abcdefghijkmnpqrstuvwxyz23456789]+$/;
function makeShare_(ids, days, note, to, opts) {
  opts = opts || {};
  ids = (ids || []).filter(Boolean);
  if (!ids.length) throw new Error('沒有選任何票券');
  if (ids.length > MAX_SHARE) throw new Error('一個連結最多 ' + MAX_SHARE + ' 張，分成兩次吧');

  var want = String(opts.token || '');
  var wantOk = want.length === SHARE_TOKEN_LEN && SHARE_TOKEN_RE_.test(want);
  var token = wantOk ? want : randomToken_(SHARE_TOKEN_LEN);
  var expiresAt = '';
  var until = String(opts.until || '');
  if (/^\d{4}-\d{2}-\d{2}$/.test(until)) {
    // v73：「跟票券一起到期」→ App 直接給日期（票券的到期日）；已經過了的當作今天
    expiresAt = until < today_() ? today_() : until;
  } else if (days && +days > 0) {
    var d = new Date();
    d.setDate(d.getDate() + (+days));
    expiresAt = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  // v76：這幾張分享出去、自己這邊也留著（只收這條連結裡真的有的）
  var keepIds = (Array.isArray(opts.keep) ? opts.keep : []).map(String).filter(function (id) { return id && ids.indexOf(id) >= 0; });
  var row = {
    token: token, createdAt: today_(), ticketIds: ids.join(','),
    note: String(note || ''), expiresAt: expiresAt, revoked: '',
    openCount: '0', lastOpenedAt: '', to: String(to || '').slice(0, 20),
    keepIds: keepIds.join(',')
  };
  var key = shareKey_(ids), hit = null, why = '', moved = 0;
  withLock_(function () {
    var all = shares_();
    if (wantOk) {
      var same = all.filter(function (s) { return s.token === row.token; })[0];
      if (same && shareKey_(same.ticketIds) === key) { hit = same; why = 'retry'; return; }   // v72：同一個碼重送（App 斷線重試）→ 回原本那條，不會多一條
      if (same) row.token = randomToken_(SHARE_TOKEN_LEN);   // 真的撞到別人的碼（幾乎不可能）：換一個，不覆蓋別人的連結
    }
    // v72：舊版 App 斷線後叫你再按一次，會帶新的碼再來一次 → 同樣幾張、同一個人、同一段話、今天建的、還沒人開過，就回原本那條
    // （新版 App 自己會先問「這張已經分享過」，要新的才帶 force；它也不會換碼重送，所以不走這條）
    if (!opts.instant && !opts.force) {
      var today = today_();
      for (var i = all.length - 1; i >= 0; i--) {
        var s = all[i];
        if (s.revoked === 'y' || s.createdAt !== today || (+s.openCount || 0) > 0) continue;
        if (s.expiresAt && s.expiresAt < today) continue;
        if (shareKey_(s.ticketIds) === key && (s.to || '') === row.to && (s.note || '') === row.note) { hit = s; why = 'again'; return; }
      }
    }
    if (opts.move) moved = moveFromShares_(all, ids);   // v75「改給新的連結」：舊連結裡的這幾張先收回
    var sh = sSh_();
    var r = sh.getLastRow() + 1;
    sh.getRange(r, 1, 1, S_HEAD.length).setNumberFormat('@').setValues([rowFrom_(row, S_HEAD)]);
  });

  if (hit) {
    if (hit.revoked === 'y') throw new Error('這個連結已經收回了');
    return { token: hit.token, url: shareUrl_(hit.token), count: String(hit.ticketIds || '').split(',').filter(Boolean).length,
             expiresAt: hit.expiresAt, to: hit.to, existed: true, reused: why === 'again' };
  }
  token = row.token;

  var url = shareUrl_(token);

  if (NOTIFY_SHARE && !opts.later) notifyShareNow_(ids, row, url);

  return { token: token, url: url, count: ids.length,
           expiresAt: expiresAt, to: row.to, moved: moved };
}

/**
 * v75「改給新的連結」：同一張不要同時在兩條連結裡（兩邊都看得到條碼，容易撞券）。
 * 把這幾張從其他還有效的連結拿掉；拿到一張都不剩的連結整個標成收回。每個連結只寫一次。
 * 要在鎖裡面呼叫（all 是鎖裡面讀的 shares_()）。回傳拿掉了幾張（同一張在兩條舊連結裡算兩次）。
 */
function moveFromShares_(all, ids) {
  var want = {}, n = 0, today = today_(), sh = sSh_();
  (ids || []).forEach(function (id) { want[id] = 1; });
  all.forEach(function (s) {
    if (s.revoked === 'y' || (s.expiresAt && s.expiresAt < today)) return;
    var have = s.ticketIds ? s.ticketIds.split(',') : [];
    var left = have.filter(function (id) { return id && !want[id]; });
    if (left.length === have.filter(Boolean).length) return;
    sh.getRange(s._row, S_HEAD.indexOf('ticketIds') + 1).setNumberFormat('@').setValue(left.join(','));
    if (!left.length) sh.getRange(s._row, S_HEAD.indexOf('revoked') + 1).setValue('y');
    n += have.filter(Boolean).length - left.length;
    s.ticketIds = left.join(',');
  });
  return n;
}

/** 連結裡是哪幾張（不管順序），用來比「是不是同一個分享」 */
function shareKey_(ids) {
  if (typeof ids === 'string') ids = ids.split(',');
  return (ids || []).map(String).filter(Boolean).sort().join(',');
}

/**
 * v72：App 先把連結給你、後端幾秒後才存好。對方剛好在這幾秒點開，就先等一下、自己重新整理（每 4 秒一次，最多 8 次）。
 * 只有長得像 App 產生的連結碼才等；收回的、過期的不等（那些有自己的說明）。次數記在網址的 #w1、#w2…（不靠瀏覽器儲存，LINE 裡也能用）。
 */
var SHARE_WAIT_TRIES = 8;
function shareWaitable_(token) {
  token = String(token || '');
  return token.length === SHARE_TOKEN_LEN && SHARE_TOKEN_RE_.test(token);
}
function shareWaitHtml_() {
  return '<p class="msg" id="w">這個連結剛建立，還在準備中⋯幾秒後會自動重新整理。</p>' +
    '<script>(function(){var m=/w(\\d+)/.exec(location.hash),n=m?+m[1]:0;' +
    'if(n>=' + SHARE_WAIT_TRIES + '){document.getElementById("w").textContent="這個分享連結不存在，或已經被刪掉了。";return;}' +
    'setTimeout(function(){location.hash="w"+(n+1);location.reload();},4000);})();</script>';
}

/** 送「你分享了 N 張票券」的 Discord 通知；失敗不影響分享本身 */
function notifyShareNow_(ids, row, url) {
  try {
    var picked = [];
    var all = tickets_();
    ids.forEach(function (id) {
      for (var i = 0; i < all.length; i++) if (all[i].id === id) picked.push(all[i]);
    });
    notify_({ kind: 'share', who: row.to, tickets: picked, url: url,
              expiresAt: row.expiresAt, note: row.note, keep: row.keepIds });
  } catch (e) {
    Logger.log('分享通知沒送出：' + e.message);   // 通知失敗不影響分享本身
  }
}

/** v62：App 顯示連結之後才來補送分享通知（同一個連結只送一次，收回的不送） */
function shareNotify_(token) {
  token = String(token || '');
  if (!NOTIFY_SHARE) return { sent: false };
  var cache = null, key = 'SHARE_NOTIFIED_' + token;
  try { cache = CacheService.getScriptCache(); } catch (e) {}
  if (cache && cache.get(key)) return { sent: false, already: true };
  var found = null;
  shares_().forEach(function (s) { if (s.token === token) found = s; });
  if (!found) throw new Error('找不到這個分享連結');
  if (found.revoked === 'y') return { sent: false };
  if (cache) { try { cache.put(key, '1', 900); } catch (e) {} }
  notifyShareNow_(String(found.ticketIds || '').split(',').filter(Boolean), found, shareUrl_(token));
  return { sent: true };
}

function listShares_() {
  var all = tickets_();
  var byId = {};
  all.forEach(function (t) { byId[t.id] = t; });

  var today = today_();
  return shares_().map(function (s) {
    var ids = s.ticketIds ? s.ticketIds.split(',') : [];
    var used = 0, titles = [], users = [], left = 0, soon = '';
    ids.forEach(function (id) {
      var t = byId[id];
      if (!t) return;
      if (t.status === 'used') {
        used++;
        if (t.usedBy && users.indexOf(t.usedBy) < 0) users.push(t.usedBy);
      } else if (!t.expiry || t.expiry >= today) {
        // 對方還沒用、也還沒過期的：記下最快哪天到期，分享紀錄要顯示
        left++;
        if (t.expiry && (!soon || t.expiry < soon)) soon = t.expiry;
      }
      if (titles.length < 3) titles.push(t.merchant || t.title || '票券');
    });
    return {
      token: s.token, createdAt: s.createdAt, to: s.to, note: s.note,
      count: ids.length, used: used, users: users, titles: titles,
      left: left, soon: soon,
      expiresAt: s.expiresAt, revoked: s.revoked === 'y',
      openCount: +s.openCount || 0, lastOpenedAt: s.lastOpenedAt,
      url: shareUrl_(s.token)
    };
  }).sort(function (a, b) { return a.createdAt < b.createdAt ? 1 : -1; });
}

function revokeShare_(token) {
  return withLock_(function () {
    var all = shares_();
    for (var i = 0; i < all.length; i++) {
      if (all[i].token === token) {
        sSh_().getRange(all[i]._row, S_HEAD.indexOf('revoked') + 1).setValue('y');
        return {};
      }
    }
    throw new Error('找不到這個分享連結');
  });
}

/** 把一張從連結裡拿掉；拿到一張都不剩就整個連結標成收回。回傳還剩幾張。要在鎖裡面呼叫。 */
function dropFromShare_(s, id) {
  var ids = s.ticketIds ? s.ticketIds.split(',') : [];
  var left = ids.filter(function (x) { return x && x !== id; });
  var sh = sSh_();
  if (left.length !== ids.length) {
    sh.getRange(s._row, S_HEAD.indexOf('ticketIds') + 1).setNumberFormat('@').setValue(left.join(','));
  }
  if (!left.length && s.revoked !== 'y') {
    sh.getRange(s._row, S_HEAD.indexOf('revoked') + 1).setValue('y');
  }
  s.ticketIds = left.join(',');
  return left.length;
}

/**
 * 只把連結裡的「這一張」收回來自己用：從那個連結拿掉它，對方還能用同一個連結裡的其他張。
 * 拿到一張都不剩，就整個連結標成收回（對方打開會看到「都已經收回或還回去了」）。
 * 對方已經用掉的就不收了，回報是誰、哪天用的。
 */
function unshareTicket_(token, id) {
  id = String(id || '');
  return withLock_(function () {
    var s = shareRow_(token);
    if (!s) throw new Error('找不到這個分享連結');
    var t = findTicket_(id);
    if (t && t.status === 'used') {
      return { used: true, usedBy: t.usedBy || '', usedAt: t.usedAt || '' };
    }
    return { used: false, left: dropFromShare_(s, id) };
  });
}

/**
 * v63：一次收回好幾張（App「已分享」多選之後按「收回」）。items = [{token, id}, ...]。
 * 整批在同一把鎖裡做，票券和連結各讀一次，每個連結只寫一次（一張一張收會慢很多）。
 * 規則跟 unshareTicket_ 一樣：只把那幾張從各自的連結拿掉，同連結沒選的對方還能用；
 * 拿到一張都不剩的連結就整個標成收回；對方已經用掉的不收，回報是誰用的。
 * 連結不在了、或那張早就不在連結裡（上次回應掉了才重送，也會碰到）：當作已經收回，不報錯。
 * 回傳 { done: 這次真的收回幾張, gone: 本來就不在連結裡幾張, used: [{id, usedBy, usedAt}],
 *        left: { 連結碼: 收完之後那個連結裡對方還能用的有幾張（用掉的不算） } }
 */
var UNSHARE_MANY_MAX = 100;
function unshareMany_(items) {
  if (!Array.isArray(items) || !items.length) throw new Error('沒有選任何票券');
  if (items.length > UNSHARE_MANY_MAX) throw new Error('一次最多收回 ' + UNSHARE_MANY_MAX + ' 張');
  return withLock_(function () {
    var shareBy = {}, ticketBy = {};
    shares_().forEach(function (s) { shareBy[s.token] = s; });
    tickets_().forEach(function (t) { ticketBy[t.id] = t; });
    var out = { done: 0, gone: 0, used: [], left: {} };
    var drop = {}, seen = {};
    items.forEach(function (it) {
      var token = String((it && it.token) || ''), id = String((it && it.id) || '');
      if (!token || !id || seen[token + '|' + id]) return;
      seen[token + '|' + id] = true;
      var s = shareBy[token];
      var ids = s && s.ticketIds ? s.ticketIds.split(',') : [];
      if (!s || ids.indexOf(id) < 0) { out.gone++; return; }
      var t = ticketBy[id];
      if (t && t.status === 'used') {
        out.used.push({ id: id, usedBy: t.usedBy || '', usedAt: t.usedAt || '' });
        return;
      }
      (drop[token] = drop[token] || []).push(id);
    });
    var sh = sSh_();
    Object.keys(drop).forEach(function (token) {
      var s = shareBy[token];
      var left = s.ticketIds.split(',').filter(function (x) { return x && drop[token].indexOf(x) < 0; });
      sh.getRange(s._row, S_HEAD.indexOf('ticketIds') + 1).setNumberFormat('@').setValue(left.join(','));
      if (!left.length && s.revoked !== 'y') sh.getRange(s._row, S_HEAD.indexOf('revoked') + 1).setValue('y');
      out.done += drop[token].length;
      out.left[token] = left.filter(function (x) { return !(ticketBy[x] && ticketBy[x].status === 'used'); }).length;
    });
    return out;
  });
}

/**
 * 朋友在分享頁按「用不到，還給他」：這張從連結拿掉、回到你的「可用」，並通知你。
 * 已經標記用掉的不能還；同一張按兩次（或網路重送）當作成功。
 */
function returnByShare_(token, id, by) {
  id = String(id || '');
  var who = String(by || '').replace(/[\r\n\t]/g, ' ').trim().slice(0, 20);
  var r = withLock_(function () {
    var s = openShare_(token);                      // 連結收回、過期、不存在都在這裡擋掉
    var ids = s.ticketIds ? s.ticketIds.split(',') : [];
    if (ids.indexOf(id) < 0) return { already: true, left: ids.length };
    var t = findTicket_(id);
    if (!t) throw new Error('找不到這張票券');
    if (t.status === 'used') return { used: true, usedAt: t.usedAt || '' };
    var left = dropFromShare_(s, id);
    return { left: left, t: t, s: s };
  });
  if (r.t) {
    try {
      notify_({ kind: 'returned', who: who || r.s.to || '對方', ticket: r.t, shareTo: r.s.to,
                token: token, share: r.s, leftCount: r.left });
    } catch (e) {
      Logger.log('還票的通知沒送出：' + e.message);   // 通知失敗不影響還票本身
    }
  }
  return { used: !!r.used, usedAt: r.usedAt || '', already: !!r.already, left: r.left || 0 };
}

function shareRow_(token) {
  var all = shares_();
  for (var i = 0; i < all.length; i++) if (all[i].token === token) return all[i];
  return null;
}

function openShare_(token) {
  var s = shareRow_(token);
  if (!s) { var nf = new Error('這個分享連結不存在，或已經被刪掉了'); nf.notFound = true; throw nf; }
  if (s.revoked === 'y') {
    throw new Error(s.ticketIds ? '分享的人已經收回這個連結了' : '這個連結裡的票券都已經收回或還回去了');
  }
  if (s.expiresAt && s.expiresAt < today_()) throw new Error('這個連結已經在 ' + s.expiresAt + ' 過期了');
  return s;
}

/** 這張票券是不是這個連結標記的、而且還在可以復原的時間內；回傳截止時間（毫秒），不行就回 0。 */
function undoUntil_(t, token) {
  if (t.status !== 'used') return 0;
  var v = String(t.usedVia || '');
  if (v.indexOf(token + '@') !== 0) return 0;
  if (!SHARE_UNDO_MINUTES) return 8.64e15;                      // 不限時間
  var when = Date.parse(v.slice(token.length + 1));
  if (isNaN(when)) return 0;
  var until = when + SHARE_UNDO_MINUTES * 60000;
  return until > new Date().getTime() ? until : 0;
}

function sharedTicket_(token, id) {
  var s = openShare_(token);
  if (s.ticketIds.split(',').indexOf(id) < 0) throw new Error('這張已經被分享的人收回了，重新整理看看');
  var t = findTicket_(id);
  if (!t) throw new Error('找不到這張票券');
  return { s: s, t: t };
}

// 回給分享頁的東西刻意不帶「誰用的」—— 那是只有你的 App 看得到的資訊
function markUsedByShare_(token, id, by) {
  var who = String(by || '').replace(/[\r\n\t]/g, ' ').trim().slice(0, 20);
  // 找票＋寫入在鎖裡面做；通知比較慢，放到鎖外面
  var r = withLock_(function () {
    var x = sharedTicket_(token, id), s = x.s, t = x.t;
    if (t.status === 'used') {
      return { out: { already: true, usedAt: t.usedAt, undoUntil: undoUntil_(t, token) } };
    }
    // v69：餘額型，對方用了不等於用光 → 不進「已使用」，只標「待確認餘額」並通知你，由你確認剩多少
    if (t.kind === 'balance' && t.pendingAt) {
      return { out: { already: true, pending: true, usedAt: t.pendingAt, undoUntil: 0 } };   // 已經通知過了，不重複吵你
    }
    if (!who) throw new Error('請填你的名字，對方才知道是誰用的');
    if (t.kind === 'balance') {
      updateTicket_(id, { pendingAt: today_(), pendingBy: who }, t);
      return { s: s, t: t, balance: true, out: { already: false, pending: true, usedAt: today_(), undoUntil: 0 } };
    }
    var stamp = new Date().toISOString();
    var merged = updateTicket_(id, {
      status: 'used', usedAt: today_(), usedBy: who, usedVia: token + '@' + stamp
    }, t);
    return { s: s, t: t, out: { already: false, usedAt: today_(), undoUntil: undoUntil_(merged, token) } };
  });

  if (r.t) notify_({ kind: r.balance ? 'balanceUsed' : 'used', who: who, ticket: r.t, shareTo: r.s.to, token: token, share: r.s });
  return r.out;
}

function unmarkUsedByShare_(token, id) {
  var r = withLock_(function () {
    var x = sharedTicket_(token, id), s = x.s, t = x.t;
    if (t.status !== 'used') return { out: { already: true } };   // 已經是未使用，就當作成功
    if (!undoUntil_(t, token)) {
      throw new Error(String(t.usedVia || '').indexOf(token + '@') === 0
        ? '超過 ' + SHARE_UNDO_MINUTES + ' 分鐘了，沒辦法自己復原，請跟分享給你的人說一聲'
        : '這張不是從這個連結標記的，請跟分享給你的人說一聲');
    }
    updateTicket_(id, { status: 'active', usedAt: '', usedBy: '', usedVia: '' }, t);
    return { s: s, t: t, out: { already: false } };
  });

  if (r.t) notify_({ kind: 'undo', who: r.t.usedBy || '對方', ticket: r.t, shareTo: r.s.to, token: token, share: r.s });
  return r.out;
}


/* ══════════════════ v75 親友連結 ══════════════════
 *
 *  第三種身分，介於「主人」（設定連結，看得到全部、什麼都能改）和「分享連結」（一次給幾張、會到期）之間：
 *    ・一個人一條長期連結。網址跟分享連結一樣是 /s/<碼>（_worker.js 不用改），碼長 16，跟分享的 12 碼分得開。
 *    ・看得到哪些：（挑的商家 ∪ 單張加的）－ 單張拿掉的，而且只有還能用的：
 *      用掉的、過期的、分享給別人的（在還有效的分享連結裡）自動不出現；以後新匯入的同一家自動出現。
 *    ・只能出示、按「我用掉了」（自動記他的名字，不用自己填）；餘額型按「我用了一些」→ 待確認，由你確認剩多少。
 *      不能改、刪、分享，看不到設定和分享紀錄。
 *    ・截圖和序號按「出示」才載入（memberImage）：頁面打開快，滑過去也不會看到條碼。
 *    ・你可以改挑選、停用、換新連結（舊的立刻失效）。
 *  資料在試算表的「Members」分頁，第一次新增親友時自動建立。
 */
var M_SHEET = 'Members';
var M_HEAD = ['id', 'token', 'name', 'createdAt', 'brands', 'addIds', 'dropIds', 'revoked', 'openCount', 'lastOpenedAt'];
// brands = 挑的商家（一行一個，跟 App 列表的商家分組同一套名字，見 BRANDS）；addIds / dropIds = 單張加的、單張拿掉的票券 id（逗號隔開）
var MEMBER_TOKEN_LEN = 16;
var MEMBER_MAX = 30;            // 最多幾個親友
var MEMBER_IDS_MAX = 500;       // 單張加減最多記幾張

/**
 * 商家分組：跟 App（index.html）的 BRANDS 是同一份，測試會逐字比對。
 * 親友挑的是「商家」，後端要用同一套規則判斷每張券屬於哪一家。
 */
var BRANDS = [
  ['7-ELEVEN', /7[\s\-]?11|7[\s\-]?eleven|統一超商|小七|city\s*cafe/i],
  ['全家',     /全家|family\s*mart|familymart/i],
  ['萊爾富',   /萊爾富|hi[\s\-]?life/i],
  ['OK 超商',  /ok\s*(超商|mart)/i],
  ['星巴克',   /星巴克|starbucks/i],
  ['麥當勞',   /麥當勞|mcdonald|mccafe|得來速/i],
  ['路易莎',   /路易莎|louisa/i],
  ['cama',     /cama/i],
  ['Mr.Wish',  /mr\.?\s*wish/i],
  ['85度C',    /85\s*度\s*c/i],
  ['摩斯漢堡', /摩斯|mos\s*burger/i],
  ['肯德基',   /肯德基|kfc/i],
  ['漢堡王',   /漢堡王|burger\s*king/i],
  ['必勝客',   /必勝客|pizza\s*hut/i],
  ['達美樂',   /達美樂|domino/i],
  ['家樂福',   /家樂福|carrefour/i],
  ['全聯',     /全聯|pxmart/i],
  ['大潤發',   /大潤發|rt[\s\-]?mart/i],
  ['新光三越', /新光三越/],
  ['SOGO',     /sogo|崇光/i],
  ['誠品',     /誠品|eslite/i],
  ['鼎泰豐',   /鼎泰豐|din\s*tai\s*fung/i],
  ['王品集團', /王品|西堤|陶板屋|聚北海道|石二鍋|藝奇/i],
  ['乾杯',     /乾杯|kanpai/i],
  ['錢櫃 / 好樂迪', /錢櫃|好樂迪|partyworld|holiday\s*ksd/i]
];
function brandOf_(t) {
  var hay = (t.merchant || '') + ' ' + (t.title || '');
  for (var i = 0; i < BRANDS.length; i++) {
    if (BRANDS[i][1].test(hay)) return BRANDS[i][0];
  }
  var m = String(t.merchant || '').trim();
  return m || '未分類';
}

function mSh_(create) {
  var sh = ss_().getSheetByName(M_SHEET);
  if (!sh && create) sh = ensureSheet_(ss_(), M_SHEET, M_HEAD);
  return sh;
}
function members_() { var sh = mSh_(false); return sh ? readRows_(sh, M_HEAD, 'id') : []; }
function memberRow_(token) {
  token = String(token || '');
  if (token.length !== MEMBER_TOKEN_LEN) return null;
  var all = members_();
  for (var i = 0; i < all.length; i++) if (all[i].token === token) return all[i];
  return null;
}
function openMember_(token) {
  var m = memberRow_(token);
  if (!m) { var nf = new Error('這個連結不存在，或已經換成新的連結了'); nf.notFound = true; throw nf; }
  if (m.revoked === 'y') throw new Error('這個連結已經停用了');
  return m;
}
function idList_(s) { return String(s || '').split(',').map(function (x) { return x.trim(); }).filter(Boolean); }
function lineList_(s) { return String(s || '').split('\n').map(function (x) { return x.trim(); }).filter(Boolean); }

/** 這張是不是他挑的（不管能不能用） */
function memberPicks_(m, t) {
  if (idList_(m.dropIds).indexOf(t.id) >= 0) return false;
  if (idList_(m.addIds).indexOf(t.id) >= 0) return true;
  return lineList_(m.brands).indexOf(brandOf_(t)) >= 0;
}
/** 他現在看得到這張嗎：挑的、沒用掉、沒過期、沒分享給別人 */
function memberSees_(m, t, shared, today) {
  if (!t || t.status === 'used') return false;
  if (t.expiry && t.expiry < today) return false;
  if (shared && shared[t.id]) return false;
  return memberPicks_(m, t);
}

/** 給 App 看的一筆親友：full = 設定畫面用（含網址、看得到幾張）；不然只給列表上標「親友」用的 */
function memberView_(m, full, tickets, shared, today) {
  var v = { id: m.id, name: m.name, brands: lineList_(m.brands), addIds: idList_(m.addIds), dropIds: idList_(m.dropIds),
            revoked: m.revoked === 'y' };
  if (!full) return v;
  v.createdAt = m.createdAt; v.openCount = +m.openCount || 0; v.lastOpenedAt = m.lastOpenedAt || '';
  v.url = m.revoked === 'y' ? '' : shareUrl_(m.token);
  v.count = v.revoked ? 0 : tickets.filter(function (t) { return memberSees_(m, t, shared, today); }).length;
  return v;
}
function listMembers_() {
  var tickets = tickets_(), shared = activeShareMap_(), today = today_();
  return members_().map(function (m) { return memberView_(m, true, tickets, shared, today); });
}
/** 列表（list）裡附上的：只要還有效的、不含網址 */
function membersBrief_() {
  try {
    return members_().filter(function (m) { return m.revoked !== 'y'; })
      .map(function (m) { return memberView_(m, false); });
  } catch (e) { return []; }
}

/** 新增或修改一個親友。in = { id?, name, brands[], addIds[], dropIds[] } */
function memberSave_(input) {
  input = input || {};
  var name = String(input.name || '').replace(/[\r\n\t]/g, ' ').trim().slice(0, 20);
  if (!name) throw new Error('請填名字');
  function clean(arr, max, len) {
    var out = [], seen = {};
    (Array.isArray(arr) ? arr : []).forEach(function (x) {
      x = String(x == null ? '' : x).replace(/[\r\n,]/g, ' ').trim().slice(0, len);
      if (x && !seen[x] && out.length < max) { seen[x] = 1; out.push(x); }
    });
    return out;
  }
  var brands = clean(input.brands, 200, 40);
  var addIds = clean(input.addIds, MEMBER_IDS_MAX, 40), dropIds = clean(input.dropIds, MEMBER_IDS_MAX, 40);
  var saved = withLock_(function () {
    var sh = mSh_(true), all = members_();
    var row = null;
    if (input.id) {
      row = all.filter(function (m) { return m.id === String(input.id); })[0];
      if (!row) throw new Error('找不到這個親友，重新整理看看');
    } else {
      if (all.filter(function (m) { return m.revoked !== 'y'; }).length >= MEMBER_MAX) throw new Error('親友最多 ' + MEMBER_MAX + ' 個');
      row = { id: 'M' + randomToken_(10), token: randomToken_(MEMBER_TOKEN_LEN), createdAt: today_(), revoked: '', openCount: '0', lastOpenedAt: '',
              _row: sh.getLastRow() + 1 };
    }
    row.name = name; row.brands = brands.join('\n'); row.addIds = addIds.join(','); row.dropIds = dropIds.join(',');
    sh.getRange(row._row, 1, 1, M_HEAD.length).setNumberFormat('@').setValues([rowFrom_(row, M_HEAD)]);
    return row;
  });
  return { member: memberView_(saved, true, tickets_(), activeShareMap_(), today_()) };
}

/** 停用（revoke=true）或換新連結（renew=true，停用的也會重新打開）。舊連結立刻失效。 */
function memberSet_(id, opts) {
  var saved = withLock_(function () {
    var m = members_().filter(function (x) { return x.id === String(id || ''); })[0];
    if (!m) throw new Error('找不到這個親友，重新整理看看');
    if (opts.renew) { m.token = randomToken_(MEMBER_TOKEN_LEN); m.revoked = ''; m.openCount = '0'; m.lastOpenedAt = ''; }
    else m.revoked = 'y';
    mSh_(true).getRange(m._row, 1, 1, M_HEAD.length).setNumberFormat('@').setValues([rowFrom_(m, M_HEAD)]);
    return m;
  });
  return { member: memberView_(saved, true, tickets_(), activeShareMap_(), today_()) };
}

/** 親友頁按「我用掉了」：自動記他的名字；餘額型只標待確認 */
function memberUse_(token, id) {
  id = String(id || '');
  var r = withLock_(function () {
    var m = openMember_(token);
    var t = findTicket_(id);
    if (!t) throw new Error('找不到這張票券');
    if (t.status === 'used') return { out: { already: true, usedAt: t.usedAt, undoUntil: undoUntil_(t, token) } };
    if (!memberSees_(m, t, activeShareMap_(), today_())) throw new Error('這張現在不能用了（可能已經給別人了），重新整理看看');
    if (t.kind === 'balance') {
      if (t.pendingAt) return { out: { already: true, pending: true, usedAt: t.pendingAt, undoUntil: 0 } };   // 已經通知過了
      updateTicket_(id, { pendingAt: today_(), pendingBy: m.name }, t);
      return { m: m, t: t, balance: true, out: { already: false, pending: true, usedAt: today_(), undoUntil: 0 } };
    }
    var merged = updateTicket_(id, { status: 'used', usedAt: today_(), usedBy: m.name, usedVia: token + '@' + new Date().toISOString() }, t);
    return { m: m, t: t, out: { already: false, usedAt: today_(), undoUntil: undoUntil_(merged, token) } };
  });
  if (r.t) {
    try { notify_({ kind: r.balance ? 'balanceUsed' : 'used', who: r.m.name, ticket: r.t, member: true }); }
    catch (e) { Logger.log('親友用券的通知沒送出：' + e.message); }   // 通知失敗不影響票券本身
  }
  return r.out;
}

/** 親友頁的「按錯了，復原」：只能復原自己從這條連結標的、而且還在時間內 */
function memberUndo_(token, id) {
  id = String(id || '');
  var r = withLock_(function () {
    var m = openMember_(token);
    var t = findTicket_(id);
    if (!t) throw new Error('找不到這張票券');
    if (t.status !== 'used') return { out: { already: true } };
    if (!undoUntil_(t, token)) {
      throw new Error(String(t.usedVia || '').indexOf(token + '@') === 0
        ? '超過 ' + SHARE_UNDO_MINUTES + ' 分鐘了，沒辦法自己復原，請跟給你連結的人說一聲'
        : '這張不是從這個連結標記的，請跟給你連結的人說一聲');
    }
    updateTicket_(id, { status: 'active', usedAt: '', usedBy: '', usedVia: '' }, t);
    return { m: m, t: t, out: { already: false } };
  });
  if (r.t) {
    try { notify_({ kind: 'undo', who: r.m.name, ticket: r.t, member: true }); }
    catch (e) { Logger.log('親友復原的通知沒送出：' + e.message); }
  }
  return r.out;
}

/** 親友頁按「出示」：這時才給截圖和序號（看得到的，或剛剛從這條連結標掉、還能復原的） */
function memberImage_(token, id) {
  var m = openMember_(token);
  var t = findTicket_(String(id || ''));
  if (!t) throw new Error('找不到這張票券');
  if (!memberSees_(m, t, activeShareMap_(), today_()) && !undoUntil_(t, token)) {
    throw new Error('這張現在不能用了（可能已經用掉或給別人了），重新整理看看');
  }
  var out = { code: t.code || '', dataUrl: null };
  try {
    if (t.fileId) {
      var b = DriveApp.getFileById(t.fileId).getBlob();
      out.dataUrl = 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes());
    }
  } catch (e) {}
  return out;
}

/** 親友頁的 HTML。me = 你自己從 App 按「打開看看」，不記開啟次數 */
function memberHtml_(m, me) {
  if (!me) {
    try {
      var sh = mSh_(false);
      sh.getRange(m._row, M_HEAD.indexOf('openCount') + 1).setValue(String((+m.openCount || 0) + 1));
      sh.getRange(m._row, M_HEAD.indexOf('lastOpenedAt') + 1).setValue(today_());
    } catch (e) {}
  }
  var token = m.token, today = today_(), shared = activeShareMap_();
  var list = tickets_().filter(function (t) {
    return memberSees_(m, t, shared, today) || (t.status === 'used' && memberPicks_(m, t) && undoUntil_(t, token));
  }).sort(function (a, b) {
    var ka = a.expiry || '9999', kb = b.expiry || '9999';
    return ka !== kb ? (ka < kb ? -1 : 1) : brandOf_(a).localeCompare(brandOf_(b));
  });

  var counts = {}, order = [];
  list.forEach(function (t) { var b = brandOf_(t); if (!counts[b]) { counts[b] = 0; order.push(b); } counts[b]++; });
  var chips = order.length > 1
    ? '<div class="chips" id="chips"><button type="button" class="chip on" data-b="">全部 ' + list.length + '</button>' +
      order.map(function (b) { return '<button type="button" class="chip" data-b="' + esc_(b) + '">' + esc_(b) + ' ' + counts[b] + '</button>'; }).join('') +
      '</div>'
    : '';

  var cards = list.map(function (t) {
    var used = t.status === 'used';
    var until = undoUntil_(t, token);
    var st = stub_(t);
    var bal = t.kind === 'balance' && !used;
    var pend = bal && !!t.pendingAt;
    var left = bal && /^\d+$/.test(String(t.balance || ''))
      ? '剩 $' + String(t.balance).replace(/\B(?=(\d{3})+(?!\d))/g, ',') : '';    // 只給剩餘金額；查餘額的網址絕不放進來
    return '' +
      '<article class="tk' + (used ? ' used' : '') + (pend ? ' pend' : '') + '" data-id="' + esc_(t.id) + '" data-b="' + esc_(brandOf_(t)) + '"' +
        (until ? ' data-undo-until="' + until + '"' : '') + (bal ? ' data-bal="1"' : '') + '>' +
        '<div class="tw"><div class="tkt" data-tier="' + st.tier + '"' +
            (isHotel_(t) ? ' data-kind="hotel"' : '') +
            (brandColor_(t) ? ' style="--brand:' + brandColor_(t) + '"' : '') + '>' +
          '<span class="edge"></span>' +
          '<div class="bd">' +
            '<span class="mc">' + esc_(t.merchant || '票券') + '</span>' +
            '<h2>' + esc_(t.title || '（未命名）') + '</h2>' +
            (left || t.value ? '<span class="val">' + esc_(left || t.value) + '</span>' : '') +
            '<span class="code" hidden></span>' +
            '<span class="ink">已使用</span>' +
          '</div>' +
          '<div class="stub">' +
            (st.top ? '<span class="s-top">' + esc_(st.top) + '</span>' : '') +
            '<span class="s-num' + (st.cls ? ' ' + st.cls : '') + '">' + esc_(st.num) + '</span>' +
            (st.unit ? '<span class="s-unit">' + esc_(st.unit) + '</span>' : '') +
            (st.date ? '<span class="s-date">' + esc_(st.date) + '</span>' : '') +
          '</div>' +
        '</div></div>' +
        '<div class="shotbox"></div>' +
        '<div class="act">' +
          '<span class="warn"></span>' +
          '<span class="stamp" ' + (used || pend ? '' : 'hidden') + '>' +
            (pend ? '已通知，他會確認剩多少' : '已使用 ' + esc_(t.usedAt || '')) + '</span>' +
          '<button class="undo" ' + (until ? '' : 'hidden') + '>按錯了，復原</button>' +
          '<button class="hide" hidden>收起</button>' +
          '<button class="use" hidden data-label="' + (bal ? '我用了一些' : '我用掉了') + '">' + (bal ? '我用了一些' : '我用掉了') + '</button>' +
          '<button class="show" ' + (used || pend ? 'hidden' : '') + '>出示</button>' +
        '</div>' +
        (bal ? '<p class="balnote">餘額型：用完一次不代表用光。按「我用了一些」，他會收到通知、再確認剩多少。</p>' : '') +
      '</article>';
  }).join('');

  var hasBal = list.some(function (t) { return t.kind === 'balance' && t.status !== 'used'; });
  var body =
    '<p class="msg">' + (list.length
      ? '要用哪張就按「出示」，截圖和條碼才會出來；兌換時把螢幕亮度調到最高，點截圖可以放大、轉橫的。用掉之後按「我用掉了」，這張就會從這裡拿掉。' +
        (SHARE_UNDO_MINUTES ? '按錯的話 ' + SHARE_UNDO_MINUTES + ' 分鐘內可以自己復原。' : '') +
        (hasBal ? '「餘額型」的券用完一次不代表用光，按「我用了一些」就好。' : '')
      : '目前沒有可以用的票券。有新的時候，再打開這個連結就會出現。') +
    '</p>' + chips + cards +
    '<div class="lb" id="lb" hidden>' +
      '<div class="lb-bar">' +
        '<button type="button" class="lb-btn" id="lbClose">關閉</button>' +
        '<button type="button" class="lb-btn" id="lbRot">轉橫向</button>' +
        '<button type="button" class="lb-btn" id="lbZoom">放大</button>' +
      '</div>' +
      '<div class="lb-stage" id="lbStage"><div class="lb-box" id="lbBox"><img id="lbImg" alt="票券截圖"></div></div>' +
    '</div>' +
    '<script>' + MEMBER_JS.replace('__API__', JSON.stringify(shareApiBase_())).replace('__TOK__', JSON.stringify(token)) + '<\/script>';

  return shell_('票券', body, { heading: m.name + '，這些票券給你用', sub: list.length ? '現在有 ' + list.length + ' 張可以用' : '' });
}

/**
 * 截圖放大：點圖全螢幕，可放大、可轉橫的。分享頁（SHARE_JS）和親友頁（MEMBER_JS）共用這一段，
 * 放在兩個頁面那個 (function(){ … })() 的最後面。要先定義在這裡，下面兩個才拿得到。
 */
var LIGHTBOX_JS_LINES = [
  '  // ── 截圖放大：點圖全螢幕，可放大、可轉橫的 ──',
  '  var lb = document.getElementById("lb"), lbImg = document.getElementById("lbImg");',
  '  var lbBox = document.getElementById("lbBox"), lbStage = document.getElementById("lbStage");',
  '  var zoom = 1, rot = false;',
  '',
  '  function layout(){',
  '    var nw = lbImg.naturalWidth, nh = lbImg.naturalHeight;',
  '    if(!nw || !nh) return;',
  '    var sw = lbStage.clientWidth - 8, sh = lbStage.clientHeight - 8;',
  '    // 直向：寬度填滿，上下捲。橫向：把截圖的寬轉成螢幕的高，條碼會變粗變長，左右捲',
  '    var fit = rot ? (sh / nw) : (sw / nw);',
  '    var k = fit * zoom, w = nw * k, h = nh * k;',
  '    lbImg.style.width = w + "px"; lbImg.style.height = h + "px";',
  '    lbImg.style.transform = "translate(-50%,-50%)" + (rot ? " rotate(90deg)" : "");',
  '    lbBox.style.width = (rot ? h : w) + "px";',
  '    lbBox.style.height = (rot ? w : h) + "px";',
  '    document.getElementById("lbZoom").textContent = zoom > 1 ? "縮小" : "放大";',
  '    document.getElementById("lbRot").textContent = rot ? "轉直向" : "轉橫向";',
  '  }',
  '',
  '  function openLb(src){',
  '    zoom = 1; rot = false;',
  '    lbImg.src = src;',
  '    lb.hidden = false;',
  '    if(lbImg.complete) layout(); else lbImg.onload = layout;',
  '  }',
  '  function closeLb(){ lb.hidden = true; lbImg.src = ""; }',
  '',
  '  document.addEventListener("click", function(e){',
  '    var sb = e.target.closest(".shotbtn");',
  '    if(sb){ openLb(sb.querySelector("img").src); return; }',
  '    if(e.target.id === "lbClose"){ closeLb(); return; }',
  '    if(e.target.id === "lbRot"){',
  '      rot = !rot; layout();',
  '      lbStage.scrollTop = (lbStage.scrollHeight - lbStage.clientHeight) / 2;',
  '      lbStage.scrollLeft = (lbStage.scrollWidth - lbStage.clientWidth) / 2;',
  '      return;',
  '    }',
  '    if(e.target.id === "lbZoom"){ zoom = zoom > 1 ? 1 : 2.5; layout(); return; }',
  '  });',
  '  window.addEventListener("resize", function(){ if(!lb.hidden) layout(); });',
  '  document.addEventListener("keydown", function(e){ if(e.key === "Escape" && !lb.hidden) closeLb(); });',
  ''
];

/** 親友頁在對方手機上跑的程式（截圖放大那段跟分享頁共用 LIGHTBOX_JS_LINES） */
var MEMBER_JS = [
  '(function(){',
  '  var API = __API__, TOK = __TOK__;',
  '  function post(body){',
  '    return fetch(API, {method:"POST", redirect:"follow", body: JSON.stringify(body)})',
  '      .then(function(r){ return r.json(); })',
  '      .then(function(j){ if(!j.ok) throw new Error(j.error || "失敗"); return j.data; });',
  '  }',
  '  function q(card, sel){ return card.querySelector(sel); }',
  '  function warn(card, msg){ q(card, ".warn").textContent = msg || ""; }',
  '  function btns(card, show){',
  '    var used = card.classList.contains("used"), pend = card.classList.contains("pend");',
  '    q(card, ".show").hidden = used || pend || show;',
  '    q(card, ".hide").hidden = !show;',
  '    q(card, ".use").hidden = used || pend || !show;',
  '  }',
  '',
  '  // 出示：這時才去拿截圖和序號',
  '  function showCard(card, el){',
  '    el.disabled = true; el.textContent = "載入中…"; warn(card, "");',
  '    post({action:"memberImage", token:TOK, id:card.dataset.id})',
  '      .then(function(d){',
  '        var box = q(card, ".shotbox");',
  '        box.innerHTML = d.dataUrl ? \'<button type="button" class="shotbtn" aria-label="放大這張截圖"><img class="shot" alt="票券截圖" src="\' + d.dataUrl + \'"></button>\'',
  '                                  : \'<p class="noshot">這張沒有截圖，給店員看上面的序號</p>\';',
  '        var c = q(card, ".code"); c.textContent = d.code || ""; c.hidden = !d.code;',
  '        card.classList.add("open"); btns(card, true);',
  '      })',
  '      .catch(function(err){ warn(card, err.message); })',
  '      .then(function(){ el.disabled = false; el.textContent = "出示"; });',
  '  }',
  '  function hideCard(card){',
  '    q(card, ".shotbox").innerHTML = ""; var c = q(card, ".code"); c.textContent = ""; c.hidden = true;',
  '    card.classList.remove("open"); btns(card, false);',
  '  }',
  '',
  '  // 我用掉了：按第一下只是確認，四秒內再按一次才送出',
  '  function useCard(card, el){',
  '    if(el.getAttribute("data-armed") !== "1"){',
  '      el.setAttribute("data-armed", "1"); el.textContent = "確定？再按一次";',
  '      setTimeout(function(){ if(el.getAttribute("data-armed") === "1"){ el.setAttribute("data-armed", "0"); el.textContent = el.getAttribute("data-label"); } }, 4000);',
  '      return;',
  '    }',
  '    el.setAttribute("data-armed", "0"); el.disabled = true; el.textContent = "回報中…"; warn(card, "");',
  '    post({action:"memberUse", token:TOK, id:card.dataset.id})',
  '      .then(function(d){',
  '        var st = q(card, ".stamp"); st.hidden = false;',
  '        if(d.pending){ card.classList.add("pend"); st.textContent = "已通知，他會確認剩多少"; }',
  '        else {',
  '          card.classList.add("used"); st.textContent = "已使用 " + (d.usedAt || "");',
  '          if(d.undoUntil && d.undoUntil > Date.now()){ q(card, ".undo").hidden = false; card.setAttribute("data-undo-until", d.undoUntil); }',
  '        }',
  '        hideCard(card);',
  '      })',
  '      .catch(function(err){ warn(card, err.message); })',
  '      .then(function(){ el.disabled = false; el.textContent = el.getAttribute("data-label"); });',
  '  }',
  '  function undoCard(card, el){',
  '    el.disabled = true; el.textContent = "復原中…"; warn(card, "");',
  '    post({action:"memberUndo", token:TOK, id:card.dataset.id})',
  '      .then(function(){',
  '        card.classList.remove("used"); q(card, ".stamp").hidden = true; el.hidden = true;',
  '        card.removeAttribute("data-undo-until"); btns(card, false);',
  '      })',
  '      .catch(function(err){ warn(card, err.message); })',
  '      .then(function(){ el.disabled = false; el.textContent = "按錯了，復原"; });',
  '  }',
  '',
  '  // 復原時間到了，按鈕自己收起來',
  '  setInterval(function(){',
  '    var cs = document.querySelectorAll(".tk[data-undo-until]");',
  '    for(var i = 0; i < cs.length; i++){',
  '      if(+cs[i].getAttribute("data-undo-until") < Date.now()){ q(cs[i], ".undo").hidden = true; cs[i].removeAttribute("data-undo-until"); }',
  '    }',
  '  }, 15000);',
  '',
  '  document.addEventListener("click", function(e){',
  '    var chip = e.target.closest(".chip");',
  '    if(chip){',
  '      var b = chip.getAttribute("data-b"), cs = document.querySelectorAll(".chip"), ts = document.querySelectorAll(".tk");',
  '      for(var i = 0; i < cs.length; i++) cs[i].classList.toggle("on", cs[i] === chip);',
  '      for(var j = 0; j < ts.length; j++) ts[j].hidden = !!b && ts[j].getAttribute("data-b") !== b;',
  '      return;',
  '    }',
  '    var el = e.target.closest("button"); if(!el) return;',
  '    var card = el.closest(".tk"); if(!card) return;',
  '    if(el.classList.contains("show")) { showCard(card, el); return; }',
  '    if(el.classList.contains("hide")) { hideCard(card); return; }',
  '    if(el.classList.contains("use"))  { useCard(card, el); return; }',
  '    if(el.classList.contains("undo")) { undoCard(card, el); return; }',
  '  });',
  '',
].concat(LIGHTBOX_JS_LINES, ['})();']).join('\n');


/* ══════════════════ 通知 ══════════════════
 *
 *  別人透過分享連結用掉（或復原）票券時通知你。有設定的管道全部都會發：
 *
 *    Discord  DISCORD_WEBHOOK   頻道設定 → 整合 → Webhook 複製的那串網址
 *    LINE     LINE_TOKEN        官方帳號的 Channel access token（長效）
 *             LINE_USER_ID      你自己的 LINE User ID（U 開頭那串）
 *
 *  都沒設就安靜地不通知。任何一個管道失敗都不會影響票券本身。
 *  設好之後在編輯器執行 testNotify 試一下。
 */

/**
 * ev = { kind: 'used' | 'undo' | 'returned' | 'balanceUsed', who, ticket, shareTo }
 *   balanceUsed（v69）= 朋友用了餘額型的券，還沒確認剩多少
 */
/** 你自己在 App 裡標記已使用／改回未使用時，要不要也通知。不想收就改成 false。 */
var NOTIFY_SELF = true;

/** 產生分享連結時的那一則：給了誰、哪幾張、連結、什麼時候失效。 */
/**
 * 現在有幾個連結還在外面、共幾張票券、其中幾張已經被用掉。
 * 收回、過期，或裡面的票都用掉（都過期）的不算。通知底下附這個，一眼就知道手上還有多少在流通。
 */
function shareStats_() {
  var today = today_();
  var byId = {};
  tickets_().forEach(function (t) { byId[t.id] = t; });
  var links = 0, seen = {}, total = 0, used = 0;
  shares_().forEach(function (s) {
    if (s.revoked === 'y') return;
    if (s.expiresAt && s.expiresAt < today) return;
    var ids = (s.ticketIds ? s.ticketIds.split(',') : []).filter(function (id) { return byId[id]; });
    // 裡面的票都被用掉、或都過期了 → 這個連結算結束，不算「在外面」（跟 App 的分享紀錄一樣）
    var live = ids.some(function (id) {
      var t = byId[id];
      return t.status !== 'used' && (!t.expiry || t.expiry >= today);
    });
    if (!live) return;
    links++;
    ids.forEach(function (id) {
      if (seen[id]) return;
      seen[id] = 1;
      total++;
      if (byId[id].status === 'used') used++;
    });
  });
  return { links: links, tickets: total, used: used, left: total - used };
}

/** 一行字的現況：3 個連結在外面，共 5 張，已用 2 張，還有 3 張沒用 */
function shareStatsLine_() {
  try {
    var st = shareStats_();
    if (!st.links) return '目前沒有連結在外面';
    return st.links + ' 個連結在外面，共 ' + st.tickets + ' 張' +
           (st.used ? '，已用 ' + st.used + ' 張，還有 ' + st.left + ' 張沒用' : '，都還沒用');
  } catch (e) { return ''; }
}

/** 某個連結裡還有幾張沒用掉 */
function shareLeftLine_(token, known) {
  try {
    var s = known || shareRow_(token);
    if (!s) return '';
    var ids = {}, n = 0, used = 0;
    (s.ticketIds ? s.ticketIds.split(',') : []).forEach(function (id) { ids[id] = 1; });
    tickets_().forEach(function (t) {
      if (!ids[t.id]) return;
      n++;
      if (t.status === 'used') used++;
    });
    return n + ' 張，已用 ' + used + ' 張，還有 ' + (n - used) + ' 張沒用';
  } catch (e) { return ''; }
}

function notifyShare_(ev) {
  var keep = String(ev.keep || '').split(',');      // v76：自己這邊也留著的，誰先用誰先按
  var list = (ev.tickets || []).map(function (t) { return '・' + ticketLabel_(t) + (keep.indexOf(t.id) >= 0 ? '（你這邊也留著）' : ''); });
  var shown = list.slice(0, 10).join('\n');
  if (list.length > 10) shown += '\n・…還有 ' + (list.length - 10) + ' 張';

  var line = '你在 ' + nowText_() + ' 分享了 ' + (ev.tickets || []).length + ' 張票券' +
             (ev.who ? '給「' + ev.who + '」' : '');

  var fields = [];
  if (shown) fields.push({ name: '票券', value: shown.slice(0, 1000) });
  if (ev.expiresAt) fields.push({ name: '連結失效日', value: ev.expiresAt, inline: true });
  if (ev.note) fields.push({ name: '給對方的話', value: String(ev.note).slice(0, 200), inline: true });
  if (ev.url) fields.push({ name: '連結', value: ev.url });
  var now = shareStatsLine_();
  if (now) fields.push({ name: '目前狀況', value: now });

  var body = {
    username: '票券',
    content: line.slice(0, 1900),
    allowed_mentions: { parse: [] },
    embeds: [{ color: 0x6A35A8, fields: fields, timestamp: new Date().toISOString() }]
  };
  postDiscord_(body);
  notifyLine_(line + (ev.url ? '\n' + ev.url : ''));
  return true;
}

function notify_(ev) {
  if (ev.kind === 'share') return notifyShare_(ev);
  var t = ev.ticket || {};
  var when = nowText_();
  var who = ev.self ? '你自己' : ev.who;
  var line = ev.kind === 'used'     ? who + ' 於 ' + when + ' 使用了「' + ticketLabel_(t) + '」'
           : ev.kind === 'balanceUsed' ? who + ' 於 ' + when + ' 用了「' + ticketLabel_(t) + '」（餘額型，還沒確認剩多少）'
           : ev.kind === 'returned' ? who + ' 於 ' + when + ' 把「' + ticketLabel_(t) + '」還給你了'
           :                          who + ' 於 ' + when + ' 復原了「' + ticketLabel_(t) + '」';
  var extra = ev.kind === 'used'     ? (t.code ? '序號 ' + t.code : '')
            : ev.kind === 'balanceUsed' ? '請打開 App 裡這張票券，確認剩餘金額' +
                                          (t.balance ? '（上次確認剩 $' + t.balance + '）' : '')
            : ev.kind === 'returned' ? '他用不到，這張回到你的「可用」了' +
                                       (t.expiry ? '（' + t.expiry.slice(5).replace('-', '/') + ' 到期）' : '')
            : (ev.self ? '這張改回未使用了，還能用' : '剛剛是按錯，這張還能用');

  notifyDiscord_(ev, line, extra);
  notifyLine_(line + (extra ? '\n' + extra : ''));
}

function notifyDiscord_(ev, line, extra) {
  var t = ev.ticket || {};
  var fields = [];
  var back = ev.kind === 'returned';
  if (t.code)   fields.push({ name: '序號', value: t.code, inline: true });
  if ((ev.kind === 'used' || back) && t.expiry) fields.push({ name: '到期', value: t.expiry, inline: true });
  if (ev.self) fields.push({ name: '來源', value: '自己在 App 裡標記', inline: true });
  if (ev.member) fields.push({ name: '來源', value: '親友連結', inline: true });   // v75
  if (ev.shareTo && ev.shareTo !== ev.who) fields.push({ name: '當初分享給', value: ev.shareTo, inline: true });
  if (ev.token) {
    if (back && !ev.leftCount) {
      fields.push({ name: '這個連結', value: '裡面的票券都還回來了，連結已經關閉' });
    } else {
      var left = shareLeftLine_(ev.token, ev.share);
      if (left) fields.push({ name: '這個連結', value: left });
      fields.push({ name: '連結', value: shareUrl_(ev.token) });
    }
  }

  // App 有帶「按下的時間」就在最底下小字註明花了幾秒送到 Discord。
  // 手機如果比這個晚很多才跳，延遲就不在票券這邊，而是 Discord 推播或手機省電設定。
  var footer = null, lag = ev.at ? (Date.now() - Number(ev.at)) / 1000 : NaN;
  if (lag >= 0 && lag < 600) footer = { text: '從按下到送出 ' + lag.toFixed(1) + ' 秒' };

  var body = {
    username: '票券',
    content: line.slice(0, 1900),
    // 名字是對方自己打的，萬一打了 @everyone 也不要真的去 tag 整個伺服器
    allowed_mentions: { parse: [] },
    embeds: (fields.length || footer || (extra && ev.kind !== 'used')) ? [{
      description: ev.kind !== 'used' ? extra : undefined,
      color: ev.kind === 'used' ? 0x6A35A8 : ev.kind === 'balanceUsed' ? 0xE0A21B : back ? 0x2A9D8F : 0x847A96,
      fields: fields,
      footer: footer || undefined,
      timestamp: new Date().toISOString()
    }] : []
  };

  return postDiscord_(body);
}

/*
 * v56：電腦上雙擊「一鍵部署」也發 Discord 通知。
 * deploy.js 跑完（成功或失敗）打 deployNotice。這個動作不用存取碼，所以：
 *   訊息內容全部由後端自己組，對方只能給「成功或失敗、哪一步、版本號、電腦名稱」這幾個固定欄位；
 *   一分鐘最多一則，免得被拿來洗版。
 * 手機「部署最新版」那條路照舊由部署小幫手回報（deployReport），deploy.js 不會重複發。
 */
var DEPLOY_NOTICE_WHY = {
  setup: '還沒做過第一次設定', files: 'web 資料夾少了檔案', version: '版本號對不起來',
  push: '上傳 Code.gs 失敗', deploy: '後端更新部署失敗', cf: '上傳 Cloudflare 失敗',
  check: '已經上傳，但網站檢查沒通過', error: '發生沒預料到的錯誤',
  stale: '這個資料夾比線上舊（OneDrive 可能還沒同步完），沒有部署'     // v69：deploy.js 部署前發現線上比本機新
};
function deployNotice_(req) {
  var p = PropertiesService.getScriptProperties();
  var before = Number(p.getProperty('DEPLOY_NOTICE_LAST') || 0);
  if (Date.now() - before < 60000) return { ignored: true };
  p.setProperty('DEPLOY_NOTICE_LAST', String(Date.now()));

  var ok = req.ok === true;
  var site = /^\d{1,4}$/.test(String(req.site)) ? 'v' + req.site : '';
  var backend = /^\d{1,6}$/.test(String(req.backend)) ? String(req.backend) : '';
  var host = String(req.host || '').replace(/[^A-Za-z0-9_-]/g, '').slice(0, 30);
  var why = DEPLOY_NOTICE_WHY[req.why] || (ok ? '' : DEPLOY_NOTICE_WHY.error);
  var last = {
    id: 'pc-' + Date.now(), ok: ok, at: Date.now(), by: '電腦', host: host, site: site, backend: backend,
    text: ok ? '網站版本：' + site : '✗ ' + why
  };
  if (req.why === 'check') last.head = '⚠ 已經上傳，但網站檢查沒通過' + (site ? '（網站 ' + site + '）' : '');
  p.setProperty('DEPLOY_LAST', JSON.stringify(last));     // App「設定 → 部署」也看得到
  LAST_DISCORD_ = null; LAST_LINE_ = null;
  notifyDeploy_(last);
  return notifyResult_();
}

/*
 * v56：手機或電腦上的 App 換了版本（強制更新、或自己更新）就發一則 Discord。
 * App 會在啟動時發現「這台上次是 v55、現在是 v56」或「剛剛按了強制更新」才打這個動作。
 */
function appUpdated_(req) {
  function v(x) { var m = /(\d{1,4})\s*$/.exec(String(x == null ? '' : x)); return m ? 'v' + m[1] : '?'; }
  var from = v(req.from), to = v(req.to), code = 'v' + CODE_VERSION;
  var dev = String(req.device || '一台裝置').replace(/[@`*_~|<>]/g, '').slice(0, 40);
  var forced = !!req.forced;
  var head = forced
    ? '🔄 強制更新：' + dev + (from === to ? '，已經是 ' + to : from === '?' ? '，現在是 ' + to : '，' + from + ' → ' + to)
    : from === '?' ? '📱 App ' + to + '：' + dev + '（這台第一次記版本，原本的版本不明）'
    : '📱 App 更新到 ' + to + '：' + dev + '（原本 ' + from + '）';
  var desc = '網站 ' + to + '、後端 ' + code +
             (to !== code ? '\n⚠ 兩邊版本不一樣，還差一邊沒更新（手機部署或電腦一鍵部署再跑一次）' : '');
  LAST_DISCORD_ = null; LAST_LINE_ = null;
  postDiscord_({
    username: '票券', allowed_mentions: { parse: [] }, content: head,
    embeds: [{ color: to !== code ? 0xC4372B : 0x6A35A8, description: desc, timestamp: new Date().toISOString() }]
  });
  notifyLine_(head + '\n' + desc);
  return notifyResult_();
}

/** 剛剛那一則沒送到的原因 */
function lastWhy_() {
  return (LAST_DISCORD_ && LAST_DISCORD_.why) || (LAST_LINE_ && LAST_LINE_.why) || '';
}

/** 每天早上的提醒排程在不在（App「最近送出的通知」會提示） */
function reminderState_() {
  var off = P_('REMINDER_OFF') === '1', has = null;
  try {
    has = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'dailyReminder'; });
  } catch (e) {}
  return { off: off, trigger: has, last: P_('REMIND_LAST') || '' };
}
/** App 按「裝回每天早上的提醒」 */
function fixReminder_() {
  if (P_('REMINDER_OFF') === '1') return { off: true };
  ensureReminderTrigger_();
  return reminderState_();
}

/**
 * 真正把一則訊息送去 Discord。沒設 webhook 或網址不對就安靜跳過。
 *   opt.noQueue：送不到也不要排進補送（測試通知、延遲測試、預覽用）
 *   opt.noFlush：送到了也不要順便補送之前沒送到的（延遲測試用，免得量到別的）
 * 送不到會先等一下重試一次；還是不行就排進補送（見下面「通知補送」），之後自動再送。
 */
function postDiscord_(body, opt) {
  opt = opt || {};
  var url = P_('DISCORD_WEBHOOK');
  if (!url) return false;
  var what = notifyLabel_(body);
  if (!discordUrlOk_(url)) {
    Logger.log('Discord 通知沒送出：DISCORD_WEBHOOK 看起來不是 Discord 的 webhook 網址');
    LAST_DISCORD_ = { ok: false, code: 0, ms: 0, at: '', atMs: 0, why: 'DISCORD_WEBHOOK 看起來不是 Discord 的 webhook 網址' };
    logNotify_('Discord', what, LAST_DISCORD_);
    return false;
  }
  var r = deliverDiscord_(url, body, true);
  if (!r.ok && r.retry && !opt.noQueue) r.queued = queueNotify_(body, what, r.why);
  LAST_DISCORD_ = r;
  logNotify_('Discord', what, r);
  if (r.ok && !opt.noFlush && notifyQueueWaiting_()) {
    try { flushNotifyQueue_(NOTIFY_FLUSH_AFTER_OK); } catch (e) { Logger.log('補送出錯：' + e.message); }
  }
  return r.ok;
}
function discordUrlOk_(url) {
  return /^https:\/\/(canary\.|ptb\.)?(discord|discordapp)\.com\/api\/webhooks\//.test(url);
}

/**
 * 送一次，不重試、不記紀錄。回傳：
 *   ok、code、ms、at（Discord 收到的時間）、why（沒送到的原因）
 *   retry：之後再送有沒有機會成功（格式錯誤 400 就沒有）
 *   soon：值得「現在等幾秒再試一次」的話是秒數，不值得是 -1
 */
function sendDiscordOnce_(url, body) {
  var t = Date.now();
  try {
    var res = UrlFetchApp.fetch(url + (url.indexOf('?') < 0 ? '?wait=true' : '&wait=true'), {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify(body)
    });
    return readDiscordReply_(res.getResponseCode(), res.getContentText(), Date.now() - t);
  } catch (e) {
    Logger.log('Discord 通知沒送出：' + e.message);
    return { ok: false, code: 0, ms: Date.now() - t, at: '', atMs: 0, why: '連不上 Discord：' + e.message,
             retry: true, soon: 1.5, blocked: false };
  }
}

/**
 * 看 Discord 回了什麼。直送和經 Cloudflare 轉送都用這個。回傳：
 *   ok、code、ms、at（Discord 收到的時間）、why（沒送到的原因）
 *   retry：之後再送有沒有機會成功（格式錯誤 400 就沒有）
 *   soon：值得「現在等幾秒再試一次」的話是秒數，不值得是 -1
 *   blocked：Discord（Cloudflare）擋掉了送出的這台伺服器（429 封鎖頁、1015、403）
 */
function readDiscordReply_(code, text, ms) {
  var r = { ok: false, code: code, ms: ms || 0, at: '', atMs: 0, why: '', retry: false, soon: -1, blocked: false };
  if (code === 200 || code === 204) {
    r.ok = true;
    // 記下 Discord 自己說「幾點收到」，用來判斷延遲是不是發生在 Discord 之後
    if (code === 200) {
      try {
        var m = JSON.parse(text);
        var at = Date.parse(String(m.timestamp || '').replace(/(\.\d{3})\d+/, '$1'));
        if (!isNaN(at)) {
          r.atMs = at;
          r.at = Utilities.formatDate(new Date(at), Session.getScriptTimeZone(), 'HH:mm:ss');
        }
      } catch (e) {}
    }
    return r;
  }
  var raw = String(text || '').slice(0, 300), detail = raw.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();
  var j = null;
  try { j = JSON.parse(raw); } catch (e) {}          // Discord 回的是 JSON：{"message": "Unknown Webhook", "code": 10015}
  if (j && j.message) detail = String(j.message) + (j.errors ? ' ' + JSON.stringify(j.errors) : '');
  var why;
  if (code === 404 || code === 401) {
    why = '這個 webhook 已經被刪掉或網址不對，到 Discord 重新建一個。';
    r.retry = true;                                  // 換了新的 webhook 之後，補送會用新的網址
  } else if (code === 429) {
    var after = j && Number(j.retry_after);
    if (j && !isNaN(after)) {
      why = 'Discord 說發太快了，等一下就好。';
      if (after <= 5) r.soon = Math.max(0.5, after);
    } else {
      why = 'Discord 暫時擋掉 Google 伺服器送來的通知。';   // 不是 JSON，通常是 Cloudflare 的封鎖頁（1015）
      r.blocked = true;
    }
    r.retry = true;
  } else if (code === 403) {
    why = 'Discord 拒絕了這台伺服器。';
    r.retry = true; r.blocked = true;
  } else if (code === 400) {
    why = 'Discord 不收這則的格式。';
  } else if (code >= 500) {
    why = 'Discord 那邊暫時出問題。';
    r.retry = true; r.soon = 1.5;
  } else {
    why = '';
    r.retry = true;
  }
  Logger.log('Discord 通知沒送出（' + code + '）：' + why + raw);
  r.why = why + '（' + code + (detail ? ' ' + detail.slice(0, 100) : '') + '）';
  return r;
}

/*
 * v57：Google 的伺服器被 Discord 擋掉時（429、error code 1015），改由自己的 Cloudflare 網站轉送。
 *   Apps Script 的 UrlFetch 走 Google 共用的 IP，很多人一起用，Discord（Cloudflare）不時會整批擋掉；
 *   Cloudflare 網站（_worker.js 的 /relay/discord）送過去就不受影響。
 *   防止別人拿這個轉送亂發：每一則都帶一個一次性的號碼，_worker.js 會回頭問後端
 *   「這一則真的是你剛剛要送的嗎？」（relayCheck，比對內容的雜湊），對得上才送。
 *   轉送的網址用 SHORT_BASE 的網域（例如 https://你的網站.pages.dev/relay/discord），也可以用指令碼屬性 DISCORD_RELAY_URL 指定。
 *
 * v58：有轉送網址就一律先走轉送（不等 Google 被擋才轉），轉送這條路壞了才改直接送。
 *   想改回「先直接送、被擋才轉送」：指令碼屬性 DISCORD_VIA 設成 direct（那時被擋過一次，30 分鐘內先走轉送）。
 */
var RELAY_PREFER_MS = 30 * 60000;
function relayUrl_() {
  var own = P_('DISCORD_RELAY_URL');
  if (own) return /^https:\/\//.test(own) ? own : '';
  var m = /^(https:\/\/[^\/?#]+)\//.exec(P_('SHORT_BASE') || '');
  return m ? m[1] + '/relay/discord' : '';
}
function hex_(bytes) {
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}
/** 經 Cloudflare 轉送一次；沒有轉送網址就回 null */
function sendViaRelay_(url, body) {
  var relay = relayUrl_();
  if (!relay) return null;
  var nonce = Utilities.getUuid();
  var text = JSON.stringify({ nonce: nonce, url: url, payload: body });
  var digest = hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8));
  try { CacheService.getScriptCache().put('RELAY_' + nonce, digest, 120); } catch (e) {}
  var t = Date.now(), r;
  try {
    var res = UrlFetchApp.fetch(relay, {
      method: 'post', contentType: 'text/plain;charset=utf-8', muteHttpExceptions: true, payload: text
    });
    var code = res.getResponseCode(), reply = res.getContentText();
    var mine = null;
    try { mine = JSON.parse(reply); } catch (e) {}
    if (mine && mine.relay) {                          // _worker.js 自己擋下來的，不是 Discord 回的
      r = { ok: false, code: code, ms: Date.now() - t, at: '', atMs: 0, retry: true, soon: -1, blocked: false, relayErr: true,
            why: mine.relay === 'denied' ? 'Cloudflare 轉送沒認出這一則（' + code + '）'
               : mine.relay === 'exec' ? 'Cloudflare 轉送連不到後端（' + code + '）'
               : mine.relay === 'discord' ? 'Cloudflare 轉送連不到 Discord（' + code + '）'
               : 'Cloudflare 轉送不收（' + code + ' ' + mine.relay + '）' };
    } else if (code === 404 && !/webhook/i.test(reply)) {
      r = { ok: false, code: code, ms: Date.now() - t, at: '', atMs: 0, retry: true, soon: -1, blocked: false, relayErr: true,
            why: '網站上還沒有轉送功能（_worker.js 要更新到 v57 以後）' };
    } else {
      r = readDiscordReply_(code, reply, Date.now() - t);
    }
  } catch (e) {
    r = { ok: false, code: 0, ms: Date.now() - t, at: '', atMs: 0, retry: true, soon: -1, blocked: false, relayErr: true,
          why: '連不上 Cloudflare 轉送：' + e.message };
  }
  r.via = 'relay';
  return r;
}
/** _worker.js 問：「這一則真的是你剛剛要送的嗎？」（不用存取碼；號碼只能用一次） */
function relayCheck_(req) {
  var nonce = String((req && req.nonce) || '');
  if (!/^[0-9a-f-]{20,64}$/i.test(nonce)) return { ok: false };
  var c = CacheService.getScriptCache(), k = 'RELAY_' + nonce;
  var want = c.get(k);
  if (!want) return { ok: false };
  c.remove(k);
  return { ok: want === String(req.digest || '') };
}
function preferRelay_() {
  return Number(P_('DISCORD_RELAY_UNTIL') || 0) > Date.now();
}
function markRelay_() {
  var until = String(Date.now() + RELAY_PREFER_MS);
  try { PropertiesService.getScriptProperties().setProperty('DISCORD_RELAY_UNTIL', until); } catch (e) {}
  if (PROPS_) PROPS_.DISCORD_RELAY_UNTIL = until;
}
/** 先走 Cloudflare 轉送嗎？（有轉送網址、沒設 DISCORD_VIA=direct；或最近 Google 被擋過） */
function relayFirst_() {
  if (!relayUrl_()) return false;
  return String(P_('DISCORD_VIA') || '').toLowerCase() !== 'direct' || preferRelay_();
}
/** 直接從 Google 送；暫時性的問題等一下再試一次 */
function sendDirect_(url, body, canWait) {
  var r = sendDiscordOnce_(url, body);
  if (!r.ok && canWait && r.soon >= 0 && !r.blocked) {
    Utilities.sleep(Math.round(r.soon * 1000));
    var why1 = r.why;
    r = sendDiscordOnce_(url, body);
    r.retried = true;
    if (!r.ok && !r.why) r.why = why1;
  }
  return r;
}
/**
 * 送一則到 Discord。
 *   預設（v58）：先經 Cloudflare 轉送；轉送這條路有問題才改直接送。
 *   DISCORD_VIA=direct：先直接送；被擋才改轉送。
 * Discord 自己說不行的（webhook 被刪 404、格式錯 400）不換路，換了也一樣。canWait=false 時不等（補送用）。
 */
function deliverDiscord_(url, body, canWait) {
  if (relayFirst_()) {
    var a = sendViaRelay_(url, body);
    if (!a.ok && canWait && a.soon >= 0) {             // Discord 說等一下（經轉送也一樣）
      Utilities.sleep(Math.round(a.soon * 1000));
      var whyA = a.why;
      a = sendViaRelay_(url, body);
      a.retried = true;
      if (!a.ok && !a.why) a.why = whyA;
    }
    if (a.ok) return a;
    if (!a.relayErr && (!a.retry || a.code === 404 || a.code === 401)) return a;
    var d = sendDirect_(url, body, canWait);
    if (d.ok) { d.fb = a.why; return d; }
    d.why = a.why + '；改直接送也沒成功：' + d.why;
    d.retry = d.retry || a.retry;
    return d;
  }
  var r = sendDirect_(url, body, canWait);
  if (r.ok || !r.retry || r.code === 404 || r.code === 401) return r;   // webhook 本身壞了，轉送也沒用
  if (r.blocked) markRelay_();
  var rr = sendViaRelay_(url, body);
  if (rr && rr.ok) { rr.direct = r.why; return rr; }
  if (rr) r.why += '；改由 Cloudflare 轉送也沒成功：' + rr.why;
  return r;
}

/*
 * 通知補送：送不到的通知先存起來，之後自動再送，補送的那則會註明原本幾點要送、晚了多久。
 *   什麼時候補送：每 15 分鐘的自動匯入、每天早上的到期提醒、下一則通知成功送到時、
 *                 App「設定 → 最近送出的通知 → 現在補送」
 *   最多排 NOTIFY_QUEUE_MAX 則，超過 NOTIFY_QUEUE_DAYS 天還送不到就放棄（會寫進通知紀錄）
 *   存在指令碼屬性：NOTIFY_QUEUE（清單）、NOTIFY_Q_<id>（每一則的內容）、NOTIFY_FLUSHING（正在補送的記號）
 */
var NOTIFY_QUEUE_MAX = 10;
var NOTIFY_QUEUE_DAYS = 3;
var NOTIFY_FLUSH_AFTER_OK = 3;     // 一則通知剛送到時，順便補送幾則（不要拖太久）
var NOTIFY_Q_BYTES = 8000;         // 指令碼屬性一個值最多 9KB，留點餘裕

function readQueue_() {
  var q = [];
  try { q = JSON.parse(PropertiesService.getScriptProperties().getProperty('NOTIFY_QUEUE') || '[]'); } catch (e) {}
  return Array.isArray(q) ? q : [];
}
function writeQueue_(q) {
  var p = PropertiesService.getScriptProperties();
  if (q.length) p.setProperty('NOTIFY_QUEUE', JSON.stringify(q));
  else p.deleteProperty('NOTIFY_QUEUE');
  if (PROPS_) PROPS_.NOTIFY_QUEUE = q.length ? JSON.stringify(q) : '';
}
/** 有沒有等著補送的（網頁請求用開頭讀好的設定，排程直接讀） */
function notifyQueueWaiting_() {
  var v = P_('NOTIFY_QUEUE');
  return !!v && v !== '[]';
}
/** 改補送清單時排隊，免得兩個請求同時改、蓋掉對方（拿不到鎖就照做，只是清單） */
function queueLock_(fn) {
  if (LOCK_DEPTH_ > 0) return fn();
  var lock = LockService.getScriptLock();
  var got = false;
  try { got = lock.tryLock(3000); } catch (e) {}
  try { return fn(); }
  finally { if (got) lock.releaseLock(); }
}
function utf8Len_(s) {
  var n = 0;
  for (var i = 0; i < s.length; i++) {
    var c = s.charCodeAt(i);
    n += c < 0x80 ? 1 : c < 0x800 ? 2 : (c >= 0xD800 && c <= 0xDBFF) ? (i++, 4) : 3;
  }
  return n;
}
/** 太長的通知先縮短小卡內容，再不行就只留第一行，好放進指令碼屬性 */
function fitQueueBody_(body) {
  var b = JSON.parse(JSON.stringify(body));
  var s = JSON.stringify(b);
  var e = b.embeds && b.embeds[0];
  while (utf8Len_(s) > NOTIFY_Q_BYTES && e && e.description && e.description.length > 200) {
    e.description = e.description.slice(0, Math.floor(e.description.length / 2)) + '\n…（補送時縮短了）';
    s = JSON.stringify(b);
  }
  if (utf8Len_(s) > NOTIFY_Q_BYTES) {
    b = { username: b.username, allowed_mentions: { parse: [] }, content: String(b.content || '').slice(0, 1500) };
    s = JSON.stringify(b);
  }
  return s;
}
/** 排進補送；回傳 true 表示排好了 */
function queueNotify_(body, what, why) {
  try {
    var id = Date.now().toString(36) + Math.random().toString(36).slice(2, 5);
    var data = fitQueueBody_(body);
    var dropped = [];
    queueLock_(function () {
      var p = PropertiesService.getScriptProperties();
      var q = readQueue_();
      p.setProperty('NOTIFY_Q_' + id, data);
      q.push({ id: id, at: Date.now(), what: String(what || '').slice(0, 60), tries: 0, why: String(why || '').slice(0, 120) });
      while (q.length > NOTIFY_QUEUE_MAX) {
        var old = q.shift();
        p.deleteProperty('NOTIFY_Q_' + old.id);
        dropped.push(old);
      }
      writeQueue_(q);
    });
    dropped.forEach(function (x) {
      logNotify_('Discord', x.what, { ok: false, why: '等著補送的太多了，最舊的這則放棄', drop: true, orig: x.at });
    });
    return true;
  } catch (e) {
    Logger.log('排進補送失敗：' + e.message);
    return false;
  }
}
function lagText_(ms) {
  var m = Math.max(1, Math.round(ms / 60000));
  if (m < 60) return m + ' 分鐘';
  var h = Math.floor(m / 60);
  if (h < 48) return h + ' 小時' + (m % 60 ? ' ' + (m % 60) + ' 分' : '');
  return Math.round(h / 24) + ' 天';
}
function mmddhhmm_(ms) {
  return Utilities.formatDate(new Date(ms), Session.getScriptTimeZone(), 'MM/dd HH:mm');
}
/** 補送的那則，最上面加一行說明原本幾點要送 */
function resendBody_(body, at) {
  var b = JSON.parse(JSON.stringify(body));
  var note = '（補送｜原本 ' + mmddhhmm_(at) + ' 要送，晚了 ' + lagText_(Date.now() - at) + '）';
  b.content = (note + '\n' + String(b.content || '')).slice(0, 2000);
  return b;
}
/**
 * 把等著補送的依序再送一次，最多 max 則。遇到送不到就停（Discord 還在擋，別一直打）。
 * 回傳 { sent, left, why }
 */
function flushNotifyQueue_(max) {
  var out = { sent: 0, left: 0, why: '' };
  var url = P_('DISCORD_WEBHOOK');
  var p = PropertiesService.getScriptProperties();
  if (!readQueue_().length) return out;
  if (!url || !discordUrlOk_(url)) { out.left = readQueue_().length; out.why = '還沒設定 Discord'; return out; }

  // 同一時間只讓一個補送在跑，免得同一則送兩次
  var mine = String(Date.now()) + Math.random().toString(36).slice(2, 5);
  var busy = false;
  queueLock_(function () {
    var cur = p.getProperty('NOTIFY_FLUSHING');
    if (cur && Date.now() - Number(cur.split('-')[0]) < 3 * 60000) { busy = true; return; }
    p.setProperty('NOTIFY_FLUSHING', Date.now() + '-' + mine);
  });
  if (busy) { out.left = readQueue_().length; out.why = '另一個補送正在跑'; return out; }

  try {
    // 先丟掉太舊的
    var stale = [];
    queueLock_(function () {
      var q = readQueue_(), keep = [];
      q.forEach(function (x) {
        if (Date.now() - x.at > NOTIFY_QUEUE_DAYS * 86400000) { stale.push(x); p.deleteProperty('NOTIFY_Q_' + x.id); }
        else keep.push(x);
      });
      if (stale.length) writeQueue_(keep);
    });
    stale.forEach(function (x) {
      logNotify_('Discord', x.what, { ok: false, why: '超過 ' + NOTIFY_QUEUE_DAYS + ' 天還送不到，放棄補送', drop: true, orig: x.at });
    });

    var todo = readQueue_().slice(0, max || NOTIFY_QUEUE_MAX);
    for (var i = 0; i < todo.length; i++) {
      var x = todo[i];
      var raw = p.getProperty('NOTIFY_Q_' + x.id);
      var body = null;
      try { body = JSON.parse(raw); } catch (e) {}
      if (!body) { removeQueued_(x.id); continue; }
      var r = deliverDiscord_(url, resendBody_(body, x.at), false);
      r.orig = x.at;
      // 補送成功一定記；補送失敗只在原因變了才記，免得同一個原因每 15 分鐘洗一次紀錄
      if (r.ok || r.why !== x.why) logNotify_('Discord', x.what, r);
      if (r.ok) { removeQueued_(x.id); out.sent++; continue; }
      out.why = r.why;
      queueLock_(function () {
        var q = readQueue_();
        q.forEach(function (y) { if (y.id === x.id) { y.tries = (y.tries || 0) + 1; y.why = String(r.why || '').slice(0, 120); } });
        if (!r.retry) q = q.filter(function (y) { if (y.id === x.id) { p.deleteProperty('NOTIFY_Q_' + y.id); return false; } return true; });
        writeQueue_(q);
      });
      break;
    }
  } finally {
    queueLock_(function () {
      var cur = p.getProperty('NOTIFY_FLUSHING');
      if (cur && cur.split('-')[1] === mine) p.deleteProperty('NOTIFY_FLUSHING');
    });
  }
  out.left = readQueue_().length;
  FLUSHED_ += out.sent;
  return out;
}
function removeQueued_(id) {
  queueLock_(function () {
    var p = PropertiesService.getScriptProperties();
    p.deleteProperty('NOTIFY_Q_' + id);
    writeQueue_(readQueue_().filter(function (y) { return y.id !== id; }));
  });
}
/** App「現在補送」 */
function notifyFlushNow_() {
  var r = flushNotifyQueue_(NOTIFY_QUEUE_MAX);
  return r;
}
/** 排程用：自動匯入、到期提醒開頭先補送 */
function flushQuietly_() {
  try {
    if (!readQueue_().length) return;
    var r = flushNotifyQueue_(NOTIFY_QUEUE_MAX);
    if (r.sent || r.left) Logger.log('通知補送：送出 ' + r.sent + ' 則，還剩 ' + r.left + ' 則' + (r.why ? '（' + r.why + '）' : ''));
  } catch (e) { Logger.log('通知補送出錯：' + e.message); }
}

/*
 * 通知紀錄：每一則有沒有送到、Discord 幾點收到、沒送到的原因，記最近 15 則。
 * App 按的動作（網頁請求）Google 不會留「執行項目」的紀錄，所以自己記，
 * 在 App「設定 → 最近送出的通知」看得到。
 */
var NOTIFY_LOG_MAX = 15;
function logNotify_(ch, what, res) {
  try {
    var p = PropertiesService.getScriptProperties();
    var list = [];
    try { list = JSON.parse(p.getProperty('NOTIFY_LOG') || '[]'); } catch (e) {}
    if (!Array.isArray(list)) list = [];
    list.unshift({
      at: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd HH:mm:ss'),
      ch: ch, what: String(what || '').slice(0, 60), ok: !!res.ok, code: res.code || 0,
      rx: res.at || '', why: res.ok ? '' : String(res.why || '').slice(0, 160),
      q: res.queued ? 1 : undefined,                  // 沒送到，排進補送了
      re: res.orig ? mmddhhmm_(res.orig) : undefined, // 這是補送，原本幾點要送
      rt: res.retried && res.ok ? 1 : undefined,      // 重試一次才送到
      via: res.via === 'relay' && res.ok ? 'relay' : undefined,   // 經 Cloudflare 轉送送到的
      dw: res.via === 'relay' && res.ok && res.direct ? 1 : undefined, // 先直接送被擋，才改轉送
      fb: res.fb ? String(res.fb).slice(0, 120) : undefined,          // 轉送沒成功，改直接送才送到
      drop: res.drop ? 1 : undefined,                 // 放棄補送
      info: res.info ? String(res.info).slice(0, 120) : undefined   // 不是通知，只是記一下（例如每日提醒沒東西發）
    });
    p.setProperty('NOTIFY_LOG', JSON.stringify(list.slice(0, NOTIFY_LOG_MAX)));
  } catch (e) { Logger.log('通知紀錄寫不進去：' + e.message); }
}
/** 一則通知的標題：訊息第一行（沒有的話拿小卡的第一行）。 */
function notifyLabel_(body) {
  var c = String((body && body.content) || '').split('\n')[0];
  var e = body && body.embeds && body.embeds[0];
  if (!c && e) c = String(e.title || e.description || '').split('\n')[0];
  return c.replace(/\*\*/g, '').trim().slice(0, 60);
}
/** 剛剛那一則的結果，回給 App 顯示（有設 Discord 看 Discord，只設 LINE 就看 LINE）。 */
function notifyResult_() {
  var d = LAST_DISCORD_, l = LAST_LINE_;
  if (d) return { ok: !!d.ok, at: d.at || '', why: d.ok ? '' : (d.why || '沒送出'), queued: !!d.queued };
  if (l) return { ok: !!l.ok, at: '', why: l.ok ? '' : (l.why || 'LINE 沒送出'), ch: 'LINE' };
  return { none: true };
}
/** App「設定 → 最近送出的通知」 */
function notifyLog_() {
  var list = [];
  try { list = JSON.parse(PropertiesService.getScriptProperties().getProperty('NOTIFY_LOG') || '[]'); } catch (e) {}
  var queue = readQueue_().map(function (x) {
    return { at: mmddhhmm_(x.at), what: x.what, tries: x.tries || 0, why: x.why || '' };
  });
  return { log: Array.isArray(list) ? list : [], queue: queue, discord: !!P_('DISCORD_WEBHOOK'),
           line: !!(P_('LINE_TOKEN') && P_('LINE_USER_ID')), reminder: reminderState_() };
}

function notifyLine_(text) {
  var token = P_('LINE_TOKEN'), to = P_('LINE_USER_ID');
  if (!token || !to) return false;
  try {
    var res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({ to: to, messages: [{ type: 'text', text: String(text).slice(0, 4900) }] })
    });
    var code = res.getResponseCode();
    LAST_LINE_ = { ok: code === 200, code: code, at: '', why: '' };
    if (code !== 200) {
      var why = res.getContentText().slice(0, 300);
      if (code === 401) why = 'LINE_TOKEN 不對或過期了。' + why;
      if (code === 429) why = '這個月的免費推播額度用完了（免費方案每月 200 則）。' + why;
      if (code === 400 && /to/i.test(why)) why = 'LINE_USER_ID 不對，或你還沒把官方帳號加為好友。' + why;
      Logger.log('LINE 通知沒送出（' + code + '）：' + why);
      LAST_LINE_.why = why + '（' + code + '）';
    }
    logNotify_('LINE', String(text).split('\n')[0], LAST_LINE_);
    return code === 200;
  } catch (e) {
    Logger.log('LINE 通知沒送出：' + e.message);   // 通知失敗絕不能讓「用掉」這個動作跟著失敗
    LAST_LINE_ = { ok: false, code: 0, at: '', why: '連不上 LINE：' + e.message };
    logNotify_('LINE', String(text).split('\n')[0], LAST_LINE_);
    return false;
  }
}

function ticketLabel_(t) {
  var m = String(t.merchant || '').trim();
  var ti = String(t.title || '').trim();
  // 券名本身常常已經帶了商家（「清心福全 烏龍綠茶」），就不要再重複一次
  if (m && ti && ti.indexOf(m) === 0) return ti;
  return [m, ti].filter(function (x) { return x; }).join(' ') || '一張票券';
}

function nowText_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd HH:mm:ss');
}

/**
 * 一次修好以前存成「7.3E+15」的序號：能從 rawText 撈回完整數字的就補回去，
 * 撈不回來的清空並標成 low，免得兩張不同的券被當成同一張。
 * 在編輯器選這個函式按執行，結果會寫在紀錄裡。
 */
function fixSciCodes() { return withLock_(fixSciCodesRun_); }   // 逐列改寫，跟其他寫入排隊
function fixSciCodesRun_() {
  var sh = tSh_();
  var all = tickets_();
  var fixed = 0, cleared = 0, lines = [];
  all.forEach(function (t) {
    if (!isSci_(t.code)) return;
    var got = fixSci_(t.code, t.rawText);
    var merged = {};
    T_HEAD.forEach(function (h) { merged[h] = t[h]; });
    if (isSci_(got)) {
      merged.code = '';
      merged.confidence = 'low';
      cleared++;
      lines.push('清空：' + ticketLabel_(t) + '（rawText 裡也沒有完整序號，請自己補）');
    } else {
      merged.code = got;
      fixed++;
      lines.push('修好：' + ticketLabel_(t) + ' → ' + got);
    }
    merged.fingerprint = fingerprint_(merged);
    sh.getRange(t._row, 1, 1, T_HEAD.length).setNumberFormat('@').setValues([rowFrom_(merged, T_HEAD)]);
  });
  Logger.log(lines.join('\n') || '沒有科學記號的序號，不用修。');
  Logger.log('修好 ' + fixed + ' 張，清空 ' + cleared + ' 張，總共 ' + all.length + ' 張。');
  return { fixed: fixed, cleared: cleared };
}

/** 在編輯器選這個函式按執行，每個有設定的管道都會收到一則測試訊息。 */
function testNotify() {
  var hasD = !!P_('DISCORD_WEBHOOK');
  var hasL = !!(P_('LINE_TOKEN') && P_('LINE_USER_ID'));
  if (!hasD && !hasL) {
    Logger.log('一個通知管道都還沒設定。到「專案設定 → 指令碼屬性」新增 DISCORD_WEBHOOK（或 LINE_TOKEN + LINE_USER_ID）。');
    return;
  }
  var msg = '票券 App 的通知設定成功。之後有人用掉你分享的券，就會在這裡通知你。';
  if (hasD) {
    var d = notifyDiscord_({ kind: 'used', who: '測試', ticket: { code: 'TEST-0001', expiry: today_() } }, msg, '');
    Logger.log('Discord：' + (d ? '已送出，去頻道看看。' : '沒送出去，看上面那行寫的原因。'));
  }
  if (hasL) {
    var l = notifyLine_(msg);
    Logger.log('LINE：' + (l ? '已送出，去 LINE 看看。' : '沒送出去，看上面那行寫的原因。'));
  }
}

/** 舊名字，留著相容。 */
function testLine() { testNotify(); }

/**
 * App 設定頁「傳一則測試通知」按鈕呼叫這裡。
 * 訊息裡寫著後端送出的時間；手機跳出來的時間比它晚多少，就是 Discord 推播＋手機那一段的延遲。
 */
function testNotifyNow_() {
  var hasD = !!P_('DISCORD_WEBHOOK');
  var hasL = !!(P_('LINE_TOKEN') && P_('LINE_USER_ID'));
  if (!hasD && !hasL) throw new Error('後端還沒設定通知管道（指令碼屬性 DISCORD_WEBHOOK）');
  var sent = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm:ss');
  var text = '🔔 測試通知｜後端 ' + sent + ' 送出。看一下手機是幾點跳出來的。';
  var ok = hasD ? postDiscord_({ username: '票券', content: text, allowed_mentions: { parse: [] } }, { noQueue: true }) : false;
  if (hasL) ok = notifyLine_(text) || ok;
  var d = LAST_DISCORD_ || {};
  var why = ok ? '' : ((LAST_DISCORD_ && LAST_DISCORD_.why) || (LAST_LINE_ && LAST_LINE_.why) || '');
  return { ok: ok, sentAt: sent, discordAt: d.at || '', ms: d.ms || 0, code: d.code || 0, why: why,
           retried: !!d.retried, via: d.via || '', flushed: FLUSHED_, waiting: readQueue_().length };
}

/**
 * 延遲診斷：在編輯器選這個函式按執行。
 * 分段量「標記已使用」在後端的每一步，並送一則寫著送出時間的測試訊息到 Discord。
 * 不會改到任何票券。
 */
function diagnoseLatency() {
  var tz = Session.getScriptTimeZone();
  var id = P_('SHEET_ID');
  var lines = ['', '───────── 延遲診斷 ─────────'];
  function sec(ms) { return (ms / 1000).toFixed(2) + ' 秒'; }

  // 先暖身一次，兩種做法都在同樣「已連上」的狀態下比
  SpreadsheetApp.openById(id).getSheetByName(T_SHEET).getLastRow();

  // 舊做法：開三次試算表、整張表讀兩次（這一版之前「標記已使用」就是這樣）
  var t = Date.now();
  var sh1 = SpreadsheetApp.openById(id).getSheetByName(T_SHEET);
  var all = readRows_(sh1, T_HEAD, 'id');
  var sample = all.length ? all[Math.floor(all.length / 2)].id : '';
  readRows_(SpreadsheetApp.openById(id).getSheetByName(T_SHEET), T_HEAD, 'id');
  SpreadsheetApp.openById(id).getSheetByName(T_SHEET);
  var oldMs = Date.now() - t;
  lines.push('找票（舊做法：開表 3 次、整張 ' + all.length + ' 張讀 2 次）：' + sec(oldMs));

  // 新做法：開一次、只讀 id 欄和那一列
  SS_ = null; TSH_ = null;
  t = Date.now();
  var found = findTicket_(sample);
  tSh_();
  var newMs = Date.now() - t;
  lines.push('找票（新做法：開表 1 次、只讀那一列）：' + sec(newMs) + (found ? '' : '（沒找到樣本票券）'));

  // 寫入：把標題列第一格原樣寫回去，不動任何票券
  t = Date.now();
  var c = tSh_().getRange(1, 1);
  c.setValue(c.getValue());
  SpreadsheetApp.flush();
  var writeMs = Date.now() - t;
  lines.push('寫入一格並存檔：' + sec(writeMs));

  // Discord：送一則帶時間的訊息，等 Discord 回「收到了」
  if (P_('DISCORD_WEBHOOK')) {
    var sent = new Date();
    var hhmmss = Utilities.formatDate(sent, tz, 'HH:mm:ss');
    postDiscord_({ username: '票券', allowed_mentions: { parse: [] },
      content: '⏱ 延遲測試｜後端 ' + hhmmss + ' 送出。看一下手機是幾點跳出來的。' }, { noQueue: true, noFlush: true });
    var d = LAST_DISCORD_ || {};
    lines.push('送到 Discord 並等它確認：' + sec(d.ms || 0) + '（HTTP ' + d.code + '）');
    if (d.atMs) {
      lines.push('Discord 伺服器記錄的收到時間：' + d.at + '（比送出晚 ' + sec(Math.max(0, d.atMs - sent.getTime())) + '）');
    }
    lines.push('');
    lines.push('→ 後端這邊：找票＋寫入＋送 Discord ≈ ' + sec(newMs + writeMs + (d.ms || 0)) +
               '（舊做法 ≈ ' + sec(oldMs + writeMs + (d.ms || 0)) + '）');
    lines.push('→ 測試訊息 ' + hhmmss + ' 已經到 Discord。手機如果晚很多才跳，延遲就在 Discord 推播或手機省電這一段。');
  } else {
    lines.push('沒設定 DISCORD_WEBHOOK，略過 Discord 測試。');
  }
  lines.push('────────────────────────────');
  Logger.log(lines.join('\n'));
}


/* ══════════════════ 分享頁 ══════════════════ */

/** 給瀏覽器直接開的版本（Apps Script 會把它包在自己的沙箱 iframe 裡） */
function sharePage_(token, me) {
  return htmlOut_(shareHtml_(token, me));
}

/** /s/<碼> 打開的頁面：16 碼的是親友連結（v75），其他是分享連結 */
function pageHtml_(token, me) {
  if (String(token || '').length === MEMBER_TOKEN_LEN) {
    var m = memberRow_(token);
    if (!m) return shell_('無法開啟', '<p class="msg">這個連結不存在，或已經換成新的連結了。請跟給你連結的人要一條新的。</p>');
    if (m.revoked === 'y') return shell_('無法開啟', '<p class="msg">這個連結已經停用了。</p>');
    return memberHtml_(m, me);
  }
  return shareHtml_(token, me);
}

/**
 * 分享頁的 HTML 字串。_worker.js 會拿這一份直接送給對方，不經過 Apps Script 的沙箱。
 * me = 你自己從 App 按「開啟」來看的，不記開啟次數（「對方還沒打開」才準）。
 */
function shareHtml_(token, me) {
  var s;
  try { s = openShare_(token); }
  catch (err) {
    if (err.notFound && shareWaitable_(token)) return shell_('連結準備中', shareWaitHtml_());   // v72：App 剛給出連結、後端還沒存好
    return shell_('無法開啟', '<p class="msg">' + esc_(err.message) + '</p>');
  }

  // 記一下開啟次數
  if (!me) {
    try {
      var sh = sSh_();
      sh.getRange(s._row, S_HEAD.indexOf('openCount') + 1).setValue(String((+s.openCount || 0) + 1));
      sh.getRange(s._row, S_HEAD.indexOf('lastOpenedAt') + 1).setValue(today_());
    } catch (e) {}
  }

  var ids = s.ticketIds.split(',');
  var all = tickets_();
  var picked = [];
  ids.forEach(function (id) {
    for (var i = 0; i < all.length; i++) if (all[i].id === id) picked.push(all[i]);
  });
  if (!picked.length) {
    return shell_('無法開啟', '<p class="msg">這個連結裡的票券都已經收回了。</p>');
  }

  var cards = picked.map(function (t) {
    var img = '';
    try {
      if (t.fileId) {
        var b = DriveApp.getFileById(t.fileId).getBlob();
        img = '<button type="button" class="shotbtn" aria-label="放大這張截圖">' +
              '<img class="shot" alt="票券截圖" src="data:' + b.getContentType() +
              ';base64,' + Utilities.base64Encode(b.getBytes()) + '"></button>';
      }
    } catch (e) {}

    var used = t.status === 'used';
    var until = undoUntil_(t, token);
    var st = stub_(t);
    // v69 餘額型：對方用完一次不代表用光，按了只是通知你去確認剩多少，不會進「已使用」
    var bal = t.kind === 'balance' && !used;
    var pend = bal && !!t.pendingAt;
    var left = bal && /^\d+$/.test(String(t.balance || ''))
      ? '剩 $' + String(t.balance).replace(/\B(?=(\d{3})+(?!\d))/g, ',') : '';    // 只給剩餘金額；查餘額的網址絕不放進分享頁
    return '' +
      '<article class="tk' + (used ? ' used' : '') + (pend ? ' pend' : '') + '" data-id="' + esc_(t.id) + '"' +
        (until ? ' data-undo-until="' + until + '"' : '') + '>' +
        '<div class="tw"><div class="tkt" data-tier="' + st.tier + '"' +
            (isHotel_(t) ? ' data-kind="hotel"' : '') +
            (brandColor_(t) ? ' style="--brand:' + brandColor_(t) + '"' : '') + '>' +
          '<span class="edge"></span>' +
          '<div class="bd">' +
            '<span class="mc">' + esc_(t.merchant || '票券') + '</span>' +
            '<h2>' + esc_(t.title || '（未命名）') + '</h2>' +
            (left || t.value ? '<span class="val">' + esc_(left || t.value) + '</span>' : '') +
            (t.code ? '<span class="code">' + esc_(t.code) + '</span>' : '') +
            '<span class="ink">已使用</span>' +
          '</div>' +
          '<div class="stub">' +
            (st.top ? '<span class="s-top">' + esc_(st.top) + '</span>' : '') +
            '<span class="s-num' + (st.cls ? ' ' + st.cls : '') + '">' + esc_(st.num) + '</span>' +
            (st.unit ? '<span class="s-unit">' + esc_(st.unit) + '</span>' : '') +
            (st.date ? '<span class="s-date">' + esc_(st.date) + '</span>' : '') +
          '</div>' +
        '</div></div>' +
        (img || '') +
        '<div class="act">' +
          '<span class="warn"></span>' +
          '<span class="stamp" ' + (used || pend ? '' : 'hidden') + '>' +
            (pend ? '已通知對方，他會確認剩餘金額' : '已使用 ' + esc_(t.usedAt || '')) + '</span>' +
          '<button class="undo" ' + (until ? '' : 'hidden') + '>按錯了，復原</button>' +
          '<button class="give" ' + (used || pend ? 'hidden' : '') + '>用不到，還給他</button>' +
          '<button class="use" ' + (used || pend ? 'hidden' : '') + '>' + (bal ? '我用了一些（通知對方確認餘額）' : '我用掉了') + '</button>' +
        '</div>' +
        (bal ? '<p class="balnote">餘額型：用完一次不代表用光，剩多少由對方確認。</p>' : '') +
        '<div class="who" hidden>' +
          '<label>是誰用的？<input class="whoname" maxlength="20" autocomplete="name" placeholder="填你的名字"></label>' +
          '<div class="whobtns">' +
            '<button class="cancel">取消</button>' +
            '<button class="confirm" data-label="' + (bal ? '確定，通知對方' : '確定用掉了') + '">' + (bal ? '確定，通知對方' : '確定用掉了') + '</button>' +
          '</div>' +
        '</div>' +
      '</article>';
  }).join('');
  var hasBal = picked.some(function (t) { return t.kind === 'balance' && t.status !== 'used'; });

  var many = picked.length === 1 ? '這張票券' : '這 ' + picked.length + ' 張票券';
  var heading = s.to
    ? s.to + '，' + many + '給你'
    : '有人分享了' + (picked.length === 1 ? '一張票券' : ' ' + picked.length + ' 張票券') + '給你';
  var body =
    '<p class="msg">兌換時把螢幕亮度調到最高，條碼比較好掃。點截圖可以放大，也能轉成橫的讓條碼更寬。用掉之後按一下「我用掉了」，對方就知道這張不能再用了；用不到的話按「用不到，還給他」，這張就回到他手上。' +
      (SHARE_UNDO_MINUTES ? '按錯的話 ' + SHARE_UNDO_MINUTES + ' 分鐘內可以自己復原。' : '按錯的話可以自己復原。') +
      (hasBal ? '標示「餘額型」的券，用完一次不代表用光：按「我用了一些」，對方會收到通知、再確認剩多少。' : '') +
      (s.expiresAt ? '這個連結 ' + esc_(s.expiresAt) + ' 之後會失效。' : '') + '</p>' +
    cards +
    '<div class="lb" id="lb" hidden>' +
      '<div class="lb-bar">' +
        '<button type="button" class="lb-btn" id="lbClose">關閉</button>' +
        '<button type="button" class="lb-btn" id="lbRot">轉橫向</button>' +
        '<button type="button" class="lb-btn" id="lbZoom">放大</button>' +
      '</div>' +
      '<div class="lb-stage" id="lbStage"><div class="lb-box" id="lbBox"><img id="lbImg" alt="票券截圖"></div></div>' +
    '</div>' +
    '<script>' +
      SHARE_JS.replace('__API__', JSON.stringify(shareApiBase_()))
              .replace('__TOK__', JSON.stringify(token))
              .replace('__TO__', JSON.stringify(s.to || '')) +
    '<\/script>';

  return shell_('分享的票券', body, { heading: heading, sub: s.note ? '「' + s.note + '」' : '' });
}

/** 票券左邊色帶的顏色，跟 App 的 BRAND_COLORS 是同一張表。 */
var BRAND_COLOR_RULES = [
  [/7[\s\-]?11|7[\s\-]?eleven|統一超商|小七|city\s*cafe/i, '#E4572E'],
  [/全家|family\s*mart|familymart/i, '#1B7DC4'],
  [/萊爾富|hi[\s\-]?life/i, '#D6242B'],
  [/ok\s*(超商|mart)/i, '#E4002B'],
  [/星巴克|starbucks/i, '#1E6B4A'],
  [/麥當勞|mcdonald|mccafe/i, '#D52B1E'],
  [/路易莎|louisa/i, '#3B2A20'],
  [/cama/i, '#6B4A2E'],
  [/85\s*度\s*c/i, '#9C2A32'],
  [/摩斯|mos\s*burger/i, '#C8102E'],
  [/肯德基|kfc/i, '#A3080C'],
  [/漢堡王|burger\s*king/i, '#D4691F'],
  [/必勝客|pizza\s*hut/i, '#C8102E'],
  [/達美樂|domino/i, '#0B6FB3'],
  [/家樂福|carrefour/i, '#1F4E9C'],
  [/全聯|pxmart/i, '#0070B8'],
  [/大潤發|rt[\s\-]?mart/i, '#D71920'],
  [/新光三越/, '#8B1E3F'],
  [/sogo|崇光/i, '#C8102E'],
  [/誠品|eslite/i, '#2F2F2F'],
  [/鼎泰豐/, '#7A1F1F'],
  [/王品|西堤|陶板屋|聚北海道|石二鍋|藝奇/, '#7A1F2B'],
  [/乾杯|kanpai/i, '#B8292F']
];
var HOTEL_RE = /飯店|酒店|hotel|晶華|君悅|喜來登|萬豪|文華|老爺|福華|國賓|寒舍|住宿|房券|一泊/i;
function isHotel_(t) { return HOTEL_RE.test((t.merchant || '') + ' ' + (t.title || '')); }
var USE_BRAND_COLORS = true;    // 改成 true，分享頁的色帶就會變成各家商店的顏色
function brandColor_(t) {
  if (!USE_BRAND_COLORS) return '';
  var hay = (t.merchant || '') + ' ' + (t.title || '');
  for (var i = 0; i < BRAND_COLOR_RULES.length; i++) {
    if (BRAND_COLOR_RULES[i][0].test(hay)) return BRAND_COLOR_RULES[i][1];
  }
  return '';
}

/** 票根上要寫什麼、用哪個顏色。跟 App 的 stubParts 是同一套規則。 */
function stub_(t) {
  if (!t.expiry) return { tier: 'none', top: '', num: '不限', cls: 'word', unit: '', date: '無到期日' };
  var today = today_();
  var d = Math.round((Date.parse(t.expiry + 'T00:00:00Z') - Date.parse(today + 'T00:00:00Z')) / 86400000);
  if (isNaN(d)) return { tier: 'none', top: '', num: '—', cls: 'word', unit: '', date: '' };
  var m = t.expiry.split('-');
  var md = (+m[1]) + '/' + (+m[2]);
  var due = (m[0] === today.slice(0, 4)) ? md + ' 到期' : m[0] + '/' + md;
  if (d < 0)  return { tier: 'gone', top: '過期', num: String(Math.min(-d, 999)), cls: '', unit: '天', date: due };
  if (d === 0) return { tier: 'urgent', top: '', num: '今天', cls: 'word', unit: '', date: '最後一天' };
  var tier = d <= 3 ? 'urgent' : d <= 7 ? 'soon' : d <= 30 ? 'ok' : 'far';
  return { tier: tier, top: '剩', num: d > 999 ? '999+' : String(d), cls: d > 99 ? 'small' : '', unit: '天', date: due };
}

/** 分享頁在對方手機上跑的程式。__API__ / __TOK__ 會在送出頁面前換掉。 */
var SHARE_JS = [
  '(function(){',
  '  var API = __API__, TOK = __TOK__, TO = __TO__;',
  '',
  '  // 記住對方填過的名字，下一張就不用再打；讀不到就用分享時填的「給誰」',
  '  function savedName(){ try{ return localStorage.getItem("tw.name") || TO; }catch(e){ return TO; } }',
  '  function saveName(n){ try{ localStorage.setItem("tw.name", n); }catch(e){} }',
  '',
  '  function post(body){',
  '    return fetch(API, {method:"POST", redirect:"follow", body: JSON.stringify(body)})',
  '      .then(function(r){ return r.json(); })',
  '      .then(function(j){ if(!j.ok) throw new Error(j.error || "失敗"); return j.data; });',
  '  }',
  '  function q(card, sel){ return card.querySelector(sel); }',
  '  function warn(card, msg){ q(card, ".warn").textContent = msg || ""; }',
  '',
  '  function openWho(card){',
  '    q(card, ".use").hidden = true;',
  '    q(card, ".give").hidden = true;',
  '    q(card, ".who").hidden = false;',
  '    var inp = q(card, ".whoname");',
  '    if(!inp.value) inp.value = savedName();',
  '    inp.focus(); inp.select();',
  '    warn(card, "");',
  '  }',
  '  function closeWho(card){',
  '    q(card, ".who").hidden = true;',
  '    if(!card.classList.contains("used") && !card.classList.contains("pend")){ q(card, ".use").hidden = false; q(card, ".give").hidden = false; }',
  '  }',
  '  // v69 餘額型：按了之後不是「已使用」，只是已經通知對方去確認剩多少',
  '  function showPending(card){',
  '    card.classList.add("pend");',
  '    q(card, ".use").hidden = true; q(card, ".give").hidden = true;',
  '    q(card, ".who").hidden = true; q(card, ".undo").hidden = true;',
  '    var st = q(card, ".stamp"); st.hidden = false; st.textContent = "已通知對方，他會確認剩餘金額";',
  '  }',
  '  function showUsed(card, usedAt, until){',
  '    card.classList.add("used");',
  '    q(card, ".use").hidden = true;',
  '    q(card, ".give").hidden = true;',
  '    q(card, ".who").hidden = true;',
  '    var st = q(card, ".stamp"); st.hidden = false; st.textContent = "已使用 " + (usedAt || "");',
  '    var u = q(card, ".undo");',
  '    if(until && until > Date.now()){ u.hidden = false; card.setAttribute("data-undo-until", until); }',
  '    else { u.hidden = true; card.removeAttribute("data-undo-until"); }',
  '  }',
  '  function showUnused(card){',
  '    card.classList.remove("used");',
  '    q(card, ".use").hidden = false;',
  '    q(card, ".give").hidden = false;',
  '    q(card, ".who").hidden = true;',
  '    q(card, ".stamp").hidden = true;',
  '    q(card, ".undo").hidden = true;',
  '    card.removeAttribute("data-undo-until");',
  '  }',
  '',
  '  function confirmUse(card){',
  '    var inp = q(card, ".whoname"), btn = q(card, ".confirm");',
  '    var name = inp.value.trim();',
  '    if(!name){ warn(card, "請填你的名字，對方才知道是誰用的"); inp.focus(); return; }',
  '    saveName(name);',
  '    btn.disabled = true; btn.textContent = "回報中…"; warn(card, "");',
  '    post({action:"markUsed", token:TOK, id:card.dataset.id, by:name})',
  '      .then(function(d){ if(d.pending) showPending(card); else showUsed(card, d.usedAt, d.undoUntil); })',
  '      .catch(function(err){ warn(card, "回報失敗：" + err.message); })',
  '      .then(function(){ btn.disabled = false; btn.textContent = btn.getAttribute("data-label") || "確定用掉了"; });',
  '  }',
  '',
  '  // 用不到，還給他：按第一下只是確認，四秒內再按一次才送出',
  '  function giveBack(card, el){',
  '    if(el.getAttribute("data-armed") !== "1"){',
  '      el.setAttribute("data-armed", "1"); el.textContent = "確定還給他？";',
  '      setTimeout(function(){',
  '        if(el.getAttribute("data-armed") === "1"){ el.setAttribute("data-armed", "0"); el.textContent = "用不到，還給他"; }',
  '      }, 4000);',
  '      return;',
  '    }',
  '    el.setAttribute("data-armed", "0");',
  '    el.disabled = true; el.textContent = "處理中…"; warn(card, "");',
  '    post({action:"returnShare", token:TOK, id:card.dataset.id, by:savedName()})',
  '      .then(function(d){',
  '        if(d.used){ showUsed(card, d.usedAt, 0); warn(card, "這張已經被標記用掉了，沒辦法還"); return; }',
  '        showBack(card);',
  '      })',
  '      .catch(function(err){ warn(card, err.message); })',
  '      .then(function(){ el.disabled = false; if(!card.classList.contains("back")) el.textContent = "用不到，還給他"; });',
  '  }',
  '  function showBack(card){',
  '    card.classList.add("back");',
  '    q(card, ".use").hidden = true; q(card, ".give").hidden = true;',
  '    q(card, ".who").hidden = true; q(card, ".undo").hidden = true;',
  '    var st = q(card, ".stamp"); st.hidden = false; st.textContent = "已經還給他了，謝謝！";',
  '  }',
  '',
  '  // 復原時間到了，按鈕自己收起來',
  '  setInterval(function(){',
  '    var cs = document.querySelectorAll(".tk[data-undo-until]");',
  '    for(var i = 0; i < cs.length; i++){',
  '      if(+cs[i].getAttribute("data-undo-until") < Date.now()){',
  '        q(cs[i], ".undo").hidden = true; cs[i].removeAttribute("data-undo-until");',
  '      }',
  '    }',
  '  }, 15000);',
  '',
  '  document.addEventListener("click", function(e){',
  '    var el = e.target.closest("button"); if(!el) return;',
  '    var card = el.closest(".tk"); if(!card) return;',
  '',
  '    if(el.classList.contains("use"))     { openWho(card); return; }',
  '    if(el.classList.contains("cancel"))  { closeWho(card); return; }',
  '    if(el.classList.contains("confirm")) { confirmUse(card); return; }',
  '    if(el.classList.contains("give"))    { giveBack(card, el); return; }',
  '',
  '    if(el.classList.contains("undo")){',
  '      el.disabled = true; el.textContent = "復原中…"; warn(card, "");',
  '      post({action:"unmarkUsed", token:TOK, id:card.dataset.id})',
  '        .then(function(){ showUnused(card); })',
  '        .catch(function(err){ warn(card, err.message); })',
  '        .then(function(){ el.disabled = false; el.textContent = "按錯了，復原"; });',
  '    }',
  '  });',
  ''
].concat(LIGHTBOX_JS_LINES, [
  '  // 名字框按 Enter 就送出',
  '  document.addEventListener("keydown", function(e){',
  '    if(e.key !== "Enter" || !e.target.classList.contains("whoname")) return;',
  '    e.preventDefault();',
  '    confirmUse(e.target.closest(".tk"));',
  '  });',
  '})();'
]).join('\n');

/**
 * 分享頁、錯誤頁共用的外框：紫色票帶 + 鋸齒邊，跟 App 同一套長相。
 * opts.heading 是票帶上的大標，opts.sub 是大標下面那行小字。
 */
function shell_(title, inner, opts) {
  opts = opts || {};
  var heading = opts.heading || title;
  return '' +
'<!doctype html><html lang="zh-Hant"><head><meta charset="utf-8">' +
'<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">' +
'<meta name="theme-color" content="#F2F2F5">' +
'<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@600&family=Noto+Sans+TC:wght@400;500;700;900&display=swap">' +
'<title>' + esc_(title) + '</title><style>' + SHARE_CSS + '</style></head><body>' +
'<header class="head"><div class="in">' +
  '<span class="logo">票券</span>' +
  '<h1>' + esc_(heading) + '</h1>' +
  (opts.sub ? '<p class="sub">' + esc_(opts.sub) + '</p>' : '') +
'</div></header>' +
'<main>' + inner + '</main></body></html>';
}

var SHARE_CSS = [
  ':root{--paper:#F2F2F5;--card:#fff;--card2:#F7F7FA;--ink:#111114;--ink2:#4A4A55;--ink3:#8A8A96;',
  '--line:#E4E4EA;--line2:#D0D0D8;--accent:#6A35A8;--accentS:#EFE7FB;--calmS:#EDEDF2;',
  '--danger:#C4372B;--dangerS:#FCE6E3;--warn:#9A6A10;--warnS:#F8EDD8;--hotel:#F7F1E6;--hotelStub:#EFE5D2;--edge:#2A2A30;',
  '--drop:drop-shadow(0 1px 1px rgba(17,17,20,.07)) drop-shadow(0 6px 14px rgba(17,17,20,.08));',
  '--sans:"Noto Sans TC","PingFang TC","Microsoft JhengHei",system-ui,sans-serif;',
  '--num:"Barlow Condensed","Arial Narrow",sans-serif;',
  '--mono:ui-monospace,SFMono-Regular,Menlo,monospace;color-scheme:light dark}',
  '@media(prefers-color-scheme:dark){:root{--paper:#0F0F12;--card:#1C1C21;--card2:#232329;--ink:#F2F2F5;',
  '--ink2:#B9B9C4;--ink3:#86868F;--line:#2C2C33;--line2:#3B3B44;--accent:#B190EC;--accentS:#2A2140;--calmS:#26262D;',
  '--danger:#F08A7E;--dangerS:#3A201D;--warn:#DDAE5C;--warnS:#342A19;--hotel:#27231C;--hotelStub:#2F2A21;--edge:#4A4A54;',
  '--drop:drop-shadow(0 1px 1px rgba(0,0,0,.5)) drop-shadow(0 6px 14px rgba(0,0,0,.45))}}',
  '*{box-sizing:border-box}[hidden]{display:none!important}',
  'body{margin:0;background:var(--paper);color:var(--ink);font:15px/1.6 var(--sans);',
  '-webkit-font-smoothing:antialiased;padding-bottom:calc(40px + env(safe-area-inset-bottom,0px))}',

  /* 頁首 */
  '.head{padding:calc(26px + env(safe-area-inset-top,0px)) 16px 6px}',
  '.head .in,main{max-width:520px;margin:0 auto}',
  '.logo{display:inline-block;font-size:13px;font-weight:900;letter-spacing:.12em;color:var(--accent)}',
  '.head h1{font-size:28px;font-weight:900;line-height:1.3;margin:4px 0 0;text-wrap:balance}',
  '.sub{margin:8px 0 0;font-size:16px;font-weight:500;color:var(--ink2)}',
  'main{padding:16px 16px 0}',
  '.msg{margin:0 0 24px;font-size:13.5px;color:var(--ink3);line-height:1.7}',

  /* 一張 = 票面（有齒孔、缺口）+ 截圖 + 按鈕 */
  '.tk{margin:0 0 30px}',
  '.tw{filter:var(--drop)}',
  '.tkt{--brand:var(--accent);--stubw:86px;--stub:var(--card2);--stubInk:var(--ink);',
  'position:relative;display:flex;min-height:124px;background:var(--card);border-radius:10px;',
  '-webkit-mask:radial-gradient(circle 3.5px at 0 50%,#0000 95%,#000) 0 0/100% 11px repeat-y,',
  'radial-gradient(circle 3.5px at 100% 50%,#0000 95%,#000) 0 0/100% 11px repeat-y,',
  'radial-gradient(circle 12px at 0 50%,#0000 97%,#000),radial-gradient(circle 12px at 100% 50%,#0000 97%,#000),',
  'radial-gradient(circle 8px at calc(100% - var(--stubw)) 0,#0000 96%,#000),',
  'radial-gradient(circle 8px at calc(100% - var(--stubw)) 100%,#0000 96%,#000);',
  '-webkit-mask-composite:source-in;',
  'mask:radial-gradient(circle 3.5px at 0 50%,#0000 95%,#000) 0 0/100% 11px repeat-y,',
  'radial-gradient(circle 3.5px at 100% 50%,#0000 95%,#000) 0 0/100% 11px repeat-y,',
  'radial-gradient(circle 12px at 0 50%,#0000 97%,#000),radial-gradient(circle 12px at 100% 50%,#0000 97%,#000),',
  'radial-gradient(circle 8px at calc(100% - var(--stubw)) 0,#0000 96%,#000),',
  'radial-gradient(circle 8px at calc(100% - var(--stubw)) 100%,#0000 96%,#000);',
  'mask-composite:intersect}',
  '.tkt[data-kind="hotel"]{background:var(--hotel);--stub:var(--hotelStub)}',
  '.tkt[data-tier="urgent"],.tkt[data-tier="soon"]{--stubInk:var(--danger)}',
  '.tkt[data-tier="none"],.tkt[data-tier="gone"]{--stubInk:var(--ink3)}',
  '.edge{position:absolute;left:0;top:0;bottom:0;width:13px;background:var(--brand)}',
  '.tkt[data-kind="hotel"] .edge{background:linear-gradient(180deg,#DCC285,#B08A45)}',
  '.bd{position:relative;flex:1;min-width:0;padding:16px 14px 15px 30px;display:flex;flex-direction:column;gap:3px}',
  '.mc{font-size:13px;font-weight:500;color:var(--ink2)}',
  '.bd h2{font-size:21px;font-weight:900;margin:0;line-height:1.3}',
  '.val{align-self:flex-start;margin-top:8px;font-size:12.5px;font-weight:700;background:var(--calmS);',
  'border-radius:999px;padding:3px 10px;color:var(--ink2)}',
  '.code{margin-top:auto;padding-top:10px;font-family:var(--mono);font-size:12.5px;color:var(--ink3);overflow-wrap:anywhere}',
  '.stub{flex:none;width:var(--stubw);display:flex;flex-direction:column;align-items:center;',
  'justify-content:center;gap:1px;padding:12px 8px;text-align:center;background:var(--stub);color:var(--stubInk);',
  'border-left:2px dashed var(--line2)}',
  '.s-top,.s-unit{font-size:12.5px;font-weight:700;line-height:1;color:var(--ink3)}',
  '.s-num{font-family:var(--num);font-weight:600;font-size:40px;line-height:.95;font-variant-numeric:tabular-nums}',
  '.s-num.word{font-family:var(--sans);font-size:20px;font-weight:900;line-height:1.2}',
  '.s-num.small{font-size:27px}',
  '.s-date{margin-top:7px;font-size:11px;font-weight:500;color:var(--ink3);white-space:nowrap}',

  /* 用掉之後：票面淡掉、蓋章 */
  '.ink{display:none;position:absolute;right:12px;top:50%;transform:translateY(-50%) rotate(-11deg);',
  'padding:5px 12px 4px;border:3px double currentColor;border-radius:8px;font-size:20px;font-weight:900;',
  'letter-spacing:.18em;line-height:1.2;white-space:nowrap;color:var(--accent);opacity:.85;pointer-events:none;',
  'animation:thunk .42s cubic-bezier(.2,1.6,.4,1) both}',
  '@keyframes thunk{0%{opacity:0;transform:translateY(-50%) rotate(-4deg) scale(1.9)}',
  '70%{opacity:.9;transform:translateY(-50%) rotate(-12deg) scale(.96)}',
  '100%{opacity:.85;transform:translateY(-50%) rotate(-11deg) scale(1)}}',
  '@media (prefers-reduced-motion:reduce){.ink{animation:none}}',
  '.tk.used .ink{display:block}',
  '.tk.used .bd > :not(.ink),.tk.used .edge{opacity:.4}.tk.used .stub,.tk.used .shot{opacity:.55}',

  '.shot{display:block;width:100%;border-radius:14px;border:1px solid var(--line);background:var(--card)}',
  '.shotbtn{position:relative;display:block;width:100%;margin-top:12px;padding:0;border:0;background:none;cursor:zoom-in}',

  /* 放大檢視：整片白底，條碼最好掃 */
  '.lb{position:fixed;inset:0;z-index:90;background:#fff;display:flex;flex-direction:column}',
  '.lb-bar{display:flex;gap:8px;padding:calc(8px + env(safe-area-inset-top,0px)) 12px 8px;border-bottom:1px solid #E4E4EA}',
  '.lb-btn{flex:1;padding:11px 8px;border:1px solid #D0D0D8;border-radius:999px;background:#fff;color:#111114;',
  'font:700 14px/1 var(--sans)}',
  '.lb-btn#lbClose{flex:0 0 auto;padding:11px 18px;background:#2B1B40;border-color:#2B1B40;color:#fff}',
  '.lb-stage{flex:1;overflow:auto;-webkit-overflow-scrolling:touch;display:flex}',
  '.lb-box{position:relative;margin:auto}',
  '.lb-box img{position:absolute;left:50%;top:50%;display:block;background:#fff}',
  '.act{display:flex;flex-wrap:wrap;align-items:center;justify-content:flex-end;gap:10px;padding:14px 2px 0}',
  '.stamp{font-size:13px;font-weight:500;color:var(--ink3);font-variant-numeric:tabular-nums}',
  '.warn{flex:1 1 100%;font-size:12.5px;color:var(--danger);text-align:left}.warn:empty{display:none}',
  /* v69 餘額型：一行小字提醒；對方按了之後票面照常（不蓋「已使用」），只在下面寫「已通知對方」 */
  '.balnote{margin:8px 2px 0;font-size:12.5px;color:var(--warn);line-height:1.5}',
  '.tk.pend .stamp{color:var(--warn);font-weight:700}',
  /* 還給他之後：條碼和截圖收起來，免得不小心又拿去用 */
  '.tk.back .tw{opacity:.45}.tk.back .shotbtn,.tk.back .code{display:none}',
  '.who{padding:14px 2px 0;display:grid;gap:10px}',
  '.who label{display:grid;gap:6px;font-size:13px;color:var(--ink2);font-weight:700}',
  '.who input{font:inherit;width:100%;padding:11px 14px;border:1px solid var(--line2);',
  'border-radius:12px;background:var(--card);color:var(--ink)}',
  '.who input:focus{outline:none;border-color:var(--accent);box-shadow:0 0 0 3px var(--accentS)}',
  '.whobtns{display:flex;gap:8px;justify-content:flex-end}',

  /* 按鈕：跟 App 一樣的膠囊，主要動作是黑的 */
  'button{font:inherit;font-weight:700;border:1px solid var(--ink);background:var(--ink);color:var(--paper);',
  'border-radius:999px;padding:11px 22px;cursor:pointer}',
  'button:disabled{opacity:.5}',
  'button.undo,button.cancel{background:var(--card);color:var(--ink2);border-color:var(--line2)}',
  'button.undo{color:var(--accent);border-color:var(--accent)}',
  'button.give{background:var(--card);color:var(--ink2);border-color:var(--line2);padding:11px 16px}',

  /* v75 親友頁：商家篩選、按「出示」才出現的截圖、收起 */
  '.chips{display:flex;gap:8px;overflow-x:auto;margin:-8px 0 22px;padding:2px 0 6px;scrollbar-width:none}',
  '.chips::-webkit-scrollbar{display:none}',
  'button.chip{flex:none;padding:7px 14px;font-size:13.5px;white-space:nowrap;background:var(--card);color:var(--ink2);border-color:var(--line2)}',
  'button.chip.on{background:var(--ink);color:var(--paper);border-color:var(--ink)}',
  '.noshot{margin:12px 2px 0;padding:18px;border-radius:14px;background:var(--card2);font-size:13.5px;color:var(--ink2);text-align:center}',
  'button.hide{background:var(--card);color:var(--ink2);border-color:var(--line2)}',
  '.tk.used .shotbox{display:none}'
].join('');

function htmlOut_(html) {
  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function esc_(s) {
  return String(s === null || s === undefined ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
