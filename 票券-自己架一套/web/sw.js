/* 票券 — 離線快取
   換版時把 VERSION 加一，舊快取就會被清掉。 */
const VERSION = 'ticket-wallet-v98';
const SHARE_CACHE = 'tw-shared';       // 從相簿分享進來的圖，先寄放在這裡
const SHELL = [
  './',
  './index.html',
  './manifest.webmanifest',
  './icon-192.png',
  './icon-512.png'
];
// v81 店家小圖示：跟 index.html 的 BRAND_MARK 對得上。一張一張放，缺哪張不影響其他（不放進 SHELL，免得一張 404 整批都失敗）
const BRAND_FILES = [
  './brands/mcdonalds.svg',
  './brands/seven-eleven.svg',
  './brands/familymart.svg',
  './brands/kfc.svg',
  './brands/mos.svg',
  './brands/starbucks.png',
  './brands/pxmart.png',
  './brands/carrefour.png',
  './brands/sogo.png',
  './brands/chingshin.png',
  './brands/hilife.png',
  './brands/misterdonut.png',
  './brands/mrwish.png',
  './brands/burgerking.png'
];

self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(VERSION)
      .then((c) => c.addAll(SHELL).then(() => Promise.all(BRAND_FILES.map((f) => c.add(f).catch(() => {})))))
      .then(() => self.skipWaiting())
      .catch(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      // 只清舊版的 App 外殼；存在手機上的截圖（tw-img）和分享暫存要留著
      .then((keys) => Promise.all(keys.filter((k) => k !== VERSION && k.indexOf('ticket-wallet-') === 0).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

/* ───────── 分享進來的處理 ─────────
 * 先把原始內容整包讀下來記錄（型別頭、長度），再用瀏覽器內建的 formData() 解析；
 * 萬一內建的解析不出東西，但內容其實不是空的，就自己切 multipart 再撈一次。
 */
async function handleShare_(req) {
  let n = 0, why = '';
  const log = [];
  const ct = req.headers.get('content-type') || '';
  log.push('型別頭=' + (ct ? ct.split(';')[0] : '沒有'));

  let raw = null;
  try {
    raw = new Uint8Array(await req.clone().arrayBuffer());
    log.push('內容=' + raw.byteLength + 'B');
  } catch (err) {
    log.push('讀不到內容(' + (err && err.message) + ')');
  }

  const files = [];
  let text = '';

  // 1) 瀏覽器內建的解析
  try {
    const fd = await req.formData();
    for (const [k, v] of fd.entries()) {
      if (v && typeof v === 'object' && typeof v.size === 'number') {
        log.push(k + '=' + (v.type || '沒型別') + '/' + v.size + 'B');
        if (v.size > 0) files.push(v);
      } else if (typeof v === 'string' && v.trim()) {
        log.push(k + '=文字');
        text += (text ? '\n' : '') + v.trim();
      }
    }
  } catch (err) {
    log.push('內建解析失敗(' + (err && err.message) + ')');
  }

  // 2) 內建的沒撈到，但內容不是空的 → 自己切
  if (!files.length && raw && raw.byteLength > 0) {
    const parts = parseMultipart_(raw, ct);
    log.push('自己切出 ' + parts.length + ' 段');
    parts.forEach((p) => {
      if (p.filename || (p.type && p.type.indexOf('text/') !== 0 && p.body.byteLength > 0)) {
        files.push(new File([p.body], p.filename || 'shared.jpg', { type: p.type || 'image/jpeg' }));
      } else if (p.body.byteLength) {
        const t = new TextDecoder().decode(p.body).trim();
        if (t) text += (text ? '\n' : '') + t;
      }
    });
  }

  try {
    const cache = await caches.open(SHARE_CACHE);
    const old = await cache.keys();
    await Promise.all(old.map((k) => cache.delete(k)));

    await Promise.all(files.map((f, i) => cache.put(
      new Request(new URL('shared/' + i, self.registration.scope)),
      new Response(f, { headers: {
        'content-type': (f.type && f.type.indexOf('image/') === 0) ? f.type : 'image/jpeg',
        'x-filename': encodeURIComponent(f.name || ('分享的截圖-' + (i + 1) + '.jpg'))
      } })
    )));
    n = files.length;

    if (!n && text) {
      await cache.put(new Request(new URL('shared/text', self.registration.scope)),
        new Response(text, { headers: { 'content-type': 'text/plain; charset=utf-8' } }));
      why = 'text';
    } else if (!n) {
      why = 'nofile';
    }

    await cache.put(new Request(new URL('shared/debug', self.registration.scope)),
      new Response(VERSION.replace('ticket-wallet-', '') + '｜' + log.join('｜'),
        { headers: { 'content-type': 'text/plain; charset=utf-8' } }));
  } catch (err) {
    why = 'error';
  }

  const back = new URL('index.html?shared=' + n + (why ? '&why=' + why : ''), self.registration.scope);
  return Response.redirect(back.href, 303);
}

/** 自己切 multipart/form-data。只在瀏覽器內建的解析失靈時才用。 */
function parseMultipart_(bytes, contentType) {
  const m = /boundary=(?:"([^"]+)"|([^;]+))/i.exec(contentType || '');
  let boundary = m ? (m[1] || m[2]).trim() : '';
  if (!boundary) {
    // 型別頭沒寫 boundary：從內容第一行猜
    const head = new TextDecoder().decode(bytes.slice(0, 200));
    const g = /^--([^\r\n]+)/.exec(head);
    if (!g) return [];
    boundary = g[1];
  }
  const enc = new TextEncoder();
  const delim = enc.encode('--' + boundary);
  const out = [];

  const find = (from) => {
    outer: for (let i = from; i <= bytes.length - delim.length; i++) {
      for (let j = 0; j < delim.length; j++) if (bytes[i + j] !== delim[j]) continue outer;
      return i;
    }
    return -1;
  };

  let pos = find(0);
  while (pos >= 0) {
    let start = pos + delim.length;
    if (bytes[start] === 45 && bytes[start + 1] === 45) break;        // "--" 結尾
    if (bytes[start] === 13) start += 2;                               // \r\n
    const next = find(start);
    if (next < 0) break;
    let end = next;
    if (bytes[end - 2] === 13 && bytes[end - 1] === 10) end -= 2;      // 去掉段尾 \r\n

    // 標頭跟內容之間是 \r\n\r\n
    let split = -1;
    for (let i = start; i < end - 3; i++) {
      if (bytes[i] === 13 && bytes[i + 1] === 10 && bytes[i + 2] === 13 && bytes[i + 3] === 10) { split = i; break; }
    }
    if (split > 0) {
      const headers = new TextDecoder().decode(bytes.slice(start, split));
      const body = bytes.slice(split + 4, end);
      const name = (/name="([^"]*)"/i.exec(headers) || [])[1] || '';
      const filename = (/filename\*?="?([^";\r\n]*)"?/i.exec(headers) || [])[1] || '';
      const type = ((/content-type:\s*([^\r\n;]+)/i.exec(headers) || [])[1] || '').trim();
      out.push({ name, filename, type, body });
    }
    pos = next;
  }
  return out;
}

// App 問「你是哪一版」時回答，設定頁會顯示出來
self.addEventListener('message', (e) => {
  if (e.data === 'version' && e.source) e.source.postMessage({ version: VERSION });
});

self.addEventListener('fetch', (e) => {
  const req = e.request;

  // Android 的「分享到票券」：系統會 POST 過來，先把圖收進快取，再把畫面帶回 App
  if (req.method === 'POST' && /\/share-target\/?$/.test(new URL(req.url).pathname)) {
    e.respondWith(handleShare_(req));
    return;
  }

  // 只管自己網域的 GET；打後端 API 的請求一律走網路，不要快取票券資料
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return;

  // HTML 走「網路優先」，這樣改版後一開就是新的；斷線才退回快取
  if (req.mode === 'navigate' || (req.headers.get('accept') || '').includes('text/html')) {
    e.respondWith(
      fetch(req)
        .then((res) => {
          const copy = res.clone();
          caches.open(VERSION).then((c) => c.put(req, copy)).catch(() => {});
          return res;
        })
        .catch(() => caches.match(req).then((r) => r || caches.match('./index.html')))
    );
    return;
  }

  // 其他靜態檔走「快取優先」
  e.respondWith(
    caches.match(req).then((hit) => hit || fetch(req).then((res) => {
      const copy = res.clone();
      caches.open(VERSION).then((c) => c.put(req, copy)).catch(() => {});
      return res;
    }))
  );
});
