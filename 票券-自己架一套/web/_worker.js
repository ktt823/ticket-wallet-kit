/* 票券 — Cloudflare Pages 的入口（進階模式）
 *
 *  /s/<token> 由這裡「代為抓取」Apps Script 的分享頁，再原樣回給對方。
 *  重點是：對方的瀏覽器完全不會連到 script.google.com，
 *  所以不管他有沒有登入 Google、登入幾個帳號，都不會再看到
 *  「Sorry, unable to open the file at this time」那張 Drive 錯誤頁。
 *
 *  後端網址（/exec）從哪裡來：
 *    1. Cloudflare Pages 專案的「Settings → Variables and Secrets」裡的 EXEC_URL（建議，更新網站不用改檔）
 *    2. 沒設的話，用下面這一行
 *  換部署（重新部署成新版本不用改；重建部署才要）就改其中一個。
 *
 *  /relay/discord（v57）：Google 的伺服器被 Discord 擋掉時，後端改從這裡轉送通知。
 *  每一則都會先回頭問後端「這一則真的是你剛剛要送的嗎？」（relayCheck），對得上才送，
 *  所以別人沒辦法拿它來亂發。
 */
// 建議用 Cloudflare 的 EXEC_URL 設定（見安裝說明），這一行就不用改；真的要改，把引號裡的字換成你的 /exec 網址
const EXEC = '在這裡貼上你的 /exec 網址';

const TOKEN_RE = /^\/s\/([A-Za-z0-9_-]{3,64})\/?$/;
const EXEC_RE = /^https:\/\/script\.google\.com\/.+\/exec$/;

// 還沒設定後端網址時，直接講清楚，不要丟出看不懂的錯誤
const NOT_SET = '分享連結還沒設定好：請到 Cloudflare Pages 專案的「Settings → Variables and Secrets」新增 EXEC_URL，' +
  '值是你的 Apps Script /exec 網址，存檔後再重新上傳一次網站。';

// 只轉送到 Discord 的 webhook
const HOOK_RE = /^https:\/\/(canary\.|ptb\.)?(discord|discordapp)\.com\/api\/webhooks\/\d+\/[A-Za-z0-9_-]+$/;

// 更新包下載時的檔名（網址上用英文，存到電腦上是中文）
const KIT_NAME = '票券-自己架一套.zip';

function execOf(env) {
  const v = String((env && env.EXEC_URL) || EXEC || '').trim();
  return EXEC_RE.test(v) ? v : '';
}

/** 原樣回傳，只多加幾個標頭 */
function withHeaders(res, headers) {
  const out = new Response(res.body, res);
  Object.keys(headers).forEach(k => out.headers.set(k, headers[k]));
  return out;
}

/** 分享頁按「我用掉了」時打的 API，也從這裡代轉，對方一樣碰不到 Google。 */
async function proxyApi(request, exec) {
  const body = await request.text();
  const res = await fetch(exec, {
    method: 'POST',
    headers: { 'content-type': 'text/plain;charset=UTF-8' },   // 跟 App 一樣，避開 CORS 預檢
    body: body,
    redirect: 'follow'
  });
  return new Response(await res.text(), {
    status: res.status,
    headers: { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store' }
  });
}

function relayError(status, kind) {
  return new Response(JSON.stringify({ relay: kind }),
    { status: status, headers: { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store' } });
}

/** 幫後端把通知送到 Discord（Google 的伺服器被擋掉的時候用） */
async function relayDiscord(request, exec) {
  const text = await request.text();
  if (text.length > 64000) return relayError(413, 'too-big');
  let req = null;
  try { req = JSON.parse(text); } catch (e) {}
  if (!req || typeof req.nonce !== 'string' || !HOOK_RE.test(String(req.url || '')) ||
      !req.payload || typeof req.payload !== 'object') return relayError(400, 'bad');

  // 回頭問後端：內容的雜湊對得上、號碼沒用過，才是它剛剛要送的
  const hash = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text));
  const digest = [...new Uint8Array(hash)].map(b => b.toString(16).padStart(2, '0')).join('');
  let check = null;
  try {
    const res = await fetch(exec, {
      method: 'POST', redirect: 'follow',
      headers: { 'content-type': 'text/plain;charset=UTF-8' },
      body: JSON.stringify({ action: 'relayCheck', nonce: req.nonce, digest: digest })
    });
    check = JSON.parse(await res.text());
  } catch (e) {
    return relayError(502, 'exec');
  }
  if (!check || !check.ok || !check.data || check.data.ok !== true) return relayError(403, 'denied');

  let res;
  try {
    res = await fetch(req.url + (req.url.indexOf('?') < 0 ? '?wait=true' : '&wait=true'), {
      method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify(req.payload)
    });
  } catch (e) {
    return relayError(502, 'discord');
  }
  // Discord 回什麼就原樣回給後端（狀態碼、內容），後端照直送的方式判斷
  return new Response(await res.text(), {
    status: res.status,
    headers: { 'content-type': res.headers.get('content-type') || 'application/json', 'cache-control': 'no-store', 'x-relay': '1' }
  });
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const exec = execOf(env);

    // 分享進來時，App 的背景程式還沒接手（剛裝好還沒開過），請求會掉到這裡。
    // 不要回錯誤頁，帶回 App 並說明原因。
    if (request.method === 'POST' && /^\/share-target\/?$/.test(url.pathname)) {
      let size = 0;
      try { size = (await request.arrayBuffer()).byteLength; } catch (e) {}
      return Response.redirect(new URL('/index.html?shared=0&why=nosw&size=' + size, url).href, 303);
    }

    if (url.pathname === '/relay/discord' && request.method === 'POST') {
      if (!exec) return relayError(503, 'not-set');
      try { return await relayDiscord(request, exec); }
      catch (err) { return relayError(500, 'error'); }
    }

    if (url.pathname === '/api' && request.method === 'POST') {
      if (!exec) {
        return new Response(JSON.stringify({ ok: false, error: NOT_SET }),
          { status: 200, headers: { 'content-type': 'application/json; charset=utf-8' } });
      }
      try { return await proxyApi(request, exec); }
      catch (err) {
        return new Response(JSON.stringify({ ok: false, error: '連不到後端，請稍後再試' }),
          { status: 200, headers: { 'content-type': 'application/json; charset=utf-8' } });
      }
    }

    // 版本說明：朋友自己架的 App 會從他們的網站來讀，所以要允許跨網站讀取，也不要被快取住
    if (url.pathname === '/release.json') {
      return withHeaders(await env.ASSETS.fetch(request), {
        'access-control-allow-origin': '*', 'cache-control': 'no-cache', 'content-type': 'application/json; charset=utf-8'
      });
    }

    // 更新包：存到電腦上用中文檔名，每次都抓最新的
    if (/^\/download\/[^/]+\.zip$/.test(url.pathname)) {
      return withHeaders(await env.ASSETS.fetch(request), {
        'content-disposition': "attachment; filename=\"tickets-kit.zip\"; filename*=UTF-8''" + encodeURIComponent(KIT_NAME),
        'cache-control': 'no-cache'
      });
    }

    const m = TOKEN_RE.exec(url.pathname);

    // 不是分享連結 → 照常給靜態檔案（index.html、圖示、404.html…）
    if (!m) return env.ASSETS.fetch(request);

    if (!exec) {
      return new Response(NOT_SET, { status: 500, headers: { 'content-type': 'text/plain; charset=utf-8' } });
    }

    const token = m[1];
    const plain = exec + '?s=' + encodeURIComponent(token);

    try {
      // raw=1 會回 JSON 包著的純 HTML，才不會拿到 Apps Script 自己的沙箱外框（那會是一片空白）
      // me=1 = 你自己從 App 按「開啟」來看的，後端不會算成對方打開過
      const me = url.searchParams.get('me') === '1' ? '&me=1' : '';
      const res = await fetch(plain + '&raw=1' + me, { redirect: 'follow' });
      const data = await res.json();
      if (!data || !data.ok || !data.data || !data.data.html) throw new Error('bad payload');
      return new Response(data.data.html, {
        status: 200,
        headers: { 'content-type': 'text/html; charset=utf-8', 'cache-control': 'no-store' }
      });
    } catch (err) {
      // 後端沒回應或版本太舊 → 退回讓對方的瀏覽器自己去 Apps Script
      return Response.redirect(plain, 302);
    }
  }
};
