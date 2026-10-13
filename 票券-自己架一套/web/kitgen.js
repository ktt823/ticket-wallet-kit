/* 票券 — 網站產生器（v98）
 *
 *  把朋友包（tickets-kit.zip）在瀏覽器裡解開，把你的 Apps Script /exec 網址填進網站，再打包成一個 zip：
 *    web/_worker.js 的 EXEC、web/404.html 的 EXEC、web/index.html 的 DEFAULT_API。
 *  解壓縮後把 web 資料夾拖進 Cloudflare 就好，不用再設 EXEC_URL、不用重新部署第二次；
 *  手機第一次打開時後端網址也已經填好，只要貼存取碼。
 *
 *  全部在這個瀏覽器裡做，/exec 網址不會傳到任何地方。
 *  用在兩個地方：Martin 網站上的 kit.html（第一次安裝），還有朋友 App 的「更新」「分享連結」頁（下載填好的新版網站）。
 *
 *  KitGen.build(kitUrl, exec)  → Promise<Blob>（zip，裡面是 web/…）
 *  KitGen.readZip / makeZip / personalize 拆開來給測試用。
 */
(function (root) {
  'use strict';
  var EXEC_RE = /^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]{20,}\/exec$/;
  var PH = '在這裡貼上你的 /exec 網址';
  var KIT_DIR = '票券-自己架一套/';

  function cleanExec(s) {
    var v = String(s == null ? '' : s).trim().replace(/^["'「]+|["'」]+$/g, '').replace(/[?#].*$/, '');
    if (!EXEC_RE.test(v)) throw new Error('這不像 Apps Script 的網址：要是 https://script.google.com/macros/s/…/exec');
    return v;
  }

  /* ── 讀 zip（只認「不壓縮」和 deflate，朋友包就是這兩種） ── */
  function u16(b, i) { return b[i] | (b[i + 1] << 8); }
  function u32(b, i) { return (b[i] | (b[i + 1] << 8) | (b[i + 2] << 16) | (b[i + 3] << 24)) >>> 0; }
  function inflateRaw(bytes) {
    if (typeof DecompressionStream !== 'function') {
      return Promise.reject(new Error('這個瀏覽器太舊，解不開壓縮檔：請用新版的 Chrome、Edge 或 Safari'));
    }
    var ds = new DecompressionStream('deflate-raw');
    var out = new Blob([bytes]).stream().pipeThrough(ds);
    return new Response(out).arrayBuffer().then(function (ab) { return new Uint8Array(ab); });
  }
  function readZip(buf) {
    var b = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
    var eocd = -1;
    for (var i = b.length - 22; i >= Math.max(0, b.length - 65557); i--) {
      if (u32(b, i) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) return Promise.reject(new Error('下載到的不是 zip 檔'));
    var n = u16(b, eocd + 10), p = u32(b, eocd + 16);
    var dec = new TextDecoder('utf-8');
    var jobs = [];
    for (var k = 0; k < n; k++) {
      if (u32(b, p) !== 0x02014b50) return Promise.reject(new Error('zip 檔壞了'));
      var method = u16(b, p + 10), csize = u32(b, p + 20), nlen = u16(b, p + 28), xlen = u16(b, p + 30), clen = u16(b, p + 32);
      var off = u32(b, p + 42);
      var name = dec.decode(b.subarray(p + 46, p + 46 + nlen)).replace(/\\/g, '/');
      p += 46 + nlen + xlen + clen;
      if (/\/$/.test(name)) continue;                                  // 資料夾
      var start = off + 30 + u16(b, off + 26) + u16(b, off + 28);
      var raw = b.subarray(start, start + csize);
      jobs.push((function (name, method, raw) {
        if (method === 0) return Promise.resolve({ name: name, data: raw.slice() });
        if (method === 8) return inflateRaw(raw).then(function (d) { return { name: name, data: d }; });
        return Promise.reject(new Error('zip 裡的「' + name + '」用了不認得的壓縮方式'));
      })(name, method, raw));
    }
    return Promise.all(jobs);
  }

  /* ── 寫 zip（不壓縮；圖片本來就壓過了，網站檔案也不大） ── */
  var CRC = null;
  function crc32(d) {
    if (!CRC) {
      CRC = new Uint32Array(256);
      for (var n = 0; n < 256; n++) { var c = n; for (var k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; CRC[n] = c >>> 0; }
    }
    var x = 0xFFFFFFFF;
    for (var i = 0; i < d.length; i++) x = CRC[(x ^ d[i]) & 0xFF] ^ (x >>> 8);
    return (x ^ 0xFFFFFFFF) >>> 0;
  }
  function makeZip(files) {
    var enc = new TextEncoder(), parts = [], central = [], off = 0;
    var now = new Date();
    var dosTime = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1);
    var dosDate = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();
    files.forEach(function (f) {
      var name = enc.encode(f.name), data = f.data, crc = crc32(data);
      var lh = new DataView(new ArrayBuffer(30));
      lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true);   // 0x0800 = 檔名是 UTF-8
      lh.setUint16(8, 0, true); lh.setUint16(10, dosTime, true); lh.setUint16(12, dosDate, true);
      lh.setUint32(14, crc, true); lh.setUint32(18, data.length, true); lh.setUint32(22, data.length, true);
      lh.setUint16(26, name.length, true); lh.setUint16(28, 0, true);
      parts.push(new Uint8Array(lh.buffer), name, data);
      var ch = new DataView(new ArrayBuffer(46));
      ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true); ch.setUint16(8, 0x0800, true);
      ch.setUint16(10, 0, true); ch.setUint16(12, dosTime, true); ch.setUint16(14, dosDate, true);
      ch.setUint32(16, crc, true); ch.setUint32(20, data.length, true); ch.setUint32(24, data.length, true);
      ch.setUint16(28, name.length, true); ch.setUint32(42, off, true);
      central.push(new Uint8Array(ch.buffer), name);
      off += 30 + name.length + data.length;
    });
    var csize = central.reduce(function (s, a) { return s + a.length; }, 0);
    var end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true); end.setUint16(8, files.length, true); end.setUint16(10, files.length, true);
    end.setUint32(12, csize, true); end.setUint32(16, off, true);
    return new Blob(parts.concat(central, [new Uint8Array(end.buffer)]), { type: 'application/zip' });
  }

  /* ── 填進後端網址：只留 web 資料夾 ── */
  function swap(text, from, to, what) {
    if (text.indexOf(from) < 0) throw new Error('朋友包裡的 ' + what + ' 找不到要填網址的地方（版本太舊？）');
    return text.split(from).join(to);
  }
  function personalize(files, exec) {
    exec = cleanExec(exec);
    var enc = new TextEncoder(), dec = new TextDecoder('utf-8');
    var out = [], seen = {};
    files.forEach(function (f) {
      var name = f.name.indexOf(KIT_DIR) === 0 ? f.name.slice(KIT_DIR.length) : f.name;
      if (name.indexOf('web/') !== 0) return;                          // Code.gs、說明檔不用放進網站
      var data = f.data;
      if (name === 'web/_worker.js') data = enc.encode(swap(dec.decode(data), "const EXEC = '" + PH + "';", "const EXEC = '" + exec + "';", '_worker.js'));
      if (name === 'web/404.html')   data = enc.encode(swap(dec.decode(data), "var EXEC = '" + PH + "';", "var EXEC = '" + exec + "';", '404.html'));
      if (name === 'web/index.html') data = enc.encode(swap(dec.decode(data), "var DEFAULT_API = '';", "var DEFAULT_API = '" + exec + "';", 'index.html'));
      seen[name] = 1;
      out.push({ name: name, data: data });
    });
    ['web/index.html', 'web/_worker.js', 'web/404.html', 'web/sw.js'].forEach(function (n) {
      if (!seen[n]) throw new Error('朋友包裡少了 ' + n);
    });
    return out;
  }

  function build(kitUrl, exec) {
    exec = cleanExec(exec);
    return fetch(kitUrl, { cache: 'no-store' })
      .then(function (r) { if (!r.ok) throw new Error('下載不到朋友包（HTTP ' + r.status + '）'); return r.arrayBuffer(); })
      .then(readZip)
      .then(function (files) { return makeZip(personalize(files, exec)); });
  }

  /** 讓瀏覽器存檔 */
  function save(blob, filename) {
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename || '票券網站.zip';
    document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
  }

  root.KitGen = { build: build, save: save, readZip: readZip, makeZip: makeZip, personalize: personalize, cleanExec: cleanExec, crc32: crc32, PH: PH };
})(typeof window !== 'undefined' ? window : globalThis);
