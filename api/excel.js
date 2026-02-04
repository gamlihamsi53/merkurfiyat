const XLSX = require("xlsx");

function uaHeaders(html = true) {
  return {
    "User-Agent":
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": html
      ? "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
      : "*/*",
    "Accept-Language": "tr-TR,tr;q=0.9,en-US;q=0.8,en;q=0.7",
  };
}

async function followRedirects(startUrl, max = 12) {
  let url = startUrl;
  const visited = [];

  for (let i = 0; i < max; i++) {
    visited.push(url);

    const r = await fetch(url, {
      method: "GET",
      redirect: "manual",
      headers: uaHeaders(true),
    });

    if (![301, 302, 303, 307, 308].includes(r.status)) break;

    const loc = r.headers.get("location");
    if (!loc) break;

    url = new URL(loc, url).toString();
  }

  return { finalUrl: url, visited };
}

function extractResid(urls) {
  for (const u of urls) {
    try {
      const url = new URL(u);
      const resid = url.searchParams.get("resid");
      if (resid) return resid;
    } catch {}
  }
  return null;
}

function looksLikeXlsx(buf) {
  // XLSX is a zip => starts with "PK"
  return buf.length >= 2 && buf[0] === 0x50 && buf[1] === 0x4b;
}

async function tryDownload(url) {
  const r = await fetch(url, {
    method: "GET",
    redirect: "follow",
    headers: uaHeaders(false),
  });
  if (!r.ok) return { ok: false, status: r.status, url };

  const buf = Buffer.from(await r.arrayBuffer());
  if (!looksLikeXlsx(buf)) {
    return {
      ok: false,
      status: r.status,
      url,
      reason: "Not XLSX (PK signature missing)",
      contentType: r.headers.get("content-type"),
      first16Hex: buf.slice(0, 16).toString("hex"),
    };
  }
  return { ok: true, status: r.status, url, buf };
}

function unescapeJsonString(s) {
  // handles \" and \u0026 etc.
  try {
    return JSON.parse(`"${s.replace(/"/g, '\\"')}"`);
  } catch {
    // fallback minimal
    return s.replace(/\\u0026/g, "&").replace(/\\"/g, '"').replace(/\\\\/g, "\\");
  }
}

function extractDownloadUrlFromHtml(html) {
  // 1) @microsoft.graph.downloadUrl
  let m = html.match(/"@microsoft\.graph\.downloadUrl"\s*:\s*"([^"]+)"/);
  if (m && m[1]) return unescapeJsonString(m[1]);

  // 2) "downloadUrl":"..."
  m = html.match(/"downloadUrl"\s*:\s*"([^"]+)"/);
  if (m && m[1]) return unescapeJsonString(m[1]);

  // 3) any onedrive.live.com/download?... found in html
  m = html.match(/https:\/\/onedrive\.live\.com\/download\?[^"' <]+/);
  if (m && m[0]) return m[0];

  // 4) any 1drv direct file host (sometimes public.dm.files.1drv.com)
  m = html.match(/https:\/\/[^"' <]+\.files\.1drv\.com\/[^"' <]+/);
  if (m && m[0]) return m[0];

  return null;
}

module.exports = async (req, res) => {
  try {
    const shareUrl = req.query.url ? req.query.url.toString() : null;
    if (!shareUrl) {
      res.status(400).json({ error: "Missing ?url=" });
      return;
    }

    // A) Redirect zinciri
    const { finalUrl, visited } = await followRedirects(shareUrl, 12);
    const resid = extractResid(visited) || extractResid([finalUrl]);

    // B) 1) Önce direkt download?resid= dene (authkey yoksa bazen yeter)
    if (resid) {
      const direct = await tryDownload(`https://onedrive.live.com/download?resid=${encodeURIComponent(resid)}`);
      if (direct.ok) {
        const wb = XLSX.read(direct.buf, { type: "buffer" });
        const sheet = wb.SheetNames[0];
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheet], { defval: null });
        res.status(200).json({ ok: true, source: "download?resid", sheet, rowCount: rows.length, rows });
        return;
      }
    }

    // C) 2) View sayfasının HTML’ini çek, içinden downloadUrl çıkar
    const pageResp = await fetch(finalUrl, {
      method: "GET",
      redirect: "follow",
      headers: uaHeaders(true),
    });

    const html = await pageResp.text();
    const downloadUrl = extractDownloadUrlFromHtml(html);

    if (!downloadUrl) {
      res.status(200).json({
        ok: false,
        error: "Download URL HTML içinde bulunamadı",
        debug: {
          finalUrl,
          resid,
          visited,
          pageStatus: pageResp.status,
          pageContentType: pageResp.headers.get("content-type"),
          htmlSnippet: html.slice(0, 400),
        },
      });
      return;
    }

    const dl = await tryDownload(downloadUrl);
    if (!dl.ok) {
      res.status(200).json({
        ok: false,
        error: "Download denemesi başarısız / XLSX değil",
        debug: { downloadUrl, dl, finalUrl, resid },
      });
      return;
    }

    const wb = XLSX.read(dl.buf, { type: "buffer" });
    const sheet = wb.SheetNames[0];
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheet], { defval: null });

    res.status(200).json({
      ok: true,
      source: "html-extract",
      sheet,
      rowCount: rows.length,
      rows,
    });
  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
};
