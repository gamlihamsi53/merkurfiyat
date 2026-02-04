const XLSX = require("xlsx");

async function safeText(resp, n = 300) {
  try {
    const t = await resp.text();
    return t.slice(0, n);
  } catch {
    return null;
  }
}

function isRedirect(status) {
  return status === 301 || status === 302 || status === 303 || status === 307 || status === 308;
}

function uaHeaders() {
  return {
    "User-Agent":
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept":
      "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "tr-TR,tr;q=0.9,en-US;q=0.8,en;q=0.7",
  };
}

// 1drv.ms / onedrive.live.com redirect zincirini çözer.
// A yöntemi çalışıyorsa finalUrl login.live.com OLMAMALI.
async function resolveRedirectChain(startUrl, maxHops = 12) {
  let url = startUrl;
  const visited = [];

  for (let i = 0; i < maxHops; i++) {
    visited.push(url);

    const r = await fetch(url, {
      method: "GET",
      redirect: "manual",
      headers: uaHeaders(),
    });

    if (!isRedirect(r.status)) {
      return {
        ok: r.ok,
        finalStatus: r.status,
        finalUrl: url,
        visited,
        contentType: r.headers.get("content-type"),
        bodySnippet: r.ok ? null : await safeText(r),
      };
    }

    const loc = r.headers.get("location");
    if (!loc) {
      return { ok: false, finalStatus: r.status, finalUrl: url, visited, error: "No Location header" };
    }

    url = new URL(loc, url).toString();
  }

  return { ok: false, finalStatus: 0, finalUrl: url, visited, error: `Max hops (${maxHops}) exceeded` };
}

module.exports = async (req, res) => {
  try {
    const mode = (req.query.mode || "resolve").toString(); // resolve | download | parse
    const shareUrl = req.query.url ? req.query.url.toString() : null;

    if (!shareUrl) {
      res.status(400).json({ error: "Missing ?url=" });
      return;
    }

    // ADIM 1: resolve (link anonim mi?)
    if (mode === "resolve") {
      const resolved = await resolveRedirectChain(shareUrl, 12);
      const host = (() => { try { return new URL(resolved.finalUrl).host; } catch { return null; } })();

      res.status(200).json({
        ...resolved,
        finalHost: host,
        isLogin: host ? host.includes("login.live.com") : false,
        note: host?.includes("login.live.com")
          ? "❌ Login istiyor: Link 'Anyone' değil veya dosya erişimi kısıtlı."
          : "✅ Login’e düşmüyor: İndirme testine geçebilirsin (mode=download).",
      });
      return;
    }

    // ADIM 2/3: download veya parse
    const resolved = await resolveRedirectChain(shareUrl, 12);
    const finalHost = (() => { try { return new URL(resolved.finalUrl).host; } catch { return null; } })();

    if (finalHost && finalHost.includes("login.live.com")) {
      res.status(200).json({
        ok: false,
        step: "resolve",
        error: "Login required. Link 'Anyone with the link' değil veya guest access kapalı.",
        resolved,
      });
      return;
    }

    const r = await fetch(resolved.finalUrl, {
      method: "GET",
      redirect: "follow",
      headers: {
        "User-Agent": uaHeaders()["User-Agent"],
        "Accept": "*/*",
      },
    });

    if (!r.ok) {
      res.status(200).json({
        ok: false,
        step: "download",
        httpStatus: r.status,
        contentType: r.headers.get("content-type"),
        bodySnippet: await safeText(r),
        resolved,
      });
      return;
    }

    const buf = Buffer.from(await r.arrayBuffer());

    // ADIM 2: download sonucu
    if (mode === "download") {
      res.status(200).json({
        ok: true,
        step: "download",
        httpStatus: r.status,
        contentType: r.headers.get("content-type"),
        bytes: buf.length,
        // XLSX genelde ZIP ile başlar: "PK" -> 504b
        first4Hex: buf.slice(0, 4).toString("hex"),
        looksLikeXlsx: buf.slice(0, 2).toString("utf8") === "PK",
        note:
          buf.slice(0, 2).toString("utf8") === "PK"
            ? "✅ XLSX/ZIP imzası (PK) görünüyor. Parse’a geçebilirsin (mode=parse)."
            : "⚠️ PK değilse HTML/redirect sayfası olabilir. contentType ve snippet’e bak.",
        resolved,
      });
      return;
    }

    // ADIM 3: parse
    if (mode === "parse") {
      const wb = XLSX.read(buf, { type: "buffer" });
      const sheet = wb.SheetNames[0];
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheet], { defval: null });

      res.status(200).json({
        ok: true,
        step: "parse",
        sheet,
        rowCount: rows.length,
        rows,
      });
      return;
    }

    res.status(400).json({ error: "Invalid mode. Use resolve|download|parse" });
  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
};
