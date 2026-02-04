const XLSX = require("xlsx");

async function followRedirects(startUrl, max = 10) {
  let url = startUrl;
  const visited = [];

  for (let i = 0; i < max; i++) {
    visited.push(url);

    const r = await fetch(url, {
      redirect: "manual",
      headers: {
        "User-Agent": "Mozilla/5.0",
        "Accept": "text/html,*/*"
      }
    });

    if (![301, 302, 303, 307, 308].includes(r.status)) break;

    const loc = r.headers.get("location");
    if (!loc) break;

    url = new URL(loc, url).toString();
  }

  return visited;
}

function extractDownloadUrl(urls) {
  for (const u of urls) {
    try {
      const url = new URL(u);
      const resid = url.searchParams.get("resid");
      const authkey = url.searchParams.get("authkey");
      if (resid && authkey) {
        return `https://onedrive.live.com/download?resid=${resid}&authkey=${authkey}`;
      }
    } catch {}
  }
  return null;
}

module.exports = async (req, res) => {
  try {
    const shareUrl = req.query.url;
    if (!shareUrl) {
      res.status(400).json({ error: "Missing ?url=" });
      return;
    }

    // 1️⃣ redirect zinciri
    const visited = await followRedirects(shareUrl);

    // 2️⃣ gerçek download linkini çıkar
    const downloadUrl = extractDownloadUrl(visited);
    if (!downloadUrl) {
      res.json({
        ok: false,
        error: "Download link bulunamadı",
        visited
      });
      return;
    }

    // 3️⃣ dosyayı indir
    const r = await fetch(downloadUrl);
    if (!r.ok) {
      res.json({
        ok: false,
        step: "download",
        status: r.status
      });
      return;
    }

    const buf = Buffer.from(await r.arrayBuffer());

    // 4️⃣ excel parse
    const wb = XLSX.read(buf, { type: "buffer" });
    const sheet = wb.SheetNames[0];
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheet], { defval: null });

    res.json({
      ok: true,
      sheet,
      rowCount: rows.length,
      rows
    });

  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
};
