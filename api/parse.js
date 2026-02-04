const XLSX = require("xlsx");

function looksLikeXlsx(buf) {
  return buf.length >= 2 && buf[0] === 0x50 && buf[1] === 0x4b; // PK
}

module.exports = async (req, res) => {
  try {
    const downloadUrl = req.query.downloadUrl ? req.query.downloadUrl.toString() : null;
    if (!downloadUrl) {
      res.status(400).json({ error: "Missing ?downloadUrl=" });
      return;
    }

    const r = await fetch(downloadUrl, {
      redirect: "follow",
      headers: { "User-Agent": "Mozilla/5.0", Accept: "*/*" },
    });

    if (!r.ok) {
      res.status(200).json({ ok: false, status: r.status, contentType: r.headers.get("content-type") });
      return;
    }

    const buf = Buffer.from(await r.arrayBuffer());
    if (!looksLikeXlsx(buf)) {
      res.status(200).json({
        ok: false,
        error: "Not an XLSX (PK missing)",
        contentType: r.headers.get("content-type"),
        first16Hex: buf.slice(0, 16).toString("hex"),
      });
      return;
    }

    const wb = XLSX.read(buf, { type: "buffer" });
    const sheet = wb.SheetNames[0];
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheet], { defval: null });

    res.status(200).json({ ok: true, sheet, rowCount: rows.length, rows });
  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
};
