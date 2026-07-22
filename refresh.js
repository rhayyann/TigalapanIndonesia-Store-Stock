const { google } = require("googleapis");

const OUTLETS = {
  442608: "TIGALAPANKAOS MAKASSAR",
  443176: "MAMU MAKASSAR",
  443177: "TIGALAPANKAOS PALU",
  443172: "TIGALAPANKAOS MANADO",
  1099711: "MAMU BTP",
  1124685: "MAMU KENDARI",
  1100076: "MAMU PALOPO",
  1125290: "TIGALAPANKAOS POLMAN",
  1128052: "TIGALAPANKAOS BONE",
  1111792: "TIGALAPANKAOS SAMARINDA",
  989704: "TIGALAPANKAOS TORAJA",
  1145319: "MAMU GOWA",
};

const ALLOWED_CATEGORIES = [
  "COMBED 24S", "COMBED 30S", "KIDS 24S",
  "PANJANG + RIB", "TUNIK 24S", "WANGKI MYNO",
];

const PER_PAGE = 900;
const TOKEN_URL = "https://api.mokapos.com/oauth/token";
const API_BASE = "https://api.mokapos.com/v1";
const TIMEZONE = "Asia/Makassar";

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const MOKA_CLIENT_ID = process.env.MOKA_CLIENT_ID;
const MOKA_CLIENT_SECRET = process.env.MOKA_CLIENT_SECRET;
const MOKA_REFRESH_TOKEN_ENV = process.env.MOKA_REFRESH_TOKEN;
const GOOGLE_CREDENTIALS_JSON = process.env.GOOGLE_CREDENTIALS_JSON;

async function getSheetsClient() {
  const credentials = JSON.parse(GOOGLE_CREDENTIALS_JSON);
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
  const authClient = await auth.getClient();
  return google.sheets({ version: "v4", auth: authClient });
}

async function ensureSheetExists(sheets, sheetTitle) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const exists = meta.data.sheets.some((s) => s.properties.title === sheetTitle);
  if (exists) return;
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: { requests: [{ addSheet: { properties: { title: sheetTitle } } }] },
  });
}

async function getValidAccessToken(sheets) {
  await ensureSheetExists(sheets, "TokenStore");
  let accessToken = null, refreshToken = MOKA_REFRESH_TOKEN_ENV, expiresAt = 0;
  try {
    const res = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: "TokenStore!A2:C2" });
    const row = res.data.values && res.data.values[0];
    if (row) { accessToken = row[0] || null; refreshToken = row[1] || MOKA_REFRESH_TOKEN_ENV; expiresAt = Number(row[2]) || 0; }
  } catch (e) {}

  if (accessToken && expiresAt > Date.now() + 5 * 60 * 1000) return accessToken;

  const resp = await fetch(TOKEN_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ client_id: MOKA_CLIENT_ID, client_secret: MOKA_CLIENT_SECRET, refresh_token: refreshToken, grant_type: "refresh_token" }),
  });
  const body = await resp.json();
  if (!body.access_token) throw new Error("Gagal refresh token: " + JSON.stringify(body));

  const newExpiresAt = Date.now() + body.expires_in * 1000;
  const newRefreshToken = body.refresh_token || refreshToken;
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID, range: "TokenStore!A1:C2", valueInputOption: "RAW",
    requestBody: { values: [["access_token", "refresh_token", "expires_at"], [body.access_token, newRefreshToken, String(newExpiresAt)]] },
  });
  return body.access_token;
}

function formatDateDDMMYYYY(date) {
  const parts = new Intl.DateTimeFormat("en-GB", { timeZone: TIMEZONE, day: "2-digit", month: "2-digit", year: "numeric" }).formatToParts(date);
  const get = (t) => parts.find((p) => p.type === t).value;
  return `${get("day")}/${get("month")}/${get("year")}`;
}

// Format "yyyy-MM-dd HH:mm:ss" dalam WITA (bukan UTC) -- dipakai di RunLog
// supaya jamnya langsung terbaca tanpa perlu dikonversi manual.
function formatTimestampWITA(date) {
  const parts = new Intl.DateTimeFormat("en-GB", {
    timeZone: TIMEZONE, year: "numeric", month: "2-digit", day: "2-digit",
    hour: "2-digit", minute: "2-digit", second: "2-digit", hour12: false,
  }).formatToParts(date);
  const get = (t) => parts.find((p) => p.type === t).value;
  return `${get("year")}-${get("month")}-${get("day")} ${get("hour")}:${get("minute")}:${get("second")}`;
}

async function fetchPage(outletId, page, dateStr, token, perPage) {
  const url = `${API_BASE}/outlets/${outletId}/inventory/item_summaries?start=${dateStr}&end=${dateStr}&page=${page}&per_page=${perPage}`;
  const resp = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!resp.ok) throw new Error(`HTTP ${resp.status}: ${await resp.text()}`);
  return (await resp.json()).data;
}

function flattenResults(results, outletName) {
  const out = [];
  (results || []).forEach((item) => {
    if (!ALLOWED_CATEGORIES.includes(item.category)) return;
    (item.item_variants || []).forEach((v) => out.push([item.item_name, item.category, outletName, Number(v.ending) || 0]));
  });
  return out;
}

async function crawlOutlet(outletId, outletName, dateStr, token) {
  const firstPage = await fetchPage(outletId, 1, dateStr, token, PER_PAGE);
  const totalPages = Math.max(1, Math.ceil((firstPage.total_counts || 0) / PER_PAGE));
  let rows = flattenResults(firstPage.results, outletName);
  if (totalPages > 1) {
    const rest = await Promise.all(Array.from({ length: totalPages - 1 }, (_, i) => fetchPage(outletId, i + 2, dateStr, token, PER_PAGE)));
    rest.forEach((pd) => { rows = rows.concat(flattenResults(pd.results, outletName)); });
  }
  return rows;
}

async function writeOutletCache(sheets, outletId, rows) {
  const sheetName = String(outletId);
  await ensureSheetExists(sheets, sheetName);
  await sheets.spreadsheets.values.clear({ spreadsheetId: SPREADSHEET_ID, range: sheetName });
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID, range: `${sheetName}!A1`, valueInputOption: "RAW",
    requestBody: { values: [["nv", "category", "outlet", "ending"], ...rows] },
  });
}

// Tulis SEMUA timestamp Meta sekaligus dalam 1 operasi, di akhir run --
// bukan 12x paralel per-outlet seperti sebelumnya (itu penyebab race
// condition yang bikin sebagian timestamp Meta "macet"/tidak ter-update).
// Outlet yang gagal di run ini TETAP mempertahankan timestamp lama miliknya
// (bukan dihapus/dikosongkan), supaya kamu tetap tahu kapan terakhir kali
// outlet itu benar-benar sukses di-fetch.
async function updateMetaBatch(sheets, allOutletIds, freshTimestamps) {
  await ensureSheetExists(sheets, "Meta");

  let existing = {};
  try {
    const res = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: "Meta!A2:B" });
    (res.data.values || []).forEach((row) => {
      if (row[0]) existing[row[0]] = row[1] || "";
    });
  } catch (e) {}

  const merged = allOutletIds.map((id) => {
    const ts = freshTimestamps[id] || existing[id] || "";
    return [String(id), ts];
  });

  await sheets.spreadsheets.values.clear({ spreadsheetId: SPREADSHEET_ID, range: "Meta" });
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID, range: "Meta!A1", valueInputOption: "RAW",
    requestBody: { values: [["OutletId", "FetchedAt"], ...merged] },
  });
}

async function logRun(sheets, entry) {
  try {
    await ensureSheetExists(sheets, "RunLog");
    const check = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: "RunLog!A1:A1" }).catch(() => ({ data: {} }));
    if (!check.data.values) {
      await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range: "RunLog!A1:F1", valueInputOption: "RAW", requestBody: { values: [["Timestamp", "OutletsSelesai", "OutletGagal", "DurasiDetik", "Platform", "Catatan"]] } });
    }
    await sheets.spreadsheets.values.append({ spreadsheetId: SPREADSHEET_ID, range: "RunLog!A:F", valueInputOption: "RAW", insertDataOption: "INSERT_ROWS", requestBody: { values: [[entry.timestamp, entry.outletsDone, entry.outletGagal, entry.durationSec, "github-actions", entry.catatan || ""]] } });
  } catch (e) { console.error("Gagal menulis RunLog:", e.message); }
}

// TriggerLog -- 1 baris per KALI GitHub Actions dipicu (bukan per refresh).
// Merangkum hasil refresh #1 dan #2 sekaligus, supaya gampang lihat apakah
// 1 kali trigger berhasil mengambil SEMUA 12 outlet di kedua refresh-nya,
// tanpa perlu bolak-balik baca 2 baris terpisah di RunLog.
async function logTrigger(sheets, entry) {
  try {
    await ensureSheetExists(sheets, "TriggerLog");
    const check = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: "TriggerLog!A1:A1" }).catch(() => ({ data: {} }));
    if (!check.data.values) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID, range: "TriggerLog!A1:H1", valueInputOption: "RAW",
        requestBody: { values: [["Timestamp", "Refresh1_Sukses", "Refresh1_OutletGagal", "Refresh2_Sukses", "Refresh2_OutletGagal", "DurasiTotalDetik", "Status", "Platform"]] },
      });
    }
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID, range: "TriggerLog!A:H", valueInputOption: "RAW", insertDataOption: "INSERT_ROWS",
      requestBody: { values: [[entry.timestamp, entry.r1Sukses, entry.r1Gagal, entry.r2Sukses, entry.r2Gagal, entry.durasiTotal, entry.status, "github-actions"]] },
    });
  } catch (e) { console.error("Gagal menulis TriggerLog:", e.message); }
}

async function refreshAllOutlets() {
  const runStart = Date.now();
  const sheets = await getSheetsClient();
  const token = await getValidAccessToken(sheets);
  const dateStr = formatDateDDMMYYYY(new Date());
  const ids = Object.keys(OUTLETS);
  let done = 0;
  const failedOutlets = [];

  const results = await Promise.allSettled(ids.map(async (outletId) => {
    const rows = await crawlOutlet(outletId, OUTLETS[outletId], dateStr, token);
    await writeOutletCache(sheets, outletId, rows);
    return { outletId, rowCount: rows.length };
  }));

  const nowWita = formatTimestampWITA(new Date());
  const freshTimestamps = {};
  results.forEach((r, i) => {
    const outletId = ids[i];
    if (r.status === "fulfilled") {
      console.log(`OK ${outletId} (${OUTLETS[outletId]}): ${r.value.rowCount} baris`);
      done++;
      freshTimestamps[outletId] = nowWita;
    } else {
      console.error(`GAGAL ${outletId} (${OUTLETS[outletId]}): ${r.reason}`);
      failedOutlets.push(OUTLETS[outletId]);
    }
  });

  // 1x tulis Meta di akhir, sekuensial -- bukan paralel per-outlet -- supaya
  // tidak ada race condition yang bikin timestamp "macet".
  await updateMetaBatch(sheets, ids, freshTimestamps);

  const durationSec = Math.round((Date.now() - runStart) / 1000);
  console.log(`Selesai: ${done}/${ids.length} outlet, ${durationSec} detik.`);
  await logRun(sheets, { timestamp: nowWita, outletsDone: done, outletGagal: failedOutlets.join(", "), durationSec });
  return { done, total: ids.length, failedOutlets, durationSec };
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

(async () => {
  const triggerStart = Date.now();
  try {
    const result1 = await refreshAllOutlets();
    console.log("Refresh #1 selesai:", JSON.stringify(result1));

    // Jeda 2,5 menit lalu refresh sekali lagi -- karena jadwal cron GitHub
    // Actions minimum 5 menit, trik ini bikin refresh AKTUAL jadi ~2,5 menit
    // sekali (refresh #1 di awal run, refresh #2 di tengah, lalu trigger
    // berikutnya 5 menit kemudian mengulang pola yang sama).
    await sleep(150 * 1000);

    const result2 = await refreshAllOutlets();
    console.log("Refresh #2 selesai:", JSON.stringify(result2));

    const sheets = await getSheetsClient();
    const totalOutlets = Object.keys(OUTLETS).length;
    const fullSuccess = result1.done === totalOutlets && result2.done === totalOutlets;
    await logTrigger(sheets, {
      timestamp: formatTimestampWITA(new Date(triggerStart)),
      r1Sukses: `${result1.done}/${totalOutlets}`,
      r1Gagal: result1.failedOutlets.join(", "),
      r2Sukses: `${result2.done}/${totalOutlets}`,
      r2Gagal: result2.failedOutlets.join(", "),
      durasiTotal: Math.round((Date.now() - triggerStart) / 1000),
      status: fullSuccess ? "Sukses penuh" : "Sebagian gagal",
    });

    const anyFailed = result1.failedOutlets.length > 0 || result2.failedOutlets.length > 0;
    process.exit(anyFailed ? 1 : 0);
  } catch (err) {
    console.error("FATAL:", err);
    process.exit(1);
  }
})();
