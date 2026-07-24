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

// Urutan prioritas -- outlet dengan transaksi paling cepat/ramai ditaruh
// duluan, supaya kalau waktu/kapasitas mepet, merekalah yang paling mungkin
// berhasil ter-refresh duluan. GANTI URUTAN INI sesuai kondisi nyata toko
// Anda -- ini baru asumsi awal berdasarkan pola kegagalan di RunLog
// (Makassar & Mamu Makassar paling sering gagal, kemungkinan karena
// item/varian-nya paling banyak = kemungkinan juga paling ramai transaksi).
const OUTLET_PRIORITY = [
  442608,  // TIGALAPANKAOS MAKASSAR
  443176,  // MAMU MAKASSAR
  443177,  // TIGALAPANKAOS PALU
  443172,  // TIGALAPANKAOS MANADO
  1099711, // MAMU BTP
  1124685, // MAMU KENDARI
  1100076, // MAMU PALOPO
  1125290, // TIGALAPANKAOS POLMAN
  1128052, // TIGALAPANKAOS BONE
  1111792, // TIGALAPANKAOS SAMARINDA
  989704,  // TIGALAPANKAOS TORAJA
  1145319, // MAMU GOWA
];

const PER_PAGE = 900;
const TOKEN_URL = "https://api.mokapos.com/oauth/token";
const API_BASE = "https://api.mokapos.com/v1";
const TIMEZONE = "Asia/Makassar";

// ====== KONFIGURASI CONCURRENCY & RETRY ======
// Titik awal yang cukup konservatif -- BUKAN angka final. Pantau kolom
// DurasiDetik/DurasiTotalDetik di RunLog/TriggerLog setelah berjalan
// beberapa hari, lalu sesuaikan naik/turun berdasarkan data nyata:
//   - Masih sering gagal (429/5xx)?      -> turunkan OUTLET_CONCURRENCY
//   - Tidak pernah gagal & durasi kecil? -> boleh naikkan sedikit
const OUTLET_CONCURRENCY = 3;   // maksimal 3 outlet di-fetch bersamaan
const PAGE_CONCURRENCY = 4;     // maksimal 4 halaman (dalam 1 outlet) bersamaan
const RETRY_ATTEMPTS = 3;       // percobaan ulang kalau gagal (di luar percobaan pertama)
const RETRY_BASE_DELAY_MS = 600; // 600ms, 1200ms, 2400ms (exponential backoff)

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const MOKA_CLIENT_ID = process.env.MOKA_CLIENT_ID;
const MOKA_CLIENT_SECRET = process.env.MOKA_CLIENT_SECRET;
const MOKA_REFRESH_TOKEN_ENV = process.env.MOKA_REFRESH_TOKEN;
const GOOGLE_CREDENTIALS_JSON = process.env.GOOGLE_CREDENTIALS_JSON;

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

// ============================================================================
// CONCURRENCY LIMITER
// Jalankan array of async task, TAPI maksimal `limit` yang jalan bersamaan
// di satu waktu -- sisanya antre sampai ada slot kosong. Item diproses
// SESUAI URUTAN ARRAY-nya (worker ambil item berikutnya secara berurutan),
// jadi kalau array-nya sudah diurutkan sesuai prioritas, item prioritas
// tinggi otomatis kebagian slot duluan.
// ============================================================================
async function mapWithConcurrency(items, limit, worker) {
  const results = new Array(items.length);
  let cursor = 0;

  async function runNext() {
    while (cursor < items.length) {
      const i = cursor++;
      results[i] = await worker(items[i], i);
    }
  }

  const runners = Array.from({ length: Math.min(limit, items.length) }, runNext);
  await Promise.all(runners);
  return results;
}

// ============================================================================
// FETCH DENGAN RETRY + EXPONENTIAL BACKOFF
// 429 (rate limit) & 5xx (error sesaat sisi server) layak dicoba ulang.
// 4xx lain (401, 404, dll) biasanya masalah permanen -- percuma diulang,
// langsung dilempar sebagai error supaya cepat diketahui akar masalahnya.
// ============================================================================
async function fetchWithRetry(url, options, retries = RETRY_ATTEMPTS, baseDelayMs = RETRY_BASE_DELAY_MS) {
  for (let attempt = 0; attempt <= retries; attempt++) {
    let resp;
    try {
      resp = await fetch(url, options);
    } catch (networkErr) {
      if (attempt < retries) {
        const wait = baseDelayMs * Math.pow(2, attempt);
        console.warn(`Network error (${networkErr.message}), coba lagi dalam ${wait}ms (percobaan ${attempt + 1}/${retries})`);
        await sleep(wait);
        continue;
      }
      throw networkErr;
    }

    if (resp.ok) return resp;

    const retryable = resp.status === 429 || resp.status >= 500;
    if (retryable && attempt < retries) {
      const wait = baseDelayMs * Math.pow(2, attempt);
      console.warn(`HTTP ${resp.status} dari ${url}, coba lagi dalam ${wait}ms (percobaan ${attempt + 1}/${retries})`);
      await sleep(wait);
      continue;
    }

    throw new Error(`HTTP ${resp.status}: ${await resp.text()}`);
  }
}

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

  const resp = await fetchWithRetry(TOKEN_URL, {
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
  const resp = await fetchWithRetry(url, { headers: { Authorization: `Bearer ${token}` } });
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

// Halaman ke-2 dst sekarang dibatasi PAGE_CONCURRENCY (bukan Promise.all
// tanpa batas seperti sebelumnya) -- outlet besar (Makassar, Mamu Makassar)
// tidak lagi menembak semua halamannya sekaligus dalam 1 gelombang.
async function crawlOutlet(outletId, outletName, dateStr, token) {
  const firstPage = await fetchPage(outletId, 1, dateStr, token, PER_PAGE);
  const totalPages = Math.max(1, Math.ceil((firstPage.total_counts || 0) / PER_PAGE));
  let rows = flattenResults(firstPage.results, outletName);

  if (totalPages > 1) {
    const remainingPageNumbers = Array.from({ length: totalPages - 1 }, (_, i) => i + 2);
    const pageResults = await mapWithConcurrency(remainingPageNumbers, PAGE_CONCURRENCY, (p) =>
      fetchPage(outletId, p, dateStr, token, PER_PAGE)
    );
    pageResults.forEach((pd) => { rows = rows.concat(flattenResults(pd.results, outletName)); });
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

// ids sekarang diurutkan sesuai OUTLET_PRIORITY (bukan Object.keys(OUTLETS)
// apa adanya) -- outlet prioritas tinggi diproses duluan oleh
// mapWithConcurrency, jadi kalau waktu/kapasitas mepet, merekalah yang
// paling mungkin sempat berhasil.
function getPrioritizedOutletIds() {
  const allIds = Object.keys(OUTLETS);
  const priorityIds = OUTLET_PRIORITY.map(String).filter((id) => allIds.includes(id));
  const remainingIds = allIds.filter((id) => !priorityIds.includes(id));
  return [...priorityIds, ...remainingIds];
}

async function refreshAllOutlets() {
  const runStart = Date.now();
  const sheets = await getSheetsClient();
  const token = await getValidAccessToken(sheets);
  const dateStr = formatDateDDMMYYYY(new Date());
  const ids = getPrioritizedOutletIds();
  let done = 0;
  const failedOutlets = [];
  const freshTimestamps = {};

  // Outlet diproses maksimal OUTLET_CONCURRENCY bersamaan, sesuai urutan
  // prioritas. Kegagalan 1 outlet ditangkap di sini (bukan lewat
  // Promise.allSettled global) supaya outlet lain tidak terpengaruh.
  await mapWithConcurrency(ids, OUTLET_CONCURRENCY, async (outletId) => {
    try {
      const rows = await crawlOutlet(outletId, OUTLETS[outletId], dateStr, token);
      await writeOutletCache(sheets, outletId, rows);
      console.log(`OK ${outletId} (${OUTLETS[outletId]}): ${rows.length} baris`);
      done++;
      freshTimestamps[outletId] = formatTimestampWITA(new Date());
    } catch (err) {
      console.error(`GAGAL ${outletId} (${OUTLETS[outletId]}): ${err.message}`);
      failedOutlets.push(OUTLETS[outletId]);
    }
  });

  // 1x tulis Meta di akhir, sekuensial -- bukan paralel per-outlet -- supaya
  // tidak ada race condition yang bikin timestamp "macet".
  await updateMetaBatch(sheets, ids, freshTimestamps);

  const durationSec = Math.round((Date.now() - runStart) / 1000);
  const nowWita = formatTimestampWITA(new Date());
  console.log(`Selesai: ${done}/${ids.length} outlet, ${durationSec} detik.`);
  await logRun(sheets, { timestamp: nowWita, outletsDone: done, outletGagal: failedOutlets.join(", "), durationSec });
  return { done, total: ids.length, failedOutlets, durationSec };
}

(async () => {
  const triggerStart = Date.now();
  try {
    const result1 = await refreshAllOutlets();
    console.log("Refresh #1 selesai:", JSON.stringify(result1));

    // Jeda ADAPTIF -- target jarak refresh #1 ke #2 tetap 150 detik (2,5
    // menit), TAPI kalau refresh #1 sudah makan waktu lebih dari itu
    // (karena throttling/retry), langsung lanjut ke refresh #2 tanpa nunggu
    // tambahan. Ini supaya cadence 2,5 menit tetap jadi target utama, tapi
    // tidak dipaksakan kaku sampai bikin total durasi kelewat timeout job.
    const targetGapMs = 150 * 1000;
    const elapsedAfterRun1 = Date.now() - triggerStart;
    const remainingWait = targetGapMs - elapsedAfterRun1;

    if (remainingWait > 0) {
      console.log(`Refresh #1 selesai dalam ${Math.round(elapsedAfterRun1 / 1000)}s, tunggu sisa ${Math.round(remainingWait / 1000)}s sebelum refresh #2.`);
      await sleep(remainingWait);
    } else {
      console.log(`Refresh #1 makan waktu ${Math.round(elapsedAfterRun1 / 1000)}s (lebih dari target 150s) -- langsung lanjut refresh #2 tanpa jeda tambahan.`);
    }

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

    // Status akhir job ditentukan HANYA dari refresh #2 (yang terakhir) --
    // kalau refresh #1 sempat gagal di beberapa outlet tapi refresh #2
    // berhasil menutupinya (gangguan sesaat sisi Moka, dll), data di akhir
    // siklus ini tetap LENGKAP, jadi tidak seharusnya ditandai gagal.
    // Kegagalan refresh #1 tetap tercatat di RunLog/TriggerLog untuk arsip,
    // cuma tidak menentukan status merah/hijau job ini.
    const anyFailed = result2.failedOutlets.length > 0;
    process.exit(anyFailed ? 1 : 0);
  } catch (err) {
    console.error("FATAL:", err);
    process.exit(1);
  }
})();
