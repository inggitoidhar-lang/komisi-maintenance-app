// ===== KONFIG =====
const GITHUB_OWNER = "inggitoidhar-lang";
const GITHUB_REPO = "komisi-maintenance-app";
const GITHUB_BRANCH = "main";

// Menu biar gampang
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Sync")
    .addItem("Setup GitHub Token", "setupGitHubToken")
    .addItem("Sync to GitHub", "syncToGitHub")
    .addToUi();
}

// Simpan token ke Script Properties (AMAN, tidak masuk repo)
function setupGitHubToken() {
  const token = Browser.inputBox(
    "GitHub Token",
    "Paste token GitHub kamu di sini (JANGAN SHARE KE SIAPAPUN):",
    Browser.Buttons.OK_CANCEL
  );
  if (!token || token === "cancel") throw new Error("Dibatalkan. Token belum disimpan.");

  PropertiesService.getScriptProperties().setProperty("GITHUB_TOKEN", token.trim());
  Browser.msgBox("✅ Token tersimpan.\nSekarang pakai menu: Sync → Sync to GitHub");
}

// Ambil semua file project via Apps Script API
function fetchProjectFiles_() {
  const scriptId = ScriptApp.getScriptId();
  const url = `https://script.googleapis.com/v1/projects/${scriptId}/content`;

  const res = UrlFetchApp.fetch(url, {
    method: "GET",
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });

  if (res.getResponseCode() !== 200) {
    throw new Error("Gagal ambil project content (HTTP " + res.getResponseCode() + "): " + res.getContentText());
  }

  return JSON.parse(res.getContentText()).files || [];
}

// Map file Apps Script -> path repo
function mapFile_(f) {
  if (f.type === "SERVER_JS") return { path: `${f.name}.gs`, content: f.source || "" };
  if (f.type === "HTML") return { path: `${f.name}.html`, content: f.source || "" };
  if (f.type === "JSON" && f.name === "appsscript") return { path: "appsscript.json", content: f.source || "" };
  return null;
}

// Helper: detect rate limit-ish message
function isRetryableGitHub_(code, bodyText) {
  const body = String(bodyText || "").toLowerCase();

  // retry untuk server error & too many requests
  if (code >= 500) return true;
  if (code === 429) return true;

  // GitHub rate limit/abuse/secondary limit kadang 403
  if (code === 403) {
    if (body.includes("rate limit")) return true;
    if (body.includes("secondary rate")) return true;
    if (body.includes("abuse detection")) return true;
    if (body.includes("temporarily blocked")) return true;
  }

  return false;
}

// Create / Update file di GitHub (dengan retry)
function upsertGitHubFile_(token, path, content, message) {
  const api = `https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPO}/contents/${encodeURI(path)}`;

  const headers = {
    Authorization: `Bearer ${token}`,
    Accept: "application/vnd.github+json",
    "X-GitHub-Api-Version": "2022-11-28",
    "User-Agent": "AppsScript-GitHubSync",
  };

  // cek sha kalau file sudah ada
  let sha = null;
  const check = UrlFetchApp.fetch(api + `?ref=${GITHUB_BRANCH}`, {
    method: "GET",
    headers,
    muteHttpExceptions: true,
  });
  if (check.getResponseCode() === 200) {
    sha = JSON.parse(check.getContentText()).sha;
  }

  // base64 yang aman: bytes -> base64
  const b64 = Utilities.base64Encode(Utilities.newBlob(content, "text/plain", path).getBytes());

  const body = {
    message,
    content: b64,
    branch: GITHUB_BRANCH,
  };
  if (sha) body.sha = sha;

  const payload = JSON.stringify(body);

  // retry loop
  const maxTry = 4;
  for (let attempt = 1; attempt <= maxTry; attempt++) {
    const res = UrlFetchApp.fetch(api, {
      method: "PUT",
      headers,
      contentType: "application/json",
      payload,
      muteHttpExceptions: true,
    });

    const code = res.getResponseCode();
    const text = res.getContentText();

    if (code === 200 || code === 201) return;

    if (isRetryableGitHub_(code, text) && attempt < maxTry) {
      // backoff + sedikit random biar gak “pola”
      const sleepMs = (attempt * attempt * 800) + Math.floor(Math.random() * 300);
      Utilities.sleep(sleepMs);
      continue;
    }

    throw new Error(`Gagal upsert ${path} (HTTP ${code}): ${text}`);
  }

  throw new Error(`Gagal upsert ${path}: retry habis.`);
}

// Tombol utama: sync semua file
function syncToGitHub() {
  const token = PropertiesService.getScriptProperties().getProperty("GITHUB_TOKEN");
  if (!token) throw new Error("Token belum ada. Jalankan: Sync → Setup GitHub Token");

  const files = fetchProjectFiles_();
  let count = 0;

  for (const f of files) {
    const mapped = mapFile_(f);
    if (!mapped) continue;

    // jeda kecil biar aman dari secondary rate limit
    Utilities.sleep(250);

    upsertGitHubFile_(token, mapped.path, mapped.content, `Sync ${mapped.path} from Apps Script`);
    count++;
  }

  Browser.msgBox(`✅ Sync selesai. ${count} file tersinkron.`);
}
