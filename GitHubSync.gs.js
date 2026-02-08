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
  if (!token || token === "cancel") {
    throw new Error("Dibatalkan. Token belum disimpan.");
  }
  PropertiesService.getScriptProperties()
    .setProperty("GITHUB_TOKEN", token.trim());
  Browser.msgBox("✅ Token tersimpan. Sekarang pakai menu: Sync → Sync to GitHub");
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
    throw new Error("Gagal ambil project content: " + res.getContentText());
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

// Create / Update file di GitHub
function upsertGitHubFile_(token, path, content, message) {
  const api = `https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPO}/contents/${encodeURI(path)}`;
  const headers = {
    Authorization: `Bearer ${token}`,
    Accept: "application/vnd.github+json",
    "X-GitHub-Api-Version": "2022-11-28",
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

  const body = {
    message,
    content: Utilities.base64Encode(content),
    branch: GITHUB_BRANCH,
  };
  if (sha) body.sha = sha;

  const res = UrlFetchApp.fetch(api, {
    method: "PUT",
    headers,
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  });
  if (![200, 201].includes(res.getResponseCode())) {
    throw new Error(`Gagal upsert ${path}: ${res.getContentText()}`);
  }
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
    upsertGitHubFile_(token, mapped.path, mapped.content, `Sync ${mapped.path} from Apps Script`);
    count++;
  }
  Browser.msgBox(`✅ Sync selesai. ${count} file tersinkron.`);
}
