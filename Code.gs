/***** CONFIG *****/
// Spreadsheet ID sekarang ambil dari Script Properties: SPREADSHEET_ID
const SHEET_NAME = "DB_Slip_Komisi_Cair";

// Sheet untuk histori pengerjaan
const SHEET_PENGERJAAN = "DB_List_Order";

// ✅ Sheet untuk database akun (sumber kebenaran user)
const SHEET_AKUN = "DB_Akun";

// Logo publik (wajib URL publik)
const LOGO_URL = "https://res.cloudinary.com/dkps3vy8m/image/upload/v1770355270/Logo_Jendela_Warna_asbatd.png";
// Logo putih untuk header beranda
const LOGO_WHITE_URL = "https://res.cloudinary.com/dkps3vy8m/image/upload/v1770465091/Logo_Jendela_Putih_yykal8.png";

/***** SPREADSHEET HELPER (WAJIB) *****/
function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  if (!id) {
    throw new Error('SPREADSHEET_ID belum diset di Script Properties. Isi dengan ID Google Sheet database.');
  }
  return SpreadsheetApp.openById(id);
}

/***** OTP CONFIG *****/
const OTP_TTL_SEC = 5 * 60;             // 5 menit
const OTP_COOLDOWN_SEC = 45;            // cooldown kirim ulang
const REMEMBER_TTL_DAYS = 30;           // ingat saya 30 hari

// Target default (placeholder; nanti bisa kamu ubah via Script Properties "TARGET_SPK_PER_MONTH")
const DEFAULT_TARGET_SPK_PER_MONTH = 100;

/***** ROUTING WEB APP *****/
function doGet(e) {
  const pageRaw = (e && e.parameter && e.parameter.page) ? String(e.parameter.page) : "login";
  const page = String(pageRaw || "").trim().toLowerCase();

  const appUrl = ScriptApp.getService().getUrl();
  const email = (e && e.parameter && e.parameter.email) ? String(e.parameter.email).trim().toLowerCase() : "";

  let file = "login";
  if (page === "home") file = "dashboard";
  if (page === "login") file = "login";
  if (page === "dashboard") file = "dashboard";
  if (page === "komisi") file = "rincianpencairankomisi";

  // ✅ histori pengerjaan
  if (page === "histori") file = "rincianpengerjaan";
  if (page === "pengerjaan") file = "rincianpengerjaan";

  if (page === "ranking") file = "ranking";
  if (page === "bersih") file = "placeholder";

  try {
    const t = HtmlService.createTemplateFromFile(file);
    t.logoUrl = LOGO_URL;
    t.logoWhiteUrl = LOGO_WHITE_URL;
    t.page = page;
    t.appUrl = appUrl;
    t.prefillEmail = email;

    t.__debug = { file: file + ".html", page: page, deployedAt: new Date().toISOString() };

    return t.evaluate()
      .setTitle("Pemeriksa Komisi – Maintenance")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (err) {
    return HtmlService.createHtmlOutput(
      `<div style="font-family:system-ui;padding:16px">
        <h2>Halaman tidak ditemukan</h2>
        <div>page=<b>${page}</b></div>
        <div>File yang diminta=<b>${file}</b> (harus ada file ${file}.html di project ini)</div>
        <hr/>
        <div style="color:#b00020;font-weight:800">Detail error:</div>
        <pre style="white-space:pre-wrap">${String(err && err.message ? err.message : err)}</pre>
        <hr/>
        <div style="font-weight:800">Checklist cepat:</div>
        <ol>
          <li>Pastikan URL /exec yang dibuka berasal dari deployment project ini.</li>
          <li>Pastikan file HTML ada: login.html, dashboard.html, rincianpencairankomisi.html, rincianpengerjaan.html, placeholder.html, ranking.html</li>
          <li>Setelah ganti file/kode: Deploy → Manage deployments → Edit → New version → Deploy</li>
          <li>Buka lagi /exec di tab baru atau hard refresh</li>
        </ol>
      </div>`
    );
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/***** UTILS *****/
function normHeader_(s) {
  return String(s || "")
    .toLowerCase()
    .trim()
    .replace(/\s+/g, " ")
    .replace(/[^\p{L}\p{N}]+/gu, "");
}
function idxOfHeader_(headerRow, headerName) {
  const target = normHeader_(headerName);
  for (let i = 0; i < headerRow.length; i++) {
    if (normHeader_(headerRow[i]) === target) return i;
  }
  return -1;
}
function idxOfAnyHeader_(headerRow, candidates) {
  for (const c of candidates) {
    const idx = idxOfHeader_(headerRow, c);
    if (idx !== -1) return idx;
  }
  return -1;
}
function asNumber_(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  const s = String(v).replace(/[^\d,-.]/g, "").replace(",", ".");
  const n = Number(s);
  return isNaN(n) ? 0 : n;
}
function asText_(v) {
  return (v === null || v === undefined) ? "" : String(v);
}
function toIso_(d) {
  if (!d) return "";
  try {
    if (Object.prototype.toString.call(d) === "[object Date]" && !isNaN(d.getTime())) return d.toISOString();
    const dt = new Date(d);
    if (!isNaN(dt)) return dt.toISOString();
    return "";
  } catch (e) {
    return "";
  }
}
function parseDate_(x) {
  if (!x) return null;
  const d = new Date(x);
  return isNaN(d) ? null : d;
}
function clampToDayStart_(d) {
  if (!d) return null;
  const x = new Date(d);
  x.setHours(0, 0, 0, 0);
  return x;
}
function clampToDayEnd_(d) {
  if (!d) return null;
  const x = new Date(d);
  x.setHours(23, 59, 59, 999);
  return x;
}
function monthDiffInclusive_(a, b) {
  if (!a || !b) return 1;
  const y1 = a.getFullYear(), m1 = a.getMonth();
  const y2 = b.getFullYear(), m2 = b.getMonth();
  return Math.max(1, (y2 - y1) * 12 + (m2 - m1) + 1);
}

// ✅ format persen: 0.6 -> 60%, 60 -> 60%, "60%" keep
function formatPercent_(v) {
  if (v === null || v === undefined || v === "") return "";
  if (typeof v === "string") {
    const s = v.trim();
    if (!s) return "";
    if (s.includes("%")) return s;
    const n = Number(s.replace(",", "."));
    if (!isNaN(n)) {
      if (n > 0 && n <= 1) return String(Math.round(n * 100)) + "%";
      return String(Math.round(n)) + "%";
    }
    return s;
  }
  if (typeof v === "number") {
    if (v > 0 && v <= 1) return String(Math.round(v * 100)) + "%";
    return String(Math.round(v)) + "%";
  }
  return asText_(v);
}

function asBool_(v) {
  if (v === true) return true;
  if (v === false) return false;
  const s = String(v || "").trim().toLowerCase();
  if (!s) return false;
  return (s === "true" || s === "1" || s === "yes" || s === "ya");
}

function normalizeRole_(role) {
  const r = String(role || "").trim().toLowerCase();
  if (!r) return "";
  if (r.includes("manage")) return "management";
  if (r.includes("leader")) return "leader";
  if (r.includes("staff")) return "staff";
  return r; // fallback
}

function emailValid_(emailInput) {
  const email = String(emailInput || "").trim().toLowerCase();
  const emailRe = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!email || !emailRe.test(email)) return "";
  return email;
}

/***** ✅ AKUN DB (sumber kebenaran user) *****/
/**
 * Cache strategy:
 * - user by email: script cache 5 menit
 * - list akun: script cache 3 menit
 */
const CACHE_USER_TTL = 300;     // 5 min
const CACHE_AKUNLIST_TTL = 180; // 3 min

function getUserByEmail_(emailInput) {
  const email = emailValid_(emailInput);
  if (!email) return null;

  const cache = CacheService.getScriptCache();
  const cKey = "akun_user_" + email;
  const cHit = cache.get(cKey);
  if (cHit) {
    try { return JSON.parse(cHit); } catch (e) {}
  }

  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(SHEET_AKUN);
  if (!sh) throw new Error(`Sheet "${SHEET_AKUN}" tidak ditemukan.`);

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return null;

  const header = values[0];

  const iNama = idxOfAnyHeader_(header, ["nama", "Nama"]);
  const iEmail = idxOfAnyHeader_(header, ["email", "Email"]);
  const iPosisi = idxOfAnyHeader_(header, ["posisi", "Posisi"]);
  const iDivisi = idxOfAnyHeader_(header, ["divisi", "Divisi"]);
  const iStatusKaryawan = idxOfAnyHeader_(header, ["status_karyawan", "status karyawan", "Status Karyawan"]);
  const iRoleAkun = idxOfAnyHeader_(header, ["role_akun", "role akun", "Role Akun"]);
  const iStatusAplikasi = idxOfAnyHeader_(header, ["status_aplikasi", "status aplikasi", "Status Aplikasi"]);

  if (iEmail === -1) throw new Error(`Header "email" tidak ditemukan di ${SHEET_AKUN}.`);

  const data = values.slice(1);
  for (const r of data) {
    const rowEmail = String(r[iEmail] || "").trim().toLowerCase();
    if (rowEmail !== email) continue;

    const user = {
      nama: (iNama !== -1) ? asText_(r[iNama]).trim() : "",
      email: rowEmail,
      posisi: (iPosisi !== -1) ? asText_(r[iPosisi]).trim() : "",
      divisi: (iDivisi !== -1) ? asText_(r[iDivisi]).trim() : "",
      status_karyawan: (iStatusKaryawan !== -1) ? asText_(r[iStatusKaryawan]).trim() : "",
      role_akun: (iRoleAkun !== -1) ? asText_(r[iRoleAkun]).trim() : "",
      status_aplikasi: (iStatusAplikasi !== -1) ? asBool_(r[iStatusAplikasi]) : false,
    };

    try { cache.put(cKey, JSON.stringify(user), CACHE_USER_TTL); } catch (e) {}
    return user;
  }

  return null;
}

/**
 * Ambil semua akun aktif dari DB_Akun (buat pilihan impersonate)
 * - includeNonActive=false => hanya status_aplikasi TRUE
 */
function getAllUsers_(includeNonActive) {
  const cache = CacheService.getScriptCache();
  const cKey = "akun_list_" + (includeNonActive ? "all" : "active");
  const cHit = cache.get(cKey);
  if (cHit) {
    try { return JSON.parse(cHit); } catch (e) {}
  }

  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(SHEET_AKUN);
  if (!sh) throw new Error(`Sheet "${SHEET_AKUN}" tidak ditemukan.`);

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const header = values[0];

  const iNama = idxOfAnyHeader_(header, ["nama", "Nama"]);
  const iEmail = idxOfAnyHeader_(header, ["email", "Email"]);
  const iPosisi = idxOfAnyHeader_(header, ["posisi", "Posisi"]);
  const iDivisi = idxOfAnyHeader_(header, ["divisi", "Divisi"]);
  const iStatusKaryawan = idxOfAnyHeader_(header, ["status_karyawan", "status karyawan", "Status Karyawan"]);
  const iRoleAkun = idxOfAnyHeader_(header, ["role_akun", "role akun", "Role Akun"]);
  const iStatusAplikasi = idxOfAnyHeader_(header, ["status_aplikasi", "status aplikasi", "Status Aplikasi"]);

  if (iEmail === -1) throw new Error(`Header "email" tidak ditemukan di ${SHEET_AKUN}.`);

  const out = [];
  const data = values.slice(1);
  for (const r of data) {
    const email = String(r[iEmail] || "").trim().toLowerCase();
    if (!email) continue;

    const u = {
      nama: (iNama !== -1) ? asText_(r[iNama]).trim() : "",
      email,
      posisi: (iPosisi !== -1) ? asText_(r[iPosisi]).trim() : "",
      divisi: (iDivisi !== -1) ? asText_(r[iDivisi]).trim() : "",
      status_karyawan: (iStatusKaryawan !== -1) ? asText_(r[iStatusKaryawan]).trim() : "",
      role_akun: (iRoleAkun !== -1) ? asText_(r[iRoleAkun]).trim() : "",
      status_aplikasi: (iStatusAplikasi !== -1) ? asBool_(r[iStatusAplikasi]) : false,
    };

    if (!includeNonActive && !u.status_aplikasi) continue;
    out.push(u);
  }

  try { cache.put(cKey, JSON.stringify(out), CACHE_AKUNLIST_TTL); } catch (e) {}
  return out;
}

// ✅ dipakai dashboard / halaman lain untuk ambil profile
function getUserByEmail(emailInput) {
  const out = { ok: false, message: "", email: "", user: null };
  const email = emailValid_(emailInput);
  out.email = email || String(emailInput || "").trim().toLowerCase();

  if (!email) {
    out.message = "Email tidak valid.";
    return JSON.stringify(out);
  }

  try {
    const user = getUserByEmail_(email);
    if (!user) {
      out.message = "Akses ditolak: email ini tidak terdaftar di DB_Akun.";
      return JSON.stringify(out);
    }
    if (!user.status_aplikasi) {
      out.message = "Akses ditolak: akun kamu non-aktif (status_aplikasi = FALSE).";
      return JSON.stringify(out);
    }
    out.ok = true;
    out.user = user;
    return JSON.stringify(out);
  } catch (e) {
    out.message = "Error server: " + (e && e.message ? e.message : e);
    return JSON.stringify(out);
  }
}

/***** SECURITY: cek email terdaftar (mencegah spam OTP) *****/
function isEmailAllowed_(email) {
  const e = emailValid_(email);
  if (!e) return false;

  try {
    const user = getUserByEmail_(e);
    if (!user) return false;
    return !!user.status_aplikasi;
  } catch (err) {
    return false;
  }
}

/***** OTP + REMEMBER TOKEN *****/
function requestOtp(emailInput) {
  const out = { ok: false, message: "", cooldownSec: 0 };

  const email = emailValid_(emailInput);
  if (!email) {
    out.message = "Email tidak valid.";
    return JSON.stringify(out);
  }

  try {
    if (!isEmailAllowed_(email)) {
      out.message = "Akses ditolak: email ini tidak terdaftar / akun non-aktif.";
      return JSON.stringify(out);
    }
  } catch (e) {
    out.message = "Error server: " + (e && e.message ? e.message : e);
    return JSON.stringify(out);
  }

  const cache = CacheService.getScriptCache();
  const cooldownKey = "otp_cd_" + email;
  const cd = cache.get(cooldownKey);
  if (cd) {
    out.message = "Tunggu sebentar sebelum kirim OTP lagi.";
    out.cooldownSec = Number(cd) || OTP_COOLDOWN_SEC;
    return JSON.stringify(out);
  }

  const otp = String(Math.floor(100000 + Math.random() * 900000));
  const otpKey = "otp_" + email;

  cache.put(otpKey, otp, OTP_TTL_SEC);
  cache.put(cooldownKey, String(OTP_COOLDOWN_SEC), OTP_COOLDOWN_SEC);

  try {
    const subject = 'Kode OTP Masuk Laman "Cek Komisi Divisi Maintenance Jendela360"';
    const body =
      'Halo,\n\n' +
      'Kamu mengakses laman masuk "Cek Komisi Divisi Maintenance Jendela360".\n\n' +
      'Silahkan masuk menggunakan Kode OTP kamu: ' + otp + '\n\n' +
      'Berlaku selama 5 menit.\n\n' +
      'Jangan bagikan kode ini ke siapa pun.\n\n' +
      'Terima Kasih - Admin Cek Komisi Divisi Maintenance Jendela360\n\n' +
      '© Internal Tools by HCGA Jendela360 (IIA)';

    const logoBottom = "https://res.cloudinary.com/dkps3vy8m/image/upload/v1770658191/5_totgo8.png";

    const htmlBody = `
      <div style="font-family:Arial,Helvetica,sans-serif; font-size:14px; color:#0b1220; line-height:1.6;">
        <div>Halo,</div>
        <br/>
        <div>Kamu mengakses laman masuk <b>"Cek Komisi Divisi Maintenance Jendela360"</b>.</div>
        <br/>
        <div>Silahkan masuk menggunakan Kode OTP kamu:</div>
        <div style="margin:10px 0 14px 0;">
          <span style="display:inline-block; font-weight:900; font-size:22px; letter-spacing:2px;">
            ${otp}
          </span>
        </div>
        <div>Berlaku selama 5 menit.</div>
        <br/>
        <div><b>Jangan bagikan kode ini ke siapa pun.</b></div>
        <br/>
        <div>Terima Kasih - Admin Cek Komisi Divisi Maintenance Jendela360</div>
        <div style="margin-top:10px; font-size:12px; color:#64748b;">
          © Internal Tools by HCGA Jendela360 (IIA)
        </div>
        <div style="margin-top:10px; text-align:center;">
          <img src="${logoBottom}" alt="Jendela360" style="height:64px; width:auto; display:inline-block;" />
        </div>
      </div>
    `;

    MailApp.sendEmail({ to: email, subject, body, htmlBody });

  } catch (e) {
    out.message = "Gagal mengirim OTP. Pastikan email bisa menerima pesan.";
    return JSON.stringify(out);
  }

  out.ok = true;
  out.message = "OTP telah dikirim ke email kamu.";
  out.cooldownSec = OTP_COOLDOWN_SEC;
  return JSON.stringify(out);
}

function verifyOtp(emailInput, otpInput, rememberMe) {
  const out = { ok: false, message: "", email: "", token: "" };

  const email = emailValid_(emailInput);
  const otp = String(otpInput || "").trim();

  if (!email) {
    out.message = "Email tidak valid.";
    return JSON.stringify(out);
  }
  if (!otp || otp.length !== 6) {
    out.message = "Kode OTP tidak valid.";
    return JSON.stringify(out);
  }

  try {
    if (!isEmailAllowed_(email)) {
      out.message = "Akses ditolak: email ini tidak terdaftar / akun non-aktif.";
      return JSON.stringify(out);
    }
  } catch (e) {
    out.message = "Error server: " + (e && e.message ? e.message : e);
    return JSON.stringify(out);
  }

  const cache = CacheService.getScriptCache();
  const otpKey = "otp_" + email;
  const saved = cache.get(otpKey);

  if (!saved) {
    out.message = "OTP sudah kadaluarsa. Silakan kirim ulang.";
    return JSON.stringify(out);
  }
  if (saved !== otp) {
    out.message = "OTP salah. Coba lagi.";
    return JSON.stringify(out);
  }

  cache.remove(otpKey);

  out.ok = true;
  out.email = email;

  const remember = String(rememberMe) === "true" || rememberMe === true;
  if (remember) {
    const token = Utilities.getUuid().replace(/-/g, "") + Utilities.getUuid().replace(/-/g, "");
    out.token = token;
    saveRememberToken_(token, email);
  }

  return JSON.stringify(out);
}

function saveRememberToken_(token, email) {
  const props = PropertiesService.getScriptProperties();
  const key = "rt_" + token;
  const exp = Date.now() + (REMEMBER_TTL_DAYS * 24 * 60 * 60 * 1000);
  props.setProperty(key, JSON.stringify({ email: String(email || "").toLowerCase(), exp: Number(exp) }));

  try {
    CacheService.getScriptCache().put("rt_ok_" + token, String(email || "").toLowerCase(), 300);
  } catch (e) {}
}

function validateRememberToken(tokenInput) {
  const out = { ok: false, email: "" };
  const token = String(tokenInput || "").trim();
  if (!token) return JSON.stringify(out);

  try {
    const c = CacheService.getScriptCache().get("rt_ok_" + token);
    if (c) {
      if (isEmailAllowed_(c)) {
        out.ok = true;
        out.email = String(c).toLowerCase();
        return JSON.stringify(out);
      }
    }
  } catch (e) {}

  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty("rt_" + token);
  if (!raw) return JSON.stringify(out);

  try {
    const obj = JSON.parse(raw);
    if (!obj || !obj.email || !obj.exp) return JSON.stringify(out);
    if (Date.now() > Number(obj.exp)) {
      props.deleteProperty("rt_" + token);
      return JSON.stringify(out);
    }
    if (!isEmailAllowed_(obj.email)) {
      props.deleteProperty("rt_" + token);
      return JSON.stringify(out);
    }
    out.ok = true;
    out.email = String(obj.email).toLowerCase();

    try {
      CacheService.getScriptCache().put("rt_ok_" + token, out.email, 300);
    } catch (e) {}

    return JSON.stringify(out);
  } catch (e) {
    props.deleteProperty("rt_" + token);
    return JSON.stringify(out);
  }
}

function cleanupRememberTokens_(limit) {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  const keys = Object.keys(all).filter(k => k.startsWith("rt_"));
  let n = 0;
  for (const k of keys) {
    if (limit && n >= limit) break;
    try {
      const obj = JSON.parse(all[k] || "{}");
      if (!obj || !obj.exp || Date.now() > Number(obj.exp)) {
        props.deleteProperty(k);
        n++;
      }
    } catch (e) {
      props.deleteProperty(k);
      n++;
    }
  }
  return n;
}

/***** ===========================
 *  ✅ IMPERSONATION
 * ============================ */

// ✅ helper izin impersonate
function canImpersonate_(authUser, effectiveUser) {
  if (!authUser || !effectiveUser) return false;

  const authRole = normalizeRole_(authUser.role_akun);
  const effRole = normalizeRole_(effectiveUser.role_akun);

  if (String(authUser.email || "").toLowerCase() === String(effectiveUser.email || "").toLowerCase()) return true;

  if (authRole === "management") {
    return effRole === "staff" || effRole === "leader";
  }

  if (authRole === "leader") {
    if (effRole !== "staff") return false;
    const aDiv = String(authUser.divisi || "").trim().toLowerCase();
    const eDiv = String(effectiveUser.divisi || "").trim().toLowerCase();
    return !!aDiv && aDiv === eDiv;
  }

  return false;
}

/**
 * ✅ INI YANG DIPAKE DASHBOARD.HTML
 * Expected return:
 *  { ok:true, items:[{email,nama,divisi,role_akun,posisi}] }
 *
 * Rules:
 * - Staff: items []
 * - Leader: hanya staff divisi auth
 * - Management:
 *    - mode="staff": semua staff (opsional filter divisi)
 *    - mode="leader": semua leader (opsional filter divisi)
 */
function getImpersonationCandidates(authEmailInput, modeInput, divisiInput) {
  const out = { ok: false, message: "", items: [] };

  const authEmail = emailValid_(authEmailInput);
  const mode = String(modeInput || "").trim().toLowerCase(); // staff | leader
  const divFilter = String(divisiInput || "").trim().toLowerCase(); // boleh kosong

  if (!authEmail) {
    out.message = "Email auth tidak valid.";
    return JSON.stringify(out);
  }

  try {
    if (!isEmailAllowed_(authEmail)) {
      out.message = "Akses ditolak: akun login tidak terdaftar / non-aktif.";
      return JSON.stringify(out);
    }

    const authUser = getUserByEmail_(authEmail);
    if (!authUser || !authUser.status_aplikasi) {
      out.message = "Akun login tidak valid / non-aktif.";
      return JSON.stringify(out);
    }

    const authRole = normalizeRole_(authUser.role_akun);
    const allActive = getAllUsers_(false);

    // Staff: tidak boleh impersonate (UI boleh tampil, backend kirim list kosong)
    if (authRole !== "leader" && authRole !== "management") {
      out.ok = true;
      out.items = [];
      return JSON.stringify(out);
    }

    // Leader: selalu staff divisi sendiri (mode apapun dari UI tetap kita paksa staff)
    if (authRole === "leader") {
      const myDiv = String(authUser.divisi || "").trim().toLowerCase();
      const items = allActive
        .filter(u =>
          normalizeRole_(u.role_akun) === "staff" &&
          String(u.divisi || "").trim().toLowerCase() === myDiv
        )
        .map(u => ({
          email: u.email,
          nama: u.nama,
          divisi: u.divisi,
          role_akun: u.role_akun,
          posisi: u.posisi
        }))
        .sort((a, b) => (a.nama || "").localeCompare(b.nama || "", "id-ID"));

      out.ok = true;
      out.items = items;
      return JSON.stringify(out);
    }

    // Management: staff atau leader tergantung mode
    if (authRole === "management") {
      const want = (mode === "leader") ? "leader" : "staff";

      let items = allActive
        .filter(u => normalizeRole_(u.role_akun) === want)
        .map(u => ({
          email: u.email,
          nama: u.nama,
          divisi: u.divisi,
          role_akun: u.role_akun,
          posisi: u.posisi
        }));

      if (divFilter) {
        items = items.filter(x => String(x.divisi || "").trim().toLowerCase() === divFilter);
      }

      items.sort((a, b) => (a.nama || "").localeCompare(b.nama || "", "id-ID"));

      out.ok = true;
      out.items = items;
      return JSON.stringify(out);
    }

    out.ok = true;
    out.items = [];
    return JSON.stringify(out);

  } catch (e) {
    out.message = "Error server: " + (e && e.message ? e.message : e);
    return JSON.stringify(out);
  }
}

/***** MAIN DATASET (rincianpencairankomisi.html) *****/
function getKomisiDatasetByEmail(emailInput) {
  const result = { ok: false, message: "", email: "", namaMitra: "", rows: [], filterOptions: { periode: [], pekerjaan: [] } };

  const email = emailValid_(emailInput);
  result.email = email || String(emailInput || "").trim().toLowerCase();

  if (!email) {
    result.message = "Email tidak valid. Pastikan format email benar.";
    return JSON.stringify(result);
  }

  try {
    if (!isEmailAllowed_(email)) {
      result.message = "Akses ditolak: email ini tidak terdaftar / akun non-aktif.";
      return JSON.stringify(result);
    }

    const user = getUserByEmail_(email);
    if (user && user.nama) result.namaMitra = user.nama;

    const ss = getSpreadsheet_();
    const sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) { result.message = `Sheet "${SHEET_NAME}" tidak ditemukan.`; return JSON.stringify(result); }

    const values = sh.getDataRange().getValues();
    if (values.length < 2) { result.message = "Database kosong."; return JSON.stringify(result); }

    const header = values[0];
    const data = values.slice(1);

    const iEmail = idxOfHeader_(header, "Email Mitra");
    const iNama = idxOfHeader_(header, "Nama Mitra");
    const iPosisi = idxOfHeader_(header, "Posisi");

    const iTglCetak = idxOfHeader_(header, "Tanggal Cetak Invoice");
    const iPeriode = idxOfHeader_(header, "Periode Cut Off Bonus");
    const iTglKerja = idxOfHeader_(header, "Tanggal Pengerjaan Pertama");
    const iSpk = idxOfHeader_(header, "No. SPK");
    const iKoord = idxOfHeader_(header, "Nama Koordinator");
    const iJenis = idxOfHeader_(header, "Jenis Pekerjaan");
    const iUnit = idxOfHeader_(header, "Kode Unit");
    const iApt = idxOfHeader_(header, "Nama Apartemen");
    const iBiaya = idxOfHeader_(header, "Biaya Penagihan Jasa");
    const iPersenPel = idxOfHeader_(header, "Persentase Komisi Pelaksana");
    const iKomisiPel = idxOfHeader_(header, "Komisi Pelaksana");

    const iStatus = idxOfHeader_(header, "Status Pencairan");
    const iWaktuCair = idxOfHeader_(header, "Waktu Pencairan");

    const required = [
      ["Email Mitra", iEmail], ["Nama Mitra", iNama], ["Posisi", iPosisi],
      ["Tanggal Cetak Invoice", iTglCetak], ["Periode Cut Off Bonus", iPeriode], ["Tanggal Pengerjaan Pertama", iTglKerja],
      ["No. SPK", iSpk], ["Nama Koordinator", iKoord], ["Jenis Pekerjaan", iJenis],
      ["Kode Unit", iUnit], ["Nama Apartemen", iApt],
      ["Biaya Penagihan Jasa", iBiaya], ["Persentase Komisi Pelaksana", iPersenPel], ["Komisi Pelaksana", iKomisiPel],
      ["Status Pencairan", iStatus], ["Waktu Pencairan", iWaktuCair],
    ];
    const missing = required.filter(x => x[1] === -1).map(x => x[0]);
    if (missing.length) { result.message = "Header tidak ketemu: " + missing.join(", "); return JSON.stringify(result); }

    let namaMitraFromSlip = "";
    const rows = [];
    const setPeriode = new Set();
    const setPekerjaan = new Set();

    for (const r of data) {
      const rowEmail = String(r[iEmail] || "").trim().toLowerCase();
      if (rowEmail !== email) continue;

      if (!namaMitraFromSlip) namaMitraFromSlip = asText_(r[iNama]).trim();

      const periode = asText_(r[iPeriode]).trim();
      const pekerjaan = asText_(r[iJenis]).trim();
      if (periode) setPeriode.add(periode);
      if (pekerjaan) setPekerjaan.add(pekerjaan);

      rows.push({
        tglCetak: toIso_(r[iTglCetak]),
        periode,
        tglKerja: toIso_(r[iTglKerja]),
        spk: asText_(r[iSpk]).trim(),
        koordinator: asText_(r[iKoord]).trim(),
        posisi: asText_(r[iPosisi]).trim(),
        pekerjaan,
        unit: asText_(r[iUnit]).trim(),
        apartemen: asText_(r[iApt]).trim(),
        biaya: asNumber_(r[iBiaya]),
        persenPelaksana: formatPercent_(r[iPersenPel]),
        komisiPelaksana: asNumber_(r[iKomisiPel]),
        statusPencairan: asText_(r[iStatus]).trim(),
        waktuPencairan: toIso_(r[iWaktuCair]),
      });
    }

    rows.sort((a, b) => new Date(b.tglCetak) - new Date(a.tglCetak));

    result.ok = true;
    if (!result.namaMitra) result.namaMitra = namaMitraFromSlip || "";
    result.rows = rows;
    result.filterOptions.periode = Array.from(setPeriode).sort();
    result.filterOptions.pekerjaan = Array.from(setPekerjaan).sort();
    return JSON.stringify(result);

  } catch (e) {
    result.message = "Error server: " + (e && e.message ? e.message : e);
    return JSON.stringify(result);
  }
}

/***** ✅ HISTORI PENGERJAAN DATASET (rincianpengerjaan.html) *****/
function getPengerjaanDatasetByEmail(emailInput) {
  const out = {
    ok: false,
    message: "",
    email: "",
    namaMitra: "",
    posisiMitra: "Pelaksana",
    rows: [],
    filterOptions: { pekerjaan: [] }
  };

  const email = emailValid_(emailInput);
  out.email = email || String(emailInput || "").trim().toLowerCase();

  if (!email) {
    out.message = "Email tidak valid.";
    return JSON.stringify(out);
  }

  try {
    if (!isEmailAllowed_(email)) {
      out.message = "Akses ditolak: email ini tidak terdaftar / akun non-aktif.";
      return JSON.stringify(out);
    }

    const ss = getSpreadsheet_();

    const user = getUserByEmail_(email);
    if (user) {
      out.namaMitra = user.nama || "";
      out.posisiMitra = (user.posisi && String(user.posisi).trim()) ? String(user.posisi).trim() : out.posisiMitra;
    }

    const sh = ss.getSheetByName(SHEET_PENGERJAAN);
    if (!sh) { out.message = `Sheet "${SHEET_PENGERJAAN}" tidak ditemukan.`; return JSON.stringify(out); }

    const values = sh.getDataRange().getValues();
    if (values.length < 2) { out.message = "Database pengerjaan kosong."; return JSON.stringify(out); }

    const header = values[0];
    const data = values.slice(1);

    const iJadwal = idxOfAnyHeader_(header, ["Tanggal Order", "Penjadwalan Pertama", "Penjadwalan", "Tgl Order", "Tanggal"]);
    const iSpk = idxOfAnyHeader_(header, ["No. SPK", "Nomor SPK", "SPK"]);
    const iApt = idxOfAnyHeader_(header, ["Nama Apartement", "Nama Apartemen", "Nama Apartment", "Apartemen"]);
    const iUnit = idxOfAnyHeader_(header, ["Kode Unit", "Unit"]);
    const iPekerjaan = idxOfAnyHeader_(header, ["Pekerjaan", "Jenis Pekerjaan"]);
    const iUci = idxOfAnyHeader_(header, ["Unique Code Invoice", "UniqueCodeInvoice", "UCI"]);
    const iPelaksana = idxOfAnyHeader_(header, ["Nama Pelaksana", "Pelaksana"]);
    const iPosisi = idxOfAnyHeader_(header, ["Posisi", "Role", "Jabatan"]);
    const iKoord = idxOfAnyHeader_(header, ["Nama Koordinator", "Koordinator"]);
    const iStatusPengerjaan = idxOfAnyHeader_(header, ["Status Penyelesaian Pekerjaan", "Status Pengerjaan", "Status Pekerjaan", "Status"]);
    const iCetakInvoice = idxOfAnyHeader_(header, ["Tanggal Invoice Selesai", "Tanggal Cetak Invoice", "Cetak Invoice", "Tgl Cetak Invoice"]);
    const iStatusJadiInvoice = idxOfAnyHeader_(header, ["Status Jadi Invoice", "Status Invoice", "Invoice Status"]);

    const must = [];
    if (iPelaksana === -1) must.push("Nama Pelaksana (kolom H)");
    if (iJadwal === -1) must.push("Tanggal Order (kolom A)");
    if (iSpk === -1) must.push("No. SPK");
    if (must.length) {
      out.message = "Header wajib belum ketemu di DB_List_Order: " + must.join(", ");
      return JSON.stringify(out);
    }

    const jobs = new Set();
    const rows = [];

    const loginName = (out.namaMitra || "").trim().toLowerCase();
    for (const r of data) {
      const pel = asText_(r[iPelaksana]).trim();
      if (!pel) continue;

      if (loginName && pel.trim().toLowerCase() !== loginName) continue;

      const pekerjaan = (iPekerjaan !== -1) ? asText_(r[iPekerjaan]).trim() : "";
      if (pekerjaan) jobs.add(pekerjaan);

      rows.push({
        penjadwalanPertama: (iJadwal !== -1) ? toIso_(r[iJadwal]) : "",
        spk: (iSpk !== -1) ? asText_(r[iSpk]).trim() : "",
        apartemen: (iApt !== -1) ? asText_(r[iApt]).trim() : "",
        unit: (iUnit !== -1) ? asText_(r[iUnit]).trim() : "",
        pekerjaan: pekerjaan,
        uniqueCodeInvoice: (iUci !== -1) ? asText_(r[iUci]).trim() : "",
        pelaksana: pel,
        posisi: (iPosisi !== -1) ? asText_(r[iPosisi]).trim() : "",
        koordinator: (iKoord !== -1) ? asText_(r[iKoord]).trim() : "",
        statusPengerjaan: (iStatusPengerjaan !== -1) ? asText_(r[iStatusPengerjaan]).trim() : "",
        cetakInvoice: (iCetakInvoice !== -1) ? toIso_(r[iCetakInvoice]) : "",
        statusJadiInvoice: (iStatusJadiInvoice !== -1) ? asText_(r[iStatusJadiInvoice]).trim() : "",
      });
    }

    out.ok = true;
    out.rows = rows;

    if (!out.namaMitra) out.namaMitra = rows.length ? (rows[0].pelaksana || "") : "";
    out.filterOptions.pekerjaan = Array.from(jobs).sort();

    return JSON.stringify(out);

  } catch (e) {
    out.message = "Error server: " + (e && e.message ? e.message : e);
    return JSON.stringify(out);
  }
}

/***** ✅ DETAIL KOMISI BY UNIQUE CODE (untuk popup CEK) *****/
function getKomisiDetailByUniqueCode(emailInput, uniqueCodeInvoiceInput) {
  const out = { ok: false, message: "", email: "", rows: [] };

  const email = emailValid_(emailInput);
  out.email = email || String(emailInput || "").trim().toLowerCase();

  const uci = String(uniqueCodeInvoiceInput || "").trim();
  if (!uci) { out.message = "Unique Code Invoice kosong."; return JSON.stringify(out); }
  if (!email) { out.message = "Email tidak valid."; return JSON.stringify(out); }

  try {
    if (!isEmailAllowed_(email)) {
      out.message = "Akses ditolak: email ini tidak terdaftar / akun non-aktif.";
      return JSON.stringify(out);
    }

    const ss = getSpreadsheet_();
    const sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) { out.message = `Sheet "${SHEET_NAME}" tidak ditemukan.`; return JSON.stringify(out); }

    const values = sh.getDataRange().getValues();
    if (values.length < 2) { out.message = "Database kosong."; return JSON.stringify(out); }

    const header = values[0];

    const iEmail = idxOfAnyHeader_(header, ["Email Mitra", "Email"]);
    const iPosisi = idxOfAnyHeader_(header, ["Posisi", "Role", "Jabatan"]);
    const iTglCetak = idxOfAnyHeader_(header, ["Tanggal Cetak Invoice", "Tgl Cetak Invoice", "Cetak Invoice"]);
    const iPeriode = idxOfAnyHeader_(header, ["Periode Cut Off Bonus", "Periode Pencairan", "Periode"]);
    const iTglKerja = idxOfAnyHeader_(header, ["Tanggal Pengerjaan Pertama", "Penjadwalan Pertama", "Tanggal Kerja"]);
    const iSpk = idxOfAnyHeader_(header, ["No. SPK", "SPK"]);
    const iKoord = idxOfAnyHeader_(header, ["Nama Koordinator", "Koordinator"]);
    const iJenis = idxOfAnyHeader_(header, ["Jenis Pekerjaan", "Pekerjaan"]);
    const iUnit = idxOfAnyHeader_(header, ["Kode Unit", "Unit"]);
    const iApt = idxOfAnyHeader_(header, ["Nama Apartemen", "Nama Apartement", "Apartemen"]);
    const iBiaya = idxOfAnyHeader_(header, ["Biaya Penagihan Jasa", "Biaya Jasa"]);
    const iPersen = idxOfAnyHeader_(header, ["Persentase Komisi Pelaksana", "% Komisi", "Persen Komisi"]);
    const iKomisi = idxOfAnyHeader_(header, ["Komisi Pelaksana", "Komisi"]);
    const iStatus = idxOfAnyHeader_(header, ["Status Pencairan", "Status Cair"]);
    const iWaktu = idxOfAnyHeader_(header, ["Waktu Pencairan", "Waktu Cair"]);
    const iUci = idxOfAnyHeader_(header, ["Unique Code Invoice", "UniqueCodeInvoice", "UCI"]);

    const must = [];
    if (iUci === -1) must.push("Unique Code Invoice");
    if (iEmail === -1) must.push("Email Mitra");
    if (must.length) { out.message = "Header wajib belum ketemu: " + must.join(", "); return JSON.stringify(out); }

    const lastRow = sh.getLastRow();
    const uciRange = sh.getRange(2, iUci + 1, Math.max(0, lastRow - 1), 1);
    const finder = uciRange.createTextFinder(uci).matchEntireCell(true);
    const cell = finder.findNext();
    if (!cell) {
      out.ok = true;
      out.rows = [];
      return JSON.stringify(out);
    }

    const rowIndex = cell.getRow();
    const row = sh.getRange(rowIndex, 1, 1, sh.getLastColumn()).getValues()[0];

    const rowEmail = String(row[iEmail] || "").trim().toLowerCase();
    if (rowEmail !== email) {
      out.message = "Akses ditolak: data komisi ini bukan milik akun kamu.";
      return JSON.stringify(out);
    }

    out.ok = true;
    out.rows = [{
      tglCetak: (iTglCetak !== -1) ? toIso_(row[iTglCetak]) : "",
      periode: (iPeriode !== -1) ? asText_(row[iPeriode]).trim() : "",
      tglKerja: (iTglKerja !== -1) ? toIso_(row[iTglKerja]) : "",
      spk: (iSpk !== -1) ? asText_(row[iSpk]).trim() : "",
      koordinator: (iKoord !== -1) ? asText_(row[iKoord]).trim() : "",
      posisi: (iPosisi !== -1) ? asText_(row[iPosisi]).trim() : "",
      pekerjaan: (iJenis !== -1) ? asText_(row[iJenis]).trim() : "",
      unit: (iUnit !== -1) ? asText_(row[iUnit]).trim() : "",
      apartemen: (iApt !== -1) ? asText_(row[iApt]).trim() : "",
      biaya: (iBiaya !== -1) ? asNumber_(row[iBiaya]) : 0,
      persenPelaksana: (iPersen !== -1) ? formatPercent_(row[iPersen]) : "",
      komisiPelaksana: (iKomisi !== -1) ? asNumber_(row[iKomisi]) : 0,
      statusPencairan: (iStatus !== -1) ? asText_(row[iStatus]).trim() : "",
      waktuPencairan: (iWaktu !== -1) ? toIso_(row[iWaktu]) : "",
    }];

    return JSON.stringify(out);

  } catch (e) {
    out.message = "Error server: " + (e && e.message ? e.message : e);
    return JSON.stringify(out);
  }
}

/***** HOMEPAGE SUMMARY (dashboard.html) *****/
function getHomepageSummaryByEmail(emailInput) {
  return getHomepageSummaryByEmailWithFilter(emailInput, "", "");
}

/***** ✅ HOMEPAGE SUMMARY + FILTER + RANKING + TARGET *****/
function getHomepageSummaryByEmailWithFilter(emailInput, dateFromIso, dateToIso) {
  const out = {
    ok: false,
    message: "",
    email: String(emailInput || "").trim().toLowerCase(),

    // profile
    namaMitra: "",
    posisi: "",

    // ✅ tambahan profile dari DB_Akun
    divisi: "",
    status_karyawan: "",
    role_akun: "",

    // ringkasan pengerjaan (DB_List_Order)
    workAll: 0,
    workDone: 0,
    workNotDone: 0,
    workDoneInvoiced: 0,
    workDoneNotInvoiced: 0,

    // ringkasan komisi (DB_Slip_Komisi_Cair)
    totalAll: 0,
    totalPaid: 0,
    totalUnpaid: 0,
    spkAll: 0,
    spkPaid: 0,
    spkUnpaid: 0,

    // pendapatan bersih
    netTotal: 0,
    netIncome: 0,
    netDeduction: 0,

    // target
    target: { achieved: 0, target: 0, percent: 0 },

    // ranking
    rankings: {
      acrossDivisions: [],
      withinDivision: []
    }
  };

  const email = emailValid_(out.email);
  out.email = email || out.email;

  if (!email) {
    out.message = "Email tidak valid.";
    return JSON.stringify(out);
  }

  try {
    if (!isEmailAllowed_(email)) {
      out.message = "Akses ditolak: email ini tidak terdaftar / akun non-aktif.";
      return JSON.stringify(out);
    }

    const ss = getSpreadsheet_();

    // ✅ ambil profile dari DB_Akun
    const user = getUserByEmail_(email);
    if (user) {
      out.namaMitra = user.nama || "";
      out.posisi = user.posisi || "";
      out.divisi = user.divisi || "";
      out.status_karyawan = user.status_karyawan || "";
      out.role_akun = user.role_akun || "";
    }

    const fromD = clampToDayStart_(parseDate_(dateFromIso || ""));
    const toD = clampToDayEnd_(parseDate_(dateToIso || ""));
    const hasFilter = !!(fromD || toD);

    const inRange = (dt) => {
      if (!dt) return false;
      if (fromD && dt < fromD) return false;
      if (toD && dt > toD) return false;
      return true;
    };

    // =========================
    // 1) Komisi + ranking dari DB_Slip_Komisi_Cair
    // =========================
    const shSlip = ss.getSheetByName(SHEET_NAME);
    if (!shSlip) {
      out.message = `Sheet "${SHEET_NAME}" tidak ditemukan.`;
      return JSON.stringify(out);
    }

    const vSlip = shSlip.getDataRange().getValues();
    if (vSlip.length < 2) {
      out.message = "Database komisi kosong.";
      return JSON.stringify(out);
    }

    const hSlip = vSlip[0];
    const dSlip = vSlip.slice(1);

    const iEmail = idxOfAnyHeader_(hSlip, ["Email Mitra", "Email"]);
    const iNama = idxOfAnyHeader_(hSlip, ["Nama Mitra", "Nama"]);
    const iPos = idxOfAnyHeader_(hSlip, ["Posisi", "Role", "Jabatan"]);
    const iTglCetak = idxOfAnyHeader_(hSlip, ["Tanggal Cetak Invoice", "Tgl Cetak Invoice", "Cetak Invoice"]);
    const iKomisi = idxOfAnyHeader_(hSlip, ["Komisi Pelaksana", "Komisi"]);
    const iStatusCair = idxOfAnyHeader_(hSlip, ["Status Pencairan", "Status Cair"]);

    for (const r of dSlip) {
      const rowEmail = String(r[iEmail] || "").trim().toLowerCase();
      if (rowEmail !== email) continue;

      if (!out.namaMitra) out.namaMitra = (iNama !== -1) ? asText_(r[iNama]).trim() : "";
      if (!out.posisi) out.posisi = (iPos !== -1) ? asText_(r[iPos]).trim() : "";

      const dtCetak = (iTglCetak !== -1) ? parseDate_(r[iTglCetak]) : null;
      if (hasFilter && !inRange(dtCetak)) continue;

      out.spkAll += 1;
      const kom = (iKomisi !== -1) ? asNumber_(r[iKomisi]) : 0;
      out.totalAll += kom;

      const status = (iStatusCair !== -1) ? asText_(r[iStatusCair]).trim().toLowerCase() : "";
      const isPaid = status.includes("sudah") && status.includes("cair");
      if (isPaid) {
        out.spkPaid += 1;
        out.totalPaid += kom;
      } else {
        out.spkUnpaid += 1;
        out.totalUnpaid += kom;
      }
    }

    const acrossAll = new Map(); // key: name||posisi
    for (const r of dSlip) {
      const dtCetak = (iTglCetak !== -1) ? parseDate_(r[iTglCetak]) : null;
      if (hasFilter && !inRange(dtCetak)) continue;

      const name = (iNama !== -1) ? asText_(r[iNama]).trim() : "";
      const pos = (iPos !== -1) ? asText_(r[iPos]).trim() : "";
      if (!name) continue;

      const key = `${name}||${pos}`;
      const cur = acrossAll.get(key) || { name, posisi: pos || "-", total: 0, spkCount: 0 };
      cur.total += (iKomisi !== -1) ? asNumber_(r[iKomisi]) : 0;
      cur.spkCount += 1;
      acrossAll.set(key, cur);
    }

    const meName = (out.namaMitra || "").trim().toLowerCase();

    out.rankings.acrossDivisions = Array.from(acrossAll.values())
      .sort((a, b) => (b.total - a.total) || (b.spkCount - a.spkCount))
      .slice(0, 10)
      .map(x => ({
        name: x.name,
        posisi: x.posisi,
        total: x.total,
        spkCount: x.spkCount,
        isMe: (meName && x.name.trim().toLowerCase() === meName)
      }));

    out.rankings.withinDivision = Array.from(acrossAll.values())
      .filter(x => out.posisi ? (String(x.posisi || "").trim().toLowerCase() === String(out.posisi || "").trim().toLowerCase()) : true)
      .sort((a, b) => (b.total - a.total) || (b.spkCount - a.spkCount))
      .slice(0, 3)
      .map(x => ({
        name: x.name,
        posisi: x.posisi,
        total: x.total,
        spkCount: x.spkCount,
        isMe: (meName && x.name.trim().toLowerCase() === meName)
      }));

    out.target.achieved = out.spkAll;

    out.netTotal = out.totalAll;
    out.netIncome = out.totalAll;
    out.netDeduction = 0;

    // =========================
    // 2) Ringkasan pengerjaan dari DB_List_Order
    // =========================
    const shWork = ss.getSheetByName(SHEET_PENGERJAAN);
    if (shWork) {
      const vW = shWork.getDataRange().getValues();
      if (vW.length >= 2) {
        const hW = vW[0];
        const dW = vW.slice(1);

        const iPel = idxOfAnyHeader_(hW, ["Nama Pelaksana", "Pelaksana"]);
        const iJadwal = idxOfAnyHeader_(hW, ["Tanggal Order", "Penjadwalan Pertama", "Penjadwalan", "Tgl Order", "Tanggal"]);
        const iStatus = idxOfAnyHeader_(hW, ["Status Penyelesaian Pekerjaan", "Status Pengerjaan", "Status Pekerjaan", "Status"]);
        const iInv = idxOfAnyHeader_(hW, ["Status Jadi Invoice", "Status Invoice", "Invoice Status"]);

        if (iPel !== -1 && iJadwal !== -1) {
          const loginName = (out.namaMitra || "").trim().toLowerCase();

          const isDoneStatus = (s) => String(s || "").trim().toLowerCase() === "pekerjaan selesai";
          const isInvoiced = (s) => {
            const t = String(s || "").trim().toLowerCase();
            return t.includes("sudah") && t.includes("invoice");
          };

          for (const r of dW) {
            const pel = String(r[iPel] || "").trim().toLowerCase();
            if (!pel) continue;
            if (loginName && pel !== loginName) continue;

            const dt = parseDate_(r[iJadwal]);
            if (hasFilter && !inRange(dt)) continue;

            out.workAll += 1;

            const st = (iStatus !== -1) ? asText_(r[iStatus]).trim() : "";
            const inv = (iInv !== -1) ? asText_(r[iInv]).trim() : "";

            if (isDoneStatus(st)) {
              out.workDone += 1;
              if (isInvoiced(inv)) out.workDoneInvoiced += 1;
              else out.workDoneNotInvoiced += 1;
            } else {
              out.workNotDone += 1;
            }
          }
        }
      }
    }

    out.ok = true;
    out.message = "";
    return JSON.stringify(out);

  } catch (e) {
    out.ok = false;
    out.message = "Error server: " + (e && e.message ? e.message : e);
    return JSON.stringify(out);
  }
}

/***** ✅✅✅ FIX UTAMA UNTUK DASHBOARD (AUTH + EFFECTIVE) ✅✅✅
 * Dashboard.html kamu manggil:
 *   getHomepageSummaryByAuthEmailWithFilter(authEmail, effectiveEmail, fromIso, toIso)
 *
 * RULE PENTING:
 * - Staff: effectiveEmail DIPAKSA = authEmail (biar "ingat saya" tidak kebawa impersonate dari cache FE)
 * - Leader/Management: boleh effective beda, tapi wajib lolos canImpersonate_()
 */
function getHomepageSummaryByAuthEmailWithFilter(authEmailInput, effectiveEmailInput, dateFromIso, dateToIso) {
  const authEmail = emailValid_(authEmailInput);
  const effEmailCandidate = emailValid_(effectiveEmailInput);

  if (!authEmail) {
    return JSON.stringify({ ok: false, message: "Email auth tidak valid." });
  }

  try {
    // pastikan auth terdaftar & aktif
    if (!isEmailAllowed_(authEmail)) {
      return JSON.stringify({ ok: false, message: "Akses ditolak: akun login tidak terdaftar / non-aktif." });
    }

    const authUser = getUserByEmail_(authEmail);
    if (!authUser || !authUser.status_aplikasi) {
      return JSON.stringify({ ok: false, message: "Akun login tidak valid / non-aktif." });
    }

    const authRole = normalizeRole_(authUser.role_akun);

    // ✅ Staff: apapun yang FE kirim, kita paksa pakai auth (prevent kebawa effectiveEmail dari localStorage)
    if (authRole !== "leader" && authRole !== "management") {
      return getHomepageSummaryByEmailWithFilter(authEmail, dateFromIso || "", dateToIso || "");
    }

    // Leader/Management: effective default = auth (kalau kosong)
    const effectiveEmail = effEmailCandidate || authEmail;

    // kalau effective beda dari auth -> cek izin impersonate
    if (effectiveEmail && effectiveEmail !== authEmail) {
      if (!isEmailAllowed_(effectiveEmail)) {
        return JSON.stringify({ ok: false, message: "Target akses tidak valid / non-aktif." });
      }

      const effUser = getUserByEmail_(effectiveEmail);
      if (!effUser || !effUser.status_aplikasi) {
        return JSON.stringify({ ok: false, message: "Target akses tidak valid / non-aktif." });
      }

      if (!canImpersonate_(authUser, effUser)) {
        return JSON.stringify({ ok: false, message: "Akses ditolak: kamu tidak punya izin untuk login sebagai akun tersebut." });
      }
    }

    // lolos -> ringkasan berdasarkan effective
    return getHomepageSummaryByEmailWithFilter(effectiveEmail, dateFromIso || "", dateToIso || "");

  } catch (e) {
    return JSON.stringify({ ok: false, message: "Error server: " + (e && e.message ? e.message : e) });
  }
}

/***** ✅ RANKING PAGE DATA (ranking.html) — biar aman kalau dibuka *****/
function getRankingPageDataByEmailWithFilter(emailInput, dateFromIso, dateToIso) {
  const res = JSON.parse(getHomepageSummaryByEmailWithFilter(emailInput, dateFromIso, dateToIso));
  if (!res || !res.ok) return JSON.stringify(res || { ok: false, message: "Gagal ambil ranking." });

  return JSON.stringify({
    ok: true,
    email: res.email,
    namaMitra: res.namaMitra || "",
    posisi: res.posisi || "",
    rankings: res.rankings || { acrossDivisions: [], withinDivision: [] }
  });
}

function TEST_openSheet() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  SpreadsheetApp.openById(id).getSheets()[0].getName();
}
