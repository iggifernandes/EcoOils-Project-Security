import React, { useState, useEffect, useCallback, useMemo, useContext, createContext } from "react";
import ECO_LOGO from "./ecoLogo.js";
import { Shield, Plus, Trash2, Printer, LogOut, Lock, ChevronLeft, ChevronRight, ChevronDown, ChevronUp, Clock, Check, X, UserPlus, Settings, ArrowLeft, MapPin, MapPinOff, FileSpreadsheet, CalendarDays, AlertTriangle, Award, Link2, Users, Camera, Download, ShieldAlert, Paperclip, Search, Eye } from "lucide-react";
// xlsx is large (~400 kB): loaded only when an Excel export is requested.
const loadXLSX = () => import("xlsx");

/* ---------- constants & helpers ---------- */

const BULAN_ID = ["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"];
const BULAN_SHORT = ["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];
const BULAN_EN = ["January","February","March","April","May","June","July","August","September","October","November","December"];
const BULAN_EN_SHORT = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const BULAN_MS = ["Januari","Februari","Mac","April","Mei","Jun","Julai","Ogos","September","Oktober","November","Disember"];
const BULAN_MS_SHORT = ["Jan","Feb","Mac","Apr","Mei","Jun","Jul","Ogo","Sep","Okt","Nov","Dis"];
const HARI_SHORT = { id: ["Min","Sen","Sel","Rab","Kam","Jum","Sab"], en: ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"], ms: ["Ahd","Isn","Sel","Rab","Kha","Jum","Sab"] };

// Picks the text for the current language (id | en | ms). Falls back to English when Malay is not given.
function pick(lang, id, en, ms) { return lang === "en" ? en : lang === "ms" ? (ms !== undefined ? ms : en) : id; }
function monthShort(m, lang) { return (lang === "en" ? BULAN_EN_SHORT : lang === "ms" ? BULAN_MS_SHORT : BULAN_SHORT)[m]; }
function monthFull(m, lang) { return (lang === "en" ? BULAN_EN : lang === "ms" ? BULAN_MS : BULAN_ID)[m]; }
function dayName(day, lang) {
  const dow = new Date(day.y, day.m, day.d).getDay();
  return HARI_SHORT[lang || "id"][dow];
}

function pad(n) { return String(n).padStart(2, "0"); }
function dateKey(y, m, d) { return `${y}-${pad(m + 1)}-${pad(d)}`; }

// Period runs 16th of a month through 15th of the next month.
// periodId = "YYYY-MM" where MM is the ENDING month (the 15th side).
function periodIdForToday() {
  const now = new Date();
  let y = now.getFullYear(), m = now.getMonth();
  if (now.getDate() < 16) {
    // we're still inside the period that ends this month
  } else {
    m += 1;
    if (m > 11) { m = 0; y += 1; }
  }
  return `${y}-${pad(m + 1)}`;
}

// lang defaults to "id" so the printed sheet (which must match the physical
// Indonesian form) can call this without passing a language.
// Used throughout the app's own screens (guard header, period dropdowns) — a plain dash reads
// cleanly in both languages. The printed sheet does NOT use this; see periodLabelPrint below.
function periodLabel(periodId, lang = "id") {
  const [yStr, mStr] = periodId.split("-");
  let y = parseInt(yStr, 10), m = parseInt(mStr, 10) - 1;
  let startM = m - 1, startY = y;
  if (startM < 0) { startM = 11; startY -= 1; }
  return `16 ${monthFull(startM, lang)} - 15 ${monthFull(m, lang)} ${y}`;
}

// Used only by the printed sheet, which must match the physical ALSOK form's exact wording
// ("s/d" = "sampai dengan"/until), regardless of the app's language toggle.
function periodLabelPrint(periodId) {
  const [yStr, mStr] = periodId.split("-");
  let y = parseInt(yStr, 10), m = parseInt(mStr, 10) - 1;
  let startM = m - 1, startY = y;
  if (startM < 0) { startM = 11; startY -= 1; }
  return `16 ${monthFull(startM, "id")} s/d 15 ${monthFull(m, "id")} ${y}`;
}

function periodDays(periodId) {
  const [yStr, mStr] = periodId.split("-");
  let y = parseInt(yStr, 10), endM = parseInt(mStr, 10) - 1;
  let startM = endM - 1, startY = y;
  if (startM < 0) { startM = 11; startY -= 1; }
  const days = [];
  const daysInStartMonth = new Date(startY, startM + 1, 0).getDate();
  for (let d = 16; d <= daysInStartMonth; d++) {
    days.push({ y: startY, m: startM, d, key: dateKey(startY, startM, d) });
  }
  for (let d = 1; d <= 15; d++) {
    days.push({ y, m: endM, d, key: dateKey(y, endM, d) });
  }
  return days;
}

function shiftPeriod(periodId, delta) {
  const [yStr, mStr] = periodId.split("-");
  let y = parseInt(yStr, 10), m = parseInt(mStr, 10) - 1;
  m += delta;
  while (m < 0) { m += 12; y -= 1; }
  while (m > 11) { m -= 12; y += 1; }
  return `${y}-${pad(m + 1)}`;
}

function recentPeriods(count = 14) {
  const cur = periodIdForToday();
  const list = [];
  for (let i = 0; i < count; i++) list.push(shiftPeriod(cur, -i));
  return list;
}

function computeHours(mulai, selesai) {
  if (!mulai || !selesai) return "";
  const [h1, m1] = mulai.split(":").map(Number);
  const [h2, m2] = selesai.split(":").map(Number);
  let mins = (h2 * 60 + m2) - (h1 * 60 + m1);
  if (mins < 0) mins += 24 * 60;
  const hrs = mins / 60;
  return Number.isInteger(hrs) ? String(hrs) : hrs.toFixed(1);
}

function uid() { return Math.random().toString(36).slice(2, 10); }

function safeSheetName(name, used) {
  let base = (name || "Guard").replace(/[\\/?*[\]:]/g, "").slice(0, 28).trim() || "Guard";
  let candidate = base;
  let i = 2;
  while (used.has(candidate)) {
    candidate = `${base} (${i})`.slice(0, 31);
    i++;
  }
  used.add(candidate);
  return candidate;
}

async function exportAllToExcel(roster, entries, schedule) {
  const XLSX = await loadXLSX();
  const wb = XLSX.utils.book_new();
  const rosterRows = roster.map((g) => ({
    Nama: g.name, NPK: g.npk || "", Jabatan: g.jabatan || "", Regu: g.regu || "", Projek: g.projek || "", Perusahaan: g.company || "", "No. HP": g.phone || "",
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rosterRows), "Roster");

  const usedNames = new Set(["Roster"]);
  roster.forEach((g) => {
    const guardEntries = entries[g.id] || {};
    const guardSchedule = schedule[g.id] || {};
    const dateKeys = Array.from(new Set([...Object.keys(guardEntries), ...Object.keys(guardSchedule)])).sort();
    const rows = dateKeys.map((dk) => {
      const e = guardEntries[dk] || {};
      return {
        Tanggal: dk,
        "Shift Terjadwal": guardSchedule[dk] || "",
        "Jam Mulai": e.jamMulai || "",
        "Jam Selesai": e.jamSelesai || "",
        "Jumlah Jam": computeHours(e.jamMulai, e.jamSelesai),
        Keterangan: e.keterangan || "",
        Lokasi: e.location?.lat != null ? `${e.location.lat}, ${e.location.lng}` : "",
        "Jarak dari Lokasi (m)": e.location?.distanceMeters ?? "",
        "Dalam Radius": e.location?.withinRadius === true ? "Ya" : e.location?.withinRadius === false ? "Tidak" : "",
      };
    });
    const sheet = XLSX.utils.json_to_sheet(rows.length ? rows : [{ Tanggal: "(belum ada data)" }]);
    XLSX.utils.book_append_sheet(wb, sheet, safeSheetName(g.name, usedNames));
  });

  XLSX.writeFile(wb, `timesheet-export-${new Date().toISOString().slice(0, 10)}.xlsx`);
}

async function exportRosterToExcel(roster) {
  const XLSX = await loadXLSX();
  const wb = XLSX.utils.book_new();
  const rows = roster.map((g) => ({
    Nama: g.name, NPK: g.npk || "", Jabatan: g.jabatan || "", Regu: g.regu || "", Projek: g.projek || "", Perusahaan: g.company || "", "No. HP": g.phone || "",
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "Roster");
  XLSX.writeFile(wb, `roster-${new Date().toISOString().slice(0, 10)}.xlsx`);
}

function captureLocation() {
  return new Promise((resolve) => {
    if (!navigator.geolocation) { resolve({ status: "unsupported" }); return; }
    navigator.geolocation.getCurrentPosition(
      (pos) => resolve({
        status: "ok",
        lat: Number(pos.coords.latitude.toFixed(6)),
        lng: Number(pos.coords.longitude.toFixed(6)),
        accuracy: Math.round(pos.coords.accuracy),
      }),
      () => resolve({ status: "denied" }),
      { enableHighAccuracy: true, timeout: 10000, maximumAge: 5 * 60 * 1000 }
    );
  });
}

// Haversine distance in meters between two lat/lng points.
function distanceMeters(lat1, lng1, lat2, lng2) {
  const R = 6371000;
  const toRad = (d) => (d * Math.PI) / 180;
  const dLat = toRad(lat2 - lat1);
  const dLng = toRad(lng2 - lng1);
  const a = Math.sin(dLat / 2) ** 2 + Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLng / 2) ** 2;
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

// Tags a captured location with distance-from-site + whether it's inside the geofence,
// if the site coordinates have been configured in Settings.
function tagGeofence(loc, settings) {
  if (!loc || !settings?.siteLat || !settings?.siteLng) return loc;
  const siteLat = parseFloat(settings.siteLat);
  const siteLng = parseFloat(settings.siteLng);
  if (Number.isNaN(siteLat) || Number.isNaN(siteLng)) return loc;
  const radius = parseFloat(settings.geofenceRadius) || 150;
  const dist = distanceMeters(siteLat, siteLng, loc.lat, loc.lng);
  return { ...loc, distanceMeters: Math.round(dist), withinRadius: dist <= radius };
}

// Shift codes from the SOS "Schedule Satpam" legend, plus HS (Half Shift — custom hours, no fixed default).
const SHIFT_CODES = {
  P: { label: "Pagi", mulai: "07:00", selesai: "15:00" },
  S: { label: "Sore", mulai: "15:00", selesai: "23:00" },
  M: { label: "Malam", mulai: "23:00", selesai: "07:00" },
  P1: { label: "Non Shift", mulai: "07:00", selesai: "15:00" },
  P2: { label: "Non Shift", mulai: "07:00", selesai: "12:00" },
  HS: { label: "Half Shift", custom: true },
  OFF: { label: "Off", mulai: "", selesai: "" },
};
const SHIFT_ORDER = ["P", "S", "M", "P1", "P2", "HS", "OFF"];

// Roster from the SOS "Schedule Satpam PT. Shield-On Service Tbk (SOS)" sheet, Lokasi: PT. ECOOILS.
// Order and spelling per Iggy's corrected list.
const SOS_ROSTER = [
  { name: "Muhammad Irsan", jabatan: "Danru", phone: "081371202710", regu: "" },
  { name: "Jhonsen Panjaitan", jabatan: "Liader shift", phone: "085835804283", regu: "" },
  { name: "Try Agung Pramana", jabatan: "Liader shift", phone: "082272505223", regu: "" },
  { name: "Rendi Irawan", jabatan: "Liader shift", phone: "082259044778", regu: "" },
  { name: "Arya Fahrul Manurung", jabatan: "Anggota", phone: "085834168556", regu: "A" },
  { name: "Agung Surya Herlambang", jabatan: "Anggota", phone: "082263162677", regu: "A" },
  { name: "Dian Akmal", jabatan: "Anggota", phone: "085327466544", regu: "B" },
  { name: "Aril Septiadi", jabatan: "Anggota", phone: "085767596552", regu: "B" },
  { name: "Muhammad Syahputra", jabatan: "Anggota", phone: "082162844725", regu: "C" },
  { name: "Pahrurrozi Nasution", jabatan: "Anggota", phone: "085933409534", regu: "C" },
  { name: "Ilham Affandi", jabatan: "Anggota", phone: "083160806631", regu: "C" },
  { name: "Dani Hardyansyah", jabatan: "Anggota", phone: "082266825787", regu: "D" },
  { name: "Binhot Eben Siahaan", jabatan: "Anggota", phone: "081339751184", regu: "D" },
  { name: "Valen Andriano", jabatan: "Anggota", phone: "085608073777", regu: "A" },
  { name: "Irvan Firmansyah", jabatan: "Anggota", phone: "082273502008", regu: "B" },
  { name: "Renaldi", jabatan: "Anggota", phone: "085636878051", regu: "D" },
  { name: "Febrianto", jabatan: "Anggota", phone: "085355738168", regu: "B" },
  { name: "Febry Ibrahim", jabatan: "Anggota", phone: "083876285486", regu: "C" },
  { name: "Aldi Ramadhan", jabatan: "Anggota", phone: "081364735611", regu: "B" },
  { name: "Deddy Satria", jabatan: "Anggota", phone: "081261938901", regu: "D" },
  { name: "Al Faroby Siregar", jabatan: "Anggota", phone: "081377120729", regu: "A" },
  { name: "Handoko Pranata", jabatan: "Anggota", phone: "082385439203", regu: "D" },
  { name: "Muhammad Siddiq", jabatan: "Anggota", phone: "085174375941", regu: "C" },
  { name: "Jona Alfian", jabatan: "Anggota", phone: "085184216456", regu: "A" },
].map((g) => ({ ...g, npk: "", projek: "EcoOils", company: "SOS" }));

/* ---------- storage layer ----------
   Shared data (roster, entries, schedule, settings, file metadata) is read
   and written through a Google Apps Script Web App, which stores it in a
   Key/Value tab of a Google Sheet and saves uploaded files as real files in
   Drive. Set VITE_SHEETS_API_URL (see .env.example) to the deployed Apps
   Script Web App URL to enable this.

   Until that's set, everything falls back to browser localStorage — fine for
   solo testing on one device, but NOT shared across guards' phones. */

export const APP_VERSION = "1.3.4";
export const APP_DATE = "2026-10-28";
// Phase 1 = guard database + documents only. Guard log-in/out, shift roster and
// timesheet printing stay in the code but are hidden from guards and Danru;
// only the Manager (PIN-protected link on the login page) can reach the timesheet side.
const PHASE_DB = true;

function addMonths(isoDate, n) {
  if (!isoDate) return "";
  const d = new Date(isoDate + "T00:00:00");
  if (isNaN(d)) return "";
  d.setMonth(d.getMonth() + n);
  return d.toISOString().slice(0, 10);
}

// Mandatory personnel documents (Indonesian Satpam rules: Perpol 4/2020 requirements,
// KTA Satpam, Gada training, plus employer-side contract and BPJS records).
// validMonths: validity counted from the issue date (SKCK = 6 months per Iggy).
const DOC_REQUIREMENTS = [
  { key: "photo", id: "Pas foto", en: "Passport photo", ms: "Gambar pasport", special: "photo" },
  { key: "ktp", id: "KTP", en: "KTP (national ID)", ms: "KTP (kad pengenalan)", number: true },
  { key: "ijazah", id: "Ijazah (min. SMA/sederajat)", en: "Diploma (min. senior high)", ms: "Diploma (min. sekolah menengah atas)" },
  { key: "skck", id: "SKCK (berlaku 6 bulan)", en: "Police clearance SKCK (valid 6 months)", ms: "Pelepasan polis SKCK (sah 6 bulan)", dated: true, validMonths: 6, number: true },
  { key: "mcu", id: "MCU / Surat sehat jasmani & rohani", en: "MCU / medical fitness (physical & mental)", ms: "MCU / kesihatan (fizikal & mental)", dated: true, validMonths: 12 },
  { key: "gada", id: "Sertifikat Gada (Pratama/Madya/Utama)", en: "Gada certificate (Pratama/Madya/Utama)", ms: "Sijil Gada (Pratama/Madya/Utama)", dated: true, number: true, hasLevel: true },
  { key: "kta", id: "KTA Satpam", en: "KTA Satpam (guard member card)", ms: "KTA Satpam (kad ahli pengawal)", dated: true, number: true },
  { key: "kontrak", id: "Kontrak kerja", en: "Employment contract", ms: "Kontrak pekerjaan" },
  { key: "bpjskes", id: "BPJS Kesehatan", en: "BPJS Kesehatan", ms: "BPJS Kesehatan", number: true },
  { key: "bpjstk", id: "BPJS Ketenagakerjaan", en: "BPJS Ketenagakerjaan", ms: "BPJS Ketenagakerjaan", number: true },
];

function docValidUntil(req, d) {
  if (!d) return "";
  if (d.expiry) return d.expiry;
  if (req.validMonths && d.issued) return addMonths(d.issued, req.validMonths);
  return "";
}

const WARN_DAYS = 90; // warn at least 3 months before any expiry
const warnDate = () => new Date(Date.now() + WARN_DAYS * 24 * 60 * 60 * 1000).toISOString().slice(0, 10);
// -> "ok" | "soon" | "missing" | "expired" | "nodate"
function docStatus(guard, req) {
  const today = new Date().toISOString().slice(0, 10);
  if (req.special === "photo") return guard.hasPhoto ? "ok" : "missing";
  const d = (guard.docs || {})[req.key];
  if (!d || !d.fileName) return "missing";
  if (req.validMonths && !d.issued) return "nodate";
  const until = docValidUntil(req, d);
  if (until && until < today) return "expired";
  if (until && until <= warnDate()) return "soon";
  return "ok";
}

function guardDocIssues(guard) {
  return DOC_REQUIREMENTS.map((r) => ({ req: r, status: docStatus(guard, r) })).filter((x) => x.status !== "ok" && x.status !== "soon");
}
function guardDocSoon(guard) {
  return DOC_REQUIREMENTS.map((r) => ({ req: r, status: docStatus(guard, r) })).filter((x) => x.status === "soon");
}

async function exportDocsChecklist(roster, lang) {
  const XLSX = await loadXLSX();
  const L = (a, b, c) => pick(lang, a, b, c);
  const head = [L("Nama", "Name", "Nama"), "NPK", L("Perusahaan", "Company", "Syarikat"), ...DOC_REQUIREMENTS.map((r) => pick(lang, r.id, r.en, r.ms)), L("Kurang", "Missing/Issues", "Kurang/Isu")];
  const word = { ok: "OK", missing: L("KURANG", "MISSING", "KURANG"), expired: L("KADALUARSA", "EXPIRED", "TAMAT TEMPOH"), nodate: L("TGL KOSONG", "NO DATE", "TIADA TARIKH"), soon: L("SEGERA HABIS (<3 BLN)", "EXPIRING (<3 MO)", "HAMPIR TAMAT (<3 BLN)") };
  const rows = roster.map((g) => {
    const st = DOC_REQUIREMENTS.map((r) => docStatus(g, r));
    return [g.name, g.npk || "", g.company || "", ...st.map((x) => word[x]), st.filter((x) => x !== "ok").length];
  });
  const ws = XLSX.utils.aoa_to_sheet([head, ...rows]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Documents");
  XLSX.writeFile(wb, "EcoOils-Security-Documents-Checklist.xlsx");
}
const API_URL = import.meta.env.VITE_SHEETS_API_URL || "";
const USE_REMOTE = !!API_URL;

async function apiCall(action, payload) {
  if (!USE_REMOTE) return null;
  const ctrl = new AbortController();
  const to = setTimeout(() => ctrl.abort(), 30000); // never hang forever on a slow/cold Apps Script
  try {
    const res = await fetch(API_URL, {
      signal: ctrl.signal,
      method: "POST",
      // text/plain avoids a CORS preflight; Apps Script reads the raw body
      // and JSON.parses it regardless of the declared content type.
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({ action, ...payload }),
    });
    return await res.json();
  } catch {
    return null;
  } finally {
    clearTimeout(to);
  }
}

function lsGet(key) {
  try {
    const raw = localStorage.getItem(key);
    return raw != null ? raw : null;
  } catch {
    return null;
  }
}
function lsSet(key, raw) {
  try {
    localStorage.setItem(key, raw);
    return true;
  } catch {
    return false;
  }
}
function lsDelete(key) {
  try {
    localStorage.removeItem(key);
    return true;
  } catch {
    return false;
  }
}

async function loadShared(key, fallback) {
  if (USE_REMOTE) {
    const remote = await apiCall("get", { key });
    if (!remote || !remote.ok) return fallback; // network/script error — don't silently diverge into local storage
    if (!remote.value) return fallback;
    try { return JSON.parse(remote.value); } catch { return fallback; }
  }
  const raw = lsGet(`shared:${key}`);
  if (raw == null) return fallback;
  try { return JSON.parse(raw); } catch { return fallback; }
}
// Like loadShared, but tells "the server said there is nothing" (ok, value null) apart from
// "the server could not be reached" (ok false). Retries, so a slow cold start is not mistaken
// for an empty database — that mistake used to re-seed and overwrite the saved guards.
async function loadSharedStrict(key) {
  if (!USE_REMOTE) {
    const raw = lsGet(`shared:${key}`);
    if (raw == null) return { ok: true, value: null };
    try { return { ok: true, value: JSON.parse(raw) }; } catch { return { ok: true, value: null }; }
  }
  for (let attempt = 1; attempt <= 3; attempt++) {
    const remote = await apiCall("get", { key });
    if (remote && remote.ok) {
      if (!remote.value) return { ok: true, value: null };
      try { return { ok: true, value: JSON.parse(remote.value) }; } catch { return { ok: false }; }
    }
    if (attempt < 3) await new Promise((r) => setTimeout(r, 1500 * attempt));
  }
  return { ok: false };
}
async function loadSharedPrefix(prefix) {
  const out = {};
  if (USE_REMOTE) {
    const remote = await apiCall("getPrefix", { prefix });
    if (!remote || !remote.ok) return out;
    for (const [k, v] of Object.entries(remote.items || {})) {
      try { out[k.slice(prefix.length)] = JSON.parse(v); } catch { /* skip */ }
    }
    return out;
  }
  try {
    for (let i = 0; i < localStorage.length; i++) {
      const k = localStorage.key(i);
      if (k && k.startsWith(`shared:${prefix}`)) {
        try { out[k.slice(`shared:${prefix}`.length)] = JSON.parse(localStorage.getItem(k)); } catch { /* skip */ }
      }
    }
  } catch { /* ignore */ }
  return out;
}
async function saveShared(key, value) {
  if (USE_REMOTE) {
    const remote = await apiCall("set", { key, value: JSON.stringify(value) });
    return !!(remote && remote.ok);
  }
  return lsSet(`shared:${key}`, JSON.stringify(value));
}
async function deleteShared(key) {
  if (USE_REMOTE) {
    const remote = await apiCall("delete", { key });
    return !!(remote && remote.ok);
  }
  return lsDelete(`shared:${key}`);
}

// "Personal" storage — used to remember which guard this specific device/viewer
// belongs to, so the full roster stops showing once bound. This ALWAYS stays in
// localStorage, even when the Sheets backend is configured — device binding is
// meant to be per-device, so syncing it centrally would defeat its purpose.
async function loadPersonal(key, fallback) {
  const raw = lsGet(`personal:${key}`);
  if (raw == null) return fallback;
  try { return JSON.parse(raw); } catch { return fallback; }
}
async function savePersonal(key, value) {
  return lsSet(`personal:${key}`, JSON.stringify(value));
}
async function deletePersonal(key) {
  return lsDelete(`personal:${key}`);
}

/* ---------- file upload/download ---------- */
// Each uploaded file gets its own storage key. With the Sheets backend
// configured, the actual file bytes go to Drive (via Apps Script) and only
// small metadata (name, size, Drive links) is kept under that key — Sheets
// cells cap out around 50,000 characters, far too small for a 2MB file.
const MAX_FILE_BYTES = 2 * 1024 * 1024; // 2MB

function readFileAsBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result).split(",")[1] || "");
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

// Phone photos are often 3-6 MB. Shrink big images to fit under the 2 MB limit
// (max 1600 px, JPEG) so guards' documents don't get rejected.
async function maybeShrinkImage(file) {
  if (!file.type.startsWith("image/") || file.size <= MAX_FILE_BYTES * 0.9) return file;
  try {
    const bmp = await createImageBitmap(file);
    const scale = Math.min(1, 1600 / Math.max(bmp.width, bmp.height));
    const c = document.createElement("canvas");
    c.width = Math.round(bmp.width * scale);
    c.height = Math.round(bmp.height * scale);
    c.getContext("2d").drawImage(bmp, 0, 0, c.width, c.height);
    for (const q of [0.85, 0.7, 0.55]) {
      const blob = await new Promise((r) => c.toBlob(r, "image/jpeg", q));
      if (blob && blob.size <= MAX_FILE_BYTES) {
        return new File([blob], file.name.replace(/\.[^.]+$/, "") + ".jpg", { type: "image/jpeg" });
      }
    }
  } catch { /* fall through to original */ }
  return file;
}

// Uploads `file` under its own storage key. Returns { ok, error? }.
async function uploadFileToStorage(file, storageKey, folderName) {
  if (!file) return { ok: false, error: "no-file" };
  file = await maybeShrinkImage(file);
  if (file.size > MAX_FILE_BYTES) return { ok: false, error: "too-large" };
  try {
    const base64 = await readFileAsBase64(file);
    const mimeType = file.type || "application/octet-stream";

    if (USE_REMOTE) {
      const remote = await apiCall("uploadFile", { fileName: file.name, mimeType, base64Data: base64, folderName: folderName || "" });
      if (!remote || !remote.ok) return { ok: false, error: "upload-failed" };
      const ok = await saveShared(storageKey, {
        fileName: file.name,
        mimeType,
        size: file.size,
        driveFileId: remote.file.id,
        driveUrl: remote.file.url,
        driveViewUrl: remote.file.viewUrl,
        driveDownloadUrl: remote.file.downloadUrl,
        uploadedAt: new Date().toISOString(),
      });
      return { ok, fileName: file.name, size: file.size, fileId: remote.file.id };
    }

    const ok = await saveShared(storageKey, {
      fileName: file.name,
      mimeType,
      size: file.size,
      base64,
      uploadedAt: new Date().toISOString(),
    });
    return { ok, fileName: file.name, size: file.size };
  } catch {
    return { ok: false, error: "read-failed" };
  }
}

async function deleteFileFromStorage(storageKey) {
  if (USE_REMOTE) {
    const data = await loadShared(storageKey, null);
    if (data?.driveFileId) {
      await apiCall("deleteFile", { fileId: data.driveFileId });
    }
    return deleteShared(storageKey);
  }
  return lsDelete(`shared:${storageKey}`);
}

// Opens/downloads a previously-uploaded file. Drive-hosted files (Sheets backend
// configured) just open their Drive link; local-fallback files get decoded from
// the base64 stored alongside them and downloaded directly.
function downloadStoredFile(fileData) {
  if (!fileData) return;
  if (fileData.driveDownloadUrl) {
    window.open(fileData.driveDownloadUrl, "_blank", "noopener");
    return;
  }
  if (!fileData.base64) return;
  try {
    const byteChars = atob(fileData.base64);
    const bytes = new Uint8Array(byteChars.length);
    for (let i = 0; i < byteChars.length; i++) bytes[i] = byteChars.charCodeAt(i);
    const blob = new Blob([bytes], { type: fileData.mimeType || "application/octet-stream" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = fileData.fileName || "file";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  } catch {
    // ignore — download is best-effort
  }
}

function formatFileSize(bytes) {
  if (bytes == null) return "";
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

/* ---------- design tokens ---------- */
// Shell-style palette: Forest Green, Shell Yellow, Solid Black, White, Shell Grey (see index.css).
const COLORS = {
  paper: "#FFFFFF",
  ink: "#231F20",
  inkSoft: "#6E6B6C",
  line: "#DCDCDC",
  amber: "#FFC600",
  amberDeep: "#231F20",
  green: "#44A276",
  red: "#B3261E",
  card: "#FFFFFF",
};

/* ---------- language / i18n ---------- */

const LangContext = createContext({ lang: "id", setLang: () => {} });

const STRINGS = {
  id: {
    loginSubtitle: "Pilih nama Anda untuk mengisi absensi hari ini",
    noPersonnel: "Belum ada personel. Hubungi manajer untuk mengatur data personel.",
    managerAccess: "Akses Manajer",
    managerPin: "PIN Manajer",
    enterPin: "Masukkan PIN",
    incorrectPin: "PIN salah",
    unlock: "Buka",
    back: "Kembali",
    pinHint: "Lupa PIN? Minta ke Manajer — dapat diubah di Manajer → Pengaturan.",
    switch: "Ganti",
    attendancePeriod: "Periode Absensi",
    saving: "Menyimpan…",
    today: "HARI INI",
    saved: "Tersimpan",
    undo: "Batal",
    keteranganPlaceholder: "Keterangan lokasi / lain-lain (opsional)",
    manager: "Manajer",
    exit: "Keluar",
    tabPrint: "Cetak / Verifikasi",
    tabRoster: "Data Personel",
    tabSettings: "Pengaturan",
    selectPersonnel: "Pilih personel",
    print: "Cetak",
    noPersonnelPrintHint: "Tambahkan personel di tab Data Personel untuk membuat lembar cetak.",
    importTitle: "Impor dari jadwal SOS",
    importButton: "Impor 24 personel",
    importAdded: (n) => `${n} personel ditambahkan dari jadwal SOS.`,
    importNone: "Semua personel dari jadwal SOS sudah ada di data.",
    addManualTitle: "Tambah personel manual",
    fieldName: "Nama",
    fieldNpk: "NPK",
    fieldJabatan: "Jabatan",
    fieldProjek: "Projek",
    fieldCompany: "Perusahaan (Vendor Keamanan)",
    reminderLoginTitle: "Waktunya Log In",
    reminderLoginBody: (time) => `Shift Anda dimulai pukul ${time}. Jangan lupa Log In.`,
    reminderLogoutTitle: "Waktunya Log Out",
    reminderLogoutBody: (time) => `Shift Anda berakhir pukul ${time}. Jangan lupa Log Out.`,
    fieldPhone: "No. HP",
    fieldRegu: "Regu (opsional)",
    add: "Tambah",
    noPersonnelYet: "Belum ada personel.",
    formTitleLabel: "Judul form (header cetak)",
    subLabelLabel: "Sub-label (opsional, mis. baris perusahaan/grup)",
    saveSettings: "Simpan pengaturan",
    loading: "Memuat data…",
    locating: "Mencari lokasi…",
    locationOn: "Lokasi aktif",
    locationOff: "Lokasi tidak tersedia",
    locationUnsupported: "Perangkat tidak mendukung lokasi",
    locationLog: "Catatan Lokasi",
    viewMap: "Lihat peta",
    noLocationYet: "Belum ada data lokasi untuk periode ini.",
    tagged: "Ditandai lokasi",
    danruAccess: "Akses Danru",
    danruPinTitle: "PIN Danru",
    danruPinLabel: "PIN Danru (akses data personel)",
    danruHeader: "Danru — Data Personel",
    halfShiftLabel: "Setengah Hari",
    exportExcel: "Unduh Excel (semua data)",
    exportRosterExcel: "Unduh Excel (data personel)",
    printPdfHint: "Untuk PDF: klik Cetak, lalu pilih \"Simpan sebagai PDF\" di tujuan pencetakan browser.",
    appOnlyHint: "Catatan: lembar cetak/PDF hanya untuk verifikasi. Semua koreksi dilakukan di aplikasi, bukan di hasil cetak.",
    printModeOne: "Satu personel",
    printModeAll: "Semua (1 orang/halaman)",
    tabSchedule: "Jadwal Shift",
    scheduleTitle: "Atur shift setiap personel per tanggal. Ketuk sel untuk mengubah.",
    duplicateShiftConfirm: (name, day, oldCode, newCode) => `${name} pada tanggal ${day} sudah diisi shift "${oldCode}". Ganti dengan "${newCode}"?`,
    duplicateEntryConfirm: (oldCode, newCode) => `Hari ini sudah tercatat shift "${oldCode}". Ganti dengan "${newCode}"?`,
    totalGuards: (n) => `${n} personel terdaftar`,
    certificatesTitle: "Sertifikat / Dokumen",
    noCertificates: "Belum ada sertifikat.",
    addCertificate: "Tambah sertifikat",
    certName: "Nama sertifikat",
    certNumber: "Nomor",
    certIssued: "Tanggal terbit",
    certExpiry: "Tanggal kadaluarsa",
    certLink: "Link file (OneDrive/Drive)",
    certNotes: "Catatan",
    certExpired: "kadaluarsa",
    certExpiringSoon: "segera kadaluarsa (<3 bulan)",
    openLink: "Buka file",
    bioTitle: "Data Diri",
    basicInfoTitle: "Info Dasar",
    fieldKtp: "No. KTP",
    fieldBirthPlace: "Tempat Lahir",
    fieldBirthDate: "Tanggal Lahir",
    fieldAddress: "Alamat",
    fieldBloodType: "Golongan Darah",
    fieldEmergencyName: "Kontak Darurat (Nama)",
    fieldEmergencyPhone: "Kontak Darurat (No. HP)",
    guardPinTitle: "Masukkan PIN Anda",
    guardPinHint: "Lupa PIN? Hubungi Danru atau Manajer.",
    fieldGuardPin: "PIN Personel (diatur oleh Danru/Manajer)",
    noPinSet: "(belum diatur)",
    welcomeBack: (name) => `Selamat datang kembali, ${name}`,
    continueAs: "Lanjutkan",
    notYou: "Bukan Anda?",
    resetDeviceTitle: "Reset perangkat ini",
    resetDeviceHint: "Memerlukan PIN Danru atau Manajer",
    deviceBoundHint: "Perangkat ini terkunci untuk personel di atas.",
    swipeHint: "← Geser untuk melihat tanggal lainnya →",
    todayLabel: "Hari ini",
    chooseShiftPrompt: "Pilih shift Anda hari ini",
    changeShift: "Ganti shift",
    logIn: "LOG IN",
    logOut: "LOG OUT",
    loggedInAt: (time) => `Masuk pukul ${time}`,
    allDoneToday: "Selesai untuk hari ini",
    doneSummary: (start, end, hrs) => `${start} – ${end} (${hrs} jam)`,
    offToday: "Hari ini: OFF",
    resetDeviceLink: "Bukan Anda? (perlu PIN Danru/Manajer)",
    notesLabel: "Catatan (opsional)",
    notesPlaceholder: "Tulis catatan di sini jika ada...",
    oneLoginHint: "Anda hanya bisa Log In satu kali per hari.",
    deleteConfirmTitle: (name) => `Hapus ${name}?`,
    deleteConfirmBody: "Personel ini akan dipindahkan ke bagian Personel Dihapus, bukan dihapus permanen.",
    deleteReasonLabel: "Alasan",
    reasonSelect: "Pilih alasan",
    reasonResigned: "Mengundurkan diri",
    reasonTerminated: "Diberhentikan",
    reasonOther: "Lainnya",
    reasonOtherPlaceholder: "Tulis alasan lain...",
    cancel: "Batal",
    continueBtn: "Lanjutkan",
    confirmDeleteBtn: "Hapus Personel",
    deletePinPrompt: "Masukkan PIN Manajer untuk konfirmasi",
    deletedSection: "Personel Dihapus",
    noDeletedGuards: "Belum ada personel yang dihapus.",
    restoreGuardBtn: "Pulihkan",
    deletedOn: (date) => `Dihapus ${date}`,
    deletedByLabel: (role) => `oleh ${role}`,
    siteLatLabel: "Latitude Lokasi",
    siteLngLabel: "Longitude Lokasi",
    geofenceRadiusLabel: "Radius Geofence (meter)",
    useCurrentLocation: "Gunakan lokasi saat ini",
    gettingLocation: "Mengambil lokasi…",
    withinRadius: (m) => `Dalam radius (${m}m)`,
    outsideRadius: (m) => `Di luar radius (${m}m)`,
    consentTitle: "Pendaftaran Perangkat",
    consentBody: "Saat Anda Log In dan Log Out, waktu dan lokasi Anda (dalam area lokasi kerja) akan dicatat untuk keperluan verifikasi kehadiran.",
    consentContinue: "Mengerti, Lanjutkan",
    photoLabel: "Foto",
    uploadPhoto: "Unggah Foto",
    changePhoto: "Ganti Foto",
    noPhoto: "Belum ada foto",
    fileTooLarge: "File maksimal 2 MB.",
    uploading: "Mengunggah…",
    uploadFile: "Unggah File",
    fileUploaded: (name) => `${name}`,
    downloadFile: "Unduh",
    removeFile: "Hapus File",
    employmentTitle: "Data Kepegawaian",
    fieldDateOfJoining: "Tanggal Bergabung",
    fieldExperience: "Pengalaman",
    experiencePlaceholder: "Riwayat pengalaman kerja...",
    skckTitle: "SKCK (Surat Keterangan Catatan Kepolisian)",
    skckNumber: "Nomor SKCK",
    skckIssued: "Tanggal Terbit",
    skckExpiry: "Tanggal Kadaluarsa",
    todayShiftTitle: "Shift Hari Ini",
    noShiftToday: "Belum dijadwalkan",
    warningsTitle: "Peringatan",
    noWarnings: "Belum ada peringatan.",
    addWarning: "Tambah Peringatan",
    warningReason: "Alasan peringatan",
    warningDate: "Tanggal",
    viewIdCard: "Lihat Kartu ID",
    shiftLabel: "Shift",
  },
  en: {
    loginSubtitle: "Select your name to submit today's attendance",
    noPersonnel: "No personnel added yet. Ask your manager to set up the roster.",
    managerAccess: "Manager access",
    managerPin: "Manager PIN",
    enterPin: "Enter PIN",
    incorrectPin: "Incorrect PIN",
    unlock: "Unlock",
    back: "Back",
    pinHint: "Forgot the PIN? Ask your Manager — it can be changed under Manager → Settings.",
    switch: "Switch",
    attendancePeriod: "Attendance Period",
    saving: "Saving…",
    today: "TODAY",
    saved: "Saved",
    undo: "Undo",
    keteranganPlaceholder: "Location notes / other (optional)",
    manager: "Manager",
    exit: "Exit",
    tabPrint: "Print / Verify",
    tabRoster: "Roster",
    tabSettings: "Settings",
    selectPersonnel: "Select personnel",
    print: "Print",
    noPersonnelPrintHint: "Add personnel under Roster to generate a printable sheet.",
    importTitle: "Import from SOS schedule",
    importButton: "Import 24 personnel",
    importAdded: (n) => `Added ${n} from the SOS schedule.`,
    importNone: "Everyone from the SOS schedule is already in the roster.",
    addManualTitle: "Add personnel manually",
    fieldName: "Name",
    fieldNpk: "NPK",
    fieldJabatan: "Position",
    fieldProjek: "Project",
    fieldCompany: "Security Company (Vendor)",
    reminderLoginTitle: "Time to Log In",
    reminderLoginBody: (time) => `Your shift starts at ${time}. Don't forget to Log In.`,
    reminderLogoutTitle: "Time to Log Out",
    reminderLogoutBody: (time) => `Your shift ends at ${time}. Don't forget to Log Out.`,
    fieldPhone: "Phone",
    fieldRegu: "Squad (optional)",
    add: "Add",
    noPersonnelYet: "No personnel yet.",
    formTitleLabel: "Form title (printed header)",
    subLabelLabel: "Sub-label (optional, e.g. company/group line)",
    saveSettings: "Save settings",
    loading: "Loading roster…",
    locating: "Finding location…",
    locationOn: "Location active",
    locationOff: "Location unavailable",
    locationUnsupported: "Device doesn't support location",
    locationLog: "Location Log",
    viewMap: "View map",
    noLocationYet: "No location data for this period yet.",
    tagged: "Location tagged",
    danruAccess: "Danru access",
    danruPinTitle: "Danru PIN",
    danruPinLabel: "Danru PIN (roster access)",
    danruHeader: "Danru — Personnel",
    halfShiftLabel: "Half Shift",
    exportExcel: "Download Excel (all data)",
    exportRosterExcel: "Download Excel (personnel)",
    printPdfHint: "For a PDF: click Print, then choose \"Save as PDF\" as the browser's print destination.",
    appOnlyHint: "Note: the printed/PDF sheet is for verification only. Make all corrections in the app, not on the printout.",
    printModeOne: "One person",
    printModeAll: "All (1 per page)",
    tabSchedule: "Shift Roster",
    scheduleTitle: "Set each guard's shift per date. Tap a cell to change it.",
    duplicateShiftConfirm: (name, day, oldCode, newCode) => `${name} on ${day} already has shift "${oldCode}" assigned. Replace with "${newCode}"?`,
    duplicateEntryConfirm: (oldCode, newCode) => `Today already has shift "${oldCode}" logged. Replace with "${newCode}"?`,
    totalGuards: (n) => `${n} personnel registered`,
    certificatesTitle: "Certificates / Documents",
    noCertificates: "No certificates yet.",
    addCertificate: "Add certificate",
    certName: "Certificate name",
    certNumber: "Number",
    certIssued: "Issued date",
    certExpiry: "Expiry date",
    certLink: "File link (OneDrive/Drive)",
    certNotes: "Notes",
    certExpired: "expired",
    certExpiringSoon: "expiring soon",
    openLink: "Open file",
    bioTitle: "Personal Details",
    basicInfoTitle: "Basic Info",
    fieldKtp: "ID Card No. (KTP)",
    fieldBirthPlace: "Birth Place",
    fieldBirthDate: "Birth Date",
    fieldAddress: "Address",
    fieldBloodType: "Blood Type",
    fieldEmergencyName: "Emergency Contact (Name)",
    fieldEmergencyPhone: "Emergency Contact (Phone)",
    guardPinTitle: "Enter your PIN",
    guardPinHint: "Forgot your PIN? Ask Danru or your Manager.",
    fieldGuardPin: "Personal PIN (set by Danru/Manager)",
    noPinSet: "(not set)",
    welcomeBack: (name) => `Welcome back, ${name}`,
    continueAs: "Continue",
    notYou: "Not you?",
    resetDeviceTitle: "Reset this device",
    resetDeviceHint: "Requires Danru or Manager PIN",
    deviceBoundHint: "This device is locked to the personnel above.",
    swipeHint: "← Swipe to see more dates →",
    todayLabel: "Today",
    chooseShiftPrompt: "Choose your shift for today",
    changeShift: "Change shift",
    logIn: "LOG IN",
    logOut: "LOG OUT",
    loggedInAt: (time) => `Logged in at ${time}`,
    allDoneToday: "All done for today",
    doneSummary: (start, end, hrs) => `${start} – ${end} (${hrs}h)`,
    offToday: "Today: OFF",
    resetDeviceLink: "Not you? (requires Danru/Manager PIN)",
    notesLabel: "Notes (optional)",
    notesPlaceholder: "Write any notes here...",
    oneLoginHint: "You can only Log In once per day.",
    deleteConfirmTitle: (name) => `Delete ${name}?`,
    deleteConfirmBody: "This person will be moved to Deleted Personnel, not permanently removed.",
    deleteReasonLabel: "Reason",
    reasonSelect: "Select a reason",
    reasonResigned: "Resigned",
    reasonTerminated: "Terminated",
    reasonOther: "Other",
    reasonOtherPlaceholder: "Write the reason...",
    cancel: "Cancel",
    continueBtn: "Continue",
    confirmDeleteBtn: "Delete Personnel",
    deletePinPrompt: "Enter Manager PIN to confirm",
    deletedSection: "Deleted Personnel",
    noDeletedGuards: "No deleted personnel yet.",
    restoreGuardBtn: "Restore",
    deletedOn: (date) => `Deleted ${date}`,
    deletedByLabel: (role) => `by ${role}`,
    siteLatLabel: "Site Latitude",
    siteLngLabel: "Site Longitude",
    geofenceRadiusLabel: "Geofence Radius (meters)",
    useCurrentLocation: "Use current location",
    gettingLocation: "Getting location…",
    withinRadius: (m) => `Within radius (${m}m)`,
    outsideRadius: (m) => `Outside radius (${m}m)`,
    consentTitle: "Device Registration",
    consentBody: "When you Log In and Log Out, your time and location (within the work site area) will be recorded for attendance verification purposes.",
    consentContinue: "Understood, Continue",
    photoLabel: "Photo",
    uploadPhoto: "Upload Photo",
    changePhoto: "Change Photo",
    noPhoto: "No photo yet",
    fileTooLarge: "File must be under 2 MB.",
    uploading: "Uploading…",
    uploadFile: "Upload File",
    fileUploaded: (name) => `${name}`,
    downloadFile: "Download",
    removeFile: "Remove File",
    employmentTitle: "Employment Details",
    fieldDateOfJoining: "Date of Joining",
    fieldExperience: "Experience",
    experiencePlaceholder: "Prior work experience...",
    skckTitle: "Police Clearance (SKCK)",
    skckNumber: "SKCK Number",
    skckIssued: "Issued Date",
    skckExpiry: "Expiry Date",
    todayShiftTitle: "Today's Shift",
    noShiftToday: "Not scheduled",
    warningsTitle: "Warnings",
    noWarnings: "No warnings yet.",
    addWarning: "Add Warning",
    warningReason: "Warning reason",
    warningDate: "Date",
    viewIdCard: "View ID Card",
    shiftLabel: "Shift",
  },
  ms: {
    loginSubtitle: "Pilih nama anda untuk menghantar kehadiran hari ini",
    noPersonnel: "Belum ada anggota ditambah. Minta pengurus anda menetapkan senarai anggota.",
    managerAccess: "Akses Pengurus",
    managerPin: "PIN Pengurus",
    enterPin: "Masukkan PIN",
    incorrectPin: "PIN salah",
    unlock: "Buka",
    back: "Kembali",
    pinHint: "Lupa PIN? Minta daripada Pengurus anda — boleh ditukar di Pengurus → Tetapan.",
    switch: "Tukar",
    attendancePeriod: "Tempoh Kehadiran",
    saving: "Menyimpan…",
    today: "HARI INI",
    saved: "Disimpan",
    undo: "Buat asal",
    keteranganPlaceholder: "Nota lokasi / lain-lain (pilihan)",
    manager: "Pengurus",
    exit: "Keluar",
    tabPrint: "Cetak / Sahkan",
    tabRoster: "Senarai Anggota",
    tabSettings: "Tetapan",
    selectPersonnel: "Pilih anggota",
    print: "Cetak",
    noPersonnelPrintHint: "Tambah anggota di Senarai Anggota untuk menjana helaian yang boleh dicetak.",
    importTitle: "Import daripada jadual SOS",
    importButton: "Import 24 anggota",
    importAdded: (n) => `${n} ditambah daripada jadual SOS.`,
    importNone: "Semua orang dalam jadual SOS sudah ada dalam senarai.",
    addManualTitle: "Tambah anggota secara manual",
    fieldName: "Nama",
    fieldNpk: "NPK",
    fieldJabatan: "Jawatan",
    fieldProjek: "Projek",
    fieldCompany: "Syarikat Keselamatan (Vendor)",
    reminderLoginTitle: "Masa untuk Log Masuk",
    reminderLoginBody: (time) => `Syif anda bermula pada ${time}. Jangan lupa Log Masuk.`,
    reminderLogoutTitle: "Masa untuk Log Keluar",
    reminderLogoutBody: (time) => `Syif anda tamat pada ${time}. Jangan lupa Log Keluar.`,
    fieldPhone: "Telefon",
    fieldRegu: "Regu (pilihan)",
    add: "Tambah",
    noPersonnelYet: "Belum ada anggota.",
    formTitleLabel: "Tajuk borang (pengepala cetakan)",
    subLabelLabel: "Sub-label (pilihan, cth. baris syarikat/kumpulan)",
    saveSettings: "Simpan tetapan",
    loading: "Memuatkan senarai…",
    locating: "Mencari lokasi…",
    locationOn: "Lokasi aktif",
    locationOff: "Lokasi tidak tersedia",
    locationUnsupported: "Peranti tidak menyokong lokasi",
    locationLog: "Log Lokasi",
    viewMap: "Lihat peta",
    noLocationYet: "Belum ada data lokasi untuk tempoh ini.",
    tagged: "Lokasi ditandakan",
    danruAccess: "Akses Danru",
    danruPinTitle: "PIN Danru",
    danruPinLabel: "PIN Danru (akses senarai anggota)",
    danruHeader: "Danru — Anggota",
    halfShiftLabel: "Separuh Syif",
    exportExcel: "Muat turun Excel (semua data)",
    exportRosterExcel: "Muat turun Excel (anggota)",
    printPdfHint: "Untuk PDF: klik Cetak, kemudian pilih \"Save as PDF\" sebagai destinasi cetakan pelayar.",
    appOnlyHint: "Nota: helaian cetakan/PDF hanya untuk pengesahan. Buat semua pembetulan dalam aplikasi, bukan pada cetakan.",
    printModeOne: "Seorang",
    printModeAll: "Semua (1 sehelai)",
    tabSchedule: "Jadual Syif",
    scheduleTitle: "Tetapkan syif setiap pengawal mengikut tarikh. Ketik sel untuk menukar.",
    duplicateShiftConfirm: (name, day, oldCode, newCode) => `${name} pada ${day} sudah ditetapkan syif "${oldCode}". Gantikan dengan "${newCode}"?`,
    duplicateEntryConfirm: (oldCode, newCode) => `Hari ini sudah direkodkan syif "${oldCode}". Gantikan dengan "${newCode}"?`,
    totalGuards: (n) => `${n} anggota berdaftar`,
    certificatesTitle: "Sijil / Dokumen",
    noCertificates: "Belum ada sijil.",
    addCertificate: "Tambah sijil",
    certName: "Nama sijil",
    certNumber: "Nombor",
    certIssued: "Tarikh dikeluarkan",
    certExpiry: "Tarikh tamat tempoh",
    certLink: "Pautan fail (OneDrive/Drive)",
    certNotes: "Nota",
    certExpired: "tamat tempoh",
    certExpiringSoon: "hampir tamat tempoh",
    openLink: "Buka fail",
    bioTitle: "Butiran Peribadi",
    basicInfoTitle: "Maklumat Asas",
    fieldKtp: "No. Kad Pengenalan (KTP)",
    fieldBirthPlace: "Tempat Lahir",
    fieldBirthDate: "Tarikh Lahir",
    fieldAddress: "Alamat",
    fieldBloodType: "Jenis Darah",
    fieldEmergencyName: "Hubungan Kecemasan (Nama)",
    fieldEmergencyPhone: "Hubungan Kecemasan (Telefon)",
    guardPinTitle: "Masukkan PIN anda",
    guardPinHint: "Lupa PIN? Tanya Danru atau Pengurus anda.",
    fieldGuardPin: "PIN Peribadi (ditetapkan oleh Danru/Pengurus)",
    noPinSet: "(belum ditetapkan)",
    welcomeBack: (name) => `Selamat kembali, ${name}`,
    continueAs: "Teruskan",
    notYou: "Bukan anda?",
    resetDeviceTitle: "Set semula peranti ini",
    resetDeviceHint: "Memerlukan PIN Danru atau Pengurus",
    deviceBoundHint: "Peranti ini dikunci kepada anggota di atas.",
    swipeHint: "← Leret untuk lihat lebih banyak tarikh →",
    todayLabel: "Hari ini",
    chooseShiftPrompt: "Pilih syif anda untuk hari ini",
    changeShift: "Tukar syif",
    logIn: "LOG MASUK",
    logOut: "LOG KELUAR",
    loggedInAt: (time) => `Log masuk pada ${time}`,
    allDoneToday: "Semua selesai untuk hari ini",
    doneSummary: (start, end, hrs) => `${start} – ${end} (${hrs}j)`,
    offToday: "Hari ini: CUTI",
    resetDeviceLink: "Bukan anda? (memerlukan PIN Danru/Pengurus)",
    notesLabel: "Nota (pilihan)",
    notesPlaceholder: "Tulis sebarang nota di sini...",
    oneLoginHint: "Anda hanya boleh Log Masuk sekali sehari.",
    deleteConfirmTitle: (name) => `Padam ${name}?`,
    deleteConfirmBody: "Orang ini akan dipindahkan ke Anggota Dipadam, bukan dibuang secara kekal.",
    deleteReasonLabel: "Sebab",
    reasonSelect: "Pilih sebab",
    reasonResigned: "Meletak jawatan",
    reasonTerminated: "Diberhentikan",
    reasonOther: "Lain-lain",
    reasonOtherPlaceholder: "Tulis sebab...",
    cancel: "Batal",
    continueBtn: "Teruskan",
    confirmDeleteBtn: "Padam Anggota",
    deletePinPrompt: "Masukkan PIN Pengurus untuk mengesahkan",
    deletedSection: "Anggota Dipadam",
    noDeletedGuards: "Belum ada anggota dipadam.",
    restoreGuardBtn: "Pulihkan",
    deletedOn: (date) => `Dipadam ${date}`,
    deletedByLabel: (role) => `oleh ${role}`,
    siteLatLabel: "Latitud Tapak",
    siteLngLabel: "Longitud Tapak",
    geofenceRadiusLabel: "Radius Geopagar (meter)",
    useCurrentLocation: "Guna lokasi semasa",
    gettingLocation: "Mendapatkan lokasi…",
    withinRadius: (m) => `Dalam radius (${m}m)`,
    outsideRadius: (m) => `Luar radius (${m}m)`,
    consentTitle: "Pendaftaran Peranti",
    consentBody: "Apabila anda Log Masuk dan Log Keluar, masa dan lokasi anda (dalam kawasan tapak kerja) akan direkodkan untuk tujuan pengesahan kehadiran.",
    consentContinue: "Faham, Teruskan",
    photoLabel: "Gambar",
    uploadPhoto: "Muat Naik Gambar",
    changePhoto: "Tukar Gambar",
    noPhoto: "Belum ada gambar",
    fileTooLarge: "Fail mestilah kurang daripada 2 MB.",
    uploading: "Memuat naik…",
    uploadFile: "Muat Naik Fail",
    fileUploaded: (name) => `${name}`,
    downloadFile: "Muat turun",
    removeFile: "Buang Fail",
    employmentTitle: "Butiran Pekerjaan",
    fieldDateOfJoining: "Tarikh Mula Bekerja",
    fieldExperience: "Pengalaman",
    experiencePlaceholder: "Pengalaman kerja terdahulu...",
    skckTitle: "Pelepasan Polis (SKCK)",
    skckNumber: "Nombor SKCK",
    skckIssued: "Tarikh Dikeluarkan",
    skckExpiry: "Tarikh Tamat Tempoh",
    todayShiftTitle: "Syif Hari Ini",
    noShiftToday: "Tidak dijadualkan",
    warningsTitle: "Amaran",
    noWarnings: "Belum ada amaran.",
    addWarning: "Tambah Amaran",
    warningReason: "Sebab amaran",
    warningDate: "Tarikh",
    viewIdCard: "Lihat Kad Pengenalan",
    shiftLabel: "Syif",
  },
};

function useT() {
  const { lang } = useContext(LangContext);
  return (key, ...args) => {
    const v = (STRINGS[lang] && STRINGS[lang][key] !== undefined) ? STRINGS[lang][key] : STRINGS.en[key];
    return typeof v === "function" ? v(...args) : v;
  };
}

function LangToggle() {
  const { lang, setLang } = useContext(LangContext);
  return (
    <header className="eh no-print">
      <img className="eh-logo" alt="EcoOils Mentari" src={ECO_LOGO} />
      <div className="eh-r">
        <div className="eh-v">v{APP_VERSION}</div>
        <div className="eh-l eco-lang" role="group" aria-label="Language">
          {["en", "id", "ms"].map((l) => (
            <button key={l} onClick={() => setLang(l)} className={lang === l ? "on" : ""} aria-pressed={lang === l}>
              {l.toUpperCase()}
            </button>
          ))}
        </div>
        <div className="eh-d">v{APP_VERSION} · {APP_DATE}</div>
      </div>
    </header>
  );
}

// 5 x 5 Shell-grey dot grid: used for the corner watermark and for empty states.
function DotGrid({ size = 56, className = "" }) {
  const dots = [];
  for (let r = 0; r < 5; r++) for (let c = 0; c < 5; c++) dots.push(<circle key={r + "-" + c} cx={6 + c * 12} cy={6 + r * 12} r="3.2" fill="#AAAAAA" />);
  return <svg className={className} width={size} height={size} viewBox="0 0 60 60" aria-hidden="true">{dots}</svg>;
}

/* ---------- root component ---------- */

export default function SecurityTimesheetApp() {
  const [loaded, setLoaded] = useState(false);
  const [roster, setRoster] = useState([]);
  const [entries, setEntries] = useState({});
  const [settings, setSettings] = useState({ pin: "0822", danruPin: "2026", formTitle: "FORM KEHADIRAN REGULER", org: "", siteLat: "", siteLng: "", geofenceRadius: "150" });
  const [view, setView] = useState("login"); // login | guard | manager-gate | manager | danru-gate | danru
  const [currentGuardId, setCurrentGuardId] = useState(null);
  const [pendingGuardId, setPendingGuardId] = useState(null);
  const [boundGuardId, setBoundGuardId] = useState(null);
  const [managerAuthed, setManagerAuthed] = useState(false);
  const [danruAuthed, setDanruAuthed] = useState(false);
  const [pinInput, setPinInput] = useState("");
  const [pinError, setPinError] = useState("");
  const [saving, setSaving] = useState(false);
  const [lang, setLang] = useState("id");
  const [schedule, setSchedule] = useState({});

  const [loadError, setLoadError] = useState(false);
  const loadAll = useCallback(async () => {
    setLoadError(false);
    const [rr, ss, sc, bound] = await Promise.all([
      loadSharedStrict("roster"),
      loadSharedStrict("settings"),
      loadSharedStrict("schedule"),
      loadPersonal("boundGuard", null),
    ]);
    // If the server could not be reached, stop here. Never seed or save anything:
    // that would overwrite the real roster with the default list.
    if (!rr.ok || !ss.ok || !sc.ok) { setLoadError(true); return; }
    const r = Array.isArray(rr.value) ? rr.value : [];
    const s = ss.value || {};
    const sch = sc.value || {};
    let finalRoster = r;
    if (r.length === 0) {
      const seeded = SOS_ROSTER.map((g) => ({ ...g, id: uid() }));
      finalRoster = seeded;
      setRoster(seeded);
      rosterRef.current = seeded;
      await saveShared("roster", seeded);
    } else {
      setRoster(r);
      rosterRef.current = r;
    }
    setEntries({});
    setSettings({ pin: "0822", danruPin: "2026", formTitle: "FORM KEHADIRAN REGULER", org: "", siteLat: "", siteLng: "", geofenceRadius: "150", ...s });
    setSchedule(sch);
    if (!PHASE_DB && bound?.guardId && finalRoster.some((g) => g.id === bound.guardId)) {
      setBoundGuardId(bound.guardId);
      setCurrentGuardId(bound.guardId);
      setView("guard");
      const mine = await loadShared(`entries:${bound.guardId}`, {});
      setEntries({ [bound.guardId]: mine });
    }
    setLoaded(true);
  }, []);
  useEffect(() => { loadAll(); }, [loadAll]);

  // Roster edits apply instantly on screen; the save to Google Sheets is debounced
  // (one write ~0.8 s after the last change) and always sends the latest roster.
  const rosterRef = React.useRef([]);
  const saveTimer = React.useRef(null);
  const flushRoster = useCallback(async () => {
    clearTimeout(saveTimer.current);
    saveTimer.current = null;
    await saveShared("roster", rosterRef.current);
    setSaving(false);
  }, []);
  const persistRoster = useCallback(async (next) => {
    rosterRef.current = next;
    setRoster(next);
    setSaving(true);
    clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(flushRoster, 800);
  }, [flushRoster]);
  useEffect(() => {
    function onHide() { if (saveTimer.current) flushRoster(); }
    document.addEventListener("visibilitychange", onHide);
    window.addEventListener("pagehide", onHide);
    return () => { document.removeEventListener("visibilitychange", onHide); window.removeEventListener("pagehide", onHide); };
  }, [flushRoster]);

  const softDeleteGuard = useCallback(async (id, reason, deletedBy) => {
    await persistRoster(rosterRef.current.map((g) => (g.id === id
      ? { ...g, deleted: true, deletedReason: reason, deletedAt: new Date().toISOString(), deletedBy }
      : g)));
  }, [roster, persistRoster]);

  const restoreGuard = useCallback(async (id) => {
    await persistRoster(rosterRef.current.map((g) => (g.id === id
      ? { ...g, deleted: false, deletedReason: null, deletedAt: null, deletedBy: null }
      : g)));
  }, [roster, persistRoster]);

  // Each guard's entries live under their own key (entries:<guardId>), so two
  // guards saving at the same moment can never overwrite each other.
  const saveGuardEntry = useCallback(async (guardId, dateKeyStr, entry) => {
    setSaving(true);
    const fresh = await loadShared(`entries:${guardId}`, null);
    const base = fresh || {};
    const merged = { ...base, [dateKeyStr]: entry };
    const ok = await saveShared(`entries:${guardId}`, merged);
    setEntries((prev) => ({ ...prev, [guardId]: merged }));
    setSaving(false);
    return ok;
  }, []);

  const loadGuardEntries = useCallback(async (guardId) => {
    const mine = await loadShared(`entries:${guardId}`, {});
    setEntries((prev) => ({ ...prev, [guardId]: mine }));
  }, []);

  const loadAllEntries = useCallback(async () => {
    const all = await loadSharedPrefix("entries:");
    setEntries((prev) => ({ ...prev, ...all }));
  }, []);

  useEffect(() => {
    if (managerAuthed) loadAllEntries();
  }, [managerAuthed, loadAllEntries]);

  useEffect(() => {
    if (currentGuardId && !managerAuthed && !danruAuthed) loadGuardEntries(currentGuardId);
  }, [currentGuardId, managerAuthed, danruAuthed, loadGuardEntries]);

  const setScheduleCell = useCallback(async (guardId, dk, code) => {
    setSchedule((prev) => {
      const next = { ...prev, [guardId]: { ...(prev[guardId] || {}), [dk]: code } };
      saveShared("schedule", next);
      return next;
    });
  }, []);

  const importSosRoster = useCallback(async () => {
    const existingNames = new Set(roster.map((g) => g.name.trim().toLowerCase()));
    const additions = SOS_ROSTER.filter((g) => !existingNames.has(g.name.trim().toLowerCase())).map((g) => ({ ...g, id: uid() }));
    if (additions.length === 0) return 0;
    await persistRoster([...roster, ...additions]);
    return additions.length;
  }, [roster, persistRoster]);

  const persistSettings = useCallback(async (next) => {
    setSettings(next);
    setSaving(true);
    await saveShared("settings", next);
    setSaving(false);
  }, []);

  const currentGuard = useMemo(
    () => roster.find((g) => g.id === currentGuardId) || null,
    [roster, currentGuardId]
  );

  // Deleted personnel are kept in `roster` (soft-delete, for the archive) but
  // hidden from login, scheduling, and print/select dropdowns.
  const activeRoster = useMemo(() => roster.filter((g) => !g.deleted), [roster]);

  function handleLogout() {
    setPendingGuardId(null);
    setManagerAuthed(false);
    setDanruAuthed(false);
    setPinInput("");
    setPinError("");
    if (boundGuardId) {
      setCurrentGuardId(boundGuardId);
      setView("guard");
    } else {
      setCurrentGuardId(null);
      setView("login");
    }
  }

  function handleRequestDeviceReset() {
    setPinInput("");
    setPinError("");
    setView("device-reset-gate");
  }

  if (loadError) {
    return (
      <LangContext.Provider value={{ lang, setLang }}>
        <div className="eco-app" style={{ minHeight: "100vh" }}>
          <LangToggle />
          <div className="max-w-[780px] mx-auto px-4">
            <h1 className="eco-title">EcoOils Security</h1>
            <div className="eco-rule" />
            <p style={{ color: COLORS.ink }} className="text-sm mt-6">
              {pick(lang, "Tidak dapat terhubung ke server. Data Anda aman — tidak ada yang diubah.", "Can't reach the server. Your data is safe — nothing was changed.", "Tidak dapat menghubungi pelayan. Data anda selamat — tiada apa yang diubah.")}
            </p>
            <button onClick={loadAll} style={{ background: COLORS.green, color: COLORS.ink }} className="rounded-lg px-6 font-bold text-sm min-h-[48px] mt-4">
              {pick(lang, "Coba lagi", "Try again", "Cuba lagi")}
            </button>
          </div>
        </div>
      </LangContext.Provider>
    );
  }

  if (!loaded) {
    return (
      <div style={{ background: COLORS.paper, minHeight: "100vh" }} className="flex items-center justify-center">
        <div className="flex flex-col items-center gap-3">
          <DotGrid className="animate-pulse" size={56} />
          <p style={{ color: COLORS.inkSoft }} className="text-sm">{STRINGS[lang].loading}</p>
        </div>
      </div>
    );
  }

  return (
    <LangContext.Provider value={{ lang, setLang }}>
    <div className="eco-app" style={{ minHeight: "100vh" }}>
      <div className="eco-dots no-print" aria-hidden="true" />
      <style>{`
        @page { size: A4; margin: 12mm; }
        @media print {
          body * { visibility: hidden; }
          .print-area, .print-area * { visibility: visible; }
          .print-area { position: absolute; top: 0; left: 0; width: 100%; }
          .no-print { display: none !important; }
        }
        input[type="time"]::-webkit-calendar-picker-indicator { opacity: 0.6; }
      `}</style>

      <LangToggle />

      {view === "login" && (
        <LoginScreen
          roster={activeRoster}
          onPickGuard={(id) => {
            const g = roster.find((r) => r.id === id);
            if (g?.pin) {
              setPendingGuardId(id);
              setView("guard-pin-gate");
            } else {
              setCurrentGuardId(id);
              setView("guard");
            }
          }}
          onManager={() => setView("manager-gate")}
          onDanru={() => setView("danru-gate")}
        />
      )}

      {view === "guard-pin-gate" && (
        <PinGate
          title={STRINGS[lang].guardPinTitle}
          pinInput={pinInput}
          setPinInput={setPinInput}
          pinError={pinError}
          hideHint
          onBack={() => { setPinInput(""); setPinError(""); setPendingGuardId(null); setView("login"); }}
          onSubmit={async () => {
            const g = roster.find((r) => r.id === pendingGuardId);
            if (g && pinInput === g.pin) {
              setCurrentGuardId(pendingGuardId);
              setBoundGuardId(pendingGuardId);
              await savePersonal("boundGuard", { guardId: pendingGuardId });
              setPendingGuardId(null);
              setPinInput("");
              setPinError("");
              setView("guard-consent");
            } else {
              setPinError(STRINGS[lang].incorrectPin);
              setPinInput("");
            }
          }}
        />
      )}

      {view === "guard-consent" && (
        <ConsentGate onContinue={() => setView("guard")} />
      )}

      {view === "device-reset-gate" && (
        <PinGate
          title={STRINGS[lang].resetDeviceTitle}
          pinInput={pinInput}
          setPinInput={setPinInput}
          pinError={pinError}
          hint={STRINGS[lang].resetDeviceHint}
          onBack={() => { setPinInput(""); setPinError(""); setView("guard"); }}
          onSubmit={async () => {
            if (pinInput === settings.pin || pinInput === settings.danruPin) {
              await deletePersonal("boundGuard");
              setBoundGuardId(null);
              setCurrentGuardId(null);
              setPinInput("");
              setPinError("");
              setView("login");
            } else {
              setPinError(STRINGS[lang].incorrectPin);
              setPinInput("");
            }
          }}
        />
      )}

      {view === "manager-gate" && (
        <PinGate
          title={STRINGS[lang].managerPin}
          pinInput={pinInput}
          setPinInput={setPinInput}
          pinError={pinError}
          onBack={() => { setPinInput(""); setPinError(""); setView(boundGuardId ? "guard" : "login"); }}
          onSubmit={() => {
            if (pinInput === settings.pin) {
              setManagerAuthed(true);
              setPinInput("");
              setPinError("");
              setView("manager");
            } else {
              setPinError(STRINGS[lang].incorrectPin);
              setPinInput("");
            }
          }}
        />
      )}

      {view === "danru-gate" && (
        <PinGate
          title={STRINGS[lang].danruPinTitle}
          pinInput={pinInput}
          setPinInput={setPinInput}
          pinError={pinError}
          onBack={() => { setPinInput(""); setPinError(""); setView(boundGuardId ? "guard" : "login"); }}
          onSubmit={() => {
            if (pinInput === settings.danruPin) {
              setDanruAuthed(true);
              setPinInput("");
              setPinError("");
              setView("danru");
            } else {
              setPinError(STRINGS[lang].incorrectPin);
              setPinInput("");
            }
          }}
        />
      )}

      {view === "guard" && currentGuard && (
        <GuardView
          guard={currentGuard}
          entries={entries[currentGuard.id] || {}}
          onSaveEntry={async (dateKeyStr, entry) => {
            await saveGuardEntry(currentGuard.id, dateKeyStr, entry);
          }}
          onLogout={boundGuardId === currentGuard.id ? handleRequestDeviceReset : handleLogout}
          boundDevice={boundGuardId === currentGuard.id}
          onDanru={() => setView("danru-gate")}
          onManager={() => setView("manager-gate")}
          settings={settings}
          schedule={schedule}
          saving={saving}
        />
      )}

      {view === "danru" && danruAuthed && (
        <DanruView
          roster={roster}
          activeRoster={activeRoster}
          entries={entries}
          settings={settings}
          schedule={schedule}
          onAddGuard={async (g) => { await persistRoster([...rosterRef.current, { ...g, id: uid() }]); }}
          onSoftDeleteGuard={softDeleteGuard}
          onUpdateGuard={async (id, patch) => { await persistRoster(rosterRef.current.map((g) => (g.id === id ? { ...g, ...patch } : g))); }}
          onImportRoster={importSosRoster}
          onSetScheduleCell={setScheduleCell}
          onLogout={handleLogout}
          saving={saving}
        />
      )}

      {view === "manager" && managerAuthed && (
        <ManagerView
          roster={roster}
          activeRoster={activeRoster}
          entries={entries}
          settings={settings}
          schedule={schedule}
          onAddGuard={async (g) => { await persistRoster([...rosterRef.current, { ...g, id: uid() }]); }}
          onSoftDeleteGuard={softDeleteGuard}
          onRestoreGuard={restoreGuard}
          onUpdateGuard={async (id, patch) => { await persistRoster(rosterRef.current.map((g) => (g.id === id ? { ...g, ...patch } : g))); }}
          onImportRoster={importSosRoster}
          onSetScheduleCell={setScheduleCell}
          onUpdateSettings={async (patch) => { await persistSettings({ ...settings, ...patch }); }}
          onLogout={handleLogout}
          saving={saving}
        />
      )}
    </div>
    </LangContext.Provider>
  );
}

/* ---------- login screen ---------- */

function LoginScreen({ roster, onPickGuard, onManager, onDanru }) {
  const t = useT();
  const showGuards = !PHASE_DB;
  return (
    <div className="min-h-screen flex flex-col px-4 pb-10 max-w-[780px] mx-auto w-full">
      <div className="mb-6">
        <h1 className="eco-title">EcoOils Security</h1>
        <div className="eco-rule" />
        <p style={{ color: COLORS.inkSoft }} className="text-sm mt-4">{PHASE_DB ? t("danruHeader") : t("loginSubtitle")}</p>
        {roster.length > 0 && (
          <p style={{ color: COLORS.ink }} className="text-xs font-medium flex items-center gap-1.5 mt-2">
            <Users size={13} /> {t("totalGuards", roster.length)}
          </p>
        )}
      </div>

      {showGuards && <div className="w-full max-w-sm flex flex-col gap-2">
        {roster.length === 0 && (
          <div style={{ background: COLORS.card, borderColor: COLORS.line }} className="border rounded-xl p-5 text-center">
            <p style={{ color: COLORS.inkSoft }} className="text-sm">{t("noPersonnel")}</p>
          </div>
        )}
        {roster.map((g) => (
          <button
            key={g.id}
            onClick={() => onPickGuard(g.id)}
            style={{ background: COLORS.card, borderColor: COLORS.line, color: COLORS.ink }}
            className="border rounded-xl px-4 py-3 text-left font-medium hover:shadow-md transition-shadow flex items-center justify-between"
          >
            <span>{g.name}</span>
            <span className="flex items-center gap-1.5">
              {g.pin && <Lock size={11} color={COLORS.inkSoft} />}
              <span style={{ color: COLORS.inkSoft }} className="text-xs font-normal">{g.jabatan}</span>
            </span>
          </button>
        ))}
      </div>}

      <div className="mt-6 flex flex-col sm:flex-row gap-3 w-full max-w-sm">
        <button
          onClick={onDanru}
          style={{ background: COLORS.green, color: COLORS.ink }}
          className="rounded-lg px-4 font-bold text-sm flex items-center justify-center gap-2 min-h-[48px] flex-1"
        >
          <Lock size={14} /> {t("danruAccess")}
        </button>
        <button
          onClick={onManager}
          style={{ background: COLORS.card, borderColor: COLORS.line, color: COLORS.ink }}
          className="border rounded-lg px-4 font-medium text-sm flex items-center justify-center gap-2 min-h-[48px] flex-1"
        >
          <Lock size={14} /> {t("managerAccess")}
        </button>
      </div>
    </div>
  );
}

/* ---------- consent gate: one-time disclosure shown once, at device registration ---------- */

function ConsentGate({ onContinue }) {
  const t = useT();
  return (
    <div className="min-h-screen flex flex-col items-center justify-center px-4">
      <div style={{ background: COLORS.card, borderColor: COLORS.line }} className="border rounded-xl p-6 w-full max-w-sm flex flex-col gap-4">
        <div className="flex flex-col items-center gap-2">
          <div style={{ background: COLORS.ink }} className="w-12 h-12 rounded-full flex items-center justify-center">
            <MapPin color={COLORS.amber} size={20} />
          </div>
          <h2 style={{ color: COLORS.ink }} className="font-semibold text-center">{t("consentTitle")}</h2>
        </div>
        <p style={{ color: COLORS.inkSoft }} className="text-sm text-center">{t("consentBody")}</p>
        <button
          onClick={onContinue}
          style={{ background: COLORS.green, color: COLORS.ink }}
          className="font-bold rounded-lg py-2.5 text-sm font-medium min-h-[44px]"
        >
          {t("consentContinue")}
        </button>
      </div>
    </div>
  );
}

/* ---------- pin gate ---------- */

function PinGate({ title, pinInput, setPinInput, pinError, onBack, onSubmit, hideHint, hint }) {
  const t = useT();
  return (
    <div className="min-h-screen flex flex-col items-center justify-center px-4">
      <div style={{ background: COLORS.card, borderColor: COLORS.line }} className="border rounded-xl p-6 w-full max-w-xs flex flex-col gap-4">
        <div className="flex items-center gap-2">
          <Lock size={18} color={COLORS.amber} />
          <h2 style={{ color: COLORS.ink }} className="font-semibold">{title || t("managerPin")}</h2>
        </div>
        <input
          type="password"
          inputMode="numeric"
          value={pinInput}
          onChange={(e) => setPinInput(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && onSubmit()}
          placeholder={t("enterPin")}
          style={{ borderColor: COLORS.line }}
          className="border rounded-lg px-3 py-2 text-center tracking-widest text-lg"
          autoFocus
        />
        {pinError && <p style={{ color: COLORS.red }} className="text-xs text-center">{pinError}</p>}
        <button onClick={onSubmit} style={{ background: COLORS.green, color: COLORS.ink }} className="font-bold rounded-lg py-2 text-sm font-medium">
          {t("unlock")}
        </button>
        <button onClick={onBack} style={{ color: COLORS.inkSoft }} className="text-xs flex items-center justify-center gap-1">
          <ArrowLeft size={12} /> {t("back")}
        </button>
        {hint !== undefined ? (
          hint && <p style={{ color: COLORS.inkSoft }} className="text-[11px] text-center">{hint}</p>
        ) : (
          <p style={{ color: COLORS.inkSoft }} className="text-[11px] text-center">{hideHint ? t("guardPinHint") : t("pinHint")}</p>
        )}
      </div>
    </div>
  );
}

/* ---------- guard view ---------- */

function nowHHMM() {
  const d = new Date();
  return `${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

// Fires a local notification. Uses the service worker (more reliable on mobile/PWA
// installs) when available, falling back to a plain Notification otherwise. Silently
// does nothing if notifications aren't supported or permission was never granted —
// reminders are a nice-to-have, never something that should interrupt the guard flow.
async function fireNotification(title, body) {
  if (typeof Notification === "undefined" || Notification.permission !== "granted") return;
  try {
    if (navigator.serviceWorker) {
      const reg = await navigator.serviceWorker.ready;
      reg.showNotification(title, { body, icon: "icon-192.png", badge: "icon-192.png" });
      return;
    }
  } catch {
    // fall through to the plain Notification API below
  }
  try { new Notification(title, { body }); } catch {}
}

function GuardView({ guard, entries, onSaveEntry, onLogout, saving, boundDevice, onDanru, onManager, settings, schedule }) {
  const t = useT();
  const { lang } = useContext(LangContext);
  const [location, setLocation] = useState(null);

  useEffect(() => {
    let cancelled = false;
    captureLocation().then((res) => {
      if (cancelled || res.status !== "ok") return;
      setLocation(res);
    });
    return () => { cancelled = true; };
  }, []);

  const now = new Date();
  const today = { y: now.getFullYear(), m: now.getMonth(), d: now.getDate(), key: dateKey(now.getFullYear(), now.getMonth(), now.getDate()) };
  const entry = entries[today.key];

  // Ask for notification permission once, quietly — never re-prompt if the guard
  // already answered (granted or denied).
  useEffect(() => {
    if (typeof Notification !== "undefined" && Notification.permission === "default") {
      Notification.requestPermission().catch(() => {});
    }
  }, []);

  // Reminders 10 minutes before the guard's ASSIGNED shift start and end (from the
  // Danru/Manager shift roster) — only for fixed-hours shift codes (not OFF, and not
  // Half Shift, which has no fixed time to count down from), and only for whichever
  // of login/logout hasn't happened yet.
  useEffect(() => {
    const assignedCode = schedule?.[guard.id]?.[today.key];
    const shift = assignedCode ? SHIFT_CODES[assignedCode] : null;
    if (!shift || shift.custom) return;

    const timers = [];
    function scheduleReminder(timeStr, isEnd, alreadyDone) {
      if (alreadyDone || !timeStr) return;
      const [h, m] = timeStr.split(":").map(Number);
      const target = new Date(today.y, today.m, today.d, h, m, 0, 0);
      if (isEnd && shift.selesai < shift.mulai) target.setDate(target.getDate() + 1); // overnight (Malam) shift
      const delay = target.getTime() - 10 * 60 * 1000 - Date.now();
      if (delay <= 0 || delay > 24 * 60 * 60 * 1000) return; // already past, or absurdly far off
      timers.push(setTimeout(() => {
        fireNotification(
          isEnd ? t("reminderLogoutTitle") : t("reminderLoginTitle"),
          isEnd ? t("reminderLogoutBody", timeStr) : t("reminderLoginBody", timeStr)
        );
      }, delay));
    }
    scheduleReminder(shift.mulai, false, !!entry?.jamMulai);
    scheduleReminder(shift.selesai, true, !!entry?.jamSelesai);

    return () => timers.forEach(clearTimeout);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [schedule, guard.id, today.key, entry?.jamMulai, entry?.jamSelesai, lang]);

  function save(patch) {
    const loc = tagGeofence(location || entry?.location || null, settings);
    onSaveEntry(today.key, {
      shiftCode: entry?.shiftCode || "",
      jamMulai: entry?.jamMulai || "",
      jamSelesai: entry?.jamSelesai || "",
      keterangan: entry?.keterangan || "",
      ...patch,
      location: loc,
      updatedAt: new Date().toISOString(),
    });
  }

  function pickShift(code) {
    if (entry?.shiftCode && entry.shiftCode !== code && !entry?.jamMulai) {
      const ok = window.confirm(t("duplicateEntryConfirm", entry.shiftCode, code));
      if (!ok) return;
    }
    if (code === "OFF") {
      save({ shiftCode: "OFF", jamMulai: "", jamSelesai: "" });
    } else {
      save({ shiftCode: code });
    }
  }

  function logIn() {
    save({ jamMulai: nowHHMM() });
  }

  function logOut() {
    save({ jamSelesai: nowHHMM() });
  }

  const isOff = entry?.shiftCode === "OFF";
  const hasShift = !!entry?.shiftCode && !isOff;
  const isLoggedIn = hasShift && !!entry?.jamMulai;
  const isDone = isLoggedIn && !!entry?.jamSelesai;
  const jam = isDone ? computeHours(entry.jamMulai, entry.jamSelesai) : null;

  return (
    <div className="min-h-screen px-4 py-6 flex flex-col">
      <div className="max-w-md mx-auto w-full flex flex-col gap-4 flex-1">
        <div className="flex items-center justify-between">
          <div>
            <h1 style={{ color: COLORS.ink }} className="text-lg font-semibold">{guard.name}</h1>
            <p style={{ color: COLORS.inkSoft }} className="text-xs">{guard.jabatan} · NPK {guard.npk} · {guard.projek}</p>
          </div>
        </div>

        {saving && <p style={{ color: COLORS.inkSoft }} className="text-xs text-right">{t("saving")}</p>}

        <div className="flex-1 flex flex-col items-center justify-center gap-5 py-6">
          <div className="text-center">
            <p style={{ color: COLORS.amberDeep }} className="text-xs font-semibold uppercase tracking-wide">{t("todayLabel")}</p>
            <p style={{ color: COLORS.ink }} className="text-xl font-semibold">{today.d} {monthShort(today.m, lang)}</p>
          </div>

          {isOff && (
            <div style={{ background: COLORS.card, borderColor: COLORS.line }} className="border rounded-2xl px-6 py-8 w-full text-center flex flex-col gap-3 items-center">
              <p style={{ color: COLORS.inkSoft }} className="text-base italic">{t("offToday")}</p>
              <button
                onClick={() => { save({ shiftCode: "" }); }}
                style={{ color: COLORS.inkSoft }}
                className="text-xs underline p-2"
              >
                {t("undo")}
              </button>
            </div>
          )}

          {!isOff && !hasShift && (
            <div className="w-full flex flex-col gap-3">
              <p style={{ color: COLORS.inkSoft }} className="text-sm text-center">{t("chooseShiftPrompt")}</p>
              <div className="grid grid-cols-4 gap-2">
                {SHIFT_ORDER.map((code) => (
                  <button
                    key={code}
                    onClick={() => pickShift(code)}
                    style={{
                      background: code === "OFF" ? "transparent" : COLORS.card,
                      borderColor: code === "OFF" ? COLORS.red : COLORS.line,
                      color: code === "OFF" ? COLORS.red : COLORS.ink,
                    }}
                    className="border rounded-xl py-3 text-sm font-bold min-h-[52px]"
                    title={SHIFT_CODES[code].label + (SHIFT_CODES[code].custom || code === "OFF" ? "" : ` ${SHIFT_CODES[code].mulai}–${SHIFT_CODES[code].selesai}`)}
                  >
                    {code}
                  </button>
                ))}
              </div>
              <p style={{ color: COLORS.inkSoft }} className="text-[11px] text-center">{t("oneLoginHint")}</p>
            </div>
          )}

          {hasShift && !isLoggedIn && (
            <div style={{ background: COLORS.card, borderColor: COLORS.line }} className="border rounded-2xl px-6 py-8 w-full text-center flex flex-col gap-4 items-center">
              <p style={{ color: COLORS.inkSoft }} className="text-xs">
                {t("shiftLabel")}: <span style={{ color: COLORS.amberDeep }} className="font-bold">{entry.shiftCode}</span>
              </p>
              <button
                onClick={logIn}
                style={{ background: COLORS.green, color: COLORS.ink }}
                className="font-bold rounded-xl px-8 py-4 text-lg font-bold w-full min-h-[56px]"
              >
                {t("logIn")}
              </button>
              <button onClick={() => save({ shiftCode: "" })} style={{ color: COLORS.inkSoft }} className="text-xs underline p-2">
                {t("changeShift")}
              </button>
              <p style={{ color: COLORS.inkSoft }} className="text-[11px]">{t("oneLoginHint")}</p>
            </div>
          )}

          {isLoggedIn && !isDone && (
            <div style={{ background: COLORS.card, borderColor: COLORS.line }} className="border rounded-2xl px-6 py-8 w-full text-center flex flex-col gap-4 items-center">
              <p style={{ color: COLORS.ink }} className="text-sm">
                <span style={{ color: COLORS.amberDeep }} className="font-bold">{entry.shiftCode}</span> — {t("loggedInAt", entry.jamMulai)}
              </p>
              <button
                onClick={logOut}
                style={{ background: COLORS.ink }}
                className="text-white rounded-xl px-8 py-4 text-lg font-bold w-full min-h-[56px]"
              >
                {t("logOut")}
              </button>
            </div>
          )}

          {isDone && (
            <div style={{ background: COLORS.card, borderColor: COLORS.line }} className="border rounded-2xl px-6 py-8 w-full text-center flex flex-col gap-2 items-center">
              <Check size={28} color={COLORS.green} />
              <p style={{ color: COLORS.ink }} className="text-sm font-medium">{t("allDoneToday")}</p>
              <p style={{ color: COLORS.inkSoft }} className="text-xs">
                <span style={{ color: COLORS.amberDeep }} className="font-bold">{entry.shiftCode}</span> · {t("doneSummary", entry.jamMulai, entry.jamSelesai, jam)}
              </p>
            </div>
          )}

          <NotesBox key={today.key} value={entry?.keterangan || ""} onSave={(v) => save({ keterangan: v })} />
        </div>

        <div className="flex flex-col items-center gap-2 pt-2">
          {boundDevice ? (
            <button onClick={onLogout} style={{ color: COLORS.inkSoft }} className="text-[11px] underline p-2">
              {t("resetDeviceLink")}
            </button>
          ) : (
            <button onClick={onLogout} style={{ color: COLORS.inkSoft }} className="text-xs flex items-center gap-1 p-2">
              <LogOut size={14} /> {t("switch")}
            </button>
          )}
          <div className="flex items-center gap-4">
            <button onClick={onDanru} style={{ color: COLORS.inkSoft }} className="text-[10px] p-2">{t("danruAccess")}</button>
            <button onClick={onManager} style={{ color: COLORS.inkSoft }} className="text-[10px] p-2">{t("managerAccess")}</button>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ---------- notes: always-open free-text section on the guard's front page ---------- */

function NotesBox({ value, onSave }) {
  const t = useT();
  const [text, setText] = useState(value);
  const [dirty, setDirty] = useState(false);

  return (
    <div style={{ background: COLORS.card, borderColor: COLORS.line }} className="border rounded-xl p-3 w-full flex flex-col gap-1.5">
      <span style={{ color: COLORS.inkSoft }} className="text-[10px] uppercase tracking-wide">{t("notesLabel")}</span>
      <textarea
        value={text}
        onChange={(e) => { setText(e.target.value); setDirty(true); }}
        onBlur={() => { if (dirty) { onSave(text); setDirty(false); } }}
        placeholder={t("notesPlaceholder")}
        rows={2}
        style={{ borderColor: COLORS.line, color: COLORS.ink }}
        className="border rounded-lg px-2.5 py-2 text-base sm:text-sm resize-none outline-none"
      />
    </div>
  );
}

/* ---------- schedule grid: Danru assigns/changes each guard's shift per day ---------- */

const SHIFT_CELL_COLORS = {
  P: "#F2C57C", S: "#8FB4D9", M: "#8578B8", P1: "#C7D98A", P2: "#D9B48F", HS: "#E9A6C0", OFF: "#E39696", "": "transparent",
};

function ScheduleGrid({ roster, schedule, onSetCell }) {
  const t = useT();
  const { lang } = useContext(LangContext);
  const [periodId, setPeriodId] = useState(periodIdForToday());
  const days = useMemo(() => periodDays(periodId), [periodId]);
  const periods = useMemo(() => recentPeriods(14), []);

  function handleChange(guardName, guardId, dk, dayLabel, newCode) {
    const existing = schedule[guardId]?.[dk] || "";
    if (existing && existing !== newCode) {
      const ok = window.confirm(t("duplicateShiftConfirm", guardName, dayLabel, existing, newCode));
      if (!ok) return;
    }
    onSetCell(guardId, dk, newCode);
  }

  return (
    <div className="flex flex-col gap-3">
      <div className="flex flex-wrap items-center justify-between gap-2">
        <p style={{ color: COLORS.inkSoft }} className="text-xs max-w-md">{t("scheduleTitle")}</p>
        <select
          value={periodId}
          onChange={(e) => setPeriodId(e.target.value)}
          style={{ borderColor: COLORS.line }}
          className="border rounded-lg px-3 py-2 text-base sm:text-sm min-h-[42px]"
        >
          {periods.map((p) => <option key={p} value={p}>{periodLabel(p, lang)}</option>)}
        </select>
      </div>

      <p style={{ color: COLORS.inkSoft }} className="text-[11px] sm:hidden">{t("swipeHint")}</p>

      <div className="overflow-x-auto border rounded-lg" style={{ borderColor: COLORS.line, WebkitOverflowScrolling: "touch" }}>
        <table className="text-xs border-collapse w-full">
          <thead>
            <tr>
              <th
                style={{ background: COLORS.card, borderColor: COLORS.line, left: 0 }}
                className="sticky border-b border-r px-2 py-1.5 text-left min-w-[120px] z-10"
              >
                {t("fieldName")}
              </th>
              {days.map((d) => (
                <th key={d.key} style={{ borderColor: COLORS.line, color: COLORS.inkSoft }} className="border-b px-1 py-1.5 font-normal min-w-[46px]">
                  {d.d}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {roster.map((g) => (
              <tr key={g.id}>
                <td
                  style={{ background: COLORS.card, borderColor: COLORS.line, left: 0 }}
                  className="sticky border-r px-2 py-1 font-medium whitespace-nowrap z-10"
                >
                  {g.name}
                </td>
                {days.map((d) => {
                  const val = schedule[g.id]?.[d.key] || "";
                  const dayLabel = `${dayName(d, lang)}, ${d.d} ${monthShort(d.m, lang)}`;
                  return (
                    <td key={d.key} className="p-0.5">
                      <select
                        value={val}
                        onChange={(e) => handleChange(g.name, g.id, d.key, dayLabel, e.target.value)}
                        style={{ background: SHIFT_CELL_COLORS[val] || "transparent", borderColor: COLORS.line, color: val ? "#1a1a1a" : COLORS.inkSoft }}
                        className="border rounded text-[11px] w-full py-2 text-center font-semibold min-h-[36px]"
                      >
                        <option value=""></option>
                        {SHIFT_ORDER.map((code) => <option key={code} value={code}>{code}</option>)}
                      </select>
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ---------- danru view: roster + schedule + print access ---------- */

function DanruView({ roster, activeRoster, entries, settings, schedule, onAddGuard, onSoftDeleteGuard, onUpdateGuard, onImportRoster, onSetScheduleCell, onLogout, saving }) {
  const t = useT();
  const [tab, setTab] = useState("roster"); // roster | docs
  const { lang } = useContext(LangContext);
  return (
    <div className="min-h-screen px-4 lg:px-8 py-4 max-w-[1700px] mx-auto flex flex-col gap-4">
      <div>
      <div className="flex items-end justify-between gap-3">
        <h1 className="eco-title">{t("danruHeader")}</h1>
        <div className="flex items-center gap-3 pb-3">
          {saving && <span style={{ color: COLORS.inkSoft }} className="text-xs">{t("saving")}</span>}
          <button onClick={onLogout} style={{ color: COLORS.inkSoft }} className="text-xs flex items-center gap-1 p-2 -m-2"><LogOut size={14} /> {t("exit")}</button>
        </div>
      </div>
      <div className="eco-rule" />
      </div>

      <div className="flex flex-wrap gap-x-5 no-print">
        {[["roster", t("tabRoster")], ["docs", pick(lang, "Daftar dokumen", "Documents checklist", "Senarai semak dokumen")]].map(([id, label]) => (
          <button
            key={id}
            onClick={() => setTab(id)}
            className={"eco-tab" + (tab === id ? " on" : "")}
          >
            {label}
          </button>
        ))}
      </div>

      {tab === "roster" && (
        <RosterTab roster={roster} role="danru" settings={settings} schedule={schedule} onAdd={onAddGuard} onSoftDelete={onSoftDeleteGuard} onUpdate={onUpdateGuard} onImportRoster={onImportRoster} />
      )}

      {tab === "docs" && <DocsOverview roster={activeRoster} />}
    </div>
  );
}

/* ---------- manager view ---------- */

function ManagerView({ roster, activeRoster, entries, settings, schedule, onAddGuard, onSoftDeleteGuard, onRestoreGuard, onUpdateGuard, onImportRoster, onSetScheduleCell, onUpdateSettings, onLogout, saving }) {
  const t = useT();
  const [tab, setTab] = useState("roster"); // roster | schedule | settings (Print / Verify hidden: timesheets live in the Patroli app)

  return (
    <div className="min-h-screen">
      <div className="no-print px-4 lg:px-8 py-4 max-w-[1700px] mx-auto flex flex-col gap-4">
        <div>
        <div className="flex items-end justify-between gap-3">
          <h1 className="eco-title">{t("manager")}</h1>
          <div className="flex items-center gap-3 pb-3">
            {saving && <span style={{ color: COLORS.inkSoft }} className="text-xs">{t("saving")}</span>}
            <button onClick={onLogout} style={{ color: COLORS.inkSoft }} className="text-xs flex items-center gap-1 p-2 -m-2"><LogOut size={14} /> {t("exit")}</button>
          </div>
        </div>
        <div className="eco-rule" />
        </div>

        <div className="flex flex-wrap gap-x-5">
          {[["roster", t("tabRoster")], ["settings", t("tabSettings")]].map(([id, label]) => (
            <button
              key={id}
              onClick={() => setTab(id)}
              className={"eco-tab" + (tab === id ? " on" : "")}
            >
              {label}
            </button>
          ))}
        </div>

        {tab === "roster" && (
          <RosterTab roster={roster} role="manager" settings={settings} schedule={schedule} onAdd={onAddGuard} onSoftDelete={onSoftDeleteGuard} onRestore={onRestoreGuard} onUpdate={onUpdateGuard} onImportRoster={onImportRoster} />
        )}

        {tab === "settings" && <SettingsTab settings={settings} onUpdate={onUpdateSettings} />}

      </div>
    </div>
  );
}

function LocationLog({ entries, periodId }) {
  const t = useT();
  const { lang } = useContext(LangContext);
  const days = periodDays(periodId).filter((d) => entries[d.key]?.location?.lat != null);

  if (days.length === 0) {
    return <p style={{ color: COLORS.inkSoft }} className="text-sm">{t("noLocationYet")}</p>;
  }

  return (
    <div style={{ background: COLORS.card, borderColor: COLORS.line }} className="border rounded-xl p-4">
      <p style={{ color: COLORS.ink }} className="text-sm font-medium flex items-center gap-1.5 mb-2"><MapPin size={14} color={COLORS.amberDeep} /> {t("locationLog")}</p>
      <div className="flex flex-col gap-1.5">
        {days.map((day) => {
          const loc = entries[day.key].location;
          return (
            <div key={day.key} className="flex flex-wrap items-center justify-between gap-x-2 gap-y-0.5 text-xs">
              <span style={{ color: COLORS.inkSoft }}>{dayName(day, lang)}, {day.d} {monthShort(day.m, lang)} — {loc.lat}, {loc.lng} {loc.accuracy != null && `(±${loc.accuracy}m)`}</span>
              <div className="flex items-center gap-2">
                {loc.withinRadius === true && (
                  <span style={{ color: COLORS.green }} className="flex items-center gap-0.5"><Check size={11} /> {t("withinRadius", loc.distanceMeters)}</span>
                )}
                {loc.withinRadius === false && (
                  <span style={{ color: COLORS.red }} className="flex items-center gap-0.5"><AlertTriangle size={11} /> {t("outsideRadius", loc.distanceMeters)}</span>
                )}
                <a
                  href={`https://maps.google.com/?q=${loc.lat},${loc.lng}`}
                  target="_blank"
                  rel="noopener noreferrer"
                  style={{ color: COLORS.amberDeep }}
                  className="underline"
                >
                  {t("viewMap")}
                </a>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function GuardRow({ guard, selected, onClick }) {
  const { lang } = useContext(LangContext);
  const [imgOk, setImgOk] = useState(true);
  const n = guardDocIssues(guard).length;
  const soonN = guardDocSoon(guard).length;
  const initials = String(guard.name || "?").split(/\s+/).filter(Boolean).slice(0, 2).map((w) => w[0]).join("").toUpperCase();
  return (
    <button
      onClick={onClick}
      style={{ borderColor: selected ? COLORS.green : COLORS.line, background: selected ? "#F1F8F4" : COLORS.card, borderLeftWidth: 4 }}
      className="eco-flat border rounded-lg px-3 py-2 flex items-center gap-3 text-left w-full flex-shrink-0 min-h-[60px]"
    >
      <span style={{ background: "#F1F1F1", color: COLORS.inkSoft }} className="w-10 h-10 rounded-full overflow-hidden flex items-center justify-center text-xs font-medium flex-shrink-0">
        {guard.photoFileId && imgOk ? <img loading="lazy" src={driveThumb(guard.photoFileId, 96)} alt="" onError={() => setImgOk(false)} className="w-full h-full object-cover" /> : initials}
      </span>
      <span className="min-w-0 flex-1">
        <span style={{ color: COLORS.ink }} className="block text-sm font-medium truncate">{guard.name}</span>
        <span style={{ color: COLORS.inkSoft }} className="block text-xs truncate">{[guard.jabatan, guard.company].filter(Boolean).join(" · ") || "—"}</span>
      </span>
      <span style={{ borderColor: COLORS.line, color: COLORS.ink }} className="border text-[11px] font-medium rounded-full px-2 py-0.5 whitespace-nowrap inline-flex items-center gap-1.5 flex-shrink-0">
        <i style={{ background: n > 0 || soonN > 0 ? COLORS.amber : COLORS.green }} className="inline-block w-2 h-2 rounded-full" />
        {n > 0 ? pick(lang, `${n} kurang`, `${n} missing`, `${n} kurang`) : soonN > 0 ? pick(lang, `${soonN} segera habis`, `${soonN} expiring`, `${soonN} hampir tamat`) : pick(lang, "Lengkap", "Complete", "Lengkap")}
      </span>
    </button>
  );
}

function RosterTab({ roster, role, settings, schedule, onAdd, onSoftDelete, onRestore, onUpdate, onImportRoster }) {
  const t = useT();
  const [form, setForm] = useState({ name: "", npk: "", jabatan: "", projek: "", company: "", phone: "", regu: "" });
  const [importMsg, setImportMsg] = useState("");
  const [pendingDelete, setPendingDelete] = useState(null); // guard object being considered for deletion
  const [showDeleted, setShowDeleted] = useState(false);

  function submit() {
    if (!form.name.trim()) return;
    onAdd({ ...form });
    setForm({ name: "", npk: "", jabatan: "", projek: "", company: "", phone: "", regu: "" });
  }

  async function handleImport() {
    const added = await onImportRoster();
    setImportMsg(added > 0 ? t("importAdded", added) : t("importNone"));
    setTimeout(() => setImportMsg(""), 4000);
  }

  const activeGuards = roster.filter((g) => !g.deleted);
  const deletedGuards = roster.filter((g) => g.deleted);
  const [selectedId, setSelectedId] = useState(null);
  const [showAdd, setShowAdd] = useState(false);
  const [query, setQuery] = useState("");
  const { lang } = useContext(LangContext);
  const q = query.trim().toLowerCase();
  const filtered = activeGuards.filter((g) => !q || [g.name, g.npk, g.jabatan, g.company].some((v) => String(v || "").toLowerCase().includes(q)));
  const selected = activeGuards.find((g) => g.id === selectedId) || null;
  const detailOpen = !!selected || showAdd;

  return (
    <div className="flex flex-col gap-4">
      <div className="lg:grid lg:grid-cols-[340px_minmax(0,1fr)] lg:gap-6 lg:h-[calc(100vh-250px)] lg:min-h-[520px]">
        <aside className={(detailOpen ? "hidden " : "flex ") + "lg:flex flex-col gap-3 min-h-0"}>
          <div className="flex gap-2">
            <label style={{ borderColor: COLORS.line }} className="border rounded-lg flex items-center gap-2 px-3 flex-1 min-w-0">
              <Search size={15} color={COLORS.inkSoft} />
              <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder={pick(lang, "Cari nama / NPK", "Search name / NPK", "Cari nama / NPK")} className="py-2.5 text-sm outline-none w-full min-w-0 bg-transparent" />
            </label>
            <button onClick={() => { setShowAdd(true); setSelectedId(null); }} style={{ background: COLORS.green, color: COLORS.ink }} className="rounded-lg px-3 font-bold text-sm flex items-center gap-1.5 min-h-[44px] flex-shrink-0">
              <Plus size={14} /> {t("add")}
            </button>
          </div>
          <p style={{ color: COLORS.inkSoft }} className="text-xs">{t("totalGuards", activeGuards.length)}</p>
          <div className="flex flex-col gap-1.5 lg:overflow-y-auto min-h-0 flex-1 lg:pr-1">
            {filtered.map((g) => (
              <GuardRow key={g.id} guard={g} selected={g.id === selectedId && !showAdd} onClick={() => { setSelectedId(g.id); setShowAdd(false); }} />
            ))}
            {activeGuards.length === 0 && <div className="flex flex-col items-center gap-3 py-8"><DotGrid size={56} /><p style={{ color: COLORS.inkSoft }} className="text-sm text-center uppercase tracking-widest font-light">{t("noPersonnelYet")}</p></div>}
          </div>
        </aside>

        <section className={(detailOpen ? "block " : "hidden ") + "lg:block lg:overflow-y-auto min-h-0 lg:pr-1"}>
          <button onClick={() => { setSelectedId(null); setShowAdd(false); }} style={{ color: COLORS.ink }} className="lg:hidden flex items-center gap-1.5 text-sm font-medium mb-3 min-h-[44px]">
            <ArrowLeft size={16} /> {pick(lang, "Kembali ke daftar", "Back to list", "Kembali ke senarai")}
          </button>
          {showAdd ? (
            <div className="flex flex-col gap-4 max-w-[780px]">
      <div style={{ background: COLORS.card, borderColor: COLORS.line }} className="border rounded-xl p-4 flex flex-col gap-2">
        <div className="flex items-center justify-between">
          <p style={{ color: COLORS.ink }} className="text-sm font-medium">{t("importTitle")}</p>
          <button
            onClick={handleImport}
            style={{ borderColor: COLORS.amberDeep, color: COLORS.amberDeep }}
            className="border rounded-lg px-3 py-1 text-xs font-medium"
          >
            {t("importButton")}
          </button>
        </div>
        {importMsg && <p style={{ color: COLORS.green }} className="text-xs">{importMsg}</p>}
      </div>

      <div style={{ background: COLORS.card, borderColor: COLORS.line }} className="border rounded-xl p-4 flex flex-col gap-2">
        <p style={{ color: COLORS.ink }} className="text-sm font-medium flex items-center gap-1.5"><UserPlus size={14} /> {t("addManualTitle")}</p>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
          <input placeholder={t("fieldName")} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} style={{ borderColor: COLORS.line }} className="border rounded px-2 py-2 text-base sm:text-sm min-h-[42px]" />
          <input placeholder={t("fieldNpk")} value={form.npk} onChange={(e) => setForm({ ...form, npk: e.target.value })} style={{ borderColor: COLORS.line }} className="border rounded px-2 py-2 text-base sm:text-sm min-h-[42px]" />
          <input placeholder={t("fieldJabatan")} value={form.jabatan} onChange={(e) => setForm({ ...form, jabatan: e.target.value })} style={{ borderColor: COLORS.line }} className="border rounded px-2 py-2 text-base sm:text-sm min-h-[42px]" />
          <input placeholder={t("fieldProjek")} value={form.projek} onChange={(e) => setForm({ ...form, projek: e.target.value })} style={{ borderColor: COLORS.line }} className="border rounded px-2 py-2 text-base sm:text-sm min-h-[42px]" />
          <input placeholder={t("fieldCompany")} value={form.company} onChange={(e) => setForm({ ...form, company: e.target.value })} style={{ borderColor: COLORS.line }} className="border rounded px-2 py-2 text-base sm:text-sm min-h-[42px]" />
          <input placeholder={t("fieldPhone")} value={form.phone} onChange={(e) => setForm({ ...form, phone: e.target.value })} style={{ borderColor: COLORS.line }} className="border rounded px-2 py-2 text-base sm:text-sm min-h-[42px]" />
          <input placeholder={t("fieldRegu")} value={form.regu} onChange={(e) => setForm({ ...form, regu: e.target.value })} style={{ borderColor: COLORS.line }} className="border rounded px-2 py-2 text-base sm:text-sm min-h-[42px]" />
        </div>
        <button onClick={submit} style={{ background: COLORS.green, color: COLORS.ink }} className="font-bold rounded-lg py-2.5 text-sm font-medium flex items-center justify-center gap-1.5 min-h-[44px]">
          <Plus size={14} /> {t("add")}
        </button>
      </div>

            </div>
          ) : selected ? (
            <GuardCard key={selected.id} guard={selected} onUpdate={onUpdate} onRequestDelete={() => setPendingDelete(selected)} schedule={schedule} role={role} />
          ) : (
            <div className="hidden lg:flex flex-col items-center justify-center gap-3 h-full min-h-[320px]">
              <DotGrid size={56} />
              <p style={{ color: COLORS.inkSoft }} className="text-sm uppercase tracking-widest font-light text-center">{pick(lang, "Pilih nama untuk melihat biodata", "Select a name to see the biodata", "Pilih nama untuk melihat biodata")}</p>
            </div>
          )}
        </section>
      </div>

      <div className="border-t pt-3" style={{ borderColor: COLORS.line }}>
        <button
          onClick={() => setShowDeleted((v) => !v)}
          style={{ color: COLORS.inkSoft }}
          className="text-xs font-medium flex items-center gap-1.5 p-2 -m-2"
        >
          {showDeleted ? <ChevronUp size={13} /> : <ChevronDown size={13} />}
          {t("deletedSection")} ({deletedGuards.length})
        </button>

        {showDeleted && (
          <div className="flex flex-col gap-2 mt-2">
            {deletedGuards.length === 0 && <p style={{ color: COLORS.inkSoft }} className="text-xs">{t("noDeletedGuards")}</p>}
            {deletedGuards.map((g) => (
              <div key={g.id} style={{ background: COLORS.paper, borderColor: COLORS.line }} className="border rounded-xl p-3 flex items-center justify-between gap-2">
                <div className="min-w-0">
                  <p style={{ color: COLORS.ink }} className="text-sm font-medium truncate">{g.name}</p>
                  <p style={{ color: COLORS.inkSoft }} className="text-xs">
                    {g.deletedReason} {g.deletedBy && `· ${t("deletedByLabel", g.deletedBy)}`}
                    {g.deletedAt && ` · ${t("deletedOn", new Date(g.deletedAt).toLocaleDateString())}`}
                  </p>
                </div>
                {role === "manager" && onRestore && (
                  <button
                    onClick={() => onRestore(g.id)}
                    style={{ borderColor: COLORS.green, color: COLORS.green }}
                    className="border rounded-lg px-3 py-1.5 text-xs font-medium flex-shrink-0"
                  >
                    {t("restoreGuardBtn")}
                  </button>
                )}
              </div>
            ))}
          </div>
        )}
      </div>

      {pendingDelete && (
        <DeleteGuardModal
          guard={pendingDelete}
          role={role}
          settings={settings}
          onCancel={() => setPendingDelete(null)}
          onConfirm={async (reason) => {
            await onSoftDelete(pendingDelete.id, reason, role);
            setSelectedId(null);
            setPendingDelete(null);
          }}
        />
      )}
    </div>
  );
}

function DeleteGuardModal({ guard, role, settings, onCancel, onConfirm }) {
  const t = useT();
  const [step, setStep] = useState("reason"); // reason | pin
  const [reason, setReason] = useState("");
  const [otherText, setOtherText] = useState("");
  const [pin, setPin] = useState("");
  const [pinError, setPinError] = useState("");

  const finalReason = reason === "other" ? (otherText.trim() || t("reasonOther")) : (
    reason === "resigned" ? t("reasonResigned") : reason === "terminated" ? t("reasonTerminated") : ""
  );
  const canContinue = reason && (reason !== "other" || otherText.trim());

  function handleContinue() {
    if (!canContinue) return;
    if (role === "danru") {
      setStep("pin");
    } else {
      onConfirm(finalReason);
    }
  }

  function handlePinSubmit() {
    if (pin === settings.pin) {
      onConfirm(finalReason);
    } else {
      setPinError(t("incorrectPin"));
      setPin("");
    }
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center px-4" style={{ background: "rgba(24,36,55,0.55)" }}>
      <div style={{ background: COLORS.card }} className="rounded-xl p-5 w-full max-w-sm flex flex-col gap-3">
        {step === "reason" && (
          <>
            <h3 style={{ color: COLORS.ink }} className="font-semibold">{t("deleteConfirmTitle", guard.name)}</h3>
            <p style={{ color: COLORS.inkSoft }} className="text-xs">{t("deleteConfirmBody")}</p>
            <label className="flex flex-col gap-1 mt-1">
              <span style={{ color: COLORS.inkSoft }} className="text-xs">{t("deleteReasonLabel")}</span>
              <select
                value={reason}
                onChange={(e) => setReason(e.target.value)}
                style={{ borderColor: COLORS.line }}
                className="border rounded-lg px-3 py-2 text-base sm:text-sm min-h-[42px]"
              >
                <option value="" disabled>{t("reasonSelect")}</option>
                <option value="resigned">{t("reasonResigned")}</option>
                <option value="terminated">{t("reasonTerminated")}</option>
                <option value="other">{t("reasonOther")}</option>
              </select>
            </label>
            {reason === "other" && (
              <input
                value={otherText}
                onChange={(e) => setOtherText(e.target.value)}
                placeholder={t("reasonOtherPlaceholder")}
                style={{ borderColor: COLORS.line }}
                className="border rounded-lg px-3 py-2 text-base sm:text-sm min-h-[42px]"
              />
            )}
            <div className="flex gap-2 mt-2">
              <button onClick={onCancel} style={{ borderColor: COLORS.line, color: COLORS.inkSoft }} className="border rounded-lg py-2.5 text-sm font-medium flex-1 min-h-[44px]">
                {t("cancel")}
              </button>
              <button
                onClick={handleContinue}
                disabled={!canContinue}
                style={{ background: COLORS.ink }}
                className="text-white rounded-lg py-2.5 text-sm font-medium flex-1 min-h-[44px] disabled:opacity-40"
              >
                {role === "danru" ? t("continueBtn") : t("confirmDeleteBtn")}
              </button>
            </div>
          </>
        )}

        {step === "pin" && (
          <>
            <h3 style={{ color: COLORS.ink }} className="font-semibold">{t("deletePinPrompt")}</h3>
            <input
              type="password"
              inputMode="numeric"
              value={pin}
              onChange={(e) => setPin(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && handlePinSubmit()}
              style={{ borderColor: COLORS.line }}
              className="border rounded-lg px-3 py-2 text-center tracking-widest text-lg"
              autoFocus
            />
            {pinError && <p style={{ color: COLORS.red }} className="text-xs text-center">{pinError}</p>}
            <div className="flex gap-2 mt-2">
              <button onClick={onCancel} style={{ borderColor: COLORS.line, color: COLORS.inkSoft }} className="border rounded-lg py-2.5 text-sm font-medium flex-1 min-h-[44px]">
                {t("cancel")}
              </button>
              <button onClick={handlePinSubmit} style={{ background: COLORS.ink }} className="text-white rounded-lg py-2.5 text-sm font-medium flex-1 min-h-[44px]">
                {t("confirmDeleteBtn")}
              </button>
            </div>
          </>
        )}
      </div>
    </div>
  );
}

const driveThumb = (id, w = 200) => `https://drive.google.com/thumbnail?id=${id}&sz=w${w}`;
const photoCache = new Map(); // guardId -> Drive file id, remembered for this session

// Peek window for a stored file. Download is a button inside it, only when needed.
function FilePreview({ storageKey, fileId, fileName, onClose }) {
  const { lang } = useContext(LangContext);
  const [data, setData] = useState(null); // stored metadata: older records / local mode
  const [loading, setLoading] = useState(!fileId);
  useEffect(() => {
    if (fileId || !storageKey) return;
    let cancelled = false;
    loadShared(storageKey, null).then((d) => { if (!cancelled) { setData(d); setLoading(false); } });
    return () => { cancelled = true; };
  }, [storageKey, fileId]);
  useEffect(() => {
    const onKey = (e) => { if (e.key === "Escape") onClose(); };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [onClose]);
  const id = fileId || data?.driveFileId;
  function doDownload() {
    if (data) downloadStoredFile(data);
    else if (id) window.open(`https://drive.google.com/uc?export=download&id=${id}`, "_blank", "noopener");
  }
  let body;
  if (loading) body = <p style={{ color: COLORS.inkSoft }} className="text-sm">{pick(lang, "Memuat…", "Loading…", "Memuatkan…")}</p>;
  else if (id) body = <iframe title={fileName || "preview"} src={`https://drive.google.com/file/d/${id}/preview`} className="w-full h-full border-0 bg-white" />;
  else if (data?.base64 && data.mimeType?.startsWith("image/")) body = <img alt="" src={`data:${data.mimeType};base64,${data.base64}`} className="max-w-full max-h-full object-contain" />;
  else if (data?.base64 && data.mimeType === "application/pdf") body = <iframe title={fileName || "preview"} src={`data:application/pdf;base64,${data.base64}`} className="w-full h-full border-0 bg-white" />;
  else body = <p style={{ color: COLORS.inkSoft }} className="text-sm px-4 text-center">{pick(lang, "Pratinjau tidak tersedia. Gunakan tombol Unduh.", "No preview available. Use the Download button.", "Pratonton tidak tersedia. Gunakan butang Muat turun.")}</p>;
  return (
    <div className="no-print fixed inset-0 z-[70] flex items-center justify-center p-2 sm:p-6" style={{ background: "rgba(35,31,32,.6)" }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: COLORS.card }} className="w-full max-w-5xl h-[90vh] rounded-lg flex flex-col overflow-hidden">
        <div style={{ borderColor: COLORS.line }} className="flex items-center gap-2 px-3 py-2 border-b">
          <span style={{ color: COLORS.ink }} className="text-sm font-medium truncate flex-1 min-w-0">{fileName}</span>
          <button onClick={doDownload} style={{ background: COLORS.green, color: COLORS.ink }} className="rounded-lg px-3 text-xs font-bold flex items-center gap-1.5 min-h-[40px]">
            <Download size={14} /> {pick(lang, "Unduh", "Download", "Muat turun")}
          </button>
          <button onClick={onClose} aria-label="Close" style={{ color: COLORS.ink }} className="eco-flat p-2"><X size={18} /></button>
        </div>
        <div style={{ background: "#F4F4F4" }} className="flex-1 min-h-0 flex items-center justify-center">{body}</div>
      </div>
    </div>
  );
}

function FileAttachment({ storageKey, fileName, fileSize, fileId, onUploaded, onRemoved, accept, folderName }) {
  const t = useT();
  const { lang } = useContext(LangContext);
  const [uploading, setUploading] = useState(false);
  const [error, setError] = useState("");
  const [peek, setPeek] = useState(false);

  async function handleChange(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setError("");
    setUploading(true);
    const res = await uploadFileToStorage(file, storageKey, folderName);
    setUploading(false);
    if (res.ok) {
      onUploaded({ fileName: res.fileName || file.name, size: res.size ?? file.size, fileId: res.fileId || null });
    } else if (res.error === "too-large") {
      setError(t("fileTooLarge"));
    }
    e.target.value = "";
  }

  async function handleRemove() {
    await deleteFileFromStorage(storageKey);
    onRemoved();
  }

  if (fileName) {
    return (
      <div style={{ borderColor: COLORS.line }} className="border rounded-lg px-2.5 py-1.5 flex items-center gap-2 text-xs">
        <button onClick={() => setPeek(true)} title={pick(lang, "Lihat", "Preview", "Pratonton")} style={{ color: COLORS.ink }} className="eco-flat flex items-center gap-2 flex-1 min-w-0 text-left py-1 hover:underline">
          <Eye size={13} color={COLORS.green} className="flex-shrink-0" />
          <span className="truncate">{fileName}{fileSize != null && <span style={{ color: COLORS.inkSoft }}> ({formatFileSize(fileSize)})</span>}</span>
        </button>
        <button onClick={handleRemove} style={{ color: COLORS.red }} className="eco-flat flex-shrink-0 p-1"><X size={13} /></button>
        {peek && <FilePreview storageKey={storageKey} fileId={fileId} fileName={fileName} onClose={() => setPeek(false)} />}
      </div>
    );
  }

  return (
    <div className="flex flex-col gap-1">
      <label style={{ borderColor: COLORS.line, color: COLORS.inkSoft }} className="border rounded-lg px-2.5 py-2 text-xs flex items-center justify-center gap-1.5 cursor-pointer min-h-[38px]">
        <Paperclip size={12} />
        {uploading ? t("uploading") : t("uploadFile")}
        <input type="file" accept={accept} onChange={handleChange} className="hidden" disabled={uploading} />
      </label>
      {error && <p style={{ color: COLORS.red }} className="text-[10px]">{error}</p>}
    </div>
  );
}

// Photo shows straight from the Drive file id kept on the guard record (no extra server call),
// and appears at once after choosing a file. Click it for a large preview.
function GuardPhoto({ guardId, hasPhoto, photoFileId, onChanged, folderName }) {
  const t = useT();
  const [fileId, setFileId] = useState(photoFileId || photoCache.get(guardId) || "");
  const [legacyUrl, setLegacyUrl] = useState(null); // base64 from local mode
  const [localUrl, setLocalUrl] = useState(null);   // instant preview of the file just chosen
  const [failed, setFailed] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [error, setError] = useState("");
  const [peek, setPeek] = useState(false);

  useEffect(() => {
    setFileId(photoFileId || photoCache.get(guardId) || "");
    setLegacyUrl(null); setLocalUrl(null); setFailed(false);
  }, [guardId, photoFileId]);

  // Older records have no photoFileId: look it up once, remember it, and save it on the guard.
  useEffect(() => {
    if (!hasPhoto || fileId || legacyUrl) return;
    let cancelled = false;
    loadShared(`photo:${guardId}`, null).then((data) => {
      if (cancelled || !data) return;
      if (data.driveFileId) { photoCache.set(guardId, data.driveFileId); setFileId(data.driveFileId); onChanged(true, data.driveFileId); }
      else if (data.base64) setLegacyUrl(`data:${data.mimeType};base64,${data.base64}`);
    });
    return () => { cancelled = true; };
  }, [guardId, hasPhoto, fileId, legacyUrl]);

  async function handleChange(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setError("");
    setLocalUrl(URL.createObjectURL(file));
    setFailed(false);
    setUploading(true);
    const res = await uploadFileToStorage(file, `photo:${guardId}`, folderName);
    setUploading(false);
    if (res.ok) {
      if (res.fileId) { photoCache.set(guardId, res.fileId); setFileId(res.fileId); }
      onChanged(true, res.fileId || null);
    } else {
      setLocalUrl(null);
      setError(res.error === "too-large" ? t("fileTooLarge") : "");
    }
    e.target.value = "";
  }

  const src = localUrl || (fileId ? driveThumb(fileId, 300) : legacyUrl);
  const hasImg = !!src && !failed;
  return (
    <div className="flex items-center gap-3">
      <button
        type="button"
        onClick={() => { if (fileId) setPeek(true); }}
        style={{ background: COLORS.paper, borderColor: COLORS.line, cursor: fileId ? "zoom-in" : "default" }}
        className="eco-flat w-24 h-24 rounded-xl border overflow-hidden flex items-center justify-center flex-shrink-0"
      >
        {hasImg ? <img src={src} alt="" className="w-full h-full object-cover" onError={() => setFailed(true)} /> : <Camera size={22} color={COLORS.inkSoft} />}
      </button>
      <div className="flex flex-col gap-1">
        <label style={{ borderColor: COLORS.ink, color: COLORS.ink }} className="border rounded-lg px-3 py-1.5 text-xs font-medium flex items-center gap-1.5 cursor-pointer w-fit min-h-[36px]">
          <Camera size={12} />
          {uploading ? t("uploading") : (hasImg || hasPhoto ? t("changePhoto") : t("uploadPhoto"))}
          <input type="file" accept="image/*" onChange={handleChange} className="hidden" disabled={uploading} />
        </label>
        {!hasImg && !hasPhoto && !uploading && <p style={{ color: COLORS.inkSoft }} className="text-[10px]">{t("noPhoto")}</p>}
        {error && <p style={{ color: COLORS.red }} className="text-[10px]">{error}</p>}
      </div>
      {peek && fileId && <FilePreview fileId={fileId} fileName="Photo" onClose={() => setPeek(false)} />}
    </div>
  );
}

function GuardCard({ guard, onUpdate, onRequestDelete, schedule, role }) {
  const t = useT();
  const [newCert, setNewCert] = useState({ name: "", number: "", issued: "", expiry: "", notes: "" });
  const [newCertFile, setNewCertFile] = useState(null);
  const [certBusy, setCertBusy] = useState(false);
  const [certErr, setCertErr] = useState("");
  const { lang } = useContext(LangContext);
  const issues = guardDocIssues(guard);
  const soonDocs = guardDocSoon(guard);
  const certificates = guard.certificates || [];
  const bio = guard.bio || {};
  const today = new Date().toISOString().slice(0, 10);
  const in30 = warnDate();
  const hasExpired = certificates.some((c) => c.expiry && c.expiry < today);
  const hasExpiringSoon = !hasExpired && certificates.some((c) => c.expiry && c.expiry >= today && c.expiry <= in30);

  function updateBio(field, value) {
    onUpdate(guard.id, { bio: { ...bio, [field]: value } });
  }

  function updateCert(certId, patch) {
    onUpdate(guard.id, { certificates: certificates.map((c) => (c.id === certId ? { ...c, ...patch } : c)) });
  }

  function updateSkck(field, value) {
    onUpdate(guard.id, { skck: { ...(guard.skck || {}), [field]: value } });
  }

  const warnings = guard.warnings || [];
  const [newWarning, setNewWarning] = useState({ reason: "", date: new Date().toISOString().slice(0, 10) });

  function addWarning() {
    if (!newWarning.reason.trim()) return;
    onUpdate(guard.id, { warnings: [{ ...newWarning, id: uid(), issuedBy: role }, ...warnings] });
    setNewWarning({ reason: "", date: new Date().toISOString().slice(0, 10) });
  }

  function removeWarning(id) {
    onUpdate(guard.id, { warnings: warnings.filter((w) => w.id !== id) });
  }

  const now = new Date();

  // Pick the file here; on "Add" it is uploaded straight to the guard's Drive folder.
  async function addCertificate() {
    if (!newCert.name.trim() || certBusy) return;
    setCertErr("");
    const id = uid();
    const cert = { ...newCert, id };
    if (newCertFile) {
      setCertBusy(true);
      const res = await uploadFileToStorage(newCertFile, `certfile:${id}`, guard.name);
      setCertBusy(false);
      if (!res.ok) { setCertErr(res.error === "too-large" ? t("fileTooLarge") : pick(lang, "Unggah gagal, coba lagi.", "Upload failed, try again.", "Muat naik gagal, cuba lagi.")); return; }
      cert.fileName = res.fileName || newCertFile.name;
      cert.fileSize = res.size ?? newCertFile.size;
      cert.fileId = res.fileId || null;
    }
    onUpdate(guard.id, { certificates: [...certificates, cert] });
    setNewCert({ name: "", number: "", issued: "", expiry: "", notes: "" });
    setNewCertFile(null);
  }

  function removeCertificate(certId) {
    onUpdate(guard.id, { certificates: certificates.filter((c) => c.id !== certId) });
  }

  return (
    <div style={{ background: COLORS.card, borderColor: COLORS.line }} className="border rounded-xl p-4 lg:p-6 flex flex-col gap-5">
      <div className="flex items-start justify-between gap-3">
        <div className="min-w-0">
          <h2 style={{ color: COLORS.ink }} className="text-xl font-medium truncate">{guard.name}</h2>
          <p style={{ color: COLORS.inkSoft }} className="text-sm">{[guard.jabatan, guard.company, guard.regu].filter(Boolean).join(" · ")}</p>
          <div className="flex items-center gap-2 mt-2 flex-wrap">
            <span style={{ borderColor: COLORS.line, color: COLORS.ink }} className="border text-[11px] font-medium rounded-full px-2 py-0.5 whitespace-nowrap inline-flex items-center gap-1.5">
              <i style={{ background: issues.length > 0 ? COLORS.amber : COLORS.green }} className="inline-block w-2 h-2 rounded-full" />
              {issues.length > 0 ? pick(lang, `${issues.length} kurang`, `${issues.length} missing`, `${issues.length} kurang`) : pick(lang, "Lengkap", "Complete", "Lengkap")}
            </span>
            {soonDocs.length > 0 && (
              <span style={{ color: COLORS.ink }} className="text-xs flex items-center gap-1"><AlertTriangle size={13} color={COLORS.ink} /> {soonDocs.map((x) => pick(lang, x.req.id, x.req.en, x.req.ms).split(" (")[0]).join(", ")} {pick(lang, "segera kadaluarsa (<3 bulan)", "expiring within 3 months", "hampir tamat (<3 bulan)")}</span>
            )}
            {(hasExpired || hasExpiringSoon) && (
              <span style={{ color: COLORS.ink }} className="text-xs flex items-center gap-1"><AlertTriangle size={13} color={hasExpired ? COLORS.red : COLORS.ink} /> {hasExpired ? t("certExpired") : t("certExpiringSoon")}</span>
            )}
          </div>
        </div>
        <button onClick={onRequestDelete} title={t("confirmDeleteBtn")} style={{ color: COLORS.red }} className="eco-flat p-2 -m-1 flex-shrink-0"><Trash2 size={18} /></button>
      </div>

      <div className="grid grid-cols-1 xl:grid-cols-2 gap-x-10 gap-y-5 items-start">
        <div className="flex flex-col gap-4 min-w-0">
          <div className="flex items-start justify-between gap-3 flex-wrap">
            <GuardPhoto guardId={guard.id} folderName={guard.name} hasPhoto={!!guard.hasPhoto} photoFileId={guard.photoFileId || ""} onChanged={(v, fid) => onUpdate(guard.id, fid ? { hasPhoto: v, photoFileId: fid } : { hasPhoto: v })} />
          </div>

          <div>
            <p style={{ color: COLORS.inkSoft }} className="text-[10px] uppercase tracking-wide mb-1">{t("basicInfoTitle")}</p>
            <div className="grid grid-cols-2 sm:grid-cols-3 gap-x-3 gap-y-1 text-xs">
              <EditableField value={guard.name} onChange={(v) => onUpdate(guard.id, { name: v })} label={t("fieldName")} />
              <EditableField value={guard.jabatan} onChange={(v) => onUpdate(guard.id, { jabatan: v })} label={t("fieldJabatan")} />
              <EditableField value={guard.regu || ""} onChange={(v) => onUpdate(guard.id, { regu: v })} label={t("fieldRegu")} />
              <EditableField value={guard.npk} onChange={(v) => onUpdate(guard.id, { npk: v })} label={t("fieldNpk")} />
              <EditableField value={guard.projek} onChange={(v) => onUpdate(guard.id, { projek: v })} label={t("fieldProjek")} />
              <EditableField value={guard.company || ""} onChange={(v) => onUpdate(guard.id, { company: v })} label={t("fieldCompany")} />
              <EditableField value={guard.phone || ""} onChange={(v) => onUpdate(guard.id, { phone: v })} label={t("fieldPhone")} />
            </div>
            <div style={{ background: COLORS.paper, borderColor: COLORS.line }} className="border rounded-lg px-2.5 py-2 mt-2 flex items-center gap-2">
              <Lock size={12} color={COLORS.amberDeep} />
              <label className="flex flex-col flex-1">
                <span style={{ color: COLORS.inkSoft }} className="text-[10px] uppercase tracking-wide">{t("fieldGuardPin")}</span>
                <input
                  defaultValue={guard.pin || ""}
                  key={guard.pin || "nopin"}
                  onBlur={(e) => { const v = e.target.value.replace(/\D/g, "").slice(0, 6); if (v !== (guard.pin || "")) onUpdate(guard.id, { pin: v }); }}
                  placeholder={t("noPinSet")}
                  inputMode="numeric"
                  style={{ color: COLORS.ink }}
                  className="bg-transparent text-base sm:text-sm py-1 outline-none tracking-widest"
                />
              </label>
            </div>
          </div>

          <div className="border-t pt-2" style={{ borderColor: COLORS.line }}>
            <p style={{ color: COLORS.ink }} className="text-xs font-medium mb-1.5">{t("bioTitle")}</p>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-3 gap-y-1 text-xs">
              <EditableField value={bio.ktp || ""} onChange={(v) => updateBio("ktp", v)} label={t("fieldKtp")} />
              <EditableField value={bio.bloodType || ""} onChange={(v) => updateBio("bloodType", v)} label={t("fieldBloodType")} />
              <EditableField value={bio.birthPlace || ""} onChange={(v) => updateBio("birthPlace", v)} label={t("fieldBirthPlace")} />
              <EditableField value={bio.birthDate || ""} onChange={(v) => updateBio("birthDate", v)} label={t("fieldBirthDate")} />
              <EditableField value={bio.emergencyName || ""} onChange={(v) => updateBio("emergencyName", v)} label={t("fieldEmergencyName")} />
              <EditableField value={bio.emergencyPhone || ""} onChange={(v) => updateBio("emergencyPhone", v)} label={t("fieldEmergencyPhone")} />
            </div>
            <div className="mt-1">
              <EditableField value={bio.address || ""} onChange={(v) => updateBio("address", v)} label={t("fieldAddress")} />
            </div>
          </div>

          <div className="border-t pt-2" style={{ borderColor: COLORS.line }}>
            <p style={{ color: COLORS.ink }} className="text-xs font-medium mb-1.5">{t("employmentTitle")}</p>
            <div className="grid grid-cols-2 gap-x-3 gap-y-1 text-xs">
              <EditableField value={bio.dateOfJoining || ""} onChange={(v) => updateBio("dateOfJoining", v)} label={t("fieldDateOfJoining")} />
              <EditableField value={bio.experience || ""} onChange={(v) => updateBio("experience", v)} label={t("fieldExperience")} />
            </div>
          </div>

        </div>
        <div className="flex flex-col gap-4 min-w-0">
          <MandatoryDocs guard={guard} onUpdate={onUpdate} />

          <div className="border-t pt-2" style={{ borderColor: COLORS.line }}>
            <p style={{ color: COLORS.ink }} className="text-xs font-medium mb-1.5 flex items-center gap-1.5"><AlertTriangle size={13} color={COLORS.red} /> {t("warningsTitle")}</p>
            {warnings.length === 0 && <p style={{ color: COLORS.inkSoft }} className="text-xs mb-2">{t("noWarnings")}</p>}
            <div className="flex flex-col gap-1.5 mb-2">
              {warnings.map((w) => (
                <div key={w.id} style={{ borderColor: COLORS.red }} className="border rounded p-2 flex items-start justify-between gap-2 text-xs">
                  <div className="min-w-0">
                    <p style={{ color: COLORS.ink }}>{w.reason}</p>
                    <p style={{ color: COLORS.inkSoft }}>{w.date} {w.issuedBy && `· ${t("deletedByLabel", w.issuedBy)}`}</p>
                  </div>
                  <button onClick={() => removeWarning(w.id)} style={{ color: COLORS.red }} className="flex-shrink-0 p-1"><X size={13} /></button>
                </div>
              ))}
            </div>
            <div className="flex flex-col gap-2">
              <input
                placeholder={t("warningReason")}
                value={newWarning.reason}
                onChange={(e) => setNewWarning({ ...newWarning, reason: e.target.value })}
                style={{ borderColor: COLORS.line }}
                className="border rounded px-2 py-2 text-base sm:text-xs min-h-[40px]"
              />
              <div className="flex gap-2">
                <input
                  type="date"
                  value={newWarning.date}
                  onChange={(e) => setNewWarning({ ...newWarning, date: e.target.value })}
                  style={{ borderColor: COLORS.line }}
                  className="border rounded px-2 py-2 text-base sm:text-xs min-h-[40px] flex-1"
                />
                <button onClick={addWarning} style={{ borderColor: COLORS.ink, color: COLORS.ink }} className="border rounded-lg px-3 text-xs font-medium min-h-[40px] flex items-center gap-1.5 max-w-full">
                  <Plus size={13} /> {t("addWarning")}
                </button>
              </div>
            </div>
          </div>

          <div className="border-t pt-2" style={{ borderColor: COLORS.line }}>
            <p style={{ color: COLORS.ink }} className="text-xs font-medium mb-1.5 flex items-center gap-1.5"><Award size={13} color={COLORS.amberDeep} /> {t("certificatesTitle")}</p>

            {certificates.length === 0 && <p style={{ color: COLORS.inkSoft }} className="text-xs mb-2">{t("noCertificates")}</p>}

            <div className="flex flex-col gap-1.5 mb-2">
              {certificates.map((c) => {
                const expired = c.expiry && c.expiry < today;
                const soon = !expired && c.expiry && c.expiry >= today && c.expiry <= in30;
                return (
                  <div key={c.id} style={{ borderColor: COLORS.line }} className="border rounded p-2 flex flex-col gap-1.5 text-xs">
                    <div className="flex items-start justify-between gap-2">
                      <div className="min-w-0">
                        <p style={{ color: COLORS.ink }} className="font-medium truncate">{c.name}{c.number && ` — ${c.number}`}</p>
                        {(c.issued || c.expiry) && (
                          <p style={{ color: expired ? COLORS.red : soon ? COLORS.amberDeep : COLORS.inkSoft }}>
                            {c.issued || "—"} → {c.expiry || "—"} {expired && `(${t("certExpired")})`} {soon && `(${t("certExpiringSoon")})`}
                          </p>
                        )}
                        {c.notes && <p style={{ color: COLORS.inkSoft }}>{c.notes}</p>}
                      </div>
                      <button onClick={() => removeCertificate(c.id)} style={{ color: COLORS.red }} className="flex-shrink-0 p-2 -m-1"><X size={14} /></button>
                    </div>
                    <FileAttachment
                      storageKey={`certfile:${c.id}`}
                      fileName={c.fileName}
                      fileSize={c.fileSize}
                      fileId={c.fileId}
                      onUploaded={({ fileName, size, fileId }) => updateCert(c.id, { fileName, fileSize: size, fileId })}
                      onRemoved={() => updateCert(c.id, { fileName: null, fileSize: null, fileId: null })}
                    />
                  </div>
                );
              })}
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
              <input placeholder={t("certName")} value={newCert.name} onChange={(e) => setNewCert({ ...newCert, name: e.target.value })} style={{ borderColor: COLORS.line }} className="border rounded px-2 py-2 text-base sm:text-xs min-h-[40px]" />
              <input placeholder={t("certNumber")} value={newCert.number} onChange={(e) => setNewCert({ ...newCert, number: e.target.value })} style={{ borderColor: COLORS.line }} className="border rounded px-2 py-2 text-base sm:text-xs min-h-[40px]" />
              <label className="flex flex-col gap-0.5">
                <span style={{ color: COLORS.inkSoft }} className="text-[10px]">{t("certIssued")}</span>
                <input type="date" value={newCert.issued} onChange={(e) => setNewCert({ ...newCert, issued: e.target.value })} style={{ borderColor: COLORS.line }} className="border rounded px-2 py-2 text-base sm:text-xs min-h-[40px] w-full" />
              </label>
              <label className="flex flex-col gap-0.5">
                <span style={{ color: COLORS.inkSoft }} className="text-[10px]">{t("certExpiry")}</span>
                <input type="date" value={newCert.expiry} onChange={(e) => setNewCert({ ...newCert, expiry: e.target.value })} style={{ borderColor: COLORS.line }} className="border rounded px-2 py-2 text-base sm:text-xs min-h-[40px] w-full" />
              </label>
              <label style={{ borderColor: COLORS.line, color: COLORS.inkSoft }} className="border rounded px-2 py-2 text-xs min-h-[40px] sm:col-span-2 flex items-center gap-2 cursor-pointer">
                <Paperclip size={13} />
                <span className="truncate">{newCertFile ? `${newCertFile.name} (${formatFileSize(newCertFile.size)})` : pick(lang, "Pilih file untuk diunggah (maks 2 MB)", "Choose file to upload (max 2 MB)", "Pilih fail untuk dimuat naik (maks 2 MB)")}</span>
                <input type="file" className="hidden" onChange={(e) => setNewCertFile(e.target.files?.[0] || null)} />
              </label>
              <input placeholder={t("certNotes")} value={newCert.notes} onChange={(e) => setNewCert({ ...newCert, notes: e.target.value })} style={{ borderColor: COLORS.line }} className="border rounded px-2 py-2 text-base sm:text-xs min-h-[40px] sm:col-span-2" />
            </div>
            <button onClick={addCertificate} style={{ borderColor: COLORS.amberDeep, color: COLORS.amberDeep }} className="border rounded-lg px-3 py-2.5 text-sm font-medium mt-2 flex items-center justify-center gap-1.5 w-full min-h-[44px]">
              <Plus size={13} /> {certBusy ? t("uploading") : t("addCertificate")}
            </button>
            {certErr && <p style={{ color: COLORS.red }} className="text-xs mt-1">{certErr}</p>}
          </div>
        </div>
      </div>
    </div>
  );
}

function MandatoryDocs({ guard, onUpdate }) {
  const { lang } = useContext(LangContext);
  const L = (a, b, c) => pick(lang, a, b, c);
  const docs = guard.docs || {};
  const levels = ["Gada Pratama", "Gada Madya", "Gada Utama"];
  function setDoc(key, patch) {
    onUpdate(guard.id, { docs: { ...docs, [key]: { ...(docs[key] || {}), ...patch } } });
  }
  const badge = { ok: [COLORS.green, COLORS.ink, L("OK", "OK", "OK")], missing: [COLORS.amber, COLORS.ink, L("Kurang", "Missing", "Kurang")], expired: [COLORS.ink, COLORS.ink, L("Kadaluarsa", "Expired", "Tamat tempoh")], soon: [COLORS.amber, COLORS.ink, L("Segera habis", "Expiring soon", "Hampir tamat")], nodate: [COLORS.amber, COLORS.ink, L("Isi tanggal", "Add date", "Isi tarikh")] };
  return (
    <div className="border-t pt-2" style={{ borderColor: COLORS.line }}>
      <p style={{ color: COLORS.ink }} className="text-xs font-medium mb-1.5 flex items-center gap-1.5">
        <ShieldAlert size={13} color={COLORS.amberDeep} /> {L("Dokumen wajib", "Mandatory documents", "Dokumen wajib")}
      </p>
      <div className="flex flex-col gap-2">
        {DOC_REQUIREMENTS.map((req) => {
          const st = docStatus(guard, req);
          const d = docs[req.key] || {};
          const [bg, fg, label] = badge[st];
          const until = docValidUntil(req, d);
          return (
            <div key={req.key} style={{ borderColor: COLORS.line }} className="border rounded-lg p-2 flex flex-col gap-1.5">
              <div className="flex items-center justify-between gap-2">
                <span style={{ color: COLORS.ink }} className="text-xs font-medium">{pick(lang, req.id, req.en, req.ms)}</span>
                <span style={{ borderColor: COLORS.line, color: fg }} className="border text-[11px] font-medium rounded-full px-2 py-0.5 inline-flex items-center gap-1.5 whitespace-nowrap"><i style={{ background: bg }} className="inline-block w-2 h-2 rounded-full" />{label}</span>
              </div>
              {req.special === "photo" ? (
                <p style={{ color: COLORS.inkSoft }} className="text-[10px]">{L("Unggah di bagian foto di atas.", "Upload in the photo box above.", "Muat naik di kotak gambar di atas.")}</p>
              ) : (
                <>
                  {(req.number || req.dated || req.hasLevel) && (
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-3 gap-y-1">
                      {req.hasLevel && (
                        <label className="flex flex-col">
                          <span style={{ color: COLORS.inkSoft }} className="text-[10px]">{L("Tingkat", "Level", "Tahap")}</span>
                          <select value={d.level || ""} onChange={(e) => setDoc(req.key, { level: e.target.value })} style={{ borderColor: COLORS.line }} className="border rounded px-2 py-2 text-base sm:text-xs min-h-[38px]">
                            <option value="">—</option>
                            {levels.map((l) => <option key={l} value={l}>{l}</option>)}
                          </select>
                        </label>
                      )}
                      {req.number && <EditableField value={d.number || ""} onChange={(v) => setDoc(req.key, { number: v })} label={L("Nomor", "Number", "Nombor")} />}
                      {req.dated && (
                        <label className="flex flex-col">
                          <span style={{ color: COLORS.inkSoft }} className="text-[10px]">{L("Tanggal terbit", "Issue date", "Tarikh dikeluarkan")}</span>
                          <input type="date" value={d.issued || ""} onChange={(e) => setDoc(req.key, { issued: e.target.value })} style={{ borderColor: COLORS.line }} className="border rounded px-2 py-2 text-base sm:text-xs min-h-[38px]" />
                        </label>
                      )}
                      {req.dated && (
                        <label className="flex flex-col">
                          <span style={{ color: COLORS.inkSoft }} className="text-[10px]">{L("Berlaku sampai", "Valid until", "Sah sehingga")}{req.validMonths ? ` (${L("otomatis", "auto", "automatik")} +${req.validMonths}${L(" bln", " mo", " bln")})` : ""}</span>
                          <input type="date" value={d.expiry || until || ""} onChange={(e) => setDoc(req.key, { expiry: e.target.value })} style={{ borderColor: COLORS.line }} className="border rounded px-2 py-2 text-base sm:text-xs min-h-[38px]" />
                        </label>
                      )}
                    </div>
                  )}
                  <FileAttachment
                    storageKey={`doc:${guard.id}:${req.key}`}
                    folderName={guard.name}
                    fileName={d.fileName}
                    fileSize={d.fileSize}
                    fileId={d.fileId}
                    onUploaded={({ fileName, size, fileId }) => setDoc(req.key, { fileName, fileSize: size, fileId })}
                    onRemoved={() => setDoc(req.key, { fileName: null, fileSize: null, fileId: null })}
                  />
                </>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

function DocsOverview({ roster }) {
  const { lang } = useContext(LangContext);
  const L = (a, b, c) => pick(lang, a, b, c);
  const sym = { ok: ["✓", COLORS.green], missing: ["✗", COLORS.ink], expired: ["!", COLORS.ink], soon: ["⚠", COLORS.ink], nodate: ["?", COLORS.inkSoft] };
  const incomplete = roster.filter((g) => guardDocIssues(g).length > 0).length;
  return (
    <div className="flex flex-col gap-3">
      <div style={{ background: COLORS.card, borderColor: COLORS.line }} className="border rounded-xl p-3 flex items-center justify-between gap-2 flex-wrap">
        <p style={{ color: COLORS.ink }} className="text-sm font-medium">
          {L(`${roster.length - incomplete} dari ${roster.length} personel lengkap`, `${roster.length - incomplete} of ${roster.length} guards complete`, `${roster.length - incomplete} daripada ${roster.length} anggota lengkap`)}
        </p>
        <button onClick={() => exportDocsChecklist(roster, lang)} style={{ borderColor: COLORS.amberDeep, color: COLORS.amberDeep }} className="border rounded-lg px-3 py-1.5 text-xs font-medium flex items-center gap-1.5 min-h-[36px]">
          <FileSpreadsheet size={13} /> {L("Unduh daftar kekurangan (Excel)", "Download missing list (Excel)", "Muat turun senarai kekurangan (Excel)")}
        </button>
      </div>
      <div style={{ background: COLORS.card, borderColor: COLORS.line }} className="border rounded-xl overflow-x-auto">
        <table className="text-xs w-full">
          <thead>
            <tr style={{ color: COLORS.inkSoft }}>
              <th className="text-left p-2 sticky left-0" style={{ background: COLORS.card }}>{L("Nama", "Name", "Nama")}</th>
              {DOC_REQUIREMENTS.map((r) => <th key={r.key} className="p-2 font-medium whitespace-nowrap">{r.key.toUpperCase()}</th>)}
            </tr>
          </thead>
          <tbody>
            {roster.map((g) => (
              <tr key={g.id} style={{ borderTop: `1px solid ${COLORS.line}` }}>
                <td className="p-2 whitespace-nowrap sticky left-0" style={{ background: COLORS.card, color: COLORS.ink }}>{g.name}</td>
                {DOC_REQUIREMENTS.map((r) => {
                  const [c, col] = sym[docStatus(g, r)];
                  return <td key={r.key} className="p-2 text-center font-bold" style={{ color: col }}>{c}</td>;
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <p style={{ color: COLORS.inkSoft }} className="text-[10px]">
        ✓ OK · ✗ {L("kurang", "missing", "kurang")} · ! {L("kadaluarsa", "expired", "tamat tempoh")} · ? {L("tanggal belum diisi", "date not filled", "tarikh belum diisi")}
      </p>
    </div>
  );
}

// Keeps typing local and saves once when the field loses focus (or Enter),
// instead of writing to Google Sheets on every keystroke.
function EditableField({ value, onChange, label }) {
  const [draft, setDraft] = useState(value ?? "");
  useEffect(() => { setDraft(value ?? ""); }, [value]);
  function commit() { if (draft !== (value ?? "")) onChange(draft); }
  return (
    <label className="flex flex-col">
      <span style={{ color: COLORS.inkSoft }} className="text-[10px] uppercase tracking-wide">{label}</span>
      <input
        value={draft}
        onChange={(e) => setDraft(e.target.value)}
        onBlur={commit}
        onKeyDown={(e) => { if (e.key === "Enter") e.currentTarget.blur(); }}
        style={{ color: COLORS.ink }}
        className="bg-transparent border-b border-transparent hover:border-gray-300 focus:border-gray-400 text-base sm:text-sm py-1.5 outline-none"
      />
    </label>
  );
}

function SettingsTab({ settings, onUpdate }) {
  const t = useT();
  const [pin, setPin] = useState(settings.pin);
  const [danruPin, setDanruPin] = useState(settings.danruPin || "2026");
  const [formTitle, setFormTitle] = useState(settings.formTitle);
  const [org, setOrg] = useState(settings.org || "");
  const [siteLat, setSiteLat] = useState(settings.siteLat || "");
  const [siteLng, setSiteLng] = useState(settings.siteLng || "");
  const [geofenceRadius, setGeofenceRadius] = useState(settings.geofenceRadius || "150");
  const [locating, setLocating] = useState(false);

  async function useCurrentLocation() {
    setLocating(true);
    const res = await captureLocation();
    if (res.status === "ok") {
      setSiteLat(String(res.lat));
      setSiteLng(String(res.lng));
    }
    setLocating(false);
  }

  return (
    <div style={{ background: COLORS.card, borderColor: COLORS.line }} className="border rounded-xl p-4 flex flex-col gap-3 max-w-sm">
      <label className="flex flex-col gap-1">
        <span style={{ color: COLORS.inkSoft }} className="text-xs">{t("managerPin")}</span>
        <input value={pin} onChange={(e) => setPin(e.target.value)} style={{ borderColor: COLORS.line }} className="border rounded px-2 py-2 text-base sm:text-sm min-h-[42px]" />
      </label>
      <label className="flex flex-col gap-1">
        <span style={{ color: COLORS.inkSoft }} className="text-xs">{t("danruPinLabel")}</span>
        <input value={danruPin} onChange={(e) => setDanruPin(e.target.value)} style={{ borderColor: COLORS.line }} className="border rounded px-2 py-2 text-base sm:text-sm min-h-[42px]" />
      </label>
      <label className="flex flex-col gap-1">
        <span style={{ color: COLORS.inkSoft }} className="text-xs">{t("formTitleLabel")}</span>
        <input value={formTitle} onChange={(e) => setFormTitle(e.target.value)} style={{ borderColor: COLORS.line }} className="border rounded px-2 py-2 text-base sm:text-sm min-h-[42px]" />
      </label>
      <label className="flex flex-col gap-1">
        <span style={{ color: COLORS.inkSoft }} className="text-xs">{t("subLabelLabel")}</span>
        <input value={org} onChange={(e) => setOrg(e.target.value)} style={{ borderColor: COLORS.line }} className="border rounded px-2 py-2 text-base sm:text-sm min-h-[42px]" />
      </label>

      <div className="border-t pt-3 flex flex-col gap-2" style={{ borderColor: COLORS.line }}>
        <div className="grid grid-cols-2 gap-2">
          <label className="flex flex-col gap-1">
            <span style={{ color: COLORS.inkSoft }} className="text-xs">{t("siteLatLabel")}</span>
            <input value={siteLat} onChange={(e) => setSiteLat(e.target.value)} inputMode="decimal" style={{ borderColor: COLORS.line }} className="border rounded px-2 py-2 text-base sm:text-sm min-h-[42px]" />
          </label>
          <label className="flex flex-col gap-1">
            <span style={{ color: COLORS.inkSoft }} className="text-xs">{t("siteLngLabel")}</span>
            <input value={siteLng} onChange={(e) => setSiteLng(e.target.value)} inputMode="decimal" style={{ borderColor: COLORS.line }} className="border rounded px-2 py-2 text-base sm:text-sm min-h-[42px]" />
          </label>
        </div>
        <label className="flex flex-col gap-1">
          <span style={{ color: COLORS.inkSoft }} className="text-xs">{t("geofenceRadiusLabel")}</span>
          <input value={geofenceRadius} onChange={(e) => setGeofenceRadius(e.target.value.replace(/\D/g, ""))} inputMode="numeric" style={{ borderColor: COLORS.line }} className="border rounded px-2 py-2 text-base sm:text-sm min-h-[42px]" />
        </label>
        <button
          onClick={useCurrentLocation}
          disabled={locating}
          style={{ borderColor: COLORS.amberDeep, color: COLORS.amberDeep }}
          className="border rounded-lg py-2 text-xs font-medium flex items-center justify-center gap-1.5 min-h-[40px] disabled:opacity-50"
        >
          <MapPin size={13} /> {locating ? t("gettingLocation") : t("useCurrentLocation")}
        </button>
      </div>

      <button
        onClick={() => onUpdate({ pin, danruPin, formTitle, org, siteLat, siteLng, geofenceRadius })}
        style={{ background: COLORS.green, color: COLORS.ink }}
        className="font-bold rounded-lg py-2.5 text-sm font-medium min-h-[44px]"
      >
        {t("saveSettings")}
      </button>
    </div>
  );
}

/* ---------- print sheet: mirrors the physical Form Kehadiran Reguler ---------- */

function PrintSheetInner({ guard, entries, periodId, settings }) {
  const days = periodDays(periodId);
  let lastMonth = null;

  return (
    <>
      <div className="grid items-start mb-2" style={{ gridTemplateColumns: "1fr 2fr 1fr" }}>
        <div></div>
        <h1 className="text-center text-lg font-bold tracking-wide">{settings.formTitle}</h1>
        <p className="text-right text-[10px] leading-tight self-center">{settings.org}</p>
      </div>
      <div style={{ borderBottom: "1px solid #111" }} className="mb-3"></div>

      <div className="grid text-sm mb-1" style={{ gridTemplateColumns: "1fr 1fr", rowGap: "2px" }}>
        <div className="flex gap-2"><span className="w-24">NAMA</span><span>: {guard.name}</span></div>
        <div className="flex gap-2 justify-end"><span>PROJEK</span><span>: {guard.projek}</span></div>
        <div className="flex gap-2"><span className="w-24">NPK</span><span>: {guard.npk}</span></div>
        <div className="flex gap-2 justify-end"><span>PERIODE ABSENSI</span><span>: {periodLabelPrint(periodId)}</span></div>
        <div className="flex gap-2"><span className="w-24">JABATAN</span><span>: {guard.jabatan}</span></div>
        <div></div>
      </div>

      <table className="w-full border-collapse text-xs mt-3" style={{ border: "1px solid #111" }}>
        <thead>
          <tr style={{ background: "#eee" }}>
            <th style={{ border: "1px solid #111" }} className="py-1 px-1 w-8">TGL</th>
            <th style={{ border: "1px solid #111" }} className="py-1 px-1 w-12">BULAN</th>
            <th style={{ border: "1px solid #111" }} className="py-1 px-1 w-16">JAM MULAI</th>
            <th style={{ border: "1px solid #111" }} className="py-1 px-1 w-16">JAM SELESAI</th>
            <th style={{ border: "1px solid #111" }} className="py-1 px-1 w-16">JUMLAH JAM</th>
            <th style={{ border: "1px solid #111" }} className="py-1 px-1 w-14">PARAF</th>
            <th style={{ border: "1px solid #111" }} className="py-1 px-1">KETERANGAN LOKASI/LAIN-LAIN</th>
          </tr>
        </thead>
        <tbody>
          {days.map((day) => {
            const e = entries[day.key];
            const showMonth = day.m !== lastMonth;
            lastMonth = day.m;
            const jam = e ? computeHours(e.jamMulai, e.jamSelesai) : "";
            return (
              <tr key={day.key}>
                <td style={{ border: "1px solid #111" }} className="text-center py-0.5">{day.d}</td>
                <td style={{ border: "1px solid #111" }} className="text-center py-0.5">{showMonth ? BULAN_SHORT[day.m] : ""}</td>
                <td style={{ border: "1px solid #111" }} className="text-center py-0.5">{e?.jamMulai || ""}</td>
                <td style={{ border: "1px solid #111" }} className="text-center py-0.5">{e?.jamSelesai || ""}</td>
                <td style={{ border: "1px solid #111" }} className="text-center py-0.5">{jam}</td>
                <td style={{ border: "1px solid #111" }} className="py-0.5"></td>
                <td style={{ border: "1px solid #111" }} className="py-0.5 px-1">{e?.keterangan && !e?.jamMulai ? e.keterangan : (e?.keterangan || "")}</td>
              </tr>
            );
          })}
        </tbody>
      </table>

      <p className="text-[10px] mt-2 italic">
        Keterangan: Yang dimasukan kedalam form ini adalah kehadiran sesuai dengan jadwal yang telah ditentukan, jika ada permintaan lembur dipisahkan dengan form lembur (form c).
      </p>

      <div className="mt-6">
        <p className="text-sm mb-3">Kolom Pengesahan</p>
        <div className="grid grid-cols-2 gap-8 text-xs">
          <div>
            <p className="mb-8">Diketahui &amp; Disetujui oleh Area Manager</p>
            <p>Paraf :</p>
            <p>Nama :</p>
            <p>Tanggal :</p>
          </div>
          <div>
            <p className="mb-8">Diketahui &amp; Disetujui oleh Pihak Client</p>
            <p>Paraf :</p>
            <p>Nama :</p>
            <p>Tanggal :</p>
          </div>
        </div>
      </div>
    </>
  );
}

const PRINT_PAGE_STYLE = { background: "#fff", color: "#111", fontFamily: "Georgia, 'Times New Roman', serif" };

function PrintSheet({ guard, entries, periodId, settings }) {
  return (
    <div className="print-area px-6 py-6 max-w-3xl mx-auto" style={PRINT_PAGE_STYLE}>
      <PrintSheetInner guard={guard} entries={entries} periodId={periodId} settings={settings} />
    </div>
  );
}

function AllPrintSheets({ roster, entries, periodId, settings }) {
  return (
    <div className="print-area">
      {roster.map((g, i) => (
        <div
          key={g.id}
          className="px-6 py-6 max-w-3xl mx-auto"
          style={{ ...PRINT_PAGE_STYLE, pageBreakAfter: i < roster.length - 1 ? "always" : "auto" }}
        >
          <PrintSheetInner guard={g} entries={entries[g.id] || {}} periodId={periodId} settings={settings} />
        </div>
      ))}
    </div>
  );
}

/* ---------- print panel: shared by Manager and Danru ---------- */

function PrintPanel({ roster, entries, settings, schedule, showFullExport, showRosterExport }) {
  const t = useT();
  const { lang } = useContext(LangContext);
  const [selectedGuardId, setSelectedGuardId] = useState(roster[0]?.id || null);
  const [periodId, setPeriodId] = useState(periodIdForToday());
  const [printMode, setPrintMode] = useState("one"); // one | all
  const periods = useMemo(() => recentPeriods(18), []);
  const selectedGuard = roster.find((g) => g.id === selectedGuardId) || null;

  return (
    <div className="flex flex-col gap-3">
      <div className="flex flex-col sm:flex-row flex-wrap gap-2 sm:items-center no-print">
        <select
          value={printMode}
          onChange={(e) => setPrintMode(e.target.value)}
          style={{ borderColor: COLORS.line }}
          className="border rounded-lg px-3 py-2 text-base sm:text-sm min-h-[42px] w-full sm:w-auto"
        >
          <option value="one">{t("printModeOne")}</option>
          <option value="all">{t("printModeAll")}</option>
        </select>

        {printMode === "one" && (
          <select
            value={selectedGuardId || ""}
            onChange={(e) => setSelectedGuardId(e.target.value)}
            style={{ borderColor: COLORS.line }}
            className="border rounded-lg px-3 py-2 text-base sm:text-sm min-h-[42px] w-full sm:w-auto"
          >
            <option value="" disabled>{t("selectPersonnel")}</option>
            {roster.map((g) => <option key={g.id} value={g.id}>{g.name}</option>)}
          </select>
        )}

        <select
          value={periodId}
          onChange={(e) => setPeriodId(e.target.value)}
          style={{ borderColor: COLORS.line }}
          className="border rounded-lg px-3 py-2 text-base sm:text-sm min-h-[42px] w-full sm:w-auto"
        >
          {periods.map((p) => <option key={p} value={p}>{periodLabel(p, lang)}</option>)}
        </select>

        <button
          onClick={() => window.print()}
          disabled={printMode === "one" && !selectedGuard}
          style={{ background: COLORS.amber, color: COLORS.ink }}
          className="font-bold rounded-lg px-4 py-2.5 text-sm font-medium flex items-center justify-center gap-1.5 disabled:opacity-40 min-h-[44px] w-full sm:w-auto"
        >
          <Printer size={15} /> {t("print")}
        </button>

        {showFullExport && (
          <button
            onClick={() => exportAllToExcel(roster, entries, schedule || {})}
            style={{ borderColor: COLORS.green, color: COLORS.green }}
            className="border rounded-lg px-3 py-2.5 text-sm font-medium flex items-center justify-center gap-1.5 min-h-[44px] w-full sm:w-auto"
          >
            <FileSpreadsheet size={15} /> {t("exportExcel")}
          </button>
        )}
        {showRosterExport && (
          <button
            onClick={() => exportRosterToExcel(roster)}
            style={{ borderColor: COLORS.green, color: COLORS.green }}
            className="border rounded-lg px-3 py-2.5 text-sm font-medium flex items-center justify-center gap-1.5 min-h-[44px] w-full sm:w-auto"
          >
            <FileSpreadsheet size={15} /> {t("exportRosterExcel")}
          </button>
        )}
      </div>

      <p className="no-print text-[11px]" style={{ color: COLORS.inkSoft }}>{t("printPdfHint")}</p>
      <p className="no-print text-[11px] flex items-start gap-1" style={{ color: COLORS.amberDeep }}>
        <AlertTriangle size={12} className="mt-0.5 flex-shrink-0" /> {t("appOnlyHint")}
      </p>

      {printMode === "one" && !selectedGuard && (
        <p style={{ color: COLORS.inkSoft }} className="text-sm no-print">{t("noPersonnelPrintHint")}</p>
      )}

      {printMode === "one" && selectedGuard && (
        <>
          <div className="no-print">
            <LocationLog entries={entries[selectedGuard.id] || {}} periodId={periodId} />
          </div>
          <div className="eco-preview"><PrintSheet guard={selectedGuard} entries={entries[selectedGuard.id] || {}} periodId={periodId} settings={settings} /></div>
        </>
      )}

      {printMode === "all" && (
        <div className="eco-preview"><AllPrintSheets roster={roster} entries={entries} periodId={periodId} settings={settings} /></div>
      )}
    </div>
  );
}
