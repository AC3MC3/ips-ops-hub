/**
 * fetch-data.js
 * Downloads SharePoint Excel files via shareable links (no Azure auth needed)
 * and fetches Jira open issues. Writes data.json for build-dashboard.js.
 */

import fetch from 'node-fetch';
import * as XLSX from 'xlsx';
import { writeFileSync } from 'fs';

// ── SharePoint direct-download URLs (shareable links + &download=1) ───────────
const VACATION_URL =
  'https://materialworkplace.sharepoint.com/:x:/s/IPS/IQANiG8tzON_SKrVCm9txisnAZEpqOJYTcnwYZZ89QoUtN4?e=gouRVL&download=1';
const AVAILABILITY_URL =
  'https://materialworkplace.sharepoint.com/:x:/s/IPS/IQCf_75fj2heT6D_jqrKe7zBAdQUsEwf9OknKKgHvZI1bEM?e=98gzU8&download=1';

// ── Helpers ───────────────────────────────────────────────────────────────────
function excelDate(serial) {
  const d = new Date(Math.round((serial - 25569) * 86400 * 1000));
  return `${d.getUTCMonth() + 1}/${d.getUTCDate()}`;
}

function weekDates(ref = new Date(), offset = 1) {
  const day = ref.getDay();
  const toMon = offset > 0 ? (day === 0 ? 1 : 8 - day) : (day === 0 ? -6 : 1 - day);
  const mon = new Date(ref);
  mon.setDate(ref.getDate() + toMon);
  return Array.from({ length: 5 }, (_, i) => {
    const d = new Date(mon);
    d.setDate(mon.getDate() + i);
    return `${d.getMonth() + 1}/${d.getDate()}`;
  });
}

async function downloadExcel(url, label) {
  const res = await fetch(url, { redirect: 'follow' });
  if (!res.ok) throw new Error(`${label}: HTTP ${res.status}`);
  return Buffer.from(await res.arrayBuffer());
}

// ── SharePoint: Vacation Calendar ─────────────────────────────────────────────
async function fetchVacation() {
  const buf = await downloadExcel(VACATION_URL, 'Vacation Calendar');
  const wb = XLSX.read(buf, { type: 'buffer' });
  const yr = new Date().getFullYear().toString();
  const ws = wb.Sheets[wb.SheetNames.find(s => s.includes(yr)) || wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
  const hdr = rows[0];
  const vac = {};
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const name = row[0];
    if (!name || typeof name !== 'string') continue;
    vac[name.trim()] = {};
    for (let c = 1; c < hdr.length; c++) {
      const dv = hdr[c];
      if (!dv) continue;
      const ds = typeof dv === 'number' ? excelDate(dv) : String(dv);
      const cv = row[c];
      if (cv && cv !== '') vac[name.trim()][ds] = String(cv).trim();
    }
  }
  return vac;
}

// ── SharePoint: SE Weekly Availability ───────────────────────────────────────
async function fetchAvailability() {
  const buf = await downloadExcel(AVAILABILITY_URL, 'Availability');
  const wb = XLSX.read(buf, { type: 'buffer' });
  const yr = new Date().getFullYear().toString();
  const ws = wb.Sheets[wb.SheetNames.find(s => s.includes(yr)) || wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
  const nw = weekDates();
  const availBlanks = [];
  let hIdx = -1;
  for (let r = rows.length - 1; r >= 0; r--) {
    const rd = rows[r].slice(4, 9).map(v =>
      v == null ? '' : (typeof v === 'number' ? excelDate(v) : String(v))
    );
    if (rd[0] === nw[0]) { hIdx = r; break; }
  }
  if (hIdx === -1) {
    console.warn('Next-week header not found in availability sheet');
    return { availBlanks, availHours: {} };
  }
  const availHours = {};
  for (let r = hIdx + 1; r < rows.length; r++) {
    const row = rows[r];
    const name = row[1];
    if (!name || typeof name !== 'string') break;
    const hrs = row.slice(4, 9).map(v => v == null ? null : v);
    if (hrs.every(v => v === null || v === '')) availBlanks.push(name.trim());
    availHours[name.trim()] = { hours: hrs, total: hrs.reduce((s, v) => s + (typeof v === 'number' ? v : 0), 0) };
  }
  return { availBlanks, availHours };
}

// ── Jira ──────────────────────────────────────────────────────────────────────
async function fetchJira() {
  const base = process.env.JIRA_BASE_URL || 'materialplus.atlassian.net';
  const email = process.env.JIRA_EMAIL;
  const tok = process.env.JIRA_API_TOKEN;
  const proj = process.env.JIRA_PROJECT_KEY || 'IPSIPS';
  const auth = Buffer.from(`${email}:${tok}`).toString('base64');
  const jql = encodeURIComponent(`project=${proj} AND status not in (Done,Closed,Resolved) ORDER BY updated DESC`);
  const res = await fetch(
    `https://${base}/rest/api/3/search?jql=${jql}&maxResults=200&fields=summary,status,assignee,issuetype,priority`,
    { headers: { Authorization: `Basic ${auth}`, Accept: 'application/json' } }
  );
  if (!res.ok) throw new Error(`Jira ${res.status}: ${await res.text()}`);
  return (await res.json()).issues.map(i => ({
    key: i.key,
    summary: i.fields.summary,
    status: i.fields.status?.name,
    assignee: i.fields.assignee?.displayName || 'Unassigned',
    type: i.fields.issuetype?.name,
    priority: i.fields.priority?.name,
  }));
}

// ── Main ──────────────────────────────────────────────────────────────────────
async function main() {
  const hasJira = process.env.JIRA_EMAIL && process.env.JIRA_API_TOKEN;
  let vac = {}, availBlanks = [], availHours = {}, jira = [];

  console.log('Fetching vacation calendar…');
  try { vac = await fetchVacation(); }
  catch (e) { console.warn('⚠ Vacation skip:', e.message); }

  console.log('Fetching availability…');
  try { ({ availBlanks, availHours } = await fetchAvailability()); }
  catch (e) { console.warn('⚠ Availability skip:', e.message); }
  console.log('Blanks:', availBlanks);

  if (hasJira) {
    console.log('Fetching Jira…');
    try { jira = await fetchJira(); }
    catch (e) { console.warn('⚠ Jira skip:', e.message); }
    console.log(`${jira.length} issues`);
  } else {
    console.warn('⚠ No Jira creds — skipping');
  }

  const now = new Date();
  writeFileSync('data.json', JSON.stringify({
    refreshedAt: now.toISOString(),
    todayStr: `${now.getMonth() + 1}/${now.getDate()}`,
    nextWeekDates: weekDates(now, 1),
    currentWeekDates: weekDates(now, 0),
    vac, availBlanks, availHours, jira,
  }, null, 2));
  console.log('✓ data.json written');
}

main().catch(e => { console.error(e); process.exit(1); });
