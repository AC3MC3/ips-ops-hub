import { ClientSecretCredential } from '@azure/identity';
import fetch from 'node-fetch';
import * as XLSX from 'xlsx';
import { writeFileSync } from 'fs';

let _cred = null;
function getCred() {
  if (!_cred) _cred = new ClientSecretCredential(
    process.env.AZURE_TENANT_ID,
    process.env.AZURE_CLIENT_ID,
    process.env.AZURE_CLIENT_SECRET,
  );
  return _cred;
}

async function graphToken() {
  const t = await getCred().getToken('https://graph.microsoft.com/.default');
  return t.token;
}
async function graphGet(url) {
  const t = await graphToken();
  const res = await fetch(`https://graph.microsoft.com/v1.0${url}`, { headers: { Authorization: `Bearer ${t}` } });
  if (!res.ok) throw new Error(`Graph ${res.status}: ${await res.text()}`);
  return res.json();
}
async function graphGetBinary(url) {
  const t = await graphToken();
  const res = await fetch(`https://graph.microsoft.com/v1.0${url}`, { headers: { Authorization: `Bearer ${t}` } });
  if (!res.ok) throw new Error(`Graph ${res.status}: ${await res.text()}`);
  return Buffer.from(await res.arrayBuffer());
}

function excelDate(serial) {
  const d = new Date(Math.round((serial - 25569) * 86400 * 1000));
  return `${d.getUTCMonth()+1}/${d.getUTCDate()}`;
}
function weekDates(ref = new Date(), offset = 1) {
  const day = ref.getDay();
  const toMon = offset > 0 ? (day === 0 ? 1 : 8 - day) : (day === 0 ? -6 : 1 - day);
  const mon = new Date(ref); mon.setDate(ref.getDate() + toMon);
  return Array.from({length:5},(_,i)=>{const d=new Date(mon);d.setDate(mon.getDate()+i);return `${d.getMonth()+1}/${d.getDate()}`;});
}

async function fetchVacation() {
  const drive = process.env.SHAREPOINT_DRIVE_ID;
  const items = await graphGet(`/drives/${drive}/root:/Key Team Docs:/children`);
  const file = items.value.find(f=>f.name==='IPS Vacation Calendar.xlsx');
  if (!file) throw new Error('IPS Vacation Calendar.xlsx not found');
  const buf = await graphGetBinary(`/drives/${drive}/items/${file.id}/content`);
  const wb = XLSX.read(buf,{type:'buffer'});
  const yr = new Date().getFullYear().toString();
  const ws = wb.Sheets[wb.SheetNames.find(s=>s.includes(yr))||wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws,{header:1,raw:true});
  const hdr = rows[0]; const vac = {};
  for (let r=1;r<rows.length;r++){
    const row=rows[r]; const name=row[0];
    if (!name||typeof name!=='string') continue;
    vac[name.trim()]={};
    for (let c=1;c<hdr.length;c++){
      const dv=hdr[c]; if(!dv) continue;
      const ds=typeof dv==='number'?excelDate(dv):String(dv);
      const cv=row[c]; if(cv&&cv!=='') vac[name.trim()][ds]=String(cv).trim();
    }
  }
  return vac;
}

async function fetchAvailability() {
  const drive = process.env.SHAREPOINT_DRIVE_ID;
  const items = await graphGet(`/drives/${drive}/root:/Key Team Docs:/children`);
  const file = items.value.find(f=>f.name==='Mat SE Weekly Availability.xlsx');
  if (!file) throw new Error('Mat SE Weekly Availability.xlsx not found');
  const buf = await graphGetBinary(`/drives/${drive}/items/${file.id}/content`);
  const wb = XLSX.read(buf,{type:'buffer'});
  const yr = new Date().getFullYear().toString();
  const ws = wb.Sheets[wb.SheetNames.find(s=>s.includes(yr))||wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws,{header:1,raw:true});
  const nw = weekDates(); const availBlanks = []; let hIdx=-1;
  for (let r=rows.length-1;r>=0;r--){
    const rd=rows[r].slice(4,9).map(v=>v==null?'':(typeof v==='number'?excelDate(v):String(v)));
    if(rd[0]===nw[0]){hIdx=r;break;}
  }
  if(hIdx===-1){console.warn('Next-week header not found');return{availBlanks,availHours:{}};}
  const availHours={};
  for(let r=hIdx+1;r<rows.length;r++){
    const row=rows[r]; const name=row[1];
    if(!name||typeof name!=='string') break;
    const hrs=row.slice(4,9).map(v=>v==null?null:v);
    if(hrs.every(v=>v===null||v==='')) availBlanks.push(name.trim());
    availHours[name.trim()]={hours:hrs,total:hrs.reduce((s,v)=>s+(typeof v==='number'?v:0),0)};
  }
  return{availBlanks,availHours};
}

async function fetchJira() {
  const base=process.env.JIRA_BASE_URL, email=process.env.JIRA_EMAIL, tok=process.env.JIRA_API_TOKEN;
  const proj=process.env.JIRA_PROJECT_KEY||'IPSIPS';
  const auth=Buffer.from(`${email}:${tok}`).toString('base64');
  const jql=encodeURIComponent(`project=${proj} AND status not in (Done,Closed,Resolved) ORDER BY updated DESC`);
  const res=await fetch(`https://${base}/rest/api/3/search?jql=${jql}&maxResults=200&fields=summary,status,assignee,issuetype,priority`,
    {headers:{Authorization:`Basic ${auth}`,Accept:'application/json'}});
  if(!res.ok) throw new Error(`Jira ${res.status}: ${await res.text()}`);
  return (await res.json()).issues.map(i=>({
    key:i.key,summary:i.fields.summary,status:i.fields.status?.name,
    assignee:i.fields.assignee?.displayName||'Unassigned',
    type:i.fields.issuetype?.name,priority:i.fields.priority?.name,
  }));
}

async function main() {
  const hasAzure = process.env.AZURE_TENANT_ID && process.env.AZURE_CLIENT_ID && process.env.AZURE_CLIENT_SECRET;
  const hasJira = process.env.JIRA_EMAIL && process.env.JIRA_API_TOKEN;
  let vac={}, availBlanks=[], availHours={}, jira=[];
  if (hasAzure) {
    console.log('Fetching vacation calendar...');
    try { vac = await fetchVacation(); } catch(e) { console.warn('Vacation skip:', e.message); }
    console.log('Fetching availability...');
    try { ({availBlanks,availHours} = await fetchAvailability()); } catch(e) { console.warn('Avail skip:', e.message); }
    console.log('Blanks:', availBlanks);
  } else { console.warn('No Azure creds — skipping SharePoint'); }
  if (hasJira) {
    console.log('Fetching Jira...');
    try { jira = await fetchJira(); } catch(e) { console.warn('Jira skip:', e.message); }
  }
  console.log(`${jira.length} issues`);
  const now = new Date();
  writeFileSync('data.json', JSON.stringify({
    refreshedAt:now.toISOString(),
    todayStr:`${now.getMonth()+1}/${now.getDate()}`,
    nextWeekDates:weekDates(now,1),
    currentWeekDates:weekDates(now,0),
    vac,availBlanks,availHours,jira
  },null,2));
  console.log('data.json written');
}
main().catch(e=>{console.error(e);process.exit(1);});
