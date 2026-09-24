import { readFileSync, writeFileSync, existsSync, copyFileSync } from 'fs';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';

const __dir = dirname(fileURLToPath(import.meta.url));
const root  = resolve(__dir, '..');
const dataIn  = resolve(__dir, 'data.json');
const tmplIn  = resolve(root,  'public', 'template.html');
const htmlOut = resolve(root,  'public', 'index.html');

if (!existsSync(dataIn))  throw new Error('data.json not found — run fetch-data.js first');
if (!existsSync(tmplIn)) {
  if (existsSync(htmlOut)) { copyFileSync(htmlOut, tmplIn); console.log('Bootstrapped template.html'); }
  else throw new Error('public/template.html not found.');
}

const data = JSON.parse(readFileSync(dataIn, 'utf8'));
const { availBlanks, vac, jira, refreshedAt } = data;
const buildDt = new Date(refreshedAt);
const ts = buildDt.toLocaleString('en-IN', {
  timeZone:'Asia/Kolkata', day:'2-digit', month:'short', year:'numeric',
  hour:'2-digit', minute:'2-digit', hour12:false,
});

const blanksArr = availBlanks.map(n => JSON.stringify(n)).join(',');
const blanksLine = `const AVAIL_BLANKS=new Set([${blanksArr}]);`;
const vacLine = `const VAC=${JSON.stringify(vac,null,0)};`;
const jiraItems = jira.map(i =>
  `  {key:${JSON.stringify(i.key)},summary:${JSON.stringify(i.summary)},` +
  `status:${JSON.stringify(i.status)},assignee:${JSON.stringify(i.assignee)},` +
  `type:${JSON.stringify(i.type)},priority:${JSON.stringify(i.priority||'')}}`
).join(',\n');
const jiraLine = `const JIRA=[\n${jiraItems}\n];`;

let html = readFileSync(tmplIn, 'utf8');
html = html.replace(/const AVAIL_BLANKS=new Set\([^)]*\);/, blanksLine);
html = html.replace(/const VAC=\{[\s\S]*?\n\};/, vacLine);
html = html.replace(/const JIRA=\[[\s\S]*?\n\];/, jiraLine);
html = html.replace(/<!-- Data refreshed:.*?-->/, `<!-- Data refreshed: ${ts} IST -->`);
html = html.replace(/(<span[^>]*id="build-ts"[^>]*>)[^<]*(<\/span>)/, `$1${ts} IST$2`);

writeFileSync(htmlOut, html);
console.log(`index.html written — ${ts} IST`);
console.log(`  ${availBlanks.length} blank-availability SE(s): ${availBlanks.join(', ')||'none'}`);
console.log(`  ${jira.length} active Jira issue(s)`);