const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const inputDir = path.join(__dirname, '..', 'reports', 'paytype');
const outputFile = path.join(__dirname, '..', 'data', 'paytype-data.json');

function num(v){
  if (v === null || v === undefined || v === '') return 0;
  return Number(String(v).replace(/[R, ]/g,'')) || 0;
}
function findKey(row, names){
  const keys = Object.keys(row);
  for (const name of names){
    const k = keys.find(x => String(x).trim().toLowerCase() === name.toLowerCase());
    if (k) return row[k];
  }
  return '';
}
function toDateString(v, filename){
  if (v instanceof Date) return v.toISOString().slice(0,10);
  if (typeof v === 'number') {
    const d = XLSX.SSF.parse_date_code(v);
    if (d) return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`;
  }
  const s = String(v || '').trim();
  const m1 = s.match(/(\d{4}-\d{2}-\d{2})/);
  if (m1) return m1[1];
  const m2 = filename.match(/(\d{4}-\d{2}-\d{2})/);
  if (m2) return m2[1];
  return s.slice(0,10);
}
function normalizeRows(rawRows, filename){
  return rawRows.map(row => {
    const date = toDateString(findKey(row, ['Date','Day']), filename);
    const time = String(findKey(row, ['Time']) || '').slice(0,8);
    return {
      date,
      time,
      hour: String(findKey(row, ['Hour']) || time.slice(0,2) || ''),
      invoice: findKey(row, ['Invoice Number','Invoice']),
      payment_method: findKey(row, ['Payment Method']),
      inclusive: num(findKey(row, ['Inclusive','Paytype Inclusive'])),
      cash: num(findKey(row, ['Cash'])),
      credit_card: num(findKey(row, ['Credit Card'])),
      accounts: num(findKey(row, ['Accounts'])),
      tips: num(findKey(row, ['Tips'])),
      employee: findKey(row, ['Employee']),
      sales_name: findKey(row, ['Sales Name']),
      source_file: filename
    };
  }).filter(x => x.date && x.inclusive);
}
function readPaytypeFile(filePath){
  const wb = XLSX.readFile(filePath, {cellDates:true});
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, {defval:''});
  return normalizeRows(rows, path.basename(filePath));
}

const files = fs.readdirSync(inputDir)
  .filter(f => /^PaytypeReport_\d{4}-\d{2}-\d{2}\.xlsx$/i.test(f))
  .sort();

let transactions = [];
for (const f of files){
  transactions = transactions.concat(readPaytypeFile(path.join(inputDir, f)));
}

const dailyMap = new Map();
for (const r of transactions){
  if (!dailyMap.has(r.date)) dailyMap.set(r.date, {day:r.date, inclusive:0, cash:0, credit_card:0, accounts:0, tips:0, txns:0});
  const d = dailyMap.get(r.date);
  d.inclusive += r.inclusive;
  d.cash += r.cash;
  d.credit_card += r.credit_card;
  d.accounts += r.accounts;
  d.tips += r.tips;
  d.txns += 1;
}
const daily = Array.from(dailyMap.values()).sort((a,b)=>a.day.localeCompare(b.day));
const output = {
  generated_at: new Date().toISOString(),
  files,
  summary: {
    total_inclusive: daily.reduce((a,x)=>a+x.inclusive,0),
    total_cash: daily.reduce((a,x)=>a+x.cash,0),
    total_credit_card: daily.reduce((a,x)=>a+x.credit_card,0),
    total_accounts: daily.reduce((a,x)=>a+x.accounts,0),
    total_tips: daily.reduce((a,x)=>a+x.tips,0),
    total_transactions: transactions.length
  },
  daily,
  transactions
};
fs.mkdirSync(path.dirname(outputFile), {recursive:true});
fs.writeFileSync(outputFile, JSON.stringify(output, null, 2));
console.log(`Processed ${files.length} files and ${transactions.length} transactions.`);
console.log(`Saved ${outputFile}`);
