const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");
const { google } = require("googleapis");

const OUT_DIR = path.join(__dirname, "..", "data");
const TMP_DIR = path.join(__dirname, "..", "tmp");
const OUT_FILE = path.join(OUT_DIR, "gaap-dashboard-data.json");

const PAYTYPE_FOLDER_ID = process.env.DRIVE_PAYTYPE_FOLDER_ID;
const AUDIT_FOLDER_ID = process.env.DRIVE_SALES_AUDIT_FOLDER_ID;
const SERVICE_ACCOUNT_JSON = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;

if (!PAYTYPE_FOLDER_ID) throw new Error("Missing GitHub secret: DRIVE_PAYTYPE_FOLDER_ID");
if (!AUDIT_FOLDER_ID) throw new Error("Missing GitHub secret: DRIVE_SALES_AUDIT_FOLDER_ID");
if (!SERVICE_ACCOUNT_JSON) throw new Error("Missing GitHub secret: GOOGLE_SERVICE_ACCOUNT_JSON");

fs.mkdirSync(OUT_DIR, { recursive: true });
fs.mkdirSync(TMP_DIR, { recursive: true });

function norm(h){ return String(h || "").trim().toLowerCase(); }
function toNumber(v){ if(v===null||v===undefined||v==="")return 0; if(typeof v==="number")return v; return Number(String(v).replace(/[R,\s]/g,""))||0; }
function parseDateFromFilename(name){ const m=String(name).match(/(\d{4}-\d{2}-\d{2})/); return m?m[1]:""; }
function dateValue(v,fallback=""){ if(v instanceof Date&&!isNaN(v))return v.toISOString().slice(0,10); if(typeof v==="number"){const p=XLSX.SSF.parse_date_code(v); if(p)return `${String(p.y).padStart(4,"0")}-${String(p.m).padStart(2,"0")}-${String(p.d).padStart(2,"0")}`;} const s=String(v||"").trim(); const m=s.match(/(\d{4}-\d{2}-\d{2})/); if(m)return m[1]; const d=new Date(s); if(!isNaN(d))return d.toISOString().slice(0,10); return fallback; }
function timeValue(v){ if(v instanceof Date&&!isNaN(v))return v.toISOString().slice(11,19); if(typeof v==="number"){const total=Math.round(v*24*60*60); const h=Math.floor(total/3600)%24,m=Math.floor((total%3600)/60),s=total%60; return `${String(h).padStart(2,"0")}:${String(m).padStart(2,"0")}:${String(s).padStart(2,"0")}`;} return String(v||""); }
function isUber(employee,salesName){ return String(employee||"").trim().toLowerCase()==="online" && String(salesName||"").trim()!==""; }
function mapRow(row){ const map={}; for(const k of Object.keys(row)) map[norm(k)]=row[k]; return map; }

async function listDriveFiles(drive, folderId){
  const res = await drive.files.list({q:`'${folderId}' in parents and trashed=false and name contains '.xlsx'`,fields:"files(id,name,modifiedTime,mimeType)",pageSize:1000});
  return (res.data.files||[]).sort((a,b)=>String(a.name).localeCompare(String(b.name)));
}
async function downloadDriveFile(drive,file,subdir){
  const dir=path.join(TMP_DIR,subdir); fs.mkdirSync(dir,{recursive:true}); const dest=path.join(dir,file.name);
  const response=await drive.files.get({fileId:file.id,alt:"media"},{responseType:"arraybuffer"});
  fs.writeFileSync(dest,Buffer.from(response.data)); return dest;
}
function readGaapSheet(filePath,range=8){ const workbook=XLSX.readFile(filePath,{cellDates:true}); const sheet=workbook.Sheets[workbook.SheetNames[0]]; return XLSX.utils.sheet_to_json(sheet,{range,defval:"",raw:true}); }

function parsePaytypeFile(file,filePath){
  const fallbackDate=parseDateFromFilename(file.name); const rows=readGaapSheet(filePath,8); const out=[];
  for(const row of rows){ const m=mapRow(row); const invoice=String(m["invoice number"]||"").trim(); if(!invoice||invoice.toLowerCase().startsWith("invoice"))continue; const date=dateValue(m["date"],fallbackDate); const inclusive=toNumber(m["inclusive"]); if(!date||!inclusive)continue; const employee=String(m["employee"]||"").trim(); const salesName=String(m["sales name"]||"").trim();
    out.push({source_file:file.name,Date:date,Time:timeValue(m["time"]),Hour:String(m["hour"]||""),HalfHour:String(m["half hour"]||""),DayOfWeek:String(m["day of week"]||""),WeekdayType:String(m["weekday type"]||""),"Invoice Number":invoice,PaymentMethod:String(m["payment method"]||""),Exclusive:toNumber(m["exclusive"]),Tax:toNumber(m["tax"]),Inclusive:inclusive,Employee:employee,GuestCount:toNumber(m["guest count"]),OrderNumber:String(m["order number"]||""),Tips:toNumber(m["tips"]),Device:String(m["device name"]||""),RecordedTime:String(m["recorded time"]||""),Cash:toNumber(m["cash"]),Card:toNumber(m["credit card"]),Accounts:toNumber(m["accounts"]),Cheque:toNumber(m["cheque"]),NonTurnover:toNumber(m["non turnover"]),SalesName:salesName,Channel:isUber(employee,salesName)?"Uber Orders":"Non-Uber / Store"});
  } return out;
}
function parseAuditFile(file,filePath){
  const fallbackDate=parseDateFromFilename(file.name); const rows=readGaapSheet(filePath,8); const out=[];
  for(const row of rows){ const m=mapRow(row); const invoice=String(m["invoice number"]||"").trim(); const type=String(m["type"]||"").trim(); if(!invoice||invoice.toLowerCase().startsWith("invoice")||!type)continue; const date=dateValue(m["date"],fallbackDate); const employee=String(m["employee"]||"").trim(); const salesName=String(m["sales name"]||"").trim(); const riskText=type.toLowerCase(); let risk=""; if(riskText.includes("price override"))risk="Price Override"; if(riskText.includes("cancel")||String(m["after"]||"").toLowerCase().includes("cancel"))risk=risk?risk+", Cancel":"Cancel"; if(riskText.includes("void"))risk=risk?risk+", Void":"Void"; if(riskText.includes("remove"))risk=risk?risk+", Removed Item":"Removed Item";
    out.push({source_file:file.name,Date:date,Time:timeValue(m["time"]),Employee:employee,Invoice:invoice,Order:String(m["order number"]||""),Product:String(m["product"]||""),Type:type,Before:String(m["before"]||""),After:String(m["after"]||""),Difference:m["difference"]===""?null:toNumber(m["difference"]),"Bal Before":toNumber(m["bal before"]||m["balance before"]),"Bal After":toNumber(m["bal after"]||m["balance after"]),"Authorized By":String(m["authorized by"]||""),"Sales Name":salesName,Device:String(m["device name"]||""),Channel:isUber(employee,salesName)?"Uber Orders":"Non-Uber / Store",Risk:risk});
  } return out;
}
function groupAuditByInvoice(events){ const map={}; for(const e of events){ const inv=e.Invoice; if(!map[inv])map[inv]={AuditRows:0,employees:new Set(),types:new Set(),price:0,payments:0,cancels:0,maxBal:0}; const x=map[inv]; x.AuditRows++; if(e.Employee)x.employees.add(e.Employee); if(e.Type)x.types.add(e.Type); const t=String(e.Type||"").toLowerCase(); if(t.includes("price override"))x.price++; if(t.includes("process payment")||t.includes("charge customer amount")||t.includes("inptmthd"))x.payments++; if(t.includes("cancel")||t.includes("void")||String(e.Risk||"").toLowerCase().includes("cancel"))x.cancels++; x.maxBal=Math.max(x.maxBal,Number(e["Bal After"]||0),Number(e["Bal Before"]||0)); } return map; }
function buildInvoices(paytypes,events){ const audit=groupAuditByInvoice(events); const pmap={}; for(const p of paytypes){ const inv=p["Invoice Number"]; if(!pmap[inv])pmap[inv]={"Invoice Number":inv,Date:p.Date,Time:p.Time,Employee:p.Employee,SalesName:p.SalesName,OrderNumber:p.OrderNumber,PaymentMethod:p.PaymentMethod,Sales:0,Cash:0,Card:0,Accounts:0,Channel:p.Channel,Device:p.Device,Rows:0}; const r=pmap[inv]; r.Sales+=p.Inclusive; r.Cash+=p.Cash; r.Card+=p.Card; r.Accounts+=p.Accounts; r.Rows++; if(!r.SalesName&&p.SalesName)r.SalesName=p.SalesName; if(!r.OrderNumber&&p.OrderNumber)r.OrderNumber=p.OrderNumber; if(!r.PaymentMethod&&p.PaymentMethod)r.PaymentMethod=p.PaymentMethod; if(p.Channel==="Uber Orders")r.Channel="Uber Orders"; }
  return Object.values(pmap).map(r=>{const a=audit[r["Invoice Number"]]||{AuditRows:0,employees:new Set(),types:new Set(),price:0,payments:0,cancels:0,maxBal:0}; return {...r,AuditRows:a.AuditRows||0,AuditEmployee:[...(a.employees||[])].join(", "),Types:[...(a.types||[])].join(", "),PriceOverrides:a.price||0,Payments:a.payments||0,Cancellations:a.cancels||0,MaxBalance:a.maxBal||0};}).sort((a,b)=>String(a.Date+a.Time).localeCompare(String(b.Date+b.Time)));
}
function buildKpi(invoices,paytypes,events){ const uber=invoices.filter(x=>x.Channel==="Uber Orders"); const store=invoices.filter(x=>x.Channel!=="Uber Orders"); const duplicatePaytypeInvoices=Object.values(paytypes.reduce((m,p)=>{m[p["Invoice Number"]]=(m[p["Invoice Number"]]||0)+1; return m;},{})).filter(n=>n>1).length; return {total_sales:invoices.reduce((s,r)=>s+r.Sales,0),total_invoices:invoices.length,pay_rows:paytypes.length,uber_sales:uber.reduce((s,r)=>s+r.Sales,0),uber_invoices:uber.length,store_sales:store.reduce((s,r)=>s+r.Sales,0),store_invoices:store.length,audit_rows:events.length,audit_invoices:new Set(events.map(e=>e.Invoice)).size,price_override_count:events.filter(e=>String(e.Type).toLowerCase().includes("price override")).length,cancel_count:events.filter(e=>String(e.Type+" "+e.Risk).toLowerCase().includes("cancel")).length,duplicate_paytype_invoices:duplicatePaytypeInvoices}; }

async function main(){
  const credentials=JSON.parse(SERVICE_ACCOUNT_JSON); const auth=new google.auth.GoogleAuth({credentials,scopes:["https://www.googleapis.com/auth/drive.readonly"]}); const drive=google.drive({version:"v3",auth});
  const payFiles=await listDriveFiles(drive,PAYTYPE_FOLDER_ID); const auditFiles=await listDriveFiles(drive,AUDIT_FOLDER_ID);
  console.log("Paytype files found:",payFiles.map(f=>f.name).join(", ")||"NONE"); console.log("Sales Audit files found:",auditFiles.map(f=>f.name).join(", ")||"NONE");
  let paytypes=[], events=[];
  for(const f of payFiles){ const fp=await downloadDriveFile(drive,f,"paytype"); const parsed=parsePaytypeFile(f,fp); console.log(`${f.name}: paytype rows loaded = ${parsed.length}`); paytypes=paytypes.concat(parsed); }
  for(const f of auditFiles){ const fp=await downloadDriveFile(drive,f,"audit"); const parsed=parseAuditFile(f,fp); console.log(`${f.name}: audit rows loaded = ${parsed.length}`); events=events.concat(parsed); }
  const invoices=buildInvoices(paytypes,events); const output={generated_at:new Date().toISOString(),files:[...payFiles.map(f=>({type:"Paytype",name:f.name,modifiedTime:f.modifiedTime})),...auditFiles.map(f=>({type:"SalesAudit",name:f.name,modifiedTime:f.modifiedTime}))],kpi:buildKpi(invoices,paytypes,events),invoices,events,paytype_rows:paytypes};
  fs.writeFileSync(OUT_FILE,JSON.stringify(output,null,2)); fs.writeFileSync(path.join(OUT_DIR,"paytype-data.json"),JSON.stringify(output,null,2));
  console.log(`Created ${OUT_FILE}`); console.log(`Paytype rows loaded: ${paytypes.length}`); console.log(`Audit rows loaded: ${events.length}`); console.log(`Invoices created: ${invoices.length}`); console.log(`Total sales: ${output.kpi.total_sales}`);
}
main().catch(err=>{console.error(err);process.exit(1);});
