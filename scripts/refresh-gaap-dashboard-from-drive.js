const fs=require("fs"),path=require("path"),XLSX=require("xlsx");const{google}=require("googleapis");
const OUT_DIR=path.join(__dirname,"..","data"),TMP_DIR=path.join(__dirname,"..","tmp"),OUT_FILE=path.join(OUT_DIR,"gaap-dashboard-data.json");
const PAYTYPE_FOLDER_ID=process.env.DRIVE_PAYTYPE_FOLDER_ID,AUDIT_FOLDER_ID=process.env.DRIVE_SALES_AUDIT_FOLDER_ID,DEPT_FOLDER_ID=process.env.DRIVE_DEPARTMENT_SALES_FOLDER_ID||"",PURCHASE_FOLDER_ID=process.env.DRIVE_PURCHASE_REPORT_FOLDER_ID||"",SERVICE_ACCOUNT_JSON=process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
if(!PAYTYPE_FOLDER_ID||!AUDIT_FOLDER_ID||!SERVICE_ACCOUNT_JSON)throw new Error("Missing required Google Drive secrets");
fs.mkdirSync(OUT_DIR,{recursive:true});fs.mkdirSync(TMP_DIR,{recursive:true});
const norm=h=>String(h||"").trim().toLowerCase();
const toNumber=v=>v==null||v===""?0:(typeof v==="number"?v:Number(String(v).replace(/[R,\s]/g,""))||0);
const parseDateFromFilename=n=>(String(n).match(/(\d{4}-\d{2}-\d{2})/)||[])[1]||"";
function dateValue(v,f=""){if(v instanceof Date&&!isNaN(v))return v.toISOString().slice(0,10);if(typeof v==="number"){const p=XLSX.SSF.parse_date_code(v);if(p)return`${String(p.y).padStart(4,"0")}-${String(p.m).padStart(2,"0")}-${String(p.d).padStart(2,"0")}`}const s=String(v||""),m=s.match(/(\d{4}-\d{2}-\d{2})/);if(m)return m[1];const d=new Date(s);return!isNaN(d)?d.toISOString().slice(0,10):f}
function timeValue(v){if(v instanceof Date&&!isNaN(v))return v.toISOString().slice(11,19);if(typeof v==="number"){const t=Math.round(v*86400),h=Math.floor(t/3600)%24,m=Math.floor((t%3600)/60),s=t%60;return`${String(h).padStart(2,"0")}:${String(m).padStart(2,"0")}:${String(s).padStart(2,"0")}`}return String(v||"")}
function hourFromTime(t){const s=String(t||"");return s.length>=2?`${s.slice(0,2)}:00`:"Unknown"}
function mapRow(row){return Object.fromEntries(Object.keys(row).map(k=>[norm(k),row[k]]))}

function val(m,names){for(const n of names){const k=norm(n);if(Object.prototype.hasOwnProperty.call(m,k)&&m[k]!=="")return m[k];}return ""}
function categoryFromItem(item){const x=String(item||"").toLowerCase();if(/chicken|lamb|mutton|fish|prawn|egg|meat|mince/.test(x))return "Meat / Protein";if(/rice|flour|atta|maida|sugar|salt|oil|ghee|dal|lentil|spice|masala/.test(x))return "Dry Goods";if(/milk|cream|yogurt|curd|cheese|paneer|butter|amasi/.test(x))return "Dairy";if(/tomato|onion|potato|coriander|chilli|vegetable|veg|mint|ginger|garlic/.test(x))return "Vegetables";if(/gas|coal|charcoal|electric|water/.test(x))return "Utilities";if(/pack|container|bag|foil|tissue|napkin|cup|lid/.test(x))return "Packaging";return "Uncategorised"}

function classifyChannel(employee,salesName,paymentMethod,orderNumber,device){
  const e=String(employee||"").trim().toLowerCase(), sn=String(salesName||"").trim(), pm=String(paymentMethod||"").toLowerCase(), on=String(orderNumber||"").toLowerCase(), dev=String(device||"").toLowerCase();
  if(e==="online" && sn) return "Uber Orders";
  if(sn.toLowerCase().includes("mrd")||sn.toLowerCase().includes("mr d")||on.includes("mrd")||on.includes("mr d")) return "MrD";
  if(e==="online") return "Online / Store Pickup";
  if(dev.includes("pos")) return "Dine-In / Store";
  return "Non-Uber / Store";
}
async function list(drive,id){if(!id)return[];const r=await drive.files.list({q:`'${id}' in parents and trashed=false and name contains '.xlsx'`,fields:"files(id,name,modifiedTime,mimeType)",pageSize:1000});return(r.data.files||[]).sort((a,b)=>String(a.name).localeCompare(String(b.name)))}
async function dl(drive,f,sub){const dir=path.join(TMP_DIR,sub);fs.mkdirSync(dir,{recursive:true});const dest=path.join(dir,f.name),r=await drive.files.get({fileId:f.id,alt:"media"},{responseType:"arraybuffer"});fs.writeFileSync(dest,Buffer.from(r.data));return dest}
function read(fp,range=8){const wb=XLSX.readFile(fp,{cellDates:true});return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{range,defval:"",raw:true})}

function parsePaytype(file,fp){
  const fd=parseDateFromFilename(file.name),out=[];
  for(const row of read(fp,8)){
    const m=mapRow(row),inv=String(m["invoice number"]||"").trim();
    if(!inv||inv.toLowerCase().startsWith("invoice"))continue;
    const date=dateValue(m["date"],fd),inc=toNumber(m["inclusive"]);if(!date||!inc)continue;
    const emp=String(m["employee"]||"").trim(),sn=String(m["sales name"]||"").trim(),pm=String(m["payment method"]||"").trim(),ord=String(m["order number"]||"").trim(),dev=String(m["device name"]||"").trim();
    out.push({source_file:file.name,Date:date,Time:timeValue(m["time"]),Hour:String(m["hour"]||hourFromTime(timeValue(m["time"]))),DayOfWeek:String(m["day of week"]||""),"Invoice Number":inv,PaymentMethod:pm,Inclusive:inc,Employee:emp,OrderNumber:ord,Tips:toNumber(m["tips"]),Device:dev,Cash:toNumber(m["cash"]),Card:toNumber(m["credit card"]),Accounts:toNumber(m["accounts"]),SalesName:sn,Channel:classifyChannel(emp,sn,pm,ord,dev)});
  }
  return out;
}
function parseAudit(file,fp){
  const fd=parseDateFromFilename(file.name),out=[];
  for(const row of read(fp,8)){
    const m=mapRow(row),inv=String(m["invoice number"]||"").trim(),type=String(m["type"]||"").trim();
    if(!inv||inv.toLowerCase().startsWith("invoice")||!type)continue;
    const emp=String(m["employee"]||"").trim(),sn=String(m["sales name"]||"").trim(),ord=String(m["order number"]||"").trim(),dev=String(m["device name"]||"").trim(),t=type.toLowerCase();
    let risk="";if(t.includes("price override"))risk="Price Override";if(t.includes("cancel")||String(m["after"]||"").toLowerCase().includes("cancel"))risk=risk?risk+", Cancel":"Cancel";if(t.includes("void"))risk=risk?risk+", Void":"Void";if(t.includes("remove"))risk=risk?risk+", Removed Item":"Removed Item";
    const tm=timeValue(m["time"]);
    out.push({source_file:file.name,Date:dateValue(m["date"],fd),Time:tm,Hour:hourFromTime(tm),Employee:emp,Invoice:inv,Order:ord,Product:String(m["product"]||""),Type:type,Before:String(m["before"]||""),After:String(m["after"]||""),Difference:m["difference"]===""?null:toNumber(m["difference"]),"Bal Before":toNumber(m["bal before"]||m["balance before"]),"Bal After":toNumber(m["bal after"]||m["balance after"]),"Authorized By":String(m["authorized by"]||""),"Sales Name":sn,Device:dev,Channel:classifyChannel(emp,sn,"",ord,dev),Risk:risk});
  }
  return out;
}
function parseDepartment(file,fp){
  const rows=read(fp,8),out=[];let dept="";
  for(const row of rows){
    const m=mapRow(row);
    const prod=String(m["product name"]||"").trim();
    if(!prod)continue;
    if(prod.toLowerCase().startsWith("dept lvl 1:")){
      dept=prod.replace(/^Dept Lvl 1:\s*/i,"").trim();
      continue;
    }
    const date=dateValue(m["date"],parseDateFromFilename(file.name));
    if(!date)continue;
    const qty=toNumber(m["sold qty"]);
    const sales=toNumber(m["total sales incl"]||m["inclusive"]);
    const discount=toNumber(m["discount"]);
    if(!prod||(!qty&&!sales))continue;
    out.push({source_file:file.name,Date:date,Department:dept||"Unmapped",Product:prod,Qty:qty,Sales:sales,Discount:discount,Inclusive:toNumber(m["inclusive"]),Cost:toNumber(m["cost"])});
  }
  return out;
}

function parsePurchase(file,fp){
  const fd=parseDateFromFilename(file.name),out=[];
  const wb=XLSX.readFile(fp,{cellDates:true});
  const sheet=wb.Sheets[wb.SheetNames[0]];
  const rows=XLSX.utils.sheet_to_json(sheet,{header:1,defval:"",raw:true});
  // PurchaseReport layout:
  // rows 1-8  = report title / filters / blank metadata
  // row 9     = real column headers: Date, Product Name, Base UOM, Base Qty, Purchase UOM, Purchase Qty, Average Cost, Last Cost, Value, Tax, Total, Supplier
  // row 10    = secondary supplier header row, so actual useful rows start after the first 10 rows
  let currentDepartment="Unmapped";
  for(const r of rows.slice(10)){
    const first=String(r[0]||"").trim();
    if(!first && !r.some(x=>String(x||"").trim())) continue;
    if(/^Dept Lvl 1:/i.test(first)){
      currentDepartment=first.replace(/^Dept Lvl 1:\s*/i,"").split("(")[0].trim()||"Unmapped";
      continue;
    }
    const date=dateValue(r[0],fd);
    // Only accept real transaction rows with a valid date. This removes subtotals, blanks, and grand totals.
    if(!date) continue;
    const product=String(r[1]||"").trim()||"Purchase";
    const supplier=String(r[11]||"").trim()||"Unknown";
    const baseQty=toNumber(r[3]);
    const purchaseQty=toNumber(r[5]);
    const avgCost=toNumber(r[6]);
    const lastCost=toNumber(r[7]);
    const value=toNumber(r[8]);
    const tax=toNumber(r[9]);
    const total=toNumber(r[10]) || (value+tax);
    out.push({
      source_file:file.name,
      Date:date,
      Department:currentDepartment,
      Category:currentDepartment,
      Product:product,
      Item:product,
      BaseUOM:String(r[2]||"").trim(),
      BaseQty:baseQty,
      PurchaseUOM:String(r[4]||"").trim(),
      PurchaseQty:purchaseQty,
      Qty:purchaseQty||baseQty,
      AverageCost:avgCost,
      AvgCost:avgCost,
      LastCost:lastCost,
      UnitPrice:lastCost||avgCost,
      Value:value,
      Tax:tax,
      Vat:tax,
      Total:total,
      Amount:total,
      Supplier:supplier,
      Reference:"",
      InvoiceNumber:"",
      PaymentMethod:"",
      Notes:""
    });
  }
  return out;
}
function purchaseKpi(rows){return{purchase_total:rows.reduce((a,r)=>a+Number(r.Amount||0),0),purchase_rows:rows.length,purchase_vat:rows.reduce((a,r)=>a+Number(r.Vat||0),0),purchase_suppliers:new Set(rows.map(r=>r.Supplier).filter(Boolean)).size,purchase_categories:new Set(rows.map(r=>r.Category).filter(Boolean)).size}}

function auditByInvoice(events){
  const m={};
  for(const e of events){
    const inv=e.Invoice;m[inv]=m[inv]||{AuditRows:0,employees:new Set(),types:new Set(),price:0,payments:0,cancels:0,maxBal:0};
    const x=m[inv];x.AuditRows++;if(e.Employee)x.employees.add(e.Employee);if(e.Type)x.types.add(e.Type);
    const t=String(e.Type||"").toLowerCase();if(t.includes("price override"))x.price++;if(t.includes("payment")||t.includes("charge customer")||t.includes("inptmthd"))x.payments++;if((t+" "+String(e.Risk||"").toLowerCase()).includes("cancel"))x.cancels++;x.maxBal=Math.max(x.maxBal,Number(e["Bal After"]||0),Number(e["Bal Before"]||0));
  }
  return m;
}
function invoices(pay,events){
  const audit=auditByInvoice(events),m={};
  for(const p of pay){
    const inv=p["Invoice Number"];m[inv]=m[inv]||{"Invoice Number":inv,Date:p.Date,Time:p.Time,Hour:p.Hour,Employee:p.Employee,SalesName:p.SalesName,OrderNumber:p.OrderNumber,PaymentMethod:p.PaymentMethod,Sales:0,Cash:0,Card:0,Accounts:0,Channel:p.Channel,Device:p.Device,Rows:0};
    const r=m[inv];r.Sales+=p.Inclusive;r.Cash+=p.Cash;r.Card+=p.Card;r.Accounts+=p.Accounts;r.Rows++;if(p.Channel==="Uber Orders")r.Channel="Uber Orders";
  }
  return Object.values(m).map(r=>{const a=audit[r["Invoice Number"]]||{AuditRows:0,employees:new Set(),types:new Set(),price:0,payments:0,cancels:0,maxBal:0};return{...r,AuditRows:a.AuditRows,AuditEmployee:[...a.employees].join(", "),Types:[...a.types].join(", "),PriceOverrides:a.price,Payments:a.payments,Cancellations:a.cancels,MaxBalance:a.maxBal}}).sort((a,b)=>String(a.Date+a.Time).localeCompare(String(b.Date+b.Time)));
}
function kpi(inv,pay,events){const u=inv.filter(x=>x.Channel==="Uber Orders"),s=inv.filter(x=>x.Channel!=="Uber Orders");return{total_sales:inv.reduce((a,r)=>a+r.Sales,0),total_invoices:inv.length,pay_rows:pay.length,uber_sales:u.reduce((a,r)=>a+r.Sales,0),uber_invoices:u.length,store_sales:s.reduce((a,r)=>a+r.Sales,0),store_invoices:s.length,audit_rows:events.length,audit_invoices:new Set(events.map(e=>e.Invoice)).size,price_override_count:events.filter(e=>String(e.Type).toLowerCase().includes("price override")).length,cancel_count:events.filter(e=>String(e.Type+" "+e.Risk).toLowerCase().includes("cancel")).length}}
function buildDepartmentTimed(deptRows,events,payByInvoice){
  const deptByProd={}; for(const r of deptRows){ if(r.Product&&!deptByProd[r.Product]) deptByProd[r.Product]=r.Department; }
  const agg={}; for(const r of deptRows){ const k=`${r.Date}|${r.Product}`; agg[k]=agg[k]||{qty:0,sales:0,discount:0,dept:r.Department}; agg[k].qty+=r.Qty; agg[k].sales+=r.Sales; agg[k].discount+=r.Discount; }
  const auditItems=events.filter(e=>String(e.Type||"").toLowerCase().includes("add order items")&&e.Product);
  const counts={}; for(const e of auditItems){ const k=`${e.Date}|${e.Product}`; counts[k]=(counts[k]||0)+1; }
  const out=[];
  for(const e of auditItems){
    const k=`${e.Date}|${e.Product}`, a=agg[k]||{qty:1,sales:0,discount:0,dept:deptByProd[e.Product]||"Unmapped"}, c=counts[k]||1, pay=payByInvoice[e.Invoice]||{};
    out.push({Date:e.Date,Time:e.Time,Hour:e.Hour||hourFromTime(e.Time),Department:a.dept||deptByProd[e.Product]||"Unmapped",Product:e.Product,Qty:a.qty/c,Sales:a.sales/c,Discount:a.discount/c,Invoice:e.Invoice,Order:e.Order,Employee:e.Employee,Channel:pay.Channel||e.Channel||"Unmapped",PaymentMethod:pay.PaymentMethod||"",SalesName:pay.SalesName||e["Sales Name"]||"",Source:"Audit+Department"});
  }
  return out;
}
async function main(){
  const auth=new google.auth.GoogleAuth({credentials:JSON.parse(SERVICE_ACCOUNT_JSON),scopes:["https://www.googleapis.com/auth/drive.readonly"]}),drive=google.drive({version:"v3",auth});
  const payFiles=await list(drive,PAYTYPE_FOLDER_ID),auditFiles=await list(drive,AUDIT_FOLDER_ID),deptFiles=await list(drive,DEPT_FOLDER_ID),purchaseFiles=await list(drive,PURCHASE_FOLDER_ID);
  console.log("Paytype files found:",payFiles.map(f=>f.name).join(", ")||"NONE");console.log("Sales Audit files found:",auditFiles.map(f=>f.name).join(", ")||"NONE");console.log("Department Sales files found:",deptFiles.map(f=>f.name).join(", ")||"NONE");console.log("Purchase files found:",purchaseFiles.map(f=>f.name).join(", ")||"NONE");
  let pay=[],events=[],department_sales=[],purchase_rows=[];
  for(const f of payFiles){const parsed=parsePaytype(f,await dl(drive,f,"paytype"));console.log(`${f.name}: paytype rows loaded = ${parsed.length}`);pay=pay.concat(parsed)}
  for(const f of auditFiles){const parsed=parseAudit(f,await dl(drive,f,"audit"));console.log(`${f.name}: audit rows loaded = ${parsed.length}`);events=events.concat(parsed)}
  for(const f of deptFiles){const parsed=parseDepartment(f,await dl(drive,f,"department"));console.log(`${f.name}: department rows loaded = ${parsed.length}`);department_sales=department_sales.concat(parsed)}
  for(const f of purchaseFiles){const parsed=parsePurchase(f,await dl(drive,f,"purchase"));console.log(`${f.name}: purchase rows loaded = ${parsed.length}`);purchase_rows=purchase_rows.concat(parsed)}
  const inv=invoices(pay,events),payByInvoice={}; for(const p of pay){payByInvoice[p["Invoice Number"]]=p;}
  const department_timed=buildDepartmentTimed(department_sales,events,payByInvoice);
  const output={generated_at:new Date().toISOString(),files:[...payFiles.map(f=>({type:"Paytype",name:f.name,modifiedTime:f.modifiedTime})),...auditFiles.map(f=>({type:"SalesAudit",name:f.name,modifiedTime:f.modifiedTime})),...deptFiles.map(f=>({type:"DepartmentSales",name:f.name,modifiedTime:f.modifiedTime})),...purchaseFiles.map(f=>({type:"PurchaseReport",name:f.name,modifiedTime:f.modifiedTime}))],kpi:{...kpi(inv,pay,events),...purchaseKpi(purchase_rows)},invoices:inv,events,paytype_rows:pay,department_sales,department_timed,purchase_rows,purchase_report:purchase_rows,department_sales_files:deptFiles.map(f=>f.name),purchase_report_files:purchaseFiles.map(f=>f.name)};
  fs.writeFileSync(OUT_FILE,JSON.stringify(output,null,2));fs.writeFileSync(path.join(OUT_DIR,"paytype-data.json"),JSON.stringify(output,null,2));
  console.log(`Created ${OUT_FILE}`);console.log(`Paytype rows loaded: ${pay.length}`);console.log(`Audit rows loaded: ${events.length}`);console.log(`Department rows loaded: ${department_sales.length}`);console.log(`Department timed rows loaded: ${department_timed.length}`);console.log(`Purchase rows loaded: ${purchase_rows.length}`);console.log(`Invoices created: ${inv.length}`);console.log(`Total sales: ${output.kpi.total_sales}`);
}
main().catch(e=>{console.error(e);process.exit(1)})
