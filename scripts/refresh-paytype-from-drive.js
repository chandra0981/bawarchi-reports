const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");
const { google } = require("googleapis");

const OUT_DIR = path.join(__dirname, "..", "data");
const DOWNLOAD_DIR = path.join(__dirname, "..", "tmp", "paytype");
const OUT_FILE = path.join(OUT_DIR, "paytype-data.json");

const FOLDER_ID = process.env.DRIVE_PAYTYPE_FOLDER_ID;
const SERVICE_ACCOUNT_JSON = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;

if (!FOLDER_ID) throw new Error("Missing GitHub secret: DRIVE_PAYTYPE_FOLDER_ID");
if (!SERVICE_ACCOUNT_JSON) throw new Error("Missing GitHub secret: GOOGLE_SERVICE_ACCOUNT_JSON");

fs.mkdirSync(OUT_DIR, { recursive: true });
fs.mkdirSync(DOWNLOAD_DIR, { recursive: true });

function toNumber(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  return Number(String(v).replace(/[R,\s]/g, "")) || 0;
}

function norm(h) {
  return String(h || "").trim().toLowerCase();
}

function parseDateFromFilename(name) {
  const m = String(name).match(/(\d{4}-\d{2}-\d{2})/);
  return m ? m[1] : "";
}

async function main() {
  const credentials = JSON.parse(SERVICE_ACCOUNT_JSON);
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ["https://www.googleapis.com/auth/drive.readonly"]
  });

  const drive = google.drive({ version: "v3", auth });

  const res = await drive.files.list({
    q: `'${FOLDER_ID}' in parents and trashed=false and (mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or mimeType='application/vnd.ms-excel')`,
    fields: "files(id,name,modifiedTime)",
    pageSize: 1000
  });

  const driveFiles = res.data.files || [];
  console.log("Files found:", driveFiles.map(f => f.name).join(", ") || "NONE");

  const allTransactions = [];

  for (const file of driveFiles) {
    const dest = path.join(DOWNLOAD_DIR, file.name);
    const response = await drive.files.get(
      { fileId: file.id, alt: "media" },
      { responseType: "arraybuffer" }
    );
    fs.writeFileSync(dest, Buffer.from(response.data));

    const reportDate = parseDateFromFilename(file.name);
    const workbook = XLSX.readFile(dest, { cellDates: true });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    const rows = XLSX.utils.sheet_to_json(sheet, {
      range: 8,
      defval: "",
      raw: true
    });

    console.log(`${file.name}: parsed rows = ${rows.length}`);
    console.log("First row sample:", JSON.stringify(rows[0] || {}));

    for (const row of rows) {
      const map = {};
      for (const k of Object.keys(row)) map[norm(k)] = row[k];

      const invoice = String(map["invoice number"] || "").trim();

      if (!invoice || invoice.toLowerCase().startsWith("invoice")) continue;
      if (String(map["date"] || "").startsWith("Node:")) continue;

      const dateValue = map["date"];
      const date = dateValue instanceof Date
        ? dateValue.toISOString().slice(0, 10)
        : String(dateValue || reportDate).slice(0, 10);

      const inclusive = toNumber(map["inclusive"]);
      if (!date || !inclusive) continue;

      allTransactions.push({
        source_file: file.name,
        date,
        time: map["time"] || "",
        hour: map["hour"] || "",
        half_hour: map["half hour"] || "",
        day_of_week: map["day of week"] || "",
        weekday_type: map["weekday type"] || "",
        invoice_number: invoice,
        payment_method: map["payment method"] || "",
        exclusive: toNumber(map["exclusive"]),
        tax: toNumber(map["tax"]),
        inclusive,
        employee: map["employee"] || "",
        guest_count: toNumber(map["guest count"]),
        order_number: map["order number"] || "",
        tips: toNumber(map["tips"]),
        device_name: map["device name"] || "",
        recorded_time: map["recorded time"] || "",
        cash: toNumber(map["cash"]),
        credit_card: toNumber(map["credit card"]),
        accounts: toNumber(map["accounts"]),
        cheque: toNumber(map["cheque"]),
        non_turnover: toNumber(map["non turnover"]),
        sales_name: map["sales name"] || ""
      });
    }
  }

  const dailyMap = {};
  const paymentMap = {};
  let total = 0, cash = 0, credit = 0, accounts = 0, tips = 0;

  for (const t of allTransactions) {
    total += t.inclusive;
    cash += t.cash;
    credit += t.credit_card;
    accounts += t.accounts;
    tips += t.tips;

    if (!dailyMap[t.date]) {
      dailyMap[t.date] = {
        date: t.date,
        day_of_week: t.day_of_week,
        inclusive: 0,
        cash: 0,
        credit_card: 0,
        accounts: 0,
        tips: 0,
        transactions: 0
      };
    }

    dailyMap[t.date].inclusive += t.inclusive;
    dailyMap[t.date].cash += t.cash;
    dailyMap[t.date].credit_card += t.credit_card;
    dailyMap[t.date].accounts += t.accounts;
    dailyMap[t.date].tips += t.tips;
    dailyMap[t.date].transactions += 1;

    const pm = t.payment_method || "Unknown";
    if (!paymentMap[pm]) paymentMap[pm] = { payment_method: pm, inclusive: 0, transactions: 0 };
    paymentMap[pm].inclusive += t.inclusive;
    paymentMap[pm].transactions += 1;
  }

  const output = {
    generated_at: new Date().toISOString(),
    files: driveFiles.map(f => ({ name: f.name, modifiedTime: f.modifiedTime })),
    summary: {
      total_inclusive: total,
      cash,
      credit_card: credit,
      accounts,
      tips,
      transactions: allTransactions.length
    },
    daily: Object.values(dailyMap).sort((a, b) => a.date.localeCompare(b.date)),
    payment_methods: Object.values(paymentMap).sort((a, b) => b.inclusive - a.inclusive),
    transactions: allTransactions
  };

  fs.writeFileSync(OUT_FILE, JSON.stringify(output, null, 2));
  console.log(`Created ${OUT_FILE}`);
  console.log(`Transactions loaded: ${allTransactions.length}`);
  console.log(`Total inclusive: ${total}`);
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});
