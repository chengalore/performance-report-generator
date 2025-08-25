import http from "http";
import fs from "fs";
import { parse } from "csv-parse/sync";
import formidable from "formidable"; // parse file uploads
import { google } from "googleapis";

const auth = new google.auth.GoogleAuth({
  keyFile: "credentials.json",
  scopes: ["https://www.googleapis.com/auth/presentations"],
});

const slides = google.slides({ version: "v1", auth });

// --- Load CSV helper ---
const monthNames = [
  "Jan",
  "Feb",
  "Mar",
  "Apr",
  "May",
  "Jun",
  "Jul",
  "Aug",
  "Sep",
  "Oct",
  "Nov",
  "Dec",
];

function normalizeMonth(raw) {
  if (!raw) return { start: null, end: null, label: "" };
  const map = {
    jan: 1,
    feb: 2,
    mar: 3,
    apr: 4,
    may: 5,
    jun: 6,
    jul: 7,
    aug: 8,
    sep: 9,
    oct: 10,
    nov: 11,
    dec: 12,
  };
  const clean = String(raw).trim();

  // Numeric month
  if (!isNaN(Number(clean))) {
    const n = Number(clean);
    return { start: n, end: n, label: monthNames[n - 1] };
  }

  // Ranges (e.g. "Feb-Jun" or "Jan–Dec")
  const parts = clean.split(/–|-/).map((p) => p.trim().toLowerCase());
  if (parts.length === 2) {
    const start = map[parts[0].substring(0, 3)];
    const end = map[parts[1].substring(0, 3)];
    const label = monthNames[start - 1] + "–" + monthNames[end - 1];
    return { start, end, label };
  }

  // Single month name
  const asText = map[clean.substring(0, 3).toLowerCase()];
  return { start: asText, end: asText, label: monthNames[asText - 1] };
}

function loadCSV(file) {
  const content = fs.readFileSync(file, "utf-8");
  return parse(content, {
    columns: (header) =>
      header.map((h) => h.trim().toLowerCase().replace(/\s+/g, "_")),
    skip_empty_lines: true,
  }).map((row) => {
    const norm = normalizeMonth(row.month);
    return {
      ...row,
      month_start: norm.start,
      month_end: norm.end,
      month_label: norm.label,
    };
  });
}

// --- Merge CSV rows into unique dataset ---
function mergeData(existing, newRows) {
  const merged = { ...existing }; // copy existing
  newRows.forEach((r) => {
    const key = [r.store_name, r.year, r.month].join("-");
    merged[key] = r; // replace or add
  });
  return merged;
}

// --- Start with base CSV(s) ---
// --- Start with saved merged data if exists, else fall back to base CSV ---
let mergedData = {};
if (fs.existsSync("./merged.csv")) {
  const rows = loadCSV("./merged.csv");
  mergedData = mergeData(mergedData, rows);
} else if (fs.existsSync("./6fcdd173-52bc-4ba7-bc36-3592faaf4f27.csv")) {
  const rows = loadCSV("./6fcdd173-52bc-4ba7-bc36-3592faaf4f27.csv");
  mergedData = mergeData(mergedData, rows);
}

// --- Convert merged object to array ---
function getData() {
  return Object.values(mergedData);
}

// --- Unique dropdown options ---
function getUnique(field) {
  const values = getData()
    .map((r) => r[field])
    .filter(Boolean);
  return [...new Set(values)];
}

// --- HTTP Server ---
http
  .createServer((req, res) => {
    if (req.url === "/" && req.method === "GET") {
      // --- Serve dashboard page ---
      const stores = getUnique("store_name");
      const years = getUnique("year");

      const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <title>Virtusize Multi-Comparison</title>
      <style>
        body { font-family: sans-serif; margin: 20px; }
        select, input { margin: 5px; padding: 3px; }
        table { border-collapse: collapse; margin-top: 20px; width: 95%; }
        th, td { border: 1px solid #ccc; padding: 6px; text-align: right; }
        th { background: #f4f4f4; text-align: left; }
        h1 { margin-bottom: 15px; }
      </style>
    </head>
    <body>
      <h1>Virtusize Multi-Comparison</h1>

     <form action="/upload" method="post" enctype="multipart/form-data">
  <label><b>Upload new CSV:</b></label>
  <input type="file" name="file" />
  <button type="submit">Upload & Merge</button>
</form>

<!-- New Download button -->
<p>
  <a href="/download" target="_blank">
    <button type="button">⬇️ Download Merged CSV</button>
  </a>
</p>
<hr/>


      <h3>Store Period 1</h3>
      <select id="store1">${stores.map((s) => `<option value="${s}">${s}</option>`).join("")}</select>
      <select id="year1">${years.map((y) => `<option value="${y}">${y}</option>`).join("")}</select>
      <input id="month1" placeholder="e.g. 6-8" />

      <h3>Store Period 2</h3>
      <select id="store2">${stores.map((s) => `<option value="${s}">${s}</option>`).join("")}</select>
      <select id="year2">${years.map((y) => `<option value="${y}">${y}</option>`).join("")}</select>
      <input id="month2" placeholder="e.g. 9-11" />

      <h3>Store Period 3</h3>
      <select id="store3">${stores.map((s) => `<option value="${s}">${s}</option>`).join("")}</select>
      <select id="year3">${years.map((y) => `<option value="${y}">${y}</option>`).join("")}</select>
      <input id="month3" placeholder="e.g. 1-12" />

      <br><br>
      <button onclick="calculate()">Compare</button>

      <script>
async function generateSlides() {
  // grab dropdown values
  const store1 = document.getElementById("store1").value;
  const year1 = document.getElementById("year1").value;
  const month1 = document.getElementById("month1").value;
  const m1 = buildMetrics(store1, year1, month1);

  const store2 = document.getElementById("store2").value;
  const year2 = document.getElementById("year2").value;
  const month2 = document.getElementById("month2").value;
  const m2 = buildMetrics(store2, year2, month2);

  const store3 = document.getElementById("store3").value;
  const year3 = document.getElementById("year3").value;
  const month3 = document.getElementById("month3").value;
  const m3 = buildMetrics(store3, year3, month3);

const replacements = {
  // --- Headers
  // --- Store Period 1 (B column)
  "{{B7}}": m1.coverage_rate_pct,
  "{{B8}}": m1.vs_usage_rate_pct,
  "{{B9}}": m1.comparisons_total,
  "{{B10}}": m1.vs_conversion_rate_pct,
  "{{B11}}": m1.purchase_compare_total,
  "{{B12}}": m1.vs_compare_share_pct,
  "{{B13}}": m1.vs_recommend_share_pct,
  "{{B14}}": m1.total_vs_share_pct,
  "{{B15}}": m1.conversion_rate_uplift_pct,
  "{{B16}}": m1.vs_user_aov,
  "{{B17}}": m1.non_vs_user_aov,
  "{{B18}}": m1.all_conversion_rate,
  "{{B19}}": m1.pv_total,
  "{{B20}}": m1.pv_vs_total,
  "{{B21}}": m1.clicks_total,
  "{{B22}}": m1.reco_total,
  "{{B23}}": m1.purchase_all_total,
  "{{B24}}": m1.purchase_recommend_total,
  "{{B25}}": m1.pv_uu,
  "{{B26}}": m1.pv_vs_uu,
  "{{B27}}": m1.clicks_uu,
  "{{B28}}": m1.comparisons_uu,
  "{{B29}}": m1.reco_uu,
  "{{B30}}": m1.purchase_all_uu,
  "{{B31}}": m1.purchase_compare_uu,
  "{{B32}}": m1.purchase_recommend_uu,
  "{{B33}}": m1.non_vs_conversion_rate,

  // --- Store Period 2 (C column)
  "{{C7}}": m2.coverage_rate_pct,
  "{{C8}}": m2.vs_usage_rate_pct,
  "{{C9}}": m2.comparisons_total,
  "{{C10}}": m2.vs_conversion_rate_pct,
  "{{C11}}": m2.purchase_compare_total,
  "{{C12}}": m2.vs_compare_share_pct,
  "{{C13}}": m2.vs_recommend_share_pct,
  "{{C14}}": m2.total_vs_share_pct,
  "{{C15}}": m2.conversion_rate_uplift_pct,
  "{{C16}}": m2.vs_user_aov,
  "{{C17}}": m2.non_vs_user_aov,
  "{{C18}}": m2.all_conversion_rate,
  "{{C19}}": m2.pv_total,
  "{{C20}}": m2.pv_vs_total,
  "{{C21}}": m2.clicks_total,
  "{{C22}}": m2.reco_total,
  "{{C23}}": m2.purchase_all_total,
  "{{C24}}": m2.purchase_recommend_total,
  "{{C25}}": m2.pv_uu,
  "{{C26}}": m2.pv_vs_uu,
  "{{C27}}": m2.clicks_uu,
  "{{C28}}": m2.comparisons_uu,
  "{{C29}}": m2.reco_uu,
  "{{C30}}": m2.purchase_all_uu,
  "{{C31}}": m2.purchase_compare_uu,
  "{{C32}}": m2.purchase_recommend_uu,
  "{{C33}}": m2.non_vs_conversion_rate,

  // --- Store Period 3 (D column)
  "{{D7}}": m3.coverage_rate_pct,
  "{{D8}}": m3.vs_usage_rate_pct,
  "{{D9}}": m3.comparisons_total,
  "{{D10}}": m3.vs_conversion_rate_pct,
  "{{D11}}": m3.purchase_compare_total,
  "{{D12}}": m3.vs_compare_share_pct,
  "{{D13}}": m3.vs_recommend_share_pct,
  "{{D14}}": m3.total_vs_share_pct,
  "{{D15}}": m3.conversion_rate_uplift_pct,
  "{{D16}}": m3.vs_user_aov,
  "{{D17}}": m3.non_vs_user_aov,
  "{{D18}}": m3.all_conversion_rate,
  "{{D19}}": m3.pv_total,
  "{{D20}}": m3.pv_vs_total,
  "{{D21}}": m3.clicks_total,
  "{{D22}}": m3.reco_total,
  "{{D23}}": m3.purchase_all_total,
  "{{D24}}": m3.purchase_recommend_total,
  "{{D25}}": m3.pv_uu,
  "{{D26}}": m3.pv_vs_uu,
  "{{D27}}": m3.clicks_uu,
  "{{D28}}": m3.comparisons_uu,
  "{{D29}}": m3.reco_uu,
  "{{D30}}": m3.purchase_all_uu,
  "{{D31}}": m3.purchase_compare_uu,
  "{{D32}}": m3.purchase_recommend_uu,
  "{{D33}}": m3.non_vs_conversion_rate,
};




  try {
    const res = await fetch("/generate-slides", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ replacements })
    });
    const data = await res.json();
    if (data.success) {
      window.open(data.url, "_blank");
    } else {
      alert("❌ Error: " + data.error);
    }
  } catch (err) {
    alert("❌ Request failed: " + err.message);
  }
}
</script>


      <div id="result"></div>
<br>
<button type="button" onclick="generateSlides()">📊 Generate Slides</button>



      <script>
        const data = ${JSON.stringify(getData())};
        const monthNames = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

        function parseMonth(m) {
          const map = {jan:1,feb:2,mar:3,apr:4,may:5,jun:6,
                       jul:7,aug:8,sep:9,oct:10,nov:11,dec:12};
          const clean = String(m).toLowerCase().trim();
          const asNum = Number(clean);
          if (!isNaN(asNum)) return asNum;
          return map[clean.substring(0,3)];
        }

        function parseMonthRange(input) {
    
          if (!input) return { start: null, end: null };
          const clean = String(input).toLowerCase().replace("--", "–");
          if (clean.includes("–") || clean.includes("-")) {
            const parts = clean.split(/–|-/).map(p => p.trim());
            return { start: parseMonth(parts[0]), end: parseMonth(parts[1]) };
          }
          const m = parseMonth(clean);
          return { start: m, end: m };
        }
    function formatMonthRange(start, end, rawInput) {
  // If the raw input looks like "Jan-Dec" or "Feb-Jun", preserve it
  const clean = String(rawInput || "").trim();
  if (/Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec/i.test(clean) && clean.includes("-")) {
    return clean.replace(/-/g, "–"); // normalize dash to en dash
  }
  if (!start || !end) return "";
  if (start === end) return monthNames[start-1];
  return monthNames[start-1] + "–" + monthNames[end-1];
}
       function getMetricSum(metric, store, year, startMonth, endMonth) {
  let total = 0;
  let found = false;

  data.forEach(r => {
    const rowM = parseMonth(r.month);
    if (String(r.year) === String(year) &&
        rowM >= startMonth && rowM <= endMonth &&
        r.store_name.toLowerCase() === store.toLowerCase()) {
      if (r[metric] !== undefined && r[metric] !== "") {
        found = true;
        total += Number(r[metric] || 0);
      }
    }
  });

  if (!found) return "N/A";   // no rows at all for this metric
  return total;
}


        function getMonthsUsed(year, startMonth, endMonth, store) {
          const available = [];
          data.forEach(r => {
            const rowM = parseMonth(r.month);
            if (String(r.year) === String(year) &&
                rowM >= startMonth && rowM <= endMonth &&
                r.store_name.toLowerCase() === store.toLowerCase() &&
                !available.includes(rowM)) {
              available.push(rowM);
            }
          });

          if (available.length === 0) return "No data";
          return available.sort((a,b)=>a-b).map(m => monthNames[m-1]).join(", ");
        }

        function formatNumber(num) {
          if (num === "N/A") return "N/A";
          if (isNaN(num)) return "N/A";
          return Number(num).toLocaleString("en-US");
        }

        function formatPercent(num) {
          if (num === "N/A") return "N/A";
          if (isNaN(num)) return "N/A";
          return parseFloat(num).toFixed(2) + "%";
        }

        function buildMetrics(store, year, monthInput) {
  const range = parseMonthRange(monthInput);
  if (!range.start || !range.end) return {};

  // --- Sum base metrics ---
  const pv_total = getMetricSum("pv_total", store, year, range.start, range.end);
  const pv_vs_total = getMetricSum("pv_vs_total", store, year, range.start, range.end);
  const comparisons_total = getMetricSum("comparisons_total", store, year, range.start, range.end);
  const purchase_all_total = getMetricSum("purchase_all_total", store, year, range.start, range.end);
  const purchase_compare_total = getMetricSum("purchase_compare_total", store, year, range.start, range.end);
  const purchase_recommend_total = getMetricSum("purchase_recommend_total", store, year, range.start, range.end);
  const clicks_total = getMetricSum("clicks_total", store, year, range.start, range.end);
  const reco_total = getMetricSum("reco_total", store, year, range.start, range.end);

  // --- UU metrics too ---
  const pv_uu = getMetricSum("pv_uu", store, year, range.start, range.end);
  const pv_vs_uu = getMetricSum("pv_vs_uu", store, year, range.start, range.end);
  const clicks_uu = getMetricSum("clicks_uu", store, year, range.start, range.end);
  const comparisons_uu = getMetricSum("comparisons_uu", store, year, range.start, range.end);
  const reco_uu = getMetricSum("reco_uu", store, year, range.start, range.end);
  const purchase_all_uu = getMetricSum("purchase_all_uu", store, year, range.start, range.end);
  const purchase_compare_uu = getMetricSum("purchase_compare_uu", store, year, range.start, range.end);
  const purchase_recommend_uu = getMetricSum("purchase_recommend_uu", store, year, range.start, range.end);

  // --- Derived metrics ---
  const coverage_rate_pct = pv_total ? (pv_vs_total / pv_total) : null;
  const vs_usage_rate_pct = pv_vs_total ? (comparisons_total / pv_vs_total) : null;
  const vs_conversion_rate_pct = comparisons_total ? (purchase_compare_total / comparisons_total) : null;
  const vs_compare_share_pct = purchase_all_total ? (purchase_compare_total / purchase_all_total) : null;
  const vs_recommend_share_pct = purchase_all_total ? (purchase_recommend_total / purchase_all_total) : null;
  const total_vs_share_pct = purchase_all_total ? ((purchase_compare_total + purchase_recommend_total) / purchase_all_total) : null;
  const all_conversion_rate = pv_total ? (purchase_all_total / pv_total) : null;

  const non_vs_purchases = purchase_all_total - purchase_compare_total - purchase_recommend_total;
  const non_vs_conversion_rate = pv_vs_total ? (non_vs_purchases / pv_vs_total) : null;

  const vsRate = comparisons_total ? (purchase_compare_total / comparisons_total) : 0;
  const nonVsRate = pv_vs_total ? (non_vs_purchases / pv_vs_total) : 0;
  const conversion_rate_uplift_pct = nonVsRate ? ((vsRate - nonVsRate) / nonVsRate) : null;

  return {
    coverage_rate_pct: coverage_rate_pct !== null ? formatPercent(coverage_rate_pct*100) : "N/A",
    vs_usage_rate_pct: vs_usage_rate_pct !== null ? formatPercent(vs_usage_rate_pct*100) : "N/A",
    comparisons_total: formatNumber(comparisons_total),
    vs_conversion_rate_pct: vs_conversion_rate_pct !== null ? formatPercent(vs_conversion_rate_pct*100) : "N/A",
    purchase_compare_total: formatNumber(purchase_compare_total),
    vs_compare_share_pct: vs_compare_share_pct !== null ? formatPercent(vs_compare_share_pct*100) : "N/A",
    vs_recommend_share_pct: vs_recommend_share_pct !== null ? formatPercent(vs_recommend_share_pct*100) : "N/A",
    total_vs_share_pct: total_vs_share_pct !== null ? formatPercent(total_vs_share_pct*100) : "N/A",
    conversion_rate_uplift_pct: conversion_rate_uplift_pct !== null ? formatPercent(conversion_rate_uplift_pct*100) : "N/A",
    vs_user_aov: "N/A",   // need sales data to compute
    non_vs_user_aov: "N/A",
    all_conversion_rate: all_conversion_rate !== null ? formatPercent(all_conversion_rate*100) : "N/A",
    pv_total: formatNumber(pv_total),
    pv_vs_total: formatNumber(pv_vs_total),
    clicks_total: formatNumber(clicks_total),
    reco_total: formatNumber(reco_total),
    purchase_all_total: formatNumber(purchase_all_total),
    purchase_recommend_total: formatNumber(purchase_recommend_total),
    pv_uu: formatNumber(pv_uu),
    pv_vs_uu: formatNumber(pv_vs_uu),
    clicks_uu: formatNumber(clicks_uu),
    comparisons_uu: formatNumber(comparisons_uu),
    reco_uu: formatNumber(reco_uu),
    purchase_all_uu: formatNumber(purchase_all_uu),
    purchase_compare_uu: formatNumber(purchase_compare_uu),
    purchase_recommend_uu: formatNumber(purchase_recommend_uu),
    non_vs_conversion_rate: non_vs_conversion_rate !== null ? formatPercent(non_vs_conversion_rate*100) : "N/A"
  };
}

        function calculate() {
          const store1 = document.getElementById("store1").value;
          const year1 = document.getElementById("year1").value;
          const month1 = document.getElementById("month1").value;

          const store2 = document.getElementById("store2").value;
          const year2 = document.getElementById("year2").value;
          const month2 = document.getElementById("month2").value;

          const store3 = document.getElementById("store3").value;
          const year3 = document.getElementById("year3").value;
          const month3 = document.getElementById("month3").value;

          const m1 = buildMetrics(store1, year1, month1);
          const m2 = buildMetrics(store2, year2, month2);
          const m3 = buildMetrics(store3, year3, month3);

          const range1 = parseMonthRange(month1);
const range2 = parseMonthRange(month2);
const range3 = parseMonthRange(month3);

const label1 = formatMonthRange(range1.start, range1.end, month1);
const label2 = formatMonthRange(range2.start, range2.end, month2);
const label3 = formatMonthRange(range3.start, range3.end, month3);

let table = "<table><tr>" +
    "<th>Metric</th>" +
    "<th>"+store1+" ("+label1+" "+year1+")</th>" +
    "<th>"+store2+" ("+label2+" "+year2+")</th>" +
    "<th>"+store3+" ("+label3+" "+year3+")</th>" +
  "</tr>";


          Object.keys(m1).forEach(k=>{
            table += "<tr><td>"+k+"</td><td>"+m1[k]+"</td><td>"+m2[k]+"</td><td>"+m3[k]+"</td></tr>";
          });
          table += "</table>";

          document.getElementById("result").innerHTML = table;
        }
      </script>
    </body>
    </html>
    `;

      res.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
      res.end(html);
    } else if (req.url === "/upload" && req.method === "POST") {
      const form = formidable({
        multiples: false,
        uploadDir: "./",
        keepExtensions: true,
      });
      form.parse(req, (err, fields, files) => {
        if (err) {
          res.writeHead(500);
          return res.end("Upload error");
        }

        const filePath = files.file[0].filepath;
        const rows = loadCSV(filePath);

        // Merge into in-memory dataset
        mergedData = mergeData(mergedData, rows);

        // Convert merged object back to array
        const allRows = Object.values(mergedData);

        // Persist back to master CSV (overwrite merged.csv)
        const header = Object.keys(allRows[0]).join(",") + "\n";
        const body = allRows.map((r) => Object.values(r).join(",")).join("\n");
        fs.writeFileSync("./merged.csv", header + body);

        // Clean up temp upload
        fs.unlinkSync(filePath);

        res.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
        res.end(`
      <p>✅ File uploaded and merged successfully (saved to merged.csv).</p>
      <a href="/">Go back to dashboard</a>
    `);
      });
      
  }
  else if (req.url === "/download" && req.method === "GET") {
  if (fs.existsSync("./merged.csv")) {
    res.writeHead(200, {
      "Content-Type": "text/csv",
      "Content-Disposition": "attachment; filename=merged.csv"
    });
    const stream = fs.createReadStream("./merged.csv");
    stream.pipe(res);
    stream.on("end", () => res.end());
    stream.on("error", (err) => {
      res.writeHead(500, { "Content-Type": "text/plain" });
      res.end("Error reading file: " + err.message);
    });
  } else {
    res.writeHead(404, { "Content-Type": "text/plain" });
    res.end("No merged.csv file available.");
  }
}

  else if (req.url === "/generate-slides" && req.method === "POST") {
  let body = "";
  req.on("data", chunk => (body += chunk));
  req.on("end", async () => {
    try {
      const { replacements } = JSON.parse(body);   // your HTML will send replacements
      const link = await generateSlidesReport(replacements);

      res.writeHead(200, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ success: true, url: link }));
    } catch (err) {
      res.writeHead(500, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ success: false, error: err.message }));
    }
  });
}

  })
  .listen(3000, () => {
    console.log("✅ Open http://localhost:3000 in your browser");
  });

// 6. Helper at bottom
async function generateSlidesReport(replacements) {
  const presentationId = "1pdQFbo6pyXXAlsPyHw1Ho1KeewU791S554W6dQ19OHc";
  const presentation = await slides.presentations.get({ presentationId });

  const totalSlides = presentation.data.slides.length;
  const template1 = presentation.data.slides[totalSlides - 2].objectId;
  const template2 = presentation.data.slides[totalSlides - 1].objectId;

  // Duplicate
 const duplicateRes = await slides.presentations.batchUpdate({
  presentationId,
  requestBody: {
    requests: [
      { duplicateObject: { objectId: template1 } },
      { duplicateObject: { objectId: template2 } },
    ],
  },
});


  const newSlides = duplicateRes.data.replies.map(r => r.duplicateObject.objectId);
 const moveRequests = [
    {
      updateSlidesPosition: {
        slideObjectIds: [newSlides[1]],
        insertionIndex: 0,
      },
    },
    {
      updateSlidesPosition: {
        slideObjectIds: [newSlides[0]],
        insertionIndex: 0,
      },
    },
  ];

// Move them before the template1 (so templates stay at the end)
 await slides.presentations.batchUpdate({
    presentationId,
    requestBody: { requests: moveRequests },
  });

  // Step 3: Replace placeholders in the new top slides
  const requests = Object.entries(replacements).map(([ph, val]) => ({
    replaceAllText: {
      containsText: { text: ph, matchCase: true },
      replaceText: String(val),
      pageObjectIds: newSlides, // only replace in the new duplicated slides
    },
  }));

  await slides.presentations.batchUpdate({
    presentationId,
    requestBody: { requests },
  });

  return `https://docs.google.com/presentation/d/${presentationId}/edit`;
}
