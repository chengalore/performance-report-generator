// index.js
import fs from "fs";
import { parse } from "csv-parse/sync";

const content = fs.readFileSync("./6fcdd173-52bc-4ba7-bc36-3592faaf4f27.csv", "utf-8");
const data = parse(content, { columns: true, skip_empty_lines: true });

// Build an HTML table
let html = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>CSV Preview</title>
  <style>
    table { border-collapse: collapse; width: 100%; }
    th, td { border: 1px solid #ccc; padding: 6px; text-align: left; }
    th { background: #f4f4f4; }
  </style>
</head>
<body>
  <h1>CSV Raw Data</h1>
  <table>
    <thead>
      <tr>${Object.keys(data[0]).map(col => `<th>${col}</th>`).join("")}</tr>
    </thead>
    <tbody>
      ${data.map(row => 
        `<tr>${Object.values(row).map(val => `<td>${val}</td>`).join("")}</tr>`
      ).join("")}
    </tbody>
  </table>
</body>
</html>
`;

// Save as HTML
fs.writeFileSync("report.html", html, "utf-8");
console.log("✅ Wrote report.html — open it in your browser.");
