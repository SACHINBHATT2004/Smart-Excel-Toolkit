const express = require("express");
const cors = require("cors");
const XLSX = require("xlsx");
const multer = require("multer");
const path = require("path");

const app = express();

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, "public")));

let data = [];

try {
  const workbook = XLSX.readFile("data.xlsx");
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  data = XLSX.utils.sheet_to_json(sheet);
} catch (err) {
  console.error("Error loading data.xlsx:", err.message);
  data = [];
}

app.get("/get-names", (req, res) => {
  const names = data.map(d => d["Name"]).filter(Boolean);
  res.json(names);
});

app.post("/generate-excel", (req, res) => {
  const inputNames = Array.isArray(req.body.names) ? req.body.names : [];

  const names = inputNames
    .map(n => String(n || "").trim().toLowerCase())
    .filter(Boolean);

  const result = data
    .filter(row => {
      const dbName = String(row["Name"] || "").toLowerCase().trim();
      return names.some(inputName => dbName.includes(inputName));
    })
    .map(row => ({
      Name: row["Name"] || "-",
      Email: row["Email"] || "-",
      "Mobile No.": row["Mobile No."] || "-",
      UTM: row["UTM "] || row["UTM"] || "-"
    }));

  res.json(result);
});

const storage = multer.memoryStorage();
const upload = multer({ storage });

app.post("/upload-excel", upload.single("file"), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: "No file uploaded" });
    }

    const workbook = XLSX.read(req.file.buffer, { type: "buffer" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet);

    const grouped = {};
    json.forEach(row => {
      const key = row["Approved By"] || "Unknown";
      if (!grouped[key]) grouped[key] = [];
      grouped[key].push(row);
    });

    res.json(grouped);
  } catch (err) {
    console.error("Upload Excel Error:", err.message);
    res.status(500).json({ error: "Invalid Excel file" });
  }
});

const PORT = 5000;
app.listen(PORT, () => {
  console.log(`Server running on ${PORT}`);
});