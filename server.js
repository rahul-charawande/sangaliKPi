require("dotenv").config();
const express = require("express");
const cors = require("cors");
const bodyParser = require("body-parser");
const { google } = require("googleapis");
const xlsx = require("xlsx");

const app = express();
app.use(cors());
app.use(bodyParser.json());

// 🔑 Service Account credentials
const KEYFILEPATH = "service-account.json";
const SCOPES = ["https://www.googleapis.com/auth/drive.readonly"];

const authold = new google.auth.GoogleAuth({
  keyFile: KEYFILEPATH,
  scopes: SCOPES,
});


const auth = new google.auth.GoogleAuth({
  credentials: {
    type: process.env.GOOGLE_TYPE,
    project_id: process.env.GOOGLE_PROJECT_ID,
    private_key_id: process.env.GOOGLE_PRIVATE_KEY_ID,
    private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, "\n"), // Fix newline issue
    client_email: process.env.GOOGLE_CLIENT_EMAIL,
    client_id: process.env.GOOGLE_CLIENT_ID,
    auth_uri: process.env.GOOGLE_AUTH_URI,
    token_uri: process.env.GOOGLE_TOKEN_URI,
    auth_provider_x509_cert_url: process.env.GOOGLE_AUTH_PROVIDER_X509_CERT_URL,
    client_x509_cert_url: process.env.GOOGLE_CLIENT_X509_CERT_URL,
    universe_domain: process.env.GOOGLE_UNIVERSE_DOMAIN,
  },
  scopes: ["https://www.googleapis.com/auth/drive.readonly"],
});

const drive = google.drive({ version: "v3", auth });

// 🔹 API: Fetch specific sheet data from Google Drive file
app.post("/api/fetch-excel", async (req, res) => {
  try {
    const { url, sheetIndex = 0 } = req.body; // sheetIndex from client
    if (!url) return res.status(400).json({ error: "Google Drive URL required" });

    // Extract fileId from URL
    const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (!match) return res.status(400).json({ error: "Invalid Google Drive URL" });
    const fileId = match[1];

    // Get file metadata
    const file = await drive.files.get({ fileId, fields: "id, name, mimeType" });

    let buffer;
    if (file.data.mimeType === "application/vnd.google-apps.spreadsheet") {
      // Export Google Sheet as XLSX in memory
      const resExport = await drive.files.export(
        { fileId, mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
        { responseType: "arraybuffer" }
      );
      buffer = Buffer.from(resExport.data);
    } else if (
      file.data.mimeType === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" ||
      file.data.mimeType === "application/vnd.ms-excel"
    ) {
      // Direct Excel download
      const resDownload = await drive.files.get({ fileId, alt: "media" }, { responseType: "arraybuffer" });
      buffer = Buffer.from(resDownload.data);
    } else {
      return res.status(400).json({ error: "Not a Google Sheet or Excel file" });
    }

    // Read Excel into workbook
    const workbook = xlsx.read(buffer, { type: "buffer" });
    const sheetName = workbook.SheetNames[sheetIndex];
    if (!sheetName) return res.status(400).json({ error: "Invalid sheet index" });

    // Return raw sheet data as array of arrays (not converting to JSON keys/values)
    const sheet = workbook.Sheets[sheetName];
    //const rawData = xlsx.utils.sheet_to_json(sheet, { header: 1 }); // array of arrays

    // With the improved version:
    const rawData = [];
    const range = xlsx.utils.decode_range(sheet['!ref']);
    for (let R = range.s.r; R <= range.e.r; ++R) {
      const row = [];
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const cellAddress = xlsx.utils.encode_cell({ r: R, c: C });
        const cell = sheet[cellAddress];
        row.push({ cell: cellAddress, value: cell ? cell.v : null });
      }
      rawData.push(row);
    }


    return res.json({ sheetName, data: rawData });
  } catch (err) {
    console.error("❌ API Error:", err.message);
    return res.status(500).json({ error: "Failed to fetch Excel sheet data" });
  }
});

const PORT = process.env.PORT || 5000;



module.exports = app;
