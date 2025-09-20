const { google } = require("googleapis");
const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");

// Path to your service account key JSON
const KEYFILEPATH = "service-account.json";
const SCOPES = ["https://www.googleapis.com/auth/drive.readonly"];

// Auth with service account
const auth = new google.auth.GoogleAuth({
  keyFile: KEYFILEPATH,
  scopes: SCOPES,
});

const drive = google.drive({ version: "v3", auth });

// List files in a folder
async function listFilesInFolder(folderId) {
  const res = await drive.files.list({
    q: `'${folderId}' in parents`,
    fields: "files(id, name, mimeType)",
    includeItemsFromAllDrives: true,
    supportsAllDrives: true,
  });
  return res.data.files || [];
}

// Download Excel file
async function downloadFile(fileId, fileName, folderPath) {
  const destPath = path.join(folderPath, fileName);
  fs.mkdirSync(path.dirname(destPath), { recursive: true });

  const dest = fs.createWriteStream(destPath);
  const res = await drive.files.get(
    { fileId, alt: "media" },
    { responseType: "stream" }
  );

  await new Promise((resolve, reject) => {
    res.data.on("end", resolve).on("error", reject).pipe(dest);
  });

  return destPath;
}

// Export Google Sheet as Excel
async function exportSheet(fileId, fileName, folderPath) {
  const safeName = fileName.endsWith(".xlsx") ? fileName : fileName + ".xlsx";
  const destPath = path.join(folderPath, safeName);
  fs.mkdirSync(path.dirname(destPath), { recursive: true });

  const dest = fs.createWriteStream(destPath);
  const res = await drive.files.export(
    {
      fileId,
      mimeType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    },
    { responseType: "stream" }
  );

  await new Promise((resolve, reject) => {
    res.data.on("end", resolve).on("error", reject).pipe(dest);
  });

  return destPath;
}

// Read Excel into JSON
function readExcel(filePath) {
  try {
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    return xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);
  } catch (e) {
    console.error("❌ Failed to read Excel:", filePath, e.message);
    return [];
  }
}

// Recursive fetch
async function fetchFolder(folderId, localPath) {
  console.log(`📂 Entering folder: ${folderId}, saving to: ${localPath}`);
  const files = await listFilesInFolder(folderId);

  for (let file of files) {
    console.log(`→ Found: ${file.name} (${file.mimeType})`);

    if (file.mimeType === "application/vnd.google-apps.folder") {
      // Subfolder → recurse
      const subFolderPath = path.join(localPath, file.name);
      await fetchFolder(file.id, subFolderPath);
    } else if (file.mimeType === "application/vnd.google-apps.spreadsheet") {
      console.log(`📥 Exporting Google Sheet: ${file.name}`);
      const filePath = await exportSheet(file.id, file.name, localPath);
      const data = readExcel(filePath);
      console.log(`✅ Data from ${file.name}:`, data.slice(0, 3));
    } else if (
      file.name.endsWith(".xlsx") ||
      file.name.endsWith(".xls")
    ) {
      console.log(`📥 Downloading Excel file: ${file.name}`);
      const filePath = await downloadFile(file.id, file.name, localPath);
      const data = readExcel(filePath);
      console.log(`✅ Data from ${file.name}:`, data.slice(0, 3));
    } else {
      console.log(`⏭ Skipping non-Excel file: ${file.name}`);
    }
  }
}

// Main
(async () => {
  try {
    const rootFolderId = "1aVKcBLFpDz0cwsBqa7efdOOEwoGiIFu7"; // replace with your shared folder ID
    const downloadRoot = path.join(__dirname, "downloads11");

    await fetchFolder(rootFolderId, downloadRoot);
    console.log("🎉 All Excel files fetched successfully.");
  } catch (err) {
    console.error("❌ Error:", err.message);
  }
})();
