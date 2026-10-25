const express = require("express");
const multer = require("multer");
const ExcelJS = require("exceljs");
const JSZip = require("jszip");
const fs = require("fs/promises");
const path = require("path");
const os = require("os");
const crypto = require("crypto");

const app = express();
const upload = multer({
  dest: path.join(os.tmpdir(), "excel-batch-uploads"),
  limits: { fileSize: 50 * 1024 * 1024 }
});

const TARGET_HEADER = "Fulfillment: Tracking Number";
const MIN_TRACKING = 1_000_000;
const MAX_TRACKING = 10_000_000;

function randomTrackingNumber() {
  // Uniform random integer in the inclusive range [MIN_TRACKING, MAX_TRACKING].
  return crypto.randomInt(MIN_TRACKING, MAX_TRACKING + 1);
}

function normalizeHeader(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "object" && value.richText) {
    return value.richText.map(part => part.text || "").join("").trim();
  }
  return String(value).trim();
}

function cellHasData(value) {
  if (value === null || value === undefined) return false;
  if (typeof value === "string" && value.trim() === "") return false;
  return true;
}

app.get("/", (_req, res) => {
  res.type("html").send(`<!doctype html>
<html lang="vi">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Excel Batch Tool</title>
  <style>
    body { font-family: system-ui, Arial, sans-serif; max-width: 760px; margin: 40px auto; padding: 0 18px; color: #202124; }
    main { border: 1px solid #ddd; border-radius: 16px; padding: 24px; }
    h1 { margin-top: 0; }
    label { display: block; font-weight: 600; margin: 18px 0 8px; }
    input, button { font: inherit; }
    input[type=file], input[type=number] { width: 100%; box-sizing: border-box; padding: 10px; }
    button { margin-top: 20px; padding: 12px 18px; border: 0; border-radius: 8px; cursor: pointer; background: #1459c7; color: white; font-weight: 600; }
    button:disabled { opacity: .6; cursor: wait; }
    .hint { color: #555; font-size: .95rem; line-height: 1.5; }
    #status { margin-top: 16px; white-space: pre-wrap; }
  </style>
</head>
<body>
  <main>
    <h1>Tạo file exel</h1>
    <p class="hint">Chọn file .xlsx, nhập số lượng file cần tạo(Max 500 file) rồi tải ZIP đầu ra.</p>
    <form id="form">
      <label for="file">File Excel gốc (.xlsx)</label>
      <input id="file" name="file" type="file" accept=".xlsx" required>
      <label for="count">Số lượng file cần tạo</label>
      <input id="count" name="count" type="number" min="1" max="500" value="100" required>
      <button id="submit" type="submit">Tạo file ZIP</button>
    </form>
    <div id="status" role="status"></div>
    <p class="hint">Cột cần có tiêu đề chính xác: <code>Fulfillment: Tracking Number</code>. Các ô trống trong cột này sẽ được giữ nguyên.</p>
  </main>
  <script>
    const form = document.getElementById("form");
    const button = document.getElementById("submit");
    const status = document.getElementById("status");
    form.addEventListener("submit", async (event) => {
      event.preventDefault();
      if (!form.file.files.length) return;
      button.disabled = true;
      status.textContent = "Đang tạo file, vui lòng đợi...";
      try {
        const data = new FormData(form);
        const response = await fetch("/generate", { method: "POST", body: data });
        if (!response.ok) {
          const error = await response.json().catch(() => ({ error: "Có lỗi xảy ra." }));
          throw new Error(error.error || "Không tạo được file.");
        }
        const blob = await response.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = response.headers.get("X-Output-Filename") || "Orders.zip";
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
        status.textContent = "Hoàn tất! File ZIP đã được tải xuống.";
      } catch (error) {
        status.textContent = "Lỗi: " + error.message;
      } finally {
        button.disabled = false;
      }
    });
  </script>
</body>
</html>`);
});

app.post("/generate", upload.single("file"), async (req, res) => {
  const uploadedPath = req.file && req.file.path;
  try {
    if (!req.file) return res.status(400).json({ error: "Vui lòng chọn file Excel." });
    if (!req.file.originalname.toLowerCase().endsWith(".xlsx")) {
      return res.status(400).json({ error: "Chỉ hỗ trợ file .xlsx." });
    }

    const count = Number.parseInt(req.body.count, 10);
    if (!Number.isInteger(count) || count < 1 || count > 500) {
      return res.status(400).json({ error: "Số lượng file phải từ 1 đến 500." });
    }

    // Load the source once to identify target columns and validate the header.
    const sourceWorkbook = new ExcelJS.Workbook();
    await sourceWorkbook.xlsx.readFile(uploadedPath);

    const targetColumns = new Map();
    for (const worksheet of sourceWorkbook.worksheets) {
      worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
          if (normalizeHeader(cell.value) === TARGET_HEADER) {
            if (!targetColumns.has(worksheet.name)) targetColumns.set(worksheet.name, []);
            targetColumns.get(worksheet.name).push({ headerRow: rowNumber, colNumber });
          }
        });
      });
    }

    if (targetColumns.size === 0) {
      return res.status(400).json({
        error: `Không tìm thấy tiêu đề cột "${TARGET_HEADER}" trong file.`
      });
    }

    const zip = new JSZip();
    for (let i = 1; i <= count; i++) {
      // Every output is freshly loaded from the original uploaded workbook.
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(uploadedPath);

      for (const [sheetName, headers] of targetColumns.entries()) {
        const worksheet = workbook.getWorksheet(sheetName);
        for (const { headerRow, colNumber } of headers) {
          for (let rowNumber = headerRow + 1; rowNumber <= worksheet.rowCount; rowNumber++) {
            const cell = worksheet.getRow(rowNumber).getCell(colNumber);
            if (cellHasData(cell.value)) {
              cell.value = randomTrackingNumber();
            }
          }
        }
      }

      const buffer = await workbook.xlsx.writeBuffer();
      const filename = `Order_${String(i).padStart(3, "0")}.xlsx`;
      zip.file(filename, buffer);
    }

    const zipBuffer = await zip.generateAsync({
      type: "nodebuffer",
      compression: "DEFLATE",
      compressionOptions: { level: 6 }
    });

    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", 'attachment; filename="Orders.zip"');
    res.setHeader("X-Output-Filename", "Orders.zip");
    res.send(zipBuffer);
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Không xử lý được file. Hãy kiểm tra file Excel và thử lại." });
  } finally {
    if (uploadedPath) await fs.unlink(uploadedPath).catch(() => {});
  }
});

const PORT = process.env.PORT || 3000;

app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on port ${PORT}`);
});