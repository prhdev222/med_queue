/**
 * Build script สำหรับ Vercel
 * ถ้ามี env QUEUE_APPSCRIPT_URL จะ inject default API เป็น /api/queue (proxy)
 * เพื่อไม่ให้ URL ของ Apps Script ไปอยู่ในโค้ดที่ส่งไปยัง browser
 */
const fs = require('fs');
const path = require('path');

const envUrl = process.env.QUEUE_APPSCRIPT_URL;
const src = path.join(__dirname, '..', 'queue-website.html');
const outDir = path.join(__dirname, '..', 'dist');
const outFile = path.join(outDir, 'index.html');

let html = fs.readFileSync(src, 'utf8');
if (envUrl) {
  html = html.replace(
    "window.QUEUE_DEFAULT_API = '';",
    "window.QUEUE_DEFAULT_API = '/api/queue';"
  );
  console.log('Injected default API: /api/queue (proxy)');
} else {
  console.log('No QUEUE_APPSCRIPT_URL — keeping default API empty (demo/local config)');
}

fs.mkdirSync(outDir, { recursive: true });
fs.writeFileSync(outFile, html);
console.log('Written:', outFile);

// คัดลอก NurseForm.html ไป dist ด้วย (ใช้คู่กับ queue-website)
const nurseFormSrc = path.join(__dirname, '..', 'NurseForm.html');
const nurseFormOut = path.join(outDir, 'NurseForm.html');
if (fs.existsSync(nurseFormSrc)) {
  let nurseHtml = fs.readFileSync(nurseFormSrc, 'utf8');
  if (envUrl) {
    // inject API URL ให้ NurseForm ใช้ proxy เหมือน queue-website
    nurseHtml = nurseHtml.replace(
      /var SCRIPT_URL = .+?;/,
      "var SCRIPT_URL = '/api/queue';"
    );
    console.log('Injected NurseForm API: /api/queue (proxy)');
  }
  fs.writeFileSync(nurseFormOut, nurseHtml);
  console.log('Written:', nurseFormOut);
}
