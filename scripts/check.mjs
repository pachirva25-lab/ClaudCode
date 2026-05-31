import { readFileSync } from "node:fs";
import { join } from "node:path";

const requiredFiles = [
  "index.html",
  "src/styles.css",
  "src/app.js",
  "README.md",
  "vba/SelectionMonitoring.bas",
  "vba/README.md",
  "GET_SELECTION_MONITORING_BAS.md",
  "Skills/VBA_MACRO_GUIDELINES.md",
  "Documentation/Техническое_описание.md",
  "Documentation/Инструкция_пользователя.md",
  "Домашнее задание/README.md",
];
const requiredText = [
  ["index.html", "СформироватьВыборкуУЗРКС"],
  ["index.html", "Отчет УЗРКС"],
  ["src/app.js", "2021–2026 Мониторинг УЗРКС"],
  ["src/app.js", "vba/SelectionMonitoring.bas"],
  ["src/styles.css", "--excel"],
  ["README.md", "Выборка из Мониторинга"],
  ["README.md", "Source\\Мониторинг УЗРКС"],
  ["vba/SelectionMonitoring.bas", "СформироватьВыборкуУЗРКС"],
  ["vba/SelectionMonitoring.bas", "SOURCE_SHEET As String = \"Закупщик\""],
  ["vba/SelectionMonitoring.bas", "OUTPUT_SHEET As String = \"Отчет УЗРКС\""],
  ["vba/SelectionMonitoring.bas", "PurchaseNumberColumns"],
  ["vba/README.md", "Alt + F11"],
  ["GET_SELECTION_MONITORING_BAS.md", "git pull origin main"],
  ["GET_SELECTION_MONITORING_BAS.md", "Welcome to pull requests!"],
  ["GET_SELECTION_MONITORING_BAS.md", "chatgpt.com/codex/cloud/tasks"],
  ["Skills/VBA_MACRO_GUIDELINES.md", "Option Explicit"],
  ["Documentation/Техническое_описание.md", "колонка 58"],
  ["Documentation/Инструкция_пользователя.md", "Alt + F8"],
  ["Домашнее задание/README.md", "скриншоты"],
];

for (const file of requiredFiles) {
  readFileSync(join(process.cwd(), file), "utf8");
}

for (const [file, text] of requiredText) {
  const content = readFileSync(join(process.cwd(), file), "utf8");
  if (!content.includes(text)) {
    throw new Error(`Expected ${file} to include ${text}`);
  }
}

console.log("Project files match the UZRKS monitoring selection implementation markers.");
