const sourceFiles = [
  {
    name: "2021–2026 Мониторинг УЗРКС*.xls*",
    type: "Книги Excel",
    role: "Исходные файлы мониторинга в папке Source/Мониторинг УЗРКС",
    status: "Источник",
  },
  {
    name: "Выборка из Мониторинга.xlsm",
    type: "Книга Excel с макросом",
    role: "Основной файл, куда импортируется VBA-модуль и выводится Отчет УЗРКС",
    status: "Целевой файл",
  },
  {
    name: "vba/SelectionMonitoring.bas",
    type: "VBA-модуль",
    role: "Макрос СформироватьВыборкуУЗРКС для выполнения ТЗ",
    status: "Реализовано",
  },
  {
    name: "Documentation/*",
    type: "Markdown-документация",
    role: "Техническое описание и инструкция пользователя",
    status: "Документация",
  },
  {
    name: "Skills/VBA_MACRO_GUIDELINES.md",
    type: "Общие указания",
    role: "Переиспользуемые подходы для будущих VBA-проектов",
    status: "Skills",
  },
];

const workbookRows = [
  {
    sheet: "Закупщик",
    action: "Прочитать исходные строки",
    details: "Лист в каждом файле 202х Мониторинг УЗРКС, объем до 3000 строк.",
    owner: "Исходные файлы",
  },
  {
    sheet: "Отчет УЗРКС",
    action: "Сформировать результат",
    details: "Лист в книге Выборка из Мониторинга.xlsm, заполняется найденными строками.",
    owner: "VBA-макрос",
  },
  {
    sheet: "Logs",
    action: "Зафиксировать ошибки",
    details: "Отсутствующие файлы, ошибки чтения, отсутствие листа или пустой источник.",
    owner: "VBA-макрос",
  },
  {
    sheet: "Домашнее задание",
    action: "Подготовить сдачу",
    details: "Список файлов и скриншотов, которые нужно приложить вручную.",
    owner: "Пользователь",
  },
];

const macroSteps = [
  "Пользователь запускает СформироватьВыборкуУЗРКС из книги Выборка из Мониторинга.xlsm.",
  "Макрос создает структуру папок Source, Output, Logs, Skills, Documentation, Backup и Домашнее задание.",
  "Пользователь вводит годы мониторинга: 2021, 2022, 2023, 2024, 2025, 2026.",
  "Макрос проверяет наличие выбранных файлов 202х Мониторинг УЗРКС*.xls*.",
  "Пользователь выбирает варианты поиска: номер закупки, наименование, номер ЦЗК, тип сделки.",
  "Макрос последовательно собирает значения критериев, ищет строки на листах Закупщик и переносит их в Отчет УЗРКС.",
  "В крайнюю колонку результата добавляется год мониторинга, ошибки пишутся в Logs.",
];

const deliverables = [
  {
    name: "vba/SelectionMonitoring.bas",
    description: "Основной VBA-модуль по ТЗ: годы 2021–2026, 4 варианта поиска, перенос в Отчет УЗРКС, журнал ошибок.",
  },
  {
    name: "Documentation/Инструкция_пользователя.md",
    description: "Пошаговая инструкция: куда положить файлы, как импортировать модуль, как запустить и проверить результат.",
  },
  {
    name: "Skills/VBA_MACRO_GUIDELINES.md",
    description: "Общие подходы разработки VBA-макросов для повторного использования в следующих проектах.",
  },
  {
    name: "Домашнее задание/README.md",
    description: "Список подготовленных файлов и скриншотов, которые пользователь должен сделать вручную для сдачи урока.",
  },
];

const fileRows = document.querySelector("#fileRows");
const workbookBody = document.querySelector("#workbookRows");
const macroList = document.querySelector("#macroSteps");
const copyButton = document.querySelector("#copySpecButton");
const statusMessage = document.querySelector("#statusMessage");
const deliverableList = document.querySelector("#deliverableList");

function renderSourceFiles() {
  fileRows.innerHTML = sourceFiles
    .map(
      (file) => `
        <tr>
          <td><strong>${file.name}</strong><span>${file.type}</span></td>
          <td>${file.role}</td>
          <td><span class="badge">${file.status}</span></td>
        </tr>
      `,
    )
    .join("");
}

function renderWorkbookRows() {
  workbookBody.innerHTML = workbookRows
    .map(
      (row) => `
        <tr>
          <td><strong>${row.sheet}</strong></td>
          <td>${row.action}</td>
          <td>${row.details}</td>
          <td>${row.owner}</td>
        </tr>
      `,
    )
    .join("");
}

function renderMacroSteps() {
  macroList.innerHTML = macroSteps.map((step) => `<li>${step}</li>`).join("");
}

function renderDeliverables() {
  deliverableList.innerHTML = deliverables
    .map((item) => `<li><strong>${item.name}</strong><span>${item.description}</span></li>`)
    .join("");
}

async function copySpec() {
  const text = [
    "Проект: Выборка из Мониторинга",
    "Цель: сформировать Отчет УЗРКС из файлов 2021–2026 Мониторинг УЗРКС.",
    "Основной макрос: vba/SelectionMonitoring.bas → СформироватьВыборкуУЗРКС.",
    "Варианты поиска: 1 номер закупки, 2 наименование закупки, 3 номер ЦЗК, 4 тип сделки.",
    "Проверка: импортировать модуль в Excel, положить файлы мониторинга в Source/Мониторинг УЗРКС, запустить Alt+F8.",
  ].join("\n");

  await navigator.clipboard.writeText(text);
  statusMessage.textContent = "Краткое ТЗ скопировано в буфер обмена.";
  setTimeout(() => {
    statusMessage.textContent = "";
  }, 3500);
}

renderSourceFiles();
renderWorkbookRows();
renderMacroSteps();
renderDeliverables();
copyButton.addEventListener("click", copySpec);
