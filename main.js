// ============================================
// JIRA Zephyr Test Steps Export to CSV/ODS
// ============================================
// Выгружает шаги из тестовых кейсов, привязанных к задачам по JQL,
// через Zephyr for Jira Server API
import { configDotenv } from "dotenv";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import * as xlsx from "xlsx";

configDotenv({ path: ".env" });

function throwEnvError() {
  throw new Error("Проверьте соответствие .env файлу .env.example");
}

if (!fs.existsSync("output")) {
  fs.mkdirSync("output");
}

// ============================================
// КОНФИГУРАЦИЯ
// ============================================
const config = {
  // URL вашего Jira сервера
  jiraUrl: process.env.jira_url || throwEnvError(),

  // Учетные данные (заполните своими данными)
  username: process.env.jira_username || throwEnvError(),
  password: process.env.jira_password || throwEnvError(),

  // JQL фильтр для поиска задач
  jql: "((fixVersion IN (13581,13582,13583,13584,13585,13567,13586,13587,13588,13589,13590,13591)))",
  // jql: "((fixVersion IN (13522,13523,13525,13526,13527,13528,13529,13530,13531,13532,13533,13534,13569,13573,13575,13577,13578,13580,13581,13582,13583,13584,13585,13567,13586,13587,13588,13589,13590,13591,13612,13613,13615,13616,13617,13618,13620,13621,13626,13629,13630,13631,13632,13633,13634,13636,13640,13642,13644,13645,13647,13648,13649,13650,13651)))",

  // Максимальное количество задач за один запрос (максимум 1000)
  maxResults: 100,

  // Имя выходного CSV файла
  outputFile: `output/${Date.now()}_test_steps.csv`,

  // Имя выходного ODS файла
  outputOdsFile: `output/${Date.now()}_test_steps.ods`,

  // Формат экспорта: "ods" или "csv"
  exportFormat: "ods",
};

// ============================================
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
// ============================================

// Получение __dirname для ES модулей
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Логирование с временной меткой
function log(message) {
  const timestamp = new Date().toISOString();
  console.log(`[${timestamp}] ${message}`);
}

// Создание заголовков для Basic Auth
function getAuthHeaders() {
  const credentials = `${config.username}:${config.password}`;
  const base64Credentials = Buffer.from(credentials).toString("base64");
  return {
    Authorization: `Basic ${base64Credentials}`,
    "Content-Type": "application/json",
  };
}

// Выполнение HTTP запроса к Jira API
async function jiraRequest(endpoint, params = {}) {
  const url = `${config.jiraUrl}${endpoint}`;

  const options = {
    method: "GET",
    headers: getAuthHeaders(),
    ...params,
  };

  try {
    const response = await fetch(url, options);

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(
        `Jira API error ${response.status}: ${response.statusText}\n${errorText}`,
      );
    }

    return await response.json();
  } catch (error) {
    log(`Ошибка при запросе к ${url}: ${error.message}`);
    throw error;
  }
}

// ============================================
// ОСНОВНАЯ ЛОГИКА
// ============================================

// Получить задачи по JQL
async function getIssuesByJql(jql, startAt = 0) {
  log(`Получение задач по JQL (startAt=${startAt})...`);

  const endpoint = `/rest/api/2/search?jql=${encodeURIComponent(jql)}&startAt=${startAt}&maxResults=${config.maxResults}&expand=changelog`;
  return await jiraRequest(endpoint);
}

// Получить testResult через traceLinks для задачи
async function getTestResultsByIssue(issueId) {
  const endpoint = `/rest/tests/1.0/issue/${issueId}/tracelinks?maxResults=50&startAt=0`;
  try {
    const data = await jiraRequest(endpoint);
    const testResults = [];

    // Проходим по всем блокам и собираем testResult
    const blocks = [
      data.testResult?.blocks?.traceLinks,
      data.testRun?.blocks?.traceLinks,
      data.testCase?.blocks?.traceLinks,
    ];

    for (const block of blocks) {
      if (block && Array.isArray(block)) {
        for (const link of block) {
          if (link.testResult && link.testResult.key) {
            testResults.push(link.testResult);
          }
        }
      }
    }

    return testResults;
  } catch (error) {
    log(
      `Ошибка при получении testResult для задачи ${issueId}: ${error.message}`,
    );
    return [];
  }
}

// Получить шаги (testScriptResults) из testResult
async function getTestScriptResults(testResultKey) {
  // Поля для запроса: testScriptResults с нужными атрибутами
  const fields =
    "id,environment(id,name),automated,estimatedTime,customFieldValues,scenarioResultIds,executionTime,iterationId,plannedStartDate,plannedEndDate,actualStartDate,actualEndDate,executionDate,jiraVersionId,comment,userKey,assignedTo,testResultStatusId,testCase(id,key,name,projectId,projectKey,objective,precondition,componentId),testRun(projectId),testScriptResults(id,testResultStatusId,executionDate,comment,index,description,expectedResult,testData,traceLinks,attachments,sourceScriptType,parameterSetId,customFieldValues,stepAttachmentsMapping,reflectRef),traceLinks,attachments,labels,customFieldValues";

  const endpoint = `/rest/tests/1.0/testresult/${testResultKey}?fields=${fields}`;
  try {
    const data = await jiraRequest(endpoint);
    return data.testScriptResults || [];
  } catch (error) {
    log(
      `Ошибка при получении testScriptResults для ${testResultKey}: ${error.message}`,
    );
    return [];
  }
}

// ============================================
// XLSX/ODS ФУНКЦИИ
// ============================================

// Заголовки для таблицы
const TABLE_HEADERS = [
  "Task Key",
  "Test Result Key",
  "Step Index",
  "Action (Description)",
  "Expected Result",
];

// Создание новой рабочей книги xlsx
function createWorkbook() {
  return xlsx.utils.book_new();
}

// Добавление листа в рабочую книгу
function addSheetToWorkbook(workbook, data, sheetName = "Test Steps") {
  const worksheet = xlsx.utils.aoa_to_sheet([TABLE_HEADERS, ...data]);
  xlsx.utils.book_append_sheet(workbook, worksheet, sheetName);
  return worksheet;
}

// Запись книги в ODS файл
function writeOdsFile(workbook, filePath) {
  xlsx.writeFile(workbook, filePath, { bookType: "ods" });
}

// Запись книги в CSV файл
function writeCsvFile(workbook, filePath) {
  xlsx.writeFile(workbook, filePath, { bookType: "csv" });
}

// Форматирование значения для экспорта (приведение к строке)
function formatFieldValue(value) {
  if (value === null || value === undefined) {
    return "";
  }
  return String(value)
    .replace(/<br ?\/>/g, "\n")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/, ">");
}

// Главная функция
async function main() {
  log("=== Запуск экспорта шагов тестов из Jira Zephyr ===");
  log(`JIRA URL: ${config.jiraUrl}`);
  log(`JQL: ${config.jql}`);
  log(`Формат экспорта: ${config.exportFormat}`);

  const outputFileName =
    config.exportFormat === "ods" ? config.outputOdsFile : config.outputFile;
  const outputFilePath = path.join(__dirname, outputFileName);

  log(`Выходной файл: ${outputFilePath}`);

  // Проверка конфигурации
  if (
    config.username === "YOUR_USERNAME" ||
    config.password === "YOUR_PASSWORD"
  ) {
    log("ОШИБКА: Пожалуйста, укажите свои учетные данные в переменной config!");
    log('Откройте файл main.js и измените секцию "КОНФИГУРАЦИЯ"');
    process.exit(1);
  }

  let totalIssuesProcessed = 0;
  let totalTestResults = 0;
  let totalSteps = 0;

  // Массив для сбора всех строк данных
  const allDataRows = [];

  // Получаем все задачи (пагинация)
  let startAt = 0;
  let allIssuesFetched = false;

  while (!allIssuesFetched) {
    try {
      const searchResult = await getIssuesByJql(config.jql, startAt);

      if (!searchResult.issues || searchResult.issues.length === 0) {
        log("Нет больше задач для обработки.");
        break;
      }

      log(
        `Получено задач: ${searchResult.issues.length} (всего обработано: ${startAt + searchResult.issues.length})`,
      );

      // Обрабатываем каждую задачу
      for (const issue of searchResult.issues) {
        const issueKey = issue.key;
        const issueId = issue.id;

        log(`Обработка задачи: ${issueKey} (ID: ${issueId})`);

        // Получаем testResult через traceLinks
        const testResults = await getTestResultsByIssue(issueId);
        totalTestResults += testResults.length;

        if (testResults.length === 0) {
          log(`  -> Нет связанных тестов для задачи ${issueKey}`);
          continue;
        }

        log(`  -> Найдено testResult: ${testResults.length}`);

        // Для каждого testResult получаем шаги
        for (const testResult of testResults) {
          const trKey = testResult.key;
          log(`    -> Получение шагов для ${trKey}...`);

          const scriptResults = await getTestScriptResults(trKey);

          if (scriptResults.length === 0) {
            log(`      -> Нет шагов для ${trKey}`);
            continue;
          }

          // Сортируем шаги по index по возрастанию
          scriptResults.sort((a, b) => {
            const idxA = a.index || 0;
            const idxB = b.index || 0;
            return idxA - idxB;
          });

          log(`      -> Найдено шагов: ${scriptResults.length}`);

          // Добавляем шаги в массив данных
          for (const script of scriptResults) {
            const row = [
              issueKey, // Task Key
              trKey, // Test Result Key
              script.index + 1, // Step Index
              formatFieldValue(script.description), // Action
              formatFieldValue(script.expectedResult), // Expected Result
            ];
            allDataRows.push(row);
            totalSteps++;
          }
        }

        totalIssuesProcessed++;
      }

      // Проверяем, есть ли еще задачи
      if (startAt + searchResult.issues.length >= searchResult.total) {
        allIssuesFetched = true;
      } else {
        startAt += searchResult.issues.length;
      }

      // Небольшая задержка между запросами (чтобы не перегружать сервер)
      await new Promise((resolve) => setTimeout(resolve, 500));
    } catch (error) {
      log(`Критическая ошибка при обработке: ${error.message}`);
      break;
    }
  }

  // Сортировка всех строк по Task Key + Test Result Key + Step Index (число)
  log("Сортировка результатов...");
  allDataRows.sort((a, b) => {
    // Сортировка по Task Key
    if (a[0] !== b[0]) return a[0].localeCompare(b[0]);
    // Затем по Test Result Key
    if (a[1] !== b[1]) return a[1].localeCompare(b[1]);
    // Внутри группы по Step Index как числу
    const idxA = Number(a[2]) || 0;
    const idxB = Number(b[2]) || 0;
    return idxA - idxB;
  });

  log("Создание рабочей книги...");
  const workbook = createWorkbook();
  addSheetToWorkbook(workbook, allDataRows);

  // Записываем файл в зависимости от формата
  if (config.exportFormat === "ods") {
    writeOdsFile(workbook, outputFilePath);
  } else {
    writeCsvFile(workbook, outputFilePath);
  }

  log(`Файл сохранен: ${outputFilePath}`);

  // Итоговая статистика
  log("=== Экспорт завершен ===");
  log(`Обработано задач: ${totalIssuesProcessed}`);
  log(`Найдено testResult: ${totalTestResults}`);
  log(`Всего шагов экспортировано: ${totalSteps}`);
  log(`Результат сохранен в: ${outputFilePath}`);
}

// Запуск
main().catch((error) => {
  console.error(error);
  log(`Фатальная ошибка: ${error.message}`);
  process.exit(1);
});
