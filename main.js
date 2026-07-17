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
  jql: process.env.jira_jql || throwEnvError(),

  outputFileName: `output/${process.env.outpit_file_name ? `${process.env.outpit_file_name}` : `${Date.now()}`}.ods`,

  // Максимальное количество задач за один запрос (максимум 1000)
  maxResults: parseInt(`${process.env.jira_task_load_chunk}`) || 100,
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

  // Добавляем объединения для одинаковых значений
  const merges = calculateMerges(data);
  if (merges.length > 0) {
    if (!worksheet["!merges"]) worksheet["!merges"] = [];
    worksheet["!merges"].push(...merges);
  }

  xlsx.utils.book_append_sheet(workbook, worksheet, sheetName);
  return worksheet;
}

// Вычисление диапазонов для объединения ячеек
// data - массив строк данных (без заголовков)
// Заголовки добавляются как первая строка в worksheet, поэтому индексы данных сдвигаются на +1
function calculateMerges(data) {
  const merges = [];

  // Для столбца Task Key (0)
  let begin = 0;
  for (let i = 1; i <= data.length; i++) {
    if (i === data.length || data[i][0] !== data[begin][0]) {
      if (i - begin > 1) {
        // begin + 1 и i - 1 + 1 - сдвиг на 1 строку для заголовков
        merges.push({ s: { r: begin + 1, c: 0 }, e: { r: i - 1 + 1, c: 0 } });
      }
      begin = i;
    }
  }

  // Для столбца Test Result Key (1)
  begin = 0;
  for (let i = 1; i <= data.length; i++) {
    if (i === data.length || data[i][1] !== data[begin][1]) {
      if (i - begin > 1) {
        merges.push({ s: { r: begin + 1, c: 1 }, e: { r: i - 1 + 1, c: 1 } });
      }
      begin = i;
    }
  }

  // Сортировка объединений для корректной работы
  merges.sort((a, b) => {
    if (a.s.r !== b.s.r) return a.s.r - b.s.r;
    if (a.s.c !== b.s.c) return a.s.c - b.s.c;
    if (a.e.r !== b.e.r) return a.e.r - b.e.r;
    return a.e.c - b.e.c;
  });

  return merges;
}

// Запись книги в ODS файл
function writeOdsFile(workbook, filePath) {
  xlsx.writeFile(workbook, filePath, { bookType: "ods" });
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

  const outputFileName = config.outputFileName;
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
  writeOdsFile(workbook, outputFilePath);

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
