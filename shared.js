const QUESTIONS_CACHE_KEY = "shikakuichi-questions-v1";
const QUESTIONS_CACHE_TTL_MS = 5 * 60 * 1000;

let questionsLoadPromise = null;

function readQuestionsCache(options = {}) {
  const { allowStale = false } = options;
  try {
    const raw = localStorage.getItem(QUESTIONS_CACHE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (!parsed.savedAt || !Array.isArray(parsed.questions)) return null;
    const age = Date.now() - parsed.savedAt;
    if (!allowStale && age > QUESTIONS_CACHE_TTL_MS) return null;
    return parsed.questions;
  } catch {
    return null;
  }
}

function writeQuestionsCache(questions) {
  localStorage.setItem(
    QUESTIONS_CACHE_KEY,
    JSON.stringify({ savedAt: Date.now(), questions }),
  );
}

async function fetchQuestionsFresh() {
  try {
    return await loadQuestionsFromCsv();
  } catch (error) {
    console.warn("問題CSV取得失敗、APIにフォールバック", error);
    if (!CONFIG.API_URL) throw error;
    const data = await loadDataFromApi();
    return data.questions || [];
  }
}

async function loadQuestionsOnly(options = {}) {
  const { forceRefresh = false } = options;

  if (!forceRefresh) {
    const cached = readQuestionsCache();
    if (cached) return cached;
  }

  if (!forceRefresh && questionsLoadPromise) {
    return questionsLoadPromise;
  }

  questionsLoadPromise = fetchQuestionsFresh()
    .then((questions) => {
      writeQuestionsCache(questions);
      return questions;
    })
    .finally(() => {
      questionsLoadPromise = null;
    });

  return questionsLoadPromise;
}

function prefetchQuestions() {
  const cached = readQuestionsCache();
  if (cached) return Promise.resolve(cached);
  return loadQuestionsOnly();
}

const QUESTION_KEYS = ["問題", "問題文"];
const ANSWER_KEYS = ["回答", "答え"];
const CHOICE_KEYS = ["選択肢1", "選択肢2", "選択肢3", "選択肢4"];
const AUTHOR_KEYS = ["作成者", "作者", "名前"];

function parseCsv(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    const next = text[i + 1];

    if (inQuotes) {
      if (ch === '"' && next === '"') {
        cell += '"';
        i += 1;
      } else if (ch === '"') {
        inQuotes = false;
      } else {
        cell += ch;
      }
      continue;
    }

    if (ch === '"') {
      inQuotes = true;
    } else if (ch === ",") {
      row.push(cell);
      cell = "";
    } else if (ch === "\n") {
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
    } else if (ch !== "\r") {
      cell += ch;
    }
  }

  if (cell.length > 0 || row.length > 0) {
    row.push(cell);
    rows.push(row);
  }

  return rows;
}

function normalizeHeader(value) {
  return String(value || "").trim();
}

function pickValue(record, keys) {
  for (const key of keys) {
    if (record[key] != null && String(record[key]).trim() !== "") {
      return String(record[key]).trim();
    }
  }
  return "";
}

function rowsToQuestions(rows) {
  if (rows.length < 2) return [];

  const headers = rows[0].map(normalizeHeader);
  const records = rows.slice(1).map((cells, index) => {
    const record = {};
    headers.forEach((header, col) => {
      record[header] = (cells[col] ?? "").trim();
    });
    return { record, row: index + 2 };
  });

  return records
    .map(({ record, row }) => ({
      row,
      question: pickValue(record, QUESTION_KEYS),
      answer: pickValue(record, ANSWER_KEYS),
      choices: CHOICE_KEYS.map((key) => record[key] ?? "").map((v) => v.trim()),
      author: pickValue(record, AUTHOR_KEYS),
    }))
    .filter((item) => item.question || item.choices.some(Boolean));
}

function escapeHtml(text) {
  return String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function shuffle(array) {
  const copy = [...array];
  for (let i = copy.length - 1; i > 0; i -= 1) {
    const j = Math.floor(Math.random() * (i + 1));
    [copy[i], copy[j]] = [copy[j], copy[i]];
  }
  return copy;
}

async function fetchCsv(url) {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`HTTP ${response.status}`);
  return parseCsv(await response.text());
}

async function loadDataFromCsv() {
  const questionsUrl = `https://docs.google.com/spreadsheets/d/${CONFIG.SHEET_ID}/export?format=csv&gid=${CONFIG.QUESTIONS_GID}`;
  const questionRows = await fetchCsv(questionsUrl);
  const questions = rowsToQuestions(questionRows);

  let users = [];
  if (CONFIG.API_URL) {
    try {
      const apiData = await loadDataFromApi();
      users = apiData.users;
    } catch {
      users = deriveUsersFromQuestions(questions);
    }
  } else {
    users = deriveUsersFromQuestions(questions);
  }

  return { users, questions };
}

async function loadApiAction(action) {
  if (!CONFIG.API_URL) throw new Error("API_URL が未設定です");
  const url = `${CONFIG.API_URL}?action=${action}&ts=${Date.now()}`;
  const response = await fetch(url);
  if (!response.ok) throw new Error(`API HTTP ${response.status}`);
  const data = await response.json();
  if (!data.ok) throw new Error(data.error || "API error");
  return data;
}

async function loadUsersFromSheetCsv() {
  const url = `https://docs.google.com/spreadsheets/d/${CONFIG.SHEET_ID}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent("ユーザー")}`;
  const response = await fetch(url);
  if (!response.ok) throw new Error(`HTTP ${response.status}`);
  const rows = parseCsv(await response.text());
  if (rows.length < 2) return [];
  return rows.slice(1).map((row) => String(row[0] || "").trim()).filter(Boolean);
}

async function loadUsersFromCsv() {
  if (CONFIG.USERS_GID) {
    const url = `https://docs.google.com/spreadsheets/d/${CONFIG.SHEET_ID}/export?format=csv&gid=${CONFIG.USERS_GID}`;
    const rows = await fetchCsv(url);
    if (rows.length < 2) return [];
    return rows.slice(1).map((row) => String(row[0] || "").trim()).filter(Boolean);
  }
  return loadUsersFromSheetCsv();
}

async function loadQuestionsFromCsv() {
  const url = `https://docs.google.com/spreadsheets/d/${CONFIG.SHEET_ID}/gviz/tq?tqx=out:csv&gid=${CONFIG.QUESTIONS_GID}&headers=1`;
  return rowsToQuestions(await fetchCsv(url));
}

async function loadUsersOnly() {
  try {
    return await loadUsersFromCsv();
  } catch (error) {
    console.warn("ユーザーCSV取得失敗、APIにフォールバック", error);
  }
  if (CONFIG.API_URL) {
    try {
      const data = await loadApiAction("getUsers");
      return data.users || [];
    } catch {
      const data = await loadDataFromApi();
      return data.users || [];
    }
  }
  return [];
}

async function loadDataFromApi() {
  if (!CONFIG.API_URL) throw new Error("API_URL が未設定です");
  const url = `${CONFIG.API_URL}?action=getData&ts=${Date.now()}`;
  const response = await fetch(url);
  if (!response.ok) throw new Error(`API HTTP ${response.status}`);
  const data = await response.json();
  if (!data.ok) throw new Error(data.error || "API error");
  return data;
}

function deriveUsersFromQuestions(questions) {
  return [...new Set(questions.map((q) => q.author).filter(Boolean))].sort();
}

async function apiPost(payload) {
  if (!CONFIG.API_URL) throw new Error("API_URL が未設定です。GAS をデプロイして config.js に URL を設定してください。");

  const response = await fetch(CONFIG.API_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify(payload),
  });

  const data = await response.json();
  if (!data.ok) throw new Error(data.error || "API error");
  return data;
}
