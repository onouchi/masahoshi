/**
 * masahoshi / shikakuichi 用スプレッドシート API
 *
 * セットアップ:
 * 1. スプレッドシートで 拡張機能 → Apps Script を開く
 * 2. このコードを貼り付け
 * 3. デプロイ → 新しいデプロイ → 種類: ウェブアプリ
 *    - 実行: 自分
 *    - アクセス: 全員
 * 4. 表示された URL を config.js の API_URL に設定
 *
 * シート構成:
 * - 「問題」(または先頭シート): 問題, 回答, 選択肢1〜4, 作成者
 * - 「ユーザー」(自動作成可): 名前
 */

const SPREADSHEET_ID = "1xFpRa3bgWPabH5dVZynlX19tFLnohjE15YcZs3jOqGQ";
const USERS_SHEET_NAME = "ユーザー";
const QUESTIONS_SHEET_NAME = "問題";

const QUESTION_HEADERS = ["問題", "回答", "選択肢1", "選択肢2", "選択肢3", "選択肢4", "作成者"];
const USER_HEADERS = ["名前"];

function doGet(e) {
  return handleRequest_(e, "GET");
}

function doPost(e) {
  return handleRequest_(e, "POST");
}

function handleRequest_(e, method) {
  try {
    const action = e.parameter.action || (e.postData ? JSON.parse(e.postData.contents).action : "");

    if (action === "getData") {
      return jsonResponse_(getData_());
    }

    if (action === "getUsers") {
      return jsonResponse_(getUsers_());
    }

    if (action === "getQuestions") {
      return jsonResponse_(getQuestions_());
    }

    if (method === "POST") {
      const body = JSON.parse(e.postData.contents);
      switch (body.action) {
        case "addUser":
          return jsonResponse_(addUser_(body.name));
        case "addQuestion":
          return jsonResponse_(addQuestion_(body));
        case "updateQuestion":
          return jsonResponse_(updateQuestion_(body));
        case "deleteQuestion":
          return jsonResponse_(deleteQuestion_(body));
        default:
          throw new Error("不明な action: " + body.action);
      }
    }

    throw new Error("不明な action: " + action);
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err.message || err) });
  }
}

function getSpreadsheet_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getQuestionsSheet_(ss) {
  const sheet = ss.getSheetByName(QUESTIONS_SHEET_NAME) || ss.getSheets()[0];
  ensureQuestionHeaders_(sheet);
  return sheet;
}

function getUsersSheet_(ss, createIfMissing) {
  let sheet = ss.getSheetByName(USERS_SHEET_NAME);
  if (!sheet && createIfMissing) {
    sheet = ss.insertSheet(USERS_SHEET_NAME);
    sheet.getRange(1, 1, 1, USER_HEADERS.length).setValues([USER_HEADERS]);
  }
  return sheet;
}

function ensureQuestionHeaders_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn() || QUESTION_HEADERS.length).getValues()[0];
  const normalized = headers.map(function (h) { return String(h || "").trim(); });

  if (normalized[0] === "") {
    sheet.getRange(1, 1, 1, QUESTION_HEADERS.length).setValues([QUESTION_HEADERS]);
    return;
  }

  if (normalized.indexOf("作成者") === -1) {
    const col = normalized.length + 1;
    sheet.getRange(1, col).setValue("作成者");
  }
}

function getData_() {
  const ss = getSpreadsheet_();
  const questionsSheet = getQuestionsSheet_(ss);
  const usersSheet = getUsersSheet_(ss, true);

  const users = readUsers_(usersSheet);
  const questions = readQuestions_(questionsSheet);

  return { ok: true, users: users, questions: questions };
}

function getUsers_() {
  const ss = getSpreadsheet_();
  const usersSheet = getUsersSheet_(ss, false);
  return { ok: true, users: usersSheet ? readUsers_(usersSheet) : [] };
}

function getQuestions_() {
  const ss = getSpreadsheet_();
  return { ok: true, questions: readQuestions_(getQuestionsSheet_(ss)) };
}

function readUsers_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow, 1).getValues();
  return values
    .map(function (row) { return String(row[0] || "").trim(); })
    .filter(Boolean);
}

function readQuestions_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
    return String(h || "").trim();
  });

  const col = function (name, aliases) {
    const names = [name].concat(aliases || []);
    for (var i = 0; i < names.length; i += 1) {
      const idx = headers.indexOf(names[i]);
      if (idx !== -1) return idx;
    }
    return -1;
  };

  const qCol = col("問題", ["問題文"]);
  const aCol = col("回答", ["答え"]);
  const authorCol = col("作成者", ["作者", "名前"]);
  const choiceCols = [
    col("選択肢1"),
    col("選択肢2"),
    col("選択肢3"),
    col("選択肢4"),
  ];

  const values = sheet.getRange(2, 1, lastRow, lastCol).getValues();
  const questions = [];

  values.forEach(function (row, index) {
    const question = qCol >= 0 ? String(row[qCol] || "").trim() : "";
    const answer = aCol >= 0 ? String(row[aCol] || "").trim() : "";
    const choices = choiceCols.map(function (c) {
      return c >= 0 ? String(row[c] || "").trim() : "";
    });
    const author = authorCol >= 0 ? String(row[authorCol] || "").trim() : "";

    if (!question && choices.every(function (c) { return !c; })) return;

    questions.push({
      row: index + 2,
      question: question,
      answer: answer,
      choices: choices,
      author: author,
    });
  });

  return questions;
}

function addUser_(name) {
  const trimmed = String(name || "").trim();
  if (!trimmed) throw new Error("名前を入力してください");

  const ss = getSpreadsheet_();
  const sheet = getUsersSheet_(ss, true);
  const users = readUsers_(sheet);

  if (users.indexOf(trimmed) !== -1) throw new Error("同じ名前が既にあります");

  sheet.appendRow([trimmed]);
  return { ok: true, users: readUsers_(sheet) };
}

function questionFromBody_(body, row) {
  return {
    row: row,
    question: String(body.question || "").trim(),
    answer: String(body.answer || "").trim(),
    choices: [
      String(body.choices?.[0] || "").trim(),
      String(body.choices?.[1] || "").trim(),
      String(body.choices?.[2] || "").trim(),
      String(body.choices?.[3] || "").trim(),
    ],
    author: String(body.author || "").trim(),
  };
}

function addQuestion_(body) {
  const author = String(body.author || "").trim();
  if (!author) throw new Error("作成者を選んでください");

  const ss = getSpreadsheet_();
  const sheet = getQuestionsSheet_(ss);
  assertAuthorExists_(ss, author);

  const row = [
    String(body.question || "").trim(),
    String(body.answer || "").trim(),
    String(body.choices?.[0] || "").trim(),
    String(body.choices?.[1] || "").trim(),
    String(body.choices?.[2] || "").trim(),
    String(body.choices?.[3] || "").trim(),
    author,
  ];

  sheet.appendRow(row);
  return { ok: true, question: questionFromBody_(body, sheet.getLastRow()) };
}

function updateQuestion_(body) {
  const author = String(body.author || "").trim();
  const row = Number(body.row);
  if (!author || !row) throw new Error("author と row が必要です");

  const ss = getSpreadsheet_();
  const sheet = getQuestionsSheet_(ss);
  assertAuthorOwnsRow_(sheet, row, author);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function (h) {
    return String(h || "").trim();
  });

  setCellByHeader_(sheet, row, headers, "問題", String(body.question || "").trim(), ["問題文"]);
  setCellByHeader_(sheet, row, headers, "回答", String(body.answer || "").trim(), ["答え"]);
  setCellByHeader_(sheet, row, headers, "選択肢1", String(body.choices?.[0] || "").trim());
  setCellByHeader_(sheet, row, headers, "選択肢2", String(body.choices?.[1] || "").trim());
  setCellByHeader_(sheet, row, headers, "選択肢3", String(body.choices?.[2] || "").trim());
  setCellByHeader_(sheet, row, headers, "選択肢4", String(body.choices?.[3] || "").trim());

  return { ok: true, question: questionFromBody_(body, row) };
}

function deleteQuestion_(body) {
  const author = String(body.author || "").trim();
  const row = Number(body.row);
  if (!author || !row) throw new Error("author と row が必要です");

  const ss = getSpreadsheet_();
  const sheet = getQuestionsSheet_(ss);
  assertAuthorOwnsRow_(sheet, row, author);
  sheet.deleteRow(row);

  return { ok: true };
}

function assertAuthorExists_(ss, author) {
  const users = readUsers_(getUsersSheet_(ss, false) || getUsersSheet_(ss, true));
  if (users.indexOf(author) === -1) throw new Error("ユーザー一覧にない名前です");
}

function assertAuthorOwnsRow_(sheet, row, author) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function (h) {
    return String(h || "").trim();
  });
  const authorCol = headers.indexOf("作成者");
  if (authorCol === -1) throw new Error("作成者列がありません");

  const owner = String(sheet.getRange(row, authorCol + 1).getValue() || "").trim();
  if (owner !== author) throw new Error("この問題は編集できません");
}

function setCellByHeader_(sheet, row, headers, name, value, aliases) {
  const names = [name].concat(aliases || []);
  for (var i = 0; i < names.length; i += 1) {
    const idx = headers.indexOf(names[i]);
    if (idx !== -1) {
      sheet.getRange(row, idx + 1).setValue(value);
      return;
    }
  }
}

function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
