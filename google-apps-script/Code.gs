function doGet() {
  return jsonResponse({
    ok: true,
    message: "BVD API is running.",
  });
}

const DATABASE_SPREADSHEET_NAME = "bvd_database";
const USER_SHEET_NAME = "user";
const COMMENT_SHEET_NAME = "comment";
const DELIVER_SHEET_NAME = "deliver";
const SOLAPI_API_KEY_PROPERTY = "SOLAPI_API_KEY";
const SOLAPI_API_SECRET_PROPERTY = "SOLAPI_API_SECRET";
const SOLAPI_FROM_PROPERTY = "SOLAPI_FROM";
const ADMIN_PHONE_PROPERTY = "ADMIN_PHONE";

function doPost(event) {
  try {
    const body = JSON.parse(event.postData.contents || "{}");
    const action = String(body.action || "").trim();

    if (action === "signup" || action === "login") {
      return handleAuth(action, body);
    }

    if (action === "list_comments") {
      return handleListComments(body);
    }

    if (action === "add_comment") {
      return handleAddComment(body);
    }

    if (action === "list_deliveries") {
      return handleListDeliveries(body);
    }

    if (action === "add_delivery") {
      return handleAddDelivery(body);
    }

    return jsonResponse({
      ok: false,
      message: "알 수 없는 요청입니다.",
    });
  } catch (error) {
    return jsonResponse({
      ok: false,
      message: error.message || "서버 오류가 발생했습니다.",
    });
  }
}

function handleAuth(action, body) {
  const userId = String(body.user_id || "").trim();
  const normalizedUserId = normalizeUserId(userId);
  const password = String(body.password || "");

  if (!userId || !password) {
    return jsonResponse({
      ok: false,
      message: "아이디와 비밀번호를 입력해주세요.",
    });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);

  try {
    const users = getSheetData(USER_SHEET_NAME, ["user_id", "password"]);
    const existingRow = users.rows
      .slice(1)
      .find(function (row) {
        return normalizeUserId(row[users.columns.user_id]) === normalizedUserId;
      });

    if (action === "signup") {
      if (existingRow) {
        return jsonResponse({
          ok: false,
          message: "이미 가입된 아이디입니다.",
        });
      }

      const nextRow = Array(users.headers.length).fill("");
      nextRow[users.columns.user_id] = userId;
      nextRow[users.columns.password] = hashPassword(password);
      users.sheet.appendRow(nextRow);

      return jsonResponse({
        ok: true,
        message: "회원가입이 완료되었습니다.",
      });
    }

    if (!existingRow || !passwordMatches(existingRow[users.columns.password], password)) {
      return jsonResponse({
        ok: false,
        message: "아이디 또는 비밀번호가 맞지 않습니다.",
      });
    }

    return jsonResponse({
      ok: true,
      message: "로그인되었습니다.",
    });
  } finally {
    lock.releaseLock();
  }
}

function handleListComments(body) {
  const cardIds = Array.isArray(body.card_ids)
    ? body.card_ids.map(String)
    : [String(body.card_id || "")];
  const uniqueCardIds = cardIds.filter(Boolean).filter(function (cardId, index, array) {
    return array.indexOf(cardId) === index;
  });

  if (!uniqueCardIds.length) {
    return jsonResponse({
      ok: true,
      comments: {},
    });
  }

  return jsonResponse({
    ok: true,
    comments: getCommentsByCardIds(uniqueCardIds),
  });
}

function handleAddComment(body) {
  const cardId = String(body.card_id || "").trim();
  const userId = String(body.user_id || "").trim();
  const comment = String(body.comment || "").trim();

  if (!cardId || !userId || !comment) {
    return jsonResponse({
      ok: false,
      message: "댓글을 입력해주세요.",
    });
  }

  if (!userExists(userId)) {
    return jsonResponse({
      ok: false,
      message: "로그인이 필요합니다.",
    });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);

  try {
    const comments = getSheetData(COMMENT_SHEET_NAME, ["timestamp", "card_id", "user_id", "comment"]);
    const nextRow = Array(comments.headers.length).fill("");

    nextRow[comments.columns.timestamp] = new Date();
    nextRow[comments.columns.card_id] = cardId;
    nextRow[comments.columns.user_id] = userId;
    nextRow[comments.columns.comment] = comment;
    comments.sheet.appendRow(nextRow);

    safeSendAdminSms(
      [
        "[비버댐] 새 댓글",
        "작성자: " + userId,
        "카드: " + cardId,
        "내용: " + comment,
      ].join("\n")
    );
  } finally {
    lock.releaseLock();
  }

  return jsonResponse({
    ok: true,
    comments: getCommentsByCardIds([cardId])[cardId] || [],
  });
}

function getCommentsByCardIds(cardIds) {
  const comments = getSheetData(COMMENT_SHEET_NAME, ["timestamp", "card_id", "user_id", "comment"]);
  const lookup = {};

  cardIds.forEach(function (cardId) {
    lookup[cardId] = [];
  });

  comments.rows.slice(1).forEach(function (row) {
    const cardId = String(row[comments.columns.card_id] || "");
    if (!Object.prototype.hasOwnProperty.call(lookup, cardId)) return;

    lookup[cardId].push({
      timestamp: row[comments.columns.timestamp],
      card_id: cardId,
      user_id: String(row[comments.columns.user_id] || ""),
      comment: String(row[comments.columns.comment] || ""),
    });
  });

  Object.keys(lookup).forEach(function (cardId) {
    lookup[cardId].sort(function (a, b) {
      return new Date(a.timestamp).getTime() - new Date(b.timestamp).getTime();
    });
  });

  return lookup;
}

function handleListDeliveries(body) {
  const teamId = String(body.team_id || "").trim();

  if (!teamId) {
    return jsonResponse({
      ok: true,
      deliveries: [],
    });
  }

  return jsonResponse({
    ok: true,
    deliveries: getDeliveriesByTeamId(teamId),
  });
}

function handleAddDelivery(body) {
  const userId = String(body.user_id || "").trim();
  const teamId = String(body.team_id || "").trim();
  const comment = String(body.comment || "").trim();
  const address = String(body.address || "").trim();

  if (!userId || !teamId || !comment || !address) {
    return jsonResponse({
      ok: false,
      message: "배달 메모와 주소를 입력해주세요.",
    });
  }

  if (!userExists(userId)) {
    return jsonResponse({
      ok: false,
      message: "로그인이 필요합니다.",
    });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);

  try {
    const deliveries = getSheetData(DELIVER_SHEET_NAME, [
      "timestamp",
      "card_id",
      "user_id",
      "team_id",
      "comment",
      "address",
    ]);
    const nextRow = Array(deliveries.headers.length).fill("");

    nextRow[deliveries.columns.timestamp] = new Date();
    nextRow[deliveries.columns.card_id] = "deliver:" + Utilities.getUuid();
    nextRow[deliveries.columns.user_id] = userId;
    nextRow[deliveries.columns.team_id] = teamId;
    nextRow[deliveries.columns.comment] = comment;
    nextRow[deliveries.columns.address] = address;
    deliveries.sheet.appendRow(nextRow);

    safeSendAdminSms(
      [
        "[비버댐] 새 배달 메모",
        "팀: " + teamId,
        "작성자: " + userId,
        "내용: " + comment,
        "주소: " + address,
      ].join("\n")
    );
  } finally {
    lock.releaseLock();
  }

  return jsonResponse({
    ok: true,
    deliveries: getDeliveriesByTeamId(teamId),
  });
}

function getDeliveriesByTeamId(teamId) {
  const deliveries = getSheetData(DELIVER_SHEET_NAME, [
    "timestamp",
    "card_id",
    "user_id",
    "team_id",
    "comment",
    "address",
  ]);
  const normalizedTeamId = String(teamId || "").trim().toUpperCase();

  return deliveries.rows
    .slice(1)
    .filter(function (row) {
      return String(row[deliveries.columns.team_id] || "").trim().toUpperCase() === normalizedTeamId;
    })
    .map(function (row) {
      return {
        timestamp: row[deliveries.columns.timestamp],
        card_id: String(row[deliveries.columns.card_id] || ""),
        user_id: String(row[deliveries.columns.user_id] || ""),
        team_id: String(row[deliveries.columns.team_id] || ""),
        comment: String(row[deliveries.columns.comment] || ""),
        address: String(row[deliveries.columns.address] || ""),
      };
    })
    .sort(function (a, b) {
      return new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime();
    });
}

function userExists(userId) {
  const users = getSheetData(USER_SHEET_NAME, ["user_id", "password"]);
  const normalizedUserId = normalizeUserId(userId);

  return users.rows.slice(1).some(function (row) {
    return normalizeUserId(row[users.columns.user_id]) === normalizedUserId;
  });
}

function getSheetData(sheetName, requiredHeaders) {
  const spreadsheet = getDatabaseSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(sheetName + " 시트를 찾을 수 없습니다.");
  }

  const rows = sheet.getDataRange().getValues();
  const headers = rows[0].map(String);
  const columns = {};

  requiredHeaders.forEach(function (header) {
    const columnIndex = headers.indexOf(header);
    if (columnIndex === -1) {
      throw new Error(sheetName + " 시트에 " + header + " 헤더가 필요합니다.");
    }
    columns[header] = columnIndex;
  });

  return {
    sheet: sheet,
    rows: rows,
    headers: headers,
    columns: columns,
  };
}

function getDatabaseSpreadsheet() {
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (activeSpreadsheet && activeSpreadsheet.getName() === DATABASE_SPREADSHEET_NAME) {
    return activeSpreadsheet;
  }

  const files = DriveApp.getFilesByName(DATABASE_SPREADSHEET_NAME);
  if (!files.hasNext()) {
    throw new Error(DATABASE_SPREADSHEET_NAME + " 구글 시트를 찾을 수 없습니다.");
  }

  return SpreadsheetApp.open(files.next());
}

function jsonResponse(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(
    ContentService.MimeType.JSON
  );
}

function safeSendAdminSms(text) {
  try {
    sendAdminSms(text);
  } catch (error) {
    console.error("SOLAPI 발송 실패: " + error.message);
  }
}

function sendAdminSms(text) {
  const properties = PropertiesService.getScriptProperties();
  const apiKey = properties.getProperty(SOLAPI_API_KEY_PROPERTY);
  const apiSecret = properties.getProperty(SOLAPI_API_SECRET_PROPERTY);
  const from = normalizePhoneNumber(properties.getProperty(SOLAPI_FROM_PROPERTY));
  const to = normalizePhoneNumber(properties.getProperty(ADMIN_PHONE_PROPERTY));

  if (!apiKey || !apiSecret || !from || !to) {
    throw new Error("SOLAPI Script Properties가 설정되지 않았습니다.");
  }

  const response = UrlFetchApp.fetch("https://api.solapi.com/messages/v4/send-many/detail", {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: createSolapiAuthHeader(apiKey, apiSecret),
    },
    payload: JSON.stringify({
      messages: [
        {
          to: to,
          from: from,
          text: text,
        },
      ],
    }),
    muteHttpExceptions: true,
  });

  const statusCode = response.getResponseCode();
  if (statusCode < 200 || statusCode >= 300) {
    throw new Error("SOLAPI " + statusCode + ": " + response.getContentText());
  }
}

function createSolapiAuthHeader(apiKey, apiSecret) {
  const date = new Date().toISOString();
  const salt = Utilities.getUuid().replace(/-/g, "");
  const signature = createHmacSha256Hex(date + salt, apiSecret);

  return (
    "HMAC-SHA256 apiKey=" +
    apiKey +
    ", date=" +
    date +
    ", salt=" +
    salt +
    ", signature=" +
    signature
  );
}

function createHmacSha256Hex(value, secret) {
  const signature = Utilities.computeHmacSha256Signature(value, secret);

  return signature
    .map(function (byte) {
      const value = byte < 0 ? byte + 256 : byte;
      return ("0" + value.toString(16)).slice(-2);
    })
    .join("");
}

function normalizePhoneNumber(phoneNumber) {
  return String(phoneNumber || "").replace(/\D/g, "");
}

function hashPassword(password) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    password,
    Utilities.Charset.UTF_8
  );
  const hex = bytes
    .map(function (byte) {
      const value = byte < 0 ? byte + 256 : byte;
      return ("0" + value.toString(16)).slice(-2);
    })
    .join("");

  return "sha256$" + hex;
}

function passwordMatches(storedPassword, password) {
  const stored = String(storedPassword || "");

  return stored === hashPassword(password) || stored === password;
}

function normalizeUserId(userId) {
  return String(userId || "")
    .trim()
    .toLowerCase();
}
