/**
 * 家族収支管理アプリ - API Backend (v6.6.0)
 * LINE IDベースのフィルタリング / クレカフラグ / クレカ確認済(L列) / 明細取込
 */

const SPREADSHEET_ID = "1GLYgC7LVp5VUlH_4QpuUboEnBNtynYojywTF_-ulw8k";
const FOLDER_ID = "1LQXWk1qWOTGhDu47hHYPS5lv9wnJiGC-";

function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const action = params.action;
    let result;

    if (action === "register") {
      result = checkAndRegisterUser(params.data);
    } else if (action === "upload") {
      result = uploadData(params.data);
    } else if (action === "getList") {
      result = getListData();
    } else if (action === "getUsers") {
      result = getUserList();
    } else if (action === "markCreditConfirmed") {
      result = markCreditConfirmed(params.data);
    } else if (action === "importCreditTransaction") {
      result = importCreditTransaction(params.data);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: "success", result: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function uploadData(payload) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const dataSheet = ss.getSheetByName('data');
  
  // 複数画像の保存処理
  let urlList = [];
  if (payload.imageBlobs && payload.imageBlobs.length > 0) {
    payload.imageBlobs.forEach((base64Data, index) => {
      if (base64Data.includes("base64")) {
        const contentType = base64Data.match(/data:(.*?);/)[1];
        const data = base64Data.split(',')[1];
        const blob = Utilities.newBlob(Utilities.base64Decode(data), contentType, `receipt_${Date.now()}_${index}.jpg`);
        const file = DriveApp.getFolderById(FOLDER_ID).createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        urlList.push(file.getUrl());
      }
    });
  }
  const fileUrls = urlList.length > 0 ? urlList.join(",") : "画像なし";

  // クレジットカード・PayPayフラグ (出金時のみ、チェック時は"○"、未チェックは空)
  const creditCard = (payload.type === "出金" && payload.creditCard === true) ? "○" : "";
  const paypay = (payload.type === "出金" && payload.paypay === true) ? "○" : "";

  // スプレッドシートへの書き込み (I:収支 / J:クレカ / K:PayPay / L:クレカ確認済)
  dataSheet.appendRow([
    new Date(),
    payload.date,
    payload.userName,
    payload.amount,
    payload.category,
    payload.shop,
    fileUrls,
    payload.lineId,
    payload.type,
    creditCard,
    paypay,
    ""
  ]);
  return "Success";
}

function getListData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('data');
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const values = sheet.getDataRange().getValues();
  const out = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const sheetRow = r + 1;
    out.push({
      sheetRow: sheetRow,
      date: row[1] instanceof Date ? Utilities.formatDate(row[1], "JST", "yyyy-MM-dd") : String(row[1] || ""),
      userName: row[2] || "",
      amount: Number(row[3]) || 0,
      category: row[4] || "",
      shop: row[5] || "",
      imageUrls: row[6] ? String(row[6]).split(",") : [],
      lineId: row[7] || "",
      type: row[8] || "出金",
      creditCard: row[9] === "○",
      paypay: row[10] === "○",
      creditCardConfirmed: row[11] === "○"
    });
  }
  return out.reverse();
}

/** L列にクレカ明細「確認済」(○)を記入（sheetRow はシートの行番号・1始まり） */
function markCreditConfirmed(data) {
  const rowIndex = Number(data.rowIndex);
  if (!rowIndex || rowIndex < 2) throw new Error("invalid rowIndex");
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("data");
  sheet.getRange(rowIndex, 12).setValue("○");
  return "OK";
}

/** CSV明細から出金＋クレカとして data に1行追加 */
function importCreditTransaction(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("data");
  sheet.appendRow([
    new Date(),
    data.date,
    data.userName,
    Number(data.amount),
    data.category || "他",
    data.shop || "クレカ明細取込",
    "画像なし",
    data.lineId,
    "出金",
    "○",
    "",
    ""
  ]);
  return "OK";
}

function checkAndRegisterUser(lineData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const userSheet = ss.getSheetByName('users');
  const data = userSheet.getDataRange().getValues();
  
  if (!data.some(row => row[0] === lineData.userId)) {
    userSheet.appendRow([lineData.userId, lineData.displayName, "", new Date()]);
  }
  return "OK";
}

function getUserList() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('users');
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  
  return values.slice(1).map(row => ({
    lineId: row[0],
    name: row[1]
  }));
}

function doGet() {
  return ContentService.createTextOutput("GAS API v6.6.0 active.");
}
