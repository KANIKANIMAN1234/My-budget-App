/**
 * 家族収支管理アプリ - API Backend (v6.2.0)
 * LINE IDベースのフィルタリング対応版 / クレジットカードフラグ対応
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

  // スプレッドシートへの書き込み (I列:収支タイプ / J列:クレカ / K列:PayPay)
  dataSheet.appendRow([
    new Date(),
    payload.date,
    payload.userName,
    payload.amount,
    payload.category,
    payload.shop,
    fileUrls,
    payload.lineId,
    payload.type,     // I列: 出金 or 入金
    creditCard,       // J列: クレジットカード支払いフラグ (○ or 空)
    paypay            // K列: PayPay支払いフラグ (○ or 空)
  ]);
  return "Success";
}

function getListData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('data');
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  
  const values = sheet.getDataRange().getValues();
  return values.slice(1).reverse().map(row => ({
    date: row[1] instanceof Date ? Utilities.formatDate(row[1], "JST", "yyyy-MM-dd") : row[1],
    userName: row[2] || "",
    amount: Number(row[3]) || 0,
    category: row[4] || "",
    shop: row[5] || "",
    imageUrls: row[6] ? row[6].split(",") : [],
    lineId: row[7] || "",
    type: row[8] || "出金",
    creditCard: row[9] === "○",  // J列: クレジットカードフラグ
    paypay: row[10] === "○"      // K列: PayPayフラグ
  }));
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
  return ContentService.createTextOutput("GAS API v6.2.0 active."); 
}
