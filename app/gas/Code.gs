/**
 * KITOKITO 在庫管理システム - Google Apps Script Web API
 *
 * スプレッドシートを簡易データベースとして使用し、
 * PWAアプリからのデータ読み書きをWeb API経由で行う。
 *
 * シート構成: "records"
 *   A: timestamp  - サーバー側タイムスタンプ（行追加時）
 *   B: date       - アプリ側の記録日
 *   C: type       - 記録種別（棚卸し, パーツ棚卸し, 持ち出し, 納品, 検品完了, 入庫, プリント, 裁断, 材料発注, 材料入荷）
 *   D: product    - 商品コードまたは材料名
 *   E: color      - カラーコードまたは詳細
 *   F: quantity   - 数量
 *   G: json       - 完全なレコードJSON（データ忠実性のため）
 */

// ===========================================
// シート初期化
// ===========================================

/**
 * "records" シートを初期化する。存在しなければヘッダー付きで作成。
 * @return {GoogleAppsScript.Spreadsheet.Sheet} recordsシート
 */
function initSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('records');

  if (!sheet) {
    sheet = ss.insertSheet('records');
    // ヘッダー行を設定
    sheet.getRange(1, 1, 1, 7).setValues([[
      'timestamp', 'date', 'type', 'product', 'color', 'quantity', 'json'
    ]]);
    // ヘッダー行を太字に
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
    // 列幅を調整
    sheet.setColumnWidth(1, 160); // timestamp
    sheet.setColumnWidth(2, 100); // date
    sheet.setColumnWidth(3, 100); // type
    sheet.setColumnWidth(4, 120); // product
    sheet.setColumnWidth(5, 100); // color
    sheet.setColumnWidth(6, 80);  // quantity
    sheet.setColumnWidth(7, 300); // json
  }

  return sheet;
}

// ===========================================
// データ変換ヘルパー
// ===========================================

/**
 * レコードオブジェクトをシートの行配列に変換する
 * @param {Object} record - レコードオブジェクト
 * @return {Array} [timestamp, date, type, product, color, quantity, json]
 */
function recordToRow(record) {
  var now = new Date();
  var timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  // 異なるレコードタイプに対応するため、複数フィールドをフォールバック
  var product = record.product || record.material || record.partCode || '';
  var color = record.color || record.category || '';
  var quantity = record.quantity || record.quantityOk || 0;

  return [
    timestamp,
    record.date || '',
    record.type || '',
    product,
    color,
    quantity,
    JSON.stringify(record)
  ];
}

/**
 * シートの行をレコードオブジェクトに変換する
 * G列のJSONをパースし、_rowNumを付与する
 * @param {Array} row - シートの行データ
 * @param {number} rowNum - スプレッドシート上の行番号（1始まり）
 * @return {Object} レコードオブジェクト（_rowNum付き）
 */
function rowToRecord(row, rowNum) {
  var record;

  try {
    // G列（index 6）のJSONをパース
    record = JSON.parse(row[6]);
  } catch (e) {
    // JSONパースに失敗した場合、列から復元
    record = {
      date: row[1],
      type: row[2],
      product: row[3],
      color: row[4],
      quantity: row[5]
    };
  }

  // 行番号を付与（後から参照できるように）
  record._rowNum = rowNum;
  return record;
}

// ===========================================
// JSON レスポンス生成
// ===========================================

/**
 * JSON形式のレスポンスを生成する
 * @param {Object} data - レスポンスデータ
 * @return {GoogleAppsScript.Content.TextOutput}
 */
function createJsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===========================================
// GET: 全レコード取得
// ===========================================

/**
 * GETリクエスト: "records"シートの全レコードをJSON配列で返す
 * 各レコードにはG列のJSONをパースした内容と_rowNumが含まれる
 * @param {Object} e - リクエストイベント
 * @return {GoogleAppsScript.Content.TextOutput}
 */
function doGet(e) {
  try {
    var sheet = initSheet();
    var lastRow = sheet.getLastRow();

    // データ行がない場合は空配列を返す
    if (lastRow <= 1) {
      return createJsonResponse({ success: true, records: [] });
    }

    // 2行目〜最終行のデータを取得（ヘッダー行を除く）
    var dataRange = sheet.getRange(2, 1, lastRow - 1, 7);
    var rows = dataRange.getValues();

    // 各行をレコードオブジェクトに変換
    var records = [];
    for (var i = 0; i < rows.length; i++) {
      var rowNum = i + 2; // スプレッドシートの行番号（1始まり、ヘッダー分+1）
      records.push(rowToRecord(rows[i], rowNum));
    }

    return createJsonResponse({ success: true, records: records });

  } catch (error) {
    return createJsonResponse({ success: false, error: error.message });
  }
}

// ===========================================
// POST: レコード追加・クリア
// ===========================================

/**
 * POSTリクエスト: action に応じて処理を分岐
 *   - "add": 1件追加
 *   - "bulkAdd": 複数件一括追加
 *   - "clear": 全データ行削除（ヘッダーは残す）
 * @param {Object} e - リクエストイベント
 * @return {GoogleAppsScript.Content.TextOutput}
 */
function doPost(e) {
  try {
    var sheet = initSheet();

    // アクションの取得: URLパラメータまたはリクエストボディから
    var action = '';
    var body = {};

    if (e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }

    action = (e.parameter && e.parameter.action) || body.action || '';

    // --- action: add（1件追加）---
    if (action === 'add') {
      var record = body.record || body;
      var row = recordToRow(record);
      sheet.appendRow(row);

      return createJsonResponse({ success: true, count: 1 });
    }

    // --- action: bulkAdd（一括追加）---
    if (action === 'bulkAdd') {
      var records = body.records || body;

      if (!Array.isArray(records)) {
        return createJsonResponse({ success: false, error: 'records は配列で指定してください' });
      }

      // 全レコードを行配列に変換
      var rows = [];
      for (var i = 0; i < records.length; i++) {
        rows.push(recordToRow(records[i]));
      }

      // まとめてシートに書き込み（パフォーマンス向上）
      if (rows.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
      }

      return createJsonResponse({ success: true, count: rows.length });
    }

    // --- action: clear（全データ削除、ヘッダーは保持）---
    if (action === 'clear') {
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.deleteRows(2, lastRow - 1);
      }

      return createJsonResponse({ success: true, count: 0 });
    }

    // --- action: delete（1件削除）---
    if (action === 'delete') {
      var delRecord = body.record || body;
      var lastRow = sheet.getLastRow();

      if (lastRow <= 1) {
        return createJsonResponse({ success: true, deleted: 0 });
      }

      var dataRange = sheet.getRange(2, 1, lastRow - 1, 7);
      var rows = dataRange.getValues();

      // 一致する行を後ろから検索（最新のものを優先削除）
      for (var i = rows.length - 1; i >= 0; i--) {
        try {
          var existing = JSON.parse(rows[i][6]);
          if (existing.date === delRecord.date &&
              existing.product === delRecord.product &&
              existing.color === delRecord.color &&
              existing.type === delRecord.type) {
            sheet.deleteRow(i + 2); // ヘッダー行分+1
            return createJsonResponse({ success: true, deleted: 1 });
          }
        } catch (e) {
          // JSONパース失敗はスキップ
        }
      }

      return createJsonResponse({ success: true, deleted: 0 });
    }

    // 不明なアクション
    return createJsonResponse({
      success: false,
      error: '不明なアクションです: ' + action + '（add, bulkAdd, delete, clear のいずれかを指定してください）'
    });

  } catch (error) {
    return createJsonResponse({ success: false, error: error.message });
  }
}
