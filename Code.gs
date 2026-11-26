/**
 * 経費・領収書管理シート 自動集計スクリプト
 * 最小権限（対象スプレッドシート＋トリガーのみ）で動作します。
 */

const INPUT_SHEET_NAME = '入力';
const INPUT_HEADERS = ['日付', '支払先', '科目', '金額', 'メモ'];
const CATEGORIES = ['旅費交通費', '消耗品費', '会議費', '通信費', '交際費', '雑費'];
const ERROR_BG_COLOR = '#FCE8E6';
const TRIGGER_HANDLER_NAME = 'runMonthlyTrigger';

/**
 * スプレッドシートを開いたときにカスタムメニューを追加する
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('経費ツール')
    .addItem('入力シートを準備', 'setupInputSheet')
    .addItem('入力チェックを実行', 'runValidationFromMenu')
    .addItem('月次集計を作る', 'promptMonthlySummary')
    .addItem('CSVを取り込む（貼り付け）', 'promptCsvImport')
    .addSeparator()
    .addItem('毎月1日の自動集計を設定', 'setupMonthlyTrigger')
    .addItem('自動集計を解除', 'removeMonthlyTrigger')
    .addToUi();
}

/**
 * 「入力」シートを初期化する（既存データは消さない）
 */
function setupInputSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(INPUT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(INPUT_SHEET_NAME);
  }

  // ヘッダーが空の場合のみ書き込む
  const firstRow = sheet.getRange(1, 1, 1, INPUT_HEADERS.length).getValues()[0];
  if (firstRow.every((cell) => cell === '')) {
    sheet.getRange(1, 1, 1, INPUT_HEADERS.length).setValues([INPUT_HEADERS]);
    sheet.getRange(1, 1, 1, INPUT_HEADERS.length).setFontWeight('bold').setBackground('#F1F3F4');
    sheet.setFrozenRows(1);
  }

  // 科目列（C列）にプルダウンを設定する
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CATEGORIES, true)
    .setAllowInvalid(true) // ミス検知のため入力自体は許可し、チェックで色付けする
    .build();
  const maxRows = Math.max(sheet.getMaxRows() - 1, 100);
  sheet.getRange(2, 3, maxRows, 1).setDataValidation(rule);

  // 日付列・金額列の表示形式を整える
  sheet.getRange(2, 1, maxRows, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(2, 4, maxRows, 1).setNumberFormat('#,##0');

  showAlert_('「入力」シートを準備しました。');
}

/**
 * メニューから入力チェックを実行する
 */
function runValidationFromMenu() {
  const sheet = getInputSheetOrAlert_();
  if (!sheet) return;

  const result = validateInputSheet_(sheet);
  if (result.totalRows === 0) {
    showAlert_('チェック対象のデータがありません。');
    return;
  }

  if (result.errorCount === 0) {
    showAlert_(`入力チェック完了：エラーはありませんでした（${result.totalRows}件）。`);
  } else {
    showAlert_(`入力ミスが ${result.errorCount} 行見つかりました。赤い背景の行を確認してください（自動修正はしていません）。`);
  }
}

/**
 * メニューから対象月を指定して月次集計を作成する
 */
function promptMonthlySummary() {
  const ui = SpreadsheetApp.getUi();
  const defaultYm = formatYearMonth_(new Date());
  const response = ui.prompt(
    '月次集計の作成',
    `集計する年月を YYYY-MM 形式で入力してください（例: ${defaultYm}）`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const targetYm = response.getResponseText().trim();
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(targetYm)) {
    showAlert_('年月の形式が正しくありません。「YYYY-MM」（例: 2026-11）で入力してください。');
    return;
  }

  buildMonthlySummary_(targetYm, false);
}

/**
 * 毎月1日に自動実行される前月分の集計処理
 */
function runMonthlyTrigger() {
  const now = new Date();
  const prevMonthDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const targetYm = formatYearMonth_(prevMonthDate);
  buildMonthlySummary_(targetYm, true);
}

/**
 * 月次集計シートを作成する本体処理
 * @param {string} targetYm - 'YYYY-MM'
 * @param {boolean} isSilent - トリガー実行時はtrue（UIを出さない）
 */
function buildMonthlySummary_(targetYm, isSilent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName(INPUT_SHEET_NAME);
  if (!inputSheet) {
    if (!isSilent) showAlert_(`「${INPUT_SHEET_NAME}」シートが見つかりません。先に「入力シートを準備」を実行してください。`);
    return;
  }

  const validation = validateInputSheet_(inputSheet);
  if (validation.errorCount > 0) {
    const msg = `入力ミスが ${validation.errorCount} 行あるため集計を中止しました。赤い行を修正してください。`;
    if (isSilent) {
      console.warn(msg);
    } else {
      showAlert_(msg);
    }
    return;
  }

  // 科目ごとの集計器を初期化
  const summaryMap = {};
  CATEGORIES.forEach((cat) => {
    summaryMap[cat] = { count: 0, amount: 0 };
  });

  let totalCount = 0;
  let totalAmount = 0;

  if (validation.totalRows > 0) {
    const values = inputSheet.getRange(2, 1, validation.totalRows, 5).getValues();
    values.forEach((row) => {
      const dateObj = parseDateValue_(row[0]);
      if (!dateObj) return;
      if (formatYearMonth_(dateObj) !== targetYm) return;

      const category = String(row[2]).trim();
      const amount = Number(row[3]);
      if (!summaryMap[category]) {
        summaryMap[category] = { count: 0, amount: 0 };
      }
      summaryMap[category].count += 1;
      summaryMap[category].amount += amount;
      totalCount += 1;
      totalAmount += amount;
    });
  }

  const summarySheetName = `集計_${targetYm}`;
  let summarySheet = ss.getSheetByName(summarySheetName);
  if (!summarySheet) {
    summarySheet = ss.insertSheet(summarySheetName);
  } else {
    // 既存の月次集計シートのみ安全にクリアして再出力する
    summarySheet.clear();
  }

  const outputRows = [
    [`月次経費集計（${targetYm}）`, '', ''],
    ['科目', '件数', '金額合計']
  ];

  CATEGORIES.forEach((cat) => {
    outputRows.push([cat, summaryMap[cat].count, summaryMap[cat].amount]);
  });
  outputRows.push(['合計', totalCount, totalAmount]);

  summarySheet.getRange(1, 1, outputRows.length, 3).setValues(outputRows);
  summarySheet.getRange(1, 1).setFontWeight('bold').setFontSize(12);
  summarySheet.getRange(2, 1, 1, 3).setFontWeight('bold').setBackground('#F1F3F4');
  summarySheet.getRange(outputRows.length, 1, 1, 3).setFontWeight('bold').setBackground('#E8F0FE');
  summarySheet.getRange(3, 2, CATEGORIES.length + 1, 1).setNumberFormat('#,##0');
  summarySheet.getRange(3, 3, CATEGORIES.length + 1, 1).setNumberFormat('#,##0');
  summarySheet.autoResizeColumns(1, 3);

  const doneMsg = totalCount === 0
    ? `「${summarySheetName}」を作成しました（対象データ 0件）。`
    : `「${summarySheetName}」を作成しました（${totalCount}件 / 合計 ${totalAmount.toLocaleString()}円）。`;

  if (isSilent) {
    console.log(doneMsg);
  } else {
    showAlert_(doneMsg);
  }
}

/**
 * 「入力」シートの入力ミスをチェックし、エラー行に色を付ける（自動修正はしない）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @return {{totalRows: number, errorCount: number}}
 */
function validateInputSheet_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return { totalRows: 0, errorCount: 0 };
  }

  const numRows = lastRow - 1;
  const range = sheet.getRange(2, 1, numRows, 5);
  const values = range.getValues();
  const backgrounds = [];
  let errorCount = 0;

  values.forEach((row) => {
    const isDateValid = parseDateValue_(row[0]) !== null;
    const isCategoryValid = CATEGORIES.includes(String(row[2]).trim());
    const isAmountValid = isValidHalfWidthNumber_(row[3]);

    const hasError = !isDateValid || !isCategoryValid || !isAmountValid;
    if (hasError) {
      errorCount += 1;
      backgrounds.push([ERROR_BG_COLOR, ERROR_BG_COLOR, ERROR_BG_COLOR, ERROR_BG_COLOR, ERROR_BG_COLOR]);
    } else {
      backgrounds.push([null, null, null, null, null]);
    }
  });

  range.setBackgrounds(backgrounds);
  return { totalRows: numRows, errorCount: errorCount };
}

/**
 * メニューからCSVテキストを貼り付けて取り込む
 */
function promptCsvImport() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'CSV取り込み',
    '「日付,支払先,金額」形式のCSVテキストを貼り付けてください（1行目ヘッダー可）。',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const csvText = response.getResponseText();
  const result = importCsvText(csvText);
  if (result.error) {
    showAlert_(result.error);
  } else {
    showAlert_(`CSV取り込み完了：追加 ${result.added}件 / 重複スキップ ${result.skipped}件`);
  }
}

/**
 * CSV文字列（日付,支払先,金額）を「入力」シートに重複なく追記する
 * @param {string} csvText
 * @return {{added: number, skipped: number, error: string|null}}
 */
function importCsvText(csvText) {
  if (!csvText || !csvText.trim()) {
    return { added: 0, skipped: 0, error: 'CSVテキストが空です。データを消さずに処理を中止しました。' };
  }

  const sheet = getInputSheetOrAlert_();
  if (!sheet) {
    return { added: 0, skipped: 0, error: `「${INPUT_SHEET_NAME}」シートが見つかりません。` };
  }

  let rows;
  try {
    rows = Utilities.parseCsv(csvText.trim());
  } catch (e) {
    return { added: 0, skipped: 0, error: 'CSVの解析に失敗しました。形式を確認してください。' };
  }

  // 既存行の重複判定用キー（日付|支払先|金額）を作成
  const existingKeys = new Set();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const existingValues = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    existingValues.forEach((r) => {
      const key = buildDuplicateKey_(r[0], r[1], r[3]);
      if (key) existingKeys.add(key);
    });
  }

  const rowsToAppend = [];
  let skipped = 0;

  rows.forEach((cols, index) => {
    // 1行目がヘッダー行ならスキップ
    if (index === 0 && String(cols[0]).trim() === '日付') return;
    if (cols.length < 3) return;

    const rawDate = String(cols[0]).trim();
    const payee = String(cols[1]).trim();
    const rawAmount = String(cols[2]).trim();
    if (!rawDate && !payee && !rawAmount) return;

    const key = buildDuplicateKey_(rawDate, payee, rawAmount);
    if (key && existingKeys.has(key)) {
      skipped += 1;
      return;
    }

    if (key) existingKeys.add(key);
    const amountVal = isValidHalfWidthNumber_(rawAmount)
      ? Number(rawAmount)
      : sanitizeFormulaText_(rawAmount);
    // 科目は空欄で追記し、入力チェックで選択漏れに気付けるようにする
    rowsToAppend.push([
      sanitizeFormulaText_(rawDate),
      sanitizeFormulaText_(payee),
      '',
      amountVal,
      'CSV取込'
    ]);
  });

  if (rowsToAppend.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToAppend.length, 5).setValues(rowsToAppend);
  }

  return { added: rowsToAppend.length, skipped: skipped, error: null };
}

/**
 * 毎月1日 9時台に前月分の集計を作るトリガーを設定する
 */
function setupMonthlyTrigger() {
  removeTriggersByHandler_(TRIGGER_HANDLER_NAME);
  ScriptApp.newTrigger(TRIGGER_HANDLER_NAME)
    .timeBased()
    .onMonthDay(1)
    .atHour(9)
    .create();

  showAlert_('毎月1日（9時頃）に前月分の集計を自動作成するトリガーを設定しました。');
}

/**
 * 自動集計トリガーを解除する
 */
function removeMonthlyTrigger() {
  const removedCount = removeTriggersByHandler_(TRIGGER_HANDLER_NAME);
  if (removedCount > 0) {
    showAlert_(`自動集計トリガーを解除しました（${removedCount}件）。`);
  } else {
    showAlert_('解除対象のトリガーはありませんでした。');
  }
}

/* ---------- ヘルパー関数 ---------- */

function getInputSheetOrAlert_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INPUT_SHEET_NAME);
  if (!sheet) {
    showAlert_(`「${INPUT_SHEET_NAME}」シートが見つかりません。先に「入力シートを準備」を実行してください。`);
    return null;
  }
  return sheet;
}

function showAlert_(message) {
  SpreadsheetApp.getUi().alert(message);
}

function removeTriggersByHandler_(handlerName) {
  const triggers = ScriptApp.getProjectTriggers();
  let count = 0;
  triggers.forEach((trigger) => {
    if (trigger.getHandlerFunction() === handlerName) {
      ScriptApp.deleteTrigger(trigger);
      count += 1;
    }
  });
  return count;
}

function parseDateValue_(val) {
  if (val instanceof Date && !isNaN(val.getTime())) {
    return val;
  }
  if (typeof val === 'string') {
    const match = val.trim().match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
    if (!match) return null;
    const year = Number(match[1]);
    const month = Number(match[2]);
    const day = Number(match[3]);
    const d = new Date(year, month - 1, day);
    if (d.getFullYear() === year && d.getMonth() === month - 1 && d.getDate() === day) {
      return d;
    }
  }
  return null;
}

function isValidHalfWidthNumber_(val) {
  if (typeof val === 'number') {
    return !isNaN(val) && isFinite(val);
  }
  if (typeof val === 'string') {
    const trimmed = val.trim();
    if (trimmed === '') return false;
    // 半角数字のみ許可（全角数字などはエラーとして検知する）
    return /^-?\d+(\.\d+)?$/.test(trimmed);
  }
  return false;
}

function sanitizeFormulaText_(text) {
  return /^[=+\-@]/.test(text) ? `'${text}` : text;
}

function formatYearMonth_(dateObj) {
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  return Utilities.formatDate(dateObj, tz, 'yyyy-MM');
}

function formatDateYmd_(dateObj) {
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  return Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd');
}

function buildDuplicateKey_(dateVal, payeeVal, amountVal) {
  const d = parseDateValue_(dateVal);
  const datePart = d ? formatDateYmd_(d) : String(dateVal).trim();
  const payeePart = String(payeeVal).trim();
  const amountPart = isValidHalfWidthNumber_(amountVal) ? String(Number(amountVal)) : String(amountVal).trim();
  if (!datePart && !payeePart && !amountPart) return '';
  return `${datePart}|${payeePart}|${amountPart}`;
}
