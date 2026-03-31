/**
 * メイン実行関数。
 * 1. 初回実行時はフォルダ内全件、2回目以降は前後2ヶ月のCSVを取り込む。
 * 2. 取り込んだデータを「csv」シートに追記する。
 * 3. 最後に年度別のクロス集計表をすべて更新する。
 */
function importCsvWithLogic() {
  const props = PropertiesService.getScriptProperties();
  const folderId = props.getProperty('FOLDER_ID');
  const targetSpecified = props.getProperty('TARGET_FILE_NAME');
  const isInitialized = props.getProperty('INITIALIZED'); // 初回実行済みフラグ
  const sheetName = 'csv';

  if (!folderId) {
    uiAlert_('エラー: スクリプトプロパティ「FOLDER_ID」を設定してください。');
    return;
  }

  if (!folderId) {
    uiAlert_('エラー: スクリプトプロパティ「FOLDER_ID」を設定してください。');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 初回のみの特別処理 ---
  if (!isInitialized && !targetSpecified) {
    // 既存のcsvシートがあれば削除
    const oldSheet = ss.getSheetByName(sheetName);
    if (oldSheet) {
      ss.deleteSheet(oldSheet);
    }
    // 再作成は processCsvFiles_ 内で行われるためここでは削除のみ
  }

  /** @type {string[]} 処理対象のファイル名リスト */
  const targetFileNames = [];

  // 1. 取得対象の決定
  if (targetSpecified && targetSpecified.trim() !== '') {
    targetFileNames.push(targetSpecified.trim());
  } else if (!isInitialized) {
    // 【初回のみ】フォルダ内のすべてのCSVを検索してリスト化
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.searchFiles(
      'title contains ".csv" and trashed = false'
    );
    while (files.hasNext()) {
      targetFileNames.push(files.next().getName());
    }
  } else {
    // 2回目以降：実行日の前後2ヶ月分を自動生成
    const targetMonths = getTargetMonths_(2);
    targetFileNames.push(...targetMonths);
  }

  // 2. CSVインポート実行
  processCsvFiles_(folderId, sheetName, targetFileNames);

  // 3. インポート後に集計表を更新
  updateAllSummaries_();

  // 4. 初回実行が成功したらフラグを立てる
  if (!isInitialized && targetFileNames.length > 0) {
    props.setProperty('INITIALIZED', 'true');
  }
}

/**
 * 指定されたファイル名のCSVを読み込み、スプレッドシートに追記する。
 * @param {string} folderId - CSVが格納されているフォルダID
 * @param {string} sheetName - 書き込み先のシート名
 * @param {string[]} targetFileNames - 処理対象のファイル名リスト
 * @private
 */
function processCsvFiles_(folderId, sheetName, targetFileNames) {
  const folder = DriveApp.getFolderById(folderId);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);

  // シートがない場合は新規作成し、見出しを追加
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow([
      'ファイル名',
      '年月日',
      '項目名',
      'filler1',
      'filler2',
      'filler3',
      '金額',
      '備考',
    ]);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'ファイル名',
      '年月日',
      '項目名',
      'filler1',
      'filler2',
      'filler3',
      '金額',
      '備考',
    ]);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
  }

  const lastRow = sheet.getLastRow();
  let existingFileNames = [];
  if (lastRow > 0) {
    existingFileNames = sheet.getRange(1, 1, lastRow, 1).getValues().flat();
  }

  let importedCount = 0;
  const skippedFiles = [];

  targetFileNames.forEach(fileName => {
    // CSVファイルのみを対象（拡張子チェック）
    if (!fileName.toLowerCase().endsWith('.csv')) return;

    if (existingFileNames.includes(fileName)) {
      skippedFiles.push(fileName);
      return;
    }

    const files = folder.getFilesByName(fileName);
    if (!files.hasNext()) return;

    const file = files.next();
    try {
      const content = file.getBlob().getDataAsString('shift_jis');
      const lines = content
        .split(/\r\n|\n|\r/)
        .filter(line => line.trim() !== '');
      if (lines.length <= 1) return;

      lines.shift(); // 1行目（顧客名等）を削除

      const finalData = [];
      lines.forEach((line, index) => {
        const parts = line.split(',');
        if (parts.length >= 6) {
          const date = parts[0];
          let amount = parts[parts.length - 2];
          let extra = parts[parts.length - 1];
          // 金額列の妥当性チェック
          if (isNaN(amount.replace(/ /g, '').trim()) || amount === '') {
            amount = parts[parts.length - 1];
            extra = '';
          }
          const f3 = parts[parts.length - 3],
            f2 = parts[parts.length - 4],
            f1 = parts[parts.length - 5];
          let itemName = parts
            .slice(1, parts.length - 5)
            .join(',')
            .trim();

          // 最終行（合計行）の判定
          if (index === lines.length - 1)
            itemName = '【合計金額】' + (itemName || '');

          finalData.push([fileName, date, itemName, f1, f2, f3, amount, extra]);
        }
      });

      if (finalData.length > 0) {
        sheet
          .getRange(
            sheet.getLastRow() + 1,
            1,
            finalData.length,
            finalData[0].length
          )
          .setValues(finalData);
        importedCount++;
        existingFileNames.push(fileName);
      }
    } catch (e) {
      console.error(fileName + ' の処理中にエラー: ' + e.message);
    }
  });

  if (importedCount > 0) {
    uiAlert_(importedCount + ' 件のファイルを取り込みました。');
  } else if (skippedFiles.length > 0 && targetFileNames.length === 1) {
    uiAlert_('指定されたファイルは既に取り込み済みです。');
  }
}

/**
 * 原データシート（csv）から年度・項目ごとに集計し、Summaryシートを更新する。
 * @private
 */
function updateAllSummaries_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName('csv');
  if (!rawSheet || rawSheet.getLastRow() < 2) return;

  const data = rawSheet.getDataRange().getValues();
  const yearlyData = {};

  for (let i = 1; i < data.length; i++) {
    const fileName = String(data[i][0]); // YYYYMM.csv
    const itemNameRaw = data[i][2];
    const amount = parseFloat(data[i][6]) || 0;
    if (itemNameRaw.includes('【合計金額】')) continue;

    let year = parseInt(fileName.substring(0, 4));
    const month = fileName.substring(4, 6);

    // 1, 2, 3月は前年度の扱いにする（会計年度対応）
    const monthNum = parseInt(month);
    if (monthNum >= 1 && monthNum <= 3) {
      year = year - 1;
    }
    const fiscalYear = year.toString();

    const itemName = normalizeItemName_(itemNameRaw);

    if (!yearlyData[fiscalYear]) yearlyData[fiscalYear] = {};
    if (!yearlyData[fiscalYear][itemName])
      yearlyData[fiscalYear][itemName] = {};
    yearlyData[fiscalYear][itemName][month] =
      (yearlyData[fiscalYear][itemName][month] || 0) + amount;
  }

  Object.keys(yearlyData)
    .sort()
    .forEach(year => {
      generateYearlySummary_(year, yearlyData[year]);
    });
}

/**
 * 特定年度のクロス集計表シートを生成・更新する。
 * 末尾に「カテゴリ」「サービス内容」の列を追加し、サービス一覧シートからVLOOKUPで参照する。
 * @param {string} year - 年度（YYYY）
 * @param {Object} itemsObj - 項目と月ごとの金額を格納したオブジェクト
 * @private
 */
function generateYearlySummary_(year, itemsObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'Summary_' + year;
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  sheet.clear();

  const prevYear = (parseInt(year) - 1).toString();
  const prevYearAvgMap = getYearlyAverageMap_(prevYear);

  // 会計年度の並び順（4月始まり）
  const monthOrder = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3];

  // 分母（経過月数）の算定
  let maxMonthIdx = 0;
  monthOrder.forEach((m, idx) => {
    const monthKey = ('0' + m).slice(-2);
    const hasDataAnywhere = Object.values(itemsObj).some(
      itemMonths => itemMonths[monthKey] !== undefined
    );
    if (hasDataAnywhere) maxMonthIdx = idx;
  });
  const elapsedMonths = maxMonthIdx + 1;

  // --- ヘッダーの作成 ---
  const header = ['項目名'];
  monthOrder.forEach(m => {
    const displayYear = m <= 3 ? parseInt(year) + 1 : year;
    header.push(`${displayYear}年${('0' + m).slice(-2)}月`);
  });
  // 右側の列を追加
  header.push(
    '計',
    '月平均',
    '前年度月平均',
    '差額',
    'カテゴリ',
    'サービス内容'
  );

  const rows = [header];
  const sortedItemNames = Object.keys(itemsObj).sort();

  // --- データの作成 ---
  sortedItemNames.forEach((name, idx) => {
    const row = [name];
    let rowTotal = 0;

    monthOrder.forEach(m => {
      const monthKey = ('0' + m).slice(-2);
      const val = itemsObj[name][monthKey] || 0;
      row.push(val);
      rowTotal += val;
    });

    const avg = elapsedMonths > 0 ? rowTotal / elapsedMonths : 0;
    const prevAvg = prevYearAvgMap[name] || 0;
    const diff = avg - prevAvg;

    // データ行に数式をセットするための準備
    // 行番号は、ヘッダーが1行目なので idx + 2 となる
    const rowNum = idx + 2;
    const categoryFormula = `=iferror(vlookup(A${rowNum},'サービス一覧'!A:D,3,false))`;
    const detailFormula = `=iferror(vlookup(A${rowNum},'サービス一覧'!A:D,4,false))`;

    row.push(rowTotal, avg, prevAvg, diff, categoryFormula, detailFormula);
    rows.push(row);
  });

  // --- 列合計（最下行）の計算 ---
  const colTotals = ['計'];
  // 数値列（月別〜差額まで）のみ合計を出す
  const lastNumColIdx = header.indexOf('差額');
  for (let c = 1; c <= lastNumColIdx; c++) {
    let sum = 0;
    for (let r = 1; r < rows.length; r++) {
      const val = rows[r][c];
      // 数式（文字列）は除外して数値のみ加算
      if (typeof val === 'number') sum += val;
    }
    colTotals.push(sum);
  }
  // カテゴリとサービス内容の列合計は空欄
  colTotals.push('', '');
  rows.push(colTotals);

  // --- シートへの書き込み ---
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);

  // --- 書式設定と列幅の指定 ---
  const lastCol = header.length;

  sheet.setColumnWidths(1, lastCol, 85);
  sheet.setColumnWidth(1, 300);
  sheet.setColumnWidth(16, 95);
  sheet.setColumnWidth(18, 168);
  sheet.setColumnWidth(19, 419);

  const fullRange = sheet.getRange(1, 1, rows.length, rows[0].length);

  // 【追加】全体に格子状の罫線を引く (上, 左, 下, 右, 垂直, 水平)
  fullRange.setBorder(
    true,
    true,
    true,
    true,
    true,
    true,
    '#999999',
    SpreadsheetApp.BorderStyle.SOLID
  );

  if (rows.length > 1) {
    // 数値範囲（月別〜差額まで）にカンマ区切りを適用
    sheet
      .getRange(2, 2, rows.length - 1, lastNumColIdx)
      .setNumberFormat('#,##0');
  }

  // デザイン調整
  sheet
    .getRange(1, 1, 1, header.length)
    .setBackground('#f3f3f3')
    .setFontWeight('bold');

  // 最下行（列合計）の太字
  sheet.getRange(rows.length, 1, 1, header.length).setFontWeight('bold');

  // 最下行の上側だけ二重線にする（合計を強調する）
  sheet
    .getRange(rows.length, 1, 1, header.length)
    .setBorder(
      true,
      null,
      null,
      null,
      null,
      null,
      '#000000',
      SpreadsheetApp.BorderStyle.DOUBLE
    );

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
}

/**
 * 項目名の名寄せを行う。キーワードマッチングまたは不要な英数字の除去。
 * @param {string} name - 元の項目名
 * @returns {string} 名寄せ後の項目名
 * @private
 */
function normalizeItemName_(name) {
  // 1. 全角英数字を半角化し、全角空白を半角空白に置換、さらに大文字へ統一
  const n = name
    .replace(/[！-～]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xfee0)) // 英数字半角化
    .replace(/\u3000/g, ' ') // 全角空白を半角へ【追加】
    .toUpperCase();

  if (n.includes('1PASSWORD')) return '1Password';
  if (n.includes('DOCKER, INC.')) return 'Docker';
  if (n.includes('DRI*PVTLTRACKER')) return 'Pivotal Tracker';
  if (n.includes('MAILTRAP')) return 'Mailtrap';
  if (n.includes('AMAZON WEB SERVICES')) return 'Amazon Web Services';
  if (n.includes('PAPERTRAIL-SOLARWINDS') || n.includes('SOLARWINDS'))
    return 'Papertrail';
  if (n.includes('PULUMI CORPORATION')) return 'Pulumi';
  if (n.includes('ROLLBAR')) return 'Rollbar';
  if (n.includes('WWW.DEEPL.COM') || n.includes('DEEPL')) return 'DeepL';
  if (n.includes('ZOOM')) return 'Zoom';
  if (n.includes('DROPBOX')) return 'Dropbox';
  if (n.includes('GOOGLE*WORKSPACE') || n.includes('GOOGLE*GSUITE'))
    return 'Google Workspace';
  if (n.includes('AMAZON')) return 'Amazon';
  if (n.includes('GITHUB')) return 'GitHub';
  if (n.includes('さくらインターネット')) return 'さくらインターネット';
  if (n.includes('OPENAI')) return 'OpenAI';
  if (n.includes('HEROKU')) return 'Heroku';
  if (n.includes('CODE CLIMATE')) return 'Code Climate';
  if (
    n.includes('カブシキガイシャボックス') ||
    n.includes('カブシキガイシヤボツクス')
  )
    return 'Box';
  if (n.includes('SKYPE')) return 'Skype';
  if (n.includes('オナマエドツトコムドメイン')) return 'お名前.COMドメイン';
  if (n.includes('LINEAR.APP')) return 'LINEAR.APP';

  // ヒットしない場合は記号やスペースより前を抽出してID等を除去
  //return n.split(/[\s\*（(]/)[0];
  return n;
}

/**
 * 指定した年度のSummaryシートから項目ごとの月平均額を取得する。
 * @param {string} year - 取得対象の年度
 * @returns {Object} 項目名をキー、月平均額を値とする連想配列
 * @private
 */
function getYearlyAverageMap_(year) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    'Summary_' + year
  );
  const map = {};
  if (!sheet) return map;
  const data = sheet.getDataRange().getValues();

  // --- ヘッダーから「月平均」のインデックスを特定 ---
  const avgColIdx = data[0].indexOf('月平均');
  if (avgColIdx === -1) return map;

  for (let i = 1; i < data.length - 1; i++) {
    map[data[i][0]] = data[i][avgColIdx];
  }
  // -------------------------------------------------------------
  return map;
}

/**
 * 現在の日付から指定された月数範囲のファイル名リスト（YYYYMM.csv）を生成する。
 * @param {number} range - 前後何ヶ月分を取得するか
 * @returns {string[]} ファイル名の配列
 * @private
 */
function getTargetMonths_(range) {
  const names = [];
  const now = new Date();
  for (let i = -range; i <= range; i++) {
    const d = new Date(now.getFullYear(), now.getMonth() + i, 1);
    const yyyy = d.getFullYear();
    const mm = ('0' + (d.getMonth() + 1)).slice(-2);
    names.push(yyyy + mm + '.csv');
  }
  return names;
}

/**
 * スプレッドシート上にアラート（ポップアップ）を表示する。
 * @param {string} msg - 表示するメッセージ
 * @private
 */
function uiAlert_(msg) {
  SpreadsheetApp.getUi().alert(msg);
}
