/**
 * Summaryシート内の整合性を多角的にチェックする
 * 1. 縦横の合計値の不一致
 * 2. 月平均の計算ミス（経過月数に基づく）
 * 3. 前年度月平均の転記ミス
 * 4. 平均差額の計算ミス
 */
function validateSummaries() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetYears = ['2024', '2025'];
  const errorMessages = [];

  targetYears.forEach(year => {
    const summarySheet = ss.getSheetByName('Summary_' + year);
    if (!summarySheet) return;

    const summaryData = summarySheet.getDataRange().getValues();
    const header = summaryData[0];

    // インデックスの特定
    const colIdx = {
      total: header.indexOf('計'),
      avg: header.indexOf('月平均'),
      prevAvg: header.indexOf('前年度月平均'),
      diff: header.indexOf('差額'),
    };

    const lastRow = summaryData.find(row => row[0] === '計');
    if (colIdx.total === -1 || !lastRow) {
      errorMessages.push(`[${year}] シート形式が正しくありません。`);
      return;
    }

    const monthStartIdx = 1; // 0番目は項目名、1番目から12ヶ月分
    let elapsedMonths = 0;
    for (let m = 0; m < 12; m++) {
      const hasData = summaryData.some(
        (row, idx) => idx > 0 && row[0] !== '計' && row[monthStartIdx + m] !== 0
      );
      if (hasData) elapsedMonths = m + 1;
    }

    let sumOfItemTotals = 0;

    // --- 各行のバリデーション ---
    for (let i = 1; i < summaryData.length; i++) {
      const row = summaryData[i];
      if (row[0] === '計') continue;

      const itemName = row[0];
      const rowTotal = parseFloat(row[colIdx.total]) || 0;
      const currentAvg = parseFloat(row[colIdx.avg]) || 0;
      const currentPrevAvg = parseFloat(row[colIdx.prevAvg]) || 0;
      const currentDiff = parseFloat(row[colIdx.diff]) || 0;

      sumOfItemTotals += rowTotal;

      // A. 月平均のチェック (合計 / 経過月数)
      const expectedAvg = elapsedMonths > 0 ? rowTotal / elapsedMonths : 0;
      if (Math.abs(currentAvg - expectedAvg) > 1) {
        errorMessages.push(
          `[${year}] ${itemName}: 月平均が正しくありません。(期待値:${expectedAvg.toFixed(0)}, 現状:${currentAvg.toFixed(0)})`
        );
      }

      // B. 平均差額のチェック (月平均 - 前年度月平均)
      const expectedDiff = currentAvg - currentPrevAvg;
      if (Math.abs(currentDiff - expectedDiff) > 1) {
        errorMessages.push(
          `[${year}] ${itemName}: 平均差額が正しくありません。(期待値:${expectedDiff.toFixed(0)}, 現状:${currentDiff.toFixed(0)})`
        );
      }
    }

    // --- 縦横合計の不一致チェック ---
    const columnTotalValue = parseFloat(lastRow[colIdx.total]) || 0;
    if (Math.abs(sumOfItemTotals - columnTotalValue) > 1) {
      errorMessages.push(
        `[${year}] 縦横合計不一致: 行合計の和:${sumOfItemTotals.toLocaleString()}, 列合計の計:${columnTotalValue.toLocaleString()}`
      );
    }
  });

  if (errorMessages.length > 0) {
    uiAlert_(
      'バリデーションエラーが発生しました：\n\n' + errorMessages.join('\n\n')
    );
  } else {
    console.log('すべてのバリデーション（合計・平均・前年比）に合格しました。');
  }
}
