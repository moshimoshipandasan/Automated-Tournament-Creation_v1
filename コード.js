/*
内容：シンプルトーナメント作成システム
作成日：2021/11/02
更新日：2025/02/26
作成・著作者：安藤昇＠青山学院中等部
改善：Cline AI
備考：プログラムの配布・改変などする場合にはメールにてご連絡ください
email:gigaschool2020@gmail.com（安藤宛）
*/

// グローバル変数の定義
const IRP = 3; // 初期行位置
const ICP = 1; // 初期列位置
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sht0 = ss.getSheetByName('大会データ');
const sht1 = ss.getSheetByName('トーナメント');
const sht2 = ss.getSheetByName('ブロック');
const sht5 = ss.getSheetByName('テーブル');

/**
 * スプレッドシートが開かれたときに実行される関数
 * メニューを作成する
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('トーナメント');
  menu.addItem('トーナメント作成', 'createTournamentFromInput');
  menu.addItem('シートクリア', 'clearSheets');
  menu.addToUi();
}

/**
 * 参加人数の入力を促し、トーナメントを作成する関数
 */
function createTournamentFromInput() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'トーナメント作成',
    '参加人数を入力してください：',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() == ui.Button.OK) {
    const entryTotalNumber = parseInt(response.getResponseText());
    
    if (isNaN(entryTotalNumber) || entryTotalNumber < 2) {
      ui.alert('エラー', '有効な参加人数を入力してください（2以上の整数）', ui.ButtonSet.OK);
      return;
    }
    
    // 参加人数を保存
    sht0.getRange(3, 2).setValue(entryTotalNumber);
    
    // シートの初期化
    clearSheet(sht1);
    sht2.clear();
    sht5.clear();
    
    // トーナメント作成
    createTournament();
  }
}

/**
 * シートを初期化する関数
 * @param {Sheet} sheet - 初期化するシート
 */
function clearSheet(sheet) {
  sheet
    .clear()
    .setHiddenGridlines(true)
    .setColumnWidths(1, 26, 25)
    .setRowHeights(1, 100, 10);
}

/**
 * すべてのシートをクリアする関数
 */
function clearSheets() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'シートクリア',
    'すべてのシートをクリアしますか？',
    ui.ButtonSet.YES_NO
  );

  if (response == ui.Button.YES) {
    clearSheet(sht1);
    sht2.clear();
    sht5.clear();
    ui.alert('シートをクリアしました');
  }
}

/**
 * 2次元配列を初期化する関数
 * @param {number} rows - 行数
 * @param {number} cols - 列数
 * @return {Array} 初期化された2次元配列
 */
function initializeArray(rows, cols) {
  const array = new Array(rows);
  for (let i = 0; i < rows; i++) {
    array[i] = new Array(cols).fill(0);
  }
  return array;
}

/**
 * トーナメント表を作成する関数
 */
function createTournament() {
  try {
    // 基本パラメータの設定
    const entryTotalNumber = sht0.getRange(3, 2).getValue();
    const entryGroup = Math.floor(Math.log2(entryTotalNumber - 1));
    const entryTableNumber = 2 ** (entryGroup + 1);
    
    // テーブルと区画の初期化
    const tableSet = initializeArray(entryGroup + 1, entryTableNumber);
    const blockSet = initializeArray(entryGroup + 1, entryTableNumber);
    
    // 初期値の設定
    tableSet[0][0] = 1;
    tableSet[0][1] = 4;
    tableSet[0][2] = 3;
    tableSet[0][3] = 2;
    
    // テーブルの設定
    setupTableSet(tableSet, entryGroup, entryTableNumber);
    
    // 最終グループの設定
    for (let i = 0; i < entryTableNumber; i++) {
      tableSet[entryGroup][i] = tableSet[entryGroup - 1][i];
      if (tableSet[entryGroup - 1][i] > entryTotalNumber) {
        tableSet[entryGroup][i] = 0;
      }
    }
    
    // トーナメント表の描画
    drawTournament(tableSet, entryGroup, entryTotalNumber, entryTableNumber);
    
    // 完了メッセージ
    Browser.msgBox("トーナメントの作成が完了しました。");
  } catch (error) {
    Browser.msgBox("エラーが発生しました: " + error.message);
    console.error(error);
  }
}

/**
 * テーブルセットを設定する関数
 * @param {Array} tableSet - テーブルセット配列
 * @param {number} entryGroup - エントリーグループ数
 * @param {number} entryTableNumber - テーブル数
 */
function setupTableSet(tableSet, entryGroup, entryTableNumber) {
  for (let h = 0; h < entryGroup - 1; h++) {
    for (let i = 0; i < entryTableNumber / 2; i++) {
      if (tableSet[h][i] !== 0) {
        tableSet[h + 1][i * 2] = tableSet[h][i];
        tableSet[h + 1][i * 2 + 1] = Math.abs(2 ** (h + 3) + 1 - tableSet[h][i]);
        
        // 特定の位置の値を入れ替え
        for (let j = 2; j <= entryTableNumber - 1; j = j + 4) {
          const tempBox = tableSet[h + 1][j];
          tableSet[h + 1][j] = tableSet[h + 1][j + 1];
          tableSet[h + 1][j + 1] = tempBox;
        }
      }
    }
  }
}

/**
 * トーナメント表を描画する関数
 * @param {Array} tableSet - テーブルセット配列
 * @param {number} entryGroup - エントリーグループ数
 * @param {number} entryTotalNumber - 参加者総数
 * @param {number} entryTableNumber - テーブル数
 */
function drawTournament(tableSet, entryGroup, entryTotalNumber, entryTableNumber) {
  // シートのアクティブ化と書式設定
  sht1.activate();
  sht1.setColumnWidth(ICP + 1, 150);
  sht1.getRange(IRP, ICP, entryTotalNumber * 4 * 1, 3)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  
  // トーナメント番号の設定（値の設定を一括で行う）
  const tournamentData = [];
  const tableData = [];
  
  let pn = 1;
  for (let i = 0; i < entryTableNumber; i++) {
    if (tableSet[entryGroup][i] !== 0) {
      // テーブルシートへの記録用データ準備
      tableData.push({
        row: tableSet[entryGroup][i],
        col: 1,
        value: tableSet[entryGroup][i]
      });
      
      tableData.push({
        row: tableSet[entryGroup][i],
        col: 2,
        value: pn
      });
      
      // トーナメントシートへの記録用データ準備
      tournamentData.push({
        row: i * 2 + IRP,
        col: ICP,
        value: pn
      });
      
      tournamentData.push({
        row: i * 2 + 1 + IRP,
        col: ICP,
        value: pn
      });
      
      tournamentData.push({
        row: i * 2 + IRP,
        col: ICP + 2,
        value: tableSet[entryGroup][i]
      });
      
      tournamentData.push({
        row: i * 2 + 1 + IRP,
        col: ICP + 2,
        value: tableSet[entryGroup][i]
      });
      
      pn++;
    }
  }
  
  // 値を一括で設定
  for (const data of tournamentData) {
    sht1.getRange(data.row, data.col).setValue(data.value);
  }
  
  for (const data of tableData) {
    sht5.getRange(data.row, data.col).setValue(data.value);
  }
  
  // 罫線の描画
  drawBorders(entryGroup, entryTableNumber);
  
  // ブロック情報の設定
  setupBlocks(entryGroup, entryTableNumber, tableSet);
  
  // 特定の罫線の調整
  adjustSpecificBorders(entryTableNumber, entryGroup);
  
  // セルのマージと調整
  mergeCells(entryTotalNumber);
  
  // 特定の位置の処理
  let tp = 0;
  for (let i = 1; i < entryTotalNumber; i++) {
    if (sht1.getRange(2 * i + IRP - 2, ICP + 2).getValue() === 3) {
      tp = i - 1;
    }
  }
  
  // 特定の罫線の設定
  sht1.getRange(2 * tp + IRP - 1, ICP + entryGroup + 4, 2, 1)
    .setBorder(null, null, null, null, null, true, "black", SpreadsheetApp.BorderStyle.SOLID);
  
  // 右側のトーナメント表の描画
  drawRightTournament(entryGroup, entryTableNumber, entryTotalNumber, tableSet, tp);
  
  // シートの先頭にフォーカスを移動
  sht1.getRange(1, 1).activate();
}

/**
 * 罫線を描画する関数
 * @param {number} entryGroup - エントリーグループ数
 * @param {number} entryTableNumber - テーブル数
 */
function drawBorders(entryGroup, entryTableNumber) {
  let k = entryTableNumber / 4;
  let l = entryTableNumber / k;
  
  for (let h = 0; h < entryGroup + 1; h++) {
    for (let i = 0; i < k * 2; i++) {
      sht1.getRange(2 ** h + l * i + IRP, ICP + 3 + h, l / 2, 1)
        .setBorder(true, null, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
    }
    l = l * 2;
    k = k / 2;
  }
}

/**
 * ブロック情報を設定する関数
 * @param {number} entryGroup - エントリーグループ数
 * @param {number} entryTableNumber - テーブル数
 * @param {Array} tableSet - テーブルセット配列
 */
function setupBlocks(entryGroup, entryTableNumber, tableSet) {
  let n = 0;
  const j = 1;
  const j2 = 2 ** j;
  
  // ブロック情報を一括で設定するためのデータを準備
  const blockData = [];
  
  for (let i = 0; i < 2 ** (entryGroup + 1); i++) {
    if (tableSet[entryGroup][i] !== 0) {
      n++;
    }
    if (i % j2 === 1) {
      blockData.push({
        row: (i + 1) / j2,
        col: 1,
        value: n
      });
    }
  }
  
  // 値を一括で設定
  for (const data of blockData) {
    sht2.getRange(data.row, data.col).setValue(data.value);
  }
}

/**
 * 特定の罫線を調整する関数
 * @param {number} entryTableNumber - テーブル数
 * @param {number} entryGroup - エントリーグループ数
 */
function adjustSpecificBorders(entryTableNumber, entryGroup) {
  const k = entryTableNumber / 2;
  
  // 値の取得を一括で行うための配列を準備
  const rangeValues = {};
  
  for (let i = 1; i < k; i++) {
    const key1 = `${4 * i - 2 + IRP}_${ICP + 2}`;
    const key2 = `${4 * (i - 1) + IRP}_${ICP + 2}`;
    
    rangeValues[key1] = sht1.getRange(4 * i - 2 + IRP, ICP + 2).getValue();
    rangeValues[key2] = sht1.getRange(4 * (i - 1) + IRP, ICP + 2).getValue();
  }
  
  // 罫線を設定
  for (let i = 1; i < k; i++) {
    const key1 = `${4 * i - 2 + IRP}_${ICP + 2}`;
    const key2 = `${4 * (i - 1) + IRP}_${ICP + 2}`;
    
    if (rangeValues[key1] === "" || rangeValues[key2] === "") {
      sht1.getRange(i * 4 - 3 + IRP, ICP + 3, 2, 1)
        .setBorder(false, false, false, false, false, true, "black", SpreadsheetApp.BorderStyle.SOLID);
    }
  }
  
  // 削除対象のセルを特定
  for (let n = 1; n <= entryTableNumber / 2; n++) {
    const i = entryTableNumber / 2 - n + 1;
    let checkOut = 0;
    
    const val1 = sht1.getRange(4 * i + IRP - 5 + 1, ICP + 2).getValue();
    const val2 = sht1.getRange(4 * i + IRP - 5 + 3, ICP + 2).getValue();
    
    if (val1 === "") {
      checkOut = 1;
    }
    if (val2 === "") {
      checkOut = 2;
    }
    
    if (checkOut === 1 || checkOut === 2) {
      sht1.getRange(4 * i + IRP - 3, ICP, 2, entryGroup + 5)
        .deleteCells(SpreadsheetApp.Dimension.ROWS);
      sht1.getRange(4 * i + IRP - 4, ICP + 3, 2, 2)
        .setBorder(null, null, null, null, null, true, "black", SpreadsheetApp.BorderStyle.SOLID);
    }
  }
}

/**
 * セルをマージする関数
 * @param {number} entryTotalNumber - 参加者総数
 */
function mergeCells(entryTotalNumber) {
  // マージ前に値をクリアする対象を特定
  const rangesToClear = [];
  
  for (let i = 1; i <= entryTotalNumber; i++) {
    if (sht1.getRange(2 * i + IRP - 2, ICP + 2).getValue() !== "") {
      rangesToClear.push({
        row: 2 * i + IRP - 1,
        col: ICP + 2
      });
      
      rangesToClear.push({
        row: 2 * i + IRP - 1,
        col: ICP
      });
    }
  }
  
  // 値を一括でクリア
  for (const range of rangesToClear) {
    sht1.getRange(range.row, range.col).clearContent();
  }
  
  // セルをマージ
  for (let i = 1; i <= entryTotalNumber; i++) {
    sht1.getRange(2 * i + IRP - 2, ICP + 2, 2, 1).merge();
    sht1.getRange(2 * i + IRP - 2, ICP + 1, 2, 1).merge();
    sht1.getRange(2 * i + IRP - 2, ICP, 2, 1).merge();
  }
}

/**
 * 右側のトーナメント表を描画する関数
 * @param {number} entryGroup - エントリーグループ数
 * @param {number} entryTableNumber - テーブル数
 * @param {number} entryTotalNumber - 参加者総数
 * @param {Array} tableSet - テーブルセット配列
 * @param {number} tp - 特定の位置
 */
function drawRightTournament(entryGroup, entryTableNumber, entryTotalNumber, tableSet, tp) {
  // 右側のトーナメント表の値を設定するためのデータを準備
  const rightTournamentData = [];
  
  let pn = 1;
  for (let i = 1; i <= entryTableNumber; i++) {
    if (tableSet[entryGroup][i - 1] !== 0) {
      rightTournamentData.push({
        row: i * 2 - 2 + IRP,
        col: 2 * entryGroup + ICP + 8,
        value: pn
      });
      
      rightTournamentData.push({
        row: i * 2 - 1 + IRP,
        col: 2 * entryGroup + ICP + 8,
        value: pn
      });
      
      rightTournamentData.push({
        row: i * 2 - 2 + IRP,
        col: 2 * entryGroup + ICP + 6,
        value: tableSet[entryGroup][i - 1]
      });
      
      rightTournamentData.push({
        row: i * 2 - 1 + IRP,
        col: 2 * entryGroup + ICP + 6,
        value: tableSet[entryGroup][i - 1]
      });
      
      pn++;
    }
  }
  
  // 値を一括で設定
  for (const data of rightTournamentData) {
    sht1.getRange(data.row, data.col).setValue(data.value);
  }
  
  // 右側の罫線の描画
  let k = entryTableNumber / 4;
  let l = entryTableNumber / k;
  
  for (let h = 0; h < entryGroup + 1; h++) {
    for (let i = 0; i < k * 2; i++) {
      sht1.getRange(2 ** h + l * i + IRP, 2 * entryGroup + ICP + 5 - h, l / 2, 1)
        .setBorder(true, true, true, null, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
    }
    l = l * 2;
    k = k / 2;
  }
  
  // 右側の特定の罫線の調整
  for (let n = 1; n <= entryTableNumber / 2; n++) {
    const i = entryTableNumber / 2 - n + 1;
    let checkOut = 0;
    
    const val1 = sht1.getRange(4 * i + IRP - 5 + 1, 2 * entryGroup + ICP + 6).getValue();
    const val2 = sht1.getRange(4 * i + IRP - 5 + 3, 2 * entryGroup + ICP + 6).getValue();
    
    if (val1 === "") {
      checkOut = 1;
    }
    if (val2 === "") {
      checkOut = 2;
    }
    
    if (checkOut === 1 || checkOut === 2) {
      sht1.getRange(4 * i + IRP - 3, entryGroup + ICP + 5, 2, entryGroup + 4)
        .deleteCells(SpreadsheetApp.Dimension.ROWS);
      sht1.getRange(4 * i + IRP - 4, 2 * entryGroup + ICP + 4, 2, 2)
        .setBorder(null, null, null, null, null, true, "black", SpreadsheetApp.BorderStyle.SOLID);
    }
  }
  
  // 右側のセルのマージと調整
  for (let i = 1; i <= entryTableNumber; i++) {
    if (sht1.getRange(2 * i + IRP - 2, 2 * entryGroup + ICP + 6).getValue() === "") {
      sht1.getRange(2 * i + IRP - 1, 2 * entryGroup + ICP + 6).getValue() === "";
      sht1.getRange(2 * i + IRP - 1, 2 * entryGroup + ICP + 8).getValue() === "";
    }
    
    sht1.getRange(2 * i + IRP - 2, 2 * entryGroup + ICP + 8, 2, 1).merge();
    sht1.getRange(2 * i + IRP - 2, 2 * entryGroup + ICP + 7, 2, 1).merge();
    sht1.getRange(2 * i + IRP - 2, 2 * entryGroup + ICP + 6, 2, 1).merge();
  }
  
  // 特定の行の削除
  sht1.getRange(IRP + tp * 2, ICP, (entryTotalNumber - tp) * 2 + 1, entryGroup + 5)
    .deleteCells(SpreadsheetApp.Dimension.ROWS);
  sht1.getRange(IRP, ICP + entryGroup + 4, 2 * tp, 2 * entryGroup - entryGroup + 5)
    .deleteCells(SpreadsheetApp.Dimension.ROWS);
  
  // 列幅と書式の設定
  sht1.setColumnWidth(ICP + 2 * entryGroup + 7, 150);
  sht1.getRange(IRP, ICP + 2 * entryGroup + 6, entryTotalNumber * 4 + 1, 3)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  
  sht1.setColumnWidth(ICP + entryGroup + 4, 20);
  sht1.getRange(IRP, ICP + entryGroup + 3, entryTotalNumber * 4 + 1, 3)
    .setBorder(false, null, false, null, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  
  // 特定の罫線の設定
  sht1.getRange(tp + IRP - 1, ICP + entryGroup + 3, 1, 3)
    .setBorder(false, null, true, null, null, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  
  sht1.getRange(tp + IRP - 3, ICP + entryGroup + 4, 6, 1)
    .merge()
    .setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
}
