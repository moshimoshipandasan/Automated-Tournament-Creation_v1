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
const INITIAL_ROW_POSITION = 3; // 初期行位置
const INITIAL_COLUMN_POSITION = 1; // 初期列位置
const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
const tournamentDataSheet = spreadsheet.getSheetByName('大会データ');
const tournamentSheet = spreadsheet.getSheetByName('トーナメント');
const blockSheet = spreadsheet.getSheetByName('ブロック');
const tableSheet = spreadsheet.getSheetByName('テーブル');

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
    const participantCount = parseInt(response.getResponseText());
    
    if (isNaN(participantCount) || participantCount < 2) {
      ui.alert('エラー', '有効な参加人数を入力してください（2以上の整数）', ui.ButtonSet.OK);
      return;
    }
    
    // 参加人数を保存
    tournamentDataSheet.getRange(3, 2).setValue(participantCount);
    
    // シートの初期化
    clearSheet(tournamentSheet);
    blockSheet.clear();
    tableSheet.clear();
    
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
    clearSheet(tournamentSheet);
    blockSheet.clear();
    tableSheet.clear();
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
    const participantCount = tournamentDataSheet.getRange(3, 2).getValue();
    const tournamentRoundCount = Math.floor(Math.log2(participantCount - 1));
    const matchCount = 2 ** (tournamentRoundCount + 1);
    
    // テーブルと区画の初期化
    const matchupTable = initializeArray(tournamentRoundCount + 1, matchCount);
    const blockTable = initializeArray(tournamentRoundCount + 1, matchCount);
    
    // 初期値の設定（シード順）
    matchupTable[0][0] = 1;
    matchupTable[0][1] = 4;
    matchupTable[0][2] = 3;
    matchupTable[0][3] = 2;
    
    // テーブルの設定
    setupMatchupTable(matchupTable, tournamentRoundCount, matchCount);
    
    // 最終グループの設定
    for (let i = 0; i < matchCount; i++) {
      matchupTable[tournamentRoundCount][i] = matchupTable[tournamentRoundCount - 1][i];
      if (matchupTable[tournamentRoundCount - 1][i] > participantCount) {
        matchupTable[tournamentRoundCount][i] = 0;
      }
    }
    
    // トーナメント表の描画
    drawTournament(matchupTable, tournamentRoundCount, participantCount, matchCount);
    
    // 完了メッセージ
    Browser.msgBox("トーナメントの作成が完了しました。");
  } catch (error) {
    Browser.msgBox("エラーが発生しました: " + error.message);
    console.error(error);
  }
}

/**
 * マッチアップテーブルを設定する関数
 * @param {Array} matchupTable - マッチアップテーブル配列
 * @param {number} tournamentRoundCount - トーナメントのラウンド数
 * @param {number} matchCount - 試合数
 */
function setupMatchupTable(matchupTable, tournamentRoundCount, matchCount) {
  for (let round = 0; round < tournamentRoundCount - 1; round++) {
    for (let matchIndex = 0; matchIndex < matchCount / 2; matchIndex++) {
      if (matchupTable[round][matchIndex] !== 0) {
        matchupTable[round + 1][matchIndex * 2] = matchupTable[round][matchIndex];
        matchupTable[round + 1][matchIndex * 2 + 1] = Math.abs(2 ** (round + 3) + 1 - matchupTable[round][matchIndex]);
        
        // 特定の位置の値を入れ替え（バランス調整）
        for (let swapIndex = 2; swapIndex <= matchCount - 1; swapIndex = swapIndex + 4) {
          const tempValue = matchupTable[round + 1][swapIndex];
          matchupTable[round + 1][swapIndex] = matchupTable[round + 1][swapIndex + 1];
          matchupTable[round + 1][swapIndex + 1] = tempValue;
        }
      }
    }
  }
}

/**
 * トーナメント表を描画する関数
 * @param {Array} matchupTable - マッチアップテーブル配列
 * @param {number} tournamentRoundCount - トーナメントのラウンド数
 * @param {number} participantCount - 参加者数
 * @param {number} matchCount - 試合数
 */
function drawTournament(matchupTable, tournamentRoundCount, participantCount, matchCount) {
  // シートのアクティブ化と書式設定
  tournamentSheet.activate();
  tournamentSheet.setColumnWidth(INITIAL_COLUMN_POSITION + 1, 150);
  tournamentSheet.getRange(INITIAL_ROW_POSITION, INITIAL_COLUMN_POSITION, participantCount * 4 * 1, 3)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  
  // トーナメント番号の設定
  let matchNumber = 1;
  for (let i = 0; i < matchCount; i++) {
    if (matchupTable[tournamentRoundCount][i] !== 0) {
      // テーブルシートへの記録
      tableSheet.getRange(matchupTable[tournamentRoundCount][i], 1).setValue(matchupTable[tournamentRoundCount][i]);
      tableSheet.getRange(matchupTable[tournamentRoundCount][i], 2).setValue(matchNumber);
      
      // トーナメントシートへの記録
      tournamentSheet.getRange(i * 2 + INITIAL_ROW_POSITION, INITIAL_COLUMN_POSITION).setValue(matchNumber);
      tournamentSheet.getRange(i * 2 + 1 + INITIAL_ROW_POSITION, INITIAL_COLUMN_POSITION).setValue(matchNumber);
      matchNumber++;
      
      tournamentSheet.getRange(i * 2 + INITIAL_ROW_POSITION, INITIAL_COLUMN_POSITION + 2).setValue(matchupTable[tournamentRoundCount][i]);
      tournamentSheet.getRange(i * 2 + 1 + INITIAL_ROW_POSITION, INITIAL_COLUMN_POSITION + 2).setValue(matchupTable[tournamentRoundCount][i]);
    }
  }
  
  // 罫線の描画
  drawBracketLines(tournamentRoundCount, matchCount);
  
  // ブロック情報の設定
  setupBlockInfo(tournamentRoundCount, matchCount, matchupTable);
  
  // 特定の罫線の調整
  adjustSpecificBracketLines(matchCount, tournamentRoundCount);
  
  // セルのマージと調整
  mergeCells(participantCount);
  
  // 特定の位置の処理（中央の位置を特定）
  let centerPosition = 0;
  for (let i = 1; i < participantCount; i++) {
    if (tournamentSheet.getRange(2 * i + INITIAL_ROW_POSITION - 2, INITIAL_COLUMN_POSITION + 2).getValue() === 3) {
      centerPosition = i - 1;
    }
  }
  
  // 特定の罫線の設定
  tournamentSheet.getRange(2 * centerPosition + INITIAL_ROW_POSITION - 1, INITIAL_COLUMN_POSITION + tournamentRoundCount + 4, 2, 1)
    .setBorder(null, null, null, null, null, true, "black", SpreadsheetApp.BorderStyle.SOLID);
  
  // 右側のトーナメント表の描画
  drawRightBracket(tournamentRoundCount, matchCount, participantCount, matchupTable, centerPosition);
  
  // 行の高さを一括で設定
  setUniformRowHeights(participantCount);
  
  // シートの先頭にフォーカスを移動
  tournamentSheet.getRange(1, 1).activate();
}

/**
 * 行の高さを一括で設定する関数
 * @param {number} participantCount - 参加者数
 */
function setUniformRowHeights(participantCount) {
  // すべての行に対して一括で高さを設定
  tournamentSheet.setRowHeights(INITIAL_ROW_POSITION, participantCount * 2, 20);
}

/**
 * 罫線を描画する関数
 * @param {number} tournamentRoundCount - トーナメントのラウンド数
 * @param {number} matchCount - 試合数
 */
function drawBracketLines(tournamentRoundCount, matchCount) {
  let groupSize = matchCount / 4;
  let groupSpacing = matchCount / groupSize;
  
  for (let round = 0; round < tournamentRoundCount + 1; round++) {
    for (let group = 0; group < groupSize * 2; group++) {
      tournamentSheet.getRange(2 ** round + groupSpacing * group + INITIAL_ROW_POSITION, INITIAL_COLUMN_POSITION + 3 + round, groupSpacing / 2, 1)
        .setBorder(true, null, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
    }
    groupSpacing = groupSpacing * 2;
    groupSize = groupSize / 2;
  }
}

/**
 * ブロック情報を設定する関数
 * @param {number} tournamentRoundCount - トーナメントのラウンド数
 * @param {number} matchCount - 試合数
 * @param {Array} matchupTable - マッチアップテーブル配列
 */
function setupBlockInfo(tournamentRoundCount, matchCount, matchupTable) {
  let blockCount = 0;
  const blockSize = 1;
  const blockSizePower = 2 ** blockSize;
  
  for (let i = 0; i < 2 ** (tournamentRoundCount + 1); i++) {
    if (matchupTable[tournamentRoundCount][i] !== 0) {
      blockCount++;
    }
    if (i % blockSizePower === 1) {
      blockSheet.getRange((i + 1) / blockSizePower, 1).setValue(blockCount);
    }
  }
}

/**
 * 特定の罫線を調整する関数
 * @param {number} matchCount - 試合数
 * @param {number} tournamentRoundCount - トーナメントのラウンド数
 */
function adjustSpecificBracketLines(matchCount, tournamentRoundCount) {
  const halfMatchCount = matchCount / 2;
  
  // 空白セルの罫線調整
  for (let i = 1; i < halfMatchCount; i++) {
    if (tournamentSheet.getRange(4 * i - 2 + INITIAL_ROW_POSITION, INITIAL_COLUMN_POSITION + 2).getValue() === "" || 
        tournamentSheet.getRange(4 * (i - 1) + INITIAL_ROW_POSITION, INITIAL_COLUMN_POSITION + 2).getValue() === "") {
      tournamentSheet.getRange(i * 4 - 3 + INITIAL_ROW_POSITION, INITIAL_COLUMN_POSITION + 3, 2, 1)
        .setBorder(false, false, false, false, false, true, "black", SpreadsheetApp.BorderStyle.SOLID);
    }
  }
  
  // 不要な行の削除と罫線調整
  for (let n = 1; n <= halfMatchCount; n++) {
    const i = halfMatchCount - n + 1;
    let emptyStatus = 0;
    
    if (tournamentSheet.getRange(4 * i + INITIAL_ROW_POSITION - 5 + 1, INITIAL_COLUMN_POSITION + 2).getValue() === "") {
      emptyStatus = 1;
    }
    if (tournamentSheet.getRange(4 * i + INITIAL_ROW_POSITION - 5 + 3, INITIAL_COLUMN_POSITION + 2).getValue() === "") {
      emptyStatus = 2;
    }
    
    if (emptyStatus === 1 || emptyStatus === 2) {
      tournamentSheet.getRange(4 * i + INITIAL_ROW_POSITION - 3, INITIAL_COLUMN_POSITION, 2, tournamentRoundCount + 5)
        .deleteCells(SpreadsheetApp.Dimension.ROWS);
      tournamentSheet.getRange(4 * i + INITIAL_ROW_POSITION - 4, INITIAL_COLUMN_POSITION + 3, 2, 2)
        .setBorder(null, null, null, null, null, true, "black", SpreadsheetApp.BorderStyle.SOLID);
    }
  }
}

/**
 * セルをマージする関数
 * @param {number} participantCount - 参加者数
 */
function mergeCells(participantCount) {
  for (let i = 1; i <= participantCount; i++) {
    if (tournamentSheet.getRange(2 * i + INITIAL_ROW_POSITION - 2, INITIAL_COLUMN_POSITION + 2).getValue() !== "") {
      tournamentSheet.getRange(2 * i + INITIAL_ROW_POSITION - 1, INITIAL_COLUMN_POSITION + 2).clearContent();
      tournamentSheet.getRange(2 * i + INITIAL_ROW_POSITION - 1, INITIAL_COLUMN_POSITION).clearContent();
    }
    
    tournamentSheet.getRange(2 * i + INITIAL_ROW_POSITION - 2, INITIAL_COLUMN_POSITION + 2, 2, 1).merge();
    tournamentSheet.getRange(2 * i + INITIAL_ROW_POSITION - 2, INITIAL_COLUMN_POSITION + 1, 2, 1).merge();
    tournamentSheet.getRange(2 * i + INITIAL_ROW_POSITION - 2, INITIAL_COLUMN_POSITION, 2, 1).merge();
  }
}

/**
 * 右側のトーナメント表を描画する関数
 * @param {number} tournamentRoundCount - トーナメントのラウンド数
 * @param {number} matchCount - 試合数
 * @param {number} participantCount - 参加者数
 * @param {Array} matchupTable - マッチアップテーブル配列
 * @param {number} centerPosition - 中央の位置
 */
function drawRightBracket(tournamentRoundCount, matchCount, participantCount, matchupTable, centerPosition) {
  let matchNumber = 1;
  
  // 右側のトーナメント表の設定
  for (let i = 1; i <= matchCount; i++) {
    if (matchupTable[tournamentRoundCount][i - 1] !== 0) {
      tournamentSheet.getRange(i * 2 - 2 + INITIAL_ROW_POSITION, 2 * tournamentRoundCount + INITIAL_COLUMN_POSITION + 8).setValue(matchNumber);
      tournamentSheet.getRange(i * 2 - 1 + INITIAL_ROW_POSITION, 2 * tournamentRoundCount + INITIAL_COLUMN_POSITION + 8).setValue(matchNumber);
      matchNumber++;
      
      tournamentSheet.getRange(i * 2 - 2 + INITIAL_ROW_POSITION, 2 * tournamentRoundCount + INITIAL_COLUMN_POSITION + 6).setValue(matchupTable[tournamentRoundCount][i - 1]);
      tournamentSheet.getRange(i * 2 - 1 + INITIAL_ROW_POSITION, 2 * tournamentRoundCount + INITIAL_COLUMN_POSITION + 6).setValue(matchupTable[tournamentRoundCount][i - 1]);
    }
  }
  
  // 右側の罫線の描画
  let groupSize = matchCount / 4;
  let groupSpacing = matchCount / groupSize;
  
  for (let round = 0; round < tournamentRoundCount + 1; round++) {
    for (let group = 0; group < groupSize * 2; group++) {
      tournamentSheet.getRange(2 ** round + groupSpacing * group + INITIAL_ROW_POSITION, 2 * tournamentRoundCount + INITIAL_COLUMN_POSITION + 5 - round, groupSpacing / 2, 1)
        .setBorder(true, true, true, null, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
    }
    groupSpacing = groupSpacing * 2;
    groupSize = groupSize / 2;
  }
  
  // 右側の特定の罫線の調整
  for (let n = 1; n <= matchCount / 2; n++) {
    const i = matchCount / 2 - n + 1;
    let emptyStatus = 0;
    
    if (tournamentSheet.getRange(4 * i + INITIAL_ROW_POSITION - 5 + 1, 2 * tournamentRoundCount + INITIAL_COLUMN_POSITION + 6).getValue() === "") {
      emptyStatus = 1;
    }
    if (tournamentSheet.getRange(4 * i + INITIAL_ROW_POSITION - 5 + 3, 2 * tournamentRoundCount + INITIAL_COLUMN_POSITION + 6).getValue() === "") {
      emptyStatus = 2;
    }
    
    if (emptyStatus === 1 || emptyStatus === 2) {
      tournamentSheet.getRange(4 * i + INITIAL_ROW_POSITION - 3, tournamentRoundCount + INITIAL_COLUMN_POSITION + 5, 2, tournamentRoundCount + 4)
        .deleteCells(SpreadsheetApp.Dimension.ROWS);
      tournamentSheet.getRange(4 * i + INITIAL_ROW_POSITION - 4, 2 * tournamentRoundCount + INITIAL_COLUMN_POSITION + 4, 2, 2)
        .setBorder(null, null, null, null, null, true, "black", SpreadsheetApp.BorderStyle.SOLID);
    }
  }
  
  // 右側のセルのマージと調整
  for (let i = 1; i <= matchCount; i++) {
    if (tournamentSheet.getRange(2 * i + INITIAL_ROW_POSITION - 2, 2 * tournamentRoundCount + INITIAL_COLUMN_POSITION + 6).getValue() === "") {
      tournamentSheet.getRange(2 * i + INITIAL_ROW_POSITION - 1, 2 * tournamentRoundCount + INITIAL_COLUMN_POSITION + 6).getValue() === "";
      tournamentSheet.getRange(2 * i + INITIAL_ROW_POSITION - 1, 2 * tournamentRoundCount + INITIAL_COLUMN_POSITION + 8).getValue() === "";
    }
    
    tournamentSheet.getRange(2 * i + INITIAL_ROW_POSITION - 2, 2 * tournamentRoundCount + INITIAL_COLUMN_POSITION + 8, 2, 1).merge();
    tournamentSheet.getRange(2 * i + INITIAL_ROW_POSITION - 2, 2 * tournamentRoundCount + INITIAL_COLUMN_POSITION + 7, 2, 1).merge();
    tournamentSheet.getRange(2 * i + INITIAL_ROW_POSITION - 2, 2 * tournamentRoundCount + INITIAL_COLUMN_POSITION + 6, 2, 1).merge();
  }
  
  // 特定の行の削除
  tournamentSheet.getRange(INITIAL_ROW_POSITION + centerPosition * 2, INITIAL_COLUMN_POSITION, (participantCount - centerPosition) * 2 + 1, tournamentRoundCount + 5)
    .deleteCells(SpreadsheetApp.Dimension.ROWS);
  tournamentSheet.getRange(INITIAL_ROW_POSITION, INITIAL_COLUMN_POSITION + tournamentRoundCount + 4, 2 * centerPosition, 2 * tournamentRoundCount - tournamentRoundCount + 5)
    .deleteCells(SpreadsheetApp.Dimension.ROWS);
  
  // 列幅と書式の設定
  tournamentSheet.setColumnWidth(INITIAL_COLUMN_POSITION + 2 * tournamentRoundCount + 7, 150);
  tournamentSheet.getRange(INITIAL_ROW_POSITION, INITIAL_COLUMN_POSITION + 2 * tournamentRoundCount + 6, participantCount * 4 + 1, 3)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  
  tournamentSheet.setColumnWidth(INITIAL_COLUMN_POSITION + tournamentRoundCount + 4, 20);
  tournamentSheet.getRange(INITIAL_ROW_POSITION, INITIAL_COLUMN_POSITION + tournamentRoundCount + 3, participantCount * 4 + 1, 3)
    .setBorder(false, null, false, null, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  
  // 特定の罫線の設定
  tournamentSheet.getRange(centerPosition + INITIAL_ROW_POSITION - 1, INITIAL_COLUMN_POSITION + tournamentRoundCount + 3, 1, 3)
    .setBorder(false, null, true, null, null, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  
  tournamentSheet.getRange(centerPosition + INITIAL_ROW_POSITION - 3, INITIAL_COLUMN_POSITION + tournamentRoundCount + 4, 6, 1)
    .merge()
    .setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
}
