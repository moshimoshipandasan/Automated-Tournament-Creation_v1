/*
内容：組合せシステムVer7
作成日：2021/11/02
作成・著作者：安藤昇＠青山学院中等部
備考：プログラムの配布・改変などする場合にはメールにてご連絡ください
email:gigaschool2020@gmail.com（安藤宛）
*/
var EntryTotalNumber;
var EntryGroup;
var EntryTableNumber;
var IRP = 3;
var ICP = 1;
var LN;
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sht0 = ss.getSheetByName('大会データ');
var sht1 = ss.getSheetByName('トーナメント');
var sht2 = ss.getSheetByName('ブロック');
var sht3 = ss.getSheetByName('参加データ');
var sht4 = ss.getSheetByName('シード入力');
var sht5 = ss.getSheetByName('テーブル');
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('組合せ');
  menu.addItem('1.トーナメント作成', 'SheetSetting');
  menu.addItem('2.シード抽選', 'ChusenStart');
  menu.addItem('3.本抽選（ランダムに抽選）', 'ChusenStart1');
  menu.addItem('0.シートクリア', 'SheetClear');
  menu.addToUi();
}
function SheetSetting(){
  var msg1 = Browser.msgBox("トーナメント作成します。よろしいですか？", Browser.Buttons.OK_CANCEL);
  if(msg1 == "cancel"){
    ;
  }else{
    sht1
      .clear()
      .setHiddenGridlines(true)
      .setColumnWidths(1, 26, 25)
      .setRowHeights(1, 100, 10);
    sht2.clear();
    sht4.clear();
    CreateTournament()
  }
}
function SheetClear(){
  var msg1 = Browser.msgBox("シートを初期化します", Browser.Buttons.OK_CANCEL);
  if(msg1 == "cancel"){
    ;
  }else{
    sht1
      .clear()
      .setHiddenGridlines(true)
      .setColumnWidths(1, 26, 25)
      .setRowHeights(1, 100, 10);
    sht2.clear();
    sht4.clear();
    sht5.clear();
  }
}
function SankaSort(){
  sht3.getRange(1,66).setValue('合計')//.setBackground('cyan');
  sht3.getRange(2,66,1000,1).clearContent()
  var lastRow = sht3.getLastRow();
  var lastCol = sht3.getLastColumn();
  for(var i = 2;i <= lastRow;i++){
    var strformula = "=counta(RC[" +(2 - lastCol) + "]:RC[-1])";
    sht3.getRange(i,66).setFormulaR1C1(strformula);
  }
  sht3.getRange(1,1,lastRow,66).activate();
  var currentCell = sht3.getCurrentCell();
  sht3.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  sht3.getActiveRange().offset(1, 0, sht3.getActiveRange().getNumRows() - 1).sort({column: 66, ascending: false});
  var data3 =  sht3.getRange(1,1,lastRow,66).getValues();
  var n = 2;
  sht4.clear();
  sht4.appendRow(["no","code", "名称", "所属", "シード番号", "トーナメント番号"]);
  for(var i = 1;i < lastRow;i++){
    for(var j = 1;j <= data3[i][65];j++){
      sht4.getRange(n,1).setValue(n - 1);
      sht4.getRange(n,2).setValue(i);
      sht4.getRange(n,3).setValue(data3[i][j]);
      sht4.getRange(n,4).setValue(data3[i][0]);
      n++;
    }
  }
  sht0.getRange(3,2).setValue(n - 2);
}
function CreateTournament() {
  SankaSort();
  var TT =1;
  var TempBox;
  EntryTotalNumber  = sht0.getRange(3,2).getValue();
  EntryGroup = Math.floor(Math.log2(EntryTotalNumber - 1));
  EntryTableNumber = 2 ** (EntryGroup + 1);
  var TableSet = new Array(EntryGroup + 1);
  var BlockSet = new Array(EntryGroup + 1);
  for (var h = 0; h < TableSet.length; h++) {
    TableSet[h] = new Array(EntryTableNumber);
    BlockSet[h] = new Array(EntryTableNumber);
  }
  for (h = 0; h < TableSet.length; h++) {
    for (i = 0; i < TableSet[h].length; i++) {
      TableSet[h][i] = 0;
      BlockSet[h][i] = 0;
    }
  }
    TableSet[0][0] = 1;
    TableSet[0][1] = 4;
    TableSet[0][2] = 3;
    TableSet[0][3] = 2;
  for(h = 0; h < EntryGroup - 1; h++){
    for(i = 0; i < EntryTableNumber/2; i++){
      if(TableSet[h][i] !== 0){
        TableSet[h + 1][i * 2] = TableSet[h][i];
        TableSet[h + 1][i * 2 + 1] = Math.abs(2 ** (h + 3) + 1 - TableSet[h][i]);
        for(j = 2;j <= EntryTableNumber - 1; j = j + 4){
          TempBox = TableSet[h + 1][j];
          TableSet[h + 1][j] = TableSet[h + 1][j + 1];
          TableSet[h + 1][j + 1] = TempBox;
        }
      }
    }
  } 
  for (i = 0; i < EntryTableNumber; i++){
    TableSet[EntryGroup][i] = TableSet[EntryGroup -　1][i];
    if(TableSet[EntryGroup　-　1][i] > EntryTotalNumber) {
      TableSet[EntryGroup][i] = 0;
    }
  }
  var PN = 1;
  sht1.activate();
  sht1.setColumnWidth(ICP + 1, 150);
  sht1.getRange(IRP,ICP,EntryTotalNumber * 4 * 1,3)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  for (i = 0; i < EntryTableNumber; i++){
    if(TableSet[EntryGroup][i] !== 0) {
      sht5.getRange(TableSet[EntryGroup][i], 1).setValue(TableSet[EntryGroup][i]);
      sht5.getRange(TableSet[EntryGroup][i], 2).setValue(PN);
      sht1.getRange(i * 2 + IRP, ICP).setValue(PN);
      sht1.getRange(i * 2 + 1 + IRP, ICP).setValue(PN);
      PN++;
      sht1.getRange(i * 2 + IRP, ICP + 2).setValue(TableSet[EntryGroup][i]);
      sht1.getRange(i * 2 + 1 + IRP, ICP + 2).setValue(TableSet[EntryGroup][i]);
    }
  }
  var k = EntryTableNumber / 4;
  var l = EntryTableNumber / k;
  for(h = 0; h < EntryGroup + 1; h++){
    for(i = 0; i < k * 2; i++){
      sht1.getRange(2 ** h + l * i + IRP,ICP + 3 + h,l / 2,1)
      .setBorder(true, null, true, true, false, false,"black",SpreadsheetApp.BorderStyle.SOLID); //細線
    }
    l = l * 2
    k = k / 2
  }
    var n = 0;
    j = 1;
    j2 = 2 ** j;
      for(i = 0; i < 2 ** (EntryGroup + 1); i++){
        if(TableSet[EntryGroup][i] !== 0){
          n++
        }
        if(i % j2 === 1){
          sht2.getRange((i + 1) / j2, 1).setValue(n);
        }
      }
  k = EntryTableNumber / 2
  for(i = 1; i < k; i++){
    if(sht1.getRange(4 * i - 2 + IRP, ICP + 2).getValue() === "" || sht1.getRange(4 * (i - 1) + IRP, ICP + 2).getValue() === ""){
      sht1.getRange(i * 4 - 3 + IRP,ICP + 3,2,1).activate()
        .setBorder(false, false, false, false, false, true,"black",SpreadsheetApp.BorderStyle.SOLID);
    }  
  }
  for(var n = 1; n <= EntryTableNumber / 2; n++){
    i = EntryTableNumber / 2 - n + 1
    if(sht1.getRange(4 * i + IRP - 5 + 1, ICP + 2).getValue() === ""){
      var CheckOut = 1
    }
    if(sht1.getRange(4 * i + IRP - 5 + 3, ICP + 2).getValue() === ""){
      var CheckOut = 2
    }
    if(CheckOut === 1 || CheckOut === 2){
      sht1.getRange(4 * i + IRP - 3, ICP, 2, EntryGroup + 5)
        .deleteCells(SpreadsheetApp.Dimension.ROWS);
      sht1.getRange(4 * i + IRP - 4, ICP + 3, 2, 2)
        .setBorder(null, null, null, null, null, true,"black",SpreadsheetApp.BorderStyle.SOLID);
      CheckOut = 0
    }
  }
  for(var i = 1; i <= EntryTotalNumber; i++){
    if(sht1.getRange(2 * i + IRP - 2, ICP + 2).getValue() !== ""){
      sht1.getRange(2 * i + IRP - 1, ICP + 2).clearContent();
      sht1.getRange(2 * i + IRP - 1, ICP).clearContent();
    }
    sht1.getRange(2 * i + IRP - 2, ICP + 2,2,1).merge();
    sht1.getRange(2 * i + IRP - 2, ICP + 1,2,1).merge();
    sht1.getRange(2 * i + IRP - 2, ICP,2,1).merge();
  }
  for(var i = 1; i < EntryTotalNumber; i++){
    if(sht1.getRange(2 * i + IRP - 2, ICP + 2).getValue() === 3){
      var TP = i - 1;
    }
  }
  sht1.getRange(2 * TP + IRP - 1, ICP + EntryGroup + 4,2,1)
    .setBorder(null, null, null, null, null, true,"black",SpreadsheetApp.BorderStyle.SOLID); //細線
  sht1.getRange(1,1).activate();
  PN = 1;
  if(TT === 1){
    for (i = 1; i <= EntryTableNumber; i++){
      if(TableSet[EntryGroup][i - 1] !== 0) {
        sht1.getRange(i * 2 - 2 + IRP, 2 * EntryGroup + ICP + 8).setValue(PN);
        sht1.getRange(i * 2 - 1 + IRP, 2 * EntryGroup + ICP + 8).setValue(PN);
        PN++;
        sht1.getRange(i * 2 - 2 + IRP, 2 * EntryGroup + ICP + 6).setValue(TableSet[EntryGroup][i-1]);
        sht1.getRange(i * 2 - 1 + IRP,2 * EntryGroup + ICP + 6).setValue(TableSet[EntryGroup][i-1]);
      }
    }
    var k = EntryTableNumber / 4;
    var l = EntryTableNumber / k;
    for(h = 0; h < EntryGroup + 1; h++){
      for(i = 0; i < k * 2; i++){
        sht1.getRange(2 ** h + l * i + IRP, 2 * EntryGroup + ICP + 5 - h,l / 2,1)
          .setBorder(true, true, true, null, false, false,"black",SpreadsheetApp.BorderStyle.SOLID); //細線
      }
      l = l * 2
      k = k / 2
    }
  for(var n = 1; n <= EntryTableNumber / 2; n++){
    i = EntryTableNumber / 2 - n + 1
    if(sht1.getRange(4 * i + IRP - 5 + 1, 2 * EntryGroup + ICP + 6).getValue() === ""){
      var CheckOut = 1
    }
    if(sht1.getRange(4 * i + IRP - 5 + 3, 2 * EntryGroup + ICP + 6).getValue() === ""){
      var CheckOut = 2
    }
    if(CheckOut === 1 || CheckOut === 2){
      sht1.getRange(4 * i + IRP - 3, EntryGroup + ICP + 5, 2, EntryGroup + 4)
        .deleteCells(SpreadsheetApp.Dimension.ROWS);
      sht1.getRange(4 * i + IRP - 4, 2 * EntryGroup + ICP + 4, 2, 2)
        .setBorder(null, null, null, null, null, true,"black",SpreadsheetApp.BorderStyle.SOLID); //細線
      CheckOut = 0
    }
  }
  for(var i = 1; i <= EntryTableNumber; i++){
    if(sht1.getRange(2 * i + IRP - 2, 2 * EntryGroup + ICP + 6).getValue() === ""){
      sht1.getRange(2 * i + IRP - 1, 2 * EntryGroup + ICP + 6).getValue() === "";
      sht1.getRange(2 * i + IRP - 1, 2 * EntryGroup + ICP + 8).getValue() === "";
    }
    sht1.getRange(2 * i + IRP - 2, 2 * EntryGroup + ICP + 8,2,1).merge();
    sht1.getRange(2 * i + IRP - 2, 2 * EntryGroup + ICP + 7,2,1).merge();
    sht1.getRange(2 * i + IRP - 2, 2 * EntryGroup + ICP + 6,2,1).merge();
  }
    sht1.getRange(IRP + TP * 2, ICP,(EntryTotalNumber - TP) * 2 + 1,EntryGroup + 5).deleteCells(SpreadsheetApp.Dimension.ROWS);
    sht1.getRange(IRP, ICP + EntryGroup + 4,2 * TP,2 * EntryGroup - EntryGroup + 5).deleteCells(SpreadsheetApp.Dimension.ROWS);
  sht1.setColumnWidth(ICP + 2 * EntryGroup + 7, 150);
  sht1.getRange(IRP, ICP + 2 * EntryGroup + 6,EntryTotalNumber * 4 + 1,3)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sht1.setColumnWidth(ICP + EntryGroup + 4, 20);
  sht1.getRange(IRP, ICP + EntryGroup + 3,EntryTotalNumber * 4 + 1,3)
    .setBorder(false, null, false, null, false, false,"black",SpreadsheetApp.BorderStyle.SOLID); //細線
  sht1.getRange(TP + IRP - 1, ICP + EntryGroup + 3,1,3)
    .activate()
    .setBorder(false, null, true, null, null, false,"black",SpreadsheetApp.BorderStyle.SOLID); //細線
  sht1.getRange(TP + IRP - 3, ICP + EntryGroup + 4,6,1)
    .merge()
    .setBorder(true, true, true, true, false, false,"black",SpreadsheetApp.BorderStyle.SOLID); //細線
  }
  sht1.getRange(1,1).activate();
  Browser.msgBox("トーナメントの作成が終了しました。次にシード入力より、シード選手にシード番号を入力してください。");
}

function ChusenStart(){
  var BlockLastNum = [];
  var BlockNum = [];
  EntryTotalNumber = sht0.getRange(3,2).getValue();
  EntryGroup = Math.floor(Math.log2(EntryTotalNumber - 1));
  EntryTableNumber = 2 ** (EntryGroup + 1);
  var lastRow4 = sht4.getLastRow();
  var lastCol4 = sht4.getLastColumn();
  var lastRow5 = sht5.getLastRow();
  var lastCol5 = sht5.getLastColumn();
  var data4 = sht4.getRange(1,1,lastRow4,lastCol4).getValues();
  var data5 = sht5.getRange(1,1,lastRow5,lastCol5).getValues();
  for(i = 1;i <= EntryTotalNumber;i++){
    if(data4[i][4] == ""){
      sht4.getRange(i+1,5).setValue(Math.floor(Math.random()*1000+1000));
    }else{
      sht4.getRange(i+1,6).setValue(data5[data4[i][4] - 1][1]);
    }
  }
  sht4.getRange(1,2,lastRow4,5).activate();
  var currentCell = sht4.getCurrentCell();
  sht4.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  sht4.getActiveRange().offset(1, 0, sht4.getActiveRange().getNumRows() - 1).sort([{column: 2, ascending: true}, {column: 5, ascending: true}]);
  var data4 = sht4.getRange(1,1,lastRow4,lastCol4).getValues();
  for(i=1;i<=EntryTotalNumber;i++){
    if(data4[i][4] >= 1000){
      sht4.getRange(i+1,5).setValue(null);
      sht4.getRange(i+1,6).setValue(null);
    }
  }
  var BlockNum = new Array(EntryGroup + 1);
  for (var h = 0; h < BlockNum.length; h++) {
    BlockNum[h] = new Array(EntryTableNumber);
  }
  var lastRow = sht2.getLastRow();
  for (h = 0; h < BlockNum.length; h++) {
    for (i = 0; i < BlockNum[h].length; i++) {
      BlockNum[h][i] = 0;
    }
  }
  for(i = 0;i < lastRow;i++){
     BlockLastNum[i] = sht2.getRange(i + 1,1).getValue();
  }
  var BlockMaxNum = lastRow;
  var Nmax = Math.floor(Math.log2(BlockMaxNum))　+　1;
  for(var n = 1;n <= Nmax;n++){
    for(var i =1;i <= BlockLastNum[BlockMaxNum - 1];i++){
      var StepCnt = -1 * (2 ** Nmax) * (1 / 2) ** n;
      var nn = Math.abs(BlockMaxNum / StepCnt);
      for(var j = 1; j <= nn; j++){ //逆ステップ
        var l = BlockMaxNum + (StepCnt * (j - 1));
        if(i <= BlockLastNum[l - 1]){
          BlockNum[n - 1][i - 1] = nn - (j - 1);
        }
      }
    }
  }
  var syumoku  = sht0.getRange(2,2).getValue();
  if(syumoku == '個人'){
    var ff = 0;//個人戦
  }else{
    var ff = 1;//団体戦
  }
      sht1.getRange(1,ICP + 1,1000).clearContent();
      sht1.getRange(1,ICP + 2 * EntryGroup + 7,1000).clearContent();
  for(i = 1;i <= BlockLastNum[lastRow-1];i++){
    var m1 = sht4.getRange(i + 1,6).getValue();
    if(ff == 0){
      var m2 = sht4.getRange(i + 1,3).getValue() +"("+sht4.getRange(i + 1,4).getValue()+")";
    }else{
      var m2 = sht4.getRange(i + 1,3).getValue();
    }
    if(BlockNum[1][m1 - 1] == 1){
      sht1.getRange(2 * m1 - 1 + (IRP - 1),ICP + 1).setValue(m2);
    }else if(BlockNum[1][m1 - 1] == 2){
      sht1.getRange(2 * (m1 - BlockLastNum[lastRow / 2 - 1]) - 1 + (IRP - 1),ICP + 2 * EntryGroup + 7).setValue(m2);
    }
  }
  sht1.activate();
  sht1.getRange(1,1).activate();
  Browser.msgBox("シード選手をトーナメント入力しましたので、次にメニューの[組合せ]から[本抽選]を行ってください。");
}
function ChusenStart1(){
  var BlockLastNum =[];
  var BlockNum =[];
  EntryTotalNumber  = sht0.getRange(3,2).getValue();
  EntryGroup = Math.floor(Math.log2(EntryTotalNumber - 1));
  EntryTableNumber = 2 ** (EntryGroup + 1);
//シードの並べ替え
  var lastRow4 = sht4.getLastRow();
  var randoms = [];
  var tmp1 = [];
  var data = [];
  var ChusenNo = [];
  var data4 = sht4.getRange(2,1,lastRow4,6).getValues();
  var EntryTotalNum = data4.length;
  for(i = 1;i < EntryTotalNum;i++){    
    tmp1.push(i);
  }
  for(i = 0;i < EntryTotalNum - 1;i++){    
    if(data4[i][5] > 0){
      tmp1[data4[i][5] - 1] = 0;
    }
  }
  for(i = 0;i < tmp1.length;i++){    
    if(tmp1[i] > 0){
      ChusenNo.push(tmp1[i]);
    }
  }
  var max = ChusenNo.length;
  for(i = 1; i <= max; i++){
    while(true){
      var min = 1;
      var tmp = intRandom(min - 1, max - 1);
      if(randoms.indexOf(tmp) == -1){
        randoms.push(tmp);
        break;
      }
    }
  }
  var cnt = 0;
  for(i = 0;i < EntryTotalNum - 1;i++){    
    if(data4[i][5] == ''){
      sht4.getRange(i + 2,6).setValue(ChusenNo[randoms[cnt]]);
      cnt++;
    }
  }
  var BlockNum = new Array(EntryGroup + 1);
  for (var h = 0; h < BlockNum.length; h++) {
    BlockNum[h] = new Array(EntryTableNumber);
  }
  var lastRow = sht2.getLastRow();
  for (h = 0; h < BlockNum.length; h++) {
    for (i = 0; i < BlockNum[h].length; i++) {
      BlockNum[h][i] = 0;
    }
  }
  for(i = 0;i < lastRow;i++){
     BlockLastNum[i] = sht2.getRange(i + 1,1).getValue();
  }
  var BlockMaxNum = lastRow; //ブロック数は4,8,16,32,64・・・になる可能性がある
  var BlockHalfNum = BlockMaxNum / 2;
  var Nmax = Math.floor(Math.log2(BlockMaxNum))　+　1;//指数部分の取得

  for(var n = 1;n <= Nmax;n++){
    for(var i =1;i <= BlockLastNum[BlockMaxNum - 1];i++){
      var StepCnt = -1 * (2 ** Nmax) * (1 / 2) ** n;
      var nn = Math.abs(BlockMaxNum / StepCnt);
      for(var j = 1; j <= nn; j++){ //逆ステップ
        var l = BlockMaxNum + (StepCnt * (j - 1));
        if(i <= BlockLastNum[l - 1]){
          BlockNum[n - 1][i - 1] = nn - (j - 1);
        }
      }
    }
  }
  var syumoku  = sht0.getRange(2,2).getValue();
  if(syumoku == '個人'){
    var ff = 0;
  }else{
    var ff = 1;
  }
  for(i　=　1;i　<=　BlockLastNum[lastRow - 1];i++){
    var m1 = sht4.getRange(i+1,6).getValue();
    if(ff == 0){
      var m2 = sht4.getRange(i + 1,3).getValue() +"("+sht4.getRange(i + 1,4).getValue()+")";
    }else{
      var m2 = sht4.getRange(i + 1,3).getValue();
    }
    if(BlockNum[1][m1 - 1] == 1){
      sht1.getRange(2 * m1 - 1 + (IRP - 1),ICP + 1).setValue(m2);
    }else if(BlockNum[1][m1 - 1] == 2){
      sht1.getRange(2 * (m1 - BlockLastNum[lastRow / 2 - 1]) - 1 + (IRP - 1),ICP + 2 * EntryGroup + 7).setValue(m2);
    }
  }
  sht1.activate();
  sht1.getRange(1,1).activate();
//  sht4.getRange(1,1).activate();
  Browser.msgBox("組み合わせが終了しました");
}
function TwoMix(array1,array2){
  var array3=[];
  for(i = 0;i < array1.length;i++){
      array3[i] = array1[i] * array2[i] 
  }
  return array3;
};
function ChangeElement(tbn,array){
  var target = 0;
  while(target >= 0){
    var target = array.indexOf(tbn);
    if(target != -1){
      array.splice(target,1,0);
    }
  }
  return array;
};
function ChangeElement3(){
  var array = [0,1,0,2,0,4,5,6,7,0,8,0,9,10];
  var tbn = 1;
  var target = 0;
  while(target >= 0){
    var target = array.indexOf(0);
    if(target != -1){
      array.splice(target,1,0);
    }
  }
};
function ChangeElement2(tbn,array1,array2){
  var target = 0;
  while(target >= 0){
    var target = array1.indexOf(tbn);
    if(target != -1){
      array2.splice(target,1,0);
    }
  }
  return array2;
};
function RndElement(array){
  var array1 = [];
  for(i = 0;i < array.length;i++){
    if(array[i] == 0){
      array1[i] = 0;
    }else{
      array1[i] = i + 1;
    }
  }
  var target = 0;
  while(target >= 0){
    var target = array1.indexOf(0);
    if(target != -1){
      array1.splice(target,1);
    }
  }
  return array1[Math.floor(array1.length * Math.random())];
};
function intRandom(min, max){
  return Math.floor(Math.random() * (max - min + 1)) + min;
};
