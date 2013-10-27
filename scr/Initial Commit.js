
function fmontemorano031(){
  //  try{
  var x = 0
  var Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var Sheet = Spreadsheet.getSheets(); 
  
  var Sheet = Spreadsheet.getSheets(); 
  var row = 1; var column = 1; var numRows = Sheet[x].getRange("A:A").getLastRow(); var numColumns = Sheet[x].getRange("1:1").getLastColumn(); 
  var range = Sheet[x].getRange(row, column, numRows, numColumns);
  
  
  SpreadsheetApp.flush();
  Sheet[x].setName('frontend');
  //borders #434343
  range.clear().breakApart().setFontSize(10).setFontColor('#151515').setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(false);
  if(numColumns>25){Sheet[x].deleteColumns(26, numColumns-25);}else{if(numColumns<25){Sheet[x].getRange(30,25,1,1).setValue("");}else{}};
  if(numRows>30){Sheet[x].deleteRows(31, numRows-30);}else{if(numRows<30){Sheet[x].getRange(30,25,1,1).setValue("");}else{}};
  
  SpreadsheetApp.flush();
  var image = '=Image("https://dl.dropboxusercontent.com/u/94960692/tiles/projects/2013%2008/borders.png",3)';
  var tm = '=Image("https://dl.dropboxusercontent.com/u/94960692/tiles/projects/2013%2008/rightsReserved.png",3)';
  
  for(var q=4;q<23;q++){
    Sheet[x].setColumnWidth(q, 25);
  }
  for(var q=1;q<4;q++){
    Sheet[x].setColumnWidth(q, 100);
  }
  for(var q=23;q<25;q++){
    Sheet[x].setColumnWidth(q, 100);
  }
  
  SpreadsheetApp.flush();
  Sheet[x].getRange("1:1").setBackground('#434343');
  Sheet[x].getRange("C1:C21").setBackground('#434343');
  Sheet[x].getRange("A12:B12").setBackground('#434343');
  Sheet[x].getRange("A14:B21").setBackground('#434343');
  Sheet[x].getRange("W:W").setBackground('#434343');
  Sheet[x].getRange("D21:V21").setBackground('#434343');
  SpreadsheetApp.flush();
  Sheet[x].getRange("A1").setFormula(image);
  Sheet[x].getRange("C1").setFormula(image);
  Sheet[x].getRange("j1").setFormula(image);
  Sheet[x].getRange("Q1").setFormula(image);
  Sheet[x].getRange("X1").setFormula(image);
  Sheet[x].getRange("C2").setFormula(image);
  Sheet[x].getRange("W2").setFormula(image);
  Sheet[x].getRange("W1").setFormula(image);
  Sheet[x].getRange("W14").setFormula(image);
  Sheet[x].getRange("A12").setFormula(image);
  Sheet[x].getRange("W21").setFormula(image);
  Sheet[x].getRange("W22").setFormula(image);
  Sheet[x].getRange("A21").setFormula(image);
  Sheet[x].getRange("C21").setFormula(image);
  Sheet[x].getRange("J21").setFormula(image);
  Sheet[x].getRange("Q21").setFormula(image);
  SpreadsheetApp.flush();
  Sheet[x].getRange("A14").setFormula(tm);
  SpreadsheetApp.flush();
  Sheet[x].getRange("A2").setFormula('="activate"');
  Sheet[x].getRange("B2").setFormula('="display"');
  Sheet[x].getRange("X2").setFormula('="activate"');
  Sheet[x].getRange("Y2").setFormula('="display"');
  Sheet[x].getRange("B2").setNote('Type "x" to delete the variable from the equation and to delete the box on the diagram. Type a word to name the variable.')
  Sheet[x].getRange("Y2").setNote('Type "x" to delete the variable from the equation. The feature of naming variables has been removed.')
  Sheet[x].getRange("A13").setNote('To create a constant, type two single digit numbers in the box to the right. Once finished go to your acions and create a file.')
  Sheet[x].getRange("A3").setFormula('=TRANSPOSE({"A1","A2","A3","A4","A5","A6","A7","A8","A9"})');
  SpreadsheetApp.flush();
  Sheet[x].getRange("X3").setFormula('=IFERROR(query(backend!E11:F,"select F where E = 1"),"")');
  Sheet[x].getRange("A13").setFormula('="add constant"');
  Sheet[x].getRange("A22").setFormula('=backend!CL2');
  Sheet[x].getRange("A23").setFormula('=backend!CL3');
  Sheet[x].getRange("A24").setFormula('=backend!CL4');
  Sheet[x].getRange("A25").setFormula('=backend!CL5');
  Sheet[x].getRange("A26").setFormula('=backend!CL6');
  Sheet[x].getRange("A27").setFormula('=backend!CL7');
  Sheet[x].getRange("A28").setFormula('=backend!CL8');
  Sheet[x].getRange("A29").setFormula('=backend!CL9');
  Sheet[x].getRange("A30").setFormula('=backend!CL10');
  SpreadsheetApp.flush();
  Sheet[x].getRange(1,1,1,2).merge();
  Sheet[x].getRange(12,1,1,2).merge();
  Sheet[x].getRange(21,1,1,2).merge();
  Sheet[x].getRange(1,24,1,2).merge();
  Sheet[x].getRange(14,1,7,3).merge();
  SpreadsheetApp.flush(); 
  Sheet[x].getRange(1,3,1,7).merge();
  Sheet[x].getRange(1,10,1,7).merge();
  Sheet[x].getRange(1,17,1,6).merge();
  Sheet[x].getRange(21,3,1,7).merge();
  Sheet[x].getRange(21,10,1,7).merge();
  Sheet[x].getRange(21,17,1,6).merge();
  SpreadsheetApp.flush(); 
  Sheet[x].getRange(2,3,12,1).merge();
  Sheet[x].getRange(2,23,12,1).merge();
  Sheet[x].getRange(14,23,7,1).merge();
  Sheet[x].getRange(22,23,9,1).merge();
  SpreadsheetApp.flush(); 
  Sheet[x].getRange(22,1,1,22).merge();
  Sheet[x].getRange(23,1,1,22).merge();
  Sheet[x].getRange(24,1,1,22).merge();
  Sheet[x].getRange(25,1,1,22).merge();
  Sheet[x].getRange(26,1,1,22).merge();
  Sheet[x].getRange(27,1,1,22).merge();
  Sheet[x].getRange(28,1,1,22).merge();
  Sheet[x].getRange(29,1,1,22).merge();
  Sheet[x].getRange(30,1,1,22).merge();
  SpreadsheetApp.flush();
  var z = 2;
  for(var q = 0; q < 16; q++){
    switch(q){ 
      case 0: var RC = "E3"; break; 
      case 1: var RC = "K3"; break; 
      case 2: var RC = "Q3"; break; 
      case 3: var RC = "E9"; break; 
      case 4: var RC = "K9"; break; 
      case 5: var RC = "Q9"; break; 
      case 6: var RC = "E15"; break; 
      case 7: var RC = "K15"; break; 
      case 8: var RC = "Q15"; break; 
    }
    
    Sheet[x].getRange(RC).offset(1,1,3,3).merge();
    
    var r = Sheet[x].getRange(RC).getRow();
    var c = Sheet[x].getRange(RC).getColumn();
    var Range2 = Sheet[x].getRange(RC).offset(4,4).getA1Notation(); 
    
    var values = new Array(5);
    for (var a = 0; a < 5; a++) {
      values[a] = new Array(5); 
      switch(a){
        case 0: 
          values[a][0] ='=if(backend!$CU$'+z+'="✖","✖",if(backend!$CU$'+z+'=1,"↖",if(backend!$CU$'+z+'="","","↘")))'; var z = z +1;
          values[a][1] ='=if(backend!$CU$'+z+'="✖","✖",if(backend!$CU$'+z+'=1,"↑",if(backend!$CU$'+z+'="","","↓")))'; var z = z +1;
          values[a][2] ='=if(backend!$CU$'+z+'="✖","✖",if(backend!$CU$'+z+'=1,"↑",if(backend!$CU$'+z+'="","","↓")))'; var z = z +1;
          values[a][3] ='=if(backend!$CU$'+z+'="✖","✖",if(backend!$CU$'+z+'=1,"↑",if(backend!$CU$'+z+'="","","↓")))'; var z = z +1;
          values[a][4] ='=if(backend!$CU$'+z+'="✖","✖",if(backend!$CU$'+z+'=1,"↗",if(backend!$CU$'+z+'="","","↙")))'; var z = z +1;
          break;
        case 1: 
          values[a][0] ='=if(backend!$CU$'+z+'="✖","✖",if(backend!$CU$'+z+'=1,"←",if(backend!$CU$'+z+'="","","→")))'; var z = z +1;
          values[a][1] = '=VLOOKUP(VLOOKUP("'+RC+'",backend!$CN$2:$CO$10,2,0),backend!$F$1:$G$10,2,0)'
          values[a][4] ='=if(backend!$CU$'+z+'="✖","✖",if(backend!$CU$'+z+'=1,"→",if(backend!$CU$'+z+'="","","←")))'; var z = z +1;
          break;
        case 2: 
          values[a][0] ='=if(backend!$CU$'+z+'="✖","✖",if(backend!$CU$'+z+'=1,"←",if(backend!$CU$'+z+'="","","→")))'; var z = z +1;
          values[a][4] ='=if(backend!$CU$'+z+'="✖","✖",if(backend!$CU$'+z+'=1,"→",if(backend!$CU$'+z+'="","","←")))'; var z = z +1;
          break;
        case 3: 
          values[a][0] ='=if(backend!$CU$'+z+'="✖","✖",if(backend!$CU$'+z+'=1,"←",if(backend!$CU$'+z+'="","","→")))'; var z = z +1;
          values[a][4] ='=if(backend!$CU$'+z+'="✖","✖",if(backend!$CU$'+z+'=1,"→",if(backend!$CU$'+z+'="","","←")))'; var z = z +1;
          break;
        case 4:
          values[a][0] ='=if(backend!$CU$'+z+'="✖","✖",if(backend!$CU$'+z+'=1,"↙",if(backend!$CU$'+z+'="","","↗")))'; var z = z +1;
          values[a][1] ='=if(backend!$CU$'+z+'="✖","✖",if(backend!$CU$'+z+'=1,"↓",if(backend!$CU$'+z+'="","","↑")))'; var z = z +1;
          values[a][2] ='=if(backend!$CU$'+z+'="✖","✖",if(backend!$CU$'+z+'=1,"↓",if(backend!$CU$'+z+'="","","↑")))'; var z = z +1;
          values[a][3] ='=if(backend!$CU$'+z+'="✖","✖",if(backend!$CU$'+z+'=1,"↓",if(backend!$CU$'+z+'="","","↑")))'; var z = z +1;
          values[a][4] ='=if(backend!$CU$'+z+'="✖","✖",if(backend!$CU$'+z+'=1,"↘",if(backend!$CU$'+z+'="","","↖")))'; var z = z +1;
          break;
      }
    }
    Sheet[x].getRange(r, c, 5, 5).setFormulas(values); 
  } 
  
  Sheet[x].getRange(22,1,9,22).setHorizontalAlignment('left');
  SpreadsheetApp.flush();
  //  }catch (error){
  //    var date = Utilities.formatDate(new Date(), Session.getTimeZone(), 'HH:MM');
  //    MailApp.sendEmail(Session.getActiveUser().getEmail(), "App Issue", "Hi ,"+ Session.getActiveUser() +", as you know the following error occured today at "+date+" on your app: "+ error +". I will try to fix it as soon as possible. Thanks. Frank");
  //    MailApp.sendEmail("fmontemorano@gmail.com", "App Issue fmonteomrano030-035", "App Issue", "Hi ,"+ Session.getActiveUser() +", as you know the following error occured today at "+date+" on your app: "+ error +". I will try to fix it as soon as possible. Thanks. Frank");
  //  }
}
//////////////////////////////////////////////////////////////////////////////ℱℳ
function fmontemorano032(){ 
  //  try{
  var x = 1
  var Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var Sheet = Spreadsheet.getSheets(); 
  
  var Sheet = Spreadsheet.getSheets(); 
  var row = 1; var column = 1; var numRows = Sheet[x].getRange("A:A").getLastRow(); var numColumns = Sheet[x].getRange("1:1").getLastColumn(); 
  var range = Sheet[x].getRange(row, column, numRows, numColumns);
  
  Sheet[x].setName('backend');
  range.clear().breakApart().setFontSize(10).setFontColor('#ff0000').setBackground('#000000').setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(false);
  if(numColumns>103){Sheet[x].deleteColumns(104, numColumns-103);}else{if(numColumns<103){Sheet[x].getRange(177,103,1,1).setValue("");}else{}};
  if(numRows>177){Sheet[x].deleteRows(178, numRows-177);}else{if(numRows<177){Sheet[x].getRange(177,103,1,1).setValue("");}else{}};
  
  var formulas = new Array(103); 
  var values = new Array(103);
  for (var y = 0; y < 1; y++) {
    values[y] = new Array(1);
    formulas[y] = new Array(1);
    for (var x = 0; x < 1; x++) {
      values[0] = '="Spreadsheet"';
      values[1] = '="SheetFE"';
      values[2] = '="SheetBE"';
      values[3] = '="☪"';
      values[4] = '="activate"';
      values[5] = '="invariable"';
      values[6] = '="userInput"';
      values[7] = '="Type"';
      values[8] = '="alfaID"';
      values[9] = '=""';
      values[10] = '="omegaID"';
      values[11] = '="alfaPar"';
      values[12] = '="omegaPar"';
      values[13] = '=if(indirect(ADDRESS(column()-12,7,1,true,$C$4))="x",69,value(right(indirect(ADDRESS(column()-12,6,1,true,$C$4)),LEN(indirect(ADDRESS(column()-12,6,1,true,$C$4)))-1)))';
      values[14] = '=if(indirect(ADDRESS(column()-12,7,1,true,$C$4))="x",69,value(right(indirect(ADDRESS(column()-12,6,1,true,$C$4)),LEN(indirect(ADDRESS(column()-12,6,1,true,$C$4)))-1)))';
      values[15] = '=if(indirect(ADDRESS(column()-12,7,1,true,$C$4))="x",69,value(right(indirect(ADDRESS(column()-12,6,1,true,$C$4)),LEN(indirect(ADDRESS(column()-12,6,1,true,$C$4)))-1)))';
      values[16] = '=if(indirect(ADDRESS(column()-12,7,1,true,$C$4))="x",69,value(right(indirect(ADDRESS(column()-12,6,1,true,$C$4)),LEN(indirect(ADDRESS(column()-12,6,1,true,$C$4)))-1)))';
      values[17] = '=if(indirect(ADDRESS(column()-12,7,1,true,$C$4))="x",69,value(right(indirect(ADDRESS(column()-12,6,1,true,$C$4)),LEN(indirect(ADDRESS(column()-12,6,1,true,$C$4)))-1)))';
      values[18] = '=if(indirect(ADDRESS(column()-12,7,1,true,$C$4))="x",69,value(right(indirect(ADDRESS(column()-12,6,1,true,$C$4)),LEN(indirect(ADDRESS(column()-12,6,1,true,$C$4)))-1)))';
      values[19] = '=if(indirect(ADDRESS(column()-12,7,1,true,$C$4))="x",69,value(right(indirect(ADDRESS(column()-12,6,1,true,$C$4)),LEN(indirect(ADDRESS(column()-12,6,1,true,$C$4)))-1)))';
      values[20] = '=if(indirect(ADDRESS(column()-12,7,1,true,$C$4))="x",69,value(right(indirect(ADDRESS(column()-12,6,1,true,$C$4)),LEN(indirect(ADDRESS(column()-12,6,1,true,$C$4)))-1)))';
      values[21] = '=if(indirect(ADDRESS(column()-12,7,1,true,$C$4))="x",69,value(right(indirect(ADDRESS(column()-12,6,1,true,$C$4)),LEN(indirect(ADDRESS(column()-12,6,1,true,$C$4)))-1)))';
      values[22] = '="Kpermit"';
      values[23] = '="MacroInput"';
      values[24] = '="Xpermit"';
      values[25] = '="in"';
      values[26] = '="out"';
      values[27] = '="☪"';
      values[28] = '="macroInput(1)"';
      values[29] = '="macroInput(2)"';
      values[30] = '=COUNTIF(AE2:AE,"<>")+2';
      values[31] = '="Concat"';
      values[32] = '="☪"';
      values[33] = '=QUERY(E:M,"select * where E = 1 and G != ""x""")';
      values[34] = '=""';
      values[35] = '=""';
      values[36] = '=""';
      values[37] = '=""';
      values[38] = '=""';
      values[39] = '=""';
      values[40] = '=""';
      values[41] = '=""';
      values[42] = '="☪"';
      values[43] = '=(COLUMN()-39)/5';
      values[44] = '=COUNTIF($AN$11:$AN,AR$1)';
      values[45] = '=COUNTIF($AL$11:$AL,AR$1)';
      values[46] = '="times"';
      values[47] = '="A1conc"';
      values[48] = '=(COLUMN()-39)/5';
      values[49] = '=COUNTIF($AN$11:$AN,AW$1)';
      values[50] = '=COUNTIF($AL$11:$AL,AW$1)';
      values[51] = '="times"';
      values[52] = '="A1conc"';
      values[53] = '=(COLUMN()-39)/5';
      values[54] = '=COUNTIF($AN$11:$AN,BB$1)';
      values[55] = '=COUNTIF($AL$11:$AL,BB$1)';
      values[56] = '="times"';
      values[57] = '="A1conc"';
      values[58] = '=(COLUMN()-39)/5';
      values[59] = '=COUNTIF($AN$11:$AN,BG$1)';
      values[60] = '=COUNTIF($AL$11:$AL,BG$1)';
      values[61] = '="times"';
      values[62] = '="A1conc"';
      values[63] = '=(COLUMN()-39)/5';
      values[64] = '=COUNTIF($AN$11:$AN,BL$1)';
      values[65] = '=COUNTIF($AL$11:$AL,BL$1)';
      values[66] = '="times"';
      values[67] = '="A1conc"';
      values[68] = '=(COLUMN()-39)/5';
      values[69] = '=COUNTIF($AN$11:$AN,BQ$1)';
      values[70] = '=COUNTIF($AL$11:$AL,BQ$1)';
      values[71] = '="times"';
      values[72] = '="A1conc"';
      values[73] = '=(COLUMN()-39)/5';
      values[74] = '=COUNTIF($AN$11:$AN,BV$1)';
      values[75] = '=COUNTIF($AL$11:$AL,BV$1)';
      values[76] = '="times"';
      values[77] = '="A1conc"';
      values[78] = '=(COLUMN()-39)/5';
      values[79] = '=COUNTIF($AN$11:$AN,CA$1)';
      values[80] = '=COUNTIF($AL$11:$AL,CA$1)';
      values[81] = '="times"';
      values[82] = '="A1conc"';
      values[83] = '=(COLUMN()-39)/5';
      values[84] = '=COUNTIF($AN$11:$AN,CF$1)';
      values[85] = '=COUNTIF($AL$11:$AL,CF$1)';
      values[86] = '="times"';
      values[87] = '="A1conc"';
      values[88] = '="☪"';
      values[89] = '="concat"';
      values[90] = '="☪"';
      values[91] = '="ref"';
      values[92] = '="invariable"';
      values[93] = '="in"';
      values[94] = '="out"';
      values[95] = '="☪"';
      values[96] = '="in"';
      values[97] = '="out"';
      values[98] = '="activation"';
      values[99] = '="ref"';
      values[100] = '="r"';
      values[101] = '="c"';
      values[102] = '="mod"';
      SpreadsheetApp.flush();
      formulas[0] = 'tH_LU-MwRMz3PXURcB7_EHg';
      formulas[1] = '1';
      formulas[2] = '2';
      formulas[3] = '="☪"';
      formulas[4] = '=IF(H2="A",if(sum(W2:Y2)=3,1,0),if(sum(W2:X2)=2,1,0))';
      formulas[5] = '=IFERROR(IF(I2="",IF(K2="","",CONCATENATE(H2,K2)),IF(K2="",CONCATENATE(H2,I2),CONCATENATE(H2,I2,$J$1,K2))),"")';
      formulas[6] = '=iferror(IF(H2="A",if(index(indirect(concatenate($B$4,"!$B3:$B11")),row()-1)="",F2,index(indirect(concatenate($B$4,"!$B3:$B11")),row()-1)),if(Vlookup(F2,indirect(concatenate($B$4,"!$X3:$Y")),2,0)="",F2,Vlookup(F2,indirect(concatenate($B$4,"!$X3:$Y")),2,0))),F2)';
      formulas[7] = '=IF(ROW()>10,"K","A")';
      formulas[8] = '=IF(ROW()<11,row()-1,if(row()=11,1,if(row()>91,if(row()>109,ROW()-109,""),if(MOD(row()-2,9)=0,I1+1,I1))))';
      formulas[9] = '';
      formulas[10] = '=IF(ROW()<11,"",if(row()=11,1,if(row()>91,if(row()>109,"",ROW()-91),if(MOD(row()-2,9)=0,1,K1+1))))';
      formulas[11] = '=iferror(if(indirect(ADDRESS(match($I2,$I$1:$I$17,0),7,1,true,$C$4))="x",indirect(ADDRESS(match($I2,$I$1:$I$17,0),6,1,true,$C$4)),indirect(ADDRESS(match($I2,$I$1:$I$17,0),7,1,true,$C$4))),"")';
      formulas[12] = '=iferror(if(indirect(ADDRESS(match($K2,$I$1:$I$17,0),7,1,true,$C$4))="x",indirect(ADDRESS(match($K2,$I$1:$I$17,0),6,1,true,$C$4)),indirect(ADDRESS(match($K2,$I$1:$I$17,0),7,1,true,$C$4))),"")';
      formulas[13] = '=if(OR($I2=N$1,$K2=N$1),1,0)';
      formulas[14] = '=if(OR($I2=O$1,$K2=O$1),1,0)';
      formulas[15] = '=if(OR($I2=P$1,$K2=P$1),1,0)';
      formulas[16] = '=if(OR($I2=Q$1,$K2=Q$1),1,0)';
      formulas[17] = '=if(OR($I2=R$1,$K2=R$1),1,0)';
      formulas[18] = '=if(OR($I2=S$1,$K2=S$1),1,0)';
      formulas[19] = '=if(OR($I2=T$1,$K2=T$1),1,0)';
      formulas[20] = '=if(OR($I2=U$1,$K2=U$1),1,0)';
      formulas[21] = '=if(OR($I2=V$1,$K2=V$1),1,0)';
      formulas[22] = '=IF(SUM(N2:V2)>0,1,0)';
      formulas[23] = '=if(H2="A",1,if(COUNTIF(AF:AF,F2)=0,0,1))';
      formulas[24] = '=IF(G2="x",0,1)';
      formulas[25] = '=if(H2="K",if(AND(E2=1,Y2=1),L2,""),"")';
      formulas[26] = '=if(H2="K",if(AND(E2=1,Y2=1),M2,""),"")';
      formulas[27] = '="☪"';
      formulas[28] = '=LEFT($AE2,1)';
      formulas[29] = '=RIGHT($AE2,1)';
      formulas[30] = '';
      formulas[31] = '=IFERROR(IF(AC2="",IF(AD2="","",CONCATENATE("K",AD2)),IF(AD2="",CONCATENATE("K",AC2),CONCATENATE("K",AC2,$J$1,AD2))),"")';
      formulas[32] = '="☪"';
      formulas[33] = '';
      formulas[34] = '';
      formulas[35] = '';
      formulas[36] = '';
      formulas[37] = '';
      formulas[38] = '';
      formulas[39] = '';
      formulas[40] = '';
      formulas[41] = '';
      formulas[42] = '="☪"';
      formulas[43] = '=IF(row()>2,"",IF(or(AS1>0,AT1>0),concatenate("DaDt(",AR$1,"): "),""))';
      formulas[44] = '=if(row()-2<AS$1," + ",if(AND(row()-2<AT$1+AS$1,row()-1>AS$1)," - ",""))';
      formulas[45] = '=iferror(if(row()-2<AS$1,index(QUERY($AI$11:$AP,concatenate("select AJ where AN = ",AR$1)),row()-1),if(AND(row()-2<AT$1+AS$1,row()-1>AS$1),index(QUERY($AI$11:$AP,concatenate("select AI where AL = ",AR$1)),row()-1-AS$1),"")),"")';
      formulas[46] = '=IF(AS2="",""," * ")';
      formulas[47] = '=IF(AS2="","",concatenate("A",mid(AT2,2,1)))';
      formulas[48] = '=IF(row()>2,"",IF(or(AX1>0,AY1>0),concatenate("DaDt(",AW$1,"): "),""))';
      formulas[49] = '=if(row()-2<AX$1," + ",if(AND(row()-2<AY$1+AX$1,row()-1>AX$1)," - ",""))';
      formulas[50] = '=iferror(if(row()-2<AX$1,index(QUERY($AI$11:$AP,concatenate("select AJ where AN = ",AW$1)),row()-1),if(AND(row()-2<AY$1+AX$1,row()-1>AX$1),index(QUERY($AI$11:$AP,concatenate("select AI where AL = ",AW$1)),row()-1-AX$1),"")),"")';
      formulas[51] = '=IF(AX2="",""," * ")';
      formulas[52] = '=IF(AX2="","",concatenate("A",mid(AY2,2,1)))';
      formulas[53] = '=IF(row()>2,"",IF(or(BC1>0,BD1>0),concatenate("DaDt(",BB$1,"): "),""))';
      formulas[54] = '=if(row()-2<BC$1," + ",if(AND(row()-2<BD$1+BC$1,row()-1>BC$1)," - ",""))';
      formulas[55] = '=iferror(if(row()-2<BC$1,index(QUERY($AI$11:$AP,concatenate("select AJ where AN = ",BB$1)),row()-1),if(AND(row()-2<BD$1+BC$1,row()-1>BC$1),index(QUERY($AI$11:$AP,concatenate("select AI where AL = ",BB$1)),row()-1-BC$1),"")),"")';
      formulas[56] = '=IF(BC2="",""," * ")';
      formulas[57] = '=IF(BC2="","",concatenate("A",mid(BD2,2,1)))';
      formulas[58] = '=IF(row()>2,"",IF(or(BH1>0,BI1>0),concatenate("DaDt(",BG$1,"): "),""))';
      formulas[59] = '=if(row()-2<BH$1," + ",if(AND(row()-2<BI$1+BH$1,row()-1>BH$1)," - ",""))';
      formulas[60] = '=iferror(if(row()-2<BH$1,index(QUERY($AI$11:$AP,concatenate("select AJ where AN = ",BG$1)),row()-1),if(AND(row()-2<BI$1+BH$1,row()-1>BH$1),index(QUERY($AI$11:$AP,concatenate("select AI where AL = ",BG$1)),row()-1-BH$1),"")),"")';
      formulas[61] = '=IF(BH2="",""," * ")';
      formulas[62] = '=IF(BH2="","",concatenate("A",mid(BI2,2,1)))';
      formulas[63] = '=IF(row()>2,"",IF(or(BM1>0,BN1>0),concatenate("DaDt(",BL$1,"): "),""))';
      formulas[64] = '=if(row()-2<BM$1," + ",if(AND(row()-2<BN$1+BM$1,row()-1>BM$1)," - ",""))';
      formulas[65] = '=iferror(if(row()-2<BM$1,index(QUERY($AI$11:$AP,concatenate("select AJ where AN = ",BL$1)),row()-1),if(AND(row()-2<BN$1+BM$1,row()-1>BM$1),index(QUERY($AI$11:$AP,concatenate("select AI where AL = ",BL$1)),row()-1-BM$1),"")),"")';
      formulas[66] = '=IF(BM2="",""," * ")';
      formulas[67] = '=IF(BM2="","",concatenate("A",mid(BN2,2,1)))';
      formulas[68] = '=IF(row()>2,"",IF(or(BR1>0,BS1>0),concatenate("DaDt(",BQ$1,"): "),""))';
      formulas[69] = '=if(row()-2<BR$1," + ",if(AND(row()-2<BS$1+BR$1,row()-1>BR$1)," - ",""))';
      formulas[70] = '=iferror(if(row()-2<BR$1,index(QUERY($AI$11:$AP,concatenate("select AJ where AN = ",BQ$1)),row()-1),if(AND(row()-2<BS$1+BR$1,row()-1>BR$1),index(QUERY($AI$11:$AP,concatenate("select AI where AL = ",BQ$1)),row()-1-BR$1),"")),"")';
      formulas[71] = '=IF(BR2="",""," * ")';
      formulas[72] = '=IF(BR2="","",concatenate("A",mid(BS2,2,1)))';
      formulas[73] = '=IF(row()>2,"",IF(or(BW1>0,BX1>0),concatenate("DaDt(",BV$1,"): "),""))';
      formulas[74] = '=if(row()-2<BW$1," + ",if(AND(row()-2<BX$1+BW$1,row()-1>BW$1)," - ",""))';
      formulas[75] = '=iferror(if(row()-2<BW$1,index(QUERY($AI$11:$AP,concatenate("select AJ where AN = ",BV$1)),row()-1),if(AND(row()-2<BX$1+BW$1,row()-1>BW$1),index(QUERY($AI$11:$AP,concatenate("select AI where AL = ",BV$1)),row()-1-BW$1),"")),"")';
      formulas[76] = '=IF(BW2="",""," * ")';
      formulas[77] = '=IF(BW2="","",concatenate("A",mid(BX2,2,1)))';
      formulas[78] = '=IF(row()>2,"",IF(or(CB1>0,CC1>0),concatenate("DaDt(",CA$1,"): "),""))';
      formulas[79] = '=if(row()-2<CB$1," + ",if(AND(row()-2<CC$1+CB$1,row()-1>CB$1)," - ",""))';
      formulas[80] = '=iferror(if(row()-2<CB$1,index(QUERY($AI$11:$AP,concatenate("select AJ where AN = ",CA$1)),row()-1),if(AND(row()-2<CC$1+CB$1,row()-1>CB$1),index(QUERY($AI$11:$AP,concatenate("select AI where AL = ",CA$1)),row()-1-CB$1),"")),"")';
      formulas[81] = '=IF(CB2="",""," * ")';
      formulas[82] = '=IF(CB2="","",concatenate("A",mid(CC2,2,1)))';
      formulas[83] = '=IF(row()>2,"",IF(or(CG1>0,CH1>0),concatenate("DaDt(",CF$1,"): "),""))';
      formulas[84] = '=if(row()-2<CG$1," + ",if(AND(row()-2<CH$1+CG$1,row()-1>CG$1)," - ",""))';
      formulas[85] = '=iferror(if(row()-2<CG$1,index(QUERY($AI$11:$AP,concatenate("select AJ where AN = ",CF$1)),row()-1),if(AND(row()-2<CH$1+CG$1,row()-1>CG$1),index(QUERY($AI$11:$AP,concatenate("select AI where AL = ",CF$1)),row()-1-CG$1),"")),"")';
      formulas[86] = '=IF(CG2="",""," * ")';
      formulas[87] = '=IF(CG2="","",concatenate("A",mid(CH2,2,1)))';
      formulas[88] = '="☪"';
      formulas[89] = '=trim(CONCATENATE(indirect(address(2,44+(5*(row()-2)),1,true,$C$4)):indirect(address(countblank(A:A)+countif(A:A,"<>"),43+(5*(row()-1)),1,true,$C$4))))';
      formulas[90] = '="☪"';
      formulas[91] = '=transpose({"E3","K3","Q3","E9","K9","Q9","E15","K15","Q15"})';
      formulas[92] = '=transpose({"A1","A2","A3","A4","A5","A6","A7","A8","A9"})';
      formulas[93] = '=COUNTIF(AA:AA,$CO2)';
      formulas[94] = '=COUNTIF(Z:Z,$CO2)';
      formulas[95] = '="☪"';
      formulas[96] = '=if(MOD(row()-2,16)=0,VLOOKUP($CV2,$CN:$CQ,3,0),if(CS1-1<0,0,CS1-1))';
      formulas[97] = '=if(MOD(row()-1,16)=0,VLOOKUP($CV2,$CN:$CQ,4,0),if(CT3-1<0,0,CT3-1))';
      formulas[98] = '=if(CS2>0,if(CT2>0,"✖",0),if(CT2>0,1,""))';
      formulas[99] = '=if(ROW()<18,"E3",if(ROW()<34,"K3",if(ROW()<50,"Q3",if(ROW()<66,"E9",if(ROW()<98,"K9",if(ROW()<114,"Q9",if(ROW()<146,"E15",if(ROW()<162,"K15",if(ROW()<178,"Q15")))))))))';
      formulas[100] = '=INDEX(transpose({0,0,0,1,2,3,4,4,4,4,4,3,2,1,0,0}),CY2)';
      formulas[101] = '=INDEX(transpose({2,3,4,4,4,4,4,3,2,1,0,0,0,0,0,1}),CY2)';
      formulas[102] = '=if(MOD(row()-2,16)=0,1,CY1+1)';
    }
  }
  Sheet[x].getRange(row, column, 1, numColumns).setFormulas([values]);
  Sheet[x].getRange(row+1, column, 1, numColumns).setFormulas([formulas]);
  Sheet[x].getRange("A2:C2").copyTo(Sheet[x].getRange("A3:C4"));
  Sheet[x].getRange("D2:CK2").copyTo(Sheet[x].getRange("D3:CK118"));
  Sheet[x].getRange("CL2:CQ2").copyTo(Sheet[x].getRange("CL3:CQ10"));
  Sheet[x].getRange("CR2:CY2").copyTo(Sheet[x].getRange("CR3:CY177"));
  
  Sheet[x].getRange(2,1).setValue(Spreadsheet.getId());
  Sheet[x].getRange(3,1).setValue('https://docs.google.com/spreadsheet/ccc?key='+Spreadsheet.getId());
  Sheet[x].getRange(4,1).setValue(Spreadsheet.getName());
  Sheet[x].getRange(2,2).setValue(Sheet[0].getSheetId());
  Sheet[x].getRange(3,2).setValue(0);
  Sheet[x].getRange(4,2).setValue('frontend');
  Sheet[x].getRange(2,3).setValue(Sheet[1].getSheetId());
  Sheet[x].getRange(3,3).setValue(1);
  Sheet[x].getRange(4,3).setValue('backend');
  //  }catch (error){
  //    var date = Utilities.formatDate(new Date(), Session.getTimeZone(), 'HH:MM');
  //    MailApp.sendEmail(Session.getActiveUser().getEmail(), "App Issue", "Hi ,"+ Session.getActiveUser() +", as you know the following error occured today at "+date+" on your app: "+ error +". I will try to fix it as soon as possible. Thanks. Frank");
  //    MailApp.sendEmail("fmontemorano@gmail.com", "App Issue fmonteomrano030-035", "App Issue", "Hi ,"+ Session.getActiveUser() +", as you know the following error occured today at "+date+" on your app: "+ error +". I will try to fix it as soon as possible. Thanks. Frank");
  //  }
}

//////////////////////////////////////////////////////////////////////////////ℱℳ
function fmontemorano033(){
  var x = 0
  var Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var Sheet = Spreadsheet.getSheets(); 
  
  var Sheet = Spreadsheet.getSheets(); 
  var row = 1; var column = 1; var numRows = Sheet[x].getRange("A:A").getLastRow(); var numColumns = Sheet[x].getRange("1:1").getLastColumn(); 
  var range = Sheet[x].getRange(row, column, numRows, numColumns);
  
  Sheet[x].getRange("A22").setFormula('=concatenate(backend!CL2,char(10),backend!CL3,char(10),backend!CL4,char(10),backend!CL5,char(10),backend!CL6,char(10),backend!CL7,char(10),backend!CL8,char(10),backend!CL9,char(10),backend!CL10)');
  
  try{
    var parent = DocsList.getFolder('Frank☪odes'); 
  }catch(err){
    var parent = DocsList.createFolder('Frank☪odes');
  }
  
  var docName = Utilities.formatDate(new Date(), Session.getTimeZone(), 'yyyy MMM dd | HH:mm');
  var doc = DocumentApp.create(docName);
  var docOnDrive = DocsList.getFileById(doc.getId());
  var body = doc.getBody();
  
  docOnDrive.addToFolder(parent);
  docOnDrive.removeFromFolder(DocsList.getRootFolder());
  
  var paragraph = body.appendParagraph(Sheet[x].getRange("A22").getValue()); 
  //  var style = {}; 
  //  style[DocumentApp.Attribute.SPACING_BEFORE] = 0;
  //  paragraph.setAttributes(style);
  
  Sheet[x].getRange("A22").setFormula('=backend!CL2');
  Sheet[x].getRange("A23").setFormula('=backend!CL3');
  Sheet[x].getRange("A24").setFormula('=backend!CL4');
  Sheet[x].getRange("A25").setFormula('=backend!CL5');
  Sheet[x].getRange("A26").setFormula('=backend!CL6');
  Sheet[x].getRange("A27").setFormula('=backend!CL7');
  Sheet[x].getRange("A28").setFormula('=backend!CL8');
  Sheet[x].getRange("A29").setFormula('=backend!CL9');
  Sheet[x].getRange("A30").setFormula('=backend!CL10');  
  
  doc.saveAndClose();
  var pdf = DocsList.createFile(doc.getAs('application/pdf'))
  pdf.rename(docName); 
  var pdfId = pdf.getId()
  var docOnDrive = DocsList.getFileById(pdf.getId());
  docOnDrive.addToFolder(parent);
  docOnDrive.removeFromFolder(DocsList.getRootFolder());
  
  var href = 'https://drive.google.com/uc?export=download&id='+pdfId
  var text = '<h2>Download your ☪ode</h2>';
  var app = UiApp.createApplication().setHeight(100).setWidth(300);
  app.setTitle('ⓕⓜⓞⓝⓣⓔⓜⓞⓡⓐⓝⓞ™')
  var widget = app.createAnchor(text, true, href);
  app.add(app.createVerticalPanel().add(widget));
  var doc = SpreadsheetApp.getActive();
  doc.show(app); 
  
}

function fmontemorano034(){
  var Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  Sheet[1].getRange("AE2:AE").clearContent();
  Sheet[0].getRange("A3:B11").clearContent();
  Sheet[0].getRange("X3:Y").clearContent();
  Sheet[0].getRange("X3").setFormula('=IFERROR(query(backend!E11:F,"select F where E = 1"),"")');
  Sheet[0].getRange("A3").setFormula('=TRANSPOSE({"A1","A2","A3","A4","A5","A6","A7","A8","A9"})');
}

//////////////////////////////////////////////////////////////////////////////ℱℳ
function onEdit(){
  var k = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getRange("B13").getValue();
  SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getRange("B13").clearContent();
  SpreadsheetApp.flush();
  var Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  Sheet[1].getRange(Sheet[1].getRange(1,31).getValue(),31).setValue(k)
  
  if(Sheet[0].getRange("B3").getValue()=='x'){Sheet[0].getRange("E3").offset(1,1,3,3).clearContent().breakApart();}else{Sheet[0].getRange("E3").offset(1,1,3,3).merge().setFormula('=VLOOKUP(VLOOKUP("E3",backend!$CN$2:$CO$10,2,0),backend!$F$1:$G$17,2,0)')}
  if(Sheet[0].getRange("B4").getValue()=='x'){Sheet[0].getRange("K3").offset(1,1,3,3).clearContent().breakApart();}else{Sheet[0].getRange("K3").offset(1,1,3,3).merge().setFormula('=VLOOKUP(VLOOKUP("K3",backend!$CN$2:$CO$10,2,0),backend!$F$1:$G$17,2,0)')}
  if(Sheet[0].getRange("B5").getValue()=='x'){Sheet[0].getRange("Q3").offset(1,1,3,3).clearContent().breakApart();}else{Sheet[0].getRange("Q3").offset(1,1,3,3).merge().setFormula('=VLOOKUP(VLOOKUP("Q3",backend!$CN$2:$CO$10,2,0),backend!$F$1:$G$17,2,0)')}
  if(Sheet[0].getRange("B6").getValue()=='x'){Sheet[0].getRange("E9").offset(1,1,3,3).clearContent().breakApart();}else{Sheet[0].getRange("E9").offset(1,1,3,3).merge().setFormula('=VLOOKUP(VLOOKUP("E9",backend!$CN$2:$CO$10,2,0),backend!$F$1:$G$17,2,0)')}
  if(Sheet[0].getRange("B7").getValue()=='x'){Sheet[0].getRange("K9").offset(1,1,3,3).clearContent().breakApart();}else{Sheet[0].getRange("K9").offset(1,1,3,3).merge().setFormula('=VLOOKUP(VLOOKUP("K9",backend!$CN$2:$CO$10,2,0),backend!$F$1:$G$17,2,0)')}
  if(Sheet[0].getRange("B8").getValue()=='x'){Sheet[0].getRange("Q9").offset(1,1,3,3).clearContent().breakApart();}else{Sheet[0].getRange("Q9").offset(1,1,3,3).merge().setFormula('=VLOOKUP(VLOOKUP("Q9",backend!$CN$2:$CO$10,2,0),backend!$F$1:$G$17,2,0)')}
  if(Sheet[0].getRange("B9").getValue()=='x'){Sheet[0].getRange("E15").offset(1,1,3,3).clearContent().breakApart();}else{Sheet[0].getRange("E15").offset(1,1,3,3).merge().setFormula('=VLOOKUP(VLOOKUP("E15",backend!$CN$2:$CO$10,2,0),backend!$F$1:$G$17,2,0)')}
  if(Sheet[0].getRange("B10").getValue()=='x'){Sheet[0].getRange("K15").offset(1,1,3,3).clearContent().breakApart();}else{Sheet[0].getRange("K15").offset(1,1,3,3).merge().setFormula('=VLOOKUP(VLOOKUP("K15",backend!$CN$2:$CO$10,2,0),backend!$F$1:$G$17,2,0)')}
  if(Sheet[0].getRange("B11").getValue()=='x'){Sheet[0].getRange("Q15").offset(1,1,3,3).clearContent().breakApart();}else{Sheet[0].getRange("Q15").offset(1,1,3,3).merge().setFormula('=VLOOKUP(VLOOKUP("Q15",backend!$CN$2:$CO$10,2,0),backend!$F$1:$G$17,2,0)')}
  
  var tm = '=Image("https://dl.dropboxusercontent.com/u/94960692/tiles/projects/2013%2008/rightsReserved.png",3)';
  if(Sheet[0].getRange("A14")!=tm){
    Sheet[0].getRange("A14").setFormula(tm);
  }
  
  Sheet[0].getRange(13,2).activate();
}

function triggerMidnight(){
  //set current triggers
  fmontemorano033();
  fmontemorano032();
  fmontemorano031();
}

function onOpen(){
  fmontemorano021();
  var Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  while(Spreadsheet.getSheets().length<2){Spreadsheet.insertSheet()};
}

function fmontemorano021(){ 

  
  var Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var menu = [
    {name: "clear data", functionName: "fmontemorano034"},
    {name: "create file", functionName: "fmontemorano033"},
    {name: "fix frontend", functionName: "fmontemorano031"},
    {name: "fix backend", functionName: "fmontemorano032"},
  ];
    Spreadsheet.addMenu("Prototype Actions", menu);
    }