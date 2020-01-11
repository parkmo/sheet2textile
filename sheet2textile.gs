var szOutString = "";

function TextileNormal()
{
  return Textile(0);
}

function TextileSmall()
{
  return Textile(1);
}

function TextileNormalNoneMerged()
{
  return TextileProcByArr(0);
}

function TextileSmallNoneMerged()
{
  return TextileProcByArr(1);
}

function getInMergeRangeByArr(activeRange, iPosX, iPosY)
{
  var szRetValue="";
  mergedRanges=activeRange.getMergedRanges();

  for (var i = 0; i < mergedRanges.length; i++) {
//    console.log("MergeDebug:[" + i + "] iPosX [" + iPosX + "] iPosY [" + iPosY + "] RowIndex[" + mergedRanges[i].getRowIndex()
//    + "] Row[" + mergedRanges[i].getRow() + "] Col[" + mergedRanges[i].getColumn() + "]");
    curCell = activeRange.getCell(iPosX+1,iPosY+1);
    szSplited=mergedRanges[i].getA1Notation().split(":");

    if ( szSplited[0] == curCell.getA1Notation() ) { // Start Pos
      if ( mergedRanges[i].getNumColumns() > 1) {
        szRetValue += "\\" + mergedRanges[i].getNumColumns();
      }
      if ( mergedRanges[i].getNumRows() > 1 ) {
        szRetValue += "/" + mergedRanges[i].getNumRows();
      }
    }
//    Logger.log(mergedRanges[i].getDisplayValue());
//    Logger.log("xxCols[" + mergedRanges[i].getNumColumns() + "] Rows[" + mergedRanges[i].getNumRows() + "]");
  }
  return szRetValue;
}

function ConvTextStyleByArr(szDisplayValue, szBgColor, szFontColor, szFontStyle, szFontWeight, szFontLine, szHorizontalAlignment, szVerticalAlignment)
{
  var szOutString = ""
  var szSetFontColor = "{";
  var szFontBold = "";
  var szFontLine = "";
  var szFontItalic = "";
  var szHorizonAlign = "";
  var szVerticalAlign = "";
  // Background Color and FontColor
  if ( szBgColor == "#ffffff" ) {
    szBgColor = "";
  }
  if ( (szBgColor.length) > 0 ) {
    szSetFontColor += "background:" + szBgColor + ";"
  }
  if ( szFontColor == "#000000" ) {
    szFontColor = "";
  }
  if ( (szFontColor.length) > 0 ) {
    szSetFontColor += "color:" + szFontColor + ";"
  }
  
  if ( szFontStyle == "italic" ) {
    szFontItalic = "_";
  }
  
  if ( szFontWeight == "bold" ) {
    szFontBold = "*";
  }
  if ( szFontLine == "line-through" ) {
    szFontLine = "-";
  }
  else if ( szFontLine == "underline" ) {
    szFontLine = "+";
  }
  if ( szHorizontalAlignment == "general-left" || szHorizontalAlignment == "left") {
    szHorizonAlign = "<";
  }
  else if ( szHorizontalAlignment == "general-right" || szHorizontalAlignment == "right") {
    szHorizonAlign = ">";
  } else if ( szHorizontalAlignment == "center" ) {
    szHorizonAlign = "=";
  }
  //szHorizonAlign="";
  if ( szVerticalAlignment == "top" ) {
    szVerticalAlign = "^";
  }
  else if ( szVerticalAlignment == "bottom" ) {
    szVerticalAlign = "~";
  }

  if ( (szSetFontColor.length) > 5 ) {
     szSetFontColor += "}";
  }
  else {
    szSetFontColor = "";
  }

  szOutString += szHorizonAlign + szVerticalAlign + szSetFontColor + ". " + szFontBold + szFontItalic + szFontLine + szDisplayValue + szFontLine + szFontItalic + szFontBold;
  return szOutString;
}

function TextileProcByArr(bSmall)
{
  var r1, r2, ro1, co1;
  var First, str1, strElm, strText;
  var selection = SpreadsheetApp.getActiveSpreadsheet().getSelection();
  var activeRange = selection.getActiveRange();
  var ui = SpreadsheetApp.getUi();
  var szTitle;
  
  if ( bSmall ) {
    szOutString += "table{valign:top;font-size:small}." + "\r\n";
    szTitle = "TextTile Table(S)";
  }
  else {
    szTitle = "TextTile Table";
  }
  
  var ar_szBackgrounds = activeRange.getBackgrounds();
  var ar_szFontColors = activeRange.getFontColors();
  var ar_szFontStyles = activeRange.getFontStyles();
  var ar_szFontWeights = activeRange.getFontWeights();
  var ar_DisplayValues = activeRange.getDisplayValues();
  var ar_szFontLines = activeRange.getFontLines();
  var ar_szHorizontalAlignment = activeRange.getHorizontalAlignments();
  var ar_szVerticalAlignment = activeRange.getVerticalAlignments();
  for (var r=0; r < ar_DisplayValues.length; r++) {
    for (var c=0; c < ar_DisplayValues[r].length; c++) {
        curRet= getInMergeRangeByArr(activeRange,r,c);
        szOutString += "|";
        szOutString += curRet;
        szOutString += ConvTextStyleByArr(ar_DisplayValues[r][c], ar_szBackgrounds[r][c],
                         ar_szFontColors[r][c],ar_szFontStyles[r][c],ar_szFontWeights[r][c]
                                         , ar_szFontLines[r][c], ar_szHorizontalAlignment[r][c], ar_szVerticalAlignment[r][c]);
    }
    szOutString += '|\r\n';
    console.log("Textile Line:[" + r + "]");
  }
  var htmlOutput=HtmlService.createHtmlOutput("<pre>" + szOutString + "</pre>").setTitle(szTitle);
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function getInMergeRange(curMergedRanges, curCell)
{
  var szRetValue="|";
  var szCellName = curCell.getA1Notation();
//  console.log( "MergedRanges.length[%d] szCellName[%s]",
//              curMergedRanges.length, szCellName);
  for (var i = 0; i < curMergedRanges.length; i++) {
    szSplited=curMergedRanges[i].getA1Notation().split(":");
    if ( szSplited[0] == szCellName ) { // Start Pos
//    console.log( "Same TryRange[%d] szSplited[0][%s] C[%d] R[%d]", i
//                , szSplited[0]
//               ,curMergedRanges[i].getNumColumns(), curMergedRanges[i].getNumRows());
      if ( curMergedRanges[i].getNumColumns() > 1) {
        szRetValue += "\\" + curMergedRanges[i].getNumColumns();
      }
      if ( curMergedRanges[i].getNumRows() > 1 ) {
        szRetValue += "/" + curMergedRanges[i].getNumRows();
      }
       szRetValue += ConvTextStyle(curCell);
       return szRetValue;
    }
//    console.log( "TryRange[%d] szSplited[0][%s] C[%d] R[%d]", i
//                , szSplited[0]
//               ,curMergedRanges[i].getNumColumns(), curMergedRanges[i].getNumRows());
  }
  if ( curCell.isPartOfMerge() ) {
    return "";
  }
  szRetValue += ConvTextStyle(curCell);
  return szRetValue;
}

function Textile(bSmall)
{
  var r1, r2, ro1, co1;
  var First, str1, strElm, strText;
  var CB; // DataObject
  var selection = SpreadsheetApp.getActiveSpreadsheet().getSelection();
  var activeRange = selection.getActiveRange();
  var ui = SpreadsheetApp.getUi();
  var szTitle;
  
  ro1 = activeRange.getNumRows();
  co1 = activeRange.getNumColumns();
  mergedRanges=activeRange.getMergedRanges();
  if ( bSmall ) {
    szOutString += "table{valign:top;font-size:small}." + "\r\n";
    szTitle = "TextTile Table(S)";
  }
  else {
    szTitle = "TextTile Table";
  }
  
  for (var i=1; i <= ro1; i++) {
	for (var j=1; j <= co1; j++ ) {      
      curCell = activeRange.getCell(i, j);
      curRet=getInMergeRange(mergedRanges, curCell);
      szOutString += curRet;
	}
    szOutString += '|\r\n';
    console.log("Textile Line:[" + i + "]");
  }
  console.log("Textile :" + szOutString);
  var htmlOutput=HtmlService.createHtmlOutput("<pre>" + szOutString + "</pre>").setTitle(szTitle);
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function onInstall(e) {
  onOpen(e);
  // Perform additional setup as needed.
}

function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu() // Or DocumentApp.
      .addItem('Make textile', 'TextileNormal')
      .addItem('Make textile(S)', 'TextileSmall')
      .addItem('Make textile_NM', 'TextileNormalNoneMerged')
      .addItem('Make textile(S)_NM', 'TextileSmallNoneMerged')
      .addToUi();
}

function ConvTextStyle(curCell)
{
  var szHorizontalAlignment = curCell.getHorizontalAlignment();
  var szVerticalAlignment = curCell.getVerticalAlignment();
  var szOutString = ""
  var szSetFontColor = "{";
  var szFontBold = "";
  var szFontLine = "";
  var szFontItalic = "";
  var szHorizonAlign = "";
  var szVerticalAlign = "";
  // Background Color and FontColor
  szBgColor = curCell.getBackground();
  szFontColor = curCell.getFontColor();
  if ( szBgColor == "#ffffff" ) {
    szBgColor = "";
  }
  if ( (szBgColor.length) > 0 ) {
    szSetFontColor += "background:" + szBgColor + ";"
  }
  if ( szFontColor == "#000000" ) {
    szFontColor = "";
  }
  if ( (szFontColor.length) > 0 ) {
    szSetFontColor += "color:" + szFontColor + ";"
  }
  
  if ( curCell.getFontStyle() == "italic" ) {
    szFontItalic = "_";
  }
  
  if ( curCell.getFontWeight() == "bold" ) {
    szFontBold = "*";
  }
  if ( curCell.getFontLine() == "line-through" ) {
    szFontLine = "-";
  }
  else if ( curCell.getFontLine() == "underline" ) {
    szFontLine = "+";
  }
  if ( szHorizontalAlignment == "general-left" || szHorizontalAlignment == "left") {
    szHorizonAlign = "<";
  }
  else if ( szHorizontalAlignment == "general-right" || szHorizontalAlignment == "right") {
    szHorizonAlign = ">";
  } else if ( szHorizontalAlignment == "center" ) {
    szHorizonAlign = "=";
  }
  //szHorizonAlign="";
  if ( szVerticalAlignment == "top" ) {
    szVerticalAlign = "^";
  }
  else if ( szVerticalAlignment == "bottom" ) {
    szVerticalAlign = "~";
  }

  if ( (szSetFontColor.length) > 5 ) {
     szSetFontColor += "}";
  }
  else {
    szSetFontColor = "";
  }

  szOutString += szHorizonAlign + szVerticalAlign + szSetFontColor + ". " + szFontBold + szFontItalic + szFontLine + curCell.getDisplayValue() + szFontLine + szFontItalic + szFontBold;
  return szOutString;
}
