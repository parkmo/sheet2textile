var globalOptWithLink = false;

function TextileNormalNoneMerged()
{
  return TextileProcByArr(0);
}

function TextileSmallNoneMerged()
{
  return TextileProcByArr(1);
}

function TextileSmallNoneMergedWithLink()
{
  globalOptWithLink = true;
  return TextileProcByArr(1);
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

  szDisplayValue =  szDisplayValue.replace(/\n\n/g, '\n&amp;nbsp;\n');  
  szOutString += szHorizonAlign + szVerticalAlign + szSetFontColor + ". " + szFontBold + szFontItalic + szFontLine + szDisplayValue + szFontLine + szFontItalic + szFontBold;
  return szOutString;
}

function createMergedInfo(curRange)
{
  var szRetValue="";
  var ar_MergedInfo = curRange.getFormulas(); // make empty string array
  var mergedRanges=curRange.getMergedRanges();
  var iBaseRow = curRange.getRow();
  var iBaseCol = curRange.getColumn();
//  console.log("curRange[%d][%d]", curRange.getRow(), curRange.getColumn());
  for (var i = 0; i < mergedRanges.length; i++) {
//    console.log("RangeNot[%s] Row[%d] NumRow[%d] Col[%d] NumCol[%d]"
//                , mergedRanges[i].getA1Notation()
//                , mergedRanges[i].getRow()
//                , mergedRanges[i].getNumRows()
//                , mergedRanges[i].getColumn()
//                , mergedRanges[i].getNumColumns()
//               );
    var iCurRow = mergedRanges[i].getRow();
    var iCurCol = mergedRanges[i].getColumn();
    var iNumRows = mergedRanges[i].getNumRows();
    var iNumCols = mergedRanges[i].getNumColumns();
    iCurRow -= iBaseRow;
    iCurCol -= iBaseCol;
    ar_MergedInfo[iCurRow][iCurCol] = Utilities.formatString("%d:%d", iNumRows, iNumCols);
    for (var iRow = 0 ; iRow < iNumRows ; iRow++ ) {
      for (var iCol = 0 ; iCol < iNumCols; iCol++ ) {
        if ( !(iRow == 0 && iCol == 0) ) {
          ar_MergedInfo[iCurRow+iRow][iCurCol+iCol] = "1:1"; // Mark Merged
        }
      }
    }
  }
//  console.log("Log : ar_MergedInfo");
//  console.log(ar_MergedInfo);
  return ar_MergedInfo;
}

function getInMergeRangeByArr(ar_MergedInfo, iPosX, iPosY)
{
  var szRetValue="";
  szOrginValue = ar_MergedInfo[iPosX][iPosY];
  szSplited = szOrginValue.split(":");
  if ( szOrginValue == "" ) { // Normal Block
    return "N"; // Normal Block
  }
  else if ( szOrginValue == "1:1" ) { // Merged Block
    return "M"; // Merged Block
  }
  else {
    if ( +szSplited[0] > 1 ) {
      szRetValue += "/" + szSplited[0];
    }
    if ( +szSplited[1] > 1 ) {
      szRetValue += "\\" + szSplited[1];
    }
  }
  return szRetValue;
}

function getCellValueWithLink(cellValue, cellRichValue)
{
  var curLinkURL = null;
  if (globalOptWithLink) {
    curLinkURL = cellRichValue.getLinkUrl();
  }
  var retValue = cellValue;
  if ( curLinkURL ) {
    retValue = `"${cellValue}":${curLinkURL}`
  }
  // Logger.log(`cellValue: , [${cellValue}]`);
  // Logger.log(`curLinkURL: [${curLinkURL}]`);
  return retValue;
}

function TextileProcByArr(bSmall)
{
  var selection = SpreadsheetApp.getActiveSpreadsheet().getSelection();
  var activeRange = selection.getActiveRange();
  var szTitle;
  var szOutString = "";
  
  if ( bSmall ) {
    szOutString += "table{valign:top;font-size:small}." + "\r\n";
    szTitle = "TextTile Table(S)";
  }
  else {
    szTitle = "TextTile Table";
  }
  // Logger.log("Start TextileProcByArr");
  
  var ar_szBackgrounds = activeRange.getBackgrounds();
  var ar_szFontColors = activeRange.getFontColors();
  var ar_szFontStyles = activeRange.getFontStyles();
  var ar_szFontWeights = activeRange.getFontWeights();
  var ar_DisplayValues = activeRange.getDisplayValues();
  var ar_RichTextValues = activeRange.getRichTextValues();

  var ar_szFontLines = activeRange.getFontLines();
  var ar_szHorizontalAlignment = activeRange.getHorizontalAlignments();
  var ar_szVerticalAlignment = activeRange.getVerticalAlignments();
  var ar_megedInfo = createMergedInfo(activeRange);

  for (var r=0; r < ar_DisplayValues.length; r++) {
    for (var c=0; c < ar_DisplayValues[r].length; c++) {
        szCurRet= getInMergeRangeByArr(ar_megedInfo,r,c);
      if ( szCurRet != "M" ) { // N or 2:1 ...
        szOutString += "|";
        if ( szCurRet == "N" ) {
          szCurRet = "";
        }
        szOutString += szCurRet;
        szOutString += ConvTextStyleByArr(getCellValueWithLink(ar_DisplayValues[r][c], ar_RichTextValues[r][c]), ar_szBackgrounds[r][c], ar_szFontColors[r][c]
                                          , ar_szFontStyles[r][c],ar_szFontWeights[r][c]
                                          , ar_szFontLines[r][c], ar_szHorizontalAlignment[r][c], ar_szVerticalAlignment[r][c]
                                         );
      }
    }
    szOutString += '|\r\n';
//    console.log("Textile Line:[" + r + "]");
  }
  var htmlOutput=HtmlService.createHtmlOutput("<pre>" + szOutString + "</pre>").setTitle(szTitle);
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function onInstall(e) {
  onOpen(e);
  // Perform additional setup as needed.
}

function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu() // Or DocumentApp.
      .addItem('Make textile', 'TextileNormalNoneMerged')
      .addItem('Make textile(S)', 'TextileSmallNoneMerged')
      .addItem('Make textile(S/Link)', 'TextileSmallNoneMergedWithLink')
//      .addItem('Make textileOrg', 'TextileNormal')
//      .addItem('Make textileOrg(S)', 'TextileSmall')
      .addToUi();
}

/// No Use
/*
function TextileNormal()
{
  return Textile(0);
}

function TextileSmall()
{
  return Textile(1);
}

function Textile(bSmall)
{
  var ro1, co1;
  var selection = SpreadsheetApp.getActiveSpreadsheet().getSelection();
  var activeRange = selection.getActiveRange();
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
//    console.log("Textile Line:[" + i + "]");
  }
//  console.log("Textile :" + szOutString);
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
*/
