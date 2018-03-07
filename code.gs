/**
 * @OnlyCurrentDoc
 */
var FIRST_COLUMN=4, FIRST_ROW=5, LAST_ROW=25;
function onOpen(e) {
  var ui=SpreadsheetApp.getUi();
  ui.createMenu("Attendance Helper").addItem("Show sidebar (please Allow access)", "showSidebar").addToUi();
  ui.alert("NEW: Attendance Helper", "The Attendance Helper is an easy and fast way to mark your attendance. At the top, click the \"Attendance Helper\" menu, then \"Show sidebar\". Select your current account and \"Allow\" access.", ui.ButtonSet.OK)
  showSidebar();
}
function showSidebar() {
  var html=HtmlService.createHtmlOutputFromFile("sidebar").setTitle("Attendance Helper");
  SpreadsheetApp.getUi().showSidebar(html);
}
function getSelected(lastRange) {
  var range=SpreadsheetApp.getActiveRange();
  if (range.getA1Notation()==lastRange) return null;
  var data={};
  if (range.getSheet().getSheetId()==1352368879 && range.getHeight()==1 && range.getWidth()==1 && range.getColumn()>=FIRST_COLUMN && range.getRow()>=FIRST_ROW && range.getRow()<=LAST_ROW) {
    data.noSelect=false;
    data.color=range.getFontColor();
    data.weight=range.getFontWeight();
    data.value=range.getValue();
  } else data.noSelect=true;
  data.lastRange=range.getA1Notation();
  return data;
}
function setValue(range, value, display) {
  if (display.indexOf(" ")==1) range.setValue(value+display.substring(1));
  else range.setValue(value);
}
function setTransport(value) {
  var range=SpreadsheetApp.getActiveRange();
  var display=range.getDisplayValue();
  if (value=="none") {
    range.setValue(null);
    range.setFontWeight(null);
    range.setFontColor(null);
  } else if (value=="sprinter") {
    setValue(range,"S",display);
    range.setFontWeight(null);
    range.setFontColor(null);
  } else if (value=="driving") {
    setValue(range,"D",display);
    range.setFontWeight(null);
    range.setFontColor(null);
  } else if (value=="maybe") {
    setValue(range,"M",display);
    range.setFontWeight(null);
    range.setFontColor("red");
  } else if (value=="notappd") {
    setValue(range,"N",display);
    range.setFontWeight(null);
    range.setFontColor("red");
  } else if (value=="approved") {
    setValue(range,"N",display);
    range.setFontWeight("bold");
    range.setFontColor(null);
  } else {
    setValue(range,"No Show",display);
    range.setFontWeight("bold");
    range.setFontColor("red");
  }
}
function setReason(value) {
  var range=SpreadsheetApp.getActiveRange();
  var display=range.getDisplayValue();
  var index=display.indexOf("(");
  if (index!=-1) {
    if (value=="") {
      if (display.substring(index-1,index)==" ") range.setValue(display.substring(0,index-1));
      else range.setValue(display.substring(0,index));
    } else range.setValue(display.substring(0,index+1)+value+")");
  } else if (value!="") range.setValue(display+" ("+value+")");
}
function setScore(value) {
  var range=SpreadsheetApp.getActiveRange();
  var display=range.getDisplayValue();
  var index=display.indexOf("(");
  if (index!=-1) {
    if (display.substring(index-1,index)==" ") range.setValue(value+display.substring(index-1));
    else range.setValue(value+display.substring(index));
  } else range.setValue(value);
}
function setMedalist(value) {
  var range=SpreadsheetApp.getActiveRange();
  if (value) {
    range.setFontWeight("bold");
    range.setFontColor("#01da00");
  } else {
    range.setFontWeight(null);
    range.setFontColor(null);
  }
}
