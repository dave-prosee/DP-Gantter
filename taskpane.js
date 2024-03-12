/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

//const IDColumn  = 1
const NameColumn  = 2
const FirstMovableColumn  = 3
const LastMovableColumn  = 19
const ColumnHeaderRow = 6
const beginRow  = 7 //first row for activities
const DarkeningDecimal = 0.3321
const settingsRange = "A1:S3"
const planningsRange = "A6:S7"
const standardRowHeight = 16
//Formatweekly uses Range that requires direct input (constants seem not to work)
//Range is up to column AZZ and row 1000

function myPlanningProperties(excel){
  const cellProperties = excel.CellPropertiesLoadOptions = {
    format: {
      font: {
        bold: true,
        color: true,
        italic: true,
        name: true,
        underline: true,
        size: true,
        strikethrough: true,
        subscript: true,
        superscript: true,
        tintAndShade: true
      }
      ,indentLevel : true
      ,fill: {
        color: true}
    }
  } 
  return cellProperties;
}

function setMessage(msg){document.getElementById("message").innerHTML = msg;}
function setProjectLabel(msg){document.getElementById("label_settings").innerHTML = msg;}
function getProjectLabel(){return document.getElementById("label_settings").innerHTML;}

function getSettingVvalue(label,array,props){
  let c = getColumn(label, array);
  if(c==-1){return c;}
  if(label.includes("color")) {return getFillColor(1,c,props)}
  else {return array.values[1][c]} 
}

function xlSerialToJsDate(xlSerial){
  // milliseconds since 1899-12-31T00:00:00Z, corresponds to Excel serial 0.
  var xlSerialOffset = -2209075200000; 

  var elapsedDays;
  // each serial up to 60 corresponds to a valid calendar date.
  // serial 60 is 1900-02-29. This date does not exist on the calendar.
  // we choose to interpret serial 60 (as well as 61) both as 1900-03-01
  // so, if the serial is 61 or over, we have to subtract 1.
  if (xlSerial < 61) {
    elapsedDays = xlSerial;
  }
  else {
    elapsedDays = xlSerial - 1;
  }

  // javascript dates ignore leap seconds
  // each day corresponds to a fixed number of milliseconds:
  // 24 hrs * 60 mins * 60 s * 1000 ms
  var millisPerDay = 86400000;
    
  var jsTimestamp = xlSerialOffset + elapsedDays * millisPerDay;
  return new Date(jsTimestamp);
}

function getFillColor(row,column, properties){return properties.m_value[row][column].format.fill.color}

function getColumn(label,array){
  return array.values[0].findIndex(searchlabel);
  function searchlabel(value, index, array) {return value == label;}
}

function getIDRow(ID,array){
  return array.findIndex(searchID);
  function searchID(value, index, array) {return value[0] == ID;}
}

function indentlevel(row,column, properties){return properties.m_value[row][column].format.indentLevel}

function darken(color){
  color = color.toString();
  var R=Number.parseInt(Number.parseInt("0x"+color.slice(1,3),16)*0.8).toString(16).toUpperCase().padStart(2,"0");
  var G=Number.parseInt(Number.parseInt("0x"+color.slice(3,5),16)*0.8).toString(16).toUpperCase().padStart(2,"0");
  var B=Number.parseInt(Number.parseInt("0x"+color.slice(5,7),16)*0.8).toString(16).toUpperCase().padStart(2,"0");
  //console.log(color+" --> #"+R+G+B)
  return "#"+R+G+B;
}

function randomgrey(){
  var y = Math.random() * (220-150) + 150 ; //(max-min)+min
  var YY=y.toString(16).toUpperCase().padStart(2,"0");
  console.log(color+" --> #"+YY+YY+YY)
  return "#"+YY+YY+YY;
}

function randomColor(){
  //color names, see : https://real-statistics.com/wp-content/uploads/2019/10/named-colors-list.png
  var colorname = ["tan", "olive", "lavender", "silver", "lightgrey", "grey"];
  var i = Number.parseInt(Math.random() * 100) % 6 ; //reminder is 0 to 5
  //console.log("i = "+i+", color = "+colorname[i]);
  return colorname[i];
}

function RGB(r, g, b){
  var R= r.toString(16).toUpperCase().padStart(2,"0");
  var G= g.toString(16).toUpperCase().padStart(2,"0");
  var B= b.toString(16).toUpperCase().padStart(2,"0");
  console.log("rgb becomes --> #"+R+G+B)
  return "#"+R+G+B;
}

//function findESColumn(value,index,array){return value == "ES";}
//function findIDColumn(value,index,array){return value == "ID";}

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("image_draw").onclick = draw;
    document.getElementById("image_move_up").onclick = move_up;
    document.getElementById("image_move_down").onclick = move_down;
    document.getElementById("image_move_left").onclick = move_left;
    document.getElementById("image_move_right").onclick = move_right;
    document.getElementById("image_indent").onclick = indent;
    document.getElementById("image_outdent").onclick = outdent;
    document.getElementById("image_collapse").onclick = collapse;
    document.getElementById("image_expand").onclick = expand;
    document.getElementById("image_grid").onclick = grid;
    document.getElementById("image_no_grid").onclick = nogrid;
    //document.getElementById("image_alternate_rows").onclick = alternateRows;
    //document.getElementById("image_clean_rows").onclick = cleanRows ;
    document.getElementById("image_take_style").onclick = run;
    document.getElementById("image_color").onclick = run;
    document.getElementById("settings").onclick = settings;
    document.getElementById("image_timeline").onclick = timeline;
    document.getElementById("image_new").onclick = testshape;
  }
});

export async function move_up() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet()
      const range = context.workbook.getSelectedRange();
      range.load("address");
      range.load("rowIndex")
      range.load("rowCount");
      await context.sync();
      //address will have format of Sheet1!B9 for single cell or Sheet1!B11:C14 for multiple cells
      //current version only support one row
      if (range.rowCount>1){
        setMessage("Select only one line to move up");
        return;
      }
      //can only move up if selected range is 8 or greater
      var row = Number.parseInt(range.rowIndex+1); //row index starts at 0, instead of line no in UI. 
      if (row < 8){ 
        setMessage("Can only move up lines beyond row 7");
        return;
      }
      //correct selection, now we perform requested function
      const rangeER = range.getEntireRow();
      rangeER.load("values")
      const destination = rangeER.getOffsetRange(-1,0).insert(Excel.InsertShiftDirection.down);
      destination.select()
      //destination.copyFrom(rangeER); //default copy type ALL does not work, hence a move to
      rangeER.moveTo(destination);  //this will change the addres of range and rangeER
      const oldrange =  sheet.getRange((row+1)+":"+(row+1)); 
      oldrange.delete(Excel.DeleteShiftDirection.up);
      await context.sync();
      console.log("row "+row+" moved up to "+ (row-1))
    });
  } catch (error) {
    console.error(error);
  }
}

export async function move_down() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet()
      const range = context.workbook.getSelectedRange();
      range.load("address");
      range.load("rowIndex");
      range.load("rowCount");
      await context.sync();
      //address will have format of Sheet1!B9 for single cell or Sheet1!B11:C14 for multiple cells
      //current version only support one row
      if (range.rowCount>1){
        setMessage("Select only one line to move down");
        return;
      }
      //can only move down if selected line is 7 or greater
      var row = Number.parseInt(range.rowIndex+1); //row index starts at 0, instead of line no in UI.
      if (row < 7){ 
        setMessage("Can only move down lines beyond row 6");
        return;
      }
      //correct selection, now we perform requested function
      const rangeER = range.getEntireRow();
      rangeER.load("values")
      const destination = rangeER.getOffsetRange(2,0).insert(Excel.InsertShiftDirection.down);
      destination.getOffsetRange(-1,0).select()
      //destination.copyFrom(rangeER); //default copy type ALL does not work, hence a move to
      rangeER.moveTo(destination);  //this will change the addres of range and rangeER
      const oldrange =  sheet.getRange(row+":"+row); //row remains the same as row inserted is below it
      oldrange.delete(Excel.DeleteShiftDirection.up);
      await context.sync()
      console.log("row "+row+" moved down to "+ (row+1))
    });
  } catch (error) {
    console.error(error);
  }
}

export async function move_left() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet()
      const range = context.workbook.getSelectedRange();
      range.load("address");
      range.load("columnIndex");
      range.load("columnCount");
      await context.sync();
      //address will have format of Sheet1!B9 for single cell or Sheet1!B11:C14 for multiple cells
      //current version only support one row
      if (range.columnCount>1){
        setMessage("Select only one column to move");
        return;
      }
      //ID and Name are not allowed to move left
      var column = Number.parseInt(range.columnIndex+1); //column index starts at 0
      if (column < 3){ 
        setMessage("Cannot move ID and Name");
        return;
      }
      //Cannot move columns T and beyond
      var column = Number.parseInt(range.columnIndex+1); //column index starts at 0
      if (column > 19){ 
        setMessage("Column T and beyond cannot be moved left");
        return;
      }
      //correct selection, now we perform requested function
      const rangeER = range.getEntireColumn();
      rangeER.load("values")
      const destination = rangeER.getOffsetRange(0,-1).insert(Excel.InsertShiftDirection.right);
      destination.select(); 
      rangeER.moveTo(destination);  //this will change the addres of range and rangeER
      var col = String.fromCharCode(64+column+1); //the original column is one more to the right because we insert one column
      const oldrange =  sheet.getRange(col+":"+col);
      oldrange.delete(Excel.DeleteShiftDirection.left);
      await context.sync()
      console.log("column "+String.fromCharCode(64+column)+" moved left to "+ String.fromCharCode(63+column))
    });
  } catch (error) {
    console.error(error);
  }
}

export async function move_right() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet()
      const range = context.workbook.getSelectedRange();
      range.load("address");
      range.load("columnIndex");
      range.load("columnCount");
      await context.sync();
      //address will have format of Sheet1!B9 for single cell or Sheet1!B11:C14 for multiple cells
      //current version only support one row
      if (range.columnCount>1){
        setMessage("Select only one column to move");
        return;
      }
      //ID and Name are not allowed to move left
      var column = Number.parseInt(range.columnIndex+1); //column index starts at 0
      if (column < 3){ 
        setMessage("Cannot move ID and Name");
        return;
      }
      //Cannot move columns T and beyond
      var column = Number.parseInt(range.columnIndex+1); //column index starts at 0
      if (column > 18){ 
        setMessage("Column S and beyond cannot be moved right");
        return;
      }
      //correct selection, now we perform requested function
      const rangeER = range.getEntireColumn();
      rangeER.load("values")
      const destination = rangeER.getOffsetRange(0,2).insert(Excel.InsertShiftDirection.right);
      destination.getOffsetRange(0,-1).select(); 
      rangeER.moveTo(destination);  //this will change the addres of range and rangeER
      var col = String.fromCharCode(64+column);
      const oldrange =  sheet.getRange(col+":"+col);
      oldrange.delete(Excel.DeleteShiftDirection.left);
      await context.sync();
      console.log("column "+String.fromCharCode(64+column)+" moved right to "+ String.fromCharCode(65+column))
    });
  } catch (error) {
    console.error(error);
  }
}

export async function indent() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("address");
      range.load("rowIndex");
      range.load("rowCount");
      await context.sync();
      //can only indent if selected line is 8 or greater
      var row = Number.parseInt(range.rowIndex+1); //row index starts at 0, instead of line no in UI.
      if (row < 8){ 
        setMessage("Cannot indent line above line 8");
        return;
      }
      //get the correct range
      const ws = context.workbook.worksheets.getActiveWorksheet();
      const cellProperties = Excel.CellPropertiesLoadOptions = {
        format: {
          font: {
            bold: true,
            color: true,
            italic: true,
            name: true,
            underline: true,
            size: true,
            strikethrough: true,
            subscript: true,
            superscript: true,
            tintAndShade: true
          }
          ,indentLevel : true
          ,fill: {
            color: true}
        }
      };
      var address = "B"+row+":B"+(row+range.rowCount-1);
      console.log("address =" + address);
      var NameProperties = ws.getRange(address).getCellProperties(cellProperties);
      var Namerange = ws.getRange(address);
      await context.sync();

      //set new indentation
      var NewProperties =[]
      for (var line=0; line<NameProperties.m_value.length; line++){
        NewProperties[line] = [{format:{indentLevel: indentlevel(line,0,NameProperties)+1}}]
      }
      Namerange.setCellProperties(NewProperties);
      await context.sync();
      console.log("Indentation done");
      return;
    });
  } catch (error) {
    console.error(error);
  }
}

export async function outdent() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("address");
      range.load("rowIndex");
      range.load("rowCount");
      await context.sync();
      //can only indent if selected line is 8 or greater
      var row = Number.parseInt(range.rowIndex+1); //row index starts at 0, instead of line no in UI.
      if (row < 7){ 
        setMessage("Cannot outdent line above line 7");
        return;
      }
      //get the correct range
      const ws = context.workbook.worksheets.getActiveWorksheet();
      const cellProperties = Excel.CellPropertiesLoadOptions = {
        format: {
          font: {
            bold: true,
            color: true,
            italic: true,
            name: true,
            underline: true,
            size: true,
            strikethrough: true,
            subscript: true,
            superscript: true,
            tintAndShade: true
          }
          ,indentLevel : true
          ,fill: {
            color: true}
        }
      };
      var address = "B"+row+":B"+(row+range.rowCount-1);
      console.log("address =" + address);
      var NameProperties = ws.getRange(address).getCellProperties(cellProperties);
      var Namerange = ws.getRange(address);
      await context.sync();

      //set new indentation
      var NewProperties =[]
      for (var line=0; line<NameProperties.m_value.length; line++){
        NewProperties[line] = [{format:{indentLevel: Math.max(0,indentlevel(line,0,NameProperties)-1)}}]
      }
      Namerange.setCellProperties(NewProperties);
      await context.sync();
      console.log("Outdentation done");
      return;
    });
  } catch (error) {
    console.error(error);
  }
}

export async function settings() {
  try {
    await Excel.run(async (context) => {
      console.log("in show_settings")
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const settingsrange = sheet.getRange("1:3");
      settingsrange.load("rowHidden");
      await context.sync()
      console.log(settingsrange.rowHidden)
      if (settingsrange.rowHidden){
        settingsrange.rowHidden = false;
        setMessage("Settings can be done in range 'A1:S2'");
        settingsrange.select();
      }
      else {
        settingsrange.rowHidden = true;
        setMessage("Settings are now hidden");
      }
      await context.sync()
      });
    } 
    catch (error) {
      console.error(error);
    }
}

export async function collapse() {
  try {
    await Excel.run(async (context) => {
      console.log("in collapse")
      const ws = context.workbook.worksheets.getActiveWorksheet()
      const range = context.workbook.getSelectedRange().getEntireRow();
      range.load("address");
      range.load("values")
      range.load("rowIndex");
      range.load("rowCount");
      const ids = ws.getRange("A7:A9999").getUsedRange(); //TODO: when fields below table are colored, they are, but should not, included in used range
      ids.load("address")
      await context.sync();

      //address will have format of Sheet1!B9 for single cell or Sheet1!B11:C14 for multiple cells
      //current version only support one row
      if (range.rowCount>1){
        setMessage("Select only one Phase to collapse");
        return;
      }
      //can only move down if selected line is 7 or greater
      var row = Number.parseInt(range.rowIndex+1); //row index starts at 0, instead of line no in UI.
      if (row < 7){ 
        setMessage("Can only collapse lines beyond row 6");
        return;
      }
      var address=ids.address;
      const planning = ws.getRange(ids.address.replace(":A", ":S").replace("A7:","A6:"))
      planning.load("values")
      planning.load(["rowCount"]);
      var planningProperties = ws.getRange(ids.address.replace(":A", ":S").replace("A7:","A6:")).getCellProperties(myPlanningProperties(Excel));
      await context.sync();

      let ColorTypeColumn = getColumn("Tasktype / Color", planning)
      if (ColorTypeColumn==-1) {setMessage("Columnheader 'Tasktype / Color' not found between C6 and S6, processing stopped. Please correct."); return;}
      let NameColumn = getColumn("Activity", planning)
      if (ColorTypeColumn==-1) {setMessage("Columnheader 'Activity' not found in B6, processing stopped. Please correct."); return;}
      let IDColumn = getColumn("ID", planning)
      if (IDColumn!=0) {setMessage("Columnheader 'ID' not found in A6, processing stopped. Please correct."); return;}
      if (planning.values[row-6][ColorTypeColumn].toString().toLowerCase() != "phase"){ 
        setMessage("Can only collapse Phase ines");
        return;
      }
      //correct selection, now we perform requested function
      var PhaseIndentLevel = indentlevel(row-6, NameColumn, planningProperties)
      var PhaseID = planning.values[row-6][IDColumn];
      var endrow = row;
      for (var i = row-5; (i<planning.rowCount && PhaseIndentLevel<indentlevel(i,NameColumn,planningProperties)); i++){
        endrow=endrow+1;
      }
      const collapserange = ws.getRange((row+1)+":"+endrow);
      console.log("collapse of lines "+(row+1)+" to "+endrow)
      collapserange.rowHidden=true;
      range.select();    
      setMessage("Subtasks of ID "+PhaseID+" collapsed");
      await context.sync();
      return;
    });
  } 
  catch (error) {
    console.error(error);
  } 
}

export async function expand() {
  try {
    await Excel.run(async (context) => {
      console.log("in expand")
      const ws = context.workbook.worksheets.getActiveWorksheet()
      const range = context.workbook.getSelectedRange().getEntireRow();
      range.load("address");
      range.load("values")
      range.load("rowIndex");
      range.load("rowCount");
      const ids = ws.getRange("A7:A9999").getUsedRange(); //TODO: when fields below table are colored, they are, but should not, included in used range
      ids.load("address")
      await context.sync();

      //address will have format of Sheet1!B9 for single cell or Sheet1!B11:C14 for multiple cells
      //current version only support one row
      if (range.rowCount>1){
        setMessage("Select only one Phase to expand");
        return;
      }
      //can only move down if selected line is 7 or greater
      var row = Number.parseInt(range.rowIndex+1); //row index starts at 0, instead of line no in UI.
      if (row < 7){ 
        setMessage("Can only expand lines beyond row 6");
        return;
      }
      var address=ids.address;
      const planning = ws.getRange(ids.address.replace(":A", ":S").replace("A7:","A6:"))
      planning.load("values")
      planning.load(["rowCount"]);
      var planningProperties = ws.getRange(ids.address.replace(":A", ":S").replace("A7:","A6:")).getCellProperties(myPlanningProperties(Excel));
      await context.sync();

      let ColorTypeColumn = getColumn("Tasktype / Color", planning)
      if (ColorTypeColumn==-1) {setMessage("Columnheader 'Tasktype / Color' not found between C6 and S6, processing stopped. Please correct."); return;}
      let NameColumn = getColumn("Activity", planning)
      if (ColorTypeColumn==-1) {setMessage("Columnheader 'Activity' not found in B6, processing stopped. Please correct."); return;}
      let IDColumn = getColumn("ID", planning)
      if (IDColumn!=0) {setMessage("Columnheader 'ID' not found in A6, processing stopped. Please correct."); return;}
      if (planning.values[row-6][ColorTypeColumn].toString().toLowerCase() != "phase"){ 
        setMessage("Can only collapse Phase ines");
        return;
      }
      //correct selection, now we perform requested function
      var PhaseIndentLevel = indentlevel(row-6, NameColumn, planningProperties)
      var PhaseID = planning.values[row-6][IDColumn];
      var endrow = row;
      for (var i = row-5; (i<planning.rowCount && PhaseIndentLevel<indentlevel(i,NameColumn,planningProperties)); i++){
        endrow=endrow+1;
      }
      const collapserange = ws.getRange((row+1)+":"+endrow);
      console.log("expansion of lines "+(row+1)+" to "+endrow)
      collapserange.rowHidden=false;
      range.select();    
      setMessage("Subtasks of ID "+PhaseID+" expanded");
      await context.sync();
      return;
    });
  } 
  catch (error) {
    console.error(error);
  } 
}

export async function grid() {
  try {
    await Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getActiveWorksheet();
      ws.showGridlines = true;
      await context.sync();
    });
  } 
  catch (error) {
      console.error(error);
  }
  }


export async function nogrid() {
  try {
    await Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getActiveWorksheet();
      ws.showGridlines = false;
      await context.sync();
    });
  } 
  catch (error) {
      console.error(error);
  }
  }

  /*

export async function alternateRows(){
  try {
      await Excel.run(async (context) => {
      console.log("in alternateRows")
      const ws = context.workbook.worksheets.getActiveWorksheet()
      const ids = ws.getRange("A7:A9999").getUsedRange(); //TODO: when fields below table are colored, they are, but should not, included in used range
      ids.load("address")
      await context.sync();

      const planning = ws.getRange(ids.address.replace(":A", ":S").replace("A7:","A6:"))
      planning.load("values")
      planning.load(["rowCount"]);
      planning.load("rowHidden")
      planning.load("hidden")
      const planningProperties = ws.getRange(ids.address.replace(":A", ":S").replace("A7:","A6:")).getCellProperties(myPlanningProperties(Excel));
      await context.sync();
      console.log(planning);
      console.log(planning.rowHidden);
      console.log(planningProperties);

      let ColorTypeColumn = getColumn("Tasktype / Color", planning)
      if (ColorTypeColumn==-1) {setMessage("Columnheader 'Tasktype / Color' not found between C6 and S6, processing stopped. Please correct."); return;}
      let IDColumn = getColumn("ID", planning)
      if (IDColumn!=0) {setMessage("Columnheader 'ID' not found in A6, processing stopped. Please correct."); return;}
      
      var row = 2;
      var ColorThisLine = true;
      while(planning.values[row][IDColumn] != "" & row<7) {  //range.rowCount
          var range = ws.getRange("A"+(row+6)+":S"+(row+6))
          range.load("hidden")
          range.load("format/fill/color")
          var rangeProperties = ws.getRange("A"+(row+6)+":S"+(row+6)).getCellProperties(myPlanningProperties(Excel))
          await context.sync()
          console.log("hidden = "+range.hidden)
          if (range.hidden == false) {
              console.log("visible")
              if (ColorThisLine){
                  console.log("color this line")
                  var shapeColor = getFillColor(row,ColorTypeColumn, planningProperties);
                  range.format.fill.color = "#F9F9F9" //.RGB = "xFAFAFA";
                  await context.sync()
                  console.log(shapeColor)
                  console.log(rangeProperties.value[0][ColorTypeColumn])
                  //planningProperties.value[row][ColorTypeColumn].format.fill.color = shapeColor;
                  rangeProperties.value[0][ColorTypeColumn].format.fill.color = shapeColor
              }
              ColorThisLine = ! ColorThisLine
          }
          row = row + 1;
          console.log("next row")
      } 
      await context.sync();

  });
} 
catch (error) {
    console.error(error);
}
}



export async function cleanRows(){
  console.log("in MarkAlternateRows")
  Row = beginRow
  ColorThisLine = truelet 
  
  
  while(Cells(Row, IDColumn) != "")
      if (Cells(Row, IDColumn).EntireRow.Hidden != true) 
        {
          if (ColorThisLine)
            {
              activity_color = Cells(Row, ColorColumn).Interior.Color
              Cells(Row, IDColumn).EntireRow.Interior.Color = RGB(250, 250, 250)
              Cells(Row, ColorColumn).Interior.Color = activity_color
            }
          ColorThisLine = ! ColorThisLine
        }
      Row = Row + 1
  Wend
  console.log("at the end of it")
}
*/


export async function run() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
     
      await context.sync();
      console.log(context);
      console.log(sheet.zoom)


    });
  } 
  catch (error) {
      console.error(error);
  }
  }


export async function do_some_test(){ //testproperties 
  try 
  {
    Excel.run (async (context)=>{
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const rangeSrc = sheet.getRange("b7:b9");
      //const rangeDst = sheet.getRange("D1:E2");
      var properties = rangeSrc.getCellProperties(cellProperties);
          
      await context.sync();
      console.log(properties)
      console.log("indentlevel regel 0 : "+properties.m_value[0][0].format.indentLevel)
      console.log("indentlevel regel 1 : "+properties.m_value[1][0].format.indentLevel)
      console.log("indentlevel regel 2 : "+indentlevel(2,0,properties) )
      console.log("fillcolor   regel 2 : "+getFillColor(2,0,properties) )
    });
  } 
  catch (error) 
  {
    console.error(error);
  }
}

export  async function draw(){
  setMessage("Starting draw")
  console.log("starting draw") 
  //TODO: set zoom level to 100%
  //TODO: set equidistant column widths 
  try 
  {
    Excel.run (async (context)=>{
        const ws = context.workbook.worksheets.getActiveWorksheet()
        
        //get IDs
        const ids = ws.getRange("A7:A9999").getUsedRange(); //TODO: when fields below table are colored, they are, but should not, included in used range
        ids.load("address")
        ids.load("values")
        ids.load(["rowCount"]);
        //Settings values and properties
        const settings = ws.getRange(settingsRange);
        settings.load("values");
        settings.load("rowHidden");
        settings.format.rowHeight=standardRowHeight;
        settings.load("format/rowHeight");
        const settingsProperties = settings.getCellProperties({
          address: true,
          format: {
              fill: {
                  color: true
              },
              font: {
                  color: true
              }
          },
          style: true,
          indentLevel: true
        });
        //drawing metrics
        const firstGraphColumn = ws.getRange("T:T")
        firstGraphColumn.load("width");
        firstGraphColumn.load("left");
        firstGraphColumn.load("format/columnWidth")
        const secondGraphColumn = ws.getRange("U:U")
        secondGraphColumn.load("left");
        const firstGraphRow = ws.getRange("T7:T7")
        firstGraphRow.load("top");
        //shapes
        const shapes = ws.shapes;
        shapes.load("items"); //   /$none");
        
        //collect
        await context.sync();
        console.log(settings)
        console.log(settingsProperties)
        console.log("no of ids found : "+ids.rowCount)
        console.log(shapes)

        //Save time axis
        const TimeUnitWidth = firstGraphColumn.width;
        const settingsHidden = settings.rowHidden;
       
        //create equidistant columns for neat drawing on the grid
        console.log("Column T has width of "+TimeUnitWidth+" and a column width of "+firstGraphColumn.format.columnWidth)
        var grid=ws.getRange("U:AZZ").format.columnWidth=TimeUnitWidth;

        //save setting values
        var PlanningStart = getSettingVvalue("Planning start", settings,settingsProperties)
        if (PlanningStart == -1){
          setMessage("No settings found for 'Planning start'. Please add label between A1 and S1, and its value below it.");
          return context;
        }
        var PlanningUnit = getSettingVvalue("Planning unit", settings,settingsProperties)
        if (PlanningUnit == -1){
          setMessage("No settings found for 'Planning unit'. Please add label between A1 and S1, and its value below it.");
          return context;
        }
        var HoursPerBlock = getSettingVvalue("Hours per unit", settings,settingsProperties)
        if (HoursPerBlock == -1){
          setMessage("No settings found for 'Hours per unit'. Please add label between A1 and S1, and its value below it.");
          return context;
        }
        var CriticalPathColor = getSettingVvalue("Critical path color", settings,settingsProperties)
        if (CriticalPathColor == -1){
          setMessage("No settings found for 'Critical path color'. Please add label between A1 and S1, and its value below it.");
          return context;
        }
        console.log("Critical path color = "+CriticalPathColor);
        var NormalPathColor = getSettingVvalue("Normal path color", settings,settingsProperties)
        if (NormalPathColor == -1){
          setMessage("No settings found for 'Normal path color'. Please add label between A1 and S1, and its value below it.");
          return context;
        }
        var DrawTimeLine = getSettingVvalue("Draw Time Line", settings,settingsProperties)
        if (DrawTimeLine == -1){
          setMessage("No settings found for 'Draw Time Line'. Please add label between A1 and S1, and its value below it.");
          return context;
        }
        var TimeLineValue = getSettingVvalue("Time line value", settings,settingsProperties)
        if (TimeLineValue == -1){
          setMessage("No settings found for 'Time line value'. Please add label between A1 and S1, and its value below it.");
          return context;
        }
        
        //All planning variables should be valid
        if (PlanningUnit ==0 || HoursPerBlock == 0 || CriticalPathColor == 0 || NormalPathColor == 0 || DrawTimeLine == 0 
            || (DrawTimeLine == "Yes" && TimeLineValue == 0)){
            setMessage("Some settings have no correct values. Please check the values in range B2 to S2.");
            return context;
          }

        //no check for abstract  planning, will be discontinued

        //get planning values and propoerties
        const planning = ws.getRange(ids.address.replace(":A", ":S").replace("A7:","A6:"))
        planning.load("values")
        planning.load(["rowCount"]);
        planning.format.rowHeight=standardRowHeight;
        planning.load("format/rowHeight")
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const rangeSrc = sheet.getRange("b7:b9");
      //const rangeDst = sheet.getRange("D1:E2");
      const cellProperties = Excel.CellPropertiesLoadOptions = {
        format: {
          font: {
            bold: true,
            color: true,
            italic: true,
            name: true,
            underline: true,
            size: true,
            strikethrough: true,
            subscript: true,
            superscript: true,
            tintAndShade: true
          }
          ,indentLevel : true
          ,fill: {
            color: true}
        }
      };
      var planningProperties = ws.getRange(ids.address.replace(":A", ":S").replace("A7:","A6:")).getCellProperties(cellProperties);
          
      await context.sync();
      console.log(planning)
      console.log(planningProperties)
      const RowHeight = planning.format.rowHeight;
      console.log("planning height"+RowHeight)
      /*console.log("indentlevel regel 0, kolom 0 : "+planningProperties.m_value[0][0].format.indentLevel)
      console.log("indentlevel regel 1, kolom 1 : "+planningProperties.m_value[1][0].format.indentLevel)
      console.log("indentlevel regel 2, kolom 1 : "+indentlevel(2,0,planningProperties) )
      console.log("fillcolor   regel 2, kolom 1 : "+getFillColor(2,0,planningProperties) )
        
    
      console.log("indentlevel regel 0, kolom 1 : "+indentlevel(0,1,planningProperties))
      console.log("indentlevel regel 1, kolom 1 : "+indentlevel(1,1,planningProperties))
      console.log("indentlevel regel 2, kolom 1 : "+indentlevel(2,1,planningProperties) )
      console.log("indentlevel regel 3, kolom 1 : "+indentlevel(3,1,planningProperties) )

      console.log("fillcolor   regel 0 : "+getFillColor(0,9 ,planningProperties) )
      console.log("fillcolor   regel 1 : "+getFillColor(1,9 ,planningProperties) )
      console.log("fillcolor   regel 2 : "+getFillColor(2,9 ,planningProperties) )
      console.log("fillcolor   regel 3 : "+getFillColor(3,9 ,planningProperties) )
       */

      //Column placement (ESColum = From, EE=Until)
      let ESColumn = getColumn("From", planning)
      if (ESColumn==-1) {setMessage("Columnheader 'From' not found between C6 and S6, processing stopped. Please correct."); return;}
      let LSColumn = getColumn("LS", planning)
      if (LSColumn==-1) {setMessage("Columnheader 'LS' not found between C6 and S6, processing stopped. Please correct."); return;}
      let EEColumn = getColumn("Until", planning)
      if (EEColumn==-1) {setMessage("Columnheader 'Until' not found between C6 and S6, processing stopped. Please correct."); return;}
      let LEColumn = getColumn("LE", planning)
      if (LEColumn==-1) {setMessage("Columnheader 'LE' not found between C6 and S6, processing stopped. Please correct."); return;}
      let DurationColumn = getColumn("Duration", planning)
      if (DurationColumn==-1) {setMessage("Columnheader 'Duration' not found between C6 and S6, processing stopped. Please correct."); return;}
      let CriticalPathColumn = getColumn("Critical path", planning)
      if (CriticalPathColumn==-1) {setMessage("Columnheader 'Critical path' not found between C6 and S6, processing stopped. Please correct."); return;}
      let DependencyColumn = getColumn("Dependency", planning)
      if (DependencyColumn==-1) {setMessage("Columnheader 'Dependency' not found between C6 and S6, processing stopped. Please correct."); return;}
      let ProgressColumn = getColumn("Progress", planning)
      if (ProgressColumn==-1) {setMessage("Columnheader 'Progress' not found between C6 and S6, processing stopped. Please correct."); return;}
      let ResponsibleColumn = getColumn("Responsible", planning)
      if (ResponsibleColumn==-1) {setMessage("Responsible 'Critical path' not found between C6 and S6, processing stopped. Please correct."); return;}
      let SNBColumn = getColumn("Start not before", planning)
      if (SNBColumn==-1) {setMessage("Columnheader 'Start not before' not found between C6 and S6, processing stopped. Please correct."); return;}
      let ColorTypeColumn = getColumn("Tasktype / Color", planning)
      if (ColorTypeColumn==-1) {setMessage("Columnheader 'Tasktype / Color' not found between C6 and S6, processing stopped. Please correct."); return;}
      let IDColumn = getColumn("ID", planning)
      if (IDColumn!=0) {setMessage("Columnheader 'ID' not found in A6, processing stopped. Please correct."); return;}
      let NameColumn = getColumn("Activity", planning)
      if (ColorTypeColumn==-1) {setMessage("Columnheader 'Activity' not found in B6, processing stopped. Please correct."); return;}
      
      console.log("ES column : "+ ESColumn)
      console.log("LS column : "+ LSColumn)
      console.log("EE column : "+ EEColumn)
      console.log("LE column : "+ LEColumn)
      console.log("Duration column : "+ DurationColumn)
      console.log("Critical path column : "+ CriticalPathColumn)
      console.log("Dependency column : "+ DependencyColumn)
      console.log("Progress column : "+ ProgressColumn)
      console.log("Responsible column : "+ ResponsibleColumn)
      console.log("SNB column : "+ SNBColumn)
      console.log("Color/Type column : "+ ColorTypeColumn)

      //validation of IDs
      setMessage("Validating IDs")
      console.log("Validating IDs")
      for (var row = 0; row < ids.rowCount; row++) {
        //validate IDs
        let ID = ids.values[row][0];
        //should not be empty
        if (ID =="")  {
          setMessage("ID value found in cell A" + (row+7) + " is missing. Processesing stopped.\n Maybe you have content or cell coloring below your planning lines. Please correct."); return;
        }
        //should be an integer
        if (isNaN(ID) ) {
          setMessage("ID value " + ID + ", found in cell A" + (row+7) + " is not a number. Processing stopped. Please correct."); return;
        }
        if (!Number.isInteger(Number(ID))) {
          setMessage("ID value " + ID + ", found in cell A" + (row+7) + " is not an integer. Processing stopped. Please correct."); return;
        }
        //check if not duplicate
        for (var rr = row+1; rr< ids.rowCount; rr++){
          if (ids.values[rr][0] == ID) {
            setMessage("ID value " + ID + ", in cell "+ (row+7) + " has a duplicate found in cell A" + (rr+7) + ". IDs should be unique. Processing stopped. Please correct."); return;
          }
        }
      }

      //phases should not have dependencies
      for (var row = 1; row < planning.rowCount; row++) {
        let ID = planning.values[row][0];
        if (planning.values[row][ColorTypeColumn].toString().toLowerCase() == "phase" 
            && planning.values[row][DependencyColumn] != "")  {
          setMessage("Activity with ID " + planning.values[row][0] +  " is a phase and should not have dependencies. Processing stopped. Please correct."); return;
        }
      }

      let NewCalculation = planning.values
      console.log("NewCalculation:");
      console.log(NewCalculation);
      console.log("NewCalculation rowcount : "+NewCalculation.length)

      setMessage("Initializing")
      //duration set to 0 if not filled, also ES, LS, EE, LS cleared
      for (var row = 1; row < NewCalculation.length; row++) {
        if (NewCalculation[row][DurationColumn] == ""){
          NewCalculation[row][DurationColumn] = 0;
          //console.log("duration op nul in row "+(row+6))
        }
        NewCalculation[row][ESColumn]="";
        NewCalculation[row][LSColumn]="";
        NewCalculation[row][EEColumn]="";
        NewCalculation[row][LEColumn]="";
        NewCalculation[row][IDColumn]= Number.parseInt(NewCalculation[row][IDColumn])
      }

      setMessage("Calculating");
      //Fill start for non dependent activities. Excluding phases
      for (var row = 1; row < planning.rowCount; row++) {
        if (NewCalculation[row][DependencyColumn] == "" && (NewCalculation[row][ColorTypeColumn].toString().toLowerCase() != "phase"))  {
          NewCalculation[row][ESColumn] = PlanningStart;
          if (NewCalculation[row][SNBColumn] > NewCalculation[row][ESColumn]) {
            NewCalculation[row][ESColumn] = NewCalculation[row][SNBColumn];
          }
          NewCalculation[row][EEColumn] = NewCalculation[row][ESColumn] + NewCalculation[row][DurationColumn];
        }
      }

      //Calculate ES and EE for activities with dependencies, excluding phases
      var Calculation_ready = false //start condition
      var MissingDependency = false
      var dependencies = [];
      var dependency = 0;
      var dependencyID;
      var dependencyRow = 0;
      while (! Calculation_ready){
        Calculation_ready = true
        for (row = 1; row<planning.rowCount; row++){
            //if dependent and not already calculated
            if (NewCalculation[row][DependencyColumn].toString() != "" && NewCalculation[row][ESColumn] == "" && NewCalculation[row][ColorTypeColumn].toString().toLowerCase() != "phase") {
              //prep for multiple dependencies by tidying up and splitting into array
                dependency = NewCalculation[row][DependencyColumn].toString()
                dependency = dependency.replaceAll(",", " ")
                dependency = dependency.replaceAll("  ", " ")
                dependencies = dependency.split(" ")    //split the words of the string, i.e. space
                
                //each dependency should be present and already filled, else we check next round
                MissingDependency = false
                for (var j = 0;  j<dependencies.length; j++){
                   dependencyID = 0+dependencies[j]
                   dependencyRow = getIDRow(dependencyID, NewCalculation)  
                    if (dependencyRow == -1) {
                        setMessage("Dependency " + dependencyID  + " does not refer to an existing ID. Please correct. Calculation halted.")
                        return;
                    }
                    if (dependencyRow == row) {
                        setMessage("Dependency on itself is not allowed for ID " + dependencyID);
                        return;
                    }
                    if (dependencyID == 0) {
                        setMessage("Dependency should not be based on ID 0. ID 0 is reserved. Please correct. Calculation stopped.");
                        return;
                    }
                    
                    if (NewCalculation[dependencyRow][EEColumn] != ""){
                        if ((NewCalculation[row][ESColumn] == "") || (NewCalculation[row][ESColumn] < NewCalculation[dependencyRow][EEColumn])) {
                          NewCalculation[row][ESColumn] = NewCalculation[dependencyRow][EEColumn];
                          NewCalculation[row][EEColumn] = NewCalculation[dependencyRow][EEColumn] + NewCalculation[row][DurationColumn];
                        }
                    }
                    else {
                        //Remember we have a missing value, but keep calculating the complete dependency list
                        MissingDependency = true
                    }
                }

                //Clear the contents after processing the complete dependencylist
                if (MissingDependency){
                        Calculation_ready = false
                        //Clear current calculations, because these are not based on all required input
                        NewCalculation[row][ESColumn]="";
                        NewCalculation[row][EEColumn]="";
                }

                //If we have a value, then align it according SNB
                if (NewCalculation[row][ESColumn] != ""){
                    if (NewCalculation[row][SNBColumn] > NewCalculation[row][ESColumn]){
                      NewCalculation[row][ESColumn] = NewCalculation[row][SNBColumn];
                      NewCalculation[row][EEColumn] = NewCalculation[row][ESColumn] + NewCalculation[row][DurationColumn];
                    }
                }
                
            }
        }
      }

      //LE for whole project equals max EE found (empty phase EE does not matter)
      var globalLE = 0
      for (row = 1; row< NewCalculation.length; row++) {
        if (NewCalculation[row][EEColumn] > globalLE) {globalLE = NewCalculation[row][EEColumn]}
      }
      console.log("Global LE :"+globalLE)

      
      //Calculate the array with successor activities
      var follower= [];
      for (row = 1; row<NewCalculation.length;row++){
          follower[row] = "0"
      }

      //Followers by row values
      var thisId = "";
      for (row=1; row<NewCalculation.length; row++) {
          thisId = NewCalculation[row][IDColumn];
          if (NewCalculation[row][DependencyColumn] != "") {
              //for every dependancy
              dependency = NewCalculation[row][DependencyColumn].toString()
              dependency = dependency.replaceAll(",", " ")
              dependency = dependency.replaceAll("  ", " ")
              dependencies = dependency.split(" ")    //split the words of the string, i.e. space
              for (var j = 0;  j<dependencies.length; j++){
                  dependency = 0 + dependencies[j]
                  dependencyRow = getIDRow(dependency, NewCalculation)
                  follower[dependencyRow] = follower[dependencyRow] + " " + thisId      //store like words in a string
                  follower[dependencyRow] = follower[dependencyRow].replaceAll(",", " ")
                  follower[dependencyRow] = follower[dependencyRow].replaceAll("  ", " ")
                  follower[dependencyRow] = follower[dependencyRow].trim()
              }
          }
      }
      console.log("follower : ");
      console.log(follower);

      //Assign LE for last activities or activities with no followers, excluding phases
      for (row = 1; row < NewCalculation.length; row++) {
          //not necessary : ID = NewCalculation[row][IDColumn];
          if (NewCalculation[row][EEColumn] == globalLE 
              || follower[row] == "0" && NewCalculation[row][ColorTypeColumn].toString().toLowerCase() != "phase" ){
              NewCalculation[row][LEColumn] = globalLE;
              NewCalculation[row][LSColumn] = globalLE - NewCalculation[row][DurationColumn];
          }
      }

      //Calculate LS and LE of  activities with followers, excluding phases
      var follower_actions = [];
      var followerID;
      var LE;
      var LS;
      var follower_index;
      var followerRow;
      Calculation_ready = false //start condition
      while (! Calculation_ready){
        Calculation_ready = true
        for (row = 1; row<NewCalculation.length; row++){
            if (follower[row] != "0" && NewCalculation[row][ColorTypeColumn].toString().toLowerCase() != "phase" ){
                follower_actions = follower[row].split(" ");    //split the words of the string
                for (follower_index = 1; follower_index < follower_actions.length; follower_index++){ //skip first entry because it is 0
                    followerID = Number.parseInt(follower_actions[follower_index])
                    followerRow = getIDRow(followerID,NewCalculation); 
                    if (NewCalculation[followerRow][LSColumn] != "") {
                        LE = NewCalculation[followerRow][LSColumn];
                        if (NewCalculation[row][LEColumn] > LE || NewCalculation[row][LEColumn] == "") { NewCalculation[row][LEColumn] = LE;} //only new value if it is more stringent
                        LS = NewCalculation[followerRow][LSColumn] - NewCalculation[row][DurationColumn];
                        if (NewCalculation[row][LSColumn] > LS || NewCalculation[row][LSColumn] == "") { NewCalculation[row][LSColumn] = LS;}
                    }
                    else {
                        Calculation_ready = false;
                    }
                  }
            }
        }
      }
 
      //Color critical path, excluding phases
      for (row = 1; row < NewCalculation.length; row++) {
          if (NewCalculation[row][ColorTypeColumn].toString().toLowerCase() != "phase"){
              if (NewCalculation[row][EEColumn] == NewCalculation[row][LEColumn]) {
                NewCalculation[row][CriticalPathColumn] = "Yes"
              }
              else {
                NewCalculation[row][CriticalPathColumn] = "No";
              }
          }
      }

      console.log("planningproperties")
      console.log(planningProperties)
      
     // TODO: check that phases have ident level that is higher then following line


      //Update phases
      var CurrentPhaseIndentLevel
      var Childrow
      for (row = 1; row<NewCalculation.length; row++) {

          if (NewCalculation[row][ColorTypeColumn].toString().toLowerCase() == "phase") {
              CurrentPhaseIndentLevel = planningProperties.value[row][NameColumn].format.indentLevel;
              //correct to values found in child tasks
              Childrow = row + 1
              while (Childrow < NewCalculation.length 
                     && planningProperties.value[Childrow][NameColumn].format.indentLevel > CurrentPhaseIndentLevel){
                  if (NewCalculation[Childrow][ColorTypeColumn].toString().toLowerCase() != "phase"){  //subphases are not yet filled
                      if (NewCalculation[Childrow][ESColumn] < NewCalculation[row][ESColumn] || NewCalculation[row][ESColumn] == "" ){ NewCalculation[row][ESColumn] = NewCalculation[Childrow][ESColumn];}
                      if (NewCalculation[Childrow][LSColumn] < NewCalculation[row][LSColumn] || NewCalculation[row][LSColumn] == "" ){ NewCalculation[row][LSColumn] = NewCalculation[Childrow][LSColumn];}
                      if (NewCalculation[Childrow][EEColumn] > NewCalculation[row][EEColumn] || NewCalculation[row][EEColumn] == "" ){ NewCalculation[row][EEColumn] = NewCalculation[Childrow][EEColumn];}
                      if (NewCalculation[Childrow][LEColumn] > NewCalculation[row][LEColumn] || NewCalculation[row][LEColumn] == "" ){ NewCalculation[row][LEColumn] = NewCalculation[Childrow][LEColumn];}
                  }
                  Childrow = Childrow + 1
              }
              //set duration
              NewCalculation[row][DurationColumn] = NewCalculation[row][EEColumn] - NewCalculation[row][ESColumn];
          }
      }

      console.log("NewCalculation:");
      console.log(NewCalculation); 
      planning.values = NewCalculation;

      //TIME TO START DRAWING 
      setMessage("Drawing the Gantt chart")
      var startingPoint;
      var verticalStart;
      var baseWidth = ((secondGraphColumn.left - firstGraphColumn.left) * 24) / HoursPerBlock //respect the calculation order, do not change brackets
      verticalStart = firstGraphRow.top + 2 ; // + ActiveSheet.Rows(ColumnHeaderRow).Height + 2
      if (PlanningUnit == "Block of x hours") {
        startingPoint = firstGraphColumn.left 
        PlanningStart = 0
      }
      else {
        startingPoint = firstGraphColumn.left + ((PlanningStart % 1) * baseWidth)
      }
      console.log("Startingpoint ="+startingPoint)
      console.log("BaseWidth = "+baseWidth)
      console.log("Verticalstart ="+verticalStart)

      //delete all activities, but keep the user ones
      var shapeName
      for (var shape of shapes.items){
        shapeName = shape._D.toString().toLowerCase()
        if (shapeName.includes("test")
            || shapeName.includes("milestone")
            || shapeName.includes("phase")
            || shapeName.includes("task") 
            || shapeName.includes("activity")
            || shapeName.includes("progress")  
            || shapeName.includes("connector") 
            || shapeName.includes("timeline") 
        ){
          shape.delete();
        }
      }

      //Task, Milestone and Phase activities are drawn first
      var ID;
      var activityDuration;
      var activityStart;
      var shapeColor;
      var shapeStyle;
      var shapeWidth;
      var horizontalOffset;
      var newShape;
      console.log("drawing shapes for tasks, milestones, phases")
      for (row = 1; row < NewCalculation.length; row++) {  
        ID = NewCalculation[row][IDColumn];
        shapeName = NewCalculation[row][NameColumn];
        activityDuration = NewCalculation[row][DurationColumn];
        activityStart = NewCalculation[row][ESColumn];
        shapeColor = getFillColor(row,ColorTypeColumn, planningProperties);
        if (NewCalculation[row][CriticalPathColumn] == "Yes"){
          shapeColor = CriticalPathColor;
        } 
        else {
            if (shapeColor  == "#FFFFFF") { //color it if not colored by user
                shapeColor = "#00FF00"  //RGB(0, 255, 0)
              }
        }
        if (NewCalculation[row][ColorTypeColumn].toString().toLowerCase() == "milestone") {
                shapeStyle  = Excel.GeometricShapeType.diamond;
                shapeWidth = 10; // same width as height
                horizontalOffset = -5;
        }
        else {
                shapeStyle = Excel.GeometricShapeType.rectangle;
                shapeWidth = activityDuration * baseWidth;
                horizontalOffset = 0;
        }
        
        newShape = shapes.addGeometricShape(shapeStyle);
        newShape.left = Math.max(1,horizontalOffset + startingPoint + (activityStart - PlanningStart) * baseWidth);  //max to be sure to have correct values
        newShape.top = verticalStart +2 + ((row-1) * standardRowHeight );
        newShape.height = 10;
        newShape.width = Math.max(shapeWidth+1,1); //TODO: remove +1 from shapeWidth, this is now done to help drawing the connectors
        newShape.name = "activity_"+ID  //also for milestones and phases
        newShape.fill.foregroundColor = shapeColor
        newShape.lineFormat.color = "grey"
        //test to get hover over text
        newShape.altTextDescription = shapeName;//comes in items._A
        //newShape.altTextTitle = shapeName; 
        //newShape.displayName = shapeName;
        
    
        //Progress bar or only color
        var Progress = NewCalculation[row][ProgressColumn];
        if (Progress != "" && Progress != 0) {
            if (isNaN(Progress)){
                setMessage("Progress should be an integer between 0 and 100. Instead we found '" + Progress + "' on row " + row + ". Please correct. Calculation halted.")
                return;
            }
            Progress = Number(Progress)
            if (NewCalculation[row][ColorTypeColumn].toString().toLowerCase() == "milestone") {
                if (NewCalculation[row][ProgressColumn]==100){
                  newShape.fill.foregroundColor = darken(shapeColor);//"#000000" //darken_color(activity_color)
                  newShape.lineFormat.color = "grey"
                }
            }
            else
            {
                //height same, otherwise line will become too fat 
                newShape = shapes.addGeometricShape(shapeStyle);  //rectangle also for progress bar
                newShape.left = Math.max(1,horizontalOffset + startingPoint + (activityStart - PlanningStart) * baseWidth);  //same start, same security
                newShape.top = verticalStart + 2 + ((row-1) * standardRowHeight );
                newShape.height = 10;
                newShape.width = Math.max(1, activityDuration * baseWidth * Progress / 100);
                newShape.name = "progress_" + ID;
                newShape.fill.foregroundColor = darken(shapeColor)  //"#000000" //darken_color(activity_color)
                newShape.lineFormat.color = "grey"
            }
        }
        
        
        //Phase beginning and ends
        if (NewCalculation[row][ColorTypeColumn].toString().toLowerCase() == "phase") {
            newShape = shapes.addGeometricShape(shapeStyle);  //rectangle also for progress bar
            newShape.left = Math.max(1, horizontalOffset + startingPoint + (activityStart - PlanningStart) * baseWidth);  //same start
            newShape.top = verticalStart + 2 + ((row-1) * standardRowHeight );
            newShape.height = 13;
            newShape.width = 1;
            newShape.name = "phase_begin_" + ID;
            newShape.altTextDescription = shapeName;//comes in items._A
            newShape.fill.foregroundColor = "grey"  //"#000000" 
            newShape.lineFormat.color = "grey"
            
            newShape = shapes.addGeometricShape(shapeStyle);  //rectangle also for progress bar
            newShape.left = Math.max(1, horizontalOffset + startingPoint + (activityStart - PlanningStart) * baseWidth + shapeWidth - 1 +1); //TODO: remove the +1 to align with activity extension of 1
            newShape.top = verticalStart + 2 + ((row-1) * standardRowHeight );
            newShape.height = 13;
            newShape.width = 1;
            newShape.name = "phase_end_" + ID;
            newShape.altTextDescription = shapeName;//comes in items._A
            newShape.fill.foregroundColor = "grey" //"#000000" 
            newShape.lineFormat.color = "grey"
        }
      }

      //connectors (only after all activity shapes are drawn)
      var predecessorID;
      var predecessorRow;
      var predecessorType;
      var predecessorName;
      var predecessorShape;
      var endShape;
      for (row = 1; row<NewCalculation.length; row++) {  
        //end point
        ID = NewCalculation[row][IDColumn];
        shapeName = "activity_"+ID //also for phase and milestone
        //if there is a dependancy, then draw connector
        if (NewCalculation[row][DependencyColumn] != "") {
            //for every dependancy
            dependency = NewCalculation[row][DependencyColumn].toString()
            dependency = dependency.replaceAll(",", " ")
            dependency = dependency.replaceAll("  ", " ")
            dependencies = dependency.split(" ")    //split the words of the string, i.e. space
            for (var j = 0;  j<dependencies.length; j++){
                //origin
                predecessorID = Number.parseInt(dependencies[j],10)
                predecessorRow = getIDRow(predecessorID, NewCalculation);
                predecessorType = NewCalculation[predecessorRow][ColorTypeColumn];
                predecessorName = "activity_"+predecessorID
                //create line
                newShape = shapes.addLine(200, 50, 300, 150, Excel.ConnectorType.curve); //some values to start with
                newShape.name = "Connector_from_"+predecessorID+"_to_"+ID;
                newShape.lineFormat.color = randomColor(); //"tan" //"khaki" //"lavender" //"silver" //"lightgrey" //"grey" //TODO: use randomgrey
                newShape.line.endArrowheadStyle = Excel.ArrowheadStyle.triangle
                newShape.line.endArrowheadWidth = Excel.ArrowheadWidth.medium
                newShape.line.endArrowheadLength = Excel.ArrowheadLength.medium
                if (NewCalculation[row][CriticalPathColumn] == "Yes"){newShape.lineFormat.color=CriticalPathColor}
                //connect to origin & end
                newShape.line.connectBeginShape(shapes.getItem("activity_"+predecessorID), 3);
                newShape.line.connectEndShape(shapes.getItem("activity_"+ID), 1);
                predecessorShape = shapes.getItem("activity_"+predecessorID);
                endShape = shapes.getItem("activity_"+ID);
                //TODO: currently fixed by applying +1 to all activity_width fields
                //if (shapes.getItem("activity_"+predecessorID).left+predecessorShape.width == endShape.width){
                //  predecessorShape.width +=1;
                //}
            }
          
        }
      }

      //Time line
      console.log("drawing timeline")
      if (DrawTimeLine == "Yes"){
        newShape = shapes.addGeometricShape("rectangle");
        newShape.left = startingPoint + (TimeLineValue - PlanningStart) * baseWidth ;  
        newShape.top = verticalStart - (4 * standardRowHeight );
        newShape.height = verticalStart + ((NewCalculation.length - 3)  * standardRowHeight );
        newShape.width = 1;
        newShape.name = "timeline"; 
        newShape.fill.foregroundColor = "orange" 
        newShape.lineFormat.color = "orange"
      }
      console.log("vertical start="+verticalStart+", NewCalculation.length="+NewCalculation.length+", standardRowHeight="+standardRowHeight+", height="+newShape.height)

      //restore visibility of settingsarea
      if (settingsHidden){
        settings.rowHidden = true;
      }
      else {
        settings.rowHidden = false;
      }

    console.log("ending draw")
    setMessage("Drawing complete")
    await context.sync();
    return ;
        
    })
  } 
  catch (error) 
  {
    console.error(error);
    setMessage("Error received, please check your drawing")
  }
}

/*
export  async function searchName(){
  console.log("start searching for 'Name'")
  var column =  await FindColumnHeader("Name");
  console.log("return value is " + column)
}
*/
export  function testshape(headertext){
  console.log("testshape");
  try {
      Excel.run (async (context)=>{
      const ws = context.workbook.worksheets.getActiveWorksheet()
      var shapes = ws.shapes;
      shapes.load("items");
      
      await context.sync();

      let NewShapes = shapes.items
      console.log("NewShapes")
      console.log(NewShapes)
      console.log("No of Shapes at start : "+NewShapes.length)
     console.log("shapes")
     console.log(shapes)
     console.log("no of shapes : "+shapes.items.length)
/*
      for (var i=NewShapes.length; i<=0; i--){
        console.log(shapes.items[i]._D)
        if (shapes.items[i]._D.toString().includes("Test")){shapes.items[i].delete();console.log("--> deleted")}
      }
      console.log(shapes)
      */
      for (var shape of shapes.items){
        if (shape._D.toString().toLowerCase().includes("test")){
          console.log(shape)
          shape.delete();
        }
      }
      await context.sync();
      console.log(shapes);
      console.log("no of shapes after delete : "+shapes.items.length)
      var newShape;

      newShape = shapes.addGeometricShape(Excel.GeometricShapeType.rectangle)
      newShape.left = 200;
      newShape.top = 100;
      newShape.height = 16;
      newShape.width = 32;
      newShape.name = "Test shape 3";
      newShape.fill.Color="red"
   
      
      newShape = shapes.addGeometricShape(Excel.GeometricShapeType.rectangle, 1100  , 1200, 1300, 1400, )
      newShape.left = 216;
      newShape.top = 116;
      newShape.height = 16;
      newShape.width = 96;
      newShape.name = "Test shape 4";
      


      await context.sync();
      console.log(shapes);
      console.log("no of shapes after add and sync: "+shapes.items.length)
      console.log("drawing done")
      


      return ;
    })
  } catch (error) {
  console.error(error);
  }
  console.log("async run fired")
}



export  function timeline(){
  try {
      Excel.run (async (context)=>{
        console.log("timeline_day");
        const ws = context.workbook.worksheets.getActiveWorksheet()    
        const range = ws.getRange("t:azz");
        range.conditionalFormats.clearAll();
        const firstGridColumn = ws.getRange("T:T");
        firstGridColumn.load("format/columnWidth");
        const IDRange = ws.getRange("A1:A9999").getUsedRange()
        IDRange.load("rowCount");
        const settings = ws.getRange(settingsRange);
        settings.load("values");
        const settingsProperties = settings.getCellProperties({
          address: true,
          format: {
              fill: {
                  color: true
              },
              font: {
                  color: true
              }
          },
          style: true,
          indentLevel: true
        });
        await context.sync();

        //TODO: columrange limited to what is necessary//save setting values
        var PlanningStart = getSettingVvalue("Planning start", settings,settingsProperties)
        if (PlanningStart == -1){
          setMessage("No settings found for 'Planning start'. Please add label between A1 and S1, and its value below it.");
          return context;
        }
        var PlanningUnit = getSettingVvalue("Planning unit", settings,settingsProperties)
        if (PlanningUnit == -1){
          setMessage("No settings found for 'Planning unit'. Please add label between A1 and S1, and its value below it.");
          return context;
        }
        var HoursPerBlock = getSettingVvalue("Hours per unit", settings,settingsProperties)
        if (HoursPerBlock == -1){
          setMessage("No settings found for 'Hours per unit'. Please add label between A1 and S1, and its value below it.");
          return context;
        }
       
        //format headers
        var timeheaders = ws.getRange("t3:azz6")
        timeheaders.load("values")
        timeheaders.load("numberFormat")
        var timeformats = timeheaders.getCellProperties({
          address: true,
          format: {
              fill: {
                  color: true
              },
              font: {
                  color: true
              }
          },
          style: true,
          indentLevel: true
        });
        //and format grid
        range.format.columnWidth=firstGridColumn.format.columnWidth;
        await context.sync();

        var th = timeheaders.values;
        var format = timeheaders.numberFormat;
        var D = PlanningStart;
        console.log(timeheaders)
        console.log("planning start = "+ PlanningStart)
        console.log("aantal kolommen = "+th[0].length)
        for (var col = 0; col < th[0].length; col++) {  
          //line 3, only day of month
          th[3][col] = D;
          format[3][col]="dd"
          //line 2, only at start and on first of month a more or less complete data
          var JSD = xlSerialToJsDate(D);
          if (JSD.getDate() == 1 || D ==PlanningStart){
            th[2][col] = JSD.toLocaleString('nl-NL', {weekday:"short", day:"numeric", month:"short"});
            format[2][col]="ddd d mmm";
          }
          else{
            th[2][col] = "";
            format[2][col]= null;
          }
          //line 1
          th[1][col] = "";
          format[1][col]=""
          //line 0
          th[0][col]= "";
          format[0][col]=""
          D = D+1;
        }

        timeheaders.values = th;
        timeheaders.numberFormat = format
        console.log(th)
        console.log(format)
        await context.sync();

        //format the grid
        var timegrid = ws.getRange("t3:azz"+(IDRange.rowCount+3));
        console.log("conditional formatting for range t3:azz"+(IDRange.rowCount+3));
        //sat + sunday + month start
        var conditionalFormat = timegrid.conditionalFormats.add(Excel.ConditionalFormatType.custom); //one of "Custom" | "DataBar" | "ColorScale" | "IconSet" | "TopBottom" | "PresetCriteria" | "ContainsText" | "CellValue"
        conditionalFormat.custom.rule.formula = "=AND(WEEKDAY(t$6, 2)>=6,day(t$6)=1)";
        conditionalFormat.custom.format.fill.color = "#F5F5F5"; //RGB(245, 245, 245);
        //conditionalFormat.custom.format.font.color = "green";
        conditionalFormat.custom.format.borders.getItem('EdgeLeft').style = 'Continuous';
        conditionalFormat.custom.format.borders.getItem('EdgeLeft').color = RGB(0, 0, 0)
        //sat + sunday
        var conditionalFormat2 = timegrid.conditionalFormats.add(Excel.ConditionalFormatType.custom); 
        conditionalFormat2.custom.rule.formula = "=WEEKDAY(t$6, 2)>=6";
        conditionalFormat2.custom.format.fill.color = "#F5F5F5" ; //RGB(245, 245, 245);
        conditionalFormat2.custom.format.borders.getItem('EdgeLeft').style = 'Continuous'; 
        conditionalFormat2.custom.format.borders.getItem('EdgeLeft').color= "#E6E6E6"; // RGB(230, 230, 230)
        //month seperator if not already done above
        var conditionalFormat3 = timegrid.conditionalFormats.add(Excel.ConditionalFormatType.custom); //one of "Custom" | "DataBar" | "ColorScale" | "IconSet" | "TopBottom" | "PresetCriteria" | "ContainsText" | "CellValue"
        conditionalFormat3.custom.rule.formula = "=day(t$6)=1";
        //conditionalFormat3.custom.format.fill.color = "#F5F5F5" ; //RGB(245, 245, 245);
        conditionalFormat3.custom.format.borders.style = "Continous" //LineStyle = xlContinous
        conditionalFormat3.custom.format.borders.getItem('EdgeLeft').color= "#000000"; 
        //weekday separator
        var conditionalFormat4 = timegrid.conditionalFormats.add(Excel.ConditionalFormatType.custom); //one of "Custom" | "DataBar" | "ColorScale" | "IconSet" | "TopBottom" | "PresetCriteria" | "ContainsText" | "CellValue"
        conditionalFormat4.custom.rule.formula = "=WEEKDAY(t$6, 2)<6";
        conditionalFormat4.custom.format.borders.style = "Continous" //LineStyle = xlContinous
        conditionalFormat4.custom.format.borders.getItem('EdgeLeft').color= "#E6E6E6"; // RGB(230, 230, 230)
        await context.sync();

      return ;
    })
  } catch (error) {
  console.error(error);
  }
  console.log("async run fired")
}


/*
Sub TimelineDay()
    'Column width from first time column, i.e. column t onwards
    TimeUnitWidth = ActiveSheet.Columns("t").ColumnWidth
    ActiveSheet.Columns("t:azz").ColumnWidth = TimeUnitWidth   '1
    

    

    
End Sub*/