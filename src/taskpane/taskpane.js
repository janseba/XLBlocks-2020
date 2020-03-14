/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    //document.getElementById("sideload-msg").style.display = "none";
    //document.getElementById("app-body").style.display = "flex";
    //document.getElementById("run").onclick = run;
    document.getElementById("validateFormula").onclick = validateFormula;
    document.getElementById("newFormula").onclick = newFormula;
    document.getElementById("changeFormula").onclick = editFormula;
    document.getElementById("cancel").onclick = cancel;
    document.getElementById("ddlFormulas").onchange = formulaSelectionChanged;
    document.getElementById("delete").onclick = deleteFormula;
    document.getElementById("pasteRange").onclick = pasteRange
    document.getElementById("inspectFormula").onclick = inspectFormula
    document.getElementById("clearMessage").onclick = clearMessage
    toggleButton('newFormula',true)
    toggleButton('validateFormula', false)
    toggleButton('cancel', false)
    getExistingFormulas().then(function () {
      if (document.getElementById('ddlFormulas').length > 1) {
        toggleButton('changeFormula', true)
      } else {
        toggleButton('changeFormula', false)
        toggleButton('delete',false)
      }
    }) 
  }
});

export async function deleteFormula() {
  try{
    await Excel.run(async context => {
      var sheets = context.workbook.worksheets
      sheets.load('items/name')
      await context.sync();
      if (sheetExists(sheets.items, 'XLBlocks')) {
        var sht = sheets.getItem('XLBlocks')
        var rngDefinitions = sht.getUsedRange();
        rngDefinitions.load('values');
        await context.sync();
        if (typeof rngDefinitions !== 'undefined') {
          var xlValues = rngDefinitions.values
          var ddlFormulas = document.getElementById('ddlFormulas')
          var formulaIds = getCol(xlValues,0)
          var formulaRowNumber = formulaIds.indexOf(ddlFormulas.value)
          var formulaRow = rngDefinitions.getRow(formulaRowNumber)
          formulaRow.delete("Up")
          await context.sync();
          xlValues.splice(formulaRowNumber,1);
          replaceFormulaDdl(xlValues);
        }
      }
    })
  } catch(error) {
    console.log(error)
  }
}

export function formulaSelectionChanged() {
  if (document.getElementById('ddlFormulas').value != 'Create a formula...') {
    toggleButton('delete', true)
    toggleButton('changeFormula', true)
  } else {
    toggleButton('delete', false)
    toggleButton('changeFormula', false)
  }
}

export async function editFormula() {
  try {
    toggleButton('validateFormula',true)
    toggleButton('newFormula', false)
    toggleButton('cancel', true)
    toggleButton('changeFormula', false)
    await Excel.run(async context => {
      var sheets = context.workbook.worksheets
      sheets.load('items/name')
      await context.sync();
      if (sheetExists(sheets.items, 'XLBlocks')) {
        var sht = sheets.getItem('XLBlocks')
        var rngDefinitions = sht.getUsedRange();
        rngDefinitions.load('values');
        await context.sync();
        if (typeof rngDefinitions !== 'undefined') {
          var xlValues = rngDefinitions.values
          var ddlFormulas = document.getElementById('ddlFormulas');
          var selectedFormulaID = ddlFormulas.value;
          var ids = getCol(xlValues,0)
          var selectedIndex = ids.indexOf(selectedFormulaID)
          workspace.clear()
          showBlockly();
          var xml_text = xlValues[selectedIndex][2];
          var xml = Blockly.Xml.textToDom(xml_text);
          Blockly.Xml.domToWorkspace(xml, workspace);
        }
      }
    })
  } catch(error) {
    console.log(error)
  }
}

export async function pasteRange() {
  try {
    await Excel.run(async context => {
      var range = context.workbook.getSelectedRange();
      range.load('address')
      await context.sync();
      var selectedAddress = range.address.slice(range.address.indexOf('!') + 1)
      var selectedBlock = Blockly.selected
      if (selectedBlock.type == "range") {
        selectedBlock.getField("range_address").setValue(selectedAddress)
      }
    })
  } catch(error) {
    console.log(error)
  }
}

export async function getExistingFormulas() {
  try {
    await Excel.run(async context => {
      // Get all worksheets
      var sheets = context.workbook.worksheets
      sheets.load('items/name')
      await context.sync(); // get names of sheets
      // Alleen als de worksheet XLBlocks al bestaat zijn er blokdefinities en moet de
      // formula dropdown worden bijgewerkt
      if (sheetExists(sheets.items, 'XLBlocks')) {
        var sht = sheets.getItem('XLBlocks')
        var rngDefinitions = sht.getUsedRange();
        rngDefinitions.load('values');
        await context.sync(); // get values of range
        if (typeof rngDefinitions !== 'undefined') {
          var xlValues = rngDefinitions.values
          //return xlValues
          replaceFormulaDdl(xlValues)
          toggleButton("changeFormula",true)
        } 
      } else {
        buildFormulaDdl();
      }
    })
  } catch(error) {
    console.log(error)
  }
}

export async function validateFormula() {
  try {
    await Excel.run(async context => {
      // Get all worksheets
      var sheets = context.workbook.worksheets
      sheets.load('items/name')

      var rngDefinitions

      await context.sync();
        if (sheetExists(sheets.items, 'XLBlocks')) {
          var sht = sheets.getItem('XLBlocks');
          rngDefinitions = sht.getUsedRange();
          rngDefinitions.load('values');
          sht.delete();
        }

      await context.sync();
      if (typeof rngDefinitions === 'undefined') {
        var xlValues = [];
      } else {
        var xlValues = rngDefinitions.values;
        xlValues.shift();
      }
      var ws = getWorkspace();
      var ids = getCol(xlValues,0)
      var index = ids.indexOf(ws.id);
      if (index === -1) {
        xlValues.push([ws.id, ws.name, ws.fullXML]);
      } else {
        var xml = Blockly.Xml.textToDom(xlValues[index][2])
        var existingOutputRange = getFormulaOutput(xml)
        xlValues[index][1] = ws.name;
        xlValues[index][2] = ws.fullXML;
      }
      xlValues.unshift(['ID', 'Name', 'XML'])
      var sht = sheets.add('XLBlocks');
      var rng = sht.getRange('A1:C' + xlValues.length)
      rng.values = xlValues;
      replaceFormulaDdl(xlValues)
      sht.visibility = Excel.SheetVisibility.hidden;

      // update formulas in Excel
      var code = getFormula()
      var activeSheet = context.workbook.worksheets.getActiveWorksheet();
      if (code.outputRange !== existingOutputRange && existingOutputRange !== undefined) {
        var oldRange = activeSheet.getRange(existingOutputRange);
        oldRange.clear('Contents')
      }
      var formulaRange = activeSheet.getRange(code.outputRange);
      for (var i = 0; i < code.statements.length; i++) {
        for (var j = 0; j < code.statements[i].length; j++) {
          code.statements[i][j] = code.statements[i][j].replace(/\|/g, ',')
        }
      }
      formulaRange.formulas = code.statements;
      await context.sync();
    });
    workspace.clear();
    hideBlockly();
    toggleButton('validateFormula', false);
    toggleButton('newFormula', true);
    toggleButton('cancel', false);
    toggleButton('changeFormula', true);
    toggleButton('delete', true);
  } catch (error) {
    console.log(error);
  }
}

export function hideBlockly() {
  try {
    var blocklyDiv = document.getElementById('blocklyDiv');
    blocklyDiv.style.display = "none";
  } catch (error) {
    console.log(error);
  }
}

export function getFormula() {
  try {
    var formulaBlock = workspace.getBlocksByType('formula', false);
    var code = Blockly.JavaScript.blockToCode(formulaBlock[0]);
    code = JSON.parse(code);
    return code;
  } catch (error) {
    console.log(error);
  }
}

export function replaceFormulaDdl(formulas) {
  try {
  var select = document.getElementById('ddlFormulas');
  select.options.length = 0;
  if (formulas !== null) { 
    for (var i = 1; i < formulas.length; i++) {
      var option = document.createElement('option');
      option.text = formulas[i][1]
      option.value = formulas[i][0]
      select.add(option);
    }
    if (formulas.length == 1) {
      toggleButton('delete', false)
      toggleButton('changeFormula', false)
    }
  }
  buildFormulaDdl();  
  } catch(error) {
    console.log(error);
  }
}

export function getFormulaOutput(xml) {
  try {
    var valueBlocks = xml.getElementsByTagName('value')
    for (var i = 0; i < valueBlocks.length; i++) {
      if (valueBlocks[i].getAttribute('name') == 'output') {
        var fields = valueBlocks[i].getElementsByTagName('field')
        var outputRange = fields[0].textContent
        break
      }
    }

  return outputRange;
  } catch (error) {
    console.log(error);
  }
}

export function getCol(matrix, col) {
  try {
    var column = [];
    for (var i = 0; i < matrix.length; i++) {
      column.push(matrix[i][col]);
    }
    return column
  } catch (error) {
    console.error(error);
  }
}

export function sheetExists(sheets, name) {
  try {
    var result = false;
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].name === name) {
        result = true;
        break;
      }
    }
    return result;
  } catch (error) {
    console.error(error);
  }
}

export function getWorkspace() {
  try {
    var xml = Blockly.Xml.workspaceToDom(workspace, false);
    console.log(Blockly.Xml.domToPrettyText(xml))
    var ws = {id:getFormulaID(xml), name:getFormulaName(xml), fullXML:Blockly.Xml.domToText(xml), outputRange:getFormulaOutput(xml)}   
    return ws;
  } catch (error) {
    console.error(error);
  }
}

export function getFormulaID(xml) {
  try {
    var allBlocks = xml.childNodes
    for (var i = 0; i < allBlocks.length; i++) {
      if (allBlocks[i].getAttribute('type') == 'formula') {
        var id = allBlocks[i].id
        break
      }
    }
  return id;
  } catch (error) {
    console.log(error);
  }
}

function getFormulaName(xml) {
  try {
  var field = xml.getElementsByTagName('field');
  for (var i = 0; i < field.length; i++) {
    if (field[i].getAttribute('name') == 'formula_name') {
      var name = field[i].textContent;
      break;
    }
  }
  return name;
  } catch (error) {
    console.log(error);
  }
}


export function toggleButton(id, isEnabled) {
  try{
    var btn = document.getElementById(id);
    if (isEnabled) {
      btn.disabled = false;
    } else {
      btn.disabled = true
    }
    
  } catch (error) {
    console.error(error);
  }
}

export function newFormula() {
  try {
    showBlockly()
    toggleButton('newFormula', false);
    toggleButton('changeFormula', false)
    toggleButton('validateFormula', true);
    toggleButton('cancel', true);
    toggleButton('delete', false);

  } catch (error) {
    console.error(error);
  }
}

export function showBlockly() {
  try {
    var blocklyDiv = document.getElementById('blocklyDiv');
    if (blocklyDiv.style.display === "none") {
      blocklyDiv.style.display = "block"
      var resizeEvent = window.document.createEvent('UIEvents');
      resizeEvent .initUIEvent('resize', true, false, window, 0);
      window.dispatchEvent(resizeEvent);
    }
  } catch (error) {
    console.log(error);
  }
}

export function buildFormulaDdl() {
  
  var ddlDiv = document.getElementById('formulaDiv');
  var child = ddlDiv.querySelector('.ms-Dropdown-title');
  if (child != null) {
    ddlDiv.removeChild(child);
  }
  child = ddlDiv.querySelector('.ms-Dropdown-items');
  if (child != null) {
    ddlDiv.removeChild(child);
  }
  child = ddlDiv.querySelector('.ms-Dropdown-truncator');
  if (child != null) {
    ddlDiv.removeChild(child);
  }

  var select = document.getElementById('ddlFormulas');
  var option = document.createElement('option');
  option.text = 'Create a formula...'
  option.value = 'Create a formula...'
  select.add(option)

  var DropdownHTMLElement = document.getElementById('formulaDiv');
  var Dropdown = new fabric['Dropdown'](DropdownHTMLElement);
}

export function cancel() {
  workspace.clear();
  hideBlockly();
  var select = document.getElementById('ddlFormulas');
  if (select.length == 1 | select.value == 'Create a formula...') {
    toggleButton('newFormula', true);
    toggleButton('changeFormula', false);
    toggleButton('cancel', false);
    toggleButton('validateFormula', false);
    toggleButton('delete', false) 
  } else {
    toggleButton('newFormula', true);
    toggleButton('changeFormula', true);
    toggleButton('cancel', false);
    toggleButton('validateFormula', false);
    toggleButton('delete', true) 
  }
}

export async function inspectFormula() {
  try{
    await Excel.run(async context => {
      var range = context.workbook.getSelectedRange();
      range.load("formulas");
      range.load("cellCount");
      await context.sync()
      var url = 'https://xlparser.perfectxl.nl/demo/Parse.json?version=120&formula'
      var formula = encodeURI(range.formulas[0]);
      url = url + formula
      fetch(url)
      .then((response) => {
        return response.json();
      })
      .then((data) => {
        console.log(data)
        console.log(xmlSUM(xmlRange('A1:A2')))
      })
    })
    var statusBar = document.getElementById('statusBar')
    statusBar.style.display = "block"
    var message = document.getElementById('messageText')
    message.innerText = 'hallo'
    document.getElementById("clearMessage").onclick=clearMessage
  } catch(error) {
    console.log(error)
  }
}

export function xmlRange(address) {
  return '<block type="range"><field name ="range_address">' + address + '</field></block>'
}

export function xmlSUM(parameters) {
  return '<block type="fn_sum"><value name=sum_parameters">' + parameters + '</value></block>'
}


export function testCreateBlock() {
  try{
    const xmlString = '<xml xmlns="https://developers.google.com/blockly/xml"><block type="formula" x="20" y="20"><field name="formula_name">SUM Test</field><value name="output"><block type="range"><field name="range_address">H5</field></block></value><value name="statements"><block type="fn_sum"><value name="sum_parameters"><block type="range"><field name="range_address">D5:G5</field></block></value></block></value></block></xml>'
    workspace.clear()
    showBlockly()
    var xml = Blockly.Xml.textToDom(xmlString)
    Blockly.Xml.domToWorkspace(xml, workspace)
  } catch(error) {
    console.log(error)
  }
}


export function clearMessage() {
  var statusBar = document.getElementById('statusBar')
  var message = document.getElementById('messageText')
  message.innerText = ""
  statusBar.style.display = "none"
}

