/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

var originalFormatting = new Object();
var testRanges = new Object();

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
    workspace.addChangeListener(handleBlocklyEvent)
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
      console.log(ws.id)
      console.log(ws.name)
      console.log(ws.fullXML)
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
      range.load("address");
      await context.sync()
      var url = 'https://xlparser.perfectxl.nl/demo/Parse.json?version=120&formula='
      var formula = encodeURIComponent(String(range.formulas[0]).substr(1));
      url = url + formula
      fetch(url)
      .then((response) => {
        return response.json();
      })
      .then((data) => {
        var output = xmlRange(range.address.split("!")[1])
        renderXML(xmlFormula("current cell", output, processFormula(data), 10, 10))
      })
    })
    // var statusBar = document.getElementById('statusBar')
    // statusBar.style.display = "block"
    // var message = document.getElementById('messageText')
    // message.innerText = 'hallo'
    // document.getElementById("clearMessage").onclick=clearMessage
  } catch(error) {
    console.log(error)
  }
}

export function processFormula(formulaTree) {
  if (formulaTree.children[0].name == 'FunctionCall') {
    return processFunctionCall(formulaTree.children[0].children)
  } else if (formulaTree.children[0].name == 'Reference') {
      return processReference(formulaTree.children[0].children)
  } else if (formulaTree.children[0].name == 'Constant') {
      return processConstant(formulaTree.children[0])
  } else if (formulaTree.children[0].name == 'Formula') {
      return processFormula(formulaTree.children[0])
  }
}

export function processConstant(constantTree) {
  if (constantTree.children[0].name == 'Number') {
    return processNumber(constantTree.children[0].children)
  } else if (constantTree.children[0].name == 'Bool') {
    return processBool(constantTree.children[0].children)
  } else if (constantTree.children[0].name == 'Text') {
    return processText(constantTree.children[0].children)
  }
}

export function processNumber(numberTree) {
  var number = numberTree[0].name
  number = number.split('"')[1]
  return xmlNumber(number)
}

export function processBool(boolTree) {
  var bool = boolTree[0].name
  bool = bool.split('"')[1]
  return xmlBool(bool)
}

export function processText(textTree) {
  var text = textTree[0].name
  text = text.split('"')[2]
  return xmlText(text)
}

export function processReference(referenceTree) {
  if (referenceTree[0].name == 'Cell') {
    return processCell(referenceTree[0].children)
  } else if (referenceTree[0].name == 'ReferenceFunctionCall') {
    return processReferenceFunctionCall(referenceTree[0].children)   
  }
}

export function processReferenceFunctionCall(ReferenceFunctionCallTree) {
  if (ReferenceFunctionCallTree[1].name == ':') {
    return processRangeReference(ReferenceFunctionCallTree)
  } else if (ReferenceFunctionCallTree[0].name == 'RefFunctionName')
    return processRefFunctionName(ReferenceFunctionCallTree)
}

export function processRangeReference(ReferenceFunctionCallTree) {
  var leftOperand = ReferenceFunctionCallTree[0].children[0].children[0].name.split('"')[1]
  var operator = ReferenceFunctionCallTree[1].name
  var rightOperand = ReferenceFunctionCallTree[2].children[0].children[0].name.split('"')[1]
  return xmlRange(leftOperand + operator + rightOperand)
}

export function processRefFunctionName(ReferenceFunctionCallTree) {
  if (ReferenceFunctionCallTree[0].children[0].name == 'ExcelConditionalRefFunctionToken["IF("]') {
    var functionArguments = ReferenceFunctionCallTree[1].children
    return processIF(functionArguments)
  }
}

export function processCell(cellTree) {
  var cell = cellTree[0].name.split('"')[1]
  return xmlRange(cell)
}

export function processFunctionCall(functionCallTree) {
  if (functionCallTree[0].name == 'FunctionName') {
    return processFunctionName(functionCallTree)
  } else if (functionCallTree[0].name == 'Formula' && containsWords(functionCallTree[1].name,['+','*','/','=','<','>','-'])) {
    return processBinOp(functionCallTree)
  }
}

export function processBinOp(functionCallTree) {
  var leftOperand = processFormula(functionCallTree[0])
  var operator = functionCallTree[1].name
  var rightOperand = processFormula(functionCallTree[2])
  return xmlBinop(leftOperand, operator, rightOperand)
}

export function processFunctionName(functionCallTree) {
  var functionArguments = functionCallTree[1].children
  if (functionCallTree[0].children[0].name == 'ExcelFunction["VLOOKUP("]') {
    return processVLOOKUP(functionArguments)
  } else if (functionCallTree[0].children[0].name == 'ExcelFunction["ROUND("]') {
    return processROUND(functionArguments)
  } else if (functionCallTree[0].children[0].name == 'ExcelFunction["OR("]') {
    return processOR(functionArguments)
  } else if (functionCallTree[0].children[0].name == 'ExcelFunction["MID("]') {
    return processMID(functionArguments);
  } else if (functionCallTree[0].children[0].name == 'ExcelFunction["AND("]') {
    return processAND(functionArguments);
  } else if (functionCallTree[0].children[0].name == 'ExcelFunction["SUM("]') {
    return processSUM(functionArguments);
  }
}

export function processSUM(functionArguments) {
  var sumArgument = processFormula(functionArguments[0].children[0]);
  return xmlSUM(sumArgument);
}

export function processAND(functionArguments) {
  var andArguments = new Array()
  for (var i = 0; i < functionArguments.length; i++) {
    andArguments[i] = processFormula(functionArguments[i].children[0])
  }
  return xmlAND(andArguments)
}

export function processMID(functionArguments) {
  var midArguments = new Array()
  for (var i = 0; i < functionArguments.length; i++) {
    midArguments[i] = processFormula(functionArguments[i].children[0]);
  }
  return xmlMID(midArguments);
}

export function processOR(functionArguments) {
  var orArguments = new Array()
  for (var i = 0; i < functionArguments.length; i++) {
    orArguments[i] = processFormula(functionArguments[i].children[0])
  }
  return xmlOR(orArguments)
}

export function processROUND(functionArguments) {
  var roundArguments = new Array()
  for (var i = 0; i < functionArguments.length; i++) {
    roundArguments.push(processFormula(functionArguments[i].children[0]))
  }
  return xmlROUND(roundArguments[0],roundArguments[1])
}

export function processVLOOKUP(functionArguments) {
  var vlookupArguments = new Array()
  for (var i = 0; i < functionArguments.length; i++) {
    vlookupArguments.push(processFormula(functionArguments[i].children[0]))
  }
  return xmlVLOOKUP(vlookupArguments[0],vlookupArguments[1],vlookupArguments[2],vlookupArguments[3])
}

export function processIF(functionArguments) {
  var ifArguments = new Array()
  for (var i = 0; i < functionArguments.length; i++) {
     ifArguments.push(processFormula(functionArguments[i].children[0]))
   }
   return xmlIF(ifArguments[0],ifArguments[1],ifArguments[2]) 
}

export function renderXML(xml) {
  try{
    workspace.clear()
    showBlockly()
    xml = Blockly.Xml.textToDom(xml)
    Blockly.Xml.domToWorkspace(xml, workspace)
  } catch(error) {
    console.log(error)
  }
}

export function xmlRange(address) {
  return '<block type="range"><field name="range_address">' + address + '</field></block>'
}

export function xmlSUM(parameters) {
  return '<block type="fn_sum"><value name="sum_parameters">' + parameters + '</value></block>'
}

export function xmlVLOOKUP(lookupValue, tableArray, colIndexNumber, rangeLookup) {
  return  '<block type="fn_vlookup">' +
            '<value name="lookup_value">' + lookupValue + '</value>' +
            '<value name="table_array">' + tableArray + '</value>' +
            '<value name="col_index_num">' + colIndexNumber + '</value>' +
            '<value name="range_lookup">' + rangeLookup + '</value>' +
          '</block>'
}

export function xmlIF(test, when_true, when_false) {
  return  '<block type="fn_if">' +
            '<value name="test">' + test + '</value>' +
            '<value name="when_true">' + when_true + '</value>' +
            '<value name="when_false">' + when_false + '</value>' +
          '</block>'
}

export function xmlROUND(number, num_digits) {
  return  '<block type="fn_round">' +
            '<value name="number">' + number + '</value>' +
            '<value name="num_digits">' + num_digits + '</value>' +
          '</block>'
}

export function xmlBinop(leftOperand, operator, rightOperand) {
  switch(operator) {
    case '+':
      operator = 'add'
      break;
    case '-':
      operator = 'subtract'
      break;
    case '/':
      operator = 'divide'
      break;
    case '>':
      operator = 'gt'
      break;
    case '<':
      operator = 'lt'
      break;
    case '*':
      operator = 'multiply'
      break;
    case '=':
      operator = 'eq'
      break;
  }
  return  '<block type="fn_binop">' +
            '<field name="operator">' +  operator + '</field>' +
            '<value name="left_operand">' + leftOperand + '</value>' +
            '<value name="right_operand">' + rightOperand + '</value>' +
          '</block>'  
}

export function xmlNumber(number) {
  try{
    return '<block type="c_number"><field name="number">' + number + "</field></block>"
  } catch(error) {
    console.log(error)
  }
}

export function xmlBool(bool) {
  try{
    return '<block type="c_bool"><field name="true_false">' + bool + "</field></block>"
  } catch(error) {
    console.log(error)
  }
}

export function xmlText(text) {
  try{
    return '<block type="c_text"><field name="text">' + text + "</field></block>"   
  } catch(error) {
    console.log(error)
  }
}

export function xmlOR(orArguments) {
  var xml = '<block type="fn_or"><statement name="logic_conditions">'
  for (var i = 0; i < orArguments.length; i++) {
    orArguments[i] = orArguments[i].replace('fn_binop', 'fn_logic_condition')
    orArguments[i] = orArguments[i].replace('left_operand', 'left_condition')
    orArguments[i] = orArguments[i].replace('right_operand', 'right_condition')
    if (i < orArguments.length -1) {
      orArguments[i] = orArguments[i].substring(0,orArguments[i].lastIndexOf('<')) + '<next>'
    }
  }
  xml += orArguments.join()
  for (var i = 0; i < orArguments.length -1; i++) {
    xml += '</next></block>'
  }
  xml += '</statement></block>'
  return xml
}

export function xmlAND(andArguments) {
  var xml = '<block type="fn_and"><statement name="logic_conditions">'
  for (var i = 0; i < andArguments.length; i++) {
    andArguments[i] = andArguments[i].replace('fn_binop', 'fn_logic_condition')
    andArguments[i] = andArguments[i].replace('left_operand', 'left_condition')
    andArguments[i] = andArguments[i].replace('right_operand', 'right_condition')
    if (i < andArguments.length -1) {
      andArguments[i] = andArguments[i].substring(0,andArguments[i].lastIndexOf('<')) + '<next>'
    }
  }
  xml += andArguments.join()
  for (var i = 0; i < andArguments.length -1; i++) {
    xml += '</next></block>'
  }
  xml += '</statement></block>'
  return xml
}

export function xmlMID(midArguments) {
  var xml = '<block type="fn_mid">' +
              '<value name="text">' +
                midArguments[0] +
              '</value>' +
              '<value name="start_num">' +
                midArguments[1] +
              '</value>' +
              '<value name="num_chars">' +
                midArguments[2] +
              '</value>' +
            '</block>';
  return xml;
}

export function assembleXml(output, statements) {
  for (var i = 0; i < statements.length; i++) {
    if (containsWords(statements[i],["FunctionCall"])) {
      if (containsWords(statements[i],["*"])) {
        output = xmlRange(output.split('!')[1])
        var leftoperand = xmlOperand(statements[i-1])
        var rightoperand = xmlOperand(statements[i+1])
        var functions = xmlMultiply(leftoperand, rightoperand)
        return xmlFormula('test',output, functions, 10, 10)
      }
    }
  }
}

export function xmlOperand(operand) {
  try{
    if (containsWords(operand,["NumberToken"])) {
      return xmlNumber(operand.split('"')[1])
    } else if (containsWords(operand,["CellToken"])) {
      return xmlRange(operand.split('"')[1])
    }
  } catch(error) {
    console.log(error)
  }
}

export function containsWords(inputString, items) {
  var found = false
  for (var i = 0; i < items.length; i++) {
    if (inputString.includes(items[i])) {
      found = true
      break
    }
  }
  return found
}

export function xmlFormula(name, output, functions, x, y) {
  return '<xml xmlns="https://developers.google.com/blockly/xml">' + 
            '<block type="formula" x="' + x + '" y="' + y + '">' +
              '<field name="formula_name">' + name + '</field>' +
              '<value name="output">' + output + '</value>' +
              '<value name="statements">' + functions + '</value>' +
            '</block>' +
         '</xml>'
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

export function visit(obj, str, output){
  for (let key in obj) {
    if (typeof obj[key] === 'object') {
      if (obj.name !== undefined) {
        str += obj.name + "."
      }
      visit(obj[key], str, output)
    } else {
      if (! obj.hasOwnProperty('children')) {
        output.push(str + obj.name)
      }
    }
  }
}
export function handleBlocklyEvent(blocklyEvent) {
  if (blocklyEvent.type == 'ui' && blocklyEvent.element == 'selected') {
    var selectedBlock = workspace.getBlockById(blocklyEvent.newValue)

    // restore cell formats if originalFormatting has been populated
    if (Object.keys(originalFormatting).length !== 0) {
      console.log('restore formats')
    } else {
        saveCurrentFormats(selectedBlock).then(
          function(){
            applyHighlightBorder(selectedBlock)
          })
      }
  }
}

export async function applyHighlightBorder(selectedBlock) {
  await Excel.run(async context => {
    var sheet = context.workbook.worksheets.getActiveWorksheet()
    var children = selectedBlock.getDescendants()
    for (var i = 0; i < children.length; i++) {
      if (children[i].type =='range') {
        var address = children[i].inputList[0].fieldRow[0].value_
        var range = sheet.getRange(address)
        var b = range.format.borders.getItem('EdgeTop')
        b.style = 'Continuous'
        b.color = 'red'
        b = range.format.borders.getItem('EdgeRight')
        b.style = 'Continuous'
        b.color = 'red'
        b = range.format.borders.getItem('EdgeBottom')
        b.style = 'Continuous'
        b.color = 'red'
        b = range.format.borders.getItem('EdgeLeft')
        b.style = 'Continuous'
        b.color = 'red'
        await context.sync()
      }
    }
  })
}

export async function saveCurrentFormats(selectedBlock) {
  await Excel.run(async context => {
    if (selectedBlock !== null) {
      var sheet = context.workbook.worksheets.getActiveWorksheet()
      var children = selectedBlock.getDescendants()
      for (var i = 0; i < children.length; i++) {
        if(children[i].type == 'range') {
          var address = children[i].inputList[0].fieldRow[0].value_
          var range = sheet.getRange(address)
          var borders = range.format.borders
          var options = Excel.CellPropertiesBorderLoadOptions={
            color: true,
            style: true
          }
          var cellProperties = borders.load(options)
          await context.sync()
          originalFormatting[address] = cellProperties
        }
      }
    }
  })
}

export async function highlightCells(selectedBlock) {
  await Excel.run(async context => {
    var sheet = context.workbook.worksheets.getActiveWorksheet()
    // change format
    if (selectedBlock !== null) {
      var children = selectedBlock.getDescendants()
      for (var i = 0; i < children.length; i++) {
        if (children[i].type == 'range') {
          var address = children[i].inputList[0].fieldRow[0].value_
          var range = sheet.getRange(address)
          var cellProperties = Excel.SettableCellProperties = {
            format: {
              fill: {
                color: "F8EAEB"
              }
            }
          }
          range.set(cellProperties)
          await context.sync()
        }
      }
    }
  })
}

export async function restoreFormat() {
  await Excel.run(async context => {
    for (var key in selectedRanges) {
      var sheet = context.workbook.worksheets.getActiveWorksheet()
      var range = sheet.getRange(key)
      await context.sync()
      range.set(selectedRanges[key].value)
      await context.sync()
    }
  })
}

export async function testFormatting() {
  Excel.run(async context => {
    var aanUit = context.workbook.worksheets.getActiveWorksheet().getRange('A1')
    var borderDefinition = context.workbook.worksheets.getActiveWorksheet().getRange('A2')
    aanUit.load('values')
    borderDefinition.load('values')
    await context.sync()
    console.log(aanUit.values[0][0])
    if (aanUit.values[0][0]=='Aan') {
      await setBorder('B2', 'EdgeTop', 'Continuous', 'red')
      await setBorder('B2', 'EdgeBottom', 'Continuous', 'red')
      await setBorder('B2', 'EdgeLeft', 'Continuous', 'red')
      await setBorder('B2', 'EdgeRight', 'Continuous', 'red')
      await setBorder('B3', 'EdgeTop', 'Continuous', 'red')
      await setBorder('B3', 'EdgeBottom', 'Continuous', 'red')
      await setBorder('B3', 'EdgeLeft', 'Continuous', 'red')
      await setBorder('B3', 'EdgeRight', 'Continuous', 'red')
      borderDefinition.values = [[JSON.stringify(testRanges)]]
      await context.sync
    } else {
      var oldBorderDefinitions = JSON.parse(borderDefinition.values[0][0])
      console.log(oldBorderDefinitions)
      const entries = Object.entries(oldBorderDefinitions)
      for (var i = 0; i < entries.length; i++) {
        console.log(entries[i])
        console.log(entries[i][1].EdgeTop.style)
        console.log(entries[i][1].EdgeTop.color)
        console.log(entries[i][0])
        await resetBorder(entries[i][0], 'EdgeTop', entries[i][1].EdgeTop.style, entries[i][1].EdgeTop.color)
      }
    }
  })
}

export async function setBorder(address, border, style, color) {
  await Excel.run(async context => {
    var range = context.workbook.worksheets.getActiveWorksheet().getRange(address)
    var rand = range.format.borders.getItem(border)
    rand.load({color: true, style: true})
    await context.sync()
    var borderFormat = new Object()
    borderFormat.color = rand.color
    borderFormat.style = rand.style
    var borderEdge = testRanges[address]
    if (borderEdge === undefined) {
      borderEdge = new Object()
    }
    borderEdge[border] = borderFormat
    testRanges[address] = borderEdge
    range.format.borders.getItem(border).style = style
    range.format.borders.getItem(border).color = color
    await context.sync()
  })
}

export async function resetBorder(address, border, style, color) {
  await Excel.run(async context => {
    var range = context.workbook.worksheets.getActiveWorksheet().getRange(address)
    console.log (address + ',' + border + ',' + style + ',' + color)
    range.format.borders.getItem(border).color = color
    range.format.borders.getItem(border).style = style
    await context.sync()
  })
}