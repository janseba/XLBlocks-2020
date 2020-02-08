'use strict';
function replaceFormulaDdl(formulas) {
	var select = document.getElementById('ddlFormulas');
	select.options.length = 0;
	if (formulas !== null) { 
		for (var i = 1; i < formulas.length; i++) {
			var option = document.createElement('option');
			option.text = formulas[i][1]
			option.value = formulas[i][0]
			select.add(option);
		}
	}
	buildFormulaDdl();	
}

function buildFormulaDdl() {
	
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
	var DropdownHTMLElement = document.getElementById('formulaDiv');
	var Dropdown = new fabric['Dropdown'](DropdownHTMLElement);
}

function sheetExists(sheets, name) {
	var result = false;
	for (var i = 0; i < sheets.length; i++) {
		if (sheets[i].name == name) {
			result = true;
			break;
		}
	}
	return result;
}

function getFormulaID(xml) {
	var block = xml.getElementsByTagName('block');
	for (var i = 0; i < block.length; i++) {
		if (block[i].getAttribute('type') == 'formula') {
			var id = block[i].getAttribute('id');
			break;
		}
	}
	return id;
}
function getFormulaName(xml) {
	var field = xml.getElementsByTagName('field');
	for (var i = 0; i < field.length; i++) {
		if (field[i].getAttribute('name') == 'formula_name') {
			var name = field[i].innerText;
			break;
		}
	}
	return name;
}

function getFormulaOutput(xml) {
	var block = xml.getElementsByTagName('block');
	for (var i = 0; i < block.length; i++) {
		if (block[i].getAttribute('type') == 'formula') {
			var xml = $(block[i])
			var outputRange = xml.find('value[name="output"]').text()
			break;
		}
	}
	return outputRange;
}

function getCol(matrix, col) {
	var column = [];
	for (var i = 0; i < matrix.length; i++) {
		column.push(matrix[i][col]);
	}
	return column
}


function initWorkspace(formulas) {
	var ddlFormulas = document.getElementById('ddlFormulas');
	var selectedFormulaID = ddlFormulas.options(ddlFormulas.selectedIndex).value;
	var ids = getCol(formulas,0)
	var selectedIndex = ids.indexOf(selectedFormulaID)
	workspace.clear();
	var xml_text = formulas[selectedIndex][2];
	var xml = Blockly.Xml.textToDom(xml_text);
	Blockly.Xml.domToWorkspace(xml, workspace);
}

function newFormula() {
	showBlockly()
	toggleButton('changeFormula',true);
	toggleButton('validateFormula', false);
	toggleButton('cancel', false);
}

function showBlockly() {
	var blocklyDiv = document.getElementById('blocklyDiv');
	if (blocklyDiv.style.display === "none") {
		blocklyDiv.style.display = "block"
		var resizeEvent = window.document.createEvent('UIEvents');
		resizeEvent .initUIEvent('resize', true, false, window, 0);
		window.dispatchEvent(resizeEvent);
	};
}

function getFormula() {
	var formulaBlock = workspace.getBlocksByType('formula', false);
	var code = Blockly.JavaScript.blockToCode(formulaBlock[0]);
	code = JSON.parse(code);
	return code;
}

function cancel() {
	clearBlockly();
	hideBlockly();
	toggleButton('newFormula', false);
	toggleButton('changeFormula', false);
	toggleButton('cancel', true);
	toggleButton('validateFormula', true);
}

function clearBlockly() {
	workspace.clear();
}

function hideBlockly() {
	var blocklyDiv = document.getElementById('blocklyDiv');
	blocklyDiv.style.display = "none";
}

function toggleButton(id,isDisabled) {
	var btnAdd = document.getElementById(id);
	btnAdd.disabled = isDisabled;
}

function getWorkspace() {
	var xml = Blockly.Xml.workspaceToDom(workspace);
	var ws = {id:getFormulaID(xml), name:getFormulaName(xml), fullXML:Blockly.Xml.domToText(xml), outputRange:getFormulaOutput(xml)}
	return ws;
}

function copyToClipboard(str) {
	const el = document.createElement('textarea');
	el.value =str;
	el.setAttribute('readonly', '');
	el.style.position ='absolute';
	el.style.left = '-9999px';
	document.body.appendChild(el);
	el.select();
	document.execCommand('copy');
	document.body.removeChild(el);
}