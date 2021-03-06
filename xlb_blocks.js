Blockly.Blocks['formula'] = {
  init: function() {
    this.appendDummyInput()
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField("Formula")
        .appendField(new Blockly.FieldTextInput("formula name"), "formula_name");
    this.appendValueInput("output")
        .setCheck(null)
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField("formula output");
    this.appendValueInput("statements")
        .setCheck(null)
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField("functions");
    this.setColour(20);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['definenamedranges'] = {
  init: function() {
    this.appendDummyInput()
        .appendField("Define Named Ranges");
    this.appendStatementInput("namedRangeDefinition")
        .setCheck(null);
    this.setColour(20);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['range'] = {
  init: function() {
    this.appendDummyInput()
        .appendField(new Blockly.FieldTextInput("range"), "range_address");
    this.setInputsInline(true);
    this.setOutput(true, "range");
    this.setColour(65);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['fn_sum'] = {
  init: function() {
    this.appendDummyInput()
        .appendField("SUM");
    this.appendValueInput("sum_parameters")
        .setCheck(null);
    this.setInputsInline(true);
    this.setOutput(true, null);
    this.setColour(120);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['for_each_row'] = {
  init: function() {
    this.appendValueInput("range_each_row_in_range")
        .setCheck("range")
        .appendField("EACH ROW");
    this.setInputsInline(true);
    this.setOutput(true, "metaRange");
    this.setColour(65);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['for_each_column'] = {
  init: function() {
    this.appendValueInput("range_each_column_in_range")
        .setCheck("range")
        .appendField("EACH COLUMN");
    this.setInputsInline(true);
    this.setOutput(true, "metaRange");
    this.setColour(65);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['lookup'] = {
  init: function() {
    this.appendDummyInput()
        .appendField("LOOKUP");
    this.appendValueInput("lookupValue")
        .setCheck(null)
        .appendField("lookup value");
    this.appendValueInput("lookupColumn")
        .setCheck(null)
        .appendField("lookup range");
    this.appendValueInput("resultColumn")
        .setCheck(null)
        .appendField("result range");
    this.setOutput(true, null);
    this.setColour(120);
 this.setTooltip("lookup value: the value to be found in the lookup range \n lookup range: range containing possilbe lookup values \n result range: range from wich the matching value should be returned.");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['fn_subtract'] = {
  init: function() {
    this.appendValueInput("left_operand")
        .setCheck(null);
    this.appendDummyInput()
        .appendField("-");
    this.appendValueInput("right_operand")
        .setCheck(null);
    this.setInputsInline(true);
    this.setOutput(true, null);
    this.setColour(120);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['fn_divide'] = {
  init: function() {
    this.appendValueInput("numerator")
        .setCheck(null)
        .setAlign(Blockly.ALIGN_CENTRE);
    this.appendDummyInput()
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField("/");
    this.appendValueInput("denominator")
        .setCheck(null);
    this.setInputsInline(true);
    this.setOutput(true, null);
    this.setColour(120);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['fn_if_error'] = {
  init: function() {
    this.appendDummyInput()
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField("IFERROR");
    this.appendValueInput("formula")
        .setCheck(null)
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField("Formula");
    this.appendValueInput("if_error")
        .setCheck(null)
        .appendField("In case of error");
    this.setOutput(true, null);
    this.setColour(120);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['fn_if'] = {
  init: function() {
    this.appendValueInput("test")
        .setCheck(null)
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField("IF");
    this.appendValueInput("when_true")
        .setCheck(null)
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField("THEN");
    this.appendValueInput("when_false")
        .setCheck(null)
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField("ELSE");
    this.setOutput(true, null);
    this.setColour(120);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['c_number'] = {
  init: function() {
    this.appendDummyInput()
        .appendField("number")
        .appendField(new Blockly.FieldNumber(0), "number");
    this.setOutput(true, "number");
    this.setColour(160);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['fn_greater_than'] = {
  init: function() {
    this.appendValueInput("left_operand")
        .setCheck(null);
    this.appendDummyInput()
        .appendField(">");
    this.appendValueInput("right_operand")
        .setCheck(null);
    this.setInputsInline(true);
    this.setOutput(true, null);
    this.setColour(120);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['c_text'] = {
  init: function() {
    this.appendDummyInput()
        .appendField("text")
        .appendField(new Blockly.FieldTextInput(""), "text");
    this.setOutput(true, null);
    this.setColour(160);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['fn_less_than'] = {
  init: function() {
    this.appendValueInput("left_operand")
        .setCheck(null);
    this.appendDummyInput()
        .appendField("<");
    this.appendValueInput("right_operand")
        .setCheck(null);
    this.setInputsInline(true);
    this.setOutput(true, null);
    this.setColour(120);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['fn_sumifs'] = {
  init: function() {
    this.appendDummyInput()
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField("SUMIFS");
    this.appendValueInput("sum_range")
        .setCheck(null)
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField("sum range");
    this.appendStatementInput("filter_statements")
        .setCheck(null)
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField("filters");
    this.setOutput(true, null);
    this.setColour(120);
 this.setTooltip("sum range: the actual cells to sum");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['fn_sumifs_filters'] = {
  init: function() {
    this.appendValueInput("filter_column")
        .setCheck(null)
        .appendField("filter column");
    this.appendValueInput("filter_value")
        .setCheck(null)
        .appendField("filter value");
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(160);
 this.setTooltip("filter column: the range of cells you want evaluated for the particular condition \n filter value: the condition that defines which cells will be added");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['comments'] = {
  init: function() {
    this.appendDummyInput()
        .appendField("Comments")
        .appendField(new Blockly.FieldTextInput(""), "comment_text");
    this.setInputsInline(true);
    this.setColour(230);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['fn_multiply'] = {
  init: function() {
    this.appendValueInput("left_operand")
        .setCheck(null);
    this.appendDummyInput()
        .appendField("*");
    this.appendValueInput("right_operand")
        .setCheck(null);
    this.setInputsInline(true);
    this.setOutput(true, null);
    this.setColour(120);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['fn_add'] = {
  init: function() {
    this.appendValueInput("left_operand")
        .setCheck(null);
    this.appendDummyInput()
        .appendField("+");
    this.appendValueInput("right_operand")
        .setCheck(null);
    this.setInputsInline(true);
    this.setOutput(true, null);
    this.setColour(120);
 this.setTooltip("");
 this.setHelpUrl("");
  }
}

Blockly.Blocks['fn_binop'] = {
  init: function() {
    this.appendValueInput("left_operand")
        .setCheck(null);
    this.appendDummyInput()
        .appendField(new Blockly.FieldDropdown([["*","multiply"], ["/","divide"], ["+","add"], ["-","subtract"], [">","gt"], ["<","lt"], ["=","eq"]]), "operator");
    this.appendValueInput("right_operand")
        .setCheck(null);
    this.setInputsInline(true);
    this.setOutput(true, null);
    this.setColour(120);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['fn_vlookup'] = {
  init: function() {
    this.appendDummyInput()
        .appendField("VLOOKUP");
    this.appendValueInput("lookup_value")
        .setCheck(["metaRange", "range"])
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField("Lookup value");
    this.appendValueInput("table_array")
        .setCheck("range")
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField("Table array");
    this.appendValueInput("col_index_num")
        .setCheck(["range", "number"])
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField("Column Index Nr");
    this.appendValueInput("range_lookup")
        .setCheck(null)
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField("Range Lookup");
    this.setOutput(true, null);
    this.setColour(120);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['c_bool'] = {
  init: function() {
    this.appendDummyInput()
        .appendField("boolean")
        .appendField(new Blockly.FieldDropdown([["TRUE","TRUE"], ["FALSE","FALSE"]]), "true_false");
    this.setInputsInline(true);
    this.setOutput(true, null);
    this.setColour(160);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['fn_round'] = {
  init: function() {
    this.appendDummyInput()
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField("ROUND");
    this.appendValueInput("number")
        .setCheck(["metaRange", "range"])
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField("number");
    this.appendValueInput("num_digits")
        .setCheck(["range", "number"])
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField("number of digits");
    this.setInputsInline(false);
    this.setOutput(true, "function");
    this.setColour(120);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['fn_or'] = {
  init: function() {
    this.appendDummyInput()
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField("OR");
    this.appendStatementInput("logic_conditions")
        .setCheck(null);
    this.setOutput(true, null);
    this.setColour(120);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['fn_logic_condition'] = {
  init: function() {
    this.appendDummyInput()
        .appendField("logic condition");
    this.appendValueInput("left_condition")
        .setCheck(null);
    this.appendDummyInput()
        .appendField(new Blockly.FieldDropdown([["=","eq"], ["<","lt"], [">","gt"], ["<=","le"], [">=","ge"], ["<>","ne"]]), "operator");
    this.appendValueInput("right_condition")
        .setCheck(null);
    this.setInputsInline(true);
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(160);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['fn_mid'] = {
  init: function() {
    this.appendDummyInput()
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField("MID");
    this.appendValueInput("text")
        .setCheck(null)
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField("text");
    this.appendValueInput("start_num")
        .setCheck(null)
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField("start_num");
    this.appendValueInput("num_chars")
        .setCheck(null)
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField("num_chars");
    this.setOutput(true, null);
    this.setColour(120);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};
Blockly.Blocks['fn_and'] = {
  init: function() {
    this.appendDummyInput()
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField("AND");
    this.appendStatementInput("logic_conditions")
        .setCheck(null);
    this.setOutput(true, null);
    this.setColour(120);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};