/**
 * @license
 * Visual Blocks Language
 *
 * Copyright 2012 Google Inc.
 * https://developers.google.com/blockly/
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *   http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * @fileoverview Helper functions for generating JavaScript for blocks.
 * @author fraser@google.com (Neil Fraser)
 */
'use strict';

goog.provide('Blockly.JavaScript');

goog.require('Blockly.Generator');


/**
 * JavaScript code generator.
 * @type {!Blockly.Generator}
 */
Blockly.JavaScript = new Blockly.Generator('JavaScript');

/**
 * List of illegal variable names.
 * This is not intended to be a security feature.  Blockly is 100% client-side,
 * so bypassing this list is trivial.  This is intended to prevent users from
 * accidentally clobbering a built-in object or function.
 * @private
 */
Blockly.JavaScript.addReservedWords(
    'Blockly,' +  // In case JS is evaled in the current window.
    // https://developer.mozilla.org/en/JavaScript/Reference/Reserved_Words
    'break,case,catch,continue,debugger,default,delete,do,else,finally,for,function,if,in,instanceof,new,return,switch,this,throw,try,typeof,var,void,while,with,' +
    'class,enum,export,extends,import,super,implements,interface,let,package,private,protected,public,static,yield,' +
    'const,null,true,false,' +
    // https://developer.mozilla.org/en/JavaScript/Reference/Global_Objects
    'Array,ArrayBuffer,Boolean,Date,decodeURI,decodeURIComponent,encodeURI,encodeURIComponent,Error,eval,EvalError,Float32Array,Float64Array,Function,Infinity,Int16Array,Int32Array,Int8Array,isFinite,isNaN,Iterator,JSON,Math,NaN,Number,Object,parseFloat,parseInt,RangeError,ReferenceError,RegExp,StopIteration,String,SyntaxError,TypeError,Uint16Array,Uint32Array,Uint8Array,Uint8ClampedArray,undefined,uneval,URIError,' +
    // https://developer.mozilla.org/en/DOM/window
    'applicationCache,closed,Components,content,_content,controllers,crypto,defaultStatus,dialogArguments,directories,document,frameElement,frames,fullScreen,globalStorage,history,innerHeight,innerWidth,length,location,locationbar,localStorage,menubar,messageManager,mozAnimationStartTime,mozInnerScreenX,mozInnerScreenY,mozPaintCount,name,navigator,opener,outerHeight,outerWidth,pageXOffset,pageYOffset,parent,performance,personalbar,pkcs11,returnValue,screen,screenX,screenY,scrollbars,scrollMaxX,scrollMaxY,scrollX,scrollY,self,sessionStorage,sidebar,status,statusbar,toolbar,top,URL,window,' +
    'addEventListener,alert,atob,back,blur,btoa,captureEvents,clearImmediate,clearInterval,clearTimeout,close,confirm,disableExternalCapture,dispatchEvent,dump,enableExternalCapture,escape,find,focus,forward,GeckoActiveXObject,getAttention,getAttentionWithCycleCount,getComputedStyle,getSelection,home,matchMedia,maximize,minimize,moveBy,moveTo,mozRequestAnimationFrame,open,openDialog,postMessage,print,prompt,QueryInterface,releaseEvents,removeEventListener,resizeBy,resizeTo,restore,routeEvent,scroll,scrollBy,scrollByLines,scrollByPages,scrollTo,setCursor,setImmediate,setInterval,setResizable,setTimeout,showModalDialog,sizeToContent,stop,unescape,updateCommands,XPCNativeWrapper,XPCSafeJSObjectWrapper,' +
    'onabort,onbeforeunload,onblur,onchange,onclick,onclose,oncontextmenu,ondevicemotion,ondeviceorientation,ondragdrop,onerror,onfocus,onhashchange,onkeydown,onkeypress,onkeyup,onload,onmousedown,onmousemove,onmouseout,onmouseover,onmouseup,onmozbeforepaint,onpaint,onpopstate,onreset,onresize,onscroll,onselect,onsubmit,onunload,onpageshow,onpagehide,' +
    'Image,Option,Worker,' +
    // https://developer.mozilla.org/en/Gecko_DOM_Reference
    'Event,Range,File,FileReader,Blob,BlobBuilder,' +
    'Attr,CDATASection,CharacterData,Comment,console,DocumentFragment,DocumentType,DomConfiguration,DOMError,DOMErrorHandler,DOMException,DOMImplementation,DOMImplementationList,DOMImplementationRegistry,DOMImplementationSource,DOMLocator,DOMObject,DOMString,DOMStringList,DOMTimeStamp,DOMUserData,Entity,EntityReference,MediaQueryList,MediaQueryListListener,NameList,NamedNodeMap,Node,NodeFilter,NodeIterator,NodeList,Notation,Plugin,PluginArray,ProcessingInstruction,SharedWorker,Text,TimeRanges,Treewalker,TypeInfo,UserDataHandler,Worker,WorkerGlobalScope,' +
    'HTMLDocument,HTMLElement,HTMLAnchorElement,HTMLAppletElement,HTMLAudioElement,HTMLAreaElement,HTMLBaseElement,HTMLBaseFontElement,HTMLBodyElement,HTMLBRElement,HTMLButtonElement,HTMLCanvasElement,HTMLDirectoryElement,HTMLDivElement,HTMLDListElement,HTMLEmbedElement,HTMLFieldSetElement,HTMLFontElement,HTMLFormElement,HTMLFrameElement,HTMLFrameSetElement,HTMLHeadElement,HTMLHeadingElement,HTMLHtmlElement,HTMLHRElement,HTMLIFrameElement,HTMLImageElement,HTMLInputElement,HTMLKeygenElement,HTMLLabelElement,HTMLLIElement,HTMLLinkElement,HTMLMapElement,HTMLMenuElement,HTMLMetaElement,HTMLModElement,HTMLObjectElement,HTMLOListElement,HTMLOptGroupElement,HTMLOptionElement,HTMLOutputElement,HTMLParagraphElement,HTMLParamElement,HTMLPreElement,HTMLQuoteElement,HTMLScriptElement,HTMLSelectElement,HTMLSourceElement,HTMLSpanElement,HTMLStyleElement,HTMLTableElement,HTMLTableCaptionElement,HTMLTableCellElement,HTMLTableDataCellElement,HTMLTableHeaderCellElement,HTMLTableColElement,HTMLTableRowElement,HTMLTableSectionElement,HTMLTextAreaElement,HTMLTimeElement,HTMLTitleElement,HTMLTrackElement,HTMLUListElement,HTMLUnknownElement,HTMLVideoElement,' +
    'HTMLCanvasElement,CanvasRenderingContext2D,CanvasGradient,CanvasPattern,TextMetrics,ImageData,CanvasPixelArray,HTMLAudioElement,HTMLVideoElement,NotifyAudioAvailableEvent,HTMLCollection,HTMLAllCollection,HTMLFormControlsCollection,HTMLOptionsCollection,HTMLPropertiesCollection,DOMTokenList,DOMSettableTokenList,DOMStringMap,RadioNodeList,' +
    'SVGDocument,SVGElement,SVGAElement,SVGAltGlyphElement,SVGAltGlyphDefElement,SVGAltGlyphItemElement,SVGAnimationElement,SVGAnimateElement,SVGAnimateColorElement,SVGAnimateMotionElement,SVGAnimateTransformElement,SVGSetElement,SVGCircleElement,SVGClipPathElement,SVGColorProfileElement,SVGCursorElement,SVGDefsElement,SVGDescElement,SVGEllipseElement,SVGFilterElement,SVGFilterPrimitiveStandardAttributes,SVGFEBlendElement,SVGFEColorMatrixElement,SVGFEComponentTransferElement,SVGFECompositeElement,SVGFEConvolveMatrixElement,SVGFEDiffuseLightingElement,SVGFEDisplacementMapElement,SVGFEDistantLightElement,SVGFEFloodElement,SVGFEGaussianBlurElement,SVGFEImageElement,SVGFEMergeElement,SVGFEMergeNodeElement,SVGFEMorphologyElement,SVGFEOffsetElement,SVGFEPointLightElement,SVGFESpecularLightingElement,SVGFESpotLightElement,SVGFETileElement,SVGFETurbulenceElement,SVGComponentTransferFunctionElement,SVGFEFuncRElement,SVGFEFuncGElement,SVGFEFuncBElement,SVGFEFuncAElement,SVGFontElement,SVGFontFaceElement,SVGFontFaceFormatElement,SVGFontFaceNameElement,SVGFontFaceSrcElement,SVGFontFaceUriElement,SVGForeignObjectElement,SVGGElement,SVGGlyphElement,SVGGlyphRefElement,SVGGradientElement,SVGLinearGradientElement,SVGRadialGradientElement,SVGHKernElement,SVGImageElement,SVGLineElement,SVGMarkerElement,SVGMaskElement,SVGMetadataElement,SVGMissingGlyphElement,SVGMPathElement,SVGPathElement,SVGPatternElement,SVGPolylineElement,SVGPolygonElement,SVGRectElement,SVGScriptElement,SVGStopElement,SVGStyleElement,SVGSVGElement,SVGSwitchElement,SVGSymbolElement,SVGTextElement,SVGTextPathElement,SVGTitleElement,SVGTRefElement,SVGTSpanElement,SVGUseElement,SVGViewElement,SVGVKernElement,' +
    'SVGAngle,SVGColor,SVGICCColor,SVGElementInstance,SVGElementInstanceList,SVGLength,SVGLengthList,SVGMatrix,SVGNumber,SVGNumberList,SVGPaint,SVGPoint,SVGPointList,SVGPreserveAspectRatio,SVGRect,SVGStringList,SVGTransform,SVGTransformList,' +
    'SVGAnimatedAngle,SVGAnimatedBoolean,SVGAnimatedEnumeration,SVGAnimatedInteger,SVGAnimatedLength,SVGAnimatedLengthList,SVGAnimatedNumber,SVGAnimatedNumberList,SVGAnimatedPreserveAspectRatio,SVGAnimatedRect,SVGAnimatedString,SVGAnimatedTransformList,' +
    'SVGPathSegList,SVGPathSeg,SVGPathSegArcAbs,SVGPathSegArcRel,SVGPathSegClosePath,SVGPathSegCurvetoCubicAbs,SVGPathSegCurvetoCubicRel,SVGPathSegCurvetoCubicSmoothAbs,SVGPathSegCurvetoCubicSmoothRel,SVGPathSegCurvetoQuadraticAbs,SVGPathSegCurvetoQuadraticRel,SVGPathSegCurvetoQuadraticSmoothAbs,SVGPathSegCurvetoQuadraticSmoothRel,SVGPathSegLinetoAbs,SVGPathSegLinetoHorizontalAbs,SVGPathSegLinetoHorizontalRel,SVGPathSegLinetoRel,SVGPathSegLinetoVerticalAbs,SVGPathSegLinetoVerticalRel,SVGPathSegMovetoAbs,SVGPathSegMovetoRel,ElementTimeControl,TimeEvent,SVGAnimatedPathData,' +
    'SVGAnimatedPoints,SVGColorProfileRule,SVGCSSRule,SVGExternalResourcesRequired,SVGFitToViewBox,SVGLangSpace,SVGLocatable,SVGRenderingIntent,SVGStylable,SVGTests,SVGTextContentElement,SVGTextPositioningElement,SVGTransformable,SVGUnitTypes,SVGURIReference,SVGViewSpec,SVGZoomAndPan');

/**
 * Order of operation ENUMs.
 * https://developer.mozilla.org/en/JavaScript/Reference/Operators/Operator_Precedence
 */
Blockly.JavaScript.ORDER_ATOMIC = 0;           // 0 "" ...
Blockly.JavaScript.ORDER_NEW = 1.1;            // new
Blockly.JavaScript.ORDER_MEMBER = 1.2;         // . []
Blockly.JavaScript.ORDER_FUNCTION_CALL = 2;    // ()
Blockly.JavaScript.ORDER_INCREMENT = 3;        // ++
Blockly.JavaScript.ORDER_DECREMENT = 3;        // --
Blockly.JavaScript.ORDER_BITWISE_NOT = 4.1;    // ~
Blockly.JavaScript.ORDER_UNARY_PLUS = 4.2;     // +
Blockly.JavaScript.ORDER_UNARY_NEGATION = 4.3; // -
Blockly.JavaScript.ORDER_LOGICAL_NOT = 4.4;    // !
Blockly.JavaScript.ORDER_TYPEOF = 4.5;         // typeof
Blockly.JavaScript.ORDER_VOID = 4.6;           // void
Blockly.JavaScript.ORDER_DELETE = 4.7;         // delete
Blockly.JavaScript.ORDER_AWAIT = 4.8;          // await
Blockly.JavaScript.ORDER_EXPONENTIATION = 5.0; // **
Blockly.JavaScript.ORDER_MULTIPLICATION = 5.1; // *
Blockly.JavaScript.ORDER_DIVISION = 5.2;       // /
Blockly.JavaScript.ORDER_MODULUS = 5.3;        // %
Blockly.JavaScript.ORDER_SUBTRACTION = 6.1;    // -
Blockly.JavaScript.ORDER_ADDITION = 6.2;       // +
Blockly.JavaScript.ORDER_BITWISE_SHIFT = 7;    // << >> >>>
Blockly.JavaScript.ORDER_RELATIONAL = 8;       // < <= > >=
Blockly.JavaScript.ORDER_IN = 8;               // in
Blockly.JavaScript.ORDER_INSTANCEOF = 8;       // instanceof
Blockly.JavaScript.ORDER_EQUALITY = 9;         // == != === !==
Blockly.JavaScript.ORDER_BITWISE_AND = 10;     // &
Blockly.JavaScript.ORDER_BITWISE_XOR = 11;     // ^
Blockly.JavaScript.ORDER_BITWISE_OR = 12;      // |
Blockly.JavaScript.ORDER_LOGICAL_AND = 13;     // &&
Blockly.JavaScript.ORDER_LOGICAL_OR = 14;      // ||
Blockly.JavaScript.ORDER_CONDITIONAL = 15;     // ?:
Blockly.JavaScript.ORDER_ASSIGNMENT = 16;      // = += -= **= *= /= %= <<= >>= ...
Blockly.JavaScript.ORDER_YIELD = 17;         // yield
Blockly.JavaScript.ORDER_COMMA = 18;           // ,
Blockly.JavaScript.ORDER_NONE = 99;            // (...)

/**
 * List of outer-inner pairings that do NOT require parentheses.
 * @type {!Array.<!Array.<number>>}
 */
Blockly.JavaScript.ORDER_OVERRIDES = [
  // (foo()).bar -> foo().bar
  // (foo())[0] -> foo()[0]
  [Blockly.JavaScript.ORDER_FUNCTION_CALL, Blockly.JavaScript.ORDER_MEMBER],
  // (foo())() -> foo()()
  [Blockly.JavaScript.ORDER_FUNCTION_CALL, Blockly.JavaScript.ORDER_FUNCTION_CALL],
  // (foo.bar).baz -> foo.bar.baz
  // (foo.bar)[0] -> foo.bar[0]
  // (foo[0]).bar -> foo[0].bar
  // (foo[0])[1] -> foo[0][1]
  [Blockly.JavaScript.ORDER_MEMBER, Blockly.JavaScript.ORDER_MEMBER],
  // (foo.bar)() -> foo.bar()
  // (foo[0])() -> foo[0]()
  [Blockly.JavaScript.ORDER_MEMBER, Blockly.JavaScript.ORDER_FUNCTION_CALL],

  // !(!foo) -> !!foo
  [Blockly.JavaScript.ORDER_LOGICAL_NOT, Blockly.JavaScript.ORDER_LOGICAL_NOT],
  // a * (b * c) -> a * b * c
  [Blockly.JavaScript.ORDER_MULTIPLICATION, Blockly.JavaScript.ORDER_MULTIPLICATION],
  // a + (b + c) -> a + b + c
  [Blockly.JavaScript.ORDER_ADDITION, Blockly.JavaScript.ORDER_ADDITION],
  // a && (b && c) -> a && b && c
  [Blockly.JavaScript.ORDER_LOGICAL_AND, Blockly.JavaScript.ORDER_LOGICAL_AND],
  // a || (b || c) -> a || b || c
  [Blockly.JavaScript.ORDER_LOGICAL_OR, Blockly.JavaScript.ORDER_LOGICAL_OR]
];

/**
 * Initialise the database of variable names.
 * @param {!Blockly.Workspace} workspace Workspace to generate code from.
 */
Blockly.JavaScript.init = function(workspace) {
  // Create a dictionary of definitions to be printed before the code.
  Blockly.JavaScript.definitions_ = Object.create(null);
  // Create a dictionary mapping desired function names in definitions_
  // to actual function names (to avoid collisions with user functions).
  Blockly.JavaScript.functionNames_ = Object.create(null);

  if (!Blockly.JavaScript.variableDB_) {
    Blockly.JavaScript.variableDB_ =
        new Blockly.Names(Blockly.JavaScript.RESERVED_WORDS_);
  } else {
    Blockly.JavaScript.variableDB_.reset();
  }

  Blockly.JavaScript.variableDB_.setVariableMap(workspace.getVariableMap());

  var defvars = [];
  // Add developer variables (not created or named by the user).
  var devVarList = Blockly.Variables.allDeveloperVariables(workspace);
  for (var i = 0; i < devVarList.length; i++) {
    defvars.push(Blockly.JavaScript.variableDB_.getName(devVarList[i],
        Blockly.Names.DEVELOPER_VARIABLE_TYPE));
  }

  // Add user variables, but only ones that are being used.
  var variables = Blockly.Variables.allUsedVarModels(workspace);
  for (var i = 0; i < variables.length; i++) {
    defvars.push(Blockly.JavaScript.variableDB_.getName(variables[i].getId(),
        Blockly.Variables.NAME_TYPE));
  }

  // Declare all of the variables.
  if (defvars.length) {
    Blockly.JavaScript.definitions_['variables'] =
        'var ' + defvars.join(', ') + ';';
  }
};

/**
 * Prepend the generated code with the variable definitions.
 * @param {string} code Generated code.
 * @return {string} Completed code.
 */
Blockly.JavaScript.finish = function(code) {
  // Convert the definitions dictionary into a list.
  var definitions = [];
  for (var name in Blockly.JavaScript.definitions_) {
    definitions.push(Blockly.JavaScript.definitions_[name]);
  }
  // Clean up temporary data.
  delete Blockly.JavaScript.definitions_;
  delete Blockly.JavaScript.functionNames_;
  Blockly.JavaScript.variableDB_.reset();
  return definitions.join('\n\n') + '\n\n\n' + code;
};

/**
 * Naked values are top-level blocks with outputs that aren't plugged into
 * anything.  A trailing semicolon is needed to make this legal.
 * @param {string} line Line of generated code.
 * @return {string} Legal line of code.
 */
Blockly.JavaScript.scrubNakedValue = function(line) {
  return line + ';\n';
};

/**
 * Encode a string as a properly escaped JavaScript string, complete with
 * quotes.
 * @param {string} string Text to encode.
 * @return {string} JavaScript string.
 * @private
 */
Blockly.JavaScript.quote_ = function(string) {
  // Can't use goog.string.quote since Google's style guide recommends
  // JS string literals use single quotes.
  string = string.replace(/\\/g, '\\\\')
                 .replace(/\n/g, '\\\n')
                 .replace(/'/g, '\\\'');
  return '\'' + string + '\'';
};

/**
 * Common tasks for generating JavaScript from blocks.
 * Handles comments for the specified block and any connected value blocks.
 * Calls any statements following this block.
 * @param {!Blockly.Block} block The current block.
 * @param {string} code The JavaScript code created for this block.
 * @param {boolean=} opt_thisOnly True to generate code for only this statement.
 * @return {string} JavaScript code with comments and subsequent blocks added.
 * @private
 */
Blockly.JavaScript.scrub_ = function(block, code, opt_thisOnly) {
  var commentCode = '';
  // Only collect comments for blocks that aren't inline.
  if (!block.outputConnection || !block.outputConnection.targetConnection) {
    // Collect comment for this block.
    var comment = block.getCommentText();
    comment = Blockly.utils.wrap(comment, Blockly.JavaScript.COMMENT_WRAP - 3);
    if (comment) {
      if (block.getProcedureDef) {
        // Use a comment block for function comments.
        commentCode += '/**\n' +
                       Blockly.JavaScript.prefixLines(comment + '\n', ' * ') +
                       ' */\n';
      } else {
        commentCode += Blockly.JavaScript.prefixLines(comment + '\n', '// ');
      }
    }
    // Collect comments for all value arguments.
    // Don't collect comments for nested statements.
    for (var i = 0; i < block.inputList.length; i++) {
      if (block.inputList[i].type == Blockly.INPUT_VALUE) {
        var childBlock = block.inputList[i].connection.targetBlock();
        if (childBlock) {
          var comment = Blockly.JavaScript.allNestedComments(childBlock);
          if (comment) {
            commentCode += Blockly.JavaScript.prefixLines(comment, '// ');
          }
        }
      }
    }
  }
  var nextBlock = block.nextConnection && block.nextConnection.targetBlock();
  var nextCode = opt_thisOnly ? '' : Blockly.JavaScript.blockToCode(nextBlock);
  return commentCode + code + nextCode;
};

/**
 * Gets a property and adjusts the value while taking into account indexing.
 * @param {!Blockly.Block} block The block.
 * @param {string} atId The property ID of the element to get.
 * @param {number=} opt_delta Value to add.
 * @param {boolean=} opt_negate Whether to negate the value.
 * @param {number=} opt_order The highest order acting on this value.
 * @return {string|number}
 */
Blockly.JavaScript.getAdjusted = function(block, atId, opt_delta, opt_negate,
    opt_order) {
  var delta = opt_delta || 0;
  var order = opt_order || Blockly.JavaScript.ORDER_NONE;
  if (block.workspace.options.oneBasedIndex) {
    delta--;
  }
  var defaultAtIndex = block.workspace.options.oneBasedIndex ? '1' : '0';
  if (delta > 0) {
    var at = Blockly.JavaScript.valueToCode(block, atId,
        Blockly.JavaScript.ORDER_ADDITION) || defaultAtIndex;
  } else if (delta < 0) {
    var at = Blockly.JavaScript.valueToCode(block, atId,
        Blockly.JavaScript.ORDER_SUBTRACTION) || defaultAtIndex;
  } else if (opt_negate) {
    var at = Blockly.JavaScript.valueToCode(block, atId,
        Blockly.JavaScript.ORDER_UNARY_NEGATION) || defaultAtIndex;
  } else {
    var at = Blockly.JavaScript.valueToCode(block, atId, order) ||
        defaultAtIndex;
  }

  if (Blockly.isNumber(at)) {
    // If the index is a naked number, adjust it right now.
    at = parseFloat(at) + delta;
    if (opt_negate) {
      at = -at;
    }
  } else {
    // If the index is dynamic, adjust it in code.
    if (delta > 0) {
      at = at + ' + ' + delta;
      var innerOrder = Blockly.JavaScript.ORDER_ADDITION;
    } else if (delta < 0) {
      at = at + ' - ' + -delta;
      var innerOrder = Blockly.JavaScript.ORDER_SUBTRACTION;
    }
    if (opt_negate) {
      if (delta) {
        at = '-(' + at + ')';
      } else {
        at = '-' + at;
      }
      var innerOrder = Blockly.JavaScript.ORDER_UNARY_NEGATION;
    }
    innerOrder = Math.floor(innerOrder);
    order = Math.floor(order);
    if (innerOrder && order >= innerOrder) {
      at = '(' + at + ')';
    }
  }
  return at;
};
Blockly.JavaScript['fn_sum'] = function(block) {
  var sumParameters = Blockly.JavaScript.valueToCode(block, 'sum_parameters', Blockly.JavaScript.ORDER_NONE);
  var parameters = sumParameters.split(',');
  var sumFormulas = new Array();
  for (var index = 0; index < parameters.length; index++) {
    sumFormulas[index] = 'SUM(' + parameters[index] + ')';
  }
  // TODO: Assemble JavaScript into code variable.
  var code = sumFormulas.join();
  return code;
};

Blockly.JavaScript['range'] = function(block) {
  var text_range_address = block.getFieldValue('range_address');
  // TODO: Assemble JavaScript into code variable.
  var code = text_range_address
  // TODO: Change ORDER_NONE to the correct strength.
  return [code, Blockly.JavaScript.ORDER_ATOMIC];
};

Blockly.JavaScript['for_each_row'] = function(block) {
  var range = Blockly.JavaScript.valueToCode(block, 'range_each_row_in_range', Blockly.JavaScript.ORDER_NONE);
  var rangeCorners = range.split(":");
  var topLeft = rangeCorners[0];
  var bottomRight = rangeCorners[1];
  var topLeftRowColumn = topLeft.split(/([0-9]+)/);
  var bottomRightRowColumn = bottomRight.split(/([0-9]+)/);
  var noRows = bottomRightRowColumn[1] - topLeftRowColumn[1] + 1;
  var ranges = new Array();
  if (topLeftRowColumn[0] === bottomRightRowColumn[0]) { // single column
    for (var row = 0; row < noRows; row++) {
      ranges[row] = topLeftRowColumn[0] + (Number(topLeftRowColumn[1]) + row)
    }
  } else {
    for (var row = 0; row < noRows; row++) { // multiple columns
      ranges[row] = topLeftRowColumn[0] + (Number(topLeftRowColumn[1]) + row) + ":" 
        + bottomRightRowColumn[0] + (Number(topLeftRowColumn[1]) + row)
    }
  }
  var code = ranges.join();
  // TODO: Change ORDER_NONE to the correct strength.
  return [code, Blockly.JavaScript.ORDER_ATOMIC];
};
Blockly.JavaScript['for_each_column'] = function(block) {
  var range = Blockly.JavaScript.valueToCode(block, 'range_each_column_in_range', Blockly.JavaScript.ORDER_NONE);
  var rangeCorners = range.split(":");
  var topLeft = rangeCorners[0];
  var bottomRight = rangeCorners[1];
  var topLeftRowColumn = topLeft.split(/([0-9]+)/);
  var bottomRightRowColumn = bottomRight.split(/([0-9]+)/);
  var noColumns = getNoColumns(range);
  var startColNr = getColumnNr(topLeftRowColumn[0])
  var ranges = new Array();
  if (topLeftRowColumn[1] === bottomRightRowColumn[1]) { // single row
    for (var col = 0; col < noColumns; col++) {
      ranges[col] = getColumnCode(startColNr + col) + topLeftRowColumn[1];
    }

  } else { // multiple rows
    for (var col = 0; col < noColumns; col++) {
      ranges[col] = getColumnCode(startColNr + col) + topLeftRowColumn[1] + ":" 
        + getColumnCode(startColNr + col) + bottomRightRowColumn[1]
    }
  }
  var code = ranges.join();
  // TODO: Change ORDER_NONE to the correct strength.
  return [code, Blockly.JavaScript.ORDER_ATOMIC];
};
Blockly.JavaScript['formula'] = function(block) {
  var text_formula_name = block.getFieldValue('formula_name');
  var outputRange = Blockly.JavaScript.valueToCode(block, 'output', Blockly.JavaScript.ORDER_ATOMIC);
  var noRows = getNoRows(outputRange);
  var noColumns = getNoColumns(outputRange);
  var statements = Blockly.JavaScript.statementToCode(block, 'statements').trim();
  var formula = new Object();
  formula.formulaName = text_formula_name;
  statements = statements.split(',');
  formula.statements = [];
  for (var i = 0; i < noRows; i++) {
    var newRow = []
    for (var j = 0; j < noColumns; j++) {
      newRow.push('=' + statements[i * noColumns + j]); 
    }
    formula.statements.push(newRow);
  }

  formula.outputRange = outputRange;
  // TODO: Assemble JavaScript into code variable.
  var code = JSON.stringify(formula)
  return code;
};

Blockly.JavaScript['lookup'] = function(block) {
  var value_lookupvalue = Blockly.JavaScript.valueToCode(block, 'lookupValue', Blockly.JavaScript.ORDER_ATOMIC);
  var value_lookupcolumn = Blockly.JavaScript.valueToCode(block, 'lookupColumn', Blockly.JavaScript.ORDER_ATOMIC);
  var value_resultcolumn = Blockly.JavaScript.valueToCode(block, 'resultColumn', Blockly.JavaScript.ORDER_ATOMIC);
  var lookupValues = value_lookupvalue.split(',');
  var lookupFormulas = new Array();
  for (var i = 0; i < lookupValues.length; i++) {
    lookupFormulas[i] = 'INDEX(' + value_resultcolumn + '|MATCH(' + lookupValues[i] + '|' + value_lookupcolumn + '|0))'
  }
  // TODO: Assemble JavaScript into code variable.
  var code = lookupFormulas.join();
  return code;
};

Blockly.JavaScript['fn_subtract'] = function(block) {
  var value_left_operand = getCode(block, 'left_operand');
  var value_right_operand = getCode(block, 'right_operand');
  var leftOperands = value_left_operand.split(',');
  var rightOperands = value_right_operand.split(',');
  // TODO: Assemble JavaScript into code variable.
  var subtractFormulas = new Array();

  if (leftOperands.length === rightOperands.length || rightOperands.length === 1) {
    for (var i = 0; i < leftOperands.length; i++) {
      if (rightOperands.length > 1) {
        subtractFormulas[i] = leftOperands[i] + '-' + rightOperands[i];  
      } else {
        subtractFormulas[i] = leftOperands[i] + '-' + rightOperands[0];          
      }
      
    }    
  } else {
    return undefined;
  }

  var code = subtractFormulas.join();
  // TODO: Change ORDER_NONE to the correct strength.
  return code;
};

Blockly.JavaScript['fn_divide'] = function(block) {
  var value_numerator = getCode(block,'numerator');
  var value_denominator = getCode(block, 'denominator');
  var numerators = value_numerator.split(',')
  var denominators = value_denominator.split(',')
  // TODO: Assemble JavaScript into code variable.
  var divideFormulas = new Array();
  if (numerators.length === denominators.length) {
    for (var i = 0; i < numerators.length; i++) {
      divideFormulas[i] = numerators[i] + '/' + denominators[i]
    }
  } else {
    return undefined
  }
  var code = divideFormulas.join();
  // TODO: Change ORDER_NONE to the correct strength.
  return code;
};

Blockly.JavaScript['fn_if_error'] = function(block) {
  var value_formula = getCode(block, 'formula');
  var value_if_error = getCode(block, 'if_error');
  var value_if_error = Blockly.JavaScript.valueToCode(block, 'if_error', Blockly.JavaScript.ORDER_NONE);
  var formulas = value_formula.split(',');
  var ifErrorFormulas = new Array();
  // TODO: Assemble JavaScript into code variable.
  for (var i = 0; i < formulas.length; i++) {
    ifErrorFormulas[i] = 'IFERROR(' + formulas[i] + "|" + value_if_error + ')'
  }
  var code = ifErrorFormulas.join();
  // TODO: Change ORDER_NONE to the correct strength.
  return code;
};

Blockly.JavaScript['fn_if'] = function(block) {
  var value_test = getCode(block, 'test');
  var value_when_true = getCode(block, 'when_true');
  var value_when_false = getCode(block, 'when_false');
  var tests = value_test.split(',');
  var whenTrues = value_when_true.split(',');
  var whenFalses = value_when_false.split(',');
  var ifFormulas = new Array();
    for (var i = 0; i < tests.length; i++) {
      if (tests.length === whenTrues.length || whenTrues.length === 1) {
        if (whenTrues.length > 1) {
          var whenTrue = whenTrues[i];
        } else {
          var whenTrue = whenTrues[0];
        }
      } else {
        return undefined
      }
      if (tests.length === whenFalses.length || whenFalses.length === 1) {
        if (whenFalses.length > 1) {
          var whenFalse = whenFalses[i];
        } else {
          var whenFalse = whenFalses[0];
        }
      } else {
        return undefined
      }
      ifFormulas[i] = 'IF(' + tests[i] + '|' + whenTrue + '|' + whenFalse + ')'
    }
  var code = ifFormulas.join();
  return code;
};

Blockly.JavaScript['c_number'] = function(block) {
  var number_number = block.getFieldValue('number');
  // TODO: Assemble JavaScript into code variable.
  var code = number_number;
  // TODO: Change ORDER_NONE to the correct strength.
  return [code, Blockly.JavaScript.ORDER_ATOMIC];
};

Blockly.JavaScript['c_text'] = function(block) {
  var text_text = block.getFieldValue('text');
  // TODO: Assemble JavaScript into code variable.
  var code = '"' + text_text + '"';
  // TODO: Change ORDER_NONE to the correct strength.
  return [code, Blockly.JavaScript.ORDER_ATOMIC];
};

Blockly.JavaScript['fn_greater_than'] = function(block) {
  var value_left_operand = getCode(block, 'left_operand');
  var value_right_operand = getCode(block, 'right_operand');
  var leftOperands = value_left_operand.split(',');
  var rightOperands = value_right_operand.split(',');
  var greaterThanFormulas = new Array();
  if (leftOperands.length === rightOperands.length || rightOperands.length === 1) {
    for (var i = 0; i < leftOperands.length; i++) {
      if (rightOperands.length > 1) {
        greaterThanFormulas[i] = leftOperands[i] + '>' + rightOperands[i];
      } else {
        greaterThanFormulas[i] = leftOperands[i] + '>' + rightOperands[0];
      }
    }
  } else {
    return undefined;
  }
  // TODO: Assemble JavaScript into code variable.
  var code = greaterThanFormulas.join();
  // TODO: Change ORDER_NONE to the correct strength.
  return code;
};

Blockly.JavaScript['fn_less_than'] = function(block) {
  var value_left_operand = getCode(block, 'left_operand');
  var value_right_operand = getCode(block, 'right_operand');
  var leftOperands = value_left_operand.split(',');
  var rightOperands = value_right_operand.split(',');
  var lessThanFormulas = new Array();
  if (leftOperands.length === rightOperands.length || rightOperands.length === 1) {
    for (var i = 0; i < leftOperands.length; i++) {
      if (rightOperands.length > 1) {
        lessThanFormulas[i] = leftOperands[i] + '<' + rightOperands[i];
      } else {
        lessThanFormulas[i] = leftOperands[i] + '<' + rightOperands[0];
      }
    }
  } else {
    return undefined;
  }
  // TODO: Assemble JavaScript into code variable.
  var code = lessThanFormulas.join();
  // TODO: Change ORDER_NONE to the correct strength.
  return code;
};

Blockly.JavaScript['fn_sumifs'] = function(block) {
  var value_sum_range = Blockly.JavaScript.valueToCode(block, 'sum_range', Blockly.JavaScript.ORDER_NONE);
  var statements_filter_statements = Blockly.JavaScript.statementToCode(block, 'filter_statements').trim();
  var filterStatements = statements_filter_statements.slice(0,-1).split('#');
  var singleFilters = new Array();
  var eachRowFilters = new Array();
  for (var i = 0; i < filterStatements.length; i++) {
    var filterArray = filterStatements[i].split(',');
    if (filterArray.length > 1) {
      eachRowFilters.push(filterArray);
    } else {
      singleFilters.push(filterArray);
    }
  }
  var sumifsFormulas = new Array();
  for (var i = 0; i < eachRowFilters[0].length; i++) {
    var sumifs = 'SUMIFS(' + value_sum_range
    for (var j = 0; j < singleFilters.length; j++) {
      sumifs += singleFilters[j][0]; 
    }
    for (var j = 0; j < eachRowFilters.length; j++) {
      sumifs += eachRowFilters[j][i]
    }
    sumifs += ')'
    sumifsFormulas.push(sumifs);
  }

  var code = sumifsFormulas.join();
  // TODO: Change ORDER_NONE to the correct strength.
  return code;
};

Blockly.JavaScript['fn_sumifs_filters'] = function(block) {
  var value_filter_column = Blockly.JavaScript.valueToCode(block, 'filter_column', Blockly.JavaScript.ORDER_NONE);
  var value_filter_value = Blockly.JavaScript.valueToCode(block, 'filter_value', Blockly.JavaScript.ORDER_NONE);
  var filterValues = value_filter_value.split(',');
  var filterFormulas = new Array();
  // TODO: Assemble JavaScript into code variable.
  for (var i = 0; i < filterValues.length; i++) {
    filterFormulas[i] = '|' + value_filter_column + '|' + filterValues[i]
  }
  var code = filterFormulas.join() + '#';
  return code;
};

// helper functions

function getCode(block, inputName) {
  var inputBlock = block.getInputTargetBlock(inputName);
  if (inputBlock.type.substring(0, 2) === 'fn') {
    return Blockly.JavaScript.statementToCode(block, inputName).trim();
  } else {
    return Blockly.JavaScript.valueToCode(block, inputName, Blockly.JavaScript.ORDER_NONE);
  }
}

function getColumnNr(ref) {
  var ALPHABET = 'abcdefghijklmnopqrstuvwxyz'.split('');
  var chars = ref.split('');
  var columnNr = 0;
  for (var j = 0; j < chars.length; j++) {
    for (var i = 0; i < ALPHABET.length; i++) {
      if (chars[j].toLowerCase()==ALPHABET[i]) {
        columnNr = columnNr + (i + 1) * Math.pow(26,chars.length - (j + 1));
        break;
      }
    }
  }
  return columnNr
};

function getColumnCode(cNr) {
  var charset = ['a','b','c','d','e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v','w','x','y','z'];
  var code = new Array();
  var stop = 0;
  while (cNr > charset.length && stop < 10) {
    cNr = cNr -1;
    var r = cNr % charset.length;
    cNr = cNr - r;
    code.push(charset[r]);
    cNr = cNr / charset.length;
    stop++
  }
  code.push(charset[cNr - 1]);
  code.reverse();
  return code.join('');
};

function getNoRows(ref) {
  if (ref.includes(':')) {
    var rangeCorners = ref.split(':');
    var topLeftCell = rangeCorners[0].split(/([0-9]+)/);
    var bottomRightCell = rangeCorners[1].split(/([0-9]+)/);
    return bottomRightCell[1] - topLeftCell[1] + 1;
  } else {
    return 1;
  }
}

function getNoColumns(ref) {
  if (ref.includes(':')) {
    var rangeCorners = ref.split(':');
    var topLeftCell = rangeCorners[0].split(/([0-9]+)/);
    var bottomRightCell = rangeCorners[1].split(/([0-9]+)/);
    var startColumn = getColumnNr(topLeftCell[0]);
    var endColumn = getColumnNr(bottomRightCell[0]);
    return endColumn - startColumn + 1;
  } else {
    return 1;
  }
}