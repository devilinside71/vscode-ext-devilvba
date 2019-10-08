/* eslint-disable no-import-assign */
/* eslint-disable no-import-assign */
/* eslint-disable no-empty-function */
/* eslint-disable capitalized-comments */
/* eslint-disable multiline-comment-style */

// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
const vscode = require('vscode');

// this method is called when your extension is activated
// your extension is activated the very first time the command is executed

/**
 * Activate command
 * @param {vscode.ExtensionContext} context
 */
function activate(context) {
  // Use the console to output diagnostic information (console.log) and errors (console.error)
  // This line of code will only be executed once when your extension is activated
  console.log('Your extension "vscode-ext-devilvba" is now active!');

  // Get indent size in space from configuration (settings)
  const indentSize = vscode.workspace
    .getConfiguration()
    .get('vscode-ext-devilvba.indentSize');
  // The command has been defined in the package.json file
  // Now provide the implementation of the command with  registerCommand
  // The commandId parameter must match the command field in package.json

  // Command, format selected text only
  let disposable = vscode.commands.registerCommand(
    'extension.VBAFormatterSelected',
    function() {
      // The code you place here will be executed every time your command is executed

      // #region Keywords

      // prettier-ignore
      var commandsUp = new Array('#Else', '#End If', '#If', 'AppActivate', 'Append',
        'Beep', 'ByRef', 'ByVal', 'Call', 'Case', 'ChDir', 'ChDrive', 'Close',
        'Const', 'Date', 'Debug.Print', 'Declare', 'DeleteSetting', 'Dim', 'Do',
        'Do While', 'Each', 'Else', 'ElseIf', 'End', 'End Function', 'End Property',
        'End Select', 'End Sub', 'Erase', 'Error', 'Exit Do', 'Exit For',
        'Exit Function', 'Exit Property', 'Exit Sub', 'FileCopy', 'For', 'For Each',
        'Function', 'Get', 'GoSub', 'GoTo', 'If', 'Iif', 'Input #', 'Kill', 'Let',
        'Line Input #', 'Load', 'Lock', 'Loop', 'Mac', 'Mid', 'MkDir', 'MsgBox',
        'Name', 'Next', 'On', 'On Error', 'Open', 'Option Base', 'Option Compare',
        'Option Explicit', 'Option Private', 'Output', 'Print #', 'Private',
        'Private Sub', 'Property Get', 'Property Let', 'Property Set', 'Public',
        'Put', 'REM', 'RaiseEvent', 'Randomize', 'ReDim', 'Reset', 'Resume', 'Return',
        'RmDir', 'SaveSetting', 'Seek', 'Select Case', 'SendKeys', 'Set', 'SetAttr',
        'Static', 'Stop', 'Sub', 'Then', 'Time', 'Type', 'Unload', 'Unlock', 'Vba6',
        'Vba7', 'Wait', 'Wend', 'While', 'Width #', 'Win32', 'Win64', 'With', 'Write #');

      // prettier-ignore
      var funcsUp = new Array('Abs', 'Array', 'Asc', 'Atn', 'CBool', 'CByte', 'CCur',
        'CDate', 'CDbl', 'CDec', 'CInt', 'CLng', 'CSng', 'CStr', 'CVErr', 'CVar',
        'Choose', 'Chr', 'Cos', 'CurDir', 'DDB', 'Date', 'DateAdd', 'DateDiff',
        'DatePart', 'DateSerial', 'DateValue', 'Day', 'Dir', 'DirectoryNameoutofPath',
        'Error', 'Exp', 'FV', 'FileAttr', 'FileDateTime', 'FileLen',
        'FileNameOutOfPath', 'Filter', 'Fix', 'Format', 'FormatCurrency',
        'FormatDateTime', 'FormatNumber', 'FormatPercent', 'GetAttr',
        'GetDocumentType', 'HasUnoInterfaces', 'Hex', 'Hour', 'IIf', 'IPmt', 'IRR',
        'InStr', 'InStrRev', 'InputBox', 'InsertNewByName', 'Int', 'IsArray',
        'IsDate', 'IsEmpty', 'IsError', 'IsMissing', 'IsNull', 'IsNumeric',
        'IsObject', 'Join', 'LBound', 'LCase', 'LTrim', 'Left', 'Len', 'LoadLibrary',
        'Log', 'MIRR', 'MacScript', 'Mid', 'Minute', 'Month', 'MonthName', 'MsgBox',
        'NPV', 'NPer', 'Now', 'Oct', 'Open', 'PPmt', 'PV', 'Pmt', 'RTrim', 'Rate',
        'Replace', 'Right', 'Rnd', 'Round', 'SLN', 'SYD', 'Second', 'Sgn', 'Sheets',
        'Sin', 'Space', 'Split', 'Sqr', 'Str', 'StrComp', 'StrConv', 'StrReverse',
        'String', 'Switch', 'Tan', 'Time', 'TimeSerial', 'TimeValue', 'Timer', 'Trim',
        'UBound', 'UCase', 'Val', 'Wait', 'Weekday', 'WeekdayName', 'Worksheets',
        'Year', 'addItem', 'callFunction', 'createEnumeration', 'findSheetIndex',
        'getByName', 'getCellByPosition', 'getCellRangeByName', 'getComponents',
        'getCount', 'getURL', 'hasLocation', 'hasMoreElements', 'loadComponentFromURL',
        'nextElement', 'setActiveSheet');
      // prettier-ignore
      var typesUp = new Array(' As Boolean', ' As Date', ' As Double', ' As Integer',
        ' As Long', ' As Object', ' As String', ' As Variant', ' As WorkBook',
        ' As WorkSheet');
      // prettier-ignore
      var objectsUp = new Array('ActiveSheet', 'ActiveWorkbook', 'BasicLibraries',
        'CurrentController', 'GlobalScope', 'RunAutoMacros', 'StarDesktop', 'ThisComponent');
      // prettier-ignore
      var activityUp = new Array('Activate', 'ActiveSheet', 'Address', 'LockControllers',
        'Name', 'Open', 'ScreenUpdating', 'Select', 'String', 'Value', 'getCurrentSelection');
      // prettier-ignore
      var subobjectsUp = new Array('Cells', 'Range', 'Sheets');

      // #endregion

      var currentIndent = 0;
      var firstCase = false;
      var underscored = true;
      var underscoreCount = 0;

      // Get the active text editor
      let editor = vscode.window.activeTextEditor;

      if (editor) {
        let document = editor.document;
        let selection = editor.selection;

        let selectedText = document.getText(selection);
        const lines = selectedText.split('\n');
        var newLines = '';

        // Format VBA code
        for (let index = 0; index < lines.length; index++) {
          let line = lines[index];
          line = addSpaceToOperators(line);
          line = formatSpecialLine(line);
          line = removeSpaces(line);
          line = formatConstDeclarationLine(line);
          line = formatSubLine(line);
          line = formatFuncLine(line);
          line = formatVBACommand(line);
          line = formatVBAFunction(line);
          line = formatVBAType(line);
          line = formatVBAObject(line);
          line = formatVBAActivity(line);
          line = formatVBASubobject(line);
          line = getIndentedLine(line);
          newLines += line + '\n';
        }
        // Display a message box to the user
        vscode.window.showInformationMessage(
          "Command 'Format selected VBA code' completed"
        );
        editor.edit(editBuilder => {
          editBuilder.replace(selection, newLines);
        });
      }

      /**
       * Create indented line
       * @param  {} line
       */
      function getIndentedLine(line) {
        var ret = getIndent(currentIndent) + line.trim();
        var regex;
        var match;
        // Lines with no indent
        regex = /^\s*(private|public|global|option explicit|end sub|end function|end property|\'#region|\'#endregion)\b/gi;
        match = regex.exec(line);
        if (match !== null) {
          ret = line.trim();
          // console.log('MainDeclare Indent:' + currentIndent + ' ' + line);
        }

        // Lines with no indent but indent after
        regex = /^\s*(private sub|private function|public sub|public function|global sub|global function|sub|function|private property|public property|global property)\b/gi;
        match = regex.exec(line);
        if (match !== null) {
          currentIndent = 1;
          ret = line.trim();
          // console.log('Main Indent:' + currentIndent + ' ' + line);
        }

        // Loop start, indent after
        regex = /^\s*(for|while|with|do while)\b/gi;
        match = regex.exec(line);
        if (match !== null) {
          ret = getIndent(currentIndent) + line.trim();
          currentIndent++;
          // console.log('ForWhile Indent:' + currentIndent + ' ' + line);
        }

        // Select Case, indent after
        regex = /^\s*(select case)\b/gi;
        match = regex.exec(line);
        if (match !== null) {
          ret = getIndent(currentIndent) + line.trim();
          currentIndent++;
          firstCase = true;
          // console.log('Select Case Indent:' + currentIndent + ' ' + firstCase + ' ' + line);
        }

        // Case, indent after, outdent after last Case
        regex = /^\s*(case)\b/gi;
        match = regex.exec(line);
        if (match !== null) {
          currentIndent--;
          // console.log('Case Indent firstCase: ' + firstCase + ' ' + line);
          if (firstCase) {
            currentIndent++;
            firstCase = false;
          }
          // console.log('Case Indent:' + currentIndent + ' ' + line);
          ret = getIndent(currentIndent) + line.trim();
          currentIndent++;
        }

        // End Select, outdent after
        regex = /^\s*(end select)\b/gi;
        match = regex.exec(line);
        if (match !== null) {
          currentIndent -= 2;
          ret = getIndent(currentIndent) + line.trim();
          // console.log('End Select Indent:' + currentIndent + ' ' + line);
        }

        // Then at end of line, indent after
        regex = /\s*(then|#then)\b\s*$/gi;
        match = regex.exec(line);
        if (match !== null) {
          ret = getIndent(currentIndent) + line.trim();
          currentIndent++;
          // console.log('Then Indent:' + currentIndent + ' ' + line);
        }

        // _ continue line, indent after, outdent after last _
        regex = /\s*( _)\b\s*$/gi;
        match = regex.exec(line);
        if (match !== null) {
          ret = getIndent(currentIndent) + line.trim();
          currentIndent++;
          underscoreCount++;
          underscored = true;
          // console.log(underscoreCount);
        } else if (underscored) {
          // console.log('Not undersored ' + line);
          underscored = false;
          currentIndent -= underscoreCount;
          underscoreCount = 0;
          // console.log('New indent:' + currentIndent);
        }

        // End of loop, outdent line
        regex = /^\s*(next|end if|#end if|wend|end with|loop)\b/gi;
        match = regex.exec(line);
        if (match !== null) {
          currentIndent--;
          ret = getIndent(currentIndent) + line.trim();
          // console.log('NextEndif Indent:' + currentIndent + ' ' + line);
        }

        // Else, outdent line and indent after
        regex = /^\s*(else|elseif|#else|#elseif)\b/gi;
        match = regex.exec(line);
        if (match !== null) {
          currentIndent--;
          ret = getIndent(currentIndent) + line.trim();
          currentIndent++;
          // console.log('Else Indent:' + currentIndent + ' ' + line);
        }
        return ret;
      }
      // #region Format keywords

      /**
       * Format special, officially not supported line,
       * i.e. #region
       * @param  {} line
       */
      function formatSpecialLine(line) {
        var ret = line;
        var mainRegexp;
        mainRegexp = /^\s*'\s*#\s*(endregion|region)\b(.*)$/gi;
        match = mainRegexp.exec(line);
        try {
          ret = "'#" + match[1] + ' ' + match[2];
        } catch (error) {}
        return ret;
      }

      /**
       * Format commands
       * @param  {} line
       */
      function formatVBACommand(line) {
        var k;
        var ret = line;
        var regex;
        for (k = 0; k < commandsUp.length; k++) {
          regex = new RegExp('\\b' + commandsUp[k] + '\\b', 'gi');
          ret = ret.replace(regex, commandsUp[k]);
        }
        return ret;
      }

      /**
       * Format functions
       * @param  {} line
       */
      function formatVBAFunction(line) {
        var k;
        var ret = line;
        var regex;
        for (k = 0; k < funcsUp.length; k++) {
          regex = new RegExp('\\b' + funcsUp[k] + '\\s*\\(', 'gi');
          ret = ret.replace(regex, funcsUp[k] + '(');
        }
        return ret;
      }

      /**
       * Format variable and function types
       * @param  {} line
       */
      function formatVBAType(line) {
        var k;
        var ret = line;
        var regex;
        for (k = 0; k < typesUp.length; k++) {
          regex = new RegExp('\\b' + typesUp[k] + '\\b', 'gi');
          ret = ret.replace(regex, typesUp[k]);
        }
        return ret;
      }

      /**
       * Format objects
       * @param  {} line
       */
      function formatVBAObject(line) {
        var k;
        var ret = line;
        var regex;
        for (k = 0; k < objectsUp.length; k++) {
          regex = new RegExp('\\b' + objectsUp[k] + '\\s*\\.', 'gi');
          ret = ret.replace(regex, objectsUp[k] + '.');
        }
        return ret;
      }

      /**
       * Format activities
       * @param  {} line
       */
      function formatVBAActivity(line) {
        var k;
        var ret = line;
        var regex;
        for (k = 0; k < activityUp.length; k++) {
          regex = new RegExp('\\.\\s*' + activityUp[k], 'gi');
          ret = ret.replace(regex, '.' + activityUp[k]);
        }
        return ret;
      }

      /**
       * Format subobjects
       * @param  {} line
       */
      function formatVBASubobject(line) {
        var k;
        var ret = line;
        var regex;
        for (k = 0; k < subobjectsUp.length; k++) {
          regex = new RegExp('\\.\\s*' + subobjectsUp[k] + '\\s*\\(', 'gi');
          ret = ret.replace(regex, '.' + subobjectsUp[k] + '(');
        }
        return ret;
      }

      // #endregion

      // #region Functions and Subs

      /**
       * Format function first line
       * @param  {} line
       */
      function formatFuncLine(line) {
        var ret = line;
        var subret = '';
        var mainRegexp;
        var subRegexp;
        var mainMatch;
        var subMatch;
        var subElems;
        var index;
        var elem;
        mainRegexp = /(private\s*function |global\s*function |function )\s*(.*)\((.*)\)\s*as\s*(\b[a-zA-Z0-9_]+\b)$/gi;
        mainMatch = mainRegexp.exec(line);
        try {
          subElems = mainMatch[3].split(',');
          for (index = 0; index < subElems.length; index++) {
            subRegexp = /(\b[a-zA-Z0-9_]+\b)\s*as\s*(\b[a-zA-Z0-9_]+\b)/gi;
            elem = subElems[index].trim();
            // console.log('E:' + elem);
            subMatch = subRegexp.exec(elem);
            // console.log(elem + ' SM:' + subMatch.length);
            try {
              subret += subMatch[1] + ' As ' + capitalize(subMatch[2]);
              if (index + 1 !== subElems.length) {
                subret += ', ';
              }
            } catch (error) {
              subret += formatOptionalPart(elem);
            }
          }
          ret =
            capitalize(mainMatch[1]) +
            mainMatch[2].trim() +
            '(' +
            subret +
            ') As ' +
            capitalize(mainMatch[4]);
          ret = ret.replace('function', 'Function');
        } catch (error) {}
        return ret;
      }

      /**
       * Format Sub first line
       * @param  {} line
       */
      function formatSubLine(line) {
        var ret = line;
        var subret = '';
        var mainRegexp;
        var subRegexp;
        var mainMatch;
        var subMatch;
        var subElems;
        var index;
        var elem;
        mainRegexp = /(private\s*sub |global\s*sub |sub )\s*(.*)\((.*)\)\s*$/gi;
        mainMatch = mainRegexp.exec(line);
        try {
          subElems = mainMatch[3].split(',');
          for (index = 0; index < subElems.length; index++) {
            subRegexp = /(\b[a-zA-Z0-9_]+\b)\s*as\s*(\b[a-zA-Z0-9_]+\b)/gi;
            elem = subElems[index].trim();
            // console.log('E:' + elem);
            subMatch = subRegexp.exec(elem);
            // console.log(elem + ' SM:' + subMatch.length);
            try {
              subret += subMatch[1] + ' As ' + capitalize(subMatch[2]);
              if (index + 1 !== subElems.length) {
                subret += ', ';
              }
            } catch (error) {
              subret += formatOptionalPart(elem);
            }
          }
          ret =
            capitalize(mainMatch[1]) + mainMatch[2].trim() + '(' + subret + ')';
          ret = ret.replace('sub', 'Sub');
        } catch (error) {}
        return ret;
      }

      /**
       * Format Function and Sub Optional parameters
       * @param  {} line
       */
      function formatOptionalPart(line) {
        var ret = line;
        var mainRegexp;
        var mainMatch;
        // console.log('Try OPTIONAL ' + line);
        mainRegexp = /(\boptional\b)\s*(.*)\s*=\s*(.*)\s*/gi;
        mainMatch = mainRegexp.exec(line);
        try {
          // console.log('MOPT:' + mainMatch.length);
          ret = 'Optional ' + mainMatch[2] + ' = ' + mainMatch[3];
        } catch (error) {
          // console.log('Not OPTIONAL');
        }
        return ret;
      }

      // #endregion

      /**
       * Format Const declaration line
       * @param  {} line
       */
      function formatConstDeclarationLine(line) {
        var ret = line;
        myRegexp = /(private |public )\s*const\s*(.*)\s*as\s*(\b[a-zA-Z0-9_]+\b)\s*=\s*(.*$)/gi;
        match = myRegexp.exec(line);
        try {
          // console.log(capitalize(match[1]));
          ret =
            capitalize(match[1]) +
            'Const ' +
            match[2].trim() +
            ' As ' +
            capitalize(match[3]) +
            ' = ' +
            match[4];
        } catch (error) {
          // console.log(error);
        }
        return ret;
      }

      // #region Format operators

      /**
       * Add space around operators
       * @param  {} line
       */
      function addSpaceToOperators(line) {
        var ret = line;
        ret = ret.replace(/\s*(>|<|=|\+|-|&|\/)\s*/gi, replaceSingleOperator);
        ret = ret.replace(/\s*(>|<|=)\s*(>|<|=)\s*/gi, replaceDoubleOperator);
        return ret;
      }
      function replaceSingleOperator(str, group1) {
        return ' ' + group1 + ' ';
      }
      function replaceDoubleOperator(str, group1, group2) {
        return ' ' + group1 + group2 + ' ';
      }

      // #endregion

      /**
       * Remove extra spaces except within quotation marks
       * @param  {string} lineText
       */
      function removeSpaces(lineText) {
        // eslint-disable-next-line id-length
        var newString = lineText.replace(/([^"]+)|("[^"]+")/g, function(
          $0,
          $1,
          $2
        ) {
          if ($1) {
            return $1.replace(/\s{2,}/g, ' ');
          }
          return $2;
        });
        return newString;
      }

      /**
       * Capitalize string, considering #
       * @param  {} string
       */
      function capitalize(string) {
        var ret;
        if (string.charAt(0) === '#') {
          ret =
            '#' +
            string.charAt(1).toUpperCase() +
            string.slice(2).toLowerCase();
        } else {
          ret = string.charAt(0).toUpperCase() + string.slice(1).toLowerCase();
        }
        return ret;
      }

      /**
       * Create indent with spaces
       * @param  {} num Number of indent
       */
      function getIndent(num) {
        var res = '';
        var indent = '';
        if (currentIndent < 0) {
          currentIndent = 0;
        }
        if (num > 0) {
          for (index = 0; index < indentSize; index++) {
            indent += ' ';
          }
          for (i = 0; i < num; i++) {
            res += indent;
          }
        }
        return res;
      }
    }
  );

  context.subscriptions.push(disposable);
}
exports.activate = activate;

// this method is called when your extension is deactivated
function deactivate() {}

module.exports = {
  activate,
  deactivate,
};
