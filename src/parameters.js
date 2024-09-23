/**
 * @file returns the parsed parameters.
 * @version 0.0.1.0
 */

/**
 * The parameters and arguments.
 * @typedef {object} ParamHash
 * @property {string} Markdown is the selected markdown file path.
 * @property {boolean} Set installs the shortcut menu.
 * @property {boolean} NoIcon installs the shortcut menu without icon.
 * @property {boolean} Unset uninstalls the shortcut menu.
 * @property {boolean} Help shows help.
 */

/** @returns {ParamHash} */
(function() {
  var WshArguments = WSH.Arguments;
  var WshNamed = WshArguments.Named;
  var paramCount = WshArguments.Count();
  if (paramCount == 1) {
    var paramMarkdown = WshNamed('Markdown');
    if (WshNamed.Exists('Markdown') && paramMarkdown && paramMarkdown.length) {
      return {
        Markdown: paramMarkdown
      }
    }
    var param = { Set: WshNamed.Exists('Set') };
    if (param.Set) {
      var noIconParam = WshNamed('Set');
      var isNoIconParam = false;
      param.NoIcon = noIconParam && (isNoIconParam = /^NoIcon$/i.test(noIconParam));
      if (noIconParam == undefined || isNoIconParam) {
        return param;
      }
    }
    param = { Unset: WshNamed.Exists('Unset') };
    if (param.Unset && WshNamed('Unset') == undefined) {
      return param;
    }
    return {
      Markdown: WshArguments(0)
    }
  } else if (paramCount == 0) {
    return {
      Set: true,
      NoIcon: false
    }
  }
  var helpText = '';
  helpText += 'The MarkdownToHtml shortcut launcher.\n';
  helpText += 'It starts the shortcut menu target script in a hidden window.\n\n';
  helpText += 'Syntax:\n';
  helpText += '  Convert-MarkdownToHtml.js /Markdown:<markdown file path>\n';
  helpText += '  Convert-MarkdownToHtml.js [/Set[:NoIcon]]\n';
  helpText += '  Convert-MarkdownToHtml.js /Unset\n';
  helpText += '  Convert-MarkdownToHtml.js /Help\n\n';
  helpText += "<markdown file path>  The selected markdown's file path.\n";
  helpText += '                 Set  Configure the shortcut menu in the registry.\n';
  helpText += '              NoIcon  Specifies that the icon is not configured.\n';
  helpText += '               Unset  Removes the shortcut menu.\n';
  helpText += '                Help  Show the help doc.\n';
  popup(helpText);
  quit(1);
})();