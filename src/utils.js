/**
 * @file some utility functions.
 * @version 0.0.1.0
 */

/** @constant */
var WINDOW_STYLE_HIDDEN = 0;
/** @constant */
var BUTTONS_OKONLY = 0;
/** @constant */
var POPUP_ERROR = 16;
/** @constant */
var POPUP_NORMAL = 0;

/** @typedef */
var FileSystemObject = new ActiveXObject('Scripting.FileSystemObject');
/** @typedef */
var WshShell = new ActiveXObject('WScript.Shell');
/** @typedef */
var StdRegProv = GetObject('winmgmts:StdRegProv');

var ScriptRoot = FileSystemObject.GetParentFolderName(WSH.ScriptFullName)

eval(include('src/package.js', true));

/** The command line arguments. */
var Command = (function() {
  var inputCommand = format('"{0}"', WSH.ScriptFullName);
  for (var index = 0; index < WSH.Arguments.Count(); index++) {
    inputCommand += format(' "{0}"', WSH.Arguments(index));
  }
  return inputCommand;
})();

/**
 * Generate a random file path.
 * @param {string} extension is the file extension.
 * @returns {string} a random file path.
 */
function generateRandomPath(extension) {
  var typeLib = new ActiveXObject('Scriptlet.TypeLib');
  try {
    return FileSystemObject.BuildPath(WshShell.ExpandEnvironmentStrings('%TEMP%'), typeLib.Guid.substr(1, 36).toLowerCase() + '.tmp' + extension);
  } finally {
    typeLib = null;
  }
}

/**
 * Create a custom icon link.
 * @param {string} linkPath is the link path.
 * @param {string} targetPath is the link target path.
 * @param {string} command is the link target arguments.
 */
function createCustomIconLink(linkPath, targetPath, command) {
  var link = WshShell.CreateShortcut(linkPath);
  link.TargetPath = targetPath;
  link.Arguments = command;
  link.IconLocation = Package.MenuIconPath;
  link.Save();
  link = null;
}

/**
 * Delete the specified file.
 * @param {string} filePath is the file path.
 */
function deleteFile(filePath) {
  try {
    FileSystemObject.DeleteFile(filePath);
  } catch (e) { }
}

/**
 * Show the application message box.
 * @param {string} messageText is the message text to show.
 * @param {number} popupType[POPUP_NORMAL] is the type of popup box.
 * @param {number} popupButtons[BUTTONS_OKONLY] are the buttons of the message box.
 */
function popup(messageText, popupType, popupButtons) {
  /** @constant */
  var WAIT_ON_RETURN = true;
  if (!popupType) {
    popupType = POPUP_NORMAL;
  }
  if (!popupButtons) {
    popupButtons = BUTTONS_OKONLY;
  }
  WshShell.Run(format('"{0}" "{1}" {2} {3}', Package.MessageBoxLinkPath, messageText.replace(/"/g, "'"), popupButtons, popupType), WINDOW_STYLE_HIDDEN, WAIT_ON_RETURN);
}

/**
 * Replace the format item "{n}" by the nth input in a list of arguments.
 * @param {string} formatStr the pattern format.
 * @param {...string} args the replacement texts.
 * @returns {string} a copy of format with the format items replaced by args.
 */
function format(formatStr, args) {
  args = Array.prototype.slice.call(arguments).slice(1);
  while (args.length > 0) {
    formatStr = formatStr.replace(new RegExp('\\{' + (args.length - 1) + '\\}', 'g'), args.pop());
  }
  return formatStr;
}

/** Destroy the COM objects. */
function dispose() {
  StdRegProv = null;
  WshShell = null;
  FileSystemObject = null;
}

/**
 * Clean up and quit.
 * @param {number} exitCode .
 */
function quit(exitCode) {
  dispose();
  WSH.Quit(exitCode);
}