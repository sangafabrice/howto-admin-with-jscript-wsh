/**
 * @file Launches the shortcut target PowerShell script with the selected markdown as an argument.
 * @version 0.0.1.0
 */

eval(include('src/utils.js', true));

RequestAdminPrivileges();

/** @type {ParamHash} */
var Param = include('src/parameters.js');

/** The application execution. */
if (Param.Markdown) {
  eval(include('src/process.js', true));
  eval(include('src/errorLog.js', true));
  var CMD_LINE_FORMAT = 'C:\\Windows\\System32\\cmd.exe /d /c ""{0}" 2> "{1}""';
  Package.IconLink.Create(Param.Markdown);
  var startInfo = new ProcessStartup(format(CMD_LINE_FORMAT, Package.IconLink.Path, ErrorLog.Path));
  startInfo.WindowStyle = WINDOW_STYLE_HIDDEN;
  if (Process.Start(startInfo).WaitForExit()) {
    with (ErrorLog) {
      Read();
      Delete();
    }
  }
  Package.IconLink.Delete();
  startInfo.Dispose();
  quit(0);
}

/** Configuration and settings. */
if (Param.Set ^ Param.Unset) {
  eval(include('src/setup.js', true));
  if (Param.Set) {
    Setup.Set();
    if (Param.NoIcon) {
      Setup.RemoveIcon();
    } else {
      Setup.AddIcon(Package.MenuIconPath);
    }
    createCustomIconLink(Package.MessageBoxLinkPath, WSH.FullName.replace(/\\c[^\\]+$/i, '\\wscript.exe'), FileSystemObject.BuildPath(ScriptRoot, 'src\\messageBox.js'));
  } else if (Param.Unset) {
    Setup.Unset();
    deleteFile(Package.MessageBoxLinkPath);
  }
  quit(0);
}

quit(1);

/** Request administrator privileges if standard user. */
function RequestAdminPrivileges() {
  if (IsCurrentProcessElevated()) return;
  var shell = new ActiveXObject('Shell.Application');
  shell.ShellExecute(WSH.FullName, Command, null, 'runas', WINDOW_STYLE_HIDDEN);
  shell = null;
  quit(0);
}

/**
 * Check if the process is elevated.
 * @returns {boolean} True if the running process is elevated, false otherwise.
 */
function IsCurrentProcessElevated() {
  var stdRegProvMethods = StdRegProv.Methods_;
  var checkAccessMethod = stdRegProvMethods('CheckAccess');
  var checkAccessMethodParams = checkAccessMethod.InParameters;
  var inParams = checkAccessMethodParams.SpawnInstance_();
  with (inParams) {
    hDefKey = 0x80000003; // HKU
    sSubKeyName = 'S-1-5-19\\Environment';
  }
  var outParams = StdRegProv.ExecMethod_(checkAccessMethod.Name, inParams);
  try {
    return outParams.bGranted;
  } finally {
    outParams = null;
    inParams = null;
    checkAccessMethodParams = null;
    checkAccessMethod = null;
    stdRegProvMethods = null;
  }
}

// Utility method for importing external JScript code.

/**
 * Import the specified jscript source file.
 * @param {string} libraryPath is the source file path.
 * @param {boolean} isModule[false] specifies that the content is not evaluated.
 * @returns {object} the object returned by the library.
 */
function include(libraryPath, isModule) {
  /** @constant */
  var FOR_READING = 1;
  var fs = new ActiveXObject('Scripting.FileSystemObject');
  try {
    var textStream = fs.OpenTextFile(fs.BuildPath(fs.GetParentFolderName(WSH.ScriptFullName), libraryPath), FOR_READING);
    with(textStream) {
      var content = ReadAll();
      Close();
    }
    if (isModule) {
      return content;
    }
    return eval(content);
  } finally {
    textStream = null;
    fs = null;
  }
}