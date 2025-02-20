/**
 * @file Launches the shortcut target PowerShell script with the selected markdown as an argument.
 * @version 0.0.1.23
 */

// #region: header of utils.js
// Constants and variables.

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

/** The command line arguments. */
var Command = (function() {
  var inputCommand = format('"{0}"', WSH.ScriptFullName);
  for (var index = 0; index < WSH.Arguments.Count(); index++) {
    inputCommand += format(' "{0}"', WSH.Arguments(index));
  }
  return inputCommand;
})();

// #endregion

// #region: main
// The main part of the program.

RequestAdminPrivileges();

/** @typedef */
var Package = getPackage();
/** @typedef */
var Param = getParameters();

/** The application execution. */
if (Param.Markdown) {
  // #region: process.js

  // #region: Process type definition

  var Process = (function() {
    /**
     * @constructor
     * @param {number} processId the process identifier.
     */
    function Process(processId) {
      /** The process identifier. */
      this.Id = processId;
    }

    /**
     * Start a specified program with its arguments.
     * @param {ProcessStartup} startInfo the process startup information.
     * @returns {Process} the process identifier.
     */
    Process.Start = function (startInfo) {
      /** @typedef */
      var Win32_Process = GetObject('winmgmts:Win32_Process');
      var createMethod = Win32_Process.Methods_('Create');
      var createMethodParams = createMethod.InParameters;
      startInfo.StartupInfo.ShowWindow = startInfo.WindowStyle;
      var inParams = createMethodParams.SpawnInstance_();
      inParams.CommandLine = startInfo.CommandLine;
      inParams.ProcessStartupInformation = startInfo.StartupInfo;
      var outParams = Win32_Process.ExecMethod_(createMethod.Name, inParams);
      try {
        return new Process(outParams.ProcessId);
      } finally {
        outParams = null;
        inParams.ProcessStartupInformation = null;
        inParams = null;
        createMethodParams = null;
        createMethod = null;
        Win32_Process = null;
        startInfo = null;
      }
    }

    /**
     * Wait for the process exit.
     * @return {number} the process exit code.
     */
    Process.prototype.WaitForExit = function () {
      // The process termination event query.
      var wqlQuery = 'SELECT * FROM Win32_ProcessStopTrace WHERE ProcessName="cmd.exe" AND ProcessId=' + this.Id;
      // Wait for the process to exit.
      /** @typedef */
      var SWbemService = GetObject('winmgmts:');
      var watcher = SWbemService.ExecNotificationQuery(wqlQuery);
      var cmdProcess = watcher.NextEvent();
      try {
        return cmdProcess.ExitStatus;
      } finally {
        cmdProcess = null;
        watcher = null;
        SWbemService = null;
      }
    }

    return Process;
  })();

  // #endregion

  // #region: ProcessStartup type definition

  var ProcessStartup = (function() {
    /** @typedef */
    var Win32_ProcessStartup = null;

    /**
     * @constructor
     * @param {string} commandLine the command line to execute.
     */
    function ProcessStartup(commandLine) {
      Win32_ProcessStartup = GetObject('winmgmts:Win32_ProcessStartup');
      /** The Win32_ProcessStartup instance. */
      this.StartupInfo = Win32_ProcessStartup.SpawnInstance_();
      /** The window style of the starting process. */
      this.WindowStyle = this.StartupInfo.ShowWindow;
      /** The command line to execute. */
      this.CommandLine = commandLine;
    }

    ProcessStartup.prototype.Dispose = function() {
      this.StartupInfo = null;
      Win32_ProcessStartup = null;
    }

    return ProcessStartup;
  })();

  // #endregion

  // #endregion

  // #region: errorLog.js
  // Manage the error log file and content.

  /** @typedef */
  var ErrorLog = {
    /** The error log file path. */
    Path: generateRandomPath('.log'),

    /** Display the content of the error log file in a message box if it is not empty. */
    Read: function () {
      /** @constant */
      var FOR_READING = 1;
      try {
        var txtStream = FileSystemObject.OpenTextFile(this.Path, FOR_READING);
        // Read the error message and remove the ANSI escaped character for red coloring.
        var errorMessage = txtStream.ReadAll().replace(/(\x1B\[31;1m)|(\x1B\[0m)/g, '');
        if (errorMessage.length) {
          popup(errorMessage, POPUP_ERROR);
        }
      } catch (e) { }
      if (txtStream) {
        txtStream.Close();
        txtStream = null;
      }
    },

    /** Delete the error log file. */
    Delete: function () {
      deleteFile(this.Path)
    }
  }

  // #endregion

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
  // #region: setup.js
  // Methods for managing the shortcut menu option: install and uninstall.

  /** @typedef */
  var Setup = (function() {
    var HKCU = 0x80000001;
    var VERB_KEY = 'SOFTWARE\\Classes\\SystemFileAssociations\\.md\\shell\\cthtml';
    var ICON_VALUENAME = 'Icon';

    return {
      /** Configure the shortcut menu in the registry. */
      Set: function () {
        var COMMAND_KEY = VERB_KEY + '\\command';
        var command = format('{0} "{1}" /Markdown:"%1"', WSH.FullName.replace(/\\cscript\.exe$/i, '\\wscript.exe'), WSH.ScriptFullName);
        StdRegProv.CreateKey(HKCU, COMMAND_KEY);
        StdRegProv.SetStringValue(HKCU, COMMAND_KEY, null, command);
        StdRegProv.SetStringValue(HKCU, VERB_KEY, null, 'Convert to &HTML');
      },

      /**
       * Add an icon to the shortcut menu in the registry.
       * @param {string} menuIconPath is the shortcut menu icon file path.
       */
      AddIcon: function (menuIconPath) {
        StdRegProv.SetStringValue(HKCU, VERB_KEY, ICON_VALUENAME, menuIconPath);
      },

      /** Remove the shortcut icon menu. */
      RemoveIcon: function () {
        StdRegProv.DeleteValue(HKCU, VERB_KEY, ICON_VALUENAME);
      },

      /** Remove the shortcut menu by removing the verb key and subkeys. */
      Unset: function () {
        var stdRegProvMethods = StdRegProv.Methods_;
        var enumKeyMethod = stdRegProvMethods('EnumKey');
        var enumKeyMethodParams = enumKeyMethod.InParameters;
        var inParams = enumKeyMethodParams.SpawnInstance_();
        inParams.hDefKey = HKCU;
        // Recursion is used because a key with subkeys cannot be deleted.
        // Recursion helps removing the leaf keys first.
        (function(key) {
          inParams.sSubKeyName = key;
          var outParams = StdRegProv.ExecMethod_(enumKeyMethod.Name, inParams);
          var sNames = outParams.sNames;
          outParams = null;
          if (sNames != null) {
            var sNamesArray = sNames.toArray();
            for (var index = 0; index < sNamesArray.length; index++) {
              arguments.callee(format('{0}\\{1}', key, sNamesArray[index]));
            }
          }
          StdRegProv.DeleteKey(HKCU, key);
        })(VERB_KEY);
        inParams = null;
        enumKeyMethodParams = null;
        enumKeyMethod = null;
        stdRegProvMethods = null;
      }
    }
  })();

  // #endregion

  if (Param.Set) {
    Setup.Set();
    if (Param.NoIcon) {
      Setup.RemoveIcon();
    } else {
      Setup.AddIcon(Package.MenuIconPath);
    }
    createCustomIconLink(Package.MessageBoxLinkPath, 'C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe', format('-NoLogo -NoProfile -NonInteractive -File "{0}"', FileSystemObject.BuildPath(ScriptRoot, 'rsc\\messageBox.ps1')));
  } else if (Param.Unset) {
    Setup.Unset();
    deleteFile(Package.MessageBoxLinkPath);
  }
  quit(0);
}

quit(1);

// #region: howto-admin-code

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

// #endregion

// #endregion

// #region: utils.js
// Utility functions.

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
  WshShell.Run(format('"{0}" """"{1}"""" {2} {3}', Package.MessageBoxLinkPath, messageText.replace(/"/g, "'"), popupButtons, popupType), WINDOW_STYLE_HIDDEN, WAIT_ON_RETURN);
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

// #endregion

// #region: package.js
// Information about the resource files used by the project.

/** Get the package type. */
function getPackage() {
  /** @constant */
  var POWERSHELL_SUBKEY = 'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\pwsh.exe\\';
  /** The project resources directory path. */
  var resourcePath = FileSystemObject.BuildPath(ScriptRoot, 'rsc');

  /** The powershell core runtime path. */
  var pwshExePath = WshShell.RegRead(POWERSHELL_SUBKEY);
  /** The shortcut target powershell script path. */
  var pwshScriptPath = FileSystemObject.BuildPath(resourcePath, 'cvmd2html.ps1');
  var msgBoxLinkPath = FileSystemObject.BuildPath(ScriptRoot, 'MsgBox.lnk');
  var menuIconPath = FileSystemObject.BuildPath(resourcePath, 'menu.ico');

  return {
    /** The shortcut menu icon path. */
    MenuIconPath: menuIconPath,
    /** The message box link path. */
    MessageBoxLinkPath: msgBoxLinkPath,

    /** Represents an adapted link object. */
    IconLink: {
      /** The icon link path. */
      Path: generateRandomPath('.lnk'),

      /**
       * Create the custom icon link file.
       * @param {string} markdownPath is the input markdown file path.
       */
      Create: function (markdownPath) {
        createCustomIconLink(this.Path, pwshExePath, format('-ep Bypass -nop -w Hidden -f "{0}" -Markdown "{1}"', pwshScriptPath, markdownPath));
      },

      /** Delete the custom icon link file. */
      Delete: function () {
        deleteFile(this.Path);
      }
    }
  }
}

// #endregion

// #region: parameters.js
// Parsed parameters.

/**
 * @typedef {object} ParamHash
 * @property {string} Markdown is the selected markdown file path.
 * @property {boolean} Set installs the shortcut menu.
 * @property {boolean} NoIcon installs the shortcut menu without icon.
 * @property {boolean} Unset uninstalls the shortcut menu.
 * @property {boolean} Help shows help.
 */

/** @returns {ParamHash} */
function getParameters() {
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
}

// #endregion