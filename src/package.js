/**
 * @file returns information about the resource files used by the project.
 * @version 0.0.1.0
 */

/** @typedef */
var Package = (function() {
  /** @constant */
  var POWERSHELL_SUBKEY = 'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\pwsh.exe\\';
  // The project resources directory path.
  var resourcePath = FileSystem.BuildPath(ScriptRoot, 'rsc');

  // The powershell core runtime path.
  var pwshExePath = Shell.RegRead(POWERSHELL_SUBKEY);
  // The message box link path.
  var msgBoxLinkPath = FileSystem.BuildPath(ScriptRoot, 'MsgBox.lnk');
  // The shortcut target powershell script path.
  var pwshScriptPath = FileSystem.BuildPath(resourcePath, 'cvmd2html.ps1');
  // The shortcut menu icon path.
  var menuIconPath = FileSystem.BuildPath(resourcePath, 'menu.ico');

  return {
    /** @type {string} */
    MenuIconPath: menuIconPath,
    /** @type {string} */
    MessageBoxLinkPath: msgBoxLinkPath,

    /** Represents an adapted link object. */
    IconLink: {
      /** @type {string} */
      Path: generateRandomPath('.lnk'),

      /**
       * Create the custom icon link file.
       * @param {string} markdownPath is the input markdown file path.
       */
      Create: function (markdownPath) {
        var link = Shell.CreateShortcut(this.Path);
        link.TargetPath = pwshExePath;
        link.Arguments = format('-ep Bypass -nop -w Hidden -f "{0}" -Markdown "{1}"', pwshScriptPath, markdownPath);
        link.IconLocation = menuIconPath;
        link.Save();
        link = null;
      },

      /** Delete the custom icon link file. */
      Delete: function () {
        deleteFile(this.Path);
      }
    }
  }
})();