/**
 * @file returns information about the resource files used by the project.
 * @version 0.0.1.0
 */

/** @typedef */
var Package = (function() {
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
})();