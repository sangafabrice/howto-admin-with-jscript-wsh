/**
 * @file returns the methods for managing the shortcut menu option: install and uninstall.
 * @version 0.0.1.0
 */

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