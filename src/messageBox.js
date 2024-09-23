/**
 * @file message box for the converter.
 * @version 0.0.1.0
 */

/** @constant {number} */
var NO_TIMEOUT = 0;

var WshShell = new ActiveXObject('WScript.Shell');
var WshArguments = WSH.Arguments
WshShell.Popup(WshArguments(0), NO_TIMEOUT, 'Convert to HTML', Number(WshArguments(1)) + Number(WshArguments(2)));
WshShell = null;