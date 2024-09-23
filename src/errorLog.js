/**
 * @file manages the error log file and content.
 * @version 0.0.1.0
 */

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