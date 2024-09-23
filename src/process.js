/**
 * @file mimic the Process and ProcessStartInfo classes in .NET.
 * @version 0.0.1.0
 */

// #region: Process type definition

/**
 * @constructor
 * @param {number} processId the process identifier.
 */
function Process(processId) {
  this.Id = processId;
}

/**
 * Start a specified program with its arguments.
 * @param {ProcessStartup} startInfo the process startup information.
 * @returns {Process} the process identifier.
 */
Process.Start = function (startInfo) {
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

// #endregion

// #region: ProcessStartup type definition

/** @typedef */
var ProcessStartup = (function() {
  var Win32_ProcessStartup = null;

  /**
   * @constructor
   * @param {string} commandLine the command line to execute.
   */
  function ProcessStartup(commandLine) {
    Win32_ProcessStartup = GetObject('winmgmts:Win32_ProcessStartup');
    this.StartupInfo = Win32_ProcessStartup.SpawnInstance_();
    this.WindowStyle = this.StartupInfo.ShowWindow;
    this.CommandLine = commandLine;
  }

  ProcessStartup.prototype.Dispose = function() {
    this.StartupInfo = null;
    Win32_ProcessStartup = null;
  }

  return ProcessStartup;
})();

// #endregion