# Literal output paths avoid Start-Process's wildcard interpretation of brackets.
if (-not ('WordPdfChild' -as [type])) {
 Add-Type -TypeDefinition @'
using System;
using System.Diagnostics;
using System.IO;
using System.Text;
public sealed class WordPdfChild {
    readonly Process process;
    StreamWriter output;
    StreamWriter error;
    bool logsClosed;
    volatile string queueResultJson;
    public string QueueResultJson { get { return queueResultJson; } }
    public string LoggingError { get; private set; }
    static StreamWriter OpenLog(string path) {
        return new StreamWriter(new FileStream(path, FileMode.CreateNew,
            FileAccess.Write, FileShare.ReadWrite | FileShare.Delete),
            new UTF8Encoding(false)) { AutoFlush = true };
    }
    public WordPdfChild(string executable, string arguments, string directory,
                        string outputPath, string errorPath) {
        process = new Process();
        try {
            output = OpenLog(outputPath);
            error = OpenLog(errorPath);
            process.StartInfo = new ProcessStartInfo(executable, arguments) {
                UseShellExecute = false, CreateNoWindow = true,
                WorkingDirectory = directory,
                RedirectStandardOutput = true, RedirectStandardError = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8
            };
            process.OutputDataReceived += (sender, e) => {
                if(e.Data != null) {
                    const string prefix="QUEUE_RESULT: ";
                    if(e.Data.StartsWith(prefix,StringComparison.Ordinal)) { queueResultJson=e.Data.Substring(prefix.Length);return; }
                    try { output.WriteLine(e.Data); }
                    catch(Exception ex) { LoggingError = ex.ToString(); }
                }
            };
            process.ErrorDataReceived += (sender, e) => {
                if(e.Data != null) {
                    try { error.WriteLine(e.Data); }
                    catch(Exception ex) { LoggingError = ex.ToString(); }
                }
            };
            if(!process.Start()) throw new IOException("Could not start conversion worker.");
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();
        } catch {
            if(output != null) output.Dispose();
            if(error != null) error.Dispose();
            process.Dispose();
            throw;
        }
    }
    public bool HasExited { get { return process.HasExited; } }
    public int ExitCode { get { return process.ExitCode; } }
    public void WaitForExit() {
        process.WaitForExit(); // Also drains asynchronous stdout/stderr callbacks.
        if(!logsClosed) {
            output.Dispose(); error.Dispose(); logsClosed = true;
        }
    }
}
'@
}
