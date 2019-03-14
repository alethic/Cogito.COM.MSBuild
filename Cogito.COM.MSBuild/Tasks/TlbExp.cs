using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

using Microsoft.Build.Framework;

namespace Cogito.COM.MSBuild
{

    public class TlbExp :
        Microsoft.Build.Utilities.Task
    {

        /// <summary>
        /// Path which contains the tlbexp.exe utility.
        /// </summary>
        public string ToolPath { get; set; }

        /// <summary>
        /// Path of source manifest file.
        /// </summary>
        public string Assembly { get; set; }

        /// <summary>
        /// File name of type library to be produced.
        /// </summary>
        public string OutputPath { get; set; }

        /// <summary>
        /// Type library used to resolve references.
        /// </summary>
        public ITaskItem[] TlbReference { get; set; }

        /// <summary>
        /// Path used to resolve referenced type libraries.
        /// </summary>
        public ITaskItem[] TlbRefPath { get; set; }

        /// <summary>
        /// Look for assembly references here.
        /// </summary>
        public ITaskItem[] AsmPath { get; set; }

        /// <summary>
        /// Escapes the argument for .NET command line processing.
        /// </summary>
        /// <param name="arg"></param>
        /// <returns></returns>
        string EscapeArg(string arg)
        {
            if (arg.Contains(" "))
                return arg.EndsWith("\\") ? "\"" + arg + "\\\"" : "\"" + arg + "\"";
            else
                return arg;
        }

        /// <summary>
        /// Gets the unique asmpaths.
        /// </summary>
        /// <returns></returns>
        IEnumerable<string> GetAsmPath()
        {
            var h = new HashSet<string>();
            if (AsmPath != null)
                foreach (var i in AsmPath)
                    if (!string.IsNullOrWhiteSpace(i.ItemSpec) && h.Add(i.ItemSpec) && Directory.Exists(i.ItemSpec))
                        yield return i.ItemSpec.TrimEnd('\\');
        }

        /// <summary>
        /// Executes the task.
        /// </summary>
        /// <returns></returns>
        public override bool Execute()
        {
            var exec = Path.Combine(ToolPath, "tlbexp.exe");
            var args = new List<string>();

            if (File.Exists(exec) == false)
            {
                Log.LogError("Could not find tlbexp.exe: {0}", exec);
                return false;
            }

            // assembly name argument
            if (!string.IsNullOrEmpty(Assembly))
                args.Add(Assembly);

            if (!string.IsNullOrWhiteSpace(OutputPath))
                args.Add("/out:" + OutputPath);

            if (TlbReference != null)
                foreach (var i in TlbReference)
                    if (!string.IsNullOrWhiteSpace(i.ItemSpec))
                        args.Add("/tlbreference:" + i.ItemSpec);

            if (TlbRefPath != null)
                foreach (var i in TlbRefPath)
                    if (!string.IsNullOrWhiteSpace(i.ItemSpec))
                        args.Add("/tlbrefpath:" + i.ItemSpec);

            foreach (var i in GetAsmPath())
                args.Add("/asmpath:" + i);

            try
            {
                using (var proc = new System.Diagnostics.Process())
                {
                    var a = new StringBuilder();
                    var b = new StringBuilder();

                    proc.StartInfo.FileName = exec;
                    proc.StartInfo.Arguments = string.Join(" ", args.Select(i => EscapeArg(i)).Where(i => i != null));
                    proc.StartInfo.CreateNoWindow = true;
                    proc.StartInfo.UseShellExecute = false;
                    proc.StartInfo.RedirectStandardOutput = true;
                    proc.StartInfo.RedirectStandardError = true;
                    proc.OutputDataReceived += (s, e) => a.AppendLine(e.Data);
                    proc.ErrorDataReceived += (s, e) => b.AppendLine(e.Data);

                    // start process
                    Log.LogMessage("{0} {1}", proc.StartInfo.FileName, proc.StartInfo.Arguments);
                    proc.Start();
                    proc.BeginOutputReadLine();
                    proc.BeginErrorReadLine();

                    // wait for exit or timeout
                    if (proc.WaitForExit((int)TimeSpan.FromMinutes(1).TotalMilliseconds) == false)
                    {
                        Log.LogError("tlbexp.exe took too long to respond.");
                        proc.Kill();
                    }

                    if (a.Length > 0)
                    {
                        var t = a.ToString().Trim();
                        if (!string.IsNullOrEmpty(t))
                            Log.LogMessage(t);
                    }

                    if (b.Length > 0)
                    {
                        var t = b.ToString().Trim();
                        if (!string.IsNullOrEmpty(t))
                            Log.LogMessage(t);
                    }

                    // success is based on exit code
                    return proc.ExitCode == 0;
                }
            }
            catch (Exception e)
            {
                Log.LogError("Exception executing TlbExp: {0}", e);
                return false;
            }
        }

    }

}
