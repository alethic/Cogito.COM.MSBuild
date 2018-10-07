using System;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;

namespace Cogito.COM.MsBuild
{

    public class EmbedTypeLib :
        Microsoft.Build.Utilities.Task
    {

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern IntPtr BeginUpdateResource(string pFileName, [MarshalAs(UnmanagedType.Bool)]bool bDeleteExistingResources);

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool UpdateResource(IntPtr hUpdate, string lpType, ushort lpName, ushort wLanguage, byte[] lpData, uint cbData);

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool EndUpdateResource(IntPtr hUpdate, bool fDiscard);

        /// <summary>
        /// Destination path.
        /// </summary>
        public string TargetPath { get; set; }

        /// <summary>
        /// Alternative input file.
        /// </summary>
        public string SourcePath { get; set; }

        /// <summary>
        /// Path of the type lib to embed.
        /// </summary>
        public string TypeLibPath { get; set; }

        /// <summary>
        /// Executes the task.
        /// </summary>
        /// <returns></returns>
        public override bool Execute()
        {
            // by default update same file
            if (SourcePath == null)
                SourcePath = TargetPath;

            if (File.Exists(TargetPath) == false)
            {
                Log.LogError("Could not find target file: {0}", TargetPath);
                return false;
            }

            if (File.Exists(SourcePath) == false)
            {
                Log.LogError("Could not find source file: {0}", SourcePath);
                return false;
            }

            // files are not the same, replace target with source
            if (SourcePath != TargetPath)
            {
                Log.LogMessage("Copying {0} to {1}.", SourcePath, TargetPath);
                File.Copy(SourcePath, TargetPath);
            }

            // type library required
            if (File.Exists(TypeLibPath) == false)
            {
                Log.LogError("Type library {0} not found.", TypeLibPath);
                return false;
            }

            // load type library and pin
            var tlb = File.ReadAllBytes(TypeLibPath);
            var hnd = IntPtr.Zero;

            try
            {
                Log.LogMessage("Embedding type library {0} into {1}.", TypeLibPath, TargetPath);

                // make temporary copy of target file
                File.Copy(TargetPath, TargetPath + ".tmp", true);

                // open target path file
                hnd = BeginUpdateResource(TargetPath + ".tmp", false);
                if (hnd == IntPtr.Zero)
                {
                    Log.LogError("BeginUpdateResource: {0}", hnd);
                    return false;
                }

                // update resource
                if (!UpdateResource(hnd, "TYPELIB", 1, 0, tlb, (uint)tlb.Length))
                    throw new Win32Exception(Marshal.GetLastWin32Error());

                // close handle
                if (!EndUpdateResource(hnd, false))
                    throw new Win32Exception(Marshal.GetLastWin32Error());
                hnd = IntPtr.Zero;

                // replace from temporary copy
                File.Copy(TargetPath + ".tmp", TargetPath, true);
                File.Delete(TargetPath + ".tmp");
            }
            catch (Exception e)
            {
                Log.LogError(e.ToString());
                return false;
            }
            finally
            {
                // try to 
                if (hnd != IntPtr.Zero)
                    EndUpdateResource(hnd, false);
            }

            return true;
        }

    }

}
