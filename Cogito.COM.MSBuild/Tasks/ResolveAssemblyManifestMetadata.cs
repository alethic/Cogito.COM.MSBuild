using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Xml;
using System.Xml.Linq;

using Microsoft.Build.Framework;
using Microsoft.Build.Utilities;

namespace Cogito.COM.MSBuild
{

    public class ResolveAssemblyManifestMetadata :
        Microsoft.Build.Utilities.Task
    {

        [Flags]
        enum LoadLibraryFlags : uint
        {
            DONT_RESOLVE_DLL_REFERENCES = 0x00000001,
            LOAD_IGNORE_CODE_AUTHZ_LEVEL = 0x00000010,
            LOAD_LIBRARY_AS_DATAFILE = 0x00000002,
            LOAD_LIBRARY_AS_DATAFILE_EXCLUSIVE = 0x00000040,
            LOAD_LIBRARY_AS_IMAGE_RESOURCE = 0x00000020,
            LOAD_LIBRARY_SEARCH_APPLICATION_DIR = 0x00000200,
            LOAD_LIBRARY_SEARCH_DEFAULT_DIRS = 0x00001000,
            LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR = 0x00000100,
            LOAD_LIBRARY_SEARCH_SYSTEM32 = 0x00000800,
            LOAD_LIBRARY_SEARCH_USER_DIRS = 0x00000400,
            LOAD_WITH_ALTERED_SEARCH_PATH = 0x00000008
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern IntPtr LoadLibraryEx(string lpFileName, IntPtr hReservedNull, LoadLibraryFlags dwFlags);

        [DllImport("kernel32.dll")]
        public static extern bool FreeLibrary(IntPtr hModule);

        [DllImport("kernel32.dll", EntryPoint = "FindResourceW", SetLastError = true, CharSet = CharSet.Unicode)]
        static extern IntPtr FindResource(IntPtr hModule, ushort pName, ushort pType);

        [DllImport("kernel32.dll", EntryPoint = "SizeofResource", SetLastError = true)]
        static extern uint SizeofResource(IntPtr hModule, IntPtr hResource);

        [DllImport("kernel32.dll", EntryPoint = "LoadResource", SetLastError = true)]
        static extern IntPtr LoadResource(IntPtr hModule, IntPtr hResource);

        [DllImport("kernel32.dll", EntryPoint = "LockResource")]
        static extern IntPtr LockResource(IntPtr hGlobal);

        readonly static XNamespace asmv1 = "urn:schemas-microsoft-com:asm.v1";

        /// <summary>
        /// Assemblies to scan.
        /// </summary>
        public ITaskItem[] Assemblies { get; set; }

        /// <summary>
        /// Assemblies that were resolved.
        /// </summary>
        [Output]
        public ITaskItem[] ResolvedAssemblies { get; set; }

        /// <summary>
        /// Returns <c>null</c> for an empty string.
        /// </summary>
        /// <param name="t"></param>
        /// <returns></returns>
        string TrimToNull(string t)
        {
            return !string.IsNullOrWhiteSpace(t) ? t.Trim() : null;
        }

        /// <summary>
        /// Reads the assembly identity from the specified manifest file.
        /// </summary>
        /// <param name="itemSpec"></param>
        /// <param name="xml"></param>
        /// <param name="manifestPath"></param>
        /// <returns></returns>
        ITaskItem ResolveIdentityFromManifest(string itemSpec, XDocument xml, string manifestPath)
        {
            var identity = xml.Root.Elements(asmv1 + "assemblyIdentity").FirstOrDefault();
            if (identity == null)
                return null;

            return new TaskItem((string)identity.Attribute("name"), new Dictionary<string, string>()
            {
                ["Name"] = (string)identity.Attribute("name") ?? "",
                ["Type"] = (string)identity.Attribute("type") ?? "",
                ["Version"] = (string)identity.Attribute("version") ?? "",
                ["ProcessorArchitecture"] = (string)identity.Attribute("processorArchitecture") ?? "",
                ["PublicKeyToken"] = (string)identity.Attribute("publicKeyToken") ?? "",
                ["Language"] = (string)identity.Attribute("language") ?? "",
                ["OriginalItemSpec"] = itemSpec,
                ["ManifestPath"] = manifestPath,
            });
        }

        /// <summary>
        /// Reads the manifest resource from the given assembly.
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        XDocument ReadManifestFromAssembly(string path)
        {
            var l = LoadLibraryEx(path, IntPtr.Zero, LoadLibraryFlags.LOAD_LIBRARY_AS_DATAFILE);
            if (l == IntPtr.Zero)
                return null;

            try
            {
                var r = FindResource(l, 1, 24);

                // check secondary resource
                if (r == IntPtr.Zero)
                    r = FindResource(l, 2, 24);

                // no go!
                if (r == IntPtr.Zero)
                    return null;

                var h = LoadResource(l, r);
                if (h == IntPtr.Zero)
                    return null;

                var b = LockResource(h);
                if (b == IntPtr.Zero)
                    return null;

                var s = SizeofResource(l, r);
                var d = new byte[s];
                Marshal.Copy(b, d, 0, (int)s);

                try
                {
                    using (var t = new StreamReader(new MemoryStream(d)))
                        return XDocument.Load(t);
                }
                catch (IOException e)
                {
                    Log.LogErrorFromException(e);
                    return null;
                }
                catch (XmlException e)
                {
                    Log.LogErrorFromException(e);
                    return null;
                }
            }
            finally
            {
                if (l != IntPtr.Zero)
                    FreeLibrary(l);
            }
        }

        /// <summary>
        /// Resolve the given COM reference.
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        ITaskItem ResolveIdentity(ITaskItem item)
        {
            if (item == null)
                return null;

            try
            {
                var path = Path.GetFullPath(TrimToNull(item.GetMetadata("FullPath")) ?? item.ItemSpec);
                if (File.Exists(path) == false)
                    return null;

                // skip unsupported items
                if (!string.Equals(Path.GetExtension(path), ".dll", StringComparison.InvariantCultureIgnoreCase) &&
                    !string.Equals(Path.GetExtension(path), ".exe", StringComparison.InvariantCultureIgnoreCase) &&
                    !string.Equals(Path.GetExtension(path), ".ocx", StringComparison.InvariantCultureIgnoreCase))
                    return null;

                // separate manifest file
                if (File.Exists(path + ".manifest"))
                    return ResolveIdentityFromManifest(item.ItemSpec, XDocument.Load(path + ".manifest"), path + ".manifest");

                // attempt to load manifest from assembly itself
                var manifest = ReadManifestFromAssembly(path);
                if (manifest != null)
                    return ResolveIdentityFromManifest(item.ItemSpec, manifest, path);
            }
            catch (Exception e)
            {
                Log.LogWarningFromException(e, true);
            }

            // oh well
            return null;
        }

        /// <summary>
        /// Executes the task.
        /// </summary>
        /// <returns></returns>
        public override bool Execute()
        {
            ResolvedAssemblies = Assemblies?.Select(i => ResolveIdentity(i)).Where(i => i != null).ToArray();
            return true;
        }

    }

}
