using System.IO;
using System.Linq;
using System.Xml.Linq;

using Microsoft.Build.Framework;

namespace Cogito.COM.MsBuild
{

    public class UpdateAssemblyManifest :
        Microsoft.Build.Utilities.Task
    {

        readonly static XNamespace asmv1 = "urn:schemas-microsoft-com:asm.v1";

        /// <summary>
        /// Path of source manifest file.
        /// </summary>
        public string ManifestSource { get; set; }

        /// <summary>
        /// Path of output manifest file.
        /// </summary>
        public string ManifestOutput { get; set; }

        /// <summary>
        /// Dependency information to add or update.
        /// </summary>
        public ITaskItem[] Dependencies { get; set; }

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
        /// Updates a single dependency in the manifest.
        /// </summary>
        /// <param name="manifest"></param>
        /// <param name="dependency"></param>
        void UpdateDependency(XDocument manifest, ITaskItem dependency)
        {
            var name = TrimToNull(dependency.GetMetadata("Name"));
            var type = TrimToNull(dependency.GetMetadata("Type"));
            var version = TrimToNull(dependency.GetMetadata("Version"));
            var processorArchitecture = TrimToNull(dependency.GetMetadata("ProcessorArchitecture"));

            // at least name is required
            if (name == null)
                return;

            Log.LogMessage("Updating Dependency '" + name + "' from '" + dependency.ItemSpec + "'.");

            // find or create dependency element
            var d = manifest.Root
                .Elements(asmv1 + "dependency")
                .Elements(asmv1 + "dependentAssembly")
                .Elements(asmv1 + "assemblyIdentity")
                .Where(i => (string)i.Attribute("name") == name)
                .FirstOrDefault();
            if (d == null)
                manifest.Root.Add(
                    new XElement(asmv1 + "dependency",
                        new XElement(asmv1 + "dependentAssembly", d =
                            new XElement(asmv1 + "assemblyIdentity"))));

            // update name
            if (name != null)
                d.SetAttributeValue("name", name);

            // update type
            if (type != null)
                d.SetAttributeValue("type", type);

            // update version
            if (version != null)
                d.SetAttributeValue("version", version);

            // update processorArchitecture
            if (processorArchitecture != null)
                d.SetAttributeValue("processorArchitecture", processorArchitecture);
        }

        /// <summary>
        /// Adds any additional dependencies to the manifest.
        /// </summary>
        /// <param name="manifest"></param>
        void UpdateDependencies(XDocument manifest)
        {
            if (Dependencies != null)
                foreach (var d in Dependencies)
                    if (d != null)
                        UpdateDependency(manifest, d);
        }

        /// <summary>
        /// Executes the task.
        /// </summary>
        /// <returns></returns>
        public override bool Execute()
        {
            // load existing manifest
            var manifest = File.Exists(ManifestSource) ? XDocument.Load(ManifestSource) : null;
            if (manifest == null)
                return false;

            UpdateDependencies(manifest);

            // save new manifest file
            manifest.Save(ManifestOutput);
            return true;
        }

    }

}
