using System.Xml.Linq;

using Microsoft.Build.Framework;

namespace Cogito.COM.MSBuild
{

    public class PatchAssemblyManifest :
        Microsoft.Build.Utilities.Task
    {

        [Required]
        public string ManifestFile { get; set; }

        /// <summary>
        /// Executes the task.
        /// </summary>
        /// <returns></returns>
        public override bool Execute()
        {
            var asmv1 = (XNamespace)"urn:schemas-microsoft-com:asm.v1";
            var xml = XDocument.Load(ManifestFile);
            //xml.Root.Element(asmv1 + "assemblyIdentity").Attributes("processorArchitecture").Remove();
            //xml.Root.Element(asmv1 + "assemblyIdentity").Attributes("type").Remove();
            //xml.Root.Element(asmv1 + "assemblyIdentity").SetAttributeValue("type", "win32");
            xml.Root.Elements(asmv1 + "file").Elements().Remove();
            xml.Root.Elements(asmv1 + "dependency").Remove();
            xml.Save(ManifestFile);

            return true;
        }

    }

}
