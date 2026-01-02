
using Autodesk.AutoCAD.Runtime;

[assembly: ExtensionApplication(typeof(Plant3DJsonMetadataAddin.ExtApp))]

namespace Plant3DJsonMetadataAddin;

public sealed class ExtApp : IExtensionApplication
{
    public void Initialize() { }
    public void Terminate() { }
}