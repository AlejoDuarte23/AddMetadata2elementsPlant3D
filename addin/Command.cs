using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System.Text.Json;

namespace Plant3DJsonMetadataAddin;

public sealed class Commands
{
    [CommandMethod("P3D_APPLYMETA", CommandFlags.Modal)]
    public void ApplyMetadataFromJson()
    {
        var doc = Application.DocumentManager.MdiActiveDocument;
        if (doc is null) return;
        var ed = doc.Editor;

        try
        {
            var pfo = new PromptOpenFileOptions("\nSelect metadata JSON file")
            {
                Filter = "JSON (*.json)|*.json|All files (*.*)|*.*"
            };

            var pfr = ed.GetFileNameForOpen(pfo);
            if (pfr.Status != PromptStatus.OK) return;

            var json = System.IO.File.ReadAllText(pfr.StringResult);
            var payload = PayloadParser.Parse(json);

            var report = MetadataApplier.Apply(payload);

            ed.WriteMessage($"\n--- Plant3D JSON Metadata Report ---");
            ed.WriteMessage($"\nProject part used: {report.UsedProjectPart}");
            ed.WriteMessage($"\nScanned entities: {report.ScannedEntities}");
            ed.WriteMessage($"\nLinked entities:  {report.LinkedEntities}");
            ed.WriteMessage($"\nUpdated rows:     {report.UpdatedRows}");

            if (report.NotFoundElements.Count > 0)
                ed.WriteMessage($"\nNot found: {string.Join(", ", report.NotFoundElements)}");

            foreach (var kv in report.MissingColumnsPerElement)
                ed.WriteMessage($"\nMissing columns for '{kv.Key}': {string.Join(", ", kv.Value)}");

            foreach (var w in report.Warnings)
                ed.WriteMessage($"\nWarning: {w}");

            ed.WriteMessage("\n-----------------------------------");
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nERROR: {ex.Message}\n{ex}");
        }
    }
}
