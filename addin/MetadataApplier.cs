
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.ProcessPower.DataObjects;
using Autodesk.ProcessPower.DataLinks;
using Autodesk.ProcessPower.PlantInstance;
using Autodesk.ProcessPower.ProjectManager;

namespace Plant3DJsonMetadataAddin;

internal sealed class ApplyReport
{
    public string UsedProjectPart { get; init; } = "";
    public int ScannedEntities { get; set; }
    public int LinkedEntities { get; set; }
    public int UpdatedRows { get; set; }
    public Dictionary<string, int> UpdatedPerElement { get; } = new(StringComparer.OrdinalIgnoreCase);
    public Dictionary<string, List<string>> MissingColumnsPerElement { get; } = new(StringComparer.OrdinalIgnoreCase);
    public List<string> NotFoundElements { get; } = new();
    public List<string> Warnings { get; } = new();
}

internal static class MetadataApplier
{
    private static readonly string[] DefaultMatchColumns =
    [
        "Tag", "TagValue", "Name", "Number", "LineNumberTag", "LineNumber"
    ];

    public static ApplyReport Apply(MetadataPayload payload)
    {
        var doc = Application.DocumentManager.MdiActiveDocument
                  ?? throw new InvalidOperationException("No active document.");

        if (PlantApplication.CurrentProject is null)
            throw new InvalidOperationException("No active Plant 3D project (PlantApplication.CurrentProject is null).");

        var candidates = !string.IsNullOrWhiteSpace(payload.ProjectPart)
            ? new[] { payload.ProjectPart! }
            : new[] { "Piping", "PnId" };

        var matchColumns = (payload.MatchColumns is { Count: > 0 } ? payload.MatchColumns : DefaultMatchColumns)
            .Select(c => c.Trim())
            .Where(c => c.Length > 0)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        foreach (var partName in candidates)
        {
            var report = TryApplyUsingProjectPart(doc, partName, payload.Elements, matchColumns);
            if (report is null) continue;

            if (!string.IsNullOrWhiteSpace(payload.ProjectPart))
                return report; // user forced the part

            if (report.UpdatedRows > 0)
                return report; // auto-mode: first successful
        }

        return new ApplyReport
        {
            UsedProjectPart = payload.ProjectPart ?? "(auto)",
            Warnings = { "No matching Plant-linked entities were updated (wrong projectPart, no links, or names not found)." }
        };
    }

    private static ApplyReport? TryApplyUsingProjectPart(
        Document doc,
        string projectPartName,
        IReadOnlyList<ElementUpdate> elements,
        IReadOnlyList<string> matchColumns)
    {
        var plantPrj = PlantApplication.CurrentProject!;
        if (!plantPrj.ProjectParts.Contains(projectPartName)) return null;
        if (plantPrj.ProjectParts[projectPartName] is not Project prj) return null;

        DataLinksManager dlm;
        try { dlm = prj.DataLinksManager; }
        catch { return null; }

        var pnpDb = dlm.GetPnPDatabase();
        if (pnpDb is null) return null;

        var report = new ApplyReport { UsedProjectPart = projectPartName };

        var byHandle = elements
            .Where(e => !string.IsNullOrWhiteSpace(e.Handle))
            .ToDictionary(e => e.Handle!, StringComparer.OrdinalIgnoreCase);

        var byName = elements
            .GroupBy(e => e.Name, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.ToList(), StringComparer.OrdinalIgnoreCase);

        var targetNames = new HashSet<string>(byName.Keys, StringComparer.OrdinalIgnoreCase);

        using (doc.LockDocument())
        using (var tr = doc.Database.TransactionManager.StartTransaction())
        {
            var bt = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
            var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

            var updatedRowIds = new HashSet<int>();

            foreach (ObjectId entId in ms)
            {
                report.ScannedEntities++;

                int rowId;
                try { rowId = dlm.FindAcPpRowId(entId); }
                catch { continue; }

                if (rowId <= 0) continue;
                report.LinkedEntities++;

                // Handle direct match
                var handleStr = entId.Handle.ToString();
                if (byHandle.TryGetValue(handleStr, out var handleUpdate))
                {
                    UpdateRow(pnpDb, rowId, handleUpdate, report, updatedRowIds);
                    continue;
                }

                // Name/Tag match
                PnPRow row;
                try { row = pnpDb.GetRow(rowId); }
                catch { continue; }

                var key = GetFirstNonEmpty(row, matchColumns);
                if (key is null || !targetNames.Contains(key)) continue;

                foreach (var upd in byName[key])
                    UpdateRow(row, upd, report, updatedRowIds);
            }

            tr.Commit();
        }

        foreach (var name in targetNames)
            if (!report.UpdatedPerElement.ContainsKey(name))
                report.NotFoundElements.Add(name);

        return report;
    }

    private static void UpdateRow(PnPDatabase pnpDb, int rowId, ElementUpdate upd, ApplyReport report, HashSet<int> updatedRowIds)
    {
        if (updatedRowIds.Contains(rowId)) return;

        PnPRow row;
        try { row = pnpDb.GetRow(rowId); }
        catch (Exception ex)
        {
            report.Warnings.Add($"Failed to get rowId={rowId} for '{upd.Name}': {ex.Message}");
            return;
        }

        UpdateRow(row, upd, report, updatedRowIds);
    }

    private static void UpdateRow(PnPRow row, ElementUpdate upd, ApplyReport report, HashSet<int> updatedRowIds)
    {
        var rowId = row.RowId;
        if (!updatedRowIds.Add(rowId)) return;

        var missing = new List<string>();

        try
        {
            row.BeginEdit();

            foreach (var kv in upd.Properties)
            {
                if (row.Table.Columns.Contains(kv.Key))
                    row[kv.Key] = kv.Value;
                else
                    missing.Add(kv.Key);
            }

            row.EndEdit();
        }
        catch (Exception ex)
        {
            report.Warnings.Add($"Failed updating '{upd.Name}' (rowId={rowId}): {ex.Message}");
            return;
        }

        report.UpdatedRows++;
        report.UpdatedPerElement[upd.Name] = report.UpdatedPerElement.TryGetValue(upd.Name, out var c) ? c + 1 : 1;

        if (missing.Count > 0)
        {
            if (!report.MissingColumnsPerElement.TryGetValue(upd.Name, out var list))
                report.MissingColumnsPerElement[upd.Name] = list = new List<string>();

            foreach (var m in missing)
                if (!list.Contains(m, StringComparer.OrdinalIgnoreCase))
                    list.Add(m);
        }
    }

    private static string? GetFirstNonEmpty(PnPRow row, IReadOnlyList<string> columns)
    {
        foreach (var col in columns)
        {
            if (!row.Table.Columns.Contains(col)) continue;
            var v = row[col]?.ToString();
            if (!string.IsNullOrWhiteSpace(v)) return v.Trim();
        }
        return null;
    }
}
