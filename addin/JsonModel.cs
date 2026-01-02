
using System.Text.Json;

namespace Plant3DJsonMetadataAddin;

public sealed record ElementUpdate(string Name, IReadOnlyDictionary<string, string> Properties, string? Handle = null);

public sealed record MetadataPayload(
    IReadOnlyList<ElementUpdate> Elements,
    string? ProjectPart = null,
    IReadOnlyList<string>? MatchColumns = null
);

internal static class PayloadParser
{
    public static MetadataPayload Parse(string json)
    {
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        string? projectPart = root.TryGetProperty("projectPart", out var pp) ? pp.GetString() : null;

        List<string>? matchColumns = null;
        if (root.TryGetProperty("matchColumns", out var mc) && mc.ValueKind == JsonValueKind.Array)
        {
            matchColumns = new List<string>();
            foreach (var v in mc.EnumerateArray())
            {
                var s = v.GetString();
                if (!string.IsNullOrWhiteSpace(s)) matchColumns.Add(s);
            }
        }

        if (!root.TryGetProperty("elements", out var elementsEl) || elementsEl.ValueKind != JsonValueKind.Array)
            throw new JsonException("Payload must contain an 'elements' array.");

        var elements = new List<ElementUpdate>();

        foreach (var el in elementsEl.EnumerateArray())
        {
            var name = el.TryGetProperty("name", out var n) ? n.GetString() : null;
            if (string.IsNullOrWhiteSpace(name))
                throw new JsonException("Each element must have a non-empty 'name'.");

            string? handle = el.TryGetProperty("handle", out var h) ? h.GetString() : null;

            if (!el.TryGetProperty("properties", out var propsEl))
                throw new JsonException($"Element '{name}' is missing 'properties'.");

            var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            if (propsEl.ValueKind == JsonValueKind.Object)
            {
                foreach (var p in propsEl.EnumerateObject())
                    props[p.Name] = p.Value.ToString();
            }
            else if (propsEl.ValueKind == JsonValueKind.Array)
            {
                foreach (var p in propsEl.EnumerateArray())
                {
                    if (!p.TryGetProperty("name", out var pn) || !p.TryGetProperty("value", out var pv))
                        throw new JsonException($"Element '{name}' has an invalid properties[] item (needs name/value).");

                    var key = pn.GetString();
                    if (!string.IsNullOrWhiteSpace(key))
                        props[key] = pv.ToString();
                }
            }
            else
            {
                throw new JsonException($"Element '{name}' has invalid 'properties' (must be object or array).");
            }

            elements.Add(new ElementUpdate(name!, props, handle));
        }

        return new MetadataPayload(elements, projectPart, matchColumns);
    }
}