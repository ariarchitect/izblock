using System.Net.Http.Json;
using System.Text.Json.Nodes;

namespace IzArchicadPlacer.App;

internal sealed class TapirClient
{
    private readonly HttpClient _http;

    public TapirClient(string baseUrl)
    {
        _http = new HttpClient
        {
            BaseAddress = new Uri(baseUrl.TrimEnd('/') + "/")
        };
    }

    public async Task<List<StoryInfo>> GetStoriesAsync(CancellationToken ct = default)
    {
        JsonObject data = await ExecuteTapirCommandAsync("GetStories", null, ct);
        JsonArray stories = data["stories"] as JsonArray ?? [];
        var result = new List<StoryInfo>();
        foreach (JsonNode? n in stories)
        {
            if (n is not JsonObject o)
            {
                continue;
            }

            int? idx = o["index"]?.GetValue<int?>();
            double? level = o["level"]?.GetValue<double?>();
            if (!idx.HasValue || !level.HasValue)
            {
                continue;
            }

            result.Add(new StoryInfo
            {
                Index = idx.Value,
                Name = o["name"]?.GetValue<string>() ?? string.Empty,
                Level = level.Value
            });
        }

        return result.OrderBy(s => s.Index).ToList();
    }

    public async Task<List<string>> CreateObjectsAsync(IReadOnlyList<ObjectPayload> objects, CancellationToken ct = default)
    {
        var objectsData = new JsonArray();
        foreach (ObjectPayload o in objects)
        {
            objectsData.Add(new JsonObject
            {
                ["libraryPartName"] = o.LibraryPartName,
                ["coordinates"] = new JsonObject
                {
                    ["x"] = o.X,
                    ["y"] = o.Y,
                    ["z"] = o.Z
                },
                ["dimensions"] = new JsonObject
                {
                    ["x"] = 1.0,
                    ["y"] = 1.0,
                    ["z"] = 1.0
                }
            });
        }

        JsonObject data = await ExecuteTapirCommandAsync("CreateObjects", new JsonObject { ["objectsData"] = objectsData }, ct);
        JsonArray elements = data["elements"] as JsonArray ?? [];
        var guids = new List<string>();
        foreach (JsonNode? n in elements)
        {
            string? guid = n?["elementId"]?["guid"]?.GetValue<string>();
            if (!string.IsNullOrWhiteSpace(guid))
            {
                guids.Add(guid);
            }
        }

        return guids;
    }

    public async Task<List<JsonObject>> GetAllPropertiesAsync(CancellationToken ct = default)
    {
        JsonObject data = await ExecuteTapirCommandAsync("GetAllProperties", null, ct);
        JsonArray arr = data["properties"] as JsonArray ?? [];
        return arr.OfType<JsonObject>().ToList();
    }

    public async Task SetPropertyValuesOfElementsAsync(JsonArray elementPropertyValues, CancellationToken ct = default)
    {
        await ExecuteTapirCommandAsync("SetPropertyValuesOfElements", new JsonObject { ["elementPropertyValues"] = elementPropertyValues }, ct);
    }

    public async Task<string?> GetBuiltInElementIdPropertyGuidAsync(CancellationToken ct = default)
    {
        JsonObject root = await ExecuteApiCommandAsync(
            "API.GetPropertyIds",
            new JsonObject
            {
                ["properties"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["type"] = "BuiltIn",
                        ["nonLocalizedName"] = "General_ElementID"
                    }
                }
            },
            ct);

        JsonArray? props = root["properties"] as JsonArray;
        return props?.FirstOrDefault()?["propertyId"]?["guid"]?.GetValue<string>();
    }

    private async Task<JsonObject> ExecuteTapirCommandAsync(string commandName, JsonObject? parameters, CancellationToken ct)
    {
        var addOnCommandId = new JsonObject
        {
            ["commandNamespace"] = "TapirCommand",
            ["commandName"] = commandName
        };

        var p = new JsonObject
        {
            ["addOnCommandId"] = addOnCommandId
        };
        if (parameters != null)
        {
            p["addOnCommandParameters"] = parameters;
        }

        JsonObject result = await ExecuteApiCommandAsync("API.ExecuteAddOnCommand", p, ct);
        return result["addOnCommandResponse"] as JsonObject ?? new JsonObject();
    }

    private async Task<JsonObject> ExecuteApiCommandAsync(string command, JsonObject? parameters, CancellationToken ct)
    {
        var body = new JsonObject
        {
            ["command"] = command
        };
        if (parameters != null)
        {
            body["parameters"] = parameters;
        }

        using HttpResponseMessage response = await _http.PostAsJsonAsync("", body, ct);
        response.EnsureSuccessStatusCode();
        JsonObject? root = await response.Content.ReadFromJsonAsync<JsonObject>(cancellationToken: ct);
        if (root is null)
        {
            throw new InvalidOperationException("Empty response from Archicad/Tapir.");
        }

        bool succeeded = root["succeeded"]?.GetValue<bool>() ?? false;
        if (!succeeded)
        {
            string message = root["error"]?["message"]?.GetValue<string>() ?? "Unknown error.";
            throw new InvalidOperationException($"Tapir RPC error: {message}");
        }

        return root["result"] as JsonObject ?? new JsonObject();
    }
}
