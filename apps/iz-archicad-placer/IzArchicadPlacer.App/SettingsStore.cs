using System.Text.Json;

namespace IzArchicadPlacer.App;

internal sealed class AppSettings
{
    public string Units { get; set; } = "mm";
    public string Prop1Set { get; set; } = "INFORMATION ABOUT MANUFACTURER";
    public string Prop1Name { get; set; } = "Manufacturer";
    public string Prop2Set { get; set; } = "EQUIPMENT";
    public string Prop2Name { get; set; } = "Model";
    public bool WriteToId { get; set; }
    public string TapirBaseUrl { get; set; } = "http://127.0.0.1:19725";
}

internal static class SettingsStore
{
    private static readonly string SettingsDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "IZBLOCK");

    private static readonly string SettingsPath = Path.Combine(SettingsDir, "archicad_placer.settings.json");

    public static AppSettings Load()
    {
        try
        {
            if (!File.Exists(SettingsPath))
            {
                return new AppSettings();
            }

            var json = File.ReadAllText(SettingsPath);
            return JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
        }
        catch
        {
            return new AppSettings();
        }
    }

    public static void Save(AppSettings settings)
    {
        Directory.CreateDirectory(SettingsDir);
        string json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(SettingsPath, json);
    }
}
