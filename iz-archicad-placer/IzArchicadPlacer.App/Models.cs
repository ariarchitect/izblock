namespace IzArchicadPlacer.App;

internal sealed class StoryInfo
{
    public int Index { get; init; }
    public string Name { get; init; } = string.Empty;
    public double Level { get; init; }

    public int DisplayIndex => Index >= 0 ? Index + 1 : Index;
}

internal sealed class SourceRow
{
    public string Filename { get; init; } = string.Empty;
    public double MinX { get; init; }
    public double MinY { get; init; }
    public string Text1 { get; init; } = string.Empty;
    public string Text2 { get; init; } = string.Empty;
}

internal sealed class SheetData
{
    public string SheetName { get; init; } = string.Empty;
    public List<SourceRow> Rows { get; init; } = [];
}

internal sealed class ObjectPayload
{
    public required string LibraryPartName { get; init; }
    public required double X { get; init; }
    public required double Y { get; init; }
    public required double Z { get; init; }
    public string? Text1 { get; init; }
    public string? Text2 { get; init; }
}
