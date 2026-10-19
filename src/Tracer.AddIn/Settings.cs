using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;

namespace Tracer.AddIn;

public enum KeyPreset
{
    /// <summary>Ctrl+Q precedents, Ctrl+Shift+Q dependents (Arixcel's defaults).</summary>
    Arixcel,
    /// <summary>Ctrl+[ precedents, Ctrl+] dependents (replaces Excel's built-in Go To Precedents/Dependents).</summary>
    Brackets,
    Both,
}

/// <summary>Per-user settings in %APPDATA%\FormulaTracer\settings.ini (key=value).</summary>
internal sealed class Settings
{
    private static readonly string Dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "FormulaTracer");
    private static readonly string FilePath = Path.Combine(Dir, "settings.ini");

    public static Settings Current { get; } = Load();

    public KeyPreset Keys { get; set; } = KeyPreset.Arixcel;
    /// <summary>Index dependents at idle after a workbook opens, so the first Ctrl+Shift+Q is instant.</summary>
    public bool Prebuild { get; set; } = true;
    /// <summary>Show timings in the status line.</summary>
    public bool ShowTimings { get; set; } = true;
    public Rectangle? WindowBounds { get; set; }

    private static Settings Load()
    {
        var s = new Settings();
        try
        {
            if (!File.Exists(FilePath)) return s;
            var kv = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var line in File.ReadAllLines(FilePath))
            {
                int eq = line.IndexOf('=');
                if (eq > 0) kv[line.Substring(0, eq).Trim()] = line.Substring(eq + 1).Trim();
            }
            if (kv.TryGetValue("Keys", out var k) && Enum.TryParse(k, true, out KeyPreset preset)) s.Keys = preset;
            if (kv.TryGetValue("Prebuild", out var p) && bool.TryParse(p, out var pb)) s.Prebuild = pb;
            if (kv.TryGetValue("ShowTimings", out var t) && bool.TryParse(t, out var st)) s.ShowTimings = st;
            if (kv.TryGetValue("WindowBounds", out var b))
            {
                var parts = b.Split(',');
                if (parts.Length == 4
                    && int.TryParse(parts[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out int x)
                    && int.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out int y)
                    && int.TryParse(parts[2], NumberStyles.Integer, CultureInfo.InvariantCulture, out int w)
                    && int.TryParse(parts[3], NumberStyles.Integer, CultureInfo.InvariantCulture, out int h))
                    s.WindowBounds = new Rectangle(x, y, w, h);
            }
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
        return s;
    }

    public void Save()
    {
        try
        {
            Directory.CreateDirectory(Dir);
            var lines = new List<string>
            {
                "Keys=" + Keys,
                "Prebuild=" + Prebuild,
                "ShowTimings=" + ShowTimings,
            };
            if (WindowBounds is Rectangle r)
                lines.Add(string.Format(CultureInfo.InvariantCulture, "WindowBounds={0},{1},{2},{3}", r.X, r.Y, r.Width, r.Height));
            File.WriteAllLines(FilePath, lines);
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }
}
