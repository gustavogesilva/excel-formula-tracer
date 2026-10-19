using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.Win32;

namespace Tracer.Setup;

/// <summary>
/// Per-user install: the XLL goes to %LOCALAPPDATA%\FormulaTracer and Excel is told to open it at startup
/// through its OPEN values (HKCU), which is exactly what ticking it in File › Options › Add-ins does.
/// Excel rewrites those values when it exits, so callers must make sure Excel is closed first.
/// </summary>
internal static class Installer
{
    public const string DisplayName = "Formula Tracer for Excel";
    private const string UninstallKey = @"Software\Microsoft\Windows\CurrentVersion\Uninstall\FormulaTracer";
    private static readonly string[] OfficeVersions = { "16.0", "15.0", "14.0" };

    public static string InstallDir => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "FormulaTracer");
    public static string XllPath => Path.Combine(InstallDir, "FormulaTracer.xll");
    public static string UninstallerPath => Path.Combine(InstallDir, "FormulaTracerSetup.exe");

    public static string Version => Assembly.GetExecutingAssembly().GetName().Version.ToString(3);

    /// <summary>Installed version from the Apps &amp; features entry, or null when not installed.</summary>
    public static string? InstalledVersion()
    {
        using var key = Registry.CurrentUser.OpenSubKey(UninstallKey);
        return key?.GetValue("DisplayVersion") as string;
    }

    public static void Install(ExcelInstall excel)
    {
        Directory.CreateDirectory(InstallDir);

        // Written from the embedded resource, so the file carries no Mark-of-the-Web and
        // Microsoft 365's "block XLLs from the internet" policy does not apply to it.
        WriteResource(excel.Is64Bit ? "FormulaTracer64.xll" : "FormulaTracer32.xll", XllPath);

        // Keep a copy of this program as the uninstaller (bytes, not File.Copy, so no Zone.Identifier stream is copied).
        string self = Assembly.GetExecutingAssembly().Location;
        if (!string.Equals(Path.GetFullPath(self), Path.GetFullPath(UninstallerPath), StringComparison.OrdinalIgnoreCase))
            File.WriteAllBytes(UninstallerPath, File.ReadAllBytes(self));

        using (var options = Registry.CurrentUser.CreateSubKey($@"Software\Microsoft\Office\{excel.RegistryVersion}\Excel\Options"))
            RewriteOpenValues(options!, current => OpenValues.Install(current, XllPath));
        RemoveFromAddInManager(excel.RegistryVersion);

        using var k = Registry.CurrentUser.CreateSubKey(UninstallKey)!;
        k.SetValue("DisplayName", DisplayName);
        k.SetValue("DisplayVersion", Version);
        k.SetValue("Publisher", "gustavogesilva");
        k.SetValue("URLInfoAbout", "https://github.com/gustavogesilva/excel-formula-tracer");
        k.SetValue("InstallLocation", InstallDir);
        k.SetValue("DisplayIcon", UninstallerPath);
        k.SetValue("UninstallString", $"\"{UninstallerPath}\" /uninstall");
        k.SetValue("QuietUninstallString", $"\"{UninstallerPath}\" /uninstall /quiet");
        k.SetValue("NoModify", 1, RegistryValueKind.DWord);
        k.SetValue("NoRepair", 1, RegistryValueKind.DWord);
        k.SetValue("EstimatedSize", (int)(DirectorySize(InstallDir) / 1024), RegistryValueKind.DWord);
    }

    /// <summary>Removes the add-in from every Excel version's lists, the Apps &amp; features entry, and the install folder.</summary>
    public static void Uninstall()
    {
        foreach (var ver in OfficeVersions)
        {
            using (var options = Registry.CurrentUser.OpenSubKey($@"Software\Microsoft\Office\{ver}\Excel\Options", writable: true))
                if (options != null) RewriteOpenValues(options, OpenValues.Uninstall);
            RemoveFromAddInManager(ver);
        }
        Registry.CurrentUser.DeleteSubKeyTree(UninstallKey, throwOnMissingSubKey: false);
        try
        {
            if (File.Exists(XllPath)) File.Delete(XllPath);
        }
        catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
        {
        }
        PendingFolderDelete = true;
    }

    /// <summary>Set by <see cref="Uninstall"/>; <c>Program</c> deletes the folder just before the process exits.</summary>
    public static bool PendingFolderDelete { get; private set; }

    /// <summary>
    /// Rewrites OPEN, OPEN1, … as a consecutive list (Excel stops reading at a gap), preserving the order and value
    /// kinds of other add-ins. Overwrites positions in place, then deletes the leftover tail.
    /// </summary>
    private static void RewriteOpenValues(RegistryKey options, Func<List<string>, List<string>> change)
    {
        var names = options.GetValueNames().Where(n => OpenValues.IndexOf(n) >= 0).ToList();
        var kinds = new Dictionary<string, RegistryValueKind>();
        var pairs = new List<KeyValuePair<string, string>>();
        foreach (var n in names)
        {
            string data = options.GetValue(n, "", RegistryValueOptions.DoNotExpandEnvironmentNames) as string ?? "";
            pairs.Add(new KeyValuePair<string, string>(n, data));
            kinds[data] = options.GetValueKind(n);
        }

        var next = change(OpenValues.Ordered(pairs));
        for (int i = 0; i < next.Count; i++)
            options.SetValue(OpenValues.NameAt(i), next[i], kinds.TryGetValue(next[i], out var kind) ? kind : RegistryValueKind.String);
        foreach (var n in names)
            if (OpenValues.IndexOf(n) >= next.Count) options.DeleteValue(n, throwOnMissingValue: false);
    }

    /// <summary>Unticked add-ins live under "Add-in Manager" with the path as the value name.</summary>
    private static void RemoveFromAddInManager(string officeVersion)
    {
        using var key = Registry.CurrentUser.OpenSubKey($@"Software\Microsoft\Office\{officeVersion}\Excel\Add-in Manager", writable: true);
        if (key == null) return;
        foreach (var name in key.GetValueNames())
            if (OpenValues.IsOurs(name)) key.DeleteValue(name, throwOnMissingValue: false);
    }

    private static void WriteResource(string resourceName, string path)
    {
        using var source = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName)
            ?? throw new InvalidOperationException("Missing embedded resource " + resourceName);
        using var target = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None);
        source.CopyTo(target);
    }

    private static long DirectorySize(string dir) =>
        Directory.Exists(dir) ? new DirectoryInfo(dir).GetFiles().Sum(f => f.Length) : 0;

    /// <summary>The running uninstaller may live in the folder, so a detached cmd deletes it a moment after we exit.</summary>
    public static void DeleteFolderAfterExit()
    {
        string dir = InstallDir;
        if (!Directory.Exists(dir)) return;
        Process.Start(new ProcessStartInfo("cmd.exe", $"/c ping -n 4 127.0.0.1 >nul & rmdir /s /q \"{dir}\"")
        {
            CreateNoWindow = true,
            UseShellExecute = false,
            WindowStyle = ProcessWindowStyle.Hidden,
            WorkingDirectory = Path.GetTempPath(),
        });
    }
}
