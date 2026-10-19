using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Microsoft.Win32;

namespace Tracer.Setup;

/// <summary>The installed Excel: where it is, which registry version folder it uses, and its bitness.</summary>
internal sealed class ExcelInstall
{
    public ExcelInstall(string exePath, string registryVersion, bool is64Bit, string productVersion)
    {
        ExePath = exePath;
        RegistryVersion = registryVersion;
        Is64Bit = is64Bit;
        ProductVersion = productVersion;
    }

    public string ExePath { get; }
    /// <summary><c>16.0</c> for Excel 2016/2019/2021/2024/365, <c>15.0</c> for 2013, <c>14.0</c> for 2010.</summary>
    public string RegistryVersion { get; }
    public bool Is64Bit { get; }
    public string ProductVersion { get; }

    public override string ToString() => $"Excel {RegistryVersion} ({(Is64Bit ? "64" : "32")}-bit), {ProductVersion}";
}

internal static class ExcelLocator
{
    /// <summary>Finds EXCEL.EXE via App Paths, then Click-to-Run, then MSI InstallRoot; reads version and bitness from the file.</summary>
    public static ExcelInstall? Find()
    {
        foreach (var candidate in Candidates())
        {
            try
            {
                string path = Environment.ExpandEnvironmentVariables(candidate.Trim().Trim('"'));
                if (path.Length == 0 || !File.Exists(path)) continue;
                var info = Describe(path);
                if (info != null) return info;
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is ArgumentException || ex is NotSupportedException)
            {
            }
        }
        return null;
    }

    public static bool IsRunning() => Process.GetProcessesByName("EXCEL").Length > 0;

    private static ExcelInstall? Describe(string exe)
    {
        var header = new byte[4096];
        int read;
        using (var fs = new FileStream(exe, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            read = fs.Read(header, 0, header.Length);
        if (read < header.Length) Array.Resize(ref header, read);

        ushort? machine = PeFile.Machine(header);
        if (machine == null) return null;

        var fvi = FileVersionInfo.GetVersionInfo(exe);
        int major = fvi.FileMajorPart;
        if (major < 14) return null;
        string folder = major >= 16 ? "16.0" : major + ".0";
        return new ExcelInstall(exe, folder, PeFile.Is64Bit(machine.Value), fvi.FileVersion ?? major + ".0");
    }

    private static IEnumerable<string> Candidates()
    {
        foreach (var hive in new[] { RegistryHive.LocalMachine, RegistryHive.CurrentUser })
            foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
                yield return Read(hive, view, @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe", "");

        foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
        {
            string c2r = Read(RegistryHive.LocalMachine, view, @"SOFTWARE\Microsoft\Office\ClickToRun\Configuration", "InstallationPath");
            if (c2r.Length > 0) yield return Path.Combine(c2r, "root", "Office16", "EXCEL.EXE");
        }

        foreach (var ver in new[] { "16.0", "15.0", "14.0" })
            foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
            {
                string root = Read(RegistryHive.LocalMachine, view, $@"SOFTWARE\Microsoft\Office\{ver}\Excel\InstallRoot", "Path");
                if (root.Length > 0) yield return Path.Combine(root, "EXCEL.EXE");
            }
    }

    private static string Read(RegistryHive hive, RegistryView view, string subKey, string valueName)
    {
        try
        {
            using var baseKey = RegistryKey.OpenBaseKey(hive, view);
            using var key = baseKey.OpenSubKey(subKey);
            return key?.GetValue(valueName) as string ?? string.Empty;
        }
        catch (Exception ex) when (ex is System.Security.SecurityException || ex is UnauthorizedAccessException || ex is IOException)
        {
            return string.Empty;
        }
    }
}
