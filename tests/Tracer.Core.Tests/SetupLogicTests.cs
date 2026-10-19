using System.Collections.Generic;
using Tracer.Setup;
using Xunit;

namespace Tracer.Core.Tests;

/// <summary>Installer logic that does not touch the registry (compiled in from src/Tracer.Setup).</summary>
public class SetupLogicTests
{
    private const string Ours = @"C:\Users\me\AppData\Local\FormulaTracer\FormulaTracer.xll";

    [Theory]
    [InlineData("OPEN", 0)]
    [InlineData("OPEN1", 1)]
    [InlineData("open12", 12)]
    [InlineData("OPENX", -1)]
    [InlineData("OPEN0", -1)]
    [InlineData("Options", -1)]
    [InlineData("DefSheets", -1)]
    public void ValueNames(string name, int index)
    {
        Assert.Equal(index, OpenValues.IndexOf(name));
        if (index >= 0) Assert.Equal(name.ToUpperInvariant(), OpenValues.NameAt(index));
    }

    [Fact]
    public void OrderedClosesGapsAndIgnoresOtherValues()
    {
        var values = new Dictionary<string, string>
        {
            ["OPEN3"] = "c",
            ["DefaultFormat"] = "51",
            ["OPEN"] = "a",
            ["OPEN1"] = "b",
            ["OPEN7"] = "d",
        };
        Assert.Equal(new[] { "a", "b", "c", "d" }, OpenValues.Ordered(values));
    }

    [Theory]
    [InlineData("/R \"C:\\x\\FormulaTracer.xll\"", "C:\\x\\FormulaTracer.xll")]
    [InlineData("\"C:\\x\\Book Tools.xlam\"", "C:\\x\\Book Tools.xlam")]
    [InlineData("/R /A \"C:\\x\\a.xll\"", "C:\\x\\a.xll")]
    [InlineData("C:\\plain.xll", "C:\\plain.xll")]
    public void PathOf(string data, string path) => Assert.Equal(path, OpenValues.PathOf(data));

    [Theory]
    [InlineData("/R \"C:\\Users\\me\\Downloads\\FormulaTracer64-packed.xll\"", true)]
    [InlineData("/R \"C:\\x\\formulatracer.xll\"", true)]
    [InlineData("C:\\x\\FormulaTracer-packed.xll", true)]
    [InlineData("/R \"C:\\x\\Other.xll\"", false)]
    [InlineData("\"C:\\x\\FormulaTracer.xlam\"", false)]
    [InlineData("/R \"C:\\x<|>\\FormulaTracer.xll\"", true)]
    public void RecognisesOurAddIn(string data, bool ours) => Assert.Equal(ours, OpenValues.IsOurs(data));

    [Fact]
    public void InstallIntoEmptyList()
    {
        Assert.Equal(new[] { "/R \"" + Ours + "\"" }, OpenValues.Install(new string[0], Ours));
    }

    [Fact]
    public void InstallKeepsOthersAndReplacesHandLoadedCopy()
    {
        var current = new[]
        {
            "\"C:\\Tools\\Analysis.xlam\"",
            "/R \"C:\\Users\\me\\Downloads\\FormulaTracer64-packed.xll\"",
            "/R \"C:\\Tools\\Other.xll\"",
        };
        Assert.Equal(new[] { "\"C:\\Tools\\Analysis.xlam\"", "/R \"C:\\Tools\\Other.xll\"", "/R \"" + Ours + "\"" },
            OpenValues.Install(current, Ours));
    }

    [Fact]
    public void ReinstallDoesNotDuplicate()
    {
        var once = OpenValues.Install(new[] { "\"C:\\a.xlam\"" }, Ours);
        Assert.Equal(once, OpenValues.Install(once, Ours));
    }

    [Fact]
    public void UninstallRemovesOursAndKeepsOrder()
    {
        var current = new[] { "\"C:\\a.xlam\"", "/R \"" + Ours + "\"", "/R \"C:\\b.xll\"" };
        Assert.Equal(new[] { "\"C:\\a.xlam\"", "/R \"C:\\b.xll\"" }, OpenValues.Uninstall(current));
    }

    private static byte[] Pe(ushort machine, int peOffset = 0x80)
    {
        var b = new byte[0x200];
        b[0] = (byte)'M';
        b[1] = (byte)'Z';
        System.BitConverter.GetBytes(peOffset).CopyTo(b, 0x3C);
        b[peOffset] = (byte)'P';
        b[peOffset + 1] = (byte)'E';
        System.BitConverter.GetBytes(machine).CopyTo(b, peOffset + 4);
        return b;
    }

    [Fact]
    public void PeMachine()
    {
        Assert.Equal(PeFile.Amd64, PeFile.Machine(Pe(PeFile.Amd64)));
        Assert.Equal(PeFile.I386, PeFile.Machine(Pe(PeFile.I386, 0xF0)));
        Assert.True(PeFile.Is64Bit(PeFile.Amd64));
        Assert.True(PeFile.Is64Bit(PeFile.Arm64));
        Assert.False(PeFile.Is64Bit(PeFile.I386));
    }

    [Fact]
    public void NotAPeFile()
    {
        Assert.Null(PeFile.Machine(new byte[10]));
        var bad = Pe(PeFile.Amd64);
        bad[0x81] = (byte)'X';
        Assert.Null(PeFile.Machine(bad));
        var truncated = Pe(PeFile.Amd64);
        System.BitConverter.GetBytes(0x1FC).CopyTo(truncated, 0x3C); // PE header would run past the bytes read
        Assert.Null(PeFile.Machine(truncated));
    }
}
