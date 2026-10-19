using System;
using System.Collections.Generic;
using System.Text;
using Tracer.Core.Model;

namespace Tracer.Core.Text;

/// <summary>A formula rendered in A1 notation, with a map from R1C1 offsets back to A1 offsets.</summary>
public sealed class RenderedFormula
{
    private readonly List<(int Src, int Dst)> _anchors;

    internal RenderedFormula(string text, List<(int Src, int Dst)> anchors)
    {
        Text = text;
        _anchors = anchors;
    }

    public string Text { get; }

    /// <summary>Maps a token boundary in the R1C1 text to the A1 text.</summary>
    public int Map(int srcIndex)
    {
        int lo = 0, hi = _anchors.Count - 1, best = 0;
        while (lo <= hi)
        {
            int mid = (lo + hi) >> 1;
            if (_anchors[mid].Src <= srcIndex) { best = mid; lo = mid + 1; }
            else hi = mid - 1;
        }
        var a = _anchors[best];
        return a.Dst + (srcIndex - a.Src);
    }

    public (int Start, int Length) MapSpan(int start, int length)
    {
        int s = Map(start);
        return (s, Map(start + length) - s);
    }
}

/// <summary>
/// Converts an R1C1 formula to A1 by replacing only the address part of each reference token.
/// Everything else (functions, names, tables, prefixes) is identical in both notations.
/// </summary>
public static class FormulaRenderer
{
    public static RenderedFormula ToA1(ParsedTemplate t, int row, int col)
    {
        string s = t.Formula;
        var anchors = new List<(int, int)> { (0, 0) };
        var sb = new StringBuilder(s.Length + 8);
        int pos = 0;
        foreach (var r in t.Refs)
        {
            // Broken tokens are rendered only when they keep an address (#REF!R1C1 from a deleted sheet).
            if (((r.Flags & RefFlags.Broken) != 0 && r.AreaStart == r.Start) || r.AreaStart < pos) continue;
            int areaEnd = r.End - ((r.Flags & RefFlags.Spill) != 0 ? 1 : 0);
            sb.Append(s, pos, r.AreaStart - pos);
            anchors.Add((r.AreaStart, sb.Length));
            sb.Append(A1.Render(r.Area, row, col));
            pos = areaEnd;
            anchors.Add((pos, sb.Length));
        }
        sb.Append(s, pos, s.Length - pos);
        return new RenderedFormula(sb.ToString(), anchors);
    }

    /// <summary>A1 text of a sub-expression, e.g. an OFFSET(...) call to hand to <c>Worksheet.Evaluate</c>.</summary>
    public static string SubExpressionA1(ParsedTemplate t, RenderedFormula rendered, int start, int length)
    {
        var (s, len) = rendered.MapSpan(start, length);
        return rendered.Text.Substring(s, Math.Min(len, rendered.Text.Length - s));
    }
}
