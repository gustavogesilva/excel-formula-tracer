using System.Text;
using Tracer.Core.Model;

namespace Tracer.Core.Text;

/// <summary>A1/R1C1 address formatting.</summary>
public static class A1
{
    public static string Column(int col)
    {
        var buf = new char[3];
        int k = 3;
        while (col > 0)
        {
            int m = (col - 1) % 26;
            buf[--k] = (char)('A' + m);
            col = (col - 1) / 26;
        }
        return new string(buf, k, 3 - k);
    }

    /// <summary>Plain address of a concrete rectangle, e.g. <c>$B$2:$C$9</c>, <c>$B:$B</c> or <c>$3:$3</c>.</summary>
    public static string Address(CellRect r, bool absolute = true)
    {
        string d = absolute ? "$" : "";
        if (r.IsWholeColumns && !r.IsWholeRows)
            return $"{d}{Column(r.Col1)}:{d}{Column(r.Col2)}";
        if (r.IsWholeRows && !r.IsWholeColumns)
            return $"{d}{r.Row1}:{d}{r.Row2}";
        string a = $"{d}{Column(r.Col1)}{d}{r.Row1}";
        return r.IsSingleCell ? a : $"{a}:{d}{Column(r.Col2)}{d}{r.Row2}";
    }

    /// <summary>
    /// Renders a formula reference in A1 notation as seen from the cell at (row, col), keeping <c>$</c>
    /// on absolute bounds, e.g. <c>R[-1]C2</c> at B5 gives <c>$B4</c>.
    /// </summary>
    public static string Render(Area a, int row, int col)
    {
        var sb = new StringBuilder(16);
        switch (a.Kind)
        {
            case AreaKind.Rows:
                AppendRow(sb, a.Row1, row);
                sb.Append(':');
                AppendRow(sb, a.Row2, row);
                break;
            case AreaKind.Cols:
                AppendCol(sb, a.Col1, col);
                sb.Append(':');
                AppendCol(sb, a.Col2, col);
                break;
            default:
                AppendCol(sb, a.Col1, col);
                AppendRow(sb, a.Row1, row);
                if (!a.IsSingleCell)
                {
                    sb.Append(':');
                    AppendCol(sb, a.Col2, col);
                    AppendRow(sb, a.Row2, row);
                }
                break;
        }
        return sb.ToString();
    }

    private static void AppendRow(StringBuilder sb, Bound b, int row)
    {
        if (!b.IsRelative) sb.Append('$');
        sb.Append(b.Resolve(row, Limits.MaxRows));
    }

    private static void AppendCol(StringBuilder sb, Bound b, int col)
    {
        if (!b.IsRelative) sb.Append('$');
        sb.Append(Column(b.Resolve(col, Limits.MaxCols)));
    }

    /// <summary>R1C1 text of an area (for diagnostics).</summary>
    public static string R1C1(Area a)
    {
        string B(char axis, Bound b) => b.IsRelative ? (b.Value == 0 ? axis.ToString() : $"{axis}[{b.Value}]") : $"{axis}{b.Value}";
        return a.Kind switch
        {
            AreaKind.Rows => a.Row1.Equals(a.Row2) ? B('R', a.Row1) : $"{B('R', a.Row1)}:{B('R', a.Row2)}",
            AreaKind.Cols => a.Col1.Equals(a.Col2) ? B('C', a.Col1) : $"{B('C', a.Col1)}:{B('C', a.Col2)}",
            _ => a.IsSingleCell ? B('R', a.Row1) + B('C', a.Col1) : $"{B('R', a.Row1)}{B('C', a.Col1)}:{B('R', a.Row2)}{B('C', a.Col2)}",
        };
    }

    /// <summary>Quotes a sheet name when Excel would, e.g. <c>'My Sheet'</c>, <c>'2024'</c>, <c>'It''s'</c>.</summary>
    public static string QuoteSheet(string sheet)
    {
        bool needs = sheet.Length == 0 || char.IsDigit(sheet[0]) || LooksLikeReference(sheet);
        if (!needs)
        {
            foreach (char c in sheet)
            {
                if (!(char.IsLetterOrDigit(c) || c == '_' || c == '.')) { needs = true; break; }
            }
        }
        return needs ? "'" + sheet.Replace("'", "''") + "'" : sheet;
    }

    private static bool LooksLikeReference(string s)
    {
        // A1-like (letters then digits, e.g. "AB12") or R1C1-like ("R", "C", "R1C1", "RC2").
        int k = 0;
        while (k < s.Length && char.IsLetter(s[k]) && s[k] < 128) k++;
        if (k > 0 && k <= 3 && k < s.Length)
        {
            int m = k;
            while (m < s.Length && char.IsDigit(s[m])) m++;
            if (m == s.Length) return true;
        }
        string u = s.ToUpperInvariant();
        if (u == "R" || u == "C") return true;
        if (u[0] == 'R' || u[0] == 'C')
        {
            int j = 0;
            if (u[j] == 'R') { j++; while (j < u.Length && char.IsDigit(u[j])) j++; }
            if (j < u.Length && u[j] == 'C') { j++; while (j < u.Length && char.IsDigit(u[j])) j++; }
            if (j == u.Length) return true;
        }
        return false;
    }

    /// <summary>Full display address: <c>[Book.xlsx]'My Sheet'!$B$2</c>, omitting the book/sheet when they match the context.</summary>
    public static string Qualified(string? book, string sheet, CellRect r, string? contextBook = null, string? contextSheet = null)
    {
        bool sameBook = book == null || contextBook != null && string.Equals(book, contextBook, System.StringComparison.OrdinalIgnoreCase);
        bool sameSheet = sameBook && contextSheet != null && string.Equals(sheet, contextSheet, System.StringComparison.OrdinalIgnoreCase);
        if (sameSheet) return Address(r);
        if (sameBook) return QuoteSheet(sheet) + "!" + Address(r);
        string inner = "[" + book + "]" + sheet;
        bool quote = QuoteSheet(sheet) != sheet || book!.IndexOfAny(new[] { ' ', '-', '(', ')', '\'' }) >= 0;
        return (quote ? "'" + inner.Replace("'", "''") + "'" : inner) + "!" + Address(r);
    }
}
