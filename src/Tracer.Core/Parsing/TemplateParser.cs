using System;
using System.Collections.Generic;
using Tracer.Core.Model;

namespace Tracer.Core.Parsing;

/// <summary>
/// Single-pass, reference-only scanner for en-US R1C1 formulas as returned by <c>Range.Formula2R1C1</c>.
/// It builds no AST: it records references, names, structured references, calls and number literals with their spans.
/// It never throws; malformed input is skipped character by character.
/// </summary>
public static class TemplateParser
{
    public static ParsedTemplate Parse(string formula) => new Scanner(formula ?? string.Empty).Run();

    private sealed class Frame
    {
        public string Name = string.Empty;
        public int Start;
        public int ArgIndex;
        public int ArgTokens;
        public bool Record;
        public List<string>? Locals;
        public bool IsLet => Name == "LET";
        public bool IsLambda => Name == "LAMBDA";
    }

    private sealed class Scanner
    {
        private readonly string _s;
        private readonly int _n;
        private int _i;

        private List<RefToken>? _refs;
        private List<NameToken>? _names;
        private List<StructToken>? _structs;
        private List<CallToken>? _calls;
        private List<Prefix>? _prefixes;
        private List<int>? _numbers;
        private readonly List<Frame> _stack = new();

        public Scanner(string s)
        {
            _s = s;
            _n = s.Length;
        }

        private char At(int k) => k < _n ? _s[k] : '\0';
        private Frame? Top => _stack.Count > 0 ? _stack[_stack.Count - 1] : null;

        public ParsedTemplate Run()
        {
            _i = _n > 0 && _s[0] == '=' ? 1 : 0;
            while (_i < _n)
            {
                char c = _s[_i];
                if (IsSpace(c)) { _i++; continue; }

                var top = Top;
                if (c == ',')
                {
                    if (top != null) { top.ArgIndex++; top.ArgTokens = 0; }
                    _i++;
                    continue;
                }
                if (c == ')')
                {
                    PopFrame();
                    _i++;
                    continue;
                }

                bool firstInArg = top != null && top.ArgTokens == 0;
                if (top != null) top.ArgTokens++;

                switch (c)
                {
                    case '"': SkipString(); break;
                    case '{': SkipArray(); break;
                    case '(': PushFrame(string.Empty, _i, record: false); _i++; break;
                    case '#': ScanHash(); break;
                    case '\'': ScanQuotedPrefix(); break;
                    case '[': ScanBracket(firstInArg); break;
                    default:
                        if (IsDigit(c) || (c == '.' && IsDigit(At(_i + 1)))) ScanNumber();
                        else if (IsIdentStart(c)) ScanIdentifier(firstInArg);
                        else _i++;
                        break;
                }
            }
            return Build();
        }

        // ---- top-level token kinds ------------------------------------------------------------

        private void ScanIdentifier(bool firstInArg)
        {
            int start = _i;
            if (TryRef(start, -1, start)) return;

            int end = ReadIdent(start);
            char next = At(end);

            if (next == '!')
            {
                int p = AddPrefix(new Prefix(null, null, _s.Substring(start, end - start), null));
                _i = end + 1;
                ScanAfterPrefix(p, start);
                return;
            }
            if (next == ':' && IsIdentStart(At(end + 1)))
            {
                int j = end + 1;
                int k = ReadIdent(j);
                if (At(k) == '!')
                {
                    int p = AddPrefix(new Prefix(null, null, _s.Substring(start, end - start), _s.Substring(j, k - j)));
                    _i = k + 1;
                    ScanAfterPrefix(p, start);
                    return;
                }
            }

            string ident = _s.Substring(start, end - start);
            if (next == '(')
            {
                bool local = IsLocal(ident);
                PushFrame(local ? string.Empty : NormalizeFunction(ident), start, record: !local);
                _i = end + 1;
                return;
            }
            if (next == '[')
            {
                _i = end;
                ScanStruct(start, ident, -1);
                return;
            }

            _i = end;
            if (ident.Equals("TRUE", StringComparison.OrdinalIgnoreCase) || ident.Equals("FALSE", StringComparison.OrdinalIgnoreCase))
                return;

            var top = Top;
            if (firstInArg && top != null && (top.IsLambda || (top.IsLet && top.ArgIndex % 2 == 0)) && NextNonSpace(end) == ',')
            {
                (top.Locals ??= new List<string>()).Add(StripParamPrefix(ident));
                return;
            }
            if (IsLocal(ident)) return;

            (_names ??= new List<NameToken>()).Add(new NameToken(ident, -1, start, end - start, isCall: false));
        }

        /// <summary>After <c>Prefix!</c>: a reference, <c>#REF!</c>, a name, a LAMBDA-name call, or a table.</summary>
        private void ScanAfterPrefix(int prefixIndex, int tokenStart)
        {
            char c = At(_i);
            if (TryRef(_i, prefixIndex, tokenStart)) return;

            if (c == '#' && string.Compare(_s, _i, "#REF!", 0, 5, StringComparison.OrdinalIgnoreCase) == 0)
            {
                _i += 5;
                AddRef(new RefToken(default, prefixIndex, tokenStart, _i - tokenStart, tokenStart, RefFlags.Broken));
                return;
            }
            if (!IsIdentStart(c)) return;

            int start = _i;
            int end = ReadIdent(start);
            string ident = _s.Substring(start, end - start);
            char next = At(end);
            if (next == '[')
            {
                _i = end;
                ScanStruct(tokenStart, ident, prefixIndex);
                return;
            }
            bool isCall = next == '(';
            (_names ??= new List<NameToken>()).Add(new NameToken(ident, prefixIndex, tokenStart, end - tokenStart, isCall));
            _i = end;
            if (isCall)
            {
                PushFrame(string.Empty, tokenStart, record: false);
                _i++;
            }
        }

        private void ScanQuotedPrefix()
        {
            int start = _i;
            int j = _i + 1;
            var sb = new System.Text.StringBuilder();
            while (j < _n)
            {
                char c = _s[j];
                if (c == '\'')
                {
                    if (At(j + 1) == '\'') { sb.Append('\''); j += 2; continue; }
                    break;
                }
                sb.Append(c);
                j++;
            }
            if (j >= _n || At(j + 1) != '!')
            {
                _i = Math.Min(j + 1, _n);
                return;
            }
            int p = AddPrefix(SplitPrefix(sb.ToString()));
            _i = j + 2;
            ScanAfterPrefix(p, start);
        }

        /// <summary><c>[</c> starts an unquoted external prefix (<c>[Book.xlsx]Sheet!</c>), a LAMBDA optional parameter, or a structured reference.</summary>
        private void ScanBracket(bool firstInArg)
        {
            int start = _i;
            char c1 = At(_i + 1);
            if (c1 != '[' && c1 != '@' && c1 != '#')
            {
                int close = _s.IndexOf(']', _i + 1);
                if (close > 0)
                {
                    int j = close + 1;
                    int k = IsIdentStart(At(j)) ? ReadIdent(j) : j;
                    string? sheet1 = k > j ? _s.Substring(j, k - j) : null;
                    string? sheet2 = null;
                    if (At(k) == ':' && IsIdentStart(At(k + 1)))
                    {
                        int m = ReadIdent(k + 1);
                        sheet2 = _s.Substring(k + 1, m - k - 1);
                        k = m;
                    }
                    if (At(k) == '!')
                    {
                        string book = _s.Substring(_i + 1, close - _i - 1);
                        int p = AddPrefix(new Prefix(book, null, sheet1, sheet2));
                        _i = k + 1;
                        ScanAfterPrefix(p, start);
                        return;
                    }

                    var top = Top;
                    if (firstInArg && top != null && top.IsLambda && sheet1 == null && NextNonSpace(close + 1) == ',')
                    {
                        (top.Locals ??= new List<string>()).Add(StripParamPrefix(_s.Substring(_i + 1, close - _i - 1).Trim()));
                        _i = close + 1;
                        return;
                    }
                }
            }
            ScanStruct(start, null, -1);
        }

        private void ScanHash()
        {
            foreach (var e in ErrorLiterals)
            {
                if (string.Compare(_s, _i, e, 0, e.Length, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    if (e == "#REF!")
                        AddRef(new RefToken(default, -1, _i, e.Length, _i, RefFlags.Broken));
                    _i += e.Length;
                    return;
                }
            }
            _i++;
        }

        private void ScanNumber()
        {
            int start = _i;
            while (IsDigit(At(_i))) _i++;
            if (At(_i) == '.') { _i++; while (IsDigit(At(_i))) _i++; }
            if ((At(_i) == 'E' || At(_i) == 'e') && (IsDigit(At(_i + 1)) || ((At(_i + 1) == '+' || At(_i + 1) == '-') && IsDigit(At(_i + 2)))))
            {
                _i += 2;
                while (IsDigit(At(_i))) _i++;
            }
            (_numbers ??= new List<int>()).Add(start);
        }

        private void SkipString()
        {
            _i++;
            while (_i < _n)
            {
                if (_s[_i] == '"')
                {
                    if (At(_i + 1) == '"') { _i += 2; continue; }
                    _i++;
                    return;
                }
                _i++;
            }
        }

        private void SkipArray()
        {
            _i++;
            while (_i < _n && _s[_i] != '}')
            {
                if (_s[_i] == '"') SkipString();
                else _i++;
            }
            if (_i < _n) _i++;
        }

        // ---- references ---------------------------------------------------------------------

        /// <summary>Tries an R1C1 reference at <paramref name="pos"/>, merging <c>a:b</c> ranges and spill <c>#</c>.</summary>
        private bool TryRef(int pos, int prefixIndex, int tokenStart)
        {
            if (!TryArea(pos, out var area, out int end) || !IsRefBoundary(end)) return false;

            if (At(end) == ':' && TryArea(end + 1, out var area2, out int end2) && IsRefBoundary(end2)
                && TryMerge(area, area2, out var merged))
            {
                area = merged;
                end = end2;
            }

            var flags = RefFlags.None;
            if (At(end) == '#' && !IsLetter(At(end + 1)))
            {
                flags |= RefFlags.Spill;
                end++;
            }
            var top = Top;
            if (top != null && top.Name == "OFFSET" && top.ArgIndex == 0) flags |= RefFlags.InDynamicCall;

            AddRef(new RefToken(area, prefixIndex, tokenStart, end - tokenStart, pos, flags));
            _i = end;
            return true;
        }

        private bool TryArea(int pos, out Area area, out int end)
        {
            area = default;
            end = pos;
            int j = pos;
            bool hasR = false, hasC = false;
            Bound rb = default, cb = default;

            char c = At(j);
            if (c == 'R' || c == 'r')
            {
                j++;
                if (!TryAxis(ref j, Limits.MaxRows, out rb)) return false;
                hasR = true;
                c = At(j);
            }
            if (c == 'C' || c == 'c')
            {
                j++;
                if (!TryAxis(ref j, Limits.MaxCols, out cb)) return false;
                hasC = true;
            }
            if (!hasR && !hasC) return false;

            area = hasR && hasC ? Area.Cell(rb, cb) : hasR ? Area.Rows(rb, rb) : Area.Cols(cb, cb);
            end = j;
            return true;
        }

        private bool TryAxis(ref int j, int max, out Bound bound)
        {
            bound = Bound.Rel(0);
            char c = At(j);
            if (c == '[')
            {
                int k = j + 1;
                bool neg = false;
                if (At(k) == '-' || At(k) == '+') { neg = At(k) == '-'; k++; }
                if (!IsDigit(At(k))) return false;
                long v = 0;
                while (IsDigit(At(k)) && v <= max) { v = v * 10 + (At(k) - '0'); k++; }
                if (At(k) != ']' || v >= max) return false;
                bound = Bound.Rel(neg ? -(int)v : (int)v);
                j = k + 1;
                return true;
            }
            if (IsDigit(c))
            {
                int k = j;
                long v = 0;
                while (IsDigit(At(k)) && v <= max) { v = v * 10 + (At(k) - '0'); k++; }
                if (v < 1 || v > max) return false;
                bound = Bound.Abs((int)v);
                j = k;
                return true;
            }
            return true;
        }

        private bool IsRefBoundary(int end)
        {
            char c = At(end);
            return !IsIdentChar(c) && c != '(' && c != '!' && c != '[';
        }

        private static bool TryMerge(Area a, Area b, out Area merged)
        {
            merged = default;
            if (a.Kind != b.Kind) return false;
            merged = a.Kind switch
            {
                AreaKind.Rows => Area.Rows(a.Row1, b.Row2),
                AreaKind.Cols => Area.Cols(a.Col1, b.Col2),
                _ => new Area(a.Row1, a.Col1, b.Row2, b.Col2),
            };
            return true;
        }

        // ---- structured references ------------------------------------------------------------

        private void ScanStruct(int tokenStart, string? table, int prefixIndex)
        {
            int open = _i;
            int depth = 0;
            int j = open;
            while (j < _n)
            {
                char c = _s[j];
                if (c == '\'') { j += 2; continue; }
                if (c == '[') depth++;
                else if (c == ']' && --depth == 0) break;
                j++;
            }
            if (j >= _n)
            {
                _i = _n;
                return;
            }
            string inner = _s.Substring(open + 1, j - open - 1);
            _i = j + 1;

            ParseStructInner(inner, out var area, out var first, out var last);
            (_structs ??= new List<StructToken>()).Add(new StructToken(table, area, first, last, prefixIndex, tokenStart, _i - tokenStart));
        }

        internal static void ParseStructInner(string inner, out TableArea area, out string? first, out string? last)
        {
            area = TableArea.None;
            first = last = null;
            string t = inner.Trim();
            if (t.Length == 0) { area = TableArea.Data; return; }

            if (t[0] == '@')
            {
                area = TableArea.ThisRow;
                t = t.Substring(1).Trim();
                if (t.Length == 0) return;
                if (t[0] != '[') { first = last = Unescape(t); return; }
            }
            else if (t[0] == '#')
            {
                area = Specifier(t);
                return;
            }
            else if (t[0] != '[')
            {
                area = TableArea.Data;
                first = last = Unescape(t);
                return;
            }

            // List of [item] / [item]:[item] separated by commas.
            int k = 0;
            while (k < t.Length)
            {
                if (t[k] != '[') { k++; continue; }
                int close = FindClose(t, k);
                string item = t.Substring(k + 1, close - k - 1);
                k = close + 1;
                if (item.StartsWith("#", StringComparison.Ordinal))
                {
                    area |= Specifier(item);
                    continue;
                }
                string col = Unescape(item);
                first ??= col;
                last = col;
                int m = k;
                while (m < t.Length && t[m] == ' ') m++;
                if (m < t.Length && t[m] == ':')
                {
                    m++;
                    while (m < t.Length && t[m] == ' ') m++;
                    if (m < t.Length && t[m] == '[')
                    {
                        int close2 = FindClose(t, m);
                        last = Unescape(t.Substring(m + 1, close2 - m - 1));
                        k = close2 + 1;
                    }
                }
            }
            if (area == TableArea.None) area = TableArea.Data;
        }

        private static int FindClose(string t, int open)
        {
            for (int j = open + 1; j < t.Length; j++)
            {
                if (t[j] == '\'') { j++; continue; }
                if (t[j] == ']') return j;
            }
            return t.Length;
        }

        private static TableArea Specifier(string s)
        {
            s = s.Trim();
            if (s.Equals("#All", StringComparison.OrdinalIgnoreCase)) return TableArea.All;
            if (s.Equals("#Data", StringComparison.OrdinalIgnoreCase)) return TableArea.Data;
            if (s.Equals("#Headers", StringComparison.OrdinalIgnoreCase)) return TableArea.Headers;
            if (s.Equals("#Totals", StringComparison.OrdinalIgnoreCase)) return TableArea.Totals;
            if (s.Equals("#This Row", StringComparison.OrdinalIgnoreCase)) return TableArea.ThisRow;
            return TableArea.Data;
        }

        private static string Unescape(string s)
        {
            if (s.IndexOf('\'') < 0) return s.Trim();
            var sb = new System.Text.StringBuilder(s.Length);
            for (int j = 0; j < s.Length; j++)
            {
                if (s[j] == '\'' && j + 1 < s.Length) j++;
                sb.Append(s[j]);
            }
            return sb.ToString().Trim();
        }

        // ---- frames (calls, LET/LAMBDA scopes) -------------------------------------------------

        private void PushFrame(string name, int start, bool record) =>
            _stack.Add(new Frame { Name = name, Start = start, Record = record && name.Length > 0 });

        private void PopFrame()
        {
            if (_stack.Count == 0) return;
            var f = _stack[_stack.Count - 1];
            _stack.RemoveAt(_stack.Count - 1);
            if (f.Record)
                (_calls ??= new List<CallToken>()).Add(new CallToken(f.Name, f.Start, _i + 1 - f.Start));
        }

        private bool IsLocal(string ident)
        {
            string name = StripParamPrefix(ident);
            for (int k = _stack.Count - 1; k >= 0; k--)
            {
                var locals = _stack[k].Locals;
                if (locals == null) continue;
                foreach (var l in locals)
                    if (l.Equals(name, StringComparison.OrdinalIgnoreCase)) return true;
            }
            return false;
        }

        private static string StripParamPrefix(string ident) =>
            ident.StartsWith("_xlpm.", StringComparison.OrdinalIgnoreCase) ? ident.Substring(6) : ident;

        internal static string NormalizeFunction(string ident)
        {
            string s = ident;
            while (true)
            {
                if (s.StartsWith("_xlfn.", StringComparison.OrdinalIgnoreCase)) s = s.Substring(6);
                else if (s.StartsWith("_xlws.", StringComparison.OrdinalIgnoreCase)) s = s.Substring(6);
                else break;
            }
            return s.ToUpperInvariant();
        }

        // ---- helpers ---------------------------------------------------------------------

        private int AddPrefix(Prefix p)
        {
            _prefixes ??= new List<Prefix>();
            _prefixes.Add(p);
            return _prefixes.Count - 1;
        }

        private void AddRef(RefToken r) => (_refs ??= new List<RefToken>()).Add(r);

        /// <summary>Splits the text of a quoted prefix: <c>C:\dir\[Book.xlsx]S1:S3</c>, <c>[Book.xlsx]S1</c>, <c>S1</c>, or <c>C:\dir\Book.xlsx</c>.</summary>
        internal static Prefix SplitPrefix(string content)
        {
            string? book = null, path = null;
            string sheetPart = content;
            int lb = content.IndexOf('[');
            int rb = lb >= 0 ? content.IndexOf(']', lb + 1) : -1;
            if (lb >= 0 && rb > lb)
            {
                path = lb > 0 ? content.Substring(0, lb) : null;
                book = content.Substring(lb + 1, rb - lb - 1);
                sheetPart = content.Substring(rb + 1);
            }
            else
            {
                int slash = Math.Max(content.LastIndexOf('\\'), content.LastIndexOf('/'));
                if (slash >= 0)
                    return new Prefix(content.Substring(slash + 1), content.Substring(0, slash + 1), null, null);
            }
            if (sheetPart.Length == 0) return new Prefix(book, path, null, null);
            int colon = sheetPart.IndexOf(':');
            return colon > 0
                ? new Prefix(book, path, sheetPart.Substring(0, colon), sheetPart.Substring(colon + 1))
                : new Prefix(book, path, sheetPart, null);
        }

        private int ReadIdent(int j)
        {
            while (j < _n && IsIdentChar(_s[j])) j++;
            return j;
        }

        private char NextNonSpace(int j)
        {
            while (j < _n && IsSpace(_s[j])) j++;
            return At(j);
        }

        private ParsedTemplate Build() => new(
            _s,
            _refs?.ToArray() ?? Array.Empty<RefToken>(),
            _names?.ToArray() ?? Array.Empty<NameToken>(),
            _structs?.ToArray() ?? Array.Empty<StructToken>(),
            _calls?.ToArray() ?? Array.Empty<CallToken>(),
            _prefixes?.ToArray() ?? Array.Empty<Prefix>(),
            _numbers?.ToArray() ?? Array.Empty<int>());
    }

    private static readonly string[] ErrorLiterals =
    {
        "#GETTING_DATA", "#BLOCKED!", "#CONNECT!", "#UNKNOWN!", "#DIV/0!", "#VALUE!", "#FIELD!", "#SPILL!",
        "#CALC!", "#BUSY!", "#NULL!", "#NAME?", "#NUM!", "#REF!", "#N/A",
    };

    private static bool IsSpace(char c) => c == ' ' || c == '\n' || c == '\r' || c == '\t' || c == '\u00A0';
    private static bool IsDigit(char c) => c >= '0' && c <= '9';
    private static bool IsLetter(char c) => char.IsLetter(c);
    private static bool IsIdentStart(char c) => char.IsLetter(c) || c == '_' || c == '\\';
    private static bool IsIdentChar(char c) => char.IsLetterOrDigit(c) || c == '_' || c == '.' || c == '\\' || c == '?';
}
