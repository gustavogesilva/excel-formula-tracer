using System.Linq;
using Tracer.Core.Model;
using Tracer.Core.Parsing;
using Tracer.Core.Text;
using Xunit;

namespace Tracer.Core.Tests;

public class ParserTests
{
    private static ParsedTemplate P(string f) => TemplateParser.Parse(f);

    /// <summary>Refs as "prefix!R1C1" strings in order, for compact assertions.</summary>
    private static string[] Refs(ParsedTemplate t) =>
        t.Refs.Select(r => (r.PrefixIndex >= 0 ? t.Prefixes[r.PrefixIndex].ToString() : "")
                           + ((r.Flags & RefFlags.Broken) != 0 ? "#REF!" : A1.R1C1(r.Area))
                           + ((r.Flags & RefFlags.Spill) != 0 ? "#" : "")).ToArray();

    private static string[] Names(ParsedTemplate t) => t.Names.Select(n => n.Name).ToArray();

    [Theory]
    [InlineData("=R[-1]C+R1C5*2", new[] { "R[-1]C", "R1C5" })]
    [InlineData("=RC[-1]", new[] { "RC[-1]" })]
    [InlineData("=SUM(R5C2:R5C)", new[] { "R5C2:R5C" })]
    [InlineData("=SUM(C[1]:C[2])+SUM(R1:R3)", new[] { "C[1]:C[2]", "R1:R3" })]
    [InlineData("=R1C1 R2C2:R3C3", new[] { "R1C1", "R2C2:R3C3" })]
    [InlineData("=(R1C1,R2C2)", new[] { "R1C1", "R2C2" })]
    [InlineData("=R1C1:INDEX(C1,5)", new[] { "R1C1", "C1" })]
    [InlineData("=_xlfn.XLOOKUP(R1C1,C2,C3)", new[] { "R1C1", "C2", "C3" })]
    [InlineData("=R[1048575]C", new[] { "R[1048575]C" })]
    [InlineData("=R", new[] { "R" })]
    [InlineData("=C", new[] { "C" })]
    public void SameSheetReferences(string formula, string[] expected) => Assert.Equal(expected, Refs(P(formula)));

    [Theory]
    [InlineData("=Rate*2", new[] { "Rate" })]
    [InlineData("=R2D2+CALC+RC2D", new[] { "R2D2", "CALC", "RC2D" })]
    [InlineData("=Revenue-Costs", new[] { "Revenue", "Costs" })]
    [InlineData("=TRUE+FALSE", new string[0])]
    public void NamesThatStartLikeReferencesAreNames(string formula, string[] names)
    {
        var t = P(formula);
        Assert.Empty(t.Refs);
        Assert.Equal(names, Names(t));
    }

    [Theory]
    [InlineData("='My Sheet'!R2C3+Sheet2!R[1]C", new[] { "My Sheet!R2C3", "Sheet2!R[1]C" })]
    [InlineData("='It''s'!R1C1", new[] { "It's!R1C1" })]
    [InlineData("=Sheet1:Sheet3!R1C1", new[] { "Sheet1:Sheet3!R1C1" })]
    [InlineData("='Jan 24:Dec 24'!R1C1:R2C2", new[] { "Jan 24:Dec 24!R1C1:R2C2" })]
    [InlineData("='[Book2.xlsx]Sheet1'!R1C1", new[] { "[Book2.xlsx]Sheet1!R1C1" })]
    [InlineData("=[Book2.xlsx]Sheet1!R1C1:R5C5", new[] { "[Book2.xlsx]Sheet1!R1C1:R5C5" })]
    [InlineData("='C:\\dir\\[Book 2.xlsx]Sheet1'!R1C1", new[] { "C:\\dir\\[Book 2.xlsx]Sheet1!R1C1" })]
    [InlineData("=Rates!R1C1", new[] { "Rates!R1C1" })]
    [InlineData("=Sheet1!#REF!+1", new[] { "Sheet1!#REF!" })]
    public void QualifiedReferences(string formula, string[] expected) => Assert.Equal(expected, Refs(P(formula)));

    [Fact]
    public void ExternalPrefixParts()
    {
        var t = P("='C:\\dir\\[Book 2.xlsx]Jan:Mar'!R1C1");
        var p = t.Prefixes[t.Refs[0].PrefixIndex];
        Assert.Equal("Book 2.xlsx", p.Book);
        Assert.Equal("C:\\dir\\", p.Path);
        Assert.Equal("Jan", p.Sheet1);
        Assert.Equal("Mar", p.Sheet2);
    }

    [Fact]
    public void ExternalNames()
    {
        var t = P("=[Book2.xlsx]!MyName+Book3.xlsx!Other+'C:\\x\\Book4.xlsx'!Third");
        Assert.Equal(new[] { "MyName", "Other", "Third" }, Names(t));
        Assert.Equal("Book2.xlsx", t.Prefixes[t.Names[0].PrefixIndex].Book);
        Assert.Equal("Book3.xlsx", t.Prefixes[t.Names[1].PrefixIndex].Sheet1);
        Assert.Equal("Book4.xlsx", t.Prefixes[t.Names[2].PrefixIndex].Book);
        Assert.Equal("C:\\x\\", t.Prefixes[t.Names[2].PrefixIndex].Path);
    }

    [Fact]
    public void SheetScopedName()
    {
        var t = P("=Sheet1!Rate+Rate");
        Assert.Equal(new[] { "Rate", "Rate" }, Names(t));
        Assert.Equal("Sheet1", t.Prefixes[t.Names[0].PrefixIndex].Sheet1);
        Assert.Equal(-1, t.Names[1].PrefixIndex);
    }

    [Fact]
    public void StringsAndArraysAreIgnored()
    {
        var t = P("=IF(R1C1>0,\"=R2C2 \"\"Rate\"\"\",{1,2;3,\"R9C9\"})");
        Assert.Equal(new[] { "R1C1" }, Refs(t));
        Assert.Empty(t.Names);
        Assert.Single(t.NumberStarts); // the 0; array constants are skipped
    }

    [Fact]
    public void SpillAndImplicitIntersection()
    {
        Assert.Equal(new[] { "R1C1#" }, Refs(P("=R1C1#")));
        Assert.Equal(new[] { "R1C1:R2C2" }, Refs(P("=@R1C1:R2C2")));
        Assert.Equal(new[] { "R1C1" }, Refs(P("=R1C1+#N/A")));
    }

    [Fact]
    public void LetLocalsAreNotNames()
    {
        var t = P("=LET(x,R1C1,y,x*2,y+Rate)");
        Assert.Equal(new[] { "R1C1" }, Refs(t));
        Assert.Equal(new[] { "Rate" }, Names(t));
    }

    [Fact]
    public void LetFinalExpressionCanBeAName()
    {
        var t = P("=LET(x,1,Rate)");
        Assert.Equal(new[] { "Rate" }, Names(t));
    }

    [Fact]
    public void LambdaParamsAndInvocation()
    {
        var t = P("=LAMBDA(a,b,a+b+Rate)(R1C1,2)");
        Assert.Equal(new[] { "R1C1" }, Refs(t));
        Assert.Equal(new[] { "Rate" }, Names(t));
    }

    [Fact]
    public void LambdaOptionalParam()
    {
        var t = P("=LAMBDA(a,[b],IF(ISOMITTED(b),a,a+b))(1)");
        Assert.Empty(t.Names);
        Assert.Empty(t.StructRefs);
    }

    [Fact]
    public void LocalLambdaCallIsNotRecordedAsCall()
    {
        var t = P("=LET(f,LAMBDA(x,x*2),f(R1C1))");
        Assert.Equal(new[] { "R1C1" }, Refs(t));
        Assert.Empty(t.Names);
        Assert.DoesNotContain(t.Calls, c => c.Name == "F");
    }

    [Fact]
    public void MapWithLambda()
    {
        var t = P("=MAP(R1C1:R5C1,LAMBDA(v,v*Rate))");
        Assert.Equal(new[] { "R1C1:R5C1" }, Refs(t));
        Assert.Equal(new[] { "Rate" }, Names(t));
    }

    [Fact]
    public void XlpmPrefixedLetNames()
    {
        var t = P("=_xlfn.LET(_xlpm.x,R1C1,_xlpm.x+1)");
        Assert.Equal(new[] { "R1C1" }, Refs(t));
        Assert.Empty(t.Names);
    }

    [Fact]
    public void DynamicCalls()
    {
        var t = P("=OFFSET(R1C1,R2C2,0)+INDIRECT(\"A1\")");
        Assert.True(t.HasDynamic);
        Assert.Equal(RefFlags.InDynamicCall, t.Refs[0].Flags);
        Assert.Equal(RefFlags.None, t.Refs[1].Flags); // OFFSET's rows argument is a real precedent
        var offset = t.Calls.Single(c => c.Name == "OFFSET");
        Assert.Equal("OFFSET(R1C1,R2C2,0)", t.Formula.Substring(offset.Start, offset.Length));
    }

    [Fact]
    public void CallsAreNormalized()
    {
        var t = P("=_xlfn._xlws.SORT(R1C1:R9C1)+MyLambda(1)");
        Assert.Contains(t.Calls, c => c.Name == "SORT");
        Assert.Contains(t.Calls, c => c.Name == "MYLAMBDA");
    }

    [Theory]
    [InlineData("=Table1[Col A]", "Table1", TableArea.Data, "Col A", "Col A")]
    [InlineData("=Table1[@Price]", "Table1", TableArea.ThisRow, "Price", "Price")]
    [InlineData("=Table1[@[Unit Price]]", "Table1", TableArea.ThisRow, "Unit Price", "Unit Price")]
    [InlineData("=Table1[[#Headers],[X]:[Y]]", "Table1", TableArea.Headers, "X", "Y")]
    [InlineData("=Table1[[#This Row],[Qty]]", "Table1", TableArea.ThisRow, "Qty", "Qty")]
    [InlineData("=Table1[#All]", "Table1", TableArea.All, null, null)]
    [InlineData("=Table1[]", "Table1", TableArea.Data, null, null)]
    [InlineData("=Table1[[#Headers],[#Data],[Amt]]", "Table1", TableArea.Headers | TableArea.Data, "Amt", "Amt")]
    [InlineData("=Table1['[Odd']]", "Table1", TableArea.Data, "[Odd]", "[Odd]")]
    [InlineData("=[@Qty]*2", null, TableArea.ThisRow, "Qty", "Qty")]
    [InlineData("=SUM([Amt])", null, TableArea.Data, "Amt", "Amt")]
    public void StructuredReferences(string formula, string? table, TableArea area, string? first, string? last)
    {
        var s = Assert.Single(P(formula).StructRefs);
        Assert.Equal(table, s.Table);
        Assert.Equal(area, s.Area);
        Assert.Equal(first, s.FirstColumn);
        Assert.Equal(last, s.LastColumn);
    }

    [Fact]
    public void HardCodedNumbers()
    {
        var t = P("=R[-1]C*1.05+2E-3+.5");
        Assert.Equal(3, t.NumberStarts.Length);
    }

    [Theory]
    [InlineData("")]
    [InlineData("=")]
    [InlineData("=SUM(")]
    [InlineData("=\"unterminated")]
    [InlineData("='unterminated")]
    [InlineData("=Table1[unterminated")]
    [InlineData("=R[abc]C")]
    [InlineData("=R1048577C1")]
    [InlineData("=))((,,")]
    [InlineData("=[")]
    public void MalformedInputDoesNotThrow(string formula) => TemplateParser.Parse(formula);

    [Fact]
    public void SpansCoverTokens()
    {
        var t = P("=SUM('My Sheet'!R1C1:R2C2)+Rate");
        Assert.Equal("'My Sheet'!R1C1:R2C2", t.Formula.Substring(t.Refs[0].Start, t.Refs[0].Length));
        Assert.Equal("R1C1:R2C2", t.Formula.Substring(t.Refs[0].AreaStart, t.Refs[0].End - t.Refs[0].AreaStart));
        Assert.Equal("Rate", t.Formula.Substring(t.Names[0].Start, t.Names[0].Length));
    }
}
