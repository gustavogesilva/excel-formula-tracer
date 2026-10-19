using Tracer.Core.Model;
using Tracer.Core.Parsing;
using Tracer.Core.Text;
using Xunit;

namespace Tracer.Core.Tests;

public class RenderTests
{
    [Theory]
    [InlineData(1, "A")]
    [InlineData(26, "Z")]
    [InlineData(27, "AA")]
    [InlineData(702, "ZZ")]
    [InlineData(703, "AAA")]
    [InlineData(16384, "XFD")]
    public void Columns(int col, string expected) => Assert.Equal(expected, A1.Column(col));

    [Theory]
    [InlineData("=R[-1]C+R1C5*2", 5, 2, "=B4+$E$1*2")]
    [InlineData("=SUM(R5C2:R5C)", 5, 4, "=SUM($B$5:D$5)")]
    [InlineData("=SUM(C[1]:C[2])+SUM(R1:R3)", 1, 1, "=SUM(B:C)+SUM($1:$3)")]
    [InlineData("='My Sheet'!R[1]C[1]+Rate", 1, 1, "='My Sheet'!B2+Rate")]
    [InlineData("=R1C1#*2", 3, 3, "=$A$1#*2")]
    [InlineData("=R[-1]C", 1, 1, "=A1048576")]
    [InlineData("=Sheet1!#REF!+R1C1", 1, 1, "=Sheet1!#REF!+$A$1")]
    public void ToA1(string r1c1, int row, int col, string expected)
    {
        var t = TemplateParser.Parse(r1c1);
        Assert.Equal(expected, FormulaRenderer.ToA1(t, row, col).Text);
    }

    [Fact]
    public void SpansMapToA1()
    {
        var t = TemplateParser.Parse("=OFFSET(R[10]C[2],1,1)+Rate");
        var r = FormulaRenderer.ToA1(t, 1, 1);
        Assert.Equal("=OFFSET(C11,1,1)+Rate", r.Text);
        var call = t.Calls[0];
        Assert.Equal("OFFSET(C11,1,1)", FormulaRenderer.SubExpressionA1(t, r, call.Start, call.Length));
        var (s, len) = r.MapSpan(t.Names[0].Start, t.Names[0].Length);
        Assert.Equal("Rate", r.Text.Substring(s, len));
    }

    [Theory]
    [InlineData("Sheet1", "Sheet1")]
    [InlineData("My Sheet", "'My Sheet'")]
    [InlineData("It's", "'It''s'")]
    [InlineData("2024", "'2024'")]
    [InlineData("AB12", "'AB12'")]
    [InlineData("R1C1", "'R1C1'")]
    [InlineData("Data-Raw", "'Data-Raw'")]
    [InlineData("Rates", "Rates")]
    public void SheetQuoting(string sheet, string expected) => Assert.Equal(expected, A1.QuoteSheet(sheet));

    [Fact]
    public void Addresses()
    {
        Assert.Equal("$B$2:$C$9", A1.Address(new CellRect(2, 2, 9, 3)));
        Assert.Equal("$B:$B", A1.Address(new CellRect(1, 2, Limits.MaxRows, 2)));
        Assert.Equal("$3:$4", A1.Address(new CellRect(3, 1, 4, Limits.MaxCols)));
        Assert.Equal("'My Sheet'!$A$1", A1.Qualified(null, "My Sheet", CellRect.Cell(1, 1), "Book.xlsx", "Other"));
        Assert.Equal("[Ext.xlsx]Data!$A$1", A1.Qualified("Ext.xlsx", "Data", CellRect.Cell(1, 1), "Book.xlsx", "Other"));
        Assert.Equal("$A$1", A1.Qualified("Book.xlsx", "S", CellRect.Cell(1, 1), "book.xlsx", "s"));
    }
}
