using System.Linq;
using Tracer.Core.Binding;
using Tracer.Core.Model;
using Tracer.Core.Tracing;
using Xunit;

namespace Tracer.Core.Tests;

public class TraceEngineTests
{
    private const string Book = "Model.xlsx";

    /// <summary>
    /// Inputs!C2 = Rate, Inputs!C3 = Tax (local to Calc), table Sales on Data!B1:E10
    /// (header row 1, data rows 2..10, columns Region/Qty/Price/Total), sheets Jan..Mar for 3D.
    /// </summary>
    private static TraceEngine Engine()
    {
        var e = new TraceEngine();
        e.UpdateBook(new WorkbookModel(Book,
            new[] { "Inputs", "Calc", "Data", "Jan", "Feb", "Mar", "Summary" },
            new[]
            {
                new NameDef("Rate", null, "=Inputs!R2C3"),
                new NameDef("Tax", "Calc", "=Inputs!R3C3"),
                new NameDef("Prev", null, "=R[-1]C"),
                new NameDef("Double", null, "=Rate*2"),
                new NameDef("AddRate", null, "=LAMBDA(x,x+Rate)"),
                new NameDef("Loop", null, "=Loop+1"),
            },
            new[] { new TableDef("Sales", "Data", 1, 2, 10, 0, 2, new[] { "Region", "Qty", "Price", "Total" }) }));
        return e;
    }

    private static string[] Labels(PrecedentResult r) => r.Items.Select(i => $"{i.Kind}:{i.Label}").ToArray();

    [Fact]
    public void PlainPrecedentsInFormulaOrder()
    {
        var r = Engine().Precedents(Book, "Calc", 5, 4, "=R[-1]C+Inputs!R2C3*SUM(R1C1:R3C2)");
        Assert.Equal("=D4+Inputs!$C$2*SUM($A$1:$B$3)", r.Formula.Text);
        Assert.Equal(new[] { "Cell:$D$4", "Cell:Inputs!$C$2", "Range:$A$1:$B$3" }, Labels(r));
        var first = r.Items[0];
        Assert.Equal("D4", r.Formula.Text.Substring(first.SpanStart, first.SpanLength));
    }

    [Fact]
    public void NamesResolveIncludingLocalNestedAndRelative()
    {
        var r = Engine().Precedents(Book, "Calc", 5, 4, "=Rate+Tax+Prev+Double");
        Assert.Equal(new[] { "Name:Rate", "Name:Tax", "Name:Prev", "Name:Double" }, Labels(r));
        Assert.Equal("Inputs!$C$2", r.Items[0].Children!.Single().Label);
        Assert.Equal("Inputs!$C$3", r.Items[1].Children!.Single().Label);
        Assert.Equal("$D$4", r.Items[2].Children!.Single().Label); // relative name floats with the using cell
        Assert.Equal("Inputs!$C$2", r.Items[3].Children!.Single().Label); // Double -> Rate -> Inputs!C2
        Assert.True(r.Items[0].HasTarget);
        Assert.Equal("=Inputs!$C$2", r.Items[0].Expression);
    }

    [Fact]
    public void LocalNameIsInvisibleFromOtherSheets()
    {
        var r = Engine().Precedents(Book, "Summary", 1, 1, "=Tax");
        Assert.Equal(TraceItemKind.Unresolved, r.Items.Single().Kind);
    }

    [Fact]
    public void LambdaNameCallAndSelfReferencingNameTerminate()
    {
        var r = Engine().Precedents(Book, "Calc", 1, 1, "=AddRate(5)+Loop");
        Assert.Contains(r.Items, i => i.Kind == TraceItemKind.Name && i.Label == "AddRate" && i.Children!.Any(c => c.Label == "Inputs!$C$2"));
        Assert.Contains(r.Items, i => i.Label == "Loop");
    }

    [Fact]
    public void StructuredReferences()
    {
        var e = Engine();
        var r = e.Precedents(Book, "Data", 5, 5, "=[@Qty]*Sales[@Price]+SUM(Sales[Qty])+Sales[[#Headers],[Region]]");
        Assert.Equal(new[] { "$C$5", "$D$5", "$C$2:$C$10", "$B$1" },
            r.Items.Select(i => i.Expression).ToArray());
        Assert.All(r.Items, i => Assert.Equal(TraceItemKind.Table, i.Kind));
    }

    [Fact]
    public void ThreeDReferencesExpandPerSheet()
    {
        var r = Engine().Precedents(Book, "Summary", 1, 1, "=SUM(Jan:Mar!R1C1)");
        Assert.Equal(new[] { "Jan!$A$1", "Feb!$A$1", "Mar!$A$1" }, r.Items.Select(i => i.Label).ToArray());
    }

    [Fact]
    public void ExternalAndDynamicAndBroken()
    {
        var r = Engine().Precedents(Book, "Calc", 2, 2, "='C:\\x\\[Ext.xlsx]Data'!R1C1+OFFSET(R1C1,1,1)+#REF!");
        Assert.Equal(TraceItemKind.External, r.Items[0].Kind);
        Assert.False(r.Items[0].HasTarget);
        Assert.Equal("[Ext.xlsx]Data!$A$1", r.Items[0].Label);
        Assert.Contains(r.Items, i => i.Kind == TraceItemKind.Dynamic && i.Expression == "OFFSET($A$1,1,1)");
        Assert.Contains(r.Items, i => i.Kind == TraceItemKind.Cell && i.IsAddressOnly);
        Assert.Contains(r.Items, i => i.Kind == TraceItemKind.Broken);
    }

    [Fact]
    public void DependentsFollowNamesTablesThreeDAndOtherBooks()
    {
        var e = Engine();
        e.UpdateBook(new WorkbookModel("Other.xlsx", new[] { "Sheet1" }));

        void Index(string book, string sheet, params (int r, int c, string f)[] cells)
        {
            var w = e.BeginSheet(book, sheet);
            foreach (var (r, c, f) in cells.OrderBy(x => x.r).ThenBy(x => x.c)) w.Add(r, c, f);
            w.Commit();
        }

        Index(Book, "Calc", (1, 1, "=Rate*2"), (2, 1, "=Double"), (3, 1, "=AddRate(1)"), (4, 1, "=Tax"));
        Index(Book, "Data", (2, 5, "=[@Qty]*[@Price]"), (3, 5, "=[@Qty]*[@Price]"), (4, 5, "=[@Qty]*[@Price]"));
        Index(Book, "Summary", (1, 1, "=SUM(Jan:Mar!R1C1)"), (2, 1, "=SUM(Sales[Qty])"));
        Index("Other.xlsx", "Sheet1", (1, 1, "='[Model.xlsx]Inputs'!R2C3"));

        string[] Deps(string sheet, int r, int c) =>
            e.Dependents(Book, sheet, CellRect.Cell(r, c)).Items.Select(i => i.Label).ToArray();

        Assert.Equal(new[] { "Calc!$A$1", "Calc!$A$2", "Calc!$A$3", "[Other.xlsx]Sheet1!$A$1" }, Deps("Inputs", 2, 3));
        Assert.Equal(new[] { "Calc!$A$4" }, Deps("Inputs", 3, 3));
        Assert.Equal(new[] { "$E$3", "Summary!$A$2" }, Deps("Data", 3, 3));
        Assert.Equal(new[] { "Summary!$A$1" }, Deps("Feb", 1, 1));
        Assert.Empty(Deps("Feb", 2, 1));
    }

    [Fact]
    public void ChangingNamesMarksBookDirty()
    {
        var e = Engine();
        var w = e.BeginSheet(Book, "Calc");
        w.Add(1, 1, "=Rate");
        w.Commit();
        Assert.True(e.IsIndexed(Book, "Calc"));
        e.UpdateBook(new WorkbookModel(Book, new[] { "Inputs", "Calc", "Data", "Jan", "Feb", "Mar", "Summary" },
            new[] { new NameDef("Rate", null, "=Inputs!R9C9") }));
        Assert.False(e.IsIndexed(Book, "Calc"));
    }
}
