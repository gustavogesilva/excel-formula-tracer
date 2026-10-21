using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using Tracer.AddIn.Excel;
using Xunit;

namespace Tracer.Core.Tests;

public class ChunkedReadsTests
{
    /// <summary>Reads that run out of memory above <paramref name="limitCells"/> cells, recording what was read.</summary>
    private static Func<int, int, bool> Reader(int cols, int limitCells, List<(int Offset, int Count)> read, int stopAfterRows = int.MaxValue)
    {
        int total = 0;
        return (offset, count) =>
        {
            if (count * cols > limitCells) throw new OutOfMemoryException("Not enough memory resources (E_OUTOFMEMORY)");
            read.Add((offset, count));
            total += count;
            return total < stopAfterRows;
        };
    }

    [Fact]
    public void ReadsEveryRowOnceInChunks()
    {
        var reads = new ChunkedReads(1_000);
        var read = new List<(int, int)>();
        reads.Walk(250, 10, Reader(10, int.MaxValue, read));
        Assert.Equal(new[] { (0, 100), (100, 100), (200, 50) }, read);
    }

    [Fact]
    public void OutOfMemoryHalvesTheChunkAndKeepsItHalved()
    {
        var reads = new ChunkedReads(1_000);
        var read = new List<(int, int)>();
        reads.Walk(100, 10, Reader(10, 300, read));
        Assert.Equal(new[] { (0, 25), (25, 25), (50, 25), (75, 25) }, read);
        Assert.Equal(250, reads.ChunkCells);

        read.Clear();
        reads.Walk(30, 10, Reader(10, int.MaxValue, read));
        Assert.Equal(new[] { (0, 25), (25, 5) }, read);
    }

    [Fact]
    public void StopsWhenTheReaderHasEnough()
    {
        var read = new List<(int, int)>();
        new ChunkedReads(100).Walk(1_000, 1, Reader(1, int.MaxValue, read, stopAfterRows: 150));
        Assert.Equal(new[] { (0, 100), (100, 100) }, read);
    }

    [Fact]
    public void ASingleRowThatStillFailsRethrows()
    {
        var read = new List<(int, int)>();
        Assert.Throws<OutOfMemoryException>(() => new ChunkedReads(1_000).Walk(10, 500, Reader(500, 100, read)));
        Assert.Empty(read);
    }

    [Fact]
    public void OtherErrorsAreNotRetried()
    {
        int calls = 0;
        Assert.Throws<COMException>(() => new ChunkedReads(1_000).Walk(100, 10, (_, _) =>
        {
            calls++;
            throw new COMException("busy", unchecked((int)0x800AC472));
        }));
        Assert.Equal(1, calls);
    }

    [Fact]
    public void RecognisesOutOfMemoryInItsForms()
    {
        Assert.True(ChunkedReads.IsOutOfMemory(new OutOfMemoryException()));
        Assert.True(ChunkedReads.IsOutOfMemory(new COMException("x", unchecked((int)0x8007000E))));
        Assert.True(ChunkedReads.IsOutOfMemory(new TargetInvocationException(new OutOfMemoryException())));
        Assert.False(ChunkedReads.IsOutOfMemory(new COMException("x", unchecked((int)0x800AC472))));
        Assert.False(ChunkedReads.IsOutOfMemory(new InvalidOperationException()));
        Assert.False(ChunkedReads.IsOutOfMemory(null));
    }
}
