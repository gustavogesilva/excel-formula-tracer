using System;
using System.Reflection;
using System.Runtime.InteropServices;

namespace Tracer.AddIn.Excel;

/// <summary>
/// Reads a block of cells in row chunks of at most <see cref="ChunkCells"/> cells, so no single COM call asks Excel
/// for millions of cells (one <c>object[,]</c> per call; per-call overhead is what makes bulk reads fast).
/// When Excel answers a chunk with out-of-memory, which large models hit in 32-bit Excel, the chunk is halved and
/// the same rows read again; the smaller size is kept from then on. Pure, unit-tested.
/// </summary>
internal sealed class ChunkedReads
{
    private const int EOutOfMemory = unchecked((int)0x8007000E);

    public ChunkedReads(int chunkCells = 65_536) => ChunkCells = Math.Max(1, chunkCells);

    public int ChunkCells { get; private set; }

    /// <summary>
    /// Calls <paramref name="read"/>(rowOffset, rowCount) for consecutive row chunks of a rows × cols block until every
    /// row is read or it returns false. It must make its COM reads before any side effect, so a chunk read again after
    /// out-of-memory is not counted twice. A single row that still runs out of memory rethrows.
    /// </summary>
    public void Walk(int rows, int cols, Func<int, int, bool> read)
    {
        cols = Math.Max(1, cols);
        for (int offset = 0; offset < rows;)
        {
            int count = Math.Min(Math.Max(1, ChunkCells / cols), rows - offset);
            try
            {
                if (!read(offset, count)) return;
                offset += count;
            }
            catch (Exception ex) when (count > 1 && IsOutOfMemory(ex))
            {
                ChunkCells = count / 2 * cols;
            }
        }
    }

    /// <summary>
    /// Excel refused for lack of memory: E_OUTOFMEMORY, which .NET raises as <see cref="OutOfMemoryException"/>
    /// (so <c>catch (COMException)</c> misses it), possibly wrapped by a late-bound call.
    /// </summary>
    public static bool IsOutOfMemory(Exception? ex) => ex switch
    {
        TargetInvocationException t => IsOutOfMemory(t.InnerException),
        OutOfMemoryException => true,
        COMException c => c.HResult == EOutOfMemory,
        _ => false,
    };
}
