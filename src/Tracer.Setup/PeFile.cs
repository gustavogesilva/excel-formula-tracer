using System;

namespace Tracer.Setup;

/// <summary>Reads the target machine of a Windows executable, to tell 32-bit from 64-bit Excel.</summary>
internal static class PeFile
{
    public const ushort I386 = 0x014C;
    public const ushort Amd64 = 0x8664;
    public const ushort Arm64 = 0xAA64;

    /// <summary>The COFF machine field, or null when <paramref name="header"/> (the first bytes of the file) is not a PE image.</summary>
    public static ushort? Machine(byte[] header)
    {
        if (header.Length < 0x40 || header[0] != 'M' || header[1] != 'Z') return null;
        int pe = BitConverter.ToInt32(header, 0x3C);
        if (pe < 0x40 || pe > header.Length - 6) return null;
        if (header[pe] != 'P' || header[pe + 1] != 'E' || header[pe + 2] != 0 || header[pe + 3] != 0) return null;
        return BitConverter.ToUInt16(header, pe + 4);
    }

    /// <summary>
    /// Whether the machine needs the 64-bit XLL. ARM64 Office runs x64 add-ins (Arm64EC), so it counts as 64-bit;
    /// that path is untested.
    /// </summary>
    public static bool Is64Bit(ushort machine) => machine == Amd64 || machine == Arm64;
}
