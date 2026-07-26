using System.IO.Compression;
using ExcelOracle.Contracts;

namespace ExcelOracle.Win;

public static class PackagePreflight
{
    private const long MaxArchiveBytes = 128L * 1024 * 1024;
    private const int MaxEntries = 10_000;
    private const long MaxEntryBytes = 64L * 1024 * 1024;
    private const long MaxTotalBytes = 256L * 1024 * 1024;

    public static void Validate(string inputPath)
    {
        var fullPath = Path.GetFullPath(inputPath);
        var extension = Path.GetExtension(fullPath);
        if (!extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase)
            && !extension.Equals(".xltx", StringComparison.OrdinalIgnoreCase))
        {
            throw new ContractException("Excel Oracle inputs must be .xlsx or .xltx");
        }
        var info = new FileInfo(fullPath);
        if (!info.Exists || info.Length > MaxArchiveBytes)
        {
            throw new ContractException("Excel Oracle input was missing or exceeded 128 MiB");
        }
        if ((info.Attributes & FileAttributes.ReparsePoint) != 0)
        {
            throw new ContractException("Excel Oracle input must not be a reparse point");
        }

        using var stream = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        if (archive.Entries.Count > MaxEntries)
        {
            throw new ContractException("Excel Oracle input exceeded the ZIP entry limit");
        }
        long total = 0;
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in archive.Entries)
        {
            var name = entry.FullName.Replace('\\', '/');
            if (!names.Add(name)
                || name.StartsWith("/", StringComparison.Ordinal)
                || name.Split('/').Any(segment => segment.Length == 0 || segment is "." or ".."))
            {
                throw new ContractException("Excel Oracle input contained an ambiguous ZIP entry name");
            }
            if (entry.Length > MaxEntryBytes)
            {
                throw new ContractException("Excel Oracle input exceeded the per-entry size limit");
            }
            total = checked(total + entry.Length);
            if (total > MaxTotalBytes)
            {
                throw new ContractException("Excel Oracle input exceeded the total size limit");
            }
            if (IsExecutablePart(name))
            {
                throw new ContractException($"Excel Oracle input contained executable part {name}");
            }
        }
    }

    private static bool IsExecutablePart(string name) =>
        name.Equals("xl/vbaProject.bin", StringComparison.OrdinalIgnoreCase)
        || name.StartsWith("xl/macrosheets/", StringComparison.OrdinalIgnoreCase)
        || name.StartsWith("xl/dialogsheets/", StringComparison.OrdinalIgnoreCase);
}
