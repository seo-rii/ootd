using System.IO.Compression;
using System.Xml;
using ExcelOracle.Contracts;

namespace ExcelOracle.Win;

public static class PackagePreflight
{
    private const long MaxArchiveBytes = 128L * 1024 * 1024;
    private const int MaxEntries = 10_000;
    private const long MaxEntryBytes = 64L * 1024 * 1024;
    private const long MaxTotalBytes = 256L * 1024 * 1024;
    private static readonly HashSet<string> ActiveContentTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        "application/vnd.ms-office.vbaProject",
        "application/vnd.ms-office.vbaProjectSignature",
        "application/vnd.ms-office.vbaProjectSignatureAgile",
        "application/vnd.ms-office.vbaProjectSignatureV3",
        "application/vnd.ms-office.vbaData+xml",
        "application/vnd.ms-excel.macrosheet",
        "application/vnd.ms-excel.macrosheet+xml",
        "application/vnd.ms-excel.intlmacrosheet",
        "application/vnd.ms-excel.intlmacrosheet+xml",
        "application/vnd.ms-excel.dialogsheet",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml",
        "application/vnd.ms-office.activeX",
        "application/vnd.ms-office.activeX+xml",
        "application/vnd.ms-excel.controlproperties+xml",
        "application/vnd.openxmlformats-officedocument.oleObject",
    };
    private static readonly HashSet<string> ActiveContentRelationshipTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        "http://schemas.microsoft.com/office/2006/relationships/vbaProject",
        "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature",
        "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignatureAgile",
        "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignatureV3",
        "http://schemas.microsoft.com/office/2006/relationships/vbaData",
        "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet",
        "http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet",
        "http://purl.oclc.org/ooxml/officeDocument/relationships/dialogsheet",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/control",
        "http://purl.oclc.org/ooxml/officeDocument/relationships/control",
        "http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary",
        "http://schemas.microsoft.com/office/2006/relationships/ctrlProp",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package",
        "http://purl.oclc.org/ooxml/officeDocument/relationships/oleObject",
        "http://purl.oclc.org/ooxml/officeDocument/relationships/package",
        "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility",
        "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility",
        "http://schemas.microsoft.com/office/2006/relationships/ui/userCustomization",
    };

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
            if (IsActiveContentPart(name))
            {
                throw new ContractException($"Excel Oracle input contained active-content part {name}");
            }
            if (name.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase)
                && XmlPartContainsMarker(entry, "ContentType", IsActiveContentType))
            {
                throw new ContractException("Excel Oracle input declared an active-content content type");
            }
            if (name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)
                && XmlPartContainsMarker(entry, "Type", ActiveContentRelationshipTypes.Contains))
            {
                throw new ContractException($"Excel Oracle input declared an active-content relationship in {name}");
            }
        }
    }

    private static bool IsActiveContentPart(string name) =>
        name.Equals("xl/vbaProject.bin", StringComparison.OrdinalIgnoreCase)
        || name.StartsWith("xl/vbaProjectSignature", StringComparison.OrdinalIgnoreCase)
        || name.Equals("xl/vbaData.xml", StringComparison.OrdinalIgnoreCase)
        || name.StartsWith("xl/macrosheets/", StringComparison.OrdinalIgnoreCase)
        || name.StartsWith("xl/dialogsheets/", StringComparison.OrdinalIgnoreCase)
        || name.StartsWith("xl/activeX/", StringComparison.OrdinalIgnoreCase)
        || name.StartsWith("xl/ctrlProps/", StringComparison.OrdinalIgnoreCase)
        || name.StartsWith("xl/embeddings/", StringComparison.OrdinalIgnoreCase)
        || name.StartsWith("customUI/", StringComparison.OrdinalIgnoreCase);

    private static bool IsActiveContentType(string value)
    {
        var separator = value.IndexOf(';');
        var mediaType = separator >= 0 ? value[..separator] : value;
        return ActiveContentTypes.Contains(mediaType.Trim());
    }

    private static bool XmlPartContainsMarker(
        ZipArchiveEntry entry,
        string attributeName,
        Func<string, bool> isMarker)
    {
        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = MaxEntryBytes,
            MaxCharactersFromEntities = 0,
        };
        try
        {
            using var stream = entry.Open();
            using var reader = XmlReader.Create(stream, settings);
            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element || !reader.HasAttributes)
                {
                    continue;
                }
                for (var hasAttribute = reader.MoveToFirstAttribute(); hasAttribute; hasAttribute = reader.MoveToNextAttribute())
                {
                    if (reader.LocalName.Equals(attributeName, StringComparison.Ordinal)
                        && isMarker(reader.Value))
                    {
                        return true;
                    }
                }
                reader.MoveToElement();
            }
            return false;
        }
        catch (XmlException error)
        {
            throw new ContractException($"Excel Oracle input contained malformed package XML in {entry.FullName}: {error.Message}");
        }
    }
}
