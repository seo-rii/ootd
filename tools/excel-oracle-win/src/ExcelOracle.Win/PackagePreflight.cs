using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json.Nodes;
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
    private static readonly HashSet<string> ExternalDataContentTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml",
        "application/vnd.ms-excel.externalLink",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml",
        "application/vnd.ms-excel.connections",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml",
        "application/vnd.ms-excel.queryTable",
        "application/vnd.ms-excel.model",
        "application/vnd.ms-excel.dataModel",
    };

    public static void Validate(string inputPath) => ValidateCore(inputPath);

    public static void ValidateAndWriteAudit(
        string inputPath,
        string auditPath,
        string inputRole)
    {
        try
        {
            var metrics = ValidateCore(inputPath);
            WriteAudit(inputPath, auditPath, inputRole, "accepted", metrics, reason: null);
        }
        catch (ContractException error)
        {
            WriteAudit(
                inputPath,
                auditPath,
                inputRole,
                "rejected",
                metrics: null,
                reason: error.Message);
            throw;
        }
        catch (InvalidDataException error)
        {
            var contractError = new ContractException(
                $"Excel Oracle input was not a valid ZIP package: {error.Message}");
            WriteAudit(
                inputPath,
                auditPath,
                inputRole,
                "rejected",
                metrics: null,
                reason: contractError.Message);
            throw contractError;
        }
    }

    private static PackagePreflightMetrics ValidateCore(string inputPath)
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
            if (IsExternalDataPart(name))
            {
                throw new ContractException($"Excel Oracle input contained external-data part {name}");
            }
            if (name.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase)
                && XmlPartContainsMarker(entry, "ContentType", IsActiveContentType))
            {
                throw new ContractException("Excel Oracle input declared an active-content content type");
            }
            if (name.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase)
                && XmlPartContainsMarker(entry, "ContentType", IsExternalDataContentType))
            {
                throw new ContractException("Excel Oracle input declared an external-data content type");
            }
            if (name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)
                && XmlPartContainsMarker(entry, "Type", ActiveContentRelationshipTypes.Contains))
            {
                throw new ContractException($"Excel Oracle input declared an active-content relationship in {name}");
            }
            if (name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)
                && XmlPartContainsMarker(entry, "Type", IsExternalDataRelationshipType))
            {
                throw new ContractException($"Excel Oracle input declared an external-data relationship in {name}");
            }
        }
        return new PackagePreflightMetrics(info.Length, archive.Entries.Count, total);
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

    private static bool IsExternalDataPart(string name) =>
        name.StartsWith("xl/externalLinks/", StringComparison.OrdinalIgnoreCase)
        || name.Equals("xl/connections.xml", StringComparison.OrdinalIgnoreCase)
        || name.StartsWith("xl/queryTables/", StringComparison.OrdinalIgnoreCase)
        || name.StartsWith("xl/model/", StringComparison.OrdinalIgnoreCase)
        || name.StartsWith("xl/customData/", StringComparison.OrdinalIgnoreCase);

    private static bool IsActiveContentType(string value)
    {
        var separator = value.IndexOf(';');
        var mediaType = separator >= 0 ? value[..separator] : value;
        return ActiveContentTypes.Contains(mediaType.Trim());
    }

    private static bool IsExternalDataContentType(string value)
    {
        var separator = value.IndexOf(';');
        var mediaType = separator >= 0 ? value[..separator] : value;
        return ExternalDataContentTypes.Contains(mediaType.Trim());
    }

    private static bool IsExternalDataRelationshipType(string value)
    {
        var relationshipType = value.Trim();
        return relationshipType.EndsWith("/externalLink", StringComparison.OrdinalIgnoreCase)
            || relationshipType.EndsWith("/externalLinkPath", StringComparison.OrdinalIgnoreCase)
            || relationshipType.Contains("/xlExternalLinkPath/", StringComparison.OrdinalIgnoreCase)
            || relationshipType.EndsWith("/externalLinkLongPath", StringComparison.OrdinalIgnoreCase)
            || relationshipType.Contains("/xlExternalLinkLongPath/", StringComparison.OrdinalIgnoreCase)
            || relationshipType.EndsWith("/oleObjectLinkLongPath", StringComparison.OrdinalIgnoreCase)
            || relationshipType.EndsWith("/connections", StringComparison.OrdinalIgnoreCase)
            || relationshipType.EndsWith("/queryTable", StringComparison.OrdinalIgnoreCase)
            || relationshipType.EndsWith("/model", StringComparison.OrdinalIgnoreCase)
            || relationshipType.EndsWith("/modelConnection", StringComparison.OrdinalIgnoreCase);
    }

    private static void WriteAudit(
        string inputPath,
        string auditPath,
        string inputRole,
        string decision,
        PackagePreflightMetrics? metrics,
        string? reason)
    {
        var fullPath = Path.GetFullPath(inputPath);
        var info = new FileInfo(fullPath);
        var audit = new JsonObject
        {
            ["schemaVersion"] = 1,
            ["inputRole"] = inputRole,
            ["inputFileName"] = Path.GetFileName(fullPath),
            ["decision"] = decision,
            ["excelActivationEligible"] = decision == "accepted",
            ["policies"] = new JsonObject
            {
                ["activeContent"] = "refuse",
                ["externalData"] = "refuse",
                ["network"] = "required-host-isolation",
            },
        };
        if (metrics is not null)
        {
            audit["archiveBytes"] = metrics.ArchiveBytes;
            audit["entryCount"] = metrics.EntryCount;
            audit["uncompressedBytes"] = metrics.UncompressedBytes;
        }
        else if (info.Exists)
        {
            audit["archiveBytes"] = info.Length;
        }
        if (CanHashInput(info))
        {
            using var stream = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.Read);
            audit["inputSha256"] = Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
        }
        if (reason is not null)
        {
            audit["reason"] = reason;
        }
        AtomicArtifacts.WriteJson(auditPath, audit);
    }

    private static bool CanHashInput(FileInfo info) =>
        info.Exists
        && info.Length <= MaxArchiveBytes
        && (info.Attributes & FileAttributes.ReparsePoint) == 0;

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

    private sealed record PackagePreflightMetrics(
        long ArchiveBytes,
        int EntryCount,
        long UncompressedBytes);
}
