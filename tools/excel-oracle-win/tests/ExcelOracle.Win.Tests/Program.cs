using System.Security.Cryptography;
using System.IO.Compression;
using System.Runtime.CompilerServices;
using System.Text.Json.Nodes;
using ExcelOracle.Contracts;
using ExcelOracle.Win;

Run("configures opens executes and closes in order", () =>
{
    using var fixture = CaseFixture.Create();
    var automation = new FakeAutomation();
    var runner = new OracleRunner(() => automation);

    var observation = runner.Run(
        fixture.CaseJson,
        fixture.InputPath,
        fixture.OutputRoot,
        EngineFingerprint.TestExcel);

    Equal(
        "configure,open,get:Worksheets,invoke:Item,invoke:Range,set:Value2,get:Value2,get:Value2,close,dispose",
        string.Join(',', automation.Calls));
    Equal("application.range", observation["caseId"]?.GetValue<string>());
    Equal("array", observation["probes"]?[0]?["result"]?["result"]?["type"]?.GetValue<string>());
});

Run("records an automation error and continues", () =>
{
    using var fixture = CaseFixture.Create(includeFailure: true);
    var automation = new FakeAutomation();
    var runner = new OracleRunner(() => automation);

    var observation = runner.Run(
        fixture.CaseJson,
        fixture.InputPath,
        fixture.OutputRoot,
        EngineFingerprint.TestExcel);

    Equal("error", observation["operations"]?[0]?["result"]?["status"]?.GetValue<string>());
    Equal("notFound", observation["operations"]?[0]?["result"]?["result"]?["kind"]?.GetValue<string>());
    Equal("array", observation["probes"]?[0]?["result"]?["result"]?["type"]?.GetValue<string>());
    True(automation.Calls.Contains("close"), "close was not called");
});

Run("closes and disposes after an unexpected probe failure", () =>
{
    using var fixture = CaseFixture.Create();
    var automation = new FakeAutomation { FailProbeUnexpectedly = true };
    var runner = new OracleRunner(() => automation);

    Throws<InvalidOperationException>(() => runner.Run(
        fixture.CaseJson,
        fixture.InputPath,
        fixture.OutputRoot,
        EngineFingerprint.TestExcel));

    True(automation.Calls.Contains("close"), "close was not called");
    Equal("dispose", automation.Calls[^1]);
});

Run("parses a strict observe command", () =>
{
    var options = RunnerOptions.Parse([
        "observe",
        "--run-id", "application-name-a",
        "--case", "case.json",
        "--input", "input.xlsx",
        "--output-root", "run",
        "--observation", "run/observations/oracle.json",
        "--channel", "Current",
        "--locale", "en-US",
        "--timezone", "UTC",
    ]);

    Equal("application-name-a", options.RunId);
    Equal("case.json", options.CasePath);
    Equal("input.xlsx", options.InputPath);
    Equal("Current", options.Channel);
    Throws<ContractException>(() => RunnerOptions.Parse(["observe", "--unknown", "value"]));
});

Run("rejects active-content paths before Excel activation", () =>
{
    foreach (var activePath in new[]
    {
        "xl/vbaProject.bin",
        "xl/vbaProjectSignatureAgile.bin",
        "xl/vbaData.xml",
        "xl/macrosheets/sheet1.xml",
        "xl/dialogsheets/sheet1.xml",
        "xl/activeX/activeX1.xml",
        "xl/ctrlProps/ctrlProp1.xml",
        "xl/embeddings/oleObject1.bin",
        "customUI/customUI.xml",
    })
    {
        var path = Path.Combine(Path.GetTempPath(), $"ootd-oracle-active-path-{Guid.NewGuid():N}.xlsx");
        try
        {
            WriteZipPackage(path, new Dictionary<string, string>
            {
                ["[Content_Types].xml"] = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>",
                ["xl/workbook.xml"] = "<workbook/>",
                [activePath] = "fixture",
            });
            Throws<ContractException>(() => PackagePreflight.Validate(path));
        }
        finally
        {
            File.Delete(path);
        }
    }
});

Run("rejects active-content content types and relationships before Excel activation", () =>
{
    foreach (var contentType in new[]
    {
        "application/vnd.ms-office.vbaProject",
        "application/vnd.ms-office.vbaProjectSignatureV3",
        "application/vnd.ms-office.vbaData+xml",
        "application/vnd.ms-excel.macrosheet+xml",
        "application/vnd.ms-excel.intlmacrosheet+xml",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml",
        "application/vnd.ms-office.activeX",
        "application/vnd.ms-office.activeX+xml",
        "application/vnd.ms-excel.controlproperties+xml",
        "application/vnd.openxmlformats-officedocument.oleObject",
    })
    {
        var path = Path.Combine(Path.GetTempPath(), $"ootd-oracle-active-type-{Guid.NewGuid():N}.xlsx");
        try
        {
            WriteZipPackage(path, new Dictionary<string, string>
            {
                ["[Content_Types].xml"] = $"<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Override PartName=\"/custom/payload.bin\" ContentType=\"{contentType}; charset=binary\"/></Types>",
                ["xl/workbook.xml"] = "<workbook/>",
            });
            Throws<ContractException>(() => PackagePreflight.Validate(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    foreach (var relationshipType in new[]
    {
        "http://schemas.microsoft.com/office/2006/relationships/vbaProject",
        "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignatureAgile",
        "http://schemas.microsoft.com/office/2006/relationships/vbaData",
        "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet",
        "http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet",
        "http://purl.oclc.org/ooxml/officeDocument/relationships/dialogsheet",
        "http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package",
        "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility",
    })
    {
        var path = Path.Combine(Path.GetTempPath(), $"ootd-oracle-active-rel-{Guid.NewGuid():N}.xlsx");
        try
        {
            WriteZipPackage(path, new Dictionary<string, string>
            {
                ["[Content_Types].xml"] = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>",
                ["xl/workbook.xml"] = "<workbook/>",
                ["_rels/.rels"] = $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rIdActive\" Type=\"{relationshipType}\" Target=\"custom/payload.bin\"/></Relationships>",
            });
            Throws<ContractException>(() => PackagePreflight.Validate(path));
        }
        finally
        {
            File.Delete(path);
        }
    }
});

Run("rejects external-data paths before Excel activation", () =>
{
    foreach (var externalDataPath in new[]
    {
        "xl/externalLinks/externalLink1.xml",
        "xl/connections.xml",
        "xl/queryTables/queryTable1.xml",
        "xl/model/model.bin",
        "xl/customData/item1.xml",
    })
    {
        var path = Path.Combine(Path.GetTempPath(), $"ootd-oracle-external-path-{Guid.NewGuid():N}.xlsx");
        try
        {
            WriteZipPackage(path, new Dictionary<string, string>
            {
                ["[Content_Types].xml"] = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>",
                ["xl/workbook.xml"] = "<workbook/>",
                [externalDataPath] = "fixture",
            });
            Throws<ContractException>(() => PackagePreflight.Validate(path));
        }
        finally
        {
            File.Delete(path);
        }
    }
});

Run("rejects external-data content types and relationships before Excel activation", () =>
{
    foreach (var contentType in new[]
    {
        "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml",
        "application/vnd.ms-excel.externalLink",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml",
        "application/vnd.ms-excel.connections",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml",
        "application/vnd.ms-excel.queryTable",
        "application/vnd.ms-excel.model",
        "application/vnd.ms-excel.dataModel",
    })
    {
        var path = Path.Combine(Path.GetTempPath(), $"ootd-oracle-external-type-{Guid.NewGuid():N}.xlsx");
        try
        {
            WriteZipPackage(path, new Dictionary<string, string>
            {
                ["[Content_Types].xml"] = $"<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Override PartName=\"/custom/payload.bin\" ContentType=\"{contentType}; charset=binary\"/></Types>",
                ["xl/workbook.xml"] = "<workbook/>",
            });
            Throws<ContractException>(() => PackagePreflight.Validate(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    foreach (var relationshipType in new[]
    {
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink",
        "http://purl.oclc.org/ooxml/officeDocument/relationships/externalLink",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath",
        "http://schemas.microsoft.com/office/2020/07/relationships/xlExternalLinkPath/remote",
        "http://schemas.microsoft.com/office/2011/relationships/externalLinkLongPath",
        "http://schemas.microsoft.com/office/2020/07/relationships/xlExternalLinkLongPath/remote",
        "http://schemas.microsoft.com/office/2011/relationships/oleObjectLinkLongPath",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable",
        "http://schemas.microsoft.com/office/2007/relationships/model",
        "http://schemas.microsoft.com/office/2011/relationships/modelConnection",
    })
    {
        var path = Path.Combine(Path.GetTempPath(), $"ootd-oracle-external-rel-{Guid.NewGuid():N}.xlsx");
        try
        {
            WriteZipPackage(path, new Dictionary<string, string>
            {
                ["[Content_Types].xml"] = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>",
                ["xl/workbook.xml"] = "<workbook/>",
                ["_rels/.rels"] = $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rIdExternal\" Type=\"{relationshipType}\" Target=\"https://external.example/data\" TargetMode=\"External\"/></Relationships>",
            });
            Throws<ContractException>(() => PackagePreflight.Validate(path));
        }
        finally
        {
            File.Delete(path);
        }
    }
});

Run("writes accepted and rejected package-preflight audits before Excel activation", () =>
{
    var root = Path.Combine(Path.GetTempPath(), $"ootd-oracle-preflight-audit-{Guid.NewGuid():N}");
    Directory.CreateDirectory(root);
    var acceptedPath = Path.Combine(root, "accepted.xlsx");
    var rejectedPath = Path.Combine(root, "rejected.xlsx");
    var acceptedAuditPath = Path.Combine(root, "manifest", "accepted.json");
    var rejectedAuditPath = Path.Combine(root, "manifest", "rejected.json");
    try
    {
        WriteZipPackage(acceptedPath, new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>",
            ["xl/workbook.xml"] = "<workbook/>",
        });
        WriteZipPackage(rejectedPath, new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>",
            ["xl/workbook.xml"] = "<workbook/>",
            ["xl/connections.xml"] = "<connections/>",
        });

        PackagePreflight.ValidateAndWriteAudit(
            acceptedPath,
            acceptedAuditPath,
            "source");
        Throws<ContractException>(() => PackagePreflight.ValidateAndWriteAudit(
            rejectedPath,
            rejectedAuditPath,
            "sandbox-copy"));

        var acceptedAudit = JsonNode.Parse(File.ReadAllBytes(acceptedAuditPath))!;
        Equal("accepted", acceptedAudit["decision"]?.GetValue<string>());
        Equal(true, acceptedAudit["excelActivationEligible"]?.GetValue<bool>());
        Equal("refuse", acceptedAudit["policies"]?["externalData"]?.GetValue<string>());
        Equal(2, acceptedAudit["entryCount"]?.GetValue<int>());
        Equal(
            Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(acceptedPath))).ToLowerInvariant(),
            acceptedAudit["inputSha256"]?.GetValue<string>());

        var rejectedAudit = JsonNode.Parse(File.ReadAllBytes(rejectedAuditPath))!;
        Equal("rejected", rejectedAudit["decision"]?.GetValue<string>());
        Equal(false, rejectedAudit["excelActivationEligible"]?.GetValue<bool>());
        Equal("sandbox-copy", rejectedAudit["inputRole"]?.GetValue<string>());
        Equal(
            Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(rejectedPath))).ToLowerInvariant(),
            rejectedAudit["inputSha256"]?.GetValue<string>());
        Equal(
            "Excel Oracle input contained external-data part xl/connections.xml",
            rejectedAudit["reason"]?.GetValue<string>());
    }
    finally
    {
        Directory.Delete(root, recursive: true);
    }
});

Run("accepts inert content types and relationships in package preflight", () =>
{
    var path = Path.Combine(Path.GetTempPath(), $"ootd-oracle-inert-{Guid.NewGuid():N}.xlsx");
    try
    {
        WriteZipPackage(path, new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/></Types>",
            ["xl/workbook.xml"] = "<workbook/>",
            ["_rels/.rels"] = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.com/\" TargetMode=\"External\"/></Relationships>",
        });
        PackagePreflight.Validate(path);
    }
    finally
    {
        File.Delete(path);
    }
});

Run("returns the exact bytes written for artifact hashing", () =>
{
    var root = Path.Combine(Path.GetTempPath(), $"ootd-oracle-artifact-{Guid.NewGuid():N}");
    var path = Path.Combine(root, "manifest", "run_manifest.json");
    try
    {
        var bytes = AtomicArtifacts.WriteJson(
            path,
            new JsonObject
            {
                ["schemaVersion"] = 1,
                ["runId"] = "application-name-a",
            });

        True(bytes.SequenceEqual(File.ReadAllBytes(path)), "returned artifact bytes differed from disk");
    }
    finally
    {
        if (Directory.Exists(root))
        {
            Directory.Delete(root, recursive: true);
        }
    }
});

Run("records every owned Excel process once for watchdog cleanup", () =>
{
    var root = Path.Combine(Path.GetTempPath(), $"ootd-oracle-processes-{Guid.NewGuid():N}");
    var path = Path.Combine(root, "manifest", "owned_processes.json");
    try
    {
        var registry = new OwnedProcessRegistry(path);
        var first = new OwnedExcelProcess(101, new DateTime(2026, 7, 26, 1, 2, 3, DateTimeKind.Utc));
        var second = new OwnedExcelProcess(202, new DateTime(2026, 7, 26, 1, 2, 4, DateTimeKind.Utc));
        registry.Record(first);
        registry.Record(second);
        registry.Record(first);

        var manifest = JsonNode.Parse(File.ReadAllBytes(path))!;
        Equal(1, manifest["schemaVersion"]?.GetValue<int>());
        Equal(2, manifest["processes"]?.AsArray().Count);
        Equal(202, manifest["processes"]?[1]?["processId"]?.GetValue<int>());
    }
    finally
    {
        if (Directory.Exists(root))
        {
            Directory.Delete(root, recursive: true);
        }
    }
});

return;

static void WriteZipPackage(string path, IReadOnlyDictionary<string, string> entries)
{
    using var stream = File.Create(path);
    using var archive = new ZipArchive(stream, ZipArchiveMode.Create);
    foreach (var (name, contents) in entries)
    {
        var entry = archive.CreateEntry(name);
        using var writer = new StreamWriter(entry.Open());
        writer.Write(contents);
    }
}

static void Run(string name, Action test)
{
    try
    {
        test();
        Console.WriteLine($"PASS {name}");
    }
    catch (Exception error)
    {
        Console.Error.WriteLine($"FAIL {name}: {error}");
        Environment.ExitCode = 1;
    }
}

static void Equal<T>(T expected, T actual)
{
    if (!EqualityComparer<T>.Default.Equals(expected, actual))
    {
        throw new Exception($"expected {expected}, got {actual}");
    }
}

static void True(bool value, string message)
{
    if (!value)
    {
        throw new Exception(message);
    }
}

static void Throws<T>(Action action)
    where T : Exception
{
    try
    {
        action();
    }
    catch (T)
    {
        return;
    }
    throw new Exception($"expected {typeof(T).Name}");
}

sealed class FakeAutomation : IExcelAutomation
{
    private readonly Token application = new("Application");
    private readonly Token workbook = new("Workbook");
    private readonly Token worksheets = new("Worksheets");
    private readonly Token worksheet = new("Worksheet");
    private readonly Token range = new("Range");
    private bool valueWasSet;

    public List<string> Calls { get; } = [];
    public bool FailProbeUnexpectedly { get; init; }
    public object Application => application;

    public void Configure()
    {
        Calls.Add("configure");
    }

    public object OpenWorkbook(string inputPath, bool readOnly)
    {
        Calls.Add("open");
        if (!File.Exists(inputPath))
        {
            throw new Exception("input did not exist");
        }
        return workbook;
    }

    public object? Get(object target, string member, object?[] arguments)
    {
        Calls.Add($"get:{member}");
        if (member == "Worksheets" && ReferenceEquals(target, workbook))
        {
            return worksheets;
        }
        if (member == "Value2" && ReferenceEquals(target, range))
        {
            if (FailProbeUnexpectedly && valueWasSet)
            {
                throw new InvalidOperationException("probe infrastructure failed");
            }
            return new object?[,] { { 1.0, "two" } };
        }
        throw AutomationCallException.NotFound(member, "fake member was not found");
    }

    public void Set(object target, string member, object? value, object?[] arguments)
    {
        Calls.Add($"set:{member}");
        valueWasSet = true;
    }

    public object? Invoke(object target, string member, object?[] arguments)
    {
        Calls.Add($"invoke:{member}");
        return member switch
        {
            "DefinitelyMissing" => throw AutomationCallException.NotFound(
                member,
                "fake member was not found"),
            "Item" => worksheet,
            "Range" => range,
            _ => null,
        };
    }

    public void Calculate()
    {
        Calls.Add("calculate");
    }

    public void SaveAs(object workbookObject, string outputPath)
    {
        Calls.Add("save");
        File.WriteAllBytes(outputPath, [1, 2, 3]);
    }

    public SaveReopenResult ReopenNormal(string outputPath)
    {
        Calls.Add("reopen");
        return new SaveReopenResult(true, false, "fake normal open");
    }

    public bool IsAutomationObject(object? value) => value is Token;

    public long GetAutomationIdentity(object value) => RuntimeHelpers.GetHashCode(value);

    public string GetAutomationTypeName(object value) => ((Token)value).TypeName;

    public void Close()
    {
        Calls.Add("close");
    }

    public void Dispose()
    {
        Calls.Add("dispose");
    }

    private sealed record Token(string TypeName);
}

sealed class CaseFixture : IDisposable
{
    private CaseFixture(string root, string inputPath, string outputRoot, string caseJson)
    {
        Root = root;
        InputPath = inputPath;
        OutputRoot = outputRoot;
        CaseJson = caseJson;
    }

    private string Root { get; }
    public string InputPath { get; }
    public string OutputRoot { get; }
    public string CaseJson { get; }

    public static CaseFixture Create(bool includeFailure = false)
    {
        var root = Path.Combine(Path.GetTempPath(), $"ootd-oracle-test-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        var input = Path.Combine(root, "input.xlsx");
        var output = Path.Combine(root, "output");
        Directory.CreateDirectory(output);
        File.WriteAllBytes(input, [1, 2, 3, 4]);
        var sha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))).ToLowerInvariant();
        var failure = includeFailure
            ? "{ \"operation\": \"invoke\", \"target\": \"application\", \"member\": \"DefinitelyMissing\" },"
            : string.Empty;
        var json = $$"""
        {
          "schemaVersion": 1,
          "id": "application.range",
          "version": 1,
          "tier": "mustMatch",
          "input": {
            "path": "ranges/input.xlsx",
            "sha256": "{{sha256}}",
            "provenance": { "source": "Microsoft Excel desktop", "producer": "ootd oracle corpus" }
          },
          "profileId": "excel-win-en-us",
          "operations": [
            {{failure}}
            { "operation": "get", "target": "workbook", "member": "Worksheets", "bind": "worksheets" },
            { "operation": "invoke", "target": "worksheets", "member": "Item", "args": [{ "type": "number", "value": 1 }], "bind": "sheet" },
            { "operation": "invoke", "target": "sheet", "member": "Range", "args": [{ "type": "text", "value": "A1:B1" }], "bind": "range" },
            { "operation": "set", "target": "range", "member": "Value2", "value": { "type": "array", "value": { "rows": 1, "cols": 2, "values": [{ "type": "number", "value": 1 }, { "type": "text", "value": "two" }] } } },
            { "operation": "get", "target": "range", "member": "Value2" }
          ],
          "probes": [
            { "id": "range-values", "target": "range", "member": "Value2" }
          ]
        }
        """;
        return new CaseFixture(root, input, output, json);
    }

    public void Dispose()
    {
        Directory.Delete(Root, recursive: true);
    }
}
