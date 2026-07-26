using System.Security.Cryptography;
using System.Text.Json.Nodes;
using ExcelOracle.Contracts;

namespace ExcelOracle.Win;

public sealed class OracleRunner(Func<IExcelAutomation> automationFactory)
{
    public JsonObject Run(
        string caseJson,
        string inputPath,
        string outputRoot,
        EngineFingerprint engine)
    {
        using var caseDocument = ContractValidator.ValidateCase(caseJson);
        var root = caseDocument.RootElement;
        var expectedSha256 = root.GetProperty("input").GetProperty("sha256").GetString();
        var actualSha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(inputPath)))
            .ToLowerInvariant();
        if (!string.Equals(expectedSha256, actualSha256, StringComparison.Ordinal))
        {
            throw new ContractException("input sha256 did not match the case");
        }

        var readOnly = !root.GetProperty("operations")
            .EnumerateArray()
            .Any(operation => operation.GetProperty("operation").GetString() == "save");
        using var automation = automationFactory();
        try
        {
            automation.Configure();
            var workbook = automation.OpenWorkbook(Path.GetFullPath(inputPath), readOnly);
            var result = new CaseExecutor(automation, workbook).Execute(root, outputRoot);
            result.Insert(0, "schemaVersion", 1);
            result.Insert(1, "caseId", root.GetProperty("id").GetString());
            result.Insert(2, "engine", EngineJson(engine));
            return result;
        }
        finally
        {
            automation.Close();
        }
    }

    private static JsonObject EngineJson(EngineFingerprint engine) => new()
    {
        ["kind"] = "excel",
        ["version"] = engine.Version,
        ["build"] = engine.Build,
        ["channel"] = engine.Channel,
        ["os"] = engine.Os,
        ["architecture"] = engine.Architecture,
        ["locale"] = engine.Locale,
        ["timezone"] = engine.Timezone,
    };
}
