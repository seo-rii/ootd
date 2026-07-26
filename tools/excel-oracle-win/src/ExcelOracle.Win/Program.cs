using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Nodes;
using ExcelOracle.Contracts;
using ExcelOracle.Win;

if (!OperatingSystem.IsWindows())
{
    Console.Error.WriteLine("{\"status\":\"unsupported_host\",\"message\":\"desktop Excel automation requires Windows\"}");
    return 2;
}

try
{
    var options = RunnerOptions.Parse(args);
    var casePath = Path.GetFullPath(options.CasePath);
    var inputPath = Path.GetFullPath(options.InputPath);
    var outputRoot = Path.GetFullPath(options.OutputRoot);
    var observationPath = Path.GetFullPath(options.ObservationPath);
    if (Directory.Exists(outputRoot) && Directory.EnumerateFileSystemEntries(outputRoot).Any())
    {
        throw new ContractException("output root must be fresh and empty");
    }
    Directory.CreateDirectory(outputRoot);
    if (!observationPath.StartsWith(outputRoot + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
    {
        throw new ContractException("observation path must be contained by the output root");
    }

    PackagePreflight.Validate(inputPath);
    var sandboxInput = Path.Combine(outputRoot, "work", "input" + Path.GetExtension(inputPath));
    Directory.CreateDirectory(Path.GetDirectoryName(sandboxInput)!);
    File.Copy(inputPath, sandboxInput, overwrite: false);
    PackagePreflight.Validate(sandboxInput);

    var caseBytes = File.ReadAllBytes(casePath);
    string caseJson;
    try
    {
        caseJson = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true)
            .GetString(caseBytes);
    }
    catch (DecoderFallbackException error)
    {
        throw new ContractException($"case must be valid UTF-8: {error.Message}");
    }
    using var caseDocument = ContractValidator.ValidateCase(caseJson);
    var caseRoot = caseDocument.RootElement;
    var inputBytes = File.ReadAllBytes(sandboxInput);
    var inputSha256 = Convert.ToHexString(SHA256.HashData(inputBytes)).ToLowerInvariant();
    var expectedInputSha256 = caseRoot.GetProperty("input").GetProperty("sha256").GetString();
    if (!string.Equals(inputSha256, expectedInputSha256, StringComparison.Ordinal))
    {
        throw new ContractException("input sha256 did not match the case");
    }

    var ownedProcesses = new OwnedProcessRegistry(
        Path.Combine(outputRoot, "manifest", "owned_processes.json"));
    using var session = new ExcelComSession(
        options.Channel,
        options.Locale,
        options.Timezone,
        ownedProcesses.Record);
    var observation = new OracleRunner(() => session).Run(
        caseJson,
        sandboxInput,
        outputRoot,
        session.Fingerprint);
    var observationBytes = AtomicArtifacts.WriteJson(observationPath, observation);
    var observationRelativePath = Path.GetRelativePath(outputRoot, observationPath).Replace('\\', '/');
    var runManifest = new JsonObject
    {
        ["schemaVersion"] = 1,
        ["runId"] = options.RunId,
        ["profileId"] = caseRoot.GetProperty("profileId").GetString(),
        ["engine"] = session.Fingerprint.ToObservationJson(),
        ["cases"] = new JsonArray
        {
            new JsonObject
            {
                ["caseId"] = caseRoot.GetProperty("id").GetString(),
                ["caseVersion"] = caseRoot.GetProperty("version").GetInt32(),
                ["tier"] = caseRoot.GetProperty("tier").GetString(),
                ["caseSha256"] = Convert.ToHexString(SHA256.HashData(caseBytes)).ToLowerInvariant(),
                ["inputSha256"] = inputSha256,
                ["status"] = "completed",
                ["observationPath"] = observationRelativePath,
                ["observationSha256"] = Convert.ToHexString(SHA256.HashData(observationBytes))
                    .ToLowerInvariant(),
            },
        },
    };
    AtomicArtifacts.WriteJson(Path.Combine(outputRoot, "manifest", "run_manifest.json"), runManifest);
    Console.WriteLine(observationPath);
    return 0;
}
catch (ContractException error)
{
    Console.Error.WriteLine($"contract_error: {error.Message}");
    return 3;
}
catch (AutomationCallException error)
{
    Console.Error.WriteLine($"automation_error: {error.Code}: {error.Message}");
    return 4;
}
catch (Exception error)
{
    Console.Error.WriteLine($"infrastructure_error: {error}");
    return 5;
}
