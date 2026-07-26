using System.Text.Json.Nodes;
using ExcelOracle.Contracts;

const string ValidCase = """
{
  "schemaVersion": 1,
  "id": "application.name",
  "version": 1,
  "tier": "mustMatch",
  "input": {
    "path": "ranges/application-name.xlsx",
    "sha256": "0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef",
    "provenance": {
      "source": "Microsoft Excel desktop",
      "producer": "ootd oracle corpus"
    }
  },
  "profileId": "excel-win-en-us",
  "operations": [
    { "operation": "calculate" }
  ],
  "probes": [
    { "id": "application-name", "target": "application", "member": "Name" }
  ]
}
""";

Run("validates the shared case contract", () =>
{
    using var document = ContractValidator.ValidateCase(ValidCase);
    Equal("application.name", document.RootElement.GetProperty("id").GetString());
});

Run("rejects unknown operation fields", () =>
{
    var invalid = ValidCase.Replace(
        "{ \"operation\": \"calculate\" }",
        "{ \"operation\": \"calculate\", \"unexpected\": true }",
        StringComparison.Ordinal);
    Throws<ContractException>(() => ContractValidator.ValidateCase(invalid));
});

Run("normalizes non-zero-based rectangular arrays in row-major order", () =>
{
    var array = Array.CreateInstance(typeof(object), [2, 2], [1, 1]);
    array.SetValue(1.0, 1, 1);
    array.SetValue("two", 1, 2);
    array.SetValue(true, 2, 1);
    array.SetValue(null, 2, 2);

    var normalized = ValueNormalizer.Normalize(array).AsObject();
    Equal("array", normalized["type"]?.GetValue<string>());
    var value = normalized["value"]?.AsObject() ?? throw new Exception("array value missing");
    Equal(2, value["rows"]?.GetValue<int>());
    Equal(2, value["cols"]?.GetValue<int>());
    var values = value["values"]?.AsArray() ?? throw new Exception("array values missing");
    Equal("number", values[0]?["type"]?.GetValue<string>());
    Equal("text", values[1]?["type"]?.GetValue<string>());
    Equal("bool", values[2]?["type"]?.GetValue<string>());
    Equal("null", values[3]?["type"]?.GetValue<string>());
});

Run("does not mistake a generic integer for a cell error", () =>
{
    var number = ValueNormalizer.Normalize(2042).AsObject();
    Equal("number", number["type"]?.GetValue<string>());
    Equal(2042.0, number["value"]?.GetValue<double>());

    var error = ValueNormalizer.NormalizeRangeCellError(2042).AsObject();
    Equal("cellError", error["type"]?.GetValue<string>());
    Equal("#N/A", error["value"]?["code"]?.GetValue<string>());
    Equal(2042, error["value"]?["cvErr"]?.GetValue<int>());
});

Run("preserves missing empty null and void as distinct values", () =>
{
    Equal("missing", ValueNormalizer.Normalize(OracleValue.Missing)["type"]?.GetValue<string>());
    Equal("empty", ValueNormalizer.Normalize(OracleValue.Empty)["type"]?.GetValue<string>());
    Equal("null", ValueNormalizer.Normalize(DBNull.Value)["type"]?.GetValue<string>());
    Equal("void", ValueNormalizer.Normalize(OracleValue.Void)["type"]?.GetValue<string>());
});

return;

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
