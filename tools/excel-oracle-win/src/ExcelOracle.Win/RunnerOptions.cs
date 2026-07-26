using ExcelOracle.Contracts;

namespace ExcelOracle.Win;

public sealed record RunnerOptions(
    string RunId,
    string CasePath,
    string InputPath,
    string OutputRoot,
    string ObservationPath,
    string Channel,
    string Locale,
    string Timezone)
{
    public static RunnerOptions Parse(string[] args)
    {
        if (args.Length == 0 || args[0] != "observe")
        {
            throw new ContractException("expected the observe command");
        }
        if ((args.Length - 1) % 2 != 0)
        {
            throw new ContractException("runner flags require values");
        }

        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        for (var index = 1; index < args.Length; index += 2)
        {
            var flag = args[index];
            var value = args[index + 1];
            if (value.Length == 0 || !string.Equals(value, value.Trim(), StringComparison.Ordinal))
            {
                throw new ContractException($"runner flag {flag} requires a trimmed value");
            }
            if (!values.TryAdd(flag, value))
            {
                throw new ContractException($"runner flag {flag} was repeated");
            }
        }

        var allowed = new HashSet<string>(
            ["--run-id", "--case", "--input", "--output-root", "--observation", "--channel", "--locale", "--timezone"],
            StringComparer.Ordinal);
        var unknown = values.Keys.FirstOrDefault(flag => !allowed.Contains(flag));
        if (unknown is not null)
        {
            throw new ContractException($"unknown runner flag {unknown}");
        }
        var runId = Required(values, "--run-id");
        if (!runId.All(character =>
                char.IsAsciiLetterOrDigit(character) || character is '.' or '_' or '-'))
        {
            throw new ContractException("runner run id must be an ASCII identifier");
        }
        return new RunnerOptions(
            runId,
            Required(values, "--case"),
            Required(values, "--input"),
            Required(values, "--output-root"),
            Required(values, "--observation"),
            Required(values, "--channel"),
            Required(values, "--locale"),
            Required(values, "--timezone"));
    }

    private static string Required(IReadOnlyDictionary<string, string> values, string flag) =>
        values.TryGetValue(flag, out var value)
            ? value
            : throw new ContractException($"runner flag {flag} is required");
}
