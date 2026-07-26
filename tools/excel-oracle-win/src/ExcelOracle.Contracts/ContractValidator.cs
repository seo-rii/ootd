using System.Text.Json;

namespace ExcelOracle.Contracts;

public static class ContractValidator
{
    private const int SchemaVersion = 1;

    public static JsonDocument ValidateCase(string json)
    {
        JsonDocument document;
        try
        {
            document = JsonDocument.Parse(json, new JsonDocumentOptions
            {
                AllowTrailingCommas = false,
                CommentHandling = JsonCommentHandling.Disallow,
                MaxDepth = 256,
            });
        }
        catch (JsonException error)
        {
            throw new ContractException($"invalid case JSON: {error.Message}");
        }

        try
        {
            ValidateCaseRoot(document.RootElement);
            return document;
        }
        catch
        {
            document.Dispose();
            throw;
        }
    }

    private static void ValidateCaseRoot(JsonElement root)
    {
        RequireObject(root, "case");
        RequireExactProperties(
            root,
            "case",
            ["schemaVersion", "id", "version", "tier", "input", "profileId", "operations", "probes"]);
        if (RequireInt32(root, "schemaVersion", "case") != SchemaVersion)
        {
            throw new ContractException("unsupported case schemaVersion");
        }

        RequireIdentifier(root, "id", "case");
        if (RequireInt32(root, "version", "case") <= 0)
        {
            throw new ContractException("case version must be positive");
        }

        var tier = RequireString(root, "tier", "case");
        if (tier is not ("mustMatch" or "informational"))
        {
            throw new ContractException("case tier was invalid");
        }

        RequireIdentifier(root, "profileId", "case");
        ValidateInput(RequireProperty(root, "input", "case"));
        ValidateOperations(RequireProperty(root, "operations", "case"));
        ValidateProbes(RequireProperty(root, "probes", "case"));
    }

    private static void ValidateInput(JsonElement input)
    {
        RequireObject(input, "case input");
        RequireExactProperties(input, "case input", ["path", "sha256", "provenance"]);
        RequireSafeRelativePath(RequireString(input, "path", "case input"), "case input path");
        RequireSha256(RequireString(input, "sha256", "case input"), "case input sha256");

        var provenance = RequireProperty(input, "provenance", "case input");
        RequireObject(provenance, "case input provenance");
        RequireExactProperties(provenance, "case input provenance", ["source", "producer"]);
        RequireTrimmedString(provenance, "source", "case input provenance");
        RequireTrimmedString(provenance, "producer", "case input provenance");
    }

    private static void ValidateOperations(JsonElement operations)
    {
        RequireArray(operations, "case operations");
        if (operations.GetArrayLength() == 0)
        {
            throw new ContractException("case operations must not be empty");
        }

        var bindings = new HashSet<string>(StringComparer.Ordinal);
        foreach (var operation in operations.EnumerateArray())
        {
            RequireObject(operation, "case operation");
            var kind = RequireString(operation, "operation", "case operation");
            switch (kind)
            {
                case "get":
                case "invoke":
                    RequireOnlyProperties(
                        operation,
                        "case operation",
                        ["operation", "target", "member", "args", "bind"],
                        ["operation", "target", "member"]);
                    RequireTrimmedString(operation, "target", "case operation");
                    RequireTrimmedString(operation, "member", "case operation");
                    ValidateOptionalValues(operation, "args", "case operation arguments");
                    if (operation.TryGetProperty("bind", out var bindElement))
                    {
                        var bind = RequireTrimmedString(bindElement, "case operation bind");
                        if (!bindings.Add(bind))
                        {
                            throw new ContractException("case operation bindings must be unique");
                        }
                    }
                    break;
                case "set":
                    RequireOnlyProperties(
                        operation,
                        "case operation",
                        ["operation", "target", "member", "value", "args"],
                        ["operation", "target", "member", "value"]);
                    RequireTrimmedString(operation, "target", "case operation");
                    RequireTrimmedString(operation, "member", "case operation");
                    ValidateObservedValue(RequireProperty(operation, "value", "case operation"));
                    ValidateOptionalValues(operation, "args", "case operation arguments");
                    break;
                case "calculate":
                    RequireExactProperties(operation, "calculate operation", ["operation"]);
                    break;
                case "save":
                    RequireExactProperties(operation, "save operation", ["operation", "workbook", "output"]);
                    RequireTrimmedString(operation, "workbook", "save operation");
                    RequireSafeRelativePath(
                        RequireString(operation, "output", "save operation"),
                        "save output");
                    break;
                default:
                    throw new ContractException($"unsupported case operation {kind}");
            }
        }
    }

    private static void ValidateProbes(JsonElement probes)
    {
        RequireArray(probes, "case probes");
        if (probes.GetArrayLength() == 0)
        {
            throw new ContractException("case probes must not be empty");
        }

        var ids = new HashSet<string>(StringComparer.Ordinal);
        foreach (var probe in probes.EnumerateArray())
        {
            RequireObject(probe, "case probe");
            RequireOnlyProperties(
                probe,
                "case probe",
                ["id", "target", "member", "args"],
                ["id", "target", "member"]);
            var id = RequireTrimmedString(probe, "id", "case probe");
            if (!ids.Add(id))
            {
                throw new ContractException("case probe ids must be unique");
            }
            RequireTrimmedString(probe, "target", "case probe");
            RequireTrimmedString(probe, "member", "case probe");
            ValidateOptionalValues(probe, "args", "case probe arguments");
        }
    }

    private static void ValidateOptionalValues(JsonElement owner, string property, string context)
    {
        if (!owner.TryGetProperty(property, out var values))
        {
            return;
        }
        RequireArray(values, context);
        foreach (var value in values.EnumerateArray())
        {
            ValidateObservedValue(value);
        }
    }

    private static void ValidateObservedValue(JsonElement value)
    {
        RequireObject(value, "observed value");
        RequireOnlyProperties(value, "observed value", ["type", "value"], ["type"]);
        var kind = RequireString(value, "type", "observed value");
        var hasValue = value.TryGetProperty("value", out var content);
        switch (kind)
        {
            case "void":
            case "missing":
            case "empty":
            case "null":
                if (hasValue)
                {
                    throw new ContractException($"observed {kind} must not include value");
                }
                break;
            case "bool":
                RequireValue(content, hasValue, JsonValueKind.True, JsonValueKind.False, kind);
                break;
            case "number":
                RequireValue(content, hasValue, JsonValueKind.Number, kind);
                if (!content.TryGetDouble(out var number) || !double.IsFinite(number))
                {
                    throw new ContractException("observed number must be finite");
                }
                break;
            case "text":
                RequireValue(content, hasValue, JsonValueKind.String, kind);
                break;
            case "cellError":
                RequireValue(content, hasValue, JsonValueKind.Object, kind);
                RequireExactProperties(content, "observed cell error", ["code", "cvErr"]);
                RequireTrimmedString(content, "code", "observed cell error");
                if (RequireInt32(content, "cvErr", "observed cell error") <= 0)
                {
                    throw new ContractException("observed cell error cvErr must be positive");
                }
                break;
            case "object":
                RequireValue(content, hasValue, JsonValueKind.Object, kind);
                RequireExactProperties(content, "observed object", ["typeName", "identity"]);
                RequireTrimmedString(content, "typeName", "observed object");
                RequireTrimmedString(content, "identity", "observed object");
                break;
            case "array":
                RequireValue(content, hasValue, JsonValueKind.Object, kind);
                ValidateObservedArray(content);
                break;
            default:
                throw new ContractException($"unsupported observed value type {kind}");
        }
    }

    private static void ValidateObservedArray(JsonElement array)
    {
        RequireExactProperties(array, "observed array", ["rows", "cols", "values"]);
        var rows = RequireInt32(array, "rows", "observed array");
        var cols = RequireInt32(array, "cols", "observed array");
        var values = RequireProperty(array, "values", "observed array");
        RequireArray(values, "observed array values");
        if (rows <= 0 || cols <= 0 || (long)rows * cols != values.GetArrayLength())
        {
            throw new ContractException("observed array dimensions must match non-empty values");
        }
        foreach (var value in values.EnumerateArray())
        {
            ValidateObservedValue(value);
        }
    }

    private static void RequireValue(
        JsonElement content,
        bool hasValue,
        JsonValueKind expected,
        string kind)
    {
        if (!hasValue || content.ValueKind != expected)
        {
            throw new ContractException($"observed {kind} has an invalid value");
        }
    }

    private static void RequireValue(
        JsonElement content,
        bool hasValue,
        JsonValueKind first,
        JsonValueKind second,
        string kind)
    {
        if (!hasValue || (content.ValueKind != first && content.ValueKind != second))
        {
            throw new ContractException($"observed {kind} has an invalid value");
        }
    }

    private static void RequireOnlyProperties(
        JsonElement value,
        string context,
        IReadOnlyCollection<string> allowed,
        IReadOnlyCollection<string> required)
    {
        var allowedSet = allowed.ToHashSet(StringComparer.Ordinal);
        foreach (var property in value.EnumerateObject())
        {
            if (!allowedSet.Contains(property.Name))
            {
                throw new ContractException($"{context} contained unknown field {property.Name}");
            }
        }
        foreach (var property in required)
        {
            if (!value.TryGetProperty(property, out _))
            {
                throw new ContractException($"{context} was missing field {property}");
            }
        }
    }

    private static void RequireExactProperties(
        JsonElement value,
        string context,
        IReadOnlyCollection<string> expected) =>
        RequireOnlyProperties(value, context, expected, expected);

    private static JsonElement RequireProperty(JsonElement owner, string property, string context)
    {
        if (!owner.TryGetProperty(property, out var value))
        {
            throw new ContractException($"{context} was missing field {property}");
        }
        return value;
    }

    private static string RequireIdentifier(JsonElement owner, string property, string context)
    {
        var value = RequireTrimmedString(owner, property, context);
        if (!value.All(character =>
                char.IsAsciiLetterOrDigit(character) || character is '.' or '_' or '-'))
        {
            throw new ContractException($"{context} {property} must be an ASCII identifier");
        }
        return value;
    }

    private static string RequireTrimmedString(JsonElement owner, string property, string context) =>
        RequireTrimmedString(RequireProperty(owner, property, context), $"{context} {property}");

    private static string RequireTrimmedString(JsonElement value, string context)
    {
        if (value.ValueKind != JsonValueKind.String)
        {
            throw new ContractException($"{context} must be a string");
        }
        var text = value.GetString() ?? string.Empty;
        if (text.Length == 0 || !string.Equals(text, text.Trim(), StringComparison.Ordinal))
        {
            throw new ContractException($"{context} must be non-empty and trimmed");
        }
        return text;
    }

    private static string RequireString(JsonElement owner, string property, string context)
    {
        var value = RequireProperty(owner, property, context);
        if (value.ValueKind != JsonValueKind.String)
        {
            throw new ContractException($"{context} {property} must be a string");
        }
        return value.GetString() ?? string.Empty;
    }

    private static int RequireInt32(JsonElement owner, string property, string context)
    {
        var value = RequireProperty(owner, property, context);
        if (value.ValueKind != JsonValueKind.Number || !value.TryGetInt32(out var number))
        {
            throw new ContractException($"{context} {property} must be an integer");
        }
        return number;
    }

    private static void RequireObject(JsonElement value, string context)
    {
        if (value.ValueKind != JsonValueKind.Object)
        {
            throw new ContractException($"{context} must be an object");
        }
    }

    private static void RequireArray(JsonElement value, string context)
    {
        if (value.ValueKind != JsonValueKind.Array)
        {
            throw new ContractException($"{context} must be an array");
        }
    }

    private static void RequireSha256(string value, string context)
    {
        if (value.Length != 64 || value.Any(character =>
                !char.IsAsciiDigit(character) && character is not (>= 'a' and <= 'f')))
        {
            throw new ContractException($"{context} must be lowercase hexadecimal");
        }
    }

    private static void RequireSafeRelativePath(string value, string context)
    {
        if (value.Length == 0
            || value.StartsWith("/", StringComparison.Ordinal)
            || value.Contains('\\')
            || value.Split('/').Any(segment => segment.Length == 0 || segment is "." or ".."))
        {
            throw new ContractException($"{context} must be a safe relative path");
        }
    }
}
