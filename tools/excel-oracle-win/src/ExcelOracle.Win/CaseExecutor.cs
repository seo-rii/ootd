using System.Text.Json;
using System.Text.Json.Nodes;
using ExcelOracle.Contracts;

namespace ExcelOracle.Win;

internal sealed class CaseExecutor
{
    private readonly IExcelAutomation automation;
    private readonly Dictionary<string, object> objects = new(StringComparer.Ordinal);
    private readonly Dictionary<object, string> identities = new(ReferenceEqualityComparer.Instance);
    private int nextIdentity = 1;

    public CaseExecutor(IExcelAutomation automation, object workbook)
    {
        this.automation = automation;
        BindInitial("application", automation.Application);
        BindInitial("workbook", workbook);
    }

    public JsonObject Execute(JsonElement caseRoot, string outputRoot)
    {
        var operations = new JsonArray();
        JsonObject? saveReopen = null;
        var operationIndex = 0;
        foreach (var operation in caseRoot.GetProperty("operations").EnumerateArray())
        {
            var result = ExecuteOperation(operation, outputRoot, ref saveReopen);
            operations.Add(new JsonObject
            {
                ["operationIndex"] = operationIndex,
                ["result"] = result,
            });
            operationIndex++;
        }

        var probes = new JsonArray();
        foreach (var probe in caseRoot.GetProperty("probes").EnumerateArray())
        {
            var target = Resolve(probe.GetProperty("target").GetString()!);
            var member = probe.GetProperty("member").GetString()!;
            var arguments = DecodeArguments(probe);
            JsonObject result;
            try
            {
                result = ValueResult(Normalize(automation.Get(target, member, arguments), null));
            }
            catch (AutomationCallException error)
            {
                result = ErrorResult(error);
            }
            probes.Add(new JsonObject
            {
                ["id"] = probe.GetProperty("id").GetString(),
                ["result"] = result,
            });
        }

        var output = new JsonObject
        {
            ["operations"] = operations,
            ["probes"] = probes,
        };
        if (saveReopen is not null)
        {
            output["saveReopen"] = saveReopen;
        }
        return output;
    }

    private JsonObject ExecuteOperation(
        JsonElement operation,
        string outputRoot,
        ref JsonObject? saveReopen)
    {
        var kind = operation.GetProperty("operation").GetString();
        try
        {
            return kind switch
            {
                "get" => ExecuteGet(operation),
                "set" => ExecuteSet(operation),
                "invoke" => ExecuteInvoke(operation),
                "calculate" => ExecuteCalculate(),
                "save" => ExecuteSave(operation, outputRoot, ref saveReopen),
                _ => throw new ContractException($"unsupported operation {kind}"),
            };
        }
        catch (AutomationCallException error)
        {
            return ErrorResult(error);
        }
    }

    private JsonObject ExecuteGet(JsonElement operation)
    {
        var target = Resolve(operation.GetProperty("target").GetString()!);
        var value = automation.Get(
            target,
            operation.GetProperty("member").GetString()!,
            DecodeArguments(operation));
        return ValueResult(Normalize(value, OptionalBind(operation)));
    }

    private JsonObject ExecuteSet(JsonElement operation)
    {
        var target = Resolve(operation.GetProperty("target").GetString()!);
        automation.Set(
            target,
            operation.GetProperty("member").GetString()!,
            DecodeValue(operation.GetProperty("value")),
            DecodeArguments(operation));
        return ValueResult(ValueNormalizer.Normalize(OracleValue.Void));
    }

    private JsonObject ExecuteInvoke(JsonElement operation)
    {
        var target = Resolve(operation.GetProperty("target").GetString()!);
        var value = automation.Invoke(
            target,
            operation.GetProperty("member").GetString()!,
            DecodeArguments(operation));
        return ValueResult(Normalize(value, OptionalBind(operation)));
    }

    private JsonObject ExecuteCalculate()
    {
        automation.Calculate();
        return ValueResult(ValueNormalizer.Normalize(OracleValue.Void));
    }

    private JsonObject ExecuteSave(
        JsonElement operation,
        string outputRoot,
        ref JsonObject? saveReopen)
    {
        if (saveReopen is not null)
        {
            throw new ContractException("a case may contain only one save operation");
        }
        var workbook = Resolve(operation.GetProperty("workbook").GetString()!);
        var relative = operation.GetProperty("output").GetString()!;
        var root = Path.GetFullPath(outputRoot);
        var output = Path.GetFullPath(Path.Combine(root, relative.Replace('/', Path.DirectorySeparatorChar)));
        if (!output.StartsWith(root + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
        {
            throw new ContractException("save output escaped the run root");
        }
        Directory.CreateDirectory(Path.GetDirectoryName(output)!);
        automation.SaveAs(workbook, output);
        var reopen = automation.ReopenNormal(output);
        saveReopen = new JsonObject
        {
            ["attempted"] = true,
            ["normalLoadSucceeded"] = reopen.NormalLoadSucceeded,
            ["repairDetected"] = reopen.RepairDetected,
            ["evidence"] = reopen.Evidence,
        };
        return ValueResult(ValueNormalizer.Normalize(OracleValue.Void));
    }

    private JsonObject Normalize(object? value, string? bind)
    {
        if (!automation.IsAutomationObject(value))
        {
            if (bind is not null)
            {
                throw new ContractException($"binding {bind} received a non-object result");
            }
            return ValueNormalizer.Normalize(value);
        }

        var objectValue = value!;
        if (bind is not null)
        {
            if (!objects.TryAdd(bind, objectValue))
            {
                throw new ContractException($"object binding {bind} already exists");
            }
        }
        if (!identities.TryGetValue(objectValue, out var identity))
        {
            identity = bind ?? $"object_{nextIdentity++:0000}";
            identities.Add(objectValue, identity);
        }
        return ValueNormalizer.NormalizeObject(automation.GetAutomationTypeName(objectValue), identity);
    }

    private object?[] DecodeArguments(JsonElement owner)
    {
        if (!owner.TryGetProperty("args", out var arguments))
        {
            return [];
        }
        return arguments.EnumerateArray().Select(DecodeValue).ToArray();
    }

    private object? DecodeValue(JsonElement value)
    {
        var kind = value.GetProperty("type").GetString();
        return kind switch
        {
            "missing" => Type.Missing,
            "empty" => null,
            "null" => DBNull.Value,
            "bool" => value.GetProperty("value").GetBoolean(),
            "number" => value.GetProperty("value").GetDouble(),
            "text" => value.GetProperty("value").GetString(),
            "cellError" => value.GetProperty("value").GetProperty("cvErr").GetInt32(),
            "object" => Resolve(value.GetProperty("value").GetProperty("identity").GetString()!),
            "array" => DecodeArray(value.GetProperty("value")),
            "void" => throw new ContractException("void cannot be used as an operation argument"),
            _ => throw new ContractException($"unsupported operation value {kind}"),
        };
    }

    private object?[,] DecodeArray(JsonElement array)
    {
        var rows = array.GetProperty("rows").GetInt32();
        var cols = array.GetProperty("cols").GetInt32();
        var result = new object?[rows, cols];
        var index = 0;
        foreach (var value in array.GetProperty("values").EnumerateArray())
        {
            result[index / cols, index % cols] = DecodeValue(value);
            index++;
        }
        return result;
    }

    private object Resolve(string identity) => objects.TryGetValue(identity, out var value)
        ? value
        : throw new ContractException($"object binding {identity} was not defined");

    private void BindInitial(string identity, object value)
    {
        objects.Add(identity, value);
        identities.Add(value, identity);
    }

    private static string? OptionalBind(JsonElement operation) =>
        operation.TryGetProperty("bind", out var bind) ? bind.GetString() : null;

    private static JsonObject ValueResult(JsonObject value) => new()
    {
        ["status"] = "value",
        ["result"] = value,
    };

    private static JsonObject ErrorResult(AutomationCallException error)
    {
        var diagnostic = new JsonObject
        {
            ["origin"] = "excelCom",
            ["message"] = error.Message,
        };
        if (error.NativeHresult is not null)
        {
            diagnostic["hresult"] = error.NativeHresult.Value;
        }
        return new JsonObject
        {
            ["status"] = "error",
            ["result"] = new JsonObject
            {
                ["kind"] = error.Kind,
                ["code"] = error.Code,
                ["diagnostic"] = diagnostic,
            },
        };
    }
}
