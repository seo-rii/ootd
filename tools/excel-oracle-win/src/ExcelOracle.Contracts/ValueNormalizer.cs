using System.Globalization;
using System.Text.Json.Nodes;

namespace ExcelOracle.Contracts;

public static class OracleValue
{
    public static object Missing { get; } = new Sentinel("missing");
    public static object Empty { get; } = new Sentinel("empty");
    public static object Void { get; } = new Sentinel("void");

    internal sealed record Sentinel(string Kind);
}

public static class ValueNormalizer
{
    public static JsonObject Normalize(object? value) => value switch
    {
        OracleValue.Sentinel sentinel => Unit(sentinel.Kind),
        null or DBNull => Unit("null"),
        bool boolean => Scalar("bool", JsonValue.Create(boolean)),
        string text => Scalar("text", JsonValue.Create(text)),
        byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal =>
            NormalizeNumber(value),
        Array array => NormalizeArray(array),
        _ => throw new ContractException($"unsupported COM value type {value.GetType().FullName}"),
    };

    public static JsonObject NormalizeRangeCellError(int cvErr)
    {
        var code = cvErr switch
        {
            2000 => "#NULL!",
            2007 => "#DIV/0!",
            2015 => "#VALUE!",
            2023 => "#REF!",
            2029 => "#NAME?",
            2036 => "#NUM!",
            2042 => "#N/A",
            2043 => "#GETTING_DATA",
            2045 => "#SPILL!",
            2046 => "#CONNECT!",
            2047 => "#BLOCKED!",
            2048 => "#UNKNOWN!",
            2049 => "#FIELD!",
            2050 => "#CALC!",
            2051 => "#BUSY!",
            2052 => "#PYTHON!",
            2053 => "#TIMEOUT!",
            _ => throw new ContractException($"unsupported Excel cell error {cvErr}"),
        };
        return new JsonObject
        {
            ["type"] = "cellError",
            ["value"] = new JsonObject
            {
                ["code"] = code,
                ["cvErr"] = cvErr,
            },
        };
    }

    public static JsonObject NormalizeObject(string typeName, string identity)
    {
        if (string.IsNullOrWhiteSpace(typeName) || string.IsNullOrWhiteSpace(identity))
        {
            throw new ContractException("object typeName and identity are required");
        }
        return new JsonObject
        {
            ["type"] = "object",
            ["value"] = new JsonObject
            {
                ["typeName"] = typeName,
                ["identity"] = identity,
            },
        };
    }

    private static JsonObject NormalizeNumber(object value)
    {
        double number;
        try
        {
            number = Convert.ToDouble(value, CultureInfo.InvariantCulture);
        }
        catch (Exception error) when (error is FormatException or InvalidCastException or OverflowException)
        {
            throw new ContractException($"invalid numeric COM value: {error.Message}");
        }
        if (!double.IsFinite(number))
        {
            throw new ContractException("numeric COM values must be finite");
        }
        return Scalar("number", JsonValue.Create(number));
    }

    private static JsonObject NormalizeArray(Array array)
    {
        if (array.Rank is < 1 or > 2)
        {
            throw new ContractException("only one- and two-dimensional COM arrays are supported");
        }
        var rows = array.GetLength(0);
        var cols = array.Rank == 1 ? 1 : array.GetLength(1);
        if (rows == 0 || cols == 0)
        {
            throw new ContractException("COM arrays must not be empty");
        }

        var values = new JsonArray();
        var rowLower = array.GetLowerBound(0);
        var colLower = array.Rank == 1 ? 0 : array.GetLowerBound(1);
        for (var row = 0; row < rows; row++)
        {
            for (var col = 0; col < cols; col++)
            {
                var item = array.Rank == 1
                    ? array.GetValue(rowLower + row)
                    : array.GetValue(rowLower + row, colLower + col);
                values.Add(Normalize(item));
            }
        }
        return new JsonObject
        {
            ["type"] = "array",
            ["value"] = new JsonObject
            {
                ["rows"] = rows,
                ["cols"] = cols,
                ["values"] = values,
            },
        };
    }

    private static JsonObject Unit(string kind) => new()
    {
        ["type"] = kind,
    };

    private static JsonObject Scalar(string kind, JsonNode? value) => new()
    {
        ["type"] = kind,
        ["value"] = value,
    };
}
