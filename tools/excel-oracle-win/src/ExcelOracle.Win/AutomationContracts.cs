namespace ExcelOracle.Win;

public sealed record EngineFingerprint(
    string Version,
    string Build,
    string Channel,
    string Os,
    string Architecture,
    string Locale,
    string Timezone)
{
    public static EngineFingerprint TestExcel { get; } = new(
        "16.0",
        "17928.20156",
        "Current",
        "Windows 11",
        "x64",
        "en-US",
        "UTC");
}

public sealed record SaveReopenResult(
    bool NormalLoadSucceeded,
    bool? RepairDetected,
    string? Evidence);

public interface IExcelAutomation : IDisposable
{
    object Application { get; }
    void Configure();
    object OpenWorkbook(string inputPath, bool readOnly);
    object? Get(object target, string member, object?[] arguments);
    void Set(object target, string member, object? value, object?[] arguments);
    object? Invoke(object target, string member, object?[] arguments);
    void Calculate();
    void SaveAs(object workbook, string outputPath);
    SaveReopenResult ReopenNormal(string outputPath);
    bool IsAutomationObject(object? value);
    string GetAutomationTypeName(object value);
    void Close();
}

public sealed class AutomationCallException : Exception
{
    public AutomationCallException(
        string kind,
        string code,
        string message,
        int? hresult = null)
        : base(message)
    {
        Kind = kind;
        Code = code;
        NativeHresult = hresult;
    }

    public string Kind { get; }
    public string Code { get; }
    public int? NativeHresult { get; }

    public static AutomationCallException NotFound(string code, string message) =>
        new("notFound", code, message);
}
