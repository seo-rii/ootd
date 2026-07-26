using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using ExcelOracle.Contracts;

namespace ExcelOracle.Win;

public sealed record OwnedExcelProcess(int ProcessId, DateTime StartTimeUtc);

[SupportedOSPlatform("windows")]
public sealed class ExcelComSession : IExcelAutomation
{
    private const int XlCalculationManual = -4135;
    private const int XlCalculationDone = 0;
    private const int XlOpenXmlWorkbook = 51;
    private const int XlNormalLoad = 0;
    private const int MsoAutomationSecurityForceDisable = 3;
    private readonly List<object> trackedObjects = [];
    private readonly HashSet<object> trackedObjectSet = new(ReferenceEqualityComparer.Instance);
    private readonly Mutex processMutex;
    private readonly Action<OwnedExcelProcess>? ownedProcessRecorder;
    private object? workbook;
    private bool closed;
    private bool disposed;

    public ExcelComSession(
        string channel,
        string locale,
        string timezone,
        Action<OwnedExcelProcess>? ownedProcessRecorder = null)
    {
        if (Process.GetProcessesByName("EXCEL").Length != 0)
        {
            throw new ContractException("Excel Oracle requires a host without a pre-existing Excel process");
        }
        processMutex = new Mutex(initiallyOwned: false, "Local\\OOTD.ExcelOracle");
        if (!processMutex.WaitOne(TimeSpan.Zero))
        {
            processMutex.Dispose();
            throw new ContractException("another Excel Oracle process owns the automation mutex");
        }
        this.ownedProcessRecorder = ownedProcessRecorder;

        object? activatedApplication = null;
        try
        {
            var excelType = Type.GetTypeFromProgID("Excel.Application", throwOnError: true)
                ?? throw new ContractException("Excel.Application COM registration was not found");
            activatedApplication = Activator.CreateInstance(excelType)
                ?? throw new ContractException("Excel.Application COM activation returned null");
            Application = Track(activatedApplication);
            OwnedProcess = ResolveOwnedProcess(Application);
            ownedProcessRecorder?.Invoke(OwnedProcess);
            Fingerprint = new EngineFingerprint(
                Convert.ToString(GetPropertyValue(Application, "Version", []), CultureInfo.InvariantCulture) ?? "unknown",
                Convert.ToString(GetPropertyValue(Application, "Build", []), CultureInfo.InvariantCulture) ?? "unknown",
                channel,
                Environment.OSVersion.VersionString,
                Environment.Is64BitProcess ? "x64" : "x86",
                locale,
                timezone);
        }
        catch
        {
            TryInvoke(activatedApplication, "Quit", []);
            ReleaseTrackedObjects();
            processMutex.ReleaseMutex();
            processMutex.Dispose();
            throw;
        }
    }

    public object Application { get; }
    public OwnedExcelProcess OwnedProcess { get; }
    public EngineFingerprint Fingerprint { get; }

    public void Configure()
    {
        SetProperty(Application, "Visible", false);
        SetProperty(Application, "DisplayAlerts", false);
        SetProperty(Application, "ScreenUpdating", false);
        SetProperty(Application, "EnableEvents", false);
        SetProperty(Application, "AskToUpdateLinks", false);
        SetProperty(Application, "DeferAsyncQueries", true);
        SetProperty(Application, "Calculation", XlCalculationManual);
        SetProperty(Application, "AutomationSecurity", MsoAutomationSecurityForceDisable);
    }

    public object OpenWorkbook(string inputPath, bool readOnly)
    {
        if (workbook is not null)
        {
            throw new ContractException("an Excel Oracle session may open only one workbook");
        }
        var workbooks = Track(GetPropertyValue(Application, "Workbooks", [])!);
        workbook = Track(InvokeMember(
            workbooks,
            "Open",
            BindingFlags.InvokeMethod,
            [
                inputPath,
                0,
                readOnly,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                true,
                Type.Missing,
                Type.Missing,
                false,
                false,
                Type.Missing,
                false,
                true,
                XlNormalLoad,
            ])!);
        return workbook;
    }

    public object? Get(object target, string member, object?[] arguments) =>
        TrackIfCom(InvokeMember(target, member, BindingFlags.GetProperty, arguments));

    public void Set(object target, string member, object? value, object?[] arguments)
    {
        var invocationArguments = new object?[arguments.Length + 1];
        arguments.CopyTo(invocationArguments, 0);
        invocationArguments[^1] = value;
        InvokeMember(target, member, BindingFlags.SetProperty, invocationArguments);
    }

    public object? Invoke(object target, string member, object?[] arguments) =>
        TrackIfCom(InvokeMember(target, member, BindingFlags.InvokeMethod, arguments));

    public void Calculate()
    {
        InvokeMember(Application, "CalculateFullRebuild", BindingFlags.InvokeMethod, []);
        var deadline = Stopwatch.StartNew();
        while (Convert.ToInt32(GetPropertyValue(Application, "CalculationState", []), CultureInfo.InvariantCulture)
               != XlCalculationDone)
        {
            if (deadline.Elapsed > TimeSpan.FromMinutes(2))
            {
                throw new AutomationCallException(
                    "external",
                    "calculationTimeout",
                    "Excel calculation did not reach xlDone within two minutes");
            }
            Thread.Sleep(50);
        }
    }

    public void SaveAs(object workbookObject, string outputPath)
    {
        InvokeMember(
            workbookObject,
            "SaveAs",
            BindingFlags.InvokeMethod,
            [
                outputPath,
                XlOpenXmlWorkbook,
                Type.Missing,
                Type.Missing,
                false,
                false,
                1,
                Type.Missing,
                false,
                Type.Missing,
                Type.Missing,
                true,
            ]);
    }

    public SaveReopenResult ReopenNormal(string outputPath)
    {
        object? reopenApplication = null;
        object? reopenWorkbooks = null;
        object? reopenWorkbook = null;
        try
        {
            var excelType = Type.GetTypeFromProgID("Excel.Application", throwOnError: true)
                ?? throw new ContractException("Excel.Application COM registration was not found");
            reopenApplication = Activator.CreateInstance(excelType)
                ?? throw new ContractException("Excel.Application reopen activation returned null");
            ownedProcessRecorder?.Invoke(ResolveOwnedProcess(reopenApplication));
            SetProperty(reopenApplication, "Visible", false);
            SetProperty(reopenApplication, "DisplayAlerts", false);
            SetProperty(reopenApplication, "EnableEvents", false);
            SetProperty(reopenApplication, "AutomationSecurity", MsoAutomationSecurityForceDisable);
            reopenWorkbooks = GetPropertyValue(reopenApplication, "Workbooks", []);
            reopenWorkbook = InvokeMember(
                reopenWorkbooks!,
                "Open",
                BindingFlags.InvokeMethod,
                [
                    outputPath,
                    0,
                    true,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    true,
                    Type.Missing,
                    Type.Missing,
                    false,
                    false,
                    Type.Missing,
                    false,
                    true,
                    XlNormalLoad,
                ]);
            return new SaveReopenResult(
                true,
                false,
                "Workbooks.Open CorruptLoad=xlNormalLoad");
        }
        catch (Exception error) when (error is COMException or TargetInvocationException or AutomationCallException)
        {
            return new SaveReopenResult(false, null, $"normal open failed: {error.GetType().Name}");
        }
        finally
        {
            TryInvoke(reopenWorkbook, "Close", [false]);
            TryInvoke(reopenApplication, "Quit", []);
            ReleaseComObject(reopenWorkbook);
            ReleaseComObject(reopenWorkbooks);
            ReleaseComObject(reopenApplication);
        }
    }

    public bool IsAutomationObject(object? value) => value is not null && Marshal.IsComObject(value);

    public long GetAutomationIdentity(object value)
    {
        var identity = Marshal.GetIUnknownForObject(value);
        try
        {
            return identity.ToInt64();
        }
        finally
        {
            Marshal.Release(identity);
        }
    }

    public string GetAutomationTypeName(object value) => "Object";

    public void Close()
    {
        if (closed)
        {
            return;
        }
        closed = true;
        Exception? cleanupError = null;
        try
        {
            if (workbook is not null)
            {
                InvokeMember(workbook, "Close", BindingFlags.InvokeMethod, [false]);
            }
        }
        catch (Exception error)
        {
            cleanupError = error;
        }
        try
        {
            InvokeMember(Application, "Quit", BindingFlags.InvokeMethod, []);
        }
        catch (Exception error)
        {
            cleanupError ??= error;
        }
        if (cleanupError is not null)
        {
            throw new InvalidOperationException("Excel COM cleanup failed", cleanupError);
        }
    }

    public void Dispose()
    {
        if (disposed)
        {
            return;
        }
        disposed = true;
        try
        {
            Close();
        }
        finally
        {
            ReleaseTrackedObjects();
            processMutex.ReleaseMutex();
            processMutex.Dispose();
        }
    }

    private object Track(object value)
    {
        if (Marshal.IsComObject(value) && trackedObjectSet.Add(value))
        {
            trackedObjects.Add(value);
        }
        return value;
    }

    private object? TrackIfCom(object? value) => value is not null && Marshal.IsComObject(value)
        ? Track(value)
        : value;

    private void ReleaseTrackedObjects()
    {
        for (var index = trackedObjects.Count - 1; index >= 0; index--)
        {
            ReleaseComObject(trackedObjects[index]);
        }
        trackedObjects.Clear();
        trackedObjectSet.Clear();
    }

    private static object? GetPropertyValue(object target, string member, object?[] arguments) =>
        InvokeMember(target, member, BindingFlags.GetProperty, arguments);

    private static void SetProperty(object target, string member, object? value) =>
        InvokeMember(target, member, BindingFlags.SetProperty, [value]);

    private static object? InvokeMember(
        object target,
        string member,
        BindingFlags operation,
        object?[] arguments)
    {
        try
        {
            return target.GetType().InvokeMember(
                member,
                operation | BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase,
                binder: null,
                target,
                arguments,
                CultureInfo.InvariantCulture);
        }
        catch (TargetInvocationException error) when (error.InnerException is not null)
        {
            throw ConvertAutomationError(member, error.InnerException);
        }
        catch (Exception error) when (error is COMException or MissingMemberException or MissingMethodException)
        {
            throw ConvertAutomationError(member, error);
        }
    }

    private static AutomationCallException ConvertAutomationError(string member, Exception error)
    {
        if (error is MissingMemberException or MissingMethodException)
        {
            return AutomationCallException.NotFound(member, error.Message);
        }
        var hresult = error.HResult;
        var kind = hresult switch
        {
            unchecked((int)0x80020003) => "notFound",
            unchecked((int)0x80020005) => "typeMismatch",
            unchecked((int)0x8002000E) => "invalidArgument",
            _ => "applicationDefined",
        };
        return new AutomationCallException(
            kind,
            $"hresult_{unchecked((uint)hresult):X8}",
            error.Message,
            hresult);
    }

    private static void TryInvoke(object? target, string member, object?[] arguments)
    {
        if (target is null)
        {
            return;
        }
        try
        {
            InvokeMember(target, member, BindingFlags.InvokeMethod, arguments);
        }
        catch
        {
            // Best-effort cleanup for the separate normal-open verification session.
        }
    }

    private static void ReleaseComObject(object? value)
    {
        if (value is not null && Marshal.IsComObject(value))
        {
            Marshal.FinalReleaseComObject(value);
        }
    }

    private static OwnedExcelProcess ResolveOwnedProcess(object excelApplication)
    {
        var hwnd = Convert.ToInt64(
            GetPropertyValue(excelApplication, "Hwnd", []),
            CultureInfo.InvariantCulture);
        if (hwnd == 0 || GetWindowThreadProcessId(new nint(hwnd), out var processId) == 0 || processId == 0)
        {
            throw new ContractException("could not resolve the owned Excel process from its window");
        }
        using var process = Process.GetProcessById(checked((int)processId));
        return new OwnedExcelProcess(process.Id, process.StartTime.ToUniversalTime());
    }

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(nint window, out uint processId);
}
