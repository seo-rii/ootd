using System.Text.Json;

if (!OperatingSystem.IsWindows())
{
    Console.Error.WriteLine(JsonSerializer.Serialize(new
    {
        status = "unsupported_host",
        message = "desktop Excel automation requires Windows",
    }));
    return 2;
}

Console.Error.WriteLine("Excel COM activation is not wired yet");
return 2;
