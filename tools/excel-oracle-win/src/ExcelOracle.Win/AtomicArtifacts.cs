using System.Text.Json;
using System.Text.Json.Nodes;

namespace ExcelOracle.Win;

public static class AtomicArtifacts
{
    public static byte[] WriteJson(string path, JsonNode value)
    {
        var fullPath = Path.GetFullPath(path);
        Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
        var temporaryPath = fullPath + $".{Guid.NewGuid():N}.tmp";
        var bytes = JsonSerializer.SerializeToUtf8Bytes(
            value,
            new JsonSerializerOptions { WriteIndented = true });
        try
        {
            File.WriteAllBytes(temporaryPath, bytes);
            File.Move(temporaryPath, fullPath, overwrite: true);
            return bytes;
        }
        finally
        {
            if (File.Exists(temporaryPath))
            {
                File.Delete(temporaryPath);
            }
        }
    }
}
