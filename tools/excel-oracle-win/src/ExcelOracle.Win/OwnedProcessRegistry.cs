using System.Text.Json.Nodes;
using ExcelOracle.Contracts;

namespace ExcelOracle.Win;

public sealed class OwnedProcessRegistry(string manifestPath)
{
    private readonly List<OwnedExcelProcess> processes = [];
    private readonly object sync = new();

    public void Record(OwnedExcelProcess process)
    {
        if (process.ProcessId <= 0)
        {
            throw new ContractException("owned Excel process id must be positive");
        }

        lock (sync)
        {
            if (!processes.Contains(process))
            {
                processes.Add(process);
            }

            var entries = new JsonArray();
            foreach (var owned in processes)
            {
                entries.Add(new JsonObject
                {
                    ["processId"] = owned.ProcessId,
                    ["startTimeUtc"] = owned.StartTimeUtc.ToUniversalTime().ToString("O"),
                });
            }
            AtomicArtifacts.WriteJson(
                manifestPath,
                new JsonObject
                {
                    ["schemaVersion"] = 1,
                    ["processes"] = entries,
                });
        }
    }
}
