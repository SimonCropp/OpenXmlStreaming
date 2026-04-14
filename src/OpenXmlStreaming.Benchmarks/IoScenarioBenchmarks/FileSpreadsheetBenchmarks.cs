[MemoryDiagnoser]
public class FileSpreadsheetBenchmarks
{
    string tempFile = string.Empty;

    [IterationSetup]
    public void Setup() =>
        tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".tmp");

    [IterationCleanup]
    public void Cleanup()
    {
        if (File.Exists(tempFile))
        {
            File.Delete(tempFile);
        }
    }

    [Benchmark(Baseline = true)]
    public void Standard()
    {
        using var stream = File.Create(tempFile);
        IoScenarioContent.WriteSpreadsheetStandard(stream);
    }

    [Benchmark]
    public void ForwardOnly()
    {
        using var stream = File.Create(tempFile);
        IoScenarioContent.WriteSpreadsheetForwardOnly(stream);
    }
}
