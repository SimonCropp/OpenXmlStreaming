// ─── File on disk ─────────────────────────────────────────────────────────
// Both paths are legal against a seekable FileStream. This measures whether
// the SDK's in-package buffering adds visible overhead vs writing direct.

[MemoryDiagnoser]
public class FileWordBenchmarks
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
        IoScenarioContent.WriteWordStandard(stream);
    }

    [Benchmark]
    public void ForwardOnly()
    {
        using var stream = File.Create(tempFile);
        IoScenarioContent.WriteWordForwardOnly(stream);
    }
}
