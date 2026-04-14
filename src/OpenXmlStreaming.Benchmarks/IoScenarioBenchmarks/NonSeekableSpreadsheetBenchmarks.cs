[MemoryDiagnoser]
public class NonSeekableSpreadsheetBenchmarks
{
    [Benchmark(Baseline = true)]
    public void Standard()
    {
        using var sink = new NonSeekableDiscardStream();
        using var buffer = new MemoryStream();

        IoScenarioContent.WriteSpreadsheetStandard(buffer);

        buffer.Position = 0;
        buffer.CopyTo(sink);
    }

    [Benchmark]
    public void ForwardOnly()
    {
        using var sink = new NonSeekableDiscardStream();
        IoScenarioContent.WriteSpreadsheetForwardOnly(sink);
    }
}
