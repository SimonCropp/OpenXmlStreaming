// ─── Non-seekable sink ────────────────────────────────────────────────────
// The headline scenario: writing a large document to a non-seekable sink
// (e.g. HTTP response body). The SDK path has to buffer the whole package
// into a MemoryStream first; the streaming path writes directly.
[MemoryDiagnoser]
public class NonSeekableWordBenchmarks
{
    [Benchmark(Baseline = true)]
    public void Standard()
    {
        using var sink = new NonSeekableDiscardStream();
        using var buffer = new MemoryStream();

        IoScenarioContent.WriteWordStandard(buffer);

        buffer.Position = 0;
        buffer.CopyTo(sink);
    }

    [Benchmark]
    public void ForwardOnly()
    {
        using var sink = new NonSeekableDiscardStream();
        IoScenarioContent.WriteWordForwardOnly(sink);
    }
}
