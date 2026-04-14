[TestFixture]
public class BufferedWriteStreamTests
{
    [Test]
    public async Task WriteAsync_SpillsViaTargetWriteAsync_NotSyncWrite()
    {
        using var stream = new MemoryStream();
        var tracker = new SyncAsyncTrackingStream(stream);

        // 16-byte buffer with a 48-byte async write forces 3 spills during WriteAsync.
        // Every spill must go through target.WriteAsync — that's what distinguishes
        // the override from the base-Stream sync fallback.
        await using var buffered = new BufferedWriteStream(tracker, bufferSize: 16, leaveOpen: true);

        await buffered.WriteAsync(new byte[48]);

        Assert.Multiple(() =>
        {
            Assert.That(tracker.AsyncWriteCalls, Is.GreaterThanOrEqualTo(2),
                "Spills inside WriteAsync should use target.WriteAsync, not sync Write");
            Assert.That(tracker.SyncWriteCalls, Is.Zero,
                "WriteAsync must not fall back to target.Write");
        });
    }

    [Test]
    public async Task WriteAsync_AccumulatesUntilDisposeAsync_WhenFitsInBuffer()
    {
        using var stream = new MemoryStream();
        var tracker = new SyncAsyncTrackingStream(stream);

        await using (var buffered = new BufferedWriteStream(tracker, bufferSize: 1024, leaveOpen: true))
        {
            await buffered.WriteAsync(new byte[100]);
            Assert.That(tracker.TotalBytesWritten, Is.Zero,
                "Small async writes that fit should accumulate, not reach the target");
        }

        Assert.That(tracker.TotalBytesWritten, Is.EqualTo(100),
            "DisposeAsync should flush the accumulated bytes");
        Assert.That(tracker.AsyncWriteCalls, Is.GreaterThanOrEqualTo(1));
    }

    [Test]
    public void Write_SyncSpillsViaTargetWrite()
    {
        using var stream = new MemoryStream();
        var tracker = new SyncAsyncTrackingStream(stream);

        using var buffered = new BufferedWriteStream(tracker, bufferSize: 16, leaveOpen: true);
        buffered.Write(new byte[48], 0, 48);
        buffered.Flush();

        Assert.Multiple(() =>
        {
            Assert.That(tracker.SyncWriteCalls, Is.GreaterThan(0),
                "Sync Write should spill via target.Write");
            Assert.That(tracker.AsyncWriteCalls, Is.Zero);
        });
    }

    [Test]
    public void Dispose_LeaveOpen_DoesNotDisposeTarget()
    {
        using var stream = new MemoryStream();
        var tracker = new SyncAsyncTrackingStream(stream);

        using (var buffered = new BufferedWriteStream(tracker, bufferSize: 16, leaveOpen: true))
        {
            buffered.Write([1, 2, 3], 0, 3);
        }

        Assert.DoesNotThrow(() => stream.WriteByte(0));
    }

    [Test]
    public async Task DisposeAsync_LeaveOpen_DoesNotDisposeTarget()
    {
        using var stream = new MemoryStream();
        var tracker = new SyncAsyncTrackingStream(stream);

        await using (var buffered = new BufferedWriteStream(tracker, bufferSize: 16, leaveOpen: true))
        {
            await buffered.WriteAsync(new byte[] { 1, 2, 3 });
        }

        Assert.DoesNotThrow(() => stream.WriteByte(0));
    }
}
