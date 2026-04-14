/// <summary>
/// Test helper that tracks whether writes come in through the sync or async
/// code paths. Used to verify that <see cref="OpenXmlPackageWriter.DisposeAsync"/>
/// actually routes the final buffer flush through <see cref="Stream.WriteAsync(ReadOnlyMemory{byte}, System.Threading.CancellationToken)"/>.
/// </summary>
class SyncAsyncTrackingStream(Stream inner) : Stream
{
    public int SyncWriteCalls { get; private set; }
    public int AsyncWriteCalls { get; private set; }
    public long TotalBytesWritten { get; private set; }

    public override bool CanRead => false;
    public override bool CanSeek => false;
    public override bool CanWrite => true;

    public override long Length => throw new NotSupportedException();

    public override long Position
    {
        get => throw new NotSupportedException();
        set => throw new NotSupportedException();
    }

    public override void Flush() => inner.Flush();
    public override Task FlushAsync(Cancel cancel) => inner.FlushAsync(cancel);

    public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();

    public override void Write(byte[] buffer, int offset, int count)
    {
        SyncWriteCalls++;
        TotalBytesWritten += count;
        inner.Write(buffer, offset, count);
    }

    public override Task WriteAsync(byte[] buffer, int offset, int count, Cancel cancel)
    {
        AsyncWriteCalls++;
        TotalBytesWritten += count;
        return inner.WriteAsync(buffer, offset, count, cancel);
    }

    public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, Cancel cancel = default)
    {
        AsyncWriteCalls++;
        TotalBytesWritten += buffer.Length;
        return inner.WriteAsync(buffer, cancel);
    }
}
