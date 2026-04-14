namespace OpenXmlStreaming;

/// <summary>
/// Write-only, forward-only buffer between a sync writer (e.g. <see cref="ZipArchive"/>
/// in Create mode) and a target stream that benefits from larger/async writes
/// (e.g. a SQL BLOB stream or HTTP response body).
///
/// Sync <see cref="Write(byte[], int, int)"/> calls accumulate in a managed
/// buffer; the buffer is pushed to the target via sync <c>Write</c> when it
/// fills. On <see cref="DisposeAsync"/> (or explicit <see cref="FlushAsync"/>)
/// the pending buffer is pushed via async <c>WriteAsync</c>, so the final —
/// and usually largest single — flush happens without blocking the thread on
/// network I/O.
/// </summary>
sealed class BufferedWriteStream : Stream
{
    readonly Stream target;
    readonly byte[] buffer;
    readonly bool leaveOpen;
    int count;
    bool disposed;

    public BufferedWriteStream(Stream target, int bufferSize, bool leaveOpen)
    {
        this.target = target;
        buffer = new byte[bufferSize];
        this.leaveOpen = leaveOpen;
    }

    public override bool CanRead => false;
    public override bool CanSeek => false;
    public override bool CanWrite => true;

    public override long Length => throw new NotSupportedException();

    public override long Position
    {
        get => throw new NotSupportedException();
        set => throw new NotSupportedException();
    }

    public override int Read(byte[] buffer, int offset, int count) =>
        throw new NotSupportedException();

    public override long Seek(long offset, SeekOrigin origin) =>
        throw new NotSupportedException();

    public override void SetLength(long value) =>
        throw new NotSupportedException();

    public override void Write(byte[] buffer, int offset, int count)
    {
        while (count > 0)
        {
            if (this.count == this.buffer.Length)
            {
                FlushBufferSync();
            }

            var toCopy = Math.Min(count, this.buffer.Length - this.count);
            Array.Copy(buffer, offset, this.buffer, this.count, toCopy);
            this.count += toCopy;
            offset += toCopy;
            count -= toCopy;
        }
    }

    public override void Flush() =>
        FlushBufferSync();

    public override Task FlushAsync(CancellationToken cancellationToken)
    {
        if (count == 0)
        {
            return Task.CompletedTask;
        }

        return FlushBufferAsyncCore(cancellationToken);
    }

    void FlushBufferSync()
    {
        if (count == 0)
        {
            return;
        }

        target.Write(buffer, 0, count);
        count = 0;
    }

    async Task FlushBufferAsyncCore(CancellationToken cancellationToken)
    {
        var toWrite = count;
        count = 0;
        await target.WriteAsync(buffer.AsMemory(0, toWrite), cancellationToken);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposed)
        {
            return;
        }

        disposed = true;

        if (disposing)
        {
            try
            {
                FlushBufferSync();
            }
            finally
            {
                if (!leaveOpen)
                {
                    target.Dispose();
                }
            }
        }

        base.Dispose(disposing);
    }

    public override async ValueTask DisposeAsync()
    {
        if (disposed)
        {
            return;
        }

        disposed = true;

        try
        {
            await FlushAsync(default);
        }
        finally
        {
            if (!leaveOpen)
            {
                await target.DisposeAsync();
            }
        }
    }
}
