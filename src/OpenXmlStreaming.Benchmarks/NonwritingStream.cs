/// <summary>
/// A stream implementation used for benchmarking that doesn't actually store
/// any data. Reports CanSeek=true so the standard OpenXml SDK path (which
/// requires a seekable stream) can operate against it, but discards all writes
/// to isolate the benchmark from file or memory-copy overhead.
/// </summary>
class NonwritingStream : Stream
{
    long length;

    public override bool CanRead => true;
    public override bool CanSeek => true;
    public override bool CanWrite => true;
    public override long Length => length;
    public override long Position { get; set; }

    public override void Flush()
    {
    }

    public override int Read(byte[] buffer, int offset, int count) =>
        throw new NotImplementedException();

    public override long Seek(long offset, SeekOrigin origin)
    {
        switch (origin)
        {
            case SeekOrigin.Begin:
                Position = offset;
                break;
            case SeekOrigin.Current:
                Position += offset;
                break;
            case SeekOrigin.End:
                throw new NotImplementedException();
        }

        return Position;
    }

    public override void SetLength(long value) =>
        length = value;

    public override void Write(byte[] buffer, int offset, int count)
    {
        length += count;
        Position += count;
    }
}
