namespace OpenXmlStreaming.Benchmarks;

/// <summary>
/// A write-only, non-seekable sink that discards everything written to it.
/// Used to isolate the cost of *producing* a package for a non-seekable
/// destination (e.g. an HTTP response body) from the cost of the destination
/// itself. Any attempt to seek or read throws.
///
/// The standard <see cref="DocumentFormat.OpenXml.Packaging.OpenXmlPackage"/>
/// path cannot write to this directly, forcing callers to buffer into a
/// <see cref="MemoryStream"/> first. The streaming writer can target it directly.
/// </summary>
internal sealed class NonSeekableDiscardStream : Stream
{
    long length;

    public override bool CanRead => false;
    public override bool CanSeek => false;
    public override bool CanWrite => true;
    public override long Length => throw new NotSupportedException();
    public override long Position
    {
        get => length;
        set => throw new NotSupportedException();
    }

    public override void Flush()
    {
    }

    public override int Read(byte[] buffer, int offset, int count) =>
        throw new NotSupportedException();

    public override long Seek(long offset, SeekOrigin origin) =>
        throw new NotSupportedException();

    public override void SetLength(long value) =>
        throw new NotSupportedException();

    public override void Write(byte[] buffer, int offset, int count) =>
        length += count;
}
