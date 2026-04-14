namespace OpenXmlStreaming;

/// <summary>
/// Represents a part being written in forward-only mode.
/// Provides access to the part's output stream and the ability to add part-level relationships.
/// </summary>
public sealed class OpenXmlPartEntry : IDisposable
{
    ZipArchive archive;
    Uri partUri;
    Stream stream;
    List<PartRelationship>? relationships;
    bool disposed;
    int nextRelId;

    internal OpenXmlPartEntry(ZipArchive archive, Uri partUri, Stream stream)
    {
        this.archive = archive;
        this.partUri = partUri;
        this.stream = stream;
    }

    /// <summary>
    /// Gets the writable stream for this part's ZIP entry.
    /// </summary>
    public Stream Stream
    {
        get
        {
            ThrowIfDisposed();
            return stream;
        }
    }

    /// <summary>
    /// Adds a relationship for this part.
    /// </summary>
    public string AddRelationship(Uri targetUri, string relationshipType, TargetMode targetMode = TargetMode.Internal, string? id = null)
    {
        ThrowIfDisposed();

        relationships ??= [];

        id ??= "rId" + (++nextRelId).ToString(CultureInfo.InvariantCulture);

        relationships.Add(new(targetUri, relationshipType, id, targetMode));
        return id;
    }

    /// <summary>
    /// Closes the part stream and writes the part's .rels file if any relationships were added.
    /// </summary>
    public void Dispose()
    {
        if (disposed)
        {
            return;
        }

        disposed = true;
        stream.Dispose();

        if (relationships is not null &&
            relationships.Count > 0)
        {
            WriteRelationships();
        }
    }

    void WriteRelationships()
    {
        var relsPath = GetRelsPath();

        var relsEntry = archive.CreateEntry(relsPath, CompressionLevel.Optimal);
        using var relsStream = relsEntry.Open();
        OpenXmlPackageWriter.WriteRelationshipsXml(relsStream, relationships!);
    }

    string GetRelsPath()
    {
        var partPath = partUri.OriginalString.AsSpan().TrimStart('/');
        var lastSlash = partPath.LastIndexOf('/');

        if (lastSlash == -1)
        {
            return string.Concat("_rels/", partPath, ".rels");
        }
        var dir = partPath[..(lastSlash + 1)];
        var file = partPath[(lastSlash + 1)..];
        return string.Concat(dir, "_rels/", file, ".rels");
    }

    void ThrowIfDisposed()
    {
        if (disposed)
        {
            throw new ObjectDisposedException(nameof(OpenXmlPartEntry));
        }
    }
}
