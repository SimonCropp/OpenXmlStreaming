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
    List<(string Id, Uri TargetUri, string RelationshipType, TargetMode TargetMode)>? relationships;
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

        if (targetUri is null)
        {
            throw new ArgumentNullException(nameof(targetUri));
        }

        if (relationshipType is null)
        {
            throw new ArgumentNullException(nameof(relationshipType));
        }

        relationships ??= [];

        id ??= "rId" + (++nextRelId).ToString(CultureInfo.InvariantCulture);

        relationships.Add((id, targetUri, relationshipType, targetMode));
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
        var partPath = partUri.OriginalString.TrimStart('/');
        var lastSlash = partPath.LastIndexOf('/');
        string relsPath;

        if (lastSlash >= 0)
        {
            var dir = partPath[..(lastSlash + 1)];
            var file = partPath[(lastSlash + 1)..];
            relsPath = string.Concat(dir, "_rels/", file, ".rels");
        }
        else
        {
            relsPath = string.Concat("_rels/", partPath, ".rels");
        }

        var relsEntry = archive.CreateEntry(relsPath, CompressionLevel.Optimal);
        using var relsStream = relsEntry.Open();
        OpenXmlPackageWriter.WriteRelationshipsXml(relsStream, relationships!);
    }

    void ThrowIfDisposed()
    {
        if (disposed)
        {
            throw new ObjectDisposedException(nameof(OpenXmlPartEntry));
        }
    }
}
