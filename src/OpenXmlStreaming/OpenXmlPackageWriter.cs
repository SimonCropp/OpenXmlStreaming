namespace OpenXmlStreaming;

/// <summary>
/// Writes an OPC package in forward-only mode to any writable stream,
/// including non-seekable streams. Produces valid .docx/.xlsx/.pptx files
/// without requiring a temporary <see cref="MemoryStream"/>.
/// </summary>
public sealed class OpenXmlPackageWriter :
    IAsyncDisposable,
    IDisposable
{
    /// <summary>
    /// Default size of the internal write buffer placed between <see cref="ZipArchive"/>
    /// and the caller-supplied target stream. Matches <see cref="Stream.CopyTo(Stream)"/>'s
    /// default copy-buffer size.
    /// </summary>
    public const int DefaultBufferSize = 81920;

    static XmlWriterSettings xmlSettings = new()
    {
        CloseOutput = false,
        Encoding = Encoding.UTF8,
    };

    BufferedWriteStream? bufferedStream;
    ZipArchive archive;
    List<(Uri PartUri, string ContentType)> contentTypes = [];
    List<StoredRelationship> relationships = [];
    HashSet<string> writtenParts = new(StringComparer.OrdinalIgnoreCase);
    OpenXmlPartEntry? currentEntry;
    bool finished;
    int nextRelId;

    /// <summary>
    /// Creates a writer targeting the given stream. The stream need not be seekable.
    /// </summary>
    /// <param name="stream">Target stream (write-only is sufficient).</param>
    /// <param name="leaveOpen">Whether to leave the stream open after disposal.</param>
    /// <param name="bufferSize">
    /// Size of the internal write buffer placed between <see cref="ZipArchive"/> and
    /// the target stream. Defaults to <see cref="DefaultBufferSize"/> (80 KB). Pass
    /// <c>0</c> to disable buffering and write directly to the target stream — this
    /// also disables async flushing on <see cref="DisposeAsync"/>, which would
    /// otherwise asynchronously push the final buffer contents to the target.
    /// Larger buffers reduce the number of writes that reach the target, which
    /// matters most for remote sinks like SQL BLOB streams or cloud upload streams.
    /// </param>
    public OpenXmlPackageWriter(Stream stream, bool leaveOpen = false, int bufferSize = DefaultBufferSize)
    {
        if (bufferSize > 0)
        {
            bufferedStream = new(stream, bufferSize, leaveOpen);
            archive = new(bufferedStream, ZipArchiveMode.Create, leaveOpen: true);
        }
        else
        {
            archive = new(stream, ZipArchiveMode.Create, leaveOpen);
        }
    }

    /// <summary>
    /// Adds a package-level relationship (written to _rels/.rels on finalization).
    /// </summary>
    public string AddRelationship(Uri partUri, string relationshipType, string? id = null)
    {
        ThrowIfFinished();

        id ??= "rId" + (++nextRelId).ToString(CultureInfo.InvariantCulture);

        relationships.Add(new(id, partUri, relationshipType, TargetMode.Internal));
        return id;
    }

    /// <summary>
    /// Creates a new part entry in the package and returns a context object
    /// for writing content and part-level relationships.
    /// Only one part may be open at a time. The previous part is
    /// automatically closed when a new part is created or when the writer is disposed.
    /// </summary>
    public OpenXmlPartEntry CreatePart(Uri partUri, string contentType)
    {
        ThrowIfFinished();

        currentEntry?.Dispose();

        var entryPath = partUri.OriginalString.TrimStart('/');

        if (!writtenParts.Add(entryPath))
        {
            throw new InvalidOperationException(string.Format(CultureInfo.InvariantCulture, "A part with URI '{0}' has already been written.", partUri));
        }

        contentTypes.Add((partUri, contentType));

        var zipEntry = archive.CreateEntry(entryPath, CompressionLevel.Optimal);
        var stream = zipEntry.Open();

        currentEntry = new(archive, partUri, stream);
        return currentEntry;
    }

    /// <summary>
    /// Creates a new part, writes the given element tree as its content,
    /// and closes the part in one call.
    /// </summary>
    public void WritePart(Uri partUri, string contentType, OpenXmlElement rootElement, IEnumerable<PartRelationship>? relationships = null)
    {
        using var entry = CreatePart(partUri, contentType);

        if (relationships is not null)
        {
            foreach (var rel in relationships)
            {
                entry.AddRelationship(rel.TargetUri, rel.RelationshipType, rel.TargetMode, rel.Id);
            }
        }

        using var xmlWriter = XmlWriter.Create(entry.Stream, xmlSettings);

        rootElement.WriteTo(xmlWriter);
    }

    /// <summary>
    /// Writes _rels/.rels and [Content_Types].xml into the ZIP. Called by
    /// <see cref="Dispose"/> and <see cref="DisposeAsync"/>; does not itself
    /// write the ZIP central directory (the archive must be disposed for that).
    /// </summary>
    internal void Finish()
    {
        if (finished)
        {
            return;
        }

        finished = true;

        currentEntry?.Dispose();
        currentEntry = null;

        WritePackageRelationships();
        WriteContentTypes();
    }

    /// <summary>
    /// Disposes the writer, finalizing the package if not already done.
    /// </summary>
    public void Dispose()
    {
        Finish();
        archive.Dispose();
        bufferedStream?.Dispose();
    }

    /// <summary>
    /// Asynchronously disposes the writer, finalizing the package if not already done.
    /// When an internal write buffer is in use (the default), the final buffer contents
    /// — which include the ZIP central directory — are flushed to the target stream
    /// via <see cref="Stream.WriteAsync(ReadOnlyMemory{byte}, CancellationToken)"/>,
    /// so the calling thread is not blocked on network I/O during the final flush.
    /// </summary>
    public async ValueTask DisposeAsync()
    {
        Finish();
        await archive.DisposeAsync();

        if (bufferedStream is not null)
        {
            await bufferedStream.DisposeAsync();
        }
    }

    /// <summary>
    /// Asynchronously flushes any bytes currently sitting in the internal write
    /// buffer to the target stream via
    /// <see cref="Stream.WriteAsync(ReadOnlyMemory{byte}, CancellationToken)"/>.
    /// Useful between <see cref="WritePart"/> calls to push accumulated bytes to
    /// a remote sink at part boundaries. A no-op when the writer is unbuffered
    /// (<c>bufferSize: 0</c>) or the buffer is empty.
    /// </summary>
    /// <remarks>
    /// This does not eliminate intermediate sync writes that occur when a single
    /// <see cref="WritePart"/> call produces more bytes than the buffer can hold
    /// — those spill synchronously during serialization regardless. Use a larger
    /// buffer if you need to avoid that.
    /// </remarks>
    public Task FlushAsync(Cancel cancel = default)
    {
        if (bufferedStream is null)
        {
            return Task.CompletedTask;
        }

        return bufferedStream.FlushAsync(cancel);
    }

    void WritePackageRelationships()
    {
        if (relationships.Count == 0)
        {
            return;
        }

        var entry = archive.CreateEntry("_rels/.rels", CompressionLevel.Optimal);
        using var stream = entry.Open();
        WriteRelationshipsXml(stream, relationships);
    }

    internal static void WriteRelationshipsXml(Stream stream, List<StoredRelationship> relationships)
    {
        using var writer = XmlWriter.Create(stream, xmlSettings);

        writer.WriteStartDocument();
        writer.WriteStartElement("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships");

        foreach (var rel in relationships)
        {
            writer.WriteStartElement("Relationship");
            writer.WriteAttributeString("Id", rel.Id);
            writer.WriteAttributeString("Type", rel.RelationshipType);
            writer.WriteAttributeString("Target", rel.TargetUri.OriginalString);

            if (rel.TargetMode == TargetMode.External)
            {
                writer.WriteAttributeString("TargetMode", "External");
            }

            writer.WriteEndElement();
        }

        writer.WriteEndElement();
    }

    void WriteContentTypes()
    {
        var entry = archive.CreateEntry("[Content_Types].xml", CompressionLevel.Optimal);
        using var stream = entry.Open();
        using var writer = XmlWriter.Create(stream, xmlSettings);

        writer.WriteStartDocument();
        writer.WriteStartElement("Types", "http://schemas.openxmlformats.org/package/2006/content-types");

        writer.WriteStartElement("Default");
        writer.WriteAttributeString("Extension", "rels");
        writer.WriteAttributeString("ContentType", "application/vnd.openxmlformats-package.relationships+xml");
        writer.WriteEndElement();

        writer.WriteStartElement("Default");
        writer.WriteAttributeString("Extension", "xml");
        writer.WriteAttributeString("ContentType", "application/xml");
        writer.WriteEndElement();

        foreach (var (partUri, contentType) in contentTypes)
        {
            writer.WriteStartElement("Override");
            var partName = GetPartName(partUri);
            writer.WriteAttributeString("PartName", partName);
            writer.WriteAttributeString("ContentType", contentType);
            writer.WriteEndElement();
        }

        writer.WriteEndElement();
    }

     static string GetPartName(Uri partUri)
    {
        var originalString = partUri.OriginalString;
        if (originalString.Length > 0 &&
            originalString[0] == '/')
        {
            return originalString;
        }

        return "/" + originalString;
    }

    void ThrowIfFinished()
    {
        if (finished)
        {
            throw new InvalidOperationException("The package writer has already been finalized. No more parts can be added.");
        }
    }
}
