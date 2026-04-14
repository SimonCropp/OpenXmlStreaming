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
    readonly ZipArchive archive;
    readonly List<(Uri PartUri, string ContentType)> contentTypes = [];
    readonly List<(string Id, Uri TargetUri, string RelationshipType, TargetMode TargetMode)> packageRelationships = [];
    readonly HashSet<string> writtenParts = new(StringComparer.OrdinalIgnoreCase);
    OpenXmlPartEntry? currentEntry;
    bool finished;
    int nextRelId;

    /// <summary>
    /// Creates a writer targeting the given stream. The stream need not be seekable.
    /// </summary>
    /// <param name="stream">Target stream (write-only is sufficient).</param>
    /// <param name="leaveOpen">Whether to leave the stream open after disposal.</param>
    public OpenXmlPackageWriter(Stream stream, bool leaveOpen = false)
    {
        if (stream is null)
        {
            throw new ArgumentNullException(nameof(stream));
        }

        archive = new(stream, ZipArchiveMode.Create, leaveOpen);
    }

    /// <summary>
    /// Adds a package-level relationship (written to _rels/.rels on finalization).
    /// </summary>
    public string AddRelationship(Uri partUri, string relationshipType, string? id = null)
    {
        ThrowIfFinished();

        if (partUri is null)
        {
            throw new ArgumentNullException(nameof(partUri));
        }

        if (relationshipType is null)
        {
            throw new ArgumentNullException(nameof(relationshipType));
        }

        id ??= "rId" + (++nextRelId).ToString(CultureInfo.InvariantCulture);

        packageRelationships.Add((id, partUri, relationshipType, TargetMode.Internal));
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

        if (partUri is null)
        {
            throw new ArgumentNullException(nameof(partUri));
        }

        if (contentType is null)
        {
            throw new ArgumentNullException(nameof(contentType));
        }

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
        if (rootElement is null)
        {
            throw new ArgumentNullException(nameof(rootElement));
        }

        using var entry = CreatePart(partUri, contentType);

        if (relationships is not null)
        {
            foreach (var rel in relationships)
            {
                entry.AddRelationship(rel.TargetUri, rel.RelationshipType, rel.TargetMode, rel.Id);
            }
        }

        using var xmlWriter = XmlWriter.Create(
            entry.Stream,
            new()
        {
            CloseOutput = false,
            Encoding = Encoding.UTF8,
        });

        rootElement.WriteTo(xmlWriter);
    }

    /// <summary>
    /// Finalizes the package by writing _rels/.rels and [Content_Types].xml.
    /// Called automatically by Dispose, but can be called explicitly.
    /// After this call, no more parts can be added.
    /// </summary>
    public void Finish()
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
    }

#if NET6_0_OR_GREATER
    /// <summary>
    /// Asynchronously disposes the writer, finalizing the package if not already done.
    /// </summary>
    public ValueTask DisposeAsync()
    {
        Dispose();
        return default;
    }
#endif

    void WritePackageRelationships()
    {
        if (packageRelationships.Count == 0)
        {
            return;
        }

        var entry = archive.CreateEntry("_rels/.rels", CompressionLevel.Optimal);
        using var stream = entry.Open();
        WriteRelationshipsXml(stream, packageRelationships);
    }

    internal static void WriteRelationshipsXml(Stream stream, List<(string Id, Uri TargetUri, string RelationshipType, TargetMode TargetMode)> relationships)
    {
        using var writer = XmlWriter.Create(
            stream,
            new()
        {
            CloseOutput = false,
            Encoding = Encoding.UTF8,
        });

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
        using var writer = XmlWriter.Create(
            stream,
            new()
        {
            CloseOutput = false,
            Encoding = Encoding.UTF8,
        });

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

        foreach (var ct in contentTypes)
        {
            writer.WriteStartElement("Override");
            var partName = ct.PartUri.OriginalString;
            writer.WriteAttributeString("PartName", partName.Length > 0 && partName[0] == '/' ? partName : "/" + partName);
            writer.WriteAttributeString("ContentType", ct.ContentType);
            writer.WriteEndElement();
        }

        writer.WriteEndElement();
    }

    void ThrowIfFinished()
    {
        if (finished)
        {
            throw new InvalidOperationException("The package writer has already been finalized. No more parts can be added.");
        }
    }
}
