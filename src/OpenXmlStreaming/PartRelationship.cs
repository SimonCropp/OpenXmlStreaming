namespace OpenXmlStreaming;

/// <summary>
/// Describes a relationship to be added to a part when using <see cref="OpenXmlPackageWriter.WritePart"/>.
/// </summary>
public readonly struct PartRelationship(Uri targetUri, string relationshipType, TargetMode targetMode = TargetMode.Internal, string? id = null)
{
    public Uri TargetUri { get; } = targetUri;

    public string RelationshipType { get; } = relationshipType;

    public TargetMode TargetMode { get; } = targetMode;

    public string? Id { get; } = id;
}
