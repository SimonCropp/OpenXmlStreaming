namespace OpenXmlStreaming;

/// <summary>
/// Describes a relationship to be added to a part when using <see cref="OpenXmlPackageWriter.WritePart"/>.
/// </summary>
// Id is required (non-nullable). Unlike OpenXmlPackageWriter.AddRelationship and
// OpenXmlPartEntry.AddRelationship — both of which accept a null id and return the
// auto-assigned value — PartRelationship is consumed by WritePart, which has no way
// to hand the assigned id back to the caller. By the time WritePart is invoked the
// part body has already been built, and that body almost always references its own
// relationships by id (hyperlinks, image refs, footer refs, …) — so the caller must
// know the id up front. Forcing it to be supplied here removes a footgun where the
// auto-assigned id would silently be discarded.
public readonly struct PartRelationship(Uri targetUri, string relationshipType, string id, TargetMode targetMode = TargetMode.Internal)
{
    public Uri TargetUri { get; } = targetUri;

    public string RelationshipType { get; } = relationshipType;

    public string Id { get; } = id;

    public TargetMode TargetMode { get; } = targetMode;
}
