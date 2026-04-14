namespace OpenXmlStreaming;

readonly record struct StoredRelationship(string Id, Uri TargetUri, string RelationshipType, TargetMode TargetMode);
