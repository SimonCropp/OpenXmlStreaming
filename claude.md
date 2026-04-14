# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Working directory

All build, test, and tool invocations run from `src/`.

## Commands

```bash
# Build
dotnet build src/OpenXmlStreaming.slnx

# Full test suite (there is no test filtering infrastructure beyond --filter)
dotnet test src/OpenXmlStreaming.Tests/OpenXmlStreaming.Tests.csproj

# A single test / class / keyword
dotnet test src/OpenXmlStreaming.Tests/OpenXmlStreaming.Tests.csproj --filter "FullyQualifiedName~MigrationGuide.WordStreaming"

# Benchmarks (default BDN job — publishable, slow)
dotnet run -c Release --project src/OpenXmlStreaming.Benchmarks -- --filter "*"

# Benchmarks (ShortRun — ~1-2 min, acceptable noise)
dotnet run -c Release --project src/OpenXmlStreaming.Benchmarks -- --job short --filter "*"

# Benchmarks (dry run — smoke-test the runner only, numbers are meaningless)
dotnet run -c Release --project src/OpenXmlStreaming.Benchmarks -- --job dry --filter "*Word_Simple*"
```

`dotnet-tools.json` installs `MarkdownSnippets.Tool` as a local tool; `mdsnippets` also runs automatically via the `MarkdownSnippets.MsBuild` package reference on the Tests project, so **every test build rewrites `/readme.md`** with the latest snippet content. If you change a snippet-marked test body, rebuild the Tests project to propagate the change to the readme.

## Architecture

### Library (`src/OpenXmlStreaming/`)

Three layers, stacked:

1. **`OpenXmlPackageWriter`** — forward-only OPC package writer built on `ZipArchive` in Create mode (which emits ZIP data descriptors, so the target stream does not need to be seekable). Writes parts via `CreatePart` / `WritePart`, tracks part-level relationships via the returned `OpenXmlPartEntry`, and writes `[Content_Types].xml` + `_rels/.rels` during disposal. `Finish()` is `internal` — callers should use `Dispose`/`DisposeAsync`, not call finalization explicitly. Tests see `Finish` via `[InternalsVisibleTo]` in the csproj (the assembly is strong-named, so the public-key token is required).

2. **`BufferedWriteStream`** — write-only buffered adapter that sits between `ZipArchive` and the caller's target stream. The `OpenXmlPackageWriter` constructor takes a `bufferSize` (default 80 KB, `OpenXmlPackageWriter.DefaultBufferSize`); passing `0` opts out entirely. `BufferedWriteStream` overrides both `Write` and `WriteAsync` (the latter is important — the base `Stream.WriteAsync` default falls back to sync `Write`, which defeats the async contract the writer needs). Sync writes from `ZipArchive` accumulate; spill flushes use `target.Write`; the final flush during `DisposeAsync` goes through `target.WriteAsync`, which is where the async win lives. Intermediate sync spills during large `WritePart` calls are unavoidable — `ZipArchive`'s write surface is sync-only, and the only mitigation is a bigger buffer.

3. **Higher-level builders** — `StreamingWorkbookBuilder`, `StreamingWordDocumentBuilder`, `StreamingPresentationBuilder`. Each wraps `OpenXmlPackageWriter` with format-aware part URI / rId allocation and composes the main part (`xl/workbook.xml`, `word/document.xml`, `ppt/presentation.xml`) in `Finish` / `DisposeAsync` using the tracked parts. `StreamingWordDocumentBuilder` is asymmetric — sub-part methods (`AddHeader`, `AddFooter`, …) return the relationship id because the caller needs it to construct content references (`FooterReference.Id`, etc.), and `WriteDocument` is explicit rather than dispose-triggered. `StreamingPresentationBuilder` embeds a minimal default theme + slide master + slide layout which is written lazily on the first `AddSlide` call (idempotent, so an empty presentation still produces a valid `.pptx`).

`StreamingDocument.CreateWord`/`CreateSpreadsheet`/`CreatePresentation` are thin factories that pre-register the package-level `officeDocument` relationship pointing at the main part URI. They exist for callers using the low-level writer directly; the builders use them internally.

### Tests (`src/OpenXmlStreaming.Tests/`)

**NUnit + Verify.OpenXml.** All round-trip tests call `Verify(stream, extension: "docx"/"xlsx"/"pptx")`; the `Verify.OpenXml` plugin opens the stream via the matching SDK `XxxDocument.Open`, normalises it via `DeterministicPackage` (from the sibling `DeterministicIoPackaging` repo), and snapshots both the binary package and extracted text/csv/info.

**Namespace collisions force partial test classes.** `DocumentFormat.OpenXml.Wordprocessing`, `DocumentFormat.OpenXml.Spreadsheet`, and `DocumentFormat.OpenXml.Presentation` (plus `DocumentFormat.OpenXml.Drawing`) share many type names (`Row`, `Cell`, `Text`, `Shape`, `TextBody`, `ColorMap`, `Bold`, `FontSize`, …). As a result:

- **`Wordprocessing` is deliberately NOT in `GlobalUsings.cs`.** Adding it back breaks any test file that also imports `Spreadsheet` or `Drawing`.
- Test fixtures that cover multiple document types (`MigrationGuide`, `BuilderTests`) are **split into partial classes** — one file per document type — with a file-level `using` for its namespace. See `MigrationGuide.Word.cs` / `.Spreadsheet.cs` / `.Presentation.cs`.
- `MigrationGuide.Presentation.cs` and `BuilderTests.Presentation.cs` additionally alias `Drawing = DocumentFormat.OpenXml.Drawing;` because `Presentation` and `Drawing` collide on ~15 type names and there is no way to import both as top-level usings. Presentation types are unqualified; Drawing types use `Drawing.TypeName`.
- Single-document-type files (e.g. `BufferedWriteStreamTests.cs`) use a file-level `using` for the relevant namespace.
- `OpenXmlPackageWriterTests.cs` and `Samples.cs` use traditional namespace aliases (`P =`, `S =`) because they intentionally mix multiple document types in a single file.

**Test-time sinks.** `NonSeekableStream.cs` wraps another stream with `CanSeek=false` to exercise the non-seekable write path. `SyncAsyncTrackingStream.cs` counts sync vs async `Write` calls and is the harness for validating async dispose / flush behaviour.

### Readme generation (`mdsnippets`)

The readme's code examples are **all snippet-backed**. Each `snippet: <name>` directive in `readme.md` is replaced by mdsnippets (at build time) with the body of a `// begin-snippet: <name>` … `// end-snippet` region in a test file. Subtle consequences:

- **Snippets live inside tests that actually run** — readme examples can't rot silently. `Samples.cs` holds the reusable "how-to" samples; `MigrationGuide.*.cs` holds the side-by-side before/after migration examples.
- `mdsnippets.json` at `src/mdsnippets.json` configures the runner (`InPlaceOverwrite`, 100-column width, TOC exclusions).
- When adding a new snippet-backed example, also add a corresponding `snippet: <name>` directive in `readme.md` and rebuild the Tests project to inject it.

## Benchmarks (`src/OpenXmlStreaming.Benchmarks/`)

Two benchmark classes:

- **`ForwardOnlyBenchmarks`** — per-format × per-size (Simple/Medium/Complex) pairs comparing `XxxDocument.Create` (Standard) against `StreamingDocument.CreateXxx` (ForwardOnly). Uses `NonwritingStream` (a discarding seekable stream) to isolate writer CPU + allocation cost from sink I/O. Ported from the upstream [Open-XML-SDK PR #2058](https://github.com/dotnet/Open-XML-SDK/pull/2058) that originally proposed this API.
- **`IoScenarioBenchmarks`** — four classes (`NonSeekableWordBenchmarks`, `NonSeekableSpreadsheetBenchmarks`, `FileWordBenchmarks`, `FileSpreadsheetBenchmarks`) measuring the real-world use case: large documents written to either a non-seekable sink (modeled with `NonSeekableDiscardStream`) or a temp file. The Standard side of the non-seekable scenario uses the idiomatic workaround — buffer to `MemoryStream` first, then `CopyTo` the sink — because the SDK can't target non-seekable streams directly.

Each IoScenario pair lives in its own class so BenchmarkDotNet can compute a meaningful per-pair `Ratio` column (one `[Baseline = true]` per class).

## Coupling with sibling repos

This library has hard behavioural dependencies on two SimonCropp repos that travel alongside it:

- **`C:\Code\Verify.OpenXml`** — provides the pptx/xlsx/docx stream converters that `Verify.OpenXml 1.7.0` uses. The pptx converter was added specifically to support this library's `CreatePresentation_RoundTrips` test. If that test fails after a Verify.OpenXml update, check whether the pptx converter was changed.
- **`C:\Code\DeterministicIoPackaging`** — provides `DeterministicPackage.Convert`, used by Verify.OpenXml to normalise the ZIP output before snapshotting. PowerPoint determinism support (the `PptxRelationshipPatcher` / `PptxContentPatcher`) was added specifically so this library's pptx snapshots don't drift run-to-run. If pptx snapshots start flapping, verify that DeterministicIoPackaging's ppt/ patchers still run.

## Strong naming and `InternalsVisibleTo`

The main assembly is strong-named (`key.snk` at `src/key.snk`, signing configured via `ProjectDefaults` package). `OpenXmlStreaming.csproj` grants `InternalsVisibleTo` to `OpenXmlStreaming.Tests` **with the full public-key token** — a bare assembly name is rejected by the compiler when the granting assembly is signed. If you add a new `InternalsVisibleTo`, extract the public key via `sn -tp key.snk` and include it.

## Intentional constraints / gotchas

- **Builder disposal inside test bodies.** Tests that use `await using var builder = …` followed by `await Verify(stream, …)` fail at runtime because the builder disposes at method exit, *after* `Verify` has already read the stream. Wrap the builder in an explicit `await using (var builder = …) { … }` scope so disposal happens before `Verify` runs. See `MigrationGuide.*Builder` tests for the pattern.
- **`WritePart` is sync by design.** There is no `WritePartAsync` because the underlying `OpenXmlElement.WriteTo(XmlWriter)` and `ZipArchive` write paths are sync, and wrapping them in a `Task` would be actively misleading — the calling thread would still block inside the serialisation call. The async surface exists only at `DisposeAsync` / `FlushAsync` boundaries, where genuine async I/O happens.
- **`bufferSize: 0` disables async flushing.** Without the internal buffer, `DisposeAsync` has nothing to flush asynchronously, so every `ZipArchive` sync write lands directly on the target. Use this only when the target is already an in-memory stream where the extra copy isn't worth it.
