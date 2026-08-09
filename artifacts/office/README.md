# Real Office review artifacts

The live verification harness produces real `.docx` and `.pptx` evidence, but
those files are local-only. Office files can carry personal author metadata and
the repository also uses source-derived presentation templates during local
review, so binary artifacts are ignored and are not published with the source.

To reproduce the evidence on a Windows machine with Office installed, follow
the live Word and PowerPoint test commands in the main README. The ignored
`.real-word-artifacts/` and `.real-powerpoint-artifacts/` directories retain
screenshots and snapshots for local review.
