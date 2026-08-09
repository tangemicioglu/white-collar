# Real Office review artifacts

These are the final post-operation files produced by the live Office verification
harness on August 8, 2026. They are actual files opened, mutated, saved, and
reopened by Microsoft Word and PowerPoint; they are not fake fixtures.

* `white-collar-word-real-v0.1.docx` — final Word document after the registered
  live operation sequence.
* `white-collar-powerpoint-real-v0.1.pptx` — final PowerPoint deck after the
  registered live operation sequence.

The ignored `.real-word-artifacts/` and `.real-powerpoint-artifacts/` directories
retain the intermediate snapshots and screenshots used by the tests. These two
files are the concise handoff artifacts.
