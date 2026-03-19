# Hermes

Hermes is a lightweight, read-only DOCX viewer built with Rust and Tauri. The goal is simple: open Word documents quickly, render them with useful formatting intact, and avoid the overhead of launching a full editor when you only need to read.

The repository currently contains two related Rust codepaths:

- a shared Rust crate at the repo root for DOCX parsing experiments and reusable types
- a Tauri desktop app in `src-tauri/` that opens `.docx` files and renders them in a native shell with an HTML/CSS/JS frontend from `src/`

## What Hermes Does

Hermes is designed to:

- open `.docx` files from a native file picker
- support drag-and-drop opening
- render document content in a page-like reading layout
- show comments in a side panel
- keep a recent files list
- provide theme toggling and in-document find
- stay read-only

Based on the current parser and document model, the app is intended to handle:

- paragraphs and styled text runs
- tables
- explicit page breaks
- comments
- headers and footers
- footnotes
- embedded images
- style extraction for rendering

## Repository Layout

```text
.
|- docs/
|  `- SPEC.md              Product and architecture notes
|- src/
|  |- index.html           Frontend shell
|  |- app.js               Viewer UI and Tauri integration
|  `- styles.css           Viewer styling
|- src-rust/
|  |- lib.rs               Shared parser library entry point
|  |- main.rs              Small parser demo binary
|  |- comment_parser.rs    Comment parsing logic
|  `- models.rs            Shared data types
|- src-tauri/
|  |- src/
|  |  |- main.rs           Tauri commands and app bootstrap
|  |  |- parser.rs         DOCX parsing for the desktop app
|  |  `- model.rs          Serialized document model
|  |- tauri.conf.json      Tauri configuration
|  `- Cargo.toml           Desktop app crate
|- Cargo.toml              Root crate manifest
`- config.yaml             Project configuration
```

## Tech Stack

- Rust for parsing, data modeling, and desktop app logic
- Tauri 2 for the desktop shell
- `zip` and `quick-xml` for DOCX archive and OOXML parsing
- `serde` and `serde_json` for document model serialization
- vanilla HTML, CSS, and JavaScript for the frontend

## Development

### Prerequisites

You will need:

- Rust and Cargo
- Tauri development prerequisites for your platform
- on Windows, Microsoft Edge WebView2

If you do not already have the Tauri CLI installed:

```bash
cargo install tauri-cli
```

### Root Crate

The root crate contains shared parsing code and a small demo binary.

Run tests:

```bash
cargo test
```

Run the demo binary:

```bash
cargo run
```

### Desktop App

The desktop app lives in `src-tauri/` and uses the static frontend files in `src/`.

Check the desktop crate:

```bash
cargo check --manifest-path src-tauri/Cargo.toml
```

Build the desktop app:

```bash
cargo tauri build --manifest-path src-tauri/Cargo.toml
```

For development, note that the checked-in Tauri config currently includes:

- `frontendDist: ../src`
- `devUrl: http://localhost:1430`

That means `cargo tauri dev` currently expects a frontend dev server at `http://localhost:1430`. If you want to run the app without a separate dev server, you will likely need to adjust `src-tauri/tauri.conf.json` so development points at the local static frontend instead.

## How It Works

Hermes opens a `.docx` file as a ZIP archive, reads the OOXML parts inside it, and converts that data into a Rust `Document` model. That model is serialized to JSON and passed to the frontend through Tauri commands, where it is rendered into a paginated reading view.

At a high level:

1. the Tauri command receives a file path
2. the Rust parser opens the DOCX archive
3. XML parts such as `word/document.xml`, styles, comments, and relationships are parsed
4. images are converted into data URIs
5. a structured document model is returned to the frontend
6. the frontend renders pages, comments, and document chrome

## Current Notes

- Hermes is a viewer, not an editor.
- The checked-in spec lives at `docs/SPEC.md`.
- The repository currently has both a shared parser crate and a Tauri-specific parser implementation.
- The app stores recent files in the platform app data directory through the Tauri backend.

## Limitations

Hermes is intentionally narrower than Microsoft Word. Some important boundaries:

- no editing or saving back to DOCX
- no goal of full OOXML compatibility
- no support for macros or embedded OLE content
- development setup is still a little rough around the Tauri dev configuration

## License

No license file is currently included in this repository.
