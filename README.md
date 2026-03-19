# Hermes

Hermes is a lightweight, read-only DOCX viewer built with Rust and Tauri. The goal is simple: open Word documents quickly, render them with useful formatting intact, and avoid the overhead of launching a full editor when you only need to read.

## Installation

This repository currently documents installation from source. You will need:

- Rust and Cargo
- Tauri prerequisites for your platform
- on Windows, Microsoft Edge WebView2

### `tauri-cli` vs. the Desktop App

- `tauri-cli` is a developer tool. You install it so commands like `cargo tauri build` and `cargo tauri dev` work.
- The Hermes desktop app is the actual program you run to open `.docx` files.
- If you are building Hermes from source, you need `tauri-cli`.
- If you are just installing a prebuilt Hermes release, you would run the app directly and would not need `tauri-cli`.

If you do not already have the Tauri CLI installed:

```bash
cargo install tauri-cli
```

Build the desktop app:

```bash
cargo tauri build --manifest-path src-tauri/Cargo.toml
```

## How to Use

1. Launch Hermes.
2. Open a `.docx` file from the welcome screen or the toolbar.
3. You can also drag and drop a `.docx` file onto the window, or reopen a file from the recent files list.
4. Use the toolbar or keyboard shortcuts to search the document, show comments, and switch themes.

Useful shortcuts:

- `Ctrl+O` open a file
- `Ctrl+F` find in the current document
- `Ctrl+]` show or hide comments
- `Ctrl+D` toggle light and dark theme
- `Ctrl+Q` quit Hermes

## What Hermes Supports

Hermes is a read-only viewer and currently supports:

- paragraphs and styled text runs
- tables
- explicit page breaks
- comments
- headers and footers
- footnotes
- embedded images
- recent files

## Useful Notes

- Hermes opens `.docx` files for reading only. It does not edit or save documents.
- Recent files are stored locally by the desktop app so you can reopen them quickly.
- For architecture, development notes, and current technical limitations, see [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md).
- The product specification lives in [docs/SPEC.md](docs/SPEC.md).

## License

No license file is currently included in this repository.
