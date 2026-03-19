# Hermes Architecture

## Overview

Hermes is a read-only DOCX viewer built as a Tauri desktop app with a Rust parsing backend and a vanilla HTML/CSS/JS frontend. The main runtime path is:

1. the frontend asks the Tauri backend to open a `.docx`
2. the Rust backend parses the DOCX archive into a structured `Document`
3. the `Document` is serialized to JSON and returned over Tauri IPC
4. the frontend renders that model into a paginated reading UI

The repository also contains an older/shared Rust crate at the repo root. The shipping desktop app uses the code under `src-tauri/`.

## System Diagram

```mermaid
flowchart LR
    U["User"] --> F["Frontend UI<br/>src/index.html + src/app.js + src/styles.css"]
    F -->|invoke('open_docx')| T["Tauri Commands<br/>src-tauri/src/main.rs"]
    T --> P["DOCX Parser<br/>src-tauri/src/parser.rs"]
    P --> Z["ZIP + OOXML Parts"]
    P --> M["Document Model<br/>src-tauri/src/model.rs"]
    M -->|JSON over IPC| F
    T --> R["Recent Files JSON<br/>app data directory"]
    F -->|invoke('get_recent_files')| T
    F -->|invoke('quit_app')| T
```

## Top-Level Layout

### Desktop App

- `src-tauri/src/main.rs`
  Tauri bootstrap, command registration, and recent-file persistence.
- `src-tauri/src/parser.rs`
  DOCX parsing pipeline for the desktop app.
- `src-tauri/src/model.rs`
  Serializable document model shared between the backend and frontend.
- `src/`
  Static frontend assets rendered inside the Tauri webview.

### Shared Root Crate

- `src-rust/lib.rs`
- `src-rust/comment_parser.rs`
- `src-rust/models.rs`
- `src-rust/main.rs`

This root crate appears to be a reusable parser/demo path rather than the primary desktop runtime. It overlaps conceptually with `src-tauri/src/parser.rs`, so the repo currently has two parsing codepaths.

## Runtime Flow

### 1. App Startup

`src-tauri/src/main.rs` builds the Tauri app, installs the dialog and filesystem plugins, and registers three commands:

- `open_docx`
- `get_recent_files`
- `quit_app`

In debug builds, the main webview opens devtools automatically.

### 2. File Selection and Loading

The frontend in `src/app.js` supports multiple entry points for loading a document:

- toolbar open button
- welcome-screen open button
- drag-and-drop via `tauri://drag-drop`
- recent files list

All of these converge on `loadDocument(path)`, which calls:

```js
invoke('open_docx', { path })
```

### 3. DOCX Parsing

`DocxParser::from_path` in `src-tauri/src/parser.rs`:

- opens the selected file
- validates basic size constraints
- ensures `word/document.xml` exists
- preloads relationship metadata

`DocxParser::parse` assembles a `Document` by calling dedicated parsing stages:

- `parse_document_body`
- `parse_styles`
- `parse_comments`
- `parse_headers_footers`
- `parse_footnotes`
- `parse_images`

This keeps the parser organized around OOXML parts instead of trying to handle everything in one pass.

### 4. IPC Boundary

The parsed Rust `Document` is serialized through `serde` and sent to the frontend as JSON. This model is the contract between backend and frontend.

The frontend assumes a tagged block model:

- `paragraph`
- `table`
- `page_break`

along with collections for comments, headers, footers, footnotes, styles, and images.

### 5. Rendering

`renderDocument(doc)` in `src/app.js` drives the UI refresh. It:

- swaps from the welcome screen to the document view
- splits body content into pages on explicit page breaks
- renders each page
- renders comments
- updates the window title and find state

Rendering is mostly composed from these functions:

- `splitIntoPages`
- `renderPage`
- `renderBlocks`
- `renderParagraph`
- `renderRun`
- `renderTable`
- `renderComments`

Headers and footers are currently rendered per page using the first available header/footer entry in the model.

## Backend Architecture

### Tauri Command Layer

The backend command layer in `src-tauri/src/main.rs` is intentionally thin:

- `open_docx` parses the file and stores it in recent files
- `get_recent_files` reads persisted history
- `quit_app` exits the application

This keeps document parsing in `parser.rs` and UI logic in the frontend.

### Parser Responsibilities

The parser is responsible for:

- opening the DOCX ZIP archive
- reading OOXML files by path
- resolving relationships for linked assets like images
- converting XML into a frontend-friendly `Document`
- failing with user-displayable error strings

Internally, the parser uses temporary context objects while walking XML, then converts them into stable model structs.

### Document Model

`src-tauri/src/model.rs` defines the central domain model:

- `Document`
- `BlockElement`
- `Run`
- `TableRow`
- `TableCell`
- `Comment`
- `HeaderFooter`
- `Footnote`
- `Style`

This model is the core architectural seam in the app. If the frontend changes, the parsing pipeline can remain mostly intact as long as this contract stays compatible.

### Persistence

The only explicit app-level persistence in the current desktop app is the recent-files list. It is stored as JSON in the Tauri app data directory:

- file name: `recent-files.json`
- maximum entries: `8`

This persistence is managed entirely in `src-tauri/src/main.rs`.

## Frontend Architecture

### UI Structure

`src/index.html` defines three main UI regions:

- the header and toolbar
- the welcome screen / recent files area
- the document view with the desk and comments sidebar

It also includes a floating find bar.

### State Model

The frontend in `src/app.js` uses module-level mutable state instead of a framework store:

- `currentDocument`
- `currentFilePath`
- `commentsVisible`
- `findMatches`
- `findIndex`
- `statusTimer`
- `findDebounceTimer`

For the current app size, this keeps the runtime simple and avoids introducing a build step or client framework.

### Rendering Strategy

The frontend renders directly into DOM nodes rather than using templates or a virtual DOM. The overall approach is:

- parse once on the backend
- render whole-page structures in the frontend
- update the full document view on each open

This is straightforward and easy to reason about, but it also means very large documents may eventually need incremental rendering or virtualization.

### Interaction Features

The frontend owns:

- keyboard shortcuts
- theme switching
- comment panel visibility
- in-document search and highlight navigation
- drag-and-drop affordances
- error and status banners

These behaviors are intentionally kept client-side because they mostly operate on already-loaded document state.

## Data and Control Boundaries

### Backend Owns

- filesystem access
- DOCX archive parsing
- OOXML interpretation
- image extraction and data URI generation
- recent-file persistence
- application lifecycle commands

### Frontend Owns

- page layout
- DOM rendering
- comment navigation UI
- theme state in the browser context
- local interaction state such as search matches and panel visibility

### Shared Contract

The `Document` JSON payload is the handoff between both sides. That contract is the most important point to protect when making changes.

## Important Design Choices

### Manual DOCX Parsing

Hermes parses DOCX files directly using `zip` and `quick-xml` rather than depending on a high-level DOCX parser abstraction. That choice gives the app control over:

- which OOXML parts are loaded
- how partial or unsupported content degrades
- how the frontend-facing model is shaped

### No Frontend Build Step

The frontend is plain HTML, CSS, and JavaScript. This keeps the app lightweight and easy to inspect, though it also means:

- less built-in structure for state management
- fewer abstractions for component reuse
- manual organization discipline matters more as the UI grows

### Read-Only Scope

The architecture is intentionally one-way:

- input file in
- parsed document model out
- render to UI

There is no editing pipeline, no document mutation layer, and no save/export path in the current design.

## Current Rough Edges

### Dual Parser Paths

The repo contains both:

- a root Rust crate in `src-rust/`
- the Tauri app parser in `src-tauri/src/`

That duplication can confuse ownership and make future parser changes harder to centralize.

### Tauri Dev Configuration

`src-tauri/tauri.conf.json` currently points development at `http://localhost:1430` while also setting `frontendDist` to `../src`. That suggests the development flow is not fully settled yet.

### Coarse Rendering Updates

The frontend rerenders the whole document on load. That is fine for the current scope, but search, comments, and very large documents may eventually benefit from more incremental strategies.

## Extension Points

If Hermes grows, the cleanest extension seams are:

- expanding the `Document` model in `src-tauri/src/model.rs`
- adding parser phases in `src-tauri/src/parser.rs`
- introducing clearer frontend view modules around rendering and interaction concerns
- consolidating the shared parser logic so the desktop app and root crate do not drift

## Suggested Near-Term Improvements

- choose a single canonical parser implementation
- make the Tauri dev/build story consistent
- separate frontend rendering logic from interaction logic into smaller modules
- add dedicated tests around the backend `Document` contract and parser stages

## Related Docs

- `docs/SPEC.md` for the product specification
- `README.md` for setup and repo-level usage
