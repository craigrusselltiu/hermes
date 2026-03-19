use serde::{Deserialize, Serialize};
use std::collections::HashMap;

/// The main document model that represents a parsed DOCX file
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Document {
    pub body: Vec<BlockElement>,
    pub comments: Vec<Comment>,
    pub headers: Vec<HeaderFooter>,
    pub footers: Vec<HeaderFooter>,
    pub footnotes: Vec<Footnote>,
    pub styles: HashMap<String, Style>,
    pub images: HashMap<String, String>, // rId -> base64 data URI
}

/// Block-level elements that can appear in the document body
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "type")]
pub enum BlockElement {
    #[serde(rename = "paragraph")]
    Paragraph {
        runs: Vec<Run>,
        style: Option<String>,
        alignment: Option<String>,
    },
    #[serde(rename = "table")]
    Table { rows: Vec<TableRow> },
    #[serde(rename = "page_break")]
    PageBreak,
}

/// A run of text with consistent formatting
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Run {
    pub text: String,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
    pub font_size: Option<f32>,        // in pt
    pub font_family: Option<String>,
    pub color: Option<String>,         // hex
    pub highlight: Option<String>,
    pub comment_ref: Option<u32>,      // links to Comment.id
    pub footnote_ref: Option<u32>,     // links to Footnote.id
    pub image_id: Option<String>,      // links to images map
}

/// A row in a table
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableRow {
    pub cells: Vec<TableCell>,
}

/// A cell in a table
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableCell {
    pub content: Vec<BlockElement>,
    pub col_span: u32,
    pub row_span: u32,
    pub shading: Option<String>, // background color hex
}

/// A comment on the document
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Comment {
    pub id: u32,
    pub author: String,
    pub date: Option<String>,
    pub text: String,
}

/// Header or footer content
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct HeaderFooter {
    pub content: Vec<BlockElement>,
    pub section: u32,
}

/// A footnote
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Footnote {
    pub id: u32,
    pub content: Vec<BlockElement>,
}

/// Style definition with inheritance
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Style {
    pub based_on: Option<String>,
    pub font_size: Option<f32>,
    pub font_family: Option<String>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub color: Option<String>,
    pub alignment: Option<String>,
    pub heading_level: Option<u8>, // 1-6 for heading styles
}

impl Document {
    /// Creates an empty document
    pub fn new() -> Self {
        Self {
            body: Vec::new(),
            comments: Vec::new(),
            headers: Vec::new(),
            footers: Vec::new(),
            footnotes: Vec::new(),
            styles: HashMap::new(),
            images: HashMap::new(),
        }
    }
}

impl Default for Document {
    fn default() -> Self {
        Self::new()
    }
}

impl Run {
    /// Creates a new text run with default formatting
    pub fn new(text: String) -> Self {
        Self {
            text,
            bold: false,
            italic: false,
            underline: false,
            strikethrough: false,
            font_size: None,
            font_family: None,
            color: None,
            highlight: None,
            comment_ref: None,
            footnote_ref: None,
            image_id: None,
        }
    }

    /// Creates a text run with basic formatting
    pub fn with_formatting(
        text: String,
        bold: bool,
        italic: bool,
        underline: bool,
    ) -> Self {
        Self {
            text,
            bold,
            italic,
            underline,
            strikethrough: false,
            font_size: None,
            font_family: None,
            color: None,
            highlight: None,
            comment_ref: None,
            footnote_ref: None,
            image_id: None,
        }
    }
}

impl TableCell {
    /// Creates a new table cell with default properties
    pub fn new(content: Vec<BlockElement>) -> Self {
        Self {
            content,
            col_span: 1,
            row_span: 1,
            shading: None,
        }
    }
}

impl Style {
    /// Creates a new empty style
    pub fn new() -> Self {
        Self {
            based_on: None,
            font_size: None,
            font_family: None,
            bold: None,
            italic: None,
            color: None,
            alignment: None,
            heading_level: None,
        }
    }
}

impl Default for Style {
    fn default() -> Self {
        Self::new()
    }
}