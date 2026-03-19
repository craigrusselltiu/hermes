use serde::{Deserialize, Serialize};
use std::collections::HashMap;

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

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "type")]
pub enum BlockElement {
    Paragraph {
        runs: Vec<Run>,
        style: Option<String>,
        alignment: Option<String>,
    },
    Table {
        rows: Vec<TableRow>,
    },
    PageBreak,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Run {
    pub text: String,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
    pub font_size: Option<f32>,       // in pt
    pub font_family: Option<String>,
    pub color: Option<String>,        // hex
    pub highlight: Option<String>,
    pub comment_ref: Option<u32>,     // links to Comment.id
    pub footnote_ref: Option<u32>,    // links to Footnote.id
    pub image_id: Option<String>,     // links to images map
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableRow {
    pub cells: Vec<TableCell>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableCell {
    pub content: Vec<BlockElement>,
    pub col_span: u32,
    pub row_span: u32,
    pub shading: Option<String>, // background color hex
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Comment {
    pub id: u32,
    pub author: String,
    pub date: Option<String>,
    pub text: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct HeaderFooter {
    pub content: Vec<BlockElement>,
    pub section: u32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Footnote {
    pub id: u32,
    pub content: Vec<BlockElement>,
}

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