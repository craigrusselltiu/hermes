use crate::models::{Comment, Run, BlockElement};
use quick_xml::events::{Event, BytesStart};
use quick_xml::Reader;
use std::collections::HashMap;
use std::io::BufRead;

/// Parser for DOCX comments
/// 
/// Handles parsing of word/comments.xml and tracking comment ranges in document.xml
pub struct CommentParser {
    /// Map of comment ID to Comment data
    comments: HashMap<u32, Comment>,
    /// Stack to track current comment range being processed  
    comment_range_stack: Vec<u32>,
    /// Map of comment ranges to track which runs are commented
    active_comment_ranges: HashMap<u32, bool>,
}

impl CommentParser {
    pub fn new() -> Self {
        Self {
            comments: HashMap::new(),
            comment_range_stack: Vec::new(),
            active_comment_ranges: HashMap::new(),
        }
    }

    /// Parse comments.xml to extract comment metadata
    pub fn parse_comments_xml<R: BufRead>(&mut self, reader: R) -> Result<(), Box<dyn std::error::Error>> {
        let mut xml_reader = Reader::from_reader(reader);
        xml_reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut current_comment: Option<Comment> = None;
        let mut current_text = String::new();
        let mut in_comment_text = false;

        loop {
            match xml_reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    match e.name().as_ref() {
                        b"w:comment" => {
                            current_comment = Some(self.parse_comment_start(e)?);
                            in_comment_text = true;
                            current_text.clear();
                        }
                        b"w:t" if in_comment_text => {
                            // Start collecting text content
                        }
                        _ => {}
                    }
                }
                Ok(Event::Text(e)) if in_comment_text => {
                    current_text.push_str(&e.unescape()?);
                }
                Ok(Event::End(ref e)) => {
                    match e.name().as_ref() {
                        b"w:comment" => {
                            if let Some(mut comment) = current_comment.take() {
                                comment.text = current_text.clone();
                                self.comments.insert(comment.id, comment);
                            }
                            in_comment_text = false;
                            current_text.clear();
                        }
                        _ => {}
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(format!("Error parsing comments XML: {}", e).into()),
                _ => {}
            }

            buf.clear();
        }

        Ok(())
    }

    /// Parse a w:comment start element to extract id, author, and date
    fn parse_comment_start(&self, element: &BytesStart) -> Result<Comment, Box<dyn std::error::Error>> {
        let mut id = None;
        let mut author = String::new();
        let mut date = None;

        for attr in element.attributes() {
            let attr = attr?;
            match attr.key.as_ref() {
                b"w:id" => {
                    id = Some(std::str::from_utf8(&attr.value)?.parse::<u32>()?);
                }
                b"w:author" => {
                    author = std::str::from_utf8(&attr.value)?.to_string();
                }
                b"w:date" => {
                    date = Some(std::str::from_utf8(&attr.value)?.to_string());
                }
                _ => {}
            }
        }

        let id = id.ok_or("Comment missing required id attribute")?;

        Ok(Comment {
            id,
            author,
            date,
            text: String::new(), // Will be filled in later
        })
    }

    /// Process comment range markers in document.xml and update runs
    /// This should be called while parsing document runs
    pub fn process_comment_range_start(&mut self, comment_id: u32) {
        self.comment_range_stack.push(comment_id);
        self.active_comment_ranges.insert(comment_id, true);
    }

    /// Process comment range end markers in document.xml
    pub fn process_comment_range_end(&mut self, comment_id: u32) {
        // Remove from stack
        if let Some(pos) = self.comment_range_stack.iter().position(|&id| id == comment_id) {
            self.comment_range_stack.remove(pos);
        }
        self.active_comment_ranges.remove(&comment_id);
    }

    /// Update a run with comment references based on active comment ranges
    pub fn update_run_with_comments(&self, run: &mut Run) {
        // If there are active comment ranges, link this run to the most recent one
        if let Some(&comment_id) = self.comment_range_stack.last() {
            run.comment_ref = Some(comment_id);
        }
    }

    /// Extract element name from XML event (helper function)
    pub fn extract_element_name(element: &BytesStart) -> Result<Option<u32>, Box<dyn std::error::Error>> {
        for attr in element.attributes() {
            let attr = attr?;
            if attr.key.as_ref() == b"w:id" {
                let id = std::str::from_utf8(&attr.value)?.parse::<u32>()?;
                return Ok(Some(id));
            }
        }
        Ok(None)
    }

    /// Get all parsed comments
    pub fn get_comments(&self) -> Vec<Comment> {
        self.comments.values().cloned().collect()
    }

    /// Get comment by ID
    pub fn get_comment(&self, id: u32) -> Option<&Comment> {
        self.comments.get(&id)
    }

    /// Check if a comment ID is currently active (within a range)
    pub fn is_comment_active(&self, comment_id: u32) -> bool {
        self.active_comment_ranges.contains_key(&comment_id)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;

    #[test]
    fn test_parse_comments_xml() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="John Doe" w:date="2023-01-15T10:30:00Z">
    <w:p>
      <w:r>
        <w:t>This is a comment</w:t>
      </w:r>
    </w:p>
  </w:comment>
  <w:comment w:id="1" w:author="Jane Smith" w:date="2023-01-16T14:45:00Z">
    <w:p>
      <w:r>
        <w:t>Another comment</w:t>
      </w:r>
    </w:p>
  </w:comment>
</w:comments>"#;

        let mut parser = CommentParser::new();
        let cursor = Cursor::new(xml);
        
        parser.parse_comments_xml(cursor).expect("Failed to parse comments XML");
        
        let comments = parser.get_comments();
        assert_eq!(comments.len(), 2);
        
        let comment0 = parser.get_comment(0).unwrap();
        assert_eq!(comment0.author, "John Doe");
        assert_eq!(comment0.date, Some("2023-01-15T10:30:00Z".to_string()));
        assert_eq!(comment0.text, "This is a comment");
        
        let comment1 = parser.get_comment(1).unwrap();
        assert_eq!(comment1.author, "Jane Smith");
        assert_eq!(comment1.text, "Another comment");
    }

    #[test]
    fn test_comment_range_tracking() {
        let mut parser = CommentParser::new();
        
        // Start a comment range
        parser.process_comment_range_start(5);
        assert!(parser.is_comment_active(5));
        
        // Create a run and check it gets the comment reference
        let mut run = Run {
            text: "commented text".to_string(),
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
        };
        
        parser.update_run_with_comments(&mut run);
        assert_eq!(run.comment_ref, Some(5));
        
        // End the comment range
        parser.process_comment_range_end(5);
        assert!(!parser.is_comment_active(5));
    }

    #[test]
    fn test_nested_comment_ranges() {
        let mut parser = CommentParser::new();
        
        // Start multiple nested comment ranges
        parser.process_comment_range_start(1);
        parser.process_comment_range_start(2);
        
        let mut run = Run {
            text: "nested comment".to_string(),
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
        };
        
        // Should reference the most recent (innermost) comment
        parser.update_run_with_comments(&mut run);
        assert_eq!(run.comment_ref, Some(2));
        
        // End inner range
        parser.process_comment_range_end(2);
        
        // Now should reference outer comment
        parser.update_run_with_comments(&mut run);
        assert_eq!(run.comment_ref, Some(1));
        
        // End outer range
        parser.process_comment_range_end(1);
        
        // No more comments
        let mut run2 = Run {
            text: "no comment".to_string(),
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
        };
        parser.update_run_with_comments(&mut run2);
        assert_eq!(run2.comment_ref, None);
    }
}