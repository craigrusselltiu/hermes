// Prevents additional console window on Windows in release, DO NOT REMOVE!!
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use hermes_app::model::Document;
use hermes_app::parser::DocxParser;
use tauri::Manager;

#[tauri::command]
fn open_docx(path: String) -> Result<Document, String> {
    println!("Opening DOCX file: {}", path);
    
    match DocxParser::from_path(&path) {
        Ok(mut parser) => {
            match parser.parse() {
                Ok(document) => {
                    println!("Successfully parsed DOCX file: {} paragraphs, {} comments, {} images", 
                        document.body.len(), 
                        document.comments.len(), 
                        document.images.len()
                    );
                    Ok(document)
                }
                Err(e) => {
                    let error_msg = format!("Failed to parse DOCX file: {}", e);
                    eprintln!("{}", error_msg);
                    Err(error_msg)
                }
            }
        }
        Err(e) => {
            let error_msg = format!("Failed to open DOCX file: {}", e);
            eprintln!("{}", error_msg);
            Err(error_msg)
        }
    }
}

fn main() {
    tauri::Builder::default()
        .plugin(tauri_plugin_fs::init())
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![open_docx])
        .setup(|app| {
            let window = app.get_webview_window("main").unwrap();
            
            #[cfg(debug_assertions)]
            {
                window.open_devtools();
            }
            
            Ok(())
        })
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}