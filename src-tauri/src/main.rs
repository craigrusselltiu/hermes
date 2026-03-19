// Prevents additional console window on Windows in release, DO NOT REMOVE!!
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use hermes_app::model::Document;
use hermes_app::parser::DocxParser;
use serde::{Deserialize, Serialize};
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::Mutex;
use tauri::Manager;

const MAX_RECENT_FILES: usize = 8;
const SETTINGS_FILE_NAME: &str = "settings.json";
const RECENT_FILES_FILE_NAME: &str = "recent-files.json";

#[derive(Debug, Default)]
struct LaunchState {
    docx_path: Mutex<Option<String>>,
}

#[derive(Debug, Default, Serialize, Deserialize)]
struct AppSettings {
    theme: Option<String>,
}

#[derive(Debug, Default, Serialize, Deserialize)]
struct RecentFiles {
    paths: Vec<String>,
}

#[tauri::command]
fn open_docx(path: String, app: tauri::AppHandle) -> Result<Document, String> {
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
                    if let Err(err) = remember_recent_file(&app, &path) {
                        eprintln!("Failed to store recent file '{}': {}", path, err);
                    }
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

#[tauri::command]
fn get_recent_files(app: tauri::AppHandle) -> Result<Vec<String>, String> {
    Ok(load_recent_files(&app)?)
}

#[tauri::command]
fn get_launch_docx_path(state: tauri::State<LaunchState>) -> Result<Option<String>, String> {
    let mut path = state
        .docx_path
        .lock()
        .map_err(|_| "Could not access launch state".to_string())?;
    Ok(path.take())
}

#[tauri::command]
fn get_theme_preference(app: tauri::AppHandle) -> Result<Option<String>, String> {
    Ok(load_settings(&app)?.theme)
}

#[tauri::command]
fn set_theme_preference(theme: String, app: tauri::AppHandle) -> Result<(), String> {
    let normalized_theme = normalize_theme(&theme)
        .ok_or_else(|| format!("Unsupported theme '{}'", theme))?
        .to_string();

    let mut settings = load_settings(&app)?;
    settings.theme = Some(normalized_theme);
    save_settings(&app, &settings)
}

#[tauri::command]
fn show_main_window(app: tauri::AppHandle) {
    if let Some(window) = app.get_webview_window("main") {
        let _ = window.show();
    }
}

#[tauri::command]
fn quit_app(app: tauri::AppHandle) {
    app.exit(0);
}

fn app_data_file_path(app: &tauri::AppHandle, filename: &str) -> Result<PathBuf, String> {
    let app_data_dir = app
        .path()
        .app_data_dir()
        .map_err(|e| format!("Could not resolve app data directory: {}", e))?;

    fs::create_dir_all(&app_data_dir)
        .map_err(|e| format!("Could not create app data directory: {}", e))?;

    Ok(app_data_dir.join(filename))
}

fn recent_files_path(app: &tauri::AppHandle) -> Result<PathBuf, String> {
    app_data_file_path(app, RECENT_FILES_FILE_NAME)
}

fn settings_path(app: &tauri::AppHandle) -> Result<PathBuf, String> {
    app_data_file_path(app, SETTINGS_FILE_NAME)
}

fn load_settings(app: &tauri::AppHandle) -> Result<AppSettings, String> {
    let path = settings_path(app)?;
    if !path.exists() {
        return Ok(AppSettings::default());
    }

    let content =
        fs::read_to_string(&path).map_err(|e| format!("Could not read settings: {}", e))?;

    serde_json::from_str(&content).map_err(|e| format!("Could not parse settings: {}", e))
}

fn save_settings(app: &tauri::AppHandle, settings: &AppSettings) -> Result<(), String> {
    let path = settings_path(app)?;
    let json = serde_json::to_string_pretty(settings)
        .map_err(|e| format!("Could not serialize settings: {}", e))?;

    fs::write(path, json).map_err(|e| format!("Could not write settings: {}", e))
}

fn load_recent_files(app: &tauri::AppHandle) -> Result<Vec<String>, String> {
    let path = recent_files_path(app)?;
    if !path.exists() {
        return Ok(Vec::new());
    }

    let content = fs::read_to_string(&path)
        .map_err(|e| format!("Could not read recent files: {}", e))?;

    let mut recent_files: RecentFiles = serde_json::from_str(&content)
        .map_err(|e| format!("Could not parse recent files: {}", e))?;
    recent_files.paths.retain(|entry| !entry.trim().is_empty());
    Ok(recent_files.paths)
}

fn save_recent_files(app: &tauri::AppHandle, paths: &[String]) -> Result<(), String> {
    let path = recent_files_path(app)?;
    let payload = RecentFiles {
        paths: paths.to_vec(),
    };
    let json = serde_json::to_string_pretty(&payload)
        .map_err(|e| format!("Could not serialize recent files: {}", e))?;

    fs::write(path, json).map_err(|e| format!("Could not write recent files: {}", e))
}

fn remember_recent_file(app: &tauri::AppHandle, path: &str) -> Result<(), String> {
    let mut recent_files = load_recent_files(app).unwrap_or_default();
    recent_files.retain(|existing| existing != path);
    recent_files.insert(0, path.to_string());
    recent_files.truncate(MAX_RECENT_FILES);
    save_recent_files(app, &recent_files)
}

fn normalize_theme(theme: &str) -> Option<&'static str> {
    match theme {
        "light" => Some("light"),
        "dark" => Some("dark"),
        _ => None,
    }
}

fn find_launch_docx_path() -> Option<String> {
    std::env::args_os()
        .skip(1)
        .find_map(|arg| docx_path_from_arg(PathBuf::from(arg)))
}

fn docx_path_from_arg(path: PathBuf) -> Option<String> {
    if !is_docx_path(&path) {
        return None;
    }

    let absolute_path = if path.is_absolute() {
        path
    } else {
        std::env::current_dir().ok()?.join(path)
    };

    if !absolute_path.exists() {
        return None;
    }

    Some(absolute_path.to_string_lossy().into_owned())
}

fn is_docx_path(path: &Path) -> bool {
    path.extension()
        .and_then(|ext| ext.to_str())
        .map(|ext| ext.eq_ignore_ascii_case("docx"))
        .unwrap_or(false)
}

fn main() {
    let launch_state = LaunchState {
        docx_path: Mutex::new(find_launch_docx_path()),
    };

    tauri::Builder::default()
        .manage(launch_state)
        .plugin(tauri_plugin_fs::init())
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![
            open_docx,
            get_recent_files,
            get_launch_docx_path,
            get_theme_preference,
            set_theme_preference,
            show_main_window,
            quit_app
        ])
        .setup(|_app| {
            #[cfg(debug_assertions)]
            {
                if let Some(window) = _app.get_webview_window("main") {
                    window.open_devtools();
                }
            }
            
            Ok(())
        })
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
