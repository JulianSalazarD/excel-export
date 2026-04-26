// Prevents additional console window on Windows in release, DO NOT REMOVE!!
#
![cfg_attr(not(debug_assertions)
, windows_subsystem = "windows")]

use tauri::command;
use std::process::Command;
use std::path::PathBuf;
use serde::{Deserialize, Serialize};

// Determinar la ruta al ejecutable Python
fn get_backend_path() -> PathBuf {
    // CARGO_MANIFEST_DIR apunta a src-tauri
    // Necesitamos subir un nivel para llegar a la raíz del proyecto
    let mut path = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    path.pop(); // Salir de src-tauri, llegar a la raíz del proyecto
    path.push("backend");
    path
}

#[derive(Serialize, Deserialize)]
struct ExcelSheets {
    sheets: Vec<String>,
}

#[derive(Serialize, Deserialize)]
struct DatosCotizacion {
    numero: Option<String>,
    nombre: Option<String>,
    empresa: Option<String>,
    telefono: Option<String>,
    correo: Option<String>,
    servicio: Option<String>,
    valor_total: Option<String>,
    estado: Option<String>,
    trabajo_realizado_en: Option<String>,
    orden_servicio: Option<String>,
    fecha: Option<String>,
    observacion: Option<String>,
}

#[command]
fn extract_cotizacion(docx_path: String) -> Result<DatosCotizacion, String> {
    let backend_path = get_backend_path();
    let extract_exe = backend_path.join("extract_cotizacion");
    
    let output = Command::new(extract_exe)
        .arg(&docx_path)
        .output()
        .map_err(|e| format!("Error ejecutando extractor: {}", e))?;
    
    if !output.status.success() {
        let stderr = String::from_utf8_lossy(&output.stderr);
        return Err(format!("Extractor falló: {}", stderr));
    }
    
    let json_str = String::from_utf8(output.stdout)
        .map_err(|e| format!("Error leyendo salida: {}", e))?;
    
    let datos: DatosCotizacion = serde_json::from_str(&json_str)
        .map_err(|e| format!("Error parseando JSON: {}", e))?;
    
    Ok(datos)
}

#[command]
fn insert_cotizacion(datos: DatosCotizacion, xlsx_path: String, sheet_name: Option<String>) -> Result<bool, String> {
    let json_str = serde_json::to_string(&datos)
        .map_err(|e| format!("Error serializando datos: {}", e))?;

    let backend_path = get_backend_path();
    let insert_exe = backend_path.join("insert_cotizacion");

    let mut cmd = Command::new(insert_exe);
    cmd.arg(&json_str);
    cmd.arg(&xlsx_path);

    if let Some(sheet) = sheet_name {
        cmd.arg(sheet);
    }
    
    let output = cmd.output()
        .map_err(|e| format!("Error ejecutando insertador: {}", e))?;

    // Leer stdout independientemente del código de salida
    let stdout_str = String::from_utf8_lossy(&output.stdout);
    
    // Si el proceso falló con un código diferente a 1 (que indica duplicado), retornar error
    if !output.status.success() && output.status.code() != Some(1) {
        let stderr = String::from_utf8_lossy(&output.stderr);
        return Err(format!("Insertador falló (código {})\nStdout: {}\nStderr: {}", 
            output.status.code().unwrap_or(-1), stdout_str, stderr));
    }

    // Parsear el resultado JSON del stdout
    let result_json: serde_json::Value = serde_json::from_str(&stdout_str)
        .map_err(|e| format!("Error parseando resultado JSON: {}\nSalida: {}", e, stdout_str))?;

    Ok(result_json["insertado"].as_bool().unwrap_or(false))
}

#[command]
fn get_excel_sheets(xlsx_path: String) -> Result<Vec<String>, String> {
    let backend_path = get_backend_path();
    let sheets_exe = backend_path.join("sheets_wrapper");

    let output = Command::new(sheets_exe)
        .arg(&xlsx_path)
        .output()
        .map_err(|e| format!("Error ejecutando sheets_wrapper: {}", e))?;

    if !output.status.success() {
        let stdout = String::from_utf8_lossy(&output.stdout);
        let stderr = String::from_utf8_lossy(&output.stderr);
        return Err(format!("sheets_wrapper falló (código {})\nStdout: {}\nStderr: {}", 
            output.status.code().unwrap_or(-1), stdout, stderr));
    }

    let json_str = String::from_utf8(output.stdout)
        .map_err(|e| format!("Error leyendo salida: {}", e))?;

    let result: ExcelSheets = serde_json::from_str(&json_str)
        .map_err(|e| format!("Error parseando JSON: {}", e))?;

    Ok(result.sheets)
}

fn main() {
    tauri::Builder::default()
        .plugin(tauri_plugin_shell::init())
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![
            extract_cotizacion,
            insert_cotizacion,
            get_excel_sheets
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
