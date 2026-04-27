// Prevents additional console window on Windows in release, DO NOT REMOVE!!
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use serde::{Deserialize, Serialize};
use tauri::{command, AppHandle};
use tauri_plugin_shell::ShellExt;

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
    medio: Option<String>,
    estado: Option<String>,
    trabajo_realizado_en: Option<String>,
    orden_servicio: Option<String>,
    observacion: Option<String>,
}

async fn run_sidecar(
    app: &AppHandle,
    name: &'static str,
    args: Vec<String>,
) -> Result<Vec<u8>, String> {
    let cmd = app
        .shell()
        .sidecar(name)
        .map_err(|e| format!("Error preparando sidecar {name}: {e}"))?
        .args(args);

    let output = cmd
        .output()
        .await
        .map_err(|e| format!("Error ejecutando {name}: {e}"))?;

    if !output.status.success() {
        let stdout = String::from_utf8_lossy(&output.stdout);
        let stderr = String::from_utf8_lossy(&output.stderr);
        return Err(format!(
            "{name} falló (código {})\nStdout: {stdout}\nStderr: {stderr}",
            output.status.code().unwrap_or(-1)
        ));
    }

    Ok(output.stdout)
}

#[command]
async fn extract_cotizacion(
    app: AppHandle,
    docx_path: String,
) -> Result<DatosCotizacion, String> {
    let stdout = run_sidecar(&app, "extract_cotizacion", vec![docx_path]).await?;
    serde_json::from_slice(&stdout)
        .map_err(|e| format!("Error parseando JSON del extractor: {e}"))
}

#[command]
async fn insert_cotizacion(
    app: AppHandle,
    datos: DatosCotizacion,
    xlsx_path: String,
    sheet_name: Option<String>,
) -> Result<bool, String> {
    let json_str = serde_json::to_string(&datos)
        .map_err(|e| format!("Error serializando datos: {e}"))?;

    let mut args = vec![json_str, xlsx_path];
    if let Some(sheet) = sheet_name {
        args.push(sheet);
    }

    let stdout = run_sidecar(&app, "insert_cotizacion", args).await?;
    let result_json: serde_json::Value = serde_json::from_slice(&stdout)
        .map_err(|e| format!("Error parseando resultado del insertador: {e}"))?;
    Ok(result_json["insertado"].as_bool().unwrap_or(false))
}

#[command]
async fn get_excel_sheets(app: AppHandle, xlsx_path: String) -> Result<Vec<String>, String> {
    let stdout = run_sidecar(&app, "sheets_wrapper", vec![xlsx_path]).await?;
    let result: ExcelSheets = serde_json::from_slice(&stdout)
        .map_err(|e| format!("Error parseando JSON de hojas: {e}"))?;
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
