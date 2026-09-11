use enigo::{
    Direction::{Click, Press, Release},
    Enigo, Key, Keyboard, Settings,
};
use serde::Deserialize;
use sqlx::{
    sqlite::{
        SqliteConnectOptions,
        SqliteJournalMode,
        SqliteSynchronous,
    },
    Connection,
    SqliteConnection,
};
use std::time::Duration;

#[cfg(desktop)]
use tauri::{
    menu::{Menu, MenuItem},
    tray::TrayIconBuilder,
    Emitter,
    Manager,
};

#[cfg(target_os = "windows")]
use windows_sys::Win32::{
    Foundation::CloseHandle,
    System::Threading::{
        OpenProcess, QueryFullProcessImageNameW, PROCESS_QUERY_LIMITED_INFORMATION,
    },
    UI::WindowsAndMessaging::{GetForegroundWindow, GetWindowThreadProcessId},
};

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct SqliteTransactionStatement {
    sql: String,
    #[serde(default)]
    params: Vec<String>,
}

#[tauri::command]
fn send_ctrl_c() -> Result<(), String> {
    let mut enigo = Enigo::new(&Settings::default()).map_err(|error| error.to_string())?;

    enigo
        .key(Key::Control, Press)
        .map_err(|error| error.to_string())?;

    let copy_result = enigo.key(Key::Unicode('c'), Click);
    let release_result = enigo.key(Key::Control, Release);

    if let Err(error) = copy_result {
        return Err(error.to_string());
    }

    if let Err(error) = release_result {
        return Err(error.to_string());
    }

    Ok(())
}

#[cfg(target_os = "windows")]
#[tauri::command]
fn get_foreground_process_name() -> Result<String, String> {
    unsafe {
        let hwnd = GetForegroundWindow();

        if hwnd.is_null() {
            return Ok(String::new());
        }

        let mut process_id = 0u32;

        GetWindowThreadProcessId(hwnd, &mut process_id);

        if process_id == 0 {
            return Ok(String::new());
        }

        let process_handle = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, 0, process_id);

        if process_handle.is_null() {
            return Err("Could not open foreground process.".to_string());
        }

        let mut buffer = vec![0u16; 1024];
        let mut size = buffer.len() as u32;

        let success = QueryFullProcessImageNameW(process_handle, 0, buffer.as_mut_ptr(), &mut size);

        CloseHandle(process_handle);

        if success == 0 {
            return Err("Could not read foreground process name.".to_string());
        }

        let path = String::from_utf16_lossy(&buffer[..size as usize]);

        let process_name = std::path::Path::new(&path)
            .file_name()
            .and_then(|name| name.to_str())
            .map(|name| name.to_string())
            .unwrap_or(path);

        Ok(process_name)
    }
}

#[cfg(not(target_os = "windows"))]
#[tauri::command]
fn get_foreground_process_name() -> Result<String, String> {
    Ok(String::new())
}

#[tauri::command]
async fn execute_sqlite_transaction(
    app: tauri::AppHandle,
    statements: Vec<SqliteTransactionStatement>,
    commit: bool,
) -> Result<(), String> {
    let database_path = app
        .path()
        .app_config_dir()
        .map_err(|error| error.to_string())?
        .join("poe2-collector.db");

    if !database_path.exists() {
        return Err("PoE 2 Unique Tracker database could not be found.".to_string());
    }

    let options = SqliteConnectOptions::new()
    .filename(database_path)
    .create_if_missing(false)
    .journal_mode(SqliteJournalMode::Wal)
    .synchronous(SqliteSynchronous::Normal)
    .busy_timeout(Duration::from_secs(15));

    let mut connection = SqliteConnection::connect_with(&options)
        .await
        .map_err(|error| error.to_string())?;

    sqlx::query("BEGIN IMMEDIATE")
        .execute(&mut connection)
        .await
        .map_err(|error| error.to_string())?;

    for statement in statements {
        let mut query = sqlx::query(&statement.sql);

        for value in statement.params {
            query = query.bind(value);
        }

        if let Err(error) = query.execute(&mut connection).await {
            let _ = sqlx::query("ROLLBACK").execute(&mut connection).await;

            return Err(error.to_string());
        }
    }

    if commit {
        sqlx::query("COMMIT")
            .execute(&mut connection)
            .await
            .map_err(|error| error.to_string())?;
    } else {
        sqlx::query("ROLLBACK")
            .execute(&mut connection)
            .await
            .map_err(|error| error.to_string())?;
    }

    Ok(())
}

#[tauri::command]
fn write_tracker_backup(
    directory: String,
    file_name: String,
    bytes: Vec<u8>,
) -> Result<String, String> {
    let directory_path =
        std::path::PathBuf::from(directory);

    if !directory_path.is_absolute() {
        return Err(
            "The backup folder must be an absolute path."
                .to_string(),
        );
    }

    if !directory_path.is_dir() {
        return Err(
            "The selected backup folder is no longer available."
                .to_string(),
        );
    }

    if file_name.is_empty()
        || file_name.contains('/')
        || file_name.contains('\\')
    {
        return Err(
            "The backup filename is invalid."
                .to_string(),
        );
    }

    let target_path =
        directory_path.join(file_name);

    std::fs::write(
        &target_path,
        bytes,
    )
    .map_err(|error| {
        format!(
            "Could not create the local backup: {}",
            error
        )
    })?;

    Ok(
        target_path
            .to_string_lossy()
            .to_string(),
    )
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    let mut builder =
        tauri::Builder::default().plugin(tauri_plugin_updater::Builder::new().build());

    #[cfg(desktop)]
    {
        // Keep PoE 2 Unique Tracker to a single running instance. This avoids
        // duplicate SQLite writers and duplicate global-hotkey registration.
        builder = builder.plugin(tauri_plugin_single_instance::init(|app, _args, _cwd| {
            if let Some(window) = app.get_webview_window("main") {
                let _ = window.show();
                let _ = window.unminimize();
                let _ = window.set_focus();
            }
        }));
    }

    builder
        .setup(|app| {
    #[cfg(desktop)]
    {
        app.handle()
            .plugin(tauri_plugin_global_shortcut::Builder::new().build())?;

        let open_tracker = MenuItem::with_id(
            app,
            "open_tracker",
            "Open tracker",
            true,
            None::<&str>,
        )?;

        let exit_application = MenuItem::with_id(
            app,
            "exit_application",
            "Exit Application",
            true,
            None::<&str>,
        )?;

        let tray_menu = Menu::with_items(
            app,
            &[&open_tracker, &exit_application],
        )?;

        TrayIconBuilder::new()
            .icon(
                app.default_window_icon()
                    .expect("PoE 2 Unique Tracker icon is missing")
                    .clone(),
            )
            .tooltip("PoE 2 Unique Tracker")
            .menu(&tray_menu)
            .show_menu_on_left_click(false)
            .on_menu_event(|app, event| match event.id.as_ref() {
                "open_tracker" => {
                    if let Some(window) =
                        app.get_webview_window("main")
                    {
                        let _ = window.show();
                        let _ = window.unminimize();
                        let _ = window.set_focus();
                    }
                }
                "exit_application" => {
                    app.exit(0);
                }
                _ => {}
            })
            .build(app)?;

        if let Some(window) = app.get_webview_window("main") {
            window.minimize()?;
        }
    }

    Ok(())
})

.on_window_event(|window, event| {
    if window.label() != "main" {
        return;
    }

    if let tauri::WindowEvent::CloseRequested { api, .. } = event {
        api.prevent_close();

        let _ = window.emit(
            "main-close-requested",
            (),
        );
    }
})
.plugin(tauri_plugin_clipboard_manager::init())
        .plugin(tauri_plugin_http::init())
        .plugin(tauri_plugin_fs::init())
        .plugin(tauri_plugin_sql::Builder::new().build())
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_opener::init())
        .invoke_handler(tauri::generate_handler![
    send_ctrl_c,
    get_foreground_process_name,
    execute_sqlite_transaction,
    write_tracker_backup
])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
