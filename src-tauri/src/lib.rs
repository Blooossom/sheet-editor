mod io;
mod menu;
mod windows;

use tauri::RunEvent;
use tauri_plugin_dialog::{DialogExt, MessageDialogKind};

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    let app = tauri::Builder::default()
        .plugin(tauri_plugin_opener::init())
        .plugin(tauri_plugin_dialog::init())
        .manage(windows::Docs::default())
        .invoke_handler(tauri::generate_handler![
            io::read_file,
            io::write_file,
            io::read_text,
            io::write_text,
            windows::open_document,
            windows::set_document_path,
            windows::clear_document_path,
            windows::set_dirty
        ])
        .on_window_event(|window, event| windows::on_window_event(window, event))
        .setup(|app| {
            let handle = app.handle();
            app.set_menu(menu::build(handle)?)?;
            app.on_menu_event(menu::on_event);
            windows::open(handle, None)?;
            Ok(())
        })
        .build(tauri::generate_context!())
        .expect("error while building tauri application");

    app.run(|app, event| match event {
        // Finder 더블클릭 등으로 들어온 파일
        RunEvent::Opened { urls } => {
            for url in urls {
                if let Some(path) = windows::file_url_to_path(&url) {
                    let shown = path.to_string_lossy().into_owned();
                    if let Err(e) = windows::open(app, Some(shown.clone())) {
                        eprintln!("{e}");
                        app.dialog()
                            .message(format!("파일을 열 수 없습니다: {shown}"))
                            .kind(MessageDialogKind::Error)
                            .show(|_| {});
                    }
                }
            }
        }
        RunEvent::ExitRequested { api, .. } => {
            if windows::begin_quit_if_dirty(app) {
                api.prevent_exit();
            }
        }
        _ => {}
    });
}
