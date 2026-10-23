use tauri::menu::{Menu, MenuEvent, MenuItem, PredefinedMenuItem, Submenu};
use tauri::{AppHandle, Emitter, EventTarget, Wry};

use crate::windows;

/// 프론트엔드로 전달하는 메뉴 id (= "menu" 이벤트 payload).
const IDS: [&str; 8] = ["new", "open", "save", "close", "undo", "redo", "insert", "delete"];

fn item(app: &AppHandle, id: &str, text: &str, accel: &str) -> tauri::Result<MenuItem<Wry>> {
    MenuItem::with_id(app, id, text, true, Some(accel))
}

pub fn build(app: &AppHandle) -> tauri::Result<Menu<Wry>> {
    let app_menu = Submenu::with_items(
        app,
        "Sheet Editor",
        true,
        &[
            &PredefinedMenuItem::about(app, None, None)?,
            &PredefinedMenuItem::separator(app)?,
            &PredefinedMenuItem::services(app, None)?,
            &PredefinedMenuItem::separator(app)?,
            &PredefinedMenuItem::hide(app, None)?,
            &PredefinedMenuItem::hide_others(app, None)?,
            &PredefinedMenuItem::show_all(app, None)?,
            &PredefinedMenuItem::separator(app)?,
            &PredefinedMenuItem::quit(app, None)?,
        ],
    )?;
    let file = Submenu::with_items(
        app,
        "파일",
        true,
        &[
            &item(app, "new", "새 파일", "CmdOrCtrl+N")?,
            &item(app, "open", "열기…", "CmdOrCtrl+O")?,
            &item(app, "save", "저장", "CmdOrCtrl+S")?,
            &item(app, "close", "닫기", "CmdOrCtrl+W")?,
        ],
    )?;
    let edit = Submenu::with_items(
        app,
        "편집",
        true,
        &[
            &item(app, "undo", "실행 취소", "CmdOrCtrl+Z")?,
            &item(app, "redo", "다시 실행", "CmdOrCtrl+Shift+Z")?,
            &PredefinedMenuItem::separator(app)?,
            &PredefinedMenuItem::cut(app, Some("잘라내기"))?,
            &PredefinedMenuItem::copy(app, Some("복사"))?,
            &PredefinedMenuItem::paste(app, Some("붙여넣기"))?,
            &PredefinedMenuItem::select_all(app, Some("전체 선택"))?,
        ],
    )?;
    let sheet = Submenu::with_items(
        app,
        "시트",
        true,
        &[
            &item(app, "insert", "행/열 삽입", "CmdOrCtrl+Shift+Equal")?,
            &item(app, "delete", "행/열 삭제", "CmdOrCtrl+Minus")?,
        ],
    )?;
    Menu::with_items(app, &[&app_menu, &file, &edit, &sheet])
}

pub fn on_event(app: &AppHandle, event: MenuEvent) {
    let id = event.id().as_ref();
    if !IDS.contains(&id) {
        return;
    }
    if let Some(label) = windows::focused_label(app) {
        let _ = app.emit_to(EventTarget::webview_window(label), "menu", id);
    }
}
