use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::sync::Mutex;
use std::sync::atomic::{AtomicUsize, Ordering};

use percent_encoding::{NON_ALPHANUMERIC, utf8_percent_encode};
use tauri::{AppHandle, Manager, Url, WebviewUrl, WebviewWindow, WebviewWindowBuilder, WindowEvent};

const UNTITLED: &str = "제목 없음";
const LABEL_PREFIX: &str = "doc-";

struct DocInfo {
    /// canonicalize된 경로. 새 문서는 None.
    path: Option<PathBuf>,
    dirty: bool,
    /// `?new=1`로 만들어진 창인지(첫 실행 빈 창 교체 판단용).
    created_blank: bool,
}

#[derive(Default)]
pub struct Docs {
    counter: AtomicUsize,
    map: Mutex<HashMap<String, DocInfo>>,
}

fn label_number(label: &str) -> Option<usize> {
    label.strip_prefix(LABEL_PREFIX)?.parse().ok()
}

fn canonical(path: &str) -> Result<PathBuf, String> {
    std::fs::canonicalize(path).map_err(|e| format!("파일을 찾을 수 없습니다: {path}: {e}"))
}

fn title_for(path: Option<&Path>) -> String {
    path.and_then(|p| p.file_name())
        .map(|n| n.to_string_lossy().into_owned())
        .unwrap_or_else(|| UNTITLED.to_string())
}

/// Finder가 넘기는 `file://` URL(퍼센트 인코딩 포함)을 경로로 바꾼다.
pub fn file_url_to_path(url: &Url) -> Option<PathBuf> {
    if url.scheme() != "file" {
        return None;
    }
    url.to_file_path().ok()
}

fn find_open(map: &HashMap<String, DocInfo>, path: &Path) -> Option<String> {
    map.iter()
        .find(|(_, d)| d.path.as_deref() == Some(path))
        .map(|(label, _)| label.clone())
}

/// 손대지 않은 빈 창. 포커스된 창을 우선하고, 아니면 번호가 가장 작은 창.
fn find_pristine_blank(map: &HashMap<String, DocInfo>, focused: Option<&str>) -> Option<String> {
    let mut blanks: Vec<(usize, &String)> = map
        .iter()
        .filter(|(_, d)| d.path.is_none() && !d.dirty && d.created_blank)
        .map(|(l, _)| (label_number(l).unwrap_or(usize::MAX), l))
        .collect();
    blanks.sort();
    blanks
        .iter()
        .find(|(_, l)| Some(l.as_str()) == focused)
        .or(blanks.first())
        .map(|(_, l)| (*l).clone())
}

fn doc_url(path: Option<&Path>) -> String {
    match path {
        Some(p) => format!(
            "index.html?path={}",
            utf8_percent_encode(&p.to_string_lossy(), NON_ALPHANUMERIC)
        ),
        None => "index.html?new=1".to_string(),
    }
}

fn latest_label(map: &HashMap<String, DocInfo>) -> Option<String> {
    map.keys().filter_map(|l| label_number(l).map(|n| (n, l))).max().map(|(_, l)| l.clone())
}

/// 포커스된 문서 창의 라벨. 없으면 가장 최근에 만든 문서 창.
pub fn focused_label(app: &AppHandle) -> Option<String> {
    let focused = app
        .webview_windows()
        .into_iter()
        .find(|(label, w)| label.starts_with(LABEL_PREFIX) && matches!(w.is_focused(), Ok(true)))
        .map(|(label, _)| label);
    focused.or_else(|| latest_label(&app.state::<Docs>().map.lock().unwrap()))
}

fn create_window(app: &AppHandle, path: Option<PathBuf>) -> Result<(), String> {
    let docs = app.state::<Docs>();
    let n = docs.counter.fetch_add(1, Ordering::SeqCst) + 1;
    let label = format!("{LABEL_PREFIX}{n}");
    let url = doc_url(path.as_deref());
    let title = title_for(path.as_deref());
    // 빌드 중 발생하는 창 이벤트가 잠금을 기다리지 않도록 등록 후 잠금을 풀고 만든다.
    docs.map.lock().unwrap().insert(
        label.clone(),
        DocInfo { created_blank: path.is_none(), path, dirty: false },
    );
    let built = WebviewWindowBuilder::new(app, &label, WebviewUrl::App(url.into()))
        .title(title)
        .inner_size(1200.0, 800.0)
        .build();
    if let Err(e) = built {
        docs.map.lock().unwrap().remove(&label);
        return Err(format!("창을 열 수 없습니다: {e}"));
    }
    Ok(())
}

pub fn open(app: &AppHandle, path: Option<String>) -> Result<(), String> {
    let Some(raw) = path else {
        return create_window(app, None);
    };
    let path = canonical(&raw)?;
    let docs = app.state::<Docs>();
    let focused = focused_label(app);
    let (existing, blank) = {
        let map = docs.map.lock().unwrap();
        (find_open(&map, &path), find_pristine_blank(&map, focused.as_deref()))
    };
    if let Some(label) = existing {
        if let Some(w) = app.get_webview_window(&label) {
            let _ = w.set_focus();
        }
        return Ok(());
    }
    create_window(app, Some(path))?;
    if let Some(label) = blank {
        if let Some(w) = app.get_webview_window(&label) {
            // 손대지 않은 빈 창이므로 확인 없이 닫는다.
            let _ = w.destroy();
        }
    }
    Ok(())
}

#[tauri::command]
pub async fn open_document(app: AppHandle, path: Option<String>) -> Result<(), String> {
    open(&app, path)
}

#[tauri::command]
pub async fn set_document_path(window: WebviewWindow, path: String) -> Result<(), String> {
    // 방금 저장한 파일이므로 보통 존재하지만, 실패해도 입력 경로를 그대로 쓴다.
    let canon = std::fs::canonicalize(&path).unwrap_or_else(|_| PathBuf::from(&path));
    let title = title_for(Some(&canon));
    if let Some(info) = window.app_handle().state::<Docs>().map.lock().unwrap().get_mut(window.label()) {
        info.path = Some(canon);
    }
    window.set_title(&title).map_err(|e| e.to_string())
}

#[tauri::command]
pub fn set_dirty(window: WebviewWindow, dirty: bool) {
    if let Some(info) = window.app_handle().state::<Docs>().map.lock().unwrap().get_mut(window.label()) {
        info.dirty = dirty;
    }
}

pub fn on_window_event(window: &tauri::Window, event: &WindowEvent) {
    if let WindowEvent::Destroyed = event {
        window.app_handle().state::<Docs>().map.lock().unwrap().remove(window.label());
    }
}

/// dirty인 문서가 있으면 true. 이때 모든 창에 close()를 걸어 각 창이 자체 확인을 띄우게 한다.
pub fn begin_quit_if_dirty(app: &AppHandle) -> bool {
    // 이미 닫힌 창의 항목이 Destroyed 처리 전에 남아 있어도 무시한다.
    let open = app.webview_windows();
    let any_dirty = app
        .state::<Docs>()
        .map
        .lock()
        .unwrap()
        .iter()
        .any(|(label, d)| d.dirty && open.contains_key(label));
    if any_dirty {
        for (_, w) in app.webview_windows() {
            let _ = w.close();
        }
    }
    any_dirty
}

#[cfg(test)]
mod tests {
    use super::*;

    fn info(path: Option<&str>, dirty: bool, blank: bool) -> DocInfo {
        DocInfo { path: path.map(PathBuf::from), dirty, created_blank: blank }
    }

    #[test]
    fn url_decodes_korean_and_space() {
        let u = Url::parse("file:///Users/me/%ED%95%9C%EA%B8%80%20%ED%8F%B4%EB%8D%94/%EB%AC%B8%EC%84%9C.xlsx").unwrap();
        assert_eq!(file_url_to_path(&u).unwrap(), PathBuf::from("/Users/me/한글 폴더/문서.xlsx"));
    }

    #[test]
    fn non_file_url_rejected() {
        assert!(file_url_to_path(&Url::parse("https://a.com/x.csv").unwrap()).is_none());
    }

    #[test]
    fn canonical_matches_symlink_and_dot_segments() {
        let dir = tempfile::tempdir().unwrap();
        let real = dir.path().join("a.csv");
        std::fs::write(&real, "x").unwrap();
        let link = dir.path().join("link.csv");
        #[cfg(unix)]
        std::os::unix::fs::symlink(&real, &link).unwrap();
        let dotted = dir.path().join("sub/../a.csv");
        std::fs::create_dir(dir.path().join("sub")).unwrap();
        let c = canonical(real.to_str().unwrap()).unwrap();
        assert_eq!(canonical(dotted.to_str().unwrap()).unwrap(), c);
        #[cfg(unix)]
        assert_eq!(canonical(link.to_str().unwrap()).unwrap(), c);
        let mut map = HashMap::new();
        map.insert("doc-1".to_string(), DocInfo { path: Some(c.clone()), dirty: false, created_blank: false });
        assert_eq!(find_open(&map, &c).as_deref(), Some("doc-1"));
        assert!(canonical(dir.path().join("nope.csv").to_str().unwrap()).is_err());
    }

    #[test]
    fn pristine_blank_rules() {
        let mut map = HashMap::new();
        map.insert("doc-1".to_string(), info(None, true, true));
        map.insert("doc-2".to_string(), info(Some("/a.csv"), false, false));
        assert!(find_pristine_blank(&map, None).is_none());
        map.get_mut("doc-1").unwrap().dirty = false;
        assert_eq!(find_pristine_blank(&map, None).as_deref(), Some("doc-1"));
        // ?path=로 만들었다가 path가 없어질 수는 없지만, 새 문서가 아닌 창은 제외된다.
        map.insert("doc-3".to_string(), info(None, false, false));
        map.remove("doc-1");
        assert!(find_pristine_blank(&map, None).is_none());
    }

    #[test]
    fn pristine_blank_prefers_focused_then_lowest() {
        let mut map = HashMap::new();
        for l in ["doc-10", "doc-2", "doc-3"] {
            map.insert(l.to_string(), info(None, false, true));
        }
        assert_eq!(find_pristine_blank(&map, None).as_deref(), Some("doc-2"));
        assert_eq!(find_pristine_blank(&map, Some("doc-10")).as_deref(), Some("doc-10"));
        assert_eq!(find_pristine_blank(&map, Some("doc-99")).as_deref(), Some("doc-2"));
    }

    #[test]
    fn doc_url_roundtrips_special_paths() {
        let base = Url::parse("tauri://localhost/").unwrap();
        for p in ["/a b/c+d&e#f%g.csv", "/Users/me/한글 폴더/문서 (1).xlsx", "/x/100%/a=b?c.csv"] {
            let url = base.join(&doc_url(Some(Path::new(p)))).unwrap();
            let got: Vec<_> = url.query_pairs().collect();
            assert_eq!(got.len(), 1);
            assert_eq!((got[0].0.as_ref(), got[0].1.as_ref()), ("path", p));
        }
        let url = base.join(&doc_url(None)).unwrap();
        assert_eq!(url.query(), Some("new=1"));
    }

    #[test]
    fn latest_label_uses_numeric_order() {
        let mut map = HashMap::new();
        for l in ["doc-2", "doc-10", "doc-9"] {
            map.insert(l.to_string(), info(None, false, true));
        }
        assert_eq!(latest_label(&map).as_deref(), Some("doc-10"));
    }

    #[test]
    fn titles() {
        assert_eq!(title_for(Some(Path::new("/a/b/한글 문서.xlsx"))), "한글 문서.xlsx");
        assert_eq!(title_for(None), "제목 없음");
    }
}
