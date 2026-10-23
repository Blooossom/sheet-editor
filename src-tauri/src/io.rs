use std::io::Write;
use std::path::{Path, PathBuf};

use encoding_rs::{EncoderResult, EUC_KR};
use percent_encoding::percent_decode_str;
use serde::Serialize;
use tauri::ipc::{InvokeBody, Request, Response};

const UTF8_BOM: [u8; 3] = [0xEF, 0xBB, 0xBF];

pub struct Decoded {
    pub text: String,
    /// "utf-8" | "euc-kr"
    pub encoding: &'static str,
    pub bom: bool,
}

#[derive(Serialize)]
pub struct TextFile {
    pub text: String,
    pub encoding: String,
    pub bom: bool,
}

/// BOM → utf-8+bom, 올바른 UTF-8 → utf-8, 그 외 → euc-kr(CP949 상위 집합).
pub fn decode_text(bytes: &[u8]) -> Decoded {
    if let Some(rest) = bytes.strip_prefix(&UTF8_BOM) {
        if let Ok(s) = std::str::from_utf8(rest) {
            return Decoded { text: s.to_string(), encoding: "utf-8", bom: true };
        }
    } else if let Ok(s) = std::str::from_utf8(bytes) {
        return Decoded { text: s.to_string(), encoding: "utf-8", bom: false };
    }
    // 잘못된 바이트는 U+FFFD가 되며, 이 글자는 EUC-KR로 저장할 수 없어 저장 시 거부된다.
    let (text, _) = EUC_KR.decode_without_bom_handling(bytes);
    Decoded { text: text.into_owned(), encoding: "euc-kr", bom: false }
}

pub fn encode_text(text: &str, encoding: &str, bom: bool) -> Result<Vec<u8>, String> {
    match encoding {
        "utf-8" => {
            let mut out = Vec::with_capacity(text.len() + 3);
            if bom {
                out.extend_from_slice(&UTF8_BOM);
            }
            out.extend_from_slice(text.as_bytes());
            Ok(out)
        }
        "euc-kr" => {
            let mut encoder = EUC_KR.new_encoder();
            let cap = encoder
                .max_buffer_length_from_utf8_without_replacement(text.len())
                .ok_or("파일이 너무 커서 변환할 수 없습니다")?;
            let mut out = vec![0u8; cap];
            let (result, _read, written) =
                encoder.encode_from_utf8_without_replacement(text, &mut out, true);
            match result {
                EncoderResult::InputEmpty => {
                    out.truncate(written);
                    Ok(out)
                }
                EncoderResult::Unmappable(c) => {
                    Err(format!("EUC-KR로 저장할 수 없는 글자가 있습니다: {c}"))
                }
                EncoderResult::OutputFull => Err("인코딩 중 내부 오류가 발생했습니다".to_string()),
            }
        }
        other => Err(format!("지원하지 않는 인코딩입니다: {other}")),
    }
}

/// 같은 폴더의 임시 파일에 쓰고 fsync 후 rename. 기존 파일의 권한은 유지하고, 실패 시 임시 파일을 남기지 않는다.
pub fn write_atomic(path: &Path, bytes: &[u8]) -> Result<(), String> {
    let fail = |e: &dyn std::fmt::Display| format!("파일을 저장할 수 없습니다: {}: {e}", path.display());
    // 심볼릭 링크는 링크 자체가 아니라 가리키는 실제 파일을 덮어쓴다.
    let target: PathBuf = std::fs::canonicalize(path).unwrap_or_else(|_| path.to_path_buf());
    let dir = target
        .parent()
        .filter(|p| !p.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    // 드롭 시 자동 삭제되므로 어느 단계에서 실패해도 임시 파일이 남지 않는다.
    let mut tmp = tempfile::Builder::new()
        .prefix(".sheet-editor-")
        .tempfile_in(dir)
        .map_err(|e| fail(&e))?;
    tmp.write_all(bytes).map_err(|e| fail(&e))?;
    if let Ok(meta) = std::fs::metadata(&target) {
        std::fs::set_permissions(tmp.path(), meta.permissions()).map_err(|e| fail(&e))?;
    }
    tmp.as_file().sync_all().map_err(|e| fail(&e))?;
    tmp.persist(&target).map_err(|e| fail(&e.error))?;
    Ok(())
}

fn read_bytes(path: &str) -> Result<Vec<u8>, String> {
    std::fs::read(path).map_err(|e| format!("파일을 읽을 수 없습니다: {path}: {e}"))
}

#[tauri::command]
pub fn read_file(path: String) -> Result<Response, String> {
    read_bytes(&path).map(Response::new)
}

#[tauri::command]
pub fn write_file(request: Request) -> Result<(), String> {
    let InvokeBody::Raw(bytes) = request.body() else {
        return Err("저장할 데이터가 올바르지 않습니다".to_string());
    };
    let header = request
        .headers()
        .get("x-path")
        .ok_or("저장할 경로가 없습니다")?
        .to_str()
        .map_err(|_| "저장할 경로가 올바르지 않습니다".to_string())?;
    let path = percent_decode_str(header)
        .decode_utf8()
        .map_err(|_| "저장할 경로가 올바르지 않습니다".to_string())?;
    write_atomic(Path::new(path.as_ref()), bytes)
}

#[tauri::command]
pub fn read_text(path: String) -> Result<TextFile, String> {
    let d = decode_text(&read_bytes(&path)?);
    Ok(TextFile { text: d.text, encoding: d.encoding.to_string(), bom: d.bom })
}

#[tauri::command]
pub fn write_text(path: String, text: String, encoding: String, bom: bool) -> Result<(), String> {
    // 인코딩에 실패하면 파일을 건드리지 않는다.
    let bytes = encode_text(&text, &encoding, bom)?;
    write_atomic(Path::new(&path), &bytes)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test] fn utf8_and_bom() {
        let d = decode_text("가,b".as_bytes()); assert_eq!((d.encoding, d.bom, d.text.as_str()), ("utf-8", false, "가,b"));
        let d = decode_text(b"\xEF\xBB\xBFa"); assert_eq!((d.encoding, d.bom, d.text.as_str()), ("utf-8", true, "a"));
    }
    #[test] fn euc_kr_roundtrip() {
        let bytes = [0xB0, 0xA1, 0xB3, 0xAA]; // "가나"
        let d = decode_text(&bytes); assert_eq!((d.encoding, d.text.as_str()), ("euc-kr", "가나"));
        assert_eq!(encode_text("가나", "euc-kr", false).unwrap(), bytes.to_vec());
    }
    #[test] fn euc_kr_unmappable_is_error() {
        let e = encode_text("가😀", "euc-kr", false).unwrap_err(); assert!(e.contains("EUC-KR") && e.contains("😀"));
    }
    #[test] fn bom_written() { assert_eq!(encode_text("a", "utf-8", true).unwrap(), b"\xEF\xBB\xBFa".to_vec()); }
    #[test] fn atomic_write_korean_path() {
        let dir = tempfile::tempdir().unwrap(); let p = dir.path().join("한글 폴더 파일.csv");
        std::fs::write(&p, b"old").unwrap(); write_atomic(&p, b"new").unwrap();
        assert_eq!(std::fs::read(&p).unwrap(), b"new");
    }
    #[cfg(unix)]
    #[test] fn atomic_write_keeps_permissions() {
        use std::os::unix::fs::PermissionsExt;
        let dir = tempfile::tempdir().unwrap(); let p = dir.path().join("a.csv");
        std::fs::write(&p, b"old").unwrap();
        std::fs::set_permissions(&p, std::fs::Permissions::from_mode(0o640)).unwrap();
        write_atomic(&p, b"new").unwrap();
        assert_eq!(std::fs::metadata(&p).unwrap().permissions().mode() & 0o777, 0o640);
    }
    #[test] fn atomic_write_failure_leaves_no_temp() {
        let dir = tempfile::tempdir().unwrap();
        let p = dir.path().join("sub"); std::fs::create_dir(&p).unwrap(); // 디렉터리 위로는 rename 불가
        assert!(write_atomic(&p, b"x").is_err());
        assert_eq!(std::fs::read_dir(dir.path()).unwrap().count(), 1);
    }
    #[test] fn atomic_write_creates_new_and_leaves_no_temp() {
        let dir = tempfile::tempdir().unwrap(); let p = dir.path().join("n.csv");
        write_atomic(&p, b"x").unwrap();
        assert_eq!(std::fs::read_dir(dir.path()).unwrap().count(), 1);
    }
    #[test] fn empty_is_utf8_without_bom() {
        let d = decode_text(b""); assert_eq!((d.encoding, d.bom, d.text.as_str()), ("utf-8", false, ""));
    }
    #[test] fn unknown_encoding_rejected() { assert!(encode_text("a", "latin1", false).is_err()); }
}
