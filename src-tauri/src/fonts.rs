/// 설치된 글꼴 계열 이름(화면 표시용으로 정렬됨). 숨김 계열(점으로 시작)은 뺀다.
#[tauri::command]
pub fn list_fonts() -> Vec<String> {
    family_names().into_iter().filter(|n| !n.starts_with('.')).collect()
}

#[cfg(target_os = "macos")]
fn family_names() -> Vec<String> {
    use objc2_core_foundation::{CFArray, CFRetained, CFString};
    use objc2_core_text::CTFontManagerCopyAvailableFontFamilyNames;

    // SAFETY: 인자가 없는 CoreText 함수이고, 문서상 CFString 배열을 돌려준다.
    let names = unsafe {
        CFRetained::cast_unchecked::<CFArray<CFString>>(CTFontManagerCopyAvailableFontFamilyNames())
    };
    names.iter().map(|s| s.to_string()).collect()
}

#[cfg(not(target_os = "macos"))]
fn family_names() -> Vec<String> {
    Vec::new()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[cfg(target_os = "macos")]
    #[test]
    fn lists_system_fonts_without_hidden_ones() {
        let fonts = list_fonts();
        assert!(fonts.iter().any(|f| f == "Helvetica"));
        assert!(fonts.iter().all(|f| !f.starts_with('.')));
    }
}
