# Sheet Editor

xlsx와 csv 파일을 바로 열어 고치고 원본에 덮어써 저장하는 가벼운 맥 앱 (Tauri 2 + React + TypeScript).

```bash
npm test                                         # vitest
cargo test --manifest-path src-tauri/Cargo.toml  # Rust 테스트
npm run tauri dev                                # 개발 모드 실행
```

## 빌드

```bash
npm install
npm run tauri build   # 결과: src-tauri/target/release/bundle/macos/Sheet Editor.app (dmg도 함께 생성)
```

## 설치

1. `Sheet Editor.app`을 `/Applications`로 복사합니다.
2. 서명하지 않은 앱이라 처음 한 번은 앱을 우클릭 → **열기**로 실행합니다 (Gatekeeper 경고에서 열기 선택).
3. 기본 앱 지정: Finder에서 xlsx(또는 csv) 파일 선택 → **정보 가져오기**(⌘I) → **다음으로 열기**에서 Sheet Editor 선택 → **모두 변경**.

## 알려진 제한

- Dock 종료·로그아웃은 저장 확인을 거치지 않습니다.
- 그림·차트·메모·피벗·확장 서식이 있는 시트는 행·열을 추가하거나 삭제할 수 없습니다.
- 표(table)의 머리글 행과 전체 데이터 행은 삭제할 수 없습니다.
- 날짜 외 숫자 서식(통화, 퍼센트 등)은 서식 없이 원래 값으로 표시됩니다.
