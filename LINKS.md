# 아침 브리핑 — 링크 모음 · 작업 지시 템플릿

> 사용자에게 할 일을 안내할 때는 **아래 링크를 하나도 빼지 말고** 단계마다 같이 준다.

## 링크

| 용도 | 링크 |
|---|---|
| 📱 앱 (폰·PC) | https://mhjang1998-code.github.io/briefing/ |
| 📊 구글 시트 「아침 브리핑」 | https://docs.google.com/spreadsheets/d/1rBpt0chrirrIrYuaREaA-3Omi9Go5wlayPWRe_ejI98/edit |
| ⚙️ Apps Script 편집기 | 시트 메뉴 **확장 프로그램 → Apps Script** · 또는 https://script.google.com/home → 「아침 브리핑」 프로젝트 (편집기 직접 링크는 서버 v11 배포 뒤 `uiLoad.scriptId`로 `https://script.google.com/d/<scriptId>/edit`) |
| 🚀 배포 관리 | Apps Script 편집기 오른쪽 위 **배포 → 배포 관리** |
| 📄 서버 코드 원본 (복붙용) | https://raw.githubusercontent.com/mhjang1998-code/briefing/main/server/Code.gs |
| ⚡ 루틴 (일일 브리핑) | https://claude.ai/code/routines → 「일일 브리핑」 (데스크톱 앱 창이 비면 크롬 주소창에 직접) |
| 🗂️ 앱 저장소 | https://github.com/mhjang1998-code/briefing |

## 서버 재배포 지시 템플릿 (그대로 복사해서 안내)

1. 서버 코드 열기 👉 https://raw.githubusercontent.com/mhjang1998-code/briefing/main/server/Code.gs → 페이지 클릭 → **Ctrl+A → Ctrl+C**
2. 구글 시트 👉 https://docs.google.com/spreadsheets/d/1rBpt0chrirrIrYuaREaA-3Omi9Go5wlayPWRe_ejI98/edit → **확장 프로그램 → Apps Script** (또는 https://script.google.com/home → 「아침 브리핑」)
3. `Code.gs` 안 클릭 → **Ctrl+A → Ctrl+V** → 💾 저장
4. (새 권한이 필요한 버전만) 위쪽 함수 칸 `onOpen ▼` → **authorize** → ▷ 실행 → 권한 검토 → 고급 → 이동 → 허용
5. **배포 → 배포 관리** → 연필 ✏️ → 버전: **새 버전** → 배포 (⚠️ 「새 배포」 금지 — 주소가 바뀜)
6. 앱 👉 https://mhjang1998-code.github.io/briefing/ → **Ctrl+Shift+R** (폰은 앱 닫았다 다시 열기)

## 새 기기 연결 템플릿

1. 앱 👉 https://mhjang1998-code.github.io/briefing/ 열기 → 「처음 연결」 화면
2. **앱 비밀번호** 입력 → 연결 (비밀번호를 잊었으면 시트 👉 https://docs.google.com/spreadsheets/d/1rBpt0chrirrIrYuaREaA-3Omi9Go5wlayPWRe_ejI98/edit → 메뉴 **브리핑 앱 → 앱 비밀번호 정하기**)

## ⚡ 루틴 토큰 다시 만들기 템플릿

1. 👉 https://claude.ai/code/routines → 「일일 브리핑」 → ⋯ → Edit → 「API로 호출」 → **토큰 생성**(한 번만 보임) → 저장
2. 앱 👉 https://mhjang1998-code.github.io/briefing/ → ⚙️ → 「⚡ 지금 브리핑 연결」에 붙여넣기 → 저장 (채팅에 붙여넣지 않기)
