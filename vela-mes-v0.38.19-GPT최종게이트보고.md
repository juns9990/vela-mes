# VELA MES v0.38.19 — 최종 Gate 증거 보고

- 버전: **v0.38.19** = v0.38.18 + 안전 잠금 1곳 추가(생산 LOT) + 테스트 안정화 · 기능 변경 없음
- 배포: index-dev.html = v0.38.19 · **index.html = v0.38.15 유지** (사용자 최종 승인 전 교체 없음)
- **검증 대상 해시(SHA-256)**: `index-dev.html` = `vela-mes-prototype.html` = `5fe1d475148b2e593302bc7048bafe9f6c6f59951dc27762fad8927e98f5f99e`
- 제출(GitHub 업로드 대상): `index-dev.html` · `_verify-otd.js` · `_verify-sec.js` · `_verify-hold3.js` · `vela-mes-v0.38.19-검증로그.txt`
- 재현: `npm i jsdom` 후 `VELA_FILE=index-dev.html node _verify-otd.js` / `_verify-sec.js` (나머지 하네스는 `vela-mes-prototype.html` 이름으로 복사해 실행)

## 1. GPT 남은 조건 4건

| # | 조건 | 결과 |
|---|---|---|
| 1 | v0.38.19 소스·테스트 원본 제출 + GitHub 일치 | 위 5개 파일 제출 · GitHub 업로드 후 해시로 대조 가능 (업로드는 사용자) |
| 2 | 실제 HTML 대상 재실행 + 해시 일치 | otd **27/27** · sec **24/24** 를 `index-dev.html` 에 직접 실행 · 해시 위 기재 |
| 3 | Firestore 개인 단락 설정 저장·재접속·사용자 변경 | **NOT VERIFIED** — 원칙상 실서버 로그인 검증 안 함 · 사용자 수동 확인 절차 3절 |
| 4 | 병렬 실행 간헐 FAIL(hold3 29h·29i) | **원인 확인·수정**: 첨부 업로드(FileReader 비동기) 완료를 **고정 600ms** 로 기다림 → CPU 부하 시 미완료 상태에서 검사. 테스트를 **첨부 반영까지 폴링(최대 15초)** 으로 변경(앱 코드 무변). 재현: 21개 동시 실행 1회 + hold3 10개 동시 + 부하 4개 → **47/47 × 10회** |

## 2. 결정 반영
- ① 업무 버튼 단락 = **숨김 허용 유지** (GPT·사용자 동일) · 조건 유지: 마지막 단락 끄기 금지 · [단락 n/m] 복구 · 주황 표시 · 안전 단락 잠금
- ② 추가 잠금 전수 조사(30화면 · 데모 데이터): 단락 본문의 HOLD·보류·격리·사고·차단·정합·저장 실패 등 문구/빨강 배지 검사
  - **생산 LOT 목록**(LOT HOLD 배지) → `data-sec-lock="safety"` + `data-sec-id="lot-list"` 추가 (이전엔 마지막 단락 규칙으로만 보호)
  - 나머지 해당 문구는 설정 화면의 설명문(예: 「품질 판정 지체 기준」)·규칙 안내문 — 경고가 아님
  - 출하 차단·HOLD 차단 메시지 · 서버 저장 거부 알림(`_fsWriteRejected`) · 동기화 상태는 **단락 밖**(오류 토스트·배너·상단 표시) → 숨김 대상 아님 · 차단 자체는 write path(`assertLotNotHeld`·`assertSoShipQty`)가 함 → 화면을 숨겨도 위험 동작은 막힘
  - 회귀 sec **10f**: HOLD·사고 배지를 보여주는 단락은 전부 명시 잠금이어야 PASS (현재 명시 잠금 3 = 생산 LOT · 부적합 목록 · 동시성 사고)

## 3. Firestore 실서버 수동 확인 (사용자 · 5분)
1. PC A: 정식 계정 로그인 → 테스트 주소 → 외주 화면 「최근 외주 입고」 끄기 → F5 → 꺼진 상태 유지 + [단락 1/2]
2. PC B(또는 시크릿 창): **같은 계정** 로그인 → 외주 화면 → 꺼져 있으면 서버 저장 OK
3. 같은 PC 에서 **다른 계정** 로그인 → 외주 화면 → 둘 다 켜져 있으면 사람별 저장 OK
4. 원복: [단락] → [모두 켜기] · 저장 거부 빨간 알림이 뜨면 Rules 확인 필요(사용자 Firebase 콘솔)

## 4. 회귀 (21개 동시 병렬 · 로그 첨부)
- v0.35.10 499 · otd 27 · sec 24 · docedit 11 · soclose 15 · mobile 39 · nav377 32 · personal 41 · q60 18 · hold3 47 · dash357 22 · forms358 26 · 2tab 26 · pilot359 29 · sim 11 · correct-guard 12 · lock-live 18 · recid 15 · r4 7 · v21 45 · 스냅샷 IDENTICAL · 파스 OK · NUL 0 · **FAIL 0**
- v21 수동 4건(기존 v0.21 항목 · 이번 변경과 무관): J7 출하 초과 차단 UI(write path 는 soclose 가 검증) · K1-3 특채/폐기/반품 UI · N2·N3 백업 파일 다운로드→복원 왕복
- 실브라우저(Playwright · 데모 · 다크): v0.38.19 · `lot:lot-list[mark]` 스위치 없음 · `qc:qc-nc-list[mark]` · 단락 끄기 F5 유지 · 납기 실적 KPI 6칸 · 페이지 오류 0

## 5. 승격 판단
- 남은 미검증 = Firestore 실서버(3절 사용자 확인) · v21 수동 4건(기존)
- 사용자 최종 승인 시 `index.html` ← v0.38.19 (승격 전 백업 `backup/index-v0.38.15-승격전-백업.html`)

---
MADE BY JUNS · VELA SYSTEM · v0.38.19 (2026-10-09)
