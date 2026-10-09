# VELA MES v0.38.22 — v0.38.21 GPT 판정(CONDITIONAL PASS) 반영 보고

- 배포: index-dev.html = **v0.38.22** · index.html = v0.38.20 유지 · 백업 `backup/vela-mes-v0.38.21-백업.html`
- index-dev.html SHA-256 = `22d3a137e6877fcd7be92cdefac124712f8729c5f8d2524693edeac5a0d24467`
- ⚠ 지난 커밋(22c69f2)에 `_verify-shipfix.js` 가 빠져 있었음 → 이번 업로드 묶음에 포함 (`github-upload-v0.38.22/` 전체를 올림)

## 1. 판정 반영

| 판정 | 조치 | 증거 |
|---|---|---|
| ① 실운영: 플래그를 켜도 실데이터 저장 차단 | `shipQtyFixEnabled()` = 플래그 ∧ (데모 ∨ `SERVER_GATES.ship_qty_fix === true`) · `SERVER_GATES.ship_qty_fix = false` 고정(서버 Gate 통과 시 별도 SPEC·판정·승인으로만 변경) · 실데이터에서는 수량 칸 없음 + 「(수량 정정 = 시험 기능 · 데모에서만)」 표시 · 검사 함수도 거부 | shipfix **G1** |
| ② 레거시 초과: 감소 허용 + 계속 표시 | 감소 정정 허용·경고 기록(유지) + **감사 › 「원장 정합성 · LOT 출하 초과」 단락 신설**(`lotShipOverage` 파생 · 저장 0 · `data-sec-lock="safety"` 숨김 불가) · 증가는 계속 차단 | shipfix **F10b · G3** |
| ③ 5회·파렛트·명세서 | 그대로 유지(시험 기간) · 정식 운영 시 「관리자 검토 후 추가 정정」 은 관리자 인증(B′) 이후 과제로 기록 | — |
| **P0 다중 문서 추적 정합성** | **정정 = 출하 문서 1건만 쓴다**(단일 문서 = 원자적) · NC `pending_out`·사고 `pending_log` 에 쓰던 `shipQtyFixTrace` **삭제** → 출하 문서의 `qty_hist` 가 유일한 권위 · NC 상세·사고 이력은 `shipQtyFixEvents` 로 **읽어서 파생 표시** · 추가 쓰기가 필요해지는 경우(증가분이 판정 대기·동시성 보류에 걸림) = **증가 정정 차단** | shipfix **F6**(NC 문서 무변 · 파생 표시 「59 → 55ea (현재 유효 55ea)」 · 판정 대기 LOT 증가 차단) |
| **P1 필수 연결** | 출하의 수주·LOT 연결 비어 있음 / 수주·LOT 품번 비어 있음 → 차단 (불일치 검사 우회 제거) | shipfix **G2** |

- 결과: 출하 수량 정정의 저장 경로는 **`DB.correct('shipments', …)` 1회뿐** — 부분 실패로 일부 기록만 남는 경로 없음 (서버 동시성·권한은 여전히 미해결 → ①로 실데이터 차단)

## 2. 검증 (22개 동시 실행 · 전부 exit 0 · 로그 첨부)
- shipfix **17/17**(신규 G1·G2·G3 + F6 재작성) · docedit 11(6번: 켜짐 검사는 데모 세션에서) · sec 24(10c 대상 단락을 정확한 제목으로 — 감사에 「정합성」 단락이 추가된 영향) · otd 28 · v0.35.10 499 · soclose 15 · mobile 39 · nav377 32 · personal 41 · q60 18 · hold3 47 · dash357 22 · forms358 26 · 2tab 26 · pilot359 29 · sim 11 · correct-guard 12 · lock-live 18 · recid 15 · r4 7 · v21 45(+수동 4) · 스냅샷 IDENTICAL
- 실브라우저(데모): 수량 정정 150 → 138 저장 · ✎ 정정 · 감사 「원장 정합성 · LOT 출하 초과」 12 LOT 표시 · 페이지 오류 0

## 3. 발견 (정직 표기)
- 데모 시드가 출하를 실적과 독립 생성 → 데모 12 LOT 이 「출하 초과」 로 표시됨(데모 데이터 품질 문제 · 시드 수정은 스냅샷 기준값이 바뀌므로 별도 판단)

## 4. 다음 (GPT 승인 범위)
- v0.38.23: **구매 입고 수량 정정** 상세 설계부터 (단일 문서 여부 · 발주 잔량 파생 여부 확인 후)
- 소재 kg(`material_receipts`+`material_lots` 2문서) · 외주 반출(추적 영향) = 원자성 설계 후 · 외주 입고 = 계속 보류

---
MADE BY JUNS · VELA SYSTEM · v0.38.22 (2026-10-09)
