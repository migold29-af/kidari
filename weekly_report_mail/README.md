# 플랫폼기획부 주간보고 메일 생성기

주간 보고 회의록 최종본(구글 시트 「플랫폼기획부 주간 보고」의 `주간 보고(M/D)` 탭)으로
매주 발송하는 `[플랫폼기획부] YY.MM.DD - YY.MM.DD 주간보고` 메일 본문을 만듭니다.
결과 `mail.html` 을 브라우저로 열고 **[본문 복사] → 메일 작성 창에 붙여넣기** 하면 표·색상·링크가 유지됩니다.

## 사용법

```bash
# 회의록 xlsx(시트 > 파일 > 다운로드 > Microsoft Excel)에서 생성
python3 build_mail.py from-xlsx 플랫폼기획부_주간보고.xlsx --sheet "주간 보고(11/9)" -o out/261109

# 지난주 메일(.eml)을 JSON으로 → 수정 → 생성
python3 build_mail.py from-eml 지난주.eml -o out/base.json
python3 build_mail.py build out/base.json --period "26.11.02 - 26.11.06" -o out/261109
```

옵션: `--period "26.11.02 - 26.11.06"`(기간 직접 지정), `--note "*24~25 추석 연휴"`(기간 아래 회색 메모), `--config 파일`.

## 출력 (`out/<폴더>/`)

| 파일 | 용도 |
|---|---|
| `mail.html` | 미리보기 + [본문 복사]/[제목 복사]/[받는 사람·참조 복사] 버튼 |
| `body.html` | 본문 HTML만 (HTML 편집 모드용) |
| `mail.txt` | 텍스트 전용 본문 |
| `subject.txt` | 제목 |
| `report.json` | 사용한 데이터 (다음 주 출발점) |

## 설정

`config.example.json` 을 `config.local.json` 으로 복사해 이름·서명·수신자를 채웁니다. `config.local.json`, `out/`, `*.eml`, `*.xlsx` 는 git에 올라가지 않습니다.

## 알려진 한계

- 보고 기간은 "회의일(탭 날짜)의 직전 주 월~금"으로 자동 계산합니다. 9/28 발송 메일 1건으로 추정한 규칙이니 매주 확인하세요.
- xlsx 파서는 `[RELEASE]` 같은 섹션 제목 행과 `JIRA KEY` 가 들어간 헤더 행을 기준으로 표를 찾습니다. 실제 시트로는 아직 검증하지 못했습니다(합성 시트와 9/28 발송 메일로만 확인). 실제 시트에서 행 수가 맞지 않으면 알려 주세요.
- 셀 안 굵게·부분 색상 같은 서식은 옮기지 않습니다. 상태 색상, 헤더 색, 배포일 굵게, JIRA 링크만 자동 적용됩니다.

Claude Code에서는 `weekly-report-mail` 스킬(`.claude/skills/weekly-report-mail/SKILL.md`)이 이 과정을 진행합니다.
