---
name: weekly-report-mail
description: 플랫폼기획부 주간 보고 회의록 최종본이 나오면, 매주 보내는 "[플랫폼기획부] YY.MM.DD - YY.MM.DD 주간보고" 메일을 메일 작성 창에 바로 붙여넣을 수 있는 HTML로 만든다. "주간보고 메일 써줘", "회의록 최종본 나왔어", "이번 주 주간보고 메일" 같은 요청에 사용.
---

# 플랫폼기획부 주간보고 메일 작성

회의록 최종본 → `weekly_report_mail/build_mail.py` → `mail.html`(본문 복사 버튼) 순서로 만든다.
기존 메일과 같은 서식(검정 헤더, 빨간 특이사항 헤더, 상태별 색상, JIRA 키 링크)을 쓰되, 메일에서는 "구분" 열을 뺀 10개 열로 만든다(설정 `hide_columns`).

## 1. 회의록 최종본 확보 (아래 중 가능한 것)

원본: Google Drive 시트 **「플랫폼기획부 주간 보고」**
(`1YVFvpkROSPiAPMKwUC_q4gmdzWK5WKJ7IPVsP6Q92u0`), 탭 이름 `주간 보고(M/D)` = 회의일(월요일).

- **A. xlsx 파일** (가장 정확): 사용자가 올린 xlsx, 또는 Drive MCP
  `download_file_content(fileId, exportMimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")`
  결과를 base64 디코딩해 저장. 파일이 크므로(약 0.5MB) 컨텍스트에 그대로 출력하지 말고 파일로만 저장한다.
  ```bash
  python3 weekly_report_mail/build_mail.py from-xlsx <파일.xlsx> --sheet "주간 보고(11/9)" -o weekly_report_mail/out/<YYMMDD>
  ```
  `--sheet` 를 빼면 가장 최근 `주간 보고(M/D)` 탭을 쓴다. xlsx 내보내기에서 탭 이름의 `/` 가 다른 문자로 바뀌어도 찾는다.
- **B. 지난주 메일(.eml) + 이번 주 변경분**: 시트를 못 읽을 때.
  ```bash
  python3 weekly_report_mail/build_mail.py from-eml <지난주.eml> -o weekly_report_mail/out/base.json
  ```
  `base.json` 을 회의록/사용자가 준 변경분대로 고친 뒤(아래 3) `build` 로 생성.
- **C. 사용자가 회의록 내용을 붙여넣은 경우**: 아래 JSON 형식으로 직접 채운다.

Confluence는 이 환경에서 IP 제한으로 막혀 있을 수 있다. 막히면 A~C로 진행하고 그 사실을 알린다.

## 2. 보고 기간

메일 제목/본문 기간 = **회의일이 속한 주의 직전 주 월~금** (예: 9/28 회의 → `26.09.21 - 26.09.25`).
from-xlsx 는 탭 이름으로 자동 계산하고 `report.json` 의 `period_inferred_from` 에 근거를 남긴다.
이 규칙은 9/28 발송 메일 1건에서 추정한 것이므로 결과를 사용자에게 확인받는다. 다르면 `--period "26.11.02 - 26.11.06"`.
연휴 등 메모는 `--note "*24~25 추석 연휴"`.

## 3. JSON 형식 (build 입력)

```json
{
  "period_label": "26.11.02 - 26.11.06",
  "notes": ["*메모(선택)"],
  "sections": [
    {"title": "RELEASE", "rows": [
      {"no": "1", "type": "TASK", "jira": "BALCONY-28103", "platform": "KR",
       "project": "[KR] ...", "detail": "줄바꿈은 \n", "schedule": "- 개발: ~11/3",
       "status": "배포", "date": "11월 3일", "owner": "담당자", "note": "특이사항"}
    ]},
    {"title": "DATA", "rows": [
      {"no": "1", "type": "DATA", "jira": "-", "platform": "KR", "project": "KPI 설계",
       "steps": [{"detail": "지표 설계", "schedule": "9/28~9/30", "status": "진행", "owner": "담당자",
                  "note": "[🔗문서](https://...)"}]}
    ]}
  ]
}
```
- 섹션: `RELEASE`, `PROJECT`, `TASK`, `UX/UI DESIGN`, `DATA` (이 순서로 출력, 빈 섹션은 생략)
- `steps` 가 있으면 앞 5개 열(NO~Project)이 세로 병합되고 단계별로 행이 나뉜다.
- 셀 안 링크는 `[라벨](URL)`. JIRA 키는 자동으로 링크가 걸린다.
- 상태 색상: 배포·배포대기(빨강), 개발, 기획, 디자인, QA, 완료, 진행, 홀딩, 대기.
- `links`(■ 업무 현황), `shared`(■ 공유 사항)를 생략하면 설정 파일 값을 쓴다.

```bash
python3 weekly_report_mail/build_mail.py build report.json -o weekly_report_mail/out/<YYMMDD>
```

## 4. 결과 확인과 전달

1. 출력 JSON의 섹션별 행 수를 회의록과 대조한다. 회의록에 없는 내용을 지어내지 않는다. 비어 있거나 애매한 칸은 비워 두고 사용자에게 목록으로 알린다.
2. `mail.html` 을 사용자에게 보낸다(SendUserFile, display=render). 사용법:
   **브라우저로 열기 → [본문 복사] → 메일 작성 창 본문에 붙여넣기(Ctrl+V)**. 제목은 [제목 복사].
3. `subject.txt`, `mail.txt`(텍스트 전용) 도 함께 있다.
4. 메일을 직접 보내지 않는다. 발송은 사용자가 한다.

## 설정 (`weekly_report_mail/config.local.json`, git 제외)

`config.example.json` 을 복사해 보내는 사람 이름, 서명, 받는 사람/참조를 채운다.
조직 규칙에 따라 연락처·메일주소는 레포에 커밋하지 않는다. 없으면 예시 설정으로 만들어지고 인사말은 "안녕하세요, 플랫폼기획부입니다."가 된다.
