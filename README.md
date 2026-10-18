# README

Jira 프로젝트를 Work Breakdown Structure의 양식으로 변환합니다.

1. Jira 프로젝트의 목록을 `Jira.csv` 파일로 내보낸다.
2. `Jira.csv` 파일의 내용을 `work-breakdown-template.xlsx`의 양식에 따라 만든다.
3. 출력 파일의 이름은 `work-breakdown-structure-2026-xx-xx.xlsx` 이다.

## 변환규칙

- 가중치에는 story point estimate 값이 들어간다.
- 에픽과 스토리에 할당된 기간은 #B4C7E7 색으로 칠한다.
- 서브태스크에 할당된 기간은 #DDEBF7 색으로 칠한다.
- 진행률은 완료되었으면 100% 아니면 0%이다.
- 계층적 구조에 따라 항목을 정렬한다.
- 에픽, 스토리, 서브태스크 순으로 정렬한다.
- WBS에는 이슈 키가 들어간다.
- 업무에는 요약이 들어간다.