# 📊 네이버 셀아웃 대시보드

엑셀 파일을 업로드하면 **자동으로** 대시보드가 업데이트됩니다.

## 🔗 대시보드 링크
👉 `https://[내 GitHub 아이디].github.io/sellout-dashboard`

---

## 📁 폴더 구조

```
sellout-dashboard/
├── data/
│   └── sellout.xlsm        ← ✅ 여기에 엑셀 파일 업로드
├── parse_dashboard.py      ← 자동 실행 파이썬 스크립트
├── index.html              ← 자동 생성되는 대시보드 (건드리지 마세요)
└── .github/
    └── workflows/
        └── update_dashboard.yml  ← 자동화 설정
```

---

## 🚀 매일 사용법 (30초)

### 방법 A: GitHub 웹에서 (가장 쉬움)
1. `data/` 폴더 클릭
2. **Add file → Upload files**
3. 엑셀 파일 드래그 앤 드롭 (파일명 아무거나 OK)
4. **Commit changes** 클릭
5. 1~2분 후 링크 자동 업데이트 ✅

### 방법 B: GitHub Desktop 앱 사용
1. `data/` 폴더에 엑셀 파일 복사
2. GitHub Desktop에서 Commit → Push
3. 자동 업데이트 ✅

---

## ⚙️ 최초 설정 (한 번만)

### 1. GitHub Pages 활성화
- repo → **Settings** → **Pages**
- Source: **Deploy from a branch**
- Branch: **main** / `/ (root)` → **Save**

### 2. Actions 권한 허용
- repo → **Settings** → **Actions** → **General**
- "Workflow permissions" → **Read and write permissions** 선택 → Save

---

## 📊 대시보드 구성

| 섹션 | 내용 |
|------|------|
| KPI 카드 | 전일 매출 / MOM / YOY / MTD / YTD |
| 일별 트렌드 | Philips + Sonicare 일별 매출 (꺾은선) |
| 카테고리 믹스 | MG / OHC / BT 주간 수량 (막대) |
| Weekly 비교 | 2026 vs 2025 YTD 누적 + CSG% |
| SKU TOP 10 | 전일 기준 / 월 누적 기준 TOP 10 |

---

## 🛠 문제 해결

| 문제 | 해결 |
|------|------|
| Actions가 실행 안 됨 | Settings → Actions → Read and write permissions 확인 |
| 대시보드 링크가 없음 | Settings → Pages 설정 확인 |
| 데이터가 이상함 | data/ 폴더의 파일이 올바른 xlsm인지 확인 |
| Actions 로그 확인 | repo → Actions 탭 → 최근 실행 클릭 |
