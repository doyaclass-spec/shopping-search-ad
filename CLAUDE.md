# AI 쇼핑검색광고 세팅 도구 — 개발 컨텍스트

## 프로젝트 개요
네이버 쇼핑검색광고(쇼검) 세팅 자동화 도구 v3.7
포쿨(Pokool) — 스타리온 냉장고 공식 판매채널 운영용

## 기술 스택
- **React 18 + Vite** (단일 파일 SPA: src/App.jsx)
- **xlsx** — 연관검색어/광고리포트 엑셀 파일 파싱
- **Claude API** (claude-sonnet-4-20250514) — 광고 세팅 AI 분석
- **localStorage** — 데이터 영속성 (window.storage 폴리필: src/main.jsx)

## 핵심 구조

### 캠페인 구조
```
Camp
├── id, number, name, adType, brand, model, mainKw
├── campTotalBudget, status (세팅중|진행중|일시정지|종료)
├── history[], memo
└── groups[4]  ← 항상 4개 고정
    ├── 0: 가격비교 메인 (GRP_META[0])
    ├── 1: 가격비교 서브 (GRP_META[1])
    ├── 2: N+스토어 메인 (GRP_META[2])
    └── 3: N+스토어 서브 (GRP_META[3])
```

### 제외 키워드 교차 계산 (중요)
- 쌍 구조: 0↔1 (가격비교), 2↔3 (N+스토어)
- 각 그룹 제외키워드 = 파트너 그룹 등록키워드 - 내 등록키워드
- `calcExcludeKeywords()` 함수로 자동 계산
- registerKeywords 변경 시 자동 재계산

### 주요 컴포넌트
- `App` — 메인 상태 관리, 사이드바 캠페인 목록
- `CampHeader` — 캠페인 헤더, ON/OFF 토글
- `HistoryPanel` — 활동 기록 (접기/펼치기)
- `MemoPanel` — 캠페인 메모
- `AIGeneratePanel` — 파일 업로드 + Claude API 분석 (4단계: input→analyzing→review→done)
- `FourGroupGrid` → `CardPair` — 4개 그룹을 2쌍으로 묶어 섹션 등높이 보장
- `CampModal` — 캠페인 추가/편집 (필수 필드 검증, 예산 쉼표 포맷)
- `ApiKeyModal` — Claude API 키 설정

### AI 분석 흐름
1. 사용자: 상품명 + 파일(연관검색어/광고리포트/노출키워드/대표이미지) 업로드
2. `callClaudeWithImage()` → Claude API 호출 (이미지 + 텍스트 멀티모달)
3. JSON 파싱 → summary(메인/서브 키워드+검색량, 입찰가표, 총예산) + groups[4]
4. `applyGeneratedGroups()` → registerKeywords 통일 → calcExcludeKeywords → setCamps 단일 호출

### 입찰가 계산
- 통합 예상 입찰가: MO×0.65 + PC×0.35 (가중평균, 10원 단위)
- 그룹별 비율: 1번 ×110%, 2번 ×55%, 3번 ×65%, 4번 ×40%
- 하루예산 비율: 1번 47%, 2번 20%, 3번 20%, 4번 13%

## 개발 명령어 (cmd 사용 — PowerShell 실행정책 문제)
```
npm run dev      # 개발서버 (http://localhost:3000)
npm run build    # 프로덕션 빌드 → dist/ 폴더 생성
npm run preview  # 빌드 결과 미리보기
```

## 배포 — GitHub Pages (무료, 인터넷 어디서나 접속 가능)

### API 키 보안 문제 없음
코드에 API 키가 없고, 각 브라우저의 localStorage에 저장됨.
GitHub에 올려도 키가 노출되지 않음.

### 배포 방법 (최초 1회 세팅)

#### 1. vite.config.js 수정 (GitHub 저장소명 맞게)
```js
export default defineConfig({
  plugins: [react()],
  base: '/forcool-ad/',  // GitHub 저장소 이름
})
```

#### 2. package.json에 deploy 스크립트 추가
```json
"scripts": {
  "dev": "vite",
  "build": "vite build",
  "preview": "vite preview",
  "deploy": "npm run build && gh-pages -d dist"
},
"devDependencies": {
  "gh-pages": "^6.0.0"
}
```

#### 3. gh-pages 설치 및 배포
```
npm install --save-dev gh-pages
npm run deploy
```

#### 4. GitHub 저장소 Settings → Pages → Source: gh-pages branch

### 배포 후 URL
```
https://[GitHub계정명].github.io/forcool-ad/
```

### 이후 업데이트 방법
코드 수정 후:
```
npm run deploy
```
1~2분 후 반영됨.

## 알려진 제약사항
- `anthropic-dangerous-direct-browser-access: true` 헤더 필요 (브라우저 CORS 우회)
- 노출 상품명: 23~25자 엄수 (네이버 광고 시스템 제한)
- window.storage 폴리필이 main.jsx에 있어야 로컬에서 정상 동작
- cmd 사용 권장 (PowerShell 실행정책 제한)

## 다음 개발 예정
- 엑셀 내보내기 (4개 그룹 세팅 → 네이버 광고 업로드 포맷)
- 성과 직접 입력 UI (클릭/비용/전환 → ROAS 자동 계산)
