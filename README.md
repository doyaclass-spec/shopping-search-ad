# forcool-ad 개발 시작 가이드

## 실행 방법 (반드시 cmd 사용)
PowerShell은 실행정책 오류가 있으므로 cmd를 사용하세요.

```
cd C:\Users\LJY\Desktop\forcool-ad
npm run dev
```

→ 브라우저에서 http://localhost:3000 자동으로 열림

## App.jsx 세팅 (최초 1회)
이 채팅에서 다운로드한 forcool-ad-tool.jsx (v3.7) 파일을
C:\Users\LJY\Desktop\forcool-ad\src\App.jsx 로 복사하세요.

## Claude Code 사용 방법
```
cd C:\Users\LJY\Desktop\forcool-ad
claude
```

## 폴더 구조
forcool-ad/
├── src/
│   ├── App.jsx      ← 메인 파일 (v3.7 내용)
│   └── main.jsx     ← window.storage 폴리필 포함
├── index.html
├── package.json
├── vite.config.js
├── CLAUDE.md        ← Claude Code 컨텍스트
└── README.md        ← 이 파일
