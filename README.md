# Duke Lab

양남규(Duke)의 실험실 저장소. GitHub Pages로 배포됩니다.

```
.
├── index.html     허브 페이지 (site/ 와 decks/ 로 연결)
├── site/          프로필 사이트 (index.html · script.js · style.css)
└── decks/         포트폴리오 덱 6종
    ├── index.html   갤러리 (슬라이드 뷰어 + pptx/pdf 다운로드)
    ├── CONTENT.md   13장 공통 원고 (site/ 내용에서 추출)
    ├── BRIEFS.md    디자인 브리프 6종 (팔레트·모티프)
    └── 0N-<slug>/   build.js · deck.pptx · deck.pdf · cover.jpg · slides/*.jpg · meta.json
```

## 덱 다시 만들기

```bash
npm install pptxgenjs sharp react react-dom react-icons   # 저장소 밖 아무 곳
cd decks/01-mono-minimal && NODE_PATH=<그 node_modules> node build.js
```

폰트는 맑은 고딕 고정. 미리보기 이미지는 LibreOffice 렌더링이라 PowerPoint와 글자 폭이 조금 다를 수 있습니다.
