# cardnews — 블로그 글 → 인스타그램 카드뉴스 렌더러

기획안 `docs/블로그_MCP_자동화_기획.md` Phase 2.5용. JSON 한 장을 채우면 1080×1350 카드 PNG(표지 + 항목 N장 + 마무리), 전체 보기 시트, 인스타 캡션 초안이 나온다.

## 사용
```bash
node tools/cardnews/render.js tools/cardnews/content.example.json tools/cardnews/out
```
결과: `out/card_01.png` … `out/card_07.png`, `out/contact_sheet.png`, `out/caption.txt`, `out/cards.html`(디버그용)

필요: Node 20+, Playwright + Chromium (`npm i -D playwright && npx playwright install chromium`)

## 내용 파일(content.json) 규칙
- `cover.title`·`items[].title`은 `<br>`(줄바꿈)과 `<em>`(강조색)만 허용. 그 외 HTML은 문자로 출력된다
- `items`는 5개 기준. 3~8개까지는 레이아웃이 유지된다(인스타 API 캐러셀 한도 10장 = 항목 8개 + 표지 + 마무리)
- `last.experience`는 **Duke가 발행 전에 직접 채운다**. 자동 파이프라인은 이 칸을 채우지 않는다(blog-write 원칙 1: 경험 지어내지 않음)
- GXC 내부 숫자·거래처·인물은 넣지 않는다(blog-write 원칙 3). 세무·노무 내용은 조문 확인 후, "작성일 기준" 문구 유지

## 글꼴
- Windows: 맑은 고딕을 자동 사용
- Linux(클라우드 세션): `tools/cardnews/fonts/`에 Noto Sans KR woff2 3개(400·700·900)를 두면 대체 사용. npm `@fontsource/noto-sans-kr` 패키지의 `files/noto-sans-kr-korean-{400,700,900}-normal.woff2`를 복사하면 된다(저장소에는 커밋하지 않음)

## 디자인 조정
`template.css`의 `:root` 변수(`--navy` 제목·표지, `--accent` 번호·포인트, `--paper` 배경, `--gold` 표지 강조)만 바꾸면 전체 톤이 바뀐다.

## 인스타 게시까지의 흐름 (Phase 2.5)
1. 블로그 글 발행 후 Claude가 content.json 생성(경험 칸 제외)
2. `render.js` 실행 → PNG·캡션
3. PNG를 공개 URL에 올림(이 저장소의 GitHub Pages `cards/` 폴더 또는 Google Drive 공개 링크 아님에 주의: 인스타 API는 직접 접근 가능한 이미지 URL이 필요)
4. Slack DM으로 contact_sheet + 캡션 미리보기 → Duke가 경험 한 줄 채우고 "게시"
5. Zapier(Instagram for Business, 캐러셀 게시) 또는 Meta Graph API로 게시
