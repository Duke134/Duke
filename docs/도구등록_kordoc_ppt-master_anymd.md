# 외부 도구 등록 안내 — kordoc · ppt-master · anymd

등록일 2026-11-01. 선별 근거는 GitHub 조사(경영지원 직무 기준)에서 채택된 3개.
이 저장소를 Claude Code로 열면 아래 설정이 자동 적용된다. 처음 열 때 폴더 신뢰(trust) 확인과 `.mcp.json` 서버 승인 프롬프트가 한 번 뜬다.

| 도구 | 역할 | 등록 위치 | 상태 |
|---|---|---|---|
| kordoc 4.24.35 | HWP·HWPX·PDF·Office → Markdown, 양식 채우기, 신구대조 (MCP 22도구) | `.mcp.json` | 설치 대기 |
| anymd 8.5.7 | PDF·Word·PPT·Excel·이미지(OCR) → Markdown (MCP 4도구: inspect·outline·read·search) | `.mcp.json` | 설치 대기 |
| ppt-master 6.8.0 | 문서 → 네이티브 편집 가능 PPTX, 자체 템플릿·슬라이드 마스터 사용 (스킬) | `.claude/settings.json` 마켓플레이스 | 설치 대기 |

## 사전 요건 (Duke PC)
- Node.js 20 이상 (kordoc·anymd 모두 npx로 실행, 별도 설치 불필요)
- Python 3.10~3.14 (ppt-master 후처리 스크립트). Windows는 설치 시 "Add to PATH" 체크
- Windows에서 Claude Desktop이 npx를 못 찾으면 `"command": "cmd", "args": ["/c", "npx", "-y", "kordoc", "mcp"]` 형태로 바꾼다 (kordoc 문서 기준)

## 1. kordoc · anymd (MCP)
이 저장소 안에서는 `.mcp.json`이 자동 인식된다. 첫 실행 시 승인만 하면 끝.
전역(어느 폴더에서나) 쓰려면 아래 한 줄씩:

```bash
claude mcp add --scope user kordoc -- npx -y kordoc mcp
```

```bash
claude mcp add --scope user anymd -- npx -y @sylphx/anymd
```

Claude Desktop까지 한 번에 등록하려면 각 패키지의 설치 마법사:

```bash
npx -y kordoc setup
```

```bash
npx -y @sylphx/anymd setup
```

접근 범위 제한: `.mcp.json`은 `KORDOC_ROOT`와 `--allow-dir`로 프로젝트 폴더 안만 읽게 묶어 두었다. 다른 폴더 문서를 다루려면 그 폴더에서 CLI를 직접 쓴다.

```bash
npx -y kordoc 공고문.hwpx -o 공고문.md
```

```bash
npx -y @sylphx/anymd 제안서.pdf > 제안서.md
```

OCR(스캔 PDF·이미지)은 kordoc `--ocr` 옵션이 내장 한국어 PP-OCRv5 모델(약 18MB)을 첫 사용 시 받는다. anymd OCR은 `anymd setup ocr`로 약 2GB 모델을 별도 옵트인해야 하므로, 한국어 스캔 문서는 kordoc을 먼저 쓴다.

## 2. ppt-master (스킬)
`.claude/settings.json`에 마켓플레이스가 등록되어 있어 Claude Code가 설치를 제안한다. 직접 설치하려면:

```bash
claude plugin marketplace add hugohe3/ppt-master
```

```bash
claude plugin install ppt-master@ppt-master
```

설치 후 반드시 Python 의존성을 넣는다. 설치된 플러그인 폴더(`claude plugin details ppt-master`로 경로 확인) 안에서:

```bash
pip install -r requirements.txt
```

의존성 19개(python-pptx·PyMuPDF·Pillow·numpy 등). 2026-11-01 Linux·Python 3.13 환경에서 설치 성공 확인.

### 한국어 폰트
스킬 내부 참조 문서가 PPT 안전 글꼴로 한국어 산세리프 `Malgun Gothic`(맑은 고딕), 세리프 `Batang`을 명시하고, CJK 글자폭을 별도 보정(calibrate)한다. GXC 규칙(docx·pptx 맑은 고딕 고정)과 맞는다. 첫 사용 시 "글꼴은 Malgun Gothic 고정"이라고 말하고 샘플 1장을 뽑아 줄바꿈·겹침을 확인한 뒤 본격 사용한다.

### 주의
- PDF 입력 변환부(PyMuPDF)는 AGPL-3.0. 사내 사용은 무관하나 ppt-master를 포함한 소프트웨어를 외부 배포할 일은 없으므로 영향 없음
- AI 이미지 생성·웹 검색·음성 내레이션은 선택 기능이며 API 키를 넣을 때만 외부로 나간다. 기본 생성은 로컬

## 검증 기록 (2026-11-01, 클라우드 컨테이너 Linux)
- kordoc: `npx -y kordoc mcp` MCP 핸드셰이크 성공, 도구 22개 확인
- anymd: `npx -y @sylphx/anymd` MCP 핸드셰이크 성공, 도구 4개 확인
- ppt-master: 저장소 clone 후 `pip install -r skills/ppt-master/requirements.txt` 성공(exit 0). 스킬 폴더 78MB라 저장소에 포함하지 않고 마켓플레이스 설치로 둠
- 미검증: Windows 환경 실행, ppt-master 실제 PPTX 산출(Duke PC에서 샘플 1장 필요)
