# 블로그 MCP 자동화 기획 (1→100) — 네이버 블로그 · 구글 Blogger

작성일 2026-11-01(일). 작성 기준: Duke 블로그 운영 스킬 `blog-write v1.0`, Duke 프롬프트 v5.0 안전 규칙, 2026-11-01 웹 조사·npm 레지스트리 확인, Duke가 공유한 Notion 가이드 「클로드로 네이버 블로그 자동 발행하기」(업리카 연결법, 2026-10-27 수정).

---

## 0. 결론 먼저

| 질문 | 답 |
|---|---|
| 1→100 전부 자동화가 되나 | **구글 Blogger는 된다.** 공식 Blogger API v3로 초안 저장·발행·통계 조회까지 코드로 가능 |
| 네이버 블로그도 되나 | **1→95까지 된다.** 네이버 공식 글쓰기 API는 종료됐지만(2020년경, 기사 기준), Duke가 공유한 Notion 가이드의 **업리카(uplika) MCP**가 Duke PC의 크롬 확장을 통해 로그인된 블로그에 글을 올린다. 비공개로 먼저 올리고 Duke가 공개 전환하는 구조라 승인 게이트가 자연스럽게 생긴다. 다만 Duke 크롬이 켜져 있어야 하고(무인 불가), 네이버가 "매크로·어뷰징"으로 제재한다고 밝힌 영역이라 **마지막 5(공개 전환)는 Duke가 누른다** |
| 클라우드 크레딧은 어디에 쓰나 | (1) Claude API 토큰, (2) 무인 실행기(Cloud Run Jobs + Cloud Scheduler), (3) 이미지 저장소. 다만 **1단계는 크레딧 없이 Claude Code 예약 작업만으로 돌릴 수 있다** |
| 지금 세팅된 것 | 없음. 이 저장소에는 kordoc·anymd만 등록되어 있고 블로그용 MCP·GitHub 저장소는 0개. 필요한 목록은 §4 |

추천안: **Blogger는 API로 완전 자동화(무인 가능) → 네이버는 업리카 MCP로 비공개 업로드까지 자동화하고 공개 전환만 Duke → Playwright 직접 자동화는 업리카가 막히거나 유료화될 때의 예비안(Phase 4).**

업리카는 Notion 가이드 기준 베타 무료, 정식 출시 후 월 20편 무료라고 쓰여 있다. 이 세션에서는 네트워크 정책에 막혀 uplika.com 접속·약관을 직접 확인하지 못했다. `[미확인]`

---

## 1. 전제 (2026-11-01 Duke 결정 반영)

1. **개인 블로그 기준.** `blog-write` 스킬 그대로 쓴다. **회사(GXC) 공식 블로그는 추후 연동 예정** → 같은 파이프라인을 재사용하되 별도 스킬(대외 표현 규칙·대표님 승인·기밀 필터 강화)로 분리한다. 개인/회사 채널이 섞이지 않게 MCP 서버 이름과 Blogger 블로그 ID를 채널별로 나눈다. (결정됨)
2. **클라우드 크레딧.** 추천안(§5-A Claude Code 예약 작업)은 Phase 1~3에서 크레딧을 쓰지 않는다. 크레딧은 Phase 4 무인화(Cloud Run 등)에서만 필요하므로 종류 확인은 그때 한다. 개인 크레딧이 아니라 회사 크레딧이면 그 시점에 대표님 확인. `[확인 필요 — Phase 4 전]`
3. **발행 게이트: Duke 승인 1클릭.** Blogger는 초안 생성 → Duke 승인 → publish, 네이버는 업리카 비공개 업로드 → Duke 승인 → 공개 전환. 완전 무인 발행은 쓰지 않는다. (결정됨)
4. **네이버: 업리카 + 공개 전환만 Duke.** Playwright 직접 자동화는 예비안으로만 둔다. (결정됨)
5. **실행기: Claude Code 예약 작업으로 시작.** (결정됨)
6. **수익화(광고·제휴)** 를 켜면 재직 중 겸업 규정 확인과 공정위 추천·보증 표기 자리가 필요하다(조문 미확인).

---

## 2. 전체 흐름 1→100 (10단계)

| # | 단계 | 하는 일 | 쓰는 도구 (MCP / 스킬) | 자동화 수준 |
|---|---|---|---|---|
| 1 | 주제·키워드 발굴 | 키워드 30개, 수요·경쟁 **실측** (네이버 검색광고 키워드도구, 데이터랩 트렌드) | naver-searchad MCP · naver-search MCP(datalab) · blog-write 모드 1-K | 자동 |
| 2 | 소재 수집 | 옵시디언 볼트(`03_SOP/`, `10_학습/`, `20_Portfolio/`)에서 Duke의 실제 경험 메모 검색 | obsidian MCP(`vault_*`) | 자동. **경험 재료가 없으면 멈추고 묻는다** |
| 3 | 제목·목차 | 제목 10개 → 3개 추천, 목차 5소제목 | blog-write 모드 2-1·2-2 | 자동(확인 지점 ★ 2곳은 설정으로 생략 가능) |
| 4 | 본문 | 2,000자 안팎, 경험 사례 1개 이상 | blog-write 모드 2-3 · Claude API | 자동 |
| 5 | 검수 | SEO 표 6항목, 기밀 필터(GXC 내부 숫자·거래처·인물), 법령 글이면 조문 확인 | blog-write 모드 2-4 · korean-law MCP | 자동 + 🔴 남으면 다음 단계 차단 |
| 6 | 이미지 | 썸네일 1장 + 본문 이미지 기획 5장 | Canva MCP(generate-design·export) · Google Drive MCP(저장) | 자동(사진은 Duke 직접 촬영분 우선) |
| 7 | 초안 저장 | 볼트 `80_블로그아이디어/초안/`에 저장 + Blogger에 `isDraft` 초안 업로드 | obsidian MCP · **blogger MCP** | 자동 |
| 8 | **승인 게이트** | Duke에게 Slack DM(본인) 또는 Gmail 초안으로 미리보기 링크 전달 → "발행" 답신 | Slack MCP(Duke 본인 DM만) | **Duke 1클릭** |
| 9 | 발행 | Blogger: `publish` 호출. 네이버: 업리카 MCP로 **비공개 업로드**(카테고리 지정, 사진 삽입) → Duke가 폰에서 확인 → "공개로 바꿔줘" | blogger MCP · **uplika MCP** · (예비: playwright MCP) | Blogger 자동 / 네이버 공개 전환만 Duke |
| 10 | 사후 처리 | `발행목록.md`·`키워드뱅크.md` 상태 갱신, 주 1회 통계 수집 → 주간 분석 → 다음 주 글 5개 제안 | obsidian MCP · Search Console MCP · Blogger pageViews · blog-write 모드 3 | 자동 |

흐름 요약: `키워드 → 소재 → 글 → 검수 → 이미지 → 초안(볼트+Blogger) → [Duke 승인] → 발행 → 기록 → 주간 분석 → 다시 키워드`.

---

## 3. 플랫폼별 발행 수단 비교

| 플랫폼 | 공식 API | 상태 | 자동 발행 가능성 | 리스크 |
|---|---|---|---|---|
| **구글 Blogger** | Blogger API v3 (`POST /v3/blogs/{blogId}/posts`, OAuth 2.0) | 운영 중 | **가능.** 초안(isDraft)·발행·수정·조회·페이지뷰 | 일 할당량 존재(기본 10,000 요청/일, 커뮤니티 보고로는 글 생성 50편/일 수준) `[미확인 — 공식 문서 직접 대조 못 함, 이 세션 네트워크 차단]` |
| **네이버 블로그 (업리카 MCP)** | 공식 API 아님. 서드파티 호스팅 MCP(`https://api.uplika.com/mcp`) + Duke PC 크롬 확장 | 베타(Notion 가이드 기준) `[미확인]` | **가능.** 비공개 업로드 → 같은 주소에서 수정 → 공개 전환. 사진 업로드 링크 제공, 카테고리 지정 | Duke 크롬이 켜져 있어야 함(무인 불가). 서드파티 서비스 지속성·유료화. 글 본문이 업리카 서버를 거친다. 네이버 제재 리스크는 아래와 동일 |
| 네이버 블로그 (직접 자동화) | 글쓰기 Open API | **종료**(기사 기준 2020년 5월 6일 이후 글쓰기 중단, 사유 "반복적·기계적 대량 발행") | 공식 경로 없음. 브라우저 자동화(Playwright)만 가능 | 네이버가 "매크로 등 어뷰징은 제재"라고 밝힘. 캡차·로그인 세션 유지·검색 노출 제한(이른바 저품질, 비공식 용어) |
| 티스토리 | Open API | 종료(2024-02) | 불가 | — |
| 워드프레스(자체 호스팅) | REST API | 운영 중 | 가능 | 호스팅 비용. 네이버 대체로 두 번째 완전자동 플랫폼이 필요할 때 후보 |

네이버는 업리카가 1순위다. 업리카가 안 되면 Playwright/browser-use로 글쓰기 화면을 직접 조작하는 길만 남는다. GitHub에 그런 도구(예: `sinmb79/blog-writer` — 전용 Chrome 프로필에 1회 수동 로그인 후 세션 재사용, MCP로 `platform="naver"` 발행, 실패 시 텔레그램 알림)가 있지만, 유지보수와 약관 리스크를 Duke가 전부 떠안는 구조다. 업리카도 결국 Duke 크롬으로 글을 쓰는 방식이라 네이버 입장에서는 같은 범주일 수 있다. 그래서 어느 쪽이든 "비공개 업로드 → Duke 공개 전환"과 주 편수 상한을 지킨다.

---

## 4. 필요한 GitHub 저장소 · MCP 서버 목록 (현재 0개 설치)

### 4-1. 이미 쓰고 있어 그대로 연결만 하는 것

| 용도 | MCP | 비고 |
|---|---|---|
| 볼트 읽기·쓰기 | obsidian-rest (`vault_*`) | Obsidian Local REST API 플러그인(`coddingtonbear/obsidian-local-rest-api`) 위에서 동작. Duke PC의 Obsidian이 켜져 있어야 하므로 **무인 실행기(§5-B)에서는 못 쓴다** → 무인 모드는 Git 동기화된 볼트 사본이나 Google Drive로 우회 |
| 법령 확인 | korean-law | 세무·노무 글 검수용 |
| 알림·승인 | Slack · Gmail (claude.ai 커넥터) | Duke 본인 DM·본인 수신 메일만 |
| 썸네일 | Canva (claude.ai 커넥터) | 이 세션에서 연결 확인됨 |
| 이미지·표 저장 | Google Drive · Sheets (claude.ai 커넥터) | 키워드뱅크를 Sheets로 복제하면 무인 모드에서도 읽을 수 있다 |

### 4-2. 새로 설치해야 하는 것 (핵심)

| 순위 | 역할 | 후보 | 설치 | 2026-11-01 확인 | 선택 이유·주의 |
|---|---|---|---|---|---|
| 0 | **네이버 발행** | 업리카(uplika) — claude.ai 사용자 정의 커넥터 `https://api.uplika.com/mcp` + 크롬 확장 | uplika.com 구글 가입 → 채널 연결(네이버) → 크롬 확장 설치 → 에이전트 연결(Claude) → 커넥터 추가·허용 | Notion 가이드 기준. 이 세션에서 접속 불가 `[미확인]` | 설치가 가장 쉽고 폰 Claude 앱에서도 동작. 연결 확인은 "내 블로그 카테고리 목록 보여줘". **GitHub 저장소 없음(호스팅 서비스)**, 약관·개인정보 처리방침 확인 후 사용 |
| 1 | **Blogger 발행** | `aleck31/mcp-blogger` (npm `mcp-blogger`) | `npx -y mcp-blogger` + `GOOGLE_CLIENT_ID/SECRET`, `DEFAULT_BLOG_ID` | npm 1.0.2(2026-02-13), MIT. README로 도구 확인 완료 | **초안 우선 설계**: `create_post`는 기본 초안, `change_post_status action=publish`로 발행, `revert`로 비공개 복귀, `list_drafts`·`get_post`로 미리보기 → 승인 게이트와 정확히 맞는다. OAuth 토큰은 `~/.config/mcp-blogger/`에 자동 저장·갱신(첫 1회 브라우저 동의). 대안 `@dalcontak/blogger-mcp-server`(1.0.3)는 초안/발행 분리 도구가 없어 2순위 |
| 2 | **키워드 검색량 실측** | `alphago2580/naver_searchad_mcp` (Python, `get_keyword_stats`) 또는 `2duckchun/naver-ads-mcp` (`npx -y github:2duckchun/naver-ads-mcp`, Node 20+) | 네이버 검색광고 API 키 필요(광고시스템 > 도구 > API 사용 관리, 무료). **npm·PyPI 미등록, GitHub에서 직접 받아야 함** | 디렉터리 등재만 확인 | blog-write 원칙 2(검색량은 추정)를 **실측**으로 바꿔 준다. 광고 계정 개설이 선행. Phase 1에서는 생략 가능(Phase 3 분석 루프에서 붙인다) |
| 3 | **검색 트렌드·블로그 검색** | `isnow890/naver-search-mcp` (npm `@isnow890/naver-search-mcp`) | `npx -y @isnow890/naver-search-mcp` + **`NCP_APIGW_API_KEY_ID/KEY`**(NAVER API HUB) | npm 1.0.54(2026-10-07), MIT. README로 도구 확인 완료 | 도구 18개: `search_blog`(경쟁 글), `search_webkr`·`search_news`, `datalab_search`(검색어 트렌드) 등. **주의: 네이버 개발자센터 키는 2026-07-31부터 신규 발급 차단(기존 키는 2027-06-30까지). 신규는 네이버클라우드 콘솔 > NAVER API HUB에서 발급** |
| 4 | **구글 유입 통계** | `ahonn/mcp-server-gsc` (npm `mcp-server-gsc`) 또는 `@sarahpark/google-search-console-mcp` | 서비스 계정 JSON + Search Console에 Blogger 속성 등록 | npm 0.3.0(2026-02-13) / 1.1.0(2026-09-19) | 모드 3 주간 분석의 데이터 소스. Blogger 자체 통계는 Blogger API `pageViews`로 보강 |
| 5 | **브라우저 자동화(네이버 예비안, Phase 4)** | `microsoft/playwright-mcp` (npm `@playwright/mcp`) | `npx -y @playwright/mcp@latest` | npm 0.0.84, 2026-10-31 | 공식·활발. 업리카가 종료·유료화되거나 막힐 때 네이버 글쓰기 화면 직접 조작용. 이 세션에는 `browser-use` 플러그인도 있어 대체 가능. **PoC 전 §6 리스크 결정 필수** |
| 6 | (참고) 네이버 발행 완성품 | `sinmb79/blog-writer` | 저장소 README 확인 | 릴리스 v3.3.0 확인 | 네이버·워드프레스 발행 + 텔레그램 알림이 이미 묶인 프로젝트. 코드 품질·유지보수는 미확인이므로 **구조 참고용** |

GitHub 직접 접속은 이 세션의 네트워크 정책에 막혀(403) 저장소 페이지는 열지 못했고, npm 레지스트리로 패키지 존재·최신 버전만 확인했다. 설치 전 Duke PC에서 각 README를 한 번 읽는다.

### 4-3. 자체 제작이 필요한 것

| 이름(가칭) | 왜 만드나 | 규모 |
|---|---|---|
| `blogger-mcp` (자체) | 후보 1(`mcp-blogger`)이 초안·발행 분리를 이미 지원하므로 **당장은 불필요**. `pageViews`(Blogger 자체 통계)가 필요해지는 Phase 3에서 Search Console로 충분하지 않을 때만 1파일 추가 | Node/TypeScript 1파일 수준, 1일 |
| `blog-pipeline` 스크립트 | 10단계를 순서대로 부르는 오케스트레이터. Claude Code 예약 작업(§5-A)에서는 프롬프트 1개로 충분하고, Cloud Run(§5-B)에서는 Claude API Tool Runner로 작성 | §5 선택에 따라 |
| `blog-write` 스킬 v1.1 | 모드 2-6 "저장"에 **Blogger 초안 업로드 + 업리카 비공개 업로드**, 모드 4 "승인·공개 전환" 추가. 공통 원칙 5를 "Duke 승인 전 공개 금지"로 다듬는다. Notion 가이드의 글쓰기 팁도 반영: ① Duke가 쓴 글 2~3편 링크로 말투 학습 ② AI 티 나는 표현 금지 목록("~해보세요", "다양한", "효과적으로", "~하는 것이 중요합니다") ③ 글마다 길이·구조 바꾸기 ④ 쓰기 전 질문은 3개 이내로 묶기 | skill-audit 절차로 .skill 파일까지 (설치 대기) |

### 4-4. 계정·키 준비 체크리스트 (Duke가 직접)

- [ ] Google Cloud 프로젝트 생성 → Blogger API 사용 설정 → OAuth 클라이언트(웹 애플리케이션, 리디렉션 `http://localhost:3000/oauth/callback`) 발급 → `mcp-blogger` 첫 쓰기 호출 때 브라우저 동의 1회
- [ ] Blogger 블로그 개설(blogspot) — API로는 블로그 생성 불가, 수동. 블로그 ID를 `DEFAULT_BLOG_ID`로
- [ ] Google Search Console에 Blogger 주소 등록 + 서비스 계정 권한 부여 (Phase 3)
- [ ] **NAVER API HUB**(네이버클라우드 콘솔 > Application Services > NAVER API HUB > Application 등록, 검색·데이터랩 선택) → `NCP_APIGW_API_KEY_ID` / `NCP_APIGW_API_KEY`
- [ ] 네이버 검색광고 계정 → API 라이선스 발급(무료) (Phase 3)

세부 절차는 `docs/블로그_MCP_세팅가이드.md`.
- [ ] 업리카: 구글 계정 가입 → 네이버 블로그 채널 연결 → 크롬 확장 설치(컴퓨터 크롬, 폰 불가) → claude.ai 사용자 정의 커넥터 추가 → 약관·개인정보 처리방침 1회 확인
- [ ] Slack: Duke 본인 DM 채널 ID 확인(승인 알림 수신용)
- [ ] 키 보관: `.env`는 저장소에 올리지 않는다. Cloud Run은 Secret Manager, 로컬은 Claude Code `settings.local.json`

---

## 5. 실행기 2안 — 어디서 돌리나

### 5-A. Claude Code 클라우드 예약 작업 (Phase 1~3 추천, 크레딧 거의 안 씀)

- 이 저장소를 Claude Code on the web 환경으로 열고, 예약 작업(매주 월·목 등)으로 `blog-write` 파이프라인 프롬프트를 실행
- MCP는 `.mcp.json`에 등록(Blogger·naver-search·searchad·gsc). 커넥터(Slack·Gmail·Canva·Drive)는 claude.ai 설정에서 연결
- 장점: 인프라 코드 0, 승인 게이트를 대화로 처리, 스킬 그대로 재사용
- 한계: 옵시디언 로컬 REST와 업리카 모두 Duke PC(크롬)가 켜져 있을 때만 동작 → 볼트를 Git으로 동기화해 저장소에서 읽거나, 키워드뱅크·발행목록을 Google Sheets로 복제. 네이버 발행은 Duke가 PC 앞에 있는 시간대(예: 평일 저녁)로 예약

### 5-B. Google Cloud 무인 실행기 (Phase 3 이후, 크레딧 사용)

| 구성 | 서비스 | 역할 | 비용 성격 |
|---|---|---|---|
| 스케줄 | Cloud Scheduler | 주 N회 트리거 | 월 수 개 잡은 무료 구간 수준 `[확인 필요]` |
| 실행 | Cloud Run Jobs | 파이프라인 컨테이너(Node/Python, Claude API Tool Runner + MCP 클라이언트) | 실행 시간만 과금 |
| 비밀 | Secret Manager | OAuth 토큰·API 키 | 소액 |
| 이미지 | Cloud Storage | 썸네일·본문 이미지 | 소액 |
| LLM | Claude API(Anthropic 직접) 또는 Vertex AI의 Claude | 본문·검수·분석 | §7 |
| 승인 | Slack 알림 + Cloud Run 승인 엔드포인트(버튼) | Duke 1클릭 → publish | — |

다른 크레딧이면: AWS = EventBridge Scheduler + Lambda/ECS + Secrets Manager + S3 / 네이버클라우드 = Cloud Functions + Object Storage(네이버 API와 같은 생태계라는 장점만 있고 Blogger 쪽 이점은 없음).

**관리형 대안**: Anthropic Managed Agents의 예약 배포(scheduled deployments)를 쓰면 스케줄러·컨테이너·Secret을 직접 운영하지 않아도 된다. Anthropic API 크레딧이 있을 때 5-B 대신 검토할 가치가 있다.

---

## 6. 네이버 자동 발행 리스크 (Phase 4 결정용)

| 무슨 일 | 영향 | 대응안 | 필요 액션 |
|---|---|---|---|
| 네이버가 자동 글쓰기를 어뷰징으로 제재 | 블로그 검색 노출 제한·계정 제재 | 발행 빈도 주 2~3편, 사람 속도의 입력, Duke 승인 후에만 발행 | Duke가 리스크 수용 여부 결정 |
| 업리카 서비스 종료·유료화·장애 | 네이버 발행 중단 | 글은 볼트+Blogger에 먼저 저장되므로 유실 없음. Playwright 예비안 또는 수동 붙여넣기 | 월 1회 서비스 상태·요금 확인 |
| 업리카 서버가 글 본문·사진을 경유 | 외부 서버에 콘텐츠 전달 | 개인 블로그 글만 보낸다. GXC 내부 숫자·거래처·인물은 blog-write 원칙 3으로 사전 차단 | 약관·개인정보 처리방침 확인 |
| 캡차·로그인 세션 만료 | 발행 실패 | 전용 Chrome 프로필 1회 수동 로그인·세션 재사용, 실패 시 Slack 알림 후 수동 | 운영 규칙 문서화 |
| 네이버 에디터 DOM 변경 | 스크립트 깨짐 | 글은 항상 볼트+Blogger에 먼저 저장되므로 유실 없음. 수동 붙여넣기로 폴백 | 월 1회 점검 |
| AI 생성 글 대량 발행으로 품질 저하 | 독자 이탈·노출 하락 | 경험 사례 1개 필수 규칙 유지, 주 편수 상한 | blog-write 원칙 1 유지 |

---

## 7. 비용 추정 (글 1편 기준)

**결과**: 글 1편당 LLM 비용 약 **0.3~0.6 USD**(Opus 5.5 기준, 추정). 주 3편이면 월 약 4~8 USD.

**계산과정** (추정값, 토큰 수는 가정):
- 1편 파이프라인 입력 합계 약 40,000토큰(볼트 소재·키워드 데이터·검수 프롬프트 반복 포함), 출력 약 12,000토큰(제목·목차·본문·SEO표·이미지 기획·분석)
- Opus 5.5 단가: 입력 4 USD/1M, 출력 20 USD/1M → 40,000×4/1,000,000 + 12,000×20/1,000,000 = 0.16 + 0.24 = **0.40 USD**
- Sonnet 5.5(입력 2, 출력 10)로 내리면 0.08 + 0.12 = **0.20 USD**
- 월 13편(주 3편) × 0.40 = **5.2 USD**

**근거**: 모델 단가는 claude-api 스킬 모델 표(캐시 기준일 2026-10-06). 토큰 수는 blog-write 모드 2 분량(본문 2,000자 안팎 + 부속물)에서 가정.

**주의사항**: Cloud Run·Scheduler·Storage는 이 사용량에서 무료 구간 또는 월 1 USD 미만일 가능성이 높지만 `[확인 필요]`. 이미지 생성(Canva/Imagen)은 별도. 프롬프트 캐싱을 쓰면 입력 비용은 더 내려간다.

✅ 검증완료 (단가·곱셈 재검산)

---

## 8. 단계별 로드맵

| Phase | 기간 | 목표 | 완료 기준 |
|---|---|---|---|
| 0 결정·계정 | 11/3(화)~11/7(토) | §1 가정 4개 결정, §4-4 계정·키 발급 | Blogger 블로그 개설 + OAuth 토큰 1회 발급 |
| 1 발행 채널 연결 | 11/10(화)~11/21(토) | **업리카 연결(가장 먼저, 반나절)** → blog-write로 쓴 글 1편을 네이버에 비공개 업로드. blogger MCP 설치 → 같은 글을 Blogger **초안**으로. 네이버 search·searchad MCP로 키워드 실측 | "볼트 초안 = 네이버 비공개 글 = Blogger 초안" 1편 |
| 2 반자동 운영 | 11/24(화)~12/5(토) | Claude Code 예약 작업 주 2회 + Slack 승인 게이트 → Blogger publish·네이버 공개 전환. 발행목록 자동 갱신 | 승인 1클릭으로 발행된 글 3편 |
| 3 분석 루프 | 12/8(화)~12/19(토) | Search Console·pageViews 수집 → 모드 3 주간 분석 자동 → 키워드뱅크 갱신 | 주간 분석 리포트 2회 자동 생성 |
| 4 예비안·무인화 (선택) | 2027-01-05(화) 이후 | 업리카 운영 2개월 평가(안정성·요금) → Playwright 예비안 PoC 여부. 크레딧 쓰는 Cloud Run 이전 여부 결정 | Duke 결정 문서 1건 |

---

## 9. 결정 기록 (2026-11-01)

| # | 항목 | 결정 |
|---|---|---|
| 1 | 대상 블로그 | **개인 블로그.** 회사 블로그는 추후 별도 스킬로 연동 |
| 2 | 크레딧 | Phase 4 전까지 불필요. 종류·소유는 그때 확인 |
| 3 | 발행 게이트 | **Duke 승인 1클릭** |
| 4 | 네이버 | **업리카 + 공개 전환만 Duke.** Playwright는 예비안 |
| 5 | 실행기 | **Claude Code 예약 작업**으로 시작 |
| 6 | 워드프레스 추가 | 보류 |

### 다음 할 일
- Duke: `docs/블로그_MCP_세팅가이드.md` 순서대로 업리카 연결 → Google Cloud OAuth → NAVER API HUB 키 발급 (Phase 0, 11/3(화)~11/7(토))
- Claude: 키가 들어오면 Phase 1 첫 글 1편을 "볼트 초안 = 네이버 비공개 = Blogger 초안"으로 올리고, 그 결과를 바탕으로 `blog-write v1.1`(.skill 파일) 작성

---

## 참고 출처 (2026-11-01 조회)

- 업리카 연결·운영 가이드: Duke 공유 Notion 「클로드로 네이버 블로그 자동 발행하기」(저작권 @bbaekong_toon, 2026-10-27 수정) — uplika.com·api.uplika.com은 이 세션에서 접속 불가

- 네이버 블로그 글쓰기 API 종료: [i-boss 기사](https://www.i-boss.co.kr/ab-2877-4595), [인프런 Q&A](https://www.inflearn.com/community/questions/1513877/)
- 네이버 어뷰징 제재 입장: [비즈한국](https://www.bizhankook.com/bk/article/29309)
- Blogger API 할당량(커뮤니티): [Make 커뮤니티](https://community.make.com/t/403-quotaexceeded-in-blogger/85292), [coupler.io](https://docs.coupler.io/sources/category/social-media/blogger/common-issues)
- Blogger MCP: [aleck31/mcp-blogger](https://playbooks.com/mcp/aleck31/mcp-blogger), dalcontak/blogger-mcp-server(npm)
- 네이버 검색광고 MCP: [2duckchun/naver-ads-mcp](https://glama.ai/mcp/servers/2duckchun/naver-ads-mcp), [alphago2580/naver_searchad_mcp](https://lobehub.com/it/mcp/alphago2580-naver_searchad_mcp)
- 네이버 검색·데이터랩 MCP: [isnow890/naver-search-mcp](https://lobehub.com/ko/mcp/isnow890-naver-search-mcp-impl)
- Search Console MCP: [ahonn/mcp-server-gsc](https://awesome.ecosyste.ms/projects/github.com%2Fahonn%2Fmcp-server-gsc)
- 네이버 발행 완성품 참고: [sinmb79/blog-writer v3.3.0](https://github.com/sinmb79/blog-writer/releases/tag/v3.3.0)
- Obsidian REST MCP 계열: [markuspfundstein/mcp-obsidian](https://awesome.ecosyste.ms/projects/github.com%2Fmarkuspfundstein%2Fmcp-obsidian), [Obsidian 포럼](https://forum.obsidian.md/t/obsidian-rest-api-mcp-server/109211)
