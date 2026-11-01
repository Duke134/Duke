# 블로그 MCP 세팅 가이드 (Phase 0~1, Duke PC에서 진행)

작성일 2026-11-01(일). 기획안 `docs/블로그_MCP_자동화_기획.md`의 결정(개인 블로그 · Duke 승인 1클릭 · 업리카 + 공개 전환만 Duke · Claude Code 예약 작업)에 맞춘 설치 순서다.
목표: 11/7(토)까지 아래 4개를 끝내고, 첫 글 1편을 "볼트 초안 = 네이버 비공개 글 = Blogger 초안"으로 올릴 수 있는 상태 만들기.

| 순서 | 할 일 | 걸리는 시간 | 결과물 |
|---|---|---|---|
| 1 | 업리카 연결 (네이버 발행) | 30분 | claude.ai 커넥터에 업리카 등록, "내 블로그 카테고리 목록 보여줘"가 답함 |
| 2 | Blogger 블로그 + Google Cloud OAuth | 40분 | `GOOGLE_CLIENT_ID` · `GOOGLE_CLIENT_SECRET` · 블로그 ID |
| 3 | NAVER API HUB 키 (검색·데이터랩) | 20분 | `NCP_APIGW_API_KEY_ID` · `NCP_APIGW_API_KEY` |
| 4 | Claude Code에 MCP 2개 등록 | 10분 | `.mcp.json`에 blogger · naver-search 추가, 핸드셰이크 확인 |
| (Phase 3) | Search Console · 검색광고 API | 각 20분 | 통계·검색량 실측용. 지금은 안 해도 됨 |

키·토큰은 저장소에 올리지 않는다. 아래 설정의 `${...}` 자리는 환경변수로 넣는다.

---

## 1. 업리카(uplika) — 네이버 블로그 발행

출처: Duke가 공유한 Notion 가이드 「클로드로 네이버 블로그 자동 발행하기」. 이 세션에서는 uplika.com 접속이 막혀 화면·약관·요금을 직접 확인하지 못했다. `[미확인]` 가입 전 약관과 개인정보 처리방침을 한 번 읽는다(글 본문·사진이 업리카 서버를 거친다).

컴퓨터 크롬에서 진행한다(폰 불가).

1. uplika.com → 오른쪽 위 로그인 → 구글 계정으로 가입
2. 대시보드 왼쪽 **채널 연결(Connections)** → 네이버 블로그 → 크롬 웹스토어에서 **uplika 확장** 설치
3. 확장을 설치한 **그 크롬**에서 네이버에 로그인(이미 로그인되어 있으면 그대로)
4. 대시보드로 돌아와 내 블로그 이름·카테고리 목록이 보이면 연결 완료
5. 대시보드 오른쪽 위 **에이전트 연결** → Claude → **Claude에 추가하기** → 클로드 창에서 추가 → 브라우저 승인 화면에서 허용(네이버 블로그 채널 체크 확인)
   - 버튼이 안 열리면: claude.ai **맞춤 설정 → 커넥터 → 사용자 정의 커넥터 추가**에 아래 주소를 붙여 넣는다

```txt
https://api.uplika.com/mcp
```

6. 연결 확인: 클로드에 "내 블로그 카테고리 목록 보여줘" → 카테고리가 나오면 완료. 이후 폰 Claude 앱에서도 쓸 수 있다.

운영 규칙(기획안 §6):
- 글은 **항상 비공개로 먼저** 올리고 글 주소를 받는다. 공개 전환은 Duke가 "공개로 바꿔줘"라고 할 때만
- 글을 올릴 때는 확장을 설치한 크롬이 켜져 있어야 한다 → 예약 작업은 Duke가 PC 앞에 있는 시간대로
- 베타 무료, 정식 출시 후 월 20편 무료(가이드 기준) `[미확인]` → 주 3편이면 월 13편으로 무료 구간 안

---

## 2. Blogger + Google Cloud OAuth (`mcp-blogger`)

### 2-1. 블로그 개설
1. blogger.com → 새 블로그 만들기(API로는 블로그 생성 불가)
2. 블로그 ID 확인: Blogger 관리 화면 주소 `https://www.blogger.com/blog/posts/<숫자>` 의 숫자. 이 값이 `BLOGGER_DEFAULT_BLOG_ID`

### 2-2. Google Cloud OAuth 클라이언트
`mcp-blogger` README 기준(npm 1.0.2). 리디렉션 주소가 다르면 동의 화면에서 실패하므로 그대로 쓴다.

1. console.cloud.google.com → 새 프로젝트(예: `duke-blog`)
2. **API 및 서비스 → 라이브러리** → "Blogger API v3" 사용 설정
3. **OAuth 동의 화면** → 사용자 유형 **외부** → 앱 이름·이메일 입력 → 범위는 기본 → **테스트 사용자**에 Duke 구글 계정 추가(게시 안 하고 테스트 상태로 둔다)
4. **사용자 인증 정보 → 사용자 인증 정보 만들기 → OAuth 클라이언트 ID** → 유형 **웹 애플리케이션** → 승인된 리디렉션 URI에 정확히:

```txt
http://localhost:3000/oauth/callback
```

5. 생성된 **클라이언트 ID**·**클라이언트 보안 비밀**을 환경변수 `GOOGLE_CLIENT_ID`, `GOOGLE_CLIENT_SECRET`으로 보관
6. 첫 쓰기 호출(초안 생성) 때 `mcp-blogger`가 브라우저를 열어 동의를 받고, 토큰을 `~/.config/mcp-blogger/tokens.json`에 저장·자동 갱신한다. 5분 안에 동의하지 않으면 타임아웃

주의: 테스트 상태 OAuth 앱의 refresh token은 7일 뒤 만료될 수 있다(Google 정책, `[미확인]`). 만료되면 쓰기 호출 때 다시 동의 창이 뜬다. 매번 뜨면 동의 화면을 "프로덕션"으로 전환하는 것을 검토한다.

### 2-3. 도구 확인(README 기준)
| 도구 | OAuth | 용도 |
|---|---|---|
| `create_post` | 필요 | **기본 초안**으로 생성. 큰 본문은 `content_file` |
| `list_drafts` · `get_post` | 필요 | 미리보기 |
| `change_post_status action=publish` | 필요 | 발행 (Duke 승인 후) |
| `change_post_status action=revert` | 필요 | 발행 취소 → 초안 |
| `update_post` | 필요 | 수정 |
| `list_posts` | 불필요(API 키) | 발행 글 목록 |

---

## 3. NAVER API HUB — 검색·데이터랩 (`@isnow890/naver-search-mcp`)

네이버 개발자센터(openapi.naver.com) 키는 **2026-07-31부터 신규 발급이 막혔다**(기존 키는 2027-06-30까지 동작). 신규는 네이버클라우드 콘솔에서 받는다. 출처: `@isnow890/naver-search-mcp` README(npm 1.0.54, 2026-10-07).

1. ncloud.com 콘솔 로그인(네이버클라우드 계정, 결제수단 등록이 필요할 수 있음 `[확인 필요]`)
2. Menu → All Services → Application Services → **NAVER API HUB** → Application 등록 → 사용할 API에 **검색**(블로그·웹문서·뉴스), **데이터랩**(검색어트렌드) 선택
3. 등록 후 **인증 정보**에서 `NCP_APIGW_API_KEY_ID`, `NCP_APIGW_API_KEY` 확보
4. 쓰게 될 도구: `search_blog`(경쟁 글 파악), `search_webkr`, `search_news`, `datalab_search`(키워드 트렌드, 날짜에 `"today"` 가능)

이미 개발자센터 키가 있으면 `NAVER_CLIENT_ID` / `NAVER_CLIENT_SECRET` 한 쌍만 넣어도 된다(둘 중 한 쌍만).

---

## 4. Claude Code에 MCP 등록

이 저장소의 `.mcp.json`에는 kordoc·anymd만 들어 있다. 아래 두 서버를 **Duke가 직접** 추가한다(Claude가 외부 MCP를 설정에 넣는 작업은 자동 모드에서 차단되어 있고, 외부 코드 등록은 Duke가 하는 것이 규칙에도 맞다).

### 4-1. `.mcp.json`의 `mcpServers`에 추가

```json
"blogger": {
  "type": "stdio",
  "command": "npx",
  "args": ["-y", "mcp-blogger"],
  "env": {
    "GOOGLE_CLIENT_ID": "${GOOGLE_CLIENT_ID}",
    "GOOGLE_CLIENT_SECRET": "${GOOGLE_CLIENT_SECRET}",
    "DEFAULT_BLOG_ID": "${BLOGGER_DEFAULT_BLOG_ID}"
  }
},
"naver-search": {
  "type": "stdio",
  "command": "npx",
  "args": ["-y", "@isnow890/naver-search-mcp"],
  "env": {
    "NCP_APIGW_API_KEY_ID": "${NCP_APIGW_API_KEY_ID}",
    "NCP_APIGW_API_KEY": "${NCP_APIGW_API_KEY}"
  }
}
```

### 4-2. `.claude/settings.json`의 `enabledMcpjsonServers`에 이름 추가

```json
"enabledMcpjsonServers": ["kordoc", "anymd", "blogger", "naver-search"]
```

### 4-3. 환경변수 (저장소 밖)
Windows PowerShell 예시(사용자 환경변수로 영구 등록):

```powershell
[Environment]::SetEnvironmentVariable("GOOGLE_CLIENT_ID", "<값>", "User")
[Environment]::SetEnvironmentVariable("GOOGLE_CLIENT_SECRET", "<값>", "User")
[Environment]::SetEnvironmentVariable("BLOGGER_DEFAULT_BLOG_ID", "<값>", "User")
[Environment]::SetEnvironmentVariable("NCP_APIGW_API_KEY_ID", "<값>", "User")
[Environment]::SetEnvironmentVariable("NCP_APIGW_API_KEY", "<값>", "User")
```

Windows에서 Claude Desktop이 npx를 못 찾으면 `"command": "cmd", "args": ["/c", "npx", "-y", "mcp-blogger"]` 형태로 바꾼다(기존 kordoc 안내와 같은 방식).

### 4-4. 확인
Claude Code를 이 폴더에서 다시 열면 `.mcp.json` 서버 승인 프롬프트가 뜬다. 승인 후:

```bash
claude mcp list
```

blogger · naver-search가 connected로 보이면 끝. 대화에서 "내 Blogger 블로그 초안 목록 보여줘", "'경영지원 연말정산' 네이버 블로그 검색 결과 5개"로 각각 한 번씩 호출해 본다.

---

## 5. 첫 글 1편 테스트 (Phase 1 시작 신호)

1~4가 끝나면 Claude에 이렇게 보낸다.

```txt
blog-write로 [키워드] 글 한 편 써줘. 한 번에 써줘.
저장은 볼트 초안 + Blogger 초안(create_post) + 업리카 비공개 업로드까지. 공개는 하지 마.
```

기대 결과: 볼트 `80_블로그아이디어/초안/` 파일 1개, Blogger 초안 1개(`list_drafts`로 확인), 네이버 비공개 글 주소 1개. Duke가 폰에서 보고 "공개로 바꿔줘" / "publish 해줘"라고 하면 발행.

---

## (Phase 3에서) Search Console · 검색광고 API

- Search Console: search.google.com/search-console에 Blogger 주소 등록 → Google Cloud에서 서비스 계정 생성 → Search Console 속성에 그 서비스 계정 이메일을 사용자로 추가 → `mcp-server-gsc`(npm 0.3.0) 또는 `@sarahpark/google-search-console-mcp`(npm 1.1.0)에 `GOOGLE_APPLICATION_CREDENTIALS` 경로 지정
- 네이버 검색광고: searchad.naver.com 광고시스템 → 도구 → **API 사용 관리**에서 라이선스(무료) 발급 → `alphago2580/naver_searchad_mcp`(Python) 또는 `2duckchun/naver-ads-mcp`(`npx -y github:2duckchun/naver-ads-mcp`). 둘 다 npm·PyPI 미등록이라 GitHub에서 받는다

---

## 확인 못 한 것 (이 세션 네트워크 정책)
- uplika.com · api.uplika.com 접속 불가 → 화면·약관·요금 `[미확인]`
- github.com 접속 불가 → MCP 저장소 README는 npm 레지스트리 메타데이터로 대신 확인(mcp-blogger 1.0.2, naver-search-mcp 1.0.54)
- developers.google.com 접속 불가 → Blogger API 할당량·OAuth 테스트 앱 토큰 만료 정책 `[미확인]`
