---
title: GitHub 경영지원 활용자료 조사
created: 2026-11-01
updated: 2026-11-01
category: 업무레퍼런스
tags: [도구, MCP, 스킬, PPT, 옵시디언, 조사]
status: 채택 3개 등록 완료(설치 대기)
source: GitHub API 2026-11-01 기준(별·갱신일·라이선스), 기능은 README 기준
---

# GitHub 경영지원 활용자료 조사

> 조사 목적: 타인 공개 저장소 중 경영지원 6카테고리(재무·회계 / 인사·노무 / IR·투자 / 정부사업 / 해외업무 / 총무)에 쓸 만한 것. 축 3개: MCP 자동화 / PPT 디자인 / 개인 자료 활용.
> 결과: 40여 개 검토 → 1차 후보 8개 → 채택 3개(kordoc · ppt-master · anymd). Duke 저장소 main에 등록 완료, PC 설치는 대기.

## 1. 결론 — 채택 3개

| 도구 | 역할 | 왜 채택 | 상태 |
|---|---|---|---|
| chrisryugj/kordoc 4.24.35 (MIT) | HWP·HWPX·PDF·Office → Markdown, 양식 자동 채우기, 신구대조, MCP 22도구 | 현재 환경에 전혀 없는 기능. 정부사업·관공서 서류가 HWP | 설치 대기 |
| hugohe3/ppt-master 6.8.0 (MIT) | 문서 → 네이티브 편집 가능 PPTX, 자체 템플릿·슬라이드 마스터 사용 | 기존 pptx 스킬보다 디자인·템플릿 준수 우위. 내부 문서에 맑은 고딕(Malgun Gothic)·바탕을 PPT 안전 글꼴로 명시 | 설치 대기 |
| SylphxAI/anymd 8.5.7 (MIT) | PDF·Word·PPT·Excel·이미지(OCR) → Markdown, MCP 4도구 | 받은 자료를 옵시디언에 넣기 전 전처리. kordoc=HWP, anymd=그 외 분담 | 설치 대기 |

등록 위치: Duke 저장소(Duke134/Duke) main
- `.mcp.json` — kordoc·anymd (접근 범위 프로젝트 폴더로 제한)
- `.claude/settings.json` — ppt-master·kordoc 마켓플레이스, ppt-master 활성화
- `docs/도구등록_kordoc_ppt-master_anymd.md` — 설치 절차·검증 기록

PC에서 할 일
1. main 받아 Claude Code로 열기 → 폴더 신뢰 + MCP 승인
2. ppt-master 설치 제안 수락 → 설치 폴더에서 `pip install -r requirements.txt`
3. ppt-master 첫 사용 시 "글꼴 Malgun Gothic 고정"으로 샘플 1장 → 줄바꿈·겹침 확인

## 2. 1차 후보 8개 선별 결과

| 저장소 | 판정 | 이유 |
|---|---|---|
| chrisryugj/kordoc | 채택 | 위 참조 |
| hugohe3/ppt-master | 채택 | 위 참조 |
| SylphxAI/anymd | 채택(보조) | 위 참조. 색인 목록(awesome-*)은 설치 대상 아님, 북마크용 |
| MarkusPfundstein/mcp-obsidian + coddingtonbear/obsidian-local-rest-api | 보류 | 현재 obsidian-rest 구성의 원본. 신규 설치 없음, 버전 확인만 |
| haris-musa/excel-mcp-server | 보류 | xlsx 스킬과 중복. 피벗·슬라이서 필요 시 |
| taylorwilsdon/google_workspace_mcp | 제외 | Gmail·Calendar·Drive·Sheets·Docs 커넥터 이미 보유 |
| GongRzhe/Office-PowerPoint-MCP-Server | 제외 | 2025-12 이후 미갱신, ppt-master와 중복 |
| korotovsky/slack-mcp-server | 제외 | Slack 커넥터 보유. 브라우저 세션 토큰 방식으로 알고 있어 보안정책 우려 `[미확인]` |

## 3. 전체 조사표

### 3-1. 한국 실무 특화 (법령·세무)
| 저장소 | ⭐ | 갱신 | 라이선스 | 요지 |
|---|---|---|---|---|
| chrisryugj/korean-law-mcp | 2,823 | 10-31 | MIT | 법제처 42 API. verify_citations(인용 실존 검증), cite_check(판례 폐기 확인). 현재 korean-law 커넥터와 동일물 가능성 → 중복 확인 |
| UpstageAI/korean-tax-mcp | 5 | 10-30 | MIT | 24도구. 국세청 해석·판례·통칙·집행기준·조세조약. 기본 키 불필요. UPSTAGE_API_KEY 설정 시 입력이 외부 전송 |
| UpstageAI/korean-tax-calc-mcp | 94 | 10-30 | MIT | 원장 → 세무조정(소득처분), 근거 조문·산식, 법정서식 |
| sungjunlee/beopsuny-skill | 3 | 10-17 | MIT | 법무 실무 스킬(법령·판례·계약서 검토). korean-law-mcp 위에서 동작. contract-review 보강 참고 |
| lbiz-partners/hometax-doum | 79 | 10-31 | 자체(비상업) | 홈택스 부가세·종소세·전자세금계산서 준비 스킬 6종. 제출 전 멈춤. 법인세 미포함 |

### 3-2. 정부사업·공공데이터
| 저장소 | ⭐ | 갱신 | 라이선스 | 요지 |
|---|---|---|---|---|
| djfksjd/ir-search | 393 | 07-30 | MIT | K-Startup·기업마당·NIPA·KOCCA·SMTECH 전수조사 → 3그룹 분류. [[GXC_지원사업_프로필]]과 결합 시 효과. 3개월 미갱신 주의 |
| leokim90/gov-data-mcp | 3 | 09-17 | MIT | 중기부·정부24·기업마당·나라장터·중소벤처24 6도구. data.go.kr 무료 키 |
| hjsh200219/korea-public-data-mcp | 11 | 10-29 | 미표기 | 법제처 21 + DART 5 + 공공데이터포털 8(사업자 진위확인). 원격 MCP 제공 |
| chrisryugj/korean-dart-mcp | 118 | 10-26 | MIT | OpenDART 15도구. 투자사·경쟁사·상장 파트너 조사용 |

### 3-3. 한국 문서(HWP)
| 저장소 | ⭐ | 갱신 | 라이선스 | 요지 |
|---|---|---|---|---|
| chrisryugj/kordoc | 2,881 | 11-01 | MIT | 채택. 표준 기안문·도장 날인·신구대조 포함 |
| treesoop/hwp-mcp | 103 | 11-01 | MIT | HWP/HWPX 읽기·쓰기 34도구. 한글 파일 직접 작성 시 |
| jkf87/hwp-mcp | 271 | 09-27 | 미표기 | 초기 프로젝트. 위 둘이 더 활발 |

### 3-4. MCP 자동화(범용)
| 저장소 | ⭐ | 갱신 | 라이선스 | 요지 | 현재 환경 |
|---|---|---|---|---|---|
| taylorwilsdon/google_workspace_mcp | 3,452 | 10-30 | MIT | Google 전체 통합 | 커넥터 보유 → 제외 |
| haris-musa/excel-mcp-server | 4,269 | 10-31 | MIT | xlsx 생성·편집, 피벗·슬라이서 | 보류 |
| GongRzhe/Office-PowerPoint-MCP-Server | 1,861 | 2025-12-31 | MIT | python-pptx MCP | 제외 |
| korotovsky/slack-mcp-server | 1,882 | 07-16 | MIT | Slack 읽기·검색·DM | 제외 |
| MarkusPfundstein/mcp-obsidian · obsidian-local-rest-api | 4,514 · 3,100 | 08-31 · 10-17 | MIT | Obsidian REST+MCP | 현재 구성 원본 |
| SylphxAI/anymd | 1,059 | 11-01 | MIT | 범용 → Markdown | 채택 |
| yctimlin/mcp_excalidraw | 2,586 | 10-30 | MIT | Excalidraw 도식 | 조직도·프로세스용 후보 |
| enescingoz/awesome-n8n-templates | 26,210 | 09-12 | 미표기 | n8n 워크플로 280+ | 서버 운영 필요, 후순위 |

색인(북마크): punkpeye/awesome-mcp-servers ⭐97,143 · hesreallyhim/awesome-claude-code ⭐57,006 · ComposioHQ/awesome-claude-skills ⭐78,467

### 3-5. PPT 디자인
| 저장소 | ⭐ | 갱신 | 라이선스 | 요지 | 판단 |
|---|---|---|---|---|---|
| hugohe3/ppt-master | 71,791 | 10-30 | MIT(PDF부 PyMuPDF AGPL) | 네이티브 PPTX, 템플릿·마스터, 기존 덱 채우기 | 채택 |
| alchaincyf/huashu-design | 28,337 | 10-30 | MIT | HTML 네이티브 디자인, 스타일 92종, 한국어 README | 스타일 탐색·인포그래픽용 |
| chuspeeism/dashi-ppt-skill | 10,226 | 09-12 | AGPL-3.0(내보내기 엔진 독점) | 브라우저 편집 덱 12테마 | 외부 배포 시 라이선스 주의 |
| ferdinandobons/brand-docs | 290 | 07-10 | MIT | 기존 docx·pptx·xlsx 템플릿 1개 → Brand Profile → 같은 서식 생성 | GXC 공문 고정 서식용 후보. 알파 |
| addsumtech/slides_maker | 645 | 10-31 | MIT | 문서 → 네이티브 차트 PPTX | ppt-master 대안 |
| marp-team/marp (+ favourhong/Awesome-Marp) | 12,709 | 07-29 | MIT | Markdown → 슬라이드 | 디자인 자유도 낮음 |
| Love-Ash/archforge | 6 | 10-07 | MIT | AI PPT 사전 점검(CJK 폰트 누락·겹침) | 보조 |

### 3-6. 개인 자료 활용(옵시디언·지식관리·보고)
| 저장소 | ⭐ | 갱신 | 라이선스 | 요지 | 판단 |
|---|---|---|---|---|---|
| YishenTu/claudian | 15,876 | 11-01 | MIT | Obsidian 안에 Claude Code 임베드 | obsidian-rest 대안 |
| AgriciDaniel/claude-obsidian | 16,039 | 10-25 | MIT | inbox 투입형 자기정리 vault, 기존 vault 채택 가능, PARA 모드 | 00_Inbox 규칙과 호환 가능성. 폴더 체계 충돌 확인 필요 |
| eugeniughelbur/obsidian-second-brain | 5,085 | 10-28 | MIT | 에이전트 기억을 vault markdown에 저장 | 개인정보 규칙 점검 후 |
| axtonliu/axton-obsidian-visual-skills | 3,673 | 06-26 | MIT | 텍스트 → Canvas·Excalidraw·Mermaid | 회고·SOP 시각화 |
| tt-a1i/archify | 92,206 | 10-31 | Apache-2.0 | 계획 → 인터랙티브 도식 | 보고용 한 장 도식 |
| Pitchdrive/investor-update-skill | 4 | 07-22 | MIT | 투자자 업데이트 8섹션, 숫자 발명 금지 규칙 | IR 월간 업데이트 틀(미국식) |
| 1984vc/cap-table · Open-Cap-Format-OCF | 142 · 194 | 09-25 · 10-26 | MIT · 미표기 | 캡테이블 워크시트 · 데이터 표준 | 라운드 시뮬레이션 · 미국법인 주식 정리 참고 |
| sara-hamza/people-ops-claude | 0 | 10-20 | MIT | HR 스킬 25종(온보딩·오프보딩·급여 예외) | 신규(별 0). 인턴 온보딩 틀 참고 |
| anthropics/knowledge-work-plugins | 34,248 | 11-01 | Apache-2.0 | 설치된 finance·HR·legal·operations 원본 | 업데이트 추적. 미국 기준 |
| alirezarezvani/claude-skills | 29,087 | 10-31 | MIT | 380+ 스킬 | 품질 편차, 골라 참고 |

## 4. 제외한 것
- 한국 노무 계산(연차·퇴직금·4대보험): 전부 ⭐0~1 개인 프로젝트. 기존 Duke 스킬 + korean-law 유지
- 미국 세무(Form 5472·델라웨어 프랜차이즈세) 전용: 없음. eCFR MCP는 ⭐1~3 → 현재 ecfr 커넥터 유지
- 홈택스 매크로·스크래핑류: 자동 로그인·제출 포함 → 외부 제출 게이트와 충돌
- duyunzhoudyz/gongwen-skills: 중국 국가표준 공문 서식

## 5. 도입 시 주의
- MCP에 넣는 데이터의 외부 전송 여부 확인(korean-tax-mcp의 UPSTAGE_API_KEY)
- 계약서 전문·급여·주주 정보는 로컬 실행 도구에만
- 라이선스 미표기 저장소는 외부 배포 전 확인
- 스킬·MCP 추가는 Duke 확인 후 "설치 대기" → 설치 완료 시 이 노트 status 갱신

## 6. 다음 액션
- [ ] PC에서 kordoc·anymd·ppt-master 설치 → status "설치 완료"로 갱신
- [ ] ppt-master 맑은 고딕 샘플 1장 검증
- [ ] 2차 검토 후보: djfksjd/ir-search, UpstageAI/korean-tax-mcp, ferdinandobons/brand-docs
