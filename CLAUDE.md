@AGENTS.md

# PPT 작성기 — 프로젝트 규칙서

사내 직원이 PowerPoint 안의 작업창(추가 기능)에서 템플릿을 골라 채팅으로 슬라이드를 만들고 고치는 앱이다. 만든 슬라이드는 열린 PowerPoint 파일에 바로 들어간다.
무엇을 만들지는 [PRD.md](PRD.md)가 기준이며, 이 파일은 "어떻게 만들지"의 규칙을 적는다.

## 대화 규칙
- 모든 설명과 주석은 한국어로 작성하고, 초보자 눈높이에서 쉬운 말로 설명한다.
- 코드를 바꾸면 반드시 무엇을 왜 바꿨는지 한 줄로 알려 준다.
- PRD에 없는 기능이나 다른 기술을 쓰고 싶으면 먼저 물어본다. 마음대로 추가하지 않는다.
- 이미 설치된 서브에이전트(예: 코드 분석, 설계-구현 비교, QA 에이전트)는 필요할 때마다 적극 활용한다.

## 명령어
- `npm run dev` — 개발 서버 실행 (PLAN 작업 1 이후 https://localhost:3000, 작업창은 HTTPS로만 열림)
- `npm run build` — 배포용 빌드
- `npm run lint` — 코드 검사 (ESLint)
- `npx tsc --noEmit` — 타입 검사

## 기술 스택 (고정)
- PRD에 정한 대로 Next.js 16 (App Router) + TypeScript + Tailwind CSS 4로 고정한다. 다른 프레임워크로 바꾸거나 마이그레이션을 제안하지 않는다.
- AI: OpenAI (`openai` 패키지, 모델은 `lib/openai.ts`의 `CHAT_MODEL`)
- PowerPoint 연결: PowerPoint 추가 기능(Office Add-in) + Office.js, 설정 파일은 XML 방식 `manifest.xml`
- 지원 PowerPoint: Windows용 PowerPoint 2024 이상. PowerPoint 기능은 2024에서 되는 PowerPointApi 1.5까지만 쓴다 (1.6 이상은 Microsoft 365 전용이라 쓰지 않는다).
- 슬라이드 만들기: `pptxgenjs`로 .pptx 내용(base64)을 만들어 PowerPoint에 넣는다 (`insertSlidesFromBase64`, DESIGN 설계안 A).
- 배포는 Vercel을 사용한다 (Part 5 예정, 작업창 화면과 API).
- Next.js 코드를 쓰기 전에 `node_modules/next/dist/docs/`의 해당 문서를 확인한다 (AGENTS.md 참고).

## 폴더 구조
- `manifest.xml` — 추가 기능 설정 (이름, 작업창 주소, 리본 버튼)
- `app/taskpane/page.tsx` — PowerPoint 작업창 페이지, Office.js를 불러오고 `PptChat`을 보여 줌
- `app/page.tsx` — PowerPoint 밖에서 열었을 때의 안내 페이지
- `app/PptChat.tsx` — 템플릿 선택, 채팅, 되돌리기 (작업창용 화면)
- `app/SlidePreview.tsx` — 템플릿 카드의 슬라이드 미리보기
- `lib/office.ts` — 추가 기능 슬라이드 찾기·넣기·지우기·표식, 파일 안 저장 (`renderDeck`)
- `app/api/chat/route.ts` — AI 규칙(프롬프트)을 담아 OpenAI에 요청하는 서버 API
- `lib/types.ts` — 슬라이드 데이터 모양과 한도(`LIMITS`)
- `lib/templates.ts` — 템플릿 3종(보고서형·회의형·교육형)의 색·글꼴·배치 (인치 단위)
- `lib/layout.ts` — 글자 크기 자동 맞춤과 다이어그램 배치 계산
- `lib/pptx.ts` — .pptx 내용(base64) 만들기
- `lib/sanitize.ts` — AI 응답을 형식·한도에 맞게 다듬기
- `lib/openai.ts` — OpenAI 연결 (서버 전용)

## 파일 규칙
- 새 파일은 이 프로젝트 폴더(my-app3) 안에만 만든다.
- 파일을 지워야 할 때는 바로 삭제하지 않는다. 프로젝트 폴더에 `trash-can` 폴더를 만들어 그 안으로 옮겨만 두고, 무엇을 옮겼는지 알려 준다. 실제 삭제는 사용자가 직접 확인한 뒤 한다.

## 핵심 설계 규칙
- **사람이 만든 슬라이드는 절대 건드리지 않는다**: 지우기·다시 넣기는 표식(태그 `PPTWRITER=1`)이 있는 추가 기능 슬라이드에만 한다. 새로 넣을 때는 넣기를 먼저 하고 예전 슬라이드는 그다음에 지운다.
- **한도는 한 곳에서만 정한다**: 슬라이드 15장, 글머리표 5개, 도형 8개는 `lib/types.ts`의 `LIMITS`만 고친다. 프롬프트와 `sanitize.ts`가 이 값을 함께 쓴다.
- **AI 응답은 믿지 않고 다듬는다**: AI가 돌려준 슬라이드는 반드시 `sanitizeDeck`을 거친 뒤 화면에 보낸다.
- **미리보기와 PowerPoint 슬라이드는 같은 계산을 쓴다**: 위치·글자 크기는 `computeSlideLayout` 하나로 계산하고, `SlidePreview`와 `pptx.ts`가 함께 쓴다. 배치를 바꿀 때는 둘이 어긋나지 않게 이 함수만 고친다.
- **템플릿 바꾸기에는 AI를 쓰지 않는다**: 같은 슬라이드 데이터를 새 템플릿으로 PowerPoint에 다시 넣기만 하므로 글자·순서·장수가 절대 바뀌지 않는다 (PRD must ②).
- **글이 넘치면 내용은 그대로, 크기만 줄인다**: `fitFontSize`로 글자 크기를 낮추고, 글자를 자르거나 고치지 않는다.
- **AI 규칙을 바꿀 때는 PRD 5번 "AI가 지킬 규칙"과 함께 맞춘다**: `route.ts`의 `INSTRUCTIONS`와 PRD 내용이 달라지지 않게 한다.

## 보안 규칙
- `.env`의 값은 절대 화면·로그·답변에 출력하지 않는다. 확인할 때는 "값 있음/없음"으로만 말한다.
- `.env` 등 비밀 정보 파일과 `node_modules` 폴더는 `.gitignore`에 등록된 상태를 유지하고, 절대 커밋하지 않는다.
- OpenAI 키는 서버 코드(`lib/openai.ts`, `server-only`)에서만 읽는다. 브라우저 코드(`"use client"` 파일)에서 `process.env`의 비밀 값을 쓰지 않는다.
- 외부 서비스 인증이 필요하면 토큰 값을 사용자에게 묻거나 채팅에 출력하지 않고, `.env`에 있는 값을 읽어서 사용한다. 명령에 토큰을 직접 적지 말고 환경 변수로 넘긴다.
  - Supabase를 쓸 상황이 생기면 Supabase CLI를 설치해 `.env`의 `SUPABASE_ACCESS_TOKEN`으로 작업한다.
  - Vercel 작업(배포 등)이 필요하면 Vercel CLI를 설치해 `.env`의 `VERCEL_TOKEN`으로 인증해 작업한다.
- 채팅 입력창 위의 "대외비·개인정보는 입력하지 마세요" 문구를 지우지 않는다.

## 커밋 규칙
- 커밋 메시지는 한국어로, 무엇을 왜 바꿨는지 한 줄로 적는다.
- 커밋하기 전에 `git status`와 스테이징된 파일 목록을 보고, `.env` 같은 비밀 값 파일이 들어가지 않았는지 확인한다.
- Windows PowerShell에서는 `-m` 메시지 안의 큰따옴표(")가 명령을 깨뜨리므로, 메시지를 UTF-8 파일로 저장해 `git commit -F 파일`로 커밋한다. 커밋 뒤 `git log -1`로 실제로 남았는지 확인한다.

## 코드 스타일
- 파일 import는 `@/lib/...` 별칭을 쓴다.
- 화면 문구와 AI 답변은 한국어 존댓말로, 마크다운 기호 없이 쓴다.
- 코드 주석은 한국어로, "왜 이렇게 했는지"가 필요한 곳에만 짧게 단다.
- 스타일은 Tailwind 클래스로 쓰고, 슬라이드 미리보기처럼 계산이 필요한 곳만 `style`을 쓴다.
- 한글 입력 중 Enter로 두 번 전송되지 않도록 `isComposing` 검사를 유지한다.

## 범위 밖 (만들지 않음)
PRD 6번 비범위를 따른다: 로그인·회원 관리, 별도 서버에 작업 저장, 사람이 만든 슬라이드나 기존 PPT 내용 AI 편집, 이미지 자동 생성, 여러 명 동시 편집, Microsoft 365 전용 PowerPoint 기능.

## 작업 절차 (검증 루프)
모든 작업은 아래 순서를 매번 반복(루프)한다.

1. **변경한다.**
2. **검사한다**: `npm run lint`와 `npm run build`를 실행해 오류가 없는지 확인한다.
3. **화면을 확인한다**: 화면이 바뀐 작업이면 `npm run dev`로 띄워 실제로 보이는지 확인한다. PowerPoint에 넣는 기능을 바꿨다면 PowerPoint에 사이드로드해 사람이 만든 슬라이드가 섞인 파일로 확인한다. AI 규칙을 바꿨다면 대표 요청(초안 생성, 범위 지정 변경, 한도 초과, 모호한 요청)으로 실제 응답도 확인한다.
4. **문제가 있으면** 고치고 다시 1번으로 돌아간다.
5. **통과하면 커밋한다**: git 커밋으로 그 상태를 남긴다. 커밋 메시지는 한국어로 무엇을 왜 바꿨는지 한 줄로 적는다. 검증을 통과하지 못한 상태는 커밋하지 않는다.
6. **알려 준다**: 무엇을 왜 바꿨는지 한 줄로 요약해 알려 준다.

## 작업 절차 (브랜치)
- 새 기능이나 눈에 띄는 변경을 시작할 때는, 먼저 `main`에서 작업용 브랜치를 만들고 그 브랜치에서 작업한다. 브랜치 이름은 `feature/기능이름` 형태로, 한글 없이 영문으로 짓는다. (예: `feature/diagram-edit`)
- 커밋은 그 작업 브랜치에만 쌓는다. `main`에서 직접 작업하지 않는다.
- 새 브랜치를 만들기 전에 `main`에 합치지 않은 `feature/` 브랜치가 있으면 목록을 알려 준다. 각 브랜치는 서로의 수정을 담고 있지 않고, 켜 둔 개발 서버(`npm run dev`)는 지금 브랜치의 코드만 보여 주기 때문이다. 개발 서버가 이미 3000번에 떠 있으면 새로 켜지 말고 그 서버로 확인한다 (PowerPoint 작업창 확인은 HTTPS 서버여야 한다).
- 검증을 통과해 작업이 끝나도 `main`에 바로 합치지 않는다. "합칠 준비가 됐다"고 사용자에게 알리고, 사용자가 확인한 뒤 병합해 달라고 말하면 그때 `main`에 합친다.
