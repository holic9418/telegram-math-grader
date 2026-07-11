# 봇 24시간 돌리기 (Railway 클라우드 배포)

내 PC를 꺼도 봇이 항상 켜져 있게 만드는 방법입니다.
크게 **① 코드를 GitHub에 올리기 → ② Railway에 연결하기** 두 단계예요.

> 비용 안내: Railway는 처음에 무료 크레딧($5)을 주고, 다 쓰면 월 약 $5(약 7천 원)입니다.

---

## ① 코드를 GitHub에 올리기 (GitHub Desktop 사용 — 가장 쉬움)

1. **GitHub 계정 만들기**: https://github.com/signup 에서 가입
2. **GitHub Desktop 설치**: https://desktop.github.com → 다운로드 후 설치
3. GitHub Desktop을 열고 방금 만든 계정으로 **로그인**
4. 상단 메뉴 `File` → `Add Local Repository`
5. 폴더 선택 창에서 이 폴더를 고릅니다:
   `C:\Users\holic\Desktop\claude`
6. 오른쪽/가운데의 **`Publish repository`** 버튼 클릭
7. 창이 뜨면:
   - **`Keep this code private`** 에 반드시 **체크** (비공개로 올리기)
   - `Publish repository` 클릭
8. 잠시 후 코드가 GitHub에 올라갑니다. ✅

> 안심하세요: 비밀 파일(`.env`, `API key.txt`)은 올라가지 않도록 이미 막아뒀습니다.

---

## ② Railway에 연결해서 24시간 켜기

1. https://railway.app 접속 → **`Login with GitHub`** (GitHub 계정으로 로그인)
2. `New Project` 클릭 → **`Deploy from GitHub repo`** 선택
3. 권한 허용 창이 뜨면 허용하고, 방금 올린 저장소(`claude`)를 선택
4. 배포가 자동으로 시작돼요. (아직 비밀 값이 없어서 한 번 실패해도 정상)
5. 프로젝트 화면에서 **`Variables`** 탭 클릭 → 아래 두 개를 추가:

   | 이름(Name) | 값(Value) |
   |------------|-----------|
   | `TELEGRAM_BOT_TOKEN` | 내 봇 토큰 |
   | `ANTHROPIC_API_KEY` | 내 API 키 |

   (값은 `.env` 파일에 넣었던 것과 똑같이 복사해 넣으면 됩니다)
6. 저장하면 Railway가 자동으로 다시 배포하고, 봇이 **24시간** 돌아갑니다. 🎉

---

## ⚠️ 꼭 지킬 것: 봇은 한 번에 하나만!

같은 봇을 동시에 두 곳(내 PC + 클라우드)에서 돌리면 서로 충돌해요.
- 클라우드(Railway)가 켜지면, 내 PC에서 돌리던 봇은 꺼주세요.
- (이 채팅 세션에서 돌던 봇은 세션이 끝나면 자동으로 꺼집니다.)

---

## 잘 됐는지 확인

Railway 프로젝트의 **`Deployments`** → 로그(Logs)에
`봇이 실행되었습니다...` 가 보이면 성공입니다.
텔레그램에서 `@Taei_lobot` 에게 메시지를 보내 확인해보세요.

문제가 생기면 Railway 로그의 빨간 글씨를 복사해서 물어봐 주세요.
