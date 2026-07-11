# 텔레그램 ↔ Claude 봇 (초보자용)

내가 텔레그램으로 보낸 메시지를 Claude(Anthropic API)에 전달하고,
그 답을 다시 텔레그램으로 돌려주는 가장 단순한 봇입니다.

---

## 준비물 2가지

봇을 켜려면 아래 2개의 "비밀 값"이 필요합니다. 코드에는 넣지 않고,
`.env` 파일에만 적어둡니다.

| 이름 | 어디서 받나요? |
|------|----------------|
| **텔레그램 봇 토큰** | 텔레그램의 `@BotFather` |
| **Anthropic API 키** | https://console.anthropic.com |

---

## 1단계 — 텔레그램 봇 토큰 받기

1. 텔레그램 앱에서 검색창에 `@BotFather` 를 검색해 대화를 엽니다.
2. `/newbot` 이라고 보냅니다.
3. 봇 이름과 사용자명(`~bot`으로 끝나야 함)을 정합니다.
4. BotFather가 `123456:ABC-DEF...` 형태의 **토큰**을 줍니다. 이 값을 복사해 두세요.

## 2단계 — Anthropic API 키 받기

1. https://console.anthropic.com 에 로그인합니다.
2. **API Keys** 메뉴에서 새 키를 만듭니다. (`sk-ant-...` 로 시작)
3. 키는 만들 때 한 번만 보이니 바로 복사해 두세요.

> 참고: API 사용에는 결제 정보 등록/크레딧이 필요할 수 있습니다.

## 3단계 — 비밀 값을 `.env` 파일에 넣기

1. 이 폴더의 `.env.example` 파일을 복사해서 이름을 `.env` 로 바꿉니다.
2. 파일을 열어 아래처럼 값을 채웁니다:

   ```
   TELEGRAM_BOT_TOKEN=123456:ABC-DEF1234567890
   ANTHROPIC_API_KEY=sk-ant-여기에_실제_키
   ```

3. 저장합니다. (`.env` 파일은 절대 다른 사람에게 공유하지 마세요.)

## 4단계 — 필요한 프로그램 설치

이 폴더에서 PowerShell 창을 열고 아래를 실행합니다.
(Python이 아직 없다면 https://www.python.org 에서 먼저 설치하세요.)

```powershell
pip install -r requirements.txt
```

## 5단계 — 봇 실행하기

```powershell
python bot.py
```

`봇이 실행되었습니다...` 라는 문구가 나오면 성공입니다.
이제 텔레그램에서 **내가 만든 봇**을 찾아 아무 메시지나 보내보세요.
잠시 후 Claude의 답이 돌아옵니다.

봇을 끄려면 PowerShell 창에서 `Ctrl + C` 를 누릅니다.

---

## 자주 겪는 문제

- **`KeyError: 'TELEGRAM_BOT_TOKEN'`**
  → `.env` 파일이 없거나 이름이 틀렸습니다. 파일 이름이 정확히 `.env` 인지 확인하세요.

- **`authentication_error` / 401**
  → `ANTHROPIC_API_KEY` 값이 틀렸습니다. 키를 다시 복사해 넣으세요.

- **봇이 응답이 없어요**
  → `python bot.py` 창이 켜져 있어야 합니다. 창을 닫으면 봇도 꺼집니다.

---

## 파일 설명

| 파일 | 역할 |
|------|------|
| `bot.py` | 봇의 실제 코드 |
| `requirements.txt` | 설치할 라이브러리 목록 |
| `.env.example` | 비밀 값을 적는 예시 파일 (복사해서 `.env` 로) |
| `.env` | 내 실제 비밀 값 (직접 만듦, 공유 금지) |
