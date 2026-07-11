"""
가장 단순한 텔레그램 → Claude 봇.

동작 흐름:
  1) 내가 텔레그램으로 메시지를 보낸다
  2) 봇이 그 메시지를 Claude(Anthropic API)에 전달한다
  3) Claude의 답을 다시 텔레그램으로 돌려준다

봇 토큰과 API 키는 코드에 직접 넣지 않고 환경변수에서 읽어옵니다.
"""

import os

from dotenv import load_dotenv

import anthropic
from telegram import Update
from telegram.ext import (
    ApplicationBuilder,
    ContextTypes,
    MessageHandler,
    filters,
)

# ── 1. 환경변수에서 비밀 값 읽기 ────────────────────────────────
# 같은 폴더의 .env 파일을 읽어 환경변수로 불러옵니다.
load_dotenv()

# 코드에 토큰/키를 직접 쓰지 않고, 실행 환경에서 가져옵니다.
TELEGRAM_BOT_TOKEN = os.environ["TELEGRAM_BOT_TOKEN"]
# ANTHROPIC_API_KEY 는 anthropic 라이브러리가 환경변수에서 자동으로 읽습니다.

# ── 2. Claude 클라이언트 준비 ──────────────────────────────────
# 인자를 비워두면 환경변수 ANTHROPIC_API_KEY 를 자동으로 사용합니다.
claude = anthropic.Anthropic()


# ── 3. 메시지가 올 때마다 실행되는 함수 ─────────────────────────
async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    user_text = update.message.text  # 사용자가 보낸 글자

    # Claude 에게 물어보기
    response = claude.messages.create(
        model="claude-opus-4-8",
        max_tokens=1024,
        messages=[{"role": "user", "content": user_text}],
    )

    # 응답에서 텍스트만 뽑아내기
    reply = "".join(
        block.text for block in response.content if block.type == "text"
    )

    # 텔레그램으로 답장 보내기
    await update.message.reply_text(reply)


# ── 4. 봇 실행 ────────────────────────────────────────────────
def main() -> None:
    app = ApplicationBuilder().token(TELEGRAM_BOT_TOKEN).build()

    # 텍스트 메시지(명령어 제외)가 오면 handle_message 를 호출
    app.add_handler(
        MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message)
    )

    print("봇이 실행되었습니다. 텔레그램에서 메시지를 보내보세요. (종료: Ctrl+C)")
    app.run_polling()


if __name__ == "__main__":
    main()
