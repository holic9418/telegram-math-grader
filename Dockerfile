FROM python:3.12-slim

WORKDIR /app

# 로그가 즉시 보이도록 출력 버퍼링 끄기 (Railway 로그 확인용)
ENV PYTHONUNBUFFERED=1

# 라이브러리 먼저 설치 (캐시 활용)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 나머지 코드 복사
COPY . .

# 봇 실행 (-u: 출력 버퍼링 없이 로그 바로 표시)
CMD ["python", "-u", "bot.py"]
