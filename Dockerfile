FROM python:3.12-slim

WORKDIR /app

# 라이브러리 먼저 설치 (캐시 활용)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 나머지 코드 복사
COPY . .

# 봇 실행
CMD ["python", "bot.py"]
