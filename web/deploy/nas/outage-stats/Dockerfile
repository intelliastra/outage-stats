FROM python:3.11-slim

WORKDIR /app/backend

COPY web/backend/requirements.txt /tmp/requirements.txt
RUN pip install --no-cache-dir -r /tmp/requirements.txt

ENV STATS_BASE=/data/统计材料
ENV PYTHONUNBUFFERED=1
ENV PYTHONIOENCODING=utf-8

EXPOSE 22229

CMD ["python", "-m", "uvicorn", "main:app", "--host", "0.0.0.0", "--port", "22229"]
