FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN mkdir -p uploads results

ENV ENV=production
ENV PORT=3333

EXPOSE 3333

CMD ["python", "app.py"]
