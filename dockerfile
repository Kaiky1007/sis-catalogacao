FROM python:3.9-slim

WORKDIR /projeto

RUN apt-get update && apt-get install -y \
    libpq-dev \
    gcc \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

ENV FLASK_APP=app
ENV FLASK_DEBUG=1
ENV PYTHONUNBUFFERED=1

CMD ["flask", "run", "--host=0.0.0.0"]
