FROM python:3.12-slim

WORKDIR /app

# Install Python dependencies first (layer-cached)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY app.py .
COPY templates/ templates/

# Create upload temp directory
RUN mkdir -p /tmp/seekprice_uploads

EXPOSE 5003

CMD ["python", "app.py"]
