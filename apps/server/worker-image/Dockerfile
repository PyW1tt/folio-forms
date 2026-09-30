FROM python:3.12.9-slim-bookworm

RUN apt-get update && apt-get install --no-install-recommends -y libseccomp2 \
  && rm -rf /var/lib/apt/lists/* \
  && pip install --no-cache-dir python-docx==1.2.0
RUN mkdir /input
COPY runner.py /runner.py
ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1
USER 65532:65532
ENTRYPOINT ["python", "/runner.py"]
