FROM python:3.12-alpine

WORKDIR /opt/src/

RUN apk update && apk add --no-cache gcc musl-dev libffi-dev

RUN pip install --upgrade pip

RUN pip install poetry
RUN poetry config virtualenvs.create false
COPY poetry.lock pyproject.toml ./
RUN poetry install --no-root

COPY ./ ./
CMD python start_app.py

