# Microsoft Generate/Analyze

## Авторизация

Для доступа к методам API требуется использовать Bearer токен. Добавьте токен в заголовки запроса следующим образом:

`Authorization: Bearer YOUR_TOKEN_HERE`

## Instruction

**Без docker-compose:**

- создать и заполнить .env файл
- установить зависимости `pip install -r requirements.txt`
- запустить сервер `python start_app.py`
- посмотреть [Swagger](http://localhost:8000/docs)

**С docker-compose:**

- создать и заполнить .env файл
- запустить контейнер `docker-compose up --build -d`
- посмотреть [Swagger](http://localhost:8000/docs)
