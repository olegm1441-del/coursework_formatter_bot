# Railway: быстрый split на Web + Worker

## Почему деплой может идти ~4 минуты
- На каждом деплое Railway пересобирает окружение и ставит зависимости из `requirements.txt`.
- Если web и worker в одном сервисе/старте, при любом патче перезапускается всё сразу.
- Это не связано с параллельной обработкой `save formatter` в рантайме: время деплоя в основном уходит на build + boot контейнера.

## Минимальная схема
- `web` процесс: `python bot.py`
- `worker` процесс: `python worker.py`

## LibreOffice для форматирования таблиц
- `worker` должен иметь доступ к `soffice`, потому что rendered table continuation
  строит PDF через LibreOffice.
- Новый Railway service не нужен: текущий `worker` остаётся самодостаточным.
- Railway сейчас собирает сервис через Railpack, поэтому `nixpacks.toml` не является
  рабочим способом поставить LibreOffice для этого окружения.
- Для `worker` используется отдельный Dockerfile: `Dockerfile.worker`.
- Root `Dockerfile` не добавляется, чтобы `web` продолжал собираться Railpack.

## Настройка Railway UI
1. Открой существующий `worker` service.
2. В Build settings включи Dockerfile-based build только для `worker`.
3. Укажи custom Dockerfile path:
   `Dockerfile.worker`
4. Если Railway использует переменную вместо UI-поля, добавь только в `worker`:
   `RAILWAY_DOCKERFILE_PATH=Dockerfile.worker`
5. Start command для `worker` оставь прежним:
   `python worker.py`
   или используй `CMD` из `Dockerfile.worker`, если Railway не переопределяет command.
6. `web` service не меняй:
   - builder остаётся Railpack
   - start command остаётся `python bot.py`
   - custom Dockerfile path для `web` не задавать
7. Не создавай новый service и не добавляй sidecar.

## Диагностика worker
При старте worker в логах должна быть одна из строк:
- `worker_libreoffice_available binary=... version=...`
- `worker_libreoffice_missing soffice_not_found_in_path`

Если видна строка `worker_libreoffice_missing`, rendered table continuation будет
пропущен, а форматтер вернёт warning:
`rendered table continuation skipped: LibreOffice (soffice) not found`.
