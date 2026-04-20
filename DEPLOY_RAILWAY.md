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
- LibreOffice ставится в Railway build через `nixpacks.toml`.
- Новый Railway service не нужен: текущий `worker` остаётся самодостаточным.

## Диагностика worker
При старте worker в логах должна быть одна из строк:
- `worker_libreoffice_available binary=... version=...`
- `worker_libreoffice_missing soffice_not_found_in_path`

Если видна строка `worker_libreoffice_missing`, rendered table continuation будет
пропущен, а форматтер вернёт warning:
`rendered table continuation skipped: LibreOffice (soffice) not found`.
