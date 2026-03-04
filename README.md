# Выгрузка счетов и позиций Т‑Инвестиций в Google Sheets через Apps Script

Решение полностью без Python: код запускается прямо в Google Таблице через Google Apps Script.

Скрипт получает из Invest API 2.0:
- список счетов;
- валютные позиции через `GetPositions`;
- рыночные позиции через `GetPortfolio`;
- метаданные фьючерсов через `InstrumentsService/Futures`;
- полное имя инструмента через `InstrumentsService/GetInstrumentBy`.

И записывает их в листы:
- `accounts`
- `positions`

## Особенности

- В таблице есть `instrument_name` (краткое) и `instrument_full_name` (полное наименование).
- Добавлен столбец `account_name`, чтобы фильтрация/сводные по счетам были нагляднее.
- Колонки `source`, `figi`, `lot`, `min_price_increment` автоматически скрываются после выгрузки (данные остаются в таблице).
- Для числовых колонок выполняется явное приведение к числу и числовой формат, чтобы сортировка работала корректно.
- Добавляется оформление таблицы: жирный заголовок, фильтры по заголовкам, чередование цветов строк (белый/серый).
- Для фьючерсов:
  - `min_price_increment` = минимальный шаг цены;
  - `point_price` = стоимость 1 пункта цены = `minPriceIncrementAmount / minPriceIncrement`.
- `current_position_value` считается для всех активов:
  - для фьючерсов: `quantity * current_price * point_price`;
  - для остальных: `quantity * current_price`.
- Значение `blocked` берется из нескольких источников (приоритетно из `GetPositions`, с fallback на `GetPortfolio`) и из нескольких полей (`blocked`, `blockedLots`, `blockedQuantity`).

## 1) Что нужно подготовить

1. Создайте токен Т‑Инвестиций (Invest API) с правами на чтение портфеля.
2. Создайте/откройте Google Таблицу.
3. Откройте в таблице: **Extensions → Apps Script**.

## 2) Добавьте код

1. Вставьте содержимое файла `tinkoff_to_gsheet.gs` в редактор Apps Script.
2. В манифест проекта (`appsscript.json`) вставьте содержимое одноимённого файла из этого репозитория (нужны scope для таблиц и внешних HTTP-запросов).

## 3) Настройте токен

В Apps Script:
- **Project Settings → Script properties → Add script property**
- Name: `TINKOFF_TOKEN`
- Value: ваш токен

## 4) Запуск

Запустите функцию:

```text
syncTinkoffToSheet
```

При первом запуске Google попросит выдать разрешения скрипту.


## 5) Автозапуск и запуск через меню

После сохранения скрипта в таблице появляется меню **Т-Инвестиции** с пунктом **Обновить данные**.

- Ручной запуск: **Т-Инвестиции → Обновить данные**.
- Автозапуск при открытии файла: добавьте Script Property
  - Name: `AUTO_SYNC_ON_OPEN`
  - Value: `true`

Если `AUTO_SYNC_ON_OPEN` не задан или равен `false`, при открытии файл только добавит меню без обновления данных.

## 6) Результат

Колонки листа `positions`:
- `account_id`
- `account_name`
- `source` (скрыта)
- `figi` (скрыта)
- `instrument_name`
- `instrument_full_name`
- `instrument_type`
- `quantity`
- `blocked`
- `average_price`
- `expected_yield`
- `current_price`
- `currency`
- `lot` (скрыт)
- `min_price_increment` (скрыт)
- `point_price`
- `current_position_value`
