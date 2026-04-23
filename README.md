# FileUp

Репозиторий: https://github.com/ljosefl/FileUp  

Настольное приложение для загрузки листов **Excel** в **Microsoft SQL Server** (создание таблиц в UI, `replace` / `append`).

## Сборка (Windows)

```text
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
pyinstaller app.spec
```

Готовый exe: `dist\FileUp.exe`.

## Проверка обновлений (GitHub Releases)

1. Создайте публичный репозиторий и публикуйте [релизы](https://docs.github.com/repositories/releasing-projects-on-github/managing-releases-in-a-repository) с тегом вида `v1.0`, `v1.1` и т.д. (см. нумерацию в `VERSIONS.md`).  
2. В `app.py` укажите репозиторий в константе `GITHUB_REPO_FOR_UPDATES` (формат `владелец/имя-репо`) **или** задайте переменную окружения `FILEUP_GITHUB_REPO`.

При более новой версии на GitHub приложение покажет плашку с кнопкой перехода на страницу релиза или прямым скачиванием exe, если файл прикреплён к релизу как asset.

Авто‑замена установленного exe без действия пользователя не выполняется (типично для Windows нужен установщик или запуск отдельного updater’а).

## Зависимости

На машине пользователя нужен драйвер **ODBC Driver for SQL Server** (например 17 или 18).
