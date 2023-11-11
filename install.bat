@echo off

:: Указываем URL вашего репозитория
set repo_url=https://github.com/mitinrs/mail_list.git

:: Проверка наличия Git
where git >nul 2>nul
IF %ERRORLEVEL% NEQ 0 (
    echo Git не установлен. Установите Git.
    pause
    exit
)

:: Проверка наличия папки с проектом
IF EXIST mail_list (
    echo Проект уже склонирован. Обновление...
    cd mail_list
    git pull
) ELSE (
    :: Клонирование вашего репозитория
    echo Клонирование репозитория...
    git clone %repo_url%
    cd mail_list
)

:: Проверка наличия Python
where python >nul 2>nul
IF %ERRORLEVEL% NEQ 0 (
    echo Python не установлен. Установите Python.
    pause
    exit
)

:: Проверка наличия виртуальной среды
IF NOT EXIST venv (
    echo Создание виртуальной среды...
    python -m venv venv
)

:: Активация виртуальной среды
call venv\Scripts\activate

:: Установка зависимостей
pip install -r requirements.txt

echo Готово!
pause
