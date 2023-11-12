# **PECHKIN**
*Автоматизированная система почтовой рассылки.* 

## Глава 1: Цели системы
Этот проект представляет собой комплексное решение для автоматизации процесса рассылки персонализированных писем. Разработан для упрощения и ускорения процедуры отправки большого количества писем, каждое из которых требует индивидуального подхода к содержанию и адресату.

## Глава 2: Описание
Этот проект предназначен для автоматизации создания и рассылки персонализированных писем. Он состоит из нескольких ключевых компонентов, которые работают вместе, чтобы упростить и ускорить процесс рассылки:

### Подготовка данных в Excel:
 Пользователи начинают с заполнения файла list.xlsx данными адресатов. Этот файл служит основой для последующей персонализации писем. Первый .bat скрипт (1. Подготовка excel.bat) анализирует эти данные и автоматически дополняет таблицу нужной информацией, такой как обращение, должность адресата и другие персонализированные детали.

### Генерация вложений:
После подготовки данных пользователь заполняет шаблон template.docx, указывая места для замены данными из таблицы с помощью плейсхолдеров. Второй .bat скрипт (2. Генератор бланков.bat) использует этот шаблон для генерации индивидуальных документов в форматах .docx и .pdf, которые будут прикреплены к каждому письму.

### Создание и отправка писем:
Пользователь затем заполняет шаблон письма mail template.docx, готовя содержание каждого отправляемого сообщения. Скрипт для Outlook (Outlook VBA script.txt) интегрируется в почтовый клиент и использует подготовленный список адресатов и шаблон письма для автоматического создания электронных писем. Скрипт сохраняет все письма в черновиках Outlook, позволяя пользователю пересмотреть их перед отправкой.

Эта система значительно упрощает процесс рассылки, снижая трудоёмкость подготовки каждого индивидуального письма и уменьшая вероятность человеческой ошибки.

## Глава 3: Технологический Стек
Этот проект использует комбинацию нескольких технологий и инструментов для эффективной работы:

Microsoft Excel: Является основным инструментом для подготовки и хранения данных адресатов. Файл list.xlsx служит источником информации для персонализации писем.

Batch (.bat) скрипты: Два скрипта (1. Подготовка excel.bat и 2. Генератор бланков.bat) автоматизируют процессы обработки данных в Excel и генерации документов. Эти скрипты являются ключевыми для управления потоком работы в проекте.

Python: Используется для написания более сложных алгоритмов обработки данных и генерации файлов. Скрипты function_excel_modify.py и function_generate_pdf.py являются примерами таких алгоритмов.

Microsoft Word: Используется для создания шаблонов документов (template.docx и mail template.docx), которые далее заполняются данными и используются в рассылке.

Microsoft Outlook и VBA (Visual Basic for Applications): VBA скрипт (Outlook VBA script.txt) интегрируется в Outlook для автоматизации процесса создания и сохранения писем в черновиках.

Форматы файлов PDF и DOCX: Используются для создания финальных версий документов, прикрепляемых к письмам.

Этот стек технологий выбран для обеспечения гибкости и масштабируемости процесса рассылки, позволяя автоматизировать ключевые этапы работы с минимальными усилиями со стороны пользователя.

## Глава 4: Установка и Настройка
Для начала работы с системой рассылки писем пользователю необходимо выполнить установку и настройку проекта:

1. **Установка Python и Git:** Перед запуском скрипта install.bat, убедитесь, что на вашем компьютере установлены Python и Git:

* Если Python не установлен, его можно скачать по ссылке: [Python 3.11.5](https://www.python.org/ftp/python/3.11.5/python-3.11.5-amd64.exe).
* Git доступен для скачивания здесь: [Git 2.42.0.2](https://github.com/git-for-windows/git/releases/download/v2.42.0.windows.2/Git-2.42.0.2-64-bit.exe).
2. **Запуск install.bat:** После установки Python и Git, запустите скрипт install.bat. [Скачать install.bat](https://raw.githack.com/mitinrs/mail_list/main/install/install.bat)

Этот скрипт автоматически выполнит следующие действия:

* Проверит наличие Git и, при необходимости, предложит его установить.
* Клонирует репозиторий проекта с GitHub, если он еще не склонирован, или обновит его, если проект уже был склонирован.
* Проверит наличие Python и, при отсутствии, предложит его установить.
* Создаст виртуальную среду Python для изоляции зависимостей проекта.
* Активирует виртуальную среду и установит необходимые зависимости из файла requirements.txt.
3. **Настройка рабочего окружения:**

* Убедитесь, что на вашем компьютере установлены Microsoft Excel и Word для работы с .xlsx и .docx файлами.
* Установите Microsoft Outlook для использования VBA скрипта и отправки писем.

После выполнения этих шагов, система будет полностью настроена и готова к использованию.

## Глава 5: Использование
После установки и настройки системы, процесс использования системы рассылки писем состоит из нескольких ключевых этапов:

1. **Подготовка данных в Excel:**

* Откройте файл list.xlsx и введите данные адресатов, которым вы хотите отправить письма.
* Запустите скрипт **"1. Подготовка excel.bat"**. Он автоматически анализирует данные и дополняет таблицу необходимыми сведениями, такими как обращение и должность в дательном падеже.
2. **Генерация вложений:**

* Заполните шаблон **template.docx**, указывая поля для замены данными из таблицы с помощью плейсхолдеров, обозначенных квадратными скобками **[...]**.
* Запустите скрипт **"2. Генератор бланков.bat"**, который создаст индивидуальные .docx и .pdf файлы для каждого адресата.
3. **Подготовка текста письма:**

* Используйте шаблон **mail template.docx** для создания текста письма. При необходимости, вставьте в него плейсхолдеры для персонализации.
4. **Интеграция и использование скрипта в Outlook:**

* Интегрируйте Outlook VBA script.txt в Microsoft Outlook через редактор VBA.
* Запустите скрипт из Outlook, он предложит выбрать файл list.xlsx и mail template.docx.
* Скрипт автоматически создаст письма с соответствующими вложениями и сохранит их в черновиках.
5. **Проверка и отправка писем:**

* Перейдите в папку "Черновики" в Outlook и пересмотрите созданные письма.
* После проверки и убеждения в корректности содержания и адресатов, отправьте письма.

Следуя этим шагам, вы сможете эффективно и быстро рассылать персонализированные письма большому количеству адресатов, минимизируя вероятность ошибок и экономя время.

