# Описание
Импортер Яндекс-метрик в Excel. По заданным метрикам и группировкам производится табличный запрос к 
[API Яндекс Метрик](https://tech.yandex.ru/metrika/doc/api2/api_v1/data-docpage/). Формируется excel-файл. 
Для каждой метрики формируется воркшит, куда построчно вносятся данные - один запрос - одна строка. В первой колонке период 
запроса. На первом воршите формируется дашборд, где построчно для каждой цели выводятся данные последнего запроса. 
В первой колонке название метрики.
# Установка и использование 
Скачайте архив с исходниками или клонируйте проект.

Установите пакеты.
```.env
pip3 install -r requirements.txt
```
Добавьте файл с параметрами params.ini используя инструкцию ниже.

Запуск.
```.env
python3 yandxl.py [-h|-i (--init)|-a(--add)]
```
- -h – помощь
- -i – создать или пересоздать заново excel-файл и добавить в него результат запроса метрик
- -a - добавить результат запроса метрик в существующий excel-файл
# Формат params.ini
[yam]

METRICS - список метрик через запятую без пробела (ym:s:users,ym:s:pageviews)

DIMENSIONS - список группировок через запятую (ym:s:paramsLevel1,ym:s:paramsLevel2)

YANDEX_TOKEN - TODO разобраться как он работает не в отладочном режиме

COUNTER - идентификатор счетчика

API_ROOT_URL - адрес API по умолчанию https://api-metrika.yandex.net/stat/v1/data

PERIOD - период, за который нужны данные. В соответствии с документацией (today, yesterday, ndaysAgo), например 10daysAgo

[excel]

DASHBOARD_WS_NAME - название первого вокршита отчета

ROW_TITLES_DASHBOARD - заголовки колонок на дашборде через запятую без пробела (пока захардкожены первые две метрики), где первая 
колонка занята под название метрики. Например Метрика,Уникальных пользователей,Всего событий

ROW_TITLES - заголовки колонок на воркшите метрики через запятую без пробела, где первая 
колонка занята под дату. Например Дата,Пользователи,Просмотры

WB_NAME - Название файла отчета, например Brokerage_ymetrics_report.xlsx

[smtp]

SERVER  SMTP-сервер почтого ящика, с которого будет производится отправка отчета

PORT - порт SMTP-сервера

FROM - адрес почтого ящика, с которого будет производится отправка отчета

TO - адресаты через запятую без пробела

SUBJECT - тема письма

TEXT - текст письма

PATH = относительный путь до файла отчета (./)

PASS = пароль от почты для приложения. [Описание](https://yandex.ru/support/passport/authorization/app-passwords.html).

# Пример ini-файла
```ini
[yam]
metrics = ym:s:users,ym:s:pageviews,ym:s:visits,ym:s:sumParams,ym:s:paramsNumber,ym:s:avgParams,ym:s:bounceRate,ym:s:pageDepth,ym:s:avgVisitDurationSeconds
dimensions = ym:s:paramsLevel1,ym:s:paramsLevel2,ym:s:paramsLevel3,ym:s:paramsLevel4,ym:s:paramsLevel5,ym:s:paramsLevel6
YANDEX_TOKEN = <TOKEN>
COUNTER = 12121212
API_ROOT_URL = https://api-metrika.yandex.net/stat/v1/data
PERIOD = yesterday

[excel]
DASHBOARD_WS_NAME = Дашборд
ROW_TITLES_DASHBOARD = Метрика,Уникальных пользователей,Всего событий
ROW_TITLES = Дата,Пользователи,Просмотры,Визиты,Сумма параметров визитов,Количество параметров визитов,Среднее параметров визитов,Отказы,Глубина просмотра,Время на сайте
WB_NAME = report.xlsx

[smtp]
SERVER = smtp.yandex.com
PORT = 465
FROM = <EMAIL>
TO = <EMAIL>,<EMAIL>,<EMAIL>...
SUBJECT = Отчет
TEXT = Ежедневный отчет о действиях пользователей
PATH = ./
PASS = <PASSWORD>
```