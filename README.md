# Robot
Process automation suite for wholesale electricity and capacity market routines

## Основные задачи
- Приём, обработка и рассылка макетов XML 80020\80040
- Приём и обработка заявок потребления определенного шаблона для преобразования в формат XML 30308
- Создание шаблонов работы на базе макетов XML 80000 и XML 6000X
- Отправка потребления рассчитанного на базе макетов XML 80020 по алгоритму XML 6000X на сервис СО ПАК Энергия 2010 (API)
- Сбор и отправка данных о вероятных часах пик в формате отчетов и XML (НПО МИР)
- Распространение файлов конфигурации для смежных проектов

## Установка
В качестве носителя этого кода выступает VBA система Outlook 2016\2019 (к сожалению, это не шутка).

1. Для этого необходимо импортировать файлы из папки src в качестве модулей макросов (Alt+F11 в Outlook).
В состав входят следующие модули:
- ThisOutlookSession.cls (основной код для сессии Outlook позволяет запускать и отключать работу ПО при запуске\закрытии Outlook)
- TimerUnit.bas (модуль для управления запуском автоматических заданий, включение, настройка и отключение автоматики происходит здесь)
- Configurator.bas (модуль управления конфигурациями, основные операции по ресурсам и источникам собраны здесь)
- Interceptor.bas (модуль обработки и отправки почтовых сообщений)
- Calendar.bas (модуль работы с календарем)
- LogWork.bas (модуль работы с логами)
- CalcRoute.bas (модуль для расчётов данных)
- BRForecast.bas (модуль для формирования отчетов о вероятных часах пик)
- XMLUtils.bas (модуль с различными функциями для работы с XML)
- UMODOv1.bas (модуль утилитивных функций для работы с файлами и т.п.)
- M_F63.bas (модуль для формирования и отправки отчетов в ПАК Энергия 2010)
- CEnergyAPI.cls (класс для работы с API ПАК Энергия 2010)
- M_CFGDrop.bas (модуль для распространения конфигураций)
- CATSDownloader.cls (класс для загрузки отчетов с сайта ООО АТС)
- CReport.cls (класс сборки состояний объектов)

2. Файлы конфигурации должны быть расположены в %HOMEPATH% (папка проекта home = %HOMEPATH%) текущего профиля пользователя.
В состав конфигураторов входят:
- Init.xml (основной файл конфигурации, здесь указываются данные о трейдере и почтовом ящике на котором будет работать проект)
- Basis.xml (конфигуратор состава ГТП и автоматизаций информобмена)
- Calendar.xml (календарь для помощи автоматам в определении рабочих и не рабочих дней, а так же часов пик СО и АТС)
- Converter.xml (шаблоны для автоматической конвертации данных входящих XML 80020)
- Credentials.xml (данные для доступа к сервисам, а также внешние интерфейсы для рассылки почты)
- Dictionary.xml (словарик общих параметров по регионам АТС и ценовым зонам АТС)
- Frame.xml (заполняется автоматически, содержит в себе кодировки ТИ и ТП активных ГТП)
- CalcRoute.xml (заполняется автоматически, содержит в себе алгоритмы для расчётов данных по ГТП)
- Набор XSD файлов необходимых для работы проверяющих механизмов проекта
