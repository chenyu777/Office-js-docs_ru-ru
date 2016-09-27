
# Жизненный цикл разработки надстроек Office


Типичный жизненный цикл разработки надстройки Office включает следующие шаги:


1.  **Определение назначения надстройки.**
    
    Задайте указанные ниже вопросы.
    
      - В чем польза от этой надстройки? 
    
      - Как оно поможет пользователям повысить производительность своего труда?
    
      - Какие сценарии поддерживают функции вашей надстройки?
    

    Определите наиболее важные возможности и сценарии и сосредоточьтесь на них при разработке надстройки. 
    
2.  **Определите данные и источник данных для новой надстройки.**
    
    Находятся ли данные в документе, книге, презентации, проекте или браузерной базе данных Access? Это данные об элементе или элементах на сервере Exchange Server или в почтовом ящике Exchange Online? Данные получены из внешнего источника (например, веб-службы)?
    
3.  **Определите тип надстройки и ведущие приложения Office, которые будут наиболее приемлемы для создаваемой надстройки.**
    
    При определении сценариев учитывайте следующее:
    
    - Будут ли клиенты использовать надстройку для улучшения содержимого документа или браузерной базы данных Access? Если это так, может быть, целесообразно создать контентную надстройку. 
    
    - Будут ли клиенты использовать надстройку во время просмотра или создания электронного сообщения или встречи? Важна ли возможность отображать надстройку с учетом текущего контекста? Важно ли сделать надстройку доступной не только на настольных компьютерах, но и на планшетах?
    
        Если вы ответили "да" на какой-либо из этих вопросов, рассмотрите возможность создания надстройки Outlook. Затем определите, в каком контексте будет активироваться надстройка (например, в форме создания, для определенных типов сообщений, при наличии вложения, адреса, предложения задачи, приглашения на собрание или определенных строковых шаблонов в тексте сообщения или сведениях о встрече). Сведения о том, как активировать надстройку Outlook в зависимости от контекста, см. в статье [Правила активации для надстроек Outlook](../outlook/manifests/activation-rules.md).
    
    - Будут ли клиенты использовать надстройку для расширения возможностей при просмотре или создании документа? Если это так, может быть, целесообразно создать надстройку области задач. 

    Поддержка некоторых API для надстроек может отличаться в зависимости от того, в каком приложении Office и на какой платформе они работают (в Windows, Mac, веб-приложениях и на мобильных устройствах). Список поддерживаемых API по клиентам и платформам представлен на странице [Доступность ведущих приложений и платформ для надстроек Office](https://dev.office.com/add-in-availability).  
    
4.  **Разработайте и реализуйте пользовательский интерфейс надстройки.**
    
    Разработайте возможности быстрый и удобный пользовательский интерфейс, который будет согласован, прост в освоении и позволит выполнять основные действия за всего несколько этапов. В зависимости от назначения надстройки используйте сторонние интерфейсы API и веб-службы.
    
    Для реализации пользовательского интерфейса можно пользоваться любым из множества доступных средств веб-разработки и применять HTML и JavaScript.
    
5.  **Создайте XML-файл манифеста на основе схемы манифеста Надстройки Office.**
    
    Создайте XML-манифест для идентификации надстройки и ее требований, укажите местоположение файлов HTML, JavaScript и CSS, которые использует надстройка. Кроме того, укажите размер и разрешения по умолчанию в зависимости от типа надстройки.
    
    Для надстроек Outlook можно указать контекст, основанный на текущем сообщении или встрече, в котором надстройка станет актуальной и будет отображаться в пользовательском интерфейсе Outlook. Кроме того, вы можете выбрать, на каких устройствах будет работать надстройка. В манифесте укажите контекст в виде правил активации и поддерживаемых устройств.
    
6.  **Установите и протестируйте надстройку.**
    
    Поместите HTML-файлы и любые файлы JavaScript и CSS на веб-серверы, указанные в файле манифеста надстройки. Процесс установки надстройки зависит от его типа.
    
    Если это надстройка Outlook, установите ее в почтовый ящик Exchange и укажите расположение манифеста надстройки в Центре администрирования Exchange (EAC). Дополнительные сведения см. в статье [Развертывание и установка надстроек Outlook для тестирования](../outlook/testing-and-tips.md).
    
7.  **Опубликуйте надстройку.**
    
    Можно отправить надстройку в магазин Office, откуда пользователи смогут ее скачать и установить. Кроме того, можно опубликовать надстройки области задач и контентные надстройки в каталоге надстроек личной папки SharePoint или в общей сетевой папке, а надстройку Outlook можно развернуть непосредственно для вашей организации. Дополнительные сведения см. в статье [Публикация надстройки Office](../publish/publish.md).
    
8.  **Обновление надстройки**
    
    Если надстройка вызывает веб-службу, а вы вносите изменения в веб-службу уже после публикации надстройки, повторно публиковать ее не нужно. Тем не менее, если изменить какие-либо элементы или данные, которые вы уже отправили для надстройки (например, манифест надстройки, снимки экрана, значки, файлы HTML или JavaScript), надстройку необходимо будет повторно опубликовать. В частности, если вы опубликовали надстройку в магазине Office, необходимо будет отправить ее повторно, чтобы магазин Office мог применить эти изменения. Повторно отправлять надстройку нужно с обновленным манифестом надстройки, содержащим новый номер версии. Необходимо также обновить номер версии надстройки в форме отправки в соответствии с новым номером версии манифеста. В случае надстроек Outlook следует убедиться, что элемент [Id](../../reference/manifest/id.md) содержит другой UUID в манифесте надстройки.
    