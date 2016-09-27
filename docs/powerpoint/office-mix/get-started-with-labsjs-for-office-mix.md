
# Начало работы с LabsJS для Office Mix



Содержимое LabsJS включает API (labs.js), примеры, документы и соответствующие файлы, с помощью которых можно разрабатывать интерактивные лаборатории, интегрировать их с Office Mix, а затем преобразовывать для просмотра в Microsoft PowerPoint. Фактически лаборатории — это Надстройки Office, создаваемые с помощью HTML5 и библиотеки JavaScript labs.js.

## Содержимое LabsJS

LabsJS включает документацию, примеры лабораторий и файлы, необходимые для создания и публикации лабораторий для Office Mix.


**Требуемые файлы**


|**Файл**|**Описание**|
|:-----|:-----|
|labs-1.0.4.js|API JavaScript LabsJS для разработки лабораторий Office Mix. Этот файл необходимо включить в проект для интеграции с Office Mix. Кроме того, он хранится в сети доставки содержимого (CDN) по адресу <code>https://az592748.vo.msecnd.net/sdk/LabsJS-1.0.4/labs-1.0.4.js</code>. При публикации приложения необходимо связать его с файлом в CDN.|
|labs-1.0.4.d.ts|Файл определения TypeScript для labs.js. Позволяет легко интегрировать код TypeScript с labs.js. Кроме того, файл определения содержит сведения обо всех компонентах labs.js. Файл TypeScript можно скачать по ссылке [http://www.typescriptlang.org/](http://www.typescriptlang.org/). Файл определения создан на основе TypeScript версии 0.9.1.1.|
|History|История выпусков библиотеки labs.js.|
|Labshost.html|Веб-страница, с помощью которой можно просматривать и отлаживать лабораторию в рамках Office Mix, но вне контекста PowerPoint. Чтобы использовать эту страницу, введите свой URL-адрес в главное поле ввода, чтобы он загрузился внутри фрейма. Данные, передаваемые между API и Office Mix при работе в PowerPoint или проигрывателе уроков Office Mix, отобразятся в полях ввода справа. Эти данные также могут быть заполнены предварительно. Обратите внимание, что в примерах лабораторий в разделе "Примеры" показаны существующие Надстройки Office Mix, выполняемые в контексте узла.|
|SampleManifest.xml|Пример манифеста Надстройки Office, используемый как шаблон для создания собственного манифеста приложения.|
|Simplelab.html|Пример лаборатории, создаваемый с помощью labs.js. Позволяет выбирать и добавлять веб-страницу, а затем отслеживать ее просмотр пользователями.|
|Simplelab.ts|Файл TypeScript, используемый для создания примера Simplelab.|
|Simplelab.js|Версия JavaScript примера Simplelab. В файлах simplelab.js и simplelab.ts используется API LabsJS.|

## Настройка среды разработки

Библиотека labs.js служит уровнем абстракции библиотеки office.js (API для Надстройки Office), поэтому лаборатории, создаваемые с помощью библиотеки labs.js, — это фактически Надстройки Office. Для работы с библиотекой labs.js и запуска этих лабораторий в Office Mix необходимо сначала зарегистрироваться в качестве разработчика Надстройки Office.


### Регистрация на сайте разработчиков для Office 365

Для начала необходимо зарегистрироваться на сайте Сайт разработчиков Office 365. Это позволит размещать и тестировать лабораторию перед ее отправкой в Магазин Office. С помощью этого сайта можно публиковать надстройку в Office Mix и тестировать ее в реальной среде.

Дополнительные сведения см. в разделе [Настройка среды для разработки надстроек SharePoint в Office 365](http://msdn.microsoft.com/library/b22ce52a-ae9e-4831-9b68-c9210af6dc54%28Office.15%29.aspx). Необходимо выполнить только первые два шага. Устанавливать средства разработки Napa необязательно.


### Настройка каталога приложений в SharePoint Online

После создания и подготовки сайта разработчика необходимо настроить каталог надстроек в SharePoint Online. Узнать больше можно в разделе [Настройка каталога надстроек в Office 365](../../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).

Для Office Mix используется каталог надстроек, что позволяет добавлять предварительные надстройки в урок и проводить сквозное тестирование лабораторий перед их отправкой в магазин.


## Создание лаборатории

Чтобы создать первую лабораторию, выполните действия, описанные в [пошаговом руководстве](../../powerpoint/office-mix/creating-your-first-lab-for-office-mix.md). В нем объясняется, как создать простой тест с ответами "правда/неправда". См. раздел [Создание первой лаборатории для Office Mix — пошаговое руководство](../../powerpoint/office-mix/creating-your-first-lab-for-office-mix.md).


## Публикация лаборатории

После создания лабораторию можно опубликовать и отправить в магазин.


### Создание и отправка манифеста приложения

Манифест приложения — это XML-документ, описывающий лабораторию LabJS. Он содержит ссылку на URL-адрес, где размещается лаборатория, а также сведения о лаборатории, включая отображаемое имя, описание, значки, размер и т. д.

Мы включили пример манифеста — SampleManifest.xml. Дополнительные сведения о схеме манифеста, а также ссылку на определение схемы см. в разделе [XML-манифест надстроек для Office](../../../docs/overview/add-in-manifests.md).

Чтобы отправить манифест на сайт SharePoint, сначала перейдите в каталог приложений, который обычно расположен по URL-адресу <code>https://\<your site\>/sites/AppCatalog</code>. Затем нажмите кнопку **Новое приложение** и следуйте инструкциям по отправке манифеста приложения.


### Обновление каталога PowerPoint 2013

Затем обновите каталог PowerPoint 2013. После этого можно войти в систему с помощью учетной записи разработчика.

Сначала обновите каталог PowerPoint 2013. Запустите PowerPoint 2013 и перейдите в меню  **Файл > Параметры > Центр управления безопасностью > Параметры центра управления безопасностью > Доверенные каталоги приложений**. Здесь добавьте ссылку на свой каталог приложений и нажмите кнопку  **ОК**. В PowerPoint 2013 вам будет предложено выйти из приложения, чтобы изменения вступили в силу. Выйдите из приложения.

В завершение снова войдите в приложение с помощью учетной записи разработчика. Выберите имя для входа в верхнем правом углу в PowerPoint 2013 и войдите в приложение, используя учетную запись разработчика. Теперь надстройку можно добавить.


### Добавление, публикация и просмотр надстройки

Чтобы добавить надстройку в каталог, выберите ленту  **Вставка**, а затем пункт  **Магазин** в разделе **Приложения**. Выберите пункт  **Моя организация**, где вы увидите надстройку в своем каталоге надстроек. Выберите надстройку и нажмите кнопку  **Вставить**, чтобы добавить надстройку (лабораторию) в документ PowerPoint 2013.

Теперь вы можете воспользоваться всеми доступными функциями Office Mix, чтобы опубликовать урок с новой лабораторией.


 >**Важно!** Чтобы просмотреть надстройку, войдите в каталог SharePoint в том же браузере, который используется для просмотра урока. Доступ к каталогам SharePoint разрешен только пользователям, которые прошли проверку подлинности, поэтому для просмотра надстройки сначала необходимо войти в каталог. 


### Отправка лаборатории в Магазин Office

Отправка лаборатории в Магазин Office описана в статье [Публикация надстройки Office](../../publish/publish.md).


## Дополнительные ресурсы



- [Надстройки Office Mix](../../powerpoint/office-mix/office-mix-add-ins.md)
    
- [Надстройки Office](../../../docs/overview/office-add-ins.md)
    
- [Создание первой лаборатории для Office Mix](../../powerpoint/office-mix/creating-your-first-lab-for-office-mix.md)
    