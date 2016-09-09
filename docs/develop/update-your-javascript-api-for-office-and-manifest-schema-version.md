
# Обновление версии API JavaScript для Office и файлов схемы манифеста



В этой статье рассказывается, как обновить файлы JavaScript (Office.js и JS-файлы для конкретной надстройки) и файл проверки манифеста надстройки в проекте надстройки Office до версии 1.1.

## Использование файлов проекта последней версии

Если для разработки надстройки вы используете Visual Studio, то чтобы можно было применять [самые новые элементы API](../../reference/what's-changed-in-the-javascript-api-for-office.md) в API JavaScript для Office и [возможности манифеста надстройки версии 1.1](../../docs/overview/add-in-manifests.md) (который проверяется на соответствие offappmanifest-1.1.xsd), вам потребуется скачать и установить [Visual Studio 2015 и последнюю версию Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs).

Если вы используете текстовый редактор или другую интегрированную среду разработки, отличную от Visual Studio, чтобы разработать надстройка, обновите ссылки на CDN для файла Office.js и версию схемы, на которую ссылается манифест приложения для Office.

Чтобы запустить надстройку, разработанную с использованием новых и обновленных компонентов манифеста надстройки и интерфейса API Office.js, ваши клиенты должны использовать локальные продукты Office 2013 с пакетом обновления 1 (SP1) или более поздней версии, а также при необходимости SharePoint Server 2013 с пакетом обновления 1 (SP1) и связанными серверными продуктами, Пакет обновления 1 (SP1) для Exchange Server 2013 или аналогичные размещенные в сети продукты: Office 365, SharePoint Online и Exchange Online.

Сведения о том, как скачать Office, SharePoint и продукты Exchange с пакетом обновления 1, см. в следующих статьях:


- [Список всех пакетов обновления 1 (SP1) для Microsoft Office 2013 и связанных продуктов для настольных систем](http://support.microsoft.com/kb/2850036)
    
- [Список всех пакетов обновления 1 (SP1) для Microsoft SharePoint Server 2013 и связанных серверных продуктов](http://support.microsoft.com/kb/2850035)
    
- [Описание пакета обновления 1 для Exchange Server 2013](http://support.microsoft.com/kb/2926248)
    

## Как обновить проект надстройки Office, созданный с помощью Visual Studio, для использования библиотеки API JavaScript для Office последней версии и схемы манифеста надстройки версии 1.1


Для проектов, созданных до выпуска версии 1.1 библиотеки JavaScript API для Office и схемы манифеста надстройки, вы можете обновить файлы проекта, используя **диспетчер пакетов NuGet**, а затем добавить ссылки на них в HTML-страницы надстройки. 

Обратите внимание, что процесс обновления применяется к _проектам по отдельности_. Вам потребуется повторить его для каждого проекта надстройки, в котором вы хотите использовать Office.js и схему манифеста надстройки версии 1.1.




### Обновление файлов библиотеки API JavaScript для Office в проекте до последней версии


1. В Visual Studio 2015 откройте или создайте проект **Надстройка Office**.
    
      - В расположенной слева области щелкните **Обновить** и завершите процесс обновления пакета.
    
  - Перейдите к этапу 6.
    
2. Выберите **Средства**  >  **Диспетчер пакетов NuGet**  >  **Управление пакетами Nuget для решения**.
    
3. В **диспетчере пакетов NuGet** выберите **nuget.org** в качестве **источника пакетов** и **Доступны обновления** в поле **Фильтр**. Затем выберите файл Microsoft.Office.js.
    
4. В расположенной слева области щелкните **Обновить** и завершите процесс обновления пакета.
    
5. В теге **head** HTML-страниц надстройки закомментируйте или удалите все существующие ссылки на скрипт office.js (например, `<script src="https://appsforoffice.microsoft.com/lib/1.0/hosted/office.js" type="text/javascript"></script>)`) и добавьте ссылки на обновленную библиотеку API JavaScript для Office, как показано ниже (изменив значение версии на 1). 

   >**Примечание.** "/1/" перед "office.js" в указанном ниже URL-адресе CDN указывает, что необходимо использовать последний добавочный выпуск в рамках Office.js версии 1.
    
```
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```


### Чтобы обновить файл манифеста в проекте для использования версии 1.1:


- В файле манифеста надстройки проекта (_projectname_ Manifest.xml) обновите атрибут **xmlns** элемента **OfficeApp**, изменив значение версии на 1.1 и оставив все атрибуты, отличные от **xmlns**, без изменений.
    
```XML
  <OfficeApp xsi:type="ContentApp" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" >
```


>
  **Примечание.** После обновления схемы манифеста надстройки до версии 1.1 вам потребуется удалить элементы **Capabilities** и **Capability** и заменить их либо [элементами Hosts и Host](http://msdn.microsoft.com/library/cff9fbdf-a530-4f6e-91ca-81bcacd90dcd%28Office.15%29.aspx), либо [элементами Requirements и Requirement](../../docs/overview/specify-office-hosts-and-api-requirements.md).

## Как обновить проект Надстройка Office, созданный с помощью текстового редактора или другой интегрированной среды разработки, чтобы использовать библиотеку API JavaScript для Office последней версии и схему манифеста надстройки версии 1.1


Если вы создали проект до выпуска схемы манифеста надстройки и API JavaScript для Office версии 1.1, обновите HTML-страницы вашей надстройки, чтобы они ссылались на CDN библиотеки версии 1.1, а также обновите файл манифеста надстройки, чтобы использовалась схема версии 1.1. 

Процесс обновления применяется к _проектам по отдельности_. Вам потребуется повторить его для каждого проекта надстройки, в котором вы хотите использовать файл Office.js и схему манифеста надстройки версии 1.1.

Вам не нужны локальные копии файлов API JavaScript для Office (Office.js и JS-файлов для конкретной надстройки), чтобы разрабатывать надстройку Office (ссылки на CDN для Office.js позволяют скачивать необходимые файлы во время выполнения). Если вам нужны файлы библиотеки, то вы можете скачать их с помощью [служебной программы командной строки NuGet](http://docs.nuget.org/consume/installing-nuget) и `Install-Package Microsoft.Office.js`.

 > **Примечание.** Чтобы получить копию XSD (определения схемы XML) для манифеста надстройки версии 1.1, см. статью [Справочник по схеме манифестов надстроек Office (версия 1.1)](../overview/add-in-manifests.md).


### Чтобы обновить файлы библиотеки API JavaScript для Office в проекте до последней версии, сделайте следующее:


1. Откройте HTML-страницы надстройки в текстовом редакторе или интегрированной среде разработки.
    
2. В теге **head** HTML-страниц надстройки закомментируйте или удалите все существующие ссылки на скрипт office.js (например, `<script src="https://appsforoffice.microsoft.com/lib/1.0/hosted/office.js" type="text/javascript"></script>)`) и добавьте ссылки на обновленную библиотеку API JavaScript для Office, как показано ниже (изменив значение версии на 1).
    
```
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```


    The  `/1/` in front of `office.js` in the CDN URL specifies to use the latest incremental release within version 1 of Office.js.
    

### Чтобы обновить файл манифеста в проекте для использования версии 1.1:


- В файле манифеста надстройки своего проекта (_имя_проекта_ Manifest.xml) обновите атрибут **xmlns** элемента **OfficeApp**, изменив значение версии на `1.1` и оставив все атрибуты, отличные от **xmlns**, без изменений:
    
```XML
<OfficeApp xsi:type="ContentApp" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" >
```

>
  **Примечание.** После обновления схемы манифеста надстройки до версии 1.1 вам потребуется удалить элементы **Capabilities** и **Capability** и заменить их либо [элементами Hosts и Host](http://msdn.microsoft.com/library/cff9fbdf-a530-4f6e-91ca-81bcacd90dcd%28Office.15%29.aspx), либо [элементами Requirements и Requirement](../../docs/overview/specify-office-hosts-and-api-requirements.md).
    

## Дополнительные ресурсы



- [Указание ведущих приложение Office и требований к API](../../docs/overview/specify-office-hosts-and-api-requirements.md)
    
- [Общие сведения об интерфейсе JavaScript API для Office](../../docs/develop/understanding-the-javascript-api-for-office.md)
    
- [API JavaScript для Office](../../reference/javascript-api-for-office.md)
    
- [Справка по схеме для манифестов надстроек Office (версия 1.1)](../overview/add-in-manifests.md)
    
