# Создание первой надстройки Excel

В этой статье описывается, как создать надстройку для Excel 2016 или Excel Online с помощью API JavaScript для Excel. C помощью описанных ниже действий вы сможете создать простую надстройку области задач, которая загружает определенные данные на лист и создает базовую диаграмму в Excel 2016.

![Надстройка для квартального отчета о продажах](../../images/QuarterlySalesReport_report.PNG)


Сначала вы научитесь создавать веб-приложения с использованием HTML и JQuery. Затем вы узнаете, как создать XML-файл манифеста, определяющий расположение вашего веб-приложения и способ его отображения в Excel.


### Создание кода

Создайте на своем локальном диске папку с именем QuarterlySalesReport (например, C:\\QuarterlySalesReport). Сохраняйте в этой папке все файлы, которые вы будете создавать далее.

Создайте HTML-страницу, которая будет загружаться в надстройку области задач. Присвойте файлу имя **Home.html** и вставьте в него указанный ниже код.

```html

    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
        <title>Quarterly Sales Report</title>

        <script src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.1.4.min.js"></script>

        <link href="Office.css" rel="stylesheet" type="text/css" />

        <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>

        <link href="Common.css" rel="stylesheet" type="text/css" />
        <script src="Notification.js" type="text/javascript"></script>

        <script src="Home.js" type="text/javascript"></script>

        <link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/1.0/fabric.min.css">
        <link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/1.0/fabric.components.min.css">

    </head>
    <body class="ms-font-m">
        <div id="content-header">
            <div class="padding">
                <h1>Welcome</h1>
            </div>
        </div>
        <div id="content-main">
            <div class="padding">
                <p>This sample shows how to load some sample data into the worksheet, and then create a chart using the Excel JavaScript API.</p>
                <br />
                <h3>Try it out</h3>
                <button class="ms-Button" id="load-data-and-create-chart">Click me!</button>
            </div>
        </div>
    </body>
    </html>

```

Создайте файл с именем **Common.css**, в котором вы будете хранить пользовательские стили, и вставьте в него указанный ниже код.

```css
    /* Common app styling */

    #content-header {
        background: #2a8dd4;
        color: #fff;
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 80px; /* Fixed header height */
        overflow: hidden; /* Disable scrollbars for header */
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px; /* Same value as #content-header's height */
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto; /* Enable scrollbars within main content section */
    }

    .padding {
        padding: 15px;
    }

    #notification-message {
        background-color: #818285;
        color: #fff;
        position: absolute;
        width: 100%;
        min-height: 80px;
        right: 0;
        z-index: 100;
        bottom: 0;
        display: none; /* Hidden until invoked */
    }

        #notification-message #notification-message-header {
            font-size: medium;
            margin-bottom: 10px;
        }

        #notification-message #notification-message-close {
            background-image: url("../../images/Close.png");
            background-repeat: no-repeat;
            width: 24px;
            height: 24px;
            position: absolute;
            right: 5px;
            top: 5px;
            cursor: pointer;
        }


```

Создайте файл, содержащий программную логику для надстройки в jQuery. Назовите этот файл **Home.js** и вставьте в него указанный ниже сценарий.

```js

    (function () {
        "use strict";

        // The initialize function must be run each time a new page is loaded
        Office.initialize = function (reason) {
            $(document).ready(function () {
                app.initialize();

                $('#load-data-and-create-chart').click(loadDataAndCreateChart);
            });
        };

        // Load some sample data into the worksheet and then create a chart
        function loadDataAndCreateChart() {
            // Run a batch operation against the Excel object model
            Excel.run(function (ctx) {

                // Create a proxy object for the active worksheet
                var sheet = ctx.workbook.worksheets.getActiveWorksheet();

                //Queue commands to set the report title in the worksheet
                sheet.getRange("A1").values = "Quarterly Sales Report";
                sheet.getRange("A1").format.font.name = "Century";
                sheet.getRange("A1").format.font.size = 26;

                //Create an array containing sample data
                var values = [["Product", "Qtr1", "Qtr2", "Qtr3", "Qtr4"],
                              ["Frames", 5000, 7000, 6544, 4377],
                              ["Saddles", 400, 323, 276, 651],
                              ["Brake levers", 12000, 8766, 8456, 9812],
                              ["Chains", 1550, 1088, 692, 853],
                              ["Mirrors", 225, 600, 923, 544],
                              ["Spokes", 6005, 7634, 4589, 8765]];

                //Queue a command to write the sample data to the specified range
                //in the worksheet and bold the header row
                var range = sheet.getRange("A2:E8");
                range.values = values;
                sheet.getRange("A2:E2").format.font.bold = true;

                //Queue a command to add a new chart
                var chart = sheet.charts.add("ColumnClustered", range, "auto");

                //Queue commands to set the properties and format the chart
                chart.setPosition("G1", "L10");
                chart.title.text = "Quarterly sales chart";
                chart.legend.position = "right"
                chart.legend.format.fill.setSolidColor("white");
                chart.dataLabels.format.font.size = 15;
                chart.dataLabels.format.font.color = "black";
                var points = chart.series.getItemAt(0).points;
                points.getItemAt(0).format.fill.setSolidColor("pink");
                points.getItemAt(1).format.fill.setSolidColor('indigo');

                //Run the queued commands, and return a promise to indicate task completion
                return ctx.sync();
            })
              .then(function () {
                  app.showNotification("Success");
                  console.log("Success!");
              })
            .catch(function (error) {
                // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
                app.showNotification("Error: " + error);
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
```


Создайте файл, содержащий программную логику для реализации уведомлений об ошибках в надстройке. Это удобно при отладке. Назовите этот файл **Notification.js** и вставьте в него указанный ниже сценарий.

```js

    /* Notification functionality */

    var app = (function () {
        "use strict";

        var app = {};

        // Initialization function (to be called from each page that needs notification)
        app.initialize = function () {
            $('body').append(
                '<div id="notification-message">' +
                    '<div class="padding">' +
                        '<div id="notification-message-close"></div>' +
                        '<div id="notification-message-header"></div>' +
                        '<div id="notification-message-body"></div>' +
                    '</div>' +
                '</div>');

            $('#notification-message-close').click(function () {
                $('#notification-message').hide();
            });


            // After initialization, expose a common notification function
            app.showNotification = function (header, text) {
                $('#notification-message-header').text(header);
                $('#notification-message-body').text(text);
                $('#notification-message').slideDown('fast');
            };
        };

        return app;
    })();
```

Создайте XML-файл манифеста, чтобы указать расположение веб-приложения и способ его отображения в Excel. Назовите этот файл **QuarterlySalesReportManifest.xml** и вставьте в него указанный ниже XML-код.

```xml
    <?xml version="1.0" encoding="UTF-8"?>
    <!--Created:cb85b80c-f585-40ff-8bfc-12ff4d0e34a9-->
    <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
      <Id>ab2991e7-fe64-465b-a2f1-c865247ef434</Id>
      <Version>1.0.0.0</Version>
      <ProviderName>Microsoft</ProviderName>
      <DefaultLocale>en-US</DefaultLocale>
      <DisplayName DefaultValue="Quarterly Sales Report Sample" />
      <Description DefaultValue="Quarterly Sales Report Sample"/>
      <Capabilities>
        <Capability Name="Workbook" />
      </Capabilities>
      <DefaultSettings>
        <SourceLocation DefaultValue="\\MyShare\QuarterlySalesReport\Home.html" />
      </DefaultSettings>
      <Permissions>ReadWriteDocument</Permissions>
    </OfficeApp>
```

Создайте GUID с помощью любого онлайн-генератора. Затем замените полученным GUID значение тега **Id**, указанного в предыдущем действии.

Сохраните все файлы. Ваша первая надстройка Excel готова.

### Проверка

Самый простой способ развернуть и проверить надстройку — скопировать эти файлы в сетевую папку.

Создайте папку в сетевой папке (например, \\\MyShare\\QuarterlySalesReport) и скопируйте в нее все файлы.

Измените элемент **SourceLocation** в файле манифеста, чтобы он указывал на расположение общей папки для HTML-страницы, упоминавшейся при выполнении шага 1.

Скопируйте манифест (QuarterlySalesReportManifest.xml) в сетевую папку (например, \\\MyShare\\MyManifests).

Теперь добавьте расположение общей папки, содержащей этот манифест, в качестве доверенного каталога приложений в Excel. Запустите Excel и откройте пустую электронную таблицу.

Перейдите на вкладку **Файл** и выберите **Параметры**.

Выберите **Центр управления безопасностью**, а затем нажмите кнопку **Параметры центра управления безопасностью**.

Выберите пункт **Доверенные каталоги надстроек**.

В поле **URL-адрес каталога** введите путь к общему сетевому ресурсу, созданному на этапе 3, и нажмите **Добавить каталог**. Установите флажок **Показывать в меню** и нажмите кнопку **ОК**. Появится сообщение о том, что параметры будут применены при следующем запуске Office.

Теперь протестируйте и запустите надстройку. На вкладке **Вставка** в Excel 2016 выберите **Мои надстройки**.

В диалоговом окне **Надстройки Office** выберите **Общая папка**.

Выберите элементы **Пример квартального отчета о продажах** > **Вставить**. Надстройка откроется в области задач справа от текущего листа, как показано на рисунке ниже.

 ![Надстройка для квартального отчета о продажах](../../images/QuarterlySalesReport_taskpane.PNG)

Нажмите кнопку **Click me!**, чтобы данные и диаграмма появились на листе, как показано на рисунке ниже.  Для динамического обновления диаграммы просто измените данные в диапазоне.

![Надстройка для квартального отчета о продажах](../../images/QuarterlySalesReport_report.PNG)


### Дополнительные ресурсы

*  [Общие сведения о программировании надстроек Excel](excel-add-ins-javascript-programming-overview.md)
*  [Обозреватель фрагментов кода для Excel](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
*  [Примеры кода надстройки Excel](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
*  [Справочник по API JavaScript для надстроек Excel](excel-add-ins-javascript-api-reference.md)
