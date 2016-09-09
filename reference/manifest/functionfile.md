# Элемент FunctionFile

Указывает файл с исходным кодом для операций, доступных через команды надстройки, которые выполняют функцию JavaScript, а не отображают пользовательский интерфейс. Элемент **FunctionFile** — дочерний элемент элемента [FormFactor](./formfactor). Атрибуту **resid** элемента **FunctionFile** присваивается значение атрибута **id** элемента **Url** в элементе **Resources**. Последний содержит URL-адрес HTML-файла, который содержит или загружает все функции JavaScript, используемые командными кнопками надстройки без пользовательского интерфейса, как определено элементом [Control](control.md).

Ниже приведен пример элемента **FunctionFile**.


```XML
<DesktopFormFactor>
          <FunctionFile resid="residDesktopFuncUrl" />
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <!-- information about this extension point -->
          </ExtensionPoint>

          <!-- You can define more than one ExtensionPoint element as needed -->

        </DesktopFormFactor>
```

Код JavaScript в HTML-файле, на который указывает элемент **FunctionFile**, должен вызывать `Office.initialize` и определять именованные функции, принимающие один параметр — `event`. Функции должны использовать API [item.notificationMessages](../../../reference/outlook/Office.context.mailbox.item.md), чтобы сообщать пользователю о ходе выполнения, успешном завершении или ошибке. Он также должен вызывать [event.completed](../../../reference/shared/event.completed.md) после выполнения. Имена функций используются в элементе **FunctionName** для кнопок без пользовательского интерфейса.

Ниже приведен пример HTML-файла для определения функции **trackMessage**.

```js
Office.intialize = function () {
    doAuth();
}

function trackMessage (event) {
    var buttonId = event.source.id;    
    var itemId = Office.context.mailbox.item.id;
    // save this message
    event.completed();
}
```

В примере кода ниже показано, как внедрить функцию, используемую элементом **FunctionName**.




```js
        // The initialize function must be run each time a new page is loaded.
        (function () {
            Office.initialize = function (reason) {
               // If you need to initialize something you can do so here.
            };
        })();

            // Your function must be in the global namespace.
        function writeText(event) {

            // Implement your custom code here. The following code is a simple example.

            Office.context.document.setSelectedDataAsync("ExecuteFunction works. Button ID=" + event.source.id,
                function (asyncResult) {
                    var error = asyncResult.error;
                    if (asyncResult.status === "failed") {
                        // Show error message.
                    }
                    else {
                        // Show success message.
                    }
                });
           // Calling event.completed is required. event.completed lets the platform know that processing has completed.
       event.completed();
        }
```


 >**Важно!** Вызов метода **event.completed** означает, что событие успешно обработано. Если функция вызывается несколько раз, например при многократном выборе одной команды надстройки, все события автоматически помещаются в очередь. Первое событие запускается автоматически, а другие ожидают в очереди. Когда функция вызывает метод **event.completed**, для нее запускается следующий вызов в очереди. Если вы не реализуете метод **event.completed**, функция не будет работать.
