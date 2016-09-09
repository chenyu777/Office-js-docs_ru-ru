
# Метод ProjectDocument.getSelectedDataAsync
Асинхронно получает текстовое значение данных из одной или нескольких выбранных ячеек в представлении диаграммы Ганта.

|||
|:-----|:-----|
|**Ведущие приложения:**|Project|
|**Доступно в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Выделение|
|**Добавлен в версии**|1.0|

```
Office.context.document.getSelectedDataAsync(coercionType[, options][, callback]);
```


## Параметры



|**Имя**|**Тип**|**Описание**|**Примечания по вопросам поддержки**|
|:-----|:-----|:-----|:-----|
| _coercionType_|[CoercionType](../../reference/shared/coerciontype-enumeration.md)|Тип возвращаемой структуры данных. Обязательный.<br/>В Project 2013 поддерживаются только перечисления **Office.CoercionType.Text** или `"text"`.||
| _options_|**object**|Задает следующие [необязательные параметры](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):||
| _valueFormat_|[ValueFormat](../../reference/shared/valueformat-enumeration.md)|Задает формат для значений даты или числовых значений.<br/>В Project 2013 этот параметр игнорируется и внутри приложения ему присваивается значение `unformatted`.||
| _filterType_|[FilterType](../../reference/shared/filtertype-enumeration.md)|Указывает, следует ли включать только видимые данные или все данные. <br/>В Project 2013 этот параметр игнорируется и внутри приложения ему присваивается значение `all`.||
| _asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string** или **undefined**|Определяемый пользователем элемент любого типа, который возвращается в объекте **AsyncResult** без изменения.||
| _callback_|**object**|Функция, вызываемая при возвращении обратного вызова, единственный параметр которой имеет тип **AsyncResult**.||

## Значение обратного вызова

При выполнении функция _callback_ получает объект [AsyncResult](../../reference/shared/asyncresult.md), к которому можно получить доступ с помощью параметра функции обратного вызова.

В случае метода **getSelectedDataAsync** возвращенный объект [AsyncResult](../../reference/shared/asyncresult.md) содержит следующие свойства:


****


|**Имя**|**Описание**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Данные, переданные в необязательном параметре _asyncContext_ (если он использовался).|
|[error](../../reference/shared/asyncresult.error.md)|Сведения об ошибке, если свойство **status** имеет значение **failed**.|
|[status](../../reference/shared/asyncresult.status.md)|Состояние **succeeded** или **failed** асинхронного вызова.|
|[value](../../reference/shared/asyncresult.value.md)|Текстовое значение выбранных ячеек.|

## Заметки

Метод **ProjectDocument.getSelectedDataAsync** переопределяет метод [Document.getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md) и возвращает текстовое значение данных, выбранных в одной или нескольких ячейках представления диаграммы Ганта. Метод **ProjectDocument.getSelectedDataAsync** поддерживает только текстовый формат параметра [CoercionType](../../reference/shared/coerciontype-enumeration.md). Он не поддерживает форматы `matrix`, `table` и другие форматы.


## Пример

Ниже приведен пример кода, который получает значения выбранных ячеек. Для передачи текста в функцию обратного вызова он использует необязательный параметр _asyncContext_.

В данном примере подразумевается, что в вашей надстройке есть ссылка на библиотеку jQuery и что указанные ниже элементы управления страницы определены в теге div контента в тексте страницы.




```HTML
<input id="get-info" type="button" value="Get info" /><br />
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // After the DOM is loaded, add-in-specific code can run.
            $('#get-info').click(getSelectedText);
        });
    };

    // Get the text from the selected cells in the document, and display it in the add-in.
    function getSelectedText() {
        Office.context.document.getSelectedDataAsync(
            Office.CoercionType.Text,
            {asyncContext: 'Some related info'},
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    var output = String.format(
                        'Selected text: {0}<br/>Passed info: {1}',
                        result.value, result.asyncContext);
                    $('#message').html(output);
                }
            }
        );
    }

    function onError(error) {
        $('#message').html(error.name + ' ' + error.code + ': ' + error.message);
    }
})();
```


## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает на то, что этот метод поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает этот метод.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office для рабочего стола Windows**|**Office Online (в браузере)**|
|:-----|:-----|:-----|
|**Project**|Y||

|||
|:-----|:-----|
|**Доступен в наборах требований**|Выделение|
|**Минимальный уровень разрешений**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Область задач|
|**Библиотека**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки



****


|**Версия**|**Изменения**|
|:-----|:-----|
|1.0|Представлено|

## См. также



#### Другие ресурсы


[Объект AsyncResult](../../reference/shared/asyncresult.md)

[Office.CoercionType](../../reference/shared/coerciontype-enumeration.md)

[Объект ProjectDocument](../../reference/shared/projectdocument.projectdocument.md)
