
# Объект Bindings
Представляет привязки, которые есть у надстройки в документе.

|||
|:-----|:-----|
|**Ведущие приложения:**|Access, Excel, Word|
|**Последнее изменение** в|1.1|

```js
Office.context.document.bindings
```


**Свойства**

|||
|:-----|:-----|
|Имя|Описание|
|[document](../../reference/shared/bindings.document.md)|Получает объект **Document**, представляющий документ, связанный с набором привязок.|

**Методы**

|||
|:-----|:-----|
|Имя|Описание|
|[addFromNamedItemAsync](../../reference/shared/bindings.addfromnameditemasync.md)|Добавляет привязку к именованному элементу в документе.|
|[addFromPromptAsync](../../reference/shared/bindings.addfrompromptasync.md)|Отображает пользовательский интерфейс, в котором пользователь может выбрать цель привязки.|
|[addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md)|Добавляет объект привязки указанного типа, привязанный к текущему фрагменту документа.|
|[getAllAsync](../../reference/shared/bindings.getallasync.md)|Получает все ранее созданные привязки.|
|[getByIdAsync](../../reference/shared/bindings.getbyidasync.md)|Получает указанную привязку по идентификатору.|
|[releaseByIdAsync](../../reference/shared/bindings.releasebyidasync.md)|Удаляет указанную привязку.|

## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает на то, что этот метод поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает этот метод.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).


|||||
|:-----|:-----|:-----|:-----|
||Office для рабочего стола Windows|Office Online (в браузере)|Office для iPad|
|**Access**||Y||
|**Excel**|Y|Да|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Типы надстроек**|Контентные надстройки и надстройки области задач|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки



|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлена поддержка Excel и Word в Office для iPad.|
|1.1|Для [addFromNamedItemAsync](../../reference/shared/bindings.addfromnameditemasync.md), [addFromPromptAsync](../../reference/shared/bindings.addfrompromptasync.md) и [addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md) добавлена поддержка привязки к данным матрицы аналогично привязке к таблице в надстройках для Excel.|
|1.1|<ul><li>Для свойства <a href="8fa0cb4a-fad1-4f2e-9a7e-5f7aa7789eca.htm">document</a> добавлена возможность доступа к объекту <span class="keyword">Document</span>, который представляет собой текущую базу данных Access в контентных надстройках для Access.</li><li>Для всех методов добавлена поддержка привязки таблиц в контентных надстройках для Access. </li></ul>|
|1.0|Представлено|
