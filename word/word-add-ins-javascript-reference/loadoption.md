# Объект LoadOption (API JavaScript для Word)

Объект, определяющий сведения о разбивке по страницам и свойства для загрузки при вызове context.sync(). 

_Область применения: Word 2016, Word для iPad, Word для Mac_

## Свойства
| Свойство   | Тип|Описание|
|:---------------|:--------|:----------|
|select|object|Содержит массив или разделенный запятыми список имен параметров и связей. Необязательный параметр.|
|expand|object|Содержит массив или разделенный запятыми список имен связей. Необязательный параметр.|
|top|int| Указывает максимальное число элементов в коллекции, которые можно включить в результат. Необязательный параметр.|
|skip|int|Укажите количество элементов в коллекции, которые необходимо пропустить и исключить из результата. Если указан параметр `top`, результирующий набор начнется после пропуска заданного числа элементов. Необязательный параметр.|

## Дополнительные сведения

Для указания свойств и сведений о разбивке на страницы рекомендуется использовать строковый литерал. В первых двух примерах показан предпочтительный способ запроса свойств размера текста и шрифта для абзацев в коллекции абзацев:

<code>context.load(paragraphs, 'text, font/size, top: 50, skip: 0');</code>

<code>paragraphs.load('text, font/size, top: 50, skip: 0');</code>

Вот эквивалент с использованием объектной нотации:

&lt;code&gt;context.load(paragraphs, {select: 'text, font/size',
                                expand: 'font',
                                top: 50,
                                skip: 0});</code>
                                
&lt;code&gt;paragraphs.load({select: 'text, font/size',
                       expand: 'font',
                       top: 50,
                       skip: 0});</code>

Обратите внимание, что если не задать определенные свойства объекта шрифта в инструкцию select, инструкция expand сама по себе означает, что загружаются все свойства шрифта. 

## Примеры

В этом примере показано, как получить 50 верхних абзацев в документе Word, а также свойства размера текста и шрифта для них.

```js
        // Run a batch operation against the Word object model.
        Word.run(function (context) {

            // Create a proxy object for the paragraphs collection.
            var paragraphs = context.document.body.paragraphs;

            // Queue a commmand to load the text and font properties for the top 50 paragraphs.
            // It is best practice to always specify the property set. Otherwise, all properties are
            // returned in on the object. 
            context.load(paragraphs, 'text, font/size, top: 50, skip: 0');

            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
            
            // Insert code that works with the paragraphs loaded by context.load().

        })
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });

```

## Сведения о поддержке

Используйте [набор требований](https://msdn.microsoft.com/EN-US/library/office/mt590206.aspx) в проверках среды выполнения, чтобы обеспечить поддержку ведущей версии Word для своего приложения. Дополнительные сведения о требованиях ведущих приложений и серверов Office см. в статье [Требования для запуска надстроек Office](https://msdn.microsoft.com/EN-US/library/office/dn833104.aspx). 
