# Объект InkAnalysisWord (API JavaScript для OneNote)

_Применяется для OneNote Online_  


Представляет данные анализа рукописного фрагмента для определенного слова, созданного росчерками пера.

## Свойства

| Свойство     | Тип   |Описание|Отзывы|
|:---------------|:--------|:----------|:-------|
|id|string|Получает идентификатор объекта InkAnalysisWord. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWord-id)|
|languageId|string|Идентификатор распознанного языка в этом inkAnalysisWord. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWord-languageId)|
|wordAlternates|string|Слова, которые были распознаны в этом слове рукописного фрагмента, в порядке вероятности. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWord-wordAlternates)|

_См. [примеры](#примеры) доступа к свойствам._

## Связи
| Связь | Тип   |Описание| Отзыв|
|:---------------|:--------|:----------|:-------|
|line|[InkAnalysisLine](inkanalysisline.md)|Ссылка на родительский объект InkAnalysisLine. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWord-line)|
|strokePointers|[InkStrokePointer](inkstrokepointer.md)|Слабые ссылки на росчерки пера, которые были распознаны как часть этого слова анализа рукописного фрагмента. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWord-strokePointers)|

## Методы

| Метод           | Возвращаемый тип    |Описание| Отзыв|
|:---------------|:--------|:----------|:-------|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWord-load)|

## Сведения о методе


### load(param: object)
Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.

#### Синтаксис
```js
object.load(param);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|param|object|Необязательный параметр. Принимает имена параметров и связей в виде строки с разделителями или массива. Либо укажите объект [loadOption](loadoption.md).|

#### Возвращаемое значение
void
### Примеры доступа к свойствам

**wordAlternates и languageId**
```js
OneNote.run(function (ctx) {        
    var app = ctx.application;
    
    // Gets the active page.
    var page = app.getActivePage();
    
    page.load('inkAnalysisOrNull/paragraphs/lines/words');
    return ctx.sync()
        .then(function() {
            var inkParagraphs = page.inkAnalysisOrNull.paragraphs;
            $.each(inkParagraphs.items, function(i, inkParagraph) {
                var inkLines = inkParagraph.lines;
                $.each(inkLines.items, function(j, inkLine) {
                    var inkWords = inkLine.words;
                    $.each(inkWords.items, function(k, inkWord) {
                    
                        // Log language Id of the word
                        console.log(inkWord.languageId);
                        
                        // Log every ink analyzed words.
                        $.each(inkWord.wordAlternates, function(l, word) {
                            console.log(word);                                  
                        })
                    })
                })
            })
        })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}); 
```