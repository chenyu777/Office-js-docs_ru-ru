# Объект InkAnalysisLine (API JavaScript для OneNote)

_Применяется для OneNote Online_   


Представляет данные анализа рукописного фрагмента для определенной текстовой строки, созданной росчерками пера.

## Свойства

| Свойство     | Тип   |Описание|Отзывы|
|:---------------|:--------|:----------|:-------|
|id|string|Получает идентификатор объекта InkAnalysisLine. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLine-id)|

_См. [примеры](#примеры) доступа к свойствам._

## Связи
| Связь | Тип   |Описание| Отзывы|
|:---------------|:--------|:----------|:-------|
|paragraph|[InkAnalysisParagraph](inkanalysisparagraph.md)|Ссылка на родительский элемент InkAnalysisParagraph. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLine-paragraph)|
|words|[InkAnalysisWordCollection](inkanalysiswordcollection.md)|Возвращает слова анализа рукописного фрагмента в этой строке анализа рукописного фрагмента. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLine-words)|

## Методы

| Метод           | Возвращаемый тип    |Описание| Отзыв|
|:---------------|:--------|:----------|:-------|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLine-load)|

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

**words**
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
                    // Word counts in a line.
                    console.log(inkLine.words.items.length);
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