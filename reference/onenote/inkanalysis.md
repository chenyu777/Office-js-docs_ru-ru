# Объект InkAnalysis (API JavaScript для OneNote)

_Применяется для OneNote Online_   


Представляет данные анализа рукописного фрагмента для заданного набора росчерков пера.

## Свойства

| Свойство     | Тип   |Описание|Отзывы|
|:---------------|:--------|:----------|:-------|
|id|string|Получает идентификатор объекта InkAnalysis. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysis-id)|

_См. [примеры](#примеры) доступа к свойствам._

## Связи
| Связь | Тип   |Описание| Отзыв|
|:---------------|:--------|:----------|:-------|
|page|[Page](page.md)|Возвращает родительский объект страницы. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysis-page)|
|paragraphs|[InkAnalysisParagraphCollection](inkanalysisparagraphcollection.md)|Возвращает объекты Paragraph после анализа рукописного фрагмента на этой странице. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysis-paragraphs)|

## Методы

| Метод           | Возвращаемый тип    |Описание| Отзыв|
|:---------------|:--------|:----------|:-------|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysis-load)|

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

**paragraphs**
```js
OneNote.run(function (ctx) {        
    var app = ctx.application;
    
    // Gets the active page.
    var page = app.getActivePage();
    
    // Load ink paragraphs.
    page.load('inkAnalysisOrNull/paragraphs');
    
    return ctx.sync()
        .then(function() {
            console.log(page.inkAnalysisOrNull.paragraphs.items.length);
        })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}); 
```