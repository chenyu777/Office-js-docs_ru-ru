# Объект PageContent (API JavaScript для OneNote)

_Относится к: OneNote Online_  


Представляет область на странице, содержащую контент верхнего уровня, например Outline или Image. Объекту PageContent можно назначить позицию по горизонтали и вертикали.

## Свойства

| Свойство     | Тип   |Описание|Отзывы|
|:---------------|:--------|:----------|:-------|
|id|string|Возвращает идентификатор объекта PageContent. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-id)|
|left|double|Получает или задает левую позицию (по оси X) объекта PageContent.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-left)|
|top|double|Получает или задает верхнюю позицию (по оси Y) объекта PageContent.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-top)|
|type|string|Получает тип объекта PageContent. Только для чтения. Возможные значения: Outline, Image, Other.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-type)|

## Связи
| Связь | Тип   |Описание| Отзывы|
|:---------------|:--------|:----------|:-------|
|image|[Изображение](image.md)|Получает объект Image в объекте PageContent. Вызывает исключение, если PageContentType не является Image. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-image)|
|ink|[FloatingInk](floatingink.md)|Получает рукописный фрагмент в объекте PageContent. Вызывает исключение, если PageContentType не является Ink. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-ink)|
|outline|[Структура](outline.md)|Получает элемент типа Outline в объекте PageContent. Вызывает исключение, если PageContentType не является Outline. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-outline)|
|parentPage|[Page](page.md)|Получает страницу, содержащую объект PageContent. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-parentPage)|

## Методы

| Метод           | Возвращаемый тип    |Описание| Отзыв|
|:---------------|:--------|:----------|:-------|
|[delete()](#delete)|void|Удаляет объект PageContent.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-delete)|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-load)|

## Сведения о методе


### delete()
Удаляет объект PageContent.

#### Синтаксис
```js
pageContentObject.delete();
```

#### Параметры
Нет

#### Возвращаемое значение
void

#### Примеры
```js
OneNote.run(function (context) {

    var page = context.application.getActivePage();
    var pageContents = page.contents;

    var firstPageContent = pageContents.getItemAt(0);
    firstPageContent.load('type');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            if(firstPageContent.isNull === false) {
                firstPageContent.delete();
                return context.sync();
            }
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
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
