# Объект ChartFont (API JavaScript для Excel)

Этот объект представляет атрибуты шрифта (название, размер, цвет и т. д.) для объекта диаграммы.

## Properties

| Свойство     | Тип   |Описание
|:---------------|:--------|:----------|
|bold|bool|Представляет параметр полужирного шрифта.|
|color|string|HTML-код цвета текста. Например, #FF0000 обозначает красный цвет.|
|italic|bool|Представляет параметр курсива.|
|name|string|Название шрифта, например Calibri.|
|size|double|Размер шрифта, например 11.|
|underline|string|Тип подчеркивания, применяемый для шрифта. Возможные значения: None, Single.|

_Ознакомьтесь с [примерами](#примерами) доступа к свойствам._

## Связи
Нет


## Методы

| Метод           | Возвращаемый тип    |Описание|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|

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

В качестве примера используйте заголовок диаграммы.

```js
Excel.run(function (ctx) { 
    var title = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").title;
    title.format.font.name = "Calibri";
    title.format.font.size = 12;
    title.format.font.color = "#FF0000";
    title.format.font.italic =  false;
    title.format.font.bold = true;
    title.format.font.underline = false;
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Отформатируйте заголовок диаграммы полужирным красным шрифтом Calibri размера 10. 

```js
Excel.run(function (ctx) { 
    var title = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").title;
    title.format.font.name = "Calibri";
    title.format.font.size = 12;
    title.format.font.color = "#FF0000";
    title.format.font.italic =  false;
    title.format.font.bold = true;
    title.format.font.underline = false;
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
