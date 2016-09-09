# Объект WorksheetProtectionOptions (API JavaScript для Excel)

_Относится к: Excel 2016, Excel Online, Excel для iOS, Office 2016_

Представляет параметры защиты листа.

## Свойства

| Свойство     | Тип   |Описание
|:---------------|:--------|:----------|
|allowAutoFilter|bool|Представляет параметр защиты листа, разрешающий использовать функцию автофильтрации.|
|allowDeleteColumns|bool|Представляет параметр защиты листа, разрешающий удалять столбцы.|
|allowDeleteRows|bool|Представляет параметр защиты листа, разрешающий удалять строки.|
|allowFormatCells|bool|Представляет параметр защиты листа, разрешающий форматировать ячейки.|
|allowFormatColumns|bool|Представляет параметр защиты листа, разрешающий форматировать столбцы.|
|allowFormatRows|bool|Представляет параметр защиты листа, разрешающий форматировать строки.|
|allowInsertColumns|bool|Представляет параметр защиты листа, разрешающий вставлять столбцы.|
|allowInsertHyperlinks|bool|Представляет параметр защиты листа, разрешающий вставлять гиперссылки.|
|allowInsertRows|bool|Представляет параметр защиты листа, разрешающий вставлять строки.|
|allowPivotTables|bool|Представляет параметр защиты листа, разрешающий использовать функцию сводных таблиц.|
|allowSort|bool|Представляет параметр защиты листа, разрешающий использовать функцию сортировки.|

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

#### Примеры
В этом примере загружаются параметры защиты активного листа.
```js
Excel.run(function (ctx) {
    var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
    worksheet.protection.load();            
    return ctx.sync()
        .then(function () {
            console.log("Active worksheet's protection options: " + worksheet.protection.options);
        });
})
.catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
