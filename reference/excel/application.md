# Объект Application (API JavaScript для Excel)

Представляет приложение Excel, которое управляет книгой.

## Свойства

| Свойство     | Тип   |Описание
|:---------------|:--------|:----------|
|calculationMode|string|Возвращает режим вычисления, который используется в книге. Только для чтения. Возможные значения: `Automatic` (Excel контролирует пересчет), `AutomaticExceptTables` (Excel контролирует пересчет, но не учитывает изменения в таблицах), `Manual` (вычисление выполняется по запросу пользователя).|

_Ознакомьтесь с [примерами](#примерами) доступа к свойствам._

## Связи
Нет


## Методы

| Метод           | Возвращаемый тип    |Описание|
|:---------------|:--------|:----------|
|[calculate(calculationType: string)](#calculatecalculationtype-string)|void|Пересчитывает данные во всех открытых в текущий момент книгах Excel.|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|

## Сведения о методе


### calculate(calculationType: string)
Пересчитывает данные во всех открытых в текущий момент книгах Excel.

#### Синтаксис
```js
applicationObject.calculate(calculationType);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|calculationType|string|Определяет тип расчета, который нужно использовать. Возможные значения: `Recalculate` (параметр по умолчанию, выполняется обычное вычисление по всем формулам в книге), `Full` (принудительное полное вычисление данных), `FullRebuild` (принудительное полное вычисление данных и пересоздание зависимостей).|

#### Возвращаемое значение
void

#### Примеры
```js
Excel.run(function (ctx) { 
    ctx.workbook.application.calculate('Full');
    return ctx.sync(); 
}).catch(function(error) {
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
|param|object|Необязательный параметр. Принимает имена параметров и связей в виде строки с разделителями или массива. Этот параметр также принимает объект [loadOption](loadoption.md).|

#### Возвращаемое значение
void
### Примеры доступа к свойствам
```js
Excel.run(function (ctx) { 
    var application = ctx.workbook.application;
    application.load('calculationMode');
    return ctx.sync().then(function() {
        console.log(application.calculationMode);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

