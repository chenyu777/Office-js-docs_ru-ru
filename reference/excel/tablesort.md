# Объект TableSort (API JavaScript для Excel)

_Относится к: Excel 2016, Excel Online, Excel для iOS, Office 2016_

Управляет операциями сортировки для объектов Table.

## Свойства

| Свойство     | Тип   |Описание
|:---------------|:--------|:----------|
|matchCase|Bool|Указывает, учитывался ли регистр при последней сортировке таблице. Только для чтения.|
|method|string|Указывает метод сортировки китайских символов, который использовался при последней сортировке таблицы. Только для чтения. Возможные значения: PinYin, StrokeCount.|

## Связи
| Связь | Тип   |Описание|
|:---------------|:--------|:----------|
|fields|[SortField](sortfield.md)|Указывает текущие условия, которые использовались при последней сортировке таблицы. Только для чтения.|

## Методы

| Метод           | Возвращаемый тип    |Описание|
|:---------------|:--------|:----------|
|[apply(fields: SortField[], matchCase: bool, method: string)](#applyfields-sortfield-matchcase-bool-method-string)|void|Выполняет сортировку.|
|[clear()](#clear)|void|Удаляет текущие параметры сортировки таблицы. При этом сбрасывается состояние кнопок в заголовках, но порядок сортировки таблицы остается неизменным.|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|
|[reapply()](#reapply)|void|Повторно применяет текущие параметры сортировки к таблице.|

## Сведения о методе


### apply(fields: SortField[], matchCase: bool, method: string)
Выполняет сортировку.

#### Синтаксис
```js
tableSortObject.apply(fields, matchCase, method);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|fields|SortField[]|Список условий для сортировки.|
|matchCase|Bool|Необязательный. Указывает, необходимо ли учитывать регистр при сортировке строк.|
|method|string|Необязательный. Метод сортировки, используемый для китайских символов.  Возможные значения: PinYin, StrokeCount|

#### Возвращаемое значение
void

#### Примеры
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.sort.apply([ 
            {
                key: 2,
                ascending: true
            },
        ], true);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### clear()
Удаляет текущие параметры сортировки таблицы. При этом сбрасывается состояние кнопок в заголовках, но порядок сортировки таблицы остается неизменным.

#### Синтаксис
```js
tableSortObject.clear();
```

#### Параметры
Нет

#### Возвращаемое значение
void

#### Примеры
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.sort.clear();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});

### load(param: object)
Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.

#### Syntax
```js
object.load(param);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|param|object|Необязательный параметр. Принимает имена параметров и связей в виде строки с разделителями или массива. Либо укажите объект [loadOption](loadoption.md).|

#### Возвращаемое значение
void

### reapply()
Повторно применяет текущие параметры сортировки к таблице.

#### Синтаксис
```js
tableSortObject.reapply();
```

#### Параметры
Нет

#### Возвращаемое значение
void

####Примеры
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.sort.reapply();   
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});