# Объект Filter (API JavaScript для Excel)

_Относится к: Excel 2016, Excel Online, Excel для iOS, Office 2016_

Управляет фильтрацией столбца таблицы.

## Свойства

Нет

## Связи
| Связь | Тип|Описание|
|:---------------|:--------|:----------|
|criteria|[FilterCriteria](filtercriteria.md)|Текущий фильтр, заданный для определенного столбца. Только для чтения.|

## Методы

| Метод   | Возвращаемый тип|Описание|
|:---------------|:--------|:----------|
|[apply(criteria: FilterCriteria)](#applycriteria-filtercriteria)|void|Применяет заданные условия фильтра для определенного столбца. Для этой задачи можно воспользоваться любым из указанных ниже вспомогательных методов.|
|[applyBottomItemsFilter(count: number)](#applybottomitemsfiltercount-number)|void|Применяет к столбцу фильтр по количеству элементов снизу.|
|[applyBottomPercentFilter(percent: number)](#applybottompercentfilterpercent-number)|void|Применяет к столбцу фильтр по проценту элементов снизу.|
|[applyCellColorFilter(color: string)](#applycellcolorfiltercolor-string)|void|Применяет к столбцу фильтр по цвету ячейки.|
|[applyCustomFilter(criteria1: string, criteria2: string, oper: FilterOperator)](#applycustomfiltercriteria1-string-criteria2-string-oper-filteroperator)|void|Применяет к столбцу фильтр по условиям.|
|[applyDynamicFilter(criteria: string)](#applydynamicfiltercriteria-string)|void|Применяет к столбцу динамический фильтр.|
|[applyFontColorFilter(color: string)](#applyfontcolorfiltercolor-string)|void|Применяет к столбцу фильтр по цвету шрифта.|
|[applyIconFilter(icon: Icon)](#applyiconfiltericon-icon)|void|Применяет к столбцу фильтр по значку.|
|[applyTopItemsFilter(count: number)](#applytopitemsfiltercount-number)|void|Применяет к столбцу фильтр по количеству элементов сверху.|
|[applyTopPercentFilter(percent: number)](#applytoppercentfilterpercent-number)|void|Применяет к столбцу фильтр по проценту элементов сверху.|
|[applyValuesFilter(values: ()[])](#applyvaluesfiltervalues-)|void|Применяет к столбцу фильтр по значениям.|
|[clear()](#clear)|void|Сбрасывает фильтр для определенного столбца.|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|

## Сведения о методе


### apply(criteria: FilterCriteria)
Применяет заданные условия фильтра для определенного столбца. Для этой задачи можно воспользоваться любым из указанных ниже вспомогательных методов. 

#### Синтаксис
```js
filterObject.apply(criteria);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|criteria|FilterCriteria|Условия, которые необходимо применить.|

#### Возвращаемое значение
void

#### Пример
В приведенном ниже примере показано, как применить настраиваемый фильтр с помощью универсального метода apply().

```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    var filterCriteria = { 
		filterOn: Excel.FilterOn.custom,
		criterion1: ">50",
		operator: Excel.FilterOperator.and,
		criterion2: "<100"
    	} 
    column.filter.apply(filterCriteria);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyBottomItemsFilter(count: number)
Применяет к столбцу фильтр по количеству элементов снизу.

#### Синтаксис
```js
filterObject.applyBottomItemsFilter(count);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|count|number|Количество элементов снизу, которые должны отображаться.|

#### Возвращаемое значение
void

#### Пример
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyBottomItemsFilter(3);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyBottomPercentFilter(percent: number)
Применяет к столбцу фильтр по проценту элементов снизу.

#### Синтаксис
```js
filterObject.applyBottomPercentFilter(percent);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|percent|number|Процент элементов снизу, которые должны отображаться.|

#### Возвращаемое значение
void

#### Пример
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyBottomPercentFilter(30);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
### applyCellColorFilter(color: string)
Применяет к столбцу фильтр по цвету ячейки.


#### Синтаксис
```js
filterObject.applyCellColorFilter(color);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|color|string|Цвет фона ячеек, которые должны отображаться.|

#### Возвращаемое значение
void

#### Пример
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyCellColorFilter('red');
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyCustomFilter(criteria1: string, criteria2: string, oper: FilterOperator)
Применяет к столбцу фильтр по условиям.

#### Синтаксис
```js
filterObject.applyCustomFilter(criteria1, criteria2, oper);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|criteria1|string|Строка первого условия.|
|criteria2|string|Необязательный. Строка второго условия.|
|oper|FilterOperator|Необязательный. Оператор, который описывает способ объединения двух условий.|

#### Возвращаемое значение
void


#### Пример
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyCustomFilter('>50','<100','and');
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyDynamicFilter(criteria: string)
Применяет к столбцу динамический фильтр.

#### Синтаксис
```js
filterObject.applyDynamicFilter(criteria);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|criteria|string|Динамические условия, которые необходимо применить.  Возможные значения: Unknown, AboveAverage, AllDatesInPeriodApril, AllDatesInPeriodAugust, AllDatesInPeriodDecember, AllDatesInPeriodFebruray, AllDatesInPeriodJanuary, AllDatesInPeriodJuly, AllDatesInPeriodJune, AllDatesInPeriodMarch, AllDatesInPeriodMay, AllDatesInPeriodNovember, AllDatesInPeriodOctober, AllDatesInPeriodQuarter1, AllDatesInPeriodQuarter2, AllDatesInPeriodQuarter3, AllDatesInPeriodQuarter4, AllDatesInPeriodSeptember, BelowAverage, LastMonth, LastQuarter, LastWeek, LastYear, NextMonth, NextQuarter, NextWeek, NextYear, ThisMonth, ThisQuarter, ThisWeek, ThisYear, Today, Tomorrow, YearToDate, Yesterday|

#### Возвращаемое значение
void

#### Пример
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyDynamicFilter(Excel.DynamicFilterCriteria.aboveAverage);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyFontColorFilter(color: string)
Применяет к столбцу фильтр по цвету шрифта.

#### Синтаксис
```js
filterObject.applyFontColorFilter(color);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|color|string|Цвет шрифта ячеек, которые должны отображаться.|

#### Возвращаемое значение
void

#### Пример
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyFontColorFilter('red');
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyIconFilter(icon: Icon)
Применяет к столбцу фильтр по значку.

#### Синтаксис
```js
filterObject.applyIconFilter(icon);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|icon|Icon|Значки ячеек, которые должны отображаться.|

#### Возвращаемое значение
void

#### Пример
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyIconFilter(Excel.icons.fiveArrows.yellowDownInclineArrow);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyTopItemsFilter(count: number)
Применяет к столбцу фильтр по количеству элементов сверху.

#### Синтаксис
```js
filterObject.applyTopItemsFilter(count);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|count|number|Количество элементов сверху, которые должны отображаться.|

#### Возвращаемое значение
void

#### Пример
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyTopItemsFilter(3);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### applyTopPercentFilter(percent: number)
Применяет к столбцу фильтр по проценту элементов сверху.

#### Синтаксис
```js
filterObject.applyTopPercentFilter(percent);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|percent|number|Процент элементов сверху, которые должны отображаться.|

#### Возвращаемое значение
void

#### Пример
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyTopPercentFilter(30);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
### applyValuesFilter(values: ()[])
Применяет к столбцу фильтр по значениям.

#### Синтаксис
```js
filterObject.applyValuesFilter(values);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|values|()[]|Список значений, которые должны отображаться.|

#### Возвращаемое значение
void

#### Пример
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyValuesFilter(['a','b']);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### clear()
Сбрасывает фильтр для определенного столбца.

#### Синтаксис
```js
filterObject.clear();
```

#### Параметры
Нет

#### Возвращаемое значение
void

#### Пример
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.clear();
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
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|param|object|Необязательный параметр. Принимает имена параметров и связей в виде строки с разделителями или массива. Либо укажите объект [loadOption](loadoption.md).|

#### Возвращаемое значение
void

