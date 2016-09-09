# Объект FilterCriteria (API JavaScript для Excel)

_Относится к: Excel 2016, Excel Online, Excel для iOS, Office 2016_

Представляет критерии фильтрации, применяемые к столбцу.

## Свойства

| Свойство     | Тип   |Описание
|:---------------|:--------|:----------|
|color|string|Строка цвета HTML, которая используется для фильтрации ячеек. Используется с фильтрацией типа "cellColor" и "fontColor".|
|criterion1|string|Первый критерий фильтрации данных. Используется в качестве оператора при фильтрации типа "custom".|
|criterion2|string|Второй критерий фильтрации данных. Используется в качестве оператора только при фильтрации типа "custom".|
|dynamicCriteria|string|Динамические критерии из набора Excel.DynamicFilterCriteria, которые необходимо применить к этому столбцу. Используется с фильтрацией типа "dynamic". Возможные значения: Unknown, AboveAverage, AllDatesInPeriodApril, AllDatesInPeriodAugust, AllDatesInPeriodDecember, AllDatesInPeriodFebruray, AllDatesInPeriodJanuary, AllDatesInPeriodJuly, AllDatesInPeriodJune, AllDatesInPeriodMarch, AllDatesInPeriodMay, AllDatesInPeriodNovember, AllDatesInPeriodOctober, AllDatesInPeriodQuarter1, AllDatesInPeriodQuarter2, AllDatesInPeriodQuarter3, AllDatesInPeriodQuarter4, AllDatesInPeriodSeptember, BelowAverage, LastMonth, LastQuarter, LastWeek, LastYear, NextMonth, NextQuarter, NextWeek, NextYear, ThisMonth, ThisQuarter, ThisWeek, ThisYear, Today, Tomorrow, YearToDate, Yesterday.|
|filterOn|string|Свойство, с помощью которого фильтр определяет, следует ли показывать значения. Возможные значения:    BottomItems, BottomPercent, CellColor, Dynamic, FontColor, Values, TopItems, TopPercent, Icon, Custom |
|values|object[]|Набор значений, который используется при фильтрации типа "values".|

## Связи
| Связь | Тип   |Описание|
|:---------------|:--------|:----------|
|icon|[Значок](icon.md)|Значок, используемый для фильтрации ячеек. Используется с фильтрацией типа "icon".|
|operator|[FilterOperator](filteroperator.md)|Оператор, который используется для объединения критериев 1 и 2 при фильтрации типа "custom".|

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
