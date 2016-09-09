# Параметры загрузки объектов (API JavaScript для Excel)

Представляет объект, который может быть передан методу загрузки, чтобы задать набор свойств и связей, загружаемых при выполнении метода sync(), который синхронизирует состояние объектов Excel и соответствующих прокси-объектов JavaScript в надстройке. Принимаются такие параметры, как select и expand. Они указывают набор свойств, загружаемых в объект, а также обеспечивают управление разбивкой на страницы в коллекции.

Кроме того, можно указать строку, содержащую свойства и связи, которые необходимо загрузить, или указать массив, содержащий список свойств и связей для загрузки. См. пример ниже.

```js   
object.load  ('<var1>,<relation1/var2>');

// Pass the parameter as an array.
object.load (["var1", "relation1/var2"]);
```

## Свойства
| Свойство     | Тип   |Описание|
|:---------------|:--------|:----------|
|select|object|Укажите разделенный запятыми список или массив имен параметров и связей, которые необходимо загрузить при вызове метода executeAsync, например "property1, relationship1", ["property1", "relationship1"]. Необязательный параметр.|
|expand|object|Укажите разделенный запятыми список или массив имен связей, которые необходимо загрузить при вызове метода executeAsync, например "relationship1, relationship2", ["relationship1", "relationship2"]. Необязательный параметр.|
|top|int| Укажите количество элементов в запрашиваемой коллекции, которые будут включены в результат. Необязательный параметр.|
|skip|int|Укажите количество элементов в коллекции, которые необходимо пропустить и исключить из результата. Если указан параметр `top`, выделение результата начнется после пропуска заданного числа элементов. Необязательный параметр.|

#### Примеры

В этом примере выбираются первые 100 строк таблицы.

```js
Excel.run(function (ctx) { 
    var table = ctx.workbook.tables.getItem("Table1");
    var tableRows = table.rows.load({"select" : "index, values","top": 100, "skip": 0 })
    return ctx.sync().then(function() {
        for (var i = 0; i < tableRows.items.length; i++)
        {
            console.log(tableRows.items[i].index);
            console.log(tableRows.items[i].values);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
})
```
