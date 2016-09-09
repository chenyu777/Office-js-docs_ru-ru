# Справочник по API JavaScript для Excel

Вы можете использовать API JavaScript для Excel, чтобы создавать надстройки для Excel 2016. Ниже перечислены объекты Excel высокого уровня, доступные в API. Каждая ссылка на страницу объекта содержит описание свойств, связей и методов, доступных для объекта. Чтобы узнать больше, перейдите по соответствующим ссылкам.

* [Workbook](../../reference/excel/workbook.md) — объект верхнего уровня, содержащий связанные объекты книг, такие как листы, таблицы, диапазоны и т. д. Его также можно использовать для вывода списка связанных ссылок.
* [Worksheet](../../reference/excel/worksheet.md)— элемент коллекции Worksheets. Коллекция Worksheets содержит все объекты Worksheet в книге.
    * [Worksheet Collection](../../reference/excel/worksheetcollection.md) коллекция всех объектов Worksheet, включенных в книгу.
* [Range](../../reference/excel/range.md): ячейка, строка, столбец или группа ячеек, содержащая один или несколько смежных блоков ячеек.
* [Table](../../reference/excel/table.md): коллекция упорядоченных ячеек для упрощения управления данными.
    * [Table Collection](../../reference/excel/tablecollection.md): коллекция таблиц в книге или листе.
    * [TableColumn Collection](../../reference/excel/tablecolumncollection.md): коллекция всех столбцов в таблице.
    * [TableRow Collection](../../reference/excel/tablerowcollection.md): коллекция всех строк в таблице.
* [Chart](../../reference/excel/chart.md): объект Chart в листе, который является визуальным представлением базовых данных.
    * [Chart Collection](../../reference/excel/chartcollection.md): коллекция диаграмм на листе.
* [TableSort](../../reference/excel/tablesort.md): объект, который управляет операциями сортировки в объектах Table.
* [RangeSort](../../reference/excel/rangesort.md): объект, который управляет операциями в объектах Range.
* [Filter](../../reference/excel/filter.md): объект, который управляет фильтрацией столбца таблицы.
* [Worksheet Protection](../../reference/excel/worksheetprotection.md): защита объекта листа.
* [Worksheet Function](../../reference/excel/functions.md): контейнер для функций листа Microsoft Excel, которые можно вызывать с помощью JavaScript.
* [NamedItem](../../reference/excel/nameditem.md): определенное имя для диапазона ячеек или значение. Имена могут быть простыми именованными объектами, объектом диапазона и т. д.
    * [NamedItem Collection](../../reference/excel/nameditemcollection.md): коллекция объектов NamedItem в книге.
* [Binding](../../reference/excel/binding.md): абстрактный класс, представляющий привязку к разделу книги.
    * [Binding Collection](../../reference/excel/bindingcollection.md): коллекция всех объектов Binding, включенных в книгу.
* [TrackedObject Collection](../../reference/excel/trackedobjectscollection.md) позволяет надстройкам управлять ссылкой на объект диапазона в пакетах sync().
* [Request Context](../../reference/excel/requestcontext.md) упрощает отправку запросов приложению Excel.


##### Дополнительные ресурсы

*  [Общие сведения о программировании надстроек Excel](excel-add-ins-javascript-programming-overview.md)
*  [Создание первой надстройки Excel](build-your-first-excel-add-in.md)
*  [Обозреватель фрагментов кода для Excel](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)

