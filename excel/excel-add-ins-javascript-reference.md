# Справочник по API JavaScript для надстроек Excel

Область применения: Excel 2016, Office 2016

По ссылкам ниже описаны объекты Excel высокого уровня, доступные в API. Каждая ссылка на страницу объекта содержит описание свойств, связей и методов, доступных для объекта. Чтобы узнать больше, перейдите по этим ссылкам.
	
* [Workbook](resources/workbook.md). Объект верхнего уровня, содержащий связанные объекты книг, такие как листы, таблицы, диапазоны и т. д. Его также можно использовать для вывода списка связанных ссылок. 
* [Worksheet](resources/worksheet.md). Состоит в коллекции Worksheets. Коллекция Worksheets содержит все объекты Worksheet в книге.
	* [Worksheet Collection](resources/worksheetcollection.md). Коллекция всех объектов Worksheet, включенных в книгу. 
* [Range](resources/range.md). Представляет ячейку, строку, столбец или группу ячеек, содержащую один или несколько смежных блоков ячеек.  
* [Table](resources/table.md). Представляет коллекцию упорядоченных ячеек для упрощения управления данными. 
	* [Table Collection](resources/tablecollection.md). Коллекция таблиц в книге или листе. 
	* [TableColumn Collection](resources/tablecolumncollection.md). Коллекция всех столбцов в таблице. 
	* [TableRow Collection](resources/tablerowcollection.md). Коллекция всех строк в таблице. 
* [Chart](resources/chart.md). Представляет объект Chart в листе, который является визуальным представлением базовых данных.   
	* [Chart Collection](resources/chartcollection.md). Коллекция диаграмм в листе.	
* [NamedItem](resources/nameditem.md). Представляет определенное имя для диапазона ячеек или значение. Имена могут быть простыми именованными объектами, объектом диапазона и т. д.
	* [NamedItem Collection](resources/nameditemcollection.md). Коллекция объектов NamedItem в книге.
* [Binding](resources/binding.md). Абстрактный класс, представляющий привязку к разделу книги.
	* [Binding Collection](resources/bindingcollection.md). Коллекция всех объектов Binding, включенных в книгу. 
* [TrackedObject Collection](resources/trackedobjectscollection.md). Позволяет надстройкам управлять ссылкой на объект диапазона в пакетах sync(). 
* [Request Context](resources/requestcontext.md). Объект RequestContext упрощает отправку запросов приложению Excel.


##### Дополнительные ресурсы

*  [Общие сведения о программировании надстроек Excel](excel-add-ins-programming-overview.md)
*  [Создание первой надстройки Excel](build-your-first-excel-add-in.md)
*  [Обозреватель фрагментов кода для Excel](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
*  [Примеры кода для надстроек Excel](excel-add-ins-code-samples.md) 


