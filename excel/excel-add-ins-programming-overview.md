# Общие сведения о программировании надстроек Excel с помощью JavaScript

Область применения: Excel 2016, Office 2016

В данной статье рассматриваются основы использования API JavaScript для создания надстроек в Excel 2016. Подробные спецификации API JavaScript для Excel см. на странице со [справочными сведениями](excel-add-ins-javascript-reference.md).

## Основы

Начнем с краткого описания основных понятий, например RequestContext, прокси-объекты JavaScript, sync(), Excel.run() и load(), имеющих ключевое значение при использовании API. В примере кода в конце данного раздела показано, как использовать эти объекты и методы.


#### RequestContext

Объект RequestContext упрощает отправку запросов приложению Excel. Так как надстройка Office и приложение Excel выполняются в виде двух различных процессов, для получения доступа из надстройки к Excel и связанным объектам, например листам, таблицам и т. д, необходим контекст запроса. Процесс создания контекста запроса показан ниже.

```js
var ctx = new Excel.RequestContext();
```

#### Прокси-объекты 

Объекты JavaScript Excel, объявленные и использованные в надстройке, — это прокси-объекты для реальных объектов в документе Excel. Никакие действия над прокси-объектами не реализуются в Excel, а состояние документа Excel не реализуется на прокси-объектах до его синхронизации. Состояние документа синхронизируется при выполнении метода context.sync(). (См. ниже). 

Например, локальный объект JavaScript `selectedRange` объявлен в качестве ссылки на выбранный диапазон. Это можно использовать для постановки в очередь настройки его свойств и вызова методов. Действия над такими объектами не реализуются до выполнения метода sync(). 

```js
var selectedRange = ctx.workbook.getSelectedRange();
```    

#### sync()

Метод sync(), доступный в контексте запроса, синхронизирует состояние прокси-объектов JavaScript и реальных объектов в Excel путем выполнения поставленных в очередь инструкций над контекстом и получения свойств загруженных объектов Office для их использования в коде. Этот метод возвращает обещание, которое выполняется после завершения синхронизации.

#### Excel.run(function(context) { batch })

Метод Excel.run() выполняет пакетный сценарий, выполняющий действия над моделью объекта Excel. Пакетные команды включают определения локальных прокси-объектов JavaScript и методов sync(), синхронизирующих состояние локальных объектов и объектов Excel, а также выполнение обещания. Преимущество пакетной обработки запросов в Excel.run() в том, что при выполнении обещания любые отслеживаемые объекты диапазона, выделенные во время выполнения, автоматически отпускаются. 

Выполняемый метод использует объект RequestContext и возвращает обещание (как правило, просто результат метода ctx.sync()). Пакетную операцию можно выполнить вне метода Excel.run(). Однако при таком сценарии любые ссылки на объекты диапазона требуют отслеживания и управления вручную. 

#### load()

Метод load() используется для заполнения прокси-объектов, созданных на уровне JavaScript надстройки. При попытке получения объекта, например листа, сначала на уровне JavaScript создается локальный прокси-объект. Такой объект можно использовать для постановки в очередь настройки его свойств и методов вызова. Но для чтения свойств или связей объекта сначала необходимо вызвать методы load() и sync(). Метод load() использует свойства и связи, которые требуется загрузить при вызове метода sync(). 

Синтаксис:

```js
object.load(string: properties);
//or 
object.load(array: properties);
//or
object.load({loadOption});
```
Где: 

* `properties` — это список имен свойств и (или) связей, которые требуется загрузить, указанных в виде строк с разделителями-запятыми или массива имен. Дополнительные сведения см. в методах .load() под каждым объектом.
* `loadOption` указывает объект, описывающий параметры "выбрать", "развернуть", "сверху" и "пропустить". Дополнительные сведения см. в разделе [параметров](resources/loadoption.md) объекта load.

##### Пример

Ниже приводится пример использования всех рассмотренных понятий. В примере кода показана запись значений из массива в объект диапазона. 

Метод Excel.run() содержит пакет инструкций. В рамках этого пакета создается прокси-объект, который ссылается на диапазон (адрес A1:B2) на активном листе. Значение этого прокси-объекта диапазона устанавливается локально. Для обратного прочтения значений свойство `text` диапазона загружается в прокси-объект. Все эти команды ставятся в очередь и выполняются при вызове метода ctx.sync(). Метод sync() возвращает обещание, с помощью которого его можно связать с другими операциями.

```js
// Run a batch operation against the Excel object model. Use the context argument to get access to the Excel document.
Excel.run(function (ctx) { 

	// Create a proxy object for the sheet
	var sheet = ctx.workbook.worksheets.getActiveWorksheet();
	// Values to be updated
	var values = [
				 ["Type", "Estimate"],
				 ["Transportation", 1670]
				 ];
	// Create a proxy object for the range
	var range = sheet.getRange("A1:B2");

	// Assign array value to the proxy object's values property.
	range.values = values;
	
	// Queue a command to load the text property for the proxy range object.	
	range.load('text');

	// Synchronizes the state between JavaScript proxy objects and real objects in Excel by executing instructions queued on the context 
	return ctx.sync().then(function() {
			console.log("Done");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

##### Пример

В следующем примере показано, как скопировать значения из диапазона от A1:A2 до B1:B2 активного листа, используя метод load() на объекте диапазона. 

```js
// Run a batch operation against the Excel object model. Use the context argument to get access to the Excel document.
Excel.run(function (ctx) { 

	// Create a proxy object for the range
	var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A1:A2");

	// Queue a command to load the following properties on the proxy range object.	
	range.load ("address, values, range/format"); 

	// Synchronizes the state between JavaScript proxy objects and real objects in Excel by executing instructions queued on the context 
	return ctx.sync().then(function() {
		// Assign the previously loaded values to the new range proxy object. The values will be updated once the following .then() function is invoked. 
		ctx.workbook.worksheets. getActiveWorksheet().getRange("B1:B2").values= range.values;
	});
}).then(function() {
	  console.log("done");
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### Выбор свойств и связей 

По умолчанию метод object.load() выбирает все скалярные и сложные свойства загружаемого объекта. По умолчанию связи не загружаются (например, формат — это объект связи объекта Range). Однако рекомендуется всегда явно помечать свойства и связи, которые требуется загрузить, чтобы увеличить производительность. Для этого нужно указать (в параметре `load()`) подмножество свойств и связей, которые требуется включить в ответ. Метод load() поддерживает два указанных ниже типа входных данных.

* Имена свойств и связей в виде имен строк с разделителями-запятыми _или_ массива строк с именами свойств или связей. 
* Объект, описывающий параметры "выбрать", "развернуть", "сверху" и "пропустить". Дополнительные сведения см. в разделе [параметров](resources/loadoption.md) объекта load.

```js	
object.load  ('<var1>,<relation1/var2>');

// Pass the parameter as an array.
object.load (["var1", "relation1/var2"]);
```

#### Примеры

Приведенная ниже инструкция загрузки загружает все свойства диапазона, а затем добавляет информацию для формата и формата/заливки.  
 
```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:B2"; 
	var myRange = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	
	myRange.load(["address", "format/*", "format/fill", "entireRow" ]);
	return ctx.sync().then(function() {
		console.log (myRange.address); //ok
		console.log (myRange.format.wrapText); //ok
		console.log (myRange.format.fill.color); //ok
		//console.log (myRange.format.font.color); //not ok as it was not loaded

	});
}).then(function() {
	  console.log("done");
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### Null-Input

#### Входное значение null в двумерном массиве


            Входное значение `null` в двумерном массиве (для значений, числового формата, формулы) игнорируется при обновлении API. Предполагаемый целевой объект не будет обновлен, если входное значение `null` отправлено в виде значений, числового формата и сетки значений формулы.

Пример: чтобы обновить только определенные фрагменты диапазона, такие как числовой формат ячейки, и сохранить существующий числовой формат в других фрагментах диапазона, установите требуемый числовой формат в нужных фрагментах и отправьте значение `null` для других ячеек. 

В запросе на присваивание ниже значения присваиваются лишь некоторым фрагментам числового формата диапазона, в то время как в остальных фрагментах сохраняется существующий числовой формат (путем отправки значения null).

```js
  range.values = [["Eurasia", "29.96", "0.25", "15-Feb" ]];
  range.numberFormat = [[null, null, null, "m/d/yyyy;@"]];
```
#### Входное значение null для свойства

`null` не может быть допустимым входным значением для всего свойства. Например, следующий пример недопустим, так как всем значениям нельзя присваивать значение null или игнорировать их. 

```js
 range.values= null;

```

Указанный ниже пример также не будет работать, так как для цвета недопустимо использовать значение null. 

```js
 range.format.fill.color =  null;
```

### Null-Response

Представление свойств форматирования, состоящее из неоднородных значений, приведет к возврату значения null в отклике. 

Пример: диапазон может состоять из одной или нескольких ячеек. Если отдельные ячейки в указанном диапазоне не содержат однородных значений форматирования, представление уровня диапазона будет неопределенным. 

```js
  "size" : null,
  "color" : null,
```

### Пустые входные и выходные данные

Пустые значения в запросах на обновление считаются указанием для очистки или сброса соответствующего свойства. Пустое значение представляется двумя двойными кавычками, не разделенными пробелом. `""`

Пример. 
* Для `values` значение диапазона очищено. Это аналогично очистке содержимого в приложении.
* Для `numberFormat` числовому формату присвоено значение `General`.
* Для `formula` и `formulaLocale` значения формулы очищены. 

При операциях чтения будьте готовы получать пустые значения, если в ячейках нет содержимого. Если ячейка не содержит данных или значений, API возвращает пустое значение. Пустое значение представляется двумя двойными кавычками, не разделенными пробелом. `""`

```js
  range.values = [["", "some", "data", "in", "other", "cells", ""]];
```

```js
  range.formula = [["", "", "=Rand()"]];
```

### Неограниченный диапазон

#### Чтение

Адрес неограниченного диапазона содержит только идентификаторы столбцов и строк, а также идентификаторы неопределенных строк и столбцов (соответственно), например:

* `C:C`, `A:F`, `A:XFD` (содержит неопределенные строки)
* `2:2`, `1:4`, `1:1048546` (содержит неопределенные столбцы)

Когда API отправляет запрос на получение неограниченного диапазона (например, `getRange('C:C')`), отклик содержит `null` для свойств уровня ячеек, например `values`, `text`, `numberFormat`, `formula` и т. д. Другие свойства диапазона, такие как `address`, `cellCount` и т. д., отражают неограниченный диапазон.

#### Запись

Задание свойств уровня ячеек (например, значений, числового формата и т. д.) для неограниченного диапазона **не допускается**, так как запрос на ввод может оказаться слишком большим для обработки. 

Пример: приведенный ниже запрос на обновление значений недопустим, поскольку запрашиваемый диапазон не ограничен. 

```js
...
	var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A:B");
	range.values = 'Due Date';
...
```

При попытке выполнения операции обновления с таким диапазоном API возвратит ошибку.


### Большой диапазон

Большой диапазон — это диапазон, размер которого слишком велик для одного вызова API. Множество факторов, например количество ячеек, значений, числовых форматов, формул и т. д. в составе диапазона, могут сделать запрос настолько большим, что он станет неподходящим для взаимодействия с API. API делает все возможное для возврата запрашиваемых данных или записи в них. Но обработка крупного запроса может привести к ошибке API из-за чрезмерного использования ресурсов. 

Чтобы избежать этого, рекомендуется выполнять операции чтения и записи с использованием нескольких диапазонов меньшего размера.


### Копирование одного входного значения

Для поддержки обновления диапазона с использованием одинаковых значений или числового формата либо для применения одной и той же формулы ко всему диапазону в установленном интерфейсе API используется следующее соглашение. В Excel этот принцип аналогичен вводу значений или формул в диапазон в режиме CTRL+ВВОД. 

API ищет *значение одной ячейки* и, если размер целевого диапазона не соответствует размеру входного диапазона, обновление применяется ко всему диапазону в режиме CTRL+ВВОД с использованием значения или формулы в запросе.

#### Примеры

Приведенный ниже запрос добавляет в выбранный диапазон текст "Due Date". Обратите внимание, что диапазон содержит 20 ячеек, в то время как входные данные — значение лишь для одной ячейки.

```js
Excel.run(function (ctx) { 
	var sheetName = 'Sheet1';
	var rangeAddress = 'A1:A20';
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	range.values = 'Due Date';
	range.load('text');
	return ctx.sync().then(function() {
		console.log(range.text);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Указанный ниже запрос добавляет в выбранный диапазон дату "11.03.2015".

```js
Excel.run(function (ctx) { 
	var sheetName = 'Sheet1';
	var rangeAddress = 'A1:A20';
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	range.numberFormat = 'm/d/yyyy';
	range.values = '3/11/2015';
	range.load('text');
	return ctx.sync().then(function() {
		console.log(range.text);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
Следующий запрос добавляет в выбранный запрос формулу, которая применяется ко всему диапазону в режиме CTRL+Enter.  

```js
Excel.run(function (ctx) { 
	var sheetName = 'Sheet1';
	var rangeAddress = 'A1:A20';
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	range.numberFormat = 'm/d/yyyy';
	range.values = '3/11/2015';
	range.load('text');
	return ctx.sync().then(function() {
		console.log(range.text);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### Сообщения об ошибках

Ошибки возвращаются с помощью объекта ошибки, состоящего из кода и сообщения. В таблице ниже перечислены возможные ошибки. 

|error.code | error.message |
|:----------|:--------------|
|InvalidArgument |Аргумент недопустим, отсутствует или имеет неправильный формат.|
|InvalidRequest  |Не удается обработать запрос.|
|InvalidReference|Эта ссылка недопустима для текущей операции.|
|InvalidBinding  |Эта привязка объектов недопустима из-за предыдущих обновлений.|
|InvalidSelection|Выбранный фрагмент недопустим для этой операции.|
|Unauthenticated |Требуемые сведения о проверке подлинности отсутствуют или недопустимы.|
|AccessDenied|Вы не можете выполнить запрашиваемую операцию.|
|ItemNotFound|Запрашиваемый ресурс не существует.|
|ActivityLimitReached|Достигнут предел действий.|
|GeneralException|При обработке запроса возникла внутренняя ошибка.|
|NotImplemented  |Запрашиваемая функция не реализована.|
|ServiceNotAvailable|Служба недоступна.|
|Conflict|Запрос не удалось обработать из-за конфликта.|
|ItemAlreadyExists|Создаваемый ресурс уже существует.|
|UnsupportedOperation|Выполняемая операция не поддерживается.|
|RequestAborted|Запрос прерван во время выполнения.|
|ApiNotAvailable|Запрашиваемый интерфейс API недоступен.|
|InsertDeleteConflict|Операция вставки или удаления привела к конфликту.|
|InvalidOperation|Выполняемая операция недопустима для этого объекта.|

### Дополнительные ресурсы

* [Создание первой надстройки Excel](build-your-first-excel-add-in.md)
* [Обозреватель фрагментов кода](https://github.com/OfficeDev/office-js-snippet-explorer)
* [Примеры кода надстройки Excel](excel-add-ins-code-samples.md)
* [Справочник по API JavaScript для надстроек Excel](excel-add-ins-javascript-reference.md)

