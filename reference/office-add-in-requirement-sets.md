
# Наборы требований для надстроек Office

Наборы требований — это именованные группы элементов API. Надстройки Office используют наборы требований, указанные в манифесте, проверки в среде выполнения, чтобы определить, поддерживает ли ведущее приложение Office необходимые надстройке API. Дополнительные сведения см. в статье [Указание ведущих приложений Office и требований к API](../docs/overview/specify-office-hosts-and-api-requirements.md).

Общие сведения о том, какие надстройки поддерживаются основными приложениями Office, см. на странице [Доступность надстроек Office в основных приложениях и на платформах](https://dev.office.com/add-in-availability).

## Наборы требований


В таблице ниже указаны имена наборов требований, методы в каждом наборе, ведущие приложения Office, поддерживающие этот набор требований, а также номера версий API.

Сведения о наборах требований для Outlook см. в статье [Общие сведения о наборах требований API Outlook](./outlook/tutorial-api-requirement-sets.md).

|  Имя набора  |  Версия  |  Основное приложение Office  |  Методы в наборе  |
|:-----|-----|:-----|:-----|
| ExcelApi   | 1.2 | Excel 2016<br>Excel Online<br>Excel для iPad<br>|Защита листа<br>Функции листа<br>Сортировка<br>Фильтр<br>Стиль ссылок R1C1<br>Объединение ячеек<br>Настройка высоты строк и ширины столбцов<br>Chart.getImage()<br>Range.getUsedRange(valuesOnly)|
| ExcelApi   | 1.1 | Excel 2016<br>Excel Online<br>Excel для iPad<br>|Все элементы в пространстве имен Excel|
| WordApi    | 1.2 | Word 2016<br>Word 2016 для Mac<br>Word для iPad<br>Word Online (ознакомительная версия) | Все элементы в пространстве имен Word. В этой версии WordApi добавлены следующие методы:<br>Body.select(selectionMode)<br>Body.insertInlinePictureFromBase64(base64EncodedImage, insertLocation)<br>contentControl.select(selectionMode)<br>contentControl.insertInlinePictureFromBase64(base64EncodedImage, insertLocation)<br>inlinePicture.paragraph<br>inlinePicture.delete<br>inlinePicture.insertBreak(breakType, insertLocation)<br>inlinePicture.insertFileFromBase64(base64file, insertLocation)<br>inlinePicture.insertHtml(html, insertLocation)<br>inlinePicture.insertInlinePictureFromBase64(base64file, insertLocation)<br>inlinePicture.insertOoxml(ooxml, insertLocation)<br>inlinePicture.insertParagraph(paragraphText, insertLocation)<br>inlinePicture.insertText(text, insertLocation)<br>inlinePicture.select(selectionMode)<br>paragraph.select(selectionMode)<br>range.inlinePictures<br>range.select(selectionMode)<br>range.insertInlinePictureFomBase64(base64EcodedImage, insertLocation)|
| WordApi    | 1.1 | Word 2016<br>Word 2016 для Mac<br>Word для iPad<br>|Все элементы в пространстве имен Word, кроме элементов API, добавленных в WordApi 1.2 и более поздних версиях, перечисленных ниже.|
| ActiveView | 1.1 | PowerPoint<br>PowerPoint Online|Document.getActiveViewAsync|
| BindingEvents  | 1.1 | Веб-приложения Access<br>Excel<br>Excel Online<br>Word 2013 и более поздних версий<br>Word 2016 для Mac<br>Word Online<br>Word для iPad|Binding.addHanderAsync<br>Binding.removeHanderAsync|
| CompressedFile    | 1.1 |PowerPoint<br>Word 2013 и более поздних версий<br>Word 2016 для Mac<br>Word Online<br>Word для iPad<br/>Excel Online<br/>PowerPoint Online|Поддерживает вывод в формате Office Open XML (OOXML) в виде байтового массива<br>(Office.FileType.Compressed) при использовании метода Document.getFileAsync.|
| CustomXmlParts    | 1.1 |Word 2013 и более поздних версий<br>Word 2016 для Mac<br>Word Online<br>Word для iPad|CustomXmlNode.getNodesAsync<br>CustomXmlNode.getNodeValueAsync<br>CustomXmlNode.getXmlAsync<br>CustomXmlNode.setNodeValueAsync<br>CustomXmlNode.setXmlAsync<br>CustomXmlPart.addHandlerAsync<br>CustomXmlPart.deleteAsync<br>CustomXmlPart.getNodesAsync<br>CustomXmlPart.getXmlAsync<br>CustomXmlPart.removeHandlerAsync<br>CustomXmlParts.addAsync<br>CustomXmlParts.getByIdAsync<br>CustomXmlParts.getByNamespaceAsync<br>CustomXmlPrefixMappings.addNamespaceAsync<br>CustomXmlPrefixMappings.getNamespaceAsync<br>CustomXmlPrefixMappings.getPrefixAsync|
| DialogAPI | 1.1 | Excel<br>PowerPoint<br>Word 2016<br>Outlook|Office.context.ui.displayDialogAsync()<br>Office.context.ui.messageParent()<br>Office.context.ui.close()|
| DocumentEvents    | 1.1 | Excel<br>Excel Online<br>PowerPoint<br>Word 2013 и более поздних версий<br>Word 2016 для Mac<br>Word Online<br>Word для iPad|Document.addHandlerAsync<br>Document.removeHandlerAsync|
| Файл  | 1.1 | PowerPoint<br>Word 2013 и более поздних версий<br>Word 2016 для Mac<br>Word Online<br>Word для iPad<br>PowerPoint Online|Document.getFileAsync<br>File.closeAsync<br>File.getSliceAsync|
| HtmlCoercion  | 1.1 | Word 2013 и более поздних версий<br>Word 2016 для Mac<br>Word Online<br>Word для iPad|Поддерживает приведение в HTML (Office.CoercionType.Html) при чтении и записи данных с использованием методов Document.getSelectedDataAsync,<br>Document.setSelectedDataAsync, Binding.getDataAsync и Binding.setDataAsync.|
| ImageCoercion | 1.1 | Word 2013 и более поздних версий<br>Word 2016 для Mac<br>Word Online<br>Word для iPad|Поддерживает преобразование в изображение (Office.CoercionType.Image) при записи данных с помощью метода Document.setSelectedDataAsync.|
| Почтовый ящик   |   | Outlook для Windows<br>Outlook для веб-браузеров<br>Outlook для Mac<br>Outlook Web App |ознакомьтесь со статьей [Общие сведения о наборах требований API Outlook](./outlook/tutorial-api-requirement-sets.md)|
| MatrixBindings    | 1.1 | Excel<br>Excel Online<br>Word|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncMatrix<br>Binding.getDataAsyncMatrix<br>Binding.setDataAsync|
| MatrixCoercion    | 1.1 | Excel<br>Excel Online<br>Word 2013 и более поздних версий<br>Word 2016 для Mac<br>Word Online<br>Word для iPad|Поддерживает приведение в структуру данных "матрица" (массив массивов, Office.CoercionType.Matrix) при чтении и записи данных с использованием методов Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync и Binding.setDataAsync.|
| OoxmlCoercion | 1.1 | Word 2013 и более поздних версий<br>Word 2016 для Mac<br>Word Online<br>Word для iPad|Поддерживает приведение в формат Open Office XML (OOXML, Office.CoercionType.Ooxml) при чтении и записи данных с использованием методов Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync и Binding.setDataAsync.|
| PartialTableBindings  | 1.1 | Веб-приложения Access||
| PdfFile   | 1.1 | PowerPoint<br/>PowerPoint Online<br/>Word 2013 и более поздних версий<br>Word 2016 для Mac<br>Word Online<br>Word для iPad|Поддерживает вывод в формате PDF (Office.FileType.Pdf)<br>при использовании метода Document.getFileAsync.|
| Выделение | 1.1 | Excel<br>Excel Online<br>PowerPoint<br>Project<br>Word 2013 и более поздних версий<br>Word 2016 для Mac<br>Word Online<br>Word для iPad|Document.getSelectedDataAsync<br>Document.setSelectedDataAsync|
| Параметры  | 1.1 | Веб-приложения Access<br>Excel<br>Excel Online<br>PowerPoint<br>PowerPoint Online<br>Word 2013 и более поздних версий<br>Word 2016 для Mac<br>Word Online<br>Word для iPad|Settings.get<br>Settings.remove<br>Settings.saveAsync<br>Settings.set|
| TableBindings | 1.1 | Веб-приложения Access<br>Excel<br>Excel Online<br>Word 2013 и более поздних версий<br>Word 2016 для Mac<br>Word Online<br>Word для iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncTable<br>Binding.addColumnsAsyncTable<br>Binding.addRowsAsyncTable<br>Binding.deleteAllDataValuesAsyncTable<br>Binding.getDataAsyncTable<br>Binding.setDataAsync|
| TableCoercion | 1.1 | Веб-приложения Access<br>Excel<br>Excel Online<br>Word 2013 и более поздних версий<br>Word 2016 для Mac<br>Word Online<br>Word для iPad|Поддерживает приведение в структуру данных "таблица" (Office.CoercionType.Table) при чтении и записи данных с использованием методов Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync и Binding.setDataAsync.|
| TextBindings  | 1.1 | Excel<br>Excel Online<br>Word 2013 и более поздних версий<br>Word 2016 для Mac<br>Word Online<br>Word для iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncText<br>Binding.getDataAsyncText<br>Binding.setDataAsync|
| TextCoercion  | 1.1 | Excel<br>Excel Online<br>PowerPoint<br>Project<br>Word 2013 и более поздних версий<br>Word 2016 для Mac<br>Word Online<br>Word для iPad|Поддерживает приведение в текстовый формат (Office.CoercionType.Text) при чтении и записи данных с использованием методов Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync и Binding.setDataAsync.|
| TextFile  | 1.1 | Word 2013 и более поздних версий<br>Word 2016 для Mac<br>Word Online<br>Word для iPad<br/>|Поддерживает вывод в текстовом формате (Office.FileType.Text) при использовании метода Document.getFileAsync.|

## Методы, отсутствующие в наборе требований


Указанные ниже методы API JavaScript API для Office не входят в состав набора требований. Если вашей надстройке необходимы какие-либо из этих методов, используйте элементы **Methods** и **Method** в манифесте надстройки, чтобы объявить их обязательными, или выполняйте проверку в среде выполнения с использованием оператора if. Дополнительные сведения см. в статье [Указание ведущих приложений Office и требований к API](../docs/overview/specify-office-hosts-and-api-requirements.md).



|**Имя метода**|**Поддержка ведущих приложений Office**|
|:-----|:-----|
|Bindings.addFromPromptAsync|Веб-приложения Access, Excel и Excel Online|
|Document.getFilePropertiesAsync|Excel, Excel Online, Word и PowerPoint|
|Document.getProjectFieldAsync|Project стандартный 2013 и Project профессиональный 2013|
|Document.getResourceFieldAsync|Project стандартный 2013 и Project профессиональный 2013|
|Document.getSelectedResourceAsync|Project стандартный 2013 и Project профессиональный 2013|
|Document.getSelectedTaskAsync|Project стандартный 2013 и Project профессиональный 2013|
|Document.getSelectedViewAsync|PowerPoint и PowerPoint Online|
|Document.getTaskAsync|Project стандартный 2013 и Project профессиональный 2013|
|Document.getTaskFieldAsync|Project стандартный 2013 и Project профессиональный 2013|
|Document.goToByIdAsync|Excel, Excel Online, Word и PowerPoint|
|Settings.addHandlerAsync|Веб-приложения Access, Excel, Excel Online, Word и PowerPoint|
|Settings.refreshAsync|Веб-приложения Access, Excel, Excel Online, Word, PowerPoint и PowerPoint Online|
|Settings.removeHandlerAsync|Веб-приложения Access, Excel, Excel Online, Word и PowerPoint|
|TableBinding.clearFormatsAsync|Excel, Excel Online|
|TableBinding.setFormatsAsync|Excel, Excel Online|
|TableBinding.setTableOptionsAsync|Excel, Excel Online|

## Дополнительные ресурсы



- [Указание основных приложений Office и требований к API](../docs/overview/specify-office-hosts-and-api-requirements.md)

