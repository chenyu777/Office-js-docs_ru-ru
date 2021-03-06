
# Создание надстройки области задач словаря


В этой статье представлен пример надстройки области задач и соответствующей веб-службы, которые предоставляют словарные статьи определений или синонимов из тезауруса к слову, выбранному пользователем в документе Word 2013. 

Надстройка словаря Office базируется на стандартной надстройке области задач с дополнительными функциональными возможностями поддержки запросов и отображения определений из словарной веб-службы XML в дополнительных расположениях пользовательского интерфейса приложения Office. 

В обычной надстройке области задач словаря пользователь выбирает слово или фразу в своем документе, после чего выделенный фрагмент передается в XML-веб-службу поставщика словаря с использованием логики JavaScript, которая лежит в основе надстройки. Затем веб-страница поставщика словаря обновляется, чтобы показать пользователю определения выделенного фрагмента.
Компонент веб-службы XML возвращает до трех определений в формате, определенном XML-схемой OfficeDefinitions, которые затем отображаются в разных местах пользовательского интерфейса ведущего приложения Office. На рисунке 1 показано, как надстройка словаря Bing, запущенная в Word 2013, отображает выделенное слово и его определения.

**Рис. 1. Надстройка словаря, отображающая определения выделенного слова**


![Приложение словаря, в котором отображается определение](../../images/DictionaryAgave01.jpg)

Вы выбираете, что отображается при переходе по ссылке **Подробнее** в пользовательском интерфейсе HTML надстройки словаря: дополнительные сведения в области задач либо полная веб-страница для выделенного слова или фразы в отдельном окне браузера. На рис. 2 приведена команда контекстного меню **Определение**, которая позволяет быстро запустить установленные словари. На рис. 3–5 перечислены все расположения в пользовательском интерфейсе Office, в которых словарные XML-службы предоставляют определения в Word 2013.

**Рис. 2. Команда определения в контекстном меню**



![Контекстное меню определения](../../images/DictionaryAgave02.jpg)

**Рис. 3. Определения в областях "Правописание" и "Грамматика"**


![Определения в областях "Правописание" и "Грамматика"](../../images/DictionaryAgave03.jpg)

**Рис. 4. Определения в области "Тезаурус"**


![Определения в области "Тезаурус"](../../images/DictionaryAgave04.jpg)

**Рис. 5. Определения в режиме чтения**


![Определения в режиме чтения](../../images/DictionaryAgave05.jpg)

Чтобы создать надстройку области задач, которая выполняет поиск в словаре, необходимо создать два основных компонента: 


- веб-службу XML, которая ищет определения в словарной службе, а затем возвращает результаты в формате XML, которые могут быть отображены в надстройке словаря;
    
- надстройку области задач, которая отправляет выбранное пользователем слово или фразу в словарную веб-службу, отображает определения и может вставить эти значения в документ.
    
В следующих разделах приведены примеры создания этих компонентов.

## Создание словарной веб-службы XML


Веб-служба XML должна возвращать запросы веб-служб в виде XML-кода, который соответствует XML-схеме OfficeDefinitions. В двух следующих разделах описывается XML-схема OfficeDefinitions и предоставлен пример возможности кодирования веб-службы XML, возвращающей запросы в этом формате XML.


### XML-схема OfficeDefinitions

В следующем коде отображается XSD для XML-схемы OfficeDefinitions.


```XML
<?xml version="1.0" encoding="utf-8"?>
<xs:schema
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  targetNamespace="http://schemas.microsoft.com/NLG/2011/OfficeDefinitions"
  xmlns="http://schemas.microsoft.com/NLG/2011/OfficeDefinitions">
  <xs:element name="Result">
    <xs:complexType>
      <xs:sequence>
        <xs:element name="SeeMoreURL" type="xs:anyURI"/>
        <xs:element name="Definitions" type="DefinitionListType"/>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
  <xs:complexType name="DefinitionListType">
    <xs:sequence>
      <xs:element name="Definition" maxOccurs="3">
        <xs:simpleType>
          <xs:restriction base="xs:normalizedString">
            <xs:maxLength value="400"/>
          </xs:restriction>
        </xs:simpleType>
      </xs:element>
    </xs:sequence>
  </xs:complexType>
</xs:schema>
```

Возвращенный XML-код, который соответствует схеме OfficeDefinitions, состоит из корневого элемента **Result**, содержащего элемент **Definitions** с дочерними элементами **Definition** в количестве от нуля до трех. Каждый из этих дочерних элементов содержит определения, длина которых не превышает 400 символов. Кроме того, URL-адрес полной страницы на сайте словаря должен быть предоставлен в элементе **SeeMoreURL**. В примере ниже показана структура возвращенного XML-кода, соответствующего схеме OfficeDefinitions.




```XML
<?xml version="1.0" encoding="utf-8"?>
<Result xmlns="http://schemas.microsoft.com/NLG/2011/OfficeDefinitions">
  <SeeMoreURL xmlns="">www.bing.com/dictionary/search?q=example</SeeMoreURL>
  <Definitions xmlns="">
    <Definition>Definition1</Definition>
    <Definition>Definition2</Definition>
    <Definition>Definition3</Definition>
  </Definitions>
 </Result>

```


### Пример словарной веб-службы XML

Приведенный ниже код C# предоставляет простой пример написания кода для веб-службы XML, которая возвращает результат запроса словаря в XML-формате OfficeDefinitions.


```C#
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Services;
using System.Xml;
using System.Text;
using System.IO;
using System.Net;

/// <summary>
/// Summary description for _Default
/// </summary>
[WebService(Namespace = "http://tempuri.org/")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
// To allow this web service to be called from script, using ASP.NET AJAX, uncomment the following line. 
// [System.Web.Script.Services.ScriptService]
public class WebService : System.Web.Services.WebService {

    public WebService () {

        // Uncomment the following line if using designed components 
        // InitializeComponent(); 
    }

    // You can replace this method entirely with your own method that gets definitions
    // from your data source, and then formats it into the OfficeDefinitions XML format. 
    // If you need a reference for constructing the returned XML, you can use this example as a basis.
    [WebMethod]
    public XmlDocument Define(string word)
    {

        StringBuilder sb = new StringBuilder();
        XmlWriter writer = XmlWriter.Create(sb);
        {
            writer.WriteStartDocument();
            
                writer.WriteStartElement("Result", "http://schemas.microsoft.com/NLG/2011/OfficeDefinitions");

            // See More URL should be changed to the dictionary publisher's page for that word on their website.
                    writer.WriteElementString("SeeMoreURL", "http://www.bing.com/search?q=" + word);

                    writer.WriteStartElement("Definitions");
            
                        writer.WriteElementString("Definition", "Definition 1 of " + word);
                        writer.WriteElementString("Definition", "Definition 2 of " + word);
                        writer.WriteElementString("Definition", "Definition 3 of " + word);
                   
                    writer.WriteEndElement();


                writer.WriteEndElement();
            
            writer.WriteEndDocument();
        }
        writer.Close();

        XmlDocument doc = new XmlDocument();
        doc.LoadXml(sb.ToString());

        return doc;
    }
   

}
```


## Создание компонентов надстройки словаря


Надстройка словаря состоит из трех файлов основных компонентов.


- XML-файл манифеста, который описывает надстройку.
    
- HTML-файл, который предоставляет пользовательский интерфейс надстройки.
    
- Файл JavaScript, который содержит логику для получения выделенного пользователем фрагмента из документа, отправки выбранного слова или фразы в веб-службу и отображения возвращенных результатов в пользовательском интерфейсе надстройки.
    

### Создание файла манифеста надстройки словаря

Ниже приведен пример файла манифеста для надстройки словаря.


```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
  <Id>DemoDict</Id>
  <Version>15.0</Version>
  <ProviderName>Microsoft Office Demo Dictionary</ProviderName>
  <DefaultLocale>en-us</DefaultLocale>
  <!--DisplayName is the name that will appear in the user's list of applications.-->
  <DisplayName DefaultValue="Microsoft Office Demo Dictionary" />
  <!--Description is a 2-3 sentence description of this dictionary. -->
  <Description DefaultValue="The Microsoft Office Demo Dictionary is an example built to demonstrate how a publisher could create a dictionary that integrates with Office. It does not return real definitions." />
  <!--IconUrl is the URI for the icon that will appear in the user's list of applications.-->
  <IconUrl DefaultValue="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg" />
  <!--Capabilities specifies the kind of host application your dictionary add-in will support. You shouldn't have to modify this area.-->
  <Capabilities>
    <Capability Name="Workbook"/>
    <Capability Name="Document"/>
    <Capability Name="Project"/>
  </Capabilities>
  <DefaultSettings>
    <!--SourceLocation is the URL for your dictionary-->
    <SourceLocation DefaultValue="http://christophernlg/ExampleDictionary/DictionaryHome.html" />
  </DefaultSettings>
  <!--Permissions is the set of permissions a user will have to give your dictionary. If you need write access, such as to allow a user to replace the highlighted word with a synonym, use ReadWriteDocument. -->
  <Permissions>ReadDocument</Permissions>
  <Dictionary>
    <!--TargetDialects is the set of dialects your dictionary contains. For example, if your dictionary applies to Spanish (Mexico) and Spanish (Peru), but not Spanish (Spain), you can specify that here. This is for different dialects of the same language. Please do NOT put more than one language (for example, Spanish and English) here. Publish separate languages as separate dictionaries. -->
    <TargetDialects>
      <TargetDialect>EN-AU</TargetDialect>
      <TargetDialect>EN-BZ</TargetDialect>
      <TargetDialect>EN-CA</TargetDialect>
      <TargetDialect>EN-029</TargetDialect>
      <TargetDialect>EN-HK</TargetDialect>
      <TargetDialect>EN-IN</TargetDialect>
      <TargetDialect>EN-ID</TargetDialect>
      <TargetDialect>EN-IE</TargetDialect>
      <TargetDialect>EN-JM</TargetDialect>
      <TargetDialect>EN-MY</TargetDialect>
      <TargetDialect>EN-NZ</TargetDialect>
      <TargetDialect>EN-PH</TargetDialect>
      <TargetDialect>EN-SG</TargetDialect>
      <TargetDialect>EN-ZA</TargetDialect>
      <TargetDialect>EN-TT</TargetDialect>
      <TargetDialect>EN-GB</TargetDialect>
      <TargetDialect>EN-US</TargetDialect>
      <TargetDialect>EN-ZW</TargetDialect>
    </TargetDialects>
    <!--QueryUri is the address of this dictionary's XML web service (which is used to put definitions in additional contexts, such as the spelling checker.)-->
    <QueryUri DefaultValue="http://christophernlg/ExampleDictionary/WebService.asmx/Define?word="/>
    <!--Citation Text, Dictionary Name, and Dictionary Home Page will be combined to form the citation line (for example, this would produce "Examples by: Microsoft", where "Microsoft" is a hyperlink to http://www.microsoft.com).-->
    <CitationText DefaultValue="Examples by: " />
    <DictionaryName DefaultValue="Microsoft" />
    <DictionaryHomePage DefaultValue="http://www.microsoft.com" />
  </Dictionary>
</OfficeApp>
```

Элемент **Dictionary** и его дочерние элементы, относящиеся непосредственно к созданию файла манифеста надстройки словаря, приведены в разделах ниже. Сведения о других элементах в файле манифеста см. в статье [XML-манифест надстроек для Office](../../docs/overview/add-in-manifests.md).


### Элемент Dictionary


Определяет параметры надстроек словаря.

 **Родительский элемент**

 `<OfficeApp>`

 **Дочерние элементы**

 `<TargetDialects>`,  `<QueryUri>`,  `<CitationText>`,  `<DictionaryName>`,  `<DictionaryHomePage>`

 **Замечания**

Элемент **Dictionary** и его дочерние элементы добавляются в манифест надстройки области задач при создании надстройки словаря.


#### Элемент TargetDialects


Определяет диалекты, поддерживаемые этим словарем. Обязательный (для надстроек словаря).

 **Родительский элемент**

 `<Dictionary>`

 **Дочерний элемент**

 `<TargetDialect>`

 **Замечания**

Элемент **TargetDialects** и его дочерние элементы определяют набор языков, содержащихся в словаре. Например, если словарь применим к диалектам испанского языка в Мексике и Перу, но не к диалекту в Испании, это можно обозначить в указанном элементе. Данный элемент предназначен исключительно для разных диалектов одного языка. Не указывайте в этом манифесте несколько языков (например, английский и испанский). Публикуйте разные языки как разные словари.

 **Пример**




```XML
<TargetDialects>
  <TargetDialect>EN-AU</TargetDialect>
  <TargetDialect>EN-BZ</TargetDialect>
  <TargetDialect>EN-CA</TargetDialect>
  <TargetDialect>EN-029</TargetDialect>
  <TargetDialect>EN-HK</TargetDialect>
  <TargetDialect>EN-IN</TargetDialect>
  <TargetDialect>EN-ID</TargetDialect>
  <TargetDialect>EN-IE</TargetDialect>
  <TargetDialect>EN-JM</TargetDialect>
  <TargetDialect>EN-MY</TargetDialect>
  <TargetDialect>EN-NZ</TargetDialect>
  <TargetDialect>EN-PH</TargetDialect>
  <TargetDialect>EN-SG</TargetDialect>
  <TargetDialect>EN-ZA</TargetDialect>
  <TargetDialect>EN-TT</TargetDialect>
  <TargetDialect>EN-GB</TargetDialect>
  <TargetDialect>EN-US</TargetDialect>
  <TargetDialect>EN-ZW</TargetDialect>
</TargetDialects>
```


#### Элемент TargetDialect


Определяет диалект, поддерживаемый этим словарем. Обязательный (для надстроек словаря).

 **Родительский элемент**

 `<TargetDialects>`

 **Замечания**

Укажите значение диалекта в формате тегов `language` RFC1766, например EN-US.

 **Пример**




```XML
<TargetDialect>EN-US</TargetDialect>
```


#### Элемент QueryUri


Указывает конечную точку для службы запросов словаря. Обязательный (для надстроек словаря).

 **Родительский элемент**

 `<Dictionary>`

 **Замечания**

Это URI XML-веб-службы поставщика словаря. К этому URI добавляется строка запроса с надлежащими escape-символами. 

 **Пример**




```XML
<QueryUri DefaultValue="http://msranlc-lingo1/proof.aspx?q="/>
```


#### Элемент CitationText


Задает текст, используемый в ссылках. Обязательный (для надстроек словаря).

 **Родительский элемент**

 `<Dictionary>`

 **Замечания**

В этом элементе указывается начальный текст ссылки, который будет отображаться в строке под контентом, возвращенным из веб-службы (например, "Источник:" или "Предоставлено:").

Для этого элемента можно указать значения в других языковых стандартах, используя для этого элемент **Override**. Например, если пользователь использует версию Office на испанском языке, но задействует английский словарь, то в строке ссылки будет написано "Resultados por: Bing", а не "Results by: Bing". Чтобы узнать, как указывать значения с использованием других языковых стандартов, см. раздел "Параметры для разных языковых стандартов" в статье [XML-манифест надстроек для Office](../../docs/overview/add-in-manifests.md).

 **Пример**




```XML
<CitationText DefaultValue="Results by: " />
```


#### Элемент DictionaryName


Задает имя словаря. Обязательный (для надстроек словаря).

 **Родительский элемент**

 `<Dictionary>`

 **Замечания**

В этом элементе указывается текст ссылки на источник. Текст ссылки на источник отображается в строчке под контентом, возвращенным веб-службой.

В этом элементе можно задать значения для дополнительных языковых стандартов.

 **Пример**




```XML
<DictionaryName DefaultValue="Bing Dictionary" />
```


#### Элемент DictionaryHomePage


Указывает URL-адрес домашней страницы для словаря. Обязательный (для надстроек словаря).

 **Родительский элемент**

 `<Dictionary>`

 **Замечания**

В этом элементе указывается URL-адрес источника. Текст ссылки на источник отображается в строчке под контентом, возвращенным веб-службой.

В этом элементе можно задать значения для дополнительных языковых стандартов.

 **Пример**




```XML
<DictionaryHomePage DefaultValue="http://www.bing.com" />
```


### Создание пользовательского интерфейса HTML для надстройки словаря


В двух следующих примерах показаны HTML- и CSS-файлы для пользовательского интерфейса демонстрационной надстройки словаря. Чтобы просмотреть, как отображается пользовательский интерфейс в надстройке области задач, изучите рис. 6, который приведен сразу после кода. Чтобы узнать, как реализация JavaScript в файле Dictionary.js предоставляет логику программирования для этого пользовательского интерфейса HTML, см. раздел "Составление реализации JavaScript" ниже.


```HTML
<!DOCTYPE html>
<html>

<head>
<meta http-equiv="X-UA-Compatible" content="IE=Edge"/>

<!--The title will not be shown but is supplied to ensure valid HTML.-->
<title>Example Dictionary</title>

<!--Required library includes.-->
<script type="text/javascript" src="http://ajax.microsoft.com/ajax/4.0/1/MicrosoftAjax.js"></script>
<script type="text/javascript" src="office.js"></script>

<!--Optional library includes.-->
<script type="text/javascript" src="http://ajax.aspnetcdn.com/ajax/jQuery/jquery-1.5.1.js"></script>

<!--App-specific CSS and JS.-->
<link rel="Stylesheet" type="text/css" href="style.css" />
<script type="text/ecmascript" src="dictionary.js"></script>
</head>

<body>
<div id="mainContainer">
    <div id="header">
        <span id="headword"></span>
        <span id="pronunciation">(<a id="pronunciationLink">Pronounce</a>)</span>
    </div>
    <ol id="definitions">
    </ol>
    <div id="SeeMore">
    <a id="SeeMoreLink">See More...</a>
    </div>
</div>
</body>

</html>
```

В приведенном ниже примере показано содержание Style.css.




```
#mainContainer
{
    font-family: Segoe UI;
    font-size: 11pt;
}

#headword
{
    font-family: Segoe UI Semibold;
    color: #262626;
}

#pronunciation
{
    margin-left: 2px;
    margin-right: 2px;
}

#definitions
{
    font-size: 8.5pt;
}
a
{
    font-size: 8pt;
    color: #336699;
    text-decoration: none;
}
a:visited
{
    color: #993366;
}
a:hover, a:active
{  
    text-decoration: underline;
}
```


**Рис. 6. Пользовательский интерфейс демоверсии словаря**

![Пользовательский интерфейс демоверсии словаря](../../images/DictionaryAgave06.jpg)


### Реализация JavaScript


В приведенном ниже примере показана реализация JavaScript в файле Dictionary.js, которая вызывается с HTML-страницы надстройки и предоставляет программную логику для надстройки Demo Dictionary. В этом сценарии используется вышеописанная XML-веб-служба. Если поместить сценарий в тот же каталог, что и пример веб-службы, он будет получать определения из этой службы. Его можно использовать с общедоступной XML-веб-службой, соответствующей схеме OfficeDefinitions. Для этого измените переменную `xmlServiceURL` в начале файла, а затем замените ключ API Bing для произношений на правильно зарегистрированный.

Ниже приведены основные элементы API JavaScript для Office (Office.js), которые вызываются в реализованном коде.


- Событие [initialize](../../reference/shared/office.initialize.md) объекта **Office**, возникающее при инициализации контекста надстройки и предоставляющее доступ к объекту [Document](../../reference/shared/document.md), представляющему собой документ, с которым взаимодействует надстройка.
    
- Метод [addHandlerAsync](../../reference/shared/document.addhandlerasync.md) объекта **Document**, который вызывается в функции **initialize** для добавления обработчика события [SelectionChanged](../../reference/shared/document.selectionchanged.event.md) документа с целью прослушивания изменений, внесенных в выделенный пользователем фрагмент.
    
- Метод [getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md) объекта **Document**, который вызывается в функции `tryUpdatingSelectedWord()` при включении обработчика событий **SelectionChanged** для получения слова или фразы по выбору пользователя, приведения их в обычный текст и выполнения асинхронной функции обратного вызова `selectedTextCallback`.
    
- При выполнении асинхронной функции обратного вызова `selectTextCallback`, которая передается как аргумент _callback_ метода **getSelectedDataAsync**, возвращается значение выделенного текста. Эта функция считывает значение из аргумента _selectedText_ (имеющего тип [AsyncResult](../../reference/shared/asyncresult.md)) с помощью свойства [value](../../reference/shared/asyncresult.status.md) возвращенного объекта **AsyncResult**.
    
- Остальной код функции `selectedTextCallback` отправляет XML-веб-службе запрос на определения. Кроме того, он вызывает API-интерфейсы Microsoft Translator для получения URL-адреса WAV-файла с произношением выделенного слова.
    
- Остальной код в файле Dictionary.js служит для отображения списка определений и ссылок на произношения в пользовательском интерфейсе HTML надстройки.
    



```
// The document the dictionary add-in is interacting with.
var _doc; 
// The last looked-up word, which is also the currently displayed word.
var lastLookup; 
// For demo purposes only!! Get an AppID if you intend to use the Pronunciation service for your feature.
var appID="3D8D4E1888B88B975484F0CA25CDD24AAC457ED8"; 

// The base URL for the OfficeDefinitions-conforming XML web service to query for definitions.
var xmlServiceUrl = "WebService.asmx/Define?Word="; 

// Initialize the add-in. 
// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    // Store a reference to the current document.
    _doc = Office.context.document; 
    // Check whether text is already selected.
    tryUpdatingSelectedWord(); 
    _doc.addHandlerAsync("documentSelectionChanged", tryUpdatingSelectedWord); //Add a handler to refresh when the user changes selection.
    });
}

// Executes when event is raised on user's selection changes, and at initialization time. 
// Gets the current selection and passes that to asynchronous callback method.
function tryUpdatingSelectedWord() {
    _doc.getSelectedDataAsync(Office.CoercionType.Text, selectedTextCallback); 
}

// Async callback that executes when the add-in gets the user's selection.
// Determines whether anything should be done. If so, it makes requests that will be passed to various functions.
function selectedTextCallback(selectedText) {
    selectedText = $.trim(selectedText.value);
    // Be sure user has selected text. The SelectionChanged event is raised every time the user moves the cursor, even if no selection.
    if (selectedText != "") { 
        // Check whether user selected the same word the pane is currently displaying to avoid unnecessary web calls.
        if (selectedText != lastLookup) { 
            // Update the lastLookup variable.
            lastLookup = selectedText; 
            // Set the "headword" span to the word you looked up.
            $("#headword").text(selectedText); 
            // AJAX request to get definitions for the selected word; pass that to refreshDefinitions.
            $.ajax(xmlServiceUrl + selectedText, { dataType: 'xml', success: refreshDefinitions, error: errorHandler }); 
            // AJAX request to the Microsoft Translator APIs. Gets the URL of a WAV file with pronunciation, which is passed to refreshPronunciation. See http://www.microsofttranslator.com/dev for details.
            $.ajax("http://api.microsofttranslator.com/V2/Ajax.svc/Speak?oncomplete=refreshPronunciation&amp;appId=" + appID + "&amp;text=" + selectedText + "&amp;language=en-us", { dataType: 'script', success: null, error: errorHandler }); 
        }
    }
}

// This function is called when the add-in gets back the definitions target word.
// It removes the old definitions and replaces them with the definitions for the current word.
// It also sets the "See More" link.
function refreshDefinitions(data, textStatus, jqXHR) {
    $(".definition").remove();
    // Make a new list item for each returned definition that was returned, set the CSS class, and append it to the definitions div.
    $(data).find("Definition").each(function () {
        $(document.createElement("li")).text($(this).text()).addClass("definition").appendTo($("#definitions"));
    });
    $("#SeeMoreLink").attr("href", $(data).find("SeeMoreURL").text()); //Change the "See More" link to direct to the correct URL.
}

// This function is called when the add-in gets back the link to the pronunciation
// to set the "Pronounce" link to the URL of the .WAV file.
function refreshPronunciation(data) {
    $("#pronunciationLink").attr("href", data);
}

// Basic error handler that writes to a div with id='message'.
function errorHandler(jqXHR, textStatus, errorThrown) {
    document.getElementById('message').innerText += errorThrown;
}

```

