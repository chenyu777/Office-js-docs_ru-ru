
# Свойство Context.contentLanguage
 Получает языковой стандарт (язык), указанный пользователем для редактирования документа или элемента.

|||
|:-----|:-----|
|**Ведущие приложения:**|Access, Excel, PowerPoint, Project, Word|
|**Последнее изменение в **|1.1|

```
var myContentLang = Office.context.contentLanguage;
```


## Возвращаемое значение

Значение **string** в формате языковых обозначений RFC 1766, например `en-US`.


## Замечания

Значение **contentLanguage** отражает значение параметра **Язык редактирования**, указанное путем выполнения команды **Файл**  >  **Параметры**  >  **Язык** ведущего приложения Office.

В контентных надстройках для веб-приложений Access свойство **contentLanguage** получает язык и региональные параметры надстройки (например, "en-GB").


## Пример




```js
function sayHelloWithContentLanguage() {
    var myContentLanguage = Office.context.contentLanguage;
    switch (myContentLanguage) {
        case 'en-US':
            write('Hello!');
            break;
        case 'en-NZ':
            write('G\'day mate!');
            break;
    }
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает на то, что этот метод поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает этот метод.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).

||**Office для рабочего стола Windows**|**Office Online (в браузере)**|**Office для iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Да|Y|
|**PowerPoint**|Y|Да|Y|
|**Project**|Y|||
|**Word**|Y|Да|Y|

|||
|:-----|:-----|
|**Минимальный уровень разрешений**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Контентные надстройки и надстройки области задач|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки



****


|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлена поддержка PowerPoint Online.|
|1.1|Добавлена поддержка Excel, PowerPoint и Word в Office для iPad.|
|1.1|Добавлен доступ к API в контентных надстройках для Access.|
|1.0|Представлено|
