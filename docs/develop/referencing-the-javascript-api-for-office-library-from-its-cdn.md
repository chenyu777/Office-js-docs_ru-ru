
# Указание ссылок на библиотеку JavaScript API для Office из сети доставки содержимого (CDN)


Библиотека [API JavaScript для Office](../../reference/javascript-api-for-office.md) состоит из файла Office.js и связанных JS-файлов ведущего приложения, например Excel-15.js и Outlook-15.js. 


Простейший способ добавить ссылку на API — использовать нашу сеть доставки содержимого (CDN), добавив следующий код `<script>` в тег `<head>` страницы:  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

`/1/` перед `office.js` в URL-адресе CDN указывает, что необходимо использовать последний добавочный выпуск файла Office.js версии 1. Так как API JavaScript для Office обеспечивает обратную совместимость, в последнем выпуске будут и дальше поддерживаться элементы API, представленные ранее в версии 1. Если вам нужно обновить существующий проект, см. статью [Обновление версии API JavaScript для Office и файлов схемы манифеста] (../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md). 

Если вы планируете опубликовать свою надстройку Office из Магазина Office, необходимо использовать эту ссылку на сеть CDN. Локальные ссылки подходят только для внутренних сценариев, а также сценариев разработки и отладки.

> **Важно!** Разрабатывая надстройку для любого ведущего приложения Office, ссылайтесь на API JavaScript для Office из раздела `<head>` страницы. Это гарантирует, что API полностью инициализируется раньше всех элементов body. Ведущим приложениям Office необходимо, чтобы надстройки инициализировались в течение 5 секунд после активации. При превышении этого порога объявляется, что надстройка не отвечает, и отображается сообщение об ошибке.       

## Дополнительные ресурсы



- [Общие сведения об интерфейсе JavaScript API для Office](../../docs/develop/understanding-the-javascript-api-for-office.md)
    
- [Обзор платформы надстроек Office](../../docs/overview/office-add-ins.md)
    
- [Жизненный цикл разработки надстроек Office](../../docs/design/add-in-development-lifecycle.md)
    
- [API JavaScript для Office](../../reference/javascript-api-for-office.md)
    
