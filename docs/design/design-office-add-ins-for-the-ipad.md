
# Разработка надстроек Office для iPad


В таблице ниже перечислены действия по созданию надстройки Office, которая будет работать в Office для iPad.


|**Действие**|**Описание**|**Ресурсы**|
|:-----|:-----|:-----|
|Обновление надстройки для поддержки Office.js версии 1.1.|Обновите до версии 1.1. файлы JavaScript (Office.js и JS-файлы приложения) и файл проверки манифеста надстройки, которые используете в проекте надстройки Office.|[Что изменилось в API JavaScript для Office](../../reference/what's-changed-in-the-javascript-api-for-office.md)|
|Следуйте рекомендациям по оформлению пользовательского интерфейса.|Органично интегрируйте в iOS пользовательский интерфейс надстройки.|[Разработка для iOS](https://developer.apple.com/library/ios/documentation/UserExperience/Conceptual/MobileHIG/)|
|Следуйте рекомендациям по оформлению надстройки.|Убедитесь, что ваша надстройка интересная, полезная и стабильная.|[Рекомендации по разработке надстроек Office](../../docs/overview/add-in-development-best-practices.md)|
|Оптимизируйте надстройку под сенсорный ввод.|Сделайте так, чтобы пользовательский интерфейс поддерживал не только клавиатуру и мышь, но и сенсорный ввод.|[Принципы разработки пользовательского интерфейса](https://msdn.microsoft.com/ru-ru/library/mt590883.aspx#Anchor_3)|
|Сделайте надстройку бесплатной.|Office на iPad — это канал, через который вы можете привлекать пользователей и рекламировать свои службы. Эти пользователи могут стать вашими клиентами.|[Политика проверки 10.8](http://msdn.microsoft.com/ru-ru/library/cd90836a-523e-42f5-ab02-5123cdf9fefe%28Office.15%29.aspx)|
|Сделайте надстройку некоммерческой.|У надстройки не должно быть пробных версий, она не должна содержать платных возможностей, рекламы платных версий или ссылок на интернет-магазины, в которых пользователи могут приобрести другой контент, приложения или надстройки. На страницах с политикой конфиденциальности и условиями использования также не должно быть рекламы и ссылок на магазины.|[Политика проверки 3.4](http://msdn.microsoft.com/ru-ru/library/cd90836a-523e-42f5-ab02-5123cdf9fefe%28Office.15%29.aspx)|
|Отправьте надстройку в Магазин Office повторно.|На Панели мониторинга продаж установите флажок **Включить эту надстройку в каталог надстроек Office для iPad** и укажите свой идентификатор разработчика Apple в поле "Идентификатор Apple ID". Просмотрите [соглашение с поставщиком приложений Магазина Office](https://sellerdashboard.microsoft.com/Assets/Content/Agreements/en-US/Office_Store_Seller_Agreement_20120927.md).|[Отправка надстроек Office и SharePoint, а также веб-приложений для Office 365 в Магазин Office](http://msdn.microsoft.com/ru-ru/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx)|
Для других платформ надстройку Office можно оставить без изменений. Кроме того, у надстройки может быть различный интерфейс в зависимости от браузера или устройства. Чтобы определить, запущена ли надстройка на iPad, можно использовать следующие API: 

- var isTouchEnabled = [Office.context.touchEnabled](../../reference/shared/office.context.touchenabled.md)
    
- var allowCommerce = [Office.context.commerceAllowed](../../reference/shared/office.context.commerceallowed.md)
    

## Рекомендации по разработке надстроек Office для iOS и Mac

Следуйте приведенным ниже рекомендациям по разработке надстроек для iOS.


-  **Разрабатывайте надстройку с помощью Visual Studio.**
    
    If you develop your add-in with Visual Studio, you can [set breakpoints and debug its code](../get-started/create-and-debug-office-add-ins-in-visual-studio.md#Test) in an Office host application running on Windows, before sideloading your add-in on the iPad or Mac. Because an add-in that runs in Office for iOS or Office for Mac supports the same APIs as an add-in running in Office for Windows, your add-in's code should run the same way on both platforms.
    
-  **Укажите требования касательно API в манифесте надстройки или с помощью проверок в среде выполнения.**
    
    When you specify API requirements in your add-in's manifest, Office will determine if the host application supports those API members. If the API members are available in the host, then your add-in will be available in that host application. Alternatively, you can perform a runtime check to determine if a method is available in the host before using it in your add-in. Runtime checks ensure that your add-in is always available in the host, and provides additional functionality if the methods are available. For more information, see [Specify Office hosts and API requirements](../../docs/overview/specify-office-hosts-and-api-requirements.md).
    
Общие рекомендации по разработке надстроек см. в статье [Рекомендации по разработке надстроек Office](../../docs/overview/add-in-development-best-practices.md).


## Дополнительные ресурсы
<a name="bk_addresources"></a>


- [Загрузка неопубликованной надстройки Office на iPad и Mac](../../docs/testing/sideload-an-office-add-in-on-ipad-and-mac.md)
    
- [Отладка надстроек Office на iPad и Mac](../../docs/testing/debug-office-add-ins-on-ipad-and-mac.md)
    

