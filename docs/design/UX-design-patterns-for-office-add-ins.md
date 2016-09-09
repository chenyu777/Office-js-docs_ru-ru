# Конструктивные шаблоны пользовательского интерфейса для надстроек Office. 

При разработке надстроек для Office следует учитывать, что дизайн пользовательского интерфейса надстройки должен привлекать пользователя и расширять возможности Office. Помимо прочего, качественная надстройка должна содержать интерфейс, используемый при первом запуске, первоклассный пользовательский интерфейс и обеспечивать плавные переходы между страницами. Понятный и современный интерфейс увеличивает популярность вашей надстройки среди пользователей и количество постоянно использующих ее пользователей. В этой статье представлены ресурсы пользовательского интерфейса для проектировщиков и разработчиков, которые позволяют выполнять указанные ниже задачи.

* Описывать стандартные конструктивные шаблоны пользовательских интерфейсов на основе рекомендаций.
* Реализовывать компоненты и стили Office Fabric.
* Реализовывать надстройки, выглядящие как естественные расширения пользовательского интерфейса Office, используемого по умолчанию. 

## Как начать работу с помощью ресурсов с примерами разработки надстроек Office?

Чтобы использовать эти ресурсы проектирования или кода, не требуется выполнять никаких предварительных условий. Чтобы начать создавать превосходные пользовательские интерфейсы для своих надстроек, выполните указанные ниже действия.

* Изучите конструктивные шаблоны пользовательского интерфейса и определите, какие из них лучше всего подходят для вашей надстройки. Например, выберите один из интерфейсов, используемых при первом запуске.
* Затем выполните одно или несколько указанных ниже действий.
	* Скопируйте файлы с кодом в проект надстройки и начните настраивать их в соответствии со своими требованиями. Вам потребуется [файл common.js](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/), [папка с ресурсами](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/assets) и папка с кодом для выбранного вами конструктивного шаблона. Воспользуйтесь указанными ниже ссылками.
	* Скачайте PDF-файлы справочников и используйте их при создании собственного дизайна пользовательского интерфейса. Воспользуйтесь указанными ниже ссылками.
	* Скачайте файлы Adobe Illustrator и измените их, чтобы создать макеты интерфейса для надстройки. Вы можете скачать их [здесь](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Source%20Files).
 

## Первый запуск

Интерфейс, используемый при первом запуске, — это интерфейс, отображаемый для пользователя, когда тот запускает надстройку в первый раз. Ниже перечислены конструктивные шаблоны интерфейса, используемого при первом запуске, которые вы можете включить в свою надстройку. Ниже показаны изображения каждого из конструктивных шаблонов.

* **Действия, необходимые для запуска**. Предоставляет пользователям упорядоченный список действий, которые необходимо выполнить, чтобы начать использовать надстройку. ([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/FirstRun_StepsToStart.pdf "PDF"), [код](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/first-run/instruction-step))
* **Решаемые задачи**. Разъясняет, какие задачи можно решить с помощью надстройки. ([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/FirstRun_ValuePlacemat.pdf "PDF"), [код](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/first-run/value-placemat))
* **Видео**. Показывает пользователям видеоролик перед тем, как они начнут использовать вашу надстройку. ([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/FirstRun_VideoPlacemat.pdf "PDF"), [код](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/first-run/video-placemat))
* **Пошаговое руководство**. Рассказывает пользователям о ряде функций или предоставляет определенные сведения, прежде чем они начнут использовать надстройку. ([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/FirstRun_PagingPanel.pdf "PDF"), [код](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/first-run/walkthrough))
* В [Магазине Office](https://msdn.microsoft.com/ru-ru/library/office/jj220033.aspx) имеется система, с помощью которой можно предоставить пользователям пробную версию надстройки, но если вы хотите полностью контролировать пользовательский интерфейс в пробной версии, используйте указанные ниже шаблоны.
	* **Пробная версия**. Показывает пользователям, как начать работу с пробной версией надстройки. ([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/FirstRun_TrialVersion.pdf "PDF"), [код](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/first-run/trial-placemat))
	* **Пробная функция**. Сообщает пользователю, что функция, которую он пытается использовать, недоступна в пробной версии надстройки ([код](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/first-run/trial-placemat-feature))


> Примечание. Решите, как часто необходимо применять интерфейс, используемый при первом запуске: один раз или много раз. Например, если пользователи нечасто применяют вашу надстройку, они могут забыть, как использовать ее. В таких случаях может быть полезно повторно отображать интерфейс, используемый при первом запуске. 

 <table>
 <tr><th>Действия, необходимые для запуска</th><th>Решаемые задачи</th><th>Видео</th></tr>
 <tr><td>![instruction steps" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/instruction.step.PNG)</td><td>![value placemat" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/value.placemat.PNG)</td><td>![video placemat" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/video.placemat.PNG)</td></tr>
 </table>

 <table>
 <tr><th>Первая страница пошагового руководства</th><th>Пробная версия</th><th>Пробная функция</th></tr>
 <tr><td>![walkthrough 1" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/walkthrough1.PNG)</td><td>![trial placemat" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/trial.placemat.PNG)</td><td>![trial placemat feature" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/trial.placemat.feature.PNG)</td></tr>
 </table> 


## Стандартная страница и фирменная символика

* **Целевая страница** — это первая страница, на которую пользователи попадают после страницы с интерфейсом, используемым при первом запуске, или после процесса входа. ([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Helpful%20Templates/AddIn_Template_Standard_Layout.pdf "PDF"), [код](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/generic/landing-page))

<table>
 <tr><th>Целевая</th></tr>
 <tr><td>![landing page" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/landing.page.PNG)</td></tr>
 </table>

## Уведомления

Существует много способов, которыми надстройка может уведомлять пользователей о событиях, например об ошибках или ходе выполнения действий. Эти методы перечислены ниже. Ниже показаны изображения для каждого из методов.

* **Внедренное диалоговое окно** отображается в области задач и предоставляет сведения и (при необходимости) средства взаимодействия в виде помощью кнопок и других элементов управления. Рекомендуется использовать диалоговое окно для подтверждения пользователем каких-либо действий. ([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Embedded_Dialog.pdf "PDF") , [код](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/notifications/embedded-dialog))
* **Встроенное сообщение** отображает информацию об ошибках, успешном выполнении действий или другие сведения. Может отображаться в указанном расположении в области задач. Например, если пользователь вводит в текстовом поле электронный адрес с неправильным форматом, то под полем отобразится соответствующее сообщение об ошибке. ([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Notification_Inline_Message.pdf "PDF"), [код](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/notifications/inline-message))
* **Баннер с сообщением** предоставляет сведения и (при необходимости) простые призывы к действиям в виде баннера, который можно свернуть в одну строку, развернуть на несколько строк или закрыть. Баннеры с сообщениями рекомендуется использовать для информирования об обновлениях служб или отображения полезных советов при запуске надстройки. ([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Notification_messagebanner.pdf "PDF"), [код](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/notifications/message-banner))
* **Индикатор выполнения** показывает ход выполнения длительных синхронных процессов, например задач по настройке, которые необходимо выполнить, прежде чем пользователь сможет выполнить дальнейшие действия. Это отдельная промежуточная страница, на которой также используется фирменная символика надстройки. Используйте индикатор выполнения, если процесс может периодически отправлять сведения о том, сколько времени осталось до его завершения. ([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Notification_progress.pdf "PDF"), [код](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/notifications/progress-bar))
* **Индикатор работы** указывает на то, что выполняется длительный синхронный процесс, но не указывает, сколько времени осталось до его завершения. Это отдельная промежуточная страница, на которой также используется фирменная символика надстройки. Используйте индикатор работы, если надстройка не может достоверно сообщить, сколько времени необходимо для завершения процесса. ([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Notification_progress.pdf "PDF"), [код](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/notifications/spinner))
* **Всплывающее уведомление** содержит краткое сообщение, исчезающее через несколько секунд. Так как пользователь может и не увидеть такое сообщение, всплывающие уведомления используются для отображения несущественной информации. Это хороший способ уведомлять пользователей о событиях в удаленной системе, например о получении электронного письма. ([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Notification_toast.pdf "PDF"), [код](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/notifications/toast))

 <table>
 <tr><th>Внедренное диалоговое окно</th><th>Встроенное сообщение</th><th>Баннер с сообщением</th></tr>
 <tr><td>![embedded dialog" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/embedded.dialog.PNG)</td><td>![inline message" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/inline.message.PNG)</td><td>![message banner" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/message.banner.PNG)</td></tr>
 </table>

 <table>
 <tr><th>Индикатор выполнения</th><th>Индикатор работы</th><th>Всплывающее уведомление</th></tr>
 <tr><td>![progress bar" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/progress.bar.PNG)</td><td>![spinner" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/spinner.PNG)</td><td>![toast" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/toast.PNG)</td></tr>
 </table>

## Известные проблемы

* При выполнении некоторых файлов с кодом за пределами проекта надстройки возникает ошибка JavaScript. 
	* Решение: добавьте эти файлы в проект Office. 
	
## Дополнительные ресурсы

* [Рекомендации по разработке надстроек Office](https://dev.office.com/docs/add-ins/design/add-in-development-best-practices)
* [Office UI Fabric](http://dev.office.com/fabric/)
