
# Шаблоны взаимодействия надстроек для Office


Надстройки Office могут улучшить работу, связанную с созданием документов и другими офисными задачами, а также подключить контент ведущих приложений Office к более крупным рабочим процессам, основанным на веб-интерфейсе. К надстройкам области задач, контентным надстройкам и надстройкам Outlook применяется ряд типичных сценариев. В этой статье описаны некоторые из самых распространенных сценариев и представлены рекомендуемые шаблоны взаимодействия для пользовательского интерфейса надстройки. Вы можете разбивать на части, совмещать и смешивать эти шаблоны согласно своим уникальным сценариям.

 **Типичные сценарии надстроек**

| Тип надстройки | Типовые сценарии |
| ------ | ------ |
|  Контент  |  визуализация данных; <br> виджеты и средства.  |
|  Область задач  |  Преобразование и обработка данных <br> Эффективная разработка <br> Поиск содержимого и вставка данных <br> Публикация и отправка контента в веб-службу  |
|  Outlook  |  взаимодействие почтового контента и внешнего приложения; <br> предоставление дополнительной информации о контенте в сообщении или сведениях о встрече; <br> предоставление сведений, позволяющих повысить производительность.  |

## Визуализация данных с помощью контентной надстройки


В этом примере показана контентная надстройка для Excel, которое создает диаграмму на основе данных в электронной таблице.

В этом шаблоне взаимодействия надстройка не активируется, пока вы не выберите и не привяжите данные для создания диаграммы. Важно сообщить цель приложения и инструкции по его активации в первоначальном представлении надстройки. 

**Контентная надстройка для Excel, создающая диаграмму на основе данных в электронной таблице**
<br>
![Контентное приложение для Excel, создающее диаграмму на основе данных в электронной таблице](../../../images/off15appUXFig01.png)
<br>
<ul><li><p>Чтобы подчеркнуть, что сначала нужно выполнить действие, покажите инструкции вместе с отключенной кнопкой (A).</p></li><li><p>После выбора диапазона ячеек кнопка <span class="ui">Создать диаграмму</span> становится активной (B–C).</p></li><li><p>Алгоритм визуализации заполняет контейнер и заменяет предыдущее представление (D).</p></li><li><p>В нижней части надстройки отображаются дополнительные элементы пользовательского интерфейса вместе с кнопкой настроек (шестеренка), позволяющей открыть представление, в котором можно сбросить или изменить параметры надстройки.</p></li></ul>Подходит в следующих случаях:
<ul><li><p>Надстройки, которые требуют, чтобы пользователь выбрал данные перед активацией.</p></li></ul>

## Преобразование контента с помощью надстройки области задач


В этом примере показана надстройка области задач, которая переводит текст в документе на другой язык.

В этом шаблоне взаимодействия сначала нужно выбрать в документе текст, который требуется перевести.

**Надстройка области задач, которая переводит текст в документе на другой язык**
<br>
![Приложение области задач, которое переводит текст в документе на другой язык](../../../images/off15appUXFig02.png)
<br>
<ul><li><p>Расскажите о назначении надстройки в заголовке и подскажите, что сначала нужно выделить текст (A).</p></li><li><p>Меню языков и кнопка <span class="ui">Перевести</span> отключены, чтобы подчеркнуть, что пользователь должен выполнить действие перед продолжением. Как только пользователь выделяет текст в документе, эти два элемента становятся активными (D).</p></li><li><p>После нажатия кнопки <span class="ui">Перевести</span> отображается пользовательский интерфейс, переведенный текст и кнопка для вставки его в документ (E).</p></li><li><p>Можно показать кнопку <span class="ui">Очистить</span> или <span class="ui">Сброс</span>, которые открывают первоначальное представление.</p></li></ul>Подходит в следующих случаях:
<ul><li><p>Надстройки, которые требуют, чтобы пользователь выбрал данные перед активацией.</p></li><li><p>Пользовательский интерфейс разворачивается или становится видимым по мере прохождения сценария.</p></li></ul>

## Обработка данных с помощью надстройки области задач


В этом примере показана надстройка области задач, проверяющая данные в Excel.

В этом шаблоне взаимодействия для начала необходимо выбрать диапазон ячеек в электронной таблице.

**Надстройка области задач, проверяющая данные в Excel**
<br>
![Приложение области задач, проверяющее данные в Excel](../../../images/off15appUXFig03.png)
<br>
<ul><li><p>Назначение надстройки описано в заголовке. Инструкции помогут вам начать работу.</p></li><li><p>Кнопка <span class="ui">Отправить выделенные данные</span> отключена, чтобы подчеркнуть, что перед продолжением нужно выполнить действие (A).</p></li><li><p>После выделения диапазона ячеек на листе (B) кнопка <span class="ui">Отправить выделенные данные</span> активируется.</p></li><li><p>После нажатия этой кнопки пользовательский интерфейс заменяется выделенным диапазоном ячеек, индикатором выполнения и кнопкой <span class="ui">Отмена</span>.</p></li><li><p>Индикатор выполнения сообщает состояние процесса, а кнопка <span class="ui">Отмена</span> позволяет прервать процесс (D).</p></li><li><p>Когда процесс заканчивается, автоматически отображаются результаты (E). При выборе элемента в списке активируется соответствующая ячейка в электронной таблице.</p></li></ul>Подходит в следующих случаях:
<ul><li><p>Процессы, выполняющиеся неопределенное время.</p></li></ul>

## Анализ контента с помощью надстройки области задач


В этом примере показана надстройка области задач, отображающая определения слов по мере их ввода.

В этом шаблоне взаимодействия сначала нужно выбрать текст в документе, чтобы увидеть результаты.

**Надстройка области задач, которая показывает определения слов по мере их ввода**
<br>
![Приложение области задач, которое показывает определения слов по мере их ввода](../../../images/off15appUXFig04.png)
<br>
<ul><li><p>В заголовке объясняется назначение надстройки и как начать с ней работу (A).</p></li><li><p>Автоматический поиск включен по умолчанию, как и возможность для его отключения (B).</p></li><li><p>После выделения надстройки отображает соответствующий контент (D).</p></li><li><p>Предоставляется ссылка для получения более подробных сведений (E).</p></li></ul>Подходит в следующих случаях:
<ul><li><p>Надстройки, которые автоматически возвращают контент по мере ввода.</p></li><li><p>Надстройки, которые требуют выделить контент перед активацией.</p></li></ul>

## Поиск контента с помощью надстройки области задач


В этом примере показана надстройка области задач для поиска контента.

В этом шаблоне взаимодействия вы вводите строку в поле поиска или выбираете элемент из списка контента.

**Надстройка области задач для поиска содержимого**
<br>
![Приложение области задач для поиска содержимого](../../../images/off15appUXFig05.png)
<br>
<ul><li><p>Главное окно содержит поле <span class="ui">Поиск</span> (A) и список рекомендуемого содержимого (B).</p></li><li><p>При вводе в поле поиска строки значок поиска заменяется значком закрытия (C).</p></li><li><p>Если щелкнуть значок закрытия, вы вернетесь в начальное представление.</p></li></ul>Подходит в следующих случаях:
<ul><li><p>Надстройки, которые автоматически возвращают контент по мере ввода.</p></li><li><p>Надстройки, которые требуют выделить контент перед активацией.</p></li></ul>

## Вставка мультимедиа с помощью надстройки области задач


В этом шаблоне взаимодействия вы можете выбрать в результатах поиска изображение для вставки в документ.

**Надстройка области задач для вставки изображения**
<br>
![Приложение области задач для вставки изображения](../../../images/off15appUXFig06.png)
<br>
<ul><li><p>Вы отфильтровали список результатов поиска (A) и выбрали содержимое для вставки (B).</p></li><li><p>Подробное представление выбранного контента отображается (C) с кнопкой, которая возвращает вас обратно в список.</p></li><li><p>Кнопка <span class="ui">Вставить фотографию</span> располагается в нижнем колонтитуле (D). Если нажать эту кнопку, изображение вставляется в документ.</p></li><li><p>Вставленный контент (E) также включает краткое описание источника изображения. </p></li><li><p>Пользовательский интерфейс надстройки визуально подтверждает успешность действия.</p></li></ul>Подходит в следующих случаях:
<ul><li><p>надстройки для вставки контента;</p></li></ul>

## Вставка выбранного текста с помощью надстройки области задач


В этом шаблоне взаимодействия вы можете выбрать в результатах поиска текст для вставки в документ.

**Надстройка области задач для вставки текста**
<br>
![Приложение области задач для вставки текста](../../../images/off15appUXFig07.png)
<br>
<ul><li><p>Вы уже нашли содержимое (A).</p></li><li><p>Отключенная кнопка <span class="ui">Вставить выделенный фрагмент</span> отображается в нижнем колонтитуле (B).</p></li><li><p>Если выбрать строку текста (C), кнопка <span class="ui">Вставить выделенный фрагмент</span> становится активной.</p></li><li><p>После нажатия этой кнопки надстройка вставляет выделенный текст в документ со ссылкой на источник контента (E).</p></li></ul>Подходит в следующих случаях:
<ul><li><p>Надстройки для проведения исследований и вставки контента.</p></li></ul>

## Публикация в веб-службе с помощью надстройки области задач


В этом примере описана надстройка области задач для публикации документа в качестве сообщения блога.

В этом шаблоне взаимодействия вы завершили написание контента в документе и хотите опубликовать его в блоге.

**Надстройка области задач для публикации документа в блоге**
<br>
![Приложение области задач для публикации документа в блоге](../../../images/off15appUXFig08.png)
<br>
<ul><li><p>Сначала отображается форма для ввода учетных данных (A).</p></li><li><p>Предоставляются ссылки для создания учетной записи и решения типичных проблем входа (B). Если выбрать ссылку, в браузере откроются следующие страницы.</p></li><li><p>После входа в систему надстройка отображает форму для создания записи блога (C).</p></li><li><p>Имя учетной записи, в которую вы вошли (и в которой будет выполняться публикация), отображается в верхней части надстройки. Надстройка предоставляет ссылку для предварительного просмотра записи (D). Если перейти по ней, в браузере будет показан документ для предварительного просмотра.</p></li><li><p>Если нажать кнопку <span class="ui">Создать запись</span>, надстройка покажет представление, подтверждающее, что контент документа опубликован (E).</p></li><li><p>Надстройка предоставляет ссылку для просмотра сообщения в браузере (F), а также кнопку для создания другого сообщения (G).</p></li></ul>Подходит в следующих случаях:
<ul><li><p>надстройки, которые выдают, публикуют или распространяют контент в социальных сетях, блогах и веб-службах;</p></li><li><p>надстройки, которые требуют входа в службу.</p></li></ul>

## Получение дополнительных сведений о пользователях с помощью надстройки Outlook


 **Пример 1**

**Страница результатов и сведений**
<br>
![Страница результатов и сведений](../../../images/off15appUXFig09.jpg)
<br>
Подходит в следующих случаях:
<ul><li><p>Раскрытие всей полноты контента, когда имеются большие наборы данных, что полезно в презентационных целях.</p></li><li><p>Страницы со сведениями, которые используют полный размер контейнера надстройки</p></li><li><p>Модели навигации, которые используют схему "назад и вперед".</p></li></ul>
 **Пример 2**

**Страница сведений с постоянной навигацией**
<br>
![Страница сведений с постоянной навигацией](../../../images/off15appUXFig10.jpg)
<br>
Подходит в следующих случаях:
<ul><li><p>Отображение по умолчанию первого результата из набора данных.</p></li><li><p>Структуры навигации, которые работают подобно вкладкам (линейная навигация на одном уровне).</p></li><li><p>Уменьшение действий по выделению благодаря постоянной доступности навигационных элементов.</p></li><li><p>Предоставление пространства для постоянного отображения навигационных элементов.</p></li></ul>

## Получение дополнительных сведений о контенте с помощью надстройки Outlook


 **Пример 1**

**Страница результатов и сведений**
<br>
![Страница результатов и сведений](../../../images/off15appUXFig11.jpg)
<br>
Подходит в следующих случаях:
<ul><li><p>Раскрытие всей полноты контента, когда имеются большие наборы данных, что полезно в презентационных целях.</p></li><li><p>Требует сделать выбор или выделить фрагмент перед отображением дополнительных сведений.</p></li><li><p>Страницы со сведениями, которые используют полный размер контейнера надстройки.</p></li><li><p>Модели навигации, которые используют схему "назад и вперед".</p></li></ul>
 **Пример 2**

**Страница сведений со дополнительным содержимым**
<br>
![Страница сведений со дополнительным содержимым](../../../images/off15appUXFig12.jpg)
<br>
Подходит в следующих случаях:
<ul><li><p>Требуется акцентировать одну из частей контента.</p></li><li><p>Предоставление контента без взаимодействия с пользователем.</p></li><li><p>Постоянная навигация (может быть добавлена в эту модель для упрощения навигации).</p></li></ul>

## Подключение к веб-службе и предоставление данных


В этих примерах показаны шаблоны взаимодействия для получения данных и контента из веб-службы. Их можно использовать во всех трех типах надстроек: контентных, надстроек Outlook и области задач.

 **Пример 1**

**Карусель**
<br>
![Карусель](../../../images/off15appUXFig13.jpg)
<br>
Подходит в следующих случаях:
<ul><li><p>Небольшие объемы данных, которые могут отображаться в виде одного элемента или в группах.</p></li><li><p>Предоставление контента в виде галереи, например слайд-шоу или галереи изображений.</p></li><li><p>Когда хорошо работает модель навигации "следующий — предыдущий".</p></li></ul>
 **Пример 2**

**Мастер**
<br>
![Мастер](../../../images/off15appUXFig14.jpg)
<br>
Подходит в следующих случаях:
<ul><li><p>Контент, который должен отображаться в определенном порядке.</p></li><li><p>Контент большого объема, который лучше воспринимается в виде последовательности небольших частей.</p></li><li><p>Книгоподобный интерфейс отображения.</p></li><li><p>Для выполнения задачи требуется последовательность действий.</p></li></ul>
 **Пример 3**

**Форма, результаты и сведения**
<br>
![Форма, результаты и сведения](../../../images/off15appUXFig15.jpg)
<br>
Подходит в следующих случаях:
<ul><li><p>надстройки, требующие ввода данных;</p></li><li><p>входная точка для шаблона результатов и сведений.</p></li></ul>

## Дополнительные ресурсы



- [Рекомендации по проектированию надстроек Office](../add-in-design.md)
    