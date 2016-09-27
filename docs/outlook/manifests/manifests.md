
# Манифесты надстроек Outlook

Надстройка Outlook состоит из двух компонентов: XML-манифеста надстройки и веб-страницы, поддерживаемой библиотекой JavaScript для надстроек Office (office.js). В манифесте описывается, как надстройка интегрируется с клиентами Outlook. В настоящее время есть три версии схемы манифеста, включая **VersionOverrides**. Рекомендуем использовать схему манифеста версии 1.1 и **VersionOverrides** 1.0. Пример приведен ниже.

 >**Примечание**  Все значения URL-адресов в следующем примере начинаются со строки "YOUR_WEB_SERVER". Это значение — заполнитель. В фактическом допустимом манифесте эти значения будут содержать действительные URL-адреса с префиксом HTTPS.




```XML
<?xml version="1.0" encoding="UTF-8" ?>
<!--Created:cb85b80c-f585-40ff-8bfc-12ff4d0e34a9-->
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0"
  xsi:type="MailApp">
  <Id>7164e750-dc86-49c0-b548-1bac57abdc7c</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Microsoft Outlook Dev Center</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Add-in Command Demo" />
  <Description DefaultValue="Adds command buttons to the ribbon in Outlook"/>
  <IconUrl DefaultValue="YOUR_WEB_SERVER/images/blue-64.png" />
  <HighResolutionIconUrl DefaultValue="YOUR_WEB_SERVER/images/blue-80.png" />
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets>
      <Set Name="MailBox" MinVersion="1.1" />
    </Sets>
  </Requirements>
  <!-- These elements support older clients that don't support add-in commands -->
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <!-- NOTE: Just reusing the read taskpane page that is invoked by the button
             on the ribbon in clients that support add-in commands. You can 
             use a completely different page if desired -->
        <SourceLocation DefaultValue="YOUR_WEB_SERVER/AppRead/TaskPane/TaskPane.html"/>
        <RequestedHeight>450</RequestedHeight>
      </DesktopSettings>
    </Form>
    <Form xsi:type="ItemEdit">
      <DesktopSettings>
        <SourceLocation DefaultValue="YOUR_WEB_SERVER/AppCompose/Home/Home.html"/>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadWriteItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" />
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit" />
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
  </Rule>
  <DisableEntityHighlighting>false</DisableEntityHighlighting>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">

    <Requirements>
      <bt:Sets DefaultMinVersion="1.3">
        <bt:Set Name="Mailbox" />
      </bt:Sets>
    </Requirements>
    <Hosts>
        <Host Name="Mailbox" />
    </Hosts>
    <Requirements>
        <Sets>
            <Set Name="MailBox" MinVersion="1.1" />
        </Sets>
    </Requirements>
    <!-- These elements support older clients that don't support add-in commands -->
    <FormSettings>
        <Form xsi:type="ItemRead">
            <DesktopSettings>
                <!-- NOTE: Just reusing the read taskpane page that is invoked by the button
                     on the ribbon in clients that support add-in commands. You can
                     use a completely different page if desired -->
                <SourceLocation DefaultValue="YOUR_WEB_SERVER/AppRead/TaskPane/TaskPane.html" />
                <RequestedHeight>450</RequestedHeight>
            </DesktopSettings>
        </Form>
        <Form xsi:type="ItemEdit">
            <DesktopSettings>
                <SourceLocation DefaultValue="YOUR_WEB_SERVER/AppCompose/Home/Home.html" />
            </DesktopSettings>
        </Form>
    </FormSettings>
    <Permissions>ReadWriteItem</Permissions>
    <Rule xsi:type="RuleCollection" Mode="Or">
        <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" />
        <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit" />
        <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
        <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
    </Rule>
    <DisableEntityHighlighting>false</DisableEntityHighlighting>

    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">

        <Requirements>
            <bt:Sets DefaultMinVersion="1.3">
                <bt:Set Name="Mailbox" />
            </bt:Sets>
        </Requirements>
        <Hosts>
            <Host xsi:type="MailHost">

                <DesktopFormFactor>
                    <FunctionFile resid="functionFile" />

                    <!-- Custom pane, only applies to read form -->
                    <ExtensionPoint xsi:type="CustomPane">
                        <RequestedHeight>100</RequestedHeight>
                        <SourceLocation resid="customPaneUrl" />
                        <Rule xsi:type="RuleCollection" Mode="Or">
                            <Rule xsi:type="ItemIs" ItemType="Message" />
                            <Rule xsi:type="ItemIs" ItemType="AppointmentAttendee" />
                        </Rule>
                    </ExtensionPoint>

                    <!-- Message compose form -->
                    <ExtensionPoint xsi:type="MessageComposeCommandSurface">
                        <OfficeTab id="TabDefault">
                            <Group id="msgComposeDemoGroup">
                                <Label resid="groupLabel" />
                                <!-- Function (UI-less) button -->
                                <Control xsi:type="Button" id="msgComposeFunctionButton">
                                    <Label resid="funcComposeButtonLabel" />
                                    <Supertip>
                                        <Title resid="funcComposeSuperTipTitle" />
                                        <Description resid="funcComposeSuperTipDescription" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="blue-icon-16" />
                                        <bt:Image size="32" resid="blue-icon-32" />
                                        <bt:Image size="80" resid="blue-icon-80" />
                                    </Icon>
                                    <Action xsi:type="ExecuteFunction">
                                        <FunctionName>addDefaultMsgToBody</FunctionName>
                                    </Action>
                                </Control>
                                <!-- Menu (dropdown) button -->
                                <Control xsi:type="Menu" id="msgComposeMenuButton">
                                    <Label resid="menuComposeButtonLabel" />
                                    <Supertip>
                                        <Title resid="menuComposeSuperTipTitle" />
                                        <Description resid="menuComposeSuperTipDescription" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="red-icon-16" />
                                        <bt:Image size="32" resid="red-icon-32" />
                                        <bt:Image size="80" resid="red-icon-80" />
                                    </Icon>
                                    <Items>
                                        <Item id="msgComposeMenuItem1">
                                            <Label resid="menuItem1ComposeLabel" />
                                            <Supertip>
                                                <Title resid="menuItem1ComposeLabel" />
                                                <Description resid="menuItem1ComposeTip" />
                                            </Supertip>
                                            <Icon>
                                                <bt:Image size="16" resid="red-icon-16" />
                                                <bt:Image size="32" resid="red-icon-32" />
                                                <bt:Image size="80" resid="red-icon-80" />
                                            </Icon>
                                            <Action xsi:type="ExecuteFunction">
                                                <FunctionName>addMsg1ToBody</FunctionName>
                                            </Action>
                                        </Item>
                                        <Item id="msgComposeMenuItem2">
                                            <Label resid="menuItem2ComposeLabel" />
                                            <Supertip>
                                                <Title resid="menuItem2ComposeLabel" />
                                                <Description resid="menuItem2ComposeTip" />
                                            </Supertip>
                                            <Icon>
                                                <bt:Image size="16" resid="red-icon-16" />
                                                <bt:Image size="32" resid="red-icon-32" />
                                                <bt:Image size="80" resid="red-icon-80" />
                                            </Icon>
                                            <Action xsi:type="ExecuteFunction">
                                                <FunctionName>addMsg2ToBody</FunctionName>
                                            </Action>
                                        </Item>
                                        <Item id="msgComposeMenuItem3">
                                            <Label resid="menuItem3ComposeLabel" />
                                            <Supertip>
                                                <Title resid="menuItem3ComposeLabel" />
                                                <Description resid="menuItem3ComposeTip" />
                                            </Supertip>
                                            <Icon>
                                                <bt:Image size="16" resid="red-icon-16" />
                                                <bt:Image size="32" resid="red-icon-32" />
                                                <bt:Image size="80" resid="red-icon-80" />
                                            </Icon>
                                            <Action xsi:type="ExecuteFunction">
                                                <FunctionName>addMsg3ToBody</FunctionName>
                                            </Action>
                                        </Item>
                                    </Items>
                                </Control>
                                <!-- Task pane button -->
                                <Control xsi:type="Button" id="msgComposeOpenPaneButton">
                                    <Label resid="paneComposeButtonLabel" />
                                    <Supertip>
                                        <Title resid="paneComposeSuperTipTitle" />
                                        <Description resid="paneComposeSuperTipDescription" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="green-icon-16" />
                                        <bt:Image size="32" resid="green-icon-32" />
                                        <bt:Image size="80" resid="green-icon-80" />
                                    </Icon>
                                    <Action xsi:type="ShowTaskpane">
                                        <SourceLocation resid="composeTaskPaneUrl" />
                                    </Action>
                                </Control>
                            </Group>
                        </OfficeTab>
                    </ExtensionPoint>

                    <!-- Appointment compose form -->
                    <ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
                        <OfficeTab id="TabDefault">
                            <Group id="apptComposeDemoGroup">
                                <Label resid="groupLabel" />
                                <!-- Function (UI-less) button -->
                                <Control xsi:type="Button" id="apptComposeFunctionButton">
                                    <Label resid="funcComposeButtonLabel" />
                                    <Supertip>
                                        <Title resid="funcComposeSuperTipTitle" />
                                        <Description resid="funcComposeSuperTipDescription" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="blue-icon-16" />
                                        <bt:Image size="32" resid="blue-icon-32" />
                                        <bt:Image size="80" resid="blue-icon-80" />
                                    </Icon>
                                    <Action xsi:type="ExecuteFunction">
                                        <FunctionName>addDefaultMsgToBody</FunctionName>
                                    </Action>
                                </Control>
                                <!-- Menu (dropdown) button -->
                                <Control xsi:type="Menu" id="apptComposeMenuButton">
                                    <Label resid="menuComposeButtonLabel" />
                                    <Supertip>
                                        <Title resid="menuComposeSuperTipTitle" />
                                        <Description resid="menuComposeSuperTipDescription" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="red-icon-16" />
                                        <bt:Image size="32" resid="red-icon-32" />
                                        <bt:Image size="80" resid="red-icon-80" />
                                    </Icon>
                                    <Items>
                                        <Item id="apptComposeMenuItem1">
                                            <Label resid="menuItem1ComposeLabel" />
                                            <Supertip>
                                                <Title resid="menuItem1ComposeLabel" />
                                                <Description resid="menuItem1ComposeTip" />
                                            </Supertip>
                                            <Icon>
                                                <bt:Image size="16" resid="red-icon-16" />
                                                <bt:Image size="32" resid="red-icon-32" />
                                                <bt:Image size="80" resid="red-icon-80" />
                                            </Icon>
                                            <Action xsi:type="ExecuteFunction">
                                                <FunctionName>addMsg1ToBody</FunctionName>
                                            </Action>
                                        </Item>
                                        <Item id="apptComposeMenuItem2">
                                            <Label resid="menuItem2ComposeLabel" />
                                            <Supertip>
                                                <Title resid="menuItem2ComposeLabel" />
                                                <Description resid="menuItem2ComposeTip" />
                                            </Supertip>
                                            <Icon>
                                                <bt:Image size="16" resid="red-icon-16" />
                                                <bt:Image size="32" resid="red-icon-32" />
                                                <bt:Image size="80" resid="red-icon-80" />
                                            </Icon>
                                            <Action xsi:type="ExecuteFunction">
                                                <FunctionName>addMsg2ToBody</FunctionName>
                                            </Action>
                                        </Item>
                                        <Item id="apptComposeMenuItem3">
                                            <Label resid="menuItem3ComposeLabel" />
                                            <Supertip>
                                                <Title resid="menuItem3ComposeLabel" />
                                                <Description resid="menuItem3ComposeTip" />
                                            </Supertip>
                                            <Icon>
                                                <bt:Image size="16" resid="red-icon-16" />
                                                <bt:Image size="32" resid="red-icon-32" />
                                                <bt:Image size="80" resid="red-icon-80" />
                                            </Icon>
                                            <Action xsi:type="ExecuteFunction">
                                                <FunctionName>addMsg3ToBody</FunctionName>
                                            </Action>
                                        </Item>
                                    </Items>
                                </Control>
                                <!-- Task pane button -->
                                <Control xsi:type="Button" id="apptComposeOpenPaneButton">
                                    <Label resid="paneComposeButtonLabel" />
                                    <Supertip>
                                        <Title resid="paneComposeSuperTipTitle" />
                                        <Description resid="paneComposeSuperTipDescription" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="green-icon-16" />
                                        <bt:Image size="32" resid="green-icon-32" />
                                        <bt:Image size="80" resid="green-icon-80" />
                                    </Icon>
                                    <Action xsi:type="ShowTaskpane">
                                        <SourceLocation resid="composeTaskPaneUrl" />
                                    </Action>
                                </Control>
                            </Group>
                        </OfficeTab>
                    </ExtensionPoint>

                    <!-- Message read form -->
                    <ExtensionPoint xsi:type="MessageReadCommandSurface">
                        <OfficeTab id="TabDefault">
                            <Group id="msgReadDemoGroup">
                                <Label resid="groupLabel" />
                                <!-- Function (UI-less) button -->
                                <Control xsi:type="Button" id="msgReadFunctionButton">
                                    <Label resid="funcReadButtonLabel" />
                                    <Supertip>
                                        <Title resid="funcReadSuperTipTitle" />
                                        <Description resid="funcReadSuperTipDescription" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="blue-icon-16" />
                                        <bt:Image size="32" resid="blue-icon-32" />
                                        <bt:Image size="80" resid="blue-icon-80" />
                                    </Icon>
                                    <Action xsi:type="ExecuteFunction">
                                        <FunctionName>getSubject</FunctionName>
                                    </Action>
                                </Control>
                                <!-- Menu (dropdown) button -->
                                <Control xsi:type="Menu" id="msgReadMenuButton">
                                    <Label resid="menuReadButtonLabel" />
                                    <Supertip>
                                        <Title resid="menuReadSuperTipTitle" />
                                        <Description resid="menuReadSuperTipDescription" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="red-icon-16" />
                                        <bt:Image size="32" resid="red-icon-32" />
                                        <bt:Image size="80" resid="red-icon-80" />
                                    </Icon>
                                    <Items>
                                        <Item id="msgReadMenuItem1">
                                            <Label resid="menuItem1ReadLabel" />
                                            <Supertip>
                                                <Title resid="menuItem1ReadLabel" />
                                                <Description resid="menuItem1ReadTip" />
                                            </Supertip>
                                            <Icon>
                                                <bt:Image size="16" resid="red-icon-16" />
                                                <bt:Image size="32" resid="red-icon-32" />
                                                <bt:Image size="80" resid="red-icon-80" />
                                            </Icon>
                                            <Action xsi:type="ExecuteFunction">
                                                <FunctionName>getItemClass</FunctionName>
                                            </Action>
                                        </Item>
                                        <Item id="msgReadMenuItem2">
                                            <Label resid="menuItem2ReadLabel" />
                                            <Supertip>
                                                <Title resid="menuItem2ReadLabel" />
                                                <Description resid="menuItem2ReadTip" />
                                            </Supertip>
                                            <Icon>
                                                <bt:Image size="16" resid="red-icon-16" />
                                                <bt:Image size="32" resid="red-icon-32" />
                                                <bt:Image size="80" resid="red-icon-80" />
                                            </Icon>
                                            <Action xsi:type="ExecuteFunction">
                                                <FunctionName>getDateTimeCreated</FunctionName>
                                            </Action>
                                        </Item>
                                        <Item id="msgReadMenuItem3">
                                            <Label resid="menuItem3ReadLabel" />
                                            <Supertip>
                                                <Title resid="menuItem3ReadLabel" />
                                                <Description resid="menuItem3ReadTip" />
                                            </Supertip>
                                            <Icon>
                                                <bt:Image size="16" resid="red-icon-16" />
                                                <bt:Image size="32" resid="red-icon-32" />
                                                <bt:Image size="80" resid="red-icon-80" />
                                            </Icon>
                                            <Action xsi:type="ExecuteFunction">
                                                <FunctionName>getItemID</FunctionName>
                                            </Action>
                                        </Item>
                                    </Items>
                                </Control>
                                <!-- Task pane button -->
                                <Control xsi:type="Button" id="msgReadOpenPaneButton">
                                    <Label resid="paneReadButtonLabel" />
                                    <Supertip>
                                        <Title resid="paneReadSuperTipTitle" />
                                        <Description resid="paneReadSuperTipDescription" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="green-icon-16" />
                                        <bt:Image size="32" resid="green-icon-32" />
                                        <bt:Image size="80" resid="green-icon-80" />
                                    </Icon>
                                    <Action xsi:type="ShowTaskpane">
                                        <SourceLocation resid="readTaskPaneUrl" />
                                    </Action>
                                </Control>
                            </Group>
                        </OfficeTab>
                    </ExtensionPoint>

                    <!-- Appointment read form -->
                    <ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
                        <OfficeTab id="TabDefault">
                            <Group id="apptReadDemoGroup">
                                <Label resid="groupLabel" />
                                <!-- Function (UI-less) button -->
                                <Control xsi:type="Button" id="apptReadFunctionButton">
                                    <Label resid="funcReadButtonLabel" />
                                    <Supertip>
                                        <Title resid="funcReadSuperTipTitle" />
                                        <Description resid="funcReadSuperTipDescription" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="blue-icon-16" />
                                        <bt:Image size="32" resid="blue-icon-32" />
                                        <bt:Image size="80" resid="blue-icon-80" />
                                    </Icon>
                                    <Action xsi:type="ExecuteFunction">
                                        <FunctionName>getSubject</FunctionName>
                                    </Action>
                                </Control>
                                <!-- Menu (dropdown) button -->
                                <Control xsi:type="Menu" id="apptReadMenuButton">
                                    <Label resid="menuReadButtonLabel" />
                                    <Supertip>
                                        <Title resid="menuReadSuperTipTitle" />
                                        <Description resid="menuReadSuperTipDescription" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="red-icon-16" />
                                        <bt:Image size="32" resid="red-icon-32" />
                                        <bt:Image size="80" resid="red-icon-80" />
                                    </Icon>
                                    <Items>
                                        <Item id="apptReadMenuItem1">
                                            <Label resid="menuItem1ReadLabel" />
                                            <Supertip>
                                                <Title resid="menuItem1ReadLabel" />
                                                <Description resid="menuItem1ReadTip" />
                                            </Supertip>
                                            <Icon>
                                                <bt:Image size="16" resid="red-icon-16" />
                                                <bt:Image size="32" resid="red-icon-32" />
                                                <bt:Image size="80" resid="red-icon-80" />
                                            </Icon>
                                            <Action xsi:type="ExecuteFunction">
                                                <FunctionName>getItemClass</FunctionName>
                                            </Action>
                                        </Item>
                                        <Item id="apptReadMenuItem2">
                                            <Label resid="menuItem2ReadLabel" />
                                            <Supertip>
                                                <Title resid="menuItem2ReadLabel" />
                                                <Description resid="menuItem2ReadTip" />
                                            </Supertip>
                                            <Icon>
                                                <bt:Image size="16" resid="red-icon-16" />
                                                <bt:Image size="32" resid="red-icon-32" />
                                                <bt:Image size="80" resid="red-icon-80" />
                                            </Icon>
                                            <Action xsi:type="ExecuteFunction">
                                                <FunctionName>getDateTimeCreated</FunctionName>
                                            </Action>
                                        </Item>
                                        <Item id="apptReadMenuItem3">
                                            <Label resid="menuItem3ReadLabel" />
                                            <Supertip>
                                                <Title resid="menuItem3ReadLabel" />
                                                <Description resid="menuItem3ReadTip" />
                                            </Supertip>
                                            <Icon>
                                                <bt:Image size="16" resid="red-icon-16" />
                                                <bt:Image size="32" resid="red-icon-32" />
                                                <bt:Image size="80" resid="red-icon-80" />
                                            </Icon>
                                            <Action xsi:type="ExecuteFunction">
                                                <FunctionName>getItemID</FunctionName>
                                            </Action>
                                        </Item>
                                    </Items>
                                </Control>
                                <!-- Task pane button -->
                                <Control xsi:type="Button" id="apptReadOpenPaneButton">
                                    <Label resid="paneReadButtonLabel" />
                                    <Supertip>
                                        <Title resid="paneReadSuperTipTitle" />
                                        <Description resid="paneReadSuperTipDescription" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="green-icon-16" />
                                        <bt:Image size="32" resid="green-icon-32" />
                                        <bt:Image size="80" resid="green-icon-80" />
                                    </Icon>
                                    <Action xsi:type="ShowTaskpane">
                                        <SourceLocation resid="readTaskPaneUrl" />
                                    </Action>
                                </Control>
                            </Group>
                        </OfficeTab>
                    </ExtensionPoint>
                </DesktopFormFactor>
            </Host>
        </Hosts>

        <Resources>
            <bt:Images>
                <!-- Blue icon -->
                <bt:Image id="blue-icon-16" DefaultValue="YOUR_WEB_SERVER/images/blue-16.png" />
                <bt:Image id="blue-icon-32" DefaultValue="YOUR_WEB_SERVER/images/blue-32.png" />
                <bt:Image id="blue-icon-80" DefaultValue="YOUR_WEB_SERVER/images/blue-80.png" />
                <!-- Red icon -->
                <bt:Image id="red-icon-16" DefaultValue="YOUR_WEB_SERVER/images/red-16.png" />
                <bt:Image id="red-icon-32" DefaultValue="YOUR_WEB_SERVER/images/red-32.png" />
                <bt:Image id="red-icon-80" DefaultValue="YOUR_WEB_SERVER/images/red-80.png" />
                <!-- Green icon -->
                <bt:Image id="green-icon-16" DefaultValue="YOUR_WEB_SERVER/images/green-16.png" />
                <bt:Image id="green-icon-32" DefaultValue="YOUR_WEB_SERVER/images/green-32.png" />
                <bt:Image id="green-icon-80" DefaultValue="YOUR_WEB_SERVER/images/green-80.png" />
            </bt:Images>
            <bt:Urls>
                <bt:Url id="functionFile" DefaultValue="YOUR_WEB_SERVER/FunctionFile/Functions.html" />
                <bt:Url id="readTaskPaneUrl" DefaultValue="YOUR_WEB_SERVER/AppRead/TaskPane/TaskPane.html" />
                <bt:Url id="composeTaskPaneUrl" DefaultValue="YOUR_WEB_SERVER/AppCompose/TaskPane/TaskPane.html" />
                <bt:Url id="customPaneUrl" DefaultValue="YOUR_WEB_SERVER/AppRead/CustomPane/CustomPane.html" />
            </bt:Urls>
            <bt:ShortStrings>
                <bt:String id="groupLabel" DefaultValue="Add-in Demo" />
                <!-- Compose mode -->
                <bt:String id="funcComposeButtonLabel" DefaultValue="Insert default message" />
                <bt:String id="menuComposeButtonLabel" DefaultValue="Insert message" />
                <bt:String id="paneComposeButtonLabel" DefaultValue="Insert custom message" />

                <bt:String id="funcComposeSuperTipTitle" DefaultValue="Inserts the default message" />
                <bt:String id="menuComposeSuperTipTitle" DefaultValue="Choose a message to insert" />
                <bt:String id="paneComposeSuperTipTitle" DefaultValue="Enter your own text to insert" />

                <bt:String id="menuItem1ComposeLabel" DefaultValue="Hello World!" />
                <bt:String id="menuItem2ComposeLabel" DefaultValue="Add-in commands are cool!" />
                <bt:String id="menuItem3ComposeLabel" DefaultValue="Visit dev.outlook.com" />

                <!-- Read mode -->
                <bt:String id="funcReadButtonLabel" DefaultValue="Get subject" />
                <bt:String id="menuReadButtonLabel" DefaultValue="Get property" />
                <bt:String id="paneReadButtonLabel" DefaultValue="Display all properties" />

                <bt:String id="funcReadSuperTipTitle" DefaultValue="Gets the subject of the message or appointment" />
                <bt:String id="menuReadSuperTipTitle" DefaultValue="Choose a property to get" />
                <bt:String id="paneReadSuperTipTitle" DefaultValue="Get all properties" />

                <bt:String id="menuItem1ReadLabel" DefaultValue="Get item class" />
                <bt:String id="menuItem2ReadLabel" DefaultValue="Get date time created" />
                <bt:String id="menuItem3ReadLabel" DefaultValue="Get item ID" />
            </bt:ShortStrings>
            <bt:LongStrings>
                <!-- Compose mode -->
                <bt:String id="funcComposeSuperTipDescription" DefaultValue="Inserts text into body of the message or appointment. This is an example of a function button." />
                <bt:String id="menuComposeSuperTipDescription" DefaultValue="Inserts your choice of text into body of the message or appointment. This is an example of a drop-down menu button." />
                <bt:String id="paneComposeSuperTipDescription" DefaultValue="Opens a pane where you can enter text to insert in the body of the message or appointment. This is an example of a button that opens a task pane." />

                <bt:String id="menuItem1ComposeTip" DefaultValue="Inserts Hello World! into the body of the message or appointment." />
                <bt:String id="menuItem2ComposeTip" DefaultValue="Inserts Add-in commands are cool! into the body of the message or appointment." />
                <bt:String id="menuItem3ComposeTip" DefaultValue="Inserts Visit dev.outlook.com into the body of the message or appointment." />

                <!-- Read mode -->
                <bt:String id="funcReadSuperTipDescription" DefaultValue="Gets the subject of the message or appointment and displays it in the info bar. This is an example of a function button." />
                <bt:String id="menuReadSuperTipDescription" DefaultValue="Gets the selected property of the message or appointment and displays it in the info bar. This is an example of a drop-down menu button." />
                <bt:String id="paneReadSuperTipDescription" DefaultValue="Opens a pane displaying all available properties of the message or appointment. This is an example of a button that opens a task pane." />

                <bt:String id="menuItem1ReadTip" DefaultValue="Gets the item class of the message or appointment and displays it in the info bar." />
                <bt:String id="menuItem2ReadTip" DefaultValue="Gets the date and time the message or appointment was created and displays it in the info bar." />
                <bt:String id="menuItem3ReadTip" DefaultValue="Gets the item ID of the message or appointment and displays it in the info bar." />
            </bt:LongStrings>
        </Resources>
    </VersionOverrides>
</OfficeApp>

```


## Версии схемы

Не все клиенты Outlook поддерживают все новые функции, а некоторые пользователи применяют устаревшую версию Outlook. Версии схемы позволяют разработчикам создавать надстройки с обратной совместимостью, которые используют новейшие функции по мере их появления, но по-прежнему работают в устаревших версиях.

Наглядный пример — элемент манифеста  **VersionOverrides**. Все элементы, определенные в  **VersionOverrides**, заменяют соответствующий элемент в другой части манифеста. Это означает, что по мере возможности Outlook будет использовать содержимое раздела  **VersionOverrides** для установки параметров надстройки. Тем не менее, если версия Outlook не поддерживает определенную версию **VersionOverrides**, Outlook пропустит ее и будет использовать остальные сведения из манифеста. 

Этот подход означает, что разработчикам не требуется создавать несколько отдельных манифестов — все параметры можно задать в одном файле.

В настоящий момент доступны следующие версии схемы:


|||
|:-----|:-----|
|Версия|Описание|
|v1.0|Поддерживает версию 1.0 API JavaScript для Office. Для надстроек Outlook поддерживается форма чтения. |
|v1.1|Поддерживает версии 1.1 API JavaScript для Office и  **VersionOverrides**. Для надстроек Outlook поддерживается форма создания.|
|**VersionOverrides** 1.0|Поддерживает более поздние версии API JavaScript для Office. Поддерживаются команды надстроек.|
Эта статья охватывает требования для манифеста версии 1.1. Даже если в манифесте вашей надстройки используется элемент  **VersionOverrides**, все равно важно включить элементы манифеста 1.1, чтобы надстройка работала со старыми клиентами, которые не поддерживают  **VersionOverrides**.


## Корневой элемент

Корневой элемент манифеста надстройки Outlook —  **OfficeApp**. Этот элемент также объявляет пространство имен по умолчанию, версию схемы и тип надстройки. Поместите все остальные элементы манифеста между его открывающим и закрывающим тегами. Ниже приводится пример корневого элемента:


```XML
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0"
  xsi:type="MailApp">

  <!-- the rest of the manifest>

</OfficeApp>
```


## Версия

Это версия конкретной надстройки. Если разработчик обновит какой-либо элемент манифеста, версия также должна быть увеличена. Благодаря этому при установке нового манифеста будет переписан существующий, а пользователь получит новые функции. Если надстройка была отправлена в магазин, то новый манифест должен быть заново отправлен и проверен. Затем пользователи надстройки автоматически получат обновленный манифест в течение нескольких часов после его утверждения.

При изменении разрешений, запрашиваемых надстройкой, пользователям предлагается выполнить обновление и повторно согласиться на предоставление надстройке разрешений. Если администратор установил эту надстройку для всей организации, он должен будет дать свое согласие первым. До этого пользователям будут доступны только старые функции.


## VersionOverrides

Элемент **VersionOverrides** — это расположение данных о командах надстройки. Дополнительные сведения об этом элементе см. в статье [Определение команд надстройки в манифесте надстройки Outlook](../../outlook/manifests/define-add-in-commands.md).


## Локализация

Некоторые элементы надстройки (например, имя, описание и загружаемый URL-адрес) необходимо локализовать для разных языковых стандартов. Эти элементы легко локализовать, указав значение по умолчанию и переопределения для языкового стандарта в дочернем элементе **Resources** элемента **VersionOverrides**. Ниже показано, как переопределить изображение, URL-адрес и строку.


```XML
<Resources>
    <bt:Images>
      <bt:Image id="icon1_16x16" DefaultValue="https://contoso.com/images/app_icon_small.png" >
        <bt:Override Locale="ar-sa" Value="https://contoso.com/images/app_icon_small_arsa.png" />
        <!-- add information for other locales -->

    <bt:Urls>
      <bt:Url id="residDesktopFuncUrl" DefaultValue="https://contoso.com/urls/page_appcmdcode.html" >
        <bt:Override Locale="ar-sa" Value="https://contoso.com/urls/page_appcmdcode.html?lcid=ar-sa" />
        <!-- add information for other locales -->

    <bt:ShortStrings> 
      </bt:String>
      <bt:String id="residViewTemplates" DefaultValue="Launch My Add-in">
        <bt:Override Locale="ar-sa" Value="<add localized value here>" />
        <!-- add information for other locales -->
    </bt:ShortStrings>

  </Resources>
```

Справочник по схеме содержит полные сведения о том, какие элементы можно локализовать.


## Hosts

Надстройки Outlook задают элемент  **Hosts** следующим образом.


```XML
<OfficeApp>
...
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
...
</OfficeApp>
```

Он отличается от элемента  **Hosts** в разделе **VersionOverrides**, который рассматривается в статье [Определение команд надстройки Outlook в манифесте](../../outlook/manifests/define-add-in-commands.md).


## Requirements

Элемент  **Requirements** указывает набор API-интерфейсов, доступный надстройке. Для надстройки Outlook требуются набор требований Mailbox и версия 1.1 или выше. Последняя версия набора требований указана в справочнике по API. Дополнительные сведения о наборах требований см. в статье [API-интерфейсы Outlook](../../outlook/apis.md).

Элемент  **Requirements** также может присутствовать в элементе **VersionOverrides**, позволяя надстройке указывать другие требования при загрузке в клиентах, поддерживающих  **VersionOverrides**.

В следующем примере используется атрибут  **DefaultMinVersion** элемента **Sets**, чтобы запрашивался файл office.js версии 1.1 или выше, и атрибут  **MinVersion** элемента **Set**, чтобы запрашивался набор требований Mailbox версии 1.1.




```XML
<OfficeApp>
...
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="MailBox" MinVersion="1.1" />
    </Sets>
  </Requirements>
...
</OfficeApp>
```


## Параметры формы

Элемент  **FormSettings** используется устаревшими клиентами Outlook, которые поддерживают только схему версии 1.1, а не **VersionOverrides**. С помощью этого элемента разработчики указывают, как надстройка будет отображаться в таких клиентах. Он состоит из двух частей: **ItemRead** и **ItemEdit**.  **ItemRead** позволяет указать, как надстройка отображается при просмотре сообщений и встреч. **ItemEdit** описывает отображение надстройки при создании ответа, сообщения или встречи либо редактировании встречи организатором.

Эти параметры напрямую связаны с правилами активации в элементе  **Rule**. Если надстройка указывает, что она должна отображаться на сообщении в форме создания, то должна быть указана форма  **ItemEdit**.

Дополнительные сведения см. в [ по схеме для манифестов надстроек Office (версия 1.1)](../../overview/add-in-manifests.md).

## Домены приложений

Домен начальной страницы надстройки, заданной в элементе  **SourceLocation**, является доменом по умолчанию для этой надстройки. Если элементы  **AppDomains** и **AppDomain** не используются, а ваша надстройка попытается перейти к другому домену, в браузере откроется новое окно за пределами области надстройки. Чтобы надстройка могла переходить на другой домен в пределах области надстройки, добавьте элемент **AppDomains** и укажите каждый дополнительный домен в отдельном дочернем элементе **AppDomain** в манифесте надстройки.

В следующем примере домен  `https://www.contoso2.com` указан как второй домен, к которому надстройка может переходить в рамках области надстройки:




```XML
<OfficeApp>
...
  <AppDomains>
    <AppDomain>https://www.contoso2.com</AppDomain>
  </AppDomains>
...
</OfficeApp>
```

Домены надстроек также необходимы для обмена файлами cookie между всплывающим окном и надстройкой, запущенной в расширенном клиенте.


## Permissions

Элемент  **Permissions** содержит необходимые надстройке разрешения. В целом вам следует указать минимальные необходимые разрешения, требуемые для вашей надстройки, в зависимости от конкретных методов, которые вы собираетесь использовать. Например, для почтовой надстройки, которая активируется в форме создания и только считывает свойства элементов типа [item.requiredAttendees](../../../reference/outlook/Office.context.mailbox.item.md), но не записывает их и не вызывает метод [mailbox.makeEwsRequestAsync](../../../reference/outlook/Office.context.mailbox.md) для получения доступа к любым операциям веб-служб Exchange, следует указать разрешение **ReadItem**. Дополнительные сведения о доступных разрешениях см. в статье [Указание разрешений для доступа надстройки Outlook к почтовому ящику пользователя](../../outlook/understanding-outlook-add-in-permissions.md).


**Четырехуровневая модель разрешений для почтовых надстроек**

![4-уровневая модель разрешений для схемы почтовых приложений версии 1.1](../../../images/olowa15wecon_Permissions_4Tier.png)
```XML
<OfficeApp>
...
  <Permissions>ReadWriteItem</Permissions>
...
</OfficeApp>

```


## Правила активации

Правила активации указываются в элементе  **Rule**. Элемент  **Rule** может быть дочерним элементом **OfficeApp** в манифестах версии 1.1, а также дочерним элементом **ExtensionPoint** в разделе **VersionOverrides**. Дополнительные сведения об использовании этого элемента в  [VersionOverrides](../../outlook/manifests/define-add-in-commands.md) см. в статье **Определение команд надстройки Outlook в манифесте**.

Правила активации позволяют активировать надстройку при соблюдении одного или нескольких из представленных ниже условий в выбранном элементе.


- Тип элемента или класс сообщения
    
- Наличие известной сущности определенного типа, например адреса или номера телефона
    
- Совпадение с регулярным выражением в основном тексте, теме или электронном адресе отправителя
    
- Наличие вложения
    
Подробные сведения и примеры правил активации см. в статье [Правила активации для надстроек Outlook](../../outlook/manifests/activation-rules.md).


## Дальнейшие действия: команды надстроек


После определения основного манифеста [определите команды для вашей надстройки](../../outlook/manifests/define-add-in-commands.md). Команды надстроек представляют собой кнопки на ленте, с помощью которых пользователи могут легко и интуитивно активировать ваши надстройки. Дополнительные сведения см. в статье [Команды надстроек Outlook](../../outlook/add-in-commands-for-outlook.md).


## Дополнительные ресурсы



- [Надстройки Outlook](../../outlook/outlook-add-ins.md)
    
- [Правила активации для надстроек Outlook](../../outlook/manifests/activation-rules.md)
    
- [Локализация надстроек для Office](../../develop/localization.md)
    
- [Создание почтовой надстройки для Outlook, работающей на настольных компьютерах, планшетах и мобильных устройствах (схема версии 1.1)](http://msdn.microsoft.com/library/8d425fb3-8a7c-429d-87b3-8046e964b153%28Office.15%29.aspx)
    
- [Конфиденциальность, разрешения и безопасность для надстроек Outlook](../../outlook/privacy-and-security.md)
    
- [API-интерфейсы Outlook](../../outlook/apis.md)
    
- [XML-манифест надстроек для Office](../../overview/add-in-manifests.md)
    
- [Справка по схеме для манифестов надстроек Office (версия 1.1)](../../overview/add-in-manifests.md)
    
- [Типы элементов и классы сообщений](http://msdn.microsoft.com/library/15b709cc-7486-b6c7-88a3-4a4d8e0ab292%28Office.15%29.aspx)
    
- [Рекомендации по проектированию надстроек Office](../../design/add-in-design.md)
    
- [Общие сведения о разрешениях для надстройки Outlook](../../outlook/understanding-outlook-add-in-permissions.md)
    
- [Использование регулярных правил активации выражений для отображения надстройки Outlook](../../outlook/use-regular-expressions-to-show-an-outlook-add-in.md)
    
- [Сопоставление строк в элементе Outlook как известных сущностей](../../outlook/match-strings-in-an-item-as-well-known-entities.md)
    