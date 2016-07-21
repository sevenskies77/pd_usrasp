<%
'ПОТОМ INCLUDE ВСТАВИТЬ В ТЕКСТ
%>
<!--#INCLUDE FILE="consts.asp" -->
<%
'Set user text constants - redefine here your own values for TEXTANSI.TXT constants
'Out "***UserText.asp"
'Set view for Bank account list

' STS Alex 20:11 19.05.2009 перевод полей для STS_PurshaseOrder и STS_PaymentOrder

If CurrentClassDoc=DOCS_Money_Accounts Or Request("ClassDoc")=DOCS_Money_Accounts Or Request("ActDoc")=DOCS_Money_Accounts Then ' - document category to be processed
If RUS()<>"RUS" Then 
			'Substitute TEXTANSI.TXT constants
			DOCS_DocID="Account No."
			DOCS_Name="Account name"
			DOCS_AmountDoc="Current balance"
			DOCS_Details1="Account details"
			DOCS_StatusPayment="Transactions"
			DOCS_CommentDelete="Delete"
Else
			'Подстановка TEXTANSI.TXT констант
			DOCS_DateActivation="Дата открытия"
			DOCS_DocID="Номер счета"
			DOCS_Name="Наименование счета"
			DOCS_AmountDoc="Текущий баланс"
			DOCS_Details1="Параметры счета"
			DOCS_StatusPayment="Транзакции"
			DOCS_CommentDelete="Удалить"
End If
End If

If InStr(UCase(CurrentClassDoc), UCase("Пропуска"))>0 Or InStr(UCase(Request("ClassDoc")), UCase("Пропуска"))>0 Then 
	DOCS_DateCompletion="Дата посещения"
	DOCS_Description="Цель посещения"
	DOCS_MakeCompleted="Назначить статус «Исполнено» - пропуск передан на проходную"
	DOCS_Details1="Реквизиты пропуска"
End If
If CurrentClassDoc="Пропуска / Разовые пропуска для иностранцев" Or Request("ClassDoc")="Пропуска / Разовые пропуска для иностранцев" Then 
	DOCS_Description="Название проводимого мероприятия"
End If

If InStr(UCase(CurrentClassDoc), UCase("Пропуска"))>0 Or InStr(UCase(Request("ClassDoc")), UCase("Пропуска"))>0 Then 
	DOCS_DocID="Номер пропуска"
	DOCS_DateCompletion="Дата посещения"
	DOCS_Description="Цель посещения"
	DOCS_MakeCompleted="Назначить статус «Исполнено» - пропуск передан на проходную"
	DOCS_Details1="Реквизиты пропуска"
	VAR_ChangeDocGetNewButton=""
	VAR_ChangeDocGenerateButton=""
	VAR_ChangeDocGetNewFromRegLogsButton=""
	VAR_ButtonsToShow	="ClickChangeDoc, ClickCreateComment, ClickCreateDepartment, ClickCreateDirectoryValues, ClickCreateDoc, ClickCreatePartner, ClickCreatePosition, ClickCreateRequest, ClickCreateType, ClickCreateUser, ClickDeleteDepartment, ClickDeleteDirectoryValues, ClickDeleteDoc, ClickDeleteMessage, ClickDeletePartner, ClickDeletePosition, ClickDeleteReporttype, ClickDeleteRequest, ClickDeleteType, ClickDeleteUser, ClickDeleteUserDirectory, ClickDeleteUserDirectoryValues, ClickDownload, ClickGetReportRefresh, ClickListDepartment, ClickListDirectoryValues, ClickListDoc, ClickListDocPrintable, ClickListDocRefresh, ClickListPartner, ClickListPositions, ClickListReporttype, ClickListRequest, ClickListType, ClickListUser, ClickListUserDirectories, ClickListUserDirectoryValues, ClickMakeActive, ClickMakeCanceled, ClickMakeArchival,  ClickMakeCompleted, ClickMakeInactive, ClickMakeOperative, ClickMSOffice, ClickMSOfficeStandard, ClickSetDeputy, "
End If
If CurrentClassDoc="Пропуска / Разовые пропуска для иностранцев" Or Request("ClassDoc")="Пропуска / Разовые пропуска для иностранцев" Then 
	DOCS_Description="Название проводимого мероприятия"
End If

Var_PossibleApplicationTypes=DOCS_BUSINESSPROCESSES+","+DOCS_Chancery+","+DOCS_Controller+","+DOCS_Viewing
If RUS()="RUS" Then 
	Var_PossibleApplicationTypes=Var_PossibleApplicationTypes+",Пропуска,"+DOCS_AppType_HelpDesk+","+DOCS_AppType_SAPR3
	'Var_PossibleApplicationTypes=""
	'Var_ApplicationType="" 'Default - Use this varyable for user view customization
	'Var_ApplicationType="Пропуска" 'Use this varyable for user view customization
	'Var_ApplicationType=DOCS_AppType_SAPR3 'Use this varyable for user view customization
	'Var_ApplicationType=DOCS_AppType_HelpDesk
Else
	Var_PossibleApplicationTypes=Var_PossibleApplicationTypes+","+DOCS_AppType_HelpDesk+","+DOCS_AppType_SAPR3
	Var_ApplicationType="" 'Use this varyable for user view customization
	'Var_ApplicationType=DOCS_AppType_SAPR3 'Use this varyable for user view customization
End If

'Out Session("UserID")
If Application("LicenseType")="PERSONAL" And Trim(Session("UserID"))<>"" Then
	Var_ApplicationType=DOCS_Chancery
	VAR_UseESignature=""
End If
'Out Var_ApplicationType

'EMailFieldList="-" 'No any document details are in e-mail notifications
'EMailFieldList="#DocID#DateActivation#DateCompletion#LISTVIEWEDDOCS#" 'Only DocID, DateActivation, DateCompletion document details are in e-mail notifications and Documents To View in PayDox E-Mail Client
'EMailFieldList="#DateActivation#DateCompletion#LISTVIEWEDDOCS#DocID#" 'Only DocID, DateActivation, DateCompletion document details are in e-mail notifications and Documents To View in PayDox E-Mail Client

If Var_ApplicationType=DOCS_Chancery Then
	VAR_UseIncomingOutgoingInTheLeftMenu="Y"
	VAR_ButtonsNotToShow="ClickCreateCommentResource, ClickCreateCommentBPStep, ClickShowReports, ClickDownloadXML"
End If

If IsHelpDeskSAP() Then
	If RUS()="RUS" Then 
		'DOCS_NameCreation="Инициатор документа"
		'If IsHelpDeskDoc() Or UCASE(Request("Type"))=UCASE("HelpDesk") Then
		If IsHelpDeskDoc() Or UCASE(Request("Type"))=UCASE("HelpDesk") Or Trim("NameRequest")<>"" Then
			DOCS_Name="Тема заявки"
			DOCS_DocID="Номер заявки"
			DOCS_PartnerName="Предприятие"
			DOCS_Correspondent="Передано в обработку - группа"
			DOCS_DateCompletion="Плановый срок исполнения"
			DOCS_Resolution="Меры по решению"
			DOCS_Resolutions="Меры по решению"
			DOCS_ResolutionAproval="Меры по решению"
			DOCS_ChangeResolution="Меры по решению"
		End If
		DOCS_NameResponsible="Исполнитель"
		But_MakeResponsible="Исполнитель"
		BUT_RESPONSIBLE="ИСПОЛНИТЕЛЬ"
		But_Resolution="Меры по решению"
		DOCS_ResolutionDocs="Требующие указания мер по решению заявки"
	End If
	If InStr(UCase(Request.ServerVariables("URL")),UCase("/ListDirectories.asp"))>0 Then
		VAR_TreeFolderSeparator="|"
	End If
End If

'Var_IsUseCheckInOut="y"

'Var_ApprovalPermitted=False 'Document aproval is NOT permitted while reconciliation process is not finished
'Var_ApprovalPermitted=True 'Document aproval is permitted even reconciliation process is not finished yet
'Out Var_ApprovalPermitted

'Var_ApprovalIfAllAgree=True 'Document aproval is NOT permitted if some reconciliation list user refused reconciliation (not agree)
'Var_ApprovalIfAllAgree=False 'Document aproval is permitted even if any reconciliation list user refused reconciliation (not agree)

'Var_ReconciliationIfAllAgree=True 'Next document reconciliation step is NOT permitted if some previous reconciliation list user refused reconciliation (not agree)

'Var_ApprovalIfAllAgree=True
'Out Var_ApprovalIfAllAgree
'Var_ReconciliationIfAllAgree=False
'//vnik
VAR_DocCreatorCanUpdateDocWithoutChangingStatuses="Y"
'//vnik

'If CurrentClassDoc="Договора" Then
'	DOCS_DocID="DocID 1"
'End If
'If CurrentClassDoc="Платежи" Then
'	DOCS_DocID="DocID 2"
'End If

'Out "VAR_ButtonsToShow:"+VAR_ButtonsToShow
'Out "VAR_ButtonsNotToShow:"+VAR_ButtonsNotToShow
'VAR_ParentConnectedToSeeDependentConnectedUserList="ListToView, NameCreation"

'VAR_FTPUploadsFullDirectoryPath=Application("PayDoxHomeDir")+"Uploads\"+Session("UserID")
'VAR_UseFTPUploads="Y"
'VAR_ReadAccess ="Y"
'Out "CurrentClassDoc:"+Session("CurrentClassDoc")
'If Session("CurrentClassDoc")="Договора" Then
'	DOCS_DocID="Номер д-та"
'End If

'If InStr(UCase(Request.ServerVariables("URL")),UCase("/ListDirectories.asp"))>0 Then 
'	If CurrentClassDoc="Договора" Then
'		VAR_ClassDocToShow="Задачи,Платежи"
'	End If
'End If

If InStr(UCase(Request.ServerVariables("SERVER_NAME")), ".COM")>0 Then
	bUseLang3=True
End If


' ------------------------------------------------ Настройки для Sitronics/STS
'Ph - 20090302 - Обнуляется Session("CurrentClassDoc")
If Trim(Session("CurrentClassDoc")) = "" Then
  Session("CurrentClassDoc") = Request("ClassDoc")
  If Trim(Session("CurrentClassDoc")) = "" Then
    Session("CurrentClassDoc") = Request("CurrentClassDoc")
  End If
End If

Var_ReconciliationIfAllAgree=True

'VAR_DocFieldsNotToShow="Description, Доп.поле 1" 'Not to show some doc record fields 
If bUseLang3 Then
' AM 12082008	DOCS_Notices="Задачи*Tasks*Feladatok"
' SAY 2008-08-26 поменял название на чешском, не соответствовало карточке
'  DOCS_Notices="Поручения*Tasks*objednávek"
  DOCS_Notices="Поручения*Tasks*Úkoly"
  
End If
VAR_UserMessageToEMail=Application("CustomerName")
'VAR_UploadFileProhibitedCommentTypes="FILE, HISTORY, REVIEW, VISA, RESOLUTION"
VAR_UploadFileProhibitedCommentTypes="VISA"
'VAR_DocFieldsNotToShow="Text 1, Text 2" 'Not to show some DOCS fields

' AM 080708
VAR_HomePageURL="Home.asp?l="+Request("l") 'Documents

sLang3="Česky"
sLanguage3="Česky"
sFlagImageLang3="czver.gif"
'Ph - 20081013 - Венгрию передвигаем на 4
sLang4="Hungarian"
sLanguage4="Hungarian"
sFlagImageLang4="huver.gif"

' SAY 2008-07-21

SIT_STS = "СТР"
SIT_STS_RU = "СТР"
SIT_SITRONICS = "СИТРОНИКС"
SIT_RTI = "РТИ"
SIT_MINC = "MINC"
SIT_VTSS = "ВТСС"
SIT_STS_ROOT_DEPARTMENT = "СТР*STS*/"
'///////////// AMW - Mikron
SIT_MIKRON = "МИКРОН"
SIT_MIKRON_ROOT_DEPARTMENT = "МИКРОН*MIKRON*/"

MIKRON_BSAP = "Закупки МИКРОН/БСАП"
MIKRON_RL_MEMO = "Закупки МИКРОН/Справка о ценах"
MIKRON_PROTOCOL = "Закупки МИКРОН/Протокол ЗК"
MIKRON_RL_PROTOCOL = "Закупки МИКРОН/Опросный лист к ПЗК"
MIKRON_PAYMENT_ORDER = "Закупки МИКРОН/Заявка на оплату"
MIKRON_PURCHASE_ORDER = "Закупки МИКРОН/Заявка на закупку"

MIKRON_CONTRACT = "Договоры МИКРОН/Договор на закупку"
MIKRON_S_CONTRACT = "Договоры МИКРОН/Договор на продажу"
MIKRON_ADD_CONTRACT = "Договоры МИКРОН/Доп. соглашение"
MIKRON_OLD_CONTRACT = "Договоры МИКРОН/Договор до 01.10.2013"
MIKRON_NDA_CONTRACT = "Договоры МИКРОН/О неразглашении (NDA)"
MIKRON_EXPADD_CONTRACT = "Договоры МИКРОН/Экспорт/Дополнение к контракту"
MIKRON_EXPORT_CONTRACT = "Договоры МИКРОН/Экспорт/Контракт"
'MIKRON_DELIVERY_CONTRACT = "Поступления МИКРОН/Доходы"
'MIKRON_STATE_CONTRACT = "Поступления МИКРОН/Гос. контракт"
MIK_EA_1 = "Разовая отгрузка"
MIK_EA_2 = "Номенклатура"
MIK_EA_3 = "Спецификация"
MIK_EA_4 = "Изменение условий"

MS_CONTRACT_INTERNAL = "Договор внутри группы Микроэлектроника"  'Это доп. поле AddField3
MS_CONTRACT_DEPEND = "Взаимозависимость сторон договора"
MS_CONTRACT_PARTY = "Бизнес-единица cторона договора"

MIKRON_CATALOG_ROLES           = "{3482B53C-018A-457A-BA91-BCDBD1EB106A}" 'справочник "Роли Микрон"
MIKRON_CATALOG_PAYMENT_TYPES   = "{857F1F9E-25B8-4AA1-A1FA-EB184EBCAA5E}" 'справочник "Закупки Микрон/Порядок оплаты"
MIKRON_CATALOG_ADD_AGREE_TYPES = "{D5FB4CF0-B8EE-4289-AD09-90800BB0CDC8}" 'справочник "Виды доп.соглашений"
MIKRON_CATALOG_CONTRACT_TYPES  = "{535C7740-CB1E-4403-ABC6-93AFC67205D5}" 'справочник "Виды договоров"
MIKRON_CATALOG_ORDERS_TYPES    = "{6E05F703-F06E-4895-A788-D991A02FEC00}" 'справочник "Виды распорядительных документов МИКРОН"
MIKRON_CATALOG_PURCHASE_TYPES  = "{B1DE9C2C-0B37-420E-9884-0C09D8CFF2DD}" 'справочник "Закупки Микрон/Способы закупки"
MIKRON_CATALOG_FINANCIAL_COSTS = "{3C047498-29BE-4F60-8940-40CFC6ED702F}" 'cправочник "Закупки МИКРОН/ЦФЗ"
MIKRON_CATALOG_RESPONCIBILITY  = "{7302B7EF-EB19-40E4-BAD8-2163207DB147}" 'cправочник "Закупки МИКРОН/ЦФO"
MIKRON_CATALOG_PROJECT_LIST    = "{c8c9cb66-f982-4266-9918-7135e14315a3}" 'cправочник "Закупки МИКРОН/Проекты МИКРОН"
MIKRON_CATALOG_SRC_FINANCING   = "{385AFB8A-28A7-4023-A87C-DC00EB19C75B}" 'справочник "Источник финансирования"
MIKRON_CATALOG_EXPENDITURE     = "{2CB6E1E0-9820-4971-93A6-AC4BD9624E5B}" 'справочник статей затрат "Бюджетный классификатор МИКРОН"
MIKRON_CATALOG_BUYERS          = "Постоянный потребитель,Пробная закупка,Дистрибутор,Дилер,Официальный представитель,PR акция"
CATALOG_CONTRACTING_PARTY      = "{63e943e3-e678-4138-900b-1e5fe96809ae}" 'справочник "Contracts/Сторона Договора"

MIKRON_REPORT_CONTRACTS_MONTH  = "{944C0408-0248-47B1-9634-B3FFF0F4E0B7}" 'отчет по договорам за период
'///////////// AMW - Mikron
'vnik_scsp
SIT_SCSP = "СПСТ"
'vnik_scsp
'Запрос №1 - СИБ - start
SIT_SIB_ROOT_DEPARTMENT = "СИБ*SIB*/"
SIT_SIB = "СИБ"
'Запрос №1 - СИБ - end

SIT_SITRU = "СИТРОНИКС ИТ" ' DmGorsky_3
SIT_SITRU_ROOT = "СИТРОНИКС ИТ*SITRONICS IT*/" ' DmGorsky_3

SIT_VHODYASCHIE = "Входящие документы"
SIT_ISHODYASCHIE = "Исходящие документы"
MINC_ISHODYASCHIE = "Исходящие документы Минц"
SIT_SLUZH_ZAPISKA = "Служебные записки"
SIT_RASP_DOCS = "Распорядительные документы"
SIT_NORM_DOCS = "Нормативные документы"
' AM 120808 SIT_ZADACHI = "Задачи"
SIT_ZADACHI = "Поручения"
'Запрос №11 - СТС - start
'SIT_DOGOVORI = "Договоры"
SIT_DOGOVORI_OLD = "Договоры до даты 21.07.2010"
SIT_DOGOVORI_NEW = "Договоры*"
'Запрос №11 - СТС - end
'20090622 - Заявка ТКП
SIT_COM_OFFERS = "Коммерческие предложения"
'vnik_protocols
SIT_PROTOCOLS = "Протоколы"
SIT_PROTOCOLS_MC_EGRB = "УК ЭПРБ"
SIT_PROTOCOLS_IT_Committee = "Комитет по ИТ" 
SIT_PROTOCOLS_Management_Board = "Правление"
SIT_PROTOCOLS_Control_And_Auditing_Committee = "Контрольно-ревизионный комитет" 
'vnik_protocols

'vnik_payment_order
SIT_PAYMENT_ORDER = "Закупки*Purchase*Nákup/Заявка на оплату УК*Payment order MC*Zadost o platbu SS"
'vnik_payment_order

'vnik_purchase_order
SIT_PURCHASE_ORDER = "Закупки*Purchase*Nákup/Заявка на закупку УК*Purchase order MC*Nákupní objednávka"
'vnik_purchase_order

'rti_purchase_order
RTI_PURCHASE_ORDER = "Закупки РТИ/Заявка на закупку"
'rti_purchase_order

'rti_payment_order
RTI_PAYMENT_ORDER = "Закупки РТИ/Заявка на оплату"
'rti_payment_order

'rti_contract
RTI_CONTRACT = "Закупки РТИ/Договор"
'rti_contract

'rti_bsap
RTI_BSAP = "Закупки РТИ/БСАП"
'rti_bsap

'rti_protocol
RTI_PROTOCOL = "Закупки РТИ/Протокол ЦЗК"
'rti_protocol

'vnik_protocolsCPC
SIT_PROTOCOLS_CPC = "Закупки*Purchase*Nákup/Протокол ЦЗК*Protocol CPC*Protocol ÚKZZ"
'vnik_protocolsCPC

'vnik_contracts
SIT_CONTRACTS_MC = "Закупки*Purchase*Nákup/Договоры УК*Contracts MC*Smlouvy SS"
'vnik_contracts

'SAY 2008-10-27
SIT_VHODYASCHIE_ACC="Входящие для бухгалтерии*Incoming correspondence for Accounting*Příchozí účetní dokumenty"

'Phil 20080817
STS_PurchaseOrder = "Заявка на закупку"
STS_PaymentOrder = "Заявка на оплату*"
'ph - 20080918 - Текст заявки на мобильный
SIT_MOBILE_CONTENT = "В связи с производственной необходимостью и на основании Положения об обеспечении корпоративными сервисами сотрудников ОАО ""СИТРОНИКС"" (приказ № 114 от 17.12.2007 г.) прошу предоставить нижеуказанному(ым) сотруднику(ам) компании право пользования служебной мобильной телефонной связью в пределах установленного лимита:"+ _
  VbCrLf+VbCrLf+"- ФИО - должность"+VbCrLf+"- ФИО - должность"

'SAY 2008-10-30
STS_PRIKAZ_TEXT = "В связи с / в целях ..."+VbCrLf+VbCrLf+VbCrLf+"П Р И К А З Ы В А Ю:"+VbCrLf+VbCrLf+"1. ..."+VbCrLf+VbCrLf+"2. ..."+VbCrLf+VbCrLf+"3. Контроль над исполнением настоящего приказа возложить на .../ либо"+VbCrLf+VbCrLf+"Контроль над исполнением настоящего приказа оставляю за собой (тогда с красной строки без нумерации пункта)"

'SAY 2008-08-25
SIT_SLUZH_ZAPISKA_COMPUTER = "Служебные записки*Office memo*Interní sdělení/На выделение компьютера*Provide computer*Provide computer" 
SIT_SLUZH_ZAPISKA_MOBILE = "Служебные записки*Office memo*Interní sdělení/На выделение мобильного*Provide cellphone*Provide cellphone" 
SIT_SLUZH_ZAPISKA_KOMANDIROVKA = "Служебные записки*Office memo*Interní sdělení/На командировку*Assignment*Assignment" 
SIT_SLUZH_ZAPISKA_OBUCHENIE = "Служебные записки*Office memo*Interní sdělení/На обучение*Training*Training" 
SIT_SLUZH_ZAPISKA_PERSONAL = "Служебные записки*Office memo*Interní sdělení/На подбор персонала*Staff recruitment*Staff recruitment" 
SIT_SLUZH_ZAPISKA_OBSCHAY = "Служебные записки*Office memo*Interní sdělení/Общая форма*Universal form*Universal form" 

'rmanyushin@sitronics.com 15.07.2009, Start
'Запрос №34 - СТС - start
'STS_SLUZH_ZAPISKA_OVERTIME = "Служебные записки*Office memo*Interní sdělení/На переработки*Overtime*Overtime"
STS_SLUZH_ZAPISKA_OVERTIME = "DELETED - Служебные записки*Office memo*Interní sdělení/На переработки*Overtime*Overtime"
'Запрос №34 - СТС - end
'rmanyushin@sitronics.com 15.07.2009, Stop

'rmanyushin 119579 19.08.2010 Start
STS_SLUZH_ZAPISKA_HOLIDAY = "Служебные записки*Office memo*Interní sdělení/На отпуск*Holiday Request*Žádost o dovolenou"
'rmanyushin 119579 19.08.2010 End

'rmanyushin 136964 08.11.2010 Start
'Запрос №34 - СТС - start
'STS_SLUZH_ZAPISKA_OVERTIME2 = "Служебные записки*Office memo*Interní sdělení/На переработки RU*Overtime RU*Overtime RU"
STS_SLUZH_ZAPISKA_OVERTIME2 = "Служебные записки*Office memo*Interní sdělení/На переработки*Overtime*Overtime"
'Запрос №34 - СТС - end
'rmanyushin 136964 08.11.2010 End
'Запрос №46 - СТС - start
STS_SLUZH_ZAPISKA_OVERTIME_PLAN = "Служебные записки*Office memo*Interní sdělení/На переработки (план)*Overtime (plan)*Overtime (plan)"
STS_SLUZH_ZAPISKA_OVERTIME_FACT = "Служебные записки*Office memo*Interní sdělení/На переработки (факт)*Overtime (actual)*Overtime (actual)"
'Запрос №46 - СТС - end

'SAY 2008-08-22
VAR_UseESignature=""
VAR_CanCreateMainVersionAddFiles=False

' прячем пользовательские спрвочники (SAY 2008-08-22)
'VAR_ShowUserDirectory="Y"
VAR_NotToShowUserDirectories = ""
If Not IsAdmin() Then
'  VAR_ShowUserDirectory=""
  VAR_NotToShowUserDirectories = "Y"
End If

'скрываем от Ситроникса чужие категории
'If InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) = 1 Then
'  VAR_ClassDocToShow = SIT_VHODYASCHIE+","+SIT_ISHODYASCHIE+","+SIT_SLUZH_ZAPISKA+","+SIT_RASP_DOCS+","+SIT_NORM_DOCS+","+SIT_ZADACHI+","+SIT_DOGOVORI+","
'End If

'скрываем от РТИ чужие категории
'If InStr(UCase(Session("Department")), UCase(SIT_RTI)) = 1 Then
'  VAR_ClassDocToShow = SIT_VHODYASCHIE+","+SIT_ISHODYASCHIE'+","+SIT_SLUZH_ZAPISKA+","+SIT_RASP_DOCS+","+SIT_NORM_DOCS+","+SIT_ZADACHI+","+SIT_DOGOVORI+","
'End If

'phil - 20080913 - Start - Переименование текстовых констант сделал с учетом языка
'If UCase(Request("l")) = "RU" Then

'Меняем надпись на кнопке "Документ MS Excel"
If InStr(UCase(Session("Department")), UCase(SIT_RTI)) = 1 and UCase(Request("l")) = "RU" Then
But_MSOfficeStandard="Шаблон"
end if

If InStr(UCase(Session("Department")), UCase(SIT_VTSS)) = 1 and UCase(Request("l")) = "RU" Then
But_MSOfficeStandard="Шаблон"
end if

'Ph - 20080922
SIT_SECURITYLEVEL_ALL = "Не ограничен списком рассылки"
SIT_SECURITYLEVEL_LISTONLY = "Ограничен списком рассылки"

'SAY 2008-09-22 Общие переводы --- start
Select Case UCase(Request("l"))
   Case "RU" 'RU
      DOCS_TaskInactive="Поручение неактивно"
      DOCS_DocIDAdd = "Проектный номер"
      'Переименовываем Кнопку "Добавить информацию"
      DOCS_CreateDocInfo = "Сохранить"
'phil - 20080906 - Start - Сообщения об ошибках при проверке срока автоматического согласования
      SIT_DateFormatError = "Ошибка в написании даты: "
      SIT_TooEarlyAutoReconcile = "Слишком малый срок для автоматического согласования: "
	  SIT_TooEarlyDate = "Слишком ранняя дата исполнения: "
'Ph - 20081211 - start
'    BUT_ALLDOCS = "ВСЕ ДОКУМЕНТЫ"
'    BUT_ALLDOCSDESCRIPTION = "Все документы к которым у меня есть доступ"
    DOCS_ALLToShowTitle = "Все документы к которым у меня есть доступ"
'Ph - 20081211 - end
    ' SAY 2008-08-14
    But_Task = "Поручение"
    DOCS_MakeNotice = "Создать поручение"
    'SAY 2008-11-10
    BUT_ALLNew="В РАБОТУ"
    DOCS_ALL1="Документы которые требуют моей обработки: согласования, утверждения, исполнения, ознакомления, регистрации или контроля"
    BUT_COMPLETION1="ПОРУЧЕНИЯ МОИ"
    DOCS_NOTCOMPLETED="Неисполненные поручения, где я поручитель или контролер"
    BUT_RESPONSIBLE="ПОРУЧЕНИЯ МНЕ"
    DOCS_YouAreResponsible="Поручения, где я исполнитель"

    DOCS_EXPIRED="Документы с наступающим или истекшим сроком"
    DOCS_YouAreCreator="Документы, которые я создал(а)"
    DOCS_UNAPPROVED="Документы, требующие моего согласования"
    DOCS_UNAPPROVED1="Документы, требующие моего утверждения"
    DOCS_ViewedStatusDocs="Документы, требующие моего ознакомления"
    

'	BUT_VISA="СОГЛАСОВАНИЕ"
'	BUT_APROVAL1="УТВЕРЖДЕНИЕ"
'	BUT_VIEWEDSTATUSDOCS="ОЗНАКОМИТЬСЯ"
'	DOCS_UNAPPROVED="Документы, требующие Вашего согласования"
'	DOCS_UNAPPROVED1="Документы, требующие Вашего утверждения"
'	DOCS_UNAPPROVED2="Документы, требующие утверждения"
'	DOCS_UNAPPROVED3="Документы, требующие согласования"
'	DOCS_YouAreCreator="Неисполненные документы, которые Вы создали"

    If InStr(UCase(Request.ServerVariables("URL")),UCase("/showdoc.asp"))>0 Then
      DOCS_Home2="Главная страница"
      But_List="Назад к списку"

      But_Create="Создать копию"
      DOCS_CreateDocRecord="Создать карточку на основе данных текущей"
  
    End If

	'Центральные кнопки
    SIT_CentralBut_CoResponsible = "СОИСПОЛНИТЕЛЬ"
    SIT_CentralBut_PurchaseOrders = "ЗАЯВКИ"

'ph - 20100416 - start - переименование поля Обязательный идентификатор в контрагентах
    if InStr(UCase(Session("Department")), UCase(SIT_MINC)) = 1 Then
         DOCS_IDRequired = "Уникальный идентификатор"
    Else 
	     DOCS_IDRequired = "Буквенный код контрагента"
    end if
    S_PostIndex_name = "Почтовый индекс"
'ph - 20100416 - end

   Case ""   'EN
      DOCS_TaskInactive="Task is inactive"
      DOCS_DocIDAdd = "Draft number"
'phil - 20080906 - Start - Сообщения об ошибках при проверке срока автоматического согласования
      SIT_DateFormatError = "Wrong date: "
      SIT_TooEarlyAutoReconcile = "The time period is too short for automated approval: "
'Ph - 20081211 - start
'    BUT_ALLDOCS = "ALL DOCUMENTS"
'    BUT_ALLDOCSDESCRIPTION = "All documents with access"
    DOCS_ALLToShowTitle = "All documents with access"
'Ph - 20081211 - end
    ' SAY 2008-08-14
    But_Task = "Task"
    DOCS_MakeNotice = "Create new Task"
    DOCS_ListToReconcile = "Reviewer"
	DOCS_NameAproval = "Approver"

	'Центральные кнопки
    SIT_CentralBut_CoResponsible = "EXECUTOR"
    SIT_CentralBut_PurchaseOrders = "ORDERS"

   Case "3"  'CZ
      DOCS_TaskInactive="Úkol je neaktivní"
      DOCS_DocIDAdd = "Číslo návrhu"
'phil - 20080906 - Start - Сообщения об ошибках при проверке срока автоматического согласования
      SIT_DateFormatError = "Wrong date: "
      SIT_TooEarlyAutoReconcile = "Časový úsek je příliš krátký pro automatické schválení: "
'Ph - 20081211 - start
'    BUT_ALLDOCS = "VŠECHNY DOKUMENTY"
'    BUT_ALLDOCSDESCRIPTION = "všechny dokumenty s přístupem"
    DOCS_ALLToShowTitle = "všechny dokumenty s přístupem"
'Ph - 20081211 - end
    ' SAY 2008-08-14
    But_Task = "úkol"
    DOCS_MakeNotice = "Vytvořit úkol"

    BUT_COMPLETION1="MNOU ZADANÉ ÚKOLY"
    BUT_RESPONSIBLE="MOJE ÚKOLY"

	'Центральные кнопки
    SIT_CentralBut_CoResponsible = "ZPRACOVATEL"
    SIT_CentralBut_PurchaseOrders = "ŽÁDOSTI"
End Select
'SAY 2008-09-22 Общие переводы --- end


  If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_VHODYASCHIE)) > 0 Then 
    Select Case UCase(Request("l"))
        Case "RU" 'RU
	DOCS_DocID = "Входящий номер"
	DOCS_DocIDParent = "Документ-основание"
	DOCS_Name = "Тема"
	DOCS_PartnerName = "Организация"
        DOCS_ListToView = "Получатели"
	DOCS_DocIDIncoming="Исходящий номер"
        Case "" 'EN
	DOCS_DocID = "Incoming number"
	DOCS_DocIDParent = "Basis document"
	DOCS_Name = "Subject"
	DOCS_PartnerName = "Company"
        DOCS_ListToView = "Addressees"
	DOCS_DocIDIncoming="Outgoing number"
        Case "3" 'CZ
	DOCS_DocID = "Příchozí číslo"
	DOCS_DocIDParent = "Osnova dokumentu"
	DOCS_Name = "Předmět"
	DOCS_PartnerName = "Společnost"
        DOCS_ListToView = "Adresy"
	DOCS_DocIDIncoming="Odchozí číslo"
    End Select
       
  End If
  
  'РТИ
      If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_VHODYASCHIE)) > 0 and InStr(UCase(Session("Department")), UCase(SIT_RTI)) = 1 Then
        DOCS_Name = "Заголовок к тексту (краткое содержание)"
        S_NameUserFieldText3 = "Подписант (внешний)"
        S_NameUserFieldText8 = "Тип входящего РТИ"
    end if
      'РТИ
'oaorti
      If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_VHODYASCHIE)) > 0 and InStr(UCase(Session("Department")), UCase(SIT_MINC)) = 1 Then
        DOCS_PartnerName = "Корреспондент"
        DOCS_Content = "Краткое содержание"
        S_NameUserFieldText7 = "Индекс дела"
        S_NameUserFieldDate2 = "Срок исполнения"
        S_NameUserFieldText8 = "Примечание"
      end if
      If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) > 0 and InStr(UCase(Session("Department")), UCase(SIT_MINC)) = 1 Then
        DOCS_PartnerName = "Получатели (Организация)"
        DOCS_Content = "Содержание"
        DOCS_Name = "Тема"
        DOCS_DocID = "Индекс документа"
        DOCS_Author = "Автор"
        DOCS_DateActivation = "Дата регистрации"
      end if
      
      

'oaorti      
      
      
'SAY 2008-10-27
  If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_VHODYASCHIE_ACC)) > 0 Then 
    Select Case UCase(Request("l"))
        Case "RU" 'RU
	DOCS_DocID = "Входящий номер"
	DOCS_Name = "Тема"
	DOCS_PartnerName = "Организация"
        DOCS_ListToView = "Получатели"
	DOCS_DocIDIncoming="Исходящий номер"
        Case ""   'EN
	DOCS_DocID = "Incoming number"
	DOCS_Name = "Subject"
	DOCS_PartnerName = "Company"
        DOCS_ListToView = "Addressees"
	DOCS_DocIDIncoming="Outgoing number"
        Case "3"  'CZ
	DOCS_DocID = "Příchozí číslo"
	DOCS_Name = "Předmět"
	DOCS_PartnerName = "Společnost"
        DOCS_ListToView = "Adresy"
	DOCS_DocIDIncoming="Odchozí číslo"
    End Select
  End If


  If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) > 0 Then 
    Select Case UCase(Request("l"))
        Case "RU" 'RU
	DOCS_DocID = "Исходящий номер"
	DOCS_DocIDParent = "Документ-основание"
	DOCS_Name = "Тема"
	DOCS_Author = "Инициатор"
	DOCS_NameAproval = "Подписант"
	DOCS_LocationPath = "Регистратор"
        DOCS_PartnerName="Получатель (Организация)"
	'SAY 2008-12-02
	DOCS_Correspondent="Получатели (ФИО)"
	VAR_AddUsersToCorrespondent=""
	S_NameUserFieldText8 = "Тип исходящего РТИ"

        Case ""   'EN
	DOCS_DocID = "Outgoing number"
	DOCS_DocIDParent = "Basis document"
	DOCS_Name = "Subject"
'	DOCS_Author = "Initiator"
	DOCS_Author = "Drafter"
'	DOCS_NameAproval = "Signee"
	DOCS_LocationPath = "Registrar"
	'SAY 2008-12-02
	DOCS_Correspondent="Addressees"
	VAR_AddUsersToCorrespondent=""
        Case "3"  'CZ
	DOCS_DocID = "Odchozí číslo"
	DOCS_DocIDParent = "Osnova dokumentu"
	DOCS_Name = "Předmět"
	DOCS_Author = "Iniciátor"
	DOCS_NameAproval = "Podepisovatel"
	DOCS_LocationPath = "Registrátor"
	'SAY 2008-12-02
	DOCS_Correspondent="Adresy"
	VAR_AddUsersToCorrespondent=""
    End Select


  End If
  'РТИ
  
      If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) > 0 and InStr(UCase(Session("Department")), UCase(SIT_RTI)) = 1 Then 
        DOCS_Name = "Заголовок к тексту (краткое содержание)"
    end if 
'РТИ
  If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_SLUZH_ZAPISKA)) > 0 Then 
    Select Case UCase(Request("l"))
        Case "RU" 'RU
	DOCS_DocID = "Номер документа"
	DOCS_DocIDParent = "Документ-основание"
	DOCS_Name = "Заголовок"
	DOCS_Author = "Инициатор"
        'SAY 2008-10-16 убираем ознакомление
        DOCS_ListToView = "Получатели"
        'DOCS_Correspondent = "Получатели"
	DOCS_NameAproval = "Подписант"
	DOCS_LocationPath = "Регистратор"
        Case ""   'EN
	DOCS_DocID = "Document number"
	DOCS_DocIDParent = "Basis document"
	DOCS_Name = "Headline"
'	DOCS_Author = "Initiator"
	DOCS_Author = "Drafter"
        DOCS_ListToView = "Addressees"
'	DOCS_NameAproval = "Signee"
	DOCS_LocationPath = "Registrar"
        Case "3"  'CZ
	DOCS_DocID = "Číslo dokumentu"
	DOCS_DocIDParent = "Osnova dokumentu"
	DOCS_Name = "Titulek"
	DOCS_Author = "Iniciátor"
        DOCS_ListToView = "Adresy"
	DOCS_NameAproval = "Podepisovatel"
	DOCS_LocationPath = "Registrátor"
    End Select
  End If
'РТИ
  If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_SLUZH_ZAPISKA)) = 1 and InStr(UCase(Session("Department")), UCase(SIT_RTI)) = 1 Then
     DOCS_Name = "Заголовок к тексту (краткое содержание)"
  end if
'РТИ

  If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_RASP_DOCS)) > 0 Then 
    Select Case UCase(Request("l"))
        Case "RU" 'RU
	DOCS_DocID = "Номер документа"
	DOCS_Name = "Название документа"
	DOCS_Author = "Разработчик"
        DOCS_ListToView = "Получатели"
	DOCS_NameAproval = "Подписант"
	DOCS_LocationPath = "Регистратор"
        'phil - 20080918
	DOCS_DocIDParent = "Номер нормативного документа"
        Case "" 'EN
	DOCS_DocID = "Document number"
	DOCS_Name = "Document name"
	DOCS_Author = "Developer"
        DOCS_ListToView = "Addressees"
'	DOCS_NameAproval = "Signee"
	DOCS_LocationPath = "Registrar"
        'phil - 20080918
	DOCS_DocIDParent = "Number of regulatory document"
        Case "3" 'CZ
	DOCS_DocID = "Číslo dokumentu"
	DOCS_Name = "Název dokumentu"
	DOCS_Author = "Vývojář"
        DOCS_ListToView = "Adresy"
	DOCS_NameAproval = "Podepisovatel"
	DOCS_LocationPath = "Registrátor"
        'phil - 20080918
	DOCS_DocIDParent = "Číslo řídícího dokumentu"
    End Select
  End If
  
       'Распорядительные РТИ   
    If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_RASP_DOCS)) > 0 and InStr(UCase(Session("Department")), UCase(SIT_RTI)) = 1 Then
        DOCS_Name = "Заголовок к тексту (краткое содержание)"
        DOCS_DocID = "Номер распорядительного документа"
        DOCS_Author = "Инициатор"
        DOCS_DocIDParent = "Номер родительского документа"
    end if

'vnik_purchase_order
If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_PURCHASE_ORDER)) > 0 Then
  Select Case UCase(Request("l"))
      Case "RU" 'RU
    DOCS_QuantityDoc = "Количество товара, работ, услуг"    
      Case "" 'EN
    
      Case "3" 'CZ
    
    End Select
 End If
'vnik_purchase_order

'vnik_protocols 
  If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_PROTOCOLS)) > 0 or InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_PROTOCOLS_CPC)) > 0 Then
  Select Case UCase(Request("l"))
      Case "RU" 'RU
    DOCS_DocID = "Номер документа"
    DOCS_Name = "Название документа"
    DOCS_Author = "Разработчик/Секретарь"
    DOCS_ListToView = "Получатели"
    DOCS_NameAproval = "Подписант/Председатель"
      Case ""   'EN
    DOCS_DocID = "Document number"
    DOCS_Name = "Document name"
    DOCS_Author = "Developer/Secretary"
    DOCS_ListToView = "Addressees"
    DOCS_NameAproval = "Approver/Chairman"
    Case "3"
        'CZ
    DOCS_DocID = "Číslo dokumentu"
    DOCS_Name = "Název dokumentu"
    DOCS_Author = "Developer/Registrátor"
    DOCS_ListToView = "Příjemci"
    DOCS_NameAproval = "Signatáři/Předseda"
    End Select
  End If
'vnik_protocols

'vnik_payment_order
If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_PAYMENT_ORDER)) > 0 Then
  Select Case UCase(Request("l"))
      Case "RU"
        'RU
    DOCS_DocID = "Номер документа"
    DOCS_Name = "Название документа"
    DOCS_ListToView = "Получатели"
    DOCS_NameAproval = "Утверждающий"
    DOCS_Description = "Основание платежа"
    DOCS_DocIDParent = "Документ основание"
    'ниже пока только на русском
    DOCS_Content = "Комментарий"
    DOCS_Author = "Инициатор"
    DOCS_Currency = "Валюта"
    
    Case ""
        'EN
    DOCS_DocID = "Document number"
    DOCS_Name = "Document name"
    DOCS_ListToView = "Addressees"
    DOCS_NameAproval = "Approver"
    DOCS_Description = "Payment justification"
    DOCS_DocIDParent = "Basis document"
    Case "3"
        'CZ
    DOCS_DocID = "Číslo dokumentu"
    DOCS_Name = "Název dokumentu"
    DOCS_ListToView = "Příjemci"
    DOCS_NameAproval = "Schvalovatel"
    DOCS_Description = "Důvod platby"
    DOCS_DocIDParent = "Osnova dokumentu"
    End Select
  End If
'vnik_payment_order

'vnik_contracts
If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_CONTRACTS_MC)) > 0 Then
  Select Case UCase(Request("l"))
      Case "RU"
        'RU
    DOCS_DocID = "Номер документа"
    DOCS_Name = "Вид обязательств"
    DOCS_ListToView = "Получатели"
    DOCS_NameAproval = "Подписант"
    DOCS_Description = "Примечания"
    DOCS_DocIDParent = "Документ основание"
    'ниже пока только на русском
    DOCS_Content = "Предмет договора"
    DOCS_Author = "Инициатор"
    DOCS_Currency = "Валюта"
    DOCS_NameResponsible = "Ответственный исполнитель"
    
    Case ""
        'EN
    
    Case "3"
        'CZ
    
    End Select
  End If
'vnik_contracts

  If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_NORM_DOCS)) > 0 Then 
    Select Case UCase(Request("l"))
      Case "RU"
        'RU
	DOCS_DocID = "Номер документа"
        'phil - 20080918 - отключено
        'DOCS_DocIDParent = "Номер приказа об утверждении"
	DOCS_Name = "Название документа"
	DOCS_Author = "Разработчик"
        DOCS_ListToView = "Получатели"
	DOCS_NameAproval = "Подписант"
	DOCS_LocationPath = "Регистратор"
      Case ""
        'EN
	DOCS_DocID = "Document number"
        'phil - 20080918 - отключено
        'DOCS_DocIDParent = "Номер приказа об утверждении"
	DOCS_Name = "Document name"
	DOCS_Author = "Developer"
        DOCS_ListToView = "Addressees"
'	DOCS_NameAproval = "Signee"
	DOCS_LocationPath = "Registrar"
      Case "3"
        'CZ
	DOCS_DocID = "Číslo dokumentu"
        'phil - 20080918 - отключено
        'DOCS_DocIDParent = "Номер приказа об утверждении"
	DOCS_Name = "Název dokumentu"
	DOCS_Author = "Vývojář"
        DOCS_ListToView = "Adresy"
	DOCS_NameAproval = "Podepisovatel"
	DOCS_LocationPath = "Registrátor"
    End Select

  End If

  If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ZADACHI)) > 0 Then 
    Select Case UCase(Request("l"))
      Case "RU"
        'RU
	DOCS_DocID = "Номер поручения"
	DOCS_DocIDParent = "Номер родительского документа"
	DOCS_Name = SIT_TaskName
	DOCS_Content = "Содержание"
	DOCS_Author = "Инициатор"
        DOCS_Rank = "Срочность"
        DOCS_DateActivation = "Дата выдачи"
        'SAY 2008-10-17 меняем поле соисполнителей (делаем Correspondent)
        'DOCS_ListToView = "Соисполнители"
        'DOCS_Correspondent = "Соисполнители"
	'DOCS_NoticesUserList = "Адресаты, список рассылки"
        DOCS_NoticesUserList = "Соисполнители"

        DOCS_DateActivationTask = "Дата выдачи"
'Ph - 20110929 - start
        DOCS_NameResponsible = "Исполнитель"
'Ph - 20110929 - end
      Case ""
        'EN
	DOCS_DocID = "Index of task"
	DOCS_Name = SIT_TaskName
'	DOCS_Author = "Initiator"
	DOCS_Author = "Drafter"
        DOCS_Rank = "Urgency"
        DOCS_DateActivation = "Date of issue"
        DOCS_ListToView = "Co-executors"
	DOCS_NoticesUserList = "Correspondents, Distribution list"
'Ph - 20110929 - start
        DOCS_NameResponsible = "Responsible person"
'Ph - 20110929 - end
      Case "3"
        'CZ
	DOCS_DocID = "Index úkolu"
	DOCS_Name = SIT_TaskName
	DOCS_Author = "Iniciátor"
        DOCS_Rank = "Urgence"
        DOCS_DateActivation = "Datum vydání"
        DOCS_ListToView = "Spolu-vykonavatel"
	DOCS_NoticesUserList = "Adresáti, seznam pro rozeslání"
'Ph - 20110929 - start
        DOCS_NameResponsible = "Odpovědná osoba"
'Ph - 20110929 - end
    End Select
  End If

'Запрос №11 - СТС - start
  If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_DOGOVORI_OLD)) = 1 Then 
'  If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_DOGOVORI)) > 0 Then 
'Запрос №11 - СТС - end
    Select Case UCase(Request("l"))
      Case "RU"
        'RU
        'AM 20082008	DOCS_DocID = "Проектный № документа"
        'DOCS_DocIDAdd = "Регистрационный № документа"
	DOCS_DocID = "№ проекта документа"
	DOCS_DocIDAdd = "Рег. № документа"
	DOCS_Name = "Вид обязательств"
	DOCS_Author = "Инициатор"
	DOCS_NameResponsible = "Ответственный исполнитель"
	DOCS_Description = "Предмет договора"
	DOCS_DocIDIncoming = "Проект"
	DOCS_QuantityDoc = "Количество листов"
	DOCS_InventoryUnit = "Примечание"
	DOCS_AmountDoc = "Сумма договора(включая НДС)"
	DOCS_Currency = "Валюта"
	DOCS_Content = "Наличие и описание штрафных санкций и премий"
      Case ""
        'EN
        'AM 20082008	DOCS_DocID = "Проектный № документа"
        'DOCS_DocIDAdd = "Регистрационный № документа"
	DOCS_DocID = "Document draft #"
	DOCS_DocIDAdd = "Document's registered#"
	DOCS_Name = "Type of liabilities"
'	DOCS_Author = "Initiator"
	DOCS_Author = "Drafter"
	DOCS_NameResponsible = "Responsible"
	DOCS_Description = "Subject of the contract"
	DOCS_DocIDIncoming = "Draft"
	DOCS_QuantityDoc = "Number of pages"
	DOCS_InventoryUnit = "Comments"
	DOCS_AmountDoc = "Price of contract (including VAT)"
	DOCS_Currency = "Currency"
	DOCS_Content = "Penalties and rewards (if any) and their description"
      Case "3"
        'CZ
        'AM 20082008	DOCS_DocID = "Проектный № документа"
        'DOCS_DocIDAdd = "Регистрационный № документа"
	DOCS_DocID = "Návrh dokumentu č."
	DOCS_DocIDAdd = "Registrovaný dokument č."
	DOCS_Name = "Druh závazků"
	DOCS_Author = "Iniciátor"
	DOCS_NameResponsible = "Zodpovědný"
	DOCS_Description = "Předmět smlouvy"
	DOCS_DocIDIncoming = "Návrh"
	DOCS_QuantityDoc = "Počet stránek"
	DOCS_InventoryUnit = "Poznámky"
	DOCS_AmountDoc = "Cena kontraktu (včetně DPH)"
	DOCS_Currency = "Měna"
	DOCS_Content = "Srážky a odměny (pokud jsou) a jejich popis"
    End Select
  End If

'Запрос №11 - СТС - start
  If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_DOGOVORI_NEW)) = 1 Then 
    Select Case UCase(Request("l"))
      Case "RU"
        DOCS_DocID = "Регистрационный номер договора/код договора"
        DOCS_Name = "Вид обязательств"
        DOCS_Author = "Ответственный исполнитель"
        DOCS_NameResponsible = "Инициатор"
        DOCS_Description = "Предмет договора"
        DOCS_DocIDIncoming = "Проект"
        DOCS_InventoryUnit = "Примечание"
        DOCS_AmountDoc = "Сумма договора(включая НДС)"
      Case ""
        'EN
        DOCS_DocID = "Contract registration No./Contract code"
        DOCS_Name = "Type of liabilities"
        
        'rmanyushin 136087 12.10.2010 End
			'DOCS_Author = "Responsible"
			'DOCS_NameResponsible = "Drafter"
			DOCS_Author = "Drafter"
			DOCS_NameResponsible = "Responsible"
        'rmanyushin 136087 12.10.2010 End
        
        DOCS_Description = "Subject of the contract"
        DOCS_DocIDIncoming = "Draft"
        DOCS_InventoryUnit = "Comments"
        DOCS_AmountDoc = "Price of contract (including VAT)"
      Case "3"
        'CZ
        DOCS_DocID = "Registrační číslo smlouvy/Kód smlouvy"
        DOCS_Name = "Druh závazků"
        
        'rmanyushin 136087 12.10.2010 End
			'DOCS_Author = "Zodpovědný"
			'DOCS_NameResponsible = "Iniciátor"
			DOCS_Author = "Iniciátor"
			DOCS_NameResponsible = "Zodpovědný"
        'rmanyushin 136087 12.10.2010 End
        
        DOCS_Description = "Předmět smlouvy"
        DOCS_DocIDIncoming = "Návrh"
        DOCS_InventoryUnit = "Poznámky"
        DOCS_AmountDoc = "Cena kontraktu (včetně DPH)"
    End Select
  End If
'Запрос №11 - СТС - end

'rmanyushin 136964 08.11.2010 Start
If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) = 1 Then 
    Select Case UCase(Request("l"))
      Case "RU"
        DOCS_NameResponsible = "Заказчик переработки"
      Case ""
        'EN
		DOCS_NameResponsible = "Overtime Requester"
      Case "3"
		DOCS_NameResponsible = "Zadavatel overtimu"
    End Select
  End If
'rmanyushin 136964 08.11.2010 End
'Запрос №46 - СТС - start
If InStr(UCase(Session("CurrentClassDoc")), UCase(STS_SLUZH_ZAPISKA_OVERTIME_PLAN)) = 1 or InStr(UCase(Session("CurrentClassDoc")), UCase(STS_SLUZH_ZAPISKA_OVERTIME_FACT)) = 1 Then
    Select Case UCase(Request("l"))
      Case "RU"
        DOCS_NameResponsible = "Руководитель департамента сотрудника, выполняющего переработку"
		DOCS_Correspondent = "Функциональные руководители"
      Case ""
        'EN
        DOCS_NameResponsible = "Department Director of the employee working overtime"
		DOCS_Correspondent = "Functional Manager"
      Case "3"
        DOCS_NameResponsible = "Vedoucí oddělení zaměstnance, který vykonává práci přesčas"
		DOCS_Correspondent = "Funkcionální vedoucí"
    End Select
  End If
'Запрос №46 - СТС - end


'Phil 20080817
'ph - 20081205 - Changed - start
  If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) > 0 Then 
    Select Case UCase(Request("l"))
      Case "RU"
        'RU
        DOCS_Author = "Инициатор"
        'ph - 20090820
        DOCS_AmountDoc = "Сумма с НДС"
        'rmanyushin 119191 17.08.2010
        DOCS_Content = "Дополнительная информация"  
      Case ""
        'EN
'        DOCS_Author = "Initiator"
         DOCS_Author = "Drafter"
        'ph - 20090820
        DOCS_AmountDoc = "Amount incl. VAT"
        'rmanyushin 119191 17.08.2010
        DOCS_Content = "Additional information"  
      Case "3"
        'CZ
        DOCS_Author = "Iniciátor"

        DOCS_Description = "Stručný popis"
		DOCS_Currency = "Kód měny"
		DOCS_AdditionalUsers = "Další uživatelé"
		DOCS_SecurityLevel = "Stupeň důvěrnosti"
		DOCS_ListToView = "Dokument četli"
		DOCS_ListToReconcile = "Schvaluje"
		DOCS_NameAproval = "Autorizuje"
        'ph - 20090820
        DOCS_AmountDoc = "Částka vč. DPH"
        'rmanyushin 119191 17.08.2010
        DOCS_Content = "Dodatečná informace"
    End Select
  End If
  If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) > 0 Then 
    Select Case UCase(Request("l"))
      Case "RU"
        'RU
        DOCS_Description = "Основание платежа"
		DOCS_DateCompletion = "Срок оплаты"
        DOCS_Author = "Инициатор"
        'ph - 20090820
        DOCS_AmountDoc = "Сумма с НДС"
      Case ""
        'EN
        DOCS_Description = "Payment justification"
		DOCS_DateCompletion = "Terms of payment"
'        DOCS_Author = "Initiator"
         DOCS_Author = "Drafter"
        'ph - 20090820
        DOCS_AmountDoc = "Amount incl. VAT"
      Case "3"
        'CZ
        DOCS_Description = "Důvod platby"
		DOCS_DateCompletion = "Datum splatnosti"
        DOCS_Author = "Iniciátor"

		DOCS_Description = "Stručný popis"
		DOCS_Currency = "Kód měny"
		DOCS_AdditionalUsers = "Další uživatelé"
		DOCS_SecurityLevel = "Stupeň důvěrnosti"
		DOCS_ListToView = "Dokument četli"
		DOCS_ListToReconcile = "Schvaluje"
		DOCS_NameAproval = "Autorizuje"
        'ph - 20090820
        DOCS_AmountDoc = "Částka vč. DPH"
    End Select
  End If
'ph - 20081205 - Changed - end
  
'20090622 - Заявка ТКП
  If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_COM_OFFERS)) > 0 Then
    'Не показывать кнопку добавления пользователей в поле ListToView при просмотре
	VAR_AddUsersToListToView=""
    'На согласование 2 дня
    Var_nDaysToReconcile = 2
    Select Case UCase(Request("l"))
      Case "RU" 'RU
        DOCS_DocID = "Номер документа"
        DOCS_DocIDParent = "Документ-основание"
        DOCS_Name = "Тема"
        DOCS_Author = "Инициатор"
        DOCS_NameAproval = "Подписант"
        DOCS_PartnerName="Получатель (Организация)"
'Запрос №1 - СИБ - start
        If InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 Then
          DOCS_ListToView="Менеджер по работе с клиентами/Руководитель проекта (получатель)"
		Else
'Запрос №1 - СИБ - end
          DOCS_ListToView="Менеджер по работе с клиентами (получатель)"
'Запрос №1 - СИБ - start
		End If
'Запрос №1 - СИБ - end
       DOCS_DateActivation="Дата и время создания"
      Case "" 'EN
        DOCS_DocID = "Document number"
        DOCS_DocIDParent = "Basis document"
        DOCS_Name = "Subject"
        DOCS_Author = "Drafter"
        DOCS_PartnerName="Recipient (Organization)"
        DOCS_ListToView="Sales manager/KAM (recipient)"
        DOCS_DateActivation="Date and time of making"
      Case "3" 'CZ
        DOCS_DocID = "Číslo dokumentu"
        DOCS_DocIDParent = "Osnova dokumentu"
        DOCS_Name = "Předmět"
        DOCS_Author = "Iniciátor"
        DOCS_NameAproval = "Podepisovatel"
        DOCS_PartnerName="Příjemce (Organizace)"
        DOCS_ListToView="Sales Manger/KAM (příjemce)"
        DOCS_DateActivation="Datum a čas vytvoření"
    End Select
  End If

'Ph - 20081109 - Start
  If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_RASP_DOCS)) > 0 or InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) > 0 or InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_SLUZH_ZAPISKA)) > 0 Then 
    If UCase(Request("l")) = "RU" Then
      But_Approve = "Подписать"
      But_RefuseApp = "Отклонить"
      DOCS_APROVALREQUIRED = "Подписание"
      DOCS_RefusedApp = "Отклонено"
	  DOCS_Approved = "Подписано"
	  DOCS_APPROVE = "Подписать"
	  DOCS_RefuseApp = "Отклонить подписание"
	  DOCS_Approving="На подписании"
	End If
  End If
'Ph - 20081109 - End

'vnik_payment_order
If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_PAYMENT_ORDER)) > 0 Then
    If UCase(Request("l")) = "RU" Then
      But_Approve = "Оплатить" 'кнопка подписать
      'But_RefuseApp = "Отклонить"
      DOCS_APROVALREQUIRED = "Ожидает оплаты" ' перед подписанием
      'DOCS_RefusedApp = "Отклонено"
	  DOCS_Approved = "Оплачено" 'после подписания
	  DOCS_APPROVE = "Оплатить" 'вопрос при утверждении 
	  'DOCS_RefuseApp = "Отклонить подписание"
	  'DOCS_Approving="На подписании"
	  DOCS_NameAproval = "Бухгалтер по оплате"
	  DOCS_APROVAL = "Оплата согласована"
	End If
End If
'vnik_payment_order

'rti_payment_order
If InStr(UCase(Session("CurrentClassDoc")),UCase(RTI_PAYMENT_ORDER)) > 0 Then
    If UCase(Request("l")) = "RU" Then
      But_Approve = "Оплатить" 'кнопка подписать
      DOCS_APROVALREQUIRED = "Ожидает оплаты" ' перед подписанием
	  DOCS_Approved = "Оплачено" 'после подписания
	  DOCS_APPROVE = "Оплатить" 'вопрос при утверждении 
	  DOCS_NameAproval = "Бухгалтер по оплате"
	  DOCS_APROVAL = "Оплата согласована"
	End If
End If
'rti_payment_order

'rti_purchase_order
If InStr(UCase(Session("CurrentClassDoc")),UCase(RTI_PURCHASE_ORDER)) > 0 Then
    If UCase(Request("l")) = "RU" Then
      DOCS_Approved = "Утверждено" 'после подписания
	End If
End If
'rti_purchase_order

'rti_contract
If InStr(UCase(Session("CurrentClassDoc")),UCase(RTI_CONTRACT)) > 0 Then
    If UCase(Request("l")) = "RU" Then
      DOCS_Content="Содержание документа"
	End If
End If
'rti_contract

'количество дней на согласование
If InStr(UCase(Session("CurrentClassDoc")),UCase(RTI_PAYMENT_ORDER)) = 1 OR InStr(UCase(Session("CurrentClassDoc")),UCase(RTI_PURCHASE_ORDER)) = 1 or InStr(UCase(Session("CurrentClassDoc")),UCase(RTI_PROTOCOL)) = 1 or InStr(UCase(Session("CurrentClassDoc")),UCase(RTI_CONTRACT)) = 1 Then
    VAR_nDaysToReconcile = 2
End If

'rti заявки
If InStr(UCase(Session("CurrentClassDoc")),UCase(RTI_PAYMENT_ORDER)) = 1 OR InStr(UCase(Session("CurrentClassDoc")),UCase(RTI_PURCHASE_ORDER)) = 1 OR InStr(UCase(Session("CurrentClassDoc")),UCase(RTI_BSAP)) = 1 or InStr(UCase(Session("CurrentClassDoc")),UCase(RTI_PROTOCOL)) = 1 or InStr(UCase(Session("CurrentClassDoc")),UCase(RTI_CONTRACT)) = 1 Then
    If UCase(Request("l")) = "RU" Then
        DOCS_ListToView="Получатели"
	End If
End If

'rti заявки
'AMW - MIKRON 
'кнопки для Заявки на оплату
If InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_PAYMENT_ORDER)) > 0 Then
   If UCase(Request("l")) = "RU" Then
      DOCS_Name = "Описание товаров/услуг/работ"
      But_Approve = "Оплатить" 'кнопка подписать
      DOCS_APROVALREQUIRED = "Ожидает оплаты" ' перед подписанием
	  DOCS_Approved = "Оплачено" 'после подписания
	  DOCS_APPROVE = "Оплатить" 'вопрос при утверждении 
	  DOCS_NameAproval = "Бухгалтер по оплате"
	  DOCS_APROVAL = "Оплата согласована"
   End If
End If

'кнопки для Протокола ЗК
If InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_PROTOCOL)) > 0 Then
   If UCase(Request("l")) = "RU" Then
      DOCS_Name = "Описание товаров/услуг/работ"
      DOCS_Author = "Секретарь ЗК"
      DOCS_NameAproval = "Председатель ЗК"
      DOCS_Content = "Содержание документа"
   End If
End If

'кнопки для опросного листа
If InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_RL_PROTOCOL)) > 0 Then
   If UCase(Request("l")) = "RU" Then
      DOCS_Name = "Описание товаров/услуг/работ"
'      DOCS_Name = "Вопрос, поставленный на голосование"
'      BUT_VISA = "Голосование"
'      But_Approve = "ЗА"
      But_AGREE = "Голосовать"
'      But_Refuse = "Против"
'      But_RefuseApp = "Воздержался"
'      But_Comment = "Воздержался 1"
'      But_Approve = "ЗА"
'      DOCS_Content = "Содержание документа"
      DOCS_NameAproval = "Секретарь ЗК"
      DOCS_ListToReconcile = "Закупочная комиссия"
      ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickRefuse"
   End If
End If

'кнопки для справки о ценах к листу согласования
If InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_RL_MEMO)) > 0 Then
   If UCase(Request("l")) = "RU" Then
      DOCS_Name = "Описание товаров/услуг/работ"
      DOCS_Author = "Уполномоченный по закупке"
      DOCS_NameAproval = "Утверждающий"
      DOCS_DocIDPrevious = "Заключенный договор"
      DOCS_DocIDIncoming = "Заключенный договор 2"
      DOCS_DocIDadd = "Заключенный договор 3"
      DOCS_InventoryUnit = "Примечание"
      DOCS_AmountDoc = "Сумма(с НДС)"
      DOCS_Description = "Выбрано по критерию(-ям)"
'      DOCS_Name = "Вопрос, поставленный на голосование"
   End If
End If

If InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_BSAP)) > 0 Then
   If UCase(Request("l")) = "RU" Then
      DOCS_Name = "Описание товаров/услуг/работ"
      DOCS_Author = "Уполномоченный по закупке"
      DOCS_NameAproval = "Утверждающий"
      DOCS_Description = "Выбрано по критерию(-ям)"
   End If
End If

If InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_PURCHASE_ORDER)) > 0 Then
   If UCase(Request("l")) = "RU" Then
      DOCS_Name = "Описание товаров/услуг/работ"
      DOCS_Author = "Уполномоченный по закупке"
   End If
End If

'amw 25-10-2013
'кнопки для заявки, БСАП, протокол ЗК, договора
'Скрываем от пользователей кнопку "Подчиненный", вместо этого вставляем свою в UserClick.asp
If UCase(Request("l")) = "RU" and _
   (InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_PAYMENT_ORDER)) > 0 or _
    InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_PURCHASE_ORDER)) > 0 or _
    InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_BSAP)) > 0 or _
    InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_PROTOCOL)) > 0 or _
    InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_S_CONTRACT)) > 0 or _
    InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_CONTRACT)) > 0 ) Then
       DOCS_ListToView="Получатели"
'       If InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_CONTRACT)) = 0 Then
'          ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickCreateDocDependant"
'       End If
End If

'кнопки для договоров Микрон
If InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_CONTRACT)) = 1 or _
   InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_S_CONTRACT)) = 1 or _
   InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_ADD_CONTRACT)) = 1 or _
   InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_EXPORT_CONTRACT)) = 1 or _
   InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_EXPADD_CONTRACT)) = 1 Then
   Select Case UCase(Request("l"))
      Case "RU" 'RU
        DOCS_DocID = "Регистрационный № договора/Код договора"
        DOCS_DocIDPrevious = "Регистрационный № основного договора"
        DOCS_Name = "Вид обязательств"
        DOCS_Author = "Уполномоченный по договору"
        DOCS_NameResponsible = "Ответственный по договору"
        DOCS_NameAproval = "Утверждающий"
        DOCS_Description = "Предмет договора"
        DOCS_DocIDIncoming = "Номер (внешний)"
        DOCS_InventoryUnit = "Примечание"
        DOCS_AmountDoc = "Сумма договора(включая НДС)"
        If InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_ADD_CONTRACT)) = 1 or _
           InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_EXPADD_CONTRACT)) = 1 Then
           DOCS_DocIDParent = "Регистрационный № основного договора"
           DOCS_DocIDPrevious = "Индекс предшествующего документа"
        End If

      Case ""   'EN
        DOCS_DocID = "Contract registration No./Contract code"
        DOCS_Name = "Type of liabilities"
        DOCS_Description = "Subject of the contract"
        DOCS_DocIDIncoming = "Number (external)"
        DOCS_InventoryUnit = "Comments"
        DOCS_AmountDoc = "Price of contract (including VAT)"

      Case "3"   'CZ
        DOCS_DocID = "Registrační číslo smlouvy/Kód smlouvy"
        DOCS_Name = "Druh závazků"
        DOCS_Description = "Předmět smlouvy"
        DOCS_DocIDIncoming = "Číslo (externí)"
        DOCS_InventoryUnit = "Poznámky"
        DOCS_AmountDoc = "Cena kontraktu (včetně DPH)"
   End Select
End If

'AMW - MIKRON


'End If
'phil - 20080913 - End

STS_SecrPravlenia = """#Секретарь правления СТС"";"
SIT_SecrPravlenia = """#Секретарь правления УК"";"

PORUCHENIA_PRAVLENIA_RU = "Поручения Правления"
PORUCHENIA_PRAVLENIA_EN = "Tasks of Management Board"
PORUCHENIA_PRAVLENIA_CZ = "Úkoly správní rady"
'Ph - 20081117 - end

' SAY 2008-08-21
'If Not IsAdmin() Then
  'Согласование
  VAR_ButtonsNotToShow=VAR_ButtonsNotToShow+"ClickReconciliationComplete,ClickReconciliationWaiting,ClickToModify,ClickReconciliationSuspend,"

  'Редактирование
  VAR_ButtonsNotToShow=VAR_ButtonsNotToShow+"ClickCreateConnected,ClickCreateConnectedCopy,ClickModifyNameCreation,ClickCopyDoc,"

  'Относящиеся
  'VAR_ButtonsNotToShow=VAR_ButtonsNotToShow+"ClickCreateMessage,ClickCreateCommentHistory,ClickESign.asp,ClickCreateCommentReview,ClickCreateCommentPartner,ClickCreateCommentResource,ClickCreateCommentLink,ClickCreateCommentBPStep,ClickCreateContact,ClickCreateEvent,ClickCheckUsers,ClickGetBarCode,"
  VAR_ButtonsNotToShow=VAR_ButtonsNotToShow+"ClickCreateMessage,ClickESign.asp,ClickCreateCommentReview,ClickCreateCommentPartner,ClickCreateCommentResource,ClickCreateCommentLink,ClickCreateCommentBPStep,ClickCreateContact,ClickCreateEvent,ClickCheckUsers,ClickGetBarCode,ClickShowReports,ClickCreateDocFollowing,"
  ' убираем "Ход Исполнения" для всех кроме Поручений
  
  'Администрирование
  VAR_ButtonsNotToShow=VAR_ButtonsNotToShow+"ClickMakeArchival,ClickHome"

  'создание ярлыка
  If InStr(UCase(Request.ServerVariables("URL")),UCase("/home.asp"))=0 Then
    VAR_ButtonsNotToShow=VAR_ButtonsNotToShow+"ClickCreateShortcut,"  
  End If

'End If

'SAY 2008-11-11 скрываем кнопки правого меню для категорий документов
    If InStr(UCase(S_ClassDoc),UCase(SIT_ZADACHI)) > 0 Then
       VAR_ButtonsNotToShow=VAR_ButtonsNotToShow+""
    Else
       'VAR_ButtonsNotToShow=VAR_ButtonsNotToShow+"ClickCreateComment,ClickCreateCommentHistory,"
       'SAY 2008-11-20 возвращаем кнопку "комментарий"
       VAR_ButtonsNotToShow=VAR_ButtonsNotToShow+"ClickCreateCommentHistory,"
    End If

' ? убираем вторую пиктограмму в прикрепленном файле
VAR_NoMSWordBookmarkInserting="Y"

' прячем кнопку генерации номера документа при регисрации
VAR_ChangeDocGetNewButton="N"

'прячем лишние поля
VAR_UseShortDocumentView="Y"
'Запрещаем редактировать категорию док-та
VAR_ChangeDocNotToChangeClassDoc="Y"

'генерация номера при регистрации
'отключаем
'ph - 20100603 - start - Код в условии ниже никогда не выполняется, закомментирован для наглядности
'If False and UCase(Request.ServerVariables("URL"))=UCase("/MakeRegistered.asp") Then
'
'  If InStr(UCase(Session("Department")),UCase(SIT_SITRONICS))=1 then
'    sDepartment = SIT_SITRONICS
'  Else
'    sDepartment = SIT_STS
'  End If
'
'  S_DocID = Request("DocID")
'  S_DocID = Right(S_DocID,Len(S_DocID)-3)
'
'  sSearchCol = "DocID"
'  sPrePrefix = ""
'
'
'  If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_ISHODYASCHIE)) > 0 Then
'
'
'    sPrefix="OUT"+Right(CStr(Year(Date)),2)+"/"
'    sPostfix = ""
'    sSufix = ""
'
'    'SAY 2008-10-07 новый алгоритм нумерации для СТС
'    If InStr(S_DocID,"_") > 0 Then
'      'sSufix = Mid(S_DocID, InStr(S_DocID,"-")+1, InStr(S_DocID,"-")-InStr(S_DocID,"/") )
'       sPrefix="OUT"+Right(CStr(Year(Date)),2)+"_"+Mid(S_DocID, InStr(S_DocID,"_")+1, InStr(S_DocID,"/")-InStr(S_DocID,"_") )
'    End If
'
'    'Call GetNewDocID_test(S_ClassDoc, sDepartmentRoot, Request("DocIDParent"), Request("UserFieldText7"),  "", "PJ-")
'
'
'  End If
'
'  'SIT_RASP_DOCS
'  If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_RASP_DOCS)) > 0 Then
'
'    'SAY 2008-10-15 переделка номера для СТС
'    If sDepartment = SIT_SITRONICS Then
'      sPrefix=left(S_DocID, InStr(S_DocID,"/"))
'    Else
'      sPrefix=left(S_DocID, InStr(S_DocID,"-"))
'    End If
'    sPostfix = ""
'    sSufix=""
'  End If
'
'  If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_NORM_DOCS)) > 0 Then
'
'    sPrefix=left(S_DocID, InStr(S_DocID,"-"))
'    sPostfix = "."+right(S_DocID, len(S_DocID)-InStr(S_DocID,"."))
'    sSufix=""
'
'    End If
'  ' 2008-12-25 временно отключаем генерацию регистрационного номера для нормативных документов
'  If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_NORM_DOCS)) = 0 Then
'    Call GetNewDocIDForClassDocWithPrefixNew(Session("CurrentClassDoc"), sSearchCol, sPrePrefix, sPrefix, sSufix, sPostfix, sDepartment)
'  End If
'' AM 14082008 Делаем рег.номер редактируемым для регистрации задним числом  S_DocID_Set = S_DocID
'
'
'End If
'ph - 20100603 - end

SIT_HelpDesk = "HelpDesk"

' SAY 2009-03-19 новая регистрация
If UCase(Request.ServerVariables("URL"))=UCase("/MakeRegistered.asp") Then

  S_DocID = Request("DocID")
  
  'vnik_scsp
  If InStr(S_DocID, "SPST") > 0 Then
    sDepartmentRoot = Trim(SIT_SCSP)
  Else
  'vnik_scsp
  sDepartmentRoot = GetRootDepartment(Session("Department"))
  'vnik_scsp
  End If
  'vnik_scsp
  'S_DocID = Replace(S_DocID,"PJ-","") 
  If InStr(S_DocID, "PJ-") = 1 Then
    S_DocID = Right(S_DocID,Len(S_DocID)-3)
  End If

  'out "S_DocID=" + S_DocID

  sParam1 = ""
  sParam2 = ""
  sParentDocID = ""

'SIT_ISHODYASCHIE - begin
  If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_ISHODYASCHIE)) > 0 Then
    If InStr(S_DocID,"5-") = 0 Then
      sParentDocID = "Param1"
    End If

'kkoshkin 02082012 begin    
'amw     If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_ISHODYASCHIE)) > 0 and inStr(sDepartmentRoot,SIT_RTI)=1  Then
     If InStr(sDepartmentRoot,SIT_RTI) = 1 Then
      'parDocIdAdd = Left(S_DocID,InStr(S_DocID,"/")-1)
      parDocIdAdd = Left(S_DocID,InStr(S_DocID,"/")) 'чтобы "/" остался
  End If
'kkoshkin 02082012 end


    If InStr(S_DocID,"_") > 0 Then
      sParam2 = Mid(S_DocID, InStr(S_DocID,"_")+1, InStr(S_DocID,"/")-InStr(S_DocID,"_")-1 ) + " Param"
    End If

'vnik mikron
  sSQL = "select code from Departments where Name=N'"+Request("DocDepartment")+"' or Name=N'"+Request("DocDepartment")+"/'"
  'sSQL = "SELECT Code FROM Departments where name like N'%"+Request("DocDepartment")+"%'"
  
  NewConnection = True
  If not IsNull(Conn) Then
    NewConnection = not IsObject(Conn)
  End If

  If NewConnection Then
    SetPayDoxPars 'Присваиваем переменные среды из Global.asa
    sConnStr = "ConnectString"
    Select Case UCase(Request("l"))
      Case "RU" sConnStr = sConnStr + "RUS"
      Case "3" sConnStr = sConnStr + "3"
    End Select

    Set MyConn = Server.CreateObject("ADODB.Connection")
    MyConn.Open Application(sConnStr)
  Else
    MyConn = Conn
  End If
  
  Set dsTempPR = Server.CreateObject("ADODB.Recordset")
  dsTempPR.Open sSQL, MyConn, 3, 1, &H1
  
    if not dsTempPR.EOF Then
        sParam1 = dsTempPR("code")
        if InStrRev(sDepartmentCode, "/") > 0 then
        sParam1 = Right(sParam1, Len(sParam1)-InStrRev(sParam1, "/"))
        End If
    End If

     dsTempPR.Close
'vnik mikron
  End If
'SIT_ISHODYASCHIE - end

'SIT_RASP_DOCS - begin
  If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_RASP_DOCS)) > 0 Then

    If sDepartmentRoot = SIT_SITRONICS Then
        sParam1 = Left(S_DocID, InStr(S_DocID,"-")-1)  + " Param"
    ElseIf sDepartmentRoot = SIT_STS Then
        sParam1 = Left(S_DocID, InStr(S_DocID,"_")-1)  + " Param"   
'Запрос №1 - СИБ - start
    ElseIf sDepartmentRoot = SIT_SIB Then
		'Проектный номер имеет вид PJ-(OR|Р)YY_SIB-001, т.е. разделителя между годом и подтипом 
		'документа нет, привязываемся к "_SIB" (к первой попавшейся цифре получалось громоздко)
        sParam1 = Left(S_DocID, InStr(S_DocID,"_SIB")-3)  + " Param"   
'Запрос №1 - СИБ - end
    Else 'Другие БН
        sParam1 = Left(S_DocID, InStr(S_DocID,"-")-1)  + " Param"   
    End If

  End If

'SIT_NORM_DOCS - begin
  If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_NORM_DOCS)) > 0 Then
        sParam1 = Left(S_DocID, InStr(S_DocID,"-")-1)   + " Param"  
        sParam2 = Right(S_DocID, Len(S_DocID) - InStr(S_DocID,"."))   
  End If
  
  'kkoshkin 06-09-2012 start
  If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_RASP_DOCS)) > 0 and inStr(sDepartmentRoot,SIT_RTI)=1 Then
    parDocIdAdd = Left(S_DocID,InStr(S_DocID,"/")) 'чтобы "/" остался
  end if
  'kkoshkin 06-09-2012 end

  'out "S_DocID="+ S_DocID + ", sParam1=" + sParam1 + ", sParam2=" + sParam2
  ' 2008-12-25 временно отключаем генерацию регистрационного номера для нормативных документов и протоколов
'vnik_protocols
  If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_NORM_DOCS)) > 0 or InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_PROTOCOLS)) > 0 or InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_PROTOCOLS_CPC)) > 0 or InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_CONTRACTS_MC)) > 0 Then
'Запрос №31 - СТС - start
'kkoshkin 02082012 begin
'ИСХОДЯЩИЕ для РТИ
  ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_ISHODYASCHIE)) = 1 and inStr(sDepartmentRoot,SIT_RTI)=1 Then
    S_DocID =  GetNewDocIDForIshodyashie(Session("CurrentClassDoc"), sDepartmentRoot, parDocIdAdd, sParam2,  "", "", "") 'BU пустая, код подставляется из sParam2
'kkoshkin 02082012 begin   
'ИСХОДЯЩИЕ для остальных
  ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_ISHODYASCHIE)) = 1 and inStr(sDepartmentRoot,SIT_RTI)=0 Then
     S_DocID =  GetNewDocIDForIshodyashie(Session("CurrentClassDoc"), sDepartmentRoot, sParam1, sParam2,  "", "", "") 'BU пустая, код подставляется из sParam2

'РАСПОРЯДИТЕЛЬНЫЕ для РТИ
  ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_RASP_DOCS)) = 1 and inStr(sDepartmentRoot,SIT_RTI)=1 Then 
     S_DocID =  GetNewDocIDForRaspDocs(Session("CurrentClassDoc"), sDepartmentRoot, parDocIdAdd, sParam2,  "", "", "") 'BU пустая, код входит в sParam1
'РАСПОРЯДИТЕЛЬНЫЕ для остальных
  ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_RASP_DOCS)) = 1 and inStr(sDepartmentRoot,SIT_RTI)=0 Then
     S_DocID =  GetNewDocIDForRaspDocs(Session("CurrentClassDoc"), sDepartmentRoot, sParam1, sParam2,  "", "", "") 'BU пустая, код входит в sParam1

'Запрос №31 - СТС - end
  Else 
'vnik_protocols
    'Call GetNewDocIDForClassDocWithPrefixNew(Session("CurrentClassDoc"), sSearchCol, sPrePrefix, sPrefix, sSufix, sPostfix, sDepartment)
    Call GetNewDocID_test(Session("CurrentClassDoc"), sDepartmentRoot, sParam1, sParam2,  "", "")
  End If

' AM 14082008 Делаем рег.номер редактируемым для регистрации задним числом
' S_DocID_Set = S_DocID

End If

'убираем кнопку "главная" в правом меню
VAR_CountRecordsForBigButtons=0

'phil- 20100118
Var_ReconciliationIfAllAgreeDoesNotExcludeCurrentLevel=""
'Var_ReconciliationIfAllAgreeDoesNotExcludeCurrentLevel="Y"
Var_ApprovalIfAllAgree=True

'phil- 20080918 - start
'При загрузке файлов убрать: кнопку "Взять файлы на сервере"
VAR_UseFTPUploads=""
'В разделе "ход исполнения" (при исполнении поручений) убрать поле с возможностью отсылки уведомлений участникам
VAR_NotToShowParticipantsFieldInCommentForm="Y"
'phil- 20080918 - end

'Ph- 20080921
'Необходимо отключить возможность входа по паролю, только Windows Аутентификация
VAR_NoPayDoxLogin="Y"
'При загрузке файлов убрать: поле "Добавить список согласования файла", флаг страны (при нажатии на него ничего не происходит), снизу флажок, кнопку сканировать и иконку установки клиентской части PayDox (появляется после попытки сканировать и не работает)
VAR_PermitToAgreeFiles=False
'Инструкции на главной странице лежат в архиве, мы договаривались, что они будут в WORD, прошу исправить
VAR_PayDoxQuickStart="PayDoxQuickStartRUS.doc"
If UCase(Request("l")) = "RU" Then
  VAR_PayDoxUserManual="PayDoxUserManualRUS.doc"
Else
  VAR_PayDoxUserManual="PayDoxUserManual.doc"
End If
'Исходящие: поле "документ от" - необходимо убрать значок часов и изменить указанный формат только на дату
VAR_NotToChooseTime="Y"
'Отключаем четвертый язык
bUseLang4 = False
'Ph- 20080921

'Ph - 20080922 - Защита шаблонов паролем
SIT_ISHODYASHIE_DOC = "Исходящие документы*Outgoing correspondence*Odchozí dokumenty"
SIT_PORUCHENYA = "Поручения*Tasks*Úkoly"
SIT_PROTOCOLY_VSTRECH = "Протоколы*Protocols*Protokoly/Встреч*Meetings*Schůze"
SIT_RASPORYADIT_DOC = "Распорядительные документы*Orders*Nařízení"

VAR_ClassDocListToProtectByPassword = ""'SIT_SLUZH_ZAPISKA_OBSCHAY+","+SIT_SLUZH_ZAPISKA_COMPUTER+","+SIT_SLUZH_ZAPISKA_MOBILE+","+SIT_SLUZH_ZAPISKA_KOMANDIROVKA+","+SIT_SLUZH_ZAPISKA_OBUCHENIE+","+SIT_SLUZH_ZAPISKA_PERSONAL+","+SIT_ISHODYASHIE_DOC+","+SIT_PORUCHENYA+","+SIT_PROTOCOLY_VSTRECH+","+SIT_RASPORYADIT_DOC

'Ph - 20081006 - Разрешить доступ к родительскому из подчиненного
VAR_PermitAccessToParentDocFromDependant= "Y"

' AM 24092008
If Request("l")="ru" Then
  DOCS_USERMANUAL="Получить описание системы PayDox"
  DOCS_AskSupport="Задайте Ваш вопрос службе поддержки"
End If

'SAY 2008-10-08
'If (InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_ISHODYASCHIE)) > 0 and (InStr(Session("UserID"),"registrator") > 0) or InStr(UCase(Session("Permitions")),"K") >0 ) Then
If InStr(Session("UserID"),"registrator") > 0 or InStr(UCase(Session("Permitions")),"K") >0  Then

  'out " reg: VAR_AgreeAgainInitiallyNotChecked="+VAR_AgreeAgainInitiallyNotChecked+", VAR_UploadFileForcesAgreeAgain="+VAR_UploadFileForcesAgreeAgain

  VAR_AgreeAgainInitiallyNotChecked="Y"
  VAR_UploadFileForcesAgreeAgain="N"
  VAR_CanCreateMainVersionFiles=True

End If


  SIT_HelpDesk = "HelpDesk"
'SAY 2008-11-18 добавил для тестового helpdesk
DOCS_AllPagesComment = "Ver. 2.0"
If UCase(Request.ServerVariables("SERVER_NAME")) <> UCase("gl-paydox-01.global.sitronics.com") and UCase(Request.ServerVariables("SERVER_NAME")) <> UCase("gl-paydox-01") and UCase(Request.ServerVariables("SERVER_NAME"))  <> UCase("paydox.oaorti.ru") and Request.ServerVariables("SERVER_NAME") <> "172.26.0.180" Then
  DOCS_AllPagesComment = DOCS_AllPagesComment + "<font color = red> - TEST SERVER</font>"

  'SAY 2008-11-06
'  Var_PossibleApplicationTypes="HelpDesk, Документооборот"
'  Var_PossibleApplicationTypes=",HelpDesk"

'Изменился интерфейс переключения между конфигурациями. HelpDesk оставляем только Администраторам
If IsAdmin() and InStr(UCase(Session("UserID")), "ADMIN") > 0 Then
    'Var_PossibleApplicationTypes="HelpDesk"
    
    'rmanyushin 63828 23.11.2009
    Var_PossibleApplicationTypes=DOCS_BUSINESSPROCESSES+","+DOCS_DOCUMENTS+","+DOCS_AppType_HelpDesk

Else
  Var_PossibleApplicationTypes=""
End If

End If

'If InStr(UCase(Session("Department")),UCase(SIT_SITRONICS))=1 then
'  sDepartment = SIT_SITRONICS
'Else
'  sDepartment = SIT_STS
'End If
sDepartment = GetRootDepartment(Session("Department"))

SIT_UsersListToAccessAllCategoryDocs = ReplaceRoleFromDir(SIT_Registrar, sDepartment)
 
' Доступ на чтение ко всем документам категории Договоры пользователям, указанным в параметре 
' SIT_UsersListToAccessAllCategoryDocs
'ph - 20090714 - start
'If Instr(UCase(SIT_UsersListToAccessAllCategoryDocs),UCase(Session("UserID")))>0 Then
'  VAR_ReadAccess="Y"
'End If

If Instr(UCase(SIT_UsersListToAccessAllCategoryDocs),"<"+UCase(Session("UserID"))+">")>0 Then
  'В списке доступ разрешаем, ограничивается он в SkipThisRecord
  If InStr(UCase(Request.ServerVariables("URL")),UCase("/ListDoc.asp"))>0 Then
    VAR_ReadAccess="Y"
  End If
  'Доступ к файлам отключаем кроме как из просмотра карточки
  If InStr(UCase(Request.ServerVariables("URL")),UCase("/ShowDoc.asp"))=0 Then
    Session("SIT_UserCanDownloadFiles") = False
  End If
End If
'ph - 20090714 - end

'Число символов поля, показываемых в списках
VAR_TextLenToShowInLists=100

'отмечать чекбокс "основная резолюция"
VAR_MainResolutionChecked="Y"
VAR_MainResolutionNoDates="Y"

'out "UserID="+Session("UserID")+", Permitions="+Session("Permitions")

'SAY 2008-10-31 срок согласования для нормативных документов 
If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_NORM_DOCS)) > 0 Then 
  Var_nDaysToReconcile=4
End If

'-------New example
'Redim S_DependantFieldNames(3)
'S_DependantFieldNames(1)="DependantFieldName1"
'S_DependantFieldNames(2)="DependantFieldName2"
'S_DependantFieldNames(3)="DependantFieldName3"
'S_DependantName="DependantName"
'S_DependantTableName="DepartmentDependants"
'S_DependantOrderBy=" Order By Parameter1, Parameter2, Parameter3"

'Ph - 20081117 - Не показывать вкладки
VAR_TabsHidden = "y"

'Ph - 20081122 - Отключение расчета маршрута в заявках на закупку
'TempVAR_DisablePurchaseOrdersReconcilationListAutoDetermination = "Y"

'SAY 2008-11-28
If not ISAdmin() Then
  VAR_ButtonsNotToShow=VAR_ButtonsNotToShow+",ClickCreateRequest"
End If

'ph - 20081203 - start
S_DependantName = "Руководители подразделения*Department leaders*Department leaders"
S_DependantTableName = "DepartmentDependants"

Redim S_DependantFieldNames(2)
S_DependantFieldNames(1) = "Бизнес единица*Business unit*Obchodní jednotka"
S_DependantFieldNames(2) = "Руководитель*Leader*Nadřízený"

S_DependantOrderBy = " order by BusinessUnit"
'ph - 20081203 - end

'rmanyushin 136151 13.10.2010 Start 
    'Заявки на закупку и оплату согласуются 2 дня
    'If (InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) > 0) or (InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) > 0) Then
    ' VAR_nDaysToReconcile = 2
    'End If

If (InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) > 0) Then
      If is5DivisionSTS(Session("Department")) Then
		    VAR_nDaysToReconcile = 1
      Else
		    VAR_nDaysToReconcile = 2
      End If
End If

If (InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) > 0) Then
  VAR_nDaysToReconcile = 2
End If
'rmanyushin 136151 13.10.2010 End 

'Запрос №35 - СТС - start
If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) = 1 Then
    VAR_nDaysToReconcile = 2
' {ph - 20120229
    If InStr(Session("Department"), SIT_STS_ROOT_DEPARTMENT) = 1 Then 'Для СТС другой срок
        VAR_nDaysToReconcile = 1
    End If
' ph - 20120229}
End If
'Запрос №35 - СТС - end

'Ph - 20090128 - start - 
'справочники прячем средствами CurrentProhibitedDirectoryGUIDs, т.к. нужно будет некоторым их 
'выборочно показывать
VAR_NotToShowUserDirectories = ""
If not IsAdmin() Then
'  CurrentProhibitedDirectoryGUIDs = "{3685D3AA-FB15-4ECF-993F-8AC5AB87F4D6},{6D42F0FB-1389-4CA0-8A00-9E0CD3F09CF7},{459D6AE1-7E6F-467E-B162-A36A4054AA5A},{BFC71550-2605-4679-8A3F-C04211891D7E},{C632B46B-3AAF-4607-BBC5-AC51C0A4971B},{535C7740-CB1E-4403-ABC6-93AFC67205D5},{3ECADCD6-0985-4659-8774-C8C9D77EE381},{2CC714EB-5836-49E9-B873-3A34EDB85098},{E1F4F724-9DB7-40E2-92E6-D0332E268346},{959D450F-9E5A-4358-B445-1D082041987A},{78961D78-DB41-4483-99AF-C36BD0A98701},{71B3A81A-0A45-41BB-8DAE-A0EB1DA292A6},{3FA80187-4062-4E50-B6E8-3118263DF690},{7C9FC910-7D81-413B-AF08-6CD6495C6BB3},{9B59137D-CB56-4FC2-AE7D-AF490ADE2A79},{F70782E3-B1A3-4AFE-B800-905763A24E70},{DA5960BE-A65D-4D21-BF89-73233FFEAEE8},{CAAA819C-DBBA-4B38-9001-58CD15FDC678},{F0103F47-DA1C-47BC-ACAD-DE69AAF0F852},{7E4C27B9-CF2A-4620-8B98-7CBF965CB93A},{E0E79CEC-5DDE-4184-92BE-85556566BD14},{D68C37BE-9EBA-4CA8-B0D2-C6369123E7C7},{37E16CD5-BC8F-4D0C-9569-D14DAA895440},{4E1DBDA5-4578-4D01-8926-86ED79EBBC9D},{FC8A0260-6B28-4F9F-BF2D-6F95DDE21E1C},{3885C48E-1CDB-4D59-A89C-DDDD6FB19FF3},{653D0ABF-6E6E-4DCA-899F-51D44A3DB2C5},{E30E310A-67DC-4BE2-8023-61F3214EAB4D},{A75D5931-546E-401E-8925-71A7EA75889E},{C9E55EA0-3AC8-4D26-9C74-17FA135EF1A5},{D6E49442-500A-42CA-AF56-37FE90DAFE3C},{4C4E59F5-3DCF-47B0-BF4E-38D95F626746},{69B7CF1D-8962-4EBC-AFCB-DF178263ADEE},{E4BF3B18-ACAE-4E77-8BDF-34CF75481C34},{B0277016-EB62-41E1-BD5A-960A98F7FEBC},{B099BE4E-13AB-403C-85B7-D5C5C2143CBE},{2FC22FA3-CCC7-41F9-8137-F907D888C999},{8E24E3EF-F350-4D29-8BA0-430E425F54E0},{ECEAF686-3552-44BD-A49B-941376AE4109},{2AE2C457-96FE-4379-BC33-BA048E4C06B8},{3A4F4557-A6E8-4382-A69F-59CF8895645F},{33F9C053-E51D-4738-91CD-45ABB82C1D8A},{2FC22FA3-CCC7-41F9-8137-F907DC9C1F24},{3012D11D-199C-4D46-8B58-6704EEF4A3EF},{521C56BD-EC92-4AF5-BE8C-229391C37673},{8F0D8C83-05F9-4148-96E9-3D015143063F},{3B0BABA9-EF20-47A0-A026-08DA34B9A7F7},{D733E4E8-0418-4B99-BCF5-8FC5CB9C5C42},{1B136CF3-83CF-471D-925A-EEB72BC6CD5B},{1EBA180D-7657-4E72-A678-9ECE4EDD58C1},{46A1E43B-E1AE-419A-AB44-6EECF93D3C7F},{F68941D5-DD4C-443F-90E3-39F5D16BFD13},{F9A6AADA-7DDD-4776-836A-A3EE4032D957},{7E9A8B94-3C6E-4597-9B09-FCABD40BB155},{6D57662F-7DD0-41E1-806B-3562412FDFAF},{0D620DAB-1B89-4E7B-BB6A-29EB77F9AEE9},{7C7058BE-F586-44C4-B5BE-47D2E05E96BD},{AF173F9A-4724-405B-AA9C-C72E1DCA7647},{2D84D273-5F30-4D17-A8C1-6308862FBBE7},{F5EB2A97-423D-4D9A-A338-67A72AC26F77},{6A200BD7-1A53-40FC-9DBB-44499F65B74C},{2F4D0C04-FD15-4321-A5E3-5AA2FCB0D70E}"
  CurrentProhibitedDirectoryGUIDs = "{3685D3AA-FB15-4ECF-993F-8AC5AB87F4D6},{6D42F0FB-1389-4CA0-8A00-9E0CD3F09CF7},{459D6AE1-7E6F-467E-B162-A36A4054AA5A},{BFC71550-2605-4679-8A3F-C04211891D7E},{C632B46B-3AAF-4607-BBC5-AC51C0A4971B},{535C7740-CB1E-4403-ABC6-93AFC67205D5},{3ECADCD6-0985-4659-8774-C8C9D77EE381},{2CC714EB-5836-49E9-B873-3A34EDB85098},{E1F4F724-9DB7-40E2-92E6-D0332E268346},{959D450F-9E5A-4358-B445-1D082041987A},{78961D78-DB41-4483-99AF-C36BD0A98701},{71B3A81A-0A45-41BB-8DAE-A0EB1DA292A6},{3FA80187-4062-4E50-B6E8-3118263DF690},{7C9FC910-7D81-413B-AF08-6CD6495C6BB3},{9B59137D-CB56-4FC2-AE7D-AF490ADE2A79},{F70782E3-B1A3-4AFE-B800-905763A24E70},{DA5960BE-A65D-4D21-BF89-73233FFEAEE8},{CAAA819C-DBBA-4B38-9001-58CD15FDC678},{F0103F47-DA1C-47BC-ACAD-DE69AAF0F852},{7E4C27B9-CF2A-4620-8B98-7CBF965CB93A},{E0E79CEC-5DDE-4184-92BE-85556566BD14},{D68C37BE-9EBA-4CA8-B0D2-C6369123E7C7},{37E16CD5-BC8F-4D0C-9569-D14DAA895440},{4E1DBDA5-4578-4D01-8926-86ED79EBBC9D},{FC8A0260-6B28-4F9F-BF2D-6F95DDE21E1C},{3885C48E-1CDB-4D59-A89C-DDDD6FB19FF3},{653D0ABF-6E6E-4DCA-899F-51D44A3DB2C5},{E30E310A-67DC-4BE2-8023-61F3214EAB4D},{A75D5931-546E-401E-8925-71A7EA75889E},{C9E55EA0-3AC8-4D26-9C74-17FA135EF1A5},{D6E49442-500A-42CA-AF56-37FE90DAFE3C},{4C4E59F5-3DCF-47B0-BF4E-38D95F626746},{69B7CF1D-8962-4EBC-AFCB-DF178263ADEE},{E4BF3B18-ACAE-4E77-8BDF-34CF75481C34},{B0277016-EB62-41E1-BD5A-960A98F7FEBC},{B099BE4E-13AB-403C-85B7-D5C5C2143CBE},{2FC22FA3-CCC7-41F9-8137-F907D888C999},{8E24E3EF-F350-4D29-8BA0-430E425F54E0},{ECEAF686-3552-44BD-A49B-941376AE4109},{2AE2C457-96FE-4379-BC33-BA048E4C06B8},{3A4F4557-A6E8-4382-A69F-59CF8895645F},{33F9C053-E51D-4738-91CD-45ABB82C1D8A},{2FC22FA3-CCC7-41F9-8137-F907DC9C1F24},{3012D11D-199C-4D46-8B58-6704EEF4A3EF},{521C56BD-EC92-4AF5-BE8C-229391C37673},{8F0D8C83-05F9-4148-96E9-3D015143063F},{3B0BABA9-EF20-47A0-A026-08DA34B9A7F7},{D733E4E8-0418-4B99-BCF5-8FC5CB9C5C42},{1B136CF3-83CF-471D-925A-EEB72BC6CD5B},{1EBA180D-7657-4E72-A678-9ECE4EDD58C1},{46A1E43B-E1AE-419A-AB44-6EECF93D3C7F},{F68941D5-DD4C-443F-90E3-39F5D16BFD13},{F9A6AADA-7DDD-4776-836A-A3EE4032D957},{7E9A8B94-3C6E-4597-9B09-FCABD40BB155},{6D57662F-7DD0-41E1-806B-3562412FDFAF},{0D620DAB-1B89-4E7B-BB6A-29EB77F9AEE9},{7C7058BE-F586-44C4-B5BE-47D2E05E96BD},{AF173F9A-4724-405B-AA9C-C72E1DCA7647},{2D84D273-5F30-4D17-A8C1-6308862FBBE7},{F5EB2A97-423D-4D9A-A338-67A72AC26F77},{6A200BD7-1A53-40FC-9DBB-44499F65B74C},{2F4D0C04-FD15-4321-A5E3-5AA2FCB0D70E},{9850A686-F991-4F36-8EF2-C0F043103276},{ACCDE453-D50A-48E0-9BFB-1BEA45D6D16E}, {8FF3157E-1099-4256-A801-51DE178950AF}, {F32E62BE-7DC2-433B-B724-394CEFA4D076},{26A64828-4E82-4398-84E1-11F9F092FDD8},{4AEC67F9-80C1-427F-904E-489920B71940},{9EB619A3-61E8-4C1A-9573-27C87DABEF76}"

'rmanyushin 25.08.2009 Start
'Скрываем справочники Seznam Adresy STS, Recipient lists STS, Списки получателей СТС RU от пользователей. При переносе с сервера на сервер уточнить GUID справочников. 
   
   'GUID справочников на тестовом сервере Paydox gl-test-02.global.sitronics.com 
   'CurrentProhibitedDirectoryGUIDs = CurrentProhibitedDirectoryGUIDs & ", {8645C6EE-448E-401C-9258-4F1513489620},{A3AAA282-3B99-4302-91D6-FCC5F92C2D45},{9741BDA3-3926-45BD-8733-0DB5EBD9E546}"   
   
   'GUID справочников на основном сервере Paydox gl-paydox-01.global.sitronics.com 
   CurrentProhibitedDirectoryGUIDs = CurrentProhibitedDirectoryGUIDs & ", {16F81D6B-DAAE-4622-9D68-37C6FE44E1FC},{53922108-81D9-4E8C-AA50-85A88319B04C},{0E026D46-852B-4B08-BA1B-8D1E1A087906}"
'rmanyushin 25.08.2009 End
  
  
'rmanyushin 60298 02.11.2009 Start
'Скрываем справочники Contracts/Contracting party, Contracts/Smluvní strana, 
'Contracts/Сторона Договора от пользователей.
'При переносе с сервера на сервер уточнить GUID справочников. 
'GUID справочников на тестовом сервере Paydox gl-test-02.global.sitronics.com:8888
   'CurrentProhibitedDirectoryGUIDs = CurrentProhibitedDirectoryGUIDs & ", {913547B9-1703-4241-8E5C-CE87656582DF}, {C2C0CC6F-6C48-41DC-9D91-C8BF2F676E5D}, {02073D43-0553-45B6-8F50-50C396DB2E14}"   
   
   'GUID справочников на основном сервере Paydox gl-paydox-01.global.sitronics.com 
   CurrentProhibitedDirectoryGUIDs = CurrentProhibitedDirectoryGUIDs & ", {260405A7-E08E-4012-B353-9447FAB64683}, {EB4C8816-901B-405C-867E-61420B4E0C30}, {63E943E3-E678-4138-900B-1E5FE96809AE}"
'rmanyushin 60298 02.11.2009 End
  

'rmanyushin 61481 09.11.2009 Start
'Скрываем справочники Distribution list STS, Seznam adresátů STS от пользователей.
'При переносе с сервера на сервер уточнить GUID справочников. 
'GUID справочников на тестовом сервере Paydox gl-test-02.global.sitronics.com:8888
   'CurrentProhibitedDirectoryGUIDs = CurrentProhibitedDirectoryGUIDs & ", {54CD6AD0-4A8E-4ADB-B902-4C4014D8E47C}, {DFB730C4-A2B6-47E9-B726-9B397ACA0B0B}"   

    'GUID справочников на основном сервере Paydox gl-paydox-01.global.sitronics.com 
   CurrentProhibitedDirectoryGUIDs = CurrentProhibitedDirectoryGUIDs & ", {FD1076B3-6CC4-474D-BB50-516C051A805F}, {582866C7-66B4-4D57-ADC2-30827DA2E54D}"
'rmanyushin 61481 09.11.2009 End

'rmanyushin 119579 19.08.2010 Start
 'Скрываем справочники Заместители ГД СТС и Директора направлений СТС от пользователей
 'GUID справочников на основном сервере Paydox it-test-08.sts.sitronics.com 
   'CurrentProhibitedDirectoryGUIDs = CurrentProhibitedDirectoryGUIDs & ", {7FC16BCD-7425-4BFC-9903-FC94259C8957}, {A8FDC863-2DAA-41B0-AA13-325456DC8237}"
    'GUID справочников на основном сервере Paydox gl-paydox-01.global.sitronics.com 
   CurrentProhibitedDirectoryGUIDs = CurrentProhibitedDirectoryGUIDs & ", {0C26E186-FED8-408A-93B4-3F1739854E92}, {DD51787D-32CB-43E6-8043-B5843E34996A}"
'rmanyushin 119579 19.08.2010 End


'rmanyushin 27.12.2010 Start
'Скрываем от пользователей справочники: STS_Rules, Заявка на оплату УК/*
	CurrentProhibitedDirectoryGUIDs = CurrentProhibitedDirectoryGUIDs & ",{365C2A1C-D404-47AF-AC76-9421A36E8E6A} , {42FD049F-2CEC-44A2-A8F2-B05EB1183F25}, {507B2058-B7C4-40B2-8EC4-D75A8E4CE28D}, {B53443BE-B30E-4C52-98D7-5ED0BEE67080}, {C62278E8-50EB-41AF-89C4-90717DF68414}, {6D6236CD-DA05-4B52-87F1-6C657F2544EE} ,{E2D63CEF-6DC4-44F3-B4BF-BD4B0D9438D0},	{15EB5243-22D8-425D-B31A-9CBA4396FCFC}" 
'rmanyushin 27.12.2010 End
  
  'Даем уполномоченным доступ к справочнику проектов STS
  If CanLoadSTSProjectList(Session("UserID")) Then
    CurrentProhibitedDirectoryGUIDs = Replace(CurrentProhibitedDirectoryGUIDs+",", "{2AE2C457-96FE-4379-BC33-BA048E4C06B8},", "")
  End If
'Запрос №1 - СИБ - start
  'Скрываем от пользователей новые справочники для СИБ (роли, отправители и списки согласующих)
  CurrentProhibitedDirectoryGUIDs = CurrentProhibitedDirectoryGUIDs & ", {95F98947-C5E4-452B-B8B1-238D887CC168}, {D0E4DEEB-A1A8-48B0-B527-9573856769DB}, {6A9F208C-D0B1-4172-8186-C8235E7FB9B9}" 
'Запрос №1 - СИБ - end
'Запрос №43 - СТС - start
  'Скрываем от пользователей справочник Проверенные контрагенты
  CurrentProhibitedDirectoryGUIDs = CurrentProhibitedDirectoryGUIDs & ", {F5E696BF-135A-4C81-9920-7F444D51CE14}" 
'Запрос №43 - СТС - end
End If
'Ph - 20090128 - end

'Формула расчета курса валюты (для заявок СТС)
'ph - 20101108 - start
'STS_CurrencyRateFormula = "CStr(CCur(dsDoc(""UserFieldMoney1"")/dsDoc(""AmountDoc""))) + ""  ("" + CStr(CCur(dsDoc(""AmountDoc"")/dsDoc(""UserFieldMoney1""))) + "")"""
STS_CurrencyRateFormula = "ShowCurrencyRate(dsDoc(""UserFieldMoney1""), dsDoc(""AmountDoc""))"
'ph - 20101108 - end

'VAR_BPTypes = ""

'rmanyushin 63828 23.11.2009
VAR_BPTypes = "@FORM"
VAR_CanMakeDocCompleted="y"

If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_PURCHASE_ORDER)) > 0 Then
    VAR_AddFieldsNames = DOCS_AdditionalUsers+VbCrLf+SIT_BusinessUnit+VbCrLf+SIT_ContractType
Elseif InStr(UCase(Session("CurrentClassDoc")), UCase(RTI_PURCHASE_ORDER)) > 0 Then
    VAR_AddFieldsNames = DOCS_AdditionalUsers
      Elseif InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_PURCHASE_ORDER)) > 0 Then
     VAR_AddFieldsNames = DOCS_AdditionalUsers
  ElseIf InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_BSAP)) > 0 or _
         InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_RL_PROTOCOL)) > 0 Then
     VAR_AddFieldsNames = DOCS_AdditionalUsers + VbCrLf + SIT_ContractType
  ElseIf InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_CONTRACT)) > 0 or _
         InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_S_CONTRACT)) > 0 Then
     VAR_AddFieldsNames = DOCS_AdditionalUsers + VbCrLf + MS_CONTRACT_PARTY + VbCrLf + _
                          MS_CONTRACT_INTERNAL + VbCrLf + MS_CONTRACT_DEPEND + VbCrLf + "Тип сделки"
  Else
     'Названия дополнительных полей
     'rmanyushin 60298 02.11.2009 ' Указываем название дополнительного поля ContractType (STS_ContractType ) в форме.
     'ph - 20120216 - start
     'VAR_AddFieldsNames = DOCS_AdditionalUsers+VbCrLf+SIT_BusinessUnit+VbCrLf+STS_ContractType
     VAR_AddFieldsNames = DOCS_AdditionalUsers+VbCrLf+SIT_BusinessUnit+VbCrLf+STS_ContractType+VbCrLf+STS_OffshoreZone+VbCrLf+"Взаимозависимость сторон договора"
     'ph - 20120216 - end
  End If

'Ph - 20090303 - Меняем основную валюту на доллар во всех интерфейсах
Var_MainSystemCurrency="USD" 'Main currency
Var_MainSystemCurrencyName="US dollar" 'Main currency name

' Необходимо дать право Регистраторам (те, кто входят в роль регистратора) в 
' распорядительных, нормативных и исходящих править поле получатели даже у подписанных
' документов
' DmGorsky_5: меняем местами проверку: вначале ищем SIT_SITRU ("СИТРОНИКС ИТ"), потом,
' если не найдено, - SIT_SITRONICS ("СИТРОНИКС")
If InStr(UCase(Session("Department")),UCase(SIT_SITRU)) = 1 then ' DmGorsky
  sDepartment = SIT_SITRU ' DmGorsky
ElseIf InStr(UCase(Session("Department")),UCase(SIT_SITRONICS)) = 1 then
  sDepartment = SIT_SITRONICS
Else
  sDepartment = SIT_STS
End If

If InStr(UCase(ReplaceRoleFromDir(SIT_Registrar, sDepartment)), "<"+UCase(Session("UserID"))+">") > 0 Then
'  If InStr(Session("CurrentClassDoc"), SIT_NORM_DOCS) > 0 or InStr(Session("CurrentClassDoc"), SIT_ISHODYASCHIE) > 0 or InStr(Session("CurrentClassDoc"), SIT_RASP_DOCS) > 0 Then
    VAR_WriteAccess="Y"
'  End If
End If

'Не показываем кнопку Календарь в пользователях
If UCase(Request.ServerVariables("URL")) = UCase("/ShowUser.asp") Then
  ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickCalendarEvent"
End If

'Для заявок на закупку и оплату выводить информацию по уточнению роли согласующих
SIT_ShowOrdersAgreesDescription = "Y"

'Для отслеживания откуда вызывается справочник и установки фильтра на SQL справочника BusinessUnits
If UCase(Request.ServerVariables("URL")) <> UCase("/ListDirectories.asp") Then
  Session("CurrentPage") = Request.ServerVariables("URL")
End If

'SAY 2009-03-19 Файлы
If UCase(Request.ServerVariables("URL"))=UCase("/UploadFileNew.asp") Then

  sNameCreation = ""
  sNameAproval = ""
  sNameAproved = "" 
  sListToReconcile = ""
  sListReconciled = ""
  sLocationPath = ""
  sIsActive = ""
  GetDocField_test Request("DocID")

  'у создателя по-умолчанию отмечена галочка "основная версия"
  If InStr(sNameCreation, "<"+Session("UserID")+">") > 0 Then
    VAR_OnlyMainVersionFiles=""
    VAR_UploadMainVersionFileByDefault="Y"
    '
  End If

  'согласующий и подписант могут загружать только версии
  'If InStr(sNameAproval, "<"+Session("UserID")+">") > 0 or InStr(sListToReconcile, "<"+Session("UserID")+">") > 0 Then
  '  VAR_CanCreateMainVersionFiles=False
  'End If

  'out "sListToReconcile="+sListToReconcile
  'out "sListReconciled="+sListReconciled

  ' если это утверждающий или согласующий в момент согласования
  'If InStr(sNameAproval, "<"+Session("UserID")+">") > 0 or (InStr(sListToReconcile, "<"+Session("UserID")+">") > 0 and InStr(sListReconciled, "<"+Session("UserID")+">") = 0) Then
  'If  InStr(sNameAproval, "<"+Session("UserID")+">") > 0 or InStr(sListToReconcile, "<"+Session("UserID")+">") > 0 Then
  'vnik_protocols
  If InStr(UCase(Session("CurrentClassDoc")), Trim(UCase("Протоколы*Protocols*Protokoly/Встреч*Meetings*Schůze"))) > 0 Then
'ph - 20111014 - start
  'ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(STS_SLUZH_ZAPISKA_OVERTIME_PLAN)) = 1 Then
  ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(STS_SLUZH_ZAPISKA_OVERTIME_PLAN)) = 1 or InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) = 1 Then
		If UCase(sIsActive)="Y" and InStr(sNameCreation, "<"+Session("UserID")+">") = 0 and (InStr(sNameAproval, "<"+Session("UserID")+">") > 0 or InStr(sListToReconcile, "<"+Session("UserID")+">") > 0) Then
			VAR_OnlyMainVersionFiles=""
			VAR_CanCreateMainVersionFiles = False
			VAR_OnlyVersionFiles = "Y"
			VAR_CanUpdateDocWithoutChangingStatuses = ""
			If Trim(Request("FileToChangeMainVersion"))<>"" Then
				VAR_CanCreateMainVersionFiles = True
				VAR_OnlyMainVersionFiles="Y"
				VAR_OnlyVersionFiles = ""
				VAR_CanUpdateDocWithoutChangingStatuses = "Y"
			End If
		End If
'ph - 20111014 - end
  Else
  'vnik_protocols
  If UCase(sIsActive)="Y" and (InStr(sNameAproval, "<"+Session("UserID")+">") > 0 or InStr(sListToReconcile, "<"+Session("UserID")+">") > 0) Then
    VAR_OnlyMainVersionFiles=""
    VAR_CanCreateMainVersionFiles=False
    VAR_OnlyVersionFiles = "Y"
    VAR_CanUpdateDocWithoutChangingStatuses = ""
    If Trim(Request("FileToChangeMainVersion"))<>"" Then
      VAR_CanCreateMainVersionFiles=True
      VAR_OnlyMainVersionFiles="Y"
      VAR_OnlyVersionFiles = ""
      VAR_CanUpdateDocWithoutChangingStatuses = "Y"
      'out "version change"
    End If
    'out "aproval or reconcile"
  End If
  'vnik_protocols
  End If
  'vnik_protocols

  ' регистратор может загружать только основные версии
  If InStr(sLocationPath, "<"+Session("UserID")+">") > 0 and ((Trim(sNameAproval) <> "" and Trim(sNameAproved) <> "") or Trim(sNameAproval) = "") Then
    VAR_OnlyVersionFiles = ""
    VAR_OnlyMainVersionFiles="Y"
    VAR_CanCreateMainVersionFiles=True
    VAR_AgreeAgainInitiallyNotChecked="Y"
    VAR_UploadFileForcesAgreeAgain="N"
    VAR_CanUpdateDocWithoutChangingStatuses = "Y"
    'out "registrator"
  End If

' включаем параметр чтобы в поручениях и входящих появилась возможность загрузить несколько 
' основных версий. Галочка "Основная версия" будет вегда стоять ON
  If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ZADACHI)) > 0 Or InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_VHODYASCHIE)) > 0 Then
    VAR_OnlyMainVersionFiles="Y"
    VAR_CanCreateMainVersionFiles=True
  End If

'out "VAR_OnlyMainVersionFiles="+VAR_OnlyMainVersionFiles
'out "VAR_CanCreateMainVersionFiles=" + CStr(VAR_CanCreateMainVersionFiles)
 
End If
'SAY 2009-03-19 Файлы

If UCase(Request.ServerVariables("URL"))=UCase("/ShowDoc.asp") Then
  sNameAproval = ""
  GetDocField_test Request("DocID")
  'подписант имеет право назначить основную версию из существующих загруженных
  If InStr(sNameAproval, "<"+Session("UserID")+">") > 0 Then
    VAR_CanChangeFileVersion="Y"
    'VAR_UploadFileForcesAgreeAgain="N"
  End If
End If


'rmanyushin 51555 16.09.2009 start
'Разрешить привилегированными пользователями СТС просмотр карточки документа, но доступ к 
'прикрепленным файлам разрешить только "STS_Auditor" и "STS_HeadOf789"
If isPrivilegedUserSTS() Then
	If UCase(Session("UserID")) = UCase(STS_Overseer) Then
		VAR_ReadAccess="Y"
		Session("SIT_UserCanDownloadFiles") = False
	End If

	If UCase(Session("UserID")) = UCase(STS_Auditor) Then
		VAR_ReadAccess="Y"
		'Доступ к прикрепленным файлам только из карточки документа
		If InStr(UCase(Request.ServerVariables("URL")),UCase("/ShowDoc.asp"))=1 Then
			Session("SIT_UserCanDownloadFiles") = True
		End If
	End If

    'rmanyushin 56781 13.10.2009 start
    If UCase(Session("UserID")) = UCase(STS_HeadOf789) Then
		VAR_ReadAccess="Y"
		'Доступ к прикрепленным файлам только из карточки документа
		If InStr(UCase(Request.ServerVariables("URL")),UCase("/ShowDoc.asp"))=1 Then
			Session("SIT_UserCanDownloadFiles") = True
		End If
	End If
	'rmanyushin 56781 13.10.2009 end
		
	'rmanyushin 133266 05.10.2010 Start
	If UCase(Session("UserID")) = UCase(STS_LegalSTS) Then
		VAR_ReadAccess="Y"
		Session("SIT_UserCanDownloadFiles") = True
	End If
	'rmanyushin 133266 05.10.2010 End
	
	'rmanyushin 79501 24.02.2010 Start
    If UCase(Session("UserID")) = UCase(STS_POViewer) Then
	    If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) > 0 or InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) > 0 or InStr(UCase(Request.ServerVariables("URL")),UCase("/GetReport.asp")) = 1 Then
			VAR_ReadAccess="Y"
			'Доступ к прикрепленным файлам только из карточки документа
			If InStr(UCase(Request.ServerVariables("URL")),UCase("/ShowDoc.asp"))=1 Then
				Session("SIT_UserCanDownloadFiles") = True
			End If
		Else
			VAR_ReadAccess=""
		End If
	End If
    'rmanyushin 79501 24.02.2010 End

End If
'rmanyushin 51555 16.09.2009 end



VAR_UseDepartmentDependantsForResponsiblesList = "Y"
VAR_FirstSymSearchNoLogin="Y"
VAR_CanSendEMailsToUnregisteredUsers="y"

'ph - 20100319 - start - Старый интерфейс в документах
bDocsShortStyleScreen = False
'ph - 20100319 - end
'ph - 20100414 - start - вставка из справочника пользователей в формате Фамилия И.О.
VAR_SurnameGN="2"
'ph - 20100414 - end


'ph - 20101020 - start

'Запрос №44 - СТС - start
STS_JurChief = "eivanova"
'Запрос №44 - СТС - end

If InStr(UCase(Request.ServerVariables("URL")), UCase("/ListDoc.asp")) > 0 and UCase(Request("VisaDocs")) = "Y" Then
'Запрос №44 - СТС - start
	'VAR_ListDocSQL = "select Comments.*, Docs.*, Comments.DateCreation as CommentsDateCreation, Docs.DateCreation as DocsDateCreation, Comments.FileName as CommentsFileName  from Docs  Left Outer Join Comments ON (Docs.DocID = Comments.DocID  And (SpecialInfo='VISAWAITING')) where UserID = N'" & Session("UserID") & "' and (StatusCompletion is NULL or (StatusCompletion<>'1' and StatusCompletion<>'0')) and (IsActive<>'N' or IsActive is Null) order by DateEventEnd, Docs.DateCreation desc, Docs.DocID desc, Comments.DateCreation"
	If InStr(Session("Department"), SIT_STS_ROOT_DEPARTMENT) = 1 Then 'Для СТС персонально подменяются запросы
		If Request("UserIDToSee") = "" Then 'запрос для списка согласование
			'VAR_ListDocSQL = "select Comments.*, Docs.*, Comments.DateCreation as CommentsDateCreation, Docs.DateCreation as DocsDateCreation, Comments.FileName as CommentsFileName  from Docs  Left Outer Join Comments ON (Docs.DocID = Comments.DocID  And (SpecialInfo='VISAWAITING')) where "
			'VAR_ListDocSQL = VAR_ListDocSQL & " (UserID = N'" & Session("UserID") & "' or CharIndex(N'<" & Session("UserID") & ">', ListToReconcile) > 0 and (CharIndex(N'#!', ListToReconcile) > 0 or CharIndex(N'-<', ListReconciled) > 0))"
			'VAR_ListDocSQL = VAR_ListDocSQL & " and (StatusCompletion is NULL or (StatusCompletion<>'1' and StatusCompletion<>'0')) and (IsActive<>'N' or IsActive is Null) order by DateEventEnd, Docs.DateCreation desc, Docs.DocID desc, Comments.DateCreation"
			VAR_ListDocSQLPrefix = "select Comments.*, Docs.*, Comments.DateCreation as CommentsDateCreation, Docs.DateCreation as DocsDateCreation, Comments.FileName as CommentsFileName  from Docs  Left Outer Join Comments ON (Docs.DocID = Comments.DocID  And (SpecialInfo='VISAWAITING')) where "
'			VAR_ListDocSQLPrefix = VAR_ListDocSQLPrefix & " (UserID = N'" & Session("UserID") & "' or CharIndex(N'<" & Session("UserID") & ">', ListToReconcile) > 0 and CharIndex(N'" & SIT_DOGOVORI_NEW & "', ClassDoc) = 1 and (CharIndex(N'#!', ListToReconcile) > 0 or CharIndex(N'-<', ListReconciled) > 0))"
			VAR_ListDocSQLPrefix = VAR_ListDocSQLPrefix & " (UserID = N'" & Session("UserID") & "' or "
			If Session("UserID") <> STS_JurChief Then
				VAR_ListDocSQLPrefix = VAR_ListDocSQLPrefix & " (CharIndex(N'<" & Session("UserID") & ">', NameResponsible) > 0 or CharIndex(N'<" & Session("UserID") & ">', NameCreation) > 0) and "
			End If
			VAR_ListDocSQLPrefix = VAR_ListDocSQLPrefix & " CharIndex(N'" & SIT_DOGOVORI_NEW & "', ClassDoc) = 1 and (CharIndex(N'#!', ListToReconcile) > 0 or CharIndex(N'-<', ListReconciled) > 0))"
			VAR_ListDocSQLPrefix = VAR_ListDocSQLPrefix & " and (StatusCompletion is NULL or (StatusCompletion<>'1' and StatusCompletion<>'0')) and (IsActive<>'N' or IsActive is Null) "
			VAR_ListDocSQLSuffix = " order by DateEventEnd, Docs.DateCreation desc, Docs.DocID desc, Comments.DateCreation"
		Else 'запрос для списка В работу
			VAR_ListDocSQL = "select Case When CommentType='VISA' And (SpecialInfo='VISASUSPENDING' or SpecialInfo='VISAOKREFUSE') Then '' Else Comments.Subject End as Subject, "
			VAR_ListDocSQL = VAR_ListDocSQL & "Case When CommentType='VISA' And (SpecialInfo='VISASUSPENDING' or SpecialInfo='VISAOKREFUSE') Then '' Else Comments.Comment End as Comment, "
			VAR_ListDocSQL = VAR_ListDocSQL & "Case When CommentType='VISA' And (SpecialInfo='VISASUSPENDING' or SpecialInfo='VISAOKREFUSE') Then '' Else Comments.Version End as Version, "
			VAR_ListDocSQL = VAR_ListDocSQL & "Comments.UserID,Comments.UserName,Comments.ContactUserName,Comments.ContactUserID,Comments.PartnerName,Comments.DocID,Comments.DateCreation,Comments.DateEvent,Comments.DateEventEnd,Comments.Address,Comments.CommentType,Comments.SpecialInfo,Comments.FileName,Comments.FileNameOriginal,Comments.KeyField,Comments.PartnerID,Comments.GUID,Comments.GUIDPartner,Comments.GUIDUser,Comments.GUIDDoc,Comments.GUIDInstanceItem,Comments.Amount,Comments.Statuses,"
			VAR_ListDocSQL = VAR_ListDocSQL & "Docs.*, BoardOrder =  CASE SUBSTRING(ISNULL(Comments.SpecialInfo, '    '), 1, 4) WHEN 'RESP' THEN SUBSTRING(Comments.SpecialInfo, 5, LEN(Comments.SpecialInfo)-4) ELSE CONVERT(varchar(24), ISNULL(Comments.DateCreation, GetDate()), 120)+' '+(CONVERT(varchar(12), ISNULL(Comments.KeyField, '            '))) END , Comments.DateCreation as CommentsDateCreation, Docs.DateCreation as DocsDateCreation, Case When CommentType='VISA' And (SpecialInfo='VISASUSPENDING' or SpecialInfo='VISAOKREFUSE') Then '' Else Comments.FileName End as CommentsFileName "
			VAR_ListDocSQL = VAR_ListDocSQL & "from Docs  Left Outer Join Comments "
			VAR_ListDocSQL = VAR_ListDocSQL & "ON (Docs.DocID = Comments.DocID  And (Comments.CommentType='VIEWED' And Comments.UserID = '" & Session("UserID") & "' Or SpecialInfo='VISAWAITING' Or (CommentType='REVIEW' And Address='REQUEST')  Or (CommentType='VISA' And (SpecialInfo='REQUESTTOMODIFY' or (SpecialInfo='VISASUSPENDING' or SpecialInfo='VISAOKREFUSE') and "

'			VAR_ListDocSQL = VAR_ListDocSQL & " (CharIndex(N'" & SIT_DOGOVORI_NEW & "', ClassDoc) = 1 and CharIndex(N'<" & Session("UserID") & ">', ListToReconcile) > 0 and (CharIndex(N'#!', ListToReconcile) > 0 or CharIndex(N'-<', ListReconciled) > 0)) "
			If Session("UserID") <> STS_JurChief Then
				VAR_ListDocSQL = VAR_ListDocSQL & " (CharIndex(N'<" & Session("UserID") & ">', NameResponsible) > 0 or CharIndex(N'<" & Session("UserID") & ">', NameCreation) > 0) and "
			End If
			VAR_ListDocSQL = VAR_ListDocSQL & " CharIndex(N'" & SIT_DOGOVORI_NEW & "', ClassDoc) = 1 and (CharIndex(N'#!', ListToReconcile) > 0 or CharIndex(N'-<', ListReconciled) > 0)"

			VAR_ListDocSQL = VAR_ListDocSQL & ")))) "
			VAR_ListDocSQL = VAR_ListDocSQL & "where (IsActive<>'N' And ((((NameAproval='' or NameAproval is Null) or (NameApproved<>'' and Not (NameApproved is Null))) and (CHARINDEX(N'<" & Session("UserID") & ">', ListToView)>0 or CHARINDEX(N'<USERS:*>', ListToView)>0  or " & CheckDepartmentInList_SQL("ListToView", Session("Department")) & ")"
			VAR_ListDocSQL = VAR_ListDocSQL & " and CHARINDEX(N'<" & Session("UserID") & ">-', ListToView)<=0 and (StatusCompletion is NULL or StatusCompletion <> '0') and CHARINDEX('-<', ISNULL(NameApproved, ''))<=0 ))) "
			VAR_ListDocSQL = VAR_ListDocSQL & " or (((( (CHARINDEX(N'<" & Session("UserID") & ">', ListToReconcile)>0 and CHARINDEX(N'<" & Session("UserID") & ">', ListReconciled)<=0)"

			VAR_ListDocSQL = VAR_ListDocSQL & " or ("
'			VAR_ListDocSQL = VAR_ListDocSQL & "CharIndex(N'" & SIT_DOGOVORI_NEW & "', ClassDoc) = 1 and CharIndex(N'<" & Session("UserID") & ">', ListToReconcile) > 0 and (CharIndex(N'#!', ListToReconcile) > 0 or CharIndex(N'-<', ListReconciled) > 0))"
			If Session("UserID") <> STS_JurChief Then
				VAR_ListDocSQL = VAR_ListDocSQL & " (CharIndex(N'<" & Session("UserID") & ">', NameResponsible) > 0 or CharIndex(N'<" & Session("UserID") & ">', NameCreation) > 0) and "
			End If
			VAR_ListDocSQL = VAR_ListDocSQL & " CharIndex(N'" & SIT_DOGOVORI_NEW & "', ClassDoc) = 1 and (CharIndex(N'#!', ListToReconcile) > 0 or CharIndex(N'-<', ListReconciled) > 0))"

			VAR_ListDocSQL = VAR_ListDocSQL & " or (CHARINDEX(N'<" & Session("UserID") & ">', NameAproval)>0 and CHARINDEX(N'<" & Session("UserID") & ">', NameApproved)<=0) "
			VAR_ListDocSQL = VAR_ListDocSQL & " or CHARINDEX(N'<" & Session("UserID") & ">', FileNameNameLastModification)>0 "
			VAR_ListDocSQL = VAR_ListDocSQL & " or PATINDEX(N'%<" & Session("UserID") & ">%', NameResponsible)>0 "
			VAR_ListDocSQL = VAR_ListDocSQL & " or CHARINDEX(N'<" & Session("UserID") & ">', ISNULL(Resolution, ''))>0 "
			VAR_ListDocSQL = VAR_ListDocSQL & " or (CHARINDEX(N'<" & Session("UserID") & ">', ISNULL(NameControl, ''))>0 and StatusCompletion='+') "
			VAR_ListDocSQL = VAR_ListDocSQL & " Or CHARINDEX(N'<" & Session("UserID") & ">', Address)>0 "
			VAR_ListDocSQL = VAR_ListDocSQL & " Or CHARINDEX(N' <" & Session("UserID") & ">', Comment)>0 "
			VAR_ListDocSQL = VAR_ListDocSQL & " or (CHARINDEX(N'<" & Session("UserID") & ">', ISNULL(LocationPath, ''))>0 And CHARINDEX(N'>+', ISNULL(LocationPath, ''))<=0 And (NameAproval is Null or NameApproved<>'')) "
			VAR_ListDocSQL = VAR_ListDocSQL & " or (CHARINDEX(N'<" & Session("UserID") & ">?', ISNULL(NameControl, ''))>0 And (ListReconciled is Null or ListReconciled=''))) "
			VAR_ListDocSQL = VAR_ListDocSQL & " and ((StatusCompletion is NULL and CHARINDEX(N'-<', ISNULL(NameApproved, ''))<=0) or (StatusCompletion<>'1' and StatusCompletion<>'0'))) and (IsActive<>'N' or IsActive is Null))) "
			VAR_ListDocSQL = VAR_ListDocSQL & " order by Docs.DateCreation desc, Docs.DocID desc, BoardOrder, Comments.DateCreation"
			' VAR_ListDocSQL = "select Case When CommentType='VISA' And (SpecialInfo='VISASUSPENDING' or SpecialInfo='VISAOKREFUSE') Then '' Else Comments.Subject End as Subject, "
			' VAR_ListDocSQL = VAR_ListDocSQL & "Case When CommentType='VISA' And (SpecialInfo='VISASUSPENDING' or SpecialInfo='VISAOKREFUSE') Then '' Else Comments.Comment End as Comment, "
			' VAR_ListDocSQL = VAR_ListDocSQL & "Case When CommentType='VISA' And (SpecialInfo='VISASUSPENDING' or SpecialInfo='VISAOKREFUSE') Then '' Else Comments.Version End as Version, "
			' VAR_ListDocSQL = VAR_ListDocSQL & "Comments.UserID,Comments.UserName,Comments.ContactUserName,Comments.ContactUserID,Comments.PartnerName,Comments.DocID,Comments.DateCreation,Comments.DateEvent,Comments.DateEventEnd,Comments.Address,Comments.CommentType,Comments.SpecialInfo,Comments.FileName,Comments.FileNameOriginal,Comments.KeyField,Comments.PartnerID,Comments.GUID,Comments.GUIDPartner,Comments.GUIDUser,Comments.GUIDDoc,Comments.GUIDInstanceItem,Comments.Amount,Comments.Statuses,"
			' VAR_ListDocSQL = VAR_ListDocSQL & "Docs.*, BoardOrder =  CASE SUBSTRING(ISNULL(Comments.SpecialInfo, '    '), 1, 4) WHEN 'RESP' THEN SUBSTRING(Comments.SpecialInfo, 5, LEN(Comments.SpecialInfo)-4) ELSE CONVERT(varchar(24), ISNULL(Comments.DateCreation, GetDate()), 120)+' '+(CONVERT(varchar(12), ISNULL(Comments.KeyField, '            '))) END , Comments.DateCreation as CommentsDateCreation, Docs.DateCreation as DocsDateCreation, Case When CommentType='VISA' And (SpecialInfo='VISASUSPENDING' or SpecialInfo='VISAOKREFUSE') Then '' Else Comments.FileName End as CommentsFileName "
			' VAR_ListDocSQL = VAR_ListDocSQL & "from Docs  Left Outer Join Comments "
			' VAR_ListDocSQL = VAR_ListDocSQL & "ON (Docs.DocID = Comments.DocID  And (Comments.CommentType='VIEWED' And Comments.UserID = '" & Session("UserID") & "' Or SpecialInfo='VISAWAITING' Or (CommentType='REVIEW' And Address='REQUEST')  Or (CommentType='VISA' And (SpecialInfo='REQUESTTOMODIFY' or (SpecialInfo='VISASUSPENDING' or SpecialInfo='VISAOKREFUSE') and (CharIndex(N'" & SIT_DOGOVORI_NEW & "', ClassDoc) = 1 and CharIndex(N'<" & Session("UserID") & ">', ListToReconcile) > 0 and (CharIndex(N'#!', ListToReconcile) > 0 or CharIndex(N'-<', ListReconciled) > 0)))))) "
			' VAR_ListDocSQL = VAR_ListDocSQL & "where (IsActive<>'N' And ((((NameAproval='' or NameAproval is Null) or (NameApproved<>'' and Not (NameApproved is Null))) and (CHARINDEX(N'<" & Session("UserID") & ">', ListToView)>0 or CHARINDEX(N'<USERS:*>', ListToView)>0  or " & CheckDepartmentInList_SQL("ListToView", Session("Department")) & ")"
			' VAR_ListDocSQL = VAR_ListDocSQL & " and CHARINDEX(N'<" & Session("UserID") & ">-', ListToView)<=0 and (StatusCompletion is NULL or StatusCompletion <> '0') and CHARINDEX('-<', ISNULL(NameApproved, ''))<=0 ))) "
			' VAR_ListDocSQL = VAR_ListDocSQL & " or (((( (CHARINDEX(N'<" & Session("UserID") & ">', ListToReconcile)>0 and CHARINDEX(N'<" & Session("UserID") & ">', ListReconciled)<=0)"
			' VAR_ListDocSQL = VAR_ListDocSQL & " or (CharIndex(N'" & SIT_DOGOVORI_NEW & "', ClassDoc) = 1 and CharIndex(N'<" & Session("UserID") & ">', ListToReconcile) > 0 and (CharIndex(N'#!', ListToReconcile) > 0 or CharIndex(N'-<', ListReconciled) > 0))"
			' VAR_ListDocSQL = VAR_ListDocSQL & " or (CHARINDEX(N'<" & Session("UserID") & ">', NameAproval)>0 and CHARINDEX(N'<" & Session("UserID") & ">', NameApproved)<=0) "
			' VAR_ListDocSQL = VAR_ListDocSQL & " or CHARINDEX(N'<" & Session("UserID") & ">', FileNameNameLastModification)>0 "
			' VAR_ListDocSQL = VAR_ListDocSQL & " or PATINDEX(N'%<" & Session("UserID") & ">%', NameResponsible)>0 "
			' VAR_ListDocSQL = VAR_ListDocSQL & " or CHARINDEX(N'<" & Session("UserID") & ">', ISNULL(Resolution, ''))>0 "
			' VAR_ListDocSQL = VAR_ListDocSQL & " or (CHARINDEX(N'<" & Session("UserID") & ">', ISNULL(NameControl, ''))>0 and StatusCompletion='+') "
			' VAR_ListDocSQL = VAR_ListDocSQL & " Or CHARINDEX(N'<" & Session("UserID") & ">', Address)>0 "
			' VAR_ListDocSQL = VAR_ListDocSQL & " Or CHARINDEX(N' <" & Session("UserID") & ">', Comment)>0 "
			' VAR_ListDocSQL = VAR_ListDocSQL & " or (CHARINDEX(N'<" & Session("UserID") & ">', ISNULL(LocationPath, ''))>0 And CHARINDEX(N'>+', ISNULL(LocationPath, ''))<=0 And (NameAproval is Null or NameApproved<>'')) "
			' VAR_ListDocSQL = VAR_ListDocSQL & " or (CHARINDEX(N'<" & Session("UserID") & ">?', ISNULL(NameControl, ''))>0 And (ListReconciled is Null or ListReconciled=''))) "
			' VAR_ListDocSQL = VAR_ListDocSQL & " and ((StatusCompletion is NULL and CHARINDEX(N'-<', ISNULL(NameApproved, ''))<=0) or (StatusCompletion<>'1' and StatusCompletion<>'0'))) and (IsActive<>'N' or IsActive is Null))) "
			' VAR_ListDocSQL = VAR_ListDocSQL & " order by Docs.DateCreation desc, Docs.DocID desc, BoardOrder, Comments.DateCreation"
		End If
	Else
		If Request("UserIDToSee") = "" Then 'запрос для списка согласование
			VAR_ListDocSQLPrefix = "select Comments.*, Docs.*, Comments.DateCreation as CommentsDateCreation, Docs.DateCreation as DocsDateCreation, Comments.FileName as CommentsFileName  from Docs  Left Outer Join Comments ON (Docs.DocID = Comments.DocID  And (SpecialInfo='VISAWAITING')) where (UserID = N'" & Session("UserID") & "' and (StatusCompletion is NULL or (StatusCompletion<>'1' and StatusCompletion<>'0')) and (IsActive<>'N' or IsActive is Null)) "
			VAR_ListDocSQLSuffix = " order by DateEventEnd, Docs.DateCreation desc, Docs.DocID desc, Comments.DateCreation"
		End If
	End If
'Запрос №44 - СТС - start
End If

'20120202 - start
If InStr(Session("Department"), SIT_STS_ROOT_DEPARTMENT) = 1 Then 'Для СТС подменяются запросы
	If InStr(UCase(Request.ServerVariables("URL")), UCase("/ListDoc.asp")) > 0 and UCase(Request("RespDocs")) = "Y" Then 'BUT_RESPONSIBLE - ПОРУЧЕНИЯ МНЕ
		VAR_ListDocSQLPrefix = "select Docs.* from Docs where (PATINDEX(N'%<" & Session("UserID") & ">%', NameResponsible)>0 and (StatusCompletion<>'1' and StatusCompletion<>'0' Or StatusCompletion is Null) and DateCompleted is Null and ClassDoc <> N'" & STS_SLUZH_ZAPISKA_OVERTIME_FACT & "' and ClassDoc <> N'" & STS_SLUZH_ZAPISKA_OVERTIME_PLAN & "' and ClassDoc <> N'" & STS_SLUZH_ZAPISKA_OVERTIME2 & "' and CharIndex(N'" & SIT_DOGOVORI_NEW & "', ClassDoc) <> 1) "
		VAR_ListDocSQLSuffix = " order by Docs.DateCreation desc, Docs.DocID desc"
	End If
End If
'20120202 - end

'сортировка для списка утверждение
If InStr(UCase(Request.ServerVariables("URL")), UCase("/ListDoc.asp")) > 0 and UCase(Request("ApprDocs")) = "Y" Then
  VAR_ListDocOrderBy = "Case When IsNull(DateCompletion, { d '1920-01-01' }) <= { d '1920-01-01' } Then { d '2100-01-01' } Else DateCompletion End, DateCreation, DocID"
End If
'ph - 20101020 - end

'ph - 20110415 - start
If InStr(UCase(Request.ServerVariables("URL")), UCase("/ListDoc.asp")) > 0 and UCase(Request("BadDocs")) = "Y" Then
  VAR_ListDocSQL = "select Docs.* from Docs where ((DateCompletion<=GetDate()+3 and Not (DateCompletion is Null) and DateCompletion > { d '1900-02-02' } and ((StatusCompletion<>'1' And StatusCompletion<>'0' ) or StatusCompletion is NULL)) or (DateExpiration<=GetDate()+3 and Not (DateExpiration is Null) and DateExpiration>{ d '1900-02-02' } and (StatusCompletion<>'1' or StatusCompletion is NULL))) and (CharIndex(N'<" & Session("UserID") & ">', NameAproval) > 0 or CharIndex(N'<" & Session("UserID") & ">', NameResponsible) > 0 or CharIndex(N'<" & Session("UserID") & ">', NameCreation)> 0 or CharIndex(N'<" & Session("UserID") & ">', NameControl)> 0) order by DateCompletion desc"
End If
'ph - 20110415 - end

'Запрос №38 - СТС - start
'Максимально допустимое число согласующих
If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_DOGOVORI_NEW)) = 1 Then
  SIT_MaxUsersInListToReconcile = 12
End If
'Запрос №38 - СТС - end

'amw 25-10-2013 (start)
' Due to some requirements from Mikron management. ( CFO, purchase deptmt, legal deptmt)
' they would like to approve the docs w/o reconcilation completed, and don't like if process should stop.
'amw 27-08-2014
' If InStr(UCase(Session("Department")), UCase(SIT_MIKRON)) = 1 Then
' некорректная работа с согласованием. Если документ создан НЕ пользователем МИКРОН, и есть отказавшие, то
' при просмотре пользователем МИКРОН, согласование возобновляется.
If InStr(UCase(GetRootDepartment(Session("CurrentDepartmentDoc"))), UCase(SIT_MIKRON)) = 1 Then
'Document aproval is NOT permitted while reconciliation process is not finished
   Var_ApprovalPermitted=False 
'Document aproval is permitted even reconciliation process is not finished yet
'   Var_ApprovalPermitted=True 

'Document aproval is NOT permitted if some reconciliation list user refused reconciliation (not agree)
'   Var_ApprovalIfAllAgree = True
'By default. Document aproval is permitted even if any reconciliation list user refused reconciliation (not agree)
   Var_ApprovalIfAllAgree=False

'Next document reconciliation step is NOT permitted if some previous reconciliation list user refused reconciliation (not agree)
'   Var_ReconciliationIfAllAgree = True
'By default
   Var_ReconciliationIfAllAgree = False
'   Var_ReconciliationIfAllAgreeDoesNotExcludeCurrentLevel="Y" 'Works in case Var_ReconciliationIfAllAgree = True

'   Var_ReconciliationNotToResetModifyDocRecord = "" 'Reconciliation should be reset if doc was changed
   Var_ReconciliationNotToResetModifyDocRecord = "Y"

   VAR_PermitToAgreeFiles = True
   VAR_DocCreatorCanUpdateDocWithoutChangingStatuses="Y"
End If
'amw 25-10-2013 - end

Select Case UCase(Request("l"))
  Case "" DOCS_CommentAgree = "Comment"
  Case "RU" DOCS_CommentAgree = "Комментарий"
  Case "3" DOCS_CommentAgree = "Komentář"
End Select

If InStr(Session("Department"), SIT_MIKRON) = 1 Then 'amw
   SIT_LinkInstructionHint = "Руководство по работе с системой в НИИМЭ и Микрон"
   SIT_LinkInstructionFile = "Manuals/UserInstruction_MIKRON.docx" 'Имя файла с инструкцией на данном языке
ElseIf InStr(Session("Department"), SIT_RTI) = 1 Then
SIT_LinkInstructionHint = "Руководство по работе с системой"
SIT_LinkInstructionFile = "Manuals/UserInstruction.pdf" 'Имя файла с инструкцией на данном языке
ElseIf InStr(Session("Department"), SIT_MINC) = 1 Then
SIT_LinkInstructionHint = "Руководство по работе с системой в ОАО РТИ"
SIT_LinkInstructionFile = "Manuals/UserInstruction_minc.pdf" 'Имя файла с инструкцией на данном языке
End If ' 


If InStr(Session("Department"), SIT_RTI) = 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(RTI_CONTRACT)) > 0 and InStr(UCase(Request.ServerVariables("URL")),UCase("/ShowDoc.asp"))>0 Then
    DOCS_AGREECancel="Отменить отказ"
    DOCS_AGREECancel1="Отменить отказ в согласовании, если замечания отказавшего устранены"
    DOCS_Cancel = "Отменить отказ"
Else
    DOCS_AGREECancel="Отменить согласование"
    DOCS_AGREECancel1="Отменить произведенное согласование некоторых участников"
    DOCS_Cancel="Отменить"
End If

'разрешаем редактировать документ Заявка на оплату пользователю с ролью..
If InStr(UCase(Request.ServerVariables("URL")),UCase("/ShowDoc.asp"))>0 Then
    'out Request.ServerVariables("URL")
    'out Session("UserID")
    'If dsDoc("ClassDoc")="Закупки*Purchase*Nákup/Заявка на оплату УК*Payment order MC*Zadost o platbu SS" Then
        If Session("UserID")="VNikulin" Then
            VAR_CanEditExtData = "Y"
            VAR_WriteAccess = "Y"
            VAR_ReadAccess = "Y"
        Else
            VAR_CanEditExtData = ""
            VAR_WriteAccess = ""
        End If
    'out VAR_ReadAccess
    'out VAR_WriteAccess
    'End If
End If

'разрешаем редактировать документ Заявка на оплату пользователю с ролью..
If InStr(UCase(Request.ServerVariables("URL")),UCase("/ChangeExtData.asp"))>0 Then
    'out Request.ServerVariables("URL")
    'out Session("UserID")
    'If dsDoc("ClassDoc")="Закупки*Purchase*Nákup/Заявка на оплату УК*Payment order MC*Zadost o platbu SS" Then
        If Session("UserID")="VNikulin" Then
            VAR_CanEditExtData = "Y"
            VAR_WriteAccess = "Y"
            VAR_ReadAccess = "Y"
        Else
            VAR_CanEditExtData = ""
            VAR_WriteAccess = ""
        End If
    'out VAR_ReadAccess
    'out VAR_WriteAccess
    'End If
End If

VAR_WindowsUsersCanLogin = "Y"
VAR_CanEditPartnerRating="Y"
bUseLangRU = True
'DOCS_CommentAgree = "_"

%>
