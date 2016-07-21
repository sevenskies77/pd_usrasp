<%
'Variables:
'CurrentClassDoc - current document category
'CurrentDocFieldOrder - current document category field order
'DependentDocFields - dependent document list field order
'
'Out CurrentClassDoc
'Select Case CurrentClassDoc
     'Case "Accounts" ' - document category to be processed
		'CurrentDocFieldOrder="DocID,Correspondent,Resolution,History,Result"
		DependentDocFields="DocID,Description,AmountDoc,DateActivation"
'End Select

'CurrentDocFieldOrder="DocID,DocIDadd,DocIDParent,DocIDPrevious,DocIDIncoming,Author,Correspondent,Resolution,History,Result,PercentCompletion,Department,Name,Description,LocationURL,LocationPaper,Currency,CurrencyRate,Rank,FileNamePrefix,FileName,FileNameNameLastModification,FileNameDateLastAccessed,FileNameDateLastModification,LocationPath,ExtInt,PartnerName,StatusDevelopment,StatusArchiv,StatusCompletion,StatusDelivery,StatusPayment,TypeDoc,ClassDoc,ActDoc,InventoryUnit,PaymentMethod,AmountDoc,QuantityDoc,DateActivation,SecurityLevel,DateCreation,DateCompletion,DateCompleted,DateExpiration,DateSigned,NameCreation,NameAproval,NameApproved,DateApproved,NameControl,ListToEdit,ListToView,ListToReconcile,ListReconciled,NameResponsible,NameLastModification,DateLastModification,UserFieldText1,UserFieldText2,UserFieldText3,UserFieldText4,UserFieldText5,UserFieldText6,UserFieldText7,UserFieldText8,UserFieldMoney1,UserFieldMoney2,UserFieldDate1,UserFieldDate2,IsActive,DateActive,BusinessProcessStep,ExtPassword,GUID,Content"

If False Then
'If IsHelpDeskDoc() Then
If Trim(Request("justcreated"))<>"" Then
	If S_NameResponsible<>"" Then
		If S_Correspondent<>"" Then
			If InStr(S_Correspondent, GetLogin(S_NameResponsible))<=0 Then
'AddLogD "SetDocField, NameResponsible="""" "
				SetDocField Request("DocID"), "NameResponsible", ""
			End If
		End If
	End If
	If Not IsAdmin() And Not IsSupervisor() Then

sUserDirName="Группы пользователей"
'nKeyField
sKeyFieldValue=S_UserFieldText4
sKeyFieldValue2=Session("Company")
'Out "UserFieldText4:"+S_UserFieldText4

Set dsTemp = Server.CreateObject("ADODB.Recordset")
        If sVersion = "MSSQL" Then
sSQL = "select * from (UserDirectories Left Outer Join UserDirValues ON UserDirValues.UDKeyField = UserDirectories.KeyField) where Name='" + sUserDirName + "' And PATINDEX('%" + sKeyFieldValue + "%', Field1) <> 0 And Field3='" + sKeyFieldValue2 + "'"
        ElseIf sVersion = "MSACCESS" Then
sSQL = "select * from (UserDirectories Left Outer Join UserDirValues ON UserDirValues.UDKeyField = UserDirectories.KeyField) where Name='" + sUserDirName + "' And InStr(Field1, '" + sKeyFieldValue + "') <> 0 And Field3='" + sKeyFieldValue2 + "'"
        End If
'Out "sSQL:" + sSQL
dsTemp.Open sSQL, Conn, 3, 1, &H1
If Not dsTemp.EOF Then
'Out "Correspondent:" + MyCStr(dsTemp("Field2"))
	SetDocField Request("DocID"), "Correspondent", MyCStr(dsTemp("Field2"))
End If
dsTemp.Close
	End If
End If
End If 'False

'------------------------------------ RTI SITRONICS ----------------------------------
'ph - 20090714 - start
'Доступ к документу привилегированным пользователям
Session("SIT_UserCanDownloadFiles") = False
If Instr(UCase(SIT_UsersListToAccessAllCategoryDocs),"<"+UCase(Session("UserID"))+">")>0 Then
  If UCase(GetRootDepartment(Session("Department"))) = UCase(GetRootDepartment(dsDoc("Department"))) Then
    VAR_ReadAccess="Y"
    oPayDox.VAR_ReadAccess=VAR_ReadAccess
	'Разрешить доступ к файлам
    Session("SIT_UserCanDownloadFiles") = True
  End If
End If
'ph - 20090714 - end

If InStr(UCase(S_ClassDoc),UCase(SIT_VHODYASCHIE)) = 1 Then
  'Порядок следования полей при просмотре
  CurrentDocFieldOrder = "DocID, DocIDParent, Name, UserFieldText1, UserFieldText7, DateActivation, UserFieldText2, ListToView, PartnerName, UserFieldText3, DocIDIncoming, UserFieldDate1, UserFieldText4, UserFieldText5, UserFieldText6, UserFieldDate2, UserFieldText8, Author, Content"
ElseIf InStr(UCase(S_ClassDoc),UCase(SIT_VHODYASCHIE_ACC)) = 1 Then
  'Порядок следования полей при просмотре
  CurrentDocFieldOrder = "DocID, Name, UserFieldText1, UserFieldText7, DateActivation, UserFieldText2, ListToView, PartnerName, UserFieldText3, DocIDIncoming, UserFieldDate1, UserFieldText4, UserFieldText5, UserFieldText6, Content"
ElseIf InStr(UCase(S_ClassDoc),UCase(SIT_ISHODYASCHIE)) = 1 or InStr(UCase(S_ClassDoc),UCase(MINC_ISHODYASCHIE)) = 1 Then 
  'Порядок следования полей при просмотре
  CurrentDocFieldOrder = "DocID, UserFieldText1, DocIDParent, UserFieldDate1, NameAproval, Name, Content, PartnerName, Correspondent, UserFieldText3, UserFieldText4, Author, ListToReconcile, DateActivation"
  '2015-04-09 kkoshkin отображаем поле "тип отправки" для минца
  if InStr(UCase(S_Department),UCase(SIT_MINC)) = 1 then
  CurrentDocFieldOrder = "DocID, UserFieldText1, DocIDParent, UserFieldDate1, NameAproval, Name, Content, PartnerName, Correspondent, UserFieldText3, UserFieldText4, Author, UserFieldText5, ListToReconcile, UserFieldText2, DateActivation"
  end if
  '2105-04-09 end
  
  'Шаблоны MS Word
  If InStr(UCase(S_Department),UCase(SIT_SITRONICS)) = 1 then
    Select case UCase(Request("l"))
      case "RU"
        VAR_DocTemplateFilename = "Letter_Sistema.doc"
      case ""
        VAR_DocTemplateFilename = "Letter_Foreign.doc"
    End Select
  ElseIf InStr(UCase(S_Department),UCase(SIT_STS)) = 1 then
    Select case UCase(Request("l"))
      case "RU"
        VAR_DocTemplateFilename = "Letter_Sistema_STS.doc"
      case ""
         VAR_DocTemplateFilename = "Letter_Foreign_STS.doc"
      End Select
   ElseIf InStr(UCase(S_Department),UCase(SIT_SIB)) = 1 then
      VAR_DocTemplateFilename = "SluzZap_SIB.doc"
   ElseIf InStr(UCase(S_Department),UCase(SIT_RTI)) = 1 then
      VAR_DocTemplateFilename = "Letter_Sistema_RTI.doc"
   ElseIf InStr(UCase(S_Department),UCase(SIT_VTSS)) = 1 then
      VAR_DocTemplateFilename = "Letter_Sistema_RTI.doc"
   ElseIf InStr(UCase(S_Department),UCase(SIT_MIKRON)) = 1 then
      VAR_DocTemplateFilename = "Letter_Sistema_Mikron.doc"
   ElseIf InStr(UCase(S_Department),UCase(SIT_MINC)) = 1 then
      VAR_DocTemplateFilename = "Letter_MINC.doc"
   End If

'Специальные виды СЗ в СТС
' rmanyushin 20.08.2009 17.27 Start
ElseIf InStr(UCase(S_ClassDoc),UCase(STS_SLUZH_ZAPISKA_OVERTIME)) = 1 Then
  'Порядок следования полей при просмотре
  CurrentDocFieldOrder = "DocID, DocIDParent, Name, Content, ListToView, NameAproval, Author, ListToReconcile, DateActivation, Resolution, UserFieldDate2, UserFieldText1, UserFieldText2, UserFieldText3, UserFieldText4, UserFieldText5, UserFieldText6"
  'Шаблоны MS Word
   VAR_DocTemplateFilename = ""
'В служебных записках убрать возможность изменения даты подписи-утверждения
'(не должно запрашивать)
   VAR_NotToAskDateDuringApproval = "Y"
'rmanyushin 20.08.2009 17.27 End


'rmanyushin 136964 08.11.2010 Start
ElseIf InStr(UCase(S_ClassDoc),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) = 1 Then
  'Порядок следования полей при просмотре
  CurrentDocFieldOrder = "DocID, DocIDParent, Name, Author, NameResponsible, UserFieldText4, Content, ListToReconcile, NameAproval, ListToView, UserFieldDate1"
  'Шаблоны MS Word
   VAR_DocTemplateFilename = ""
'В служебных записках убрать возможность изменения даты подписи-утверждения
'(не должно запрашивать)
   VAR_NotToAskDateDuringApproval = "Y"
'rmanyushin 136964 08.11.2010 End

'Запрос №46 - СТС - start
ElseIf InStr(UCase(S_ClassDoc),UCase(STS_SLUZH_ZAPISKA_OVERTIME_PLAN)) = 1 Then
	'Порядок следования полей при просмотре
	CurrentDocFieldOrder = "DocID, Name, Author, Content, NameResponsible, UserFieldText3, Correspondent, ListToReconcile, NameAproval, ListToView"
	'Запрет добавления пользователей в адресатов (функциональные руководители), скрыть кнопку с плюсом
	VAR_AddUsersToCorrespondent = ""
	'Запрет добавления пользователей в получателей, скрыть кнопку с плюсом
	VAR_AddUsersToListToView = ""
	'Шаблоны MS Word
	VAR_DocTemplateFilename = ""
	'В служебных записках убрать возможность изменения даты подписи-утверждения (не должно запрашивать)
	VAR_NotToAskDateDuringApproval = "Y"
ElseIf InStr(UCase(S_ClassDoc),UCase(STS_SLUZH_ZAPISKA_OVERTIME_FACT)) = 1 Then
	'Порядок следования полей при просмотре
	CurrentDocFieldOrder = "DocID, DocIDParent, Name, Author, Content, NameResponsible, UserFieldText3, Correspondent, ListToReconcile, NameAproval, ListToView"
	'Запрет добавления пользователей в адресатов (функциональные руководители), скрыть кнопку с плюсом
	VAR_AddUsersToCorrespondent = ""
	'Запрет добавления пользователей в получателей, скрыть кнопку с плюсом
	VAR_AddUsersToListToView = ""
	'Шаблоны MS Word
	VAR_DocTemplateFilename = ""
	'В служебных записках убрать возможность изменения даты подписи-утверждения (не должно запрашивать)
	VAR_NotToAskDateDuringApproval = "Y"
'Запрос №46 - СТС - end

'rmanyushin 119579 19.08.2010 Start
ElseIf InStr(UCase(S_ClassDoc),UCase(STS_SLUZH_ZAPISKA_HOLIDAY)) = 1 Then
  'Порядок следования полей при просмотре
  CurrentDocFieldOrder = "DocID, DocIDParent, Name, Content, ListToView, NameAproval, Author, ListToReconcile, DateActivation, Resolution, UserFieldDate1, UserFieldText1, UserFieldText2, UserFieldText3, UserFieldText4, UserFieldText5, UserFieldText6"
  'Шаблоны MS Word
   VAR_DocTemplateFilename = ""
  'В служебных записках убрать возможность изменения даты подписи-утверждения (не должно запрашивать)
  VAR_NotToAskDateDuringApproval = "Y"
'rmanyushin 119579 19.08.2010 End   
' Специальные виды СЗ в СТС

' Служебная записка (общая форма)
ElseIf InStr(UCase(S_ClassDoc),UCase(SIT_SLUZH_ZAPISKA)) = 1 Then
  'Порядок следования полей при просмотре
  CurrentDocFieldOrder = "DocID, DocIDParent, Name, Content, ListToView, NameAproval, Author, ListToReconcile, DateActivation, Resolution, UserFieldDate2, UserFieldText1, UserFieldText2, UserFieldText3, UserFieldText4, UserFieldText5, UserFieldText6"
  'Шаблоны MS Word
  If InStr(UCase(S_Department),UCase(SIT_SITRONICS)) = 1 then
    VAR_DocTemplateFilename = "SluzZap.doc"
  ElseIf InStr(UCase(S_Department),UCase(SIT_MIKRON)) = 1 then
    VAR_DocTemplateFilename = "SZ_Mikron.doc"
  ElseIf InStr(UCase(S_Department),UCase(SIT_STS)) = 1 then
    VAR_DocTemplateFilename = "SluzZap_STS.doc"
  ElseIf InStr(UCase(S_Department),UCase(SIT_RTI)) = 1 then
    Select case S_UserFieldText2
    case "Служебная записка"
    VAR_DocTemplateFilename = "SluzZap_RTI.doc"
    case "Докладная записка"
    VAR_DocTemplateFilename = "DoklZap_RTI.doc"
  End Select
  ElseIf InStr(UCase(S_Department),UCase(SIT_VTSS)) = 1 then
    VAR_DocTemplateFilename = "SluzZap_VTSS.doc"
 
  End If
  'В служебных записках убрать возможность изменения даты подписи-утверждения (не должно запрашивать)
  VAR_NotToAskDateDuringApproval = "Y"
ElseIf InStr(UCase(S_ClassDoc),UCase(SIT_RASP_DOCS)) = 1 Then
  'Порядок следования полей при просмотре
  CurrentDocFieldOrder = "DocID, Name, UserFieldText1, DateActivation, Content, UserFieldText2, Author, ListToReconcile, NameAproval, ListToView, LocationPath, DocIDParent, SecurityLevel"	
  'Шаблоны MS Word
  ' DmGorsky_5: меняем местами проверку: вначале ищем SIT_SITRU ("СИТРОНИКС ИТ"), потом, если не найдено, - SIT_SITRONICS ("СИТРОНИКС")
  If InStr(UCase(S_Department),UCase(SIT_SITRU)) = 1 then ' DmGorsky
    Select case S_UserFieldText1
      case SIT_Orders_Prikaz
        VAR_DocTemplateFilename = "Prikaz_SITRU.doc"
      case SIT_Orders_Rasporyajenie
        VAR_DocTemplateFilename = "Raspor_SITRU.doc"
      case SIT_Orders_Prikaz_ND ' DmGorsky_6
        VAR_DocTemplateFilename = "Prikaz_ND_SITRU.doc" ' DmGorsky_6
    End Select
    VAR_StandardFileNameReconciliationList = "ReconciliationListRUS_SITRU.doc"
  ElseIf InStr(UCase(S_Department),UCase(SIT_STS)) = 1 then
    Select case S_UserFieldText1
      case SIT_Orders_Prikaz
        VAR_DocTemplateFilename = "Prikaz_STS_RU.doc"
      case SIT_Orders_Rasporyajenie
        VAR_DocTemplateFilename = "Raspor_STS_RU.doc"
    End Select
  ElseIf InStr(UCase(S_Department),UCase(SIT_SITRONICS)) = 1 then
    Select case S_UserFieldText1
      case SIT_Orders_Prikaz
        VAR_DocTemplateFilename = "Prikaz_RU.doc"
      case SIT_Orders_Rasporyajenie
        VAR_DocTemplateFilename = "Raspor_RU.doc"
    End Select
    '///////////// AMW - Mikron
  ElseIf InStr(UCase(S_Department),UCase(SIT_MIKRON)) = 1 then
    Select case S_UserFieldText1
      case SIT_Orders_Prikaz_MIKRON
        VAR_DocTemplateFilename = "Order_Mikron.doc"
      case SIT_Orders_Prikaz_NIIME
        VAR_DocTemplateFilename = "Order_NIIME.doc"
      case SIT_Orders_Rasporyajenie
        VAR_DocTemplateFilename = "ByLaw_Mikron.doc"
    End Select
'///////////// AMW - Mikron

  ElseIf InStr(UCase(S_Department),UCase(SIT_RTI)) = 1 then
    Select case S_UserFieldText1
      case SIT_Orders_Prikaz_RTI
        VAR_DocTemplateFilename = "Prikaz_RU_RTI.doc"
      case SIT_Orders_Rasporyajenie
        VAR_DocTemplateFilename = "Raspor_RU_RTI.doc"
    End Select

  End If
'vnik_protocols
'vnik_protocolsCPC
ElseIf InStr(UCase(S_ClassDoc),UCase(SIT_PROTOCOLS)) > 0 or InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_PROTOCOLS_CPC)) > 0 Then
'Порядок следования полей при просмотре
  CurrentDocFieldOrder = "DocID, Name, UserFieldDate2, UserFieldText1, DateActivation, UserFieldDate1, Content, UserFieldText2, Author, ListToReconcile, NameAproval, ListToView, DocIDParent"	
'vnik_protocolsCPC

'rti_protocol
ElseIf InStr(UCase(S_ClassDoc),UCase(RTI_PROTOCOL)) = 1 Then '"Закупки РТИ/Протокол ЦЗК"
'Порядок следования полей при просмотре
  CurrentDocFieldOrder = "DocID, DocIDParent, Name, UserFieldDate2, Content, Author, ListToReconcile, NameAproval, ListToView"	
  VAR_DocTemplateFilename = "Protocol_RTI.doc"
'rti_protocol
'///////////// AMW - Mikron
'"Закупки МИКРОН/Протокол ЗК"
ElseIf InStr(UCase(S_ClassDoc),UCase(MIKRON_PROTOCOL)) = 1 Then
  'Порядок следования полей при просмотре
  CurrentDocFieldOrder = "Name,DocID,DocIDParent,UserFieldText1,UserFieldMoney1,UserFieldText2," + _
                         "UserFieldText3,UserFieldDate1,UserFieldDate2,NameAproval,ListToView," + _
                         "PartnerName,AmountDoc,Currency,Resolution"
  VAR_DocTemplateFilename = "PZK_Mikron.doc"

'"Закупки МИКРОН/Опросный лист для ПЗК"
ElseIf InStr(UCase(S_ClassDoc),UCase(MIKRON_RL_PROTOCOL)) = 1 Then
  'Порядок следования полей при просмотре
  CurrentDocFieldOrder = "DocID,Name,DocIDParent,UserFieldText6,UserFieldText3,UserFieldText4," + _
                         "UserFieldText5,Currency,QuantityDoc,PartnerName,AmountDoc," + _
                         "UserFieldText1,UserFieldMoney1,UserFieldText2,UserFieldMoney2," + _
                         "UserFieldText7,NameAproval,ListToReconcile,ListToView,Content,Author,Department"
  VAR_DocTemplateFilename = "RL_for_PCP_Mikron.doc"
'  'Прячем кнопку отказать в утверждении
'  If InStr(VAR_ButtonsNotToShow & ",", "ClickRefuseApp,") = 0 Then
'     VAR_ButtonsNotToShow = VAR_ButtonsNotToShow & "ClickRefuseApp,"
'  EndIf
  If InStr(VAR_ButtonsNotToShow & ",", "ClickRefuse,") = 0 Then
     VAR_ButtonsNotToShow = VAR_ButtonsNotToShow & "ClickRefuse,"
  End If
'  If InStr(VAR_ButtonsNotToShow & ",", "ClickVisaAdd,") = 0 Then
'     VAR_ButtonsNotToShow = VAR_ButtonsNotToShow & "ClickVisaAdd,"
'  EndIf
'  If InStr(VAR_ButtonsNotToShow & ",", "ClickVisaDelegate,") = 0 Then
'     VAR_ButtonsNotToShow = VAR_ButtonsNotToShow & "ClickVisaDelegate,"
'  EndIf
'  If InStr(VAR_ButtonsNotToShow & ",", "ClickToReview,") = 0 Then
'     VAR_ButtonsNotToShow = VAR_ButtonsNotToShow & "ClickToReview,"
'  EndIf
'///////////// AMW - Mikron

'vnik_protocols
'vnik_payment_order
ElseIf InStr(UCase(S_ClassDoc),UCase(SIT_PAYMENT_ORDER)) > 0 Then
'Порядок следования полей при просмотре
  CurrentDocFieldOrder = "DocID, Name, DocIDParent, UserFieldDate2, UserFieldText2, Author, ListToReconcile, NameAproval, UserFieldText5, ListToView, Department, UserFieldText4, Description, UserFieldMoney1, Currency, PartnerName"
'vnik_payment_order
'vnik_purchase_order
ElseIf InStr(UCase(S_ClassDoc),UCase(SIT_PURCHASE_ORDER)) > 0 Then
'Порядок следования полей при просмотре
  'CurrentDocFieldOrder = "DocID, Name, DocIDParent, UserFieldDate2, UserFieldText2, Author, ListToReconcile, NameAproval, UserFieldText5, ListToView, Department, UserFieldText4, Description, UserFieldMoney1, Currency, PartnerName"
  CurrentDocFieldOrder = "DocID, Name, UserFieldText6, PartnerName, DocIDParent, Author, Department, ContractType, UserFieldText8, Currency, UserFieldText7, UserFieldMoney1, UserFieldMoney2, QuantityDoc, UserFieldDate2, UserFieldText4, UserFieldText3, UserFieldText2, UserFieldText1, ListToReconcile, NameAproval, ListToView, UserFieldText5, Content"	
  VAR_DocTemplateFilename = "Purchase_Order_MC.doc"
'vnik_purchase_order

'rti_purchase_order
ElseIf InStr(UCase(S_ClassDoc),UCase(RTI_PURCHASE_ORDER)) = 1 Then
'Порядок следования полей при просмотре
  CurrentDocFieldOrder = "DocID, Name, DocIDParent, Author, Department, UserFieldMoney1, UserFieldText4, UserFieldText3, UserFieldText2, UserFieldText1, ListToReconcile, NameAproval, ListToView, Content"	
  VAR_DocTemplateFilename = "Purchase_Order_RTI.doc"
  'rti_purchase_order
  
'rti_payment_order
ElseIf InStr(UCase(S_ClassDoc),UCase(RTI_PAYMENT_ORDER)) > 0 Then
'Порядок следования полей при просмотре
  CurrentDocFieldOrder = "DocID, Name, DocIDParent, Author, UserFieldText1, UserFieldText2, UserFieldText3, UserFieldText4,UserFieldText6, UserFieldMoney1, UserFieldText5, ListToReconcile, NameAproval, ListToView, Department, Description, PartnerName, Content"
  VAR_DocTemplateFilename = "Payment_Order_RTI.doc"
'rti_payment_order

'rti_contract
ElseIf InStr(UCase(S_ClassDoc),UCase(RTI_CONTRACT)) > 0 Then
'Порядок следования полей при просмотре
    CurrentDocFieldOrder = "DocID, DocIDParent, Author, UserFieldText1, UserFieldText2, UserFieldText3, UserFieldMoney1, UserFieldText4, UserFieldText5, ListToReconcile, NameAproval, ListToView, Department,  PartnerName, Content"
  'VAR_DocTemplateFilename = "ReconciliationListRUSContractsHQ.doc"
'rti_contract
'rti_bsap
ElseIf InStr(UCase(S_ClassDoc),UCase(RTI_BSAP)) > 0 Then
'Порядок следования полей при просмотре  
  CurrentDocFieldOrder = "DocID, Name, DocIDParent, Author, Department,  ListToReconcile, NameAproval, ListToView, Content"	
  VAR_DocTemplateFilename = "BSAP_RTI.xls"
'rti_bsap
'///////////// AMW - Mikron
'Порядок следования полей при просмотре
'MIKRON_purchase_order
ElseIf InStr(UCase(S_ClassDoc),UCase(MIKRON_PURCHASE_ORDER)) = 1 Then
  CurrentDocFieldOrder = "DocID,Name,UserFieldText3,Author,Department,UserFieldMoney1,UserFieldText1," + _
                         "UserFieldText2,UserFieldText5,UserFieldText4,UserFieldText6,ListToReconcile," + _
                         "NameAproval,Resolution"
  VAR_DocTemplateFilename = "Purchase_Order_Mikron.doc"
  
'MIKRON_payment_order
ElseIf InStr(UCase(S_ClassDoc),UCase(MIKRON_PAYMENT_ORDER)) > 0 Then
  CurrentDocFieldOrder = "DocID,Name,DocIDParent,Author,UserFieldText1,UserFieldText2,UserFieldText3," + _
                         "UserFieldText4,UserFieldMoney1,UserFieldText5,ListToReconcile,NameAproval," + _
                         "ListToView,Department,Description,PartnerName"
  VAR_DocTemplateFilename = "Payment_Order_Mikron.doc"

'MIKRON_contract
ElseIf InStr(UCase(S_ClassDoc),UCase(MIKRON_CONTRACT)) > 0 or _
       InStr(UCase(S_ClassDoc),UCase(MIKRON_S_CONTRACT)) > 0 Then
  CurrentDocFieldOrder = "DocID,DocIDParent,Author,NameResponsible,Department,DocIDIncoming," + _
                         "BusinessUnit,PartnerName,UserFieldText1,Description,Name,UserFieldText2," + _
                         "UserFieldText3,AmountDoc,Currency,UserFieldText5,UserFieldDate4," + _
                         "UserFieldDate6,ListToReconcile,NameAproval,SecurityLevel,UserFieldDate5," + _
                         "ContractType,AddFieldText1,AddFieldText2,UserFieldText4"
'amw 04-08-2014
  VAR_DocTemplateFilename = ""
'  VAR_DocTemplateFilename = "RL_MikronContract.docx"

'MIKRON_additional_agreement
ElseIf InStr(UCase(S_ClassDoc),UCase(MIKRON_ADD_CONTRACT)) > 0 Then
  CurrentDocFieldOrder = "DocID,DocIDParent,DocIDPrevious,Author,NameResponsible,Department,DocIDIncoming," + _
                         "BusinessUnit,PartnerName,UserFieldText1,Name,UserFieldText2,UserFieldText3," + _
                         "Description,AmountDoc,Currency,UserFieldDate4,UserFieldDate6," + _
                         "ListToReconcile,NameAproval,SecurityLevel,UserFieldDate5"
'amw 04-08-2014
  VAR_DocTemplateFilename = ""
'  VAR_DocTemplateFilename = "RL_MikronContract.docx"
ElseIf InStr(UCase(S_ClassDoc),UCase(MIKRON_NDA_CONTRACT)) > 0 Then
  CurrentDocFieldOrder = "Name,DocID,Author,Department,UserFieldText1,Description,PartnerName," + _
                         "DocIDIncoming,UserFieldDate4,UserFieldDate5,UserFieldDate6," + _
                         "ListToReconcile,NameAproval,Content"
  VAR_DocTemplateFilename = "NDA_Mikron.doc"

'MIKRON_BSAP
ElseIf InStr(UCase(S_ClassDoc),UCase(MIKRON_BSAP)) > 0 Then
  CurrentDocFieldOrder = "DocID,Name,DocIDParent,Author,Department,UserFieldText3,Description," + _
                         "BusinessUnit,ListToView,UserFieldText4,NameAproval,Currency,PartnerName," + _
                         "AmountDoc,UserFieldText1,UserFieldMoney1,UserFieldText2,UserFieldMoney2"
  VAR_DocTemplateFilename = "BSAP_Mikron.xlsx"

'MIKRON_RL_MEMO
ElseIf InStr(UCase(S_ClassDoc),UCase(MIKRON_RL_MEMO)) > 0 Then
  CurrentDocFieldOrder = "DocID,Name,Author,Department,UserFieldText3,Description,QuantityDoc,Currency," + _
                         "PartnerName,AmountDoc,DocIDPrevious,UserFieldDate1,UserFieldText5," + _
                         "UserFieldText1,UserFieldMoney1,DocIDIncoming,UserFieldDate2,UserFieldText6," + _
                         "UserFieldText2,UserFieldMoney2,DocIDadd,UserFieldDate3,UserFieldText7," + _
                         "NameAproval,ListToView,UserFieldText4,UserFieldText8,ListToReconcile"
  VAR_DocTemplateFilename = "MEMO_Mikron.xlsx"
ElseIf InStr(UCase(S_ClassDoc),UCase(MIKRON_OLD_CONTRACT)) > 0 Then
  CurrentDocFieldOrder = "Name,DocID,UserFieldDate5,Author,Department,UserFieldText1,UserFieldText8," + _
                         "Description,UserFieldText3,PartnerName,UserFieldText4,AmountDoc,Currency," + _
                         "UserFieldText5,UserFieldDate4,UserFieldDate6,UserFieldText2,UserFieldText6," + _
                         "UserFieldText7,UserFieldDate1,NameAproval"
'amw 04-08-2014
  VAR_DocTemplateFilename = ""
ElseIf InStr(UCase(S_ClassDoc),UCase(MIKRON_EXPORT_CONTRACT)) > 0 or _
       InStr(UCase(S_ClassDoc),UCase(MIKRON_EXPADD_CONTRACT)) > 0 Then
  CurrentDocFieldOrder = "Name,DocID,DocIDParent,DocIDPrevious,Author,NameResponsible,Department,DocIDIncoming," + _
                         "PartnerName,UserFieldText1,UserFieldText2,UserFieldText3," + _
                         "Description,AmountDoc,Currency,UserFieldDate4,UserFieldDate6," + _
                         "ListToReconcile,NameAproval,SecurityLevel,UserFieldDate5"
  VAR_DocTemplateFilename = "EXPORT_Mikron.doc"
'///////////// AMW - Mikron


'vnik_contracts
ElseIf InStr(UCase(S_ClassDoc),UCase(SIT_CONTRACTS_MC)) > 0 Then
'Порядок следования полей при просмотре
    CurrentDocFieldOrder = "DocID, Name, DocIDParent, UserFieldDate2, UserFieldText2, Author, ListToReconcile, NameAproval, UserFieldText5, ListToView, Department, UserFieldText4, Description, UserFieldMoney1, Currency, PartnerName, UserFieldText3, AddFieldText2, AddFieldText1"
  VAR_DocTemplateFilename = "ReconciliationListRUSContractsHQ.doc"
'vnik_contracts
ElseIf InStr(UCase(S_ClassDoc),UCase(SIT_NORM_DOCS)) = 1 Then
  'Порядок следования полей при просмотре
  CurrentDocFieldOrder = "DocID, Name, UserFieldText1, UserFieldDate1, Author, DateActivation, Content, Description, UserFieldText3, UserFieldText4, UserFieldDate2 ,UserFieldText2, ListToReconcile, NameAproval, LocationPath, ListToView, SecurityLevel, Department "	
ElseIf InStr(UCase(S_ClassDoc),UCase(SIT_ZADACHI)) = 1 Then
  'Порядок следования полей при просмотре
  CurrentDocFieldOrder = "DocID, DocIDParent, Name, Author, DateActivation, DateCompletion, Content, NameResponsible, Correspondent, NameControl, UserFieldText1"
    If InStr(UCase(Session("Department")), UCase(SIT_MINC)) = 1 Then
    CurrentDocFieldOrder = "DocID, DocIDParent, Name, Author, DateActivation, DateCompletion, Content, NameResponsible, Correspondent, NameControl, UserFieldText1, UserFieldText8"
  end if
  
'Запрос №11 - СТС - start
'ElseIf InStr(UCase(S_ClassDoc),UCase(SIT_DOGOVORI)) = 1 Then
ElseIf InStr(UCase(S_ClassDoc),UCase(SIT_DOGOVORI_OLD)) = 1 Then
'Запрос №11 - СТС - end
  'Порядок следования полей при просмотре
  CurrentDocFieldOrder = "DocID, DocIDAdd, Author, NameResponsible, UserFieldText4, UserFieldDate1, PartnerName, UserFieldText1, Description, Name, DocIDIncoming, QuantityDoc, UserFieldText8, InventoryUnit, AmountDoc, UserFieldMoney1, Currency, UserFieldText5, UserFieldText6, UserFieldText7, UserFieldDate4, UserFieldDate5, UserFieldDate6, Content, ListToReconcile, NameAproval, UserFieldText2, UserFieldText3 "	
'Запрос №11 - СТС - start
ElseIf InStr(UCase(S_ClassDoc),UCase(SIT_DOGOVORI_NEW)) = 1 Then
  'Порядок следования полей при просмотре
'ph - 20120216 - start
'  CurrentDocFieldOrder = "DocID, Author, NameResponsible, UserFieldText4, PartnerName, UserFieldText1, Description, Name, UserFieldText8, UserFieldText2, InventoryUnit, AmountDoc, Currency, UserFieldText5, UserFieldDate4, UserFieldDate6, ListToReconcile, NameAproval, SecurityLevel, Department, UserFieldText3, DocIDParent, UserFieldDate5, ContractType "	
  CurrentDocFieldOrder = "DocID, Author, NameResponsible, UserFieldText4, PartnerName, UserFieldText1, Description, Name, UserFieldText8, UserFieldText2, InventoryUnit, AmountDoc, Currency, UserFieldText5, UserFieldDate4, UserFieldDate6, ListToReconcile, NameAproval, SecurityLevel, Department, UserFieldText3, DocIDParent, UserFieldDate5, ContractType, UserFieldText6, UserFieldText7, AddFieldText1"
'ph - 20120216 - end
'Запрос №11 - СТС - end
'20090622 - Заявка ТКП
ElseIf InStr(UCase(S_ClassDoc),UCase(SIT_COM_OFFERS)) = 1 Then
  'Порядок следования полей при просмотре
  CurrentDocFieldOrder = "DocID, UserFieldText1, NameAproval, Name, Content, PartnerName, Author, ListToView, UserFieldText2, ListToReconcile, DateActivation, "
ElseIf InStr(UCase(S_ClassDoc), UCase(STS_PaymentOrder)) = 1 Then
  'Показ расчетного поля с курсом валюты
  ReDim AdditionalCalculatedFieldNames(1)
  ReDim AdditionalCalculatedFieldFormulas(1)
  AdditionalCalculatedFieldNames(1) = SIT_CurrencyRateToUSD
  AdditionalCalculatedFieldFormulas(1) = STS_CurrencyRateFormula
  'Порядок следования полей при просмотре
  CurrentDocFieldOrder = "DocID, DocIDParent, Author, NameResponsible, PartnerName, Description, AmountDoc, Currency, UserFieldMoney1, "+AdditionalCalculatedFieldNames(1)+", ListToReconcile, NameAproval, Department, UserFieldText1, BusinessUnit, UserFieldText3, UserFieldText4, UserFieldText5, UserFieldText6, UserFieldText8, DateCompletion, ListToView"
  'Шаблоны MS Word
  VAR_DocTemplateFilename = "PaymentOrder_RU.doc"
ElseIf InStr(UCase(S_ClassDoc), UCase(STS_PurchaseOrder)) = 1 Then
  'Показ расчетного поля с курсом валюты
  ReDim AdditionalCalculatedFieldNames(1)
  ReDim AdditionalCalculatedFieldFormulas(1)
  AdditionalCalculatedFieldNames(1) = SIT_CurrencyRateToUSD
  AdditionalCalculatedFieldFormulas(1) = STS_CurrencyRateFormula
  
  'rmanyushin 119191 17.08.2010 Start
  'Порядок следования полей при просмотре
  'CurrentDocFieldOrder = "DocID, Author, NameResponsible, PartnerName, Description, AmountDoc, Currency, UserFieldMoney1, "+AdditionalCalculatedFieldNames(1)+", ListToReconcile, NameAproval, Department, UserFieldText1, BusinessUnit, UserFieldText3, UserFieldText4, UserFieldText5, UserFieldText6, UserFieldText8, DateCompletion, ListToView"
    CurrentDocFieldOrder = "DocID, Author, NameResponsible, PartnerName, Description, Content, AmountDoc, Currency, UserFieldMoney1, "+AdditionalCalculatedFieldNames(1)+", ListToReconcile, NameAproval, Department, UserFieldText1, BusinessUnit, UserFieldText3, UserFieldText4, UserFieldText5, UserFieldText6, UserFieldText8, DateCompletion, ListToView"
  'rmanyushin 119191 17.08.2010 End
  
  'Шаблоны MS Word
  VAR_DocTemplateFilename = "PurchaseOrder_RU.doc"
End If

'Запрос №11 - СТС - start
'Отдельный лист согласования
If InStr(UCase(S_ClassDoc),UCase(SIT_DOGOVORI_NEW)) = 1 Then
  VAR_StandardFileNameReconciliationList = "ReconciliationListContractRUS.doc"
  Else if InStr(UCase(S_ClassDoc),UCase(SIT_CONTRACTS_MC)) = 1 then
  VAR_StandardFileNameReconciliationList = "ReconciliationListRUSContractsHQ.doc"

Else
  'VAR_StandardFileNameReconciliationList = ""
End If
End If
'Запрос №11 - СТС - end

' порядок следования полей при редактировании/создании
'    If InStr(UCase(S_ClassDoc),UCase(SIT_VHODYASCHIE)) > 0 Then 
'      CurrentDocFieldOrder = "DocID, Name, UserFieldText1, UserFieldText7, DateActivation, UserFieldText2, ListToView, PartnerName, UserFieldText3, DocIDIncoming, UserFieldDate1, UserFieldText4, UserFieldText5, UserFieldText6, Content"
'    End If
'
'    'SAY 2008-10-27
'    If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_VHODYASCHIE_ACC)) > 0 Then 
'      CurrentDocFieldOrder = "DocID, Name, UserFieldText1, UserFieldText7, DateActivation, UserFieldText2, ListToView, PartnerName, UserFieldText3, DocIDIncoming, UserFieldDate1, UserFieldText4, UserFieldText5, UserFieldText6, Content"
'    End If
'
'    If InStr(UCase(S_ClassDoc),UCase(SIT_ISHODYASCHIE)) > 0 Then 
'      CurrentDocFieldOrder = "DocID, UserFieldText1, DocIDParent, UserFieldDate1, NameAproval, Name, Content, PartnerName, UserFieldText5, UserFieldText3, UserFieldText4, Author, ListToReconcile, DateActivation"
'      ' SAY 2008-12-03 меняем UserFieldText на Correspondents
'      CurrentDocFieldOrder = "DocID, UserFieldText1, DocIDParent, UserFieldDate1, NameAproval, Name, Content, PartnerName, Correspondent, UserFieldText3, UserFieldText4, Author, ListToReconcile, DateActivation"
'    End If
'
'    If InStr(UCase(S_ClassDoc),UCase(SIT_SLUZH_ZAPISKA)) > 0 Then
'      'SAY 2008-10-08 добавлено поле основной резолюции
'  '   CurrentDocFieldOrder = "Name, UserFieldText1, DocIDParent, UserFieldDate1, DocID, Author, DateActivation, Content, ListToReconcile, NameAproval, LocationPath, ListToView, Resolution, SecurityLevel, Department "
'      CurrentDocFieldOrder = "DocID, DocIDParent, Name, Content, ListToView, NameAproval, Author, ListToReconcile, DateActivation, Resolution, UserFieldDate2, UserFieldText1, UserFieldText2, UserFieldText3, UserFieldText4, UserFieldText5, UserFieldText6"
'    End If
'
'    If InStr(UCase(S_ClassDoc),UCase(SIT_RASP_DOCS)) > 0 Then
'      CurrentDocFieldOrder = "DocID, Name, UserFieldText1, DateActivation, Content, UserFieldText2, Author, ListToReconcile, NameAproval, ListToView, LocationPath, DocIDParent, SecurityLevel"	
'    End If
'
'    If InStr(UCase(S_ClassDoc),UCase(SIT_NORM_DOCS)) > 0 Then
'      CurrentDocFieldOrder = "DocID, Name, UserFieldText1, UserFieldDate1, Author, DateActivation, Content, Description, UserFieldText3, UserFieldText4, UserFieldDate2 ,UserFieldText2, ListToReconcile, NameAproval, LocationPath, ListToView, SecurityLevel, Department "	
'    End If
'
'    If InStr(UCase(S_ClassDoc),UCase(SIT_ZADACHI)) > 0 Then
'      'SAY 2008-10-23 
'      'CurrentDocFieldOrder = "DocID, DocIDParent, Name, UserFieldText1, Author, DateActivation, DateCompletion, Content, Rank, NameResponsible, ListToView, Context, NameControl, SecurityLevel, Department "	
'      CurrentDocFieldOrder = "DocID, DocIDParent, Name, Author, DateActivation, DateCompletion, Content, NameResponsible, Correspondent, NameControl, UserFieldText1"
'    End If
'
'    If InStr(UCase(S_ClassDoc),UCase(SIT_DOGOVORI)) > 0 Then
'      CurrentDocFieldOrder = "DocID, DocIDAdd, Author, NameResponsible, UserFieldText4, UserFieldDate1, PartnerName, UserFieldText1, Description, Name, DocIDIncoming, QuantityDoc, UserFieldText8, InventoryUnit, AmountDoc, UserFieldMoney1, Currency, UserFieldText5, UserFieldText6, UserFieldText7, UserFieldDate4, UserFieldDate5, UserFieldDate6, Content, ListToReconcile, NameAproval, UserFieldText2, UserFieldText3 "	
'    End If
'
'Ph - Start - 20080922 - ПОТОМ СДЕЛАТЬ ЧЕРЕЗ ПЕРЕМЕННЫЕ
'If InStr(UCase(S_Department),UCase(SIT_SITRONICS)) = 1 then
'  If InStr(UCase(S_ClassDoc), UCase(SIT_RASP_DOCS)) = 1 Then
'	Select case S_UserFieldText1
'		case SIT_Orders_Prikaz
'			VAR_DocTemplateFilename = "Prikaz_RU.doc"
'		case SIT_Orders_Rasporyajenie
'			VAR_DocTemplateFilename = "Raspor_RU.doc"
'	End Select
'  End If
'  If InStr(UCase(S_ClassDoc), UCase(SIT_ISHODYASCHIE)) = 1 Then
'	Select case Session("Lang")
'		case "RUS"
'			VAR_DocTemplateFilename = "Letter_Sistema.doc"
'		case ""
'			VAR_DocTemplateFilename = "Letter_Foreign.doc"
'	End Select
'  End If
'  If InStr(UCase(S_ClassDoc), UCase(SIT_SLUZH_ZAPISKA)) = 1 Then
'	VAR_DocTemplateFilename = "SluzZap.doc"
'  End If
'ElseIf InStr(UCase(S_Department),UCase(SIT_STS_RU)) = 1 Then
'  If InStr(UCase(S_ClassDoc), UCase(SIT_RASP_DOCS)) = 1 Then
'	Select case S_UserFieldText1
'		case SIT_Orders_Prikaz
'			VAR_DocTemplateFilename = "Prikaz_STS_RU.doc"
'		case SIT_Orders_Rasporyajenie
'			VAR_DocTemplateFilename = "Raspor_STS_RU.doc"
'	End Select
'  End If
'  If InStr(UCase(S_ClassDoc), UCase(SIT_ISHODYASCHIE)) = 1 Then
'	Select case Session("Lang")
'		case "RUS"
'			VAR_DocTemplateFilename = "Letter_Sistema_STS.doc"
'		case ""
'			VAR_DocTemplateFilename = "Letter_Foreign_STS.doc"
'	End Select
'  End If
'  If InStr(UCase(S_ClassDoc), UCase(SIT_SLUZH_ZAPISKA)) = 1 Then
'	VAR_DocTemplateFilename = "SluzZap_STS.doc"
'  End If
''Ph - 20090201 - Start
'  If InStr(UCase(S_ClassDoc), UCase(STS_PaymentOrder)) = 1 Then
'    VAR_DocTemplateFilename = "PaymentOrder_RU.doc"
'  End If
'  If InStr(UCase(S_ClassDoc), UCase(STS_PurchaseOrder)) = 1 Then
'    VAR_DocTemplateFilename = "PurchaseOrder_RU.doc"
'  End If
''Ph - 20090201 - End
'End If
''Ph - End - 20080922
'
'Ph - 20080918 - start
''В служебных записках убрать возможность изменения даты подписи-утверждения (не должно запрашивать)
'If InStr(UCase(S_ClassDoc), UCase(SIT_SLUZH_ZAPISKA)) = 1 Then
'  VAR_NotToAskDateDuringApproval="Y"
'End If
''Ph - 20080918 - end

'Phil - VV
'20080918 - Может быть уже не нужно, механизм выдачи доступа изменился, надо проверить
'20090318 - точно не нужно, уровень доступа во всех документах поставлен только для лиц...
'addlogd "@@@Session(""OtherCompanies""): "+Session("OtherCompanies")
'addlogd "@@@dsDoc(""Department""): "+dsDoc("Department")
'If InStr(UCase(Session("OtherCompanies")) + vbCrLf, UCase(dsDoc("Department")) + vbCrLf) > 0 Then
'Addlogd "@@@WORK"
'  VAR_ReadAccess = "Y"
'End If
'Phil - VV

'SAY 2008-10-08
'AddLogD "W2W: S_ListToView="+S_ListToView+", UserID="+"<"+Session("UserID")+">"
   VAR_CanMakeMainResolution=""
'SAY 2008-10-17 меняем поле ListToView на Correspondent 2008-10-23 меняем обратно...
   If InStr(S_ListToView, "<"+Session("UserID")+">") > 0 and InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_SLUZH_ZAPISKA)) >0 Then
   'If InStr(S_Correspondent, "<"+Session("UserID")+">") > 0 and InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_SLUZH_ZAPISKA)) >0 Then
      VAR_CanMakeMainResolution="Y"
      VAR_MainResolutionChecked="Y"
'amw 12-12-2013 включение кнопки "Резолюция" для "Закупки Микрон/Заявка на закупку" и "Протокол ЗК"
   ElseIf (InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_PURCHASE_ORDER)) > 0 or _
           InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_PROTOCOL)) > 0 ) and _
          InStr(UCase(ReplaceRoleFromDir(MIKRON_HeadKFIE,SIT_MIKRON)),UCase(Session("UserID"))) > 0 Then
      VAR_CanMakeMainResolution="Y"
      VAR_MainResolutionChecked="Y"
      VAR_MainResolutionNoDates="Y"
      VAR_CanMakeDocCompleted="Y"
'amw 12-12-2013
   Else
      VAR_ButtonsNotToShow=VAR_ButtonsNotToShow + "ClickCreateCommentResolution,ClickResolution,"
   End If

'   ' *** ЗАЯВКИ ДЛЯ СТС
'    If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) > 0 Then
'      ReDim AdditionalCalculatedFieldNames(1)
'      ReDim AdditionalCalculatedFieldFormulas(1)
'	  AdditionalCalculatedFieldNames(1) = SIT_CurrencyRateToUSD
'      AdditionalCalculatedFieldFormulas(1) = STS_CurrencyRateFormula
''      CurrentDocFieldOrder = "DocID, DocIDParent, Author, NameResponsible, PartnerName, Description, AmountDoc, Currency, UserFieldMoney1, "+AdditionalCalculatedFieldNames(1)+", ListToReconcile, NameAproval, Department, UserFieldText1, UserFieldText2, UserFieldText3, UserFieldText4, UserFieldText5, UserFieldText6, UserFieldText8, DateCompletion, ListToView"
'      CurrentDocFieldOrder = "DocID, DocIDParent, Author, NameResponsible, PartnerName, Description, AmountDoc, Currency, UserFieldMoney1, "+AdditionalCalculatedFieldNames(1)+", ListToReconcile, NameAproval, Department, UserFieldText1, BusinessUnit, UserFieldText3, UserFieldText4, UserFieldText5, UserFieldText6, UserFieldText8, DateCompletion, ListToView"
'    End If
'    If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) > 0 Then
'      ReDim AdditionalCalculatedFieldNames(1)
'      ReDim AdditionalCalculatedFieldFormulas(1)
'	  AdditionalCalculatedFieldNames(1) = SIT_CurrencyRateToUSD
'      AdditionalCalculatedFieldFormulas(1) = STS_CurrencyRateFormula
''      CurrentDocFieldOrder = "DocID, Author, NameResponsible, PartnerName, Description, AmountDoc, Currency, UserFieldMoney1, "+AdditionalCalculatedFieldNames(1)+", ListToReconcile, NameAproval, Department, UserFieldText1, UserFieldText2, UserFieldText3, UserFieldText4, UserFieldText5, UserFieldText6, UserFieldText8, DateCompletion, ListToView"
'      CurrentDocFieldOrder = "DocID, Author, NameResponsible, PartnerName, Description, AmountDoc, Currency, UserFieldMoney1, "+AdditionalCalculatedFieldNames(1)+", ListToReconcile, NameAproval, Department, UserFieldText1, BusinessUnit, UserFieldText3, UserFieldText4, UserFieldText5, UserFieldText6, UserFieldText8, DateCompletion, ListToView"
'    End If

'SAY 2008-10-21 разрешить регистратору загружать основные версии файлов
If (InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_RASP_DOCS)) > 0 or InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_NORM_DOCS)) > 0 ) and (InStr(Session("UserID"),"registrator") > 0 or CheckPermit(Session("Permitions"),"REGISTRAR")) Then
  VAR_AgreeAgainInitiallyNotChecked="Y"
  VAR_UploadFileForcesAgreeAgain="N"
  VAR_CanCreateMainVersionFiles=True
End If

'Заявки на закупку и оплату Исполнитель может отмечать исполненными
If (InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) > 0) or (InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) > 0) Then
  VAR_ResponsiblePersonCanMarkDocAsCompleted = "Y"
Else
  VAR_ResponsiblePersonCanMarkDocAsCompleted = ""
End If

'rmanyushin 142913 08-11-2010 Start
'If (InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) > 0) Then
'	If S_NameAproval<>"" and InStr(S_NameAproval, "<"+Session("UserID")+">")>0 Then
'		VAR_ButtonsNotToShow=VAR_ButtonsNotToShow+"ClickMakeCompleted, "
'	End If
'End If
'rmanyushin 142913 08-11-2010 End

'Инициатор может добавлять согласующих
If InStr(Session("UserID"), Request("Author")) > 0 Then
  VAR_ThisUserCanAddVisa = "Y"
Else
  VAR_ThisUserCanAddVisa = ""
End If

'If InStr(dsDoc("NameCreation"), "<"+Session("UserID")+">") > 0 or InStr(dsDoc("Author"), "<"+Session("UserID")+">") > 0 Then
'  ShowDoc_ShowButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickVisaAdd"
'End If

'Запрос №17 - СТС - start
'Запоминаем подразделение документа (для рассылок, зависимых от БН)
Session("CurrentDepartmentDoc") = S_Department
'Запрос №17 - СТС - end

'Запрос №33 - СТС - start
'Изначально возобновление отмененных документов разрешено
bHideClickMakeCanceledCancel = False
'Если после анализа комментариев в UserShowListComments.asp выяснится, что документ отменен автоматически, то будет установлено в True
'Запрос №33 - СТС - end

'ph - 20111109 - start
If InStr(dsDoc("LocationPath"), ">+") > 0 Then
  If InStr(Session("CurrentClassDoc"), SIT_VHODYASCHIE) > 0 or InStr(Session("CurrentClassDoc"), SIT_ISHODYASCHIE) > 0 or InStr(Session("CurrentClassDoc"), SIT_RASP_DOCS) > 0 Then
    ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickMakeCanceled"
  End If
End If
'ph - 20111109 - end

if InStr(UCase(S_Department), UCase(SIT_RTI)) = 1 then
  VAR_StandardFileNameReconciliationList = "ReconciliationListRUS_RTI.doc"
    if InStr(Session("CurrentClassDoc"), RTI_CONTRACT) = 1 Then
    VAR_StandardFileNameReconciliationList = "ReconciliationListRUSContract_RTI.doc"
  End If
End If

'vtss лист согласования
if InStr(UCase(S_Department), UCase(SIT_VTSS)) = 1 then
  VAR_StandardFileNameReconciliationList = "ReconciliationListRUS_VTSS.doc"
End If
'vtss лист согласования

If InStr(UCase(S_Department), UCase(SIT_MIKRON)) = 1 Then
   VAR_StandardFileNameReconciliationList = "RL_Mikron.doc"
   If InStr(Session("CurrentClassDoc"), MIKRON_CONTRACT) = 1 or _
      InStr(Session("CurrentClassDoc"), MIKRON_S_CONTRACT) = 1 or _ 
      InStr(Session("CurrentClassDoc"), MIKRON_EXPORT_CONTRACT) = 1 or _ 
      InStr(Session("CurrentClassDoc"), MIKRON_EXPADD_CONTRACT) = 1 or _ 
      InStr(Session("CurrentClassDoc"), MIKRON_ADD_CONTRACT) = 1 Then
      VAR_StandardFileNameReconciliationList = "RL_MikronContract.doc"
   End If
   If InStr(UCase(Request.ServerVariables("URL")),UCase("/ShowDoc.asp"))>0 Then
      If dsDoc("ClassDoc")=MIKRON_CONTRACT or dsDoc("ClassDoc")=MIKRON_ADD_CONTRACT Then
         If InStr(Session("UserComment"),UCase("Финансовый Контролер"))>0 Then
            VAR_ReadAccess ="Y"
         End If
      End If
   End If
End If

if InStr(UCase(S_Department), UCase(SIT_MINC)) = 1 then
  VAR_StandardFileNameReconciliationList = "ReconciliationListRUS_MINC.doc"
End If



'ph - 20120311 - start
S_AdditionalUsers = Trim(MyCStr(dsDoc("AdditionalUsers")))
'ph - 20120311 - end

'kkoshkin 04032014
                    But_AGREE = "Согласовать"
                    DOCS_AGREE = "Согласовать документ"

			If InStr(UCase(Session("CurrentClassDoc")),UCase(RTI_PAYMENT_ORDER)) > 0  Then 
			   	
                Set dsDocVNIK2 = Server.CreateObject("ADODB.Recordset")
                sSQL = "select UserFieldText6 from Docs where Docid = N'" + Trim(Request("DocID")) + "'"
                dsDocVNIK2.CursorLocation = 3
                dsDocVNIK2.Open sSQL, Conn, 3, 1, &H1
                If not dsDocVNIK2.EOF Then
                    var_DogovorSdanNaXranenie = UCase(Trim(dsDocVNIK2("UserFieldText6").Value))
                Else
                    var_DogovorSdanNaXranenie = ""
                End If
                dsDocVNIK2.Close
    
                If var_DogovorSdanNaXranenie = "НЕТ" Then
                    But_AGREE = "Согласовать(!)"
                    DOCS_AGREE = "Согласовать документ БЕЗ ДОГОВОРА (НЕ СДАН НА ХРАНЕНИЕ В БУХГАЛТЕРИЮ)"
                End If
			End If
'kkoshkin 04032014

%>
