<%
'Set output columns for document lists
'Variables:
'CurrentClassDoc - current document category
'Document variables (each must set from 0 to 8, if set to 0 it means not to show this document field):
'nCol_DocID
'nCol_DocIDadd
'nCol_DocIDParent
'nCol_DocIDPrevious
'nCol_DocIDIncoming
'nCol_Author
'nCol_Correspondent
'nCol_Resolution
'nCol_History
'nCol_Result
'nCol_PercentCompletion
'nCol_Department
'nCol_Name
'nCol_Description
'nCol_LocationPaper
'nCol_Currency
'nCol_CurrencyRate
'nCol_Rank
'nCol_ExtInt
'nCol_PartnerName
'nCol_StatusDevelopment
'nCol_StatusArchiv
'nCol_StatusCompletion
'nCol_StatusPayment
'nCol_TypeDoc
'nCol_ClassDoc
'nCol_ActDoc
'nCol_InventoryUnit
'nCol_PaymentMethod
'nCol_AmountDoc
'nCol_QuantityDoc
'nCol_SecurityLevel
'nCol_DateActivation
'nCol_DateCreation
'nCol_DateCompletion
'nCol_DateCompleted
'nCol_DateExpiration
'nCol_DateSigned
'nCol_NameCreation
'nCol_NameAproval
'nCol_NameApproved
'nCol_DateApproved
'nCol_NameControl
'nCol_ListToEdit
'nCol_ListToView
'nCol_ListToReconcile
'nCol_ListReconciled
'nCol_NameResponsible
'nCol_NameLastModification
'nCol_DateLastModification
'nCol_UserFieldText1
'nCol_UserFieldText2
'nCol_UserFieldText3
'nCol_UserFieldText4
'nCol_UserFieldText5
'nCol_UserFieldText6
'nCol_UserFieldText7
'nCol_UserFieldText8
'nCol_UserFieldMoney1
'nCol_UserFieldMoney2
'nCol_UserFieldDate1
'nCol_UserFieldDate2
'nCol_IsActive
'nCol_Content
'
'Example:
Select Case CurrentClassDoc
'    Case "Invoices" ' - document category to be processed
'			nCol_DocID = 1 			'Show document ID in the first column of the document list table
'			nCol_NameCreation = 2	'Show document creator name in the second column of the document list table
'			nCol_DateExpiration = 0 'Not to show expiration date	
'			sColAlign(1)="left"		'Set column 1 align value to "left"
'			sColWidth(1)="40"			'Set column 1 width value to "40"
'    Case "???" ' - other directory GUID to be processed
'    			'....
End Select

'Set view for Bank account list
If CurrentClassDoc="Accounts" Or Request("ActDoc")="Bank accounts" Then ' - document category to be processed
			nCol_LocationPaper=0
			nCol_Description=0
			nCol_UserFieldText1=0
			nCol_UserFieldText2=0
			nCol_UserFieldText3=0
			nCol_UserFieldText4=0
			nCol_UserFieldText5=0
			nCol_UserFieldText6=0
			nCol_UserFieldText7=0
			nCol_UserFieldText8=0
			nCol_DocIDParent=0
			nCol_DateCompletion=0
End If

sColWidth(1)="120"
sColWidth(2)="200"
sColWidth(3)="110"
sColWidth(4)="110"
sColWidth(5)="110"
sColWidth(6)="110"
sColWidth(7)="110"

'nCol_Result=1

If IsHelpDeskDoc() Then
	nCol_DocID=1
	nCol_DocIDParent=0
	nCol_Name=2
	nCol_Description=2
	nCol_DateCreation=3
	nCol_DateCompletion=3
	nCol_Department=4
	nCol_NameResponsible=4
	nCol_Correspondent=5
	nCol_DateActivation=0
	nCol_ListToView=0
	nCol_AmountDoc=0
	nCol_NameAproval=0
	nCol_ListToReconcile=0
	nCol_LocationPaper=0
	nCol_NameCreation=0
	nCol_PartnerName=0
	nCol_UserFieldText1=0

	'If IsHelpDeskAdminOrConsultant() Then
		nCol_NameResponsible=4
		NColumns=5
	'Else
	'	nCol_NameResponsible=0
	'	NColumns=3
	'End If
	sColWidth(5)="250"
End If

If CurrentClassDoc=DOCS_Notices Then
	sColWidth(1)="130"
	sColWidth(2)="400"
	sColWidth(3)="250"
	sColWidth(4)="20"
	NColumns=4
	nCol_LocationPaper=0
End If

'If InStr(UCase(CurrentClassDoc), UCase("РџСЂРѕРїСѓСЃРєР°"))>0 Or InStr(UCase(Request("ClassDoc")), UCase("РџСЂРѕРїСѓСЃРєР°"))>0 Then 
'	nCol_DocID=1
'	nCol_DateActivation=0
'	nCol_DateCreation=0
'	nCol_DocIDParent=0
'	nCol_Name=0
'	nCol_LocationPaper=0
'	nCol_DateActivation=0
'	nCol_PartnerName=0
'	nCol_ListToReconcile=0
'	nCol_AmountDoc=0
'	nCol_Department=4
'	nCol_NameResponsible=5
'	nCol_NameAproval=0
'	NColumns=5
'End If

'If False Then 'test
'	nCol_Name=3
'	nCol_Name=0
'
'	nCol_Description=3
'	nCol_DateActivation=2
'	nCol_DateCreation=2
'	nCol_LocationPaper=0
'	nCol_BusinessProcessStep=0
'	nCol_Resolution=0
'	nCol_FileNameNameLastModification=0
'	nCol_DateCompletion=0
'	sColWidth(2)="10"
'	nCol_History=0
'	nCol_DateActivation=0
'	nCol_PartnerName=0
'	nCol_ListToReconcile=0
'	nCol_AmountDoc=0
'	nCol_Department=0
'	nCol_NameResponsible=0
'	nCol_NameAproval=0
'End If

'If Request("depdocs") = "y" Then
'If False Then
'	nCol_DocID = 0
'	nCol_DocIDadd = 0
'	nCol_DocIDParent = 0
'	nCol_DocIDPrevious = 0
'	nCol_DocIDIncoming = 0
'	nCol_Author = 0
'	nCol_Correspondent = 0
'	nCol_Resolution = 0
'	nCol_History = 0
'	nCol_Result = 0
'	nCol_PercentCompletion = 0
'	nCol_Department = 0
'	nCol_Name = 0
'	nCol_Description = 0
'	nCol_LocationPaper = 0
'	nCol_Currency = 0
'	nCol_CurrencyRate = 0
'	nCol_Rank = 0
'	nCol_ExtInt = 0
'	nCol_PartnerName = 0
'	nCol_StatusDevelopment = 0
'	nCol_StatusArchiv = 0
'	nCol_StatusCompletion = 0
'	nCol_StatusPayment = 0
'	nCol_TypeDoc = 0
'	nCol_ClassDoc = 0
'	nCol_ActDoc = 0
'	nCol_InventoryUnit = 0
'	nCol_PaymentMethod = 0
'	nCol_AmountDoc = 0
'	nCol_QuantityDoc = 0
'	nCol_SecurityLevel = 0
'	nCol_DateActivation = 0
'	nCol_DateCreation = 0
'	nCol_DateCompletion = 0
'	nCol_DateCompleted = 0
'	nCol_DateExpiration = 0
'	nCol_DateSigned = 0
'	nCol_NameCreation = 0
'	nCol_NameAproval = 0
'	nCol_NameApproved = 0
'	nCol_DateApproved = 0
'	nCol_NameControl = 0
'	nCol_ListToEdit = 0
'	nCol_ListToView = 0
'	nCol_ListToReconcile = 0
'	nCol_ListReconciled = 0
'	nCol_NameResponsible = 0
'	nCol_NameLastModification = 0
'	nCol_DateLastModification = 0
'	nCol_UserFieldText1 = 0
'	nCol_UserFieldText2 = 0
'	nCol_UserFieldText3 = 0
'	nCol_UserFieldText4 = 0
'	nCol_UserFieldText5 = 0
'	nCol_UserFieldText6 = 0
'	nCol_UserFieldText7 = 0
'	nCol_UserFieldText8 = 0
'	nCol_UserFieldMoney1 = 0
'	nCol_UserFieldMoney2 = 0
'	nCol_UserFieldDate1 = 0
'	nCol_UserFieldDate2 = 0
'	nCol_IsActive = 0
'	nCol_Content = 0
'
'    nCol_DocID = 1
'    nCol_UserFieldText2 = 2
'    nCol_UserFieldText6 = 3
'    nCol_UserFieldText7 = 3
'    nCol_UserFieldText8 = 4
'
'	'S_NameUserFieldText2="Test2"
'	'S_NameUserFieldText6="Test6"
'	'S_NameUserFieldText7="Test7"
'	'S_NameUserFieldText8="Test8"
'
'	S_ClassDoc="РџР»Р°С‚РµР¶Рё"
'
'    NColumns = 4
'
'	sColWidth(1)="100"
'	sColWidth(2)="200"
'	sColWidth(3)="300"
'	sColWidth(4)="200"
'End If

'If InStr(UCase(CurrentClassDoc), UCase("Протокол"))>0 Then 
'  ClearFieldsnCol
'  select case Request("l")
'    case ""
'      AddFields_Names(2) = "Description"
'      nCol_AddFields(2) = 2
'    case "ru"
'      AddFields_Names(1) = "Краткое содержание"
'      nCol_AddFields(1) = 2
'    case "3"
'      nCol_Description=2
'  end select
'  nCol_DocID=1
'  nCol_DocIDParent=1
'  nCol_Name=2
'  nCol_LocationPaper=2
'  nCol_DateActivation=3
'  nCol_DateCompletion=3
'  nCol_Department=4
'  nCol_PartnerName=4
'  NColumns=4
'End If

'SAY 2008-10-30  Служебные записки
If InStr(UCase(CurrentClassDoc), UCase(SIT_SLUZH_ZAPISKA))>0 Then 
  ClearFieldsnCol
  NColumns=5
  
  nCol_DocID=1
  nCol_DocIDParent=1
  sColWidth(1)="80"
  
  nCol_Name=2
  sColWidth(2)="150"
  
  nCol_DateActivation=3
  sColWidth(3)="50"
  
  nCol_Author=4
  sColWidth(4)="80"
  
'  NameApproved_Names(5) = "Подписант"
  nCol_NameAproval=5
  sColWidth(5)="80"
End If

'SAY 2008-10-30  Поручения
If InStr(UCase(CurrentClassDoc), UCase(SIT_ZADACHI))>0 Then 
  ClearFieldsnCol
  NColumns=5
  
  nCol_DocID=1
  sColWidth(1)="80"
  
  nCol_UserFieldText1=2
  nCol_Name=2
  nCol_UserFieldText2=2
  sColWidth(2)="150"
  
  nCol_DateActivation=3
  sColWidth(3)="50"
  
  nCol_Author=4
  sColWidth(4)="80"
  
  nCol_NameResponsible=5
  sColWidth(5)="80"
End If
'поручения Минц
If InStr(UCase(CurrentClassDoc), UCase(SIT_ZADACHI))>0 and InStr(UCase(Session("Department")), UCase(SIT_MINC)) = 1 Then 
  ClearFieldsnCol
  NColumns=7
  
  nCol_DocID=1
  sColWidth(1)="80"
  
  nCol_UserFieldText1=2
  nCol_Name=2
  nCol_Content=2
  nCol_UserFieldText8=2
  sColWidth(2)="150"
  
  nCol_UserFieldText8=3
  sColWidth(3)="50"
  
  nCol_DateActivation=4
  sColWidth(4)="50"
  
  nCol_DateCompletion=5
  sColWidth(5)="50"

  nCol_Author=6
  sColWidth(6)="80"
  nCol_NameResponsible=7
  sColWidth(7)="80"
End If

'SAY 2008-10-30  Распорядительные документы
If InStr(UCase(CurrentClassDoc), UCase(SIT_RASP_DOCS))>0 Then 
  ClearFieldsnCol
  NColumns=5
  
  nCol_DocID=1
  sColWidth(1)="80"
  
  nCol_UserFieldText1=2
  nCol_Name=2
  sColWidth(2)="150"
  
  nCol_DateActivation=3
  sColWidth(3)="50"
  
  nCol_Author=4
  sColWidth(4)="80"
  
  nCol_NameAproval=5
  sColWidth(5)="80"
End If

'vnik_protocols
If InStr(UCase(CurrentClassDoc), UCase(SIT_PROTOCOLS))>0 or InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_PROTOCOLS_CPC)) > 0 Then 
  ClearFieldsnCol
  NColumns=5
  
  nCol_DocID=1
  sColWidth(1)="80"
  
  nCol_UserFieldText1=2
  nCol_Name=2
  sColWidth(2)="150"
  
  nCol_DateActivation=3
  sColWidth(3)="50"
  
  nCol_Author=4
  sColWidth(4)="80"
  
  nCol_NameAproval=5
  sColWidth(5)="80"
End If
'vnik_protocols

'vnik_payment_order
If InStr(UCase(CurrentClassDoc), UCase(SIT_PAYMENT_ORDER))>0 Then 
  ClearFieldsnCol
  NColumns=5
   
  nCol_DocID=1
  sColWidth(1)="80"
  
  nCol_UserFieldText1=2
  nCol_Name=2
  sColWidth(2)="150"
  
  nCol_DateActivation=3
  sColWidth(3)="50"
  
  nCol_UserFieldMoney1=4
  sColWidth(4)="80"
  
  nCol_Author=5
  sColWidth(5)="80"
  
  nCol_NameAproval=6
  sColWidth(6)="80"
End If
'vnik_payment_order

'vnik_purchase_order
If InStr(UCase(CurrentClassDoc), UCase(SIT_PURCHASE_ORDER))>0 Then 
  ClearFieldsnCol
  NColumns=5
   
  nCol_DocID=1
  sColWidth(1)="80"
  
  nCol_UserFieldText1=2
  nCol_Name=2
  sColWidth(2)="150"
  
  nCol_DateActivation=3
  sColWidth(3)="50"
  
  nCol_UserFieldMoney1=4
  sColWidth(4)="80"
  
  nCol_Author=5
  sColWidth(5)="80"
  
  nCol_NameAproval=6
  sColWidth(6)="80"
End If
'vnik_purchase_order

'vnik_contracts
If InStr(UCase(CurrentClassDoc), UCase(SIT_CONTRACTS_MC))>0 Then 
  ClearFieldsnCol
  NColumns=5
   
  nCol_DocID=1
  sColWidth(1)="80"
  
  nCol_UserFieldText1=2
  nCol_Name=2
  sColWidth(2)="150"
  
  nCol_DateActivation=3
  sColWidth(3)="50"
  
  nCol_UserFieldMoney1=4
  sColWidth(4)="80"
  
  nCol_Author=5
  sColWidth(5)="80"
  
  nCol_NameAproval=6
  sColWidth(6)="80"
End If
'vnik_contracts

'SAY 2008-10-30  Нормативные документы
If InStr(UCase(CurrentClassDoc), UCase(SIT_NORM_DOCS))>0 Then 
  ClearFieldsnCol
  NColumns=5
  
  nCol_DocID=1
  sColWidth(1)="80"
  
  nCol_UserFieldText1=2
  nCol_Name=2
  sColWidth(2)="150"
  
  nCol_DateActivation=3
  sColWidth(3)="50"
  
  nCol_Author=4
  sColWidth(4)="80"
  
  nCol_NameAproval=5
  sColWidth(5)="80"
End If

'SAY 2008-10-30  Входящие
If InStr(UCase(CurrentClassDoc), UCase(SIT_VHODYASCHIE))>0 Then 
  ClearFieldsnCol
  NColumns=5
  
  nCol_DocID=1
  sColWidth(1)="80"
  
  nCol_Name=2
  sColWidth(2)="150"
  
  nCol_DateActivation=3
  sColWidth(3)="50"
  
  nCol_PartnerName=4
  sColWidth(4)="100"
  
  nCol_UserFieldText3=5
  sColWidth(5)="100"
End If

'SAY Входящие МИНЦ
If InStr(UCase(CurrentClassDoc), UCase(SIT_VHODYASCHIE))>0 and InStr(UCase(Session("Department")), UCase(SIT_MINC)) = 1 Then 
  ClearFieldsnCol
  NColumns=8
  
  nCol_DocID=1
  nCol_DocIDParent=1
  sColWidth(1)="80"
  
  nCol_DocIDIncoming=2
  sColWidth(2)="40"

  nCol_PartnerName=3
  sColWidth(3)="250"
  
  nCol_Content=4
  nCol_UserFieldText8=4
  S_NameUserFieldText8 = "Примечание"
  sColWidth(4)="350"

  nCol_UserFieldText3=5
  sColWidth(5)="100"

  nCol_DateActivation=6
  sColWidth(6)="40"
  
  
  nCol_UserFieldDate1=7
  sColWidth(8)="40"

  S_NameUserFieldDate2 = "Срок исполнения"
  nCol_UserFieldDate2=8
  sColWidth(8)="40"

End If

'входящие РТИ
      If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_VHODYASCHIE)) > 0 and InStr(UCase(Session("Department")), UCase(SIT_RTI)) = 1 Then
        DOCS_Name = "Заголовок к тексту (краткое содержание)"
        S_NameUserFieldText3 = "Подписант (внешний)"
    end if
'Исходящие РТИ    
    If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) > 0 and InStr(UCase(Session("Department")), UCase(SIT_RTI)) = 1 Then
        DOCS_Name = "Заголовок к тексту (краткое содержание)"
    end if
 'Распорядительные РТИ   
    If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_RASP_DOCS)) > 0 and InStr(UCase(Session("Department")), UCase(SIT_RTI)) = 1 Then
        DOCS_Name = "Заголовок к тексту (краткое содержание)"
        DOCS_DocID = "Номер распорядительного документа"
        DOCS_Author = "Инициатор"
    end if
 'Служебные записки РТИ    (общая форма)
    If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_SLUZH_ZAPISKA)) > 0 and InStr(UCase(Session("Department")), UCase(SIT_RTI)) = 1 Then
        DOCS_Name = "Заголовок к тексту (краткое содержание)"
    end if
'SAY 2008-10-30 Исходящие
If InStr(UCase(CurrentClassDoc), UCase(SIT_ISHODYASCHIE))>0 Then 
  ClearFieldsnCol
  NColumns=5
  
  nCol_DocID=1
  nCol_DocIDParent=1
  sColWidth(1)="80"
  
  nCol_Name=2
  sColWidth(2)="150"
  
  nCol_DateActivation=3
  sColWidth(3)="50"
  
  nCol_PartnerName=4
  sColWidth(4)="100"
  
  nCol_NameAproval=5
  sColWidth(5)="80"

End If
'Исходящие Минц   
If InStr(UCase(CurrentClassDoc), UCase(SIT_ISHODYASCHIE))>0 and InStr(UCase(Session("Department")), UCase(SIT_MINC)) = 1 Then 
  ClearFieldsnCol
  NColumns=5
  
  nCol_DocID=1
  nCol_DocIDParent=1
  sColWidth(1)="80"
  
  nCol_Name=2
  nCol_Content=2
  sColWidth(2)="150"
 
  nCol_DateActivation=3
  sColWidth(3)="50"
  
  nCol_PartnerName=4
  sColWidth(4)="100"
  
  nCol_NameAproval=5
  sColWidth(5)="80"

End If
' ******************************    MIKRON - start           ***************************
If InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_CONTRACT)) = 1 or _
   InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_S_CONTRACT)) = 1 or _
   InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_ADD_CONTRACT)) = 1 or _
   InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_EXPORT_CONTRACT)) = 1 or _
   InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_EXPADD_CONTRACT)) = 1 Then 
   ClearFieldsnCol
   NColumns=6
  
   nCol_DocID=1
'   nCol_UserFieldText1=1
'   sColWidth(1)="100"

   nCol_Name=2
   nCol_Description=2
   nCol_Content=2
   If InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_ADD_CONTRACT)) = 1 or _
      InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_EXPADD_CONTRACT)) = 1 Then 
      nCol_DocIDParent=1
   End If
'   sColWidth(2)="200"

   nCol_UserFieldDate4=3
   nCol_UserFieldDate5=3
   nCol_UserFieldDate6=3
'   nCol_DateActivation=3
'   nCol_DateCreation=3
'   nCol_DateCompletion=3
'   nCol_DateCompleted=3
'   nCol_DateExpiration=3
'   nCol_DateSigned=3
'   nCol_UserFieldDate5=3
'   sColWidth(3)="80"
  
   nCol_AmountDoc=4
   nCol_Currency=4
'   sColWidth(4)="80"
  
   nCol_Department=5
   nCol_PartnerName=5
'   nCol_UserFieldText1 = 5
'   sColWidth(5)="40"
    
   nCol_Author=6
   nCol_NameAproval=6
'   sColWidth(6)="40"
End If
  
If InStr(UCase(CurrentClassDoc), UCase(MIKRON_OLD_CONTRACT)) > 0 Then 
   ClearFieldsnCol
   NColumns=6
  
   nCol_DocID=1
   nCol_UserFieldText1=1
   sColWidth(1)="100"
  
   nCol_Name=2
   sColWidth(2)="200"
  
   nCol_UserFieldDate5 = 3
   sColWidth(3)="40"

   nCol_AmountDoc = 4
   nCol_Currency = 4
   sColWidth(4)="80"
  
   nCol_Department=5
   nCol_PartnerName=5
   sColWidth(5)="40"

   nCol_NameAproval=6
   sColWidth(6)="40"
End If

If InStr(UCase(CurrentClassDoc), UCase(MIKRON_NDA_CONTRACT)) > 0 Then
   ClearFieldsnCol
   NColumns=5
 
   nCol_DocID=1
   nCol_Name=2

   nCol_UserFieldDate4=3
   nCol_UserFieldDate5=3
   nCol_UserFieldDate6=3

   nCol_Department=4
   nCol_PartnerName=4

   nCol_NameAproval=5
   sColWidth(5)="40"
End If

If InStr( UCase(Session("CurrentClassDoc")), UCase(MIKRON_RL_MEMO) ) = 1 Then 
   ClearFieldsnCol
   NColumns=6
  
   nCol_DocID=1
'   nCol_DocIDadd = 1
'   nCol_DocIDParent = 1
   nCol_DocIDPrevious = 1
'   nCol_DocIDIncoming = 1
'   nCol_UserFieldText1=1
'   sColWidth(1)="100"

   nCol_Name=2
   nCol_Description=2
   nCol_Content=2
'   sColWidth(2)="200"

   nCol_UserFieldText8=3
   nCol_UserFieldText4=3
'   nCol_UserFieldDate4=3
'   nCol_UserFieldDate5=3
'   nCol_UserFieldDate6=3
'   nCol_DateActivation=3
'   nCol_DateCreation=3
'   nCol_DateCompletion=3
'   nCol_DateCompleted=3
'   nCol_DateExpiration=3
'   nCol_DateSigned=3
'   nCol_UserFieldDate5=3
'   sColWidth(3)="80"
  
   nCol_AmountDoc=4
   nCol_Currency=4
'   sColWidth(4)="80"
  
   nCol_Department=5
   nCol_PartnerName=5
'   nCol_UserFieldText1 = 5
'   sColWidth(5)="40"

   nCol_NameAproval=6
'   sColWidth(6)="40"
End If

If InStr( UCase(Session("CurrentClassDoc")), UCase(MIKRON_PROTOCOL) ) = 1 or _
   InStr( UCase(Session("CurrentClassDoc")), UCase(MIKRON_RL_PROTOCOL) ) = 1 Then 
   ClearFieldsnCol
   NColumns=5

   nCol_DocID=1
   nCol_DocIDParent = 1
'   nCol_DocIDPrevious = 1
'   sColWidth(1)="100"

   nCol_Name=2
   nCol_Description=2
   nCol_Content=2
'   sColWidth(2)="200"

'   nCol_UserFieldText8=3
'   nCol_UserFieldText4=3
'   nCol_UserFieldDate4=3
'   nCol_UserFieldDate5=3
'   nCol_UserFieldDate6=3
'   nCol_DateActivation=3
'   nCol_DateCreation=3
'   nCol_DateCompletion=3
'   nCol_DateCompleted=3
'   nCol_DateExpiration=3
'   nCol_DateSigned=3
'   nCol_UserFieldDate5=3
   nCol_AmountDoc=3
   nCol_Currency=3
'   sColWidth(3)="80"
  
   nCol_Department=4
   nCol_PartnerName=4
'   sColWidth(4)="80"
  
   If InStr( UCase(Session("CurrentClassDoc")), UCase(MIKRON_PROTOCOL) ) = 1 Then
'      nCol_UserFieldText1
'      nCol_UserFieldText2
      nCol_Resolution = 5
      nCol_UserFieldText3 = 5
      nCol_UserFieldMoney1 = 5
'      nCol_ListToView = 5
      sColWidth(5)="200"
   Else
      nCol_UserFieldText3=2
      nCol_UserFieldText4=2
      nCol_UserFieldText5=2
      nCol_ListToReconcile = 5
'      nCol_ListReconciled = 5
'      nCol_UserFieldText1 = 5
      sColWidth(5)="300"
   End If
'   sColWidth(5)="200"

End If
' ******************************    MIKRON - end           ***************************

'SAY 2008-08-22
nCol_DocIDadd=0
nCol_LocationPaper=0
'SAY 2008-10-08
nCol_Resolution=0

'SAY 2008-10-22 для списка поручений АФК
If UCase(Request("T_AFK")) ="Y" Then 
  ClearFieldsnCol
  
  nCol_DocID=1
  nCol_Name=2
  nCol_DateCompletion=3
  nCol_NameResponsible=4
  'NColumns=3

End If

'SAY 2008-12-02
'Запрос №11 - СТС - start
If InStr(UCase(CurrentClassDoc), UCase(SIT_DOGOVORI_OLD)) = 1 or InStr(UCase(CurrentClassDoc), UCase(SIT_DOGOVORI_NEW)) = 1 Then 
'If InStr(UCase(CurrentClassDoc), UCase(SIT_DOGOVORI))>0 Then 
'Запрос №11 - СТС - end
  sColWidth(1)="40"
  sColWidth(2)="50"
  sColWidth(3)="200"
  sColWidth(4)="100"
  sColWidth(5)="80"
  sColWidth(5)="80"
End If
'Запрос №11 - СТС - start
'Запретить создание документов в старой категории договоров
If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_DOGOVORI_OLD)) = 1 Then
	ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickCreateDoc"
	ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickCreateDocTemplate"
End If
'Запрос №11 - СТС - end
'Запрос №46 - СТС - start
'Запретить сотрудникам СТС создавать документы в старой категории на переработки
If InStr(UCase(Session("CurrentClassDoc")), UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) = 1 Then
	If InStr(UCase(Session("Department")), UCase(SIT_STS_ROOT_DEPARTMENT)) = 1 Then
		ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickCreateDoc"
		ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickCreateDocTemplate"
	End If
End If
'Служебные на переработки (факт) можно создавать только из плановых
If InStr(UCase(Session("CurrentClassDoc")), UCase(STS_SLUZH_ZAPISKA_OVERTIME_FACT)) = 1 Then
	ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickCreateDoc"
	ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickCreateDocTemplate"
End If
'Запрос №46 - СТС - end

'Заявки
If InStr(UCase(CurrentClassDoc), UCase(STS_PurchaseOrder))=1 or InStr(UCase(CurrentClassDoc), UCase(STS_PaymentOrder))>0 Then 
  ClearFieldsnCol
  NColumns=6
  
  nCol_DocID=1
  nCol_DocIDParent=1
  
  nCol_Name=2
  nCol_Description=2
  
  nCol_DateActivation=3
  nCol_DateCompletion=3

  nCol_AmountDoc=4

  nCol_Department=5
  nCol_PartnerName=5

  nCol_Author=6
End If

'20090622 - Заявка ТКП
'Коммерческие предложения
If InStr(UCase(CurrentClassDoc), UCase(SIT_COM_OFFERS))>0 Then 
  ClearFieldsnCol
  NColumns=6
  
  nCol_DocID=1
  nCol_DocIDParent=1
  sColWidth(1)="80"
  
  nCol_Name=2
  sColWidth(2)="100"
  
  nCol_UserFieldText1=3
  S_NameUserFieldText1 = SIT_ComOfferNumber
  sColWidth(3)="50"
  
  nCol_DateActivation=4
  sColWidth(4)="50"
  
  nCol_PartnerName=5
  sColWidth(5)="100"
  
  nCol_NameAproval=6
  sColWidth(6)="80"

End If

'Центральные кнопки: Срочные, Регистрация, Все документы, В работу, Мои документы, Согласование, Утверждение, Ознакомление
If UCase(Request("BadDocs")) = "Y" or UCase(Request("registration")) = "Y" or UCase(Request("AllDocs")) = "Y" or (UCase(Request("VisaDocs")) = "Y" and Request("UserIDToSee") <> "") or UCase(Request("CreatedDocs")) = "Y" or UCase(Request("VisaDocs")) = "Y" or UCase(Request("ApprDocs")) = "Y" or UCase(Request("ViewedDocs")) = "Y" Then
  ClearFieldsnCol
  NColumns=6
  
  nCol_DocID=1
  sColWidth(1)="80"
  
  nCol_Name=2
  nCol_ClassDoc=2
  sColWidth(2)="150"
  
  nCol_Description=3
  nCol_UserFieldText1=3
  S_NameUserFieldText1 = SIT_Subject
  sColWidth(3)="170"

  nCol_DateActivation=4
  nCol_DateCompletion=4
  sColWidth(4)="50"
  
  DOCS_Author = SIT_Initiator
  nCol_Author=5
  sColWidth(5)="80"
  
  nCol_NameAproval=6
  nCol_NameResponsible=6
  sColWidth(6)="80"
End If

'Центральные кнопки: Мои поручения, Поручения мне, Соисполнитель
If UCase(Request("RespDocs")) = "Y" or UCase(Request("CompleteDocs")) = "Y" or Request("UserBtn")="CoResponsible" Then
  ClearFieldsnCol
  NColumns=6
  
  nCol_DocID=1
  sColWidth(1)="80"
  
  DOCS_Name = SIT_TaskName
  nCol_Name=2
  sColWidth(2)="150"
  
  S_NameUserFieldText1 = SIT_ReportTypeOfTask
  nCol_UserFieldText1=3
  sColWidth(3)="170"

  nCol_DateActivation=4
  nCol_DateCompletion=4
  sColWidth(4)="50"
  
  DOCS_Author = SIT_Initiator
  nCol_Author=5
  sColWidth(5)="80"
  
  nCol_NameResponsible=6
  sColWidth(6)="80"
End If



Sub ClearFieldsnCol 'Clear all field numbers
  nCol_DocID=0
  nCol_DocIDadd=0
  nCol_DocIDParent=0
  nCol_DocIDPrevious=0
  nCol_DocIDIncoming=0
  nCol_Author=0
  nCol_Correspondent=0
  nCol_Resolution=0
  nCol_History=0
  nCol_Result=0
  nCol_PercentCompletion=0
  nCol_Department=0
  nCol_Name=0
  nCol_Description=0
  nCol_LocationURL=0
  nCol_LocationPaper=0
  nCol_Currency=0
  nCol_CurrencyRate=0
  nCol_Rank=0
  nCol_LocationPath=0
  nCol_ExtInt=0
  nCol_PartnerName=0
  nCol_StatusDevelopment=0
  nCol_StatusArchiv=0
  nCol_StatusCompletion=0
  nCol_StatusDelivery=0
  nCol_StatusPayment=0
  nCol_TypeDoc=0
  nCol_ClassDoc=0
  nCol_ActDoc=0
  nCol_InventoryUnit=0
  nCol_PaymentMethod=0
  nCol_AmountDoc=0
  nCol_QuantityDoc=0
  nCol_DateActivation=0
  nCol_SecurityLevel=0
  nCol_DateCreation=0
  nCol_DateActive=0
  nCol_DateCompletion=0
  nCol_DateCompleted=0
  nCol_DateExpiration=0
  nCol_DateSigned=0
  nCol_NameCreation=0
  nCol_NameAproval=0
  nCol_NameApproved=0
  nCol_DateApproved=0
  nCol_NameControl=0
  nCol_ListToEdit=0
  nCol_ListToView=0
  nCol_ListToReconcile=0
  nCol_ListReconciled=0
  nCol_NameResponsible=0
  nCol_NameLastModification=0
  nCol_DateLastModification=0
  nCol_UserFieldText1=0
  nCol_UserFieldText2=0
  nCol_UserFieldText3=0
  nCol_UserFieldText4=0
  nCol_UserFieldText5=0
  nCol_UserFieldText6=0
  nCol_UserFieldText7=0
  nCol_UserFieldText8=0
  nCol_UserFieldMoney1=0
  nCol_UserFieldMoney2=0
  nCol_UserFieldDate1=0
  nCol_UserFieldDate2=0
  nCol_UserFieldDate3=0
  nCol_UserFieldDate4=0
  nCol_UserFieldDate5=0
  nCol_UserFieldDate6=0
  nCol_UserFieldDate7=0
  nCol_UserFieldDate8=0
  nCol_IsActive=0
  nCol_DateActive=0
  nCol_BusinessProcessStep=0
  nCol_ExtPassword=0
  nCol_Content=0
  nCol_GUID=0
  nCol_FileNameNameLastModification=0
End Sub

'{ph - 20120628
Function UserPrepareListDocField(ByRef parField)
	Dim sColor

	UserPrepareListDocField = ""
	Select Case parField.Name
		Case "Name"
			UserPrepareListDocField = ShowContextHTMLEncode(HTMLEncode(parField.Value))
			If InStr(UCase(dsDoc("ClassDoc")), UCase(SIT_ZADACHI)) = 1 Then
				sColor = ""
				If MyCStr(dsDoc("StatusCompletion")) = VAR_StatusCompletion Then
					sColor = "green"
				ElseIf dsDoc("DateCompletion") < Date() Then
					sColor = "red"
				End If
				If sColor <> "" Then
					UserPrepareListDocField = "<font color = """ & sColor & """>" & UserPrepareListDocField & "</font>"
				End If
			End If
		Case "UserFieldText1", "UserFieldText2", "UserFieldText3", "UserFieldText4", "UserFieldText5", "UserFieldText6", "UserFieldText7", "UserFieldText8"
			UserPrepareListDocField = ShowContextHTMLEncode(HTMLEncode(parField.Value))

        Case "UserFieldText2"
            UserPrepareListDocField = ShowContextHTMLEncode(HTMLEncode(parField.Value))
            If InStr(UCase(dsDoc("ClassDoc")), UCase(SIT_ZADACHI)) = 1 Then
                sLink = "\\gl-fs-01\RTI\СЭД 2016"
                UserPrepareListDocField = "<a href=""" & sLink & """>" & UserPrepareListDocField & "</a>"
                
            End If

	End Select
End Function
'ph - 20120628}
%>