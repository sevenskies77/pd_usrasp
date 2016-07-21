<%
'Place here ASP code to set the document record preinstalled values
'Use expression Request("DocID") to get the value of the DocID field
'Use expression Request("ClassDoc") to get the value of the Doument category field
'Variables for preinstalled values
'S_DocID_Set
'S_Department_Set
'S_Correspondent_Set
'S_NameAproval_Set
'S_NameControl_Set
'S_ListToView_Set
'S_ListToEdit_Set
'S_ListToReconcile_Set
'S_NameResponsible_Set
'S_UserFieldText1_Set
'S_UserFieldText2_Set
'S_UserFieldText3_Set
'S_UserFieldText4_Set
'S_UserFieldText5_Set
'S_UserFieldText6_Set
'Variables for current values
'S_AmountDoc
'S_ClassDoc - current document category
'CurrentClassDoc - current document category
'CurrentDocFieldOrder - current document category field order

'If InStr(UCase(S_ClassDoc), UCase("Пропуска"))>0 Or InStr(UCase(Request("ClassDoc")), UCase("Пропуска"))>0 Then 
'	If Request("create")="y" Then
'		If Trim(S_DocID)<>"" Then
'			S_DocID_Set=GenNewDocIDIncrement(S_DocID)
'		Else
'			S_DocID_Set="ПР"+CStr(Year(Date))+CStr(Month(Date))+CStr(Day(Date))+"/0001"
'		End If
'		If Trim(Session("Role"))="" Then
'			S_UserFieldText3=SurnameGN(Session("Name"))+", "+Session("Position")
'			S_UserFieldText4=Session("Phone")
'		Else
'			oPayDox.GetUserDetails Trim(Session("Role")), sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
'			S_UserFieldText3=SurnameGN(sName)+", "+sPosition
'			S_UserFieldText4=sPhone
'		End If		
'	Else
'		S_DocID_Set=S_DocID
'	End If
'End If
' sts alex 20:19 19.05.2009 - PartnerName field is requred now for purchaseOrder

If IsHelpDeskDoc() Then
	VAR_ChangeDocGetNewButton=""
	VAR_ChangeDocGenerateButton=""
	VAR_ChangeDocGetNewFromRegLogsButton=""
	CurrentDocFieldOrder="DocID,DocIDadd,DocIDParent,DocIDPrevious,DocIDIncoming,Author,Correspondent,Resolution,History,Result,PercentCompletion,Department,Name,Description,LocationURL,LocationPaper,Currency,CurrencyRate,Rank,FileNamePrefix,FileName,FileNameNameLastModification,FileNameDateLastAccessed,FileNameDateLastModification,LocationPath,ExtInt,PartnerName,StatusDevelopment,StatusArchiv,StatusCompletion,StatusDelivery,StatusPayment,TypeDoc,ClassDoc,ActDoc,InventoryUnit,PaymentMethod,AmountDoc,QuantityDoc,DateActivation,SecurityLevel,DateCreation,DateCompletion,DateCompleted,DateExpiration,DateSigned,NameCreation,NameAproval,NameApproved,DateApproved,NameControl,ListToEdit,ListToView,ListToReconcile,ListReconciled,NameResponsible,NameLastModification,DateLastModification,UserFieldText1,UserFieldText2,UserFieldText3,UserFieldText4,UserFieldText5,UserFieldText8,UserFieldMoney1,UserFieldMoney2,UserFieldDate1,UserFieldDate2,IsActive,DateActive,BusinessProcessStep,ExtPassword,Content,GUID,UserFieldText6,UserFieldText7"
	S_ActDoc = DOCS_HelpDesk
	DOCS_PartnerName=DOCS_Company	
	If Request("UpdateDoc")<>"YES" Then
	If Request("create") = "y" Then
		S_Department_Set=Session("Department")
		S_PartnerName_Set=GetCompanyName(Session("Department"))
		S_IsActive_Set = VAR_ActiveTask		
		If IsHelpDeskAdmin() Or IsSupervisor() Then
			S_DateCompletion=MyDate(Date+3)
		End If
		If Not IsHelpDeskAdmin() And Not IsSupervisor() Then
			'bRank=""
			'If InStr("ДИРЕКТОР # МЕНЕДЖЕР # НАЧАЛЬНИК", UCase(Session("Position")))>0 Then
			'	S_Rank_Set="Срочный"
			'	S_DateCompletion_Set=MyDate(Date+4)
			'Else
			'	S_Rank_Set="Обычный"
			'	'S_Rank_Set="-"
			'	S_DateCompletion_Set=MyDate(Date+7)
			'End If
'Out "S_Rank_Set:"+S_Rank_Set			
'Out "S_DateCompletion_Set:"+S_DateCompletion_Set			
			bNameResponsible=""
			bListToReconcile=""
			bResolution=""
			bCorrespondent=""
			bDateCompletion=""
			bResolution=""
			S_NameUserFieldText6=""
			S_NameUserFieldText7=""
		End If
		If IsHelpDeskAdminOrConsultant() Then
			bResolution="Y"
		End If
		S_Rank_Set=""
		S_Rank=""
		S_Correspondent_Set=""
		S_Correspondent=""
		S_Resolution_Set=""
		S_Resolution=""
	Else
		If Not IsAdmin() And Not IsSupervisor() Then
			S_Department_Set=S_Department
			S_PartnerName_Set=S_PartnerName
			S_Rank_Set=S_Rank
			S_DateCompletion_Set=S_DateCompletion
			If S_DateCompletion="" Then
				bDateCompletion=""
			End If
			S_DocID_Set=S_DocID
			If S_NameResponsible="" Then
				bNameResponsible=""
			Else
				S_NameResponsible_Set=S_NameResponsible
			End If
			If S_ListToReconcile="" Then
				bListToReconcile=""
			Else
				S_ListToReconcile_Set=S_ListToReconcile
			End If
			If S_Correspondent="" Then
				bCorrespondent=""
			Else
				S_Correspondent_Set=S_Correspondent
			End If
			If S_Correspondent="" Then
				bResolution=""
			Else
				S_Resolution_Set=S_Resolution
			End If
		End If
	End If 'Request("create") = "y" Then
	End If 'Request("UpdateDoc")<>"YES" Then
	'If InStr(S_Correspondent, "<"+Session("UserID")+">")>0 Then
	If IsHelpDeskAdminOrConsultant() Then
		bResolution="Y"
		S_Resolution_Set=""
	End If
If Request("create") = "y" Then
	If Request("UpdateDoc")<>"YES" Then
		S_SecurityLevel=4
		If Not IsHelpDeskAdminOrConsultant() Then
                        'SAY 2008-10-17 закомментирована строка, доступ на изменение при создании карточки документа
			'S_SecurityLevel_Set=4
		End If
		If Request("task") = "y" Then
			S_NameResponsible=""
			'S_NameControl=GetFullName(Session("Name"), Session("UserID"))
			'S_NameControl=S_NameAproval
		End If
	End If
End If
End If 'If IsHelpDeskDoc() Then

'If Request("create") = "y" Then
'	S_SecurityLevel=4
'	If S_ClassDoc="Договора" Then
'		If S_ActDoc<>"" Then
'			S_ListToReconcile=GetUserDirValue("Группы пользователей", S_ActDoc, 1, 2)
'		End If			
'	End If			
'End If			

'VAR_ChangeDocGenerateButton =""
'VAR_ChangeDocGetNewButton =""
'VAR_ChangeDocGetNewFromRegLogsButton ="Y"

'VAR_ChangeDocGetNewButtonDocIDAdd =""
'VAR_ChangeDocGetNewFromRegLogsButtonDocIDAdd ="Y"

'S_ConnectedDocCommonFields="AmountDoc, Currency"
'CurrentDocRequiredFields ="Author, Correspondent, Department, Description, UserFieldText1, UserFieldText8"

'If Request("paste")<>"" Then
'	If Request("ClassDoc")="Договора" Then
'		S_PartnerName="Вставьте наименование контрагента"
'	End If
'End If

'If Request("ClassDocDependant")<>"" Then
'	S_ListToReconcile=S_Correspondent
'	S_Correspondent="Вставьте список"
'End If
'If S_ClassDoc="Служебные записки/СЗ на командировку" And UCase(Request("UpdateDoc")) <> "YES" Then
'	S_NameAproval_Set="""#Директор департамента инициатора д-та"""
'	S_ListToReconcile_Set=oPayDox.GetExtTableValue("Agree", "Category", "Служебные записки/СЗ на командировку", "List")
'End If


' ******** Настройки для Sitronics/STS ***************************************************
if not (InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_CONTRACTS_MC)) > 0) Then 
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow + ", AddFieldText2"
end if

' ********************************* ВХОДЯЩИЕ ДОКУМЕНТЫ
' *** 
If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_VHODYASCHIE)) = 1 Then 
  'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
  S_SecurityLevel = 4
  If not IsAdmin() Then
    S_SecurityLevel_Set = S_SecurityLevel
  End If

  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
    S_AdditionalUsers = ""
  End If

'rti_vhodyaschie  
  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" and InStr(UCase(Session("Department")), UCase(SIT_RTI)) = 1 Then
    S_AdditionalUsers = """Боев С. Ф."" <boev_oaorti>;"
    If InStr(UCase(Session("Department")), UCase("Управление ИТ")) > 0 Then
      S_AdditionalUsers = S_AdditionalUsers + VbCrLf + """Иванов В. В."" <vivanon_oaorti>;"
    End If
    DocIDParent_RTI = SIT_GetDocField(S_DocIDParent, "Correspondent", Conn)
    DocIDPar_RTI = SIT_GetDocField(S_DocIDParent, "DocId", Conn)
    if len(DocIDParent_RTI) >0 and InStr(UCase(DocIDPar_RTI),"T_") > 0 Then
        S_AdditionalUsers = S_AdditionalUsers + VbCrLf + DocIDParent_RTI
    End If
  End If
'rti_vhodyaschie  

'minc_vhodyaschie  
  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" and InStr(UCase(Session("Department")), UCase(SIT_MINC)) = 1 Then
    S_AdditionalUsers = """Селютина О. А."" <usr_oselyutina>;" + VbCrLf + """Бабушкин А. Н."" <usr_ababushkin>;" + VbCrLf + """Теппер А. Б."" <usr_atepper>;" + VbCrLf
    DocIDParent_RTI = SIT_GetDocField(S_DocIDParent, "Correspondent", Conn)
    DocIDPar_RTI = SIT_GetDocField(S_DocIDParent, "DocId", Conn)
    if len(DocIDParent_RTI) >0 and InStr(UCase(DocIDPar_RTI),"К_") > 0 Then
        S_AdditionalUsers = S_AdditionalUsers + VbCrLf + DocIDParent_RTI
    End If
  End If
'minc_vhodyaschie  

'vtss_vhodyaschie  
  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" and InStr(UCase(Session("Department")), UCase(SIT_VTSS)) = 1 Then   
    S_AdditionalUsers = """Подольский А. Е."" <vtss_a.podolskii>;" + VbCrLf
    DocIDParent_RTI = SIT_GetDocField(S_DocIDParent, "Correspondent", Conn)
    DocIDPar_RTI = SIT_GetDocField(S_DocIDParent, "DocId", Conn)
    if len(DocIDParent_RTI) >0 and InStr(UCase(DocIDPar_RTI),"T_") > 0 Then
        S_AdditionalUsers = S_AdditionalUsers + VbCrLf + DocIDParent_RTI
    End If
  End If
'vtss_vhodyaschie  


  If S_AdditionalUsers <> "" Then
     S_AdditionalUsers_Set = S_AdditionalUsers
  Else
     S_AdditionalUsers_Set = " "
  End If

  S_DocID_Set = S_DocID
  S_DocIDAdd_Set = S_DocIDAdd
  S_Department_Set = S_Department
  S_Author_Set = S_Author

  ' Создание карточки документа
  If UCase(Request("create")) = "Y" Then
     If Request("UpdateDoc") <> "YES" Then
        'SAY 2009-02-20
        VAR_ChangeDocGenerateButton=""
        S_DocID = " "
        If S_DocID_Set<>"" Then
           S_DocID_Set = S_DocID
        End If

        'SAY 2008-11-06
        S_Department_Set = Session("Department")

        'SAY 2008-09-22 для подчиненных документов оставляем весь список согласования родителя
        If Trim(S_DocIDParent)<>"" then
           S_ListToReconcile = Replace(S_ListToReconcile,"##;","")
        End If

        S_DocID = ""
        S_DocID_Set = " "
        S_DateActivation = MyDate(Date)
        If not IsAdmin() Then
           S_DateActivation_Set = S_DateActivation
        End If
        S_Author_Set = InsertionName(Session("Name"), Session("UserID"))
        If InStr(UCase(Session("Department")), UCase(SIT_MINC)) = 1 Then
          S_Name = ""
        End If
        S_Department_Set = Session("Department")
     End If

'     If InStr(UCase(Session("Department")), UCase(SIT_SITRU)) = 1 then ' DmGorsky
'	     ' По умолчанию "Вид отправителя" = "Компании бизнес-направления"
'        S_UserFieldText7_Set = "3 - Компании бизнес-направления"
'     End If
  End If 'UCase(Request("create")) = "Y"

  S_ListToReconcile = DeleteConstPrefixFromList(S_ListToReconcile)

  CurrentDocFieldOrder = "Name, UserFieldText1, UserFieldText7, DateActivation, DocID, Author, UserFieldText2, ListToView, DocIDIncoming, UserFieldDate1, PartnerName, UserFieldText3, UserFieldText4, UserFieldText5, UserFieldText6, Content, SecurityLevel, Department "
  CurrentDocRequiredFields = "Name, UserFieldText1, ListToView, DocIDIncoming, UserFieldDate1, PartnerName, UserFieldText3, UserFieldText5, Content," ' SecurityLevel"
  
  'ph - 20080914 - У админа поле DateActivation редактируемое, т.ч. делаем обязательным
  If IsAdmin() Then
     CurrentDocRequiredFields = CurrentDocRequiredFields + ", DateActivation"
  End If

  'SAY 2008-10-03 добавляем поле "Вид отправителя" и для СТС
  CurrentDocRequiredFields = CurrentDocRequiredFields+",UserFieldText7" 

  'отмена справочников для пользовательских полей
  VAR_DirPictNotToShow = "UserFieldText5, UserFieldText6"

'Запрос №31 - СТС - start
  If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 then
     If UCase(Request("create")) <> "Y" Then
        'Запрещаем редактировать БЕ, от нее зависит номер
        If Request("UpdateDoc")<>"YES" Then
           If S_AddField2 = "" Then
              S_AddField2 = MyCStr(dsDoc("BusinessUnit"))
           End If
        Else
           If S_AddField2 = "" Then
              S_AddField2 = Request("BusinessUnit")
           End If
        End If
        If S_AddField2 = "" Then
           S_AddField2 = " "
        End If
        S_AddField_Set2 = S_AddField2
     End If
  End If
'Запрос №31 - СТС - end

  VAR_DirPictNotToShow = VAR_DirPictNotToShow+", DocName"
  'rmanyushin 60298 02.11.2009 Start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", ContractType"
  'rmanyushin 60298 02.11.2009 End
'ph - 20120216 - start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", AddFieldText1"
'ph - 20120216 - end

  'Выпадающие списки
  Select Case UCase(Request("l"))
    Case "RU" 'RU
      S_UserFieldText1_Select = GetUserDirValues("{459D6AE1-7E6F-467E-B162-A36A4054AA5A}")
      S_UserFieldText2_Select = GetUserDirValues("{6D42F0FB-1389-4CA0-8A00-9E0CD3F09CF7}")
      If InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) = 1 then
         S_UserFieldText7_Select = GetUserDirValues("{D68C37BE-9EBA-4CA8-B0D2-C6369123E7C7}")
'Запрос №1 - СИБ - start
      ElseIf InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 then
         S_UserFieldText7_Select = GetUserDirValues("{D0E4DEEB-A1A8-48B0-B527-9573856769DB}")
'Запрос №1 - СИБ - start
      Else
         S_UserFieldText7_Select = GetUserDirValues("{da5960be-a65d-4d21-bf89-73233ffeaee8}")
      End If
    Case "" 'EN
      S_UserFieldText1_Select = GetUserDirValues("{8F0D8C83-05F9-4148-96E9-3D015143063F}")
      S_UserFieldText2_Select = GetUserDirValues("{6D57662F-7DD0-41E1-806B-3562412FDFAF}")
      If InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) = 1 then
         S_UserFieldText7_Select = GetUserDirValues("{FC8A0260-6B28-4F9F-BF2D-6F95DDE21E1C}")
'Запрос №1 - СИБ - start
      ElseIf InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 then
         S_UserFieldText7_Select = GetUserDirValues("{D0E4DEEB-A1A8-48B0-B527-9573856769DB}")
'Запрос №1 - СИБ - start
      Else
         S_UserFieldText7_Select = GetUserDirValues("{6A200BD7-1A53-40FC-9DBB-44499F65B74C}") '------------- новый справочник
      End If
    Case "3" 'CZ
      S_UserFieldText1_Select = GetUserDirValues("{3885C48E-1CDB-4D59-A89C-DDDD6FB19FF3}")
      S_UserFieldText2_Select = GetUserDirValues("{0D620DAB-1B89-4E7B-BB6A-29EB77F9AEE9}")
      If InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) = 1 then
         S_UserFieldText7_Select = GetUserDirValues("{521C56BD-EC92-4AF5-BE8C-229391C37673}")
'Запрос №1 - СИБ - start
      ElseIf InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 then
         S_UserFieldText7_Select = GetUserDirValues("{D0E4DEEB-A1A8-48B0-B527-9573856769DB}")
'Запрос №1 - СИБ - start
      Else
         S_UserFieldText7_Select = GetUserDirValues("{2F4D0C04-FD15-4321-A5E3-5AA2FCB0D70E}") '------------- новый справочник
      End If
  End Select

' Изменения для СИТРУ ' DmGorsky_3
If InStr(UCase(Session("Department")), UCase(SIT_SITRU)) = 1 and UCase(Request("create")) = "Y" Then ' DmGorsky_3
   S_UserFieldText7_Select = GetUserDirValues("{D68C37BE-9EBA-4CA8-B0D2-C6369123E7C7}") ' DmGorsky_3
   'S_UserFieldText7_Select2 = GetUserDirValues("{D0E4DEEB-A1A8-48B0-B527-9573856769DB}")
   CurrentDocFieldOrder = CurrentDocFieldOrder + ", Resolution, UserFieldText8" ' DmGorsky_3

   'DOCS_Resolution = "Резолюция"
   'S_Resolution_Set = "Резолюция:"
   'S_Resolution = "Резолюция:"

ElseIf InStr(UCase(Session("Department")), UCase(SIT_SITRU)) <> 1 Then ' DmGorsky_3
   VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow + ", Resolution" ' DmGorsky_3
End If ' DmGorsky_3

  'Для пользователей СТС везде добавляем поле бизнес единица
  If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
     If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 Then
        If UCase(Request("create")) = "Y" Then
           sCreator = Session("UserID")
        Else
           sCreator = GetUserID(S_Author)
        End If
        BusinessUnitsList = GetUsersBusinessUnits(sCreator)
        If InStr(BusinessUnitsList, VbCrLf) > 0 Then
           S_AddField2 = " "
           S_AddField_Select2 = BusinessUnitsList
           'Поле обязательно для заполнения
           CurrentDocRequiredFields = CurrentDocRequiredFields+", BusinessUnit"
        Else
           S_AddField_Set2 = BusinessUnitsList
        End If
     End If
  Else
     'Поле показывается только для пользователей СТС и администратора (у админа не из СТС оно необязательное, редактируемое)
     If not IsAdmin() Then
        VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", BusinessUnit"
     End If
  End If
 
'rti_vhodyaschie  
  If InStr(UCase(Session("Department")), UCase(SIT_RTI)) = 1 Then
     'VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow + ", UserFieldText1, Content, AddFieldText2, UserFieldText7, DocIDParent"
     VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow + ", UserFieldText1, Content, AddFieldText2, UserFieldText7"
     CurrentDocFieldOrder = "Name, DateActivation, DocID, DocIDParent, Author, UserFieldText8, UserFieldText2, ListToView, DocIDIncoming, UserFieldDate1, PartnerName, UserFieldText3, UserFieldText4, UserFieldText5, UserFieldText6, SecurityLevel, Department "
     CurrentDocRequiredFields = "Name, ListToView, DocIDIncoming, UserFieldDate1, PartnerName, UserFieldText3, UserFieldText5, UserFieldText8" ' SecurityLevel"
     S_UserFieldText8_Select = GetUserDirValues("{9EB619A3-61E8-4C1A-9573-27C87DABEF76}")
  End If
'rti_vhodyaschie  

'vtss_vhodyaschie  
  If InStr(UCase(Session("Department")), UCase(SIT_VTSS)) = 1 Then
     'VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow + ", UserFieldText1, Content, AddFieldText2, UserFieldText7, DocIDParent"
     VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow + ", UserFieldText1, Content, AddFieldText2, UserFieldText7"
     CurrentDocFieldOrder = "Name, DateActivation, DocID, DocIDParent, Author, UserFieldText2, ListToView, DocIDIncoming, UserFieldDate1, PartnerName, UserFieldText3, UserFieldText4, UserFieldText5, UserFieldText6, SecurityLevel, Department "
     CurrentDocRequiredFields = "Name, ListToView, DocIDIncoming, UserFieldDate1, PartnerName, UserFieldText3, UserFieldText5, UserFieldText8" ' SecurityLevel"
  End If
'vtss_vhodyaschie  

'oaorti_vhodyaschie
    If InStr(UCase(Session("Department")), UCase(SIT_MINC)) = 1 Then
        VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow + ", AddFieldText2"
        VAR_DirPictNotToShow = VAR_DirPictNotToShow + ",UserFieldText7, UserFieldText8"
        S_UserFieldText1_Select = GetUserDirValues("{450BA4D3-65B9-4C2D-9224-88CC4C9B1CD1}")
        S_Name_Select = GetUserDirValues("{7CD9F979-A1CB-48C5-8A6C-AC4C1E950693}")        
        If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
          S_UserFieldText1 = "Письмо"
          S_UserFieldText2 = "Почта"       
        end if
        CurrentDocFieldOrder = "Name, UserFieldText1, DateActivation, Author, DocID, DocIDParent, UserFieldText2, ListToView, DocIDIncoming, UserFieldDate1, PartnerName, UserFieldText3, UserFieldText4, UserFieldText5, UserFieldText6, Content, SecurityLevel, Department, UserFieldText7, UserFieldDate2, UserFieldText8"
        CurrentDocRequiredFields = "Name, UserFieldText1, PartnerName, ListToView, DocIDIncoming, UserFieldText5, Content, UserFieldDate1"       
        S_UserFieldText7_Select = ""
        S_UserFieldText8_Select = ""
    End If
'oaorti_vhodyaschie


' ********************************* ВХОДЯЩИЕ ДОКУМЕНТЫ ДЛЯ БУХГАЛТЕРИИ
' *** 
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_VHODYASCHIE_ACC)) = 1 Then
AddLogD "amw UserChangeDocSetValue "
  'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
  S_SecurityLevel = 4
  If not IsAdmin() Then
    S_SecurityLevel_Set = S_SecurityLevel
  End If

  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
    S_AdditionalUsers = ""
  End If

  If S_AdditionalUsers <> "" Then
    S_AdditionalUsers_Set = S_AdditionalUsers
  Else
    S_AdditionalUsers_Set = " "
  End If

  S_DocID_Set = S_DocID
  S_DocIDAdd_Set = S_DocIDAdd
  S_Department_Set = S_Department
  S_Author_Set = S_Author

  ' Создание карточки документа
  If UCase(Request("create")) = "Y" Then
    If Request("UpdateDoc") <> "YES" Then
      'SAY 2009-02-20
      VAR_ChangeDocGenerateButton=""
      S_DocID = " "
      If S_DocID_Set<>"" Then
        S_DocID_Set = S_DocID
      End If

      'SAY 2008-11-06
      S_Department_Set = Session("Department")

      'SAY 2008-09-22 для подчиненных документов оставляем весь список согласования родителя
      If Trim(S_DocIDParent)<>"" then
        S_ListToReconcile = Replace(S_ListToReconcile,"##;","")
      End If

      S_DocID = ""
      S_DocID_Set = " "
      S_DateActivation = MyDate(Date)
      If not IsAdmin() Then
        S_DateActivation_Set = S_DateActivation
      End If
      S_Department_Set = Session("Department")
    End If
  End If 'UCase(Request("create")) = "Y"

  S_ListToReconcile = DeleteConstPrefixFromList(S_ListToReconcile)

  CurrentDocFieldOrder = "Name, UserFieldText1, UserFieldText7, DateActivation, DocID, UserFieldText2, ListToView, DocIDIncoming, UserFieldDate1, PartnerName, UserFieldText3, UserFieldText4, UserFieldText5, UserFieldText6, Content, SecurityLevel, Department "
  CurrentDocRequiredFields = ""
  'phil - 20080914 - У админа поле DateActivation редактируемое, т.ч. делаем обязательным
  If IsAdmin Then
    CurrentDocRequiredFields = CurrentDocRequiredFields + ", DateActivation"
  End If

  'отмена справочников для пользовательских полей
  VAR_DirPictNotToShow = "UserFieldText5, UserFieldText6"

  VAR_DirPictNotToShow = VAR_DirPictNotToShow+", DocName"
  'rmanyushin 60298 02.11.2009 Start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", ContractType"
  'rmanyushin 60298 02.11.2009 End
'ph - 20120216 - start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", AddFieldText1"
'ph - 20120216 - end

  'Для пользователей СТС везде добавляем поле бизнес единица
  If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
    If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 Then
      If UCase(Request("create")) = "Y" Then
        sCreator = Session("UserID")
      Else
        sCreator = GetUserID(S_Author)
      End If
      BusinessUnitsList = GetUsersBusinessUnits(sCreator)
      If InStr(BusinessUnitsList, VbCrLf) > 0 Then
        S_AddField2 = " "
        S_AddField_Select2 = BusinessUnitsList
        'Поле обязательно для заполнения
        CurrentDocRequiredFields = CurrentDocRequiredFields+", BusinessUnit"
      Else
        S_AddField_Set2 = BusinessUnitsList
      End If
    End If
  Else
    'Поле показывается только для пользователей СТС и администратора (у админа не из СТС оно необязательное, редактируемое)
    If not IsAdmin() Then
      VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", BusinessUnit"
    End If
  End If

' ********************************************************************************* 
' ***                       ИСХОДЯЩИЕ ДОКУМЕНТЫ                                 ***
' *********************************************************************************
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_ISHODYASCHIE)) = 1 Then
  'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
  S_SecurityLevel = 4
  If not IsAdmin() Then
    S_SecurityLevel_Set = S_SecurityLevel
  End If

  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
    S_AdditionalUsers = ""
  End If

'rti_ishodyaschie
  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" and InStr(UCase(Session("Department")), UCase(SIT_RTI)) = 1 Then
     S_Name = "О "
     S_AdditionalUsers = """Боев С. Ф."" <boev_oaorti>;"
    If InStr(UCase(Session("Department")), UCase("Управление ИТ")) > 0 Then
      S_AdditionalUsers = S_AdditionalUsers + VbCrLf + """Иванов В. В."" <vivanon_oaorti>;"
    End If
     DocIDParent_RTI = SIT_GetDocField(S_DocIDParent, "Correspondent", Conn)
     DocIDPar_RTI = SIT_GetDocField(S_DocIDParent, "DocId", Conn)
     if len(DocIDParent_RTI) >0 and InStr(UCase(DocIDPar_RTI),"T_") > 0 Then
        S_AdditionalUsers = S_AdditionalUsers + VbCrLf + DocIDParent_RTI
     End If
  End If
'rti_ishodyaschie

'minc
  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" and InStr(UCase(Session("Department")), UCase(SIT_MINC)) = 1 Then
    S_AdditionalUsers = """Селютина О. А."" <usr_oselyutina>;" + VbCrLf + """Бабушкин А. Н."" <usr_ababushkin>;" + VbCrLf + """Теппер А. Б."" <usr_atepper>;" + VbCrLf
    DocIDParent_RTI = SIT_GetDocField(S_DocIDParent, "Correspondent", Conn)
    DocIDPar_RTI = SIT_GetDocField(S_DocIDParent, "DocId", Conn)
    if len(DocIDParent_RTI) >0 and InStr(UCase(DocIDPar_RTI),"К_") > 0 Then
        S_AdditionalUsers = S_AdditionalUsers + VbCrLf + DocIDParent_RTI
    End If
    'если исходящее создается из поручения, то даем доступ автору поручения,исполнителю, соисполнителю.
    If S_DocIDParent <>"" and InStr(UCase(S_DocIDParent),"K_") > 0 Then 
        S_AdditionalUsers = S_AdditionalUsers + S_Author + VbCrLf + S_NameResponsible + VbCrLf + S_Correspondent + VbCrLf + S_NameControl
    End If
    'если исходящее создается из поручения, то даем доступ автору поручения,исполнителю, соисполнителю.

  End If
'minc

'vtss
  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" and InStr(UCase(Session("Department")), UCase(SIT_VTSS)) = 1 Then
    S_AdditionalUsers = """Подольский А. Е."" <vtss_a.podolskii>;" + VbCrLf
    DocIDParent_RTI = SIT_GetDocField(S_DocIDParent, "Correspondent", Conn)
    DocIDPar_RTI = SIT_GetDocField(S_DocIDParent, "DocId", Conn)
    if len(DocIDParent_RTI) >0 and InStr(UCase(DocIDPar_RTI),"T_") > 0 Then
        S_AdditionalUsers = S_AdditionalUsers + VbCrLf + DocIDParent_RTI
    End If
    'если исходящее создается из поручения, то даем доступ автору поручения,исполнителю, соисполнителю.
    If S_DocIDParent <>"" and InStr(UCase(S_DocIDParent),"T_") > 0 Then 
        S_AdditionalUsers = S_AdditionalUsers + S_Author + VbCrLf + S_NameResponsible + VbCrLf + S_Correspondent + VbCrLf + S_NameControl
    End If
    'если исходящее создается из поручения, то даем доступ автору поручения,исполнителю, соисполнителю.

  End If
'vtss

  If S_AdditionalUsers <> "" Then
     S_AdditionalUsers_Set = S_AdditionalUsers
  Else
     S_AdditionalUsers_Set = " "
  End If

  S_DocID_Set = S_DocID
  S_DocIDAdd_Set = S_DocIDAdd
  S_Department_Set = S_Department
  S_Author_Set = S_Author

  ' Создание карточки документа
  If UCase(Request("create")) = "Y" Then
     If Request("UpdateDoc") <> "YES" Then
        'SAY 2009-02-20
        VAR_ChangeDocGenerateButton=""
        S_DocID = " "
        If S_DocID_Set<>"" Then
           S_DocID_Set = S_DocID
        End If

        'SAY 2008-11-06
        S_Department_Set = Session("Department")

        'SAY 2008-09-22 для подчиненных документов оставляем весь список согласования родителя
        If Trim(S_DocIDParent)<>"" then
           S_ListToReconcile = Replace(S_ListToReconcile,"##;","")
        End If

        S_DocID = ""
        S_DocID_Set = " "
        S_DocIDadd_Set = " "
        VAR_DocFieldsNotToShow = "DocIDAdd"

        If S_DocIDParent = "" Then
           S_UserFieldText1 = SIT_Letter_Initiative
        Else
           S_DocIDParent_Set = S_DocIDParent
           S_UserFieldText1 = SIT_Letter_AnswerFor + S_DocIDParent
           'SAY 2009-03-20
           S_Correspondent = S_UserFieldText3
           S_UserFieldText3 = ""
           S_UserFieldText4 = ""
        End If
        S_UserFieldDate1 = MyDate(Date)
        S_DateActivation = MyDate(Date)

'Запрос №31 - СТС - start
        If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
           S_NameAproval = STS_Orders_CEO_SC
           S_LocationPath = " "
           S_LocationPath_Set = S_LocationPath
        Else
'Запрос №31 - СТС - end
           'SAY 2008-11-10 переделываем регистратора с роли на пользователя
		   If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
		      sDepartment1 = SIT_STS
		   ElseIf InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 Then
		      sDepartment1 = SIT_SIB
		   ElseIf InStr(UCase(Session("Department")), UCase(SIT_SITRU)) = 1 Then ' DmGorsky
		      sDepartment1 = SIT_SITRU
		   ElseIf InStr(UCase(Session("Department")), UCase(SIT_RTI)) = 1 Then
		      sDepartment1 = SIT_RTI
		   ElseIf InStr(UCase(Session("Department")), UCase(SIT_MINC)) = 1 Then
		      sDepartment1 = SIT_MINC		      
		   ElseIf InStr(UCase(Session("Department")), UCase(SIT_MIKRON)) = 1 Then
		      sDepartment1 = SIT_MIKRON
		   ElseIf InStr(UCase(Session("Department")), UCase(SIT_VTSS)) = 1 Then
		      sDepartment1 = SIT_VTSS
		   Else
		      sDepartment1 = SIT_SITRONICS
		   End If
		   
           S_LocationPath = ReplaceRoleFromDir(SIT_Registrar, sDepartment1)
           S_LocationPath_Set = S_LocationPath
'Запрос №31 - СТС - start
        End If
'Запрос №31 - СТС - end

        S_Author_Set = InsertionName(Session("Name"), Session("UserID"))
        S_Department_Set = Session("Department")
     End If
  End If 'UCase(Request("create")) = "Y"

  S_ListToReconcile = DeleteConstPrefixFromList(S_ListToReconcile)

  CurrentDocFieldOrder = "Name, UserFieldText1, DocIDParent, UserFieldDate1, DocID,DocIDAdd, NameCreation, Author, DateActivation, Content, UserFieldText3, UserFieldText4, ListToReconcile, NameAproval, LocationPath, Correspondent, PartnerName, UserFieldText7, UserFieldText2, SecurityLevel, Department"
  CurrentDocRequiredFields = "Name, UserFieldText1, UserFieldDate1, DateActivation, UserFieldText3, NameAproval, DocListToRegister, UserFieldText5"
  
  'отмена справочников для пользовательских полей
  VAR_DirPictNotToShow = "UserFieldText1, UserFieldText3, UserFieldText4"

  'SAY 2008-09-03 доступ только у рук. канцелярии и помощника президента
  '20110212 ReplaceRoleFromDir. Учесть, если будут меняться справочники ролей
  sChiefName1 = GetUserDirValue(SIT_RolesDirSitronics, SIT_HeadOfDocControl, 1, 2)
  sChiefName2 = GetUserDirValue(SIT_RolesDirSitronics, SIT_AssistantOfPresident, 1, 2)

  If InStr(UCase(sChiefName1), UCase(Session("UserID"))) = 0 and InStr(UCase(sChiefName2), UCase(Session("UserID"))) = 0 and not IsAdmin() Then
     S_DateActivation_Set = S_DateActivation
     CurrentDocRequiredFields = Replace(CurrentDocRequiredFields, "DateActivation", "")
  End If

  'SAY 2008-09-09
  S_LocationPath_Set = S_LocationPath

  'SAY 2008-09-14, SAY 2008-10-07
  ' DmGorsky_5: меняем местами проверку: вначале ищем SIT_SITRU ("СИТРОНИКС ИТ"), потом, если не найдено, - SIT_SITRONICS ("СИТРОНИКС")
  If InStr(UCase(Session("Department")), UCase(SIT_SITRU)) = 1 then ' DmGorsky
     S_ListToReconcile_Comment = oPayDox.GetExtTableValue("AgreeSITRU","Category",Session("CurrentClassDoc")+"/","List") ' DmGorsky
     VAR_DocFieldsNotToShow = "UserFieldText7" ' DmGorsky
  ElseIf InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) = 1 Then  
     S_ListToReconcile_Comment = oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")
  ElseIf InStr(UCase(Session("Department")), UCase(SIT_RTI)) = 1 Then  
     S_ListToReconcile_Comment =SIT_RequiredAgrees + oPayDox.GetExtTableValue("AgreeRTIRU","Category",Session("CurrentClassDoc")+"/","List")    
  ElseIf InStr(UCase(Session("Department")), UCase(SIT_MINC)) = 1 Then  
     S_ListToReconcile_Comment =SIT_RequiredAgrees + oPayDox.GetExtTableValue("AgreeMINCRU","Category",Session("CurrentClassDoc"),"List")
 
     
     'SAY 2008-10-07
     VAR_DocFieldsNotToShow = "UserFieldText7"
  ElseIf InStr(UCase(Session("Department")), UCase(SIT_MIKRON)) = 1 Then  
     S_ListToReconcile_Comment =SIT_RequiredAgrees + oPayDox.GetExtTableValue("AgreeMIKRON","Category",Session("CurrentClassDoc")+"/","List")
     VAR_DocFieldsNotToShow = "UserFieldText7"
  ElseIf InStr(UCase(Session("Department")), UCase(SIT_VTSS)) = 1 Then  
     S_ListToReconcile_Comment =SIT_RequiredAgrees + oPayDox.GetExtTableValue("AgreeVTSSRU","Category",Session("CurrentClassDoc"),"List")
'Запрос №31 - СТС - start
  ElseIf InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 then
'ph - 20101231 - start
     CurrentDocRequiredFields = CurrentDocRequiredFields & ", UserFieldText7"
'ph - 20101231 - end
     If UCase(Request("create")) <> "Y" Then
        'Запрещаем редактировать БЕ, от нее зависит номер
        If Request("UpdateDoc") <> "YES" Then
           If S_AddField2 = "" Then
              S_AddField2 = MyCStr(dsDoc("BusinessUnit"))
           End If
        Else
           If S_AddField2 = "" Then
              S_AddField2 = Request("BusinessUnit")
           End If
        End If
        If S_AddField2 = "" Then
           S_AddField2 = " "
        End If
        S_AddField_Set2 = S_AddField2
     End If
'Запрос №31 - СТС - end
  End If

  VAR_DirPictNotToShow = VAR_DirPictNotToShow+", DocName"
  'rmanyushin 60298 02.11.2009 Start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", ContractType"
  'rmanyushin 60298 02.11.2009 End
'ph - 20120216 - start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", AddFieldText1, AddFieldText2"
'ph - 20120216 - end

  'Выпадающие списки
  Select Case UCase(Request("l"))
    Case "RU" 'RU
      S_UserFieldText2_Select = GetUserDirValues("{6D42F0FB-1389-4CA0-8A00-9E0CD3F09CF7}")
'Запрос №1 - СИБ - start
      If InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 then
         S_UserFieldText7_Select = GetUserDirValues("{D0E4DEEB-A1A8-48B0-B527-9573856769DB}")
      'If InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) <> 1 then
      ElseIf InStr(UCase(Session("Department")), UCase(SIT_SITRU)) = 1 then ' DmGorsky_3
         S_UserFieldText7_Select = GetUserDirValues("{D68C37BE-9EBA-4CA8-B0D2-C6369123E7C7}") ' DmGorsky_3
'Запрос №1 - СИБ - start
      ElseIf InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) <> 1 then
         S_UserFieldText7_Select = GetUserDirValues("{da5960be-a65d-4d21-bf89-73233ffeaee8}")
      End If
    Case "" 'EN
      S_UserFieldText2_Select = GetUserDirValues("{6D57662F-7DD0-41E1-806B-3562412FDFAF}")
'Запрос №1 - СИБ - start
      If InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 then
         S_UserFieldText7_Select = GetUserDirValues("{D0E4DEEB-A1A8-48B0-B527-9573856769DB}")
      'If InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) <> 1 then
'Запрос №1 - СИБ - start
      ElseIf InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) <> 1 then
         S_UserFieldText7_Select = GetUserDirValues("{6A200BD7-1A53-40FC-9DBB-44499F65B74C}") '------------- новый справочник
      End If
    Case "3" 'CZ
      S_UserFieldText2_Select = GetUserDirValues("{0D620DAB-1B89-4E7B-BB6A-29EB77F9AEE9}")
'Запрос №1 - СИБ - start
      If InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 then
         S_UserFieldText7_Select = GetUserDirValues("{D0E4DEEB-A1A8-48B0-B527-9573856769DB}")
      'If InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) <> 1 then
'Запрос №1 - СИБ - start
      ElseIf InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) <> 1 then
         S_UserFieldText7_Select = GetUserDirValues("{2F4D0C04-FD15-4321-A5E3-5AA2FCB0D70E}") '------------- новый справочник
      End If
  End Select

  'Для пользователей СТС везде добавляем поле бизнес единица
  If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
     If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 Then
        If UCase(Request("create")) = "Y" Then
           sCreator = Session("UserID")
        Else
           sCreator = GetUserID(S_Author)
        End If
        BusinessUnitsList = GetUsersBusinessUnits(sCreator)
        If InStr(BusinessUnitsList, VbCrLf) > 0 Then
           S_AddField2 = " "
           S_AddField_Select2 = BusinessUnitsList
           'Поле обязательно для заполнения
           CurrentDocRequiredFields = CurrentDocRequiredFields+", BusinessUnit"
        Else
           S_AddField_Set2 = BusinessUnitsList
        End If
     End If
  Else
     'Поле показывается только для пользователей СТС и администратора (у админа не из СТС оно необязательное, редактируемое)
     If not IsAdmin() Then
        VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", BusinessUnit, UserFieldText7"
     End If
  End If

'rti_ishodyaschie
  If InStr(UCase(Session("Department")), UCase(SIT_RTI)) = 1 Then
     CurrentDocFieldOrder = "Name, UserFieldText1, DocIDParent,UserFieldText8, UserFieldDate1, DocID,DocIDAdd, NameCreation, Author, DateActivation, Content, UserFieldText3, UserFieldText4, ListToReconcile, NameAproval, LocationPath, Correspondent, PartnerName, UserFieldText7, UserFieldText2, SecurityLevel, Department"
     CurrentDocRequiredFields = "Name, UserFieldText1, UserFieldDate1,UserFieldText8, DateActivation, UserFieldText3, NameAproval, DocListToRegister, UserFieldText5"
     S_UserFieldText8_Select = GetUserDirValues("{9EB619A3-61E8-4C1A-9573-27C87DABEF76}")
  End If
'rti_ishodyaschie

'vtss_ishodyaschie
  If InStr(UCase(Session("Department")), UCase(SIT_VTSS)) = 1 Then
     CurrentDocFieldOrder = "Name, UserFieldText1, DocIDParent,UserFieldDate1, DocID,DocIDAdd, NameCreation, Author, DateActivation, Content, UserFieldText3, UserFieldText4, ListToReconcile, NameAproval, LocationPath, Correspondent, PartnerName, UserFieldText7, UserFieldText2, SecurityLevel, Department"
     CurrentDocRequiredFields = "Name, UserFieldText1, UserFieldDate1,DateActivation, UserFieldText3, NameAproval, DocListToRegister, UserFieldText5"   
     VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", UserFieldText8"
  End If
'vtss_ishodyaschie

'oaorti_ishodyaschie
  If InStr(UCase(Session("Department")), UCase(SIT_MINC)) = 1 Then
  S_Name_Select = GetUserDirValues("{7CD9F979-A1CB-48C5-8A6C-AC4C1E950693}")
  CurrentDocFieldOrder = "Name, UserFieldText1, DocIDParent, UserFieldDate1, DocID,DocIDAdd, NameCreation, Author, DateActivation, Content, UserFieldText3, UserFieldText4, ListToReconcile, NameAproval, LocationPath, Correspondent, PartnerName, UserFieldText7, UserFieldText2, SecurityLevel, Department"
  '2015-04-09 kkoskhin скрываем поля "утверждающий" и "список согласующих"
  CurrentDocRequiredFields = "Name, UserFieldText1, UserFieldDate1, DateActivation, UserFieldText3, NameAproval, DocListToRegister, UserFieldText5, PartnerName, Content"
  'CurrentDocRequiredFields = "Name, UserFieldText1, UserFieldDate1, DateActivation, UserFieldText3, DocListToRegister, UserFieldText5, PartnerName, Content"
  'VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", NameAproval, ListToReconcile"
  '2015-04-09 end
  S_UserFieldText2 = ""
          If Trim(S_DocIDParent)<>"" and   UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" then
           S_Name = ""
          End If
  End If
'oaorti_ishodyaschie
'amw 06-08-2014
'  If InStr(UCase(Session("Department")), UCase(SIT_MIKRON)) = 1 Then
'     CurrentDocFieldOrder = "Name, UserFieldText1, DocIDParent,UserFieldText8, UserFieldDate1, DocID,DocIDAdd, NameCreation, Author, DateActivation, Content, UserFieldText3, UserFieldText4, ListToReconcile, NameAproval, LocationPath, Correspondent, PartnerName, UserFieldText7, UserFieldText2, SecurityLevel, Department"
'     CurrentDocRequiredFields = "Name, UserFieldText1, UserFieldDate1,UserFieldText8, DateActivation, UserFieldText3, NameAproval, DocListToRegister, UserFieldText5"
'     S_UserFieldText8_Select = GetUserDirValues("{9EB619A3-61E8-4C1A-9573-27C87DABEF76}")
'  End If
'mikron_ishodyaschie

' ********************************************************************************* 
' ***                       РАСПОРЯДИТЕЛЬНЫЕ ДОКУМЕНТЫ                          ***
' ********************************************************************************* 
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_RASP_DOCS)) = 1 Then
  'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
  S_SecurityLevel = 4
  If not IsAdmin() Then
     S_SecurityLevel_Set = S_SecurityLevel
  End If

  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
     S_AdditionalUsers = ""
  End If
  
    If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" and InStr(UCase(Session("Department")), UCase(SIT_RTI)) = 1 Then
    S_AdditionalUsers = """Боев С. Ф."" <boev_oaorti>;" + VbCrLf + """Двойченкова О. А."" <Dvoychenkova_oaorti>;" + VbCrLf + """Нестерова А. Н."" <nesterova>;"
    If InStr(UCase(Session("Department")), UCase("Управление ИТ")) > 0 Then
      S_AdditionalUsers = S_AdditionalUsers + VbCrLf + """Иванов В. В."" <vivanon_oaorti>;"
    End If
    DocIDParent_RTI = SIT_GetDocField(S_DocIDParent, "Correspondent", Conn)
    DocIDPar_RTI = SIT_GetDocField(S_DocIDParent, "DocId", Conn)
    if len(DocIDParent_RTI) >0 and InStr(UCase(DocIDPar_RTI),"T_") > 0 Then
        S_AdditionalUsers = S_AdditionalUsers + VbCrLf + DocIDParent_RTI
     End If
  End If

'amw mikron - start
  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" and InStr(UCase(Session("Department")), UCase(SIT_MIKRON)) = 1 Then
     S_AdditionalUsers = """Козырева Н. В."" <nkozyreva_ms>;" + VbCrLf + ReplaceRoleFromDir(MIKRON_GenDirector,SIT_MIKRON)
     DocIDParent_RTI = SIT_GetDocField(S_DocIDParent, "Correspondent", Conn)
     DocIDPar_RTI = SIT_GetDocField(S_DocIDParent, "DocId", Conn)
     if len(DocIDParent_RTI) >0 and InStr(UCase(DocIDPar_RTI),"T_") > 0 Then
        S_AdditionalUsers = S_AdditionalUsers + VbCrLf + DocIDParent_RTI
     End If
  End If
'amw mikron - end

  If S_AdditionalUsers <> "" Then
     S_AdditionalUsers_Set = S_AdditionalUsers
  Else
     S_AdditionalUsers_Set = " "
  End If

  S_DocID_Set = S_DocID
  S_DocIDAdd_Set = S_DocIDAdd
  S_Department_Set = S_Department
  S_Author_Set = S_Author

  ' Создание карточки документа
  If UCase(Request("create")) = "Y" Then
     If Request("UpdateDoc") <> "YES" Then
        'SAY 2009-02-20
        VAR_ChangeDocGenerateButton=""
        S_DocID = " "
        If S_DocID_Set<>"" Then
           S_DocID_Set = S_DocID
        End If

        'SAY 2008-11-06
        S_Department_Set = Session("Department")

        'SAY 2008-09-22 для подчиненных документов оставляем весь список согласования родителя
        If Trim(S_DocIDParent)<>"" then
           S_ListToReconcile = Replace(S_ListToReconcile,"##;","")
        End If

        S_DocID_Set = " "
        S_DocIDadd_Set = " "
        'vnik_rasp_norm_doc
        VAR_DocFieldsNotToShow = "DocIDAdd,DocIDParent, AddFieldText2"    
        'vnik_rasp_norm_doc
        S_DateActivation = MyDate(Date)
        S_UserFieldDate1 = MyDate(Date)
        S_Author_Set = InsertionName(Session("Name"), Session("UserID"))
        S_Department_Set = Session("Department")

        'vnik_rasp_norm_doc
        VNIK_WeekDay_Number = Trim(WeekDay(Date))
        If (VNIK_WeekDay_Number = 4) or (VNIK_WeekDay_Number = 5) or (VNIK_WeekDay_Number = 6) Then
           S_UserFieldDate2 = MyDate(Date+5)
        ElseIf (VNIK_WeekDay_Number = 7) Then
           S_UserFieldDate2 = MyDate(Date+4)
        Else
           S_UserFieldDate2 = MyDate(Date+3)
        End If
        'vnik_rasp_norm_doc  

        ' DmGorsky_5: меняем местами проверку: вначале ищем SIT_SITRU ("СИТРОНИКС ИТ"), потом, если не найдено, - SIT_SITRONICS ("СИТРОНИКС")
        If InStr(UCase(Session("Department")), UCase(SIT_SITRU)) = 1 Then ' DmGorsky
	       S_NameAproval = SITRU_GenDirector ' DmGorsky
           S_LocationPath = ReplaceRoleFromDir(SITRU_Registrar, SIT_SITRU) ' DmGorsky
           S_LocationPath_Set = S_LocationPath
        ElseIf InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
           S_LocationPath = " "
           S_LocationPath_Set = S_LocationPath
           S_NameAproval = STS_Orders_CEO_SC
           If UCase(Request("l")) = "RU" Then
              S_Content = STS_PRIKAZ_TEXT
           End If
        ElseIf InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 Then
           S_NameAproval = SIB_GenDirector
           S_LocationPath = ReplaceRoleFromDir(SIB_Registrar, SIT_SIB)
           S_LocationPath_Set = S_LocationPath
        ElseIf InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) = 1 Then
           S_NameAproval = SIT_President
           S_LocationPath = ReplaceRoleFromDir(SIT_Registrar, iif(InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1, SIT_STS, SIT_SITRONICS))
           S_LocationPath_Set = S_LocationPath
        ElseIf InStr(UCase(Session("Department")), UCase(SIT_RTI)) = 1 Then
           S_NameAproval = RTI_President
           S_LocationPath = ReplaceRoleFromDir(SIT_Registrar, SIT_RTI)
           'If InStr(UCase(Session("Department")), UCase(RTI_DVKiA)) > 0 Then
           '     S_ListToReconcile_Comment =SIT_RequiredAgrees + VbCrLf + """#Управляющий делами"";""#ЗГД – Начальник управления безопасности и режима"";"""
           ' else
                S_ListToReconcile_Comment =SIT_RequiredAgrees + VbCrLf + oPayDox.GetExtTableValue("AgreeRTIRU","Category",Session("CurrentClassDoc")+"/","List")
           'End If
        ElseIf InStr(UCase(Session("Department")), UCase(SIT_MIKRON)) = 1 Then
           S_NameAproval = MIKRON_GenDirector
           S_LocationPath = ReplaceRoleFromDir(SIT_Registrar, SIT_MIKRON)
           S_LocationPath_Set = S_LocationPath
           S_ListToReconcile_Comment =SIT_RequiredAgrees + VbCrLf + oPayDox.GetExtTableValue("AgreeMIKRON","Category",Session("CurrentClassDoc")+"/","List")
        End If
     End If
  End If 'UCase(Request("create")) = "Y"

  S_ListToReconcile = DeleteConstPrefixFromList(S_ListToReconcile)

  CurrentDocFieldOrder = "Name, UserFieldText1, DocIDParent, UserFieldDate1, DocID, Author, DateActivation, Content, UserFieldText2, ListToReconcile, NameAproval, LocationPath, ListToView, SecurityLevel, Department, UserFieldDate2"	
  CurrentDocRequiredFields = "Name, UserFieldText1, UserFieldDate1, UserFieldText2, NameAproval, DocListToRegister"

  'SAY 2008-09-09
  S_LocationPath_Set = S_LocationPath

  ' DmGorsky_5: меняем местами проверку: вначале ищем SIT_SITRU ("СИТРОНИКС ИТ"), потом, если не найдено, - SIT_SITRONICS ("СИТРОНИКС")
  If InStr(UCase(Session("Department")), UCase(SIT_SITRU)) = 1 Then ' DmGorsky

   ' DmGorsky жесткая часть маршрута согласования распорядительных документов
   ' (обязательные согласующие)
    'If InStr(UCase(S_UserFieldText1), "OR") > 0 Then ' DmGorsky_5
    '  If Trim(S_UserFieldText1) = SIT_Orders_Prikaz_ND Then ' DmGorsky_5
    '    S_ListToReconcile_Comment = "Для Приказа об утверждении Нормативного документа список обязательных согласующих: " + oPayDox.GetExtTableValue("AgreeSITRU","Category","Нормативные документы*Regulations*Ridici dokumenty/","List") +VbCrLf+ "Для остальных видов Приказов список обязательных согласующих: " + oPayDox.GetExtTableValue("AgreeSITRU","Category",Session("CurrentClassDoc")+"/","List") ' DmGorsky_5
    '  Else ' добавляем еще и роль "#Руководитель ОАРБП""; ' DmGorsky_5
    '    S_ListToReconcile_Comment = oPayDox.GetExtTableValue("AgreeSITRU","Category",Session("CurrentClassDoc")+"/","List")+VbCrLf+"""#Руководитель ОАРБП"";" ' DmGorsky_5
    '  End If ' DmGorsky_5
    'Else ' DmGorsky_5
      S_ListToReconcile_Comment = "Для Приказа об утверждении Нормативного документа список обязательных согласующих: " + oPayDox.GetExtTableValue("AgreeSITRU","Category","Нормативные документы*Regulations*Řídicí dokumenty/","List") +VbCrLf+ VbCrLf+ "Для остальных видов Приказов список обязательных согласующих: " + oPayDox.GetExtTableValue("AgreeSITRU","Category",Session("CurrentClassDoc")+"/","List") ' DmGorsky_5
    'End If ' DmGorsky_5

  ElseIf InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) = 1 Then
    ' SAY 2008-09-14 жесткая часть маршрута согласования
     If InStr(UCase(S_UserFieldText1), "OR") > 0 Then
        'vnik_rasp_norm_doc
        If Trim(S_UserFieldText1) = SIT_Orders_Prikaz_ND Then
           S_ListToReconcile_Comment = "Для Приказа об утверждении Нормативного документа список обязательных согласующих: " + oPayDox.GetExtTableValue("Agree"+Request("l"),"Category","Нормативные документы*Regulations*Ridici dokumenty/","List") +VbCrLf+ "Для остальных видов Приказов список обязательных согласующих: " + oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")
        Else
           S_ListToReconcile_Comment = oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+VbCrLf+"""#Вице-президент инициатора"";"
        End If
     Else
        S_ListToReconcile_Comment = "Для Приказа об утверждении Нормативного документа список обязательных согласующих: " + oPayDox.GetExtTableValue("Agree"+Request("l"),"Category","Нормативные документы*Regulations*Ridici dokumenty/","List") +VbCrLf+ "Для остальных видов Приказов список обязательных согласующих: " + oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")
        'vnik_rasp_norm_doc  
     End If   

'Запрос №31 - СТС - start
  ElseIf InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 then
     If UCase(Request("create")) <> "Y" Then
        'Запрещаем редактировать БЕ, от нее зависит номер
        If Request("UpdateDoc")<>"YES" Then
           If S_AddField2 = "" Then
              S_AddField2 = MyCStr(dsDoc("BusinessUnit"))
           End If
        Else
           If S_AddField2 = "" Then
              S_AddField2 = Request("BusinessUnit")
           End If
        End If
        If S_AddField2 = "" Then
           S_AddField2 = " "
        End If
        S_AddField_Set2 = S_AddField2
     End If
'Запрос №31 - СТС - end

'Запрос №1 - СИБ - start
  ElseIf InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 Then
     If InStr(UCase(S_UserFieldText1), "OR") > 0 Then
        If Trim(S_UserFieldText1) = SIT_Orders_Prikaz_ND Then
           S_ListToReconcile_Comment = "Для Приказа об утверждении Нормативного документа список обязательных согласующих: " + oPayDox.GetExtTableValue("AgreeSIBRU","Category","Нормативные документы*Regulations*Řídicí dokumenty/","List") +VbCrLf+ "Для остальных видов Приказов список обязательных согласующих: " + oPayDox.GetExtTableValue("AgreeSIBRU","Category",Session("CurrentClassDoc")+"/","List")
        Else
           S_ListToReconcile_Comment = oPayDox.GetExtTableValue("AgreeSIBRU","Category",Session("CurrentClassDoc")+"/","List")
        End If
     Else
        S_ListToReconcile_Comment = "Для Приказа об утверждении Нормативного документа список обязательных согласующих: " + oPayDox.GetExtTableValue("AgreeSIBRU","Category","Нормативные документы*Regulations*Ridici dokumenty/","List") +VbCrLf+ "Для остальных видов Приказов список обязательных согласующих: " + oPayDox.GetExtTableValue("AgreeSIBRU","Category",Session("CurrentClassDoc")+"/","List")
     End If   
  End If
'Запрос №1 - СИБ - end

  If InStr(UCase(Session("Department")), UCase(SIT_RTI)) = 1 Then
     VAR_DocFieldsNotToShow = "DocIDAdd,AddFieldText2"
  end if

  VAR_DirPictNotToShow = VAR_DirPictNotToShow+", DocName"
'rmanyushin 60298 02.11.2009 Start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", ContractType"
'rmanyushin 60298 02.11.2009 End
'ph - 20120216 - start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", AddFieldText1"
'ph - 20120216 - end

  'Выпадающие списки
  Select Case UCase(Request("l"))
     Case "RU" 'RU
        S_UserFieldText1_Select = GetUserDirValues("{E0E79CEC-5DDE-4184-92BE-85556566BD14}")
        S_UserFieldText2_Select = GetUserDirValues("{37E16CD5-BC8F-4D0C-9569-D14DAA895440}")
     Case "" 'EN
        S_UserFieldText1_Select = GetUserDirValues("{3B0BABA9-EF20-47A0-A026-08DA34B9A7F7}")
        S_UserFieldText2_Select = GetUserDirValues("{7C7058BE-F586-44C4-B5BE-47D2E05E96BD}")
     Case "3" 'CZ
        S_UserFieldText1_Select = GetUserDirValues("{D733E4E8-0418-4B99-BCF5-8FC5CB9C5C42}")
        S_UserFieldText2_Select = GetUserDirValues("{AF173F9A-4724-405B-AA9C-C72E1DCA7647}")
  End Select
'Запрос №31 - СТС - start
'Для СТС убираем распоряжения из списка возможных типов распорядительных докуметов
  If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
     S_UserFieldText1_Select = Replace(S_UserFieldText1_Select, SIT_Orders_Rasporyajenie, "")
     S_UserFieldText1_Select = Replace(S_UserFieldText1_Select, VbCrLf&VbCrLf, VbCrLf)
  End If
'Запрос №31 - СТС - end

  'Для пользователей СТС везде добавляем поле бизнес единица
  If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
     If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 Then
        If UCase(Request("create")) = "Y" Then
           sCreator = Session("UserID")
        Else
           sCreator = GetUserID(S_Author)
        End If
        BusinessUnitsList = GetUsersBusinessUnits(sCreator)
        If InStr(BusinessUnitsList, VbCrLf) > 0 Then
           S_AddField2 = " "
           S_AddField_Select2 = BusinessUnitsList
           'Поле обязательно для заполнения
           CurrentDocRequiredFields = CurrentDocRequiredFields+", BusinessUnit"
        Else
           S_AddField_Set2 = BusinessUnitsList
        End If
     End If
  Else
'Поле показывается только для пользователей СТС и администратора
'(у админа не из СТС оно необязательное, редактируемое)
     If not IsAdmin() Then
        VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", BusinessUnit"
     End If
  End If
  
  If InStr(UCase(Session("Department")), UCase(SIT_RTI)) = 1 Then
     VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow + ", AddFieldText2, UserFieldText2"
     S_UserFieldText1_Select = GetUserDirValues("{8CF4C5D4-C76A-46A8-B043-7E6FD9466372}")
     CurrentDocRequiredFields = "Name, UserFieldText1, UserFieldDate1,  NameAproval, DocListToRegister"
  End If
' amw 16-04-2013
  If InStr(UCase(Session("Department")), UCase(SIT_MIKRON)) = 1 Then
     VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow + ", AddFieldText2, UserFieldText2"
     S_UserFieldText1_Select = GetUserDirValues(MIKRON_CATALOG_ORDERS_TYPES)
     CurrentDocRequiredFields = "Name, UserFieldText1, UserFieldDate1,  NameAproval, DocListToRegister"
  End If

' ********************************************************************************* 
' ***                               ПРОТОКОЛЫ                                   ***
' ********************************************************************************* 
' *** ПРОТОКОЛ ЦЗК РТИ
' rti_protocol 
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(RTI_PROTOCOL)) > 0 Then
  'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
  S_SecurityLevel = 4
  If not IsAdmin() Then
    S_SecurityLevel_Set = S_SecurityLevel
  End If

  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
    S_AdditionalUsers = ""
  End If
  
  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" and InStr(UCase(Session("Department")), UCase(SIT_RTI)) = 1 Then
    S_AdditionalUsers = """Боев С. Ф."" <boev_oaorti>;"
    If InStr(UCase(Session("Department")), UCase("Управление ИТ")) > 0 Then
      S_AdditionalUsers = S_AdditionalUsers + VbCrLf + """Иванов В. В."" <vivanon_oaorti>;"
    End If
  End If
  
  If S_AdditionalUsers <> "" Then
    S_AdditionalUsers_Set = S_AdditionalUsers
  Else
    S_AdditionalUsers_Set = " "
  End If

  S_DocID_Set = S_DocID
  S_DocIDAdd_Set = S_DocIDAdd
  S_Department_Set = S_Department
  S_Author_Set = S_Author

  ' Создание карточки документа
  If UCase(Request("create")) = "Y" Then
    If Request("UpdateDoc") <> "YES" Then
      'SAY 2009-02-20
      VAR_ChangeDocGenerateButton=""
      S_DocID = " "
      If S_DocID_Set<>"" Then
        S_DocID_Set = S_DocID
      End If

      'SAY 2008-11-06
      S_Department_Set = Session("Department")

      'SAY 2008-09-22 для подчиненных документов оставляем весь список согласования родителя
      If Trim(S_DocIDParent)<>"" and InStr(UCase(Session("CurrentClassDoc")),UCase(RTI_PROTOCOL)) = 0 then
        S_ListToReconcile = Replace(S_ListToReconcile,"##;","")
      End If

      If Trim(S_DocIDParent)<>"" and InStr(UCase(Session("CurrentClassDoc")),UCase(RTI_PROTOCOL)) > 0 then
        S_ListToView = S_ListToView + " " + S_Author  
      End If

      S_DocID_Set = " "
      S_DocIDadd_Set = " "
      VAR_DocFieldsNotToShow = "DocIDAdd"
      S_UserFieldDate2 = MyDate(Date)
      
      S_Author_Set = InsertionName(Session("Name"), Session("UserID"))
      S_Department = Session("Department")
      S_Department_Set = Session("Department")
      S_NameAproval = RTI_ChairmanOfCPC
      S_NameAproval_Set = S_NameAproval
    End If
  End If 'UCase(Request("create")) = "Y"
  

  S_ListToReconcile = DeleteConstPrefixFromList(S_ListToReconcile)
  'Добавим комментарий по обязательным согласующим
  S_ListToReconcile_Comment = SIT_RequiredAgrees + vbCrLf + oPayDox.GetExtTableValue("AgreeRTI"+Request("l"),"Category",Session("CurrentClassDoc"),"List") 

  CurrentDocFieldOrder = "Name, DocIDParent, DocID, Author, UserFieldDate2, Content, ListToReconcile, NameAproval, ListToView, SecurityLevel, Department"	
  ' поля помечаются звездочкой - обязательно к заполнению
  CurrentDocRequiredFields = "Name, UserFieldDate2, NameAproval"

  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", DateActivation, BusinessUnit, DocName, ContractType, AddFieldText1, AddFieldText2"
'rti_protocol

' ********************************************************************************* 
' ***                               ПРОТОКОЛЫ                                   ***
' ********************************************************************************* 
' *** ПРОТОКОЛ ЗК МИКРОН
' mikron_protocol
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_PROTOCOL)) > 0 Then
    'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
    S_SecurityLevel = 4
    If not IsAdmin() Then
       S_SecurityLevel_Set = S_SecurityLevel
    End If

    'Создание карточки документа
    If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
       S_DocID = " "
       If S_DocID_Set <> "" Then
          S_DocID_Set = S_DocID
       End If

       'Дополнительные пользователи: Финансовый директор
       S_AdditionalUsers = ""
       
       sTemp = SIT_GetDocField(S_DocIDParent,"NameCreation",Conn)
       S_AddFieldText1 = Left(sTemp, InStr(sTemp,"*") - 1) + Mid(sTemp, InStrRev( sTemp,"*") +1 )
       If S_AddFieldText1 <> "" Then
          sTemp = S_AddFieldText1
       End If
       S_AdditionalUsers = sTemp +  ";" + VbCrLf + ReplaceRoleFromDir(MIKRON_CFO,SIT_MIKRON)

       VAR_ChangeDocGenerateButton = ""
       VAR_DocFieldsNotToShow = "DocIDAdd"

       'Утверждающий - Председатель ЗК
       S_NameAproval = MIKRON_ChairmanOfPC

       S_UserFieldText2 = "" 'Вопрос, вынесенный на голосование
       S_UserFieldMoney1 = 0 'Всего проголосовало

       S_UserFieldText3 = """ЗА""- , ""ПРОТИВ""- , ""ВОЗДЕРЖАЛСЯ""- "
       S_UserFieldDate1 = MyDate(Date) 'Дата окончания голосования
       S_UserFieldDate2 = MyDate(Date) 'Дата проведения заседания комиссии

       S_ListToView = oPayDox.GetExtTableValue("AgreeMIKRON","Name","Список согласующих протокол ЗК МИКРОН","List")
       
       'Если родительский документ Опросный лист, форма проведения ЗАОЧНАЯ
       'заполняем поля из родительского документа
       If InStr( UCase(Trim(S_DocIDParent)), UCase("RL") ) <> 0 Then
          S_UserFieldText1 = MIKRON_Form_OfPC_1  'ЗАОЧНЫЙ способ проведения ЗК
          
          'Вычисляем имя заявки на закупку, контрагента и сумму закупки
          vnik_DocIDParent = SIT_GetDocField(S_DocIDParent, "DocIDParent", Conn)
          S_Name = SIT_GetDocField(vnik_DocIDParent, "Name", Conn) + " Протокол ЗК"
          S_Partner = SIT_GetDocField(vnik_DocIDParent, "PartnerName", Conn)
          S_PartnerName_Set = S_Partner

          'Вычисляем вопрос, поставленный на голосование
          S_UserFieldText2 = "Выбор КП от поставщика " + S_PartnerName + _
                             ". Сумма предложения " + CStr(S_AmountDoc) + S_Currency
          
       'Если родительский документ Заявка на закупку, форма проведения ОЧНАЯ 
       'создатель - секретарь ЗК, надо заполнять поля руками.
       Else
          S_UserFieldText1 = MIKRON_Form_OfPC_2   'ОЧНЫЙ способ проведения ЗК
          S_Name = S_Name + " Протокол ЗК"
          S_UserFieldText2 = "Выбор предложения от контрагента  . Сумма предложения "
          VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow + ",UserFieldDate1"
'amw 28.09.2013 start      
          'Дополнительные пользователи: Руководители ЦФЗ и ЦФО (вычисляем из Заявки на закупку)
          S_AddFieldText1 = SIT_GetDocField(S_DocIDParent, "UserFieldText2", Conn)
          'Берем поле №3 'Руководитель ЦФЗ' из Справочника "Закупки МИКРОН/ЦФЗ"
          VNIK_TempValueAgree = GetUserDirValuesVNIK("{3C047498-29BE-4F60-8940-40CFC6ED702F}","Field3","Field1",S_AddFieldText1)
          S_AdditionalUsers = S_AdditionalUsers + VbCrLf + VNIK_TempValueAgree

          S_AddFieldText1 = SIT_GetDocField(S_DocIDParent, "UserFieldText5", Conn)
          'Берем поле №3 'Руководитель ЦФO' из Справочника "Закупки МИКРОН/ЦФO"
          VNIK_TempValueAgree = GetUserDirValuesVNIK("{7302B7EF-EB19-40E4-BAD8-2163207DB147}","Field3","Field1",S_AddFieldText1)
          S_AdditionalUsers = S_AdditionalUsers + VbCrLf + VNIK_TempValueAgree
          S_AddFieldText1 = ""
'amw 28.09.2013 end
       End If
    End If 'UCase(Request("create")) = "Y"

    'заполнение нередактируемых полей
    S_NameAproval_Set = S_NameAproval
    S_AdditionalUsers_Set = DeleteUserDoublesInList(S_AdditionalUsers)
    S_DocID_Set = S_DocID
    S_UserFieldText1_Select = MIKRON_Form_OfPC
    S_Department_Set = S_Department

    'определяем порядок следования полей
    CurrentDocFieldOrder = "Name,DocID,DocIdParent,UserFieldText1,UserFieldText2,UserFieldMoney1," + _
                           "UserFieldText3,UserFieldDate1,UserFieldDate2,NameAproval,ListToView,Content"

    'поля помечаются звездочкой - обязательно к заполнению
    CurrentDocRequiredFields = "Name,UserFieldText2,UserFieldMoney1,UserFieldText6,UserFieldText3," + _
                               "UserFieldText4,Currency,PartnerName,AmountDoc"

    'Не показываем кнопку выбора из справочника
    VAR_DirPictNotToShow = "DocName,AmountDoc,QuantityDoc,UserFieldMoney1,UserFieldText3," + _
                           "UserFieldText4,UserFieldText5,UserFieldText6,UserFieldText7"

    VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+",NameControl"
    VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ",AddFieldText1,AddFieldText2,BusinessUnit,ContractType"
' mikron_protocol
   
' ********************************************************************************* 
' ***                      ОПРОСНЫЙ ЛИСТ ДЛЯ ПЗК МИКРОН                         ***
' *** ОЛ согласуется членами ЗК в формате ЗА, ПРОТИВ, ВОЗДЕРЖАЛСЯ и приходит 
' *** на утверждение Секретарю ЗК, для создания протокола. Необходимо собрать и 
' *** посчитать все голоса, а также не дать нажать кнопку "Отказать", чтобы 
' *** согласование не остановилось.
' ********************************************************************************* 
' mikron_RL_protocol 
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_RL_PROTOCOL)) > 0 Then
    'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
    S_SecurityLevel = 4
    If not IsAdmin() Then
       S_SecurityLevel_Set = S_SecurityLevel
    End If

    'Дополнительные пользователи: Председатель ЗК, Руководители ЦФЗ и ЦФО (вычисляем из Заявки на закупку)
    If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
       S_AdditionalUsers = ""
       S_AdditionalUsers = ReplaceRoleFromDir(MIKRON_ChairmanOfPC,SIT_MIKRON)

       S_AddFieldText1 = SIT_GetDocField(S_DocIDParent, "UserFieldText2", Conn)
       'Берем поле №3 'Руководитель ЦФЗ' из Справочника "Закупки МИКРОН/ЦФЗ"
       VNIK_TempValueAgree = GetUserDirValuesVNIK(MIKRON_CATALOG_FINANCIAL_COSTS,"Field3","Field1",S_AddFieldText1)
       S_AdditionalUsers = S_AdditionalUsers + VbCrLf + VNIK_TempValueAgree

       S_AddFieldText1 = SIT_GetDocField(S_DocIDParent, "UserFieldText5", Conn)
       'Берем поле №3 'Руководитель ЦФO' из Справочника "Закупки МИКРОН/ЦФO"
       VNIK_TempValueAgree = GetUserDirValuesVNIK(MIKRON_CATALOG_RESPONCIBILITY,"Field3","Field1",S_AddFieldText1)
       S_AdditionalUsers = S_AdditionalUsers + VbCrLf + VNIK_TempValueAgree
       S_AddFieldText1 = ""

       VAR_ChangeDocGenerateButton = ""
        
       S_DocID = " "
       If S_DocID_Set <> "" Then
          S_DocID_Set = S_DocID
       End If

       S_DocIDadd_Set = " "
       VAR_DocFieldsNotToShow = "DocIDAdd"
      
       S_Author_Set = InsertionName(Session("Name"), Session("UserID"))
       S_Department = Session("Department")
       S_Name = S_Name + " Опросный лист к ПЗК"
       S_NameAproval = MIKRON_SecretaryOfPC
       S_AmountDoc = 0
       S_QuantityDoc = 0

       S_UserFieldText1 = ""  'контрагент 2
       S_UserFieldText1_Set = S_UserFieldText1
       S_UserFieldMoney1 = 0
         
       S_UserFieldText2 = ""  'контрагент 3
       S_UserFieldText2_Set = S_UserFieldText2
       S_UserFieldMoney2 = 0

       S_UserFieldText3 = ""  'способ закупки
       S_UserFieldText4 = ""  'обоснование способа закупки
       S_UserFieldText4_Set = S_UserFieldText4      
       S_UserFieldText5 = ""  'условия оплаты
       S_UserFieldText6 = ""  'вопрос, поставленный на голосование

       'Добавим комментарий по обязательным согласующим
       S_ListToReconcile = SIT_RequiredAgrees + vbCrLf + oPayDox.GetExtTableValue("AgreeMIKRON","Category",Session("CurrentClassDoc"),"List")
    End If

    'заполнение нередактируемых полей
    S_AdditionalUsers_Set = iif (S_AdditionalUsers <> "", S_AdditionalUsers, " ")
    S_DocID_Set = S_DocID
    S_DocIDAdd_Set = S_DocIDAdd
    S_Department_Set = S_Department
    S_Author_Set = S_Author
    S_NameAproval_Set = S_NameAproval

    S_ListToReconcile_Set = S_ListToReconcile

    'порядок следования полей
    CurrentDocFieldOrder = "Name,DocID,DocIdParent,Author,Department,UserFieldText6,UserFieldText3,UserFieldText4,UserFieldText5,Currency,QuantityDoc,PartnerName,AmountDoc,UserFieldText1,UserFieldMoney1,UserFieldText2,UserFieldMoney2,UserFieldText7,NameAproval,ListToReconcile,ListToView,Content"

    'поля помечаются звездочкой - обязательно к заполнению
    CurrentDocRequiredFields = "Author,NameAproval,Name,UserFieldText5,UserFieldText6,UserFieldText3,UserFieldText4,Currency,PartnerName,AmountDoc"

    'Не показываем кнопку выбора из справочника
    VAR_DirPictNotToShow = "DocName,AmountDoc,QuantityDoc,UserFieldMoney1,UserFieldMoney2,UserFieldText3,UserFieldText4,UserFieldText5,UserFieldText6,UserFieldText7"

    'Не показываем поля
    VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+",NameControl"
    VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ",AddFieldText1,AddFieldText2,BusinessUnit,ContractType"

    'выбор способа закупки, ссылка на справочник "Закупки Микрон/Способы закупки"
    S_UserFieldText3_Select = GetUserDirValues(MIKRON_CATALOG_PURCHASE_TYPES)
    'Поле "Порядок оплаты", ссылка на справочник "Закупки Микрон/Порядок оплаты"
    S_UserFieldText5_Select = GetUserDirValues(MIKRON_CATALOG_PAYMENT_TYPES)
    'Вопрос, поставленный на голосование
    S_UserFieldText6_Select = MIKRON_CHOISE_PC
    'Поле необходимость доп.затрат на СМР и т.п.
    S_UserFieldText7_Select = SIT_YesNo   
   
'mikron_RL_protocol PC

' *********************************************************************************
' ***                               ПРОТОКОЛЫ                                   ***
' *********************************************************************************
' *** ПРОТОКОЛЫ Комитетов, ЦЗК СИТРОНИКС
'vnik_protocolsCPC
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_PROTOCOLS)) = 1 or InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_PROTOCOLS_CPC)) > 0 Then
  'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
  S_SecurityLevel = 4
  If not IsAdmin() Then
     S_SecurityLevel_Set = S_SecurityLevel
  End If

  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
     S_AdditionalUsers = ""
  End If
  
  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" and InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_PROTOCOLS_CPC)) > 0 Then
     S_AdditionalUsers = """Успенский А. В."" <uspensky>;"
  End If
  
  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" and InStr(UCase(Session("Department")), UCase(SIT_RTI)) = 1 Then
     S_AdditionalUsers = """Боев С. Ф."" <boev_oaorti>;"
    If InStr(UCase(Session("Department")), UCase("Управление ИТ")) > 0 Then
      S_AdditionalUsers = S_AdditionalUsers + VbCrLf + """Иванов В. В."" <vivanon_oaorti>;"
    End If
     DocIDParent_RTI = SIT_GetDocField(S_DocIDParent, "Correspondent", Conn)
     DocIDPar_RTI = SIT_GetDocField(S_DocIDParent, "DocId", Conn)
     if len(DocIDParent_RTI) >0 and InStr(UCase(DocIDPar_RTI),"T_") > 0 Then
        S_AdditionalUsers = S_AdditionalUsers + VbCrLf + DocIDParent_RTI
     End If
  End If
  
  If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_PROTOCOLS_CPC)) > 0 Then
     If S_AdditionalUsers <> "" Then
        S_AdditionalUsers = S_AdditionalUsers + VbCrLf + ReplaceRoleFromDir(SIT_ChiefOfPurchaseDepartment,GetRootDepartment(S_Department))
     Else
        S_AdditionalUsers = S_AdditionalUsers + ReplaceRoleFromDir(SIT_ChiefOfPurchaseDepartment,sDepartmentRoot)
     End If
  End If

  If S_AdditionalUsers <> "" Then
     S_AdditionalUsers_Set = S_AdditionalUsers
  Else
     S_AdditionalUsers_Set = " "
  End If

  S_DocID_Set = S_DocID
  S_DocIDAdd_Set = S_DocIDAdd
  S_Department_Set = S_Department
  S_Author_Set = S_Author

  ' Создание карточки документа
  If UCase(Request("create")) = "Y" Then
     If Request("UpdateDoc") <> "YES" Then
        'SAY 2009-02-20
        VAR_ChangeDocGenerateButton=""
        S_DocID = " "
        If S_DocID_Set<>"" Then
           S_DocID_Set = S_DocID
        End If

        'SAY 2008-11-06
        S_Department_Set = Session("Department")

        'SAY 2008-09-22 для подчиненных документов оставляем весь список согласования родителя
        If Trim(S_DocIDParent)<>"" and InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_PROTOCOLS_CPC)) = 0 then
           S_ListToReconcile = Replace(S_ListToReconcile,"##;","")
        End If

        If Trim(S_DocIDParent)<>"" and InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_PROTOCOLS_CPC)) > 0 then
           S_ListToView = S_Author  
        End If

        S_DocID_Set = " "
        S_DocIDadd_Set = " "
        VAR_DocFieldsNotToShow = "DocIDAdd"
        S_UserFieldDate1 = MyDate(Date)
        S_UserFieldDate2 = MyDate(Date)
      
        S_Author_Set = InsertionName(Session("Name"), Session("UserID"))
        S_Department = Session("Department")
        S_Department_Set = Session("Department")
      
        'Подставим роль Председателя и наименование в зависимости от вида Протокола
        If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_PROTOCOLS_MC_EGRB)) > 0 Then
           S_NameAproval = SIT_ChairmanOfManagingCommitteeOnTheProgramEGRB
           S_Name = "Протокол Управляющего комитета по программе ""Электронное правительство Республики Башкортостан"""
        Else 
           If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_PROTOCOLS_Management_Board)) > 0 Then
              S_NameAproval = SIT_ChairmanOfBoard
              S_Name = "Протокол Правления"
           Else
              If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_PROTOCOLS_IT_Committee)) > 0 Then
                 S_NameAproval = SIT_ChairmanOfCommitteeOnIT
                 S_Name = "Протокол Комитета по ИТ"
              Else
                 If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_PROTOCOLS_Control_And_Auditing_Committee)) > 0 Then
                    S_NameAproval = SIT_ChairmanOfCommitteeOnControlAndAuditing
                    S_Name = "Протокол Контрольно-ревизионного комитета"
                 Else
                    If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_PROTOCOLS_CPC)) > 0 Then
                       S_NameAproval = SIT_ChairmanOfCentralPurchasingCommission
                       S_Name = "Протокол ЦЗК"
                    Else             
                       'Для протокола Встреч этого не нужно
                    End If
                 End If
              End If
           End If
        End If

        'Делаем регистратором разработчика Протокола
        If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_PROTOCOLS_CPC)) > 0 Then
           S_LocationPath = S_Author_Set
        End If
     End If
  End If 'UCase(Request("create")) = "Y"

  S_ListToReconcile = DeleteConstPrefixFromList(S_ListToReconcile)

  CurrentDocFieldOrder = "Name, UserFieldDate2, UserFieldText1, DocIDParent, DocID, Author, Content, UserFieldText2, ListToReconcile, NameAproval, LocationPath, ListToView, UserFieldDate1, DateActivation, SecurityLevel, Department "	
  ' поля помечаются звездочкой - обязательно к заполнению
  CurrentDocRequiredFields = "Name, UserFieldText1, UserFieldDate1, UserFieldText2, NameAproval"

  If InStr(UCase(Session("Department")), UCase(SIT_RTI)) = 1 Then
     S_LocationPath = S_Author_Set
     VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", DateActivation"
     S_UserFieldText3_Select = GetUserDirValues("{4AEC67F9-80C1-427F-904E-489920B71940}")
     CurrentDocRequiredFields = CurrentDocRequiredFields & ", UserFieldText3"
     CurrentDocFieldOrder = "Name, UserFieldText1, UserFieldText3, DocIDParent, DocID, UserFieldDate2, Author, Content, UserFieldText2, ListToReconcile, NameAproval, LocationPath, ListToView, DateActivation, SecurityLevel,Department, UserFieldDate1"	
  End If

  'Делаем поля нередактируемыми
  S_LocationPath_Set = S_LocationPath
  '?
  If InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) = 1 Then
     S_ListToReconcile_Comment = oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")
  Else
    '
  End If
  
  VAR_DirPictNotToShow = VAR_DirPictNotToShow+", DocName"
  'rmanyushin 60298 02.11.2009 Start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", ContractType"
  'rmanyushin 60298 02.11.2009 End
'ph - 20120216 - start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", AddFieldText1, AddFieldText2"
'ph - 20120216 - end

  'Для пользователей СТС везде добавляем поле бизнес единица
  If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
     If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 Then
        If UCase(Request("create")) = "Y" Then
           sCreator = Session("UserID")
        Else
           sCreator = GetUserID(S_Author)
        End If
        BusinessUnitsList = GetUsersBusinessUnits(sCreator)
        If InStr(BusinessUnitsList, VbCrLf) > 0 Then
           S_AddField2 = " "
           S_AddField_Select2 = BusinessUnitsList
           'Поле обязательно для заполнения
           CurrentDocRequiredFields = CurrentDocRequiredFields+", BusinessUnit"
        Else
           S_AddField_Set2 = BusinessUnitsList
        End If
     End If
  Else
     'Поле показывается только для пользователей СТС и администратора (у админа не из СТС оно необязательное, редактируемое)
     If not IsAdmin() Then
        VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", BusinessUnit"
     End If
  End If
'vnik_protocolsCPC

' *********************************************************************************
' ***                               ЗАЯВКА НА ОПЛАТУ                            ***
' *********************************************************************************
' *** 1. УК (СИТРОНИКС)
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_PAYMENT_ORDER)) = 1 Then
  'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
  S_SecurityLevel = 4
  If not IsAdmin() Then
    S_SecurityLevel_Set = S_SecurityLevel
  End If


  If S_AdditionalUsers <> "" Then
    S_AdditionalUsers_Set = S_AdditionalUsers
  Else
    S_AdditionalUsers_Set = " "
  End If

  S_DocID_Set = S_DocID
  S_DocIDAdd_Set = S_DocIDAdd
  S_Department_Set = S_Department
  S_Author_Set = S_Author

  ' Создание карточки документа
  If UCase(Request("create")) = "Y" Then
    If Request("UpdateDoc") <> "YES" Then
      'SAY 2009-02-20
      VAR_ChangeDocGenerateButton=""
      S_DocID = " "
      If S_DocID_Set<>"" Then
        S_DocID_Set = S_DocID
      End If

      'SAY 2008-11-06
      S_Department_Set = Session("Department")

      'SAY 2008-09-22 для подчиненных документов оставляем весь список согласования родителя
      If Trim(S_DocIDParent)<>"" then
        S_ListToReconcile = Replace(S_ListToReconcile,"##;","")
      End If

      S_DocID_Set = " "
            
      S_Author_Set = InsertionName(Session("Name"), Session("UserID"))
      S_Department = Session("Department")
      S_Department_Set = S_Department
      
      'S_NameAproval = SIT_TreasuryController
      S_NameAproval = SIT_Account_manager
      S_NameAproval_Set = S_NameAproval
      S_Name = "Заявка на оплату"
      
      S_Currency = "RUR"
      S_UserFieldText2 = "Банковский перевод"
      S_UserFieldText6 = "Средний"
      S_UserFieldText7 = "Месяц"
      
      'Заполнение на основании договора
      S_DocIDParent_Set = S_DocIDParent
      If Trim(S_DocIDParent)<>"" and InStr(UCase(S_DocIDParent), UCase("CNT-HQ-")) > 0 Then
        S_ListToReconcile = ""
        S_Content = ""
        S_PartnerName_Set = S_Partner
        'не работает
        S_UserFieldText1_Set = S_UserFieldText1
        
        'получаем родителя заявки на оплату УК
        'out S_DocIDParent
        vnik_DocIDParent_Contract_MC = SIT_GetDocField(S_DocIDParent, "DocIDParent", Conn)
        'out vnik_DocIDParent_Contract_MC
        If InStr(UCase(vnik_DocIDParent_Contract_MC), UCase("POHQ-")) > 0 Then
            S_UserFieldMoney1 = SIT_GetDocField(S_DocIDParent, "UserFieldMoney1", Conn) + SIT_GetDocField(S_DocIDParent, "UserFieldMoney2", Conn)
            S_UserFieldMoney2 = "1"
            S_Currency_Set = SIT_GetDocField(vnik_DocIDParent, "Currency", Conn)
            S_UserFieldText5_Set = SIT_GetDocField(vnik_DocIDParent_Contract_MC, "UserFieldText5", Conn) 
            
            'не работает
            S_UserFieldText3_Set = SIT_GetDocField(vnik_DocIDParent_Contract_MC, "UserFieldText3", Conn)
            'out Trim(SIT_GetDocField(vnik_DocIDParent_Contract_MC, "UserFieldText3", Conn))
            S_UserFieldText4_Set = SIT_GetDocField(vnik_DocIDParent_Contract_MC, "UserFieldText4", Conn)
       
        ElseIf InStr(UCase(vnik_DocIDParent_Contract_MC), UCase("PR-CPC-")) > 0 Then
            vnik_DocIDParent_PurchaseOrder = SIT_GetDocField(vnik_DocIDParent_Contract_MC, "DocIDParent", Conn)
            'out vnik_DocIDParent_PurchaseOrder
            S_UserFieldMoney1 = SIT_GetDocField(S_DocIDParent, "UserFieldMoney1", Conn) + SIT_GetDocField(S_DocIDParent, "UserFieldMoney2", Conn)
            S_UserFieldMoney2 = "1"
            S_Currency_Set = SIT_GetDocField(vnik_DocIDParent_PurchaseOrder, "Currency", Conn)
            S_UserFieldText5_Set = SIT_GetDocField(vnik_DocIDParent_PurchaseOrder, "UserFieldText5", Conn)
            
            'не работает
            S_UserFieldText3_Set = SIT_GetDocField(vnik_DocIDParent_PurchaseOrder, "UserFieldText3", Conn)
            'out Trim(SIT_GetDocField(vnik_DocIDParent_PurchaseOrder, "UserFieldText3", Conn))
            S_UserFieldText4_Set = SIT_GetDocField(vnik_DocIDParent_PurchaseOrder, "UserFieldText4", Conn)       
            
        End If
      ElseIf Trim(S_DocIDParent)<>"" and InStr(UCase(S_DocIDParent), UCase("POHQ-")) > 0 Then
      
            S_ListToReconcile = ""
            S_Content = ""
            S_PartnerName_Set = S_Partner
            'не работает
            S_UserFieldText1_Set = S_UserFieldText1
        
            S_UserFieldMoney1 = SIT_GetDocField(S_DocIDParent, "UserFieldMoney1", Conn) + SIT_GetDocField(S_DocIDParent, "UserFieldMoney2", Conn)
            S_UserFieldMoney2 = "1"
            S_Currency_Set = SIT_GetDocField(vnik_DocIDParent, "Currency", Conn)
            S_UserFieldText5_Set = SIT_GetDocField(vnik_DocIDParent_Contract_MC, "UserFieldText5", Conn) 
            
            'не работает
            S_UserFieldText3_Set = SIT_GetDocField(vnik_DocIDParent_Contract_MC, "UserFieldText3", Conn)
            'out Trim(SIT_GetDocField(vnik_DocIDParent_Contract_MC, "UserFieldText3", Conn))
            S_UserFieldText4_Set = SIT_GetDocField(vnik_DocIDParent_Contract_MC, "UserFieldText4", Conn)
      End If
    End If
  End If 'UCase(Request("create")) = "Y"   

    'kkoshkin 04042012 start
    If InStr(UCase(Session("Department")),UCase("СИТРОНИКС*SITRONICS*/Комплекс маркетинга и развития бизнеса*Complex marketing and business development*/"))=1 and UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" then
      S_ListToReconcile = S_ListToReconcile + " ""Кутуков К. В."" <kutukov>;"        
    end If
    'kkoshkin 04042012 end

  S_ListToReconcile = DeleteConstPrefixFromList(S_ListToReconcile)

  'Добавим комментарий по обязательным согласующим
  If InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) = 1 Then
    S_ListToReconcile_Comment = oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")
  ElseIf InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 then
    ' и т.д.
  End If

  CurrentDocFieldOrder = "Name, DocID, DocIDParent, Author, Department, PartnerName, Description, Currency, UserFieldMoney2, UserFieldMoney1, UserFieldText2, UserFieldDate2, UserFieldText4, UserFieldText3, UserFieldText1, UserFieldText7, UserFieldText6, NameAproval, ListToReconcile, ListToView, UserFieldText5"
  ' поля помечаются звездочкой - обязательно к заполнению
  CurrentDocRequiredFields = "Description, Author, PartnerName, UserFieldMoney1, Currency, UserFieldText2, NameAproval, UserFieldText4, UserFieldText3, UserFieldText6, UserFieldText7, Name"
  'Основание платежа ,Срок оплаты ,Дата и время создания ,Инициатор,Контрагент ,Сумма с НДС ,Валюта ,Форма расчета ,Утверждающий ,ЦФЗ ,Приоритет ,Планируемый срок платежа
   
  'Не показываем кнопку выбора из справочника
  If Trim(S_DocIDParent)<>"" Then
  VAR_DirPictNotToShow = "UserFieldText1, UserFieldMoney1, UserFieldMoney2"
  Else
    VAR_DirPictNotToShow = "UserFieldMoney1, UserFieldMoney2"
  End If
  '?
  If InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) = 1 Then  
    S_ListToReconcile_Comment = oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")
  Else
    '
  End If

  VAR_DirPictNotToShow = VAR_DirPictNotToShow+", DocName"
  'rmanyushin 60298 02.11.2009 Start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", ContractType"
  'rmanyushin 60298 02.11.2009 End
'ph - 20120216 - start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", AddFieldText1"
'ph - 20120216 - end

  'vnik_payment_order
  S_UserFieldText2_Select = GetUserDirValues("{6D6236CD-DA05-4B52-87F1-6C657F2544EE}")
  S_UserFieldText3_Select = GetUserDirValues("{15EB5243-22D8-425D-B31A-9CBA4396FCFC}")
  S_UserFieldText4_Select = GetUserDirValues("{E2D63CEF-6DC4-44F3-B4BF-BD4B0D9438D0}") 
  'S_UserFieldText5_Select = GetUserDirValues("{507B2058-B7C4-40B2-8EC4-D75A8E4CE28D}")
  S_UserFieldText6_Select = GetUserDirValues("{3ECADCD6-0985-4659-8774-C8C9D77EE381}")
  S_UserFieldText7_Select = GetUserDirValues("{C62278E8-50EB-41AF-89C4-90717DF68414}")
  S_Currency_Select       = GetUserDirValues("{B53443BE-B30E-4C52-98D7-5ED0BEE67080}")   
  'vnik_payment_order

  'Для пользователей СТС везде добавляем поле бизнес единица
  If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
    If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 Then
      If UCase(Request("create")) = "Y" Then
        sCreator = Session("UserID")
      Else
        sCreator = GetUserID(S_Author)
      End If
      BusinessUnitsList = GetUsersBusinessUnits(sCreator)
      If InStr(BusinessUnitsList, VbCrLf) > 0 Then
        S_AddField2 = " "
        S_AddField_Select2 = BusinessUnitsList
        'Поле обязательно для заполнения
        CurrentDocRequiredFields = CurrentDocRequiredFields+", BusinessUnit"
      Else
        S_AddField_Set2 = BusinessUnitsList
      End If
    End If
  Else
    'Поле показывается только для пользователей СТС и администратора (у админа не из СТС оно необязательное, редактируемое)
    If not IsAdmin() Then
      VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", BusinessUnit"
    End If
  End If

' *********************************************************************************
' ***                               ЗАЯВКА НА ОПЛАТУ                            ***
' *********************************************************************************
' *** 2. РТИ
'rti_payment_order
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(RTI_PAYMENT_ORDER)) = 1 Then
  'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
  S_SecurityLevel = 4
  If not IsAdmin() Then
     S_SecurityLevel_Set = S_SecurityLevel
  End If

  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
     S_AdditionalUsers = """Боев С. Ф."" <boev_oaorti>;"+ VbCrLf + """Иванов В. В."" <vivanon_oaorti>;"
     S_AdditionalUsers_Set = S_AdditionalUsers
  End If 

  If S_AdditionalUsers <> "" Then
     S_AdditionalUsers_Set = S_AdditionalUsers
  Else
     S_AdditionalUsers_Set = " "
  End If    

  S_DocID_Set = S_DocID
  S_DocIDAdd_Set = S_DocIDAdd
  S_Department_Set = S_Department
  S_Author_Set = S_Author

  if S_DocIDParent<>"" Then
    S_DocIDParent_Set = S_DocIDParent
  Else 
    S_DocIDParent_Set = " "
  End If

  ' Создание карточки документа
  If UCase(Request("create")) = "Y" Then
     If Request("UpdateDoc") <> "YES" Then
        'SAY 2009-02-20
        VAR_ChangeDocGenerateButton=""
        S_DocID = " "
        If S_DocID_Set<>"" Then
           S_DocID_Set = S_DocID
        End If

        'SAY 2008-11-06
        S_Department_Set = Session("Department")

        'SAY 2008-09-22 для подчиненных документов оставляем весь список согласования родителя
        If Trim(S_DocIDParent)<>"" then
           S_ListToReconcile = Replace(S_ListToReconcile,"##;","")
        End If

        S_DocID_Set = " "
            
      S_Author_Set = InsertionName(Session("Name"), Session("UserID"))
      S_Department = Session("Department")
      S_Department_Set = S_Department
      
      S_NameAproval = RTI_PaymentHead
      S_NameAproval_Set = S_NameAproval
      S_Name = "Заявка на оплату"
      S_UserFieldText5_Select = SIT_YesNo
     
        'Заполнение на основании заявки на закупку
        S_DocIDParent_Set = S_DocIDParent
        'kkoshkin 20151001
        'если создается подчиненный документ - поле "документ-основание" оставляем пустым
        'если создается копия заявки за оплату, то поле "документ-основание" автозаполняется из исходного документа
        If not (Request("ClassDocDependant") = "" and Request("empty") = "") Then
            S_UserFieldText3 = ""
        End If
        'kkoshkin 20151001
        S_ListToReconcile = ""
        S_ListToView = SIT_GetDocField(S_DocIDParent, "Author", Conn)
     End If
  End If 'UCase(Request("create")) = "Y"   


  S_ListToReconcile = DeleteConstPrefixFromList(S_ListToReconcile)

  'Добавим комментарий по обязательным согласующим
  'S_ListToReconcile_Comment = SIT_RequiredAgrees + vbCrLf + oPayDox.GetExtTableValue("AgreeRTI"+Request("l"),"Category",Session("CurrentClassDoc"),"List") + VbCrLf + RTI_ChiefOfPurchaseDepartment + " " +RTI_HeadKFIE
  S_ListToReconcile_Comment = SIT_RequiredAgrees + vbCrLf + oPayDox.GetExtTableValue("AgreeRTI"+Request("l"),"Category",Session("CurrentClassDoc"),"List") + " " +RTI_HeadKFIE
  CurrentDocFieldOrder = "Name, DocID, DocIDParent, Author, Department, PartnerName, UserFieldText1, UserFieldText2, UserFieldText4, UserFieldText6, UserFieldText3, UserFieldMoney1, UserFieldText5, NameAproval, ListToReconcile, ListToView"
  ' поля помечаются звездочкой - обязательно к заполнению
  CurrentDocRequiredFields = "Author, PartnerName, UserFieldMoney1, UserFieldText2, NameAproval, UserFieldText4, UserFieldText3, UserFieldText5,UserFieldText6, Name"
 
   
  'Не показываем кнопку выбора из справочника
  VAR_DirPictNotToShow = "UserFieldMoney1, UserFieldMoney2, UserFieldText1,UserFieldText3,UserFieldText6"

  VAR_DirPictNotToShow = VAR_DirPictNotToShow+", DocName"
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", NameControl"
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", AddFieldText1, BusinessUnit, ContractType"
  
   'UserFieldText2 - центр затрат
  S_UserFieldText2_Select = GetUserDirValues("{6A8607D5-88A1-4706-87D4-B37D633B2671}")
  S_UserFieldText4_Select = GetUserDirValues("{365C2A1C-D404-47AF-AC76-9421A36E8E6A}")
  S_UserFieldText1_Select = GetUserDirValues("{92C89199-89B9-4DE1-9D8A-CBE0C7A20081}")
  S_UserFieldText6_Select = SIT_YesNo

'rti_payment_order

' *********************************************************************************
' ***                               ЗАЯВКА НА ЗАКУПКУ                           ***
' *********************************************************************************
' *** РТИ
'rti_purchase_order
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(RTI_PURCHASE_ORDER)) > 0 Then

  'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
  S_SecurityLevel = 4
  If not IsAdmin() Then
     S_SecurityLevel_Set = S_SecurityLevel
  End If
    
  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
    S_AdditionalUsers = """Боев С. Ф."" <boev_oaorti>;" + VbCrLf + """Искоростенский С. П."" <isp_oaorti>;" + VbCrLf + ReplaceRoleFromDir(RTI_SecretaryOfCPC,SIT_RTI)
    If InStr(UCase(Session("Department")), UCase("Управление ИТ")) > 0 Then
      S_AdditionalUsers = S_AdditionalUsers + VbCrLf + """Иванов В. В."" <vivanon_oaorti>;"
    End If
    S_AdditionalUsers_Set = S_AdditionalUsers
  End If 

  If S_AdditionalUsers <> "" Then
     S_AdditionalUsers_Set = S_AdditionalUsers
  Else
     S_AdditionalUsers_Set = " "
  End If

  S_DocID_Set = S_DocID
  S_DocIDAdd_Set = S_DocIDAdd
  S_Department_Set = S_Department
  S_Author_Set = S_Author

  if S_DocIDParent<>"" Then
    S_DocIDParent_Set = S_DocIDParent
  Else 
    S_DocIDParent_Set = " "
  End If

  ' Создание карточки документа
  If UCase(Request("create")) = "Y" Then
     If Request("UpdateDoc") <> "YES" Then
        'SAY 2009-02-20
        VAR_ChangeDocGenerateButton=""
        S_DocID = " "
        If S_DocID_Set<>"" Then
           S_DocID_Set = S_DocID
        End If

        'SAY 2008-11-06
        S_Department_Set = Session("Department")

        S_DocID_Set = " "
            
        S_Author_Set = InsertionName(Session("Name"), Session("UserID"))
        S_Department = Session("Department")
        S_Department_Set = S_Department
      
      S_NameAproval = RTI_HeadKFIE
      S_NameAproval_Set = S_NameAproval
      S_Name = "Заявка на закупку"

     
      'сделаем поле наименование проекта нередактируемым вручную
      'S_UserFieldText2_Set = ""
      
        'S_UserFieldText2 = "Банковский перевод"
        'S_UserFieldText6 = "Средний"
        'S_UserFieldText7 = "Месяц"
     End If
  End If 'UCase(Request("create")) = "Y"

  'сделаем поле наименование проекта нередактируемым вручную
  'S_UserFieldText2_Set = S_UserFieldText2
  
  
  S_ListToReconcile = DeleteConstPrefixFromList(S_ListToReconcile)
  CurrentDocFieldOrder = "Name, DocID, UserFieldText3, Author, UserFieldMoney1, Department, PartnerName, UserFieldText1, UserFieldText2, UserFieldText4, NameAproval, ListToReconcile, ListToView"
  
  ' поля помечаются звездочкой - обязательно к заполнению
  CurrentDocRequiredFields = "Author, UserFieldMoney1, NameAproval, UserFieldText3, Name, UserFieldText2"
     
  'Не показываем кнопку выбора из справочника
  VAR_DirPictNotToShow = "UserFieldMoney1, UserFieldMoney2, UserFieldText1,UserFieldText3"
  '?
    
  'Добавим комментарий по обязательным согласующим
  S_ListToReconcile_Comment = SIT_RequiredAgrees + vbCrLf + oPayDox.GetExtTableValue("AgreeRTI"+Request("l"),"Category",Session("CurrentClassDoc"),"List") + " "  + RTI_BudgetController' + VbCrLf + RTI_ChiefOfPurchaseDepartment 

  VAR_DirPictNotToShow = VAR_DirPictNotToShow+", DocName"
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", NameControl"
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", AddFieldText1, BusinessUnit, ContractType"

 'UserFieldText2 - центр затрат
  S_UserFieldText2_Select = GetUserDirValues("{6A8607D5-88A1-4706-87D4-B37D633B2671}")
  S_UserFieldText4_Select = GetUserDirValues("{365C2A1C-D404-47AF-AC76-9421A36E8E6A}")
  S_UserFieldText1_Select = GetUserDirValues("{92C89199-89B9-4DE1-9D8A-CBE0C7A20081}")
  'If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
  'S_UserFieldText5 = "Прочее"
  'end If
  'test server
  'S_UserFieldText5_Select = GetUserDirValues("{3DA4F9EE-4732-4395-84B1-2715B2AA62E7}")
  'prod server
  'S_UserFieldText5_Select = GetUserDirValues("{84E1A1BB-0CBB-4258-9017-9D92EEEE2522}")
  
  
  
'rti_purchase_order

' ********************************* БСАП
' *** РТИ
'rti_bsap
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(RTI_BSAP)) > 0 Then

  'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
  S_SecurityLevel = 4
  If not IsAdmin() Then
    S_SecurityLevel_Set = S_SecurityLevel
  End If

  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
    S_AdditionalUsers = """Боев С. Ф."" <boev_oaorti>;"
    If InStr(UCase(Session("Department")), UCase("Управление ИТ")) > 0 Then
      S_AdditionalUsers = S_AdditionalUsers + VbCrLf + """Иванов В. В."" <vivanon_oaorti>;"
    End If
    S_AdditionalUsers_Set = S_AdditionalUsers
  End If 

  If S_AdditionalUsers <> "" Then
    S_AdditionalUsers_Set = S_AdditionalUsers
  Else
    S_AdditionalUsers_Set = " "
  End If

  S_DocID_Set = S_DocID
  S_DocIDAdd_Set = S_DocIDAdd
  S_Department_Set = S_Department
  S_Author_Set = S_Author

  if S_DocIDParent<>"" Then
    S_DocIDParent_Set = S_DocIDParent
  Else 
    S_DocIDParent_Set = " "
  End If

  ' Создание карточки документа
  If UCase(Request("create")) = "Y" Then
    If Request("UpdateDoc") <> "YES" Then
      'SAY 2009-02-20
      VAR_ChangeDocGenerateButton=""
      S_DocID = " "
      If S_DocID_Set<>"" Then
        S_DocID_Set = S_DocID
      End If

      'SAY 2008-11-06
      S_Department_Set = Session("Department")

      S_DocID_Set = " "
            
      S_Author_Set = InsertionName(Session("Name"), Session("UserID"))
      S_Department = Session("Department")
      S_Department_Set = S_Department
      
      S_NameAproval = RTI_DirectorOfSecurity
      S_NameAproval_Set = S_NameAproval
      S_Name = "БСАП"
      
      S_UserFieldDate1 = MyDate(Date)
      S_UserFieldDate1_Set = S_UserFieldDate1

                      
    End If
  End If 'UCase(Request("create")) = "Y"
  
  S_ListToReconcile = DeleteConstPrefixFromList(S_ListToReconcile)

  CurrentDocFieldOrder = "Name, DocID, DocIdParent, Author, Content, Department, NameAproval, ListToReconcile, ListToView"
  ' поля помечаются звездочкой - обязательно к заполнению
  CurrentDocRequiredFields = "Author, NameAproval, Name"
     
  'Не показываем кнопку выбора из справочника
  VAR_DirPictNotToShow = "UserFieldMoney1, UserFieldMoney2, UserFieldText1,UserFieldText3,UserFieldText4"
  '?
  
  'Добавим комментарий по обязательным согласующим
  S_ListToReconcile_Comment = SIT_RequiredAgrees + vbCrLf + oPayDox.GetExtTableValue("AgreeRTI"+Request("l"),"Category",Session("CurrentClassDoc"),"List") + " " +RTI_HeadOfPurchaseCenter

  VAR_DirPictNotToShow = VAR_DirPictNotToShow+", DocName"
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", NameControl"
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", AddFieldText1, BusinessUnit, ContractType" 
'rti_bsap

' <!--#INCLUDE FILE="amwMIKRON_UserChangeDocSetValues.asp" -->
' AMW - MIKRON

' *********************************************************************************
' ***                               ЗАКУПКИ МИКРОН                              ***
' *** 1. МИКРОН ЗАЯВКА НА ЗАКУПКУ                                               ***
' *********************************************************************************
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_PURCHASE_ORDER)) > 0 Then

  'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
  S_SecurityLevel = 4
  If not IsAdmin() Then
     S_SecurityLevel_Set = S_SecurityLevel
  End If

  ' Создание карточки документа
  If UCase(Request("create")) = "Y" Then
     If Request("UpdateDoc") <> "YES" Then
        VAR_ChangeDocGenerateButton=""
        S_DocID = " "
        S_AdditionalUsers = ReplaceRoleFromDir(MIKRON_CFO,SIT_MIKRON) + VbCrLf + _
                            ReplaceRoleFromDir(MIKRON_SecretaryOfPC,SIT_MIKRON) + VbCrLf + MIKRON_Auditor

        S_Author = InsertionName(Session("Name"), Session("UserID"))      
        S_Department = Session("Department")
        S_NameAproval = MIKRON_HeadKFIE
        S_Name = ""

        S_UserFieldText1 = ""    'готовим поле Проект
        S_UserFieldText2 = ""    'готовим поле наименование ЦФЗ
        S_UserFieldText5 = ""    'готовим поле наименование ЦФO
        S_UserFieldText4 = ""    'готовим поле статья затрат        
        S_UserFieldText6 = ""    'готовим поле источник финансирования
        'готовим список согласующих
        S_ListToReconcile = SIT_RequiredAgrees + vbCrLf + _
                oPayDox.GetExtTableValue("AgreeMIKRON","Category",Session("CurrentClassDoc"),"List") + _
                SIT_AdditionalAgreesDelimeter
     End If
  End If 'UCase(Request("create")) = "Y"

  S_DocID_Set = S_DocID
  S_Department_Set = S_Department
  S_Author_Set = S_Author
  S_NameAproval_Set = S_NameAproval
  S_ListToReconcile_Set = S_ListToReconcile
  S_AdditionalUsers_Set = iif (S_AdditionalUsers <> "", S_AdditionalUsers, " " )
  
  CurrentDocFieldOrder = "Name,DocID,UserFieldText3,Author,UserFieldMoney1,Department,PartnerName," + _
                         "UserFieldText1,UserFieldText2,UserFieldText5,UserFieldText4,UserFieldText6," + _
                         "NameAproval,ListToReconcile"

  'поля помечаются звездочкой - обязательно к заполнению
  CurrentDocRequiredFields = "Author,UserFieldMoney1,NameAproval,Name,UserFieldText2,UserFieldText3," + _
                             "UserFieldText4,UserFieldText5"

  'Не показываем кнопку выбора из справочника
  VAR_DirPictNotToShow = "DocName,UserFieldMoney1,UserFieldMoney2,UserFieldText1,UserFieldText3"

  'Скрываем от редактирования список имеющих доступ к документу
  bListToView = ""

  'Не показываем поля
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow + ",NameControl,AddFieldText1,BusinessUnit,ContractType"

  '"Проект"
  S_UserFieldText1_Select = GetUserDirValues(MIKRON_CATALOG_PROJECT_LIST)
  '"ЦФЗ МИКРОН"
  S_UserFieldText2_Select = GetUserDirValues(MIKRON_CATALOG_FINANCIAL_COSTS)
  '"ЦФO МИКРОН"
  S_UserFieldText5_Select = GetUserDirValues(MIKRON_CATALOG_RESPONCIBILITY)
  '"Источник финансирования"
  S_UserFieldText6_Select = GetUserDirValues(MIKRON_CATALOG_SRC_FINANCING)

' *********************************************************************************
' ***                               ЗАКУПКИ МИКРОН                              ***
' *** 2. МИКРОН БСАП                                                            ***
' *********************************************************************************
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_BSAP)) > 0 Then
   'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
   S_SecurityLevel = 4
   If not IsAdmin() Then
      S_SecurityLevel_Set = S_SecurityLevel
   End If

   'Дополнительные пользователи: Финансовый директор и Секретарь ЗК
   If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
      S_AdditionalUsers = ReplaceRoleFromDir(MIKRON_CFO,SIT_MIKRON) + VbCrLf + _
                          ReplaceRoleFromDir(MIKRON_SecretaryOfPC,SIT_MIKRON) + VbCrLf + MIKRON_Auditor
'amw 28.09.2013 start      
      'Дополнительные пользователи: Руководители ЦФЗ и ЦФО (вычисляем из Заявки на закупку)
      S_AddFieldText1 = SIT_GetDocField(S_DocIDParent, "UserFieldText2", Conn)
      'Берем поле №3 'Руководитель ЦФЗ' из Справочника "Закупки МИКРОН/ЦФЗ"
      VNIK_TempValueAgree = GetUserDirValuesVNIK(MIKRON_CATALOG_FINANCIAL_COSTS,"Field3","Field1",S_AddFieldText1)
      S_AdditionalUsers = S_AdditionalUsers + VbCrLf + VNIK_TempValueAgree

      S_AddFieldText1 = SIT_GetDocField(S_DocIDParent, "UserFieldText5", Conn)
      'Берем поле №3 'Руководитель ЦФO' из Справочника "Закупки МИКРОН/ЦФO"
      VNIK_TempValueAgree = GetUserDirValuesVNIK(MIKRON_CATALOG_RESPONCIBILITY,"Field3","Field1",S_AddFieldText1)
      S_AdditionalUsers = S_AdditionalUsers + VbCrLf + VNIK_TempValueAgree
      S_AddFieldText1 = ""
'amw 28.09.2013 end
      S_AdditionalUsers_Set = S_AdditionalUsers
   End If

   S_AdditionalUsers_Set = iif (S_AdditionalUsers <> "", S_AdditionalUsers, " " )

   S_DocID_Set = S_DocID
   S_DocIDAdd_Set = S_DocIDAdd
   S_Department_Set = S_Department
   S_Author_Set = S_Author

   ' Создание карточки документа
   If UCase(Request("create")) = "Y" Then
      If Request("UpdateDoc") <> "YES" Then
         'SAY 2009-02-20
         VAR_ChangeDocGenerateButton=""
         S_DocID = " "
         If S_DocID_Set<>"" Then
            S_DocID_Set = S_DocID
         End If
         'SAY 2008-11-06
         
         S_Department_Set = Session("Department")
         S_DocID_Set = " "
         S_Author_Set = InsertionName(Session("Name"), Session("UserID"))
         
         S_Department = Session("Department")
         S_Department_Set = S_Department

         S_NameAproval = MIKRON_HeadKFIE
         S_NameAproval_Set = S_NameAproval

         S_Name = SIT_GetDocField(S_DocIDParent, "Name", Conn) + ", БСАП"
         S_AmountDoc = 0
         S_QuantityDoc = 0

         S_UserFieldDate1 = MyDate(Date)
         S_UserFieldDate1_Set = S_UserFieldDate1

         S_UserFieldText1 = ""  'контрагент 2
         S_UserFieldText1_Set = S_UserFieldText1
         S_UserFieldMoney1 = 0
         
         S_UserFieldText2 = ""  'контрагент 3
         S_UserFieldText2_Set = S_UserFieldText2
         S_UserFieldMoney2 = 0
         
         S_UserFieldText3 = Left(MIKRON_CHOISE_PC, InStr(MIKRON_CHOISE_PC,",")-1)   'Выбор предложения
         S_AddField2 = ""  'способ закупки
         
         S_UserFieldText4 = ""  'обоснование способа закупки
         S_UserFieldText4_Set = S_UserFieldText4         
         S_UserFieldText5 = ""  'Порядок оплаты      
         S_UserFieldText6 = ""  'меры по снижению аванса
         S_UserFieldText6_Set = S_UserFieldText6
         S_UserFieldText7 = ""  'объем незакрытых авансов
         S_UserFieldText7_Set = S_UserFieldText7
         S_UserFieldText8 = ""  'срок поставки
         S_UserFieldText8_Set = S_UserFieldText8
         S_Description = MIKRON_TEXT_KP_SELECT
      End If  'Request("UpdateDoc") <> "YES"
   End If 'UCase(Request("create")) = "Y"
  
   CurrentDocFieldOrder = "Name,DocID,DocIdParent,Author,Department,UserFieldText3,Description," + _
                          "BusinessUnit,UserFieldText4,QuantityDoc,Currency,PartnerName,AmountDoc," + _
                          "UserFieldText1,UserFieldMoney1,UserFieldText2,UserFieldMoney2,NameAproval," + _
                          "ListToView"

   'поля помечаются звездочкой - обязательно к заполнению
   CurrentDocRequiredFields = "Author,NameAproval,Name,BusinessUnit,UserFieldText4,QuantityDoc," + _
                              "Currency,PartnerName,AmountDoc,UserFieldText5"

'amw 10/09/2013 Согласующих НЕТ!!!

   'Не показываем кнопку выбора из справочника
   VAR_DirPictNotToShow = "DocName,AmountDoc,QuantityDoc,UserFieldMoney1,UserFieldMoney2," + _
                          "UserFieldText4,UserFieldText5,UserFieldText6,UserFieldText7,UserFieldText8"

   'Не показываем поля
   VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow + ",NameControl,AddFieldText1,ContractType"

   'выбор способа закупки, ссылка на справочник "Закупки Микрон/Способы закупки"
   S_AddField_Select2 = GetUserDirValues(MIKRON_CATALOG_PURCHASE_TYPES)
   
   'Поле "Порядок оплаты", ссылка на справочник "Закупки Микрон/Порядок оплаты"
   S_UserFieldText5_Select = GetUserDirValues(MIKRON_CATALOG_PAYMENT_TYPES)
   
   'Вопрос, поставленный на голосование
   S_UserFieldText3_Select = MIKRON_CHOISE_PC

'mikron_BSAP_end

' *********************************************************************************
' ***                               ЗАКУПКИ МИКРОН                              ***
' *** 3. ДОГОВОР МИКРОН                                                         ***
' *********************************************************************************
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_CONTRACT)) = 1 or _
       InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_S_CONTRACT)) = 1 Then

   S_DocID_Set = S_DocID
   S_DocIDAdd_Set = S_DocIDAdd
   S_Department_Set = S_Department
   S_Author_Set = S_Author

   VNIK_TempValueAgree = S_ListToReconcile
   S_ListToReconcile = "" 'список согласующих
   If InStr(VNIK_TempValueAgree,"##") <= 0 Then
      VNIK_TempValueAgree = ""
   End If

   ' Создание карточки документа
   If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
      VAR_ChangeDocGenerateButton = ""
      S_DocID = " "
      If S_DocID_Set <> "" Then
         S_DocID_Set = S_DocID
      End If
'14-02-14
'Создаём договор доходный или безвозмездный
      If S_DocIDParent = "" and _
         InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_S_CONTRACT)) = 1 Then
         S_AdditionalUsers = ReplaceRoleFromDir(SIT_HeadOfDocControl,SIT_MIKRON)
         ' Доходный,Расходный,Безвозмездный
         S_UserFieldText3_Select = STS_ContractPaymentDirection_In & "," & STS_ContractPaymentDirection_Free
         S_DocIDParent_Set = " "
'Создаём договор для закупки
      Else
         'Дополнительные пользователи: канцелярия и ОВКиА
         S_AdditionalUsers = ReplaceRoleFromDir(SIT_HeadOfDocControl,SIT_MIKRON) + VbCrLf + _
                             ReplaceRoleFromDir(MIKRON_SecretaryOfPC,SIT_MIKRON) + VbCrLf + MIKRON_Auditor
         S_DocIDParent_Set = S_DocIDParent

         ' Доходный,Расходный,Безвозмездный
         S_UserFieldText3_Set = STS_ContractPaymentDirection_Out
         ' Контрагента берем из родительского документа.
         S_Partner = SIT_GetDocField(S_DocIDParent, "PartnerName", Conn)
         ' Cписок получателей обнуляем
         S_ListToView = ""
         ' Если родительский документ непустой, то лезет всякая дрянь из поля AddField2 (блокируем её)
         S_AddField2 = MIKRON_BU_1
      End If

      S_DocID_Set = " "
      S_Author_Set = InsertionName(Session("Name"), Session("UserID"))

      S_Department = Session("Department")
      S_Department_Set = S_Department

      S_NameAproval = MIKRON_GenDirector
      S_NameAproval_Set = S_NameAproval

      S_Name = ""            '"Вид обязательств"
      S_DocIDIncoming = ""   '№договора (внешний)

      S_UserFieldText1 = ""  'контрагент 2
      S_UserFieldText2 = ""  'название проекта

      S_UserFieldText4 = ""  'Статья бюджета
      S_UserFieldText5 = ""  'тип сделки
   End If 'UCase(Request("create")) = "Y"

   S_AdditionalUsers_Set = iif (S_AdditionalUsers <> "", S_AdditionalUsers, "" )

   CurrentDocFieldOrder = "DocID,DocIDParent,Author,NameResponsible,Department,DocIDIncoming," + _
          "BusinessUnit,PartnerName,UserFieldText1,Description,Name,UserFieldText2," + _
          "UserFieldText3,AmountDoc,Currency,UserFieldText5,UserFieldDate4,UserFieldDate6," + _
          "ListToReconcile,NameAproval,SecurityLevel,UserFieldDate5,ContractType,UserFieldText4"

   'поля помечаются звездочкой - обязательно к заполнению
   CurrentDocRequiredFields = "Description,AmountDoc,Currency,PartnerName,UserFieldText2,UserFieldText3," + _
                              "UserFieldText4,UserFieldText5,ContractType,BusinessUnit,NameResponsible,AddFieldText2"

   'Убрать иконку справочников у полей "№ внешний", "Проект", "Предмет договора"
   VAR_DirPictNotToShow = "DocIDIncoming,UserFieldText2,Description"
   VAR_DocFieldsNotToShow = ""

   'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
   S_SecurityLevel = 4
   If not IsAdmin() Then
'amw 17/07/2014 - start
      'При редактировании документа менять значение поля PartnerName можно только Admin 
      If Request("create") <> "y" Then
         'Контрагента не даем редактировать - берется из родительского документа
         S_PartnerName_Set = S_PartnerName       'контрагент
      End If
'amw 17/07/2014 - end
      'Всем запрещено, переводим на правила
''      S_NameAproval_Set = S_NameAproval
      S_SecurityLevel_Set = S_SecurityLevel
      'Добавим комментарий по обязательным согласующим
      If InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_S_CONTRACT)) = 1 Then
         S_ListToReconcile_Comment = SIT_PreliminaryAgrees + _
                   oPayDox.GetExtTableValue("AgreeMIKRON","Name",SIT_PreliminaryAgrees,"List") + _
                   vbCrLf + SIT_RequiredAgrees + _ 
                   oPayDox.GetExtTableValue("AgreeMIKRON","Name",MIKRON_SalesAgrees,"List") + _
                   oPayDox.GetExtTableValue("AgreeMIKRON","Name",SIT_RequiredAgrees,"List") + _
                   vbCrLf + SIT_AdditionalAgrees
      Else
         S_ListToReconcile_Comment = SIT_PreliminaryAgrees + _
                   oPayDox.GetExtTableValue("AgreeMIKRON","Name",SIT_PreliminaryAgrees,"List") + _
                   vbCrLf + SIT_RequiredAgrees + _
                   oPayDox.GetExtTableValue("AgreeMIKRON","Name",SIT_RequiredAgrees,"List") + _
                   vbCrLf + SIT_AdditionalAgrees
      End If
'amw 19/06/2014      S_ListToReconcile = AdditionalAgreeFromList(VNIK_TempValueAgree)
      If VNIK_TempValueAgree <> "" Then
         S_ListToReconcile = AdditionalAgreeFromList(VNIK_TempValueAgree)
      End If
'amw 19/06/2014
   Else
      S_ListToReconcile = VNIK_TempValueAgree
   End If

   'Поле "Вид обязательств", ссылка на справочник "Виды договоров"
   S_Name_Select = GetUserDirValues(MIKRON_CATALOG_CONTRACT_TYPES)
   'Поле "Проект", справочник "Закупки Микрон/Проекты"
   S_UserFieldText2_Select = GetUserDirValues(MIKRON_CATALOG_PROJECT_LIST)
   'Поле "Порядок оплаты(условия платежа)", справочник "Закупки Микрон/Порядок оплаты"
   S_UserFieldText5_Select = GetUserDirValues(MIKRON_CATALOG_PAYMENT_TYPES)
   'Поле "Cтатья бюджета", ссылка на справочник "Закупки Микрон/Статьи бюджета"
   S_UserFieldText4_Set = GetUserDirValues(MIKRON_CATALOG_EXPENDITURE)
  
   'Поле Сторона договора
   S_AddField_Select2 = MIKRON_BUs
   'Поле Договор внутри группы МЭ
   S_AddField_Select3 = SIT_YesNo
   'Поле Зависимость сторон договора
   S_AddField_Select4 = SIT_YesNo
   'выбор статьи бюджета, ссылка на справочник "Закупки Микрон/Статьи бюджета"
    S_AddField_Select5 = MIKRON_DEAL_TYPES
'mikron_contract_end

' *********************************************************************************
' ***                               ЗАКУПКИ МИКРОН                              ***
' *** 3.1 Доп. соглашение к договору МИКРОН                                     ***
' *********************************************************************************
'mikron_additional_agreement
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_ADD_CONTRACT)) = 1 Then

   S_DocID_Set = S_DocID
   S_DocIDAdd_Set = S_DocIDAdd
   S_Department_Set = S_Department
   S_Author_Set = S_Author

   VNIK_TempValueAgree = S_ListToReconcile
   S_ListToReconcile = "" 'список согласующих
   If InStr(VNIK_TempValueAgree,"##") <= 0 Then
      VNIK_TempValueAgree = ""
   End If

   S_DocIDPrevious = Request("DocIDPrevious")
   S_DocIDPrevious_Set = S_DocIDPrevious

   If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
      VAR_ChangeDocGenerateButton = ""

      'Дополнительные пользователи: канцелярия, ОВКиА, секретарь ЗК
      S_AdditionalUsers = ReplaceRoleFromDir(SIT_HeadOfDocControl,SIT_MIKRON) + VbCrLf + _
                          ReplaceRoleFromDir(MIKRON_SecretaryOfPC,SIT_MIKRON) + VbCrLf + MIKRON_Auditor
      S_DocID = " "
      S_DocID_Set = S_DocID

      If Trim(S_DocIDParent) <> "" Then
         If InStr(UCase(SIT_GetDocField(S_DocIDPrevious,"ClassDoc",Conn)),UCase(MIKRON_RL_MEMO)) = 1 Then
            S_Currency = SIT_GetDocField(S_DocIDPrevious, "Currency", Conn)
            S_AmountDoc = SIT_GetDocField(S_DocIDPrevious, "AmountDoc", Conn)
         
            S_ListToReconcile = ""
            VNIK_TempValueAgree = ""
            S_Description = ""
         End If
         S_NameAproval = SIT_GetDocField(S_DocIDParent, "NameAproval", Conn)
      Else
         S_DocIDParent = "Договор до даты 01.10.2013"
         S_NameAproval = MIKRON_GenDirector
      End If

      S_Author = InsertionName(Session("Name"), Session("UserID"))
      S_Department = Session("Department")
'      S_NameAproval = MIKRON_GenDirector

      S_Name = ""            '"Вид обязательств"
      S_DocIDIncoming = ""   '№договора (внешний)

      S_UserFieldText1 = ""  'контрагент 2
      S_UserFieldText2 = ""  'Проект
      S_UserFieldText3 = ""  'Доходный\расходный
   End If 'UCase(Request("create")) = "Y"

   S_AdditionalUsers_Set = iif (S_AdditionalUsers <> "", S_AdditionalUsers, "" )

   CurrentDocFieldOrder = "DocID,DocIDParent,DocIDPrevious,Author,NameResponsible,Department,DocIDIncoming," + _
                          "BusinessUnit,PartnerName,UserFieldText1,Name,UserFieldText2,UserFieldText3," + _
                          "Description,AmountDoc,Currency,UserFieldDate4,UserFieldDate6," + _
                          "ListToReconcile,NameAproval,SecurityLevel,UserFieldDate5"

   'поля помечаются звездочкой - обязательно к заполнению
   CurrentDocRequiredFields = "AmountDoc,Currency,PartnerName,UserFieldText2,Name,UserFieldText3," + _
                              "BusinessUnit,NameResponsible"

   'Убрать иконку справочников у полей
   VAR_DirPictNotToShow = "DocIDIncoming,UserFieldText2,Name,Description,AmountDoc"

   'Скрываем ненужные поля
   VAR_DocFieldsNotToShow = "ContractType,AddFieldText1,AddFieldText2"

   'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
   'Всем запрещено, переводим на правила
   S_SecurityLevel = 4
   If not IsAdmin() Then
      S_SecurityLevel_Set = S_SecurityLevel
      S_NameAproval_Set = S_NameAproval
      S_DocIDParent_Set = S_DocIDParent
      S_DocIDPrevious_Set = S_DocIDPrevious
      If S_DocIDPrevious = "" Then
         bDocIDPrevious = ""
      End If

      'Добавим комментарий по обязательным согласующим
      S_ListToReconcile_Comment = SIT_PreliminaryAgrees + oPayDox.GetExtTableValue("AgreeMIKRON","Name",SIT_PreliminaryAgrees,"List") + vbCrLf + _
                                  SIT_RequiredAgrees + oPayDox.GetExtTableValue("AgreeMIKRON","Name",SIT_RequiredAgrees,"List") + vbCrLf + _
                                  SIT_AdditionalAgrees
      If VNIK_TempValueAgree <> "" Then
         S_ListToReconcile = AdditionalAgreeFromList(VNIK_TempValueAgree)
      End If
      'Контрагента не даем редактировать - берется из родительского документа
      S_PartnerName_Set = S_PartnerName       'контрагент
   Else
      S_ListToReconcile_Comment = ""
      S_ListToReconcile = VNIK_TempValueAgree
   End If

   'Поле "Вид обязательств", ссылка на справочник "Виды доп.соглашений"
   S_Name_Select = GetUserDirValues(MIKRON_CATALOG_ADD_AGREE_TYPES)
   'Поле "Проект", справочник "Закупки Микрон/Проекты"
   S_UserFieldText2_Select = GetUserDirValues(MIKRON_CATALOG_PROJECT_LIST)
   'Поле "Порядок оплаты", ссылка на справочник "Закупки Микрон/Порядок оплаты"
   S_UserFieldText4_Select = GetUserDirValues(MIKRON_CATALOG_PAYMENT_TYPES)
   'Поле Сторона договора
   S_AddField_Select2 = MIKRON_BUs
   'Поле "Доходный\расходный"
   S_UserFieldText3_Select = STS_ContractPaymentDirection_Out & "," & STS_ContractPaymentDirection_In & "," & STS_ContractPaymentDirection_Free
'mikron_additional_agreement_end

' *********************************************************************************
' ***                               ЗАЯВКА НА ОПЛАТУ                            ***
' *********************************************************************************
' *** 4. МИКРОН
'mikron_payment_order
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_PAYMENT_ORDER)) = 1 Then
   'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
   S_SecurityLevel = 4
   If not IsAdmin() Then
      S_SecurityLevel_Set = S_SecurityLevel
   End If

   'Дополнительные пользователи: Финансовый директор
   If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
'      S_AdditionalUsers = ReplaceRoleFromDir(MIKRON_CFO,SIT_MIKRON) + VbCrLf + ReplaceRoleFromDir(MIKRON_SecretaryOfPC,SIT_MIKRON)
      S_AdditionalUsers = Request("AdditionalUsers") + VbCrLf + S_NameResponsible
      S_AdditionalUsers_Set = S_AdditionalUsers
   End If

   If S_AdditionalUsers <> "" Then
      S_AdditionalUsers_Set = S_AdditionalUsers
   Else
      S_AdditionalUsers_Set = " "
   End If    

   S_DocID_Set = S_DocID
   S_DocIDAdd_Set = S_DocIDAdd
   S_Department_Set = S_Department
   S_Author_Set = S_Author

   ' Создание карточки документа
   If UCase(Request("create")) = "Y" Then
      If Request("UpdateDoc") <> "YES" Then
        'SAY 2009-02-20
         VAR_ChangeDocGenerateButton=""
         S_DocID = " "
         If S_DocID_Set<>"" Then
            S_DocID_Set = S_DocID
         End If

         'SAY 2008-11-06
         S_Department_Set = Session("Department")

         'SAY 2008-09-22 для подчиненных документов оставляем весь список согласования родителя
         If Trim(S_DocIDParent)<>"" then
            S_ListToReconcile = Replace(S_ListToReconcile,"##;","")
         End If

         S_DocID_Set = " "            
         S_Author_Set = InsertionName(Session("Name"), Session("UserID"))
         S_Department = Session("Department")
         S_Department_Set = S_Department
    
         S_NameAproval = MIKRON_CFO
         S_NameAproval_Set = S_NameAproval
         S_Name = "Заявка на оплату"
         S_UserFieldText5_Select = SIT_YesNo
     
         'Заполнение на основании заявки на закупку
         S_DocIDParent_Set = S_DocIDParent

         S_UserFieldText3 = ""
         S_ListToReconcile = ""
         S_ListToView = SIT_GetDocField(S_DocIDParent, "Author", Conn)
      End If
   End If 'UCase(Request("create")) = "Y"   

   S_ListToReconcile = DeleteConstPrefixFromList(S_ListToReconcile)

   'Добавим комментарий по обязательным согласующим
   S_ListToReconcile_Comment = SIT_RequiredAgrees + vbCrLf + _
                               oPayDox.GetExtTableValue("AgreeMIKRON","Category",Session("CurrentClassDoc"),"List")

   CurrentDocFieldOrder = "Name,DocID,DocIDParent,Author,Department,PartnerName,UserFieldText1," + _
                          "UserFieldText2,UserFieldText4,UserFieldText3,UserFieldMoney1,UserFieldText5," + _
                          "NameAproval, ListToReconcile, ListToView"

   ' поля помечаются звездочкой - обязательно к заполнению
   CurrentDocRequiredFields = "Author,PartnerName,UserFieldMoney1,UserFieldText2,NameAproval," + _
                              "UserFieldText4, UserFieldText3, UserFieldText5, Name"
 
   'Не показываем кнопку выбора из справочника
   VAR_DirPictNotToShow = "DocName,UserFieldMoney1,UserFieldMoney2,UserFieldText1,UserFieldText3"

   'Не показываем поля
   VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow + ",NameControl,AddFieldText1,BusinessUnit,ContractType"
  
   'UserFieldText2 - центр затрат
   S_UserFieldText2_Select = GetUserDirValues("{6A8607D5-88A1-4706-87D4-B37D633B2671}")
   S_UserFieldText4_Select = GetUserDirValues("{365C2A1C-D404-47AF-AC76-9421A36E8E6A}")
'mikron_payment_order

' *********************************************************************************
' ***                               ЗАКУПКИ МИКРОН                              ***
' *** 5. МИКРОН Справка о ценах к ЛС                                            ***
' *********************************************************************************
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_RL_MEMO)) > 0 Then
   'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
   S_SecurityLevel = 4
   S_DocID_Set = S_DocID
   ' Создание карточки документа
   If UCase(Request("create")) = "Y" Then
      If Request("UpdateDoc") <> "YES" Then
         VAR_ChangeDocGenerateButton = ""

         'Дополнительные пользователи: Финансовый директор и Секретарь ЗК
         S_AdditionalUsers = ReplaceRoleFromDir(MIKRON_CFO,SIT_MIKRON) + VbCrLf + _
                             ReplaceRoleFromDir(MIKRON_SecretaryOfPC,SIT_MIKRON)
         S_Author = InsertionName(Session("Name"), Session("UserID"))
         S_Department = Session("Department")
         S_NameAproval = MIKRON_HeadKFIE

         S_DocID = ""
         S_DocID_Set = " "
         S_AmountDoc = 0
         S_QuantityDoc = 0

         S_UserFieldText1 = ""  'контрагент 2
         S_UserFieldText1_Set = S_UserFieldText1
         S_UserFieldMoney1 = 0
         
         S_UserFieldText2 = ""  'контрагент 3
         S_UserFieldText2_Set = S_UserFieldText2
         S_UserFieldMoney2 = 0
         
         S_UserFieldText3 = Left(MIKRON_CHOISE_PC, InStr(MIKRON_CHOISE_PC,",")-1)   'Выбор предложения

         S_UserFieldText4 = ""  'статья расходов
         S_UserFieldText4_Set = S_UserFieldText4
         S_UserFieldText5 = ""  'способ закупки      
         S_UserFieldText6 = ""  'способ закупки
         S_UserFieldText7 = ""  'способ закупки
         S_UserFieldText8 = ""  'ЦФЗ

         'готовим список согласующих
         S_ListToReconcile = SIT_RequiredAgrees + vbCrLf + _
                oPayDox.GetExtTableValue("AgreeMIKRON","Category",Session("CurrentClassDoc"),"List") + _
                SIT_AdditionalAgreesDelimeter
         S_Description = MIKRON_TEXT_KP_SELECT
      End If  'Request("UpdateDoc") <> "YES"
   End If 'UCase(Request("create")) = "Y"
  
   'S_DocID_Set = iif (S_DocID <> "", S_DocID, " " )
   S_DocIDAdd_Set = S_DocIDAdd
   S_Department_Set = S_Department
   S_Author_Set = S_Author

   If not IsAdmin() Then
      S_NameAproval_Set = S_NameAproval
      S_ListToReconcile_Set = S_ListToReconcile
      S_SecurityLevel_Set = S_SecurityLevel
      S_AdditionalUsers_Set = iif (S_AdditionalUsers <> "", S_AdditionalUsers, " " )
   End If

   CurrentDocFieldOrder = "Name,DocID,Author,Department,UserFieldText3,Description,QuantityDoc," + _
                          "UserFieldText4,UserFieldText8," + _
                          "PartnerName,AmountDoc,Currency,DocIDPrevious,UserFieldDate1,UserFieldText5," + _
                          "UserFieldText1,UserFieldMoney1,DocIDIncoming,UserFieldDate2,UserFieldText6," + _
                          "UserFieldText2,UserFieldMoney2,DocIDadd,UserFieldDate3,UserFieldText7," + _
                          "NameAproval,ListToReconcile"

   'поля помечаются звездочкой - обязательно к заполнению
   CurrentDocRequiredFields = "Name,QuantityDoc,Currency,PartnerName,AmountDoc,UserFieldText4,DocIDPrevious,UserFieldText8"

'amw 10/09/2013 Согласующих НЕТ!!!

   'Не показываем кнопку выбора из справочника
   VAR_DirPictNotToShow = "DocName,AmountDoc,QuantityDoc,UserFieldMoney1,UserFieldMoney2,UserFieldText5,UserFieldText6,UserFieldText7,UserFieldText8"

   'Не показываем поля
   VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow + ",NameControl,AddFieldText1,AddFieldText2,ContractType,BusinessUnit"

   'Поле "Порядок оплаты", ссылка на справочник "Закупки Микрон/Порядок оплаты"
   S_UserFieldText5_Select = GetUserDirValues(MIKRON_CATALOG_PAYMENT_TYPES)
   S_UserFieldText6_Select = GetUserDirValues(MIKRON_CATALOG_PAYMENT_TYPES)
   S_UserFieldText7_Select = GetUserDirValues(MIKRON_CATALOG_PAYMENT_TYPES)
   
   'Вопрос, поставленный на голосование
   S_UserFieldText3_Select = MIKRON_CHOISE_PC

   '"ЦФЗ МИКРОН"
   S_UserFieldText8_Select = GetUserDirValues(MIKRON_CATALOG_FINANCIAL_COSTS)   
   
'mikron_RL_MEMO_end

' *********************************************************************************
' ***                               ЗАКУПКИ МИКРОН                              ***
' *** 6. МИКРОН contracts before date 01/10/2013                                ***
' *********************************************************************************
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_OLD_CONTRACT)) > 0 Then
   'Уровень доступа - 1 общедоступный, 4-только указанным в документе.
   S_SecurityLevel = 4
   'S_DocID_Set = S_DocID
   ' Создание карточки документа
   VAR_ChangeDocGenerateButton = ""

   If UCase(Request("create")) = "Y" Then
      If UCase(Request("UpdateDoc")) <> "YES" Then
         'Дополнительные пользователи: указываем явно
         S_AdditionalUsers = ""
         S_Author = InsertionName(Session("Name"), Session("UserID"))
         S_Department = Session("Department")
         S_NameAproval = MIKRON_Legal

         S_DocID = ""
         'S_DocID_Set = " "
         S_UserFieldText1 = ""
         S_UserFieldText2 = ""
         S_UserFieldText4 = ""
         S_UserFieldText8 = "" 

      End If  'Request("UpdateDoc") <> "YES"
   End If 'UCase(Request("create")) = "Y"

   S_DocIDAdd_Set = S_DocIDAdd
   S_Department_Set = S_Department
   S_Author_Set = S_Author

   If not IsAdmin() Then
      S_NameAproval_Set = S_NameAproval
      S_SecurityLevel_Set = S_SecurityLevel
      If UCase(Request("create")) <> "Y" Then
         S_DocID_Set = S_DocID
      End If
   End If

   CurrentDocFieldOrder = "Name,DocID,UserFieldDate5,Author,Department,UserFieldText1,UserFieldText8," + _
                          "Description,UserFieldText3,PartnerName,UserFieldText4,AmountDoc,Currency," + _
                          "UserFieldText5,UserFieldDate4,UserFieldDate6,UserFieldText2,UserFieldText6," + _
                          "UserFieldText7,UserFieldDate1,NameAproval"

   'поля помечаются звездочкой - обязательно к заполнению
   CurrentDocRequiredFields = "Name,DocID,UserFieldDate5,UserFieldText1,Description,UserFieldText3," + _
                          "PartnerName,AmountDoc,Currency,UserFieldText5,UserFieldText2,UserFieldDate4," + _
                          "NameAproval"

   'Не показываем кнопку выбора из справочника
   VAR_DirPictNotToShow = "DocName,UserFieldText8,Description,UserFieldText4,UserFieldText5,AmountDoc," + _
                          "UserFieldText1,UserFieldText2,QuantityDoc,UserFieldText6,UserFieldText7"

   'Не показываем поля
   VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow + ",NameControl,AddFieldText1,AddFieldText2,ContractType,BusinessUnit"

   'Поле "Доходный\расходный"
   S_UserFieldText3_Select = STS_ContractPaymentDirection_Out & "," & STS_ContractPaymentDirection_In & "," & STS_ContractPaymentDirection_Free
   'Поле "Порядок оплаты", ссылка на справочник "Закупки Микрон/Порядок оплаты"
   S_UserFieldText5_Select = GetUserDirValues(MIKRON_CATALOG_PAYMENT_TYPES)
   'Поле "внутри группы МЭ"
   S_UserFieldText6_Select = SIT_YesNo
   'Поле "Зависимость сторон".
   S_UserFieldText7_Select = SIT_YesNo

'mikron_MIKRON_OLD_CONTRACT_end

' *********************************************************************************
' ***                               ЗАКУПКИ МИКРОН                              ***
' *** 7. МИКРОН NDA contracts                                                   ***
' *********************************************************************************
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_NDA_CONTRACT)) > 0 Then
   'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
   S_SecurityLevel = 4

   S_DocID_Set = S_DocID
   S_DocIDAdd_Set = S_DocIDAdd
   S_Department_Set = S_Department
   S_Author_Set = S_Author

   ' Создание карточки документа
   If UCase(Request("create")) = "Y" Then
      If Request("UpdateDoc") <> "YES" Then
         'убираем возможность сгенерировать для Договора номер в СЭД
         VAR_ChangeDocGenerateButton = ""

         'Дополнительные пользователи: канцелярия, ОВКиА и Секретарь ЗК
         S_AdditionalUsers = ReplaceRoleFromDir(SIT_HeadOfDocControl,SIT_MIKRON)
         S_Author = InsertionName(Session("Name"), Session("UserID"))
         S_Department = Session("Department")
         S_NameAproval = MIKRON_GenDirector

         S_Description = ""
         S_DocID = ""
         S_DocID_Set = " "
      End If  'Request("UpdateDoc") <> "YES"
   End If 'UCase(Request("create")) = "Y"
  
   S_Author_Set = S_Author
   'готовим список согласующих
   S_ListToReconcile_Comment = SIT_RequiredAgrees + _
              oPayDox.GetExtTableValue("AgreeMIKRON","Name",SIT_PreliminaryAgrees,"List") + _
              oPayDox.GetExtTableValue("AgreeMIKRON","Category",Session("CurrentClassDoc"),"List") + _
              SIT_AdditionalAgreeDelimiter
   S_ListToReconcile = DeleteConstPrefixFromList(S_ListToReconcile)
   S_Department_Set = S_Department

   If not IsAdmin() Then
      S_SecurityLevel_Set = S_SecurityLevel
      S_AdditionalUsers_Set = iif (S_AdditionalUsers <> "", S_AdditionalUsers, " " )
'3. >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'      S_DocID_Set = iif (S_DocID <> "", S_DocID, " " )
   End If

   CurrentDocFieldOrder = "Name,DocID,Author,Department,UserFieldText1,Description,PartnerName," + _
                          "DocIDIncoming,UserFieldDate4,UserFieldDate5,UserFieldDate6,NameAproval"

   'поля помечаются звездочкой - обязательно к заполнению
   CurrentDocRequiredFields = "Name,PartnerName,UserFieldText1,UserFieldDate4,UserFieldDate5,UserFieldDate6"

   'Не показываем кнопку выбора из справочника
   VAR_DirPictNotToShow = "DocName,Description"

   'Не показываем поля
   VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow + ",NameControl,AddFieldText1,AddFieldText2,ContractType,BusinessUnit"

   'Поле "Сторона договора", ссылка на справочник "Contracts/Сторона договора"
   S_UserFieldText1_Select = GetUserDirValues(CATALOG_CONTRACTING_PARTY)
   
'mikron_MIKRON_NDA_CONTRACT_end

' *********************************************************************************
' ***                                 ПРОДАЖИ МИКРОН                            ***
' *** 8.  МИКРОН EXPORT CONTRACTS                                               ***
' *********************************************************************************
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_EXPORT_CONTRACT)) = 1 or _
       InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_EXPADD_CONTRACT)) = 1 Then

   S_DocID_Set = S_DocID
   S_DocIDAdd_Set = S_DocIDAdd
   S_Department_Set = S_Department
   S_Author_Set = S_Author

   VNIK_TempValueAgree = S_ListToReconcile
   S_ListToReconcile = "" 'список согласующих
   If InStr(VNIK_TempValueAgree,"##") <= 0 Then
      VNIK_TempValueAgree = ""
   End If

   S_DocIDPrevious = Request("DocIDPrevious")
   S_DocIDPrevious_Set = S_DocIDPrevious

   If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
      VAR_ChangeDocGenerateButton = ""

      'Дополнительные пользователи: начальник отдела экспортных продаж
      S_AdditionalUsers = ReplaceRoleFromDir(SIT_HeadOfDocControl,SIT_MIKRON) + VbCrLf + """Дмитриев С. А."" <sdmitriev_ms>;"
      S_DocID = " "
      S_DocID_Set = S_DocID

      If Trim(S_DocIDParent) <> "" Then
         S_NameAproval = SIT_GetDocField(S_DocIDParent, "NameAproval", Conn)
         S_NameResponsible = SIT_GetDocField(S_DocIDParent, "NameResponsible", Conn)
         S_ListToReconcile = ""
         VNIK_TempValueAgree = ""
         S_Description = ""
      Else
         S_NameAproval = MIKRON_DeputyGDmarketing
         S_NameResponsible = """Ветчинин А. Е."" <avetchinin_ms>;"
         S_DocIDParent = " "
      End If

      S_Author = InsertionName(Session("Name"), Session("UserID"))
      S_Department = Session("Department")

'      S_Name = ""            'Вид обязательств
      S_DocIDIncoming = ""   '№договора (внешний)

      S_UserFieldText1 = ""  'контрагент (соисполнитель)
      S_UserFieldText2 = ""  'Проект
      S_UserFieldText3 = ""  'Доходный\расходный
   End If 'UCase(Request("create")) = "Y"

   S_AdditionalUsers_Set = iif (S_AdditionalUsers <> "", S_AdditionalUsers, "" )

   CurrentDocFieldOrder = "Name,DocID,DocIDParent,DocIDPrevious,Author,NameResponsible,Department,DocIDIncoming," + _
                          "PartnerName,UserFieldText1,UserFieldText2,UserFieldText3,Description," + _
                          "AmountDoc,Currency,UserFieldDate4,UserFieldDate6,ListToReconcile,NameAproval," + _
                          "SecurityLevel,UserFieldDate5"

   'поля помечаются звездочкой - обязательно к заполнению
   CurrentDocRequiredFields = "AmountDoc,Currency,Name,NameResponsible,PartnerName," + _
                              "UserFieldText2,UserFieldText3"

   'Убрать иконку справочников у полей
   VAR_DirPictNotToShow = "DocIDIncoming,Description,UserFieldText2,Name,Description,AmountDoc"

   'Скрываем ненужные поля
   VAR_DocFieldsNotToShow = "ContractType,AddFieldText1,AddFieldText2,BusinessUnit"

   'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
   'Всем запрещено, переводим на правила
   S_SecurityLevel = 4
   If not IsAdmin() Then
      S_Author_Set = S_Author
      S_Department_Set = S_Department
      S_NameAproval_Set = S_NameAproval
      S_SecurityLevel_Set = S_SecurityLevel
      S_NameResponsible_Set = S_NameResponsible
      S_DocIDParent_Set = S_DocIDParent
      S_DocIDPrevious_Set = S_DocIDPrevious

      If Trim(S_DocIDParent) <> "" Then  'здесь идут дополнения к эксп. контрактам
         S_UserFieldText4_Set = S_UserFieldText4
         'Контрагента не даем редактировать - берется из родительского документа
         S_PartnerName_Set = S_PartnerName       'контрагент
         'Поле "Наименование дополнения"
         If (S_Name = "") Then
            S_Name = Request("DocName")
         End If
         S_Name_Set = S_Name
         S_ListToReconcile_Comment = " Привет NNM " + S_Name + " Привет AMW"
         Select Case S_Name
            Case MIK_EA_1
               S_ListToReconcile_Comment = SIT_RequiredAgrees + oPayDox.GetExtTableValue("AgreeMIKRON","Name","Дополнение на разовую отгрузку (экспорт):","List") + vbCrLf + _
                          SIT_AdditionalAgrees
            Case MIK_EA_2
               S_ListToReconcile_Comment = SIT_PreliminaryAgrees + oPayDox.GetExtTableValue("AgreeMIKRON","Name",SIT_PreliminaryAgrees,"List") + vbCrLf + _
                          SIT_RequiredAgrees + oPayDox.GetExtTableValue("AgreeMIKRON","Name","Дополнение на расширение номенклатуры (экспорт):","List") + vbCrLf + _
                          SIT_AdditionalAgrees
            Case MIK_EA_3
               S_ListToReconcile_Comment = SIT_PreliminaryAgrees + oPayDox.GetExtTableValue("AgreeMIKRON","Name",SIT_PreliminaryAgrees,"List") + vbCrLf + _
                          SIT_RequiredAgrees + oPayDox.GetExtTableValue("AgreeMIKRON","Name","Дополнение на добавление спецификации (экспорт):","List") + vbCrLf + _
                          SIT_AdditionalAgrees
            Case MIK_EA_4
               S_ListToReconcile_Comment = SIT_PreliminaryAgrees + oPayDox.GetExtTableValue("AgreeMIKRON","Name",SIT_PreliminaryAgrees,"List") + vbCrLf + _
                          SIT_RequiredAgrees + oPayDox.GetExtTableValue("AgreeMIKRON","Name","Дополнение на изменение условий  и пр.(экспорт):","List") + vbCrLf + _
                          SIT_AdditionalAgrees
         End Select
         If VNIK_TempValueAgree <> "" Then
            S_ListToReconcile = AdditionalAgreeFromList(VNIK_TempValueAgree)
         End If
      Else
         bName = ""
         bDocIDParent = ""
         bDocIDPrevious = ""
         'Добавим комментарий по обязательным согласующим
         S_ListToReconcile_Comment = SIT_PreliminaryAgrees + oPayDox.GetExtTableValue("AgreeMIKRON","Name",SIT_PreliminaryAgrees,"List") + vbCrLf + _
                                     SIT_RequiredAgrees + oPayDox.GetExtTableValue("AgreeMIKRON","Name","Контракт (экспорт):","List") + vbCrLf + _
                                     SIT_AdditionalAgrees
         If VNIK_TempValueAgree <> "" Then
            S_ListToReconcile = AdditionalAgreeFromList(VNIK_TempValueAgree)
         End If
         CurrentDocRequiredFields = "AmountDoc,Currency,Name,NameResponsible,PartnerName," + _
                                    "UserFieldText2,UserFieldText3,UserFieldText4"
      End If
   Else
      S_ListToReconcile_Comment = ""
      S_ListToReconcile = VNIK_TempValueAgree
   End If

   'Поле "Проект", справочник "Закупки Микрон/Проекты"
   S_UserFieldText2_Select = GetUserDirValues(MIKRON_CATALOG_PROJECT_LIST)
   'Поле "Категория покупателя"
   S_UserFieldText3_Select = MIKRON_CATALOG_BUYERS
   'Поле "Cтатья бюджета", ссылка на справочник "Закупки Микрон/Статьи бюджета"
   S_UserFieldText4_Select = GetUserDirValues(MIKRON_CATALOG_EXPENDITURE)
   'Поле "Порядок оплаты", ссылка на справочник "Закупки Микрон/Порядок оплаты"
   S_UserFieldText5_Select = GetUserDirValues(MIKRON_CATALOG_PAYMENT_TYPES)

'mikron_MIKRON_EXP_CONTRACT_end

' AMW - MIKRON - END

' ********************************* ЗАЯВКА НА ЗАКУПКУ
' *** УК (СИТРОНИКС)
'vnik_purchase_order
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_PURCHASE_ORDER)) > 0 Then
  'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
  S_SecurityLevel = 4
  If not IsAdmin() Then
    S_SecurityLevel_Set = S_SecurityLevel
  End If

  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
    S_AdditionalUsers = """Успенский А. В."" <uspensky>;"
  End If

  If S_AdditionalUsers <> "" Then
    S_AdditionalUsers_Set = S_AdditionalUsers
  Else
    S_AdditionalUsers_Set = " "
  End If

  S_DocID_Set = S_DocID
  S_DocIDAdd_Set = S_DocIDAdd
  S_Department_Set = S_Department
  S_Author_Set = S_Author

  ' Создание карточки документа
  If UCase(Request("create")) = "Y" Then
    If Request("UpdateDoc") <> "YES" Then
      'SAY 2009-02-20
      VAR_ChangeDocGenerateButton=""
      S_DocID = " "
      If S_DocID_Set<>"" Then
        S_DocID_Set = S_DocID
      End If

      'SAY 2008-11-06
      S_Department_Set = Session("Department")

      S_DocID_Set = " "
            
      S_Author_Set = InsertionName(Session("Name"), Session("UserID"))
      S_Department = Session("Department")
      S_Department_Set = S_Department
      
      S_NameAproval = SIT_ChairmanOfCentralPurchasingCommission
      S_NameAproval_Set = S_NameAproval
      S_Name = "Заявка на закупку"
      
      S_Currency = "RUR"
      
      'сделаем поле наименование проекта нередактируемым вручную
      S_UserFieldText2_Set = ""
      
      'S_UserFieldText2 = "Банковский перевод"
      'S_UserFieldText6 = "Средний"
      'S_UserFieldText7 = "Месяц"
    End If
  End If 'UCase(Request("create")) = "Y"

  'сделаем поле наименование проекта нередактируемым вручную
  S_UserFieldText2_Set = S_UserFieldText2
    
  S_ListToReconcile = DeleteConstPrefixFromList(S_ListToReconcile)

  CurrentDocFieldOrder = "Name, DocID, DocIDParent, Author, Department, ContractType, UserFieldText8, PartnerName, Currency, UserFieldMoney1, UserFieldMoney2, QuantityDoc, UserFieldDate2, UserFieldText4, UserFieldText3, UserFieldText1, UserFieldText2, UserFieldText7, UserFieldText6, NameAproval, ListToReconcile, ListToView, UserFieldText5"
  ' поля помечаются звездочкой - обязательно к заполнению
  CurrentDocRequiredFields = "Author, PartnerName, UserFieldMoney1, UserFieldMoney2, Currency, NameAproval, UserFieldText4, UserFieldText3, UserFieldText7, Name, UserFieldText1, UserFieldText8, ContractType"
   
  'Не показываем кнопку выбора из справочника
  VAR_DirPictNotToShow = "UserFieldMoney1, UserFieldMoney2, UserFieldText6, UserFieldText2, UserFieldText7, DocQuantityDoc, UserFieldText8"
  '?
    
  'Добавим комментарий по обязательным согласубщим
  If InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) = 1 Then
     S_ListToReconcile_Comment = oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")
  ElseIf InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 then
     ' и т.д.
  End If

'kkoshkin 04042012 start
  If InStr(UCase(Session("Department")),UCase("СИТРОНИКС*SITRONICS*/Комплекс маркетинга и развития бизнеса*Complex marketing and business development*/"))=1 and UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" then
    S_ListToReconcile = S_ListToReconcile + " ""Кутуков К. В."" <kutukov>;"        
  end If
'kkoshkin 04042012 end

  VAR_DirPictNotToShow = VAR_DirPictNotToShow + ", DocName"
  'rmanyushin 60298 02.11.2009 Start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow + ", NameControl"
  'rmanyushin 60298 02.11.2009 End
'ph - 20120216 - start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", AddFieldText1"
'ph - 20120216 - end

 'S_UserFieldText1_Select = 'Код проекта
  S_UserFieldText3_Select = GetUserDirValues("{15EB5243-22D8-425D-B31A-9CBA4396FCFC}")
  S_UserFieldText4_Select = GetUserDirValues("{E2D63CEF-6DC4-44F3-B4BF-BD4B0D9438D0}") 
  'S_UserFieldText5_Select = GetUserDirValues("{507B2058-B7C4-40B2-8EC4-D75A8E4CE28D}")
  'S_UserFieldText6_Select = GetUserDirValues("{3ECADCD6-0985-4659-8774-C8C9D77EE381}")
  'S_UserFieldText7_Select = GetUserDirValues("{C62278E8-50EB-41AF-89C4-90717DF68414}")
  S_Currency_Select = GetUserDirValues("{B53443BE-B30E-4C52-98D7-5ED0BEE67080}")   
  S_AddField_Select3 = GetUserDirValues("{B1DE9C2C-0B37-420E-9884-0C09D8CFF2DD}") 'Способ закупки
 
  'Для пользователей СТС везде добавляем поле бизнес единица
  If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
     If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 Then
        If UCase(Request("create")) = "Y" Then
           sCreator = Session("UserID")
        Else
           sCreator = GetUserID(S_Author)
        End If
        BusinessUnitsList = GetUsersBusinessUnits(sCreator)
        If InStr(BusinessUnitsList, VbCrLf) > 0 Then
           S_AddField2 = " "
           S_AddField_Select2 = BusinessUnitsList
           'Поле обязательно для заполнения
           CurrentDocRequiredFields = CurrentDocRequiredFields+", BusinessUnit"
        Else
           S_AddField_Set2 = BusinessUnitsList
        End If
     End If
  Else
     ' Поле показывается только для пользователей СТС и администратора
     ' (у админа не из СТС оно необязательное, редактируемое)
     If not IsAdmin() Then
        VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow + ", BusinessUnit"
     End If
  End If
'vnik_purchase_order

' ********************************* ДОГОВОРЫ
' **** УК
'vnik_contracts
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_CONTRACTS_MC)) > 0 Then
  'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
  S_SecurityLevel = 4
  If not IsAdmin() Then
     S_SecurityLevel_Set = S_SecurityLevel
  End If

  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
    S_AdditionalUsers = """Успенский А. В."" <uspensky>;"
  End If

  If S_AdditionalUsers <> "" Then
     S_AdditionalUsers = S_AdditionalUsers + VbCrLf + ReplaceRoleFromDir(SIT_ChiefOfPurchaseDepartment,GetRootDepartment(S_Department))
  Else
     S_AdditionalUsers = S_AdditionalUsers + ReplaceRoleFromDir(SIT_ChiefOfPurchaseDepartment,sDepartmentRoot)
  End If

  If S_AdditionalUsers <> "" Then
     S_AdditionalUsers_Set = S_AdditionalUsers
  Else
     S_AdditionalUsers_Set = " "
  End If

  S_DocID_Set = S_DocID
  S_DocIDAdd_Set = S_DocIDAdd
  S_Department_Set = S_Department
  S_Author_Set = S_Author
  
  ' Создание карточки документа
  If UCase(Request("create")) = "Y" Then
     If Request("UpdateDoc") <> "YES" Then
        'SAY 2009-02-20
        VAR_ChangeDocGenerateButton=""
        S_DocID = " "
        If S_DocID_Set <> "" Then
           S_DocID_Set = S_DocID
        End If
        'SAY 2008-11-06

        S_Department_Set = Session("Department")

        S_DocID_Set = " "
      
        S_DocIDParent_Set = S_DocIDParent
            
        S_Author_Set = InsertionName(Session("Name"), Session("UserID"))
        S_Department = Session("Department")
        S_Department_Set = S_Department
      
        S_NameAproval = ""
       'S_NameAproval = SIT_SignatoryOfTheContractsMC
      'S_NameAproval_Set = S_NameAproval
      
        S_LocationPath = ReplaceRoleFromDir(SIT_Registrar, iif(InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1, SIT_STS, SIT_SITRONICS))
        S_LocationPath_Set = S_LocationPath
      
        If Trim(S_DocIDParent)<>"" Then
           If InStr(UCase(S_DocIDParent), UCase("POHQ-")) > 0 Then
              S_Name = ""
              S_PartnerName_Set = S_Partner
              S_Currency_Set = S_Currency
              S_UserFieldText1_Set = S_UserFieldText1
              S_UserFieldText2_Set = S_UserFieldText2
              'сделаем поля с суммами нередактируемыми вручную, пока не работает
              S_UserFieldMoney1_Set = S_UserFieldMoney1
              S_UserFieldMoney2_Set = S_UserFieldMoney2
              S_UserFieldText4 = ""
              S_UserFieldText5 = ""
              S_UserFieldText6 = ""
              S_UserFieldText7 = ""
              S_ListToReconcile = ""
           ElseIf InStr(UCase(S_DocIDParent), UCase("PR-CPC-")) > 0 Then
              vnik_DocIDParent = SIT_GetDocField(S_DocIDParent, "DocIDParent", Conn)
              S_Name = ""
              S_Partner = SIT_GetDocField(vnik_DocIDParent, "PartnerName", Conn)
              S_PartnerName_Set = S_Partner 
              S_Currency_Set = SIT_GetDocField(vnik_DocIDParent, "Currency", Conn)
              S_Content = ""  'Содержание д-та
              S_UserFieldMoney1 = SIT_GetDocField(vnik_DocIDParent, "UserFieldMoney1", Conn)
              S_UserFieldMoney1_Set = S_UserFieldMoney1    
              S_UserFieldMoney2 = SIT_GetDocField(vnik_DocIDParent, "UserFieldMoney2", Conn)
              S_UserFieldMoney2_Set = S_UserFieldMoney2 
              S_UserFieldText1_Set = SIT_GetDocField(vnik_DocIDParent, "UserFieldText1", Conn)
              S_UserFieldText2_Set = SIT_GetDocField(vnik_DocIDParent, "UserFieldText2", Conn)
              S_QuantityDoc_Set = SIT_GetDocField(vnik_DocIDParent, "QuantityDoc", Conn)           
           End If
           S_UserFieldText8_Set = "Не рамочный"
        Else
           S_UserFieldText8_Set = "Рамочный"
           S_UserFieldMoney1 = 0
           S_UserFieldMoney1_Set = S_UserFieldMoney1
           S_UserFieldMoney2 = 0
           S_UserFieldMoney2_Set = S_UserFieldMoney2
        End If
        S_ListToView = ""  
        S_LocationPath_Set = S_LocationPath 
     End If
  End If 'UCase(Request("create")) = "Y"
  
  'If InStr(UCase(ReplaceRoleFromDir(SIT_OldContractOperator,Session("Department"))),UCase(Session("UserID"))) = 0 Then
    'не работает
  '  S_PartnerName_Set = S_Partner
  '  S_UserFieldMoney1_Set = S_UserFieldMoney1
  '  S_UserFieldMoney2_Set = S_UserFieldMoney2
  'End If
  
  'Ответственный исполнитель равно автор
  S_NameResponsible = S_Author_Set
  
  S_ListToReconcile = DeleteConstPrefixFromList(S_ListToReconcile)

  CurrentDocFieldOrder = "Name, UserFieldText8, DocID, DocIDParent, Author, NameResponsible, Department, PartnerName, Content, Currency, UserFieldMoney1, UserFieldMoney2, UserFieldText4, UserFieldDate3, UserFieldText1, UserFieldText2, Description, UserFieldText7, UserFieldText6, NameAproval, ListToReconcile, ListToView, UserFieldText5, UserFieldText3, AddFieldText2, AddFieldText1"
  ' поля помечаются звездочкой - обязательно к заполнению
  CurrentDocRequiredFields = "Author, NameResponsible, PartnerName, UserFieldMoney1, UserFieldMoney2, Currency, NameAproval, UserFieldText5, Name, UserFieldText1, UserFieldText2, UserFieldDate4, UserFieldDate5, UserFieldDate6, Content"
   
  'Не показываем кнопку выбора из справочника
  VAR_DirPictNotToShow = "UserFieldMoney1, UserFieldMoney2, UserFieldText4, UserFieldText2, UserFieldText5, QuantityDoc, DocDescription, Content, UserFieldText6, UserFieldText7, UserFieldText3, UserFieldText8, AddFieldText2"
  
  'повтор
  'If InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) = 1 Then  
  '  S_ListToReconcile_Comment = oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")
  'Else
  '  '
  'End If
  
  'Добавим комментарий по обязательным согласубщим
  If InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) = 1 Then
     S_ListToReconcile_Comment = oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")
  ElseIf InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 then
     ' и т.д.
  End If

  VAR_DirPictNotToShow = VAR_DirPictNotToShow+", DocName"
  'rmanyushin 60298 02.11.2009 Start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", ContractType, NameControl,DateApproved,Docs_DateRegistered"
  'rmanyushin 60298 02.11.2009 End
'ph - 20120216 - start
  'VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", AddFieldText1"
'ph - 20120216 - end

  ' Вид Обязательств
  Select Case UCase(Request("l"))
    Case "RU" 'RU
       S_Name_Select = GetUserDirValues("{535C7740-CB1E-4403-ABC6-93AFC67205D5}")
    Case "" 'EN
       S_Name_Select = GetUserDirValues("{653D0ABF-6E6E-4DCA-899F-51D44A3DB2C5}")
    Case "3" 'CZ
       S_Name_Select = GetUserDirValues("{E30E310A-67DC-4BE2-8023-61F3214EAB4D}")
  End Select

  'S_UserFieldText1_Select = 'Код проекта
  'S_UserFieldText3_Select = GetUserDirValues("{15EB5243-22D8-425D-B31A-9CBA4396FCFC}")
  'S_UserFieldText4_Select = GetUserDirValues("{E2D63CEF-6DC4-44F3-B4BF-BD4B0D9438D0}") 
  'S_UserFieldText5_Select = GetUserDirValues("{507B2058-B7C4-40B2-8EC4-D75A8E4CE28D}")
  'S_UserFieldText6_Select = GetUserDirValues("{3ECADCD6-0985-4659-8774-C8C9D77EE381}")
  'S_UserFieldText7_Select = GetUserDirValues("{C62278E8-50EB-41AF-89C4-90717DF68414}")
  'S_Currency_Select       = GetUserDirValues("{B53443BE-B30E-4C52-98D7-5ED0BEE67080}")   

  S_UserFieldText3_Select = SIT_YesNo
  S_AddField_Select5 = SIT_YesNo
  S_AddField_Select4 = SIT_YesNo
 
  'Для пользователей СТС везде добавляем поле бизнес единица
  If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
     If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 Then
        If UCase(Request("create")) = "Y" Then
           sCreator = Session("UserID")
        Else
           sCreator = GetUserID(S_Author)
        End If
        BusinessUnitsList = GetUsersBusinessUnits(sCreator)
        If InStr(BusinessUnitsList, VbCrLf) > 0 Then
           S_AddField2 = " "
           S_AddField_Select2 = BusinessUnitsList
           'Поле обязательно для заполнения
           CurrentDocRequiredFields = CurrentDocRequiredFields+", BusinessUnit"
        Else
           S_AddField_Set2 = BusinessUnitsList
        End If
     End If
  Else
     ' Поле показывается только для пользователей СТС и администратора
     ' (у админа не из СТС оно необязательное, редактируемое)
     If not IsAdmin() Then
        VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow + ", BusinessUnit"
     End If
  End If
'vnik_contracts

' ********************************* НОРМАТИВНЫЕ ДОКУМЕНТЫ
' ***
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_NORM_DOCS)) = 1 Then
  'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
  S_SecurityLevel = 4
  If not IsAdmin() Then
    S_SecurityLevel_Set = S_SecurityLevel
  End If

  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
    S_AdditionalUsers = ""
  End If

  If S_AdditionalUsers <> "" Then
    S_AdditionalUsers_Set = S_AdditionalUsers
  Else
    S_AdditionalUsers_Set = " "
  End If

  S_DocID_Set = S_DocID
  S_DocIDAdd_Set = S_DocIDAdd
  S_Department_Set = S_Department
  S_Author_Set = S_Author

  ' Создание карточки документа
  If UCase(Request("create")) = "Y" Then
    If Request("UpdateDoc") <> "YES" Then
      'SAY 2009-02-20
      VAR_ChangeDocGenerateButton=""
      S_DocID = " "
      If S_DocID_Set<>"" Then
        S_DocID_Set = S_DocID
      End If

      'SAY 2008-11-06
      S_Department_Set = Session("Department")

      'SAY 2008-09-22 для подчиненных документов оставляем весь список согласования родителя
      If Trim(S_DocIDParent)<>"" then
        S_ListToReconcile = Replace(S_ListToReconcile,"##;","")
      End If

      S_DocID_Set = " "
      S_DocIDadd_Set = " "
      S_DateActivation = MyDate(Date)
      S_UserFieldDate1 = MyDate(Date)
      S_Author_Set = InsertionName(Session("Name"), Session("UserID"))
      S_Department_Set = Session("Department")

      'SAY 2008-11-10 переделываем регистратора с роли на пользователя
      S_LocationPath = ReplaceRoleFromDir(SIT_Registrar, iif(InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1, SIT_STS, SIT_SITRONICS))
      S_LocationPath_Set = S_LocationPath

      'phil - 20080906 - Start - Дата автоматического согласования
      If InStr(Session("Department"), "СИТРОНИКС*SITRONICS") = 1 then
        ' Sitronics
        S_UserFields(14) = MyDate(GetNormDocLastReconcileDate)
      Else 
        ' остальные БН - не показываем это поле
        S_UserFields(14) = ""
        VAR_DocFieldsNotToShow = "UserFieldDate3"
      End If
      'phil - 20080906 - End
    End If
  End If 'UCase(Request("create")) = "Y"

  S_ListToReconcile = DeleteConstPrefixFromList(S_ListToReconcile)

  CurrentDocFieldOrder = "Name, UserFieldText1, UserFieldDate1, DocID, Author, DateActivation, Content, Description, UserFieldText3, UserFieldText4, UserFieldDate2 ,UserFieldText2, ListToReconcile, NameAproval, LocationPath, ListToView, SecurityLevel, Department "	
  CurrentDocRequiredFields = "Name, UserFieldText1, UserFieldDate1, DateActivation, UserFieldText4, UserFieldDate2 ,UserFieldText2, NameAproval, DocListToRegister, UserFieldDate3 "

  'SAY 2008-09-16 закрываем от записи поле "версия"
  If UCase(Request("create")) <> "Y" Then
    S_UserFieldText3_Set = S_UserFieldText3
  End If

  S_LocationPath_Set=S_LocationPath

  If InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) = 1 Then  
    S_ListToReconcile_Comment = oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")
  Else
    S_ListToReconcile = oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")
  End IF

  S_UserFieldText1_Set = S_UserFieldText1
  VAR_DirPictNotToShow = "UserFieldText3"

  VAR_DirPictNotToShow = VAR_DirPictNotToShow+", DocName"
  'rmanyushin 60298 02.11.2009 Start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", ContractType"
  'rmanyushin 60298 02.11.2009 End
'ph - 20120216 - start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", AddFieldText1"
'ph - 20120216 - end

  'Выпадающие списки
  Select Case UCase(Request("l"))
    Case "RU" 'RU
      S_UserFieldText1_Select = GetUserDirValues("{3685D3AA-FB15-4ECF-993F-8AC5AB87F4D6}")
      S_UserFieldText2_Select = GetUserDirValues("{37E16CD5-BC8F-4D0C-9569-D14DAA895440}")
    Case "" 'EN
      S_UserFieldText1_Select = GetUserDirValues("{D6E49442-500A-42CA-AF56-37FE90DAFE3C}")
      S_UserFieldText2_Select = GetUserDirValues("{7C7058BE-F586-44C4-B5BE-47D2E05E96BD}")
    Case "3" 'CZ
      S_UserFieldText1_Select = GetUserDirValues("{4C4E59F5-3DCF-47B0-BF4E-38D95F626746}")
      S_UserFieldText2_Select = GetUserDirValues("{AF173F9A-4724-405B-AA9C-C72E1DCA7647}")
  End Select

  'Для пользователей СТС везде добавляем поле бизнес единица
  If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
    If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 Then
      If UCase(Request("create")) = "Y" Then
        sCreator = Session("UserID")
      Else
        sCreator = GetUserID(S_Author)
      End If
      BusinessUnitsList = GetUsersBusinessUnits(sCreator)
      If InStr(BusinessUnitsList, VbCrLf) > 0 Then
        S_AddField2 = " "
        S_AddField_Select2 = BusinessUnitsList
        'Поле обязательно для заполнения
        CurrentDocRequiredFields = CurrentDocRequiredFields+", BusinessUnit"
      Else
        S_AddField_Set2 = BusinessUnitsList
      End If
    End If
  Else
    'Поле показывается только для пользователей СТС и администратора (у админа не из СТС оно необязательное, редактируемое)
    If not IsAdmin() Then
      VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", BusinessUnit"
    End If
  End If

' ********************************* ПОРУЧЕНИЯ
' ****
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_ZADACHI)) = 1 Then
  'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
  S_SecurityLevel = 4
  If not IsAdmin() Then
    S_SecurityLevel_Set = S_SecurityLevel
  End If

  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
    S_AdditionalUsers = ""
  End If

  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" and InStr(UCase(Session("Department")), UCase(SIT_RTI)) = 1 Then
     S_AdditionalUsers = """Боев С. Ф."" <boev_oaorti>;"
    If InStr(UCase(Session("Department")), UCase("Управление ИТ")) > 0 Then
      S_AdditionalUsers = S_AdditionalUsers + VbCrLf + """Иванов В. В."" <vivanon_oaorti>;"
    End If
     DocIDParent_RTI = SIT_GetDocField(S_DocIDParent, "Correspondent", Conn)
     DocIDPar_RTI = SIT_GetDocField(S_DocIDParent, "DocId", Conn)
     If len(DocIDParent_RTI) > 0 and InStr(UCase(DocIDPar_RTI),"T_") > 0 Then
        S_AdditionalUsers = S_AdditionalUsers + VbCrLf + DocIDParent_RTI
     End If
  End If
  
  'minc 
  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" and InStr(UCase(Session("Department")), UCase(SIT_MINC)) = 1 Then
    S_AdditionalUsers = """Селютина О. А."" <usr_oselyutina>;" + VbCrLf + """Бабушкин А. Н."" <usr_ababushkin>;" + VbCrLf + """Теппер А. Б."" <usr_atepper>;" + VbCrLf
    DocIDParent_RTI = SIT_GetDocField(S_DocIDParent, "Correspondent", Conn)
    DocIDPar_RTI = SIT_GetDocField(S_DocIDParent, "DocId", Conn)
    if len(DocIDParent_RTI) >0 and InStr(UCase(DocIDPar_RTI),"К_") > 0 Then
        S_AdditionalUsers = S_AdditionalUsers + VbCrLf + DocIDParent_RTI
    End If
  End If
  'minc

  'vtss
  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" and InStr(UCase(Session("Department")), UCase(SIT_VTSS)) = 1 Then
    S_AdditionalUsers = """Подольский А. Е."" <vtss_a.podolskii>;" + VbCrLf
    DocIDParent_RTI = SIT_GetDocField(S_DocIDParent, "Correspondent", Conn)
    DocIDPar_RTI = SIT_GetDocField(S_DocIDParent, "DocId", Conn)
    if len(DocIDParent_RTI) >0 and InStr(UCase(DocIDPar_RTI),"T_") > 0 Then
        S_AdditionalUsers = S_AdditionalUsers + VbCrLf + DocIDParent_RTI
    End If
  End If
  'vtss

  If S_AdditionalUsers <> "" Then
     S_AdditionalUsers_Set = S_AdditionalUsers
  Else
     S_AdditionalUsers_Set = " "
  End If

  S_DocID_Set = S_DocID
  S_DocIDAdd_Set = S_DocIDAdd
  S_Department_Set = S_Department
  S_Author_Set = S_Author

  ' Создание карточки документа
  If UCase(Request("create")) = "Y" Then
    If Request("UpdateDoc") <> "YES" Then
      'SAY 2009-02-20
      VAR_ChangeDocGenerateButton=""
      S_DocID = " "
      If S_DocID_Set<>"" Then
        S_DocID_Set = S_DocID
      End If

      'SAY 2008-11-06
      S_Department_Set = Session("Department")

      'SAY 2008-09-22 для подчиненных документов оставляем весь список согласования родителя
      If Trim(S_DocIDParent)<>"" then
        S_ListToReconcile = Replace(S_ListToReconcile,"##;","")
      End If

      S_Name = SIT_GetDocField(S_DocIDParent, "Name", Conn)
      S_DocID = " "
      S_DocID_Set = S_DocID
      S_UserFieldText1 = ""
      
      sSufix = ""
      If Trim(S_DocIDParent) <> "" then
        'Ph - 20090303 - При создании подчиненного поручения убрать копирование контролера
        S_NameControl = ""
		if InStr(UCase(Session("Department")), UCase(SIT_MINC)) = 1 Then
        S_Name = ""
        end if
		if InStr(UCase(Session("Department")), UCase(SIT_VTSS)) = 1 Then
        S_Name = ""
        end if
        sSufix = "("+S_DocIDParent + ")_"
        'SAY 2008-11-10 Настроить автоматическую подстановку значения по умолчанию для поля «Вид поручения»
        If InStr(UCase(S_DocIDParent),"IN") > 0 Then
          If InStr(S_UserFieldText7, "1 - ") = 1 Then
            S_UserFieldText1 = SIT_SistemaTasks
          Else
            S_UserFieldText1 = SIT_IncomingMailTasks
          End If
        End If
        If InStr(UCase(S_DocIDParent),"OR") > 0 or InStr(UCase(S_DocIDParent),"P") > 0 Then
          S_UserFieldText1 = SIT_TasksOnOrders
        End If
        If InStr(UCase(S_DocIDParent),"IH") > 0 Then
          S_UserFieldText1 = SIT_TasksOnMemos
        End If
      Else
        'SAY 2008-11-10 Настроить автоматическую подстановку значения по умолчанию для поля «Вид поручения»        
         S_UserFieldText1 = SIT_TasksOnOtherDocs
      End If
      S_DateActivation = MyDate(Date)
      S_Author_Set = InsertionName(Session("Name"), Session("UserID"))
      S_Department_Set = Session("Department")
      
      S_Content = ""
      'Ph - 20081109 - Очищать при создании подчиненного поручения список соисполнителей
      S_Correspondent = ""

      'SAY 2008-11-10
      S_DateCompletion = MyDate(Date+14)
      If InStr(UCase(Session("Department")), UCase(SIT_RTI)) = 1 Then S_DateCompletion = MyDate(Date+30)
      If InStr(UCase(Session("Department")), UCase(SIT_MINC)) = 1 Then S_DateCompletion = MyDate(Date+10)
      If InStr(UCase(Session("Department")), UCase(SIT_VTSS)) = 1 Then S_DateCompletion = MyDate(Date+5)
    End If
  End If 'UCase(Request("create")) = "Y"

  S_ListToReconcile = DeleteConstPrefixFromList(S_ListToReconcile)

  CurrentDocFieldOrder = "DocID, DocIDParent, Name, UserFieldText1, Author, DateActivation, DateCompletion, Content, Rank, NameResponsible, Correspondent, Context, NameControl, SecurityLevel, Department "	
  bResult = ""
  CurrentDocRequiredFields = "Name, UserFieldText1, DateActivation, DateCompletion, NameResponsible"	
  'SAY 2008-11-10 
  S_DateActivation_Set = S_DateActivation

  VAR_DirPictNotToShow = VAR_DirPictNotToShow+", DocName"
  'rmanyushin 60298 02.11.2009 Start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", ContractType"
  'rmanyushin 60298 02.11.2009 End
'ph - 20120216 - start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", AddFieldText1"
'ph - 20120216 - end

  'Выпадающие списки
  Select Case UCase(Request("l"))
    Case "RU" 'RU
      S_UserFieldText1_Select = GetUserDirValues("{C632B46B-3AAF-4607-BBC5-AC51C0A4971B}")
      If Session("UserID") <> GetUserID(MyGetUserDirValue("Роли СТС RU", STS_SecrPravlenia, 1, 2)) and not InStr(UCase(MyGetUserDirValue("Роли", SIT_SecrPravlenia, 1, 2)),UCase(Session("UserID"))) > 0 and not IsAdmin() Then
        S_UserFieldText1_Select = Replace(S_UserFieldText1_Select, PORUCHENIA_PRAVLENIA_RU+VbCrLf, "")
      End If
    Case "" 'EN
      S_UserFieldText1_Select = GetUserDirValues("{A75D5931-546E-401E-8925-71A7EA75889E}")
      If Session("UserID") <> GetUserID(MyGetUserDirValue("Роли СТС RU", STS_SecrPravlenia, 1, 2)) and not InStr(UCase(MyGetUserDirValue("Роли", SIT_SecrPravlenia, 1, 2)),UCase(Session("UserID"))) > 0 and not IsAdmin() Then
        S_UserFieldText1_Select = Replace(S_UserFieldText1_Select, PORUCHENIA_PRAVLENIA_EN+VbCrLf, "")
      End If
    Case "3" 'CZ
      S_UserFieldText1_Select = GetUserDirValues("{C9E55EA0-3AC8-4D26-9C74-17FA135EF1A5}")
      If Session("UserID") <> GetUserID(MyGetUserDirValue("Роли СТС RU", STS_SecrPravlenia, 1, 2)) and not InStr(UCase(MyGetUserDirValue("Роли", SIT_SecrPravlenia, 1, 2)),UCase(Session("UserID"))) > 0 and not IsAdmin() Then       
        S_UserFieldText1_Select = Replace(S_UserFieldText1_Select, PORUCHENIA_PRAVLENIA_CZ+VbCrLf, "")
      End If
  End Select
  
  'для пользователей РТИ свой справочник видов поручений
  If InStr(UCase(Session("Department")), UCase(SIT_RTI)) = 1 Then
    S_UserFieldText1_Select = GetUserDirValues("{940A2530-081D-4E5A-9C3D-751083E0BBF5}")
    'скрыть поле Примечание
    VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", UserFieldText8"
  end if

  'для пользователей МИКРОН свой справочник видов поручений
  If InStr(UCase(Session("Department")), UCase(SIT_MIKRON)) = 1 Then
    S_UserFieldText1_Select = GetUserDirValues("{940A2530-081D-4E5A-9C3D-751083E0BBF5}")
    'скрыть поле Примечание
    VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", UserFieldText8"
  end if
  
  'для пользователей MINC свой справочник видов поручений
  If InStr(UCase(Session("Department")), UCase(SIT_MINC)) = 1 Then
    S_UserFieldText1_Select = GetUserDirValues("{C8CA5264-E304-4AE2-9735-3FC5048313F7}")
    CurrentDocFieldOrder = "DocID, DocIDParent, Name, UserFieldText1, Author, DateActivation, DateCompletion, Content, Rank, NameResponsible, Correspondent, Context, NameControl, UserFieldText8, SecurityLevel, Department "
  end if
  
    'для пользователей ВТСС свой справочник видов поручений
  If InStr(UCase(Session("Department")), UCase(SIT_VTSS)) = 1 Then
    S_UserFieldText1_Select = GetUserDirValues("{60934932-64EC-444E-8E66-C41689E14107}")
    'скрыть поле Примечание
    VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", UserFieldText8"
  end if

  'Для пользователей СТС везде добавляем поле бизнес единица
  If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
    If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 Then
      If UCase(Request("create")) = "Y" Then
        sCreator = Session("UserID")
      Else
        sCreator = GetUserID(S_Author)
      End If
      BusinessUnitsList = GetUsersBusinessUnits(sCreator)
      If InStr(BusinessUnitsList, VbCrLf) > 0 Then
        S_AddField2 = " "
        S_AddField_Select2 = BusinessUnitsList
        'Поле обязательно для заполнения
        CurrentDocRequiredFields = CurrentDocRequiredFields+", BusinessUnit"
      Else
        S_AddField_Set2 = BusinessUnitsList
      End If
    End If
  Else
    'Поле показывается только для пользователей СТС и администратора (у админа не из СТС оно необязательное, редактируемое)
    If not IsAdmin() Then
      VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow + ", BusinessUnit"
    End If
  End If
  
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", AddFieldText2"

' ********************************* ДОГОВОРЫ ДО ДАТЫ ...
' ****
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_OLD)) = 1 Then
  'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
  S_SecurityLevel = 4
  If not IsAdmin() Then
    S_SecurityLevel_Set = S_SecurityLevel
  End If

  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
    S_AdditionalUsers = ""
  End If

  If S_AdditionalUsers <> "" Then
    S_AdditionalUsers_Set = S_AdditionalUsers
  Else
    S_AdditionalUsers_Set = " "
  End If

  S_DocID_Set = S_DocID
  S_DocIDAdd_Set = S_DocIDAdd
  S_Department_Set = S_Department
  S_Author_Set = S_Author

  ' Создание карточки документа
  If UCase(Request("create")) = "Y" Then
    'Предупреждение пользователю
    S_UserInstructions = "<font color = red><b>" & SIT_CannotCreateOldContracts & "</b></font>"

    If Request("UpdateDoc") <> "YES" Then
      'SAY 2009-02-20
      VAR_ChangeDocGenerateButton=""
      S_DocID = " "
      If S_DocID_Set<>"" Then
        S_DocID_Set = S_DocID
      End If

      'SAY 2008-11-06
      S_Department_Set = Session("Department")

      'SAY 2008-09-22 для подчиненных документов оставляем весь список согласования родителя
      If Trim(S_DocIDParent)<>"" then
        S_ListToReconcile = Replace(S_ListToReconcile,"##;","")
      End If

      S_DocID_Set = " "
      S_DocIDAdd_Set = " "
      S_DateActivation = MyDate(Date)
      S_Author_Set = InsertionName(Session("Name"), Session("UserID"))
      S_Department = Session("Department")
      S_Department_Set = S_Department
	  
      'rmanyushin 60298 02.11.2009 Start
      'Список обязательных согласующих вызываем из справочника
      If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
        S_ListToReconcile_Comment = oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")
        S_ListToReconcile = ""
        S_ListToReconcile_Set = ""
      End If
      'rmanyushin 60298 02.11.2009 
    End If
  End If 'UCase(Request("create")) = "Y"

  'rmanyushin 60298 02.11.2009  Устанавливаем порядок следования полей в документе
  CurrentDocFieldOrder = "DocID, DocIDAdd, Author, NameResponsible, UserFieldText4, UserFieldDate1, PartnerName, UserFieldText1, ContractType, Description, Name, DocIDIncoming, QuantityDoc, UserFieldText8, InventoryUnit, AmountDoc, UserFieldMoney1, Currency, UserFieldText5, UserFieldText6, UserFieldText7, UserFieldDate4, UserFieldDate5, UserFieldDate6, Content, ListToReconcile, NameAproval, UserFieldText2, UserFieldText3, Department "

  'rmanyushin 60298, 110100 13.07.2010 Утверждающий обязательное поле.
  'Если пользователь из СТС, то обязательным полем сделать и поле ContractType.
  If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
    CurrentDocRequiredFields = "Name, ContractType, NameAproval"
  Else
    CurrentDocRequiredFields = "Name" 
    'SAY 2008-09-15
  End If

  '20091110 - start - В договорах при редактировании жесткую часть не обновляем, могло быть делегирование
  If UCase(Request("create")) = "Y" Then
		'rmanyushin 60298 02.11.2009 Если пользователь из СТС, то для определения согласующих использовать справочник AgreeSTSRU
		If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
			S_ListToReconcile_Comment = oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")
		Else
			S_ListToReconcile_Comment = oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")
		End If
  Else
		If Request("UpdateDoc") = "YES" or not bUserCheckOK Then
      S_ListToReconcile_Comment = Session("S_ListToReconcile_Comment")
		Else
      If InStr(S_ListToReconcile, SIT_AdditionalAgreesDelimeter) > 0 Then
        S_ListToReconcile_Comment = Left(S_ListToReconcile, InStr(S_ListToReconcile, SIT_AdditionalAgreesDelimeter)-1)
      Else
        S_ListToReconcile_Comment = ""
      End If
      Session("S_ListToReconcile_Comment") = S_ListToReconcile_Comment
    End If
    S_ListToReconcile = DeleteConstPrefixFromList(S_ListToReconcile)
  End If
  '20091110 - end
	  
  'Ph - 20080921 - Убрать иконку справочников у полей
  VAR_DirPictNotToShow = "UserFieldText4, DocQuantityDoc"

  'rmanyushin 60298 02.11.2009 Start 'Скрываем поле ContractType в договорах УК и в договорах СТС, созданных до ноября 2009 г.
  If InStr(UCase(Session("Department")), UCase(SIT_STS)) <> 1 Then 
    VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", ContractType"
  Else
    If Request("create") <> "y" Then 'при создании нового документа условие по дате не проверяем
      Select Case UCase(Request("l"))
        Case "RU" 'RU
          dS_DateCreation = Replace(S_DateCreation, ".", ",")
          dS_DateCreation = Day(dS_DateCreation) & "/" & Month(dS_DateCreation) & "/" & Year (dS_DateCreation)
        Case "" 'EN
          dS_DateCreation = FormatDateTime(S_DateCreation,2)
        Case "3" 'CZ
          dS_DateCreation = Replace(S_DateCreation, ".", ",")
      End Select

      dLineDate = DateSerial(2009, 11, 1)
      dDateDiff = DateDiff("d", dLineDate, dS_DateCreation) 

      If dDateDiff < 0 Then
        VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", ContractType"
      End If
    End If

    VAR_DirPictNotToShow = VAR_DirPictNotToShow+", ContractType"
  End If
  'rmanyushin 60298 02.11.2009 End
'ph - 20120216 - start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", AddFieldText1"
'ph - 20120216 - end

  Select Case UCase(Request("l"))
    Case "RU" 'RU
      S_Name_Select = GetUserDirValues("{535C7740-CB1E-4403-ABC6-93AFC67205D5}")
    Case "" 'EN
      S_Name_Select = GetUserDirValues("{653D0ABF-6E6E-4DCA-899F-51D44A3DB2C5}")
    Case "3" 'CZ
      S_Name_Select = GetUserDirValues("{E30E310A-67DC-4BE2-8023-61F3214EAB4D}")
  End Select
  S_Currency_Select = GetCurrencyList()

  'rmanyushin 29.12.2009 start 'При редактировании документа нельзя поменять значение поля ContractType 
  If Request("create") = "y" Then
    S_AddField_Select3 = GetUserDirValues(STS_ContractPartyDirGUID) 
  Else
    Set dsContractType = Server.CreateObject("ADODB.Recordset")
    strSQL = "select ContractType from Docs where DocID = N'"+Request("DocID")+"'"
    dsContractType.Open strSQL, Conn, 3, 1, &H1
    If Not dsContractType.EOF Then str_S_AddField3 = dsContractType("ContractType")
    dsContractType.Close
    S_AddField_Set3 = str_S_AddField3
  End If
  'rmanyushin 29.12.2009 end 

  'Для пользователей СТС везде добавляем поле бизнес единица
  If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
    If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 Then
      If UCase(Request("create")) = "Y" Then
        sCreator = Session("UserID")
      Else
        sCreator = GetUserID(S_Author)
      End If
      BusinessUnitsList = GetUsersBusinessUnits(sCreator)
      If InStr(BusinessUnitsList, VbCrLf) > 0 Then
        S_AddField2 = " "
        S_AddField_Select2 = BusinessUnitsList
        'Поле обязательно для заполнения
        CurrentDocRequiredFields = CurrentDocRequiredFields+", BusinessUnit"
      Else
        S_AddField_Set2 = BusinessUnitsList
      End If
    End If
  Else
    'Поле показывается только для пользователей СТС и администратора (у админа не из СТС оно необязательное, редактируемое)
    If not IsAdmin() Then
      VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", BusinessUnit"
    End If
  End If
  
' ********************************* ДОГОВОРЫ РТИ ...
' ****
'rti_contract
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(RTI_CONTRACT)) = 1 Then

  'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
  S_SecurityLevel = 4
  If not IsAdmin() Then
    S_SecurityLevel_Set = S_SecurityLevel
  End If

  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
    S_AdditionalUsers = """Боев С. Ф."" <boev_oaorti>;"
    If InStr(UCase(Session("Department")), UCase("Управление ИТ")) > 0 Then
      S_AdditionalUsers = S_AdditionalUsers + VbCrLf + """Иванов В. В."" <vivanon_oaorti>;"
    End If
    S_AdditionalUsers_Set = S_AdditionalUsers
  End If 

  If S_AdditionalUsers <> "" Then
    S_AdditionalUsers_Set = S_AdditionalUsers
  Else
    S_AdditionalUsers_Set = " "
  End If

  S_DocID_Set = S_DocID
  S_DocIDAdd_Set = S_DocIDAdd
  S_Department_Set = S_Department
  S_Author_Set = S_Author

  if S_DocIDParent<>"" Then
  S_DocIDParent_Set = S_DocIDParent
  Else 
  S_DocIDParent_Set = " "
  End If

  ' Создание карточки документа
  If UCase(Request("create")) = "Y" Then
    If Request("UpdateDoc") <> "YES" Then
      'SAY 2009-02-20
      VAR_ChangeDocGenerateButton=""
      S_DocID = " "
      If S_DocID_Set<>"" Then
        S_DocID_Set = S_DocID
      End If
      
      S_DocID_Set = " "
      S_Author_Set = InsertionName(Session("Name"), Session("UserID"))
      S_Department = Session("Department")
      S_Department_Set = S_Department
	  S_NameAproval = " "
      S_ListToReconcile = ""
	  S_ListToReconcile_Set = ""
	  S_UserFieldText2 = ""
    End If
  End If 'UCase(Request("create")) = "Y"

 

'ph - 20120216 - start
  CurrentDocFieldOrder = "DocID, DocIDParent, Author, PartnerName, UserFieldText1, UserFieldText2, UserFieldText3, ListToReconcile, NameAproval, Content, UserFieldText4, UserFieldText5, UserFieldMoney1, ListToView, SecurityLevel, Department"
'ph - 20120216 - end
  CurrentDocRequiredFields = "PartnerName, UserFieldText2,NameAproval,Content"
  'Ph - 20080921 - Убрать иконку справочников у полей
  VAR_DirPictNotToShow = "UserFieldText1,UserFieldText2,UserFieldText3,UserFieldText4,UserFieldMoney1,UserFieldText5"
  
    'Добавим комментарий по обязательным согласующим
  S_ListToReconcile_Comment = SIT_RequiredAgrees + vbCrLf + oPayDox.GetExtTableValue("AgreeRTI"+Request("l"),"Category",Session("CurrentClassDoc"),"List")+vbcrlf+"""Гартван К. Р."" <gartvan_oaorti>;"+ +vbcrlf+RTI_DirectorOfSecurity+RTI_HeadOfUpravDelami+RTI_HeadOfAccounting+RTI_HeadKFIE+RTI_HeadOfPurchaseCenter+RTI_HeadPriceforming+RTI_BudgetController+vbcrlf+SIT_RTI_DirectorPravovogoUprav_RU


    VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", Name, ContractType, BusinessUnit"
    VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", AddFieldText1, AddFieldText2"    
	S_ListToReconcile = DeleteConstPrefixFromList(S_ListToReconcile)
'rti_contract  
  

' ********************************* ДОГОВОРЫ NEW ...
' *** ДОГОВОРЫ
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) = 1 Then
  'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
  S_SecurityLevel = 4
  If not IsAdmin() Then
    S_SecurityLevel_Set = S_SecurityLevel
  End If

  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
    S_AdditionalUsers = ""
  End If 

  If S_AdditionalUsers <> "" Then
    S_AdditionalUsers_Set = S_AdditionalUsers
  Else
    S_AdditionalUsers_Set = " "
  End If

  S_DocID_Set = S_DocID
  S_DocIDAdd_Set = S_DocIDAdd
  S_Department_Set = S_Department
  S_Author_Set = S_Author

  ' Создание карточки документа
  If UCase(Request("create")) = "Y" Then
    If Request("UpdateDoc") <> "YES" Then
      'SAY 2009-02-20
      VAR_ChangeDocGenerateButton=""
      S_DocID = " "
      If S_DocID_Set<>"" Then
        S_DocID_Set = S_DocID
      End If

'Запрос №43 - СТС - start
      ''SAY 2008-09-22 для подчиненных документов оставляем весь список согласования родителя
      'If Trim(S_DocIDParent)<>"" then
      '  S_ListToReconcile = Replace(S_ListToReconcile,"##;","")
      'End If
'Запрос №43 - СТС - end

      S_DocID_Set = " "
      S_Author_Set = InsertionName(Session("Name"), Session("UserID"))
      S_Department = Session("Department")
      S_Department_Set = S_Department
'Запрос №35 - СТС - start
'Запрос №43 - СТС - start
		  'S_NameAproval = STS_GenDirector
		  S_NameAproval = " "
'Запрос №43 - СТС - end
'Запрос №43 - СТС - start
'		  S_NameAproval_Set = ""
		  S_NameAproval_Set = S_NameAproval
'Запрос №43 - СТС - end
			S_ListToReconcile = ""
			S_ListToReconcile_Set = ""
'Запрос №35 - СТС - end
    End If
  End If 'UCase(Request("create")) = "Y"

'ph - 20120216 - start
'  CurrentDocFieldOrder = "DocID, Author, NameResponsible, UserFieldText4, PartnerName, UserFieldText1, Description, Name, UserFieldText8, UserFieldText2, InventoryUnit, AmountDoc, Currency, UserFieldText5, UserFieldDate4, UserFieldDate6, ListToReconcile, NameAproval, SecurityLevel, Department, UserFieldText3, DocIDParent, UserFieldDate5, ContractType "	
  CurrentDocFieldOrder = "DocID, Author, NameResponsible, UserFieldText4, PartnerName, UserFieldText1, Description, Name, UserFieldText8, UserFieldText2, InventoryUnit, AmountDoc, Currency, UserFieldText5, UserFieldDate4, UserFieldDate6, ListToReconcile, NameAproval, SecurityLevel, Department, UserFieldText3, DocIDParent, UserFieldDate5, ContractType, UserFieldText6, UserFieldText7, AddFieldText1"
'ph - 20120216 - end
  CurrentDocRequiredFields = "Description, AmountDoc, Currency"
  If UCase(Request("create")) = "Y" Then
    CurrentDocRequiredFields = CurrentDocRequiredFields & ", PartnerName, UserFieldText8, UserFieldText3"
    'Если пользователь из СТС, то обязательным полем сделать и поле ContractType
    If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
      CurrentDocRequiredFields = CurrentDocRequiredFields & ", ContractType"
    End If
  End If
  'Название проекта
  If S_UserFieldText2 = "" Then
    S_UserFieldText2 = " "
  End If
  S_UserFieldText2_Set = S_UserFieldText2
  'Ph - 20080921 - Убрать иконку справочников у полей
  VAR_DirPictNotToShow = "UserFieldText4"
'Запрос №35 - СТС - start
  If not IsAdmin() Then
'Запрос №43 - СТС - start
'    sUsersBUs = GetUsersField(Session("UserID"), "BusinessUnits")
''ph - 20101227 - start
'    'Для BU 2090 меняем подписанта
'    If InStr(sUsersBUs, "2090") > 0 Then
'      S_NameAproval = """Красовский А. В."" <akrasovski>;"
'    End If
'ph - 20101227 - end
    'Всем запрещено, переводим на правила
    S_NameAproval_Set = S_NameAproval
'    'Пользователям не из указанных БЕ запрещено править утверждающего
'    If InStr(sUsersBUs, "1010") = 0 and InStr(sUsersBUs, "2090") = 0 Then
'      S_NameAproval_Set = S_NameAproval
'    End If
'    'В активных документах не администраторам нельзя править список согласования
'    If UCase(S_IsActive) = "Y" Then
'      S_ListToReconcile_Set = S_ListToReconcile
'      'Даже если он пустой
'      If S_ListToReconcile_Set = "" Then
'        S_ListToReconcile = " "
'        S_ListToReconcile_Set = " "
'      End If
'    End If
'Запрос №43 - СТС - end
  End If

  If InStr(UCase(Session("Department")), UCase(SIT_STS)) <> 1 Then 
    VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", ContractType"
'ph - 20120216 - start
    VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", AddFieldText1"
'ph - 20120216 - end
  Else
    VAR_DirPictNotToShow = VAR_DirPictNotToShow+", ContractType"
  End If

  'rmanyushin 29.12.2009 start - С ДОБАВЛЕНИЯМИ 'При редактировании документа нельзя поменять значение поля ContractType 
  If Request("create") = "y" Then
    S_AddField_Select3 = GetUserDirValues(STS_ContractPartyDirGUID) 
    S_UserFieldText3_Select = STS_ContractPaymentDirection_In&","&STS_ContractPaymentDirection_Out&","&STS_ContractPaymentDirection_Free
    	'{ ph - 20120601 - При создании как подчиненного из другой категории, может лезть значение не содержащееся в списке допустимых
	If S_UserFieldText3 <> "" Then
		If InStr(S_UserFieldText3_Select & ",", S_UserFieldText3 & ",") = 0 Then
			S_UserFieldText3 = ""
		End If
	End If
	' ph - 20120601 }

  Else
    S_AddField_Set3 = Request("ContractType")
    If S_AddField3 <> "" And bUserCheckOK Then
      S_AddField_Set3 = S_AddField3
    ElseIf IsObject(dsDoc) And bUserCheckOK And UCase(Request("empty"))<>"Y" Then
	  If Not dsDoc.EOF Then
        S_AddField_Set3 = MyCStr(dsDoc("ContractType"))
	  End If
    End If
    'также нельзя менять поля, влияющие на нумерацию - Контрагент, Проект, Доходный/расходный
    S_PartnerName_Set = S_PartnerName
    S_UserFieldText3_Set = S_UserFieldText3
    S_UserFieldText8_Set = S_UserFieldText8
  End If
  'rmanyushin 29.12.2009 end   

  Select Case UCase(Request("l"))
    Case "RU" 'RU
      S_Name_Select = GetUserDirValues("{535C7740-CB1E-4403-ABC6-93AFC67205D5}")
    Case "" 'EN
      S_Name_Select = GetUserDirValues("{653D0ABF-6E6E-4DCA-899F-51D44A3DB2C5}")
    Case "3" 'CZ
      S_Name_Select = GetUserDirValues("{E30E310A-67DC-4BE2-8023-61F3214EAB4D}")
  End Select

  'Для пользователей СТС везде добавляем поле бизнес единица
  If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
    If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 Then
      If UCase(Request("create")) = "Y" Then
        sCreator = Session("UserID")
      Else
        sCreator = GetUserID(S_Author)
      End If
      BusinessUnitsList = GetUsersBusinessUnits(sCreator)
      If InStr(BusinessUnitsList, VbCrLf) > 0 Then
        S_AddField2 = " "
        S_AddField_Select2 = BusinessUnitsList
        'Поле обязательно для заполнения
        CurrentDocRequiredFields = CurrentDocRequiredFields+", BusinessUnit"
      Else
        S_AddField_Set2 = BusinessUnitsList
      End If
    End If
  Else
    'Поле показывается только для пользователей СТС и администратора (у админа не из СТС оно необязательное, редактируемое)
    If not IsAdmin() Then
      VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", BusinessUnit"
    End If
  End If

  'Заполнение списка согласования делаем после разбирательства с BU, т.к. на них опираемся при выборе Россия/Чехия
  If Request("create") = "y" Then
    If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
'Запрос №43 - СТС - start
	   'Для СТС работают правила
      If Request("UpdateDoc") <> "YES" Then
        S_ListToReconcile_Comment = ""
        S_ListToReconcile = ""
        S_ListToReconcile_Set = ""
      End If
'      sCountry = GetLangByBusinessUnits(BusinessUnitsList)
'      If sCountry = "?" Then
'        S_ListToReconcile_Comment = "?" 'Не удалось на этой стадии определить список обязательных согласующих, будет определен после выбора БЕ
'      Else
'        S_ListToReconcile_Comment = oPayDox.GetExtTableValue("AgreeSTSRU"&sCountry, "Category", Session("CurrentClassDoc")&"/", "List")
'      End If
'      If Request("UpdateDoc") <> "YES" Then
'        S_ListToReconcile = ""
'        S_ListToReconcile_Set = ""
'      End If
'Запрос №43 - СТС - end
    Else
      'Для остальных БН не определена БЕ, т.ч. привязка языка по интерфейсу
      S_ListToReconcile_Comment = oPayDox.GetExtTableValue("Agree"&Request("l"), "Category", Session("CurrentClassDoc")&"/", "List")
    End If
  Else
'Запрос №43 - СТС - start
    If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
	  S_ListToReconcile_Comment = ""
	Else
'Запрос №43 - СТС - end
		If Request("UpdateDoc") = "YES" or not bUserCheckOK Then
		  S_ListToReconcile_Comment = Session("S_ListToReconcile_Comment")
		Else
		  If InStr(S_ListToReconcile, SIT_AdditionalAgreesDelimeter) > 0 Then
			S_ListToReconcile_Comment = Left(S_ListToReconcile, InStr(S_ListToReconcile, SIT_AdditionalAgreesDelimeter)-1)
		  Else
			S_ListToReconcile_Comment = ""
		  End If
		  Session("S_ListToReconcile_Comment") = S_ListToReconcile_Comment
		End If
		S_ListToReconcile = DeleteConstPrefixFromList(S_ListToReconcile)
'Запрос №43 - СТС - start
	End If
'Запрос №43 - СТС - end
  End If
'Запрос №43 - СТС - start
  'Поле Договор внутри группы СТС
  S_UserFieldText6_Select = SIT_YesNo
  'Делаем обязательным
  CurrentDocRequiredFields = CurrentDocRequiredFields+", UserFieldText6"
'Запрос №43 - СТС - end

'20120203 - start
  'Поле Взаимодействие сторон договора
  S_UserFieldText7_Select = SIT_YesNo
'20120203 - end
'ph - 20120216 - start
  S_AddField_Select4 = SIT_YesNo
'ph - 20120216 - end

' *** КОММЕРЧЕСКИЕ ПРЕДЛОЖЕНИЯ
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_COM_OFFERS)) = 1 Then
  'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
  S_SecurityLevel = 4
  If not IsAdmin() Then
    S_SecurityLevel_Set = S_SecurityLevel
  End If

  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
    S_AdditionalUsers = ""
  End If

  If S_AdditionalUsers <> "" Then
    S_AdditionalUsers_Set = S_AdditionalUsers
  Else
    S_AdditionalUsers_Set = " "
  End If

  S_DocID_Set = S_DocID
  S_DocIDAdd_Set = S_DocIDAdd
  S_Department_Set = S_Department
  S_Author_Set = S_Author

  ' Создание карточки документа
  If UCase(Request("create")) = "Y" Then
    If Request("UpdateDoc") <> "YES" Then
      'SAY 2009-02-20
      VAR_ChangeDocGenerateButton=""
      S_DocID = " "
      If S_DocID_Set<>"" Then
        S_DocID_Set = S_DocID
      End If

      'SAY 2008-11-06
      S_Department_Set = Session("Department")

      'SAY 2008-09-22 для подчиненных документов оставляем весь список согласования родителя
      If Trim(S_DocIDParent)<>"" then
        S_ListToReconcile = Replace(S_ListToReconcile,"##;","")
      End If

      S_DocID_Set = " "
      S_DateActivation_Set = MyDate(Now)
      S_Author_Set = InsertionName(Session("Name"), Session("UserID"))
      S_Department_Set = Session("Department")
      'Очищаем поля, чтобы не были заполнены при создании как подчиненного из документа другой категории
      If UCase(Request("ClassDoc")) <> UCase(Session("CurrentClassDoc")) Then
        S_Name = ""
        S_UserFieldText1 = ""
        S_UserFieldText2 = ""
        S_Content = ""
        S_ListToReconcile = ""
        S_NameAproval = ""
        S_PartnerName = ""
        S_ListToView = ""
      End If
    End If
  End If 'UCase(Request("create")) = "Y"

  S_ListToReconcile = DeleteConstPrefixFromList(S_ListToReconcile)

  'Порядок следования полей при редактировании
  CurrentDocFieldOrder = "Name, DocIDParent, DateActivation, DocID, UserFieldText1, UserFieldText2, Author, Content, ListToReconcile, NameAproval, PartnerName, ListToView, SecurityLevel, Department, "
  'Обязательные поля (нередактируемые здесь не нужны)
  CurrentDocRequiredFields = "Name, NameAproval, PartnerName, ListToView"
  'Отмена показа справочников для полей
  VAR_DirPictNotToShow = "UserFieldText1, UserFieldText2"

  VAR_DirPictNotToShow = VAR_DirPictNotToShow+", DocName"
  'rmanyushin 60298 02.11.2009 Start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", ContractType"
  'rmanyushin 60298 02.11.2009 End
'ph - 20120216 - start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", AddFieldText1"
'ph - 20120216 - end

  'Для пользователей СТС везде добавляем поле бизнес единица
  If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
    If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 Then
      If UCase(Request("create")) = "Y" Then
        sCreator = Session("UserID")
      Else
        sCreator = GetUserID(S_Author)
      End If
      BusinessUnitsList = GetUsersBusinessUnits(sCreator)
      If InStr(BusinessUnitsList, VbCrLf) > 0 Then
        S_AddField2 = " "
        S_AddField_Select2 = BusinessUnitsList
        'Поле обязательно для заполнения
        CurrentDocRequiredFields = CurrentDocRequiredFields+", BusinessUnit"
      Else
        S_AddField_Set2 = BusinessUnitsList
      End If
    End If
  Else
    'Поле показывается только для пользователей СТС и администратора (у админа не из СТС оно необязательное, редактируемое)
    If not IsAdmin() Then
      VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", BusinessUnit"
    End If
  End If

' *** HELPDESK
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_HelpDesk)) = 1 Then 
  'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
  S_SecurityLevel = 4
  If not IsAdmin() Then
    S_SecurityLevel_Set = S_SecurityLevel
  End If

  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
    S_AdditionalUsers = ""
  End If

  If S_AdditionalUsers <> "" Then
    S_AdditionalUsers_Set = S_AdditionalUsers
  Else
    S_AdditionalUsers_Set = " "
  End If

  S_DocID_Set = S_DocID
  S_DocIDAdd_Set = S_DocIDAdd
  S_Department_Set = S_Department
  S_Author_Set = S_Author

  ' Создание карточки документа
  If UCase(Request("create")) = "Y" Then
    If Request("UpdateDoc") <> "YES" Then
      'SAY 2009-02-20
      VAR_ChangeDocGenerateButton=""
      S_DocID = " "
      If S_DocID_Set<>"" Then
        S_DocID_Set = S_DocID
      End If

      'SAY 2008-11-06
      S_Department_Set = Session("Department")

      'SAY 2008-09-22 для подчиненных документов оставляем весь список согласования родителя
      If Trim(S_DocIDParent)<>"" then
        S_ListToReconcile = Replace(S_ListToReconcile,"##;","")
      End If

      S_DocID = ""
      S_DocID_Set = " "
      'S_DateActivation = MyDate(Date)
      'If not IsAdmin() Then
      '  S_DateActivation_Set = S_DateActivation
      'End If
      S_Department_Set = Session("Department")
    End If
  End If 'UCase(Request("create")) = "Y"

  S_ListToReconcile = DeleteConstPrefixFromList(S_ListToReconcile)

  VAR_DirPictNotToShow = VAR_DirPictNotToShow+", DocName"
  'rmanyushin 60298 02.11.2009 Start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", ContractType"
  'rmanyushin 60298 02.11.2009 End
'ph - 20120216 - start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", AddFieldText1"
'ph - 20120216 - end

  'Для пользователей СТС везде добавляем поле бизнес единица
  If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
    If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 Then
      If UCase(Request("create")) = "Y" Then
        sCreator = Session("UserID")
      Else
        sCreator = GetUserID(S_Author)
      End If
      BusinessUnitsList = GetUsersBusinessUnits(sCreator)
      If InStr(BusinessUnitsList, VbCrLf) > 0 Then
        S_AddField2 = " "
        S_AddField_Select2 = BusinessUnitsList
        'Поле обязательно для заполнения
        CurrentDocRequiredFields = CurrentDocRequiredFields+", BusinessUnit"
      Else
        S_AddField_Set2 = BusinessUnitsList
      End If
    End If
  Else
    'Поле показывается только для пользователей СТС и администратора (у админа не из СТС оно необязательное, редактируемое)
    If not IsAdmin() Then
      VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", BusinessUnit"
    End If
  End If

' *** ЗАЯВКИ НА ЗАКУПКУ СТС
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) = 1 Then
  'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
  S_SecurityLevel = 4
  If not IsAdmin() Then
    S_SecurityLevel_Set = S_SecurityLevel
  End If

  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
    S_AdditionalUsers = ""
  End If

  If S_AdditionalUsers <> "" Then
    S_AdditionalUsers_Set = S_AdditionalUsers
  Else
    S_AdditionalUsers_Set = " "
  End If

  S_DocID_Set = S_DocID
  S_DocIDAdd_Set = S_DocIDAdd
  S_Department_Set = S_Department
  S_Author_Set = S_Author

  If Request("create") = "y" and Request("UpdateDoc") <> "YES" Then
    'SAY 2009-02-20
    VAR_ChangeDocGenerateButton=""
    S_DocID = " "
    If S_DocID_Set<>"" Then
      S_DocID_Set = S_DocID
    End If

    'SAY 2008-11-06
    S_Department_Set = Session("Department")

    'SAY 2008-09-22 для подчиненных документов оставляем весь список согласования родителя
    If Trim(S_DocIDParent)<>"" then
      S_ListToReconcile = Replace(S_ListToReconcile,"##;","")
    End If

    S_DocID = " "
    S_DocID_Set = S_DocID
    S_Author = InsertionName(Session("Name"), Session("UserID"))
    S_Department = Session("Department")
    S_UserFieldText4 = " "
    S_ListToReconcile = ""
    S_NameAproval = " "
    S_UserFieldMoney1 = "0"
  End If

  S_ListToReconcile = DeleteConstPrefixFromList(S_ListToReconcile)

  S_UserFieldText4_Set = S_UserFieldText4
'Запрос №36 - СТС - start - Список согласования редактируемый, дополняется правилами при сохранении
'  S_ListToReconcile_Set = S_ListToReconcile
  S_ListToReconcile_Set = ""
'Запрос №36 - СТС - end
  S_NameAproval_Set = S_NameAproval
  S_Department_Set = S_Department
  S_Author_Set = S_Author
  S_NameUserFieldMoney1 = "" 'прячем поле (всегда)
  'Список выбора для наличия в бюджете
  S_UserFieldText6_Select = "Yes,No"
  CurrentDocFieldOrder = "DocID, Author, NameResponsible, PartnerName, Description, Content, AmountDoc, Currency, ListToReconcile, NameAproval, Department, UserFieldText1, BusinessUnit, UserFieldText3, UserFieldText4, UserFieldText5, UserFieldText6, UserFieldText8, DateCompletion"
  VAR_DirPictNotToShow = "DocAmountDoc, DocDescription"
'Запрос №30 - Start
  CurrentDocRequiredFields = "Description, AmountDoc, Currency, UserFieldText1, BusinessUnit, UserFieldText3, UserFieldText5, UserFieldText6, UserFieldText8, DateCompletion"
  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
'Запрос №36 - СТС - start - Теперь ставится правилами
'    S_NameResponsible = STS_Purchase_Logistics_Department
    S_NameResponsible = " "
  End If
  S_NameResponsible_Set = S_NameResponsible
'Запрос №36 - СТС - end
'Запрос №30 - End

  VAR_DirPictNotToShow = VAR_DirPictNotToShow+", DocName"
  'rmanyushin 60298 02.11.2009 Start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", ContractType"
  'rmanyushin 60298 02.11.2009 End
'ph - 20120216 - start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", AddFieldText1"
'ph - 20120216 - end

  'Для пользователей СТС везде добавляем поле бизнес единица
  If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
    If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 Then
      If UCase(Request("create")) = "Y" Then
        sCreator = Session("UserID")
      Else
        sCreator = GetUserID(S_Author)
      End If
      BusinessUnitsList = GetUsersBusinessUnits(sCreator)
      If InStr(BusinessUnitsList, VbCrLf) > 0 Then
        S_AddField2 = " "
        S_AddField_Select2 = BusinessUnitsList
        'Поле обязательно для заполнения
        CurrentDocRequiredFields = CurrentDocRequiredFields+", BusinessUnit"
      Else
        S_AddField_Set2 = BusinessUnitsList
      End If
    End If
  Else
    'Поле показывается только для пользователей СТС и администратора (у админа не из СТС оно необязательное, редактируемое)
    If not IsAdmin() Then
      VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", BusinessUnit"
    End If
  End If

' *** ЗАЯВКИ НА ОПЛАТУ СТС
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) = 1 Then
  'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
  S_SecurityLevel = 4
  If not IsAdmin() Then
    S_SecurityLevel_Set = S_SecurityLevel
  End If

  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
    S_AdditionalUsers = ""
  End If

  If S_AdditionalUsers <> "" Then
    S_AdditionalUsers_Set = S_AdditionalUsers
  Else
    S_AdditionalUsers_Set = " "
  End If

  S_DocID_Set = S_DocID
  S_DocIDAdd_Set = S_DocIDAdd
  S_Department_Set = S_Department
  S_Author_Set = S_Author

  If Request("create") = "y" and Request("UpdateDoc") <> "YES" Then
    'SAY 2009-02-20
    VAR_ChangeDocGenerateButton=""
    S_DocID = " "
    If S_DocID_Set<>"" Then
      S_DocID_Set = S_DocID
    End If

    'SAY 2008-11-06
    S_Department_Set = Session("Department")

    'SAY 2008-09-22 для подчиненных документов оставляем весь список согласования родителя
    If Trim(S_DocIDParent)<>"" then
      S_ListToReconcile = Replace(S_ListToReconcile,"##;","")
    End If

    S_DocID = " "
    S_DocID_Set = S_DocID
    S_Author = InsertionName(Session("Name"), Session("UserID"))
    'rmanyushin 105583 25.10.2010 Start
    'S_NameResponsible = STS_Orders_Accounting
    S_NameResponsible = STS_Orders_ResponsiblePayO
    'rmanyushin 105583 25.10.2010 End

    S_Department = Session("Department")
    S_ListToReconcile = " "
    S_NameAproval = " "
    S_UserFieldMoney1 = "0"
    If Trim(S_UserFieldText1) = "" Then
      S_UserFieldText1 = " "
    End If
    If S_AddField2 <> "" and bUserCheckOK Then
    ElseIf not dsDoc.EOF and bUserCheckOK and UCase(Request("empty")) <> "Y" Then
      S_AddField2 = TDValue(dsDoc("BusinessUnit"))
    Else
      S_AddField2 = Request("BusinessUnit")
    End If
    If Trim(S_AddField2) = "" Then
      S_AddField2 = " "
    End If
    If Trim(S_UserFieldText3) = "" Then
      S_UserFieldText3 = " "
    End If
    If Trim(S_UserFieldText4) = "" Then
      S_UserFieldText4 = " "
    End If
    If Trim(S_UserFieldText6) = "" Then
      S_UserFieldText6 = " "
    End If
    If Trim(S_UserFieldText8) = "" Then
      S_UserFieldText8 = " "
    End If
  End If

  S_ListToReconcile = DeleteConstPrefixFromList(S_ListToReconcile)

  CurrentDocFieldOrder = "DocID, DocIDParent, Author, NameResponsible, PartnerName, Description, AmountDoc, Currency, ListToReconcile, NameAproval, Department, UserFieldText1, BusinessUnit, UserFieldText3, UserFieldText4, UserFieldText5, UserFieldText6, UserFieldText8, DateCompletion"
  CurrentDocRequiredFields = "DocIDParent, AmountDoc, Currency, DateCompletion, Description, UserFieldText5, PartnerName"
  VAR_DirPictNotToShow = "DocAmountDoc, DocDescription"

  S_NameUserFieldMoney1 = "" 'прячем поле
  S_ListToReconcile_Set = S_ListToReconcile
  S_NameAproval_Set = S_NameAproval
  S_Department_Set = S_Department
  S_NameResponsible_Set = S_NameResponsible
  S_Author_Set = S_Author
  S_UserFieldText1_Set = S_UserFieldText1
  S_AddField_Set2 = S_AddField2
  S_UserFieldText3_Set = S_UserFieldText3
  S_UserFieldText4_Set = S_UserFieldText4
  S_UserFieldText6_Set = S_UserFieldText6
  S_UserFieldText8_Set = S_UserFieldText8

  If Trim(S_DocIDParent) <> "" Then
    If Request("create") = "y" and Request("UpdateDoc")<>"YES" Then
      Set dsTemp1 = Server.CreateObject("ADODB.Recordset")
      sSQL = "select * from Docs where DocID = "+sUnicodeSymbol+"'" + S_DocIDParent + "'"
AddLogD "@@@SearchingParentDoc SQL: "+sSQL
      dsTemp1.Open sSQL, Conn, 3, 1, &H1
      If Not dsTemp1.EOF Then
        nPurchaseOrderAmountUSD = CCur(dsTemp1("UserFieldMoney1"))
        dsTemp1.Close
        sSQL = "select IsNull(Sum(UserFieldMoney1), 0) as SumUSD from Docs where DocIDParent = "+sUnicodeSymbol+"'" + S_DocIDParent + "' and ClassDoc like "+sUnicodeSymbol+"'"+STS_PaymentOrder+"%'"
AddLogD "@@@SumOfChildPaymentOrders SQL: "+sSQL
        dsTemp1.Open sSQL, Conn, 3, 1, &H1
        nPaymentOrdersSumUSD = CCur(dsTemp1("SumUSD"))
        S_AmountDoc = CStr(Round((nPurchaseOrderAmountUSD-nPaymentOrdersSumUSD)*CurrencyConvertionFactor("USD", S_Currency), 2))
      End If
      dsTemp1.Close
    End If
  End If

  VAR_DirPictNotToShow = VAR_DirPictNotToShow+", DocName"
  'rmanyushin 60298 02.11.2009 Start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", ContractType"
  'rmanyushin 60298 02.11.2009 End
'ph - 20120216 - start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", AddFieldText1"
'ph - 20120216 - end

  'Для пользователей СТС везде добавляем поле бизнес единица
  If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
    If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 Then
      If UCase(Request("create")) = "Y" Then
        sCreator = Session("UserID")
      Else
        sCreator = GetUserID(S_Author)
      End If
      BusinessUnitsList = GetUsersBusinessUnits(sCreator)
      If InStr(BusinessUnitsList, VbCrLf) > 0 Then
        S_AddField2 = " "
        S_AddField_Select2 = BusinessUnitsList
        'Поле обязательно для заполнения
        CurrentDocRequiredFields = CurrentDocRequiredFields+", BusinessUnit"
      Else
        S_AddField_Set2 = BusinessUnitsList
      End If
    End If
  Else
    'Поле показывается только для пользователей СТС и администратора (у админа не из СТС оно необязательное, редактируемое)
    If not IsAdmin() Then
      VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", BusinessUnit"
    End If
  End If

' *** ПОДТИПЫ СЛУЖЕБНЫХ ЗАПИСОК
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_SLUZH_ZAPISKA_COMPUTER)) = 1 Then
  'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
  S_SecurityLevel = 4
  If not IsAdmin() Then
    S_SecurityLevel_Set = S_SecurityLevel
  End If

  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
    S_AdditionalUsers = ""
  End If

  If S_AdditionalUsers <> "" Then
    S_AdditionalUsers_Set = S_AdditionalUsers
  Else
    S_AdditionalUsers_Set = " "
  End If

  S_DocID_Set = S_DocID
  S_DocIDAdd_Set = S_DocIDAdd
  S_Department_Set = S_Department
  S_Author_Set = S_Author

  ' Создание карточки документа
  If UCase(Request("create")) = "Y" Then
    If Request("UpdateDoc") <> "YES" Then
      'SAY 2009-02-20
      VAR_ChangeDocGenerateButton=""
      S_DocID = " "
      If S_DocID_Set<>"" Then
        S_DocID_Set = S_DocID
      End If

      'SAY 2008-11-06
      S_Department_Set = Session("Department")

      'SAY 2008-09-22 для подчиненных документов оставляем весь список согласования родителя
      If Trim(S_DocIDParent)<>"" then
        S_ListToReconcile = Replace(S_ListToReconcile,"##;","")
      End If

      S_DocID_Set = " "
      S_DateActivation = MyDate(Date)
      S_UserFieldDate1 = MyDate(Date)
      S_Author_Set = InsertionName(Session("Name"), Session("UserID"))
      S_Department_Set = Session("Department")
    End If
  End If 'UCase(Request("create")) = "Y"

  If UCase(Request("create")) = "Y" Then
    If Request("UpdateDoc") <> "YES" Then
'Запрос №1 - СИБ - start
      If InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 Then
        S_NameAproval_Set = SIB_AssistantDirector
        S_ListToView = SIB_HeadOfSectorIS
        S_ListToView_Set = S_ListToView
	  Else
'Запрос №1 - СИБ - end
        S_NameAproval_Set = SIT_VicePresidentOfInitiator
        S_ListToView = SIT_President
        S_ListToView_Set = S_ListToView
'Запрос №1 - СИБ - start
      End If
'Запрос №1 - СИБ - end
    End If
  End If 'UCase(Request("create")) = "Y"

  S_ListToReconcile = DeleteConstPrefixFromList(S_ListToReconcile)

  CurrentDocFieldOrder = "Name, UserFieldText1, DocIDParent, UserFieldDate1, DocID, Author, DateActivation, Content, ListToReconcile, NameAproval, Correspondent, SecurityLevel, Department "
  CurrentDocRequiredFields = "Name, UserFieldText1, UserFieldDate1, DateActivation, NameAproval, ListToView"

'Запрос №1 - СИБ - start
  If InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 Then
'  If InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) = 1 Then
    S_ListToReconcile_Comment = oPayDox.GetExtTableValue("AgreeSIBRU","Category",Session("CurrentClassDoc")+"/","List")
  Else
    S_ListToReconcile_Comment = oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")
  End If
'Запрос №1 - СИБ - end

  If InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) = 1 then
    S_LocationPath = ""
    bLocationPath = ""
  Else
    CurrentDocRequiredFields = CurrentDocRequiredFields+", DocListToRegister" 
  End If

  VAR_DirPictNotToShow = VAR_DirPictNotToShow+", DocName"
  'rmanyushin 60298 02.11.2009 Start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", ContractType"
  'rmanyushin 60298 02.11.2009 End
'ph - 20120216 - start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", AddFieldText1, AddFieldText2"
'ph - 20120216 - end
'Запрос №1 - СИБ - start
  If InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 Then
    VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", UserFieldText1"
    CurrentDocRequiredFields = Replace(CurrentDocRequiredFields, "UserFieldText1", "")
  End If
'Запрос №1 - СИБ - end

'Запрос №1 - СИБ - start
  If InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) <> 1 Then
'Запрос №1 - СИБ - end
    'Выпадающие списки
    Select Case UCase(Request("l"))
      Case "RU" 'RU
        S_UserFieldText1_Select = GetUserDirValues("{BFC71550-2605-4679-8A3F-C04211891D7E}")
      Case "" 'EN
        S_UserFieldText1_Select = GetUserDirValues("{1B136CF3-83CF-471D-925A-EEB72BC6CD5B}")
      Case "3" 'CZ
        S_UserFieldText1_Select = GetUserDirValues("{1EBA180D-7657-4E72-A678-9ECE4EDD58C1}")
    End Select
'Запрос №1 - СИБ - start
  End If
'Запрос №1 - СИБ - end

  'Для пользователей СТС везде добавляем поле бизнес единица
  If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
    If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 Then
      If UCase(Request("create")) = "Y" Then
        sCreator = Session("UserID")
      Else
        sCreator = GetUserID(S_Author)
      End If
      BusinessUnitsList = GetUsersBusinessUnits(sCreator)
      If InStr(BusinessUnitsList, VbCrLf) > 0 Then
        S_AddField2 = " "
        S_AddField_Select2 = BusinessUnitsList
        'Поле обязательно для заполнения
        CurrentDocRequiredFields = CurrentDocRequiredFields+", BusinessUnit"
      Else
        S_AddField_Set2 = BusinessUnitsList
      End If
    End If
  Else
    'Поле показывается только для пользователей СТС и администратора (у админа не из СТС оно необязательное, редактируемое)
    If not IsAdmin() Then
      VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", BusinessUnit"
    End If
  End If

ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_SLUZH_ZAPISKA_MOBILE)) = 1 Then
  'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
  S_SecurityLevel = 4
  If not IsAdmin() Then
    S_SecurityLevel_Set = S_SecurityLevel
  End If

  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
    S_AdditionalUsers = ""
  End If

  If S_AdditionalUsers <> "" Then
    S_AdditionalUsers_Set = S_AdditionalUsers
  Else
    S_AdditionalUsers_Set = " "
  End If

  S_DocID_Set = S_DocID
  S_DocIDAdd_Set = S_DocIDAdd
  S_Department_Set = S_Department
  S_Author_Set = S_Author

  ' Создание карточки документа
  If UCase(Request("create")) = "Y" Then
    If Request("UpdateDoc") <> "YES" Then
      'SAY 2009-02-20
      VAR_ChangeDocGenerateButton=""
      S_DocID = " "
      If S_DocID_Set<>"" Then
        S_DocID_Set = S_DocID
      End If

      'SAY 2008-11-06
      S_Department_Set = Session("Department")

      'SAY 2008-09-22 для подчиненных документов оставляем весь список согласования родителя
      If Trim(S_DocIDParent)<>"" then
        S_ListToReconcile = Replace(S_ListToReconcile,"##;","")
      End If

      S_DocID_Set = " "
      S_DateActivation = MyDate(Date)
      S_UserFieldDate1 = MyDate(Date)
      S_Author_Set = InsertionName(Session("Name"), Session("UserID"))
      S_Department_Set = Session("Department")
    End If
  End If 'UCase(Request("create")) = "Y"

  If UCase(Request("create")) = "Y" Then
    If Request("UpdateDoc") <> "YES" Then
      S_NameAproval_Set = SIT_VicePresidentOfInitiator
      S_ListToView = SIT_FinancialVicePresident
      S_ListToView_Set = S_ListToView
      S_Content = SIT_SLUZH_ZAPISKA_MOBILE_CONTENT
      'отмена справочников для пользовательских полей
      VAR_DirPictNotToShow = "UserFieldText2, UserFieldText4, UserFieldText5, UserFieldText6"
    End If
  End If 'UCase(Request("create")) = "Y"

  S_ListToReconcile = DeleteConstPrefixFromList(S_ListToReconcile)

  CurrentDocFieldOrder = "Name, UserFieldText1, DocIDParent, UserFieldDate1, DocID, Author, DateActivation, Content, ListToReconcile, NameAproval, Correspondent, SecurityLevel, Department "
  CurrentDocRequiredFields = "Name, UserFieldText1, UserFieldDate1, DateActivation, NameAproval, ListToView"

  If InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) = 1 Then
    S_ListToReconcile_Comment = oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")
  End If

  CurrentDocFieldOrder = "Name, UserFieldText1, DocIDParent, UserFieldDate1, DocID, Author, DateActivation, Content, UserFieldText2, UserFieldText3, UserFieldText4, UserFieldText5, UserFieldText6, ListToReconcile, NameAproval, ListToView, SecurityLevel, Department "
  CurrentDocRequiredFields = "Name, UserFieldText1, UserFieldDate1, DateActivation, UserFieldText2, UserFieldText3, UserFieldText4, UserFieldText5, NameAproval, ListToView"
  S_ListToReconcile_Comment = oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")
  'S_UserFieldText3_Select = GetUserDirValues("{b0277016-eb62-41e1-bd5a-960a98f7febc}")

  If InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) = 1 then
    S_LocationPath = ""
    bLocationPath = ""
  Else
    CurrentDocRequiredFields = CurrentDocRequiredFields+", DocListToRegister" 
  End If

  VAR_DirPictNotToShow = VAR_DirPictNotToShow+", DocName"
  'rmanyushin 60298 02.11.2009 Start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", ContractType"
  'rmanyushin 60298 02.11.2009 End
'ph - 20120216 - start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", AddFieldText1"
'ph - 20120216 - end

'Запрос №1 - СИБ - start
  If InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 Then
    VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", UserFieldText1"
    CurrentDocRequiredFields = Replace(CurrentDocRequiredFields, "UserFieldText1", "")
  End If
'Запрос №1 - СИБ - end

'Запрос №1 - СИБ - start
  If InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) <> 1 Then
'Запрос №1 - СИБ - end
    'Выпадающие списки
    Select Case UCase(Request("l"))
      Case "RU" 'RU
        S_UserFieldText1_Select = GetUserDirValues("{BFC71550-2605-4679-8A3F-C04211891D7E}")
      Case "" 'EN
        S_UserFieldText1_Select = GetUserDirValues("{1B136CF3-83CF-471D-925A-EEB72BC6CD5B}")
      Case "3" 'CZ
        S_UserFieldText1_Select = GetUserDirValues("{1EBA180D-7657-4E72-A678-9ECE4EDD58C1}")
    End Select
'Запрос №1 - СИБ - start
  End If
'Запрос №1 - СИБ - end

  'Для пользователей СТС везде добавляем поле бизнес единица
  If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
    If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 Then
      If UCase(Request("create")) = "Y" Then
        sCreator = Session("UserID")
      Else
        sCreator = GetUserID(S_Author)
      End If
      BusinessUnitsList = GetUsersBusinessUnits(sCreator)
      If InStr(BusinessUnitsList, VbCrLf) > 0 Then
        S_AddField2 = " "
        S_AddField_Select2 = BusinessUnitsList
        'Поле обязательно для заполнения
        CurrentDocRequiredFields = CurrentDocRequiredFields+", BusinessUnit"
      Else
        S_AddField_Set2 = BusinessUnitsList
      End If
    End If
  Else
    'Поле показывается только для пользователей СТС и администратора (у админа не из СТС оно необязательное, редактируемое)
    If not IsAdmin() Then
      VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", BusinessUnit"
    End If
  End If

ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_SLUZH_ZAPISKA_KOMANDIROVKA)) = 1 Then
  'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
  S_SecurityLevel = 4
  If not IsAdmin() Then
    S_SecurityLevel_Set = S_SecurityLevel
  End If

  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
    S_AdditionalUsers = ""
  End If

  If S_AdditionalUsers <> "" Then
    S_AdditionalUsers_Set = S_AdditionalUsers
  Else
    S_AdditionalUsers_Set = " "
  End If

  S_DocID_Set = S_DocID
  S_DocIDAdd_Set = S_DocIDAdd
  S_Department_Set = S_Department
  S_Author_Set = S_Author

  ' Создание карточки документа
  If UCase(Request("create")) = "Y" Then
    If Request("UpdateDoc") <> "YES" Then
      'SAY 2009-02-20
      VAR_ChangeDocGenerateButton=""
      S_DocID = " "
      If S_DocID_Set<>"" Then
        S_DocID_Set = S_DocID
      End If

      'SAY 2008-11-06
      S_Department_Set = Session("Department")

      'SAY 2008-09-22 для подчиненных документов оставляем весь список согласования родителя
      If Trim(S_DocIDParent)<>"" then
        S_ListToReconcile = Replace(S_ListToReconcile,"##;","")
      End If

      S_DocID_Set = " "
      S_DateActivation = MyDate(Date)
      S_UserFieldDate1 = MyDate(Date)
      S_Author_Set = InsertionName(Session("Name"), Session("UserID"))
      S_Department_Set = Session("Department")
    End If
  End If 'UCase(Request("create")) = "Y"

  If UCase(Request("create")) = "Y" Then
    If Request("UpdateDoc") <> "YES" Then
'Запрос №1 - СИБ - start
      If InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 Then
        S_NameAproval_Set = SIB_GenDirector
        S_ListToView = SIB_HRManager & VbCrLf & SIB_AssistantDirectorCorpDevelopment & VbCrLf & SIB_HeadAccounting
        S_ListToView_Set = S_ListToView
	  Else
'Запрос №1 - СИБ - end
        S_NameAproval = SIT_VicePresidentOfInitiator
      S_ListToView = SIT_President
        S_ListToView_Set = S_ListToView
'Запрос №1 - СИБ - start
      End If
'Запрос №1 - СИБ - end
    End If
  End If 'UCase(Request("create")) = "Y"

  S_ListToReconcile = DeleteConstPrefixFromList(S_ListToReconcile)

  CurrentDocFieldOrder = "Name, UserFieldText1, DocIDParent, UserFieldDate1, DocID, Author, DateActivation, Content, ListToReconcile, NameAproval, Correspondent, SecurityLevel, Department "
  CurrentDocRequiredFields = "Name, UserFieldText1, UserFieldDate1, DateActivation, NameAproval, ListToView"

'Запрос №1 - СИБ - start
  If InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 Then
'  If InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) = 1 Then
    S_ListToReconcile_Comment = oPayDox.GetExtTableValue("AgreeSIBRU","Category",Session("CurrentClassDoc")+"/","List")
  Else
    S_ListToReconcile_Comment = oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")
  End If
'Запрос №1 - СИБ - end

  CurrentDocFieldOrder = "Name, UserFieldText1, DocIDParent, UserFieldDate1, DocID, Author, DateActivation, Content, UserFieldDate2, UserFieldDate3, ListToReconcile, NameAproval, ListToView, SecurityLevel, Department "
  CurrentDocRequiredFields = "Name, UserFieldText1, UserFieldDate1, DateActivation, UserFieldDate2, UserFieldDate3, NameAproval, ListToView"

  If InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) = 1 then
    S_LocationPath = ""
    bLocationPath = ""
  Else
    CurrentDocRequiredFields = CurrentDocRequiredFields+", DocListToRegister" 
  End If

  VAR_DirPictNotToShow = VAR_DirPictNotToShow+", DocName"
  'rmanyushin 60298 02.11.2009 Start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", ContractType"
  'rmanyushin 60298 02.11.2009 End
'ph - 20120216 - start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", AddFieldText1"
'ph - 20120216 - end

'Запрос №1 - СИБ - start
  If InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 Then
    VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", UserFieldText1"
    CurrentDocRequiredFields = Replace(CurrentDocRequiredFields, "UserFieldText1", "")
  End If
'Запрос №1 - СИБ - end

'Запрос №1 - СИБ - start
  If InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) <> 1 Then
'Запрос №1 - СИБ - end
    'Выпадающие списки
    Select Case UCase(Request("l"))
      Case "RU" 'RU
        S_UserFieldText1_Select = GetUserDirValues("{BFC71550-2605-4679-8A3F-C04211891D7E}")
      Case "" 'EN
        S_UserFieldText1_Select = GetUserDirValues("{1B136CF3-83CF-471D-925A-EEB72BC6CD5B}")
      Case "3" 'CZ
        S_UserFieldText1_Select = GetUserDirValues("{1EBA180D-7657-4E72-A678-9ECE4EDD58C1}")
    End Select
'Запрос №1 - СИБ - start
  End If
'Запрос №1 - СИБ - end

  'Для пользователей СТС везде добавляем поле бизнес единица
  If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
    If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 Then
      If UCase(Request("create")) = "Y" Then
        sCreator = Session("UserID")
      Else
        sCreator = GetUserID(S_Author)
      End If
      BusinessUnitsList = GetUsersBusinessUnits(sCreator)
      If InStr(BusinessUnitsList, VbCrLf) > 0 Then
        S_AddField2 = " "
        S_AddField_Select2 = BusinessUnitsList
        'Поле обязательно для заполнения
        CurrentDocRequiredFields = CurrentDocRequiredFields+", BusinessUnit"
      Else
        S_AddField_Set2 = BusinessUnitsList
      End If
    End If
  Else
    'Поле показывается только для пользователей СТС и администратора (у админа не из СТС оно необязательное, редактируемое)
    If not IsAdmin() Then
      VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", BusinessUnit"
    End If
  End If

ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_SLUZH_ZAPISKA_OBUCHENIE)) = 1 Then
  'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
  S_SecurityLevel = 4
  If not IsAdmin() Then
    S_SecurityLevel_Set = S_SecurityLevel
  End If

  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
    S_AdditionalUsers = ""
  End If

  If S_AdditionalUsers <> "" Then
    S_AdditionalUsers_Set = S_AdditionalUsers
  Else
    S_AdditionalUsers_Set = " "
  End If

  S_DocID_Set = S_DocID
  S_DocIDAdd_Set = S_DocIDAdd
  S_Department_Set = S_Department
  S_Author_Set = S_Author

  ' Создание карточки документа
  If UCase(Request("create")) = "Y" Then
    If Request("UpdateDoc") <> "YES" Then
      'SAY 2009-02-20
      VAR_ChangeDocGenerateButton=""
      S_DocID = " "
      If S_DocID_Set<>"" Then
        S_DocID_Set = S_DocID
      End If

      'SAY 2008-11-06
      S_Department_Set = Session("Department")

      'SAY 2008-09-22 для подчиненных документов оставляем весь список согласования родителя
      If Trim(S_DocIDParent)<>"" then
        S_ListToReconcile = Replace(S_ListToReconcile,"##;","")
      End If

      S_DocID_Set = " "
      S_DateActivation = MyDate(Date)
      S_UserFieldDate1 = MyDate(Date)
      S_Author_Set = InsertionName(Session("Name"), Session("UserID"))
      S_Department_Set = Session("Department")
    End If
  End If 'UCase(Request("create")) = "Y"

  If UCase(Request("create")) = "Y" Then
    If Request("UpdateDoc") <> "YES" Then
'Запрос №1 - СИБ - start
      If InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 Then
        S_NameAproval_Set = SIB_AssistantDirector
        S_ListToView = SIB_HRManager
        S_ListToView_Set = S_ListToView
	  Else
'Запрос №1 - СИБ - end
        S_NameAproval_Set = SIT_VicePresidentOfInitiator
        S_ListToView = SIT_HRDirector
        S_ListToView_Set = S_ListToView
'Запрос №1 - СИБ - start
      End If
'Запрос №1 - СИБ - end
    End If
  End If 'UCase(Request("create")) = "Y"

  S_ListToReconcile = DeleteConstPrefixFromList(S_ListToReconcile)

  CurrentDocFieldOrder = "Name, UserFieldText1, DocIDParent, UserFieldDate1, DocID, Author, DateActivation, Content, ListToReconcile, NameAproval, Correspondent, SecurityLevel, Department "
  CurrentDocRequiredFields = "Name, UserFieldText1, UserFieldDate1, DateActivation, NameAproval, ListToView"

'Запрос №1 - СИБ - start
  If InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 Then
'  If InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) = 1 Then
    S_ListToReconcile_Comment = oPayDox.GetExtTableValue("AgreeSIBRU","Category",Session("CurrentClassDoc")+"/","List")
  Else
    S_ListToReconcile_Comment = oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")
  End If
'Запрос №1 - СИБ - end

  If InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) = 1 then
    S_LocationPath = ""
    bLocationPath = ""
  Else
    CurrentDocRequiredFields = CurrentDocRequiredFields+", DocListToRegister" 
  End If

  VAR_DirPictNotToShow = VAR_DirPictNotToShow+", DocName"
  'rmanyushin 60298 02.11.2009 Start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", ContractType"
  'rmanyushin 60298 02.11.2009 End
'ph - 20120216 - start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", AddFieldText1"
'ph - 20120216 - end

'Запрос №1 - СИБ - start
  If InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 Then
    VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", UserFieldText1"
    CurrentDocRequiredFields = Replace(CurrentDocRequiredFields, "UserFieldText1", "")
  End If
'Запрос №1 - СИБ - end

'Запрос №1 - СИБ - start
  If InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) <> 1 Then
'Запрос №1 - СИБ - end
    'Выпадающие списки
    Select Case UCase(Request("l"))
      Case "RU" 'RU
        S_UserFieldText1_Select = GetUserDirValues("{BFC71550-2605-4679-8A3F-C04211891D7E}")
      Case "" 'EN
        S_UserFieldText1_Select = GetUserDirValues("{1B136CF3-83CF-471D-925A-EEB72BC6CD5B}")
      Case "3" 'CZ
        S_UserFieldText1_Select = GetUserDirValues("{1EBA180D-7657-4E72-A678-9ECE4EDD58C1}")
    End Select
'Запрос №1 - СИБ - start
  End If
'Запрос №1 - СИБ - end

  'Для пользователей СТС везде добавляем поле бизнес единица
  If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
    If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 Then
      If UCase(Request("create")) = "Y" Then
        sCreator = Session("UserID")
      Else
        sCreator = GetUserID(S_Author)
      End If
      BusinessUnitsList = GetUsersBusinessUnits(sCreator)
      If InStr(BusinessUnitsList, VbCrLf) > 0 Then
        S_AddField2 = " "
        S_AddField_Select2 = BusinessUnitsList
        'Поле обязательно для заполнения
        CurrentDocRequiredFields = CurrentDocRequiredFields+", BusinessUnit"
      Else
        S_AddField_Set2 = BusinessUnitsList
      End If
    End If
  Else
    'Поле показывается только для пользователей СТС и администратора (у админа не из СТС оно необязательное, редактируемое)
    If not IsAdmin() Then
      VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", BusinessUnit"
    End If
  End If

ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_SLUZH_ZAPISKA_PERSONAL)) = 1 Then
  'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
  S_SecurityLevel = 4
  If not IsAdmin() Then
    S_SecurityLevel_Set = S_SecurityLevel
  End If

  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
    S_AdditionalUsers = ""
  End If

  If S_AdditionalUsers <> "" Then
    S_AdditionalUsers_Set = S_AdditionalUsers
  Else
    S_AdditionalUsers_Set = " "
  End If

  S_DocID_Set = S_DocID
  S_DocIDAdd_Set = S_DocIDAdd
  S_Department_Set = S_Department
  S_Author_Set = S_Author

  ' Создание карточки документа
  If UCase(Request("create")) = "Y" Then
    If Request("UpdateDoc") <> "YES" Then
      'SAY 2009-02-20
      VAR_ChangeDocGenerateButton=""
      S_DocID = " "
      If S_DocID_Set<>"" Then
        S_DocID_Set = S_DocID
      End If

      'SAY 2008-11-06
      S_Department_Set = Session("Department")

      'SAY 2008-09-22 для подчиненных документов оставляем весь список согласования родителя
      If Trim(S_DocIDParent)<>"" then
        S_ListToReconcile = Replace(S_ListToReconcile,"##;","")
      End If

      S_DocID_Set = " "
      S_DateActivation = MyDate(Date)
      S_UserFieldDate1 = MyDate(Date)
      S_Author_Set = InsertionName(Session("Name"), Session("UserID"))
      S_Department_Set = Session("Department")
    End If
  End If 'UCase(Request("create")) = "Y"

  If UCase(Request("create")) = "Y" Then
    If Request("UpdateDoc") <> "YES" Then
'Запрос №1 - СИБ - start
      If InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 Then
        S_NameAproval_Set = SIB_AssistantDirector
        S_ListToView = SIB_HRManager
        S_ListToView_Set = S_ListToView
	  Else
'Запрос №1 - СИБ - end
        S_NameAproval_Set = SIT_VicePresidentOfInitiator
        S_ListToView = SIT_HRDirector
        S_ListToView_Set = S_ListToView
'Запрос №1 - СИБ - start
      End If
'Запрос №1 - СИБ - end
    End If
  End If 'UCase(Request("create")) = "Y"

  S_ListToReconcile = DeleteConstPrefixFromList(S_ListToReconcile)

  CurrentDocFieldOrder = "Name, UserFieldText1, DocIDParent, UserFieldDate1, DocID, Author, DateActivation, Content, ListToReconcile, NameAproval, Correspondent, SecurityLevel, Department "
  CurrentDocRequiredFields = "Name, UserFieldText1, UserFieldDate1, DateActivation, NameAproval, ListToView"

'Запрос №1 - СИБ - start
  If InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 Then
'  If InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) = 1 Then
    S_ListToReconcile_Comment = oPayDox.GetExtTableValue("AgreeSIBRU","Category",Session("CurrentClassDoc")+"/","List")
  Else
    S_ListToReconcile_Comment = oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")
  End If
'Запрос №1 - СИБ - end

  CurrentDocFieldOrder = "Name, UserFieldText1, DocIDParent, UserFieldDate1, DocID, Author, DateActivation, Content, UserFieldDate2, ListToReconcile, NameAproval, ListToView, SecurityLevel, Department "
  CurrentDocRequiredFields = "Name, UserFieldText1, UserFieldDate1, DateActivation, UserFieldDate2, NameAproval, ListToView"

  If InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) = 1 then
    S_LocationPath = ""
    bLocationPath = ""
  Else
    CurrentDocRequiredFields = CurrentDocRequiredFields+", DocListToRegister" 
  End If

  VAR_DirPictNotToShow = VAR_DirPictNotToShow+", DocName"
  'rmanyushin 60298 02.11.2009 Start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", ContractType"
  'rmanyushin 60298 02.11.2009 End
'ph - 20120216 - start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", AddFieldText1"
'ph - 20120216 - end

'Запрос №1 - СИБ - start
  If InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 Then
    VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", UserFieldText1"
    CurrentDocRequiredFields = Replace(CurrentDocRequiredFields, "UserFieldText1", "")
  End If
'Запрос №1 - СИБ - end

'Запрос №1 - СИБ - start
  If InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) <> 1 Then
'Запрос №1 - СИБ - end
    'Выпадающие списки
    Select Case UCase(Request("l"))
      Case "RU" 'RU
        S_UserFieldText1_Select = GetUserDirValues("{BFC71550-2605-4679-8A3F-C04211891D7E}")
      Case "" 'EN
        S_UserFieldText1_Select = GetUserDirValues("{1B136CF3-83CF-471D-925A-EEB72BC6CD5B}")
      Case "3" 'CZ
        S_UserFieldText1_Select = GetUserDirValues("{1EBA180D-7657-4E72-A678-9ECE4EDD58C1}")
    End Select
'Запрос №1 - СИБ - start
  End If
'Запрос №1 - СИБ - end

  'Для пользователей СТС везде добавляем поле бизнес единица
  If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
    If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 Then
      If UCase(Request("create")) = "Y" Then
        sCreator = Session("UserID")
      Else
        sCreator = GetUserID(S_Author)
      End If
      BusinessUnitsList = GetUsersBusinessUnits(sCreator)
      If InStr(BusinessUnitsList, VbCrLf) > 0 Then
        S_AddField2 = " "
        S_AddField_Select2 = BusinessUnitsList
        'Поле обязательно для заполнения
        CurrentDocRequiredFields = CurrentDocRequiredFields+", BusinessUnit"
      Else
        S_AddField_Set2 = BusinessUnitsList
      End If
    End If
  Else
    'Поле показывается только для пользователей СТС и администратора (у админа не из СТС оно необязательное, редактируемое)
    If not IsAdmin() Then
      VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", BusinessUnit"
    End If
  End If

' rmanyushin 93755, 22.04.2009, Start
'ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(STS_SLUZH_ZAPISKA_OVERTIME)) = 1 Then 'УПРАЗДНЕНА
'  'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
'  S_SecurityLevel = 4
'  If not IsAdmin() Then
'    S_SecurityLevel_Set = S_SecurityLevel
'  End If
'
'  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
'    S_AdditionalUsers = ""
'  End If
'
'  If S_AdditionalUsers <> "" Then
'    S_AdditionalUsers_Set = S_AdditionalUsers
'  Else
'    S_AdditionalUsers_Set = " "
'  End If
'
'  S_DocID_Set = S_DocID
'  S_DocIDAdd_Set = S_DocIDAdd
'  S_Department_Set = S_Department
'  S_Author_Set = S_Author
'
'  ' Создание карточки документа
'  If UCase(Request("create")) = "Y" Then
'    If Request("UpdateDoc") <> "YES" Then
'      'SAY 2009-02-20
'      VAR_ChangeDocGenerateButton=""
'      S_DocID = " "
'      If S_DocID_Set<>"" Then
'        S_DocID_Set = S_DocID
'      End If
'
'      'SAY 2008-11-06
'      S_Department_Set = Session("Department")
'
'      'SAY 2008-09-22 для подчиненных документов оставляем весь список согласования родителя
'      If Trim(S_DocIDParent)<>"" then
'        S_ListToReconcile = Replace(S_ListToReconcile,"##;","")
'      End If
'
'      S_DocID_Set = " "
'      S_DateActivation = MyDate(Date)
'      S_UserFieldDate1 = MyDate(Date)
'      S_Author_Set = InsertionName(Session("Name"), Session("UserID"))
'      S_Department_Set = Session("Department")
'    End If
'  End If 'UCase(Request("create")) = "Y"
'
'  If UCase(Request("create")) = "Y" Then
'    If Request("UpdateDoc") <> "YES" Then
'      S_Name = STS_SLUZH_ZAPISKA_OVERTIME_TITLE
'      S_Name_Set = ""
'      S_NameAproval_Set = STS_GenDirector
'      S_ListToReconcile_Comment = oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")
'      S_ListToReconcile = ""
'      S_ListToReconcile_Set = ""
'      S_ListToView_Set= oPayDox.GetExtTableValue("RecipientSTS"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")
'      S_Content = STS_SLUZH_ZAPISKA_OVERTIME_PRIKAZ_TEXT
'    End If
'  End If 'UCase(Request("create")) = "Y"
'
'  S_ListToReconcile = DeleteConstPrefixFromList(S_ListToReconcile)
'
'  CurrentDocFieldOrder = "Name, UserFieldText1, DocIDParent, UserFieldDate1, DocID, Author, DateActivation, Content, ListToReconcile, NameAproval, Correspondent, SecurityLevel, Department "
'  CurrentDocRequiredFields = "Name, UserFieldText1, UserFieldDate1, DateActivation, NameAproval, ListToView"
'
'  If InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) = 1 Then
'    S_ListToReconcile_Comment = oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")
'  End If
'' rmanyushin 93755, 22.04.2009, Stop
'
'  If InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) = 1 then
'    S_LocationPath = ""
'    bLocationPath = ""
'  Else
'    CurrentDocRequiredFields = CurrentDocRequiredFields+", DocListToRegister" 
'  End If
'
'  VAR_DirPictNotToShow = VAR_DirPictNotToShow+", DocName"
'  'rmanyushin 60298 02.11.2009 Start
'  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", ContractType"
'  'rmanyushin 60298 02.11.2009 End
'
'  'Выпадающие списки
'  Select Case UCase(Request("l"))
'    Case "RU" 'RU
'      S_UserFieldText1_Select = GetUserDirValues("{BFC71550-2605-4679-8A3F-C04211891D7E}")
'    Case "" 'EN
'      S_UserFieldText1_Select = GetUserDirValues("{1B136CF3-83CF-471D-925A-EEB72BC6CD5B}")
'    Case "3" 'CZ
'      S_UserFieldText1_Select = GetUserDirValues("{1EBA180D-7657-4E72-A678-9ECE4EDD58C1}")
'  End Select
'
'  'Для пользователей СТС везде добавляем поле бизнес единица
'  If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
'    If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 Then
'      If UCase(Request("create")) = "Y" Then
'        sCreator = Session("UserID")
'      Else
'        sCreator = GetUserID(S_Author)
'      End If
'      BusinessUnitsList = GetUsersBusinessUnits(sCreator)
'      If InStr(BusinessUnitsList, VbCrLf) > 0 Then
'        S_AddField2 = " "
'        S_AddField_Select2 = BusinessUnitsList
'        'Поле обязательно для заполнения
'        CurrentDocRequiredFields = CurrentDocRequiredFields+", BusinessUnit"
'      Else
'        S_AddField_Set2 = BusinessUnitsList
'      End If
'    End If
'  Else
'    'Поле показывается только для пользователей СТС и администратора (у админа не из СТС оно необязательное, редактируемое)
'    If not IsAdmin() Then
'      VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", BusinessUnit"
'    End If
'  End If
'
''rmanyushin 136964 08.11.2010 Start
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) = 1 Then
  'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
  S_SecurityLevel = 4
  If not IsAdmin() Then
    S_SecurityLevel_Set = S_SecurityLevel
  End If

  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
    S_AdditionalUsers = ""
  End If

  If S_AdditionalUsers <> "" Then
    S_AdditionalUsers_Set = S_AdditionalUsers
  Else
    S_AdditionalUsers_Set = " "
  End If

  S_DocID_Set = S_DocID
  S_DocIDAdd_Set = S_DocIDAdd
  S_Department_Set = S_Department
  S_Author_Set = S_Author

  ' Создание карточки документа
  If UCase(Request("create")) = "Y" Then
    'Запрос №46 - СТС - start
    'Предупреждение пользователю
	If InStr(UCase(Session("Department")), UCase(SIT_STS_ROOT_DEPARTMENT)) = 1 Then
		S_UserInstructions = "<font color = red><b>" & SIT_CannotCreateDocInThisCategory & "</b></font>"
	End If
	'Запрос №46 - СТС - end

    If Request("UpdateDoc") <> "YES" Then
      'SAY 2009-02-20
      VAR_ChangeDocGenerateButton=""
      S_DocID = " "
      If S_DocID_Set<>"" Then
        S_DocID_Set = S_DocID
      End If

      'SAY 2008-11-06
      S_Department_Set = Session("Department")

      'SAY 2008-09-22 для подчиненных документов оставляем весь список согласования родителя
      If Trim(S_DocIDParent)<>"" then
        S_ListToReconcile = Replace(S_ListToReconcile,"##;","")
      End If

      S_DocID_Set = " "
      S_DateActivation = MyDate(Date)
      S_UserFieldDate1 = MyDate(Date)
      S_Author_Set = InsertionName(Session("Name"), Session("UserID"))
      S_Department_Set = Session("Department")
    End If
  End If 'UCase(Request("create")) = "Y"

  If UCase(Request("create")) = "Y" Then
    If Request("UpdateDoc") <> "YES" Then
'Запрос №1 - СИБ - start
      If InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 Then
        S_NameAproval = SIB_AssistantDirector
        S_NameAproval_Set = S_NameAproval
        S_Name = SIB_SLUZH_ZAPISKA_OVERTIME_TITLE
        S_Name_Set = ""
        S_Content = SIB_SLUZH_ZAPISKA_OVERTIME_PRIKAZ_TEXT
        S_ListToView = SIB_HRManager
        S_ListToView_Set = S_ListToView
        S_UserFieldText4 = " "
        S_ListToReconcile_Comment = oPayDox.GetExtTableValue("AgreeSIBRU","Category",Session("CurrentClassDoc")+"/","List")
	  Else
'Запрос №1 - СИБ - end
        S_NameAproval = " "
        S_NameAproval_Set = ""
        S_Name = STS_SLUZH_ZAPISKA_OVERTIME_TITLE
        S_Name_Set = ""
        S_Content = STS_SLUZH_ZAPISKA_OVERTIME_PRIKAZ_TEXT
        'Запрос №34 - СТС - start
        S_ListToView = " "
        'Запрос №34 - СТС - end
        S_ListToView_Set = ""
        S_UserFieldText4 = " "
        'Запрос №34 - СТС - start
        S_ListToReconcile = ""
        'Запрос №34 - СТС - end
        S_ListToReconcile_Set = ""
'Запрос №1 - СИБ - start
      End If 'InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 Then
'Запрос №1 - СИБ - end
    End If
  End If 'UCase(Request("create")) = "Y"
'rmanyushin 136964 08.11.2010 End

  S_ListToReconcile = DeleteConstPrefixFromList(S_ListToReconcile)

  CurrentDocFieldOrder = "Name, UserFieldText1, DocIDParent, UserFieldDate1, DocID, Author, DateActivation, Content, ListToReconcile, NameAproval, Correspondent, SecurityLevel, Department "
  CurrentDocRequiredFields = "Name, UserFieldText1, UserFieldDate1, DateActivation, NameAproval, ListToView"

  If InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) = 1 Then
    S_ListToReconcile_Comment = oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")
'Запрос №1 - СИБ - start
  ElseIf InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 Then
    S_ListToReconcile_Comment = oPayDox.GetExtTableValue("AgreeSIBRU","Category",Session("CurrentClassDoc")+"/","List")
'Запрос №1 - СИБ - end
  End If

'rmanyushin 136964 08.11.2010 Start
  S_UserFieldText4_Set = S_UserFieldText4
  S_NameAproval_Set = S_NameAproval
  S_ListToReconcile_Set = S_ListToReconcile
  S_ListToView_Set= S_ListToView

  CurrentDocFieldOrder = "DocID, DocIDParent, Name, Author, NameResponsible, Content, UserFieldText3, UserFieldText4, NameAproval, ListToReconcile, ListToView, UserFieldDate1, DateActivation, Resolution, UserFieldText1, UserFieldDate2, UserFieldText5, UserFieldText6"
  CurrentDocRequiredFields = CurrentDocRequiredFields+", UserFieldText3" 
'Запрос №34 - СТС - start
  'Получатели необязательны, добавляются правилом
  CurrentDocRequiredFields = Replace(CurrentDocRequiredFields, "ListToView", "")
  'Подписант необязателен, добавляется правилом
  CurrentDocRequiredFields = Replace(CurrentDocRequiredFields, "NameAproval", "")
'Запрос №34 - СТС - end
'ph - 20110128 - start
  'Заказчик переработки обязателен
  CurrentDocRequiredFields = CurrentDocRequiredFields+", NameResponsible" 
'ph - 20110128 - end
'rmanyushin 136964 08.11.2010 End

  If InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) = 1 then
    S_LocationPath = ""
    bLocationPath = ""
  Else
    CurrentDocRequiredFields = CurrentDocRequiredFields+", DocListToRegister" 
  End If

  VAR_DirPictNotToShow = VAR_DirPictNotToShow+", DocName"
  'rmanyushin 60298 02.11.2009 Start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", ContractType"
  'rmanyushin 60298 02.11.2009 End
'ph - 20120216 - start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", AddFieldText1, AddFieldText2"
'ph - 20120216 - end

'Запрос №1 - СИБ - start
  If InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 Then
    VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", UserFieldText3, UserFieldText4"
    CurrentDocRequiredFields = Replace(CurrentDocRequiredFields, "UserFieldText3", "")
  End If
'Запрос №1 - СИБ - end
'ph - 20110629 - start - Нет этого поля в категории
'  'Выпадающие списки
'  Select Case UCase(Request("l"))
'    Case "RU" 'RU
'      S_UserFieldText1_Select = GetUserDirValues("{BFC71550-2605-4679-8A3F-C04211891D7E}")
'    Case "" 'EN
'      S_UserFieldText1_Select = GetUserDirValues("{1B136CF3-83CF-471D-925A-EEB72BC6CD5B}")
'    Case "3" 'CZ
'      S_UserFieldText1_Select = GetUserDirValues("{1EBA180D-7657-4E72-A678-9ECE4EDD58C1}")
'  End Select
'ph - 20110629 - end - Нет этого поля в категории

  'Для пользователей СТС везде добавляем поле бизнес единица
  If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
    If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 Then
      If UCase(Request("create")) = "Y" Then
        sCreator = Session("UserID")
      Else
        sCreator = GetUserID(S_Author)
      End If
      BusinessUnitsList = GetUsersBusinessUnits(sCreator)
      If InStr(BusinessUnitsList, VbCrLf) > 0 Then
        S_AddField2 = " "
        S_AddField_Select2 = BusinessUnitsList
        'Поле обязательно для заполнения
        CurrentDocRequiredFields = CurrentDocRequiredFields+", BusinessUnit"
      Else
        S_AddField_Set2 = BusinessUnitsList
      End If
    End If
  Else
    'Поле показывается только для пользователей СТС и администратора (у админа не из СТС оно необязательное, редактируемое)
    If not IsAdmin() Then
      VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", BusinessUnit"
    End If
  End If

'Запрос №46 - СТС - start
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(STS_SLUZH_ZAPISKA_OVERTIME_PLAN)) = 1 Then
	'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
	S_SecurityLevel = 4
	If not IsAdmin() Then
		S_SecurityLevel_Set = S_SecurityLevel
	End If

	If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then 'Пока только пользователи СТС
		' Создание карточки документа
		If UCase(Request("create")) = "Y" Then
			If Request("UpdateDoc") <> "YES" Then
				'SAY 2009-02-20
				VAR_ChangeDocGenerateButton = ""
				S_DocID = " "
				S_Author = InsertionName(Session("Name"), Session("UserID"))
				S_Department = Session("Department")
				S_Name = STS_SLUZH_ZAPISKA_OVERTIME_PLAN_TITLE
				S_Content = STS_SLUZH_ZAPISKA_OVERTIME_PLAN_TEXT
				S_ListToReconcile = " "
				S_NameAproval = " "
				S_ListToView = " "
				S_AdditionalUsers = " "
			End If
		End If 'UCase(Request("create")) = "Y"

		S_DocID_Set = S_DocID
		S_DocIDAdd_Set = S_DocIDAdd
		S_Department_Set = S_Department
		S_Author_Set = S_Author
		If S_NameAproval = "" Then
			S_NameAproval = " "
		End If
		S_NameAproval_Set = S_NameAproval
		If S_ListToReconcile = "" Then
			S_ListToReconcile = " "
		End If
		S_ListToReconcile_Set = S_ListToReconcile
		If S_ListToView = "" Then
			S_ListToView = " "
		End If
		S_ListToView_Set = S_ListToView
		If S_AdditionalUsers = "" Then
			S_AdditionalUsers = " "
		End If
		S_AdditionalUsers_Set = S_AdditionalUsers
		'Скрыть кнопки изменения отдельных полей
		VAR_NotToShowSaveButtonsInChangeDoc = "Y"
		'If UCase(Request("create")) <> "Y" Then
		'	If Request("UpdateDoc") <> "YES" Then
		'		S_Correspondent_Set = S_Correspondent
		'	End If
		'End If

		CurrentDocFieldOrder = "DocID, Name, Author, Content, NameResponsible, UserFieldText3, Correspondent, ListToReconcile, NameAproval, ListToView"
		CurrentDocRequiredFields = "Name, NameResponsible, UserFieldText3, Correspondent"
		VAR_DirPictNotToShow = "DocName"
		VAR_DocFieldsNotToShow = "ContractType"
'ph - 20120216 - start
		VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", AddFieldText1, AddFieldText2"
'ph - 20120216 - end
	End If 'InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then 'Пока только пользователи СТС

	'Для пользователей СТС везде добавляем поле бизнес единица
	If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
		If UCase(Request("create")) = "Y" Then
			sCreator = Session("UserID")
		Else
			sCreator = GetUserID(S_Author)
		End If
		BusinessUnitsList = GetUsersBusinessUnits(sCreator)
		If InStr(BusinessUnitsList, VbCrLf) > 0 Then
			S_AddField2 = " "
			S_AddField_Select2 = BusinessUnitsList
			'Поле обязательно для заполнения
			CurrentDocRequiredFields = CurrentDocRequiredFields+", BusinessUnit"
		Else
			S_AddField_Set2 = BusinessUnitsList
		End If
	Else
		'Поле показывается только для пользователей СТС и администратора (у админа не из СТС оно необязательное, редактируемое)
		If not IsAdmin() Then
			VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", BusinessUnit"
		End If
	End If
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(STS_SLUZH_ZAPISKA_OVERTIME_FACT)) = 1 Then
	'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
	S_SecurityLevel = 4
	If not IsAdmin() Then
		S_SecurityLevel_Set = S_SecurityLevel
	End If

	If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then 'Пока только пользователи СТС
		' Создание карточки документа
		If UCase(Request("create")) = "Y" Then
			If Request("UpdateDoc") <> "YES" Then
				'SAY 2009-02-20
				VAR_ChangeDocGenerateButton = ""
				S_DocID = " "
				S_Author = InsertionName(Session("Name"), Session("UserID"))
				S_Department = Session("Department")
				S_ListToReconcile = " "
				S_NameAproval = " "
				S_ListToView = " "
				S_AdditionalUsers = " "
			End If
		End If 'UCase(Request("create")) = "Y"

		S_DocID_Set = S_DocID
		S_DocIDAdd_Set = S_DocIDAdd
		S_Department_Set = S_Department
		S_Author_Set = S_Author
		If S_DocIDParent = "" Then
			S_DocIDParent = " "
		End If
		S_DocIDParent_Set = S_DocIDParent
		'S_Name_Set = S_Name
		'S_Content_Set = S_Content
		If S_Correspondent = "" Then
			S_Correspondent = " "
		End If
		S_Correspondent_Set = S_Correspondent
'{ph - 20120327
'		If S_NameResponsible = "" Then
'			S_NameResponsible = " "
'		End If
'		S_NameResponsible_Set = S_NameResponsible
		S_NameResponsible_Set = ""
'ph - 20120327}
		If S_UserFieldText3 = "" Then
			S_UserFieldText3 = " "
		End If
		S_UserFieldText3_Set = S_UserFieldText3
		If S_NameAproval = "" Then
			S_NameAproval = " "
		End If
		S_NameAproval_Set = S_NameAproval
		If S_ListToReconcile = "" Then
			S_ListToReconcile = " "
		End If
		S_ListToReconcile_Set = S_ListToReconcile
		If S_ListToView = "" Then
			S_ListToView = " "
		End If
		S_ListToView_Set = S_ListToView
		If S_AdditionalUsers = "" Then
			S_AdditionalUsers = " "
		End If
		S_AdditionalUsers_Set = S_AdditionalUsers

		CurrentDocFieldOrder = "DocID, DocIDParent, Name, Author, Content, NameResponsible, UserFieldText3, Correspondent, ListToReconcile, NameAproval, ListToView"
		'Закрыты от редактирования обязательные поля, берутся из родительской записки
		'CurrentDocRequiredFields = "DocIDParent, Name, NameResponsible, UserFieldText3, Correspondent"
		CurrentDocRequiredFields = "Name"

		VAR_DirPictNotToShow = "DocName"
		VAR_DocFieldsNotToShow = "ContractType"
'ph - 20120216 - start
		VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", AddFieldText1, AddFieldText2"
'ph - 20120216 - end
	End If 'InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then 'Пока только пользователи СТС

	'Для пользователей СТС везде добавляем поле бизнес единица
	If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
		If UCase(Request("create")) = "Y" Then
			sCreator = Session("UserID")
		Else
			sCreator = GetUserID(S_Author)
		End If
		BusinessUnitsList = GetUsersBusinessUnits(sCreator)
		If InStr(BusinessUnitsList, VbCrLf) > 0 Then
			S_AddField2 = " "
			S_AddField_Select2 = BusinessUnitsList
			'Поле обязательно для заполнения
			CurrentDocRequiredFields = CurrentDocRequiredFields+", BusinessUnit"
		Else
			S_AddField_Set2 = BusinessUnitsList
		End If
	Else
		'Поле показывается только для пользователей СТС и администратора (у админа не из СТС оно необязательное, редактируемое)
		If not IsAdmin() Then
			VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", BusinessUnit"
		End If
	End If
'Запрос №46 - СТС - end

'rmanyushin 119579 19.08.2010 Start
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(STS_SLUZH_ZAPISKA_HOLIDAY)) = 1 Then
  'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
  S_SecurityLevel = 4
  If not IsAdmin() Then
    S_SecurityLevel_Set = S_SecurityLevel
  End If

  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
    S_AdditionalUsers = ""
  End If

  If S_AdditionalUsers <> "" Then
    S_AdditionalUsers_Set = S_AdditionalUsers
  Else
    S_AdditionalUsers_Set = " "
  End If

  S_DocID_Set = S_DocID
  S_DocIDAdd_Set = S_DocIDAdd
  S_Department_Set = S_Department
  S_Author_Set = S_Author

  ' Создание карточки документа
  If UCase(Request("create")) = "Y" Then
    If Request("UpdateDoc") <> "YES" Then
      'SAY 2009-02-20
      VAR_ChangeDocGenerateButton=""
      S_DocID = " "
      If S_DocID_Set<>"" Then
        S_DocID_Set = S_DocID
      End If

      'SAY 2008-11-06
      S_Department_Set = Session("Department")

      'SAY 2008-09-22 для подчиненных документов оставляем весь список согласования родителя
      If Trim(S_DocIDParent)<>"" then
        S_ListToReconcile = Replace(S_ListToReconcile,"##;","")
      End If

      S_DocID_Set = " "
      S_DateActivation = MyDate(Date)
      S_UserFieldDate1 = MyDate(Date)
      S_Author_Set = InsertionName(Session("Name"), Session("UserID"))
      S_Department_Set = Session("Department")
    End If
  End If 'UCase(Request("create")) = "Y"

  If UCase(Request("create")) = "Y" Then
    If Request("UpdateDoc") <> "YES" Then
      S_Name = STS_SLUZH_ZAPISKA_HOLIDAY_TITLE
      S_Name_Set = ""
      '{ph - 20120326
		If InStr(UCase(Session("Department")), UCase(SIT_STS)) <> 1 Then

      S_NameAproval_Set = GetApproverForSTS_HolidayRequest(Session("Department"))
      S_ListToReconcile_Comment = oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")
      S_ListToReconcile = ""
      S_ListToReconcile_Set = ""
      S_ListToView_Set = oPayDox.GetExtTableValue("RecipientSTS"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")
      		Else
			S_NameAproval = " "
			S_ListToReconcile = ""
			S_ListToView = " "
		End If
'ph - 20120326}

    End If
  End If 'UCase(Request("create")) = "Y"
'rmanyushin 119579 19.08.2010 Stop

  '{ph - 20120326
	If Request("UpdateDoc") <> "YES" Then
		If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
			S_NameAproval_Set = S_NameAproval
			'S_ListToReconcile_Set = S_ListToReconcile
			S_ListToView_Set = S_ListToView
		End If
	End If
'ph - 20120326}

'{ph - 20120326
	If InStr(UCase(Session("Department")), UCase(SIT_STS)) <> 1 Then
		S_ListToReconcile = DeleteConstPrefixFromList(S_ListToReconcile)
	End If
'ph - 20120326}

  CurrentDocFieldOrder = "Name, UserFieldText1, DocIDParent, UserFieldDate1, DocID, Author, DateActivation, Content, ListToReconcile, NameAproval, Correspondent, SecurityLevel, Department "
'{ph - 20120326
'  CurrentDocRequiredFields = "Name, UserFieldText1, UserFieldDate1, DateActivation, NameAproval, ListToView"
  CurrentDocRequiredFields = "Name, UserFieldText1, UserFieldDate1, DateActivation"
  If InStr(UCase(Session("Department")), UCase(SIT_STS)) <> 1 Then
    CurrentDocRequiredFields = CurrentDocRequiredFields & ",NameAproval,ListToView"
  End If
'ph - 20120326}


  If InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) = 1 Then
    S_ListToReconcile_Comment = oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")
  End If

  'rmanyushin 119579 19.08.2010 Start
  CurrentDocRequiredFields = CurrentDocRequiredFields+", Content" 
  'rmanyushin 119579 19.08.2010 End

  If InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) = 1 then
    S_LocationPath = ""
    bLocationPath = ""
  Else
    CurrentDocRequiredFields = CurrentDocRequiredFields+", DocListToRegister" 
  End If

  VAR_DirPictNotToShow = VAR_DirPictNotToShow+", DocName"
  'rmanyushin 60298 02.11.2009 Start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", ContractType"
  'rmanyushin 60298 02.11.2009 End
'ph - 20120216 - start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", AddFieldText1"
'ph - 20120216 - end

'ph - 20110629 - start - Нет этого поля в категории
'  'Выпадающие списки
'  Select Case UCase(Request("l"))
'    Case "RU" 'RU
'      S_UserFieldText1_Select = GetUserDirValues("{BFC71550-2605-4679-8A3F-C04211891D7E}")
'    Case "" 'EN
'      S_UserFieldText1_Select = GetUserDirValues("{1B136CF3-83CF-471D-925A-EEB72BC6CD5B}")
'    Case "3" 'CZ
'      S_UserFieldText1_Select = GetUserDirValues("{1EBA180D-7657-4E72-A678-9ECE4EDD58C1}")
'  End Select
'ph - 20110629 - end - Нет этого поля в категории

  'Для пользователей СТС везде добавляем поле бизнес единица
  If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
    If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 Then
      If UCase(Request("create")) = "Y" Then
        sCreator = Session("UserID")
      Else
        sCreator = GetUserID(S_Author)
      End If
      BusinessUnitsList = GetUsersBusinessUnits(sCreator)
      If InStr(BusinessUnitsList, VbCrLf) > 0 Then
        S_AddField2 = " "
        S_AddField_Select2 = BusinessUnitsList
        'Поле обязательно для заполнения
        CurrentDocRequiredFields = CurrentDocRequiredFields+", BusinessUnit"
      Else
        S_AddField_Set2 = BusinessUnitsList
      End If
    End If
  Else
    'Поле показывается только для пользователей СТС и администратора (у админа не из СТС оно необязательное, редактируемое)
    If not IsAdmin() Then
      VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", BusinessUnit"
    End If
  End If

' *** ОСТАЛЬНЫЕ СЛУЖЕБНЫЕ ЗАПИСКИ
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_SLUZH_ZAPISKA)) = 1 Then
  'Убрать из всех типов документов поле уровень доступа. Оставляем только админам
  S_SecurityLevel = 4
  If not IsAdmin() Then
    S_SecurityLevel_Set = S_SecurityLevel
  End If

  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
    S_AdditionalUsers = ""
  End If
  
  'rti
  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" and InStr(UCase(Session("Department")), UCase(SIT_RTI)) = 1 Then
    S_AdditionalUsers = """Боев С. Ф."" <boev_oaorti>;"
    If InStr(UCase(Session("Department")), UCase("Управление ИТ")) > 0 Then
      S_AdditionalUsers = S_AdditionalUsers + VbCrLf + """Иванов В. В."" <vivanon_oaorti>;"
    End If
    If S_DocIDParent <>"" and InStr(UCase(S_DocIDParent),"T_") > 0 Then 
        S_AdditionalUsers = S_AdditionalUsers + S_Author + VbCrLf + S_NameResponsible + VbCrLf + S_Correspondent + VbCrLf + S_NameControl
    End If
    'если служебка создается из поручения, то даем доступ автору поручения,исполнителю, соисполнителю,контролеру
  End If
  'rti

    'vtss
  If UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" and InStr(UCase(Session("Department")), UCase(SIT_VTSS)) = 1 Then
    S_AdditionalUsers = """Подольский А. Е."" <vtss_a.podolskii>;" + VbCrLf
    If S_DocIDParent <>"" and InStr(UCase(S_DocIDParent),"T_") > 0 Then 
        S_AdditionalUsers = S_AdditionalUsers + S_Author + VbCrLf + S_NameResponsible + VbCrLf + S_Correspondent + VbCrLf + S_NameControl
    End If
    'если служебка создается из поручения, то даем доступ автору поручения,исполнителю, соисполнителю,контролеру
  End If
  'vtss

  If S_AdditionalUsers <> "" Then
    S_AdditionalUsers_Set = S_AdditionalUsers
  Else
    S_AdditionalUsers_Set = " "
  End If

  S_DocID_Set = S_DocID
  S_DocIDAdd_Set = S_DocIDAdd
  S_Department_Set = S_Department
  S_Author_Set = S_Author
  S_DocIDParent_Set = S_DocIDParent

  ' Создание карточки документа
  If UCase(Request("create")) = "Y" Then
    If Request("UpdateDoc") <> "YES" Then
      'SAY 2009-02-20
      VAR_ChangeDocGenerateButton=""
      S_DocID = " "
      If S_DocID_Set<>"" Then
        S_DocID_Set = S_DocID
      End If

      'SAY 2008-11-06
      S_Department_Set = Session("Department")

      'SAY 2008-09-22 для подчиненных документов оставляем весь список согласования родителя
      If Trim(S_DocIDParent)<>"" then
        S_ListToReconcile = Replace(S_ListToReconcile,"##;","")
      End If

      S_DocID_Set = " "
      S_DateActivation = MyDate(Date)
      S_UserFieldDate1 = MyDate(Date)
      S_Author_Set = InsertionName(Session("Name"), Session("UserID"))
      S_Department_Set = Session("Department")
    End If
  End If 'UCase(Request("create")) = "Y"

'ph - 20110506 - start
  S_ListToReconcile = DeleteConstPrefixFromList(S_ListToReconcile)

  CurrentDocFieldOrder = "Name, UserFieldText1, DocIDParent, UserFieldDate1, DocID, Author, DateActivation, Content, ListToReconcile, NameAproval, Correspondent, SecurityLevel, Department "
  CurrentDocRequiredFields = "Name, UserFieldText1, UserFieldDate1, DateActivation, NameAproval, ListToView"

  If InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) = 1 Then
    S_ListToReconcile_Comment = oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")
  End If

  'rmanyushin 119579 19.08.2010 Start
  CurrentDocRequiredFields = CurrentDocRequiredFields+", Content" 
  'rmanyushin 119579 19.08.2010 End

  If InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) = 1 or InStr(UCase(Session("Department")), UCase(SIT_RTI)) = 1 then
    S_LocationPath = ""
    bLocationPath = ""
  Else
    CurrentDocRequiredFields = CurrentDocRequiredFields+", DocListToRegister" 
  End If

  VAR_DirPictNotToShow = VAR_DirPictNotToShow+", DocName"
  'rmanyushin 60298 02.11.2009 Start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", ContractType"
  'rmanyushin 60298 02.11.2009 End
'ph - 20120216 - start
  VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", AddFieldText1, AddFieldText2"
'ph - 20120216 - end

'Запрос №1 - СИБ - start
  If InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 Then
    VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", UserFieldText1"
    CurrentDocRequiredFields = Replace(CurrentDocRequiredFields, "UserFieldText1", "")
  End If
'Запрос №1 - СИБ - end

'Запрос №1 - СИБ - start
  If InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) <> 1 Then
'Запрос №1 - СИБ - end
    'Выпадающие списки
    Select Case UCase(Request("l"))
      Case "RU" 'RU
        S_UserFieldText1_Select = GetUserDirValues("{BFC71550-2605-4679-8A3F-C04211891D7E}")
      Case "" 'EN
        S_UserFieldText1_Select = GetUserDirValues("{1B136CF3-83CF-471D-925A-EEB72BC6CD5B}")
      Case "3" 'CZ
        S_UserFieldText1_Select = GetUserDirValues("{1EBA180D-7657-4E72-A678-9ECE4EDD58C1}")
    End Select
'Запрос №1 - СИБ - start
  End If
'Запрос №1 - СИБ - end

'ph - 20110506 - end

 'Служебные записки РТИ    (общая форма)
    If InStr(UCase(Session("Department")), UCase(SIT_RTI)) = 1 Then
        CurrentDocFieldOrder = "Name, UserFieldText2, DocIDParent, UserFieldDate1, DocID, Author, DateActivation, Content, ListToReconcile, NameAproval, Correspondent, SecurityLevel, Department "
        CurrentDocRequiredFields = "Name, UserFieldText2, UserFieldDate1, DateActivation, NameAproval, ListToView"
        if UCase(Request("create")) = "Y" and Request("UpdateDoc") <> "YES" Then
          S_UserFieldText2 = " "
        End If
        S_UserFieldText2_Select = GetUserDirValues("{1C5F3936-D1DE-492B-9736-FB5E5C71DE6D}")
        S_ListToReconcile_Comment = oPayDox.GetExtTableValue("AgreeRTIRU","Category",Session("CurrentClassDoc")+"/","List")        
        VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", UserFieldText1"
    end if
     
 'vtss
    If InStr(UCase(Session("Department")), UCase(SIT_VTSS)) = 1 Then
        CurrentDocFieldOrder = "Name, DocIDParent, UserFieldDate1, DocID, Author, DateActivation, Content, ListToReconcile, NameAproval, Correspondent, SecurityLevel, Department "
        CurrentDocRequiredFields = "Name, UserFieldDate1, DateActivation, NameAproval, ListToView"
        VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", UserFieldText2"
        S_ListToReconcile_Comment = oPayDox.GetExtTableValue("AgreeVTSSRU","Category",Session("CurrentClassDoc"),"List")        
        VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow & ", UserFieldText1"
    end if
 'vtss     

  'Для пользователей СТС везде добавляем поле бизнес единица
  If InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
    If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 Then
      If UCase(Request("create")) = "Y" Then
        sCreator = Session("UserID")
      Else
        sCreator = GetUserID(S_Author)
      End If
      BusinessUnitsList = GetUsersBusinessUnits(sCreator)
      If InStr(BusinessUnitsList, VbCrLf) > 0 Then
        S_AddField2 = " "
        S_AddField_Select2 = BusinessUnitsList
        'Поле обязательно для заполнения
        CurrentDocRequiredFields = CurrentDocRequiredFields+", BusinessUnit"
      Else
        S_AddField_Set2 = BusinessUnitsList
      End If
    End If
  Else
    'Поле показывается только для пользователей СТС и администратора (у админа не из СТС оно необязательное, редактируемое)
    If not IsAdmin() Then
      VAR_DocFieldsNotToShow = VAR_DocFieldsNotToShow+", BusinessUnit"
    End If
  End If

End If
'Запрос №17 - СТС - start
'Запоминаем подразделение документа (для рассылок, зависимых от БН)
Session("CurrentDepartmentDoc") = S_Department
'Запрос №17 - СТС - end

AddLogD "@@@@@@@@@@@@@S_ListToReconcile = '" & S_ListToReconcile&"'"
AddLogD "@@@@@@@@@@@@@S_ListToReconcile_Set = '" & S_ListToReconcile_Set&"'"
%>