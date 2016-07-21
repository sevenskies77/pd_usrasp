<%
Function UserChangeDocCheck()
AddLogD "UserChangeDocCheck"

'Place here ASP code for your ChangeDoc validation
'Use expression Request("DocID") to get the value of the DocID field
'Use expression Request("DocFieldName") to get the value of the other fields named FieldName 

'If InStr(Request("DocID"), "-")=0 Then
'	UserChangeDocCheck=False
'	Session("Message")="Error message"
'Else
'	UserChangeDocCheck=True
'	Session("Message")="User data validation OK"
'End If

'If Request("create")="y" And UCase(Request("UpdateDoc"))="YES" Then
'	S_DocID = "12345" 'Auto-resign doc id at the edit stage
'	S_DocIDAdd = "12345" 'Auto-resign additional doc id at the edit stage
'End If

AddLogD "UserChangeDocCheck 1"
If IsHelpDeskDoc() Then
If Request("create")="y" And UCase(Request("UpdateDoc"))="YES" Then
'AddLogD "UserChangeDocCheck 2"
'AddLogD "UserChangeDocCheck 3"
'If Trim(Request("justcreated"))<>"" Then
	If Trim(Request("DocNameResponsible"))<>"" Then
		If Trim(Request("DocCorrespondent"))<>"" Then
			If InStr(Trim(Request("DocCorrespondent")), GetLogin(Trim(Request("DocNameResponsible"))))<=0 Then
'AddLogD "SetDocField, NameResponsible="""" "
				'SetDocField Request("DocID"), "NameResponsible", ""
				S_NameResponsible_Set=""
				S_NameResponsible=""
			End If
		End If
	End If
	If (Not IsHelpDeskAdmin() And Not IsSupervisor()) Then
		If Trim(Request("DocRank"))="Срочный" Then
			S_DateCompletion=MyDate(Date+3)
		ElseIf Trim(Request("DocRank"))="Обычный" Then
			S_DateCompletion=MyDate(Date+5)
		Else
			S_DateCompletion=MyDate(Date+3)
		End If
	End If
	If (Not IsHelpDeskAdmin() And Not IsSupervisor()) Or Trim(Request("DocCorrespondent"))="" Then
sUserDirName="Группы пользователей"
'sUserDirName="Группы консультантов"
'nKeyField
sKeyFieldValue=Trim(Request("UserFieldText4"))
sKeyFieldValue2=Session("Company")
'Out "UserFieldText4:"+S_UserFieldText4

Set dsTemp = Server.CreateObject("ADODB.Recordset")
        If sVersion = "MSSQL" Then
sSQL = "select * from (UserDirectories Left Outer Join UserDirValues ON UserDirValues.UDKeyField = UserDirectories.KeyField) where Name='" + sUserDirName + "' And PATINDEX('%" + sKeyFieldValue + "%', Field1) <> 0 And Field3='" + sKeyFieldValue2 + "'"
        ElseIf sVersion = "MSACCESS" Then
sSQL = "select * from (UserDirectories Left Outer Join UserDirValues ON UserDirValues.UDKeyField = UserDirectories.KeyField) where Name='" + sUserDirName + "' And InStr(Field1, '" + sKeyFieldValue + "') <> 0 And Field3='" + sKeyFieldValue2 + "'"
        End If
AddLogD "***sSQL for UserDirectories:" + sSQL
dsTemp.Open sSQL, Conn, 3, 1, &H1
	If Not dsTemp.EOF Then
AddLogD "Correspondent:" + MyCStr(dsTemp("Field2"))
	'SetDocField Request("DocID"), "Correspondent", MyCStr(dsTemp("Field2"))
	S_Correspondent_Set=MyCStr(dsTemp("Field2"))
	S_Correspondent=S_Correspondent_Set
	AddLogD "FOUND"
Else
	AddLogD "EOF"
End If
dsTemp.Close
	End If
	
'End If
End If 'Request("create")="y" And UCase(Request("UpdateDoc"))="YES" Then

	If Request("UpdateDoc")="YES" Then
		
		If S_Correspondent_Set<>"" Then
			S_CorrespondentNew=S_Correspondent_Set
		Else
			S_CorrespondentNew=Request("DocCorrespondent")
		End If
		
		If S_CorrespondentNew<>Request("DocCorrespondentOld") And S_CorrespondentNew<>"" Then
			CreateAutoCommentType Request("DocID"), "Передано в обработку - группа: "+S_CorrespondentNew, "HISTORY"
			'sGetNextUserIDInList="<"+oPayDox.GetNextUserIDInList(S_CorrespondentNew, 1)+">"
'AddLogD "sGetNextUserIDInList: "+sGetNextUserIDInList
			'sNotificationList="D"
			'sNotificationSubject="Передано в обработку"
			'SendNotificationToUsers "", Null, Request("DocID"), sGetNextUserIDInList, "Передано в обработку"
		End If
		'If Request("DocResponsible")<>Request("DocResponsibleOld") And Request("DocResponsible")<>"" Then
		'	CreateAutoCommentType Request("DocID"), "Назначен ответственный исполнитель: "+Request("DocResponsible"), "HISTORY"
		'End If

	End If

End If 'IsHelpDeskDoc() Then

'If Not oPayDox.IsContainRightSymbolsOnly(Request("DocID"), DOCS_RightSymbols, WrongSymbol) Then
'	Session("Message")=Error_DoNotUseSpecSymbols1+": В«"+WrongSymbol+"В»"
'	Exit Function 
'End If 

UserChangeDocCheck=True 	'If everything is OK and data in DB can be changed
'UserChangeDocCheck=False 'If something wrong and data in DB can NOT be changed
'Session("Message")="Error message!" 'Message to display


'-------------------------------------------------------------- СИТРОНИКС -----------------------------------------------------------------
AddLogD "vnikvnik 1234" + Trim(S_ListToReconcile)

' *** ВХОДЯЩИЕ ДОКУМЕНТЫ
If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_VHODYASCHIE)) = 1 Then
	'Создатель документа
	If UCase(Request("create")) = "Y" Then
		sDocCreator = Session("UserID")
	Else
		sDocCreator = GetUserID(Request("DocAuthor"))
	End If

	S_NameAproval=Request("DocNameAproval")
	S_NameResponsible=Request("DocNameResponsible")
	S_NameControl=Request("DocNameControl")
	S_ListToReconcile=Request("DocListToReconcile")
	S_ListToEdit=Request("DocListToEdit")
	S_ListToView=Request("DocListToView")
	S_Correspondent=Request("DocCorrespondent")

	'Определяем бизнес направление
	sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))

	' SAY 2008-07-21
	Set dsTempPR = Server.CreateObject("ADODB.Recordset")

' DmGorsky_5: меняем местами проверку: вначале ищем SIT_SITRU ("СИТРОНИКС ИТ"), потом, если не найдено, - SIT_SITRONICS ("СИТРОНИКС")
	If InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRU))=1 then ' DmGorsky
		sDepartment = SIT_SITRU ' DmGorsky
		sDepartmentCode = "" ' DmGorsky
'Запрос №1 - СИБ - start
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SIB_ROOT_DEPARTMENT))=1 then
		sDepartment = SIT_SIB
		sDepartmentCode = ""
'Запрос №1 - СИБ - end
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRONICS))=1 then
		sDepartment = SIT_SITRONICS
		sDepartmentCode = ""
	Else
		sDepartment = SIT_STS

		sSQL = "select code from Departments where Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"' or Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"/'"
		AddLogD "WAY sSQL="+sSQL
		AddLogD "WAY Conn="+Conn
		dsTempPR.Open sSQL, Conn, 3, 1, &H1

		if not dsTempPR.EOF Then
			sDepartmentCode = dsTempPR("code")
			if InStrRev(sDepartmentCode, "/") > 0 then
				sDepartmentCode = Right(sDepartmentCode, Len(sDepartmentCode)-InStrRev(sDepartmentCode, "/"))
			End If
		Else
			sDepartmentCode="000000"
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInDepartmentCode1 + DelOtherLangFromFolder(Request("DocDepartment")) + SIT_ErrorInDepartmentCode2)
			UserChangeDocCheck=False
			Exit Function
		End If

		dsTempPR.Close
	End If

	If Request("create") = "y" Then
		sPostfix = ""
		sSearchCol = "DocID"
		sPrePrefix = ""
		'SAY 2009-02-20
		sSufix = ""

		'SAY 2009-03-12
		S_DocID = ""

		If InStr(UCase(S_ClassDoc), UCase(SIT_VHODYASCHIE)) > 0 Then
			S_DocID = GetNewDocIDForVhodyashie(S_ClassDoc, sDepartmentRoot, Request("UserFieldText7"), "", "", "", GetDZKCode(GetCodeFromCode_NameString(Request("BusinessUnit")), "LetterCode"))
		End If
		S_DocIDAdd = S_DocID
	End If
	If sDepartmentRoot <> SIT_STS or (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_RASP_DOCS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1) Then
		'реализация полужесткого маршрута
		S_ListToReconcile = Replace(S_ListToReconcile, Trim(SIT_AdditionalAgrees), "")
		'Для коммерческих предложений и заявок на закупку и оплату не нужна приписка про доп. согласующих  
		If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_COM_OFFERS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) <> 1 Then
			If S_ListToReconcile <> "" Then
				S_ListToReconcile = " " + SIT_AdditionalAgrees + DeleteConstPrefixFromList(S_ListToReconcile)
			End If
		End If
	End If
	'Для всех категорий кроме заявок подставляем пользователей вместо ролей
	If (InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1) and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) <> 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_DOGOVORI_NEW)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIS)) <> 0) Then
		'Получаем список ролей
		If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) = 1 Then 'Для договоров получаем список ролей на всех языках (могут быть роли на разных языках)
			sRoleList = GetFullRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"), "", "", "")
		Else
			sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))
		End If
		'Подставляем роли в поля
		S_NameAproval = ReplaceRolesInList(S_NameAproval, sRoleList)
		S_NameResponsible = ReplaceRolesInList(S_NameResponsible, sRoleList)
		S_NameControl = ReplaceRolesInList(S_NameControl, sRoleList)
		S_ListToEdit = ReplaceRolesInList(S_ListToEdit, sRoleList)
		S_ListToView = ReplaceRolesInList(S_ListToView, sRoleList)
		S_ListToReconcile = ReplaceRolesInList(S_ListToReconcile, sRoleList)

		'Удаляем повторы в списке согласования
		S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)

		'vnik_error
		'Убираем автора документа из списка согласующих, т.к. после активации он не сможет загрузить новый документ
		'(работает пока только, если пользователь создает документ не под ролью)
		S_ListToReconcile = Replace(S_ListToReconcile,Request("DocAuthor"),"")
		'vnik_error
	End If
	'Проверка полей с пользователями
	'Для заявок отключаем проверку утверждающего, т.к. он рассчитывается
	If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 Then
		UserChangeDocCheck = CheckSingleUserField(S_NameAproval)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameAproval")+SIT_ErrorInUserField2)
			Exit Function
		End If
		'Для заявки на закупку отключаем проверку ответственного, т.к. можно указывать подразделение
		'Для заявки на оплату тоже, т.к. автоматом ставится роль из специального справочника
		UserChangeDocCheck = CheckSingleUserField(S_NameResponsible)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameResponsible")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
	UserChangeDocCheck = CheckSingleUserField(S_NameControl)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameControl")+SIT_ErrorInUserField2)
		Exit Function
	End If
	If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) = 0 Then 'В исходящих может быть любой текст
		UserChangeDocCheck = CheckMultiUserField(S_Correspondent)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("Correspondent")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
'Запрос №1 - СИБ - start
	IF InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_SLUZH_ZAPISKA)) = 0) THEN
'Запрос №1 - СИБ - end
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) = 0 Then 'В СЗ на переработки работают правила
			UserChangeDocCheck = CheckMultiUserField(S_ListToReconcile)
			If not UserChangeDocCheck Then
				Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToReconcile")+SIT_ErrorInUserField2)
				Exit Function
			End If
'Запрос №1 - СИБ - start
		End If
	END IF
'Запрос №1 - СИБ - end
	UserChangeDocCheck = CheckMultiUserField(S_ListToEdit)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToEdit")+SIT_ErrorInUserField2)
		Exit Function
	End If
	'UserChangeDocCheck = CheckMultiUserField(S_ListToView)
	'If not UserChangeDocCheck Then
	'	Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToView")+SIT_ErrorInUserField2)
	'	Exit Function
	'End If

'{ph - 20120812
	'S_AdditionalUsers = GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit"))
	S_AdditionalUsers = DeleteUserDoublesInListNew(Request("AdditionalUsers") & VbCrLf & GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit")), True)
'ph - 20120812}

	AddLogD "@@@Search For Leaders - S_AdditionalUsers: "+S_AdditionalUsers
	'vnik_payment_order
	'If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_PAYMENT_ORDER)) > 0 Then
	'	S_AdditionalUsers = S_AdditionalUsers + ReplaceRoleFromDir(SIT_Account_manager,sDepartmentRoot)       
	'End If        
	'vnik_payment_order
	S_AdditionalUsers = DeleteUserDoublesInList(S_AdditionalUsers)

	'vnik
	If S_AdditionalUsers = "" Then
		S_AdditionalUsers = "-"
	End If
	S_AdditionalUsers_Set = S_AdditionalUsers

	S_ListToReconcile_Set = S_ListToReconcile
	S_NameAproval_Set = S_NameAproval
	S_ListToEdit = S_ListToEdit
	S_NameResponsible_Set = S_NameResponsible
	S_NameControl_Set = S_NameControl
	S_ListToView_Set=S_ListToView

	If UCase(Request("UpdateDoc")) = "YES" And (Request("create") = "y"  Or Request("DocIDPrev") <> Request("DocID")) Then
		'out "123"
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		dsTemp.Open "select * from Docs where DocID='" + MakeSQLSafe(Request("DocID")) + "'", Conn 
		If Not dsTemp.EOF Then
			Session("Message") = DOCS_ALREADYEXISTS + ": " + Request("DocID")
			dsTemp.Close
			UserChangeDocCheck=False
			Exit Function
		End If
		dsTemp.Close
	End If

	'ph - 20090808 - start
	If Request("DocIDParent") <> "" Then
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		bMyDocIDParentChanged = False
		If Request("create") <> "y" Then
			'Проверяем изменился ли номер родительского документа, если идет редактирование
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocID"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				bMyDocIDParentChanged = Request("DocIDParent")<>dsTemp("DocIDParent")
			End If
			dsTemp.Close
		End If
		If bMyDocIDParentChanged Then
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocIDParent"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				'sMessage = ""
				'UserChangeDocCheck = oPayDox.IsReadAccessUser(Session("UserID"), sMessage, dsTemp, DOCS_All, DOCS_NoReadAccess, DOCS_NoAccess, USER_Department, VAR_ExtInt, VAR_AdminSecLevel, VAR_StatusActiveUser, NULL)
				UserChangeDocCheck = IsReadAccessRS(dsTemp)' IsReadAccessUser(Session("UserID"), dsTemp, NULL)
				If UserChangeDocCheck Then
					S_DocIDParent = dsTemp("DocID")
				End If
				dsTemp.Close
				If not UserChangeDocCheck Then
					Session("Message") = SIT_NoAccessToParentDoc'sMessage
					Exit Function
				End If
			End If
		End If
	End If
	'ph - 20090808 - end

	if sDepartmentRoot = SIT_SITRU then
		'Проверка числовых значений в полях 
		'* Входящие:
		'* Количество листов (NameUserFieldText5)
		'* Количество листов в приложении (NameUserFieldText6)
		If Not (IsNumeric("0" & Request("UserFieldText5")) and IsNumeric("0" & Request("UserFieldText6"))) Then
			Session("Message") = "Поля ""Количество листов"" и ""Количество листов в приложении"" являются числовыми. Проверьте правильность заполнения."
			UserChangeDocCheck = False
			Exit Function
		End If
	end if

' *** ВХОДЯЩИЕ ДЛЯ БУХГАЛТЕРИИ
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_VHODYASCHIE_ACC)) = 1 Then
	'Создатель документа
	If UCase(Request("create")) = "Y" Then
		sDocCreator = Session("UserID")
	Else
		sDocCreator = GetUserID(Request("DocAuthor"))
	End If

	S_NameAproval=Request("DocNameAproval")
	S_NameResponsible=Request("DocNameResponsible")
	S_NameControl=Request("DocNameControl")
	S_ListToReconcile=Request("DocListToReconcile")
	S_ListToEdit=Request("DocListToEdit")
	S_ListToView=Request("DocListToView")
	S_Correspondent=Request("DocCorrespondent")

	'Определяем бизнес направление
	sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))

	' SAY 2008-07-21
	Set dsTempPR = Server.CreateObject("ADODB.Recordset")

' DmGorsky_5: меняем местами проверку: вначале ищем SIT_SITRU ("СИТРОНИКС ИТ"), потом, если не найдено, - SIT_SITRONICS ("СИТРОНИКС")
	If InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRU))=1 then ' DmGorsky
		sDepartment = SIT_SITRU ' DmGorsky
		sDepartmentCode = "" ' DmGorsky
'Запрос №1 - СИБ - start
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SIB_ROOT_DEPARTMENT))=1 then
		sDepartment = SIT_SIB
		sDepartmentCode = ""
'Запрос №1 - СИБ - end
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRONICS))=1 then
		sDepartment = SIT_SITRONICS
		sDepartmentCode = ""
	Else
		sDepartment = SIT_STS

		sSQL = "select code from Departments where Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"' or Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"/'"
		AddLogD "WAY sSQL="+sSQL
		AddLogD "WAY Conn="+Conn
		dsTempPR.Open sSQL, Conn, 3, 1, &H1

		if not dsTempPR.EOF Then
			sDepartmentCode = dsTempPR("code")
			if InStrRev(sDepartmentCode, "/") > 0 then
				sDepartmentCode = Right(sDepartmentCode, Len(sDepartmentCode)-InStrRev(sDepartmentCode, "/"))
			End If
		Else
			sDepartmentCode="000000"
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInDepartmentCode1 + DelOtherLangFromFolder(Request("DocDepartment")) + SIT_ErrorInDepartmentCode2)
			UserChangeDocCheck=False
			Exit Function
		End If

		dsTempPR.Close
	End If

	If Request("create") = "y" Then
		sPostfix = ""
		sSearchCol = "DocID"
		sPrePrefix = ""
		'SAY 2009-02-20
		sSufix = ""

		'SAY 2009-03-12
		S_DocID = ""

		If InStr(UCase(S_ClassDoc), UCase(SIT_VHODYASCHIE_ACC)) > 0 Then
			Call GetNewDocID_test(S_ClassDoc, sDepartmentRoot, "", "",  "", "")
		End If
		S_DocIDAdd = S_DocID
	End If
	If sDepartmentRoot <> SIT_STS or (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_RASP_DOCS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1) Then
		'реализация полужесткого маршрута
		S_ListToReconcile = Replace(S_ListToReconcile, Trim(SIT_AdditionalAgrees), "")
		'Для коммерческих предложений и заявок на закупку и оплату не нужна приписка про доп. согласующих  
		If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_COM_OFFERS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) <> 1 Then
			If S_ListToReconcile <> "" Then
				S_ListToReconcile = " " + SIT_AdditionalAgrees + DeleteConstPrefixFromList(S_ListToReconcile)
			End If
		End If
	End If
	'Для всех категорий кроме заявок подставляем пользователей вместо ролей
	If (InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1) and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) <> 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_DOGOVORI_NEW)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIS)) <> 0) Then
		'Получаем список ролей
		If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) = 1 Then 'Для договоров получаем список ролей на всех языках (могут быть роли на разных языках)
			sRoleList = GetFullRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"), "", "", "")
		Else
			sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))
		End If
		'Подставляем роли в поля
		S_NameAproval = ReplaceRolesInList(S_NameAproval, sRoleList)
		S_NameResponsible = ReplaceRolesInList(S_NameResponsible, sRoleList)
		S_NameControl = ReplaceRolesInList(S_NameControl, sRoleList)
		S_ListToEdit = ReplaceRolesInList(S_ListToEdit, sRoleList)
		S_ListToView = ReplaceRolesInList(S_ListToView, sRoleList)
		S_ListToReconcile = ReplaceRolesInList(S_ListToReconcile, sRoleList)

		'Удаляем повторы в списке согласования
		S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)

		'vnik_error
		'Убираем автора документа из списка согласующих, т.к. после активации он не сможет загрузить новый документ
		'(работает пока только, если пользователь создает документ не под ролью)
		S_ListToReconcile = Replace(S_ListToReconcile,Request("DocAuthor"),"")
		'vnik_error
	End If
	'Проверка полей с пользователями
	'Для заявок отключаем проверку утверждающего, т.к. он рассчитывается
	If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 Then
		UserChangeDocCheck = CheckSingleUserField(S_NameAproval)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameAproval")+SIT_ErrorInUserField2)
			Exit Function
		End If
		'Для заявки на закупку отключаем проверку ответственного, т.к. можно указывать подразделение
		'Для заявки на оплату тоже, т.к. автоматом ставится роль из специального справочника
		UserChangeDocCheck = CheckSingleUserField(S_NameResponsible)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameResponsible")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
	UserChangeDocCheck = CheckSingleUserField(S_NameControl)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameControl")+SIT_ErrorInUserField2)
		Exit Function
	End If
	If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) = 0 Then 'В исходящих может быть любой текст
		UserChangeDocCheck = CheckMultiUserField(S_Correspondent)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("Correspondent")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
'Запрос №1 - СИБ - start
	IF InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_SLUZH_ZAPISKA)) = 0) THEN
'Запрос №1 - СИБ - end
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) = 0 Then 'В СЗ на переработки работают правила
			UserChangeDocCheck = CheckMultiUserField(S_ListToReconcile)
			If not UserChangeDocCheck Then
				Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToReconcile")+SIT_ErrorInUserField2)
				Exit Function
			End If
'Запрос №1 - СИБ - start
		End If
	END IF
'Запрос №1 - СИБ - end
	UserChangeDocCheck = CheckMultiUserField(S_ListToEdit)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToEdit")+SIT_ErrorInUserField2)
		Exit Function
	End If
	'UserChangeDocCheck = CheckMultiUserField(S_ListToView)
	'If not UserChangeDocCheck Then
	'	Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToView")+SIT_ErrorInUserField2)
	'	Exit Function
	'End If

'{ph - 20120812
	'S_AdditionalUsers = GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit"))
	S_AdditionalUsers = DeleteUserDoublesInListNew(Request("AdditionalUsers") & VbCrLf & GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit")), True)
'ph - 20120812}
	AddLogD "@@@Search For Leaders - S_AdditionalUsers: "+S_AdditionalUsers
	'vnik_payment_order
	'If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_PAYMENT_ORDER)) > 0 Then
	'	S_AdditionalUsers = S_AdditionalUsers + ReplaceRoleFromDir(SIT_Account_manager,sDepartmentRoot)       
	'End If        
	'vnik_payment_order
	S_AdditionalUsers = DeleteUserDoublesInList(S_AdditionalUsers)

	'vnik
	If S_AdditionalUsers = "" Then
		S_AdditionalUsers = "-"
	End If
	S_AdditionalUsers_Set = S_AdditionalUsers

	S_ListToReconcile_Set = S_ListToReconcile
	S_NameAproval_Set = S_NameAproval
	S_ListToEdit = S_ListToEdit
	S_NameResponsible_Set = S_NameResponsible
	S_NameControl_Set = S_NameControl
	S_ListToView_Set=S_ListToView

	If UCase(Request("UpdateDoc")) = "YES" And (Request("create") = "y"  Or Request("DocIDPrev") <> Request("DocID")) Then
		'out "123"
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		dsTemp.Open "select * from Docs where DocID='" + MakeSQLSafe(Request("DocID")) + "'", Conn 
		If Not dsTemp.EOF Then
			Session("Message") = DOCS_ALREADYEXISTS + ": " + Request("DocID")
			dsTemp.Close
			UserChangeDocCheck=False
			Exit Function
		End If
		dsTemp.Close
	End If

	'ph - 20090808 - start
	If Request("DocIDParent") <> "" Then
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		bMyDocIDParentChanged = False
		If Request("create") <> "y" Then
			'Проверяем изменился ли номер родительского документа, если идет редактирование
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocID"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				bMyDocIDParentChanged = Request("DocIDParent")<>dsTemp("DocIDParent")
			End If
			dsTemp.Close
		End If
		If bMyDocIDParentChanged Then
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocIDParent"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				'sMessage = ""
				'UserChangeDocCheck = oPayDox.IsReadAccessUser(Session("UserID"), sMessage, dsTemp, DOCS_All, DOCS_NoReadAccess, DOCS_NoAccess, USER_Department, VAR_ExtInt, VAR_AdminSecLevel, VAR_StatusActiveUser, NULL)
				UserChangeDocCheck = IsReadAccessRS(dsTemp)' IsReadAccessUser(Session("UserID"), dsTemp, NULL)
				If UserChangeDocCheck Then
					S_DocIDParent = dsTemp("DocID")
				End If
				dsTemp.Close
				If not UserChangeDocCheck Then
					Session("Message") = SIT_NoAccessToParentDoc'sMessage
					Exit Function
				End If
			End If
		End If
	End If
	'ph - 20090808 - end

' *** ИСХОДЯЩИЕ ДОКУМЕНТЫ
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_ISHODYASCHIE)) = 1 Then
	'Создатель документа
	If UCase(Request("create")) = "Y" Then
		sDocCreator = Session("UserID")
	Else
		sDocCreator = GetUserID(Request("DocAuthor"))
	End If

	S_NameAproval=Request("DocNameAproval")
	S_NameResponsible=Request("DocNameResponsible")
	S_NameControl=Request("DocNameControl")
	S_ListToReconcile=Request("DocListToReconcile")
	S_ListToEdit=Request("DocListToEdit")
	S_ListToView=Request("DocListToView")
	S_Correspondent=Request("DocCorrespondent")

	'Определяем бизнес направление
	sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))

	' SAY 2008-07-21
	Set dsTempPR = Server.CreateObject("ADODB.Recordset")

' DmGorsky_5: меняем местами проверку: вначале ищем SIT_SITRU ("СИТРОНИКС ИТ"), потом, если не найдено, - SIT_SITRONICS ("СИТРОНИКС")
	If InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRU))=1 then ' DmGorsky
		sDepartment = SIT_SITRU ' DmGorsky
		sDepartmentCode = "" ' DmGorsky
'Запрос №1 - СИБ - start
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SIB_ROOT_DEPARTMENT))=1 then
		sDepartment = SIT_SIB
		sDepartmentCode = ""
'Запрос №1 - СИБ - end
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRONICS))=1 then
		sDepartment = SIT_SITRONICS
		sDepartmentCode = ""
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_MINC))=1 then
		sDepartment = SIT_MINC
		sDepartmentCode = ""		
	Else
		sDepartment = SIT_STS

		sSQL = "select code from Departments where Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"' or Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"/'"
		AddLogD "WAY sSQL="+sSQL
		AddLogD "WAY Conn="+Conn
		dsTempPR.Open sSQL, Conn, 3, 1, &H1

		if not dsTempPR.EOF Then
			sDepartmentCode = dsTempPR("code")
			if InStrRev(sDepartmentCode, "/") > 0 then
				sDepartmentCode = Right(sDepartmentCode, Len(sDepartmentCode)-InStrRev(sDepartmentCode, "/"))
			End If
		Else
			sDepartmentCode="000000"
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInDepartmentCode1 + DelOtherLangFromFolder(Request("DocDepartment")) + SIT_ErrorInDepartmentCode2)
			UserChangeDocCheck=False
			Exit Function
		End If

		dsTempPR.Close
	End If

	If Request("create") = "y" Then
		sPostfix = ""
		sSearchCol = "DocID"
		sPrePrefix = ""
		'SAY 2009-02-20
		sSufix = ""

		'SAY 2009-03-12
		S_DocID = ""

		if sDepartment = SIT_SITRU then ' DmGorsky
			S_DocID = GetNewDocIDForIshodyashie(S_ClassDoc, sDepartmentRoot, sDepartmentCode, Request("UserFieldText7"), Request("DocIDParent"), "", GetDZKCode(GetCodeFromCode_NameString(Request("BusinessUnit")), "LetterCode")) ' DmGorsky
		else ' DmGorsky
			S_DocID = GetNewDocIDForIshodyashie(S_ClassDoc, sDepartmentRoot, sDepartmentCode, Request("UserFieldText7"), Request("DocIDParent"), "PJ-", GetDZKCode(GetCodeFromCode_NameString(Request("BusinessUnit")), "LetterCode"))
		end if ' DmGorsky
		S_DocIDAdd = S_DocID
	End If
	If sDepartmentRoot <> SIT_STS or _
	   (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) <> 1 and _
	    InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_RASP_DOCS)) <> 1 and _
	    InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1) Then
		'реализация полужесткого маршрута
		S_ListToReconcile = Replace(S_ListToReconcile, Trim(SIT_AdditionalAgrees), "")
		'Для коммерческих предложений и заявок на закупку и оплату не нужна приписка про доп. согласующих  
		If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_COM_OFFERS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) <> 1 Then
			If S_ListToReconcile <> "" Then
				S_ListToReconcile = " " + SIT_AdditionalAgrees + DeleteConstPrefixFromList(S_ListToReconcile)
			End If
		End If
	End If
If sDepartmentRoot = SIT_SITRU Then ' DmGorsky
		S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSITRU","Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile ' DmGorsky
	
	ElseIf sDepartmentRoot = SIT_STS Then
'{Запрос №50 - СТС
		'GetReconciliationByRules(parDocID, parClassDoc, parAmount, parChartOfAccounts, parCostCenter, parProjectCode, parBusinessUnit, parPartnerName, parKindOfPayment, parInitiator, parProjectManager, parOvertimeRequester, parIncomeExpenceContract, parContranctInSTS, parCurrency, parContractType, parOvertimeFuncLeaders, ByRef parListToReconcile, ByRef parNameAproval, ByRef parNameResponsible, ByRef parCorrespondent, ByRef parListToView, ByRef parRegistrar, ByRef parNameControl, ByRef parListToEdit)
		'GetReconciliationByRules iif(Trim(Request("DocID")) = "", S_DocID, Request("DocID")), Session("CurrentClassDoc"), 0, "", "", "", GetCodeFromCode_NameString(Request("BusinessUnit")), S_PartnerName, "", iif(UCase(Request("create")) = "Y", Session("UserID"), Request("DocAuthor")), "", "", NULL, NULL, NULL, NULL, NULL, S_ListToReconcile, S_NameAproval, Null, S_Correspondent, Null, S_LocationPath, Null, Null
		par_DocID = iif(Trim(Request("DocID")) = "", S_DocID, Request("DocID"))
		par_ClassDoc = Session("CurrentClassDoc")
		par_Amount = 0
		par_ChartOfAccounts = ""
		par_CostCenter = ""
		par_ProjectCode = ""
		par_BusinessUnit = GetCodeFromCode_NameString(Request("BusinessUnit"))
		par_PartnerName = S_PartnerName
		par_KindOfPayment = ""
		par_Initiator = iif(UCase(Request("create")) = "Y", Session("UserID"), Request("DocAuthor"))
		par_ProjectManager = ""
		par_OvertimeRequester = ""
		par_IncomeExpenceContract = NULL
		par_ContranctInSTS = NULL
		par_Currency = NULL
		par_ContractType = NULL
		par_OvertimeFuncLeaders = NULL
		par_TypeOfDocument = ""
		par_FunctionArea = ""
       GetReconciliationByRules par_DocID, par_ClassDoc, par_Amount, par_ChartOfAccounts, par_CostCenter, _
            par_ProjectCode, par_BusinessUnit, par_PartnerName, par_KindOfPayment, par_Initiator, _
            par_ProjectManager, par_OvertimeRequester, par_IncomeExpenceContract, par_ContranctInSTS, _
            par_Currency, par_ContractType, par_OvertimeFuncLeaders, par_TypeOfDocument, par_FunctionArea, _
			S_ListToReconcile, S_NameAproval, Null, S_Correspondent, Null, S_LocationPath, Null, Null
		'	par_ListToReconcile, par_NameAproval, par_NameResponsible, par_Correspondent, par_ListToView, par_Registrar, par_NameControl, par_ListToEdit
'Запрос №50 - СТС}
		S_LocationPath_Set = S_LocationPath
	ElseIf sDepartmentRoot = SIT_SIB Then 'СИБ
		S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSIBRU","Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
	ElseIf sDepartmentRoot = SIT_SITRONICS Then
		S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
    ElseIf sDepartmentRoot = SIT_RTI Then
        S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeRTIRU","Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
    ElseIf sDepartmentRoot = SIT_MINC Then
        S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeMINCRU","Category",Session("CurrentClassDoc"),"List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
    ElseIf sDepartmentRoot = SIT_VTSS Then
        S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeVTSSRU","Category",Session("CurrentClassDoc"),"List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
    ElseIf sDepartmentRoot = SIT_MIKRON Then
        S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeMIKRON","Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
	Else 'Другие БН
		S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
	End If
	'Для всех категорий кроме заявок подставляем пользователей вместо ролей
	If (InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1) and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) <> 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_DOGOVORI_NEW)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIS)) <> 0) Then
		'Получаем список ролей
		If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) = 1 Then 'Для договоров получаем список ролей на всех языках (могут быть роли на разных языках)
			sRoleList = GetFullRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"), "", "", "")
		Else
			sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))
		End If

		
		'Подставляем роли в поля
		S_NameAproval = ReplaceRolesInList(S_NameAproval, sRoleList)
		S_NameResponsible = ReplaceRolesInList(S_NameResponsible, sRoleList)
		S_NameControl = ReplaceRolesInList(S_NameControl, sRoleList)
		S_ListToEdit = ReplaceRolesInList(S_ListToEdit, sRoleList)
		S_ListToView = ReplaceRolesInList(S_ListToView, sRoleList)
		S_ListToReconcile = ReplaceRolesInList(S_ListToReconcile, sRoleList)

        S_MINC_Director = ReplaceRolesInList(MINC_Director, sRoleList)
        
        'для минца в исходящих убираем первого согласующего, если он совпадает с подписантом.
         If sDepartmentRoot = SIT_MINC Then
            S_ListToReconcile = Replace(S_ListToReconcile,S_NameAproval,"")
                    'для минца всегда убираем генерального директора из списка согласующих'
            S_ListToReconcile = Replace(S_ListToReconcile,S_MINC_Director,"")
                    'для минца всегда убираем генерального директора из списка согласующих'

         End If
        'для минца в исходящих убираем первого согласующего, если он совпадает с подписантом.

		'Удаляем повторы в списке согласования
		S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)

		'vnik_error
		'Убираем автора документа из списка согласующих, т.к. после активации он не сможет загрузить новый документ
		'(работает пока только, если пользователь создает документ не под ролью)
		S_ListToReconcile = Replace(S_ListToReconcile,Request("DocAuthor"),"")
		'vnik_error
	End If
	'Проверка полей с пользователями
	'Для заявок отключаем проверку утверждающего, т.к. он рассчитывается
	If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 Then
		UserChangeDocCheck = CheckSingleUserField(S_NameAproval)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameAproval")+SIT_ErrorInUserField2)
			Exit Function
		End If
		'Для заявки на закупку отключаем проверку ответственного, т.к. можно указывать подразделение
		'Для заявки на оплату тоже, т.к. автоматом ставится роль из специального справочника
		UserChangeDocCheck = CheckSingleUserField(S_NameResponsible)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameResponsible")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
	UserChangeDocCheck = CheckSingleUserField(S_NameControl)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameControl")+SIT_ErrorInUserField2)
		Exit Function
	End If
	If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) = 0 Then 'В исходящих может быть любой текст
		UserChangeDocCheck = CheckMultiUserField(S_Correspondent)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("Correspondent")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
'Запрос №1 - СИБ - start
	IF InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_SLUZH_ZAPISKA)) = 0) THEN
'Запрос №1 - СИБ - end
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) = 0 Then 'В СЗ на переработки работают правила
			UserChangeDocCheck = CheckMultiUserField(S_ListToReconcile)
			If not UserChangeDocCheck Then
				Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToReconcile")+SIT_ErrorInUserField2)
				Exit Function
			End If
'Запрос №1 - СИБ - start
		End If
	END IF
'Запрос №1 - СИБ - end
	UserChangeDocCheck = CheckMultiUserField(S_ListToEdit)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToEdit")+SIT_ErrorInUserField2)
		Exit Function
	End If
	'UserChangeDocCheck = CheckMultiUserField(S_ListToView)
	'If not UserChangeDocCheck Then
	'	Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToView")+SIT_ErrorInUserField2)
	'	Exit Function
	'End If

'{ph - 20120812
	'S_AdditionalUsers = GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit"))
	S_AdditionalUsers = DeleteUserDoublesInListNew(Request("AdditionalUsers") & VbCrLf & GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit")), True)
'ph - 20120812}
	AddLogD "@@@Search For Leaders - S_AdditionalUsers: "+S_AdditionalUsers
	'vnik_payment_order
	'If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_PAYMENT_ORDER)) > 0 Then
	'	S_AdditionalUsers = S_AdditionalUsers + ReplaceRoleFromDir(SIT_Account_manager,sDepartmentRoot)       
	'End If        
	'vnik_payment_order
	S_AdditionalUsers = DeleteUserDoublesInList(S_AdditionalUsers)

	'vnik
	If S_AdditionalUsers = "" Then
		S_AdditionalUsers = "-"
	End If
	S_AdditionalUsers_Set = S_AdditionalUsers

	S_ListToReconcile_Set = S_ListToReconcile
	S_NameAproval_Set = S_NameAproval
	S_ListToEdit = S_ListToEdit
	S_NameResponsible_Set = S_NameResponsible
	S_NameControl_Set = S_NameControl
	S_ListToView_Set=S_ListToView

	If UCase(Request("UpdateDoc")) = "YES" And (Request("create") = "y"  Or Request("DocIDPrev") <> Request("DocID")) Then
		'out "123"
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		dsTemp.Open "select * from Docs where DocID='" + MakeSQLSafe(Request("DocID")) + "'", Conn 
		If Not dsTemp.EOF Then
			Session("Message") = DOCS_ALREADYEXISTS + ": " + Request("DocID")
			dsTemp.Close
			UserChangeDocCheck=False
			Exit Function
		End If
		dsTemp.Close
	End If

	'ph - 20090808 - start
	If Request("DocIDParent") <> "" Then
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		bMyDocIDParentChanged = False
		If Request("create") <> "y" Then
			'Проверяем изменился ли номер родительского документа, если идет редактирование
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocID"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				bMyDocIDParentChanged = Request("DocIDParent")<>dsTemp("DocIDParent")
			End If
			dsTemp.Close
		End If
		If bMyDocIDParentChanged Then
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocIDParent"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				'sMessage = ""
				'UserChangeDocCheck = oPayDox.IsReadAccessUser(Session("UserID"), sMessage, dsTemp, DOCS_All, DOCS_NoReadAccess, DOCS_NoAccess, USER_Department, VAR_ExtInt, VAR_AdminSecLevel, VAR_StatusActiveUser, NULL)
				UserChangeDocCheck = IsReadAccessRS(dsTemp)' IsReadAccessUser(Session("UserID"), dsTemp, NULL)
				If UserChangeDocCheck Then
					S_DocIDParent = dsTemp("DocID")
				End If
				dsTemp.Close
				If not UserChangeDocCheck Then
					Session("Message") = SIT_NoAccessToParentDoc'sMessage
					Exit Function
				End If
			End If
		End If
	End If
	'ph - 20090808 - end

	if sDepartmentRoot = SIT_SITRU then
		'Проверка числовых значений в полях 
		'* Исходящие:
		'* Количество листов (NameUserFieldText3)
		'* Количество листов в приложении (NameUserFieldText4)
		If Not (IsNumeric("0" & Request("UserFieldText3")) and IsNumeric("0" & Request("UserFieldText4"))) Then
			Session("Message") = "Поля ""Количество листов"" и ""Количество листов в приложении"" являются числовыми. Проверьте правильность заполнения."
			UserChangeDocCheck = False
			Exit Function
		End If
	end if

' *** РАСПОРЯДИТЕЛЬНЫЕ ДОКУМЕНТЫ
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_RASP_DOCS)) = 1 Then
	'Создатель документа
	If UCase(Request("create")) = "Y" Then
		sDocCreator = Session("UserID")
	Else
		sDocCreator = GetUserID(Request("DocAuthor"))
	End If

	S_NameAproval=Request("DocNameAproval")
	S_NameResponsible=Request("DocNameResponsible")
	S_NameControl=Request("DocNameControl")
	S_ListToReconcile=Request("DocListToReconcile")
	S_ListToEdit=Request("DocListToEdit")
	S_ListToView=Request("DocListToView")
	S_Correspondent=Request("DocCorrespondent")

	'Определяем бизнес направление
	sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))

	' SAY 2008-07-21
	Set dsTempPR = Server.CreateObject("ADODB.Recordset")

' DmGorsky_5: меняем местами проверку: вначале ищем SIT_SITRU ("СИТРОНИКС ИТ"), потом, если не найдено, - SIT_SITRONICS ("СИТРОНИКС")
	If InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRU))=1 then ' DmGorsky
		sDepartment = SIT_SITRU ' DmGorsky
		sDepartmentCode = "" ' DmGorsky
'Запрос №1 - СИБ - start
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SIB_ROOT_DEPARTMENT))=1 then
		sDepartment = SIT_SIB
		sDepartmentCode = ""
'Запрос №1 - СИБ - end
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRONICS))=1 then
		sDepartment = SIT_SITRONICS
		sDepartmentCode = ""
	Else
		sDepartment = SIT_STS

		sSQL = "select code from Departments where Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"' or Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"/'"
		AddLogD "WAY sSQL="+sSQL
		AddLogD "WAY Conn="+Conn
		dsTempPR.Open sSQL, Conn, 3, 1, &H1

		if not dsTempPR.EOF Then
			sDepartmentCode = dsTempPR("code")
			if InStrRev(sDepartmentCode, "/") > 0 then
				sDepartmentCode = Right(sDepartmentCode, Len(sDepartmentCode)-InStrRev(sDepartmentCode, "/"))
			End If
		Else
			sDepartmentCode="000000"
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInDepartmentCode1 + DelOtherLangFromFolder(Request("DocDepartment")) + SIT_ErrorInDepartmentCode2)
			UserChangeDocCheck=False
			Exit Function
		End If

		dsTempPR.Close
	End If
	
	If Request("create") = "y" Then
		sPostfix = ""
		sSearchCol = "DocID"
		sPrePrefix = ""
		'SAY 2009-02-20
		sSufix = ""

		'SAY 2009-03-12
		S_DocID = ""

		if sDepartment = SIT_SITRU then ' DmGorsky
			S_DocID = GetNewDocIDForRaspDocs(S_ClassDoc, sDepartmentRoot, Request("UserFieldText1"), "", "", "", GetDZKCode(GetCodeFromCode_NameString(Request("BusinessUnit")), "OrderCode")) ' DmGorsky
		else ' DmGorsky
			S_DocID = GetNewDocIDForRaspDocs(S_ClassDoc, sDepartmentRoot, Request("UserFieldText1"), "", "", "PJ-", GetDZKCode(GetCodeFromCode_NameString(Request("BusinessUnit")), "OrderCode"))
		end if ' DmGorsky
		S_DocIDAdd = S_DocID
	End If
	If sDepartmentRoot <> SIT_STS or (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_RASP_DOCS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1) Then
		'реализация полужесткого маршрута
		S_ListToReconcile = Replace(S_ListToReconcile, Trim(SIT_AdditionalAgrees), "")
		'Для коммерческих предложений и заявок на закупку и оплату не нужна приписка про доп. согласующих  
		If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_COM_OFFERS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) <> 1 Then
			If S_ListToReconcile <> "" Then
				S_ListToReconcile = " " + SIT_AdditionalAgrees + DeleteConstPrefixFromList(S_ListToReconcile)
			End If
		End If
	End If
	AddLogD "vnikvnik 12" + Trim(S_ListToReconcile)
	AddLogD "vnikvnik 123" + Trim(S_ListToReconcile) 
	'vnik_rasp_norm_doc
	'проверка крайней даты, работает корректно только если срок согласования по умолчанию 3 дня
	If Trim(Request("UserFieldText1")) = SIT_Orders_Prikaz_ND Then
		' Получим дату активации/создания
		AddLogD "vnik654 1"
		Set VNIK_dsTemp = Server.CreateObject("ADODB.Recordset")
		sSQL = "select DateActive,DateLastModification from Docs where DocID = "+sUnicodeSymbol+"'" + Request("DocID") + "'"
		VNIK_dsTemp.Open sSQL, Conn, 3, 1, &H1
		AddLogD "vnik654 2"
		If not VNIK_dsTemp.EOF Then
			AddLogD "vnik654 3"
			AddLogD Trim(VNIK_dsTemp("DateActive"))
			AddLogD Trim(VNIK_dsTemp("DateLastModification"))
			AddLogD Request("UserFieldDate1")
			AddLogD "vnik654 4"
			s_DateCreateActive = VNIK_dsTemp("DateLastModification")
			VNIK_UserFieldDate1 = FormatDateTime(s_DateCreateActive,2)
		Else
			AddLogD "vnik654 6"
			s_DateCreateActive = Request("UserFieldDate1")
			VNIK_UserFieldDate1 = Right(Left(s_DateCreateActive,5),2) + "/" + Left(MyDate(s_DateCreateActive),2) + "/" + Right(MyDate(s_DateCreateActive),4)                   
		End If 
		' получили
		AddLogD "vnik654 7" + Trim(VNIK_UserFieldDate1)
		VNIK_WeekDay_Number = Trim(WeekDay(cdate(VNIK_UserFieldDate1)))
		If (VNIK_WeekDay_Number = 4) or (VNIK_WeekDay_Number = 5) or (VNIK_WeekDay_Number = 6) Then
			VNIK_UserFieldDate2 = MyDate(cdate(VNIK_UserFieldDate1)+5)
		ElseIf (VNIK_WeekDay_Number = 7) Then
			VNIK_UserFieldDate2 = MyDate(cdate(VNIK_UserFieldDate1)+4)
		Else
			VNIK_UserFieldDate2 = MyDate(cdate(VNIK_UserFieldDate1)+3)
			'AddLogD "vnik654 " + Trim(VNIK_UserFieldDate2)
		End If

		AddLogD "vnik654 " + Trim(VNIK_WeekDay_Number)
		AddLogD "vnik654 " + Trim(cdate(Replace(s_DateCreateActive,".","/")))
		AddLogD "vnik654 " + Trim(MyDate(Request("UserFieldDate2")))
		AddLogD "vnik654 " + Trim(Right(Left(MyDate(Request("UserFieldDate2")),5),2))
		AddLogD "vnik654 " + Trim(Right(Left(VNIK_UserFieldDate2,5),2))
		AddLogD "vnik654 " + Trim(Left(MyDate(Request("UserFieldDate2")),2))
		AddLogD "vnik654 " + Trim(VNIK_UserFieldDate2)
		If Right(Left(MyDate(Request("UserFieldDate2")),5),2) < Right(Left(VNIK_UserFieldDate2,5),2) Then 
			AddLogD "vnik654 " + Trim(SIT_NORM_DOCS_WARNING+VNIK_UserFieldDate2)
			Session("Message") = SIT_NORM_DOCS_WARNING+VNIK_UserFieldDate2
			UserChangeDocCheck = False
			Exit Function
		ElseIf Right(Left(MyDate(Request("UserFieldDate2")),5),2) = Right(Left(VNIK_UserFieldDate2,5),2) Then
			If Left(MyDate(Request("UserFieldDate2")),2) < Left(VNIK_UserFieldDate2,2) Then
				AddLogD "vnik6541 " + Trim(SIT_NORM_DOCS_WARNING+VNIK_UserFieldDate2)
				Session("Message") = SIT_NORM_DOCS_WARNING+VNIK_UserFieldDate2
				UserChangeDocCheck = False
				Exit Function    
			End If
		End If
	End If
	'vnik_rasp_norm_doc 

	If sDepartmentRoot = SIT_SITRU Then ' DmGorsky_9
    If Trim(Request("UserFieldText1")) = SIT_Orders_Prikaz_ND Then ' DmGorsky_9
    'выбор по категории нормативных документов ' DmGorsky_9
        S_ListToReconcile = oPayDox.GetExtTableValue("AgreeSITRU","Category","Нормативные документы*Regulations*Řídicí dokumenty/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile' DmGorsky_9
    Else ' DmGorsky_9
    	S_ListToReconcile = oPayDox.GetExtTableValue("AgreeSITRU","Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile' DmGorsky_9
    End If ' DmGorsky_9
    
    ElseIf sDepartmentRoot = SIT_RTI Then
       AddLogD "vnik0123 " + Trim(S_ListToReconcile)
       if Request("DocNameAproval") = RTI_President or Request("DocNameAproval") = """Боев С. Ф."" <boev_oaorti>;" Then
         If Request("create") = "y" Then
'           If InStr(UCase(Session("Department")), UCase(RTI_DVKiA)) > 0 Then
'              S_ListToReconcile = SIT_RequiredAgrees+RTI_HeadOfUpravDelami+RTI_DirectorOfSecurity+SIT_AdditionalAgreesDelimeter+S_ListToReconcile+VbCrLf+SIT_RTI_DirectorPravovogoUprav_RU+VbCrLf+SIT_RTI_DirectorUpravDelami_RU+VbCrLf+SIT_RTI_DirectorApparatGD_RU
'           else
              S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeRTIRU","Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile+VbCrLf+SIT_RTI_DirectorPravovogoUprav_RU+VbCrLf+SIT_RTI_DirectorUpravDelami_RU+VbCrLf+SIT_RTI_DirectorApparatGD_RU
'           End if
         Else
            AddLogD "vnik0123 " + Trim(S_ListToReconcile)
            sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))
            S_ListToReconcile = replace(S_ListToReconcile, ReplaceRolesInList(SIT_RTI_DirectorPravovogoUprav_RU, sRoleList), "")
            S_ListToReconcile = replace(S_ListToReconcile, VbCrLf+ReplaceRolesInList(SIT_RTI_DirectorUpravDelami_RU, sRoleList), "")
            S_ListToReconcile = replace(S_ListToReconcile, VbCrLf+ReplaceRolesInList(SIT_RTI_DirectorApparatGD_RU, sRoleList), "")
            S_ListToReconcile = replace(S_ListToReconcile, VbCrLf, "")
'            If InStr(UCase(Session("Department")), UCase(RTI_DVKiA)) > 0 Then
'              S_ListToReconcile = SIT_RequiredAgrees+RTI_HeadOfUpravDelami+RTI_DirectorOfSecurity+SIT_AdditionalAgreesDelimeter+S_ListToReconcile+VbCrLf+SIT_RTI_DirectorPravovogoUprav_RU+VbCrLf+SIT_RTI_DirectorUpravDelami_RU+VbCrLf+SIT_RTI_DirectorApparatGD_RU
'            else
            S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeRTIRU","Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile+VbCrLf+SIT_RTI_DirectorPravovogoUprav_RU+VbCrLf+SIT_RTI_DirectorUpravDelami_RU+VbCrLf+SIT_RTI_DirectorApparatGD_RU
'            End if            
         End If
         S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)
       else
         If Request("create") = "y" Then

'           If InStr(UCase(Session("Department")), UCase(RTI_DVKiA)) > 0 Then
'              S_ListToReconcile = SIT_RequiredAgrees+RTI_HeadOfUpravDelami+RTI_DirectorOfSecurity+SIT_AdditionalAgreesDelimeter+S_ListToReconcile+VbCrLf+SIT_RTI_DirectorPravovogoUprav_RU+VbCrLf+SIT_RTI_DirectorUpravDelami_RU
'           else
              S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeRTIRU","Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile +VbCrLf+SIT_RTI_DirectorPravovogoUprav_RU+VbCrLf+SIT_RTI_DirectorUpravDelami_RU
'           End if

	     Else      
	        sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))
            S_ListToReconcile = replace(S_ListToReconcile, ReplaceRolesInList(SIT_RTI_DirectorPravovogoUprav_RU, sRoleList), "")
            S_ListToReconcile = replace(S_ListToReconcile, VbCrLf+ReplaceRolesInList(SIT_RTI_DirectorUpravDelami_RU, sRoleList), "")
            S_ListToReconcile = replace(S_ListToReconcile, VbCrLf, "")
'           If InStr(UCase(Session("Department")), UCase(RTI_DVKiA)) > 0 Then
'              S_ListToReconcile = SIT_RequiredAgrees+RTI_HeadOfUpravDelami+RTI_DirectorOfSecurity+SIT_AdditionalAgreesDelimeter+S_ListToReconcile+VbCrLf+SIT_RTI_DirectorPravovogoUprav_RU+VbCrLf+SIT_RTI_DirectorUpravDelami_RU
'           else
              S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeRTIRU","Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile +VbCrLf+SIT_RTI_DirectorPravovogoUprav_RU+VbCrLf+SIT_RTI_DirectorUpravDelami_RU
'           End if
         End If
         S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)
	   end if

	   S_ListToView = S_ListToView +" ;" + vbCrLf + RTI_RaspDocsViewList
	   	
	ElseIf sDepartmentRoot = SIT_SITRONICS Then
			'vnik_rasp_norm_doc
		AddLogD "vnikvnik 0" + Trim(S_ListToReconcile)
		If Trim(Request("UserFieldText1")) = SIT_Orders_Prikaz_ND Then
			'vnik здесь пишем выбор по категории нормативных документов
			S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category","Нормативные документы*Regulations*Řídicí dokumenty/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile+VbCrLf+SIT_VicePresidentOfInitiator
			S_ListToReconcile = Replace(S_ListToReconcile,"""Хачатуров К. К."" <kkhachaturov>;","")
			AddLogD "vnikvnik 1" + Trim(S_ListToReconcile)    
		Else
			AddLogD "vnikvnik 2" + Trim(S_ListToReconcile)
			S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile+VbCrLf+SIT_VicePresidentOfInitiator
			S_ListToReconcile = Replace(S_ListToReconcile,"""Хачатуров К. К."" <kkhachaturov>;","")
		End If
		AddLogD "vnikvnik 3" + Trim(S_ListToReconcile)
		'vnik_rasp_norm_doc
	ElseIf sDepartmentRoot = SIT_STS Then
'{Запрос №50 - СТС
		'GetReconciliationByRules(parDocID, parClassDoc, parAmount, parChartOfAccounts, parCostCenter, parProjectCode, parBusinessUnit, parPartnerName, parKindOfPayment, parInitiator, parProjectManager, parOvertimeRequester, parIncomeExpenceContract, parContranctInSTS, parCurrency, parContractType, parOvertimeFuncLeaders, ByRef parListToReconcile, ByRef parNameAproval, ByRef parNameResponsible, ByRef parCorrespondent, ByRef parListToView, ByRef parRegistrar, ByRef parNameControl, ByRef parListToEdit)
		'GetReconciliationByRules iif(Trim(Request("DocID")) = "", S_DocID, Request("DocID")), Session("CurrentClassDoc"), 0, "", "", "", GetCodeFromCode_NameString(Request("BusinessUnit")), S_PartnerName, "", iif(UCase(Request("create")) = "Y", Session("UserID"), Request("DocAuthor")), "", "", NULL, NULL, NULL, NULL, NULL, S_ListToReconcile, S_NameAproval, Null, S_Correspondent, Null, S_LocationPath, Null, Null
		par_DocID = iif(Trim(Request("DocID")) = "", S_DocID, Request("DocID"))
		par_ClassDoc = Session("CurrentClassDoc")
		par_Amount = 0
		par_ChartOfAccounts = ""
		par_CostCenter = ""
		par_ProjectCode = ""
		par_BusinessUnit = GetCodeFromCode_NameString(Request("BusinessUnit"))
		par_PartnerName = S_PartnerName
		par_KindOfPayment = ""
		par_Initiator = iif(UCase(Request("create")) = "Y", Session("UserID"), Request("DocAuthor"))
		par_ProjectManager = ""
		par_OvertimeRequester = ""
		par_IncomeExpenceContract = NULL
		par_ContranctInSTS = NULL
		par_Currency = NULL
		par_ContractType = NULL
		par_OvertimeFuncLeaders = NULL
		par_TypeOfDocument = Request("UserFieldText1")
		par_FunctionArea = Request("UserFieldText2")

		GetReconciliationByRules par_DocID, par_ClassDoc, par_Amount, par_ChartOfAccounts, par_CostCenter, par_ProjectCode, par_BusinessUnit, par_PartnerName, par_KindOfPayment, par_Initiator, par_ProjectManager, par_OvertimeRequester, par_IncomeExpenceContract, par_ContranctInSTS, par_Currency, par_ContractType, par_OvertimeFuncLeaders, par_TypeOfDocument, par_FunctionArea, _
			S_ListToReconcile, S_NameAproval, Null, S_Correspondent, Null, S_LocationPath, Null, Null
		'	par_ListToReconcile, par_NameAproval, par_NameResponsible, par_Correspondent, par_ListToView, par_Registrar, par_NameControl, par_ListToEdit
'Запрос №50 - СТС}
		S_LocationPath_Set = S_LocationPath
'Запрос №1 - СИБ - start
	ElseIf sDepartmentRoot = SIT_SIB Then
		If Trim(Request("UserFieldText1")) = SIT_Orders_Prikaz_ND Then
			S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSIBRU","Category","Нормативные документы*Regulations*Řídicí dokumenty/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
		Else
			S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSIBRU","Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
		End If
'Запрос №1 - СИБ - end
'Запрос МИКРОН - start
	ElseIf sDepartmentRoot = SIT_MIKRON Then
		If Trim(Request("UserFieldText1")) = SIT_Orders_Prikaz_ND Then
			S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeMIKRON","Category","Нормативные документы*Regulations*Řídicí dokumenty/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
		Else
			S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeMIKRON","Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
		End If
		S_ListToReconcile = Replace(S_ListToReconcile,VbCrLf,"")
'Запрос МИКРОН - end
	Else 'Другие БН
		S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile+VbCrLf+SIT_VicePresidentOfInitiator
	End If
	'Для всех категорий кроме заявок подставляем пользователей вместо ролей
	If (InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1) and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) <> 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_DOGOVORI_NEW)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIS)) <> 0) Then
		'Получаем список ролей
		If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) = 1 Then 'Для договоров получаем список ролей на всех языках (могут быть роли на разных языках)
			sRoleList = GetFullRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"), "", "", "")
		Else
			sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))
		End If
		'Подставляем роли в поля
		S_NameAproval = ReplaceRolesInList(S_NameAproval, sRoleList)
		S_NameResponsible = ReplaceRolesInList(S_NameResponsible, sRoleList)
		S_NameControl = ReplaceRolesInList(S_NameControl, sRoleList)
		S_ListToEdit = ReplaceRolesInList(S_ListToEdit, sRoleList)
		S_ListToView = ReplaceRolesInList(S_ListToView, sRoleList)
		S_ListToReconcile = ReplaceRolesInList(S_ListToReconcile, sRoleList)
		S_ListToReconcile = Replace(S_ListToReconcile,"""Хачатуров К. К."" <kkhachaturov>;","")

		'Удаляем повторы в списке согласования
		S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)
        S_ListToView = DeleteUserDoublesInList(S_ListToView)
		'vnik_error
		'Убираем автора документа из списка согласующих, т.к. после активации он не сможет загрузить новый документ
		'(работает пока только, если пользователь создает документ не под ролью)
		S_ListToReconcile = Replace(S_ListToReconcile,Request("DocAuthor"),"")
		'vnik_error
	End If
	'Проверка полей с пользователями
	'Для заявок отключаем проверку утверждающего, т.к. он рассчитывается
	If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 Then
		UserChangeDocCheck = CheckSingleUserField(S_NameAproval)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameAproval")+SIT_ErrorInUserField2)
			Exit Function
		End If
		'Для заявки на закупку отключаем проверку ответственного, т.к. можно указывать подразделение
		'Для заявки на оплату тоже, т.к. автоматом ставится роль из специального справочника
		UserChangeDocCheck = CheckSingleUserField(S_NameResponsible)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameResponsible")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
	UserChangeDocCheck = CheckSingleUserField(S_NameControl)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameControl")+SIT_ErrorInUserField2)
		Exit Function
	End If
	If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) = 0 Then 'В исходящих может быть любой текст
		UserChangeDocCheck = CheckMultiUserField(S_Correspondent)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("Correspondent")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
'Запрос №1 - СИБ - start
	IF InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_SLUZH_ZAPISKA)) = 0) THEN
'Запрос №1 - СИБ - end
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) = 0 Then 'В СЗ на переработки работают правила
			UserChangeDocCheck = CheckMultiUserField(S_ListToReconcile)
			If not UserChangeDocCheck Then
				Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToReconcile")+SIT_ErrorInUserField2)
				Exit Function
			End If
'Запрос №1 - СИБ - start
		End If
	END IF
'Запрос №1 - СИБ - end
	UserChangeDocCheck = CheckMultiUserField(S_ListToEdit)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToEdit")+SIT_ErrorInUserField2)
		Exit Function
	End If
	'UserChangeDocCheck = CheckMultiUserField(S_ListToView)
	'If not UserChangeDocCheck Then
	'	Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToView")+SIT_ErrorInUserField2)
	'	Exit Function
	'End If

'{ph - 20120812
	'S_AdditionalUsers = GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit"))
	S_AdditionalUsers = DeleteUserDoublesInListNew(Request("AdditionalUsers") & VbCrLf & GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit")), True)
'ph - 20120812}
	AddLogD "@@@Search For Leaders - S_AdditionalUsers: "+S_AdditionalUsers
	'vnik_payment_order
	'If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_PAYMENT_ORDER)) > 0 Then
	'	S_AdditionalUsers = S_AdditionalUsers + ReplaceRoleFromDir(SIT_Account_manager,sDepartmentRoot)       
	'End If        
	'vnik_payment_order
	S_AdditionalUsers = DeleteUserDoublesInList(S_AdditionalUsers)

	'vnik
	If S_AdditionalUsers = "" Then
		S_AdditionalUsers = "-"
	End If
	S_AdditionalUsers_Set = S_AdditionalUsers

	S_ListToReconcile_Set = S_ListToReconcile
	S_NameAproval_Set = S_NameAproval
	S_ListToEdit = S_ListToEdit
	S_NameResponsible_Set = S_NameResponsible
	S_NameControl_Set = S_NameControl
	S_ListToView_Set=S_ListToView

	If UCase(Request("UpdateDoc")) = "YES" And (Request("create") = "y"  Or Request("DocIDPrev") <> Request("DocID")) Then
		'out "123"
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		dsTemp.Open "select * from Docs where DocID='" + MakeSQLSafe(Request("DocID")) + "'", Conn 
		If Not dsTemp.EOF Then
			Session("Message") = DOCS_ALREADYEXISTS + ": " + Request("DocID")
			dsTemp.Close
			UserChangeDocCheck=False
			Exit Function
		End If
		dsTemp.Close
	End If

	'ph - 20090808 - start
	If Request("DocIDParent") <> "" Then
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		bMyDocIDParentChanged = False
		If Request("create") <> "y" Then
			'Проверяем изменился ли номер родительского документа, если идет редактирование
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocID"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				bMyDocIDParentChanged = Request("DocIDParent")<>dsTemp("DocIDParent")
			End If
			dsTemp.Close
		End If
		If bMyDocIDParentChanged Then
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocIDParent"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				'sMessage = ""
				'UserChangeDocCheck = oPayDox.IsReadAccessUser(Session("UserID"), sMessage, dsTemp, DOCS_All, DOCS_NoReadAccess, DOCS_NoAccess, USER_Department, VAR_ExtInt, VAR_AdminSecLevel, VAR_StatusActiveUser, NULL)
				UserChangeDocCheck = IsReadAccessRS(dsTemp)' IsReadAccessUser(Session("UserID"), dsTemp, NULL)
				If UserChangeDocCheck Then
					S_DocIDParent = dsTemp("DocID")
				End If
				dsTemp.Close
				If not UserChangeDocCheck Then
					Session("Message") = SIT_NoAccessToParentDoc'sMessage
					Exit Function
				End If
			End If
		End If
	End If
	'ph - 20090808 - end

' *** ПРОТОКОЛЫ

'rti_protocol
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(RTI_PROTOCOL)) = 1 Then
	'Создатель документа
	If UCase(Request("create")) = "Y" Then
		sDocCreator = Session("UserID")
	Else
		sDocCreator = GetUserID(Request("DocAuthor"))
	End If
          
	S_NameAproval=Request("DocNameAproval")
	S_NameResponsible=Request("DocNameResponsible")
	S_NameControl=Request("DocNameControl")
	S_ListToReconcile=Request("DocListToReconcile")
	S_ListToEdit=Request("DocListToEdit")
	S_ListToView=Request("DocListToView")
	S_Correspondent=Request("DocCorrespondent")

	'Определяем бизнес направление
	sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))

		sDepartment = SIT_RTI ' DmGorsky
		sDepartmentCode = "" ' DmGorsky


	If Request("create") = "y" Then
		sPostfix = ""
		sSearchCol = "DocID"
		sPrePrefix = ""
		'SAY 2009-02-20
		sSufix = ""

		'SAY 2009-03-12
		S_DocID = ""
		
		S_DocID = GetNewDocIDForPurchaseOrder(S_ClassDoc, sDepartmentRoot, "PRCPCRTI", "", "", "", "")
		S_DocIDAdd = S_DocID
		
	
	End If
	If sDepartmentRoot <> SIT_STS or (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_RASP_DOCS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1) Then
		'реализация полужесткого маршрута
		S_ListToReconcile = Replace(S_ListToReconcile, Trim(SIT_AdditionalAgrees), "")
		'Для коммерческих предложений и заявок на закупку и оплату не нужна приписка про доп. согласующих  
		If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_COM_OFFERS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) <> 1 Then
			If S_ListToReconcile <> "" Then
				S_ListToReconcile = " " + SIT_AdditionalAgrees + DeleteConstPrefixFromList(S_ListToReconcile)
			End If
		End If
	End If

			    'Проверяем документ основание - протокол можно создавать только на основе заявк ина закупку и только пользователем с ролью Секретарь ЦЗК	    

			'	Set vnikdsTemp1 = Server.CreateObject("ADODB.Recordset")
			'	sSQL = "Select * from Docs where DocID = "+sUnicodeSymbol+"'"+Request("DocIDParent")+"'" + " and ClassDoc = " + sUnicodeSymbol +"'"+RTI_PURCHASE_ORDER+"'"
			'	vnikdsTemp1.Open sSQL, Conn, 3, 1, &H1 
			'	If vnikdsTemp1.EOF Then
			'		Session("Message") = AddNewLineToMessage(Session("Message"), RTI_PROTOCOL_WARNING1)
			'		UserChangeDocCheck = False
			'		Exit Function
			'	End If   
			'	vnikdsTemp1.Close  

	S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeRTI"+Request("l"),"Category",Session("CurrentClassDoc"),"List")+S_ListToReconcile
		'Получаем список ролей
			sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))
		'Подставляем роли в поля
		S_NameAproval = ReplaceRolesInList(S_NameAproval, sRoleList)
		S_NameResponsible = ReplaceRolesInList(S_NameResponsible, sRoleList)
		S_NameControl = ReplaceRolesInList(S_NameControl, sRoleList)
		S_ListToEdit = ReplaceRolesInList(S_ListToEdit, sRoleList)
		S_ListToView = ReplaceRolesInList(S_ListToView, sRoleList)
		S_ListToReconcile = ReplaceRolesInList(S_ListToReconcile, sRoleList)

		'Удаляем повторы в списке согласования
		S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)

		'vnik_error
		'Убираем автора документа из списка согласующих, т.к. после активации он не сможет загрузить новый документ
		'(работает пока только, если пользователь создает документ не под ролью)
		S_ListToReconcile = Replace(S_ListToReconcile,Request("DocAuthor"),"")
		'vnik_error


		S_ListToView = DeleteUserDoublesInList(S_ListToView)    

	'Проверка полей с пользователями
	'Для заявок отключаем проверку утверждающего, т.к. он рассчитывается
	If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 Then
		UserChangeDocCheck = CheckSingleUserField(S_NameAproval)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameAproval")+SIT_ErrorInUserField2)
			Exit Function
		End If
		'Для заявки на закупку отключаем проверку ответственного, т.к. можно указывать подразделение
		'Для заявки на оплату тоже, т.к. автоматом ставится роль из специального справочника
		UserChangeDocCheck = CheckSingleUserField(S_NameResponsible)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameResponsible")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
	UserChangeDocCheck = CheckMultiUserField(S_ListToEdit)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToEdit")+SIT_ErrorInUserField2)
		Exit Function
	End If

	S_AdditionalUsers = DeleteUserDoublesInList(S_AdditionalUsers)

	'vnik
	If S_AdditionalUsers = "" Then
		S_AdditionalUsers = "-"
	End If
	S_AdditionalUsers_Set = S_AdditionalUsers

	S_ListToReconcile_Set = S_ListToReconcile
	S_NameAproval_Set = S_NameAproval
	S_ListToEdit = S_ListToEdit
	S_NameResponsible_Set = S_NameResponsible
	S_NameControl_Set = S_NameControl
	S_ListToView_Set=S_ListToView

	If UCase(Request("UpdateDoc")) = "YES" And (Request("create") = "y"  Or Request("DocIDPrev") <> Request("DocID")) Then
		'out "123"
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		dsTemp.Open "select * from Docs where DocID='" + MakeSQLSafe(Request("DocID")) + "'", Conn 
		If Not dsTemp.EOF Then
			Session("Message") = DOCS_ALREADYEXISTS + ": " + Request("DocID")
			dsTemp.Close
			UserChangeDocCheck=False
			Exit Function
		End If
		dsTemp.Close
	End If

	'ph - 20090808 - start
	If Request("DocIDParent") <> "" Then
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		bMyDocIDParentChanged = False
		If Request("create") <> "y" Then
			'Проверяем изменился ли номер родительского документа, если идет редактирование
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocID"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				bMyDocIDParentChanged = Request("DocIDParent")<>dsTemp("DocIDParent")
			End If
			dsTemp.Close
		End If
		If bMyDocIDParentChanged Then
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocIDParent"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				'sMessage = ""
				'UserChangeDocCheck = oPayDox.IsReadAccessUser(Session("UserID"), sMessage, dsTemp, DOCS_All, DOCS_NoReadAccess, DOCS_NoAccess, USER_Department, VAR_ExtInt, VAR_AdminSecLevel, VAR_StatusActiveUser, NULL)
				UserChangeDocCheck = IsReadAccessRS(dsTemp)' IsReadAccessUser(Session("UserID"), dsTemp, NULL)
				If UserChangeDocCheck Then
					S_DocIDParent = dsTemp("DocID")
				End If
				dsTemp.Close
				If not UserChangeDocCheck Then
					Session("Message") = SIT_NoAccessToParentDoc'sMessage
					Exit Function
				End If
			End If
		End If
	End If
	'ph - 20090808 - end
'rti_protocol

'vnik_protoclos_cpc
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_PROTOCOLS)) = 1 or InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_PROTOCOLS_CPC)) > 0 Then
	'Создатель документа
	If UCase(Request("create")) = "Y" Then
		sDocCreator = Session("UserID")
	Else
		sDocCreator = GetUserID(Request("DocAuthor"))
	End If
          
	S_NameAproval=Request("DocNameAproval")
	S_NameResponsible=Request("DocNameResponsible")
	S_NameControl=Request("DocNameControl")
	S_ListToReconcile=Request("DocListToReconcile")
	S_ListToEdit=Request("DocListToEdit")
	S_ListToView=Request("DocListToView")
	S_Correspondent=Request("DocCorrespondent")

	'Определяем бизнес направление
	sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))

	' SAY 2008-07-21
	Set dsTempPR = Server.CreateObject("ADODB.Recordset")

' DmGorsky_5: меняем местами проверку: вначале ищем SIT_SITRU ("СИТРОНИКС ИТ"), потом, если не найдено, - SIT_SITRONICS ("СИТРОНИКС")
	If InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRU))=1 then ' DmGorsky
		sDepartment = SIT_SITRU ' DmGorsky
		sDepartmentCode = "" ' DmGorsky
'Запрос №1 - СИБ - start
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SIB_ROOT_DEPARTMENT))=1 then
		sDepartment = SIT_SIB
		sDepartmentCode = ""
'Запрос №1 - СИБ - end
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRONICS))=1 then
		sDepartment = SIT_SITRONICS
		sDepartmentCode = ""
	Else
		sDepartment = SIT_STS

		sSQL = "select code from Departments where Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"' or Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"/'"
		AddLogD "WAY sSQL="+sSQL
		AddLogD "WAY Conn="+Conn
		dsTempPR.Open sSQL, Conn, 3, 1, &H1

		if not dsTempPR.EOF Then
			sDepartmentCode = dsTempPR("code")
			if InStrRev(sDepartmentCode, "/") > 0 then
				sDepartmentCode = Right(sDepartmentCode, Len(sDepartmentCode)-InStrRev(sDepartmentCode, "/"))
			End If
		Else
			sDepartmentCode="000000"
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInDepartmentCode1 + DelOtherLangFromFolder(Request("DocDepartment")) + SIT_ErrorInDepartmentCode2)
			UserChangeDocCheck=False
			Exit Function
		End If

		dsTempPR.Close
	End If

	If Request("create") = "y" Then
		sPostfix = ""
		sSearchCol = "DocID"
		sPrePrefix = ""
		'SAY 2009-02-20
		sSufix = ""

		'SAY 2009-03-12
		S_DocID = ""
		
		S_DocID = GetNewDocIDForProtocols(S_ClassDoc, sDepartmentRoot)   
		S_DocIDAdd = S_DocID
	End If
	If sDepartmentRoot <> SIT_STS or (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_RASP_DOCS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1) Then
		'реализация полужесткого маршрута
		S_ListToReconcile = Replace(S_ListToReconcile, Trim(SIT_AdditionalAgrees), "")
		'Для коммерческих предложений и заявок на закупку и оплату не нужна приписка про доп. согласующих  
		If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_COM_OFFERS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) <> 1 Then
			If S_ListToReconcile <> "" Then
				S_ListToReconcile = " " + SIT_AdditionalAgrees + DeleteConstPrefixFromList(S_ListToReconcile)
			End If
		End If
	End If

	'vnik_protocols
	If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_PROTOCOLS)) > 0 Then
		If sDepartmentRoot = SIT_SITRONICS Then
			S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
			S_ListToView = S_ListToView + Replace(Replace(Replace(S_ListToReconcile,"Обязательные согласующие: ",""), "Дополнительные согласующие: ", ""), "##;", "")
		ElseIf sDepartmentRoot = SIT_STS Then
		'Else 'Другие БН
			'S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile+VbCrLf+SIT_VicePresidentOfInitiator
		End If
	End If
	'vnik_protocols
	'vnik_protocolsCPC
	If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_PROTOCOLS_CPC)) > 0 Then
		If sDepartmentRoot = SIT_SITRONICS Then
			S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
			'S_ListToView = S_ListToView + Replace(Replace(Replace(S_ListToReconcile,"Обязательные согласующие: ",""), "Дополнительные согласующие: ", ""), "##;", "")
		ElseIf sDepartmentRoot = SIT_STS Then
		'Else 'Другие БН
			'S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile+VbCrLf+SIT_VicePresidentOfInitiator
		End If
	End If
	'vnik_protocolsCPC
	'Для всех категорий кроме заявок подставляем пользователей вместо ролей
	If (InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1) and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) <> 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_DOGOVORI_NEW)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIS)) <> 0) Then
		'Получаем список ролей
		If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) = 1 Then 'Для договоров получаем список ролей на всех языках (могут быть роли на разных языках)
			sRoleList = GetFullRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"), "", "", "")
		Else
			sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))
		End If
		'Подставляем роли в поля
		S_NameAproval = ReplaceRolesInList(S_NameAproval, sRoleList)
		S_NameResponsible = ReplaceRolesInList(S_NameResponsible, sRoleList)
		S_NameControl = ReplaceRolesInList(S_NameControl, sRoleList)
		S_ListToEdit = ReplaceRolesInList(S_ListToEdit, sRoleList)
		S_ListToView = ReplaceRolesInList(S_ListToView, sRoleList)
		S_ListToReconcile = ReplaceRolesInList(S_ListToReconcile, sRoleList)

		'Удаляем повторы в списке согласования
		S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)

		'vnik_error
		'Убираем автора документа из списка согласующих, т.к. после активации он не сможет загрузить новый документ
		'(работает пока только, если пользователь создает документ не под ролью)
		S_ListToReconcile = Replace(S_ListToReconcile,Request("DocAuthor"),"")
		'vnik_error
	End If
'vnik_protocols
	If sDepartmentRoot = SIT_SITRONICS Then
		S_ListToView = DeleteUserDoublesInList(S_ListToView)    
	End If
'vnik_protocols
	'Проверка полей с пользователями
	'Для заявок отключаем проверку утверждающего, т.к. он рассчитывается
	If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 Then
		UserChangeDocCheck = CheckSingleUserField(S_NameAproval)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameAproval")+SIT_ErrorInUserField2)
			Exit Function
		End If
		'Для заявки на закупку отключаем проверку ответственного, т.к. можно указывать подразделение
		'Для заявки на оплату тоже, т.к. автоматом ставится роль из специального справочника
		UserChangeDocCheck = CheckSingleUserField(S_NameResponsible)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameResponsible")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
	UserChangeDocCheck = CheckSingleUserField(S_NameControl)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameControl")+SIT_ErrorInUserField2)
		Exit Function
	End If
	If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) = 0 Then 'В исходящих может быть любой текст
		UserChangeDocCheck = CheckMultiUserField(S_Correspondent)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("Correspondent")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
'Запрос №1 - СИБ - start
	IF InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_SLUZH_ZAPISKA)) = 0) THEN
'Запрос №1 - СИБ - end
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) = 0 Then 'В СЗ на переработки работают правила
			UserChangeDocCheck = CheckMultiUserField(S_ListToReconcile)
			If not UserChangeDocCheck Then
				Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToReconcile")+SIT_ErrorInUserField2)
				Exit Function
			End If
'Запрос №1 - СИБ - start
		End If
	END IF
'Запрос №1 - СИБ - end
	UserChangeDocCheck = CheckMultiUserField(S_ListToEdit)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToEdit")+SIT_ErrorInUserField2)
		Exit Function
	End If
	'UserChangeDocCheck = CheckMultiUserField(S_ListToView)
	'If not UserChangeDocCheck Then
	'	Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToView")+SIT_ErrorInUserField2)
	'	Exit Function
	'End If

'{ph - 20120812
	'S_AdditionalUsers = GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit"))
	S_AdditionalUsers = DeleteUserDoublesInListNew(Request("AdditionalUsers") & VbCrLf & GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit")), True)
'ph - 20120812}
	AddLogD "@@@Search For Leaders - S_AdditionalUsers: "+S_AdditionalUsers
	'vnik_payment_order
	'If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_PAYMENT_ORDER)) > 0 Then
	'	S_AdditionalUsers = S_AdditionalUsers + ReplaceRoleFromDir(SIT_Account_manager,sDepartmentRoot)       
	'End If        
	'vnik_payment_order
	S_AdditionalUsers = DeleteUserDoublesInList(S_AdditionalUsers)

	'vnik
	If S_AdditionalUsers = "" Then
		S_AdditionalUsers = "-"
	End If
	S_AdditionalUsers_Set = S_AdditionalUsers

	S_ListToReconcile_Set = S_ListToReconcile
	S_NameAproval_Set = S_NameAproval
	S_ListToEdit = S_ListToEdit
	S_NameResponsible_Set = S_NameResponsible
	S_NameControl_Set = S_NameControl
	S_ListToView_Set=S_ListToView

	If UCase(Request("UpdateDoc")) = "YES" And (Request("create") = "y"  Or Request("DocIDPrev") <> Request("DocID")) Then
		'out "123"
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		dsTemp.Open "select * from Docs where DocID='" + MakeSQLSafe(Request("DocID")) + "'", Conn 
		If Not dsTemp.EOF Then
			Session("Message") = DOCS_ALREADYEXISTS + ": " + Request("DocID")
			dsTemp.Close
			UserChangeDocCheck=False
			Exit Function
		End If
		dsTemp.Close
	End If

	'ph - 20090808 - start
	If Request("DocIDParent") <> "" Then
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		bMyDocIDParentChanged = False
		If Request("create") <> "y" Then
			'Проверяем изменился ли номер родительского документа, если идет редактирование
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocID"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				bMyDocIDParentChanged = Request("DocIDParent")<>dsTemp("DocIDParent")
			End If
			dsTemp.Close
		End If
		If bMyDocIDParentChanged Then
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocIDParent"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				'sMessage = ""
				'UserChangeDocCheck = oPayDox.IsReadAccessUser(Session("UserID"), sMessage, dsTemp, DOCS_All, DOCS_NoReadAccess, DOCS_NoAccess, USER_Department, VAR_ExtInt, VAR_AdminSecLevel, VAR_StatusActiveUser, NULL)
				UserChangeDocCheck = IsReadAccessRS(dsTemp)' IsReadAccessUser(Session("UserID"), dsTemp, NULL)
				If UserChangeDocCheck Then
					S_DocIDParent = dsTemp("DocID")
				End If
				dsTemp.Close
				If not UserChangeDocCheck Then
					Session("Message") = SIT_NoAccessToParentDoc'sMessage
					Exit Function
				End If
			End If
		End If
	End If
	'ph - 20090808 - end
	
'rti_payment_order
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(RTI_PAYMENT_ORDER)) = 1 Then
	'Создатель документа
	If UCase(Request("create")) = "Y" Then
		sDocCreator = Session("UserID")
	Else
		sDocCreator = GetUserID(Request("DocAuthor"))
	End If

	S_NameAproval=Request("DocNameAproval")
	S_NameResponsible=Request("DocNameResponsible")
	S_NameControl=Request("DocNameControl")
	S_ListToReconcile=Request("DocListToReconcile")
	S_ListToEdit=Request("DocListToEdit")
	S_ListToView=Request("DocListToView")
	S_Correspondent=Request("DocCorrespondent")

	'Определяем бизнес направление
	sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))
	sDepartment = SIT_RTI
	sDepartmentCode = ""


	If Request("create") = "y" Then
        addlogd "eXor777 :" + S_ClassDoc + " " + sDepartmentRoot
		S_DocID = GetNewDocIDForPaymentOrder(S_ClassDoc, sDepartmentRoot, "PRTI", "", "", "", "")
		S_DocIDAdd = S_DocID
			
	End If
	If sDepartmentRoot <> SIT_STS or (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_RASP_DOCS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1) Then
		'реализация полужесткого маршрута
		S_ListToReconcile = Replace(S_ListToReconcile, Trim(SIT_AdditionalAgrees), "")
		'Для коммерческих предложений и заявок на закупку и оплату не нужна приписка про доп. согласующих  
		If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_COM_OFFERS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) <> 1 Then
			If S_ListToReconcile <> "" Then
				S_ListToReconcile = " " + SIT_AdditionalAgrees + DeleteConstPrefixFromList(S_ListToReconcile)
			End If
		End If
	End If
	
 

		'добавление обязательных согласующих
		'S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeRTIRU","Category",Session("CurrentClassDoc"),"List")+SIT_AdditionalAgreesDelimeter+Trim(S_ListToReconcile)
		'AddLogD "vnik468 " + Trim(SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeRTI"+Request("l"),"Category",Session("CurrentClassDoc"),"List")+Trim(S_ListToReconcile) + VbCrLf + RTI_ChiefOfPurchaseDepartment + VbCrLf+RTI_HeadKFIE)
		'S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeRTI"+Request("l"),"Category",Session("CurrentClassDoc"),"List")+Trim(S_ListToReconcile) + VbCrLf + RTI_ChiefOfPurchaseDepartment + VbCrLf+RTI_HeadKFIE
				If Request("create") = "y" Then
            'S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeRTI"+Request("l"),"Category",Session("CurrentClassDoc"),"List")+ SIT_AdditionalAgrees + Trim(S_ListToReconcile) +VbCrLf+ RTI_ChiefOfPurchaseDepartment + VbCrLf + RTI_HeadKFIE
            S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeRTI"+Request("l"),"Category",Session("CurrentClassDoc"),"List")+ SIT_AdditionalAgrees + Trim(S_ListToReconcile) + VbCrLf + RTI_HeadKFIE
         Else
            AddLogD "vnik01234 " + Trim(S_ListToReconcile)
            sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))
            'S_ListToReconcile = replace(S_ListToReconcile, ReplaceRolesInList(RTI_ChiefOfPurchaseDepartment, sRoleList), "")
            S_ListToReconcile = replace(S_ListToReconcile, ReplaceRolesInList(RTI_HeadKFIE, sRoleList), "")
            S_ListToReconcile = replace(S_ListToReconcile, ReplaceRolesInList(oPayDox.GetExtTableValue("AgreeRTI"+Request("l"),"Category",Session("CurrentClassDoc"),"List"), sRoleList), "")
            S_ListToReconcile = replace(S_ListToReconcile, SIT_AdditionalAgrees, "")
            S_ListToReconcile = replace(S_ListToReconcile, SIT_RequiredAgrees, "")                      
            S_ListToReconcile = replace(S_ListToReconcile, VbCrLf, "")         
		    S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)
            'S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeRTI"+Request("l"),"Category",Session("CurrentClassDoc"),"List")+SIT_AdditionalAgrees + Trim(S_ListToReconcile) +VbCrLf+ RTI_ChiefOfPurchaseDepartment + VbCrLf + RTI_HeadKFIE
            S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeRTI"+Request("l"),"Category",Session("CurrentClassDoc"),"List")+SIT_AdditionalAgrees + Trim(S_ListToReconcile) + VbCrLf + RTI_HeadKFIE
         End If
        'Удаляем повторы в списке согласования и в получателях
		S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)
		S_ListToView = S_ListToView + RTI_PaymentChief
		S_ListToView = DeleteUserDoublesInList(S_ListToView)
		'S_ListToView = S_ListToView + Replace(Replace(Replace(S_ListToReconcile,"Обязательные согласующие: ",""), "Дополнительные согласующие: ", ""), "##;", "")

	'Для всех категорий кроме заявок подставляем пользователей вместо ролей
	If (InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1) and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) <> 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_DOGOVORI_NEW)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIS)) <> 0) Then
		'Получаем список ролей
		If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) = 1 Then 'Для договоров получаем список ролей на всех языках (могут быть роли на разных языках)
			sRoleList = GetFullRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"), "", "", "")
		Else
			sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))
		End If
		'Подставляем роли в поля
		S_NameAproval = ReplaceRolesInList(S_NameAproval, sRoleList)
		S_NameResponsible = ReplaceRolesInList(S_NameResponsible, sRoleList)
		S_NameControl = ReplaceRolesInList(S_NameControl, sRoleList)
		S_ListToEdit = ReplaceRolesInList(S_ListToEdit, sRoleList)
		S_ListToView = ReplaceRolesInList(S_ListToView, sRoleList)
		S_ListToReconcile = ReplaceRolesInList(S_ListToReconcile, sRoleList)
         AddLogD "eXor235 " + S_ListToReconcile
		'Удаляем повторы в списке согласования
		S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)
        AddLogD "eXor236 " + S_ListToReconcile
		'vnik_error
		'Убираем автора документа из списка согласующих, т.к. после активации он не сможет загрузить новый документ
		'(работает пока только, если пользователь создает документ не под ролью)
		S_ListToReconcile = Replace(S_ListToReconcile,Request("DocAuthor"),"")
		AddLogD "eXor237 " + S_ListToReconcile
		'vnik_error
	End If
	'Проверка полей с пользователями
	'Для заявок отключаем проверку утверждающего, т.к. он рассчитывается
	If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 Then
		UserChangeDocCheck = CheckSingleUserField(S_NameAproval)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameAproval")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
	UserChangeDocCheck = CheckMultiUserField(S_ListToEdit)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToEdit")+SIT_ErrorInUserField2)
		Exit Function
	End If
	
	'29-04-2013 проверка поля центр затрат и статья расходов

	sSQL = "select * from RTI_CostCenter where FullName = "+sUnicodeSymbol+"'"+ Request("UserFieldText2") + "'"
    Set dsTemp = Server.CreateObject("ADODB.Recordset")
    dsTemp.Open sSQL, Conn, 3, 1, &H1
    UserChangeDocCheck = not dsTemp.EOF
    dsTemp.Close
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+" Центр затрат " + SIT_ErrorInUserField2)
		Exit Function
	End If
	

    'проверка статьи расходов
      sSQL = "select * from RTI_CostItem2 where Name = "+sUnicodeSymbol+"'"+ Request("UserFieldText4") + "'"
      Set dsTemp = Server.CreateObject("ADODB.Recordset")
      dsTemp.Open sSQL, Conn, 3, 1, &H1
      if not dsTemp.EOF Then
       isValid = dsTemp("isValid")
      End If       
      UserChangeDocCheck = not dsTemp.EOF    
      if IsValid = "1" Then
        UserChangeDocCheck = false
      End If
      dsTemp.Close 
    End If
    
	If not UserChangeDocCheck  Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+" Статья расходов " + SIT_ErrorInUserField2 + ". Либо вместо статьи выбрана категория статей (в этом случае выбранная строка не содержит номера статьи расходов)")
		Exit Function
		
	'проверка статьи расходов

'{ph - 20120812
	'S_AdditionalUsers = GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit"))
	'S_AdditionalUsers = DeleteUserDoublesInListNew(Request("AdditionalUsers") & VbCrLf & GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit")), True)
'ph - 20120812}
	AddLogD "@@@Search For Leaders - S_AdditionalUsers: "+S_AdditionalUsers

	S_AdditionalUsers = DeleteUserDoublesInList(S_AdditionalUsers)

	'vnik
	If S_AdditionalUsers = "" Then
		S_AdditionalUsers = "-"
	End If
	S_AdditionalUsers_Set = S_AdditionalUsers

	S_ListToReconcile_Set = S_ListToReconcile
	S_NameAproval_Set = S_NameAproval
	S_ListToEdit = S_ListToEdit
	S_NameResponsible_Set = S_NameResponsible
	S_NameControl_Set = S_NameControl
	S_ListToView_Set=S_ListToView
	

	If UCase(Request("UpdateDoc")) = "YES" And (Request("create") = "y"  Or Request("DocIDPrev") <> Request("DocID")) Then
		'out "123"
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		dsTemp.Open "select * from Docs where DocID='" + MakeSQLSafe(Request("DocID")) + "'", Conn 
		If Not dsTemp.EOF Then
			Session("Message") = DOCS_ALREADYEXISTS + ": " + Request("DocID")
			dsTemp.Close
			UserChangeDocCheck=False
			Exit Function
		End If
		dsTemp.Close
	End If

	'ph - 20090808 - start
	If Request("DocIDParent") <> "" Then
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		bMyDocIDParentChanged = False
		If Request("create") <> "y" Then
			'Проверяем изменился ли номер родительского документа, если идет редактирование
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocID"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				bMyDocIDParentChanged = Request("DocIDParent")<>dsTemp("DocIDParent")
			End If
			dsTemp.Close
		End If
		If bMyDocIDParentChanged Then
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocIDParent"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				'sMessage = ""
				'UserChangeDocCheck = oPayDox.IsReadAccessUser(Session("UserID"), sMessage, dsTemp, DOCS_All, DOCS_NoReadAccess, DOCS_NoAccess, USER_Department, VAR_ExtInt, VAR_AdminSecLevel, VAR_StatusActiveUser, NULL)
				UserChangeDocCheck = IsReadAccessRS(dsTemp)' IsReadAccessUser(Session("UserID"), dsTemp, NULL)
				If UserChangeDocCheck Then
					S_DocIDParent = dsTemp("DocID")
				End If
				dsTemp.Close
				If not UserChangeDocCheck Then
					Session("Message") = SIT_NoAccessToParentDoc'sMessage
					Exit Function
				End If
			End If
		End If
	End If
	'ph - 20090808 - end	
'rti_payment_order

' ***ЗАЯВКА НА ОПЛАТУ УК
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_PAYMENT_ORDER)) = 1 Then
	'Создатель документа
	If UCase(Request("create")) = "Y" Then
		sDocCreator = Session("UserID")
	Else
		sDocCreator = GetUserID(Request("DocAuthor"))
	End If

	S_NameAproval=Request("DocNameAproval")
	S_NameResponsible=Request("DocNameResponsible")
	S_NameControl=Request("DocNameControl")
	S_ListToReconcile=Request("DocListToReconcile")
	S_ListToEdit=Request("DocListToEdit")
	S_ListToView=Request("DocListToView")
	S_Correspondent=Request("DocCorrespondent")

	'Определяем бизнес направление
	sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))

	' SAY 2008-07-21
	Set dsTempPR = Server.CreateObject("ADODB.Recordset")

' DmGorsky_5: меняем местами проверку: вначале ищем SIT_SITRU ("СИТРОНИКС ИТ"), потом, если не найдено, - SIT_SITRONICS ("СИТРОНИКС")
	If InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRU))=1 then ' DmGorsky
		sDepartment = SIT_SITRU ' DmGorsky
		sDepartmentCode = "" ' DmGorsky
'Запрос №1 - СИБ - start
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SIB_ROOT_DEPARTMENT))=1 then
		sDepartment = SIT_SIB
		sDepartmentCode = ""
'Запрос №1 - СИБ - end
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRONICS))=1 then
		sDepartment = SIT_SITRONICS
		sDepartmentCode = ""
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_RTI))=1 then
		sDepartment = SIT_RTI
		sDepartmentCode = ""
	Else
		sDepartment = SIT_STS

		sSQL = "select code from Departments where Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"' or Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"/'"
		AddLogD "WAY sSQL="+sSQL
		AddLogD "WAY Conn="+Conn
		dsTempPR.Open sSQL, Conn, 3, 1, &H1

		if not dsTempPR.EOF Then
			sDepartmentCode = dsTempPR("code")
			if InStrRev(sDepartmentCode, "/") > 0 then
				sDepartmentCode = Right(sDepartmentCode, Len(sDepartmentCode)-InStrRev(sDepartmentCode, "/"))
			End If
		Else
			sDepartmentCode="000000"
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInDepartmentCode1 + DelOtherLangFromFolder(Request("DocDepartment")) + SIT_ErrorInDepartmentCode2)
			UserChangeDocCheck=False
			Exit Function
		End If

		dsTempPR.Close
	End If

	If Request("create") = "y" Then
		sPostfix = ""
		sSearchCol = "DocID"
		sPrePrefix = ""
		'SAY 2009-02-20
		sSufix = ""

		'SAY 2009-03-12
		S_DocID = ""

		Call GetNewDocID_test(S_ClassDoc, sDepartmentRoot, Right(CStr(Year(Date)),2) + "/", "PHQ-", "", "")
		S_DocIDAdd = S_DocID
	End If
	If sDepartmentRoot <> SIT_STS or (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_RASP_DOCS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1) Then
		'реализация полужесткого маршрута
		S_ListToReconcile = Replace(S_ListToReconcile, Trim(SIT_AdditionalAgrees), "")
		'Для коммерческих предложений и заявок на закупку и оплату не нужна приписка про доп. согласующих  
		If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_COM_OFFERS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) <> 1 Then
			If S_ListToReconcile <> "" Then
				S_ListToReconcile = " " + SIT_AdditionalAgrees + DeleteConstPrefixFromList(S_ListToReconcile)
			End If
		End If
	End If
'vnik_payment_order
	'Проверка коррректности ввода Заявки на оплату УК
	AddLogD "" + Trim(MyDate(Date))
	AddLogD "" + Trim(MyDate("31.07.2011"))
	If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_PAYMENT_ORDER)) = 1 and (MyDate(Date) >= MyDate("31.07.2011")) Then
		AddLogD "" + Trim(MyDate("01.08.2011"))
		If UCase(Request("UpdateDoc")) = "YES" Then
			If Trim(Request("DocIDParent")) = "" Then  
				vnik_SimplePurchase = ""
				If Request("DocCurrency") = "RUR" Then
					If Request("UserFieldMoney1") > 30000 Then
						vnik_SimplePurchase = "0"
					Else
						vnik_SimplePurchase = "1"
					End If
				ElseIf Request("DocCurrency") = "USD" Then
					If Request("UserFieldMoney1") > 1000 Then
						vnik_SimplePurchase = "0"
					Else
						vnik_SimplePurchase = "1"
					End If
				ElseIf Request("DocCurrency") = "EUR" Then
					If Request("UserFieldMoney1") > 750 Then
						vnik_SimplePurchase = "0"
					Else
						vnik_SimplePurchase = "1"
					End If
				End If      

				If vnik_SimplePurchase = "" Then
					Session("Message") = AddNewLineToMessage(Session("Message"), SIT_CONTRACTS_MC_WARNING4)
					UserChangeDocCheck = False
					Exit Function
				ElseIf vnik_SimplePurchase = "0" Then
					Session("Message") = AddNewLineToMessage(Session("Message"), SIT_PAYMENT_ORDER_WARNING1)
					UserChangeDocCheck = False
					Exit Function    
				End If
			Else
				'Проверяем документ основание
				Set vnikdsTemp1 = Server.CreateObject("ADODB.Recordset")
				sSQL = "Select * from Docs where DocID = "+sUnicodeSymbol+"'"+Request("DocIDParent")+"'" + " and StatusDevelopment = 4"
				vnikdsTemp1.Open sSQL, Conn, 3, 1, &H1 
				If vnikdsTemp1.EOF Then
					Session("Message") = AddNewLineToMessage(Session("Message"), SIT_PAYMENT_ORDER_WARNING2)
					UserChangeDocCheck = False
					Exit Function
				Else
					If vnikdsTemp1("UserFieldText8") = "Рамочный" Then
						Session("Message") = AddNewLineToMessage(Session("Message"), SIT_PAYMENT_ORDER_WARNING3)
						UserChangeDocCheck = False
						Exit Function    
					End If
				End If   
				vnikdsTemp1.Close           
			End If
		End If
	End If
'vnik_payment_order
'vnik_payment_order
	If sDepartmentRoot = SIT_SITRONICS or sDepartmentRoot = SIT_RTI Then
		AddLogD "vnik123" + Trim(S_ListToReconcile) 
		'добавление обязательных согласующих
		S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+Trim(S_ListToReconcile)

		AddLogD "vnik123" + Trim(S_ListToReconcile)

		VNIK_TempValueAgree = ReplaceRoleFromDir(SIT_Account_manager,sDepartmentRoot)
		If InStr(UCase(S_ListToReconcile),UCase(VNIK_TempValueAgree)) < 1 Then
			S_ListToReconcile = S_ListToReconcile+VbCrLf+ VNIK_TempValueAgree 'Бухгалтер по заявкам на оплату
		Else
			If VNIK_TempValueAgree <> "" Then
				S_ListToReconcile = Replace(S_ListToReconcile, VNIK_TempValueAgree, "")
				S_ListToReconcile = S_ListToReconcile+VbCrLf+ VNIK_TempValueAgree
			End If
		End If

		VNIK_TempValueAgree = ReplaceRoleFromDir(SIT_BudgetController,sDepartmentRoot)
		If InStr(UCase(S_ListToReconcile),UCase(VNIK_TempValueAgree)) < 1 Then
			S_ListToReconcile = S_ListToReconcile+VbCrLf+ VNIK_TempValueAgree' + ";""Мякотникова Е. А."" <myakotnikova_oaorti>;" 'Бюджетный контролер
		Else
			If VNIK_TempValueAgree <> "" Then
				S_ListToReconcile = Replace(S_ListToReconcile, VbCrLf+VNIK_TempValueAgree, "")
				S_ListToReconcile = S_ListToReconcile+VbCrLf+ VNIK_TempValueAgree' + ";""Мякотникова Е. А."" <myakotnikova_oaorti>;"
			End If
		End If

		VNIK_TempValueAgree = GetUserDirValuesVNIK("{E2D63CEF-6DC4-44F3-B4BF-BD4B0D9438D0}","Field3","Field1",Request("UserFieldText4"))
		If InStr(UCase(S_ListToReconcile),UCase(VNIK_TempValueAgree)) < 1 Then
			S_ListToReconcile = S_ListToReconcile+VbCrLf+ VNIK_TempValueAgree 'РуководительЦФЗ
		Else
			If VNIK_TempValueAgree <> "" Then
				S_ListToReconcile = Replace(S_ListToReconcile, VbCrLf+VNIK_TempValueAgree, "")        
				S_ListToReconcile = S_ListToReconcile+VbCrLf+ VNIK_TempValueAgree    
			End If
		End If

		VNIK_TempValueAgree = GetUserDirValuesVNIK("{15EB5243-22D8-425D-B31A-9CBA4396FCFC}","Field3","Field1",Request("UserFieldText3"))
		If InStr(UCase(S_ListToReconcile),UCase(VNIK_TempValueAgree)) < 1 Then
			S_ListToReconcile = S_ListToReconcile+VbCrLf+ VNIK_TempValueAgree 'РуководительЦФО
		Else    'это чтобы порядок согласующих не менялся
			If VNIK_TempValueAgree <> "" Then
				S_ListToReconcile = Replace(S_ListToReconcile, VbCrLf+VNIK_TempValueAgree, "")
				S_ListToReconcile = S_ListToReconcile+VbCrLf+ VNIK_TempValueAgree
			End If
		End If

		VNIK_TempValueAgree = GetUserDirValuesVNIK("{E2D63CEF-6DC4-44F3-B4BF-BD4B0D9438D0}","Field4","Field1",Request("UserFieldText4"))
		If InStr(UCase(S_ListToReconcile),UCase(VNIK_TempValueAgree)) < 1 Then
			S_ListToReconcile = S_ListToReconcile+VbCrLf+ VNIK_TempValueAgree 'ВПкураторЦФЗ
		Else
			If VNIK_TempValueAgree <> "" Then
				S_ListToReconcile = Replace(S_ListToReconcile, VbCrLf+VNIK_TempValueAgree, "")      
				S_ListToReconcile = S_ListToReconcile+VbCrLf+ VNIK_TempValueAgree
			End If    
		End If    

		'Удаляем повторы в списке согласования
		AddLogD "vnik123" + Trim(S_ListToReconcile)

		'Заполняем список получателей
		'S_ListToView = S_ListToView + ReplaceRoleFromDir(SIT_Account_manager,sDepartmentRoot)
		AddLogD "vnik987" + Trim(Request("DocDepartment"))
		VNIK_ChiefDocDepartment = GetChiefOfDepartment(Request("DocDepartment"),"")
		AddLogD "vnik987" + Trim(VNIK_ChiefDocDepartment)
		If VNIK_ChiefDocDepartment <> "#EMPTY" and VNIK_ChiefDocDepartment <> "#MANY" and VNIK_ChiefDocDepartment <> "" Then
			AddLogD "vnik987" + Trim(S_ListToView)
			S_ListToView = S_ListToView + VNIK_ChiefDocDepartment
		End If
		'Удаляем повторы в списке согласования и в получателях
		S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)
		S_ListToView = DeleteUserDoublesInList(S_ListToView)
		'S_ListToView = S_ListToView + Replace(Replace(Replace(S_ListToReconcile,"Обязательные согласующие: ",""), "Дополнительные согласующие: ", ""), "##;", "")
	ElseIf sDepartmentRoot = SIT_STS Then
	'Else 'Другие БН
		'S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile+VbCrLf+SIT_VicePresidentOfInitiator
	End If
'vnik_payment_order
	'Для всех категорий кроме заявок подставляем пользователей вместо ролей
	If (InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1) and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) <> 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_DOGOVORI_NEW)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIS)) <> 0) Then
		'Получаем список ролей
		If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) = 1 Then 'Для договоров получаем список ролей на всех языках (могут быть роли на разных языках)
			sRoleList = GetFullRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"), "", "", "")
		Else
			sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))
		End If
		'Подставляем роли в поля
		S_NameAproval = ReplaceRolesInList(S_NameAproval, sRoleList)
		S_NameResponsible = ReplaceRolesInList(S_NameResponsible, sRoleList)
		S_NameControl = ReplaceRolesInList(S_NameControl, sRoleList)
		S_ListToEdit = ReplaceRolesInList(S_ListToEdit, sRoleList)
		S_ListToView = ReplaceRolesInList(S_ListToView, sRoleList)
		S_ListToReconcile = ReplaceRolesInList(S_ListToReconcile, sRoleList)

		'Удаляем повторы в списке согласования
		S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)

		'vnik_error
		'Убираем автора документа из списка согласующих, т.к. после активации он не сможет загрузить новый документ
		'(работает пока только, если пользователь создает документ не под ролью)
		S_ListToReconcile = Replace(S_ListToReconcile,Request("DocAuthor"),"")
		'vnik_error
	End If
	'Проверка полей с пользователями
	'Для заявок отключаем проверку утверждающего, т.к. он рассчитывается
	If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 Then
		UserChangeDocCheck = CheckSingleUserField(S_NameAproval)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameAproval")+SIT_ErrorInUserField2)
			Exit Function
		End If
		'Для заявки на закупку отключаем проверку ответственного, т.к. можно указывать подразделение
		'Для заявки на оплату тоже, т.к. автоматом ставится роль из специального справочника
		UserChangeDocCheck = CheckSingleUserField(S_NameResponsible)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameResponsible")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
	UserChangeDocCheck = CheckSingleUserField(S_NameControl)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameControl")+SIT_ErrorInUserField2)
		Exit Function
	End If
	If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) = 0 Then 'В исходящих может быть любой текст
		UserChangeDocCheck = CheckMultiUserField(S_Correspondent)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("Correspondent")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
'Запрос №1 - СИБ - start
	IF InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_SLUZH_ZAPISKA)) = 0) THEN
'Запрос №1 - СИБ - end
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) = 0 Then 'В СЗ на переработки работают правила
			UserChangeDocCheck = CheckMultiUserField(S_ListToReconcile)
			If not UserChangeDocCheck Then
				Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToReconcile")+SIT_ErrorInUserField2)
				Exit Function
			End If
'Запрос №1 - СИБ - start
		End If
	END IF
'Запрос №1 - СИБ - end
	UserChangeDocCheck = CheckMultiUserField(S_ListToEdit)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToEdit")+SIT_ErrorInUserField2)
		Exit Function
	End If
	'UserChangeDocCheck = CheckMultiUserField(S_ListToView)
	'If not UserChangeDocCheck Then
	'	Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToView")+SIT_ErrorInUserField2)
	'	Exit Function
	'End If

'{ph - 20120812
	'S_AdditionalUsers = GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit"))
	S_AdditionalUsers = DeleteUserDoublesInListNew(Request("AdditionalUsers") & VbCrLf & GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit")), True)
'ph - 20120812}
	AddLogD "@@@Search For Leaders - S_AdditionalUsers: "+S_AdditionalUsers
	'vnik_payment_order
	'If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_PAYMENT_ORDER)) > 0 Then
	'	S_AdditionalUsers = S_AdditionalUsers + ReplaceRoleFromDir(SIT_Account_manager,sDepartmentRoot)       
	'End If        
	'vnik_payment_order
	S_AdditionalUsers = DeleteUserDoublesInList(S_AdditionalUsers)

	'vnik
	If S_AdditionalUsers = "" Then
		S_AdditionalUsers = "-"
	End If
	S_AdditionalUsers_Set = S_AdditionalUsers

	S_ListToReconcile_Set = S_ListToReconcile
	S_NameAproval_Set = S_NameAproval
	S_ListToEdit = S_ListToEdit
	S_NameResponsible_Set = S_NameResponsible
	S_NameControl_Set = S_NameControl
	S_ListToView_Set=S_ListToView

	If UCase(Request("UpdateDoc")) = "YES" And (Request("create") = "y"  Or Request("DocIDPrev") <> Request("DocID")) Then
		'out "123"
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		dsTemp.Open "select * from Docs where DocID='" + MakeSQLSafe(Request("DocID")) + "'", Conn 
		If Not dsTemp.EOF Then
			Session("Message") = DOCS_ALREADYEXISTS + ": " + Request("DocID")
			dsTemp.Close
			UserChangeDocCheck=False
			Exit Function
		End If
		dsTemp.Close
	End If

	'ph - 20090808 - start
	If Request("DocIDParent") <> "" Then
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		bMyDocIDParentChanged = False
		If Request("create") <> "y" Then
			'Проверяем изменился ли номер родительского документа, если идет редактирование
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocID"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				bMyDocIDParentChanged = Request("DocIDParent")<>dsTemp("DocIDParent")
			End If
			dsTemp.Close
		End If
		If bMyDocIDParentChanged Then
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocIDParent"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				'sMessage = ""
				'UserChangeDocCheck = oPayDox.IsReadAccessUser(Session("UserID"), sMessage, dsTemp, DOCS_All, DOCS_NoReadAccess, DOCS_NoAccess, USER_Department, VAR_ExtInt, VAR_AdminSecLevel, VAR_StatusActiveUser, NULL)
				UserChangeDocCheck = IsReadAccessRS(dsTemp)' IsReadAccessUser(Session("UserID"), dsTemp, NULL)
				If UserChangeDocCheck Then
					S_DocIDParent = dsTemp("DocID")
				End If
				dsTemp.Close
				If not UserChangeDocCheck Then
					Session("Message") = SIT_NoAccessToParentDoc'sMessage
					Exit Function
				End If
			End If
		End If
	End If
	'ph - 20090808 - end

'rti_bsap
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(RTI_BSAP)) > 0 Then
	'Создатель документа
	If UCase(Request("create")) = "Y" Then
		sDocCreator = Session("UserID")
	Else
		sDocCreator = GetUserID(Request("DocAuthor"))
	End If

	S_NameAproval=Request("DocNameAproval")
	S_NameResponsible=Request("DocNameResponsible")
	S_NameControl=Request("DocNameControl")
	S_ListToReconcile=Request("DocListToReconcile")
	S_ListToEdit=Request("DocListToEdit")
	S_ListToView=Request("DocListToView")
	S_Correspondent=Request("DocCorrespondent")

	'Определяем бизнес направление
	sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))
    sDepartment = SIT_RTI
    sDepartmentCode = ""


	If Request("create") = "y" Then
		sPostfix = ""
		sSearchCol = "DocID"
		sPrePrefix = ""
		'SAY 2009-02-20
		sSufix = ""

		'SAY 2009-03-12
		S_DocID = ""

		S_DocID = GetNewDocIDForPurchaseOrder(S_ClassDoc, sDepartmentRoot, "FCAS", "", "", "", "")
		S_DocIDAdd = S_DocID
	End If
	If sDepartmentRoot <> SIT_STS or (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_RASP_DOCS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1) Then
		'реализация полужесткого маршрута
		S_ListToReconcile = Replace(S_ListToReconcile, Trim(SIT_AdditionalAgrees), "")
		'Для коммерческих предложений и заявок на закупку и оплату не нужна приписка про доп. согласующих  
		If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_COM_OFFERS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) <> 1 Then
			If S_ListToReconcile <> "" Then
				S_ListToReconcile = " " + SIT_AdditionalAgrees + DeleteConstPrefixFromList(S_ListToReconcile)
			End If
		End If
	End If

		'добавление обязательных согласующих
		'S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeRTI"+Request("l"),"Category",Session("CurrentClassDoc"),"List")+SIT_AdditionalAgreesDelimeter+Trim(S_ListToReconcile)
        S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeRTI"+Request("l"),"Category",Session("CurrentClassDoc"),"List")+Trim(S_ListToReconcile) +VbCrLf+RTI_HeadOfPurchaseCenter

		'Удаляем повторы в списке согласования
		S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)
		S_ListToView = DeleteUserDoublesInList(S_ListToView)


	'Для всех категорий кроме заявок подставляем пользователей вместо ролей
		'Получаем список ролей
		If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) = 1 Then 'Для договоров получаем список ролей на всех языках (могут быть роли на разных языках)
			sRoleList = GetFullRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"), "", "", "")
		Else
			sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))
		End If
		'Подставляем роли в поля
		S_NameAproval = ReplaceRolesInList(S_NameAproval, sRoleList)
		S_NameResponsible = ReplaceRolesInList(S_NameResponsible, sRoleList)
		S_NameControl = ReplaceRolesInList(S_NameControl, sRoleList)
		S_ListToEdit = ReplaceRolesInList(S_ListToEdit, sRoleList)
		S_ListToView = ReplaceRolesInList(S_ListToView, sRoleList)
		S_ListToReconcile = ReplaceRolesInList(S_ListToReconcile, sRoleList)
		S_ListToReconcile = Replace(S_ListToReconcile,"""Боев С. Ф."" <boev_oaorti>;","")

		'Удаляем повторы в списке согласования
		S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)

		'vnik_error
		'Убираем автора документа из списка согласующих, т.к. после активации он не сможет загрузить новый документ
		'(работает пока только, если пользователь создает документ не под ролью)
		S_ListToReconcile = Replace(S_ListToReconcile,Request("DocAuthor"),"")
		'vnik_error
	'Проверка полей с пользователями
	'Для заявок отключаем проверку утверждающего, т.к. он рассчитывается
	If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 Then
		UserChangeDocCheck = CheckSingleUserField(S_NameAproval)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameAproval")+SIT_ErrorInUserField2)
			Exit Function
		End If
		'Для заявки на закупку отключаем проверку ответственного, т.к. можно указывать подразделение
		'Для заявки на оплату тоже, т.к. автоматом ставится роль из специального справочника
		UserChangeDocCheck = CheckSingleUserField(S_NameResponsible)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameResponsible")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
	UserChangeDocCheck = CheckSingleUserField(S_NameControl)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameControl")+SIT_ErrorInUserField2)
		Exit Function
	End If
	UserChangeDocCheck = CheckMultiUserField(S_ListToEdit)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToEdit")+SIT_ErrorInUserField2)
		Exit Function
	End If

	S_AdditionalUsers = DeleteUserDoublesInList(S_AdditionalUsers)

	'vnik
	If S_AdditionalUsers = "" Then
		S_AdditionalUsers = "-"
	End If
	S_AdditionalUsers_Set = S_AdditionalUsers

	S_ListToReconcile_Set = S_ListToReconcile
	S_NameAproval_Set = S_NameAproval
	S_ListToEdit = S_ListToEdit
	S_NameResponsible_Set = S_NameResponsible
	S_NameControl_Set = S_NameControl
	S_ListToView_Set=S_ListToView

	If UCase(Request("UpdateDoc")) = "YES" And (Request("create") = "y"  Or Request("DocIDPrev") <> Request("DocID")) Then
		'out "123"
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		dsTemp.Open "select * from Docs where DocID='" + MakeSQLSafe(Request("DocID")) + "'", Conn 
		If Not dsTemp.EOF Then
			Session("Message") = DOCS_ALREADYEXISTS + ": " + Request("DocID")
			dsTemp.Close
			UserChangeDocCheck=False
			Exit Function
		End If
		dsTemp.Close
	End If

	'ph - 20090808 - start
	If Request("DocIDParent") <> "" Then
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		bMyDocIDParentChanged = False
		If Request("create") <> "y" Then
			'Проверяем изменился ли номер родительского документа, если идет редактирование
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocID"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				bMyDocIDParentChanged = Request("DocIDParent")<>dsTemp("DocIDParent")
			End If
			dsTemp.Close
		End If
		If bMyDocIDParentChanged Then
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocIDParent"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				'sMessage = ""
				'UserChangeDocCheck = oPayDox.IsReadAccessUser(Session("UserID"), sMessage, dsTemp, DOCS_All, DOCS_NoReadAccess, DOCS_NoAccess, USER_Department, VAR_ExtInt, VAR_AdminSecLevel, VAR_StatusActiveUser, NULL)
				UserChangeDocCheck = IsReadAccessRS(dsTemp)' IsReadAccessUser(Session("UserID"), dsTemp, NULL)
				If UserChangeDocCheck Then
					S_DocIDParent = dsTemp("DocID")
				End If
				dsTemp.Close
				If not UserChangeDocCheck Then
					Session("Message") = SIT_NoAccessToParentDoc'sMessage
					Exit Function
				End If
			End If
		End If
	End If
	'ph - 20090808 - end

'rti_bsap


'rti_purchase_order
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(RTI_PURCHASE_ORDER)) > 0 Then
	'Создатель документа


	If UCase(Request("create")) = "Y" Then
		sDocCreator = Session("UserID")
	Else
		sDocCreator = GetUserID(Request("DocAuthor"))
	End If

	S_NameAproval=Request("DocNameAproval")
	S_NameResponsible=Request("DocNameResponsible")
	S_NameControl=Request("DocNameControl")
	S_ListToReconcile=Request("DocListToReconcile")
	S_ListToEdit=Request("DocListToEdit")
	S_ListToView=Request("DocListToView")
	S_Correspondent=Request("DocCorrespondent")

	'Определяем бизнес направление
	sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))
    sDepartment = SIT_RTI
    sDepartmentCode = ""


	If Request("create") = "y" Then
		sPostfix = ""
		sSearchCol = "DocID"
		sPrePrefix = ""
		'SAY 2009-02-20
		sSufix = ""

		'SAY 2009-03-12
		S_DocID = ""

		S_DocID = GetNewDocIDForPurchaseOrder(S_ClassDoc, sDepartmentRoot, "PORTI", "", "", "", "")
		S_DocIDAdd = S_DocID
	End If
	If sDepartmentRoot <> SIT_STS or (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_RASP_DOCS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1) Then
		'реализация полужесткого маршрута
		S_ListToReconcile = Replace(S_ListToReconcile, Trim(SIT_AdditionalAgrees), "")
		'Для коммерческих предложений и заявок на закупку и оплату не нужна приписка про доп. согласующих  
		If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_COM_OFFERS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) <> 1 Then
			If S_ListToReconcile <> "" Then
				S_ListToReconcile = " " + SIT_AdditionalAgrees + DeleteConstPrefixFromList(S_ListToReconcile)
			End If
		End If
	End If

		'добавление обязательных согласующих
		'S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeRTI"+Request("l"),"Category",Session("CurrentClassDoc"),"List")+SIT_AdditionalAgreesDelimeter+Trim(S_ListToReconcile)
		'S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeRTI"+Request("l"),"Category",Session("CurrentClassDoc"),"List")+Trim(S_ListToReconcile) +VbCrLf+ RTI_BudgetController + VbCrLf + RTI_ChiefOfPurchaseDepartment
		
		If Request("create") = "y" Then
            S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeRTI"+Request("l"),"Category",Session("CurrentClassDoc"),"List")+ SIT_AdditionalAgrees + Trim(S_ListToReconcile) +VbCrLf+ RTI_BudgetController' + VbCrLf + RTI_ChiefOfPurchaseDepartment
         Else
            AddLogD "vnik01234 " + Trim(S_ListToReconcile)
            sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))
            S_ListToReconcile = replace(S_ListToReconcile, ReplaceRolesInList(RTI_BudgetController, sRoleList), "")
            'S_ListToReconcile = replace(S_ListToReconcile, ReplaceRolesInList(RTI_ChiefOfPurchaseDepartment, sRoleList), "")
            S_ListToReconcile = replace(S_ListToReconcile, ReplaceRolesInList(oPayDox.GetExtTableValue("AgreeRTI"+Request("l"),"Category",Session("CurrentClassDoc"),"List"), sRoleList), "")
            S_ListToReconcile = replace(S_ListToReconcile, SIT_AdditionalAgrees, "")
            S_ListToReconcile = replace(S_ListToReconcile, SIT_RequiredAgrees, "")                      
            S_ListToReconcile = replace(S_ListToReconcile, VbCrLf, "")         
		    S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)
            S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeRTI"+Request("l"),"Category",Session("CurrentClassDoc"),"List")+SIT_AdditionalAgrees + Trim(S_ListToReconcile) +VbCrLf+ RTI_BudgetController' + VbCrLf + RTI_ChiefOfPurchaseDepartment
         End If
		
		'Удаляем повторы в списке согласования
		S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)
		S_ListToView = DeleteUserDoublesInList(S_ListToView)
		'S_NameControl = SIT_SecretaryCPC


	'Для всех категорий кроме заявок подставляем пользователей вместо ролей

     	sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))

		'Подставляем роли в поля
		S_NameAproval = ReplaceRolesInList(S_NameAproval, sRoleList)
		S_NameResponsible = ReplaceRolesInList(S_NameResponsible, sRoleList)
		S_NameControl = ReplaceRolesInList(S_NameControl, sRoleList)
		S_ListToEdit = ReplaceRolesInList(S_ListToEdit, sRoleList)
		S_ListToView = ReplaceRolesInList(S_ListToView, sRoleList)
		S_ListToReconcile = ReplaceRolesInList(S_ListToReconcile, sRoleList)
        S_ListToReconcile = Replace(S_ListToReconcile,"""Боев С. Ф."" <boev_oaorti>;","")

		'Удаляем повторы в списке согласования
		S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)

		'vnik_error
		'Убираем автора документа из списка согласующих, т.к. после активации он не сможет загрузить новый документ
		'(работает пока только, если пользователь создает документ не под ролью)
		S_ListToReconcile = Replace(S_ListToReconcile,Request("DocAuthor"),"")
		'vnik_error
	'Проверка полей с пользователями
	'Для заявок отключаем проверку утверждающего, т.к. он рассчитывается
	If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 Then
		UserChangeDocCheck = CheckSingleUserField(S_NameAproval)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameAproval")+SIT_ErrorInUserField2)
			Exit Function
		End If
		'Для заявки на закупку отключаем проверку ответственного, т.к. можно указывать подразделение
		'Для заявки на оплату тоже, т.к. автоматом ставится роль из специального справочника
		UserChangeDocCheck = CheckSingleUserField(S_NameResponsible)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameResponsible")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
	UserChangeDocCheck = CheckSingleUserField(S_NameControl)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameControl")+SIT_ErrorInUserField2)
		Exit Function
	End If
	UserChangeDocCheck = CheckMultiUserField(S_ListToEdit)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToEdit")+SIT_ErrorInUserField2)
		Exit Function
	End If
	
		'29-04-2013 проверка поля центр затрат и статья расходов

	sSQL = "select * from RTI_CostCenter where FullName = "+sUnicodeSymbol+"'"+ Request("UserFieldText2") + "'"
    Set dsTemp = Server.CreateObject("ADODB.Recordset")
    dsTemp.Open sSQL, Conn, 3, 1, &H1
    UserChangeDocCheck = not dsTemp.EOF
    dsTemp.Close
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+" Центр затрат " + SIT_ErrorInUserField2)	
		Exit Function
	End If


    'проверка статьи расходов
    if Request("UserFieldText4") = "" Then
      UserChangeDocCheck = true
    Else   
      sSQL = "select * from RTI_CostItem2 where Name = "+sUnicodeSymbol+"'"+ Request("UserFieldText4") + "'"
      Set dsTemp = Server.CreateObject("ADODB.Recordset")
      dsTemp.Open sSQL, Conn, 3, 1, &H1
      if not dsTemp.EOF Then
       isValid = dsTemp("isValid")
      End If       
      UserChangeDocCheck = not dsTemp.EOF    
      if IsValid = "1" Then
        UserChangeDocCheck = false
      End If
      dsTemp.Close 
    End If
    
	If not UserChangeDocCheck  Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+" Статья расходов " + SIT_ErrorInUserField2 + ". Либо вместо статьи выбрана категория статей (в этом случае выбранная строка не содержит номера статьи расходов)")
		Exit Function
	End If
	'проверка статьи расходов


	'29-04-2013 проверка поля центр затрат и статья расходов
	
	
'{ph - 20120812
	'S_AdditionalUsers = GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit"))
	S_AdditionalUsers = DeleteUserDoublesInListNew(Request("AdditionalUsers") & VbCrLf & GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit")), True)
'ph - 20120812}
	AddLogD "@@@Search For Leaders - S_AdditionalUsers: "+S_AdditionalUsers

	S_AdditionalUsers = DeleteUserDoublesInList(S_AdditionalUsers)

	'vnik
	If S_AdditionalUsers = "" Then
		S_AdditionalUsers = "-"
	End If
	S_AdditionalUsers_Set = S_AdditionalUsers

	S_ListToReconcile_Set = S_ListToReconcile
	S_NameAproval_Set = S_NameAproval
	S_ListToEdit = S_ListToEdit
	S_NameResponsible_Set = S_NameResponsible
	S_NameControl_Set = S_NameControl
	S_ListToView_Set=S_ListToView

	If UCase(Request("UpdateDoc")) = "YES" And (Request("create") = "y"  Or Request("DocIDPrev") <> Request("DocID")) Then
		'out "123"
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		dsTemp.Open "select * from Docs where DocID='" + MakeSQLSafe(Request("DocID")) + "'", Conn 
		If Not dsTemp.EOF Then
			Session("Message") = DOCS_ALREADYEXISTS + ": " + Request("DocID")
			dsTemp.Close
			UserChangeDocCheck=False
			Exit Function
		End If
		dsTemp.Close
	End If

	'ph - 20090808 - start
	If Request("DocIDParent") <> "" Then
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		bMyDocIDParentChanged = False
		If Request("create") <> "y" Then
			'Проверяем изменился ли номер родительского документа, если идет редактирование
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocID"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				bMyDocIDParentChanged = Request("DocIDParent")<>dsTemp("DocIDParent")
			End If
			dsTemp.Close
		End If
		If bMyDocIDParentChanged Then
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocIDParent"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				'sMessage = ""
				'UserChangeDocCheck = oPayDox.IsReadAccessUser(Session("UserID"), sMessage, dsTemp, DOCS_All, DOCS_NoReadAccess, DOCS_NoAccess, USER_Department, VAR_ExtInt, VAR_AdminSecLevel, VAR_StatusActiveUser, NULL)
				UserChangeDocCheck = IsReadAccessRS(dsTemp)' IsReadAccessUser(Session("UserID"), dsTemp, NULL)
				If UserChangeDocCheck Then
					S_DocIDParent = dsTemp("DocID")
				End If
				dsTemp.Close
				If not UserChangeDocCheck Then
					Session("Message") = SIT_NoAccessToParentDoc'sMessage
					Exit Function
				End If
			End If
		End If
	End If
	'ph - 20090808 - end

'rti_purchase_order


'vnik_purchase_order
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_PURCHASE_ORDER)) > 0 Then
	'Создатель документа
	If UCase(Request("create")) = "Y" Then
		sDocCreator = Session("UserID")
	Else
		sDocCreator = GetUserID(Request("DocAuthor"))
	End If

	S_NameAproval=Request("DocNameAproval")
	S_NameResponsible=Request("DocNameResponsible")
	S_NameControl=Request("DocNameControl")
	S_ListToReconcile=Request("DocListToReconcile")
	S_ListToEdit=Request("DocListToEdit")
	S_ListToView=Request("DocListToView")
	S_Correspondent=Request("DocCorrespondent")

	'Определяем бизнес направление
	sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))

	' SAY 2008-07-21
	Set dsTempPR = Server.CreateObject("ADODB.Recordset")

' DmGorsky_5: меняем местами проверку: вначале ищем SIT_SITRU ("СИТРОНИКС ИТ"), потом, если не найдено, - SIT_SITRONICS ("СИТРОНИКС")
	If InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRU))=1 then ' DmGorsky
		sDepartment = SIT_SITRU ' DmGorsky
		sDepartmentCode = "" ' DmGorsky
'Запрос №1 - СИБ - start
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SIB_ROOT_DEPARTMENT))=1 then
		sDepartment = SIT_SIB
		sDepartmentCode = ""
'Запрос №1 - СИБ - end
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRONICS))=1 then
		sDepartment = SIT_SITRONICS
		sDepartmentCode = ""
	Else
		sDepartment = SIT_STS

		sSQL = "select code from Departments where Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"' or Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"/'"
		AddLogD "WAY sSQL="+sSQL
		AddLogD "WAY Conn="+Conn
		dsTempPR.Open sSQL, Conn, 3, 1, &H1

		if not dsTempPR.EOF Then
			sDepartmentCode = dsTempPR("code")
			if InStrRev(sDepartmentCode, "/") > 0 then
				sDepartmentCode = Right(sDepartmentCode, Len(sDepartmentCode)-InStrRev(sDepartmentCode, "/"))
			End If
		Else
			sDepartmentCode="000000"
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInDepartmentCode1 + DelOtherLangFromFolder(Request("DocDepartment")) + SIT_ErrorInDepartmentCode2)
			UserChangeDocCheck=False
			Exit Function
		End If

		dsTempPR.Close
	End If

	If Request("create") = "y" Then
		sPostfix = ""
		sSearchCol = "DocID"
		sPrePrefix = ""
		'SAY 2009-02-20
		sSufix = ""

		'SAY 2009-03-12
		S_DocID = ""

		S_DocID = GetNewDocIDForPurchaseOrder(S_ClassDoc, sDepartmentRoot, "POHQ", "", "", "", "")
		S_DocIDAdd = S_DocID
	End If
	If sDepartmentRoot <> SIT_STS or (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_RASP_DOCS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1) Then
		'реализация полужесткого маршрута
		S_ListToReconcile = Replace(S_ListToReconcile, Trim(SIT_AdditionalAgrees), "")
		'Для коммерческих предложений и заявок на закупку и оплату не нужна приписка про доп. согласующих  
		If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_COM_OFFERS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) <> 1 Then
			If S_ListToReconcile <> "" Then
				S_ListToReconcile = " " + SIT_AdditionalAgrees + DeleteConstPrefixFromList(S_ListToReconcile)
			End If
		End If
	End If
'vnik_purchase_order


	If sDepartmentRoot = SIT_SITRONICS Then
		AddLogD "vnik123" + Trim(S_ListToReconcile) 
		'добавление обязательных согласующих
		'S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+Trim(S_ListToReconcile)
		S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+Trim(S_ListToReconcile)+VbCrLf+SIT_VicePresidentOfInitiator
		S_ListToReconcile = Replace(S_ListToReconcile,"""Хачатуров К. К."" <kkhachaturov>;","")
		'S_ListToView = S_ListToView + Replace(Replace(Replace(S_ListToReconcile,"Обязательные согласующие: ",""), "Дополнительные согласующие: ", ""), "##;", "")
		AddLogD "vnik123" + Trim(S_ListToReconcile)
		'Удаляем повторы в списке согласования
		S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)
		S_ListToView = DeleteUserDoublesInList(S_ListToView)

		S_NameControl = SIT_SecretaryCPC
	ElseIf sDepartmentRoot = SIT_STS Then
		'
	Else 'Другие БН
		'S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile+VbCrLf+SIT_VicePresidentOfInitiator
	End If
'vnik_purchase_order
'vnik_purchase_order
	'Подставляем название проекта
	Set dsTemp1 = Server.CreateObject("ADODB.Recordset")
	sSQL = "select * from MC_ProjectList where Code = "+sUnicodeSymbol+"'" + Request("UserFieldText1") + "'"
	AddLogD "@@@GetProjectName SQL: "+sSQL
	dsTemp1.Open sSQL, Conn, 3, 1, &H1
	If Not dsTemp1.EOF Then
		S_UserFieldText2 = Trim(CStr(dsTemp1("Name")))
		'S_UserFieldText2_Set = S_UserFieldText2
	Else 'Код проекта не соответствует справочнику
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInProjectNumber)
		UserChangeDocCheck = False
		dsTemp1.Close
		Exit Function
	End If
	dsTemp1.Close
'vnik_purchase_order
	'Для всех категорий кроме заявок подставляем пользователей вместо ролей
	If (InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1) and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) <> 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_DOGOVORI_NEW)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIS)) <> 0) Then
		'Получаем список ролей
		If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) = 1 Then 'Для договоров получаем список ролей на всех языках (могут быть роли на разных языках)
			sRoleList = GetFullRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"), "", "", "")
		Else
			sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))
		End If
		'Подставляем роли в поля
		S_NameAproval = ReplaceRolesInList(S_NameAproval, sRoleList)
		S_NameResponsible = ReplaceRolesInList(S_NameResponsible, sRoleList)
		S_NameControl = ReplaceRolesInList(S_NameControl, sRoleList)
		S_ListToEdit = ReplaceRolesInList(S_ListToEdit, sRoleList)
		S_ListToView = ReplaceRolesInList(S_ListToView, sRoleList)
		S_ListToReconcile = ReplaceRolesInList(S_ListToReconcile, sRoleList)

		'Удаляем повторы в списке согласования
		S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)

		'vnik_error
		'Убираем автора документа из списка согласующих, т.к. после активации он не сможет загрузить новый документ
		'(работает пока только, если пользователь создает документ не под ролью)
		S_ListToReconcile = Replace(S_ListToReconcile,Request("DocAuthor"),"")
		'vnik_error
	End If
	'Проверка полей с пользователями
	'Для заявок отключаем проверку утверждающего, т.к. он рассчитывается
	If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 Then
		UserChangeDocCheck = CheckSingleUserField(S_NameAproval)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameAproval")+SIT_ErrorInUserField2)
			Exit Function
		End If
		'Для заявки на закупку отключаем проверку ответственного, т.к. можно указывать подразделение
		'Для заявки на оплату тоже, т.к. автоматом ставится роль из специального справочника
		UserChangeDocCheck = CheckSingleUserField(S_NameResponsible)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameResponsible")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
	UserChangeDocCheck = CheckSingleUserField(S_NameControl)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameControl")+SIT_ErrorInUserField2)
		Exit Function
	End If
	If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) = 0 Then 'В исходящих может быть любой текст
		UserChangeDocCheck = CheckMultiUserField(S_Correspondent)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("Correspondent")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
'Запрос №1 - СИБ - start
	IF InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_SLUZH_ZAPISKA)) = 0) THEN
'Запрос №1 - СИБ - end
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) = 0 Then 'В СЗ на переработки работают правила
			UserChangeDocCheck = CheckMultiUserField(S_ListToReconcile)
			If not UserChangeDocCheck Then
				Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToReconcile")+SIT_ErrorInUserField2)
				Exit Function
			End If
'Запрос №1 - СИБ - start
		End If
	END IF
'Запрос №1 - СИБ - end
	UserChangeDocCheck = CheckMultiUserField(S_ListToEdit)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToEdit")+SIT_ErrorInUserField2)
		Exit Function
	End If
	'UserChangeDocCheck = CheckMultiUserField(S_ListToView)
	'If not UserChangeDocCheck Then
	'	Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToView")+SIT_ErrorInUserField2)
	'	Exit Function
	'End If

'{ph - 20120812
	'S_AdditionalUsers = GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit"))
	S_AdditionalUsers = DeleteUserDoublesInListNew(Request("AdditionalUsers") & VbCrLf & GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit")), True)
'ph - 20120812}
	AddLogD "@@@Search For Leaders - S_AdditionalUsers: "+S_AdditionalUsers
	'vnik_payment_order
	'If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_PAYMENT_ORDER)) > 0 Then
	'	S_AdditionalUsers = S_AdditionalUsers + ReplaceRoleFromDir(SIT_Account_manager,sDepartmentRoot)       
	'End If        
	'vnik_payment_order
	S_AdditionalUsers = DeleteUserDoublesInList(S_AdditionalUsers)

	'vnik
	If S_AdditionalUsers = "" Then
		S_AdditionalUsers = "-"
	End If
	S_AdditionalUsers_Set = S_AdditionalUsers

	S_ListToReconcile_Set = S_ListToReconcile
	S_NameAproval_Set = S_NameAproval
	S_ListToEdit = S_ListToEdit
	S_NameResponsible_Set = S_NameResponsible
	S_NameControl_Set = S_NameControl
	S_ListToView_Set=S_ListToView

	If UCase(Request("UpdateDoc")) = "YES" And (Request("create") = "y"  Or Request("DocIDPrev") <> Request("DocID")) Then
		'out "123"
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		dsTemp.Open "select * from Docs where DocID='" + MakeSQLSafe(Request("DocID")) + "'", Conn 
		If Not dsTemp.EOF Then
			Session("Message") = DOCS_ALREADYEXISTS + ": " + Request("DocID")
			dsTemp.Close
			UserChangeDocCheck=False
			Exit Function
		End If
		dsTemp.Close
	End If

	'ph - 20090808 - start
	If Request("DocIDParent") <> "" Then
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		bMyDocIDParentChanged = False
		If Request("create") <> "y" Then
			'Проверяем изменился ли номер родительского документа, если идет редактирование
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocID"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				bMyDocIDParentChanged = Request("DocIDParent")<>dsTemp("DocIDParent")
			End If
			dsTemp.Close
		End If
		If bMyDocIDParentChanged Then
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocIDParent"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				'sMessage = ""
				'UserChangeDocCheck = oPayDox.IsReadAccessUser(Session("UserID"), sMessage, dsTemp, DOCS_All, DOCS_NoReadAccess, DOCS_NoAccess, USER_Department, VAR_ExtInt, VAR_AdminSecLevel, VAR_StatusActiveUser, NULL)
				UserChangeDocCheck = IsReadAccessRS(dsTemp)' IsReadAccessUser(Session("UserID"), dsTemp, NULL)
				If UserChangeDocCheck Then
					S_DocIDParent = dsTemp("DocID")
				End If
				dsTemp.Close
				If not UserChangeDocCheck Then
					Session("Message") = SIT_NoAccessToParentDoc'sMessage
					Exit Function
				End If
			End If
		End If
	End If
	'ph - 20090808 - end

'vnik_contracts
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_CONTRACTS_MC)) > 0 Then
	'Создатель документа
	If UCase(Request("create")) = "Y" Then
		sDocCreator = Session("UserID")
	Else
		sDocCreator = GetUserID(Request("DocAuthor"))
	End If

	S_NameAproval=Request("DocNameAproval")
	S_NameResponsible=Request("DocNameResponsible")
	S_NameControl=Request("DocNameControl")
	S_ListToReconcile=Request("DocListToReconcile")
	S_ListToEdit=Request("DocListToEdit")
	S_ListToView=Request("DocListToView")
	S_Correspondent=Request("DocCorrespondent")

	'Определяем бизнес направление
	sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))

	' SAY 2008-07-21
	Set dsTempPR = Server.CreateObject("ADODB.Recordset")

' DmGorsky_5: меняем местами проверку: вначале ищем SIT_SITRU ("СИТРОНИКС ИТ"), потом, если не найдено, - SIT_SITRONICS ("СИТРОНИКС")
	If InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRU))=1 then ' DmGorsky
		sDepartment = SIT_SITRU ' DmGorsky
		sDepartmentCode = "" ' DmGorsky
'Запрос №1 - СИБ - start
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SIB_ROOT_DEPARTMENT))=1 then
		sDepartment = SIT_SIB
		sDepartmentCode = ""
'Запрос №1 - СИБ - end
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRONICS))=1 then
		sDepartment = SIT_SITRONICS
		sDepartmentCode = ""
	Else
		sDepartment = SIT_STS

		sSQL = "select code from Departments where Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"' or Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"/'"
		AddLogD "WAY sSQL="+sSQL
		AddLogD "WAY Conn="+Conn
		dsTempPR.Open sSQL, Conn, 3, 1, &H1

		if not dsTempPR.EOF Then
			sDepartmentCode = dsTempPR("code")
			if InStrRev(sDepartmentCode, "/") > 0 then
				sDepartmentCode = Right(sDepartmentCode, Len(sDepartmentCode)-InStrRev(sDepartmentCode, "/"))
			End If
		Else
			sDepartmentCode="000000"
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInDepartmentCode1 + DelOtherLangFromFolder(Request("DocDepartment")) + SIT_ErrorInDepartmentCode2)
			UserChangeDocCheck=False
			Exit Function
		End If

		dsTempPR.Close
	End If

	If Request("create") = "y" Then
		sPostfix = ""
		sSearchCol = "DocID"
		sPrePrefix = ""
		'SAY 2009-02-20
		sSufix = ""

		'SAY 2009-03-12
		S_DocID = ""

		S_DocID = GetNewDocIDForContractsMC(S_ClassDoc, sDepartmentRoot, "CNT-HQ", "", "", "", "")
		S_DocIDAdd = S_DocID
	End If
	If sDepartmentRoot <> SIT_STS or (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_RASP_DOCS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1) Then
		'реализация полужесткого маршрута
		S_ListToReconcile = Replace(S_ListToReconcile, Trim(SIT_AdditionalAgrees), "")
		'Для коммерческих предложений и заявок на закупку и оплату не нужна приписка про доп. согласующих  
		If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_COM_OFFERS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) <> 1 Then
			If S_ListToReconcile <> "" Then
				S_ListToReconcile = " " + SIT_AdditionalAgrees + DeleteConstPrefixFromList(S_ListToReconcile)
			End If
		End If
	End If
'vnik_contracts
	If sDepartmentRoot = SIT_SITRONICS Then
		AddLogD "vnik123" + Trim(S_ListToReconcile) 
		'добавление обязательных согласующих
		'S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+Trim(S_ListToReconcile)
		S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+Trim(S_ListToReconcile)
		'S_ListToView = S_ListToView + Replace(Replace(Replace(S_ListToReconcile,"Обязательные согласующие: ",""), "Дополнительные согласующие: ", ""), "##;", "")
		AddLogD "vnik123" + Trim(S_ListToReconcile)

		'Удаляем повторы в списке согласования
		S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)
		S_ListToView = DeleteUserDoublesInList(S_ListToView)
	ElseIf sDepartmentRoot = SIT_STS Then
		'
	Else 'Другие БН
		'S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile+VbCrLf+SIT_VicePresidentOfInitiator
	End If
'vnik_contracts
	'Для всех категорий кроме заявок подставляем пользователей вместо ролей
	If (InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1) and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) <> 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_DOGOVORI_NEW)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIS)) <> 0) Then
		'Получаем список ролей
		If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) = 1 Then 'Для договоров получаем список ролей на всех языках (могут быть роли на разных языках)
			sRoleList = GetFullRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"), "", "", "")
		Else
			sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))
		End If
		'Подставляем роли в поля
		S_NameAproval = ReplaceRolesInList(S_NameAproval, sRoleList)
		S_NameResponsible = ReplaceRolesInList(S_NameResponsible, sRoleList)
		S_NameControl = ReplaceRolesInList(S_NameControl, sRoleList)
		S_ListToEdit = ReplaceRolesInList(S_ListToEdit, sRoleList)
		S_ListToView = ReplaceRolesInList(S_ListToView, sRoleList)
		S_ListToReconcile = ReplaceRolesInList(S_ListToReconcile, sRoleList)
		S_ListToReconcile = Replace(S_ListToReconcile,"""Хачатуров К. К."" <kkhachaturov>;","")

		'Удаляем повторы в списке согласования
		S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)

		'vnik_error
		'Убираем автора документа из списка согласующих, т.к. после активации он не сможет загрузить новый документ
		'(работает пока только, если пользователь создает документ не под ролью)
		S_ListToReconcile = Replace(S_ListToReconcile,Request("DocAuthor"),"")
		'vnik_error
	End If
	'Проверка полей с пользователями
	'Для заявок отключаем проверку утверждающего, т.к. он рассчитывается
	If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 Then
		UserChangeDocCheck = CheckSingleUserField(S_NameAproval)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameAproval")+SIT_ErrorInUserField2)
			Exit Function
		End If
		'Для заявки на закупку отключаем проверку ответственного, т.к. можно указывать подразделение
		'Для заявки на оплату тоже, т.к. автоматом ставится роль из специального справочника
		UserChangeDocCheck = CheckSingleUserField(S_NameResponsible)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameResponsible")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
	UserChangeDocCheck = CheckSingleUserField(S_NameControl)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameControl")+SIT_ErrorInUserField2)
		Exit Function
	End If
	If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) = 0 Then 'В исходящих может быть любой текст
		UserChangeDocCheck = CheckMultiUserField(S_Correspondent)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("Correspondent")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
'Запрос №1 - СИБ - start
	IF InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_SLUZH_ZAPISKA)) = 0) THEN
'Запрос №1 - СИБ - end
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) = 0 Then 'В СЗ на переработки работают правила
			UserChangeDocCheck = CheckMultiUserField(S_ListToReconcile)
			If not UserChangeDocCheck Then
				Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToReconcile")+SIT_ErrorInUserField2)
				Exit Function
			End If
'Запрос №1 - СИБ - start
		End If
	END IF
'Запрос №1 - СИБ - end
	UserChangeDocCheck = CheckMultiUserField(S_ListToEdit)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToEdit")+SIT_ErrorInUserField2)
		Exit Function
	End If
	'UserChangeDocCheck = CheckMultiUserField(S_ListToView)
	'If not UserChangeDocCheck Then
	'	Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToView")+SIT_ErrorInUserField2)
	'	Exit Function
	'End If

'{ph - 20120812
	'S_AdditionalUsers = GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit"))
	S_AdditionalUsers = DeleteUserDoublesInListNew(Request("AdditionalUsers") & VbCrLf & GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit")), True)
'ph - 20120812}
	AddLogD "@@@Search For Leaders - S_AdditionalUsers: "+S_AdditionalUsers
	'vnik_payment_order
	'If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_PAYMENT_ORDER)) > 0 Then
	'	S_AdditionalUsers = S_AdditionalUsers + ReplaceRoleFromDir(SIT_Account_manager,sDepartmentRoot)       
	'End If        
	'vnik_payment_order
	S_AdditionalUsers = DeleteUserDoublesInList(S_AdditionalUsers)

	'vnik
	If S_AdditionalUsers = "" Then
		S_AdditionalUsers = "-"
	End If
	S_AdditionalUsers_Set = S_AdditionalUsers

	S_ListToReconcile_Set = S_ListToReconcile
	S_NameAproval_Set = S_NameAproval
	S_ListToEdit = S_ListToEdit
	S_NameResponsible_Set = S_NameResponsible
	S_NameControl_Set = S_NameControl
	S_ListToView_Set=S_ListToView

	If UCase(Request("UpdateDoc")) = "YES" And (Request("create") = "y"  Or Request("DocIDPrev") <> Request("DocID")) Then
		'out "123"
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		dsTemp.Open "select * from Docs where DocID='" + MakeSQLSafe(Request("DocID")) + "'", Conn 
		If Not dsTemp.EOF Then
			Session("Message") = DOCS_ALREADYEXISTS + ": " + Request("DocID")
			dsTemp.Close
			UserChangeDocCheck=False
			Exit Function
		End If
		dsTemp.Close
	End If

	'ph - 20090808 - start
	If Request("DocIDParent") <> "" Then
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		bMyDocIDParentChanged = False
		If Request("create") <> "y" Then
			'Проверяем изменился ли номер родительского документа, если идет редактирование
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocID"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				bMyDocIDParentChanged = Request("DocIDParent")<>dsTemp("DocIDParent")
			End If
			dsTemp.Close
		End If
		If bMyDocIDParentChanged Then
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocIDParent"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				'sMessage = ""
				'UserChangeDocCheck = oPayDox.IsReadAccessUser(Session("UserID"), sMessage, dsTemp, DOCS_All, DOCS_NoReadAccess, DOCS_NoAccess, USER_Department, VAR_ExtInt, VAR_AdminSecLevel, VAR_StatusActiveUser, NULL)
				UserChangeDocCheck = IsReadAccessRS(dsTemp)' IsReadAccessUser(Session("UserID"), dsTemp, NULL)
				If UserChangeDocCheck Then
					S_DocIDParent = dsTemp("DocID")
				End If
				dsTemp.Close
				If not UserChangeDocCheck Then
					Session("Message") = SIT_NoAccessToParentDoc'sMessage
					Exit Function
				End If
			End If
		End If
	End If
	'ph - 20090808 - end
' *** НОРМАТИВНЫЕ ДОКУМЕНТЫ
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_NORM_DOCS)) = 1 Then
	'Создатель документа
	If UCase(Request("create")) = "Y" Then
		sDocCreator = Session("UserID")
	Else
		sDocCreator = GetUserID(Request("DocAuthor"))
	End If

	S_NameAproval=Request("DocNameAproval")
	S_NameResponsible=Request("DocNameResponsible")
	S_NameControl=Request("DocNameControl")
	S_ListToReconcile=Request("DocListToReconcile")
	S_ListToEdit=Request("DocListToEdit")
	S_ListToView=Request("DocListToView")
	S_Correspondent=Request("DocCorrespondent")

	'Определяем бизнес направление
	sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))

	' SAY 2008-07-21
	Set dsTempPR = Server.CreateObject("ADODB.Recordset")

' DmGorsky_5: меняем местами проверку: вначале ищем SIT_SITRU ("СИТРОНИКС ИТ"), потом, если не найдено, - SIT_SITRONICS ("СИТРОНИКС")
	If InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRU))=1 then ' DmGorsky
		sDepartment = SIT_SITRU ' DmGorsky
		sDepartmentCode = "" ' DmGorsky
'Запрос №1 - СИБ - start
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SIB_ROOT_DEPARTMENT))=1 then
		sDepartment = SIT_SIB
		sDepartmentCode = ""
'Запрос №1 - СИБ - end
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRONICS))=1 then
		sDepartment = SIT_SITRONICS
		sDepartmentCode = ""
	Else
		sDepartment = SIT_STS

		sSQL = "select code from Departments where Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"' or Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"/'"
		AddLogD "WAY sSQL="+sSQL
		AddLogD "WAY Conn="+Conn
		dsTempPR.Open sSQL, Conn, 3, 1, &H1

		if not dsTempPR.EOF Then
			sDepartmentCode = dsTempPR("code")
			if InStrRev(sDepartmentCode, "/") > 0 then
				sDepartmentCode = Right(sDepartmentCode, Len(sDepartmentCode)-InStrRev(sDepartmentCode, "/"))
			End If
		Else
			sDepartmentCode="000000"
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInDepartmentCode1 + DelOtherLangFromFolder(Request("DocDepartment")) + SIT_ErrorInDepartmentCode2)
			UserChangeDocCheck=False
			Exit Function
		End If

		dsTempPR.Close
	End If

	If Request("create") = "y" Then
		sPostfix = ""
		sSearchCol = "DocID"
		sPrePrefix = ""
		'SAY 2009-02-20
		sSufix = ""

		'SAY 2009-03-12
		S_DocID = ""

		Call GetNewDocID_test(S_ClassDoc, sDepartmentRoot, Request("UserFieldText1"), Request("UserFieldText3"),  "", "PJ-")
		S_DocIDAdd = S_DocID
	End If
	If sDepartmentRoot <> SIT_STS or (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_RASP_DOCS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1) Then
		'реализация полужесткого маршрута
		S_ListToReconcile = Replace(S_ListToReconcile, Trim(SIT_AdditionalAgrees), "")
		'Для коммерческих предложений и заявок на закупку и оплату не нужна приписка про доп. согласующих  
		If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_COM_OFFERS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) <> 1 Then
			If S_ListToReconcile <> "" Then
				S_ListToReconcile = " " + SIT_AdditionalAgrees + DeleteConstPrefixFromList(S_ListToReconcile)
			End If
		End If
	End If
	If sDepartmentRoot = SIT_SITRONICS Then
		S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
	ElseIf sDepartmentRoot = SIT_STS Then
		'
	Else 'Другие БН
		S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
	End If
	'Для всех категорий кроме заявок подставляем пользователей вместо ролей
	If (InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1) and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) <> 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_DOGOVORI_NEW)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIS)) <> 0) Then
		'Получаем список ролей
		If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) = 1 Then 'Для договоров получаем список ролей на всех языках (могут быть роли на разных языках)
			sRoleList = GetFullRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"), "", "", "")
		Else
			sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))
		End If
		'Подставляем роли в поля
		S_NameAproval = ReplaceRolesInList(S_NameAproval, sRoleList)
		S_NameResponsible = ReplaceRolesInList(S_NameResponsible, sRoleList)
		S_NameControl = ReplaceRolesInList(S_NameControl, sRoleList)
		S_ListToEdit = ReplaceRolesInList(S_ListToEdit, sRoleList)
		S_ListToView = ReplaceRolesInList(S_ListToView, sRoleList)
		S_ListToReconcile = ReplaceRolesInList(S_ListToReconcile, sRoleList)

		'Удаляем повторы в списке согласования
		S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)

		'vnik_error
		'Убираем автора документа из списка согласующих, т.к. после активации он не сможет загрузить новый документ
		'(работает пока только, если пользователь создает документ не под ролью)
		S_ListToReconcile = Replace(S_ListToReconcile,Request("DocAuthor"),"")
		'vnik_error
	End If
	'Проверка полей с пользователями
	'Для заявок отключаем проверку утверждающего, т.к. он рассчитывается
	If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 Then
		UserChangeDocCheck = CheckSingleUserField(S_NameAproval)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameAproval")+SIT_ErrorInUserField2)
			Exit Function
		End If
		'Для заявки на закупку отключаем проверку ответственного, т.к. можно указывать подразделение
		'Для заявки на оплату тоже, т.к. автоматом ставится роль из специального справочника
		UserChangeDocCheck = CheckSingleUserField(S_NameResponsible)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameResponsible")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
	UserChangeDocCheck = CheckSingleUserField(S_NameControl)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameControl")+SIT_ErrorInUserField2)
		Exit Function
	End If
	If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) = 0 Then 'В исходящих может быть любой текст
		UserChangeDocCheck = CheckMultiUserField(S_Correspondent)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("Correspondent")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
'Запрос №1 - СИБ - start
	IF InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_SLUZH_ZAPISKA)) = 0) THEN
'Запрос №1 - СИБ - end
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) = 0 Then 'В СЗ на переработки работают правила
			UserChangeDocCheck = CheckMultiUserField(S_ListToReconcile)
			If not UserChangeDocCheck Then
				Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToReconcile")+SIT_ErrorInUserField2)
				Exit Function
			End If
'Запрос №1 - СИБ - start
		End If
	END IF
'Запрос №1 - СИБ - end
	UserChangeDocCheck = CheckMultiUserField(S_ListToEdit)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToEdit")+SIT_ErrorInUserField2)
		Exit Function
	End If
	'UserChangeDocCheck = CheckMultiUserField(S_ListToView)
	'If not UserChangeDocCheck Then
	'	Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToView")+SIT_ErrorInUserField2)
	'	Exit Function
	'End If

	' DmGorsky_7 Включать всех вышестоящих руководителей каждого из участников согласования документов
  	' в список дополнительных пользователей в СИТРУ не предполагалось
  	If InStr(UCase(Session("Department")), UCase(SIT_SITRU)) <> 1 Then ' DmGorsky_7 "СИТРОНИКС ИТ" не участвует
'{ph - 20120812
	'S_AdditionalUsers = GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit"))
	S_AdditionalUsers = DeleteUserDoublesInListNew(Request("AdditionalUsers") & VbCrLf & GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit")), True)
'ph - 20120812}
	End If ' DmGorsky_7

	AddLogD "@@@Search For Leaders - S_AdditionalUsers: "+S_AdditionalUsers
	'vnik_payment_order
	'If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_PAYMENT_ORDER)) > 0 Then
	'	S_AdditionalUsers = S_AdditionalUsers + ReplaceRoleFromDir(SIT_Account_manager,sDepartmentRoot)       
	'End If        
	'vnik_payment_order
	S_AdditionalUsers = DeleteUserDoublesInList(S_AdditionalUsers)

	'vnik
	If S_AdditionalUsers = "" Then
		S_AdditionalUsers = "-"
	End If
	S_AdditionalUsers_Set = S_AdditionalUsers

	'phil - 20080906 - Start - Проверка даты автоматического согласования
	If sDepartmentRoot = SIT_SITRONICS Then
		If Trim(Request("UserFieldDate3")) <> "" Then
			If ConvertToDate(Request("UserFieldDate3")) <> VAR_BeginOfTimes Then
				If ConvertToDate(Request("UserFieldDate3")) < ConvertToDate(GetNormDocLastReconcileDate) Then
					Session("Message") = SIT_TooEarlyAutoReconcile+Request("UserFieldDate3")
					UserChangeDocCheck = False
					Exit Function
				End If
			Else
				Session("Message") = SIT_DateFormatError+Request("UserFieldDate3")
				UserChangeDocCheck = False
				Exit Function
			End If
		End If
	End If
	'phil - 20080906 - End

	S_ListToReconcile_Set = S_ListToReconcile
	S_NameAproval_Set = S_NameAproval
	S_ListToEdit = S_ListToEdit
	S_NameResponsible_Set = S_NameResponsible
	S_NameControl_Set = S_NameControl
	S_ListToView_Set=S_ListToView

	If UCase(Request("UpdateDoc")) = "YES" And (Request("create") = "y"  Or Request("DocIDPrev") <> Request("DocID")) Then
		'out "123"
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		dsTemp.Open "select * from Docs where DocID='" + MakeSQLSafe(Request("DocID")) + "'", Conn 
		If Not dsTemp.EOF Then
			Session("Message") = DOCS_ALREADYEXISTS + ": " + Request("DocID")
			dsTemp.Close
			UserChangeDocCheck=False
			Exit Function
		End If
		dsTemp.Close
	End If

	'ph - 20090808 - start
	If Request("DocIDParent") <> "" Then
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		bMyDocIDParentChanged = False
		If Request("create") <> "y" Then
			'Проверяем изменился ли номер родительского документа, если идет редактирование
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocID"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				bMyDocIDParentChanged = Request("DocIDParent")<>dsTemp("DocIDParent")
			End If
			dsTemp.Close
		End If
		If bMyDocIDParentChanged Then
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocIDParent"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				'sMessage = ""
				'UserChangeDocCheck = oPayDox.IsReadAccessUser(Session("UserID"), sMessage, dsTemp, DOCS_All, DOCS_NoReadAccess, DOCS_NoAccess, USER_Department, VAR_ExtInt, VAR_AdminSecLevel, VAR_StatusActiveUser, NULL)
				UserChangeDocCheck = IsReadAccessRS(dsTemp)' IsReadAccessUser(Session("UserID"), dsTemp, NULL)
				If UserChangeDocCheck Then
					S_DocIDParent = dsTemp("DocID")
				End If
				dsTemp.Close
				If not UserChangeDocCheck Then
					Session("Message") = SIT_NoAccessToParentDoc'sMessage
					Exit Function
				End If
			End If
		End If
	End If
	'ph - 20090808 - end

' *** ПОРУЧЕНИЯ
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_ZADACHI)) = 1 Then
	'Создатель документа
	If UCase(Request("create")) = "Y" Then
		sDocCreator = Session("UserID")
	Else
		sDocCreator = GetUserID(Request("DocAuthor"))
	End If

	S_NameAproval=Request("DocNameAproval")
	S_NameResponsible=Request("DocNameResponsible")
	S_NameControl=Request("DocNameControl")
	S_ListToReconcile=Request("DocListToReconcile")
	S_ListToEdit=Request("DocListToEdit")
	S_ListToView=Request("DocListToView")
	S_Correspondent=Request("DocCorrespondent")

	'Определяем бизнес направление
	sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))

	' SAY 2008-07-21
	Set dsTempPR = Server.CreateObject("ADODB.Recordset")

' DmGorsky_5: меняем местами проверку: вначале ищем SIT_SITRU ("СИТРОНИКС ИТ"), потом, если не найдено, - SIT_SITRONICS ("СИТРОНИКС")
	If InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRU))=1 then ' DmGorsky
		sDepartment = SIT_SITRU ' DmGorsky
		sDepartmentCode = "" ' DmGorsky
'Запрос №1 - СИБ - start
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SIB_ROOT_DEPARTMENT))=1 then
		sDepartment = SIT_SIB
		sDepartmentCode = ""
'Запрос №1 - СИБ - end
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRONICS))=1 then
		sDepartment = SIT_SITRONICS
		sDepartmentCode = ""
	Else
		sDepartment = SIT_STS

		sSQL = "select code from Departments where Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"' or Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"/'"
		AddLogD "WAY sSQL="+sSQL
		AddLogD "WAY Conn="+Conn
		dsTempPR.Open sSQL, Conn, 3, 1, &H1

		if not dsTempPR.EOF Then
			sDepartmentCode = dsTempPR("code")
			if InStrRev(sDepartmentCode, "/") > 0 then
				sDepartmentCode = Right(sDepartmentCode, Len(sDepartmentCode)-InStrRev(sDepartmentCode, "/"))
			End If
		Else
			sDepartmentCode="000000"
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInDepartmentCode1 + DelOtherLangFromFolder(Request("DocDepartment")) + SIT_ErrorInDepartmentCode2)
			UserChangeDocCheck=False
			Exit Function
		End If

		dsTempPR.Close
	End If

	If Request("create") = "y" Then
		sPostfix = ""
		sSearchCol = "DocID"
		sPrePrefix = ""
		'SAY 2009-02-20
		sSufix = ""

		'SAY 2009-03-12
		S_DocID = ""

		S_DocID = GetNewDocIDForZadachi(Session("CurrentClassDoc"), sDepartmentRoot, Request("UserFieldText1"), Request("DocIDParent"), "")
		S_DocIDAdd = S_DocID
	End If
	If sDepartmentRoot <> SIT_STS or (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_RASP_DOCS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1) Then
		'реализация полужесткого маршрута
		S_ListToReconcile = Replace(S_ListToReconcile, Trim(SIT_AdditionalAgrees), "")
		'Для коммерческих предложений и заявок на закупку и оплату не нужна приписка про доп. согласующих  
		If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_COM_OFFERS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) <> 1 Then
			If S_ListToReconcile <> "" Then
				S_ListToReconcile = " " + SIT_AdditionalAgrees + DeleteConstPrefixFromList(S_ListToReconcile)
			End If
		End If
	End If
	'Для всех категорий кроме заявок подставляем пользователей вместо ролей
	If (InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1) and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) <> 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_DOGOVORI_NEW)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIS)) <> 0) Then
		'Получаем список ролей
		If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) = 1 Then 'Для договоров получаем список ролей на всех языках (могут быть роли на разных языках)
			sRoleList = GetFullRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"), "", "", "")
		Else
			sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))
		End If
		'Подставляем роли в поля
		S_NameAproval = ReplaceRolesInList(S_NameAproval, sRoleList)
		S_NameResponsible = ReplaceRolesInList(S_NameResponsible, sRoleList)
		S_NameControl = ReplaceRolesInList(S_NameControl, sRoleList)
		S_ListToEdit = ReplaceRolesInList(S_ListToEdit, sRoleList)
		S_ListToView = ReplaceRolesInList(S_ListToView, sRoleList)
		S_ListToReconcile = ReplaceRolesInList(S_ListToReconcile, sRoleList)

		'Удаляем повторы в списке согласования
		S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)

		'vnik_error
		'Убираем автора документа из списка согласующих, т.к. после активации он не сможет загрузить новый документ
		'(работает пока только, если пользователь создает документ не под ролью)
		S_ListToReconcile = Replace(S_ListToReconcile,Request("DocAuthor"),"")
		'vnik_error
	End If
	'Проверка полей с пользователями
	'Для заявок отключаем проверку утверждающего, т.к. он рассчитывается
	If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 Then
		UserChangeDocCheck = CheckSingleUserField(S_NameAproval)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameAproval")+SIT_ErrorInUserField2)
			Exit Function
		End If
		'Для заявки на закупку отключаем проверку ответственного, т.к. можно указывать подразделение
		'Для заявки на оплату тоже, т.к. автоматом ставится роль из специального справочника
		UserChangeDocCheck = CheckSingleUserField(S_NameResponsible)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameResponsible")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
	UserChangeDocCheck = CheckSingleUserField(S_NameControl)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameControl")+SIT_ErrorInUserField2)
		Exit Function
	End If
	If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) = 0 Then 'В исходящих может быть любой текст
		UserChangeDocCheck = CheckMultiUserField(S_Correspondent)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("Correspondent")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
'Запрос №1 - СИБ - start
	IF InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_SLUZH_ZAPISKA)) = 0) THEN
'Запрос №1 - СИБ - end
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) = 0 Then 'В СЗ на переработки работают правила
			UserChangeDocCheck = CheckMultiUserField(S_ListToReconcile)
			If not UserChangeDocCheck Then
				Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToReconcile")+SIT_ErrorInUserField2)
				Exit Function
			End If
'Запрос №1 - СИБ - start
		End If
	END IF
'Запрос №1 - СИБ - end
	UserChangeDocCheck = CheckMultiUserField(S_ListToEdit)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToEdit")+SIT_ErrorInUserField2)
		Exit Function
	End If
	'UserChangeDocCheck = CheckMultiUserField(S_ListToView)
	'If not UserChangeDocCheck Then
	'	Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToView")+SIT_ErrorInUserField2)
	'	Exit Function
	'End If

'{ph - 20120812
	'S_AdditionalUsers = GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit"))
	S_AdditionalUsers = DeleteUserDoublesInListNew(Request("AdditionalUsers") & VbCrLf & GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit")), True)
'ph - 20120812}
	AddLogD "@@@Search For Leaders - S_AdditionalUsers: "+S_AdditionalUsers
	'vnik_payment_order
	'If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_PAYMENT_ORDER)) > 0 Then
	'	S_AdditionalUsers = S_AdditionalUsers + ReplaceRoleFromDir(SIT_Account_manager,sDepartmentRoot)       
	'End If        
	'vnik_payment_order
	S_AdditionalUsers = DeleteUserDoublesInList(S_AdditionalUsers)

	'vnik - добавим к доп. пользователям подчиненного поручения исполнителя по родительскому поручению 
	Set dsTempPRVNIK = Server.CreateObject("ADODB.Recordset")
	VNIK_Tmp_DocIDParent = Trim(Request("DocIDParent"))
	Do While True
		If VNIK_Tmp_DocIDParent = "" Then
			Exit Do
		End If

		sSQL = "select NameResponsible,DocIDParent from Docs where DocID=N'"+VNIK_Tmp_DocIDParent+"' and ClassDoc like N'Поручения*%'"
		dsTempPRVNIK.Open sSQL, Conn, 3, 1, &H1
		If not dsTempPRVNIK.EOF Then
			S_AdditionalUsers = S_AdditionalUsers +VbCrLf+ Trim(dsTempPRVNIK("NameResponsible"))
			VNIK_Tmp_DocIDParent = Trim(dsTempPRVNIK("DocIDParent"))
			AddLogD "@@@Add Parents Responsibles - S_AdditionalUsers: "+S_AdditionalUsers
		Else
			Exit Do
			AddLogD "@@@VNIK_ERROR Add Parents Responsibles - S_AdditionalUsers: No parent doc class like Поручение" + Trim(VNIK_Tmp_DocIDParent)
		End If    
		dsTempPRVNIK.Close
	Loop

	'vnik
	If S_AdditionalUsers = "" Then
		S_AdditionalUsers = "-"
	End If
	S_AdditionalUsers_Set = S_AdditionalUsers

	'SAY 2008-09-16 проверка даты исполнения
	If ConvertToDate(Request("DocDateCompletion")) < ConvertToDate(Request("DocDateActivation")) Then
		Session("Message") = SIT_ErrorInDateCompletion+Request("DocDateCompletion")
		UserChangeDocCheck = False
		Exit Function
	End If

	S_ListToReconcile_Set = S_ListToReconcile
	S_NameAproval_Set = S_NameAproval
	S_ListToEdit = S_ListToEdit
	S_NameResponsible_Set = S_NameResponsible
	S_NameControl_Set = S_NameControl
	S_ListToView_Set=S_ListToView

	If UCase(Request("UpdateDoc")) = "YES" And (Request("create") = "y"  Or Request("DocIDPrev") <> Request("DocID")) Then
	    'out "123"
		'out UCase(Request("UpdateDoc"))
		'out Request("create")
		'out Request("DocIDPrev")
		'out Request("DocID")
		'out "select * from Docs where DocID='" + MakeSQLSafe(Request("DocID")) + "'"
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		dsTemp.Open "select * from Docs where DocID='" + MakeSQLSafe(Request("DocID")) + "'", Conn 
		If Not dsTemp.EOF Then
			Session("Message") = DOCS_ALREADYEXISTS + ": " + Request("DocID")
			dsTemp.Close
			UserChangeDocCheck=False
			Exit Function
		End If
		dsTemp.Close
	End If

	'ph - 20090808 - start
	If Request("DocIDParent") <> "" Then
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		bMyDocIDParentChanged = False
		If Request("create") <> "y" Then
			'Проверяем изменился ли номер родительского документа, если идет редактирование
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocID"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				bMyDocIDParentChanged = Request("DocIDParent")<>dsTemp("DocIDParent")
			End If
			dsTemp.Close
		End If
		If bMyDocIDParentChanged Then
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocIDParent"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				'sMessage = ""
				'UserChangeDocCheck = oPayDox.IsReadAccessUser(Session("UserID"), sMessage, dsTemp, DOCS_All, DOCS_NoReadAccess, DOCS_NoAccess, USER_Department, VAR_ExtInt, VAR_AdminSecLevel, VAR_StatusActiveUser, NULL)
				UserChangeDocCheck = IsReadAccessRS(dsTemp)' IsReadAccessUser(Session("UserID"), dsTemp, NULL)
				If UserChangeDocCheck Then
					S_DocIDParent = dsTemp("DocID")
				End If
				dsTemp.Close
				If not UserChangeDocCheck Then
					Session("Message") = SIT_NoAccessToParentDoc'sMessage
					Exit Function
				End If
			End If
		End If
	End If
	'ph - 20090808 - end

' *** ДОГОВОРЫ ДО ДАТЫ ...
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_OLD)) = 1 Then
	'Запрет создания документов в категории
	If UCase(Request("create")) = "Y" Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_CannotCreateOldContracts)
		UserChangeDocCheck = False
		Exit Function
	End If
	'Создатель документа
	If UCase(Request("create")) = "Y" Then
		sDocCreator = Session("UserID")
	Else
		sDocCreator = GetUserID(Request("DocAuthor"))
	End If

	S_NameAproval=Request("DocNameAproval")
	S_NameResponsible=Request("DocNameResponsible")
	S_NameControl=Request("DocNameControl")
	S_ListToReconcile=Request("DocListToReconcile")
	S_ListToEdit=Request("DocListToEdit")
	S_ListToView=Request("DocListToView")
	S_Correspondent=Request("DocCorrespondent")

	'Определяем бизнес направление
	sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))

	' SAY 2008-07-21
	Set dsTempPR = Server.CreateObject("ADODB.Recordset")

' DmGorsky_5: меняем местами проверку: вначале ищем SIT_SITRU ("СИТРОНИКС ИТ"), потом, если не найдено, - SIT_SITRONICS ("СИТРОНИКС")
	If InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRU))=1 then ' DmGorsky
		sDepartment = SIT_SITRU ' DmGorsky
		sDepartmentCode = "" ' DmGorsky
'Запрос №1 - СИБ - start
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SIB_ROOT_DEPARTMENT))=1 then
		sDepartment = SIT_SIB
		sDepartmentCode = ""
'Запрос №1 - СИБ - end
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRONICS))=1 then
		sDepartment = SIT_SITRONICS
		sDepartmentCode = ""
	Else
		sDepartment = SIT_STS

		sSQL = "select code from Departments where Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"' or Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"/'"
		AddLogD "WAY sSQL="+sSQL
		AddLogD "WAY Conn="+Conn
		dsTempPR.Open sSQL, Conn, 3, 1, &H1

		if not dsTempPR.EOF Then
			sDepartmentCode = dsTempPR("code")
			if InStrRev(sDepartmentCode, "/") > 0 then
				sDepartmentCode = Right(sDepartmentCode, Len(sDepartmentCode)-InStrRev(sDepartmentCode, "/"))
			End If
		Else
			sDepartmentCode="000000"
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInDepartmentCode1 + DelOtherLangFromFolder(Request("DocDepartment")) + SIT_ErrorInDepartmentCode2)
			UserChangeDocCheck=False
			Exit Function
		End If

		dsTempPR.Close
	End If

	If Request("create") = "y" Then
		sPostfix = ""
		sSearchCol = "DocID"
		sPrePrefix = ""
		'SAY 2009-02-20
		sSufix = ""

		'SAY 2009-03-12
		S_DocID = ""

		Call GetNewDocID_test(S_ClassDoc, sDepartmentRoot, Request("DocName"), "",  "", "PJ-")
		S_DocIDAdd = S_DocID
	End If
	If sDepartmentRoot <> SIT_STS or (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_RASP_DOCS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1) Then
		'реализация полужесткого маршрута
		S_ListToReconcile = Replace(S_ListToReconcile, Trim(SIT_AdditionalAgrees), "")
		'Для коммерческих предложений и заявок на закупку и оплату не нужна приписка про доп. согласующих  
		If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_COM_OFFERS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) <> 1 Then
			If S_ListToReconcile <> "" Then
				S_ListToReconcile = " " + SIT_AdditionalAgrees + DeleteConstPrefixFromList(S_ListToReconcile)
			End If
		End If
	End If
	If sDepartmentRoot = SIT_SITRONICS Then
		S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
	ElseIf sDepartmentRoot = SIT_STS Then
		'20091110 - start - В договорах при редактировании жесткую часть не обновляем, могло быть делегирование
		If UCase(Request("create")) = "Y" Then
			'rmanyushin 60298 02.11.2009 Start
			'Если в значение поля ContractType, то использовать расширенный список согласующих 
			strContactType = Request("ContractType")
			If strContactType = "Физическое лицо" or strContactType = "Natural person" or strContactType = "Fyzická osoba" Then
				S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+STS_HRAdministrationManager+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
			Else
				S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
			End If
			'rmanyushin 60298 02.11.2009 End
		Else
			S_ListToReconcile = Session("S_ListToReconcile_Comment")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
		End If
		'20091110 - end
	Else 'Другие БН
		S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
	End If
	'Для всех категорий кроме заявок подставляем пользователей вместо ролей
	If (InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1) and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) <> 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_DOGOVORI_NEW)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIS)) <> 0) Then
		'Получаем список ролей
		If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) = 1 Then 'Для договоров получаем список ролей на всех языках (могут быть роли на разных языках)
			sRoleList = GetFullRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"), "", "", "")
		Else
			sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))
		End If
		'Подставляем роли в поля
		S_NameAproval = ReplaceRolesInList(S_NameAproval, sRoleList)
		S_NameResponsible = ReplaceRolesInList(S_NameResponsible, sRoleList)
		S_NameControl = ReplaceRolesInList(S_NameControl, sRoleList)
		S_ListToEdit = ReplaceRolesInList(S_ListToEdit, sRoleList)
		S_ListToView = ReplaceRolesInList(S_ListToView, sRoleList)
		S_ListToReconcile = ReplaceRolesInList(S_ListToReconcile, sRoleList)

		'Удаляем повторы в списке согласования
		S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)

		'vnik_error
		'Убираем автора документа из списка согласующих, т.к. после активации он не сможет загрузить новый документ
		'(работает пока только, если пользователь создает документ не под ролью)
		S_ListToReconcile = Replace(S_ListToReconcile,Request("DocAuthor"),"")
		'vnik_error
	End If
	'Проверка полей с пользователями
	'Для заявок отключаем проверку утверждающего, т.к. он рассчитывается
	If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 Then
		UserChangeDocCheck = CheckSingleUserField(S_NameAproval)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameAproval")+SIT_ErrorInUserField2)
			Exit Function
		End If
		'Для заявки на закупку отключаем проверку ответственного, т.к. можно указывать подразделение
		'Для заявки на оплату тоже, т.к. автоматом ставится роль из специального справочника
		UserChangeDocCheck = CheckSingleUserField(S_NameResponsible)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameResponsible")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
	UserChangeDocCheck = CheckSingleUserField(S_NameControl)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameControl")+SIT_ErrorInUserField2)
		Exit Function
	End If
	If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) = 0 Then 'В исходящих может быть любой текст
		UserChangeDocCheck = CheckMultiUserField(S_Correspondent)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("Correspondent")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
'Запрос №1 - СИБ - start
	IF InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_SLUZH_ZAPISKA)) = 0) THEN
'Запрос №1 - СИБ - end
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) = 0 Then 'В СЗ на переработки работают правила
			UserChangeDocCheck = CheckMultiUserField(S_ListToReconcile)
			If not UserChangeDocCheck Then
				Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToReconcile")+SIT_ErrorInUserField2)
				Exit Function
			End If
'Запрос №1 - СИБ - start
		End If
	END IF
'Запрос №1 - СИБ - end
	UserChangeDocCheck = CheckMultiUserField(S_ListToEdit)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToEdit")+SIT_ErrorInUserField2)
		Exit Function
	End If
	'UserChangeDocCheck = CheckMultiUserField(S_ListToView)
	'If not UserChangeDocCheck Then
	'	Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToView")+SIT_ErrorInUserField2)
	'	Exit Function
	'End If

'{ph - 20120812
	'S_AdditionalUsers = GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit"))
	S_AdditionalUsers = DeleteUserDoublesInListNew(Request("AdditionalUsers") & VbCrLf & GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit")), True)
'ph - 20120812}
	AddLogD "@@@Search For Leaders - S_AdditionalUsers: "+S_AdditionalUsers
	'vnik_payment_order
	'If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_PAYMENT_ORDER)) > 0 Then
	'	S_AdditionalUsers = S_AdditionalUsers + ReplaceRoleFromDir(SIT_Account_manager,sDepartmentRoot)       
	'End If        
	'vnik_payment_order
	S_AdditionalUsers = DeleteUserDoublesInList(S_AdditionalUsers)

	'vnik
	If S_AdditionalUsers = "" Then
		S_AdditionalUsers = "-"
	End If
	S_AdditionalUsers_Set = S_AdditionalUsers

	S_ListToReconcile_Set = S_ListToReconcile
	S_NameAproval_Set = S_NameAproval
	S_ListToEdit = S_ListToEdit
	S_NameResponsible_Set = S_NameResponsible
	S_NameControl_Set = S_NameControl
	S_ListToView_Set=S_ListToView

	If UCase(Request("UpdateDoc")) = "YES" And (Request("create") = "y"  Or Request("DocIDPrev") <> Request("DocID")) Then
		'out "123"
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		dsTemp.Open "select * from Docs where DocID='" + MakeSQLSafe(Request("DocID")) + "'", Conn 
		If Not dsTemp.EOF Then
			Session("Message") = DOCS_ALREADYEXISTS + ": " + Request("DocID")
			dsTemp.Close
			UserChangeDocCheck=False
			Exit Function
		End If
		dsTemp.Close
	End If

	'ph - 20090808 - start
	If Request("DocIDParent") <> "" Then
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		bMyDocIDParentChanged = False
		If Request("create") <> "y" Then
			'Проверяем изменился ли номер родительского документа, если идет редактирование
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocID"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				bMyDocIDParentChanged = Request("DocIDParent")<>dsTemp("DocIDParent")
			End If
			dsTemp.Close
		End If
		If bMyDocIDParentChanged Then
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocIDParent"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				'sMessage = ""
				'UserChangeDocCheck = oPayDox.IsReadAccessUser(Session("UserID"), sMessage, dsTemp, DOCS_All, DOCS_NoReadAccess, DOCS_NoAccess, USER_Department, VAR_ExtInt, VAR_AdminSecLevel, VAR_StatusActiveUser, NULL)
				UserChangeDocCheck = IsReadAccessRS(dsTemp)' IsReadAccessUser(Session("UserID"), dsTemp, NULL)
				If UserChangeDocCheck Then
					S_DocIDParent = dsTemp("DocID")
				End If
				dsTemp.Close
				If not UserChangeDocCheck Then
					Session("Message") = SIT_NoAccessToParentDoc'sMessage
					Exit Function
				End If
			End If
		End If
	End If
	'ph - 20090808 - end

'rti_contract
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(RTI_CONTRACT)) = 1 Then
	'Создатель документа
	If UCase(Request("create")) = "Y" Then
		sDocCreator = Session("UserID")
	Else
		sDocCreator = GetUserID(Request("DocAuthor"))
	End If

	S_NameAproval=Request("DocNameAproval")
	S_NameResponsible=Request("DocNameResponsible")
	S_NameControl=Request("DocNameControl")
	S_ListToReconcile=Request("DocListToReconcile")
	S_ListToEdit=Request("DocListToEdit")
	S_ListToView=Request("DocListToView")
	S_Correspondent=Request("DocCorrespondent")

	'Определяем бизнес направление
	sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))
	sDepartment = SIT_RTI
	sDepartmentCode = ""

	If Request("create") = "y" Then
		sPostfix = ""
		sSearchCol = "DocID"
		sPrePrefix = ""
		'SAY 2009-02-20
		sSufix = ""
        S_ClassDoc = UCase(Session("CurrentClassDoc"))
		'SAY 2009-03-12
		S_DocID = ""
		S_DocID = GetNewDocIDForPaymentOrder(S_ClassDoc, sDepartmentRoot, "CNTRTI", "", "", "", "")
		S_DocIDAdd = S_DocID
	   '2013-09-03 добавляем в получатели правовое управление
	    S_ListToView = S_ListToView + " " + RTI_ContractViewList

	End If
	
	
	If sDepartmentRoot <> SIT_STS or (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_RASP_DOCS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1) Then
		'реализация полужесткого маршрута
		S_ListToReconcile = Replace(S_ListToReconcile, Trim(SIT_AdditionalAgrees), "")
		'Для коммерческих предложений и заявок на закупку и оплату не нужна приписка про доп. согласующих  
		If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_COM_OFFERS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) <> 1 Then
			If S_ListToReconcile <> "" Then
				S_ListToReconcile = " " + SIT_AdditionalAgrees + DeleteConstPrefixFromList(S_ListToReconcile)
			End If
		End If
	End If
	

	'Формируем список согласования
			'S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeRTI"+Request("l"),"Category",Session("CurrentClassDoc"),"List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
		 If Request("create") = "y" Then
            S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeRTI"+Request("l"),"Category",Session("CurrentClassDoc"),"List")+SIT_AdditionalAgreesDelimeter+Trim(S_ListToReconcile)+vbcrlf+"""Гартван К. Р."" <gartvan_oaorti>;"+vbcrlf+RTI_DirectorOfSecurity+RTI_HeadOfUpravDelami+RTI_HeadOfAccounting+RTI_HeadKFIE+RTI_HeadOfPurchaseCenter+RTI_HeadPriceforming+RTI_BudgetController+vbcrlf+SIT_RTI_DirectorPravovogoUprav_RU
         Else
            AddLogD "vnik01234 " + Trim(S_ListToReconcile)
            sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))
            S_ListToReconcile = replace(S_ListToReconcile, ReplaceRolesInList(RTI_DirectorOfSecurity, sRoleList), "")
            S_ListToReconcile = replace(S_ListToReconcile, ReplaceRolesInList(RTI_HeadOfUpravDelami, sRoleList), "")
            S_ListToReconcile = replace(S_ListToReconcile, ReplaceRolesInList(RTI_HeadOfAccounting, sRoleList), "")
            S_ListToReconcile = replace(S_ListToReconcile, ReplaceRolesInList(RTI_HeadKFIE, sRoleList), "")
            S_ListToReconcile = replace(S_ListToReconcile, ReplaceRolesInList(RTI_HeadOfPurchaseCenter, sRoleList), "")
            S_ListToReconcile = replace(S_ListToReconcile, ReplaceRolesInList(SIT_RTI_DirectorPravovogoUprav_RU, sRoleList), "")
            S_ListToReconcile = replace(S_ListToReconcile, """Гартван К. Р."" <gartvan_oaorti>;", "")
            S_ListToReconcile = replace(S_ListToReconcile, ReplaceRolesInList(RTI_BudgetController, sRoleList), "")
            S_ListToReconcile = replace(S_ListToReconcile, ReplaceRolesInList(RTI_HeadPriceforming, sRoleList), "")
            S_ListToReconcile = replace(S_ListToReconcile, VbCrLf, "")
            S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeRTI"+Request("l"),"Category",Session("CurrentClassDoc"),"List")+SIT_AdditionalAgreesDelimeter+Trim(S_ListToReconcile)+vbcrlf+"""Гартван К. Р."" <gartvan_oaorti>;"+vbcrlf+RTI_DirectorOfSecurity+RTI_HeadOfUpravDelami+RTI_HeadOfAccounting+RTI_HeadKFIE+RTI_HeadOfPurchaseCenter+RTI_HeadPriceforming+RTI_BudgetController+vbcrlf+SIT_RTI_DirectorPravovogoUprav_RU
         End If
         S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)
 
	'Для всех категорий кроме заявок подставляем пользователей вместо ролей
	If (InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1) and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) <> 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_DOGOVORI_NEW)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIS)) <> 0) Then
		'Получаем список ролей
		If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) = 1 Then 'Для договоров получаем список ролей на всех языках (могут быть роли на разных языках)
			sRoleList = GetFullRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"), "", "", "")
		Else
			sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))
		End If
		'Подставляем роли в поля
		S_NameAproval = ReplaceRolesInList(S_NameAproval, sRoleList)
		S_NameResponsible = ReplaceRolesInList(S_NameResponsible, sRoleList)
		S_NameControl = ReplaceRolesInList(S_NameControl, sRoleList)
		S_ListToEdit = ReplaceRolesInList(S_ListToEdit, sRoleList)
		S_ListToView = ReplaceRolesInList(S_ListToView, sRoleList)
		S_ListToReconcile = ReplaceRolesInList(S_ListToReconcile, sRoleList)
		S_ListToReconcile = Replace(S_ListToReconcile,"""Боев С. Ф."" <boev_oaorti>;","")

		'Удаляем повторы в списке согласования
		

		'vnik_error
		'Убираем автора документа из списка согласующих, т.к. после активации он не сможет загрузить новый документ
		'(работает пока только, если пользователь создает документ не под ролью)
		S_ListToReconcile = Replace(S_ListToReconcile,Request("DocAuthor"),"")
		If Request("create") <> "y" Then
		  'S_ListToReconcile = replace(S_ListToReconcile, ReplaceRolesInList(RTI_DirectorOfSecurity, sRoleList), "")
		  'S_ListToReconcile = replace(S_ListToReconcile, VbCrLf, "")
		end if
		S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)
		S_ListToView = DeleteUserDoublesInList(S_ListToView)
		'vnik_error
	End If

'{ph - 20120812
	S_AdditionalUsers = DeleteUserDoublesInListNew(Request("AdditionalUsers") & VbCrLf & GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit")), True)
'ph - 20120812}
	AddLogD "@@@Search For Leaders - S_AdditionalUsers: "+S_AdditionalUsers
	'vnik_payment_order
	'If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_PAYMENT_ORDER)) > 0 Then
	'	S_AdditionalUsers = S_AdditionalUsers + ReplaceRoleFromDir(SIT_Account_manager,sDepartmentRoot)       
	'End If        
	'vnik_payment_order
	S_AdditionalUsers = DeleteUserDoublesInList(S_AdditionalUsers)

	'vnik
	If S_AdditionalUsers = "" Then
		S_AdditionalUsers = "-"
	End If
	S_AdditionalUsers_Set = S_AdditionalUsers

	S_ListToReconcile_Set = S_ListToReconcile
	S_NameAproval_Set = S_NameAproval
	S_ListToEdit = S_ListToEdit
	S_NameResponsible_Set = S_NameResponsible
	S_NameControl_Set = S_NameControl
	S_ListToView_Set=S_ListToView

	If UCase(Request("UpdateDoc")) = "YES" And (Request("create") = "y"  Or Request("DocIDPrev") <> Request("DocID")) Then
		'out "123"
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		dsTemp.Open "select * from Docs where DocID='" + MakeSQLSafe(Request("DocID")) + "'", Conn 
		If Not dsTemp.EOF Then
			Session("Message") = DOCS_ALREADYEXISTS + ": " + Request("DocID")
			dsTemp.Close
			UserChangeDocCheck=False
			Exit Function
		End If
		dsTemp.Close
	End If

	'ph - 20090808 - start
	If Request("DocIDParent") <> "" Then
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		bMyDocIDParentChanged = False
		If Request("create") <> "y" Then
			'Проверяем изменился ли номер родительского документа, если идет редактирование
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocID"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				bMyDocIDParentChanged = Request("DocIDParent")<>dsTemp("DocIDParent")
			End If
			dsTemp.Close
		End If
		If bMyDocIDParentChanged Then
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocIDParent"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				'sMessage = ""
				'UserChangeDocCheck = oPayDox.IsReadAccessUser(Session("UserID"), sMessage, dsTemp, DOCS_All, DOCS_NoReadAccess, DOCS_NoAccess, USER_Department, VAR_ExtInt, VAR_AdminSecLevel, VAR_StatusActiveUser, NULL)
				UserChangeDocCheck = IsReadAccessRS(dsTemp)' IsReadAccessUser(Session("UserID"), dsTemp, NULL)
				If UserChangeDocCheck Then
					S_DocIDParent = dsTemp("DocID")
				End If
				dsTemp.Close
				If not UserChangeDocCheck Then
					Session("Message") = SIT_NoAccessToParentDoc'sMessage
					Exit Function
				End If
			End If
		End If
	End If
	'ph - 20090808 - end
'rti_contract


' *** ДОГОВОРЫ
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) = 1 Then
	'Создатель документа
	If UCase(Request("create")) = "Y" Then
		sDocCreator = Session("UserID")
	Else
		sDocCreator = GetUserID(Request("DocAuthor"))
	End If

	S_NameAproval=Request("DocNameAproval")
	S_NameResponsible=Request("DocNameResponsible")
	S_NameControl=Request("DocNameControl")
	S_ListToReconcile=Request("DocListToReconcile")
	S_ListToEdit=Request("DocListToEdit")
	S_ListToView=Request("DocListToView")
	S_Correspondent=Request("DocCorrespondent")

	'Определяем бизнес направление
	sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))

	' SAY 2008-07-21
	Set dsTempPR = Server.CreateObject("ADODB.Recordset")

' DmGorsky_5: меняем местами проверку: вначале ищем SIT_SITRU ("СИТРОНИКС ИТ"), потом, если не найдено, - SIT_SITRONICS ("СИТРОНИКС")
	If InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRU))=1 then ' DmGorsky
		sDepartment = SIT_SITRU ' DmGorsky
		sDepartmentCode = "" ' DmGorsky
'Запрос №1 - СИБ - start
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SIB_ROOT_DEPARTMENT))=1 then
		sDepartment = SIT_SIB
		sDepartmentCode = ""
'Запрос №1 - СИБ - end
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRONICS))=1 then
		sDepartment = SIT_SITRONICS
		sDepartmentCode = ""
	Else
		sDepartment = SIT_STS

		sSQL = "select code from Departments where Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"' or Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"/'"
		AddLogD "WAY sSQL="+sSQL
		AddLogD "WAY Conn="+Conn
		dsTempPR.Open sSQL, Conn, 3, 1, &H1

		if not dsTempPR.EOF Then
			sDepartmentCode = dsTempPR("code")
			if InStrRev(sDepartmentCode, "/") > 0 then
				sDepartmentCode = Right(sDepartmentCode, Len(sDepartmentCode)-InStrRev(sDepartmentCode, "/"))
			End If
		Else
			sDepartmentCode="000000"
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInDepartmentCode1 + DelOtherLangFromFolder(Request("DocDepartment")) + SIT_ErrorInDepartmentCode2)
			UserChangeDocCheck=False
			Exit Function
		End If

		dsTempPR.Close
	End If

	If Request("create") = "y" Then
		sPostfix = ""
		sSearchCol = "DocID"
		sPrePrefix = ""
		'SAY 2009-02-20
		sSufix = ""

		'SAY 2009-03-12
		S_DocID = ""

		'sContractPartnerCode используется дальше при проверке наличия кода
		sContractPartnerCode = Trim(MyCStr(GetPartnerCode(Request("DocPartnerName"))))
'{ph - Запрос №47 - СТС
		sDZKCode = Trim(GetDZKCode(GetCodeFromCode_NameString(Request("BusinessUnit")), "ContractCode"))
		If sDZKCode <> "" Then
			sDZKCode = "_" & sDZKCode
		End If
		S_DocID = GetNewDocIDForContracts(S_ClassDoc, sDepartmentRoot, sContractPartnerCode & sDZKCode, Request("UserFieldText8"), Request("UserFieldText3"))
		'S_DocID = GetNewDocIDForContracts(S_ClassDoc, sDepartmentRoot, sContractPartnerCode, Request("UserFieldText8"), Request("UserFieldText3"))
'ph - Запрос №47 - СТС}
		S_DocIDAdd = S_DocID
	End If
	If sDepartmentRoot <> SIT_STS or (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_RASP_DOCS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1) Then
		'реализация полужесткого маршрута
		S_ListToReconcile = Replace(S_ListToReconcile, Trim(SIT_AdditionalAgrees), "")
		'Для коммерческих предложений и заявок на закупку и оплату не нужна приписка про доп. согласующих  
		If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_COM_OFFERS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) <> 1 Then
			If S_ListToReconcile <> "" Then
				S_ListToReconcile = " " + SIT_AdditionalAgrees + DeleteConstPrefixFromList(S_ListToReconcile)
			End If
		End If
	End If

	'Формируем список согласования
	Select Case sDepartmentRoot
		Case SIT_STS
'{Запрос №50 - СТС
			'GetReconciliationByRules(parDocID, parClassDoc, parAmount, parChartOfAccounts, parCostCenter, parProjectCode, parBusinessUnit, parPartnerName, parKindOfPayment, parInitiator, parProjectManager, parOvertimeRequester, parIncomeExpenceContract, parContranctInSTS, parCurrency, parContractType, parOvertimeFuncLeaders, ByRef parListToReconcile, ByRef parNameAproval, ByRef parNameResponsible, ByRef parCorrespondent, ByRef parListToView, ByRef parRegistrar, ByRef parNameControl, ByRef parListToEdit)
			'GetReconciliationByRules iif(Trim(Request("DocID")) = "", S_DocID, Request("DocID")), Session("CurrentClassDoc"), CCur(CorrectSum(Request("DocAmountDoc"))), "", "", Request("UserFieldText8"), GetCodeFromCode_NameString(Request("BusinessUnit")), Request("DocPartnerName"), "", iif(UCase(Request("create")) = "Y", Session("UserID"), Request("DocAuthor")), "", "", Request("UserFieldText3"), Request("UserFieldText6"), Request("DocCurrency"), Request("ContractType"), NULL, S_ListToReconcile, S_NameAproval, Null, Null, Null, Null, Null, Null
			par_DocID = iif(Trim(Request("DocID")) = "", S_DocID, Request("DocID"))
			par_ClassDoc = Session("CurrentClassDoc")
			par_Amount = CCur(CorrectSum(Request("DocAmountDoc")))
			par_ChartOfAccounts = ""
			par_CostCenter = ""
			par_ProjectCode = Request("UserFieldText8")
			par_BusinessUnit = GetCodeFromCode_NameString(Request("BusinessUnit"))
			par_PartnerName = Request("DocPartnerName")
			par_KindOfPayment = ""
			par_Initiator = iif(UCase(Request("create")) = "Y", Session("UserID"), Request("DocAuthor"))
			par_ProjectManager = ""
			par_OvertimeRequester = ""
			par_IncomeExpenceContract = Request("UserFieldText3")
			par_ContranctInSTS = Request("UserFieldText6")
			par_Currency = Request("DocCurrency")
			par_ContractType = Request("ContractType")
			par_OvertimeFuncLeaders = NULL
			par_TypeOfDocument = ""
			par_FunctionArea = ""

			GetReconciliationByRules par_DocID, par_ClassDoc, par_Amount, par_ChartOfAccounts, par_CostCenter, par_ProjectCode, par_BusinessUnit, par_PartnerName, par_KindOfPayment, par_Initiator, par_ProjectManager, par_OvertimeRequester, par_IncomeExpenceContract, par_ContranctInSTS, par_Currency, par_ContractType, par_OvertimeFuncLeaders, par_TypeOfDocument, par_FunctionArea, _
				S_ListToReconcile, S_NameAproval, Null, Null, Null, Null, Null, Null
			'	par_ListToReconcile, par_NameAproval, par_NameResponsible, par_Correspondent, par_ListToView, par_Registrar, par_NameControl, par_ListToEdit
'Запрос №50 - СТС}
			'Запрос №38 - СТС - start
			If InStr(Request("DocDepartment"), SIT_STS_ROOT_DEPARTMENT) = 1 Then
				If not CheckUsersInListToReconcile(S_ListToReconcile, SIT_MaxUsersInListToReconcile) Then
					Session("Message") = AddNewLineToMessage(Session("Message"), Replace(SIT_MaxUsersInListToReconcileExceeded, "#MAX", CStr(SIT_MaxUsersInListToReconcile)))
					UserChangeDocCheck = False
					Exit Function
				End If
			End If
			'Запрос №38 - СТС - end
		Case SIT_SITRONICS
			S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
		Case Else 'Другие БН
			S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
	End Select

	'Подставляем название проекта
	Set dsTemp1 = Server.CreateObject("ADODB.Recordset")
	sSQL = "select * from ProjectList where ProjectID = "+sUnicodeSymbol+"'" + Request("UserFieldText8") + "'"
	dsTemp1.Open sSQL, Conn, 3, 1, &H1
	If not dsTemp1.EOF Then
		S_UserFieldText2 = Trim(CStr(dsTemp1("ProjectCode")))+" "+Trim(CStr(dsTemp1("ProjectName")))
		'S_UserFieldText2_Set = S_UserFieldText2
	Else 'Код проекта не соответствует справочнику
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInProjectNumber)
		UserChangeDocCheck = False
		dsTemp1.Close
		Exit Function
	End If
	dsTemp1.Close
  
	'Проверяем родительский док-т
	If Request("UserFieldText3") = STS_ContractPaymentDirection_Out_RU or Request("UserFieldText3") = STS_ContractPaymentDirection_Out_EN or Request("UserFieldText3") = STS_ContractPaymentDirection_Out_CZ Then
		If Trim(Request("DocIDParent")) = "" Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInParentOrderNumber)
			UserChangeDocCheck = False
			Exit Function
		Else
			'Проверка наличия родительского документа
			Set dsTemp1 = Server.CreateObject("ADODB.Recordset")
			sSQL = "select * from Docs where DocID = "+sUnicodeSymbol+"'" + Request("DocIDParent") + "'"
			AddLogD "@CheckParentDocID SQL: "+sSQL
			dsTemp1.Open sSQL, Conn, 3, 1, &H1
			If not dsTemp1.EOF Then
				'Родительский документ - не заявка на закупку
				If InStr(MyCStr(dsTemp1("ClassDoc")), STS_PurchaseOrder) <> 1 Then
					Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInParentOrderNumber)
					UserChangeDocCheck = False
					dsTemp1.Close
					Exit Function
				End If
				'Заявка отменена
				If dsTemp1("StatusCompletion") = "0" Then
					Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ParentOrderCanceled)
					UserChangeDocCheck = False
					dsTemp1.Close
					Exit Function
				End If
				'Заявка не утверждена
				If MyCStr(dsTemp1("NameApproved")) = "" Then
					Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ParentOrderNotApproved)
					UserChangeDocCheck = False
					dsTemp1.Close
					Exit Function
				End If
			Else
				Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInParentOrderNumber)
				UserChangeDocCheck = False
				dsTemp1.Close
				Exit Function
			End If
			dsTemp1.Close
		End If
	End If
	'sContractPartnerCode определяется выше при присвоении номера
	sContractPartnerCode = Trim(MyCStr(GetPartnerCode(Request("DocPartnerName"))))
	If sContractPartnerCode = "" Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorNoPartnerCode)
		UserChangeDocCheck = False
		Exit Function
	End If
	'Для всех категорий кроме заявок подставляем пользователей вместо ролей
	If (InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1) and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) <> 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_DOGOVORI_NEW)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIS)) <> 0) Then
		'Получаем список ролей
		If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) = 1 Then 'Для договоров получаем список ролей на всех языках (могут быть роли на разных языках)
			sRoleList = GetFullRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"), "", "", "")
		Else
			sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))
		End If
		'Подставляем роли в поля
		S_NameAproval = ReplaceRolesInList(S_NameAproval, sRoleList)
		S_NameResponsible = ReplaceRolesInList(S_NameResponsible, sRoleList)
		S_NameControl = ReplaceRolesInList(S_NameControl, sRoleList)
		S_ListToEdit = ReplaceRolesInList(S_ListToEdit, sRoleList)
		S_ListToView = ReplaceRolesInList(S_ListToView, sRoleList)
		S_ListToReconcile = ReplaceRolesInList(S_ListToReconcile, sRoleList)

		'Удаляем повторы в списке согласования
		S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)

		'vnik_error
		'Убираем автора документа из списка согласующих, т.к. после активации он не сможет загрузить новый документ
		'(работает пока только, если пользователь создает документ не под ролью)
		S_ListToReconcile = Replace(S_ListToReconcile,Request("DocAuthor"),"")
		'vnik_error
	End If
	'Проверка полей с пользователями
	'Для заявок отключаем проверку утверждающего, т.к. он рассчитывается
	If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 Then
		UserChangeDocCheck = CheckSingleUserField(S_NameAproval)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameAproval")+SIT_ErrorInUserField2)
			Exit Function
		End If
		'Для заявки на закупку отключаем проверку ответственного, т.к. можно указывать подразделение
		'Для заявки на оплату тоже, т.к. автоматом ставится роль из специального справочника
		UserChangeDocCheck = CheckSingleUserField(S_NameResponsible)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameResponsible")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
	UserChangeDocCheck = CheckSingleUserField(S_NameControl)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameControl")+SIT_ErrorInUserField2)
		Exit Function
	End If
	If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) = 0 Then 'В исходящих может быть любой текст
		UserChangeDocCheck = CheckMultiUserField(S_Correspondent)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("Correspondent")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
'Запрос №1 - СИБ - start
	IF InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_SLUZH_ZAPISKA)) = 0) THEN
'Запрос №1 - СИБ - end
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) = 0 Then 'В СЗ на переработки работают правила
			UserChangeDocCheck = CheckMultiUserField(S_ListToReconcile)
			If not UserChangeDocCheck Then
				Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToReconcile")+SIT_ErrorInUserField2)
				Exit Function
			End If
'Запрос №1 - СИБ - start
		End If
	END IF
'Запрос №1 - СИБ - end
	UserChangeDocCheck = CheckMultiUserField(S_ListToEdit)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToEdit")+SIT_ErrorInUserField2)
		Exit Function
	End If
	'UserChangeDocCheck = CheckMultiUserField(S_ListToView)
	'If not UserChangeDocCheck Then
	'	Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToView")+SIT_ErrorInUserField2)
	'	Exit Function
	'End If

'{ph - 20120812
	'S_AdditionalUsers = GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit"))
	S_AdditionalUsers = DeleteUserDoublesInListNew(Request("AdditionalUsers") & VbCrLf & GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit")), True)
'ph - 20120812}
	AddLogD "@@@Search For Leaders - S_AdditionalUsers: "+S_AdditionalUsers
	'vnik_payment_order
	'If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_PAYMENT_ORDER)) > 0 Then
	'	S_AdditionalUsers = S_AdditionalUsers + ReplaceRoleFromDir(SIT_Account_manager,sDepartmentRoot)       
	'End If        
	'vnik_payment_order
	S_AdditionalUsers = DeleteUserDoublesInList(S_AdditionalUsers)

	'vnik
	If S_AdditionalUsers = "" Then
		S_AdditionalUsers = "-"
	End If
	S_AdditionalUsers_Set = S_AdditionalUsers

	'Добавляем в дополнительных пользователей финансового контролера из справочника ролей для заявок
	S_AdditionalUsers = GetRoleForOrders(STS_Orders_FinancialControl, "", "", GetCodeFromCode_NameString(Request("BusinessUnit")))
	If S_AdditionalUsers_Set = "-" Then
		S_AdditionalUsers_Set = S_AdditionalUsers
	Else
		S_AdditionalUsers_Set = DeleteUserDoublesInList(S_AdditionalUsers_Set & VbCrLf & S_AdditionalUsers)
	End If
	S_AdditionalUsers = S_AdditionalUsers_Set

	S_ListToReconcile_Set = S_ListToReconcile
	S_NameAproval_Set = S_NameAproval
	S_ListToEdit = S_ListToEdit
	S_NameResponsible_Set = S_NameResponsible
	S_NameControl_Set = S_NameControl
	S_ListToView_Set=S_ListToView

	If UCase(Request("UpdateDoc")) = "YES" And (Request("create") = "y"  Or Request("DocIDPrev") <> Request("DocID")) Then
		'out "123"
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		dsTemp.Open "select * from Docs where DocID='" + MakeSQLSafe(Request("DocID")) + "'", Conn 
		If Not dsTemp.EOF Then
			Session("Message") = DOCS_ALREADYEXISTS + ": " + Request("DocID")
			dsTemp.Close
			UserChangeDocCheck=False
			Exit Function
		End If
		dsTemp.Close
	End If

	'ph - 20090808 - start
	If Request("DocIDParent") <> "" Then
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		bMyDocIDParentChanged = False
		If Request("create") <> "y" Then
			'Проверяем изменился ли номер родительского документа, если идет редактирование
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocID"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				bMyDocIDParentChanged = Request("DocIDParent")<>dsTemp("DocIDParent")
			End If
			dsTemp.Close
		End If
		If bMyDocIDParentChanged Then
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocIDParent"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				'sMessage = ""
				'UserChangeDocCheck = oPayDox.IsReadAccessUser(Session("UserID"), sMessage, dsTemp, DOCS_All, DOCS_NoReadAccess, DOCS_NoAccess, USER_Department, VAR_ExtInt, VAR_AdminSecLevel, VAR_StatusActiveUser, NULL)
				UserChangeDocCheck = IsReadAccessRS(dsTemp)' IsReadAccessUser(Session("UserID"), dsTemp, NULL)
				If UserChangeDocCheck Then
					S_DocIDParent = dsTemp("DocID")
				End If
				dsTemp.Close
				If not UserChangeDocCheck Then
					Session("Message") = SIT_NoAccessToParentDoc'sMessage
					Exit Function
				End If
			End If
		End If
	End If
	'ph - 20090808 - end

' *** КОММЕРЧЕСКИЕ ПРЕДЛОЖЕНИЯ
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_COM_OFFERS)) = 1 Then
	'Создатель документа
	If UCase(Request("create")) = "Y" Then
		sDocCreator = Session("UserID")
	Else
		sDocCreator = GetUserID(Request("DocAuthor"))
	End If

	S_NameAproval=Request("DocNameAproval")
	S_NameResponsible=Request("DocNameResponsible")
	S_NameControl=Request("DocNameControl")
	S_ListToReconcile=Request("DocListToReconcile")
	S_ListToEdit=Request("DocListToEdit")
	S_ListToView=Request("DocListToView")
	S_Correspondent=Request("DocCorrespondent")

	'Определяем бизнес направление
	sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))

	' SAY 2008-07-21
	Set dsTempPR = Server.CreateObject("ADODB.Recordset")

' DmGorsky_5: меняем местами проверку: вначале ищем SIT_SITRU ("СИТРОНИКС ИТ"), потом, если не найдено, - SIT_SITRONICS ("СИТРОНИКС")
	If InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRU))=1 then ' DmGorsky
		sDepartment = SIT_SITRU ' DmGorsky
		sDepartmentCode = "" ' DmGorsky
'Запрос №1 - СИБ - start
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SIB_ROOT_DEPARTMENT))=1 then
		sDepartment = SIT_SIB
		sDepartmentCode = ""
'Запрос №1 - СИБ - end
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRONICS))=1 then
		sDepartment = SIT_SITRONICS
		sDepartmentCode = ""
	Else
		sDepartment = SIT_STS

		sSQL = "select code from Departments where Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"' or Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"/'"
		AddLogD "WAY sSQL="+sSQL
		AddLogD "WAY Conn="+Conn
		dsTempPR.Open sSQL, Conn, 3, 1, &H1

		if not dsTempPR.EOF Then
			sDepartmentCode = dsTempPR("code")
			if InStrRev(sDepartmentCode, "/") > 0 then
				sDepartmentCode = Right(sDepartmentCode, Len(sDepartmentCode)-InStrRev(sDepartmentCode, "/"))
			End If
		Else
			sDepartmentCode="000000"
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInDepartmentCode1 + DelOtherLangFromFolder(Request("DocDepartment")) + SIT_ErrorInDepartmentCode2)
			UserChangeDocCheck=False
			Exit Function
		End If

		dsTempPR.Close
	End If

	If Request("create") = "y" Then
		sPostfix = ""
		sSearchCol = "DocID"
		sPrePrefix = ""
		'SAY 2009-02-20
		sSufix = ""

		'SAY 2009-03-12
		S_DocID = ""

		S_DocID = GetNewDocIDForComOffers(Session("CurrentClassDoc"), sDepartmentRoot)
		S_DocIDAdd = S_DocID
	End If
	If sDepartmentRoot <> SIT_STS or (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_RASP_DOCS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1) Then
		'реализация полужесткого маршрута
		S_ListToReconcile = Replace(S_ListToReconcile, Trim(SIT_AdditionalAgrees), "")
		'Для коммерческих предложений и заявок на закупку и оплату не нужна приписка про доп. согласующих  
		If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_COM_OFFERS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) <> 1 Then
			If S_ListToReconcile <> "" Then
				S_ListToReconcile = " " + SIT_AdditionalAgrees + DeleteConstPrefixFromList(S_ListToReconcile)
			End If
		End If
	End If
	'Для всех категорий кроме заявок подставляем пользователей вместо ролей
	If (InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1) and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) <> 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_DOGOVORI_NEW)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIS)) <> 0) Then
		'Получаем список ролей
		If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) = 1 Then 'Для договоров получаем список ролей на всех языках (могут быть роли на разных языках)
			sRoleList = GetFullRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"), "", "", "")
		Else
			sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))
		End If
		'Подставляем роли в поля
		S_NameAproval = ReplaceRolesInList(S_NameAproval, sRoleList)
		S_NameResponsible = ReplaceRolesInList(S_NameResponsible, sRoleList)
		S_NameControl = ReplaceRolesInList(S_NameControl, sRoleList)
		S_ListToEdit = ReplaceRolesInList(S_ListToEdit, sRoleList)
		S_ListToView = ReplaceRolesInList(S_ListToView, sRoleList)
		S_ListToReconcile = ReplaceRolesInList(S_ListToReconcile, sRoleList)

		'Удаляем повторы в списке согласования
		S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)

		'vnik_error
		'Убираем автора документа из списка согласующих, т.к. после активации он не сможет загрузить новый документ
		'(работает пока только, если пользователь создает документ не под ролью)
		S_ListToReconcile = Replace(S_ListToReconcile,Request("DocAuthor"),"")
		'vnik_error
	End If
	'Проверка полей с пользователями
	'Для заявок отключаем проверку утверждающего, т.к. он рассчитывается
	If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 Then
		UserChangeDocCheck = CheckSingleUserField(S_NameAproval)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameAproval")+SIT_ErrorInUserField2)
			Exit Function
		End If
		'Для заявки на закупку отключаем проверку ответственного, т.к. можно указывать подразделение
		'Для заявки на оплату тоже, т.к. автоматом ставится роль из специального справочника
		UserChangeDocCheck = CheckSingleUserField(S_NameResponsible)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameResponsible")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
	UserChangeDocCheck = CheckSingleUserField(S_NameControl)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameControl")+SIT_ErrorInUserField2)
		Exit Function
	End If
	If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) = 0 Then 'В исходящих может быть любой текст
		UserChangeDocCheck = CheckMultiUserField(S_Correspondent)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("Correspondent")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
'Запрос №1 - СИБ - start
	IF InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_SLUZH_ZAPISKA)) = 0) THEN
'Запрос №1 - СИБ - end
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) = 0 Then 'В СЗ на переработки работают правила
			UserChangeDocCheck = CheckMultiUserField(S_ListToReconcile)
			If not UserChangeDocCheck Then
				Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToReconcile")+SIT_ErrorInUserField2)
				Exit Function
			End If
'Запрос №1 - СИБ - start
		End If
	END IF
'Запрос №1 - СИБ - end
	UserChangeDocCheck = CheckMultiUserField(S_ListToEdit)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToEdit")+SIT_ErrorInUserField2)
		Exit Function
	End If
	'UserChangeDocCheck = CheckMultiUserField(S_ListToView)
	'If not UserChangeDocCheck Then
	'	Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToView")+SIT_ErrorInUserField2)
	'	Exit Function
	'End If

'{ph - 20120812
	'S_AdditionalUsers = GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit"))
	S_AdditionalUsers = DeleteUserDoublesInListNew(Request("AdditionalUsers") & VbCrLf & GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit")), True)
'ph - 20120812}
	AddLogD "@@@Search For Leaders - S_AdditionalUsers: "+S_AdditionalUsers
	'vnik_payment_order
	'If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_PAYMENT_ORDER)) > 0 Then
	'	S_AdditionalUsers = S_AdditionalUsers + ReplaceRoleFromDir(SIT_Account_manager,sDepartmentRoot)       
	'End If        
	'vnik_payment_order
	S_AdditionalUsers = DeleteUserDoublesInList(S_AdditionalUsers)

	'vnik
	If S_AdditionalUsers = "" Then
		S_AdditionalUsers = "-"
	End If
	S_AdditionalUsers_Set = S_AdditionalUsers

	S_ListToReconcile_Set = S_ListToReconcile
	S_NameAproval_Set = S_NameAproval
	S_ListToEdit = S_ListToEdit
	S_NameResponsible_Set = S_NameResponsible
	S_NameControl_Set = S_NameControl
	S_ListToView_Set=S_ListToView

	If UCase(Request("UpdateDoc")) = "YES" And (Request("create") = "y"  Or Request("DocIDPrev") <> Request("DocID")) Then
		'out "123"
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		dsTemp.Open "select * from Docs where DocID='" + MakeSQLSafe(Request("DocID")) + "'", Conn 
		If Not dsTemp.EOF Then
			Session("Message") = DOCS_ALREADYEXISTS + ": " + Request("DocID")
			dsTemp.Close
			UserChangeDocCheck=False
			Exit Function
		End If
		dsTemp.Close
	End If

	'ph - 20090808 - start
	If Request("DocIDParent") <> "" Then
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		bMyDocIDParentChanged = False
		If Request("create") <> "y" Then
			'Проверяем изменился ли номер родительского документа, если идет редактирование
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocID"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				bMyDocIDParentChanged = Request("DocIDParent")<>dsTemp("DocIDParent")
			End If
			dsTemp.Close
		End If
		If bMyDocIDParentChanged Then
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocIDParent"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				'sMessage = ""
				'UserChangeDocCheck = oPayDox.IsReadAccessUser(Session("UserID"), sMessage, dsTemp, DOCS_All, DOCS_NoReadAccess, DOCS_NoAccess, USER_Department, VAR_ExtInt, VAR_AdminSecLevel, VAR_StatusActiveUser, NULL)
				UserChangeDocCheck = IsReadAccessRS(dsTemp)' IsReadAccessUser(Session("UserID"), dsTemp, NULL)
				If UserChangeDocCheck Then
					S_DocIDParent = dsTemp("DocID")
				End If
				dsTemp.Close
				If not UserChangeDocCheck Then
					Session("Message") = SIT_NoAccessToParentDoc'sMessage
					Exit Function
				End If
			End If
		End If
	End If
	'ph - 20090808 - end

' *** HELPDESK
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_HelpDesk)) = 1 Then 
	'Создатель документа
	If UCase(Request("create")) = "Y" Then
		sDocCreator = Session("UserID")
	Else
		sDocCreator = GetUserID(Request("DocAuthor"))
	End If

	S_NameAproval=Request("DocNameAproval")
	S_NameResponsible=Request("DocNameResponsible")
	S_NameControl=Request("DocNameControl")
	S_ListToReconcile=Request("DocListToReconcile")
	S_ListToEdit=Request("DocListToEdit")
	S_ListToView=Request("DocListToView")
	S_Correspondent=Request("DocCorrespondent")

	'Определяем бизнес направление
	sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))

	' SAY 2008-07-21
	Set dsTempPR = Server.CreateObject("ADODB.Recordset")

' DmGorsky_5: меняем местами проверку: вначале ищем SIT_SITRU ("СИТРОНИКС ИТ"), потом, если не найдено, - SIT_SITRONICS ("СИТРОНИКС")
	If InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRU))=1 then ' DmGorsky
		sDepartment = SIT_SITRU ' DmGorsky
		sDepartmentCode = "" ' DmGorsky
'Запрос №1 - СИБ - start
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SIB_ROOT_DEPARTMENT))=1 then
		sDepartment = SIT_SIB
		sDepartmentCode = ""
'Запрос №1 - СИБ - end
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRONICS))=1 then
		sDepartment = SIT_SITRONICS
		sDepartmentCode = ""
	Else
		sDepartment = SIT_STS

		sSQL = "select code from Departments where Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"' or Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"/'"
		AddLogD "WAY sSQL="+sSQL
		AddLogD "WAY Conn="+Conn
		dsTempPR.Open sSQL, Conn, 3, 1, &H1

		if not dsTempPR.EOF Then
			sDepartmentCode = dsTempPR("code")
			if InStrRev(sDepartmentCode, "/") > 0 then
				sDepartmentCode = Right(sDepartmentCode, Len(sDepartmentCode)-InStrRev(sDepartmentCode, "/"))
			End If
		Else
			sDepartmentCode="000000"
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInDepartmentCode1 + DelOtherLangFromFolder(Request("DocDepartment")) + SIT_ErrorInDepartmentCode2)
			UserChangeDocCheck=False
			Exit Function
		End If

		dsTempPR.Close
	End If

	If Request("create") = "y" Then
		sPostfix = ""
		sSearchCol = "DocID"
		sPrePrefix = ""
		'SAY 2009-02-20
		sSufix = ""

		'SAY 2009-03-12
		S_DocID = ""

		Call GetNewDocID_test(Session("CurrentClassDoc"), sDepartmentRoot, "", "",  "", "")
		S_DocIDAdd = S_DocID
	End If
	If sDepartmentRoot <> SIT_STS or (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_RASP_DOCS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1) Then
		'реализация полужесткого маршрута
		S_ListToReconcile = Replace(S_ListToReconcile, Trim(SIT_AdditionalAgrees), "")
		'Для коммерческих предложений и заявок на закупку и оплату не нужна приписка про доп. согласующих  
		If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_COM_OFFERS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) <> 1 Then
			If S_ListToReconcile <> "" Then
				S_ListToReconcile = " " + SIT_AdditionalAgrees + DeleteConstPrefixFromList(S_ListToReconcile)
			End If
		End If
	End If
	'Для всех категорий кроме заявок подставляем пользователей вместо ролей
	If (InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1) and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) <> 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_DOGOVORI_NEW)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIS)) <> 0) Then
		'Получаем список ролей
		If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) = 1 Then 'Для договоров получаем список ролей на всех языках (могут быть роли на разных языках)
			sRoleList = GetFullRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"), "", "", "")
		Else
			sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))
		End If
		'Подставляем роли в поля
		S_NameAproval = ReplaceRolesInList(S_NameAproval, sRoleList)
		S_NameResponsible = ReplaceRolesInList(S_NameResponsible, sRoleList)
		S_NameControl = ReplaceRolesInList(S_NameControl, sRoleList)
		S_ListToEdit = ReplaceRolesInList(S_ListToEdit, sRoleList)
		S_ListToView = ReplaceRolesInList(S_ListToView, sRoleList)
		S_ListToReconcile = ReplaceRolesInList(S_ListToReconcile, sRoleList)

		'Удаляем повторы в списке согласования
		S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)

		'vnik_error
		'Убираем автора документа из списка согласующих, т.к. после активации он не сможет загрузить новый документ
		'(работает пока только, если пользователь создает документ не под ролью)
		S_ListToReconcile = Replace(S_ListToReconcile,Request("DocAuthor"),"")
		'vnik_error
	End If
	'Проверка полей с пользователями
	'Для заявок отключаем проверку утверждающего, т.к. он рассчитывается
	If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 Then
		UserChangeDocCheck = CheckSingleUserField(S_NameAproval)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameAproval")+SIT_ErrorInUserField2)
			Exit Function
		End If
		'Для заявки на закупку отключаем проверку ответственного, т.к. можно указывать подразделение
		'Для заявки на оплату тоже, т.к. автоматом ставится роль из специального справочника
		UserChangeDocCheck = CheckSingleUserField(S_NameResponsible)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameResponsible")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
	UserChangeDocCheck = CheckSingleUserField(S_NameControl)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameControl")+SIT_ErrorInUserField2)
		Exit Function
	End If
	If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) = 0 Then 'В исходящих может быть любой текст
		UserChangeDocCheck = CheckMultiUserField(S_Correspondent)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("Correspondent")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
'Запрос №1 - СИБ - start
	IF InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_SLUZH_ZAPISKA)) = 0) THEN
'Запрос №1 - СИБ - end
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) = 0 Then 'В СЗ на переработки работают правила
			UserChangeDocCheck = CheckMultiUserField(S_ListToReconcile)
			If not UserChangeDocCheck Then
				Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToReconcile")+SIT_ErrorInUserField2)
				Exit Function
			End If
'Запрос №1 - СИБ - start
		End If
	END IF
'Запрос №1 - СИБ - end
	UserChangeDocCheck = CheckMultiUserField(S_ListToEdit)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToEdit")+SIT_ErrorInUserField2)
		Exit Function
	End If
	'UserChangeDocCheck = CheckMultiUserField(S_ListToView)
	'If not UserChangeDocCheck Then
	'	Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToView")+SIT_ErrorInUserField2)
	'	Exit Function
	'End If

'{ph - 20120812
	'S_AdditionalUsers = GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit"))
	S_AdditionalUsers = DeleteUserDoublesInListNew(Request("AdditionalUsers") & VbCrLf & GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit")), True)
'ph - 20120812}
	AddLogD "@@@Search For Leaders - S_AdditionalUsers: "+S_AdditionalUsers
	'vnik_payment_order
	'If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_PAYMENT_ORDER)) > 0 Then
	'	S_AdditionalUsers = S_AdditionalUsers + ReplaceRoleFromDir(SIT_Account_manager,sDepartmentRoot)       
	'End If        
	'vnik_payment_order
	S_AdditionalUsers = DeleteUserDoublesInList(S_AdditionalUsers)

	'vnik
	If S_AdditionalUsers = "" Then
		S_AdditionalUsers = "-"
	End If
	S_AdditionalUsers_Set = S_AdditionalUsers

	S_ListToReconcile_Set = S_ListToReconcile
	S_NameAproval_Set = S_NameAproval
	S_ListToEdit = S_ListToEdit
	S_NameResponsible_Set = S_NameResponsible
	S_NameControl_Set = S_NameControl
	S_ListToView_Set=S_ListToView

	If UCase(Request("UpdateDoc")) = "YES" And (Request("create") = "y"  Or Request("DocIDPrev") <> Request("DocID")) Then
		'out "123"
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		dsTemp.Open "select * from Docs where DocID='" + MakeSQLSafe(Request("DocID")) + "'", Conn 
		If Not dsTemp.EOF Then
			Session("Message") = DOCS_ALREADYEXISTS + ": " + Request("DocID")
			dsTemp.Close
			UserChangeDocCheck=False
			Exit Function
		End If
		dsTemp.Close
	End If

	'ph - 20090808 - start
	If Request("DocIDParent") <> "" Then
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		bMyDocIDParentChanged = False
		If Request("create") <> "y" Then
			'Проверяем изменился ли номер родительского документа, если идет редактирование
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocID"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				bMyDocIDParentChanged = Request("DocIDParent")<>dsTemp("DocIDParent")
			End If
			dsTemp.Close
		End If
		If bMyDocIDParentChanged Then
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocIDParent"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				'sMessage = ""
				'UserChangeDocCheck = oPayDox.IsReadAccessUser(Session("UserID"), sMessage, dsTemp, DOCS_All, DOCS_NoReadAccess, DOCS_NoAccess, USER_Department, VAR_ExtInt, VAR_AdminSecLevel, VAR_StatusActiveUser, NULL)
				UserChangeDocCheck = IsReadAccessRS(dsTemp)' IsReadAccessUser(Session("UserID"), dsTemp, NULL)
				If UserChangeDocCheck Then
					S_DocIDParent = dsTemp("DocID")
				End If
				dsTemp.Close
				If not UserChangeDocCheck Then
					Session("Message") = SIT_NoAccessToParentDoc'sMessage
					Exit Function
				End If
			End If
		End If
	End If
	'ph - 20090808 - end

' *** ЗАЯВКИ НА ЗАКУПКУ СТС
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) = 1 Then
	'Создатель документа
	If UCase(Request("create")) = "Y" Then
		sDocCreator = Session("UserID")
	Else
		sDocCreator = GetUserID(Request("DocAuthor"))
	End If

	S_NameAproval=Request("DocNameAproval")
	S_NameResponsible=Request("DocNameResponsible")
	S_NameControl=Request("DocNameControl")
	S_ListToReconcile=Request("DocListToReconcile")
	S_ListToEdit=Request("DocListToEdit")
	S_ListToView=Request("DocListToView")
	S_Correspondent=Request("DocCorrespondent")

	'Определяем бизнес направление
	sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))

	' SAY 2008-07-21
	Set dsTempPR = Server.CreateObject("ADODB.Recordset")

' DmGorsky_5: меняем местами проверку: вначале ищем SIT_SITRU ("СИТРОНИКС ИТ"), потом, если не найдено, - SIT_SITRONICS ("СИТРОНИКС")
	If InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRU))=1 then ' DmGorsky
		sDepartment = SIT_SITRU ' DmGorsky
		sDepartmentCode = "" ' DmGorsky
'Запрос №1 - СИБ - start
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SIB_ROOT_DEPARTMENT))=1 then
		sDepartment = SIT_SIB
		sDepartmentCode = ""
'Запрос №1 - СИБ - end
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRONICS))=1 then
		sDepartment = SIT_SITRONICS
		sDepartmentCode = ""
	Else
		sDepartment = SIT_STS

		sSQL = "select code from Departments where Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"' or Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"/'"
		AddLogD "WAY sSQL="+sSQL
		AddLogD "WAY Conn="+Conn
		dsTempPR.Open sSQL, Conn, 3, 1, &H1

		if not dsTempPR.EOF Then
			sDepartmentCode = dsTempPR("code")
			if InStrRev(sDepartmentCode, "/") > 0 then
				sDepartmentCode = Right(sDepartmentCode, Len(sDepartmentCode)-InStrRev(sDepartmentCode, "/"))
			End If
		Else
			sDepartmentCode="000000"
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInDepartmentCode1 + DelOtherLangFromFolder(Request("DocDepartment")) + SIT_ErrorInDepartmentCode2)
			UserChangeDocCheck=False
			Exit Function
		End If

		dsTempPR.Close
	End If

	If Request("create") = "y" Then
		sPostfix = ""
		sSearchCol = "DocID"
		sPrePrefix = ""
		'SAY 2009-02-20
		sSufix = ""

		'SAY 2009-03-12
		S_DocID = ""

		S_DocID = GetNewOrderDocID(Session("CurrentClassDoc"))
		S_DocIDAdd = S_DocID
	End If
	If sDepartmentRoot <> SIT_STS or (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_RASP_DOCS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1) Then
		'реализация полужесткого маршрута
		S_ListToReconcile = Replace(S_ListToReconcile, Trim(SIT_AdditionalAgrees), "")
		'Для коммерческих предложений и заявок на закупку и оплату не нужна приписка про доп. согласующих  
		If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_COM_OFFERS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) <> 1 Then
			If S_ListToReconcile <> "" Then
				S_ListToReconcile = " " + SIT_AdditionalAgrees + DeleteConstPrefixFromList(S_ListToReconcile)
			End If
		End If
	End If
	'Для всех категорий кроме заявок подставляем пользователей вместо ролей
	If (InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1) and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) <> 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_DOGOVORI_NEW)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIS)) <> 0) Then
		'Получаем список ролей
		If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) = 1 Then 'Для договоров получаем список ролей на всех языках (могут быть роли на разных языках)
			sRoleList = GetFullRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"), "", "", "")
		Else
			sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))
		End If
		'Подставляем роли в поля
		S_NameAproval = ReplaceRolesInList(S_NameAproval, sRoleList)
		S_NameResponsible = ReplaceRolesInList(S_NameResponsible, sRoleList)
		S_NameControl = ReplaceRolesInList(S_NameControl, sRoleList)
		S_ListToEdit = ReplaceRolesInList(S_ListToEdit, sRoleList)
		S_ListToView = ReplaceRolesInList(S_ListToView, sRoleList)
		S_ListToReconcile = ReplaceRolesInList(S_ListToReconcile, sRoleList)

		'Удаляем повторы в списке согласования
		S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)

		'vnik_error
		'Убираем автора документа из списка согласующих, т.к. после активации он не сможет загрузить новый документ
		'(работает пока только, если пользователь создает документ не под ролью)
		S_ListToReconcile = Replace(S_ListToReconcile,Request("DocAuthor"),"")
		'vnik_error
	End If
	'Проверка полей с пользователями
	'Для заявок отключаем проверку утверждающего, т.к. он рассчитывается
	If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 Then
		UserChangeDocCheck = CheckSingleUserField(S_NameAproval)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameAproval")+SIT_ErrorInUserField2)
			Exit Function
		End If
		'Для заявки на закупку отключаем проверку ответственного, т.к. можно указывать подразделение
		'Для заявки на оплату тоже, т.к. автоматом ставится роль из специального справочника
		UserChangeDocCheck = CheckSingleUserField(S_NameResponsible)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameResponsible")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
	UserChangeDocCheck = CheckSingleUserField(S_NameControl)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameControl")+SIT_ErrorInUserField2)
		Exit Function
	End If
	If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) = 0 Then 'В исходящих может быть любой текст
		UserChangeDocCheck = CheckMultiUserField(S_Correspondent)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("Correspondent")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
'Запрос №1 - СИБ - start
	IF InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_SLUZH_ZAPISKA)) = 0) THEN
'Запрос №1 - СИБ - end
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) = 0 Then 'В СЗ на переработки работают правила
			UserChangeDocCheck = CheckMultiUserField(S_ListToReconcile)
			If not UserChangeDocCheck Then
				Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToReconcile")+SIT_ErrorInUserField2)
				Exit Function
			End If
'Запрос №1 - СИБ - start
		End If
	END IF
'Запрос №1 - СИБ - end
	UserChangeDocCheck = CheckMultiUserField(S_ListToEdit)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToEdit")+SIT_ErrorInUserField2)
		Exit Function
	End If
	'UserChangeDocCheck = CheckMultiUserField(S_ListToView)
	'If not UserChangeDocCheck Then
	'	Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToView")+SIT_ErrorInUserField2)
	'	Exit Function
	'End If

'{ph - 20120812
	'S_AdditionalUsers = GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit"))
	S_AdditionalUsers = DeleteUserDoublesInListNew(Request("AdditionalUsers") & VbCrLf & GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit")), True)
'ph - 20120812}
	AddLogD "@@@Search For Leaders - S_AdditionalUsers: "+S_AdditionalUsers
	'vnik_payment_order
	'If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_PAYMENT_ORDER)) > 0 Then
	'	S_AdditionalUsers = S_AdditionalUsers + ReplaceRoleFromDir(SIT_Account_manager,sDepartmentRoot)       
	'End If        
	'vnik_payment_order
	S_AdditionalUsers = DeleteUserDoublesInList(S_AdditionalUsers)

	'vnik
	If S_AdditionalUsers = "" Then
		S_AdditionalUsers = "-"
	End If
	S_AdditionalUsers_Set = S_AdditionalUsers

	S_ListToReconcile_Set = S_ListToReconcile
	S_NameAproval_Set = S_NameAproval
	S_ListToEdit = S_ListToEdit
	S_NameResponsible_Set = S_NameResponsible
	S_NameControl_Set = S_NameControl
	S_ListToView_Set=S_ListToView

	If UCase(Request("UpdateDoc")) = "YES" And (Request("create") = "y"  Or Request("DocIDPrev") <> Request("DocID")) Then
		'out "123"
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		dsTemp.Open "select * from Docs where DocID='" + MakeSQLSafe(Request("DocID")) + "'", Conn 
		If Not dsTemp.EOF Then
			Session("Message") = DOCS_ALREADYEXISTS + ": " + Request("DocID")
			dsTemp.Close
			UserChangeDocCheck=False
			Exit Function
		End If
		dsTemp.Close
	End If

	'Проверка даты исполнения заявок
	If Trim(Request("DocDateCompletion")) <> "" Then
		If ConvertToDate(Request("DocDateCompletion")) <> VAR_BeginOfTimes Then
			If ConvertToDate(Request("DocDateCompletion")) < Date Then
				Session("Message") = SIT_TooEarlyDate+Request("DocDateCompletion")
				UserChangeDocCheck = False
				Exit Function
			End If
		Else
			Session("Message") = SIT_DateFormatError+Request("DocDateCompletion")
			UserChangeDocCheck = False
			Exit Function
		End If
	End If

	'Запрос №36 - СТС - start
	'Получаем лимиты из БД
	'STS_HeadOfSector_Limit = 0
	'STS_HeadOfDepartment_Limit = 0
	'STS_HeadOfDivision_Limit = 0
	'STS_FinancialControl_Limit = 0
	'STS_FinDirector_Limit = 0
	'STS_GenDirector_Limit = 0
	'STS_ProjectManager_Limit = 0
	'STS_Accounting_Limit = 0
	'GetLimitsForOrders STS_HeadOfSector_Limit, STS_HeadOfDepartment_Limit, STS_HeadOfDivision_Limit, STS_FinancialControl_Limit, STS_FinDirector_Limit, STS_GenDirector_Limit, STS_ProjectManager_Limit, STS_Accounting_Limit
	'Запрос №36 - СТС - end

	'Подставляем название проекта и определям менеджера проекта
	Set dsTemp1 = Server.CreateObject("ADODB.Recordset")
	sSQL = "select * from ProjectList where ProjectID = "+sUnicodeSymbol+"'" + Request("UserFieldText3") + "'"
	AddLogD "@@@GetProjectName SQL: "+sSQL
	dsTemp1.Open sSQL, Conn, 3, 1, &H1
	If Not dsTemp1.EOF Then
		S_UserFieldText4 = Trim(CStr(dsTemp1("ProjectCode")))+" "+Trim(CStr(dsTemp1("ProjectName")))
		S_UserFieldText4_Set = S_UserFieldText4
		If IsNull(dsTemp1("ProjectManagerUser")) Then
			ProjectManager = ""
		Else
			ProjectManager = Trim(CStr(dsTemp1("ProjectManagerUser")))
		End If
	Else 'Код проекта не соответствует справочнику
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInProjectNumber)
		UserChangeDocCheck = False
		dsTemp1.Close
		Exit Function
	End If
	dsTemp1.Close

	'Проверка, что бизнес единица соответствует справочнику и замена текущего значения на англ. вариант
	sSQL = BusinessUnitSelectForCheckValue(Request("BusinessUnit"))
	AddLogD "@@@CheckBusinessUnit SQL: "+sSQL
	dsTemp1.Open sSQL, Conn, 3, 1, &H1
	If dsTemp1.EOF Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInBU)
		UserChangeDocCheck = False
		dsTemp1.Close
		Exit Function
	End If
	S_AddField2 = dsTemp1("BusinessUnit")&" - "&dsTemp1("Company_EN")
	dsTemp1.Close

	'Проверка, что Статья расходов соответствует справочнику и замена текущего значения на англ. вариант
	sSQL = ChartOfAccountsSelectForCheckValue(Request("UserFieldText8"))
	AddLogD "@@@CheckChartOfAccount SQL: "+sSQL
	dsTemp1.Open sSQL, Conn, 3, 1, &H1
	If dsTemp1.EOF Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInChartOfAccount)
		UserChangeDocCheck = False
		dsTemp1.Close
		Exit Function
	End If
	S_UserFieldText8 = dsTemp1("STS_Account_No")&" - "&dsTemp1("AccountName_EN")
	dsTemp1.Close

	'Проверка, что Форма расчета соответствует справочнику и замена текущего значения на англ. вариант
	sSQL = PaymentTypesSelectForCheckValue(Request("UserFieldText5"))
	AddLogD "@@@CheckPaymentType SQL: "+sSQL
	dsTemp1.Open sSQL, Conn, 3, 1, &H1
	If dsTemp1.EOF Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInPaymentType)
		UserChangeDocCheck = False
		dsTemp1.Close
		Exit Function
	End If
	S_UserFieldText5 = dsTemp1("PaymentType_EN")
	dsTemp1.Close

	'Проверка, что Центр затрат соответствует справочнику и замена текущего значения на англ. вариант
	sSQL =  CostCentersSelectForCheckValue(Trim(Request("UserFieldText1")))
	AddLogD "@@@CheckCostCenter SQL: "+sSQL
	dsTemp1.Open sSQL, Conn, 3, 1, &H1
	If dsTemp1.EOF Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInCostCenter)
		UserChangeDocCheck = False
		dsTemp1.Close
		Exit Function
	End If
	S_UserFieldText1 = dsTemp1("CostCenterEng")
	dsTemp1.Close

	'Запрос №43 - СТС - start
	'Вообще убрали - 20110728
	''Проверка, что Контрагент проверенный
	'sSQL =  "select Name from Partners where Name = N'" & MakeSQLSafeSimple((Trim(Request("DocPartnerName")))) & "' and IsNull(Rating, N'') = N'+'"
	'AddLogD "@@@CheckPartnerName SQL: "+sSQL
	'dsTemp1.Open sSQL, Conn, 3, 1, &H1
	'If dsTemp1.EOF Then
	'	Session("Message") = AddNewLineToMessage(Session("Message"), STS_ErrorPartnerIsNotChecked)
	'	'Переиграли - ругаемся, но сохранять даем
	'	UserChangeDocCheck = False
	'	dsTemp1.Close
	'	Exit Function
	'End If
	'dsTemp1.Close
	'Запрос №43 - СТС - end

	AddLogD "@@@Request(""DocAmountDoc""): "+Request("DocAmountDoc")
	ErrorInAmount = True
	If Request("DocAmountDoc") <> "" Then
		If IsNumeric(Request("DocAmountDoc")) Then
			If Request("DocCurrency") = "USD" Then
				USD_Amount = CCur(CorrectSum(Request("DocAmountDoc")))
			Else
				cCurrencyConvertionFactor = CurrencyConvertionFactor(Request("DocCurrency"), "USD")
				AddLogD "@@@cCurrencyConvertionFactor: "+CStr(cCurrencyConvertionFactor)
				USD_Amount = CCur(CorrectSum(Request("DocAmountDoc")))*cCurrencyConvertionFactor
			End If
			ErrorInAmount = USD_Amount <= 0
		End If
	End If
	AddlogD "@@@USD_Amount: "+CStr(USD_Amount)
	If ErrorInAmount Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInSumOrCurrency)
		UserChangeDocCheck = False
		Exit Function
	Else
		S_UserFieldMoney1 = CStr(USD_Amount)
		'Запрос №43 - СТС - start
		'Проверка возможности создания заявки на указанную сумму
		If USD_Amount > 10000 Then
			'Проверяем руководителей подразделений
			sSQL = "select Leader, IsNull(Statuses, '') as Statuses from DepartmentDependants left join Departments on (Departments.GUID = DepartmentDependants.DependantGUID) where CharIndex(N'<" & Session("UserID") & ">', Leader) > 0"
			sSQL = sSQL & "union select Name as Leader, N'#LEV2' as Statuses from STS_DirectorOfDirectionRU where CharIndex(N'<" & Session("UserID") & ">', Name) > 0"
			sSQL = sSQL & "union select Name as Leader, N'#LEV2' as Statuses from STS_AssistantDirectorRU where CharIndex(N'<" & Session("UserID") & ">', Name) > 0"
			sSQL = sSQL & "union select Users as Leader, CASE Role WHEN N'" & STS_Assistant_PO_50000 & "' THEN N'#LEV4' ELSE N'#LEV2' END as Statuses from RolesForOrders_STS where (Role in (N'" & STS_Orders_CEO_SC & "', N'" & STS_Assistant_PO_50000 & "', N'" & STS_Assistant_PO_more_than_50000 & "')) and CharIndex(N'<" & Session("UserID") & ">', Users) > 0"
			AddLogD "Check Amount in PO SQL: " & sSQL
			dsTemp1.CursorLocation = 3
			dsTemp1.Open sSQL, Conn, 3, 1, &H1
			dsTemp1.ActiveConnection = Nothing
			If dsTemp1.EOF Then
				dsTemp1.Close
				Session("Message") = AddNewLineToMessage(Session("Message"), STS_ErrorCantCreatePOLargerLimit1)
				UserChangeDocCheck = False
				Exit Function
			Else
				If USD_Amount > 50000 Then
					'Если сумма выше второго рубежа шерстим записи на предмет уровня подразделения
					bDepDirector = False
					dsTemp1.MoveFirst
					Do While not dsTemp1.EOF and not bDepDirector
						iPos = InStr(dsTemp1("Statuses"), "#LEV")
						If iPos > 0 Then
							sLev = Trim(mid(dsTemp1("Statuses") & " ", iPos+4, 1))
							If IsNumeric(sLev) and CInt(sLev) <= 2 Then
								bDepDirector = True
							End If
						End If
						dsTemp1.MoveNext
					Loop
					If not bDepDirector Then
						dsTemp1.Close
						Session("Message") = AddNewLineToMessage(Session("Message"), STS_ErrorCantCreatePOLargerLimit2)
						UserChangeDocCheck = False
						Exit Function
					End If
				End If
				dsTemp1.Close
			End If
		End If
		'Запрос №43 - СТС - end

		sBusinessUnit = GetCodeFromCode_NameString(S_AddField2)
		AddLogD "@@@OrdersReconcilation - sBusinessUnit: "+sBusinessUnit

		'Запрос №36 - СТС - start - Подстановка по справочнику правил
		sMyNameResponsible = ""
		'Утверждающего всегда пересчитываем
		S_NameAproval = ""
		''Находим директора департамента инициатора
		'sHeadOfDepartment = GetRoleForOrders(STS_Orders_HeadOfDepartment, Request("DocDepartment"), sDocCreator, sBusinessUnit)
		''Если директор департамента инициатора отсутствует, ставим директора дивизиона
		'If Trim(GetUserID(sHeadOfDepartment)) = "" Then
		'	sHeadOfDepartment = GetRoleForOrders(STS_Orders_HeadOfDivision, Request("DocDepartment"), sDocCreator, sBusinessUnit)
		'End If
		''Согласующих берем из пользовательского ввода и добавляем директора департамент/дивизиона на новый уровень
		'S_ListToReconcile = S_ListToReconcile & VbCrLf & sHeadOfDepartment
		'Не обязательно сделано правилами      'Далее подставляем согласующих из правил. Первым правилом должен добавляться #USERINPUT#
'{Запрос №50 - СТС
		'GetReconciliationByRules(parDocID, parClassDoc, parAmount, parChartOfAccounts, parCostCenter, parProjectCode, parBusinessUnit, parPartnerName, parKindOfPayment, parInitiator, parProjectManager, parOvertimeRequester, parIncomeExpenceContract, parContranctInSTS, parCurrency, parContractType, parOvertimeFuncLeaders, ByRef parListToReconcile, ByRef parNameAproval, ByRef parNameResponsible, ByRef parCorrespondent, ByRef parListToView, ByRef parRegistrar, ByRef parNameControl, ByRef parListToEdit)
		'GetReconciliationByRules iif(Trim(Request("DocID")) = "", S_DocID, Request("DocID")), Session("CurrentClassDoc"), USD_Amount, GetCodeFromCode_NameString(S_UserFieldText8), GetCodeFromCode_NameString(S_UserFieldText1), GetCodeFromCode_NameString(Request("UserFieldText3")), sBusinessUnit, Request("DocPartnerName"), S_UserFieldText5, iif(UCase(Request("create")) = "Y", Session("UserID"), Request("DocAuthor")), ProjectManager, "", NULL, NULL, "USD", NULL, NULL, S_ListToReconcile, S_NameAproval, sMyNameResponsible, Null, S_ListToView, Null, Null, Null
		par_DocID = iif(Trim(Request("DocID")) = "", S_DocID, Request("DocID"))
		par_ClassDoc = Session("CurrentClassDoc")
		par_Amount = USD_Amount
		par_ChartOfAccounts = GetCodeFromCode_NameString(S_UserFieldText8)
		par_CostCenter = GetCodeFromCode_NameString(S_UserFieldText1)
		par_ProjectCode = GetCodeFromCode_NameString(Request("UserFieldText3"))
		par_BusinessUnit = sBusinessUnit
		par_PartnerName = Request("DocPartnerName")
		par_KindOfPayment = S_UserFieldText5
		par_Initiator = iif(UCase(Request("create")) = "Y", Session("UserID"), Request("DocAuthor"))
		par_ProjectManager = ProjectManager
		par_OvertimeRequester = ""
		par_IncomeExpenceContract = NULL
		par_ContranctInSTS = NULL
		par_Currency = "USD"
		par_ContractType = NULL
		par_OvertimeFuncLeaders = NULL
		par_TypeOfDocument = ""
		par_FunctionArea = ""

		GetReconciliationByRules par_DocID, par_ClassDoc, par_Amount, par_ChartOfAccounts, par_CostCenter, par_ProjectCode, par_BusinessUnit, par_PartnerName, par_KindOfPayment, par_Initiator, par_ProjectManager, par_OvertimeRequester, par_IncomeExpenceContract, par_ContranctInSTS, par_Currency, par_ContractType, par_OvertimeFuncLeaders, par_TypeOfDocument, par_FunctionArea, _
			S_ListToReconcile, S_NameAproval, sMyNameResponsible, Null, S_ListToView, Null, Null, Null
		'	par_ListToReconcile, par_NameAproval, par_NameResponsible, par_Correspondent, par_ListToView, par_Registrar, par_NameControl, par_ListToEdit
'Запрос №50 - СТС}
		'Могло быть переназначение исполнителя, подставляем из правила только если исполнитель не содержит конкретного пользователя
		S_NameResponsible = Trim(Request("DocNameResponsible"))
		If UCase(Request("create")) = "Y" Then
			S_NameResponsible = sMyNameResponsible
		End If
		S_NameResponsible_Set = S_NameResponsible

		If S_ListToReconcile = "" Then
			S_ListToReconcile = " "
		End If
		If S_NameAproval = "" Then
			S_NameAproval = " "
		End If
		'Запрос №36 - СТС - end

		If not IsAdmin() Then
			S_ListToReconcile_Set = S_ListToReconcile
			S_NameAproval_Set = S_NameAproval
		End If
	End If
	'Старый код - выброшен
	<!--INCLUDE FILE="PurchaseOrders.asp" -->

	'ph - 20090808 - start
	If Request("DocIDParent") <> "" Then
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		bMyDocIDParentChanged = False
		If Request("create") <> "y" Then
			'Проверяем изменился ли номер родительского документа, если идет редактирование
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocID"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				bMyDocIDParentChanged = Request("DocIDParent")<>dsTemp("DocIDParent")
			End If
			dsTemp.Close
		End If
		If bMyDocIDParentChanged Then
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocIDParent"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				'sMessage = ""
				'UserChangeDocCheck = oPayDox.IsReadAccessUser(Session("UserID"), sMessage, dsTemp, DOCS_All, DOCS_NoReadAccess, DOCS_NoAccess, USER_Department, VAR_ExtInt, VAR_AdminSecLevel, VAR_StatusActiveUser, NULL)
				UserChangeDocCheck = IsReadAccessRS(dsTemp)' IsReadAccessUser(Session("UserID"), dsTemp, NULL)
				If UserChangeDocCheck Then
					S_DocIDParent = dsTemp("DocID")
				End If
				dsTemp.Close
				If not UserChangeDocCheck Then
					Session("Message") = SIT_NoAccessToParentDoc'sMessage
					Exit Function
				End If
			End If
		End If
	End If
	'ph - 20090808 - end

' *** ЗАЯВКИ НА ОПЛАТУ СТС
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) = 1 Then
	'Создатель документа
	If UCase(Request("create")) = "Y" Then
		sDocCreator = Session("UserID")
	Else
		sDocCreator = GetUserID(Request("DocAuthor"))
	End If

	S_NameAproval=Request("DocNameAproval")
	S_NameResponsible=Request("DocNameResponsible")
	S_NameControl=Request("DocNameControl")
	S_ListToReconcile=Request("DocListToReconcile")
	S_ListToEdit=Request("DocListToEdit")
	S_ListToView=Request("DocListToView")
	S_Correspondent=Request("DocCorrespondent")

	'Определяем бизнес направление
	sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))

	' SAY 2008-07-21
	Set dsTempPR = Server.CreateObject("ADODB.Recordset")

' DmGorsky_5: меняем местами проверку: вначале ищем SIT_SITRU ("СИТРОНИКС ИТ"), потом, если не найдено, - SIT_SITRONICS ("СИТРОНИКС")
	If InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRU))=1 then ' DmGorsky
		sDepartment = SIT_SITRU ' DmGorsky
		sDepartmentCode = "" ' DmGorsky
'Запрос №1 - СИБ - start
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SIB_ROOT_DEPARTMENT))=1 then
		sDepartment = SIT_SIB
		sDepartmentCode = ""
'Запрос №1 - СИБ - end
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRONICS))=1 then
		sDepartment = SIT_SITRONICS
		sDepartmentCode = ""
	Else
		sDepartment = SIT_STS

		sSQL = "select code from Departments where Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"' or Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"/'"
		AddLogD "WAY sSQL="+sSQL
		AddLogD "WAY Conn="+Conn
		dsTempPR.Open sSQL, Conn, 3, 1, &H1

		if not dsTempPR.EOF Then
			sDepartmentCode = dsTempPR("code")
			if InStrRev(sDepartmentCode, "/") > 0 then
				sDepartmentCode = Right(sDepartmentCode, Len(sDepartmentCode)-InStrRev(sDepartmentCode, "/"))
			End If
		Else
			sDepartmentCode="000000"
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInDepartmentCode1 + DelOtherLangFromFolder(Request("DocDepartment")) + SIT_ErrorInDepartmentCode2)
			UserChangeDocCheck=False
			Exit Function
		End If

		dsTempPR.Close
	End If

	If Request("create") = "y" Then
		sPostfix = ""
		sSearchCol = "DocID"
		sPrePrefix = ""
		'SAY 2009-02-20
		sSufix = ""

		'SAY 2009-03-12
		S_DocID = ""

		S_DocID = GetNewOrderDocID(Session("CurrentClassDoc"))
		S_DocIDAdd = S_DocID
	End If
	If sDepartmentRoot <> SIT_STS or (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_RASP_DOCS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1) Then
		'реализация полужесткого маршрута
		S_ListToReconcile = Replace(S_ListToReconcile, Trim(SIT_AdditionalAgrees), "")
		'Для коммерческих предложений и заявок на закупку и оплату не нужна приписка про доп. согласующих  
		If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_COM_OFFERS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) <> 1 Then
			If S_ListToReconcile <> "" Then
				S_ListToReconcile = " " + SIT_AdditionalAgrees + DeleteConstPrefixFromList(S_ListToReconcile)
			End If
		End If
	End If
	'Для всех категорий кроме заявок подставляем пользователей вместо ролей
	If (InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1) and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) <> 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_DOGOVORI_NEW)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIS)) <> 0) Then
		'Получаем список ролей
		If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) = 1 Then 'Для договоров получаем список ролей на всех языках (могут быть роли на разных языках)
			sRoleList = GetFullRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"), "", "", "")
		Else
			sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))
		End If
		'Подставляем роли в поля
		S_NameAproval = ReplaceRolesInList(S_NameAproval, sRoleList)
		S_NameResponsible = ReplaceRolesInList(S_NameResponsible, sRoleList)
		S_NameControl = ReplaceRolesInList(S_NameControl, sRoleList)
		S_ListToEdit = ReplaceRolesInList(S_ListToEdit, sRoleList)
		S_ListToView = ReplaceRolesInList(S_ListToView, sRoleList)
		S_ListToReconcile = ReplaceRolesInList(S_ListToReconcile, sRoleList)

		'Удаляем повторы в списке согласования
		S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)

		'vnik_error
		'Убираем автора документа из списка согласующих, т.к. после активации он не сможет загрузить новый документ
		'(работает пока только, если пользователь создает документ не под ролью)
		S_ListToReconcile = Replace(S_ListToReconcile,Request("DocAuthor"),"")
		'vnik_error
	End If
	'Проверка полей с пользователями
	'Для заявок отключаем проверку утверждающего, т.к. он рассчитывается
	If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 Then
		UserChangeDocCheck = CheckSingleUserField(S_NameAproval)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameAproval")+SIT_ErrorInUserField2)
			Exit Function
		End If
		'Для заявки на закупку отключаем проверку ответственного, т.к. можно указывать подразделение
		'Для заявки на оплату тоже, т.к. автоматом ставится роль из специального справочника
		UserChangeDocCheck = CheckSingleUserField(S_NameResponsible)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameResponsible")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
	UserChangeDocCheck = CheckSingleUserField(S_NameControl)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameControl")+SIT_ErrorInUserField2)
		Exit Function
	End If
	If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) = 0 Then 'В исходящих может быть любой текст
		UserChangeDocCheck = CheckMultiUserField(S_Correspondent)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("Correspondent")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
'Запрос №1 - СИБ - start
	IF InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_SLUZH_ZAPISKA)) = 0) THEN
'Запрос №1 - СИБ - end
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) = 0 Then 'В СЗ на переработки работают правила
			UserChangeDocCheck = CheckMultiUserField(S_ListToReconcile)
			If not UserChangeDocCheck Then
				Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToReconcile")+SIT_ErrorInUserField2)
				Exit Function
			End If
'Запрос №1 - СИБ - start
		End If
	END IF
'Запрос №1 - СИБ - end
	UserChangeDocCheck = CheckMultiUserField(S_ListToEdit)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToEdit")+SIT_ErrorInUserField2)
		Exit Function
	End If
	'UserChangeDocCheck = CheckMultiUserField(S_ListToView)
	'If not UserChangeDocCheck Then
	'	Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToView")+SIT_ErrorInUserField2)
	'	Exit Function
	'End If

'{ph - 20120812
	'S_AdditionalUsers = GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit"))
	S_AdditionalUsers = DeleteUserDoublesInListNew(Request("AdditionalUsers") & VbCrLf & GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit")), True)
'ph - 20120812}
	AddLogD "@@@Search For Leaders - S_AdditionalUsers: "+S_AdditionalUsers
	'vnik_payment_order
	'If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_PAYMENT_ORDER)) > 0 Then
	'	S_AdditionalUsers = S_AdditionalUsers + ReplaceRoleFromDir(SIT_Account_manager,sDepartmentRoot)       
	'End If        
	'vnik_payment_order
	S_AdditionalUsers = DeleteUserDoublesInList(S_AdditionalUsers)

	'vnik
	If S_AdditionalUsers = "" Then
		S_AdditionalUsers = "-"
	End If
	S_AdditionalUsers_Set = S_AdditionalUsers

	S_ListToReconcile_Set = S_ListToReconcile
	S_NameAproval_Set = S_NameAproval
	S_ListToEdit = S_ListToEdit
	S_NameResponsible_Set = S_NameResponsible
	S_NameControl_Set = S_NameControl
	S_ListToView_Set=S_ListToView

	If UCase(Request("UpdateDoc")) = "YES" And (Request("create") = "y"  Or Request("DocIDPrev") <> Request("DocID")) Then
		'out "123"
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		dsTemp.Open "select * from Docs where DocID='" + MakeSQLSafe(Request("DocID")) + "'", Conn 
		If Not dsTemp.EOF Then
			Session("Message") = DOCS_ALREADYEXISTS + ": " + Request("DocID")
			dsTemp.Close
			UserChangeDocCheck=False
			Exit Function
		End If
		dsTemp.Close
	End If

	'Проверка даты исполнения заявок
	If Trim(Request("DocDateCompletion")) <> "" Then
		If ConvertToDate(Request("DocDateCompletion")) <> VAR_BeginOfTimes Then
			If ConvertToDate(Request("DocDateCompletion")) < Date Then
				Session("Message") = SIT_TooEarlyDate+Request("DocDateCompletion")
				UserChangeDocCheck = False
				Exit Function
			End If
		Else
			Session("Message") = SIT_DateFormatError+Request("DocDateCompletion")
			UserChangeDocCheck = False
			Exit Function
		End If
	End If

	'Запрос №36 - СТС - start
	'Получаем лимиты из БД
	'STS_HeadOfSector_Limit = 0
	'STS_HeadOfDepartment_Limit = 0
	'STS_HeadOfDivision_Limit = 0
	'STS_FinancialControl_Limit = 0
	'STS_FinDirector_Limit = 0
	'STS_GenDirector_Limit = 0
	'STS_ProjectManager_Limit = 0
	'STS_Accounting_Limit = 0
	'GetLimitsForOrders STS_HeadOfSector_Limit, STS_HeadOfDepartment_Limit, STS_HeadOfDivision_Limit, STS_FinancialControl_Limit, STS_FinDirector_Limit, STS_GenDirector_Limit, STS_ProjectManager_Limit, STS_Accounting_Limit
	'Запрос №36 - СТС - end

	'Проверка наличия родительского документа
	Set dsTemp1 = Server.CreateObject("ADODB.Recordset")
	sSQL = "select * from Docs where DocID = "+sUnicodeSymbol+"'" + Request("DocIDParent") + "'"
	AddLogD "@@@CheckParentDocID SQL: "+sSQL
	dsTemp1.Open sSQL, Conn, 3, 1, &H1
	If not dsTemp1.EOF Then
		'ph - 20101028 - start
		'Родительский документ - не заявка на закупку
		If InStr(MyCStr(dsTemp1("ClassDoc")), STS_PurchaseOrder) <> 1 Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInParentOrderNumber)
			UserChangeDocCheck = False
			dsTemp1.Close
			Exit Function
		End If
		'ph - 20101028 - end
		'По отмененным заявкам на закупку нельзя создавать заявки на оплату
		If dsTemp1("StatusCompletion") = "0" Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ParentOrderCanceled)
			UserChangeDocCheck = False
			dsTemp1.Close
			Exit Function
		End If
		'По неутвержденным заявкам на закупку нельзя создавать заявки на оплату
		If MyCStr(dsTemp1("NameApproved")) = "" Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ParentOrderNotApproved)
			UserChangeDocCheck = False
			dsTemp1.Close
			Exit Function
		End If

		'Часть полей берем из Заявки на закупку
		If Trim(S_UserFieldText1) = "" Then
			S_UserFieldText1 = Trim(CStr(dsTemp1("UserFieldText1")))
			S_UserFieldText1_Set = S_UserFieldText1
		End If
		If Trim(S_AddField2) = "" Then
			S_AddField2 = Trim(CStr(dsTemp1("BusinessUnit")))
			S_AddField_Set2 = S_AddField2
		End If
		If Trim(S_UserFieldText3) = "" Then
			S_UserFieldText3 = Trim(CStr(dsTemp1("UserFieldText3")))
			S_UserFieldText3_Set = S_UserFieldText3
		End If
		If Trim(S_UserFieldText4) = "" Then
			S_UserFieldText4 = Trim(CStr(dsTemp1("UserFieldText4")))
			S_UserFieldText4_Set = S_UserFieldText4
		End If
		If Trim(S_UserFieldText6) = "" Then
			S_UserFieldText6 = Trim(CStr(dsTemp1("UserFieldText6")))
			S_UserFieldText6_Set = S_UserFieldText6
		End If
		If Trim(S_UserFieldText7) = "" Then
			S_UserFieldText7 = Trim(CStr(dsTemp1("UserFieldText7")))
			S_UserFieldText7_Set = S_UserFieldText7
		End If
		If Trim(S_UserFieldText8) = "" Then
			S_UserFieldText8 = Trim(CStr(dsTemp1("UserFieldText8")))
			S_UserFieldText8_Set = S_UserFieldText8
		End If
		sBusinessUnit = GetCodeFromCode_NameString(S_AddField2)
		sCostCenter = GetCostCenterByCode(GetCodeFromCode_NameString(S_UserFieldText1))
	Else
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInParentOrderNumber)
		UserChangeDocCheck = False
		dsTemp1.Close
		Exit Function
	End If
	dsTemp1.Close

	'Проверка, что Форма расчета соответствует справочнику и замена текущего значения на англ. вариант
	sSQL = PaymentTypesSelectForCheckValue(Request("UserFieldText5"))
	AddLogD "@@@CheckPaymentType PaymentOrders SQL: "+sSQL
	dsTemp1.Open sSQL, Conn, 3, 1, &H1
	If dsTemp1.EOF Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInPaymentType)
		UserChangeDocCheck = False
		dsTemp1.Close
		Exit Function
	End If
	S_UserFieldText5 = dsTemp1("PaymentType_EN")
	dsTemp1.Close

	ErrorInAmount = True
	If Request("DocAmountDoc") <> "" Then
		If IsNumeric(Request("DocAmountDoc")) Then
			If Request("Currency") = "USD" Then
				USD_Amount = CCur(CorrectSum(Request("DocAmountDoc")))
			Else
				USD_Amount = CCur(CorrectSum(Request("DocAmountDoc")))*CurrencyConvertionFactor(Request("DocCurrency"), "USD")
			End If
			ErrorInAmount = USD_Amount <= 0
		End If
	End If
	If ErrorInAmount Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInSumOrCurrency)
		UserChangeDocCheck = False
		Exit Function
	Else
		Set dsTemp1 = Server.CreateObject("ADODB.Recordset")
		sSQL = "select * from Docs where DocID = "+sUnicodeSymbol+"'" + Request("DocIDParent") + "'"
		AddLogD "@@@SearchingParentDoc SQL: "+sSQL
		dsTemp1.Open sSQL, Conn, 3, 1, &H1
		If Not dsTemp1.EOF Then
			nPurchaseOrderAmountUSD = CCur(dsTemp1("UserFieldMoney1"))
			dsTemp1.Close
			sSQL = "select IsNull(Sum(UserFieldMoney1), 0) as SumUSD from Docs where DocIDParent = "+sUnicodeSymbol+"'" + Request("DocIDParent") + "' and ClassDoc like "+sUnicodeSymbol+"'"+STS_PaymentOrder+"%'"
			If Request("create") <> "y" Then
				sSQL = sSQL + " and DocID <> "+sUnicodeSymbol+"'"+Request("DocID")+"'"
			End If
			AddLogD "@@@SumOfChildPaymentOrders SQL: "+sSQL
			dsTemp1.Open sSQL, Conn, 3, 1, &H1
			nPaymentOrdersSumUSD = CCur(dsTemp1("SumUSD"))+USD_Amount
			'If nPaymentOrdersSumUSD > nPurchaseOrderAmountUSD Then
			If nPaymentOrdersSumUSD-nPurchaseOrderAmountUSD > 0.005 Then
				Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorSumExceeding)
				'UserChangeDocCheck = False
				'Exit Function
			End If
		End If
		dsTemp1.Close
	End If
	S_UserFieldMoney1 = CStr(USD_Amount)

	'Запрос №23 - СТС - start
	'Кусок из старого алгоритма - start
	'Если нулевой проект добавляем финансового контролера в дополнительные пользователи
	If not IsProject(S_UserFieldText3) Then
		S_AdditionalUsers = GetRoleForOrders(STS_Orders_FinancialControl, sCostCenter, "", sBusinessUnit)
		If S_AdditionalUsers_Set = "-" Then
			S_AdditionalUsers_Set = S_AdditionalUsers
		Else
			S_AdditionalUsers_Set = S_AdditionalUsers_Set & VbCrLf & S_AdditionalUsers
		End If
	End If
	'Запрос №31 - СТС - start - исполнитель тоже устанавливается через правила
	sMyNameResponsible = ""
'{Запрос №50 - СТС
	'GetReconciliationByRules(parDocID, parClassDoc, parAmount, parChartOfAccounts, parCostCenter, parProjectCode, parBusinessUnit, parPartnerName, parKindOfPayment, parInitiator, parProjectManager, parOvertimeRequester, parIncomeExpenceContract, parContranctInSTS, parCurrency, parContractType, parOvertimeFuncLeaders, ByRef parListToReconcile, ByRef parNameAproval, ByRef parNameResponsible, ByRef parCorrespondent, ByRef parListToView, ByRef parRegistrar, ByRef parNameControl, ByRef parListToEdit)
	'GetReconciliationByRules iif(Trim(Request("DocID")) = "", S_DocID, Request("DocID")), Session("CurrentClassDoc"), USD_Amount, GetCodeFromCode_NameString(S_UserFieldText8), GetCodeFromCode_NameString(S_UserFieldText1), GetCodeFromCode_NameString(S_UserFieldText3), sBusinessUnit, Request("DocPartnerName"), S_UserFieldText5, iif(UCase(Request("create")) = "Y", Session("UserID"), Request("DocAuthor")), "", "", NULL, NULL, "USD", NULL, NULL, S_ListToReconcile, S_NameAproval, sMyNameResponsible, Null, Null, Null, Null, Null
	par_DocID = iif(Trim(Request("DocID")) = "", S_DocID, Request("DocID"))
	par_ClassDoc = Session("CurrentClassDoc")
	par_Amount = USD_Amount
	par_ChartOfAccounts = GetCodeFromCode_NameString(S_UserFieldText8)
	par_CostCenter = GetCodeFromCode_NameString(S_UserFieldText1)
	par_ProjectCode = GetCodeFromCode_NameString(S_UserFieldText3)
	par_BusinessUnit = sBusinessUnit
	par_PartnerName = Request("DocPartnerName")
	par_KindOfPayment = S_UserFieldText5
	par_Initiator = iif(UCase(Request("create")) = "Y", Session("UserID"), Request("DocAuthor"))
	par_ProjectManager = ""
	par_OvertimeRequester = ""
	par_IncomeExpenceContract = NULL
	par_ContranctInSTS = NULL
	par_Currency = "USD"
	par_ContractType = NULL
	par_OvertimeFuncLeaders = NULL
	par_TypeOfDocument = ""
	par_FunctionArea = ""

	GetReconciliationByRules par_DocID, par_ClassDoc, par_Amount, par_ChartOfAccounts, par_CostCenter, par_ProjectCode, par_BusinessUnit, par_PartnerName, par_KindOfPayment, par_Initiator, par_ProjectManager, par_OvertimeRequester, par_IncomeExpenceContract, par_ContranctInSTS, par_Currency, par_ContractType, par_OvertimeFuncLeaders, par_TypeOfDocument, par_FunctionArea, _
		S_ListToReconcile, S_NameAproval, sMyNameResponsible, Null, Null, Null, Null, Null
	'	par_ListToReconcile, par_NameAproval, par_NameResponsible, par_Correspondent, par_ListToView, par_Registrar, par_NameControl, par_ListToEdit
'Запрос №50 - СТС}
	'Могло быть переназначение исполнителя, подставляем из правила только если исполнитель не содержит конкретного пользователя
	S_NameResponsible = Trim(Request("DocNameResponsible"))
	If UCase(Request("create")) = "Y" Then
		S_NameResponsible = sMyNameResponsible
	End If
	'ph - 20101223 - end
	If S_ListToReconcile = "" Then
		S_ListToReconcile = "-"
	End If
	'Запрос №31 - СТС - end
	If S_NameAproval = "" Then
		S_NameAproval = "-"
	End If
	'Запрос №23 - СТС  - end

	'ph - 20090808 - start
	If Request("DocIDParent") <> "" Then
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		bMyDocIDParentChanged = False
		If Request("create") <> "y" Then
			'Проверяем изменился ли номер родительского документа, если идет редактирование
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocID"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				bMyDocIDParentChanged = Request("DocIDParent")<>dsTemp("DocIDParent")
			End If
			dsTemp.Close
		End If
		If bMyDocIDParentChanged Then
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocIDParent"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				'sMessage = ""
				'UserChangeDocCheck = oPayDox.IsReadAccessUser(Session("UserID"), sMessage, dsTemp, DOCS_All, DOCS_NoReadAccess, DOCS_NoAccess, USER_Department, VAR_ExtInt, VAR_AdminSecLevel, VAR_StatusActiveUser, NULL)
				UserChangeDocCheck = IsReadAccessRS(dsTemp)' IsReadAccessUser(Session("UserID"), dsTemp, NULL)
				If UserChangeDocCheck Then
					S_DocIDParent = dsTemp("DocID")
				End If
				dsTemp.Close
				If not UserChangeDocCheck Then
					Session("Message") = SIT_NoAccessToParentDoc'sMessage
					Exit Function
				End If
			End If
		End If
	End If
	'ph - 20090808 - end

' *** ПОДТИПЫ СЛУЖЕБНЫХ ЗАПИСОК
'SIT_SLUZH_ZAPISKA_COMPUTER
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_SLUZH_ZAPISKA_COMPUTER)) = 1 Then
	'Создатель документа
	If UCase(Request("create")) = "Y" Then
		sDocCreator = Session("UserID")
	Else
		sDocCreator = GetUserID(Request("DocAuthor"))
	End If

	S_NameAproval=Request("DocNameAproval")
	S_NameResponsible=Request("DocNameResponsible")
	S_NameControl=Request("DocNameControl")
	S_ListToReconcile=Request("DocListToReconcile")
	S_ListToEdit=Request("DocListToEdit")
	S_ListToView=Request("DocListToView")
	S_Correspondent=Request("DocCorrespondent")

	'Определяем бизнес направление
	sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))

	' SAY 2008-07-21
	Set dsTempPR = Server.CreateObject("ADODB.Recordset")

' DmGorsky_5: меняем местами проверку: вначале ищем SIT_SITRU ("СИТРОНИКС ИТ"), потом, если не найдено, - SIT_SITRONICS ("СИТРОНИКС")
	If InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRU))=1 then ' DmGorsky
		sDepartment = SIT_SITRU ' DmGorsky
		sDepartmentCode = "" ' DmGorsky
'Запрос №1 - СИБ - start
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SIB_ROOT_DEPARTMENT))=1 then
		sDepartment = SIT_SIB
		sDepartmentCode = ""
'Запрос №1 - СИБ - end
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRONICS))=1 then
		sDepartment = SIT_SITRONICS
		sDepartmentCode = ""
	Else
		sDepartment = SIT_STS

		sSQL = "select code from Departments where Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"' or Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"/'"
		AddLogD "WAY sSQL="+sSQL
		AddLogD "WAY Conn="+Conn
		dsTempPR.Open sSQL, Conn, 3, 1, &H1

		if not dsTempPR.EOF Then
			sDepartmentCode = dsTempPR("code")
			if InStrRev(sDepartmentCode, "/") > 0 then
				sDepartmentCode = Right(sDepartmentCode, Len(sDepartmentCode)-InStrRev(sDepartmentCode, "/"))
			End If
		Else
			sDepartmentCode="000000"
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInDepartmentCode1 + DelOtherLangFromFolder(Request("DocDepartment")) + SIT_ErrorInDepartmentCode2)
			UserChangeDocCheck=False
			Exit Function
		End If

		dsTempPR.Close
	End If

	If Request("create") = "y" Then
		sPostfix = ""
		sSearchCol = "DocID"
		sPrePrefix = ""
		'SAY 2009-02-20
		sSufix = ""

		'SAY 2009-03-12
		S_DocID = ""

		sClassDoc = "Служебные записки%"
		S_DocID = GetNewDocIDForSluzhZapiski(sClassDoc, sDepartmentRoot, Request("UserFieldText1"), "")
		S_DocIDAdd = S_DocID
	End If
	If sDepartmentRoot <> SIT_STS or (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_RASP_DOCS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1) Then
		'реализация полужесткого маршрута
		S_ListToReconcile = Replace(S_ListToReconcile, Trim(SIT_AdditionalAgrees), "")
		'Для коммерческих предложений и заявок на закупку и оплату не нужна приписка про доп. согласующих  
		If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_COM_OFFERS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) <> 1 Then
			If S_ListToReconcile <> "" Then
				S_ListToReconcile = " " + SIT_AdditionalAgrees + DeleteConstPrefixFromList(S_ListToReconcile)
			End If
		End If
	End If

	If sDepartmentRoot = SIT_SITRONICS Then
		S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
	ElseIf sDepartmentRoot = SIT_STS Then
		'rmanyushin 89142 01.04.2010 Start
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME)) > 0 Then
			If is789DivisionSTS(Session("Department")) Then
				DocSession = "Служебные записки*Office memo*Interní sdělení/На переработки 7-8-9*Overtime 7-8-9*Overtime 7-8-9"
				S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",DocSession + "/","List")+STS_DptOSSBSS+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
				'S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+STS_DptOSSBSS+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
			Else
				S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
			End if
		End If
		'rmanyushin 89142 01.04.2010 Stop

		'rmanyushin 119579 19.08.2010 Start
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_HOLIDAY)) > 0 Then
			S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
		End If
		'rmanyushin 119579 19.08.2010 End
	ElseIf sDepartmentRoot = SIT_SIB Then'СИБ
		S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSIBRU","Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
	Else 'Другие БН
		S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
	End If

	'Для всех категорий кроме заявок подставляем пользователей вместо ролей
	If (InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1) and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) <> 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_DOGOVORI_NEW)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIS)) <> 0) Then
		'Получаем список ролей
		If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) = 1 Then 'Для договоров получаем список ролей на всех языках (могут быть роли на разных языках)
			sRoleList = GetFullRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"), "", "", "")
		Else
			sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))
		End If
		'Подставляем роли в поля
		S_NameAproval = ReplaceRolesInList(S_NameAproval, sRoleList)
		S_NameResponsible = ReplaceRolesInList(S_NameResponsible, sRoleList)
		S_NameControl = ReplaceRolesInList(S_NameControl, sRoleList)
		S_ListToEdit = ReplaceRolesInList(S_ListToEdit, sRoleList)
		S_ListToView = ReplaceRolesInList(S_ListToView, sRoleList)
		S_ListToReconcile = ReplaceRolesInList(S_ListToReconcile, sRoleList)

		'Удаляем повторы в списке согласования
		S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)

		'vnik_error
		'Убираем автора документа из списка согласующих, т.к. после активации он не сможет загрузить новый документ
		'(работает пока только, если пользователь создает документ не под ролью)
		S_ListToReconcile = Replace(S_ListToReconcile,Request("DocAuthor"),"")
		'vnik_error
	End If
	'Проверка полей с пользователями
	'Для заявок отключаем проверку утверждающего, т.к. он рассчитывается
	If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 Then
		UserChangeDocCheck = CheckSingleUserField(S_NameAproval)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameAproval")+SIT_ErrorInUserField2)
			Exit Function
		End If
		'Для заявки на закупку отключаем проверку ответственного, т.к. можно указывать подразделение
		'Для заявки на оплату тоже, т.к. автоматом ставится роль из специального справочника
		UserChangeDocCheck = CheckSingleUserField(S_NameResponsible)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameResponsible")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
	UserChangeDocCheck = CheckSingleUserField(S_NameControl)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameControl")+SIT_ErrorInUserField2)
		Exit Function
	End If
	If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) = 0 Then 'В исходящих может быть любой текст
		UserChangeDocCheck = CheckMultiUserField(S_Correspondent)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("Correspondent")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
'Запрос №1 - СИБ - start
	IF InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_SLUZH_ZAPISKA)) = 0) THEN
'Запрос №1 - СИБ - end
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) = 0 Then 'В СЗ на переработки работают правила
			UserChangeDocCheck = CheckMultiUserField(S_ListToReconcile)
			If not UserChangeDocCheck Then
				Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToReconcile")+SIT_ErrorInUserField2)
				Exit Function
			End If
'Запрос №1 - СИБ - start
		End If
	END IF
'Запрос №1 - СИБ - end
	UserChangeDocCheck = CheckMultiUserField(S_ListToEdit)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToEdit")+SIT_ErrorInUserField2)
		Exit Function
	End If
	'UserChangeDocCheck = CheckMultiUserField(S_ListToView)
	'If not UserChangeDocCheck Then
	'	Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToView")+SIT_ErrorInUserField2)
	'	Exit Function
	'End If

'{ph - 20120812
	'S_AdditionalUsers = GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit"))
	S_AdditionalUsers = DeleteUserDoublesInListNew(Request("AdditionalUsers") & VbCrLf & GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit")), True)
'ph - 20120812}
	AddLogD "@@@Search For Leaders - S_AdditionalUsers: "+S_AdditionalUsers
	'vnik_payment_order
	'If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_PAYMENT_ORDER)) > 0 Then
	'	S_AdditionalUsers = S_AdditionalUsers + ReplaceRoleFromDir(SIT_Account_manager,sDepartmentRoot)       
	'End If        
	'vnik_payment_order
	S_AdditionalUsers = DeleteUserDoublesInList(S_AdditionalUsers)

	'vnik
	If S_AdditionalUsers = "" Then
		S_AdditionalUsers = "-"
	End If
	S_AdditionalUsers_Set = S_AdditionalUsers

	S_ListToReconcile_Set = S_ListToReconcile
	S_NameAproval_Set = S_NameAproval
	S_ListToEdit = S_ListToEdit
	S_NameResponsible_Set = S_NameResponsible
	S_NameControl_Set = S_NameControl
	S_ListToView_Set=S_ListToView

	If UCase(Request("UpdateDoc")) = "YES" And (Request("create") = "y"  Or Request("DocIDPrev") <> Request("DocID")) Then
		'out "123"
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		dsTemp.Open "select * from Docs where DocID='" + MakeSQLSafe(Request("DocID")) + "'", Conn 
		If Not dsTemp.EOF Then
			Session("Message") = DOCS_ALREADYEXISTS + ": " + Request("DocID")
			dsTemp.Close
			UserChangeDocCheck=False
			Exit Function
		End If
		dsTemp.Close
	End If

	'ph - 20090808 - start
	If Request("DocIDParent") <> "" Then
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		bMyDocIDParentChanged = False
		If Request("create") <> "y" Then
			'Проверяем изменился ли номер родительского документа, если идет редактирование
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocID"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				bMyDocIDParentChanged = Request("DocIDParent")<>dsTemp("DocIDParent")
			End If
			dsTemp.Close
		End If
		If bMyDocIDParentChanged Then
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocIDParent"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				'sMessage = ""
				'UserChangeDocCheck = oPayDox.IsReadAccessUser(Session("UserID"), sMessage, dsTemp, DOCS_All, DOCS_NoReadAccess, DOCS_NoAccess, USER_Department, VAR_ExtInt, VAR_AdminSecLevel, VAR_StatusActiveUser, NULL)
				UserChangeDocCheck = IsReadAccessRS(dsTemp)' IsReadAccessUser(Session("UserID"), dsTemp, NULL)
				If UserChangeDocCheck Then
					S_DocIDParent = dsTemp("DocID")
				End If
				dsTemp.Close
				If not UserChangeDocCheck Then
					Session("Message") = SIT_NoAccessToParentDoc'sMessage
					Exit Function
				End If
			End If
		End If
	End If
	'ph - 20090808 - end

'SIT_SLUZH_ZAPISKA_MOBILE
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_SLUZH_ZAPISKA_MOBILE)) = 1 Then
	'Создатель документа
	If UCase(Request("create")) = "Y" Then
		sDocCreator = Session("UserID")
	Else
		sDocCreator = GetUserID(Request("DocAuthor"))
	End If

	S_NameAproval=Request("DocNameAproval")
	S_NameResponsible=Request("DocNameResponsible")
	S_NameControl=Request("DocNameControl")
	S_ListToReconcile=Request("DocListToReconcile")
	S_ListToEdit=Request("DocListToEdit")
	S_ListToView=Request("DocListToView")
	S_Correspondent=Request("DocCorrespondent")

	'Определяем бизнес направление
	sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))

	' SAY 2008-07-21
	Set dsTempPR = Server.CreateObject("ADODB.Recordset")

' DmGorsky_5: меняем местами проверку: вначале ищем SIT_SITRU ("СИТРОНИКС ИТ"), потом, если не найдено, - SIT_SITRONICS ("СИТРОНИКС")
	If InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRU))=1 then ' DmGorsky
		sDepartment = SIT_SITRU ' DmGorsky
		sDepartmentCode = "" ' DmGorsky
'Запрос №1 - СИБ - start
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SIB_ROOT_DEPARTMENT))=1 then
		sDepartment = SIT_SIB
		sDepartmentCode = ""
'Запрос №1 - СИБ - end
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRONICS))=1 then
		sDepartment = SIT_SITRONICS
		sDepartmentCode = ""
	Else
		sDepartment = SIT_STS

		sSQL = "select code from Departments where Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"' or Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"/'"
		AddLogD "WAY sSQL="+sSQL
		AddLogD "WAY Conn="+Conn
		dsTempPR.Open sSQL, Conn, 3, 1, &H1

		if not dsTempPR.EOF Then
			sDepartmentCode = dsTempPR("code")
			if InStrRev(sDepartmentCode, "/") > 0 then
				sDepartmentCode = Right(sDepartmentCode, Len(sDepartmentCode)-InStrRev(sDepartmentCode, "/"))
			End If
		Else
			sDepartmentCode="000000"
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInDepartmentCode1 + DelOtherLangFromFolder(Request("DocDepartment")) + SIT_ErrorInDepartmentCode2)
			UserChangeDocCheck=False
			Exit Function
		End If

		dsTempPR.Close
	End If

	If Request("create") = "y" Then
		sPostfix = ""
		sSearchCol = "DocID"
		sPrePrefix = ""
		'SAY 2009-02-20
		sSufix = ""

		'SAY 2009-03-12
		S_DocID = ""

		sClassDoc = "Служебные записки%"
		S_DocID = GetNewDocIDForSluzhZapiski(sClassDoc, sDepartmentRoot, Request("UserFieldText1"), "")
		S_DocIDAdd = S_DocID
	End If
	If sDepartmentRoot <> SIT_STS or (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_RASP_DOCS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1) Then
		'реализация полужесткого маршрута
		S_ListToReconcile = Replace(S_ListToReconcile, Trim(SIT_AdditionalAgrees), "")
		'Для коммерческих предложений и заявок на закупку и оплату не нужна приписка про доп. согласующих  
		If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_COM_OFFERS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) <> 1 Then
			If S_ListToReconcile <> "" Then
				S_ListToReconcile = " " + SIT_AdditionalAgrees + DeleteConstPrefixFromList(S_ListToReconcile)
			End If
		End If
	End If

	If sDepartmentRoot = SIT_SITRONICS Then
		S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
	ElseIf sDepartmentRoot = SIT_STS Then
		'rmanyushin 89142 01.04.2010 Start
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME)) > 0 Then
			If is789DivisionSTS(Session("Department")) Then
				DocSession = "Служебные записки*Office memo*Interní sdělení/На переработки 7-8-9*Overtime 7-8-9*Overtime 7-8-9"
				S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",DocSession + "/","List")+STS_DptOSSBSS+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
				'S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+STS_DptOSSBSS+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
			Else
				S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
			End if
		End If
		'rmanyushin 89142 01.04.2010 Stop

		'rmanyushin 119579 19.08.2010 Start
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_HOLIDAY)) > 0 Then
			S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
		End If
		'rmanyushin 119579 19.08.2010 End
	ElseIf sDepartmentRoot = SIT_SIB Then'СИБ
		S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSIBRU","Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
	Else 'Другие БН
		S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
	End If

	'Для всех категорий кроме заявок подставляем пользователей вместо ролей
	If (InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1) and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) <> 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_DOGOVORI_NEW)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIS)) <> 0) Then
		'Получаем список ролей
		If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) = 1 Then 'Для договоров получаем список ролей на всех языках (могут быть роли на разных языках)
			sRoleList = GetFullRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"), "", "", "")
		Else
			sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))
		End If
		'Подставляем роли в поля
		S_NameAproval = ReplaceRolesInList(S_NameAproval, sRoleList)
		S_NameResponsible = ReplaceRolesInList(S_NameResponsible, sRoleList)
		S_NameControl = ReplaceRolesInList(S_NameControl, sRoleList)
		S_ListToEdit = ReplaceRolesInList(S_ListToEdit, sRoleList)
		S_ListToView = ReplaceRolesInList(S_ListToView, sRoleList)
		S_ListToReconcile = ReplaceRolesInList(S_ListToReconcile, sRoleList)

		'Удаляем повторы в списке согласования
		S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)

		'vnik_error
		'Убираем автора документа из списка согласующих, т.к. после активации он не сможет загрузить новый документ
		'(работает пока только, если пользователь создает документ не под ролью)
		S_ListToReconcile = Replace(S_ListToReconcile,Request("DocAuthor"),"")
		'vnik_error
	End If
	'Проверка полей с пользователями
	'Для заявок отключаем проверку утверждающего, т.к. он рассчитывается
	If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 Then
		UserChangeDocCheck = CheckSingleUserField(S_NameAproval)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameAproval")+SIT_ErrorInUserField2)
			Exit Function
		End If
		'Для заявки на закупку отключаем проверку ответственного, т.к. можно указывать подразделение
		'Для заявки на оплату тоже, т.к. автоматом ставится роль из специального справочника
		UserChangeDocCheck = CheckSingleUserField(S_NameResponsible)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameResponsible")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
	UserChangeDocCheck = CheckSingleUserField(S_NameControl)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameControl")+SIT_ErrorInUserField2)
		Exit Function
	End If
	If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) = 0 Then 'В исходящих может быть любой текст
		UserChangeDocCheck = CheckMultiUserField(S_Correspondent)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("Correspondent")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
'Запрос №1 - СИБ - start
	IF InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_SLUZH_ZAPISKA)) = 0) THEN
'Запрос №1 - СИБ - end
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) = 0 Then 'В СЗ на переработки работают правила
			UserChangeDocCheck = CheckMultiUserField(S_ListToReconcile)
			If not UserChangeDocCheck Then
				Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToReconcile")+SIT_ErrorInUserField2)
				Exit Function
			End If
'Запрос №1 - СИБ - start
		End If
	END IF
'Запрос №1 - СИБ - end
	UserChangeDocCheck = CheckMultiUserField(S_ListToEdit)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToEdit")+SIT_ErrorInUserField2)
		Exit Function
	End If
	'UserChangeDocCheck = CheckMultiUserField(S_ListToView)
	'If not UserChangeDocCheck Then
	'	Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToView")+SIT_ErrorInUserField2)
	'	Exit Function
	'End If

'{ph - 20120812
	'S_AdditionalUsers = GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit"))
	S_AdditionalUsers = DeleteUserDoublesInListNew(Request("AdditionalUsers") & VbCrLf & GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit")), True)
'ph - 20120812}
	AddLogD "@@@Search For Leaders - S_AdditionalUsers: "+S_AdditionalUsers
	'vnik_payment_order
	'If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_PAYMENT_ORDER)) > 0 Then
	'	S_AdditionalUsers = S_AdditionalUsers + ReplaceRoleFromDir(SIT_Account_manager,sDepartmentRoot)       
	'End If        
	'vnik_payment_order
	S_AdditionalUsers = DeleteUserDoublesInList(S_AdditionalUsers)

	'vnik
	If S_AdditionalUsers = "" Then
		S_AdditionalUsers = "-"
	End If
	S_AdditionalUsers_Set = S_AdditionalUsers

	S_ListToReconcile_Set = S_ListToReconcile
	S_NameAproval_Set = S_NameAproval
	S_ListToEdit = S_ListToEdit
	S_NameResponsible_Set = S_NameResponsible
	S_NameControl_Set = S_NameControl
	S_ListToView_Set=S_ListToView

	If UCase(Request("UpdateDoc")) = "YES" And (Request("create") = "y"  Or Request("DocIDPrev") <> Request("DocID")) Then
		'out "123"
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		dsTemp.Open "select * from Docs where DocID='" + MakeSQLSafe(Request("DocID")) + "'", Conn 
		If Not dsTemp.EOF Then
			Session("Message") = DOCS_ALREADYEXISTS + ": " + Request("DocID")
			dsTemp.Close
			UserChangeDocCheck=False
			Exit Function
		End If
		dsTemp.Close
	End If

	'ph - 20090808 - start
	If Request("DocIDParent") <> "" Then
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		bMyDocIDParentChanged = False
		If Request("create") <> "y" Then
			'Проверяем изменился ли номер родительского документа, если идет редактирование
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocID"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				bMyDocIDParentChanged = Request("DocIDParent")<>dsTemp("DocIDParent")
			End If
			dsTemp.Close
		End If
		If bMyDocIDParentChanged Then
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocIDParent"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				'sMessage = ""
				'UserChangeDocCheck = oPayDox.IsReadAccessUser(Session("UserID"), sMessage, dsTemp, DOCS_All, DOCS_NoReadAccess, DOCS_NoAccess, USER_Department, VAR_ExtInt, VAR_AdminSecLevel, VAR_StatusActiveUser, NULL)
				UserChangeDocCheck = IsReadAccessRS(dsTemp)' IsReadAccessUser(Session("UserID"), dsTemp, NULL)
				If UserChangeDocCheck Then
					S_DocIDParent = dsTemp("DocID")
				End If
				dsTemp.Close
				If not UserChangeDocCheck Then
					Session("Message") = SIT_NoAccessToParentDoc'sMessage
					Exit Function
				End If
			End If
		End If
	End If
	'ph - 20090808 - end

'SIT_SLUZH_ZAPISKA_KOMANDIROVKA
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_SLUZH_ZAPISKA_KOMANDIROVKA)) = 1 Then
	'Создатель документа
	If UCase(Request("create")) = "Y" Then
		sDocCreator = Session("UserID")
	Else
		sDocCreator = GetUserID(Request("DocAuthor"))
	End If

	S_NameAproval=Request("DocNameAproval")
	S_NameResponsible=Request("DocNameResponsible")
	S_NameControl=Request("DocNameControl")
	S_ListToReconcile=Request("DocListToReconcile")
	S_ListToEdit=Request("DocListToEdit")
	S_ListToView=Request("DocListToView")
	S_Correspondent=Request("DocCorrespondent")

	'Определяем бизнес направление
	sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))

	' SAY 2008-07-21
	Set dsTempPR = Server.CreateObject("ADODB.Recordset")

' DmGorsky_5: меняем местами проверку: вначале ищем SIT_SITRU ("СИТРОНИКС ИТ"), потом, если не найдено, - SIT_SITRONICS ("СИТРОНИКС")
	If InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRU))=1 then ' DmGorsky
		sDepartment = SIT_SITRU ' DmGorsky
		sDepartmentCode = "" ' DmGorsky
'Запрос №1 - СИБ - start
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SIB_ROOT_DEPARTMENT))=1 then
		sDepartment = SIT_SIB
		sDepartmentCode = ""
'Запрос №1 - СИБ - end
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRONICS))=1 then
		sDepartment = SIT_SITRONICS
		sDepartmentCode = ""
	Else
		sDepartment = SIT_STS

		sSQL = "select code from Departments where Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"' or Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"/'"
		AddLogD "WAY sSQL="+sSQL
		AddLogD "WAY Conn="+Conn
		dsTempPR.Open sSQL, Conn, 3, 1, &H1

		if not dsTempPR.EOF Then
			sDepartmentCode = dsTempPR("code")
			if InStrRev(sDepartmentCode, "/") > 0 then
				sDepartmentCode = Right(sDepartmentCode, Len(sDepartmentCode)-InStrRev(sDepartmentCode, "/"))
			End If
		Else
			sDepartmentCode="000000"
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInDepartmentCode1 + DelOtherLangFromFolder(Request("DocDepartment")) + SIT_ErrorInDepartmentCode2)
			UserChangeDocCheck=False
			Exit Function
		End If

		dsTempPR.Close
	End If

	If Request("create") = "y" Then
		sPostfix = ""
		sSearchCol = "DocID"
		sPrePrefix = ""
		'SAY 2009-02-20
		sSufix = ""

		'SAY 2009-03-12
		S_DocID = ""

		sClassDoc = "Служебные записки%"
		S_DocID = GetNewDocIDForSluzhZapiski(sClassDoc, sDepartmentRoot, Request("UserFieldText1"), "")
		S_DocIDAdd = S_DocID
	End If
	If sDepartmentRoot <> SIT_STS or (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_RASP_DOCS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1) Then
		'реализация полужесткого маршрута
		S_ListToReconcile = Replace(S_ListToReconcile, Trim(SIT_AdditionalAgrees), "")
		'Для коммерческих предложений и заявок на закупку и оплату не нужна приписка про доп. согласующих  
		If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_COM_OFFERS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) <> 1 Then
			If S_ListToReconcile <> "" Then
				S_ListToReconcile = " " + SIT_AdditionalAgrees + DeleteConstPrefixFromList(S_ListToReconcile)
			End If
		End If
	End If

	If sDepartmentRoot = SIT_SITRONICS Then
		S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
	ElseIf sDepartmentRoot = SIT_STS Then
		'rmanyushin 89142 01.04.2010 Start
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME)) > 0 Then
			If is789DivisionSTS(Session("Department")) Then
				DocSession = "Служебные записки*Office memo*Interní sdělení/На переработки 7-8-9*Overtime 7-8-9*Overtime 7-8-9"
				S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",DocSession + "/","List")+STS_DptOSSBSS+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
				'S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+STS_DptOSSBSS+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
			Else
				S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
			End if
		End If
		'rmanyushin 89142 01.04.2010 Stop

		'rmanyushin 119579 19.08.2010 Start
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_HOLIDAY)) > 0 Then
			S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
		End If
		'rmanyushin 119579 19.08.2010 End
	ElseIf sDepartmentRoot = SIT_SIB Then'СИБ
		S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSIBRU","Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
	Else 'Другие БН
		S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
	End If

	'Для всех категорий кроме заявок подставляем пользователей вместо ролей
	If (InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1) and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) <> 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_DOGOVORI_NEW)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIS)) <> 0) Then
		'Получаем список ролей
		If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) = 1 Then 'Для договоров получаем список ролей на всех языках (могут быть роли на разных языках)
			sRoleList = GetFullRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"), "", "", "")
		Else
			sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))
		End If
		'Подставляем роли в поля
		S_NameAproval = ReplaceRolesInList(S_NameAproval, sRoleList)
		S_NameResponsible = ReplaceRolesInList(S_NameResponsible, sRoleList)
		S_NameControl = ReplaceRolesInList(S_NameControl, sRoleList)
		S_ListToEdit = ReplaceRolesInList(S_ListToEdit, sRoleList)
		S_ListToView = ReplaceRolesInList(S_ListToView, sRoleList)
		S_ListToReconcile = ReplaceRolesInList(S_ListToReconcile, sRoleList)

		'Удаляем повторы в списке согласования
		S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)

		'vnik_error
		'Убираем автора документа из списка согласующих, т.к. после активации он не сможет загрузить новый документ
		'(работает пока только, если пользователь создает документ не под ролью)
		S_ListToReconcile = Replace(S_ListToReconcile,Request("DocAuthor"),"")
		'vnik_error
	End If
	'Проверка полей с пользователями
	'Для заявок отключаем проверку утверждающего, т.к. он рассчитывается
	If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 Then
		UserChangeDocCheck = CheckSingleUserField(S_NameAproval)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameAproval")+SIT_ErrorInUserField2)
			Exit Function
		End If
		'Для заявки на закупку отключаем проверку ответственного, т.к. можно указывать подразделение
		'Для заявки на оплату тоже, т.к. автоматом ставится роль из специального справочника
		UserChangeDocCheck = CheckSingleUserField(S_NameResponsible)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameResponsible")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
	UserChangeDocCheck = CheckSingleUserField(S_NameControl)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameControl")+SIT_ErrorInUserField2)
		Exit Function
	End If
	If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) = 0 Then 'В исходящих может быть любой текст
		UserChangeDocCheck = CheckMultiUserField(S_Correspondent)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("Correspondent")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
'Запрос №1 - СИБ - start
	IF InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_SLUZH_ZAPISKA)) = 0) THEN
'Запрос №1 - СИБ - end
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) = 0 Then 'В СЗ на переработки работают правила
			UserChangeDocCheck = CheckMultiUserField(S_ListToReconcile)
			If not UserChangeDocCheck Then
				Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToReconcile")+SIT_ErrorInUserField2)
				Exit Function
			End If
'Запрос №1 - СИБ - start
		End If
	END IF
'Запрос №1 - СИБ - end
	UserChangeDocCheck = CheckMultiUserField(S_ListToEdit)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToEdit")+SIT_ErrorInUserField2)
		Exit Function
	End If
	'UserChangeDocCheck = CheckMultiUserField(S_ListToView)
	'If not UserChangeDocCheck Then
	'	Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToView")+SIT_ErrorInUserField2)
	'	Exit Function
	'End If

'{ph - 20120812
	'S_AdditionalUsers = GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit"))
	S_AdditionalUsers = DeleteUserDoublesInListNew(Request("AdditionalUsers") & VbCrLf & GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit")), True)
'ph - 20120812}
	AddLogD "@@@Search For Leaders - S_AdditionalUsers: "+S_AdditionalUsers
	'vnik_payment_order
	'If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_PAYMENT_ORDER)) > 0 Then
	'	S_AdditionalUsers = S_AdditionalUsers + ReplaceRoleFromDir(SIT_Account_manager,sDepartmentRoot)       
	'End If        
	'vnik_payment_order
	S_AdditionalUsers = DeleteUserDoublesInList(S_AdditionalUsers)

	'vnik
	If S_AdditionalUsers = "" Then
		S_AdditionalUsers = "-"
	End If
	S_AdditionalUsers_Set = S_AdditionalUsers

	S_ListToReconcile_Set = S_ListToReconcile
	S_NameAproval_Set = S_NameAproval
	S_ListToEdit = S_ListToEdit
	S_NameResponsible_Set = S_NameResponsible
	S_NameControl_Set = S_NameControl
	S_ListToView_Set=S_ListToView

	If UCase(Request("UpdateDoc")) = "YES" And (Request("create") = "y"  Or Request("DocIDPrev") <> Request("DocID")) Then
		'out "123"
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		dsTemp.Open "select * from Docs where DocID='" + MakeSQLSafe(Request("DocID")) + "'", Conn 
		If Not dsTemp.EOF Then
			Session("Message") = DOCS_ALREADYEXISTS + ": " + Request("DocID")
			dsTemp.Close
			UserChangeDocCheck=False
			Exit Function
		End If
		dsTemp.Close
	End If

	'ph - 20090808 - start
	If Request("DocIDParent") <> "" Then
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		bMyDocIDParentChanged = False
		If Request("create") <> "y" Then
			'Проверяем изменился ли номер родительского документа, если идет редактирование
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocID"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				bMyDocIDParentChanged = Request("DocIDParent")<>dsTemp("DocIDParent")
			End If
			dsTemp.Close
		End If
		If bMyDocIDParentChanged Then
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocIDParent"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				'sMessage = ""
				'UserChangeDocCheck = oPayDox.IsReadAccessUser(Session("UserID"), sMessage, dsTemp, DOCS_All, DOCS_NoReadAccess, DOCS_NoAccess, USER_Department, VAR_ExtInt, VAR_AdminSecLevel, VAR_StatusActiveUser, NULL)
				UserChangeDocCheck = IsReadAccessRS(dsTemp)' IsReadAccessUser(Session("UserID"), dsTemp, NULL)
				If UserChangeDocCheck Then
					S_DocIDParent = dsTemp("DocID")
				End If
				dsTemp.Close
				If not UserChangeDocCheck Then
					Session("Message") = SIT_NoAccessToParentDoc'sMessage
					Exit Function
				End If
			End If
		End If
	End If
	'ph - 20090808 - end

'SIT_SLUZH_ZAPISKA_OBUCHENIE
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_SLUZH_ZAPISKA_OBUCHENIE)) = 1 Then
	'Создатель документа
	If UCase(Request("create")) = "Y" Then
		sDocCreator = Session("UserID")
	Else
		sDocCreator = GetUserID(Request("DocAuthor"))
	End If

	S_NameAproval=Request("DocNameAproval")
	S_NameResponsible=Request("DocNameResponsible")
	S_NameControl=Request("DocNameControl")
	S_ListToReconcile=Request("DocListToReconcile")
	S_ListToEdit=Request("DocListToEdit")
	S_ListToView=Request("DocListToView")
	S_Correspondent=Request("DocCorrespondent")

	'Определяем бизнес направление
	sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))

	' SAY 2008-07-21
	Set dsTempPR = Server.CreateObject("ADODB.Recordset")

' DmGorsky_5: меняем местами проверку: вначале ищем SIT_SITRU ("СИТРОНИКС ИТ"), потом, если не найдено, - SIT_SITRONICS ("СИТРОНИКС")
	If InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRU))=1 then ' DmGorsky
		sDepartment = SIT_SITRU ' DmGorsky
		sDepartmentCode = "" ' DmGorsky
'Запрос №1 - СИБ - start
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SIB_ROOT_DEPARTMENT))=1 then
		sDepartment = SIT_SIB
		sDepartmentCode = ""
'Запрос №1 - СИБ - end
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRONICS))=1 then
		sDepartment = SIT_SITRONICS
		sDepartmentCode = ""
	Else
		sDepartment = SIT_STS

		sSQL = "select code from Departments where Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"' or Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"/'"
		AddLogD "WAY sSQL="+sSQL
		AddLogD "WAY Conn="+Conn
		dsTempPR.Open sSQL, Conn, 3, 1, &H1

		if not dsTempPR.EOF Then
			sDepartmentCode = dsTempPR("code")
			if InStrRev(sDepartmentCode, "/") > 0 then
				sDepartmentCode = Right(sDepartmentCode, Len(sDepartmentCode)-InStrRev(sDepartmentCode, "/"))
			End If
		Else
			sDepartmentCode="000000"
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInDepartmentCode1 + DelOtherLangFromFolder(Request("DocDepartment")) + SIT_ErrorInDepartmentCode2)
			UserChangeDocCheck=False
			Exit Function
		End If

		dsTempPR.Close
	End If

	If Request("create") = "y" Then
		sPostfix = ""
		sSearchCol = "DocID"
		sPrePrefix = ""
		'SAY 2009-02-20
		sSufix = ""

		'SAY 2009-03-12
		S_DocID = ""

		sClassDoc = "Служебные записки%"
		S_DocID = GetNewDocIDForSluzhZapiski(sClassDoc, sDepartmentRoot, Request("UserFieldText1"), "")
		S_DocIDAdd = S_DocID
	End If
	If sDepartmentRoot <> SIT_STS or (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_RASP_DOCS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1) Then
		'реализация полужесткого маршрута
		S_ListToReconcile = Replace(S_ListToReconcile, Trim(SIT_AdditionalAgrees), "")
		'Для коммерческих предложений и заявок на закупку и оплату не нужна приписка про доп. согласующих  
		If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_COM_OFFERS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) <> 1 Then
			If S_ListToReconcile <> "" Then
				S_ListToReconcile = " " + SIT_AdditionalAgrees + DeleteConstPrefixFromList(S_ListToReconcile)
			End If
		End If
	End If

	If sDepartmentRoot = SIT_SITRONICS Then
		S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
	ElseIf sDepartmentRoot = SIT_STS Then
		'rmanyushin 89142 01.04.2010 Start
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME)) > 0 Then
			If is789DivisionSTS(Session("Department")) Then
				DocSession = "Служебные записки*Office memo*Interní sdělení/На переработки 7-8-9*Overtime 7-8-9*Overtime 7-8-9"
				S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",DocSession + "/","List")+STS_DptOSSBSS+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
				'S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+STS_DptOSSBSS+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
			Else
				S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
			End if
		End If
		'rmanyushin 89142 01.04.2010 Stop

		'rmanyushin 119579 19.08.2010 Start
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_HOLIDAY)) > 0 Then
			S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
		End If
		'rmanyushin 119579 19.08.2010 End
	ElseIf sDepartmentRoot = SIT_SIB Then'СИБ
		S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSIBRU","Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
	Else 'Другие БН
		S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
	End If

	'Для всех категорий кроме заявок подставляем пользователей вместо ролей
	If (InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1) and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) <> 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_DOGOVORI_NEW)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIS)) <> 0) Then
		'Получаем список ролей
		If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) = 1 Then 'Для договоров получаем список ролей на всех языках (могут быть роли на разных языках)
			sRoleList = GetFullRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"), "", "", "")
		Else
			sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))
		End If
		'Подставляем роли в поля
		S_NameAproval = ReplaceRolesInList(S_NameAproval, sRoleList)
		S_NameResponsible = ReplaceRolesInList(S_NameResponsible, sRoleList)
		S_NameControl = ReplaceRolesInList(S_NameControl, sRoleList)
		S_ListToEdit = ReplaceRolesInList(S_ListToEdit, sRoleList)
		S_ListToView = ReplaceRolesInList(S_ListToView, sRoleList)
		S_ListToReconcile = ReplaceRolesInList(S_ListToReconcile, sRoleList)

		'Удаляем повторы в списке согласования
		S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)

		'vnik_error
		'Убираем автора документа из списка согласующих, т.к. после активации он не сможет загрузить новый документ
		'(работает пока только, если пользователь создает документ не под ролью)
		S_ListToReconcile = Replace(S_ListToReconcile,Request("DocAuthor"),"")
		'vnik_error
	End If
	'Проверка полей с пользователями
	'Для заявок отключаем проверку утверждающего, т.к. он рассчитывается
	If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 Then
		UserChangeDocCheck = CheckSingleUserField(S_NameAproval)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameAproval")+SIT_ErrorInUserField2)
			Exit Function
		End If
		'Для заявки на закупку отключаем проверку ответственного, т.к. можно указывать подразделение
		'Для заявки на оплату тоже, т.к. автоматом ставится роль из специального справочника
		UserChangeDocCheck = CheckSingleUserField(S_NameResponsible)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameResponsible")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
	UserChangeDocCheck = CheckSingleUserField(S_NameControl)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameControl")+SIT_ErrorInUserField2)
		Exit Function
	End If
	If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) = 0 Then 'В исходящих может быть любой текст
		UserChangeDocCheck = CheckMultiUserField(S_Correspondent)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("Correspondent")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
'Запрос №1 - СИБ - start
	IF InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_SLUZH_ZAPISKA)) = 0) THEN
'Запрос №1 - СИБ - end
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) = 0 Then 'В СЗ на переработки работают правила
			UserChangeDocCheck = CheckMultiUserField(S_ListToReconcile)
			If not UserChangeDocCheck Then
				Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToReconcile")+SIT_ErrorInUserField2)
				Exit Function
			End If
'Запрос №1 - СИБ - start
		End If
	END IF
'Запрос №1 - СИБ - end
	UserChangeDocCheck = CheckMultiUserField(S_ListToEdit)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToEdit")+SIT_ErrorInUserField2)
		Exit Function
	End If
	'UserChangeDocCheck = CheckMultiUserField(S_ListToView)
	'If not UserChangeDocCheck Then
	'	Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToView")+SIT_ErrorInUserField2)
	'	Exit Function
	'End If

'{ph - 20120812
	'S_AdditionalUsers = GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit"))
	S_AdditionalUsers = DeleteUserDoublesInListNew(Request("AdditionalUsers") & VbCrLf & GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit")), True)
'ph - 20120812}
	AddLogD "@@@Search For Leaders - S_AdditionalUsers: "+S_AdditionalUsers
	'vnik_payment_order
	'If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_PAYMENT_ORDER)) > 0 Then
	'	S_AdditionalUsers = S_AdditionalUsers + ReplaceRoleFromDir(SIT_Account_manager,sDepartmentRoot)       
	'End If        
	'vnik_payment_order
	S_AdditionalUsers = DeleteUserDoublesInList(S_AdditionalUsers)

	'vnik
	If S_AdditionalUsers = "" Then
		S_AdditionalUsers = "-"
	End If
	S_AdditionalUsers_Set = S_AdditionalUsers

	S_ListToReconcile_Set = S_ListToReconcile
	S_NameAproval_Set = S_NameAproval
	S_ListToEdit = S_ListToEdit
	S_NameResponsible_Set = S_NameResponsible
	S_NameControl_Set = S_NameControl
	S_ListToView_Set=S_ListToView

	If UCase(Request("UpdateDoc")) = "YES" And (Request("create") = "y"  Or Request("DocIDPrev") <> Request("DocID")) Then
		'out "123"
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		dsTemp.Open "select * from Docs where DocID='" + MakeSQLSafe(Request("DocID")) + "'", Conn 
		If Not dsTemp.EOF Then
			Session("Message") = DOCS_ALREADYEXISTS + ": " + Request("DocID")
			dsTemp.Close
			UserChangeDocCheck=False
			Exit Function
		End If
		dsTemp.Close
	End If

	'ph - 20090808 - start
	If Request("DocIDParent") <> "" Then
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		bMyDocIDParentChanged = False
		If Request("create") <> "y" Then
			'Проверяем изменился ли номер родительского документа, если идет редактирование
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocID"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				bMyDocIDParentChanged = Request("DocIDParent")<>dsTemp("DocIDParent")
			End If
			dsTemp.Close
		End If
		If bMyDocIDParentChanged Then
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocIDParent"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				'sMessage = ""
				'UserChangeDocCheck = oPayDox.IsReadAccessUser(Session("UserID"), sMessage, dsTemp, DOCS_All, DOCS_NoReadAccess, DOCS_NoAccess, USER_Department, VAR_ExtInt, VAR_AdminSecLevel, VAR_StatusActiveUser, NULL)
				UserChangeDocCheck = IsReadAccessRS(dsTemp)' IsReadAccessUser(Session("UserID"), dsTemp, NULL)
				If UserChangeDocCheck Then
					S_DocIDParent = dsTemp("DocID")
				End If
				dsTemp.Close
				If not UserChangeDocCheck Then
					Session("Message") = SIT_NoAccessToParentDoc'sMessage
					Exit Function
				End If
			End If
		End If
	End If
	'ph - 20090808 - end

'SIT_SLUZH_ZAPISKA_PERSONAL
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_SLUZH_ZAPISKA_PERSONAL)) = 1 Then
	'Создатель документа
	If UCase(Request("create")) = "Y" Then
		sDocCreator = Session("UserID")
	Else
		sDocCreator = GetUserID(Request("DocAuthor"))
	End If

	S_NameAproval=Request("DocNameAproval")
	S_NameResponsible=Request("DocNameResponsible")
	S_NameControl=Request("DocNameControl")
	S_ListToReconcile=Request("DocListToReconcile")
	S_ListToEdit=Request("DocListToEdit")
	S_ListToView=Request("DocListToView")
	S_Correspondent=Request("DocCorrespondent")

	'Определяем бизнес направление
	sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))

	' SAY 2008-07-21
	Set dsTempPR = Server.CreateObject("ADODB.Recordset")

' DmGorsky_5: меняем местами проверку: вначале ищем SIT_SITRU ("СИТРОНИКС ИТ"), потом, если не найдено, - SIT_SITRONICS ("СИТРОНИКС")
	If InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRU))=1 then ' DmGorsky
		sDepartment = SIT_SITRU ' DmGorsky
		sDepartmentCode = "" ' DmGorsky
'Запрос №1 - СИБ - start
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SIB_ROOT_DEPARTMENT))=1 then
		sDepartment = SIT_SIB
		sDepartmentCode = ""
'Запрос №1 - СИБ - end
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRONICS))=1 then
		sDepartment = SIT_SITRONICS
		sDepartmentCode = ""
	Else
		sDepartment = SIT_STS

		sSQL = "select code from Departments where Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"' or Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"/'"
		AddLogD "WAY sSQL="+sSQL
		AddLogD "WAY Conn="+Conn
		dsTempPR.Open sSQL, Conn, 3, 1, &H1

		if not dsTempPR.EOF Then
			sDepartmentCode = dsTempPR("code")
			if InStrRev(sDepartmentCode, "/") > 0 then
				sDepartmentCode = Right(sDepartmentCode, Len(sDepartmentCode)-InStrRev(sDepartmentCode, "/"))
			End If
		Else
			sDepartmentCode="000000"
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInDepartmentCode1 + DelOtherLangFromFolder(Request("DocDepartment")) + SIT_ErrorInDepartmentCode2)
			UserChangeDocCheck=False
			Exit Function
		End If

		dsTempPR.Close
	End If

	If Request("create") = "y" Then
		sPostfix = ""
		sSearchCol = "DocID"
		sPrePrefix = ""
		'SAY 2009-02-20
		sSufix = ""

		'SAY 2009-03-12
		S_DocID = ""

		sClassDoc = "Служебные записки%"
		S_DocID = GetNewDocIDForSluzhZapiski(sClassDoc, sDepartmentRoot, Request("UserFieldText1"), "")
		S_DocIDAdd = S_DocID
	End If
	If sDepartmentRoot <> SIT_STS or (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_RASP_DOCS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1) Then
		'реализация полужесткого маршрута
		S_ListToReconcile = Replace(S_ListToReconcile, Trim(SIT_AdditionalAgrees), "")
		'Для коммерческих предложений и заявок на закупку и оплату не нужна приписка про доп. согласующих  
		If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_COM_OFFERS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) <> 1 Then
			If S_ListToReconcile <> "" Then
				S_ListToReconcile = " " + SIT_AdditionalAgrees + DeleteConstPrefixFromList(S_ListToReconcile)
			End If
		End If
	End If

	If sDepartmentRoot = SIT_SITRONICS Then
		S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
	ElseIf sDepartmentRoot = SIT_STS Then
		'rmanyushin 89142 01.04.2010 Start
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME)) > 0 Then
			If is789DivisionSTS(Session("Department")) Then
				DocSession = "Служебные записки*Office memo*Interní sdělení/На переработки 7-8-9*Overtime 7-8-9*Overtime 7-8-9"
				S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",DocSession + "/","List")+STS_DptOSSBSS+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
				'S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+STS_DptOSSBSS+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
			Else
				S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
			End if
		End If
		'rmanyushin 89142 01.04.2010 Stop

		'rmanyushin 119579 19.08.2010 Start
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_HOLIDAY)) > 0 Then
			S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
		End If
		'rmanyushin 119579 19.08.2010 End
	ElseIf sDepartmentRoot = SIT_SIB Then'СИБ
		S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSIBRU","Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
	Else 'Другие БН
		S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
	End If

	'Для всех категорий кроме заявок подставляем пользователей вместо ролей
	If (InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1) and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) <> 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_DOGOVORI_NEW)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIS)) <> 0) Then
		'Получаем список ролей
		If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) = 1 Then 'Для договоров получаем список ролей на всех языках (могут быть роли на разных языках)
			sRoleList = GetFullRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"), "", "", "")
		Else
			sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))
		End If
		'Подставляем роли в поля
		S_NameAproval = ReplaceRolesInList(S_NameAproval, sRoleList)
		S_NameResponsible = ReplaceRolesInList(S_NameResponsible, sRoleList)
		S_NameControl = ReplaceRolesInList(S_NameControl, sRoleList)
		S_ListToEdit = ReplaceRolesInList(S_ListToEdit, sRoleList)
		S_ListToView = ReplaceRolesInList(S_ListToView, sRoleList)
		S_ListToReconcile = ReplaceRolesInList(S_ListToReconcile, sRoleList)

		'Удаляем повторы в списке согласования
		S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)

		'vnik_error
		'Убираем автора документа из списка согласующих, т.к. после активации он не сможет загрузить новый документ
		'(работает пока только, если пользователь создает документ не под ролью)
		S_ListToReconcile = Replace(S_ListToReconcile,Request("DocAuthor"),"")
		'vnik_error
	End If
	'Проверка полей с пользователями
	'Для заявок отключаем проверку утверждающего, т.к. он рассчитывается
	If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 Then
		UserChangeDocCheck = CheckSingleUserField(S_NameAproval)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameAproval")+SIT_ErrorInUserField2)
			Exit Function
		End If
		'Для заявки на закупку отключаем проверку ответственного, т.к. можно указывать подразделение
		'Для заявки на оплату тоже, т.к. автоматом ставится роль из специального справочника
		UserChangeDocCheck = CheckSingleUserField(S_NameResponsible)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameResponsible")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
	UserChangeDocCheck = CheckSingleUserField(S_NameControl)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameControl")+SIT_ErrorInUserField2)
		Exit Function
	End If
	If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) = 0 Then 'В исходящих может быть любой текст
		UserChangeDocCheck = CheckMultiUserField(S_Correspondent)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("Correspondent")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
'Запрос №1 - СИБ - start
	IF InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_SLUZH_ZAPISKA)) = 0) THEN
'Запрос №1 - СИБ - end
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) = 0 Then 'В СЗ на переработки работают правила
			UserChangeDocCheck = CheckMultiUserField(S_ListToReconcile)
			If not UserChangeDocCheck Then
				Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToReconcile")+SIT_ErrorInUserField2)
				Exit Function
			End If
'Запрос №1 - СИБ - start
		End If
	END IF
'Запрос №1 - СИБ - end
	UserChangeDocCheck = CheckMultiUserField(S_ListToEdit)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToEdit")+SIT_ErrorInUserField2)
		Exit Function
	End If
	'UserChangeDocCheck = CheckMultiUserField(S_ListToView)
	'If not UserChangeDocCheck Then
	'	Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToView")+SIT_ErrorInUserField2)
	'	Exit Function
	'End If

'{ph - 20120812
	'S_AdditionalUsers = GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit"))
	S_AdditionalUsers = DeleteUserDoublesInListNew(Request("AdditionalUsers") & VbCrLf & GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit")), True)
'ph - 20120812}
	AddLogD "@@@Search For Leaders - S_AdditionalUsers: "+S_AdditionalUsers
	'vnik_payment_order
	'If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_PAYMENT_ORDER)) > 0 Then
	'	S_AdditionalUsers = S_AdditionalUsers + ReplaceRoleFromDir(SIT_Account_manager,sDepartmentRoot)       
	'End If        
	'vnik_payment_order
	S_AdditionalUsers = DeleteUserDoublesInList(S_AdditionalUsers)

	'vnik
	If S_AdditionalUsers = "" Then
		S_AdditionalUsers = "-"
	End If
	S_AdditionalUsers_Set = S_AdditionalUsers

	S_ListToReconcile_Set = S_ListToReconcile
	S_NameAproval_Set = S_NameAproval
	S_ListToEdit = S_ListToEdit
	S_NameResponsible_Set = S_NameResponsible
	S_NameControl_Set = S_NameControl
	S_ListToView_Set=S_ListToView

	If UCase(Request("UpdateDoc")) = "YES" And (Request("create") = "y"  Or Request("DocIDPrev") <> Request("DocID")) Then
		'out "123"
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		dsTemp.Open "select * from Docs where DocID='" + MakeSQLSafe(Request("DocID")) + "'", Conn 
		If Not dsTemp.EOF Then
			Session("Message") = DOCS_ALREADYEXISTS + ": " + Request("DocID")
			dsTemp.Close
			UserChangeDocCheck=False
			Exit Function
		End If
		dsTemp.Close
	End If

	'ph - 20090808 - start
	If Request("DocIDParent") <> "" Then
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		bMyDocIDParentChanged = False
		If Request("create") <> "y" Then
			'Проверяем изменился ли номер родительского документа, если идет редактирование
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocID"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				bMyDocIDParentChanged = Request("DocIDParent")<>dsTemp("DocIDParent")
			End If
			dsTemp.Close
		End If
		If bMyDocIDParentChanged Then
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocIDParent"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				'sMessage = ""
				'UserChangeDocCheck = oPayDox.IsReadAccessUser(Session("UserID"), sMessage, dsTemp, DOCS_All, DOCS_NoReadAccess, DOCS_NoAccess, USER_Department, VAR_ExtInt, VAR_AdminSecLevel, VAR_StatusActiveUser, NULL)
				UserChangeDocCheck = IsReadAccessRS(dsTemp)' IsReadAccessUser(Session("UserID"), dsTemp, NULL)
				If UserChangeDocCheck Then
					S_DocIDParent = dsTemp("DocID")
				End If
				dsTemp.Close
				If not UserChangeDocCheck Then
					Session("Message") = SIT_NoAccessToParentDoc'sMessage
					Exit Function
				End If
			End If
		End If
	End If
	'ph - 20090808 - end

'STS_SLUZH_ZAPISKA_OVERTIME2
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) = 1 Then
	'Запрос №46 - СТС - start
	'Запрет создания документов в категории
	If UCase(Request("create")) = "Y" Then
		If InStr(UCase(Request("DocDepartment")), UCase(SIT_STS_ROOT_DEPARTMENT)) = 1 Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_CannotCreateDocInThisCategory)
			UserChangeDocCheck = False
			Exit Function
		End If
	End If
	'Запрос №46 - СТС - end
	'Создатель документа
	If UCase(Request("create")) = "Y" Then
		sDocCreator = Session("UserID")
	Else
		sDocCreator = GetUserID(Request("DocAuthor"))
	End If

	S_NameAproval=Request("DocNameAproval")
	S_NameResponsible=Request("DocNameResponsible")
	S_NameControl=Request("DocNameControl")
	S_ListToReconcile=Request("DocListToReconcile")
	S_ListToEdit=Request("DocListToEdit")
	S_ListToView=Request("DocListToView")
	S_Correspondent=Request("DocCorrespondent")

	'Определяем бизнес направление
	sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))

	' SAY 2008-07-21
	Set dsTempPR = Server.CreateObject("ADODB.Recordset")

' DmGorsky_5: меняем местами проверку: вначале ищем SIT_SITRU ("СИТРОНИКС ИТ"), потом, если не найдено, - SIT_SITRONICS ("СИТРОНИКС")
	If InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRU))=1 then ' DmGorsky
		sDepartment = SIT_SITRU ' DmGorsky
		sDepartmentCode = "" ' DmGorsky
'Запрос №1 - СИБ - start
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SIB_ROOT_DEPARTMENT))=1 then
		sDepartment = SIT_SIB
		sDepartmentCode = ""
'Запрос №1 - СИБ - end
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRONICS))=1 then
		sDepartment = SIT_SITRONICS
		sDepartmentCode = ""
	Else
		sDepartment = SIT_STS

		sSQL = "select code from Departments where Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"' or Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"/'"
		AddLogD "WAY sSQL="+sSQL
		AddLogD "WAY Conn="+Conn
		dsTempPR.Open sSQL, Conn, 3, 1, &H1

		if not dsTempPR.EOF Then
			sDepartmentCode = dsTempPR("code")
			if InStrRev(sDepartmentCode, "/") > 0 then
				sDepartmentCode = Right(sDepartmentCode, Len(sDepartmentCode)-InStrRev(sDepartmentCode, "/"))
			End If
		Else
			sDepartmentCode="000000"
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInDepartmentCode1 + DelOtherLangFromFolder(Request("DocDepartment")) + SIT_ErrorInDepartmentCode2)
			UserChangeDocCheck=False
			Exit Function
		End If

		dsTempPR.Close
	End If

	If Request("create") = "y" Then
		sPostfix = ""
		sSearchCol = "DocID"
		sPrePrefix = ""
		'SAY 2009-02-20
		sSufix = ""

		'SAY 2009-03-12
		S_DocID = ""

		sClassDoc = "Служебные записки%"
		S_DocID = GetNewDocIDForSluzhZapiski(sClassDoc, sDepartmentRoot, Request("UserFieldText1"), "")
		S_DocIDAdd = S_DocID
	End If
	If sDepartmentRoot <> SIT_STS or (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_RASP_DOCS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1) Then
		'реализация полужесткого маршрута
		S_ListToReconcile = Replace(S_ListToReconcile, Trim(SIT_AdditionalAgrees), "")
		'Для коммерческих предложений и заявок на закупку и оплату не нужна приписка про доп. согласующих  
		If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_COM_OFFERS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) <> 1 Then
			If S_ListToReconcile <> "" Then
				S_ListToReconcile = " " + SIT_AdditionalAgrees + DeleteConstPrefixFromList(S_ListToReconcile)
			End If
		End If
	End If

	If sDepartmentRoot = SIT_SITRONICS Then
		S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
	ElseIf sDepartmentRoot = SIT_STS Then
		'rmanyushin 89142 01.04.2010 Start
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME)) > 0 Then
			If is789DivisionSTS(Session("Department")) Then
				DocSession = "Служебные записки*Office memo*Interní sdělení/На переработки 7-8-9*Overtime 7-8-9*Overtime 7-8-9"
				S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",DocSession + "/","List")+STS_DptOSSBSS+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
				'S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+STS_DptOSSBSS+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
			Else
				S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
			End if
		End If
		'rmanyushin 89142 01.04.2010 Stop

		'rmanyushin 119579 19.08.2010 Start
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_HOLIDAY)) > 0 Then
			S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
		End If
		'rmanyushin 119579 19.08.2010 End
	ElseIf sDepartmentRoot = SIT_SIB Then'СИБ
		S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSIBRU","Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
	Else 'Другие БН
		S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
	End If

	'Для всех категорий кроме заявок подставляем пользователей вместо ролей
	If (InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1) and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) <> 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_DOGOVORI_NEW)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIS)) <> 0) Then
		'Получаем список ролей
		If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) = 1 Then 'Для договоров получаем список ролей на всех языках (могут быть роли на разных языках)
			sRoleList = GetFullRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"), "", "", "")
		Else
			sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))
		End If
		'Подставляем роли в поля
		S_NameAproval = ReplaceRolesInList(S_NameAproval, sRoleList)
		S_NameResponsible = ReplaceRolesInList(S_NameResponsible, sRoleList)
		S_NameControl = ReplaceRolesInList(S_NameControl, sRoleList)
		S_ListToEdit = ReplaceRolesInList(S_ListToEdit, sRoleList)
		S_ListToView = ReplaceRolesInList(S_ListToView, sRoleList)
		S_ListToReconcile = ReplaceRolesInList(S_ListToReconcile, sRoleList)

		'Удаляем повторы в списке согласования
		S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)

		'vnik_error
		'Убираем автора документа из списка согласующих, т.к. после активации он не сможет загрузить новый документ
		'(работает пока только, если пользователь создает документ не под ролью)
		S_ListToReconcile = Replace(S_ListToReconcile,Request("DocAuthor"),"")
		'vnik_error
	End If
	'Проверка полей с пользователями
	'Для заявок отключаем проверку утверждающего, т.к. он рассчитывается
	If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 Then
		UserChangeDocCheck = CheckSingleUserField(S_NameAproval)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameAproval")+SIT_ErrorInUserField2)
			Exit Function
		End If
		'Для заявки на закупку отключаем проверку ответственного, т.к. можно указывать подразделение
		'Для заявки на оплату тоже, т.к. автоматом ставится роль из специального справочника
		UserChangeDocCheck = CheckSingleUserField(S_NameResponsible)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameResponsible")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
	UserChangeDocCheck = CheckSingleUserField(S_NameControl)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameControl")+SIT_ErrorInUserField2)
		Exit Function
	End If
	If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) = 0 Then 'В исходящих может быть любой текст
		UserChangeDocCheck = CheckMultiUserField(S_Correspondent)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("Correspondent")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
'Запрос №1 - СИБ - start
	IF InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_SLUZH_ZAPISKA)) = 0) THEN
'Запрос №1 - СИБ - end
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) = 0 Then 'В СЗ на переработки работают правила
			UserChangeDocCheck = CheckMultiUserField(S_ListToReconcile)
			If not UserChangeDocCheck Then
				Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToReconcile")+SIT_ErrorInUserField2)
				Exit Function
			End If
'Запрос №1 - СИБ - start
		End If
	END IF
'Запрос №1 - СИБ - end
	UserChangeDocCheck = CheckMultiUserField(S_ListToEdit)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToEdit")+SIT_ErrorInUserField2)
		Exit Function
	End If
	'UserChangeDocCheck = CheckMultiUserField(S_ListToView)
	'If not UserChangeDocCheck Then
	'	Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToView")+SIT_ErrorInUserField2)
	'	Exit Function
	'End If

'{ph - 20120812
	'S_AdditionalUsers = GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit"))
	S_AdditionalUsers = DeleteUserDoublesInListNew(Request("AdditionalUsers") & VbCrLf & GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit")), True)
'ph - 20120812}
	AddLogD "@@@Search For Leaders - S_AdditionalUsers: "+S_AdditionalUsers
	'vnik_payment_order
	'If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_PAYMENT_ORDER)) > 0 Then
	'	S_AdditionalUsers = S_AdditionalUsers + ReplaceRoleFromDir(SIT_Account_manager,sDepartmentRoot)       
	'End If        
	'vnik_payment_order
	S_AdditionalUsers = DeleteUserDoublesInList(S_AdditionalUsers)

	'vnik
	If S_AdditionalUsers = "" Then
		S_AdditionalUsers = "-"
	End If
	S_AdditionalUsers_Set = S_AdditionalUsers

	S_ListToReconcile_Set = S_ListToReconcile
	S_NameAproval_Set = S_NameAproval
	S_ListToEdit = S_ListToEdit
	S_NameResponsible_Set = S_NameResponsible
	S_NameControl_Set = S_NameControl
	S_ListToView_Set=S_ListToView

	If UCase(Request("UpdateDoc")) = "YES" And (Request("create") = "y"  Or Request("DocIDPrev") <> Request("DocID")) Then
		'out "123"
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		dsTemp.Open "select * from Docs where DocID='" + MakeSQLSafe(Request("DocID")) + "'", Conn, 3, 1, &H1
		If Not dsTemp.EOF Then
			Session("Message") = DOCS_ALREADYEXISTS + ": " + Request("DocID")
			dsTemp.Close
			UserChangeDocCheck=False
			Exit Function
		End If
		dsTemp.Close
	End If

	'rmanyushin 136964 08.11.2010 Start
	'Запрос №1 - СИБ - start - Проверку номера проекта для СИБ не делаем
	If InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) <> 1 Then
		Set dsTemp1 = Server.CreateObject("ADODB.Recordset")
		sSQL = "select * from ProjectList where ProjectID = "+sUnicodeSymbol+"'" + Request("UserFieldText3") + "'"
		AddLogD "@@@GetProjectName SQL: "+sSQL
		dsTemp1.Open sSQL, Conn, 3, 1, &H1
		If Not dsTemp1.EOF Then
			S_UserFieldText4 = Trim(CStr(dsTemp1("ProjectCode")))+" "+Trim(CStr(dsTemp1("ProjectName")))
			S_UserFieldText4_Set = S_UserFieldText4
			If IsNull(dsTemp1("ProjectManagerUser")) Then
				ProjectManager = ""
			Else
				ProjectManager = Trim(CStr(dsTemp1("ProjectManagerUser")))
			End If
		Else 'Код проекта не соответствует справочнику
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInProjectNumber)
			UserChangeDocCheck = False
			dsTemp1.Close
			Exit Function
		End If
		dsTemp1.Close
	End If

	'Запрос №1 - СИБ - start - Проверку номера проекта для СИБ не делаем
	If InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 Then
		S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSIBRU","Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
	Else
'{Запрос №50 - СТС
		'GetReconciliationByRules(parDocID, parClassDoc, parAmount, parChartOfAccounts, parCostCenter, parProjectCode, parBusinessUnit, parPartnerName, parKindOfPayment, parInitiator, parProjectManager, parOvertimeRequester, parIncomeExpenceContract, parContranctInSTS, parCurrency, parContractType, parOvertimeFuncLeaders, ByRef parListToReconcile, ByRef parNameAproval, ByRef parNameResponsible, ByRef parCorrespondent, ByRef parListToView, ByRef parRegistrar, ByRef parNameControl, ByRef parListToEdit)
		'GetReconciliationByRules iif(Trim(Request("DocID")) = "", S_DocID, Request("DocID")), Session("CurrentClassDoc"), Null, Null, sCostCenter, GetCodeFromCode_NameString(Request("UserFieldText3")), GetCodeFromCode_NameString(Request("BusinessUnit")), Null, "", iif(UCase(Request("create")) = "Y", Session("UserID"), Request("DocAuthor")), ProjectManager, Request("DocNameResponsible"), NULL, NULL, NULL, NULL, NULL, S_ListToReconcile, S_NameAproval, Null, Null, S_ListToView, Null, Null, Null
		par_DocID = iif(Trim(Request("DocID")) = "", S_DocID, Request("DocID"))
		par_ClassDoc = Session("CurrentClassDoc")
		par_Amount = NULL
		par_ChartOfAccounts = NULL
		par_CostCenter = sCostCenter
		par_ProjectCode = GetCodeFromCode_NameString(Request("UserFieldText3"))
		par_BusinessUnit = GetCodeFromCode_NameString(Request("BusinessUnit"))
		par_PartnerName = NULL
		par_KindOfPayment = ""
		par_Initiator = iif(UCase(Request("create")) = "Y", Session("UserID"), Request("DocAuthor"))
		par_ProjectManager = ProjectManager
		par_OvertimeRequester = Request("DocNameResponsible")
		par_IncomeExpenceContract = NULL
		par_ContranctInSTS = NULL
		par_Currency = NULL
		par_ContractType = NULL
		par_OvertimeFuncLeaders = NULL
		par_TypeOfDocument = ""
		par_FunctionArea = ""

		GetReconciliationByRules par_DocID, par_ClassDoc, par_Amount, par_ChartOfAccounts, par_CostCenter, par_ProjectCode, par_BusinessUnit, par_PartnerName, par_KindOfPayment, par_Initiator, par_ProjectManager, par_OvertimeRequester, par_IncomeExpenceContract, par_ContranctInSTS, par_Currency, par_ContractType, par_OvertimeFuncLeaders, par_TypeOfDocument, par_FunctionArea, _
			S_ListToReconcile, S_NameAproval, Null, Null, S_ListToView, Null, Null, Null
		'	par_ListToReconcile, par_NameAproval, par_NameResponsible, par_Correspondent, par_ListToView, par_Registrar, par_NameControl, par_ListToEdit
'Запрос №50 - СТС}
	End If
	S_ListToView_Set = S_ListToView
	'If not IsAdmin() Then
		S_ListToReconcile_Set = S_ListToReconcile
		S_NameAproval_Set = S_NameAproval
	'End If
	'rmanyushin 136964 08.11.2010 End

	'ph - 20090808 - start
	If Request("DocIDParent") <> "" Then
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		bMyDocIDParentChanged = False
		If Request("create") <> "y" Then
			'Проверяем изменился ли номер родительского документа, если идет редактирование
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocID"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				bMyDocIDParentChanged = Request("DocIDParent")<>dsTemp("DocIDParent")
			End If
			dsTemp.Close
		End If
		If bMyDocIDParentChanged Then
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocIDParent"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				'sMessage = ""
				'UserChangeDocCheck = oPayDox.IsReadAccessUser(Session("UserID"), sMessage, dsTemp, DOCS_All, DOCS_NoReadAccess, DOCS_NoAccess, USER_Department, VAR_ExtInt, VAR_AdminSecLevel, VAR_StatusActiveUser, NULL)
				UserChangeDocCheck = IsReadAccessRS(dsTemp)' IsReadAccessUser(Session("UserID"), dsTemp, NULL)
				If UserChangeDocCheck Then
					S_DocIDParent = dsTemp("DocID")
				End If
				dsTemp.Close
				If not UserChangeDocCheck Then
					Session("Message") = SIT_NoAccessToParentDoc'sMessage
					Exit Function
				End If
			Else
				dsTemp.Close
			End If
		End If
	End If
	'ph - 20090808 - end

'STS_SLUZH_ZAPISKA_OVERTIME_PLAN
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(STS_SLUZH_ZAPISKA_OVERTIME_PLAN)) = 1 Then
	S_NameAproval=Request("DocNameAproval")
	S_NameResponsible=Request("DocNameResponsible")
	S_NameControl=Request("DocNameControl")
	S_ListToReconcile=Request("DocListToReconcile")
	S_ListToEdit=Request("DocListToEdit")
	S_ListToView=Request("DocListToView")
	S_Correspondent=Request("DocCorrespondent")

	'Определяем бизнес направление
	sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))

	' SAY 2008-07-21
	Set dsTempPR = Server.CreateObject("ADODB.Recordset")

	If InStr(UCase(Request("DocDepartment")),  UCase(SIT_STS)) = 1 Then
		sDepartment = SIT_STS

		sSQL = "select code from Departments where Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"' or Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"/'"
		AddLogD "WAY sSQL="+sSQL
		AddLogD "WAY Conn="+Conn
		dsTempPR.Open sSQL, Conn, 3, 1, &H1

		if not dsTempPR.EOF Then
			sDepartmentCode = dsTempPR("code")
			if InStrRev(sDepartmentCode, "/") > 0 then
				sDepartmentCode = Right(sDepartmentCode, Len(sDepartmentCode)-InStrRev(sDepartmentCode, "/"))
			End If
		Else
			sDepartmentCode="000000"
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInDepartmentCode1 + DelOtherLangFromFolder(Request("DocDepartment")) + SIT_ErrorInDepartmentCode2)
			UserChangeDocCheck=False
			Exit Function
		End If

		dsTempPR.Close
	End If

	'Генерация номера
	If Request("create") = "y" Then
		sClassDoc = "Служебные записки%"
		S_DocID = GetNewDocIDForSluzhZapiski(sClassDoc, sDepartmentRoot, Request("UserFieldText1"), "")
		S_DocIDAdd = S_DocID
	End If

	'Проверка полей с пользователями
	UserChangeDocCheck = CheckMultiUserField(S_Correspondent)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("Correspondent")+SIT_ErrorInUserField2)
		Exit Function
	End If

	'Проверяем номера проектов
	S_UserFieldText3 = Request("UserFieldText3")
	ProjectManagers = ""
	If not CheckAndEnhanceProjectList(S_UserFieldText3, ProjectManagers, Conn) Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInProjectNumbers & S_UserFieldText3)
		'Возвращаем исходное значение в S_UserFieldText3
		S_UserFieldText3 = Request("UserFieldText3")
		UserChangeDocCheck = False
		Exit Function
	End If

'{Запрос №50 - СТС
	'GetReconciliationByRules(parDocID, parClassDoc, parAmount, parChartOfAccounts, parCostCenter, parProjectCode, parBusinessUnit, parPartnerName, parKindOfPayment, parInitiator, parProjectManager, parOvertimeRequester, parIncomeExpenceContract, parContranctInSTS, parCurrency, parContractType, parOvertimeFuncLeaders, ByRef parListToReconcile, ByRef parNameAproval, ByRef parNameResponsible, ByRef parCorrespondent, ByRef parListToView, ByRef parRegistrar, ByRef parNameControl, ByRef parListToEdit)
	'GetReconciliationByRules iif(Trim(Request("DocID")) = "", S_DocID, Request("DocID")), Session("CurrentClassDoc"), Null, Null, sCostCenter, GetCodeFromCode_NameString(Request("UserFieldText3")), GetCodeFromCode_NameString(Request("BusinessUnit")), Null, "", iif(UCase(Request("create")) = "Y", Session("UserID"), Request("DocAuthor")), Replace(ProjectManagers, VbCrLf, " "), Request("DocNameResponsible"), NULL, NULL, NULL, NULL, Replace(Request("DocCorrespondent"), VbCrLf, " "), S_ListToReconcile, S_NameAproval, Null, Null, S_ListToView, Null, Null, Null
	par_DocID = iif(Trim(Request("DocID")) = "", S_DocID, Request("DocID"))
	par_ClassDoc = Session("CurrentClassDoc")
	par_Amount = NULL
	par_ChartOfAccounts = NULL
	par_CostCenter = sCostCenter
	par_ProjectCode = GetCodeFromCode_NameString(Request("UserFieldText3"))
	par_BusinessUnit = GetCodeFromCode_NameString(Request("BusinessUnit"))
	par_PartnerName = NULL
	par_KindOfPayment = ""
	par_Initiator = iif(UCase(Request("create")) = "Y", Session("UserID"), Request("DocAuthor"))
	par_ProjectManager = Replace(ProjectManagers, VbCrLf, " ")
	par_OvertimeRequester = Request("DocNameResponsible")
	par_IncomeExpenceContract = NULL
	par_ContranctInSTS = NULL
	par_Currency = NULL
	par_ContractType = NULL
	par_OvertimeFuncLeaders = Replace(Request("DocCorrespondent"), VbCrLf, " ")
	par_TypeOfDocument = ""
	par_FunctionArea = ""

	GetReconciliationByRules par_DocID, par_ClassDoc, par_Amount, par_ChartOfAccounts, par_CostCenter, par_ProjectCode, par_BusinessUnit, par_PartnerName, par_KindOfPayment, par_Initiator, par_ProjectManager, par_OvertimeRequester, par_IncomeExpenceContract, par_ContranctInSTS, par_Currency, par_ContractType, par_OvertimeFuncLeaders, par_TypeOfDocument, par_FunctionArea, _
		S_ListToReconcile, S_NameAproval, Null, Null, S_ListToView, Null, Null, Null
	'	par_ListToReconcile, par_NameAproval, par_NameResponsible, par_Correspondent, par_ListToView, par_Registrar, par_NameControl, par_ListToEdit
'Запрос №50 - СТС}

	'В доп. пользователей заносим всех руководителей исполнителя и согласующих, чтобы они имели доступ
'{ph - 20120812
	'S_AdditionalUsers = GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit"))
	S_AdditionalUsers = DeleteUserDoublesInListNew(Request("AdditionalUsers") & VbCrLf & GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit")), True)
'ph - 20120812}
	AddLogD "@@@Search For Leaders - S_AdditionalUsers: "+S_AdditionalUsers
	S_AdditionalUsers = DeleteUserDoublesInList(S_AdditionalUsers)
	If S_AdditionalUsers = "" Then
		S_AdditionalUsers = " "
	End If
	S_AdditionalUsers_Set = S_AdditionalUsers

'STS_SLUZH_ZAPISKA_OVERTIME_FACT
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(STS_SLUZH_ZAPISKA_OVERTIME_FACT)) = 1 Then
	S_NameAproval=Request("DocNameAproval")
	S_NameResponsible=Request("DocNameResponsible")
	S_NameControl=Request("DocNameControl")
	S_ListToReconcile=Request("DocListToReconcile")
	S_ListToEdit=Request("DocListToEdit")
	S_ListToView=Request("DocListToView")
	S_Correspondent=Request("DocCorrespondent")

	'Определяем бизнес направление
	sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))

	' SAY 2008-07-21
	Set dsTempPR = Server.CreateObject("ADODB.Recordset")

	If InStr(UCase(Request("DocDepartment")),  UCase(SIT_STS)) = 1 Then
		sDepartment = SIT_STS

		sSQL = "select code from Departments where Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"' or Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"/'"
		AddLogD "WAY sSQL="+sSQL
		AddLogD "WAY Conn="+Conn
		dsTempPR.Open sSQL, Conn, 3, 1, &H1

		if not dsTempPR.EOF Then
			sDepartmentCode = dsTempPR("code")
			if InStrRev(sDepartmentCode, "/") > 0 then
				sDepartmentCode = Right(sDepartmentCode, Len(sDepartmentCode)-InStrRev(sDepartmentCode, "/"))
			End If
		Else
			sDepartmentCode="000000"
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInDepartmentCode1 + DelOtherLangFromFolder(Request("DocDepartment")) + SIT_ErrorInDepartmentCode2)
			UserChangeDocCheck=False
			Exit Function
		End If

		dsTempPR.Close
	End If

	'Проверки родительского документа
	If Trim(Request("DocIDParent")) = "" Then
		'Родительский документ не указан
		UserChangeDocCheck = False
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInParentDocNumber)
		Exit Function
	Else
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		sSQL = "select * from Docs where DocID=" & sUnicodeSymbol & "'" & MakeSQLSafeSimple(Request("DocIDParent")) & "'"
		dsTemp.Open sSQL, Conn, 3, 1, &H1
		If dsTemp.EOF Then
			'Родительский документ не найден
			UserChangeDocCheck = False
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInParentDocNumber)
			dsTemp.Close
			Exit Function
		Else
			'Проверяем доступ к родительскому документу
			UserChangeDocCheck = IsReadAccessRS(dsTemp) 'IsReadAccessUser(Session("UserID"), dsTemp, NULL)
			If UserChangeDocCheck Then
				S_DocIDParent = dsTemp("DocID")
			End If
			If not UserChangeDocCheck Then
				Session("Message") = AddNewLineToMessage(Session("Message"), SIT_NoAccessToParentDoc)
				dsTemp.Close
				Exit Function
			End If
			'Родительский документ - не плановая служебка на переработки
			If InStr(MyCStr(dsTemp("ClassDoc")), STS_SLUZH_ZAPISKA_OVERTIME_PLAN) <> 1 Then
				Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInParentDocNumber)
				UserChangeDocCheck = False
				dsTemp.Close
				Exit Function
			End If
			'По отмененным плановым служебкам нельзя создавать фактические
			If MyCStr(dsTemp("StatusCompletion")) = "0" Then
				Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ParentDocCancelled)
				UserChangeDocCheck = False
				dsTemp.Close
				Exit Function
			End If
			'По неутвержденным  плановым служебкам нельзя создавать фактические
			If MyCStr(dsTemp("NameApproved")) = "" Then
				Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ParentDocNotApproved)
				UserChangeDocCheck = False
				dsTemp.Close
				Exit Function
			End If
		End If
		dsTemp.Close
	End If

	'Генерация номера
	If Request("create") = "y" Then
		sClassDoc = "Служебные записки%"
		S_DocID = GetNewDocIDForSluzhZapiski(sClassDoc, sDepartmentRoot, Request("UserFieldText1"), "")
		S_DocIDAdd = S_DocID
	End If

	'Проверяем номера проектов и получаем список руководителей
	S_UserFieldText3 = Request("UserFieldText3")
	ProjectManagers = ""
	If not CheckAndEnhanceProjectList(S_UserFieldText3, ProjectManagers, Conn) Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInProjectNumbers & S_UserFieldText3)
		'Возвращаем исходное значение в S_UserFieldText3
		S_UserFieldText3 = Request("UserFieldText3")
		UserChangeDocCheck = False
		Exit Function
	End If

'{Запрос №50 - СТС
	'GetReconciliationByRules(parDocID, parClassDoc, parAmount, parChartOfAccounts, parCostCenter, parProjectCode, parBusinessUnit, parPartnerName, parKindOfPayment, parInitiator, parProjectManager, parOvertimeRequester, parIncomeExpenceContract, parContranctInSTS, parCurrency, parContractType, parOvertimeFuncLeaders, ByRef parListToReconcile, ByRef parNameAproval, ByRef parNameResponsible, ByRef parCorrespondent, ByRef parListToView, ByRef parRegistrar, ByRef parNameControl, ByRef parListToEdit)
	'GetReconciliationByRules iif(Trim(Request("DocID")) = "", S_DocID, Request("DocID")), Session("CurrentClassDoc"), Null, Null, sCostCenter, GetCodeFromCode_NameString(Request("UserFieldText3")), GetCodeFromCode_NameString(Request("BusinessUnit")), Null, "", iif(UCase(Request("create")) = "Y", Session("UserID"), Request("DocAuthor")), Replace(ProjectManagers, VbCrLf, " "), Request("DocNameResponsible"), NULL, NULL, NULL, NULL, Replace(Request("DocCorrespondent"), VbCrLf, " "), S_ListToReconcile, S_NameAproval, Null, Null, S_ListToView, Null, Null, Null
	par_DocID = iif(Trim(Request("DocID")) = "", S_DocID, Request("DocID"))
	par_ClassDoc = Session("CurrentClassDoc")
	par_Amount = NULL
	par_ChartOfAccounts = NULL
	par_CostCenter = sCostCenter
	par_ProjectCode = GetCodeFromCode_NameString(Request("UserFieldText3"))
	par_BusinessUnit = GetCodeFromCode_NameString(Request("BusinessUnit"))
	par_PartnerName = NULL
	par_KindOfPayment = ""
	par_Initiator = iif(UCase(Request("create")) = "Y", Session("UserID"), Request("DocAuthor"))
	par_ProjectManager = Replace(ProjectManagers, VbCrLf, " ")
	par_OvertimeRequester = Request("DocNameResponsible")
	par_IncomeExpenceContract = NULL
	par_ContranctInSTS = NULL
	par_Currency = NULL
	par_ContractType = NULL
	par_OvertimeFuncLeaders = Replace(Request("DocCorrespondent"), VbCrLf, " ")
	par_TypeOfDocument = ""
	par_FunctionArea = ""

	GetReconciliationByRules par_DocID, par_ClassDoc, par_Amount, par_ChartOfAccounts, par_CostCenter, par_ProjectCode, par_BusinessUnit, par_PartnerName, par_KindOfPayment, par_Initiator, par_ProjectManager, par_OvertimeRequester, par_IncomeExpenceContract, par_ContranctInSTS, par_Currency, par_ContractType, par_OvertimeFuncLeaders, par_TypeOfDocument, par_FunctionArea, _
		S_ListToReconcile, S_NameAproval, Null, Null, S_ListToView, Null, Null, Null
	'	par_ListToReconcile, par_NameAproval, par_NameResponsible, par_Correspondent, par_ListToView, par_Registrar, par_NameControl, par_ListToEdit
'Запрос №50 - СТС}

	'В доп. пользователей заносим всех руководителей исполнителя и согласующих, чтобы они имели доступ
'{ph - 20120812
	'S_AdditionalUsers = GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit"))
	S_AdditionalUsers = DeleteUserDoublesInListNew(Request("AdditionalUsers") & VbCrLf & GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit")), True)
'ph - 20120812}
	AddLogD "@@@Search For Leaders - S_AdditionalUsers: "+S_AdditionalUsers
	S_AdditionalUsers = DeleteUserDoublesInList(S_AdditionalUsers)
	If S_AdditionalUsers = "" Then
		S_AdditionalUsers = " "
	End If
	S_AdditionalUsers_Set = S_AdditionalUsers

'STS_SLUZH_ZAPISKA_HOLIDAY
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(STS_SLUZH_ZAPISKA_HOLIDAY)) = 1 Then
	'Создатель документа
	If UCase(Request("create")) = "Y" Then
		sDocCreator = Session("UserID")
	Else
		sDocCreator = GetUserID(Request("DocAuthor"))
	End If

	S_NameAproval=Request("DocNameAproval")
	S_NameResponsible=Request("DocNameResponsible")
	S_NameControl=Request("DocNameControl")
	S_ListToReconcile=Request("DocListToReconcile")
	S_ListToEdit=Request("DocListToEdit")
	S_ListToView=Request("DocListToView")
	S_Correspondent=Request("DocCorrespondent")

	'Определяем бизнес направление
	sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))

	' SAY 2008-07-21
	Set dsTempPR = Server.CreateObject("ADODB.Recordset")

' DmGorsky_5: меняем местами проверку: вначале ищем SIT_SITRU ("СИТРОНИКС ИТ"), потом, если не найдено, - SIT_SITRONICS ("СИТРОНИКС")
	If InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRU))=1 then ' DmGorsky
		sDepartment = SIT_SITRU ' DmGorsky
		sDepartmentCode = "" ' DmGorsky
'Запрос №1 - СИБ - start
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SIB_ROOT_DEPARTMENT))=1 then
		sDepartment = SIT_SIB
		sDepartmentCode = ""
'Запрос №1 - СИБ - end
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRONICS))=1 then
		sDepartment = SIT_SITRONICS
		sDepartmentCode = ""
	Else
		sDepartment = SIT_STS

		sSQL = "select code from Departments where Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"' or Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"/'"
		AddLogD "WAY sSQL="+sSQL
		AddLogD "WAY Conn="+Conn
		dsTempPR.Open sSQL, Conn, 3, 1, &H1

		if not dsTempPR.EOF Then
			sDepartmentCode = dsTempPR("code")
			if InStrRev(sDepartmentCode, "/") > 0 then
				sDepartmentCode = Right(sDepartmentCode, Len(sDepartmentCode)-InStrRev(sDepartmentCode, "/"))
			End If
		Else
			sDepartmentCode="000000"
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInDepartmentCode1 + DelOtherLangFromFolder(Request("DocDepartment")) + SIT_ErrorInDepartmentCode2)
			UserChangeDocCheck=False
			Exit Function
		End If

		dsTempPR.Close
	End If

	If Request("create") = "y" Then
		sPostfix = ""
		sSearchCol = "DocID"
		sPrePrefix = ""
		'SAY 2009-02-20
		sSufix = ""

		'SAY 2009-03-12
		S_DocID = ""

		sClassDoc = "Служебные записки%"
		S_DocID = GetNewDocIDForSluzhZapiski(sClassDoc, sDepartmentRoot, Request("UserFieldText1"), "")
		S_DocIDAdd = S_DocID
	End If
	If sDepartmentRoot <> SIT_STS or (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_RASP_DOCS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1) Then
		'реализация полужесткого маршрута
		S_ListToReconcile = Replace(S_ListToReconcile, Trim(SIT_AdditionalAgrees), "")
		'Для коммерческих предложений и заявок на закупку и оплату не нужна приписка про доп. согласующих  
		If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_COM_OFFERS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) <> 1 Then
			If S_ListToReconcile <> "" Then
				S_ListToReconcile = " " + SIT_AdditionalAgrees + DeleteConstPrefixFromList(S_ListToReconcile)
			End If
		End If
	End If

	If sDepartmentRoot = SIT_SITRONICS Then
		S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
	ElseIf sDepartmentRoot = SIT_STS Then
		'rmanyushin 89142 01.04.2010 Start
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME)) > 0 Then
			If is789DivisionSTS(Session("Department")) Then
				DocSession = "Служебные записки*Office memo*Interní sdělení/На переработки 7-8-9*Overtime 7-8-9*Overtime 7-8-9"
				S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",DocSession + "/","List")+STS_DptOSSBSS+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
				'S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+STS_DptOSSBSS+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
			Else
				S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
			End if
		End If
		'rmanyushin 89142 01.04.2010 Stop
'{ph - 20120326
		'rmanyushin 119579 19.08.2010 Start
'		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_HOLIDAY)) > 0 Then
'			S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
'		End If
		'rmanyushin 119579 19.08.2010 End
'{Запрос №50 - СТС
		'GetReconciliationByRules iif(Trim(Request("DocID")) = "", S_DocID, Request("DocID")), Session("CurrentClassDoc"), Null, Null, GetNearestCostCenterCodeByCode(sDepartmentCode), Null, GetCodeFromCode_NameString(Request("BusinessUnit")), Null, "", iif(UCase(Request("create")) = "Y", Session("UserID"), Request("DocAuthor")), Null, Null, NULL, NULL, NULL, NULL, NULL, S_ListToReconcile, S_NameAproval, Null, Null, S_ListToView, Null, Null, Null
		par_DocID = iif(Trim(Request("DocID")) = "", S_DocID, Request("DocID"))
		par_ClassDoc = Session("CurrentClassDoc")
		par_Amount = NULL
		par_ChartOfAccounts = NULL
		par_CostCenter = GetNearestCostCenterCodeByCode(sDepartmentCode)
		par_ProjectCode = NULL
		par_BusinessUnit = GetCodeFromCode_NameString(Request("BusinessUnit"))
		par_PartnerName = NULL
		par_KindOfPayment = ""
		par_Initiator = iif(UCase(Request("create")) = "Y", Session("UserID"), Request("DocAuthor"))
		par_ProjectManager = NULL
		par_OvertimeRequester = NULL
		par_IncomeExpenceContract = NULL
		par_ContranctInSTS = NULL
		par_Currency = NULL
		par_ContractType = NULL
		par_OvertimeFuncLeaders = NULL
		par_TypeOfDocument = ""
		par_FunctionArea = ""

		GetReconciliationByRules par_DocID, par_ClassDoc, par_Amount, par_ChartOfAccounts, par_CostCenter, par_ProjectCode, par_BusinessUnit, par_PartnerName, par_KindOfPayment, par_Initiator, par_ProjectManager, par_OvertimeRequester, par_IncomeExpenceContract, par_ContranctInSTS, par_Currency, par_ContractType, par_OvertimeFuncLeaders, par_TypeOfDocument, par_FunctionArea, _
			S_ListToReconcile, S_NameAproval, Null, Null, S_ListToView, Null, Null, Null
		'	par_ListToReconcile, par_NameAproval, par_NameResponsible, par_Correspondent, par_ListToView, par_Registrar, par_NameControl, par_ListToEdit
'Запрос №50 - СТС}
'ph - 20120326}
	ElseIf sDepartmentRoot = SIT_SIB Then'СИБ
		S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSIBRU","Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
	Else 'Другие БН
		S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
	End If

	'Для всех категорий кроме заявок подставляем пользователей вместо ролей
	If (InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1) and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) <> 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_DOGOVORI_NEW)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIS)) <> 0) Then
'{ph - 20120326
	  If sDepartmentRoot <> SIT_STS Then
'ph - 20120326}
		'Получаем список ролей
		If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) = 1 Then 'Для договоров получаем список ролей на всех языках (могут быть роли на разных языках)
			sRoleList = GetFullRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"), "", "", "")
		Else
			sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))
		End If
		'Подставляем роли в поля
		S_NameAproval = ReplaceRolesInList(S_NameAproval, sRoleList)
		S_NameResponsible = ReplaceRolesInList(S_NameResponsible, sRoleList)
		S_NameControl = ReplaceRolesInList(S_NameControl, sRoleList)
		S_ListToEdit = ReplaceRolesInList(S_ListToEdit, sRoleList)
		S_ListToView = ReplaceRolesInList(S_ListToView, sRoleList)
		S_ListToReconcile = ReplaceRolesInList(S_ListToReconcile, sRoleList)

		'Удаляем повторы в списке согласования
		S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)

		'vnik_error
		'Убираем автора документа из списка согласующих, т.к. после активации он не сможет загрузить новый документ
		'(работает пока только, если пользователь создает документ не под ролью)
		S_ListToReconcile = Replace(S_ListToReconcile,Request("DocAuthor"),"")
		'vnik_error
'{ph - 20120326
	  End If
'ph - 20120326}
	End If
	'Проверка полей с пользователями
	'Для STS отключаем проверку утверждающего, т.к. он рассчитывается
'{ph - 20120326
	'If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 Then
	If sDepartmentRoot <> SIT_STS Then
'ph - 20120326}
		UserChangeDocCheck = CheckSingleUserField(S_NameAproval)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameAproval")+SIT_ErrorInUserField2)
			Exit Function
		End If
		'Для заявки на закупку отключаем проверку ответственного, т.к. можно указывать подразделение
		'Для заявки на оплату тоже, т.к. автоматом ставится роль из специального справочника
		UserChangeDocCheck = CheckSingleUserField(S_NameResponsible)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameResponsible")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
	UserChangeDocCheck = CheckSingleUserField(S_NameControl)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameControl")+SIT_ErrorInUserField2)
		Exit Function
	End If
'	If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) = 0 Then 'В исходящих может быть любой текст
		UserChangeDocCheck = CheckMultiUserField(S_Correspondent)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("Correspondent")+SIT_ErrorInUserField2)
			Exit Function
		End If
'	End If
''Запрос №1 - СИБ - start
'	IF InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_SLUZH_ZAPISKA)) = 0) THEN
''Запрос №1 - СИБ - end
'		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) = 0 Then 'В СЗ на переработки работают правила
'			UserChangeDocCheck = CheckMultiUserField(S_ListToReconcile)
'			If not UserChangeDocCheck Then
'				Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToReconcile")+SIT_ErrorInUserField2)
'				Exit Function
'			End If
''Запрос №1 - СИБ - start
'		End If
'	END IF
'Запрос №1 - СИБ - end
	UserChangeDocCheck = CheckMultiUserField(S_ListToEdit)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToEdit")+SIT_ErrorInUserField2)
		Exit Function
	End If
	'UserChangeDocCheck = CheckMultiUserField(S_ListToView)
	'If not UserChangeDocCheck Then
	'	Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToView")+SIT_ErrorInUserField2)
	'	Exit Function
	'End If

'{ph - 20120812
	'S_AdditionalUsers = GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit"))
	S_AdditionalUsers = DeleteUserDoublesInListNew(Request("AdditionalUsers") & VbCrLf & GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit")), True)
'ph - 20120812}
	AddLogD "@@@Search For Leaders - S_AdditionalUsers: "+S_AdditionalUsers
	'vnik_payment_order
	'If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_PAYMENT_ORDER)) > 0 Then
	'	S_AdditionalUsers = S_AdditionalUsers + ReplaceRoleFromDir(SIT_Account_manager,sDepartmentRoot)       
	'End If        
	'vnik_payment_order
	S_AdditionalUsers = DeleteUserDoublesInList(S_AdditionalUsers)

	'vnik
	If S_AdditionalUsers = "" Then
		S_AdditionalUsers = "-"
	End If
	S_AdditionalUsers_Set = S_AdditionalUsers

	S_ListToReconcile_Set = S_ListToReconcile
	S_NameAproval_Set = S_NameAproval
	S_ListToEdit = S_ListToEdit
	S_NameResponsible_Set = S_NameResponsible
	S_NameControl_Set = S_NameControl
	S_ListToView_Set=S_ListToView

	If UCase(Request("UpdateDoc")) = "YES" And (Request("create") = "y"  Or Request("DocIDPrev") <> Request("DocID")) Then
		'out "123"
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		dsTemp.Open "select * from Docs where DocID='" + MakeSQLSafe(Request("DocID")) + "'", Conn 
		If Not dsTemp.EOF Then
			Session("Message") = DOCS_ALREADYEXISTS + ": " + Request("DocID")
			dsTemp.Close
			UserChangeDocCheck=False
			Exit Function
		End If
		dsTemp.Close
	End If

'	'rmanyushin 119579 19.08.2010 Start  
'	AddLogD "@@@STS_Holiday - S_ListToReconcile 7: "+S_ListToReconcile
'	S_ListToReconcile = RemoveUserFromListWithDescriptions(S_ListToReconcile, S_NameAproval)
'	S_ListToReconcile = Replace(S_ListToReconcile, VbCrLf+VbCrLf, VbCrLf)
'	If InStr(S_ListToReconcile, VbCrLf) = 1 Then
'		S_ListToReconcile = Replace(S_ListToReconcile, VbCrLf, "", 1 ,1)
'	End If
'	AddLogD "@@@STS_Holiday - S_ListToReconcile 8: "+S_ListToReconcile

	'Удаляем повторы в списке согласования
	AddLogD "vnik123" + Trim(S_ListToReconcile)
	S_ListToReconcile_Set = DeleteUserDoublesInList(S_ListToReconcile)
	AddLogD "vnik123" + Trim(S_ListToReconcile)
	AddLogD "vnik123" + Trim(S_ListToReconcile_Set)
	'rmanyushin 119579 19.08.2010 End  

	'ph - 20090808 - start
	If Request("DocIDParent") <> "" Then
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		bMyDocIDParentChanged = False
		If Request("create") <> "y" Then
			'Проверяем изменился ли номер родительского документа, если идет редактирование
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocID"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				bMyDocIDParentChanged = Request("DocIDParent")<>dsTemp("DocIDParent")
			End If
			dsTemp.Close
		End If
		If bMyDocIDParentChanged Then
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocIDParent"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				'sMessage = ""
				'UserChangeDocCheck = oPayDox.IsReadAccessUser(Session("UserID"), sMessage, dsTemp, DOCS_All, DOCS_NoReadAccess, DOCS_NoAccess, USER_Department, VAR_ExtInt, VAR_AdminSecLevel, VAR_StatusActiveUser, NULL)
				UserChangeDocCheck = IsReadAccessRS(dsTemp) 'IsReadAccessUser(Session("UserID"), dsTemp, NULL)
				If UserChangeDocCheck Then
					S_DocIDParent = dsTemp("DocID")
				End If
				dsTemp.Close
				If not UserChangeDocCheck Then
					Session("Message") = SIT_NoAccessToParentDoc'sMessage
					Exit Function
				End If
			End If
		End If
	End If
	'ph - 20090808 - end
	
	'{ph - 20120517
	If sDepartmentRoot = SIT_STS and InStr(Request("BusinessUnit"), "2010") <> 1 Then
		If Trim(Request("UserFieldDate1")) <> "" Then
			If ConvertToDate(Request("UserFieldDate1")) <= Date() Then
				Session("Message") = STS_ErrorInHolidayFirstDate
				UserChangeDocCheck = False
				Exit Function
			End If
		End If
	End If
'ph - 20120517}


' *** ОСТАЛЬНЫЕ СЛУЖЕБНЫЕ ЗАПИСКИ
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_SLUZH_ZAPISKA)) = 1 Then
	'Создатель документа
	If UCase(Request("create")) = "Y" Then
		sDocCreator = Session("UserID")
	Else
		sDocCreator = GetUserID(Request("DocAuthor"))
	End If

	S_NameAproval=Request("DocNameAproval")
	S_NameResponsible=Request("DocNameResponsible")
	S_NameControl=Request("DocNameControl")
	S_ListToReconcile=Request("DocListToReconcile")
	S_ListToEdit=Request("DocListToEdit")
	S_ListToView=Request("DocListToView")
	S_Correspondent=Request("DocCorrespondent")

	'Определяем бизнес направление
	sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))

	' SAY 2008-07-21
	Set dsTempPR = Server.CreateObject("ADODB.Recordset")

' DmGorsky_5: меняем местами проверку: вначале ищем SIT_SITRU ("СИТРОНИКС ИТ"), потом, если не найдено, - SIT_SITRONICS ("СИТРОНИКС")
	If InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRU))=1 then ' DmGorsky
		sDepartment = SIT_SITRU ' DmGorsky
		sDepartmentCode = "" ' DmGorsky
'Запрос №1 - СИБ - start
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SIB_ROOT_DEPARTMENT))=1 then
		sDepartment = SIT_SIB
		sDepartmentCode = ""
'Запрос №1 - СИБ - end
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_SITRONICS))=1 then
		sDepartment = SIT_SITRONICS
		sDepartmentCode = ""
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_RTI))=1 then
		sDepartment = SIT_RTI
		sDepartmentCode = ""		
	ElseIf InStr(UCase(Request("DocDepartment")),UCase(SIT_VTSS))=1 then
		sDepartment = SIT_VTSS
		sDepartmentCode = ""	
	Else
		sDepartment = SIT_STS

		sSQL = "select code from Departments where Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"' or Name="+sUnicodeSymbol+"'"+Request("DocDepartment")+"/'"
		AddLogD "WAY sSQL="+sSQL
		AddLogD "WAY Conn="+Conn
		dsTempPR.Open sSQL, Conn, 3, 1, &H1

		if not dsTempPR.EOF Then
			sDepartmentCode = dsTempPR("code")
			if InStrRev(sDepartmentCode, "/") > 0 then
				sDepartmentCode = Right(sDepartmentCode, Len(sDepartmentCode)-InStrRev(sDepartmentCode, "/"))
			End If
		Else
			sDepartmentCode="000000"
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInDepartmentCode1 + DelOtherLangFromFolder(Request("DocDepartment")) + SIT_ErrorInDepartmentCode2)
			UserChangeDocCheck=False
			Exit Function
		End If

		dsTempPR.Close
	End If

	If Request("create") = "y" Then
		sPostfix = ""
		sSearchCol = "DocID"
		sPrePrefix = ""
		'SAY 2009-02-20
		sSufix = ""

		'SAY 2009-03-12
		S_DocID = ""

		sClassDoc = "Служебные записки%"
		S_DocID = GetNewDocIDForSluzhZapiski(sClassDoc, sDepartmentRoot, Request("UserFieldText1"), "")
		S_DocIDAdd = S_DocID
	End If
	If sDepartmentRoot <> SIT_STS or (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_RASP_DOCS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1) Then
		'реализация полужесткого маршрута
		S_ListToReconcile = Replace(S_ListToReconcile, Trim(SIT_AdditionalAgrees), "")
		'Для коммерческих предложений и заявок на закупку и оплату не нужна приписка про доп. согласующих  
		If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_COM_OFFERS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) <> 1 Then
			If S_ListToReconcile <> "" Then
				S_ListToReconcile = " " + SIT_AdditionalAgrees + DeleteConstPrefixFromList(S_ListToReconcile)
			End If
		End If
	End If

	If sDepartmentRoot = SIT_SITRONICS Then
		S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
    ElseIf sDepartmentRoot = SIT_RTI Then
        S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeRTIRU","Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
    ElseIf sDepartmentRoot = SIT_VTSS Then
        S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeVTSSRU","Category",Session("CurrentClassDoc"),"List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
	ElseIf sDepartmentRoot = SIT_STS Then
		'rmanyushin 89142 01.04.2010 Start
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME)) > 0 Then
			If is789DivisionSTS(Session("Department")) Then
				DocSession = "Служебные записки*Office memo*Interní sdělení/На переработки 7-8-9*Overtime 7-8-9*Overtime 7-8-9"
				S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",DocSession + "/","List")+STS_DptOSSBSS+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
				'S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+STS_DptOSSBSS+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
			Else
				S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
			End if
		End If
		'rmanyushin 89142 01.04.2010 Stop

		'rmanyushin 119579 19.08.2010 Start
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_HOLIDAY)) > 0 Then
			S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSTSRU"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
		End If
		'rmanyushin 119579 19.08.2010 End
	ElseIf sDepartmentRoot = SIT_SIB Then'СИБ
		S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeSIBRU","Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
	Else 'Другие БН
		S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("Agree"+Request("l"),"Category",Session("CurrentClassDoc")+"/","List")+SIT_AdditionalAgreesDelimeter+S_ListToReconcile
	End If

	'Для всех категорий кроме заявок подставляем пользователей вместо ролей
	If (InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1) and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) <> 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_DOGOVORI_NEW)) <> 1 or InStr(UCase(Request("DocDepartment")), UCase(SIT_SIS)) <> 0) Then
		'Получаем список ролей
		If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) = 1 Then 'Для договоров получаем список ролей на всех языках (могут быть роли на разных языках)
			sRoleList = GetFullRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"), "", "", "")
		Else
			sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))
		End If
		'Подставляем роли в поля
		S_NameAproval = ReplaceRolesInList(S_NameAproval, sRoleList)
		S_NameResponsible = ReplaceRolesInList(S_NameResponsible, sRoleList)
		S_NameControl = ReplaceRolesInList(S_NameControl, sRoleList)
		S_ListToEdit = ReplaceRolesInList(S_ListToEdit, sRoleList)
		S_ListToView = ReplaceRolesInList(S_ListToView, sRoleList)
		S_ListToReconcile = ReplaceRolesInList(S_ListToReconcile, sRoleList)

		'Удаляем повторы в списке согласования
		S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)

		'vnik_error
		'Убираем автора документа из списка согласующих, т.к. после активации он не сможет загрузить новый документ
		'(работает пока только, если пользователь создает документ не под ролью)
		S_ListToReconcile = Replace(S_ListToReconcile,Request("DocAuthor"),"")
		'vnik_error
	End If
	'Проверка полей с пользователями
	'Для заявок отключаем проверку утверждающего, т.к. он рассчитывается
	If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 Then
		UserChangeDocCheck = CheckSingleUserField(S_NameAproval)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameAproval")+SIT_ErrorInUserField2)
			Exit Function
		End If
		'Для заявки на закупку отключаем проверку ответственного, т.к. можно указывать подразделение
		'Для заявки на оплату тоже, т.к. автоматом ставится роль из специального справочника
		UserChangeDocCheck = CheckSingleUserField(S_NameResponsible)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameResponsible")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
	UserChangeDocCheck = CheckSingleUserField(S_NameControl)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameControl")+SIT_ErrorInUserField2)
		Exit Function
	End If
	If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) = 0 Then 'В исходящих может быть любой текст
		UserChangeDocCheck = CheckMultiUserField(S_Correspondent)
		If not UserChangeDocCheck Then
			Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("Correspondent")+SIT_ErrorInUserField2)
			Exit Function
		End If
	End If
'Запрос №1 - СИБ - start
	IF InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 and (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_SLUZH_ZAPISKA)) = 0) THEN
'Запрос №1 - СИБ - end
		If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) = 0 Then 'В СЗ на переработки работают правила
			UserChangeDocCheck = CheckMultiUserField(S_ListToReconcile)
			If not UserChangeDocCheck Then
				Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToReconcile")+SIT_ErrorInUserField2)
				Exit Function
			End If
'Запрос №1 - СИБ - start
		End If
	END IF
'Запрос №1 - СИБ - end
	UserChangeDocCheck = CheckMultiUserField(S_ListToEdit)
	If not UserChangeDocCheck Then
		Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToEdit")+SIT_ErrorInUserField2)
		Exit Function
	End If
	'UserChangeDocCheck = CheckMultiUserField(S_ListToView)
	'If not UserChangeDocCheck Then
	'	Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToView")+SIT_ErrorInUserField2)
	'	Exit Function
	'End If

'{ph - 20120812
	'S_AdditionalUsers = GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit"))
	S_AdditionalUsers = DeleteUserDoublesInListNew(Request("AdditionalUsers") & VbCrLf & GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit")), True)
'ph - 20120812}
	AddLogD "@@@Search For Leaders - S_AdditionalUsers: "+S_AdditionalUsers
	'vnik_payment_order
	'If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_PAYMENT_ORDER)) > 0 Then
	'	S_AdditionalUsers = S_AdditionalUsers + ReplaceRoleFromDir(SIT_Account_manager,sDepartmentRoot)       
	'End If        
	'vnik_payment_order
	S_AdditionalUsers = DeleteUserDoublesInList(S_AdditionalUsers)

	'vnik
	If S_AdditionalUsers = "" Then
		S_AdditionalUsers = "-"
	End If
	S_AdditionalUsers_Set = S_AdditionalUsers

	S_ListToReconcile_Set = S_ListToReconcile
	S_NameAproval_Set = S_NameAproval
	S_ListToEdit = S_ListToEdit
	S_NameResponsible_Set = S_NameResponsible
	S_NameControl_Set = S_NameControl
	S_ListToView_Set=S_ListToView

	If UCase(Request("UpdateDoc")) = "YES" And (Request("create") = "y"  Or Request("DocIDPrev") <> Request("DocID")) Then
		'out "123"
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		dsTemp.Open "select * from Docs where DocID='" + MakeSQLSafe(Request("DocID")) + "'", Conn 
		If Not dsTemp.EOF Then
			Session("Message") = DOCS_ALREADYEXISTS + ": " + Request("DocID")
			dsTemp.Close
			UserChangeDocCheck=False
			Exit Function
		End If
		dsTemp.Close
	End If

	'ph - 20090808 - start
	If Request("DocIDParent") <> "" Then
		Set dsTemp = Server.CreateObject("ADODB.Recordset")
		bMyDocIDParentChanged = False
		If Request("create") <> "y" Then
			'Проверяем изменился ли номер родительского документа, если идет редактирование
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocID"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				bMyDocIDParentChanged = Request("DocIDParent")<>dsTemp("DocIDParent")
			End If
			dsTemp.Close
		End If
		If bMyDocIDParentChanged Then
			dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocIDParent"), "'", "''")+"'", Conn
			If not dsTemp.EOF Then
				'sMessage = ""
				'UserChangeDocCheck = oPayDox.IsReadAccessUser(Session("UserID"), sMessage, dsTemp, DOCS_All, DOCS_NoReadAccess, DOCS_NoAccess, USER_Department, VAR_ExtInt, VAR_AdminSecLevel, VAR_StatusActiveUser, NULL)
				UserChangeDocCheck = IsReadAccessRS(dsTemp)' IsReadAccessUser(Session("UserID"), dsTemp, NULL)
				If UserChangeDocCheck Then
					S_DocIDParent = dsTemp("DocID")
				End If
				dsTemp.Close
				If not UserChangeDocCheck Then
					Session("Message") = SIT_NoAccessToParentDoc'sMessage
					Exit Function
				End If
			End If
		End If
	End If
	'ph - 20090808 - end
'MIKRON protocol PC
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_PROTOCOL)) = 1 Then
   UserChangeDocCheck = Sit_PROTOCOL_PC_MIKRON()
   If not UserChangeDocCheck Then
      Exit Function
   End if

'MIKRON RL for protocol PC
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_RL_PROTOCOL)) = 1 Then
   UserChangeDocCheck = Sit_RLforPC_MIKRON()
   If not UserChangeDocCheck Then
      Exit Function
   End if

'MIKRON payment order
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_PAYMENT_ORDER)) = 1 Then
   UserChangeDocCheck = Sit_MIKRON_PAYMENT_ORDER()
   If not UserChangeDocCheck Then
      Exit Function
   End if

'MIKRON BSAP
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_BSAP)) > 0 Then
   UserChangeDocCheck = Sit_BSAP_MIKRON()
   If not UserChangeDocCheck Then
      Exit Function
   End if

'MIKRON purchase order
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_PURCHASE_ORDER)) > 0 Then
   UserChangeDocCheck = Sit_PURCHASE_ORDER_MIKRON()
   If not UserChangeDocCheck Then
      Exit Function
   End if
	
'MIKRON contracts
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_CONTRACT)) = 1 or _
       InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_S_CONTRACT)) = 1 or _
       InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_NDA_CONTRACT)) = 1 or _
       InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_OLD_CONTRACT)) = 1 or _
       InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_ADD_CONTRACT)) = 1 Then
   UserChangeDocCheck = Sit_CONTRACTS_MIKRON()
   If not UserChangeDocCheck Then
      Exit Function
   End if

'MIKRON RL_MEMO
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_RL_MEMO)) = 1 Then
   UserChangeDocCheck = Sit_RL_MEMO_MIKRON()
   If not UserChangeDocCheck Then
      Exit Function
   End if
'MIKRON EXPORT CONTRACTS
ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_EXPORT_CONTRACT)) = 1 or _
       InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_EXPADD_CONTRACT)) = 1 Then
   UserChangeDocCheck = Sit_EXPORT_CONTRACTS_MIKRON()
   If not UserChangeDocCheck Then
      Exit Function
   End if	

End If

End Function 'UserChangeDocCheck

' *********************************************************************************
' ***                            ПРОТОКОЛ ЗК МИКРОН                             ***
' ***                                                                           ***
' *** Retuns:                                                                   ***
' ***         True if successful                                                ***
' ***         False in case something wrong                                     ***
' *********************************************************************************
Function Sit_PROTOCOL_PC_MIKRON()

   Sit_PROTOCOL_PC_MIKRON = True

   'Создатель документа
   sDocCreator = iif (UCase(Request("create"))="Y", Session("UserID"), GetUserID(Request("DocAuthor")) )

   S_NameAproval=Request("DocNameAproval")
   S_ListToView=Request("DocListToView")

   'Определяем бизнес направление
   sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))
   sDepartment = SIT_MIKRON
   sDepartmentCode = ""

   If Request("create") = "y" Then
      sPostfix = ""
      sSearchCol = "DocID"
      sPrePrefix = ""
      sSufix = ""
      S_DocID = ""
      S_DocID = GetNewDocIDForPurchaseOrder(S_ClassDoc, sDepartmentRoot, "PRPCM-", "", "", "", "")
   End If

   'Получаем список ролей
   sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))

   'Подставляем роли в поля
   S_NameAproval = ReplaceRolesInList(S_NameAproval, sRoleList)
   S_ListToView = ReplaceRolesInList(S_ListToView, sRoleList)

   'Удаляем повторы в списке пользователей
   S_ListToView = DeleteUserDoublesInList(S_ListToView)    
   S_NameAproval_Set = S_NameAproval
   S_ListToView_Set=S_ListToView

   If UCase(Request("UpdateDoc")) = "YES" And (Request("create") = "y"  Or Request("DocIDPrev") <> Request("DocID")) Then
      Set dsTemp = Server.CreateObject("ADODB.Recordset")
      dsTemp.Open "select * from Docs where DocID='" + MakeSQLSafe(Request("DocID")) + "'", Conn 
      If Not dsTemp.EOF Then
         Session("Message") = DOCS_ALREADYEXISTS + ": " + Request("DocID")
         dsTemp.Close
         Sit_PROTOCOL_PC_MIKRON=False
         Exit Function
      End If
      dsTemp.Close
   End If

   If Request("DocIDParent") <> "" Then
      Set dsTemp = Server.CreateObject("ADODB.Recordset")
      bMyDocIDParentChanged = False
      If Request("create") <> "y" Then
         'Проверяем изменился ли номер родительского документа, если идет редактирование
         dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocID"), "'", "''")+"'", Conn
         If not dsTemp.EOF Then
            bMyDocIDParentChanged = Request("DocIDParent")<>dsTemp("DocIDParent")
         End If
         dsTemp.Close
      End If
      
      If bMyDocIDParentChanged Then
         dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocIDParent"), "'", "''")+"'", Conn
         If not dsTemp.EOF Then
			Sit_PROTOCOL_PC_MIKRON = IsReadAccessRS(dsTemp)
			If Sit_PROTOCOL_PC_MIKRON Then
               S_DocIDParent = dsTemp("DocID")
			End If
			dsTemp.Close
			If not Sit_PROTOCOL_PC_MIKRON Then
               Session("Message") = SIT_NoAccessToParentDoc'sMessage
               Exit Function
			End If
         End If
      End If
   End If

End Function '---------------------------       Sit_PROTOCOL_PC_MIKRON()

' *********************************************************************************
' ***                  ОПРОСНЫЙ ЛИСТ ДЛЯ ПРОТОКОЛА ЗК МИКРОН                    ***
' ***                                                                           ***
' *** Retuns:                                                                   ***
' ***         True if successful                                                ***
' ***         False in case something wrong                                     ***
' *********************************************************************************
Function Sit_RLforPC_MIKRON()

   Sit_RLforPC_MIKRON = True

   'Создатель документа
   sDocCreator = iif (UCase(Request("create"))="Y", Session("UserID"), GetUserID(Request("DocAuthor")) )

   S_NameAproval=Request("DocNameAproval")
   S_NameResponsible=Request("DocNameResponsible")
   S_NameControl=Request("DocNameControl")
   S_ListToReconcile=Request("DocListToReconcile")
   S_ListToEdit=Request("DocListToEdit")
   S_ListToView=Request("DocListToView")
   S_Correspondent=Request("DocCorrespondent")

   'Определяем бизнес направление
   sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))
   sDepartment = SIT_MIKRON
   sDepartmentCode = ""

   If Request("create") = "y" Then
      sPostfix = ""
      sSearchCol = "DocID"
      sPrePrefix = ""
      sSufix = ""
      S_DocID = ""
      S_DocID = GetNewDocIDForPurchaseOrder(S_ClassDoc, sDepartmentRoot, "RL_PRPCM-", "", "", "", "")
      S_DocIDAdd = S_DocID
   End If

   'Список согласования ЖЕСТКИЙ!!
   S_ListToReconcile = Replace(S_ListToReconcile, Trim(SIT_AdditionalAgrees), "")

   'Получаем список ролей
   sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))

   'Подставляем роли в поля
   S_NameAproval = ReplaceRolesInList(S_NameAproval, sRoleList)
   S_NameResponsible = ReplaceRolesInList(S_NameResponsible, sRoleList)
   S_NameControl = ReplaceRolesInList(S_NameControl, sRoleList)
   S_ListToEdit = ReplaceRolesInList(S_ListToEdit, sRoleList)
   S_ListToView = ReplaceRolesInList(S_ListToView, sRoleList)
   S_ListToReconcile = ReplaceRolesInList(S_ListToReconcile, sRoleList)

   'Удаляем повторы в списке согласования
   S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)

   'Убираем автора документа из списка согласующих, т.к. после активации он не сможет загрузить
   'новый документ (работает пока только, если пользователь создает документ не под ролью)
   S_ListToReconcile = Replace(S_ListToReconcile,Request("DocAuthor"),"")

   'Удаляем повторы в списке пользователей
   S_ListToView = DeleteUserDoublesInList(S_ListToView)    

   'Проверка полей с пользователями
   'Для заявок отключаем проверку утверждающего, т.к. он рассчитывается
   Sit_RLforPC_MIKRON = CheckSingleUserField(S_NameAproval)
   If not Sit_RLforPC_MIKRON Then
      Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameAproval")+SIT_ErrorInUserField2)
      Exit Function
   End If
   'Для заявки на закупку отключаем проверку ответственного, т.к. можно указывать подразделение
   'Для заявки на оплату тоже, т.к. автоматом ставится роль из специального справочника
   Sit_RLforPC_MIKRON = CheckSingleUserField(S_NameResponsible)
   If not Sit_RLforPC_MIKRON Then
      Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameResponsible")+SIT_ErrorInUserField2)
      Exit Function
   End If
   Sit_RLforPC_MIKRON = CheckMultiUserField(S_ListToEdit)
   If not Sit_RLforPC_MIKRON Then
      Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToEdit")+SIT_ErrorInUserField2)
      Exit Function
   End If

   ' Добавляем дополнительных пользователей здесь если сумма > 10.000.000
   If Request("AmountDoc") > 10000000 Then
      VNIK_ChiefDocDepartment = GetNearestChief(Request("DocDepartment"), sDocCreator, "")
      If VNIK_ChiefDocDepartment <> "#EMPTY" and VNIK_ChiefDocDepartment <> "#MANY" and VNIK_ChiefDocDepartment <> "" Then
         S_ListToView = Replace(S_ListToView, VbCrLf+VNIK_ChiefDocDepartment, "")
         S_ListToView = S_ListToView + VNIK_ChiefDocDepartment
      End If
   End If
   
   S_AdditionalUsers = DeleteUserDoublesInList(S_AdditionalUsers)
   S_AdditionalUsers_Set = S_AdditionalUsers

   S_NameAproval_Set = S_NameAproval
   S_ListToEdit = S_ListToEdit
   S_NameResponsible_Set = S_NameResponsible
   S_NameControl_Set = S_NameControl
   S_ListToView_Set=S_ListToView

   If UCase(Request("UpdateDoc")) = "YES" And (Request("create") = "y"  Or Request("DocIDPrev") <> Request("DocID")) Then
      Set dsTemp = Server.CreateObject("ADODB.Recordset")
      dsTemp.Open "select * from Docs where DocID='" + MakeSQLSafe(Request("DocID")) + "'", Conn 
      If Not dsTemp.EOF Then
         Session("Message") = DOCS_ALREADYEXISTS + ": " + Request("DocID")
         dsTemp.Close
         Sit_RLforPC_MIKRON=False
         Exit Function
      End If
      dsTemp.Close
   End If

   If Request("DocIDParent") <> "" Then
      Set dsTemp = Server.CreateObject("ADODB.Recordset")
      bMyDocIDParentChanged = False
      If Request("create") <> "y" Then
         'Проверяем изменился ли номер родительского документа, если идет редактирование
         dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocID"), "'", "''")+"'", Conn
         If not dsTemp.EOF Then
            bMyDocIDParentChanged = Request("DocIDParent")<>dsTemp("DocIDParent")
         End If
         dsTemp.Close
      End If
      
      If bMyDocIDParentChanged Then
         dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocIDParent"), "'", "''")+"'", Conn
         If not dsTemp.EOF Then
			Sit_RLforPC_MIKRON = IsReadAccessRS(dsTemp)
			If Sit_RLforPC_MIKRON Then
               S_DocIDParent = dsTemp("DocID")
			End If
			dsTemp.Close
			If not Sit_RLforPC_MIKRON Then
               Session("Message") = SIT_NoAccessToParentDoc'sMessage
               Exit Function
			End If
         End If
      End If
   End If

End Function '---------------------------       Sit_RLforPC_MIKRON()

' *********************************************************************************
' ***                       ЗАЯВКА  НА  ЗАКУПКУ  (МИКРОН)                       ***
' ***                                                                           ***
' *** Retuns:                                                                   ***
' ***         True if successful                                                ***
' ***         False in case something wrong                                     ***
' *********************************************************************************
Function Sit_PURCHASE_ORDER_MIKRON()

   Sit_PURCHASE_ORDER_MIKRON = True

   'Создатель документа	
   sDocCreator = iif (UCase(Request("create"))="Y", Session("UserID"), GetUserID(Request("DocAuthor")) )

   S_NameAproval=Request("DocNameAproval")
   S_NameResponsible=Request("DocNameResponsible")
   S_NameControl=Request("DocNameControl")
   S_ListToReconcile=Request("DocListToReconcile")
   S_ListToEdit=Request("DocListToEdit")
   S_ListToView=Request("DocListToView")
   S_Correspondent=Request("DocCorrespondent")

   'Определяем бизнес направление
   sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))
   sDepartment = SIT_MIKRON
   sDepartmentCode = ""
   
   If Request("create") = "y" Then
      sPostfix = ""
      sSearchCol = "DocID"
      sPrePrefix = ""
      sSufix = ""
      S_DocID = ""
      S_DocID = GetNewDocIDForPurchaseOrder(S_ClassDoc, sDepartmentRoot, "POM-", "", "", "", "")
   End If

  'формируем жесткий список согласования
   S_ListToReconcile = DeleteConstPrefixFromList(S_ListToReconcile)
   S_ListToReconcile = SIT_RequiredAgrees + VbCrLf + _
            oPayDox.GetExtTableValue("AgreeMIKRON","Category",Session("CurrentClassDoc"),"List") + _
            SIT_AdditionalAgreesDelimeter + Trim(S_ListToReconcile)

   'Руководитель ЦФЗ МИКРОН
   VNIK_TempValueAgree = GetUserDirValuesVNIK(MIKRON_CATALOG_FINANCIAL_COSTS,"Field3","Field1",Request("UserFieldText2"))
   If VNIK_TempValueAgree <> "" Then
      S_ListToReconcile = Replace(S_ListToReconcile, VbCrLf+VNIK_TempValueAgree, "")
      S_ListToReconcile = Replace(S_ListToReconcile, """#Руководитель ЦФЗ"";", VNIK_TempValueAgree)
   End If

   'Руководитель ЦФО МИКРОН
   VNIK_TempValueAgree = GetUserDirValuesVNIK(MIKRON_CATALOG_RESPONCIBILITY,"Field3","Field1",Request("UserFieldText5"))
   If VNIK_TempValueAgree <> "" Then
      S_ListToReconcile = Replace(S_ListToReconcile, VbCrLf+VNIK_TempValueAgree, "")
      S_ListToReconcile = Replace(S_ListToReconcile, """#Руководитель ЦФО"";", VNIK_TempValueAgree)
   End If

   'Руководитель направления = ЗГД кураторЦФЗ
   VNIK_TempValueAgree = GetUserDirValuesVNIK(MIKRON_CATALOG_FINANCIAL_COSTS,"Field4","Field1",Request("UserFieldText2"))
   If VNIK_TempValueAgree <> "" Then
      S_ListToReconcile = Replace(S_ListToReconcile, VbCrLf+VNIK_TempValueAgree, "")
   End If

   'убираем согласование ЗГД если сумма меньше 500000
   If Request("UserFieldMoney1") < 500000 Then
      VNIK_TempValueAgree = ""
   End If
   S_ListToReconcile = Replace(S_ListToReconcile, MIKRON_HeadOfInitiatorUnit, VNIK_TempValueAgree)

   'Дополняем список получателей
'   VNIK_ChiefDocDepartment = GetChiefOfDepartment(Request("DocDepartment"),"")
'   VNIK_ChiefDocDepartment = GetNearestChief(Request("DocDepartment"), sDocCreator, "")
'   VNIK_ChiefDocDepartment = GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit"))
'   VNIK_ChiefDocDepartment = GetAllUpperChiefsOfUsersFromList(S_NameResponsible,"")
'   If VNIK_ChiefDocDepartment <> "#EMPTY" and VNIK_ChiefDocDepartment <> "#MANY" and VNIK_ChiefDocDepartment <> "" Then
'      S_ListToView = Replace(S_ListToView, VbCrLf+VNIK_ChiefDocDepartment, "")
'      S_ListToView = S_ListToView + VNIK_ChiefDocDepartment
'   End If
'amw 08/07/2013

   'Удаляем повторы в списке согласования
   S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)
   S_ListToView = DeleteUserDoublesInList(S_ListToView)
   
   'Для всех категорий кроме заявок подставляем пользователей вместо ролей
   sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))

   'Подставляем роли в поля
   S_NameAproval = ReplaceRolesInList(S_NameAproval, sRoleList)
   S_NameResponsible = ReplaceRolesInList(S_NameResponsible, sRoleList)
   S_NameControl = ReplaceRolesInList(S_NameControl, sRoleList)
   S_ListToEdit = ReplaceRolesInList(S_ListToEdit, sRoleList)
   S_ListToView = ReplaceRolesInList(S_ListToView, sRoleList)
   S_ListToReconcile = ReplaceRolesInList(S_ListToReconcile, sRoleList)

   'Удаляем повторы в списке согласования
   S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)

   'Убираем автора документа из списка согласующих, т.к. после активации он не сможет загрузить
   'новый документ (работает пока только, если пользователь создает документ не под ролью)
   S_ListToReconcile = Replace(S_ListToReconcile,Request("DocAuthor"),"")
   'Убираем утверждающего документ из списка согласующих
   S_ListToReconcile = Replace(S_ListToReconcile,S_NameAproval,"")

   'Проверка полей с пользователями
   'Для заявок отключаем проверку утверждающего, т.к. он рассчитывается
   Sit_PURCHASE_ORDER_MIKRON = CheckSingleUserField(S_NameAproval)
   If not Sit_PURCHASE_ORDER_MIKRON Then
      Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameAproval")+SIT_ErrorInUserField2)
      Exit Function
   End If

   'Для заявки на закупку отключаем проверку ответственного, т.к. можно указывать подразделение
   'Для заявки на оплату тоже, т.к. автоматом ставится роль из специального справочника
   Sit_PURCHASE_ORDER_MIKRON = CheckSingleUserField(S_NameResponsible)
   If not Sit_PURCHASE_ORDER_MIKRON Then
      Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameResponsible")+SIT_ErrorInUserField2)
      Exit Function
   End If
	
   Sit_PURCHASE_ORDER_MIKRON = CheckSingleUserField(S_NameControl)
   If not Sit_PURCHASE_ORDER_MIKRON Then
      Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameControl")+SIT_ErrorInUserField2)
      Exit Function
   End If
	
   Sit_PURCHASE_ORDER_MIKRON = CheckMultiUserField(S_ListToEdit)
   If not Sit_PURCHASE_ORDER_MIKRON Then
      Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToEdit")+SIT_ErrorInUserField2)
      Exit Function
   End If

'amw 25-10-2013
'добавляем в список доп.пользователей ЗГД если сумма больше 82000000
   If Request("UserFieldMoney1") > 82000000 Then
      S_AdditionalUsers = DeleteUserDoublesInListNew(Request("AdditionalUsers") & VbCrLf & GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit")), True)
   End If

   S_ListToReconcile_Set = S_ListToReconcile
'amw 25-10-2013
'   Session("Message") = "<font color=red> S_ListToReconcile !</font> -> " & S_ListToReconcile
   S_NameAproval_Set = S_NameAproval
   S_NameResponsible_Set = S_NameResponsible
   S_ListToView_Set=S_ListToView

   If UCase(Request("UpdateDoc"))="YES" And (Request("create")="y" Or Request("DocIDPrev")<>Request("DocID")) Then
      Set dsTemp = Server.CreateObject("ADODB.Recordset")
      dsTemp.Open "select * from Docs where DocID='" + MakeSQLSafe(Request("DocID")) + "'", Conn
      bError = False
      If Not dsTemp.EOF Then
         bError = True
      End If
      dsTemp.Close
      If bError Then
         bError = False
         Session("Message") = DOCS_ALREADYEXISTS + ": " + Request("DocID")
         Sit_PURCHASE_ORDER_MIKRON = False
         Exit Function
      End If
   End If

   If Request("DocIDParent") <> "" Then
      Set dsTemp = Server.CreateObject("ADODB.Recordset")
      bMyDocIDParentChanged = False
      If Request("create") <> "y" Then
		 'Проверяем изменился ли номер родительского документа, если идет редактирование
		 dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocID"), "'", "''")+"'", Conn
		 If not dsTemp.EOF Then
		    bMyDocIDParentChanged = Request("DocIDParent")<>dsTemp("DocIDParent")
		 End If
		 dsTemp.Close
      End If
      If bMyDocIDParentChanged Then
         dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocIDParent"), "'", "''")+"'", Conn
         If not dsTemp.EOF Then
			Sit_PURCHASE_ORDER_MIKRON = IsReadAccessRS(dsTemp)
			If Sit_PURCHASE_ORDER_MIKRON Then
			   S_DocIDParent = dsTemp("DocID")
			End If
			dsTemp.Close
			If not Sit_PURCHASE_ORDER_MIKRON Then
			   Session("Message") = SIT_NoAccessToParentDoc  'sMessage
			   Exit Function
			End If
         End If
      End If
   End If
'amw 25-10-2013 (start)
   'Код статьи не соответствует справочнику
   If oPayDox.GetExtTableValue("Mikron_BudgetCode","Name",Request("UserFieldText4"),"Code") = "" Then
      Session("Message") = "<font color=red>ОШИБКА!</font> Не указана статья затрат" & VbCrLf & "-->" & Request("UserFieldText4")
      Sit_PURCHASE_ORDER_MIKRON=False
   End If
'amw 25-10-2013 (end)

End Function '---------------------------       Sit_PURCHASE_ORDER_MIKRON()

' ***
' *** БСАП МИКРОН
' *********************************************************************************
' ***        Бланк Сравнительного Анализа Поставщиков (БСАП) МИКРОН             ***
' ***  Заполняется при покупках на суммы меньше 500000. Утверждается ЗГД по ФиИ ***
' *** Retuns:                                                                   ***
' ***         True if successful                                                ***
' ***         False in case something wrong                                     ***
' *********************************************************************************
Function Sit_BSAP_MIKRON()
    
   Sit_BSAP_MIKRON = True
	
   'Создатель документа
   sDocCreator = iif (UCase(Request("create"))="Y", Session("UserID"), GetUserID(Request("DocAuthor")) )
   
   S_NameAproval=Request("DocNameAproval")
   S_ListToReconcile=Request("DocListToReconcile")
   S_ListToEdit=Request("DocListToEdit")
   S_ListToView=Request("DocListToView")
   S_Correspondent=Request("DocCorrespondent")

   'Определяем бизнес направление
   sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))
   sDepartment = SIT_MIKRON
   sDepartmentCode = ""

   If Request("create") = "y" Then
      sPostfix = ""
      sSearchCol = "DocID"
      sPrePrefix = ""
      sSufix = ""

      S_DocID = ""
            S_DocID = GetNewDocIDForPurchaseOrder(S_ClassDoc, sDepartmentRoot, "BSAP-", "", "", "", "")
      S_DocIDAdd = S_DocID
   End If

   If sDepartmentRoot <> SIT_STS or (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_RASP_DOCS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1) Then
      'реализация полужесткого маршрута
      S_ListToReconcile = Replace(S_ListToReconcile, Trim(SIT_AdditionalAgrees), "")
      'Для коммерческих предложений и заявок на закупку и оплату не нужна приписка про доп. согласующих  
      If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_COM_OFFERS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) <> 1 Then
         If S_ListToReconcile <> "" Then
            S_ListToReconcile = " " + SIT_AdditionalAgrees + DeleteConstPrefixFromList(S_ListToReconcile)
         End If
      End If
   End If

   'добавление обязательных согласующих
'amw 10/09/2013
'   S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeMIKRON","Category",Session("CurrentClassDoc"),"List")+SIT_AdditionalAgreesDelimeter+Trim(S_ListToReconcile)
   S_ListToReconcile = SIT_AdditionalAgreesDelimeter+Trim(S_ListToReconcile)
   'Удаляем повторы в списке согласования
   S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)
   S_ListToView = DeleteUserDoublesInList(S_ListToView)

   'Для всех категорий кроме заявок подставляем пользователей вместо ролей
   'Получаем список ролей
   If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) = 1 Then 'Для договоров получаем список ролей на всех языках (могут быть роли на разных языках)
      sRoleList = GetFullRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"), "", "", "")
   Else
      sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))
   End If
   
   'Подставляем роли в поля
   S_NameAproval = ReplaceRolesInList(S_NameAproval, sRoleList)
   S_ListToEdit = ReplaceRolesInList(S_ListToEdit, sRoleList)
   S_ListToView = ReplaceRolesInList(S_ListToView, sRoleList)
   S_ListToReconcile = ReplaceRolesInList(S_ListToReconcile, sRoleList)

   'Удаляем повторы в списке согласования
   S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)

   'Убираем автора документа из списка согласующих, т.к. после активации он не сможет загрузить новый документ
   '(работает пока только, если пользователь создает документ не под ролью)
   S_ListToReconcile = Replace(S_ListToReconcile,Request("DocAuthor"),"")
'amw 05/09/2013 start
   VNIK_ChiefDocDepartment = GetNearestChief(Request("DocDepartment"), sDocCreator, "")
   If VNIK_ChiefDocDepartment <> "#EMPTY" and VNIK_ChiefDocDepartment <> "#MANY" and VNIK_ChiefDocDepartment <> "" Then
      S_ListToReconcile = Replace(S_ListToReconcile, VbCrLf+VNIK_ChiefDocDepartment, "")   
      S_ListToReconcile = Replace(S_ListToReconcile, MIKRON_HeadOfInitiatorUnit, VNIK_ChiefDocDepartment)
   End If
'amw 05/09/2013 end
	
   'Проверка полей с пользователями
   'Для заявок отключаем проверку утверждающего, т.к. он рассчитывается
   Sit_BSAP_MIKRON = CheckSingleUserField(S_NameAproval)
   If not Sit_BSAP_MIKRON Then
      Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameAproval")+SIT_ErrorInUserField2)
      Exit Function
   End If
      
   Sit_BSAP_MIKRON = CheckMultiUserField(S_ListToEdit)
   If not Sit_BSAP_MIKRON Then
      Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToEdit")+SIT_ErrorInUserField2)
      Exit Function
   End If

   S_AdditionalUsers = DeleteUserDoublesInList(S_AdditionalUsers)

'vnik
   If S_AdditionalUsers = "" Then
      S_AdditionalUsers = "-"
   End If
   S_AdditionalUsers_Set = S_AdditionalUsers

   S_ListToReconcile_Set = S_ListToReconcile
   S_NameAproval_Set = S_NameAproval
   S_ListToEdit_Set = S_ListToEdit
   S_ListToView_Set=S_ListToView

   If UCase(Request("UpdateDoc"))="YES" And (Request("create")="y" Or Request("DocIDPrev") <> Request("DocID")) Then
      Set dsTemp = Server.CreateObject("ADODB.Recordset")
      dsTemp.Open "select * from Docs where DocID='" + MakeSQLSafe(Request("DocID")) + "'", Conn 
      If Not dsTemp.EOF Then
         Session("Message") = DOCS_ALREADYEXISTS + ": " + Request("DocID")
         dsTemp.Close
         Sit_BSAP_MIKRON=False
         Exit Function
      End If
      dsTemp.Close
   End If

   'ph - 20090808 - start
   If Request("DocIDParent") <> "" Then
      Set dsTemp = Server.CreateObject("ADODB.Recordset")
      bMyDocIDParentChanged = False
      If Request("create") <> "y" Then
         'Проверяем изменился ли номер родительского документа, если идет редактирование
         dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocID"), "'", "''")+"'", Conn
         If not dsTemp.EOF Then
            bMyDocIDParentChanged = Request("DocIDParent")<>dsTemp("DocIDParent")
         End If
         dsTemp.Close
      End If
      If bMyDocIDParentChanged Then
         dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocIDParent"), "'", "''")+"'", Conn
         If not dsTemp.EOF Then
            Sit_BSAP_MIKRON = IsReadAccessRS(dsTemp)
			If Sit_BSAP_MIKRON Then
               S_DocIDParent = dsTemp("DocID")
            End If
			dsTemp.Close
			If not Sit_BSAP_MIKRON Then
               Session("Message") = SIT_NoAccessToParentDoc
               Exit Function
            End If
         End If
      End If
   End If

End Function '---------------------------       Sit_BSAP_MIKRON()

' ***
' *** mikron_payment_order
' *********************************************************************************
' ***                             ЗАЯВКА НА ОПЛАТУ МИКРОН                       ***
' ***                                                                           ***
' *** Retuns:                                                                   ***
' ***         True if successful                                                ***
' ***         False in case something wrong                                     ***
' *********************************************************************************
Function Sit_MIKRON_PAYMENT_ORDER()

   Sit_MIKRON_PAYMENT_ORDER = True

   'Создатель документа
   sDocCreator = iif (UCase(Request("create"))="Y", Session("UserID"), GetUserID(Request("DocAuthor")) )

   S_NameAproval=Request("DocNameAproval")
   S_NameResponsible=Request("DocNameResponsible")
   S_NameControl=Request("DocNameControl")
   S_ListToReconcile=Request("DocListToReconcile")
   S_ListToEdit=Request("DocListToEdit")
   S_ListToView=Request("DocListToView")
   S_Correspondent=Request("DocCorrespondent")

   'Определяем бизнес направление
   sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))
   sDepartment = SIT_MIKRON
   sDepartmentCode = ""

   If Request("create") = "y" Then
      S_DocID = ""
            S_DocID = GetNewDocIDForPaymentOrder(S_ClassDoc, sDepartmentRoot, "PMIK-", "", "", "", "")
      S_DocIDAdd = S_DocID
   End If

   If sDepartmentRoot <> SIT_STS or (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ISHODYASCHIE)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_RASP_DOCS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1) Then
      'реализация полужесткого маршрута
      S_ListToReconcile = Replace(S_ListToReconcile, Trim(SIT_AdditionalAgrees), "")
      'Для коммерческих предложений и заявок на закупку и оплату не нужна приписка про доп. согласующих  
      If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_COM_OFFERS)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) <> 1 and InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) <> 1 Then
         If S_ListToReconcile <> "" Then
            S_ListToReconcile = " " + SIT_AdditionalAgrees + DeleteConstPrefixFromList(S_ListToReconcile)
         End If
      End If
   End If

   'добавление обязательных согласующих
   S_ListToReconcile = SIT_RequiredAgrees+oPayDox.GetExtTableValue("AgreeMIKRON","Category",Session("CurrentClassDoc"),"List")+SIT_AdditionalAgreesDelimeter+Trim(S_ListToReconcile)

   'Удаляем повторы в списке согласования и в получателях
   S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)
   S_ListToView = DeleteUserDoublesInList(S_ListToView)
   'S_ListToView = S_ListToView + Replace(Replace(Replace(S_ListToReconcile,"Обязательные согласующие: ",""), "Дополнительные согласующие: ", ""), "##;", "")

   'Для всех категорий кроме заявок подставляем пользователей вместо ролей
   If (InStr(UCase(Session("CurrentClassDoc")),UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) <> 1 or _
       InStr(UCase(Request("DocDepartment")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1) and _
       InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) <> 1 and _
       InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) <> 1 and _
       (InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_DOGOVORI_NEW)) <> 1 or _
        InStr(UCase(Request("DocDepartment")), UCase(SIT_STS)) <> 0) Then
      'Получаем список ролей
      If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) = 1 Then 'Для договоров получаем список ролей на всех языках (могут быть роли на разных языках)
         sRoleList = GetFullRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"), "", "", "")
      Else
         sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))
      End If
       
      'Подставляем роли в поля
      S_NameAproval = ReplaceRolesInList(S_NameAproval, sRoleList)
      S_NameResponsible = ReplaceRolesInList(S_NameResponsible, sRoleList)
      S_NameControl = ReplaceRolesInList(S_NameControl, sRoleList)
      S_ListToEdit = ReplaceRolesInList(S_ListToEdit, sRoleList)
      S_ListToView = ReplaceRolesInList(S_ListToView, sRoleList)
      S_ListToReconcile = ReplaceRolesInList(S_ListToReconcile, sRoleList)
       
      'Удаляем повторы в списке согласования
      S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)
       
      'Убираем автора документа из списка согласующих, т.к. после активации он не сможет загрузить новый документ
      '(работает пока только, если пользователь создает документ не под ролью)
      S_ListToReconcile = Replace(S_ListToReconcile,Request("DocAuthor"),"")
   End If

   'Проверка полей с пользователями
   'Для заявок отключаем проверку утверждающего, т.к. он рассчитывается
   If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 0 and InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) = 0 Then
      Sit_MIKRON_PAYMENT_ORDER = CheckSingleUserField(S_NameAproval)
      If not Sit_MIKRON_PAYMENT_ORDER Then
         Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameAproval")+SIT_ErrorInUserField2)
         Exit Function
      End If
   End If

   Sit_MIKRON_PAYMENT_ORDER = CheckMultiUserField(S_ListToEdit)
   If not Sit_MIKRON_PAYMENT_ORDER Then
      Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToEdit")+SIT_ErrorInUserField2)
      Exit Function
   End If

'amw   S_AdditionalUsers = GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit"))
'amw   S_AdditionalUsers = DeleteUserDoublesInListNew(Request("AdditionalUsers") & VbCrLf & GetAllUpperChiefsOfUsersFromList(S_NameResponsible + "; " + S_ListToReconcile, Request("BusinessUnit")), True)
   S_AdditionalUsers = DeleteUserDoublesInListNew(Request("AdditionalUsers")+VbCrLf+GetAllUpperChiefsOfUsersFromList(S_NameResponsible, ""), True)
'   AddLogD "@@@Search For Leaders - S_AdditionalUsers: "+S_AdditionalUsers
'amw   S_AdditionalUsers = DeleteUserDoublesInList(S_AdditionalUsers)

   If S_AdditionalUsers = "" Then
      S_AdditionalUsers = "-"
   End If
   S_AdditionalUsers_Set = S_AdditionalUsers

   S_ListToReconcile_Set = S_ListToReconcile
   S_NameAproval_Set = S_NameAproval
   S_ListToEdit = S_ListToEdit
   S_NameResponsible_Set = S_NameResponsible
   S_NameControl_Set = S_NameControl
   S_ListToView_Set=S_ListToView

   If UCase(Request("UpdateDoc")) = "YES" And (Request("create") = "y"  Or Request("DocIDPrev") <> Request("DocID")) Then
      Set dsTemp = Server.CreateObject("ADODB.Recordset")
      dsTemp.Open "select * from Docs where DocID='" + MakeSQLSafe(Request("DocID")) + "'", Conn 
      If Not dsTemp.EOF Then
         Session("Message") = DOCS_ALREADYEXISTS + ": " + Request("DocID")
         dsTemp.Close
         Sit_MIKRON_PAYMENT_ORDER=False
         Exit Function
      End If
      dsTemp.Close
   End If

   If Request("DocIDParent") <> "" Then
      Set dsTemp = Server.CreateObject("ADODB.Recordset")
      bMyDocIDParentChanged = False
      If Request("create") <> "y" Then
         'Проверяем изменился ли номер родительского документа, если идет редактирование
         dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocID"), "'", "''")+"'", Conn
         If not dsTemp.EOF Then
            bMyDocIDParentChanged = Request("DocIDParent")<>dsTemp("DocIDParent")
         End If
         dsTemp.Close
      End If
      If bMyDocIDParentChanged Then
         dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocIDParent"), "'", "''")+"'", Conn
         If not dsTemp.EOF Then
            Sit_MIKRON_PAYMENT_ORDER = IsReadAccessRS(dsTemp)' IsReadAccessUser(Session("UserID"), dsTemp, NULL)
			If Sit_MIKRON_PAYMENT_ORDER Then
               S_DocIDParent = dsTemp("DocID")
            End If
			dsTemp.Close
			If not Sit_MIKRON_PAYMENT_ORDER Then
               Session("Message") = SIT_NoAccessToParentDoc'sMessage
               Exit Function
            End If
         End If
      End If
   End If
End Function '---------------------------       Sit_MIKRON_PAYMENT_ORDER()

' *********************************************************************************
' ***                    ДОГОВОРЫ     М_И_К_Р_О_Н                               ***
' ***                                                                           ***
' *** Retuns:                                                                   ***
' ***         True if successful                                                ***
' ***         False in case something wrong                                     ***
' *********************************************************************************
Function Sit_CONTRACTS_MIKRON()

   Sit_CONTRACTS_MIKRON = True

   'Создатель документа	
   sDocCreator = iif (UCase(Request("create"))="Y", Session("UserID"), GetUserID(Request("DocAuthor")) )

   S_NameAproval=Request("DocNameAproval")
   S_NameResponsible=Request("DocNameResponsible")
   S_ListToReconcile=Request("DocListToReconcile")
   'Определяем бизнес направление
   sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))
   sDepartment = SIT_MIKRON

   If Request("create") = "y" Then
      S_DocID = ""
      If InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_NDA_CONTRACT)) = 1 Then
         S_DocID = GetNewDocIDForPaymentOrder(S_ClassDoc, sDepartmentRoot, "NDA-", "", "", "", "")
      ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_OLD_CONTRACT)) = 1 Then
         S_DocID = Request("DocID")
      Else
      S_DocID = GetNewDocIDForPaymentOrder(S_ClassDoc, sDepartmentRoot, "CNTM-", "", "", "", "")
      End If   
      S_DocIDAdd = S_DocID
   End If

   'Формируем список согласования. Админу даем возможность редактирования списка согласующих
''   If not IsAdmin() Then
''      S_ListToReconcile = AdditionalAgreeFromList(S_ListToReconcile)
      If InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_S_CONTRACT)) = 1 Then
         S_ListToReconcile = AdditionalAgreeFromList(S_ListToReconcile)
         S_ListToReconcile = SIT_PreliminaryAgrees + oPayDox.GetExtTableValue("AgreeMIKRON","Name",SIT_PreliminaryAgrees,"List") + vbCrLf + _
                       iif( S_ListToReconcile <> "", SIT_AdditionalAgrees + S_ListToReconcile + SIT_AdditionalAgreesDelimeter + vbCrLf, "") + _
                       SIT_RequiredAgrees + _
                       oPayDox.GetExtTableValue("AgreeMIKRON","Name",MIKRON_SalesAgrees,"List") + _
                       oPayDox.GetExtTableValue("AgreeMIKRON","Name",SIT_RequiredAgrees,"List")
'amw 25-08-2014 (start)
         'Код статьи не соответствует справочнику
         Num = CInt( oPayDox.GetExtTableValue("Mikron_BudgetCode","Name",Request("UserFieldText4"),"Code") )
         If Num > 1075 or Num < 1030 Then 
            S_ListToReconcile = Replace(S_ListToReconcile,oPayDox.GetExtTableValue("AgreeMIKRON","Name",MIKRON_SalesAgrees,"List"),"")
         End If
'amw 25-08-2014 (end)
      ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_NDA_CONTRACT)) = 1 Then
         S_ListToReconcile = SIT_RequiredAgrees + _
                      oPayDox.GetExtTableValue("AgreeMIKRON","Name",SIT_PreliminaryAgrees,"List") + _
                      oPayDox.GetExtTableValue("AgreeMIKRON","Category",Session("CurrentClassDoc"),"List") + _
                      SIT_AdditionalAgreesDelimeter + DeleteConstPrefixFromList(S_ListToReconcile)
      ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_OLD_CONTRACT)) = 1 Then
         S_ListToReconcile = S_ListToReconcile
      Else
         S_ListToReconcile = AdditionalAgreeFromList(S_ListToReconcile)
         S_ListToReconcile = SIT_PreliminaryAgrees + oPayDox.GetExtTableValue("AgreeMIKRON","Name",SIT_PreliminaryAgrees,"List") + vbCrLf + _
                       iif( S_ListToReconcile <> "", SIT_AdditionalAgrees + S_ListToReconcile + SIT_AdditionalAgreesDelimeter + vbCrLf, "") + _
                       SIT_RequiredAgrees + oPayDox.GetExtTableValue("AgreeMIKRON","Name",SIT_RequiredAgrees,"List")
      End If
''   End If

   'Для всех категорий кроме заявок подставляем пользователей вместо ролей
   'Получаем список ролей
   sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))

   'Подставляем роли в поля
   S_NameAproval = ReplaceRolesInList(S_NameAproval, sRoleList)
   S_NameResponsible = ReplaceRolesInList(S_NameResponsible, sRoleList)
   S_NameControl = ReplaceRolesInList(S_NameControl, sRoleList)
   S_ListToEdit = ReplaceRolesInList(S_ListToEdit, sRoleList)
   S_ListToView = ReplaceRolesInList(S_ListToView, sRoleList)
   S_ListToReconcile = ReplaceRolesInList(S_ListToReconcile, sRoleList)

   'Убираем автора документа из списка согласующих, т.к. после активации он не сможет загрузить
   'новый документ(работает пока только, если пользователь создает документ не под ролью)
   S_ListToReconcile = Replace(S_ListToReconcile,Request("DocAuthor"),"")

   'Руководитель направления = ЗГД кураторЦФЗ
   VNIK_ChiefDocDepartment = GetNearestChief(Request("DocDepartment"), S_NameResponsible, "")
   If VNIK_ChiefDocDepartment <> "#EMPTY" and VNIK_ChiefDocDepartment <> "#MANY" and VNIK_ChiefDocDepartment <> "" Then
      S_ListToReconcile = Replace(S_ListToReconcile, VbCrLf+VNIK_ChiefDocDepartment, "")   
      S_ListToReconcile = Replace(S_ListToReconcile, MIKRON_HeadOfInitiatorUnit, VNIK_ChiefDocDepartment)
   End If
   'Удаляем утверждающего (подписанта) Договора из списка согласующих.
   S_ListToReconcile = Replace(S_ListToReconcile, S_NameAproval, "")

   'Удаляем повторы в списке согласования
   S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   'Проверка полей с пользователями
   Sit_CONTRACTS_MIKRON = CheckSingleUserField(S_NameAproval)
   If not Sit_CONTRACTS_MIKRON Then
      Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameAproval")+SIT_ErrorInUserField2)
      Exit Function
   End If
      
   Sit_CONTRACTS_MIKRON = CheckMultiUserField(S_ListToEdit)
   If not Sit_CONTRACTS_MIKRON Then
      Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToEdit")+SIT_ErrorInUserField2)
      Exit Function
   End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'  Дополнительных пользователей здесь (ДОГОВОРа!) не надо
   S_AdditionalUsers = DeleteUserDoublesInList(S_AdditionalUsers)
   S_AdditionalUsers_Set = S_AdditionalUsers

   S_ListToReconcile_Set = S_ListToReconcile
''   S_NameAproval_Set = S_NameAproval
   S_NameResponsible_Set = S_NameResponsible
   S_NameControl_Set = S_NameControl
   S_ListToView_Set=S_ListToView

'дополнительные поля: обязательные
   S_PartnerName_Set = S_PartnerName ' контрагент
   S_Description_Set = S_Description
   S_AmountDoc_Set = S_AmountDoc 'Money
   S_Currency_Set = S_Currency 'Валюта

   If UCase(Request("UpdateDoc"))="YES" And _
      (Request("create")="y" Or Request("DocIDPrev") <> Request("DocID")) Then
      Set dsTemp = Server.CreateObject("ADODB.Recordset")
      dsTemp.Open "select * from Docs where DocID='" + MakeSQLSafe(Request("DocID")) + "'", Conn 
      If Not dsTemp.EOF Then
         Session("Message") = DOCS_ALREADYEXISTS + ": " + Request("DocID")
         dsTemp.Close
         Sit_CONTRACTS_MIKRON=False
         Exit Function
      End If
      dsTemp.Close
   End If

   If Request("DocIDParent") <> "" Then
      Set dsTemp = Server.CreateObject("ADODB.Recordset")
      bMyDocIDParentChanged = False
      If Request("create") <> "y" Then
         'Проверяем изменился ли номер родительского документа, если идет редактирование
         dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocID"), "'", "''")+"'", Conn
         If not dsTemp.EOF Then
            bMyDocIDParentChanged = Request("DocIDParent")<>dsTemp("DocIDParent")
         End If
         dsTemp.Close
      End If
      If bMyDocIDParentChanged Then
         dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocIDParent"), "'", "''")+"'", Conn
         If not dsTemp.EOF Then
            Sit_CONTRACTS_MIKRON = IsReadAccessRS(dsTemp)
            If Sit_CONTRACTS_MIKRON Then
               S_DocIDParent = dsTemp("DocID")
            End If
            dsTemp.Close
            
            If not Sit_CONTRACTS_MIKRON Then
               Session("Message") = SIT_NoAccessToParentDoc
               Exit Function
            End If
         End If
      End If
   End If

End Function '---------------------------       Sit_CONTRACTS_MIKRON()

' *********************************************************************************
' ***            ЭКСПОРТНЫЕ КОНТРАКТЫ     М_И_К_Р_О_Н                           ***
' ***                                                                           ***
' *** Retuns:                                                                   ***
' ***         True if successful                                                ***
' ***         False in case something wrong                                     ***
' *********************************************************************************
Function Sit_EXPORT_CONTRACTS_MIKRON()

   Sit_EXPORT_CONTRACTS_MIKRON = True

   'Создатель документа	
   sDocCreator = iif (UCase(Request("create"))="Y", Session("UserID"), GetUserID(Request("DocAuthor")) )

   S_NameAproval=Request("DocNameAproval")
   S_NameResponsible=Request("DocNameResponsible")
   S_ListToReconcile=Request("DocListToReconcile")
   S_Name = Request("DocName")
   'Определяем бизнес направление
   sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))
   sDepartment = SIT_MIKRON

   If Request("create") = "y" Then
      S_DocID = ""
      S_DocID = GetNewDocIDForPaymentOrder(S_ClassDoc, sDepartmentRoot, "EXPM-", "", "", "", "")
      S_DocIDAdd = S_DocID
   End If
   
   'Формируем список согласования. Админу даем возможность редактирования списка согласующих
'   If not IsAdmin() Then
'      S_ListToReconcile = AdditionalAgreeFromList(S_ListToReconcile)
      If InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_EXPORT_CONTRACT)) = 1 Then
         S_ListToReconcile = SIT_PreliminaryAgrees + oPayDox.GetExtTableValue("AgreeMIKRON","Name",SIT_PreliminaryAgrees,"List") + vbCrLf + _
                       iif( S_ListToReconcile <> "", SIT_AdditionalAgrees + S_ListToReconcile + SIT_AdditionalAgreesDelimeter + vbCrLf, "") + _
                       SIT_RequiredAgrees + oPayDox.GetExtTableValue("AgreeMIKRON","Name","Контракт (экспорт):","List")
         'Код статьи не соответствует справочнику
         If Request("UserFieldText4") <> "" Then
            Num = CInt( oPayDox.GetExtTableValue("Mikron_BudgetCode","Name",Request("UserFieldText4"),"Code") )
            If Num > 1063 or Num < 1060 Then
               Session("Message") = "<font color=red>ОШИБКА!</font> Не указана статья затрат" & VbCrLf & "-->" & Request("UserFieldText4")
               Sit_EXPORT_CONTRACTS_MIKRON=False
               Exit Function
            End If
         End If
      Else 'ДОПОЛНЕНИЯ к Экспортному контракту
         S_ListToReconcile = SIT_PreliminaryAgrees + oPayDox.GetExtTableValue("AgreeMIKRON","Name",SIT_PreliminaryAgrees,"List")
         If InStr(S_Name, MIK_EA_1) = 1 Then 'Здесь Юристов НЕ НЕДО !!
            S_ListToReconcile = oPayDox.GetExtTableValue("AgreeMIKRON","Name","Дополнение на разовую отгрузку (экспорт):","List")
         ElseIf InStr(S_Name, MIK_EA_2) = 1 Then
            S_ListToReconcile = S_ListToReconcile + vbCrLf + SIT_RequiredAgrees + _
                                oPayDox.GetExtTableValue("AgreeMIKRON","Name","Дополнение на расширение номенклатуры (экспорт):","List")
         ElseIf InStr(S_Name, MIK_EA_3) = 1 Then
            S_ListToReconcile = S_ListToReconcile + vbCrLf + SIT_RequiredAgrees + _
                                oPayDox.GetExtTableValue("AgreeMIKRON","Name","Дополнение на добавление спецификации (экспорт):","List")
         ElseIf InStr(S_Name, MIK_EA_4) = 1 Then
            S_ListToReconcile = S_ListToReconcile + vbCrLf + SIT_RequiredAgrees + _
                                oPayDox.GetExtTableValue("AgreeMIKRON","Name","Дополнение на изменение условий  и пр.(экспорт):","List")
         Else
            S_ListToReconcile = ""
         End If
      End If
'   End If

   'Для всех категорий кроме заявок подставляем пользователей вместо ролей
   'Получаем список ролей
   sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))

   'Подставляем роли в поля
   S_NameAproval = ReplaceRolesInList(S_NameAproval, sRoleList)
   S_NameResponsible = ReplaceRolesInList(S_NameResponsible, sRoleList)
   S_NameControl = ReplaceRolesInList(S_NameControl, sRoleList)
   S_ListToEdit = ReplaceRolesInList(S_ListToEdit, sRoleList)
   S_ListToView = ReplaceRolesInList(S_ListToView, sRoleList)
   S_ListToReconcile = ReplaceRolesInList(S_ListToReconcile, sRoleList)

'   'Удаляем повторы в списке согласования
'   S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)

   'Убираем автора документа из списка согласующих, т.к. после активации он не сможет загрузить
   'новый документ(работает пока только, если пользователь создает документ не под ролью)
   S_ListToReconcile = Replace(S_ListToReconcile,Request("DocAuthor"),"")

   'Убираем ответственного по договору из списка согласующих, т.к. после активации он не сможет загрузить
   'новый документ(работает пока только, если пользователь создает документ не под ролью)
   S_ListToReconcile = Replace(S_ListToReconcile,S_NameResponsible,"")

   'Удаляем утверждающего (подписанта) Договора из списка согласующих.
   S_ListToReconcile = Replace(S_ListToReconcile, S_NameAproval, "")

   'Удаляем повторы в списке согласования
   S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   'Проверка полей с пользователями
   Sit_EXPORT_CONTRACTS_MIKRON = CheckSingleUserField(S_NameAproval)
   If not Sit_EXPORT_CONTRACTS_MIKRON Then
      Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameAproval")+SIT_ErrorInUserField2)
      Exit Function
   End If
      
   Sit_EXPORT_CONTRACTS_MIKRON = CheckMultiUserField(S_ListToEdit)
   If not Sit_EXPORT_CONTRACTS_MIKRON Then
      Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("ListToEdit")+SIT_ErrorInUserField2)
      Exit Function
   End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'  Дополнительных пользователей здесь (ДОГОВОРа!) не надо
   S_AdditionalUsers = DeleteUserDoublesInList(S_AdditionalUsers)
   S_AdditionalUsers_Set = S_AdditionalUsers

   S_ListToReconcile_Set = S_ListToReconcile
   S_NameAproval_Set = S_NameAproval
   S_NameResponsible_Set = S_NameResponsible
   S_NameControl_Set = S_NameControl
   S_ListToView_Set=S_ListToView

'дополнительные поля: обязательные
   S_PartnerName_Set = S_PartnerName ' контрагент
   S_Description_Set = S_Description
   S_AmountDoc_Set = S_AmountDoc 'Money
   S_Currency_Set = S_Currency 'Валюта

   If UCase(Request("UpdateDoc"))="YES" And _
      (Request("create")="y" Or Request("DocIDPrev") <> Request("DocID")) Then
      Set dsTemp = Server.CreateObject("ADODB.Recordset")
      dsTemp.Open "select * from Docs where DocID='" + MakeSQLSafe(Request("DocID")) + "'", Conn 
      If Not dsTemp.EOF Then
         Session("Message") = DOCS_ALREADYEXISTS + ": " + Request("DocID")
         dsTemp.Close
         Sit_EXPORT_CONTRACTS_MIKRON=False
         Exit Function
      End If
      dsTemp.Close
   End If

   If Request("DocIDParent") <> "" Then
      Set dsTemp = Server.CreateObject("ADODB.Recordset")
      bMyDocIDParentChanged = False
      If Request("create") <> "y" Then
         'Проверяем изменился ли номер родительского документа, если идет редактирование
         dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocID"), "'", "''")+"'", Conn
         If not dsTemp.EOF Then
            bMyDocIDParentChanged = Request("DocIDParent")<>dsTemp("DocIDParent")
         End If
         dsTemp.Close
      End If
      If bMyDocIDParentChanged Then
         dsTemp.Open "select * from Docs where DocID="+sUnicodeSymbol+"'"+Replace(Request("DocIDParent"), "'", "''")+"'", Conn
         If not dsTemp.EOF Then
            Sit_EXPORT_CONTRACTS_MIKRON = IsReadAccessRS(dsTemp)
            If Sit_EXPORT_CONTRACTS_MIKRON Then
               S_DocIDParent = dsTemp("DocID")
            End If
            dsTemp.Close
            
            If not Sit_EXPORT_CONTRACTS_MIKRON Then
               Session("Message") = SIT_NoAccessToParentDoc
               Exit Function
            End If
         End If
      End If
   End If

End Function '---------------------------       Sit_EXPORT_CONTRACTS_MIKRON()

' *********************************************************************************
' ***  Закупки Микрон. Справка к листу согласования в рамках ранее заключенных  ***
' ***  договоров. Используется для создания доп.соглашений. Нужна проверка на   ***
' ***  наличие основного договора. В будущем добавить проверку на лимит закупок ***
' ***                                                                           ***
' *** Retuns:                                                                   ***
' ***         True if successful                                                ***
' ***         False in case something wrong                                     ***
' *********************************************************************************
Function Sit_RL_MEMO_MIKRON()

   Sit_RL_MEMO_MIKRON = True

   'Создатель документа
   sDocCreator = iif (UCase(Request("create"))="Y",Session("UserID"),GetUserID(Request("DocAuthor")) )

   S_NameAproval=Request("DocNameAproval")
   S_NameResponsible=Request("DocNameResponsible")
   S_ListToReconcile=Request("DocListToReconcile")

   'Определяем бизнес направление
   sDepartmentRoot = GetRootDepartment(Request("DocDepartment"))
   sDepartment = SIT_MIKRON

   If Request("create") = "y" Then
      sPostfix = ""
      sSearchCol = "DocID"
      sPrePrefix = ""
      sSufix = ""
      S_DocID = ""
      S_DocID = GetNewDocIDForPurchaseOrder(S_ClassDoc, sDepartmentRoot, "MEMO-", "", "", "", "")
      S_DocIDAdd = S_DocID
   End If   

  'формируем жесткий список согласования
   S_ListToReconcile = DeleteConstPrefixFromList(S_ListToReconcile)
   S_ListToReconcile = SIT_RequiredAgrees + VbCrLf + _
            oPayDox.GetExtTableValue("AgreeMIKRON","Category",Session("CurrentClassDoc"),"List") + _
            SIT_AdditionalAgreesDelimeter + Trim(S_ListToReconcile)

   'Руководитель ЦФЗ МИКРОН
   VNIK_TempValueAgree = GetUserDirValuesVNIK(MIKRON_CATALOG_FINANCIAL_COSTS,"Field3","Field1",Request("UserFieldText8"))
   If VNIK_TempValueAgree <> "" Then
      S_ListToReconcile = Replace(S_ListToReconcile, VbCrLf+VNIK_TempValueAgree, "")
      S_ListToReconcile = Replace(S_ListToReconcile, """#Руководитель ЦФЗ"";", VNIK_TempValueAgree)
   End If

   'Удаляем повторы в списке согласования
   S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)
   S_ListToView = DeleteUserDoublesInList(S_ListToView)
   
   'Для всех категорий кроме заявок подставляем пользователей вместо ролей
   sRoleList = GetRolesList(Request("DocDepartment"), sDocCreator, Request("BusinessUnit"))

   'Подставляем роли в поля
   S_NameAproval = ReplaceRolesInList(S_NameAproval, sRoleList)
   S_NameResponsible = ReplaceRolesInList(S_NameResponsible, sRoleList)
   S_NameControl = ReplaceRolesInList(S_NameControl, sRoleList)
   S_ListToEdit = ReplaceRolesInList(S_ListToEdit, sRoleList)
   S_ListToView = ReplaceRolesInList(S_ListToView, sRoleList)
   S_ListToReconcile = ReplaceRolesInList(S_ListToReconcile, sRoleList)

   'Удаляем повторы в списке согласования
   S_ListToReconcile = DeleteUserDoublesInList(S_ListToReconcile)

   'Убираем автора документа из списка согласующих, т.к. после активации он не сможет загрузить
   'новый документ (работает пока только, если пользователь создает документ не под ролью)
   S_ListToReconcile = Replace(S_ListToReconcile,Request("DocAuthor"),"")
   'Убираем утверждающего документ из списка согласующих
   S_ListToReconcile = Replace(S_ListToReconcile,S_NameAproval,"")

   'Проверка полей с пользователями
   'Для заявок отключаем проверку утверждающего, т.к. он рассчитывается
   Sit_RL_MEMO_MIKRON = CheckSingleUserField(S_NameAproval)
   If not Sit_RL_MEMO_MIKRON Then
      Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1+GetDocFieldDescription("NameAproval")+SIT_ErrorInUserField2)
      Exit Function
   End If

   'Для заявки на закупку отключаем проверку ответственного, т.к. можно указывать подразделение
   'Для заявки на оплату тоже, т.к. автоматом ставится роль из специального справочника
   Sit_RL_MEMO_MIKRON = CheckSingleUserField(S_NameResponsible)
   If not Sit_RL_MEMO_MIKRON Then
      Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1 + GetDocFieldDescription("NameResponsible")+SIT_ErrorInUserField2)
      Exit Function
   End If
	
   Sit_RL_MEMO_MIKRON = CheckSingleUserField(S_NameControl)
   If not Sit_RL_MEMO_MIKRON Then
      Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1 + GetDocFieldDescription("NameControl")+SIT_ErrorInUserField2)
      Exit Function
   End If
	
   Sit_RL_MEMO_MIKRON = CheckMultiUserField(S_ListToEdit)
   If not Sit_RL_MEMO_MIKRON Then
      Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInUserField1 + GetDocFieldDescription("ListToEdit")+SIT_ErrorInUserField2)
      Exit Function
   End If

'amw 25-10-2013 (start)
   'Код статьи не соответствует справочнику
   If oPayDox.GetExtTableValue("Mikron_BudgetCode","Name",Request("UserFieldText4"),"Code") = "" Then
      Session("Message") = "<font color=red>ОШИБКА!</font> Не указана статья затрат" & VbCrLf & "-->" & Request("UserFieldText4")
      Sit_RL_MEMO_MIKRON=False
      Exit Function
   End If
   If InStr( Request("DocDescription"), MIKRON_TEXT_KP_SELECT) = 1 Then 
      Session("Message") = "<font color=red>ОШИБКА!</font> Не указан критерий выбора"
      Sit_RL_MEMO_MIKRON=False
      Exit Function
   End If
'amw 25-10-2013 (end)
   If Request("DocIDPrevious") <> "" Then
     'Получим документ- договор основание
      Set vnikdsTemp1 = Server.CreateObject("ADODB.Recordset")
      sSQL = "SELECT * FROM Docs where DocID="+sUnicodeSymbol+"'"+MakeSQLSafeSimple(Request("DocIDPrevious"))+"'" + _
             " and (ClassDoc = "+sUnicodeSymbol+"'"+MakeSQLSafeSimple(MIKRON_CONTRACT)+ _
             "' or ClassDoc = "+sUnicodeSymbol+"'"+MakeSQLSafeSimple(MIKRON_OLD_CONTRACT)+"') and StatusDevelopment = 4"
      vnikdsTemp1.Open sSQL,Conn,3,1,&H1
      If vnikdsTemp1.EOF Then
         Session("Message") = "Договор основание <font color=red>"+Request("DocIDPrevious")+"</font> в базе не обнаружен"
         Sit_RL_MEMO_MIKRON = False
      Else If StrComp(vnikdsTemp1("PartnerName"),Request("DocPartnerName")) <> 0 Then
         Session("Message") = "Контрагент <font color=red>"+Request("DocPartnerName")+"</font> не соответствует договору"
         Sit_RL_MEMO_MIKRON = False
         End If
      End If
      vnikdsTemp1.Close        
   End If

'amw 06-10-2014 (end)

End Function '---------------------------       Sit_RL_MEMO_MIKRON()

' *********************************************************************************
' ***  Закупки Микрон.                                                          ***
' ***  Формируем лист согласования в виде:                                      ***
' ***  Предварительное согласование: Юристы                                     ***
' ***  Дополнительно согласующие (переменная часть):                            ***
' ***  Обязательные согласующие (постоянная часть                               ***
' ***  Разграничитель : "##"                                                    ***
' ***                                                                           ***
' *** Retuns:                                                                   ***
' ***         True if successful                                                ***
' ***         False in case something wrong                                     ***
' ***                                                                           ***
' *** Changes:                                                                  ***
' ***         19/06/2014 - STATUS OPEN. There were errors while                 ***
' *** AddListCorrespondebtRet.asp has used. Например: добавить дополнительного  ***
' *** пользователя или "На рецензию"                                            ***
' ***         04/08/2014 - FIXED. Added IsNull() additiondl check/              ***
' *********************************************************************************
Function AdditionalAgreeFromList(ByVal sUsersList)
   If Trim(sUsersList) = "" or IsNull(sUsersList) Then
      AdditionalAgreeFromList = ""
      Exit Function     
   End If

  sListAgree = Replace(sUsersList, vbCrLf, "")
  If InStr(sListAgree,"##") > 0 Then
     sListAgree = Left(sUsersList,InStr(sUsersList,"##") - 1)
     If InStr(sListAgree,SIT_AdditionalAgrees) > 0 Then
        sListAgree = Right(sListAgree, Len(sListAgree) - InStr(sListAgree,SIT_AdditionalAgrees) - Len(SIT_AdditionalAgrees) + 1)
     End If
  End If
  AdditionalAgreeFromList = sListAgree
End Function '---------------------------       AdditionalAgreeFromList()

'rmanyushin 60298 02.11.2009 Добавил S_AddField3
'Надо объявить за пределами функции, иначе переменная будет локальной
Dim S_AddField1, S_AddField2, S_AddField3

'Запрос №17 - СТС - start
'Запоминаем подразделение документа (для рассылок, зависимых от БН)
Session("CurrentDepartmentDoc") = S_Department
'Запрос №17 - СТС - end
%>