<%
'Set output columns for document lists
'Variables:
'dsDoc - current document recordset iterating in the loop
'Request("SomeKey") - some http get request key defining the showing document list, i.e. Request("VisaDocs") for reviewing docs or Request("ViewedDocs") for docs required to be viewed
'srSearchComments - if srSearchComments = "ON" then dsDoc recordset contains the Docs table and the joined Comment table fields, so you can use impressions like dsDoc("Comment") or dsDoc("Subject")

'Example:
'If Trim(Request("SomeKey")) = "y" Then
'	If IsNull(dsDoc("StatusCompletion")) Then
'		SkipThisRecord=False
'AddLogD "Skipped SomeKey for DocID:"+dsDoc("DocID")
'		Exit Function
'	End If	
'	If srSearchComments = "ON" Then
'		If dsDoc("Subject")="MySubject" Then
'			SkipThisRecord=False
'AddLogD "Skipped MySubject for DocID:"+dsDoc("DocID")
'			Exit Function
'		End If	
'	End If	
'End If	

  sDepartment = UCase(GetRootDepartment(Session("Department")))
  sDepartmentDoc = UCase(GetRootDepartment(dsDoc("Department")))

'ph - 20090714 - start - Механизм предоставления доступа переделан
  'SAY 2008-11-11 SAY 2008-10-15 предоставляем доступ пользователю ко всей категории
  'SAY 2009-02-20 добавлено условие активности документа.
'  If Instr(SIT_UsersListToAccessAllCategoryDocs,Session("UserID"))>0 and sDepartmentDoc=sDepartment and dsDoc("IsActive")="Y" Then
'    If Trim(Request("ClassDoc")) <> "" or UCase(Trim(Request("AllDocs"))) = "Y" Then 'Ph - 20090303 - добавлено условие, чтобы документы не лезли во всех списках
'      SkipThisRecord=False
'      Exit Function
'	End If
'  elseIf (VAR_ReadAccess="Y") Then
'    SkipThisRecord=True
'    Exit Function
'  End If 


AddLogD "####Access to doc - DocID: "+dsDoc("DocID")
AddLogD "####Instr(SIT_UsersListToAccessAllCategoryDocs,Session(""UserID""))>0: "+CStr(Instr(SIT_UsersListToAccessAllCategoryDocs,"<"+Session("UserID")+">")>0)
AddLogD "####dsDoc(""IsActive""): "+dsDoc("IsActive")
AddLogD "####sDepartmentDoc: "+sDepartmentDoc
AddLogD "####sDepartment: "+sDepartment
  If Instr(SIT_UsersListToAccessAllCategoryDocs,"<"+Session("UserID")+">")>0 and dsDoc("IsActive")="Y" Then
AddLogD "####1"
    If sDepartmentDoc = sDepartment and (Trim(Request("ClassDoc")) <> "" or UCase(Trim(Request("AllDocs"))) = "Y") Then 'Ph - 20090303 - добавлено условие по спискам, чтобы документы не лезли во всех списках
AddLogD "####2 - Not skipped"
      SkipThisRecord = False
      Exit Function
    Else
AddLogD "####3"
      oPayDox.VAR_ReadAccess = ""
      SkipThisRecord = Not IsReadAccessRS(dsDoc)
      oPayDox.VAR_ReadAccess = "Y"

      If SkipThisRecord Then
AddLogD "####4"
        If Request("supervisor")="y" Then
AddLogD "####5"
          If CheckPermit(S_Permitions,"SUPERVISORDEPARTMENT") And Request("UserIDToSeeDept")=Session("Department") Then
AddLogD "####6"
            If IsReadAccessUser(Request("UserIDToSee"), dsDoc, rsUser) Then
AddLogD "####7"
              SkipThisRecord=False
            End If
          End If
        End If
      Else
'ph - 20090823 - в ознакомление идут документы с которыми уже ознакомились - start
        If Trim(Request("ViewedDocs")) = "y" or (Request("VisaDocs") = "y" and Request("UserIDToSee") <> "") Then
AddLogD "####8"
          If Not IsNull(dsDoc("CommentType")) Then
AddLogD "####9"
            If dsDoc("CommentType")="VIEWED" And dsDoc("UserID")=Session("UserID") Then
AddLogD "####10"
              SkipThisRecord = True
            End If
          End If
        End If
'ph - 20090823 - в ознакомление идут документы с которыми уже ознакомились - end
      End If
AddLogD "####SkipThisRecord: "+CStr(SkipThisRecord)
      Exit Function
    End If
  End If
'ph - 20090714 - end

'Со списка Мои документы снимаем фильтр
If Trim(Request("CreatedDocs")) = "y" Then
	SkipThisRecord=False
	Exit Function
End If	

'Ph - 20081022 - показ утверждающему документов, требующих его утверждениея в списке Все новые
If Trim(Request("VisaDocs")) = "y" Then
	If (InStr(dsDoc("ListToReconcile"), "#!") > 0 or InStr(dsDoc("ListReconciled"), "-<") > 0) and (Session("UserID") = "mbondarenko" or Session("UserID") = "dmitry.kolesnikov" or Session("UserID") = "vsbrodov") Then
		SkipThisRecord = True
AddLogD "SkipThisRecord=True: TEMP - mbondarenko"
		Exit Function
	End If

	If Trim(Request("UserIDToSee"))<>"" Then
		'SAY 2008-10-29 добавлено условие and dsDoc("StatusDevelopment")="3" для отсечения всех кроме требующих утверждения
		If InStr(MyCStr(dsDoc("NameAproval")), "<" + Trim(Request("UserIDToSee")) + ">") > 0 and dsDoc("StatusDevelopment")="3" Then
			SkipThisRecord=False
			Exit Function
		End If
	End If

 'Запрос №44 - СТС - start
'ph - 20111208 - start
	'СОГЛАСОВАНИЕ - Для СТС приостановленные в согласовании договоры показываем инициатору, исполнителю и нач. юр. отдела (STS_JurChief)
	If Trim(Request("UserIDToSee"))="" Then
		If InStr(Session("Department"), SIT_STS_ROOT_DEPARTMENT) = 1 Then
				If (InStr(dsDoc("ListToReconcile"), "#!") > 0 or InStr(dsDoc("ListReconciled"), "-<") > 0) and (Session("UserID") = STS_JurChief or InStr(dsDoc("NameCreation"), "<" & Session("UserID") & ">") or InStr(dsDoc("NameResponsible"), "<" & Session("UserID") & ">")) Then
					SkipThisRecord = False
					Exit Function
				End If
		End If
	End If	
'ph - 20111208 - end

	If Trim(Request("UserIDToSee"))<>"" Then
		sUserIDToSee = Trim(Request("UserIDToSee"))
	Else
		sUserIDToSee = Session("UserID")
	End If
	
	If srSearchComments = "ON" Then
		If Trim(MyCStr(dsDoc("CommentType")))="VISA" Then
			If Trim(MyCStr(dsDoc("SpecialInfo")))="VISAWAITING" Then
				If dsDoc("UserID") <> sUserIDToSee Then
					SkipThisRecord = True
AddLogD "SkipThisRecord=True: VISAWAITING UserID<>sUserIDToSee"
					Exit Function
				End If
			End If
		End If

		If Trim(MyCStr(dsDoc("CommentType")))="REVIEW" Then
			If Trim(MyCStr(dsDoc("Address")))="REQUEST" And InStr(MyCStr(dsDoc("Comment")), "+<") <= 0 And (InStr(MyCStr(dsDoc("Comment")), "<" + sUserIDToSee + ">") > 0 Or dsDoc("UserID")=sUserIDToSee) Then
AddLogD "SkipThisRecord=False: REVIEW REQUEST: "+MyCStr(dsDoc("Comment"))+" -> "+sUserIDToSee
				SkipThisRecord = False
				Exit Function
			ElseIf Trim(MyCStr(dsDoc("Address")))="REQUEST" And (InStr(MyCStr(dsDoc("Comment")), "+<") > 0 Or IsReconciliationCompleteWithOptions(dsDoc("ListToReconcile"), dsDoc("ListReconciled")) Or ShowStatusDevelopment(dsDoc("StatusDevelopment"))=DOCS_RefusedApp Or ShowStatusDevelopment(dsDoc("StatusDevelopment"))=DOCS_Approved Or (not(IsNull(dsDoc("StatusCompletion"))) and (dsDoc("StatusCompletion")="1" or dsDoc("StatusCompletion")="0"))) Then
AddLogD "SkipThisRecord=True: REVIEW REQUEST ANSWERED"
					SkipThisRecord = True
					Exit Function
			End If
		End If
	End If

	If Trim(Request("UserIDToSee"))<>"" Then
		If InStr(Session("Department"), SIT_STS_ROOT_DEPARTMENT) = 1 Then
			If InStr(dsDoc("ClassDoc"), SIT_DOGOVORI_NEW) = 1 or InStr(dsDoc("ClassDoc"), SIT_DOGOVORI_OLD) = 1 Then
				If (InStr(dsDoc("ListToReconcile"), "#!") > 0 or InStr(dsDoc("ListReconciled"), "-<") > 0) and (Session("UserID") = STS_JurChief or InStr(dsDoc("NameCreation"), "<" & sUserIDToSee & ">") > 0 or InStr(dsDoc("NameResponsible"), "<" & sUserIDToSee & ">") > 0) Then
		SkipThisRecord = False
		Exit Function
	End If
				If InStr(dsDoc("NameResponsible"), "<" & sUserIDToSee & ">") > 0 and MyCStr(dsDoc("NameApproved") <> "") and InStr(MyCStr(dsDoc("NameApproved") <> ""), "-<") = 0 Then
					SkipThisRecord = True
AddLogD "SkipThisRecord=True: not to show approved docs to responsibles"
					Exit Function
				End If
			End If
		End If
	End If	
'	If (InStr(dsDoc("ListToReconcile"), "#!") > 0 or InStr(dsDoc("ListReconciled"), "-<") > 0) and InStr(dsDoc("ListToReconcile"), sUserIDToSee) > 0 Then
'		SkipThisRecord = False
'		Exit Function
'	End If
'Запрос №44 - СТС - end
End If

'rmanyushin 51555, 56781, 79501, 133266 05.10.2010 Start
If isPrivilegedUserSTS() Then
'Показать привилегированному пользователю СТС список активных документов с фильтром - только из СТС. 
	If sDepartmentDoc = sDepartment and dsDoc("IsActive")="Y" Then
		If UCase(Session("UserID")) = UCase(STS_Overseer) or UCase(Session("UserID")) = UCase(STS_Auditor) or UCase(Session("UserID")) = UCase(STS_POViewer) or UCase(Session("UserID")) = UCase(STS_LegalSTS) Then
			SkipThisRecord = False
		    Exit Function
        ElseIf UCase(Session("UserID")) = UCase(STS_HeadOf789) Then
		    If is789DivisionSTS(dsDoc("Department")) Then
		        SkipThisRecord = False
		        Exit Function
		    Else
		        SkipThisRecord = True
		        Exit Function
		    End If    
		End If       
	Else
		SkipThisRecord = True
		Exit Function
	End If
End If
'rmanyushin 51555, 56781, 79501, 133266 05.10.2010 End

'ph - 20101113 - start - для кнопки Рецензии
If UCase(Request("UserBtn")) = UCase("Reviews") Then
  If Trim(MyCStr(dsDoc("CommentType"))) = "REVIEW" Then
    If Trim(MyCStr(dsDoc("Address"))) = "REQUEST" And InStr(MyCStr(dsDoc("Comment")), "+<") <= 0 And (InStr(MyCStr(dsDoc("Comment")), "<" + Session("UserID") + ">") > 0 Or dsDoc("UserID") = Session("UserID")) Then
AddLogD "SkipThisRecord=False: REVIEW REQUEST: " & MyCStr(dsDoc("Comment")) & " -> " & Session("UserID")
      SkipThisRecord = False
      Exit Function
    ElseIf Trim(MyCStr(dsDoc("Address"))) = "REQUEST" And (InStr(MyCStr(dsDoc("Comment")), "+<") > 0 Or IsReconciliationCompleteWithOptions(dsDoc("ListToReconcile"), dsDoc("ListReconciled")) Or ShowStatusDevelopment(dsDoc("StatusDevelopment")) = DOCS_RefusedApp Or ShowStatusDevelopment(dsDoc("StatusDevelopment")) = DOCS_Approved Or (not(IsNull(dsDoc("StatusCompletion"))) and (dsDoc("StatusCompletion") = "1" or dsDoc("StatusCompletion") = "0"))) Then
AddLogD "SkipThisRecord=True: REVIEW REQUEST ANSWERED"
      SkipThisRecord = True
      Exit Function
    End If
  End If
End If
'ph - 20101113 - end

SkipThisRecord = True

%>