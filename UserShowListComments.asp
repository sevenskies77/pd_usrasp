<%
'Set output columns for document lists
'Variables:
'dsComments - Comment table recordset 
'VAR_CanDeleteComment - If True than current user can delete current comment in the comment list loop 
VAR_CanDeleteComment="" 
VAR_CanEditComment=""
VAR_CanViewComment=""
'Out "Comment:"+dsComments("Comment")
'VAR_NotToShowFiles="y" 'Not to download files to client

If UCase(Request.ServerVariables("URL")) = UCase("/ShowDoc.asp") Then
  'Нельзя удалять прикрепленные файлы, если началось согласование (первый человек согласовал)
  If dsComments("CommentType") = "FILE" Then
    If not IsAdmin() Then
      If S_IsActive = "Y" and ((S_ListReconciled <> "" and S_ListReconciled <> "-") or S_NameApproved <> "" or S_DateCompleted <> "") Then
        VAR_CanDeleteComment = "N"
      Else
        VAR_CanDeleteComment = ""
      End If
    End If
  End If

    If dsComments("CommentType") = "FILE" Then
    If CheckPermit(Session("Permitions"),"REGISTRAR") and InStr(UCase(CurrentClassDoc), UCase(SIT_ISHODYASCHIE))>0 and InStr(UCase(Session("Department")), UCase(SIT_MINC)) = 1 and instr(ucase(dsComments("UserId")),ucase(request("UserId"))) Then
        VAR_CanDeleteComment = "Y"
      Else
        VAR_CanDeleteComment = ""
      End If
    End If
  'rmanyushin 79439 16.02.2010 Start
  'Если пользователь привелигированный, то в карточке показывать раздел "Согласование".
  If isPrivilegedUserSTS() Then
    If UCase(Request("bVisa")) = "Y" or UCase(Request("showall")) = "Y" or UCase(Session("bViewDet"))= "Y" Then
      If dsComments("CommentType") = "VISA" Then
        VAR_CanViewComment = "Y"
      End If	
    End if
  End If
  'rmanyushin 79439 16.02.2010 End

'Запрос №33 - СТС - start
  'Проверяем не комментарий ли это об автоотмене
  If dsComments("CommentType") = "HISTORY" and dsComments("SpecialInfo") = "DOCS_Cancelled" and InStr(dsComments("Comment"), STS_AutoCancelledComment) > 0 Then
    'Устанавливаем флаг запрета возобновления документа, в UserShowDocAvailableButtons.asp будут спрятаны соответствующие кнопки
    bHideClickMakeCanceledCancel = True
  End If
'Запрос №33 - СТС - end
End If

'ph - 20120311 - start
If UCase(Request.ServerVariables("URL")) = UCase("/ShowDoc.asp") Then
	If InStr(S_AdditionalUsers, "<" & Session("UserID") & ">") > 0 And Not IsReconciliationCompleteWithOptions(S_ListToReconcile, S_ListReconciled) Then
		VAR_ShowVersionFilesForAll = "Y"
	End If
End If
'ph - 20120311 - end
%>