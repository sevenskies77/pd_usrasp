<%
'Place here your ASP-code running after user form input and AFTER PayDox database updating

'ph - 20090825 - start
'Собственная переменная для управления редиректом
SIT_bRedirect = False
'ph - 20090825 - end

'SAY 2009-03-12 - start - присвоение регистрационного номера при активации
'SAY 2009-06-05 алгоритм оставлен работающим для ранее созданных и неактивированных документов
If UCase(Request.ServerVariables("URL")) = UCase("/MakeActive.asp") Then 

'SAY 2009-06-02 добавлено условие на отдельные категории (Входящие, Поручения, Служебные)
'ниже есть действия для других категорий (не должны пересекаться)
   If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_VHODYASCHIE)) > 0 or InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_ZADACHI)) > 0 or InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_SLUZH_ZAPISKA)) > 0 Then
      If InStr(Request("DocID"),"PJ-") > 0 and Request("Active")<>"" Then

    Set dsTempPR1 = Server.CreateObject("ADODB.Recordset")
    sSQL = "SELECT * FROM Docs WHERE DocID=N'"+Request("DocID")+"'"
    dsTempPR1.Open sSQL, Conn, 3, 1, &H1

         If not dsTempPR1.EOF Then
'Определяем бизнес направление
' DmGorsky_5: меняем местами проверку: вначале ищем SIT_SITRU ("СИТРОНИКС ИТ"), потом, если не найдено, - SIT_SITRONICS ("СИТРОНИКС")
            If Instr(UCase(dsTempPR1("Department")), UCase(SIT_SITRU)) = 1 Then ' DmGorsky
 	           sDepartmentRoot = SIT_SITRU
            ElseIf InStr(UCase(dsTempPR1("Department")), UCase(SIT_STS)) = 1 Then
               sDepartmentRoot = SIT_STS
            ElseIf InStr(UCase(dsTempPR1("Department")), UCase(SIT_SITRONICS)) = 1 Then
               sDepartmentRoot = SIT_SITRONICS
            ElseIf InStr(UCase(dsTempPR1("Department")), UCase(SIT_RTI)) = 1 Then
               sDepartmentRoot = SIT_RTI
            ElseIf InStr(UCase(dsTempPR1("Department")), UCase(SIT_MIKRON)) = 1 Then
               sDepartmentRoot = SIT_MIKRON
            Else
               sDepartmentRoot = ""
         End If

'Ph - 20090330 - start 
'- первая из трех вставок для посылки ручного уведомления
'(также отключено автоматическое уведомление в UserSendNotificationList)
	     S_UserList = ""
'Ph - 20090330 - end

'1. Входящие документы
         If InStr(UCase(dsTempPR1("ClassDoc")), UCase(SIT_VHODYASCHIE)) > 0 Then
	        S_DocID = ""
            Call GetNewDocID_test(dsTempPR1("ClassDoc"), sDepartmentRoot, dsTempPR1("UserFieldText7"), "",  "", "")
         End If

'2. Поручения
         If InStr(UCase(dsTempPR1("ClassDoc")), UCase(SIT_ZADACHI)) > 0 Then
	        S_DocID = ""
            Call GetNewDocID_test(dsTempPR1("ClassDoc"), sDepartmentRoot, dsTempPR1("UserFieldText1"), "",  dsTempPR1("DocIDParent"), "")
'Ph - 20090330 - start
	S_UserList = dsTempPR1("NameControl")+"; "+dsTempPR1("NameResponsible")+"; "+dsTempPR1("Correspondent")
'Ph - 20090330 - end
      End If

'3. Служебная записка
         If InStr(UCase(dsTempPR1("ClassDoc")), UCase(SIT_SLUZH_ZAPISKA)) > 0 Then
            S_DocID = ""
            sClassDoc = "Служебные записки%"
            Call GetNewDocID_test(sClassDoc, sDepartmentRoot, dsTempPR1("UserFieldText1"), "",  "", "")
         End If
      End If

    dsTempPR1.Close

    If Trim(S_DocID) <> "" Then
'      Set dsTemp1 = Server.CreateObject("ADODB.Recordset")
'      dsTemp1.Open "select top 1 * from Docs where DocID = N'" & Request("DocID") & "'", Conn, 1, 3, &H1
'	   If not dsTemp1.EOF Then
'	     dsTemp1("DocID") = S_DocID
' 	     dsTemp1.Update
'	   End If
'	   dsTemp1.Close

dsDoc.Close
      Conn.Execute ("Update docs set docid=N'"+S_DocID+"' Where docid = N'"+Request("DocID")+"'")
dsDoc.Open "select * from Docs where DocID = N'" & Request("DocID") & "'", Conn, 3, 1, &H1
      'Conn.Execute ("Update docs set docidparent='"+S_DocID+"' Where docidparent = '"+Request("DocID")+"'")
      'Conn.Execute ("Update comments set docid='"+S_DocID+"' Where docid = '"+Request("DocID")+"'")
      
      Call oPayDox.ChangeDocIdInDepandants(Request("DocID"), S_DocID)

      If Request("Active") = "y" Then
	susercheck = "&usercheck=y"
      Else
	susercheck = ""
      End If
      'out GetURL("showdoc.asp", "?docid=", S_DocID)
'Ph - 20090330 - start
'При активации проектный номер заменяется на окончательный, поэтому рассылку стандартным
'способом делать нельзя, делаем вручную
         TempMessage = Session("Message")
         Session("Message") = ""
         S_MessageSubject = ""
	
         If S_UserList <> "" Then
            oPayDox.SendNotificationCore sMessageBody, Null, sS_Description, S_UserList, S_MessageSubject, _
                                         "", S_DocID, 0, "Yes", DOCS_KeyWordDepartment, DOCS_NOTFOUND, _
                                         DOCS_DocID, DOCS_DateActivation, DOCS_DateCompletion, DOCS_Name, _
                                         DOCS_PartnerName, DOCS_ACT, DOCS_Description, DOCS_Author, _
                                         DOCS_Correspondent, DOCS_Resolution, DOCS_NotificationSentTo, _
                                         DOCS_SendNotification, DOCS_UsersNotFound, DOCS_NotificationDoc, _
                                         USER_NOEMail, DOCS_NoAccess, DOCS_EXPIREDSEC, DOCS_STATUSHOLD, _
                                         VAR_StatusActiveUser, VAR_BeginOfTimes, VAR_ExtInt, DOCS_FROM1, _
                                         DOCS_Reconciliation, DOCS_NotificationNotCompletedDoc, _
                                         DOCS_ErrorSMTP, DOCS_Sender, DOCS_All, DOCS_NoReadAccess, _
                                         VAR_AdminSecLevel, False, MailTexts
	     End If

    If TempMessage <> "" Then
      Session("Message") = TempMessage+VbCrLf+Session("Message")
    End If
'Ph - 20090330 - end

      Response.Redirect GetURL("showdoc.asp", "?docid=", S_DocID) + "&template=" + Trim(Request("template")) + susercheck

    End If

  End If

  ' добавлено условие по категориям
  End IF
'SAY 2009-03-12 - end - присвоение регистрационного номера при активации

'Запрос №17 - СТС - start
  'Отправка уведомления получателям
  If Request("Active")<>"" Then
'{ ph - 20120601
'   If InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) = 1 or InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) = 1 Then
   If InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) = 1 or InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) = 1 or InStr(UCase(Session("CurrentClassDoc")), UCase(STS_SLUZH_ZAPISKA_HOLIDAY)) = 1 Then
'ph - 20120601 }
      If GetRootDepartment(Session("CurrentDepartmentDoc")) = SIT_STS Then
        S_UserList = ""
        Set dsTempPR1 = Server.CreateObject("ADODB.Recordset")
        sSQL = "select * from Docs where DocID = "&sUnicodeSymbol&"'"&Request("DocID")&"'"
        dsTempPR1.Open sSQL, Conn, 3, 1, &H1
        If not dsTempPR1.EOF Then
          S_UserList = Trim(MyCStr(dsTempPR1("ListToView")))
        End If
        dsTempPR1.Close

        If S_UserList <> "" Then
          TempMessage = Session("Message")
          Session("Message") = ""
          S_MessageSubject = DOCS_Active & " - " & GetFullName(SurnameGN(DelOtherLangFromFolder(Session("Name"))), Session("UserID"))
          sMessageBody = ""
          oPayDox.SendNotificationCore sMessageBody, Null, sS_Description, S_UserList, S_MessageSubject, "", Request("DocID"), 0, "Yes", DOCS_KeyWordDepartment, DOCS_NOTFOUND, DOCS_DocID, DOCS_DateActivation, DOCS_DateCompletion, DOCS_Name, DOCS_PartnerName, DOCS_ACT, DOCS_Description, DOCS_Author, DOCS_Correspondent, DOCS_Resolution, DOCS_NotificationSentTo, DOCS_SendNotification, DOCS_UsersNotFound, DOCS_NotificationDoc, USER_NOEMail, DOCS_NoAccess, DOCS_EXPIREDSEC, DOCS_STATUSHOLD, VAR_StatusActiveUser, VAR_BeginOfTimes, VAR_ExtInt, DOCS_FROM1, DOCS_Reconciliation, DOCS_NotificationNotCompletedDoc, DOCS_ErrorSMTP, DOCS_Sender, DOCS_All, DOCS_NoReadAccess, VAR_AdminSecLevel, False, MailTexts
          If TempMessage <> "" Then
            Session("Message") = TempMessage+VbCrLf+Session("Message")
          End If
        End If
      End If
    End If
  End If
'Запрос №17 - СТС - end
'20100111 - Запрос №13 из СТС - start
  'Рассылка уведомленй ответственным при активации
'ph - 20100712 - start
'  If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI)) > 0 or InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) > 0 Then
  If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) = 1 or InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) = 1 Then
'ph - 20100712 - end
    If Request("Active") <> "" Then
      S_UserList = SIT_GetDocField(Request("DocID"), "NameResponsible", Conn)
      TempMessage = Session("Message")
      Session("Message") = ""
      If S_UserList <> "" Then
        sMessageBody = STS_YouAreResponsibleBody
        sMessageBody = "<b>" & sMessageBody & "</b>" & "<br>"
        oPayDox.SendNotificationCore sMessageBody, Null, sS_Description, S_UserList, STS_YouAreResponsibleSubject, "", Request("DocID"), 0, "Yes", DOCS_KeyWordDepartment, DOCS_NOTFOUND, DOCS_DocID, DOCS_DateActivation, DOCS_DateCompletion, DOCS_Name, DOCS_PartnerName, DOCS_ACT, DOCS_Description, DOCS_Author, DOCS_Correspondent, DOCS_Resolution, DOCS_NotificationSentTo, DOCS_SendNotification, DOCS_UsersNotFound, DOCS_NotificationDoc, USER_NOEMail, DOCS_NoAccess, DOCS_EXPIREDSEC, DOCS_STATUSHOLD, VAR_StatusActiveUser, VAR_BeginOfTimes, VAR_ExtInt, DOCS_FROM1, DOCS_Reconciliation, DOCS_NotificationNotCompletedDoc, DOCS_ErrorSMTP, DOCS_Sender, DOCS_All, DOCS_NoReadAccess, VAR_AdminSecLevel, False, MailTexts
      End If
      If TempMessage <> "" Then
        Session("Message") = TempMessage & VbCrLf & Session("Message")
      End If
      Response.Redirect GetURL("showdoc.asp", "?docid=", Request("DocID")) & "&usercheck=y&template=" & Trim(Request("template"))
    End If
  End If
'20100111 - Запрос №13 из СТС - end
End If




'Ph - 20080918 - start - Уведомление
'vnik
vnik_SIT_NotificationDocReconciled4 = ""
'vnik
'Согласование и Утверждение
'Одобрение или отказ в утверждении
If UCase(Request.ServerVariables("URL")) = UCase("/Visa.asp") Then
   If UCase(Trim(Request("app")))<>"Y" Then 'СОГЛАСОВАНИЕ
      If Trim(Request("r")) <> "y" Then 'Согласовано
         TempMessage = Session("Message")
         Session("Message") = ""

      'Отправляем отдельно уведомление инициатору при согласовании
      If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_ISHODYASCHIE)) > 0 Then
        S_UserList = SIT_GetDocField(Request("DocID"), "NameCreation", Conn) 'Выполнить рассылку создателю (инициатор)
      ElseIf InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) > 0 or InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) > 0 Then
'Ph - 20081209 - Инициатор переехал в создателя и автора
'        S_UserList = SIT_GetDocField(Request("DocID"), "NameResponsible", Conn) 'Выполнить рассылку ответственному (инициатор)
        S_UserList = SIT_GetDocField(Request("DocID"), "Author", Conn)
'20090622 - Заявка ТКП
         ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_COM_OFFERS)) = 1 Then
'Уведомление списку ознакомления(менеджер) и автору
'20090731 - start 
'- Добавлено условие - рассылка только если согласование завершено
            Set dsTemp1 = Server.CreateObject("ADODB.Recordset")
            sSQL = "select * from Docs where DocID=N'"+Request("DocID")+"'"
            dsTemp1.Open sSQL, Conn, 3, 1, &H1
    
        S_UserList = "-"
        If not dsTemp1.EOF Then
          If IsReconciliationComplete(dsTemp1("ListToReconcile"), dsTemp1("ListReconciled")) Then
            S_UserList = dsTemp1("Author")+dsTemp1("ListToView")
          Else
            S_UserList = dsTemp1("Author")
          End If
        End If
        dsTemp1.Close
        '20090731 - end
	ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_CONTRACTS_MC)) = 1 Then
            Set vnikdsTemp1 = Server.CreateObject("ADODB.Recordset")
            sSQL = "SELECT * FROM Docs where DocID = " +sUnicodeSymbol+"'"+Request("DocID")+"'"
            AddLogD "vnik456 " + Trim(sSQL)
            vnikdsTemp1.Open sSQL, Conn, 3, 1, &H1
            If not vnikdsTemp1.EOF Then
                If IsReconciliationComplete(vnikdsTemp1("ListToReconcile"), vnikdsTemp1("ListReconciled")) Then
                    S_UserList = vnikdsTemp1("Author")+vnikdsTemp1("LocationPath")
                    vnik_SIT_NotificationDocReconciled4 = SIT_NotificationDocReconciled4
                Else
                    S_UserList = vnikdsTemp1("Author")
                End If
            End If
	        vnikdsTemp1.Close
         Else
'Выполнить рассылку автору (инициатор)
            S_UserList = SIT_GetDocField(Request("DocID"), "Author", Conn)
         End If

'Если документ согласован с комментариями, то в теме уведомления автору документа в
'конце добавляем " с комментариями "
'20090827 - vnik - start 
         If Trim(Request("reason")) <> "" Then
            S_MessageSubject=SIT_NotificationDocReconciled1 + Request("DocID") + SIT_NotificationDocReconciled2 + GetFullName(SurnameGN(DelOtherLangFromFolder(Session("Name"))), Session("UserID")) + SIT_NotificationDocReconciled3
         Else
'20090827 - vnik - end
            S_MessageSubject=SIT_NotificationDocReconciled1 + Request("DocID") + SIT_NotificationDocReconciled2 + GetFullName(SurnameGN(DelOtherLangFromFolder(Session("Name"))), Session("UserID"))
'20090827 - vnik - start
         End If
'vnik
         S_MessageSubject=S_MessageSubject + vnik_SIT_NotificationDocReconciled4
'vnik
'20090827 - vnik - end
      
'Запрос №17 - СТС - start
      If GetRootDepartment(Session("CurrentDepartmentDoc")) = SIT_STS Then
        If Trim(Request("reason")) <> "" Then
          sMessageBody = NotificationWithReason(DOCS_Reconciled, Request("reason"))
        End If
      End If
'Запрос №17 - СТС - end
      oPayDox.SendNotificationCore sMessageBody, Null, sS_Description, S_UserList, S_MessageSubject, "", Request("DocID"), 0, "Yes", DOCS_KeyWordDepartment, DOCS_NOTFOUND, DOCS_DocID, DOCS_DateActivation, DOCS_DateCompletion, DOCS_Name, DOCS_PartnerName, DOCS_ACT, DOCS_Description, DOCS_Author, DOCS_Correspondent, DOCS_Resolution, DOCS_NotificationSentTo, DOCS_SendNotification, DOCS_UsersNotFound, DOCS_NotificationDoc, USER_NOEMail, DOCS_NoAccess, DOCS_EXPIREDSEC, DOCS_STATUSHOLD, VAR_StatusActiveUser, VAR_BeginOfTimes, VAR_ExtInt, DOCS_FROM1, DOCS_Reconciliation, DOCS_NotificationNotCompletedDoc, DOCS_ErrorSMTP, DOCS_Sender, DOCS_All, DOCS_NoReadAccess, VAR_AdminSecLevel, False, MailTexts

      If TempMessage <> "" Then
        Session("Message") = TempMessage+VbCrLf+Session("Message")
      End If

    End If
  End If
End If
'Ph - 20080918 - end

'ph - 20090717 - start
'Отмена статуса Отменено
If UCase(Request.ServerVariables("URL")) = UCase("/MakeCanceledCancel.asp") Then
   If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_COM_OFFERS)) > 0 Then
      Set dsTempPR1 = Server.CreateObject("ADODB.Recordset")
      sSQL = "select * from Docs where DocID=N'"+Request("DocID")+"'"
      dsTempPR1.Open sSQL, Conn, 3, 1, &H1
    
      If dsTempPR1.EOF Then
         S_UserList = "-"
      Else
         'Инициатору и менеджеру (список ознакомления) и согласующим уведомления всегда
         S_UserList = dsTempPR1("Author")+";"+dsTempPR1("ListToView")+";"+dsTempPR1("ListToReconcile")
         'Если согласование закончено, то уведомление еще и утверждающему
         If IsReconciliationComplete(dsTempPR1("ListToReconcile"), dsTempPR1("ListReconciled")) Then
            S_UserList = S_UserList+";"+dsTempPR1("NameAproval")
         End If
      End If
      dsTempPR1.Close
      If S_UserList <> "-" Then
         TempMessage = Session("Message")
         Session("Message") = ""
         S_MessageSubject= DOCS_CancellationCancelled
         oPayDox.SendNotificationCore sMessageBody, Null, sS_Description, S_UserList, S_MessageSubject, "", Request("DocID"), 0, "Yes", DOCS_KeyWordDepartment, DOCS_NOTFOUND, DOCS_DocID, DOCS_DateActivation, DOCS_DateCompletion, DOCS_Name, DOCS_PartnerName, DOCS_ACT, DOCS_Description, DOCS_Author, DOCS_Correspondent, DOCS_Resolution, DOCS_NotificationSentTo, DOCS_SendNotification, DOCS_UsersNotFound, DOCS_NotificationDoc, USER_NOEMail, DOCS_NoAccess, DOCS_EXPIREDSEC, DOCS_STATUSHOLD, VAR_StatusActiveUser, VAR_BeginOfTimes, VAR_ExtInt, DOCS_FROM1, DOCS_Reconciliation, DOCS_NotificationNotCompletedDoc, DOCS_ErrorSMTP, DOCS_Sender, DOCS_All, DOCS_NoReadAccess, VAR_AdminSecLevel, False, MailTexts

      If TempMessage <> "" Then
        Session("Message") = TempMessage+VbCrLf+Session("Message")
      End If
    End If
  End If
End If

'Отправка письма Менеджеру при редактировании с повторным согласованием 
' - НЕ ПРОШЛО ИЗ-ЗА РЕДИРЕКТА
'If UCase(Request.ServerVariables("URL")) = UCase("/ChangeDoc.asp") Then
'  If UCase(Request("UpdateDoc")) = "YES" Then
'    If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_COM_OFFERS)) > 0 Then
'      If UCase(Request("ClearStatusReconciled")) = "ON" Then
'        TempMessage = Session("Message")
'        Session("Message") = ""
'        S_UserList = SIT_GetDocField(Request("DocID"), "ListToView", Conn) 'Выполнить рассылку менеджеру (в списке ознакомления)
'        S_MessageSubject=DOCS_AGREERepeate
'        oPayDox.SendNotificationCore sMessageBody, Null, sS_Description, S_UserList, S_MessageSubject, "", Request("DocID"), 0, "Yes", DOCS_KeyWordDepartment, DOCS_NOTFOUND, DOCS_DocID, DOCS_DateActivation, DOCS_DateCompletion, DOCS_Name, DOCS_PartnerName, DOCS_ACT, DOCS_Description, DOCS_Author, DOCS_Correspondent, DOCS_Resolution, DOCS_NotificationSentTo, DOCS_SendNotification, DOCS_UsersNotFound, DOCS_NotificationDoc, USER_NOEMail, DOCS_NoAccess, DOCS_EXPIREDSEC, DOCS_STATUSHOLD, VAR_StatusActiveUser, VAR_BeginOfTimes, VAR_ExtInt, DOCS_FROM1, DOCS_Reconciliation, DOCS_NotificationNotCompletedDoc, DOCS_ErrorSMTP, DOCS_Sender, DOCS_All, DOCS_NoReadAccess, VAR_AdminSecLevel, False, MailTexts
'        If TempMessage <> "" Then
'          Session("Message") = TempMessage+VbCrLf+Session("Message")
'        End If
'      End If
'    End If
'  End If
'End If
'ph - 20090717 - end

'Ph - 20081209 - start - 
'В исходящем документе получатель не получает уведомление после регистрации
'  (получатель из СТС)
If UCase(Request.ServerVariables("URL")) = UCase("/MakeRegisteredRet.asp") Then
  If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_ISHODYASCHIE)) > 0 Then
    TempMessage = Session("Message")
    Session("Message") = ""

    S_UserList = GetCorrectUsersFromList(SIT_GetDocField(Request("DocID"), "Correspondent", Conn), Conn)
    S_MessageSubject = SIT_NotificationLetterForYou
	If S_UserList <> "" Then
      oPayDox.SendNotificationCore sMessageBody, Null, sS_Description, S_UserList, S_MessageSubject, "", Request("DocID"), 0, "Yes", DOCS_KeyWordDepartment, DOCS_NOTFOUND, DOCS_DocID, DOCS_DateActivation, DOCS_DateCompletion, DOCS_Name, DOCS_PartnerName, DOCS_ACT, DOCS_Description, DOCS_Author, DOCS_Correspondent, DOCS_Resolution, DOCS_NotificationSentTo, DOCS_SendNotification, DOCS_UsersNotFound, DOCS_NotificationDoc, USER_NOEMail, DOCS_NoAccess, DOCS_EXPIREDSEC, DOCS_STATUSHOLD, VAR_StatusActiveUser, VAR_BeginOfTimes, VAR_ExtInt, DOCS_FROM1, DOCS_Reconciliation, DOCS_NotificationNotCompletedDoc, DOCS_ErrorSMTP, DOCS_Sender, DOCS_All, DOCS_NoReadAccess, VAR_AdminSecLevel, False, MailTexts
	End If

    If TempMessage <> "" Then
      Session("Message") = TempMessage+VbCrLf+Session("Message")
    End If
  End If
End If
'Ph - 20081209 - end

'отправка уведомления Наливных И.В. при согласовании заявки Заболотневой М.В.
'заявка на оплату
If InStr(UCase(Session("CurrentClassDoc")), UCase(RTI_Payment_Order)) > 0 Then
  'была нажата кнопка "согласовать"
  If UCase(Request.ServerVariables("URL")) = UCase("/Visa.asp") and Trim(Request("app")) <> "y" and Trim(Request("r")) <> "y" Then
    'пользователь Заболотнева М.В. и автор заявки не наливных и.в.
    If Session("UserId") = "mzabolotneva_oaorti" and SIT_GetDocField(Request("DocID"), "Author", Conn) <> """Наливных И. В."" <inalivnyh_oaorti>;" Then
      'addlogD "exor777 отправляем уведомление Наливных И.В. если статья расходов к нему относится, а также предоставляем доступ к заявке на оплату (если необходимо)"
      'проверяем, что статья расходов относится к Игорю Наливных
      Set dsTemp1 = Server.CreateObject("ADODB.Recordset")
      sSQL = "select UserFieldText4 from docs where DocID = " +sUnicodeSymbol+"'" + Trim(Request("DocID")) + "'"+" and UserFieldText4 in (select name from rti_costitem2 where Responsible like N'%inalivnyh_oaorti%')"
      dsTemp1.Open sSQL, Conn, 3, 1, &H1
      If Not dsTemp1.EOF Then
      'если статья расходов относится к игорю наливных, то проверяем, что у него есть доступ к заявке (поле "получатели")
        dsTemp1.Close
        sSQL = "select ListToView from docs where DocID = " +sUnicodeSymbol+"'" + Request("DocID") + "'"+" and ListToView like N'%inalivnyh_oaorti%'"
        dsTemp1.Open sSQL, Conn, 3, 1, &H1
        'если доступа нет, то перед отправкой уведомления предоставляем доступ (поле "получатели" + поле "дополнительные пользователи")      
        varIsListToViewUser = ""
        varAdditionalUser = ""
        If dsTemp1.EOF Then
          varListToViewUser = "N"
        End If
        dsTemp1.Close
        sSQL = "select AdditionalUsers from docs where DocID = " +sUnicodeSymbol+"'" + Request("DocID") + "'"+" and AdditionalUsers like N'%inalivnyh_oaorti%'"
        dsTemp1.Open sSQL, Conn, 3, 1, &H1
        If dsTemp1.EOF Then
          varAdditionalUser = "N"
        End If
        'добавляем в получатели
        If varListToViewUser = "N" Then
          Conn.Execute ("update docs set ListToView = ListToView + N' ""Наливных И. В."" <inalivnyh_oaorti>;' where DocID = " +sUnicodeSymbol+"'"  + Trim(Request("DocID")) + "'")
        End If
        'добавляем в AdditionalUsers (чтобы документ был виден в системе до утверждения)
        If varAdditionalUser = "N" Then
          Conn.Execute ("update docs set AdditionalUsers = AdditionalUsers + N' ""Наливных И. В."" <inalivnyh_oaorti>;' where DocID = " +sUnicodeSymbol+"'"  + Trim(Request("DocID")) + "'")
        End If
        'отправка уведомления
        TempMessage = Session("Message")
        Session("Message") = ""
        S_UserList = """Наливных И. В."" <inalivnyh_oaorti>;"     
        S_MessageSubject = "Заявка на оплату " +Request("DocID")+ " согласована и ожидает оплаты"' + varAccessGranted
        oPayDox.SendNotificationCore sMessageBody, Null, sS_Description, S_UserList, S_MessageSubject, _
             "", Request("DocID"), 0, "Yes", DOCS_KeyWordDepartment, DOCS_NOTFOUND, DOCS_DocID, DOCS_DateActivation, _
            DOCS_DateCompletion, DOCS_Name, DOCS_PartnerName, DOCS_ACT, DOCS_Description, DOCS_Author, DOCS_Correspondent, _
            DOCS_Resolution, DOCS_NotificationSentTo, DOCS_SendNotification, DOCS_UsersNotFound, DOCS_NotificationDoc, _
            USER_NOEMail, DOCS_NoAccess, DOCS_EXPIREDSEC, DOCS_STATUSHOLD, VAR_StatusActiveUser, VAR_BeginOfTimes, _
            VAR_ExtInt, DOCS_FROM1, DOCS_Reconciliation, DOCS_NotificationNotCompletedDoc, DOCS_ErrorSMTP, _
            DOCS_Sender, DOCS_All, DOCS_NoReadAccess, VAR_AdminSecLevel, False, MailTexts
        If TempMessage <> "" Then
            Session("Message") = TempMessage+VbCrLf+Session("Message")
        End If        
      End If
      dsTemp1.Close     
    End If
  End If
End If
'end if отправка уведомления Наливных И.В. при согласовании заявки Заболотневой М.В.


'Ph - 20081212 - start
If UCase(Request.ServerVariables("URL")) = UCase("/MakeActive.asp") and Request("Active")<>"" Then
  If InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) > 0 Then
    sDocIDParent = SIT_GetDocField(Request("DocID"), "DocIDParent", Conn)
    Set dsTemp1 = Server.CreateObject("ADODB.Recordset")
    sSQL = "select * from Docs where DocID = "+sUnicodeSymbol+"'" + sDocIDParent + "'"
AddLogD "@@@DBUpdateAfter SearchingParentDoc SQL: "+sSQL
      dsTemp1.Open sSQL, Conn, 3, 1, &H1
      If Not dsTemp1.EOF Then
        nPurchaseOrderAmountUSD = CCur(dsTemp1("UserFieldMoney1"))
		dsTemp1.Close
        sSQL = "select IsNull(Sum(UserFieldMoney1), 0) as SumUSD from Docs where DocIDParent = "+sUnicodeSymbol+"'" + sDocIDParent + "' and ClassDoc like "+sUnicodeSymbol+"'"+STS_PaymentOrder+"%' and IsActive = 'Y'"
AddLogD "@@@DBUpdateAfter SumOfChildPaymentOrders SQL: "+sSQL
        dsTemp1.Open sSQL, Conn, 3, 1, &H1
		If CCur(dsTemp1("SumUSD")) > nPurchaseOrderAmountUSD Then
          Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorSumExceeding)
		End If
	  End If
	  dsTemp1.Close
  End If
End If
'Ph - 20081212 - end

'Ph - 20090123 - start
'Если изменился утверждающий на этапе утверждения, делаем рассылку
'(Sitronics_OldAproval устанавливается в DBUpdateBefore)
If UCase(Request.ServerVariables("URL")) = UCase("/ChangeDoc.asp") or UCase(Request.ServerVariables("URL")) = UCase("/ChangeField.asp")Then
  If Request("create") <> "y" and UCase(Request("UpdateDoc")) = "YES" Then
    If Sitronics_OldAproval <> "" Then
AddLogD "Sitronics_OldAproval: "+Sitronics_OldAproval
      If GetUserID(S_NameAproval) <> GetUserID(Sitronics_OldAproval) Then
        Set dsTemp2 = Server.CreateObject("ADODB.Recordset")
        sSQL = "Select * from Docs where DocID = "+sUnicodeSymbol+"'"+Request("DocID")+"'"
        dsTemp2.Open sSQL, Conn, 3, 1, &H1
        If not dsTemp2.EOF Then
'          bApprovalRequired = MyCStr(dsTemp2("NameApproved")) = "" and IsReconciliationComplete(dsTemp2("ListToReconcile"), dsTemp2("ListReconciled"))
'{ph - 20120607
'          bApprovalRequired = dsTemp2("IsActive") = "Y" and MyCStr(dsTemp2("NameApproved")) = "" and IsReconciliationComplete(dsTemp2("ListToReconcile"), dsTemp2("ListReconciled"))
          bApprovalRequired = dsTemp2("IsActive") = "Y" and MyCStr(dsTemp2("NameApproved")) = "" and IsReconciliationCompleteWithOptions(dsTemp2("ListToReconcile"), dsTemp2("ListReconciled"))
'ph - 20120607}
        End If
        dsTemp2.Close
        
'Документ находится на стадии утверждения
            If bApprovalRequired Then 
               sSessionMessage = Session("Message")
               Session("Message") = ""

          S_UserList = Sitronics_OldAproval
          S_MessageSubject = SIT_NotificationApprovalChanged1 + Sitronics_OldAproval + SIT_NotificationApprovalChanged2 + S_NameAproval
          oPayDox.SendNotificationCore sMessageBody, Null, sS_Description, S_UserList, S_MessageSubject, "", Request("DocID"), 0, "Yes", DOCS_KeyWordDepartment, DOCS_NOTFOUND, DOCS_DocID, DOCS_DateActivation, DOCS_DateCompletion, DOCS_Name, DOCS_PartnerName, DOCS_ACT, DOCS_Description, DOCS_Author, DOCS_Correspondent, DOCS_Resolution, DOCS_NotificationSentTo, DOCS_SendNotification, DOCS_UsersNotFound, DOCS_NotificationDoc, USER_NOEMail, DOCS_NoAccess, DOCS_EXPIREDSEC, DOCS_STATUSHOLD, VAR_StatusActiveUser, VAR_BeginOfTimes, VAR_ExtInt, DOCS_FROM1, DOCS_Reconciliation, DOCS_NotificationNotCompletedDoc, DOCS_ErrorSMTP, DOCS_Sender, DOCS_All, DOCS_NoReadAccess, VAR_AdminSecLevel, False, MailTexts

          If Session("Message") <> "" Then
            sSessionMessage = sSessionMessage + "<BR>" + Session("Message")
            Session("Message") = ""
          End If
          S_UserList = S_NameAproval
          S_MessageSubject = DOCS_APROVAL
          oPayDox.SendNotificationCore sMessageBody, Null, sS_Description, S_UserList, S_MessageSubject, "", Request("DocID"), 0, "Yes", DOCS_KeyWordDepartment, DOCS_NOTFOUND, DOCS_DocID, DOCS_DateActivation, DOCS_DateCompletion, DOCS_Name, DOCS_PartnerName, DOCS_ACT, DOCS_Description, DOCS_Author, DOCS_Correspondent, DOCS_Resolution, DOCS_NotificationSentTo, DOCS_SendNotification, DOCS_UsersNotFound, DOCS_NotificationDoc, USER_NOEMail, DOCS_NoAccess, DOCS_EXPIREDSEC, DOCS_STATUSHOLD, VAR_StatusActiveUser, VAR_BeginOfTimes, VAR_ExtInt, DOCS_FROM1, DOCS_Reconciliation, DOCS_NotificationNotCompletedDoc, DOCS_ErrorSMTP, DOCS_Sender, DOCS_All, DOCS_NoReadAccess, VAR_AdminSecLevel, False, MailTexts
          Sitronics_OldAproval = ""

          If Session("Message") <> "" Then
            Session("Message") = sSessionMessage + "<BR>" + Session("Message")
          End If
        End If
      End If
    End If
'ph - 20090825 - start
'      Response.Redirect GetURL("ShowDoc.asp", "?docid=", S_DocID) + "&justmodified=" + CStr(Time())
'Редирект делаем потом, чтобы отработал код ниже
      SIT_bRedirect = True
'ph - 20090825 - end
  End If
End If
'Ph - 20090123 - end

'ph - 20090825 - start
If UCase(Request.ServerVariables("URL")) = UCase("/ChangeDoc.asp") Then
  If UCase(Request("UpdateDoc")) = "YES" Then
    If InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) > 0 or InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) > 0 Then
      If Trim(SIT_SessionMessageSaveBeforeSendNotitfication) <> "" Then
        'Восстанавливаем сообщение, чтобы показать пользователям
        Session("Message") = DOCS_Changed + SIT_SessionMessageSaveBeforeSendNotitfication + "<br>" + Replace(Session("Message"), DOCS_Changed, "")
        SIT_bRedirect = True
      End If
    End If
  End If
End If

'Собственно сам редирект
If SIT_bRedirect Then
  Response.Redirect GetURL3("showdoc.asp", "?docid=", S_DocID, "&docidparent=", S_DocIDParent, "&ClassDoc=", S_ClassDoc) + IIf(Request("create") = "y", "&justcreated=", "&justmodified=") + CStr(Time()) + "&template=" + Trim(Request("template"))
End If
'ph - 20090825 - end

'ph - 20111220 - start
	If UCase(Request.ServerVariables("URL")) = UCase("/CreateComment.asp") Then
		If InStr(Session("CurrentClassDoc"), SIT_DOGOVORI_NEW) = 1 and InStr(Session("Department"), SIT_STS_ROOT_DEPARTMENT) = 1 Then
			If UCase(Request("CommentType")) = "REVIEW" Then
				If UCase(Request("reviewtype")) = "NEGATIVE" Then
					TempMessage = Session("Message")
					Session("Message") = ""

					S_UserList = SIT_GetDocField(Request("DocID"), "NameCreation", Conn)
					S_MessageSubject = "NEGATIVE review obtained"
					If S_UserList <> "" Then
					  oPayDox.SendNotificationCore sMessageBody, Null, sS_Description, S_UserList, S_MessageSubject, "", Request("DocID"), 0, "Yes", DOCS_KeyWordDepartment, DOCS_NOTFOUND, DOCS_DocID, DOCS_DateActivation, DOCS_DateCompletion, DOCS_Name, DOCS_PartnerName, DOCS_ACT, DOCS_Description, DOCS_Author, DOCS_Correspondent, DOCS_Resolution, DOCS_NotificationSentTo, DOCS_SendNotification, DOCS_UsersNotFound, DOCS_NotificationDoc, USER_NOEMail, DOCS_NoAccess, DOCS_EXPIREDSEC, DOCS_STATUSHOLD, VAR_StatusActiveUser, VAR_BeginOfTimes, VAR_ExtInt, DOCS_FROM1, DOCS_Reconciliation, DOCS_NotificationNotCompletedDoc, DOCS_ErrorSMTP, DOCS_Sender, DOCS_All, DOCS_NoReadAccess, VAR_AdminSecLevel, False, MailTexts
					End If

					If TempMessage <> "" Then
					  Session("Message") = TempMessage+VbCrLf+Session("Message")
					End If
				End If
			End If
		End If
	End If
'ph - 20111220 - end

%>