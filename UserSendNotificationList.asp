<%
'AddLogD "sNotificationList *:"+sNotificationList
'AddLogD "sNotificationSubject *:"+sNotificationSubject

'Define your own e-mail notification list for various events
'
'sNotificationSubject - e-mail notification subject
'sNotificationList - e-mail notification list 
'
'C - document creator
'U - document author
'R - responsible person
'L - controlling person
'A - approving person
'V - viewers list
'D - correspondents, notification list
'E - reconciliation (agree) list - not agreed yet
'G - reconciliation (agree) list - agreed
'S - registrars list
'
'sNotificationList="CRLAVDE"

'ph - 20090825 - start
SIT_SessionMessageSaveBeforeSendNotitfication = ""
'ph - 20090825 - end

AddLogD "UserSendNotificationList.asp URL:"+UCase(Request.ServerVariables("URL"))
'AddLogD UCase(Request.ServerVariables("URL"))=UCase("/Visa.asp")

sNotificationList=""
sNotificationSubject=""

Select Case UCase(Request.ServerVariables("URL"))
	Case UCase("/MakeActive.asp") ' make document status as Active or Inactive
		If Request("Active")<>"" Then 'make document status as Active 
'phil - 20080918 - start
			If InStr(UCase(Request("ClassDoc")), UCase(SIT_ZADACHI)) = 1 Then
              			If InStr(Request("DocID"),"PJ-") > 0 Then '20090330 - добавлено условие, чтобы при перебивке номера при активации уведомления здесь не рассылались
			    		sNotificationList="-"
		      		Else
					sNotificationList="RLD" 'Уведомление исполнителю, контролеру, Correspondent (соисполнителям)
'					sNotificationList="RLV" 'Уведомление исполнителю, контролеру, Viewers list (соисполнителям)
'amw                  If GetRootDepartment(Session("CurrentDepartmentDoc"))=SIT_SITRONICS or GetRootDepartment(Session("CurrentDepartmentDoc"))=SIT_RTI Then
                  If GetRootDepartment(Session("CurrentDepartmentDoc"))=SIT_SITRONICS or GetRootDepartment(Session("CurrentDepartmentDoc"))=SIT_RTI or GetRootDepartment(Session("CurrentDepartmentDoc"))=SIT_MIKRON or GetRootDepartment(Session("CurrentDepartmentDoc"))=SIT_VTSS Then
                      sNotificationSubject = "Поручение выдано"
                      DOCS_Active = "Поручение выдано"
                    end if

              			End If
			End If
'20090622 - Заявка ТКП
			If InStr(UCase(Request("ClassDoc")), UCase(SIT_COM_OFFERS)) = 1 Then
				sNotificationList="VE" 'Уведомление списку ознакомления(менеджер), согласующим
			End If
'phil - 20080918 - start

			'SAY 2009-06-02 рассылка при активации (для старых документов)
			If InStr(UCase(Request("ClassDoc")), UCase(SIT_VHODYASCHIE)) = 1 or InStr(UCase(Request("ClassDoc")), UCase(SIT_SLUZH_ZAPISKA)) = 1 Then
              			If InStr(Request("DocID"),"PJ-") > 0 Then '20090330 - добавлено условие, чтобы при перебивке номера при активации уведомления здесь не рассылались
			    		sNotificationList="-"
              			End If
			End If
			
			If InStr(UCase(Session("CurrentClassDoc")),UCase(RTI_PAYMENT_ORDER)) > 0  Then 'and UCase(Request("UserFieldText4")) = "НЕТ" Then
			   	
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
                    sNotificationSubject = "Активен - требует согласования (ДОГОВОР НЕ СДАН НА ХРАНЕНИЕ В БУХГАЛТЕРИЮ!)"
                End If
			End If
			'SAY End
			
			'vnik_payment_order
            'If InStr(UCase(Request("ClassDoc")),UCase(SIT_PAYMENT_ORDER)) > 0 Then
            '    sNotificationList="VE"       
            'End If
            'vnik_payment_order

			'rti_payment_order
            'If InStr(UCase(Request("ClassDoc")),UCase(RTI_PAYMENT_ORDER)) > 0 Then
            '    sNotificationList="E"       
            'End If
            'rti_payment_order
            'vnik_purchase_order
            If InStr(UCase(Request("ClassDoc")),UCase(SIT_PURCHASE_ORDER)) > 0 Then
                sNotificationList="VE"       
            End If
            'vnik_purchase_order
            
            'rti_purchase_order
            'If InStr(UCase(Request("ClassDoc")),UCase(RTI_PURCHASE_ORDER)) > 0 Then
            '    sNotificationList="E"       
            'End If
            'vnik_purchase_order
            'mikron_purchase_order
            If InStr(UCase(Request("ClassDoc")),UCase(MIKRON_PURCHASE_ORDER)) > 0 Then
                sNotificationList="VE"
            End If
'mikron_purchase_order
'08-06-2015
            If InStr(UCase(Request("ClassDoc")),UCase(SIT_ISHODYASCHIE)) > 0 and GetRootDepartment(Session("CurrentDepartmentDoc")) = SIT_RTI Then
                sNotificationList="E"       
                sNotificationSubject = "Активен - требует согласования"
            End If
'08-06-2015

'SAY 2008-10-07
AddLogD "@2@:" + Request("ClassDoc") + ", "+ Session("CurrentClassDoc")+", "+Session("ClassDoc")
			If InStr(UCase(Request("ClassDoc")), UCase(SIT_VHODYASCHIE)) = 1 Then
				'sNotificationList="V" 'Уведомление исполнителю, контролеру, Viewers list (соисполнителям)
			End If


		Else								'make document status as InActive 
		End If
	'	If InStr(UCase(Request("ClassDoc")), "INVOICES")<=0 Then 
   ' 		sNotificationList="R" 
	'	Else 
   '		sNotificationList="-" 
	'	End If
	'		sNotificationList="CRLAVDE"
	Case UCase("/Visa.asp") 'Согласование и Утверждение
		'Одобрение или отказ в утверждении
	
		If Trim(Request("app"))="y" Then 'УТВЕРЖДЕНИЕ
			sNotificationList=""
			sNotificationSubject=""
			
			If GetRootDepartment(Session("CurrentDepartmentDoc")) = SIT_RTI Then
			sNotificationList="A"
			
			end if
			
			If Trim(Request("r")) <> "y" Then 'Утвержден
				sNotificationList="CUS" 'Выполнить рассылку создателю, автору и регистратору
                                'SAY 2008-10-07
'   				If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_SLUZH_ZAPISKA)) = 1 Then
   				If InStr(UCase(Request("ClassDoc")), UCase(SIT_SLUZH_ZAPISKA)) = 1 Then
					sNotificationList=sNotificationList+"V" 'Уведомление исполнителю, контролеру, Viewers list (получателю)
                                        'SAY 2008-10-23 адресату
                                        'sNotificationList=sNotificationList+"D" 'Уведомление адресату
				End If
'20090622 - Заявка ТКП
                If InStr(UCase(Request("ClassDoc")), UCase(SIT_COM_OFFERS)) = 1 Then
                    sNotificationList="VC" 'Уведомление списку ознакомления(менеджер), автору
                End If

                                'sNotificationList=sNotificationList+"D" 'Уведомление адресату

                'vnik_purchase_order
                If InStr(UCase(Request("ClassDoc")),UCase(SIT_PURCHASE_ORDER)) > 0 Then
                    sNotificationList="UL"       
                    sNotificationSubject = "Заявка на закупку утверждена, если сумма заявки больше 300 000 рублей, то Секретарь ЦЗК инциирует протокол ЦЗК."
                End If
                'vnik_purchase_order
                
                'rti_purchase_order
                If InStr(UCase(Request("ClassDoc")),UCase(RTI_PURCHASE_ORDER)) > 0 Then
                    sNotificationList="ULV"       
                    sNotificationSubject = "Заявка на закупку утверждена"
                End If
                'rti_purchase_order
'mikron_purchase_order
               If InStr(UCase(Request("ClassDoc")),UCase(MIKRON_PURCHASE_ORDER)) > 0 Then
                  sNotificationList="UL"
                  sNotificationSubject = "Заявка на закупку утверждена, если сумма заявки больше 300 000 рублей, требуется Протокол ЗК."
               End If
'mikron_purchase_order
'rti_payment_order
               If InStr(UCase(Request("ClassDoc")),UCase(RTI_PAYMENT_ORDER)) > 0 Then
                   sNotificationList="UV"       
                   sNotificationSubject = "Оплачено"
               End If
'rti_payment_order	
               If InStr(UCase(Request("ClassDoc")),UCase(RTI_BSAP)) > 0 Then
                   sNotificationList="UV"       
               End If
               If InStr(UCase(Request("ClassDoc")),UCase(RTI_CONTRACT)) > 0 Then
                   sNotificationList="UV"       
               End If
               sNotificationSubject="" '"Требуется регистрация документа"
			Else 'Отказано, Не утвержден	
'Запрос №17 - СТС - start
				sNotificationBody = DOCS_RefusedApp
'Запрос №17 - СТС - end
'20090622 - Заявка ТКП
                If InStr(UCase(Request("ClassDoc")), UCase(SIT_COM_OFFERS)) = 1 Then
                    sNotificationList="VC" 'Уведомление списку ознакомления(менеджер), автору
                Else
                    sNotificationList="C" 'Направить уведомление автору
                End If
				sNotificationSubject=SIT_NotificationApproveRefused
			End If	
		Else 'СОГЛАСОВАНИЕ
		 	If Trim(Request("r")) <> "y" Then 'Согласовано
'Запрос №17 - СТС - start
				sNotificationBody = DOCS_Reconciled
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
                    sNotificationSubject = "Согласование (ДОГОВОР НЕ СДАН НА ХРАНЕНИЕ В БУХГАЛТЕРИЮ!)"
                End If
			   
			End If
				
'Запрос №17 - СТС - end
'Отправляем все по умолчанию + отдельно уведомление инициатору (в DBUpdateAfter.asp)
'20090622 - Заявка ТКП
'Для коммерческих предложений в DBUpdateAfter.asp автору и списку ознакомления (менеджеру)

                   ' if GetRootDepartment(Session("CurrentDepartmentDoc")) = SIT_RTI then
                   '   sNotificationList = sNotificationList + "G"
                   ' end if

			Else 'Отказано
'Запрос №17 - СТС - start
				sNotificationBody = DOCS_Refused
'Запрос №17 - СТС - end
				If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_ISHODYASCHIE)) > 0 Then
					sNotificationList="C" 'Направить уведомление создателю (инициатор)
'20090622 - Заявка ТКП
                ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_COM_OFFERS)) = 1 Then
                    sNotificationList="VC" 'Уведомление списку ознакомления(менеджер), автору
				Else
					sNotificationList="U" 'Направить уведомление автору (инициатор)
				End If
				sNotificationSubject=SIT_NotificationAgreeRefusedByUser+Session("UserID")
			End If
			
		End If
'Запрос №17 - СТС - start
		If GetRootDepartment(Session("CurrentDepartmentDoc")) = SIT_STS and (Trim(Request("app")) <> "y" or Trim(Request("r")) = "y") Then
			sNotificationBody = NotificationWithReason(sNotificationBody, Request("reason"))
		Else
			sNotificationBody = ""
		End If
'Запрос №17 - СТС - end

    Case UCase("/VisaCancel.asp")
        If Request("sop")="refuse" Then
            sNotificationList="TU" 'CU
            sNotificationSubject="Требует согласования – отменен отказ в согласовании"
            sNotificationBody="Требует согласования – отменен отказ в согласовании"
        End If

	Case UCase("/MakeCompleted.asp") ' make document status as completed
			If InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) > 0 or InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) > 0 Then
			    sNotificationList="CUR" 'Для заявок рассылка создателю, автору и исполнителю
			Else
				sNotificationList="CURVD" 'Рассылка инициатору, автору, исполнителю и по ознакомлению (соисполнителям)
			End If
	Case UCase("/MakeCanceled.asp") ' make document status as canceled
'Ph - 20090206 - start
'			sNotificationList="CURVD"
			If InStr(UCase(Session("CurrentClassDoc")), UCase(DOCS_Notices)) > 0 Then
			    sNotificationList="CURD" 'Для поручений рассылка создателю, автору, исполнителю и соисполнителям (D)
'20090622 - Заявка ТКП
            ElseIf InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_COM_OFFERS)) = 1 Then
                sNotificationList="VEG" 'Уведомление списку ознакомления(менеджер), всем согласующим
			Else
			    sNotificationList="CUR" 'Для поручений рассылка создателю, автору и исполнителю
			End If
'Ph - 20090206 - end
	Case UCase("/RequestCompleted.asp") ' запрос статуса исполнено
			sNotificationList="CURVL"
'20090622 - Заявка ТКП
	Case UCase("/UploadRetNew.asp") ' upload file into the document record
            If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_COM_OFFERS)) = 1 Then
              If Trim(Request("AgreeAgain")) = "ON" And Trim(Request("MainVersion")) = "ON" Then 'Agree process will be restarted
                sNotificationList="VEG" 'Уведомление списку ознакомления(менеджер), всем согласующим
                sNotificationSubject=DOCS_AGREERepeate
              End If
            End If
'			sNotificationList="CE"
'			If Trim(Request("AgreeAgain")) = "ON" And Trim(Request("MainVersion")) = "ON" Then 'Agree process will be restarted
'				If RUS()="RUS" Then
'					sNotificationSubject="Загружен файл основной версии документа. Согласование производится повторно"
'				Else
'					sNotificationSubject="Document main version file has been uploaded. Agree process restarted"
'				End If
'			Else
'				sNotificationSubject=DOCS_VersionFileUploaded
'			End If	
	'Case UCase("/MakeArchival.asp") ' make document status as archival or operative
	'	If Trim(Request("Archival")) = "y" Then 'make document status as archival 
	'		sNotificationList="C"
	'	Else 'make document status as operative
	'		sNotificationList="C"
	'	End If
	'Case UCase("/CreateComment.asp") ' create comment, create comment message
	'		sNotificationList="CR" 'document creator, responsible person
	'Case UCase("/MakeSigned.asp") ' document signed
	'		sNotificationList="CR" 'document creator, responsible person
	Case UCase("/MakeRegisteredRet.asp") ' document registered
			sNotificationList="CUV" 'Уведомление создателю, автору и всем получателям ViewList
'			sNotificationList="CRL" 'document creator, responsible person, controlling person
			'vnik_change_req
			sNotificationList=sNotificationList + "S"   'добавил уведомление для регистратора
			'vnik_change_req
			sNotificationSubject=DOCS_Registered + Request("DocID")
			'SAY 2008-12-03
			If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_ISHODYASCHIE)) > 0 Then
'				sNotificationList="CUVD" 'Направить уведомление создателю (инициатор)
'Ph - 20081209 - Уведомление адресатам с другой темой в DBUpdateAfter
				sNotificationList="CU" 'Направить уведомление создателю и автору
			End If
            
            'vnik_purchase_order
            If InStr(UCase(Request("ClassDoc")),UCase(SIT_PROTOCOLS_CPC)) > 0 Then
                sNotificationList="V"       
            End If
            'vnik_purchase_order


'	Case UCase("/ModifyPaymentStatus.asp") ' modify document payment status
'    		If Request("status") = "NotPaid" Then
'				sNotificationList="-"
'    		ElseIf Request("status") = "ToBePaid" Then
'				sNotificationList="-"
'    		ElseIf Request("status") = "SentToBePaid" Then
'				sNotificationList="-"
'    		ElseIf Request("status") = "ToPay" Then
'				sNotificationList="CR"
'    		ElseIf Request("status") = "ToPayCancel" Then
'				sNotificationList="CR"
'    		ElseIf Request("status") = "Paid" Then
'				sNotificationList="-"
'    		End If
'ph - 20090717 - start
	Case UCase("/MakeCanceledCancel.asp") 'Отмена статуса Отменено
          If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_COM_OFFERS)) > 0 Then
            sNotificationList = "-" 'Стандартную рассылку отключаем
          End If
'ph - 20090717 - end
	Case UCase("/ChangeDoc.asp") ' change document record
			If UCase(Request("UpdateDoc")) = "YES" Then
'ph - 20090717 - start
              If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_COM_OFFERS)) > 0 Then
                If UCase(Request("ClearStatusReconciled")) = "ON" Then
                  sNotificationList = "V" 'рассылка менеджеру
                End If
              End If
'ph - 20090717 - end
			  If InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) > 0 or InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) > 0 Then
			    If not IsReconciliationComplete(S_ListToReconcile, S_ListReconciled) Then
				  If Request("NoResetStatuses")="ON" Then 'Только если статусы сохранены, иначе идет стандартная рассылка и получается 2 письма
'ph - 20100726 - start
                  If Request("IsActive2") <> VAR_InActiveTask Then
'ph - 20100726 - end

			        sNotificationList = "E" 'рассылка еще не согласовавшим документ
		            sNotificationSubject = SIT_DocChangedReconcilationRequired
'ph - 20090825 - start
                    'Запоминаем сообщение, чтобы восстановить его в DBUpdateAfter и очищаем, чтобы оно не попадало в письмо
                    SIT_SessionMessageSaveBeforeSendNotitfication = Session("Message")
                    Session("Message") = ""
                    AddLogD SIT_SessionMessageSaveBeforeSendNotitfication
'ph - 20090825 - end
'ph - 20100726 - start
                  End If
'ph - 20100726 - end

			      End If
			    End If
			  End If
'				'If Session("UserID")<>GetLogin(Request("NameCreation")) Then
'				'	sNotificationList="C"
'				'	sNotificationSubject="Document "+Request("DocID")+" is updated by "+Session("UserID")+" user"
'				'Else
'				'	sNotificationList="-"
'				'	sNotificationSubject=""
'				'End If
'				If Request("NoResetStatuses")<>"ON" Or Request("ClearStatusReconciled")="ON" Then 'Reconciliation statuses cleared
'					sNotificationList="C"
'					'sNotificationSubject="Document "+Request("DocID")+" - Reconciliation statuses cleared"
'					'sNotificationSubject="Документ изменен "+Request("DocID")+" - Согласование надо повторить"
'				End If
			End If
'	Case UCase("/ReconciliationSuspend.asp")  'Приостановить или возобновить согласование
'   	 		If Request("stop") = "y" Then 'Приостановить согласование
'        		sNotificationList="U" 'Уведомление автору
'        		sNotificationSubject="Согласование приостановлено"
'    		Else  'Возобновить согласование
'        		sNotificationList="E" 'Уведомление текущим согласующим
'        		sNotificationSubject="Согласование возобновлено"
'    		End If
'20090622 - Заявка ТКП
	Case UCase("/ReconciliationCancelRet.asp")  'Отменить согласование некоторых пользователей
        If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_COM_OFFERS)) = 1 Then
            sNotificationList="VEG" 'Уведомление списку ознакомления(менеджер), всем согласующим
        End If
	Case UCase("/ReconciliationForce.asp")  'Возобновить согласование несмотря на наличие отказавшего
        If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_COM_OFFERS)) = 1 Then
            sNotificationList="VEG" 'Уведомление списку ознакомления(менеджер), всем согласующим
        End If
End Select
'sNotificationList="-" ' switch off e-mail notifications

'If UCase(Request.ServerVariables("URL"))=UCase("/MakeActive.asp") Then
'If InStr(UCase(Request.ServerVariables("URL")),UCase("/MakeActive.asp"))>0 Then
'	If IsHelpDeskSAP() Then
'		'sResult=SetDocFieldIfNotSet(Request("DocID"), "Rank", "Обычный")
'		'sResult=SetDocFieldIfNotSet(Request("DocID"), "DateCompletion", Date+7)
'		'sResult=SetDocFieldIfNotSet(Request("DocID"), "NameResponsible", """Администратор"" <Admin>")
'	End If
'End If

'If IsHelpDeskDoc() Then
'AddLogD "IsHelpDeskDoc 7"
'	If UCase(Request("UpdateDoc"))="YES" Then
'AddLogD "IsHelpDeskDoc 8"
'		sGetNewUserIDsInList=oPayDox.GetNewUserIDsInList(Request("DocCorrespondentOld"), S_Correspondent_Set, False)
'		If sGetNewUserIDsInList<>"" Or (Request("create")="y" And S_Correspondent_Set<>"") Then
'AddLogD "IsHelpDeskDoc 9"
'			sNotificationList=sNotificationList+"D"
'			sNotificationSubject="Передано в обработку"
'		End If
'	End If
'End If

If InStr(UCase(Request.ServerVariables("URL")),UCase("/CreateComment.asp"))>0 and Request("resol")="y" Then
  sNotificationList="CAG" 
  sNotificationSubject="Поступила резолюция на документ"
End If

'C - document creator
'U - document author
'R - responsible person
'L - controlling person
'A - approving person
'V - viewers list
'D - correspondents, notification list
'E - reconciliation (agree) list - not agreed yet
'G - reconciliation (agree) list - agreed
'S - registrars list


AddLogD "sNotificationList:"+sNotificationList
AddLogD "sNotificationSubject:"+sNotificationSubject
'vnik_send_notification
Function IsSendFilesUser(ByRef dsUser, ByRef dsDoc)
    IsSendFilesUser = False 'по умолчанию файлы не аттачим
        AddLogD "IsSendFilesUser.URL:" & Request.ServerVariables("URL")
         AddLogD "IsSendFilesUser.URL:" & dsDoc("ListToReconcile")
          AddLogD "IsSendFilesUser.URL:" & dsDoc("ListReconciled") '+ " <" & dsUser("UserID") & ">;"
          AddLogD "IsSendFilesUser.URL:" & IsReconciliationCompleteWithOptions(dsDoc("ListToReconcile"), dsDoc("ListReconciled"))
            AddLogD "IsSendFilesUser.URL:" & Trim(Request("r"))
    If (UCase(Request.ServerVariables("URL")) = UCase("/MakeActive.asp")) OR (UCase(Request.ServerVariables("URL")) = UCase("/Visa.asp") and IsReconciliationCompleteWithOptions(dsDoc("ListToReconcile"), dsDoc("ListReconciled")) and Trim(Request("r")) <> "y") OR (UCase(Request.ServerVariables("URL")) = UCase("/MakeRegisteredRet.asp")) OR (UCase(Request.ServerVariables("URL")) = UCase("/AddListCorrespondentRet.asp") and IsRegistered(dsDoc("LocationPath"))) Then 'АКТИВАЦИЯ'УТВЕРЖДЕНИЕ'РЕГИСТРАЦИЯ
        AddLogD "IsSendFilesUser.DocDepartment:" & dsDoc("Department")
        AddLogD "IsSendFilesUser.LocationPath:" & dsDoc("LocationPath")
        AddLogD "IsSendFilesUser.IsRegistered:" & IsRegistered(dsDoc("LocationPath"))
        If InStr(dsDoc("Department"), SIT_SITRONICS) = 1 Then 'условие по подразделению в документе
            AddLogD "IsSendFilesUser.SIT_AttachFilesToNotofication:" & SIT_AttachFilesToNotofication
            If InStr(UCase(ReplaceRoleFromDir(SIT_AttachFilesToNotofication,SIT_SITRONICS)),UCase(dsUser("UserID"))) > 0 Then 'условие по пользователю
                sUser = "<" & dsUser("UserID") & ">"
                AddLogD "IsSendFilesUser.sUser:" & sUser
                AddLogD "IsSendFilesUser.ClassDoc:" & dsDoc("ClassDoc")
                AddLogD "IsSendFilesUser.NameResponsible:" & dsDoc("NameResponsible")
                AddLogD "IsSendFilesUser.ListToView:" & dsDoc("ListToView")
                'условия по категориям документа и вхождению пользователя в ту или иную группу
                If InStr(dsDoc("ClassDoc"), SIT_PAYMENT_ORDER) > 0 or InStr(dsDoc("Author"), sUser)Then
                    IsSendFilesUser = False
                Else
                    IsSendFilesUser = True 'если все условия выполнены, то цепляем файлы (только основные версии!!!)
                End If
                'If InStr(dsDoc("ClassDoc"), SIT_ZADACHI) > 0 and InStr(dsDoc("NameResponsible"), sUser) > 0 or InStr(dsDoc("ClassDoc"), SIT_VHODYASCHIE) > 0 and InStr(dsDoc("ListToView"), sUser) > 0 Then           
                'End If
            End If
        End If
    End If
    AddLogD "IsSendFilesUser:" & CStr(IsSendFilesUser)
End Function
'vnik_send_notification

%>
