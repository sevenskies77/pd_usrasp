<%
'Make some right pane buttons available
'
'bStatusPaymentPaid - make button "Status «Paid»" available
'If Session("UserID")="Cashier" Then 'Provide your actual user login instead of "Cashier"
'	bUserStatusPaymentPaid=True
'End If
'
'bUserNOTApprove - make button "Approve" unavailable
'If S_ClassDoc="Payments" Then
'	If Session("UserID")="BigBoss" Then 'Provide your actual user login instead of "BigBoss"
'		VAR_ButtonsNotToShow="ClickApprove" 'Use "To pay" button instead of "Approve" for payment documents
'	End If
'End If
'
'bUserNOTChange - make button "Modify" unavailable
'If Session("UserID")="RestrictedUser" Then 'Provide your actual user login instead of "RestrictedUser"
'	bUserNOTChange=True
'End If

'Make "Register" button available only for approved contracts
'If CurrentClassDoc="Contracts" Then
'	If ShowStatusDevelopment(S_StatusDevelopment)<>DOCS_Approved Then 
'		VAR_ButtonsNotToShow="ClickMakeRegistered, ClickUpdateRegLog"
'	End If
'End If

'VAR_ButtonsNotToShow="ClickVisaDelegate"
If IsHelpDeskDoc() Then
	'ClickRequestCompleted, ClickRefuseCompletion
	VAR_ButtonsToShow	=But_Statuses+", "+But_Admin1+", "+But_Filing+", "+But_Linked+", "+But_Modification+", "+But_Aproval+", "+But_MSWord+", "+But_Completion+", ClickApprove, ClickUserEval, ClickCreateDoc, ClickDistribute, ClickRedistribute, ClickSetResponsible, ClickReSetResponsible, ClickMakeResponsible, ClickResolution, ClickMakeCompleted, ClickMakeCompletedNot, ClickMakeCanceled, ClickSendNotification, ClickUploadFileNew, ClickAuditingDoc, ClickRestoreDoc, ClickMakeArchival, ClickMakeOperative, ClickChangeDoc, ClickModifyPercent, ClickCreateComment, ClickCreateCommentHelpDesk, ClickCreateCommentHistory, ClickCreateNotice, ClickDeleteDoc, ClickDeleteMessage, ClickMakeActive, ClickMakePublic, ClickMakeCanceled, ClickMakeArchival,  ClickMakeCompleted, ClickMakeOperative, ClickMSOffice, ClickMSOfficeStandard"
	'VAR_ButtonsToShow	=But_Reconciliation+", "+VAR_ButtonsToShow
	If Not IsHelpDeskAdminOrConsultant() Then
		VAR_ButtonsNotToShow	="ClickCreateCommentHistory, ClickCreateCommentHelpDesk, ClickCreateCommentHistory"
		If Trim(MyCStr(S_NameResponsible))<>"" Then
			VAR_ButtonsNotToShow	=VAR_ButtonsNotToShow+", ClickChangeDoc"
		End If
	End If
End If

'ph - 20080824
'Прячем кнопку "Приостановить" от всех кроме инициатора и Админов
If not IsAdmin() Then
  If (InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_ISHODYASCHIE)) > 0 and InStr(S_NameCreation, "<"+Session("UserID")+">") > 0) or _
    ((InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_SLUZH_ZAPISKA)) > 0 or _
    InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_ZADACHI)) > 0 or _
    InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI)) > 0) and _
    InStr(S_Author, "<"+Session("UserID")+">") > 0) _
  Then
    ShowDoc_ShowButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickReconciliationSuspend"
  Else
    ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickReconciliationSuspend"
  End If

  'SAY - 2008-08-25
  'ShowDoc_ShowButton VAR_ButtonsToShow, VAR_ButtonsNotToShow,"ClickRequestCompleted,ClickMakeCompleted"
  If InStr(S_NameControl, "<"+Session("UserID")+">") > 0 Then
    VAR_ButtonsNotToShow = VAR_ButtonsNotToShow+",ClickRequestCompleted"
  End If

  If S_StatusCompletion="1" Then
    VAR_ButtonsNotToShow = VAR_ButtonsNotToShow+",ClickModifyDateCompletion,ClickModifyNameControl"
  End If
  
  'ph - 20081118 - Прячем раздел Администрирование
  ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, But_Admin1
End If

'ph - 20080921 - Убрать кнопку комментарий в Поручениях
If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_ZADACHI)) > 0 Then
  ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickCreateComment"
'Ph - 20080922 - Убрать кнопку Ознакомлен у Соисполнителей
  If InStr(S_ListToView, "<"+Session("UserID")+">") > 0 Then
    ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickCreateViewed"
  End If

'показать кнопку Ход исполнения поручений РТИ для Боева'
  If InStr(Session("UserID"),"boev_oaorti") > 0 and InStr(Session("Department"), UCase(SIT_RTI)) >0 and InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_ZADACHI)) > 0 Then
    ShowDoc_ShowButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickCreateCommentHistory"
  End If

'Ph - 20080922 - Показать кнопку Ход исполнения у Ответственного SAY 2008-11-19 И у соисполнителей
  If InStr(S_NameResponsible, "<"+Session("UserID")+">") > 0 or InStr(S_Correspondent, "<"+Session("UserID")+">") > 0 or InStr(S_Author, "<"+Session("UserID")+">") > 0 or InStr(S_NameControl, "<"+Session("UserID")+">") > 0 Then
    ShowDoc_ShowButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickCreateCommentHistory"
  End If

  
'AM - 08122008 - Покажем кнопку "Загрузить файл" для списка соисполнителей
  If InStr(S_Correspondent, "<"+Session("UserID")+">") Then
    VAR_CanCreateMainVersionFiles=True
  End If
End If

If not IsAdmin() Then
  ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickUpdateRegLog"
End If

'rmanyushin 03.09.2009 9:10 Start "Показать кнопку "Ссылка" в карточке документа только для пользователей СТС"
If InStr(Session("Department"), UCase(SIT_STS)) =1 Then
	ShowDoc_ShowButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickCreateCommentLink"
End If
'rmanyushin 03.09.2009 9:10 End

'rmanyushin 51555, 56781 13.10.2009 Start
'Скрыть кнопоки для привелигированных пользователей СТС - Аудитор СТС и Контролер СТС
If isPrivilegedUserSTS() Then
    VAR_ButtonsNotToShow = VAR_ButtonsNotToShow+",ShowListComments,ShowListNotices,ShowListSpecItems,ShowListDependantDoc,ShowListFollowingDoc,ClickDeleteDocFromFolder,ClickCopyDocToFolder"
	VAR_ButtonsNotToShow = VAR_ButtonsNotToShow+",ClickMakeRegistered,ClickUpdateRegLog,ClickMakeRegisteredNot,ClickHome,ClickListDoc,ClickShowBPs,ClickCreateViewed,ClickMSOfficeStandard"
	VAR_ButtonsNotToShow = VAR_ButtonsNotToShow+",ClickMSWordPrinting,ClickMSExcelStandard,ClickMSOfficeReconciliationList,ClickMSOfficeChangesList,ClickMSOfficeViewedList,ClickEMailingTest"
	VAR_ButtonsNotToShow = VAR_ButtonsNotToShow+",ClickEMailingCheck,ClickEMailing,ClickCreateCommentReviewRed,ClickChangeDocRed,ClickVisa,ClickVisaAdd,ClickVisaDelegate,ClickRefuse"
	VAR_ButtonsNotToShow = VAR_ButtonsNotToShow+",ClickToReview,ClickToModify,ClickVisaNot,ClickRefuseNot,ClickRefuseCancel,ClickVisaCancel,ClickReconciliationCancel,ClickReconciliationSuspend"
	VAR_ButtonsNotToShow = VAR_ButtonsNotToShow+",ClickReconciliationRelease,ClickReconciliationAgain,ClickReconciliationForce,ClickReconciliationComplete,ClickStatusPaymentNotPaid" 
	VAR_ButtonsNotToShow = VAR_ButtonsNotToShow+",ClickStatusPaymentPaidPart,ClickStatusPaymentToBePaid,ClickStatusPaymentToBePaidPart,ClickStatusPaymentSentToBePaid"
	VAR_ButtonsNotToShow = VAR_ButtonsNotToShow+",ClickStatusPaymentSentToBePaidUnavailable,ClickStatusPaymentToBePaidUnavailable,ClickSetDeputy,ClickDeleteMessage"
	VAR_ButtonsNotToShow = VAR_ButtonsNotToShow+",ClickStatusPaymentToPay,ClickStatusPaymentToPayPart,ClickStatusPaymentToPayPartPart,ClickStatusPaymentToPayPartRest,ClickStatusPaymentToPayCancel"
	VAR_ButtonsNotToShow = VAR_ButtonsNotToShow+",ClickStatusPaymentToPayCancelPart,ClickStatusPaymentPaid,ClickStatusPaymentPaidPartIncoming,ClickTransactionTransfer,ClickTransactionWithdrawal,ClickTransactionDeposit"
	VAR_ButtonsNotToShow = VAR_ButtonsNotToShow+",ClickTransactionTransferUnavailable,ClickTransactionWithdrawalUnavailable,ClickTransactionDepositUnavailable,ClickApprove,ClickApproveNot,ClickRefuseApp,ClickRefuseAppCancel"
	VAR_ButtonsNotToShow = VAR_ButtonsNotToShow+",ClickCreateCommentResolution,ClickCreateCommentHistory,ClickMakeSigned,ClickMakeSignedCancel,ClickDistribute,ClickRedistribute,ClickSetResponsible,ClickReSetResponsible"
	VAR_ButtonsNotToShow = VAR_ButtonsNotToShow+",ClickMakeResponsible,ClickResolution,ClickMakeCompleted,ClickRefuseCompletion,ClickMakeCompletedNot,ClickRequestCompleted,ClickMakeCanceled"
	VAR_ButtonsNotToShow = VAR_ButtonsNotToShow+",ClickMakeCanceledCancel,ClickMakeActive,ClickMakeActiveNot,ClickUserEval,ClickMakeInactive,ClickMakeInactiveNot,Click1C,Click1CNew"
	VAR_ButtonsNotToShow = VAR_ButtonsNotToShow+",ClickModifyDateCompletion,ClickMakePublic,ClickCreateDoc,ClickCreateDocTemplate,ClickCreateDocUnavailable,ClickCopyDoc,ClickPasteDoc,ClickReCalcSpec,ClickCopySpec" 
	VAR_ButtonsNotToShow = VAR_ButtonsNotToShow+",ClickPasteSpec,ClickShift,ClickAutofill,ClickPasteComponentsSpec,ClickCreateConnected,ClickCreateConnectedCopy,ClickCreateDisconnected"
	VAR_ButtonsNotToShow = VAR_ButtonsNotToShow+",ClickChangeDoc,ClickDeleteDoc,ClickModifyPercent,ClickModifyNameCreation,ClickModifyNameControl,ClickCreateComment,ClickCreateCommentHelpDesk,ClickCreateMessage,ClickESign"
	VAR_ButtonsNotToShow = VAR_ButtonsNotToShow+",ClickCreateCommentHelpDesk,ClickCreateMessage,ClickCreateCommentHistory,ClickCreateCommentReview,ClickCreateCommentPARTNER,ClickCreateCommentResource,ClickCreateCommentLink,ClickCreateCommentBPStep,ClickCreateBPInstance"
	VAR_ButtonsNotToShow = VAR_ButtonsNotToShow+",BPClickConnectToBP,BPClickDisconnectFromBP,ClickCreateDocDependant,ClickCreateDocFollowing,ClickCreateContact,ClickCreateEvent"
	VAR_ButtonsNotToShow = VAR_ButtonsNotToShow+",ClickCreateNotice,ClickShowReports,ClickCheckUsers,ClickSendNotification,ClickSendAll,ClickMSOffice,ClickGetBarCode,ClickUploadFileNew,ClickeBay"
	VAR_ButtonsNotToShow = VAR_ButtonsNotToShow+",ClickChangePaperFileName,ClickGoToPaperFile,ClickMakeWaiting,ClickMakeSent,ClickMakeDelivered,ClickMakeReturnedToFile,ClickMakeFiled,ClickMakeReturned,ClickAuditingDoc"
	VAR_ButtonsNotToShow = VAR_ButtonsNotToShow+",ClickRestoreDoc,ClickMakeArchival,ClickMakeOperative,ClickDownloadXML,ClickCreateShortcut"
	
	'Скрываем разделители
	ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, But_MSWord
	ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, But_Reconciliation
	ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, But_Aproval
	ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, But_Completion
	ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, But_Modification
	ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, But_Linked
	ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, But_Admin1
	'ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, But_Statuses
	
	If UCase(Session("UserID")) = UCase(STS_HeadOf789) Then
	    ShowDoc_ShowButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, But_Linked
	    ShowDoc_ShowButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickSendNotification"
	End IF
End If
'rmanyushin 51555, 56781 13.10.2009 End

'vnik_archive
If Session("Archive")="YES" Then
    VAR_ButtonsNotToShow = VAR_ButtonsNotToShow+",ClickMakeRegistered,ClickCreateViewed"        
End If
'vnik_archive

'ph - 20091224 - start
'Скрыть кнопки по БП
ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickCreateBPInstance"
ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "BPClickConnectToBP"
ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "BPClickDisconnectFromBP"
'ph - 20091224 - end

'Запрос №11 - СТС - start
'Запретить создание документов в старой категории договоров
If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_DOGOVORI_OLD)) = 1 Then
	ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickCreateDoc"
	ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickCreateDocTemplate"
End If

'Запрос №11 - СТС - end

'rmanyushin 110094 14.07.2010 start
'Скрыть кнопку "Requested completed" у всех, кроме Администраторов
 If not IsAdmin() AND (InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) > 0 OR InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) > 0) Then
    ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickRequestCompleted"
 End If
'rmanyushin 110094 14.07.2010 end

'Запрос №30 - СТС - start
'Показать в заявках на оплату кнопку Переназначить пользователям из ролей STS_Purchase_Logistics_Department справочника RolesForOrders_STS
'Запрос №36 - СТС - с роли STS_Purchase_Logistics_Department переключен показ на роль STS_UsersToShowResetResponsibleButton из RolesForOrders_STS
If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) > 0 and InStr(UCase(GetUsersToShowClickReSetResponsible), "<" & UCase(Session("UserID")) & ">") Then
  'Устанавливаем флаг необходимости показа кнопки, сам показ в UserClick в обход системных правил
  'Ставить условие с обращением к БД в UserClick нельзя, он выполняется многократно
  bSTS_ShowClickReSetResponsible = "Y"
  'Обязательно показываем раздел Исполнение
  ShowDoc_ShowButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, But_Completion
  'Отключаем показ кнопки Переназначить, чтобы она не была показана системно второй раз
  ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickReSetResponsible"
  '...и Назначить заодно
  ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickSetResponsible"
End If
'Запрос №30 - СТС - end

'Прячем кнопку "На правку" в Заявках на закупку и оплату
If (InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) > 0 or InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) > 0) and not IsAdmin() Then
  ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickRefuseCompletion"
End If

'Запрос №34 - СТС - start
'Заплатка!!! Прячем кнопку редактировать для Служебок на переработки созданных до определенной даты
If UCase(Request.ServerVariables("URL")) = UCase("/ShowDoc.asp") Then
  If InStr(UCase(Session("CurrentClassDoc")), UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) = 1 and dsDoc("DateCreation") < DateSerial(2010,12,20) Then
    ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickChangeDoc"
    ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickChangeDocRed"
  End If
End If
'Запрос №34 - СТС - end

'Запрос №33 - СТС - start
'Автоматически отмененный документ (определяется в UserShowListComments.asp), прячем кнопки позволяющие возобновить
If bHideClickMakeCanceledCancel Then
  ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickMakeCanceledCancel" 'Отмена отмены
  ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickChangeDoc" 'Редактирование
  ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickChangeDocRed" 'Перестраховка, у отмененного документа не должно быть этой кнопки
  ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickMakeInactive" 'Деактивация (после деактивации и активации документ опять будет возобновлен)
End If
'Запрос №33 - СТС - end

'ph - 20110128 - start
If not IsAdmin() Then
  If InStr(S_ListReconciled, "<" & Session("UserID") & ">") > 0 Then
    ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickVisaAdd" 'Запрет добавления согласующих после сделанного согласования
  End If
End If
'ph - 20110128 - end

'Запрос №38 - СТС - start
If UCase(Request.ServerVariables("URL")) = UCase("/ShowDoc.asp") Then
  If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_DOGOVORI_NEW)) = 1 and InStr(dsDoc("Department"), SIT_STS_ROOT_DEPARTMENT) = 1 Then
    If IsNumeric(SIT_MaxUsersInListToReconcile) Then
      If not CheckUsersInListToReconcile(dsDoc("ListToReconcile"), CInt(SIT_MaxUsersInListToReconcile)-1) Then
	    'Прячем настоящую кнопку Добавить
        ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickVisaAdd"
	    'Ставим флаг показа подмененной кнопки Добавить
		bSTS_ShowClickVisaAddFake = "Y"
      End If
	End If
  End If
End If
'Запрос №38 - СТС - end

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

If InStr(UCase(Session("CurrentClassDoc")), UCase(RTI_PURCHASE_ORDER)) = 1 Then
  ShowDoc_ShowButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, But_MSOfficeStandard
End If
If InStr(UCase(Session("CurrentClassDoc")), UCase(RTI_CONTRACT)) = 1 Then
  ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, But_MSOfficeStandard
End If

'MIKRON start
'В Заявке на закупку, Опросном листе, БСАП, Протоколе ЗК прячем кнопку "Подчиненный"
'Вместо этого вставляем кнопки БСАП и Опросный лист и тд. см. файл UserClick.asp
'If InStr( UCase(Session("CurrentClassDoc")),UCase(MIKRON_PURCHASE_ORDER) ) > 0 or _
'   InStr( UCase(Session("CurrentClassDoc")),UCase(MIKRON_PROTOCOL) ) > 0 or _
'   InStr( UCase(Session("CurrentClassDoc")),UCase(MIKRON_RL_PROTOCOL) ) > 0 or _
'   InStr( UCase(Session("CurrentClassDoc")),UCase(MIKRON_RL_MEMO) ) > 0 or _
'   InStr( UCase(Session("CurrentClassDoc")),UCase(MIKRON_BSAP) ) > 0 Then
If InStr( Session("Department"),SIT_MIKRON ) > 0 Then
   ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickCreateDocDependant"
End If
'В Договоре прячем кнопку Согласование и добавить согласующего, т.к они добавляют только в конец списка
'чтобы изменить список согласующих, надо войти в редактирование документа
If UCase(Request.ServerVariables("URL")) = UCase("/ShowDoc.asp") Then
   If InStr( UCase(Session("CurrentClassDoc")),UCase(MIKRON_CONTRACT) ) > 0 or _
      InStr( UCase(Session("CurrentClassDoc")),UCase(MIKRON_S_CONTRACT) ) > 0 or _
      InStr( UCase(Session("CurrentClassDoc")),UCase(MIKRON_ADD_CONTRACT) ) > 0 Then
      ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickVisaAdd"
   End If
End If
'MIKRON end

'Всегда и от всех прячем кнопку отмены подписания, она не нужна
ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickApproveCancel"

If UCase(Request.ServerVariables("URL")) = UCase("/ShowDoc.asp") Then
  If InStr(S_Department, SIT_SIB_ROOT_DEPARTMENT) = 1 Then
    If InStr(S_NameCreation, "<" & Session("UserID") & ">") > 0 and InStr(S_NameApproved, "-<") > 0 Then
      'Показываем кнопку Редактирования
      VAR_ButtonsMustShow = VAR_ButtonsMustShow & ",ClickChangeDoc"
    End If
  End If
End If

'прячем кнопку отмены согласования, если нет отказа в согласовании
If UCase(Request.ServerVariables("URL")) = UCase("/ShowDoc.asp") Then
  If InStr(S_Department, SIT_RTI) = 1 Then
    If InStr(Trim(S_ListReconciled), "-<") = 0 Then
      'cскрываем кнопку отмены отказа
      ShowDoc_HideButton VAR_ButtonsToShow, VAR_ButtonsNotToShow, "ClickReconciliationCancel"
    End If
  End If
End If

'запрещаем создавать документы категории Договоры МИКРОН/Договор до 01.10.2013
If InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_OLD_CONTRACT)) = 1 Then
  VAR_ButtonsNotToShow = VAR_ButtonsNotToShow & ",ClickCreateDoc,ClickCreateDocTemplate" 'Прячем кнопку создать и создать по шаблону
End If
'запрещаем создавать документы категории Договоры МИКРОН/Договор до 01.10.2013

'out "bNonUpdatable = " & CStr(bNonUpdatable)
%>