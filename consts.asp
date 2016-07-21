<%
' для русского языка меняем название кнопки
If RUS()="RUS" Then 
  BUT_VIEWEDSTATUSDOCS = "ОЗНАКОМЛЕНИЕ"
End If

'Запрос №33 - СТС - start
'НЕ МЕНЯТЬ - текст комментария при автоматической отмене документа, он же используется для проверки, что документ отменен автоматически, чтобы не дать перезапустить
STS_AutoCancelledComment = "Automatic cancellation after long period of users inactivity"

STS_AutoCancellationNotificationSubj = "Your document #DOCID# cancelled"
STS_AutoCancellationNotificationBody = "The document is rejected because of no agreement actions were performed within 15 working days. Cannot resume the agreement." '"Документ отменен, по причине отсутствия действий по согласованию документа в течении 15-ти рабочих дней. Возобновление согласования невозможно"
STS_AutoCancellationWarningSubj = "Your document will be cancelled after 3 days"
STS_AutoCancellationWarningBody = "No agreement actions for the document. Please perform actions to continue the agreement. If the agreement is not continued, the document will be automatically rejected within 3 working days without the possibility to resume the agreement process." '"Отсутствуют действия по согласованию документа, пожалуйста, предпримите меры для продолжения согласования. Если согласование не продолжится, документ будет отменен через 3 рабочих дня, без возможности возобновления процесса согласования"
STS_ExpirationWarningSubj = "Your document will be cancelled after 12 days"
STS_ExpirationWarningBody = "You violate the term of the document agreement specified in the normative documents of the Company. Please perform actions to continue the agreement. If the agreement is not continued, the document will be automatically rejected within 12 working days without the possibility to resume the agreement process." '"Вы нарушаете срок согласования документа, установленный в нормативных документах компании. Пожалуйста, предпримите меры для продолжения согласования. Если согласование не продолжится, документ будет отменен через 12 рабочих дней, без возможности возобновления процесса согласования"
'Запрос №33 - СТС - end

'Поручения АФК Система для запроса по кнопке ПОРУЧЕНИЯ АФК (у Ушаковой)
SIT_AFK_Tasks = sUnicodeSymbol+"'Поручения АФК',"+sUnicodeSymbol+"'SISTEMA tasks',"+sUnicodeSymbol+"'Úkoly SISTEMA'"
SIT_AdditionalAgreesDelimeter = "##;"

'Список типов поручений по которым идет специальная рассылка о просрочках
SIT_TasksTypesForSpecialNotification = "Поручения АФК, Поручения Правления, Поручения СД, Поручения Президента"
'Разделитель в этом списке
SIT_TasksTypesDelimeter = ","

'Перевод не нужен, используется только русский
SIT_CorporateControlChief = """#Начальник отдела корпоративного контроля"";"

'Роли для заявок. Только на английском
STS_Orders_Accounting = """#Accounting department"";"
STS_Orders_Treasury = """#Treasury"";"
STS_Orders_FinDirector = """#Financial director"";"
STS_Orders_FinancialControl = """#Financial controller"";"
STS_Orders_HeadOfDepartment = """#Department director"";"
STS_Orders_HeadOfSector = """#Sector manager"";"
STS_Orders_HeadOfDivision = """#Division director"";"
STS_Orders_GenDirector = """#CEO"";"
STS_Orders_ProjectManager = """#Owner of the project budget"";"

'ph - 20101225 - end
STS_HRAdministrationManager = """#HR Administration Manager"";"
'ph - 20101225 - start

'rmanyushin 96545 06.05.2010 Start
STS_Orders_DptOSSBSS = """#Deputy CEO for OSS/BSS"";"
'rmanyushin 96545 06.05.2010 End

'rmanyushin 136151 14.10.2010 Start
STS_Orders_DptEMS = """#Operational Deputy Division Director EMS"";"
'rmanyushin 136151 14.10.2010 End

'rmanyushin 158840 18.12.2010 Start
STS_Orders_DirNS = """#Division Director NS"";"
'rmanyushin 158840 18.12.2010 End

'rmanyushin 105583 21.06.2010 Start
STS_Orders_ResponsiblePayO = """#Responsible person PayO"";"
'rmanyushin 105583 21.06.2010 End

'rmanyushin 51555 17.09.2009 Start
STS_Overseer = "OverseerSTS" 'Account пользовательской роли - "Контролер СТС" в справочнике пользователей
STS_Auditor = "AuditorSTS" 'Account пользовательской роли - "Аудитор СТС" в справочнике пользователей
'rmanyushin 51555 17.09.2009 End

'rmanyushin 133266 05.10.2010 Start
STS_LegalSTS = "LegalSTS" 'Account пользовательской роли - "Юрист СТС" в справочнике пользователей
'rmanyushin 133266 05.10.2010 End

'rmanyushin 56781 13.10.2009 Start
STS_HeadOf789 = "HeadOf789" 'Account пользовательской роли - "Дивизионы 7-8-9" в справочнике пользователей
'rmanyushin 56781 13.09.2009 End

'rmanyushin 79501 24.02.2010 Start
STS_POViewer = "POViewer" 'Role for Mr. Cupic, director of Purchasing and Logistics department to see all PO´s.
'rmanyushin 79501 24.02.2010 End

'Запрос №30 - Start
STS_Purchase_Logistics_Department = """#Purchase & Logistics Department"";"
'Запрос №30 - End

'Запрос №31 - СТС - start
STS_Orders_CEO_SC = """#CEO SC"";"
'Запрос №31 - СТС - end
'Запрос №43 - СТС - start
STS_Assistant_PO_50000 = """#Assistant PO 50000"";"
STS_Assistant_PO_more_than_50000 = """#Assistant PO more than 50000"";"
'Запрос №43 - СТС - end

'Запрос №36 - СТС - start
STS_UsersToShowResetResponsibleButton = """#Users to show ResetResponsible button"";"
'Запрос №36 - СТС - end

'rmanyushin 119579 19.08.2010 Start
STS_HolidayRequest_Refused = "The holiday request has not been approved"
'rmanyushin 119579 19.08.2010 End

'Приписки об автосогласовании. По умолчанию на английском, нужный язык ставится при вызове SetLangConstsForAutoReconcilation
SIT_Agreed = "Agreed"
SIT_AutoAgreed = " / Agreed by default"

'Тексты почтовых уведомлений. По умолчанию на английском, нужный язык ставится при вызове SetLangConstsForEmail
SIT_AgreeTimeExceeded = "You have not met the document agreement deadline. Your superior will be informed."
'20090622 - Заявка ТКП
SIT_AgreeTimeExceededComOffer = "Time for commercial offer endorsement is up"
SIT_AgreeDelaying = "You are delaying the document agreement process"
SIT_UserDelayedAgree1 = "User "
SIT_UserDelayedAgree2 = " has not met the document agreement deadline."
SIT_OneDayForTaskCompletionSubj = "The task fulfilment deadline will expire in 1 day."
SIT_2DaysForTaskCompletionSubj = "The task fulfilment deadline will expire in 2 days."
SIT_3DaysForTaskCompletionSubj = "The task fulfilment deadline will expire in 3 days."
SIT_OneDayForTaskCompletionBody1 = "<BR><B>Dear colleague, we hereby inform you that ("
SIT_OneDayForTaskCompletionBody2 = ") the task fulfilment time limit expires tomorrow (the task card and the link to it in the document management system are below). The task is now being checked by "
SIT_OneDayForTaskCompletionBody3 = ". Please, complete the task fulfilment report.</B><BR>"

SIT_2DaysForTaskCompletionBody1 = "<BR><B>Dear colleague, we hereby inform you that ("
SIT_2DaysForTaskCompletionBody2 = ") the task fulfilment time limit expires in 2 days (the task card and the link to it in the document management system are below). The task is now being checked by "
SIT_2DaysForTaskCompletionBody3 = ". Please, complete the task fulfilment report.</B><BR>"

SIT_3DaysForTaskCompletionBody1 = "<BR><B>Dear colleague, we hereby inform you that ("
SIT_3DaysForTaskCompletionBody2 = ") the task fulfilment time limit expires in 3 days (the task card and the link to it in the document management system are below). The task is now being checked by "
SIT_3DaysForTaskCompletionBody3 = ". Please, complete the task fulfilment report.</B><BR>"

SIT_TheDayOfTaskCompletionSubj = "Task fulfilment deadline today!"

SIT_AfterTheDayOfTaskCompletionSubj = "Task fulfilment has expired!"

SIT_TheDayOfTaskCompletionBody1 = "<BR><B>Dear colleague! It is ("
SIT_TheDayOfTaskCompletionBody2 = ") the task fulfilment report deadline today (the task card and the link to it in the document management system are below). The task is now being fulfilled by "
SIT_TheDayOfTaskCompletionBody3 = " and checked by "

SIT_AfterTheDayOfTaskCompletionBody1 = "<BR><B>Dear colleague! "
SIT_AfterTheDayOfTaskCompletionBody2 = " the task fulfilment report has expired (the task card and the link to it in the document management system are below). The task is now being fulfilled by "
SIT_AfterTheDayOfTaskCompletionBody3 = " and checked by "


'vnik_fix_begin
SIT_TheDayOfTaskCompletionBody4 = ". Please create the task fulfilment report. If the report is not created within 5 calendar days, punitive sanctions will be applied to you.</B><BR>"
'vnik_fix_end
SIT_TaskCompletionDelayedSubj = "Based on the sanctions, your bonus amount will be decreased."
'vnik_fix_begin
SIT_TaskCompletionDelayedBody = "<BR><B>Dear colleague, as you have not met the task fulfilment deadline (the task card and the link to it in the document management system are below) and have not submitted the order fulfilment report within 5 days of the deadline date (including the report submission date set by the manager), sanctions in compliance with the order № 153 from 30.09.2008 will be applied to you. The sanctions will impact the amount of your quarterly bonus. Nevertheless, this does not relieve you of the duty to fulfil the order and submit the order fulfilment report.</B><BR>"
'vnik_fix_end
SIT_TaskCompletionExpiredSubj = "The task fulfilment deadline has expired. Your superior will be informed about it."
SIT_TaskCompletionExpiredBody = "<BR><B>Dear colleague, we hereby inform you that the task fulfilment deadline has expired (the task card and the link to it in the document management system are below). Please, complete the task fulfilment report.</B><BR>"
SIT_PaymentCompletionExpiredSubj = "The order fulfilment deadline has expired. Your superior will be informed about it."
SIT_TaskCompletionExpiresSoonSubj = "The task fulfilment deadline is coming."
SIT_TaskCompletionExpiresSoonBody = "<BR><B>Dear colleague, we hereby inform you that the task fulfilment deadline is coming (the task card and the link to it in the document management system are below). Please, complete the task fulfilment report.</B><BR>"
SIT_OneDayForPaymentCompletionSubj = "The order fulfilment deadline expires in 1 day."
SIT_OneDayForPaymentCompletionBody = "<BR><B>Dear colleague, we hereby inform you that the order fulfilment deadline is coming (the order card and the link to it in the document management system are below).</B><BR>"
SIT_UserDelayedTask1 = "Employee "
SIT_UserDelayedTask2 = " did not fulfil the task by the set deadline."
SIT_UserDelayedPayment1 = "Employee "
SIT_UserDelayedPayment2 = " did not fulfil the order by the set deadline."
'SIT_MoreThanOneLeader1 = "There is more than one manager in the organizational unit "
'SIT_MoreThanOneLeader2 = ""
'Запрос №32 - СТС - start
'Ниже можно добавить на разных языках по аналогии с другими. #RESPONSIBLE# заменяется на имя
SIT_UserDelayedTaskWithRequestCompletedToChief = "Employee #RESPONSIBLE# did not fulfil the order by the set deadline. Employee #RESPONSIBLE# has requested the status «Completed», but it was not accepted."
SIT_UserDelayedTaskWithRequestCompletedToControl = "Employee #RESPONSIBLE# has requested the status «Completed», please accept execution."
'Запрос №32 - СТС - end

'Названия справочников ролей (присваиваются переменным в зависимости от языка и используются в функции CheckRoleExistence)
SIT_RolesDirSitronics_RU = "Роли"
'SIT_RolesDirSTS_RU = "Роли СТС RU"
SIT_RolesDirRTI = "Роли РТИ"
SIT_RolesDirMinc = "Роли MINC"
SIT_RolesDirVTSS = "Роли ВТСС"
SIT_RolesDirSTS_RU = "Роли СТС"
SIT_RolesDirSITRU_RU = "Роли СИТРУ" ' DmGorsky
SIT_RolesDirSITRU_EN = "Роли СИТРУ" ' DmGorsky
SIT_RolesDirSITRU_CZ = "Роли СИТРУ" ' DmGorsky

SIT_RolesDirSitronics_EN = "Roles"
'SIT_RolesDirSTS_EN = "Roles STS RU"
SIT_RolesDirSTS_EN = "Roles STS"
SIT_RolesDirSitronics_CZ = "Role"
'SIT_RolesDirSTS_CZ = "Role STS RU"
SIT_RolesDirSTS_CZ = "Role STS"
'Запрос №1 - СИБ - start
SIT_RolesDirSIB_RU = "Роли СИБ"
SIT_RolesDirSIB_EN = SIT_RolesDirSIB_RU
SIT_RolesDirSIB_CZ = SIT_RolesDirSIB_RU
'Запрос №1 - СИБ - end

'Запрос №11 - СТС - start
STS_ContractPaymentDirection_In_RU = "Доходный"
STS_ContractPaymentDirection_Out_RU = "Расходный"
STS_ContractPaymentDirection_Free_RU = "Безвозмездный"

STS_ContractPaymentDirection_In_EN = "Income"
STS_ContractPaymentDirection_Out_EN = "Expense"
STS_ContractPaymentDirection_Free_EN = "Gratuitous"

STS_ContractPaymentDirection_In_CZ = "Příjmová"
STS_ContractPaymentDirection_Out_CZ = "Výdajová"
STS_ContractPaymentDirection_Free_CZ = "Bezplatná"

SIT_HeadOfInitiatorsUnit_RU = """#Начальник отдела инициатора"";"
RTI_HeadOfInitiatorsUnit = """#Руководитель инициатора"";"
MINC_HeadOfInitiatorsUnit = """#Руководитель инициатора"";"
VTSS_HeadOfInitiatorsUnit = """#Руководитель инициатора"";"
SIT_DirectorOfInitiatorsDepartment_RU = """#Директор департамента инициатора"";"
SIT_DirectorOfInitiatorsDivision_RU = """#Директор Дивизиона инициатора"";" 'В СТС
SIT_VicePresidentOfInitiator_RU = """#Вице-президент инициатора"";" 'В УК
'RTI
RTI_HeadOfPurchaseCenter = """#Директор центра организации и управления закупками"";"  
RTI_HeadOfUpravDelami = """#Управляющий делами"";" 
RTI_HeadOfAccounting = """#Главный бухгалтер"";"
RTI_HeadOfUpravEconomy = """#Начальник управления экономики"";"
RTI_DVKiA = "Департамент внутреннего контроля и аудита"
'RTI


SIT_RTI_DirectorPravovogoUprav_RU = """#Начальник правового управления"";" 'В РТИ
SIT_RTI_DirectorUpravDelami_RU = """#Начальник отдела канцелярии"";" 'В РТИ
SIT_RTI_DirectorApparatGD_RU = """#Руководитель аппарата управления ГД"";" 'В РТИ
RTI_DirectorOfSecurity = """#ЗГД – Начальник управления безопасности и режима"";" 'В РТИ

SIT_HeadOfInitiatorsUnit_EN = """#Head of the initiator's sector"";"
SIT_DirectorOfInitiatorsDepartment_EN = """#Director of the initiator's department"";"
SIT_DirectorOfInitiatorsDivision_EN = """#Director of the initiator's division"";"
SIT_VicePresidentOfInitiator_EN = """#Vice-President of the initiator"";"

SIT_HeadOfInitiatorsUnit_CZ = """#Vedoucí sektoru iniciátora"";"
SIT_DirectorOfInitiatorsDepartment_CZ = """#Vedoucí oddělení iniciátora"";"
SIT_DirectorOfInitiatorsDivision_CZ = """#Ředitel divize iniciátora"";"
SIT_VicePresidentOfInitiator_CZ = """#Viceprezident iniciátora"";"
'Запрос №11 - СТС - end

'Запрос №34 - СТС - start
STS_AssistantDirector_RU = """#Заместитель Генерального директора"";"
STS_DirectorOfDirection_RU = """#Директор по направлению"";"
STS_HeadOfInitiatorsGroup_RU = """#Начальник группы инициатора"";"
STS_AssistantDirector_EN = """#Deputy CEO"";"
STS_DirectorOfDirection_EN = """#Line of Business Director"";"
STS_HeadOfInitiatorsGroup_EN = """#Manager of initiator's group"";"
STS_AssistantDirector_CZ = """#Zástupce generálního ředitele"";"
STS_DirectorOfDirection_CZ = """#Ředitel divize"";"
STS_HeadOfInitiatorsGroup_CZ = """#Team Leader iniciátor"";"
'Запрос №34 - СТС - end

'rmanyushin 136964 08.11.2010 Start
STS_SecurityManager = """#Начальник Группы безопасности"";"
STS_HRDirector = """#Директор по персоналу"";"
'Запрос №34 - СТС - start - добавлен чуть ниже в трехъязычном варианте, оставшиеся две роли выше пока нет, все равно нет переводов
'STS_Overtime_Requester ="""#Заказчик переработки"";"
'Запрос №34 - СТС - end
'rmanyushin 136964 08.11.2010 End

'Запрос №34 - СТС - start
STS_ProjectManager_RU = """#Владелец бюджета проекта"";"
STS_ProjectManager_EN = """#Owner of the project budget"";"
STS_ProjectManager_CZ = """#Vlastnik rozpoctu projektu"";"

STS_Overtime_Requester_RU = """#Заказчик переработки"";"
STS_Overtime_Requester_EN = """#Заказчик переработки"";"
STS_Overtime_Requester_CZ = """#Заказчик переработки"";"

STS_DirectorOfProjectManagersDepartment_RU = """#Директор департамента руководителя проекта"";"
STS_DirectorOfProjectManagersDepartment_EN = """#Директор департамента руководителя проекта"";"
STS_DirectorOfProjectManagersDepartment_CZ = """#Директор департамента руководителя проекта"";"

STS_DirectorOfOvertimeRequestersDepartment_RU = """#Директор департамента заказчика переработки"";"
STS_DirectorOfOvertimeRequestersDepartment_EN = """#Директор департамента заказчика переработки"";"
STS_DirectorOfOvertimeRequestersDepartment_CZ = """#Директор департамента заказчика переработки"";"

STS_DirectorOfOvertimeRequestersDirection_RU = """#Директор по направлению заказчика переработки"";"
STS_DirectorOfOvertimeRequestersDirection_EN = """#Директор по направлению заказчика переработки"";"
STS_DirectorOfOvertimeRequestersDirection_CZ = """#Директор по направлению заказчика переработки"";"

STS_OvertimeRequestersAssistantDirector_RU = """#Заместитель генерального директора заказчика переработки"";"
STS_OvertimeRequestersAssistantDirector_EN = """#Заместитель генерального директора заказчика переработки"";"
STS_OvertimeRequestersAssistantDirector_CZ = """#Заместитель генерального директора заказчика переработки"";"
'Запрос №34 - СТС - end

'Запрос №36 - СТС - start
STS_CostCenterDirectorOfDepartment_RU = """#Директор департамента центра затрат"";"
STS_CostCenterDirectorOfDepartment_EN = """#Cost center department director"";"
STS_CostCenterDirectorOfDepartment_CZ = """#Директор департамента центра затрат"";"

STS_CostCenterDirectorOfDivision_RU = """#Директор дивизиона центра затрат"";"
STS_CostCenterDirectorOfDivision_EN = """#Cost center division director"";"
STS_CostCenterDirectorOfDivision_CZ = """#Директор дивизиона центра затрат"";"

STS_DirectorOfProjectManagersDivision_RU = """#Директор Дивизиона руководителя проекта"";"
STS_DirectorOfProjectManagersDivision_EN = """#Директор Дивизиона руководителя проекта"";"
STS_DirectorOfProjectManagersDivision_CZ = """#Директор Дивизиона руководителя проекта"";"
'Запрос №36 - СТС - end

'Запрос №46 - СТС - start
STS_OvertimeFuncLeader_RU = """#Функциональный руководитель"";"
STS_OvertimeFuncLeader_EN = """#Функциональный руководитель"";"
STS_OvertimeFuncLeader_CZ = """#Функциональный руководитель"";"
STS_Initiator_RU = """#Инициатор"";"
STS_Initiator_EN = """#Инициатор"";"
STS_Initiator_CZ = """#Инициатор"";"
'Запрос №46 - СТС - end

'Запрос №1 - СИБ - start - только русский
SIB_SLUZH_ZAPISKA_OVERTIME_PRIKAZ_TEXT = "В связи с / в целях ... ,"+VbCrLf+VbCrLf+"прошу привлечь к работе в выходные дни по письменному согласию нижеперечисленных работников ООО ""СИТРОНИКС Башкортостан"" с оплатой в соответствии с Трудовым законодательством:"
SIB_SLUZH_ZAPISKA_OVERTIME_TITLE = "О привлечении работников к работе в выходные дни"
SIB_HeadOfSector = """#Начальник отдела инициатора"";"
SIB_HeadOfDepartment = """#Директор департамента инициатора"";"
SIB_AssistantDirector = """#Заместитель генерального директора инициатора"";"
SIB_GenDirector = """#Генеральный директор"";"
SITRU_GenDirector = """#Генеральный директор СИТРУ"";" ' DmGorsky_3
SIB_Registrar = """#Регистратор"";"
SIB_HeadOfSectorIS = """#Начальник отдела эксплуатации ИС"";"
SIB_HRManager = """#Менеджер по персоналу"";"
SIB_HeadAccounting = """#Главный бухгалтер"";"
SIB_AssistantDirectorCorpDevelopment = """#Зам. генерального директора по корпоративному развитию"";"
'Запрос №1 - СИБ - end
'Запрос №46 - СТС - start - только английский
STS_AutoCancelledOvertimeMemo = "Automatic cancellation of not approved memo"
'Запрос №46 - СТС - end

Select Case UCase(Request("l"))
  Case "RU" '-------------------------------------------------------------------- РУССКИЙ
SIT_ErrorUnrecognizedRoles = "ОШИБКА! Есть нераспознанные роли в маршруте следования документа"
SIT_NotificationDocReconciled1 = "Документ "
SIT_NotificationDocReconciled2 = " согласован пользователем "
SIT_NotificationDocReconciled3 = " с КОММЕНТАРИЯМИ."
SIT_NotificationDocReconciled4 = " Документ готов к подписанию!"
SIT_NotificationLetterForYou = "Вам направлено письмо из другой организации"
SIT_ErrorSumExceeding = "<font color = red>Сумма оплат превышает стоимость закупки</font>"
SIT_NotificationApprovalChanged1 = "Изменен подписант "
SIT_NotificationApprovalChanged2 = " --> "
SIT_ErrorInDepartmentCode1 = "У подразделения пользователя: "
SIT_ErrorInDepartmentCode2 = " отсутствует код (либо неверно указано подразделение). Обратитесь к администратору системы!"
SIT_ErrorInDateCompletion = "Неверная дата исполнения: "
SIT_ErrorInProjectNumber = "Ошибка в указании номера проекта."
SIT_ErrorInBU = "Значение Бизнес единицы не соответствует справочнику"
SIT_ErrorInChartOfAccount = "Значение Статьи расходов не соответствует справочнику"
SIT_ErrorInPaymentType = "Значение Формы расчета не соответствует справочнику"
SIT_ErrorInCostCenter = "Значение Центра затрат не соответствует справочнику"
SIT_NoHeaderInResponsibleDepartment = "Не найден руководитель подразделения-исполнителя"
SIT_ErrorInSumOrCurrency = "<font color = red>ВНИМАНИЕ! </font>Ошибка в указании суммы платежа или валюты"
SIT_ErrorInProjectManager = "ВНИМАНИЕ! У менеджера проекта нет логина"
SIT_ErrorNotAllUsersFound = "ВНИМАНИЕ! Найдены не все участники согласования. Обратитесь к администратору"
SIT_ParentOrderCanceled = "Родительская заявка на закупку отменена"
SIT_ParentOrderNotApproved = "Родительская заявка на закупку не утверждена"
SIT_ErrorInParentOrderNumber = "Ошибка в указании номера родительской заявки"
SIT_AdditionalAgrees = "Дополнительные согласующие: "
SIT_RequiredAgrees = "Обязательные согласующие: "
SIT_PreliminaryAgrees = "Предварительное согласование: "
SIT_BUT_AFKTasks = "ПОРУЧЕНИЯ АФК"
SIT_BUT_AFKTasksHint = SIT_BUT_AFKTasks
SIT_CentralBut_CoResponsibleHint = "Неисполненные поручения, где я соисполнитель"
SIT_CentralBut_PurchaseOrdersHint = "Незакрытые заявки на закупку"
SIT_LastReconcilationDate = "Срок согласования: "
SIT_NotificationApproveRefused = "Отказано в утверждении"
SIT_NotificationAgreeRefusedByUser = "Отказано в согласовании пользователем "
SIT_ButtonLoadProjects = "Загрузить проекты"
SIT_ButtonLoadProjectsHint = "Загрузить справочник проектов из XML"
SIT_ButtonLoadProjectsConfirm = "Загрузить справочник проектов из XML-файла?"
SIT_ErrorInUserField1 = "Значение в поле "
SIT_ErrorInUserField2 = " выбрано не из справочника"
'Запрос №11 - СТС - start
SIT_ErrorNoPartnerCode = "Отсутствует код контрагента. Укажите буквенный код в справочнике Контрагенты"
SIT_CannotCreateOldContracts = "В этой категории запрещено создавать документы, пользуйтесь категорией Договоры"
'Запрос №11 - СТС - end
'Запрос №43 - СТС - start
STS_ErrorCantCreatePOLargerLimit1 = "Вы не имеете права создавать заявки на закупку свыше $10000"
STS_ErrorCantCreatePOLargerLimit2 = "Вы не имеете права создавать заявки на закупку свыше $50000"
STS_ErrorPartnerIsNotChecked = "Нельзя создавать заявки на закупку с непроверенными контрагентами"
'Запрос №43 - СТС - end
'Запрос №46 - СТС - start
SIT_CannotCreateDocInThisCategory = "В этой категории запрещено создавать документы"
SIT_ErrorInProjectNumbers = "Ошибка в указании номеров проектов: "
SIT_ParentDocCancelled = "Родительский документ отменен"
SIT_ParentDocNotApproved = "Родительский документ не утвержден"
SIT_ErrorInParentDocNumber = "Ошибка в указании номера родительского документа"
STS_SLUZH_ZAPISKA_OVERTIME_PLAN_TITLE = "О привлечении работников к работе в выходные дни"
STS_SLUZH_ZAPISKA_OVERTIME_PLAN_TEXT = "В связи с проведением работ по проекту … прошу привлечь к работе в выходные/праздничные дни по письменному согласию нижеперечисленных работников ЗАО «СИТРОНИКС Телеком Солюшнс» с оплатой в соответствии с Трудовым законодательством:"
'Запрос №46 - СТС - end

'20100111 - Запрос №13 из СТС - start
STS_YouAreResponsibleSubject = "Вы назначены исполнителем по документу"
STS_YouAreResponsibleBody = "Информируем Вас о том, что Вы назначены исполнителем по данному документу. Если у Вас возникли вопросы, свяжитесь, пожалуйста, с инициатором данного документа или системным администратором."
'20100111 - Запрос №13 из СТС - end

STS_HeadOfSector = """#Начальник отдела"";"
STS_HeadOfDepartment = """#Директор департамента"";"
STS_HeadOfDivision = """#Директор Дивизиона"";"
STS_FinancialControl = """#Финансовый Контролер"";"
STS_FinDirector = """#Финансовый Директор"";"
STS_GenDirector = """#Генеральный Директор"";"
'Запрос №34 - СТС - start
'STS_ProjectManager = """#Владелец бюджета проекта"";"
STS_ProjectManager = STS_ProjectManager_RU
STS_Overtime_Requester = STS_Overtime_Requester_RU
STS_DirectorOfProjectManagersDepartment = STS_DirectorOfProjectManagersDepartment_RU
STS_DirectorOfOvertimeRequestersDepartment = STS_DirectorOfOvertimeRequestersDepartment_RU
STS_DirectorOfOvertimeRequestersDirection = STS_DirectorOfOvertimeRequestersDirection_RU
STS_OvertimeRequestersAssistantDirector = STS_OvertimeRequestersAssistantDirector_RU
'Запрос №34 - СТС - end
'Запрос №36 - СТС - start
STS_CostCenterDirectorOfDepartment = STS_CostCenterDirectorOfDepartment_RU
STS_CostCenterDirectorOfDivision = STS_CostCenterDirectorOfDivision_RU
STS_DirectorOfProjectManagersDivision = STS_DirectorOfProjectManagersDivision_RU
'Запрос №36 - СТС - end
'Запрос №46 - СТС - start
STS_OvertimeFuncLeader = STS_OvertimeFuncLeader_RU
STS_Initiator = STS_Initiator_RU
'Запрос №46 - СТС - end
'Запрос №38 - СТС - start
SIT_MaxUsersInListToReconcileExceeded = "Превышено максимальное количество согласующих. Согласующих должно быть не более #MAX. Сократите количество согласующих и попробуйте сохранить карточку."
SIT_CannotAddUsersToListToReconcile = "Добавление пользователей невозможно, достигнуто предельное число согласующих."
'Запрос №38 - СТС - end
STS_Accounting = """#Бухгалтерия"";"
STS_Treasury = """#Казначейство"";"
STS_SecrPravlenia = """#Секретарь правления СТС"";"
SIT_SecrPravlenia = """#Секретарь правления УК"";"
SIT_Registrar = """#Регистратор"";"

SITRU_Registrar = """#Регистратор СИТ"";" ' DmGorsky
SITOAORTI_Registar = """#Регистратор"";"
SIT_ReportControl = "Контроль"
SIT_ReportControlValues = "На контроле, Без контроля"
SIT_ReportStartingDate = "Начальная дата"
SIT_ReportFinishingDate = "Конечная дата"
SIT_ReportTaskStartingDateFrom = "Дата выдачи поручения (от)"
SIT_ReportTaskStartingDateTo = "Дата выдачи поручения (до)"
SIT_ReportTypeOfRequest = "Тип запроса"
SIT_ReportRequestTypes = "Исполнено в срок,Исполнено с нарушением срока,Не исполнено(срок не наступил),Не исполнено(с нарушением срока)"

SIT_ErrorInFieldValue1 = "<font color=red>Ошибка!</font> Необходимо выбрать значения для поля ["
SIT_ErrorInFieldValue2 = "] из справочника"

SIT_ReportTypeOfTask = "Вид задания"
SIT_TaskName = "Текст резолюции / поручения"

'Запрос №11 - СТС - start
SIT_HeadOfInitiatorsUnit = SIT_HeadOfInitiatorsUnit_RU
SIT_DirectorOfInitiatorsDepartment = SIT_DirectorOfInitiatorsDepartment_RU
SIT_DirectorOfInitiatorsDivision = SIT_DirectorOfInitiatorsDivision_RU 'В СТС
SIT_VicePresidentOfInitiator = SIT_VicePresidentOfInitiator_RU 'В УК
'SIT_HeadOfInitiatorsUnit = """#Начальник отдела инициатора"";"
'SIT_DirectorOfInitiatorsDepartment = """#Директор департамента инициатора"";"
'SIT_DirectorOfInitiatorsDivision = """#Директор Дивизиона инициатора"";" 'В СТС
'SIT_VicePresidentOfInitiator = """#Вице-президент инициатора"";" 'В УК
'Запрос №11 - СТС - end
SIT_President = """#Президент"";"
RTI_President = """#Генеральный директор"";"
MINC_Director = """#Генеральный директор"";"

SIT_FinancialVicePresident = """#Вице-президент по финансам и инвестициям"";"
SIT_HRDirector = """#Директор по персоналу"";"
SIT_GenDirector = """#Генеральный Директор"";"
'vnik_protocols
SIT_ChairmanOfBoard = """#Председатель Правления"";"
SIT_ChairmanOfCommitteeOnIT = """#Председатель Комитета по ИТ"";"
SIT_ChairmanOfCommitteeOnControlAndAuditing  = """#Председатель контрольно-ревизионного комитета"";"
SIT_ChairmanOfManagingCommitteeOnTheProgramEGRB = """#Председатель управляющего комитета по программе ЭПРБ"";"
SIT_SecretaryOfBoard = """#Секретарь Правления"";"
SIT_SecretaryOfCommitteeOnIT = """#Секретарь Комитета по ИТ"";"
SIT_SecretaryOfCommitteeOnControlAndAuditing  = """#Секретарь контрольно-ревизионного комитета"";"
SIT_SecretaryOfManagingCommitteeOnTheProgramEGRB = """#Секретарь управляющего комитета по программе ЭПРБ"";"
'vnik_protocols

'vnik_send_notification
SIT_AttachFilesToNotofication = """#Список получателей уведомлений с приложенными файлами документа"";"
'vnik_send_notification

'vnik_payment_order
SIT_BudgetController = """#Бюджетный контролер"";"
SIT_TreasuryController = """#Казначейский контролер"";"
SIT_Docs_Priority = "Приоритет платежа"
SIT_Budget_classification = "Бюджетный классификатор"
SIT_Account_manager = """#Бухгалтер по заявкам на оплату"";"
'vnik_payment_order

'vnik_purchase_order
SIT_SecretaryCPC = """#Секретарь ЦЗК"";"
SIT_SecurityController = """#Контролер по безопасности"";" 
'vnik_purchase_order

'rti_rasp_docs
RTI_RaspDocsViewList = """#Получатели распорядительных документов РТИ"";"
'rti_rasp_docs

'rti_contract
RTI_ContractViewList = """#Правовое управление РТИ"";"
'rti_contract


'rti_purchase_order
RTI_ChiefOfPurchaseDepartment = """#Начальник Управления планирования и управления отчетности"";"
RTI_BudgetController = """#Руководитель по бюджетному контролю"";"
RTI_HeadPriceforming = """#Начальник отдела ценообразования"";"
'rti_purchase_order

'rti_payment_order
RTI_HeadKFIE = """#ЗГД - Руководитель КФИЭ"";"
RTI_PaymentChief = """#Бухгалтер по оплате"";"
RTI_PaymentHead = """#Руководитель по казначейству"";"
'rti_payment_order

'vnik_protocolsCPC
SIT_ChairmanOfCentralPurchasingCommission = """#Председатель Центральной Закупочной Комиссии"";"
SIT_ChiefOfPurchaseDepartment = """#Начальник отдела закупок"";"
'vnik_protocolsCPC

'rti_protocol
RTI_ChairmanOfCPC = """#Председатель ЦЗК"";" 
RTI_SecretaryOfCPC = """#Секретарь ЦЗК"";"
'rti_protocol

'oaorti_roles
OAORTI_GeneralDirector = """#Генеральный директор"";"
'oaorti_roles


'vnik_contracts
SIT_SignatoryOfTheContractsMC = """#Подписант договоров УК по умолчанию"";"
SIT_OldContractOperator = """#Право на ввод старых договоров УК"";"
'vnik_contracts

'vnik_archive
SIT_AccessToArchive = """#Доступ в архив"";"
'vnik_archive

SIT_LinkInstruction = "Инструкция"
SIT_LinkInstructionHint = "Руководство по работе с системой"
SIT_LinkInstructionFile = "Manuals/UserInstruction.doc" 'Имя файла с инструкцией на данном языке

SIT_Initiator = "Инициатор"
SIT_CurrencyRateToUSD = "Курс валюты относительно USD"
SIT_Orders_Prikaz = "OR - приказ"
SIT_Orders_Rasporyajenie = "Р - распоряжение"
SIT_Orders_Prikaz_RTI = "OR - приказ ОАО ""РТИ"""
SIT_Orders_Rasporyajenie_RTI = "Р - распоряжение"
SIT_Orders_Prikaz_MIKRON = "OR - приказ ОАО ""НИИМЭ и Микрон"""
SIT_Orders_Prikaz_NIIME = "OR - приказ АО ""НИИМЭ"""

'vnik_rasp_norm_doc
SIT_Orders_Prikaz_ND = "OR - Приказ об утверждении Нормативного документа" 'при изменении данного наименования поменять условие запроса в файле AgentUser.asp Function NormatDocsAutoReconcilation()
SIT_NORM_DOCS_CLOSED = "Категория Нормативные документы закрыта для создания новых документов с 01.07.2010, используйте категорию распорядительных документов с видом документа Приказ на утверждение нормативного документа!!!"
SIT_NORM_DOCS_WARNING = "Крайний срок согласования должен быть больше или равен "
'vnik_rasp_norm_doc
'vnik_protocolsCPC
SIT_PROTOCOLS_CPC_RESTRICT = "Документ Протокол ЦЗК может создать только пользователь входящий в роль Секретарь ЦЗК!!!"
'vnik_protocolsCPC

'01.07.2013 kkoshkin sts_purchase_payment_order_restrict
STS_ORDER_RESTRICT = "Пользователям Вашего БН запрещено создавать заявки на закупку и оплату!!!"
STS_ORDER_BN_RESTRICT = "1010 - СИТРОНИКС Телеком Солюшнс, ЗАО 1011 - Дальневосточный филиал 1012 - Уральский филиал 1013 - Филиал МЕДИАТЕЛ-Кубань 1014 - Орловский филиал 1015 - Поволжский филиал 1016 - Новосибирский филиал 1017 - Санкт-Петербургский филиал 1018 - Нижегородский филиал 1060 - СИТРОНИКС Телеком Софтвэа "
'01.07.2013 kkoshkin sts_purchase_payment_order_restrict
'vnik_contracts
SIT_CONTRACTS_MC_WARNING1 = "Договор УК можно создать только на основании подписанной Заявки на закупку УК при условии (Сумма заявки < 300 000 рублей) или на основании зарегистрированного Протокола ЦЗК!!!"
SIT_CONTRACTS_MC_WARNING2 = "Не найден документ основание или документ основание не подписан/зарегистрирован!!!"
SIT_CONTRACTS_MC_WARNING3 = "Протокол ЦЗК введен не на основании Заявки на закупку УК или Заявка на закупку УК не найдена!!!"
SIT_CONTRACTS_MC_WARNING4 = "Для данной валюты не прописаны условия, обратитесь к администратору системы!!!"
SIT_CONTRACTS_MC_WARNING5 = "Договора УК по Заявке на закупку УК суммой свыше 300 000 рублей или её эквиваленту в другой валюте создаются на основании Протокола ЦЗК!!!"
SIT_CONTRACTS_MC_WARNING6 = "Договор УК без документа основания можно создать только с видом Рамочный, без возможности указать сумму договора отличную от 0!!!"
'vnik_contracts
'vnik_payment_order
SIT_PAYMENT_ORDER_WARNING1 = "Заявку на оплату УК превышающую по сумме 30 000 рублей или её эквиваленту в другой валюте создают на основании зарегистрированного Договора УК, укажите документ основание!!!"
SIT_PAYMENT_ORDER_WARNING2 = "Некорректно указан документ основание, Заявку на оплату УК превышающую по сумме 30 000 рублей или её эквиваленту в другой валюте создают на основании зарегистрированного Договора УК!!!"
SIT_PAYMENT_ORDER_WARNING3 = "На основании Рамочного договора УК нельзя создавать заявки на оплату УК!!!"
'vnik_payment_order

'rti_payment_order
RTI_PAYMENT_ORDER_WARNING1 = "Заявку на оплату РТИ можно создавать только на основе утвержденой заявки на закупку!!!"
RTI_PAYMENT_ORDER_WARNING2 = "Заявку на оплату РТИ можно создавать только на основе утвержденой заявки на закупку с суммой менее 59000 рублей (для суммы более 59000 заявка на оплату создается на основе документа БСАП)!!!"
RTI_PAYMENT_ORDER_WARNING3 = "Заявка на оплату РТИ по статье расходов ""13072 Канцелярские и хозяйственные товары, офисные принадлежности"" создается организатором закупки!!!"
'rti_payment_order

'rti_protocol
RTI_PROTOCOL_WARNING1 = "Протокол ЦЗК РТИ можно создавать только на основе утвержденной заявки на закупку с суммой свыше 354000 рублей!!!"
RTI_PROTOCOL_WARNING2 = "Протокол ЦЗК РТИ может создавать только Секретарь ЦЗК!!!"
'rti_protocol

'rti_bsap
RTI_BSAP_WARNING1 = "Документ БСАП можно создавать только на основе утвержденной заявки на закупку с суммой от 59000 до 354000 рублей (при сумме менее 59000 следует создавать заявку на оплату, при сумме свыше 354000 следует создавать протокол ЦЗК)!!!"
'rti_bsap

'rti_contract
RTI_CONTRACT_WARNING1 = "Договор РТИ можно создавать только на основе утвержденной заявки на закупку, либо сам по себе!!!"
RTI_CONTRACT_WARNING2 = "Протокол ЦЗК введен не на основании Заявки на закупку РТИ или Заявка на закупку РТИ не найдена!!!"
RTI_CONTRACT_WARNING3 = "Договора РТИ по Заявке на закупку РТИ суммой свыше 354 000 рублей создаются на основании Протокола ЦЗК!!!"
'rti_contract

'
RTI_SIT_VHODYASCHIE_RESTRICT = "Входящие документы ОАО ""РТИ"" может создавать только сотрудник канцелярии!!!"
'AMW - MIKRON - Start
SIT_RolesDirMIKRON = "Роли МИКРОН"
MIKRON_GenDirector = """#Генеральный директор"";"
MIKRON_GenDesigner = """#Генеральный конструктор"";"
MIKRON_MainAuditor = """#Руководитель по внутреннему контролю"";"
MIKRON_MainDesigner = """#Главный конструктор"";"
MIKRON_MainAccountant = """#Главный бухгалтер"";"
MIKRON_DeputyGDmarketing = """#ЗГД - по маркетингу"";"
MIKRON_DeputySecurityHead = """#Зам. директора по безопасности"";"
MIKRON_HeadOfInitiatorUnit = """#Руководитель направления"";"
MIKRON_Aprovals_Contracts = MIKRON_GenDirector + vbCrLf + """#ЗГД - по производству"";" + vbCrLf+ _
      """#ЗГД - по науке"";" + vbCrLf + """#ЗГД - по маркетингу"";" + vbCrLf+ _
      """#1-й заместитель Ген.директора"";" + vbCrLf + """#Директор по персоналу"";" + vbCrLf + """#Управляющий делами"";"
MIKRON_Aprovals_NDA = MIKRON_GenDirector + vbCrLf + """#ЗГД - по науке"";" + vbCrLf + """#ЗГД - по маркетингу"";" + vbCrLf+ _
MIKRON_Overseer = """Контролер МИКРОН"" <OverseerMIKRON>;" 'Account пользовательской роли - "Контролер МИКРОН" в справочнике пользователей
MIKRON_Auditor  = """Аудитор МИКРОН"" <AuditorMIKRON>;"    'Account пользовательской роли - "Аудитор МИКРОН" в справочнике пользователей
MIKRON_Legal    = """Юрист МИКРОН"" <LegalMIKRON>;"        'Account пользовательской роли - "Юрист МИКРОН" в справочнике пользователей
'rmanyushin 133266 05.10.2010 End
 
'purchase_order
MIKRON_ChiefOfPurchaseDepartment = """#Начальник отдела обеспечения"";"
MIKRON_BudgetController = """#Руководитель по бюджетному контролю"";"

'payment_order
MIKRON_HeadKFIE = """#ЗГД - по Финансам и Инвестициям"";"
MIKRON_PaymentChief = """#Начальник бюджетного отдела"";"
MIKRON_SalesChief = """#Директор по продажам"";"
MIKRON_SalesAgrees = "Согласующие договор на продажу: "
MIKRON_CFO = """#Финансовый директор"";"
MIKRON_PAYMENT_ORDER_WARNING1 = "Заявку на оплату МИКРОН можно создавать только на основе утвержденой заявки на закупку."
MIKRON_PAYMENT_ORDER_WARNING2 = "Заявку на оплату МИКРОН можно создавать только на основе утвержденой заявки на закупку с суммой менее 59,000 рублей (для суммы более 59,000 заявка на оплату создается на основе документа БСАП)"

'МИКРОН Протокол ЗК
MIKRON_ChairmanOfPC = """#Председатель ЗК"";" 
MIKRON_SecretaryOfPC = """#Секретарь ЗК"";"
MIKRON_Form_OfPC_1 = "Заочная"
MIKRON_Form_OfPC_2 = "Очная"
MIKRON_Form_OfPC = MIKRON_Form_OfPC_1 + vbCrLf + MIKRON_Form_OfPC_2
MIKRON_CHOISE_PC = "Выбор предложения от поставщика №1,Выбор предложения от поставщика №2,Выбор предложения от поставщика №3"

'WARNING - Start
MIKRON_PROTOCOL_WARNING1 = "Протокол ЗК МИКРОН можно создавать только на основе утвержденной заявки на закупку с суммой свыше 500,000 рублей!!!"
MIKRON_PROTOCOL_WARNING2 = "Протокол ЗК МИКРОН может создавать только Секретарь ЗК"
MIKRON_PROTOCOL_WARNING3 = "Протокол ЗК МИКРОН можно создавать только на основе утвержденного Опросного Листа"

'WARNING
MIKRON_DOCNOTFOUND_WARNING1 = "Документ основание не найден."
MIKRON_WRONGSUM_WARNING1 = "Сумма неправильная"
'Mikron_BSAP
MIKRON_BSAP_WARNING1 = "Документ БСАП можно создавать только на основе утвержденной заявки на закупку с суммой от 50,000 до 500,000 рублей (при сумме менее 50,000 следует создавать заявку на оплату, при сумме свыше 500,000 следует создавать протокол ЗК)"
'Mikron_contract
MIKRON_CONTRACT_WARNING1 = "Договор МИКРОН можно создавать только на основе утвержденного Протокола ЗК или документа БСАП."
MIKRON_CONTRACT_WARNING2 = "Протокол ЗК введен не на основании Заявки на закупку МИКРОН или Заявка на закупку МИКРОН не найдена."
MIKRON_CONTRACT_WARNING3 = "Договор МИКРОН по Заявке на закупку МИКРОН суммой свыше 500,000 рублей создается на основании Протокола ЗК"
MIKRON_CONTRACT_WARNING4 = "Создавать Договор МИКРОН можно только типа <font color = red>ДОХОДНЫЙ</font> или <font color = red>БЕЗВОЗМЕЗДНЫЙ</font>."+vbCrLf+"Договор на закупку создается на основе утвержденного Протокола ЗК или документа БСАП."
MIKRON_CONTRACT_WARNING5 = "Создавать Доп.соглашения можно только на основе <font color = red>СУЩЕСТВУЮЩИХ</font> в системе документов категории <font color = red>ДОГОВОРЫ</font>."
MIKRON_CONTRACT_WARNING6 = "Доп.соглашения c <font color = red>УВЕЛИЧЕНИЕМ</font> стоимости создаются на основе <font color = red>СПРАВКИ О ЦЕНАХ</font>."
'Mikron_export
MIKRON_EXP_WARNING1 = "Дополнение к экспортному контракту можно создавать только на основе утвержденного контракта"
'WARNING - End
MIKRON_BU_1 = "ОАО ""НИИМЭ и Микрон"""
MIKRON_BU_2 = "АО ""НИИМЭ"""
MIKRON_BU_3 = "ЗАО ""РТИ-Микроэлектроника"""
MIKRON_BUs = MIKRON_BU_1 + vbCrLf + MIKRON_BU_2 + vbCrLf + MIKRON_BU_3
MIKRON_DEAL_1 = "Экспорт"
MIKRON_DEAL_2 = "Импорт"
MIKRON_DEAL_3 = "Внутренний рынок"
MIKRON_DEAL_TYPES = MIKRON_DEAL_1 + vbCrLf + MIKRON_DEAL_2 + vbCrLf + MIKRON_DEAL_3
MIKRON_TEXT_KP_SELECT = "-цена; -качество; -условия оплаты; -срок поставки/условия поставки; - сервис/надежность; - Другое:"
'AMW - MIKRON - End

SIT_BusinessUnit = "Бизнес единица"
SIT_CostCenterCode = "Код Центра Затрат"
SIT_Project = "Проект"
SIT_ExpenseItem = "Статья расходов"
SIT_ProjectCode = "Код проекта"
SIT_PaymentType = "Форма расчета"
SIT_Budgeted = "Есть в бюджете (Да\Нет)"
SIT_USDAmount = "Сумма по документу в USD"
SIT_PurchaseOrder = "Заявка на закупку"
SIT_PaymentOrder = "Заявка на оплату"

RTI_PurchaseOrder = "Заявка на закупку"
RTI_PaymentOrder = "Заявка на оплату"

SIT_Subject = "Тема"
'20090622 - Заявка ТКП
SIT_ComOfferNumber = "Номер ТКП"

SIT_RolesDirSitronics = SIT_RolesDirSitronics_RU
SIT_RolesDirSTS = SIT_RolesDirSTS_RU
'Запрос №1 - СИБ - start
SIT_RolesDirSIB = SIT_RolesDirSIB_RU
'Запрос №1 - СИБ - end

SIT_Letter_Initiative = "инициативное"
SIT_Letter_AnswerFor = "письмо-ответ на № "
SIT_SistemaTasks = "Поручения АФК"
SIT_IncomingMailTasks = "Поручения по входящим письмам"
SIT_TasksOnOrders = "Поручения по приказам/распоряжениям"
SIT_TasksOnMemos = "Поручения по служебным запискам"
SIT_TasksOnOtherDocs = "Поручения (на основании других документов или устных заданий)"
SIT_YesNo = "Да,Нет"

SIT_ErrorInDepartmentLeader = "Ошибка в указании руководителя подразделения "
SIT_DocChangedReconcilationRequired = "Документ изменен, требует согласования"
SIT_Language = "Язык"
SIT_AssistantOfPresident = """#Помощник Президента"";"
SIT_HeadOfDocControl = """#Заведующий канцелярией"";"
SIT_SortType = "Тип сортировки"
SIT_SortTypeAscending = "по возрастанию"
SIT_SortTypeDescending = "по убыванию"

SIT_MoreThanOneLeader1 = "В подразделении "
SIT_MoreThanOneLeader2 = " больше одного руководителя"
SIT_NoAccessToParentDoc = "У Вас нет доступа к родительскому документу, Вы не можете на него ссылаться"

'rmanyushin 60298, 89142 01.04.2010 Start
STS_ContractType = "Сторона Договора"
'vnik_purchase_order
SIT_ContractType = "Способ закупки"
'vnik_purchase_order
'ph - 20120216 - start
STS_OffshoreZone = "Местонахождение контрагента в оффшорной зоне"
'ph - 20120216 - end
'ph - 20101225 - start
'Делаем только на англ. будет в справочнике Ролей для PO
'STS_HRAdministrationManager = """#Начальник отдела кадрового администрирования"";"
'ph - 20101225 - end
STS_DptOSSBSS = """#Заместитель генерального директора по OSS/BSS"";"
'STS_ContractPartyDirGUID = "{1A10FE95-E22A-4413-9394-00120C0C12C1}" ' Справочник на IT-TEST-07
'STS_ContractPartyDirGUID = "{02073D43-0553-45B6-8F50-50C396DB2E14}" ' Справочник на 8888
STS_ContractPartyDirGUID = "{63E943E3-E678-4138-900B-1E5FE96809AE}" ' Справочник на GL-PAYDOX-01
'rmanyushin 60298, 89142 01.04.2010 End

'rmanyushin 93755, 22.04.2010 Start
STS_SLUZH_ZAPISKA_OVERTIME_PRIKAZ_TEXT = "В связи с / в целях ... ,"+VbCrLf+VbCrLf+"прошу привлечь к работе в выходные дни по письменному согласию нижеперечисленных работников ЗАО ""СИТРОНИКС Телеком Солюшнс"" с оплатой в соответствии с Трудовым законодательством:"
STS_SLUZH_ZAPISKA_OVERTIME_TITLE = "О привлечении работников к работе в выходные дни"
'rmanyushin 93755, 22.04.2010 End

'rmanyushin 119579 19.08.2010 Start
STS_SLUZH_ZAPISKA_HOLIDAY_TITLE = "Заявление на отпуск"
'rmanyushin 119579 19.08.2010 End

'rmanyushin 119579 19.08.2010 Start
'Запрос №34 - СТС - start
'STS_AssistantDirector = """#Заместитель Генерального директора"";"
'STS_DirectorOfDirection = """#Директор по направлению"";"
'STS_HeadOfInitiatorsGroup = """#Начальник группы инициатора"";"
STS_AssistantDirector = STS_AssistantDirector_RU
STS_DirectorOfDirection = STS_DirectorOfDirection_RU
STS_HeadOfInitiatorsGroup = STS_HeadOfInitiatorsGroup_RU
'Запрос №34 - СТС - end
'rmanyushin 119579 19.08.2010 Stop


'Запрос №11 - СТС - start
STS_ContractPaymentDirection_In = STS_ContractPaymentDirection_In_RU
STS_ContractPaymentDirection_Out = STS_ContractPaymentDirection_Out_RU
STS_ContractPaymentDirection_Free = STS_ContractPaymentDirection_Free_RU
'Запрос №11 - СТС - end
'{ph - 20120517
STS_ErrorInHolidayFirstDate = "Укажите правильную дату начала отпуска"
'ph - 20120517}

  Case "" 'EN ------------------------------------------------------- АНГЛИЙСКИЙ

SIT_ErrorUnrecognizedRoles = "ERROR! Some of the document processing roles are unrecognized"
SIT_NotificationDocReconciled1 = "Document "
SIT_NotificationDocReconciled2 = " has been approved by the user "
SIT_NotificationDocReconciled3 = " with COMMENTS."
SIT_NotificationLetterForYou = "You have received a letter from an external organization"
SIT_ErrorSumExceeding = "<font color = red>The payment amount exceeds the purchase value</font>"
SIT_NotificationApprovalChanged1 = "Signatory has been changed "
SIT_NotificationApprovalChanged2 = " --> "
SIT_ErrorInDepartmentCode1 = "For the user's organizational unit: "
SIT_ErrorInDepartmentCode2 = " the code is missing (or the an incorrect department was entered). Contact the system administrator!"
SIT_ErrorInDateCompletion = "Incorrect completion date: "
SIT_ErrorInProjectNumber = "Incorrect project number"
SIT_ErrorInBU = "The Business unit does not correspond to the catalogue"
SIT_ErrorInChartOfAccount = "The cost item does not correspond to the catalogue"
SIT_ErrorInPaymentType = "The payment method does not correspond to the catalogue"
SIT_ErrorInCostCenter = "The cost centre does not correspond to the catalogue"
SIT_NoHeaderInResponsibleDepartment = "The responsible unit head was not found"
SIT_ErrorInSumOrCurrency = "<font color = red>ATTENTION!</font>Incorrect payment amount or currency"
SIT_ErrorInProjectManager = "ATTENTION! No username for project management."
SIT_ErrorNotAllUsersFound = "ATTENTION! Some of the agreement process participants not found. Contact the administrator."
SIT_ParentOrderCanceled = "Parent purchase order was cancelled"
SIT_ParentOrderNotApproved = "Parent purchase order was not approved"
SIT_ErrorInParentOrderNumber = "Incorrect parent order number"
SIT_AdditionalAgrees = "Additional agreement process participants: "
SIT_RequiredAgrees = "Mandatory agreement process participants: "
SIT_PreliminaryAgrees = "Preliminary agreement: "
SIT_BUT_AFKTasks = "TASKS FROM AFK"
SIT_BUT_AFKTasksHint = SIT_BUT_AFKTasks
SIT_CentralBut_CoResponsibleHint = "Unfulfiled tasks, for whom I am the co-executor"
SIT_CentralBut_PurchaseOrdersHint = "Open purchase orders"
SIT_LastReconcilationDate = "Agreement deadline: "
SIT_NotificationApproveRefused = "Not approved"
SIT_NotificationAgreeRefusedByUser = "Not agreed by a user "
SIT_ButtonLoadProjects = "Download projects"
SIT_ButtonLoadProjectsHint = "Download the project catalogue from XML?"
SIT_ButtonLoadProjectsConfirm = "Download the project catalogue from the XML file?"
SIT_ErrorInUserField1 = "The entered value in the field "
SIT_ErrorInUserField2 = " was not selected from the list"
'Запрос №11 - СТС - start
SIT_ErrorNoPartnerCode = "The contract partner code is missing. Select the alphabetic code in the Partner directory."
SIT_CannotCreateOldContracts = "It’s forbidden to create documents in this category, use Contracts category"
'Запрос №11 - СТС - end
'Запрос №43 - СТС - start
STS_ErrorCantCreatePOLargerLimit1 = "You don’t have rights to create Purchase Orders  for amounts exceeding  $10000"
STS_ErrorCantCreatePOLargerLimit2 = "You don’t have rights to create Purchase Orders  for amounts exceeding  $50000"
STS_ErrorPartnerIsNotChecked = "It is not allowed to create Purchase Orders with unverified contractors"
'Запрос №43 - СТС - end
'Запрос №46 - СТС - start
SIT_CannotCreateDocInThisCategory = "It is forbidden to create documents in this section"
SIT_ErrorInProjectNumbers = "Error in project number introduction: "
SIT_ParentDocCancelled = "Parent document was cancelled"
SIT_ParentDocNotApproved = "Parent document wasn’t approved"
SIT_ErrorInParentDocNumber = "Error in parent document introduction"
STS_SLUZH_ZAPISKA_OVERTIME_PLAN_TITLE = "Concerning overtime approval"
STS_SLUZH_ZAPISKA_OVERTIME_PLAN_TEXT = "With regard to the works on the project … I request engagement of the below mentioned employees of the company SITRONICS Telecom Solutions for overtime work upon written agreement with remuneration according to the Labour Code"
'Запрос №46 - СТС - end
'20100111 - Запрос №13 из СТС - start
STS_YouAreResponsibleSubject = "You have been appointed the responsible person for this document"
STS_YouAreResponsibleBody = "We hereby inform you that you have been appointed the responsible person for this document. If you have any questions, please contact the document drafter or the system administrator."
'20100111 - Запрос №13 из СТС - end

STS_HeadOfSector = """#Sector manager"";"
STS_HeadOfDepartment = """#Department director"";"
STS_HeadOfDivision = """#Division director"";"
STS_FinancialControl = """#Financial controller"";"
STS_FinDirector = """#Financial director"";"
STS_GenDirector = """#CEO"";"
'Запрос №34 - СТС - start
'STS_ProjectManager = """#Owner of the project budget"";"
STS_ProjectManager = STS_ProjectManager_EN
STS_Overtime_Requester = STS_Overtime_Requester_EN
STS_DirectorOfProjectManagersDepartment = STS_DirectorOfProjectManagersDepartment_EN
STS_DirectorOfOvertimeRequestersDepartment = STS_DirectorOfOvertimeRequestersDepartment_EN
STS_DirectorOfOvertimeRequestersDirection = STS_DirectorOfOvertimeRequestersDirection_EN
STS_OvertimeRequestersAssistantDirector = STS_OvertimeRequestersAssistantDirector_EN
'Запрос №34 - СТС - end
'Запрос №36 - СТС - start
STS_CostCenterDirectorOfDepartment = STS_CostCenterDirectorOfDepartment_EN
STS_CostCenterDirectorOfDivision = STS_CostCenterDirectorOfDivision_EN
STS_DirectorOfProjectManagersDivision = STS_DirectorOfProjectManagersDivision_EN
'Запрос №36 - СТС - end
'Запрос №46 - СТС - start
STS_OvertimeFuncLeader = STS_OvertimeFuncLeader_EN
STS_Initiator = STS_Initiator_EN
'Запрос №46 - СТС - end
STS_Accounting = """#Accounting department"";"
STS_Treasury = """#Treasury"";"
STS_SecrPravlenia = """#STS management secretary"";"
SIT_SecrPravlenia = """#Sitronics management secretary"";"
SIT_Registrar = """#Registrar"";"

SIT_ReportControl = "Check"
SIT_ReportControlValues = "Checked, Without check"
SIT_ReportStartingDate = "Start date"
SIT_ReportFinishingDate = "End date"
SIT_ReportTaskStartingDateFrom = "The task setting date (from)"
SIT_ReportTaskStartingDateTo = "The task setting date (until)"
SIT_ReportTypeOfRequest = "Request type"
SIT_ReportRequestTypes = "Fulfiled by the deadline,Fulfiled but deadline not met,Not fulfiled (before the deadline),Not fulfiled (deadline not met)"

SIT_ErrorInFieldValue1 = "<font color=red>Error!</font> Values in the field ["
SIT_ErrorInFieldValue2 = "] should be selected from the directory"

SIT_ReportTypeOfTask = "Type of task"
SIT_TaskName = "Name"

'Запрос №11 - СТС - start
SIT_HeadOfInitiatorsUnit = SIT_HeadOfInitiatorsUnit_EN
SIT_DirectorOfInitiatorsDepartment = SIT_DirectorOfInitiatorsDepartment_EN
SIT_DirectorOfInitiatorsDivision = SIT_DirectorOfInitiatorsDivision_EN 'В СТС
SIT_VicePresidentOfInitiator = SIT_VicePresidentOfInitiator_EN 'В УК
'SIT_HeadOfInitiatorsUnit = """#Head of the initiator's sector"";"
'SIT_DirectorOfInitiatorsDepartment = """#Director of the initiator's department"";"
'SIT_DirectorOfInitiatorsDivision = """#Director of the initiator's division"";"
'SIT_VicePresidentOfInitiator = """#Vice-President of the initiator"";"
'Запрос №11 - СТС - end
'SIT_HeadOfInitiatorsUnit = """#Head of initiator`s department"";"'-----------------------old
'SIT_DirectorOfInitiatorsDepartment = """#Director of initiator`s department"";"'-----------------------old
'SIT_DirectorOfInitiatorsDivision = """#Director of initiator`s division"";"'-----------------------old
'SIT_VicePresidentOfInitiator = """#Vice-President of initiator"";"'-----------------------old
SIT_President = """#President"";"
SIT_FinancialVicePresident = """#Vice-President, Finances and Investment"";"
SIT_HRDirector = """#HR Director"";"
SIT_GenDirector = """#General Director"";"

SIT_LinkInstruction = "Instruction"
SIT_LinkInstructionHint = "User manual for Sitronics and STS users"
SIT_LinkInstructionFile = "Manuals/UserInstruction.doc" 'Имя файла с инструкцией на данном языке

'SIT_Initiator = "Initiator"
SIT_Initiator = "Drafter"
SIT_CurrencyRateToUSD = "Currency rate"
SIT_Orders_Prikaz = "OR - Order"
SIT_Orders_Rasporyajenie = "Р - Task"
'vnik_rasp_norm_doc
SIT_Orders_Prikaz_ND = "OR - Order approving regulations" 'при изменении данного наименования поменять условие запроса в файле AgentUser.asp Function NormatDocsAutoReconcilation()
SIT_NORM_DOCS_CLOSED = "Regulations category is closed for new document creation as of 01.07.2010, use Orders category and „Order to approve a regulation“ document type"
SIT_NORM_DOCS_WARNING = "Deadline for approval must be more or equal to "
'vnik_rasp_norm_doc
SIT_BusinessUnit = "Business unit"
SIT_CostCenterCode = "Cost Center Code"
SIT_Project = "Project"
SIT_ExpenseItem = "Expense item"
SIT_ProjectCode = "Code of project"
SIT_PaymentType = "Calculation form"
SIT_Budgeted = "Budgeted (Yes\No)"
SIT_USDAmount = "USD Amount"
SIT_PurchaseOrder = "Purchase Order"
SIT_PaymentOrder = "Payment Order"
SIT_Subject = "Subject"
'20090622 - Заявка ТКП
SIT_ComOfferNumber = "Commercial offer number"

SIT_RolesDirSitronics = SIT_RolesDirSitronics_EN
SIT_RolesDirSTS = SIT_RolesDirSTS_EN
'Запрос №1 - СИБ - start
SIT_RolesDirSIB = SIT_RolesDirSIB_EN
'Запрос №1 - СИБ - end

SIT_Letter_Initiative = "initiative"
SIT_Letter_AnswerFor = "in response to letter # "
SIT_SistemaTasks = "SISTEMA tasks"
SIT_IncomingMailTasks = "Tasks according to incoming mail"
SIT_TasksOnOrders = "Tasks according to orders"
SIT_TasksOnMemos = "Tasks of Memos"
SIT_TasksOnOtherDocs = "Tasks (based on other documents or verbal)"
SIT_YesNo = "Yes,No"

SIT_ErrorInDepartmentLeader = "Incorrect department manager's name "
SIT_DocChangedReconcilationRequired = "The document has been changed, agreement needed"
SIT_Language = "Language"
SIT_AssistantOfPresident = """#President's Assistant "";"
SIT_HeadOfDocControl = """#Office manager"";"
SIT_SortType = "Sorting Type"
SIT_SortTypeAscending = "ascending"
SIT_SortTypeDescending = "descending"

SIT_MoreThanOneLeader1 = "There is more than one manager in the organizational unit "
SIT_MoreThanOneLeader2 = ""
SIT_NoAccessToParentDoc = "You have no access to the parent document, you cannot refer to it"

'rmanyushin 60298, 89142 01.04.2010 Start
STS_ContractType = "Contracting party"
'vnik_purchase_order
SIT_ContractType = "How to purchase"
'vnik_purchase_order
'ph - 20120216 - start
STS_OffshoreZone = "The Contractor’s location in the offshore area"
'ph - 20120216 - end
'ph - 20101225 - start
'Делаем только на англ. будет в справочнике Ролей для PO
'STS_HRAdministrationManager = """#HR Administration Manager"";"
'ph - 20101225 - end
STS_DptOSSBSS = """#Deputy CEO for OSS/BSS"";" 
 
'STS_ContractPartyDirGUID = "{73AA9A7D-FAC2-4B75-8C42-0C8B7453DC9B}"' Справочник на IT-TEST-07
'STS_ContractPartyDirGUID = "{913547B9-1703-4241-8E5C-CE87656582DF}" ' Справочник на 8888
STS_ContractPartyDirGUID = "{260405A7-E08E-4012-B353-9447FAB64683}" ' Справочник на GL-PAYDOX-01 
'rmanyushin 60298, 89142 01.04.2010 End

'rmanyushin 93755, 22.04.2010 Start
STS_SLUZH_ZAPISKA_OVERTIME_PRIKAZ_TEXT = "In the context of …"+VbCrLf+VbCrLf+"I hereby request that the following SITRONICS Telecom Solutions employees are ordered to work on the weekend based on their written consent. They will be compensated in compliance with the Labour Code: …"
STS_SLUZH_ZAPISKA_OVERTIME_TITLE = "Ordering work on weekends"
'rmanyushin 93755, 22.04.2010 End

'rmanyushin 119579 19.08.2010 Start
STS_SLUZH_ZAPISKA_HOLIDAY_TITLE = "Holiday Request"
'rmanyushin 119579 19.08.2010 End

'rmanyushin 119579 19.08.2010 Start
'Запрос №34 - СТС - start
'STS_AssistantDirector = """#Deputy CEO"";"
'STS_DirectorOfDirection = """#Line of Business Director"";"
'STS_HeadOfInitiatorsGroup = """#Manager of initiator's group"";"
STS_AssistantDirector = STS_AssistantDirector_EN
STS_DirectorOfDirection = STS_DirectorOfDirection_EN
STS_HeadOfInitiatorsGroup = STS_HeadOfInitiatorsGroup_EN
'Запрос №34 - СТС - end
'rmanyushin 119579 19.08.2010 Stop

'Запрос №11 - СТС - start
STS_ContractPaymentDirection_In = STS_ContractPaymentDirection_In_EN
STS_ContractPaymentDirection_Out = STS_ContractPaymentDirection_Out_EN
STS_ContractPaymentDirection_Free = STS_ContractPaymentDirection_Free_EN
'Запрос №11 - СТС - end
'{ph - 20120517
STS_ErrorInHolidayFirstDate = "Specify the right date of the beginning of your vacations"
'ph - 20120517}

  Case "3" 'CZ ------------------------------------------------------- ЧЕШСКИЙ

SIT_ErrorUnrecognizedRoles = "CHYBA! Některé role v procesu zpracování dokumentu nejsou rozpoznány."
SIT_NotificationDocReconciled1 = "Dokument "
SIT_NotificationDocReconciled2 = " byl uživatelem schválen"
SIT_NotificationDocReconciled3 = " s KOMENTÁŘE."
SIT_NotificationLetterForYou = "Obdržel/-a jste dopis z externí organizace"
SIT_ErrorSumExceeding = "<font color = red>Částka platby je vyšší než cena nákupu</font>"
SIT_NotificationApprovalChanged1 = "Podepisovatel byl změněn "
SIT_NotificationApprovalChanged2 = " --> "
SIT_ErrorInDepartmentCode1 = "U organizační jednotky uživatele: "
SIT_ErrorInDepartmentCode2 = " chybí kód (nebo je uvedeno chybné oddělení). Kontaktujte správce systému!"
SIT_ErrorInDateCompletion = "Chybný termín dokončení: "
SIT_ErrorInProjectNumber = "Chybné číslo projektu"
SIT_ErrorInBU = "Obchodní jednotka neodpovídá katalogu"
SIT_ErrorInChartOfAccount = "Nákladová položka neodpovídá katalogu"
SIT_ErrorInPaymentType = "Způsob platby neodpovídá katalogu"
SIT_ErrorInCostCenter = "Nákladové středisko neodpovídá katalogu"
SIT_NoHeaderInResponsibleDepartment = "Vedoucí odpovědné jednotky nenalezen"
SIT_ErrorInSumOrCurrency = "<font color = red>POZOR!</font>Chybná částka nebo měna platby "
SIT_ErrorInProjectManager = "POZOR! Vedoucí projektu nemá uživatelské jméno"
SIT_ErrorNotAllUsersFound = "POZOR! Někteří účastníci schvalovacího procesu nebyli nalezeni. Kontaktujte správce systému."
SIT_ParentOrderCanceled = "Nadřazená nákupní objednávka byla zrušena"
SIT_ParentOrderNotApproved = "Nadřazená nákupní objednávka nebyla schválena"
SIT_ErrorInParentOrderNumber = "Chybné číslo nadřazené objednávky"
SIT_AdditionalAgrees = "Další účastníci schvalovacího procesu: "
SIT_RequiredAgrees = "Povinní účastníci schvalovacího procesu: "
SIT_PreliminaryAgrees = "Předběžná koordinace: "
SIT_BUT_AFKTasks = "ÚKOLY OD AFK"
SIT_BUT_AFKTasksHint = SIT_BUT_AFKTasks
SIT_CentralBut_CoResponsibleHint = "Nesplněné úkoly, u nichž jsem spoluvykonavatelem"
SIT_CentralBut_PurchaseOrdersHint = "Neuzavřené nákupní objednávky"
SIT_LastReconcilationDate = "Termín schvalování: "
SIT_NotificationApproveRefused = "Neschváleno"
SIT_NotificationAgreeRefusedByUser = "Neschváleno uživatelem "
SIT_ButtonLoadProjects = "Načíst projekty"
SIT_ButtonLoadProjectsHint = "Načíst katalog projektů z XML?"
SIT_ButtonLoadProjectsConfirm = "Načíst katalog projektů z XML souboru?"
SIT_ErrorInUserField1 = "Zadana hodnota v polozce "
SIT_ErrorInUserField2 = " nebyla vybrana se seznamu"
'Запрос №11 - СТС - start
SIT_ErrorNoPartnerCode = "Chybí kód smluvního partnera. Vyberte příslušný abecední kód ze seznamu smluvních partnerů."
SIT_CannotCreateOldContracts = "V této kategorii není povoleno vytvářet dokumenty, použijte kategorii Smlouvy"
'Запрос №11 - СТС - end
'Запрос №43 - СТС - start
STS_ErrorCantCreatePOLargerLimit1 = "Nejste oprávněn (a) vytvářet Nákupní Objednávky (PO) nad  $10000"
STS_ErrorCantCreatePOLargerLimit2 = "Nejste oprávněn (a) vytvářet Nákupní Objednávky (PO) nad  $50000"
STS_ErrorPartnerIsNotChecked = "Nelze vytvářet Nákupní Objednávky (PO) s neověřenými smluvními stranami"
'Запрос №43 - СТС - end
'Запрос №46 - СТС - start
SIT_CannotCreateDocInThisCategory = "V této kategorii je zakázáno vytvářet dokumenty"
SIT_ErrorInProjectNumbers = "Chyba v uvedení čísel projektů: "
SIT_ParentDocCancelled = "Rodičovský dokument je zrušen"
SIT_ParentDocNotApproved = "Rodičovský dokument není schválen"
SIT_ErrorInParentDocNumber = "Chyba v uvedení čísla rodičovského dokumentu"
STS_SLUZH_ZAPISKA_OVERTIME_PLAN_TITLE = "Žádost o schválení práce přesčas"
STS_SLUZH_ZAPISKA_OVERTIME_PLAN_TEXT = "V souvislosti s provedením prací na projektu … žádám o schválení práce přesčas během víkendu/svátku s písemným souhlasem níže uvedených zaměstnanců společnosti SITRONICS Telecom Solutions s nárokem na mzdu podle Zákoníku práce:"
'Запрос №46 - СТС - end
'20100111 - Запрос №13 из СТС - start
STS_YouAreResponsibleSubject = "Byl jste určen jako osoba odpovědná za realizaci dokumentu"
STS_YouAreResponsibleBody = "Tímto Vás informujeme, že jste byl určen jako osoba odpovědná za realizaci tohoto dokumentu. S případnými dotazy se obraťte na iniciátora tohoto dokumentu nebo administrátora systému."
'20100111 - Запрос №13 из СТС - end

STS_HeadOfSector = """#Vedoucí sektoru"";"
STS_HeadOfDepartment = """#Ředitel oddělení"";"
STS_HeadOfDivision = """#Ředitel divize"";"
STS_FinancialControl = """#Finanční kontrolor"";"
STS_FinDirector = """#Finanční ředitel"";"
STS_GenDirector = """#Generální ředitel"";"
'Запрос №34 - СТС - start
'STS_ProjectManager = """#Vlastník rozpočtu projektu"";"
STS_ProjectManager = STS_ProjectManager_CZ
STS_Overtime_Requester = STS_Overtime_Requester_CZ
STS_DirectorOfProjectManagersDepartment = STS_DirectorOfProjectManagersDepartment_CZ
STS_DirectorOfOvertimeRequestersDepartment = STS_DirectorOfOvertimeRequestersDepartment_CZ
STS_DirectorOfOvertimeRequestersDirection = STS_DirectorOfOvertimeRequestersDirection_CZ
STS_OvertimeRequestersAssistantDirector = STS_OvertimeRequestersAssistantDirector_CZ
'Запрос №34 - СТС - end
'Запрос №36 - СТС - start
STS_CostCenterDirectorOfDepartment = STS_CostCenterDirectorOfDepartment_CZ
STS_CostCenterDirectorOfDivision = STS_CostCenterDirectorOfDivision_CZ
STS_DirectorOfProjectManagersDivision = STS_DirectorOfProjectManagersDivision_CZ
'Запрос №36 - СТС - end
'Запрос №46 - СТС - start
STS_OvertimeFuncLeader = STS_OvertimeFuncLeader_CZ
STS_Initiator = STS_Initiator_CZ
'Запрос №46 - СТС - end

STS_Accounting = """#Účtárna"";"
STS_Treasury = """#Pokladna"";"
STS_SecrPravlenia = """#Asistent/-ka vedení STS"";"
SIT_SecrPravlenia = """#Asistent/-ka vedení společnosti Sitronics"";"
SIT_Registrar = """#Registrátor"";"

SIT_ReportControl = "Kontrola"
SIT_ReportControlValues = "Kontrolován, Nekontrolován"
SIT_ReportStartingDate = "Datum začátku"
SIT_ReportFinishingDate = "Datum konce"
SIT_ReportTaskStartingDateFrom = "Datum zadání úkolu (od)"
SIT_ReportTaskStartingDateTo = "Datum zadání úkolu (do)"
SIT_ReportTypeOfRequest = "Typ požadavku"
SIT_ReportRequestTypes = "Splněno do termínu,Splněno bez dodržení termínu,Nesplněno (před termínem),Nesplněno (termín nebyl dodržen)"


SIT_ErrorInFieldValue1 = "<font color=red>Chyba!</font> Hodnoty pro toto pole ["
SIT_ErrorInFieldValue2 = "] musí být vybrány ze seznamu"

SIT_ReportTypeOfTask = "Druh úkolu"
SIT_TaskName = "Jméno"

'Запрос №11 - СТС - start
SIT_HeadOfInitiatorsUnit = SIT_HeadOfInitiatorsUnit_CZ
SIT_DirectorOfInitiatorsDepartment = SIT_DirectorOfInitiatorsDepartment_CZ
SIT_DirectorOfInitiatorsDivision = SIT_DirectorOfInitiatorsDivision_CZ 'В СТС
SIT_VicePresidentOfInitiator = SIT_VicePresidentOfInitiator_CZ 'В УК
'SIT_HeadOfInitiatorsUnit = """#Vedoucí sektoru iniciátora"";"
'SIT_DirectorOfInitiatorsDepartment = """#Vedoucí oddělení iniciátora"";"
'SIT_DirectorOfInitiatorsDivision = """#Ředitel divize iniciátora"";"
'SIT_VicePresidentOfInitiator = """#Viceprezident iniciátora"";"
'Запрос №11 - СТС - end
'SIT_HeadOfInitiatorsUnit = """#Vedouci iniciatorova useku"";" '--------------------old
'SIT_DirectorOfInitiatorsDepartment = """#Vedouci iniciatorova useku"";" '--------------------old
'SIT_DirectorOfInitiatorsDivision = """#Reditel iniciatorovy divize"";" '--------------------old
'SIT_VicePresidentOfInitiator = """#Iniciatoruv Vice-President"";" '--------------------old
SIT_President = """#President"";"
SIT_FinancialVicePresident = """#Vice-President, Finance a Investice"";"
SIT_HRDirector = """#HR ředitel"";"
SIT_GenDirector = """#Generální ředitel"";"

SIT_LinkInstruction = "Instrukce"
SIT_LinkInstructionHint = "Uživatelská příručka pro Sitronics a STS"
SIT_LinkInstructionFile = "Manuals/UserInstruction.doc" 'Имя файла с инструкцией на данном языке

SIT_Initiator = "Iniciátor"
SIT_CurrencyRateToUSD = "Kurz měny"
'SIT_Orders_Prikaz = "OR - Objednávka" '--------------------old
SIT_Orders_Prikaz = "OR - Nařízení"
SIT_Orders_Rasporyajenie = "Р - Úkol"
'vnik_rasp_norm_doc
SIT_Orders_Prikaz_ND = "OR - Schvalování objednávek Řídicí dokumenty" 'при изменении данного наименования поменять условие запроса в файле AgentUser.asp Function NormatDocsAutoReconcilation()
SIT_NORM_DOCS_CLOSED = "Kategorie Řídicí dokumenty je od 01.07.2010 uzavřena pro vytváření nových dokumentů, použijte kategorii Nařízení a typ dokumentu „Nařízení schválení normativního dokumentu“"
SIT_NORM_DOCS_WARNING = "Termín schválení musí být větší nebo roven "
'vnik_rasp_norm_doc
SIT_BusinessUnit = "Obchodní jednotka"
SIT_CostCenterCode = "Kód nákladového střediska"
SIT_Project = "Projekt"
SIT_ExpenseItem = "Nákladová položka"
SIT_ProjectCode = "Kód projektu"
SIT_PaymentType = "Forma platby"
SIT_Budgeted = "Je v rozpočtu (Ano\Ne)"
SIT_USDAmount = "Částka dokumentu USD"
SIT_PurchaseOrder = "Nákupní objednávka"
SIT_PaymentOrder = "Požadavek na platbu"
SIT_Subject = "Předmět"
'20090622 - Заявка ТКП
SIT_ComOfferNumber = "Číslo obchodní nabídky"

SIT_RolesDirSitronics = SIT_RolesDirSitronics_CZ
SIT_RolesDirSTS = SIT_RolesDirSTS_CZ
'Запрос №1 - СИБ - start
SIT_RolesDirSIB = SIT_RolesDirSIB_CZ
'Запрос №1 - СИБ - end

SIT_Letter_Initiative = "Iniciativní"
SIT_Letter_AnswerFor = "V odpovědi na dopis č. "
SIT_SistemaTasks = "Úkoly SISTEMA"
SIT_IncomingMailTasks = "Úkoly dle příchozí pošty"
SIT_TasksOnOrders = "Úkoly dle objednávek"
SIT_TasksOnMemos = "úkoly sdělení"
SIT_TasksOnOtherDocs = "Úkoly (dle dalších dokumentů či ústní)"
SIT_YesNo = "Ano,Ne"

SIT_ErrorInDepartmentLeader = "Chybné jméno vedoucího oddělení "
SIT_DocChangedReconcilationRequired = "Dokument byl změněn, je nutné jeho odsouhlasení"
SIT_Language = "Jazyk"
SIT_AssistantOfPresident = """#Asistent prezidenta"";"
SIT_HeadOfDocControl = """#Vedoucí kanceláře"";"
SIT_SortType = "Typ řazení"
SIT_SortTypeAscending = "vzestupně"
SIT_SortTypeDescending = "sestupně"

SIT_MoreThanOneLeader1 = "V této organizační jednotce je "
SIT_MoreThanOneLeader2 = " více než jeden vedoucí pracovník."
SIT_NoAccessToParentDoc = "Nemáte přístup k rodičovskému dokumentu, nemůžete na něj odkazovat"

'rmanyushin 60298, 89142 01.04.2010 Start
STS_ContractType = "Smluvní strana"
'vnik_purchase_order
SIT_ContractType = "Jak nakupovat"
'vnik_purchase_order
'ph - 20120216 - start
STS_OffshoreZone = "Sídlo smluvní strany v zahraničí"
'ph - 20120216 - end
'ph - 20101225 - start
'Делаем только на англ. будет в справочнике Ролей для PO
'STS_HRAdministrationManager = """#Vedoucí administrativního sektoru Lidské zdroje"";"
'ph - 20101225 - end
STS_DptOSSBSS = """#Zástupce generálního ředitele pro OSS/BSS"";"
'STS_ContractPartyDirGUID = "{4D630A10-9532-4F8E-9155-FE22104C012C}"' Справочник на IT-TEST-07
'STS_ContractPartyDirGUID = "{C2C0CC6F-6C48-41DC-9D91-C8BF2F676E5D} " ' Справочник на 8888
STS_ContractPartyDirGUID = "{EB4C8816-901B-405C-867E-61420B4E0C30}" ' Справочник на GL-PAYDOX-01 
'rmanyushin 60298, 89142 01.04.2010 End

'rmanyushin 93755, 22.04.2010 Start
STS_SLUZH_ZAPISKA_OVERTIME_PRIKAZ_TEXT = "V souvislosti s …"+VbCrLf+VbCrLf+"žádám, aby byla níže uvedeným zaměstnancům SITRONICS Telecom Solutions na základě písemného souhlasu nařízena práce o víkendu s odměnou podle Zákoníku práce: …"
STS_SLUZH_ZAPISKA_OVERTIME_TITLE = "O nařízení práce o víkendu"
'rmanyushin 93755, 22.04.2010 End

'rmanyushin 119579 19.08.2010 Start
STS_SLUZH_ZAPISKA_HOLIDAY_TITLE = "Žádost o dovolenou"
'rmanyushin 119579 19.08.2010 End

'rmanyushin 119579 19.08.2010 Start
'Запрос №34 - СТС - start
'STS_AssistantDirector = """#Zástupce generálního ředitele"";"
'STS_DirectorOfDirection = """#Ředitel divize"";"
'STS_HeadOfInitiatorsGroup = """#Team Leader iniciátor"";"
STS_AssistantDirector = STS_AssistantDirector_CZ
STS_DirectorOfDirection = STS_DirectorOfDirection_CZ
STS_HeadOfInitiatorsGroup = STS_HeadOfInitiatorsGroup_CZ
'Запрос №34 - СТС - end
'rmanyushin 119579 19.08.2010 Stop



'Запрос №11 - СТС - start
STS_ContractPaymentDirection_In = STS_ContractPaymentDirection_In_CZ
STS_ContractPaymentDirection_Out = STS_ContractPaymentDirection_Out_CZ
STS_ContractPaymentDirection_Free = STS_ContractPaymentDirection_Free_CZ
'Запрос №11 - СТС - end
'{ph - 20120517
STS_ErrorInHolidayFirstDate = "Uveďte správné datum začátku dovolené"
'ph - 20120517}

End Select

'Запрос №37 - СТС - start - определяем переменные за пределеми Sub, чтобы они не оказались локальными
STS_TaskCompletionExpires = ""
STS_TaskCompletionExpired1 = ""
STS_TaskCompletionExpired4 = ""
STS_TaskCompletionExpiresControl = ""
STS_ReconciliationExpires = ""
STS_ReconciliationExpired1 = ""
STS_ReconciliationExpired4 = ""
STS_PenaltyMessage = ""
'Запрос №37 - СТС - end

Sub SetLangConstsForEmail(ByVal Lang)
'Запрос №37 - СТС - start
  If VAR_CurrentL = Lang and oPayDox.VAR_CurrentL = Lang Then
    Exit Sub
  End If
'Запрос №37 - СТС - end
  VAR_CurrentL = Lang
  oPayDox.VAR_CurrentL=VAR_CurrentL
  'Переприсваивание стандартных констант, используемых в почтовых сообщениях
  Select case UCase(Lang)
  case "RU" '--------------------------------------------------------------------------- RU
    DOCS_UnderDevelopment = "На стадии разработки"
    DOCS_Mark = "Сделать отметку"
    DOCS_Viewed = "Ознакомлен(а) с документом"
    DOCS_View = "Ознакомлен"
    DOCS_MakeCompleted = "Назначить статус «Исполнено»"
    But_MakeCompleted = "Статус «Исполнено»"
    DOCS_CreateComment = "Создать комментарий"
    But_Comment = "Комментарий"
    DOCS_AGREE = "Согласовать"
    But_AGREE = "Согласовать"
    DOCS_Refuse = "Отказать в согласовании"
    But_Refuse = "Отказать"
    DOCS_RequestCompleted = "Запросить назначение статуса «Исполнено»"
    But_RequestCompleted = "Запрос «Исполнено»"
    DOCS_NameAproval = "Утверждающий"
    DOCS_NameCreation = "Создатель документа"
    DOCS_NameControl = "Контролер"
    DOCS_NameResponsible = "Исполнитель"
    DOCS_AmountDoc = "Сумма по документу"
    DOCS_Signing = "На согласовании"
    DOCS_Approving = "На утверждении"
    DOCS_Approved = "Утверждено"
    DOCS_RefusedApp = "Отказано в утверждении"
    DOCS_Completed = "Исполнено"
    DOCS_Actual = "Исполняется"
    DOCS_Cancelled = "Отменено"
    But_Statuses = "Состояние"
    But_StatusesDesc = "Текущее состояние документа"
    But_Actions = "Действия"
    But_ActionsDesc = "Действия над документом"
    DOCS_Inactive = "Неактивен"
    DOCS_Active = "Активен"
    DOCS_UNAPPROVED = "Документы, требующие Вашего согласования"
    BUT_VISA = "СОГЛАСОВАНИЕ"
    DOCS_YouAreResponsible = "Документы, для которых данный пользователь является ответственным исполнителем"
    BUT_RESPONSIBLE = "ОТВЕТСТВЕННЫЙ"
    DOCS_ListToReconcile = "Список согласующих"
    DOCS_NextStepToReconcile = "Последовательное согласование - следующий уровень"
    DOCS_Reconciled = "Согласовано"
    DOCS_Refused = "Отказано в согласовании"
    DOCS_VersionFile = "Файл документа"
    DOCS_GoDoc = "Перейти к карточке документа "
    DOCS_Go = "Перейти"
    BUT_URGENT = "СРОЧНЫЕ"
    DOCS_EXPIRED = "С наступающим сроком и просроченные"
    BUT_CREATED = "СОЗДАННЫЕ"
    DOCS_YouAreCreator = "Неисполненные документы, которые Вы создали"
    BUT_COMPLETION1 = "ИСПОЛНЕНИЕ"
    DOCS_NOTCOMPLETED = "Требующие Вашей отметки об исполнении (Вы являетесь утверждающим или контролером)"
    BUT_CONTROL = "КОНТРОЛЬ"
    DOCS_UnderControl = "Документы на КОНТРОЛЕ"
    BUT_NEWDOCS = "НОВЫЕ"
    DOCS_NewDocs = "Новые документы"
    DOCS_InOffice = "Нахожусь сейчас в офисе"
    DOCS_OutOfOffice = "Нахожусь сейчас вне офиса"
    DOCS_PushToChange = "Нажмите, чтобы изменить "
    BUT_OUTOFFICE = "Вне офиса"
    BUT_INOFFICE = "В офисе"
    DOCS_PaymentOutgoingIncompleted = "Документы с неисполненными исходящими платежами"
    DOCS_StatusRequireToBePaid = " - требующие оплаты"
    BUT_Payments = "ПЛАТЕЖИ"
    DOCS_PaymentIncomingIncompleted = "Документы с неисполненными входящими платежами"
    DOCS_GetListDocEMail = "Получить по e-mail список документов"
    DOCS_SendAction = "Послать по e-mail указание в PayDox выполнить действие"
    DOCS_eMailClientWarning = "PayDox E-Mail Клиент - Взаимодействие с PayDox через e-mail. Пожалуйста, не изменяйте содержание этого e-mail!"
    DOCS_NotPreviousYet = "Предыдущий шаг согласования еще не закончен"
    DOCS_StatusPaymentPaid = "Оплачен"
    DOCS_StatusPaymentNotPaid = "Неоплачен"
    DOCS_StatusPaymentToBePaid = "Неоплачен, требует оплаты"
    DOCS_StatusPaymentSentToBePaid = "Неоплачен, выставлен к оплате"
    DOCS_StatusPaymentToPay = "Неоплачен. Дано указание оплатить"
    DOCS_StatusExistsButNotDefined = "Статус не задан"
    DOCS_Comments = "Комментарии"
    DOCS_Contacts = "Контакты"
    DOCS_NotificationDoc = "Уведомление о документе"
    DOCS_ListToView = "Список ознакомления с  д-том"
    DOCS_UploadFile = "Загрузить файл в систему"
    But_SendFile = "Послать файл"
    DOC_SendFile = "Присоедините файл к e-mail. При необходимости укажите Ваш комментарий после ключевого слова COMMENT ="
    DOCS_FilesReceived = "Файлов принято"
    DOCS_FileNotUploaded = "Файл документа НЕ загружен - пожалуйста, используйте только стандартные английские символы для имени файла. Не используйте пробелы в имени файла."
    DOCS_FileReceived = "Файл принят по e-mail"
    DOCS_GetDocEMail = "Запрос на получение этого документа по e-mail (включая присоединенные файлы)"
    But_GetDoc = "Получить документ"
    USER_NOSMS = "Отправка SMS не задана"
    DOCS_SMSError = "Ошибка связи с GSM-модемом"
    DOCS_GSMModemBusy = "GSM-модем занят. Сообщение не отправлено"
    DOCS_WRONGPHONECELL = "Неверный номер сотового телефона. используйте полный номер с символом «+» в начале номера"
    DOCS_GSMModemError1 = "GSM-модем не подключен или неверный номер порта: "
    DOCS_SMSSent = "SMS-сообщение отправлено"
    DOCS_StatusPaymentToBePaidPart = "Оплачен частично, требует дальнейшей оплаты"
    DOCS_StatusPaymentToPayPart = "Неоплачен. Дано указание оплатить ЧАСТИЧНО"
    DOCS_StatusPaymentToPayPartRest = "Дано указание оплатить ОСТАТОК"
    DOCS_StatusPaymentPaidPart = "Оплачен ЧАСТИЧНО"
    DOCS_StatusPaymentToPayPartPart = "Дано указание оплатить еще ЧАСТЬ"
    DOCS_ViewedStatusDocs = "Документы, требующие ознакомления"
    BUT_VIEWEDSTATUSDOCS = "ОЗНАКОМИТЬСЯ"
    DOCS_PaymentOutgoing = "Исходящий платеж"
    DOCS_PaymentIncoming = "Входящий платеж"
    DOCS_NotPaymentDoc = "Платеж не задан"
    DOCS_APPROVE = "Утвердить"
    But_Approve = "Утвердить"
    DOCS_NotAgreedYet = "Согласование документа еще не завершено"
    DOCS_RefuseApp = "Отказать в утверждении"
    But_RefuseApp = "Отказать"
    DOCS_BusinessProcessSteps = "Этапы бизнес-процесса"
    DOCS_StatusCompletion = "Статус исполнения"
    DOCS_Currency = "Код валюты"
    DOCS_ToView = "Ознакомиться"
    DOCS_ParallelReconciliation = "Параллельное согласование"
    DOCS_Messages = "Сообщения"
    But_DEPUTY = "Заместитель"
    DOCS_VersionFiles = "Файлы документа"
    DOCS_LocationPaper = "Расположение подлинника документа"
    DOCS_Notifications = "Уведомления"
    DOCS_Resolutions = "Резолюции"
    DOCS_APROVAL = "Утверждение"
    DOCS_COMPLETION = "Исполнение"
    DOCS_History = "Ход исполнения"
    DOCS_StatusPayment = "Статус платежа"
    DOCS_Visa = "Согласование"
    DOCS_MEETINGS = "Совещания, встречи"
    DOCS_Viewed1 = "Ознакомлены с документом"
    DOCS_SystemMessage = "Системные комментарии"
    DOCS_Resources = "Ресурсы"
    DOCS_PARTNERS = "Контрагенты"
    DOCS_Reviews = "Рецензии"
    DOCS_Review = "Рецензия"
    DOCS_ToReview1 = "Запрос на рецензию"
    DOCS_ReconciliationREQUIRED = "Требует согласования"
    DOCS_Department = "Подразделение"
    DOCS_VersionMainAdd = "Дополнение к основной версии"
    DOCS_NOTFOUND = "Информация не найдена"
    DOCS_ACT = "Вид деятельности"
    DOCS_All = "Все"
    DOCS_Author = "Автор"
    DOCS_Correspondent = "Адресаты, список рассылки"
    DOCS_DateActivation = "Дата регистрации"
    DOCS_DateCompletion = "Дата исполнения"
    DOCS_Description = "Краткое содержание"
    DOCS_DocID = "Индекс документа"
    DOCS_ErrorSMTP = "ОШИБКА при вызове SMTP - проверьте, установлена ли стандартная компонента Windows Default SMTP Virtual Server для IIS"
    DOCS_EXPIREDSEC = "Полномочия доступа истекли"
    DOCS_FROM1 = "из"
    DOCS_Name = "Наименование документа"
    DOCS_NoAccess = "Доступ к данному документу не разрешен"
    DOCS_NotificationNotCompletedDoc = "Уведомление о неисполненном документе"
    DOCS_NotificationSentTo = "Уведомление отправлено пользователям"
    DOCS_PartnerName = "Контрагент"
    DOCS_Reconciliation = "Согласование"
    DOCS_Resolution = "Резолюция"
    DOCS_Sender = "Отправитель"
    DOCS_SendNotification = "Отправить e-mail уведомление"
    DOCS_STATUSHOLD = "Доступ к системе приостановлен"
    DOCS_UsersNotFound = "Пользователи не найдены"
    USER_NOEMail = "Отсутствует или ошибочный e-mail адрес"
  case "" '----------------------------------------------------------------------------- EN
    DOCS_UnderDevelopment = "Under Development"
    DOCS_Mark = "Mark as"
    DOCS_Viewed = "Document viewed"
    DOCS_View = "Viewed"
    DOCS_MakeCompleted = "Change status to «Completed»"
    But_MakeCompleted = "Status «Completed»"
    DOCS_CreateComment = "Create comment"
    But_Comment = "Comment"
    DOCS_AGREE = "Agree"
    But_AGREE = "Agree"
    DOCS_Refuse = "Refuse"
    But_Refuse = "Refuse"
    DOCS_RequestCompleted = "Request for status «Completed»"
    But_RequestCompleted = "Request «Completed»"
    DOCS_NameAproval = "Approval person"
    DOCS_NameCreation = "Created by"
    DOCS_NameControl = "Controlling person"
    DOCS_NameResponsible = "Responsible person"
    DOCS_AmountDoc = "Document amount"
    DOCS_Signing = "Signing"
    DOCS_Approving = "Approving"
    DOCS_Approved = "Approved"
    DOCS_RefusedApp = "Approval refused"
    DOCS_Completed = "Completed"
    DOCS_Actual = "Actual"
    DOCS_Cancelled = "Cancelled"
    But_Statuses = "Statuses"
    But_StatusesDesc = "Current document statuses"
    But_Actions = "Actions"
    But_ActionsDesc = "Actions under the document"
    DOCS_Inactive = "Inactive"
    DOCS_Active = "Active"
    DOCS_UNAPPROVED = "Required to be agreed by you"
    BUT_VISA = "AGREE"
    DOCS_YouAreResponsible = "Documents where the current user is the responsible person"
    BUT_RESPONSIBLE = "RESPONSIBLE"
    DOCS_ListToReconcile = "Agree list (participants)"
    DOCS_NextStepToReconcile = "Sequential Agree process - Next step"
    DOCS_Reconciled = "Agreed"
    DOCS_Refused = "Refused"
    DOCS_VersionFile = "Document file"
    DOCS_GoDoc = "Go to the document record "
    DOCS_Go = "Go"
    BUT_URGENT = "URGENT"
    DOCS_EXPIRED = "Having near expiration date or expired"
    BUT_CREATED = "CREATED"
    DOCS_YouAreCreator = "Actual documents you have created"
    BUT_COMPLETION1 = "COMPLETION"
    DOCS_NOTCOMPLETED = "Not completed yet (you are approving or controlling person)"
    BUT_CONTROL = "CONTROL"
    DOCS_UnderControl = "Documents under CONTROL"
    BUT_NEWDOCS = "NEW DOCS"
    DOCS_NewDocs = "New documents"
    DOCS_InOffice = "I am in the office NOW"
    DOCS_OutOfOffice = "I am out of office NOW"
    DOCS_PushToChange = "Push to change"
    BUT_OUTOFFICE = "Out of office"
    BUT_INOFFICE = "In office"
    DOCS_PaymentOutgoingIncompleted = "Documents having outgoing incomplete payments"
    DOCS_StatusRequireToBePaid = " - require to be paid"
    BUT_Payments = "PAYMENTS"
    DOCS_PaymentIncomingIncompleted = "Documents having incoming incomplete payments"
    DOCS_GetListDocEMail = "Receive list of documents by e-mail"
    DOCS_SendAction = "Send this action to PayDox by e-mail"
    DOCS_eMailClientWarning = "PayDox E-Mail Client - Interaction to PayDox via e-mail. Please DO NOT ALTER this e-mail content!"
    DOCS_NotPreviousYet = "Previous agree step not completed yet"
    DOCS_StatusPaymentPaid = "Paid"
    DOCS_StatusPaymentNotPaid = "Unpaid"
    DOCS_StatusPaymentToBePaid = "Unpaid, requires to be paid"
    DOCS_StatusPaymentSentToBePaid = "Unpaid, sent to be paid"
    DOCS_StatusPaymentToPay = "Unpaid. To pay"
    DOCS_StatusExistsButNotDefined = "Status not defined"
    DOCS_Comments = "Comments"
    DOCS_Contacts = "Contact Us"
    DOCS_NotificationDoc = "Document notification"
    DOCS_ListToView = "Document viewers list"
    DOCS_UploadFile = "Upload file into the system"
    But_SendFile = "Send file"
    DOC_SendFile = "Attach file to e-mail. You can provide your comment after the key word COMMENT ="
    DOCS_FilesReceived = "Files received"
    DOCS_FileNotUploaded = "Document file NOT uploaded - please use only standard English symbols for the file name. Do not use spaces in the file name."
    DOCS_FileReceived = "File received by e-mail"
    DOCS_GetDocEMail = "Request to receive this document by e-mail (including attachments)"
    But_GetDoc = "Receive document"
    USER_NOSMS = "SMS-notification not set"
    DOCS_SMSError = "GSM-modem connection error"
    DOCS_GSMModemBusy = "GSM-modem busy. SMS-message not sent"
    DOCS_WRONGPHONECELL = "Wrong cellphone number. Use full number having symbol «+» in the beginning"
    DOCS_GSMModemError1 = "GSM-modem not connected or invalid port number: "
    DOCS_SMSSent = "SMS-message sent"
    DOCS_StatusPaymentToBePaidPart = "Paid in part, requires to be paid more"
    DOCS_StatusPaymentToPayPart = "Unpaid. To pay in part"
    DOCS_StatusPaymentToPayPartRest = "To pay the REST"
    DOCS_StatusPaymentPaidPart = "Paid in part"
    DOCS_StatusPaymentToPayPartPart = "To pay one PART more"
    DOCS_ViewedStatusDocs = "Required to be «Viewed» documents"
    BUT_VIEWEDSTATUSDOCS = "TO VIEW"
    DOCS_PaymentOutgoing = "Outgoing payment 'Account payable"
    DOCS_PaymentIncoming = "Incoming payment 'Account receivable"
    DOCS_NotPaymentDoc = "Payment not set"
    DOCS_APPROVE = "Approve"
    But_Approve = "Approve"
    DOCS_NotAgreedYet = "Document not agreed yet"
    DOCS_RefuseApp = "Refuse approval"
    But_RefuseApp = "Refuse"
    DOCS_BusinessProcessSteps = "Business process steps"
    DOCS_StatusCompletion = "Completion status"
    DOCS_Currency = "Currency code"
    DOCS_ToView = "To view"
    DOCS_ParallelReconciliation = "Parallel agree process"
    DOCS_Messages = "Messages"
    But_DEPUTY = "Deputy"
    DOCS_VersionFiles = "Document files"
    DOCS_LocationPaper = "Hard copy location"
    DOCS_Notifications = "Notifications"
    DOCS_Resolutions = "Resolutions"
    DOCS_APROVAL = "Approval"
    DOCS_COMPLETION = "Completion"
    DOCS_History = "History"
    DOCS_StatusPayment = "Payment status"
    DOCS_Visa = "Agree process"
    DOCS_MEETINGS = "Meetings"
    DOCS_Viewed1 = "Document viewed"
    DOCS_SystemMessage = "System comments"
    DOCS_Resources = "Resources"
    DOCS_PARTNERS = "Partners"
    DOCS_Reviews = "Reviews"
    DOCS_Review = "Review"
    DOCS_ToReview1 = "Request for review"
    DOCS_ReconciliationREQUIRED = "Agree process is required"
    DOCS_Department = "Department"
    DOCS_VersionMainAdd = "Addition to main version"
    DOCS_NOTFOUND = "Information not found"
    DOCS_ACT = "Activity type"
    DOCS_All = "All"
    DOCS_Author = "Author"
    DOCS_Correspondent = "Correspondents, Distribution list"
    DOCS_DateActivation = "Date received"
    DOCS_DateCompletion = "Completion date"
    DOCS_Description = "Short description"
    DOCS_DocID = "Document ID"
    DOCS_ErrorSMTP = "ERROR during SMTP call "
    DOCS_EXPIREDSEC = "Access is expired"
    DOCS_FROM1 = "from"
    DOCS_Name = "Document title"
    DOCS_NoAccess = "Access denied"
    DOCS_NotificationNotCompletedDoc = "Notification - NOT completed document"
    DOCS_NotificationSentTo = "Notification has been sent to the following users"
    DOCS_PartnerName = "Partner Name"
    DOCS_Reconciliation = "Agree process"
    DOCS_Resolution = "Resolution"
    DOCS_Sender = "Sender"
    DOCS_SendNotification = "Send e-mail notification"
    DOCS_STATUSHOLD = "Registration is inactive "
    DOCS_UsersNotFound = "Users not found"
    USER_NOEMail = "No e-mail address or wrong e-mail address"
  case "3" '---------------------------------------------------------------------------- CZ
    DOCS_UnderDevelopment = "Ve stádiu vývoje"
    DOCS_Mark = "Vytvořit poznámku"
    DOCS_Viewed = "Přečetl/a dokument"
    DOCS_View = "Přečetl"
    DOCS_MakeCompleted = "Nastavit status «Dokončeno»"
    But_MakeCompleted = "Status «Dokončeno»"
    DOCS_CreateComment = "Vytvořit komentář"
    But_Comment = "Komentář"
    DOCS_AGREE = "Odsouhlasit"
    But_AGREE = "Odsouhlasit"
    DOCS_Refuse = "Zamítnout odsouhlasení"
    But_Refuse = "Odmítnout"
    DOCS_RequestCompleted = "Poslat požadavek na nastavení statusu «Dokončeno»"
    But_RequestCompleted = "Požadavek «Dokončeno»"
    DOCS_NameAproval = "Autorizující"
    DOCS_NameCreation = "Záznam vytvořil"
    DOCS_NameControl = "Kontrolu provedl"
    DOCS_NameResponsible = "Odpovědná osoba"
    DOCS_AmountDoc = "Částka dokumentu"
    DOCS_Signing = "K odsouhlasení"
    DOCS_Approving = "Ke schválení"
    DOCS_Approved = "Schváleno"
    DOCS_RefusedApp = "Schválení zamítnuto"
    DOCS_Completed = "Dokončeno"
    DOCS_Actual = "V procesu"
    DOCS_Cancelled = "Zrušeno"
    But_Statuses = "Stav"
    But_StatusesDesc = "Aktuální stav dokumentu"
    But_Actions = "Činnosti"
    But_ActionsDesc = "Činnosti spojené s dokumentem"
    DOCS_Inactive = "Dokument je neaktivní"
    DOCS_Active = "Dokument je aktivní"
    DOCS_UNAPPROVED = "Dokumenty vyžadující vaše odsouhlasení"
    BUT_VISA = "K REVIZI"
    DOCS_YouAreResponsible = "Dokumenty, jejichž současný uživatel je zároveň jejich odpovědným zpracovatelem"
    BUT_RESPONSIBLE = "ODPOVÍDÁ"
    DOCS_ListToReconcile = "Seznam schvalujících"
    DOCS_NextStepToReconcile = "Postupné odsouhlasování - další úroveň"
    DOCS_Reconciled = "Odsouhlaseno"
    DOCS_Refused = "Odsouhlasení zamítnuto"
    DOCS_VersionFile = "Dokument"
    DOCS_GoDoc = "Přejít ke kartě dokumentu"
    DOCS_Go = "Přejít"
    BUT_URGENT = "URGENTNÍ"
    DOCS_EXPIRED = "S blížícím se termínem dokončení a po termínu dokončení"
    BUT_CREATED = "VYTVOŘENÍ"
    DOCS_YouAreCreator = "Vámi vytvořené nedokončené dokumenty,"
    BUT_COMPLETION1 = "DOKONČENÍ"
    DOCS_NOTCOMPLETED = "Dokumenty vyžadující vaši poznámku o dokončení (vy tyto dokumenty schvalujete nebo kontrolujete)"
    BUT_CONTROL = "KONTROLA"
    DOCS_UnderControl = "Dokumenty jsou KONTROLOVÁNY"
    BUT_NEWDOCS = "NOVÉ"
    DOCS_NewDocs = "Nové dokumenty"
    DOCS_InOffice = "Momentálně jsem v kanceláři"
    DOCS_OutOfOffice = "Momentálně nejsem v kanceláři"
    DOCS_PushToChange = "Stiskněte pro změnu "
    BUT_OUTOFFICE = "Mimo kancelář"
    BUT_INOFFICE = "V kanceláři"
    DOCS_PaymentOutgoingIncompleted = "Doklady s neprovedenými odchozími platbami"
    DOCS_StatusRequireToBePaid = " - k úhradě"
    BUT_Payments = "PLATBY"
    DOCS_PaymentIncomingIncompleted = "Doklady s neprovedenými příchozími platbami"
    DOCS_GetListDocEMail = "Obdržet e-mailem seznam dokumentů"
    DOCS_SendAction = "Poslat e-mailem pokyn do PayDox k provedení aktivity"
    DOCS_eMailClientWarning = "E-mailový klient PayDoxu - Komunikace s PayDox přes e-mail. neměňte, prosím, obsah tohoto e-mailu!"
    DOCS_NotPreviousYet = "Předchozí krok odsouhlasení není dokončen"
    DOCS_StatusPaymentPaid = "Zaplaceno"
    DOCS_StatusPaymentNotPaid = "Nezaplaceno"
    DOCS_StatusPaymentToBePaid = "Nezaplaceno, nutné zaplatit"
    DOCS_StatusPaymentSentToBePaid = "Nezaplaceno, vystaveno k úhradě"
    DOCS_StatusPaymentToPay = "Nezaplaceno. Vydáno nařízení k platbě"
    DOCS_StatusExistsButNotDefined = "Status nebyl zadán"
    DOCS_Comments = "Komentáře"
    DOCS_Contacts = "Kontakty"
    DOCS_NotificationDoc = "Upozornění na dokument"
    DOCS_ListToView = "Seznam osob, které dokument četly"
    DOCS_UploadFile = "Načíst soubor do systému"
    But_SendFile = "Odeslat soubor"
    DOC_SendFile = "Připojte soubor k e-mailu. V případě potřeby uveďte svůj komentář po klíčovém slově COMMENT ="
    DOCS_FilesReceived = "přijato souborů"
    DOCS_FileNotUploaded = "Soubor nebyl načten - v názvu souboru, prosím, používejte jen standardní anglické znaky. Nepoužívejte v názvu souboru mezery."
    DOCS_FileReceived = "Soubor byl obdržen prostřednictvím e-mailu"
    DOCS_GetDocEMail = "Požadavek na obdržení tohoto dokumentu e-mailem (včetně připojených souborů)"
    But_GetDoc = "Obdržet dokument"
    USER_NOSMS = "Odeslání SMS nebylo zadáno"
    DOCS_SMSError = "Chyba spojení s modemem GSM"
    DOCS_GSMModemBusy = "Modem GSM je obsazen. Zpráva nebyla odeslána"
    DOCS_WRONGPHONECELL = "Chybné číslo mobilního telefonu. Zadejte celé číslo včetně znaku + na začátku čísla"
    DOCS_GSMModemError1 = "Modem GSM není připojen nebo je zadáno chybné číslo portu: "
    DOCS_SMSSent = "SMS odeslána"
    DOCS_StatusPaymentToBePaidPart = "Zaplaceno částečně, nutno doplatit"
    DOCS_StatusPaymentToPayPart = "Nezaplaceno. Vydán pokyn k ČÁSTEČNÉ úhradě"
    DOCS_StatusPaymentToPayPartRest = "Vydán pokyn k zaplaceí ZBÝVAJÍCÍ částky"
    DOCS_StatusPaymentPaidPart = "Zaplaceno ČÁSTEČNĚ"
    DOCS_StatusPaymentToPayPartPart = "Vydán pokyn k zaplacení další ČÁSTI"
    DOCS_ViewedStatusDocs = "Dokumenty, které je třeba přečíst"
    BUT_VIEWEDSTATUSDOCS = "PŘEČÍST"
    DOCS_PaymentOutgoing = "Odchozí platba"
    DOCS_PaymentIncoming = "Příchozí platba"
    DOCS_NotPaymentDoc = "Platba nebyla zadána"
    DOCS_APPROVE = "Schválit"
    But_Approve = "Schválit"
    DOCS_NotAgreedYet = "Odsouhlasení dokumentu není dokončené"
    DOCS_RefuseApp = "Zamítnout schválení"
    But_RefuseApp = "Odmítnout"
    DOCS_BusinessProcessSteps = "Etapy obchodního procesu"
    DOCS_StatusCompletion = "Status dokončení"
    DOCS_Currency = "Kód valuty"
    DOCS_ToView = "Přečíst"
    DOCS_ParallelReconciliation = "Paralelní odsouhlasování"
    DOCS_Messages = "Zprávy"
    But_DEPUTY = "Zástupce"
    DOCS_VersionFiles = "Soubory s dokumentem"
    DOCS_LocationPaper = "Umístění originálního dokumentu"
    DOCS_Notifications = "Oznámení"
    DOCS_Resolutions = "Rezoluce"
    DOCS_APROVAL = "Schválení"
    DOCS_COMPLETION = "Dokončení"
    DOCS_History = "Průběh vytváření dokumentu"
    DOCS_StatusPayment = "Status platby"
    DOCS_Visa = "Odsouhlasení"
    DOCS_MEETINGS = "Porady, meetingy"
    DOCS_Viewed1 = "Přečetli dokument"
    DOCS_SystemMessage = "Systémové zprávy"
    DOCS_Resources = "Zdroje"
    DOCS_PARTNERS = "Smluvní partněři"
    DOCS_Reviews = "Recenze"
    DOCS_Review = "Recenze"
    DOCS_ToReview1 = "Požadavek na recenzi"
    DOCS_ReconciliationREQUIRED = "Vyžaduje odsouhlasení"
    DOCS_Department = "Útvar"
    DOCS_VersionMainAdd = "Dodatek k základní verzi"
    DOCS_NOTFOUND = "Informace nebyla nalezena"
    DOCS_ACT = "Typ aktivity"
    DOCS_All = "Všechny"
    DOCS_Author = "Autor"
    DOCS_Correspondent = "Adresáti, seznam pro rozeslání"
    DOCS_DateActivation = "Datum registrace"
    DOCS_DateCompletion = "Datum dokončení"
    DOCS_Description = "Krátký obsah"
    DOCS_DocID = "Číslo dokumentu"
    DOCS_ErrorSMTP = "CHYBA při volání SMTP - zkontrolujte, jestli je ustanoven standardní komponent Windows Default SMTP Virtual Server pro IIS"
    DOCS_EXPIREDSEC = "Přístupová práva vypršela"
    DOCS_FROM1 = "z"
    DOCS_Name = "Název dokumentu"
    DOCS_NoAccess = "Přístup k tomuto dokumentu není povolen"
    DOCS_NotificationNotCompletedDoc = "Upozornění na nedokončený dokument"
    DOCS_NotificationSentTo = "Uživatelům bylo odesláno upozornění"
    DOCS_PartnerName = "Smluvní strana"
    DOCS_Reconciliation = "Odsouhlasení"
    DOCS_Resolution = "Rezoluce"
    DOCS_Sender = "Odesílatel"
    DOCS_SendNotification = "Odeslat e-mailové oznámení"
    DOCS_STATUSHOLD = "Přístup k systému je pozastaven"
    DOCS_UsersNotFound = "Uživatelé nebyli nalezeni"
    USER_NOEMail = "Chybí e-mailová adresa"
  End Select

  MailTexts ( 1)=DOCS_UnderDevelopment
  MailTexts ( 2)=DOCS_Mark
  MailTexts ( 3)=DOCS_Viewed
  MailTexts ( 4)=DOCS_View
  MailTexts ( 5)=DOCS_MakeCompleted
  MailTexts ( 6)=But_MakeCompleted
  MailTexts ( 7)=DOCS_CreateComment
  MailTexts ( 8)=But_Comment
  MailTexts ( 9)=DOCS_AGREE
  MailTexts (10)=But_AGREE
  MailTexts (11)=DOCS_Refuse
  MailTexts (12)=But_Refuse
  MailTexts (13)=DOCS_RequestCompleted
  MailTexts (14)=But_RequestCompleted
  MailTexts (15)=DOCS_RequestCompleted
  MailTexts (16)=But_RequestCompleted
  MailTexts (17)=DOCS_NameAproval
  MailTexts (18)=DOCS_NameCreation
  MailTexts (19)=DOCS_NameControl
  MailTexts (20)=DOCS_NameResponsible
  MailTexts (21)=DOCS_AmountDoc 
  MailTexts (22)=DOCS_Signing
  MailTexts (23)=DOCS_Approving
  MailTexts (24)=DOCS_Approved
  MailTexts (25)=DOCS_RefusedApp
  MailTexts (26)=VAR_StatusCompletion
  MailTexts (27)=VAR_StatusCancelled
  MailTexts (28)=DOCS_Completed
  MailTexts (29)=DOCS_Actual
  MailTexts (30)=DOCS_Cancelled
  MailTexts (31)=But_Statuses
  MailTexts (32)=But_StatusesDesc
  MailTexts (33)=But_Actions
  MailTexts (34)=But_ActionsDesc
  MailTexts (35)=VAR_InActiveTask
  MailTexts (36)=DOCS_Inactive
  MailTexts (37)=DOCS_Active
  MailTexts (38)=DOCS_UNAPPROVED
  MailTexts (39)=BUT_VISA
  MailTexts (40)=DOCS_YouAreResponsible
  MailTexts (41)=BUT_RESPONSIBLE
  MailTexts (42)=DOCS_ListToReconcile
  MailTexts (43)=DOCS_NextStepToReconcile
  MailTexts (44)=DOCS_Reconciled
  MailTexts (45)=DOCS_Refused
  MailTexts (46)=DOCS_VersionFiles
  MailTexts (47)=DOCS_VersionFile
  MailTexts (48)=DOCS_GoDoc 
  MailTexts (49)=DOCS_Go

  MailTexts(50)=BUT_URGENT
  MailTexts(51)=DOCS_EXPIRED
  MailTexts(52)=BUT_CREATED
  MailTexts(53)=DOCS_YouAreCreator
  MailTexts(54)=BUT_COMPLETION1
  MailTexts(55)=DOCS_NOTCOMPLETED
  MailTexts(56)=BUT_CONTROL
  MailTexts(57)=DOCS_UnderControl
  MailTexts(58)=BUT_NEWDOCS
  MailTexts(59)=DOCS_NewDocs
  MailTexts(60)=DOCS_InOffice
  MailTexts(61)=DOCS_OutOfOffice
  MailTexts(62)=DOCS_PushToChange
  MailTexts(63)=BUT_OUTOFFICE
  MailTexts(64)=BUT_INOFFICE
  MailTexts(65)=DOCS_PaymentOutgoingIncompleted
  MailTexts(66)=DOCS_StatusRequireToBePaid
  MailTexts(67)=BUT_Payments
  MailTexts(68)=DOCS_PaymentIncomingIncompleted
  MailTexts(69)=DOCS_GetListDocEMail
  MailTexts(70)=DOCS_SendAction
  MailTexts(71)=DOCS_eMailClientWarning
  MailTexts(72)=DOCS_Refuse
  MailTexts(73)=But_Refuse
  MailTexts(74)=DOCS_NotPreviousYet
  MailTexts(75)=DOCS_StatusPaymentPaid
  MailTexts(76)=VAR_StatusRequestCompletion
  MailTexts(77) = DOCS_StatusPaymentNotPaid
  MailTexts(78) = DOCS_StatusPaymentToBePaid
  MailTexts(79) = DOCS_StatusPaymentSentToBePaid
  MailTexts(80) = DOCS_StatusPaymentToPay
  MailTexts(81) = DOCS_StatusPaymentPaid
  MailTexts(82) = DOCS_StatusExistsButNotDefined
  MailTexts(83) = DOCS_Comments 
  MailTexts(84) = DOCS_Contacts
  MailTexts(85) = DOCS_NotificationDoc
  MailTexts(86) = DOCS_Viewed1
  MailTexts(87) = DOCS_ListToView
  MailTexts(88) = DOCS_UploadFile
  MailTexts(89) = But_SendFile
  MailTexts(90) = DOC_SendFile
  MailTexts(91) = DOCS_FilesReceived
  MailTexts(92) = DOCS_FileNotUploaded
  MailTexts(93) = DOCS_FileReceived
  MailTexts(94) = DOCS_GetDocEMail
  MailTexts(95) = But_GetDoc
  MailTexts(96) = USER_NOSMS
  MailTexts(97) = DOCS_SMSError
  MailTexts(98) = DOCS_GSMModemBusy
  MailTexts(99) = DOCS_WRONGPHONECELL
  MailTexts(100) = DOCS_GSMModemError1
  MailTexts(101) = DOCS_SMSSent
  MailTexts(102) = DOCS_OutOfOffice

  MailTexts(103) = DOCS_StatusPaymentToBePaidPart
  MailTexts(104) = DOCS_StatusPaymentToPayPart
  MailTexts(105) = DOCS_StatusPaymentToPayPartRest
  MailTexts(106) = DOCS_StatusPaymentPaidPart
  MailTexts(107) = DOCS_StatusPaymentToPayPartPart
  MailTexts(108) = DOCS_ViewedStatusDocs
  MailTexts(109) = BUT_VIEWEDSTATUSDOCS

  MailTexts(110) = DOCS_PaymentOutgoing
  MailTexts(111) = DOCS_PaymentIncoming
  MailTexts(112) = DOCS_NotPaymentDoc
  MailTexts(113) = VAR_MaxEMailAttachedFileSize
  MailTexts(114) = Var_ApprovalPermitted
  MailTexts(115) = DOCS_APPROVE
  MailTexts(116) = But_Approve
  MailTexts(117) = DOCS_NotAgreedYet
  MailTexts(118) = DOCS_RefuseApp
  MailTexts(119) = But_RefuseApp
  MailTexts(120) = Var_nDaysToReconcile
  MailTexts(121) = DOCS_BusinessProcessSteps
  MailTexts(122) = DOCS_StatusCompletion
  MailTexts(123) = DOCS_Currency
  MailTexts(124) = EMailFieldList
  MailTexts(125) = DOCS_ToView
  MailTexts(126) = DOCS_ParallelReconciliation 
  MailTexts(127) = Var_ApprovalIfAllAgree
  MailTexts(128) = DOCS_Messages
  MailTexts(129) = But_DEPUTY

  MailTexts(130) = DOCS_VersionFiles
  MailTexts(131) = DOCS_LocationPaper
  MailTexts(132) = DOCS_Notifications
  MailTexts(133) = DOCS_Resolutions
  MailTexts(134) = DOCS_APROVAL
  MailTexts(135) = DOCS_COMPLETION
  MailTexts(136) = DOCS_History
  MailTexts(137) = DOCS_StatusPayment
  MailTexts(138) = DOCS_Visa
  MailTexts(139) = DOCS_MEETINGS
  MailTexts(140) = DOCS_Viewed1
  MailTexts(141) = DOCS_SystemMessage
  MailTexts(142) = DOCS_Resources
  MailTexts(143) = DOCS_Messages
  MailTexts(144) = Var_ReconciliationIfAllAgree
  MailTexts(145) = DOCS_KeyWordUsers
  MailTexts(146) = Var_ReconciliationRoleToo
  MailTexts(147) = DOCS_PARTNERS
  MailTexts(148) = DOCS_Reviews
  MailTexts(149) = Var_NotToSendAttachments
  MailTexts(150) = Var_NotToSendAttachmentsXML
  MailTexts(151) = DOCS_Review
  MailTexts(152) = DOCS_ToReview1
  MailTexts(153) = DOCS_ReconciliationREQUIRED
  MailTexts(154) = DOCS_Department
  MailTexts(155) = VAR_ContentType
  MailTexts(156) = VAR_Charset
  MailTexts(157) = DOCS_VersionMainAdd
  MailTexts(158) = VAR_CharsetWebPages
  MailTexts(159) = VAR_AttachFilesToEMail
  MailTexts(160) = VAR_UserMessageToEMail
  MailTexts(161) = DOCS_NOTFOUND
  MailTexts(162) = VAR_USESSL 
  MailTexts(163) = VAR_LanguageSeparator

  'Тексты почтовых уведомлений
  Select case UCase(Lang)
  case "RU" '--------------------------------------------------------------------------- RU
    SIT_AgreeTimeExceeded = "Вы просрочили согласование документа. Об этом будет уведомлен Ваш руководитель."
'20090622 - Заявка ТКП
    SIT_AgreeTimeExceededComOffer = "Истек срок, отведенный на согласование коммерческого предложения"
    SIT_AgreeDelaying = "Вы задерживаете согласование документа"
    SIT_UserDelayedAgree1 = "Пользователь "
    SIT_UserDelayedAgree2 = " просрочил согласование документа."
    SIT_OneDayForTaskCompletionSubj = "До истечения срока исполнения поручения остался 1 день"
    SIT_2DaysForTaskCompletionSubj = "До истечения срока исполнения поручения осталось 2 дня"
    SIT_3DaysForTaskCompletionSubj = "До истечения срока исполнения поручения осталось 3 дня"
    SIT_OneDayForTaskCompletionBody1 = "<BR><B>Уважаемый коллега, информируем Вас, что завтра ("
    SIT_OneDayForTaskCompletionBody2 = ") истекает срок выполнения поручения (карточка поручения и ссылка на него в систему документооборота ниже в письме), находящегося на контроле у "
    SIT_OneDayForTaskCompletionBody3 = ", просим Вас найти время заполнить отчёт об исполнении поручения.</B><BR>"

    SIT_2DaysForTaskCompletionBody1 = "<BR><B>Уважаемый коллега, информируем Вас, что через 2 дня ("
    SIT_2DaysForTaskCompletionBody2 = ") истекает срок выполнения поручения (карточка поручения и ссылка на него в систему документооборота ниже в письме), находящегося на контроле у "
    SIT_2DaysForTaskCompletionBody3 = ", просим Вас найти время заполнить отчёт об исполнении поручения.</B><BR>"

    SIT_3DaysForTaskCompletionBody1 = "<BR><B>Уважаемый коллега, информируем Вас, что через 3 дня ("
    SIT_3DaysForTaskCompletionBody2 = ") истекает срок выполнения поручения (карточка поручения и ссылка на него в систему документооборота ниже в письме), находящегося на контроле у "
    SIT_3DaysForTaskCompletionBody3 = ", просим Вас найти время заполнить отчёт об исполнении поручения.</B><BR>"

    SIT_TheDayOfTaskCompletionSubj = "Сегодня истекает срок исполнения поручения!"

    SIT_AfterTheDayOfTaskCompletionSubj = "Истек срок исполнения поручения!"

    SIT_TheDayOfTaskCompletionBody1 = "<BR><B>Уважаемый коллега! Сегодня ("
    SIT_TheDayOfTaskCompletionBody2 = ") истекает срок отчётности по поручению (карточка поручения и ссылка на него в систему документооборота ниже в письме), находящемуся на исполнении у "
    SIT_TheDayOfTaskCompletionBody3 = " и на контроле у "

    SIT_AfterTheDayOfTaskCompletionBody1 = "<BR><B>Уважаемый коллега! "
    SIT_AfterTheDayOfTaskCompletionBody2 = " истек срок отчётности по поручению (карточка поручения и ссылка на него в систему документооборота ниже в письме), находящемуся на исполнении у "
    SIT_AfterTheDayOfTaskCompletionBody3 = " и на контроле у "

    SIT_RTI_TheDayOfTaskCompletionBody4 = ". Просим Вас отчитаться об исполнении поручения.</B><BR>"
    SIT_RTI_TaskCompletionDelayedSubj = "Неисполнение поручения в срок"
    SIT_RTI_TaskCompletionDelayedBody = "<BR><B>Уважаемый коллега, истек срок исполнения поручения (карточка поручения и ссылка на него в систему документооборота ниже в письме). Просим вас отчитаться о выполнении поручения</B><BR>"
    'vnik_fix_begin
    SIT_TheDayOfTaskCompletionBody4 = ". Просим Вас отчитаться об исполнении, иначе по истечении 5-ти календарных дней к Вам будут применены штрафные санкции.</B><BR>"
    'vnik_fix_end
'Запрос №1 - СИБ - start
    SIT_TheDayOfTaskCompletionBody4_SIB = ". Просим Вас отчитаться об исполнении, иначе по истечении 3-ёх дней к Вам будут применены штрафные санкции.</B><BR>"
'Запрос №1 - СИБ - end
    SIT_TaskCompletionDelayedSubj = "Понижающее значение при выплате бонуса"
    'vnik_fix_begin
    SIT_TaskCompletionDelayedBody = "<BR><B>Уважаемый коллега, из-за срыва сроков исполнения поручения (карточка поручения и ссылка на него в систему документооборота ниже в письме) и задержки предоставления отчётности более чем на 5 дней (включая назначенный руководителем отчётный день), к Вам будут применены санкции в соответствии с приказом № 153 от 30.09.2008, что отразится на размере Вашего квартального бонуса. Тем не менее, это не освобождает Вас от необходимости выполнения поручения и предоставления отчета о выполнении.</B><BR>"
    'vnik_fix_end
'Запрос №1 - СИБ - start
    SIT_TaskCompletionDelayedBody_SIB = "<BR><B>Уважаемый коллега, из-за срыва сроков исполнения поручения (карточка поручения и ссылка на него в систему документооборота ниже в письме) и задержки предоставления отчётности более чем на 3 дня (включая назначенный руководителем отчётный день), к Вам будут применены санкции, что отразится на размере Вашего квартального бонуса. Тем не менее, это не освобождает Вас от необходимости выполнения поручения и предоставления отчета о выполнении.</B><BR>"
'Запрос №1 - СИБ - end
    SIT_TaskCompletionExpiredSubj = "Истек срок выполнения поручения. Об этом будет уведомлен Ваш руководитель."
    SIT_TaskCompletionExpiredBody = "<BR><B>Уважаемый коллега, информируем Вас, что истек срок выполнения поручения (карточка поручения и ссылка на него в систему документооборота ниже в письме), просим Вас найти время заполнить отчёт об исполнении поручения.</B><BR>"
    SIT_PaymentCompletionExpiredSubj = "Истек срок выполнения заявки. Об этом будет уведомлен Ваш руководитель."
    SIT_TaskCompletionExpiresSoonSubj = "Истекает срок выполнения поручения."
    SIT_TaskCompletionExpiresSoonBody = "<BR><B>Уважаемый коллега, информируем Вас, что истекает срок выполнения поручения (карточка поручения и ссылка на него в систему документооборота ниже в письме), просим Вас найти время заполнить отчёт об исполнении поручения.</B><BR>"
    SIT_OneDayForPaymentCompletionSubj = "До истечения срока выполнение заявки остался 1 день."
    SIT_OneDayForPaymentCompletionBody = "<BR><B>Уважаемый коллега, информируем Вас, что истекает срок выполнения заявки (карточка заявки и ссылка на нее в систему документооборота ниже в письме).</B><BR>"
    SIT_UserDelayedTask1 = "Сотрудник "
    SIT_UserDelayedTask2 = " не выполнил поручение в срок."
    SIT_UserDelayedPayment1 = "Сотрудник "
    SIT_UserDelayedPayment2 = " не выполнил заявку в срок."
'    SIT_MoreThanOneLeader1 = "В подразделении "
'    SIT_MoreThanOneLeader2 = " больше одного руководителя"

    'rmanyushin 119579 19.08.2010 Start
	STS_HolidayRequest_Refused = "Заявление на отпуск не утверждено."
	'rmanyushin 119579 19.08.2010 End

'Запрос №37 - СТС - start
  STS_TaskCompletionExpires = "<BR>Уважаемый коллега! Сегодня истекает срок отчётности по поручению #DOCID#, находящемуся у Вас на исполнении#CONTROL#. Просим Вас отчитаться об исполнении, иначе по истечении трех рабочих дней к Вам будет применена система частичного депремирования, согласно Положению «О премировании работников»."
  STS_TaskCompletionExpiresControl = " и на контроле у #NAMECONTROL#"
  STS_TaskCompletionExpired1 = "<BR>Уважаемый коллега, Вы сорвали сроки исполнения поручения #DOCID#, напоминаем Вам, что при задержке предоставления отчётности более чем на три рабочих дня (включая назначенный руководителем отчётный день), к Вам будут применены санкции, что может отразиться на размере Вашей премии. Тем не менее, это не освобождает Вас от необходимости выполнения данного поручения и предоставления отчета о его выполнении."
  STS_TaskCompletionExpired4 = "<BR>Уважаемый коллега, из-за срыва сроков исполнения поручения #DOCID# и задержки предоставления отчётности более чем на три рабочих дня (включая назначенный руководителем отчётный день), к Вам будут применены санкции, что может отразиться на размере Вашей премии. Тем не менее, это не освобождает Вас от необходимости выполнения данного поручения и предоставления отчета о его выполнении."
  STS_ReconciliationExpires =  "<BR>Уважаемый коллега! Сегодня истекает срок согласования документа #DOCID#. Просим Вас согласовать документ, иначе по истечении трех рабочих дней к Вам будет применена система частичного депремирования, согласно Положению «О премировании работников»." 
  STS_ReconciliationExpired1 = "<BR>Уважаемый коллега, Вы сорвали сроки согласования документа #DOCID#, напоминаем Вам, что при задержке согласования более чем на три рабочих дня, к Вам будут применены санкции, что может отразиться на размере Вашей премии. Тем не менее, это не освобождает Вас от необходимости согласования данного документа." 
  STS_ReconciliationExpired4 = "<BR>Уважаемый коллега, из-за срыва сроков согласования документа #DOCID# более чем на три рабочих дня, к Вам будут применены санкции, что может отразиться на размере Вашей премии. Тем не менее, это не освобождает Вас от необходимости согласования данного документа."  
  STS_PenaltyMessage = " Понижающее значение при выплате премий составляет #SUM#"
'Запрос №37 - СТС - end

  case "" '----------------------------------------------------------------------------- EN
    SIT_AgreeTimeExceeded = "You have not met the document agreement deadline. Your superior will be informed."
'20090622 - Заявка ТКП
    SIT_AgreeTimeExceededComOffer = "Time for commercial offer endorsement is up"
    SIT_AgreeDelaying = "You are delaying the document agreement process"
    SIT_UserDelayedAgree1 = "User "
    SIT_UserDelayedAgree2 = " has not met the document agreement deadline."
    SIT_OneDayForTaskCompletionSubj = "The task fulfilment deadline will expire in 1 day."
    SIT_OneDayForTaskCompletionBody1 = "<BR><B>Dear colleague, we hereby inform you that ("
    SIT_OneDayForTaskCompletionBody2 = ") the task fulfilment time limit expires tomorrow (the task card and the link to it in the document management system are below). The task is now being checked by "
    SIT_OneDayForTaskCompletionBody3 = ". Please, complete the task fulfilment report.</B><BR>"
    SIT_TheDayOfTaskCompletionSubj = "Task fulfilment deadline today!"
    SIT_TheDayOfTaskCompletionBody1 = "<BR><B>Dear colleague! It is ("
    SIT_TheDayOfTaskCompletionBody2 = ") the task fulfilment report deadline today (the task card and the link to it in the document management system are below). The task is now being fulfilled by "
    SIT_TheDayOfTaskCompletionBody3 = " and checked by "
    'vnik_fix_begin
    SIT_TheDayOfTaskCompletionBody4 = ". Please create the task fulfilment report. If the report is not created within 5 calendar days, punitive sanctions will be applied to you.</B><BR>"
    'vnik_fix_end
'Запрос №1 - СИБ - start
    SIT_TheDayOfTaskCompletionBody4_SIB = ". Please create the task fulfilment report. If the report is not created within 3 days, punitive sanctions will be applied to you.</B><BR>"
'Запрос №1 - СИБ - end
    SIT_TaskCompletionDelayedSubj = "Based on the sanctions, your bonus amount will be decreased."
    'vnik_fix_begin
    SIT_TaskCompletionDelayedBody = "<BR><B>Dear colleague, as you have not met the task fulfilment deadline (the task card and the link to it in the document management system are below) and have not submitted the order fulfilment report within 5 days of the deadline date (including the report submission date set by the manager), sanctions in compliance with the order № 153 from 30.09.2008 will be applied to you. The sanctions will impact the amount of your quarterly bonus. Nevertheless, this does not relieve you of the duty to fulfil the order and submit the order fulfilment report.</B><BR>"
    'vnik_fix_end
'Запрос №1 - СИБ - start
    SIT_TaskCompletionDelayedBody_SIB = "<BR><B>Dear colleague, as you have not met the task fulfilment deadline (the task card and the link to it in the document management system are below) and have not submitted the order fulfilment report within 3 days of the deadline date (including the report submission date set by the manager), sanctions will be applied to you. The sanctions will impact the amount of your quarterly bonus. Nevertheless, this does not relieve you of the duty to fulfil the order and submit the order fulfilment report.</B><BR>"
'Запрос №1 - СИБ - end
    SIT_TaskCompletionExpiredSubj = "The task fulfilment deadline has expired. Your superior will be informed about it."
    SIT_TaskCompletionExpiredBody = "<BR><B>Dear colleague, we hereby inform you that the task fulfilment deadline has expired (the task card and the link to it in the document management system are below). Please, complete the task fulfilment report.</B><BR>"
    SIT_PaymentCompletionExpiredSubj = "The order fulfilment deadline has expired. Your superior will be informed about it."
    SIT_TaskCompletionExpiresSoonSubj = "The task fulfilment deadline is coming."
    SIT_TaskCompletionExpiresSoonBody = "<BR><B>Dear colleague, we hereby inform you that the task fulfilment deadline is coming (the task card and the link to it in the document management system are below). Please, complete the task fulfilment report.</B><BR>"
    SIT_OneDayForPaymentCompletionSubj = "The order fulfilment deadline expires in 1 day."
    SIT_OneDayForPaymentCompletionBody = "<BR><B>Dear colleague, we hereby inform you that the order fulfilment deadline is coming (the order card and the link to it in the document management system are below).</B><BR>"
    SIT_UserDelayedTask1 = "Employee "
    SIT_UserDelayedTask2 = " did not fulfil the task by the set deadline."
    SIT_UserDelayedPayment1 = "Employee "
    SIT_UserDelayedPayment2 = " did not fulfil the order by the set deadline."
'    SIT_MoreThanOneLeader1 = "There is more than one manager in the organizational unit "
'    SIT_MoreThanOneLeader2 = ""

    'rmanyushin 119579 19.08.2010 Start
	STS_HolidayRequest_Refused = "The holiday request has not been approved."
	'rmanyushin 119579 19.08.2010 End

'Запрос №37 - СТС - start
  STS_TaskCompletionExpires = "<BR>Dear Colleague! Today is the deadline for the fulfillment of your order #DOCID#.#CONTROL# Please report on the order fulfillment. If you do not submit the order fulfillment report within 3 (three) following working days, you will be subject to partial bonus reduction in accordance with the Regulation “Concerning bonuses in CJSC SITRONICS Telecom Solutions”."
  STS_TaskCompletionExpiresControl = " The order fulfillment is being controlled by #NAMECONTROL#."
  STS_TaskCompletionExpired1 = "<BR>Dear Colleague! As you have not met the order #DOCID# fulfillment deadline, we notify you that if you do not submit the order fulfillment report within 3 (three) following working days (including the report submission date set by the manager), the appropriate sanctions will be applied to you. This will impact the amount of your bonus. Nevertheless, this does not acquit you of the duty to fulfill the order and submit the order fulfillment report."
  STS_TaskCompletionExpired4 = "<BR>Dear Colleague! As you have not met the order #DOCID# fulfillment deadline and have not submitted the order fulfillment report within 3 (three) working days (including the report submission date set by the manager), the appropriate sanctions will be applied to you. This will impact the amount of your bonus. Nevertheless, this does not acquit you of the duty to fulfill the order and submit the order fulfillment report."
  STS_ReconciliationExpires =  "<BR>Dear Colleague! Today is the deadline for your agreement with the document #DOCID#. Please agree with the document. If you do not agree within 3 (three) following working days, you will be subject to partial bonus reduction in accordance with the Regulation “Concerning bonuses in CJSC SITRONICS Telecom Solutions”." 
  STS_ReconciliationExpired1 = "<BR>Dear Colleague! As you have not met the document #DOCID# agreement deadline, we notify you that if you do not agree within 3 (three) following working days, the appropriate sanctions will be applied to you. This will impact the amount of your bonus. Nevertheless, this does not acquit you of the duty to agree with the document." 
  STS_ReconciliationExpired4 = "<BR>Dear Colleague! As you have not agreed with the document #DOCID# within 3 (three) working days after exceeding the deadline, the appropriate sanctions will be applied to you. This will impact the amount of your bonus. Nevertheless, this does not acquit you of the duty to agree with the document."  
  STS_PenaltyMessage = " Your bonus will be reduced by #SUM#"
'Запрос №37 - СТС - end

  case "3" '---------------------------------------------------------------------------- CZ
    SIT_AgreeTimeExceeded = "Nedodržel/-a jste lhůtu pro schvalování dokumentu. O tomto faktu bude informován Váš nadřízený."
'20090622 - Заявка ТКП
    SIT_AgreeTimeExceededComOffer = "Lhůta pro schválení obchodní nabídky vypršela"
    SIT_AgreeDelaying = "Zdržujete proces schválení dokumentu"
    SIT_UserDelayedAgree1 = "Uživatel "
    SIT_UserDelayedAgree2 = " nedodržel jste lhůtu pro schvalování dokumentu."
    SIT_OneDayForTaskCompletionSubj = "Do konce lhůty pro splnění úkolu zbývá 1 den"
    SIT_OneDayForTaskCompletionBody1 = "<BR><B>Dobrý den, tímto Vám oznamujeme, že ("
    SIT_OneDayForTaskCompletionBody2 = ") časová lhůta pro splnění úkolu vyprší zítra (karta úkolu a odkaz na úkol v systému pro správu dokumentů jsou uvedeny níže). Úkol nyní kontroluje "
    SIT_OneDayForTaskCompletionBody3 = ". Vytvořte, prosím, zprávu o splnění úkolu.</B><BR>"
    SIT_TheDayOfTaskCompletionSubj = "Dnes je poslední den lhůty pro splnění úkolu!"
    SIT_TheDayOfTaskCompletionBody1 = "<BR><B>Dobrý den, dnes ("
    SIT_TheDayOfTaskCompletionBody2 = ") vyprší lhůta pro dodání zprávy o splnění úkolu (karta úkolu a odkaz na úkol v systému pro správu dokumentů jsou uvedeny níže), který nyní plní "
    SIT_TheDayOfTaskCompletionBody3 = " a kontroluje "
    'vnik_fix_begin
    SIT_TheDayOfTaskCompletionBody4 = ". Vytvořte, prosím, zprávu o splnění úkolu. Pokud nebude zpráva vytvořena do 5 kalendářních dnů, bude vůči Vám uplatněn postih,</B><BR>"
    'vnik_fix_end
'Запрос №1 - СИБ - start
    SIT_TheDayOfTaskCompletionBody4_SIB = ". Vytvořte, prosím, zprávu o splnění úkolu. Pokud nebude zpráva vytvořena do 3 dnů, bude vůči Vám uplatněn postih,</B><BR>"
'Запрос №1 - СИБ - end
    SIT_TaskCompletionDelayedSubj = "v jehož důsledku Vám bude snížen bonus."
    'vnik_fix_begin
    SIT_TaskCompletionDelayedBody = "<BR><B>Dobrý den, vzhledem k tomu, že jste nedodržel/-a termín pro splnění úkolu (karta úkolu a odkaz na úkol v systému pro správu dokumentů jsou uvedeny níže) a nedodal jste zprávu o splnění úkolu do 5 dnů od termínu (včetně data stanoveného nadřízeným pro dodání této zprávy), bude vůči Vám uplatněn postih v souladu s nařízením č. 153 ze 30.09.2008, což ovlivní výši Vašeho čtvrtletního bonusu. Vaše povinnost splnit nařízení a poskytnout zprávu o splnění úkolu však nadále trvá.</B><BR>"
    'vnik_fix_end
'Запрос №1 - СИБ - start
    SIT_TaskCompletionDelayedBody_SIB = "<BR><B>Dobrý den, vzhledem k tomu, že jste nedodržel/-a termín pro splnění úkolu (karta úkolu a odkaz na úkol v systému pro správu dokumentů jsou uvedeny níže) a nedodal jste zprávu o splnění úkolu do 3 dnů od termínu (včetně data stanoveného nadřízeným pro dodání této zprávy), bude vůči Vám uplatněn postih, což ovlivní výši Vašeho čtvrtletního bonusu. Vaše povinnost splnit nařízení a poskytnout zprávu o splnění úkolu však nadále trvá.</B><BR>"
'Запрос №1 - СИБ - end
    SIT_TaskCompletionExpiredSubj = "Lhůta pro splnění úkolu vypršela. O tomto faktu bude informován Váš nadřízený."
    SIT_TaskCompletionExpiredBody = "<BR><B>Dobrý den, tímto Vám oznamujeme, že lhůta pro splnění úkolu vypršela (karta úkolu a odkaz na úkol v systému pro správu dokumentů jsou uvedeny níže). Vytvořte, prosím, zprávu o splnění úkolu.</B><BR>"
    SIT_PaymentCompletionExpiredSubj = "Lhůta pro splnění objednávky vypršela. O tomto faktu bude informován Váš nadřízený."
    SIT_TaskCompletionExpiresSoonSubj = "Blíží se konec lhůty pro splnění úkolu."
    SIT_TaskCompletionExpiresSoonBody = "<BR><B>Dobrý den, tímto Vám oznamujeme, že se blíží konec lhůty pro splnění úkolu (karta úkolu a odkaz na úkol v systému pro správu dokumentů jsou uvedeny níže). Vytvořte, prosím, zprávu o splnění úkolu.</B><BR>"
    SIT_OneDayForPaymentCompletionSubj = "Do konce lhůty pro splnění objednávky zbývá 1 den."
    SIT_OneDayForPaymentCompletionBody = "<BR><B>Dobrý den, tímto Vám oznamujeme, že se blíží konec lhůty pro splnění objednávky (karta objednávky a odkaz na úkol v systému pro správu dokumentů jsou uvedeny níže).</B><BR>"
    SIT_UserDelayedTask1 = "Zaměstnanec "
    SIT_UserDelayedTask2 = " nesplnil úkol do stanoveného termínu."
    SIT_UserDelayedPayment1 = "Zaměstnanec "
    SIT_UserDelayedPayment2 = " nesplnil objednávku do stanoveného termínu."
'    SIT_MoreThanOneLeader1 = "V této organizační jednotce je "
'    SIT_MoreThanOneLeader2 = " více než jeden vedoucí pracovník."

    'rmanyushin 119579 19.08.2010 Start
	STS_HolidayRequest_Refused = "Žádost o dovolenou nebyla schválena."
	'rmanyushin 119579 19.08.2010 End

'Запрос №37 - СТС - start
  STS_TaskCompletionExpires = "<BR>Vážená kolegyně, kolego! Dnes vyprší termín dodání zprávy o dokončení úkolu #DOCID#, u něhož jste byl stanoven jako odpovědná osoba a jehož kontrolou je pověřen #CONTROL#. Žádáme Vás, abyste podal zprávu o dokončení, jinak bude vůči Vám po vypršení 3 (tří) pracovních dnů uplatněna opatření systému částečného depremiování v souladu se směrnicí „O udělování prémií zaměstnancům“."
  STS_TaskCompletionExpiresControl = "#NAMECONTROL#"
  STS_TaskCompletionExpired1 = "<BR>Vážená kolegyně, kolego, nedodržel (a) jste termín dokončení úkolu #DOCID#. Připomínáme Vám, že v případě nedodání zprávy o dokončení v průběhu 3 (tří) pracovních dnů (včetně data dodání zprávy stanoveného vedoucím), budou vůči Vám uplatněny sankce, které mohou ovlivnit výši Vaší prémie. Nicméně, toto Vás neosvobozuje od povinnosti dokončení tohoto úkolu a doručení zprávy o jeho dokončení."
  STS_TaskCompletionExpired4 = "<BR>Vážená kolegyně, kolego, z důvodu nedodržení termínu dokončení úkolu #DOCID# a opoždění dodání zprávy o dokončení úkolu o více než 3 (tři) pracovní dny (včetně data dodání zprávy stanoveného vedoucím), budou vůči Vám uplatněny sankce, které mohou ovlivnit výši Vaší prémie. Nicméně toto Vás neosvobozuje od povinnosti dokončení tohoto úkolu a doručení zprávy o jeho dokončení."
  STS_ReconciliationExpires =  "<BR>Vážená/ý kolegyně/kolego! Dnes vyprší termín odsouhlasení dokumentu #DOCID#. Žádáme Vás, abyste odsouhlasil dokument, jinak bude vůči Vám po vypršení 3 (třech) pracovních dnů uplatněno opatření systému částečného snížení prémií v souladu se směrnicí „O udělování prémií zaměstnancům“." 
  STS_ReconciliationExpired1 = "<BR>Vážená/ý kolegyně/kolego, nedodržel (a) jste termín odsouhlasení dokumentu #DOCID#. Připomínáme Vám, že v případě zpoždění odsouhlasení dokumentu o více než 3 (tři) pracovní dny, budou vůči Vám uplatněny sankce, které mohou ovlivnit výši Vaší prémie. Nicméně, toto Vás neosvobozuje od povinnosti odsouhlasit tento dokument." 
  STS_ReconciliationExpired4 = "<BR>Vážená/ý kolegyně/kolego, z důvodu nedodržení termínu odsouhlasení dokumentu #DOCID# o více než 3 (tři) pracovní dny, budou vůči Vám uplatněny sankce, které mohou ovlivnit výši Vaší prémie. Nicméně toto Vás neosvobozuje od povinnosti odsouhlasit tento dokument."  
  STS_PenaltyMessage = " Vaše prémie bude snížena o #SUM#"
'Запрос №37 - СТС - end

  End Select
End Sub

Sub SetLangConstsForAutoReconcilation(ByVal Lang)
  Select case UCase(Lang)
  case "RU" '----- RU
    SIT_Agreed = "Согласовано"
    SIT_AutoAgreed = " / Согласовано по умолчанию"
  case ""   '----- EN
    SIT_Agreed = "Agreed"
    SIT_AutoAgreed = " / Agreed by default"
  case "3"  '----- CZ
    SIT_Agreed = "Odsouhlaseno"
    SIT_AutoAgreed = " / Schváleno defaultně"
  End Select
End Sub

%>