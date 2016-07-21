<%
'Define some user parameters before report run

'Possible variables:
'sReportGUID - Report GUID
'sReport_Name - Report name
'
'Use the following functions:
'
'ReportRequestFormFieldInput sTitle, sValue, sSQLContext, sDirGUID
'Inserts user report text form field to replace report SQL expression context
'Parameters:
'	sTitle - string value, parameter's title
'	sValue - string value, default parameter's value
'	sSQLContext - string value, context to be replaced by parameter's value in report SQL expression
'	sDirGUID - string value, directory GUID connected to this user report text form field
'
'ReportRequestTitle sTitle
'Inserts user report parameters title
'Parameters:
'	sTitle - string value
'
'Example:

If sReportGUID="{E2AF7EF1-29F0-A70B-7897-6B7814332992}" Then
	ReportRequestTitle ("<b>User defined report parameters title</b>")
	ReportRequestFormFieldInput "Document ID ", "ИСПР20031129/63", "XXX", "{EAB2C1BF-2676-E606-B671-7D7B051A5DC4}"
	ReportRequestTitle ("")
End If	
If sReportGUID="{3E828A0D-E7CD-5DDC-A977-FD6C4EE8638F}" Then
	ReportRequestTitle ("<b>Укажите логин пользователя, по которому делается отчет (Только логин, без Ф.И.О. и без угловых скобок)</b>")
	ReportRequestFormFieldInput "Логин пользователя", Session("UserID"), "XXX", ""
	ReportRequestTitle ("")
End If	
If sReportGUID="{9CBDB693-05AE-3609-8DC2-7328BB881B40}" Then
	ReportRequestTitle ("<b>Укажите период</b>")
	ReportRequestFormFieldInputDate "Начальная дата", MyDate(Date), "XXX", ""
	ReportRequestFormFieldInputDate "Конечная дата", MyDate(Date), "YYY", ""
	ReportRequestTitle ("")

	'ReportRequestFormFieldInputSelect "Example", "Name 1, Name 2", "Value 1, Value 2", "ZZZ"
	
End If	

If sReportGUID="{82B6B463-AF02-C317-DF48-864981B48F70}" Then 'Сводный отчет исполнительности подразделений
	ReportRequestTitle ("<b>Укажите подразделение, по которому делается отчет</b>")
	ReportRequestFormFieldInput "Подразделение", Session("Department"), "#DEP", "D"
	ReportRequestTitle ("<b>Укажите период</b>")
	dDateFrom=DateSerial(Year(Date), Month(Date), 1)
	nMonthNext=Month(Date)+1
	If nMonthNext>12 Then
		nMonthNext=1
	End If
	dDateTo=DateSerial(Year(Date), nMonthNext, 1)-1
	'dDateTo=dDateTo+CDate("23:59:59")
	ReportRequestFormFieldInputDate "Начальная дата", MyDate(dDateFrom), "#DATE1", ""
	ReportRequestFormFieldInputDate "Конечная дата", MyDate(dDateTo)+" 23:59:59", "#DATE2", ""
	ReportRequestTitle ("<b>Укажите категории документа, по которым делается отчет</b>")
	ReportRequestFormClassDocInput "", "#CLASSDOC"
	ReportRequestTitle ("")
End If	

If sReportGUID="{A4F06FFA-96E6-36FD-7D9D-1E542975549F}" Then 'Показатели качества выполнения поручений по подразделению
	'ReportRequestTitle ("<b>Укажите ЛОГИН исполнителя(без угловых скобок), по которому делается отчет</b>")
	'ReportRequestFormFieldInput "Исполнитель", Session("UserID"), "#RESP", "U"
	'ReportRequestTitle ("")
	ReportRequestTitle ("<b>Укажите подразделение, по которому делается отчет</b>")
	ReportRequestFormFieldInput "Подразделение", Session("Department"), "#DEP", "D"
	ReportRequestTitle ("<b>Укажите период</b>")
	dDateFrom=DateSerial(Year(Date), Month(Date), 1)
	nMonthNext=Month(Date)+1
	If nMonthNext>12 Then
		nMonthNext=1
	End If
	dDateTo=DateSerial(Year(Date), nMonthNext, 1)-1
	'dDateTo=dDateTo+CDate("23:59:59")
	ReportRequestFormFieldInputDate "Начальная дата", MyDate(dDateFrom), "#DATE1", ""
	ReportRequestFormFieldInputDate "Конечная дата", MyDate(dDateTo)+" 23:59:59", "#DATE2", ""
	ReportRequestTitle ("<b>Укажите категории документа, по которым делается отчет</b>")
	ReportRequestFormClassDocInput "", "#CLASSDOC"
	ReportRequestTitle ("")
	
End If	

If sReportGUID="{B81ED646-5350-3714-CEB0-443C1234B4B3}" Then 'Состояние хода работ по исполнителю (по документам на контроле)	ReportRequestTitle ("<b>Укажите ЛОГИН исполнителя(без угловых скобок), по которому делается отчет</b>")
	ReportRequestFormFieldInput "Исполнитель", Session("UserID"), "#RESP", "U"
	ReportRequestTitle ("<b>Укажите период</b>")
	dDateFrom=DateSerial(Year(Date), Month(Date), 1)
	nMonthNext=Month(Date)+1
	If nMonthNext>12 Then
		nMonthNext=1
	End If
	dDateTo=DateSerial(Year(Date), nMonthNext, 1)-1
	'dDateTo=dDateTo+CDate("23:59:59")
	ReportRequestFormFieldInputDate "Начальная дата", MyDate(dDateFrom), "#DATE1", ""
	ReportRequestFormFieldInputDate "Конечная дата", MyDate(dDateTo)+" 23:59:59", "#DATE2", ""
	ReportRequestTitle ("<b>Укажите категории документа, по которым делается отчет</b>")
	ReportRequestFormClassDocInput "", "#CLASSDOC"
	ReportRequestTitle ("")
End If	


' -------------------------------------------------------  Отчеты Ситроникс и СТС

' Сводка по контролируемым документам исполненным
If sReportGUID="{C4BFC7A8-6448-4182-A191-9FE84A5021D9}" Then
  ReportRequestFormFieldInputSelect SIT_ReportControl, SIT_ReportControlValues, "NameControl <> '',NameControl = ''", "#CONTROL"
  ReportRequestFormFieldInputSelect DOCS_ClassDoc, GetCategoriesList(), GetCategoriesList, "#CLASSDOC"
  ReportRequestFormFieldInputDate SIT_ReportStartingDate, "", "#DATE1", ""
  ReportRequestFormFieldInputDate SIT_ReportFinishingDate, "", "#DATE2", ""
  ReportRequestFormFieldInput DOCS_NameResponsible, GetFullName(SurnameGN(Session("Name")), Session("UserID")), "#RESPONSE", "U"
  ReportRequestFormFieldInput DOCS_DEPARTMENT, DelOtherLangFromFolder(Session("Department")), "#DEPARTMENT", "D"
End If

'02.11.2011 Отчет по действующим поручениям и поручениям, закрытым с нарушением срока
'02.11.2011 Отчет по нарушениям срока согласования
If sReportGUID="{9D1939A0-DFF3-43B4-9666-36543E381E23}" or sReportGUID="{D5146BB9-8C2F-434D-90B1-A6F18282624C}" Then
   ReportRequestFormFieldInput DOCS_NameResponsible, GetFullName(SurnameGN(DelOtherLangFromFolder(Session("Name"))),Session("UserID")), "#RESPONSIBLE_USER", "U"
  dDateTo=Date-1  
  If Month(dDateTo)< 4 Then
     dDateFrom=DateSerial(Year(dDateTo), 1, 1)
  Else
     If Month(Date)> 3 and Month(dDateTo)< 7 Then
        dDateFrom=DateSerial(Year(dDateTo), 4, 1)
     Else
        If Month(Date-1)> 6 and Month(dDateTo)< 10 Then
           dDateFrom=DateSerial(Year(dDateTo), 7, 1)
        Else
           If Month(Date-1)> 9 Then
              dDateFrom=DateSerial(Year(dDateTo), 10, 1)
           End If 
        End If
     End If
  End If
  
  ReportRequestFormFieldInputDate SIT_ReportTaskStartingDateFrom, MyDate(dDateFrom), "#DATE_BEGIN", ""
  ReportRequestFormFieldInputDate SIT_ReportTaskStartingDateTo, MyDate(dDateTo), "#DATE_END", ""
End If


' Справка по неисполненным срочным поручениям 
If sReportGUID="{B5E1B002-9AD2-49E2-8915-B6EF43B5602D}" Then
  ReportRequestFormFieldInputSelect SIT_ReportTypeOfTask, "<Любое>," + GetExtDirListOfValues("{C632B46B-3AAF-4607-BBC5-AC51C0A4971B}","Field1"), "," + GetExtDirListOfValues("{C632B46B-3AAF-4607-BBC5-AC51C0A4971B}","Field1"), "#TASK_TYPE"
  ReportRequestFormFieldInputSelect "Срочность", "<Любая>,Высокая,Средняя,Низкая", ",ВЫСОК,СРЕДН,НИЗК", "#PRIORITY"
  ReportRequestFormFieldInput SIT_Initiator, GetFullName(SurnameGN(Session("Name")), Session("UserID")), "#INITIAL_USER", "U"
  ReportRequestFormFieldInput DOCS_DEPARTMENT, DelOtherLangFromFolder(Session("Department")), "#DEPARTMENT", "D"
  ReportRequestFormFieldInput DOCS_NameResponsible, GetFullName(SurnameGN(Session("Name")), Session("UserID")), "#RESPONSIBLE_USER", "U"
  dDateFrom=DateSerial(Year(Date), Month(Date), 1)
  nMonthNext=Month(Date)+1
  If nMonthNext>12 Then
    nMonthNext=1
  End If
  dDateTo=DateSerial(Year(Date), nMonthNext, 1)-1
  ReportRequestFormFieldInputDate SIT_ReportTaskStartingDateFrom, MyDate(dDateFrom), "#DATE_BEGIN", ""
  ReportRequestFormFieldInputDate SIT_ReportTaskStartingDateTo, MyDate(dDateTo), "#DATE_END", ""
  ReportRequestFormFieldInputDate DOCS_DateCompletion, "", "#DATE_COMPLETION", ""
  ReportRequestFormFieldInputSelect SIT_ReportControl, SIT_ReportControlValues, "NameControl <> '',NameControl = ''", "#CONTROL"
End If

' Справка по неисполненным поручениям с нарушением срока 
If sReportGUID="{92416A95-B39C-486D-B508-67859FEB9222}" Then
  ReportRequestFormFieldInputSelect SIT_ReportTypeOfTask, "<Любое>," + GetExtDirListOfValues("{C632B46B-3AAF-4607-BBC5-AC51C0A4971B}","Field1"), "," + GetExtDirListOfValues("{C632B46B-3AAF-4607-BBC5-AC51C0A4971B}","Field1"), "#TASK_TYPE"
'  ReportRequestFormFieldInputSelect "Срочность", "<Любая>,Высокая,Средняя,Низкая", ",ВЫСОК,СРЕДН,НИЗК", "#PRIORITY"
'  ReportRequestFormFieldInput "Инициатор", GetFullName(SurnameGN(Session("Name")), Session("UserID")), "#INITIAL_USER", "U"
  ReportRequestFormFieldInput DOCS_DEPARTMENT, DelOtherLangFromFolder(Session("Department")), "#DEPARTMENT", "D"
  ReportRequestFormFieldInput DOCS_NameResponsible, GetFullName(SurnameGN(Session("Name")), Session("UserID")), "#RESPONSIBLE_USER", "U"
  dDateFrom=DateSerial(Year(Date), Month(Date), 1)
  nMonthNext=Month(Date)+1
  If nMonthNext>12 Then
    nMonthNext=1
  End If
  dDateTo=DateSerial(Year(Date), nMonthNext, 1)-1
  ReportRequestFormFieldInputDate SIT_ReportTaskStartingDateFrom, MyDate(dDateFrom), "#DATE_BEGIN", ""
  ReportRequestFormFieldInputDate SIT_ReportTaskStartingDateTo, MyDate(dDateTo), "#DATE_END", ""
'  ReportRequestFormFieldInputDate "Дата исполнения", "", "#DATE_COMPLETION", ""
  ReportRequestFormFieldInputSelect SIT_ReportControl, SIT_ReportControlValues, "NameControl <> '',NameControl = ''", "#CONTROL"
End If

' Сводка по срочным контрольным поручениям с нарушением срока 
If sReportGUID="{FC0BE1C4-2580-48DB-AEB1-BFDE94D4E3B5}" Then
'  ReportRequestFormFieldInputSelect "Вид поручения", GetExtDirListOfValues("{C632B46B-3AAF-4607-BBC5-AC51C0A4971B}","Field1"), GetExtDirListOfValues("{C632B46B-3AAF-4607-BBC5-AC51C0A4971B}","Field1"), "#TASK_TYPE"
'  ReportRequestFormFieldInputSelect "Срочность", "<Любая>,Высокая,Средняя,Низкая", ",ВЫСОК,СРЕДН,НИЗК", "#PRIORITY"
  ReportRequestFormFieldInput SIT_Initiator, GetFullName(SurnameGN(Session("Name")), Session("UserID")), "#INITIAL_USER", "U"
  ReportRequestFormFieldInput DOCS_DEPARTMENT, DelOtherLangFromFolder(Session("Department")), "#DEPARTMENT", "D"
  ReportRequestFormFieldInput DOCS_NameResponsible, GetFullName(SurnameGN(Session("Name")), Session("UserID")), "#RESPONSIBLE_USER", "U"
  dDateFrom=DateSerial(Year(Date), Month(Date), 1)
  nMonthNext=Month(Date)+1
  If nMonthNext>12 Then
    nMonthNext=1
  End If
  dDateTo=DateSerial(Year(Date), nMonthNext, 1)-1
  ReportRequestFormFieldInputDate SIT_ReportTaskStartingDateFrom, MyDate(dDateFrom), "#DATE_BEGIN", ""
  ReportRequestFormFieldInputDate SIT_ReportTaskStartingDateTo, MyDate(dDateTo), "#DATE_END", ""
  ReportRequestFormFieldInputSelect SIT_ReportTypeOfRequest, SIT_ReportRequestTypes, "StatusCompletion='1' And DateCompletion>=GetDate(),StatusCompletion='1' And DateCompletion<GetDate(),StatusCompletion<>'1' And DateCompletion>=GetDate(),StatusCompletion<>'1' And DateCompletion<GetDate()", "#REPORT_TYPE"
'  ReportRequestFormFieldInputDate "Дата исполнения", "", "#DATE_COMPLETION", ""
'  ReportRequestFormFieldInputSelect "Контроль", "На контроле, Без контроля", "NameControl <> '',NameControl = ''", "#CONTROL"
End If

'SAY 2008-12-19
'тест пиктограмм
If sReportGUID="{783B269D-2424-46BD-AEB2-46A0B58582FB}" or  sReportGUID="{EE5B34E5-75DD-434B-9040-6BB57D02B99D}" or  sReportGUID="{27E35FA5-08E6-455E-9B8E-323592AD59E3}" Then
  ReportRequestFormFieldInput DOCS_USER, GetFullName(SurnameGN(Session("Name")), Session("UserID")), "#RESPONSIBLE", "U"
  dDateFrom=DateSerial(Year(Date), Month(Date), 1)
  nMonthNext=Month(Date)+1
  nYearNext=0
  If nMonthNext>12 Then
    nMonthNext=1
    nYearNext=1
  End If
  dDateTo=DateSerial(Year(Date)+nYearNext, nMonthNext, 1)-1
  ReportRequestFormFieldInputDate SIT_ReportTaskStartingDateFrom, MyDate(dDateFrom), "#DATE_BEGIN", ""
  ReportRequestFormFieldInputDate SIT_ReportTaskStartingDateTo, MyDate(dDateTo), "#DATE_END", ""
End If

'Отчет по заявкам - Отчет 1 Операционная деятельность
If sReportGUID = "{C9FCFC11-4C43-4C84-B5F1-C9FD47F325E7}" Then

'rmanyushin@sitronics.com 15.07.2009, Start
  ReportRequestFormFieldInputSelect DOCS_ClassDoc, SIT_PurchaseOrder+","+SIT_PaymentOrder, "ClassDoc like N'%"+SIT_PurchaseOrder+"%',ClassDoc like N'%"+SIT_PaymentOrder+"%'", "#CLASSDOC"
'  ReportRequestFormFieldInputSelect DOCS_ClassDoc, DOCS_All+","+SIT_PurchaseOrder+","+SIT_PaymentOrder, "(ClassDoc like N'%"+SIT_PurchaseOrder+"%' or ClassDoc like N'%"+SIT_PaymentOrder+"%'),ClassDoc like N'%"+SIT_PurchaseOrder+"%',ClassDoc like N'%"+SIT_PaymentOrder+"%'", "#CLASSDOC"
'rmanyushin@sitronics.com 15.07.2009, Stop  
  ReportRequestFormFieldInput SIT_BusinessUnit, "", "#BUSINESSUNIT", "{8E24E3EF-F350-4D29-8BA0-430E425F54E0}"
  ReportRequestFormFieldInput SIT_CostCenterCode, "", "#COSTCENTER", "{33F9C053-E51D-4738-91CD-45ABB82C1D8A}"
'Ph - 20090403 - Убран параметр, сделали только по 0-м проектам
'  ReportRequestFormFieldInputSelect SIT_Project, DOCS_All+","+SIT_YesNo, ",and UserFieldText3 not like '[a-z0]00000',and UserFieldText3 like '[a-z0]00000'", "#ISPROJECT"
  ReportRequestFormFieldInput SIT_ExpenseItem, "", "#CHARTOFACCOUT", "{3A4F4557-A6E8-4382-A69F-59CF8895645F}"
  dDateFrom = DateSerial(Year(Date), Month(Date), 1)
  nMonthNext = Month(Date)+1
  nYearNext = 0
  If nMonthNext > 12 Then
    nMonthNext = 1
    nYearNext = 1
  End If
  dDateTo = DateSerial(Year(Date)+nYearNext, nMonthNext, 1)-1
  ReportRequestFormFieldInputDate DOCS_DateCompletion+" "+Trim(DOCS_FROM_Date), MyDate(dDateFrom), "#DATE_BEGIN", ""
  ReportRequestFormFieldInputDate DOCS_DateCompletion+" "+Trim(DOCS_TO_Date), MyDate(dDateTo), "#DATE_END", ""
  ReportRequestFormFieldInputSelect DOCS_APROVAL, DOCS_All+","+DOCS_Approved1+","+DOCS_ApprovedNot1, "and (NameApproved='' or NameApproved is Null or (NameApproved<>'' and NameApproved is Not Null and DateApproved >= #DATEAPPR1 and DateApproved <= #DATEAPPR2)),and NameApproved<>'' and NameApproved is Not Null and DateApproved >= #DATEAPPR1 and DateApproved <= #DATEAPPR2,and (NameApproved='' or NameApproved is Null)", "#APPROVED"
  ReportRequestFormFieldInputDate DOCS_DateApproval+" "+Trim(DOCS_FROM_Date), MyDate(dDateFrom), "#DATEAPPR1", ""
  ReportRequestFormFieldInputDate DOCS_DateApproval+" "+Trim(DOCS_TO_Date), MyDate(dDateTo), "#DATEAPPR2", ""
End If

'Отчет по заявкам - Отчет 2 Проектная деятельность
If sReportGUID = "{3CDF1CF5-9FF3-4637-91B6-3C1525AE5026}" Then
'rmanyushin@sitronics.com 15.07.2009, Start  
  ReportRequestFormFieldInputSelect DOCS_ClassDoc, SIT_PurchaseOrder+","+SIT_PaymentOrder, "ClassDoc like N'%"+SIT_PurchaseOrder+"%',ClassDoc like N'%"+SIT_PaymentOrder+"%'", "#CLASSDOC"
'  ReportRequestFormFieldInputSelect DOCS_ClassDoc, DOCS_All+","+SIT_PurchaseOrder+","+SIT_PaymentOrder, "(ClassDoc like N'%"+SIT_PurchaseOrder+"%' or ClassDoc like N'%"+SIT_PaymentOrder+"%'),ClassDoc like N'%"+SIT_PurchaseOrder+"%',ClassDoc like N'%"+SIT_PaymentOrder+"%'", "#CLASSDOC"
'rmanyushin@sitronics.com 15.07.2009, Stop
  ReportRequestFormFieldInput SIT_BusinessUnit, "", "#BUSINESSUNIT", "{8E24E3EF-F350-4D29-8BA0-430E425F54E0}"
  ReportRequestFormFieldInput SIT_CostCenterCode, "", "#COSTCENTER", "{33F9C053-E51D-4738-91CD-45ABB82C1D8A}"
  ReportRequestFormFieldInput SIT_Project, "", "#PROJECT", "{ACCDE453-D50A-48E0-9BFB-1BEA45D6D16E}"
  ReportRequestFormFieldInput SIT_ExpenseItem, "", "#CHARTOFACCOUT", "{3A4F4557-A6E8-4382-A69F-59CF8895645F}"
  dDateFrom = DateSerial(Year(Date), Month(Date), 1)
  nMonthNext = Month(Date)+1
  nYearNext = 0
  If nMonthNext > 12 Then
    nMonthNext = 1
    nYearNext = 1
  End If
  dDateTo = DateSerial(Year(Date)+nYearNext, nMonthNext, 1)-1
  ReportRequestFormFieldInputDate DOCS_DateCompletion+" "+Trim(DOCS_FROM_Date), MyDate(dDateFrom), "#DATE_BEGIN", ""
  ReportRequestFormFieldInputDate DOCS_DateCompletion+" "+Trim(DOCS_TO_Date), MyDate(dDateTo), "#DATE_END", ""
  ReportRequestFormFieldInputSelect DOCS_APROVAL, DOCS_All+","+DOCS_Approved1+","+DOCS_ApprovedNot1, "and (NameApproved='' or NameApproved is Null or (NameApproved<>'' and NameApproved is Not Null and DateApproved >= #DATEAPPR1 and DateApproved <= #DATEAPPR2)),and NameApproved<>'' and NameApproved is Not Null and DateApproved >= #DATEAPPR1 and DateApproved <= #DATEAPPR2,and (NameApproved='' or NameApproved is Null)", "#APPROVED"
  ReportRequestFormFieldInputDate DOCS_DateApproval+" "+Trim(DOCS_FROM_Date), MyDate(dDateFrom), "#DATEAPPR1", ""
  ReportRequestFormFieldInputDate DOCS_DateApproval+" "+Trim(DOCS_TO_Date), MyDate(dDateTo), "#DATEAPPR2", ""
End If

'Отчет по заявкам - Отчет 3 Заявки на закупку, согласованные/утвержденные фин.контролером
If sReportGUID = "{B2683E4E-EC38-44CB-BDBC-1BFE016F6687}" Then
  ReportRequestFormFieldInputSelect DOCS_ClassDoc, DOCS_All+","+SIT_PurchaseOrder+","+SIT_PaymentOrder, "(ClassDoc like N'%"+SIT_PurchaseOrder+"%' or ClassDoc like N'%"+SIT_PaymentOrder+"%'),ClassDoc like N'%"+SIT_PurchaseOrder+"%',ClassDoc like N'%"+SIT_PaymentOrder+"%'", "#CLASSDOC"
  ReportRequestFormFieldInput SIT_Project, "", "#PROJECT", "{ACCDE453-D50A-48E0-9BFB-1BEA45D6D16E}"
  ReportRequestFormFieldInputUserID Replace(Replace(Replace(STS_FinancialControl,"#",""),";",""),"""",""), "", "#CONTROLER"
  dDateFrom = DateSerial(Year(Date), Month(Date), 1)
  nMonthNext = Month(Date)+1
  nYearNext = 0
  If nMonthNext > 12 Then
    nMonthNext = 1
    nYearNext = 1
  End If
  dDateTo = DateSerial(Year(Date)+nYearNext, nMonthNext, 1)-1
  ReportRequestFormFieldInputDate SIT_ReportStartingDate, MyDate(dDateFrom), "#DATE_BEGIN", ""
  ReportRequestFormFieldInputDate SIT_ReportFinishingDate, MyDate(dDateTo), "#DATE_END", ""
End If

'Отчет по заявкам - Отчет 4 Отчет с произвольными настройками
If sReportGUID = "{821FF706-FDE5-4E52-B38B-CB9A3F00932D}" Then
  ReportRequestFormFieldInputSelect DOCS_ClassDoc, DOCS_All+","+SIT_PurchaseOrder+","+SIT_PaymentOrder, "(ClassDoc like N'%"+SIT_PurchaseOrder+"%' or ClassDoc like N'%"+SIT_PaymentOrder+"%'),ClassDoc like N'%"+SIT_PurchaseOrder+"%',ClassDoc like N'%"+SIT_PaymentOrder+"%'", "#CLASSDOC"
  ReportRequestFormFieldInput DOCS_DocID, "", "#DOCID", "!"
  ReportRequestFormFieldInputUserID SIT_Initiator, "", "#AUTHOR"
  ReportRequestFormFieldInputUserID DOCS_NameResponsible, "", "#NAMERESPONSIBLE"
  ReportRequestFormFieldInput DOCS_PartnerName, "", "#PARTNER", "P"
  ReportRequestFormFieldInput DOCS_Description, "", "#DESCRIPTION", "!"
  ReportRequestTitle ("<b>"+DOCS_AmountDoc+"</b>")
  ReportRequestFormFieldInput Trim(DOCS_FROM_Date), "0", "#AMOUNTDOC1", "!"
  ReportRequestFormFieldInput Trim(DOCS_TO_Date), "1000000000", "#AMOUNTDOC2", "!"
  ReportRequestFormFieldInput DOCS_Currency, "", "#CURRENCY", "E"
  ReportRequestTitle ("")
  ReportRequestTitle ("<b>"+SIT_USDAmount+"</b>")
  ReportRequestFormFieldInput Trim(DOCS_FROM_Date), "0", "#AMOUNTUSD1", "!"
  ReportRequestFormFieldInput DOCS_TO_Date, "1000000000", "#AMOUNTUSD2", "!"
  ReportRequestTitle ("")
  ReportRequestFormFieldInputUserID DOCS_ListToReconcile, "", "#LISTTORECONCILE"
  ReportRequestFormFieldInputUserID DOCS_NameAproval, "", "#NAMEAPROVAL"
  ReportRequestFormFieldInput DOCS_Department, "", "#DEPARTMENT", "D"
  ReportRequestFormFieldInput SIT_CostCenterCode, "", "#COSTCENTERCODE", "{33F9C053-E51D-4738-91CD-45ABB82C1D8A}"
  ReportRequestFormFieldInput SIT_BusinessUnit, "", "#BUSINESSUNIT", "{8E24E3EF-F350-4D29-8BA0-430E425F54E0}"
  ReportRequestFormFieldInput SIT_ProjectCode, "", "#PROJECTCODE", "{ACCDE453-D50A-48E0-9BFB-1BEA45D6D16E}"
  ReportRequestFormFieldInputSelect SIT_Budgeted, "<"+DOCS_All+">,"+SIT_YesNo, ",Yes,No", "#BUDGETED"
  ReportRequestFormFieldInputSelect SIT_PaymentType, "<"+DOCS_All+">,"+Replace(GetExtTableValues("PaymentTypes", "PaymentType_"+iif(Request("l") = "", "EN", iif(Request("l") = "3", "CZ", "RU"))), VbCrLf, ","), ","+Replace(GetExtTableValues("PaymentTypes", "PaymentType_EN"), VbCrLf, ","), "#PAYMENTTYPE"
  ReportRequestFormFieldInput SIT_ExpenseItem, "", "#CHARTOFACCOUT", "{3A4F4557-A6E8-4382-A69F-59CF8895645F}"
  ReportRequestTitle ("<b>"+DOCS_DateActivation+"</b>")
  'Так не хорошо, могут быть проблемы при преобразовании на разных языках
  'ReportRequestFormFieldInputDate SIT_ReportStartingDate, "01.01.2000", "#DATEACTIVATION1", ""
  'ReportRequestFormFieldInputDate SIT_ReportFinishingDate, "01.01.2100", "#DATEACTIVATION2", ""
  ReportRequestFormFieldInputDate SIT_ReportStartingDate, MyDate(DateSerial(2000, 1, 1)), "#DATEACTIVATION1", ""
  ReportRequestFormFieldInputDate SIT_ReportFinishingDate, MyDate(DateSerial(2100, 1, 1)), "#DATEACTIVATION2", ""
  ReportRequestTitle ("")
  ReportRequestTitle ("<b>"+DOCS_DateCompletion+"</b>")
'  ReportRequestFormFieldInputDate SIT_ReportStartingDate, "01.01.2000", "#DATECOMPLETION1", ""
'  ReportRequestFormFieldInputDate SIT_ReportFinishingDate, "01.01.2100", "#DATECOMPLETION2", ""
  ReportRequestFormFieldInputDate SIT_ReportStartingDate, MyDate(DateSerial(2000, 1, 1)), "#DATECOMPLETION1", ""
  ReportRequestFormFieldInputDate SIT_ReportFinishingDate, MyDate(DateSerial(2100, 1, 1)), "#DATECOMPLETION2", ""
  ReportRequestTitle ("")
  ReportRequestFormFieldInputSelect DOCS_OrderBy, DOCS_DocID+","+DOCS_Author+","+DOCS_NameResponsible+","+DOCS_PartnerName+","+DOCS_Description+","+DOCS_AmountDoc+","+DOCS_Currency+","+SIT_USDAmount+","+DOCS_ListToReconcile+","+DOCS_NameAproval+","+DOCS_Department+","+SIT_CostCenterCode+","+SIT_BusinessUnit+","+SIT_ProjectCode+","+SIT_Budgeted+","+SIT_PaymentType+","+SIT_ExpenseItem+","+DOCS_DateActivation+","+DOCS_DateCompletion, "DocID,Author,NameResponsible,PartnerName,Description,AmountDoc,Currency,UserFieldMoney1,ListToReconcile,NameAproval,Department,UserFieldText1,BusinessUnit,UserFieldText3,UserFieldText6,UserFieldText5,UserFieldText8,DateActivation,DateCompletion", "#SORTFIELD1"
  ReportRequestFormFieldInputSelect SIT_SortType, SIT_SortTypeAscending+","+SIT_SortTypeDescending, ",desc", "#SORTTYPE1"
End If

'Отчет Анализ активности пользователей за месяц
If sReportGUID = "{BBC97D49-B04E-49F3-B273-1C24B7DA0216}" Then
  ReportRequestFormFieldInputSelect DOCS_PERIOD_Month, DOCS_PERIOD_JAN+","+DOCS_PERIOD_FEB+","+DOCS_PERIOD_MAR+","+DOCS_PERIOD_APR+","+DOCS_PERIOD_MAY+","+DOCS_PERIOD_JUN+","+DOCS_PERIOD_JUL+","+DOCS_PERIOD_AUG+","+DOCS_PERIOD_SEP+","+DOCS_PERIOD_OCT+","+DOCS_PERIOD_NOV+","+DOCS_PERIOD_DEC, "01,02,03,04,05,06,07,08,09,10,11,12", "#MONTH"
  ReportRequestFormFieldInput DOCS_PERIOD_YEAR, MyCStr(Year(Date)), "#YEAR", "!"
  ReportRequestFormFieldHidden SIT_Language, iif(Request("l") = "", "EN", iif(Request("l") = "3", "CZ", "RU")), "#LANGUAGE"
End If

'Отчет по заданному пользователю и Перечень процессов, в которых участвует данный пользователь 
If sReportGUID = "{1DDD6183-B848-4213-B197-C11F38F1301A}" or sReportGUID = "{118D9DF4-9350-49C7-9CBF-4DC50235F843}" Then
'  ReportRequestFormFieldInput DOCS_USER, GetFullName(SurnameGN(DelOtherLangFromFolder(Session("Name"))), Session("UserID")), "#USER", "U"
  ReportRequestFormFieldInputUserID DOCS_USER, GetFullName(SurnameGN(DelOtherLangFromFolder(Session("Name"))), Session("UserID")), "#USERID"
End If

'20090810 - Запрос №4 из СТС - Отчеты
'TOP_20_longest_approvals_list
If sReportGUID = "{66B2774B-07A1-4CEF-895E-BAED8351D709}" Then
  'Параметр только для отображения
  ReportRequestFormFieldHidden "Request from", GetFullName(SurnameGN(DelOtherLangFromFolder(Session("Name"))), Session("UserID")), "#EMPTY#"
  ReportRequestFormFieldInputDate "Date from", MyDate(DateSerial(2000, 1, 1)), "#DATE_BEGIN", ""
  ReportRequestFormFieldInputDate "Date till", MyDate(DateSerial(2100, 1, 1)), "#DATE_END", ""
End If

'Report_XXX
'If sReportGUID = "{3A0FA388-AF8A-4C04-B359-335A8EC18194}" or sReportGUID = "{768D0BCE-2EB2-4D57-90DC-71FE3368C99A}" or sReportGUID = "{E571BBB1-A8BA-4820-9752-E775574BC0EF}" or sReportGUID = "{CE71BD6D-FDF3-4885-AC08-19D09FD5DD1E}" Then
'20091110 - Запрос №7 из СТС - добавлен отчет по договорам
'If sReportGUID = "{3A0FA388-AF8A-4C04-B359-335A8EC18194}" or sReportGUID = "{768D0BCE-2EB2-4D57-90DC-71FE3368C99A}" or sReportGUID = "{E571BBB1-A8BA-4820-9752-E775574BC0EF}" or sReportGUID = "{CE71BD6D-FDF3-4885-AC08-19D09FD5DD1E}" or sReportGUID = "{C16BE030-1EA3-4D13-99F2-00E63E510F3B}" Then
'20100720 - Запрос №11 из СТС - добавлен отчет по старым договорам
If sReportGUID = "{3A0FA388-AF8A-4C04-B359-335A8EC18194}" or sReportGUID = "{768D0BCE-2EB2-4D57-90DC-71FE3368C99A}" or sReportGUID = "{E571BBB1-A8BA-4820-9752-E775574BC0EF}" or sReportGUID = "{CE71BD6D-FDF3-4885-AC08-19D09FD5DD1E}" or sReportGUID = "{C16BE030-1EA3-4D13-99F2-00E63E510F3B}" or sReportGUID = "{92915495-CCD5-4ADF-9D3B-2CA4771DF250}" Then
  'Параметр только для отображения
  ReportRequestFormFieldHidden "Request from", GetFullName(SurnameGN(DelOtherLangFromFolder(Session("Name"))), Session("UserID")), "#EMPTY#"
  ReportRequestFormFieldInputDate "Date from", MyDate(DateSerial(2000, 1, 1)), "#DATE_BEGIN", ""
  ReportRequestFormFieldInputDate "Date till", MyDate(DateSerial(2100, 1, 1)), "#DATE_END", ""
End If

'rmanyushin 93489 21.04.2010
'Documents in the Paydox Archive
If sReportGUID = "{3D00D517-6264-4A36-A4E5-2AB33FD5EB77}" Then
  ReportRequestFormFieldHidden "Request from", GetFullName(SurnameGN(DelOtherLangFromFolder(Session("Name"))), Session("UserID")), "#EMPTY#"
  ReportRequestFormFieldInputSelect DOCS_ClassDoc, GetCategoriesList(), GetCategoriesList2, "#CLASSDOC"
  ReportRequestFormFieldInputDate "Date from", MyDate("01.09.2008"), "#DATE_BEGIN", ""
  ReportRequestFormFieldInputDate "Date till", MyDate("01.04.2009"), "#DATE_END", ""
End If
'rmanyushin 93489 21.04.2010

'rmanyushin 119429 18.08.2010 Start
If sReportGUID = "{9007AB69-CB5C-48FD-AE67-8C7F53EB2BA1}" Then
 ReportRequestFormFieldInputDate "Date from", MyDate(DateSerial(2000, 1, 1)), "#DATE_BEGIN", ""
 ReportRequestFormFieldInputDate "Date till", MyDate(DateSerial(2100, 1, 1)), "#DATE_END", ""
End If
'rmanyushin 119429 18.08.2010 End

'rmanyushin 133351 05.10.2010 Start
If sReportGUID = "{61DF768C-B20E-46D7-92F9-82716D52AFEC}" Then
 ReportRequestFormFieldInputDate "Date from", MyDate(DateSerial(2009, 1, 1)), "#DATE_BEGIN", ""
 ReportRequestFormFieldInputDate "Date till", MyDate(DateSerial(2010, 1, 1)), "#DATE_END", ""
End If
'rmanyushin 133351 05.10.2010 End

'Запрос №23 - СТС - start
'Параметр запроса лога использования правил
If sReportGUID = "{DE37E6E8-58E7-468E-A635-1F3BE63BBE9C}" Then
  ReportRequestFormFieldInput DOCS_DocID, "", "#DOCID#", "!"
End If
'Запрос №23 - СТС - end

'Запрос №41 - СТС - start
If sReportGUID = "{1812959E-14E5-40CE-9ECA-F9218EEC6DE0}" Then
  ReportRequestFormFieldHidden "Company", GetRootDepartmentFull(Session("Department")), "#DEPARTMENT#" 'Отчет строится по БН вызывающего пользователя
  GetCategoriesListForReport sSelectList, sReportList
  'включаем оператор and в значения из списка для вставки в отчет (чтобы учесть отсутствие условия по категории)
  sReportList = Replace(sReportList, "ClassDoc = ", "and ClassDoc = ")
  ReportRequestFormFieldInputSelect DOCS_ClassDoc, DOCS_All & "," & sSelectList, "," & sReportList, "#CLASSDOC#"
  ReportRequestFormFieldInputDate "Date from", MyDate(Date()-7), "#DATE_BEGIN#", ""
  ReportRequestFormFieldInputDate "Date till", MyDate(Date()), "#DATE_END#", ""
End If
'Запрос №41 - СТС - end

%>
