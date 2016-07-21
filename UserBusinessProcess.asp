<%
'Set user business processes functions
Function GetConnection()
  Set GetConnection = Conn
End Function

'Получить букву в названии ячейки по ее номеру
Function GetCellXName(parNo)
  Dim iNo

  GetCellXName = ""
  If parNo <= 0 Then
    Exit Function
  End If

  iNo = parNo
  Do While iNo > 0
    GetCellXName = Chr(Asc("A")-1+iNo Mod 26)+GetCellXName
    iNo = iNo \ 26
  Loop
End Function

Sub MyInsertRangeText(parSheet, Text, sRangeName)
  parSheet.Range(sRangeName).Value=Text
End Sub 

Function Col2(par)
  Col2 = "123"
End Function

Function Col3(par)
  Col3 = "456"
End Function

Function Col4(par)
  Col4 = "789"
End Function

'Вырезать английское название из трехъязычного
Function GetEngNameFromFolder(ByVal parStr)
  VAR_CurrentL = ""
  objTreeView.VAR_CurrentL=VAR_CurrentL
  GetEngNameFromFolder = objTreeView.DelOtherLangFromFolder(parStr)
  VAR_CurrentL = "-"
  objTreeView.VAR_CurrentL=VAR_CurrentL
End Function

'Получить конечное звено в иерархии подразделений на англ. языке
Function MyGetShortDepartment(parDepartment)
  MyGetShortDepartment = ""
  If Trim(parDepartment) = "" Then
    Exit Function
  End If
  MyGetShortDepartment = GetEngNameFromFolder(Trim(parDepartment))
  If MyGetShortDepartment = "" Then
    Exit Function
  End If
  If Right(MyGetShortDepartment, 1) = "/" Then
    MyGetShortDepartment = Left(MyGetShortDepartment, Len(MyGetShortDepartment)-1)
  End If
  iPos = InStrRev(MyGetShortDepartment,"/")
  If iPos <> 0 Then
    MyGetShortDepartment = Mid(MyGetShortDepartment, iPos+1, Len(MyGetShortDepartment)-iPos)
  End If
End Function

Function LeadSymbolNVal(cPar, symbol, N)
  cPar=Trim(MyCStr(cPar))
  LeadSymbolNVal = IIf(Len(cPar) < N, String(N - Len(cPar), symbol) + cPar, cPar)
End Function

'Форматирование даты в формат DD.MM.YYYY
Function DateInDDMMYYYY(ByVal parDate)
  If IsDate(parDate) Then
    DateInDDMMYYYY = LeadSymbolNVal(CStr(Day(parDate)),"0",2)+"."+LeadSymbolNVal(CStr(Month(parDate)),"0",2)+"."+CStr(Year(parDate))
  Else
    DateInDDMMYYYY = ""
  End If
End Function

Function GetFullName(ByVal cParName, ByVal cParID)
  GetFullName = """" & MyCStr(cParName) & """ <" & MyCStr(cParID) & ">"
End Function 

Function IsAdmin()
  If Session("WriteSecurityLevel")>=VAR_AdminSecLevel Or Session("ReadSecurityLevel")>=VAR_AdminSecLevel Then
    IsAdmin=True
  Else
    IsAdmin=False
  End If
End Function

Function UniDate(parDate)
  If IsNull(parDate) Then
    UniDate = "Not a date"
    Exit Function
  End If
  If not IsDate(parDate) Then
    UniDate = "Not a date"
    Exit Function
  End If
  UniDate = "{ d '"+CStr(Year(parDate))+"-"+LeadSymbolNVal(CStr(Month(parDate)),"0",2)+"-"+LeadSymbolNVal(CStr(Day(parDate)),"0",2)+"' }"
End Function

'Получить название корневого подразделения (на всех языках)
Function GetRootDepartment(ByVal sDepartment)
  Dim iPos
  
  GetRootDepartment = Trim(sDepartment)
  If GetRootDepartment = "" Then
    Exit Function
  End If
  iPos = InStr(GetRootDepartment, "/")
  If iPos > 0 Then
    GetRootDepartment = Left(GetRootDepartment, iPos-1)
  End If
End Function

'Вырезать название из трехъязычного
Function DelOtherLangFromFolder(ByVal parStr)
  DelOtherLangFromFolder = objTreeView.DelOtherLangFromFolder(parStr)
End Function

'Посчитать число рабочих дней между двумя датами
Function WorkingDaysBetweenTwoDates(DateFrom, DateTill)
  FullWeeks = DateDiff("d", DateFrom, DateTill) \ 7
  AddWeekDays = DateDiff("d", DateFrom, DateTill) mod 7

  DayFrom = Weekday(DateFrom, 2) '1 - понедельник
  DayTill = Weekday(DateTill, 2)
  If DayTill >= DayFrom Then
    Correction = 0
  Else
    If DayFrom = 7 Then
      Correction = -1
    Else
      Correction = -2
    End If
  End If
  WorkingDaysBetweenTwoDates = FullWeeks*5 + AddWeekDays + Correction
End Function

'Получить время первой активации документа
Function GetFirstActivationDate(parDocID, parDateActivation, parActivated)
  Dim sSQL, dsTemp

  sSQL = "select top 1 DateCreation from Comments where DocID = N'"+Replace(parDocID, "'", "''")+"' and CommentType = 'system' and (SpecialInfo = 'DOCS_Active' or CharIndex(N'Dokument je aktivní', Comment) = 1 or CharIndex(N'Active', Comment) = 1 or CharIndex(N'Активен', Comment) = 1 or CharIndex(N'Document active', Comment) = 1 or CharIndex(N'Документ активен', Comment) = 1) order by DateCreation"
'BP.AddLog "GetFirstActivationDate SQL: "+sSQL
  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  dsTemp.Open sSQL, Conn, 3, 1, &H1
  GetFirstActivationDate = iif(dsTemp.EOF, parDateActivation, dsTemp("DateCreation"))
  dsTemp.Close
End Function

Function IsLogin(parStr)
  Dim iLBr

  iLBr = InStr(parStr, "<")
  IsLogin = iLBr > 0 and iLBr < InStr(parStr, ">")
End Function

'Определить число уровней согласования/утверждения
Function GetReconciliationLevels(parListToReconcile, parNameAproval)
  Dim i, arLevels
  
  GetReconciliationLevels = 0
  If Trim(MyCStr(parListToReconcile)) <> "" Then
    arLevels = Split(parListToReconcile, VbCrLf)
    For i = 0 To UBound(arLevels)
      If Trim(arLevels(i)) <> "" and IsLogin(arLevels(i)) Then
        GetReconciliationLevels = GetReconciliationLevels+1
      End If
    Next
  End If

  If Trim(MyCStr(parNameAproval)) <> "" Then
    If IsLogin(parNameAproval) Then
      GetReconciliationLevels = GetReconciliationLevels+1
    End If
  End If
End Function

'Получить задержку при визировании документа (в рабочих днях)
Function GetDelayInAproval(parDocID, parDateActivation, parDateApproved, parNameAproval, parListToReconcile, parActivated)
  Dim dDateActivation, dEndDate, iDelay

'  dDateActivation = GetFirstActivationDate(parDocID, parDateActivation, parActivated)
  dDateActivation = iif(IsNull(parActivated), parDateActivation, parActivated)
  dEndDate = iif(IsNull(parDateApproved) or parDateApproved = VAR_BeginOfTimes, Date(), parDateApproved)
  iDelay = WorkingDaysBetweenTwoDates(dDateActivation, dEndDate) - GetReconciliationLevels(parListToReconcile, parNameAproval)*3
  GetDelayInAproval = iif(iDelay > 0, iDelay, 0)
End Function

'Получить задержку в исполнении документа (в рабочих днях)
Function GetDelayInCompletion(parDateCompletion, parDateCompleted, parStatusCompletion)
  Dim dDateCompleted
  
  dDateCompleted = iif(MyCStr(parStatusCompletion)<>VAR_StatusCompletion And MyCStr(parStatusCompletion)<>VAR_StatusCancelled And IsNull(parDateCompleted), Date(), parDateCompleted)
  GetDelayInCompletion = iif(dDateCompleted > parDateCompletion, WorkingDaysBetweenTwoDates(parDateCompletion, dDateCompleted), 0)
End Function

'Получить логин пользователя из поля - SQL операторы
Function SQL_GetUserID(parFieldName)
  SQL_GetUserID = " CASE WHEN "+parFieldName+" is not NULL and CharIndex('<', "+parFieldName+") > 0 and CharIndex('>', "+parFieldName+") > CharIndex('<', "+parFieldName+") THEN SubString("+parFieldName+", CharIndex('<', "+parFieldName+")+1, CharIndex('>', "+parFieldName+")-CharIndex('<', "+parFieldName+")-1) ELSE '' END "
End Function

'Показать в отчете High_level_report строчку со статистикой по указанному подразделению (включая вложенные подразделения)
Sub ShowDepartmentStatistics(parDepartmentCondition, parDepartmentToShow, parSheet, parX, parY, dDateFrom, dDateTill, bAverage)
  Dim sSQL, dsTemp, iDelay
  Dim i, sCode, sName
  Dim iValue
  'Счетчики
  Dim iPOs, iPOsDelayed, iPODelayTime 'Заявки на закупку
  Dim iPs 'Заявки на оплату
  Dim iCEOs, iCEOsDelayed, iCEODelayTime 'Распорядительные
  Dim iTasks, iTasksDelayed, iTaskDelayTime 'Поручения
  Dim iContracts, iContractsDelayed, iContractDelayTime 'Договоры
  Dim RootDepartment
  Dim sPreviousDocID

  iPOs = 0
  iPOsDelayed = 0
  iPODelayTime = 0
  iPs = 0
  iCEOs = 0
  iCEOsDelayed = 0
  iCEODelayTime = 0
  iTasks = 0
  iTasksDelayed = 0
  iTaskDelayTime = 0
  iContracts = 0
  iContractsDelayed = 0
  iContractDelayTime = 0
  
  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  dsTemp.CursorLocation = 3
'  sSQL = "select DocID, DateActivation, DateApproved, NameAproval, ListToReconcile, DateCompletion, DateCompleted, StatusCompletion, Docs.ClassDoc"
  sSQL = "select Comments.DateCreation as Activated, Docs.DocID, DateActivation, DateApproved, NameAproval, ListToReconcile, DateCompletion, DateCompleted, StatusCompletion, Docs.ClassDoc"
  sSQL = sSQL + ", Users.UserID, Users.Department as UserDepartment from Docs left join Users on (UserID = Case When CharIndex(N'Поручения', Docs.ClassDoc) = 1 or CharIndex(N'Договоры', Docs.ClassDoc) = 1 Then "+SQL_GetUserID("NameResponsible")+" Else "+SQL_GetUserID("Author")+" End) "
  sSQL = sSQL + " left join Comments on (Docs.DocID = Comments.DocID and CommentType = 'system' and (SpecialInfo = 'DOCS_Active' or CharIndex(N'Dokument je aktivní', Comments.Comment) = 1 or CharIndex(N'Active', Comments.Comment) = 1 or CharIndex(N'Активен', Comments.Comment) = 1 or CharIndex(N'Document active', Comments.Comment) = 1 or CharIndex(N'Документ активен', Comments.Comment) = 1)) "
  sSQL = sSQL + " where IsActive <> 'N' "
  sSQL = sSQL + " and (CharIndex(N'Заявка на закупку', Docs.ClassDoc) = 1 or CharIndex(N'Заявка на оплату', Docs.ClassDoc) = 1 or CharIndex(N'Распорядительные документы', Docs.ClassDoc) = 1 or CharIndex(N'Поручения', Docs.ClassDoc) = 1 or CharIndex(N'Договоры', Docs.ClassDoc) = 1) "
'ph 20100323 - start - Смена условия по датам
  sSQL = sSQL + " and DateActivation >= "+UniDate(dDateFrom)+" and DateActivation <= "+UniDate(dDateTill)+"+1 "
'  sSQL = sSQL + " and DateActivation >= "+UniDate(dDateFrom)+" and ((StatusCompletion is not NULL and StatusCompletion <> '0' and StatusCompletion <> '1') or (DateApproved is NULL and IsNull(StatusCompletion, '') <> '0') or DateActivation <= "+UniDate(dDateTill)+"+1) "
'ph 20100323 - end
'  sSQL = sSQL + " and CharIndex(N'"+Replace(parDepartmentCondition, "'", "''")+"', Users.Department) = 1 "
  sSQL = sSQL + " and "+parDepartmentCondition
  sSQL = sSQL + " and CharIndex(N'СТР*STS*/', Docs.Department) = 1 " 'Отсекаем документы, созданные не в СТР
  sSQL = sSQL + " order by Docs.DocID, Comments.DateCreation"

BP.AddLog "ShowDepartmentStatistics - sSQL: " & sSQL
  dsTemp.Open sSQL, Conn, 3, 1, &H1
  dsTemp.ActiveConnection = Nothing

  sPreviousDocID = ""
  Do While not dsTemp.EOF
    If sPreviousDocID <> MyCStr(dsTemp("DocID")) Then
      sPreviousDocID = MyCStr(dsTemp("DocID"))
      'If InStr(dsTemp("ClassDoc"), "Заявка на закупку") = 1 Then
	  If InStr(dsTemp("ClassDoc"), "Заявка на закупку") = 1 and (dsTemp("StatusCompletion") <> "0" or IsNull(dsTemp("StatusCompletion"))) Then 'rmanyushin 88625 31.03.2010
        iPOs = iPOs+1
        iDelay = GetDelayInAproval(dsTemp("DocID"), dsTemp("DateActivation"), dsTemp("DateApproved"), dsTemp("NameAproval"), dsTemp("ListToReconcile"), dsTemp("Activated"))
        iPODelayTime = iPODelayTime+iDelay
        If iDelay > 0 Then
          iPOsDelayed = iPOsDelayed+1
        End If
      'ElseIf InStr(dsTemp("ClassDoc"), "Заявка на оплату") = 1 Then
	  ElseIf InStr(dsTemp("ClassDoc"), "Заявка на оплату") = 1 and (dsTemp("StatusCompletion") <> "0" or IsNull(dsTemp("StatusCompletion"))) Then 'rmanyushin 88625 31.03.2010 
        iPs = iPs+1
      'ElseIf InStr(dsTemp("ClassDoc"), "Распорядительные документы") = 1 Then
      ElseIf InStr(dsTemp("ClassDoc"), "Распорядительные документы") = 1 and (dsTemp("StatusCompletion") <> "0" or IsNull(dsTemp("StatusCompletion"))) Then 'rmanyushin 88625 31.03.2010  
        iCEOs = iCEOs+1
        iDelay = GetDelayInAproval(dsTemp("DocID"), dsTemp("DateActivation"), dsTemp("DateApproved"), dsTemp("NameAproval"), dsTemp("ListToReconcile"), dsTemp("Activated"))
        iCEODelayTime = iCEODelayTime+iDelay
        If iDelay > 0 Then
          iCEOsDelayed = iCEOsDelayed+1
        End If
      'ElseIf InStr(dsTemp("ClassDoc"), "Поручения") = 1 Then
	  ElseIf InStr(dsTemp("ClassDoc"), "Поручения") = 1 and (dsTemp("StatusCompletion") <> "0" or IsNull(dsTemp("StatusCompletion"))) Then 'rmanyushin 88625 31.03.2010  
        iTasks = iTasks+1
        iDelay = GetDelayInCompletion(dsTemp("DateCompletion"), dsTemp("DateCompleted"), dsTemp("StatusCompletion"))
        iTaskDelayTime = iTaskDelayTime+iDelay
        If iDelay > 0 Then
          iTasksDelayed = iTasksDelayed+1
        End If
      'ElseIf InStr(dsTemp("ClassDoc"), "Договоры") = 1 Then
	  ElseIf InStr(dsTemp("ClassDoc"), "Договоры") = 1 and (dsTemp("StatusCompletion") <> "0" or IsNull(dsTemp("StatusCompletion"))) Then 'rmanyushin 88625 31.03.2010  
        iContracts = iContracts+1
        iDelay = GetDelayInAproval(dsTemp("DocID"), dsTemp("DateActivation"), dsTemp("DateApproved"), dsTemp("NameAproval"), dsTemp("ListToReconcile"), dsTemp("Activated"))
        iContractDelayTime = iContractDelayTime+iDelay
        If iDelay > 0 Then
          iContractsDelayed = iContractsDelayed+1
        End If
      End If
    End If

    dsTemp.MoveNext
  Loop
  dsTemp.Close
  Set dsTemp = Nothing
  
  'Вывод данных в форму
  'Вывод информации о подразделениях по которым строка
  sCode = ""
  sName = MyGetShortDepartment(parDepartmentToShow)
  i = InStr(sName, " ")
  If i <> 0 Then
    sCode = Left(sName, i-1)
    sName = Mid(sName, i+1)
  Else
    sCode = "???"
  End If

  parSheet.Range(GetCellXName(parX)+CStr(parY)).Value = sName
  parSheet.Range(GetCellXName(parX+1)+CStr(parY)).Value = sCode

  parSheet.Range(GetCellXName(parX+2)+CStr(parY)).Value = CStr(iPOs)
  parSheet.Range(GetCellXName(parX+5)+CStr(parY)).Value = CStr(iPs)
  parSheet.Range(GetCellXName(parX+6)+CStr(parY)).Value = CStr(iCEOs)
  parSheet.Range(GetCellXName(parX+9)+CStr(parY)).Value = CStr(iTasks)
  parSheet.Range(GetCellXName(parX+12)+CStr(parY)).Value = CStr(iContracts)

  If bAverage Then
    If iPOs > 0 Then
      iValue = 100*(iPOs-iPOsDelayed)\iPOs
      parSheet.Range(GetCellXName(parX+3)+CStr(parY)).Value = CStr(iValue)+"%"
      parSheet.Range(GetCellXName(parX+3)+CStr(parY)).Interior.Color = PercentColor(iValue)
      iValue = iPODelayTime \ iif(iPOsDelayed = 0, 1, iPOsDelayed)
      parSheet.Range(GetCellXName(parX+4)+CStr(parY)).Value = CStr(iValue)
      parSheet.Range(GetCellXName(parX+4)+CStr(parY)).Interior.Color = WDColor(iValue)
    Else
      parSheet.Range(GetCellXName(parX+3)+CStr(parY)).Value = "N/A"
      parSheet.Range(GetCellXName(parX+4)+CStr(parY)).Value = "N/A"
    End If
    If iCEOs > 0 Then
      iValue = 100*(iCEOs-iCEOsDelayed)\iCEOs
      parSheet.Range(GetCellXName(parX+7)+CStr(parY)).Value = CStr(iValue)+"%"
      parSheet.Range(GetCellXName(parX+7)+CStr(parY)).Interior.Color = PercentColor(iValue)
      iValue = iCEODelayTime \ iif(iCEOsDelayed = 0, 1, iCEOsDelayed)
      parSheet.Range(GetCellXName(parX+8)+CStr(parY)).Value = CStr(iValue)
      parSheet.Range(GetCellXName(parX+8)+CStr(parY)).Interior.Color = WDColor(iValue)
    Else
      parSheet.Range(GetCellXName(parX+7)+CStr(parY)).Value = "N/A"
      parSheet.Range(GetCellXName(parX+8)+CStr(parY)).Value = "N/A"
    End If
    If iTasks > 0 Then
      iValue = 100*(iTasks-iTasksDelayed)\iTasks
      parSheet.Range(GetCellXName(parX+10)+CStr(parY)).Value = CStr(iValue)+"%"
      parSheet.Range(GetCellXName(parX+10)+CStr(parY)).Interior.Color = PercentColor(iValue)
      iValue = iTaskDelayTime \ iif(iTasksDelayed = 0, 1, iTasksDelayed)
      parSheet.Range(GetCellXName(parX+11)+CStr(parY)).Value = CStr(iValue)
      parSheet.Range(GetCellXName(parX+11)+CStr(parY)).Interior.Color = WDColor(iValue)
    Else
      parSheet.Range(GetCellXName(parX+10)+CStr(parY)).Value = "N/A"
      parSheet.Range(GetCellXName(parX+11)+CStr(parY)).Value = "N/A"
    End If
    If iContracts > 0 Then
      iValue = 100*(iContracts-iContractsDelayed)\iContracts
      parSheet.Range(GetCellXName(parX+13)+CStr(parY)).Value = CStr(iValue)+"%"
      parSheet.Range(GetCellXName(parX+13)+CStr(parY)).Interior.Color = PercentColor(iValue)
      iValue = iContractDelayTime \ iif(iContractsDelayed = 0, 1, iContractsDelayed)
      parSheet.Range(GetCellXName(parX+14)+CStr(parY)).Value = CStr(iValue)
      parSheet.Range(GetCellXName(parX+14)+CStr(parY)).Interior.Color = WDColor(iValue)
    Else
      parSheet.Range(GetCellXName(parX+13)+CStr(parY)).Value = "N/A"
      parSheet.Range(GetCellXName(parX+14)+CStr(parY)).Value = "N/A"
    End If
  Else
    parSheet.Range(GetCellXName(parX+3)+CStr(parY)).Value = CStr(iPOs-iPOsDelayed)
	'parSheet.Range(GetCellXName(parX+4)+CStr(parY)).Value = CStr(iPODelayTime \ iif(iPOsDelayed = 0, 1, iPOsDelayed))
	'rmanyushin 88625 31.03.2010 Start 
	parSheet.Range(GetCellXName(parX+4)+CStr(parY)).NumberFormat = "0.00"
	parSheet.Range(GetCellXName(parX+4)+CStr(parY)).Value = CDbl(iPODelayTime / iif(iPOsDelayed = 0, 1, iPOsDelayed))
	'rmanyushin 88625 31.03.2010 End
	
    parSheet.Range(GetCellXName(parX+7)+CStr(parY)).Value = CStr(iCEOs-iCEOsDelayed)
    'parSheet.Range(GetCellXName(parX+8)+CStr(parY)).Value = CStr(iCEODelayTime \ iif(iCEOsDelayed = 0, 1, iCEOsDelayed))
	'rmanyushin 88625 31.03.2010 Start 
	parSheet.Range(GetCellXName(parX+8)+CStr(parY)).NumberFormat = "0.00"
	parSheet.Range(GetCellXName(parX+8)+CStr(parY)).Value = CDbl(iCEODelayTime / iif(iCEOsDelayed = 0, 1, iCEOsDelayed))
	'rmanyushin 88625 31.03.2010 End
	
	parSheet.Range(GetCellXName(parX+10)+CStr(parY)).Value = CStr(iTasks-iTasksDelayed)
	'parSheet.Range(GetCellXName(parX+11)+CStr(parY)).Value = CStr(iTaskDelayTime \ iif(iTasksDelayed = 0, 1, iTasksDelayed))
	'rmanyushin 88625 31.03.2010 Start 
	parSheet.Range(GetCellXName(parX+11)+CStr(parY)).NumberFormat = "0.00"
	parSheet.Range(GetCellXName(parX+11)+CStr(parY)).Value = CDbl(iTaskDelayTime / iif(iTasksDelayed = 0, 1, iTasksDelayed))
	'rmanyushin 88625 31.03.2010 End
	
	parSheet.Range(GetCellXName(parX+13)+CStr(parY)).Value = CStr(iContracts-iContractsDelayed)
    'parSheet.Range(GetCellXName(parX+14)+CStr(parY)).Value = CStr(iContractDelayTime \ iif(iContractsDelayed = 0, 1, iContractsDelayed))
	'rmanyushin 88625 31.03.2010 Start
	parSheet.Range(GetCellXName(parX+14)+CStr(parY)).NumberFormat = "0.00"
	parSheet.Range(GetCellXName(parX+14)+CStr(parY)).Value = CDbl(iContractDelayTime / iif(iContractsDelayed = 0, 1, iContractsDelayed))
	'rmanyushin 88625 31.03.2010 End
  End If
End Sub

'Получить цвет ячейки в зависимости от числа рабочих дней
Function WDColor(parWD)
  If parWD >= 11 Then
    WDColor = RGB(255,0,0)
  ElseIf parWD >= 6 Then
    WDColor = RGB(255,255,0)
  Else
    WDColor = RGB(153,204,0)
  End If
End Function

'Получить цвет ячейки в зависимости от процента
Function PercentColor(parPercent)
  If parPercent >= 90 Then
    PercentColor = RGB(153,204,0)
  ElseIf parPercent >= 75 Then
    PercentColor = RGB(255,255,0)
  Else
    PercentColor = RGB(255,0,0)
  End If
End Function

Sub ShowLegend(parSheet, parX, parY)
  parSheet.Range(GetCellXName(parX)+CStr(parY)).Value = "Legend"
  parSheet.Range(GetCellXName(parX)+CStr(parY+1)).Value = "On time rate"
  parSheet.Range(GetCellXName(parX)+CStr(parY+2)).Value = "Average delay"
  parSheet.Range(GetCellXName(parX)+CStr(parY)+":"+GetCellXName(parX)+CStr(parY+2)).Font.Bold = True
  parSheet.Range(GetCellXName(parX)+CStr(parY)).Font.Italic = True
  parSheet.Range(GetCellXName(parX+1)+CStr(parY+1)).Value = "0-74%"
  parSheet.Range(GetCellXName(parX+1)+CStr(parY+2)).Value = "> 11 WD"
  parSheet.Range(GetCellXName(parX+1)+CStr(parY+1)+":"+GetCellXName(parX+1)+CStr(parY+2)).Interior.Color = RGB(255,0,0)
  parSheet.Range(GetCellXName(parX+2)+CStr(parY+1)).Value = "75-89%"
  parSheet.Range(GetCellXName(parX+2)+CStr(parY+2)).Value = "6-10 WD"
  parSheet.Range(GetCellXName(parX+2)+CStr(parY+1)+":"+GetCellXName(parX+2)+CStr(parY+2)).Interior.Color = RGB(255,255,0)
  parSheet.Range(GetCellXName(parX+3)+CStr(parY+1)).Value = "90-100%"
  parSheet.Range(GetCellXName(parX+3)+CStr(parY+2)).Value = "0-5 WD"
  parSheet.Range(GetCellXName(parX+3)+CStr(parY+1)+":"+GetCellXName(parX+3)+CStr(parY+2)).Interior.Color = RGB(153,204,0)
  parSheet.Range(GetCellXName(parX+1)+CStr(parY+1)+":"+GetCellXName(parX+3)+CStr(parY+2)).HorizontalAlignment = -4108
  parSheet.Range(GetCellXName(parX)+CStr(parY+1)+":"+GetCellXName(parX+3)+CStr(parY+2)).Borders.LineStyle = 1
  parSheet.Range(GetCellXName(parX)+CStr(parY+1)+":"+GetCellXName(parX+3)+CStr(parY+2)).Borders.Weight = 2
End Sub

Function MyDate(parDate)
  MyDate = objTreeView.MyDate(parDate)
End Function

%>
