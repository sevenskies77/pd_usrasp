<%

'on error resume next
'dim i
'For Each i in Session.Contents
'  Response.Write(i & ": " & Cstr(Session.Contents(i)) & "<br />")
'Next

Function NotNegative(parNum)
   NotNegative = parNum
   If IsNumeric(NotNegative) Then
      If NotNegative < 0 Then
         NotNegative = 0
	  End If
   End If
End Function

Sub GetDocField_test(sDocID1)

  NewConnection = True
  If not IsNull(Conn) Then
     NewConnection = not IsObject(Conn)
  End If

  If NewConnection Then
     SetPayDoxPars 'Присваиваем переменные среды из Global.asa
     sConnStr = "ConnectString"
     Select Case UCase(Request("l"))
        Case "RU" sConnStr = sConnStr + "RUS"
        Case "3" sConnStr = sConnStr + "3"
     End Select

     Set MyConn = Server.CreateObject("ADODB.Connection")
     MyConn.Open Application(sConnStr)
  Else
     Set MyConn = Conn
  End If

  Dim dsTemp
  
  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  sSQL = "select * from Docs where DocID = "+sUnicodeSymbol+"'"+sDocID1+"'"
  'out sSQL
  dsTemp.Open sSQL, MyConn, 3, 1, &H1
  If dsTemp.EOF Then
     'GetDocField_test = ""
  Else
     sNameCreation = dsTemp("NameCreation")
     sNameAproval = dsTemp("NameAproval")
     sNameAproved = dsTemp("NameApproved")
     sListToReconcile = dsTemp("ListToReconcile")
     sListReconciled = dsTemp("ListReconciled")
     sLocationPath = dsTemp("LocationPath")
     sIsActive = dsTemp("IsActive")

  End If
  dsTemp.Close

  If NewConnection Then
     MyConn.Close
  End If

End Sub


Sub GetNewDocID_test(sClassDoc,sDepartment,sSubClassParameter,sSubClassParameter2,sParentDocID,sIsProjectDocID)

'out "sSubClassParameter=" & sSubClassParameter
'out "sSubClassParameter2=" & sSubClassParameter2
'out "sDepartment=" & sDepartment
'out "sClassDoc=" & sClassDoc

  NewConnection = True
  If not IsNull(Conn) Then
     NewConnection = not IsObject(Conn)
  End If

  If NewConnection Then
     SetPayDoxPars 'Присваиваем переменные среды из Global.asa
     sConnStr = "ConnectString"
     Select Case UCase(Request("l"))
        Case "RU" sConnStr = sConnStr + "RUS"
        Case "3" sConnStr = sConnStr + "3"
     End Select

     Set MyConn = Server.CreateObject("ADODB.Connection")
     MyConn.Open Application(sConnStr)
  Else
     Set MyConn = Conn
  End If

  If sIsProjectDocID="PJ-" Then
     sDocIDPJ="PJ-"
     sSearchCol="DocIDadd"
  Else
     sDocIDPJ=""
     sSearchCol="DocID"
  End If

  Set dsTempPR = Server.CreateObject("ADODB.Recordset")
'if false then

  Select Case sDepartment
   Case SIT_SITRONICS
     'нумерация для ситроникса
     If InStr(UCase(sClassDoc), UCase(SIT_VHODYASCHIE)) > 0 Then
	    sSeparator = "-"
'Ph - 20100311 - start
	  sPrefix="IN"+Right(CStr(Year(Date)),2)+"/"+Left(sSubClassParameter, NotNegative(InStr(sSubClassParameter, " ")-1))+sSeparator
'	     sPrefix="IN"+Right(CStr(Year(Date)),2)+"/"+Left(sSubClassParameter, InStr(sSubClassParameter, " ")-1)+sSeparator
'Ph - 20100311 - end
	    sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%' AND Department like N'"+sDepartment+"%'"
	 End If

	 If InStr(UCase(sClassDoc), UCase(SIT_VHODYASCHIE_ACC)) > 0 Then
	    sSeparator = "/"
	    sPrefix="ACC"+Right(CStr(Year(Date)),2)+sSeparator
	    'sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%' AND Department like N'"+sDepartment+"%'"
	    sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%'"
	 End If

	 If InStr(UCase(sClassDoc), UCase(SIT_ISHODYASCHIE)) > 0 Then
	    'vnik mikron   
	    If InStr(UCase(sDepartment), UCase(SIT_MIKRON)) > 0 Then
	        sPrefix="MS-OUT"+Right(CStr(Year(Date)),2)+"-"+ sSubClassParameter +"/"
	        sSQL="select ISNULL(Max(cast(case CharIndex('5-', "+sSearchCol+") when 0 then right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1)-2) else right("+sSearchCol+",len("+sSearchCol+")-charindex('-',"+sSearchCol+",len('"+sDocIDPJ+"')+1)) end as int)),0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%' AND Department like N'"+sDepartment+"%'"
	    Else
	    'vnik mikron
	        sPrefix="OUT"+Right(CStr(Year(Date)),2)+"/"
 	        If sParentDocID="" Then
               sSufix="5-"
	           'sSeparator="-"
	        else
	           'sSeparator="/"
	        End If 
	        sSeparator="/"
	        'SAY 2009-03-16
	        'sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%' AND Department like N'"+sDepartment+"%'"
	        sSQL="select ISNULL(Max(cast(case CharIndex('5-', "+sSearchCol+") when 0 then right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1)-2) else right("+sSearchCol+",len("+sSearchCol+")-charindex('-',"+sSearchCol+",len('"+sDocIDPJ+"')+1)) end as int)),0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%' AND Department like N'"+sDepartment+"%'"
	        'sSQL="select ISNULL(Max(cast(case CharIndex('-', right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1))) when 0 then right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1)) else right(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1)),len(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1)))-charindex('-',right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1)),len('"+sDocIDPJ+"')+1)) end as int)),0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%' AND Department like N'"+sDepartment+"%'"

            'out "sSQL1=" + sSQL
        End If
	 End If

	 If InStr(UCase(sClassDoc), UCase(SIT_SLUZH_ZAPISKA)) > 0 Then
	    sSeparator = "-"
'Ph - 20100311 - start
	  sPrefix="IH."+Right(CStr(Year(Date)),2)+"/"+Left(sSubClassParameter, NotNegative(InStr(sSubClassParameter, " ")-1))+sSeparator
'	     sPrefix="IH."+Right(CStr(Year(Date)),2)+"/"+Left(sSubClassParameter, InStr(sSubClassParameter, " ")-1)+sSeparator
'Ph - 20100311 - end
	    sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc like N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%' AND Department like N'"+sDepartment+"%'"
	    AddLogD "vnikvnik " + Trim(ssql)
	 End If

	 If InStr(UCase(sClassDoc), UCase(SIT_RASP_DOCS)) > 0 Then
	    sSeparator = "/"
'Ph - 20100311 - start
	  sPrefix=Left(sSubClassParameter, NotNegative(InStr(sSubClassParameter, " ")-1))+"-"+Right(CStr(Year(Date)),2)+sSeparator
'	     sPrefix=Left(sSubClassParameter, InStr(sSubClassParameter, " ")-1)+"-"+Right(CStr(Year(Date)),2)+sSeparator
'Ph - 20100311 - end
	    sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%' AND Department like N'"+sDepartment+"%'"
	 End If
'vnik_protocols
     If InStr(UCase(sClassDoc), UCase(SIT_PROTOCOLS)) > 0 Then
	    sPrefix= Trim(sSubClassParameter2 + "" + sSubClassParameter)
	    sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-len('"+sPrefix+"'))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sPrefix+"%' AND Department like N'"+sDepartment+"%'"
	 End If
'vnik_protocols
    
'vnik_payment_order
     If InStr(UCase(sClassDoc), UCase(SIT_PAYMENT_ORDER)) > 0 Then
	    sPrefix= Trim(sSubClassParameter2 + "" + sSubClassParameter)
	    sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-len('"+sPrefix+"'))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sPrefix+"%' AND Department like N'"+sDepartment+"%'"
	 End If
'vnik_payment_order
	 If InStr(UCase(sClassDoc), UCase(SIT_NORM_DOCS)) > 0 Then
	    sSeparator = "-"
'Ph - 20100311 - start
	  sPrefix=Left(sSubClassParameter, NotNegative(InStr(sSubClassParameter, " ")-1))+sSeparator
'	     sPrefix=Left(sSubClassParameter, InStr(sSubClassParameter, " ")-1)+sSeparator
'Ph - 20100311 - end
	    sPostfix = "."+sSubClassParameter2
	    sSQL="select ISNULL(Max(cast(left(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1)),charindex('.',right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1)),0)-1) as int)),0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%'"
	 End If

'Запрос №11 - СТС - start
  	If InStr(UCase(sClassDoc), UCase(SIT_DOGOVORI_OLD)) = 1 Then
'  	  If InStr(UCase(sClassDoc), UCase(SIT_DOGOVORI)) > 0 Then
'Запрос №11 - СТС - end
	    sSeparator = "-"
'Ph - 20100311 - start
	    sPrefix= Left(sSubClassParameter, NotNegative(InStr(sSubClassParameter, " ")-1))+sSeparator
'	     sPrefix= Left(sSubClassParameter, InStr(sSubClassParameter, " ")-1)+sSeparator
'Ph - 20100311 - end
	    sPostfix = "/"+LeadSymbolNVal(CStr(Month(Date)),"0",2)+"/"+Right(CStr(Year(Date)),2)
	    sSQL="select ISNULL(Max(cast(left(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1)),charindex('/',right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1)),0)-1) as int)),0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%'"
	 End If

	 If InStr(UCase(sClassDoc), UCase(SIT_ZADACHI)) > 0 Then
	    sSeparator = "_"
	    sPrefix= "T_"

	    If UCase(Trim(sSubClassParameter))=UCase("Поручения АФК") or UCase(Trim(sSubClassParameter))=UCase("SISTEMA tasks") Then
	       sPrefix= "T_AFK_"
	    End If

	    sSufix = ""
	    If Trim(sParentDocID)<>"" then
	       sSufix = "("+sParentDocID + ")_"
	    End If

'Ph - 20091204 - start
        sSQL="select ISNULL(Max(cast(Right("+sSearchCol+", charindex('"+sSeparator+"',reverse("+sSearchCol+"),0)-1) as int)),0) as MaxDocID from Docs where ClassDoc like N'"+sClassDoc+"' and "
	    If InStr(sPrefix,"AFK")>0 Then
	       sSQL = sSQL+" CharIndex(N'"+sDocIDPJ+"T_AFK_', "+sSearchCol+") = 1"
'	        sSQL="select ISNULL(Max(cast(Right("+sSearchCol+", charindex('"+sSeparator+"',reverse("+sSearchCol+"),0)-1) as int)),0) as MaxDocID from Docs where "+sSearchCol+" like N'"+sDocIDPJ+"T_AFK_%' and ClassDoc like N'"+sClassDoc+"'"
	    Else
	       sSQL = sSQL+" CharIndex(N'"+sDocIDPJ+"T_', "+sSearchCol+") = 1 and CharIndex(N'"+sDocIDPJ+"T_AFK_', "+sSearchCol+") <> 1"
'	        sSQL="select ISNULL(Max(cast(Right("+sSearchCol+", charindex('"+sSeparator+"',reverse("+sSearchCol+"),0)-1) as int)),0) as MaxDocID from Docs where "+sSearchCol+" like N'"+sDocIDPJ+"T_%' AND "+sSearchCol+" not like N'%AFK%' and ClassDoc like N'"+sClassDoc+"'"
	    End If
	 End If
'Ph - 20091204 - end

	 If InStr(UCase(sClassDoc), UCase(SIT_HelpDesk)) > 0 Then
	    sSeparator = "/"
	    sPrefix="HD."+Right(CStr(Year(Date)),2)+sSeparator
	    sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%'"
          'out "SIT_HelpDesk=" + SIT_HelpDesk
     End If

'нумерация для СТС
   Case SIT_STS
     If InStr(UCase(sClassDoc), UCase(SIT_VHODYASCHIE)) > 0 Then
	    sSeparator = "/"
'Ph - 20100311 - start
	    sPrefix="IN"+Right(CStr(Year(Date)),2)+"_"+Left(sSubClassParameter, NotNegative(InStr(sSubClassParameter, " ")-1))+sSeparator
'	     sPrefix="IN"+Right(CStr(Year(Date)),2)+"_"+Left(sSubClassParameter, InStr(sSubClassParameter, " ")-1)+sSeparator
'Ph - 20100311 - end
	    sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%' AND Department like N'"+sDepartment+"%'"
	 End If

	 If InStr(UCase(sClassDoc), UCase(SIT_VHODYASCHIE_ACC)) > 0 Then
	    sSeparator = "/"
	    sPrefix="ACC"+Right(CStr(Year(Date)),2)+sSeparator
'        sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%' AND Department like N'"+sDepartment+"%'"
	    sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%'"
     End If

	 If InStr(UCase(sClassDoc), UCase(SIT_ISHODYASCHIE)) > 0 Then
	    sSeparator = "/"
'Ph - 20090313 - start
'Ph - 20100311 - start
	    sPrefix="OUT"+Right(CStr(Year(Date)),2)+"_"+Left(sSubClassParameter2, NotNegative(InStr(sSubClassParameter2, " ")-1))+sSeparator
'	     sPrefix="OUT"+Right(CStr(Year(Date)),2)+"_"+Left(sSubClassParameter2, iif(InStr(sSubClassParameter2, " ") = 0, 0, InStr(sSubClassParameter2, " ")-1))+sSeparator
'	     sPrefix="OUT"+Right(CStr(Year(Date)),2)+"_"+Left(sSubClassParameter2, InStr(sSubClassParameter2, " ")-1)+sSeparator
'Ph - 20100311 - end
'	     sPrefix="OUT"+Right(CStr(Year(Date)),2)+"_"+Left(sSubClassParameter2, InStr(sSubClassParameter, " ")-1)+sSeparator
'Ph - 20090313 - end
'SAY 2009-03-30
'	     sPrefix="OUT"+Right(CStr(Year(Date)),2)+"_"+Left(sSubClassParameter2, InStr(sSubClassParameter2, " ")-1)'+sSeparator

'        sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%' AND Department like N'"+sDepartment+"%'"
'SAY 2009-07-03 изменен запрос (игнорирование некастомизируемых в число номеров)
'        sSQL="select ISNULL(Max(cast(right("+sSearchCol+",charindex('"+sSeparator+"',reverse("+sSearchCol+"),0)-1)as int)),0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%' AND Department like N'"+sDepartment+"%'"
	    sSQL = "select ISNULL(Max(case IsNumeric(right("+sSearchCol+",charindex('"+sSeparator+"',reverse("+sSearchCol+"),0)-1)) when 1 then cast(right("+sSearchCol+",charindex('"+sSeparator+"',reverse("+sSearchCol+"),0)-1)as int) else 0 end),0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%' AND Department like N'"+sDepartment+"%'"
	 End If

	 If InStr(UCase(sClassDoc), UCase(SIT_SLUZH_ZAPISKA)) > 0 Then
	    sSeparator = "/"
	    sPrefix="IH."+Right(CStr(Year(Date)),2)+sSeparator
'SAY 2009-06-23 исправлен запрос на генерирование номера
'        sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc like N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%' AND Department like N'"+sDepartment+"%'"
        sSQL="select ISNULL(Max(case IsNumeric(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1))) when 1 then cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1))as int) else -1 end),0) as MaxDocID from Docs where  ClassDoc like N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%' AND Department like N'"+sDepartment+"%'"

'Session("Message") = "sSubClassParameter=" & sSubClassParameter & "<br>" & "sSubClassParameter2=" & sSubClassParameter2 & "<br>" & "sDepartment=" & sDepartment & "<br>" & "sClassDoc=" & sClassDoct & "<br>" & "sSQL=" & sSQL
'RedirMessage
     End If

     If InStr(UCase(sClassDoc), UCase(SIT_RASP_DOCS)) > 0 Then
	    sSeparator="-"
'Ph - 20100311 - start
	    sPrefix=Left(sSubClassParameter, NotNegative(InStr(sSubClassParameter, " ")-1))+"_"+Right(CStr(Year(Date)),2)+sSeparator
'	     sPrefix=Left(sSubClassParameter, InStr(sSubClassParameter, " ")-1)+"_"+Right(CStr(Year(Date)),2)+sSeparator
'Ph - 20100311 - end
'        sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%' AND Department like N'"+sDepartment+"%'"
'SAY 2009-03-30
        sSQL="select ISNULL(Max(cast(right("+sSearchCol+",charindex('"+sSeparator+"',reverse("+sSearchCol+"),0)-1)as int)),0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%' AND Department like N'"+sDepartment+"%'"
     End If

     If InStr(UCase(sClassDoc), UCase(SIT_NORM_DOCS)) > 0 Then
	    sSeparator = "-"
'Ph - 20100311 - start
	    sPrefix=Left(sSubClassParameter, NotNegative(InStr(sSubClassParameter, " ")-1))+sSeparator
'	     sPrefix=Left(sSubClassParameter, InStr(sSubClassParameter, " ")-1)+sSeparator
'Ph - 20100311 - end
	    sPostfix = "."+sSubClassParameter2
	    sSQL="select ISNULL(Max(cast(left(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1)),charindex('.',right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1)),0)-1) as int)),0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%'"
     End If

'Запрос №11 - СТС - start
  	 If InStr(UCase(sClassDoc), UCase(SIT_DOGOVORI_OLD)) = 1 Then
'  	  If InStr(UCase(sClassDoc), UCase(SIT_DOGOVORI)) > 0 Then
'Запрос №11 - СТС - end
	    sSeparator = "-"
'Ph - 20100311 - start
	    sPrefix= Left(sSubClassParameter, NotNegative(InStr(sSubClassParameter, " ")-1))+sSeparator
'	     sPrefix= Left(sSubClassParameter, InStr(sSubClassParameter, " ")-1)+sSeparator
'Ph - 20100311 - end
	    sPostfix = "/"+LeadSymbolNVal(CStr(Month(Date)),"0",2)+"/"+Right(CStr(Year(Date)),2)
	    sSQL="select ISNULL(Max(cast(left(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1)),charindex('/',right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1)),0)-1) as int)),0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%'"
	 End If

	 If InStr(UCase(sClassDoc), UCase(SIT_ZADACHI)) > 0 Then
	    sSeparator = "_"
	    sPrefix= "T_"

	    If UCase(Trim(sSubClassParameter)) = UCase("Поручения АФК") or UCase(Trim(sSubClassParameter)) = UCase("SISTEMA tasks") Then
	       sPrefix= "T_AFK_"
	    End If 

	    sSufix = ""
	    If Trim(sParentDocID)<>"" then
	       sSufix = "("+sParentDocID + ")_"
	    End If

	    If InStr(sPrefix,"AFK")>0 Then
	       sSQL="select ISNULL(Max(cast(Right("+sSearchCol+", charindex('"+sSeparator+"',reverse("+sSearchCol+"),0)-1) as int)),0) as MaxDocID from Docs where "+sSearchCol+" like N'"+sDocIDPJ+"T[_]AFK[_]%' and ClassDoc like N'"+sClassDoc+"'"
	    Else
	       sSQL="select ISNULL(Max(cast(Right("+sSearchCol+", charindex('"+sSeparator+"',reverse("+sSearchCol+"),0)-1) as int)),0) as MaxDocID from Docs where "+sSearchCol+" like N'"+sDocIDPJ+"T[_]%' AND "+sSearchCol+" not like N'%AFK%' and ClassDoc like N'"+sClassDoc+"'"
	    End If
	End If

	If InStr(UCase(sClassDoc), UCase(SIT_HelpDesk)) > 0 Then
	  sSeparator = "/"
	  sPrefix="HD."+Right(CStr(Year(Date)),2)+sSeparator
	  sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%'"
    End If
'vnik mikron
  Case SIT_MIKRON    
    'нумерация для МИКРОНА
	If InStr(UCase(sClassDoc), UCase(SIT_VHODYASCHIE)) > 0 Then
       sSeparator = "-"
'Ph - 20100311 - start
	   sPrefix="MS-IN"+Right(CStr(Year(Date)),2)+sSeparator
'	    sPrefix="IN"+Right(CStr(Year(Date)),2)+"/"+Left(sSubClassParameter, InStr(sSubClassParameter, " ")-1)+sSeparator
'Ph - 20100311 - end
	   sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+4))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%' AND Department like N'"+sDepartment+"%'"
	End If

	If InStr(UCase(sClassDoc), UCase(SIT_VHODYASCHIE_ACC)) > 0 Then
	   sSeparator = "/"
	   sPrefix="ACC"+Right(CStr(Year(Date)),2)+sSeparator
	   'sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%' AND Department like N'"+sDepartment+"%'"
	   sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%'"
	End If

	If InStr(UCase(sClassDoc), UCase(SIT_ISHODYASCHIE)) > 0 Then
	   sPrefix="MS-OUT"+Right(CStr(Year(Date)),2)+"-"+ sSubClassParameter +"/"
	   sSQL="select ISNULL(Max(cast(case CharIndex('5-', "+sSearchCol+") when 0 then Replace(right("+sSearchCol+",len("+sSearchCol+")-charindex('-',"+sSearchCol+",len('"+sDocIDPJ+"')+4)),'"+sSubClassParameter+"' + '/','') else right("+sSearchCol+",len("+sSearchCol+")-charindex('-',"+sSearchCol+",len('"+sDocIDPJ+"')+1)) end as int)),0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%' AND Department like N'"+sDepartment+"%'"
	   AddLogD "WAY sSQL="+sSQL
	End If

	If InStr(UCase(sClassDoc), UCase(SIT_SLUZH_ZAPISKA)) > 0 Then
	   sSeparator = "-"
'Ph - 20100311 - start
	   sPrefix="IH."+Right(CStr(Year(Date)),2)+"/"+Left(sSubClassParameter, NotNegative(InStr(sSubClassParameter, " ")-1))+sSeparator
'	    sPrefix="IH."+Right(CStr(Year(Date)),2)+"/"+Left(sSubClassParameter, InStr(sSubClassParameter, " ")-1)+sSeparator
'Ph - 20100311 - end
	   sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc like N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%' AND Department like N'"+sDepartment+"%'"
	End If

	If InStr(UCase(sClassDoc), UCase(SIT_RASP_DOCS)) > 0 Then
	   sSeparator = "/"
'Ph - 20100311 - start
	   sPrefix=Left(sSubClassParameter, NotNegative(InStr(sSubClassParameter, " ")-1))+"-"+Right(CStr(Year(Date)),2)+sSeparator
'	    sPrefix=Left(sSubClassParameter, InStr(sSubClassParameter, " ")-1)+"-"+Right(CStr(Year(Date)),2)+sSeparator
'Ph - 20100311 - end
	   sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%' AND Department like N'"+sDepartment+"%'"
	End If
	
	If InStr(UCase(sClassDoc), UCase(SIT_ZADACHI)) > 0 Then
	   sSeparator = "_"
	   sPrefix= "T_"

	   If UCase(Trim(sSubClassParameter)) = UCase("Поручения АФК") or UCase(Trim(sSubClassParameter)) = UCase("SISTEMA tasks") Then
	      sPrefix= "T_AFK_"
	   End If 

	   sSufix = ""
	   If Trim(sParentDocID)<>"" then
	      sSufix = "("+sParentDocID + ")_"
	   End If

'Ph - 20091204 - start
       sSQL="select ISNULL(Max(cast(Right("+sSearchCol+", charindex('"+sSeparator+"',reverse("+sSearchCol+"),0)-1) as int)),0) as MaxDocID from Docs where ClassDoc like N'"+sClassDoc+"' and "
	   If InStr(sPrefix,"AFK")>0 Then
	      sSQL = sSQL+" CharIndex(N'"+sDocIDPJ+"T_AFK_', "+sSearchCol+") = 1"
'	       sSQL="select ISNULL(Max(cast(Right("+sSearchCol+", charindex('"+sSeparator+"',reverse("+sSearchCol+"),0)-1) as int)),0) as MaxDocID from Docs where "+sSearchCol+" like N'"+sDocIDPJ+"T_AFK_%' and ClassDoc like N'"+sClassDoc+"'"
	   Else
	      sSQL = sSQL+" CharIndex(N'"+sDocIDPJ+"T_', "+sSearchCol+") = 1 and CharIndex(N'"+sDocIDPJ+"T_AFK_', "+sSearchCol+") <> 1"
'	       sSQL="select ISNULL(Max(cast(Right("+sSearchCol+", charindex('"+sSeparator+"',reverse("+sSearchCol+"),0)-1) as int)),0) as MaxDocID from Docs where "+sSearchCol+" like N'"+sDocIDPJ+"T_%' AND "+sSearchCol+" not like N'%AFK%' and ClassDoc like N'"+sClassDoc+"'"
	   End If
	End If
'Ph - 20091204 - end
'vnik mikron

  Case SIT_RTI
AddLogD "999"
    If InStr(UCase(sClassDoc), UCase(SIT_ZADACHI)) > 0 Then
       sSeparator = "_"
	   sPrefix= "T_"

	   If UCase(Trim(sSubClassParameter)) = UCase("Поручения АФК") or UCase(Trim(sSubClassParameter)) = UCase("SISTEMA tasks") Then
	      sPrefix= "T_AFK_"
	   End If 

	   sSufix = ""
	   If Trim(sParentDocID)<>"" then
	      sSufix = "("+sParentDocID + ")_"
	   End If

'Ph - 20091204 - start
      sSQL="select ISNULL(Max(cast(Right("+sSearchCol+", charindex('"+sSeparator+"',reverse("+sSearchCol+"),0)-1) as int)),0) as MaxDocID from Docs where ClassDoc like N'"+sClassDoc+"' and "
	  If InStr(sPrefix,"AFK")>0 Then
	    sSQL = sSQL+" CharIndex(N'"+sDocIDPJ+"T_AFK_', "+sSearchCol+") = 1"
'	       sSQL="select ISNULL(Max(cast(Right("+sSearchCol+", charindex('"+sSeparator+"',reverse("+sSearchCol+"),0)-1) as int)),0) as MaxDocID from Docs where "+sSearchCol+" like N'"+sDocIDPJ+"T_AFK_%' and ClassDoc like N'"+sClassDoc+"'"
	  Else
	    sSQL = sSQL+" CharIndex(N'"+sDocIDPJ+"T_', "+sSearchCol+") = 1 and CharIndex(N'"+sDocIDPJ+"T_AFK_', "+sSearchCol+") <> 1"
'	       sSQL="select ISNULL(Max(cast(Right("+sSearchCol+", charindex('"+sSeparator+"',reverse("+sSearchCol+"),0)-1) as int)),0) as MaxDocID from Docs where "+sSearchCol+" like N'"+sDocIDPJ+"T_%' AND "+sSearchCol+" not like N'%AFK%' and ClassDoc like N'"+sClassDoc+"'"
	  End If
	  End If    
    Case "OTHER B.U."
	'для других бизнес направлений
  End Select

'20090622 - Заявка ТКП
  'Коммерческие предложения - сквозная нумерация для всех БН
  If InStr(UCase(sClassDoc),UCase(SIT_COM_OFFERS)) > 0 Then
     sDocIDPJ = ""
     sSufix = ""
     sPostfix = ""
     sPrefix = "CO-"
     sSQL="select ISNULL(MAX(cast(REPLACE(DocID,"+sUnicodeSymbol+"'"+sPrefix+"','') as int)),0) as MaxDocID from Docs where ClassDoc="+sUnicodeSymbol+"'"+sClassDoc+"'"
  End If

'out sSQL

  dsTempPR.Open sSQL, MyConn, 2, 1, &H1
  If dsTempPR.EOF Then
    sNumberNext = 0
  Else
    sNumberNext = dsTempPR("MaxDocID")
  End If
  dsTempPR.Close

'  out "sNumberNext=" & CStr(sNumberNext)
'  out "DocID (first iteration): " + sDocIDPJ + sPrefix + sSufix + LeadSymbolNVal(sNumberNext, "0", 3)+sPostfix
  
  StopIncrement = False
  do while not StopIncrement
    sNumberNext=sNumberNext+1
'20090622 - Заявка ТКП
    'Для коммерческих предложений разрядность 6, для остальных - 3
    'vnik_protocols для протоколов тоже 6
    'vnik_payment_order для заявок на оплату УК тоже 6
    If InStr(UCase(sClassDoc),UCase(SIT_COM_OFFERS)) > 0 or InStr(UCase(sClassDoc),UCase(SIT_PROTOCOLS)) > 0 or InStr(UCase(sClassDoc),UCase(SIT_PAYMENT_ORDER)) > 0 Then
    'vnik_payment_order
    'vnik_protocols
      if sNumberNext<1000000 Then
        stringNumberNext = LeadSymbolNVal(sNumberNext, "0", 6)
      Else
        stringNumberNext = CStr(sNumberNext)
      End If 
    'vnik mikron
    ElseIf (InStr(UCase(sClassDoc),UCase(SIT_VHODYASCHIE)) > 0 or InStr(UCase(sClassDoc),UCase(SIT_ISHODYASCHIE)) > 0) and InStr(sDepartment, SIT_MIKRON) > 0 Then
        if sNumberNext<10000 Then
            stringNumberNext = LeadSymbolNVal(sNumberNext, "0", 4)
        Else
            stringNumberNext = CStr(sNumberNext)
        End If
    'vnik mikron
    Else
      'SAY 2009-05-13 предотвращение переполнения номера
      if sNumberNext<1000 Then
        stringNumberNext = LeadSymbolNVal(sNumberNext, "0", 3)
      Else
        stringNumberNext = CStr(sNumberNext)
      End If
      'SAY 2009-05-13 предотвращение переполнения номера    
    End If
    'S_DocID = sDocIDPJ + sPrefix + sSufix + LeadSymbolNVal(sNumberNext, "0", 3)+sPostfix
    S_DocID = sDocIDPJ + sPrefix + sSufix + stringNumberNext + sPostfix

    sSQL = "SELECT DocID FROM Docs WHERE DocID=N'"+S_DocID+"'"
    dsTempPR.Open sSQL, MyConn, 2, 1, &H1
    If dsTempPR.EOF Then
      StopIncrement = True
    End If
    dsTempPR.Close  
  loop
    
'end If
  'out "sSQL="&sSQL
  'out "S_DocID="&S_DocID
  'iTemp = 1/0

  If NewConnection Then
    MyConn.Close
  End If

End Sub


'Добавить текстовую строку к сообщению (новой строкой)
Function AddNewLineToMessage(ByVal sMessage, ByVal sAddMessage)
  AddNewLineToMessage = Trim(sMessage)
  If AddNewLineToMessage <> "" Then
    AddNewLineToMessage = AddNewLineToMessage+VbCrLf
  End If
  AddNewLineToMessage = AddNewLineToMessage+sAddMessage
End Function

'Функция определения, что проект не нулевой (для заявок СТС)
Function IsProject(ByVal ProjectNo)
  IsProject = InStr("123456789", Left(ProjectNo, 1)) <> 0 or Mid(ProjectNo, 2) <> "00000"
End Function

'ph - 20100623 - start
Function NotNegativeSQL(ByVal Par)
  NotNegativeSQL = "Case When "+Par+" > 0 Then "+Par+" Else 0 End"
End Function
'ph - 20100623 - end

'Запросы к БД для источников справочников Заявок и для проверки введенных значений
Function ProjectListSelectForInsert()
'  ProjectListSelectForInsert = "select ProjectID, ProjectCode,ProjectName from ProjectList where ProjectCurrentStatus <> N'Closed' and ProjectCurrentStatus <> N'Canceled' and ProjectCurrentStatus <> N'Frozen' order by ProjectID"
'  Exit Function
'  ProjectManagerLogin = "Case CharIndex(N'<', ProjectManagerUser) When 0 Then N'' Else Case CharIndex(N'>', ProjectManagerUser) When 0 Then N'' Else SubString(ProjectManagerUser, CharIndex(N'<', ProjectManagerUser)+1, CharIndex(N'>', ProjectManagerUser)-CharIndex(N'<', ProjectManagerUser)-1) End End"
  ProjectManagerLogin = "SubString(ProjectManagerUser, CharIndex(N'<', ProjectManagerUser)+1, Case CharIndex(N'>', ProjectManagerUser)-CharIndex(N'<', ProjectManagerUser) When 0 Then 0 Else CharIndex(N'>', ProjectManagerUser)-CharIndex(N'<', ProjectManagerUser)-1 End)"
  ShortDepartment1 = "Case Right(Department, 1) When '/' Then Case CharIndex(N'/', Reverse(Left(Department, Len(Department)-1))) When 0 Then Left(Department, Len(Department)-1) Else Right(Left(Department, Len(Department)-1), CharIndex(N'/', Reverse(Left(Department, Len(Department)-1)))-1) End Else Case CharIndex(N'/', Reverse(Department)) When 0 Then Department Else Right(Department, CharIndex(N'/', Reverse(Department))-1) End End"
  ShortDepartment = "Case Department When NULL Then '' When '' Then '' Else "+ShortDepartment1+" End"
'  RusDepartment = "Left("+ShortDepartment+", CharIndex(N'*', "+ShortDepartment+")-1)"
  RusDepartment = "Left("+ShortDepartment+", Case CharIndex(N'*', "+ShortDepartment+") When 0 Then Len("+ShortDepartment+") Else CharIndex(N'*', "+ShortDepartment+")-1 End)"
'ph - 20100623 - start
  EngDepartment = "SubString("+ShortDepartment+", CharIndex(N'*', "+ShortDepartment+")+1, "+NotNegativeSQL("Len("+ShortDepartment+")-CharIndex(N'*', Reverse("+ShortDepartment+"))-CharIndex(N'*', "+ShortDepartment+")")+")"
'  EngDepartment = "SubString("+ShortDepartment+", CharIndex(N'*', "+ShortDepartment+")+1, Len("+ShortDepartment+")-CharIndex(N'*', Reverse("+ShortDepartment+"))-CharIndex(N'*', "+ShortDepartment+"))"
'ph - 20100623 - end
'  CZDepartment = "Right("+ShortDepartment+", CharIndex(N'*', Reverse("+ShortDepartment+"))-1)"
  CZDepartment = "Right("+ShortDepartment+", Case CharIndex(N'*', Reverse("+ShortDepartment+")) When 0 Then Len("+ShortDepartment+") Else CharIndex(N'*', Reverse("+ShortDepartment+"))-1 End)"
  ReplaceEmptyCZByEng = "Case "+CZDepartment+" When '' Then "+EngDepartment+" Else "+CZDepartment+" End"
  DepartmentInCurrentLang = "Case '"+UCase(Request("l"))+"' When 'RU' Then "+RusDepartment+" When '' Then "+EngDepartment+" When '3' Then "+ReplaceEmptyCZByEng+" End"

  RusName = "Left(Name, Case CharIndex(N'*', Name) When 0 Then Len(Name) Else CharIndex(N'*', Name)-1 End)"
'ph - 20100623 - start
  EngName = "SubString(Name, CharIndex(N'*', Name)+1, "+NotNegativeSQL("Len(Name)-CharIndex(N'*', Reverse(Name))-CharIndex(N'*', Name)")+")"
'  EngName = "SubString(Name, CharIndex(N'*', Name)+1, Len(Name)-CharIndex(N'*', Reverse(Name))-CharIndex(N'*', Name))"
'ph - 20100623 - end
  CZName = "Right(Name, Case CharIndex(N'*', Reverse(Name)) When 0 Then Len(Name) Else CharIndex(N'*', Reverse(Name))-1 End)"
  ReplaceEmptyCZByEng = "Case "+CZName+" When '' Then "+EngName+" Else "+CZName+" End"
  NameInCurrentLang = "Case '"+UCase(Request("l"))+"' When 'RU' Then "+RusName+" When '' Then "+EngName+" When '3' Then "+ReplaceEmptyCZByEng+" End"
  FullName = "Case UserID When '' Then '' Else N'""'+"+NameInCurrentLang+"+N'"" <'+UserID+N'>;' End"

  'Conditions = "where ProjectCurrentStatus <> N'Closed' and ProjectCurrentStatus <> N'Canceled' and ProjectCurrentStatus <> N'Frozen' order by ProjectID"

  'rmanyushin 63140, 178534 16.02.2011
  Conditions = "where ProjectCurrentStatus <> N'NotForPO' and ProjectCurrentStatus <> N'Closed' and ProjectCurrentStatus <> N'Canceled' and ProjectCurrentStatus <> N'Frozen' order by substring(ProjectID,2,1)+substring(ProjectID,4,1)+substring(ProjectID,5,1)+substring(ProjectID,6,1)"
  ProjectListSelectForInsert = "select ProjectID, ProjectCode, ProjectName, "+FullName+" as Manager, Case CharIndex(N'*', Department) When 0 Then Department Else "+DepartmentInCurrentLang+" End as CostCenter from ProjectList left join Users on ("+ProjectManagerLogin+" = UserID) "+Conditions
End Function

Function ProjectListSelectForReport()
'  ProjectListSelectForReport = "select distinct UserFieldText3+N' - '+UserFieldText4 as Project from Docs where ClassDoc like N'"+STS_PaymentOrder+"%' order by Project"
  ProjectListSelectForReport = "select distinct UserFieldText3+N' - '+UserFieldText4 as Project from Docs where ClassDoc like N'"+STS_PaymentOrder+"%' or ClassDoc like N'"+STS_PurchaseOrder+"%' order by Project"
End Function

Function BusinessUnitSelectForInsert()
  BusinessUnitSelectForInsert = "select BusinessUnit+' - '+Company_"+iif(Request("l") = "", "EN", iif(Request("l") = "3", "CZ", "RU"))+" as BusinessUnit from BusinessUnits"
'  If Request("DocID") <> "" Then
'  If Request("DocID") <> "" or Request("UserID") <> "" Then
  If UCase(Session("CurrentPage")) <> UCase("/ModifyDependants.asp") Then 'Session("CurrentPage")  устанавлиается в UserText
    BusinessUnitSelectForInsert = BusinessUnitSelectForInsert + " where BusinessUnits.BusinessUnit not like "+sUnicodeSymbol+"'%000' and BusinessUnits.BusinessUnit <> "+sUnicodeSymbol+"'9999'"
  End If
  BusinessUnitSelectForInsert = BusinessUnitSelectForInsert + " order by BusinessUnit"
End Function

Function BusinessUnitSelectForCheckValue(ByVal Value)
'  BusinessUnitSelectForCheckValue = "select * from BusinessUnits where BusinessUnit+' - '+Company_"+iif(Request("l") = "", "EN", iif(Request("l") = "3", "CZ", "RU"))+" = "+sUnicodeSymbol+"'" + Value + "' or BusinessUnit+' - '+Company_EN = "+sUnicodeSymbol+"'" + Value + "'"
'  BusinessUnitSelectForCheckValue = "select * from BusinessUnits where (BusinessUnit+' - '+Company_"+iif(Request("l") = "", "EN", iif(Request("l") = "3", "CZ", "RU"))+" = "+sUnicodeSymbol+"'" + Value + "' or BusinessUnit+' - '+Company_EN = "+sUnicodeSymbol+"'" + Value + "') and (BusinessUnit not like "+sUnicodeSymbol+"'%000' and BusinessUnit <> "+sUnicodeSymbol+"'9999')"
  BusinessUnitSelectForCheckValue = "select * from BusinessUnits where (BusinessUnit = "+sUnicodeSymbol+"'"+Value+"' or BusinessUnit+' - '+Company_"+iif(Request("l") = "", "EN", iif(Request("l") = "3", "CZ", "RU"))+" = "+sUnicodeSymbol+"'" + Value + "' or BusinessUnit+' - '+Company_EN = "+sUnicodeSymbol+"'" + Value + "') and (BusinessUnit not like "+sUnicodeSymbol+"'%000' and BusinessUnit <> "+sUnicodeSymbol+"'9999')"
End Function

Function ChartOfAccountsSelectForInsert()
'  ChartOfAccountsSelectForInsert = "select GroupName, STS_Account_No+' - '+AccountName_"+iif(Request("l") = "", "EN", iif(Request("l") = "3", "CZ", "RU"))+" as AccountName from ChartOfAccounts order by GroupName,[Order]"
'  ChartOfAccountsSelectForInsert = "select GroupName_"+iif(Request("l") = "", "EN", iif(Request("l") = "3", "CZ", "RU"))+" as GroupName, STS_Account_No+' - '+AccountName_"+iif(Request("l") = "", "EN", iif(Request("l") = "3", "CZ", "RU"))+" as AccountName from ChartOfAccounts order by GroupName,[Order]"
  ChartOfAccountsSelectForInsert = "select GroupName_"+iif(Request("l") = "", "EN", iif(Request("l") = "3", "CZ", "RU"))+" as GroupName, STS_Account_No+' - '+AccountName_"+iif(Request("l") = "", "EN", iif(Request("l") = "3", "CZ", "RU"))+" as AccountName from ChartOfAccounts order by GroupOrder,GroupName,[Order]"
End Function

Function ChartOfAccountsSelectForCheckValue(ByVal Value)
'  ChartOfAccountsSelectForCheckValue = "select * from ChartOfAccounts where STS_Account_No+' - '+AccountName_"+iif(Request("l") = "", "EN", iif(Request("l") = "3", "CZ", "RU"))+" = "+sUnicodeSymbol+"'" + Value + "' or STS_Account_No+' - '+AccountName_EN = "+sUnicodeSymbol+"'" + Value + "'"
  ChartOfAccountsSelectForCheckValue = "select * from ChartOfAccounts where STS_Account_No = "+sUnicodeSymbol+"'"+Value+"' or STS_Account_No+' - '+AccountName_"+iif(Request("l") = "", "EN", iif(Request("l") = "3", "CZ", "RU"))+" = "+sUnicodeSymbol+"'" + Value + "' or STS_Account_No+' - '+AccountName_EN = "+sUnicodeSymbol+"'" + Value + "'"
End Function

Function PaymentTypesSelectForInsert()
  PaymentTypesSelectForInsert = "select PaymentType_"+iif(Request("l") = "", "EN", iif(Request("l") = "3", "CZ", "RU"))+" as PaymentType from PaymentTypes order by PaymentType"
End Function
Function PaymentTypesSelectForCheckValue(ByVal Value)
  PaymentTypesSelectForCheckValue = "select * from PaymentTypes where PaymentType_"+iif(Request("l") = "", "EN", iif(Request("l") = "3", "CZ", "RU"))+" = "+sUnicodeSymbol+"'" + Value + "' or PaymentType_EN = "+sUnicodeSymbol+"'" + Value + "'"
End Function

Function CostCentersSelectForInsert()
  ShortName = "Case Right(Name, 1) When '/' Then Case CharIndex(N'/', Reverse(Left(Name, Len(Name)-1))) When 0 Then Left(Name, Len(Name)-1) Else Right(Left(Name, Len(Name)-1), CharIndex(N'/', Reverse(Left(Name, Len(Name)-1)))-1) End Else Case CharIndex(N'/', Reverse(Name)) When 0 Then Name Else Right(Name, CharIndex(N'/', Reverse(Name))-1) End End"
'  RusName = "Left("+ShortName+", CharIndex(N'*', "+ShortName+")-1)"
  RusName = "Left("+ShortName+", Case CharIndex(N'*', "+ShortName+") When 0 Then Len("+ShortName+") Else CharIndex(N'*', "+ShortName+")-1 End)"
  EngName = "SubString("+ShortName+", CharIndex(N'*', "+ShortName+")+1, Len("+ShortName+")-CharIndex(N'*', Reverse("+ShortName+"))-CharIndex(N'*', "+ShortName+"))"
'  CZName = "Right("+ShortName+", CharIndex(N'*', Reverse("+ShortName+"))-1)"
  CZName = "Right("+ShortName+", Case CharIndex(N'*', Reverse("+ShortName+")) When 0 Then Len("+ShortName+") Else CharIndex(N'*', Reverse("+ShortName+"))-1 End)"
  ReplaceEmptyCZByEng = "Case "+CZName+" When '' Then "+EngName+" Else "+CZName+" End"
  Conditions = "where Name like N'СТР%' and Statuses like N'%#BL=Yes%'"
  CostCentersSelectForInsert = "select Case '"+UCase(Request("l"))+"' When 'RU' Then "+RusName+" When '' Then "+EngName+" When '3' Then "+ReplaceEmptyCZByEng+" End as CostCenter from Departments "+Conditions+"  order by CostCenter"
End Function

Function CostCentersSelectForCheckValue(ByVal Value)
  ShortName = "Case Right(Name, 1) When '/' Then Case CharIndex(N'/', Reverse(Left(Name, Len(Name)-1))) When 0 Then Left(Name, Len(Name)-1) Else Right(Left(Name, Len(Name)-1), CharIndex(N'/', Reverse(Left(Name, Len(Name)-1)))-1) End Else Case CharIndex(N'/', Reverse(Name)) When 0 Then Name Else Right(Name, CharIndex(N'/', Reverse(Name))-1) End End"
'  RusName = "Left("+ShortName+", CharIndex(N'*', "+ShortName+")-1)"
  RusName = "Left("+ShortName+", Case CharIndex(N'*', "+ShortName+") When 0 Then Len("+ShortName+") Else CharIndex(N'*', "+ShortName+")-1 End)"
  EngName = "SubString("+ShortName+", CharIndex(N'*', "+ShortName+")+1, Len("+ShortName+")-CharIndex(N'*', Reverse("+ShortName+"))-CharIndex(N'*', "+ShortName+"))"
'  CZName = "Right("+ShortName+", CharIndex(N'*', Reverse("+ShortName+"))-1)"
  CZName = "Right("+ShortName+", Case CharIndex(N'*', Reverse("+ShortName+")) When 0 Then Len("+ShortName+") Else CharIndex(N'*', Reverse("+ShortName+"))-1 End)"
  ReplaceEmptyCZByEng = "Case "+CZName+" When '' Then "+EngName+" Else "+CZName+" End"
  Conditions = "where Name like N'СТР%' and Statuses like N'%#BL=Yes%'"
'  CostCentersSelectForCheckValue = "select "+EngName+" as CostCenterEng from Departments "+Conditions+" and "+iif(Request("l") = "", EngName, iif(Request("l") = "3", CZName, RusName))+" = "+sUnicodeSymbol+"'" + Value + "' or "+EngName+" = "+sUnicodeSymbol+"'" + Value + "'"
  CostCentersSelectForCheckValue = "select "+EngName+" as CostCenterEng from Departments "+Conditions+" and ("+iif(Request("l") = "", EngName, iif(Request("l") = "3", CZName, RusName))+" = "+sUnicodeSymbol+"'" + Value + "' or "+EngName+" = "+sUnicodeSymbol+"'" + Value + "' or CharIndex("+sUnicodeSymbol+"'" + Value + "'+' ', "+EngName+") = 1)"
End Function

'Определение даты автоматического согласования для нормативных документов
Function GetNormDocLastReconcileDate
  Select Case Weekday(Date)
    Case 1 'Sunday
      GetNormDocLastReconcileDate = Date+8
    Case 2 'Monday
      GetNormDocLastReconcileDate = Date+8
    Case 3
      GetNormDocLastReconcileDate = Date+8
    Case 4
      GetNormDocLastReconcileDate = Date+8
    Case 5
      GetNormDocLastReconcileDate = Date+8
    Case 6
      GetNormDocLastReconcileDate = Date+10
    Case 7
      GetNormDocLastReconcileDate = Date+10
  End Select
End Function

'Функция преобразования Фамилии сотрудника к одноязычному формату
Function InsertionName(Name, UserID)
  InsertionName = GetFullName(SurnameGN(DelOtherLangFromFolder(Name)), UserID)+";"
End Function

'Показать кнопку в ShowDoc
Sub ShowDoc_ShowButton(VAR_ButtonsToShow, VAR_ButtonsNotToShow, ByVal Button)
'  If InStr(VAR_ButtonsToShow, Button) = 0 Then
'    VAR_ButtonsToShow = VAR_ButtonsToShow + Button + ","
'  End If
  VAR_ButtonsNotToShow = Replace(VAR_ButtonsNotToShow, Button + ",", "")
End Sub

'Спрятать кнопку в ShowDoc
Sub ShowDoc_HideButton(VAR_ButtonsToShow, VAR_ButtonsNotToShow, ByVal Button)
  If InStr(VAR_ButtonsNotToShow+",", Button+",") = 0 Then
    VAR_ButtonsNotToShow = VAR_ButtonsNotToShow + Button + ","
  End If
'  VAR_ButtonsToShow = Replace(VAR_ButtonsToShow, Button + ",", "")
End Sub

'Определение необходимости перевода строки (при добавлении значения к тексту)
Function IsNeedVbCrLf(ByVal sString)
  If Trim(sString)="" Then
    IsNeedVbCrLf = ""
  Else
    IsNeedVbCrLf = VbCrLf
  End If
End Function

'Получить поле документа из БД (поле должно существовать, иначе ошибка 500 :))
Function SIT_GetDocField(sDocID, FieldName, Conn)
  Dim dsTemp
  
  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  dsTemp.Open "select * from Docs where DocID = "+sUnicodeSymbol+"'"+sDocID+"'", Conn, 3, 1, &H1
  If dsTemp.EOF Then
    SIT_GetDocField = ""
  Else
  	SIT_GetDocField = dsTemp(FieldName)
  End If
  dsTemp.Close
End Function

'Поменять местами фамилию и инициалы (инициалы в конец)
Function SIT_SurnameGN(ByVal SurnameGN)
  dim i
  i = InstrRev(SurnameGN, ".")
  SIT_SurnameGN = Trim(Mid(SurnameGN, i+1))+" "+Left(SurnameGN, i)
End Function

Function GivenNames(ByVal UserName)
  GivenNames = Trim(Mid(UserName, Instr(UserName, " ")+1))
End Function

'Множитель для перевода одной валюты в другую
Function CurrencyConvertionFactor(ByVal FromCurrency, ByVal ToCurrency)
  CurrencyConvertionFactor = CCur(0)
  If UCase(FromCurrency) = UCase(ToCurrency) Then
    CurrencyConvertionFactor = CCur(1)
	Exit Function
  End If
  
  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  sSQL = "Select IsNULL(Rate, 0) as Rate from CurrencyRates where Code = "+sUnicodeSymbol+"'"+ToCurrency+"' or Code2 = "+sUnicodeSymbol+"'"+ToCurrency+"'"
AddlogD "CurrencyConvertionFactor SQL1: "+sSQL
  dsTemp.Open sSQL, Conn, 3, 1, &H1
  If dsTemp.EOF Then
    CurrencyConvertionFactor = CCur(0)
    dsTemp.Close
	Exit Function
  Else
    CurrencyConvertionFactor = CCur(dsTemp("Rate"))
  End If
  dsTemp.Close
  If CurrencyConvertionFactor = CCur(0) Then
    Exit Function
  End If

  sSQL = "Select IsNULL(Rate, 0) as Rate from CurrencyRates where Code = "+sUnicodeSymbol+"'"+FromCurrency+"' or Code2 = "+sUnicodeSymbol+"'"+FromCurrency+"'"
AddlogD "CurrencyConvertionFactor SQL2: "+sSQL
  dsTemp.Open sSQL, Conn, 3, 1, &H1
  If dsTemp.EOF Then
    CurrencyConvertionFactor = CCur(0)
    dsTemp.Close
	Exit Function
  Else
    CurrencyConvertionFactor = CCur(dsTemp("Rate"))/CurrencyConvertionFactor
  End If
AddlogD "CurrencyConvertionFactor: "+CStr(CurrencyConvertionFactor)
  dsTemp.Close
End Function

'Запрос №36 - СТС - start - Больше НЕ используется, переделано на правила
'Получить лимиты из БД (возвращются в переменных)
'Sub GetLimitsForOrders(STS_HeadOfSector_Limit, STS_HeadOfDepartment_Limit, STS_HeadOfDivision_Limit, STS_FinancialControl_Limit, STS_FinDirector_Limit, STS_GenDirector_Limit, STS_ProjectManager_Limit, STS_Accounting_Limit)
'  Set dsTemp = Server.CreateObject("ADODB.Recordset")
'  sSQL = "Select * from AmountLimits order by Limit"
'  dsTemp.Open sSQL, Conn, 3, 1, &H1
'  do while not dsTemp.EOF
'    Select Case UCase(CStr(dsTemp("Role")))
'	  Case UCase(STS_Orders_HeadOfSector)
'        STS_HeadOfSector_Limit = dsTemp("Limit")
'	  Case UCase(STS_Orders_HeadOfDepartment)
'        STS_HeadOfDepartment_Limit = dsTemp("Limit")
'	  Case UCase(STS_Orders_HeadOfDivision)
'        STS_HeadOfDivision_Limit = dsTemp("Limit")
'	  Case UCase(STS_Orders_FinancialControl)
'        STS_FinancialControl_Limit = dsTemp("Limit")
'	  Case UCase(STS_Orders_FinDirector)
'        STS_FinDirector_Limit = dsTemp("Limit")
'	  Case UCase(STS_Orders_GenDirector)
'        STS_GenDirector_Limit = dsTemp("Limit")
'	  Case UCase(STS_Orders_ProjectManager)
'        STS_ProjectManager_Limit = dsTemp("Limit")
'	  Case UCase(STS_Orders_Accounting)
'        STS_Accounting_Limit = dsTemp("Limit")
'	End Select
'	dsTemp.MoveNext
'  loop
'  dsTemp.Close
'End Sub
'Запрос №36 - СТС - end

'Получение индекса заявок и коммерческих предложений
Function GetNewOrderDocID(sClassDoc)
  iDigits = 6
  sYear = "/"+Right(CStr(Year(Date)), 2)

  If InStr(UCase(sClassDoc),UCase(STS_PurchaseOrder)) > 0 Then
    sPrefix = "PO-"
  ElseIf InStr(UCase(sClassDoc),UCase(STS_PaymentOrder)) > 0 Then
    sPrefix = "P-"
   
  Else
    GetNewOrderDocID = "0"
    Exit Function
  End If

  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  sSQL="select ISNULL(MAX(cast(left(REPLACE(DocID,"+sUnicodeSymbol+"'"+sPrefix+"',''),"+CStr(iDigits)+") as int)),0) as MaxDocID from Docs where ClassDoc="+sUnicodeSymbol+"'"+sClassDoc+"'"
AddlogD "@@@GetOrderDocID SQL: "+sSQL
  dsTemp.Open sSQL, Conn, 3, 1, &H1
  GetNewOrderDocID = CStr(dsTemp("MaxDocID")+1)
  dsTemp.Close
  
  do while Len(GetNewOrderDocID) < iDigits
    GetNewOrderDocID = "0"+GetNewOrderDocID
  loop

  GetNewOrderDocID = sPrefix + GetNewOrderDocID + sYear
AddlogD "@@@GetOrderDocID DocID: "+GetNewOrderDocID
End Function

'Вычленить код из строки вида "код (- )значение"
Function GetCodeFromCode_NameString(ByVal sValue)
  Dim iPos
  iPos = InStr(sValue, " ")
  If iPos = 0 Then
'ph - 20090704 - start
    GetCodeFromCode_NameString = sValue
'    GetCodeFromCode_NameString = ""
'ph - 20090704 - end
  Else
    GetCodeFromCode_NameString = Left(sValue, iPos-1)
  End If
End Function

Function GetCostCenterByCode(ByVal sCode)
  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  sSQL = "Select Name from Departments where Code = "+sUnicodeSymbol+"'"+sCode+"' and Statuses like "+sUnicodeSymbol+"'%#BL=Yes%'"
AddlogD "GetCostCenterByCode SQL: "+sSQL
  dsTemp.Open sSQL, Conn, 3, 1, &H1
  If dsTemp.EOF Then
    GetCostCenterByCode = ""
  Else
    GetCostCenterByCode = dsTemp("Name")
  End If
  dsTemp.Close
End Function

'Вычленить из списка пользователей на рассылку уведомлений (для исходящих в другое бизнес-направление)
Function GetCorrectUsersFromList(sUserList, Conn)
Dim iPos, dsTemp, sUserID, sUserIDList, sAdditionalList

  iPos = 1
  sAdditionalList = ""
  sUserIDList = ""
  sUserID = oPayDox.GetNextUserIDInList(sUserList, iPos)
  If sUserID <> "" Then
    If InStr(sUserID, "USERS:") = 1 or InStr(sUserID, "DEPARTMENTS:") = 1 Then
      sAdditionalList = sAdditionalList + "<"+sUserID+">; "
    Else
      sUserIDList = "'"+sUserID+"'"
    End If
  Else
    GetCorrectUsersFromList = ""
    Exit Function
  End If
  Do While sUserID <> ""
    sUserID = oPayDox.GetNextUserIDInList(sUserList, iPos)
    If sUserID <> "" Then
      If InStr(sUserID, "USERS:") = 1 or InStr(sUserID, "DEPARTMENTS:") = 1 Then
        sAdditionalList = sAdditionalList + "<"+sUserID+">; "
      Else
        If sUserIDList <> "" Then
          sUserIDList = sUserIDList + ",'"+sUserID+"'"
        Else
          sUserIDList = "'"+sUserID+"'"
        End If
      End If
    End If
  Loop

  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  dsTemp.CursorLocation = 3
  dsTemp.Open "select * from users where UserID in ("+sUserIDList+") and WinLogin <> '' and WinLogin is not NULL and StatusActive <> '' and StatusActive is not NULL", Conn, 3, 1, &H1
  dsTemp.ActiveConnection = Nothing
  
  GetCorrectUsersFromList = ""
  Do While not dsTemp.EOF
    GetCorrectUsersFromList = GetCorrectUsersFromList+"<"+dsTemp("UserID")+">; "
	dsTemp.MoveNext
  Loop
  
  dsTemp.Close
  GetCorrectUsersFromList = GetCorrectUsersFromList + sAdditionalList
  AddlogD "GetCorrectUsersFromList: "+GetCorrectUsersFromList
End Function

'Определение можно ли пользователю загружать справочник проектов
Function CanLoadSTSProjectList(ByVal sUser)
Dim sCheckUser
  sCheckUser = UCase(sUser)
  CanLoadSTSProjectList = InStr(sCheckUser, UCase("Admin")) = 1 or InStr(sCheckUser, UCase("fincontrol")) = 1
'out "CanLoadSTSProjectList - UserID: """+sUser+""" CanLoadSTSProjectList = "+CStr(CanLoadSTSProjectList)
End Function

'Получить имя пользователя на текущем языке для вставки в поля
Function GetInsertionNameInCurrentLanguage(ByVal sUser)
  UserID = GetUserID(sUser)
  If UserID = "" Then
    GetInsertionNameInCurrentLanguage = sUser
    Exit Function
  End If
  oPayDox.GetUserDetails UserID, sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
  GetInsertionNameInCurrentLanguage = InsertionName(sName, UserID)
End Function

'Получить список значений из справочника для заполнения выпадающего списка
Function GetUserDirValues2(ByVal DirGUID, ByVal DataFieldName)
  Dim DirKeyField

  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  sSQL = "Select KeyField from UserDirectories where GUID = '"&DirGUID&"'"
  AddlogD "GetUserDirValues2 FindDirKeyField SQL: "+sSQL
  dsTemp.Open sSQL, Conn, 3, 1, &H1
  
  GetUserDirValues2 = ""
  If dsTemp.EOF Then
    dsTemp.Close
	Exit Function
  Else
    DirKeyField = dsTemp("KeyField")
	dsTemp.Close
  End If

  sSQL = "Select "+DataFieldName+" from UserDirValues where UDKeyField = "&CStr(DirKeyField)&" order by "+DataFieldName
  AddlogD "GetUserDirValues2 GetValues SQL: "+sSQL
  dsTemp.Open sSQL, Conn, 3, 1, &H1

  do while not dsTemp.EOF
    If GetUserDirValues2 <> "" Then
	  GetUserDirValues2 = GetUserDirValues2 + VbCrLf
	End If
    GetUserDirValues2 = GetUserDirValues2+dsTemp(DataFieldName)
    dsTemp.MoveNext
  loop

  dsTemp.Close
End Function

'vnik_payment_order
'получить свойства элемента справочника по переданному значению
'DirGUID - идентификатор справочника свойства из которого хотим получить
'DataFieldName - свойства которые хотим получить
'FieldSelection - поле по которому отбираем
'ValueSelection - значение по которому отбираем
Function GetUserDirValuesVNIK(ByVal DirGUID, ByVal DataFieldName, ByVal FieldSelection, ByVal ValueSelection)
Dim DirKeyFieldVNIK

  Set dsTempVNIK = Server.CreateObject("ADODB.Recordset")
  sSQL = "Select KeyField from UserDirectories where GUID = '"&DirGUID&"'"
  AddlogD "GetUserDirValuesVNIK FindDirKeyField SQL: "+sSQL
  dsTempVNIK.Open sSQL, Conn, 3, 1, &H1
  AddlogD "GetUserDirValuesVNIK FindDirKeyField SQL: "+Trim(dsTempVNIK("KeyField"))
  GetUserDirValuesVNIK = ""

  If dsTempVNIK.EOF Then
     dsTempVNIK.Close
	 Exit Function
  Else
     DirKeyFieldVNIK = Trim(dsTempVNIK("KeyField"))

     AddlogD "GetUserDirValuesVNIK:" + DirKeyFieldVNIK
	 dsTempVNIK.Close
  End If

'amw 10/07/2013
'  sSQL = "Select "+DataFieldName+" from UserDirValues where UDKeyField = '"&CStr(DirKeyFieldVNIK)&"' and "&FieldSelection&" = '"&CStr(ValueSelection)&"' order by "+DataFieldName
'  sSQL = "Select "+DataFieldName+" from UserDirValues where UDKeyField='"+CStr(DirKeyFieldVNIK)+"' and "+FieldSelection+" like N'%"+CStr(ValueSelection)+"%' order by "+DataFieldName
  sSQL = "Select "+DataFieldName+" from UserDirValues where UDKeyField = '"&CStr(DirKeyFieldVNIK)&"' and "&FieldSelection&" = N'"&CStr(ValueSelection)&"' order by "+DataFieldName
  AddlogD "GetUserDirValuesVNIK GetValues SQL: "+sSQL
  dsTempVNIK.Open sSQL, Conn, 3, 1, &H1
  
  do while not dsTempVNIK.EOF
     If GetUserDirValuesVNIK <> "" Then
	    GetUserDirValuesVNIK = GetUserDirValuesVNIK + VbCrLf
	 End If
     GetUserDirValuesVNIK = GetUserDirValuesVNIK + dsTempVNIK(DataFieldName)
     dsTempVNIK.MoveNext
  loop
  
  dsTempVNIK.Close
End Function
'vnik_payment_order

'Аналогично предыдущей, но значения берутся из первого поля справочника
Function GetUserDirValues(ByVal DirGUID)
  GetUserDirValues = GetUserDirValues2(DirGUID, "Field1")
End Function

'Получить список всех категорий документов (через запятую)
Function GetCategoriesList
  Dim Conn, ds
  
  Set Conn = CreateObject("ADODB.Connection")
  Conn.Open Application("ConnectStringRUS")
  Set ds = CreateObject("ADODB.Recordset")
  ds.Open "select * from DocTypes order by Name", Conn, 1, 3, &H1
  GetCategoriesList = ""
  do while not ds.EOF
    GetCategoriesList = GetCategoriesList + DelOtherLangFromFolder(ds("Name")) + ","
    ds.MoveNext
  loop
  ds.Close
  GetCategoriesList = Left(GetCategoriesList, Len(GetCategoriesList)-1)

  Conn.Close
End Function

'Получить список валют (для выпадающего списка)
Function GetCurrencyList()
  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  sSQL = "Select Code from CurrencyRates"
  dsTemp.Open sSQL, Conn, 3, 1, &H1

  GetCurrencyList = ""
  do while not dsTemp.EOF
    If GetCurrencyList <> "" Then
	  GetCurrencyList = GetCurrencyList + VbCrLf
	End If
    GetCurrencyList = GetCurrencyList+dsTemp("Code")
    dsTemp.MoveNext
  loop

  dsTemp.Close
End Function


'rmanyushin 93489 21.04.2010
Function GetCategoriesList2
  Dim Conn, ds
  Set Conn = CreateObject("ADODB.Connection")
  Conn.Open Application("ConnectStringRUS")
  Set ds = CreateObject("ADODB.Recordset")
  ds.Open "select * from DocTypes order by Name", Conn, 1, 3, &H1
  GetCategoriesList2 = ""
  do while not ds.EOF
    GetCategoriesList2 = GetCategoriesList2 + GetRelativeName(DelOtherLangFromFolder(ds("Name"))) + ","
    ds.MoveNext
  loop
  ds.Close
  GetCategoriesList2 = Left(GetCategoriesList2, Len(GetCategoriesList2)-1)

  Conn.Close
End Function

Function GetRelativeName(ByVal strInputName)
Dim strFind	
	strFind = Len(strInputName) - InStrRev(strInputName, "/") 
	GetRelativeName = Right(strInputName, strFind)
End Function
'rmanyushin 93489 21.04.2010


'Получить список значений выбранного поля и указанной таблицы (для выпадающего списка)
Function GetExtTableValues(ByVal sTableName, ByVal sFieldName)
  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  sSQL = "Select "+sFieldName+" from "+sTableName
  dsTemp.Open sSQL, Conn, 3, 1, &H1

  GetExtTableValues = ""
  do while not dsTemp.EOF
    If GetExtTableValues <> "" Then
	   GetExtTableValues = GetExtTableValues+VbCrLf
	End If
    GetExtTableValues = GetExtTableValues+dsTemp(sFieldName)
    dsTemp.MoveNext
  loop

  dsTemp.Close
End Function

'Добавить согласующего с проверкой, что он не создатель (передается параметром 
'sCurrentLogin) и без дублирования
'sDelimeter - пробел или VbCrLf в зависимости от типа добавления согласующего
Function AddUserToReconcileList(ByVal sListToReconcile, ByVal sCurrentLogin, ByVal sDelimeter, ByVal sNewUser)
  AddUserToReconcileList = sListToReconcile
  If InStr(UCase(sNewUser), "<"+UCase(sCurrentLogin)+">") > 0 Then
	Exit Function
  End If
  'В данном варианте при задвоении добавляем нового согласующего, а старого убираем
'ph - 20090704 - start
  If InStr(UCase(AddUserToReconcileList), "<"+UCase(GetUserID(sNewUser))+">") > 0 Then
'  If InStr(UCase(AddUserToReconcileList), UCase(sNewUser)) > 0 Then
'ph - 20090704 - end
'    AddUserToReconcileList = Replace(AddUserToReconcileList, sNewUser, "")
    AddUserToReconcileList = RemoveUserFromListWithDescriptions(AddUserToReconcileList, sNewUser)
  End If
  AddUserToReconcileList = AddUserToReconcileList+sDelimeter+sNewUser
End Function

' ==============================
' AM 09122008 - функция проверки соответствия введеных в поле данных значениям из справочника
' sFieldName - поле таблицы, по которому искать; 
' sTableName - таблица, по которой искать; 
' sFormFieldName - имя поля формы, для которого производится поиск. Брать надо стандартное имя поля таблицы Docs, например, ListToView
' sValueToCheck - значение по которому искать (может быть несколько значений через разделитель); 
' sDelimiter - разделитель в случае поиска по нескольким значениям. Оставить пустые, если предполагается, что поиск только по одному значению
' sLangCode - язык, на котором выдавать сообщение
Function CheckIfValueIsInDirectory(sFieldName,sTableName,sFormFieldName,sValueToCheck,sDelimiter,sLangCode)
  sFormFieldName=Trim(sFormFieldName)
  sFieldName=Trim(sFieldName)
  sTableName=Trim(sTableName)
  sValueToCheck=Trim(sValueToCheck)
  If (sFormFieldName<>"" And sFieldName<>"" And sTableName<>"" And sValueToCheck<>"") Then
    sMessage=SIT_ErrorInFieldValue1 + GetDocFieldDescription(sFormFieldName) + SIT_ErrorInFieldValue2
    sValues=Split(Replace(sValueToCheck,VbCrLf,""),sDelimiter)
	nCounter=UBound(sValues)
    If nCounter=0 Then
	  CheckIfValueIsInDirectory=sMessage
	  Exit Function
	End If
    For i=0 To UBound(sValues)
	  If i=UBound(sValues) Then
	    Exit For
	  End If
	  ' Специфические вещи для пользователей, остальные виды проверяемых данных возможно работают проще
	  If sTableName="Users" Then
        sValuesList = sValuesList + iif(i=0,"'",",'") + GetUserID(sValues(i)+">;") + "'"
	  Else
        sValuesList = sValuesList + iif(i=0,"'",",'") + sValues(i) + "'"
	  End If
    Next
	If sTableName="Users" Then
      sSQL = "select " + sFieldName + " from " + sTableName + " where " + sFieldName + " in (" + sValuesList + ") and WinLogin <> '' and WinLogin is not NULL and StatusActive <> '' and StatusActive is not NULL"
    Else
      sSQL = "select " + sFieldName + " from " + sTableName + " where " + sFieldName + " in (" + sValuesList + ")"
	End If
    Set dsTempCheck=Server.CreateObject("ADODB.Recordset") 
    dsTempCheck.Open sSQL, Conn, 3, 1, &H1
    If Not dsTempCheck.EOF Then
	  If nCounter>0 Then
        nRecords=dsTempCheck.RecordCount
		If nCounter <> nRecords Then
		  bCheckFailed=True
		End If
	  End If
    Else
      bCheckFailed=True
    End If
    dsTempCheck.Close
    If bCheckFailed Then
      CheckIfValueIsInDirectory=sMessage
	Else
	  CheckIfValueIsInDirectory=""
    End If
  End If
End Function

' AM 10092008 Функция для отчетов. Вывод выпадающего списка значений пользовательского справочника по известному GUID и названию поля справочника
Function GetExtDirListOfValues(sGUID,sFieldName)
  Dim Conn, dsExtDir
  Set Conn = CreateObject("ADODB.Connection")
  Conn.Open Application("ConnectStringRUS")
  Set dsExtDir = CreateObject("ADODB.Recordset")
  sGUID=LCase(sGUID)
  sSQL="select "+sFieldName+" from userdirvalues where UDKeyField=(select KeyField from UserDirectories where GUID='"+sGUID+"')"
  dsExtDir.Open sSQL, Conn, 3, 1, &H1
  GetExtDirListOfValues=""
  do while not dsExtDir.EOF
    GetExtDirListOfValues = GetExtDirListOfValues + dsExtDir(sFieldName) + ","
    dsExtDir.MoveNext
  Loop
  GetExtDirListOfValues = "," + GetExtDirListOfValues
  dsExtDir.Close
  Conn.Close
End Function

' AM 12092008 Функция для отчетов, для визуального представления статусов исполнения
Function CheckStatusExpired(S_DateCompletion)
  If Request("l")="ru" Then
    sMessage="Просрочено"
  Else
    sMessage="Expired"
  End If
  If dsDoc("StatusCompletion")<>"1" And DateFullTime(dsDoc("DateCompletion"))<Now And Not IsNull(dsDoc("DateCompletion")) And dsDoc("DateCompletion")>VAR_BeginOfTimes Then
    CheckStatusExpired="<b><font color=red>" + CStr(MyDate(dsDoc("DateCompletion"))) + "</font></b>" + "&nbsp; <img border=0 src=" + GetPayDoxURL() + "IMAGES/pict_expired.gif alt=DOCS_EXPIRED3 width=16 height=16>"
  ElseIf dsDoc("StatusCompletion")<>"1" And DateFullTime(dsDoc("DateCompletion"))=Now And Not IsNull(dsDoc("DateCompletion")) And dsDoc("DateCompletion")>VAR_BeginOfTimes Then
    CheckStatusExpired="<font color=red>" + CStr(MyDate(dsDoc("DateCompletion"))) + "</font>"
  End If
End Function 

'Удаление дубликатов из списка пользователей (для списка согласующих)
Function DeleteUserDoublesInList(ByVal sUsersList)
  If Trim(sUsersList)="" or InStr(sUsersList,";")=0 Then
    DeleteUserDoublesInList = sUsersList    
    Exit Function
  End If

  nStr=Split(sUsersList,";")
  ReDim vArr(UBound(nStr))
  ReDim vArrOut(UBound(nStr))
  For i=0 To UBound(nStr)
    vArr(i)=Trim(nStr(i))
  Next
  vArrOut(0)=vArr(0)
  n=1
  For i=1 To UBound(vArr)
    k=0
    For j=0 To n-1
      'If vArr(i)=vArrOut(j) Then
'      If replace(vArr(i),VbCrLf,"")=replace(vArrOut(j),VbCrLf,"") Then
      If Trim(replace(vArr(i),VbCrLf,""))=Trim(replace(vArrOut(j),VbCrLf,"")) Then
        Exit For
      End If
      k=j+1
    Next
    If k=n Then
      vArrOut(n)=vArr(i)
      n=n+1
    End If
  Next
  'ReDim Preserve vArrOut(n-2)
  ReDim Preserve vArrOut(n-1)
  DeleteUserDoublesInList = ""
  For i=0 To UBound(vArrOut)
    If vArrOut(i)<>"" Then
      DeleteUserDoublesInList = DeleteUserDoublesInList + vArrOut(i) + ";"
    End If
  Next
  'DeleteUserDoublesInList = DeleteUserDoublesInList + vArrOut(i)
End Function

' ==============================

'Показ дней просрочки для отчетов
Function DaysOfExpiration(ByVal iDays)
  If iDays <= 0 Then
    DaysOfExpiration = ""
  Else
    DaysOfExpiration = CStr(iDays)
  End If
End Function

'Показ ссылки на документ для отчета (если вывод в Word/Excel, то без ссылки)
Function MyLinkToDocument(ByVal DocID, ByVal DocName)
  If Trim(Request("R1"))="MSWord" or Trim(Request("R1"))="MSExcel" Then
    MyLinkToDocument = DocID
  Else
    MyLinkToDocument = "<a href="""+GetURLEncode("showdoc.asp", "?docid=", DocID)+""" title="""+DocName+""" target=""_blank"">"+DocID+"</a>"
  End If
End Function

'Показ ссылки на документ для отчета (если вывод в Word/Excel, то без ссылки)
Function MyLinkToDocAndParent(ByVal DocID, ByVal DocName, ByVal DocIDParent)
  If Trim(Request("R1"))="MSWord" or Trim(Request("R1"))="MSExcel" Then
    MyLinkToDocAndParent = DocID
  Else
    MyLinkToDocAndParent = "<a href="""+GetURLEncode("showdoc.asp", "?docid=", DocID)+""" title="""+DocName+""" target=""_blank"">"+DocID+"</a>"
    If Trim(DocIDParent) <> "" Then
      MyLinkToDocAndParent = MyLinkToDocAndParent + "&nbsp;&nbsp;<a href="""+GetURLEncode("showdoc.asp", "?docid=", DocIDParent)+""" title="""+DOCS_DocParent+": "+DocIDParent+""" target=""_blank""><img border=""0"" src=""IMAGES/pict_document.GIF""></a>"
    End If
  End If
End Function

'SAY 2008-09-11 функция отсечения префикса перед ##
Function DeleteConstPrefixFromList(ByVal sUsersList)
  If InStr(sUsersList,"##") > 0 Then
    DeleteConstPrefixFromList = Right(sUsersList, Len(sUsersList) - InStr(sUsersList,"##")-2)
  Else
    DeleteConstPrefixFromList = sUsersList
  End If
End Function 

'Получение списка доступных пользователю бизнес-единиц
Function GetUsersBusinessUnits(ByVal sUserID)
  Set dsUsers = CreateObject("ADODB.Recordset")
  sSQL="select BusinessUnits from Users where UserID = N'"+sUserID+"'"
AddLogD "GetUsersBusinessUnits SQL: "+sSQL
  dsUsers.Open sSQL, Conn, 3, 1, &H1
  If dsUsers.EOF Then
    GetUsersBusinessUnits = ""
  Else
    GetUsersBusinessUnits = Trim(MyCStr(dsUsers("BusinessUnits")))
	'Если не указаны, то добавляем все
	If GetUsersBusinessUnits = "" Then
	  GetUsersBusinessUnits = GetAllAllowableBusinessUnits()
	Else
'      'Вырезаем пустые строки, убираем лишние концы строк
'	  arBusinessUnits = Split(GetUsersBusinessUnits, VbCrLf)
'	  GetUsersBusinessUnits = ""
'	  For i = LBound(arBusinessUnits) to UBound(arBusinessUnits)
'	    If Trim(arBusinessUnits(i)) <> "" Then
'		  If GetUsersBusinessUnits <> "" Then
'		    GetUsersBusinessUnits = GetUsersBusinessUnits + VbCrLf
'		  End If
'		  GetUsersBusinessUnits = GetUsersBusinessUnits + arBusinessUnits(i)
'		End If
'	  Next
      'Получаем список бизнес единиц через запятую
	  arBusinessUnits = Split(GetUsersBusinessUnits, VbCrLf)
	  GetUsersBusinessUnits = ""
	  For i = LBound(arBusinessUnits) to UBound(arBusinessUnits)
	    If Trim(arBusinessUnits(i)) <> "" Then
		  If GetUsersBusinessUnits <> "" Then
		    GetUsersBusinessUnits = GetUsersBusinessUnits + ", "
		  End If
		  GetUsersBusinessUnits = GetUsersBusinessUnits + "'" + Left(arBusinessUnits(i), InStr(arBusinessUnits(i)+" ", " ")-1) + "'"
		End If
	  Next
	  GetUsersBusinessUnits = GetBusinessUnitsInCurrentLanguage(GetUsersBusinessUnits)
	End If
  End If
  dsUsers.Close
End Function

'Получение списка бизнес единиц на нужном языке (вызывается из GetUsersBusinessUnits)
Function GetBusinessUnitsInCurrentLanguage(ByVal sBusinessUnitList)
  If Trim(sBusinessUnitList) = "" Then
    GetBusinessUnitsInCurrentLanguage = ""
	Exit Function
  End If
  Set dsTemp = CreateObject("ADODB.Recordset")
  sSQL = "select BusinessUnit+' - '+Company_"+iif(Request("l") = "", "EN", iif(Request("l") = "3", "CZ", "RU"))+" as FullBusinessUnit from BusinessUnits where BusinessUnit in ("+sBusinessUnitList+") order by BusinessUnit"
AddLogD "GetBusinessUnitsInCurrentLanguage - SQL: " + sSQL
  dsTemp.Open sSQL, Conn, 3, 1, &H1
  GetBusinessUnitsInCurrentLanguage = ""
  do while not dsTemp.EOF
    If GetBusinessUnitsInCurrentLanguage <> "" Then
      GetBusinessUnitsInCurrentLanguage = GetBusinessUnitsInCurrentLanguage+VbCrLf
	End If
    GetBusinessUnitsInCurrentLanguage = GetBusinessUnitsInCurrentLanguage+dsTemp("FullBusinessUnit")
	dsTemp.MoveNext
  loop
  dsTemp.Close
End Function

'Получение списка всех бизнес единиц (не консолидированных)
Function GetAllAllowableBusinessUnits()
  Set dsTemp = CreateObject("ADODB.Recordset")
  sSQL = BusinessUnitSelectForInsert()
AddLogD "GetAllAllowableBusinessUnits - SQL: " + sSQL
  dsTemp.Open sSQL, Conn, 3, 1, &H1
  GetAllAllowableBusinessUnits = ""
  do while not dsTemp.EOF
    GetAllAllowableBusinessUnits = GetAllAllowableBusinessUnits+dsTemp("BusinessUnit")+VbCrLf
	dsTemp.MoveNext
  loop
  dsTemp.Close
End Function

'Функция получения уровня подразделения по его наименованию (из поля Statuses по ключу #LEVX, где X - уровень)
Function GetDepartmentLevel(ByVal sDepartment)
  Set dsTemp = Server.CreateObject("ADODB.Recordset")

  sSQL = "select Statuses from Departments where Name="+sUnicodeSymbol+"'" + sDepartment + "' or Name=N'" + sDepartment + "/'"
AddLogD "GetDepartmentLevel, sSQL="+sSQL

  dsTemp.Open sSQL, Conn, 3, 1, &H1
  If Not dsTemp.EOF Then
    GetDepartmentLevel = dsTemp("Statuses")
  Else
    GetDepartmentLevel = ""
  End If
  dsTemp.Close

  If GetDepartmentLevel <> "" Then
    iLevPos = InStr(GetDepartmentLevel, "#LEV")
	If iLevPos = 0 Then
      GetDepartmentLevel = ""
	Else
      GetDepartmentLevel = Mid(GetDepartmentLevel, iLevPos, 5)
	End If
  End If
End Function

'Вспомогательная функция для GetChiefOfDepartment
Function CorrectLeaderName(ByVal sLeader)
  CorrectLeaderName = Trim(MyCStr(sLeader))
  If CorrectLeaderName = "" Then
    CorrectLeaderName = "#EMPTY"
  ElseIf GetUserID(CorrectLeaderName) = "" Then
    CorrectLeaderName = "#EMPTY"
  Else
    CorrectLeaderName = GetInsertionNameInCurrentLanguage(sLeader)
  End If
End Function

'Вывести сообщение об ошибке поиска руководителя
Sub ShowChiefErrorMessage(ByVal sErrCode, ByVal sDepartment, ByVal sBusinessUnit)
  Dim sMessage
  
  Select Case sErrCode
    Case "#EMPTY"
	  sMessage = SIT_ErrorInDepartmentLeader + sDepartment + ", " + SIT_BusinessUnit + ": " + sBusinessUnit
    Case "#MANY"
	  sMessage = SIT_MoreThanOneLeader1 + sDepartment + SIT_MoreThanOneLeader2 + ", " + SIT_BusinessUnit + ": " + sBusinessUnit
	Case Else
	  sMessage = ""
  End Select
  
  If sMessage <> "" Then
AddLogD "ShowChiefErrorMessage: " + sMessage
    Session("Message") = AddNewLineToMessage(Session("Message"), sMessage)
  End If
End Sub

'Получить руководителя данного подразделения (без движений по структуре)
'Коды ошибок (обрабатываются в ShowChiefErrorMessage):
'#EMPTY - запись есть, но пустая или без логина
'#MANY - больше одного руководителя
'Пустое значение - руководитель не найден (нет записи)
Function GetChiefOfDepartment(ByVal sDepartment, ByVal sBusinessUnit)
  Dim BusinessUnit
  
  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  sSQL = "Select Leader from Departments inner Join DepartmentDependants ON (Departments.GUID = DepartmentDependants.DependantGUID) where (Name="+sUnicodeSymbol+"'" + sDepartment+ "' or Name="+sUnicodeSymbol+"'" + sDepartment+ "/')"
  BusinessUnit = Left(sBusinessUnit, InStr(sBusinessUnit+" ", " ")-1)
  If sBusinessUnit = "" Then
    sSQLAdd = ""
  Else
    sSQLAdd = " and BusinessUnit like "+sUnicodeSymbol+"'"+BusinessUnit+"%'"
  End If
AddLogD "GetChiefOfDepartment SQL1: "+sSQL+sSQLAdd
  dsTemp.Open sSQL+sSQLAdd, Conn, 3, 1, &H1
  If Not dsTemp.EOF Then
    If dsTemp.RecordCount = 1 Then
      GetChiefOfDepartment = CorrectLeaderName(dsTemp("Leader"))
    Else
      GetChiefOfDepartment = "#MANY"
    End If
	ShowChiefErrorMessage GetChiefOfDepartment, sDepartment, BusinessUnit
  Else
    dsTemp.Close

	If sBusinessUnit = "" Then
      'Для пустой бизнес единицы больше делать нечего
	  GetChiefOfDepartment = ""
	  Exit Function
	End If

    'Ищем по таблице руководителей на совпадение агрегатных БЕ
    sSQLAdd = " and BusinessUnit like "+sUnicodeSymbol+"'"+Left(BusinessUnit, 1)+"000%'"
AddLogD "GetChiefOfDepartment SQL2: "+sSQL+sSQLAdd
    dsTemp.Open sSQL+sSQLAdd, Conn, 3, 1, &H1
    If Not dsTemp.EOF Then
      If dsTemp.RecordCount = 1 Then
        GetChiefOfDepartment = CorrectLeaderName(dsTemp("Leader"))
      Else
        GetChiefOfDepartment = "#MANY"
      End If
      ShowChiefErrorMessage GetChiefOfDepartment, sDepartment, Left(BusinessUnit, 1)+"000"
    Else
      dsTemp.Close
      'Ищем по таблице руководителей по БЕ 9999
      sSQLAdd = " and BusinessUnit like "+sUnicodeSymbol+"'9999%'"
AddLogD "GetChiefOfDepartment SQL3: "+sSQL+sSQLAdd
      dsTemp.Open sSQL+sSQLAdd, Conn, 3, 1, &H1
      If Not dsTemp.EOF Then
        If dsTemp.RecordCount = 1 Then
          GetChiefOfDepartment = CorrectLeaderName(dsTemp("Leader"))
        Else
          GetChiefOfDepartment = "#MANY"
        End If
        ShowChiefErrorMessage GetChiefOfDepartment, sDepartment, "9999"
      Else
        GetChiefOfDepartment = "" 
      End If
    End If
  End If
  dsTemp.Close
End Function

'Получить руководителя определенного уровня для данного подразделения (внутренняя)
'Специальные возвращаемые значения:
' "-" - уровень переданного подразделения выше запрошенного
' "=" - руководитель совпадает с переданным пользователем
Function GetChiefOfDepUpperByLevelSys(ByVal sDepartment, ByVal nDepartmentLevel, ByVal sUser, ByVal sBusinessUnit)
  Dim sTemp

  sTemp = sDepartment
  GetChiefOfDepUpperByLevelSys = ""

  'цикл пока не найдем нужный уровень
  nCurDepLevel = 9
  Do while nCurDepLevel > nDepartmentLevel and sTemp <> ""
    sCurDepLevel = GetDepartmentLevel(sTemp)
    If sCurDepLevel <> "" Then
      nCurDepLevel = CInt(Right(sCurDepLevel,1))
    Else
      nCurDepLevel = 0
    End If

    If nCurDepLevel < nDepartmentLevel Then
      GetChiefOfDepUpperByLevelSys = "-" 'Такое значение обрабатывается в GetRoleForOrders_WithCheck и GetChiefOfDepUpperByLevel
AddLogD "GetChiefOfDepUpperByLevelSys - Exiting: nCurDepLevel = "+CStr(nCurDepLevel)+" < nDepartmentLevel = "+CStr(nDepartmentLevel)
      Exit Function
    End If

    If nCurDepLevel = nDepartmentLevel Then
	  GetChiefOfDepUpperByLevelSys = GetChiefOfDepartment(sTemp, sBusinessUnit)
	  'Совпадает с текущим сотрудником
      If InStr(sUser, "<") > 0 Then
        sUserID = GetUserID(sUser)
	  Else
        sUserID = sUser
	  End If
	  If UCase(GetUserID(GetChiefOfDepUpperByLevelSys)) = UCase(sUserID) Then
	    GetChiefOfDepUpperByLevelSys = "=" 'Такое значение обрабатывается в GetRoleForOrders_WithCheck и GetChiefOfDepUpperByLevel
	  End If
AddLogD "GetChiefOfDepUpperByLevelSys(sDepartment="""+sDepartment+""", nDepartmentLevel="""+CStr(nDepartmentLevel)+""", sUser="""+sUser+""", sBusinessUnit="""+sBusinessUnit+""") = """+GetChiefOfDepUpperByLevelSys+""""
      Exit Function
    End If

    If InStrRev(sTemp,"/") > 0 Then
      sTemp = Left(sTemp, InStrRev(sTemp,"/") - 1)
    Else
      sTemp = ""
    End If
  Loop
End Function

'Получить руководителя определенного уровня для данного подразделения (для внешнего вызова)
Function GetChiefOfDepUpperByLevel(ByVal sDepartment, ByVal nDepartmentLevel, ByVal sUser, ByVal sBusinessUnit)
  GetChiefOfDepUpperByLevel = GetChiefOfDepUpperByLevelSys(sDepartment, nDepartmentLevel, sUser, sBusinessUnit)
  If GetChiefOfDepUpperByLevel = "-" or GetChiefOfDepUpperByLevel = "=" Then
    GetChiefOfDepUpperByLevel = ""
  End If
End Function

'Получить непосредственного начальника пользователя (с движением вверх, пока не будет найден)
Function GetNearestChief(ByVal sDepartment, ByVal sUserID, ByVal sBusinessUnit)
  Dim sTemp

AddLogD "GetNearestChief, sDepartment:" + sDepartment + ", sUserID:" + sUserID + ", sBusinessUnit:" + sBusinessUnit
  'цикл пока не найдем начальника
  sTemp = sDepartment
  Do
    GetNearestChief = GetChiefOfDepartment(sTemp, sBusinessUnit)
	If GetUserID(GetNearestChief) <> "" and UCase(GetUserID(GetNearestChief)) <> UCase(sUserID) Then
AddLogD "GetNearestChief - found sDepartment: " + sTemp + ", Chief: " + GetNearestChief
	  Exit Function
    End If

    N = InStrRev(sTemp, "/")
    If N < 2 Then
      Exit Do
    End If
    sTemp = Left(sTemp, N - 1)
  Loop Until sTemp = "" 'InStrRev(sTemp, "/") = 0
  'Обнуляем значение, чтобы для высшего руководителя не стоял он сам
  GetNearestChief = ""
AddLogD "GetNearestChief - not found sDepartment: " + sTemp
End Function

'Получить всех вышестоящих руководителей подразделения
Function GetAllUpperChiefs(ByVal sDepartment, ByVal sUserID, ByVal sBusinessUnit)
  Dim sTemp
  Dim sChief

AddLogD "GetAllUpperChiefs, sDepartment:" + sDepartment + ", sUserID:" + sUserID + ", sBusinessUnit:" + sBusinessUnit
  sTemp = sDepartment
  GetAllUpperChiefs = ""
  Do
    sChief = GetChiefOfDepartment(sTemp, sBusinessUnit)
AddLogD "GetAllUpperChiefs - CurrentDepartment = " + sTemp + ", Chief = " + sChief
	If GetUserID(sChief) <> "" and UCase(GetUserID(sChief)) <> UCase(sUserID) Then
	  If InStr(GetAllUpperChiefs, "<"+GetUserID(sChief)+">") = 0 Then
	    If GetAllUpperChiefs <> "" Then
		  GetAllUpperChiefs = GetAllUpperChiefs + VbCrLf
        End If
        GetAllUpperChiefs = GetAllUpperChiefs + sChief
	  End If
    End If

    N = InStrRev(sTemp, "/")
    If N < 2 Then
      Exit Do
    End If
    sTemp = Left(sTemp, N - 1)
  Loop Until sTemp = "" 'InStrRev(sTemp, "/") = 0
AddLogD "GetAllUpperChiefs: " + GetAllUpperChiefs
End Function

'Получить всех непосредственных руководителей сотрудников из списка
Function GetChiefsOfUsersFromList(ByVal sUsersList, ByVal sBusinessUnit)
  Dim iPos
  Dim sUserID, sChief
AddLogD "GetChiefsOfUsersFromList, sUsersList = " + sUsersList + ", BusinessUnit = " + sBusinessUnit

  GetChiefsOfUsersFromList = ""
  iPos = 1
  sUserID = oPayDox.GetNextUserIDInList(sUsersList, iPos)
  Do While sUserID <> ""
    sChief = GetNearestChief(oPayDox.GetUserDepartment(sUserID), sUserID, sBusinessUnit)
	If sChief <> "" Then
	  If InStr(GetChiefsNameFromList, "<"+GetUserID(sChief)+">") = 0 Then 'Сразу проверяем дубли
	    If GetChiefsOfUsersFromList <> "" Then
		  GetChiefsOfUsersFromList = GetChiefsOfUsersFromList + VbCrLf
        End If
        GetChiefsOfUsersFromList = GetChiefsOfUsersFromList + sChief
	  End If
	End If
    sUserID = oPayDox.GetNextUserIDInList(sUsersList, iPos)
  Loop
AddLogD "GetChiefsOfUsersFromList = " + GetChiefsOfUsersFromList
End Function

'Получить всех вышестоящих руководителей сотрудников из списка
Function GetAllUpperChiefsOfUsersFromList(ByVal sUsersList, ByVal sBusinessUnit)
  Dim iPos
  Dim sUserID, sChiefs
AddLogD "GetAllUpperChiefsOfUsersFromList, sUsersList = " + sUsersList + ", BusinessUnit = " + sBusinessUnit

  GetAllUpperChiefsOfUsersFromList = ""
  iPos = 1
  sUserID = oPayDox.GetNextUserIDInList(sUsersList, iPos)
  Do While sUserID <> ""
    sChiefs = GetAllUpperChiefs(oPayDox.GetUserDepartment(sUserID), sUserID, sBusinessUnit)
    If GetAllUpperChiefsOfUsersFromList <> "" Then
      GetAllUpperChiefsOfUsersFromList = GetAllUpperChiefsOfUsersFromList + VbCrLf
    End If
    GetAllUpperChiefsOfUsersFromList = GetAllUpperChiefsOfUsersFromList + sChiefs
    sUserID = oPayDox.GetNextUserIDInList(sUsersList, iPos)
  Loop
  'Удаляем лишнее
  GetAllUpperChiefsOfUsersFromList = RemoveDubsFromUserListWithCrLf(GetAllUpperChiefsOfUsersFromList)
AddLogD "GetAllUpperChiefsOfUsersFromList = " + GetAllUpperChiefsOfUsersFromList
End Function

'Почистить список пользователей от дублей и пустых значений (для GetAllUpperChiefsOfUsersFromList)
'считается, что список разделен концами строк
Function RemoveDubsFromUserListWithCrLf(ByVal sList)
  Dim arList
  Dim i
  Dim sUserID

AddLogD "RemoveDubsFromUserListWithCrLf initial list: " + sList
  RemoveDubsFromUserListWithCrLf = ""
  If Trim(sList) <> "" Then
    arList = Split(sList, VbCrLf)
	For i = LBound(arList) To UBound(arList)
	  sUserID = GetUserID(arList(i))
	  If sUserID <> "" Then
	    If InStr(RemoveDubsFromUserListWithCrLf, "<"+sUserID+">") = 0 Then
		  If RemoveDubsFromUserListWithCrLf <> "" Then
		    RemoveDubsFromUserListWithCrLf = RemoveDubsFromUserListWithCrLf + VbCrLf
		  End If
		  RemoveDubsFromUserListWithCrLf = RemoveDubsFromUserListWithCrLf + arList(i)
		End If
	  End If
	Next
  End If
AddLogD "RemoveDubsFromUserListWithCrLf result list: " + RemoveDubsFromUserListWithCrLf
End Function

'Вычленить код из названия подразделения СТС
Function GetSTSDepartmentCode(ByVal sDepartment)
  Dim arTemp
  GetSTSDepartmentCode = Trim(sDepartment)
  If GetSTSDepartmentCode = "" Then
    Exit Function
  End If
  If Right(GetSTSDepartmentCode, 1) = "/" Then
    GetSTSDepartmentCode = Left(GetSTSDepartmentCode, Len(GetSTSDepartmentCode)-1)
  End If
  arTemp = Split(GetSTSDepartmentCode, "/")
  GetSTSDepartmentCode = Trim(arTemp(UBound(arTemp)))
  GetSTSDepartmentCode = Left(GetSTSDepartmentCode, InStr(GetSTSDepartmentCode+" ", " ")-1)
  If not IsNumeric(GetSTSDepartmentCode) Then
    GetSTSDepartmentCode = ""
  End If
End Function

'Заменить роль в полях документа на найденного пользователя 
'Sub ReplaceInLists(sString, sNewString)
'  S_NameAproval=Replace(S_NameAproval, sString, sNewString) 
'  S_ListToEdit=Replace(S_ListToEdit, sString, sNewString)
'  S_NameResponsible=Replace(S_NameResponsible, sString, sNewString)
'  S_NameControl=Replace(S_NameControl, sString, sNewString)
'  S_ListToView=Replace(S_ListToView, sString, sNewString)
'  S_Correspondent=Replace(S_Correspondent, sString, sNewString)
'  S_ListToReconcile=Replace(S_ListToReconcile, sString, sNewString)
'End Sub

'Заменить роли в списке на пользователей из справочника ролей
Function ReplaceRoleFromDir(ByVal sUsersList, ByVal sDepartment)
Dim iPos1, iPos2
Dim sUser, sRole

'out "ReplaceRoleFromDir - initial list: "+sUsersList
  If sDepartment = SIT_SITRONICS Then
     sDirName = SIT_RolesDirSitronics
  ElseIf sDepartment = SIT_STS Then
     sDirName = SIT_RolesDirSTS
  ElseIf sDepartment = SIT_SIB Then
     sDirName = SIT_RolesDirSIB
  ElseIf sDepartment = SIT_RTI Then
     sDirName = SIT_RolesDirRTI
  ElseIf sDepartment = SIT_MIKRON Then
     sDirName = SIT_RolesDirMIKRON
  ElseIf sDepartment = SIT_MINC Then
     sDirName = SIT_RolesDirMinc    
  ElseIf sDepartment = SIT_VTSS Then
     sDirName = SIT_RolesDirVTSS    
  Else
     sDirName = SIT_RolesDirSitronics
  End If
'out "ReplaceRoleFromDir - sDepartment: "+sDepartment+"  sDirName: "+sDirName

  If Trim(sUsersList) = "" Then
     ReplaceRoleFromDir = sUsersList
     Exit Function
  End If

  iPos1 = 1
  ReplaceRoleFromDir = ""
  Do
     iPos2 = InStr(iPos1, sUsersList, """#")
     If iPos2 = 0 Then
        ReplaceRoleFromDir = ReplaceRoleFromDir + Mid(sUsersList, iPos1)
     Else
        ReplaceRoleFromDir = ReplaceRoleFromDir + Mid(sUsersList, iPos1, iPos2-iPos1)
        iPos1 = iPos2
        iPos2 = InStr(iPos1, sUsersList, """;")
        If iPos2 = 0 Then
           ReplaceRoleFromDir = ReplaceRoleFromDir + Mid(sUsersList, iPos1)
        Else
           sRole = Mid(sUsersList, iPos1, iPos2-iPos1+2)
           sUser = MyGetUserDirValue(sDirName, sRole, 1, 2)
           If sUser = "" Then
              sUser = sRole
           End If
'out "ReplaceRoleFromDir - Found role: '"+sRole+"'  replacing role by: '"+sUser+"'"
           ReplaceRoleFromDir = ReplaceRoleFromDir + sUser
           iPos1 = iPos2+2
        End If
     End If
  Loop Until iPos2 = 0
'out "ReplaceRoleFromDir - final list: "+ReplaceRoleFromDir
End Function

'Показ пиктограмм документа для отчета
Function MyShowPicts(DocID)
  Set dsDoc1 = Server.CreateObject("ADODB.Recordset")
  sSQL="select * from Docs where docid=N'"+DocID+"'"
  'MyShowPicts=sSQL
  dsDoc1.Open sSQL, Conn, 2, 1, &H1
  'ShowPicts dsDoc1

  If dsDoc.EOF Then
    Exit Function
  End If

  MyShowPicts=""

  'неактивен
  If dsDoc("IsActive")=VAR_InActiveTask and Not IsNull(dsDoc("IsActive")) Then
    MyShowPicts="<img border=""0"" src="""+GetPayDoxURL()+"IMAGES/Pict_Inactive.gif"" title="""+DOCS_Inactive+""" width=""16"" height=""16"">"
  End If

  'исполнен, просрочен
  If Not IsNull(dsDoc("StatusCompletion")) and dsDoc("StatusCompletion")<>"U" and Trim(dsDoc("StatusCompletion"))=VAR_StatusCompletion Then
    MyShowPicts=MyShowPicts+"<img border=""0"" src="""+GetPayDoxURL()+"IMAGES/Pict_completed.GIF"" title="""+DOCS_Completed+""" width=""16"" height=""16"">"
  ElseIf Not IsNull(dsDoc("StatusCompletion")) and dsDoc("StatusCompletion")<>"U" and Trim(dsDoc("StatusCompletion"))=VAR_StatusCancelled Then
    MyShowPicts=MyShowPicts+"<img border=""0"" src="""+GetPayDoxURL()+"IMAGES/Pict_canceled.GIF"" title="""+DOCS_Cancelled+""" width=""16"" height=""16"">"
  ElseIf Not IsNull(dsDoc("StatusCompletion")) and DateFullTime(dsDoc("DateCompletion"))<Now And Not IsNull(dsDoc("DateCompletion")) And dsDoc("DateCompletion")>VAR_BeginOfTimes Then
    MyShowPicts=MyShowPicts+"<img border=""0"" src="""+GetPayDoxURL()+"IMAGES/pict_expired.gif"" title="""+DOCS_EXPIRED2+""" width=""16"" height=""16"">"
  End If

  'контроль
  If Trim(MyCStr(dsDoc("NameControl")))<>"" Then
    If (MyCStr(dsDoc("StatusCompletion"))<>VAR_StatusCompletion And MyCStr(dsDoc("StatusCompletion"))<>VAR_StatusCancelled And IsNull(dsDoc("DateCompleted"))) Then
      MyShowPicts=MyShowPicts+"<img border=""0"" src="""+GetPayDoxURL()+"IMAGES/pict_UnderControl1.GIF"" title="""+DOCS_UnderControl1+" - "+HTMLEncode(Trim(MyCStr(dsDoc("NameControl"))))+""" width=""16"" height=""16"">"
    Else
      MyShowPicts=MyShowPicts+"<img border=""0"" src="""+GetPayDoxURL()+"IMAGES/pict_UnderControl3.GIF"" title="""+DOCS_UnderControl2+" - "+HTMLEncode(Trim(MyCStr(dsDoc("NameControl"))))+""" width=""16"" height=""16"">"
    End If
  End If

  If ShowStatusDevelopment(dsDoc("StatusDevelopment"))=DOCS_Approved Then
    MyShowPicts=MyShowPicts+"<img border=""0"" src="""+GetPayDoxURL()+"IMAGES/Pict_approved.gif"" title="""+DOCS_Approved+"&nbsp;"+MyDateShort(dsDoc("DateApproved"))+" - "+HTMLEncode(dsDoc("NameApproved"))+""" width=""16"" height=""16"">"
  Else
    If ShowStatusDevelopment(dsDoc("StatusDevelopment"))=DOCS_RefusedApp Then
      MyShowPicts=MyShowPicts+"<img border=""0"" src="""+GetPayDoxURL()+"IMAGES/Pict_Refused.GIF"" title="""+DOCS_RefusedApp+""" width=""16"" height=""16"">"
'    Else
    End If
  End If

  If MyCStr(dsDoc("StatusCompletion"))=VAR_StatusRequestCompletion Then
    MyShowPicts=MyShowPicts+"<img border=""0"" title="""+DOCS_RequestedCompleted+""" src="""+GetPayDoxURL()+"IMAGES/Pict_question.GIF"" width=""16"" height=""16"">"
  End If

  dsDoc1.Close
End Function

'Получить список всех комментариев о ходе исполнения
Function GetDocHistoryComments(ByVal sDocID)
  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  sSQL = "select UserName, DateCreation, Comment from comments where docid=N'"+sDocID+"' and CommentType='HISTORY' order by DateCreation"
AddlogD "GetDocHistoryComments SQL: "+sSQL
  dsTemp.Open sSQL, Conn, 3, 1, &H1
  GetDocHistoryComments = ""

  do while not dsTemp.EOF
    GetDocHistoryComments=GetDocHistoryComments+""+DelOtherLangFromFolder(dsTemp("UserName")) + "("+MyDateTime(dsTemp("DateCreation"))+"): " +Trim(MyCStr(DelOtherLangFromNames(dsTemp("Comment")))) + "<br>"
    dsTemp.MoveNext
  loop

  dsTemp.Close
End Function

'Аналог GetUserDirValue, не требующий oPaydox (для ReplaceRoleFromDir)
'ph - 20090705 - start
'Function MyGetUserDirValue(ByVal DirName, ByVal sKeyFieldValue, ByVal nKeyField, ByVal nValueField)
Function MyGetUserDirValue(ByVal DirName, ByVal sKeyFieldValue, ByVal parKeyField, ByVal parValueField)
'ph - 20090705 - end
  Dim MyConn
  Dim sConnStr
  Dim dsTemp
  Dim NewConnection

'ph - 20090705 - start
  MyGetUserDirValue = ""
  If IsNumeric(parKeyField) Then
    nKeyField = CInt(parKeyField)
    If nKeyField < 1 or nKeyField > 6 Then
      Exit Function
    End If
  Else
    Exit Function
  End If
  If IsNumeric(parValueField) Then
    nValueField = CInt(parValueField)
    If nValueField < 1 or nValueField > 6 Then
      Exit Function
    End If
  Else
    Exit Function
  End If
'ph - 20090705 - end

  NewConnection = True
  If not IsNull(Conn) Then
    NewConnection = not IsObject(Conn)
  End If
'out "NewConnection = "+CStr(NewConnection)

  If NewConnection Then
    SetPayDoxPars 'Присваиваем переменные среды из Global.asa
    sConnStr = "ConnectString"
    Select Case UCase(Request("l"))
      Case "RU" sConnStr = sConnStr + "RUS"
      Case "3" sConnStr = sConnStr + "3"
    End Select

    Set MyConn = Server.CreateObject("ADODB.Connection")
    MyConn.Open Application(sConnStr)
  Else
    Set MyConn = Conn
  End If

'on error resume next
'If Err.Number <> 0 Then
'    out "<font color = red>Error: "+ Err.Description + " (" + CStr(Err.Number) + ")</font>"
'End If
'On Error GoTo 0

  Set dsTemp = Server.CreateObject("ADODB.Recordset")
'Ph - 20090303 - GUID is NULL
'  sSQL = "Select UserDirValues.Field"+CStr(nValueField)+" as Value from UserDirValues inner join UserDirectories on (UserDirValues.GUIDUD = UserDirectories.GUID) where UserDirectories.Name = N'"+DirName+"' and UserDirValues.Field"+CStr(nKeyField)+" = N'"+sKeyFieldValue+"'"
'  sSQL = "Select UserDirValues.Field"+CStr(nValueField)+" as Value from UserDirValues inner join UserDirectories on (UserDirValues.UDKeyField = UserDirectories.KeyField) where UserDirectories.Name = N'"+DirName+"' and UserDirValues.Field"+CStr(nKeyField)+" = N'"+sKeyFieldValue+"'"
'ph - 20090705
  sSQL = "Select UserDirValues.Field"+CStr(nValueField)+" as Value from UserDirValues inner join UserDirectories on (UserDirValues.UDKeyField = UserDirectories.KeyField) where UserDirectories.Name = N'"+Replace(DirName,"'","''")+"' and UserDirValues.Field"+CStr(nKeyField)+" = N'"+Replace(sKeyFieldValue,"'","''")+"'"
'out "MyGetUserDirValue - SQL: " + sSQL
  dsTemp.Open sSQL, MyConn, 3, 1, &H1
'ph - 20090705 - start
'  If dsTemp.EOF Then
'    MyGetUserDirValue = ""
'  Else
  If not dsTemp.EOF Then
'ph - 20090705 - end
    MyGetUserDirValue = MyCStr(dsTemp("Value"))
  End If
  dsTemp.Close
'out "MyGetUserDirValue: "+MyGetUserDirValue

  If NewConnection Then
    MyConn.Close
  End If
End Function

'Получить пользователя по роли (для заявок)
Function GetRoleForOrders(ByVal sRole, ByVal sDepartment, ByVal sUserID, ByVal BusinessUnit)
'ph - 20081205 - start
  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  sSQL = "Select * from RolesForOrders_STS where Role="+sUnicodeSymbol+"'"+sRole+"'"
  'Ищем на полное соответствие БЕ
  sSQLAdd = " and BusinessUnit like "+sUnicodeSymbol+"'"+BusinessUnit+"%'"
AddLogD "GetRoleForOrders SQL1: "+sSQL+sSQLAdd
  dsTemp.Open sSQL+sSQLAdd, Conn, 3, 1, &H1
  If Not dsTemp.EOF Then
    If dsTemp.RecordCount = 1 Then
      GetRoleForOrders = GetInsertionNameInCurrentLanguage(dsTemp("Users"))
    Else
      GetRoleForOrders = ""
    End If
    Exit Function
  Else
    dsTemp.Close
    'Ищем на совпадение агрегатных БЕ
    sSQLAdd = " and BusinessUnit like "+sUnicodeSymbol+"'"+Left(BusinessUnit, 1)+"000%'"
AddLogD "GetRoleForOrders SQL2: "+sSQL+sSQLAdd
    dsTemp.Open sSQL+sSQLAdd, Conn, 3, 1, &H1
    If Not dsTemp.EOF Then
      If dsTemp.RecordCount = 1 Then
        GetRoleForOrders = GetInsertionNameInCurrentLanguage(dsTemp("Users"))
      Else
        GetRoleForOrders = ""
      End If
      Exit Function
    Else
      dsTemp.Close
      'Ищем по БЕ 9999
      sSQLAdd = " and BusinessUnit like "+sUnicodeSymbol+"'9999%'"
AddLogD "GetRoleForOrders SQL3: "+sSQL+sSQLAdd
      dsTemp.Open sSQL+sSQLAdd, Conn, 3, 1, &H1
      If Not dsTemp.EOF Then
        If dsTemp.RecordCount = 1 Then
          GetRoleForOrders = GetInsertionNameInCurrentLanguage(dsTemp("Users"))
        Else
          GetRoleForOrders = ""
        End If
        Exit Function
      Else
        GetRoleForOrders = ""
      End If	  
    End If	  
  End If	  
  dsTemp.Close
'ph - 20081205 - end
  If GetRoleForOrders <> "" Then
    Exit Function
  End If

  'Поиск  по ролям руководителей подразделений
  If InStr(UCase(Trim(sRole)), UCase(STS_Orders_HeadOfSector)) = 1 Then
    iLevel = 3
  ElseIf InStr(UCase(Trim(sRole)), UCase(STS_Orders_HeadOfDepartment)) = 1 Then
    iLevel = 2
  ElseIf InStr(UCase(Trim(sRole)), UCase(STS_Orders_HeadOfDivision)) = 1 Then
    iLevel = 1
  Else iLevel = 0
  End If
  'Неизвестная роль
  If iLevel = 0 Then
    GetRoleForOrders = sRole
  Else
'    GetRoleForOrders = GetChiefNameUpperWithDepLevelSTSNew(sDepartment, iLevel, sUserID, BusinessUnit)
    GetRoleForOrders = GetChiefOfDepUpperByLevelSys(sDepartment, iLevel, sUserID, BusinessUnit)
  End If
End Function

'Предыдущая функция с обработкой ошибок
Function GetRoleForOrders_WithCheck(ByVal sRole, ByVal sDepartment, ByVal sUserID, ByVal BusinessUnit, bErrorFlag)
Dim iPos, tempGetRoleForOrders
  bErrorFlag = False
  tempGetRoleForOrders = GetRoleForOrders(sRole, sDepartment, sUserID, BusinessUnit)
  GetRoleForOrders_WithCheck = tempGetRoleForOrders
  If GetRoleForOrders_WithCheck = "" Then
	bErrorFlag = True
    GetRoleForOrders_WithCheck = sRole
  ElseIf GetRoleForOrders_WithCheck = "-" Then 'Текущее подразделение более высокого уровня
    GetRoleForOrders_WithCheck = ""
  ElseIf GetRoleForOrders_WithCheck = "=" Then 'Пользователь совпадает с sUserID
    GetRoleForOrders_WithCheck = ""
  Else
	iPos = 1
	If oPayDox.GetNextUserIDInList(GetRoleForOrders_WithCheck, iPos) = "" Then 'Нет логина - некорректные данные
      bErrorFlag = True
      GetRoleForOrders_WithCheck = sRole
	End If
  End If
  If bErrorFlag Then
AddlogD "ROLE NOT FOUND --- Role = """+sRole+""", Department = """+sDepartment+""", UserID = """+sUserID+""", BusinessUnit = "+BusinessUnit+""", ReturnedRole = """+tempGetRoleForOrders+""""
'    Session("Message") =  Session("Message")+VbCrLf+"<font color = red>ROLE NOT FOUND:</font> """+sRole+""""+VbCrLf+"Department: """+sDepartment+""""+VbCrLf+"UserID: """+sUserID+""""++VbCrLf+"BusinessUnit: """+BusinessUnit+""""+VbCrLf+"ReturnedRole: """+tempGetRoleForOrders+""""
  End If
End Function

'Получение расшифровки роли согласующего для заявок
'UseLang = True - расшифровка идет на текущем языке
'UseLang = False - расшифровка идет на английском языке
Function GetOrderRoleDescription(ByVal parRole, ByVal parDepartment, ByVal UseLang, parBusinessUnit)
  Dim sRole

  If SIT_ShowOrdersAgreesDescription = "" Then
    GetOrderRoleDescription = ""
	Exit Function
  End If

  If UseLang Then
    Select case parRole
      case STS_Orders_HeadOfSector
        sRole = STS_HeadOfSector
      case STS_Orders_HeadOfDepartment
        sRole = STS_HeadOfDepartment
      case STS_Orders_HeadOfDivision
        sRole = STS_HeadOfDivision
      case STS_Orders_FinancialControl
        sRole = STS_FinancialControl
      case STS_Orders_FinDirector
        sRole = STS_FinDirector
      case STS_Orders_GenDirector
        sRole = STS_GenDirector
      case STS_Orders_ProjectManager
        sRole = STS_ProjectManager
      case STS_Orders_Accounting
        sRole = STS_Accounting
      case STS_Orders_Treasury
        sRole = STS_Treasury
      case else
        sRole = "UNKNOWN"
    End Select
  Else
    sRole = parRole
  End If

  GetOrderRoleDescription = " [" + Replace(Replace(sRole, """#", ""), """;", "")
  If parRole <> STS_Orders_ProjectManager Then
    GetOrderRoleDescription = GetOrderRoleDescription + " - Dep:" + GetSTSDepartmentCode(parDepartment)
  End If
  If Trim(parBusinessUnit) <> "" Then
    GetOrderRoleDescription = GetOrderRoleDescription + " BU:" + Left(parBusinessUnit, InStr(parBusinessUnit+" ", " ")-1)
  End If
  GetOrderRoleDescription = GetOrderRoleDescription + "]"
End Function

'Добавление пользователя, соответствующего роли, в список согласования с расшифровкой (если включена), плюс ведение полного списка согласования (кидаются расшифровки, даже если пользователь не добавлен)
Function AddUserToReconcileList2(ByVal sListToReconcile, ByVal sDelimeter, ByVal sRole, ByVal sDepartment, ByVal sUserID, ByVal sBusinessUnit, bErrorFlag, parFullListToReconcile)
  Dim sAgreePerson, sAgreePersonDescription
  If sRole = STS_Orders_ProjectManager Then
    sAgreePerson = GetInsertionNameInCurrentLanguage(sDepartment)
    'rmanyushin 136964 08.11.2010 Start
	ElseIf sRole = STS_Overtime_Requester Then
		sAgreePerson = GetInsertionNameInCurrentLanguage(sUserID)
	'rmanyushin 136964 08.11.2010 End
  Else
    sAgreePerson = GetRoleForOrders_WithCheck(sRole, sDepartment, sUserID, sBusinessUnit, bErrorFlag)
  End If
  
  
'rmanyushin 136964 08.11.2010 Start
  If sRole = STS_DirectorOfDirection Then
    sAgreePerson = GetInsertionNameInCurrentLanguage(GetSTSDirectorOfDirection(sDepartment))
  AddLogD "@@@OVERTIME2Reconcilation - STS_DirectorOfDirection: "+sAgreePerson
  End If
  
  If sRole = STS_AssistantDirector Then
    sAgreePerson = GetInsertionNameInCurrentLanguage(GetSTSAssistantDirector(sDepartment))
  AddLogD "@@@OVERTIME2Reconcilation - STS_AssistantDirector: "+sAgreePerson
  End If
 
  If sRole = STS_SecurityManager Then
    sAgreePerson = GetInsertionNameInCurrentLanguage(GetSTSSecurityManager(sDepartment))
  AddLogD "@@@OVERTIME2Reconcilation - STS_SecurityManager: "+sAgreePerson
  End If
 
  If sRole = STS_HRDirector Then
    sAgreePerson = GetInsertionNameInCurrentLanguage(GetSTSHRDirector(sDepartment))
  AddLogD "@@@OVERTIME2Reconcilation - STS_HRDirector: "+sAgreePerson
  End If
'rmanyushin 136964 08.11.2010 Start
  
  sAgreePersonDescription = GetOrderRoleDescription(sRole, sDepartment, False, sBusinessUnit)
  parFullListToReconcile = parFullListToReconcile+sDelimeter+sAgreePerson+GetOrderRoleDescription(sRole, sDepartment, True, sBusinessUnit)
  If Trim(sAgreePerson) <> "" Then
    AddUserToReconcileList2 = AddUserToReconcileList(S_ListToReconcile, sUserID, sDelimeter, sAgreePerson+sAgreePersonDescription)
  Else
    AddUserToReconcileList2 = S_ListToReconcile
  End If
End Function

'ph - 20090704 - start
'Удаление пользователя в стандартном формате с имененем и логином из списка (используется при корректировке листа согласования Заявок)
'Удаляется при полном совпадении имени (не только логина)
'Function RemoveUserFromListWithDescriptions(ByVal sUserList, ByVal sRemovingUser)
'  Dim iPos, iQuotPos, iLBracketPos, iRBracketPos
'  Dim arList, i
'
'  RemoveUserFromListWithDescriptions = sUserList
'  Do While True
'    iPos = InStr(RemoveUserFromListWithDescriptions, sRemovingUser)
'    If iPos = 0 Then
'      'Удаляем пустые строки
'      If Trim(RemoveUserFromListWithDescriptions) <> "" Then
'        arList = Split(RemoveUserFromListWithDescriptions, VbCrLf)
'        RemoveUserFromListWithDescriptions = ""
'        For i = LBound(arList) To UBound(arList)
'          If Trim(arList(i)) <> "" Then
'            If RemoveUserFromListWithDescriptions <> "" Then
'              RemoveUserFromListWithDescriptions = RemoveUserFromListWithDescriptions + VbCrLf
'            End If
'            RemoveUserFromListWithDescriptions = RemoveUserFromListWithDescriptions + arList(i)
'          End If
'        Next
'      End If
'	  Exit Function
'    End If
'    RemoveUserFromListWithDescriptions = Replace(RemoveUserFromListWithDescriptions, sRemovingUser, "", 1, 1)
'    iQuotPos = InStr(iPos, RemoveUserFromListWithDescriptions, """")
'    iLBracketPos = InStr(iPos, RemoveUserFromListWithDescriptions, "[")
'    If iLBracketPos <> 0 and iLBracketPos < iQuotPos Then 'Расшифровка у текущего пользователя есть
'      iRBracketPos = InStr(iPos, RemoveUserFromListWithDescriptions, "]")
'      If iRBracketPos > iLBracketPos and iRBracketPos < iQuotPos Then 'Ошибок нет, вырезаем расшифровку
'        RemoveUserFromListWithDescriptions = Left(RemoveUserFromListWithDescriptions, iPos-1)+Mid(RemoveUserFromListWithDescriptions, iRBracketPos+1)
'      End If
'    End If
'  Loop
'End Function

'Удаление пользователя в стандартном формате с имененем и логином из списка (используется при корректировке листа согласования Заявок)
'Удаляется при полном совпадении имени (не только логина)
Function RemoveUserFromListWithDescriptions(ByVal sUserList, ByVal parRemovingUser)
  Dim iPos, iQuotPos, iLBracketPos, iRBracketPos
  Dim arList, i

  sRemovingUser = parRemovingUser
  iLBracketPos = InStr(sRemovingUser, "[")
  If iLBracketPos <> 0 Then 'есть расшифровка
    iRBracketPos = InStr(sRemovingUser, "]")
    If iRBracketPos > iLBracketPos Then 'скобки в верной последовательности - отрезаем расшифровку
      sRemovingUser = Trim(Left(sRemovingUser, iLBracketPos-1)+Mid(sRemovingUser, iRBracketPos+1))
    End If
  End If
  RemoveUserFromListWithDescriptions = sUserList
  Do While True
    iPos = InStr(RemoveUserFromListWithDescriptions, sRemovingUser)
    If iPos = 0 Then
      'Удаляем пустые строки
      If Trim(RemoveUserFromListWithDescriptions) <> "" Then
        arList = Split(RemoveUserFromListWithDescriptions, VbCrLf)
        RemoveUserFromListWithDescriptions = ""
        For i = LBound(arList) To UBound(arList)
          If Trim(arList(i)) <> "" Then
            If RemoveUserFromListWithDescriptions <> "" Then
              RemoveUserFromListWithDescriptions = RemoveUserFromListWithDescriptions + VbCrLf
            End If
            RemoveUserFromListWithDescriptions = RemoveUserFromListWithDescriptions + arList(i)
          End If
        Next
      End If
	  Exit Function
    End If
    RemoveUserFromListWithDescriptions = Replace(RemoveUserFromListWithDescriptions, sRemovingUser, "", 1, 1)
    iQuotPos = InStr(iPos, RemoveUserFromListWithDescriptions, """")
    iLBracketPos = InStr(iPos, RemoveUserFromListWithDescriptions, "[")
    If iLBracketPos <> 0 and (iLBracketPos < iQuotPos or iQuotPos = 0) Then 'Расшифровка у текущего пользователя есть
      iRBracketPos = InStr(iPos, RemoveUserFromListWithDescriptions, "]")
      If iRBracketPos > iLBracketPos and (iRBracketPos < iQuotPos or iQuotPos = 0) Then 'Ошибок нет, вырезаем расшифровку
        RemoveUserFromListWithDescriptions = Left(RemoveUserFromListWithDescriptions, iPos-1)+Mid(RemoveUserFromListWithDescriptions, iRBracketPos+1)
      End If
    End If
  Loop
End Function
'ph - 20090704 - end

'ph - 20100623 - start - больше не используется
'Возвращает сокращенное имя на нужном языке
'Lang - языковой параметр в формате Request("l")
'Function SurnameGNLang(ByVal Par, ByVal Lang)

  ' Dim i1, s, sSuffix, s0, cPar
  
  ' sSuffix = ""
  ' cPar = Trim(MyCStr(Par))
  ' SurnameGNLang = cPar
  ' If cPar = "" Then
    ' Exit Function
  ' End If

  ' i1 = InStr(cPar, "*")
  ' If i1 > 1 Then
    ' CheckNamesByLang cPar, sName1, sName2, sName3
	' select case UCase(Lang)
	  ' case "RU" cPar = sName1
	  ' case "" cPar = sName2
	  ' case "3" cPar = sName3
	' end select
  ' End If

  ' If (UCase(Lang) <> "RU" And i1 <= 0) Or Trim(VAR_NotToUseNameInitials) <> "" Then
    ' Exit Function
  ' End If

  ' i1 = InStr(cPar, ",")
  ' If i1 > 1 Then
    ' sSuffix = Mid(cPar, i1)
    ' cPar = Left(cPar, i1 - 1)
  ' End If
  ' cPar = GetName(cPar)
  ' SurnameGNLang = cPar + sSuffix
  ' If cPar = "" Then
    ' SurnameGNLang = cPar
    ' Exit Function
  ' End If
  ' If InStr(cPar, ".") > 0 Then
    ' SurnameGNLang = cPar
    ' Exit Function
  ' End If
  ' i1 = InStr(cPar, " ")
  ' If i1 <= 0 Then
    ' SurnameGNLang = cPar
    ' Exit Function
  ' End If
  ' i2 = InStr(i1 + 1, cPar + "*", " ")
  ' If i2 <= 0 Then
    ' SurnameGNLang = cPar
    ' Exit Function
  ' End If
  ' s = Trim(Left(cPar, i1))
  ' If s = "" Then
    ' SurnameGNLang = cPar
    ' Exit Function
  ' End If
  ' cPar = Trim(Mid(cPar, i1))
  ' If cPar = "" Then
    ' SurnameGNLang = s + sSuffix
    ' Exit Function
  ' End If
  ' SurnameGNLang = s
  ' i1 = InStr(cPar, " ")
  ' If i1 <= 0 Then
    ' SurnameGNLang = Left(cPar, 1) + ". " + SurnameGNLang + sSuffix
    ' Exit Function
  ' End If
  ' s = Trim(Mid(cPar, i1))
  ' If s = "" Then
    ' SurnameGNLang = Left(cPar, 1) + ". " + SurnameGNLang + sSuffix
    ' Exit Function
  ' End If
  ' SurnameGNLang = Left(cPar, 1) + ". " + Left(s, 1) + ". " + SurnameGNLang + sSuffix
'End Function
'ph - 20100623 - end

'Проверка однопользовательского поля на соответствие справочнику
Function CheckSingleUserField(ByVal sUser)
  Dim sUserID
  Dim sName1, sName2, sName3
'ph - 20100623 - start
  Dim Save_VAR_SurnameGN
'ph - 20100623 - end

AddLogD "CheckSingleUserField - sUser: "+MyCStr(sUser)
'amw debug
  CheckSingleUserField = True
'amw debug

  If Trim(sUser) = "" Then
     CheckSingleUserField = True
'amw 1 <start>
     Exit Function
  End If
'amw 1 <end> 
'  Else
'amw 1 <end> 
     If InStr(sUser, """#") = 1 Then
        CheckSingleUserField = CheckRoleExistence(Trim(sUser))
'amw 2 <start>
        Exit Function
     End If
'amw 2 <end>   
'     Else
'amw 2 <end>   
        sUserID = GetUserID(sUser)
        If sUserID = "" Then
           CheckSingleUserField = False
'amw 3 <start>
           Exit Function
        End If
'amw 3 <end>       
'        Else
'amw 3 <end>       
           oPayDox.GetUserDetails sUserID, sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
'Ph - 20090312 - start
           If InStr(sName, "*") > 1 Then
              CheckNamesByLang sName, sName1, sName2, sName3
           Else
              sName1 = sName
              sName2 = sName
              sName3 = sName
           End If
'ph - 20100623 - start

''Ph - 20090312 - end
'        If GetFullName(SurnameGNLang(sName1, "RU"), sUserID)+";" = Trim(sUser) Then
'          CheckSingleUserField = True
'        ElseIf GetFullName(SurnameGNLang(sName2, ""), sUserID)+";" = Trim(sUser) Then
'          CheckSingleUserField = True
'        ElseIf GetFullName(SurnameGNLang(sName3, "3"), sUserID)+";" = Trim(sUser) Then
'          CheckSingleUserField = True
''Ph - 20090312 - start
'        ElseIf GetFullName(SurnameGN(sName), sUserID)+";" = Trim(sUser) Then
'          CheckSingleUserField = True
''Ph - 20090312 - end
''Ph - 20090313 - start - Вариант возвращаемый списком подчиненных в русском интерфейсе
'        ElseIf GetFullName(sName1, sUserID)+";" = Trim(sUser) Then
'          CheckSingleUserField = True
''Ph - 20090313 - end
'        Else
'          CheckSingleUserField = False
'        End If

           CheckSingleUserField = False
		   Save_VAR_SurnameGN = oPayDox.VAR_SurnameGN

		   oPayDox.VAR_SurnameGN = ""
           'Вариант возвращаемый списком подчиненных в русском интерфейсе
           If GetFullName(sName1, sUserID)+";" = Trim(sUser) Then
              CheckSingleUserField = True
		   End If
		   If not CheckSingleUserField Then
              oPayDox.VAR_CurrentL = "RU"
              If GetFullName(SurnameGN(sName1), sUserID)+";" = Trim(sUser) Then
                 CheckSingleUserField = True
              End If
		   End If
		   If not CheckSingleUserField Then
              If GetFullName(SurnameGN(sName), sUserID)+";" = Trim(sUser) Then
                 CheckSingleUserField = True
              End If
		   End If
		   If not CheckSingleUserField Then
              oPayDox.VAR_CurrentL = ""
              If GetFullName(SurnameGN(sName2), sUserID)+";" = Trim(sUser) Then
                 CheckSingleUserField = True
              End If
		   End If
		   If not CheckSingleUserField Then
              If GetFullName(SurnameGN(sName), sUserID)+";" = Trim(sUser) Then
                 CheckSingleUserField = True
              End If
		   End If
           If not CheckSingleUserField Then
              oPayDox.VAR_CurrentL = "3"
              If GetFullName(SurnameGN(sName3), sUserID)+";" = Trim(sUser) Then
                 CheckSingleUserField = True
              End If
           End If
		   If not CheckSingleUserField Then
              If GetFullName(SurnameGN(sName), sUserID)+";" = Trim(sUser) Then
                 CheckSingleUserField = True
              End If
		   End If

           oPayDox.VAR_SurnameGN = "2"
		   If not CheckSingleUserField Then
              oPayDox.VAR_CurrentL = "RU"
              If GetFullName(SurnameGN(sName1), sUserID)+";" = Trim(sUser) Then
                 CheckSingleUserField = True
              End If
		   End If
		   If not CheckSingleUserField Then
              If GetFullName(SurnameGN(sName), sUserID)+";" = Trim(sUser) Then
                 CheckSingleUserField = True
              End If
		   End If
		   If not CheckSingleUserField Then
              oPayDox.VAR_CurrentL = ""
              If GetFullName(SurnameGN(sName2), sUserID)+";" = Trim(sUser) Then
                 CheckSingleUserField = True
              End If
		   End If
		   If not CheckSingleUserField Then
              If GetFullName(SurnameGN(sName), sUserID)+";" = Trim(sUser) Then
                 CheckSingleUserField = True
              End If
		   End If
		   If not CheckSingleUserField Then
              oPayDox.VAR_CurrentL = "3"
              If GetFullName(SurnameGN(sName3), sUserID)+";" = Trim(sUser) Then
                 CheckSingleUserField = True
              End If
		   End If
		   If not CheckSingleUserField Then
              If GetFullName(SurnameGN(sName), sUserID)+";" = Trim(sUser) Then
                 CheckSingleUserField = True
              End If
		   End If
		  
		   oPayDox.VAR_CurrentL = Request("l")
		   oPayDox.VAR_SurnameGN = Save_VAR_SurnameGN
'ph - 20100623 - end
'amw 3      
'        End If
'amw 2    
'     End If
'amw 1 
'  End If

'AddLogD "CheckSingleUserField - result: "+CStr(CheckSingleUserField)
End Function

Function CheckRoleExistence(sRole)
  Dim sSQL, dsTemp
  sSQL = "select UserDirValues.Field1 from (UserDirectories Left Outer Join UserDirValues ON UserDirValues.UDKeyField = UserDirectories.KeyField) where Field1 = "+sUnicodeSymbol+"'"+sRole+"'"
'Запрос №1 - СИБ - start
  'sSQL = sSQL + "and UserDirectories.Name in ("+sUnicodeSymbol+"'"+SIT_RolesDirSitronics_RU+"',"+sUnicodeSymbol+"'"+SIT_RolesDirSTS_RU+"',"+sUnicodeSymbol+"'"+SIT_RolesDirSitronics_EN+"',"+sUnicodeSymbol+"'"+SIT_RolesDirSTS_EN+"',"+sUnicodeSymbol+"'"+SIT_RolesDirSitronics_CZ+"',"+sUnicodeSymbol+"'"+SIT_RolesDirSTS_CZ+"')" 'SIT_RolesDirSIB
  sSQL = sSQL + "and UserDirectories.Name in ("+sUnicodeSymbol+"'"+SIT_RolesDirSitronics_RU+"',"+sUnicodeSymbol+"'"+SIT_RolesDirSTS_RU++"',"+sUnicodeSymbol+"'"+SIT_RolesDirSIB_RU+"',"+sUnicodeSymbol+"'"+SIT_RolesDirSitronics_EN+"',"+sUnicodeSymbol+"'"+SIT_RolesDirSTS_EN+"',"+sUnicodeSymbol+"'"+SIT_RolesDirSIB_EN+"',"+sUnicodeSymbol+"'"+SIT_RolesDirSitronics_CZ+"',"+sUnicodeSymbol+"'"+SIT_RolesDirSTS_CZ+"',"+sUnicodeSymbol+"'"+SIT_RolesDirSIB_CZ+"')"
'Запрос №1 - СИБ - end
AddLogD "CheckRoleExistence - sSQL: " + sSQL
  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  dsTemp.Open sSQL, Conn, 3, 1, &H1
  CheckRoleExistence = not dsTemp.EOF
  dsTemp.Close
AddLogD "CheckRoleExistence("+MyCStr(sRole)+") = " + CStr(CheckRoleExistence)
End Function

'Проверка многопользовательского поля
Function CheckMultiUserField(ByVal sUserList)
  Dim sList

AddLogD "CheckMultiUserField - sUserList: "+MyCStr(sUserList)
  CheckMultiUserField = Trim(sUserList) = "" or oPayDox.GetNextUserIDInList(sUserList, 1) <> "" or InStr(sUserList, """#") > 0
  'Список не пуст и в нем нет логинов и ролей
  If not CheckMultiUserField Then
    'Удаляем приписки об обязательных/дополнительных согласующих,
    sList = Replace(sUserList, SIT_AdditionalAgrees, "")
    sList = Replace(sList, SIT_RequiredAgrees, "")
    'отделитель дополнительных согласующих
    sList = Replace(sList, SIT_AdditionalAgreesDelimeter, "")
    'кавычки
    sList = Replace(sList, """", "")
    'Если ничего значимого не осталось, значит список пуст, проверка пройдена
    CheckMultiUserField = Trim(sList) = ""
  End If
AddLogD "CheckMultiUserField - result: "+CStr(CheckMultiUserField)
End Function

'Устаревшая, нужно от нее уходить
Function GetUserLanguage(ByVal sUser)
  Dim dsTemp
  Dim sUserID, sDomain, iPos
  
  'По умолчанию англ. язык
  GetUserLanguage = ""
  
  sUserID = Trim(sUser)
  If InStr(sUserID, "<") > 0 Then
    sUserID = GetUserID(sUserID)
  End If
  If sUserID = "" Then
    Exit Function
  End If
  
  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  sSQL = "select WinLogin from Users where UserID = "+sUnicodeSymbol+"'"+sUserID+"'"
AddLogD "GetUserLanguage - sSQL: " + sSQL
  dsTemp.Open sSQL, Conn, 3, 1, &H1
  If not dsTemp.EOF Then
    sDomain = MyCStr(dsTemp("WinLogin"))
  Else
    sDomain = ""
  End If
  dsTemp.Close
AddLogD "GetUserLanguage - WinLogin: " + sDomain

  iPos = InStr(sDomain, "\")
  If iPos > 0 Then
    sDomain = Left(sDomain, iPos-1)
  Else
    sDomain = ""
  End If
AddLogD "GetUserLanguage - Domain: " + sDomain
  
  Select case UCase(sDomain)
    case "ROOT"
	  GetUserLanguage = "RU"
    case "GLOBAL"
	  GetUserLanguage = "RU"
    case "STROMTELECOM"
	  GetUserLanguage = "3"
    case else
	  GetUserLanguage = ""
  End Select
AddLogD "GetUserLanguage: " + GetUserLanguage
End Function

'Получить список ролей и соответствующих пользователей
'(список возвращаемого формата используется функцией ReplaceRolesInList)
Function GetRolesList(ByVal sDepartment, ByVal sUser, ByVal sBusinessUnit)
  Dim MyConn, sConnStr, dsTemp
  Dim NewConnection, i
  Dim sRole, sUsers

  NewConnection = True
  If not IsNull(Conn) Then
    NewConnection = not IsObject(Conn)
  End If

  If NewConnection Then
    SetPayDoxPars 'Присваиваем переменные среды из Global.asa
    sConnStr = "ConnectString"
    Select Case UCase(Request("l"))
      Case "RU" sConnStr = sConnStr + "RUS"
      Case "3" sConnStr = sConnStr + "3"
    End Select

    Set MyConn = Server.CreateObject("ADODB.Connection")
    MyConn.Open Application(sConnStr)
  Else
    Set MyConn = Conn
  End If

  ' DmGorsky_5: меняем местами проверку: вначале ищем SIT_SITRU ("СИТРОНИКС ИТ"), потом, если не найдено, - SIT_SITRONICS ("СИТРОНИКС")
  If InStr(sDepartment, SIT_SITRU) = 1 Then ' DmGorsky_5
    sDirName = SIT_RolesDirSITRU_RU ' DmGorsky_5
  ElseIf InStr(sDepartment, SIT_SITRONICS) = 1 Then ' DmGorsky_5
    sDirName = SIT_RolesDirSitronics
  ElseIf InStr(sDepartment, SIT_STS) = 1 Then
    sDirName = SIT_RolesDirSTS
'Запрос №1 - СИБ - start
  ElseIf InStr(sDepartment, SIT_SIB_ROOT_DEPARTMENT) = 1 Then
    sDirName = SIT_RolesDirSIB
'Запрос №1 - СИБ - end
  ElseIf InStr(sDepartment, SIT_RTI) = 1 Then
    sDirName = SIT_RolesDirRTI
  ElseIf InStr(sDepartment, SIT_MIKRON) = 1 Then
    sDirName = SIT_RolesDirMIKRON
  ElseIf InStr(sDepartment, SIT_MINC) = 1 Then
    sDirName = SIT_RolesDirMinc
  ElseIf InStr(sDepartment, SIT_VTSS) = 1 Then
    sDirName = SIT_RolesDirVTSS
  Else
    sDirName = SIT_RolesDirSitronics
  End If

  GetRolesList = ""

  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  sSQL = "Select UserDirValues.Field1,UserDirValues.Field2 from UserDirValues inner join UserDirectories on (UserDirValues.UDKeyField = UserDirectories.KeyField) where UserDirectories.Name = "+sUnicodeSymbol+"'"+sDirName+"'"
AddLogD "GetRolesList - SQL: "+sSQL

  dsTemp.Open sSQL, MyConn, 3, 1, &H1
  If not dsTemp.EOF Then
    i = 0
    Do While not dsTemp.EOF
      sRole = Trim(MyCStr(dsTemp("Field1")))
      Select case sRole
        case SIT_HeadOfInitiatorsUnit
          sUsers = GetChiefOfDepUpperByLevel(sDepartment, 3, sUser, sBusinessUnit)
        case RTI_HeadOfInitiatorsUnit
          sUsers = GetNearestChief(sDepartment, sUser, sBusinessUnit)
        case MINC_HeadOfInitiatorsUnit
          sUsers = GetNearestChief(sDepartment, sUser, sBusinessUnit)
        case VTSS_HeadOfInitiatorsUnit
          sUsers = GetNearestChief(sDepartment, sUser, sBusinessUnit)
        case MIKRON_HeadOfInitiatorsUnit
          sUsers = GetNearestChief(sDepartment, sUser, sBusinessUnit)
        case SIT_DirectorOfInitiatorsDepartment
          sUsers = GetChiefOfDepUpperByLevel(sDepartment, 2, sUser, sBusinessUnit)
        case SIT_DirectorOfInitiatorsDivision
          sUsers = GetChiefOfDepUpperByLevel(sDepartment, 1, sUser, sBusinessUnit)
        case SIT_VicePresidentOfInitiator
          sUsers = GetChiefOfDepUpperByLevel(sDepartment, 1, sUser, sBusinessUnit)
'Запрос №1 - СИБ - start
        case SIB_AssistantDirector
          sUsers = GetChiefOfDepUpperByLevel(sDepartment, 1, sUser, sBusinessUnit)
        case SIB_HeadOfDepartment
          sUsers = GetChiefOfDepUpperByLevel(sDepartment, 2, sUser, sBusinessUnit)
        case SIB_HeadOfSector
          sUsers = GetChiefOfDepUpperByLevel(sDepartment, 3, sUser, sBusinessUnit)
'Запрос №1 - СИБ - end
        'rmanyushin 119579 19.08.2010 Start
		case STS_HeadOfInitiatorsGroup
          sUsers = GetChiefOfDepUpperByLevel(sDepartment, 4, sUser, sBusinessUnit)
		case STS_AssistantDirector
          sUsers = GetInsertionNameInCurrentLanguage(GetSTSAssistantDirector(sDepartment))
		case STS_DirectorOfDirection
          sUsers = GetInsertionNameInCurrentLanguage(GetSTSDirectorOfDirection(sDepartment))
		'rmanyushin 119579 19.08.2010 End
        case else
          sUsers = Trim(MyCStr(dsTemp("Field2")))
      End Select
AddLogD "GetRolesList - "+CStr(i)+"  Role: "+sRole+"  Value: "+sUsers
      If GetRolesList <> "" Then
        GetRolesList = GetRolesList + VbCrLf
      End If
      GetRolesList = GetRolesList + sRole + "|" + sUsers
      dsTemp.MoveNext
      i = i+1
    Loop
  End If
  dsTemp.Close

  If NewConnection Then
    MyConn.Close
  End If
End Function

'Заменить роли в списке на соответствующих пользователей.
'Данные о ролях передаются параметром sRoles, который нужно получить вызовом GetRolesList
Function ReplaceRolesInList(ByVal sList, ByVal sRoles)
  Dim i, arRoles, arVals
  
  arRoles = Split(sRoles, VbCrLf)
  ReplaceRolesInList = sList

  If InStr(ReplaceRolesInList, """#") = 0 Then
    'Нет ни одной роли
    Exit Function
  End If

  For i = 0 to UBound(arRoles)
    arVals = Split(arRoles(i), "|")
    ReplaceRolesInList = Replace(ReplaceRolesInList, arVals(0), arVals(1))
  Next
End Function

'Получить название корневого подразделения на русском языке
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
  iPos = InStr(GetRootDepartment, "*")
  If iPos > 0 Then
    GetRootDepartment = Left(GetRootDepartment, iPos-1)
  End If
End Function

'Получить код подразделения СТС
Function GetSTSDepartmentCode(ByVal sDepartment)
  Dim iPos

  GetSTSDepartmentCode = Trim(sDepartment)
  If GetSTSDepartmentCode = "" Then
    Exit Function
  End If

  If Right(GetSTSDepartmentCode, 1) = "/" Then
    GetSTSDepartmentCode = Left(GetSTSDepartmentCode, Len(GetSTSDepartmentCode)-1)
  End If
  iPos = InStrRev(GetSTSDepartmentCode, "/")
  If iPos > 0 Then
    GetSTSDepartmentCode = Right(GetSTSDepartmentCode, Len(GetSTSDepartmentCode)-iPos)
  End If
  iPos = InStr(GetSTSDepartmentCode, " ")
  If iPos > 0 Then
    GetSTSDepartmentCode = Left(GetSTSDepartmentCode, iPos-1)
  End If
  If InStr("0123456789", Left(GetSTSDepartmentCode, 1)) = 0  Then
    GetSTSDepartmentCode  = ""
  End If
End Function

''Определение бизнес направления
'Function GetDepartmentRoot(ByVal sDepartment)
'  If InStr(UCase(sDepartment), UCase(SIT_SITRONICS)) = 1 Then
'    GetDepartmentRoot = SIT_SITRONICS
'  ElseIf InStr(UCase(sDepartment), UCase(SIT_STS)) = 1 Then
'    GetDepartmentRoot = SIT_STS
'  Else
'    GetDepartmentRoot = ""
'  End If
'End Function


'Функция SQL для вырезания логина из полного имени (только для MSSQL)
Function GetLoginSQL_MSSQL(ByVal sFieldName)
'  GetLoginSQL_MSSQL = "SubString("+sFieldName+", CharIndex(N'<', "+sFieldName+")+1, Case CharIndex(N'>', "+sFieldName+")-CharIndex(N'<', "+sFieldName+") When 0 Then 0 Else CharIndex(N'>', "+sFieldName+")-CharIndex(N'<', "+sFieldName+")-1 End)"
  GetLoginSQL_MSSQL = "SubString("+sFieldName+", CharIndex(N'<', "+sFieldName+")+1, Case When CharIndex(N'>', "+sFieldName+")-CharIndex(N'<', "+sFieldName+") <= 0 Then 0 Else CharIndex(N'>', "+sFieldName+")-CharIndex(N'<', "+sFieldName+")-1 End)"
End Function

' ------------------- Формирование отчета "Анализ активности пользователей за месяц" -------------------- START
'Накопительная переменная для подсчета Суммы по строкам
iReportUsersActivityRowSum = 0

'Функция формирующая SQL-запрос для ячеек отчета и возвращающая его результат
Function ReportUsersActivity(ByVal RequestType, ByVal DomainType, ByVal parDay, ByVal parMonth, ByVal parYear, ByVal sClassDoc)
  Dim sDateField, sDomainAdd, sDateAdd, sSQL
  Dim sDay, sMonth, sYear, iYear
  Dim dsTemp
  
  'Обнуляем сумму
  If CInt(parDay) = 1 Then
    iReportUsersActivityRowSum = 0
  End If
  
  'Обработка входящих параметров даты
  On Error Resume Next
  iYear = CInt(parYear)
  If Err.Number <> 0 Then
    ReportUsersActivity = "<font color = red>Error: "+ Err.Description + " (" + CStr(Err.Number) + ")</font>"
    On Error GoTo 0
    Exit Function
  End If
  On Error GoTo 0
  If iYear < 100 Then
    iYear = 2000+iYear
  End If
  If iYear < 2008 or iYear > Year(Date) Then
    ReportUsersActivity = ""
    Exit Function
  End If
  sYear = CStr(iYear)

  sMonth = CStr(parMonth)
  If Len(sMonth) = 1 Then
    sMonth = "0"+sMonth
  End If

  sDay = CStr(parDay)
  If Len(sDay) = 1 Then
    sDay = "0"+sDay
  End If

  'Проверка существования даты
  If not IsDate(sMonth&"/"&sDay&"/"&sYear) Then
    ReportUsersActivity = ""
    Exit Function
  End If

  'Формирование SQL-запросов в зависимости от типа запрашиваемой информации
  Select Case RequestType
    Case "Login"
      sSQL = "select count(distinct [Log].UserID) as Value from [Log] left outer join Users on [Log].UserID = Users.UserID where [Log].DocID = N'SysLogin' and "
      'Поле даты по которому идет поиск
      sDateField = "DateTime"
    Case "Activity"
      sSQL = "select count(distinct [Log].UserID) as Value from [Log] left outer join Users on [Log].UserID = Users.UserID where [Log].DocID <> N'SysLogin' and [Log].DocID <> N'SysLogout' and "
      sDateField = "DateTime"
    Case "DocCreation"
      sSQL = "select count(distinct "+GetLoginSQL_MSSQL("Docs.NameCreation")+") as Value from Docs left outer join Users on "+GetLoginSQL_MSSQL("Docs.NameCreation")+" = Users.UserID where "
      sDateField = "Docs.DateCreation"
    Case "DocTypes"
      sSQL = "select count(DocID) as Value from Docs left outer join Users on "+GetLoginSQL_MSSQL("Docs.NameCreation")+" = Users.UserID where CharIndex(Docs.ClassDoc, N'"+sClassDoc+"') > 0 and "
      sDateField = "Docs.DateCreation"
    Case "BlockTitle" 'Заголовок блока выводимой информации
'      ReportUsersActivity = "<div align=""center"">-----</div>"
      ReportUsersActivity = "<hr>"
	  'rmanyushin 26.08.2009 Не выводить сумму в строке разделитель
	  iReportUsersActivityRowSum = "<hr>"
      Exit Function
    Case "" 'Ничего не выводим (для формирования разделителей)
      ReportUsersActivity = "&nbsp;"
	  'rmanyushin 26.08.2009 Не выводить сумму в строке разделитель
	  iReportUsersActivityRowSum = "&nbsp;"
      Exit Function
    Case Else
      ReportUsersActivity = "Unknown RequestType"
      Exit Function
  End Select

  'Фильтр по дате
  sDateAdd = " DateDiff(day, "+sDateField+", {d '"+sYear+"-"+sMonth+"-"+sDay+"'}) = 0 "

  'Фильтр по домену
  Select Case DomainType
    Case "RU"
	'rmanyushin 26.08.2009 Изменяем критерий фильтрации пользователей Россия/Чехия. Поскольку все в одном домене STS, то фильтруем по полю Users.Company
		'sDomainAdd = " and WinLogin is not NULL and WinLogin like N'root\%' "
		sDomainAdd = " AND (Users.Company ='JSC Sitronics Telecom Solutions'" 
		sDomainADD = sDomainADD & " OR Users.Company ='SITRONICS Telecom Solutions, Ukraine'"
		sDomainADD = sDomainADD & " OR Users.Company ='""ИнтерТел Сибирь""')"
    Case "CZ"
	'rmanyushin 26.08.2009 Изменяем критерий фильтрации пользователей Россия/Чехия. Поскольку все в одном домене STS, то фильтруем по полю Users.Company
		'sDomainAdd = " and WinLogin is not NULL and WinLogin like N'stromtelecom\%' "
		sDomainADD = " AND (Users.Company ='Sitronics TS'"
		sDomainADD = sDomainADD & " OR Users.Company ='Sitronics TS nonemp'"
		sDomainADD = sDomainADD & " OR Users.Company ='AnnexNet')"
    Case "ALL"
	'rmanyushin 26.08.2009 Изменяем критерий фильтрации пользователей Россия/Чехия. Поскольку все в одном домене STS, то фильтруем по полю Users.Company
		'sDomainAdd = " and WinLogin is not NULL and (WinLogin like N'root\%' or WinLogin like N'stromtelecom\%') "
		sDomainAdd = " AND (Users.Company ='JSC Sitronics Telecom Solutions'" 
		sDomainADD = sDomainADD & " OR Users.Company ='SITRONICS Telecom Solutions, Ukraine'"
		sDomainADD = sDomainADD & " OR Users.Company ='""ИнтерТел Сибирь""'"
		sDomainADD = sDomainADD & " OR Users.Company ='Sitronics TS'"
		sDomainADD = sDomainADD & " OR Users.Company ='Sitronics TS nonemp'"
		sDomainADD = sDomainADD & " OR Users.Company ='AnnexNet')"
    Case Else
      ReportUsersActivity = "Unknown DomainType"
      Exit Function
  End Select

  'Окончательный запрос
  sSQL = sSQL+sDateAdd+sDomainAdd
AddLogD "ReportUsersActivity - RequestType = "+MyCStr(RequestType)+"  DomainType = "+MyCStr(DomainType)+"  Day = "+MyCStr(parDay)+"  Month = "+MyCStr(parMonth)+"  Year = "+MyCStr(parYear)+"  sClassDoc = "+MyCStr(sClassDoc)
AddLogD "ReportUsersActivity - sSQL = "+sSQL

  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  dsTemp.Open sSQL, Conn, 3, 1, &H1
  'Все запросы с Count, поэтому на EOF не проверяем
  ReportUsersActivity = MyCStr(dsTemp("Value"))
  iReportUsersActivityRowSum = iReportUsersActivityRowSum + dsTemp("Value")
  
  dsTemp.Close

  'Раскраска результатов
  'Для отчетов по пользователям 0 - красный
  If RequestType <> "DocTypes" Then
    If ReportUsersActivity = "0" Then
      ReportUsersActivity = "<font color=red>"+ReportUsersActivity+"</font>"
    End If
  Else
    If InStr(sClassDoc, ",") > 0 Then 'Отчет по нескольким категориям
      ReportUsersActivity = "<b>"+ReportUsersActivity+"</b>"
      If DomainType  <> "ALL" Then
        ReportUsersActivity = "<font color=""#003366"">"+ReportUsersActivity+"</font>"
      End If
    End If
  End If

  'Суммы по доменам - жирные
  If DomainType = "ALL" Then
    ReportUsersActivity = "<b>"+ReportUsersActivity+"</b>"
  End If
End Function

Function ReportUsersActivityRowSum()
  ReportUsersActivityRowSum = "<b>"+CStr(iReportUsersActivityRowSum)+"</b>"
End Function
' ------------------- Формирование отчета "Анализ активности пользователей за месяц" -------------------- END

'Функция, возвращающая список подчиненных руководителей. Для отчета "Отчет по заданному пользователю"
Function GetAllLowerLeaders(ByVal UserID)
  Dim dsTemp, sDepartment, dsTemp2
  
'  sSQL = "Select distinct Users.Name from DepartmentDependants left join Departments on DepartmentDependants.DependantGUID = Departments.GUID left join Users on Users.UserID = "+GetLoginSQL_MSSQL("Leader")+" where Departments.Name like IsNull((select Department from Users where UserID = N'"+UserID+"'), '')+'%' and Departments.Name <> IsNull((select Department from Users where UserID = N'"+UserID+"'), '') and Users.UserID <> N'"+UserID+"' order by Users.Name"
'AddLogD "GetAllLowerLeaders - SQL: "+sSQL

  sSQL = "Select distinct Departments.Name as Name,Leader from DepartmentDependants left join Departments on DepartmentDependants.DependantGUID = Departments.GUID where Leader like N'%<"+UserID+">%' order by Departments.Name"
AddLogD "GetAllLowerLeaders - Leader of departments SQL: "+sSQL
  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  dsTemp.Open sSQL, Conn, 3, 1, &H1
  GetAllLowerLeaders = ""
  Do While not dsTemp.EOF
    sDepartment = dsTemp("Name")
'запрос возвращает всех подчиненных руководителей по веткам вниз
'    sSQL = "Select distinct Users.Name from DepartmentDependants left join Departments on DepartmentDependants.DependantGUID = Departments.GUID left join Users on Users.UserID = "+GetLoginSQL_MSSQL("Leader")+" where Departments.Name like N'"+sDepartment+"%' and Departments.Name <> N'"+sDepartment+"' and Users.UserID <> N'"+UserID+"' order by Users.Name"
'запрос возвращает подчиненных руководителей только первой ступени (находящихся в непосредственном подчинении)
    sSQL = "Select distinct Users.Name from DepartmentDependants left join Departments on DepartmentDependants.DependantGUID = Departments.GUID left join Users on Users.UserID = "+GetLoginSQL_MSSQL("Leader")+" where Departments.Name like N'"+sDepartment+"%' and Departments.Name <> N'"+sDepartment+"' and Users.UserID <> N'"+UserID+"' and CharIndex(N'/', Departments.Name, Len(N'"+sDepartment+"')+1) in (0,Len(Departments.Name)) order by Users.Name"
AddLogD "GetAllLowerLeaders - Department: "+MyCStr(sDepartment)
AddLogD "GetAllLowerLeaders - Get lower leaders SQL: "+sSQL

    Set dsTemp2 = Server.CreateObject("ADODB.Recordset")
    dsTemp2.Open sSQL, Conn, 3, 1, &H1
    If GetAllLowerLeaders <> "" Then
      GetAllLowerLeaders = GetAllLowerLeaders+VbCrLf
    End If
    GetAllLowerLeaders = GetAllLowerLeaders+"<b>"+DelOtherLangFromFolder(sDepartment)+"</b>"

    If dsTemp2.EOF Then
      GetAllLowerLeaders = GetAllLowerLeaders+VbCrLf+"<font color = red>-----</font>"
    Else
      Do While not dsTemp2.EOF
        GetAllLowerLeaders = GetAllLowerLeaders+VbCrLf+"<font color=blue>"+DelOtherLangFromFolder(MyCStr(dsTemp2("Name")))+"</font>"
        dsTemp2.MoveNext
      Loop
    End If
    dsTemp2.Close
    dsTemp.MoveNext
  Loop
  dsTemp.Close
  GetAllLowerLeaders = Replace(GetAllLowerLeaders, VbCrLf, "<br>")
End Function

'Получить роли которые исполняет пользователь
Function GetUserRole(ByVal parUserID, ByVal sDepartment)
  Dim sUserID, dsTemp, i

  GetUserRole = ""
  If parUserID = "" Then
    Exit Function
  End If
  sUserID = "<"+parUserID+">"
'  If InStr(UCase(sDepartment), UCase(SIT_SITRONICS)) = 1 Then
'    sDir = SIT_RolesDirSitronics
'  ElseIf InStr(UCase(sDepartment), UCase(SIT_STS)) = 1 Then
'    sDir = SIT_RolesDirSTS
'  Else
'    sDir = SIT_RolesDirSitronics
'  End If
'
'  sSQL = "Select UserDirValues.Field1 from UserDirValues inner join UserDirectories on (UserDirValues.UDKeyField = UserDirectories.KeyField) where UserDirectories.Name = "+sUnicodeSymbol+"'"+sDir+"' and UserDirValues.Field2 like "+sUnicodeSymbol+"'%"+sUserID+"%' order by UserDirValues.Field1"
'AddLogD "GetUserRole - SQL: "+sSQL
'
'  Set dsTemp = Server.CreateObject("ADODB.Recordset")
'  dsTemp.Open sSQL, Conn, 3, 1, &H1
'  Do While not dsTemp.EOF
'    If GetUserRole <> "" Then
'      GetUserRole = GetUserRole+VbCrLf
'    End If
'    GetUserRole = GetUserRole+MyCStr(dsTemp("Field1"))
'    dsTemp.MoveNext
'  Loop
'  dsTemp.Close

  For i = 1 to 4
    If i <> 4 Then
      If InStr(UCase(sDepartment), UCase(SIT_SITRONICS)) = 1 Then
        Select Case i
          Case 1
            sDir = SIT_RolesDirSitronics_RU
          Case 2
            sDir = SIT_RolesDirSitronics_EN
          Case 3
            sDir = SIT_RolesDirSitronics_CZ
        End Select
      ElseIf InStr(UCase(sDepartment), UCase(SIT_STS)) = 1 Then
        Select Case i
          Case 1
            sDir = SIT_RolesDirSTS_RU
          Case 2
            sDir = SIT_RolesDirSTS_EN
          Case 3
            sDir = SIT_RolesDirSTS_CZ
        End Select
      Else
        Select Case i
          Case 1
            sDir = SIT_RolesDirSitronics_RU
          Case 2
            sDir = SIT_RolesDirSitronics_EN
          Case 3
            sDir = SIT_RolesDirSitronics_CZ
        End Select
      End If
      sSQL = "Select UserDirValues.Field1 as Role from UserDirValues inner join UserDirectories on (UserDirValues.UDKeyField = UserDirectories.KeyField) where UserDirectories.Name = "+sUnicodeSymbol+"'"+sDir+"' and UserDirValues.Field2 like "+sUnicodeSymbol+"'%"+sUserID+"%' order by UserDirValues.Field1"
AddLogD "GetUserRole - SQL"+CStr(i)+": "+sSQL
    Else
      sDir = "Roles for Orders"
      sSQL = "Select Role, BusinessUnit from RolesForOrders_STS where Users like "+sUnicodeSymbol+"'%"+sUserID+"%' order by Role,BusinessUnit"
AddLogD "GetUserRole - SQL"+CStr(i)+": "+sSQL
    End If

    Set dsTemp = Server.CreateObject("ADODB.Recordset")
    dsTemp.Open sSQL, Conn, 3, 1, &H1
    If not dsTemp.EOF Then
      If GetUserRole <> "" Then
        GetUserRole = GetUserRole+VbCrLf
      End If
      GetUserRole = GetUserRole+"<b>"+sDir+"</b>"
    End If
    Do While not dsTemp.EOF
      GetUserRole = GetUserRole+VbCrLf+MyCStr(dsTemp("Role"))
      If i = 4 Then
        GetUserRole = GetUserRole+" <i>(BU: "+GetCodeFromCode_NameString(Trim(dsTemp("BusinessUnit")))+")</i>"
      End If
      dsTemp.MoveNext
    Loop
    dsTemp.Close
  Next
  GetUserRole = Replace(GetUserRole, VbCrLf, "<br>")
End Function

'Получить непосредственного начальника (если бизнес единиц у пользователя несколько, то берется первая)
Function GetNearestChiefForReport(ByVal parDepartment, ByVal parUserID, ByVal parBusinessUnits)
  Dim sChief, arBUs, i, sChiefLogin, dsTemp
  
'  sChief = GetNearestChief(parDepartment, parUserID, parBusinessUnits)
'  GetNearestChiefForReport = GetUserID(sChief)
'  If GetNearestChiefForReport = "" Then
'    GetNearestChiefForReport = sChief
'  Else
'    oPayDox.GetUserDetails GetNearestChiefForReport, sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
'    GetNearestChiefForReport = DelOtherLangFromFolder(sName)
'  End If
  GetNearestChiefForReport = ""
  If parDepartment <> "" Then
    GetNearestChiefForReport = "<b>"+DelOtherLangFromFolder(parDepartment)+"</b>"
    If parBusinessUnits = "" Then
      sChief = GetNearestChief(parDepartment, parUserID, "")
      sChiefLogin = GetUserID(sChief)
      If sChiefLogin = "" Then
        If sChief = "" Then
          GetNearestChiefForReport = GetNearestChiefForReport+VbCrLf+"<font color = red>-----</font>"
        Else
          GetNearestChiefForReport = GetNearestChiefForReport+VbCrLf+"<font color=blue>"+sChief+"</font>"
        End If
      Else
        oPayDox.GetUserDetails sChiefLogin, sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
        GetNearestChiefForReport = GetNearestChiefForReport+VbCrLf+"<font color=blue>"+DelOtherLangFromFolder(sName)+"</font>"
      End If
    Else
      arBUs = Split(parBusinessUnits, VbCrLf)
      For i = 0 To UBound(arBUs)
        If Trim(arBUs(i) <> "") Then
          GetNearestChiefForReport = GetNearestChiefForReport+VbCrLf+"<i>BU: "+GetCodeFromCode_NameString(Trim(arBUs(i)))+"</i>"
          sChief = GetNearestChief(parDepartment, parUserID, arBUs(i))

          sChiefLogin = GetUserID(sChief)
          If sChiefLogin = "" Then
            If sChief = "" Then
              GetNearestChiefForReport = GetNearestChiefForReport+VbCrLf+"<font color = red>-----</font>"
            Else
              GetNearestChiefForReport = GetNearestChiefForReport+VbCrLf+"<font color=blue>"+sChief+"</font>"
            End If
          Else
            oPayDox.GetUserDetails sChiefLogin, sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
            GetNearestChiefForReport = GetNearestChiefForReport+VbCrLf+"<font color=blue>"+DelOtherLangFromFolder(sName)+"</font>"
          End If
        End If
      Next
    End If
  End If
  'Кусок по сбору руководителей по подчиненным (руководимым) подразделениям------------------------------------
  sSQL = "Select Departments.Name as Department,Leader,BusinessUnit from DepartmentDependants left join Departments on DepartmentDependants.DependantGUID = Departments.GUID where Leader like N'%<"+parUserID+">%' order by Departments.Name,BusinessUnit"
AddLogD "GetNearestChiefForReport - Leader of departments SQL: "+sSQL
  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  dsTemp.Open sSQL, Conn, 3, 1, &H1
  Do While not dsTemp.EOF
    If GetNearestChiefForReport <> "" Then
      GetNearestChiefForReport = GetNearestChiefForReport+VbCrLf
    End If
    GetNearestChiefForReport = GetNearestChiefForReport+"<b>"+DelOtherLangFromFolder(MyCStr(dsTemp("Department")))+"</b> <i>(BU: "+GetCodeFromCode_NameString(MyCStr(dsTemp("BusinessUnit")))+")</i>"
    sChief = GetNearestChief(MyCStr(dsTemp("Department")), parUserID, MyCStr(dsTemp("BusinessUnit")))

    sChiefLogin = GetUserID(sChief)
    If sChiefLogin = "" Then
      If sChief = "" Then
        GetNearestChiefForReport = GetNearestChiefForReport+VbCrLf+"<font color = red>-----</font>"
      Else
        GetNearestChiefForReport = GetNearestChiefForReport+VbCrLf+"<font color=blue>"+sChief+"</font>"
      End If
    Else
      oPayDox.GetUserDetails sChiefLogin, sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
      GetNearestChiefForReport = GetNearestChiefForReport+VbCrLf+"<font color=blue>"+DelOtherLangFromFolder(sName)+"</font>"
    End If
    dsTemp.MoveNext
  Loop
  dsTemp.Close
  '------------------------------------------------------------------------------------------------------------
  GetNearestChiefForReport = Replace(GetNearestChiefForReport, VbCrLf, "<br>")
End Function

'Получить кто от имени фин. контролера согласовал/утвердил документ
'(берется последний по дате комментарий, считается, что фин. контролер либо согласует,
' либо утверждает)
Function GetFincontrolInfo(ByVal sDocID)
  Dim dsTemp,sUserID
  sSQL = "select Comments.Comment,Comments.DateCreation,Comments.CommentType,Comments.SpecialInfo,Docs.NameApproved from Docs left outer join Comments on Docs.DocID = Comments.DocID left outer join RolesForOrders_STS on CharIndex(Comments.UserID, RolesForOrders_STS.Users) > 0 where Docs.DocID = N'"+sDocID+"' and Role = N'""#Financial controller"";' and (Comments.CommentType = N'APROVAL' or (Comments.CommentType = N'VISA' and (Comments.SpecialInfo = N'VISAOK' or Comments.SpecialInfo = N'VISACANCEL'))) order by Comments.DateCreation desc"
AddLogD "GetFincontrolInfo SQL: "+sSQL
  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  dsTemp.Open sSQL, Conn, 3, 1, &H1
  GetFincontrolInfo = ""
  If dsTemp.EOF Then
    Exit Function
  End If
  If dsTemp("CommentType") = "APROVAL" Then
    If InStr("-<", dsTemp("NameApproved")) = 0 Then 'Утвержден
      GetFincontrolInfo = DOCS_Approved
    Else 'Отказ утверждения
      GetFincontrolInfo = DOCS_RefusedApp
    End If
  ElseIf dsTemp("CommentType") = "VISA" Then
    If dsTemp("SpecialInfo") = "VISAOK" Then 'Согласован
      GetFincontrolInfo = DOCS_Reconciled
    ElseIf dsTemp("SpecialInfo") = "VISACANCEL" Then 'Отказ согласования
      GetFincontrolInfo = DOCS_Refused
    End If
  End If
  
  If InStr(dsTemp("Comment"), "/") = 0 Then
    sUserID = GetUserID(dsTemp("Comment"))
  Else
    sUserID = GetUserID(Mid(dsTemp("Comment"), InStrRev(dsTemp("Comment"), "/")))
  End If
  oPayDox.GetUserDetails sUserID, sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
  GetFincontrolInfo = GetFincontrolInfo+" - "+GetFullName(SurnameGN(DelOtherLangFromFolder(sName)), sUserID)+" ("+MyDate(dsTemp("DateCreation"))+")"
  dsTemp.Close
End Function

Function LinkToUser(ByVal sUserID)
  If Trim(Request("R1"))="MSWord" or Trim(Request("R1"))="MSExcel" Then
    LinkToUser = sUserID
  Else
    LinkToUser = "<a href="""+GetURLEncode("ShowUser.asp", "?UserID=", sUserID)+"&DocPrintableView=&bMessages="" target=""_blank"">"+sUserID+"</a>"
  End If
End Function

'###########################################################################################################################################
'###########################################################################################################################################
'###########################################################################################################################################
'###########################################################################################################################################
'###########################################################################################################################################
'###########################################################################################################################################

' Процедура генерирования индекса документа
'ph - 20100603 - start - Более нигде не используется, закомментирована для наглядности
'Sub GetNewDocIDForClassDocWithPrefixNew(sClassDoc, sSearchCol, sPrePrefix, sPrefix, sSufix, sPostfix, sDepartment)
'
'  sNumberLen = "3"
'
'  'для генерации регистрационного номера
'  If sPrePrefix="" Then
'    Set oPayDox1 = Server.CreateObject("PayDox.Common")
'    Set Conn = oPayDox1.Conn
'  End If
'
'  Set dsTempPR = Server.CreateObject("ADODB.Recordset")
'
'  If InStr(sPrefix,"%")>0 Then
'    sPrefixSearch = sPrefix
'    sPrefix = Left(sPrefix, InStr(sPrefix,"%")-1)
'  Else
'    sPrefixSearch = sPrefix+"%"
'  End If
'
'  If InStr(UCase(sClassDoc),UCase(SIT_DOGOVORI))>0 Then
'    sSQL="select ISNULL(MAX(cast(left(REPLACE("+sSearchCol+",N'"+sPrePrefix+sPrefix+"',''),len(REPLACE("+sSearchCol+",N'"+sPrePrefix+sPrefix+"',''))-6)as int)),0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sPrePrefix+sPrefixSearch+"' AND Department like N'"+sDepartment+"%'"
'  ElseIf InStr(UCase(sClassDoc),UCase(SIT_NORM_DOCS))>0 Then
'    'sSQL="select ISNULL(MAX(cast(Right(REPLACE("+sSearchCol+",N'.'+UserFieldText3,''), "+sNumberLen+") as int)), 0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sPrePrefix+sPrefix+"%' AND Department like N'"+sDepartment+"%'"
''SAY 2008-12-26
'sSQL="select ISNULL(MAX(cast(substring(REPLACE("+sSearchCol+",N'"+sPrePrefix+"',''),charindex('-',REPLACE("+sSearchCol+",N'"+sPrePrefix+"',''),0)+1, charindex('.',REPLACE("+sSearchCol+",N'"+sPrePrefix+"',''),0)-charindex('-',REPLACE("+sSearchCol+",N'"+sPrePrefix+"',''),0)-1) as int)), 0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sPrePrefix+sPrefix+"%' AND Department like N'"+sDepartment+"%'"
'
'  ElseIf InStr(UCase(sClassDoc),UCase(SIT_ZADACHI))>0 Then
'    If InStr(sPrefix,"AFK") Then
'      sSQL="select ISNULL(MAX(cast(Right("+sSearchCol+", "+sNumberLen+") as int)), 0) as MaxDocID from Docs where ClassDoc like N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sPrePrefix+sPrefix+"%' AND Department like N'"+sDepartment+"%'"    
'    Else
'      'sSQL="select ISNULL(MAX(cast(Right("+sSearchCol+", "+sNumberLen+") as int)), 0) as MaxDocID from Docs where ClassDoc like N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sPrePrefix+sPrefix+"%' AND "+sSearchCol+" not like N'%AFK%' AND Department like N'"+sDepartment+"%'"    
'      sSQL="select ISNULL(Right(DocID, 3), '0') as MaxDocID from Docs where DocID like N'T_%' AND DocID not like N'%AFK%' and ClassDoc like N'"+sClassDoc+"' order by DateCreation DESC, MaxDocID DESC"
'    End If
'  Else
'   sSQL="select ISNULL(MAX(cast(Right("+sSearchCol+", "+sNumberLen+") as int)), 0) as MaxDocID from Docs where ClassDoc like N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sPrePrefix+sPrefix+"%' AND Department like N'"+sDepartment+"%'"
'  End If
'
''!!!!!!!!!!!!!!!!
''Session("Message")=Session("Message")+" sSQL="+sSQL
''AddlogD "W@Y sSQL="+sSQL
'  dsTempPR.Open sSQL, Conn, 2, 1, &H1
'  If dsTempPR.EOF Then
'    sNumber = 0
'  Else
'    sNumber=dsTempPR("MaxDocID")
'  End If
'  dsTempPR.Close
'
'  If sNumber=0 Then
'    sNumberNext=1
'  Else
'    sNumberNext=sNumber+1
'  End If
'
'  S_DocID = sPrePrefix + sPrefix + sSufix + LeadSymbolNVal(sNumberNext, "0", sNumberLen)+sPostfix
'
'  'phil - 20080909 - start - TESTTESTTESTTESTTESTTESTTESTTESTTESTTESTTESTTESTTEST
'  'AddlogD "@@@@sSQL: " + sSQL
'  'AddlogD "@@@@S_DocID: " + S_DocID
'  'phil - 20080909 - start - TESTTESTTESTTESTTESTTESTTESTTESTTESTTESTTESTTESTTEST
'End Sub
'ph - 20100603 - end

'Статусы документа для отчета
Function MyGetStatuses(docIsActive,docListToReconcile,docListReconciled,docStatusDevelopment,docStatusPayment,docStatusCompletion,docLocationPath,docDateCompletion,docNameApproved, docClassDoc)
sMessage=""
sMessage1=""
sMessage2=""
sColor=""
sColor1=""
sColor2=""
sColor3=""
'If Trim(MyCStr(docNameAproval))="" Or IsNull(docStatusDevelopment) Or ShowStatusDevelopment(docStatusDevelopment)=DOCS_Approved Or ShowStatusDevelopment(docStatusDevelopment)=DOCS_RefusedApp Then
'	sIsAprovalRequired=False
'Else
	sIsAprovalRequired=True
'End If

If docIsActive=VAR_InActiveTask Then

	If docClassDoc=DOCS_Notices Then
		sMessage=DOCS_TaskInactive
    Else
		sMessage=DOCS_Inactive
    End If
	sColor="#808080"

Else

sColor="green"
If docListToReconcile<>"" And IsReconciliationRequired(docListToReconcile, docListReconciled, docStatusCompletion) Then
	If oPayDox.IsReconciliationWaiting(docListToReconcile) Then
		sMessage=DOCS_ReconciliationWaiting1
	ElseIf (oPayDox.IsReconciliationPendingWithOptions(docListToReconcile, docListReconciled, Var_ReconciliationIfAllAgree) And Not IsVisaNowWithOptions(Session("UserID"), docListToReconcile, docListReconciled, docNameApproved)) Or oPayDox.IsReconciliationWaitingAfterRefuse(docListToReconcile, docListReconciled, Var_ReconciliationIfAllAgree) Then
		sMessage=DOCS_ReconciliationPending
	Else
		sMessage=DOCS_ReconciliationREQUIRED
	End If
	sColor="red"
ElseIf IsReconciliationCompleteWithOptions(docListToReconcile, docListReconciled) And sIsAprovalRequired Then
	sMessage=DOCS_APROVALREQUIRED
	If oPayDox.IsReconciliationInterrupted(docListToReconcile) Then
		sMessage=DOCS_ReconciliationComplete1+". "+DOCS_APROVALREQUIRED
	End If
	sColor="red"
End If

If ShowStatusDevelopment(docStatusDevelopment)=DOCS_RefusedApp Then
	sMessage=DOCS_RefusedApp
	sColor="red"
ElseIf ShowStatusDevelopment(docStatusDevelopment)=DOCS_Approved Then
	sMessage=DOCS_Approved
	sColor="green"
End If

sColor1="green"
If bStatusPayment="Y" Then
	sMessage1=ShowStatusPayment(docStatusPayment)
	If sMessage1<>DOCS_StatusPaymentPaid Then
		sColor1="red"
	End If
End If

If InStr(MyCStr(docListToReconcile), "#@") > 0 Then
	If sMessageParent<>"" Then
		sMessage=sMessageParent
		sColor=sColorMessageParent
	End If
End If

sColor2="green"
If Not IsNull(docStatusCompletion) Or ShowStatusCompletion(docStatusCompletion)=DOCS_Cancelled Or ShowStatusCompletion(docStatusCompletion)=DOCS_Completed Then
	sMessage2=ShowStatusCompletion(docStatusCompletion)
	If sMessage<>DOCS_Approved And sMessage<>"" And sMessage2<>DOCS_Cancelled Then
		sMessage2=""
	End If
	If sMessage2=DOCS_Cancelled Then
		sMessage=""
		sMessage1=""
	End If
	If sMessage2<>DOCS_Completed and sMessage2<>DOCS_Cancelled Then
		If docDateCompletion>VAR_BeginOfTimes and Not IsNull(docDateCompletion) Then
			If DateFullTime(docDateCompletion)<Now Then
				sColor2="red"
				sMessage2=DOCS_EXPIRED2
			End If
		End If
	End If
End If
sMessage3=""
sLocationPath=Trim(MyCStr(docLocationPath))
If Not IsNull(docLocationPath) And (ShowStatusDevelopment(docStatusDevelopment)=DOCS_Approved Or IsNull(docNameApproved)) And InStr(sLocationPath, ">+")<=0 Then
	sMessage3=DOCS_REGISTRATIONREQUIRED
	sColor3="red"
End If

If bReviewRequired Then
	sMessage=DOCS_ReviewRequired
	sColor="red"
End If

End If

sFormatStr1 = "<font color="
sFormatStr2 = ">"
sFormatStr3 = "</font>"
MyGetStatuses = "<b>"
If sMessage<>"" Then
  MyGetStatuses = MyGetStatuses+sFormatStr1&sColor&sFormatStr2&sMessage&sFormatStr3
End If
If sMessage1<>"" Then
  If MyGetStatuses <> "" Then
    MyGetStatuses = MyGetStatuses+"<br>"
  End If
  MyGetStatuses = MyGetStatuses+sFormatStr1&sColor1&sFormatStr2&sMessage1&sFormatStr3
End If
If sMessage2<>"" Then
  If MyGetStatuses <> "" Then
    MyGetStatuses = MyGetStatuses+"<br>"
  End If
  MyGetStatuses = MyGetStatuses+sFormatStr1&sColor2&sFormatStr2&sMessage2&sFormatStr3
End If
If sMessage3<>"" Then
  If MyGetStatuses <> "" Then
    MyGetStatuses = MyGetStatuses+"<br>"
  End If
  MyGetStatuses = MyGetStatuses+sFormatStr1&sColor3&sFormatStr2&sMessage3&sFormatStr3
End If
If sMessage="" And sMessage1="" And sMessage2="" And sMessage3="" Then
  MyGetStatuses = MyGetStatuses+sFormatStr1&sColor2&sFormatStr2&DOCS_Active&sFormatStr3
End If
MyGetStatuses = MyGetStatuses+"</b>"
End Function

'Запрос логина пользователя в отчет
Sub ReportRequestFormFieldInputUserID(sTitle, sValue, sSQLContext)
nUserPars=nUserPars+1
S_DirGUID="U"
sUserID = GetUserID(sValue)
If sUserID = "" Then
  sUserID = sValue
End If
%>
<tr>
    <td width="35%" valign="top" align="left"><font <%=StyleDetailName%>><%=sTitle%></font>&nbsp;</td>
    <td width="65%">
        <input type="text" name="UserParName<%=Trim(CStr(nUserPars))%>" size="<%=VAR_TextFieldSize-14%>" value="<%=HTMLEncode(sValue)%>" onpropertychange="javascript:forma.UserPar<%=Trim(CStr(nUserPars))%>.value=jsGetLogin(forma.UserParName<%=Trim(CStr(nUserPars))%>.value);">
        <%DirectoryCall "UserParName"+Trim(CStr(nUserPars)), S_DirGUID, "", ""%>
        <input type="hidden" name="SQLContext<%=Trim(CStr(nUserPars))%>" value="<%=MyCStr(sSQLContext)%>">
        <input type="hidden" name="UserParTitle<%=Trim(CStr(nUserPars))%>" value="<%=HTMLEncode(sTitle)%>">
        <input type="hidden" name="UserPar<%=Trim(CStr(nUserPars))%>" value="<%=HTMLEncode(sUserID)%>">
    </td>
</tr>
<%
%>
<script language="JavaScript"><!--
    function jsGetLogin(sStr) 
{
        poslt = sStr.indexOf('<');
        posgt = sStr.indexOf('>');
        if (poslt == -1 || posgt == -1 || posgt - poslt < 0) 
{
            return sStr;
        }
        return sStr.substring(poslt + 1, posgt);
    }
// --></script>
<%
End Sub 

'ph - 20090524 - start
'Запрос по кнопке ВСЕ ДОКУМЕНТЫ
If UCase(Request.ServerVariables("URL"))=UCase("/ListDoc.asp") and Request("AllDocs")="y" Then
  If IsAdmin() Then
    If Application("sVersion") = "MSSQL" Then
      VAR_ListDocSQL = "select * from Docs where DateCompleted is null or DateCompleted > GetDate()-90"
    ElseIf Application("sVersion") = "MSACCESS" Then
      VAR_ListDocSQL = "select * from Docs where DateCompleted is null or DateCompleted > Date()-90"
    End If
  Else
    sCheckDepartmentInListSQL = opaydox.CheckDepartmentInListSQL(Session("Department"), "ListToView")
    bAdditionalUsersExists = IsAdditionalUsersFieldExist()
    If Application("sVersion") = "MSSQL" Then
      VAR_ListDocSQL = "select * from Docs where (CharIndex(N'<"+Session("UserID")+">', NameCreation)<>0 or (IsActive<>'N' and (SecurityLevel < 4 or CharIndex(N'<"+Session("UserID")+">', ListToReconcile)<>0 or CharIndex(N'<"+Session("UserID")+">', NameAproval)<>0 or CharIndex(N'<"+Session("UserID")+">', NameResponsible)<>0 or CharIndex(N'<"+Session("UserID")+">', NameControl)<>0 or CharIndex(N'<"+Session("UserID")+">', LocationPath)<>0 or CharIndex(N'<"+Session("UserID")+">', Author)<>0 or CharIndex(N'<"+Session("UserID")+">', Correspondent)<>0 or CharIndex(N'<"+Session("UserID")+">', ListToEdit)<>0 or CharIndex(N'<"+Session("UserID")+">', ListToView)<>0 or CharIndex('<USERS:*>', ListToView)<>0 "
      VAR_ListDocSQL = VAR_ListDocSQL+sCheckDepartmentInListSQL
      If bAdditionalUsersExists Then
        VAR_ListDocSQL = VAR_ListDocSQL+"  or CharIndex(N'<"+Session("UserID")+">', AdditionalUsers)<>0"
      End If
      VAR_ListDocSQL = VAR_ListDocSQL+"))) and (DateCompleted is null or DateCompleted > GetDate()-90)"
    ElseIf Application("sVersion") = "MSACCESS" Then
      VAR_ListDocSQL = "select * from Docs where (InStr(NameCreation, '<"+Session("UserID")+">')<>0 or (IsActive<>'N' and (SecurityLevel < 4 or InStr(ListToReconcile, '<"+Session("UserID")+">')<>0 or InStr(NameAproval, '<"+Session("UserID")+">')<>0 or InStr(NameResponsible, '<"+Session("UserID")+">')<>0 or InStr(NameControl, '<"+Session("UserID")+">')<>0 or InStr(LocationPath, '<"+Session("UserID")+">')<>0 or InStr(Author, '<"+Session("UserID")+">')<>0 or InStr(Correspondent, '<"+Session("UserID")+">')<>0 or InStr(ListToEdit, '<"+Session("UserID")+">')<>0 or InStr(ListToView, '<"+Session("UserID")+">')<>0 or InStr(ListToView, '<USERS:*>')<>0 "
      VAR_ListDocSQL = VAR_ListDocSQL+sCheckDepartmentInListSQL
      If bAdditionalUsersExists Then
        VAR_ListDocSQL = VAR_ListDocSQL+" or InStr(AdditionalUsers, '<"+Session("UserID")+">')<>0"
      End If
      VAR_ListDocSQL = VAR_ListDocSQL+"))) and (DateCompleted is null or DateCompleted > Date()-90)"
    End If
  End If
  VAR_ListDocSQL = VAR_ListDocSQL+" order by DateActivation Desc, DateCreation Desc"
End If
'ph - 20090524 - end

'20090810 - Запрос №4 из СТС - Отчеты
SIT_ReportDateActivation = VAR_BeginOfTimes

'Определить время первой активации
Function GetFirstActivationDate(parDocID, parDateActivation)
  Dim dsTemp

  sSQL = "select top 1 DateCreation from Comments where DocID = N'"+parDocID+"' and CommentType = 'system' and (SpecialInfo = 'DOCS_Active' or CharIndex(N'Dokument je aktivní', Comment) = 1 or CharIndex(N'Active', Comment) = 1 or CharIndex(N'Активен', Comment) = 1 or CharIndex(N'Document active', Comment) = 1 or CharIndex(N'Документ активен', Comment) = 1) order by DateCreation"
AddLogD "GetFirstActivationDate SQL: "+sSQL
  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  dsTemp.Open sSQL, Conn, 3, 1, &H1
  If dsTemp.EOF Then
    SIT_ReportDateActivation = parDateActivation
  Else
    SIT_ReportDateActivation = dsTemp("DateCreation")
  End If
  GetFirstActivationDate = SIT_ReportDateActivation
  dsTemp.Close
End Function

'Вырезать английское название из трехъязычного
Function GetEngNameFromFolder(ByVal parStr)
  VAR_CurrentL = ""
  oPayDox.VAR_CurrentL=VAR_CurrentL
  GetEngNameFromFolder = DelOtherLangFromFolder(parStr)
  VAR_CurrentL = "-"
  oPayDox.VAR_CurrentL=VAR_CurrentL
End Function

'Показ статуса исполнения для отчетов
Function MyShowStatusCompletion(parStatusCompletion)
  Select Case parStatusCompletion
    Case VAR_StatusCompletion
      MyShowStatusCompletion = "<font color=green><b>Completed</b></font>"
    Case VAR_StatusCancelled
      MyShowStatusCompletion = "<font color=black><b>Cancelled</b></font>"
    Case VAR_StatusRequestCompletion
      MyShowStatusCompletion = "<font color=FFB000><b>«Completed» requested</b></font>"

    'rmanyushin 83560 09.03.2010 Start
    Case "13" 
		MyShowStatusCompletion = "<font color=red><b>Approval refused</b></font>" 
    'rmanyushin 83560 09.03.2010 End
      
    Case Else
      MyShowStatusCompletion = "<font color=red><b>Actual</b></font>"
  End Select
End Function

'Показ списка ознакомления для отчетов
Function MyShowListToView(parListToView)
  VAR_CurrentL = ""
  oPayDox.VAR_CurrentL=VAR_CurrentL
  MyShowListToView=oPayDox.DelOtherLangDelimiters(parListToView, "<DEPARTMENTS: ", ">", VAR_TreeFolderSeparator)
  MyShowListToView=OutDoc(DelOtherLangFromNames(MyShowListToView))
  VAR_CurrentL = "-"
  oPayDox.VAR_CurrentL=VAR_CurrentL

  MyShowListToView=Trim(MyShowListToView)
  MyShowListToView=Replace(MyShowListToView,"<","&lt;")
  MyShowListToView=Replace(MyShowListToView,">","&gt;")
  MyShowListToView=Replace(MyShowListToView, VbCrLf,"<br>")

  If (Trim(Request("R1"))="MSWord" or Trim(Request("R1"))="MSExcel") Then
    sOK = VAR_OK_GIF
  Else
    sOK = "<img border=""0"" src=""IMAGES/OK.GIF"" alt=""" + DOCS_Viewed + """ width=""13"" height=""12"">"
  End If
  MyShowListToView=Replace(MyShowListToView,"&gt;-", "&gt;"+sOK)
  MyShowListToView=Replace(MyShowListToView,">-", ">"+sOK)

  MyShowListToView=Replace(MyShowListToView, ">-","")
End Function

'Заявка СТС №7 (10.11.2009) - start
'Из MyGetUserDepartment вычленена функция GetLastDepartmentLevelEng

'Получить конечное звено в иерархии подразделения на англ. языке
Function GetLastDepartmentLevelEng(parDepartment)
  If parDepartment = "" Then
    GetLastDepartmentLevelEng = ""
    Exit Function
  End If
  GetLastDepartmentLevelEng = GetEngNameFromFolder(Trim(parDepartment))
  If GetLastDepartmentLevelEng = "" Then
    Exit Function
  End If
  If Right(GetLastDepartmentLevelEng, 1) = "/" Then
    GetLastDepartmentLevelEng = Left(GetLastDepartmentLevelEng, Len(GetLastDepartmentLevelEng)-1)
  End If
  iPos = InStrRev(GetLastDepartmentLevelEng,"/")
  If iPos <> 0 Then
    GetLastDepartmentLevelEng = Mid(GetLastDepartmentLevelEng, iPos+1, Len(GetLastDepartmentLevelEng)-iPos)
  End If
End Function

'Получить подразделение пользователя (конечное звено в иерархии)
Function MyGetUserDepartment(parUser)
  MyGetUserDepartment = ""
  If Trim(parUser) = "" Then
    Exit Function
  End If
  oPayDox.GetUserDetails GetUserID(parUser), sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
  MyGetUserDepartment = GetLastDepartmentLevelEng(sDepartment)
End Function
'Заявка СТС №7 (10.11.2009) - end

'Форматирование даты в формат DD.MM.YYYY
Function DateInDDMMYYYY(ByVal parDate)
  If IsDate(parDate) Then
    DateInDDMMYYYY = LeadSymbolNVal(CStr(Day(parDate)),"0",2)+"."+LeadSymbolNVal(CStr(Month(parDate)),"0",2)+"."+CStr(Year(parDate))
  Else
    DateInDDMMYYYY = ""
  End If
End Function

'Получить дату последнего запроса статуса Исполнено
Function GetLastRequestCompleted(ByVal parDocID)
  Dim dsTemp
  sSQL = "select DateCreation from Comments where DocID = N'"+parDocID+"' and CommentType = 'HISTORY' and (SpecialInfo = 'DOCS_RequestedCompleted' or CharIndex(N'Byl poslán požadavek na nastavení statusu «Dokončeno»', Comment) = 1 or CharIndex(N'Status «Completed» requested', Comment) = 1 or CharIndex(N'Запрошено назначение статуса «Исполнено»', Comment) = 1) order by DateCreation Desc"
AddLogD "GetLastRequestCompleted SQL: "+sSQL
  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  dsTemp.Open sSQL, Conn, 3, 1, &H1
  If dsTemp.EOF Then
    GetLastRequestCompleted = ""
  Else
    GetLastRequestCompleted = DateInDDMMYYYY(dsTemp("DateCreation"))
  End If
  dsTemp.Close
End Function

'Определить число отказов в приеме поручения
Function GetRefuseCompletion(ByVal parDocID)
  Dim dsTemp
  sSQL = "select Count(*) as RefuseCount from Comments where DocID = N'"+parDocID+"' and CommentType = 'HISTORY' and (CharIndex(N'K dopracování', Comment) = 1 or CharIndex(N'Refuse completion', Comment) = 1 or CharIndex(N'На доработку', Comment) = 1)"
AddLogD "GetRefuseCompletion SQL: "+sSQL
  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  dsTemp.Open sSQL, Conn, 3, 1, &H1
  GetRefuseCompletion = dsTemp("RefuseCount")
  dsTemp.Close
End Function

'Вырезать дату без времени
Function GetDateFromDateTime(parDate)
  If IsDate(parDate) Then
    GetDateFromDateTime = DateSerial(Year(parDate), Month(parDate), Day(parDate))
  End If
End Function

'Определить не просрочено ли (было) исполнение документа
Function GetCompletionOnTime(parStatusCompletion, parDateCompletion, parDateCompleted)
  If IsNULL(parStatusCompletion) or IsNULL(parDateCompletion) or parDateCompletion = VAR_BeginOfTimes Then
    GetCompletionOnTime = ""
    Exit Function
  End If
  If parStatusCompletion = VAR_StatusCancelled Then
    GetCompletionOnTime = "<font color=red><b>X</b></font>"
  ElseIf parStatusCompletion = VAR_StatusCompletion Then
    If parDateCompleted <= parDateCompletion Then
      GetCompletionOnTime = "Yes"
    Else
      GetCompletionOnTime = "No"
    End If
  Else
    If parDateCompletion <= Date Then
      GetCompletionOnTime = "No"
    Else
      GetCompletionOnTime = ""
    End If
  End If
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

'Показать разницу дат в рабочих днях (для отчетов)
Function ShowWorkingDaysBetweenTwoDates(DateFrom, DateTill)
  If IsNull(DateTill) or IsNull(DateFrom) or DateTill = VAR_BeginOfTimes or DateFrom = VAR_BeginOfTimes Then
    ShowWorkingDaysBetweenTwoDates = ""
  Else
    ShowWorkingDaysBetweenTwoDates = CStr(WorkingDaysBetweenTwoDates(DateFrom, DateTill))
  End If
End Function

'Проверить Отказ в утверждении
Function IsRefused(parNameApproved)
  IsRefused = not(IsNull(parNameApproved)) and InStr(parNameApproved, "-<") > 0
End Function

'Определить число отказов в ходе согласования/утверждения
Function GetRefuse(parDocID)
  Dim dsTemp
'ph - 20090910 - отказ в согласовании с последующим нажатием перечеркнутого отказа тоже считать отказом
'  sSQL = "select Count(*) as RefuseCount from Comments where DocID = N'"+parDocID+"' and (CommentType = 'VISA' and SpecialInfo = 'VISAOKREFUSE' or CommentType = 'APROVAL' and (SpecialInfo = 'DOCS_RefusedApp' or CharIndex(N'Schválení zamítnuto', Comment) = 1 or CharIndex(N'Approval refused', Comment) = 1 or CharIndex(N'Отказано в утверждении', Comment) = 1))"
  sSQL = "select Count(*) as RefuseCount from Comments where DocID = N'"+parDocID+"' and (CommentType = 'VISA' and CharIndex('VISAOKREFUSE', SpecialInfo) = 1 or CommentType = 'APROVAL' and (SpecialInfo = 'DOCS_RefusedApp' or CharIndex(N'Schválení zamítnuto', Comment) = 1 or CharIndex(N'Approval refused', Comment) = 1 or CharIndex(N'Отказано в утверждении', Comment) = 1))"
AddLogD "GetRefuse SQL: "+sSQL
  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  dsTemp.Open sSQL, Conn, 3, 1, &H1
  GetRefuse = dsTemp("RefuseCount")
  dsTemp.Close
End Function

'Определить число уровней согласования
Function GetReconciliationLevels(ByVal parListToReconcile)
  Dim i, arLevels
  
  GetReconciliationLevels = 0
  If Trim(parListToReconcile) <> "" Then
    arLevels = Split(parListToReconcile, VbCrLf)
    For i = 0 To UBound(arLevels)
      If Trim(arLevels(i)) <> "" and GetUserID(arLevels(i)) <> "" Then
        GetReconciliationLevels = GetReconciliationLevels+1
      End If
    Next
  End If
End Function

'Показать разницу даты начала и указанной даты в рабочих днях (вызывать после GetFirstActivationDate, где устанавливается SIT_ReportDateActivation)
Function ShowWorkingDaysFromStartToDate(parDateActivation, DateTill)
  Dim dDateActivation

  If SIT_ReportDateActivation = VAR_BeginOfTimes Then
    dDateActivation = parDateActivation
  Else
    dDateActivation = SIT_ReportDateActivation
  End If

  If IsNull(DateTill) or IsNull(DateFrom) or DateTill = VAR_BeginOfTimes or DateFrom = VAR_BeginOfTimes Then
    ShowWorkingDaysFromStartToDate = ""
  Else
    ShowWorkingDaysFromStartToDate = CStr(WorkingDaysBetweenTwoDates(dDateActivation, DateTill))
  End If
End Function

'Показать наличие просрочки в визировании (вызывать после GetFirstActivationDate, где устанавливается SIT_ReportDateActivation)
Function ShowApprovalProcessOnTime(parDateActivation, parDateApproved, parNameApproved, parListToReconcile, parStatusCompletion)
  Dim EndDate, bActual, sColor, dDateActivation
  If SIT_ReportDateActivation = VAR_BeginOfTimes Then
    dDateActivation = parDateActivation
  Else
    dDateActivation = SIT_ReportDateActivation
  End If
  
  bActual = IsNull(parDateApproved) or parDateApproved = VAR_BeginOfTimes
  If bActual Then
    EndDate = Date
  Else
    EndDate = parDateApproved
  End If

  If bActual Then
    If not IsNull(parStatusCompletion) and parStatusCompletion = VAR_StatusCancelled Then
      ShowApprovalProcessOnTime = "<font color=black><b>—</b></font>" 'Документ отменен до утверждения
      Exit Function
    Else
      sColor = "FFB000" 'В процессе
    End If
  Else
    If IsRefused(parNameApproved) Then
      sColor = "red" 'Отклонено
    Else
      sColor = "green" 'Утверждено
    End If
  End If

  ShowApprovalProcessOnTime = "<font color="+sColor+"><b>"+iif(WorkingDaysBetweenTwoDates(dDateActivation, EndDate) <= (GetReconciliationLevels(parListToReconcile)+1)*3,"Yes","No")+"</b></font>"
End Function

'Показать просрочку при визировании (вызывать после GetFirstActivationDate, где устанавливается SIT_ReportDateActivation)
Function ShowDelayInApproval(parDateActivation, parDateApproved, parNameApproved, parListToReconcile, parStatusCompletion)
  Dim EndDate, bActual, sColor, iDelay, dDateActivation

  If SIT_ReportDateActivation = VAR_BeginOfTimes Then
    dDateActivation = parDateActivation
  Else
    dDateActivation = SIT_ReportDateActivation
  End If
  
  bActual = IsNull(parDateApproved) or parDateApproved = VAR_BeginOfTimes
  If bActual Then
    EndDate = Date
  Else
    EndDate = parDateApproved
  End If

  If bActual Then
    If not IsNull(parStatusCompletion) and parStatusCompletion = VAR_StatusCancelled Then
      ShowDelayInApproval = "<font color=black><b>—</b></font>" 'Документ отменен до утверждения
      Exit Function
    Else
      sColor = "FFB000" 'В процессе
    End If
  Else
    If IsRefused(parNameApproved) Then
      sColor = "red" 'Отклонено
    Else
      sColor = "green" 'Утверждено
    End If
  End If
  
  iDelay = WorkingDaysBetweenTwoDates(dDateActivation, EndDate) - (GetReconciliationLevels(parListToReconcile)+1)*3

  ShowDelayInApproval = "<font color="+sColor+"><b>"+iif(iDelay > 0,CStr(iDelay),"0")+"</b></font>"
End Function

'rmanyushin 88625 31.03.2010 Start
'Заменяет функцию ShowDelayInApproval в отчетах. Отличается тем, что в workflow еще учитывает и утверждающего (аналогично функциям в HLR), а не просто + 1. 
Function ShowDelayInApproval2(parDateActivation, parDateApproved, parNameApproved, parNameAproval, parListToReconcile, parStatusCompletion)
  Dim EndDate, bActual, sColor, iDelay, dDateActivation

  If SIT_ReportDateActivation = VAR_BeginOfTimes Then
    dDateActivation = parDateActivation
  Else
    dDateActivation = SIT_ReportDateActivation
  End If
  
  bActual = IsNull(parDateApproved) or parDateApproved = VAR_BeginOfTimes
  If bActual Then
    EndDate = Date
  Else
    EndDate = parDateApproved
  End If

  If bActual Then
    If not IsNull(parStatusCompletion) and parStatusCompletion = VAR_StatusCancelled Then
      ShowDelayInApproval = "<font color=black><b>—</b></font>" 'Документ отменен до утверждения
      Exit Function
    Else
      sColor = "FFB000" 'В процессе
    End If
  Else
    If IsRefused(parNameApproved) Then
      sColor = "red" 'Отклонено
    Else
      sColor = "green" 'Утверждено
    End If
  End If
  
  If parNameAproval <> "" Then
      ApprovalLevel = 1
	Else  
	  ApprovalLevel = 0
  End If
  
  iDelay = WorkingDaysBetweenTwoDates(dDateActivation, EndDate) - (GetReconciliationLevels(parListToReconcile)+ ApprovalLevel)*3
  
  ShowDelayInApproval2 = "<font color="+sColor+"><b>"+iif(iDelay > 0,CStr(iDelay),"0")+"</b></font>"
End Function
'rmanyushin 31.03.2010 End


'Определить число делегирований и запросов рецензий
Function GetDelegation(parDocID)
  Dim dsTemp
  sSQL = "select Count(*) as DelegationCount from Comments where DocID = N'"+parDocID+"' and (CommentType = 'VISA' and SpecialInfo = 'DELEGATE' or CommentType = 'REVIEW' and Address = 'REQUEST')"
AddLogD "GetDelegation SQL: "+sSQL
  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  dsTemp.Open sSQL, Conn, 3, 1, &H1
  GetDelegation = dsTemp("DelegationCount")
  dsTemp.Close
End Function

'Определить дату последнего согласования
Function ShowLastVisaDate(parDocID)
  Dim dsTemp, Duration, MaxDuration
  sSQL = "select * from Comments where DocID = N'"+parDocID+"' and CommentType = 'VISA' and (SpecialInfo = 'VISAOK' or SpecialInfo = 'VISAOKREFUSE') order by DateCreation desc"
AddLogD "ShowLastVisaDate SQL: "+sSQL
  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  dsTemp.Open sSQL, Conn, 3, 1, &H1
  If dsTemp.EOF Then
    ShowLastVisaDate = ""
  Else
    ShowLastVisaDate = DateInDDMMYYYY(dsTemp("DateCreation"))
  End If
  dsTemp.Close
End Function

'Посчитать число уникальных согласующих/утверждающих
Function CountUniqueReviewers(parListToReconcile, parNameAproval)
  Dim iPos, sUserID, sLogins
  iPos = 1
  sLogins = VbCrLf
  CountUniqueReviewers = 0
  sUserID = oPayDox.GetNextUserIDInList(parListToReconcile, iPos)
  Do While sUserID <> ""
    If InStr(sLogins, VbCrLf+sUserID+VbCrLf) = 0 Then
      sLogins = sLogins+sUserID+VbCrLf
      CountUniqueReviewers = CountUniqueReviewers+1
    End If
    sUserID = oPayDox.GetNextUserIDInList(parListToReconcile, iPos)
  Loop
  sUserID = GetUserID(parNameAproval)
  If sUserID <> "" Then
    If InStr(sLogins, VbCrLf+sUserID+VbCrLf) = 0 Then
      CountUniqueReviewers = CountUniqueReviewers+1
    End If
  End If
End Function

'Определить время последнего утверждения/отказа в утв. В parRealNameApproved возвращается пользователь, который был под ролью
Function GetLastApprovedDate(parDocID, parDateApproved, ByRef parRealNameApproved)
  Dim dsTemp

  sSQL = "select top 1 DateCreation, Comment from Comments where DocID = N'"+parDocID+"' and CommentType = 'APROVAL' and (SpecialInfo = 'DOCS_RefusedApp' or SpecialInfo = 'DOCS_Approved' or CharIndex(N'Schválení zamítnuto - ', Comment) = 1 or CharIndex(N'Schváleno - ', Comment) = 1 or CharIndex(N'Approval refused - ', Comment) = 1 or CharIndex(N'Approved - ', Comment) = 1 or CharIndex(N'Отклонено - ', Comment) = 1 or CharIndex(N'Подписано - ', Comment) = 1 or CharIndex(N'Отказано в утверждении - ', Comment) = 1 or CharIndex(N'Утверждено - ', Comment) = 1) order by DateCreation desc"
AddLogD "GetLastApprovedDate SQL: "+sSQL
  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  dsTemp.Open sSQL, Conn, 3, 1, &H1
  If dsTemp.EOF Then
    GetLastApprovedDate = parDateApproved
    parRealNameApproved = ""
  Else
    GetLastApprovedDate = dsTemp("DateCreation")
    parRealNameApproved = GetRealUser(dsTemp("Comment"))
  End If
  dsTemp.Close
End Function

'Определение дня недели (1 - понедельник, 7 - воскресенье) для SQL-запросов
Function SQL_CorrectWeekDay(parDate)
  SQL_CorrectWeekDay = " Case DatePart(weekday, { d '2006-01-01' }) When 7 Then DatePart(weekday, "+parDate+") Else Case DatePart(weekday, "+parDate+") When 1 Then 7 Else DatePart(weekday, "+parDate+")-1 End End "
End Function

'Определение числа рабочих дней между двумя датами для SQL-запросов
Function SQL_WorkingDaysBetweenTwoDates(parDateStart, parDateEnd)
  FullWeekDays = " ROUND(DateDiff(day, "+parDateStart+", "+parDateEnd+")/7, 0, 1)*5 "
  AddWeekDays = " DateDiff(day, "+parDateStart+", "+parDateEnd+") % 7 "
  SQL_WorkingDaysBetweenTwoDates = FullWeekDays+" + "+AddWeekDays+" - Case When "+SQL_CorrectWeekDay(parDateEnd)+" >= "+SQL_CorrectWeekDay(parDateStart)+" Then 0 Else Case When "+SQL_CorrectWeekDay(parDateStart)+" = 7 Then 1 Else 2 End End"
End Function

'Определить кому было делегировано согласование
Function GetDelegationTo(parStr)
  Dim iPos1, iPos2
  
  GetDelegationTo = ""
  If IsNull(parStr) Then
    Exit Function
  End If
  iPos1 = InStr(parStr, """")
  iPos2 = InStr(parStr, ">")
  If iPos2 > iPos1 Then
    GetDelegationTo = Mid(parStr, iPos1, iPos2-iPos1+1)
  End If
End Function

'Определить реального пользователя, если работа была под ролью
Function GetRealUser(sComment)
  Dim arParts, iLastPart

  GetRealUser = ""
  arParts = Split(MyCStr(sComment), "/")
  iLastPart = UBound(arParts)
  If iLastPart > 0 Then
'Ph - 20091202 - start
    arParts(iLastPart) = Trim(arParts(iLastPart))
'    If GetUserID(arParts(iLastPart)) <> "" Then
    If InStr(arParts(iLastPart), """") = 1 and InStr(arParts(iLastPart), ">") = Len(arParts(iLastPart)) Then
'      GetRealUser = " / " + Trim(DelOtherLangFromNames(arParts(iLastPart)))
      GetRealUser = " / " + DelOtherLangFromNames(arParts(iLastPart))
'Ph - 20091202 - end
    End If
  End If
End Function

'Определить время самого долгого согласования/утверждения, в параметре LongestVisaUser возвращается ответственный за самую длинную просрочку
Function GetLongestVisaTime(parDocID, parDateActivation, parDateApproved, parNameAproval, parNameApproved, parStatusCompletion, parDateCompleted, parListReconciled, ByRef LongestVisaUser)
  Dim bDocumentCancelled
  Dim dsTemp
  Dim arDelegationTo(), arDelegationFrom(), arDelegationWhen(), iDelegations, i
  Dim sSQL_EndDate, dEndDate, iDuration, iMaxDuration, sUser, sMaxUser
  Dim dLastReconciliationDate, bIsVisaWaiting
  Dim sRealNameApproved
  
  bDocumentCancelled = not IsNull(parStatusCompletion) and parStatusCompletion = VAR_StatusCancelled

  sSQL_EndDate = "Case When SpecialInfo = 'VISAWAITING' Then Case When '"+iif(bDocumentCancelled,"0","")+"'='0' Then "+UniDate(parDateCompleted)+" Else GetDate() End Else DateEventEnd End"
  sSQL = "select Case When SpecialInfo = 'DELEGATE' Then -1 Else "+SQL_WorkingDaysBetweenTwoDates("DateEvent", sSQL_EndDate)+" End as Duration,"+sSQL_EndDate+" as EndDate,* from Comments where DocID = N'"+parDocID+"' and CommentType = 'VISA' and (SpecialInfo = 'DELEGATE' or SpecialInfo = 'VISAWAITING' or SpecialInfo = 'VISAOK' or SpecialInfo = 'VISAOKREFUSE')"
AddLogD "GetLongestVisaTime SQL: "+sSQL
  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  dsTemp.CursorLocation = 3
  dsTemp.Open sSQL, Conn, 3, 1, &H1
  dsTemp.ActiveConnection = Nothing

  dsTemp.Filter = "SpecialInfo = 'DELEGATE'"
  dsTemp.Sort = "DateCreation desc"
  If not dsTemp.EOF Then
    dsTemp.MoveFirst
  End If
  iDelegations = dsTemp.RecordCount
  If iDelegations > 0 Then
    ReDim arDelegationTo(iDelegations)
    ReDim arDelegationFrom(iDelegations)
    ReDim arDelegationWhen(iDelegations)
    i = 1
    Do While not dsTemp.EOF
      arDelegationTo(i) = GetDelegationTo(dsTemp("Comment"))
      arDelegationFrom(i) = GetFullName(SurnameGN(GetEngNameFromFolder(MyCStr(dsTemp("UserName")))), MyCStr(dsTemp("UserID")))
      arDelegationWhen(i) = dsTemp("DateCreation")
      dsTemp.MoveNext
AddLogD "Delegation: "+MyDate(arDelegationWhen(i))+" - "+arDelegationFrom(i)+" -> "+arDelegationTo(i)
      i = i + 1
    Loop
  End If

  dsTemp.Filter = "SpecialInfo = 'VISAWAITING' or SpecialInfo = 'VISAOK' or SpecialInfo = 'VISAOKREFUSE'"
  dsTemp.Sort = "Duration desc, DateCreation"
  If not dsTemp.EOF Then
    dsTemp.MoveFirst
  End If
  iMaxDuration = -1
  sMaxUser = ""
  Do While not dsTemp.EOF
    If iMaxDuration >= dsTemp("Duration") Then
      Exit Do
    Else
      If iDelegations = 0 Then
        iMaxDuration = dsTemp("Duration")
        sMaxUser = GetFullName(SurnameGN(GetEngNameFromFolder(MyCStr(dsTemp("UserName")))), MyCStr(dsTemp("UserID")))+GetRealUser(dsTemp("Comment"))
      Else
        dEndDate = dsTemp("EndDate")
        sUser = GetFullName(SurnameGN(GetEngNameFromFolder(MyCStr(dsTemp("UserName")))), MyCStr(dsTemp("UserID")))
        For i = 1 To UBound(arDelegationTo)
          If GetUserID(arDelegationTo(i)) = GetUserID(sUser) and arDelegationWhen(i) >= dsTemp("DateEvent") and arDelegationWhen(i) <= dEndDate Then
            iDuration = WorkingDaysBetweenTwoDates(arDelegationWhen(i), dEndDate)
            If iDuration > iMaxDuration Then
              iMaxDuration = iDuration
              sMaxUser = sUser+GetRealUser(dsTemp("Comment"))
            End If
            sUser = arDelegationFrom(i)
            dEndDate = arDelegationWhen(i)
          End If
        Next
        'Время затраченное первоначальным согласующим
        iDuration = WorkingDaysBetweenTwoDates(dsTemp("DateEvent"), dEndDate)
        If iDuration > iMaxDuration Then
          iMaxDuration = iDuration
          sMaxUser = sUser+GetRealUser(dsTemp("Comment"))
        End If
      End If
    End If
    dsTemp.MoveNext
  Loop

  dsTemp.Sort = "DateCreation desc"
  If not dsTemp.EOF Then
    dsTemp.MoveFirst
    dLastReconciliationDate = dsTemp("DateCreation")
    bIsVisaWaiting = dsTemp("SpecialInfo") = "VISAWAITING"
  Else
    bIsVisaWaiting = False
    dLastReconciliationDate = GetFirstActivationDate(parDocID, parDateActivation)
  End If

  dsTemp.Close
  'Утверждение
'ph - 20090911 - добавлено условие отсутствия отказа в утверждении (если отказ есть, то виновный ищется только среди согласовавших, утверждение в расчет не берется)
  If not bIsVisaWaiting and (not IsNull(parNameAproval) or Trim(parNameAproval) <> "") and InStr(parListReconciled, "-<") = 0 Then
'  If not bIsVisaWaiting and (not IsNull(parNameAproval) or Trim(parNameAproval) <> "") Then
    If IsNull(parNameApproved) or Trim(parNameApproved) = "" Then
      If bDocumentCancelled Then
        dEndDate = parDateCompleted 'Для отмененных документов дата окончания - дата отмены
      Else
        dEndDate = Date()
      End If
      sRealNameApproved = ""
    Else
      dEndDate = GetLastApprovedDate(parDocID, parDateApproved, sRealNameApproved)
      If IsNull(dEndDate) or dEndDate = VAR_BeginOfTimes Then
        If bDocumentCancelled Then
          dEndDate = parDateCompleted
        Else
          dEndDate = Date()
        End If
      End If
    End If
    iDuration = WorkingDaysBetweenTwoDates(dLastReconciliationDate, dEndDate)
    If iDuration > iMaxDuration Then
      iMaxDuration = iDuration
      sMaxUser = parNameAproval+sRealNameApproved
    End If
  End If
  
  If iMaxDuration = -1 Then
    GetLongestVisaTime = "—"
    LongestVisaUser = "—"
  Else
    GetLongestVisaTime = CStr(iMaxDuration)
    LongestVisaUser = sMaxUser
  End If
End Function

'Переменная для сохранения пользователя, ответственного за самую длинную просрочку
SIT_LongestVisaUser = ""

'Аналог GetLongestVisaTime для вызова из отчета, сохраняющая ответственного за самую длинную просрочку в SIT_LongestVisaUser
Function GetLongestVisaTimeForReport(parDocID, parDateActivation, parDateApproved, parNameAproval, parNameApproved, parStatusCompletion, parDateCompleted, parListReconciled)
  GetLongestVisaTimeForReport = GetLongestVisaTime(parDocID, parDateActivation, parDateApproved, parNameAproval, parNameApproved, parStatusCompletion, parDateCompleted, parListReconciled, SIT_LongestVisaUser)
End Function

'Получить виноватого в самой большой просрочке, можно вызывать только после GetLongestVisaTimeForReport, т.к. определяется он там
Function GetLongestVisaUser()
  GetLongestVisaUser = HTMLEncode(SIT_LongestVisaUser)
End Function

'Определить подразделение пользователя, ответственного за просрочку (если была работа под ролью, то подразделение того, кто был под ролью)
Function GetRealLongestVisaUserDepartment()
  Dim arParts, sUser

  GetRealLongestVisaUserDepartment = ""
  If MyCStr(SIT_LongestVisaUser) = "" Then
    Exit Function
  End If
  arParts = Split(MyCStr(SIT_LongestVisaUser), "/")
  sUser = Trim(arParts(UBound(arParts)))
  If GetUserID(sUser) = "" Then
    Exit Function
  End If
  GetRealLongestVisaUserDepartment = MyGetUserDepartment(sUser)
End Function

'Получить виноватого в самой большой просрочке, сама вызывает GetLongestVisaTimeForReport
Function GetLongestVisaUser2(parDocID, parDateActivation, parDateApproved, parNameAproval, parNameApproved, parStatusCompletion, parDateCompleted, parListReconciled)
  Dim iLongestVisaTime, sLongestVisaUser
  
  iLongestVisaTime = GetLongestVisaTime(parDocID, parDateActivation, parDateApproved, parNameAproval, parNameApproved, parStatusCompletion, parDateCompleted, parListReconciled, sLongestVisaUser)
  GetLongestVisaUser2 = HTMLEncode(sLongestVisaUser)
End Function

'Заявка СТС №6 (19.10.2009) - start
'Определить время самого долгого текущего согласования/утверждения, в параметре LongestVisaUser возвращаются те, кто дольше всех держат документ
Function GetActualLongestVisaTime(parDocID, parDateActivation, parNameAproval, parNameApproved, parStatusCompletion, parListReconciled, ByRef LongestVisaUser)
  Dim dsTemp
  Dim arDelegationTo(), arDelegationFrom(), arDelegationWhen(), iDelegations, i
  Dim sSQL_EndDate, dEndDate, iDuration, iMaxDuration, sUser, sMaxUser
  Dim dLastReconciliationDate, bIsVisaWaiting
  
  'Проверка на отмену/исполненность документа
  If not IsNull(parStatusCompletion) and (parStatusCompletion = VAR_StatusCancelled or parStatusCompletion = VAR_StatusCompletion) Then
    'Документ никто не держит
    LongestVisaUser = ""
    GetActualLongestVisaTime = ""
    Exit Function
  End If
  'Есть отказ в согласовании -> документ никто не держит, он исправляется
  If InStr(parListReconciled, "-<") <> 0 Then
    LongestVisaUser = ""
    GetActualLongestVisaTime = ""
    Exit Function
  End If

  sSQL_EndDate = "Case When SpecialInfo = 'VISAWAITING' Then GetDate() Else DateEventEnd End"
  sSQL = "select Case When SpecialInfo = 'DELEGATE' Then -1 Else "+SQL_WorkingDaysBetweenTwoDates("DateEvent", sSQL_EndDate)+" End as Duration,"+sSQL_EndDate+" as EndDate,* from Comments where DocID = N'"+parDocID+"' and CommentType = 'VISA' and (SpecialInfo = 'DELEGATE' or SpecialInfo = 'VISAWAITING' or SpecialInfo = 'VISAOK' or SpecialInfo = 'VISAOKREFUSE')"
AddLogD "GetActualLongestVisaTime SQL: "+sSQL
  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  dsTemp.CursorLocation = 3
  dsTemp.Open sSQL, Conn, 3, 1, &H1
  dsTemp.ActiveConnection = Nothing

  dsTemp.Filter = "SpecialInfo = 'VISAWAITING'"
  If dsTemp.EOF Then
    'Документ не находится на согласовании
    bIsVisaWaiting = False

    dsTemp.Filter = "SpecialInfo = 'VISAOK' or SpecialInfo = 'VISAOKREFUSE'"
    If dsTemp.EOF Then
      'Документ вообще не согласовывался
      dLastReconciliationDate = GetFirstActivationDate(parDocID, parDateActivation)
      iMaxDuration = -1
      sMaxUser = ""
    Else
      dsTemp.Sort = "DateCreation desc"
      dsTemp.MoveFirst
      dLastReconciliationDate = dsTemp("DateCreation")
      iMaxDuration = -1
      sMaxUser = ""
    End If
  Else
    'Документ на согласовании
    bIsVisaWaiting = True

    'Учет делегирования
    dsTemp.Filter = "SpecialInfo = 'DELEGATE'"
    dsTemp.Sort = "DateCreation desc"
    If not dsTemp.EOF Then
      dsTemp.MoveFirst
    End If
    iDelegations = dsTemp.RecordCount
    If iDelegations > 0 Then
      ReDim arDelegationTo(iDelegations)
      ReDim arDelegationFrom(iDelegations)
      ReDim arDelegationWhen(iDelegations)
      i = 1
      Do While not dsTemp.EOF
        arDelegationTo(i) = GetDelegationTo(dsTemp("Comment"))
        arDelegationFrom(i) = GetFullName(SurnameGN(GetEngNameFromFolder(MyCStr(dsTemp("UserName")))), MyCStr(dsTemp("UserID")))
        arDelegationWhen(i) = dsTemp("DateCreation")
        dsTemp.MoveNext
AddLogD "Delegation: "+MyDate(arDelegationWhen(i))+" - "+arDelegationFrom(i)+" -> "+arDelegationTo(i)
        i = i + 1
      Loop
    End If

    dsTemp.Filter = "SpecialInfo = 'VISAWAITING'"
    dsTemp.Sort = "Duration desc, DateCreation"
    dsTemp.MoveFirst
    iMaxDuration = -1
    sMaxUser = ""
    Do While not dsTemp.EOF
      If iMaxDuration > dsTemp("Duration") Then
        Exit Do
      Else
        If iDelegations = 0 Then
          sUser = GetFullName(SurnameGN(GetEngNameFromFolder(MyCStr(dsTemp("UserName")))), MyCStr(dsTemp("UserID")))+GetRealUser(dsTemp("Comment"))
          If iMaxDuration = dsTemp("Duration") Then
            If sMaxUser <> "" Then
              sMaxUser = sMaxUser+VbCrLf
            End If
            sMaxUser = sMaxUser+sUser
          Else
            iMaxDuration = dsTemp("Duration")
            sMaxUser = sUser
          End If
        Else
          dEndDate = dsTemp("EndDate")
          sUser = GetFullName(SurnameGN(GetEngNameFromFolder(MyCStr(dsTemp("UserName")))), MyCStr(dsTemp("UserID")))
          For i = 1 To UBound(arDelegationTo)
            If GetUserID(arDelegationTo(i)) = GetUserID(sUser) and arDelegationWhen(i) >= dsTemp("DateEvent") and arDelegationWhen(i) <= dEndDate Then
              iDuration = WorkingDaysBetweenTwoDates(arDelegationWhen(i), dEndDate)
              If iDuration = iMaxDuration Then
                If sMaxUser <> "" Then
                  sMaxUser = sMaxUser+VbCrLf
                End If
                sMaxUser = sMaxUser+sUser+GetRealUser(dsTemp("Comment"))
              ElseIf iDuration > iMaxDuration Then
                iMaxDuration = iDuration
                sMaxUser = sUser+GetRealUser(dsTemp("Comment"))
              End If
              sUser = arDelegationFrom(i)
              dEndDate = arDelegationWhen(i)
            End If
          Next
          'Время затраченное первоначальным согласующим
          iDuration = WorkingDaysBetweenTwoDates(dsTemp("DateEvent"), dEndDate)
          If iDuration = iMaxDuration Then
            If sMaxUser <> "" Then
              sMaxUser = sMaxUser+VbCrLf
            End If
            sMaxUser = sMaxUser+sUser+GetRealUser(dsTemp("Comment"))
          ElseIf iDuration > iMaxDuration Then
            iMaxDuration = iDuration
            sMaxUser = sUser+GetRealUser(dsTemp("Comment"))
          End If
        End If
      End If
      dsTemp.MoveNext
    Loop
  End If
  
  dsTemp.Close
  
  'Утверждение
  If not bIsVisaWaiting and (not IsNull(parNameAproval) or Trim(parNameAproval) <> "") Then
    If IsNull(parNameApproved) or Trim(parNameApproved) = "" Then
      'Документ на утверждении
      dEndDate = Date()
    Else 'Документ утвержден/отклонен -> его никто не держит
      GetActualLongestVisaTime = ""
      LongestVisaUser = ""
      Exit Function
    End If
    iDuration = WorkingDaysBetweenTwoDates(dLastReconciliationDate, dEndDate)
    If iDuration > iMaxDuration Then
      iMaxDuration = iDuration
      sMaxUser = parNameAproval
    End If
  End If
  
  If iMaxDuration = -1 Then
    GetActualLongestVisaTime = ""
    LongestVisaUser = ""
  Else
    GetActualLongestVisaTime = CStr(iMaxDuration)
    LongestVisaUser = sMaxUser
  End If
End Function

'Переменная для сохранения пользователей, ответственных за самое долгое текущее согласование/утверждение
SIT_ActualLongestVisaUser = ""

'Аналог GetActualLongestVisaTime для вызова из отчета, сохраняющая ответственных за самое долгое текущее согласование/утверждение в SIT_ActualLongestVisaUser
Function GetActualLongestVisaTimeForReport(parDocID, parDateActivation, parNameAproval, parNameApproved, parStatusCompletion, parListReconciled)
  GetActualLongestVisaTimeForReport = GetActualLongestVisaTime(parDocID, parDateActivation, parNameAproval, parNameApproved, parStatusCompletion, parListReconciled, SIT_ActualLongestVisaUser)
End Function

'Получить виноватых в самом долгом текущем согласовании/утверждении, можно вызывать только после GetActualLongestVisaTimeForReport, т.к. определяются они там
Function GetActualLongestVisaUser()
'  GetActualLongestVisaUser = HTMLEncode(SIT_ActualLongestVisaUser)
  GetActualLongestVisaUser=Replace(SIT_ActualLongestVisaUser,"<","&lt;")
  GetActualLongestVisaUser=Replace(GetActualLongestVisaUser,">","&gt;")
  GetActualLongestVisaUser=Replace(GetActualLongestVisaUser, VbCrLf,"<br>")
End Function
'Заявка СТС №6 (19.10.2009) - end

'Заявка СТС №7 (10.11.2009) - start
'Получить департамент или дивизион СТС (с движением вверх, если указано подразделение более низкого уровня)
Function GetDepartmentOrDivision(parDepartment)
  Dim dsTemp, sTemp, sSQL

  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  dsTemp.CursorLocation = 3
  sSQL = "select * from Departments where CharIndex(Name, N'"+Replace(parDepartment, "'", "''")+"') > 0 order by Name desc"
AddLogD "GetDepartmentOrDivision - sSQL: " + sSQL
  dsTemp.Open sSQL, Conn, 3, 1, &H1
  dsTemp.ActiveConnection = Nothing
  
  GetDepartmentOrDivision = ""
  'Проверяем существует ли заданное подразделение. Если нет, возвращаем пустое значение
  If dsTemp.EOF Then
    dsTemp.Close
AddLogD "GetDepartmentOrDivision - department not found"
    Exit Function
  ElseIf UCase(dsTemp("Name")) <> UCase(parDepartment) Then
    dsTemp.Close
AddLogD "GetDepartmentOrDivision - department name doesn't match"
    Exit Function
  End If
  Do While InStr(MyCStr(dsTemp("Statuses")), "#LEV1") = 0 and InStr(MyCStr(dsTemp("Statuses")), "#LEV2") = 0
    dsTemp.MoveNext
    If dsTemp.EOF Then
      dsTemp.Close
AddLogD "GetDepartmentOrDivision - no department with required level"
      Exit Function
    End If
  Loop

  GetDepartmentOrDivision = dsTemp("Name")
  dsTemp.Close
End Function

'Определить подразделения пользователей, ответственных за самое долгое текущее согласование/утверждение
Function GetActualLongestVisaUserDepartment()
  Dim sUserID, iPos, sDepartmentOrDivision

  GetActualLongestVisaUserDepartment = ""
  If MyCStr(SIT_ActualLongestVisaUser) = "" Then
    Exit Function
  End If
  iPos = 1
  sUserID = oPayDox.GetNextUserIDInList(SIT_ActualLongestVisaUser, iPos)
  Do While sUserID <> ""
    If GetActualLongestVisaUserDepartment <> "" Then
      GetActualLongestVisaUserDepartment = GetActualLongestVisaUserDepartment + VbCrLf
    End If
    oPayDox.GetUserDetails sUserID, sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
    sDepartmentOrDivision = GetDepartmentOrDivision(sDepartment)
    If sDepartmentOrDivision = "" Then
      'Если департамента/дивизиона не найдено, ставим подразделение пользователя и выделяем красным
      sDepartmentOrDivision = "<font color=red>"+GetLastDepartmentLevelEng(sDepartment)+"</font>"
    Else
      sDepartmentOrDivision = GetLastDepartmentLevelEng(sDepartmentOrDivision)
    End If
    GetActualLongestVisaUserDepartment = GetActualLongestVisaUserDepartment+sDepartmentOrDivision
    sUserID = oPayDox.GetNextUserIDInList(SIT_ActualLongestVisaUser, iPos)
  Loop
  'Преобразование концов строк для вывода
  GetActualLongestVisaUserDepartment=Replace(GetActualLongestVisaUserDepartment, VbCrLf,"<br>")
End Function

'Получить число рабочих дней с момента приостановки в согласовании из-за отказа
Function GetActualRefusal(parDocID, parListToReconcile, parListReconciled, parNameApproved)
  Dim dsTemp, dDateRefused, sSQL

  dDateRefused = VAR_BeginOfTimes

  If (oPayDox.IsReconciliationPendingWithOptions(parListToReconcile, parListReconciled, Var_ReconciliationIfAllAgree) and not IsVisaNowWithOptions(Session("UserID"), parListToReconcile, parListReconciled, parNameApproved)) or oPayDox.IsReconciliationWaitingAfterRefuse(parListToReconcile, parListReconciled, Var_ReconciliationIfAllAgree) Then
IF FALSE THEN
    Set dsTemp = Server.CreateObject("ADODB.Recordset")
    dsTemp.CursorLocation = 3
    sSQL = "select * from Comments where DocID = N'"+Replace(parDocID, "'", "''")+"' and CommentType = 'VISA' order by DateCreation desc"
AddLogD "GetActualRefusal - sSQL: " + sSQL
    dsTemp.Open sSQL, Conn, 3, 1, &H1
    dsTemp.ActiveConnection = Nothing
    Do While not dsTemp.EOF
      If dsTemp("SpecialInfo") = "VISAOKREFUSE" Then
        dDateRefused = dsTemp("DateCreation")
      ElseIf dsTemp("SpecialInfo") = "DOCS_Changed" or dsTemp("SpecialInfo") = "INFOCHANGED" or _
      InStr(dsTemp("Comment"), "Согласование отменено") = 1 or InStr(dsTemp("Comment"), "Agree cancelled") = 1 or InStr(dsTemp("Comment"), "Odsouhlasení bylo zrušeno") = 1 or _
      InStr(dsTemp("Comment"), "Повторное согласование") = 1 or InStr(dsTemp("Comment"), "Repeated agree") = 1 or InStr(dsTemp("Comment"), "Znovu odsouhlasit") = 1 or _
      InStr(dsTemp("Comment"), "Информация сохранена") = 1 or InStr(dsTemp("Comment"), "Information updated") = 1 or InStr(dsTemp("Comment"), "Informace byly změněny") = 1 Then
        Exit Do
      End If
      dsTemp.MoveNext
    Loop
    dsTemp.Close
END IF 'FALSE
'Другой способ
    iPos = InStr(parListReconciled, "-<")
    If iPos > 0 Then
      sFirstRefusedUserID = oPayDox.GetNextUserIDInList(parListReconciled, iPos)
      Set dsTemp = Server.CreateObject("ADODB.Recordset")
      sSQL = "select * from Comments where DocID = N'"+Replace(parDocID, "'", "''")+"' and UserID = N'"+Replace(sFirstRefusedUserID, "'", "''")+"' and SpecialInfo = 'VISAOKREFUSE' order by DateCreation desc"
AddLogD "GetActualRefusal - sSQL: " + sSQL
      dsTemp.Open sSQL, Conn, 3, 1, &H1
      If not dsTemp.EOF Then
        dDateRefused = dsTemp("DateCreation")
      End If
      dsTemp.Close
    End If
  End If
  
  If dDateRefused = VAR_BeginOfTimes Then
    GetActualRefusal = ""
  Else
    GetActualRefusal = CStr(WorkingDaysBetweenTwoDates(dDateRefused, Date()))
  End If
End Function

'Определить задержку в выполнении поручения (в рабочих днях)
Function GetCompletionDelay(parDateCompletion, parDateCompleted, parStatusCompletion)
  Dim dDateCompleted
  
  If MyCStr(parStatusCompletion)<>VAR_StatusCompletion And MyCStr(parStatusCompletion)<>VAR_StatusCancelled And IsNull(parDateCompleted) Then
    dDateCompleted = Date()
  Else
    dDateCompleted = parDateCompleted
  End If
  If dDateCompleted > parDateCompletion Then
    GetCompletionDelay = CStr(WorkingDaysBetweenTwoDates(parDateCompletion, dDateCompleted))
  Else
    If IsNull(parDateCompleted) Then
    GetCompletionDelay = ""
    Else
      GetCompletionDelay = "0"
    End If
  End If
End Function

'Заявка СТС №7 (10.11.2009) - end

'rmanyushin 51555, 56781, 79501, 133266 05.10.2010 Start
' Проверка, является ли пользователь привелигированным пользователем СТС - Контролер СТС или Аудитор СТС или Юрист СТС.
Function isPrivilegedUserSTS
	If UCase(Session("UserID")) = UCase(STS_Auditor) or UCase(Session("UserID")) = UCase(STS_Overseer) or UCase(Session("UserID")) = UCase(STS_HeadOf789) or UCase(Session("UserID")) = UCase(STS_POViewer) or UCase(Session("UserID")) = UCase(STS_LegalSTS) Then
		isPrivilegedUserSTS = True
	Else
		isPrivilegedUserSTS = False
	End If
End Function
'rmanyushin 51555, 56781, 79501, 133266 05.10.2010 End

'rmanyushin 56781 13.10.2009 Start
Function is789DivisionSTS(strDepartmentName)
	Set objRegExp = CreateObject("VBScript.RegExp")
	objRegExp.Global = False
	objRegExp.Pattern = "\d{5}"
	Set objMatches = objRegExp.Execute(strDepartmentName)
    
    If objMatches.Count <> 0 Then
        strDivisionID = objMatches.Item(0).Value
		strDivisionID = Trim(strDivisionID)
        strDivisionID = Left(strDivisionID,1)
		
	    If CInt(strDivisionID) >= 7 and CInt(strDivisionID) =< 9 Then
		    is789DivisionSTS = True
	    Else
		    is789DivisionSTS = False
	    End If
	Else
	    is789DivisionSTS = False
	End If
End Function
'rmanyushin 56781 13.10.2009 End


'rmanyushin 119579 19.08.2010 Start
Function GetApproverForSTS_HolidayRequest(UserDepartment)
'{ph - 20120323
'	If GetSTSAssistantDirector(UserDepartment) = "" Then
'		GetApproverForSTS_HolidayRequest = STS_DirectorOfDirection
'	Else
'		GetApproverForSTS_HolidayRequest = STS_AssistantDirector
'	End If		
	If UCase(Request("l")) <> "RU" Then
		GetApproverForSTS_HolidayRequest = SIT_DirectorOfInitiatorsDepartment
	Else

	If GetSTSAssistantDirector(UserDepartment) = "" Then
			GetApproverForSTS_HolidayRequest = STS_DirectorOfDirection
	Else
			GetApproverForSTS_HolidayRequest = STS_AssistantDirector
				End If
	End If		
	'ph - 20120323}

End Function
'rmanyushin 119579 19.08.2010 End


'Запрос №11 - СТС - start
'Возвращает язык (аналогично параметру языка) на основании списка БЕ. Определение идет по первому символу кода БЕ: 1 - рус., 2 - чеш., в остальных - англ.
'Если в списке несколько БЕ с разными первыми символами или ничего нет, возвращается "?"
Function GetLangByBusinessUnits(parBusinessUnitsList)
  Dim sBusinessUnitsList, sBUClass, sBUClass2, arBUs, i

  GetLangByBusinessUnits = "?"

  sBusinessUnitsList =  Trim(MyCStr(parBusinessUnitsList))
  If sBusinessUnitsList = "" Then
    Exit Function
  End If

  If InStr(sBusinessUnitsList, VbCrLf) = 0 Then 'Только одна БЕ
    sBUClass = Left(sBusinessUnitsList, 1)
  Else
    arBUs = Split(sBusinessUnitsList, VbCrLf)
    sBUClass = Left(arBUs(0), 1)
    For i = 1 To UBound(arBUs)
      If arBUs(i) <> "" Then
        sBUClass2 = Left(arBUs(i), 1)
        If sBUClass <> sBUClass2 Then 'В списке есть БЕ разных классов
          Exit Function
        End If
      End If
    Next
  End If

  Select Case sBUClass
    Case "1" GetLangByBusinessUnits = "RU"
    Case "2" GetLangByBusinessUnits = "3"
    Case Else GetLangByBusinessUnits = ""
  End Select
End Function

'Запрос №34 - СТС - start
'Получить список ролей и соответствующих пользователей (список возвращаемого формата используется функцией ReplaceRolesInList)
'(!) В целях экономии ресурсов предполагается, что на всех языках система работает с одной БД, берется ConnectStringRUS
'Также предполагается, что на разных языках названия ролей различны
'Запрос №34 - СТС - Добавлены новые роли, для их поддержки дополнительно передаются менеджер проекта и заказчик переработки
'Вместо менеджера проекта в parProjectManager может быть передан номер проекта, тогда менеджер будет доставаться из БД
'Запрос №36 - СТС - start - Переименованы параметры, изменилось назначение
'В parCostCenter должно быть передано подразделение центра затрат, по нему считаются руководители центра затрат
'Руководители инициатора считаются по подразделению пользователя, переданного в parInitiator
'Запрос №46 - СТС - Добавлен очередной параметр для расшифровки роли STS_OvertimeFuncLeader - parOvertimeFuncLeaders
'Function GetFullRolesList(ByVal sDepartment, ByVal sUser, ByVal sBusinessUnit, ByVal parProjectManager, ByVal parOvertimeRequester)
'Function GetFullRolesList(ByVal parCostCenter, ByVal parInitiator, ByVal parBusinessUnit, ByVal parProjectManager, ByVal parOvertimeRequester)
Function GetFullRolesList(ByVal parCostCenter, ByVal parInitiator, ByVal parBusinessUnit, ByVal parProjectManager, ByVal parOvertimeRequester, ByVal parOvertimeFuncLeaders)
  Dim MyConn, dsTemp
  Dim sDirNameRU, sDirNameEN, sDirNameCZ
  Dim NewConnection, i
  Dim sRole, sUsers
  Dim sProjectManager, sProjectManagersDepartment, sOvertimeRequestersDepartment, sInitiatorID, sInitiatorsDepartment
  Dim sHeadOfInitiatorsUnit, sDirectorOfInitiatorsDepartment, sDirectorOfInitiatorsDivision
  Dim sHeadOfInitiatorsGroup, sAssistantDirector, sDirectorOfDirection
  Dim sDirectorOfProjectManagersDepartment, sDirectorOfOvertimeRequestersDepartment, sDirectorOfOvertimeRequestersDirection, sOvertimeRequestersAssistantDirector 'Остальные можно добавлять по мере необходимости
  Dim sCostCenterDirectorOfDepartment, sCostCenterDirectorOfDivision, sDirectorOfProjectManagersDivision, sProjectManagersNearestChief
  Dim sFullInitiator

  NewConnection = True
  If not IsNull(Conn) Then
    NewConnection = not IsObject(Conn)
  End If

  If NewConnection Then
    SetPayDoxPars 'Присваиваем переменные среды из Global.asa
    Set MyConn = Server.CreateObject("ADODB.Connection")
    MyConn.Open Application("ConnectStringRUS")
  Else
    Set MyConn = Conn
  End If
  
  If parCostCenter = "" Then 'Чтобы не сбилась привязка справочников когда не указан CostCenter
    ' DmGorsky_5: меняем местами проверку: вначале ищем SIT_SITRU ("СИТРОНИКС ИТ"), потом, если не найдено, - SIT_SITRONICS ("СИТРОНИКС")
    If InStr(Session("Department"), SIT_SITRU) = 1 Then ' DmGorsky
      sDirNameRU = SIT_RolesDirSITRU_RU ' DmGorsky
      sDirNameEN = SIT_RolesDirSITRU_EN ' DmGorsky
      sDirNameCZ = SIT_RolesDirSITRU_CZ ' DmGorsky
    ElseIf InStr(Session("Department"), SIT_STS) = 1 Then
      sDirNameRU = SIT_RolesDirSTS_RU
      sDirNameEN = SIT_RolesDirSTS_EN
      sDirNameCZ = SIT_RolesDirSTS_CZ
    ElseIf InStr(Session("Department"), SIT_SITRONICS) = 1 Then
      sDirNameRU = SIT_RolesDirSitronics_RU
      sDirNameEN = SIT_RolesDirSitronics_EN
      sDirNameCZ = SIT_RolesDirSitronics_CZ
    ElseIf InStr(Session("Department"), SIT_RTI) = 1 Then
      sDirNameRU = SIT_RolesDirRTI
      sDirNameEN = SIT_RolesDirRTI
      sDirNameCZ = SIT_RolesDirRTI
    ElseIf InStr(Session("Department"), SIT_MIKRON) = 1 Then
      sDirNameRU = SIT_RolesDirMIKRON
      sDirNameEN = SIT_RolesDirMIKRON
      sDirNameCZ = SIT_RolesDirMIKRON
    Else
      sDirNameRU = SIT_RolesDirSitronics_RU
      sDirNameEN = SIT_RolesDirSitronics_EN
      sDirNameCZ = SIT_RolesDirSitronics_CZ
    End If
  Else
    If InStr(parCostCenter, SIT_SITRU) = 1 Then ' DmGorsky_5
      sDirNameRU = SIT_RolesDirSITRU_RU ' DmGorsky_5
      sDirNameEN = SIT_RolesDirSITRU_RU ' DmGorsky_5
      sDirNameCZ = SIT_RolesDirSITRU_RU ' DmGorsky_5
    ElseIf InStr(parCostCenter, SIT_SITRONICS) = 1 Then ' DmGorsky_5
      sDirNameRU = SIT_RolesDirSitronics_RU
      sDirNameEN = SIT_RolesDirSitronics_EN
      sDirNameCZ = SIT_RolesDirSitronics_CZ
    ElseIf InStr(parCostCenter, SIT_STS) = 1 Then
      sDirNameRU = SIT_RolesDirSTS_RU
      sDirNameEN = SIT_RolesDirSTS_EN
      sDirNameCZ = SIT_RolesDirSTS_CZ
    ElseIf InStr(parCostCenter, SIT_RTI) = 1 Then
      sDirNameRU = SIT_RolesDirRTI
      sDirNameEN = SIT_RolesDirRTI
      sDirNameCZ = SIT_RolesDirRTI
    ElseIf InStr(parCostCenter, SIT_MIKRON) = 1 Then
      sDirNameRU = SIT_RolesDirMIKRON
      sDirNameEN = SIT_RolesDirMIKRON
      sDirNameCZ = SIT_RolesDirMIKRON
    Else

    End If
  End If

  GetFullRolesList = ""

  'определяем менеджера проекта
  sProjectManager = Trim(MyCStr(parProjectManager))
  If sProjectManager <> "" Then
     If GetUserID(sProjectManager) = "" Then 'Вместо менеджера передан номер проекта, имя нужно достать из БД
        Set dsTemp = Server.CreateObject("ADODB.Recordset")
        sSQL = "select * from ProjectList where ProjectID = " & sUnicodeSymbol & "'" & sProjectManager & "'"
        dsTemp.Open sSQL, MyConn, 3, 1, &H1
        If not dsTemp.EOF Then
           sProjectManager = Trim(MyCStr(dsTemp("ProjectManagerUser")))
        Else
           sProjectManager = ""
        End If
        dsTemp.Close
     End If
  End If

'Запрос №36 - СТС - start
  sInitiatorID = GetUserID(MyCStr(parInitiator))
  If sInitiatorID = "" Then
     sInitiatorID = MyCStr(parInitiator)
  End If
  'определяем департамент инициатора, чтобы потом рассчитывать его руководителей
  sInitiatorsDepartment = ""
  If sInitiatorID <> "" Then
     oPayDox.GetUserDetails sInitiatorID, sName, sPhone, sEMail, sICQ, sInitiatorsDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
	 'И полное имя для вставки в роль "#Инициатор";
	 sFullInitiator = InsertionName(sName, sInitiatorID)
  End If
'Запрос №36 - СТС - end

  'определяем департамент менеджера проекта, чтобы потом рассчитывать его руководителей
  sProjectManagersDepartment = ""
  If GetUserID(sProjectManager) <> "" Then
     oPayDox.GetUserDetails GetUserID(sProjectManager), sName, sPhone, sEMail, sICQ, sProjectManagersDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
  End If

  'определяем департамент заказчика переработки, чтобы потом рассчитывать его руководителей
  sOvertimeRequestersDepartment = ""
  If GetUserID(MyCStr(parOvertimeRequester)) <> "" Then
     oPayDox.GetUserDetails GetUserID(MyCStr(parOvertimeRequester)), sName, sPhone, sEMail, sICQ, sOvertimeRequestersDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
  End If
AddLogD "sProjectManagersDepartment: " & sProjectManagersDepartment
AddLogD "sOvertimeRequestersDepartment: " & sOvertimeRequestersDepartment

  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  sSQL = "Select UserDirValues.Field1,UserDirValues.Field2 from UserDirValues inner join UserDirectories on (UserDirValues.UDKeyField = UserDirectories.KeyField) "
  sSQL = sSQL & " where UserDirectories.Name = "&sUnicodeSymbol&"'"&sDirNameRU&"' or UserDirectories.Name = "&sUnicodeSymbol&"'"&sDirNameEN&"' or UserDirectories.Name = "&sUnicodeSymbol&"'"&sDirNameCZ&"'"
AddLogD "GetFullRolesList - SQL: "+sSQL

  dsTemp.Open sSQL, MyConn, 3, 1, &H1
  If not dsTemp.EOF Then
     i = 0
     sHeadOfInitiatorsUnit = ""
     sDirectorOfInitiatorsDepartment = ""
     sDirectorOfInitiatorsDivision = ""
     Do While not dsTemp.EOF
        sRole = Trim(MyCStr(dsTemp("Field1")))
        Select case sRole
           case SIT_HeadOfInitiatorsUnit_RU,SIT_HeadOfInitiatorsUnit_EN,SIT_HeadOfInitiatorsUnit_CZ,RTI_HeadOfInitiatorsUnit,MIKRON_HeadOfInitiatorsUnit
                If sInitiatorsDepartment = "" Then
                   sUsers = ""
                Else
                   If sHeadOfInitiatorsUnit = "" Then
                      sHeadOfInitiatorsUnit = GetChiefOfDepUpperByLevel(sInitiatorsDepartment, 3, sInitiatorID, parBusinessUnit)
                   End If
                   sUsers = sHeadOfInitiatorsUnit
                End If
           case SIT_DirectorOfInitiatorsDepartment_RU, SIT_DirectorOfInitiatorsDepartment_EN, SIT_DirectorOfInitiatorsDepartment_CZ
                If sInitiatorsDepartment = "" Then
                   sUsers = ""
                Else
                   If sDirectorOfInitiatorsDepartment = "" Then
                      sDirectorOfInitiatorsDepartment = GetChiefOfDepUpperByLevel(sInitiatorsDepartment, 2, sInitiatorID, parBusinessUnit)
                   End If
                   sUsers = sDirectorOfInitiatorsDepartment
                End If
           case SIT_DirectorOfInitiatorsDivision_RU, SIT_DirectorOfInitiatorsDivision_EN, SIT_DirectorOfInitiatorsDivision_CZ, SIT_VicePresidentOfInitiator_RU, SIT_VicePresidentOfInitiator_EN, SIT_VicePresidentOfInitiator_CZ
                If sInitiatorsDepartment = "" Then
                   sUsers = ""
                Else
                   If sDirectorOfInitiatorsDivision = "" Then
                      sDirectorOfInitiatorsDivision = GetChiefOfDepUpperByLevel(sInitiatorsDepartment, 1, sInitiatorID, parBusinessUnit)
                   End If
                   sUsers = sDirectorOfInitiatorsDivision
                End If
'Запрос №34 - СТС - новые роли - start
           case STS_HeadOfInitiatorsGroup_RU, STS_HeadOfInitiatorsGroup_EN, STS_HeadOfInitiatorsGroup_CZ
                If sInitiatorsDepartment = "" Then
                   sUsers = ""
                Else
                   If sHeadOfInitiatorsGroup = "" Then
                      sHeadOfInitiatorsGroup = GetChiefOfDepUpperByLevel(sInitiatorsDepartment, 4, sInitiatorID, parBusinessUnit)
                   End If
                   sUsers = sHeadOfInitiatorsGroup
                End If
           case STS_AssistantDirector_RU, STS_AssistantDirector_EN, STS_AssistantDirector_CZ
                If sAssistantDirector = "" Then
                   sAssistantDirector = GetInsertionNameInCurrentLanguage(GetSTSAssistantDirector(parCostCenter))
                End If
                sUsers = sAssistantDirector
           case STS_DirectorOfDirection_RU, STS_DirectorOfDirection_EN, STS_DirectorOfDirection_CZ
                If sDirectorOfDirection = "" Then
                   sDirectorOfDirection = GetInsertionNameInCurrentLanguage(GetSTSDirectorOfDirection(parCostCenter))
                End If
                sUsers = sDirectorOfDirection
           case STS_ProjectManager_RU, STS_ProjectManager_EN, STS_ProjectManager_CZ
                sUsers = sProjectManager
           case STS_Overtime_Requester_RU, STS_Overtime_Requester_EN, STS_Overtime_Requester_CZ
                sUsers = MyCStr(parOvertimeRequester)
'Запрос №46 - СТС - start
           case STS_OvertimeFuncLeader_RU, STS_OvertimeFuncLeader_EN, STS_OvertimeFuncLeader_CZ
                sUsers = MyCStr(parOvertimeFuncLeaders)
           case STS_Initiator_RU, STS_Initiator_EN, STS_Initiator_CZ
                sUsers = MyCStr(sFullInitiator)
'Запрос №46 - СТС - end
           case STS_DirectorOfProjectManagersDepartment_RU, STS_DirectorOfProjectManagersDepartment_EN, STS_DirectorOfProjectManagersDepartment_CZ
                If sProjectManagersDepartment = "" Then
                   sUsers = ""
                Else
                   If sDirectorOfProjectManagersDepartment = "" Then
                      sDirectorOfProjectManagersDepartment = GetChiefOfDepUpperByLevel(sProjectManagersDepartment, 2, sProjectManager, parBusinessUnit)
                   End If
                   sUsers = sDirectorOfProjectManagersDepartment
                End If
           case STS_DirectorOfProjectManagersDivision_RU, STS_DirectorOfProjectManagersDivision_EN, STS_DirectorOfProjectManagersDivision_CZ
                If sProjectManagersDepartment = "" Then
                   sUsers = ""
                Else
                   If sDirectorOfProjectManagersDivision = "" Then
                      sDirectorOfProjectManagersDivision = GetChiefOfDepUpperByLevel(sProjectManagersDepartment, 1, sProjectManager, parBusinessUnit)
                   End If
                   sUsers = sDirectorOfProjectManagersDivision
                End If
        '{ph - 20120607
           case STS_ProjectManagersNearestChief_RU, STS_ProjectManagersNearestChief_EN, STS_ProjectManagersNearestChief_CZ
                If sProjectManagersDepartment = "" Then
                   sUsers = ""
                Else
                   If sProjectManagersNearestChief = "" Then
                      sProjectManagersNearestChief = GetNearestChief(sProjectManagersDepartment, GetUserID(sProjectManager), parBusinessUnit)
                   End If
                   sUsers = sProjectManagersNearestChief
                End If
        'ph - 20120607}
           case STS_DirectorOfOvertimeRequestersDepartment_RU, STS_DirectorOfOvertimeRequestersDepartment_EN, STS_DirectorOfOvertimeRequestersDepartment_CZ
                If sOvertimeRequestersDepartment = "" Then
                   sUsers = ""
                Else
                   If sDirectorOfOvertimeRequestersDepartment = "" Then
                      sDirectorOfOvertimeRequestersDepartment = GetChiefOfDepUpperByLevel(sOvertimeRequestersDepartment, 2, parOvertimeRequester, parBusinessUnit)
                   End If
                   sUsers = sDirectorOfOvertimeRequestersDepartment
                End If
           case STS_DirectorOfOvertimeRequestersDirection_RU, STS_DirectorOfOvertimeRequestersDirection_EN, STS_DirectorOfOvertimeRequestersDirection_CZ
                If sOvertimeRequestersDepartment = "" Then
                   sUsers = ""
                Else
                   If sDirectorOfOvertimeRequestersDirection = "" Then
                      sDirectorOfOvertimeRequestersDirection = GetInsertionNameInCurrentLanguage(GetSTSDirectorOfDirection(sOvertimeRequestersDepartment))
                   End If
                   sUsers = sDirectorOfOvertimeRequestersDirection
                End If
           case STS_OvertimeRequestersAssistantDirector_RU, STS_OvertimeRequestersAssistantDirector_EN, STS_OvertimeRequestersAssistantDirector_CZ
                If sOvertimeRequestersDepartment = "" Then
                   sUsers = ""
                Else
                   If sOvertimeRequestersAssistantDirector = "" Then
                      sOvertimeRequestersAssistantDirector = GetInsertionNameInCurrentLanguage(GetSTSAssistantDirector(sOvertimeRequestersDepartment))
                   End If
                   sUsers = sOvertimeRequestersAssistantDirector
                End If
'Запрос №34 - СТС - новые роли - end
'Запрос №36 - СТС - новые роли - start
           case STS_CostCenterDirectorOfDepartment_RU, STS_CostCenterDirectorOfDepartment_EN, STS_CostCenterDirectorOfDepartment_CZ
                If parCostCenter = "" Then
                   sUsers = ""
                Else
                   If sCostCenterDirectorOfDepartment = "" Then
                      sCostCenterDirectorOfDepartment = GetChiefOfDepUpperByLevel(parCostCenter, 2, sInitiatorID, parBusinessUnit)
                   End If
                   sUsers = sCostCenterDirectorOfDepartment
                End If
           case STS_CostCenterDirectorOfDivision_RU, STS_CostCenterDirectorOfDivision_EN, STS_CostCenterDirectorOfDivision_CZ
                If parCostCenter = "" Then
                   sUsers = ""
                Else
                   If sCostCenterDirectorOfDivision = "" Then
                      sCostCenterDirectorOfDivision = GetChiefOfDepUpperByLevel(parCostCenter, 1, sInitiatorID, parBusinessUnit)
                   End If
                   sUsers = sCostCenterDirectorOfDivision
                End If
'Запрос №36 - СТС - новые роли - end
           case else
                sUsers = Trim(MyCStr(dsTemp("Field2")))
        End Select
AddLogD "GetFullRolesList - "+CStr(i)+"  Role: "+sRole+"  Value: "+sUsers
        If GetFullRolesList <> "" Then
           GetFullRolesList = GetFullRolesList + VbCrLf
        End If
        GetFullRolesList = GetFullRolesList + sRole + "|" + sUsers
        dsTemp.MoveNext
        i = i+1
     Loop
  End If
  dsTemp.Close

  If NewConnection Then
     MyConn.Close
  End If
End Function
'Запрос №34 - СТС - end

'Ph - 20101125 - НЕ ИСПОЛЬЗУЕТСЯ - start
'Получить шестизначный номер (для форматирования номера проекта в нумерации договоров)
'Function SixSymbol(parStr)
'  SixSymbol = Trim(MyCStr(parStr))
'  If Len(SixSymbol) > 6 Then
'    SixSymbol = Left(SixSymbol, 6)
'  ElseIf Len(SixSymbol) < 6 Then
'    SixSymbol =  String(6-Len(SixSymbol), "0") & SixSymbol
'  End If
'End Function
'Ph - 20101125 - end

'Получить код контрагента для нумерации договоров
'Используется Conn
Function GetPartnerCode(parPartnerName)
  Dim dsTemp, sSQL

  GetPartnerCode = ""
  sSQL = "select IDRequired from Partners where Name = " & sUnicodeSymbol &"'" & MakeSQLSafeSimple(parPartnerName) & "'"
  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  dsTemp.Open sSQL, Conn, 3, 1, &H1
  If not dsTemp.EOF Then
    GetPartnerCode = Trim(MyCStr(dsTemp("IDRequired")))
  End If
  dsTemp.Close
End Function

'Получить индекс документа для Договоров
Function GetNewDocIDForContracts(parClassDoc, parDepartment, parPartnerCode, parProjectNo, parPaymentDirection)
  Dim sSQL, sSuffix, sRightPart

  sSuffix = ""
  Select Case parPaymentDirection
    Case STS_ContractPaymentDirection_In_RU, STS_ContractPaymentDirection_In_EN, STS_ContractPaymentDirection_In_CZ
      sSuffix = "-R"
    Case STS_ContractPaymentDirection_Out_RU, STS_ContractPaymentDirection_Out_EN, STS_ContractPaymentDirection_Out_CZ
      sSuffix = "-C"
  End Select
  
  Select Case parDepartment
    Case SIT_SITRONICS 'нумерация для ситроникса
      sRightPart = "Right(DocID, Len(DocID)-CharIndex(N'/', DocID))"
	  'Проверка на скобки, чтобы отсечь поручения, где в скобках может быть номер любого формата, что приведет к ошибке в Cast
      sSQL="select IsNull(Max(Cast(Case When CharIndex(N'(', DocID) > 0 and CharIndex(N')', DocID) > CharIndex(N'(', DocID) Then 0 Else Case CharIndex(N'-', "&sRightPart&") When 0 Then "&sRightPart&" Else Left("&sRightPart&", CharIndex(N'-',"&sRightPart&")-1) End End as int)), 0) as MaxDocID from Docs where ClassDoc = N'"&parClassDoc&"' and DocID like N'%-"&parProjectNo&"/%'"
      'sSQL="select IsNull(Max(Cast(Case CharIndex(N'-', "&sRightPart&") When 0 Then "&sRightPart&" Else Left("&sRightPart&", CharIndex(N'-',"&sRightPart&")-1) End as int)), 0) as MaxDocID from Docs where ClassDoc = N'"&parClassDoc&"' and DocID like N'%-"&parProjectNo&"/%'"
    Case SIT_STS 'нумерация для СТС
      sRightPart = "Right(DocID, Len(DocID)-CharIndex(N'/', DocID))"
      sSQL="select IsNull(Max(Cast(Case When CharIndex(N'(', DocID) > 0 and CharIndex(N')', DocID) > CharIndex(N'(', DocID) Then 0 Else Case CharIndex(N'-', "&sRightPart&") When 0 Then "&sRightPart&" Else Left("&sRightPart&", CharIndex(N'-',"&sRightPart&")-1) End End as int)), 0) as MaxDocID from Docs where ClassDoc = N'"&parClassDoc&"' and DocID like N'%-"&parProjectNo&"/%'"
      'sSQL="select IsNull(Max(Cast(Case CharIndex(N'-', "&sRightPart&") When 0 Then "&sRightPart&" Else Left("&sRightPart&", CharIndex(N'-',"&sRightPart&")-1) End as int)), 0) as MaxDocID from Docs where ClassDoc = N'"&parClassDoc&"' and DocID like N'%-"&parProjectNo&"/%'"
    Case "OTHER B.U." 'для других бизнес направлений
  End Select
'  out sSQL 
  GetNewDocIDForContracts = GetNewDocID(sSQL, parPartnerCode&"-"&parProjectNo&"/", sSuffix, 3)
End Function

'Функция получения уникального номера по заданным правилам
'Нужно передать запрос получения из БД MaxDocID (максимальной инкрементируемой части номера)
Function GetNewDocID(parSQL, parPrefix, parSuffix, parDigits)
  Dim bNewConnection, sConnStr, MyConn, dsTemp, iNumberNext, bStop, sSQL

  bNewConnection = True
  If not IsNull(Conn) Then
    bNewConnection = not IsObject(Conn)
  End If
  If bNewConnection Then
    SetPayDoxPars 'Присваиваем переменные среды из Global.asa
    sConnStr = "ConnectString"
    Select Case UCase(Request("l"))
      Case "RU" sConnStr = sConnStr + "RUS"
      Case "3" sConnStr = sConnStr + "3"
    End Select
    Set MyConn = Server.CreateObject("ADODB.Connection")
    MyConn.Open Application(sConnStr)
  Else
    Set MyConn = Conn
  End If

  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  dsTemp.Open parSQL, MyConn, 3, 1, &H1
  If dsTemp.EOF Then
    iNumberNext = 0
  Else
    iNumberNext = dsTemp("MaxDocID")
  End If
  dsTemp.Close

  bStop = False
  Do While not bStop
    iNumberNext = iNumberNext+1
    GetNewDocID = parPrefix & MyLeadSymbolNVal(iNumberNext, "0", parDigits) & parSuffix

    sSQL = "select DocID from Docs where DocID = " & sUnicodeSymbol & "'" & GetNewDocID & "'"
    dsTemp.Open sSQL, MyConn, 3, 1, &H1
    If dsTemp.EOF Then
      bStop = True
    End If
    dsTemp.Close  
  Loop
    
  If bNewConnection Then
    MyConn.Close
  End If
End Function
'Запрос №11 - СТС - end

'Запрос №17 - СТС - start
Function NotificationWithReason(parAction, parReason)
  NotificationWithReason = "<br><b>" & parAction & " - " & HTMLEncode(GetFullName(SurnameGN(DelOtherLangFromFolder(Session("Name"))), Session("UserID"))) & iif(Trim(MyCStr(parReason)) = "", "", " - <font color = red>" & Replace(HTMLEncode(parReason), VbCrLf, "<br>") & "</font>") & "</b><br>"
End Function
'Запрос №17 - СТС - end


'rmanyushin 119579 19.08.2010 Start
Function GetSTSDirectorOfDirection(ByVal sDepartment)
  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  sSQL = "SELECT CostCenter, Name FROM STS_DirectorOfDirectionRU ORDER BY CostCenter"
AddLogD "GetSTSDirectorOfDirection SQL1: "+sSQL
  dsTemp.Open sSQL, Conn, 3, 1, &H1
AddLogD "GetSTSDirectorOfDirection sDepartment: " & sDepartment
  Do While not dsTemp.EOF
	AddLogD "GetSTSDirectorOfDirection CC: " & TRIM(dsTemp("CostCenter"))
	IF INSTR(sDepartment, TRIM(dsTemp("CostCenter"))) > 0 Then
		'GetSTSDirectorOfDirection = GetInsertionNameInCurrentLanguage(dsTemp("Name"))
		GetSTSDirectorOfDirection = dsTemp("Name")
		AddLogD "GetSTSDirectorOfDirection Name: " & dsTemp("Name")
    End If
  dsTemp.MoveNext
  Loop
  dsTemp.Close
End Function
'rmanyushin 119579 19.08.2010 End

'rmanyushin 119579 19.08.2010 Start
Function GetSTSAssistantDirector(ByVal sDepartment)
  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  sSQL = "SELECT CostCenter, Name FROM STS_AssistantDirectorRU ORDER BY CostCenter"
AddLogD "GetSTSAssistantDirector SQL1: "+sSQL
  dsTemp.Open sSQL, Conn, 3, 1, &H1
AddLogD "GetSTSAssistantDirector sDepartment: " & sDepartment
Do While not dsTemp.EOF
	AddLogD "GetSTSAssistantDirector CC: " & TRIM(dsTemp("CostCenter"))
	IF INSTR(sDepartment, TRIM(dsTemp("CostCenter"))) > 0 Then
		'GetSTSAssistantDirector = GetInsertionNameInCurrentLanguage(dsTemp("Name")) 
		GetSTSAssistantDirector = dsTemp("Name") 
		AddLogD "GetSTSAssistantDirector Name: " & dsTemp("Name")
    End If
  dsTemp.MoveNext
  Loop
  dsTemp.Close
End Function
'rmanyushin 119579 19.08.2010 End


'rmanyushin 136964 08.11.2010 Start
Function GetSTSSecurityManager(ByVal sDepartment)
	Set dsTemp = Server.CreateObject("ADODB.Recordset")
	sSQL = "SELECT * FROM UserDirValues Where UDKeyfield = '19' AND Field1 = N'" + STS_SecurityManager + "'"
	AddLogD "GetSTSSecurityManager SQL1: "+sSQL
	dsTemp.Open sSQL, Conn, 3, 1, &H1
	'AddLogD "GetSTSSecurityManager sDepartment: " & sDepartment
	Do While not dsTemp.EOF
		GetSTSSecurityManager = dsTemp("Field2") 
		AddLogD "GetSTSSecurityManager Name: " & dsTemp("Field2")
	dsTemp.MoveNext
	Loop
	dsTemp.Close
End Function

Function GetSTSHRDirector(ByVal sDepartment)
	Set dsTemp = Server.CreateObject("ADODB.Recordset")
	sSQL = "SELECT * FROM UserDirValues Where UDKeyfield = '19' AND Field1 = N'" + STS_HRDirector + "'"
	AddLogD "GetSTSHRDirector SQL1: "+sSQL
	dsTemp.Open sSQL, Conn, 3, 1, &H1
	'AddLogD "GetSTSHRDirector sDepartment: " & sDepartment
	Do While not dsTemp.EOF
		GetSTSHRDirector = dsTemp("Field2") 
		AddLogD "GetSTSHRDirector Name: " & dsTemp("Field2")
	dsTemp.MoveNext
	Loop
	dsTemp.Close
End Function
'rmanyushin 136964 08.11.2010 End


'rmanyushin 136151 13.10.2010 Start
Function is5DivisionSTS(strDepartmentName)
	Set objRegExp = CreateObject("VBScript.RegExp")
	objRegExp.Global = False
	objRegExp.Pattern = "\d{5}"
	Set objMatches = objRegExp.Execute(strDepartmentName)
    
    If objMatches.Count <> 0 Then
        strDivisionID = objMatches.Item(0).Value
		strDivisionID = Trim(strDivisionID)
        strDivisionID = Left(strDivisionID,1)
		
	    If CInt(strDivisionID) = 5 Then
		    is5DivisionSTS = True
	    Else
		    is5DivisionSTS = False
	    End If
	Else
	    is5DivisionSTS = False
	End If
End Function

'rmanyushin 158840 18.12.2010 Start
Function is3DivisionSTS(strDepartmentName)
	Set objRegExp = CreateObject("VBScript.RegExp")
	objRegExp.Global = False
	objRegExp.Pattern = "\d{5}"
	Set objMatches = objRegExp.Execute(strDepartmentName)
    
    If objMatches.Count <> 0 Then
        strDivisionID = objMatches.Item(0).Value
		strDivisionID = Trim(strDivisionID)
        strDivisionID = Left(strDivisionID,1)
		
	    If CInt(strDivisionID) = 3 Then
		    is3DivisionSTS = True
	    Else
		    is3DivisionSTS = False
	    End If
	Else
	    is3DivisionSTS = False
	End If
End Function
'rmanyushin 158840 18.12.2010 End

'Запрос №23 - СТС - start - Функции поддержки справочника правил
'Для справки структуры таблиц правил и лога их использования
'--Правила
' CREATE TABLE [dbo].[STS_Rules](
	' [OrderNo] [nvarchar](16) NOT NULL,
	' [Active] [nvarchar](1) NULL,
	' [ClassDoc] [nvarchar](255) NULL,
	' [AmountMin] [money] NULL,
	' [AmountMax] [money] NULL,
	' [ChartOfAccounts] [nvarchar](255) NULL,
	' [CostCenter] [nvarchar](255) NULL,
	' [ProjectCode] [nvarchar](255) NULL,
	' [BusinessUnit] [nvarchar](255) NULL,
	' [PartnerName] [nvarchar](255) NULL,
	' [ListToReconcile] [nvarchar](1024) NULL,
	' [NameAproval] [nvarchar](128) NULL,
	' [Comment] [nvarchar](1024) NULL,
 ' CONSTRAINT [PK_STS_Rules] PRIMARY KEY CLUSTERED 
' (
	' [OrderNo] ASC
' )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
' ) ON [PRIMARY]

'--Лог
' CREATE TABLE [dbo].[STS_RulesLog](
	' [GUID] [uniqueidentifier] NOT NULL,
	' [DocID] [nvarchar](128) NULL,
	' [DateCreation] [datetime] NULL,
	' [RuleOrderNo] [nvarchar](16) NULL,
	' [Amount] [money] NULL,
	' [ChartOfAccounts] [nvarchar](1024) NULL,
	' [CostCenter] [nvarchar](1024) NULL,
	' [ProjectCode] [nvarchar](1024) NULL,
	' [BusinessUnit] [nvarchar](64) NULL,
	' [PartnerName] [nvarchar](255) NULL,
 ' CONSTRAINT [PK_STS_RulesLog] PRIMARY KEY CLUSTERED 
' (
	' [GUID] ASC
' )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
' ) ON [PRIMARY]

MASK_SYMBOL = "?" 'Одиночный символ в маске
MASK_SUBSTRING = "*" 'Подстрока в маске
NOT_CONTEXT = "#NOT#" 'Ключевое слово, показывающее, что нужно взять обратный результат сравнения (если удовлетворяет маске, то отсечь)
EMPTY_CONTEXT = "#EMPTY#" 'Ключевое слово, показывающее, что параметр должен быть пустой

'Основная функция для вызова извне. Принимает параметры документа, возвращает согласующих и утверждающего (parListToReconcile, parNameAproval) по справочнику правил
'Sub GetReconciliationByRules(parDocID, parClassDoc, parAmount, parChartOfAccounts, parCostCenter, parProjectCode, parBusinessUnit, parPartnerName, ByRef parListToReconcile, ByRef parNameAproval)
'  Dim sSQL
'  Dim NewConnection, MyConn, dsTemp, dsLog
'  Dim bMatch
'  Dim dCurrentDate
'
'  sSQL = "Select * from STS_Rules where ClassDoc = N'" & parClassDoc & "' and "
'  If MyCStr(parAmount) <> "" and IsNumeric(parAmount) Then
'    sSQL = sSQL & "(IsNull(AmountMin, 0) = 0 or IsNull(AmountMin, 0) <> 0 and AmountMin <= " & MyCStr(parAmount) & ") and (IsNull(AmountMax, 0) = 0 or IsNull(AmountMax, 0) <> 0 and AmountMax >= " & MyCStr(parAmount) & ") and "
'  End If
'  If MyCStr(parPartnerName) <> "" Then
'    sSQL = sSQL & " PartnerName like " & sUnicodeSymbol & "'%" & MyCStr(parPartnerName) & "%' and "
'  End If
'
'  If LCase(Trim(Right(sSQL, 4))) = "and" Then
'    sSQL = Left(sSQL, Len(sSQL)-4)
'  End If
'  sSQL = sSQL & " and Active = 'Y' order by OrderNo"
'AddLogD "GetReconciliationByRules - SQL: "+sSQL
'
'  NewConnection = True
'  If not IsNull(Conn) Then
'    NewConnection = not IsObject(Conn)
'  End If
'
'  If NewConnection Then
'    SetPayDoxPars 'Присваиваем переменные среды из Global.asa
'    Set MyConn = Server.CreateObject("ADODB.Connection")
'    MyConn.Open Application("ConnectStringRUS")
'  Else
'    Set MyConn = Conn
'  End If
'
'  Set dsTemp = Server.CreateObject("ADODB.Recordset")
'  dsTemp.CursorLocation = 3
'  dsTemp.Open sSQL, MyConn, 3, 1, &H1
'  dsTemp.ActiveConnection = Nothing
'AddLogD "Rules in recordset: " & CStr(dsTemp.RecordCount)
'
'  'Таблица для лога
'  Set dsLog = Server.CreateObject("ADODB.Recordset")
'  dsLog.Open "select top 1 * from STS_RulesLog", MyConn, 1, 2, &H1
'  dCurrentDate = Now
'
'  parListToReconcile = ""
'  parNameAproval = ""
'  Do While not dsTemp.EOF
'AddLogD "-------------- Rule OrderNo: " & MyCStr(dsTemp("OrderNo"))
'    bMatch = True
'AddLogD "parChartOfAccounts: """ & MyCStr(parChartOfAccounts) & """"
'AddLogD "dsTemp(ChartOfAccounts): """ & MyCStr(dsTemp("ChartOfAccounts")) & """"
'    If MyCStr(parChartOfAccounts) <> "" Then
'      If MyCStr(dsTemp("ChartOfAccounts")) <> "" Then
'	    If InStr(MyCStr(dsTemp("ChartOfAccounts")), NOT_CONTEXT) > 0 Then
'          bMatch = not CheckCode(MyCStr(parChartOfAccounts), Replace(MyCStr(dsTemp("ChartOfAccounts")), NOT_CONTEXT, ""))
'		Else
'          bMatch = CheckCode(MyCStr(parChartOfAccounts), MyCStr(dsTemp("ChartOfAccounts")))
'		End If
'      End If
'    End If
'AddLogD "bMatch: " & CStr(bMatch)
'    If bMatch Then
'AddLogD "parCostCenter: """ & MyCStr(parCostCenter) & """"
'AddLogD "dsTemp(CostCenter): """ & MyCStr(dsTemp("CostCenter")) & """"
'      If MyCStr(parCostCenter) <> "" Then
'        If MyCStr(dsTemp("CostCenter")) <> "" Then
'	      If InStr(MyCStr(dsTemp("CostCenter")), NOT_CONTEXT) > 0 Then
'            bMatch = not CheckCode(MyCStr(parCostCenter), Replace(MyCStr(dsTemp("CostCenter")), NOT_CONTEXT, ""))
'		  Else
'            bMatch = CheckCode(MyCStr(parCostCenter), MyCStr(dsTemp("CostCenter")))
'		  End If
'        End If
'      End If
'AddLogD "bMatch: " & CStr(bMatch)
'    End If
'    If bMatch Then
'AddLogD "parProjectCode: """ & MyCStr(parProjectCode) & """"
'AddLogD "dsTemp(ProjectCode): """ & MyCStr(dsTemp("ProjectCode")) & """"
'      If MyCStr(parProjectCode) <> "" Then
'        If MyCStr(dsTemp("ProjectCode")) <> "" Then
'	      If InStr(MyCStr(dsTemp("ProjectCode")), NOT_CONTEXT) > 0 Then
'            bMatch = not CheckCode(MyCStr(parProjectCode), Replace(MyCStr(dsTemp("ProjectCode")), NOT_CONTEXT, ""))
'		  Else
'            bMatch = CheckCode(MyCStr(parProjectCode), MyCStr(dsTemp("ProjectCode")))
'		  End If
'        End If
'      End If
'AddLogD "bMatch: " & CStr(bMatch)
'    End If
'    If bMatch Then
'AddLogD "parBusinessUnit: """ & MyCStr(parBusinessUnit) & """"
'AddLogD "dsTemp(BusinessUnit): """ & MyCStr(dsTemp("BusinessUnit")) & """"
'      If MyCStr(parBusinessUnit) <> "" Then
'        If MyCStr(dsTemp("BusinessUnit")) <> "" Then
'	      If InStr(MyCStr(dsTemp("BusinessUnit")), NOT_CONTEXT) > 0 Then
'            bMatch = not CheckCode(MyCStr(parBusinessUnit), Replace(MyCStr(dsTemp("BusinessUnit")), NOT_CONTEXT, ""))
'		  Else
'            bMatch = CheckCode(MyCStr(parBusinessUnit), MyCStr(dsTemp("BusinessUnit")))
'		  End If
'        End If
'      End If
'AddLogD "bMatch: " & CStr(bMatch)
'    End If
'    If bMatch Then
'	  'Пишем в лог, что правило задействовано
'    dsLog.AddNew
'    dsLog("GUID") = oPayDox.GetGUID1()
'    dsLog("DocID") = parDocID
'    dsLog("DateCreation") = dCurrentDate
'    dsLog("RuleOrderNo") = MyCStr(dsTemp("OrderNo"))
'    dsLog("Amount") = parAmount
'    dsLog("ChartOfAccounts") = MyCStr(parChartOfAccounts)
'    dsLog("CostCenter") = MyCStr(parCostCenter)
'    dsLog("ProjectCode") = MyCStr(parProjectCode)
'    dsLog("BusinessUnit") = MyCStr(parBusinessUnit)
'    dsLog("PartnerName") = MyCStr(parPartnerName)
'    dsLog.Update
'
'AddLogD "Initial ListToReconcile: " & MyCStr(parListToReconcile)
'AddLogD "Adding ListToReconcile: " & MyCStr(dsTemp("ListToReconcile"))
'      If MyCStr(dsTemp("ListToReconcile")) <> "" Then
'        parListToReconcile = parListToReconcile & MyCStr(dsTemp("ListToReconcile")) & " "
'      End If
'AddLogD "Result ListToReconcile1: " & MyCStr(parListToReconcile)
'AddLogD "Initial NameAproval: " & MyCStr(parNameAproval)
'AddLogD "Adding NameAproval: " & MyCStr(dsTemp("NameAproval"))
'      If MyCStr(dsTemp("NameAproval")) <> "" Then
'        'Если уже есть утверждающий, то ставим его последовательно в согласующих, т.е. приоритет правил растет с ростом порядкового номера
'        If Trim(parNameAproval) <> "" Then
'          parListToReconcile = parListToReconcile & VbCrLf & parNameAproval & VbCrLf
'        End If
'        parNameAproval = MyCStr(dsTemp("NameAproval"))
'      End If
'AddLogD "Result NameAproval: " & MyCStr(parNameAproval)
'AddLogD "Result ListToReconcile2: " & MyCStr(parListToReconcile)
'    End If
'    dsTemp.MoveNext
'  Loop
'  dsLog.Close
'  dsTemp.Close
'
'  If NewConnection Then
'    MyConn.Close
'  End If
'
'  'Окончательная ListToReconcile обработка, чтобы убрать лишние концы строк и пробелы
'  parListToReconcile = CleanListToReconcile(parListToReconcile)
'End Sub

'Чистка списка согласования от пустых строк
Function CleanListToReconcile(parStr)
  Dim LenCrLf

  LenCrLf = Len(VbCrLf)
  CleanListToReconcile = Trim(parStr)
  Do While InStr(CleanListToReconcile, VbCrLf & " ") > 0
    CleanListToReconcile = Replace(CleanListToReconcile, VbCrLf & " ", VbCrLf)
  Loop
  Do While InStr(CleanListToReconcile, " " & VbCrLf) > 0
    CleanListToReconcile = Replace(CleanListToReconcile, " " & VbCrLf, VbCrLf)
  Loop
  Do While InStr(CleanListToReconcile, VbCrLf & VbCrLf) > 0
    CleanListToReconcile = Replace(CleanListToReconcile, VbCrLf & VbCrLf, VbCrLf)
  Loop
  Do While Left(CleanListToReconcile, LenCrLf) = VbCrLf
    CleanListToReconcile = Right(CleanListToReconcile, Len(CleanListToReconcile)-LenCrLf)
    CleanListToReconcile = Trim(CleanListToReconcile)
  Loop
  Do While Right(CleanListToReconcile, LenCrLf) = VbCrLf
    CleanListToReconcile = Left(CleanListToReconcile, Len(CleanListToReconcile)-LenCrLf)
    CleanListToReconcile = Trim(CleanListToReconcile)
  Loop
End Function

'Проверить удовлетворяет ли код значениям из справочника
Function CheckCode(parDocCode, parDictCodes)
  Dim arDictCodes, i, iMax

AddLogD "CheckCode(""" & parDocCode & """, """ & parDictCodes & """)"
  If InStr(parDictCodes, VbCrLf) > 0 Then
    arDictCodes = Split(parDictCodes, VbCrLf)
  Else
    arDictCodes = Split(parDictCodes, ",")
  End If
  iMax = UBound(arDictCodes)

  CheckCode = False
  For i = 0 To iMax
    CheckCode = CompareByMask(parDocCode, Trim(arDictCodes(i)))
AddLogD CStr(i+1) & "/" & CStr(iMax+1) & " - CompareByMask(""" & parDocCode & """, """ & Trim(arDictCodes(i)) & """) = " & CStr(CheckCode)
    If CheckCode Then
      Exit Function
    End If
  Next
End Function

'Поддерживаются MASK_SUBSTRING только слева и справа, не в середине маски
Function CompareByMask(parCode, parMask)
  Dim iPos

  If Replace(parMask, MASK_SUBSTRING, "") = "" Then
    CompareByMask = True
    Exit Function
  End If
  If Len(parCode) < Len(parMask) Then
    CompareByMask = False
    Exit Function
  End If

  If Left(parMask, 1) = MASK_SUBSTRING Then
    If Right(parMask, 1) = MASK_SUBSTRING Then
      CompareByMask = False
      iPos = 1
      Do While iPos < Len(parCode)-Len(parMask)+2
        CompareByMask = CompareByMask(Mid(parCode, iPos, Len(parMask)-3+iPos), Mid(parMask, 2, Len(parMask)-2))
        If CompareByMask Then
          Exit Function
        End If
        iPos = iPos+1
      Loop
      Exit Function
    Else
      CompareByMask = CompareByMask(Right(parCode, Len(parMask)-1), Right(parMask, Len(parMask)-1))
      Exit Function
    End If
  ElseIf Right(parMask, 1) = MASK_SUBSTRING Then
    CompareByMask = CompareByMask(Left(parCode, Len(parMask)-1), Left(parMask, Len(parMask)-1))
    Exit Function
  End If

  If Len(parCode) <> Len(parMask) Then
    CompareByMask = False
    Exit Function
  End If

  iPos = 1
  CompareByMask = True
  Do While iPos <= Len(parCode)
    CompareByMask = CompareByMask and (Mid(parMask, iPos, 1) = MASK_SYMBOL or Mid(parMask, iPos, 1) = Mid(parCode, iPos, 1))
    If not CompareByMask Then
      Exit Function
    End If
    iPos = iPos+1
  Loop
End Function


'Получить список вообще всех ролей (из языковых справочников плюс из справочника ролей для 
'заявок СТС) и соответствующих пользователей (список возвращаемого формата используется функцией
'ReplaceRolesInList)
'(!) В целях экономии ресурсов предполагается, что на всех языках система работает с одной БД,
'берется ConnectStringRUS
'Также предполагается, что на разных языках названия ролей различны
'Запрос №46 - СТС - start
'Function GetAllRolesList(ByVal parCostCenter, ByVal parInitiator, ByVal parBusinessUnit, ByVal parProjectManager, ByVal parOvertimeRequester)
Function GetAllRolesList(ByVal parCostCenter, ByVal parInitiator, ByVal parBusinessUnit, ByVal parProjectManager, ByVal parOvertimeRequester, ByVal parOvertimeFuncLeaders)
'Запрос №46 - СТС - end
  Dim MyConn, dsTemp
  Dim NewConnection
  Dim sRole

AddLogD "GetAllRolesList START"
  NewConnection = True
  If not IsNull(Conn) Then
    NewConnection = not IsObject(Conn)
  End If

  If NewConnection Then
    SetPayDoxPars 'Присваиваем переменные среды из Global.asa
    Set MyConn = Server.CreateObject("ADODB.Connection")
    MyConn.Open Application("ConnectStringRUS")
  Else
    Set MyConn = Conn
  End If

'Запрос №34 - СТС - start
'  GetAllRolesList = GetFullRolesList(sDepartment, sUser, sBusinessUnit)
'Запрос №46 - СТС - start
'  GetAllRolesList = GetFullRolesList(parCostCenter, parInitiator, parBusinessUnit, parProjectManager, parOvertimeRequester)
  GetAllRolesList = GetFullRolesList(parCostCenter, parInitiator, parBusinessUnit, parProjectManager, parOvertimeRequester, parOvertimeFuncLeaders)
'Запрос №46 - СТС - end
'Запрос №34 - СТС - end

  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  sSQL = "Select Role, Users, Case When CharIndex(N'9999', BusinessUnit) = 1 Then '0000' Else BusinessUnit End as BusinessUnit from RolesForOrders_STS where CharIndex(N'" & GetCodeFromCode_NameString(parBusinessUnit) & "', BusinessUnit)= 1 or CharIndex(N'" & Left(parBusinessUnit, 1) & "000', BusinessUnit)= 1 or CharIndex(N'9999', BusinessUnit) = 1 order by Role, BusinessUnit desc"
AddLogD "GetAllRolesList - SQL: "+sSQL

  dsTemp.Open sSQL, MyConn, 3, 1, &H1
  If not dsTemp.EOF Then
    i = 0
    sRole = ""
    Do While not dsTemp.EOF
      If sRole <> Trim(MyCStr(dsTemp("Role"))) Then
        sRole = Trim(MyCStr(dsTemp("Role")))
AddLogD "GetAllRolesList - "+CStr(i)+"  Role: "+sRole+"  Value: "+Trim(MyCStr(dsTemp("Users")))
        If GetAllRolesList <> "" Then
          GetAllRolesList = GetAllRolesList + VbCrLf
        End If
        GetAllRolesList = GetAllRolesList + sRole + "|" + Trim(MyCStr(dsTemp("Users")))
        i = i+1
      End If
      dsTemp.MoveNext
    Loop
  End If
  dsTemp.Close

  If NewConnection Then
    MyConn.Close
  End If
AddLogD "GetAllRolesList END"
'AddLogD "GetAllRolesList: " & GetAllRolesList
End Function
'Запрос №23 - СТС - end

'ph - 20101108 - start
Function ShowCurrencyRate(parSumUSD, parSum)
  If CCur(parSum) = 0 or CCur(parSumUSD) = 0 Then
    ShowCurrencyRate = "0 (0)"
  Else
    ShowCurrencyRate = CStr(CCur(parSumUSD/parSum)) & "  (" & CStr(CCur(parSum/parSumUSD)) & ")"
  End If
End Function
'ph - 20101108 - end

'Запрос №30 - СТС - start
'Получить список пользователей, которым нужно показать кнопку Переназначить
Function GetUsersToShowClickReSetResponsible()
  Dim MyConn, dsTemp
  Dim NewConnection
  Dim sSQL

AddLogD "GetUsersToShowClickReSetResponsible START"
  NewConnection = True
  If not IsNull(Conn) Then
    NewConnection = not IsObject(Conn)
  End If

  If NewConnection Then
    SetPayDoxPars 'Присваиваем переменные среды из Global.asa
    Set MyConn = Server.CreateObject("ADODB.Connection")
    MyConn.Open Application("ConnectStringRUS")
  Else
    Set MyConn = Conn
  End If

  Set dsTemp = Server.CreateObject("ADODB.Recordset")
'Запрос №36 - СТС - start
'  sSQL = "Select Users from RolesForOrders_STS where Role = " & sUnicodeSymbol & "'" & STS_Purchase_Logistics_Department & "'"
  sSQL = "Select Users from RolesForOrders_STS where Role = " & sUnicodeSymbol & "'" & STS_UsersToShowResetResponsibleButton & "'"
'Запрос №36 - СТС - end
AddLogD "GetUsersToShowClickReSetResponsible - SQL: "+sSQL

  dsTemp.CursorLocation = 3
  dsTemp.Open sSQL, MyConn, 3, 1, &H1
  dsTemp.ActiveConnection = Nothing

  GetUsersToShowClickReSetResponsible = ""
  Do While not dsTemp.EOF
    GetUsersToShowClickReSetResponsible = GetUsersToShowClickReSetResponsible & Trim(MyCStr(dsTemp("Users")))
    dsTemp.MoveNext
  Loop
  dsTemp.Close
AddLogD "GetUsersToShowClickReSetResponsible: " & GetUsersToShowClickReSetResponsible

  If NewConnection Then
    MyConn.Close
  End If
AddLogD "GetUsersToShowClickReSetResponsible END"
End Function
'Запрос №30 - СТС - end

'Запрос №31 - СТС - start
'Получить код дочерней компании по бизнес единице (для вставки в индекс документа)
'parFieldToReturn - возвращаемое поле LetterCode (для писем) или OrderCode (для приказов)
'Алгоритм сравнения как в справочнике правил, поиск идет до первого совпадения
Function GetDZKCode(parBusinessUnit, parFieldToReturn)
  Dim NewConnection, MyConn, dsTemp
  Dim bMatch

  GetDZKCode = ""
  If MyCStr(parBusinessUnit) = "" Then
    Exit Function
  End If

  NewConnection = True
  If not IsNull(Conn) Then
    NewConnection = not IsObject(Conn)
  End If

  If NewConnection Then
    SetPayDoxPars 'Присваиваем переменные среды из Global.asa
    Set MyConn = Server.CreateObject("ADODB.Connection")
    MyConn.Open Application("ConnectStringRUS")
  Else
    Set MyConn = Conn
  End If

  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  dsTemp.CursorLocation = 3
  dsTemp.Open "Select * from STS_DZKCodes", MyConn, 3, 1, &H1
  dsTemp.ActiveConnection = Nothing

AddLogD "-------------- STS_DZKCodes --------------"
  Do While not dsTemp.EOF
AddLogD "parBusinessUnit: """ & MyCStr(parBusinessUnit) & """"
AddLogD "dsTemp(BusinessUnit): """ & MyCStr(dsTemp("BusinessUnit")) & """"
    If InStr(MyCStr(dsTemp("BusinessUnit")), NOT_CONTEXT) > 0 Then
      bMatch = not CheckCode(MyCStr(parBusinessUnit), Replace(MyCStr(dsTemp("BusinessUnit")), NOT_CONTEXT, ""))
    Else
      bMatch = CheckCode(MyCStr(parBusinessUnit), MyCStr(dsTemp("BusinessUnit")))
    End If
AddLogD "bMatch: " & CStr(bMatch)
    If bMatch Then
      GetDZKCode = MyCStr(dsTemp(parFieldToReturn))
AddLogD "GetDZKCode: " & GetDZKCode
      Exit Do
    End If
    dsTemp.MoveNext
  Loop
  dsTemp.Close

  If NewConnection Then
    MyConn.Close
  End If
End Function

'----------------------- Вспомогательные функции для расчета индексов документов - START
'--------------------- Функция заменяющая в SQL-выражении отрицательные значения на нулевые
Function NotNegative_SQL(par)
  NotNegative_SQL = "Case When " & par & " < 0 Then 0 Else " & par & " End"
End Function

'Функция преобразующая SQL-выражение в число
Function IsNumeric_SQL(par)
  IsNumeric_SQL = "Case IsNumeric(" & par & ") When 1 Then Cast(" & par & " as int) Else 0 End"
End Function

'Функция возвращающая запрос для поиска максимальной инкрементируемой части индекса
'parIncrementedPart - SQL-выражение для вырезания числовой инкрементируемой части номера
'parClassDoc - категория документа
'parSearchCol - поле для поиска номера (DocID или DocIDAdd)
'parPrefix - префикс искомого номера
'parDepartment - бизнес-направление, может быть пустым, если нумерация сквозная по всем предприятиям
Function GenerateSQLForMaxDocIDSearch(parIncrementedPart, parClassDoc, parSearchCol, parPrefix, parDepartment)
  GenerateSQLForMaxDocIDSearch = "Select IsNull(Max(" & IsNumeric_SQL(parIncrementedPart) & "), 0) as MaxDocID from Docs where "
  GenerateSQLForMaxDocIDSearch = GenerateSQLForMaxDocIDSearch & " ClassDoc = N'" & parClassDoc & "' "
  GenerateSQLForMaxDocIDSearch = GenerateSQLForMaxDocIDSearch & " And CharIndex(N'" & parPrefix & "', " & parSearchCol & ") = 1 "
  If MyCStr(parDepartment) <> "" Then
     GenerateSQLForMaxDocIDSearch = GenerateSQLForMaxDocIDSearch & " And CharIndex(N'" & parDepartment & "', Department) = 1 "
  End If
End Function

'Выше определена еще функция
'Function GetNewDocID(parSQL, parPrefix, parSuffix, parDigits)
' ----------------------------- Вспомогательные функции для расчета индексов документов - END

'Расчет номеров входящих документов
Function GetNewDocIDForVhodyashie(parClassDoc, parDepartment, parSubClassParameter, parSubClassParameter2,  parParentDocID, parIsProjectDocID, parDZKCode)
Dim iDigits, sSearchCol, sSQL, sPrefix, sSeparator, sSubClassParameter, sIncrementedPart

  If parIsProjectDocID = "PJ-" Then
     sSearchCol = "DocIDadd"
  Else
     sSearchCol = "DocID"
  End If

  Select Case parDepartment
    Case SIT_SITRONICS
        'нумерация для ситроникса
	    sSeparator = "-"
	    sPrefix="IN"+Right(CStr(Year(Date)),2)+"/"+Left(parSubClassParameter, NotNegative(InStr(parSubClassParameter, " ")-1))+sSeparator
	    sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+parIsProjectDocID+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+parClassDoc+"' AND "+sSearchCol+" like N'"+parIsProjectDocID+sPrefix+"%' AND Department like N'"+parDepartment+"%'"
        iDigits = 3
    case SIT_RTI
        'нумерация для РТИ
        sSeparator = "/"
        select case Trim(Request("UserFieldText8"))
        case "ОАО ""РТИ"""
          sPrefix="IN"+Right(CStr(Year(Date)),2)+"_RTI/"
        case "ОАО ""Концерн РТИ"""
          sPrefix="IN"+Right(CStr(Year(Date)),2)+"_CRS/"
        case "ГК ОАО ""РТИ"""
          sPrefix="IN"+Right(CStr(Year(Date)),2)+"_RTI.GK/"
        end select 

	    'sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+parIsProjectDocID+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+parClassDoc+"' AND "+sSearchCol+" like N'"+parIsProjectDocID+sPrefix+"%' AND Department like N'"+parDepartment+"%'"
        sSQL="select ISNULL(Max(cast(right(replace("+sSearchCol+",'-',''),len(replace("+sSearchCol+",'-',''))-charindex('"+sSeparator+"',replace("+sSearchCol+",'-',''),len('"+parIsProjectDocID+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+parClassDoc+"' AND "+sSearchCol+" like N'"+parIsProjectDocID+sPrefix+"%' AND Department like N'"+parDepartment+"%'"
        'out sSql
        iDigits = 6
    case SIT_MINC
        'нумерация входящих для ОАО РТИ
        sSeparator = "/"
        sPrefix="Bx."+Right(CStr(Year(Date)),2)+sSeparator
        sSQL="select ISNULL(Max(cast(right(replace("+sSearchCol+",'-',''),len(replace("+sSearchCol+",'-',''))-charindex('"+sSeparator+"',replace("+sSearchCol+",'-',''),len('"+parIsProjectDocID+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+parClassDoc+"' AND "+sSearchCol+" like N'"+parIsProjectDocID+sPrefix+"%' AND Department like N'"+parDepartment+"%'"
        iDigits = 6
    case SIT_VTSS
        'нумерация входящих для ВТСС
        sSeparator = "/"
        sPrefix="IN"+Right(CStr(Year(Date)),2)+"_VTSS/"
	    'sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+parIsProjectDocID+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+parClassDoc+"' AND "+sSearchCol+" like N'"+parIsProjectDocID+sPrefix+"%' AND Department like N'"+parDepartment+"%'"
        sSQL="select ISNULL(Max(cast(right(replace("+sSearchCol+",'-',''),len(replace("+sSearchCol+",'-',''))-charindex('"+sSeparator+"',replace("+sSearchCol+",'-',''),len('"+parIsProjectDocID+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+parClassDoc+"' AND "+sSearchCol+" like N'"+parIsProjectDocID+sPrefix+"%' AND Department like N'"+parDepartment+"%'"
        'out sSql
        iDigits = 6
    Case SIT_STS
        'нумерация для СТС
        If parDZKCode = "" Then
           sSubClassParameter = Left(parSubClassParameter, NotNegative(InStr(parSubClassParameter, " ")-1)) 'Старая нумерация с подклассами
        Else
           sSubClassParameter = parDZKCode
        End If
	    sSeparator = "/"
	    sPrefix="IN" & Right(CStr(Year(Date)), 2) & "_" & sSubClassParameter & sSeparator
        sIncrementedPart = "Right(" & sSearchCol & ", " & NotNegative_SQL("CharIndex(Reverse(N'" & sSeparator & "'), Reverse(" & sSearchCol & "))-1") & ")"
        sSQL = GenerateSQLForMaxDocIDSearch(sIncrementedPart, parClassDoc, sSearchCol, parIsProjectDocID & sPrefix, parDepartment)
        'If parIsProjectDocID = "" Then
        '   sNotProjectPart = sSearchCol
        'Else
        '   sNotProjectPart = "Case When CharIndex(N'" & parIsProjectDocID & "', " & sSearchCol & ") = 1 Then Right(" & sSearchCol & ", " & Len(parIsProjectDocID) & ") Else " & sSearchCol & " End"
        'End If
	    'sSQL = "Select IsNull(Max(Cast(Right(" & sNotProjectPart & ", Len(" & sNotProjectPart & ")-CharIndex(N'" & sSeparator & "', " & sNotProjectPart & ")+" & Len(sSeparator) & "-1) as int)), 0) as MaxDocID from Docs where ClassDoc=N'"+parClassDoc+"' AND "+sSearchCol+" like N'"+parIsProjectDocID+sPrefix+"%' AND Department like N'"+parDepartment+"%'"
'sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+parIsProjectDocID+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+parClassDoc+"' AND "+sSearchCol+" like N'"+parIsProjectDocID+sPrefix+"%' AND Department like N'"+parDepartment+"%'"
      iDigits = 3
        Case SIT_MIKRON
      sSeparator = "-"
      sPrefix="MS-IN"+Right(CStr(Year(Date)),2)+sSeparator
      sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+parIsProjectDocID+"')+4))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+parClassDoc+"' AND "+sSearchCol+" like N'"+parIsProjectDocID+sPrefix+"%' AND Department like N'"+parDepartment+"%'"
      iDigits = 4
'Запрос №1 - СИБ - start
    Case SIT_SIB
        sSeparator = "/"
        sPrefix="IN"+Right(CStr(Year(Date)),2)+"_SIB"+sSeparator
        sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+parIsProjectDocID+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+parClassDoc+"' AND "+sSearchCol+" like N'"+parIsProjectDocID+sPrefix+"%' AND Department like N'"+parDepartment+"%'"
        iDigits = 3
'Запрос №1 - СИБ - end
    Case "OTHER B.U."
        'для других бизнес направлений
	    sPrefix = "N/A"
    Case SIT_SITRU ' DmGorsky
        'Нумерация для Ситроникс ИТ
	    sSeparator = "-"
	    sPrefix="IN"+Right(CStr(Year(Date)),2)+"_SITRU"+sSeparator
	    sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+parIsProjectDocID+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+parClassDoc+"' AND "+sSearchCol+" like N'"+parIsProjectDocID+sPrefix+"%' AND Department like N'"+parDepartment+"%'"
        'sSQL="select ISNULL(Max(cast(right(replace("+sSearchCol+",'-',''),len(replace("+sSearchCol+",'-',''))-charindex('"+sSeparator+"',replace("+sSearchCol+",'-',''),len('"+parIsProjectDocID+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+parClassDoc+"' AND "+sSearchCol+" like N'"+parIsProjectDocID+sPrefix+"%' AND Department like N'"+parDepartment+"%'"
        iDigits = 6
   
  End Select

  GetNewDocIDForVhodyashie = GetNewDocID(sSQL, parIsProjectDocID & sPrefix, "", iDigits)
AddLogD "GetNewDocIDForVhodyashie - sSQL: " & sSQL
End Function

'Расчет номеров исходящих документов
Function GetNewDocIDForIshodyashie(parClassDoc, parDepartment, parSubClassParameter, parSubClassParameter2,  parParentDocID, parIsProjectDocID, parDZKCode)
  Dim iDigits, sSearchCol, sSQL, sPrefix, sSuffix, sSeparator, sSubClassParameter, sIncrementedPart

  If parIsProjectDocID="PJ-" Then
    sSearchCol="DocIDadd"
  Else
    sSearchCol="DocID"
  End If

  Select Case parDepartment
    Case SIT_SITRONICS
      'нумерация для ситроникса
      sPrefix="OUT"+Right(CStr(Year(Date)),2)+"/"
      If parParentDocID="" Then
        sSuffix="5-"
      else
      End If 
      sSeparator="/"
      sSQL="select ISNULL(Max(cast(case CharIndex('5-', "+sSearchCol+") when 0 then right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+parIsProjectDocID+"')+1)-2) else right("+sSearchCol+",len("+sSearchCol+")-charindex('-',"+sSearchCol+",len('"+parIsProjectDocID+"')+1)) end as int)),0) as MaxDocID from Docs where ClassDoc=N'"+parClassDoc+"' AND "+sSearchCol+" like N'"+parIsProjectDocID+sPrefix+"%' AND Department like N'"+parDepartment+"%'"
      iDigits = 3
      
'kkoshkin begin       
    Case SIT_RTI
      'нумерация для РТИ
        select case Trim(Request("UserFieldText8"))
        case "ОАО ""РТИ"""
          sPrefix="OUT"+Right(CStr(Year(Date)),2)+"_RTI/"
        case "ОАО ""Концерн РТИ"""
          sPrefix="OUT"+Right(CStr(Year(Date)),2)+"_CRS/"
        case "ГК ОАО ""РТИ"""
          sPrefix="OUT"+Right(CStr(Year(Date)),2)+"_RTI.GK/"
        end select 
        
      if Trim(Request("UserFieldText8")) = "" Then
         sPrefix = parSubClassParameter
      End If
      sSeparator="/"
      sSQL="select ISNULL(Max(cast(case CharIndex('5-', "+sSearchCol+") when 0 then right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+parIsProjectDocID+"')+1)-2) else right("+sSearchCol+",len("+sSearchCol+")-charindex('-',"+sSearchCol+",len('"+parIsProjectDocID+"')+1)) end as int)),0) as MaxDocID from Docs where ClassDoc=N'"+parClassDoc+"' AND "+sSearchCol+" like N'"+parIsProjectDocID+sPrefix+"%' AND Department like N'"+parDepartment+"%'"
      iDigits = 6
   
'kkoshkin end   
'kkoshkin begin       
    Case SIT_MINC
      'нумерация для OAO РТИ
      sPrefix="Ucx."+Right(CStr(Year(Date)),2)+"/"      
      sSeparator="/"
      sSQL="select ISNULL(Max(cast(case CharIndex('5-', "+sSearchCol+") when 0 then right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+parIsProjectDocID+"')+1)-2) else right("+sSearchCol+",len("+sSearchCol+")-charindex('-',"+sSearchCol+",len('"+parIsProjectDocID+"')+1)) end as int)),0) as MaxDocID from Docs where ClassDoc=N'"+parClassDoc+"' AND "+sSearchCol+" like N'"+parIsProjectDocID+sPrefix+"%' AND Department like N'"+parDepartment+"%'"
      iDigits = 6
   
'kkoshkin end   
'kkoshkin begin       
    Case SIT_VTSS
      'нумерация для ВТСС
      sPrefix="OUT"+Right(CStr(Year(Date)),2)+"_VTSS/"     
      sSeparator="/"
      sSQL="select ISNULL(Max(cast(case CharIndex('5-', "+sSearchCol+") when 0 then right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+parIsProjectDocID+"')+1)-2) else right("+sSearchCol+",len("+sSearchCol+")-charindex('-',"+sSearchCol+",len('"+parIsProjectDocID+"')+1)) end as int)),0) as MaxDocID from Docs where ClassDoc=N'"+parClassDoc+"' AND "+sSearchCol+" like N'"+parIsProjectDocID+sPrefix+"%' AND Department like N'"+parDepartment+"%'"
      iDigits = 6
   
'kkoshkin end   
    Case SIT_STS
      'нумерация для СТС
      If parDZKCode = "" Then
         sSubClassParameter = Left(parSubClassParameter2, NotNegative(InStr(parSubClassParameter2, " ")-1)) 'Старая нумерация с подклассами
      Else
         sSubClassParameter = parDZKCode
      End If
      sSuffix = ""
      sSeparator = "/"
      sPrefix = "OUT" & Right(CStr(Year(Date)), 2) & "_" & sSubClassParameter & sSeparator
      sIncrementedPart = "Right(" & sSearchCol & ", " & NotNegative_SQL("CharIndex(Reverse(N'" & sSeparator & "'), Reverse(" & sSearchCol & "))-1") & ")"
      sSQL = GenerateSQLForMaxDocIDSearch(sIncrementedPart, parClassDoc, sSearchCol, parIsProjectDocID & sPrefix, parDepartment)
'      sSQL = "select ISNULL(Max(case IsNumeric(right("+sSearchCol+",charindex('"+sSeparator+"',reverse("+sSearchCol+"),0)-1)) when 1 then cast(right("+sSearchCol+",charindex('"+sSeparator+"',reverse("+sSearchCol+"),0)-1)as int) else 0 end),0) as MaxDocID from Docs where ClassDoc=N'"+parClassDoc+"' AND "+sSearchCol+" like N'"+parIsProjectDocID+sPrefix+"%' AND Department like N'"+parDepartment+"%'"
      iDigits = 3
'amw Mikron
    Case SIT_MIKRON
      sSuffix = ""
      sPrefix="MS-OUT"+Right(CStr(Year(Date)),2)+"-"+ parSubClassParameter +"/"
      sSQL="select ISNULL(Max(cast(case CharIndex('5-', "+sSearchCol+") when 0 then Replace(right("+sSearchCol+",len("+sSearchCol+")-charindex('-',"+sSearchCol+",len('"+parIsProjectDocID+"')+4)),'"+parSubClassParameter+"' + '/','') else right("+sSearchCol+",len("+sSearchCol+")-charindex('-',"+sSearchCol+",len('"+parIsProjectDocID+"')+1)) end as int)),0) as MaxDocID from Docs where ClassDoc=N'"+parClassDoc+"' AND "+sSearchCol+" like N'"+parIsProjectDocID+sPrefix+"%' AND Department like N'"+parDepartment+"%'"
      iDigits = 4
'vnik_scsp    
    Case SIT_SCSP
      'нумерация для СПСТ
      sPrefix="SCSP-OUT"+Right(CStr(Year(Date)),2)+"/"
      If parParentDocID="" Then
        sSuffix="5-"
      else
      End If 
      sSeparator="/"
      sSQL="select ISNULL(Max(cast(case CharIndex('5-', "+sSearchCol+") when 0 then right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+parIsProjectDocID+"')+1)-2) else right("+sSearchCol+",len("+sSearchCol+")-charindex('5-',"+sSearchCol+",len('"+parIsProjectDocID+"')+1)-1) end as int)),0) as MaxDocID from Docs where ClassDoc=N'"+parClassDoc+"' AND "+sSearchCol+" like N'"+parIsProjectDocID+sPrefix+"%' AND Department like N'"+parDepartment+"%'"
      iDigits = 3
     'vnik_scsp
'Запрос №1 - СИБ - start
    Case SIT_SIB
      sSeparator = "/"
      sPrefix="OUT"+Right(CStr(Year(Date)),2)+"_SIB"+sSeparator
      sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+parIsProjectDocID+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+parClassDoc+"' AND "+sSearchCol+" like N'"+parIsProjectDocID+sPrefix+"%' AND Department like N'"+parDepartment+"%'"
      iDigits = 3
'Запрос №1 - СИБ - end
    Case "OTHER B.U."
      'для других бизнес направлений
    Case SIT_SITRU ' DmGorsky
      'Нумерация для Ситроникс ИТ
      sPrefix="OUT"+Right(CStr(Year(Date)),2)+"_SITRU/"
      If parParentDocID="" Then
        sSuffix="5-"
      else
        sSuffix=""
      End If 
      sSeparator="/"
      sSQL="select ISNULL(Max(cast(case CharIndex('5-', "+sSearchCol+") when 0 then right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+parIsProjectDocID+"')+1)-2) else right("+sSearchCol+",len("+sSearchCol+")-charindex('-',"+sSearchCol+",len('"+parIsProjectDocID+"')+1)) end as int)),0) as MaxDocID from Docs where ClassDoc=N'"+parClassDoc+"' AND "+sSearchCol+" like N'"+parIsProjectDocID+sPrefix+"%' AND Department like N'"+parDepartment+"%'"
      iDigits = 3
  End Select

  GetNewDocIDForIshodyashie = GetNewDocID(sSQL, parIsProjectDocID & sPrefix & sSuffix, "", iDigits)
End Function

Function GetNewDocIDForRaspDocs(parClassDoc, parDepartment, parSubClassParameter, parSubClassParameter2,  parParentDocID, parIsProjectDocID, parDZKCode)
Dim iDigits, sSearchCol, sSQL, sPrefix, sSuffix, sSeparator, sDZKCode, sIncrementedPart

  If parIsProjectDocID="PJ-" Then
     sSearchCol="DocIDadd"
  Else
     sSearchCol="DocID"
  End If

  Select Case parDepartment
    Case SIT_SITRONICS
    'нумерация для ситроникса
      sSeparator = "/"
      sPrefix=Left(parSubClassParameter, NotNegative(InStr(parSubClassParameter, " ")-1))+"-"+Right(CStr(Year(Date)),2)+sSeparator
      sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+parIsProjectDocID+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+parClassDoc+"' AND "+sSearchCol+" like N'"+parIsProjectDocID+sPrefix+"%' AND Department like N'"+parDepartment+"%'"
      iDigits = 3
    Case SIT_RTI
    'нумерация для РТИ ОАО ""РТИ""
      sSeparator = "/"
      select case Trim(Request("UserFieldText1"))
        case "OR - приказ ОАО ""Концерн РТИ"""
          sPrefix=Left(parSubClassParameter, NotNegative(InStr(parSubClassParameter, " ")-1))+"-"+Right(CStr(Year(Date)),2)+"_CRS"+sSeparator
        case "OR - приказ ОАО ""РТИ"""
          sPrefix=Left(parSubClassParameter, NotNegative(InStr(parSubClassParameter, " ")-1))+"-"+Right(CStr(Year(Date)),2)+"_RTI"+sSeparator
        case "Р - распоряжение"
          sPrefix=Left(parSubClassParameter, NotNegative(InStr(parSubClassParameter, " ")-1))+"-"+Right(CStr(Year(Date)),2)+"_RTI"+sSeparator
        case "Р - распоряжение ГК"
          sPrefix=Left(parSubClassParameter, NotNegative(InStr(parSubClassParameter, " ")-1))+"-"+Right(CStr(Year(Date)),2)+"_RTI.GK"+sSeparator
      end select              
      if Trim(Request("UserFieldText1")) = "" Then
      sPrefix = parSubClassParameter
      End If
      
          'sPrefix=Left(parSubClassParameter, NotNegative(InStr(parSubClassParameter, " ")-1))+"-"+Right(CStr(Year(Date)),2)+"_RTI"+sSeparator
          sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+parIsProjectDocID+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+parClassDoc+"' AND "+sSearchCol+" like N'"+parIsProjectDocID+sPrefix+"%' AND Department like N'"+parDepartment+"%'"
          iDigits = 3
     Case SIT_STS
          'нумерация для СТС
          If parDZKCode <> "" Then
             sDZKCode = "-" & parDZKCode
          Else
             sDZKCode = ""
          End If
          sSeparator = "-"
          sPrefix = Left(parSubClassParameter, NotNegative(InStr(parSubClassParameter, " ")-1)) & sDZKCode & "_" & Right(CStr(Year(Date)), 2) & sSeparator
          sIncrementedPart = "Right(" & sSearchCol & ", " & NotNegative_SQL("CharIndex(Reverse(N'" & sSeparator & "'), Reverse(" & sSearchCol & "))-1") & ")"
          sSQL = GenerateSQLForMaxDocIDSearch(sIncrementedPart, parClassDoc, sSearchCol, parIsProjectDocID & sPrefix, parDepartment)
'          sSQL="select ISNULL(Max(cast(right("+sSearchCol+",charindex('"+sSeparator+"',reverse("+sSearchCol+"),0)-1)as int)),0) as MaxDocID from Docs where ClassDoc=N'"+parClassDoc+"' AND "+sSearchCol+" like N'"+parIsProjectDocID+sPrefix+"%' AND Department like N'"+parDepartment+"%'"
          iDigits = 3
     Case SIT_MIKRON
          sSeparator = "/"
          sPrefix=Left(parSubClassParameter, NotNegative(InStr(parSubClassParameter, " ")-1))+"-"+Right(CStr(Year(Date)),2)+sSeparator
          sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+parIsProjectDocID+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+parClassDoc+"' AND "+sSearchCol+" like N'"+parIsProjectDocID+sPrefix+"%' AND Department like N'"+parDepartment+"%'"
          iDigits = 3
'Запрос №1 - СИБ - start
     Case SIT_SIB
          sSeparator = "-"
          sPrefix=Left(parSubClassParameter, NotNegative(InStr(parSubClassParameter, " ")-1))+Right(CStr(Year(Date)),2)+"_SIB"+sSeparator
          sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+parIsProjectDocID+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+parClassDoc+"' AND "+sSearchCol+" like N'"+parIsProjectDocID+sPrefix+"%' AND Department like N'"+parDepartment+"%'"
          iDigits = 3
'Запрос №1 - СИБ - end
     Case "OTHER B.U."
'для других бизнес направлений
     Case SIT_SITRU ' DmGorsky
'Нумерация для Ситроникс ИТ
          sSeparator = "/"
          sPrefix=Left(parSubClassParameter, NotNegative(InStr(parSubClassParameter, " ")-1))+"-"+Right(CStr(Year(Date)),2)+"_SITRU"+sSeparator
          sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+parIsProjectDocID+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+parClassDoc+"' AND "+sSearchCol+" like N'"+parIsProjectDocID+sPrefix+"%' AND Department like N'"+parDepartment+"%'"
          iDigits = 3
  End Select

   GetNewDocIDForRaspDocs = GetNewDocID(sSQL, parIsProjectDocID & sPrefix, "", iDigits)
End Function

'vnik_protocols
'vnik_protocolsCPC
'Function GetNewDocIDForProtocols(parClassDoc, parDepartment, parSubClassParameter, parSubClassParameter2,  parParentDocID, parIsProjectDocID, parDZKCode)
'    Dim iDigits, sSearchCol, sSQL, sPrefix, sSuffix, sSeparator, sDZKCode, sIncrementedPart
'    
'  If parIsProjectDocID="PJ-" Then
'      sSearchCol="DocIDadd"
'  Else
'      sSearchCol="DocID"
'  End If
'
'  Select Case parDepartment
'    Case SIT_SITRONICS
'    'нумерация для ситроникса
'      sPrefix= Trim(parSubClassParameter2 + "" + parSubClassParameter)
'	  sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-len('"+sPrefix+"'))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sPrefix+"%' AND Department like N'"+sDepartment+"%'"
'	  iDigits = 6
'  Case "OTHER B.U."
'    'для других бизнес направлений
'  End Select
'GetNewDocIDForProtocols = GetNewDocID(sSQL, parIsProjectDocID & sPrefix, "", iDigits)
'End Function
'vnik_protocolsCPC
'vnik_protocols

'vnik_purchase_order
Function GetNewDocIDForPurchaseOrder(parClassDoc, parDepartment, parSubClassParameter, parSubClassParameter2,  parParentDocID, parIsProjectDocID, parDZKCode)
Dim iDigits, sSearchCol, sSQL, sPrefix, sSuffix, sSeparator, sDZKCode, sIncrementedPart

  If parIsProjectDocID="PJ-" Then
     sSearchCol="DocIDadd"
  Else
     sSearchCol="DocID"
  End If

  Select Case parDepartment
    Case SIT_SITRONICS
    'нумерация для ситроникса
      iDigits = 6
      sSeparator = "/"
      sPrefix=parSubClassParameter+"-"+Right(CStr(Year(Date)),2)+sSeparator
      sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-len('"+sPrefix+"'))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sPrefix+"%' AND Department like N'"+sDepartment+"%'"
    Case SIT_RTI
      iDigits = 6
      sSeparator = "/"
      sPrefix=parSubClassParameter+Right(CStr(Year(Date)),2)+sSeparator
      sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-len('"+sPrefix+"'))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sPrefix+"%' AND Department like N'"+sDepartment+"%'"
      addlogd "888999" + Trim(sSQL)
    Case SIT_MIKRON
        'нумерация для Mikron
'amw 03/09/2013 изменена нумерация на CODE-MMYY/NNNN
        iDigits = 4
        sSeparator = "/"
        sPrefix=parSubClassParameter + LeadSymbolNVal(CStr(Month(Date)),"0",2) + Right(CStr(Year(Date)),2) + sSeparator
        sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-len('"+sPrefix+"'))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sPrefix+"%' AND Department like N'"+sDepartment+"%'"
     Case "OTHER B.U."
        'для других бизнес направлений
  End Select
  
  GetNewDocIDForPurchaseOrder = GetNewDocID(sSQL, parIsProjectDocID & sPrefix, "", iDigits)

End Function
'vnik_purchase_order

'vnik_payment_order
Function GetNewDocIDForPaymentOrder(parClassDoc, parDepartment, parSubClassParameter, parSubClassParameter2,  parParentDocID, parIsProjectDocID, parDZKCode)
Dim iDigits, sSearchCol, sSQL, sPrefix, sSuffix, sSeparator, sDZKCode, sIncrementedPart

  If parIsProjectDocID="PJ-" Then
    sSearchCol="DocIDadd"
  Else
    sSearchCol="DocID"
  End If

  Select Case parDepartment
    Case SIT_RTI
      sSeparator = "/"
      sPrefix=parSubClassParameter+Right(CStr(Year(Date)),2)+sSeparator
      sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-len('"+sPrefix+"'))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sPrefix+"%' AND Department like N'"+sDepartment+"%'"
      addlogd "exor777: " + sSQL      
      iDigits = 6
    Case SIT_MIKRON
'amw 03/09/2013 изменена нумерация на CODE-MMYY/NNNN
      sSeparator = "/"
      sPrefix=parSubClassParameter + LeadSymbolNVal(CStr(Month(Date)),"0",2) + Right(CStr(Year(Date)),2) + sSeparator
      sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-len('"+sPrefix+"'))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sPrefix+"%' AND Department like N'"+sDepartment+"%'"
      iDigits = 4
  End Select
GetNewDocIDForPaymentOrder = GetNewDocID(sSQL, parIsProjectDocID & sPrefix, "", iDigits)
End Function
'vnik_payment_order


'vnik_contracts
Function GetNewDocIDForContractsMC(parClassDoc, parDepartment, parSubClassParameter, parSubClassParameter2,  parParentDocID, parIsProjectDocID, parDZKCode)
    Dim iDigits, sSearchCol, sSQL, sPrefix, sSuffix, sSeparator, sDZKCode, sIncrementedPart

  If parIsProjectDocID="PJ-" Then
    sSearchCol="DocIDadd"
  Else
    sSearchCol="DocID"
  End If

  Select Case parDepartment
    Case SIT_SITRONICS
    'нумерация для ситроникса
      sSeparator = "/"
      sPrefix=parSubClassParameter+"-"+Right(CStr(Year(Date)),2)+sSeparator
      sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-len('"+sPrefix+"'))as int)),0) as MaxDocID from Docs where ClassDoc=N'"+sClassDoc+"' AND "+sSearchCol+" like N'"+sPrefix+"%' AND Department like N'"+sDepartment+"%'"
      iDigits = 6
  Case "OTHER B.U."
    'для других бизнес направлений
  End Select
GetNewDocIDForContractsMC = GetNewDocID(sSQL, parIsProjectDocID & sPrefix, "", iDigits)
End Function
'vnik_contracts

'Получить согласующих и регистраторов по БЕ
Sub GetDZKVisors(ByVal parBusinessUnit, ByVal parClassDoc, ByRef parListToReconcile, ByRef parRegistrar)
  Dim NewConnection, MyConn, dsTemp
  Dim sRegistrarField, sListToReconcileField
  Dim bMatch

  parListToReconcile = ""
  parRegistrar = ""
  If MyCStr(parBusinessUnit) = "" Then
    Exit Sub
  End If
  If InStr(UCase(parClassDoc), UCase(SIT_VHODYASCHIE)) = 1 Then
    sRegistrarField = "RegistrarIn"
    sListToReconcileField = ""
  ElseIf InStr(UCase(parClassDoc), UCase(SIT_ISHODYASCHIE)) = 1 Then
    sRegistrarField = "RegistrarOut"
    sListToReconcileField = "ListToReconcileOut"
  ElseIf InStr(UCase(parClassDoc), UCase(SIT_RASP_DOCS)) = 1 Then
    sRegistrarField = "RegistrarOrder"
    sListToReconcileField = "ListToReconcileOrder"
  Else
    Exit Sub
  End If

  NewConnection = True
  If not IsNull(Conn) Then
    NewConnection = not IsObject(Conn)
  End If

  If NewConnection Then
    SetPayDoxPars 'Присваиваем переменные среды из Global.asa
    Set MyConn = Server.CreateObject("ADODB.Connection")
    MyConn.Open Application("ConnectStringRUS")
  Else
    Set MyConn = Conn
  End If

  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  dsTemp.CursorLocation = 3
  dsTemp.Open "Select * from STS_DZKVisors", MyConn, 3, 1, &H1
  dsTemp.ActiveConnection = Nothing

AddLogD "-------------- STS_DZKVisors --------------"
  Do While not dsTemp.EOF
AddLogD "parBusinessUnit: """ & MyCStr(parBusinessUnit) & """"
AddLogD "dsTemp(BusinessUnit): """ & MyCStr(dsTemp("BusinessUnit")) & """"
    If InStr(MyCStr(dsTemp("BusinessUnit")), NOT_CONTEXT) > 0 Then
      bMatch = not CheckCode(MyCStr(parBusinessUnit), Replace(MyCStr(dsTemp("BusinessUnit")), NOT_CONTEXT, ""))
    Else
      bMatch = CheckCode(MyCStr(parBusinessUnit), MyCStr(dsTemp("BusinessUnit")))
    End If
AddLogD "bMatch: " & CStr(bMatch)
    If bMatch Then
      parRegistrar = MyCStr(dsTemp(sRegistrarField))
AddLogD "parRegistrar: " & parRegistrar
      parListToReconcile = MyCStr(dsTemp(sListToReconcileField))
AddLogD "parListToReconcile: " & parListToReconcile
      Exit Do
    End If
    dsTemp.MoveNext
  Loop
  dsTemp.Close

  If NewConnection Then
    MyConn.Close
  End If
End Sub

'Для справки структуры таблиц правил и лога их использования
'--Правила
' CREATE TABLE [dbo].[STS_Rules](
	' [OrderNo] [nvarchar](16) NOT NULL,
	' [Active] [nvarchar](1) NULL,
	' [ClassDoc] [nvarchar](255) NULL,
	' [AmountMin] [money] NULL,
	' [AmountMax] [money] NULL,
	' [ChartOfAccounts] [nvarchar](255) NULL,
	' [CostCenter] [nvarchar](255) NULL,
	' [ProjectCode] [nvarchar](255) NULL,
	' [BusinessUnit] [nvarchar](255) NULL,
	' [PartnerName] [nvarchar](255) NULL,
	' [KindOfPayment] [nvarchar](255) NULL,
	' [ListToReconcile] [nvarchar](1024) NULL,
	' [NameAproval] [nvarchar](128) NULL,
	' [NameResponsible] [nvarchar](128) NULL,
	' [Correspondent] [nvarchar](1024) NULL,
	' [ListToView] [nvarchar](1024) NULL,
	' [Registrar] [nvarchar](255) NULL,
	' [NameControl] [nvarchar](128) NULL,
	' [ListToEdit] [nvarchar](1024) NULL,
	' [Comment] [nvarchar](1024) NULL,
 ' CONSTRAINT [PK_STS_Rules] PRIMARY KEY CLUSTERED 
' (
	' [OrderNo] ASC
' )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
' ) ON [PRIMARY]


'--Лог
' CREATE TABLE [dbo].[STS_RulesLog](
	' [GUID] [uniqueidentifier] NOT NULL,
	' [DocID] [nvarchar](128) NULL,
	' [DateCreation] [datetime] NULL,
	' [RuleOrderNo] [nvarchar](16) NULL,
	' [Amount] [money] NULL,
	' [ChartOfAccounts] [nvarchar](1024) NULL,
	' [CostCenter] [nvarchar](1024) NULL,
	' [ProjectCode] [nvarchar](1024) NULL,
	' [BusinessUnit] [nvarchar](64) NULL,
	' [PartnerName] [nvarchar](255) NULL,
	' [KindOfPayment] [nvarchar](255) NULL,
 ' CONSTRAINT [PK_STS_RulesLog] PRIMARY KEY CLUSTERED 
' (
	' [GUID] ASC
' )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
' ) ON [PRIMARY]

'Подстрока, задающая введенную пользователем часть списка согласования
USER_INPUT = "#USERINPUT#"

'Основная подпрограмма для вызова извне. Принимает параметры документа, возвращает списки пользователей (parListToReconcile, parNameAproval, ...) по справочнику правил
'Поля уже могут содержать значения. Эти значения объединяются с расчетными.
'ListToReconcile собирается из правил, если есть контекст, заданный переменной USER_INPUT, то он заменяется на введенный в документ список.
'Из пользовательского списка удаляются пользователи добавленные правилами после него, потом удаляются дубли от начала полного списка до конца. 
'Остальные списки дополняются через VbCrLf, однострочные поля через пробел.
'При появлении нового значения в поле NameAproval, старое добавляется в список согласования и ставится новое.
'parInitiator может содержать логин или полное имя с логином.

'Запрос №43 - СТС - start
'{Запрос №50 - СТС
Sub GetReconciliationByRules(parDocID, parClassDoc, parAmount, parChartOfAccounts, parCostCenter, parProjectCode, parBusinessUnit, parPartnerName, parKindOfPayment, parInitiator, parProjectManager, parOvertimeRequester, parIncomeExpenceContract, parContranctInSTS, parCurrency, parContractType, parOvertimeFuncLeaders, parTypeOfDocument, parFunctionArea, ByRef parListToReconcile, ByRef parNameAproval, ByRef parNameResponsible, ByRef parCorrespondent, ByRef parListToView, ByRef parRegistrar, ByRef parNameControl, ByRef parListToEdit)
'Sub GetReconciliationByRules(parDocID, parClassDoc, parAmount, parChartOfAccounts, parCostCenter, parProjectCode, parBusinessUnit, parPartnerName, parKindOfPayment, parInitiator, parProjectManager, parOvertimeRequester, parIncomeExpenceContract, parContranctInSTS, parCurrency, parContractType, parOvertimeFuncLeaders, ByRef parListToReconcile, ByRef parNameAproval, ByRef parNameResponsible, ByRef parCorrespondent, ByRef parListToView, ByRef parRegistrar, ByRef parNameControl, ByRef parListToEdit)
'Запрос №50 - СТС}
'Sub GetReconciliationByRules(parDocID, parClassDoc, parAmount, parChartOfAccounts, parCostCenter, parProjectCode, parBusinessUnit, parPartnerName, parKindOfPayment, parInitiator, parProjectManager, parOvertimeRequester, parIncomeExpenceContract, parContranctInSTS, parCurrency, parContractType, ByRef parListToReconcile, ByRef parNameAproval, ByRef parNameResponsible, ByRef parCorrespondent, ByRef parListToView, ByRef parRegistrar, ByRef parNameControl, ByRef parListToEdit)
  Dim sSQL
  Dim NewConnection, MyConn, dsTemp, dsLog
  Dim sInitiatorsDepartment, sInitiatorsDepartmentCode, sInitiatorsDepartmentLevel, sCostCenterLevel, sCostCenterDepartment
  Dim bMatch
  Dim dCurrentDate
  Dim sListToReconcile, sListToReconcileLast, iPos
  Dim sRoles
  Dim sInitiator
  Dim nAmount, sAmount
  Dim nUSDRate, sUSDRate
  Dim AmountMinUSD, AmountMaxUSD
  Dim sPartnerRating, sPartnerName

  NewConnection = True
  If not IsNull(Conn) Then
    NewConnection = not IsObject(Conn)
  End If

  If NewConnection Then
    SetPayDoxPars 'Присваиваем переменные среды из Global.asa
    Set MyConn = Server.CreateObject("ADODB.Connection")
    MyConn.Open Application("ConnectStringRUS")
  Else
    Set MyConn = Conn
  End If

  'Разбирательство с курсами валют
  nAmount = 0
  'Вне всяких условий, т.к. в любом случае идет запись в лог коэффициента пересчета и может быть деление на ноль
  nUSDRate = CCur(1)
  If MyCStr(parAmount) <> "" and IsNumeric(parAmount) Then
    nAmount = parAmount
    Set dsTemp = Server.CreateObject("ADODB.Recordset")
    sSQL = "Select IsNULL(Rate, 0) as Rate from CurrencyRates where Code = "+sUnicodeSymbol+"'USD' or Code2 = "+sUnicodeSymbol+"'USD'"
    dsTemp.Open sSQL, MyConn, 3, 1, &H1
    If not dsTemp.EOF Then
	  If not IsNull(dsTemp("Rate")) Then
	    If CCur(dsTemp("Rate")) > 0 Then
          nUSDRate = CCur(dsTemp("Rate"))
        End If
	  End If
    End If
    dsTemp.Close

	If MyCStr(parCurrency) <> "" and MyCStr(parCurrency) <> "USD" Then
      sSQL = "Select IsNULL(Rate, 0) as Rate from CurrencyRates where Code = "+sUnicodeSymbol+"'"+parCurrency+"' or Code2 = "+sUnicodeSymbol+"'"+parCurrency+"'"
      dsTemp.Open sSQL, MyConn, 3, 1, &H1
      If not dsTemp.EOF Then
	    nAmount = nAmount*CCur(dsTemp("Rate"))/nUSDRate
      End If
      dsTemp.Close
	End If
  End If

  sSQL = "Select STS_Rules.*, CurrencyRates.Rate from STS_Rules left join CurrencyRates on (CurrencyRates.Code = STS_Rules.Currency or CurrencyRates.Code2 = STS_Rules.Currency) where ClassDoc = N'" & parClassDoc & "' and "
  If MyCStr(parAmount) <> "" and IsNumeric(parAmount) Then
    sUSDRate = CorrectSum(MyCStr(nUSDRate))
    AmountMinUSD = "IsNull(AmountMin, 0)*IsNull(Rate, 0)/" & sUSDRate
	AmountMaxUSD = "IsNull(AmountMax, 0)*IsNull(Rate, 0)/" & sUSDRate
	sAmount = CorrectSum(MyCStr(nAmount))
    sSQL = sSQL & "(" & AmountMinUSD & " = 0 or " & AmountMinUSD & " <> 0 and " & AmountMinUSD & " <= " & sAmount & ") and (" & AmountMaxUSD & " = 0 or " & AmountMaxUSD & " <> 0 and " & AmountMaxUSD & " >= " & sAmount & ") and "
  End If
  
  If not IsNull(parPartnerName) Then
    sPartnerName = MyCStr(parPartnerName)
    If Trim(sPartnerName) = "" Then
      sPartnerName = EMPTY_CONTEXT
    End If
    sSQL = sSQL & "(IsNull(PartnerName, N'') = N'' or (CharIndex(N'" & NOT_CONTEXT & "', PartnerName) = 0 and CharIndex(N'" & MakeSQLSafeSimple(MyCStr(sPartnerName)) & "', PartnerName) > 0 or CharIndex(N'" & NOT_CONTEXT & "', PartnerName) <> 0 and CharIndex(N'" & MakeSQLSafeSimple(MyCStr(sPartnerName)) & "', Replace(PartnerName, N'" & NOT_CONTEXT & "', N'')) <= 0 or CharIndex(N'" & NOT_CONTEXT & "', PartnerName) <> 0 and CharIndex(N'" & EMPTY_CONTEXT & "', PartnerName) <> 0 and N'" & MakeSQLSafeSimple(MyCStr(sPartnerName)) & "' <> N'" & EMPTY_CONTEXT & "')) and "
  End If

  sSQLPattern = "(IsNull(#FIELDNAME#, N'') = N'' or (CharIndex(N'" & NOT_CONTEXT & "', #FIELDNAME#) = 0 and CharIndex(N'#FIELDVALUE#,', #FIELDNAME#+N',') > 0 or CharIndex(N'" & NOT_CONTEXT & "', #FIELDNAME#) <> 0 and CharIndex(N'#FIELDVALUE#,', Replace(#FIELDNAME#, N'" & NOT_CONTEXT & "', N'')+N',') <= 0)) and "
  If MyCStr(parKindOfPayment) <> "" Then
    sSQL = sSQL & Replace(Replace(sSQLPattern, "#FIELDNAME#", "KindOfPayment"), "#FIELDVALUE#", MakeSQLSafeSimple(MyCStr(parKindOfPayment)))
  End If
  If MyCStr(parContractType) <> "" Then
    sSQL = sSQL & Replace(Replace(sSQLPattern, "#FIELDNAME#", "ContractType"), "#FIELDVALUE#", MakeSQLSafeSimple(MyCStr(parContractType)))
  End If
  If MyCStr(parIncomeExpenceContract) <> "" Then
    sSQL = sSQL & Replace(Replace(sSQLPattern, "#FIELDNAME#", "IncomeExpenceContract"), "#FIELDVALUE#", MakeSQLSafeSimple(MyCStr(parIncomeExpenceContract)))
  End If
  If MyCStr(parContranctInSTS) <> "" Then
    sSQL = sSQL & Replace(Replace(sSQLPattern, "#FIELDNAME#", "ContranctInSTS"), "#FIELDVALUE#", MakeSQLSafeSimple(MyCStr(parContranctInSTS)))
  End If
  If MyCStr(parTypeOfDocument) <> "" Then
    sSQL = sSQL & Replace(Replace(sSQLPattern, "#FIELDNAME#", "TypeOfDocument"), "#FIELDVALUE#", MakeSQLSafeSimple(MyCStr(parTypeOfDocument)))
  End If
  If MyCStr(parFunctionArea) <> "" Then
    sSQL = sSQL & Replace(Replace(sSQLPattern, "#FIELDNAME#", "FunctionArea"), "#FIELDVALUE#", MakeSQLSafeSimple(MyCStr(parFunctionArea)))
  End If
'  If MyCStr(parKindOfPayment) <> "" Then
'    sSQL = sSQL & "(IsNull(KindOfPayment, N'') = N'' or (CharIndex(N'" & NOT_CONTEXT & "', KindOfPayment) = 0 and CharIndex(N'" & MakeSQLSafeSimple(MyCStr(parKindOfPayment)) & ",', KindOfPayment+N',') > 0 or CharIndex(N'" & NOT_CONTEXT & "', KindOfPayment) <> 0 and CharIndex(N'" & MakeSQLSafeSimple(MyCStr(parKindOfPayment)) & ",', Replace(KindOfPayment, N'" & NOT_CONTEXT & "', N'')+N',') <= 0)) and "
'  End If
'  If MyCStr(parContractType) <> "" Then
'    sSQL = sSQL & "(IsNull(ContractType, N'') = N'' or (CharIndex(N'" & NOT_CONTEXT & "', ContractType) = 0 and CharIndex(N'" & MakeSQLSafeSimple(MyCStr(parContractType)) & ",', ContractType+N',') > 0 or CharIndex(N'" & NOT_CONTEXT & "', ContractType) <> 0 and CharIndex(N'" & MakeSQLSafeSimple(MyCStr(parContractType)) & ",', Replace(ContractType, N'" & NOT_CONTEXT & "', N'')+N',') <= 0)) and "
'  End If
'  If MyCStr(parIncomeExpenceContract) <> "" Then
'    sSQL = sSQL & "(IsNull(IncomeExpenceContract, N'') = N'' or (CharIndex(N'" & NOT_CONTEXT & "', IncomeExpenceContract) = 0 and CharIndex(N'" & MakeSQLSafeSimple(MyCStr(parIncomeExpenceContract)) & ",', IncomeExpenceContract+N',') > 0 or CharIndex(N'" & NOT_CONTEXT & "', IncomeExpenceContract) <> 0 and CharIndex(N'" & MakeSQLSafeSimple(MyCStr(parIncomeExpenceContract)) & ",', Replace(IncomeExpenceContract, N'" & NOT_CONTEXT & "', N'')+N',') <= 0)) and "
'  End If
'  If MyCStr(parContranctInSTS) <> "" Then
'    sSQL = sSQL & "(IsNull(ContranctInSTS, N'') = N'' or (CharIndex(N'" & NOT_CONTEXT & "', ContranctInSTS) = 0 and CharIndex(N'" & MakeSQLSafeSimple(MyCStr(parContranctInSTS)) & ",', ContranctInSTS+N',') > 0 or CharIndex(N'" & NOT_CONTEXT & "', ContranctInSTS) <> 0 and CharIndex(N'" & MakeSQLSafeSimple(MyCStr(parContranctInSTS)) & ",', Replace(ContranctInSTS, N'" & NOT_CONTEXT & "', N'')+N',') <= 0)) and "
'  End If
'  If MyCStr(parTypeOfDocument) <> "" Then
'    sSQL = sSQL & "(IsNull(TypeOfDocument, N'') = N'' or (CharIndex(N'" & NOT_CONTEXT & "', TypeOfDocument) = 0 and CharIndex(N'" & MakeSQLSafeSimple(MyCStr(parTypeOfDocument)) & ",', TypeOfDocument+N',') > 0 or CharIndex(N'" & NOT_CONTEXT & "', TypeOfDocument) <> 0 and CharIndex(N'" & MakeSQLSafeSimple(MyCStr(parTypeOfDocument)) & ",', Replace(TypeOfDocument, N'" & NOT_CONTEXT & "', N'')+N',') <= 0)) and "
'  End If
'  If MyCStr(parFunctionArea) <> "" Then
'    sSQL = sSQL & "(IsNull(FunctionArea, N'') = N'' or (CharIndex(N'" & NOT_CONTEXT & "', FunctionArea) = 0 and CharIndex(N'" & MakeSQLSafeSimple(MyCStr(parFunctionArea)) & ",', FunctionArea+N',') > 0 or CharIndex(N'" & NOT_CONTEXT & "', FunctionArea) <> 0 and CharIndex(N'" & MakeSQLSafeSimple(MyCStr(parFunctionArea)) & ",', Replace(FunctionArea, N'" & NOT_CONTEXT & "', N'')+N',') <= 0)) and "
'  End If


  If LCase(Trim(Right(sSQL, 4))) = "and" Then
    sSQL = Left(sSQL, Len(sSQL)-4)
  End If
  sSQL = sSQL & " and Active = 'Y' order by OrderNo"
AddLogD "GetReconciliationByRules - SQL: "+sSQL

  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  dsTemp.CursorLocation = 3
  dsTemp.Open sSQL, MyConn, 3, 1, &H1
  dsTemp.ActiveConnection = Nothing
AddLogD "Rules in recordset: " & CStr(dsTemp.RecordCount)

  'Таблица для лога
  Set dsLog = Server.CreateObject("ADODB.Recordset")
  dsLog.Open "select top 1 * from STS_RulesLog", MyConn, 1, 2, &H1
  dCurrentDate = Now

  'parNameAproval = ""
  'parNameResponsible = ""
  'parRegistrar = ""
  'parNameControl = ""
  parNameAproval = MyCStr(parNameAproval)
  parNameResponsible = MyCStr(parNameResponsible)
  parRegistrar = MyCStr(parRegistrar)
  parNameControl = MyCStr(parNameControl)
  parCorrespondent = MyCStr(parCorrespondent)
  parListToView = MyCStr(parListToView)
  parListToEdit = MyCStr(parListToEdit)
  'список согласования
  sListToReconcile = ""

  'Определяем вспомогательные параметры для условий
  sInitiatorsDepartmentCode = ""
  sInitiatorsDepartmentLevel = ""
  If Trim(MyCStr(parInitiator)) <> "" Then
    sInitiator = GetUserID(parInitiator)
    If sInitiator = "" Then
      sInitiator = Trim(MyCStr(parInitiator))
    End If

    oPayDox.GetUserDetails sInitiator, sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
    sInitiatorsDepartment = Trim(MyCStr(sDepartment))
    If sInitiatorsDepartment <> "" Then
      sInitiatorsDepartmentCode = GetSTSDepartmentCode(sInitiatorsDepartment)
      sInitiatorsDepartmentLevel = Replace(GetDepartmentLevel(sInitiatorsDepartment), "#LEV", "")
    End If
  End If
  AddLogD "#################################parCostCenter: " & parCostCenter
  If MyCStr(parCostCenter) <> "" Then
    sCostCenterDepartment = GetCostCenterByCode(MyCStr(parCostCenter))
    AddLogD "#################################sCostCenterDepartment: " & sCostCenterDepartment
    sCostCenterLevel = Replace(GetDepartmentLevel(sCostCenterDepartment), "#LEV", "")
  Else
    sCostCenterDepartment = ""
    sCostCenterLevel = ""
  End If
AddLogD "#################################sCostCenterDepartment: " & sCostCenterDepartment

  Do While not dsTemp.EOF
AddLogD "-------------- Rule OrderNo: " & MyCStr(dsTemp("OrderNo"))
    bMatch = True
AddLogD "parChartOfAccounts: """ & MyCStr(parChartOfAccounts) & """"
AddLogD "dsTemp(ChartOfAccounts): """ & MyCStr(dsTemp("ChartOfAccounts")) & """"
    If MyCStr(parChartOfAccounts) <> "" Then
      If MyCStr(dsTemp("ChartOfAccounts")) <> "" Then
        If InStr(MyCStr(dsTemp("ChartOfAccounts")), NOT_CONTEXT) > 0 Then
          bMatch = not CheckCode(MyCStr(parChartOfAccounts), Replace(MyCStr(dsTemp("ChartOfAccounts")), NOT_CONTEXT, ""))
        Else
          bMatch = CheckCode(MyCStr(parChartOfAccounts), MyCStr(dsTemp("ChartOfAccounts")))
        End If
      End If
    End If
AddLogD "bMatch: " & CStr(bMatch)

    If bMatch Then
AddLogD "parCostCenter: """ & MyCStr(parCostCenter) & """"
AddLogD "dsTemp(CostCenter): """ & MyCStr(dsTemp("CostCenter")) & """"
      If MyCStr(parCostCenter) <> "" Then
        If MyCStr(dsTemp("CostCenter")) <> "" Then
          If InStr(MyCStr(dsTemp("CostCenter")), NOT_CONTEXT) > 0 Then
            bMatch = not CheckCode(MyCStr(parCostCenter), Replace(MyCStr(dsTemp("CostCenter")), NOT_CONTEXT, ""))
          Else
            bMatch = CheckCode(MyCStr(parCostCenter), MyCStr(dsTemp("CostCenter")))
          End If
        End If
      End If
AddLogD "bMatch: " & CStr(bMatch)
    End If

    If bMatch Then
AddLogD "parProjectCode: """ & MyCStr(parProjectCode) & """"
AddLogD "dsTemp(ProjectCode): """ & MyCStr(dsTemp("ProjectCode")) & """"
      If MyCStr(parProjectCode) <> "" Then
        If MyCStr(dsTemp("ProjectCode")) <> "" Then
          If InStr(MyCStr(dsTemp("ProjectCode")), NOT_CONTEXT) > 0 Then
            bMatch = not CheckCode(MyCStr(parProjectCode), Replace(MyCStr(dsTemp("ProjectCode")), NOT_CONTEXT, ""))
          Else
            bMatch = CheckCode(MyCStr(parProjectCode), MyCStr(dsTemp("ProjectCode")))
          End If
        End If
      End If
AddLogD "bMatch: " & CStr(bMatch)
    End If

    If bMatch Then
AddLogD "parBusinessUnit: """ & MyCStr(parBusinessUnit) & """"
AddLogD "dsTemp(BusinessUnit): """ & MyCStr(dsTemp("BusinessUnit")) & """"
      If MyCStr(parBusinessUnit) <> "" Then
        If MyCStr(dsTemp("BusinessUnit")) <> "" Then
          If InStr(MyCStr(dsTemp("BusinessUnit")), NOT_CONTEXT) > 0 Then
            bMatch = not CheckCode(MyCStr(parBusinessUnit), Replace(MyCStr(dsTemp("BusinessUnit")), NOT_CONTEXT, ""))
          Else
            bMatch = CheckCode(MyCStr(parBusinessUnit), MyCStr(dsTemp("BusinessUnit")))
          End If
        End If
      End If
AddLogD "bMatch: " & CStr(bMatch)
    End If

    If bMatch Then
AddLogD "sInitiatorsDepartmentCode: """ & MyCStr(sInitiatorsDepartmentCode) & """"
AddLogD "dsTemp(InitiatorsDepartmentCode): """ & MyCStr(dsTemp("InitiatorsDepartmentCode")) & """"
      If MyCStr(sInitiatorsDepartmentCode) <> "" Then
        If MyCStr(dsTemp("InitiatorsDepartmentCode")) <> "" Then
          If InStr(MyCStr(dsTemp("InitiatorsDepartmentCode")), NOT_CONTEXT) > 0 Then
            bMatch = not CheckCode(MyCStr(sInitiatorsDepartmentCode), Replace(MyCStr(dsTemp("InitiatorsDepartmentCode")), NOT_CONTEXT, ""))
          Else
            bMatch = CheckCode(MyCStr(sInitiatorsDepartmentCode), MyCStr(dsTemp("InitiatorsDepartmentCode")))
          End If
        End If
      End If
AddLogD "bMatch: " & CStr(bMatch)
    End If

    If bMatch Then
AddLogD "sInitiatorsDepartmentLevel: """ & MyCStr(sInitiatorsDepartmentLevel) & """"
AddLogD "dsTemp(InitiatorsDepartmentLevel): """ & MyCStr(dsTemp("InitiatorsDepartmentLevel")) & """"
      If MyCStr(sInitiatorsDepartmentLevel) <> "" Then
        If MyCStr(dsTemp("InitiatorsDepartmentLevel")) <> "" Then
          'Уровни перечисляются через запятую, перед поиском удаляем пробелы
          bMatch = InStr("," & Replace(MyCStr(dsTemp("InitiatorsDepartmentLevel")), " ", "") & ",", "," & MyCStr(sInitiatorsDepartmentLevel) & ",") > 0
        End If
      End If
AddLogD "bMatch: " & CStr(bMatch)
    End If

    If bMatch Then
AddLogD "sCostCenterLevel: """ & MyCStr(sCostCenterLevel) & """"
AddLogD "dsTemp(CostCenterLevel): """ & MyCStr(dsTemp("CostCenterLevel")) & """"
      If MyCStr(sCostCenterLevel) <> "" Then
        If MyCStr(dsTemp("CostCenterLevel")) <> "" Then
          'Уровни перечисляются через запятую, перед поиском удаляем пробелы
          bMatch = InStr("," & Replace(MyCStr(dsTemp("CostCenterLevel")), " ", "") & ",", "," & MyCStr(sCostCenterLevel) & ",") > 0
        End If
      End If
AddLogD "bMatch: " & CStr(bMatch)
    End If
	
    If bMatch Then
AddLogD "parPartnerName: """ & MyCStr(parPartnerName) & """"
AddLogD "dsTemp(PartnerRating): """ & MyCStr(dsTemp("PartnerRating")) & """"
	  If MyCStr(parPartnerName) <> "" and MyCStr(dsTemp("PartnerRating")) <> "" Then
	    sPartnerRating = GetPartnerRating(parPartnerName)
		If Trim(sPartnerRating) = "" Then
		  sPartnerRating = EMPTY_CONTEXT
		End If
AddLogD "sPartnerRating: """ & sPartnerRating & """"
        'Варианты перечисляются через запятую (+,-,#EMPTY#), перед поиском удаляем пробелы
        bMatch = InStr("," & Replace(MyCStr(dsTemp("PartnerRating")), " ", "") & ",", "," & sPartnerRating & ",") > 0
      End If
AddLogD "bMatch: " & CStr(bMatch)
    End If

    If bMatch Then
      'Пишем в лог, что правило задействовано
      dsLog.AddNew
      dsLog("GUID") = oPayDox.GetGUID1()
      dsLog("DocID") = parDocID
      dsLog("DateCreation") = dCurrentDate
      dsLog("RuleOrderNo") = MyCStr(dsTemp("OrderNo"))
      dsLog("Amount") = parAmount
      dsLog("ChartOfAccounts") = MyCStr(parChartOfAccounts)
      dsLog("CostCenter") = MyCStr(parCostCenter)
      dsLog("ProjectCode") = MyCStr(parProjectCode)
      dsLog("BusinessUnit") = MyCStr(parBusinessUnit)
      dsLog("PartnerName") = MyCStr(sPartnerName)
      dsLog("KindOfPayment") = MyCStr(parKindOfPayment)
      dsLog("CostCenterLevel") = MyCStr(sCostCenterLevel)
      dsLog("InitiatorsDepartmentLevel") = MyCStr(sInitiatorsDepartmentLevel)
      dsLog("InitiatorsDepartmentCode") = MyCStr(sInitiatorsDepartmentCode)
      dsLog("ContractType") = MyCStr(parContractType)
      dsLog("ContranctInSTS") = MyCStr(parContranctInSTS)
      dsLog("TypeOfDocument") = MyCStr(parTypeOfDocument)
      dsLog("FunctionArea") = MyCStr(parFunctionArea)
      dsLog("IncomeExpenceContract") = MyCStr(parIncomeExpenceContract)
      dsLog("PartnerRating") = MyCStr(sPartnerRating)
      dsLog("Currency") = MyCStr(parCurrency)
      dsLog("USDAmount") = nAmount
	  dsLog("RuleCurrencyToUSDConvFactor") = iif(IsNull(dsTemp("Rate")), 0, dsTemp("Rate")) / nUSDRate
      dsLog.Update

      'Список согласующих и утверждающий
AddLogD "Initial ListToReconcile: " & MyCStr(sListToReconcile)
AddLogD "Adding ListToReconcile: " & MyCStr(dsTemp("ListToReconcile"))
      If MyCStr(dsTemp("ListToReconcile")) <> "" Then
        sListToReconcile = sListToReconcile & MyCStr(dsTemp("ListToReconcile")) & " "
      End If
AddLogD "Result ListToReconcile1: " & MyCStr(sListToReconcile)
AddLogD "Initial NameAproval: " & MyCStr(parNameAproval)
AddLogD "Adding NameAproval: " & MyCStr(dsTemp("NameAproval"))
      If MyCStr(dsTemp("NameAproval")) <> "" Then
        'Если уже есть утверждающий, то ставим его последовательно в согласующих, т.е. приоритет правил растет с ростом порядкового номера
        If Trim(parNameAproval) <> "" Then
          sListToReconcile = sListToReconcile & VbCrLf & parNameAproval & VbCrLf
        End If
        parNameAproval = MyCStr(dsTemp("NameAproval"))
      End If
AddLogD "Result NameAproval: " & MyCStr(parNameAproval)
AddLogD "Result ListToReconcile2: " & MyCStr(sListToReconcile)

      'Список адресатов
      parCorrespondent = Trim(MyCStr(parCorrespondent))
AddLogD "Initial Correspondent: " & MyCStr(parCorrespondent)
AddLogD "Adding Correspondent: " & MyCStr(dsTemp("Correspondent"))
      If MyCStr(dsTemp("Correspondent")) <> "" Then
        If parCorrespondent = "" or InStr(MyCStr(dsTemp("Correspondent")), parCorrespondent) = 0 Then
          If parCorrespondent <> "" Then
            parCorrespondent = parCorrespondent & VbCrLf
          End If
          parCorrespondent = parCorrespondent & MyCStr(dsTemp("Correspondent"))
        End If
      End If
AddLogD "Result Correspondent: " & MyCStr(parCorrespondent)

      'Список ознакомления
      parListToView = Trim(MyCStr(parListToView))
AddLogD "Initial ListToView: " & MyCStr(parListToView)
AddLogD "Adding ListToView: " & MyCStr(dsTemp("ListToView"))
      If MyCStr(dsTemp("ListToView")) <> "" Then
        If parListToView = "" or InStr(MyCStr(dsTemp("ListToView")), parListToView) = 0 Then
          If parListToView <> "" Then
            parListToView = parListToView & VbCrLf
          End If
          parListToView = parListToView & MyCStr(dsTemp("ListToView"))
        End If
      End If
AddLogD "Result ListToView: " & MyCStr(parListToView)

      'Список имеющих право редактирования
      parListToEdit = Trim(MyCStr(parListToEdit))
AddLogD "Initial ListToEdit: " & MyCStr(parListToEdit)
AddLogD "Adding ListToEdit: " & MyCStr(dsTemp("ListToEdit"))
      If MyCStr(dsTemp("ListToEdit")) <> "" Then
        If parListToEdit = "" or InStr(MyCStr(dsTemp("ListToEdit")), parListToEdit) = 0 Then
          If parListToEdit <> "" Then
            parListToEdit = parListToEdit & VbCrLf
          End If
          parListToEdit = parListToEdit & MyCStr(dsTemp("ListToEdit"))
        End If
      End If
AddLogD "Result ListToEdit: " & MyCStr(parListToEdit)

      'Исполнитель(и)
      parNameResponsible = Trim(MyCStr(parNameResponsible))
AddLogD "Initial NameResponsible: " & MyCStr(parNameResponsible)
AddLogD "Adding NameResponsible: " & MyCStr(dsTemp("NameResponsible"))
      If MyCStr(dsTemp("NameResponsible")) <> "" Then
        If parNameResponsible = "" or InStr(MyCStr(dsTemp("NameResponsible")), parNameResponsible) = 0 Then
          If parNameResponsible <> "" Then
            parNameResponsible = parNameResponsible & " "
          End If
          parNameResponsible = parNameResponsible & MyCStr(dsTemp("NameResponsible"))
        End If
      End If
AddLogD "Result NameResponsible: " & MyCStr(parNameResponsible)

      'Регистраторы
      parRegistrar = Trim(MyCStr(parRegistrar))
AddLogD "Initial Registrar: " & MyCStr(parRegistrar)
AddLogD "Adding Registrar: " & MyCStr(dsTemp("Registrar"))
      If MyCStr(dsTemp("Registrar")) <> "" Then
        If parRegistrar = "" or InStr(MyCStr(dsTemp("Registrar")), parRegistrar) = 0 Then
          If parRegistrar <> "" Then
            parRegistrar = parRegistrar & " "
          End If
          parRegistrar = parRegistrar & MyCStr(dsTemp("Registrar"))
        End If
      End If
AddLogD "Result Registrar: " & MyCStr(parRegistrar)

      'Контролеры
      parNameControl = Trim(MyCStr(parNameControl))
AddLogD "Initial NameControl: " & MyCStr(parNameControl)
AddLogD "Adding NameControl: " & MyCStr(dsTemp("NameControl"))
      If MyCStr(dsTemp("NameControl")) <> "" Then
        If parNameControl = "" or InStr(MyCStr(dsTemp("NameControl")), parNameControl) = 0 Then
          If parNameControl <> "" Then
            parNameControl = parNameControl & " "
          End If
          parNameControl = parNameControl & MyCStr(dsTemp("NameControl"))
        End If
      End If
AddLogD "Result NameControl: " & MyCStr(parNameControl)
    End If

    dsTemp.MoveNext
  Loop
  dsLog.Close
  dsTemp.Close

  If NewConnection Then
    MyConn.Close
  End If

  'Получаем список ролей
  sRoles = GetAllRolesList(sCostCenterDepartment, sInitiator, parBusinessUnit, parProjectManager, parOvertimeRequester, parOvertimeFuncLeaders)
  'Заменяем роли во всех списках и удаляем дубли
  sListToReconcile = ReplaceRolesInList(sListToReconcile, sRoles)
AddLogD "ListToReconcile after replacing roles: " & MyCStr(sListToReconcile)
  iPos = InStr(sListToReconcile, USER_INPUT)
  If iPos > 0 Then
    sListToReconcileLast = Right(sListToReconcile, Len(sListToReconcile)-iPos-Len(USER_INPUT)+1)
AddLogD "ListToReconcile last part: " & MyCStr(sListToReconcileLast)
    'Если вхождений USER_INPUT больше одного, то последующие игнорируем
    sListToReconcileLast = Replace(sListToReconcileLast, USER_INPUT, "")
AddLogD "ListToReconcile last part after removing another " & USER_INPUT & ": " & MyCStr(sListToReconcileLast)
    sListToReconcile = Left(sListToReconcile, iPos-1)
AddLogD "ListToReconcile first part: " & MyCStr(sListToReconcile)
    'Соединяем пользовательскую часть и заднюю из правил и удаляем дубли с конца
    sListToReconcileLast = DeleteUserDoublesInListNew(ReplaceRolesInList(parListToReconcile, sRoles) & sListToReconcileLast, False)
AddLogD "ListToReconcile last part after adding user input: " & MyCStr(sListToReconcileLast)
    'Соединяем весь список и удаляем дубли с начала
    parListToReconcile = DeleteUserDoublesInListNew(sListToReconcile & sListToReconcileLast, True)
  Else
    parListToReconcile = DeleteUserDoublesInListNew(ReplaceRolesInList(sListToReconcile, sRoles), True)
  End If
AddLogD "Final ListToReconcile: " & MyCStr(parListToReconcile)

  parNameAproval = DeleteUserDoublesInListNew(ReplaceRolesInList(parNameAproval, sRoles), True)
AddLogD "NameAproval after replacing roles: " & MyCStr(parNameAproval)
  'Убираем приписку со временем из утверждающего
  parNameAproval = DeleteReconciliationTimeFromUser(parNameAproval)
AddLogD "NameAproval after removing time of reconciliation: " & MyCStr(parNameAproval)

'------------ Выбрасываем из списка согласования утверждающего (может быть это не всегда будет нужно)
'  parListToReconcile = Replace(parListToReconcile, parNameAproval, "")
  sNameAprovalLogin = """DELETE"" <" & GetUserID(parNameAproval) & ">;"
  parListToReconcile = DeleteUserDoublesInListNew(parListToReconcile & sNameAprovalLogin, False)
  parListToReconcile = Replace(parListToReconcile, sNameAprovalLogin, "")
AddLogD "ListToReconcile after deleting NameAproval: " & MyCStr(parListToReconcile)
'---------------------------------------------------------------------------------------------------------

  parNameResponsible = DeleteUserDoublesInList(ReplaceRolesInList(parNameResponsible, sRoles))
AddLogD "NameResponsible after replacing roles: " & MyCStr(parNameResponsible)
  parCorrespondent = DeleteUserDoublesInList(ReplaceRolesInList(parCorrespondent, sRoles))
AddLogD "Correspondent after replacing roles: " & MyCStr(parCorrespondent)
  parListToView = DeleteUserDoublesInList(ReplaceRolesInList(parListToView, sRoles))
AddLogD "ListToView after replacing roles: " & MyCStr(parListToView)
  parRegistrar = DeleteUserDoublesInList(ReplaceRolesInList(parRegistrar, sRoles))
AddLogD "Registrar after replacing roles: " & MyCStr(parRegistrar)
  parNameControl = DeleteUserDoublesInList(ReplaceRolesInList(parNameControl, sRoles))
AddLogD "NameControl after replacing roles: " & MyCStr(parNameControl)
  parListToEdit = DeleteUserDoublesInList(ReplaceRolesInList(parListToEdit, sRoles))
AddLogD "ListToEdit after replacing roles: " & MyCStr(parListToEdit)

  'Окончательная обработка списков, чтобы убрать лишние концы строк и пробелы
  parListToReconcile = CleanListToReconcile(parListToReconcile)
AddLogD "Final ListToReconcile after cleaning: " & MyCStr(parListToReconcile)
  parCorrespondent = CleanListToReconcile(parCorrespondent)
AddLogD "Final Correspondent after cleaning: " & MyCStr(parCorrespondent)
  parListToView = CleanListToReconcile(parListToView)
AddLogD "Final ListToView after cleaning: " & MyCStr(parListToView)
  parListToEdit = CleanListToReconcile(parListToEdit)
AddLogD "Final ListToEdit after cleaning: " & MyCStr(parListToEdit)
End Sub
'Запрос №43 - СТС - end
'Запрос №36 - СТС - end

'Запрос №35 - СТС - start
'Удаление дубликатов из списка согласующих
'parFromStart = True - поиск с начала списка (дубли удаляются в конце)
'parFromStart = False - поиск с конца списка (дубли удаляются в начале)
' Function DeleteUserDoublesInListNew(ByVal parUsersList, ByVal parFromStart)
  ' Dim nStr, vArr
  ' Dim iLast
  ' Dim i, j
  ' Dim sValue
  ' Dim LenCrLf

  ' If Trim(parUsersList) = "" or InStr(parUsersList, ";") = 0 Then
    ' DeleteUserDoublesInListNew = parUsersList
    ' Exit Function
  ' End If

  ' nStr = Split(parUsersList, ";")
  ' iLast = UBound(nStr)

  ' ReDim vArr(iLast)
  ' For i = 0 To iLast
    ' If parFromStart Then
      ' vArr(i) = Trim(nStr(i))
    ' Else
      ' vArr(iLast-i) = Trim(nStr(i))
    ' End If
  ' Next

  ' For i = 1 To iLast
    ' For j = 0 To i-1
      ' sValue = Trim(Replace(vArr(j), VbCrLf, ""))
      ' If sValue = Trim(Replace(vArr(i), VbCrLf, "")) Then
        ' vArr(i) = Trim(Replace(vArr(i), sValue, ""))
      ' End If
    ' Next
  ' Next

  ' DeleteUserDoublesInListNew = ""
  ' For i = 0 To UBound(vArr)
    ' If vArr(i) <> "" Then
      ' If parFromStart Then
        ' DeleteUserDoublesInListNew = DeleteUserDoublesInListNew & vArr(i) & ";"
      ' Else
        ' DeleteUserDoublesInListNew = vArr(i) & ";" & DeleteUserDoublesInListNew
      ' End If
    ' End If
  ' Next
  ' 'Чистка мусора
  ' DeleteUserDoublesInListNew = Replace(DeleteUserDoublesInListNew, VbCrLf & ";", VbCrLf)
  ' Do While InStr(DeleteUserDoublesInListNew, ";;") > 0
    ' DeleteUserDoublesInListNew = Replace(DeleteUserDoublesInListNew, ";;", ";")
  ' Loop
  ' Do While InStr(DeleteUserDoublesInListNew, VbCrLf & VbCrLf) > 0
    ' DeleteUserDoublesInListNew = Replace(DeleteUserDoublesInListNew, VbCrLf & VbCrLf, VbCrLf)
  ' Loop
  ' LenCrLf = Len(VbCrLf)
  ' Do While Right(DeleteUserDoublesInListNew, LenCrLf) = VbCrLf
    ' DeleteUserDoublesInListNew = Left(DeleteUserDoublesInListNew, Len(DeleteUserDoublesInListNew)-LenCrLf)
    ' DeleteUserDoublesInListNew = Trim(DeleteUserDoublesInListNew)
  ' Loop
' End Function

'Удаление дубликатов из списка согласующих
'parFromStart = True - поиск с начала списка (дубли удаляются в конце)
'parFromStart = False - поиск с конца списка (дубли удаляются в начале)
Function DeleteUserDoublesInListNew(ByVal parUsersList, ByVal parFromStart)
  Dim nStr, vArr
  Dim iLast
  Dim i, j
  Dim LenCrLf
  Dim iQuotPos, iLBracketPos, iRBracketPos, iSemicolonPos, iLTPos, iPos1, iPos2

  If Trim(parUsersList) = "" or InStr(parUsersList, ";") = 0 Then
    DeleteUserDoublesInListNew = parUsersList
    Exit Function
  End If

  nStr = Split(parUsersList, ">")
  iLast = UBound(nStr)
  For i = 0 To iLast-1
    iLTPos = InStr(nStr(i+1), "<")
    If iLTPos = 0 Then
      nStr(i) = nStr(i) & ">" & nStr(i+1)
      nStr(i+1) = ""
    Else
      iQuotPos = InStr(nStr(i+1), """")
      iLBracketPos = InStr(nStr(i+1), "(")
      iRBracketPos = InStr(nStr(i+1), ")")
      iSemicolonPos = InStr(nStr(i+1), ";")
      iPos1 = iLTPos
      If iQuotPos < iLTPos Then
        iPos1 = iQuotPos
      End If
      iPos2 = 0
      If iLBracketPos < iRBracketPos and iRBracketPos < iPos1 Then
        iPos2 = iRBracketPos
      End If
      If iPos2 < iSemicolonPos and iSemicolonPos < iPos1 Then
        iPos2 = iSemicolonPos
      End If
      If iPos2 > 0 Then
        nStr(i) = nStr(i) & ">" & Left(nStr(i+1), iPos2)
        nStr(i+1) = Mid(nStr(i+1), iPos2+1)
      End If
    End If
  Next
  If nStr(iLast) = "" Then
    iLast = iLast-1
  ElseIf InStr(nStr(iLast), "<") > 0 and InStr(nStr(iLast), ">") = 0 Then
    nStr(iLast) = nStr(iLast) & ">"
  End If

  ReDim vArr(iLast)
  For i = 0 To iLast
    If parFromStart Then
      vArr(i) = Trim(nStr(i))
    Else
      vArr(iLast-i) = Trim(nStr(i))
    End If
  Next

  For i = 1 To iLast
    For j = 0 To i-1
      If UCase(GetUserID(vArr(i))) = UCase(GetUserID(vArr(j))) Then
        If InStr(vArr(i), VbCrLf) > 0 Then
          vArr(i) = VbCrLf
        Else
          vArr(i) = ""
        End If
      End If
    Next
  Next

  DeleteUserDoublesInListNew = ""
  For i = 0 To UBound(vArr)
    If vArr(i) <> "" Then
      If parFromStart Then
        DeleteUserDoublesInListNew = DeleteUserDoublesInListNew & vArr(i)
      Else
        DeleteUserDoublesInListNew = vArr(i) & DeleteUserDoublesInListNew
      End If
    End If
  Next
  'Чистка мусора
  DeleteUserDoublesInListNew = Replace(DeleteUserDoublesInListNew, VbCrLf & ";", VbCrLf)
  Do While InStr(DeleteUserDoublesInListNew, ";;") > 0
    DeleteUserDoublesInListNew = Replace(DeleteUserDoublesInListNew, ";;", ";")
  Loop
  Do While InStr(DeleteUserDoublesInListNew, VbCrLf & VbCrLf) > 0
    DeleteUserDoublesInListNew = Replace(DeleteUserDoublesInListNew, VbCrLf & VbCrLf, VbCrLf)
  Loop
  LenCrLf = Len(VbCrLf)
  Do While Right(DeleteUserDoublesInListNew, LenCrLf) = VbCrLf
    DeleteUserDoublesInListNew = Left(DeleteUserDoublesInListNew, Len(DeleteUserDoublesInListNew)-LenCrLf)
    DeleteUserDoublesInListNew = Trim(DeleteUserDoublesInListNew)
  Loop
  Do While Left(DeleteUserDoublesInListNew, LenCrLf) = VbCrLf
    DeleteUserDoublesInListNew = Right(DeleteUserDoublesInListNew, Len(DeleteUserDoublesInListNew)-LenCrLf)
    DeleteUserDoublesInListNew = Trim(DeleteUserDoublesInListNew)
  Loop
  DeleteUserDoublesInListNew = Replace(DeleteUserDoublesInListNew, ";""", "; """)
  DeleteUserDoublesInListNew = Replace(DeleteUserDoublesInListNew, ")""", ") """)
  'Очистка от лишних времен согласования в скобках
  DeleteUserDoublesInListNew = CleanDoubleReconciliationTimes(DeleteUserDoublesInListNew)
End Function

'Очистка списка согласующих от лишних указаний времени согласования (в скобках), остается только первое указание из идущих подряд (2)(3)(8) -> (2)
Function CleanDoubleReconciliationTimes(parListToReconcile)
  Dim arStr, i, iLast, iLBPos

  CleanDoubleReconciliationTimes = ""
  If Len(parListToReconcile) = 0 Then
    Exit Function
  End If

  arStr = Split(parListToReconcile, ")")
  iLast = UBound(arStr)
  For i = 0 To iLast
    If i <> iLast or arStr(i) <> "" Then
      iLBPos = InStrRev(arStr(i), "(")
      If iLBPos = 0 Then
        CleanDoubleReconciliationTimes = CleanDoubleReconciliationTimes & arStr(i)
        If i <> iLast Then
          CleanDoubleReconciliationTimes = CleanDoubleReconciliationTimes & ")"
        End If
      Else
        If Trim(Replace(Left(arStr(i), iLBPos-1), VbCrLf, " ")) <> "" Then
          CleanDoubleReconciliationTimes = CleanDoubleReconciliationTimes & arStr(i)
          If i <> iLast Then
            CleanDoubleReconciliationTimes = CleanDoubleReconciliationTimes & ")"
          End If
        End If
      End If
	End If
  Next
End Function

Function DeleteReconciliationTimeFromUser(parSingleUser)
  Dim iLBPos, iRBPos

  DeleteReconciliationTimeFromUser = Trim(parSingleUser)
  iLBPos = InStrRev(DeleteReconciliationTimeFromUser, "(")
  iRBPos = InStrRev(DeleteReconciliationTimeFromUser, ")")
  If iLBPos > 0 and iRBPos > iLBPos Then
    DeleteReconciliationTimeFromUser = Trim(Left(DeleteReconciliationTimeFromUser, iLBPos-1) & Mid(DeleteReconciliationTimeFromUser, iRBPos+1, Len(DeleteReconciliationTimeFromUser)-iRBPos))
  End If
End Function

'Получить поле из справочника пользователей
Function GetUsersField(parUser, parField)
  Dim sUserID
  Dim NewConnection, MyConn, dsTemp

  sUserID = GetUserID(MyCStr(parUser))
  If sUserID = "" Then
    sUserID = Trim(MyCStr(parUser))
  End If

  GetUsersField = ""
  If sUserID <> "" Then
    NewConnection = True
    If not IsNull(Conn) Then
      NewConnection = not IsObject(Conn)
    End If

    If NewConnection Then
      SetPayDoxPars 'Присваиваем переменные среды из Global.asa
      sConnStr = "ConnectString"
      Select Case UCase(Request("l"))
        Case "RU" sConnStr = sConnStr & "RUS"
        Case "3" sConnStr = sConnStr & "3"
      End Select

      Set MyConn = Server.CreateObject("ADODB.Connection")
      MyConn.Open Application(sConnStr)
    Else
      Set MyConn = Conn
    End If

    Set dsTemp = Server.CreateObject("ADODB.Recordset")
    sSQL = "select " & parField & " from Users where UserID = " & sUnicodeSymbol & "'" & sUserID & "'"
    dsTemp.Open sSQL, MyConn, 3, 1, &H1
    If not dsTemp.EOF Then
      GetUsersField = dsTemp(parField)
    End If
    dsTemp.Close

    If NewConnection Then
      MyConn.Close
    End If
  End If
End Function
'Запрос №35 - СТС - end

'Запрос №38 - СТС - start
'Проверка числа согласующих на превышение максимального числа
Function CheckUsersInListToReconcile(parListToReconcile, parMaxUsers)
  Dim nMaxUsers, nUsers, iPos, sUserID

  nMaxUsers = 0
  If parMaxUsers <> "" and IsNumeric(parMaxUsers) Then
    nMaxUsers = CInt(parMaxUsers)
  End If
  
  If nMaxUsers = 0 Then
    CheckUsersInListToReconcile = True
	Exit Function
  End If
  
  iPos = 1
  nUsers = 0
  sUserID = oPayDox.GetNextUserIDInList(parListToReconcile, iPos)
  Do While sUserID <> ""
    nUsers = nUsers + 1
    sUserID = oPayDox.GetNextUserIDInList(parListToReconcile, iPos)
  Loop
'  AddLogD "CheckUsersInListToReconcile - nUsers = " & nUsers
'  AddLogD "CheckUsersInListToReconcile - nMaxUsers = " & nMaxUsers
  CheckUsersInListToReconcile = nUsers <= nMaxUsers
End Function
'Запрос №38 - СТС - end

'Запрос №41 - СТС - start
'Получить трехъязычное корневое подразделение
Function GetRootDepartmentFull(ByVal sDepartment)
  Dim iPos
  
  GetRootDepartmentFull = Trim(sDepartment)
  If GetRootDepartmentFull = "" Then
    Exit Function
  End If
  iPos = InStr(GetRootDepartmentFull, "/")
  If iPos > 0 Then
    GetRootDepartmentFull = Left(GetRootDepartmentFull, iPos-1)
  End If
End Function

'Получить список всех категорий документов для выпадающего списка в отчетах (возвращается два списка: отображаемый и вставляемых в отчет значений)
Sub GetCategoriesListForReport(ByRef parSelectList, ByRef parReportList)
  Dim ds, List
  
  Set ds = CreateObject("ADODB.Recordset")
  ds.Open "select * from DocTypes order by Name", Conn, 1, 3, &H1
  parSelectList = ""
  parReportList = ""
  If not ds.EOF Then
    parSelectList = DelOtherLangFromFolder(ds("Name"))
    parReportList = "ClassDoc = N'" & ds("Name") & "'"
  End If
  Do while not ds.EOF
    parSelectList = parSelectList & "," & DelOtherLangFromFolder(ds("Name"))
	parReportList = parReportList & "," & "ClassDoc = N'" & ds("Name") & "'"
    ds.MoveNext
  Loop
  ds.Close
End Sub

Function LinkToComment(parDocID, parKeyField)
  LinkToComment = "<a href=""ShowDoc.asp?l=" & Request("l") & "&DocID=" & URLEncode(parDocID) & "&bVISA=y&#c" & MyCStr(parKeyField) & """ title=""Перейти"" target=""_blank"">" & HTMLEncode(parDocID) & "</a>"
End Function
'Запрос №41 - СТС - end

'Запрос №1 - СИБ - start
'Расчет номеров служебных записок
Function GetNewDocIDForSluzhZapiski(parClassDoc, parDepartment, parSubClassParameter, parIsProjectDocID)
  Dim iDigits, sSearchCol, sSQL, sPrefix, sSeparator

  If parIsProjectDocID="PJ-" Then
    sSearchCol="DocIDadd"
  Else
    sSearchCol="DocID"
  End If

  Select Case parDepartment
    Case SIT_SITRONICS
	  sSeparator = "-"
	  sPrefix="IH."+Right(CStr(Year(Date)),2)+"/"+Left(parSubClassParameter, NotNegative(InStr(parSubClassParameter, " ")-1))+sSeparator
	  sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc like N'"+parClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%' AND Department like N'"+parDepartment+"%'"
      iDigits = 3
    Case SIT_RTI
	  sSeparator = "/"
	  sPrefix="IH"+Right(CStr(Year(Date)),2)+"_RTI"+sSeparator'+Left(parSubClassParameter, NotNegative(InStr(parSubClassParameter, " ")-1))+sSeparator
	  sSQL="select ISNULL(Max(cast(right(replace("+sSearchCol+",'-',''),len(replace("+sSearchCol+",'-',''))-charindex('"+sSeparator+"',replace("+sSearchCol+",'-',''),len('"+sDocIDPJ+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc like N'"+parClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%' AND Department like N'"+parDepartment+"%'"
      iDigits = 5    
    Case SIT_VTSS
	  sSeparator = "/"
	  sPrefix="IH"+Right(CStr(Year(Date)),2)+"_VTSS"+sSeparator'+Left(parSubClassParameter, NotNegative(InStr(parSubClassParameter, " ")-1))+sSeparator
	  sSQL="select ISNULL(Max(cast(right(replace("+sSearchCol+",'-',''),len(replace("+sSearchCol+",'-',''))-charindex('"+sSeparator+"',replace("+sSearchCol+",'-',''),len('"+sDocIDPJ+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc like N'"+parClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%' AND Department like N'"+parDepartment+"%'"
      iDigits = 5   
    Case SIT_STS
      sSeparator = "/"
      sPrefix="IH."+Right(CStr(Year(Date)),2)+sSeparator
      sSQL="select ISNULL(Max(case IsNumeric(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1))) when 1 then cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1))as int) else -1 end),0) as MaxDocID from Docs where  ClassDoc like N'"+parClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%' AND Department like N'"+parDepartment+"%'"
      iDigits = 3
    Case SIT_MIKRON
	  sSeparator = "-"
	  sPrefix="IH."+Right(CStr(Year(Date)),2)+"_MIK"+sSeparator'+Left(parSubClassParameter, NotNegative(InStr(parSubClassParameter, " ")-1))+sSeparator
	  sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc like N'"+parClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%' AND Department like N'"+parDepartment+"%'"
      iDigits = 3
    Case SIT_SIB
      sSeparator = "/"
      sPrefix="IH."+Right(CStr(Year(Date)),2)+"_SIB"+sSeparator
      sSQL="select ISNULL(Max(case IsNumeric(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1))) when 1 then cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1))as int) else -1 end),0) as MaxDocID from Docs where  ClassDoc like N'"+parClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%' AND Department like N'"+parDepartment+"%'"
      iDigits = 3
    Case "OTHER B.U."
    'для других бизнес направлений
      sSeparator = "-"
    Case SIT_SITRU ' DmGorsky
	  sSeparator = "-"
	  sPrefix="IH."+Right(CStr(Year(Date)),2)+"_SITRU/"+Left(parSubClassParameter, NotNegative(InStr(parSubClassParameter, " ")-1))+sSeparator
	  sSQL="select ISNULL(Max(cast(right("+sSearchCol+",len("+sSearchCol+")-charindex('"+sSeparator+"',"+sSearchCol+",len('"+sDocIDPJ+"')+1))as int)),0) as MaxDocID from Docs where ClassDoc like N'"+parClassDoc+"' AND "+sSearchCol+" like N'"+sDocIDPJ+sPrefix+"%' AND Department like N'"+parDepartment+"%'"
      iDigits = 3
  End Select

  GetNewDocIDForSluzhZapiski = GetNewDocID(sSQL, parIsProjectDocID & sPrefix, "", iDigits)
End Function

'Расчет номеров ТКП
Function GetNewDocIDForComOffers(parClassDoc, parDepartment)
  Dim iDigits, sSQL, sPrefix

  Select Case parDepartment
    Case SIT_SIB
      sPrefix = "CO_SIB-"
      sSQL="select ISNULL(MAX(cast(REPLACE(DocID,"+sUnicodeSymbol+"'"+sPrefix+"','') as int)),0) as MaxDocID from Docs where ClassDoc="+sUnicodeSymbol+"'"+parClassDoc+"' and DocID like "+sUnicodeSymbol+"'"+sPrefix+"%'"
      iDigits = 6
    Case "OTHER B.U."
    'для других бизнес направлений
	Case Else 'Сейчас для всех остальных одинаково
      sPrefix = "CO-"
      sSQL="select ISNULL(MAX(cast(REPLACE(DocID,"+sUnicodeSymbol+"'"+sPrefix+"','') as int)),0) as MaxDocID from Docs where ClassDoc="+sUnicodeSymbol+"'"+parClassDoc+"' and DocID like "+sUnicodeSymbol+"'"+sPrefix+"%'"
      iDigits = 6
  End Select

  GetNewDocIDForComOffers = GetNewDocID(sSQL, sPrefix, "", iDigits)
End Function

'Расчет номеров Протоколов
Function GetNewDocIDForProtocols(parClassDoc, parDepartment)
Dim iDigits, sSQL, sPrefix

  Select Case parDepartment
    Case SIT_SITRONICS
         sPrefix = "PR-"
	     If InStr(UCase(parClassDoc), UCase(SIT_PROTOCOLS_MC_EGRB)) > 0 Then
            sPrefix = sPrefix & "УК_ЭПРБ-"
         ElseIf InStr(UCase(parClassDoc), UCase(SIT_PROTOCOLS_IT_Committee)) > 0 Then
            sPrefix = sPrefix & "ИТ-"
	     ElseIf InStr(UCase(parClassDoc), UCase(SIT_PROTOCOLS_Management_Board)) > 0 Then
            sPrefix = sPrefix & "MAN-"
	     ElseIf InStr(UCase(parClassDoc), UCase(SIT_PROTOCOLS_Control_And_Auditing_Committee)) > 0 Then
            sPrefix = sPrefix & "КРК-"
	     ElseIf InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_PROTOCOLS_CPC)) > 0 Then
            sPrefix = sPrefix & "CPC-"
	     Else
            sPrefix = sPrefix & "OTH-"
         End If
         sSQL="select ISNULL(MAX(cast(REPLACE(DocID,"+sUnicodeSymbol+"'"+sPrefix+"','') as int)),0) as MaxDocID from Docs where ClassDoc="+sUnicodeSymbol+"'"+parClassDoc+"' and DocID like "+sUnicodeSymbol+"'"+sPrefix+"%'"
         iDigits = 6
    Case SIT_SIB
	     sPrefix = "PR_SIB-"
	     'Пока только один тип протоколов - прочие
         sPrefix = sPrefix & "OTH-"
         sSQL="select ISNULL(MAX(cast(REPLACE(DocID,"+sUnicodeSymbol+"'"+sPrefix+"','') as int)),0) as MaxDocID from Docs where ClassDoc="+sUnicodeSymbol+"'"+parClassDoc+"' and DocID like "+sUnicodeSymbol+"'"+sPrefix+"%'"
         iDigits = 6
    Case SIT_RTI                     
      Select case Request("UserFieldText3")
          case "Встречи"
            protokolType = "OTH"
          case "Правления"
            protokolType = "MAN"
          case "Совета Директоров"
            protokolType = "BD"
      End Select
   	  sPrefix = "PR-"+Right(CStr(Year(Date)),2)+"_RTI-"+protokolType+"/"
      sSQL="select ISNULL(MAX(cast(REPLACE(DocID,"+sUnicodeSymbol+"'"+sPrefix+"','') as int) + 0),0) as MaxDocID from Docs where ClassDoc="+sUnicodeSymbol+"'"+parClassDoc+"' and DocID like "+sUnicodeSymbol+"'"+sPrefix+"%'"
      iDigits = 6
    Case SIT_MIKRON
         Select case Request("UserFieldText3")
            case "Встречи"
                 protokolType = "OTH"
            case "Правления"
                 protokolType = "MAN"
            case "Совета Директоров"
                 protokolType = "BD"
         End Select
   	     sPrefix = "PRMIK"+Right(CStr(Year(Date)),2)+"-"+protokolType+"/"
         sSQL="select ISNULL(MAX(cast(REPLACE(DocID,"+sUnicodeSymbol+"'"+sPrefix+"','') as int)),0) as MaxDocID from Docs where ClassDoc="+sUnicodeSymbol+"'"+parClassDoc+"' and DocID like "+sUnicodeSymbol+"'"+sPrefix+"%'"
         iDigits = 6
    Case "OTHER B.U."
         'для других бизнес направлений
  End Select

  GetNewDocIDForProtocols = GetNewDocID(sSQL, sPrefix, "", iDigits)
End Function

'Расчет номеров Поручений
Function GetNewDocIDForZadachi(parClassDoc, parDepartment, parSubClassParameter, parParentDocID, parIsProjectDocID)
Dim iDigits, sSearchCol, sSQL, sPrefix, sSufix, sSeparator

  If parIsProjectDocID = "PJ-" Then
     sSearchCol = "DocIDadd"
     sDocIDPJ = "PJ-"
  Else
     sSearchCol = "DocID"
     sDocIDPJ = ""
  End If

  Select Case parDepartment
    Case SIT_SITRONICS
	     sSeparator = "_"

         sPrefix= "T_"
	     If UCase(Trim(parSubClassParameter)) = UCase("Поручения АФК") or UCase(Trim(parSubClassParameter)) = UCase("SISTEMA tasks") Then
	        sPrefix= "T_AFK_"
	     End If 

         sSufix = ""
	     If Trim(parParentDocID) <> "" then
	        sSufix = "(" & parParentDocID & ")_"
	     End If

         sSQL = "select ISNULL(Max(cast(Right("+sSearchCol+", charindex('"+sSeparator+"',reverse("+sSearchCol+"),0)-1) as int)),0) as MaxDocID from Docs where ClassDoc like N'"+parClassDoc+"' and "
	     If InStr(sPrefix, "AFK") > 0 Then
	        sSQL = sSQL+" CharIndex(N'"+sDocIDPJ+"T_AFK_', "+sSearchCol+") = 1"
	     Else
	        sSQL = sSQL+" CharIndex(N'"+sDocIDPJ+"T_', "+sSearchCol+") = 1 and CharIndex(N'"+sDocIDPJ+"T_AFK_', "+sSearchCol+") <> 1"
	     End If
         iDigits = 5
    Case SIT_RTI
	     sSeparator = "_"

	  sPrefix= "T_"
	  If UCase(Trim(parSubClassParameter)) = UCase("Поручения АФК") or UCase(Trim(parSubClassParameter)) = UCase("SISTEMA tasks") Then
	    sPrefix= "T_AFK_"
	  End If 

	  sSufix = ""
	  If Trim(parParentDocID) <> "" then
	    sSufix = "(" & parParentDocID & ")_"
	  End If

      sSQL = "select ISNULL(Max(cast(Right("+sSearchCol+", charindex('"+sSeparator+"',reverse("+sSearchCol+"),0)-1) as int)),0) as MaxDocID from Docs where ClassDoc like N'"+parClassDoc+"' and "
	  If InStr(sPrefix, "AFK") > 0 Then
	    sSQL = sSQL+" CharIndex(N'"+sDocIDPJ+"T_AFK_', "+sSearchCol+") = 1"
	  Else
	    sSQL = sSQL+" CharIndex(N'"+sDocIDPJ+"T_', "+sSearchCol+") = 1 and CharIndex(N'"+sDocIDPJ+"T_AFK_', "+sSearchCol+") <> 1"
	  End If
      iDigits = 5
    Case SIT_MINC
     addlogd "exor666 - сюда дошли"
	     sSeparator = "_"

	  sPrefix= "K_"

	  sSufix = ""
	  If Trim(parParentDocID) <> "" then
	    sSufix = "(" & parParentDocID & ")_"
	  End If

      sSQL = "select ISNULL(Max(cast(Right("+sSearchCol+", charindex('"+sSeparator+"',reverse("+sSearchCol+"),0)-1) as int)),0) as MaxDocID from Docs where ClassDoc like N'"+parClassDoc+"' and "
	  If InStr(sPrefix, "AFK") > 0 Then
	    sSQL = sSQL+" CharIndex(N'"+sDocIDPJ+"T_AFK_', "+sSearchCol+") = 1"
	  Else
	    sSQL = sSQL+" CharIndex(N'"+sDocIDPJ+"K_', "+sSearchCol+") = 1 and CharIndex(N'"+sDocIDPJ+"T_AFK_', "+sSearchCol+") <> 1"
	  End If
      iDigits = 5
    Case SIT_VTSS
     addlogd "exor666 - сюда дошли"
	     sSeparator = "_"

	  sPrefix= "T_VTSS_"

	  sSufix = ""
	  If Trim(parParentDocID) <> "" then
	    sSufix = "(" & parParentDocID & ")_"
	  End If

      sSQL = "select ISNULL(Max(cast(Right("+sSearchCol+", charindex('"+sSeparator+"',reverse("+sSearchCol+"),0)-1) as int)),0) as MaxDocID from Docs where ClassDoc like N'"+parClassDoc+"' and "
	  If InStr(sPrefix, "AFK") > 0 Then
	    sSQL = sSQL+" CharIndex(N'"+sDocIDPJ+"T_AFK_', "+sSearchCol+") = 1"
	  Else
	    sSQL = sSQL+" CharIndex(N'"+sDocIDPJ+"T_VTSS_', "+sSearchCol+") = 1 and CharIndex(N'"+sDocIDPJ+"T_AFK_', "+sSearchCol+") <> 1"
	  End If
      iDigits = 5

    Case SIT_STS
	     sSeparator = "_"

	     sPrefix= "T_"
	     If UCase(Trim(parSubClassParameter)) = UCase("Поручения АФК") or UCase(Trim(parSubClassParameter)) = UCase("SISTEMA tasks") Then
	        sPrefix= "T_AFK_"
	     End If 

	     sSufix = ""
	     If Trim(parParentDocID) <> "" then
	        sSufix = "(" & parParentDocID & ")_"
	     End If

	     If InStr(sPrefix,"AFK") > 0 Then
	        sSQL = "select ISNULL(Max(cast(Right("+sSearchCol+", charindex('"+sSeparator+"',reverse("+sSearchCol+"),0)-1) as int)),0) as MaxDocID from Docs where "+sSearchCol+" like N'"+sDocIDPJ+"T[_]AFK[_]%' and ClassDoc like N'"+parClassDoc+"'"
	     Else
	        sSQL = "select ISNULL(Max(cast(Right("+sSearchCol+", charindex('"+sSeparator+"',reverse("+sSearchCol+"),0)-1) as int)),0) as MaxDocID from Docs where "+sSearchCol+" like N'"+sDocIDPJ+"T[_]%' AND "+sSearchCol+" not like N'%AFK%' and ClassDoc like N'"+parClassDoc+"'"
	     End If
         iDigits = 5
    Case SIT_MIKRON
	     sSeparator = "_"

	     sPrefix= "T_"
	     If UCase(Trim(parSubClassParameter)) = UCase("Поручения АФК") or UCase(Trim(parSubClassParameter)) = UCase("SISTEMA tasks") Then
	        sPrefix= "T_AFK_"
	     End If 

	     sSufix = ""
	     If Trim(parParentDocID) <> "" then
	        sSufix = "(" & parParentDocID & ")_"
	     End If

         sSQL = "select ISNULL(Max(cast(Right("+sSearchCol+", charindex('"+sSeparator+"',reverse("+sSearchCol+"),0)-1) as int)),0) as MaxDocID from Docs where ClassDoc like N'"+parClassDoc+"' and "
	     If InStr(sPrefix, "AFK") > 0 Then
	        sSQL = sSQL+" CharIndex(N'"+sDocIDPJ+"T_AFK_', "+sSearchCol+") = 1"
	     Else
	        sSQL = sSQL+" CharIndex(N'"+sDocIDPJ+"T_', "+sSearchCol+") = 1 and CharIndex(N'"+sDocIDPJ+"T_AFK_', "+sSearchCol+") <> 1"
	     End If
         iDigits = 5
    Case SIT_SIB
	     sSeparator = "_"
	     sPrefix= "T_SIB_"

	     sSufix = ""
	     If Trim(parParentDocID) <> "" then
	        sSufix = "(" & parParentDocID & ")_"
	     End If

         sSQL = "select ISNULL(Max(cast(Right(" & sSearchCol & ", charindex('" & sSeparator & "',reverse(" & sSearchCol & "),0)-1) as int)),0) as MaxDocID from Docs where ClassDoc like N'" & parClassDoc & "' and "
         sSQL = sSQL & " CharIndex(N'" & sDocIDPJ & "T_SIB_', " & sSearchCol & ") = 1"
         iDigits = 5
    Case "OTHER B.U."
    'для других бизнес направлений
    Case SIT_SITRU ' DmGorsky
    ' Нумерация для Ситроникс ИТ
	     sSeparator = "_"      

	     sPrefix= "T_SITRU_"
	     If UCase(Trim(parSubClassParameter)) = UCase("Поручения АФК") or UCase(Trim(parSubClassParameter)) = UCase("SISTEMA tasks") Then
	        sPrefix= "T_AFK_SITRU_"
	     End If 

	     sSufix = ""
	     If Trim(parParentDocID) <> "" then
	        sSufix = "(" & parParentDocID & ")_"
	     End If

         sSQL = "select ISNULL(Max(cast(Right("+sSearchCol+", charindex('"+sSeparator+"',reverse("+sSearchCol+"),0)-1) as int)),0) as MaxDocID from Docs where ClassDoc like N'"+parClassDoc+"' and "
	     If InStr(sPrefix, "AFK") > 0 Then
	        sSQL = sSQL+" CharIndex(N'"+sDocIDPJ+"T_AFK_SITRU_', "+sSearchCol+") = 1"
	     Else
	        sSQL = sSQL+" CharIndex(N'"+sDocIDPJ+"T_SITRU_', "+sSearchCol+") = 1 and CharIndex(N'"+sDocIDPJ+"T_AFK_SITRU_', "+sSearchCol+") <> 1"
	     End If
         iDigits = 5
  End Select

  GetNewDocIDForZadachi = GetNewDocID(sSQL, parIsProjectDocID & sPrefix & sSufix, "", iDigits)
End Function

'Замена системной функции
Function MyLeadSymbolNVal(cPar, symbol, N)
  cPar = Trim(MyCStr(cPar))
  If Len(cPar) < N Then
    MyLeadSymbolNVal = String(N - Len(cPar), symbol) & cPar
  Else
    MyLeadSymbolNVal = cPar
  End If
End Function

'Запрос №1 - СИБ - end


'Запрос №43 - СТС - start
'Получить рейтинг контрагента (для справочника правил)
Function GetPartnerRating(ByVal parPartnerName)
  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  sSQL = "Select Rating from Partners where Name = " & sUnicodeSymbol & "'" & MakeSQLSafeSimple(parPartnerName) & "'"
  dsTemp.Open sSQL, Conn, 3, 1, &H1

  GetPartnerRating = MyCStr(dsTemp("Rating"))

  dsTemp.Close
End Function

'Проверить наличие у контрагента ИНН
'Если ИНН пустой или повторяется, то ERROR
'Используется Conn
Function CheckPartnerTaxID(ByVal parPartnerName)
   CheckPartnerTaxID = ""
   sSQL = "select * from Partners where Name = " & sUnicodeSymbol & "'" & MakeSQLSafeSimple(parPartnerName) & "'"
   Set dsTemp = Server.CreateObject("ADODB.Recordset")
   dsTemp.Open sSQL, Conn, 3, 1, &H1
   If not dsTemp.EOF Then
      If dsTemp("TaxID") <> "" and dsTemp("Country") <> "" Then
         CheckPartnerTaxID = dsTemp("TaxID")
         If (dsTemp("Country") = "Россия" or dsTemp("Country") = "РФ" or dsTemp("Country") = "RU" or dsTemp("Country") = "RUS") Then
            If InStr(dsTemp("TaxID"), dsTemp("Area")) <> 1 Then
               CheckPartnerTaxID = ""
            End If
         Else
            If InStr(dsTemp("TaxID"),dsTemp("Country")) <> 1 Then
               CheckPartnerTaxID = ""
            End If
         End If
      End If
   End If
   dsTemp.Close
   If CheckPartnerTaxID <> "" Then
      sSQL = "select * from Partners where TaxID = " & sUnicodeSymbol & "'" & MakeSQLSafeSimple(CheckPartnerTaxID) & "'"
      Set dsTemp1 = Server.CreateObject("ADODB.Recordset")
      dsTemp1.Open sSQL, Conn, 3, 1, &H1
      iDigits = 0
      Do While not dsTemp1.EOF
         dsTemp1.MoveNext
         iDigits = iDigits + 1
      Loop
      dsTemp1.Close
      If iDigits <> 1 Then
         CheckPartnerTaxID = ""
      End If
   End If
End Function
Function CorrectSum(parSum)
  CorrectSum = MyCStr(parSum)
  CorrectSum = Replace(CorrectSum, " ", "")
  CorrectSum = Replace(CorrectSum, ",", ".")
End Function

'Запрос №43 - СТС - end

'Запрос №46 - СТС - start
'Дополнить список проектов названиями и проверить правильность указания номеров (в случае ошибки возвращается список несуществующих проектов)
Function CheckAndEnhanceProjectList(ByRef parProjectCodeList, ByRef parProjectManagers, ByRef parConn)
Dim arProjects, i, iPos
Dim sProjectCodes

	arProjects = Split(parProjectCodeList, VbCrLf)
	sProjectCodes = "#"
	'Чистим список, удаляем возможные приписки названий проектов (через пробел или пробел+дефис)
	For i = 0 To UBound(arProjects)
		arProjects(i) = Trim(arProjects(i))
		iPos = InStr(arProjects(i), " ")
		If iPos > 0 Then
			arProjects(i) = Left(arProjects(i), iPos-1)
		End If
		sProjectCodes = sProjectCodes & arProjects(i) & "#"
	Next
	Set dsTemp = Server.CreateObject("ADODB.Recordset")
	dsTemp.CursorLocation = 3
	sSQL = "select * from ProjectList where CharIndex(N'#'+ProjectID+N'#', " & sUnicodeSymbol & "'" & MakeSQLSafeSimple(sProjectCodes) & "') > 0 order by CharIndex(N'#'+ProjectID+N'#', " & sUnicodeSymbol & "'" & MakeSQLSafeSimple(sProjectCodes) & "')"
	dsTemp.Open sSQL, parConn, 3, 1, &H1
	dsTemp.ActiveConnection = Nothing

	'Выбрасываем найденные номера проектов из строки с номерами, формируем список проектов "номер - название" и список менеджеров проектов
	parProjectCodeList = ""
	parProjectManagers = ""
	Do While not dsTemp.EOF
		sProjectCodes = Replace(sProjectCodes, dsTemp("ProjectID"), "")

		If parProjectCodeList <> "" Then
			parProjectCodeList = parProjectCodeList & VbCrLf
		End If
		parProjectCodeList = parProjectCodeList & dsTemp("ProjectID") & " - " & dsTemp("ProjectName")

		If parProjectManagers <> "" Then
			parProjectManagers = parProjectManagers & VbCrLf
		End If
		If not IsNull(dsTemp("ProjectManagerUser")) Then
			parProjectManagers = parProjectManagers & Trim(CStr(dsTemp("ProjectManagerUser")))
		End If

		dsTemp.MoveNext
	Loop
	dsTemp.Close
	Set dsTemp = Nothing

	sProjectCodes = Trim(Replace(sProjectCodes, "#", " "))
	CheckAndEnhanceProjectList = sProjectCodes = ""
	'Если оказались ненайденные номера проектов, в parProjectCodeList возвращаем отсутствующие номера (через запятую)
	If not CheckAndEnhanceProjectList Then
		parProjectCodeList = Replace(sProjectCodes, " ", ", ")
	End If
End Function

'Запрос №46 - СТС - end

'Запрос №44 - СТС - start
Function CheckDepartmentInList_SQL(ByVal sField, ByVal sDepartment)
  Dim arDeps, sDep, i
  
  CheckDepartmentInList_SQL = ""
  If sDepartment <> "" Then
    arDeps = Split(sDepartment, VAR_TreeFolderSeparator)

    sDep = ""
    For i = 0 to UBound(arDeps)
      If arDeps(i) = "" Then
        Exit Function
      End If

      If sDep <> "" Then
        sDep = sDep + VAR_TreeFolderSeparator
      End If
      sDep = sDep + arDeps(i)

      If CheckDepartmentInList_SQL <> "" Then
        CheckDepartmentInList_SQL = CheckDepartmentInList_SQL & "or "
      End If
      CheckDepartmentInList_SQL = CheckDepartmentInList_SQL & "CHARINDEX(N'<DEPARTMENTS: " & sDep & ">', " & sField+")>0 "
    Next
  End If
End Function
'Запрос №44 - СТС - end

'{ph - 20120326
Function GetNearestCostCenterCodeByCode(ByVal sCode)
  Set dsTemp = Server.CreateObject("ADODB.Recordset")
  sSQL = "Select Code from Departments where CharIndex(Name, (Select Name from Departments where Code = N'"+sCode+"')) > 0 and Statuses like "+sUnicodeSymbol+"'%#BL=Yes%' order by Name desc"
AddlogD "GetNearestCostCenterCodeByCode SQL: "+sSQL
  dsTemp.Open sSQL, Conn, 3, 1, &H1
  If dsTemp.EOF Then
    GetNearestCostCenterCodeByCode = ""
  Else
    GetNearestCostCenterCodeByCode = dsTemp("Code")
  End If
  dsTemp.Close
End Function
'ph - 20120326}


%>