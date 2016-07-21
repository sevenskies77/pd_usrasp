<%
'vnik_protocol
'Протоколы определенных комитетов могут создавать только секретари этих комитетов
If UCase(Request.ServerVariables("URL")) = UCase("/ChangeDoc.asp") Then
    If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_PROTOCOLS)) > 0 Then
        If Request("create") = "y" Then
           
            If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_PROTOCOLS_MC_EGRB)) > 0 Then
                If InStr(UCase(ReplaceRoleFromDir(SIT_SecretaryOfManagingCommitteeOnTheProgramEGRB,Session("Department"))),UCase(Session("UserID"))) > 0 Then
                        'можно создавать
                Else
                        Session("Message") = AddNewLineToMessage(Session("Message"), "Только секретарь данного комитета может создавать протокол!!!")
                        Response.Redirect GetURL("listdoc.asp", "?ClassDoc=",Session("CurrentClassDoc"))                            
                End If    
            Else 
                If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_PROTOCOLS_IT_Committee)) > 0 Then
                    If InStr(UCase(ReplaceRoleFromDir(SIT_SecretaryOfCommitteeOnIT,Session("Department"))),UCase(Session("UserID"))) > 0 Then
                        'можно создавать
                    Else
                        Session("Message") = AddNewLineToMessage(Session("Message"), "Только секретарь данного комитета может создавать протокол!!!")
                        Response.Redirect GetURL("listdoc.asp", "?ClassDoc=",Session("CurrentClassDoc"))                            
                    End If      
                Else
                    If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_PROTOCOLS_Management_Board)) > 0 Then
                        If InStr(UCase(ReplaceRoleFromDir(SIT_SecretaryOfBoard,Session("Department"))),UCase(Session("UserID"))) > 0 Then
                            'можно создавать
                        Else
                            Session("Message") = AddNewLineToMessage(Session("Message"), "Только секретарь данного комитета может создавать протокол!!!")
                            Response.Redirect GetURL("listdoc.asp", "?ClassDoc=",Session("CurrentClassDoc"))                            
                        End If    
                    Else
                        If InStr(UCase(Session("CurrentClassDoc")),UCase(SIT_PROTOCOLS_Control_And_Auditing_Committee)) > 0 Then
                            If InStr(UCase(ReplaceRoleFromDir(SIT_SecretaryOfCommitteeOnControlAndAuditing,Session("Department"))),UCase(Session("UserID"))) > 0 Then
                                'можно создавать
                            Else
                                Session("Message") = AddNewLineToMessage(Session("Message"), "Только секретарь данного комитета может создавать протокол!!!")
                                Response.Redirect GetURL("listdoc.asp", "?ClassDoc=",Session("CurrentClassDoc"))                            
                            End If 
                        Else
                           'Протокол Встреч доступен всем
                        End If
                    End If
                End If
            End If
        End If
    End If
End If
'vnik_protocol

'vnik_rasp_norm_doc
'Категория Нормативные документы закрыта для создания новых документов с 01.07.2010
If UCase(Request.ServerVariables("URL")) = UCase("/ChangeDoc.asp") Then
    If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_NORM_DOCS)) > 0 Then
        If Request("create") = "y" Then
            Session("Message") = AddNewLineToMessage(Session("Message"), SIT_NORM_DOCS_CLOSED)
            Response.Redirect GetURL("listdoc.asp", "?ClassDoc=",Session("CurrentClassDoc"))    
        End If
    End If
End If
'vnik_rasp_norm_doc

'vnik_protocolsCPC
'Протоколы ЦЗК может делать только пользователь входящий в роль СекретарьЦЗК
If UCase(Request.ServerVariables("URL")) = UCase("/ChangeDoc.asp") Then
    If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_PROTOCOLS_CPC)) > 0 Then
        If Request("create") = "y" Then
            If InStr(UCase(ReplaceRoleFromDir(SIT_SecretaryCPC,Session("Department"))),UCase(Session("UserID"))) = 0 Then
                'AddLogD "vnik444 " + Trim(Request("DocID"))
                'AddLogD "vnik444 " + Trim(Request("DocIDParent"))
                Session("Message") = AddNewLineToMessage(Session("Message"), RTI_PROTOCOL_WARNING2)
                Response.Redirect GetURL("listdoc.asp", "?ClassDoc=",Session("CurrentClassDoc"))    
            End If
        End If
    End If
End If
'vnik_protocolsCPC

'rti_purchase_order
AddLogD "456"
If UCase(Request.ServerVariables("URL")) = UCase("/Visa.asp") Then
    If InStr(UCase(Session("CurrentClassDoc")), UCase(RTI_PURCHASE_ORDER)) = 1 Then
    'vnik - purchase_order_rti
        Set dsDocVNIK = Server.CreateObject("ADODB.Recordset")
        sSQL = "select Field2 from UserDirValues where UDKeyField='150' and Field1 like N'%#Руководитель по бюджетному контролю%'"
        dsDocVNIK.CursorLocation = 3
        dsDocVNIK.Open sSQL, Conn, 3, 1, &H1
        If not dsDocVNIK.EOF Then
            RoleBudgContrValue = Trim(dsDocVNIK("Field2").Value)
        Else
            RoleBudgContrValue = ""
        End If
        dsDocVNIK.Close
        VNIK_Current_User = "<"+Session("UserID")+">"
    'vnik - purchase_order_rti
            
    If InStr(RoleBudgContrValue, VNIK_Current_User)>0 Then
        If Trim(Request("Cost_item")) = "" and Trim(Request("r")) <> "y" Then
            Session("Message") = AddNewLineToMessage(Session("Message"), "ДОКУМЕНТ НЕ СОГЛАСОВАН!!! Не указана статья расходов")
            Response.Redirect GetURL("showdoc.asp", "?docid=", Request("DocID"))
        End If
     End If   
   End If
End If
'rti_purchase_order

'rti_bsap
If UCase(Request.ServerVariables("URL")) = UCase("/ChangeDoc.asp") Then
    If InStr(UCase(Session("CurrentClassDoc")), UCase(RTI_BSAP)) = 1 Then
        If Request("create") = "y" Then
        'Session("Message") = AddNewLineToMessage(Session("Message"), "123:" + Trim(Request("DocID")))
        'Session("Message") = AddNewLineToMessage(Session("Message"), "456:" + Trim(Request("DocIDParent")))
        'Session("Message") = AddNewLineToMessage(Session("Message"), Request("ClassDoc"))
            If (InStr(UCase(Request("ClassDoc")), UCase(RTI_PURCHASE_ORDER)) <> 1) Then
                If (Trim(Request("DocIDParent")) = "" and InStr(UCase(Request("ClassDoc")), UCase(RTI_BSAP)) = 1) Then
                    Session("Message") = AddNewLineToMessage(Session("Message"), Trim(RTI_BSAP_WARNING1))
                    Response.Redirect GetURL("listdoc.asp", "?ClassDoc=",Session("CurrentClassDoc"))  
                End If            
            Else
                'Получим документ основание
                Set vnikdsTemp1 = Server.CreateObject("ADODB.Recordset")
                sSQL = "SELECT * FROM Docs where DocID = " +sUnicodeSymbol+"'"+Request("DocID")+"'" + "and ClassDoc = N'Закупки РТИ/Заявка на закупку' and StatusDevelopment = 4"
                AddLogD "vnik444 " + Trim(sSQL)
                vnikdsTemp1.Open sSQL, Conn, 3, 1, &H1             
	            If vnikdsTemp1.EOF Then                
                            Session("Message") = AddNewLineToMessage(Session("Message"), Trim(RTI_BSAP_WARNING1))
                            Response.Redirect GetURL("listdoc.asp", "?ClassDoc=",Session("CurrentClassDoc")) 
                  Else If vnikdsTemp1("UserFieldMoney1") >= 354000 or vnikdsTemp1("UserFieldMoney1") <= 59000 Then
                            Session("Message") = AddNewLineToMessage(Session("Message"), Trim(RTI_BSAP_WARNING1))
                            Response.Redirect GetURL("listdoc.asp", "?ClassDoc=",Session("CurrentClassDoc"))  
                  End If 	            
                End If               
                vnikdsTemp1.Close        
            End If
        End If
    End If
End If
'rti_bsap

'rti_payment_order
If UCase(Request.ServerVariables("URL")) = UCase("/ChangeDoc.asp") Then
    If InStr(UCase(Session("CurrentClassDoc")), UCase(RTI_PAYMENT_ORDER)) = 1 Then
        If Request("create") = "y" Then
            If not (Request("ClassDocDependant") = "" and Request("empty") = "") Then
                'Session("Message") = AddNewLineToMessage(Session("Message"), "123:" + Trim(Request("DocID")))
                'Session("Message") = AddNewLineToMessage(Session("Message"), "456:" + Trim(Request("DocIDParent")) + "456")
                'Session("Message") = AddNewLineToMessage(Session("Message"), "ClassDoc = " + Request("ClassDoc"))
                'Session("Message") = AddNewLineToMessage(Session("Message"), "ClassDocDependant = " + Request("ClassDocDependant"))
                'Session("Message") = AddNewLineToMessage(Session("Message"), "empty = " + Request("empty"))
             If (InStr(UCase(Request("ClassDoc")), UCase(RTI_PURCHASE_ORDER)) <> 1) Then
                If (Trim(Request("DocIDParent")) = "" and InStr(UCase(Request("ClassDoc")), UCase(RTI_PAYMENT_ORDER)) = 1) Then                   
                    Session("Message") = AddNewLineToMessage(Session("Message"), Trim(RTI_PAYMENT_ORDER_WARNING1))
                    Response.Redirect GetURL("listdoc.asp", "?ClassDoc=",Session("CurrentClassDoc"))  
                End If  
            Else
                'Получим документ основание
                Set vnikdsTemp1 = Server.CreateObject("ADODB.Recordset")
                sSQL = "SELECT * FROM Docs where DocID = " +sUnicodeSymbol+"'"+Request("DocID")+"'" + "and ClassDoc = N'Закупки РТИ/Заявка на закупку' and StatusDevelopment = 4"
                AddLogD "vnik444 " + Trim(sSQL)
                vnikdsTemp1.Open sSQL, Conn, 3, 1, &H1   
                'addlogd "eXor999 " + vnikdsTemp1("UserFieldText4") + " " + UCase(Session("UserId")) + UCase(vnikdsTemp1("ListToView"))          
	            If vnikdsTemp1.EOF Then                
                            Session("Message") = AddNewLineToMessage(Session("Message"), Trim(RTI_PAYMENT_ORDER_WARNING1))
                            Response.Redirect GetURL("listdoc.asp", "?ClassDoc=",Session("CurrentClassDoc"))
                Else If vnikdsTemp1("UserFieldText4") = "13072 Канцелярские и хозяйственные товары, офисные принадлежности" and InStr(UCase(vnikdsTemp1("ListToView")), UCase(Session("UserId"))) < 1 Then
                            Session("Message") = AddNewLineToMessage(Session("Message"), Trim(RTI_PAYMENT_ORDER_WARNING3))
                            Response.Redirect GetURL("listdoc.asp", "?ClassDoc=",Session("CurrentClassDoc"))  
                 End If
                 End If 	                         
                vnikdsTemp1.Close        
            End If
          End If
        End If
    End If
End If
'rti_payment_order

'rti_входящие
'Входящие документы РТИ может создавать только пользователь с ролью Регистратор
If UCase(Request.ServerVariables("URL")) = UCase("/ChangeDoc.asp") and not IsAdmin() Then
    If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_VHODYASCHIE)) = 1 and InStr(UCase(Session("Department")), UCase(SIT_RTI)) = 1 Then
        If Request("create") = "y" Then
        'sRoleList = GetRolesList(Request("Department"), Session("UserID"), Request("BusinessUnit"))
        'Session("Message") = AddNewLineToMessage(Session("Message"), ReplaceRolesInList(SIT_Registrar, sRoleList))
            If InStr(UCase(ReplaceRoleFromDir(SIT_Registrar,SIT_RTI)),UCase(Session("UserID"))) = 0 Then
                Session("Message") = AddNewLineToMessage(Session("Message"), RTI_SIT_VHODYASCHIE_RESTRICT)
                Response.Redirect GetURL("listdoc.asp", "?ClassDoc=",Session("CurrentClassDoc"))    
            End If
        End If
    End If
End If
'rti_входящие

'rti_protocol
'Протоколы ЦЗК может делать только пользователь входящий в роль СекретарьЦЗК
If UCase(Request.ServerVariables("URL")) = UCase("/ChangeDoc.asp") Then
    If InStr(UCase(Session("CurrentClassDoc")), UCase(RTI_PROTOCOL)) = 1 Then
        If Request("create") = "y" Then
            If InStr(UCase(ReplaceRoleFromDir(RTI_SecretaryOfCPC,SIT_RTI)),UCase(Session("UserID"))) = 0 Then
                Session("Message") = AddNewLineToMessage(Session("Message"), SIT_PROTOCOLS_CPC_RESTRICT)
                Response.Redirect GetURL("listdoc.asp", "?ClassDoc=",Session("CurrentClassDoc"))    
            End If
        End If
    End If
End If
'rti_protocol

'rti_protocol
If UCase(Request.ServerVariables("URL")) = UCase("/ChangeDoc.asp") Then
    If InStr(UCase(Session("CurrentClassDoc")), UCase(RTI_PROTOCOL)) = 1 Then
        If Request("create") = "y" Then
            If (InStr(UCase(Request("ClassDoc")), UCase(RTI_PURCHASE_ORDER)) <> 1) Then
                If (Trim(Request("DocIDParent")) = "" and InStr(UCase(Request("ClassDoc")), UCase(RTI_PROTOCOL)) = 1) Then
                    Session("Message") = AddNewLineToMessage(Session("Message"), Trim(RTI_PROTOCOL_WARNING1))
                    Response.Redirect GetURL("listdoc.asp", "?ClassDoc=",Session("CurrentClassDoc"))  
                End If            
            Else
                'Получим документ основание
                Set vnikdsTemp1 = Server.CreateObject("ADODB.Recordset")
                sSQL = "SELECT * FROM Docs where DocID = " +sUnicodeSymbol+"'"+Request("DocID")+"'" + "and ClassDoc = N'Закупки РТИ/Заявка на закупку' and StatusDevelopment = 4"
                AddLogD "vnik444 " + Trim(sSQL)
                vnikdsTemp1.Open sSQL, Conn, 3, 1, &H1             
	            If vnikdsTemp1.EOF Then                
                            Session("Message") = AddNewLineToMessage(Session("Message"), RTI_PROTOCOL_WARNING1)
                            Response.Redirect GetURL("listdoc.asp", "?ClassDoc=",Session("CurrentClassDoc"))  	            
                  Else If vnikdsTemp1("UserFieldMoney1") < 354000 Then
                            Session("Message") = AddNewLineToMessage(Session("Message"), RTI_PROTOCOL_WARNING1)
                            Response.Redirect GetURL("listdoc.asp", "?ClassDoc=",Session("CurrentClassDoc"))  
                 End If
                End If               
                vnikdsTemp1.Close        
            End If
        End If
    End If
End If
'rti_protocol

'rti_contract
'Договоры РТИ можно делать только на основании утвержденной заявки на закупку РТИ, либо сам по себе, либо на основе утвержденного БСАП
If UCase(Request.ServerVariables("URL")) = UCase("/ChangeDoc.asp") Then
    If InStr(UCase(Session("CurrentClassDoc")), UCase(RTI_CONTRACT)) = 1 Then
        If Request("create") = "y" Then
            'If InStr(UCase(Request("ClassDoc")), UCase(RTI_PURCHASE_ORDER)) <> 1 Then
            '    If Trim(Request("DocIDParent")) = "" and InStr(UCase(Request("ClassDoc")), UCase(RTI_CONTRACT)) <> 1  Then
            '        Session("Message") = AddNewLineToMessage(Session("Message"), Trim(RTI_CONTRACT_WARNING1))
            '        Response.Redirect GetURL("listdoc.asp", "?ClassDoc=",Session("CurrentClassDoc"))  
            '    End If  
            'Else
                'Если договор создается как подчиненный документ - получим документ основание  - договор можно создать на основе утвержденной заявки на закупку РТИ или утвержденного БСАП
             if InStr(UCase(Request("ClassDoc")), UCase(RTI_CONTRACT)) <> 1 Then
                Set vnikdsTemp1 = Server.CreateObject("ADODB.Recordset")
                sSQL = "SELECT * FROM Docs where DocID = " +sUnicodeSymbol+"'"+Request("DocID")+"'" + "and ((ClassDoc = N'Закупки РТИ/Заявка на закупку' or ClassDoc = N'Закупки РТИ/БСАП') and StatusDevelopment = 4)"
                vnikdsTemp1.Open sSQL, Conn, 3, 1, &H1             
	            If vnikdsTemp1.EOF Then                
                            Session("Message") = AddNewLineToMessage(Session("Message"), RTI_CONTRACT_WARNING1)
                            Response.Redirect GetURL("listdoc.asp", "?ClassDoc=",Session("CurrentClassDoc"))  	            
                End If
                vnikdsTemp1.Close        
            End If
        End If
    End If
End If
'rti_contract

'делегирование документов РТИ
If UCase(Request.ServerVariables("URL")) = UCase("/DelegateVisaToUser.asp") Then
      bDocID = Request("DocID")
      bNOUserID = "<" + Session("UserID") + ">; "  

	  bTypeNVarChar = 202 ' adVarWChar
	  bDirectionInput = 1 ' adParamInput	  	
      Set bUpdateCommand = Server.CreateObject("ADODB.Command")
	  bUpdateCommand.ActiveConnection = Conn
	  bUpdateCommand.Prepared = True
	  bUpdateCommand.CommandText = "UPDATE Docs Set AdditionalUsers = AdditionalUsers + ? WHERE DocID = ?"	  
	      
	  Set bParamNOUserID = bUpdateCommand.CreateParameter("@nouserid", bTypeNVarChar, bDirectionInput, 1024, bNOUserID)
	  bUpdateCommand.Parameters.Append bParamNOUserID
	   
	  Set bParamDocId = bUpdateCommand.CreateParameter("@docId", bTypeNVarChar, bDirectionInput, 128, bDocID)
	  bUpdateCommand.Parameters.Append bParamDocId
 
	  bUpdateCommand.Execute
	  Set bUpdateCommand = Nothing
     
End If
'делегирование документов РТИ

''**************************************************************************************************
''**                                   MIKRON ЗАКУПКИ (start)                                     **
''** AMW - 159 - это код справочника с ролями МИКРОНА                                                                  **
''**************************************************************************************************
' <!--#INCLUDE FILE="DBUpdateBefore_MIKRON.asp" -->
'amw 25-10-2013 (start)
' Не позволять активировать документ, если к нему не приложены файлы
'If UCase(Request.ServerVariables("URL")) = UCase("/MakeActive.asp") Then 'Страница активации документа
'   If Request("Active") <> "" Then 'Активация
'      If InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_RL_PROTOCOL)) <> 0 or _
'         InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_PROTOCOL)) <> 0 and _
'          InStr(UCase(Request("DocIDParent")), UCase("POM-")) = 1) or _
'         InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_CONTRACT)) <> 0 Then 'Ограничение по категории
'         Set dsTemp1 = Server.CreateObject("ADODB.Recordset") 'Создаем рекордсет
'         'Проверяем только основные версии файлов
'         sSQL = "select * from Comments where DocID = " & sUnicodeSymbol & "'" & MakeSQLSafeSimple(Request("DocID")) & "' and CommentType = 'FILE' and (Subject is NULL or Subject <> 'ESIGNATURE') and Amount = 0"
'         dsTemp1.CursorLocation = 3
'         dsTemp1.Open sSQL, Conn, 3, 1, &H1
'         dsTemp1.ActiveConnection = Nothing
'         bError = False
'         If dsTemp1.EOF Then
'            Session("Message") = Session("Message") & VbCrLf & "<font color=red>ОШИБКА!</font> Нет приложенных файлов"
'            bError = True 'Устанавливаем флаг ошибки
'         End If
'         dsTemp1.Close 'Закрываем рекордсет
'         If bError Then
'            'Делаем редирект на страницу просмотра документа , чтобы предотвратить запись в БД
'            Response.Redirect GetURL("showdoc.asp", "?docid=", Request("DocID")) & "&template=" & Trim(Request("template"))
'         End If
'      End If
'   End If
'End If
'amw 25-10-2013 (end)
   
'amw 25-11-2014 (start)
'Не позволять активировать документ, в котором указан контрагент, в карточке которого не
'заполнены необходимые поля: ИНН, ОГРН
If UCase(Request.ServerVariables("URL")) = UCase("/MakeActive.asp") Then 'Страница активации документа
   If Request("Active") <> "" Then 'Активация
      If InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_BSAP)) = 1 or _
         InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_RL_MEMO)) = 1 or _
         InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_PROTOCOL)) = 1 or _
         InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_RL_PROTOCOL)) = 1 or _
         InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_CONTRACT)) = 1 or _
         InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_S_CONTRACT)) = 1 or _
         InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_ADD_CONTRACT)) = 1 or _
         InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_OLD_CONTRACT)) = 1 or _
         InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_NDA_CONTRACT)) = 1 or _
         InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_EXPADD_CONTRACT)) = 1 or _
         InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_EXPORT_CONTRACT)) = 1 Then  'Ограничение по категории

         bError = False
         Set dsTemp = Server.CreateObject("ADODB.Recordset") 'Создаем рекордсет
         sSQL = "select PartnerName from Docs where DocID = N'" & MakeSQLSafeSimple(Trim(Request("DocID"))) & "'"
         dsTemp.Open sSQL, Conn, 3, 1, &H1
         If dsTemp.EOF Then
            sparPartnerName = Request("DocID")
            bError = True 'Устанавливаем флаг ошибки
         Else
            sparPartnerName = dsTemp("PartnerName")
         End If
         dsTemp.Close 'Закрываем рекордсет
         
         If not bError and CheckPartnerTaxID(sparPartnerName) = "" Then
            bError = True 'Устанавливаем флаг ошибки
         End If

         If bError Then
            Session("Message") = "Реквизиты контрагента (ИНН, страна, регион) отсуствуют или неправильные"
            'Делаем редирект на страницу просмотра документа , чтобы предотвратить запись в БД
            Response.Redirect GetURL("showdoc.asp", "?docid=", Request("DocID")) & "&template=" & Trim(Request("template"))
         End If

      End If
   End If
End If
'amw 25-11-2014 (end)

'mikron_purchase_order
If UCase(Request.ServerVariables("URL")) = UCase("/Visa.asp") Then
   If InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_PURCHASE_ORDER)) = 1 Then
      Set dsDocVNIK = Server.CreateObject("ADODB.Recordset")
      sSQL = "select Field2 from UserDirValues where UDKeyField='159' and Field1 like N'%#Руководитель по бюджетному контролю%'"
      dsDocVNIK.CursorLocation = 3
      dsDocVNIK.Open sSQL, Conn, 3, 1, &H1
      If not dsDocVNIK.EOF Then
         RoleBudgContrValue = Trim(dsDocVNIK("Field2").Value)
      Else
         RoleBudgContrValue = ""
      End If
      dsDocVNIK.Close
      
      VNIK_Current_User = "<"+Session("UserID")+">"
      If InStr(RoleBudgContrValue, VNIK_Current_User)>0 Then
         If Trim(Request("Cost_item")) = "" and Trim(Request("r")) <> "y" Then
            Session("Message") = AddNewLineToMessage(Session("Message"), "ДОКУМЕНТ НЕ СОГЛАСОВАН!!! Не указана статья расходов")
            Response.Redirect GetURL("showdoc.asp", "?docid=", Request("DocID"))
         End If
      End If   
   End If
End If
'mikron_purchase_order

'mikron_BSAP
If UCase(Request.ServerVariables("URL")) = UCase("/ChangeDoc.asp") Then
   If InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_BSAP)) = 1 Then
      If Request("create") = "y" Then
         bError = False
         If (InStr(UCase(Request("ClassDoc")), UCase(MIKRON_PURCHASE_ORDER)) <> 1) Then
            If (Trim(Request("DocIDParent")) = "" and InStr(UCase(Request("ClassDoc")), UCase(MIKRON_BSAP)) = 1) Then
               Session("Message") = AddNewLineToMessage(Session("Message"), Trim(MIKRON_BSAP_WARNING1))
               bError = True
            End If            
         Else
            'Получим документ основание
            Set vnikdsTemp1 = Server.CreateObject("ADODB.Recordset")
            sSQL = "SELECT * FROM Docs where DocID = "+sUnicodeSymbol+"'"+Request("DocID")+"' and ClassDoc=" + _
                   sUnicodeSymbol+"'"+MIKRON_PURCHASE_ORDER+"' and StatusDevelopment = 4"
            vnikdsTemp1.Open sSQL, Conn, 3, 1, &H1
            If vnikdsTemp1.EOF Then
               Session("Message") = AddNewLineToMessage(Session("Message"), Trim(MIKRON_BSAP_WARNING1) + " " + Trim(MIKRON_DOCNOTFOUND_WARNING1))
               bError = True
            Else If vnikdsTemp1("UserFieldMoney1") >= 500000 or vnikdsTemp1("UserFieldMoney1") <= 50000 Then
                    Session("Message") = AddNewLineToMessage(Session("Message"), Trim(MIKRON_BSAP_WARNING1) + " " + Trim(MIKRON_WRONGSUM_WARNING1))
                    bError = True
                 End If 	            
            End If               
            vnikdsTemp1.Close        
         End If
         If bError Then
            bError = False
            Response.Redirect GetURL("listdoc.asp", "?ClassDoc=",Session("CurrentClassDoc"))
         End If
      End If
   End If
End If
'mikron_BSAP

'mikron_payment_order
If UCase(Request.ServerVariables("URL")) = UCase("/ChangeDoc.asp") Then
   If InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_PAYMENT_ORDER)) = 1 Then
      If Request("create") = "y" Then
         bError = False
         If (InStr(UCase(Request("ClassDoc")),UCase(MIKRON_PURCHASE_ORDER)) <> 1) Then
            If (Trim(Request("DocIDParent")) = "" and InStr(UCase(Request("ClassDoc")), UCase(MIKRON_PAYMENT_ORDER)) = 1) Then
               Session("Message") = AddNewLineToMessage(Session("Message"), Trim(MIKRON_PAYMENT_ORDER_WARNING1))
               bError = True
            End If  
         Else
            'Получим документ основание
            Set vnikdsTemp1 = Server.CreateObject("ADODB.Recordset")
            sSQL = "SELECT * FROM Docs where DocID = "+sUnicodeSymbol+"'"+Request("DocID")+"' and ClassDoc=" + _
                   sUnicodeSymbol+"'"+MIKRON_PURCHASE_ORDER+"' and StatusDevelopment = 4"
            vnikdsTemp1.Open sSQL, Conn, 3, 1, &H1             
            If vnikdsTemp1.EOF Then                
               Session("Message") = AddNewLineToMessage(Session("Message"), Trim(MIKRON_PAYMENT_ORDER_WARNING1))
               bError = True
            Else If vnikdsTemp1("UserFieldMoney1") > 59000 Then
                    Session("Message") = AddNewLineToMessage(Session("Message"), Trim(MIKRON_PAYMENT_ORDER_WARNING2))
                    bError = True
                 End If 	            
            End If               
            vnikdsTemp1.Close        
         End If
         If bError Then
            bError = False
            Response.Redirect GetURL("listdoc.asp", "?ClassDoc=",Session("CurrentClassDoc"))
         End If
      End If
   End If
End If
'mikron_payment_order

'mikron_protocol CP
If UCase(Request.ServerVariables("URL")) = UCase("/ChangeDoc.asp") Then
   If InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_PROTOCOL)) = 1 Then
      If Request("create") = "y" Then
         bError = False
         'Протоколы ЗК может делать только пользователь входящий в роль Секретарь ЗК
         If InStr(UCase(ReplaceRoleFromDir(MIKRON_SecretaryOfPC,SIT_MIKRON)),UCase(Session("UserID"))) = 0 Then
            Session("Message") = AddNewLineToMessage(Session("Message"), MIKRON_PROTOCOL_WARNING2)
            Response.Redirect GetURL("listdoc.asp", "?ClassDoc=",Session("CurrentClassDoc"))
         End If
         'Протокол ЗК на основе Заявки на закупку. Форма проведения ОЧНАЯ
         If (InStr(UCase(Request("ClassDoc")), UCase(MIKRON_RL_PROTOCOL)) <> 1) Then
            If (InStr(UCase(Request("ClassDoc")), UCase(MIKRON_PURCHASE_ORDER)) <> 1) Then
               If (Trim(Request("DocIDParent")) = "" and InStr(UCase(Request("ClassDoc")), UCase(MIKRON_PROTOCOL)) = 1) Then
                  Session("Message") = AddNewLineToMessage(Session("Message"), Trim(MIKRON_PROTOCOL_WARNING1))
                  bError = True
               End If
            Else
               'Получим документ основание
               Set vnikdsTemp1 = Server.CreateObject("ADODB.Recordset")
               sSQL = "SELECT * FROM Docs where DocID = "+sUnicodeSymbol+"'"+Request("DocID")+"' and ClassDoc=" + _
                      sUnicodeSymbol+"'"+MIKRON_PURCHASE_ORDER+"' and StatusDevelopment = 4"
               vnikdsTemp1.Open sSQL, Conn, 3, 1, &H1
               If vnikdsTemp1.EOF Then
                  Session("Message") = AddNewLineToMessage(Session("Message"), MIKRON_PROTOCOL_WARNING1)
                  bError = True
'amw (start) 17-10-2013 К проверке на 500.000, добавлено исключение "0" - для особых договоров
               Else If vnikdsTemp1("UserFieldMoney1") < 500000 and vnikdsTemp1("UserFieldMoney1") <> 0 Then
                       Session("Message") = AddNewLineToMessage(Session("Message"), MIKRON_PROTOCOL_WARNING1)
                       bError = True
                    End If
'amw (start) 17-10-2013 К проверке на 500.000, добавлено исключение "0" - для особых договоров
               End If               
               vnikdsTemp1.Close        
            End If
         'Протокол ЗК на основе опросного листа. Форма проведения ЗАОЧНАЯ
         Else 'Родительский документ - Опросный лист.
            'Получим документ основание
            Set vnikdsTemp1 = Server.CreateObject("ADODB.Recordset")
            sSQL = "SELECT * FROM Docs where DocID = "+sUnicodeSymbol+"'"+Request("DocID")+"' and ClassDoc=" + _
                   sUnicodeSymbol+"'"+MIKRON_RL_PROTOCOL+"' and StatusDevelopment = 4"
            vnikdsTemp1.Open sSQL, Conn, 3, 1, &H1
            If vnikdsTemp1.EOF Then
               Session("Message") = AddNewLineToMessage(Session("Message"), MIKRON_PROTOCOL_WARNING3)
               bError = True
'amw (start) 17-10-2013 Сумму не проверяем, т.к. она проверялась при создании опросного листа
            End If               
            vnikdsTemp1.Close        
         End If
         If bError Then
            bError = False
            Response.Redirect GetURL("listdoc.asp", "?ClassDoc=",Session("CurrentClassDoc"))
         End If
      End If
   End If
End If
'mikron_protocol CP

'mikron RL for protocol CP
If UCase(Request.ServerVariables("URL")) = UCase("/ChangeDoc.asp") Then
   If InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_RL_PROTOCOL)) = 1 Then
      If Request("create") = "y" Then
         bError = False
         If (InStr(UCase(Request("ClassDoc")), UCase(MIKRON_PURCHASE_ORDER)) <> 1) Then
            If (Trim(Request("DocIDParent")) = "" and InStr(UCase(Request("ClassDoc")), UCase(MIKRON_RL_PROTOCOL)) = 1) Then
               Session("Message") = AddNewLineToMessage(Session("Message"), Trim(MIKRON_PROTOCOL_WARNING1))
               bError = True
            End If
         Else
            'Получим документ основание
            Set vnikdsTemp1 = Server.CreateObject("ADODB.Recordset")
            sSQL = "SELECT * FROM Docs where DocID = "+sUnicodeSymbol+"'"+Request("DocID")+"' and ClassDoc=" + _
                   sUnicodeSymbol+"'"+MIKRON_PURCHASE_ORDER+"' and StatusDevelopment = 4"
            vnikdsTemp1.Open sSQL, Conn, 3, 1, &H1
	        If vnikdsTemp1.EOF Then
               Session("Message") = AddNewLineToMessage(Session("Message"), MIKRON_PROTOCOL_WARNING1)
               bError = True
'amw (start) 17-10-2013 К проверке на 500.000, добавлено исключение "0" - для особых договоров
            Else If vnikdsTemp1("UserFieldMoney1") < 500000 and vnikdsTemp1("UserFieldMoney1") <> 0 Then
                    Session("Message") = AddNewLineToMessage(Session("Message"), MIKRON_PROTOCOL_WARNING1)
                    bError = True
                 End If
'amw (start) 17-10-2013 К проверке на 500.000, добавлено исключение "0" - для особых договоров
            End If               
            vnikdsTemp1.Close        
         End If
         If bError Then
            bError = False
            Response.Redirect GetURL("listdoc.asp", "?ClassDoc=",Session("CurrentClassDoc"))
         End If
      End If
   End If
End If
'mikron RL for protocol CP

'mikron_contract
'Договоры МИКРОН можно делать только на основании утвержденного БСАП или Протокола ЗК
If UCase(Request.ServerVariables("URL")) = UCase("/ChangeDoc.asp") Then
   If InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_CONTRACT)) = 1 Then
      If Request("create") = "y" Then
         bError = False
         If (InStr(UCase(Request("ClassDoc")), UCase(MIKRON_BSAP)) <> 1) or (InStr(UCase(Request("ClassDoc")), UCase(MIKRON_PROTOCOL)) <> 1) Then
            If (Trim(Request("DocIDParent")) = "" and InStr(UCase(Request("ClassDoc")), UCase(MIKRON_CONTRACT)) = 1) Then
               Session("Message") = AddNewLineToMessage(Session("Message"), Trim(MIKRON_CONTRACT_WARNING4))
               bError = True
            End If
         Else
            'Получим документ основание - 
            'договор можно создать на основе утвержденного Протокола ЗК или БСАП
            Set vnikdsTemp1 = Server.CreateObject("ADODB.Recordset")
            sSQL = "SELECT * FROM Docs where DocID = "+sUnicodeSymbol+"'"+Request("DocID")+"' and ((ClassDoc = " + _
                   sUnicodeSymbol+"'"+MIKRON_BSAP+ "' and StatusDevelopment = 4) or (ClassDoc = " + _
                   sUnicodeSymbol+"'"+MIKRON_PROTOCOL+"' and StatusDevelopment = 4))"
            vnikdsTemp1.Open sSQL, Conn, 3, 1, &H1
	        If vnikdsTemp1.EOF Then                
               Session("Message") = AddNewLineToMessage(Session("Message"), MIKRON_CONTRACT_WARNING1)
               bError = True
            End If
            vnikdsTemp1.Close        
         End If
         If bError Then
            bError = False
            Response.Redirect GetURL("listdoc.asp", "?ClassDoc=",Session("CurrentClassDoc"))
         End If
      End If
   End If
End If
''amw_contract

'mikron_add_contract
'Доп. соглашения МИКРОН с увеличением суммы можно делать только на основании утвержденной 
'справки о ценах.
If UCase(Request.ServerVariables("URL")) = UCase("/ChangeDoc.asp") Then
   If InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_ADD_CONTRACT)) = 1 Then
      If Request("create") = "y" Then
         bError = False
         If (Trim(Request("DocIDParent")) = "" and InStr(UCase(Request("ClassDoc")), UCase(MIKRON_ADD_CONTRACT)) = 1) Then
            Session("Message") = AddNewLineToMessage(Session("Message"), MIKRON_CONTRACT_WARNING5)
            bError = True
         ElseIf (Request("DocAmountDoc") > 0 and Request("UserFieldText3") = "Расходный") Then
            'Получим документ основание -
            'Расходное доп.соглашение можно создать на основе утвержденной справки о ценах.
            Set vnikdsTemp1 = Server.CreateObject("ADODB.Recordset")
            sSQL = "SELECT * FROM Docs where DocIDPrevious="+sUnicodeSymbol+"'"+Request("DocIDParent")+"' and ClassDoc="+_
                   sUnicodeSymbol+"'"+MIKRON_RL_MEMO+"' and StatusDevelopment=4"
            vnikdsTemp1.Open sSQL, Conn, 3, 1, &H1
            If vnikdsTemp1.EOF Then
               Session("Message") = AddNewLineToMessage(Session("Message"), MIKRON_CONTRACT_WARNING6)
               bError = True
            End If
            vnikdsTemp1.Close        
         End If
         If bError Then
            bError = False
            Response.Redirect GetURL("listdoc.asp", "?ClassDoc=",Session("CurrentClassDoc"))
         End If
      End If
   End If
End If
''amw_contract

'mikron_export_contract_Additional_Start
'Дополнения к экспортным контрактам МИКРОН можно делать только на основании утвержденного Контракта
If UCase(Request.ServerVariables("URL")) = UCase("/ChangeDoc.asp") Then
   If InStr(UCase(Session("CurrentClassDoc")), UCase(MIKRON_EXPADD_CONTRACT)) = 1 Then
      If Request("create") = "y" Then
         bError = False
         If (InStr(UCase(Request("ClassDoc")), UCase(MIKRON_EXPORT_CONTRACT)) <> 1) Then
            If (Trim(Request("DocIDParent")) = "" and InStr(UCase(Request("ClassDoc")), UCase(MIKRON_EXPADD_CONTRACT)) = 1) Then
               Session("Message") = AddNewLineToMessage(Session("Message"), Trim(MIKRON_EXP_WARNING1))
               bError = True
            End If            
         Else
            'Получим документ основание
            Set vnikdsTemp1 = Server.CreateObject("ADODB.Recordset")
            sSQL = "SELECT * FROM Docs where DocID = "+sUnicodeSymbol+"'"+Request("DocID")+"' and ClassDoc=" + _
                   sUnicodeSymbol+"'"+MIKRON_EXPORT_CONTRACT+"' and StatusDevelopment = 4"
            vnikdsTemp1.Open sSQL, Conn, 3, 1, &H1
            If vnikdsTemp1.EOF Then
               Session("Message") = AddNewLineToMessage(Session("Message"), Trim(MIKRON_EXP_WARNING1) + " " + Trim(MIKRON_DOCNOTFOUND_WARNING1))
               bError = True
            End If               
            vnikdsTemp1.Close        
         End If
         If bError Then
            bError = False
            Response.Redirect GetURL("listdoc.asp", "?ClassDoc=",Session("CurrentClassDoc"))
         End If
      End If
   End If
End If
'mikron_export_contract_Additional_End

''** AMW -                             MIKRON ЗАКУПКИ (end)                                     **
''************************************************************************************************


'vnik_contracts
'Договоры УК можно делать только на основании заявки на закупку УК или Протокола ЦЗК
If UCase(Request.ServerVariables("URL")) = UCase("/ChangeDoc.asp") Then
    If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_CONTRACTS_MC)) > 0 Then
        If Request("create") = "y" Then
            'If Request("DocID") = "" and InStr(UCase(ReplaceRoleFromDir(SIT_OldContractOperator,Session("Department"))),UCase(Session("UserID"))) = 0 Then
            If Request("DocID") = "" Then
                'Session("Message") = AddNewLineToMessage(Session("Message"), SIT_CONTRACTS_MC_WARNING1)
                Session("Message") = AddNewLineToMessage(Session("Message"), SIT_CONTRACTS_MC_WARNING6)
                'Response.Redirect GetURL("listdoc.asp", "?ClassDoc=",Session("CurrentClassDoc")) 
            Else
                'Получим документ основание и в зависимости от вида есть 2 варианта
                Set vnikdsTemp1 = Server.CreateObject("ADODB.Recordset")
                sSQL = "SELECT * FROM Docs where DocID = " +sUnicodeSymbol+"'"+Request("DocID")+"'" + "and ((ClassDoc = N'Закупки*Purchase*Nákup/Протокол ЦЗК*Protocol CPC*Protocol ÚKZZ' and StatusDevelopment = 4) or (ClassDoc = N'Закупки*Purchase*Nákup/Заявка на закупку УК*Purchase order MC*Nákupní objednávka' and StatusDevelopment = 4))"
                AddLogD "vnik444 " + Trim(sSQL)
                vnikdsTemp1.Open sSQL, Conn, 3, 1, &H1             
	            If not vnikdsTemp1.EOF Then                
                    If vnikdsTemp1("ClassDoc") = SIT_PROTOCOLS_CPC Then 
                        Set vnikdsTemp2 = Server.CreateObject("ADODB.Recordset")      
                        sSQL = "Select * from Docs where DocID = "+sUnicodeSymbol+"'"+vnikdsTemp1("DocIDParent")+"'" + " and StatusDevelopment = 4"
                        AddLogD "vnik444 " + Trim(sSQL)
                        vnikdsTemp2.Open sSQL, Conn, 3, 1, &H1
                        If not vnikdsTemp2.EOF Then
                            If vnikdsTemp2("ClassDoc") <> SIT_PURCHASE_ORDER Then
                                Session("Message") = AddNewLineToMessage(Session("Message"), SIT_CONTRACTS_MC_WARNING3)
                                Response.Redirect GetURL("listdoc.asp", "?ClassDoc=",Session("CurrentClassDoc"))
                            End If
                        Else
                            Session("Message") = AddNewLineToMessage(Session("Message"), SIT_CONTRACTS_MC_WARNING3)
                            Response.Redirect GetURL("listdoc.asp", "?ClassDoc=",Session("CurrentClassDoc"))    
                        End If
                        vnikdsTemp2.Close
                    ElseIf vnikdsTemp1("ClassDoc") = SIT_PURCHASE_ORDER Then
                        vnik_SimplePurchase = ""
                        If vnikdsTemp1("Currency") = "RUR" Then
                            If vnikdsTemp1("UserFieldMoney1") > 300000 Then
                                vnik_SimplePurchase = "0"
                            Else
                                vnik_SimplePurchase = "1"
                            End If
                        ElseIf vnikdsTemp1("Currency") = "USD" Then
                            If vnikdsTemp1("UserFieldMoney1") > 10000 Then
                                vnik_SimplePurchase = "0"
                            Else
                                vnik_SimplePurchase = "1"
                            End If
                        ElseIf vnikdsTemp1("Currency") = "EUR" Then
                            If vnikdsTemp1("UserFieldMoney1") > 7500 Then
                                vnik_SimplePurchase = "0"
                            Else
                                vnik_SimplePurchase = "1"
                            End If
                        End If
                        
                        If vnik_SimplePurchase = "" Then
                            Session("Message") = AddNewLineToMessage(Session("Message"), SIT_CONTRACTS_MC_WARNING4)
                            Response.Redirect GetURL("listdoc.asp", "?ClassDoc=",Session("CurrentClassDoc"))
                        ElseIf vnik_SimplePurchase = "0" Then
                            Session("Message") = AddNewLineToMessage(Session("Message"), SIT_CONTRACTS_MC_WARNING5)
                            Response.Redirect GetURL("listdoc.asp", "?ClassDoc=",Session("CurrentClassDoc"))    
                        End If
                    Else
                        'Session("Message") = AddNewLineToMessage(Session("Message"), SIT_CONTRACTS_MC_WARNING1)
                        Session("Message") = AddNewLineToMessage(Session("Message"), SIT_CONTRACTS_MC_WARNING6)
                        'Response.Redirect GetURL("listdoc.asp", "?ClassDoc=",Session("CurrentClassDoc"))
                    End If
                Else
                    'Session("Message") = AddNewLineToMessage(Session("Message"), SIT_CONTRACTS_MC_WARNING2)
                    Session("Message") = AddNewLineToMessage(Session("Message"), SIT_CONTRACTS_MC_WARNING6)
                    'Response.Redirect GetURL("listdoc.asp", "?ClassDoc=",Session("CurrentClassDoc"))    
                End If
                vnikdsTemp1.Close        
            End If
        End If
    End If
End If
'vnik_contracts

'vnik_payment_order
'Договоры УК можно делать только на основании заявки на закупку УК или Протокола ЦЗК
If UCase(Request.ServerVariables("URL")) = UCase("/ChangeDoc.asp") Then
    If InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_PAYMENT_ORDER)) = 1 Then
        If Request("create") = "y" Then
            If Request("DocID") <> "" Then
                Set vnikdsTemp1 = Server.CreateObject("ADODB.Recordset")
                sSQL = "SELECT * FROM Docs where DocID = " +sUnicodeSymbol+"'"+Request("DocID")+"'" + "and (ClassDoc = N'Закупки*Purchase*Nákup/Договоры УК*Contracts MC*Smlouvy SS' and StatusDevelopment = 4)"
                AddLogD "vnik555 " + Trim(sSQL)
                vnikdsTemp1.Open sSQL, Conn, 3, 1, &H1             
	            If not vnikdsTemp1.EOF Then 
	                If vnikdsTemp1("UserFieldText8") = "Рамочный" Then
                        Session("Message") = AddNewLineToMessage(Session("Message"), SIT_PAYMENT_ORDER_WARNING3)
                        Response.Redirect GetURL("listdoc.asp", "?ClassDoc=",Session("CurrentClassDoc"))    
                    End If          
	            End If
	            vnikdsTemp1.Close
            End If
        End If
    End If
End If
'vnik_payment_order

'01.07.2013 kkoshkin sts_Purchase_Payment_Order
'запретить создавать заявки на закупку и оплату СТС пользователям определенных БН
If UCase(Request.ServerVariables("URL")) = UCase("/ChangeDoc.asp") Then
    If InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) = 1 or InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) = 1 Then
        If Request("create") = "y" Then
          sCreator = Session("UserID")
          BusinessUnitsList = GetUsersBusinessUnits(sCreator)
          If InStr(BusinessUnitsList, VbCrLf) > 0 Then
            BusinessUnitsList1 = Split(BusinessUnitsList,VbCrLf)
            BusinessUnitsList = BusinessUnitsList1(0)
          End If     
          If InStr(UCase(STS_ORDER_BN_RESTRICT),UCase(BusinessUnitsList)) > 0 Then
            Session("Message") = AddNewLineToMessage(Session("Message"), STS_ORDER_RESTRICT)
            Response.Redirect GetURL("listdoc.asp", "?ClassDoc=",Session("CurrentClassDoc"))    
          End If
        End If
    End If
End If
'01.07.2013 kkoshkin sts_Purchase_Payment_Order

'Place here your ASP-code running after user form input and BEFORE PayDox database updating
'Ph - 20081206 - start
'Не давать активировать заявки, если найдены не все роли
If UCase(Request.ServerVariables("URL")) = UCase("/MakeActive.asp") Then
  If Request("Active")<>"" Then 'активация
    If InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) = 1 or InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) = 1 Then
      Set dsTemp1 = Server.CreateObject("ADODB.Recordset")
      sSQL = "Select * from Docs where DocID = "+sUnicodeSymbol+"'"+Request("DocID")+"'"
      dsTemp1.Open sSQL, Conn, 3, 1, &H1
	  bUnrecognizedRolesFound = False
      If not dsTemp1.EOF Then
        bUnrecognizedRolesFound = InStr(dsTemp1("ListToReconcile"), """#") > 0 or InStr(dsTemp1("NameAproval"), """#") > 0 or InStr(dsTemp1("NameResponsible"), """#") > 0
      End If
      dsTemp1.Close
      If bUnrecognizedRolesFound Then
        Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorUnrecognizedRoles)
        Response.Redirect GetURL("showdoc.asp", "?docid=", Request("DocID")) + "&template=" + Trim(Request("template"))
      End If
    End If
  End If
End If
'Ph - 20081206 - end

'Ph - 20090123 - start
'Отлавливаем смену утверждающего при редактировании, чтобы потом сделать рассылку 
'(в DBUpdateAfter)
If UCase(Request.ServerVariables("URL")) = UCase("/ChangeDoc.asp") or UCase(Request.ServerVariables("URL")) = UCase("/ChangeField.asp")Then
   If Request("create") <> "y" and UCase(Request("UpdateDoc")) = "YES" Then
      Set dsTemp2 = Server.CreateObject("ADODB.Recordset")
      sSQL = "Select * from Docs where DocID = "+sUnicodeSymbol+"'"+Request("DocID")+"'"
      dsTemp2.Open sSQL, Conn, 3, 1, &H1
      If not dsTemp2.EOF Then
'         If MyCStr(dsTemp2("NameApproved")) = "" and IsReconciliationComplete(dsTemp2("ListToReconcile"), dsTemp2("ListReconciled")) Then
         bApprovalRequired = MyCStr(dsTemp2("NameApproved"))="" and IsReconciliationComplete(dsTemp2("ListToReconcile"), dsTemp2("ListReconciled"))
'         bApprovalRequired = dsTemp2("IsActive")="Y" and MyCStr(dsTemp2("NameApproved"))="" and IsReconciliationCompleteWithOptions(dsTemp2("ListToReconcile"), dsTemp2("ListReconciled"))
'         bApprovalRequired = MyCStr(dsTemp2("NameApproved"))="" and IsReconciliationCompleteWithOptions(dsTemp2("ListToReconcile"), dsTemp2("ListReconciled"))
'         bApprovalRequired = MyCStr(dsTemp2("NameApproved"))=""
         If bApprovalRequired Then
            'Документ находится на стадии утверждения
            Sitronics_OldAproval = dsTemp2("NameAproval")
            bRedirect = False
'amw
'         Else
'            bRedirect = True
''            bRedirect = False
'            S_DocID_Set = dsTemp2("DocID")
'            S_DocIDAdd_Set = dsTemp2("DocID")
'            S_NameAproval_Set = S_NameAproval
'            Session("Message") = AddNewLineToMessage(Session("Message"), S_DocID_Set + " !!! " + dsTemp2("NameAproval") + " !!! " + S_NameAproval)
'            dsTemp2.Close
'            Response.Redirect GetURL("showdoc.asp", "?docid=", Request("DocID")) + "&template=" + Trim(Request("template"))
'amw
         End If
      End If
      dsTemp2.Close
   End If
End If
'Ph - 20090123 - end

'ph - 20090825 - start
If UCase(Request.ServerVariables("URL")) = UCase("/ChangeDoc.asp") Then
  If UCase(Request("UpdateDoc")) = "YES" Then
    If InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PurchaseOrder)) > 0 or InStr(UCase(Session("CurrentClassDoc")), UCase(STS_PaymentOrder)) > 0 Then
      If Trim(SIT_SessionMessageSaveBeforeSendNotitfication) <> "" Then
        'Убираем редирект, чтобы отработал DBUpdateAfter
        bRedirect = False
      End If
    End If
  End If
End If
'ph - 20090825 - end

'minc - поручения - start
If UCase(Request.ServerVariables("URL")) = UCase("/MakeCompleted.asp") and InStr(UCase(Session("Department")), UCase(SIT_MINC)) = 1 Then
    Set dsDocVNIK = Server.CreateObject("ADODB.Recordset")
    sSQL = "select case when convert(datetime,convert(varchar,datecreation,104),104) > convert(datetime,'" + Trim(Request("datecompleted")) + "',104) then N'N' else N'Y' end as DateCompletionValid from docs where docid = N'"+Trim(Request("DocId"))+"'"
    dsDocVNIK.CursorLocation = 3
    dsDocVNIK.Open sSQL, Conn, 3, 1, &H1
    If not dsDocVNIK.EOF Then
        DateCompletionValid = Trim(dsDocVNIK("DateCompletionValid").Value)
    Else
        DateCompletionValid = ""
    End If
    dsDocVNIK.Close   
    if DateCompletionValid = "N" Then
       Session("Message") = AddNewLineToMessage(Session("Message"), "ФАКТИЧЕСКАЯ ДАТА ИСПОЛНЕНИЯ ПОРУЧЕНИЯ НЕ МОЖЕТ БЫТЬ РАНЬШЕ ДАТЫ ВЫДАЧИ!!!")
       Response.Redirect GetURL("showdoc.asp", "?docid=", Request("DocID")) + "&template=" + Trim(Request("template"))
    End If

End If
'minc - поручения - end

%>