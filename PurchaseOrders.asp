<%





'--------------------------------------------------------------------------------------------------проверка лимитов
      'Флаг, что не найдена одна из ролей
      bRoleNotFound = False
	  
      S_ListToReconcile = ""
	  sFullListToReconcile = ""
'Запрос №24 СТС - start





'rmanyushin 136151 13.10.2010 Start
	If is5DivisionSTS(Session("Department")) Then
            AddLogD "@@@OrdersReconcilation - EMS User or CC."
		        S_ListToReconcile = AddUserToReconcileList2(S_ListToReconcile, "", STS_Orders_HeadOfDepartment, Request("DocDepartment"), sDocCreator, sBusinessUnit, bRoleErrorFlag, sFullListToReconcile)
		          If Trim(GetUserID(S_ListToReconcile)) = "" Then
			        S_ListToReconcile = AddUserToReconcileList2(S_ListToReconcile, "", STS_Orders_HeadOfDivision, Request("DocDepartment"), sDocCreator, sBusinessUnit, bRoleErrorFlag, sFullListToReconcile)
		          End If
    			
		         bRoleNotFound = bRoleNotFound or bRoleErrorFlag
	        AddLogD "@@@OrdersReconcilation - S_ListToReconcile EMS1-1: "+S_ListToReconcile

	        AddLogD "@@@PurchaseOrdersReconcilation - IsProject("+Request("UserFieldText3")+") = "+CStr(IsProject(Request("UserFieldText3")))
		          If IsProject(Request("UserFieldText3")) Then
			        'Если не нулевой проект, то следующий согласующий - менеджер проекта
			        S_ListToReconcile = AddUserToReconcileList2(S_ListToReconcile, "", STS_Orders_ProjectManager, ProjectManager, sDocCreator, sBusinessUnit, bRoleErrorFlag, sFullListToReconcile)
	        AddLogD "@@@OrdersReconcilation - S_ListToReconcile EMS2-1: "+S_ListToReconcile
			        iPos = 1
			        If oPayDox.GetNextUserIDInList(ProjectManager, iPos) = "" Then
			          Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInProjectManager)
			        End If
			        oPayDox.GetUserDetails GetUserID(ProjectManager), sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
			        sCostCenter = sDepartment
	        AddLogD "@@@OrdersReconcilation - sCostCenter (If IsProject): "+sCostCenter
		          Else
			        sCostCenter = GetCostCenterByCode(GetCodeFromCode_NameString(Request("UserFieldText1")))
	        AddLogD "@@@OrdersReconcilation - sCostCenter (If not IsProject): "+sCostCenter
		          End If

		          'Директор департамента центра затрат
		          S_ListToReconcile = AddUserToReconcileList2(S_ListToReconcile, "", STS_Orders_HeadOfDepartment, sCostCenter, sDocCreator, sBusinessUnit, bRoleErrorFlag, sFullListToReconcile)
		          bRoleNotFound = bRoleNotFound or bRoleErrorFlag
	        AddLogD "@@@OrdersReconcilation - S_ListToReconcile EMS3-1: "+S_ListToReconcile

    		
	        S_ListToReconcile = AddUserToReconcileList2(S_ListToReconcile, "", STS_Orders_DptEMS, sCostCenter, sDocCreator, sBusinessUnit, bRoleErrorFlag, sFullListToReconcile)
	        AddLogD "@@@OrdersReconcilation - S_ListToReconcile EMS4-1: "+ S_ListToReconcile
    		
	        '20090703 - Заявка на изменение списка согласования PO - start
		          sCurDepLevel = GetDepartmentLevel(sCostCenter)
		          If sCurDepLevel <> "" Then
			        nCurDepLevel = CInt(Right(sCurDepLevel, 1))
		          Else
			        nCurDepLevel = 0
		          End If
		          'Директор дивизиона центра затрат, если превышен лимит директора департамента или проект нулевой и центром затрат является дивизион
		          If (USD_Amount > STS_HeadOfDepartment_Limit) or (not(IsProject(Request("UserFieldText3"))) and nCurDepLevel = 1) Then
	        '20090703 - Заявка на изменение списка согласования PO - end
	        '      'Директор дивизиона центра затрат, если превышен лимит директора департамента
	        '      If USD_Amount > STS_HeadOfDepartment_Limit Then
			        S_ListToReconcile = AddUserToReconcileList2(S_ListToReconcile, "", STS_Orders_HeadOfDivision, sCostCenter, sDocCreator, sBusinessUnit, bRoleErrorFlag, sFullListToReconcile)
			        bRoleNotFound = bRoleNotFound or bRoleErrorFlag
	        AddLogD "@@@OrdersReconcilation - S_ListToReconcile EMS5-1: "+S_ListToReconcile
		          End If
    		
	        'rmanyushin 105583 21.06.2010 Start
	        If is789DivisionSTS(Session("Department")) or is789DivisionSTS(S_UserFieldText1) Then
		        S_ListToReconcile = AddUserToReconcileList2(S_ListToReconcile, "", STS_Orders_DptOSSBSS, sCostCenter, sDocCreator, sBusinessUnit, bRoleErrorFlag, sFullListToReconcile)
		        AddLogD "@@@OrdersReconcilation - S_ListToReconcile EMS6-1 STS_Orders_DptOSSBSS: "+ S_ListToReconcile
	        End if
	        'rmanyushin 105583 21.06.2010 End	
    		
	        '20090703 - Заявка на изменение списка согласования PO - end

'Запрос №30 - СТС - start
		      S_ListToReconcile = AddUserToReconcileList2(S_ListToReconcile, " ", STS_Purchase_Logistics_Department, sCostCenter, sDocCreator, sBusinessUnit, bRoleErrorFlag, sFullListToReconcile)
		      AddLogD "@@@OrdersReconcilation - S_ListToReconcile STS_Purchase_Logistics_Department: "+ S_ListToReconcile
'Запрос №30 - СТС - end

	        If USD_Amount > STS_FinancialControl_Limit Then
			        S_ListToReconcile = AddUserToReconcileList2(S_ListToReconcile, VbCrLf, STS_Orders_FinancialControl, sCostCenter, sDocCreator, sBusinessUnit, bRoleErrorFlag, sFullListToReconcile)
			        bRoleNotFound = bRoleNotFound or bRoleErrorFlag
	        AddLogD "@@@OrdersReconcilation - S_ListToReconcile EMS6-3: "+S_ListToReconcile
			        If USD_Amount > STS_FinDirector_Limit Then
			          S_ListToReconcile = AddUserToReconcileList2(S_ListToReconcile, "", STS_Orders_FinDirector, sCostCenter, sDocCreator, sBusinessUnit, bRoleErrorFlag, sFullListToReconcile)
			          bRoleNotFound = bRoleNotFound or bRoleErrorFlag
	        AddLogD "@@@OrdersReconcilation - S_ListToReconcile 6: "+S_ListToReconcile
			          S_NameAproval = GetRoleForOrders_WithCheck(STS_Orders_GenDirector, sCostCenter, "", sBusinessUnit, bRoleErrorFlag)
			          sFullListToReconcile = sFullListToReconcile+VbCrLf+DOCS_NameAproval+VbCrLf+S_NameAproval+GetOrderRoleDescription(STS_Orders_GenDirector, sCostCenter, True, sBusinessUnit)
			          bRoleNotFound = bRoleNotFound or bRoleErrorFlag
			        Else
			          S_NameAproval = GetRoleForOrders_WithCheck(STS_Orders_FinDirector, sCostCenter, "", sBusinessUnit, bRoleErrorFlag)
			          sFullListToReconcile = sFullListToReconcile+DOCS_NameAproval+VbCrLf+S_NameAproval+GetOrderRoleDescription(STS_Orders_FinDirector, sCostCenter, True, sBusinessUnit)
			          bRoleNotFound = bRoleNotFound or bRoleErrorFlag
			        End If
		          Else
			        S_NameAproval = GetRoleForOrders_WithCheck(STS_Orders_FinancialControl, sCostCenter, "", sBusinessUnit, bRoleErrorFlag)
			        sFullListToReconcile = sFullListToReconcile+VbCrLf+DOCS_NameAproval+VbCrLf+S_NameAproval+GetOrderRoleDescription(STS_Orders_FinancialControl, sCostCenter, True, sBusinessUnit)
			        bRoleNotFound = bRoleNotFound or bRoleErrorFlag
		          End If
	
	Else 'rmanyushin 136151
	
	         S_ListToReconcile = AddUserToReconcileList2(S_ListToReconcile, "", STS_Orders_HeadOfDepartment, Request("DocDepartment"), sDocCreator, sBusinessUnit, bRoleErrorFlag, sFullListToReconcile)
	          If Trim(GetUserID(S_ListToReconcile)) = "" Then
                S_ListToReconcile = AddUserToReconcileList2(S_ListToReconcile, "", STS_Orders_HeadOfDivision, Request("DocDepartment"), sDocCreator, sBusinessUnit, bRoleErrorFlag, sFullListToReconcile)
              End If
        'Запрос №24 СТС - end
              bRoleNotFound = bRoleNotFound or bRoleErrorFlag
        AddLogD "@@@OrdersReconcilation - S_ListToReconcile 1: "+S_ListToReconcile

        AddLogD "@@@PurchaseOrdersReconcilation - IsProject("+Request("UserFieldText3")+") = "+CStr(IsProject(Request("UserFieldText3")))
              If IsProject(Request("UserFieldText3")) Then
                'Если не нулевой проект, то следующий согласующий - менеджер проекта
                S_ListToReconcile = AddUserToReconcileList2(S_ListToReconcile, VbCrLf, STS_Orders_ProjectManager, ProjectManager, sDocCreator, sBusinessUnit, bRoleErrorFlag, sFullListToReconcile)
        AddLogD "@@@OrdersReconcilation - S_ListToReconcile 2: "+S_ListToReconcile
		        iPos = 1
		        If oPayDox.GetNextUserIDInList(ProjectManager, iPos) = "" Then
                  Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorInProjectManager)
		        End If
		        oPayDox.GetUserDetails GetUserID(ProjectManager), sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
		        sCostCenter = sDepartment
        AddLogD "@@@OrdersReconcilation - sCostCenter (If IsProject): "+sCostCenter
              Else
		        sCostCenter = GetCostCenterByCode(GetCodeFromCode_NameString(Request("UserFieldText1")))
        AddLogD "@@@OrdersReconcilation - sCostCenter (If not IsProject): "+sCostCenter
	          End If

              'Директор департамента центра затрат
              S_ListToReconcile = AddUserToReconcileList2(S_ListToReconcile, VbCrLf, STS_Orders_HeadOfDepartment, sCostCenter, sDocCreator, sBusinessUnit, bRoleErrorFlag, sFullListToReconcile)
              bRoleNotFound = bRoleNotFound or bRoleErrorFlag
        AddLogD "@@@OrdersReconcilation - S_ListToReconcile 3: "+S_ListToReconcile

        '20090703 - Заявка на изменение списка согласования PO - start
              sCurDepLevel = GetDepartmentLevel(sCostCenter)
              If sCurDepLevel <> "" Then
                nCurDepLevel = CInt(Right(sCurDepLevel, 1))
              Else
                nCurDepLevel = 0
              End If
              'Директор дивизиона центра затрат, если превышен лимит директора департамента или проект нулевой и центром затрат является дивизион
              If (USD_Amount > STS_HeadOfDepartment_Limit) or (not(IsProject(Request("UserFieldText3"))) and nCurDepLevel = 1) Then
        '20090703 - Заявка на изменение списка согласования PO - end
        '      'Директор дивизиона центра затрат, если превышен лимит директора департамента
        '      If USD_Amount > STS_HeadOfDepartment_Limit Then
                S_ListToReconcile = AddUserToReconcileList2(S_ListToReconcile, VbCrLf, STS_Orders_HeadOfDivision, sCostCenter, sDocCreator, sBusinessUnit, bRoleErrorFlag, sFullListToReconcile)
                bRoleNotFound = bRoleNotFound or bRoleErrorFlag
        AddLogD "@@@OrdersReconcilation - S_ListToReconcile 4: "+S_ListToReconcile
              End If
        '20090703 - Заявка на изменение списка согласования PO - end
          
          'rmanyushin 158840 18.12.2010 Start
	        If is3DivisionSTS(S_UserFieldText1) or is3DivisionSTS(Session("Department")) Then
		        S_ListToReconcile = AddUserToReconcileList2(S_ListToReconcile, VbCrLf, STS_Orders_DirNS, sCostCenter, sDocCreator, sBusinessUnit, bRoleErrorFlag, sFullListToReconcile)
		        AddLogD "@@@OrdersReconcilation - S_ListToReconcile STS_Orders_DirNS: "+ S_ListToReconcile
	        End If
	      'rmanyushin 158840 18.12.2010 End  
	        
	        If is5DivisionSTS(S_UserFieldText1) Then
		        S_ListToReconcile = AddUserToReconcileList2(S_ListToReconcile, VbCrLf, STS_Orders_DptEMS, sCostCenter, sDocCreator, sBusinessUnit, bRoleErrorFlag, sFullListToReconcile)
		        AddLogD "@@@OrdersReconcilation - S_ListToReconcile STS_Orders_DptEMS: "+ S_ListToReconcile
	        End If
        	
	        'rmanyushin 105583 21.06.2010 Start
		        If is789DivisionSTS(Session("Department")) or is789DivisionSTS(S_UserFieldText1) Then
			        S_ListToReconcile = AddUserToReconcileList2(S_ListToReconcile, VbCrLf, STS_Orders_DptOSSBSS, sCostCenter, sDocCreator, sBusinessUnit, bRoleErrorFlag, sFullListToReconcile)
			        AddLogD "@@@OrdersReconcilation - S_ListToReconcile STS_Orders_DptOSSBSS: "+ S_ListToReconcile
		        End if
            'rmanyushin 105583 21.06.2010 End	
'Запрос №30 - СТС - start
		      S_ListToReconcile = AddUserToReconcileList2(S_ListToReconcile, VbCrLf, STS_Purchase_Logistics_Department, sCostCenter, sDocCreator, sBusinessUnit, bRoleErrorFlag, sFullListToReconcile)
		      AddLogD "@@@OrdersReconcilation - S_ListToReconcile STS_Purchase_Logistics_Department: "+ S_ListToReconcile
'Запрос №30 - СТС - end
          
	       If USD_Amount > STS_FinancialControl_Limit Then
                S_ListToReconcile = AddUserToReconcileList2(S_ListToReconcile, VbCrLf, STS_Orders_FinancialControl, sCostCenter, sDocCreator, sBusinessUnit, bRoleErrorFlag, sFullListToReconcile)
                bRoleNotFound = bRoleNotFound or bRoleErrorFlag
        AddLogD "@@@OrdersReconcilation - S_ListToReconcile 5: "+S_ListToReconcile
                If USD_Amount > STS_FinDirector_Limit Then
                  S_ListToReconcile = AddUserToReconcileList2(S_ListToReconcile, VbCrLf, STS_Orders_FinDirector, sCostCenter, sDocCreator, sBusinessUnit, bRoleErrorFlag, sFullListToReconcile)
                  bRoleNotFound = bRoleNotFound or bRoleErrorFlag
        AddLogD "@@@OrdersReconcilation - S_ListToReconcile 6: "+S_ListToReconcile
                  S_NameAproval = GetRoleForOrders_WithCheck(STS_Orders_GenDirector, sCostCenter, "", sBusinessUnit, bRoleErrorFlag)
		          sFullListToReconcile = sFullListToReconcile+VbCrLf+DOCS_NameAproval+VbCrLf+S_NameAproval+GetOrderRoleDescription(STS_Orders_GenDirector, sCostCenter, True, sBusinessUnit)
                  bRoleNotFound = bRoleNotFound or bRoleErrorFlag
                Else
                  S_NameAproval = GetRoleForOrders_WithCheck(STS_Orders_FinDirector, sCostCenter, "", sBusinessUnit, bRoleErrorFlag)
                  sFullListToReconcile = sFullListToReconcile+VbCrLf+DOCS_NameAproval+VbCrLf+S_NameAproval+GetOrderRoleDescription(STS_Orders_FinDirector, sCostCenter, True, sBusinessUnit)
                  bRoleNotFound = bRoleNotFound or bRoleErrorFlag
                End If
	          Else
                S_NameAproval = GetRoleForOrders_WithCheck(STS_Orders_FinancialControl, sCostCenter, "", sBusinessUnit, bRoleErrorFlag)
                sFullListToReconcile = sFullListToReconcile+VbCrLf+DOCS_NameAproval+VbCrLf+S_NameAproval+GetOrderRoleDescription(STS_Orders_FinancialControl, sCostCenter, True, sBusinessUnit)
                bRoleNotFound = bRoleNotFound or bRoleErrorFlag
              End If
	End If 'rmanyushin 136151
	
'rmanyushin 136151 13.10.2010 End	

      'Проверяем все ли участники найдены
      If bRoleNotFound Then
        Session("Message") = AddNewLineToMessage(Session("Message"), SIT_ErrorNotAllUsersFound)
      Else
AddLogD "@@@OrdersReconcilation - S_NameAproval: "+S_NameAproval
	    'Убираем из согласующих утверждающего, если есть
        S_ListToReconcile = RemoveUserFromListWithDescriptions(S_ListToReconcile, S_NameAproval)
'        S_ListToReconcile = Replace(S_ListToReconcile, S_NameAproval, "")
AddLogD "@@@OrdersReconcilation - S_ListToReconcile 7: "+S_ListToReconcile
        'Убираем лишние строки
	    S_ListToReconcile = Replace(S_ListToReconcile, VbCrLf+VbCrLf, VbCrLf)
		If InStr(S_ListToReconcile, VbCrLf) = 1 Then
		  S_ListToReconcile = Replace(S_ListToReconcile, VbCrLf, "", 1 ,1)
		End If
AddLogD "@@@OrdersReconcilation - S_ListToReconcile 8: "+S_ListToReconcile
      End If

	  'Выводим в сообщение полный маршрут документа
      Session("Message") = AddNewLineToMessage(Session("Message"), VbCrLf+"</b><font color=white size=""-2"">"+DOCS_ListToReconcile+VbCrLf+HTMLEncode(sFullListToReconcile)+"</font>")
	
%>