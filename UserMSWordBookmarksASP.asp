<% 'Insert your Case branches here

'SAY 2008-10-31
        
        'vnik_purchase_order
        Case "PO_NameAproval"
            iPos = 1
			sUserID = oPayDox.GetNextUserIDInList(ds("NameAproval"), iPos)
            'AddLogD "vnik897 " + Trim(sUserID)
            oPayDox.GetUserDetails sUserID, sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
			'AddLogD "vnik897 " + Trim(sPosition + " " + SIT_SurnameGN(SurnameGN(DelOtherLangFromFolder(sName))))
			oPayDox.InsertBookmarkText sPosition + " " + SIT_SurnameGN(SurnameGN(DelOtherLangFromFolder(sName))), "PO_NameAproval"
		Case "PO_HeadOfDocControl"
		    AddLogD "vnik897 " + Trim(SIT_RolesDirSitronics)
		    AddLogD "vnik897 " + Trim(SIT_HeadOfDocControl)
		    AddLogD "vnik897 " + Trim(GetUserDirValue(SIT_RolesDirSitronics, SIT_HeadOfDocControl + ";", 1, 2))
		    
		    iPos = 1
			sUserID = oPayDox.GetNextUserIDInList(GetUserDirValue(SIT_RolesDirSitronics, SIT_HeadOfDocControl, 1, 2), iPos)
		    oPayDox.GetUserDetails sUserID, sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
			oPayDox.InsertBookmarkText sPosition + " " + SIT_SurnameGN(SurnameGN(DelOtherLangFromFolder(sName))), "PO_HeadOfDocControl"
		    'oPayDox.InsertBookmarkText "Заведующая канцелярией " + SIT_SurnameGN(SurnameGN(DelOtherLangFromFolder(GetUserDirValue(SIT_RolesDirSitronics, SIT_HeadOfDocControl, 1, 2)))), "PO_HeadOfDocControl"    		    
        Case "PO_DateActivation"
            oPayDox.InsertBookmarkText Left(MyDate(ds("DateActivation")),10), "PO_DateActivation"   
        Case "PO_MyViewerGN"
			oPayDox.GetUserDetails oPayDox.GetNextUserIDInList(ds("NameAproval"), 1), sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
	 		oPayDox.InsertBookmarkText " " + GivenNames(DelOtherLangFromFolder(sName)), "PO_MyViewerGN"      
        Case "PO_ContractType"
            oPayDox.InsertBookmarkText ds("ContractType"), "PO_ContractType"
        Case "PO_UserFieldText8"
            oPayDox.InsertBookmarkText ds("UserFieldText8"), "PO_UserFieldText8"
        Case "PO_UserFieldText2"
            oPayDox.InsertBookmarkText ds("UserFieldText2"), "PO_UserFieldText2"
        Case "PO_UserFieldText1"
            oPayDox.InsertBookmarkText ds("UserFieldText1"), "PO_UserFieldText1"
            
        Case "PO_UserFieldText3"
            oPayDox.InsertBookmarkText ds("UserFieldText3"), "PO_UserFieldText3"
        Case "PO_UserFieldText4"
            oPayDox.InsertBookmarkText ds("UserFieldText3"), "PO_UserFieldText4"  
        Case "PO_UserFieldText5"
            oPayDox.InsertBookmarkText ds("UserFieldText5"), "PO_UserFieldText5"   
            
        Case "DateProtocolRTI" 
	 		oPayDox.InsertBookmarkText Replace(MyDateLong(ds("UserFieldDate2")),",","") + "г.", "DateProtocolRTI"       

'rti_purchase_order
        Case "RTI_HeadOfAuthor"            
            iPos = 1
			sUserID = oPayDox.GetNextUserIDInList(GetNearestChief(ds("Department"), ds("Author"), sBusinessUnit),iPos)
			oPayDox.GetUserDetails sUserID, sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
			oPayDox.InsertBookmarkText SIT_SurnameGN(SurnameGN(DelOtherLangFromFolder(sName))), "RTI_HeadOfAuthor"
'rti_contract
        'аффилированность
        Case "RTIC_UserFieldText1"
            oPayDox.InsertBookmarkText ds("UserFieldText1"), "RTIC_UserFieldText1"
        'предмет договора
        Case "RTIC_UserFieldText2"
            oPayDox.InsertBookmarkText ds("UserFieldText2"), "RTIC_UserFieldText2"
        'Код заказа
        Case "RTIC_UserFieldText3"
            oPayDox.InsertBookmarkText ds("UserFieldText3"), "RTIC_UserFieldText3"          
        'контрагент
        Case "RTIC_PartnerName"
            oPayDox.InsertBookmarkText ds("PartnerName"), "RTIC_PartnerName"
            
'rti_contract
        Case "PO_UserFieldText6"
            oPayDox.InsertBookmarkText ds("UserFieldText6"), "PO_UserFieldText6"
        Case "PO_UserFieldMoney1"
            oPayDox.InsertBookmarkText ds("UserFieldMoney1"), "PO_UserFieldMoney1"
        Case "PO_UserFieldMoney2"
            oPayDox.InsertBookmarkText ds("UserFieldMoney2"), "PO_UserFieldMoney2"
        Case "PO_Currency"
            oPayDox.InsertBookmarkText ds("Currency"), "PO_Currency"
        Case "PO_UserFieldText7"
            oPayDox.InsertBookmarkText ds("UserFieldText7"), "PO_UserFieldText7"
'mikron
        Case "MIKRON_Description"
            oPayDox.InsertBookmarkText ds("Description"), "MIKRON_Description"     
        Case "MIKRON_AmountDoc"
            oPayDox.InsertBookmarkText ds("AmountDoc"), "MIKRON_AmountDoc"    
        Case "MIKRON_Currency"
            oPayDox.InsertBookmarkText ds("Currency"), "MIKRON_Currency"    
        Case "MIKRON_PartnerName"
            oPayDox.InsertBookmarkText ds("PartnerName"), "MIKRON_PartnerName"    
'mikron            
        Case "VicePresident"
            iPos = 1
			sUserID = oPayDox.GetNextUserIDInList(GetChiefOfDepUpperByLevel(ds("Department"), 1, ds("Author"), sBusinessUnit),iPos)
			oPayDox.GetUserDetails sUserID, sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
			oPayDox.InsertBookmarkText sPosition + " " + SIT_SurnameGN(SurnameGN(DelOtherLangFromFolder(sName))), "VicePresident"
		Case "NameVicePresident"
		    iPos = 1
			sUserID = oPayDox.GetNextUserIDInList(GetChiefOfDepUpperByLevel(ds("Department"), 1, ds("Author"), sBusinessUnit),iPos)
			oPayDox.GetUserDetails sUserID, sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
			oPayDox.InsertBookmarkText SIT_SurnameGN(SurnameGN(DelOtherLangFromFolder(sName))), "NameVicePresident"
		Case "PositionVicePresident"
            iPos = 1
			sUserID = oPayDox.GetNextUserIDInList(GetChiefOfDepUpperByLevel(ds("Department"), 1, ds("Author"), sBusinessUnit),iPos)
			oPayDox.GetUserDetails sUserID, sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
			oPayDox.InsertBookmarkText sPosition, "PositionVicePresident"
		Case "PO_UserFieldText5"
            oPayDox.InsertBookmarkText ds("UserFieldText5"), "PO_UserFieldText5"
        Case "PO_QuantityDoc"
            oPayDox.InsertBookmarkText ds("QuantityDoc"), "PO_QuantityDoc"
        'vnik_purchase_order
        
		Case "MyAddresseName" '20080922 - Ph
	 		oPayDox.InsertBookmarkText SIT_SurnameGN(SurnameGN(ds("UserFieldText5"))), "MyAddresseName"
		Case "MyAddressePosition"
	 		oPayDox.InsertBookmarkText oPayDox.NamesIn3ndForm(GetPosition(GetLogin(ds("UserFieldText5")))), "MyAddressePosition"
		Case "MyAddresseNameGN"
			oPayDox.GetUserDetails oPayDox.GetNextUserIDInList(ds("UserFieldText5"), 1), sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
	 		oPayDox.InsertBookmarkText GivenNames(DelOtherLangFromFolder(sName)), "MyAddresseNameGN"



		Case "NameAprovalSignature"
			MSWordInsertUserSignatureToDoc GetUserID(ds("NameAproval")), "NameAprovalSignature"

		Case "DateCreationLong" 
	 		oPayDox.InsertBookmarkText MyDateLong(ds("DateCreation")), "DateCreationLong"
		Case "DateCompletionLong" 
	 		oPayDox.InsertBookmarkText MyDateLong(ds("DateCompletion")), "DateCompletionLong"
		Case "NameCreationPhone" 'Your bookmark name
			oPayDox.GetUserDetails GetUserID(ds("NameCreation")), sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
	 		oPayDox.InsertBookmarkText sPhone, "NameCreationPhone"
		Case "UserFieldDate1Long" 'Your bookmark name
	 		oPayDox.InsertBookmarkText MyDateLong(ds("UserFieldDate1")), "UserFieldDate1Long" 'Your code
		Case "UserFieldDate2Long" 'Your bookmark name
	 		oPayDox.InsertBookmarkText MyDateLong(ds("UserFieldDate2")), "UserFieldDate2Long" 'Your code
		Case "DateActivationLong" 'Your bookmark name
	 		oPayDox.InsertBookmarkText MyDateLong(ds("DateActivation")), "DateActivationLong" 'Your code
		Case "DateActivationLongOfficial" 'Your bookmark name
	 		oPayDox.InsertBookmarkText MyCStr(Day(ds("DateActivation")))+" day of "+MyMonthName(Month(ds("DateActivation")))+", "+MyCStr(Year(ds("DateActivation"))), "DateActivationLongOfficial" 'Your code
		Case "ApprovalName" 
	 		oPayDox.InsertBookmarkText GetName(ds("NameAproval")), "ApprovalName" 
		Case "ApprovedName" 
	 		oPayDox.InsertBookmarkText GetName(ds("NameApproved")), "ApprovedName" 
		Case "WhoIsAproval" 
	 		oPayDox.InsertBookmarkText GetName(ds("UserFieldText8")), "WhoIsAproval" 
		Case "WhoIsAprovalPosition" 
			oPayDox.GetUserDetails GetUserID(ds("UserFieldText8")), sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
	 		oPayDox.InsertBookmarkText sPosition, "WhoIsAprovalPosition" 
		Case "ControlName" 
	 		oPayDox.InsertBookmarkText GetName(ds("NameControl")), "ControlName" 
		Case "ResponsibleName" 
	 		oPayDox.InsertBookmarkText GetName(ds("NameResponsible")), "ResponsibleName" 
		Case "LISTCOMMENTS" 'Your bookmark name
			InsertBookmarkComments "comment", "LISTCOMMENTS" 'outputs all the doc comments having "comment" as a comment type
		Case "VISA" 'Your bookmark name
			InsertBookmarkComments "VISA", "VISA" 
		Case "NameResponsibleIDentification" 'Your bookmark name
			oPayDox.GetUserDetails GetUserID(ds("NameResponsible")), sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
	 		oPayDox.InsertBookmarkText sIDentification, "NameResponsibleIDentification" 'Your code
		Case "NameResponsibleIDNo" 'Your bookmark name
			oPayDox.GetUserDetails GetUserID(ds("NameResponsible")), sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
	 		oPayDox.InsertBookmarkText sIDNo, "NameResponsibleIDNo" 'Your code
		Case "NameResponsibleIDIssuedBy" 'Your bookmark name
			oPayDox.GetUserDetails GetUserID(ds("NameResponsible")), sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
	 		oPayDox.InsertBookmarkText sIDIssuedBy, "NameResponsibleIDIssuedBy" 'Your code
		Case "NameResponsibleIDIssueDate" 'Your bookmark name
			oPayDox.GetUserDetails GetUserID(ds("NameResponsible")), sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
	 		oPayDox.InsertBookmarkText MyDateLong(dIDIssueDate), "NameResponsibleIDIssueDate" 'Your code
		Case "MASSPRINT_FIO" 'Your bookmark name
			 oPayDox.InsertBookmarkTextDoc doc, dsDoc("FIO"), "MASSPRINT_FIO" 
		Case "MASSPRINT_F" 'Your bookmark name
			 oPayDox.InsertBookmarkTextDoc doc, Surname(dsDoc("FIO")), "MASSPRINT_F" 
		Case "MASSPRINT_IO" 'Your bookmark name
			 oPayDox.InsertBookmarkTextDoc doc, GivenNames(dsDoc("FIO")), "MASSPRINT_IO" 
		Case "MyAuthor" '20080922 - Ph
	 		oPayDox.InsertBookmarkText SIT_SurnameGN(SurnameGN(ds("Author"))), "MyAuthor"
		Case "MyAuthor_1" 
	 		oPayDox.InsertBookmarkText SIT_SurnameGN(SurnameGN(ds("Author"))), "MyAuthor_1"
		Case "MyAuthor_2" 
	 		oPayDox.InsertBookmarkText SIT_SurnameGN(SurnameGN(ds("Author"))), "MyAuthor_2"
        Case "LSoglPhone"
			oPayDox.GetUserDetails GetUserID(ds("Author")), sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
	 		oPayDox.InsertBookmarkText sPhone, "LSoglPhone"
	    
	    Case "UserDepartment"
			oPayDox.GetUserDetails GetUserID(ds("NameAproval")), sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
			aValues=Split(sDepartment, "/")

            sDptmnt = ""

            If inStr(Right(sDepartment,1),"/") = 1 and UBound(aValues) > 0  then
              sDptmnt = aValues(UBound(aValues)-1)
              Else if inStr(Right(sDepartment,1),"/") = 0 and UBound(aValues) > 0  then
              sDptmnt = aValues(UBound(aValues))
              Else sDptmnt = sDepartment
            End If
            End If
            
	 		oPayDox.InsertBookmarkText sDptmnt, "UserDepartment"
	 		
	 	 Case "AuthorDepartment"
			oPayDox.GetUserDetails GetUserID(ds("Author")), sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
			aValues=Split(sDepartment, "/")

            sDptmnt = ""

            If inStr(Right(sDepartment,1),"/") = 1 and UBound(aValues) > 0  then
              sDptmnt = aValues(UBound(aValues)-1)
              Else if inStr(Right(sDepartment,1),"/") = 0 and UBound(aValues) > 0  then
              sDptmnt = aValues(UBound(aValues))
              Else sDptmnt = sDepartment
            End If
            End If
            
	 		oPayDox.InsertBookmarkText sDptmnt, "AuthorDepartment"

	 		
        Case "Contract_MC_UserFieldText3"
            oPayDox.InsertBookmarkText ds("UserFieldText3"), "Contract_MC_UserFieldText3"

        Case "Contract_MC_AddFieldText1"
            oPayDox.InsertBookmarkText ds("AddFieldText1"), "Contract_MC_AddFieldText1"

        Case "Contract_MC_AddFieldText2"
            oPayDox.InsertBookmarkText ds("AddFieldText2"), "Contract_MC_AddFieldText2"

	 		
		Case "AuthorPosition"
	 		oPayDox.InsertBookmarkText GetPosition(GetLogin(ds("Author"))), "AuthorPosition"
		Case "AuthorPosition_1"
	 		oPayDox.InsertBookmarkText GetPosition(GetLogin(ds("Author"))), "AuthorPosition_1"
		Case "AuthorPosition_2"
	 		oPayDox.InsertBookmarkText GetPosition(GetLogin(ds("Author"))), "AuthorPosition_2"
		Case "MyNameAproval"
	 		oPayDox.InsertBookmarkText SIT_SurnameGN(SurnameGN(ds("NameAproval"))), "MyNameAproval"
		Case "MyNameAproval3"
				sName=SurnameGN(ds("NameAproval"))
				iTemp1=InStr(sName, " ")
				If iTemp1>0 Then
					sName=Trim(Mid(sName, iTemp1) +" "+Left(sName, iTemp1))
				End If
	 		oPayDox.InsertBookmarkText sName, "MyNameAproval3"
		
		Case "MyNameAproval1"
	 		oPayDox.InsertBookmarkText SIT_SurnameGN(SurnameGN(ds("NameAproval"))), "MyNameAproval1"
		Case "NameAprovalPosition1"
	 		oPayDox.InsertBookmarkText GetPosition(GetLogin(ds("NameAproval"))), "NameAprovalPosition1"
'Ph - 20090201 - start
		Case "MyNameResponsible"
	 		oPayDox.InsertBookmarkText SIT_SurnameGN(SurnameGN(ds("NameResponsible"))), "MyNameResponsible"
		Case "STS_OrderCurrencyRate"
'	 		oPayDox.InsertBookmarkText Eval(Replace(STS_CurrencyRateFormula, "dsDoc(", "ds(")), "STS_OrderCurrencyRate"
            On Error Resume Next
            rVal = Eval(Replace(STS_CurrencyRateFormula, "dsDoc(", "ds("))
            If Err.Number <> 0 Then
                rVal = "ERROR! - " + Err.Description
            End If	
            On Error GoTo 0
	 		oPayDox.InsertBookmarkText rVal, "STS_OrderCurrencyRate"
'Ph - 20090201 - end
		Case "MyViewer"
			oPayDox.GetUserDetails oPayDox.GetNextUserIDInList(ds("ListToView"), 1), sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
	 		oPayDox.InsertBookmarkText SIT_SurnameGN(SurnameGN(DelOtherLangFromFolder(sName))), "MyViewer"
		Case "MyViewerPosition"
	 		oPayDox.InsertBookmarkText GetPosition(oPayDox.GetNextUserIDInList(ds("ListToView"), 1)), "MyViewerPosition"
		Case "MyViewerGN"
			oPayDox.GetUserDetails oPayDox.GetNextUserIDInList(ds("ListToView"), 1), sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
	 		oPayDox.InsertBookmarkText GivenNames(DelOtherLangFromFolder(sName)), "MyViewerGN"

		Case "MyRecipient" 'Ph - 20081112 - Like MyViewer but with position and setting MyRecipientCopy bookmark if several recipients
			iPos = 1
			sUserID = oPayDox.GetNextUserIDInList(ds("ListToView"), iPos)
			If sUserID <> "" Then
				oPayDox.GetUserDetails sUserID, sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
				oPayDox.InsertBookmarkText sPosition + " " + SIT_SurnameGN(SurnameGN(DelOtherLangFromFolder(sName))), "MyRecipient"
				sCopyList = ""
				sUserID = oPayDox.GetNextUserIDInList(ds("ListToView"), iPos)
				Do While sUserID <> ""
					If sCopyList <> "" Then
					  sCopyList = sCopyList + VbCrLf
					End If
					oPayDox.GetUserDetails sUserID, sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
					sCopyList = sCopyList + VbTab + sPosition + " " + SIT_SurnameGN(SurnameGN(DelOtherLangFromFolder(sName)))
					sUserID = oPayDox.GetNextUserIDInList(ds("ListToView"), iPos)
				Loop
				oPayDox.InsertBookmarkText sCopyList, "MyRecipientCopy"
			End If
		Case "MyNameAprovalSignature" 'Ph - 20081112
			If InStr(UCase(ds("UserFieldText2")), "ФАКС") or InStr(UCase(ds("UserFieldText2")), "FAX") Then
				MSWordInsertUserSignatureToDoc GetUserID(ds("NameAproval")), "MyNameAprovalSignature"
			End If

		Case "MySecurityLevel" '20080922 - Ph
		    If dsDoc("SecurityLevel") = 1 Then
		      sMySecurityLevel = SIT_SECURITYLEVEL_ALL
		    Else
		      sMySecurityLevel = SIT_SECURITYLEVEL_LISTONLY
		    End If
	 		oPayDox.InsertBookmarkText sMySecurityLevel, "MySecurityLevel"
		Case "LastReconcilationDate" '20081019 - Ph
		    If InStr(Session("CurrentClassDoc"), SIT_NORM_DOCS) > 0 Then
    	 		oPayDox.InsertBookmarkText SIT_LastReconcilationDate+MyDate(ds("UserFieldDate3")), "LastReconcilationDate"
		    End If
		Case "ListToView_Position"
			S_ListToView=dsDoc("ListToView")
			iPos = 1
    		sUserID = oPayDox.GetNextUserIDInList(S_ListToView, iPos)
			Do While sUserID<>""
				sPosition=GetPosition(sUserID)
				sFullName=oPayDox.GetFullUserFromList(S_ListToView, sUserID)
				sName=GetName(sFullName)
				sName=SurnameGN(sName)
				iTemp1=InStr(sName, " ")
				If iTemp1>0 Then
					sName=Trim(Mid(sName, iTemp1) +" "+Left(sName, iTemp1))
				End If
			 	oPayDox.InsertBookmarkText sPosition, "ListToView_Position"
			 	oPayDox.InsertBookmarkText sName, "ListToView_Name"
    			sUserID = oPayDox.GetNextUserIDInList(S_ListToView, iPos)
    			If sUserID = "" Then
        			Exit Do
				Else
			   		MSWordInsertRowInTable 1 'Insert correct MS Word document table number
			   		MSWordAddBookmarksToTable Array("ListToView_Position","ListToView_Name"), 1, 0, 1 'Insert correct MS Word document table number
    			End If
			Loop
		Case "AuthorPhone"
			oPayDox.GetUserDetails GetLogin(ds("Author")), sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
	 		oPayDox.InsertBookmarkText sPhone, "AuthorPhone"
		Case "CostCenterCode"
		    CostCenter = Trim(ds("UserFieldText1"))
	 		oPayDox.InsertBookmarkText Left(CostCenter, InStr(CostCenter, " ")-1), "CostCenterCode"
		Case "CostCenterName"
		    CostCenter = Trim(ds("UserFieldText1"))
	 		oPayDox.InsertBookmarkText Right(CostCenter, Len(CostCenter)-InStr(CostCenter, " ")), "CostCenterName"

		Case "XXX" 'Your bookmark name
			 oPayDox.InsertBookmarkText "This text will be inserted", "XXX" 'Your code
			 
  Case "LISTVISA1"
    Set dsTemp1 = Server.CreateObject("ADODB.Recordset")
    dsTemp1.Open "select * from Comments where DocID='" + Trim(Request("DocID")) + "' and CommentType='VISA' order by DateCreation desc", Conn, 3, 1, &H1
    stemp = ""
    Do While Not dsTemp1.EOF
        stemp = stemp + MyDate(dsTemp1("DateCreation")) + " " + MyCStr(dsTemp1("UserName")) + " " + MyCStr(dsTemp1("Comment")) + vbCrLf
        dsTemp1.MoveNext
    Loop
    dsTemp1.Close
    oPayDox.InsertBookmarkText stemp, "LISTVISA1" 

  Case "VISATABLEPOSITION"
'Ph - 20090202 - start
    If UCase(Request("StandardFile")) = "R" Then
'Запрос №11 - СТС - start
      If Request("StandardFileName") = "ReconciliationListContractRUS.doc" Then 'шаблон для Договоров
        ReconcilationTableNo = 2
      Else
        ReconcilationTableNo = 1 'Данные о согласующих в первой таблице шаблона
      End If
'Запрос №11 - СТС - end
    Else
      ReconcilationTableNo = 2 'Данные о согласующих во второй таблице шаблона (PaymentOrder)
    End If
'Ph - 20090202 - end

    'Поиск последней отмены согласования
'    sSQL = "select DateCreation from Comments where DocID=N'" + Trim(Request("DocID")) + "' and CommentType=N'VISA' and (SpecialInfo = N'INFOCHANGED' or PATINDEX(N'%Отмена статуса «Согласовано» или «Отказано»%', Comment) > 0 or PATINDEX(N'%Повторное согласование%', Comment) > 0 or PATINDEX(N'%Согласование: Отменено%', Comment) > 0) order by DateCreation desc"
'    sSQL = "select DateCreation from Comments where DocID=N'" + Trim(Request("DocID")) + "' and CommentType=N'VISA' and (SpecialInfo = N'INFOCHANGED' or PATINDEX(N'%Отмена статуса «Согласовано» или «Отказано»%', Comment) > 0 or PATINDEX(N'%Повторное согласование%', Comment) > 0 or PATINDEX(N'%Согласование: Отменено%', Comment) > 0 or PATINDEX(N'%Cancellation «Agreed» or «Refused» status%', Comment) > 0 or PATINDEX(N'%Repeated agree%', Comment) > 0 or PATINDEX(N'%Agree process: Cancelled%', Comment) > 0) order by DateCreation desc"
'Ph - 20081031 - Новый селект. Убрано "Согласование: отменено" из поиска записей, начинающих согласование сначала
    sSQL = "select DateCreation from Comments where DocID=N'" + Trim(Request("DocID")) + "' and CommentType=N'VISA' and (SpecialInfo = N'INFOCHANGED' or PATINDEX(N'%Отмена статуса «Согласовано» или «Отказано»%', Comment) > 0 or PATINDEX(N'%Повторное согласование%', Comment) > 0 or  PATINDEX(N'%Cancellation «Agreed» or «Refused» status%', Comment) > 0 or PATINDEX(N'%Repeated agree%', Comment) > 0) order by DateCreation desc"
AddLogD "UserMSWordBookmarksASP - Search Reconcilation Renew SQL: "+sSQL
    Set dsTemp1 = Server.CreateObject("ADODB.Recordset")
    dsTemp1.Open sSQL, Conn, 3, 1, &H1
    If dsTemp1.EOF Then
      sStartDateCreation = UniDate(VAR_BeginOfTimes)
    Else
      sStartDateCreation = UniDate(dsTemp1("DateCreation"))
    End If
    dsTemp1.Close                      

    ' запрос для поиска по Comments и вывода закладок        
    sSQL = "select * from Comments where DocID=N'" + Trim(Request("DocID")) + "' and CommentType=N'VISA' and DateCreation > "+sStartDateCreation+" order by DateCreation"
AddLogD "UserMSWordBookmarksASP - ReconcilationList SQL: "+sSQL
    dsTemp1.Open sSQL, Conn, 3, 1, &H1

    bPrint=True

    iRecTemp=0	  
    j=0                     
    Do While Not dsTemp1.EOF
      iRecTemp=iRecTemp+1
      sC=MyCStr(dsTemp1("Comment"))
'ph - 20081008 - start
iSITCommentPosition = InStr(MyCStr(dsTemp1("Comment")), "|")
If iSITCommentPosition = 0 Then
  sSITComment = ""
Else
  sSITComment = Trim(Right(MyCStr(dsTemp1("Comment")), Len(MyCStr(dsTemp1("Comment")))-iSITCommentPosition-1))
End If
'ph - 20081008 - end

'kkoshkin - 20120710 - start убираем из комментария фио и логин пользователя
iUserCommentPos = InStr(sSITComment, "/")
If iUserCommentPos > 0 Then
  sSITComment = Trim(Left(sSITComment, iUserCommentPos-1))
End If
'kkoshkin - 20120710 - end

      sSpecialInfo=UCase(MyCStr(dsTemp1("SpecialInfo")))

      If sSpecialInfo="VISAOK" or sSpecialInfo="VISAOKREFUSE" Or InStr(UCase(dsTemp1("Comment")),UCase("делегировать")) > 0 Then
        If Not bPrint Then
          If Not dsTemp1.EOF Then
'Ph - 20090202 - start
'            MSWordInsertRowInTable 1
'            MSWordAddBookmarksToTable Array("VISATABLEPOSITION","VISATABLEFIO","VISATABLERESULT","VISATABLEDATE","VISATABLESIGNATURE","VISATABLECOMMENT"), 1, 0, 1
            MSWordInsertRowInTable ReconcilationTableNo
'Запрос №11 - СТС - start
            If Request("StandardFileName") = "ReconciliationListContractRUS.doc" Then 'лист согласования для Договоров
              MSWordAddBookmarksToTable Array("VISATABLEPOSITION","VISATABLEFIO","VISATABLERESULT","VISATABLEDATE","VISATABLECOMMENT"), ReconcilationTableNo, 0, 1
            Else
              MSWordAddBookmarksToTable Array("VISATABLEPOSITION","VISATABLEFIO","VISATABLERESULT","VISATABLEDATE","VISATABLESIGNATURE","VISATABLECOMMENT"), ReconcilationTableNo, 0, 1
            End If
'Запрос №11 - СТС - end
'Ph - 20090202 - end
          End If
        End If
    
        oPayDox.GetUserDetails dsTemp1("UserID"), sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment

        sFio = SurnameGN(DelOtherLangFromFolder(MyCStr(dsTemp1("UserName"))))

        If InStr(UCase(sC), UCase("Согласовано по умолчанию")) > 0 Then
          sC="Согласовано по умолчанию"
        ElseIf InStr(UCase(sC), UCase("Согласовано")) > 0 or InStr(UCase(sC), UCase("Agreed")) > 0 Then
          sC="Согласовано"        
        End If
        If InStr(UCase(sC), UCase("(#!)")) > 0 Then
          sC="Согласование приостановлено"        
        End If
        If InStr(UCase(sC), UCase("Отказано")) > 0 or InStr(UCase(sC), UCase("Refused")) > 0 Then
          sC="Отказано"
        End If

' перенос на новую строку
          oPayDox.InsertBookmarkText sPosition + "  ", "VISATABLEPOSITION"

        oPayDox.InsertBookmarkText sFio, "VISATABLEFIO"
    
		If InStr(UCase(dsTemp1("Comment")),UCase("делегировать")) > 0 Then
          ' используем уже существующую функцию перевода в падеж вместо недавно написанной
          sC="Делегировано "  + oPayDox.NamesIn3rdForm(GetName(Replace(dsTemp1("Comment"),"Делегировать свои полномочия по согласованию другому пользователю: ","")))
		End If

        oPayDox.InsertBookmarkText sC+VbCrLf, "VISATABLERESULT"
        oPayDox.InsertBookmarkText MyDate(dsTemp1("DateCreation")), "VISATABLEDATE"

'        If sSpecialInfo="VISAOK" Then
'	      MSWordInsertUserSignatureToDoc dsTemp1("UserID"), "VISATABLESIGNATURE"
'        Else
'	      oPayDox.InsertBookmarkText "", "VISATABLESIGNATURE"
'        End If
        
        sUserIDTemp=dsTemp1("UserID")
	    iPosTemp = dsTemp1.Bookmark
	    dsTemp1.MoveFirst
         													  
	    Do While Not dsTemp1.EOF
	      If dsTemp1("UserID")=sUserIDTemp And Not (UCase(MyCStr(dsTemp1("SpecialInfo")))="VISAOK" or UCase(MyCStr(dsTemp1("SpecialInfo")))="VISAOKREFUSE") Then
	         Exit Do
	      End If
	      dsTemp1.MoveNext
	    Loop

        dsTemp1.Bookmark = iPosTemp
        oPayDox.InsertBookmarkText sSITComment, "VISATABLECOMMENT"
        bPrint=False
      End If

' am 08102006 Ищем время поступления док-та для следующего согласующего из списка еще не согласовавших
      If sSpecialInfo="VISAWAITING" Then
        sTimeNextRecon=MyDate(dsTemp1("DateEvent"))      
      End If
      j=j+1      
      dsTemp1.MoveNext
    Loop										    
	
    ' еще не согласовавшие, берем из Docs
    ' 21102006
    If ds("IsActive")="Y" Then

      sUserList=oPayDox.NotYetAgreedList(dsDoc("ListToReconcile"), dsDoc("ListReconciled"))

      If sUserList <> "" Then
     	iTempComment = 1
        i=0

        Do While True
          sUser = oPayDox.GetUserIDFromList(sUserList, iTempComment)

          If sUser = "" Then 
            Exit Do
          End If

          If Not bPrint Then 
'Ph - 20090202 - start
'  	        MSWordInsertRowInTable 1
'	        MSWordAddBookmarksToTableColumn Array("VISATABLEPOSITION"), 1, 0, 1
'            MSWordAddBookmarksToTableColumn Array("VISATABLEFIO"), 1, 0, 1
  	        MSWordInsertRowInTable ReconcilationTableNo
	        MSWordAddBookmarksToTableColumn Array("VISATABLEPOSITION"), ReconcilationTableNo, 0, 1
            MSWordAddBookmarksToTableColumn Array("VISATABLEFIO"), ReconcilationTableNo, 0, 1
'Ph - 20090202 - end
	      End If

	      oPayDox.GetUserDetails sUser, sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment

          oPayDox.InsertBookmarkText sPosition + "  ", "VISATABLEPOSITION"          
	      oPayDox.InsertBookmarkText SurnameGN(DelOtherLangFromFolder(sName)), "VISATABLEFIO"

          If i=0 Then
'Ph - 20090202 - start
'            MSWordAddBookmarksToTableColumn Array("VISATABLERESULT"), 1, 0, 1
'            MSWordAddBookmarksToTableColumn Array("VISATABLEDATE"), 1, 0, 1
            MSWordAddBookmarksToTableColumn Array("VISATABLERESULT"), ReconcilationTableNo, 0, 1
            MSWordAddBookmarksToTableColumn Array("VISATABLEDATE"), ReconcilationTableNo, 0, 1
'Ph - 20090202 - end
                        
            oPayDox.InsertBookmarkText "На согласовании" + VbCrLf, "VISATABLERESULT"          
            oPayDox.InsertBookmarkText sTimeNextRecon, "VISATABLEDATESTART"
          End If
          
          i=i+1
  	      bPrint=False
        Loop
      End If 
	End If  

    dsTemp1.Close

'Запрос №11 - СТС - start - add20100720
    If Request("StandardFileName") = "ReconciliationListContractRUS.doc" Then 'лист согласования для Договоров - добавляем утверждающего
      sSQL = "select *, Comments.DateCreation as CommentDateCreation from Docs left join Comments on (Docs.DocID = Comments.DocID and CommentType='APROVAL') where Docs.DocID = " & sUnicodeSymbol & "'" & Trim(Request("DocID")) & "' order by Comments.DateCreation desc"
      dsTemp1.Open sSQL, Conn, 3, 1, &H1
      If not dsTemp1.EOF and Trim(MyCStr(dsTemp1("NameAproval"))) <> "" Then
        If MyCStr(dsTemp1("NameApproved")) = "" Then
         sNameApproved = GetUserID(MyCStr(dsTemp1("NameAproval")))
        Else
         sNameApproved = GetUserID(MyCStr(dsTemp1("NameApproved")))
        End If
        MSWordInsertRowInTable ReconcilationTableNo
        MSWordAddBookmarksToTable Array("VISATABLEPOSITION","VISATABLEFIO","VISATABLERESULT","VISATABLEDATE","VISATABLECOMMENT"), ReconcilationTableNo, 0, 1
        oPayDox.GetUserDetails sNameApproved, sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
        oPayDox.InsertBookmarkText sPosition, "VISATABLEPOSITION"
        oPayDox.InsertBookmarkText SurnameGN(DelOtherLangFromFolder(sName)), "VISATABLEFIO"
        If MyCStr(dsTemp1("NameApproved")) = "" Then
          oPayDox.InsertBookmarkText DOCS_Approving, "VISATABLERESULT"
        Else
          If InStr(MyCStr(dsTemp1("NameApproved")), "-<") > 0 Then
          sResult = "DOCS_RefusedApp"
            oPayDox.InsertBookmarkText DOCS_RefusedApp, "VISATABLERESULT"
          Else
            sResult = "DOCS_Approved"
            oPayDox.InsertBookmarkText DOCS_Approved, "VISATABLERESULT"
          End If
          Do While UCase(MyCStr(dsTemp1("SpecialInfo"))) <> UCase(sResult) and not dsTemp1.EOF
            dsTemp1.MoveNext
          Loop
          If not dsTemp1.EOF Then
            oPayDox.InsertBookmarkText MyDate(dsTemp1("CommentDateCreation")), "VISATABLEDATE"
            sComment = Trim(MyCStr(dsTemp1("Comment")))
            If sComment <> "" Then
              iPos = InStr(sComment, VbCrLf)
              If iPos > 0 Then 'Отказ
                sComment = Mid(sComment, iPos+Len(VbCrLf), Len(sComment))
                sComment = Trim(sComment)
              Else 'Утверждение
                sComment = Replace(sComment, GetFullName(SurnameGN(DelOtherLangFromFolder(sName)), sNameApproved), "")
                sComment = Replace(sComment, "  ", " ")
                sComment = Replace(sComment, "- -", "-")
              End If
            End If
            oPayDox.InsertBookmarkText sComment, "VISATABLECOMMENT"
          End If
        End If
      End If
      dsTemp1.Close
    End If
'Запрос №11 - СТС - end - add20100720



'  Case "VISATABLEPOSITION"
'   	Set dsTemp1 = Server.CreateObject("ADODB.Recordset")
'   	dsTemp1.CursorLocation = 3 'adUseClient  
'   	'dsTemp1.Open "select * from Comments where DocID='" + Trim(Request("DocID")) + "' and (CommentType='VISA' Or CommentType='FILE') order by DateCreation desc", Conn, 3, 1, &H1
'   	dsTemp1.Open "select * from Comments where DocID='" + Trim(Request("DocID")) + "' and CommentType='VISA' order by DateCreation", Conn, 1, 2, &H1
'   	dsTemp1.ActiveConnection = Nothing
'
'	CurrentRSRecordNumber=0
'   	If Not dsTemp1.EOF Then
'   		dsTemp1.MoveLast
'   		'dsTemp1.MoveFirst
'		CurrentRSRecordCount=dsTemp1.RecordCount
'	Else
'		CurrentRSRecordCount=0
'	End If
'AddLogD "VISATABLEPOSITION, CurrentRSRecordCount:"+CStr(CurrentRSRecordCount)
'	bPrint=True
'AddLogD "LOOP START"
'   	iRecTemp=0
'   	Do While Not dsTemp1.EOF And Not dsTemp1.BOF
'   		iRecTemp=iRecTemp+1
'		sC=MyCStr(dsTemp1("Comment"))
'AddLogD "iRecTemp: "+CStr(iRecTemp)
'AddLogD dsTemp1("UserID")+", Comment:"+sC
'
'		If InStr(sC, "Отмена статуса «Согласовано» или «Отказано»")>0 Then
'			Exit Do
'		End If
'
'		'If InStr(sComment, UCase("Согласовано"))>0 Or InStr(sComment, UCase("Отказано"))>0 Then
'		sSpecialInfo=UCase(MyCStr(dsTemp1("SpecialInfo")))
'		sCommentType=UCase(MyCStr(dsTemp1("CommentType")))
'AddLogD "sSpecialInfo:"+sSpecialInfo	
'		'If sSpecialInfo="VISAOK" or sSpecialInfo="VISAOKREFUSE" Or InStr(UCase(sC), UCase("Отмена статуса"))>0 Then
'		If sSpecialInfo="VISAOK" or sSpecialInfo="VISAOKREFUSE" Or (sCommentType="FILE" And (InStr(MyCStr(dsTemp1("Address")), ">+")>0 Or InStr(MyCStr(dsTemp1("Address")), ">-")>0)) Then
'AddLogD "Print"	
'	       If Not bPrint Then
'	       'If True Then
'	   	    	If Not dsTemp1.EOF Then
'					MSWordInsertRowInTable 1
'					MSWordAddBookmarksToTable Array("VISATABLEPOSITION","VISATABLEFIO","VISATABLERESULT", "VISATABLEDATE","VISATABLESIGNATURE", "VISATABLECOMMENT","VISATABLETIME"), 1, 0, 1
'				End If
'       	End If
'			oPayDox.GetUserDetails dsTemp1("UserID"), sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
'	 		oPayDox.InsertBookmarkText sPosition, "VISATABLEPOSITION"
'			iTempComment2=0
'			iTempComment=InStr(UCase(sC), UCase("Согласовано"))
'			If iTempComment>1 Then
'				sD=Replace(Trim(Left(sC, iTempComment-1)), "(",VbCrLf+"(")
'			Else
'				sD=""
'				iTempComment2=InStr(UCase(sC), UCase("Отказано в согласовании"))
'				If iTempComment2>1 Then
'					sD=Replace(Trim(Left(sC, iTempComment2-1)), "(",VbCrLf+"(")
'					sC=Replace(sC, "Отказано в согласовании","Не согласовано")
'				End If
'			End If
'
'			If InStr(sC, ", дней:")>0 Then 
'			If iTempComment>0 Then
'				sC=Trim(Mid(sC, iTempComment))
'AddLogD "sComment 2:"+sComment	
'			ElseIf iTempComment2>0 Then
'				sC=Trim(Mid(sC, iTempComment2))
'AddLogD "sComment 2:"+sComment	
'			End If
'			End If
'			sC=Replace(sC, "| ",VbCrLf)
'
'			sRole=""
'			sUserIDRole=""
'			sUserName=""
'			iTemp3=InStrRev(sC, "/")
'			If iTemp3>1 Then
'				sTemp=Mid(sC+" ", iTemp3+1)
'    			sUserIDRole = oPayDox.GetNextUserIDInList(sTemp, 1)
'				If sUserIDRole<>"" Then
'					sC=Left(sC, iTemp3-1)
'					sUserName=GetName(sTemp)
'				End If
'			End If
'			sUserNameOut=MyCStr(dsTemp1("UserName"))
'			If sUserIDRole<>"" Then
'				sUserNameOut=sUserNameOut+" /"+VbCrLf+sUserName
'			End If
'			
'	 		oPayDox.InsertBookmarkText sUserNameOut, "VISATABLEFIO"
'			
'			sC=Replace(sC, "| ",VbCrLf)
'
'			'If sCommentType="FILE" Then
'			'	If InStr(MyCStr(dsTemp1("Address")), ">+")>0 Then
'			'		sC=DOCS_ReconciliationFile+" "+MyCStr(dsTemp1("Version"))+": "+DOCS_Reconciled+VbCrLf+sC
'			'	End If
'			'	If InStr(MyCStr(dsTemp1("Address")), ">-")>0 Then
'			'		sC=DOCS_ReconciliationFile+" "+MyCStr(dsTemp1("Version"))+": "+DOCS_Refused+VbCrLf+sC
'			'	End If
'			'End If
'	 		oPayDox.InsertBookmarkText sC+VbCrLf, "VISATABLERESULT"
'	 		oPayDox.InsertBookmarkText MyDate(dsTemp1("DateCreation")), "VISATABLEDATE"
'			If iTempComment>0 Then
'				MSWordInsertUserSignatureToDoc IIF(sUserIDRole<>"", sUserIDRole, dsTemp1("UserID")), "VISATABLESIGNATURE"
'			End If
'			oPayDox.InsertBookmarkText sD, "VISATABLETIME"
'			
'			sUserIDTemp=dsTemp1("UserID")
'			iPosTemp = dsTemp1.Bookmark
'			dsTemp1.MoveFirst
'			sUserComment=DOCS_No
'			Do While Not dsTemp1.EOF
'				If dsTemp1("UserID")=sUserIDTemp And Not (UCase(MyCStr(dsTemp1("SpecialInfo")))="VISAOK" or UCase(MyCStr(dsTemp1("SpecialInfo")))="VISAOKREFUSE") Then
'					sUserComment=DOCS_Yes
'					'Exit Do
'				End If
'				If dsTemp1("UserID")=sUserIDTemp Then
'				If iPosTemp <> dsTemp1.Bookmark Then
'					sSpecialInfo1=UCase(MyCStr(dsTemp1("SpecialInfo")))
'					sCommentType1=UCase(MyCStr(dsTemp1("CommentType")))
'					If sSpecialInfo1="VISAOK" or sSpecialInfo1="VISAOKREFUSE" And sCommentType1<>"FILE" Then
'AddLogD "dsTemp1.Delete:"+MyCStr(dsTemp1("Comment"))
'						dsTemp1.Delete
'						'dsTemp1.Update
'					End If
'				End If
'				End If
'				If dsTemp1.EOF Or dsTemp1.BOF Then
'					Exit Do
'				End If
'				dsTemp1.MoveNext
'			Loop
'			dsTemp1.Bookmark = iPosTemp
'			oPayDox.InsertBookmarkText sUserComment, "VISATABLECOMMENT"
'
'			bPrint=False
'		Else
'AddLogD "No print"	
'       End If
'       'dsTemp1.MoveNext
'       dsTemp1.MovePrevious
'   	Loop
'   	dsTemp1.Close
'AddLogD "LOOP END"
'
'   	sUserList=oPayDox.NotYetAgreedList(dsDoc("ListToReconcile"), dsDoc("ListReconciled"))
'AddLogD "sUserList:"+sUserList
'
'   	If sUserList<>"" Then
'   	'If False Then
'   	iTempComment = 1
'	Do While True
'    	sUser = oPayDox.GetUserIDFromList(sUserList, iTempComment)
'AddLogD "sUser:"+sUser
'    	If sUser = "" Then 
'    		Exit Do
'    	End If
'       If Not bPrint Then
'AddLogD "MSWordInsertRowInTable"
'			MSWordInsertRowInTable 1
'			MSWordAddBookmarksToTable Array("VISATABLEPOSITION","VISATABLEFIO","VISATABLERESULT", "VISATABLEDATE","VISATABLESIGNATURE", "VISATABLECOMMENT","VISATABLETIME"), 1, 0, 1
'		End If
'		oPayDox.GetUserDetails sUser, sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
'	 	oPayDox.InsertBookmarkText sPosition, "VISATABLEPOSITION"
'	 	oPayDox.InsertBookmarkText SurnameGN(sName), "VISATABLEFIO"
'		bPrint=False
'	Loop
'   	End If 'sUserList<>"" Then
'AddLogD "VISATABLEPOSITION End"
'
'  Case "APROVALTABLEPOSITION"
'   	Set dsTemp1 = Server.CreateObject("ADODB.Recordset")
'   	dsTemp1.Open "select * from Comments where DocID='" + Trim(Request("DocID")) + "' and (CommentType='APROVAL' Or CommentType='VISA') order by DateCreation desc", Conn, 3, 1, &H1
'	CurrentRSRecordNumber=0
'   	If Not dsTemp1.EOF Then
'   		dsTemp1.MoveLast
'   		dsTemp1.MoveFirst
'		CurrentRSRecordCount=dsTemp1.RecordCount
'	Else
'		CurrentRSRecordCount=0
'	End If
'	bPrint=True
'   	Do While Not dsTemp1.EOF
'		bPrint=False
'		sComment=UCase(MyCStr(dsTemp1("Comment")))
'		sCommentType=UCase(MyCStr(dsTemp1("CommentType")))
'		If (InStr(sComment, UCase("Утверждено -"))>0 Or InStr(sComment, UCase("Отказано"))>0) And sCommentType="APROVAL" Then
'	       If Not bPrint Then
'	   	    	If Not dsTemp1.EOF Then
'					MSWordInsertRowInTable 2
'					MSWordAddBookmarksToTable Array("APROVALTABLEPOSITION","APROVALTABLEFIO", "APROVALTABLERESULT", "APROVALTABLEDATE","APROVALTABLESIGNATURE", "APROVALTABLECOMMENT"), 2, 0, 1
'				End If
'       	End If
'			oPayDox.GetUserDetails dsTemp1("UserID"), sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
'	 		oPayDox.InsertBookmarkText sPosition, "APROVALTABLEPOSITION"
'	 		oPayDox.InsertBookmarkText MyDate(dsTemp1("DateCreation")), "APROVALTABLEDATE"
'
'			sC=MyCStr(dsTemp1("Comment"))
'			iTempComment=InStr(UCase(sC), UCase(DOCS_Approved))
'AddLogD "iTempComment:"+CStr(iTempComment)			
'			If iTempComment>=1 Then
'				sC=Left(sC, iTempComment+9)
'			Else
'				sC=""
'			End If
'
'	 		'oPayDox.InsertBookmarkText MyCStr(dsTemp1("Comment")), "APROVALTABLERESULT"
'	 		oPayDox.InsertBookmarkText sC, "APROVALTABLERESULT"
'	 		oPayDox.InsertBookmarkText MyCStr(dsTemp1("UserName")), "APROVALTABLEFIO"
'			MSWordInsertUserSignatureToDoc dsTemp1("UserID"), "APROVALTABLESIGNATURE"
'			
'			'APROVALTABLECOMMENT
'			sUserIDTemp=dsTemp1("UserID")
'			iPosTemp = dsTemp1.Bookmark
'			dsTemp1.MoveFirst
'			sUserComment=DOCS_No
'			Do While Not dsTemp1.EOF
'				If dsTemp1("UserID")=sUserIDTemp And UCase(MyCStr(dsTemp1("CommentType")))<>"APROVAL" And Not (UCase(MyCStr(dsTemp1("SpecialInfo")))="VISAOK" or UCase(MyCStr(dsTemp1("SpecialInfo")))="VISAOKREFUSE") Then
'					sUserComment=DOCS_Yes
'					Exit Do
'				End If
'				dsTemp1.MoveNext
'			Loop
'			dsTemp1.Bookmark = iPosTemp
'			oPayDox.InsertBookmarkText sUserComment, "APROVALTABLECOMMENT"
'			
'			bPrint=False
'       End If
'       dsTemp1.MoveNext
'   	Loop
'   	dsTemp1.Close
'   	
'  Case "COMMENTVISATABLEPOSITION"
'   	Set dsTemp1 = Server.CreateObject("ADODB.Recordset")
'   'dsTemp1.Open "select * from Comments where DocID='" + Trim(Request("DocID")) + "' and CommentType='VISA' Or CommentType='APROVAL' order by DateCreation desc", Conn, 3, 1, &H1
'	oPayDox.ShowListComments Trim(Request("DocID")), dsTemp1	
'	CurrentRSRecordNumber=0
'   	If Not dsTemp1.EOF Then
'   		dsTemp1.MoveLast
'   		dsTemp1.MoveFirst
'		CurrentRSRecordCount=dsTemp1.RecordCount
'	Else
'		CurrentRSRecordCount=0
'	End If
'AddLogD "COMMENTVISATABLEPOSITION dsTemp1.RecordCount:"+CStr(dsTemp1.RecordCount)
'	bPrint=True
'   	Do While Not dsTemp1.EOF
'		sCommentType=UCase(MyCStr(dsTemp1("CommentType")))
'   		If sCommentType="VISA" Or sCommentType="APROVAL" Then
'		bPrint=False
''		sCommentType=UCase(MyCStr(dsTemp1("CommentType")))
''		sComment=UCase(MyCStr(dsTemp1("Comment")))
''		If InStr(sComment, UCase("Утверждено"))>0 Or InStr(sComment, UCase("Отказано"))>0 Then
'		sSpecialInfo=UCase(MyCStr(dsTemp1("SpecialInfo")))
'		If Not (sSpecialInfo="VISAOK" or sSpecialInfo="VISAOKREFUSE") And InStr(UCase(MyCStr(dsTemp1("Comment"))), UCase(DOCS_Approved))<=0 Then
'	       If Not bPrint Then
'	   	    	If Not dsTemp1.EOF Then
'					MSWordInsertRowInTable 3
'					MSWordAddBookmarksToTable Array("COMMENTVISATABLEPOSITION","COMMENTVISATABLEFIO","COMMENTVISATABLERESULT", "COMMENTVISATABLEDATE", "COMMENTVISATABLESIGNATURE"), 3, 0, 1
'				End If
'       	End If
'       	
'       	IF Left(MyCStr(dsTemp1("SpecialInfo"))+"    ", 4)="RESP" Then
'       		sPrefixTemp="  ^  "
'       	Else
'       		sPrefixTemp=""
'       	End If
'			oPayDox.GetUserDetails dsTemp1("UserID"), sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
'	 		oPayDox.InsertBookmarkText sPosition, "COMMENTVISATABLEPOSITION"
'	 		oPayDox.InsertBookmarkText MyDate(dsTemp1("DateCreation")), "COMMENTVISATABLEDATE"
'	 		oPayDox.InsertBookmarkText sPrefixTemp+MyCStr(dsTemp1("Comment")), "COMMENTVISATABLERESULT"
'	 		oPayDox.InsertBookmarkText MyCStr(dsTemp1("UserName")), "COMMENTVISATABLEFIO"
'			'MSWordInsertUserSignatureToDoc dsTemp1("UserID"), "COMMENTVISATABLESIGNATURE"
'			bPrint=False
'       End If
'   		End If 'sCommentType="VISA" Then
'       dsTemp1.MoveNext
'   	Loop
'   	dsTemp1.Close
%>
