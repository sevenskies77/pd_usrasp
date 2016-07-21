<% 'Insert your Case branches here
		Case "UserFieldDate1Long" 'Your Range name
	 		oPayDox.InsertRangeDate MyDateLong(ds("UserFieldDate1")), "UserFieldDate1Long" 'Your code
		Case "NameResponsibleIDentification" 'Your Range name
			oPayDox.GetUserDetails GetUserID(ds("NameResponsible")), sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
	 		oPayDox.InsertRangeDate sIDentification, "NameResponsibleIDentification" 'Your code
		Case "NameResponsibleIDNo" 'Your Range name
			oPayDox.GetUserDetails GetUserID(ds("NameResponsible")), sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
	 		oPayDox.InsertRangeDate sIDNo, "NameResponsibleIDNo" 'Your code
		Case "NameResponsibleIDIssuedBy" 'Your Range name
			oPayDox.GetUserDetails GetUserID(ds("NameResponsible")), sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
	 		oPayDox.InsertRangeDate sIDIssuedBy, "NameResponsibleIDIssuedBy" 'Your code
		Case "NameResponsibleIDIssueDate" 'Your Range name
			oPayDox.GetUserDetails GetUserID(ds("NameResponsible")), sName, sPhone, sEMail, sICQ, sDepartment, sPartnerName, sPosition, sIDentification, sIDNo, sIDIssuedBy, dIDIssueDate, dIDExpDate, dBirthDate, sCorporateIDNo, sAddInfo, sComment
	 		oPayDox.InsertRangeDate MyDateLong(dIDIssueDate), "NameResponsibleIDIssueDate" 'Your code
		Case "XXX" 'Your Range name
			 oPayDox.InsertRangeDate "This text will be inserted", "XXX" 'Your code
		Case "YYY" 'Your Range name
			 oPayDox.InsertRangeDate "This text will be inserted too", "YYY" 'Your code
	    Case "MIKRON_AMW"
	         oPayDox.InsertRangeText "This text will be inserted too", "MIKRON_AMW"
%>