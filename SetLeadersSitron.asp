<!--#INCLUDE VIRTUAL="\TextANSI.txt" -->
<!--#INCLUDE VIRTUAL="\Common.asp" -->
<%
Sub LoadLeadersToNewStructure
  Set dsUsers = CreateObject("ADODB.Recordset")
  Set dsDeps = CreateObject("ADODB.Recordset")
  Set dsDependants = CreateObject("ADODB.Recordset")
  sSQLUsers = "SELECT UserID, Name, HeadOfDepartments FROM Users WHERE (HeadOfDepartments LIKE N'СИТРОНИКС%')"
  sSQLDepartments = "Select GUID from Departments where Name = N'XXX' or Name = N'XXX/'"
  dsUsers.Open sSQLUsers, Conn, 3, 1, &H1
  Do While not dsUsers.EOF
    dsDeps.Open Replace(sSQLDepartments, "XXX", dsUsers("HeadOfDepartments")), Conn, 3, 1, &H1
    If dsDeps.EOF Then
      Response.Write "<font color = red>ERROR - Department Not Found: '"+CStr(dsUsers("HeadOfDepartments"))+"'</font><BR>"
    Else
      dsDependants.Open "Select * from DepartmentDependants where DependantGUID = '"+dsDeps("GUID")+"'", Conn, 1, 3, &H1
      If dsDependants.EOF Then
        Response.Write "Adding Leader Of Department: '"+CStr(dsUsers("HeadOfDepartments"))+"'<BR>"
        dsDependants.AddNew
        dsDependants("GUID") = oPaydox.GetGUID()
        dsDependants("DependantGUID") = dsDeps("GUID")
        dsDependants("BusinessUnit") = ""
        dsDependants("Leader") = InsertionName(dsUsers("Name"), dsUsers("UserID"))
        dsDependants.Update
      Else
        Response.Write "<font color = green>Already Exists Department: '"+CStr(dsUsers("HeadOfDepartments"))+"'</font><BR>"
      End If
      dsDependants.Close
    End If
    dsDeps.Close
	dsUsers.MoveNext
  Loop
  dsUsers.Close
End Sub


On Error Resume Next
Response.Write "STARTING...<BR>"
LoadLeadersToNewStructure
Response.Write "FINISHING<BR>"

%>
