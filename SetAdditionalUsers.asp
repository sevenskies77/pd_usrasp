<!--#INCLUDE VIRTUAL="\TextANSI.txt" -->
<!--#INCLUDE VIRTUAL="\Common.asp" -->
<%
Sub SetAdditionalUsers
  Set dsDocs = CreateObject("ADODB.Recordset")
  sSQL = "SELECT * FROM Docs WHERE (Department LIKE N'СИТРОНИКС%')"
  AddLogD "QQQQ"&sSQL
  dsDocs.Open sSQL, Conn, 1, 3, &H1
  AddLogD "QQQQ"&dsDocs("DocID")

  Response.Write "<TABLE BORDER=1 COLS=3 BGCOLOR=white align=left><TR><TH>DocID</TH><TH>OLD AdditionalUsers</TH><TH>NEW AdditionalUsers</TH></TR>"

  Do While not dsDocs.EOF
    sAdditionalUsers = GetAllUpperChiefsOfUsersFromList(MyCStr(dsDocs("NameResponsible")) + "; " + MyCStr(dsDocs("ListToReconcile")), MyCStr(dsDocs("BusinessUnit")))
    If Trim(Replace(Replace(MyCStr(dsDocs("AdditionalUsers")), ";", ""), "-", "")) <> "" Then
	  sAdditionalUsers = sAdditionalUsers+VbCrLf+MyCStr(dsDocs("AdditionalUsers"))
    End If
    sAdditionalUsers = DeleteUserDoublesInList(sAdditionalUsers)
    Response.Write "<TR>"
    Response.Write "<TD><FONT size=3 face=""courier new"">"+MyCStr(dsDocs("DocID"))+"</FONT></TD>"
    Response.Write "<TD><FONT size=3 face=""courier new"">"+Replace(Replace(MyCStr(dsDocs("AdditionalUsers")), ">", "&gt;"), "<", "&lt;")+";"+"</FONT></TD>"

    dsDocs("AdditionalUsers") = sAdditionalUsers
    dsDocs.Update
    Response.Write "<TD><FONT size=3 face=""courier new"" color = green>"+Replace(Replace(MyCStr(dsDocs("AdditionalUsers")), ">", "&gt;"), "<", "&lt;")+";"+"</FONT></TD>"
    Response.Write "</TR>"
	dsDocs.MoveNext
  Loop

  Response.Write "</TABLE>"

  dsDocs.Close
End Sub


'On Error Resume Next
SetAdditionalUsers
%>
