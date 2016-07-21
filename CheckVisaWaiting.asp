<!--#include virtual="\TextANSI.txt" -->
<!--#include virtual="\Common.asp" -->
<%
  Set rs = CreateObject("ADODB.Recordset")
  sSQL = "select *, Docs.DocID as DocDocID, Docs.DateCreation as DocDateCreation from Docs left join Comments on (Docs.DocID=Comments.DocID and SpecialInfo = 'VISAWAITING') where IsActive <> 'N' and (StatusCompletion is NULL or (StatusCompletion <> '0' and StatusCompletion <> '1')) and CharIndex('<', ListToReconcile) > 0 and CharIndex('-<', ListReconciled) = 0 and CharIndex('#!', ListToReconcile) = 0 and NameApproved is not NULL and NameApproved = '' and Comments.DocID is NULL"
  rs.Open sSQL, Conn, 3, 1, &H1
  Response.Write "<table border=1><tr><th>DocID</th><th>ListToReconcile</th><th>ListReconciled</th><th>DateCreation</th><th>Department</th></tr>"
  Do While not rs.EOF
    If not IsReconciliationCompleteWithOptions(rs("ListToReconcile"), rs("ListReconciled")) Then
		Response.Write "<tr><td>"&_
		HTMLEncode(MyCStr(rs("DocDocID")))&"</td><td>"&_
		HTMLEncode(MyCStr(rs("ListToReconcile")))&"</td><td>"&_
		HTMLEncode(MyCStr(rs("ListReconciled")))&"&nbsp</td><td>"&_
		HTMLEncode(MyDate(rs("DocDateCreation")))&"</td><td>"&_
		HTMLEncode(MyCStr(DelOtherLangFromFolder(rs("Department"))))&"&nbsp</td></tr>"
	End If
	rs.MoveNext
  Loop
  Response.Write "</table>"
  rs.Close
%>
