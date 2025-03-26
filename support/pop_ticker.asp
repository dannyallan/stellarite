<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_support.asp" -->
<%
	Call pageFunctions(5,1)

	Dim strClass		'as String
	Dim strClassUrgent	'as String
	Dim blnRecords		'as Boolean

	strTitle = Application("IDS_ModSupport")
	blnRecords = True

	Call DisplayHeader(3)

	Response.Write(vbCrLf & vbCrLf & "<meta http-equiv=""refresh"" content=""120"">" & vbCrLf & _
			"<div id=""contentDiv"" class=""dvNoBorder"" style=""height:400px;"">" & vbCrLf & _
			"<table border=0 cellpadding=2 cellspacing=0 width=""100%"">" & vbCrLf)

	Set objRS = objConn.Execute(getTicketOwners)
	If not (objRS.BOF and objRS.EOF) Then arrRS = objRS.GetRows()

	If not isArray(arrRS) Then
		Response.Write("<tr class=""hrow""><td class=""hfont"">" & Application("IDS_NoneSpecified") & ".</td></tr>" & vbCrLf)
	Else
		For i = 0 to UBound(arrRS,2)

			If arrRS(1,i) <> "" Then

				If not blnRecords Then
					Response.Write("<tr><td class=""dfont"" colspan=2>" & Application("IDS_NoneSpecified") & ".</td></tr>" & vbCrLf)
				Else
					blnRecords = False
				End If

				If i <> 0 Then Response.Write("<tr><td class=""dfont"" colspan=2>&nbsp;</td></tr>" & vbCrLf)
				Response.Write("<tr class=""hrow""><td class=""hfont"" colspan=2>" & trimString(arrRS(1,i),40) & "</td></tr>" & vbCrLf)

			Else
				blnRecords = True

				strClass = toggleRowColor(strClass)
				If arrRS(3,i) = 1 Then strClassUrgent = "drow3" Else strClassUrgent = strClass

				Response.Write("<tr class=""" & strClassUrgent & """><td class=""dfont"">" & showLink(5,"Javascript:window.opener.location.href='../support/ticket.asp?id=" & arrRS(2,i) & "';window.location.reload();",bigDigitNum(7,arrRS(2,i))) & "</td>" & _
						"<td class=""dfont"">" & showLink(2,"Javascript:window.opener.location.href='../sales/client.asp?id=" & arrRS(4,i) & "';window.location.reload();",trimString(arrRS(5,i),30)) & "</td></tr>" & vbCrLf)
			End If
		Next
	End If

	Response.Write("</table>" & vbCrLf & "</div>" & vbCrLf & vbCrLf)

	Call DisplayFooter(3)
%>
