<%
Sub DisplayFooter(fType)

	If Err <> 0 Then Call logError(0,0)

	Call closeConn()

	'Response.Write("<div id=""timerDiv"" style=""display:none""><form name=""frmTimer"" method=""post"">" & getHidden("hdnTimer",Timer-intStartTime) & "</form></div>" & vbCrLf & vbCrLf)

	Response.Write("</body>" & vbCrLf & "</html>")
End Sub
%>