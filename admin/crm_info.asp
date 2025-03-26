<!--#include file="..\_inc\functions.asp" -->
<%
	Call pageFunctions(90,5)

	Dim strKey      'as String
	Dim strClass    'as String

	If not blnAdmin Then Call logError(2,1)

	strTitle = getIDS("IDS_ServerInformation")

	Call DisplayHeader(1)

	Response.Write("<div id=""contentDiv"" class=""dvBorder"" style=""height=:" & intScreenH - 60 & "px;"">" & vbCrLf & vbCrLf & _
			"<table border=0 cellspacing=0 cellpadding=5 width=""100%"">" & vbCrLf & _
			"<tr class=""hRow"" width=""25%""><th class=""hFont"">" & getIDS("IDS_VariableName") & "</th><th class=""hFont"" width=""75%"">" & getIDS("IDS_Value") & "</th></tr>" & vbCrLf)

	For Each strKey in Request.ServerVariables

		strClass = toggleRowColor(strClass)

		Response.Write("<tr class=""" & strClass & """><td class=""bFont"">" & strKey & "</td><td class=""dFont"">")
		If Request.ServerVariables(strKey) = "" then Response.Write("&nbsp;") Else Response.Write("Disabled for Security")
		Response.Write("</td></tr>" & vbCrLf)
	Next

	Response.Write("<tr><td colspan=""2"">&nbsp;</td></tr>" & vbCrLf & "<tr class=""hRow""><th class=""hFont"" colspan=""3"">" & getIDS("IDS_DatabaseInfo") & "</th></tr>" & vbCrLf)

	For Each strKey in objConn.Properties

		strClass = toggleRowColor(strClass)

		Response.Write("<tr class=""" & strClass & """><td class=""bFont"">" & strKey.Name & "</td><td class=""dFont"">")
		If strKey.Value = "" then Response.Write("&nbsp;") Else Response.Write("Disabled for Security")
		Response.Write("</td></tr>" & vbCrLf)
	Next

	Response.Write("</table>" & vbCrLf & vbCrLf & "</div>" & vbCrLf)

	Call DisplayFooter(1)
%>