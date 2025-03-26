<!--#include file="..\_inc\functions.asp" -->
<%
	Call pageFunctions(90,5)

	Dim strKey      'as String
	Dim strSubKey   'as String
	Dim strClass    'as String

	If not blnAdmin Then Call logError(2,1)

	strTitle = getIDS("IDS_ServerVariables")

	Call DisplayHeader(1)

	Response.Write("<div id=""contentDiv"" class=""dvBorder"" style=""height:" & intScreenH - 60 & "px;"">" & vbCrLf & vbCrLf & _
			"<table border=0 cellspacing=0 cellpadding=5 width=""100%"">" & vbCrLf)

	Response.Write("<tr class=""hRow""><th class=""hFont"" colspan=2>" & getIDS("IDS_Cookies") & "</th></tr>" & vbCrLf)

	For Each strKey in Request.Cookies

		If Request.Cookies(strKey).HasKeys then
			For Each strSubKey in Request.Cookies(strKey)

				strClass = toggleRowColor(strClass)

				Response.Write("<tr class=""" & strClass & """><td class=""bFont"" width=""20%"">" & strKey & " (" & strSubKey & ")</td><td class=""dFont"" width=""80%"">")
				If Request.Cookies(strKey)(strSubKey) = "" then Response.Write("&nbsp;") Else Response.Write(showString(Request.Cookies(strKey)(strSubKey)))
				Response.Write("</td></tr>" & vbCrLf)
			Next
		Else
			strClass = toggleRowColor(strClass)

			Response.Write("<tr class=""" & strClass & """><td class=""bFont"">" & strKey & "</td><td class=""dFont"">")
			If Request.Cookies(strKey) = "" then Response.Write("&nbsp;") Else Response.Write(showString(Request.Cookies(strKey)))
			Response.Write("</td></tr>" & vbCrLf)
		End if
	Next

	Response.Write("<tr class=""hRow""><th class=""hFont"" colspan=2>" & getIDS("IDS_SessionVariables") & "</th></tr>" & vbCrLf)

	For Each strKey in Session.Contents

		strClass = toggleRowColor(strClass)

		Response.Write("<tr class=""" & strClass & """><td class=""bFont"">" & strKey & "</td><td class=""dFont"">")
		If Session.Contents(strKey) <> "" Then Response.Write(showString(Session.Contents(strKey))) Else Response.Write("&nbsp;")
		Response.Write("</td></tr>" & vbCrLf)
	Next

	Response.Write("<tr class=""hRow""><th class=""hFont"" colspan=2>" & getIDS("IDS_ApplicationVariables") & "</th></tr>" & vbCrLf)

	For Each strKey in Application.Contents

		strClass = toggleRowColor(strClass)

		Response.Write("<tr class=""" & strClass & """><td class=""bFont"">" & strKey & "</td><td class=""dFont"">")

		If isArray(Application.Contents(strKey)) Then
			Response.Write("{Array}")
		Elseif Application.Contents(strKey) = "" Then
			Response.Write("---")
		Else
			Response.Write(showString(Application.Contents(strKey)))
		End If

		Response.Write("</td></tr>" & vbCrLf)
	Next

	Response.Write("</table>" & vbCrLf & vbCrLf & "</div>" & vbCrLf)

	Call DisplayFooter(1)
%>