<%
Sub DisplayHeader(fType)

'	Response.write("<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.1//EN"" ""http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd"">" & vbCrLf)
'	Response.write("<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">" & vbCrLf)
'	Response.write("<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3.org/TR/html4/loose.dtd"">" & vbCrLf)
	Response.write("<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 3.2 Final//EN"">" & vbCrLf)
	Response.Write("<html lang=""" & strLanguage & """>" & vbCrLf & vbCrLf & "<head>" & vbCrLf)

	If strModItem <> "" and strModName <> strTitle and bytRealMod <> 0 Then
		Response.Write("<title>" & strModItem & ": " & showString(strTitle) & "</title>" & vbCrLf)
	Else
		Response.Write("<title>" & showString(strTitle) & "</title>" & vbCrLf)
	End If

	Response.Write("<meta http-equiv=""Content-Type"" content=""text/html; charset=windows-1252"" />" & vbCrLf & _
			"<script language=""JavaScript"" type=""text/javascript"">" & vbCrLf & _
				vbTab & "var sLang=""" & strLanguage & """;" & vbCrLf & _
				vbTab & "var sDateFormat=""" & Application("av_DateFormat") & """;" & vbCrLf & _
				vbTab & "var sCRMUrl=""" & Application("av_CRMDir") & """;" & vbCrLf & _
				vbTab & "var iMode=" & intMode & ";" & vbCrLf & _
			"</script>" & vbCrLf & _
			"<script language=""JavaScript"" type=""text/javascript"" src=""" & Application("av_CRMDir") & "common/js/crm.js""></script>" & vbCrLf)

	If fType = 1 or fType = 0 Then

		strTab = Request.Cookies("ModTab")(CStr(bytRealMod))
		If strTab = "" Then strTab = "m" & bytRealMod & "Tab0"

		Response.Write("<script language=""JavaScript"" type=""text/javascript"" src=""" & Application("av_CRMDir") & "common/js/menu.js""></script>" & vbCrLf)
		If Session("ChngPass") = 1 Then Response.Write("<script language=""JavaScript"" type=""text/javascript"">" & getEditURL("W","?id="&lngUserId) & "</script>" & vbCrLf)

	ElseIf fType = 2 Then

		strTab = valString(Request.QueryString("tab"),7,0,0)
		If Len(strTab) > 5 Then Response.Cookies("ModTab")(CStr(Mid(strTab,2,Instr(strTab,"T")-2))) = strTab

	End If

	If strIncHead <> "" Then Response.Write(strIncHead & vbCrLf)

	Response.Write("<noscript><meta http-equiv=""refresh"" content=""0;URL=" & Application("av_CRMDir") & "require.asp?prob=js"" /></noscript>" & vbCrLf & _
				"<link href=""" & Application("av_CRMDir") & "common/css.asp"" rel=""stylesheet"" type=""text/css"" />" & vbCrLf & _
				"</head>" & vbCrLf & vbCrLf)

	Response.Write("<body bgcolor=""#FFFFFF"" leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"" scroll=""no"">" & vbCrLf & vbCrLf)

	If fType = 1 or (fType = 0 and Abs(1-valNum(Request.QueryString("menu"),1,0)) = 0) Then
		Call DisplayMenu()
		bytMenu = 1
	Else
		bytMenu = 0
	End If

End Sub
%>