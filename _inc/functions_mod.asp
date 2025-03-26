<!--#include file="sql\sql_modules.asp" -->
<%
Function makeTabs(f_strTabs)
	Dim f_arrTabs1
	Dim f_arrTabs2
	Dim f_intCount
	Dim f_strTemp

	f_arrTabs1 = split(f_strTabs,"||")

	Response.Write("<table border=0 cellspacing=0 cellpadding=0><tr>" & vbCrLf)

	For f_intCount = 0 to Ubound(f_arrTabs1)

		f_arrTabs2 = split(f_arrTabs1(f_intCount),"|")

		If strTab = "m" & bytRealMod & "Tab" & f_intCount Then
			f_strTemp = "eTab"
			strTabURL = f_arrTabs2(1) & "&tab=m" & bytRealMod & "Tab" & f_intCount
			makeTabs = Replace(f_arrTabs2(0)," ","_")
		Else
			f_strTemp = "dTab"
		End if

		Response.Write(vbTab & "<td class=""" & f_strTemp & """ id=""m" & bytRealMod & "Tab" & f_intCount & """ onClick=""doSetTabColor('m" & bytRealMod & "Tab" & f_intCount & "','" & f_arrTabs2(1) & "&tab=m" & bytRealMod & "Tab" & f_intCount & "','" & Replace(f_arrTabs2(0)," ","_") & "');"">" & vbCrLf & _
				vbTab & "<img src=""../images/lefttab.gif"" alt="""" height=18 width=4 align=absmiddle />" & vbCrLf & _
				vbTab & f_arrTabs2(0) & vbCrLf & _
				vbTab & "<img src=""../images/righttab.gif"" alt="""" height=18 width=4 align=absmiddle /></td>" & vbCrLf & _
				vbTab & "<td>" & getSpacer(3,3) & "</td>" & vbCrLf)

	Next

	Response.Write("</tr></table>" & vbCrLf & vbCrLf)
End Function


Sub showToolBar()

	Dim fFile, fTemp

	fFile = Request.ServerVariables("PATH_INFO")

	If strModItem = getIDS("IDS_Event") Then fTemp = "&m=" & bytMod & "&mid=" & lngModId

	Response.Write("<form name=""frmTabs"" method=""post"" action=""" & fFile & "?id=" & lngRecordId & fTemp & """>" & vbCrLf & _
			"<table border=0 cellspacing=0 width=""100%"">" & vbCrLf & _
			getHidden("hdnClicked",strTab) & vbCrLf & _
			"  <tr class=""hRow"">" & vbCrLf & _
			"    <td class=""tFont""><img src=""../images/" & strModImage & ".gif"" alt=""" & strModItem & """ width=32 height=32 hspace=10 align=absmiddle />" & strTitle & "</td>" & vbCrLf & _
			"    <td align=right>" & vbCrLf)

	If CLng(lngPrevId) <> 0 Then
		Response.Write(getIconPrev(fFile & "?id=" & lngPrevId & fTemp))
	End If

	If CLng(lngNextId) <> 0 Then
		Response.Write(getIconNext(fFile & "?id=" & lngNextId & fTemp))
	Else
		Response.Write(getSpacer(1,28))
	End If

	Response.Write(vbTab & getSpacer(3,3) & vbCrLf)

	If intPerm >= 2 Then
		If strModItem = getIDS("IDS_Event") Then
			Response.Write(getIconNew(getEditURL(50,"?" & Mid(fTemp,2))))
		Else
			Response.Write(getIconNew(getEditURL(bytMod,"")))
		End If
	End If

	If intPerm >= 3 Then
		If strModItem = getIDS("IDS_Event") Then
			Response.Write(getIconEdit(getEditURL(50,"?id=" & lngRecordId & fTemp)))
		Else
			Response.Write(getIconEdit(getEditURL(bytMod,"?id="&lngRecordId)))
		End If
	End If

	If intPerm >= 4 Then
		Response.Write(getIconDelete() & vbTab & getHidden("hdnAction",""))
	End If

	If fTemp = "" Then
		Response.Write(getIconSearch(getSearchURL("?m="&bytMod)))
		If Application("av_EnableEmail") <> "0" Then Response.Write(getIcon("Javascript:openWindow('../common/pop_email.asp?m=" & bytMod & "&mid=" & lngRecordId & "','sw_Email','300','130');","E","email.gif",getIDS("IDS_EmailNotification")))
		Response.Write(getIconPrint("Javascript:openWindow('../common/pop_print.asp?m=" & bytMod & "&mid=" & lngRecordId & "&title=" & Server.URLEncode(strTitle) & "','sw_Print','200','340');"))
	Else
		Response.Write(getIconSearch(getSearchURL("?m=50")))
	End If

	Response.Write("    </td>" & vbCrLf & "  </tr>" & vbCrLf & "</table>" & vbCrLf & "</form>" & vbCrLf)
End Sub
%>