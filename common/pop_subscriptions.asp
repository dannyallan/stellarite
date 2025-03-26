<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_email.asp" -->
<%
	Call pageFunctions(0,1)

	Dim strClass	'as String
	Dim intTickets	'as Integer
	Dim intBugs		'as Integer
	Dim lngDelete	'as Long

	strTitle = Application("IDS_EmailSubscriptions")
	lngDelete = valNum(Request.Form("hdnDelete"),3,0)

	Set objRS = objConn.Execute(getEmailSub(lngUserId,5,0))
	If objRS.BOF and objRS.EOF Then intTickets = 0 Else intTickets = 1
	Set objRS = objConn.Execute(getEmailSub(lngUserId,6,0))
	If objRS.BOF and objRS.EOF Then intBugs = 0 Else intBugs = 1

	If valNum(Request.Form("hdnSubmit"),0,0) = 1 Then

		If lngDelete <> "" Then
			objConn.Execute(delEmailSub(lngUserId,lngDelete,bytMod,lngModId))
		End If

		If valNum(Request.Form("chkTickets"),0,0) = 1 Then
			If intTickets = 0 Then
				objConn.Execute(insertEmailSub(lngUserId,5,0))
				intTickets = 1
			End If
		Else
			If intTickets = 1 Then
				objConn.Execute(delEmailSub(lngUserId,0,5,0))
				intTickets = 0
			End If
		End If

		If valNum(Request.Form("chkBugs"),0,0) = 1 Then
			If intBugs = 0 Then
				objConn.Execute(insertEmailSub(lngUserId,6,0))
				intBugs = 1
			End If
		Else
			If intBugs = 1 Then
				objConn.Execute(delEmailSub(lngUserId,0,6,0))
				intBugs = 0
			End If
		End if
	End If

	Set objRS = objConn.Execute(getSubscriptions(lngUserId))
	If not (objRS.BOF and objRS.EOF) Then arrRS = objRS.GetRows()

	Call DisplayHeader(3)
	Call showEditHeader(strTitle,"","","","")
%>
<div id="contentDiv" class="dvBorder" style="height:336px;"><br>

<table border=0 cellspacing=0 cellpadding=3 width="100%">
<form name="frmSubscriptions" method="post" action="pop_subscriptions.asp">
<%
	Response.Write(getHidden("hdnSubmit","1"))
	Response.Write(getHidden("hdnDelete",""))

	If pTickets >= 1 Then Response.Write("<tr class=""drow1""><td>" & getLabel(Application("IDS_HotTickets"),"chkTickets") & "</td><td align=right>" & getCheckbox("chkTickets",intTickets,"onClick=""document.forms[0].submit();""") & "</td></tr>" & vbCrLf)
	If pBugs >= 1 Then Response.Write("<tr class=""drow2""><td>" & getLabel(Application("IDS_HotBugs"),"chkBugs") & "</td><td align=right>" & getCheckbox("chkBugs",intBugs,"onClick=""document.forms[0].submit();""") & "</td></tr>" & vbCrLf)


	If isArray(arrRS) Then
		For i = 0 to UBound(arrRS,2)
			strClass = toggleRowColor(strClass)

			Response.Write("<tr class=""" & strClass & """><td class=""dfont"">" & showString(arrRS(arrRS(1,i)+1,i)) & "</td>" & _
					"<td class=""dfont"" align=right><a href=""Javascript:document.forms[0].hdnDelete.value='" & arrRS(0,i) & "';document.forms[0].submit();"">" & _
					"<img src=""../images/del2.gif"" alt=""" & Application("IDS_Delete") & " " & showString(arrRS(arrRS(1,i)+1,i)) & """ border=0 height=16 width=16></a></td></tr>" & vbCrLf)

		Next
	End If
%>
</form>
</table>
</div>

<div id="footerDiv" class="dvFooter">
<%
	Response.Write(getIconCancel())
%>
</div>

<%

	Call DisplayFooter(3)
%>
