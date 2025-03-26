<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_email.asp" -->
<%
	Call pageFunctions(0,1)

	Dim strClass    'as String
	Dim intTickets  'as Integer
	Dim intBugs     'as Integer
	Dim lngDelete   'as Long

	strTitle = getIDS("IDS_EmailSubscriptions")
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

	Call DisplayHeader(1)
	Call showEditHeader(strTitle,"","","","")
%>
<form name="frmSubscriptions" method="post" action="edit_subscriptions.asp">
<div id="contentDiv" class="dvBorder" style="height:<% =intScreenH-170 %>px;"><br />

<table border=0 cellspacing=0 cellpadding=3 width="100%">
<%
	Response.Write(getHidden("hdnSubmit","1"))
	Response.Write(getHidden("hdnDelete",""))

	If pTickets >= 1 Then Response.Write("<tr class=""dRow1""><td>" & getLabel(getIDS("IDS_HotTickets"),"chkTickets") & "</td><td align=right>" & getCheckbox("chkTickets",intTickets,"onClick=""document.forms[0].submit();""") & "</td></tr>" & vbCrLf)
	If pBugs >= 1 Then Response.Write("<tr class=""dRow2""><td>" & getLabel(getIDS("IDS_HotBugs"),"chkBugs") & "</td><td align=right>" & getCheckbox("chkBugs",intBugs,"onClick=""document.forms[0].submit();""") & "</td></tr>" & vbCrLf)


	If isArray(arrRS) Then
		For i = 0 to UBound(arrRS,2)
			strClass = toggleRowColor(strClass)

			Response.Write("<tr class=""" & strClass & """><td class=""dFont"">" & showString(arrRS(arrRS(1,i)+1,i)) & "</td>" & _
					"<td class=""dFont"" align=right>" & getIconImport(4,"Javascript:document.forms[0].hdnDelete.value='" & arrRS(0,i) & "';document.forms[0].submit();",showString(arrRS(arrRS(1,i)+1,i))) & "</td></tr>" & vbCrLf)

		Next
	End If
%>
</table>
</div>

<div id="footerDiv" class="dvFooter">
<%
	Response.Write(getIconCancel("../main.asp"))
%>
</div>
</form>

<%
	Call DisplayFooter(1)
%>
