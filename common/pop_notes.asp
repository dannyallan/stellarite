<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_notes.asp" -->
<%
	Call pageFunctions(0,2)

	Dim lngEventId		'as Long
	Dim strContactType	'as String
	Dim intContactType	'as Integer
	Dim intPermissions	'as Integer
	Dim strInfo			'as String
	Dim strCreatedBy	'as String
	Dim strModBy		'as String
	Dim datModDate		'as Date
	Dim datCreatedDate	'as Date

	strTitle = Application("IDS_JournalEntry")
	lngEventId = valNum(Request.QueryString("eid"),3,0)

	If bytMod = "" or lngModId = "" Then Call logError(3,1)

	If strDoAction <> "" then

		strInfo = valString(Request.Form("txtInfo"),-1,0,5)
		intContactType = valNum(Request.Form("selContactType"),2,-1)
		intPermissions = valNum(Request.Form("selPermissions"),1,1)

		If strDoAction = "del" and intPerm >= 4 Then

			Call delNote(lngUserId,lngRecordId,bytMod,lngModId)

		Elseif strDoAction = "edit" and intPerm >= 3 then

			Call updateNote(lngUserId,lngRecordId,bytMod,lngModId,strInfo,intContactType,intPermissions)

		ElseIf strDoAction = "new" Then

			lngRecordId = insertNote(lngUserId,lngRecordId,bytMod,lngModId,strInfo,intContactType,intPermissions,lngEventId)
		End If

		Call closeWindow(strOpenerURL)
	Else

		If blnRS and intPerm >= 3 Then

			Set objRS = objConn.Execute(getNote(lngRecordId,intMember))

			If not (objRS.BOF and objRS.EOF) then
				intContactType = objRS.fields("N_ContactType").value
				intPermissions = objRS.fields("N_Permissions").value
				strInfo = objRS.fields("N_Info").value
				strCreatedBy = showString(objRS.fields("CreatedBy").value)
				datCreatedDate = objRS.fields("N_CreatedDate").value
				strModBy = showString(objRS.fields("ModBy").value)
				datModDate = objRS.fields("N_ModDate").value
			End If

		Elseif blnRS Then
			Call doRedirect("pop_notes.asp?m=" & bytMod & "&mid=" & lngModId & "eid=" & lngEventId)
		Else
			strCreatedBy = strFullName
			datCreatedDate = Now
			strModBy = strFullName
			datModDate = Now
		End If
		strContactType = getOptionDropDown(150,False,"selContactType","Contact Type",intContactType)
	End If

	strIncHead = getEditorScripts()

	Call DisplayHeader(3)
	Call showEditHeader(strTitle,strCreatedBy,datCreatedDate,strModBy,datModDate)
%>
<div id="contentDiv" class="dvBorder" style="height:430px;"><br>

<table border=0 cellspacing=5 cellpadding=0 width="100%">
<form name="frmNote" method="post" action="pop_notes.asp?id=<% =lngRecordId %>&m=<% =bytMod %>&mid=<% =lngModId %>&eid=<% =lngEventId %>">
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
<% =getHidden("hdnWinOpen","") %>
  <tr>
    <td colspan=2>
<% =getTextArea("txtInfo","oText",strInfo,"98%",23,"") %>
    </td>
  </tr>
  <tr>
    <td width="25%"><% =getLabel(Application("IDS_Type"),"selContactType") %></td>
    <td width="75%"><% =strContactType %></td>
  </tr>
  <tr>
    <td width="25%"><% =getLabel(Application("IDS_Permissions"),"selPermissions") %></td>
    <td width="75%"><% =getPermissionsDropDown(intPermissions,intMember) %></td>
  </tr>
</form>
</table>

</div>

<div id="footerDiv" class="dvFooter">
<%
	If blnRS Then
		Response.Write(getIconNew("pop_notes.asp?m=" & bytMod & "&mid=" & lngModId))
		If intPerm >= 4 Then Response.Write(getIconDelete())
	End If
	Response.Write(getIcon("Javascript:document.forms[0].onsubmit();confirmAction('" & strAction & "');","S","save.gif",Application("IDS_Save")))
	Response.Write(getIconCancel())
%>
</div>

<script language="Javascript">
	var editor = new HTMLArea("txtInfo");
	editor.generate();
</script>

<%
	Call DisplayFooter(3)
%>

