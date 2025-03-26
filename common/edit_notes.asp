<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_notes.asp" -->
<%
	Call pageFunctions(0,2)

	Dim lngEventId      'as Long
	Dim strContactType  'as String
	Dim intContactType  'as Integer
	Dim intPermissions  'as Integer
	Dim strInfo         'as String
	Dim strCreatedBy    'as String
	Dim strModBy        'as String
	Dim datModDate      'as Date
	Dim datCreatedDate  'as Date

	strTitle = getIDS("IDS_JournalEntry")
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

		Session("LastPage") = "i_notes.asp?id=" & lngRecordId & "&m=" & bytMod & "&mid=" & lngModId & "&eid=" & lngEventId
		Call closeEdit()
	Else

		If blnRS and intPerm >= 3 Then

			Set objRS = objConn.Execute(getNote(lngRecordId,intMember))

			If not (objRS.BOF and objRS.EOF) then
				intContactType = objRS.fields("N_ContactType").value
				intPermissions = objRS.fields("N_Permissions").value
				strInfo = objRS.fields("N_Info").value
				strCreatedBy = objRS.fields("CreatedBy").value
				datCreatedDate = objRS.fields("N_CreatedDate").value
				strModBy = objRS.fields("ModBy").value
				datModDate = objRS.fields("N_ModDate").value
			End If

		Elseif blnRS Then
			Call logError(2,1)
		Else
			strCreatedBy = strFullName
			datCreatedDate = Now
			strModBy = strFullName
			datModDate = Now
		End If
		strContactType = getOptionDropDown(150,False,"selContactType","IDS_NoteType",intContactType)
	End If

	strIncHead = getEditorScripts()

	Call DisplayHeader(3)
	Call showEditHeader(strTitle,strCreatedBy,datCreatedDate,strModBy,datModDate)
%>
<form name="frmNote" method="post" action="edit_notes.asp?id=<% =lngRecordId %>&m=<% =bytMod %>&mid=<% =lngModId %>&eid=<% =lngEventId %>">
<div id="contentDiv" class="dvNoBorder" style="height:430px;">

<table border=0 cellspacing=5 cellpadding=0 width="100%">
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
  <tr>
	<td colspan=2>
<% =getTextArea("txtInfo","oText",strInfo,"98%",20,"") %>
	</td>
  </tr>
  <tr>
	<td width="25%"><% =getLabel(getIDS("IDS_Type"),"selContactType") %></td>
	<td width="75%"><% =strContactType %></td>
  </tr>
  <tr>
	<td width="25%"><% =getLabel(getIDS("IDS_Permissions"),"selPermissions") %></td>
	<td width="75%"><% =getPermissionsDropDown(intPermissions,intMember) %></td>
  </tr>
</table>

</div>

<div id="footerDiv" class="dvFooter">
<%
	If blnRS Then
		Response.Write(getIconNew(getEditURL("N","?m="&bytMod&"&mid="&lngModId)))
		If intPerm >= 4 Then Response.Write(getIconDelete())
	End If
	Response.Write(getIcon("Javascript:document.forms[0].onsubmit();confirmAction('" & strAction & "');","S","save.gif",getIDS("IDS_Save")))
	Response.Write(getIconCancel("back"))
%>
</div>
</form>

<script language="JavaScript" type="text/javascript">
	var editor = new HTMLArea("txtInfo");
	editor.generate();
</script>

<%
	Call DisplayFooter(3)
%>

