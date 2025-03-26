<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_users.asp" -->
<%
	Call pageFunctions(90,5)

	Dim strUsrMem       'as String
	Dim strUsrPerm      'as String
	Dim strName         'as String
	Dim bytAdmin        'as Byte
	Dim bytPass         'as Byte
	Dim bytLock         'as Byte
	Dim blnThisMod		'as Boolean
	Dim bytThisMod		'as Byte

	strTitle = getIDS("IDS_Permissions")

	If lngRecordId = "" Then Call logError(3,1)

	If strDoAction = "edit" Then

		For i = 1 to bytModCount
			strUsrMem = strUsrMem & valNum(Request.Form("chk" & i),0,0)
			strUsrPerm = strUsrPerm & valNum(Request.Form("sel" & i),1,1)
		Next

		If blnAdmin Then
			bytAdmin = valNum(Request.Form("chkAdmin"),0,0)
			bytPass     = valNum(Request.Form("chkPass"),0,0)
			If valNum(Request.Form("chkLock"),0,0) = 1 Then bytLock = CInt(Application("av_LoginAttempts")) Else bytLock = 0
		Else
			bytAdmin = "NULL"
		End If

		objConn.Execute(upDateUserPerm(lngRecordId,strUsrMem,strUsrPerm,bytAdmin,bytPass,bytLock))
		Call doRedirect("profile.asp?id="&lngRecordId)
	Else
		Set objRS = objConn.Execute(getUserDetails(lngRecordId))

		If not (objRS.BOF and objRS.EOF) then
			strUsrMem     = objRS.fields("U_Member").value
			strUsrPerm    = objRS.fields("U_Permissions").value

			If objRS.fields("U_Admin").value = 1 Then bytAdmin = 1 Else bytAdmin = 0
			If objRS.fields("U_ChangePassword").value = 1 Then bytPass = 1 Else bytPass = 0
			If objRS.fields("U_LoginAttempts").value >= CInt(Application("av_LoginAttempts")) Then bytLock = 1 Else bytLock = 0

			strName = objRS.fields(0).value
		End If
	End If

	Function showPermissions(fTitle,fDefault)

		showPermissions = "<select name=""" & fTitle & """ id=""" & fTitle & """ class=""oByte"" onChange=""doChange();"">" & vbCrLf & _
				vbTab & "<option value=0" & getDefault(0,fDefault,0) & ">" & getIDS("IDS_AccessDeny") & "</option>" & vbCrLf & _
				vbTab & "<option value=1" & getDefault(0,fDefault,1) & ">" & getIDS("IDS_AccessRead") & "</option>" & vbCrLf & _
				vbTab & "<option value=2" & getDefault(0,fDefault,2) & ">" & getIDS("IDS_New") & "</option>" & vbCrLf & _
				vbTab & "<option value=3" & getDefault(0,fDefault,3) & ">" & getIDS("IDS_Edit") & "</option>" & vbCrLf & _
				vbTab & "<option value=4" & getDefault(0,fDefault,4) & ">" & getIDS("IDS_Delete") & "</option>" & vbCrLf & _
				vbTab & "<option value=5" & getDefault(0,fDefault,5) & ">" & getIDS("IDS_Administrator") & "</option>" & vbCrLf & _
				"    </select>"
	End Function

	strTitle = getIDS("IDS_Permissions") & " - " & strName
	Call DisplayHeader(1)
	Call showEditHeader(strTitle,"","","","")
%>
<form name="frmAdmin" method="post" action="edit_permissions.asp?id=<% =lngRecordId %>">
<div id="contentDiv" class="dvBorder" style="height:<% =intScreenH-170 %>px;"><br />

<table border=0 cellspacing=0 cellpadding=3 width="99%">
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
  <% If blnAdmin Then %>
  <tr>
	<td><% =getLabel(getIDS("IDS_CRMAdministrator"),"chkAdmin") %></td>
	<td colspan=2><% =getCheckbox("chkAdmin",bytAdmin,"") %></td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_PasswordChange"),"chkPass") %></td>
	<td colspan=2><% =getCheckbox("chkPass",bytPass,"") %></td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_LockAccount"),"chkLock") %><br /><br /></td>
	<td colspan=2><% =getCheckbox("chkLock",bytLock,"") %><br /><br /></td>
 </tr>
  <% End If %>
  <tr class="hRow">
	<th align=left class="bFont"><% =getIDS("IDS_Group") %></td>
	<th align=left class="bFont"><% =getIDS("IDS_Member") %></td>
	<th align=center class="bFont"><% =getIDS("IDS_Permissions") %></td>
  </tr>
<%
	For i = 1 to bytModCount
		blnThisMod = valNum(Mid(strUsrMem,i,1),0,0)
		bytThisMod = valNum(Mid(strUsrPerm,i,1),1,0)

		If Application("av_Module" & i & "On") and (CByte(valNum(Mid(Session("Permissions"),i,1),1,0)) = 5 or blnAdmin) Then
			Response.Write("  <tr><td>" & getLabel(getIDS("IDS_ModName" & i),"chk" & i) & "</td><td align=center>" & getCheckbox("chk" & i,blnThisMod,"") & "</td><td align=right>" & showPermissions("sel" & i,bytThisMod) & "</td></tr>" & vbCrLf)
		Else
			Response.Write(getHidden("chk" & i,blnThisMod) & getHidden("sel" & i,bytThisMod) & vbCrLf)
		End If
	Next
%>
</table>

</div>

<div id="footerDiv" class="dvFooter">
<%
	Response.Write(getIconSave("edit"))
	Response.Write(getIconCancel("back"))
%>
</div>
</form>

<%
	Call DisplayFooter(1)
%>