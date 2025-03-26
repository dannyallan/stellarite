<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_users.asp" -->
<%
	Call pageFunctions(0,1)

	Dim strMsg    'as String

	strTitle = getIDS("IDS_PasswordEdit")

	If not (lngUserId=lngRecordId) Then Call logError(2,1)

	If Session("ChngPass") = 1 Then
		strMsg = getIDS("IDS_PasswordChange")
	Else
		strMsg = ""
	End If


	If strDoAction = "edit" Then

		Set objRS = objConn.Execute(getUserDetails(lngRecordId))

		If valString(Request.Form("txtOldPassword"),35,1,0) <> objRS.fields("U_Password").value Then
			strMsg = getIDS("IDS_MsgIncorrectPassword")
		Else
			objConn.Execute(updateUserPass(lngRecordId,valString(Request.Form("txtNewPassword"),35,1,0)))
			Session.Contents.Remove("ChngPass")
			Call closeEdit()
		End if
	End If

	strIncHead = "<script language=""JavaScript"" type=""text/javascript"" src=""../common/js/md5.js""></script>"

	Call DisplayHeader(1)
	Call showEditHeader(strTitle,"","","","")
%>

<script language="JavaScript" type="text/javascript">
	function checkPwd() {
		var msg = "";
		var newPass = document.forms[0].txtNewPassword.value;
		var oldPass = document.forms[0].txtOldPassword.value;
		var conPass = document.forms[0].txtConPassword.value;
		var illegalChars = /[\W]/;

		if (newPass != conPass) {
			msg += "<% =getIDS("IDS_MsgPasswordMatch") %>\n";
		}
		if (newPass == oldPass) {
			msg += "<% =getIDS("IDS_MsgPasswordSame") %>\n";
		}
		if (newPass.length < 6) {
			msg += "<% =getIDS("IDS_MsgPasswordLength") %>\n";
		}
		if (illegalChars.test(newPass)) {
			msg += "<% =getIDS("IDS_MsgPasswordChars") %>\n";
		}
		if (((newPass.search(/[a-z]+/)==-1) || (newPass.search(/[A-Z]+/)==-1) || (newPass.search(/[0-9]+/)==-1))) {
			msg += "<% =getIDS("IDS_MsgPasswordSecure") %>\n";
		}
		if (msg != "") {
			doWarning("txtConPassword");
			doWarning("txtNewPassword");
			alert(msg);
		}
		else {
			document.forms[0].txtOldPassword.value = calcMD5(oldPass);
			document.forms[0].txtNewPassword.value = calcMD5(newPass);
			document.forms[0].txtConPassword.value = calcMD5(conPass);
			confirmAction('<% =strAction %>');
		}
	}
</script>

<form name="frmPassword" method="post" action="edit_password.asp?id=<% =lngRecordId %>">
<div id="contentDiv" class="dvBorder" style="height:<% =intScreenH-170 %>px;"><br />

<table border=0 width="100%">
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
	<tr>
	  <td class="bFont" width=170><% =getIDS("IDS_Name") %></td>
	  <td class="dFont"><% =strFullName %></td>
	</tr>
	<tr><td class="dFont" colspan=2>&nbsp;</td></tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_PasswordOld"),"txtOldPassword") %></td>
	  <td><% =getPassword("txtOldPassword","mText","",20,35,"") %></td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_PasswordNew"),"txtNewPassword") %></td>
	  <td><% =getPassword("txtNewPassword","mText","",20,35,"") %></td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_PasswordConfirm"),"txtConPassword") %></td>
	  <td><% =getPassword("txtConPassword","mText","",20,35,"") %></td>
	</tr>
	<tr><td class="wFont" colspan=2><br /><% =strMsg %>&nbsp;</td></tr>
</table>

</div>

<div id="footerDiv" class="dvFooter">
<%
	Response.Write(getIcon("Javascript:checkPwd();","S","save.gif",getIDS("IDS_Save")))
	Response.Write(getIconCancel("back"))
%>
</div>
</form>

<%
	Call DisplayFooter(1)
%>