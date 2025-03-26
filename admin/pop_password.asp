<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_users.asp" -->
<%
	Call pageFunctions(0,1)

	Dim strMsg	'as String

	strTitle = Application("IDS_PasswordEdit")

	If not (lngUserId=lngRecordId) Then Call logError(2,1)

	If Session("ChngPass") = 1 Then
		strMsg = Application("IDS_PasswordChange")
	Else
		strMsg = ""
	End If


	If strDoAction = "edit" Then

		Set objRS = objConn.Execute(getUserDetails(lngRecordId))

		If valString(Request.Form("txtOldPassword"),35,1,0) <> objRS.fields("U_Password").value Then
			strMsg = Application("IDS_MsgIncorrectPassword")
		Else
			objConn.Execute(updateUserPass(lngRecordId,valString(Request.Form("txtNewPassword"),35,1,0)))
			Session.Contents.Remove("ChngPass")
			Call closeWindow(strOpenerURL)
		End if
	End If

	strIncHead = "<script language=""Javascript"" src=""../common/js/md5.js""></script>"

	Call DisplayHeader(3)
	Call showEditHeader(strTitle,"","","","")
%>

<script language="Javascript">
	function checkPwd() {
		var msg = "";
		var newPass = document.forms[0].txtNewPassword.value;
		var oldPass = document.forms[0].txtOldPassword.value;
		var conPass = document.forms[0].txtConPassword.value;
		var illegalChars = /[\W]/;

		if (newPass != conPass) {
			msg += "<% =Application("IDS_MsgPasswordMatch") %>\n";
		}
		if (newPass == oldPass) {
			msg += "<% =Application("IDS_MsgPasswordSame") %>\n";
		}
		if (newPass.length < 6) {
			msg += "<% =Application("IDS_MsgPasswordLength") %>\n";
		}
		if (illegalChars.test(newPass)) {
			msg += "<% =Application("IDS_MsgPasswordChars") %>\n";
		}
		if (((newPass.search(/[a-z]+/)==-1) || (newPass.search(/[A-Z]+/)==-1) || (newPass.search(/[0-9]+/)==-1))) {
			msg += "<% =Application("IDS_MsgPasswordSecure") %>\n";
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

<div id="contentDiv" class="dvBorder" style="height:330px;"><br>

<table border=0 width="100%">
<form name="frmPassword" method="post" action="pop_password.asp?id=<% =lngRecordId %>">
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
<% =getHidden("hdnWinOpen","") %>
    <tr>
      <td class="dlabel"><% =Application("IDS_Name") %></td>
      <td class="dfont"><% =strFullName %></td>
    </tr>
    <tr><td class="dfont" colspan=2>&nbsp;</td></tr>
    <tr>
      <td><% =getLabel(Application("IDS_PasswordOld"),"txtOldPassword") %></td>
      <td><% =getPassword("txtOldPassword","mText","",20,35,"") %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_PasswordNew"),"txtNewPassword") %></td>
      <td><% =getPassword("txtNewPassword","mText","",20,35,"") %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_PasswordConfirm"),"txtConPassword") %></td>
      <td><% =getPassword("txtConPassword","mText","",20,35,"") %></td>
    </tr>
    <tr><td class="wfont" colspan=2><br><% =strMsg %>&nbsp;</td></tr>
</form>
</table>

</div>

<div id="footerDiv" class="dvFooter">
<%
	Response.Write(getIcon("Javascript:checkPwd();","S","save.gif",Application("IDS_Save")))
	Response.Write(getIconCancel())
%>
</div>

<%
	Call DisplayFooter(3)
%>