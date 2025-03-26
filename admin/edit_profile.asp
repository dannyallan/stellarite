<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\states.asp" -->
<!--#include file="..\_inc\timezone.asp" -->
<!--#include file="..\_inc\sql\sql_users.asp" -->

<%
	Call pageFunctions(0,1)

	Dim strFirstName    'as String
	Dim strLastName     'as String
	Dim strAddress1     'as String
	Dim strAddress2     'as String
	Dim strAddress3     'as String
	Dim strCity         'as String
	Dim strState        'as String
	Dim strCountry      'as String
	Dim strZIP          'as String
	Dim strEmail        'as String
	Dim lngPhone1       'as Long
	Dim intExt1         'as Integer
	Dim lngPhone2       'as Long
	Dim intExt2         'as Integer
	Dim intTimeZone     'as Integer
	Dim strUsrMem       'as String
	Dim strUsrPerm      'as String
	Dim strPortal       'as String
	Dim strUserName     'as String
	Dim strPassword     'as String

	strTitle = getIDS("IDS_UserProfile")

	If not (lngUserId = lngRecordId or blnAdmin) Then Call logError(2,1)

	If strDoAction <> "" Then

		strFirstName = valString(Request.Form("txtFirstName"),20,1,0)
		strLastName = valString(Request.Form("txtLastName"),20,1,0)
		strAddress1 = valString(Request.Form("txtAddress1"),60,0,0)
		strAddress2 = valString(Request.Form("txtAddress2"),60,0,0)
		strAddress3 = valString(Request.Form("txtAddress3"),60,0,0)
		strCity = valString(Request.Form("txtCity"),20,0,0)
		strState = valString(Request.Form("selState"),2,0,0)
		strCountry = valString(Request.Form("selCountry"),20,0,0)
		strZip = valString(Request.Form("txtZIP"),7,0,0)
		intTimeZone = valNum(Request.Form("selTimeZone"),2,-1)
		strEmail = valString(Request.Form("txtEmail"),255,1,1)
		lngPhone1 = valNum(Request.Form("txtPhone1"),4,-1)
		intExt1 = valNum(Request.Form("txtExt1"),2,-1)
		lngPhone2 = valNum(Request.Form("txtPhone2"),4,-1)
		intExt2 = valNum(Request.Form("txtExt2"),2,-1)
		strUserName = valString(Request.Form("txtUserName"),20,1,0)
		strPassword = valString(Request.Form("txtNewPassword"),35,1,0)

		If strDoAction = "del" and blnAdmin Then

			Call delUser(lngUserId,lngRecordId)

		ElseIf strDoAction = "edit" Then

			Set objRS = objConn.Execute(getUserIdSql(strUserName))
			If not (objRS.BOF and objRS.EOF) Then
				If lngRecordId <> CLng(objRS.fields(0).value) Then Call sendBack(getIDS("IDS_UserExists"))
			End If

			Call updateUser(lngRecordId,strFirstName,strLastName,strAddress1,strAddress2,strAddress3,strCity, _
					strState,strCountry,strZIP,intTimeZone,strEmail,lngPhone1,intExt1,lngPhone2, _
					intExt2,strUserName,strPassword,blnAdmin)

		ElseIf strDoAction = "new" Then

			Set objRS = objConn.Execute(getUserIdSql(strUserName))
			If not (objRS.BOF and objRS.EOF) Then    Call sendBack(getIDS("IDS_UserExists"))

			Set objRS = objConn.Execute(getDefUserPerm)

			If not (objRS.BOF and objRS.EOF) Then
				strUsrMem = objRS.fields(0).value
				strUsrPerm = objRS.fields(1).value

				If Mid(strUsrMem,1,1) = "1" Then strPortal = strPortal & "1" Else strPortal = strPortal & "0"
				If Mid(strUsrMem,2,1) = "1" Then strPortal = strPortal & "1" Else strPortal = strPortal & "0"
				If Mid(strUsrMem,3,1) = "1" Then strPortal = strPortal & "1" Else strPortal = strPortal & "0"
				If Mid(strUsrMem,4,1) = "1" Then strPortal = strPortal & "10" Else strPortal = strPortal & "00"
				If Mid(strUsrMem,5,1) = "1" Then strPortal = strPortal & "1" Else strPortal = strPortal & "0"
				If Mid(strUsrMem,6,1) = "1" Then strPortal = strPortal & "1" Else strPortal = strPortal & "0"
				If Mid(strUsrMem,7,1) = "1" Then strPortal = strPortal & "11" Else strPortal = strPortal & "01"
				If Mid(strUsrMem,5,1) = "1" Then strPortal = strPortal & "0000010" Else strPortal = strPortal & "0000000"
				If Mid(strUsrMem,6,1) = "1" Then strPortal = strPortal & "100000" Else strPortal = strPortal & "000000"

			End If

			lngRecordId = insertUser(strFirstName,strLastName,strAddress1,strAddress2,strAddress3,strCity, _
					strState,strCountry,strZIP,intTimeZone,strEmail,lngPhone1,intExt1,lngPhone2, _
					intExt2,strUserName,strPassword,strUsrMem,strUsrPerm,strPortal)
		End If
		Call doRedirect("profile.asp?id="&lngRecordId)
	Else

		If blnRS Then
			Set objRS = objConn.Execute(getUser(lngRecordId))

			If not (objRS.BOF and objRS.EOF) then
				strFirstName = objRS.fields("U_FirstName").value
				strLastName = objRS.fields("U_LastName").value
				strAddress1 = objRS.fields("U_Address1").value
				strAddress2 = objRS.fields("U_Address2").value
				strAddress3 = objRS.fields("U_Address3").value
				strCity = objRS.fields("U_City").value
				strState = objRS.fields("U_State").value
				strCountry = objRS.fields("U_Country").value
				strZIP = objRS.fields("U_ZIP").value
				intTimeZone = objRS.fields("U_TimeZone").value
				strEmail = objRS.fields("U_Email").value
				lngPhone1 = objRS.fields("U_Phone1").value
				intExt1 = objRS.fields("U_Ext1").value
				lngPhone2 = objRS.fields("U_Phone2").value
				intExt2 = objRS.fields("U_Ext2").value
				strUserName = objRS.fields("U_UserName").value
			End If
		Else
			intTimeZone = -5
		End If
	End If

	strIncHead = "<script language=""JavaScript"" type=""text/javascript"" src=""../common/js/md5.js""></script>"

	Call DisplayHeader(1)
	Call showEditHeader(strTitle,"","","","")
%>

<% If blnAdmin Then %>

<script language="JavaScript" type="text/javascript">
	function checkPwd() {
		var pass = document.forms[0].txtNewPassword.value;
		if (pass != document.forms[0].txtConPassword.value) {
			alert('<% =getIDS("IDS_MsgPasswordMatch") %>');
			doWarning("txtNewPassword");
			doWarning("txtConPassword");
		}
		else {
			document.forms[0].txtNewPassword.value = calcMD5(pass);
			document.forms[0].txtConPassword.value = calcMD5(pass);
			confirmAction('<% =strAction %>');
		}
	}
</script>

<% End If %>

<form name="frmProfile" method="post" action="edit_profile.asp?id=<% =lngRecordId %>">
<div id="contentDiv" class="dvBorder" style="height:<% =intScreenH-170 %>px;"><br />

<table border=0>
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
	<tr>
	  <td><% =getLabel(getIDS("IDS_Name"),"txtFirstName") %></td>
	  <td>
	  <% =getTextField("txtFirstName","mText",strFirstName,18,20,"") %>
	  <% =getTextField("txtLastName","mText",strLastName,19,20,"") %>
	  </td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_Address"),"txtAddress1") %></td>
	  <td><% =getTextField("txtAddress1","oText",strAddress1,40,60,"") %></td>
	</tr>
	<tr>
	  <td></td>
	  <td><% =getTextField("txtAddress2","oText",strAddress2,40,60,"") %></td>
	</tr>
	<tr>
	  <td></td>
	  <td><% =getTextField("txtAddress3","oText",strAddress3,40,60,"") %></td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_City"),"txtCity") %></td>
	  <td><% =getTextField("txtCity","oText",strCity,40,20,"") %></td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_State"),"selState") %></td>
	  <td><% =getStates(140,"selState",strState) %></td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_Country"),"selCountry") %></td>
	  <td><% =getCountries(140,"selCountrry",strCountry) %>
	  &nbsp;&nbsp;&nbsp;&nbsp;
	  <% =getLabel(getIDS("IDS_ZIP"),"txtZip") & getTextField("txtZIP","oText",strZIP,7,7,"") %></td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_TimeZone"),"selTimeZone") %></td>
	  <td><% =getTimeZone(intTimeZone) %></td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_Email"),"txtEmail") %></td>
	  <td><% =getTextField("txtEmail","mText",strEmail,40,255,"") %></td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_Phone") & " 1","txtPhone1") %></td>
	  <td><% =getTextField("txtPhone1","oPhone",lngPhone1,15,255,"") & " " & getLabel("Ext.","txtExt1") & " " & getTextField("txtExt1","oInt",intExt1,6,255,"") %></td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_Phone") & " 2","txtPhone2") %></td>
	  <td><% =getTextField("txtPhone2","oPhone",lngPhone2,15,255,"") & " " & getLabel("Ext.","txtExt2") & " " & getTextField("txtExt2","oInt",intExt2,6,255,"") %></td>
	</tr>
	<% If blnAdmin Then %>
	<tr><td colspan=2 class="dFont">&nbsp;</td></tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_UserName"),"txtUserName") %></td>
	  <td><% =getTextField("txtUserName","mText",strUserName,20,20,"") %></td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_Password"),"txtNewPassword") %></td>
	  <td><% =getPassword("txtNewPassword","mText","password",35,35,"") %></td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_PasswordConfirm"),"txtConPassword") %></td>
  <td><% =getPassword("txtConPassword","mText","password",35,35,"") %></td>
	</tr>
	<% End If %>
</table>

</div>

<div id="footerDiv" class="dvFooter">
<%
	If blnRS and blnAdmin Then
		Response.Write(getIconNew(getEditURL("U","")))
		Response.Write(getIconDelete())
	End If

	If blnAdmin Then
		Response.Write(getIcon("Javascript:checkPwd();","S","save.gif",getIDS("IDS_Save")))
	Else
		Response.Write(getIconSave("edit"))
	End If

	Response.Write(getIconCancel("back"))
%>
</div>
</form>

<%
	Call DisplayFooter(1)
%>