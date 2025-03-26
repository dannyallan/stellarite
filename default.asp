<!--#include file="_inc\functions.asp" -->
<!--#include file="_inc\sql\sql_login.asp" -->

<script language="JavaScript" type="text/javascript" runat="server">
function getTimeOffset() {
	var lnOffset = new Date();
	lnOffset = lnOffset.getTimezoneOffset();
	lnOffset = -(lnOffset);
	return(lnOffset);
}
</script>

<%
	Dim strUsername         'as String
	Dim strPassword         'as String
	Dim blnPopup            'as Boolean
	Dim strLastIP           'as String
	Dim strDest             'as String

	strTitle = "Stellarite " & getIDS("IDS_Login")
	strAction = valString(Request.QueryString,2048,0,0)

	If Request.Form.Count > 0 Then
		strUsername = valString(Request.Form("txtUsername"),20,0,0)
		strPassword = valString(Request.Form("txtPassword"),35,0,0)
		blnPopup = valNum(Request.Form("chkPopup"),0,0)
		strLastIP = Request.ServerVariables("REMOTE_ADDR")
	Else
		blnPopup = Request.Cookies("Popup")
		If blnPopup = "" Then blnPopup = 1
	End If

'	strUsername = "admin"
'	strPassword = "5f4dcc3b5aa765d61d8327deb882cf99"
'	blnPopup = 0
'	strLastIP = "127.0.0.1"
'	strAction = "login"

	Response.Cookies("Popup") = blnPopup
	Response.Cookies("Popup").Expires = DateAdd("d",30,Date)

	Call DisplayHeader(3)

	If intMode <> 2 and strUsername <> "" and strPassword <> "" and strAction = "login" Then

		strDest = "default.asp"
		Session.Contents.RemoveAll()

		Set objRS = objConn.Execute(doLogin(strUsername,strPassword,strLastIP,CInt(Application("av_LoginAttempts"))))

		If objRS.BOF and objRS.EOF Then
			Call logError(0,0)
			Session("ErrorMsg") = getIDS("IDS_ErrorUnspecified")
		Else
			If CByte(objRS.fields(0).value) = 1 Then                        'Username Error
				Session("ErrorMsg") = getIDS("IDS_MsgIncorrectLogin")
			Elseif CByte(objRS.fields(0).value) = 2 Then                    'Admin Locked Out
				Call logError(4,0)
				Session("ErrorMsg") = getIDS("IDS_MsgLockedOutAdmin")
			Elseif CByte(objRS.fields(0).value) = 3 Then                    'User Locked Out
				Call logError(4,0)
				Session("ErrorMsg") = getIDS("IDS_MsgLockedOutUser")
			Elseif CByte(objRS.fields(0).value) = 4 Then                    'Password Error
				Session("ErrorMsg") = getIDS("IDS_MsgIncorrectLogin")
			Elseif CByte(objRS.fields(0).value) = 5 Then                    'Unspecified
				Call logError(0,0)
				Session("ErrorMsg") = getIDS("IDS_ErrorUnspecified")
			Else
				Session("UserId")       = CLng(objRS.fields(1).value)
				Session("UserName")     = CStr(objRS.fields(2).value)
				Session("Admin")        = objRS.fields(3).value
				Session("ChngPass")     = objRS.fields(4).value
				Session("Member")       = CStr(objRS.fields(5).value & "")
				Session("Permissions")  = CStr(objRS.fields(6).value & "")
				Session("PortalList")   = CStr(objRS.fields(7).value & "")
				Session("PortalCount")  = UBound(Split(Session("PortalList"),"1"))
				Session("TimeOffset")   = 0-valNum(Request.Form("tz"),2,0)-getTimeOffset()
				Session("screenH")      = valNum(Request.Form("sh"),2,0)
				Session("screenW")      = valNum(Request.Form("sw"),2,0)

				If Request.Cookies("Popup") = "" Then Call doRedirect("require.asp?prob=ck")

				'Disable the checking for screen resolution
				'If Session("screenH") < 500 or Session("screenW") < 700 Then Call doRedirect("require.asp?prob=sr")

				For i = 1 to bytModCount
					If not Application("av_Module" & i & "On") Then
						Session("Permissions") = Left(Session("Permissions"),i-1) & "0" & Mid(Session("Permissions"),i+1,1)
						Session("Member") = Left(Session("Member"),i-1) & "0" & Mid(Session("Member"),i+1,1)
					End If
				Next

				strDest = valString(Request.Form("hdnDest"),2048,0,0)
				If strDest = "" Then
					Session("Destination") = "main.asp"
				Else
					Session("Destination") = strDest
					strDest = valString("http://localhost" & strDest,2048,1,2)
				End If

				If blnPopup = 0 Then
					strDest = Session("Destination")
					Session.Contents.Remove("Destination")
				Else
					strDest = "default.asp?new"
				End If
			End If
		End If
		Call doRedirect(strDest)
	End If

	If strAction = "logout" Then

		Session.Contents.RemoveAll()
		Call doRedirect("default.asp")

	Elseif intMode = 2 Then
%>

<br /><br />

<table border=0 width="100%" height="90%"><tr><td valign="middle" align="center">

<table border=0 cellspacing=10 cellpadding=0 width=300 height=150>
  <tr class="hRow">
	<td class="dFont">
	<span class="hFont"><% =getIDS("IDS_Maintenance") %></span><br /><br />
	<% =getIDS("IDS_MsgMaintDown") %><br />
	</td>
  </tr>
</table>

</td></tr></table>

<script language="JavaScript" type="text/javascript">
	closeWindow(null);
</script>
<%
	Elseif Session("UserId") <> "" and Session("Destination") <> "" and strAction = "new" Then

		strDest = Session("Destination")
		Session.Contents.Remove("Destination")
%>
<script language="JavaScript" type="text/javascript">
	window.name = "sw_Login";
	var oWin = window.open("<% =strDest %>","sw_CRM","status=0,left=0,top=0,menubar=0,resizable,width=<% =Session("screenW") %>,height=<% =Session("screenH") %>");
	if (oWin.opener == null)
		oWin.opener = self;
	oWin.focus();
</script>

<table border=0 width="100%" height="90%"><tr><td valign="middle" align="center">

<table border=0 cellspacing=10 cellpadding=0 width=300 height=150>
  <tr class="hRow">
	<td class="hFont"><center><% =getIDS("IDS_ThankYou") %>.<br />
	<a href="default.asp"><% =getIDS("IDS_LogBackIn") %>.</a></center></td>
  </tr>
</table>

</td></tr></table>
<%
	Else

		Session.Abandon

		strDest = valString(Request.QueryString("dest"),2048,0,0)
%>

<script language="JavaScript" type="text/javascript" src="common/js/md5.js"></script>

<table border=0 width="100%" height="90%"><tr><td valign="middle" align="center">

<form name="frmLogin" method="post" action="<% =strCRMURL %>default.asp?login" autocomplete="off" onSubmit="document.forms[0].txtPassword.value = calcMD5(txtPassword.value);">
<% =getHidden("hdnDest",strDest) %>
<% =getHidden("tz","") %>
<% =getHidden("sw","") %>
<% =getHidden("sh","") %>
<% If Session("ErrorMsg") <> "" Then Response.Write(Session("ErrorMsg") & "<br /><br />" & vbCrLf) %>

<table border=0 cellspacing=10 cellpadding=0 width=300 height=150 class="hRow">
  <tr>
	<td><% =getLabel(getIDS("IDS_UserName"),"txtUserName") %></td>
	<td><% =getTextField("txtUserName","oText","",20,20,"") %></td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_Password"),"txtPassword") %></td>
	<td><% =getPassword("txtPassword","oText","",20,35,"") %></td>
  </tr>
<% If strDest = "" Then %>
<tr><td colspan="2"><% =getCheckbox("chkPopup",blnPopup,"") %>&nbsp;&nbsp;<% =getLabel(getIDS("IDS_MsgPopup"),"chkPopup") %></td></tr>
<% End If %>
  <tr>
	<td colspan=2 align=center><% =getSubmit("btnSubmit",getIDS("IDS_Login"),100,"S","onClick=""doHidden();""") %></td>
  </tr>
</table>

<table border=0 cellspacing=10 cellpadding=0 width=350>
  <tr>
    <td class="dfont">
      <p>You can access this demonstration with any of the following accounts.  If you
      have any questions, please contact us at <a href="http://www.stellarite.com/">Stellarite, Inc</a>.</p>

      <table border=0 width="100%">
        <tr><td class="dfont">UserName: <span class="dlabel">admin</span></td><td class="dfont">Password: <span class="dlabel">password</span></td></tr>
        <tr><td class="dfont">UserName: <span class="dlabel">sales</span></td><td class="dfont">Password: <span class="dlabel">password</span></td></tr>
        <tr><td class="dfont">UserName: <span class="dlabel">services</span></td><td class="dfont">Password: <span class="dlabel">password</span></td></tr>
        <tr><td class="dfont">UserName: <span class="dlabel">support</span></td><td class="dfont">Password: <span class="dlabel">password</span></td></tr>
        <tr><td class="dfont">UserName: <span class="dlabel">quality</span></td><td class="dfont">Password: <span class="dlabel">password</span></td></tr>
        <tr><td class="dfont">UserName: <span class="dlabel">finance</span></td><td class="dfont">Password: <span class="dlabel">password</span></td></tr>
        <tr><td class="dfont">UserName: <span class="dlabel">knowledge</span></td><td class="dfont">Password: <span class="dlabel">password</span></td></tr>
      </table>

      <p>All names, addresses and phone numbers in this demonstration have been generated from
      random data sources.  The client list has been generated from the Fortune 500 index and
      has no relation to real sales or contacts.</p>
    </td>
  </tr>
</table>

</form>

</td></tr></table>

<script language="JavaScript" type="text/javascript">
	if (self != top)
		top.location.reload();

	if ((window.opener != null) && (window.opener.name == 'sw_CRM'))
		closeWindow("refresh");

	function doHidden() {
		document.forms[0].tz.value = new Date().getTimezoneOffset();
		if ((document.forms[0].chkPopup != null) && (document.forms[0].chkPopup.checked)) {
			document.forms[0].sw.value = eval(screen.availWidth-10);
			document.forms[0].sh.value = eval(screen.availHeight-30);
		} else {
			window.name = "sw_CRM";
			document.forms[0].sw.value = getWindowWidth();
			document.forms[0].sh.value = getWindowHeight();
		}
	}
</script>

<%
	End If

	Call DisplayFooter(1)
%>