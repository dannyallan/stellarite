<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_admin.asp" -->
<%
	Call pageFunctions(90,5)

	Dim objEmail        'as Object
	Dim blnEnableEmail  'as Boolean
	Dim bytEmailType    'as Byte
	Dim strEmailHost    'as String
	Dim strEmailFrom    'as String
	Dim lngEmailPort    'as Long

	On Error Resume Next

	strTitle = getIDS("IDS_EmailOptions")

	If not blnAdmin Then Call logError(2,1)

	If strDoAction = "edit" Then

		blnEnableEmail = CStr(valNum(Request.Form("chkEmail"),0,0))
		bytEmailType = CStr(valNum(Request.Form("selEmailType"),1,blnEnableEmail))
		strEmailHost = valString(Request.Form("txtEmailHost"),255,blnEnableEmail,0)
		strEmailFrom = valString(Request.Form("txtEmailFrom"),255,blnEnableEmail,1)
		lngEmailPort = CStr(valNum(Request.Form("txtEmailPort"),3,blnEnableEmail))

		Select Case bytEmailType
			Case "1"
				Set objEmail = Server.CreateObject("CDO.message")
			Case "2"
				Set objEmail = Server.CreateObject("Persits.MailSender")
			Case "3"
				Set objEmail = Server.CreateObject("JMail.SMTPMail")
		End Select

		If IsObject(objEmail) Then Set objEmail = Nothing
		If blnEnableEmail = "1" and Err <> 0 Then Call sendBack("Component Not Installed")

		Application.Lock

		Call setAppVar("av_EnableEmail",blnEnableEmail)
		Call setAppVar("av_EmailType",bytEmailType)
		Call setAppVar("av_EmailHost",strEmailHost)
		Call setAppVar("av_EmailFrom",strEmailFrom)
		Call setAppVar("av_EmailPort",lngEmailPort)

		Application.Unlock

		Call closeEdit()
	Else
		blnEnableEmail    = Application("av_EnableEmail")
		bytEmailType    = Application("av_EmailType")
		strEmailHost    = Application("av_EmailHost")
		strEmailFrom    = Application("av_EmailFrom")
		lngEmailPort    = Application("av_EmailPort")
	End If

	Call DisplayHeader(1)
	Call showEditHeader(strTitle,"","","","")
%>
<form name="frmAdmin" method="post" action="edit_email.asp">
<div id="contentDiv" class="dvBorder" style="height:<% =intScreenH-170 %>px;"><br />

<table border=0 cellspacing=5>
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
  <tr>
	<td><% =getLabel(getIDS("IDS_EmailNotification"),"chkEmail") %></td>
	<td><% =getCheckbox("chkEmail",blnEnableEmail,"onClick=""doClassChange();""") %></td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_EmailType"),"selEmailType") %></td>
	<td>
	  <select name="selEmailType" id="selEmailType" class="oByte" onChange="doChange();" style="width:195px;">
		<option value="1"<% =getDefault(0,bytEmailType,"1") %>>CDONTS</option>
		<option value="2"<% =getDefault(0,bytEmailType,"2") %>>ASPEmail</option>
		<option value="3"<% =getDefault(0,bytEmailType,"3") %>>JMail</option>
	  </select>
	</td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_EmailHost"),"txtEmailHost") %></td>
	<td><% =getTextField("txtEmailHost","oText",strEmailHost,30,255,"") %></td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_EmailPort"),"txtEmailPort") %></td>
	<td><% =getTextField("txtEmailPort","oLong",lngEmailPort,4,4,"") %></td>
  </tr>
  <tr>
	<td colspan=2>&nbsp;</td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_EmailFrom"),"txtEmailFrom") %></td>
	<td><% =getTextField("txtEmailFrom","oEmail",strEmailFrom,30,255,"") %></td>
  </tr>
</table>

</div>

<div id="footerDiv" class="dvFooter">
<%
	Response.Write(getIconSave("edit"))
	Response.Write(getIconCancel("back"))
%>
</div>
</form>

<script language="JavaScript" type="text/javascript">
	doClassChange();
	function doClassChange() {
		if (document.forms[0].chkEmail.checked) {
			document.forms[0].txtEmailHost.className = "mText";
			document.forms[0].txtEmailPort.className = "mLong";
			document.forms[0].txtEmailFrom.className = "mEmail";
		}
		else {
			document.forms[0].txtEmailHost.className = "oText";
			document.forms[0].txtEmailPort.className = "oLong";
			document.forms[0].txtEmailFrom.className = "oEmail";
		}
	}
</script>

<%
	Call DisplayFooter(1)
%>

