<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_admin.asp" -->
<%
	Call pageFunctions(90,5)

	Dim objEmail		'as Object
	Dim blnEnableEmail	'as Boolean
	Dim bytEmailType	'as Byte
	Dim strEmailHost	'as String
	Dim strEmailFrom	'as String
	Dim lngEmailPort	'as Long

	On Error Resume Next

	strTitle = Application("IDS_EmailOptions")

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

		Call setConfigValue("av_EnableEmail",blnEnableEmail)
		Call setConfigValue("av_EmailType",bytEmailType)
		Call setConfigValue("av_EmailHost",strEmailHost)
		Call setConfigValue("av_EmailFrom",strEmailFrom)
		Call setConfigValue("av_EmailPort",lngEmailPort)

		Application.Unlock

		Call closeWindow(strOpenerURL)
	Else
		blnEnableEmail	= Application("av_EnableEmail")
		bytEmailType	= Application("av_EmailType")
		strEmailHost	= Application("av_EmailHost")
		strEmailFrom	= Application("av_EmailFrom")
		lngEmailPort	= Application("av_EmailPort")
	End If

	Call DisplayHeader(3)
	Call showEditHeader(strTitle,"","","","")
%>
<div id="contentDiv" class="dvBorder" style="height:330px;"><br>

<table border=0 cellspacing=5 width="100%">
<form name="frmAdmin" method="post" action="pop_email.asp">
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
<% =getHidden("hdnWinOpen","") %>
  <tr>
    <td><% =getLabel(Application("IDS_EmailNotification"),"chkEmail") %></td>
    <td><% =getCheckbox("chkEmail",blnEnableEmail,"onClick=""doClassChange();""") %></td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_EmailType"),"selEmailType") %></td>
    <td>
      <select name="selEmailType" id="selEmailType" class="oByte" onChange="doChange();" style="width:195px;">
        <option value="1"<% =getDefault(0,bytEmailType,"1") %>>CDONTS</option>
        <option value="2"<% =getDefault(0,bytEmailType,"2") %>>ASPEmail</option>
        <option value="3"<% =getDefault(0,bytEmailType,"3") %>>JMail</option>
      </select>
    </td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_EmailHost"),"txtEmailHost") %></td>
    <td><% =getTextField("txtEmailHost","oText",strEmailHost,30,255,"") %></td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_EmailPort"),"txtEmailPort") %></td>
    <td><% =getTextField("txtEmailPort","oLong",lngEmailPort,4,4,"") %></td>
  </tr>
  <tr>
    <td colspan=2>&nbsp;</td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_EmailFrom"),"txtEmailFrom") %></td>
    <td><% =getTextField("txtEmailFrom","oEmail",strEmailFrom,30,255,"") %></td>
  </tr>
</form>
</table>

</div>

<div id="footerDiv" class="dvFooter">
<%
	Response.Write(getIconSave("edit"))
	Response.Write(getIconCancel())
%>
</div>

<script language="Javascript">
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
	Call DisplayFooter(3)
%>

