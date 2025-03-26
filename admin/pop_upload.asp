<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_admin.asp" -->
<%
	Call pageFunctions(90,5)

	Dim objUpload		'as Object
	Dim blnEnableUpload	'as Boolean
	Dim bytUploadType	'as Integer
	Dim strUploadPath	'as String
	Dim strUploadURL	'as String
	Dim strUploadLog	'as String
	Dim lngUploadLimit	'as String

	strTitle = Application("IDS_UploadOptions")

	If not blnAdmin Then Call logError(2,1)

	If strDoAction = "edit" Then

		blnEnableUpload = CStr(valNum(Request.Form("chkUpload"),0,0))
		bytUploadType = CStr(valNum(Request.Form("selUploadType"),1,1))
		strUploadPath = valString(Request.Form("txtUploadPath"),255,blnEnableUpload,2)
		strUploadURL = valString(Request.Form("txtUploadURL"),255,blnEnableUpload,2)
		strUploadPath = valString(Request.Form("txtUploadPath"),255,blnEnableUpload,2)
		strUploadLog = valString(Request.Form("txtUploadLog"),255,blnEnableUpload,0)
		lngUploadLimit = CStr(valNum(Request.Form("txtUploadLimit"),3,blnEnableUpload))

		If strUploadURL <> "" and Right(strUploadURL,1) <> "/" Then
			strUploadURL = strUploadURL & "/"
		End If

		On Error Resume Next
		Select Case bytUploadType
			Case "1"
				Set objUpload = Server.CreateObject("Persits.Upload")
			Case "2"
				Set objUpload = Server.CreateObject("AspSmartUpLoad.SmartUpLoad")
			Case "3"
				Set objUpload = Server.CreateObject("NET2DATABASE.AspFileUp")
			Case "4"
				Set objUpload = Server.CreateObject("SoftArtisans.FileUp")
		End Select

		If IsObject(objUpload) Then Set objUpload = Nothing
		If Err <> 0 Then Call sendBack(Application("IDS_MsgUploadComponent"))

		'PureASP has a maximum freeware upload size of 10MB.  Do not allow the upload
		'size to exceed this or uploads will cause 500 Server errors.

		If bytUploadType = 0 and lngUploadLimit > 10240000 Then
			Call sendBack(Application("IDS_MsgUploadSize"))
		End If

		Application.Lock

		Call setConfigValue("av_Upload",blnEnableUpload)
		Call setConfigValue("av_UploadType",bytUploadType)
		Call setConfigValue("av_UploadPath",strUploadPath)
		Call setConfigValue("av_UploadURL",strUploadURL)
		Call setConfigValue("av_UploadLog",strUploadLog)
		Call setConfigValue("av_UploadLimit",lngUploadLimit)

		Application.Unlock

		Call closeWindow(strOpenerURL)
	Else
		blnEnableUpload	= showString(Application("av_Upload"))
		bytUploadType	= showString(Application("av_UploadType"))
		strUploadPath	= showString(Application("av_UploadPath"))
		strUploadURL	= showString(Application("av_UploadURL"))
		strUploadLog	= showString(Application("av_UploadLog"))
		lngUploadLimit	= Application("av_UploadLimit")
	End If

	Call DisplayHeader(3)
	Call showEditHeader(strTitle,"","","","")
%>
<div id="contentDiv" class="dvBorder" style="height:330px;"><br>

<table border=0 cellspacing=5 width="100%">
<form name="frmAdmin" method="post" action="pop_upload.asp">
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
<% =getHidden("hdnWinOpen","") %>
  <tr>
    <td><% =getLabel(Application("IDS_UploadEnabled"),"chkUpload") %></td>
    <td><% =getCheckbox("chkUpload",blnEnableUpload,"onClick=""doClassChange();""") %></td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_UploadType"),"selUploadType") %></td>
    <td>
      <select name="selUploadType" id="selUploadType" class="oByte" onChange="doChange();" style="width:195px;">
        <option value="0"<% =getDefault(0,bytUploadType,0) %>>Pure ASP Upload</option>
        <option value="1"<% =getDefault(0,bytUploadType,1) %>>ASPUpload</option>
        <option value="2"<% =getDefault(0,bytUploadType,2) %>>aspSmartUpload</option>
        <option value="3"<% =getDefault(0,bytUploadType,3) %>>Net2Database</option>
        <option value="4"<% =getDefault(0,bytUploadType,4) %>>SA FileUp</option>
      </select>
    </td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_UploadPath"),"txtUploadPath") %></td>
    <td><% =getTextField("txtUploadPath","oLink",strUploadPath,30,255,"") %></td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_UploadURL"),"txtUploadURL") %></td>
    <td><% =getTextField("txtUploadURL","oLink",strUploadURL,30,255,"") %></td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_UploadLog"),"txtUploadLog") %></td>
    <td><% =getTextField("txtUploadLog","oLink",strUploadLog,30,255,"") %></td>
  </tr>
  <tr>
   <td><% =getLabel(Application("IDS_UploadLimit"),"txtUploadLimit") %></td>
    <td><% =getTextField("txtUploadLimit","oLong",lngUploadLimit,30,255,"") %></td>
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
		if (document.forms[0].chkUpload.checked) {
			document.forms[0].txtUploadPath.className = "mLink";
			document.forms[0].txtUploadURL.className = "mLink";
			document.forms[0].txtUploadLog.className = "mLink";
			document.forms[0].txtUploadLimit.className = "mLong";
		}
		else {
			document.forms[0].txtUploadPath.className = "oLink";
			document.forms[0].txtUploadURL.className = "oLink";
			document.forms[0].txtUploadLog.className = "oLink";
			document.forms[0].txtUploadLimit.className = "oLong";
		}
	}
</script>

<%
	Call DisplayFooter(3)
%>

