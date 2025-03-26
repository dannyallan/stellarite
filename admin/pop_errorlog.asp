<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\timezone.asp" -->
<!--#include file="..\_inc\sql\sql_admin.asp" -->
<%
	Call pageFunctions(90,5)

	Dim blnDetailed			'as Boolean
	Dim strLogFile 			'as String
	Dim strEmails			'as String

	strTitle = Application("IDS_ErrorOptions")

	If not blnAdmin Then Call logError(2,1)

	If strDoAction = "edit" Then


		blnDetailed = valNum(Request.Form("selDetailed"),0,0)
		strLogFile = valString(Request.Form("txtLogFile"),255,0,2)
		strEmails = valString(Request.Form("txtEmail"),255,0,0)

		Application.Lock

		Call setConfigValue("av_ErrorDetailed",blnDetailed)
		Call setConfigValue("av_ErrorLog",strLogFile)
		Call setConfigValue("av_ErrorEmail",strEmails)

		Application.Unlock

		Call closeWindow(strOpenerURL)
	Else
		blnDetailed		= Application("av_ErrorDetailed")
		strLogFile		= Application("av_ErrorLog")
		strEmails		= Application("av_ErrorEmail")
	End If

	Call DisplayHeader(3)
	Call showEditHeader(strTitle,"","","","")
%>
<div id="contentDiv" class="dvBorder" style="height:330px;"><br>

<table border=0 cellspacing=5 width="100%">
<form name="frmAdmin" method="post" action="pop_errorlog.asp">
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
<% =getHidden("hdnWinOpen","") %>
  <tr>
    <td colspan=2><p class="dfont"><% =Application("IDS_MsgErrorLogs") %></p></td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_ErrorLogLevel"),"selDetailed") %></td>
    <td>
      <select name="selDetailed" id="selDetailed" class="oBool" style="width:195px;">
        <option value="0"<% =getDefault(0,0,blnDetailed) %>><% =Application("IDS_ErrorLogLevel0") %></option>
        <option value="1"<% =getDefault(0,1,blnDetailed) %>><% =Application("IDS_ErrorLogLevel1") %></option>
      </select>
    </td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_ErrorLog"),"txtLogFile") %></td>
    <td><% =getTextField("txtLogFile","oLink",strLogFile,30,255,"") %></td>
  </tr>
<% If Application("av_EnableEmail") = "1" Then %>
  <tr>
    <td><% =getLabel(Application("IDS_Email"),"txtEmail") %></td>
    <td><% =getTextField("txtEmail","oText",strEmails,30,255,"") %></td>
  </tr>
<% End If %>
</form>
</table>

</div>

<div id="footerDiv" class="dvFooter">
<%
	Response.Write(getIconSave("edit"))
	Response.Write(getIconCancel())
%>
</div>

<%
	Call DisplayFooter(3)
%>