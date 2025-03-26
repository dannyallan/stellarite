<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_admin.asp" -->
<%
	Call pageFunctions(90,5)

	Dim strInfo	'as String

	If lngRecordId = "" Then Call logError(3,1)
	strTitle = Application("IDS_EditMessage")

	If strDoAction = "edit" Then

		strInfo = valString(Request.Form("txtInfo"),255,0,4)

		Application.Lock
		Call setConfigValue("av_ModuleMsg"&lngRecordId,strInfo)
		Application.Unlock

		Call closeWindow(strOpenerURL)
	Else
		strInfo = Application("av_ModuleMsg" & lngRecordId)
	End If

	Call DisplayHeader(3)
	Call showEditHeader(strTitle,"","","","")
%>
<div id="contentDiv" class="dvBorder" style="height:330px;"><br>

<table border=0 cellspacing=5 width="100%">
<form name="frmAdmin" method="post" action="pop_msg.asp?id=<% =lngRecordId %>">
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
<% =getHidden("hdnWinOpen","") %>
  <tr>
    <td>
      <% =getLabel(Application("IDS_CharLimit255"),"txtInfo") %>
      <br>
      <% =getTextArea("txtInfo","oMemo",strInfo,"100%",15,"") %>
    </td>
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

<%
	Call DisplayFooter(3)
%>