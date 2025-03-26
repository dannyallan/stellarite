<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_admin.asp" -->
<%
	Call pageFunctions(90,5)

	Dim strInfo    'as String

	If lngRecordId = "" Then Call logError(3,1)
	strTitle = getIDS("IDS_EditMessage")

	If strDoAction = "edit" Then

		strInfo = valString(Request.Form("txtInfo"),255,0,4)

		Application.Lock
		Call setAppVar("av_ModuleMsg"&lngRecordId,strInfo)
		Application.Unlock

		Call closeEdit()
	Else
		strInfo = Application("av_ModuleMsg" & lngRecordId)
	End If

	Call DisplayHeader(1)
	Call showEditHeader(strTitle,"","","","")
%>
<form name="frmAdmin" method="post" action="edit_msg.asp?id=<% =lngRecordId %>">
<div id="contentDiv" class="dvBorder" style="height:<% =intScreenH-170 %>px;"><br />

<table border=0 cellspacing=5 width="100%">
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
  <tr>
	<td>
	  <% =getLabel(getIDS("IDS_CharLimit255"),"txtInfo") %>
	  <br />
	  <% =getTextArea("txtInfo","oMemo",strInfo,"100%",15,"") %>
	</td>
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

<%
	Call DisplayFooter(1)
%>