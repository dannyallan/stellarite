<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_admin.asp" -->
<%
	Call pageFunctions(90,5)

	Dim strModules  'as String

	strTitle = getIDS("IDS_EnableModules")

	If not blnAdmin Then Call logError(2,1)

	If strDoAction = "edit" Then

		For i = 1 to bytModCount
			strModules = strModules & valNum(Request.Form("chk" & i),0,0)
		Next

		Application.Lock
		Call setAppVar("av_Modules",strModules)
		Application.Unlock

		Application.Contents.RemoveAll()

		Call closeEdit()
	End If

	Call DisplayHeader(1)
	Call showEditHeader(strTitle,"","","","")
%>
<form name="frmAdmin" method="post" action="edit_modules.asp">
<div id="contentDiv" class="dvBorder" style="height:<% =intScreenH-170 %>px;"><br />

<table border=0 cellspacing=0 cellpadding=4 width="99%">
<% =getHidden("hdnAction","") %>
<%=getHidden("hdnChange","") %>
  <tr class="hRow">
	<th align=left class="bFont"><% =getIDS("IDS_Module") %></td>
	<th align=center class="bFont"><% =getIDS("IDS_Enabled") %></td>
  </tr>
<%
	For i = 1 to bytModCount
		Response.Write("  <tr>" & vbCrLf & _
						"    <td>" & getLabel(getIDS("IDS_ModName" & i),"chk" & i) & "</td>" & vbCrLf & _
						"    <td align=center>" & getCheckbox("chk" & i,valNum(Mid(Application("av_Modules"),i,1),1,0),"") & "</td>" & vbCrLf & _
						"  </tr>" & vbCrLf)
	Next
%>
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