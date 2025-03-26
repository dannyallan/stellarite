<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_admin.asp" -->
<%
	Server.ScriptTimeout = 6000

	Call pageFunctions(90,5)

	Dim intCount        'as Integer

	If not blnAdmin Then Call logError(2,1)

	strTitle = getIDS("IDS_CRMAdministration")
	Call DisplayHeader(1)
	Call showEditHeader(strTitle,"","","","")

%>
<form name="frmAdmin" method="post" action="edit_admin.asp" onSubmit="hideDiv.style.visibility='hidden';">
<div id="contentDiv" class="dvBorder" style="height:<% =intScreenH-170 %>px;"><br />

	<p class="dFont"><% =getIDS("IDS_MsgWaitForEnd") %></p>

	<div id="hideDiv">
		<% =getSubmit("btnReload",getIDS("IDS_ApplicationReload"),200,"1","") %>
		<br /><br />
		<% =getSubmit("btnCount",getIDS("IDS_UpdateCounts"),200,"2","") %>

<%
	If Request.Form("btnCount") = getIDS("IDS_UpdateCounts") Then
		For i = 1 to bytModCount
			Response.Write("<span class=""dFont"">" & getIDS("IDS_Updating") & " " & getIDS("IDS_ModItem" & i) & " ... </span>" & vbCrLf)
			Response.Flush

			Set objRS = objConn.Execute(getModuleList(i))
			If not (objRS.BOF and objRS.EOF) Then

				arrRS = objRS.GetRows()

				For intCount = 0 to UBound(arrRS,2)
					objConn.Execute(updateModuleCount(i,arrRS(0,intCount)))
				Next
			End If

			Response.Write("<span class=""bFont"">" & getIDS("IDS_Done") & "</span><br />" & vbCrLf)
		Next
		Response.Write("<p class=""bFont"">" & getIDS("IDS_Complete") & "</p>" & vbCrLf)

	Elseif Request.Form("btnReload") = getIDS("IDS_ApplicationReload") Then

		Application.Contents.RemoveAll()
		doRedirect("edit_admin.asp?unload")

	Elseif Request.QueryString = "unload" Then

		Response.Write("<span class=""bFont"">" & getIDS("IDS_Done") & "</span><br />" & vbCrLf)

	End If
%>
	</div>
</div>

<div id="footerDiv" class="dvFooter">
<%
	Response.Write(getIconCancel("default.asp"))
%>
</div>
</form>

<%
	Call DisplayFooter(1)
%>