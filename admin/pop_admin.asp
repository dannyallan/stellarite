<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_admin.asp" -->
<%
	Server.ScriptTimeout = 6000

	Call pageFunctions(90,5)

	Dim intCount		'as Integer

	If not blnAdmin Then Call logError(2,1)

	strTitle = Application("IDS_CRMAdministration")
	Call DisplayHeader(3)
	Call showEditHeader(strTitle,"","","","")

%>
	<div id="contentDiv" class="dvBorder" style="height:330px;"><br>

	<p class="dfont"><% =Application("IDS_MsgWaitForEnd") %></p>

	<div id="hideDiv">
		<form name="frmAdmin" method="post" action="pop_admin.asp" onSubmit="hideDiv.style.visibility='hidden';">
		<% =getSubmit("btnReload",Application("IDS_ApplicationReload"),200,"1","") %>
		<br><br>
		<% =getSubmit("btnCount",Application("IDS_UpdateCounts"),200,"2","") %>
		</form>

<%
	If Request.Form("btnCount") = Application("IDS_UpdateCounts") Then
		For i = 1 to bytModCount
			Response.Write("<span class=""dfont"">" & Application("IDS_Updating") & " " & Application("IDS_Module" & i) & " ... </span>" & vbCrLf)
			Response.Flush

			Set objRS = objConn.Execute(getModuleList(i))
			If not (objRS.BOF and objRS.EOF) Then

				arrRS = objRS.GetRows()

				For intCount = 0 to UBound(arrRS,2)
					objConn.Execute(updateModuleCount(i,arrRS(0,intCount)))
				Next
			End If

			Response.Write("<span class=""dlabel"">" & Application("IDS_Done") & "</span><br>" & vbCrLf)
		Next
		Response.Write("<p class=""dlabel"">" & Application("IDS_Complete") & "</p>" & vbCrLf)

	Elseif Request.Form("btnReload") = Application("IDS_ApplicationReload") Then

		Application.Contents.RemoveAll()
		doRedirect("pop_admin.asp?unload")

	Elseif Request.QueryString = "unload" Then

		Response.Write("<span class=""dlabel"">" & Application("IDS_Done") & "</span><br>" & vbCrLf)

	End If
%>
	</div>
</div>

<div id="footerDiv" class="dvFooter">
<%
	Response.Write(getIconCancel())
%>
</div>

<%
	Call DisplayFooter(3)
%>