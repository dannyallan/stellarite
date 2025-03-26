<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_admin.asp" -->
<%
	Call pageFunctions(0,5)

	Dim arrGroups	'as Array
	Dim arrOptions	'as Array
	Dim strPerm	'as String
	Dim strGroup	'as String
	Dim lngGroup	'as Long
	Dim lngOption	'as Long
	Dim strClass	'as String
	Dim strMsg	'as String
	Dim strNewOption 'as String

	If not (blnAdmin or Instr(Session("Permissions"),"5") > 0) Then
		Call logError(2,1)
	Elseif blnAdmin Then
		strPerm = "(0,1,2,3,4,5,6,7,8)"
	Else
		strPerm = "("
		For i = 1 to Len(Session("Permissions"))
			If Mid(Session("Permissions"),i,1) = 5 Then strPerm = strPerm & i & ","
		Next
		strPerm = Left(strPerm,Len(strPerm)-1) & ")"
	End If

	lngGroup = valNum(Request.QueryString("group"),3,0)
	lngOption = valNum(Request.Form("selOption"),3,0)
	strMsg = Application("IDS_MsgOptionValues")

	strNewOption = valString(Request.Form("txtNewOption"),100,0,0)

	If valString(Request.Form("btnSubmit"),-1,0,0) = Application("IDS_New") and strNewOption <> "" Then

		objConn.Execute(insertOptionValue(lngGroup,strNewOption))
		Call remAppVar(Request.Form("hdnGroup"))

	Elseif valString(Request.Form("btnSubmit"),-1,0,0) = Application("IDS_Delete") and lngOption <> 0 Then

		objConn.Execute(delOptionValue(lngOption))
		Call remAppVar(Request.Form("hdnGroup"))
	End If

	Set objRS = objConn.Execute(getOptionGroups(strPerm))
	If not (objRS.BOF and objRS.EOF) Then arrGroups = objRS.GetRows()

	If lngGroup <> 0 Then
		Set objRS = objConn.Execute(getOptionValues(lngGroup))
		If not (objRS.BOF and objRS.EOF) Then arrOptions = objRS.GetRows()
	End If

	strTitle = Application("IDS_EditOptionValues")
	Call DisplayHeader(3)
	Call showEditHeader(strTitle,"","","","")
%>

<div id="contentDiv" class="dvBorder" style="height:330px;"><br>

<table border=0 cellspacing=10>
<form name="frmAdmin" action="pop_menus.asp?group=<% =lngGroup %>" method="post">
  <tr>
    <td class="dfont" colspan=2>
	<select name="selGroup" id="selGroup" class="oLong" style="width:300px;" onChange="window.location.href='pop_menus.asp?group='+document.forms[0].selGroup.value;">
	<option></option>
<%
	If IsArray(arrGroups) Then
		For i = 0 to UBound(arrGroups,2)
			Response.Write(vbTab & "  <option value=""" & arrGroups(0,i) & """" & getDefault(0,lngGroup,arrGroups(0,i)) & ">" & showString(arrGroups(1,i)) & "</option>" & vbCrLf)
			If CStr(arrGroups(0,i)) = CStr(lngGroup) Then
				strGroup = arrGroups(1,i)
				strMsg = arrGroups(2,i)
			End If
		Next
	End If
%>
	</select>

<div style="width:310;height:55px;overflow:auto;margin:0px;padding:10px;">
<% =strMsg %>
</div>

    </td>
  </tr>

<% If lngGroup <> 0 Then %>

  <tr>
    <td>
        <% =getTextField("txtNewOption","oText","",47,100,"") %>
	<% =getHidden("hdnGroup",strGroup) %>
    </td>
    <td>
	<% =getSubmit("btnSubmit",Application("IDS_New"),70,"N","") %>
    </td>
  </tr>
  <tr>
    <td>
	<select name="selOption" id="selOption" size="5" class="oLong" style="width:300;">
<%
		If not isArray(arrOptions) Then
			Response.Write(vbTab & "  <option>" & Application("IDS_NoneSpecified") & "</option>" & vbCrLf)
		Else
			For i = 0 to UBound(arrOptions,2)

				Response.Write(vbTab & "  <option value=""" & arrOptions(0,i) & """>" & showString(arrOptions(1,i)) & "</option>" & vbCrLf)
			Next
		End If
%>
	</select>
    </td>
    <td valign="top">
	<% =getSubmit("btnSubmit",Application("IDS_Delete"),70,"D","") %>
    </td>
  </tr>
<% End If %>

</form>
</table>

</div>

<div id="footerDiv" class="dvFooter">
<%
	Response.Write(getIconCancel())
%>
</div>

<%
	Call DisplayFooter(3)
%>