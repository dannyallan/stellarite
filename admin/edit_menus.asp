<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_admin.asp" -->
<%
	Call pageFunctions(90,5)

	Dim arrGroups   'as Array
	Dim arrOptions  'as Array
	Dim strPerm     'as String
	Dim strGroup    'as String
	Dim lngGroup    'as Long
	Dim lngOption   'as Long
	Dim strClass    'as String
	Dim strMsg      'as String
	Dim strNewOption 'as String
	Dim blnNewOption 'as Boolean

	strTitle = getIDS("IDS_EditOptionValues")
	blnNewOption = False

	If blnAdmin Then
		strPerm = "0,1,2,3,4,5,6,7,8"
	Else
		For i = 1 to Len(Session("Permissions"))
			If Mid(Session("Permissions"),i,1) = 5 Then strPerm = strPerm & i & ","
		Next
		strPerm = Left(strPerm,Len(strPerm)-1)
	End If

	lngGroup = valNum(Request.QueryString("group"),3,0)
	lngOption = valNum(Request.Form("selOption"),3,0)
	strMsg = getIDS("IDS_MsgOptionValues")

	strNewOption = valString(Request.Form("txtNewOption"),100,0,0)

	If valString(Request.Form("btnSubmit"),-1,0,0) = getIDS("IDS_New") and strNewOption <> "" Then

		objConn.Execute(insertOptionValue(lngGroup,strNewOption))
		Call remAppVar(Request.Form("hdnGroup"))
		blnNewOption = True

	Elseif valString(Request.Form("btnSubmit"),-1,0,0) = getIDS("IDS_Delete") and lngOption <> 0 Then

		objConn.Execute(delOptionValue(lngOption))
		Call remAppVar(Request.Form("hdnGroup"))
	End If

	Set objRS = objConn.Execute(getOptionGroups(strPerm,0))
	If not (objRS.BOF and objRS.EOF) Then arrGroups = objRS.GetRows()

	If lngGroup <> 0 Then
		Set objRS = objConn.Execute(getOptionValues(lngGroup))
		If not (objRS.BOF and objRS.EOF) Then arrOptions = objRS.GetRows()
	End If

	Call DisplayHeader(1)
	Call showEditHeader(strTitle,"","","","")
%>

<form name="frmAdmin" method="post" action="edit_menus.asp?group=<% =lngGroup %>">
<div id="contentDiv" class="dvBorder" style="height:<% =intScreenH-170 %>px;"><br />

<table border=0 cellspacing=10>
<% =getHidden("hdnChange","") %>
  <tr>
	<td colspan=2>
	<select name="selGroup" id="selGroup" class="oLong" style="width:300px;" onChange="window.location.href='edit_menus.asp?group='+document.forms[0].selGroup.value;">
	<option></option>
<%
	If IsArray(arrGroups) Then
		For i = 0 to UBound(arrGroups,2)
			Response.Write(vbTab & "  <option value=""" & arrGroups(0,i) & """" & getDefault(0,lngGroup,arrGroups(0,i)) & ">" & showString(getIDS(arrGroups(1,i))) & "</option>" & vbCrLf)
			If CStr(arrGroups(0,i)) = CStr(lngGroup) Then
				strGroup = arrGroups(1,i)
				strMsg = arrGroups(2,i)
			End If
		Next
	End If
%>
	</select>

<div style="width:310;height:55px;overflow:auto;margin:0px;padding:10px;">
<p class="dFont"><% =strMsg %></p>
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
	<% =getSubmit("btnSubmit",getIDS("IDS_New"),70,"N","") %>
	</td>
  </tr>
  <tr>
	<td>
	<select name="selOption" id="selOption" size="5" class="oLong" style="width:300;">
<%
		If not isArray(arrOptions) Then
			Response.Write(vbTab & "  <option>" & getIDS("IDS_NoneSpecified") & "</option>" & vbCrLf)
		Else
			For i = 0 to UBound(arrOptions,2)
				If blnNewOption and strNewOption = arrOptions(1,i) Then Call setAppVar("ao_Option" & arrOptions(0,i),strNewOption)
				Response.Write(vbTab & "  <option value=""" & arrOptions(0,i) & """>" & showString(arrOptions(1,i)) & "</option>" & vbCrLf)
			Next
		End If
%>
	</select>
	</td>
	<td valign="top">
	<% =getSubmit("btnSubmit",getIDS("IDS_Delete"),70,"D","") %>
	</td>
  </tr>
<% End If %>

</table>

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