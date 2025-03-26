<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_admin.asp" -->
<%
	Call pageFunctions(90,5)

	Dim arrProducts   'as Array
	Dim lngGroupId    'as Long
	Dim strProduct    'as String
	Dim lngOptionId   'as Long
	Dim strClass      'as String
	Dim strNewOption  'as String

	strTitle = getIDS("IDS_EditProducts")

	Set objRS = objConn.Execute(getOptionGroups("",1))
	If not (objRS.BOF and objRS.EOF) Then
		lngGroupId = objRS.fields("OptGroupId").value
		strProduct = objRS.fields("G_Name").value
	End If

	lngOptionId = valNum(Request.Form("selOption"),3,0)
	strNewOption = valString(Request.Form("txtNewProduct"),100,0,0)

	If valString(Request.Form("btnSubmit"),-1,0,0) = getIDS("IDS_New") and strNewOption <> "" Then

		objConn.Execute(insertOptionValue(lngGroupId,strNewOption))
		Call remAppVar(Request.Form("hdnGroup"))

	Elseif valString(Request.Form("btnSubmit"),-1,0,0) = getIDS("IDS_Delete") and lngOptionId <> 0 Then

		objConn.Execute(delOptionValue(lngOptionId))
		Call remAppVar(Request.Form("hdnGroup"))
	End If

	Set objRS = objConn.Execute(getOptionValues(lngGroupId))
	If not (objRS.BOF and objRS.EOF) Then arrProducts = objRS.GetRows()

	Call DisplayHeader(1)
	Call showEditHeader(strTitle,"","","","")
%>

<form name="frmAdmin" method="post" action="edit_products.asp">
<div id="contentDiv" class="dvBorder" style="height:<% =intScreenH-170 %>px;"><br />

<table border=0 cellspacing=10>
<% =getHidden("hdnChange","") %>
  <tr>
    <td colspan="2" class="dFont"><% =getIDS("IDS_ProductNew") %></td>
  </tr>
  <tr>
	<td>
	<% =getTextField("txtNewProduct","oText","",47,100,"") %>
	</td>
	<td>
	<% =getSubmit("btnSubmit",getIDS("IDS_New"),70,"N","") %>
	</td>
  </tr>
  <tr>
	<td>
	<select name="selOption" id="selOption" size="5" class="oLong" style="width:300;">
<%
		If not isArray(arrProducts) Then
			Response.Write(vbTab & "  <option>" & getIDS("IDS_NoneSpecified") & "</option>" & vbCrLf)
		Else
			For i = 0 to UBound(arrProducts,2)
				Response.Write(vbTab & "  <option value=""" & arrProducts(0,i) & """>" & showString(arrProducts(1,i)) & "</option>" & vbCrLf)
			Next
		End If
%>
	</select>
	</td>
	<td valign="top">
	<% =getSubmit("btnSubmit",getIDS("IDS_Delete"),70,"D","") %>
	</td>
  </tr>
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