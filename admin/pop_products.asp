<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_admin.asp" -->
<%
	Call pageFunctions(90,5)

	Dim lngProductId	'as Long
	Dim lngVersionId	'as Long
	Dim strNewProduct	'as String
	Dim strNewVersion	'as String
	Dim arrProducts		'as Array
	Dim arrVersions		'as Array

	If not blnAdmin and pTickets < 5 and pBugs < 5 Then Call logError(2,1)

	lngProductId = valNum(Request.QueryString("prod"),3,0)
	lngVersionId = valNum(Request.Form("selVersion"),3,0)

	If valString(Request.Form("btnProduct"),-1,0,0) = Application("IDS_New") Then
		strNewProduct = valString(Request.Form("txtNewProduct"),255,0,0)
		If strNewProduct <> "" Then objConn.Execute(editProducts(1,strNewProduct,0))
		Call remAppVar("Products")
	Elseif valString(Request.Form("btnVersion"),-1,0,0) = Application("IDS_New") Then
		strNewVersion = valString(Request.Form("txtNewVersion"),50,0,0)
		If strNewVersion <> "" Then objConn.Execute(editProducts(2,lngProductId,strNewVersion))
		Call remAppVar("Products")
		Call remAppVar("Product Versions")
	Elseif valString(Request.Form("btnProduct"),-1,0,0) = Application("IDS_Delete") Then
		If lngProductId <> "" Then objConn.Execute(editProducts(3,lngProductId,0))
		Call remAppVar("Products")
	Elseif valString(Request.Form("btnVersion"),-1,0,0) = Application("IDS_Delete") Then
		If lngVersionId <> "" Then objConn.Execute(editProducts(4,0,lngVersionId))
		Call remAppVar("Products")
		Call remAppVar("Product Versions")
	End If

	Set objRS = objConn.Execute(editProducts(5,0,0))
	If not (objRS.BOF and objRS.EOF) Then arrProducts = objRS.GetRows()

	If lngProductId <> "" Then
		Set objRS = objConn.Execute(editProducts(6,lngProductId,0))
		If not (objRS.BOF and objRS.EOF) Then arrVersions = objRS.GetRows()
	End If

	strTitle = Application("IDS_EditProducts")
	Call DisplayHeader(3)
	Call showEditHeader(strTitle,"","","","")
%>
<div id="contentDiv" class="dvBorder" style="height:330px;">

<table border=0 cellspacing=10>
<form name="frmAdmin" action="pop_products.asp?prod=<% =lngProductId %>" method="post">
  <tr><td colspan=2><% =getLabel(Application("IDS_Product"),"txtNewProduct") %></td></tr>
  <tr>
    <td><% =getTextField("txtNewProduct","oText","",39,255,"") %></td>
    <td><% =getSubmit("btnProduct",Application("IDS_New"),70,"","") %></td>
  </tr>
  <tr>
    <td>
      <select name="selProduct" id="selProduct" size="5" class="oLong" style="width:250;" onClick="window.location.href='pop_products.asp?prod='+document.forms[0].selProduct.value;">
<%
	If not isArray(arrProducts) Then
		Response.Write("<option>" & Application("IDS_NoneSpecified") & "</option>" & vbCrLf)
	Else
		For i = 0 to UBound(arrProducts,2)

			Response.Write("          <option value=""" & arrProducts(0,i) & """" & _
					getDefault(0,lngProductId,arrProducts(0,i)) & ">" & _
					showString(arrProducts(1,i)) & "</option>" & vbCrLf)
		Next
	End If
%>
      </select>
    </td>
    <td valign="top"><% =getSubmit("btnProduct",Application("IDS_Delete"),70,"","") %></td>
  </tr>
<%	If lngProductId <> 0 Then %>
  <tr><td colspan=2><% =getLabel(Application("IDS_Version"),"txtNewVersion") %></td></tr>
  <tr>
    <td><% =getTextField("txtNewVersion","oText","",39,50,"") %></td>
    <td><% =getSubmit("btnVersion",Application("IDS_New"),70,"","") %></td>
  </tr>
  <tr>
    <td>
      <select name="selVersion" id="selVersion" size="5" class="oLong" style="width:250;">
<%
		If lngProductId <> "" Then
			If not isArray(arrVersions) Then
				Response.Write("        <option>" & Application("IDS_NoneSpecified") & "</option>" & vbCrLf)
			Else
				For i = 0 to UBound(arrVersions,2)

					Response.Write("        <option value=""" & arrVersions(0,i) & """>" & showString(arrVersions(1,i)) & "</option>" & vbCrLf)
				Next
			End If
		End If
%>
      </select>
    </td>
    <td valign="top"><% =getSubmit("btnVersion",Application("IDS_Delete"),70,"","") %></td>
  </tr>
<%	End If %>
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