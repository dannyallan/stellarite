<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_products.asp" -->
<%
	Call pageFunctions(0,2)

	Dim strInvoiceDD    'as String
	Dim lngInvoiceId    'as String
	Dim lngDivId        'as Long
	Dim strProduct      'as String
	Dim lngProductId    'as Integer
	Dim strSerial       'as String
	Dim strPIN          'as String
	Dim datExpiry       'as Date
	Dim strCreatedBy    'as String
	Dim datCreatedDate  'as Date
	Dim strModBy        'as String
	Dim datModDate      'as Date

	strTitle = getIDS("IDS_Edit") & " " & getIDS("IDS_Products")

	If strDoAction <> "" then

		lngInvoiceId = valNum(Request.Form("selInvoice"),3,-1)
		lngProductId = valNum(Request.Form("selProduct"),3,-1)
		strSerial = valString(Request.Form("txtSerial"),100,0,0)
		strPIN = valString(Request.Form("txtPIN"),100,0,0)
		datExpiry = valDate(Request.Form("txtExpiry"),0)

		If strDoAction = "del" and intPerm >= 4 Then

			Call delProduct(lngUserId,lngRecordId)

		Elseif strDoAction = "edit" and intPerm >= 3 then

			Call updateProduct(lngUserId,lngRecordId,bytMod,lngModId,lngInvoiceId,lngProductId,strSerial,strPIN,datExpiry)

		ElseIf strDoAction = "new" Then

			lngRecordId = insertProduct(lngUserId,lngRecordId,bytMod,lngModId,lngInvoiceId,lngProductId,strSerial,strPIN,datExpiry)
		End If

		Session("LastPage") = "i_products.asp?id=" & lngRecordId & "&m=" & bytMod & "&mid=" & lngModId
		Call closeEdit()
	Else

		If blnRS and intPerm >= 3 Then

			Set objRS = objConn.Execute(getProducts(0,lngRecordId,0,0))

			If not (objRS.BOF and objRS.EOF) then
				lngRecordId = objRS.fields("SerialzId").value
				lngDivId = objRS.fields("DivId").value
				lngInvoiceId = objRS.fields("InvoiceId").value
				lngProductId = objRS.fields("Z_ProductId").value
				strSerial = objRS.fields("Z_Serial").value
				strPIN = objRS.fields("Z_PIN").value
				datExpiry = objRS.fields("Z_Expiry").value
				strCreatedBy = objRS.fields("CreatedBy").value
				datCreatedDate = objRS.fields("Z_CreatedDate").value
				strModBy = objRS.fields("ModBy").value
				datModDate = objRS.fields("Z_ModDate").value
			End If

		Elseif blnRS Then
			Call logError(2,1)
		Else
			Select Case bytMod
				Case 1
					lngDivId = getValue("DivId","CRM_Contacts","ContactId="&lngModId,0)
				Case 2
					lngDivId = lngModId
				Case 7
					lngDivId = getValue("DivId","CRM_Invoices","InvoiceId="&lngModId,0)
					lngInvoiceId = lngModId
			End Select
			strCreatedBy = strFullName
			datCreatedDate = Now
			strModBy = strFullName
			datModDate = Now
		End If
		strInvoiceDD = getInvoiceDropDown(150,True,"selInvoice",lngDivId,lngInvoiceId)
		strProduct = getOptionDropDown(150,False,"selProduct","IDS_Product",lngProductId)
	End If

	strIncHead = getCalendarScripts()

	Call DisplayHeader(3)
	Call showEditHeader(strTitle,strCreatedBy,datCreatedDate,strModBy,datModDate)
%>
<form name="frmProduct" method="post" action="edit_product.asp?id=<% =lngRecordId %>&m=<% =bytMod %>&mid=<% =lngModId %>">
<div id="contentDiv" class="dvBorder" style="height:230px;"><br />

<table border=0 cellspacing=5 width="100%">
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
  <tr>
	<td><% =getLabel(getIDS("IDS_InvoiceId"),"selInvoice") %></td>
	<td><% =strInvoiceDD %></td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_Product"),"selProductId") %></td>
	<td><% =strProduct %></td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_Serial"),"txtSerial") %></td>
	<td><% =getTextField("txtSerial","oText",strSerial,20,100,"") %></td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_PIN"),"txtPIN") %></td>
	<td><% =getTextField("txtPIN","oText",strPIN,20,100,"") %></td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_Expiry"),"txtExpiry") %></td>
	<td><% =getDateField("txtExpiry","oDate",datExpiry,getIDS("IDS_Expiry")) %></td>
  </tr>
</table>

</div>

<div id="footerDiv" class="dvFooter">
<%
	If blnRS Then
		Response.Write(getIconNew(getEditURL("Z","?m="&bytMod&"&mid="&lngModId)))
		If intPerm >= 4 Then Response.Write(getIconDelete())
	End If
	Response.Write(getIconSave(strAction))
	Response.Write(getIconCancel("back"))
%>
</div>
</form>

<%
	Call DisplayFooter(3)
%>

