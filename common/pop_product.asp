<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_products.asp" -->
<%
	Call pageFunctions(3,2)

	Dim strInvoiceDD	'as String
	Dim lngInvoiceId	'as String
	Dim lngDivId		'as Long
	Dim strProductDD	'as String
	Dim intProductId	'as Integer
	Dim strSerial		'as String
	Dim strPIN			'as String
	Dim datExpiry		'as Date
	Dim strCreatedBy	'as String
	Dim datCreatedDate	'as Date
	Dim strModBy		'as String
	Dim datModDate		'as Date

	strTitle = Application("IDS_Edit") & " " & Application("IDS_Products")

	If strDoAction <> "" then

		lngInvoiceId = valNum(Request.Form("selInvoice"),3,-1)
		intProductId = valNum(Request.Form("selProductId"),3,-1)
		strSerial = valString(Request.Form("txtSerial"),100,0,0)
		strPIN = valString(Request.Form("txtPIN"),100,0,0)
		datExpiry = valDate(Request.Form("txtExpiry"),0)

		If strDoAction = "del" and intPerm >= 4 Then

			Call delProduct(lngUserId,lngRecordId)

		Elseif strDoAction = "edit" and intPerm >= 3 then

			Call updateProduct(lngUserId,lngRecordId,bytMod,lngModId,lngInvoiceId,intProductId,strSerial,strPIN,datExpiry)

		ElseIf strDoAction = "new" Then

			lngRecordId = insertProduct(lngUserId,lngRecordId,bytMod,lngModId,lngInvoiceId,intProductId,strSerial,strPIN,datExpiry)
		End If
		Call closeWindow(strOpenerURL)
	Else

		If blnRS and intPerm >= 3 Then

			Set objRS = objConn.Execute(getProducts(0,lngRecordId,0,0))

			If not (objRS.BOF and objRS.EOF) then
				lngRecordId = objRS.fields("SerialzId").value
				lngDivId = objRS.fields("DivId").value
				lngInvoiceId = objRS.fields("InvoiceId").value
				intProductId = objRS.fields("ProductId").value
				strSerial = showString(objRS.fields("Z_Serial").value)
				strPIN = showString(objRS.fields("Z_PIN").value)
				datExpiry = showDate(0,objRS.fields("Z_Expiry").value)
				strCreatedBy = showString(objRS.fields("CreatedBy").value)
				datCreatedDate = objRS.fields("Z_CreatedDate").value
				strModBy = showString(objRS.fields("ModBy").value)
				datModDate = objRS.fields("Z_ModDate").value
			End If

		Elseif blnRS Then
			Call doRedirect("pop_product.asp?m=" & bytMod & "&mid=" & lngModId)
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
		strProductDD = getProductDropDown(150,False,"selProductId",intProductId)
	End If

	strIncHead = getCalendarScripts()

	Call DisplayHeader(3)
	Call showEditHeader(strTitle,strCreatedBy,datCreatedDate,strModBy,datModDate)
%>
<div id="contentDiv" class="dvBorder" style="height:230px;"><br>

<table border=0 cellspacing=5 width="100%">
<form name="frmProduct" method="post" action="pop_product.asp?id=<% =lngRecordId %>&m=<% =bytMod %>&mid=<% =lngModId %>">
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
<% =getHidden("hdnWinOpen","") %>
  <tr>
    <td><% =getLabel(Application("IDS_InvoiceId"),"selInvoice") %></td>
    <td><% =strInvoiceDD %></td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_Product"),"selProductId") %></td>
    <td><% =strProductDD %></td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_Serial"),"txtSerial") %></td>
    <td><% =getTextField("txtSerial","oText",strSerial,20,100,"") %></td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_PIN"),"txtPIN") %></td>
    <td><% =getTextField("txtPIN","oText",strPIN,20,100,"") %></td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_Expiry"),"txtExpiry") %></td>
    <td><% =getTextField("txtExpiry","oDate",datExpiry,12,12,"") %>
    <a href="Javascript:showCalendar('txtExpiry');"><img src="../images/cal.gif" alt="<% =getImport("IDS_Expiry") %>" border=0 height=16 width=16></a></td>
  </tr>
</form>
</table>

</div>

<div id="footerDiv" class="dvFooter">
<%
	If blnRS Then
		Response.Write(getIconNew("pop_product.asp?m=" & bytMod & "&mid=" & lngModId))
		If intPerm >= 4 Then Response.Write(getIconDelete())
	End If
	Response.Write(getIconSave(strAction))
	Response.Write(getIconCancel())
%>
</div>

<%
	Call DisplayFooter(3)
%>

