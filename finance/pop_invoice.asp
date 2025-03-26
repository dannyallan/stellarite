<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_finance.asp" -->
<!--#include file="..\_inc\currency.asp" -->
<%
	Call pageFunctions(7,2)

	Dim strInvoiceNum	'as String
	Dim strClient		'as String
	Dim lngDivId		'as Long
	Dim strContact		'as String
	Dim lngContactId	'as Long
	Dim strOwner		'as String
	Dim lngOwnerId		'as Long
	Dim strPO			'as String
	Dim blnReceived		'as Integer
	Dim strType			'as String
	Dim intType			'as Integer
	Dim strPhase		'as String
	Dim intPhase		'as String
	Dim strCurrency		'as String
	Dim decValue		'as Decimal
	Dim decTax			'as Decimal
	Dim datSendDate		'as Date
	Dim datDueDate		'as Date
	Dim datPaidDate		'as Date
	Dim blnClosed		'as String
	Dim strCreatedBy	'as String
	Dim datCreatedDate	'as Date
	Dim strModBy		'as String
	Dim datModDate		'as Date

	strTitle = Application("IDS_Edit") & " " & Application("IDS_Invoice")

	If strDoAction <> "" Then

		lngContactId = valNum(Request.Form("hdnContact"),3,1)
		lngDivId = valNum(Request.Form("hdnDivision"),3,1)
		strOwner = valString(Request.Form("txtOwner"),100,1,0)
		lngOwnerId = getUserId(7,strOwner)
		strPO = valString(Request.Form("txtPurchaseOrder"),25,0,0)
		blnReceived = valNum(Request.Form("selReceived"),1,1)
		intType = valNum(Request.Form("selType"),2,-1)
		intPhase = valNum(Request.Form("selPhase"),2,-1)
		strCurrency = valString(Request.Form("selCurrency"),3,1,0)
		decValue = valNum(Request.Form("txtValue"),5,1)
		decTax = valNum(Request.Form("txtTax"),5,1)
		datSendDate = valDate(Request.Form("txtInvoiceSent"),0)
		datDueDate = valDate(Request.Form("txtInvoiceDue"),0)
		datPaidDate = valDate(Request.Form("txtInvoicePaid"),0)
		blnClosed = valNum(Request.Form("chkClosed"),0,0)

		If strDoAction = "del" and intPerm >= 4 Then

			Call delInvoice(lngUserId,lngRecordId)

		ElseIf strDoAction = "edit" and intPerm >= 3 Then

			Call updateInvoice(lngUserId,lngRecordId,lngContactId,lngDivId,lngOwnerId, strPO, blnClosed, _
					blnReceived,intType,intPhase,strCurrency,decValue,decTax,datSendDate,datDueDate,datPaidDate)

		ElseIf strDoAction = "new" Then

			lngRecordId = insertInvoice(lngUserId,lngRecordId,lngContactId,lngDivId,lngOwnerId, strPO, blnClosed, _
					blnReceived,intType,intPhase,strCurrency,decValue,decTax,datSendDate,datDueDate,datPaidDate)
		End If
		Call closeWindow(strOpenerURL)
	Else

		If blnRS and intPerm >= 3 Then
			Set objRS = objConn.Execute(getInvoice(0,lngRecordId))

			If not (objRS.BOF and objRS.EOF) then
				strContact = showString(objRS.fields("Contact").value)
				lngContactId = objRS.fields("ContactId").value
				lngDivId = objRS.fields("DivId").value
				strOwner = showString(objRS.fields("Owner").value)
				strPO = showString(objRS.fields("I_PurchaseOrder").value)
				blnReceived = objRS.fields("I_Received").value
				intType = objRS.fields("I_Type").value
				intPhase = objRS.fields("I_Phase").value
				strCurrency = showString(objRS.fields("I_Currency").value)
				decValue = objRS.fields("I_Value").value
				decTax = objRS.fields("I_Tax").value
				datSendDate = showDate(0,objRS.fields("I_SendDate").value)
				datDueDate = showDate(0,objRS.fields("I_DueDate").value)
				datPaidDate = showDate(0,objRS.fields("I_PaidDate").value)
				strCreatedBy = showString(objRS.fields("CreatedBy").value)
				datCreatedDate = objRS.fields("I_CreatedDate").value
				blnClosed = objRS.fields("I_Closed").value
				strModBy = showString(objRS.fields("ModBy").value)
				datModDate = objRS.fields("I_ModDate").value
			End If
		Elseif blnRS Then
			Call doRedirect("pop_invoice.asp?m=" & bytMod & "&mid=" & lngModId)
		Else
			If bytMod = 1 Then
				lngContactId = lngModId
				lngDivId = getValue("DivId","CRM_Contacts","ContactId = "&lngContactId,0)
				strContact = getValue(doConCat(doConCat("K_FirstName","' '"),"K_LastName"),"CRM_Contacts","ContactId = "&lngModId,"")
			Elseif bytMod = 2 Then
				lngDivId = lngModId
				lngContactId = getValue("ContactId","CRM_Contacts","DivId = "&lngDivId,"")
				If lngContactId = "" Then strContact = "" Else strContact = getValue(doConCat(doConCat("K_FirstName","' '"),"K_LastName"),"CRM_Contacts","ContactId = " & lngContactId,"")
			End If
			If mInvoices Then strOwner = strFullName
			strCurrency = Application("av_Currency")
			strCreatedBy = strFullName
			datCreatedDate = Now
			strModBy = strFullName
			datModDate = Now
		End If
		strType = getOptionDropDown(260,False,"selType","Invoice Type",intType)
		strPhase = getOptionDropDown(260,False,"selPhase","Invoice Phase",intPhase)
	End If

	strIncHead = getCalendarScripts()

	Call DisplayHeader(3)
	Call showEditHeader(strTitle,strCreatedBy,datCreatedDate,strModBy,datModDate)
%>
<div id="contentDiv" class="dvBorder" style="height:380px;"><br>

<table border=0 width="100%">
<form name="frmInvoice" method="post" action="pop_invoice.asp?id=<% =lngRecordId %>&m=<% =bytMod %>&mid=<% =lngModId %>">
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
<% =getHidden("hdnWinOpen","") %>
<% =getHidden("hdnContact",lngContactId) %>
<% =getHidden("hdnDivision",lngDivId) %>
  <tr>
    <td><% =getLabel(Application("IDS_Contact"),"txtContact") %></td>
    <td colspan=3><% =getTextField("txtContact","mText",strContact,40,100,"readonly=""readonly""") %>
    <% If pContacts >= 1 Then %>
    <a href="<% =newWindow("S","?m=1&rVal=K") %>"><img src="../images/import.gif" alt="<% =getImport("IDS_Contact") %>" border=0 height=16 width=16></a>
    <% End If %>
    <% If pContacts >= 2 Then %>
    <a href="<% =newWindow(1,"?m=7") %>"><img src="../images/new2.gif" alt="<% =getImport("IDS_ContactNew") %>" border=0 height=16 width=16></a></td>
    <% End If %>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_Owner"),"txtOwner") %></td>
    <td><% =getTextField("txtOwner","mText",strOwner,40,100,"") %>
    <a href="<% =newWindow("S","?m=0&rVal=txtOwner") %>"><img src="../images/import.gif" alt="<% =getImport("IDS_Owner") %>" border=0 height=16 width=16></a></td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_PurchaseOrder"),"txtPurchaseOrder") %></td>
    <td><% =getTextField("txtPurchaseOrder","oText",strPO,40,25,"") %></td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_Invoice"),"selReceived") %></td>
    <td>
	  <select name="selReceived" id="selReceived" class="oBool" onChange="doChange();" style="width:260px">
		<option value="0"<% =getDefault(0,blnReceived,0) %>><% =Application("IDS_InvoiceReceived") %></option>
		<option value="1"<% =getDefault(0,blnReceived,1) %>><% =Application("IDS_InvoiceSent") %></option>
	  </select>
    </td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_Type"),"selType") %></td>
    <td><% =strType %></td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_Phase"),"selPhase") %></td>
    <td><% =strPhase %></td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_Currency"),"selCurrency") %></td>
    <td><% =getCurrency(260,"selCurrency",strCurrency) %></td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_Value"),"txtValue") %></td>
    <td><% =getTextField("txtValue","mCurrency",decValue,12,25,"onBlur=""calcTax();""") %></td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_Tax"),"txtTax") %></td>
    <% If Application("av_TaxChange") = "1" Then %>
    <td><% =getTextField("txtTax","dCurrency",decTax,12,25,"readonly=""readonly""") %> <% =getTextField("txtTaxRate","mCurrency",Application("av_TaxRate"),4,4,"onBlur=""calcTax();""") %><span class="dfont">%</span></td>
	<% Else %>
    <td><% =getTextField("txtTax","dCurrency",decTax,12,25,"readonly=""readonly""") %> <% =getTextField("txtTaxRate","dCurrency",Application("av_TaxRate"),4,4,"readonly=""readonly""") %><span class="dfont">%</span></td>
	<% End If %>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_InvoiceDate"),"txtInvoiceSent") %></td>
    <td><% =getTextField("txtInvoiceSent","oDate",datSendDate,12,12,"") %>
    <a href="Javascript:showCalendar('txtInvoiceSent');"><img src="../images/cal.gif" alt="<% =getImport("IDS_InvoiceDate") %>" border=0 height=16 width=16></a></td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_InvoiceDue"),"txtInvoiceDue") %></td>
    <td><% =getTextField("txtInvoiceDue","oDate",datDueDate,12,12,"") %>
    <a href="Javascript:showCalendar('txtInvoiceDue');"><img src="../images/cal.gif" alt="<% =getImport("IDS_InvoiceDue") %>" border=0 height=16 width=16></a></td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_InvoicePaid"),"txtInvoicePaid") %></td>
    <td><% =getTextField("txtInvoicePaid","oDate",datPaidDate,12,12,"") %>
    <a href="Javascript:showCalendar('txtInvoicePaid');"><img src="../images/cal.gif" alt="<% =getImport("IDS_InvoicePaid") %>" border=0 height=16 width=16></a></td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_Closed"),"chkClosed") %></td>
    <td><% =getCheckbox("chkClosed",blnClosed,"onClick=""doClassChange();""") %></td>
  </tr>
</form>
</table>

</div>

<div id="footerDiv" class="dvFooter">
<%
	If blnRS Then
		Response.Write(getIconNew("pop_invoice.asp"))
		If intPerm >= 4 Then Response.Write(getIconDelete())
	End If
	Response.Write(getIconSave(strAction))
	Response.Write(getIconCancel())
%>
</div>

<script>
function calcTax() {
	var dTax = (document.forms[0].txtValue.value * document.forms[0].txtTaxRate.value)/100;
	document.forms[0].txtTax.value = dTax.toFixed(2);
}
</script>

<%
	Call DisplayFooter(3)
%>

