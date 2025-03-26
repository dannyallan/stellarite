<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_finance.asp" -->
<!--#include file="..\_inc\currency.asp" -->
<%
	Call pageFunctions(7,2)

	Dim strInvoiceNum   'as String
	Dim strClient       'as String
	Dim lngDivId        'as Long
	Dim strContact      'as String
	Dim lngContactId    'as Long
	Dim strOwner        'as String
	Dim lngOwnerId      'as Long
	Dim strPO           'as String
	Dim blnReceived     'as Integer
	Dim strType         'as String
	Dim intType         'as Integer
	Dim strPhase        'as String
	Dim intPhase        'as String
	Dim strCurrency     'as String
	Dim decValue        'as Decimal
	Dim decTax          'as Decimal
	Dim strInfo         'as String
	Dim datSendDate     'as Date
	Dim datDueDate      'as Date
	Dim datPaidDate     'as Date
	Dim blnClosed       'as String
	Dim strCreatedBy    'as String
	Dim datCreatedDate  'as Date
	Dim strModBy        'as String
	Dim datModDate      'as Date
	Dim strExtraFields  'as String

	strTitle = getIDS("IDS_Edit") & " " & getIDS("IDS_Invoice")

	If strDoAction <> "" Then

		Select Case strDoAction
			Case "del"
				If intPerm >= 4 Then Call delInvoice(lngUserId,lngRecordId)

			Case "new", "edit"

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
				strInfo = valString(Request.Form("txtInfo"),-1,0,5)
				datSendDate = valDate(Request.Form("txtInvoiceSent"),0)
				datDueDate = valDate(Request.Form("txtInvoiceDue"),0)
				datPaidDate = valDate(Request.Form("txtInvoicePaid"),0)
				blnClosed = valNum(Request.Form("chkClosed"),0,0)

				If strDoAction = "edit" and intPerm >= 3 Then

					Call updateInvoice(lngUserId,lngRecordId,lngContactId,lngDivId,lngOwnerId, strPO, blnClosed, _
							blnReceived,intType,intPhase,strCurrency,decValue,decTax,strInfo,datSendDate,datDueDate,datPaidDate)

				ElseIf strDoAction = "new" Then

					lngRecordId = insertInvoice(lngUserId,lngRecordId,lngContactId,lngDivId,lngOwnerId, strPO, blnClosed, _
							blnReceived,intType,intPhase,strCurrency,decValue,decTax,strInfo,datSendDate,datDueDate,datPaidDate)
				End If

				Call saveCustomFields(7,lngRecordId)
		End Select

		If bytMenu = 0 Then Session("LastPage") = strCRMURL & "common/i_invoices.asp?m=" & bytMod & "&mid=" & lngModId
		Call closeEdit()
	Else

		If blnRS and intPerm >= 3 Then
			Set objRS = objConn.Execute(getInvoice(0,lngRecordId))

			If not (objRS.BOF and objRS.EOF) then
				strContact = objRS.fields("Contact").value
				lngContactId = objRS.fields("ContactId").value
				lngDivId = objRS.fields("DivId").value
				strOwner = objRS.fields("Owner").value
				strPO = objRS.fields("I_PurchaseOrder").value
				blnReceived = objRS.fields("I_Received").value
				intType = objRS.fields("I_Type").value
				intPhase = objRS.fields("I_Phase").value
				strCurrency = objRS.fields("I_Currency").value
				decValue = objRS.fields("I_Value").value
				decTax = objRS.fields("I_Tax").value
				strInfo = objRS.fields("I_PayInfo").value
				datSendDate = objRS.fields("I_SendDate").value
				datDueDate = objRS.fields("I_DueDate").value
				datPaidDate = objRS.fields("I_PaidDate").value
				strCreatedBy = objRS.fields("CreatedBy").value
				datCreatedDate = objRS.fields("I_CreatedDate").value
				blnClosed = objRS.fields("I_Closed").value
				strModBy = objRS.fields("ModBy").value
				datModDate = objRS.fields("I_ModDate").value
			End If
		Elseif blnRS Then
			Call logError(2,1)
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
		strExtraFields = editCustomFields(7)
		strType = getOptionDropDown(260,False,"selType","IDS_InvoiceType",intType)
		strPhase = getOptionDropDown(260,False,"selPhase","IDS_InvoicePhase",intPhase)
	End If

	strIncHead = getCalendarScripts()

	Call DisplayHeader(0)
	Call showEditHeader(strTitle,strCreatedBy,datCreatedDate,strModBy,datModDate)
%>
<form name="frmInvoice" method="post" action="edit_invoice.asp?id=<% =lngRecordId %>&m=<% =bytMod %>&mid=<% =lngModId %>&menu=<% =bytMenu %>">
<div id="contentDiv" class="dvBorder" style="height:<% =intScreenH-170 %>px;"><br />

<table border=0>
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
<% =getHidden("hdnContact",lngContactId) %>
<% =getHidden("hdnDivision",lngDivId) %>
  <tr>
	<td width=170><% =getLabel(getIDS("IDS_Contact"),"txtContact") %></td>
	<td colspan=3><% =getTextField("txtContact","mText",strContact,40,100,"readonly=""readonly""") %>
<%
	If pContacts >= 1 Then
		Response.Write(getIconImport(1,getSearchURL("?m=1&rVal=K"),getIDS("IDS_Contact")))
	End If
	If pContacts >= 2 Then
		Response.Write(getIconImport(3,getEditURL(1,"?m="&bytMod&"&mid="&lngModId),getIDS("IDS_Contact")))
	End If
%>
	</td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_Owner"),"txtOwner") %></td>
	<td><% =getTextField("txtOwner","mText",strOwner,40,100,"") %>
	<% =getIconImport(1,getSearchURL("?m=0&rVal=txtOwner"),getIDS("IDS_Owner")) %>
	</td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_PurchaseOrder"),"txtPurchaseOrder") %></td>
	<td><% =getTextField("txtPurchaseOrder","oText",strPO,40,25,"") %></td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_Invoice"),"selReceived") %></td>
	<td>
	  <select name="selReceived" id="selReceived" class="oBool" onChange="doChange();" style="width:260px">
		<option value="0"<% =getDefault(0,blnReceived,0) %>><% =getIDS("IDS_InvoiceReceived") %></option>
		<option value="1"<% =getDefault(0,blnReceived,1) %>><% =getIDS("IDS_InvoiceSent") %></option>
	  </select>
	</td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_Type"),"selType") %></td>
	<td><% =strType %></td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_Phase"),"selPhase") %></td>
	<td><% =strPhase %></td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_Currency"),"selCurrency") %></td>
	<td><% =getCurrency(260,"selCurrency",strCurrency) %></td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_InvoiceDetails"),"txtInfo") %></td>
	<td><% =getTextArea("txtInfo","oText",strInfo,"260",3,"") %></td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_Value"),"txtValue") %></td>
	<td><% =getTextField("txtValue","mCurrency",decValue,12,25,"onBlur=""calcTax();""") %></td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_Tax"),"txtTax") %></td>
	<% If Application("av_TaxChange") = "1" Then %>
	<td><% =getTextField("txtTax","dCurrency",decTax,12,25,"readonly=""readonly""") %> <% =getTextField("txtTaxRate","mCurrency",Application("av_TaxRate"),4,4,"onBlur=""calcTax();""") %><span class="dFont">%</span></td>
	<% Else %>
	<td><% =getTextField("txtTax","dCurrency",decTax,12,25,"readonly=""readonly""") %> <% =getTextField("txtTaxRate","dCurrency",Application("av_TaxRate"),4,4,"readonly=""readonly""") %><span class="dFont">%</span></td>
	<% End If %>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_InvoiceDate"),"txtInvoiceSent") %></td>
	<td><% =getDateField("txtInvoiceSent","oDate",datSendDate,getIDS("IDS_InvoiceDate")) %></td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_InvoiceDue"),"txtInvoiceDue") %></td>
	<td><% =getDateField("txtInvoiceDue","oDate",datDueDate,getIDS("IDS_InvoiceDue")) %></td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_InvoicePaid"),"txtInvoicePaid") %></td>
	<td><% =getDateField("txtInvoicePaid","oDate",datPaidDate,getIDS("IDS_InvoicePaid")) %></td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_Closed"),"chkClosed") %></td>
	<td><% =getCheckbox("chkClosed",blnClosed,"onClick=""doClassChange();""") %></td>
  </tr>
<%	=strExtraFields %>
</table>

</div>

<div id="footerDiv" class="dvFooter">
<%
	If blnRS Then
		Response.Write(getIconNew(getEditURL(7,"?m="&bytMod&"&mid="&lngModId)))
		If intPerm >= 4 Then Response.Write(getIconDelete())
	End If
	Response.Write(getIconSave(strAction))
	Response.Write(getIconCancel("back"))
%>
</div>
</form>

<script language="JavaScript" type="text/javascript">
function calcTax() {
	var dTax = (document.forms[0].txtValue.value * document.forms[0].txtTaxRate.value)/100;
	document.forms[0].txtTax.value = dTax.toFixed(2);
}
</script>

<%
	Call DisplayFooter(0)
%>

