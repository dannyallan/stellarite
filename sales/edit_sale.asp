<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_sales.asp" -->
<!--#include file="..\_inc\currency.asp" -->
<%
	Call pageFunctions(3,2)

	Dim strSalesRep     'as String
	Dim lngSalesRepId   'as Long
	Dim lngInvoiceId    'as Long
	Dim strPhase        'as String
	Dim intPhase        'as Integer
	Dim intPipe         'as Integer
	Dim strCurrency     'as String
	Dim decSaleValue    'as Decimal
	Dim blnClosed       'as Boolean
	Dim strContact      'as String
	Dim lngContactId    'as Long
	Dim lngDivId        'as Long
	Dim datCloseDate    'as Date
	Dim strCreatedBy    'as String
	Dim datCreatedDate  'as Date
	Dim strModBy        'as String
	Dim datModDate      'as Date
	Dim strExtraFields  'as String

	strTitle = getIDS("IDS_Edit") & " " & getIDS("IDS_Sale")

	If strDoAction <> "" then

		Select Case strDoAction
			Case "del"
				If intPerm >= 4 Then Call delSale(lngUserId,lngRecordId)

			Case "new","edit"

				lngContactId = valNum(Request.Form("hdnContact"),3,1)
				lngDivId = valNum(Request.Form("hdnDivision"),3,1)
				strSalesRep = valString(Request.Form("txtSalesRep"),100,1,0)
				lngSalesRepId = getUserId(3,strSalesRep)
				intPhase = valNum(Request.Form("selPhase"),2,-1)
				intPipe = valNum(Request.Form("rdoPipe"),1,1)
				blnClosed = valNum(Request.Form("chkClosed"),0,0)
				datCloseDate = valDate(Request.Form("txtCloseDate"),blnClosed)
				lngInvoiceId = valNum(Request.Form("txtInvoice"),3,-1)
				strCurrency = valString(Request.Form("selCurrency"),3,1,0)
				decSaleValue = valNum(Request.Form("txtSaleValue"),5,blnClosed)
				If blnClosed = "1" Then intPipe = 100

				If strDoAction = "edit" and intPerm >= 3 then

					Call updateSale(lngUserId,lngRecordId,bytMod,lngModId,lngDivId,lngContactId,intPhase,intPipe,lngSalesRepId,lngInvoiceId,blnClosed,datCloseDate,strCurrency,decSaleValue)

				ElseIf strDoAction = "new" Then

					lngRecordId = insertSale(lngUserId,lngRecordId,bytMod,lngModId,lngDivId,lngContactId,intPhase,intPipe,lngSalesRepId,lngInvoiceId,blnClosed,datCloseDate,strCurrency,decSaleValue)
				End If

				Call saveCustomFields(3,lngRecordId)
		End Select

		If bytMenu = 0 Then Session("LastPage") = strCRMURL & "common/i_sales.asp?m=" & bytMod & "&mid=" & lngModId
		Call closeEdit()
	Else
		If blnRS and intPerm >= 3 Then
			Set objRS = objConn.Execute(getSale(0,lngRecordId))

			If not (objRS.BOF and objRS.EOF) then
				lngContactId = objRS.fields("ContactId").value
				strContact = objRS.fields("Contact").value
				lngDivId = objRS.fields("DivId").value
				lngInvoiceId = objRS.fields("InvoiceId").value
				intPhase = objRS.fields("S_Phase").value
				intPipe = objRS.fields("S_Pipe").value
				strSalesRep = objRS.fields("SalesRep").value
				blnClosed = objRS.fields("S_Closed").value
				datCloseDate = objRS.fields("S_CloseDate").value
				strCurrency = objRS.fields("S_Currency").value
				decSaleValue = objRS.fields("S_SaleValue").value
				strCreatedBy = objRS.fields("CreatedBy").value
				datCreatedDate = objRS.fields("S_CreatedDate").value
				strModBy = objRS.fields("ModBy").value
				datModDate = objRS.fields("S_ModDate").value
			End If

		Elseif blnRS Then
			Call logError(2,1)
		Else
			Select Case bytMod
				Case 1
					lngContactId = lngModId
					lngDivId = getValue("DivId","CRM_Contacts","ContactId = "&lngContactId,0)
					strContact = getValue(doConCat(doConCat("K_FirstName","' '"),"K_LastName"),"CRM_Contacts","ContactId = " & lngContactId,"")
				Case 2
					lngDivId = lngModId
					lngContactId = getValue("ContactId","CRM_Contacts","DivId = "&lngDivId,0)
					strContact = getValue(doConCat(doConCat("K_FirstName","' '"),"K_LastName"),"CRM_Contacts","ContactId = " & lngContactId,"")
				Case 7
					lngDivId = getValue("DivId","CRM_Invoices","InvoiceId="&bytMod,0)
					lngContactId = getValue("ContactId","CRM_Invoices","InvoiceId="&lngModId,0)
					strContact = getValue(doConCat(doConCat("K_FirstName","' '"),"K_LastName"),"CRM_Contacts","ContactId = " & lngContactId,"")
					lngInvoiceId = lngModId
			End Select
			intPipe = 10
			If mSales Then strSalesRep = strFullName
			strCreatedBy = strFullName
			datCreatedDate = Now
			strModBy = strFullName
			datModDate = Now
		End If
		strExtraFields = editCustomFields(3)
		If strCurrency = "" Then strCurrency = Application("av_Currency")
		strPhase = getOptionDropDown(260,False,"selPhase","IDS_SalesPhase",intPhase)
	End If

	strIncHead = getCalendarScripts()

	Call DisplayHeader(0)
	Call showEditHeader(strTitle,strCreatedBy,datCreatedDate,strModBy,datModDate)
%>
<form name="frmSale" method="post" action="edit_sale.asp?id=<% =lngRecordId %>&m=<% =bytMod %>&mid=<% =lngModId %>&menu=<% =bytMenu %>">
<div id="contentDiv" class="dvBorder" style="height:<% =intScreenH-170 %>px;"><br />

<table border=0 width="100%">
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
<% =getHidden("hdnContact",lngContactId) %>
<% =getHidden("hdnDivision",lngDivId) %>
  <tr>
	<td width=170><% =getLabel(getIDS("IDS_Contact"),"txtContact") %></td>
	<td colspan=3><% =getTextField("txtContact","mText",strContact,40,100,"readonly=""readonly""") %>

<%	If pContacts >= 1 Then
		Response.Write(getIconImport(1,getSearchURL("?m=1&rVal=K"),getIDS("IDS_Contact")))
	End If
	If pContacts >= 2 Then
		Response.Write(getIconImport(3,getEditURL(1,"?m="&bytMod&"&mid="&lngModId),getIDS("IDS_Contact")))
	End If
%>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_SalesRep"),"txtSalesRep") %></td>
	<td><% =getTextField("txtSalesRep","mText",strSalesRep,40,100,"") %>
	<% =getIconImport(1,getSearchURL("?m=0&rVal=txtSalesRep"),getIDS("IDS_Sale")) %>
	</td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_Phase"),"selPhase") %></td>
	<td><% =strPhase %></td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_Pipeline"),"rdoPipe10") %></td>
	<td class="dFont">
<%
	For i = 1 to 10
		Response.Write(vbTab & getRadio("rdoPipe",i*10,intPipe,"") & vbCrLf)
	Next
%>
	</td>
  </tr>
  <tr>
	<td></td>
	<td>
	  <table border=0 cellspacing=0 cellpadding=0 width=260>
		<tr>
		  <td class="dFont">0%</td>
		  <td class="dFont" align=right>100%</td>
		</tr>
	  </table>
	</td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_Currency"),"selCurrency") %></td>
	<td><% =getCurrency(260,"selCurrency",strCurrency) %></td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_SaleValue"),"txtSaleValue") %></td>
	<td><% =getTextField("txtSaleValue","oCurrency",decSaleValue,12,255,"") %></td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_InvoiceId"),"txtInvoice") %></td>
	<td><% =getTextField("txtInvoice","oLong",lngInvoiceId,12,20,"") %>
<%
	If pInvoices >= 1 Then
		Response.Write(getIconImport(1,getSearchURL("?m=7&rVal=txtInvoice"),getIDS("IDS_Invoice")))
	End If
%>
	</td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_CloseDate"),"txtCloseDate") %></td>
	<td><% =getDateField("txtCloseDate","oDate",datCloseDate,getIDS("IDS_CloseDate")) %></td>
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
		Response.Write(getIconNew(getEditURL(3,"?m="&bytMod&"&mid="&lngModId)))
		If intPerm >= 4 Then Response.Write(getIconDelete())
	End If
	Response.Write(getIconSave(strAction))
	Response.Write(getIconCancel("back"))
%>
</div>
</form>

<script language="JavaScript" type="text/javascript">
	doClassChange();
	function doClassChange() {
		if (document.forms[0].chkClosed.checked) {
			if (document.forms[0].txtCloseDate.value == "")
				document.forms[0].txtCloseDate.value = getToday();
			document.forms[0].txtSaleValue.className = "mCurrency";
			document.forms[0].txtCloseDate.className = "mDate";
			document.forms[0].rdoPipe[9].checked = "1";
		}
		else {
			document.forms[0].txtSaleValue.className = "oCurrency";
			document.forms[0].txtCloseDate.className = "oDate";
		}
	}
</script>

<%
	Call DisplayFooter(0)
%>

