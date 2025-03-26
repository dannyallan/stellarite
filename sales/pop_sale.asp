<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_sales.asp" -->
<!--#include file="..\_inc\currency.asp" -->
<%
	Call pageFunctions(3,2)

	Dim strSalesRep		'as String
	Dim lngSalesRepId	'as Long
	Dim lngInvoiceId	'as Long
	Dim strPhase		'as String
	Dim intPhase		'as Integer
	Dim intPipe			'as Integer
	Dim strCurrency		'as String
	Dim decSaleValue	'as Decimal
	Dim blnClosed		'as Boolean
	Dim strContact		'as String
	Dim lngContactId	'as Long
	Dim lngDivId		'as Long
	Dim datCloseDate	'as Date
	Dim strCreatedBy	'as String
	Dim datCreatedDate	'as Date
	Dim strModBy		'as String
	Dim datModDate		'as Date

	strTitle = Application("IDS_Edit") & " " & Application("IDS_Sale")

	If strDoAction <> "" then

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

		If strDoAction = "del" and intPerm >= 4 Then

			Call delSale(lngUserId,lngRecordId)

		Elseif strDoAction = "edit" and intPerm >= 3 then

			Call updateSale(lngUserId,lngRecordId,bytMod,lngModId,lngDivId,lngContactId,intPhase,intPipe,lngSalesRepId,lngInvoiceId,blnClosed,datCloseDate,strCurrency,decSaleValue)

		ElseIf strDoAction = "new" Then

			lngRecordId = insertSale(lngUserId,lngRecordId,bytMod,lngModId,lngDivId,lngContactId,intPhase,intPipe,lngSalesRepId,lngInvoiceId,blnClosed,datCloseDate,strCurrency,decSaleValue)
		End If
		Call closeWindow(strOpenerURL)

	Else
		If blnRS and intPerm >= 3 Then
			Set objRS = objConn.Execute(getSale(0,lngRecordId))

		If not (objRS.BOF and objRS.EOF) then
				strContact = showString(objRS.fields("Contact").value)
				lngContactId = objRS.fields("ContactId").value
				lngDivId = objRS.fields("DivId").value
				lngInvoiceId = showString(objRS.fields("InvoiceId").value)
				intPhase = objRS.fields("S_Phase").value
				intPipe = objRS.fields("S_Pipe").value
				strSalesRep = showString(objRS.fields("SalesRep").value)
				blnClosed = objRS.fields("S_Closed").value
				datCloseDate = showDate(0,objRS.fields("S_CloseDate").value)
				strCurrency = showString(objRS.fields("S_Currency").value)
				decSaleValue = objRS.fields("S_SaleValue").value
				strCreatedBy = showString(objRS.fields("CreatedBy").value)
				datCreatedDate = objRS.fields("S_CreatedDate").value
				strModBy = showString(objRS.fields("ModBy").value)
				datModDate = objRS.fields("S_ModDate").value
			End If

		Elseif blnRS Then
			Call doRedirect("pop_sale.asp?m=" & bytMod & "&mid=" & lngModId)
		Else
			Select Case bytMod
				Case 1
					lngContactId = lngModId
					lngDivId = getValue("DivId","CRM_Contacts","ContactId = "&lngContactId,0)
					strContact = getValue(doConCat(doConCat("K_FirstName","' '"),"K_LastName"),"CRM_Contacts","ContactId = "&lngContactId,"")
				Case 2
					lngDivId = lngModId
					lngContactId = getValue("ContactId","CRM_Contacts","DivId = "&lngDivId,"")
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
		If strCurrency = "" Then strCurrency = Application("av_Currency")
		strPhase = getOptionDropDown(260,False,"selPhase","Sales Phase",intPhase)
	End If

	strIncHead = getCalendarScripts()

	Call DisplayHeader(3)
	Call showEditHeader(strTitle,strCreatedBy,datCreatedDate,strModBy,datModDate)
%>
<div id="contentDiv" class="dvBorder" style="height:330px;"><br>

<table border=0 cellspacing=5 width="100%">
<form name="frmSale" method="post" action="pop_sale.asp?id=<% =lngRecordId %>&m=<% =bytMod %>&mid=<% =lngModId %>">
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
    <a href="<% =newWindow(1,"?m=3") %>"><img src="../images/new2.gif" alt="<% =getImport("IDS_ContactNew") %>" border=0 height=16 width=16></a></td>
    <% End If %>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_SalesRep"),"txtSalesRep") %></td>
    <td><% =getTextField("txtSalesRep","mText",strSalesRep,40,100,"") %>
    <a href="<% =newWindow("S","?m=0&rVal=txtSalesRep") %>"><img src="../images/import.gif" alt="<% =getImport("IDS_SalesRep") %>" border=0 height=16 width=16></a></td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_Phase"),"selPhase") %></td>
    <td><% =strPhase %></td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_Pipeline"),"rdoPipe10") %></td>
    <td class="dfont">
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
      <table border=0 cellspacing=0 cellpadding=0 width="90%">
        <tr>
          <td class="dfont">0%</td>
          <td class="dfont" align=right>100%</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_Currency"),"selCurrency") %></td>
    <td><% =getCurrency(260,"selCurrency",strCurrency) %></td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_SaleValue"),"txtSaleValue") %></td>
    <td><% =getTextField("txtSaleValue","oCurrency",decSaleValue,12,255,"") %></td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_InvoiceId"),"txtInvoice") %></td>
    <td><% =getTextField("txtInvoice","oLong",lngInvoiceId,12,20,"") %>
    <% If pInvoices >= 1 Then %>
    <a href="<% =newWindow("S","?m=7&rVal=txtInvoice") %>"><img src="../images/import.gif" alt="<% =getImport("IDS_InvoiceId") %>" border=0 height=16 width=16></a></td>
    <% End If %>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_CloseDate"),"txtCloseDate") %></td>
    <td><% =getTextField("txtCloseDate","oDate",datCloseDate,12,12,"") %>
    <a href="Javascript:showCalendar('txtCloseDate');"><img src="../images/cal.gif" alt="<% =getImport("IDS_CloseDate") %>" border=0 height=16 width=16></a></td>
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
		Response.Write(getIconNew("pop_sale.asp"))
		If intPerm >= 4 Then Response.Write(getIconDelete())
	End If
	Response.Write(getIconSave(strAction))
	Response.Write(getIconCancel())
%>
</div>

<script language="Javascript">
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
	Call DisplayFooter(3)
%>

