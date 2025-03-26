<!--#include file="..\_inc\functions.asp" -->
<!--#include file="..\_inc\sql\sql_finance.asp" -->
<%
	Call pageFunctions(7,1)

	Dim strPaymentInfo	'as String
	Dim strCreatedBy	'as String
	Dim strModBy		'as String

	If bytMod = "" or lngModId = "" Then Call logError(3,1)

	Set objRS = objConn.Execute(getInvoice(0,lngModId))

	If not (objRS.BOF and objRS.EOF) then
		strPaymentInfo = showParagraph(objRS.fields("I_PayInfo").value)
		strCreatedBy = showDate(0,objRS.fields("I_CreatedDate").value) & " - " & showString(objRS.fields("CreatedBy").value)
		strModBy = showDate(0,objRS.fields("I_ModDate").value) & " - " & showString(objRS.fields("ModBy").value)
	End If

	strTitle = Application("IDS_InvoiceDetails")
	Call DisplayHeader(2)
%>

<div id="headerDiv" class="dvNoBorder">

<table border=0 cellspacing=3 width="100%">
  <tr>
    <td>
<%
	If intPerm >= 3 Then
		Response.Write(getIconEdit("Javascript:openWindow('pop_payment.asp?id=" & lngRecordId & "','sw_Payment','500','400');"))
	End If
%>
    </td>
  </tr>
</table>

<table border=0 cellspacing=0 cellpadding=0 width="100%">
  <tr class="hrow">
    <th class="hfont">&nbsp;&nbsp;&nbsp;<% =Application("IDS_InvoiceDetails") %></th>
  </tr>
</table>

</div>

<div id="contentDiv" class="dvBorder" style="height:<% =intScreenH-276 %>px;">

<span class="dfont"><% =strPaymentInfo %></font>

</div>

<%
	Call DisplayFooter(2)
%>