<!--#include file="..\_inc\functions.asp" -->
<!--#include file="..\_inc\functions_mod.asp" -->
<!--#include file="..\_inc\sql\sql_finance.asp" -->
<%
	Call pageFunctions(7,1)

	Dim strClient       'as String
	Dim lngDivId        'as Long
	Dim strContact      'as String
	Dim lngContactId    'as Long
	Dim strOwner        'as String
	Dim strPO           'as String
	Dim strType         'as String
	Dim strPhase        'as String
	Dim strValue        'as String
	Dim strTax          'as String
	Dim datSendDate     'as Date
	Dim datDueDate      'as Date
	Dim datPaidDate     'as Date
	Dim strClosed       'as String
	Dim strFrame        'as String

	lngRecordId = valNum(lngRecordId,3,1)
	strTitle = bigDigitNum(7,lngRecordId)

	If strDoAction = "del" and intPerm >= 4 then
		Call delInvoice(lngUserId,lngRecordId)
		lngPrevId = doPrevNext(0,4,lngRecordId,0,0)
		lngNextId = doPrevNext(1,4,lngRecordId,0,0)
		strTitle = getIDS("IDS_Deleted")
	Else
		Set objRS = objConn.Execute(getInvoice(1,lngRecordId))

		If not (objRS.BOF and objRS.EOF) then
			lngPrevId = objRS.fields("PrevId").value
			lngNextId = objRS.fields("NextId").value
			strClient = showString(objRS.fields("C_Client").value)
			lngDivId = objRS.fields("DivId").value
			strContact = showString(objRS.fields("Contact").value)
			lngContactId = objRS.fields("ContactId").value
			strOwner = showString(objRS.fields("Owner").value)
			strPO = showString(objRS.fields("I_PurchaseOrder").value)
			strType = getAOS(objRS.fields("I_Type").value)
			strPhase = getAOS(objRS.fields("I_Phase").value)
			strValue = showString(objRS.fields("I_Currency").value & " " & FormatCurrency(objRS.fields("I_Value").value))
			strTax = showString(objRS.fields("I_Currency").value & " " & FormatCurrency(objRS.fields("I_Tax").value))
			datSendDate = showDate(0,objRS.fields("I_SendDate").value)
			datDueDate = showDate(0,objRS.fields("I_DueDate").value)
			datPaidDate = showDate(0,objRS.fields("I_PaidDate").value)
			strClosed = showTrueFalse(objRS.fields("I_Closed").value)

			If objRS.fields("I_Received").value = 1 Then
				strType = strType & " [" & getIDS("IDS_InvoiceReceived") & "]"
			Else
				strType = strType & " [" & getIDS("IDS_InvoiceSent") & "]"
			End if
		Else
			strTitle = getIDS("IDS_Deleted")
		End If
	End If

	Call DisplayHeader(1)
%>

<div id="modDiv" class="dvMod">

<% Call showToolBar() %>

<table border="0" cellspacing="10" width="100%">
  <tr><td width="34%" valign=top>

  <table border=0>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Account") %></td>
	  <td class="dFont"><% =showLink(2,"../sales/client.asp?id="&lngDivId,strClient) %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Contact") %></td>
	  <td class="dFont"><% =showLink(1,"../sales/contact.asp?id="&lngContactId,strContact) %>&nbsp;</td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Owner") %></td>
	  <td class="dFont"><% =strOwner %>&nbsp;</td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_PurchaseOrder") %></td>
	  <td class="dFont"><% =strPO %></td>
	</tr>
  </table>

  </td>
  <td width="33%" valign=top>

  <table border=0>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Type") %></td>
	  <td class="dFont" width=200><% =strType %>&nbsp;</td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Phase") %></td>
	  <td class="dFont"><% =strPhase %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Value") %></td>
	  <td class="dFont"><% =strValue %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Tax") %></td>
	  <td class="dFont"><% =strTax %></td>
	</tr>
  </table>

  </td>
  <td width="33%" valign=top>

  <table border=0>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_InvoiceDate") %></td>
	 <td class="dFont" width=200><% =datSendDate %>&nbsp;</td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_InvoiceDue") %></td>
	  <td class="dFont"><% =datDueDate %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_InvoicePaid") %></td>
	  <td class="dFont"><% =datPaidDate %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Closed") %></td>
	  <td class="dFont"><% =strClosed %></td>
	</tr>
  </table>

  </td></tr>
</table>

<%
	If strTitle <> "Deleted" Then

		strTabBuilder = getIDS("IDS_Summary") & "|i_summary.asp?m=7&mid=" & lngRecordId & _
			"||" & getIDS("IDS_Notes") & "|../common/i_notes.asp?m=7&mid=" & lngRecordId & _
			"||" & getIDS("IDS_Attachments") & "|../common/i_attach.asp?m=7&mid=" & lngRecordId


		'Enable following line allows event logging
		'strTabBuilder = strTabBuilder & "||" & getIDS("IDS_Events") & "|../common/i_events.asp?m=7&mid=" & lngRecordId

		If pSales >= 1 Then strTabBuilder = strTabBuilder & "||" & getIDS("IDS_Sales") & "|../common/i_sales.asp?m=7&mid=" & lngRecordId
		strTabBuilder = strTabBuilder & "||" & getIDS("IDS_Products") & "|../common/i_products.asp?m=7&mid=" & lngRecordId
		If pProjects >= 1 Then strTabBuilder = strTabBuilder & "||" & getIDS("IDS_Projects") & "|../common/i_projects.asp?m=7&mid=" & lngRecordId

		strFrame = makeTabs(strTabBuilder)

		Response.Write("</div><iframe id=""contentDiv"" class=""iBorder"" src=""" & strTabURL & """ title=""" & strFrame & """ style=""height:" & intScreenH-210 & "px;"" width=""100%"" scrolling=""no""></iframe>" & vbCrLf)
	End If

	Call DisplayFooter(1)
%>