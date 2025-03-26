<!--#include file="..\_inc\functions.asp" -->
<!--#include file="..\_inc\functions_mod.asp" -->
<!--#include file="..\_inc\sql\sql_sales.asp" -->
<%
	Call pageFunctions(3,1)

	Dim lngDivId        'as Long
	Dim strFrame        'as String
	Dim lngContactId    'as Long
	Dim strContact      'as String
	Dim strDivision     'as String
	Dim strSalesRep     'as String
	Dim strPhase        'as String
	Dim strPipe         'as String
	Dim strInvoiceId    'as String
	Dim strSaleValue    'as Integer
	Dim strClosed       'as Integer

	lngRecordId = valNum(lngRecordId,3,1)

	If strDoAction = "del" and intPerm >= 4 Then
		Call delSale(lngUserId,lngRecordId)
		lngPrevId = doPrevNext(0,3,lngRecordId,0,0)
		lngNextId = doPrevNext(1,3,lngRecordId,0,0)
		strTitle = getIDS("IDS_Deleted")
	Else
		Set objRS = objConn.Execute(getSale(1,lngRecordId))

		If not (objRS.BOF and objRS.EOF) then
			lngPrevId       = objRS.fields("PrevId").value
			lngNextId       = objRS.fields("NextId").value
			strContact      = showString(objRS.fields("Contact").value)
			strTitle        = showString(objRS.fields("C_Client").value)
			strDivision     = showString(objRS.fields("D_Division").value)
			strSalesRep     = showString(objRS.fields("SalesRep").value)
			lngContactId    = objRS.fields("ContactId").value
			lngDivId        = objRS.fields("DivId").value
			strInvoiceId    = bigDigitNum(7,objRS.fields("InvoiceId").value)
			strPhase        = getAOS(objRS.fields("S_Phase").value)
			strPipe         = objRS.fields("S_Pipe").value & " %"
			strSaleValue    = FormatCurrency(objRS.fields("S_SaleValue").value) & " " & objRS.fields("S_Currency").value
			strClosed       = showTrueFalse(objRS.fields("S_Closed").value)
		Else
			strTitle = getIDS("IDS_Deleted")
		End If
	End If

	Call DisplayHeader(1)
%>

<div id="modDiv" class="dvMod">

<% Call showToolBar() %>

<table border="0" cellspacing="10" width="100%">
  <tr><td width="50%" valign=top>

  <table border=0>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Account") %></td>
	  <td class="dFont"><% =showLink(2,"client.asp?id="&lngDivId,strTitle) %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Contact") %></td>
	  <td class="dFont"><% =showLink(1,"contact.asp?id="&lngContactId,strContact) %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_InvoiceId") %></td>
	  <td class="dFont"><% =strInvoiceId %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_SaleValue") %></td>
	  <td class="dFont"><% =strSaleValue %></td>
	</tr>
  </table>

  </td>
  <td width="50%" valign=top>

  <table border=0>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_SalesRep") %></td>
	  <td class="dFont"><% =strSalesRep %></td>
	</tr>
	<tr>
	  <td nowrap class="bFont""><% =getIDS("IDS_Phase") %></td>
	  <td class="dFont"><% =strPhase %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Pipeline") %></td>
	  <td class="dFont"><% =strPipe %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Closed") %></td>
	  <td class="dFont"><% =strClosed %></td>
	</tr>
  </table>

  </td></tr>
</table>

<%
	If strTitle <> getIDS("IDS_Deleted") Then

		strTabBuilder = getIDS("IDS_Summary") & "|i_summary.asp?m=3&mid=" & lngRecordId & _
			"||" & getIDS("IDS_Notes") & "|../common/i_notes.asp?m=3&mid=" & lngRecordId & _
			"||" & getIDS("IDS_Attachments") & "|../common/i_attach.asp?m=3&mid=" & lngRecordId

		'Enable following line allows event logging
		'strTabBuilder = strTabBuilder & "||" & getIDS("IDS_Events") & "|../common/i_events.asp?m=3&mid=" & lngRecordId

		strFrame = makeTabs(strTabBuilder)

		Response.Write("</div><iframe id=""contentDiv"" class=""iBorder"" src=""" & strTabURL & """ title=""" & strFrame & """ style=""height:" & intScreenH-210 & "px;"" width=""100%"" scrolling=""no""></iframe>" & vbCrLf)
	End If

	Call DisplayFooter(1)
%>