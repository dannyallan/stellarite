<!--#include file="..\_inc\functions.asp" -->
<!--#include file="..\_inc\sql\sql_finance.asp" -->
<%
	Call pageFunctions(7,1)

	Dim bytOrder        'as Byte
	Dim bytLastOrder    'as Byte
	Dim strSortOrder    'as String
	Dim strClass        'as String

	strTitle = getIDS("IDS_Tickets")
	bytOrder = valNum(Request.QueryString("o"),1,0)
	bytLastOrder = valNum(Request.QueryString("lo"),1,0)
	strSortOrder = valString(Request.QueryString("so"),4,0,0)

	If bytMod = "" or lngModId = "" Then Call logError(3,1)
	If bytOrder <> 2 and bytOrder <> 3 Then bytOrder = 1
	strSortOrder = getOrder(bytOrder,bytLastOrder,strSortOrder)

	Set objRS = objConn.Execute(getInvoicesBy(bytMod,lngModId,bytOrder,strSortOrder))
	If not (objRS.BOF and objRS.EOF) Then arrRS = objRS.GetRows()

	Call DisplayHeader(2)

%>

<div id="headerDiv" class="dvNoBorder">

<table border=0 cellspacing=3 width="100%">
  <tr>
	<td>
<%
	If pInvoices >= 2 Then
		Response.Write(getIconNew(getEditURL(7,"?m="&bytMod&"&mid="&lngModId)))
	End If
%>
	</td>
  </tr>
</table>

<table border=0 cellspacing=0 cellpadding=0 width="100%">
  <tr class="hRow">
	<th class="hFont lIndent" width="29%"><a href="i_invoices.asp?m=<% =bytMod %>&mid=<% =lngModId %>&o=1&lo=<% =bytOrder %>&so=<% =strSortOrder %>"><% =getIDS("IDS_InvoiceId") %></a></th>
	<th class="hFont" width="38%"><a href="i_invoices.asp?m=<% =bytMod %>&mid=<% =lngModId %>&o=2&lo=<% =bytOrder %>&so=<% =strSortOrder %>"><% =getIDS("IDS_InvoiceDue") %></a></th>
	<th class="hFont" width="27%"><a href="i_invoices.asp?m=<% =bytMod %>&mid=<% =lngModId %>&o=3&lo=<% =bytOrder %>&so=<% =strSortOrder %>"><% =getIDS("IDS_Closed") %></a></th>
	<th class="hFont" width="5%">&nbsp;</th>
  </tr>
</table>

</div>

<div id="contentDiv" class="dvNoBorder" style="height:<% =intScreenH-276 %>px;">
<table border=0 cellspacing=0 cellpadding=0 width="100%">
<%
		If isArray(arrRS) Then
			For i = 0 to UBound(arrRS,2)
				strClass = toggleRowColor(strClass)
%>
  <tr class="<%=strClass%>">
	<td class="dFont lIndent" width="29%"><a href="../finance/invoice.asp?id=<% =arrRS(0,i) %>" target="_top"><% =bigDigitNum(7,arrRS(0,i)) %></a></td>
	<td class="dFont" width="38%"><% =showDate(0,arrRS(1,i)) %></td>
	<td class="dFont" width="27%"><% If arrRS(2,i) = 1 Then Response.Write(getIDS("IDS_True")) Else Response.Write(getIDS("IDS_False")) %></td>
	<td class="dFont rIndent" width="5%" align=right>&nbsp;
<%
				If pInvoices >= 3 Then
					Response.Write(vbTab & getIconImport(2,getEditURL(7,"?id="&arrRS(0,i)&"&m="&bytMod&"&mid="&lngModId),bigDigitNum(7,arrRS(0,i))) & vbCrLf)
				End If
%>
	</td>
  </tr>
<%            Next
		End If
%>
</table>
</div>
<%
	Call DisplayFooter(2)
%>

