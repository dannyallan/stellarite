<!--#include file="..\_inc\functions.asp" -->
<!--#include file="..\_inc\sql\sql_sales.asp" -->
<%
	Call pageFunctions(3,1)

	Dim bytOrder        'as Byte
	Dim bytLastOrder    'as Byte
	Dim strSortOrder    'as String
	Dim strClass        'as String

	strTitle = getIDS("IDS_Sales")
	bytOrder = valNum(Request.QueryString("o"),1,0)
	bytLastOrder = valNum(Request.QueryString("lo"),1,0)
	strSortOrder = valString(Request.QueryString("so"),4,0,0)

	If bytMod = "" or lngModId = "" Then Call logError(3,1)
	If bytOrder <> 2 and bytOrder <> 3 Then bytOrder = 1
	strSortOrder = getOrder(bytOrder,bytLastOrder,strSortOrder)

	Set objRS = objConn.Execute(getSalesByMod(bytMod,lngModId,bytOrder,strSortOrder))
	If not (objRS.EOF and objRS.BOF) Then arrRS = objRS.GetRows()

	Call DisplayHeader(2)

%>

<div id="headerDiv" class="dvNoBorder">

<table border=0 cellspacing=3 width="100%">
  <tr>
	<td>
<%
	If pSales >= 2 Then
		Response.Write(getIconNew(getEditURL(3,"?m="&bytMod&"&mid="&lngModId)))
	End If
%>
	</td>
  </tr>
</table>

<table border=0 cellspacing=0 cellpadding=0 width="100%">
  <tr class="hRow">
	<th class="hFont lIndent" width="22%"><a href="i_sales.asp?m=<% =bytMod %>&mid=<% =lngModId %>&o=1&lo=<% =bytOrder %>&so=<% =strSortOrder %>"><% =getIDS("IDS_Modified") %></a></th>
	<th class="hFont" width="31%"><a href="i_sales.asp?m=<% =bytMod %>&mid=<% =lngModId %>&o=2&lo=<% =bytOrder %>&so=<% =strSortOrder %>"><% =getIDS("IDS_SalesRep") %></a></th>
	<th class="hFont" width="41%"><a href="i_sales.asp?m=<% =bytMod %>&mid=<% =lngModId %>&o=3&lo=<% =bytOrder %>&so=<% =strSortOrder %>"><% =getIDS("IDS_CloseDate") %></a></th>
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
	<td class="dFont lIndent" width="22%"><a href="../sales/sale.asp?id=<% =arrRS(0,i) %>" target="_top"><% =bigDigitNum(7,arrRS(0,i)) %><a/></td>
	<td class="dFont" width="31%"><% =showString(arrRS(1,i)) %></td>
	<td class="dFont" width="41%"><% =showDate(0,arrRS(2,i)) %></td>
	<td class="dFont rIndent" width="5%" align=right>&nbsp;
<%
				If pSales >= 3 Then
					Response.Write(vbTab & getIconImport(2,getEditURL(3,"?id="&arrRS(0,i)&"&m="&bytMod&"&mid="&lngModId),bigDigitNum(7,arrRS(0,i))) & vbCrLf)
				End If
%>
	</td>
  </tr>
<%        Next
	End If
%>
</table>
</div>
<%
	Call DisplayFooter(2)
%>