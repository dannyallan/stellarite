<!--#include file="..\_inc\functions.asp" -->
<!--#include file="..\_inc\sql\sql_products.asp" -->
<%
	Call pageFunctions(0,1)

	Dim bytOrder        'as Byte
	Dim bytLastOrder    'as Byte
	Dim strSortOrder    'as String
	Dim strClass        'as String

	strTitle = getIDS("IDS_Products")
	bytOrder = valNum(Request.QueryString("o"),1,0)
	bytLastOrder = valNum(Request.QueryString("lo"),1,0)
	strSortOrder = valString(Request.QueryString("so"),4,0,0)

	If bytMod = "" or lngModId = "" Then Call logError(3,1)
	If bytOrder <> 2 and bytOrder <> 3 and bytOrder <> 4 Then bytOrder = 1
	strSortOrder = getOrder(bytOrder,bytLastOrder,strSortOrder)

	Set objRS = objConn.Execute(getProducts(bytMod,lngModId,bytOrder,strSortOrder))
	If not (objRS.BOF and objRS.EOF) Then arrRS = objRS.GetRows()

	Call DisplayHeader(2)

%>
<div id="headerDiv" class="dvNoBorder">

<table border=0 cellspacing=3 width="100%">
  <tr>
	<td>
<%
	If intPerm >= 2 Then
		Response.Write(getIconNew(getEditURL("Z","?m=" & bytMod & "&mid=" & lngModId)))
	End If
%>
	</td>
  </tr>
</table>

<table border=0 cellspacing=0 cellpadding=0 width="100%">
  <tr class="hRow">
	<th class="hFont lIndent" width="14%"><a href="i_products.asp?m=<% =bytMod %>&mid=<% =lngModId %>&o=1&lo=<% =bytOrder %>&so=<% =strSortOrder %>"><% =getIDS("IDS_Modified") %></a></th>
	<th class="hFont" width="35%"><a href="i_products.asp?m=<% =bytMod %>&mid=<% =lngModId %>&o=2&lo=<% =bytOrder %>&so=<% =strSortOrder %>"><% =getIDS("IDS_Product") %></a></th>
	<th class="hFont" width="15%"><a href="i_products.asp?m=<% =bytMod %>&mid=<% =lngModId %>&o=3&lo=<% =bytOrder %>&so=<% =strSortOrder %>"><% =getIDS("IDS_Serial") %></a></th>
	<th class="hFont" width="15%"><a href="i_products.asp?m=<% =bytMod %>&mid=<% =lngModId %>&o=4&lo=<% =bytOrder %>&so=<% =strSortOrder %>"><% =getIDS("IDS_PIN") %></a></th>
	<th class="hFont" width="15%"><a href="i_products.asp?m=<% =bytMod %>&mid=<% =lngModId %>&o=5&lo=<% =bytOrder %>&so=<% =strSortOrder %>"><% =getIDS("IDS_Expiry") %></a></th>
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
	<td class="dFont lIndent" width="14%"><% =showDate(0,arrRS(1,i)) %></td>
	<td class="dFont" width="35%"><% =trimString(arrRS(2,i),40) %></td>
	<td class="dFont" width="15%"><% =showString(arrRS(3,i)) %></td>
	<td class="dFont" width="15%"><% =showString(arrRS(4,i)) %></td>
	<td class="dFont" width="15%"><% =showDate(0,arrRS(5,i)) %></td>
	<td class="dFont rIndent" width="5%" align=right>&nbsp;
<%
				If intPerm >= 3 Then
					Response.Write(vbTab & getIconImport(2,getEditURL("Z","?id=" & arrRS(0,i) & "&m=" & bytMod & "&mid=" & lngModId),showString(arrRS(3,i))) & vbCrLf)
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