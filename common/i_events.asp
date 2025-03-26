<!--#include file="..\_inc\functions.asp" -->
<!--#include file="..\_inc\sql\sql_events.asp" -->
<%
	Call pageFunctions(50,1)

	Dim bytOrder        'as Byte
	Dim bytLastOrder    'as Byte
	Dim strSortOrder    'as String
	Dim strClass        'as String
	Dim blnLook         'as Boolean

	strTitle = getIDS("IDS_Events")
	bytOrder = valNum(Request.QueryString("o"),1,0)
	bytLastOrder = valNum(Request.QueryString("lo"),1,0)
	strSortOrder = valString(Request.QueryString("so"),4,0,0)
	blnLook = True

	If bytMod = "" or lngModId = "" Then Call logError(3,1)
	If bytOrder <> 2 and bytOrder <> 3 Then bytOrder = 1
	strSortOrder = getOrder(bytOrder,bytLastOrder,strSortOrder)


	Set objRS = objConn.Execute(getEventsBy(bytMod,lngModId,intMember,bytOrder,strSortOrder))
	If not (objRS.BOF and objRS.EOF) Then arrRS = objRS.GetRows()

	Call DisplayHeader(2)

%>

<div id="headerDiv" class="dvNoBorder">

<table border=0 cellspacing=3 width="100%">
  <tr>
	<td>
<%
	If intPerm >= 2 Then
		Response.Write(getIconNew(getEditURL(50,"?m="&bytMod&"&mid="&lngModId)))
	End If
%>
	</td>
  </tr>
</table>

<table border=0 cellspacing=0 cellpadding=0 width="100%">
  <tr class="hRow">
	<th class="hFont lIndent" width="22%"><a href="i_events.asp?m=<% =bytMod %>&mid=<% =lngModId %>&o=1&lo=<% =bytOrder %>&so=<% =strSortOrder %>"><% =getIDS("IDS_StartDate") %></a></th>
	<th class="hFont" width="47%"><a href="i_events.asp?m=<% =bytMod %>&mid=<% =lngModId %>&o=2&lo=<% =bytOrder %>&so=<% =strSortOrder %>"><% =getIDS("IDS_Title") %></a></th>
	<th class="hFont" width="20%"><a href="i_events.asp?m=<% =bytMod %>&mid=<% =lngModId %>&o=3&lo=<% =bytOrder %>&so=<% =strSortOrder %>"><% =getIDS("IDS_Type") %></a></th>
	<th class="hFont" width="10%">&nbsp;</td>
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
	<td class="dFont lIndent" width="22%"><% =showDate(0,arrRS(3,i)) %></td>
	<td class="dFont" width="47%">
	<a href="event.asp?id=<% =arrRS(0,i) %>&m=<% =bytMod %>&mid=<% =lngModId %>" target="_top"><% =showString(arrRS(1,i)) %></a>
	</td>
	<td class="dFont" width="20%"><% =showString(arrRS(2,i)) %></td>
	<td class="dFont rIndent" width="10%" align=right>&nbsp;
<%
				If intPerm >= 3 Then
					Response.Write(vbTab & getIconImport(2,getEditURL(50,"?id="&arrRS(0,i)&"&m="&bytMod&"&mid="&lngModId),showString(arrRS(1,i))) & vbCrLf)
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