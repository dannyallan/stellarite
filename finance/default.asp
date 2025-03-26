<!--#include file="..\_inc\functions.asp" -->
<!--#include file="..\_inc\functions_list.asp" -->
<%
	Call pageFunctions(7,1)

	Dim strClass        'as String
	Dim intMaxRecords   'as Integer
	Dim intDivSize      'as Integer

	strTitle = getIDS("IDS_ModName7")

	intMaxRecords = 20
	intDivSize = CInt((intScreenH-70)/2)

	Call DisplayHeader(1)
%>

<div id="contentDiv" class="dvBorder">

	<div style="float:left;width:75%;">
<%
	Call showList(18,intDivSize,intMaxRecords,3)
	Call showList(19,intDivSize,intMaxRecords,3)
%>
	</div>

	<div style="margin-left:75%;padding-left:10px;">
	  <div id="div13" class="dvRightMenu" style="height:<% =CInt((intScreenH-70)*0.3) %>px;">
		<div class="hRow hScr hFont"><% =getIDS("IDS_UsefulLinks") %></div>
		<div>
		  <ul>
			<% If pInvoices > 2 Then %>
			<li><% =getEditLink(7,"",getIDS("IDS_InvoiceNew")) %></li>
			<div class="hr"></div>
			<% End If %>
			<li><% =getPopLink("S","?m=7",getIDS("IDS_InvoiceSearch")) %></li>
			<div class="hr"></div>
			<li><a href="../common/calendar.asp?m=7"><% =getIDS("IDS_Calendar") %></a></li>
		  </ul>
		</div>
	  </div>

<%	  Call showList(31,CInt((intScreenH-70)*0.6),"7",1) %>

	</div>
</div>

<%
	Call DisplayFooter(1)
%>