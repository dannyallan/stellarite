<!--#include file="..\_inc\functions.asp" -->
<!--#include file="..\_inc\functions_list.asp" -->
<%
	Call pageFunctions(5,1)

	Dim strClass        'as String
	Dim strClassUrgent  'as String
	Dim intMaxRecords   'as Integer
	Dim intDivSize      'as Integer

	strTitle = getIDS("IDS_ModName5")

	intMaxRecords = 20
	intDivSize = CInt(intScreenH-100)

	Call DisplayHeader(1)
%>

<div id="contentDiv" class="dvBorder">

	<div style="float:left;width:75%;">
<%
	Call showList(15,intDivSize,intMaxRecords,3)
%>
	</div>

	<div style="margin-left:75%;padding-left:10px;">
	  <div id="div13" class="dvRightMenu" style="height:<% =CInt((intScreenH-70)*0.3) %>px;">
		<div class="hRow hScr hFont"><% =getIDS("IDS_UsefulLinks") %></div>
		<div>
		  <ul>
			<% If pTickets > 2 Then %>
			<li><% =getEditLink(5,"",getIDS("IDS_TicketNew")) %></li>
			<div class="hr"></div>
			<% End If %>
			<li><% =getPopLink("S","?m=5",getIDS("IDS_TicketSearch")) %></li>
			<div class="hr"></div>
			<li><a href="../common/calendar.asp?m=5"><% =getIDS("IDS_Calendar") %></a></li>
		  </ul>
		</div>
	  </div>

<%	  Call showList(31,CInt((intScreenH-70)*0.6),"5",1) %>

	</div>
</div>

</div>

<%
	Call DisplayFooter(1)
%>
