<!--#include file="..\_inc\functions.asp" -->
<!--#include file="..\_inc\functions_list.asp" -->
<%
	Call pageFunctions(6,1)

	Dim strClass        'as String
	Dim strClassUrgent  'as String
	Dim intMaxRecords   'as Integer
	Dim intDivSize      'as Integer

	strTitle = getIDS("IDS_ModName6")

	intMaxRecords = 20
	intDivSize = CInt(intScreenH-100)

	Call DisplayHeader(1)
%>

<div id="contentDiv" class="dvBorder">

	<div style="float:left;width:75%;">
<%
	Call showList(17,intDivSize,intMaxRecords,3)
%>
	</div>

	<div style="margin-left:75%;padding-left:10px;">
	  <div id="div13" class="dvRightMenu" style="height:<% =CInt((intScreenH-70)*0.3) %>px;">
		<div class="hRow hScr hFont"><% =getIDS("IDS_UsefulLinks") %></div>
		<div>
		  <ul>
			<% If pBugs > 2 Then %>
			<li><% =getEditLink(6,"",getIDS("IDS_BugNew")) %></li>
			<div class="hr"></div>
			<% End If %>
			<li><% =getPopLink("S","?m=6",getIDS("IDS_BugSearch")) %></li>
			<div class="hr"></div>
			<li><a href="../common/calendar.asp?m=6"><% =getIDS("IDS_Calendar") %></a></li>
		  </ul>
		</div>
	  </div>

<%	  Call showList(31,CInt((intScreenH-70)*0.6),"6",1) %>

	</div>
</div>

<%
	Call DisplayFooter(1)
%>
