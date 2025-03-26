<!--#include file="..\_inc\functions.asp" -->
<!--#include file="..\_inc\functions_list.asp" -->
<%
	Call pageFunctions(4,1)

	Dim strClass        'as String
	Dim intMaxRecords   'as Integer
	Dim intDivSize      'as Integer

	strTitle = getIDS("IDS_ModName4")

	intMaxRecords = 20
	intDivSize = CInt((intScreenH-100)/2)

	Call DisplayHeader(1)
%>

<div id="contentDiv" class="dvBorder">

	<div style="float:left;width:75%;">
<%
	Call showList(12,intDivSize,intMaxRecords,3)
	Call showList(13,intDivSize,intMaxRecords,3)
%>
	</div>

	<div style="margin-left:75%;padding-left:10px;">
	  <div id="div13" class="dvRightMenu" style="height:<% =CInt((intScreenH-70)*0.3) %>px;">
		<div class="hRow hScr hFont"><% =getIDS("IDS_UsefulLinks") %></div>
		<div>
		  <ul>
			<% If pProjects > 2 Then %>
			<li><% =getEditLink(4,"",getIDS("IDS_ProjectNew")) %></li>
			<div class="hr"></div>
			<% End If %>
			<li><% =getPopLink("S","?m=4",getIDS("IDS_ProjectSearch")) %></li>
			<div class="hr"></div>
			<li><a href="../common/calendar.asp?m=4"><% =getIDS("IDS_Calendar") %></a></li>
		  </ul>
		</div>
	  </div>

<%	  Call showList(31,CInt((intScreenH-70)*0.6),"4",1) %>

	</div>
</div>

<%
	Call DisplayFooter(1)
%>