<!--#include file="..\_inc\functions.asp" -->
<!--#include file="..\_inc\functions_list.asp" -->
<%
	Call pageFunctions(0,1)

	Dim strClass        'as String
	Dim intDivSize      'as Integer
	Dim intMaxRecords   'as Integer

	strTitle = getIDS("IDS_Reports")
	strDir = Application("av_CRMDir") & "reports/"
	strModItem = getIDS("IDS_Reports")
	strModName = getIDS("IDS_Reports")

	intMaxRecords = 20
	intDivSize = CInt((intScreenH-100)/2)

	Call DisplayHeader(1)
%>

<div id="contentDiv" class="dvBorder">

	<div style="float:left;width:75%;">
<%
	Call showList(8,intDivSize,intMaxRecords,2)
	Call showList(20,intDivSize,intMaxRecords,2)
%>
	</div>

	<div style="margin-left:75%;padding-left:10px;">
	  <div id="div13" class="dvRightMenu" style="height:<% =CInt((intScreenH-70)*0.3) %>px;">
		<div class="hRow hScr hFont"><% =getIDS("IDS_UsefulLinks") %></div>
		<div>
			<ul>
			  <li><a href="edit_report.asp"><% =getIDS("IDS_ReportNew") %></a></li>
			</ul>
			<br /><br />
		</div>
	  </div>
	</div>
</div>

<%
	Call DisplayFooter(1)
%>