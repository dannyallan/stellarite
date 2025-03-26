<!--#include file="..\_inc\functions.asp" -->
<!--#include file="..\_inc\functions_list.asp" -->
<%
	Call pageFunctions(0,1)

	Dim strClass            'as String
	Dim intCount            'as Integer
	Dim intDivSize          'as Integer
	Dim intMaxRecords       'as Integer

	If not pContacts >= 1 and not pClients >= 1 and not pSales >= 1 Then Call logError(2,1)

	strModName = getIDS("IDS_ModName123")
	strTitle = getIDS("IDS_ModName123")

	If pContacts >= 1 Then intCount = intCount + 1
	If pClients >= 1 Then intCount = intCount + 1
	If pSales >= 1 Then intCount = intCount + 1

	intMaxRecords = 20
	intDivSize = CInt((intScreenH-50-(intCount*20))/intCount)

	Call DisplayHeader(1)
%>

<div id="contentDiv" class="dvBorder">

	<div style="float:left;width:75%;">
<%
	If pClients >= 1 Then Call showList(10,intDivSize,intMaxRecords,3)
	If pContacts >= 1 Then Call showList(9,intDivSize,intMaxRecords,3)
	If pSales >= 1 Then Call showList(11,intDivSize,intMaxRecords,3)
%>
	</div>

	<div style="margin-left:75%;padding-left:10px;">
	  <div id="div13" class="dvRightMenu" style="height:<% =CInt((intScreenH-70)*0.3) %>px;">
		<div class="hRow hScr hFont"><% =getIDS("IDS_UsefulLinks") %></div>
		<div>
		  <ul>
<%
	If pContacts >= 2 Then Response.Write("<li>" & getEditLink(1,"",getIDS("IDS_ContactNew")) & "</li>" & vbCrLf)
	If pClients >= 2 Then Response.Write("<li>" & getEditLink(2,"",getIDS("IDS_AccountNew")) & "</li>" & vbCrLf)
	If pSales >= 2 Then Response.Write("<li>" & getEditLink(3,"",getIDS("IDS_SaleNew")) & "</li>" & vbCrLf)
	Response.Write("<div class=""hr""></div>" & vbCrLf)
	If pContacts >= 1 Then Response.Write("<li>" & getPopLink("S","?m=1",getIDS("IDS_ContactSearch")) & "</li>" & vbCrLf)
	If pClients >= 1 Then Response.Write("<li>" & getPopLink("S","?m=2",getIDS("IDS_AccountSearch")) & "</li>" & vbCrLf)
	If pSales >= 1 Then Response.Write("<li>" & getPopLink("S","?m=3",getIDS("IDS_SaleSearch")) & "</li>" & vbCrLf)
%>
			<div class="hr"></div>
			<li><a href="../common/calendar.asp?m=2"><% =getIDS("IDS_Calendar") %></a></li>
		  </ul>
		</div>
	  </div>

<%	  Call showList(31,CInt((intScreenH-70)*0.6),"1,2,3",1) %>

	</div>
</div>

<%
	Call DisplayFooter(1)
%>
