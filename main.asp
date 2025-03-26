<!--#include file="_inc\functions.asp" -->
<!--#include file="_inc\functions_list.asp" -->
<%
	Call pageFunctions(0,1)

	Dim intCount        'as Integer
	Dim strClass        'as String
	Dim strClassUrgent  'as String
	Dim strNews         'as String
	Dim blnColWrap      'as Boolean
	Dim intMaxRecords   'as Integer
	Dim intDivSize      'as Integer

	strTitle = "Stellarite " & getIDS("IDS_Home")
	strNews = Application("av_ModuleMsg0")
	blnColWrap = False
	intMaxRecords = 20

	If mContacts Then strNews = strNews & "<hr width=""80%"" size=1 noshade />" & vbCrLf & Application("av_ModuleMsg1")
	If mClients Then strNews = strNews & "<hr width=""80%"" size=1 noshade />" & vbCrLf & Application("av_ModuleMsg2")
	If mSales Then strNews = strNews & "<hr width=""80%"" size=1 noshade />" & vbCrLf & Application("av_ModuleMsg3")
	If mProjects Then strNews = strNews & "<hr width=""80%"" size=1 noshade />" & vbCrLf & Application("av_ModuleMsg4")
	If mTickets Then strNews = strNews & "<hr width=""80%"" size=1 noshade />" & vbCrLf & Application("av_ModuleMsg5")
	If mBugs Then strNews = strNews & "<hr width=""80%"" size=1 noshade />" & vbCrLf & Application("av_ModuleMsg6")
	If mInvoices Then strNews = strNews & "<hr width=""80%"" size=1 noshade />" & vbCrLf & Application("av_ModuleMsg7")

	If Session("PortalCount") <> 0 Then
		intCount = Session("PortalCount")
		If intCount > 3 Then intCount = Fix((Session("PortalCount")+1)/2)
		intMaxRecords = 20
		intDivSize = CInt((intScreenH-50-(intCount*20))/intCount)
	End If

	intCount = 0

	Call DisplayHeader(1)
%>

<div id="contentDiv" class="dvBorder">

  <div style="float:left;width:75%;">
    <div style="float:left;width:<% If Session("PortalCount") > 3 Then Response.Write("50") Else Response.Write("100") %>%;">
<%
		For i = 0 to Len(Session("PortalList"))
			If Mid(Session("PortalList"),i+1,1) = "1" Then
				Call showList(i,intDivSize,intMaxRecords,2)
				intCount = intCount + 1
			End If
			If intCount = Fix((Session("PortalCount")+1)/2) and Session("PortalCount") > 3 and NOT blnColWrap Then
				blnColWrap = True
				Response.Write("    </div>" & vbCrLf & "    <div style=""margin-left:50%;padding-left:10px;"">" & vbCrLf)
			End If
		Next
%>
    </div>
  </div>

  <div style="margin-left:75%;padding-left:10px;">

	  <div id="div13" class="dvRightMenu" style="height:<% =CInt((intScreenH-70)*0.3) %>px;">
	      <div class="hRow hScr hFont"><% =getIDS("IDS_NewsUpdates") %></div>
		  <div><p class="dFont"><% =strNews %></p><br /></div>
	  </div>

	  <div id="div23" class="dvRightMenu" style="height:<% =CInt((intScreenH-70)*0.6) %>px;">
	  	<div class="hRow hScr hFont"><% =getIDS("IDS_UsefulLinks") %></div>
	    <div>
	        <ul>
<%
	If pContacts >= 1 or pClients >= 1 or pSales >= 1 Then Response.Write("<li><a href=""sales/default.asp"">" & getIDS("IDS_ModName123") & "</a></li>" & vbCrLf)
	If pProjects >= 1 Then Response.Write("<li><a href=""services/default.asp"">" & getIDS("IDS_ModName4") & "</a></li>" & vbCrLf)
	If pTickets >= 1 Then Response.Write("<li><a href=""support/default.asp"">" & getIDS("IDS_ModName5") & "</a></li>" & vbCrLf)
	If pBugs >= 1 Then Response.Write("<li><a href=""qa/default.asp"">" & getIDS("IDS_ModName6") & "</a></li>" & vbCrLf)
	If pInvoices >= 1 Then Response.Write("<li><a href=""finance/default.asp"">" & getIDS("IDS_ModName7") & "</a></li>" & vbCrLf)
	If pArticles >= 1 Then Response.Write("<li><a href=""kb/default.asp"">" & getIDS("IDS_ModName8") & "</a></li>" & vbCrLf)
	Response.Write("<li><a href=""reports/default.asp"">" & getIDS("IDS_Reports") & "</a></li>" & vbCrLf)
	Response.Write("<div class=""hr""></div>" & vbCrLf)

	If mContacts and pContacts >= 2 Then Response.Write("<li>" & getEditLink(1,"",getIDS("IDS_ContactNew")) & "</li>" & vbCrLf)
	If mClients and pClients >= 2 Then Response.Write("<li>" & getEditLink(2,"",getIDS("IDS_AccountNew")) & "</li>" & vbCrLf)
	If mSales and pSales >= 2 Then Response.Write("<li>" & getEditLink(3,"",getIDS("IDS_SaleNew")) & "</li>" & vbCrLf)
	If mProjects and pProjects >= 2 Then Response.Write("<li>" & getEditLink(4,"",getIDS("IDS_ProjectNew")) & "</li>" & vbCrLf)
	'Response.Write("<li>" & getEditLink(50,"?m=0&mid=0",getIDS("IDS_EventNew")) & "</li>" & vbCrLf)
	If mTickets and pTickets >= 2 Then Response.Write("<li>" & getEditLink(5,"",getIDS("IDS_TicketNew")) & "</li>" & vbCrLf)
	If mBugs and pBugs >= 2 Then Response.Write("<li>" & getEditLink(6,"",getIDS("IDS_BugNew")) & "</li>" & vbCrLf)
	If mInvoices and pInvoices >= 2 Then Response.Write("<li>" & getEditLink(7,"",getIDS("IDS_InvoiceNew")) & "</li>" & vbCrLf)
	If mArticles and pArticles >= 2 Then Response.Write("<li>" & getEditLink(8,"",getIDS("IDS_ArticleNew")) & "</li>" & vbCrLf)
	Response.Write("<div class=""hr""></div>" & vbCrLf)

	If mContacts Then Response.Write("<li>" & getPopLink("S","?m=1",getIDS("IDS_ContactSearch")) & "</li>" & vbCrLf)
	If mClients Then Response.Write("<li>" & getPopLink("S","?m=2",getIDS("IDS_AccountSearch")) & "</li>" & vbCrLf)
	If mSales Then Response.Write("<li>" & getPopLink("S","?m=3",getIDS("IDS_SaleSearch")) & "</li>" & vbCrLf)
	If mProjects Then Response.Write("<li>" & getPopLink("S","?m=4",getIDS("IDS_ProjectSearch")) & "</li>" & vbCrLf)
	If mTickets Then Response.Write("<li>" & getPopLink("S","?m=5",getIDS("IDS_TicketSearch")) & "</li>" & vbCrLf)
	If mBugs Then Response.Write("<li>" & getPopLink("S","?m=6",getIDS("IDS_BugSearch")) & "</li>" & vbCrLf)
	If mInvoices Then Response.Write("<li>" & getPopLink("S","?m=7",getIDS("IDS_InvoiceSearch")) & "</li>" & vbCrLf)
	If mArticles Then Response.Write("<li><a href=""kb/search.asp"">" & getIDS("IDS_ArticleSearch") & "</a></li>" & vbCrLf)
	Response.Write("<div class=""hr""></div>" & vbCrLf)

	Response.Write("<li><a href=""common/calendar.asp"">" & getIDS("IDS_Calendar") & "</a></li>" & vbCrLf)
	If Application("av_EnableEmail") <> "0" Then Response.Write("<li>" & getEditLink("B","",getIDS("IDS_EmailSubscriptions")) & "</li>" & vbCrLf)
	Response.Write("<li>" & getEditLink("O","",getIDS("IDS_EditHomePage")) & "</li>" & vbCrLf)
	If blnAdmin or getSecurity(0,5) Then Response.Write("<li><a href=""admin/default.asp"">" & getIDS("IDS_Administration") & "</a></li>" & vbCrLf)

%>
	        </ul>
		</div>
	  </div>
	</div>
</div>

<%
	Call DisplayFooter(1)
%>
