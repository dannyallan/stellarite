<% Sub DisplayMenu() %>

<div id="SkipNav" style="display:none;">
  <a href="#content"><img src="<% =Application("av_CRMDir") %>images/spacer.gif" alt="<% =getIDS("IDS_SkipNav") %>" border=0 height=0 width=0 /></a>
</div>

<div id="navDiv">
<div class="menuBar">
  <div style="float:left;">
	<a class="menuButton" href="Javascript:resetButton(activeButton);" onClick="return buttonClick(event,'fileMenu');" onMouseOver="buttonMouseover(event,'fileMenu');"><% =getIDS("IDS_File") %></a>
	<a class="menuButton" href="Javascript:resetButton(activeButton);" onClick="return buttonClick(event,'searchMenu');" onMouseOver="buttonMouseover(event,'searchMenu');"><% =getIDS("IDS_Search") %></a>
	<a class="menuButton" href="Javascript:resetButton(activeButton);" onClick="return buttonClick(event,'modMenu');" onMouseOver="buttonMouseover(event,'modMenu');"><% =getIDS("IDS_Modules") %></a>
	<a class="menuButton" href="Javascript:resetButton(activeButton);" onClick="return buttonClick(event,'helpMenu');" onMouseOver="buttonMouseover(event,'helpMenu');"><% =getIDS("IDS_Help") %></a>
  </div>
  <div class="userName"><% =getIDS("IDS_User") & ": " & showString(Session("UserName")) %></div>
</div>

<div id="fileMenu" class="menu">
  <a class="menuItem" href="<% =Application("av_CRMDir") %>main.asp"><% =getIDS("IDS_Home") %></a>
  <div class="menuItemSep"></div>
  <a class="menuItem" href="<% =Application("av_CRMDir") %>admin/profile.asp?id=<% =lngUserId %>&o=1&so=ASC"><% =getIDS("IDS_MyProfile") %></a>
  <% If Application("av_HideUsers") <> "1" or (blnAdmin or getSecurity(0,5)) Then %>
  <a class="menuItem" href="<% =Application("av_CRMDir") %>admin/profile.asp"><% =getIDS("IDS_UserProfiles") %></a>
  <% End If %>
  <div class="menuItemSep"></div>
  <a class="menuItem" href="<% =getEditURL("W","?id="&lngUserId) %>"><% =getIDS("IDS_PasswordEdit") %></a>
  <a class="menuItem" href="<% =Application("av_CRMDir") %>admin/edit_portal.asp"><% =getIDS("IDS_EditHomePage") %></a>
  <a class="menuItem" href="<% =Application("av_CRMDir") %>common/edit_subscriptions.asp"><% =getIDS("IDS_EmailSubscriptions") %></a>
  <div class="menuItemSep"></div>
  <a class="menuItem" href="<% =Application("av_CRMDir") %>common/calendar.asp"><% =getIDS("IDS_Calendar") %></a>
  <div class="menuItemSep"></div>
  <% If blnAdmin or getSecurity(0,5) Then %>
  <a class="menuItem" href="<% =Application("av_CRMDir") %>admin/default.asp"><% =getIDS("IDS_Administration") %></a>
 <% End If %>
  <div class="menuItemSep"></div>
  <a class="menuItem" href="<% =Application("av_CRMDir") %>default.asp?logout" onClick="closeWindow('<% =Application("av_CRMDir") %>default.asp?logout');return false;"><% =getIDS("IDS_Logout") %></a>
</div>

<div id="searchMenu" class="menu">
<%
If pContacts >= 1 Then Response.Write("  <a class=""menuItem"" href=""" & getEditURL("S","?m=1") & """ onClick=""resetButton(activeButton);" & getSearchURL("?m=1") & "return false;"">" & getIDS("IDS_Contacts") & "</a>" & vbCrLf)
If pClients >= 1 Then Response.Write("  <a class=""menuItem"" href=""" & getEditURL("S","?m=2") & """ onClick=""resetButton(activeButton);" & getSearchURL("?m=2") & "return false;"">" & getIDS("IDS_Accounts") & "</a>" & vbCrLf)
If pSales >= 1 Then Response.Write("  <a class=""menuItem"" href=""" & getEditURL("S","?m=3") & """ onClick=""resetButton(activeButton);" & getSearchURL("?m=3") & "return false;"">" & getIDS("IDS_Sales") & "</a>" & vbCrLf)
If pProjects >= 1 Then Response.Write("  <a class=""menuItem"" href=""" & getEditURL("S","?m=4") & """ onClick=""resetButton(activeButton);" & getSearchURL("?m=4") & "return false;"">" & getIDS("IDS_Projects") & "</a>" & vbCrLf)
If pTickets >= 1 Then Response.Write("  <a class=""menuItem"" href=""" & getEditURL("S","?m=5") & """ onClick=""resetButton(activeButton);" & getSearchURL("?m=5") & "return false;"">" & getIDS("IDS_Tickets") & "</a>" & vbCrLf)
If pBugs >= 1 Then Response.Write("  <a class=""menuItem"" href=""" & getEditURL("S","?m=6") & """ onClick=""resetButton(activeButton);" & getSearchURL("?m=6") & "return false;"">" & getIDS("IDS_Bugs") & "</a>" & vbCrLf)
If pInvoices >= 1 Then Response.Write("  <a class=""menuItem"" href=""" & getEditURL("S","?m=7") & """ onClick=""resetButton(activeButton);" & getSearchURL("?m=7") & "return false;"">" & getIDS("IDS_Invoices") & "</a>" & vbCrLf)
Response.Write("  <div class=""menuItemSep""></div>" & vbCrLf)
If pArticles >= 1 Then Response.Write("  <a class=""menuItem"" href=""" & Application("av_CRMDir") & "kb/search.asp"">" & getIDS("IDS_ModName8") & "</a>" & vbCrLf)
%>
</div>

<div id="modMenu" class="menu">
<%
If pContacts >= 1 or pClients >= 1 or pSales >= 1 Then Response.Write("  <a class=""menuItem"" href=""" & Application("av_CRMDir") & "sales/default.asp"">" & getIDS("IDS_ModName123") & "</a>" & vbCrLf)
If pProjects >= 1 Then Response.Write("  <a class=""menuItem"" href=""" & Application("av_CRMDir") & "services/default.asp"">" & getIDS("IDS_ModName4") & "</a>" & vbCrLf)
If pTickets >= 1 Then Response.Write("  <a class=""menuItem"" href=""" & Application("av_CRMDir") & "support/default.asp"">" & getIDS("IDS_ModName5") & "</a>" & vbCrLf)
If pBugs >= 1 Then Response.Write("  <a class=""menuItem"" href=""" & Application("av_CRMDir") & "qa/default.asp"">" & getIDS("IDS_ModName6") & "</a>" & vbCrLf)
If pInvoices >= 1 Then Response.Write("  <a class=""menuItem"" href=""" & Application("av_CRMDir") & "finance/default.asp"">" & getIDS("IDS_ModName7") & "</a>" & vbCrLf)
Response.Write("  <div class=""menuItemSep""></div>" & vbCrLf)
If pArticles >= 1 Then Response.Write("  <a class=""menuItem"" href=""" & Application("av_CRMDir") & "kb/default.asp"">" & getIDS("IDS_ModName8") & "</a>" & vbCrLf)
%>
  <div class="menuItemSep"></div>
  <a class="menuItem" href="<% =Application("av_CRMDir") %>reports/default.asp"><% =getIDS("IDS_Reports") %></a>
</div>

<div id="helpMenu" class="menu">
  <a class="menuItem" href="<% =Application("av_CRMDir") & "faq.asp"">" & getIDS("IDS_FAQs") %></a>
  <a class="menuItem" href="<% =Application("av_CRMDir") %>require.asp" onClick="resetButton(activeButton);window.open('<% =Application("av_CRMDir") %>require.asp','require');return false;"><% =getIDS("IDS_Requirements") %></a>
  <a class="menuItem" href="<% =Application("av_CRMDir") %>about.asp" onClick="resetButton(activeButton);openWindow('<% =Application("av_CRMDir") %>about.asp','sw_About','350','250');return false;"><% =getIDS("IDS_About") %></a>
</div>

<div class="breadcrumb">
  <span class="breadcrumbtext">
<%
	If strTitle <> "Stellarite " & getIDS("IDS_Home") Then
		Response.Write("    <a class=""breadcrumbtext"" href=""" & Application("av_CRMDir") & "main.asp"">" & getIDS("IDS_Home") & "</a>")
	Else
		Response.Write(getIDS("IDS_Home"))
	End If

	If strModName <> "" and strModName <> strTitle Then
		Response.Write(" &raquo; <a class=""breadcrumbtext"" href=""" & strDir & "default.asp"">" & showString(strModName) & "</a>")
	End If

	If strModName <> "" Then
		Response.Write(" &raquo; ")
		If bytMod <> 90 and strModItem <> "" and strModName <> strTitle and bytRealMod <> 0 Then Response.Write(strModItem & ": ")
		Response.Write(showString(strTitle))
	End If
%>
  </span>
</div>
</div>

<a name="content"></a>

<% End Sub %>