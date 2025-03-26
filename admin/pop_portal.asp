<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\functions_list.asp" -->
<%
	Call pageFunctions(0,1)

	Dim arrDisplay		'as Array
	Dim bytVal		'as Byte
	Dim intListCount	'as Integer
	Dim strList		'as String
	Dim blnShow		'as Boolean

	arrDisplay = split("chkMyContacts,chkMyClients,chkMySales,chkMyProjects,chkMyEvents,chkMyTickets,chkMyBugs,chkMyInvoices,chkMyReports,chkContacts,chkClients,chkSales,chkProjects,chkEvents,chkHotTickets,chkTickets,chkHotBugs,chkBugs,chkDueInvoices,chkInvoices,chkReports,chkNewArticles",",")

	strTitle = Application("IDS_EditHomePage")

	If strDoAction = "edit" Then

		intListCount = 0
		strList = ""

		For i = 0 to UBound(arrDisplay)
			If valNum(Request.Form(arrDisplay(i)),0,0) = 0 Then
				strList = strList & "0"
			Else
				intListCount = intListCount + 1
				strList = strList & "1"
			End If
		Next

		Session("PortalCount") = intListCount
		Session("PortalList") = strList
		objConn.Execute(insertPortalView(lngUserId,strList))
		Call closeWindow(strOpenerURL)
	End If

	Call DisplayHeader(3)
	Call showEditHeader(strTitle,"","","","")
%>
<div id="contentDiv" class="dvBorder" style="height:430px;"><br>

<table border=0 cellspacing=5 width="100%">
<form name="frmAdmin" method="post" action="pop_portal.asp">
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
<% =getHidden("hdnWinOpen","") %>
  <tr class="hrow">
    <td class="dlabel" colspan=2><% =strFullName %></td>
  </tr>
<%
	For i = 0 to UBound(arrDisplay)

		blnShow = False
		bytVal = Mid(Session("PortalList"),i+1,1)
		If bytVal <> "" Then bytVal = CByte(bytVal)

		Select Case i
			Case 0
				If mContacts Then blnShow = True
			Case 9
				If pContacts >= 1 Then blnShow = True
			Case 1
				If mClients Then blnShow = True
			Case 10
				If pClients >= 1 Then blnShow = True
			Case 2
				If mSales Then blnShow = True
			Case 11
				If pSales >= 1 Then blnShow = True
			Case 3
				If mProjects Then blnShow = True
			Case 12,13
				If pProjects >= 1 Then blnShow = True
			Case 5
				If mTickets Then blnShow = True
			Case 14,15
				If pTickets >= 1 Then blnShow = True
			Case 6
				If mBugs Then blnShow = True
			Case 16,17
				If pBugs >= 1 Then blnShow = True
			Case 7
				If mInvoices Then blnShow = True
			Case 18,19
				If pInvoices >= 1 Then blnShow = True
			Case 4,8,20,21
				blnShow = True
		End Select

		If i = 9 or i = 21 Then
%>
  <tr class="hrow">
    <td class="dfont" colspan=2>&nbsp;&nbsp;</td>
  </tr>
<%
		End If

		If blnShow Then
%>
  <tr>
    <td><% =getLabel(getListName(i),arrDisplay(i)) %></td>
    <td align=center><% =getCheckbox(arrDisplay(i),bytVal,"") %></td>
  </tr>
<%
		End If
	Next
%>
</form>
</table>

</div>

<div id="footerDiv" class="dvFooter">
<%
	Response.Write(getIconSave("edit"))
	Response.Write(getIconCancel())
%>
</div>
<%
	Call DisplayFooter(3)
%>

