<!--#include file="..\_inc\functions.asp" -->
<!--#include file="..\_inc\functions_mod.asp" -->
<!--#include file="..\_inc\sql\sql_clients.asp" -->
<%
	Call pageFunctions(2,1)

	Dim strFrame        'as String
	Dim strDivision     'as String
	Dim strAccount      'as String
	Dim strAccountType  'as String
	Dim strRefAccount   'as String
	Dim strSalesRep     'as String
	Dim strRegion       'as String
	Dim strWebsite      'as String
	Dim strProbFlag     'as String

	lngRecordId = valNum(lngRecordId,3,1)

	If strDoAction = "del" and intPerm >= 4 Then
		Call delClient(lngUserId,lngRecordId)
		lngPrevId = doPrevNext(0,2,lngRecordId,0,0)
		lngNextId = doPrevNext(1,2,lngRecordId,0,0)
		strTitle = getIDS("IDS_Deleted")
	Else
		Set objRS = objConn.Execute(getClient(1,lngRecordId))

		If not (objRS.BOF and objRS.EOF) then
			lngPrevId = objRS.fields("PrevId").value
			lngNextId = objRS.fields("NextId").value
			strTitle = showString(objRS.fields("C_Client").value)
			strDivision = showString(objRS.fields("D_Division").value)
			strSalesRep = showString(objRS.fields("SalesRep").value)
			strRegion = getAOS(objRS.fields("D_Region").value)
			strWebsite = showString(objRS.fields("D_Website").value)
			strAccount = showString(objRS.fields("D_Account").value)
			strAccountType = getAOS(objRS.fields("D_AccountType").value)
			strAccount = showString(objRS.fields("D_Account").value)
			strRefAccount = showTrueFalse(objRS.fields("D_RefAccount").value)
			strProbFlag = showTrueFalse(objRS.fields("D_ProbFlag").value)
		Else
			strTitle = getIDS("IDS_Deleted")
		End If
	End If

	Call DisplayHeader(1)
%>

<div id="modDiv" class="dvMod">

<% Call showToolBar() %>

<table border="0" cellspacing="10" width="100%">
  <tr><td width="50%" valign=top>

  <table border=0>
	<tr>
	 <td class="bFont"><% =getIDS("IDS_Division") %></td>
	  <td class="dFont"><% =strDivision %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_SalesRep") %></td>
	  <td class="dFont"><% =strSalesRep %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_SalesRegion") %></td>
	  <td class="dFont"><% =strRegion %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Website") %></td>
	  <td class="dFont"><% =showLink(0,strWebsite,strWebsite) %></td>
	</tr>
  </table>

  </td>
  <td width="50%" valign=top>

  <table border=0>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_AccountType") %></td>
	  <td class="dFont"><% =strAccountType %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_AccountId") %></td>
	  <td class="dFont"><% =strAccount %></td>
	</tr>

	<tr>
	  <td class="bFont"><% =getIDS("IDS_Reference") %></td>
	  <td class="dFont"><% =strRefAccount %></td>
	</tr>
	<tr>
	  <td class="bFont""><% =getIDS("IDS_ProblemFlag") %></td>
	  <td class="dFont" width=100><% =strProbFlag %>&nbsp;</td>
	</tr>
  </table>

  </td></tr>
</table>


<%
	If strTitle <> getIDS("IDS_Deleted") Then

		strTabBuilder = getIDS("IDS_Summary") & "|i_summary.asp?m=2&mid=" & lngRecordId & _
			"||" & getIDS("IDS_Notes") & "|../common/i_notes.asp?m=2&mid=" & lngRecordId & _
			"||" & getIDS("IDS_Attachments") & "|../common/i_attach.asp?m=2&mid=" & lngRecordId

		'Enable following line allows event logging
		strTabBuilder = strTabBuilder & "||" & getIDS("IDS_Events") & "|../common/i_events.asp?m=2&mid=" & lngRecordId

		If pContacts >= 1 Then strTabBuilder = strTabBuilder & "||" & getIDS("IDS_Contacts") & "|i_contacts.asp?m=2&mid=" & lngRecordId
		If pSales >= 1 Then strTabBuilder = strTabBuilder & "||" & getIDS("IDS_Sales") & "|../common/i_sales.asp?m=2&mid=" & lngRecordId
		strTabBuilder = strTabBuilder & "||" & getIDS("IDS_Products") & "|../common/i_products.asp?m=2&mid=" & lngRecordId
		If pProjects >= 1 Then strTabBuilder = strTabBuilder & "||" & getIDS("IDS_Projects") & "|../common/i_projects.asp?m=2&mid=" & lngRecordId
		If pTickets >= 1 Then strTabBuilder = strTabBuilder & "||" & getIDS("IDS_Tickets") & "|../common/i_tickets.asp?m=2&mid=" & lngRecordId
		If pInvoices >= 1 Then strTabBuilder = strTabBuilder & "||" & getIDS("IDS_Invoices") & "|../common/i_invoices.asp?m=2&mid=" & lngRecordId

		strFrame = makeTabs(strTabBuilder)

		Response.Write("</div><iframe id=""contentDiv"" class=""iBorder"" src=""" & strTabURL & """ title=""" & strFrame & """ style=""height:" & intScreenH-210 & "px;"" width=""100%"" scrolling=""no""></iframe>" & vbCrLf)
	End If

	Call DisplayFooter(1)
%>