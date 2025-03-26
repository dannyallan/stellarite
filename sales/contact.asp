<!--#include file="..\_inc\functions.asp" -->
<!--#include file="..\_inc\functions_mod.asp" -->
<!--#include file="..\_inc\sql\sql_contacts.asp" -->
<%
	Call pageFunctions(1,1)

	Dim strFrame        'as String
	Dim strClient       'as String
	Dim strAddress1     'as String
	Dim strAddress2     'as String
	Dim strAddress3     'as String
	Dim strCity         'as String
	Dim strState        'as String
	Dim strCountry      'as String
	Dim strZIP          'as String
	Dim lngDivId        'as Long
	Dim strDivision     'as String
	Dim strDept         'as String
	Dim strJobTitle     'as String
	Dim strEmail        'as String
	Dim strPhone1       'as String
	Dim strExt1         'as String
	Dim strPhone2       'as String
	Dim strExt2         'as String
	Dim strFax          'as String

	lngRecordId = valNum(lngRecordId,3,1)

	If strDoAction = "del" and intPerm >= 4 Then
		Call delContact(lngUserId,lngRecordId)
		lngPrevId = doPrevNext(0,1,lngRecordId,0,0)
		lngNextId = doPrevNext(1,1,lngRecordId,0,0)
		strTitle = getIDS("IDS_Deleted")
	Else
		Set objRS = objConn.Execute(getContact(1,lngRecordId))

		If not (objRS.BOF and objRS.EOF) then
			lngPrevId = objRS.fields("PrevId").value
			lngNextId = objRS.fields("NextId").value
			strTitle = getAOS(objRS.fields("K_Prefix").value) & " " & objRS.fields("K_FirstName").value & " " & objRS.fields("K_LastName").value
			strAddress1 = objRS.fields("K_Address1").value
			strAddress2 = objRS.fields("K_Address2").value
			strAddress3 = objRS.fields("K_Address3").value
			strCity = objRS.fields("K_City").value
			strState = objRS.fields("K_State").value
			strCountry = objRS.fields("K_Country").value
			strZIP = objRS.fields("K_ZIP").value
			strClient = objRS.fields("C_Client").value
			lngDivId = objRS.fields("DivId").value
			strDivision = objRS.fields("D_Division").value
			strDept = objRS.fields("K_Dept").value
			strJobTitle = objRS.fields("K_JobTitle").value
			strEmail = objRS.fields("K_Email").value
			strPhone1 = objRS.fields("K_Phone1").value
			strExt1 = objRS.fields("K_Ext1").value
			strPhone2 = objRS.fields("K_Phone2").value
			strExt2 = objRS.fields("K_Ext2").value
			strFax = objRS.fields("K_Fax").value
		Else
			strTitle = getIDS("IDS_Deleted")
		End If
	End If

	Call DisplayHeader(1)
%>

<div id="modDiv" class="dvMod">

<%    Call showToolBar() %>

<table border="0" cellspacing="10" width="100%">
  <tr><td width="33%" valign=top>

  <table border=0>
	<tr>
	  <td nowrap class="bFont"><% =getIDS("IDS_Account") %></td>
	  <td class="dFont"><% =showLink(2,"client.asp?id=" & lngDivId,strClient) %>&nbsp;</td>
	</tr>
	<tr>
	  <td nowrap class="bFont"><% =getIDS("IDS_Division") %></td>
	  <td class="dFont"><% =showString(strDivision) %>&nbsp;</td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Department") %></td>
	  <td class="dFont"><% =showString(strDept) %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_JobTitle") %></td>
	  <td class="dFont"><% =showString(strJobTitle) %></td>
	</tr>
  </table>

  </td>
  <td width="33%" valign=top>

  <table border=0>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Email") %></td>
	  <td class="dFont"><% =showEmail(strEmail) %></td>
	</tr>
	<tr>
	  <td class="bFont""><% =getIDS("IDS_Phone") %> 1</td>
	  <td class="dFont"><% =showPhone(strPhone1) %><% If strExt1 <> "" Then Response.Write("&nbsp;&nbsp;<span class=""bFont"">Ext.</span> " & strExt1) %></td>
	</tr>
	<tr>
	  <td class="bFont""><% =getIDS("IDS_Phone") %> 2</td>
	  <td class="dFont"><% =showPhone(strPhone2) %><% If strExt2 <> "" Then Response.Write("&nbsp;&nbsp;<span class=""bFont"">Ext.</span> " & strExt2) %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Fax") %></td>
	  <td class="dFont"><% =showPhone(strFax) %></td>
	</tr>
  </table>

  </td>
  <td width="33%" valign=top>

  <table border=0>
	<tr>
	  <td valign=top class="bFont"><% =getIDS("IDS_Address") %></td>
	  <td valign=top class="dFont" width=250>
<%
	If strAddress1 <> "" Then Response.Write(vbTab & showString(strAddress1) & "<br />" & vbCrLf)
	If strAddress2 <> "" Then Response.Write(vbTab & showString(strAddress2) & "<br />" & vbCrLf)
	If strAddress3 <> "" Then Response.Write(vbTab & showString(strAddress3) & "<br />" & vbCrLf)
	If strCity <> "" Then Response.Write(vbTab & showString(strCity) & "<br />" & vbCrLf)
	Response.Write(showString(strState))
	If strState <> "" and strCountry <> "" Then Response.Write(", ")
	If strState <> "" or strCountry <> "" Then Response.Write(showString(strCountry) & "&nbsp;&nbsp;&nbsp;")
	Response.Write(showString(strZIP) & vbCrLf)
%>
	  &nbsp;</td>

	</tr>
  </table>

  </td></tr>
</table>

<%
	If strTitle <> getIDS("IDS_Deleted") Then

		strTabBuilder = getIDS("IDS_Summary") & "|i_summary.asp?m=1&mid=" & lngRecordId & _
			"||" & getIDS("IDS_Notes") & "|../common/i_notes.asp?m=1&mid=" & lngRecordId & _
			"||" & getIDS("IDS_Attachments") & "|../common/i_attach.asp?m=1&mid=" & lngRecordId

		'Enable following line allows event logging
		'strTabBuilder = strTabBuilder & "||" & getIDS("IDS_Events") & "|../common/i_events.asp?m=1&mid=" & lngRecordId

		If pSales >= 1 Then strTabBuilder = strTabBuilder & "||" & getIDS("IDS_Sales") & "|../common/i_sales.asp?m=1&mid=" & lngRecordId
		strTabBuilder = strTabBuilder & "||" & getIDS("IDS_Products") & "|../common/i_products.asp?m=1&mid=" & lngRecordId
		If pProjects >= 1 Then strTabBuilder = strTabBuilder & "||" & getIDS("IDS_Projects") & "|../common/i_projects.asp?m=2&mid=" & lngDivId
		If pTickets >= 1 Then strTabBuilder = strTabBuilder & "||" & getIDS("IDS_Tickets") & "|../common/i_tickets.asp?m=1&mid=" & lngRecordId
		If pInvoices >= 1 Then strTabBuilder = strTabBuilder & "||" & getIDS("IDS_Invoices") & "|../common/i_invoices.asp?m=1&mid=" & lngRecordId

		strFrame = makeTabs(strTabBuilder)

		Response.Write("</div><iframe id=""contentDiv"" class=""iBorder"" src=""" & strTabURL & """ title=""" & strFrame & """ style=""height:" & intScreenH-210 & "px;"" width=""100%"" scrolling=""no""></iframe>" & vbCrLf)
	End If

	Call DisplayFooter(1)
%>