<!--#include file="..\_inc\functions.asp" -->
<!--#include file="..\_inc\functions_mod.asp" -->
<!--#include file="..\_inc\sql\sql_services.asp" -->
<%
	Call pageFunctions(4,1)

	Dim strClient       'as String
	Dim strDivision     'as String
	Dim lngInvoiceId    'as Long
	Dim strOwner        'as String
	Dim strFrame        'as String
	Dim intDaysTotal    'as Integer
	Dim intDaysOwed     'as Integer
	Dim lngDivId        'as Long
	Dim strClosed       'as String

	lngRecordId = valNum(lngRecordId,3,1)

	If strDoAction = "del" and intPerm >= 4 then
		Call delProject(lngUserId,lngRecordId)
		lngPrevId = doPrevNext(0,4,lngRecordId,0,0)
		lngNextId = doPrevNext(1,4,lngRecordId,0,0)
		strTitle = getIDS("IDS_Deleted")
	Else
		Set objRS = objConn.Execute(getProject(1,lngRecordId))

		If not (objRS.BOF and objRS.EOF) then
			lngPrevId = objRS.fields("PrevId").value
			lngNextId = objRS.fields("NextId").value
			strClient = showString(objRS.fields("C_Client").value)
			strDivision = showString(objRS.fields("D_Division").value)
			lngDivId = objRS.fields("DivId").value
			strTitle = showString(objRS.fields("P_Title").value)
			lngInvoiceId = objRS.fields("InvoiceId").value
			strOwner = showString(objRS.fields("Owner").value)
			intDaysTotal = objRS.fields("P_DaysTotal").value
			intDaysOwed = objRS.fields("P_DaysOwed").value
			strClosed = showTrueFalse(objRS.fields("P_Closed").value)
		Else
			strTitle = getIDS("IDS_Deleted")
		End If
	End If

	Call DisplayHeader(1)
%>

<div id="modDiv" class="dvMod">

<% Call showToolBar() %>

<table border="0" cellspacing="10" width="100%">
  <tr><td width="50%" align="left" valign="top">

  <table border=0>
	<tr>
	  <td nowrap class="bFont"><% =getIDS("IDS_Account") %></td>
	  <td class="dFont"><% =showLink(2,"../sales/client.asp?id="&lngDivId,strClient) %>&nbsp;</td>
	</tr>
	<tr>
	  <td nowrap class="bFont"><% =getIDS("IDS_Division") %></td>
	  <td class="dFont"><% =strDivision %>&nbsp;</td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_InvoiceId") %></td>
	  <td class="dFont"><% =showLink(7,"../finance/invoice.asp?id="&lngInvoiceId,bigDigitNum(7,lngInvoiceId)) %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Owner") %></td>
	  <td class="dFont"><% =strOwner %></td>
	</tr>
  </table>

  </td>
  <td width="50%" valign="top">

  <table border=0>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_DaysTotal") %></td>
	  <td class="dFont" width=200><% =intDaysTotal %>&nbsp;</td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_DaysOwed") %></td>
	  <td class="dFont"><% =intDaysOwed %></td>
	</tr>
	<tr><td colspan=2 class="dFont">&nbsp;</td></tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Closed") %></td>
	  <td class="dFont"><% =strClosed %></td>
	</tr>
  </table>

  </td></tr>
</table>

<%
	If strTitle <> "Deleted" Then

		strTabBuilder = getIDS("IDS_Summary") & "|i_summary.asp?m=4&mid=" & lngRecordId & _
			"||" & getIDS("IDS_Notes") & "|../common/i_notes.asp?m=4&mid=" & lngRecordId & _
			"||" & getIDS("IDS_Attachments") & "|../common/i_attach.asp?m=4&mid=" & lngRecordId

		'Enable following line allows event logging
		strTabBuilder = strTabBuilder & "||" & getIDS("IDS_Events") & "|../common/i_events.asp?m=4&mid=" & lngRecordId

		strFrame = makeTabs(strTabBuilder)

		Response.Write("</div><iframe id=""contentDiv"" class=""iBorder"" src=""" & strTabURL & """ title=""" & strFrame & """ style=""height:" & intScreenH-210 & "px;"" width=""100%"" scrolling=""no""></iframe>" & vbCrLf)
	End If

	Call DisplayFooter(1)
%>