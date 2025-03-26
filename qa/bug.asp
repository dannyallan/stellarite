<!--#include file="..\_inc\functions.asp" -->
<!--#include file="..\_inc\functions_mod.asp" -->
<!--#include file="..\_inc\sql\sql_qa.asp" -->
<%
	Call pageFunctions(6,1)

	Dim datOpen         'as Date
	Dim datUpdate       'as Date
	Dim datClosed       'as Date
	Dim strOwner        'as String
	Dim strPriority     'as String
	Dim strDuration     'as String
	Dim strProduct      'as String
	Dim strBuild        'as String
	Dim strFrame        'as String

	lngRecordId = valNum(lngRecordId,3,1)
	strTitle = bigDigitNum(7,lngRecordId)

	If strDoAction = "del" and intPerm >= 4 then
		Call delBug(lngUserId,lngRecordId)
		lngPrevId = doPrevNext(0,6,lngRecordId,0,0)
		lngNextId = doPrevNext(1,6,lngRecordId,0,0)
		strTitle = getIDS("IDS_Deleted")
	Else
		Set objRS = objConn.Execute(getBug(1,lngRecordId))

		If not (objRS.BOF and objRS.EOF) then
			lngPrevId = objRS.fields("PrevId").value
			lngNextId = objRS.fields("NextId").value
			datOpen = showDate(1,objRS.fields("B_CreatedDate").value) & " - " & showString(objRS.fields("CreatedBy").value)
			datUpdate = showDate(1,objRS.fields("B_ModDate").value) & " - " & showString(objRS.fields("ModBy").value)
			If objRS.fields("B_Closed").value = 1 Then
				datClosed = showDate(3,objRS.fields("B_CloseDate").value)
				strDuration = showDuration(objRS.fields("B_CreatedDate").value,objRS.fields("B_CloseDate").value)
			Else
				strDuration = showDuration(objRS.fields("B_CreatedDate").value,Now)
			End If
			strOwner = showString(objRS.fields("Owner").value)
			strPriority = getAOS(objRS.fields("B_Priority").value)
			strProduct = getAOS(objRS.fields("B_ProductId").value)
			strBuild = showString(objRS.fields("B_Build").value)
		Else
			strTitle = getIDS("IDS_Deleted")
		End If
	End If

	Call DisplayHeader(1)
%>

<div id="modDiv" class="dvMod">

<%    Call showToolBar() %>


<table border="0" cellspacing="10" width="100%">
  <tr><td width="50%" valign=top>

  <table border=0>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Owner") %></td>
	  <td class="dFont"><% =strOwner %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Priority") %></td>
	  <td class="dFont"><% =strPriority %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Product") %></td>
	  <td class="dFont"><% =strProduct %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Build") %></td>
	  <td class="dFont"><% =strBuild %></td>
	</tr>
  </table>

  </td>
  <td width="50%" valign=top>

  <table border=0>
	<tr>
	  <td nowrap class="bFont""><% =getIDS("IDS_Opened") %></td>
	  <td class="dFont"><% =datOpen %>&nbsp;</td>
	</tr>
	<tr>
	  <td nowrap class="bFont""><% =getIDS("IDS_Updated") %></td>
	  <td class="dFont"><% =datUpdate %></td>
	</tr>
	<tr>
	  <td nowrap class="bFont""><% =getIDS("IDS_Closed") %></td>
	  <td class="dFont"><% =datClosed %></td>
	</tr>
	<tr>
	  <td nowrap class="bFont"><% =getIDS("IDS_Duration") %></td>
	  <td class="dFont"><% =strDuration %></td>
	</tr>
  </table>

  </td></tr>
</table>

<%
	If strTitle <> getIDS("IDS_Deleted") Then

		strTabBuilder = getIDS("IDS_Summary") & "|i_summary.asp?m=6&mid=" & lngRecordId & _
			"||" & getIDS("IDS_Notes") & "|../common/i_notes.asp?m=6&mid=" & lngRecordId & _
			"||" & getIDS("IDS_Attachments") & "|../common/i_attach.asp?m=6&mid=" & lngRecordId

		'Enable following line allows event logging
		'strTabBuilder = strTabBuilder & "||" & getIDS("IDS_Events") & "|../common/i_events.asp?m=6&mid=" & lngRecordId

		If pTickets >= 1 Then strTabBuilder = strTabBuilder & "||" & getIDS("IDS_Tickets") & "|../common/i_tickets.asp?m=" & bytMod & "&mid=" & lngRecordId

		strFrame= makeTabs(strTabBuilder)

		Response.Write("</div><iframe id=""contentDiv"" class=""iBorder"" src=""" & strTabURL & """ title=""" & strFrame & """ style=""height:" & intScreenH-210 & "px;"" width=""100%"" scrolling=""no""></iframe>" & vbCrLf)
	End If

	Call DisplayFooter(1)
%>