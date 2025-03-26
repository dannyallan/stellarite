<!--#include file="..\_inc\functions.asp" -->
<!--#include file="..\_inc\functions_mod.asp" -->
<!--#include file="..\_inc\sql\sql_support.asp" -->
<%
	Call pageFunctions(5,1)

	Dim strContact      'as String
	Dim strClient       'as String
	Dim strEmail        'as String
	Dim strPhone        'as String
	Dim strDuration     'as String
	Dim datOpen         'as Date
	Dim datUpdate       'as Date
	Dim datClosed       'as Date
	Dim strFrame        'as String

	lngRecordId = valNum(lngRecordId,3,1)
	strTitle = bigDigitNum(7,lngRecordId)

	If strDoAction = "del" and intPerm >= 4 then
		Call delTicket(lngUserId,lngRecordId)
		lngPrevId = doPrevNext(0,5,lngRecordId,0,0)
		lngNextId = doPrevNext(1,5,lngRecordId,0,0)
		strTitle = getIDS("IDS_Deleted")
	Else
		Set objRS = objConn.Execute(getTicket(1,lngRecordId))

		If not (objRS.BOF and objRS.EOF) then
			lngPrevId = objRS.fields("PrevId").value
			lngNextId = objRS.fields("NextId").value
			strClient = showLink(2,"../sales/client.asp?id=" & objRS.fields("DivId").value,showString(objRS.fields("C_Client").value))
			If objRS.fields("D_Division").value <> "" Then strClient = strClient & " - " & showString(objRS.fields("D_Division").value)
			datOpen = showDate(1,objRS.fields("T_CreatedDate").value) & " - " & showString(objRS.fields("CreatedBy").value)
			datUpdate = showDate(1,objRS.fields("T_ModDate").value) & " - " & showString(objRS.fields("ModBy").value)
			If objRS.fields("T_Closed").value = 1 Then
				datClosed = showDate(1,objRS.fields("T_CloseDate").value)
				strDuration = showDuration(objRS.fields("T_CreatedDate").value,objRS.fields("T_CloseDate").value)
			Else
				strDuration = showDuration(objRS.fields("T_CreatedDate").value,Now)
			End If
			strContact = showLink(1,"../sales/contact.asp?id=" & objRS.fields("ContactId").value,showString(objRS.fields("Contact").value))
			strPhone = showPhone(objRS.fields("K_Phone1").value)
			If objRS.fields("K_Ext1").value <> "" Then strPhone = strPhone & "&nbsp;&nbsp;<span class=""bFont"">Ext.</span>" & objRS.fields("K_Ext1").value
			strEmail = showEmail(objRS.fields("K_Email").value)
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
	  <td class="bFont"><% =getIDS("IDS_Account") %></td>
	  <td class="dFont"><% =strClient %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Contact") %></td>
	  <td class="dFont"><% =strContact %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Phone") %></td>
	  <td class="dFont"><% =strPhone %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Email") %></td>
	  <td class="dFont"><% =strEmail %></td>
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

		strTabBuilder = getIDS("IDS_Summary") & "|i_summary.asp?m=5&mid=" & lngRecordId & _
			"||" & getIDS("IDS_Notes") & "|../common/i_notes.asp?m=5&mid=" & lngRecordId & _
			"||" & getIDS("IDS_Attachments") & "|../common/i_attach.asp?m=5&mid=" & lngRecordId

		'Enable following line allows event logging
		'strTabBuilder = strTabBuilder & "||" & getIDS("IDS_Events") & "|../common/i_events.asp?m=5&mid=" & lngRecordId

		strFrame = makeTabs(strTabBuilder)

		Response.Write("</div><iframe id=""contentDiv"" class=""iBorder"" src=""" & strTabURL & """ title=""" & strFrame & """ style=""height:" & intScreenH-210 & "px;"" scrolling=""no""></iframe>" & vbCrLf)
	End If

	Call DisplayFooter(1)
%>
