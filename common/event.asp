<!--#include file="..\_inc\functions.asp" -->
<!--#include file="..\_inc\functions_mod.asp" -->
<!--#include file="..\_inc\sql\sql_events.asp" -->
<%
	Call pageFunctions(50,1)

	Dim strParent       'as String
	Dim strOwner        'as String
	Dim strEventType    'as String
	Dim strBillable     'as String
	Dim datStartTime    'as Date
	Dim datEndTime      'as Date
	Dim strFrame        'as String

	strModImage = "event"
	strModItem = getIDS("IDS_Event")

	lngRecordId = valNum(lngRecordId,3,1)
	bytMod = getValue("E_Module","CRM_Events","EventId="&lngRecordId,0)
	lngModId = getValue("E_ModuleId","CRM_Events","EventId="&lngRecordId,0)

	If strDoAction = "del" and intPerm >= 4 Then
		Call delEvent(lngUserId,lngRecordId,bytMod,lngModId)
		lngPrevId = doPrevNext(0,50,lngRecordId,bytMod,lngModId)
		lngNextId = doPrevNext(1,50,lngRecordId,bytMod,lngModId)
		strTitle = getIDS("IDS_Deleted")
	Else
		Set objRS = objConn.Execute(getEvent(1,lngRecordId,bytMod,lngModId))

		If not (objRS.BOF and objRS.EOF) then
			If objRS.fields("E_Permissions").value = 1 and intMember > 1 Then Call sendBack(getIDS("IDS_MsgMembersOnly"))
			lngPrevId = objRS.fields("PrevId").value
			lngNextId = objRS.fields("NextId").value
			strEventType = getAOS(objRS.fields("E_EventType").value)
			strTitle = showString(objRS.fields("E_Title").value)
			strOwner = showString(objRS.fields("Owner").value)
			strParent = showString(objRS.fields("Title").value)
			datStartTime = showDate(1,objRS.fields("E_StartTime").value)
			datEndTime = showDate(1,objRS.fields("E_EndTime").value)
			strBillable = showTrueFalse(objRS.fields("E_Billable").value)
			If objRS.fields("E_Onsite").value = 1 Then strEventType = strEventType & " - " & getIDS("IDS_Onsite")
		Else
			strTitle = getIDS("IDS_Deleted")
		End If
	End If

	Call DisplayHeader(1)
%>

<div id="modDiv" class="dvMod">

<% Call showToolBar() %>

<table border="0" cellspacing="10" width="100%">
  <tr><td width="70%" valign=top>

  <table border=0>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_ModItem" & bytMod) %></td>
	  <td class="dFont">
<%
	Select Case bytMod
		Case 1
			Response.Write(showLink(1,"../sales/contact.asp?id="&lngModId,strParent))
		Case 2
			Response.Write(showLink(2,"../sales/client.asp?id="&lngModId,strParent))
		Case 3
			Response.Write(showLink(3,"../sales/sale.asp?id="&lngModId,strParent))
		Case 4
			Response.Write(showLink(4,"../services/project.asp?id="&lngModId,strParent))
		Case 5
			Response.Write(showLink(5,"../support/ticket.asp?id="&lngModId,strParent))
		Case 6
			Response.Write(showLink(6,"../qa/bug.asp?id="&lngModId,strParent))
		Case 7
			Response.Write(showLink(7,"../finance/invoice.asp?id="&lngModId,strParent))
	End Select
%>
	  </td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Owner") %></td>
	  <td class="dFont"><% =strOwner %></td>
	</tr>
  </table>

  </td>
  <td width="30%" valign=top>

  <table border=0>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Event") %></td>
	  <td class="dFont"><% =strEventType %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Billable") %></td>
	  <td class="dFont"><% =strBillable %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_StartTime") %></td>
	  <td class="dFont"><% =datStartTime %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_EndTime") %></td>
	  <td class="dFont"><% =datEndTime %></td>
	</tr>
  </table>

  </td></tr>
</table>

<%    If strTitle <> getIDS("IDS_Deleted") Then

		strTabBuilder = getIDS("IDS_Notes") & "|i_notes.asp?m=" & bytMod & "&mid=" & lngModId & "&eid=" & lngRecordId & _
			"||" & getIDS("IDS_Attachments") & "|i_attach.asp?m=" & bytMod & "&mid=" & lngModId & "&eid=" & lngRecordId

		strFrame = makeTabs(strTabBuilder)

		Response.Write("</div><iframe id=""contentDiv"" class=""iBorder"" src=""" & strTabURL & """ title=""" & strFrame & """ style=""height:" & intScreenH-210 & "px;"" width=""100%"" scrolling=""no""></iframe>" & vbCrLf)
	End If

	Call DisplayFooter(1)
%>