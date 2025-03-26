<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_support.asp" -->
<%
	Call pageFunctions(5,2)

	Dim strTicket       'as String
	Dim strContact      'as String
	Dim strOwner        'as String
	Dim lngOwnerId      'as Long
	Dim strPriority     'as String
	Dim lngPriorityId   'as Integer
	Dim blnHotIssue     'as String
	Dim strBuild        'as String
	Dim lngProductId    'as Long
	Dim strProduct      'as String
	Dim strTicketType   'as String
	Dim lngTicketTypeId 'as Integer
	Dim strTicketSource 'as String
	Dim lngTicketSourceId   'as Integer
	Dim strSupportType      'as String
	Dim lngSupportTypeId    'as Integer
	Dim strDescription  'as String
	Dim strSolution     'as String
	Dim blnClosed       'as Boolean
	Dim strCause        'as String
	Dim lngCauseId      'as Integer
	Dim lngBugId        'as Long
	Dim lngDivId        'as Long
	Dim lngContactId    'as Long
	Dim strCreatedBy    'as String
	Dim strModBy        'as String
	Dim datCreatedDate  'as Date
	Dim datModDate      'as Date
	Dim datClosed       'as Date
	Dim datCloseDate    'as Date
	Dim strExtraFields  'as String

	strTitle = getIDS("IDS_Edit") & " " & getIDS("IDS_Ticket")

	If strDoAction <> "" Then

		Select Case strDoAction
			Case "del"
				If intPerm >= 4 Then Call delTicket(lngUserId,lngRecordId)

			Case "new","edit"

				lngOwnerId = getUserId(5,valString(Request.Form("txtOwner"),100,1,0))
				lngContactId = valNum(Request.Form("hdnContact"),3,1)
				lngDivId = valNum(Request.Form("hdnDivision"),3,1)
				blnHotIssue = valNum(Request.Form("chkHotIssue"),0,0)
				lngPriorityId = valNum(Request.Form("selPriority"),2,-1)
				lngProductId = valNum(Request.Form("selProduct"),2,-1)
				lngTicketTypeId = valNum(Request.Form("selTicketType"),2,-1)
				lngTicketSourceId = valNum(Request.Form("selTicketSource"),2,-1)
				lngSupportTypeId = valNum(Request.Form("selSupportType"),2,-1)
				strBuild = valString(Request.Form("txtBuild"),10,0,0)
				lngBugId = valNum(Request.Form("txtBugNumber"),3,-1)
				strDescription = valString(Request.Form("txtDescription"),255,0,4)
				strSolution = valString(Request.Form("txtSolution"),255,0,4)
				blnClosed = valNum(Request.Form("chkClosed"),0,0)
				lngCauseId = valNum(Request.Form("selCause"),2,blnClosed)
				datCloseDate = valDate(Request.Form("txtCloseDate"),blnClosed)

				If strDoAction = "edit" and intPerm >= 3 Then

					Call updateTicket(lngUserId,lngRecordId,bytMod,lngModId,lngContactId,lngDivId,lngOwnerId,blnHotIssue,lngPriorityId, _
							lngTicketTypeId,lngTicketSourceId,lngSupportTypeId,lngProductId,strBuild,lngBugId, _
							strDescription,strSolution,lngCauseId,blnClosed,datCloseDate)

				ElseIf strDoAction = "new" Then

					lngRecordId = insertTicket(lngUserId,lngRecordId,bytMod,lngModId,lngContactId,lngDivId,lngOwnerId,blnHotIssue,lngPriorityId, _
							lngTicketTypeId,lngTicketSourceId,lngSupportTypeId,lngProductId,strBuild,lngBugId, _
							strDescription,strSolution,lngCauseId,blnClosed,datCloseDate)
				End If

				Call saveCustomFields(5,lngRecordId)
		End Select

		If bytMenu = 0 Then Session("LastPage") = strCRMURL & "common/i_tickets.asp?m=" & bytMod & "&mid=" & lngModId
		Call closeEdit()
	Else
		If blnRS and intPerm >= 3 Then
			Set objRS = objConn.Execute(getTicket(0,lngRecordId))

			If not (objRS.BOF and objRS.EOF) then
				lngRecordId = objRS.fields("TicketId").value
				strContact = objRS.fields("Contact").value
				lngContactId = objRS.fields("ContactId").value
				lngDivId = objRS.fields("DivId").value
				strOwner = objRS.fields("Owner").value
				blnHotIssue = objRS.fields("T_HotIssue").value
				lngPriorityId = objRS.fields("T_Priority").value
				lngTicketTypeId = objRS.fields("T_TicketType").value
				lngSupportTypeId = objRS.fields("T_SupportType").value
				lngProductId = objRS.fields("T_ProductId").value
				strBuild = objRS.fields("T_Build").value
				lngBugId = objRS.fields("T_BugId").value
				strDescription = objRS.fields("T_Description").value
				strSolution = objRS.fields("T_Solution").value
				lngCauseId = objRS.fields("T_Cause").value
				blnClosed = objRS.fields("T_Closed").value
				datCloseDate = objRS.fields("T_CloseDate").value
				strCreatedBy = objRS.fields("CreatedBy").value
				datCreatedDate = objRS.fields("T_CreatedDate").value
				strModBy = objRS.fields("ModBy").value
				datModDate = objRS.fields("T_ModDate").value
			End If
		Elseif blnRS Then
			Call logError(2,1)
		Else
			If mTickets Then strOwner = strFullName
			If bytMod = 1 Then
				lngContactId = lngModId
				lngDivId = getValue("DivId","CRM_Contacts","ContactId = " & lngContactId,0)
				strContact = getValue(doConCat(doConCat("K_FirstName","' '"),"K_LastName"),"CRM_Contacts","ContactId = " & lngModId,"")
			Elseif bytMod = 2 Then
				lngDivId = lngModId
				lngContactId = getValue("ContactId","CRM_Contacts","DivId = " & lngDivId,"")
				If lngContactId = "" Then strContact = "" Else strContact = getValue(doConCat(doConCat("K_FirstName","' '"),"K_LastName"),"CRM_Contacts","ContactId = " & lngContactId,"")
			End If
			strCreatedBy = strFullName
			datCreatedDate = Now
			strModBy = strFullName
			datModDate = Now
		End If
		strExtraFields = editCustomFields(5)
		strPriority = getOptionDropDown(60,False,"selPriority","IDS_Priority",lngPriorityId)
		strProduct = getOptionDropDown(150,False,"selProduct","IDS_Product",lngProductId)
		strSupportType = getOptionDropDown(120,False,"selSupportType","IDS_TicketSupport",lngSupportTypeId)
		strTicketType = getOptionDropDown(120,False,"selTicketType","IDS_TicketType",lngTicketTypeId)
		strTicketSource = getOptionDropDown(120,False,"selTicketSource","IDS_TicketSource",lngTicketSourceId)
		strCause = getOptionDropDown(120,True,"selCause","IDS_TicketCause",lngCauseId)
	End if

	strIncHead = getCalendarScripts()

	Call DisplayHeader(0)
	Call showEditHeader(strTitle,strCreatedBy,datCreatedDate,strModBy,datModDate)
%>
<form name="frmTicket" method="post" action="edit_ticket.asp?id=<% =lngRecordId %>&m=<% =bytMod %>&mid=<% =lngModId %>&menu=<% =bytMenu %>">
<div id="contentDiv" class="dvBorder" style="height:<% =intScreenH-170 %>px;"><br />

<table border=0>
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
<% =getHidden("hdnContact",lngContactId) %>
<% =getHidden("hdnDivision",lngDivId) %>
	<tr>
	  <td width=170><% =getLabel(getIDS("IDS_Contact"),"txtContact") %></td>
	  <td colspan=3><% =getTextField("txtContact","mText",strContact,63,100,"readonly=""readonly""") %>
<%
	If pContacts >= 1 Then
		Response.Write(getIconImport(1,getSearchURL("?m=1&rVal=K"),getIDS("IDS_Contact")))
	End If
	If pContacts >= 2 Then
		Response.Write(getIconImport(3,getEditURL(1,"?m="&bytMod&"&mid="&lngModId),getIDS("IDS_Contact")))
	End If
%>
	  </td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_Owner"),"txtOwner") %></td>
  <td colspan=3><% =getTextField("txtOwner","mText",strOwner,67,100,"") %>
	  <% =getIconImport(1,getSearchURL("?m=0&rVal=txtOwner"),getIDS("IDS_Owner")) %>
	  </td>
	</tr>
	<tr><td class="dFont" colspan=4>&nbsp;</td></tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_Priority"),"selPriority") %></td>
	  <td><% =strPriority %></td>
	  <td><% =getLabel(getIDS("IDS_SupportType"),"selSupportType") %></td>
	  <td><% =strSupportType %></td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_HotIssue"),"chkHotIssue") %></td>
	  <td><% =getCheckbox("chkHotIssue",blnHotIssue,"") %>
	  <td><% =getLabel(getIDS("IDS_TicketType"),"selTicketType") %></td>
	  <td><% =strTicketType %></td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_Product"),"selProduct") %></td>
	  <td><% =strProduct %></td>
	  <td><% =getLabel(getIDS("IDS_TicketSource"),"selTicketSource") %></td>
	  <td><% =strTicketSource %></td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_Build"),"txtBuild") %></td>
	  <td><% =getTextField("txtBuild","oText",strBuild,22,10,"") %></td>
	  <td><% =getLabel(getIDS("IDS_BugId"),"txtBugNumber") %></td>
	  <td><% =getTextField("txtBugNumber","oLong",lngBugId,8,8,"readonly=""readonly""") %>
<%
	If pBugs >= 1 Then
		Response.Write(getIconImport(1,getSearchURL("?m=6&rVal=txtBugNumber"),getIDS("IDS_BugId")))
	End If
%>
	  </td>
	</tr>
	<tr><td class="dFont" colspan=4>&nbsp;</td></tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_Description"),"txtDescription") %></td>
	  <td colspan=3><% =getTextArea("txtDescription","oMemo",strDescription,"420px",4,"") %></td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_Solution"),"txtSolution") %></td>
	  <td colspan=3><% =getTextArea("txtSolution","oMemo",strDescription,"420px",4,"") %></td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_Closed"),"chkClosed") %></td>
	  <td><% =getCheckbox("chkClosed",blnClosed,"onClick=""doClassChange();""") %></td>
	  <td><% =getLabel(getIDS("IDS_CloseDate"),"txtCloseDate") %></td>
	  <td><% =getDateField("txtCloseDate","oDate",datCloseDate,getIDS("IDS_CloseDate")) %></td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_Cause"),"selCause") %></td>
	  <td colspan=3><% =strCause %></td>
	</tr>
<%	=strExtraFields %>
</table>

</div>

<div id="footerDiv" class="dvFooter">
<%
	If blnRS Then
		Response.Write(getIconNew(getEditURL(5,"?m="&bytMod&"&mid="&lngModId)))
		If intPerm >= 4 Then Response.Write(getIconDelete())
	End If
	Response.Write(getIconSave(strAction))
	Response.Write(getIconCancel("back"))
%>
</div>
</form>

<script language="JavaScript" type="text/javascript">
	doClassChange();
	function doClassChange() {
		if (document.forms[0].chkClosed.checked) {
			if (document.forms[0].txtCloseDate.value == "")
				document.forms[0].txtCloseDate.value = getToday();
			document.forms[0].txtCloseDate.className = "mDate";
			document.forms[0].selCause.className = "mText";
		}
		else {
			document.forms[0].txtCloseDate.className = "oDate";
			document.forms[0].selCause.className = "oText";
		}
	}
</script>

<%
	Call DisplayFooter(0)
%>
