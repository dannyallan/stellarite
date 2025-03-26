<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_support.asp" -->
<%
	Call pageFunctions(5,2)

	Dim strTicket		'as String
	Dim strContact		'as String
	Dim strOwner		'as String
	Dim lngOwnerId		'as Long
	Dim strPriority		'as String
	Dim intPriorityId	'as Integer
	Dim blnHotIssue		'as String
	Dim strBuild		'as String
	Dim strProduct		'as String
	Dim strTicketType	'as String
	Dim intTicketTypeId	'as Integer
	Dim strTicketSource	'as String
	Dim intTicketSourceId	'as Integer
	Dim strSupportType	'as String
	Dim intSupportTypeId	'as Integer
	Dim strDescription	'as String
	Dim strSolution		'as String
	Dim blnClosed		'as Boolean
	Dim strCause		'as String
	Dim intCauseId		'as Integer
	Dim lngProductId	'as Long
	Dim lngVersionId	'as Long
	Dim lngBugId		'as Long
	Dim lngDivId		'as Long
	Dim lngContactId	'as Long
	Dim strCreatedBy	'as String
	Dim strModBy		'as String
	Dim datCreatedDate	'as Date
	Dim datModDate		'as Date
	Dim datClosed		'as Date
	Dim datCloseDate	'as Date

	strTitle = Application("IDS_Edit") & " " & Application("IDS_Ticket")

	If strDoAction <> "" Then

		lngOwnerId = getUserId(5,valString(Request.Form("txtOwner"),100,1,0))
		lngContactId = valNum(Request.Form("hdnContact"),3,1)
		lngDivId = valNum(Request.Form("hdnDivision"),3,1)
		blnHotIssue = valNum(Request.Form("chkHotIssue"),0,0)
		intPriorityId = valNum(Request.Form("selPriority"),2,-1)
		intTicketTypeId = valNum(Request.Form("selTicketType"),2,-1)
		intTicketSourceId = valNum(Request.Form("selTicketSource"),2,-1)
		intSupportTypeId = valNum(Request.Form("selSupportType"),2,-1)
		strBuild = valString(Request.Form("txtBuild"),10,0,0)
		lngBugId = valNum(Request.Form("txtBugNumber"),3,-1)
		strDescription = valString(Request.Form("txtDescription"),255,0,4)
		strSolution = valString(Request.Form("txtSolution"),255,0,4)
		blnClosed = valNum(Request.Form("chkClosed"),0,0)
		intCauseId = valNum(Request.Form("selCause"),2,blnClosed)
		datCloseDate = valDate(Request.Form("txtCloseDate"),blnClosed)
		strProduct = valString(Request.Form("selProduct"),20,0,0)

		If Instr(strProduct,".") > 0 Then
			lngProductId = valNum(Left(strProduct,Instr(strProduct,".")-1),3,1)
			lngVersionId = valNum(Mid(strProduct,Instr(strProduct,".")+1),3,1)
		Else
			lngProductId = "NULL"
			lngVersionId = "NULL"
		End If

		If strDoAction = "del" and intPerm >= 5 Then

			Call delTicket(lngUserId,lngRecordId)

		ElseIf strDoAction = "edit" and intPerm >= 3 Then

			Call updateTicket(lngUserId,lngRecordId,bytMod,lngModId,lngContactId,lngDivId,lngOwnerId,blnHotIssue,intPriorityId, _
					intTicketTypeId,intTicketSourceId,intSupportTypeId,lngProductId,lngVersionId,strBuild,lngBugId, _
					strDescription,strSolution,intCauseId,blnClosed,datCloseDate)

		ElseIf strDoAction = "new" Then

			lngRecordId = insertTicket(lngUserId,lngRecordId,bytMod,lngModId,lngContactId,lngDivId,lngOwnerId,blnHotIssue,intPriorityId, _
					intTicketTypeId,intTicketSourceId,intSupportTypeId,lngProductId,lngVersionId,strBuild,lngBugId, _
					strDescription,strSolution,intCauseId,blnClosed,datCloseDate)
		End If
		Call closeWindow(strOpenerURL)
	Else
		If blnRS and intPerm >= 3 Then
			Set objRS = objConn.Execute(getTicket(0,lngRecordId))

			If not (objRS.BOF and objRS.EOF) then
				lngRecordId = objRS.fields("TicketId").value
				strContact = showString(objRS.fields("Contact").value)
				lngContactId = objRS.fields("ContactId").value
				lngDivId = objRS.fields("DivId").value
				strOwner = showString(objRS.fields("Owner").value)
				blnHotIssue = objRS.fields("T_HotIssue").value
				intPriorityId = objRS.fields("T_Priority").value
				intTicketTypeId = objRS.fields("T_TicketType").value
				intSupportTypeId = objRS.fields("T_SupportType").value
				strProduct = objRS.fields("ProductId").value & "." & objRS.fields("VersionId").value
				strBuild = showString(objRS.fields("T_Build").value)
				lngBugId = objRS.fields("T_BugId").value
				strDescription = showString(objRS.fields("T_Description").value)
				strSolution = showString(objRS.fields("T_Solution").value)
				intCauseId = objRS.fields("T_Cause").value
				blnClosed = objRS.fields("T_Closed").value
				datCloseDate = showDate(0,objRS.fields("T_CloseDate").value)
				strCreatedBy = showString(objRS.fields("CreatedBy").value)
				datCreatedDate = objRS.fields("T_CreatedDate").value
				strModBy = showString(objRS.fields("ModBy").value)
				datModDate = objRS.fields("T_ModDate").value
			End If
		Elseif blnRS Then
			Call doRedirect("pop_ticket.asp?m=" & bytMod & "&mid=" & lngModId)
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
		strPriority = getOptionDropDown(60,False,"selPriority","Priority",intPriorityId)
		strSupportType = getOptionDropDown(120,False,"selSupportType","Support Type",intSupportTypeId)
		strTicketType = getOptionDropDown(120,False,"selTicketType","Ticket Type",intTicketTypeId)
		strProduct = getProductVersionDropDown(150,False,"selProduct",strProduct)
		strTicketSource = getOptionDropDown(120,False,"selTicketSource","Ticket Source",intTicketSourceId)
		strCause = getOptionDropDown(120,True,"selCause","Ticket Cause",intCauseId)
	End if

	strIncHead = getCalendarScripts()

	Call DisplayHeader(3)
	Call showEditHeader(strTitle,strCreatedBy,datCreatedDate,strModBy,datModDate)
%>
<div id="contentDiv" class="dvBorder" style="height:470px;"><br>

<table border=0 width="100%">
<form name="frmTicket" method="post" action="pop_ticket.asp?id=<% =lngRecordId %>&m=<% =bytMod %>&mid=<% =lngModId %>">
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
<% =getHidden("hdnWinOpen","") %>
<% =getHidden("hdnContact",lngContactId) %>
<% =getHidden("hdnDivision",lngDivId) %>
    <tr>
      <td><% =getLabel(Application("IDS_Contact"),"txtContact") %></td>
      <td colspan=3><% =getTextField("txtContact","mText",strContact,63,100,"readonly=""readonly""") %>
      <% If pContacts >= 1 Then %>
      <a href="<% =newWindow("S","?m=1&rVal=K") %>"><img src="../images/import.gif" alt="<% =getImport("IDS_Contact") %>" border=0 height=16 width=16></a>
      <% End If %>
      <% If pContacts >= 2 Then %>
      <a href="<% =newWindow(1,"?m=6") %>"><img src="../images/new2.gif" alt="<% =getImport("IDS_ContactNew") %>" border=0 height=16 width=16></a></td>
      <% End If %>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_Owner"),"txtOwner") %></td>
      <td colspan=3><% =getTextField("txtOwner","mText",strOwner,67,100,"") %>
      <a href="<% =newWindow("S","?m=0&rVal=txtOwner") %>"><img src="../images/import.gif" alt="<% =getImport("IDS_Owner") %>" border=0 height=16 width=16></a></td>
    </tr>
    <tr><td class="dfont" colspan=4>&nbsp;</td></tr>
    <tr>
      <td><% =getLabel(Application("IDS_Priority"),"selPriority") %></td>
      <td><% =strPriority %></td>
      <td><% =getLabel(Application("IDS_SupportType"),"selSupportType") %></td>
      <td><% =strSupportType %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_HotIssue"),"chkHotIssue") %></td>
      <td><% =getCheckbox("chkHotIssue",blnHotIssue,"") %>
      <td><% =getLabel(Application("IDS_TicketType"),"selTicketType") %></td>
      <td><% =strTicketType %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_Product"),"selProduct") %></td>
      <td><% =strProduct %></td>
      <td><% =getLabel(Application("IDS_TicketSource"),"selTicketSource") %></td>
      <td><% =strTicketSource %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_Build"),"txtBuild") %></td>
      <td><% =getTextField("txtBuild","oText",strBuild,22,10,"") %></td>
      <td><% =getLabel(Application("IDS_BugId"),"txtBugNumber") %></td>
      <td><% =getTextField("txtBugNumber","oLong",lngBugId,8,8,"readonly=""readonly""") %>
      <% If pBugs >= 1 Then %>
      <a href="<% =newWindow("S","?m=6&rVal=txtBugNumber") %>"><img src="../images/import.gif" alt="<% =getImport("IDS_BugId") %>" border=0 height=16 width=16></a>
      <% End If %>
      </td>
    </tr>
    <tr><td class="dfont" colspan=4>&nbsp;</td></tr>
    <tr>
      <td><% =getLabel(Application("IDS_Description"),"txtDescription") %></td>
      <td colspan=3><% =getTextArea("txtDescription","oMemo",strDescription,"420px",4,"") %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_Solution"),"txtSolution") %></td>
      <td colspan=3><% =getTextArea("txtSolution","oMemo",strDescription,"420px",4,"") %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_Closed"),"chkClosed") %></td>
      <td><% =getCheckbox("chkClosed",blnClosed,"onClick=""doClassChange();""") %></td>
      <td><% =getLabel(Application("IDS_CloseDate"),"txtCloseDate") %></td>
      <td><% =getTextField("txtCloseDate","oDate",datCloseDate,12,12,"") %>
      <a href="Javascript:showCalendar('txtCloseDate');"><img src="../images/cal.gif" alt="<% =getImport("IDS_CloseDate") %>" border=0 height=16 width=16></a></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_Cause"),"selCause") %></td>
      <td colspan=3><% =strCause %></td>
    </tr>
</form>
</table>

</div>

<div id="footerDiv" class="dvFooter">
<%
	If blnRS Then
		Response.Write(getIconNew("pop_ticket.asp"))
		If intPerm >= 4 Then Response.Write(getIconDelete())
	End If
	Response.Write(getIconSave(strAction))
	Response.Write(getIconCancel())
%>
</div>

<script language="Javascript">
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
	Call DisplayFooter(3)
%>
