<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_qa.asp" -->
<%
	Call pageFunctions(6,2)

	Dim strOwner        'as String
	Dim lngOwnerId      'as Long
	Dim strPriority     'as String
	Dim lngPriorityId   'as Integer
	Dim blnHotIssue     'as String
	Dim strBugType      'as String
	Dim lngBugTypeId    'as Integer
	Dim strBugSource    'as String
	Dim lngBugSourceId  'as Integer
	Dim lngProductId    'as Integer
	Dim strProduct      'as String
	Dim strBuild        'as String
	Dim strDescription  'as String
	Dim strSolution     'as String
	Dim strCause        'as String
	Dim lngCauseId      'as Integer
	Dim blnClosed       'as Boolean
	Dim strCreatedBy    'as String
	Dim strModBy        'as String
	Dim datCreatedDate  'as Date
	Dim datModDate      'as Date
	Dim datCloseDate    'as Date
	Dim strExtraFields  'as String

	strTitle = getIDS("IDS_Edit") & " " & getIDS("IDS_Bug")

	If strDoAction <> "" Then

		Select Case strDoAction
			Case "del"
				If intPerm >= 4 Then Call delBug(lngUserId,lngRecordId)

			Case "new","edit"

				lngOwnerId = getUserId(6,valString(Request.Form("txtOwner"),100,1,0))
				blnHotIssue = valNum(Request.Form("chkHotIssue"),0,0)
				lngPriorityId = valNum(Request.Form("selPriority"),2,-1)
				lngProductId = valNum(Request.Form("selProduct"),2,-1)
				lngBugTypeId = valNum(Request.Form("selBugType"),2,-1)
				lngBugSourceId = valNum(Request.Form("selBugSource"),2,-1)
				strBuild = valString(Request.Form("txtBuild"),10,0,0)
				strDescription = valString(Request.Form("txtDescription"),255,0,4)
				strSolution = valString(Request.Form("txtSolution"),255,0,4)
				blnClosed = valNum(Request.Form("chkClosed"),0,0)
				lngCauseId = valNum(Request.Form("selCause"),2,blnClosed)
				datCloseDate = valDate(Request.Form("txtCloseDate"),blnClosed)

				If strDoAction = "edit" and intPerm >= 3 Then
					Call updateBug(lngUserId,lngRecordId,lngOwnerId,blnHotIssue,lngPriorityId,lngBugTypeId,lngBugSourceId,lngProductId,strBuild,strDescription,strSolution,lngCauseId,blnClosed,datCloseDate)
				ElseIf strDoAction = "new" Then
					lngRecordId = insertBug(lngUserId,lngRecordId,lngOwnerId,blnHotIssue,lngPriorityId,lngBugTypeId,lngBugSourceId,lngProductId,strBuild,strDescription,strSolution,lngCauseId,blnClosed,datCloseDate)
				End If

				Call saveCustomFields(6,lngRecordId)
		End Select
		Call closeEdit()
	Else
		If blnRS and intPerm >= 3 Then

			Set objRS = objConn.Execute(getBug(0,lngRecordId))

			If not (objRS.BOF and objRS.EOF) then
				lngRecordId = objRS.fields("BugId").value
				strOwner = objRS.fields("Owner").value
				blnHotIssue = objRS.fields("B_HotIssue").value
				lngPriorityId = objRS.fields("B_Priority").value
				lngBugTypeId = objRS.fields("B_BugType").value
				lngBugSourceId = objRS.fields("B_BugSource").value
				lngProductId = objRS.fields("B_ProductId").value
				strBuild = objRS.fields("B_Build").value
				strDescription = objRS.fields("B_Description").value
				strSolution = objRS.fields("B_Solution").value
				lngCauseId = objRS.fields("B_Cause").value
				blnClosed = objRS.fields("B_Closed").value
				datCloseDate = objRS.fields("B_CloseDate").value
				strCreatedBy = objRS.fields("CreatedBy").value
				datCreatedDate = objRS.fields("B_CreatedDate").value
				strModBy = objRS.fields("ModBy").value
				datModDate = objRS.fields("B_ModDate").value
			End If
		Elseif blnRS Then
			Call logError(2,1)
		Else
			strCreatedBy = strFullName
			datCreatedDate = Now
			strModBy = strFullName
			datModDate = Now
			If mBugs Then strOwner = strFullName
		End If
		strExtraFields = editCustomFields(6)
		strPriority = getOptionDropDown(60,False,"selPriority","IDS_Priority",lngPriorityId)
		strProduct = getOptionDropDown(150,False,"selProduct","IDS_Product",lngProductId)
		strBugType = getOptionDropDown(150,False,"selBugType","IDS_BugType",lngBugTypeId)
		strBugSource = getOptionDropDown(150,False,"selBugSource","IDS_BugSource",lngBugSourceId)
		strCause = getOptionDropDown(120,True,"selCause","IDS_BugCause",lngCauseId)
	End if

	strIncHead = getCalendarScripts()

	Call DisplayHeader(0)
	Call showEditHeader(strTitle,strCreatedBy,datCreatedDate,strModBy,datModDate)
%>
<form name="frmBug" method="post" action="edit_bug.asp?id=<% =lngRecordId %>">
<div id="contentDiv" class="dvBorder" style="height:<% =intScreenH-170 %>px;"><br />

<table border=0>
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
	<tr>
	  <td width=170><% =getLabel(getIDS("IDS_Owner"),"txtOwner") %></td>
	  <td colspan=3><% =getTextField("txtOwner","mText",strOwner,67,100,"") %>
	  <% =getIconImport(1,getSearchURL("?m=0&rVal=txtOwner"),getIDS("IDS_Owner")) %>
	  </td>
	</tr>
	<tr><td class="dFont" colspan=4>&nbsp;</td></tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_Priority"),"selPriority") %></td>
	  <td><% =strPriority %></td>
	  <td><% =getLabel(getIDS("IDS_HotIssue"),"chkHotIssue") %></td>
	  <td><% =getCheckbox("chkHotIssue",blnHotIssue,"") %></td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_Product"),"selProduct") %></td>
	  <td><% =strProduct %></td>
	  <td><% =getLabel(getIDS("IDS_BugType"),"selBugType") %></td>
	  <td><% =strBugType %></td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_Build"),"txtBuild") %></td>
	  <td><% =getTextField("txtBuild","oText",strBuild,22,10,"") %></td>
	  <td><% =getLabel(getIDS("IDS_BugSource"),"selBugSource") %></td>
	  <td><% =strBugSource %></td>
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
	  <td><% =getCheckbox("chkClosed",blnClosed,"onClick=""doClassChange();""") %>
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
		Response.Write(getIconNew(getEditURL(6,"?m="&bytMod&"&mid="&lngModId)))
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

