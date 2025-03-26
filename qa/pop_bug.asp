<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_qa.asp" -->
<%
	Call pageFunctions(6,2)

	Dim strOwner		'as String
	Dim lngOwnerId		'as Long
	Dim strPriority		'as String
	Dim intPriorityId	'as Integer
	Dim blnHotIssue		'as String
	Dim strBugType		'as String
	Dim intBugTypeId	'as Integer
	Dim strBugSource	'as String
	Dim intBugSourceId	'as Integer
	Dim strProduct		'as String
	Dim strBuild		'as String
	Dim intProductId	'as Integer
	Dim intVersionId	'as Integer
	Dim strDescription	'as String
	Dim strSolution		'as String
	Dim strCause		'as String
	Dim intCauseId		'as Integer
	Dim blnClosed		'as Boolean
	Dim strCreatedBy	'as String
	Dim strModBy		'as String
	Dim datCreatedDate	'as Date
	Dim datModDate		'as Date
	Dim datCloseDate	'as Date

	strTitle = Application("IDS_Edit") & " " & Application("IDS_Bug")

	If strDoAction <> "" Then

		lngOwnerId = getUserId(6,valString(Request.Form("txtOwner"),100,1,0))
		blnHotIssue = valNum(Request.Form("chkHotIssue"),0,0)
		intPriorityId = valNum(Request.Form("selPriority"),2,-1)
		intBugTypeId = valNum(Request.Form("selBugType"),2,-1)
		intBugSourceId = valNum(Request.Form("selBugSource"),2,-1)
		strBuild = valString(Request.Form("txtBuild"),10,0,0)
		strDescription = valString(Request.Form("txtDescription"),255,0,4)
		strSolution = valString(Request.Form("txtSolution"),255,0,4)
		blnClosed = valNum(Request.Form("chkClosed"),0,0)
		intCauseId = valNum(Request.Form("selCause"),2,blnClosed)
		datCloseDate = valDate(Request.Form("txtCloseDate"),blnClosed)
		strProduct = valString(Request.Form("selProduct"),20,0,0)

		If Instr(strProduct,".") > 0 Then
			intProductId = Left(strProduct,Instr(strProduct,".")-1)
			intVersionId = Mid(strProduct,Instr(strProduct,".")+1)
		Else
			intProductId = "NULL"
			intVersionId = "NULL"
		End If

		If strDoAction = "del" and intPerm >= 4 Then

			Call delBug(lngUserId,lngRecordId)

		ElseIf strDoAction = "edit" and intPerm >= 3 Then

			Call updateBug(lngUserId,lngRecordId,lngOwnerId,blnHotIssue,intPriorityId,intBugTypeId,intBugSourceId,intProductId, _
					intVersionId,strBuild,strDescription,strSolution,intCauseId,blnClosed,datCloseDate)

		ElseIf strDoAction = "new" Then

			lngRecordId = insertBug(lngUserId,lngRecordId,lngOwnerId,blnHotIssue,intPriorityId,intBugTypeId,intBugSourceId,intProductId, _
					intVersionId,strBuild,strDescription,strSolution,intCauseId,blnClosed,datCloseDate)
		End If
		Call closeWindow(strOpenerURL)
	Else
		If blnRS and intPerm >= 3 Then

			Set objRS = objConn.Execute(getBug(0,lngRecordId))

			If not (objRS.BOF and objRS.EOF) then
				lngRecordId = objRS.fields("BugId").value
				strOwner = showString(objRS.fields("Owner").value)
				blnHotIssue = showString(objRS.fields("B_HotIssue").value)
				intPriorityId = showString(objRS.fields("B_Priority").value)
				intBugTypeId = showString(objRS.fields("B_BugType").value)
				intBugSourceId = showString(objRS.fields("B_BugSource").value)
				strProduct = objRS.fields("ProductId").value & "." & objRS.fields("VersionId").value
				strBuild = showString(objRS.fields("B_Build").value)
				strDescription = showString(objRS.fields("B_Description").value)
				strSolution = showString(objRS.fields("B_Solution").value)
				intCauseId = showString(objRS.fields("B_Cause").value)
				blnClosed = showString(objRS.fields("B_Closed").value)
				datCloseDate = showDate(0,objRS.fields("B_CloseDate").value)
				strCreatedBy = showString(objRS.fields("CreatedBy").value)
				datCreatedDate = objRS.fields("B_CreatedDate").value
				strModBy = showString(objRS.fields("ModBy").value)
				datModDate = objRS.fields("B_ModDate").value
			End If
		Elseif blnRS Then
			Call doRedirect("pop_bug.asp")
		Else
			strCreatedBy = strFullName
			datCreatedDate = Now
			strModBy = strFullName
			datModDate = Now
			If mBugs Then strOwner = strFullName
		End If
		strPriority = getOptionDropDown(60,False,"selPriority","Priority",intPriorityId)
		strProduct = getProductVersionDropDown(150,False,"selProduct",strProduct)
		strBugType = getOptionDropDown(150,False,"selBugType","Bug Type",intBugTypeId)
		strBugSource = getOptionDropDown(150,False,"selBugSource","Bug Source",intBugSourceId)
		strCause = getOptionDropDown(120,True,"selCause","Bug Cause",intCauseId)
	End if

	strIncHead = getCalendarScripts()

	Call DisplayHeader(3)
	Call showEditHeader(strTitle,strCreatedBy,datCreatedDate,strModBy,datModDate)
%>
<div id="contentDiv" class="dvBorder" style="height:405px;"><br>

<table border=0 width="100%">
<form name="frmBug" method="post" action="pop_bug.asp?id=<% =lngRecordId %>">
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
<% =getHidden("hdnWinOpen","") %>
    <tr>
      <td><% =getLabel(Application("IDS_Owner"),"txtOwner") %></td>
      <td colspan=3><% =getTextField("txtOwner","mText",strOwner,67,100,"") %>
      <a href="<% =newWindow("S","?m=0&rVal=txtOwner") %>"><img src="../images/import.gif" alt="<% =getImport("IDS_Owner") %>" border=0 height=16 width=16></a></td>
    </tr>
    <tr><td class="dfont" colspan=4>&nbsp;</td></tr>
    <tr>
      <td><% =getLabel(Application("IDS_Priority"),"selPriority") %></td>
      <td><% =strPriority %></td>
      <td><% =getLabel(Application("IDS_HotIssue"),"chkHotIssue") %></td>
      <td><% =getCheckbox("chkHotIssue",blnHotIssue,"") %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_Product"),"selProduct") %></td>
      <td><% =strProduct %></td>
      <td><% =getLabel(Application("IDS_BugType"),"selBugType") %></td>
      <td><% =strBugType %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_Build"),"txtBuild") %></td>
      <td><% =getTextField("txtBuild","oText",strBuild,22,10,"") %></td>
      <td><% =getLabel(Application("IDS_BugSource"),"selBugSource") %></td>
      <td><% =strBugSource %></td>
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
      <td><% =getCheckbox("chkClosed",blnClosed,"onClick=""doClassChange();""") %>
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
		Response.Write(getIconNew("pop_bug.asp"))
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

