<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_services.asp" -->
<%
	Call pageFunctions(4,2)

	Dim strProject		'as String
	Dim strClient		'as String
	Dim strDivision		'as String
	Dim lngSaleId		'as Long
	Dim lngInvoice		'as String
	Dim strOwner		'as String
	Dim lngOwnerId		'as Long
	Dim strFrame		'as String
	Dim intDaysTotal	'as Integer
	Dim intDaysOwed		'as Integer
	Dim strShortDesc	'as String
	Dim lngDivId		'as Long
	Dim blnClosed		'as Boolean
	Dim datCloseDate	'as Date
	Dim strCreatedBy	'as String
	Dim datCreatedDate	'as Date
	Dim strModBy		'as String
	Dim datModDate		'as Date

	strTitle = Application("IDS_Edit") & " " & Application("IDS_Project")

	If strDoAction <> "" Then

		strOwner = valString(Request.Form("txtOwner"),100,1,0)
		lngOwnerId = getUserId(4,strOwner)
		strProject = valString(Request.Form("txtProject"),40,1,0)
		lngInvoice = valNum(Request.Form("txtInvoice"),3,-1)
		intDaysTotal = valNum(Request.Form("txtDaysTotal"),2,0)
		intDaysOwed = valNum(Request.Form("txtDaysOwed"),2,-1)
		strShortDesc = valString(Request.Form("txtShortDesc"),255,0,4)
		blnClosed = valNum(Request.Form("chkClosed"),0,0)
		datCloseDate = valDate(Request.Form("txtCloseDate"),blnClosed)

		If bytMod = 4 Then
			strClient = valString(Request.Form("txtClient"),40,1,0)
			strDivision = valString(Request.Form("txtDivision"),40,0,0)
		End If

		If strDoAction = "del" and intPerm >= 4 Then

			Call delProject(lngUserId,lngRecordId)

		ElseIf strDoAction = "edit" and intPerm >= 3 Then

			Call updateProject(lngUserId,lngRecordId,bytMod,lngModId,strProject,lngOwnerId,datCloseDate, _
					lngInvoice,intDaysTotal,intDaysOwed,blnClosed,strClient,strDivision,strShortDesc)

		ElseIf strDoAction = "new" Then

			lngRecordId = insertProject(lngUserId,lngRecordId,bytMod,lngModId,strProject,lngOwnerId,datCloseDate, _
					lngInvoice,intDaysTotal,intDaysOwed,blnClosed,strClient,strDivision,strShortDesc)
		End If
		Call closeWindow(strOpenerURL)
	Else

		If blnRS and intPerm >= 3 Then
			Set objRS = objConn.Execute(getProject(0,lngRecordId))

			If not (objRS.BOF and objRS.EOF) then
				lngRecordId = objRS.fields("ProjectId").value
				strProject = showString(objRS.fields("P_Title").value)
				strClient = showString(objRS.fields("C_Client").value)
				strDivision = showString(objRS.fields("D_Division").value)
				lngDivId = objRS.fields("DivId").value
				lngInvoice = showString(objRS.fields("InvoiceId").value)
				strOwner = showString(objRS.fields("Owner").value)
				intDaysTotal = objRS.fields("P_DaysTotal").value
				intDaysOwed = objRS.fields("P_DaysOwed").value
				strShortDesc = showString(objRS.fields("P_ShortDesc").value)
				blnClosed = objRS.fields("P_Closed").value
				datCloseDate = showDate(0,objRS.fields("P_CloseDate").value)
				strCreatedBy = showString(objRS.fields("CreatedBy").value)
				datCreatedDate = objRS.fields("P_CreatedDate").value
				strModBy = showString(objRS.fields("ModBy").value)
				datModDate = objRS.fields("P_ModDate").value
			End If
		Elseif blnRS Then
			Call doRedirect("pop_project.asp?m=" & bytMod & "&mid=" & lngModId)
		Else
			If mProjects Then strOwner = strFullName
			strCreatedBy = strFullName
			datCreatedDate = Now
			strModBy = strFullName
			datModDate = Now
		End If
	End If

	strIncHead = getCalendarScripts()

	Call DisplayHeader(3)
	Call showEditHeader(strTitle,strCreatedBy,datCreatedDate,strModBy,datModDate)
%>
<div id="contentDiv" class="dvBorder" style="height:380px;"><br>

<table border=0 width="100%">
<form name="frmProject" method="post" action="pop_project.asp?id=<% =lngRecordId %>&m=<% =bytMod %>&mid=<% =lngModId %>">
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
<% =getHidden("hdnWinOpen","") %>
    <tr>
      <td><% =getLabel(Application("IDS_Project"),"txtProject") %></td>
      <td><% =getTextField("txtProject","mText",strProject,40,40,"") %></td>
    </tr>
    <% If bytMod = 4 Then %>
    <tr>
      <td><% =getLabel(Application("IDS_Account"),"txtClient") %></td>
      <td><% =getTextField("txtClient","mText",strClient,40,40,"") %>
      <% If pClients >= 1 Then %>
      <a href="<% =newWindow("S","?m=2&rVal=C") %>"><img src="../images/import.gif" alt="<% =getImport("IDS_Account") %>" border=0 height=16 width=16></a></td>
      <% End If %>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_Division"),"txtDivision") %></td>
      <td><% =getTextField("txtDivision","oText",strDivision,40,40,"") %></td>
    </tr>
    <% End If %>
    <tr>
      <td><% =getLabel(Application("IDS_Owner"),"txtOwner") %></td>
      <td><% =getTextField("txtOwner","mText",strOwner,40,100,"") %>
      <a href="<% =newWindow("S","?m=0&rVal=txtOwner") %>"><img src="../images/import.gif" alt="<% =getImport("IDS_Owner") %>" border=0 height=16 width=16></a></td>
    </tr>
    <% If bytMod <> 7 Then %>
    <tr>
      <td><% =getLabel(Application("IDS_InvoiceId"),"txtInvoice") %></td>
      <td><% =getTextField("txtInvoice","oLong",lngInvoice,40,255,"") %>
      <% If pInvoices >= 1 Then %>
      <a href="<% =newWindow("S","?m=7&rVal=txtInvoice") %>"><img src="../images/import.gif" alt="<% =getImport("IDS_InvoiceId") %>" border=0 height=16 width=16></a></td>
      <% End If %>
    </tr>
	<% End If %>
    <tr>
      <td><% =getLabel(Application("IDS_Description"),"txtShortDesc") %></td>
      <td><% =getTextArea("txtShortDesc","oMemo",strShortDesc,"255",4,"") %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_DaysTotal"),"txtDaysTotal") %></td>
      <td><% =getTextField("txtDaysTotal","oInt",intDaysTotal,3,255,"") %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_DaysOwed"),"txtDaysOwed") %></td>
      <td><% =getTextField("txtDaysOwed","oInt",intDaysOwed,3,255,"") %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_Closed"),"chkClosed") %></td>
      <td><% =getCheckbox("chkClosed",blnClosed,"onClick=""doClassChange();""") %></td>
    </tr>
  <tr>
    <td><% =getLabel(Application("IDS_CloseDate"),"txtCloseDate") %></td>
    <td><% =getTextField("txtCloseDate","oDate",datCloseDate,12,12,"") %>
    <a href="Javascript:showCalendar('txtCloseDate');"><img src="../images/cal.gif" alt="<% =getImport("IDS_CloseDate") %>" border=0 height=16 width=16></a></td>
  </tr>
</form>
</table>

</div>

<div id="footerDiv" class="dvFooter">
<%
	If blnRS Then
		Response.Write(getIconNew("pop_project.asp"))
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
		} else {
			document.forms[0].txtCloseDate.className = "oDate";
		}
	}
</script>

<%
	Call DisplayFooter(3)
%>

