<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_services.asp" -->
<%
	Call pageFunctions(4,2)

	Dim strProject      'as String
	Dim strClient       'as String
	Dim strDivision     'as String
	Dim lngSaleId       'as Long
	Dim lngInvoice      'as String
	Dim strOwner        'as String
	Dim lngOwnerId      'as Long
	Dim strFrame        'as String
	Dim intDaysTotal    'as Integer
	Dim intDaysOwed     'as Integer
	Dim strShortDesc    'as String
	Dim lngDivId        'as Long
	Dim blnClosed       'as Boolean
	Dim datCloseDate    'as Date
	Dim strCreatedBy    'as String
	Dim datCreatedDate  'as Date
	Dim strModBy        'as String
	Dim datModDate      'as Date
	Dim strExtraFields  'as String

	strTitle = getIDS("IDS_Edit") & " " & getIDS("IDS_Project")

	If strDoAction <> "" Then

		Select Case strDoAction
			Case "del"
				If intPerm >= 4 Then Call delProject(lngUserId,lngRecordId)

			Case "new","edit"
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

				If strDoAction = "edit" and intPerm >= 3 Then

					Call updateProject(lngUserId,lngRecordId,bytMod,lngModId,strProject,lngOwnerId,datCloseDate, _
							lngInvoice,intDaysTotal,intDaysOwed,blnClosed,strClient,strDivision,strShortDesc)

				ElseIf strDoAction = "new" Then

					lngRecordId = insertProject(lngUserId,lngRecordId,bytMod,lngModId,strProject,lngOwnerId,datCloseDate, _
							lngInvoice,intDaysTotal,intDaysOwed,blnClosed,strClient,strDivision,strShortDesc)
				End If

				Call saveCustomFields(4,lngRecordId)
		End Select

		If bytMenu = 0 Then Session("LastPage") = strCRMURL & "common/i_projects.asp?m=" & bytMod & "&mid=" & lngModId
		Call closeEdit()
	Else

		If blnRS and intPerm >= 3 Then
			Set objRS = objConn.Execute(getProject(0,lngRecordId))

			If not (objRS.BOF and objRS.EOF) then
				lngRecordId = objRS.fields("ProjectId").value
				strProject = objRS.fields("P_Title").value
				strClient = objRS.fields("C_Client").value
				strDivision = objRS.fields("D_Division").value
				lngDivId = objRS.fields("DivId").value
				lngInvoice = objRS.fields("InvoiceId").value
				strOwner = objRS.fields("Owner").value
				intDaysTotal = objRS.fields("P_DaysTotal").value
				intDaysOwed = objRS.fields("P_DaysOwed").value
				strShortDesc = objRS.fields("P_ShortDesc").value
				blnClosed = objRS.fields("P_Closed").value
				datCloseDate = objRS.fields("P_CloseDate").value
				strCreatedBy = objRS.fields("CreatedBy").value
				datCreatedDate = objRS.fields("P_CreatedDate").value
				strModBy = objRS.fields("ModBy").value
				datModDate = objRS.fields("P_ModDate").value
			End If
		Elseif blnRS Then
			Call logError(2,1)
		Else
			If mProjects Then strOwner = strFullName
			strCreatedBy = strFullName
			datCreatedDate = Now
			strModBy = strFullName
			datModDate = Now
		End If
		strExtraFields = editCustomFields(4)
	End If

	strIncHead = getCalendarScripts()

	Call DisplayHeader(0)
	Call showEditHeader(strTitle,strCreatedBy,datCreatedDate,strModBy,datModDate)
%>
<form name="frmProject" method="post" action="edit_project.asp?id=<% =lngRecordId %>&m=<% =bytMod %>&mid=<% =lngModId %>&menu=<% =bytMenu %>">
<div id="contentDiv" class="dvBorder" style="height:<% =intScreenH-170 %>px;"><br />

<table border=0 width="100%">
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
	<tr>
	  <td width=170><% =getLabel(getIDS("IDS_Project"),"txtProject") %></td>
	  <td><% =getTextField("txtProject","mText",strProject,40,40,"") %></td>
	</tr>
	<% If bytMod = 4 Then %>
	<tr>
	  <td><% =getLabel(getIDS("IDS_Account"),"txtClient") %></td>
	  <td><% =getTextField("txtClient","mText",strClient,40,40,"") %>
<%
	If pClients >= 1 Then
		Response.Write(getIconImport(1,getSearchURL("?m=2&rVal=C"),getIDS("IDS_Account")))
	End If
%>
	  </td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_Division"),"txtDivision") %></td>
	  <td><% =getTextField("txtDivision","oText",strDivision,40,40,"") %></td>
	</tr>
	<% End If %>
	<tr>
	  <td><% =getLabel(getIDS("IDS_Owner"),"txtOwner") %></td>
	  <td><% =getTextField("txtOwner","mText",strOwner,40,100,"") %>
	  <% =getIconImport(1,getSearchURL("?m=0&rVal=txtOwner"),getIDS("IDS_Owner")) %>
	  </td>
	</tr>
	<% If bytMod <> 7 Then %>
	<tr>
	  <td><% =getLabel(getIDS("IDS_InvoiceId"),"txtInvoice") %></td>
	  <td><% =getTextField("txtInvoice","oLong",lngInvoice,40,255,"") %>
<%
	If pInvoices >= 1 Then
		Response.Write(getIconImport(1,getSearchURL("?m=7&rVal=txtInvoice"),getIDS("IDS_Invoice")))
	End If
%>
	  </td>
	</tr>
	<% End If %>
	<tr>
	  <td><% =getLabel(getIDS("IDS_Description"),"txtShortDesc") %></td>
	  <td><% =getTextArea("txtShortDesc","oMemo",strShortDesc,"255",4,"") %></td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_DaysTotal"),"txtDaysTotal") %></td>
	  <td><% =getTextField("txtDaysTotal","oInt",intDaysTotal,3,255,"") %></td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_DaysOwed"),"txtDaysOwed") %></td>
	  <td><% =getTextField("txtDaysOwed","oInt",intDaysOwed,3,255,"") %></td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_Closed"),"chkClosed") %></td>
	  <td><% =getCheckbox("chkClosed",blnClosed,"onClick=""doClassChange();""") %></td>
	</tr>
    <tr>
	  <td><% =getLabel(getIDS("IDS_CloseDate"),"txtCloseDate") %></td>
	  <td><% =getDateField("txtCloseDate","oDate",datCloseDate,getIDS("IDS_CloseDate")) %></td>
    </tr>
<%	=strExtraFields %>
</table>

</div>

<div id="footerDiv" class="dvFooter">
<%
	If blnRS Then
		Response.Write(getIconNew(getEditURL(4,"?m="&bytMod&"&mid="&lngModId)))
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
		} else {
			document.forms[0].txtCloseDate.className = "oDate";
		}
	}
</script>

<%
	Call DisplayFooter(0)
%>

