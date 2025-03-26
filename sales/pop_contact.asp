<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\states.asp" -->
<!--#include file="..\_inc\sql\sql_contacts.asp" -->
<%
	Call pageFunctions(1,2)

	Dim lngDivId			'as Long
	Dim strPrefix			'as String
	Dim intPrefixId			'as Integer
	Dim strFirstName		'as String
	Dim strMiddleInitial	'as String
	Dim strLastName			'as String
	Dim strAddress1			'as String
	Dim strAddress2			'as String
	Dim strAddress3			'as String
	Dim strCity				'as String
	Dim strState			'as String
	Dim strCountry			'as String
	Dim strZIP				'as String
	Dim strClient			'as String
	Dim strDivision			'as String
	Dim strEmail			'as String
	Dim strJobTitle			'as String
	Dim strDept				'as String
	Dim lngPhone1			'as Long
	Dim intExt1				'as Integer
	Dim lngPhone2			'as Long
	Dim intExt2				'as Integer
	Dim lngFax				'as Long
	Dim strReportsTo		'as String
	Dim lngReportsTo		'as Long
	Dim strAssistant		'as String
	Dim lngAssistant		'as Long
	Dim strCreatedBy		'as String
	Dim strModBy			'as String
	Dim datCreatedDate		'as Date
	Dim datModDate			'as Date

	strTitle = Application("IDS_Edit") & " " & Application("IDS_Contact")

	If strDoAction <> "" Then

		intPrefixId = valNum(Request.Form("selPrefix"),2,-1)
		strFirstName = valString(Request.Form("txtFirstName"),30,1,0)
		strMiddleInitial = valString(Request.Form("txtMiddleInitial"),1,0,0)
		strLastName = valString(Request.Form("txtLastName"),30,1,0)
		strAddress1 = valString(Request.Form("txtAddress1"),60,0,0)
		strAddress2 = valString(Request.Form("txtAddress2"),60,0,0)
		strAddress3 = valString(Request.Form("txtAddress3"),60,0,0)
		strCity = valString(Request.Form("txtCity"),20,0,0)
		strState = valString(Request.Form("selState"),2,0,0)
		strCountry = valString(Request.Form("selCountry"),20,0,0)
		strZip = valString(Request.Form("txtZIP"),7,0,0)
		strEmail = valString(Request.Form("txtEmail"),255,0,1)
		strDept = valString(Request.Form("txtDept"),40,0,0)
		strJobTitle = valString(Request.Form("txtJobTitle"),30,0,0)
		lngPhone1 = valNum(Request.Form("txtPhone1"),4,-1)
		intExt1 = valNum(Request.Form("txtExt1"),2,-1)
		lngPhone2 = valNum(Request.Form("txtPhone2"),4,-1)
		intExt2 = valNum(Request.Form("txtExt2"),2,-1)
		lngFax = valNum(Request.Form("txtFax"),4,-1)
		lngReportsTo = valNum(Request.Form("hdnReportsTo"),3,-1)
		lngAssistant = valNum(Request.Form("hdnAssistant"),3,-1)
		If bytMod <> 2 Then
			strClient = valString(Request.Form("txtClient"),40,1,0)
			strDivision = valString(Request.Form("txtDivision"),40,0,0)
		End If

		If strDoAction = "del" and intPerm >= 4 Then

			Call delContact(lngUserId,lngRecordId)

		ElseIf strDoAction = "edit" and intPerm >= 3 Then

			Call updateContact(lngUserId,lngRecordId,bytMod,lngModId,intPrefixId,strFirstName, _
					strMiddleInitial,strLastName,strAddress1,strAddress2,strAddress3,strCity,strState, _
					strCountry,strZIP,strEmail,strDept,strJobTitle,lngPhone1,intExt1,lngPhone2,intExt2, _
					lngFax,strClient,strDivision,lngReportsTo,lngAssistant)

		ElseIf strDoAction = "new" Then

			lngRecordId = insertContact(lngUserId,lngRecordId,bytMod,lngModId,intPrefixId,strFirstName, _
					strMiddleInitial,strLastName,strAddress1,strAddress2,strAddress3,strCity,strState, _
					strCountry,strZIP,strEmail,strDept,strJobTitle,lngPhone1,intExt1,lngPhone2,intExt2, _
					lngFax,strClient,strDivision,lngReportsTo,lngAssistant)
		End If

		Select Case bytMod
			Case 6,7
				lngDivId = getValue("DivId","CRM_Contacts","ContactId = " & lngRecordId,0)
				Response.Write("<html><head><script>" & _
						"window.opener.document.forms[0].txtContact.value = '" & strFirstName & " " & strLastName & "';" & _
						"window.opener.document.forms[0].hdnContact.value = '" & lngRecordId & "';" & _
						"window.opener.document.forms[0].hdnDivision.value = '" & lngDivId & "';" & _
						"window.close();" & _
						"</script></head></html>")
				Call endResponse()
			Case Else
				Call closeWindow(strOpenerURL)
		End Select
	Else
		If blnRS and intPerm >= 3 Then

			Set objRS = objConn.Execute(getContact(0,lngRecordId))

			If not (objRS.BOF and objRS.EOF) then
				intPrefixId = showString(objRS.fields("K_Prefix").value)
				strFirstName = showString(objRS.fields("K_FirstName").value)
				strMiddleInitial = showString(objRS.fields("K_MiddleInitial").value)
				strLastName = showString(objRS.fields("K_LastName").value)
				strAddress1 = showString(objRS.fields("K_Address1").value)
				strAddress2 = showString(objRS.fields("K_Address2").value)
				strAddress3 = showString(objRS.fields("K_Address3").value)
				strCity = showString(objRS.fields("K_City").value)
				strState = showString(objRS.fields("K_State").value)
				strCountry = showString(objRS.fields("K_Country").value)
				strZIP = showString(objRS.fields("K_ZIP").value)
				strClient = showString(objRS.fields("C_Client").value)
				strDivision = showString(objRS.fields("D_Division").value)
				strDept = showString(objRS.fields("K_Dept").value)
				strJobTitle = showString(objRS.fields("K_JobTitle").value)
				strEmail = showString(objRS.fields("K_Email").value)
				lngPhone1 = showPhone(objRS.fields("K_Phone1").value)
				intExt1 = showString(objRS.fields("K_Ext1").value)
				lngPhone2 = showPhone(objRS.fields("K_Phone2").value)
				intExt2 = showString(objRS.fields("K_Ext2").value)
				lngFax = showPhone(objRS.fields("K_Fax").value)
			'	strReportsTo = showString(objRS.fields("ReportsTo").value)
			'	strAssistant = showString(objRS.fields("Assistant").value)
				strCreatedBy = showString(objRS.fields("CreatedBy").value)
				datCreatedDate = objRS.fields("K_CreatedDate").value
				strModBy = showString(objRS.fields("ModBy").value)
				datModDate = objRS.fields("K_ModDate").value
			End If
		Elseif blnRS Then
			Call doRedirect("pop_contact.asp?m=" & bytMod & "&mid=" & lngModId)
		Else
			If bytMod = 2 Then
				strClient = showString(getClientName(0,lngModId))
				strDivision = showString(getDivName(lngModId))
			End If
			strCreatedBy = strFullName
			datCreatedDate = Now
			strModBy = strFullName
			datModDate = Now
		End If
		strPrefix = getOptionDropDown(45,True,"selPrefix","Prefix",intPrefixId)
	End If

	Call DisplayHeader(3)
	Call showEditHeader(strTitle,strCreatedBy,datCreatedDate,strModBy,datModDate)
%>
<div id="contentDiv" class="dvBorder"><br>

<table border=0 width="100%">
<form name="frmContact" method="post" action="pop_contact.asp?id=<% =lngRecordId %>&m=<% =bytMod %>&mid=<% =lngModId %>">
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
<% =getHidden("hdnWinOpen","") %>
    <tr>
      <td><% =getLabel(Application("IDS_Contact"),"txtFirstName") %></td>
      <td>
      <% =strPrefix %>
      <% =getTextField("txtFirstName","mText",strFirstName,11,30,"") %>
      <% =getTextField("txtMiddleInitial","oText",strMiddleInitial,1,1,"") %>
      <% =getTextField("txtLastName","mText",strLastName,14,30,"") %>
      </td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_Address"),"txtAddress1") %></td>
      <td><% =getTextField("txtAddress1","oText",strAddress1,40,60,"") %></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><% =getTextField("txtAddress2","oText",strAddress2,40,60,"") %></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><% =getTextField("txtAddress3","oText",strAddress3,40,60,"") %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_City"),"txtCity") %></td>
      <td><% =getTextField("txtCity","oText",strCity,40,20,"") %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_State"),"selState") %></td>
      <td><% =getStates(140,"selState",strState) %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_Country"),"selCountry") %></td>
      <td><% =getCountries(140,"selCountrry",strCountry) %>
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<% =getLabel(Application("IDS_ZIP"),"txtZIP") %>
      <% =getTextField("txtZIP","oText",strZIP,7,7,"") %></td>
    </tr>
    <% If bytMod <> 2 Then %>
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
    <% Else %>
    <tr>
      <td><% =getLabel(Application("IDS_Account"),"txtClient") %></td>
      <td><% =getTextField("txtClient","dText",strClient,40,40,"readonly=""readonly""") %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_Division"),"txtDivision") %></td>
      <td><% =getTextField("txtDivision","dText",strDivision,40,40,"readonly=""readonly""") %></td>
    </tr>
    <% End If %>
    <tr>
      <td><% =getLabel(Application("IDS_Department"),"txtDept") %></td>
      <td><% =getTextField("txtDept","oText",strDept,40,40,"") %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_JobTitle"),"txtJobTitle") %></td>
      <td><% =getTextField("txtJobTitle","oText",strJobTitle,40,30,"") %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_Email"),"txtEmail") %></td>
      <td><% =getTextField("txtEmail","oEmail",strEmail,40,255,"") %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_Phone") & " 1","txtPhone1") %></td>
      <td><% =getTextField("txtPhone1","oPhone",lngPhone1,15,255,"") %>&nbsp;&nbsp;
      <% =getLabel("Ext.","txtExt1") %>
      <% =getTextField("txtExt1","oInt",intExt1,6,255,"") %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_Phone") & " 2","txtPhone2") %></td>
      <td><% =getTextField("txtPhone2","oPhone",lngPhone2,15,255,"") %>&nbsp;&nbsp;
      <% =getLabel("Ext.","txtExt2") %>
      <% =getTextField("txtExt2","oInt",intExt2,6,255,"") %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_Fax"),"txtFax") %></td>
      <td><% =getTextField("txtFax","oPhone",lngFax,15,255,"") %></td>
    </tr>
</form>
</table>

</div>

<div id="footerDiv" class="dvFooter">
<%
	If blnRS Then
		Response.Write(getIconNew("pop_contact.asp"))
		If intPerm >= 4 Then Response.Write(getIconDelete())
	End If
	Response.Write(getIconSave(strAction))
	Response.Write(getIconCancel())
%>
</div>

<%
	Call DisplayFooter(3)
%>

