<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_clients.asp" -->
<%
	Call pageFunctions(2,2)

	Dim strClient       'as String
	Dim strDivision     'as String
	Dim strAccount      'as String
	Dim blnRefAccount   'as Boolean
	Dim strRegion       'as String
	Dim intRegionId     'as Integer
	Dim strWebsite      'as String
	Dim strSalesRep     'as String
	Dim lngSalesRepId   'as Long
	Dim strAccountType  'as String
	Dim intAccountTypeID'as Integer
	Dim intVerticalId   'as Integer
	Dim strVertical     'as String
	Dim intSizeId       'as Integer
	Dim strSize         'as String
	Dim blnProbFlag     'as Boolean
	Dim strShortDesc    'as String
	Dim strCreatedBy    'as String
	Dim datCreatedDate  'as Date
	Dim strModBy        'as String
	Dim datModDate      'as Date
	Dim strExtraFields  'as String

	strTitle = getIDS("IDS_Edit") & " " & getIDS("IDS_Account")

	If strDoAction <> "" Then

		Select Case strDoAction
			Case "del"
				If intPerm >= 4 Then Call delClient(lngUserId,lngRecordId)

			Case "new","edit"

				strSalesRep = valString(Request.Form("txtSalesRep"),100,1,0)
				lngSalesRepId = getUserId(3,strSalesRep)
				blnRefAccount = valNum(Request.Form("chkRefAccount"),0,0)
				strClient = valString(Request.Form("txtClient"),40,1,0)
				strDivision = valString(Request.Form("txtDivision"),40,0,0)
				strAccount = valString(Request.Form("txtAccount"),25,0,0)
				intAccountTypeID = valNum(Request.Form("selAccountType"),2,-1)
				intRegionId = valNum(Request.Form("selRegion"),2,-1)
				strWebsite = valString(Request.Form("txtWebsite"),255,0,0)
				intVerticalId = valNum(Request.Form("selVertical"),2,-1)
				intSizeId = valNum(Request.Form("selSize"),2,-1)
				blnProbFlag = valNum(Request.Form("chkProbFlag"),0,0)
				strShortDesc = valString(Request.Form("txtShortDesc"),255,0,4)

				If strDoAction = "edit" and intPerm >= 3 Then

					Call updateClient(lngUserId,lngRecordId,blnRefAccount,strClient,strDivision,strAccount, _
							intAccountTypeID,lngSalesRepId,intRegionId,strWebsite,intVerticalId,blnProbFlag, _
							intSizeId,strShortDesc)

				ElseIf strDoAction = "new" Then

					lngRecordId = insertClient(lngUserId,lngRecordId,blnRefAccount,strClient,strDivision,strAccount, _
							intAccountTypeID,lngSalesRepId,intRegionId,strWebsite,intVerticalId,blnProbFlag, _
							intSizeId,strShortDesc)
				End If

				Call saveCustomFields(2,lngRecordId)
		End Select
		Call closeEdit()
	Else

		If blnRS and intPerm >= 3 Then
			Set objRS = objConn.Execute(getClient(0,lngRecordId))

			If not (objRS.BOF and objRS.EOF) then
				strClient = objRS.fields("C_Client").value
				strDivision = objRS.fields("D_Division").value
				strSalesRep = objRS.fields("SalesRep").value
				intRegionId = objRS.fields("D_Region").value
				strWebsite = objRS.fields("D_Website").value
				strAccount = objRS.fields("D_Account").value
				intAccountTypeID = objRS.fields("D_AccountType").value
				blnRefAccount = objRS.fields("D_RefAccount").value
				intVerticalId = objRS.fields("D_Vertical").value
				intSizeId = objRS.fields("D_Size").value
				blnProbFlag = objRS.fields("D_ProbFlag").value
				strShortDesc = objRS.fields("D_ShortDesc").value
				strCreatedBy = objRS.fields("CreatedBy").value
				datCreatedDate = objRS.fields("D_CreatedDate").value
				strModBy = objRS.fields("ModBy").value
				datModDate = objRS.fields("D_ModDate").value
			End If
		Elseif blnRS Then
			Call logError(2,1)
		Else
			If mSales Then strSalesRep = strFullName
			strCreatedBy = strFullName
			datCreatedDate = Now
			strModBy = strFullName
			datModDate = Now
		End If
		strExtraFields = editCustomFields(2)
		strAccountType = getOptionDropDown(255,False,"selAccountType","IDS_AccountType",intAccountTypeID)
		strRegion = getOptionDropDown(255,False,"selRegion","IDS_SalesRegion",intRegionId)
		strVertical = getOptionDropDown(255,True,"selVertical","IDS_IndustrySector",intVerticalId)
		strSize = getOptionDropDown(255,True,"selSize","IDS_AccountSize",intVerticalId)
	End If

	strIncHead = getCalendarScripts()

	Call DisplayHeader(1)
	Call showEditHeader(strTitle,strCreatedBy,datCreatedDate,strModBy,datModDate)
%>
<form name="frmClient" method="post" action="edit_client.asp?id=<% =lngRecordId %>">
<div id="contentDiv" class="dvBorder" style="height:<% =intScreenH-170 %>px;"><br />

<table border=0 width="100%">
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
	<tr>
	  <td width=170><% =getLabel(getIDS("IDS_Account"),"txtClient") %></td>
	  <td><% =getTextField("txtClient","mText",strClient,40,40,"") %>
	  <% =getIconImport(1,getSearchURL("?m=2&rVal=C"),getIDS("IDS_Account")) %>
	  </td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_Division"),"txtDivision") %></td>
	  <td><% =getTextField("txtDivision","oText",strDivision,40,40,"") %></td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_AccountId"),"txtAccount") %></td>
	  <td><% =getTextField("txtAccount","oText",strAccount,40,25,"") %></td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_SalesRep"),"txtSalesRep") %></td>
	  <td><% =getTextField("txtSalesRep","mText",strSalesRep,40,100,"") %>
	  <% =getIconImport(1,getSearchURL("?m=0&rVal=txtSalesRep"),getIDS("IDS_SalesRep")) %>
	  </td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_SalesRegion"),"selRegion") %></td>
	  <td><% =strRegion %></td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_Website"),"txtWebsite") %></td>
	  <td><% =getTextField("txtWebsite","oLink",strWebsite,40,255,"") %></td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_AccountType"),"selAccountType") %></td>
	  <td><% =strAccountType %></td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_IndustrySector"),"selVertical") %></td>
	  <td><% =strVertical %></td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_AccountSize"),"selSize") %></td>
	  <td><% =strSize %></td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_Description"),"txtShortDesc") %></td>
	  <td><% =getTextArea("txtShortDesc","oMemo",strShortDesc,"255",4,"") %></td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_Reference"),"chkRefAccount") %></td>
	  <td><% =getCheckbox("chkRefAccount",blnRefAccount,"") %></td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_ProblemFlag"),"chkProbFlag") %></td>
	  <td><% =getCheckbox("chkProbFlag",blnProbFlag,"") %></td>
	</tr>
<%	=strExtraFields %>
</table>

</div>

<div id="footerDiv" class="dvFooter">
<%
	If blnRS Then
		Response.Write(getIconNew(getEditURL(2,"")))
		If intPerm >= 4 Then Response.Write(getIconDelete())
	End If
	Response.Write(getIconSave(strAction))
	Response.Write(getIconCancel("back"))
%>
</div>
</form>

<%
	Call DisplayFooter(1)
%>