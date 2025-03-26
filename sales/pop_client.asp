<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_clients.asp" -->
<%
	Call pageFunctions(2,2)

	Dim strClient		'as String
	Dim strDivision		'as String
	Dim strAccount		'as String
	Dim blnRefAccount	'as Boolean
	Dim strRegion		'as String
	Dim intRegionId		'as Integer
	Dim strWebsite		'as String
	Dim strSalesRep		'as String
	Dim lngSalesRepId	'as Long
	Dim strAccountType	'as String
	Dim intAccountTypeID'as Integer
	Dim intVerticalId	'as Integer
	Dim strVertical		'as String
	Dim blnProbFlag		'as Boolean
	Dim strShortDesc	'as String
	Dim strCreatedBy	'as String
	Dim datCreatedDate	'as Date
	Dim strModBy		'as String
	Dim datModDate		'as Date

	strTitle = Application("IDS_Edit") & " " & Application("IDS_Account")

	If strDoAction <> "" Then

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
		blnProbFlag = valNum(Request.Form("chkProbFlag"),0,0)
		strShortDesc = valString(Request.Form("txtShortDesc"),255,0,4)

		If strDoAction = "del" and intPerm >= 4 Then

			Call delClient(lngUserId,lngRecordId)

		ElseIf strDoAction = "edit" and intPerm >= 3 Then

			Call updateClient(lngUserId,lngRecordId,blnRefAccount,strClient,strDivision,strAccount, _
					intAccountTypeID,lngSalesRepId,intRegionId,strWebsite,intVerticalId,blnProbFlag, _
					strShortDesc)

		ElseIf strDoAction = "new" Then

			lngRecordId = insertClient(lngUserId,lngRecordId,blnRefAccount,strClient,strDivision,strAccount, _
					intAccountTypeID,lngSalesRepId,intRegionId,strWebsite,intVerticalId,blnProbFlag, _
					strShortDesc)
		End If
		Call closeWindow(strOpenerURL)
	Else

		If blnRS and intPerm >= 3 Then
			Set objRS = objConn.Execute(getClient(0,lngRecordId))

			If not (objRS.BOF and objRS.EOF) then
				strClient = showString(objRS.fields("C_Client").value)
				strDivision = showString(objRS.fields("D_Division").value)
				strSalesRep = showString(objRS.fields("SalesRep").value)
				intRegionId = objRS.fields("D_Region").value
				strWebsite = showString(objRS.fields("D_Website").value)
				strAccount = showString(objRS.fields("D_Account").value)
				intAccountTypeID = objRS.fields("D_AccountType").value
				blnRefAccount = objRS.fields("D_RefAccount").value
				intVerticalId = objRS.fields("D_Vertical").value
				blnProbFlag = objRS.fields("D_ProbFlag").value
				strShortDesc = showString(objRS.fields("D_ShortDesc").value)
				strCreatedBy = showString(objRS.fields("CreatedBy").value)
				datCreatedDate = objRS.fields("D_CreatedDate").value
				strModBy = showString(objRS.fields("ModBy").value)
				datModDate = objRS.fields("D_ModDate").value
			End If
		Elseif blnRS Then
			Call doRedirect("pop_client.asp")
		Else
			If mSales Then strSalesRep = strFullName
			strCreatedBy = strFullName
			datCreatedDate = Now
			strModBy = strFullName
			datModDate = Now
		End If
		strAccountType = getOptionDropDown(255,False,"selAccountType","Account Type",intAccountTypeID)
		strRegion = getOptionDropDown(255,False,"selRegion","Sales Region",intRegionId)
		strVertical = getOptionDropDown(255,True,"selVertical","Sales Vertical",intVerticalId)
	End If

	Call DisplayHeader(3)
	Call showEditHeader(strTitle,strCreatedBy,datCreatedDate,strModBy,datModDate)
%>
<div id="contentDiv" class="dvBorder" style="height:380px;"><br>

<table border=0 width="100%">
<form name="frmClient" method="post" action="pop_client.asp?id=<% =lngRecordId %>">
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
<% =getHidden("hdnWinOpen","") %>
    <tr>
      <td><% =getLabel(Application("IDS_Account"),"txtClient") %></td>
      <td><% =getTextField("txtClient","mText",strClient,40,40,"") %>
      <a href="<% =newWindow("S","?m=2&rVal=C") %>"><img src="../images/import.gif" alt="<% =getImport("IDS_Account") %>" border=0 height=16 width=16></a></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_Division"),"txtDivision") %></td>
      <td><% =getTextField("txtDivision","oText",strDivision,40,40,"") %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_AccountId"),"txtAccount") %></td>
      <td><% =getTextField("txtAccount","oText",strAccount,40,25,"") %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_SalesRep"),"txtSalesRep") %></td>
      <td><% =getTextField("txtSalesRep","mText",strSalesRep,40,100,"") %>
      <a href="<% =newWindow("S","?m=0&rVal=txtSalesRep") %>"><img src="../images/import.gif" alt="<% =getImport("IDS_SalesRep") %>" border=0 height=16 width=16></a></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_Region"),"selRegion") %></td>
      <td><% =strRegion %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_Website"),"txtWebsite") %></td>
      <td><% =getTextField("txtWebsite","oLink",strWebsite,40,255,"") %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_AccountType"),"selAccountType") %></td>
      <td><% =strAccountType %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_IndustrySector"),"selVertical") %></td>
      <td><% =strVertical %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_Description"),"txtShortDesc") %></td>
      <td><% =getTextArea("txtShortDesc","oMemo",strShortDesc,"255",4,"") %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_Reference"),"chkRefAccount") %></td>
      <td><% =getCheckbox("chkRefAccount",blnRefAccount,"") %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_ProblemFlag"),"chkProbFlag") %></td>
      <td><% =getCheckbox("chkProbFlag",blnProbFlag,"") %></td>
    </tr>
</form>
</table>

</div>

<div id="footerDiv" class="dvFooter">
<%
	If blnRS Then
		Response.Write(getIconNew("pop_client.asp"))
		If intPerm >= 4 Then Response.Write(getIconDelete())
	End If
	Response.Write(getIconSave(strAction))
	Response.Write(getIconCancel())
%>
</div>

<%
	Call DisplayFooter(3)
%>