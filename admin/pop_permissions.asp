<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_users.asp" -->
<%
	Call pageFunctions(90,5)

	Dim strUsrMem		'as String
	Dim strUsrPerm		'as String
	Dim strName 		'as String
	Dim bytAdmin		'as Byte
	Dim bytPass			'as Byte
	Dim bytLock			'as Byte
	Dim blnContact, blnClient, blnSale, blnServices, blnSupport, blnQA, blnKB, blnFinance 	'as Boolean
	Dim bytContact, bytClient, bytSale, bytServices, bytSupport, bytQA, bytKB, bytFinance	'as Byte

	strTitle = Application("IDS_Permissions")

	If lngRecordId = "" Then Call logError(3,1)

	If strDoAction = "edit" Then

		strUsrMem = valNum(Request.Form("chkContact"),0,0)
		strUsrMem = strUsrMem & valNum(Request.Form("chkClient"),0,0)
		strUsrMem = strUsrMem & valNum(Request.Form("chkSale"),0,0)
		strUsrMem = strUsrMem & valNum(Request.Form("chkService"),0,0)
		strUsrMem = strUsrMem & valNum(Request.Form("chkSupport"),0,0)
		strUsrMem = strUsrMem & valNum(Request.Form("chkQA"),0,0)
		strUsrMem = strUsrMem & valNum(Request.Form("chkFinance"),0,0)
		strUsrMem = strUsrMem & valNum(Request.Form("chkKB"),0,0)

		strUsrPerm = valNum(Request.Form("selContact"),1,1)
		strUsrPerm = strUsrPerm & valNum(Request.Form("selClient"),1,1)
		strUsrPerm = strUsrPerm & valNum(Request.Form("selSale"),1,1)
		strUsrPerm = strUsrPerm & valNum(Request.Form("selService"),1,1)
		strUsrPerm = strUsrPerm & valNum(Request.Form("selSupport"),1,1)
		strUsrPerm = strUsrPerm & valNum(Request.Form("selQA"),1,1)
		strUsrPerm = strUsrPerm & valNum(Request.Form("selFinance"),1,1)
		strUsrPerm = strUsrPerm & valNum(Request.Form("selKB"),1,1)

		If blnAdmin Then
			bytAdmin = valNum(Request.Form("chkAdmin"),0,0)
			bytPass	 = valNum(Request.Form("chkPass"),0,0)
			If valNum(Request.Form("chkLock"),0,0) = 1 Then bytLock = CInt(Application("av_LoginAttempts")) Else bytLock = 0
		Else
			bytAdmin = "NULL"
		End If

		objConn.Execute(upDateUserPerm(lngRecordId,strUsrMem,strUsrPerm,bytAdmin,bytPass,bytLock))
		Call closeWindow(strOpenerURL)
	Else
		Set objRS = objConn.Execute(getUserDetails(lngRecordId))

		If not (objRS.BOF and objRS.EOF) then
			bytContact 	= valNum(Mid(objRS.fields("U_Permissions").value,1,1),1,0)
			bytClient 	= valNum(Mid(objRS.fields("U_Permissions").value,2,1),1,0)
			bytSale	 	= valNum(Mid(objRS.fields("U_Permissions").value,3,1),1,0)
			bytServices = valNum(Mid(objRS.fields("U_Permissions").value,4,1),1,0)
			bytSupport 	= valNum(Mid(objRS.fields("U_Permissions").value,5,1),1,0)
			bytQA	 	= valNum(Mid(objRS.fields("U_Permissions").value,6,1),1,0)
			bytFinance	= valNum(Mid(objRS.fields("U_Permissions").value,7,1),1,0)
			bytKB	 	= valNum(Mid(objRS.fields("U_Permissions").value,8,1),1,0)

			blnContact	= valNum(Mid(objRS.fields("U_Member").value,1,1),0,0)
			blnClient	= valNum(Mid(objRS.fields("U_Member").value,2,1),0,0)
			blnSale		= valNum(Mid(objRS.fields("U_Member").value,3,1),0,0)
			blnServices = valNum(Mid(objRS.fields("U_Member").value,4,1),0,0)
			blnServices = valNum(Mid(objRS.fields("U_Member").value,5,1),0,0)
			blnQA	 	= valNum(Mid(objRS.fields("U_Member").value,6,1),0,0)
			blnFinance 	= valNum(Mid(objRS.fields("U_Member").value,7,1),0,0)
			blnKB	 	= valNum(Mid(objRS.fields("U_Member").value,8,1),0,0)

			If objRS.fields("U_Admin").value = 1 Then bytAdmin = 1 Else bytAdmin = 0
			If objRS.fields("U_ChangePassword").value = 1 Then bytPass = 1 Else bytPass = 0
			If objRS.fields("U_LoginAttempts").value >= CInt(Application("av_LoginAttempts")) Then bytLock = 1 Else bytLock = 0

			strName = showString(objRS.fields(0).value)
		End If
	End If

	Function showPermissions(fTitle,fDefault)

		showPermissions = "<select name=""" & fTitle & """ id=""" & fTitle & """ class=""oByte"" onChange=""doChange();"">" & vbCrLf & _
				vbTab & "<option value=0" & getDefault(0,fDefault,0) & ">" & Application("IDS_AccessDeny") & "</option>" & vbCrLf & _
				vbTab & "<option value=1" & getDefault(0,fDefault,1) & ">" & Application("IDS_AccessRead") & "</option>" & vbCrLf & _
				vbTab & "<option value=2" & getDefault(0,fDefault,2) & ">" & Application("IDS_New") & "</option>" & vbCrLf & _
				vbTab & "<option value=3" & getDefault(0,fDefault,3) & ">" & Application("IDS_Edit") & "</option>" & vbCrLf & _
				vbTab & "<option value=4" & getDefault(0,fDefault,4) & ">" & Application("IDS_Delete") & "</option>" & vbCrLf & _
				vbTab & "<option value=5" & getDefault(0,fDefault,5) & ">" & Application("IDS_Administrator") & "</option>" & vbCrLf & _
				"    </select>"
	End Function

	strTitle = Application("IDS_Permissions") & " - " & strName
	Call DisplayHeader(3)
	Call showEditHeader(strTitle,"","","","")
%>
<div id="contentDiv" class="dvBorder" style="height:378px;"><br>

<table border=0 cellspacing=0 cellpadding=3 width="99%">
<form name="frmAdmin" method="post" action="pop_permissions.asp?id=<% =lngRecordId %>">
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
<% =getHidden("hdnWinOpen","") %>
  <% If blnAdmin Then %>
  <tr>
    <td colspan=2><% =getLabel(Application("IDS_CRMAdministrator"),"chkAdmin") %></td>
    <td><% =getCheckbox("chkAdmin",bytAdmin,"") %></td>
  </tr>
  <tr>
    <td colspan=2><% =getLabel(Application("IDS_PasswordChange"),"chkPass") %></td>
    <td><% =getCheckbox("chkPass",bytPass,"") %></td>
  </tr>
  <tr>
    <td colspan=2><% =getLabel(Application("IDS_LockAccount"),"chkLock") %><br><br></td>
    <td><% =getCheckbox("chkLock",bytLock,"") %><br><br></td>
  </tr>
  <% End If %>
  <tr class="hrow">
    <th align=left class="dlabel"><% =Application("IDS_Group") %></td>
    <th align=center class="dlabel"><% =Application("IDS_Member") %></td>
    <th align=right class="dlabel"><% =Application("IDS_Permissions") %></td>
  </tr>
  <% If Application("av_Module1On") and (pContacts = 5 or blnAdmin) Then %>
  <tr>
    <td><% =getLabel(Application("IDS_Contacts"),"chkContact") %></td>
    <td align=center><% =getCheckbox("chkContact",blnContact,"") %></td>
    <td align=right><% =showPermissions("selContact",bytContact) %></td>
  </tr>
  <% Else %>
  <% =getHidden("chkContact",blnContact) %>
  <% =getHidden("selContact",bytContact) %>
  <% End If
     If Application("av_Module2On") and (pClients = 5 or blnAdmin) Then %>
  <tr>
    <td class="dfont"><% =getLabel(Application("IDS_Accounts"),"chkClient") %></td>
    <td align=center><% =getCheckbox("chkClient",blnClient,"") %></td>
    <td align=right><% =showPermissions("selClient",bytClient) %></td>
  </tr>
  <% Else %>
  <% =getHidden("chkClient",blnClient) %>
  <% =getHidden("selClient",bytClient) %>
  <% End If
     If Application("av_Module3On") and (pSales = 5 or blnAdmin) Then %>
  <tr>
    <td class="dfont"><% =getLabel(Application("IDS_Sales"),"chkSale") %></td>
    <td align=center><% =getCheckbox("chkSale",blnSale,"") %></td>
    <td align=right><% =showPermissions("selSale",bytSale) %></td>
  </tr>
  <% Else %>
  <% =getHidden("chkSale",blnSale) %>
  <% =getHidden("selSale",bytSale) %>
  <% End If
     If Application("av_Module4On") and (pProjects = 5 or blnAdmin) Then %>
  <tr>
    <td class="dfont"><% =getLabel(Application("IDS_ModServices"),"chkService") %></td>
    <td align=center><% =getCheckbox("chkService",blnServices,"") %></td>
    <td align=right><% =showPermissions("selService",bytServices) %></td>
  </tr>
  <% Else %>
  <% =getHidden("chkService",blnServices) %>
  <% =getHidden("selService",bytServices) %>
  <% End If
  If Application("av_Module5On") and (pTickets = 5 or blnAdmin) Then %>
  <tr>
    <td class="dfont"><% =getLabel(Application("IDS_ModSupport"),"chkSupport") %></td>
    <td align=center><% =getCheckbox("chkSupport",blnServices,"") %></td>
    <td align=right><% =showPermissions("selSupport",bytSupport) %></td>
  </tr>
  <% Else %>
  <% =getHidden("chkSupport",blnServices) %>
  <% =getHidden("selSupport",bytSupport) %>
  <% End If
  If Application("av_Module6On") and (pBugs = 5 or blnAdmin) Then %>
  <tr>
    <td class="dfont"><% =getLabel(Application("IDS_ModQualityAssurance"),"chkQA") %></td>
    <td align=center><% =getCheckbox("chkQA",blnQA,"") %></td>
    <td align=right><% =showPermissions("selQA",bytQA) %></td>
  </tr>
  <% Else %>
  <% =getHidden("chkQA",blnQA) %>
  <% =getHidden("selQA",bytQA) %>
  <% End If
  If Application("av_Module7On") and (pInvoices = 5 or blnAdmin) Then %>
  <tr>
    <td class="dfont"><% =getLabel(Application("IDS_ModFinance"),"chkFinance") %></td>
    <td align=center><% =getCheckbox("chkFinance",blnFinance,"") %></td>
    <td align=right><% =showPermissions("selFinance",bytFinance) %></td>
  </tr>
  <% Else %>
  <% =getHidden("chkFinance",blnFinance) %>
  <% =getHidden("selFinance",bytFinance) %>
  <% End If
  If Application("av_Module8On") and (pArticles = 5 or blnAdmin) Then %>
  <tr>
    <td class="dfont"><% =getLabel(Application("IDS_KnowledgeBase"),"chkKB") %></td>
    <td align=center><% =getCheckbox("chkKB",blnKB,"") %></td>
    <td align=right><% =showPermissions("selKB",bytKB) %></td>
  </tr>
  <% Else %>
  <% =getHidden("chkKB",blnKB) %>
  <% =getHidden("selKB",bytKB) %>
  <% End If %>
</form>
</table>

</div>

<div id="footerDiv" class="dvFooter">
<%
	Response.Write(getIconSave("edit"))
	Response.Write(getIconCancel())
%>
</div>

<%
	Call DisplayFooter(3)
%>