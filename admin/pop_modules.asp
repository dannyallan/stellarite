<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_admin.asp" -->
<%
	Call pageFunctions(90,5)

	Dim strModules	'as String
	Dim bytContact	'as Byte
	Dim bytClient	'as Byte
	Dim bytSale		'as Byte
	Dim bytService	'as Byte
	Dim bytSupport	'as Byte
	Dim bytQA		'as Byte
	Dim bytKB		'as Byte
	Dim bytFinance	'as Byte

	strTitle = Application("IDS_EnableModules")

	If not blnAdmin Then Call logError(2,1)

	If strDoAction = "edit" Then

		strModules = valNum(Request.Form("chkContact"),0,0)
		strModules = strModules & valNum(Request.Form("chkClient"),0,0)
		strModules = strModules & valNum(Request.Form("chkSale"),0,0)
		strModules = strModules & valNum(Request.Form("chkService"),0,0)
		strModules = strModules & valNum(Request.Form("chkSupport"),0,0)
		strModules = strModules & valNum(Request.Form("chkQA"),0,0)
		strModules = strModules & valNum(Request.Form("chkFinance"),0,0)
		strModules = strModules & valNum(Request.Form("chkKB"),0,0)

		Application.Lock
		Call setConfigValue("av_Modules",strModules)
		Application.Unlock

		Application.Contents.RemoveAll()

		Call closeWindow(strOpenerURL)
	Else
		bytContact	= valNum(Mid(Application("av_Modules"),1,1),1,0)
		bytClient	= valNum(Mid(Application("av_Modules"),2,1),1,0)
		bytSale		= valNum(Mid(Application("av_Modules"),3,1),1,0)
		bytService 	= valNum(Mid(Application("av_Modules"),4,1),1,0)
		bytSupport 	= valNum(Mid(Application("av_Modules"),5,1),1,0)
		bytQA		= valNum(Mid(Application("av_Modules"),6,1),1,0)
		bytFinance	= valNum(Mid(Application("av_Modules"),7,1),1,0)
		bytKB		= valNum(Mid(Application("av_Modules"),8,1),1,0)
	End If

	Call DisplayHeader(3)
	Call showEditHeader(strTitle,"","","","")
%>
<div id="contentDiv" class="dvBorder" style="height:330px;"><br>

<table border=0 cellspacing=0 cellpadding=4 width="99%">
<form name="frmAdmin" method="post" action="pop_modules.asp">
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
<% =getHidden("hdnWinOpen","") %>
  <tr class="hrow">
    <th align=left class="dlabel"><% =Application("IDS_Module") %></td>
    <th align=center class="dlabel"><% =Application("IDS_Enabled") %></td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_Contacts"),"chkContact") %></td>
    <td align=center><% =getCheckbox("chkContact",bytContact,"") %></td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_Accounts"),"chkClient") %></td>
    <td align=center><% =getCheckbox("chkClient",bytClient,"") %></td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_Sales"),"chkSale") %></td>
    <td align=center><% =getCheckbox("chkSale",bytSale,"") %></td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_ModServices"),"chkService") %></td>
    <td align=center><% =getCheckbox("chkService",bytService,"") %></td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_ModSupport"),"chkSupport") %></td>
    <td align=center><% =getCheckbox("chkSupport",bytSupport,"") %></td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_ModQualityAssurance"),"chkQA") %></td>
    <td align=center><% =getCheckbox("chkQA",bytQA,"") %></td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_ModFinance"),"chkFinance") %></td>
    <td align=center><% =getCheckbox("chkFinance",bytFinance,"") %></td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_KnowledgeBase"),"chkKB") %></td>
    <td align=center><% =getCheckbox("chkKB",bytKB,"") %></td>
  </tr>
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