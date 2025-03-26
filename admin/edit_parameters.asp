<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_admin.asp" -->
<!--#include file="..\_inc\timezone.asp" -->
<!--#include file="..\_inc\currency.asp" -->
<%
	Call pageFunctions(90,5)

	Dim strURL              'as String
	Dim intTimeZone         'as String
	Dim strCurrency         'as String
	Dim decTaxRate          'as String
	Dim blnTaxChange        'as Boolean
	Dim bytLoginAttempts    'as Byte
	Dim intMaxRecords       'as Integer
	Dim blnNavigation       'as Boolean
	Dim blnHideUsers        'as Boolean


	strTitle = getIDS("IDS_CRMParameters")

	If not blnAdmin Then Call logError(2,1)

	If strDoAction = "edit" Then

		intTimeZone = CStr(valNum(Request.Form("selTimeZone"),2,1))
		strCurrency = valString(Request.Form("selCurrency"),3,0,0)
		decTaxRate = CStr(valNum(Request.Form("txtTaxRate"),5,0))
		blnTaxChange = CStr(valNum(Request.Form("chkTaxChange"),0,0))
		bytLoginAttempts = CStr(valNum(Request.Form("txtLoginAttempts"),1,1))
		intMaxRecords = CStr(valNum(Request.Form("txtMaxRecords"),2,1))
		blnNavigation = CStr(valNum(Request.Form("selNavigation"),0,0))
		blnHideUsers = CStr(valNum(Request.Form("chkHideUsers"),0,0))

		strURL = valString(Request.Form("txtURL"),255,1,2)
		If Right(strURL,1) <> "/" Then strURL = strURL & "/"

		Application.Lock

		Call setAppVar("av_CRMURL",strURL)
		Call setAppVar("av_TimeZone",intTimeZone)
		Call setAppVar("av_Navigation",blnNavigation)
		Call setAppVar("av_Currency",strCurrency)
		Call setAppVar("av_TaxRate",decTaxRate)
		Call setAppVar("av_TaxChange",blnTaxChange)
		Call setAppVar("av_LoginAttempts",bytLoginAttempts)
		Call setAppVar("av_MaxRecords",intMaxRecords)
		Call setAppVar("av_HideUsers",blnHideUsers)

		Application.Unlock

		Call closeEdit()
	Else
		strURL                = Application("av_CRMURL")
		intTimeZone            = Application("av_TimeZone")
		blnNavigation        = Application("av_Navigation")
		strCurrency            = Application("av_Currency")
		decTaxRate            = Application("av_TaxRate")
		blnTaxChange        = Application("av_TaxChange")
		bytLoginAttempts    = Application("av_LoginAttempts")
		intMaxRecords        = Application("av_MaxRecords")
		blnHideUsers        = Application("av_HideUsers")

		If intTimeZone = "" Then intTimeZone = 0
		If strCurrency = "" Then strCurrency = "USD"
		If decTaxRate = "" Then decTaxRate = "0.0"
	End If

	Call DisplayHeader(1)
	Call showEditHeader(strTitle,"","","","")
%>
<form name="frmAdmin" method="post" action="edit_parameters.asp">
<div id="contentDiv" class="dvBorder" style="height:<% =intScreenH-170 %>px;"><br />

<table border=0 cellspacing=5>
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
  <tr>
	<td><% =getLabel(getIDS("IDS_CRMURL"),"txtURL") %></td>
	<td><% =getTextField("txtURL","oLink",strURL,40,255,"") %></td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_TimeZone"),"selTimeZone") %></td>
	<td><% =getTimeZone(intTimeZone) %></td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_Currency"),"selCurrency") %></td>
	<td><% =getCurrency(260,"selCurrency",strCurrency) %></td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_Navigation"),"selNavigation") %></td>
	<td>
	  <select name="selNavigation" id="selNavigation" class="oBool" onChange="doChange();" style="width:260px;">
		<option value="0"<% =getDefault(0,blnNavigation,"0") & ">" & getIDS("IDS_NavIcons") %></option>
		<option value="1"<% =getDefault(0,blnNavigation,"1") & ">" & getIDS("IDS_NavButtons") %></option>
	  </select>
   </td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_TaxRate"),"txtTaxRate") %></td>
	<td><% =getTextField("txtTaxRate","mCurrency",decTaxRate,4,4,"") %></td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_LoginAttempts"),"txtLoginAttempts") %></td>
	<td><% =getTextField("txtLoginAttempts","mByte",bytLoginAttempts,4,4,"") %></td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_MaximumRecords"),"txtMaxRecords") %></td>
	<td><% =getTextField("txtMaxRecords","mInt",intMaxRecords,4,6,"") %></td>
  </tr>
  <tr>
	<td colspan="2"><hr /><% =getCheckbox("chkTaxChange",blnTaxChange,"") %>&nbsp;&nbsp;&nbsp;<% =getLabel(getIDS("IDS_MsgTaxChange"),"chkTaxChange") %></td>
  </tr>
  <tr>
	<td colspan="2"><% =getCheckbox("chkHideUsers",blnHideUsers,"") %>&nbsp;&nbsp;&nbsp;<% =getLabel(getIDS("IDS_MsgHideUsers"),"chkHideUsers") %></td>
  </tr>
</table>

</div>

<div id="footerDiv" class="dvFooter">
<%
	Response.Write(getIconSave("edit"))
	Response.Write(getIconCancel("back"))
%>
</div>
</form>

<%
	Call DisplayFooter(1)
%>