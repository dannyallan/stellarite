<!--#include file="..\_inc\functions.asp" -->
<!--#include file="..\_inc\sql\sql_search.asp" -->
<%
	Call pageFunctions(0,1)

	Dim strRetField     'as String
	Dim strValue        'as String
	Dim strCol1         'as String
	Dim strCol2         'as String
	Dim strCol3         'as String
	Dim strParent       'as String
	Dim bytField        'as Byte
	Dim intOrder        'as Integer
	Dim intThird        'as Integer
	Dim arrName         'as Array
	Dim strDoSelect     'as String

	strRetField = valString(Request.QueryString("rVal"),100,0,0)
	strValue = valString(Request.QueryString("txtValue"),25,0,0)
	bytField = valNum(Request.QueryString("selField"),1,0)

	Select Case bytMod
		Case 1
			strCol1 = getIDS("IDS_Contacts")
			strCol2 = getIDS("IDS_Account")
			strParent = "sales/contact.asp?id="
			arrName = split(getIDS("IDS_NameFull") & "|" & getIDS("IDS_NameLast") & "|" & getIDS("IDS_NameFirst") & "|" & getIDS("IDS_Account") & "|" & getIDS("IDS_Division") & "|" & getIDS("IDS_City") & "|" & getIDS("IDS_State") & "|" & getIDS("IDS_Country") & "|" & getIDS("IDS_ZIP") & "|" & getIDS("IDS_Email") & "|" & getIDS("IDS_Phone"),"|")
			If bytField = 0 or bytField = 3 Then intThird = 4 Else intThird = bytField

		Case 2
			strCol1 = getIDS("IDS_Accounts")
			strCol2 = getIDS("IDS_Division")
			strParent = "sales/client.asp?id="
			arrName = split(getIDS("IDS_Account") & "|" & getIDS("IDS_Division") & "|" & getIDS("IDS_AccountId") & "|" & getIDS("IDS_SalesRep") & "|" & getIDS("IDS_Description"),"|")
			If bytField = 0 or bytField = 1 Then intThird = 2 Else intThird = bytField

		Case 3
			strCol1 = getIDS("IDS_PurchaseOrder")
			strCol2 = getIDS("IDS_Account")
			strParent = "sales/sale.asp?id="
			arrName = split(getIDS("IDS_Sale") & "|" & getIDS("IDS_Account") & "|" & getIDS("IDS_Division") & "|" & getIDS("IDS_SalesRep"),"|")
			If bytField = 0 or bytField = 1 Then intThird = 2 Else intThird = bytField

		Case 4
			strCol1 = getIDS("IDS_Project")
			strCol2 = getIDS("IDS_Account")
			strParent = "services/project.asp?id="
			arrName = split(getIDS("IDS_Project") & "|" & getIDS("IDS_Account") & "|" & getIDS("IDS_Division") & "|" & getIDS("IDS_Owner") & "|" & getIDS("IDS_Description"),"|")
			If bytField = 0 or bytField = 1 Then intThird = 3 Else intThird = bytField

		Case 5
			strCol1 = getIDS("IDS_Ticket")
			strCol2 = getIDS("IDS_Account")
			strParent = "support/ticket.asp?id="
			arrName = split(getIDS("IDS_Ticket") & "|" & getIDS("IDS_Account") & "|" & getIDS("IDS_Division") & "|" & getIDS("IDS_Contact") & "|" & getIDS("IDS_Owner") & "|" & getIDS("IDS_Priority") & "|" & getIDS("IDS_Description") & "|" & getIDS("IDS_Solution"),"|")
			If bytField = 0 or bytField = 1 Then intThird = 4 Else intThird = bytField

		Case 6
			strCol1 = getIDS("IDS_BugId")
			strCol2 = getIDS("IDS_Owner")
			strParent = "qa/bug.asp?id="
			arrName = split(getIDS("IDS_BugId") & "|" & getIDS("IDS_Owner") & "|" & getIDS("IDS_Priority") & "|" & getIDS("IDS_Description") & "|" & getIDS("IDS_Solution"),"|")
			If bytField = 0 or bytField = 1 Then intThird = 2 Else intThird = bytField

		Case 7
			strCol1 = getIDS("IDS_InvoiceId")
			strCol2 = getIDS("IDS_Account")
			strParent = "finance/invoice.asp?id="
			arrName = split(getIDS("IDS_InvoiceId") & "|" & getIDS("IDS_Account") & "|" & getIDS("IDS_Division") & "|" & getIDS("IDS_Contact") & "|" & getIDS("IDS_Owner"),"|")
			If bytField = 0 or bytField = 1 Then intThird = 2 Else intThird = bytField

		Case 50
			strCol1 = getIDS("IDS_Events")
			strCol2 = getIDS("IDS_Owner")
			strParent = "common/event.asp?id="
			arrName = split(getIDS("IDS_Title") & "|" & getIDS("IDS_Owner") & "|" & getIDS("IDS_StartTime"),"|")
			intThird = 2

		Case 0
			strCol1 = getIDS("IDS_NameFull")
			strCol2 = getIDS("IDS_Email")
			strParent = "admin/profile.asp?id="
			arrName = split(getIDS("IDS_NameFull") & "|" & getIDS("IDS_NameLast") & "|" & getIDS("IDS_NameFirst"),"|")
			If bytField = 0 Then intThird = 1 Else intThird = bytField

		Case Else
			Call logError(1,1)
	End Select


	strTitle     = getIDS("IDS_Search")
	strCol3     = arrName(intThird)

	If strValue <> "" Then
		Set objRS = objConn.Execute(getSearch(bytMod,intThird,bytField,strValue,Application("av_MaxRecords")))
		If not (objRS.BOF and objRS.EOF) Then arrRS = objRS.GetRows()
	End If

	If strValue = "" Then strValue = "*" Else strValue = showString(strValue)

	Sub getSearchDropDown(sDefault)

		For i = 0 to UBound(arrName)
			Response.Write(vbCrLf & vbTab & vbTab & "<option value=""" & i & """" & getDefault(0,CStr(i),sDefault) & ">" & arrName(i) & "</option>")
		Next

	End Sub

	Call DisplayHeader(3)
%>
<div id="headerDiv" class="dvNoBorder">

<script language="JavaScript" type="text/javascript">
if (window.opener==null) { sendBack('<% =getIDS("IDS_MsgNoParent") %>'); }
<%    If strRetField = "" Then
		Select Case bytMod
			Case 0,1,2,3,4,5,6,7     %>
function doSelect(sVal) {
	window.opener.location.href = "<% =Application("av_CRMDir") & strParent %>"+sVal;
<%             Case 50         %>
function doSelect(sVal1,sVal2,sVal3) {
	window.opener.location.href = "<% =Application("av_CRMDir") & strParent %>"+sVal1+"&m="+sVal2+"&mid="+sVal3;
<%         End Select
	Else
		Select Case strRetField
			Case "C"        %>
function doSelect(sVal1,sVal2) {
	window.opener.document.forms[0].txtClient.value = sVal1;
	window.opener.document.forms[0].txtDivision.value = sVal2;
	window.opener.document.forms[0].txtClient.select();
<%             Case "K"        %>
function doSelect(sVal1,sVal2,sVal3) {
	window.opener.document.forms[0].txtContact.value = sVal1;
	window.opener.document.forms[0].hdnContact.value = sVal2;
	window.opener.document.forms[0].hdnDivision.value = sVal3;
<%             Case Else         %>
function doSelect(sVal1) {
	window.opener.document.forms[0].<% =showString(strRetField) %>.value = sVal1;
	window.opener.document.forms[0].<% =showString(strRetField) %>.select();
<%         End Select
	End If
%>
	closeWindow(null);
}
</script>

<form name="frmSearch" method="get" action="pop_search.asp">
<table border=0 cellspacing=0 cellpadding=10 width="100%">
<% =getHidden("m",bytMod) %>
<% =getHidden("rVal",showString(strRetField)) %>
  <tr class="hRow">
	<td>
	  <table border=0>
		<tr>
		  <td><% =getLabel(getIDS("IDS_FindValue"),"txtValue") %></td>
	  <td><% =getTextField("txtValue","oText",strValue,22,25,"") %></td>
		</tr>
		<tr>
		  <td><% =getLabel(getIDS("IDS_Search"),"selField") %></td>
		  <td><select name="selField" id="selField" class="oByte" style="width:150px"><% Call getSearchDropDown(bytField) %>
		  </select></td>
		</tr>
	  </table>
	</td>
	<td>
	  <% =getSubmit("btnSearch",getIDS("IDS_Search"),100,"S","") %><br />
	  <% =getSubmit("btnCancel",getIDS("IDS_Cancel"),100,"X","onClick=""window.close();""") %>
	</td>
  </tr>
</table>
</form>

<table border=0 cellspacing=0 cellpadding=2 width="100%">
  <tr class="eTab">
	<th width="32%">&nbsp;&nbsp;<% =strCol1 %></td>
	<th width="32%">&nbsp;&nbsp;<% =strCol2 %></td>
	<th width="36%">&nbsp;&nbsp;<% =strCol3 %></td>
  </tr>
</table>

</div>

<div id="contentDiv" class="dvNoBorder" style="height:313px;">

<table border=0 cellspacing=1 cellpadding=1 width="100%">
<%
	If isArray(arrRS) Then
		For i = 0 to UBound(arrRS,2)

			Select Case bytMod
				Case 1
					If strRetField = "" Then strDoSelect = arrRS(0,i) Else strDoSelect = arrRS(1,i) & "','" & arrRS(0,i) & "','" & arrRS(3,i)
					strCol1 = arrRS(1,i)
					strCol2 = arrRS(2,i)
					strCol3 = arrRS(4,i)
				Case 2
					If strRetField = "" Then strDoSelect = arrRS(0,i) Else strDoSelect = arrRS(1,i) & "','" & arrRS(2,i)
					strCol1 = arrRS(1,i)
					strCol2 = arrRS(2,i)
					strCol3 = arrRS(4,i)
				Case 3
					strDoSelect = arrRS(0,i)
					strCol1 = bigDigitNum(7,arrRS(0,i))
					strCol2 = arrRS(1,i)
					strCol3 = arrRS(3,i)
				Case 4
					If strRetField = "" Then strDoSelect = arrRS(0,i) Else strDoSelect = arrRS(1,i)
					strCol1 = arrRS(1,i)
					strCol2 = arrRS(2,i)
					strCol3 = arrRS(4,i)
				Case 5
					strDoSelect = arrRS(0,i)
					strCol1 = bigDigitNum(7,arrRS(0,i))
					strCol2 = arrRS(1,i)
					strCol3 = arrRS(3,i)
				Case 6
					strDoSelect = arrRS(0,i)
					strCol1 = bigDigitNum(7,arrRS(0,i))
					strCol2 = arrRS(1,i)
					strCol3 = arrRS(2,i)
				Case 7
					strDoSelect = arrRS(0,i)
					strCol1 = bigDigitNum(7,arrRS(0,i))
					strCol2 = arrRS(1,i)
					strCol3 = arrRS(2,i)
				Case 50
					If strRetField = "" Then strDoSelect = arrRS(0,i) & "','" & arrRS(1,i) & "','" & arrRS(2,i) Else strDoSelect = arrRS(3,i)
					strCol1 = arrRS(3,i)
					strCol2 = arrRS(4,i)
					strCol3 = showDate(0,arrRS(5,i))
				Case 0
					If strRetField = "" Then strDoSelect = arrRS(0,i) Else strDoSelect = arrRS(1,i)
					strCol1 = arrRS(1,i)
					strCol2 = arrRS(2,i)
					strCol3 = arrRS(3,i)
			End Select
%>
 <tr class="dRow1" onClick="doSelect('<% =strDoSelect %>');" onMouseOver="this.style.cursor='hand';this.className='dRow2';" onMouseOut="this.style.cursor='default';this.className='dRow1';">
	<td class="dFont" width="33%">&nbsp;&nbsp;<a href="Javascript:doSelect('<% =strDoSelect %>');" class="dFont"><% =trimString(strCol1,20) %></a></td>
	<td class="dFont" width="33%">&nbsp;&nbsp;<% =trimString(strCol2,20) %></td>
	<td class="dFont" width="33%">&nbsp;&nbsp;<% =trimString(strCol3,20) %></td>
  </tr>
<%

		Next
	End If
%>
</table>

</div>

<%
	Call DisplayFooter(3)
%>

