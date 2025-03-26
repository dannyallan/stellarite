<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_finance.asp" -->
<%
	Call pageFunctions(7,3)

	Dim strInfo			'as String
	Dim strCreatedBy	'as String
	Dim datCreatedDate	'as Date
	Dim strModBy		'as String
	Dim datModDate		'as Date

	If lngRecordId = "" Then Call logError(3,1)
	strTitle = Application("IDS_InvoiceDetails")

	If strDoAction = "edit" Then

		strInfo = valString(Request.Form("txtInfo"),-1,0,5)

		Call updatePaymentDetails(lngUserId,lngRecordId,strInfo)

		Call closeWindow(strOpenerURL)
	Else
		Set objRS = objConn.Execute(getInvoice(0,lngRecordId))
		If not (objRS.BOF and objRS.EOF) Then
			strInfo = objRS.fields("I_PayInfo").value
			strCreatedBy = showString(objRS.fields("CreatedBy").value)
			datCreatedDate = objRS.fields("I_CreatedDate").value
			strModBy = showString(objRS.fields("ModBy").value)
			datModDate = objRS.fields("I_ModDate").value
		End If
	End If

	Call DisplayHeader(3)
	Call showEditHeader(strTitle,strCreatedBy,datCreatedDate,strModBy,datModDate)
%>
<div id="contentDiv" class="dvBorder" style="height:330px;"><br>

<table border=0 cellspacing=5 width="100%">
<form name="frmAdmin" method="post" action="pop_payment.asp?id=<% =lngRecordId %>">
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
<% =getHidden("hdnWinOpen","") %>
  <tr>
    <td>
      <% =getLabel(Application("IDS_CharLimit255"),"txtInfo") %>
      <br>
      <% =getTextArea("txtInfo","oText",strInfo,"100%",15,"") %>
    </td>
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