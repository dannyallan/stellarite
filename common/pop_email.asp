<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_email.asp" -->
<%
	Call pageFunctions(0,1)

	Dim strMsg    'as String

	strTitle = getIDS("IDS_EmailNotification")
	bytMod = valNum(bytMod,1,1)
	lngModId = valNum(lngModId,3,1)

	Set objRS = objConn.Execute(getEmailSub(lngUserId,bytMod,lngModId))

	If objRS.BOF and objRS.EOF Then
		If valString(Request.Form("btnYes"),-1,0,0) = getIDS("IDS_Yes") Then
			objConn.Execute(insertEmailSub(lngUserId,bytMod,lngModId))
		End If
		strMsg = getIDS("IDS_MsgEmailNo") & vbCrLf
	Else
		If valString(Request.Form("btnNo"),-1,0,0) = getIDS("IDS_No") Then
			objConn.Execute(delEmailSub(lngUserId,objRS.fields(0).value,bytMod,lngModId))
		End If
		strMsg = getIDS("IDS_MsgEmailYes") & vbCrLf
	End If

	Call DisplayHeader(3)
%>


<form name="frmEmail" method="post" action="pop_email.asp?m=<% =bytMod %>&mid=<% =lngModId %>">
<table border=0 cellpadding=10 width="100%">
  <tr>
	<td class="dFont">
		<% =strMsg %>

	<br /><br />
	<center>
	<% =getSubmit("btnYes",getIDS("IDS_Yes"),70,"Y","") %>
	<% =getSubmit("btnNo",getIDS("IDS_No"),70,"N","") %>
	</center>
	</td>
  </tr>
</table>
</form>

<% If Request.Form.Count > 0 Then %>
<script language="JavaScript" type="text/javascript">
	window.close();
</script>
<% End If

	Call DisplayFooter(3)
%>
