<!--#include file="..\_inc\functions.asp" -->
<!--#include file="..\_inc\sql\sql_support.asp" -->
<%
	Call pageFunctions(5,1)

	Dim strOwner        'as String
	Dim strPriority     'as String
	Dim strTicketType   'as String
	Dim strTicketSource 'as String
	Dim strSupportType  'as String
	Dim strProduct      'as String
	Dim strBuild        'as String
	Dim intNotes        'as Integer
	Dim intAttachments  'as Integer
	Dim intEvents       'as Integer
	Dim lngBugId        'as Long
	Dim strDescription  'as String
	Dim strSolution     'as String
	Dim strCause        'as String

	If bytMod = "" or lngModId = "" Then Call logError(3,1)

	Set objRS = objConn.Execute(getTicket(0,lngModId))

	If not (objRS.BOF and objRS.EOF) then
		strOwner = showString(objRS.fields("Owner").value)
		strPriority = getAOS(objRS.fields("T_Priority").value)
		strTicketType = getAOS(objRS.fields("T_TicketType").value)
		strTicketSource = getAOS(objRS.fields("T_TicketSource").value)
		strSupportType = getAOS(objRS.fields("T_SupportType").value)
		strProduct = getAOS(objRS.fields("T_ProductId").value)
		strBuild = showString(objRS.fields("T_Build").value)
		lngBugId = objRS.fields("T_BugId").value
		intNotes = showString(objRS.fields("T_Notes").value)
		intAttachments = showString(objRS.fields("T_Attach").value)
		intEvents = showString(objRS.fields("T_Events").value)
		strDescription = showString(objRS.fields("T_Description").value)
		strSolution = showString(objRS.fields("T_Solution").value)
		strCause = getAOS(objRS.fields("T_Cause").value)
	End If

	strTitle = getIDS("IDS_Summary")
	Call DisplayHeader(2)
%>

<div id="contentDiv" class="dvNoBorder">

<table border="0" cellspacing="10" width="100%">
  <tr><td valign=top width="30%">

  <table border=0>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Priority") %></td>
	  <td class="dFont"><% =strPriority %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_TicketType") %></td>
	  <td class="dFont"><% =strTicketType %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_TicketSource") %></td>
	  <td class="dFont"><% =strTicketSource %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_BugId") %></td>
	  <td class="dFont"><% =showLink(6,"../qa/bug.asp?id="&lngBugId,bigDigitNum(8,lngBugId)) %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_SupportType") %></td>
	  <td class="dFont"><% =strSupportType %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Owner") %></td>
	  <td class="dFont"><% =strOwner %></td>
	</tr>
	<tr><td colspan=2>&nbsp;</td></tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Product") %></td>
	  <td class="dFont"><% =strProduct %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Build") %></td>
	  <td class="dFont"><% =strBuild %></td>
	</tr>
	<tr><td colspan=2>&nbsp;</td></tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Notes") %></td>
	  <td class="dFont"><% =intNotes %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Attachments") %></td>
	  <td class="dFont"><% =intAttachments %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Events") %></td>
	  <td class="dFont"><% =intEvents %></td>
	</tr>
    <tr><td class="dFont" colspan=2>&nbsp;</td></tr>
<%
	Call showCustomFields(5)
%>
  </table>

  </td>
  <td valign=top align=right width="70%">

  <table border=0 width="100%">
	<tr>
	  <td valign=top class="bFont""><% =getIDS("IDS_Description") %></td>
	  <td width="90%"><% =getTextArea("txtDescription","dText",strDescription,"100%",8,"readonly=""readonly""") %></td>
	</tr>
	<tr>
	  <td valign=top class="bFont""><% =getIDS("IDS_Solution") %></td>
	  <td width="90%"><% =getTextArea("txtSolution","dText",strSolution,"100%",8,"readonly=""readonly""") %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Cause") %></td>
	  <td class="dFont"><% =strCause %></td>
	</tr>
  </table>

  </td></tr>
</table>

</div>

<%

	Call DisplayFooter(2)
%>
