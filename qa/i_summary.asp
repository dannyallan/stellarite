<!--#include file="..\_inc\functions.asp" -->
<!--#include file="..\_inc\sql\sql_qa.asp" -->
<%
	Call pageFunctions(6,1)

	Dim strBugType          'as String
	Dim strBugSource        'as String
	Dim strDescription      'as String
	Dim strSolution         'as String
	Dim strCause            'as String
	Dim intNotes            'as Integer
	Dim intAttachments      'as Integer
	Dim intEvents           'as Integer
	Dim intTickets          'as Integer

	If bytMod = "" or lngModId = "" Then Call logError(3,1)

	Set objRS = objConn.Execute(getBug(0,lngModId))

	If not (objRS.BOF and objRS.EOF) then
		strBugType = getAOS(objRS.fields("B_BugType").value)
		strBugSource = getAOS(objRS.fields("B_BugSource").value)
		intNotes = objRS.fields("B_Notes").value
		intAttachments = objRS.fields("B_Attach").value
		intEvents = objRS.fields("B_Events").value
		intTickets = objRS.fields("B_Tickets").value
		strDescription = showString(objRS.fields("B_Description").value)
		strSolution = showString(objRS.fields("B_Solution").value)
		strCause = getAOS(objRS.fields("B_Cause").value)
	End If

	strTitle = getIDS("IDS_Summary")
	Call DisplayHeader(2)
%>

<div id="contentDiv" class="dvNoBorder">

<table border="0" cellspacing="10" width="100%">
  <tr><td valign=top width="30%">

  <table border=0>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_BugType") %></td>
	  <td class="dFont"><% =strBugType %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_BugSource") %></td>
	  <td class="dFont"><% =strBugSource %></td>
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
	  <td class="bFont"><% =getIDS("IDS_Tickets") %></td>
	  <td class="dFont"><% =intTickets %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Events") %></td>
	  <td class="dFont"><% =intEvents %></td>
	</tr>
    <tr><td class="dFont" colspan=2>&nbsp;</td></tr>
<%
	Call showCustomFields(6)
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

