<!--#include file="..\_inc\functions.asp" -->
<!--#include file="..\_inc\sql\sql_services.asp" -->
<%
	Call pageFunctions(4,1)

	Dim strCreatedBy    'as String
	Dim strModBy        'as String
	Dim dtmCloseDate    'as Date
	Dim strDescription  'as String
	Dim intNotes        'as Integer
	Dim intAttachments  'as Integer
	Dim intEvents       'as Integer

	If bytMod = "" or lngModId = "" Then Call logError(3,1)

	Set objRS = objConn.Execute(getProject(0,lngModId))

	If not (objRS.BOF and objRS.EOF) then
		strCreatedBy = showDate(0,objRS.fields("P_CreatedDate").value) & " - " & showString(objRS.fields("CreatedBy").value)
		strModBy = showDate(0,objRS.fields("P_ModDate").value) & " - " & showString(objRS.fields("ModBy").value)
		dtmCloseDate = showDate(0,objRS.fields("P_CloseDate").value)
		strDescription = showParagraph(objRS.fields("P_ShortDesc").value)
		intNotes = objRS.fields("P_Notes").value
		intAttachments = objRS.fields("P_Attach").value
		intEvents = objRS.fields("P_Events").value
	End If

	strTitle = getIDS("IDS_Summary")
	Call DisplayHeader(2)
%>

<div id="contentDiv" class="dvNoBorder">

<table border="0" cellspacing="10" width="100%">
  <tr><td valign=top width="50%">

  <table border=0>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Created") %></td>
	  <td class="dFont"><% =strCreatedBy %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Modified") %></td>
	  <td class="dFont"><% =strModBy %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Closed") %></td>
	  <td class="dFont"><% =dtmCloseDate %></td>
	</tr>
	<tr>
	  <td class="dFont" colspan=2>
	  <br /><br />
	  <span class="bFont"><% =getIDS("IDS_Description") %></span><br />
	  <% =showParagraph(strDescription) %>
	  </td>
	</tr>
    <tr><td class="dFont" colspan=2>&nbsp;</td></tr>
<%
	Call showCustomFields(4)
%>
  </table>

  </td>
  <td valign=top width="50%">

  <table border=0>
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
  </table>

  </td></tr>
</table>

</div>

<%
	Call DisplayFooter(2)
%>