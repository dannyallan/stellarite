<!--#include file="..\_inc\functions.asp" -->
<!--#include file="..\_inc\sql\sql_finance.asp" -->
<%
	Call pageFunctions(7,1)

	Dim intNotes        'as Integer
	Dim intAttachments  'as Integer
	Dim intEvents       'as Integer
	Dim intSales        'as Integer
	Dim intProducts     'as Integer
	Dim intProjects     'as Integer
	Dim strPayInfo		'as String
	Dim strCreatedBy    'as String
	Dim strModBy        'as String

	If bytMod = "" or lngModId = "" Then Call logError(3,1)

	Set objRS = objConn.Execute(getInvoice(0,lngModId))

	If not (objRS.BOF and objRS.EOF) then
		intNotes = objRS.fields("I_Notes").value
		intAttachments = objRS.fields("I_Attach").value
		intEvents = objRS.fields("I_Events").value
		intSales = objRS.fields("I_Sales").value
		intProducts = objRS.fields("I_Serials").value
		intProjects = objRS.fields("I_Projects").value
		strPayInfo = objRS.fields("I_PayInfo").value
		strCreatedBy = showDate(0,objRS.fields("I_CreatedDate").value) & " - " & showString(objRS.fields("CreatedBy").value)
		strModBy = showDate(0,objRS.fields("I_ModDate").value) & " - " & showString(objRS.fields("ModBy").value)
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
	  <td class="dFont" colspan=2>
	  <br /><br />
	  <span class="bFont"><% =getIDS("IDS_InvoiceDetails") %></span><br />
	  <% =showParagraph(strPayInfo) %>
	  </td>
	</tr>
    <tr><td class="dFont" colspan=2>&nbsp;</td></tr>
<%
	Call showCustomFields(7)
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
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Sales") %></td>
	  <td class="dFont"><% =intSales %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Products") %></td>
	  <td class="dFont"><% =intProducts %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Projects") %></td>
	  <td class="dFont"><% =intProjects %></td>
	</tr>
  </table>

  </td></tr>
</table>

</div>

<%
	Call DisplayFooter(2)
%>