<!--#include file="..\_inc\functions.asp" -->
<!--#include file="..\_inc\sql\sql_contacts.asp" -->
<!--#include file="..\_inc\sql\sql_clients.asp" -->
<!--#include file="..\_inc\sql\sql_sales.asp" -->
<%
	Call pageFunctions(0,1)

	Dim strPrefix           'as String
	Dim strCreated          'as String
	Dim strModified         'as String
	Dim dtmClosed           'as Date
	Dim strDescription      'as String
	Dim intNotes            'as Integer
	Dim intAttach           'as Integer
	Dim intEvents           'as Integer
	Dim intContacts         'as Integer
	Dim intSales            'as Integer
	Dim intProducts         'as Integer
	Dim intProjects         'as Integer
	Dim intTickets          'as Integer
	Dim intInvoices         'as Integer
	Dim strVertical         'as String
	Dim strAccountSize      'as String

	If bytMod = 0 or bytMod > 3 or lngModId = "" Then Call logError(3,1)

	Select Case bytMod
		Case 1
			Set objRS = objConn.Execute(getContact(0,lngModId))
			strPrefix = "K_"
		Case 2
			Set objRS = objConn.Execute(getClient(0,lngModId))
			strPrefix = "D_"
		Case 3
			Set objRS = objConn.Execute(getSale(0,lngModId))
			strPrefix = "S_"
	End Select

	If (objRS.BOF and objRS.EOF) Then
		Call logError(1,1)
	Else
		strCreated = showDate(0,objRS.fields(strPrefix & "CreatedDate").value) & " - " & showString(objRS.fields("CreatedBy").value)
		strModified = showDate(0,objRS.fields(strPrefix & "ModDate").value) & " - " & showString(objRS.fields("ModBy").value)
		intNotes = objRS.fields(strPrefix & "Notes").value
		intAttach = objRS.fields(strPrefix & "Attach").value
		intEvents = objRS.fields(strPrefix & "Events").value

		Select Case bytMod
			Case 1
				intSales = objRS.fields("K_Sales").value
				intProducts = objRS.fields("K_Serials").value
				intTickets = objRS.fields("K_Tickets").value
				intInvoices = objRS.fields("K_Invoices").value
			Case 2
				intContacts = objRS.fields("D_Contacts").value
				intSales = objRS.fields("D_Sales").value
				intProducts = objRS.fields("D_Serials").value
				intProjects = objRS.fields("D_Projects").value
				intTickets = objRS.fields("D_Tickets").value
				intInvoices = objRS.fields("D_Invoices").value
				strDescription = showParagraph(objRS.fields("D_ShortDesc").value)
				strVertical = getAOS(objRS.fields("D_Vertical").value)
				strAccountSize = getAOS(objRS.fields("D_Size").value)
			Case 3
				dtmClosed = showDate(0,objRS.fields("S_CloseDate").value)
		End Select

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
	  <td class="dFont"><% =strCreated %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Modified") %></td>
	  <td class="dFont"><% =strModified %></td>
	</tr>
<%  If bytMod = 3 Then %>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Closed") %></td>
	  <td class="dFont"><% =dtmClosed %></td>
	</tr>
<%
	End If
	If bytMod = 2 Then
%>
	<tr><td class="dFont" colspan=2>&nbsp;</td></tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_IndustrySector") %></td>
	  <td class="dFont"><% =strVertical %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_AccountSize") %></td>
	  <td class="dFont"><% =strAccountSize %></td>
	</tr>
	<tr><td class="dFont" colspan=2>&nbsp;</td></tr>
	<tr>
	  <td class="dFont" colspan=2>
	  <span class="bFont"><% =getIDS("IDS_Description") %></span><br />
	  <% =strDescription %>
	  </td>
	</tr>
<%
	End If
%>
    <tr><td class="dFont" colspan=2>&nbsp;</td></tr>
<%
	Call showCustomFields(bytMod)
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
	  <td class="dFont"><% =intAttach %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Events") %></td>
	  <td class="dFont"><% =intEvents %></td>
	</tr>
	<% If bytMod = 2 Then %>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Contacts") %></td>
	  <td class="dFont"><% =intContacts %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Projects") %></td>
	  <td class="dFont"><% =intProjects %></td>
	</tr>
	<% End If
	   If bytMod <> 3 Then %>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Sales") %></td>
	  <td class="dFont"><% =intSales %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Products") %></td>
	  <td class="dFont"><% =intProducts %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Tickets") %></td>
	  <td class="dFont"><% =intTickets %></td>
	</tr>
	<tr>
	  <td class="bFont"><% =getIDS("IDS_Invoices") %></td>
	  <td class="dFont"><% =intInvoices %></td>
	</tr>
	<% End If %>
  </table>

  </td></tr>
</table>

</div>

<%

	Call DisplayFooter(2)
%>

