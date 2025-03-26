<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_contacts.asp" -->
<!--#include file="..\_inc\sql\sql_clients.asp" -->
<!--#include file="..\_inc\sql\sql_sales.asp" -->
<!--#include file="..\_inc\sql\sql_services.asp" -->
<!--#include file="..\_inc\sql\sql_support.asp" -->
<!--#include file="..\_inc\sql\sql_qa.asp" -->
<!--#include file="..\_inc\sql\sql_finance.asp" -->
<!--#include file="..\_inc\sql\sql_notes.asp" -->
<!--#include file="..\_inc\sql\sql_attachments.asp" -->
<!--#include file="..\_inc\sql\sql_events.asp" -->
<!--#include file="..\_inc\sql\sql_products.asp" -->
<%
	Call pageFunctions(0,1)

	Dim lngEventId  'as Long
	Dim lngDivId    'as Long
	Dim objRS2      'as Object

	strTitle = valString(Request.QueryString("title"),-1,0,0)
	lngEventId = valNum(Request.QueryString("eid"),3,0)

	If bytMod = "" or lngModId = "" Then Call logError(3,1)

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<html lang="<% =strLanguage %>">
<head>
<title><% =showString(strModItem & ": " & strTitle) %></title>
<link href="../common/css.asp" rel="stylesheet" type="text/css" />
<script language="JavaScript" type="text/javascript">
	function printWin() {
		window.print();
		window.close();
	}
</script>
</head>

<%
	If valString(Request.Form("btnPrint"),-1,0,0) = getIDS("IDS_Print") Then


		Response.Write("<body onLoad='printWin()'>" & vbCrLf & vbCrLf & _
				"<span class=""pFont"">" & vbCrLf & vbCrLf & _
				"<h4>" & showString(strModItem & ": " & strTitle) & "</h4><hr />" & vbCrLf & vbCrLf)

		If bytMod = 1 Then lngDivId = getValue("DivId","CRM_Contacts","ContactId="&lngModId,0)
		If bytMod = 2 Then lngDivId = lngModId

		If valNum(Request.Form("chkDetails"),0,0) = 1 Then
			Select Case bytMod
				Case 1
					Set objRS = objConn.Execute(getContact(0,lngModId))

					If objRS.EOF Then
						Response.Write(getIDS("IDS_Deleted") & "<hr /></body></html>")
						Call endResponse()
					Else
						Response.Write("<b>" & getIDS("IDS_Name") & ":</b> &nbsp;" & showString(objRS.fields("K_FirstName").value & " " & objRS.fields("K_LastName").value) & "<br />" & vbCrLf & _
								"<b>Address:</b><br />" & vbCrLf)
						If objRS.fields("K_Address1").value <> "" Then Response.Write(showString(objRS.fields("K_Address1").value) & "<br />" & vbCrLf)
						If objRS.fields("K_Address2").value <> "" Then Response.Write(showString(objRS.fields("K_Address2").value) & "<br />" & vbCrLf)
						If objRS.fields("K_Address3").value <> "" Then Response.Write(showString(objRS.fields("K_Address3").value) & "<br />" & vbCrLf)
						If objRS.fields("K_City").value <> "" Then Response.Write(showString(objRS.fields("K_City").value))
						If objRS.fields("K_City").value <> "" and showString(objRS.fields("K_State").value) <> "" Then Response.Write(", ")
						If objRS.fields("K_State").value <> "" Then Response.Write(showString(objRS.fields("K_State").value))
						Response.Write("<br />" & vbCrLf)
						If objRS.fields("K_Country").value <> "" Then Response.Write(showString(objRS.fields("K_Country").value))
						If objRS.fields("K_Country").value <> "" and showString(objRS.fields("K_ZIP").value) <> "" Then Response.Write(", ")
						If objRS.fields("K_ZIP").value <> "" Then Response.Write(showString(objRS.fields("K_ZIP").value))
						Response.Write("<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Account") & ":</b> &nbsp;" & showString(objRS.fields("C_Client").value))
						If objRS.fields("D_Division").value <> "" Then Response.Write(" - " & showString(objRS.fields("D_Division").value))
						Response.Write("<br /><b>" & getIDS("IDS_Department") & ":</b> &nbsp;" & showString(objRS.fields("K_Dept").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_JobTitle") & ":</b> &nbsp;" & showString(objRS.fields("K_JobTitle").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Email") & ":</b> &nbsp;" & showString(objRS.fields("K_Email").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_EmailOptOut") & ":</b> &nbsp;" & showString(objRS.fields("K_EmailOptOut").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Phone") & " 1:</b> &nbsp;" & showPhone(objRS.fields("K_Phone1").value) & vbCrLf)
						If objRS.fields("K_Phone1").value <> "" and objRS.fields("K_Ext1").value <> "" Then Response.Write(" x." & objRS.fields("K_Ext1").value)
						Response.Write("<br /><b>" & getIDS("IDS_Phone") & " 2:</b> &nbsp;" & showPhone(objRS.fields("K_Phone2").value) & vbCrLf)
						If objRS.fields("K_Phone2").value <> "" and objRS.fields("K_Ext2").value <> "" Then Response.Write(" x." & objRS.fields("K_Ext2").value)
						Response.Write("<br /><b>" & getIDS("IDS_Fax") & ":</b> &nbsp;" & showPhone(objRS.fields("K_Fax").value) & "<hr />" & vbCrLf)
						Response.Write("<br /><b>" & getIDS("IDS_DoNotCall") & ":</b> &nbsp;" & showTrueFalse(objRS.fields("K_DoNotCall").value) & "<hr />" & vbCrLf)
					End if
				Case 2
					Set objRS = objConn.Execute(getClient(0,lngModId))

					If objRS.EOF Then
						Response.Write(getIDS("IDS_Deleted") & "<hr /></body></html>")
						Call endResponse()
					Else
						Response.Write("<b>" & getIDS("IDS_Account") & ":</b> &nbsp;" & showString(strTitle) & "<br />" & vbCrLf)
						If objRS.fields("D_Division").value <> "" Then Response.Write("<b>" & getIDS("IDS_Division") & ":</b> &nbsp;" & showString(objRS.fields("D_Division").value) & "<br />" & vbCrLf)
						Response.Write("<b>" & getIDS("IDS_SalesRep") & ":</b> &nbsp;" & showString(objRS.fields("SalesRep").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_AccountType") & ":</b> &nbsp;" & getAOS(objRS.fields("D_AccountType").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Description") & ":</b> &nbsp;" & showParagraph(objRS.fields("D_ShortDesc").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_IndustrySector") & ":</b> &nbsp;" & getAOS(objRS.fields("D_Vertical").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_SalesRegion") & ":</b> &nbsp;" & getAOS(objRS.fields("D_Region").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Website") & ":</b> &nbsp;" & showString(objRS.fields("D_Website").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Reference") & ":</b> &nbsp;" & showString(objRS.fields("D_RefAccount").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_ProblemFlag") & ":</b> &nbsp;" & showTrueFalse(getIDS("D_ProbFlag")) )
					End if
				Case 3
					Set objRS = objConn.Execute(getSale(0,lngModId))

					If objRS.EOF Then
						Response.Write(getIDS("IDS_Deleted") & "<hr /></body></html>")
						Call endResponse()
					Else
						Response.Write("<b>" & getIDS("IDS_Sale") & ":</b> &nbsp;" & bigDigitNum(7,objRS.fields("SaleId").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Account") & ":</b> &nbsp;" & showString(objRS.fields("C_Client").value) & "<br />" & vbCrLf)
						If objRS.fields("D_Division").value <> "" Then Response.Write("<b>" & getIDS("IDS_Division") & ":</b> &nbsp;" & showString(objRS.fields("D_Division").value) & "<br />" & vbCrLf)
						If objRS.fields("S_CloseDate").value <> "" Then Response.Write("<b>" & getIDS("IDS_CloseDate") & ":</b> &nbsp;" & showDate(0,objRS.fields("S_CloseDate").value))
						Response.Write("<b>" & getIDS("IDS_SaleValue") & ":</b> &nbsp;" & FormatCurrency(objRS.fields("S_SaleValue").value) & "<hr />" & vbCrLf)
					End if
				Case 4
					Set objRS = objConn.Execute(getProject(0,lngModId))

					If objRS.EOF Then
						Response.Write(getIDS("IDS_Deleted") & "<hr /></body></html>")
						Call endResponse()
					Else
						Response.Write("<b>" & getIDS("IDS_Project") & ":</b> &nbsp;" & showString(objRS.fields("P_Title").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Account") & ":</b> &nbsp;" & showString(objRS.fields("C_Client").value) & "<br />" & vbCrLf)
						If objRS.fields("D_Division").value <> "" Then Response.Write("<b>" & getIDS("IDS_Division") & ":</b> &nbsp;" & showString(objRS.fields("D_Division").value) & "<br />" & vbCrLf)
						Response.Write("<b>" & getIDS("IDS_Owner") & ":</b> &nbsp;" & showString(objRS.fields("Owner").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Description") & ":</b> &nbsp;" & showParagraph(objRS.fields("P_ShortDesc").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_DaysTotal") & ":</b> &nbsp;" & objRS.fields("P_DaysTotal").value & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_DaysOwed") & ":</b> &nbsp;" & objRS.fields("P_DaysOwed").value & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Created") & ":</b> &nbsp;" & showDate(0,objRS.fields("P_CreatedDate").value) & " - " & showString(objRS.fields("CreatedBy").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Modified") & ":</b> &nbsp;" & showDate(0,objRS.fields("P_ModDate").value) & " - " & showString(objRS.fields("ModBy").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_CloseDate") & ":</b> &nbsp;" & showDate(0,objRS.fields("P_CloseDate").value) & "<hr />" & vbCrLf)
					End if
				Case 5
					Set objRS = objConn.Execute(getTicket(0,lngModId))

					If objRS.EOF Then
						Response.Write(getIDS("IDS_Deleted") & "<hr /></body></html>")
						Call endResponse()
					Else
						Response.Write("<b>" & getIDS("IDS_Ticket") & ":</b> &nbsp;" & bigDigitNum(7,objRS.fields("TicketId").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Contact") & ":</b> &nbsp;" & showString(objRS.fields("Contact").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Account") & ":</b> &nbsp;" & showString(objRS.fields("C_Client").value) & "<br />" & vbCrLf)
						If objRS.fields("D_Division").value <> "" Then Response.Write("<b>" & getIDS("IDS_Division") & ":</b> &nbsp;" & showString(objRS.fields("D_Division").value) & "<br />" & vbCrLf)
						Response.Write("<b>" & getIDS("IDS_Owner") & ":</b> &nbsp;" & showString(objRS.fields("Owner").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_HotIssue") & ":</b> &nbsp;" & showString(objRS.fields("T_HotIssue").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Priority") & ":</b> &nbsp;" & getAOS(objRS.fields("T_Priority").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_TicketType") & ":</b> &nbsp;" & getAOS(objRS.fields("T_TicketType").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_TicketSource") & ":</b> &nbsp;" & getAOS(objRS.fields("T_TicketSource").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_ModName4") & ":</b> &nbsp;" & getAOS(objRS.fields("T_SupportType").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Product") & ":</b> &nbsp;" & getAOS(objRS.fields("T_ProductId").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Build") & ":</b> &nbsp;" & showString(objRS.fields("T_Build").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Bug") & ":</b> &nbsp;" & objRS.fields("T_BugId").value & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Description") & ":</b> &nbsp;" & showString(objRS.fields("T_Description").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Solution") & ":</b> &nbsp;" & showString(objRS.fields("T_Solution").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Cause") & ":</b> &nbsp;" & getAOS(objRS.fields("T_Cause").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Created") & ":</b> &nbsp;" & showDate(0,objRS.fields("T_CreatedDate").value) & " - " & showString(objRS.fields("CreatedBy").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Modified") & ":</b> &nbsp;" & showDate(0,objRS.fields("T_ModDate").value) & " - " & showString(objRS.fields("ModBy").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_CloseDate") & ":</b> &nbsp;" & showDate(0,objRS.fields("T_CloseDate").value) & "<hr />" & vbCrLf)
					End if
				Case 6
					Set objRS = objConn.Execute(getBug(0,lngModId))

					If objRS.EOF Then
						Response.Write(getIDS("IDS_Deleted") & "<hr /></body></html>")
						Call endResponse()
					Else
						Response.Write("<b>" & getIDS("IDS_Bug") & ":</b> &nbsp;" & bigDigitNum(7,objRS.fields("BugId").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Owner") & ":</b> &nbsp;" & showString(objRS.fields("Owner").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_HotIssue") & ":</b> &nbsp;" & showString(objRS.fields("B_HotIssue").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Priority") & ":</b> &nbsp;" & getAOS(objRS.fields("B_Priority").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_BugType") & ":</b> &nbsp;" & getAOS(objRS.fields("B_BugType").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_BugSource") & ":</b> &nbsp;" & getAOS(objRS.fields("B_BugSource").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Product") & ":</b> &nbsp;" & getAOS(objRS.fields("B_ProductId").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Build") & ":</b> &nbsp;" & showString(objRS.fields("B_Build").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Description") & ":</b> &nbsp;" & showString(objRS.fields("B_Description").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Solution") & ":</b> &nbsp;" & showString(objRS.fields("B_Solution").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Cause") & ":</b> &nbsp;" & getAOS(objRS.fields("B_Cause").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Created") & ":</b> &nbsp;" & showDate(0,objRS.fields("B_CreatedDate").value) & " - " & showString(objRS.fields("CreatedBy").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Modified") & ":</b> &nbsp;" & showDate(0,objRS.fields("B_ModDate").value) & " - " & showString(objRS.fields("ModBy").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_CloseDate") & ":</b> &nbsp;" & showDate(0,objRS.fields("B_CloseDate").value) & "<hr />" & vbCrLf)
					End if
				Case 7
					Set objRS = objConn.Execute(getInvoice(0,lngModId))

					If objRS.EOF Then
						Response.Write(getIDS("IDS_Deleted") & "<hr /></body></html>")
						Call endResponse()
					Else
						Response.Write("<b>" & getIDS("IDS_InvoiceId") & ":</b> &nbsp;" & bigDigitNum(7,objRS.fields("InvoiceId").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Contact") & ":</b> &nbsp;" & showString(objRS.fields("Contact").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Account") & ":</b> &nbsp;" & showString(objRS.fields("C_Client").value) & "<br />" & vbCrLf)
						If objRS.fields("D_Division").value <> "" Then Response.Write("<b>" & getIDS("IDS_Division") & ":</b> &nbsp;" & showString(objRS.fields("D_Division").value) & "<br />" & vbCrLf)
						Response.Write("<b>" & getIDS("IDS_Owner") & ":</b> &nbsp;" & showString(objRS.fields("Owner").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_PurchaseOrder") & ":</b> &nbsp;" & showString(objRS.fields("I_PurchaseOrder").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Invoice") & ":</b> &nbsp;" & showString(objRS.fields("I_Received").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Type") & ":</b> &nbsp;" & getAOS(objRS.fields("I_Type").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Phase") & ":</b> &nbsp;" & getAOS(objRS.fields("I_Phase").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Value") & ":</b> &nbsp;" & showString(objRS.fields("I_Currency").value & FormatCurrency(objRS.fields("I_Value").value)) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Tax") & ":</b> &nbsp;" & showString(objRS.fields("I_Currency").value & FormatCurrency(objRS.fields("I_Tax").value)) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_InvoiceDate") & ":</b> &nbsp;" & showDate(0,objRS.fields("I_SendDate").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_InvoiceDue") & ":</b> &nbsp;" & showDate(0,objRS.fields("I_DueDate").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_InvoicePaid") & ":</b> &nbsp;" & showDate(0,objRS.fields("I_PaidDate").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Created") & ":</b> &nbsp;" & showDate(0,objRS.fields("I_CreatedDate").value) & " - " & showString(objRS.fields("CreatedBy").value) & "<br />" & vbCrLf & _
								"<b>" & getIDS("IDS_Modified") & ":</b> &nbsp;" & showDate(0,objRS.fields("I_ModDate").value) & " - " & showString(objRS.fields("ModBy").value) & "<hr />" & vbCrLf)
					End if
			End Select
		End If

		If valNum(Request.Form("chkNotes"),0,0) = 1 Then

			Set objRS = objConn.Execute(getNotesBy(bytMod,lngModId,lngEventId,intMember,"1","DESC"))
			Response.Write("<h5>" & getIDS("IDS_Notes") & "</h5><hr />")

			If objRS.EOF Then
				Response.Write(getIDS("IDS_None") & ".<hr />")
			Else
				Do while NOT objRS.EOF
					Response.Write(    "<b>" & getIDS("IDS_Author") & ":</b> &nbsp;" & showString(objRS.fields("ModBy").value) & " &nbsp;&nbsp;" & _
							showDate(1,showString(objRS.fields("N_ModDate").value)) & "<br />" & vbCrLf & _
							"<b>" & getIDS("IDS_Type") & ":</b> &nbsp;" & getAOS(objRS.fields("N_ContactType").value) & "<br /><br />" & vbCrLf & _
							showString(objRS.fields("N_Info").value) & "<hr />" & vbCrLf)
					objRS.MoveNext
				Loop
			End If
		End if

		If valNum(Request.Form("chkAttach"),0,0) = 1 Then

			Set objRS = objConn.Execute(getAttachBy(bytMod,lngModId,lngEventId,intMember,"1","DESC"))
			Response.Write("<h5>" & getIDS("IDS_Attachments") & "</h5><hr />")

			If objRS.EOF Then
				Response.Write(getIDS("IDS_None") & ".<hr />")
			Else
				Do while NOT objRS.EOF
					Response.Write(    "<b>" & getIDS("IDS_Author") & ":</b> &nbsp;" & showString(objRS.fields("ModBy").value) & " &nbsp;&nbsp;" & vbCrLf & _
							showDate(1,showString(objRS.fields("A_ModDate").value)) & "<br />" & vbCrLf & _
							"<b>" & getIDS("IDS_Title") & ":</b> &nbsp;" & showString(objRS.fields("A_Title").value) & "<br />" & vbCrLf & _
							"<b>" & getIDS("IDS_Type") & ":</b> &nbsp;" & showString(objRS.fields("O_Value").value) & "<br /><br />" & vbCrLf & _
							showString(objRS.fields("A_Info").value) & "<br />" & vbCrLf)

					Set objRS2 = objConn.Execute(getAttachLinks(objRS.fields("AttachId").value))

					If not objRS2.EOF Then
						Do while not objRS2.EOF
							Response.Write("<li><a href=""" & objRS2(0) & """ target=""_top"">" & objRS2(0) & "</a></li>" & vbCrLf)
							objRS2.MoveNext
						Loop
					End If

					objRS.MoveNext

					Response.Write("<hr />" & vbCrLf & vbCrLf)
				Loop

				If isObject(objRS2) Then
					objRS2.Close
					Set objRS2 = Nothing
				End If
			End If
		End if

		If valNum(Request.Form("chkContacts"),0,0) = 1 Then

			Set objRS = objConn.Execute(getContactsByDiv(lngModId,"1","ASC"))
			Response.Write("<h5>" & getIDS("IDS_Contacts") & "</h5><hr />")

			If objRS.EOF Then
				Response.Write(getIDS("IDS_None") & ".<hr />")
			Else
				Do while NOT objRS.EOF
					Response.Write(    "<b>" & getIDS("IDS_Name") & ":</b> &nbsp;" & showString(objRS.fields(1).value) & "<br />" & vbCrLf)
					Response.Write("<b>" & getIDS("IDS_Email") & ":</b> &nbsp;" & showEmail(objRS.fields("K_Email").value) & "<br />" & vbCrLf)
					Response.Write("<b>" & getIDS("IDS_Phone") & ":</b> &nbsp;" & showPhone(objRS.fields("K_Phone1").value) & vbCrLf)
					If objRS.fields("K_Phone1").value <> "" and objRS.fields("K_Ext1").value <> "" Then Response.Write(" x." & objRS.fields("K_Ext1").value)
					Response.Write("<hr />" & vbCrLf)

					objRS.MoveNext
				Loop
			End If
		End if

		If valNum(Request.Form("chkEvents"),0,0) = 1 Then

			Set objRS = objConn.Execute(getEventsBy(bytMod,lngModId,intMember,"1","ASC"))
			Response.Write("<h5>" & getIDS("IDS_Events") & "</h5><hr />")

			If objRS.EOF Then
				Response.Write(getIDS("IDS_None") & ".<hr />")
			Else
				Do while NOT objRS.EOF
					Response.Write("<b>" & getIDS("IDS_Owner") & ":</b> &nbsp;" & showString(objRS.fields("Owner").value) & "<br />" & vbCrLf & _
							"<b>" & getIDS("IDS_Title") & ":</b> &nbsp;" & showString(objRS.fields("E_Title").value) & "<br />" & vbCrLf & _
							"<b>" & getIDS("IDS_Value") & ":</b> &nbsp;" & showString(objRS.fields("O_Value").value))

					If objRS.fields("E_Onsite").value = 1 Then Response.Write(" &nbsp; - " & getIDS("IDS_Onsite"))
					If objRS.fields("E_Billable").value = 1 Then Response.Write(" &nbsp; - " & getIDS("IDS_Billable"))

					Response.Write("<br /><b>" & getIDS("IDS_Start") & ":</b> &nbsp;" & showDate(1,objRS.fields("E_StartTime").value) & "<br />" & vbCrLf & _
							"<b>" & getIDS("IDS_End") & ":</b> &nbsp;" & showDate(1,objRS.fields("E_EndTime").value) & "<hr />" & vbCrLf)
					objRS.MoveNext
				Loop
			End If
		End if

		If valNum(Request.Form("chkSales"),0,0) = 1 Then

			Set objRS = objConn.Execute(getSalesByMod(bytMod,lngModId,"1","ASC"))
			Response.Write("<h5>" & getIDS("IDS_Sales") & "</h5><hr />")

			If objRS.EOF Then
				Response.Write(getIDS("IDS_None") & ".<hr />")
			Else
				Do while NOT objRS.EOF
					Response.Write("<b>" & getIDS("IDS_Sale") & ":</b> &nbsp;" & bigDigitNum(7,objRS.fields("SaleId").value) & "<br />" & vbCrLf & _
							"<b>" & getIDS("IDS_CloseDate") & ":</b> &nbsp;" & showDate(0,objRS.fields("S_CloseDate").value) & "<br />" & vbCrLf & _
							"<b>" & getIDS("IDS_Owner") & ":</b> &nbsp;" & showString(objRS.fields("Owner").value) & "<hr />" & vbCrLf)
					objRS.MoveNext
				Loop
			End If
		End if

		If valNum(Request.Form("chkProjects"),0,0) = 1 Then

			Set objRS = objConn.Execute(getProjectsByDiv(lngModId,"1","ASC"))
			Response.Write("<h5>" & getIDS("IDS_Projects") & "</h5><hr />")

			If objRS.EOF Then
				Response.Write(getIDS("IDS_None") & ".<hr />")
			Else
				Do while NOT objRS.EOF
					Response.Write("<b>" & getIDS("IDS_Title") & ":</b> &nbsp;" & showString(objRS.fields("P_Title").value) & "<br />" & vbCrLf & _
							"<b>" & getIDS("IDS_Owner") & ":</b> &nbsp;" & showString(objRS.fields("Owner").value) & "<br />" & vbCrLf & _
							"<b>" & getIDS("IDS_TotalDays") & ":</b> &nbsp;" & objRS.fields("P_DaysTotal").value & "<br />" & vbCrLf & _
							"<b>" & getIDS("IDS_DaysOwed") & ":</b> &nbsp;" & objRS.fields("P_DaysOwed").value & "<br />" & vbCrLf & _
							"<b>" & getIDS("IDS_CloseDate") & ":</b> &nbsp;" & showDate(0,objRS.fields("P_CloseDate").value) & "<hr />" & vbCrLf)
					objRS.MoveNext
				Loop
			End If
		End if

		If valNum(Request.Form("chkTickets"),0,0) = 1 Then

			Set objRS = objConn.Execute(getTicketsBy(bytMod,lngModId,"1","ASC"))
			Response.Write("<h5>" & getIDS("IDS_Tickets") & "</h5><hr />")

			If objRS.EOF Then
				Response.Write(getIDS("IDS_None") & ".<hr />")
			Else
				Do while NOT objRS.EOF
					Response.Write("<b>" & getIDS("IDS_Ticket") & ":</b> &nbsp;" & bigDigitNum(7,objRS.fields("TicketId").value) & "<br />" & vbCrLf & _
							"<b>" & getIDS("IDS_Owner") & ":</b> &nbsp;" & showString(objRS.fields("Owner").value) & "<br />" & vbCrLf & _
							"<b>" & getIDS("IDS_Created") & ":</b> &nbsp;" & showDate(1,objRS.fields("T_CreatedDate").value) & "<br />" & vbCrLf & _
							"<b>" & getIDS("IDS_Closed") & ":</b> &nbsp;" & showDate(1,objRS.fields("T_CloseDate").value) & "<hr />" & vbCrLf)
					objRS.MoveNext
				Loop
			End If
		End if

		If valNum(Request.Form("chkProducts"),0,0) = 1 Then

			Set objRS = objConn.Execute(getProducts(bytMod,lngModId,"1","DESC"))
			Response.Write("<h5>" & getIDS("IDS_Products") & "</h5><hr />")

			If objRS.EOF Then
				Response.Write(getIDS("IDS_None") & ".<hr />")
			Else
				Do while NOT objRS.EOF
					Response.Write("<b>" & getIDS("IDS_Product") & ":</b> &nbsp;" & showString(objRS.fields("R_Name").value) & "<br />" & vbCrLf & _
							"<b>" & getIDS("IDS_Serial") & ":</b> &nbsp;" & showString(objRS.fields("Z_Serial").value) & "<br />" & vbCrLf & _
							"<b>" & getIDS("IDS_PIN") & ":</b> &nbsp;" & showString(objRS.fields("Z_PIN").value) & "<hr />" & vbCrLf)
					objRS.MoveNext
				Loop
			End If
		End If

		If valNum(Request.Form("chkInvoices"),0,0) = 1 Then

			Set objRS = objConn.Execute(getInvoicesBy(bytMod,lngModId,"1","DESC"))
			Response.Write("<h5>" & getIDS("IDS_Invoices") & "</h5><hr />")

			If objRS.EOF Then
				Response.Write(getIDS("IDS_None") & ".<hr />")
			Else
				Do while NOT objRS.EOF
					Response.Write("<b>" & getIDS("IDS_InvoiceId") & ":</b> &nbsp;" & bigDigitNum(7,objRS.fields("InvoiceId").value) & "<br />" & vbCrLf & _
							"<b>" & getIDS("IDS_InvoiceDue") & ":</b> &nbsp;" & showDate(0,objRS.fields("I_DueDate").value) & "<br />" & vbCrLf & _
							"<b>" & getIDS("IDS_Closed") & ":</b> &nbsp;")
					If objRS.fields("I_Closed") = 1 Then Response.Write(getIDS("IDS_True")) Else Response.Write(getIDS("IDS_False"))
					Response.Write("<hr />" & vbCrLf)
					objRS.MoveNext
				Loop
			End If
		End If
	Else
%>
<body>

<div style="width:100%;height:200;margin:0px;padding:10px">

<center>
<form name="frmPrint" method="get" action="pop_print.asp">
<table border=0>
<% =getHidden("m",bytMod) %>
<% =getHidden("mid",lngModId) %>
<% =getHidden("title",showString(strTitle)) %>
  <tr>
	<td><% =getCheckbox("chkDetails","1","") %><% =getLabel(getIDS("IDS_AllDetails"),"chkDetails") %></td>
  </tr>
  <tr>
	<td><% =getCheckbox("chkNotes","","") %><% =getLabel(getIDS("IDS_Notes"),"chkNotes") %></td>
  </tr>
  <tr>
	<td><% =getCheckbox("chkAttach","","") %><% =getLabel(getIDS("IDS_Attachments"),"chkAttach") %></td>
  </tr>
  <tr>
	<td><% =getCheckbox("chkEvents","","") %><% =getLabel(getIDS("IDS_Events"),"chkEvents") %></td>
  </tr>
<% If bytMod = 2 Then %>
  <tr>
	<td><% =getCheckbox("chkContacts","","") %><% =getLabel(getIDS("IDS_Contacts"),"chkContacts") %></td>
  </tr>
<% End If

   If bytMod = 1 or bytMod = 2 or bytMod = 7 Then %>
  <tr>
	<td><% =getCheckbox("chkSales","","") %><% =getLabel(getIDS("IDS_Sales"),"chkSales") %></td>
  </tr>
  <tr>
   <td><% =getCheckbox("chkProjects","","") %><% =getLabel(getIDS("IDS_Projects"),"chkProjects") %></td>
  </tr>
<% End If

   If bytMod = 1 or bytMod = 2 or bytMod = 6 Then %>
  <tr>
	<td><% =getCheckbox("chkTickets","","") %><% =getLabel(getIDS("IDS_Tickets"),"chkTickets") %></td>
  </tr>
<% End If

   If bytMod = 1 or bytMod = 2 or bytMod = 7 Then %>
  <tr>
	<td><% =getCheckbox("chkProducts","","") %><% =getLabel(getIDS("IDS_Products"),"chkProducts") %></td>
  </tr>
  <tr>
	<td><% =getCheckbox("chkInvoices","","") %><% =getLabel(getIDS("IDS_Invoices"),"chkInvoices") %></td>
  </tr>
<% End If %>
  <tr>
	<td>
		<% =getSubmit("btnPrint",getIDS("IDS_Print"),80,"P","") %><br />
		<% =getSubmit("btnCancel",getIDS("IDS_Cancel"),80,"X","onClick=""window.close();""") %>
	</td>
  </tr>
</table>
</form>
</center>

</div>

<%
	End If

	Response.Write("</body>" & vbCrLf & "</html>")

	Call endResponse()
%>