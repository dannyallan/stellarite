<%
Sub doSendMail(fSubject, fMsg, fRecipients)
	Dim fArray, fCount, bSend, oEmail

	On Error Resume Next

	If Application("av_EnableEmail") = "0" Then Exit Sub
	bSend = False

	If fRecipients <> "" Then fArray = Split(fRecipients,",")

	If isArray(fArray) Then

		Select Case Application("av_EmailType")
			Case "1"
				Set oEmail = Server.CreateObject("CDO.message")
				oEmail.BodyFormat = 1
				o'Mail.MailFormat = 0
				oEmail.From = Application("av_EmailFrom")
				oEmail.Subject = fSubject
				oEmail.TextBody = fMsg
			Case "2"
				Set oEmail = Server.CreateObject("Persits.MailSender")
				oEmail.Host = Application("av_EmailHost")
				oEmail.Port = Application("av_EmailPort")
				oEmail.From = Application("av_EmailFrom")
				oEmail.Subject = fSubject
				oEmail.TextBody = fMsg
			Case "3"
				Set oEmail = Server.CreateObject("JMail.SMTPMail")
				oEmail.ServerAddress = Application("av_EmailHost") & ":" & Application("av_EmailPort")
				oEmail.Sender = Application("av_EmailFrom")
				oEmail.Subject = fSubject
				oEmail.TextBody = fMsg
		End Select

		For fCount = 0 to UBound(fArray)
			If valString(Trim(fArray(fCount),255,0,1)) <> "" Then
				If Application("av_EmailType") = "1" Then
					oEmail.bcc = Trim(fArray(fCount))
				Elseif Application("av_EmailType") = "2" Then
					oEmail.AddBcc = Trim(fArray(fCount))
				Elseif Application("av_EmailType") = "3" Then
					oEmail.AddRecipientBCC = Trim(fArray(fCount))
				End If
				bSend = True
			End If
		Next

		If bSend Then
			Select Case Application("av_EmailType")
				Case "1","2"
					oEmail.Send
				Case "3"
					oEmail.Execute
			End Select
		End If

		'oEmail.Close
		Set oEmail = Nothing
	End If

End Sub

Sub doNotification(fMod,fModId)
	Dim fSubject, fMsg, fRecipients
	Dim fURL, fHotIssue

	If Application("av_EnableEmail") = "0" Then Exit Sub

	Select Case fMod
		Case 1
			fSubject = getValue(doConcat(doConCat("K_FirstName","' '"),"K_LastName"),"CRM_Contacts","ContactId="&fModId,"NULL")
			fURL = "sales/contact.asp?id=" & fModId
		Case 2
			fSubject = getValue("C_Client","CRM_Clients C, CRM_Divisions D","D.ClientId=C.ClientId AND D.DivId="&fModId,"NULL")
			fURL = "sales/client.asp?id=" & fModId
		Case 3
			fSubject = getIDS("IDS_Sale") & " " & fModId
			fURL = "sales/sale.asp?id=" & fModId
		Case 4
			fSubject = fSubject & getValue("P_Title","CRM_Projects","ProjectId="&fModId,"NULL")
			fURL = "services/project.asp?id=" & fModId
		Case 5
			fSubject = getIDS("IDS_TicketId") & " " & fModId
			If getValue("T_HotIssue","CRM_Tickets","TicketId="&fModId,"NULL") = 1 Then
				fSubject = fSubject & " (" & getIDS("IDS_HotTickets") & ")"
				fHotIssue = True
			End If
			fURL = "support/ticket.asp?id=" & fModId
		Case 6
			fSubject = getIDS("IDS_BugId") & " " & fModId
			If getValue("B_HotIssue","CRM_Bugs","BugId="&fModId,"NULL") = 1 Then
				fSubject = fSubject & " (" & getIDS("IDS_HotBugs") & ")"
				fHotIssue = True
			End If
			fURL = "qa/bug.asp?id=" & fModId
		Case 7
			fSubject = getIDS("IDS_InvoiceId") & " " & fModId
			fURL = "finance/invoice.asp?id=" & fModId
	End Select

	fURL = strCRMURL & fURL

	fMsg = getIDS("IDS_MsgNotification")
	fMsg = Replace(fMsg,"[SUBJECT]",fSubject)
	fMsg = Replace(fMsg,"[MODBY]",strFullName)
	fMsg = Replace(fMsg,"[URL]",fURL)

	Set objFRS = objConn.Execute(getSubscriptionSQL(fMod,fModId,fHotIssue))

	If not (objFRS.BOF and objFRS.EOF) Then
		Do while not objFRS.EOF
			fRecipients = fRecipients & "," & objFRS.fields(0).value
			objFRS.MoveNext
		Loop

		Call doSendMail(fSubject, fMsg, fRecipients)
	End If

End Sub
%>