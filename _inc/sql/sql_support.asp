<%
Function getTicket(fType,fId)

	Dim fPrevId, fNextId

	If fType = 1 Then
		fPrevId = doPrevNext(0,5,fId,0,0)
		fNextId = doPrevNext(1,5,fId,0,0)
	Else
		fPrevId = CLng(0)
		fNextId = CLng(0)
	End If

	getTicket = "SELECT " & fPrevId & " AS PrevId," & fNextId & " AS NextId, " & _
			"T.*, C.C_Client, D.D_Division, " & _
			doConCat(doConCat("K.K_FirstName","' '"),"K.K_LastName") & " AS Contact, " & _
			"K.K_Phone1, K.K_Ext1, K.K_Email, " & _
			doConCat(doConCat("UO.U_FirstName","' '"),"UO.U_LastName") & " AS Owner, " & _
			doConCat(doConCat("UC.U_FirstName","' '"),"UC.U_LastName") & " AS CreatedBy, " & _
			doConCat(doConCat("UM.U_FirstName","' '"),"UM.U_LastName") & " AS ModBy " & _
		"FROM ((((((CRM_Tickets T " & _
			"INNER JOIN ALL_Users UO ON T.T_Owner = UO.UserId) " & _
			"INNER JOIN ALL_Users UC ON T.T_CreatedBy = UC.UserId) " & _
			"INNER JOIN ALL_Users UM ON T.T_ModBy = UM.UserId) " & _
			"INNER JOIN CRM_Contacts K ON T.ContactId = K.ContactId) " & _
			"INNER JOIN CRM_Divisions D ON T.DivId = D.DivId) " & _
			"INNER JOIN CRM_Clients C ON D.ClientId = C.ClientId) " & _
		"WHERE T.T_Status = 1 AND T.TicketId = " & fId

End Function

Sub delTicket(fUser,fId)
	objConn.Execute("UPDATE CRM_Tickets SET T_Status = 0, T_ModBy="&fUser&", T_ModDate=" & Application("av_DateNow") & " WHERE TicketId = "&fId)
	objConn.Execute("UPDATE CRM_Contacts SET K_Tickets = K_Tickets-1 WHERE ContactId = "&getValue("ContactId","CRM_Tickets","TicketId="&fId,0))
	objConn.Execute("UPDATE CRM_Divisions SET D_Tickets = D_Tickets-1 WHERE DivId = "&getValue("DivId","CRM_Tickets","TicketId="&fId,0))
	objConn.Execute("UPDATE CRM_Bugs SET B_Tickets = B_Tickets-1 WHERE BugId = "&getValue("T_BugId","CRM_Tickets","TicketId="&fId,0))
	Call doNotification(5,fId)
End Sub

Sub updateTicket(fUser,fId,fMod,fModId,fContact,fDivision,fOwner,fHotIssue,fPriority,fTicketType,fTicketSource,fSupportType,fProduct,fBuild,fBugNum,fDescription,fSolution,fCause,fClosed,fCloseDate)
	Dim fSQL

	fSQL = "UPDATE CRM_Tickets SET " & _
		"DivId = " & fDivision & _
		",ContactId = " & fContact & _
		",T_Owner = " & fOwner & _
		",T_HotIssue = " & fHotIssue & _
		",T_Priority = " & fPriority & _
		",T_TicketType = " & fTicketType & _
		",T_TicketSource = " & fTicketSource & _
		",T_SupportType = " & fSupportType & _
		",T_ProductId = " & fProduct & _
		",T_Build = " & sqlText(fBuild) & _
		",T_BugId = " & fBugNum & _
		",T_Description = " & sqlText(fDescription) & _
		",T_Solution = " & sqlText(fSolution) & _
		",T_Cause = " & fCause & _
		",T_Closed = " & fClosed & _
		",T_CloseDate = " & sqlDate(fCloseDate) & _
		",T_ModBy = " & fUser & ",T_ModDate = " & Application("av_DateNow") &  " WHERE  TicketId = " & fId

	objConn.Execute(fSQL)
	Call doNotification(5,fId)
End Sub


Function insertTicket(fUser,fId,fMod,fModId,fContact,fDivision,fOwner,fHotIssue,fPriority,fTicketType,fTicketSource,fSupportType,fProduct,fBuild,fBugNum,fDescription,fSolution,fCause,fClosed,fCloseDate)
	Dim fSQL

	fSQL = "INSERT INTO CRM_Tickets (DivId,ContactId,T_Owner,T_HotIssue,T_Priority,T_TicketType" & _
		",T_TicketSource,T_SupportType,T_ProductId,T_Build,T_Description,T_Solution" & _
		",T_Cause,T_BugId,T_Closed,T_CloseDate,T_CreatedBy,T_CreatedDate,T_ModBy" & _
		",T_ModDate,T_Status) VALUES (" & _
		fDivision & "," & fContact & _
		"," & fOwner & "," & fHotIssue & _
		"," & fPriority & "," & fTicketType & _
		"," & fTicketSource & "," & fSupportType & _
		"," & fProduct & "," & sqlText(fBuild) & _
		"," & sqlText(fDescription) & "," & sqlText(fSolution) & _
		"," & fCause & "," & fBugNum & _
		"," & fClosed & "," & sqlDate(fCloseDate) & _
		"," & fUser & "," & Application("av_DateNow") & _
		"," & fUser & "," & Application("av_DateNow") & ",1)"

	objConn.Execute(fSQL)
	insertTicket = getLastInsert(fUser,5)
	objConn.Execute("UPDATE CRM_Contacts SET K_Tickets = K_Tickets+1 WHERE ContactId = "&fContact)
	objConn.Execute("UPDATE CRM_Divisions SET D_Tickets = D_Tickets+1 WHERE DivId = "&fDivision)
	If fBugNum <> "" Then objConn.Execute("UPDATE CRM_Bugs SET B_Tickets = B_Tickets+1 WHERE BugId = "&fBugNum)
End Function

Function getTicketsBy(fMod,fModId,fOrder,fSort)
	Dim fSQL

	fSQL = "SELECT TicketId, T_CloseDate, T_ModDate, T_CreatedDate, " & doConCat(doConCat("U.U_FirstName","' '"),"U.U_LastName") & " AS Owner " & _
		"FROM (CRM_Tickets T INNER JOIN ALL_Users U ON T.T_Owner = U.UserId) " & _
		"WHERE T_Status = 1 "

	If fMod = 2 Then
		fSQL =     fSQL & " AND DivId = " & fModId
	Elseif fMod = 1 Then
		fSQL =     fSQL & " AND ContactId = " & fModId
	Elseif fMod = 6 Then
		fSQL =     fSQL & " AND T_BugId = " & fModId
	End If

	If fOrder = "2" Then
		fSQL = fSQL & " ORDER BY T_CloseDate " & fSort
	ElseIf fOrder = "3" Then
		fSQL = fSQL & " ORDER BY T_ModDate " & fSort
	Else
		fSQL = fSQL & " ORDER BY TicketId " & fSort
	End If

	getTicketsBy = fSQL
End Function

%>