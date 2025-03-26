<%
Function getContacts(fType,fUser,fMax)
	getContacts = " ContactId, " & doConCat(doConCat("K_FirstName","' '"),"K_LastName") & ", " & _
				"K_ModDate, " & doConCat(doConCat("U.U_FirstName","' '"),"U.U_LastName") & " AS Owner " & _
			"FROM (CRM_Contacts K INNER JOIN ALL_Users U ON K.K_ModBy = U.UserId) " & _
			"WHERE K.K_Status = 1 "

	If fType = 0 Then getContacts = getContacts & "AND U.UserId = " & fUser & " "

	getContacts = getSelectTop(getContacts & "ORDER BY K.K_ModDate DESC, K.ContactId DESC",fMax)
End Function

Function getClients(fType,fUser,fMax)
	getClients = " D.DivId, C.C_Client, D.D_ModDate, " & doConCat(doConCat("U.U_FirstName","' '"),"U.U_LastName") & " AS Owner " & _
			"FROM CRM_Clients C, CRM_Divisions D, ALL_Users U " & _
			"WHERE D.D_ModBy = U.UserId AND D.ClientId = C.ClientId " & _
			"AND D.D_Status = 1 "

	If fType = 1 Then getClients = getClients & "AND U.UserId = " & fUser & " "

	getClients = getSelectTop(getClients & "ORDER BY D_ModDate DESC, D.DivId DESC",fMax)
End Function

Function getSales(fType,fUser,fMax)
	getSales = " S.SaleId, C.C_Client, S.S_ModDate, " & doConCat(doConCat("U.U_FirstName","' '"),"U.U_LastName") & " AS Owner " & _
			"FROM CRM_Sales S, CRM_Divisions D, CRM_Clients C, ALL_Users U " & _
			"WHERE S.S_SalesRep = U.UserId AND S.DivId = D.DivId " & _
			"AND D.ClientId = C.ClientId AND S.S_Status = 1 "
	If fType = 2 Then getSales = getSales & "AND U.UserId = " & fUser & " "

	getSales = getSelectTop(getSales & "ORDER BY S.S_ModDate DESC, S.SaleId DESC",fMax)
End Function

Function getEvents(fType,fUser,fMax)
	getEvents = " EventId, E_Module, E_ModuleId, E_Title, E_StartTime, " & doConCat(doConCat("U.U_FirstName","' '"),"U.U_LastName") & " AS Owner " & _
			" FROM (CRM_Events E INNER JOIN ALL_Users U ON E.E_Owner = U.UserId) " & _
			" WHERE E.E_Status = 1 AND E.E_StartTime >= " & Application("av_DateNow")

	If fType = 13 Then getEvents = getEvents & " AND E.E_Onsite = 1 "
	If fType = 4 Then getEvents = getEvents & " AND E.E_Owner = " & fUser & " "

	getEvents = getSelectTop(getEvents & "ORDER BY E.E_StartTime ASC, E.EventId DESC",fMax)
End Function

Function getProjects(fType,fUser)
	getProjects = "SELECT ProjectId, P_Title, P_ModDate, " & doConCat(doConCat("U.U_FirstName","' '"),"U.U_LastName") & " AS Owner " & _
			"FROM (CRM_Projects P INNER JOIN ALL_Users U ON P.P_Owner = U.UserId) " & _
			"WHERE P_Status = 1 AND P_Closed <> 1 "

	If fType = 3 Then getProjects = getProjects & "AND P_Owner = " & fUser
End Function

Function getTickets(fType,fUser)
	getTickets = "SELECT TicketId, T.DivId, C.C_Client, T_HotIssue, " & doConCat(doConCat("U.U_FirstName","' '"),"U.U_LastName") & " AS Owner " & _
			"FROM CRM_Tickets T, CRM_Divisions D, CRM_Clients C, ALL_Users U " & _
			"WHERE T.T_Owner = U.UserId AND T.DivId = D.DivId " & _
			"AND D.ClientId = C.ClientId AND T.T_Status = 1 " & _
			"AND T.T_Closed <> 1 "
	If fType = 14 Then getTickets = getTickets & "AND T.T_HotIssue = 1 "
	If fType = 5 Then getTickets = getTickets & "AND U.UserId = " & fUser & " "

	getTickets = getTickets & "ORDER BY T.T_CreatedDate ASC, T.TicketId DESC"
End Function

Function getBugs(fType,fUser)
	getBugs = "SELECT BugId, B_ModDate, B_HotIssue, " & doConCat(doConCat("U.U_FirstName","' '"),"U.U_LastName") & " AS Owner " & _
		"FROM (CRM_Bugs B INNER JOIN ALL_Users U ON B.B_Owner = U.UserId) " & _
		"WHERE B_Status = 1 AND B_Closed <> 1 "

	If fType = 16 Then getBugs = getBugs & "AND B_HotIssue = 1 "
	If fType = 6 Then getBugs = getBugs & "AND B_Owner = " & fUser & " "

	getBugs = getBugs & "ORDER BY B_CreatedDate ASC, B.BugId DESC"
End Function

Function getInvoices(fType,fUser)
	getInvoices = "SELECT I.InvoiceId, D.DivId, C.C_Client, I.I_ModDate, I.I_DueDate " & _
		"FROM ((CRM_Invoices I INNER JOIN CRM_Divisions D ON I.DivId = D.DivId) " & _
			" INNER JOIN CRM_Clients C ON D.ClientId = C.ClientId) " & _
		"WHERE I.I_Status = 1 AND I.I_Closed <> 1 "

	If fType = 18 Then getInvoices = getInvoices & "AND I.I_DueDate < " & Application("av_DateNow")
	If fType = 7 Then getInvoices = getInvoices & "AND I.I_Owner = " & fUser & " "

	getInvoices = getInvoices & " ORDER BY I_DueDate ASC, I.InvoiceId DESC"
End Function

Function getReports(fType,fUser,fMod,fMember,fPermissions)
	Dim fSQL,fCount

	fSQL = "SELECT R.ReportId, R.R_Module, R.R_Type, R.R_Title, " & doConCat(doConCat("U.U_FirstName","' '"),"U.U_LastName") & " AS Owner " & _
		"FROM (CRM_Reports R INNER JOIN ALL_Users U ON U.UserId = R.R_Owner) " & _
		"WHERE R.R_Status = 1 AND R.R_Module IN (" & fMod & ") AND "

	If fType = 0 Then fSQL = fSQL & "("

	Select Case fType
		Case 0,20
			fSQL = fSQL & "("
			For fCount = 1 to Len(fPermissions)
				If CByte(Mid(fPermissions,fCount,1)) >= 1 Then fSql = fSql & "(R.R_Module = " & fCount & " AND R.R_Permissions = 3) OR "
			Next

			For fCount = 1 to Len(fMember)
				If Mid(fMember,fCount,1) = "1" Then fSQL = fSQL & "(R.R_Module = " & fCount & " AND R.R_Permissions = 2) OR "
			Next
			fSQL = Left(fSQL,Len(fSQL)-4) & ")"
	End Select

	If fType = 0 Then fSQL = fSQL & " OR "

	Select Case fType
		Case 0,8
			fSql = fSql & "(R.R_Owner = " & fUser & " or R.R_ModBy = " & fUser & ")"
	End Select

	If fType = 0 Then fSQL = fSQL & ")"

	getReports = fSQL

End Function

Function getArticles(fUser,fMax,fCat)

	getArticles = " ka.ArticleId, ka.H_Title, ka.H_ModDate " & _
		"FROM (KB_Articles ka INNER JOIN KB_Categories kc ON ka.CatId = kc.CatId) " & _
		"WHERE ka.H_Status = 1 "

	If fCat <> "" Then getArticles = getArticles & " AND ka.CatId = " & fCat & " "

	getArticles = getArticles & " ORDER BY ka.H_ModDate DESC "

	getArticles = getSelectTop(getArticles,fMax)
End Function

Function insertPortalView(fUser,fVal)
	insertPortalView = "UPDATE ALL_Users SET U_Portal = '" & fVal & "' WHERE UserId = " & fUser
End Function
%>