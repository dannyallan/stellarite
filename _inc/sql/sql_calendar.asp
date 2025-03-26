<%
Function getCalendar(fType,fBill,fOff,fMine,fDateLast,fDateNext,fUser)
	Dim fSQL, fAddCols, fAddTables, fWhere

	Select Case fType
		Case 0
			fAddCols = ", '' AS Unspecified "
			fAddTables = "))"
		Case 1
			fAddCols = ", " & doConCat(doConCat("K.K_FirstName","' '"),"K.K_LastName") & " AS Contact "
			fAddTables = " LEFT JOIN CRM_Contacts K ON E.E_ModuleId = K.ContactId)) "
			fWhere = " AND E.E_Module = 1 "
		Case 2
			fAddCols = ", C_Client "
			fAddTables = " LEFT JOIN CRM_Divisions D ON E.E_ModuleId = D.DivId) " & _
						" LEFT JOIN CRM_Clients C ON D.ClientId = C.ClientId) "
			fWhere = " AND E.E_Module = 2"
		Case 3
			fAddCols = ", SaleId "
			fAddTables = " LEFT JOIN CRM_Sales S ON E.E_ModuleId = S.SaleId)) "
			fWhere = " AND E.E_Module = 3 "
		Case 4
			fAddCols = ", P_Title "
			fAddTables = " LEFT JOIN CRM_Projects P ON E.E_ModuleId = P.ProjectId)) "
			fWhere = " AND E.E_Module = 4 "
		Case 5
			fAddCols = ", TicketId "
			fAddTables = " LEFT JOIN CRM_Tickets T ON E.E_ModuleId = T.TicketId)) "
			fWhere = " AND E.E_Module = 5 "
		Case 6
			fAddCols = ", BugId "
			fAddTables = " LEFT JOIN CRM_Bugs B ON E.E_ModuleId = B.BugId)) "
			fWhere = " AND E.E_Module = 6 "
		Case 7
			fAddCols = ", InvoiceId "
			fAddTables = " LEFT JOIN CRM_Invoices I ON E.E_ModuleId = I.InvoiceId)) "
			fWhere = " AND E.E_Module = 7 "
	End Select

	fSQL =     "SELECT EventId, E_Title, E_StartTime, E_EndTime, " & doConCat(doConCat("U.U_FirstName","' '"),"U.U_LastName") & " AS Owner, " & _
			"E_Module, E_ModuleId " & fAddCols & _
		" FROM (((CRM_Events E " & _
			" INNER JOIN ALL_Users U ON E.E_Owner = U.UserId) " & fAddTables & _
		" WHERE E_Status = 1 " & _
			" AND E_EndTime > " & sqlDate(fDateLast) & _
			" AND E_StartTime < " & sqlDate(fDateNext) & fWhere

	If fBill = 0 Then fSQL = fSQL & " AND E_Billable = 1"
	If fOff = 0 Then fSQL = fSQL & " AND E_Onsite = 1"
	If fMine = 1 Then fSQL = fSQL & " AND E_Owner = " & fUser

	getCalendar = fSQL

End Function
%>