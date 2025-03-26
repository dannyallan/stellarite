<%
Sub delEvent(fUser,fId,fMod,fModId)

	objConn.Execute("UPDATE CRM_Events SET E_Status = 0,E_ModBy="&fUser&",E_ModDate=" & Application("av_DateNow") & " WHERE EventId = " & fId)
	Call updateMasterRecord("Events", fMod, fModId)
End Sub

Function getEventsBy(fMod,fModId,fMember,fOrder,fSort)

	getEventsBy =     "SELECT EventId, E_Title, E_StartTime, E_EndTime, " & _
			doConCat(doConCat("U.U_FirstName","' '"),"U.U_LastName") & " AS Owner, E_Onsite, E_Billable " & _
		"FROM (CRM_Events E INNER JOIN ALL_Users U ON E.E_Owner = U.UserId) " & _
		"WHERE  E.E_Status = 1 " & _
			" AND E.E_Module = " & fMod & _
			" AND E.E_ModuleId = " & fModId & _
			" AND E.E_Permissions >= " & fMember

	If fOrder = "2" Then
		getEventsBy = getEventsBy & " ORDER BY E.E_Title " & fSort
	ElseIf fOrder = "3" Then
		getEventsBy = getEventsBy & " ORDER BY O.O_Value " & fSort
	Else
		getEventsBy = getEventsBy & " ORDER BY E.E_StartTime " & fSort
	End If

End Function

Function getEvent(fType,fId,fMod,fModId)

	Dim fSQL, fPrevId, fNextId

	If fType = 1 Then
		fPrevId = doPrevNext(0,50,fId,fMod,fModId)
		fNextId = doPrevNext(1,50,fId,fMod,fModId)
	Else
		fPrevId = CLng(0)
		fNextId = CLng(0)
	End If

	fSQL =    "SELECT " & fPrevId & " AS PrevId," & fNextId & " AS NextId, " & _
				"E_Title, E_StartTime, E_EndTime, E_Billable, E_Onsite, E_Permissions, " & _
				doConCat(doConCat("U.U_FirstName","' '"),"U.U_LastName") & " AS Owner, E.E_EventType, " & _
				doConCat(doConCat("UC.U_FirstName","' '"),"UC.U_LastName") & " AS CreatedBy, E_CreatedDate, " & _
				doConCat(doConCat("UM.U_FirstName","' '"),"UM.U_LastName") & " AS ModBy, E_ModDate, "

	Select Case fMod
		Case 1
			fSQL = fSQL & doConCat(doConCat("K_FirstName","' '"),"K_LastName") & " as Title "
		Case 2
			fSQL = fSQL & "C.C_Client as Title "
		Case 3
			fSQL = fSQL & "S.SaleId as Title "
		Case 4
			fSQL = fSQL & "P.P_Title as Title "
		Case 5
			fSQL = fSQL & "T.TicketId as Title "
		Case 6
			fSQL = fSQL & "B.BugId as Title "
		Case Else
			fSQL = fSQL & "'' as Title "
	End Select

	fSql = fSql & "FROM (((((CRM_Events E " & _
				"INNER JOIN ALL_Users U ON E.E_Owner = U.UserId) " & _
				"INNER JOIN ALL_Users UC ON E.E_CreatedBy = UC.UserId) " & _
				"INNER JOIN ALL_Users UM ON E.E_ModBy = UM.UserId) "

	Select Case fMod
		Case 1
			fSql = fSql & "LEFT JOIN CRM_Contacts K ON E.E_ModuleId = K.ContactId)) "
		Case 2
			fSql = fSql & "LEFT JOIN CRM_Divisions D ON E.E_ModuleId = D.DivId) " & _
						"LEFT JOIN CRM_Clients C ON D.ClientId = C.ClientId) "
		Case 3
			fSql = fSql & "LEFT JOIN CRM_Sales S ON E.E_ModuleId = S.SaleId)) "
		Case 4
			fSql = fSql & "LEFT JOIN CRM_Projects P ON E.E_ModuleId = P.ProjectId)) "
		Case 5
			fSql = fSql & "LEFT JOIN CRM_Tickets T ON E.E_ModuleId = T.TicketId)) "
		Case 6
			fSql = fSql & "LEFT JOIN CRM_Bugs B ON E.E_ModuleId = B.BugId)) "
		Case Else
			fSql = fSQL & ")) "
	End Select

	getEvent = fSql & " WHERE E_Status = 1 AND EventId = " & fId

End Function

Function updateEvent(fUser,fId,fMod,fModId,fOwner,fOnsite,fBill,fEvent,fPerm,fTitle,fStart,fEnd)

	objConn.Execute("UPDATE CRM_Events SET " & _
				"E_Module = " & fMod & _
				",E_ModuleId = " & fModId & _
				",E_Owner = " & fOwner & _
				",E_Onsite = " & fOnsite & _
				",E_Billable = " & fBill & _
				",E_EventType = " & fEvent & _
				",E_Permissions = " & fPerm & _
				",E_Title = " & sqlText(fTitle) & _
				",E_StartTime = " & sqlDate(fStart) & _
				",E_EndTime = " & sqlDate(fEnd) & _
				",E_ModBy = " & fUser & _
				",E_ModDate = " & Application("av_DateNow") & _
			"WHERE  EventId = " & fId)

	Call updateMasterRecord("Events", fMod, fModId)
End Function

Function insertEvent(fUser,fMod,fModId,fOwner,fOnsite,fBill,fEvent,fPerm,fTitle,fStart,fEnd)

	objConn.Execute("INSERT INTO CRM_Events (E_Module,E_ModuleId,E_Onsite,E_Billable,E_EventType,E_Permissions,E_Title" & _
			",E_Owner,E_StartTime,E_EndTime,E_CreatedBy,E_CreatedDate,E_ModBy,E_ModDate,E_Status) VALUES (" & _
			fMod & "," & fModId & "," & fOnsite & _
			"," & fBill & "," & fEvent & "," & fPerm & _
			"," & sqlText(fTitle) & "," & fOwner & "," & sqlDate(fStart) & _
			"," & sqlDate(fEnd) & "," & fUser & _
			"," & Application("av_DateNow") &  "," & fUser & "," & Application("av_DateNow") & ",1)")

	insertEvent = getLastInsert(fUser,50)
	Call updateMasterRecord("Events", fMod, fModId)
End Function
%>