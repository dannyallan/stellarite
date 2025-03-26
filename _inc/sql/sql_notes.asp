<%
Function getNotesBy(fMod,fModId,fEvent,fMember,fOrder,fSort)

	getNotesBy =     "SELECT NotesId, N_ContactType, O.O_Value AS ContactType, N_Info, N_Permissions, " & _
			doConCat(doConCat("UC.U_FirstName","' '"),"UC.U_LastName") & " AS CreatedBy, N_CreatedDate, " & _
			doConCat(doConCat("UM.U_FirstName","' '"),"UM.U_LastName") & " AS ModBy, N_ModDate " & _
		"FROM (((CRM_Notes N " & _
			"INNER JOIN ALL_Users UM ON N.N_ModBy = UM.UserId) " & _
			"INNER JOIN ALL_Users UC ON N.N_CreatedBy = UC.UserId) " & _
			"LEFT JOIN ALL_Options O ON N.N_ContactType = O.OptionId) " & _
		"WHERE N.N_Module = " & fMod & _
			" AND N.N_ModuleId = " & fModId & _
			" AND N.N_Status = 1 " & _
			" AND N.N_Permissions >= " & fMember
		If fEvent <> 0 Then getNotesBy = getNotesBy & " AND N.N_EventId = " & fEvent

	If fOrder = "2" Then
		getNotesBy = getNotesBy & " ORDER BY O.O_Value " & fSort
	ElseIf fOrder = "3" Then
		getNotesBy = getNotesBy & " ORDER BY UM.U_FirstName " & fSort
	Else
		getNotesBy = getNotesBy & " ORDER BY N_ModDate " & fSort
	End If

End Function



Sub delNote(fUser,fId,fMod,fModId)

	objConn.Execute("UPDATE CRM_Notes SET N_Status = 0,N_ModBy ="&fUser&",N_ModDate=" & Application("av_DateNow") & " WHERE NotesId = " & fId)
	Call updateMasterRecord("Notes",fMod,fModId)
End Sub

Sub updateNote(fUser,fId,fMod,fModId,fInfo,fType,fPerm)

	objConn.Execute("UPDATE CRM_Notes SET " & _
			"N_Info = " & sqlText(fInfo) & _
			",N_ContactType = " & fType & _
			",N_Permissions = " & fPerm & _
			",N_ModBy = " & fUser & _
			",N_ModDate = " & Application("av_DateNow") &  _
			" WHERE NotesId = " & fId)

	Call updateMasterRecord("Notes",fMod,fModId)

End Sub

Function insertNote(fUser,fId,fMod,fModId,fInfo,fType,fPerm,fEvent)

	objConn.Execute("INSERT INTO CRM_Notes (N_Module,N_ModuleId,N_EventId,N_Info,N_CreatedBy," & _
			"N_CreatedDate,N_ModBy,N_ModDate,N_ContactType,N_Permissions,N_Status) VALUES (" & _
			fMod & "," & fModId  & _
			"," & fEvent & "," & sqlText(fInfo) & _
			"," & fUser & "," & Application("av_DateNow") & _
			"," & fUser & "," & Application("av_DateNow") &  _
			"," & fType & "," & fPerm & ",1)")

	insertNote = getLastInsert(fUser,"N")
	Call updateMasterRecord("Notes",fMod,fModId)
End Function

Function getNote(fId,fPerm)
	getNote = "SELECT N.*, " & doConCat(doConCat("UC.U_FirstName","' '"),"UC.U_LastName") & " AS CreatedBy, " & doConCat(doConCat("UM.U_FirstName","' '"),"UM.U_LastName") & " AS ModBy " & _
		"FROM ((CRM_Notes N " & _
			"INNER JOIN ALL_Users UC ON UC.UserId = N.N_CreatedBy) " & _
			"INNER JOIN ALL_Users UM ON UM.UserId = N.N_ModBy) " & _
		"WHERE NotesId = " & fId & " AND N_Permissions >= " & fPerm
End Function
%>