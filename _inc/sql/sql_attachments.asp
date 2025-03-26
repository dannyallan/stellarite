<%
Function getAttachBy(fMod,fModId,fEvent,fMember,fOrder,fSort)

	getAttachBy =     "SELECT A.AttachId, A.A_Title, O.O_Value, A.A_Info, " & _
			"A.A_CreatedDate, " & doConCat(doConCat("UC.U_FirstName","' '"),"UC.U_LastName") & " AS CreatedBy, " & _
			"A.A_ModDate, " & doConCat(doConCat("UM.U_FirstName","' '"),"UM.U_LastName") & " AS ModBy " & _
		"FROM (((CRM_Attach A " & _
			"INNER JOIN ALL_Users UM ON A.A_ModBy = UM.UserId) " & _
			"INNER JOIN All_Users UC ON A.A_CreatedBy = UC.UserId) " & _
			"LEFT JOIN ALL_Options O ON A.A_DocType = O.OptionId) " & _
		"WHERE A.A_Status = 1 " & _
			" AND A.A_Module = " & fMod & _
			" AND A.A_ModuleId = " & fModId & _
			" AND A.A_Permissions >= " & fMember
		If fEvent <> 0 Then getAttachBy = getAttachBy & " AND A.A_EventId = " & fEvent

	If fOrder = "2" Then
		getAttachBy = getAttachBy & " ORDER BY O.O_Value " & fSort
	ElseIf fOrder = "3" Then
		getAttachBy = getAttachBy & " ORDER BY A.A_Title " & fSort
	ElseIf fOrder = "4" Then
		getAttachBy = getAttachBy & " ORDER BY UM.U_FirstName " & fSort
	Else
		getAttachBy = getAttachBy & " ORDER BY A.A_ModDate " & fSort
	End If

End Function

Function getAttachLinks(fId)
	getAttachLinks = "SELECT L_URL FROM CRM_Links WHERE AttachId = " & fId
End Function

Sub delAttach(fUser,fId,fMod,fModId)

	objConn.Execute("UPDATE CRM_Attach SET A_Status = 0" & _
			",A_ModBy = " & fUser & _
			",A_ModDate = " & Application("av_DateNow") & _
			" WHERE AttachId = " & fId)

	Call updateMasterRecord("Attach", fMod, fModId)
End Sub

Function insertAttach(fUser,fMod,fModId,fEvent,fType,fPerm,fTitle,fInfo)

	objConn.Execute("INSERT INTO CRM_Attach (A_Module,A_ModuleId,A_EventId,A_CreatedBy,A_CreatedDate,A_ModBy,A_ModDate" & _
			",A_DocType,A_Permissions,A_Title,A_Info,A_Status) VALUES (" & _
			fMod & "," & fModId & "," & fEvent & "," & fUser & _
			"," & Application("av_DateNow") & "," & fUser & _
			"," & Application("av_DateNow") & "," & fType & _
			"," & fPerm & "," & sqlText(fTitle) & "," & sqlText(fInfo) & ",1)")

	insertAttach = getLastInsert(fUser,"A")
	Call updateMasterRecord("Attach", fMod, fModId)
End Function

Function insertAttachLinks(fId,fURL)

	insertAttachLinks = "INSERT INTO CRM_Links (AttachId,L_URL) VALUES (" & fId & "," & sqlText(fURL) & ")"

End Function

Function getAttach(fId,fPerm)

	getAttach = "SELECT A.*, " & doConCat(doConCat("UC.U_FirstName","' '"),"UC.U_LastName") & " AS CreatedBy, " & doConCat(doConCat("UM.U_FirstName","' '"),"UM.U_LastName") & " AS ModBy " & _
			"FROM ((CRM_Attach A " & _
				"INNER JOIN ALL_Users UC ON UC.UserId = A.A_CreatedBy) " & _
				"INNER JOIN ALL_Users UM ON UM.UserId = A.A_ModBy) " & _
			"WHERE AttachId = " & fId & " AND A_Permissions >= " & fPerm
End Function
%>