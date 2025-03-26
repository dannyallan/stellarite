<%
Function getClientId(fClientName)

	Set objFRS = objConn.Execute("SELECT ClientId FROM CRM_Clients WHERE C_Client = " & sqlText(fClientName))

	If not (objFRS.BOF and objFRS.EOF) Then
		getClientId = objFRS.fields(0).value
	Else
		objConn.Execute("INSERT INTO CRM_Clients (C_Client) VALUES (" & sqlText(fClientName) & ")")

		Set objFRS = objConn.Execute("SELECT MAX(ClientId) FROM CRM_Clients WHERE C_Client = " & sqlText(fClientName))
		getClientId = objFRS.fields(0).value
	End If
End Function

Function getDivId(fUser,fDivName,fClientId)
	Dim fSQL

	fSQL = "SELECT DivId FROM CRM_Divisions WHERE ClientId = " & fClientId
	If Len(fDivName & "") > 0 Then fSQL = fSQL & " AND D_Division = " & sqlText(fDivName)

	Set objFRS = objConn.Execute(fSQL)

	If not (objFRS.BOF and objFRS.EOF) Then
		getDivId = objFRS.fields(0).value
	Else
		fSQL =     "INSERT INTO CRM_Divisions (ClientId,D_Division,D_CreatedBy,D_CreatedDate,D_ModBy,D_ModDate,D_Status) " & _
			"VALUES (" & fClientId & "," & sqlText(fDivName) & _
			"," & fUser & "," & Application("av_DateNow") & _
			"," & fUser & "," & Application("av_DateNow") &  ",1)"

		objConn.Execute(fSQL)

		Set objFRS = objConn.Execute("SELECT MAX(DivId) FROM CRM_Divisions WHERE ClientId = " & fClientId)
		getDivId = objFRS.fields(0).value
	End If

End Function


Function getClientName(fType,fId)
	Select Case fType
		Case 0
			'For Div ID
			If fId <> "" Then getClientName = getValue("C.C_Client","CRM_Clients C, CRM_Divisions D","D.DivId=" & fId & " AND D.ClientId=C.ClientId","")
		Case 1
			'For Client ID
			If fId <> "" Then getClientName = getValue("C_Client","CRM_Client","ClientId=" & fId,"")
	End Select
End Function

Function getDivName(fDivId)
	If fDivId <> "" Then getDivName = getValue("D_Division","CRM_Divisions","DivId=" & fDivId,"")
End Function

Function getOptionSQL(fOptGroup)
	getOptionSQL = "SELECT O.OptionID, O.O_Value FROM ALL_Options O,ALL_OptGroups G " & _
			" WHERE O.O_Status = 1 AND O.OptGroupId = G.OptGroupId AND G.G_Name = " & sqlText(fOptGroup) & " " & _
			" ORDER BY O.O_Value"
End Function

Function getUserIdSQL(fVal)
	getUserIdSQL = "SELECT UserId,U_Member FROM ALL_Users Where U_UserName=" & sqlText(fVal) & " OR " & doConCat(doConCat("U_FirstName","' '"),"U_LastName") & "=" & sqlText(fVal)
End Function

Function getLastInsertSQL(fUser,fType)

	If strDatabase = "MySQL" Then
		Select Case CStr(fType)
			Case "1"
				getLastInsertSQL = "SELECT MAX(ContactId) FROM CRM_Contacts WHERE K_ModBy = " & fUser
			Case "2"
				getLastInsertSQL = "SELECT MAX(DivId) FROM CRM_Divisions WHERE D_ModBy = " & fUser
			Case "3"
				getLastInsertSQL = "SELECT MAX(SaleId) FROM CRM_Sales WHERE S_ModBy = " & fUser
			Case "4"
				getLastInsertSQL = "SELECT MAX(ProjectId) FROM CRM_Projects WHERE P_ModBy = " & fUser
			Case "5"
				getLastInsertSQL = "SELECT MAX(TicketId) FROM CRM_Tickets WHERE T_ModBy = " & fUser
			Case "6"
				getLastInsertSQL = "SELECT MAX(BugId) FROM CRM_Bugs WHERE B_ModBy = " & fUser
			Case "7"
				getLastInsertSQL = "SELECT MAX(InvoiceId) FROM CRM_Invoices WHERE I_ModBy = " & fUser
			Case "8"
				getLastInsertSQL = "SELECT MAX(ArticleId) FROM KB_Articles WHERE H_ModBy = " & fUser
			Case "50"
				getLastInsertSQL = "SELECT MAX(EventId) FROM CRM_Events WHERE E_ModBy = " & fUser
			Case "A"
				getLastInsertSQL = "SELECT MAX(AttachId) FROM CRM_Attach WHERE A_ModBy = " & fUser
			Case "C"
				getLastInsertSQL = "SELECT MAX(CatId) FROM KB_Categories"
			Case "N"
				getLastInsertSQL = "SELECT MAX(NotesId) FROM CRM_Notes WHERE N_ModBy = " & fUser
			Case "R"
				getLastInsertSQL = "SELECT MAX(ReportId) FROM CRM_Reports WHERE R_ModBy = " & fUser
			Case "U"
				getLastInsertSQL = "SELECT MAX(UserId) FROM ALL_Users WHERE U_UserName = " & fUserName
			Case "Z"
				getLastInsertSQL = "SELECT MAX(SerialzId) FROM CRM_Serialz WHERE Z_ModBy = " & fUser
		End Select
	Else
		getLastInsertSQL = "SELECT @@IDENTITY"
	End If

End Function
%>