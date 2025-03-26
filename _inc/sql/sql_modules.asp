<%
Function doPrevNext(fWhich,fType,fId,fMod,fModId)

	Dim fSql, fUpDown, fMinMax, fRS

	If fWhich = 0 Then
		fUpDown = "<"
		fMinMax = "MAX"
	Else
		fUpDown = ">"
		fMinMax = "MIN"
	End If

	Select Case CStr(fType)
		Case "1"
			fSql = "SELECT " & fMinMax & "(ContactId) FROM CRM_Contacts WHERE ContactId " & fUpDown & fId & " AND K_Status = 1"
		Case "2"
			fSql = "SELECT " & fMinMax & "(DivId) FROM CRM_Divisions WHERE DivId " & fUpDown & fId & " AND D_Status = 1"
		Case "3"
			fSql = "SELECT " & fMinMax & "(SaleId) FROM CRM_Sales WHERE SaleId " & fUpDown & fId & " AND S_Status = 1"
		Case "4"
			fSql = "SELECT " & fMinMax & "(ProjectId) FROM CRM_Projects WHERE  ProjectId " & fUpDown & fId & " AND P_Status = 1"
		Case "5"
			fSql = "SELECT " & fMinMax & "(TicketId) FROM CRM_Tickets WHERE  TicketId " & fUpDown & fId & " AND T_Status = 1"
		Case "6"
			fSql = "SELECT " & fMinMax & "(BugId) FROM CRM_Bugs WHERE BugId " & fUpDown & fId & " AND B_Status = 1"
		Case "7"
			fSql = "SELECT " & fMinMax & "(InvoiceId) FROM CRM_Invoices WHERE InvoiceId " & fUpDown & fId & " AND I_Status = 1"
		Case "50"
			fSql = "SELECT " & fMinMax & "(EventId) FROM CRM_Events WHERE EventId " & fUpDown & fId & " AND E_Status = 1 and E_Permissions >= " & intMember
			If fMod <> "" and fModId <> "" Then fSQL = fSQL & " AND E_Module = " & fMod & " AND E_ModuleId = " & fModId
	End Select

	Set fRS = objConn.Execute(fSQL)

	If not IsNull(fRS.fields(0).value) Then doPrevNext = fRS.fields(0).value Else doPrevNext = 0

	fRS.Close
	Set fRS = Nothing

End Function
%>