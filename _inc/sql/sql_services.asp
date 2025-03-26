<%
Function getProject(fType,fId)

	Dim fPrevId, fNextId

	If fType = 1 Then
		fPrevId = doPrevNext(0,4,fId,0,0)
		fNextId = doPrevNext(1,4,fId,0,0)
	Else
		fPrevId = CLng(0)
		fNextId = CLng(0)
	End If

	getProject = "SELECT " & fPrevId & " AS PrevId," & fNextId & " AS NextId, " & _
			"P.*, C.C_Client, D.D_Division, " & _
			doConCat(doConCat("UO.U_FirstName","' '"),"UO.U_LastName") & " AS Owner, " & _
			doConCat(doConCat("UC.U_FirstName","' '"),"UC.U_LastName") & " AS CreatedBy, " & _
			doConCat(doConCat("UM.U_FirstName","' '"),"UM.U_LastName") & " AS ModBy " & _
		"FROM (((((CRM_Projects P " & _
			"INNER JOIN CRM_Divisions D ON P.DivId = D.DivId) " & _
			"INNER JOIN CRM_Clients C ON D.ClientId = C.ClientId) " & _
			"INNER JOIN ALL_Users UO ON P.P_Owner = UO.UserId) " & _
			"INNER JOIN ALL_Users UC ON P.P_CreatedBy = UC.UserId) " & _
			"INNER JOIN ALL_Users UM ON P.P_ModBy = UM.UserId) " & _
		"WHERE P_Status = 1 AND ProjectId = " & fId
End Function

Sub delProject(fUser,fId)
	objConn.Execute("UPDATE CRM_Projects SET P_Status = 0, P_ModBy="&fUser&", P_ModDate=" & Application("av_DateNow") & " WHERE ProjectId = "&fId)
	objConn.Execute("UPDATE CRM_Divisions SET D_Projects = D_Projects-1 WHERE DivId = "&getValue("DivId","CRM_Projects","ProjectId="&fId,0))
	objConn.Execute("UPDATE CRM_Invoices SET I_Projects = I_Projects-1 WHERE InvoiceId = "&getValue("InvoiceId","CRM_Projects","ProjectId="&fId,0))
	Call doNotification(4,fId)
End Sub

Sub updateProject(fUser,fId,fMod,fModId,fProject,fOwner,fCloseDate,fInvoice,fDaysTotal,fDaysOwed,fClosed,fClient,fDivision,fShortDesc)
	Dim fSQL, fDivId

	Select Case fMod
		Case 2
			fDivId = fModId
		Case 4
			fDivId = getDivId(fUser,fDivision,getClientId(fClient))
		Case 7
			fDivId = getValue("DivId","CRM_Invoices","InvoiceId="&fModId,"0")
			fInvoice = fModId
	End Select

	fSQL = "UPDATE CRM_Projects SET " & _
		"P_Title = " & sqlText(fProject) & _
		",DivId = " & fDivId & _
		",InvoiceId = " & fInvoice & _
		",P_Owner = " & fOwner & _
		",P_ShortDesc = " & sqlText(fShortDesc) & _
		",P_DaysTotal = " & fDaysTotal & _
		",P_DaysOwed = " & fDaysOwed & _
		",P_Closed = " & fClosed & _
		",P_CloseDate = " & sqlDate(fCloseDate) & _
		",P_ModBy = " & fUser & _
		",P_ModDate = " & Application("av_DateNow") &  " WHERE ProjectId = " & fId

	objConn.Execute(fSQL)
	Call doNotification(4,fId)
End Sub

Function insertProject(fUser,fId,fMod,fModId,fProject,fOwner,fCloseDate,fInvoice,fDaysTotal,fDaysOwed,fClosed,fClient,fDivision,fShortDesc)
	Dim fSQL, fDivId

	Select Case fMod
		Case 2
			fDivId = fModId
		Case 4
			fDivId = getDivId(fUser,fDivision,getClientId(fClient))
		Case 7
			fDivId = getValue("DivId","CRM_Invoices","InvoiceId="&fModId,"0")
			fInvoice = fModId
	End Select

	fSQL = "INSERT INTO CRM_Projects (P_Title,DivId,InvoiceId,P_Owner,P_ShortDesc,P_DaysTotal,P_DaysOwed" & _
		",P_Closed,P_CloseDate,P_CreatedBy,P_CreatedDate,P_ModBy,P_ModDate,P_Status) VALUES (" & _
		"" & sqlText(fProject) & "," & fDivId & _
		"," & fInvoice & "," & fOwner & "," & sqlText(fShortDesc) & _
		"," & fDaysTotal & "," & fDaysOwed & _
		"," & fClosed & "," & sqlDate(fCloseDate) & _
		"," & fUser & "," & Application("av_DateNow") & _
		"," & fUser & "," & Application("av_DateNow") &  ",1)"

	objConn.Execute(fSQL)
	insertProject = getLastInsert(fUser,4)
	objConn.Execute("UPDATE CRM_Divisions SET D_Projects = D_Projects+1 WHERE DivId = " & fDivId)
	objConn.Execute("UPDATE CRM_Invoices SET I_Projects = I_Projects+1 WHERE InvoiceId = " & fInvoice)
End Function

Function getProjectsByDiv(fMod,fModId,fOrder,fSort)
	Dim fSQL

	fSQL = "SELECT ProjectId,P_Title,P_CloseDate,P_ModDate,P_DaysTotal,P_DaysOwed," & doConCat(doConCat("U.U_FirstName","' '"),"U.U_LastName") & " AS Owner " & _
		"FROM (CRM_Projects P LEFT JOIN ALL_Users U ON P.P_Owner = U.UserId) " & _
		"WHERE P_Status = 1 "

	If fMod = 2 Then fSQL = fSQL & " AND DivId = " & fModId
	If fMod = 7 Then fSQL = fSQL & " AND InvoiceId = " & fModId

	If fOrder = "2" Then
		getProjectsByDiv = fSQL & " ORDER BY P_CloseDate " & fSort
	ElseIf fOrder = "3" Then
		getProjectsByDiv = fSQL & " ORDER BY P_ModDate " & fSort
	Else
		getProjectsByDiv = fSQL & " ORDER BY P_Title " & fSort
	End If
End Function
%>