<%
Function getInvoice(fType,fId)

	Dim fPrevId, fNextId

	If fType = 1 Then
		fPrevId = CLng(doPrevNext(0,7,fId,0,0))
		fNextId = CLng(doPrevNext(1,7,fId,0,0))
	Else
		fPrevId = CLng(0)
		fNextId = CLng(0)
	End If

	getInvoice = "SELECT " & fPrevId & " AS PrevId," & fNextId & " AS NextId, " & _
			"I.*, C.C_Client, D.D_Division, " & _
			doConCat(doConCat("K.K_FirstName","' '"),"K.K_LastName") & " AS Contact, " & _
			doConCat(doConCat("UO.U_FirstName","' '"),"UO.U_LastName") & " AS Owner, " & _
			doConCat(doConCat("UC.U_FirstName","' '"),"UC.U_LastName") & " AS CreatedBy, " & _
			doConCat(doConCat("UM.U_FirstName","' '"),"UM.U_LastName") & " AS ModBy " & _
		"FROM ((((((CRM_Invoices I " & _
			"INNER JOIN ALL_Users UO ON I.I_Owner = UO.UserId) " & _
			"INNER JOIN ALL_Users UC ON I.I_CreatedBy = UC.UserId) " & _
			"INNER JOIN ALL_Users UM ON I.I_ModBy = UM.UserId) " & _
			"INNER JOIN CRM_Contacts K ON I.ContactId = K.ContactId) " & _
			"INNER JOIN CRM_Divisions D ON I.DivId = D.DivId) " & _
			"INNER JOIN CRM_Clients C ON D.ClientId = C.ClientId) " & _
		"WHERE I.I_Status = 1 AND I.InvoiceId = " & fId

End Function

Sub delInvoice(fUser,fId)
	objConn.Execute("UPDATE CRM_Invoices SET I_Status = 0, I_ModBy="&fUser&", I_ModDate=" & Application("av_DateNow") & " WHERE InvoiceId = "&fId)
	objConn.Execute("UPDATE CRM_Contacts SET K_Invoices = K_Invoices-1 WHERE ContactId = "&getValue("ContactId","CRM_Invoices","InvoiceId="&fId,0))
	objConn.Execute("UPDATE CRM_Divisions SET D_Invoices = D_Invoices-1 WHERE DivId = "&getValue("DivId","CRM_Invoices","InvoiceId="&fId,0))
	Call doNotification(7,fId)
End Sub

Sub updateInvoice(fUserId,fRecordId,fContactId,fDivId,fOwner,fPO,fClosed,fReceived,fType,fPhase,fCurrency,fValue,fTax,fInfo,fSendDate,fDueDate,fPaidDate)
	Dim fSQL

	fSQL = "UPDATE CRM_Invoices SET " & _
		"DivId = " & fDivId & _
		",ContactId = " & fContactId & _
		",I_Owner = " & fOwner & _
		",I_PurchaseOrder = " & sqlText(fPO) & _
		",I_Received = " & fReceived & _
		",I_Type = " & fType & _
		",I_Phase = " & fPhase & _
		",I_Currency = " & sqlText(fCurrency) & _
		",I_Value = " & fValue & _
		",I_Tax = " & fTax & _
		",I_PayInfo = " & sqlText(fInfo) & _
		",I_SendDate = " & sqlDate(fSendDate) & _
		",I_DueDate = " & sqlDate(fDueDate) & _
		",I_PaidDate = " & sqlDate(fPaidDate) & _
		",I_Closed = " & fClosed & _
		",I_ModBy = " & fUserId & ",I_ModDate = " & Application("av_DateNow") &  " WHERE  InvoiceId = " & fRecordId

	objConn.Execute(fSQL)
	Call doNotification(7,fRecordId)
End Sub

Function insertInvoice(fUserId,fRecordId,fContactId,fDivId,fOwner,fPO,fClosed,fReceived,fType,fPhase,fCurrency,fValue,fTax,fInfo,fSendDate,fDueDate,fPaidDate)
	Dim fSQL

	fSQL = "INSERT INTO CRM_Invoices (DivId,ContactId,I_Owner,I_PurchaseOrder,I_Received" & _
		",I_Type,I_Phase,I_Currency,I_Value,I_Tax,I_PayInfo,I_SendDate,I_DueDate,I_PaidDate" & _
		",I_Closed,I_CreatedBy,I_CreatedDate,I_ModBy,I_ModDate,I_Status) VALUES (" & _
		fDivId & "," & fContactId & _
		"," & fOwner & "," & sqlText(fPO) & _
		"," & fReceived & "," & fType & _
		"," & fPhase & "," & sqlText(fCurrency) & _
		"," & fValue & ","& fTax & "," & sqlText(fInfo) & _
		"," & sqlDate(fSendDate) & "," & sqlDate(fDueDate) & _
		"," & sqlDate(fPaidDate) & "," & fClosed & _
		"," & fUserId & "," & Application("av_DateNow") & _
		"," & fUserId & "," & Application("av_DateNow") & ",1)"

	objConn.Execute(fSQL)
	insertInvoice = getLastInsert(fUserId,7)
	objConn.Execute("UPDATE CRM_Contacts SET K_Invoices = K_Invoices+1 WHERE ContactId = "&fContactId)
	objConn.Execute("UPDATE CRM_Divisions SET D_Invoices = D_Invoices+1 WHERE DivId = "&fDivId)
End Function

Function getInvoicesBy(fMod,fModId,fOrder,fSort)
	Dim fSQL

	fSQL = "SELECT InvoiceId, I_DueDate, I_Closed " & _
		"FROM CRM_Invoices WHERE I_Status = 1 "

	If fMod = 2 Then
		fSQL =     fSQL & " AND DivId = " & fModId
	Elseif fMod = 1 Then
		fSQL =     fSQL & " AND ContactId = " & fModId
	End If

	If fOrder = "2" Then
		fSQL = fSQL & " ORDER BY I_DueDate " & fSort
	ElseIf fOrder = "3" Then
		fSQL = fSQL & " ORDER BY I_Closed " & fSort
	Else
		fSQL = fSQL & " ORDER BY InvoiceId " & fSort
	End If

	getInvoicesBy = fSQL
End Function
%>