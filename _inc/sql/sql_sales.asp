<%
Function getSale(fType,fId)

	Dim fPrevId, fNextId

	If fType = 1 Then
		fPrevId = doPrevNext(0,3,fId,0,0)
		fNextId = doPrevNext(1,3,fId,0,0)
	Else
		fPrevId = CLng(0)
		fNextId = CLng(0)
	End If

	getSale = "SELECT " & fPrevId & " AS PrevId," & fNextId & " AS NextId, S.*, " & _
			"C.C_Client, D.D_Division, " & _
			doConCat(doConCat("K.K_FirstName","' '"),"K.K_LastName") & " AS Contact, " & _
			doConCat(doConCat("US.U_FirstName","' '"),"US.U_LastName") & " AS SalesRep, " & _
			doConCat(doConCat("UC.U_FirstName","' '"),"UC.U_LastName") & " AS CreatedBy, " & _
			doConCat(doConCat("UM.U_FirstName","' '"),"UM.U_LastName") & " AS ModBy " & _
		"FROM ((((((CRM_Sales S " & _
			"LEFT JOIN CRM_Contacts K ON S.ContactId = K.ContactId) " & _
			"LEFT JOIN CRM_Divisions D ON S.DivId = D.DivId) " & _
			"LEFT JOIN CRM_Clients C ON D.ClientId = C.ClientId) " & _
			"INNER JOIN ALL_Users US ON S.S_SalesRep = US.UserId) " & _
			"INNER JOIN ALL_Users UC ON S.S_CreatedBy = UC.UserId) " & _
			"INNER JOIN ALL_Users UM ON S.S_ModBy = UM.UserId) " & _
		"WHERE S_Status = 1 AND SaleId = " & fId

End Function

Sub delSale(fUser,fId)
	objConn.Execute("UPDATE CRM_Sales SET S_Status = 0, S_ModBy="&fUser&", S_ModDate=" & Application("av_DateNow") & " WHERE SaleId = "&fId)
	objConn.Execute("UPDATE CRM_Contacts SET K_Sales = K_Sales-1 WHERE ContactId = "&getValue("ContactId","CRM_Sales","SaleId="&fId,0))
	objConn.Execute("UPDATE CRM_Divisions SET D_Sales = D_Sales-1 WHERE DivId = "&getValue("DivId","CRM_Sales","SaleId="&fId,0))
	objConn.Execute("UPDATE CRM_Invoices SET I_Sales = I_Sales-1 WHERE InvoiceId = "&getValue("InvoiceId","CRM_Sales","SaleId="&fId,0))
	Call doNotification(3,fId)
End Sub

Sub updateSale(fUser,fId,fMod,fModId,fDivId,fContactId,fPhase,fPipe,fSalesRep,fInvoice,fClosed,fCloseDate,fCurrency,fSaleValue)
	Dim fSQL

	fSQL = "UPDATE CRM_Sales SET ContactId = " & fContactId & _
			",DivId = " & fDivId & _
			",S_Phase = " & fPhase & _
			",S_Pipe = " & fPipe & _
			",S_SalesRep = " & fSalesRep & _
			",InvoiceId = " & fInvoice & _
			",S_Closed = " & fClosed & _
			",S_CloseDate = " & sqlDate(fCloseDate) & _
			",S_Currency = " & sqlText(fCurrency) & _
			",S_SaleValue = " & fSaleValue & _
			",S_ModBy = " & fUser & "," & _
			"S_ModDate = " & Application("av_DateNow") & _
		"WHERE  SaleId = " & fId

	objConn.Execute(fSQL)
	Call doNotification(3,fId)
End Sub

Function insertSale(fUser,fId,fMod,fModId,fDivId,fContactId,fPhase,fPipe,fSalesRep,fInvoice,fClosed,fCloseDate,fCurrency,fSaleValue)
	Dim fSQL

	fSQL = "INSERT INTO CRM_Sales (ContactId,DivId,S_SalesRep,S_Closed,S_CloseDate,InvoiceId" & _
		",S_Phase,S_Pipe,S_Currency,S_SaleValue,S_CreatedBy,S_CreatedDate,S_ModBy,S_ModDate,S_Status) VALUES (" & _
		fContactId & "," & fDivId & "," & fSalesRep & "," & fClosed & "," & sqlDate(fCloseDate) & _
		"," & fInvoice & "," & fPhase & "," & fPipe & "," & sqlText(fCurrency) & "," & fSaleValue & _
		"," & fUser & "," & Application("av_DateNow") & _
		"," & fUser & "," & Application("av_DateNow") & ",1)"

	objConn.Execute(fSQL)
	insertSale = getLastInsert(fUser,3)
	objConn.Execute("UPDATE CRM_Contacts SET K_Sales = K_Sales+1 WHERE ContactId = " & fContactId)
	objConn.Execute("UPDATE CRM_Divisions SET D_Sales = D_Sales+1 WHERE DivId = " & fDivId)
	objConn.Execute("UPDATE CRM_Invoices SET I_Sales = I_Sales+1 WHERE InvoiceId = " & fInvoice)
End Function

Function getSalesByMod(fMod,fModId,fOrder,fSort)
	Dim fSQL

	fSQL = "SELECT SaleId," & doConCat(doConCat("U.U_FirstName","' '"),"U.U_LastName") & " AS Owner, S_CloseDate " & _
		"FROM (CRM_Sales S INNER JOIN ALL_Users U ON S.S_SalesRep = U.UserId) WHERE S_Status = 1 "

	If fMod = 1 Then fSQL = fSQL & " AND S.ContactId = " & fModId
	If fMod = 2 Then fSQL = fSQL & " AND S.DivId = " & fModId
	If fMod = 7 Then fSQL = fSQL & " AND S.InvoiceId = " & fModId

	If fOrder = "2" Then
		getSalesByMod = fSQL & " ORDER BY U.U_FirstName " & fSort
	ElseIf fOrder = "3" Then
		getSalesByMod = fSQL & " ORDER BY S_CloseDate " & fSort
	Else
		getSalesByMod = fSQL & " ORDER BY S_ModDate " & fSort
	End If
End Function
%>