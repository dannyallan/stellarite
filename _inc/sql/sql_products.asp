<%
Function getProducts(fMod,fModId,fOrder,fSort)

	getProducts =    "SELECT Z.SerialzId, Z.Z_ModDate, R.O_Value, Z.Z_Serial, Z.Z_PIN, Z.Z_Expiry, Z.Z_ProductId, Z.InvoiceId, " & _
				doConCat(doConCat("UC.U_FirstName","' '"),"UC.U_LastName") & " AS CreatedBy, Z_CreatedDate, " & _
				doConCat(doConCat("UM.U_FirstName","' '"),"UM.U_LastName") & " AS ModBy, Z.DivId " & _
			"FROM ((((CRM_Serialz Z INNER JOIN ALL_Options R ON Z.Z_ProductId = R.OptionId) " & _
				"INNER JOIN ALL_Users UC on Z.Z_CreatedBy = UC.UserId) " & _
				"INNER JOIN ALL_Users UM on Z.Z_ModBy = UM.UserId) " & _
				"LEFT JOIN CRM_Sales S on Z.InvoiceId = S.InvoiceId) " & _
			"WHERE Z.Z_Status = 1 "

	If fMod = 0 Then getProducts = getProducts & " AND SerialzId = " & fModId
	If fMod = 1 Then getProducts = getProducts & " AND Z.ContactId = " & fModId
	If fMod = 2 Then getProducts = getProducts & " AND Z.DivId = " & fModId
	If fMod = 3 Then getProducts = getProducts & " AND S.SaleId = " & fModId
	If fMod = 7 Then getProducts = getProducts & " AND Z.InvoiceId = " & fModId

	If fOrder = "2" Then
		getProducts = getProducts & " ORDER BY R.O_Value " & fSort
	ElseIf fOrder = "3" Then
		getProducts = getProducts & " ORDER BY Z.Z_Serial " & fSort
	ElseIf fOrder = "4" Then
		getProducts = getProducts & " ORDER BY Z.Z_PIN " & fSort
	ElseIf fOrder = "5" Then
		getProducts = getProducts & " ORDER BY Z_Expiry " & fSort
	ElseIf fOrder = "1" Then
		getProducts = getProducts & " ORDER BY Z_ModDate " & fSort
	End If

End Function

Function delProduct(fUser,fId)
	objConn.Execute("UPDATE CRM_Serialz SET Z_Status = 0, Z_ModBy = " & fUser & ", Z_ModDate = " & Application("av_DateNow") & " WHERE SerialzId = " & fId)
	objConn.Execute("UPDATE CRM_Contacts SET K_Serials = K_Serials-1 WHERE ContactId = "&getValue("ContactId","CRM_Serialz","SerialzId="&fId,0))
	objConn.Execute("UPDATE CRM_Divisions SET D_Serials = D_Serials-1 WHERE DivId = "&getValue("DivId","CRM_Serialz","SerialzId="&fId,0))
	objConn.Execute("UPDATE CRM_Invoices SET I_Serials = I_Serials-1 WHERE InvoiceId = "&getValue("InvoiceId","CRM_Serialzs","SerialzId="&fId,0))
End Function

Function insertProduct(fUser,fId,fMod,fModId,fInvoiceId,fProd,fSerial,fPIN,fExpiry)
	Dim fSQL, fContactId, fDivId

	Select Case fMod
		Case 1
			fContactId = fModId
			fDivId = getValue("DivId","CRM_Contacts","ContactId="&fContactId,"NULL")
		Case 2
			fContactId = "NULL"
			fDivId = fModId
		Case 7
			fContactId = getValue("ContactId","CRM_Invoices","InvoiceId="&fModId,"NULL")
			fDivId = getValue("DivId","CRM_Invoices","InvoiceId="&fModId,"NULL")
	End Select

	fSQL = "INSERT INTO CRM_Serialz (ContactId,DivId,InvoiceId,Z_ProductId,Z_Serial,Z_PIN,Z_Expiry" & _
			",Z_CreatedBy,Z_CreatedDate,Z_ModBy,Z_ModDate,Z_Status) VALUES (" & _
			fContactId & "," & fDivId & "," & fInvoiceId & _
			"," & fProd & "," & sqlText(fSerial) & "," & sqlText(fPIN) & "," & sqlDate(fExpiry) & _
			"," & fUser & "," & Application("av_DateNow") & "," & fUser & _
			"," & Application("av_DateNow") & ",1)"

	objConn.Execute(fSQL)
	insertProduct = getLastInsert(fUser,"Z")
	objConn.Execute("UPDATE CRM_Contacts SET K_Serials = K_Serials+1 WHERE ContactId = " & fContactId)
	objConn.Execute("UPDATE CRM_Divisions SET D_Serials = D_Serials+1 WHERE DivId = " & fDivId)
	objConn.Execute("UPDATE CRM_Invoices SET I_Serials = I_Serials+1 WHERE InvoiceId = " & fInvoiceId)
End Function

Function updateProduct(fUser,fId,fMod,fModId,fInvoiceId,fProd,fSerial,fPIN,fExpiry)
	Dim fSQL, fContactId, fDivId

	Select Case fMod
		Case 1
			fContactId = fModId
			fDivId = getValue("DivId","CRM_Contacts","ContactId="&fContactId,"NULL")
		Case 2
			fContactId = "NULL"
			fDivId = fModId
		Case 7
			fContactId = getValue("ContactId","CRM_Invoices","InvoiceId="&fModId,"NULL")
			fDivId = getValue("DivId","CRM_Invoices","InvoiceId="&fModId,"NULL")
	End Select

	objConn.Execute("UPDATE CRM_Serialz SET ContactId = " & fContactId & _
				",DivId = " & fDivId & _
				",InvoiceId = " & fInvoiceId & _
				",Z_ProductId = " & fProd & _
				",Z_Serial = " & sqlText(fSerial) & _
				",Z_PIN = " & sqlText(fPIN) & _
				",Z_Expiry = " & sqlDate(fExpiry) & _
				",Z_ModBy = " & fUser & _
				",Z_ModDate = " & Application("av_DateNow") & _
			" WHERE  SerialzId = " & fId)
End Function

Function getInvoiceSQL(fDivId)

	getInvoiceSQL = "SELECT InvoiceId,InvoiceId FROM CRM_Invoices WHERE DivId = " & fDivId & " ORDER BY InvoiceId"

End Function
%>