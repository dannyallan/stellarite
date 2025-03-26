<%
Function getSummary(fMod,fModId)

	Select Case fMod
		Case 1
			getSummary = "SELECT K_CreatedDate, " & doConCat(doConCat("UC.U_FirstName","' '"),"UC.U_LastName") & " AS CreatedBy, " & _
					"K_ModDate, " & doConCat(doConCat("UM.U_FirstName","' '"),"UM.U_LastName") & " AS ModBy, " & _
					"K_Notes, K_Attach, K_Events, 0, K_Sales, K_Serials, D_Projects, K_Tickets, K_Invoices, NULL " & _
				"FROM CRM_Contacts K, CRM_Divisions D, ALL_Users UC, ALL_Users UM " & _
				"WHERE UC.UserId = K.K_CreatedBy " & _
					"AND UM.UserId = K.K_ModBy " & _
					"AND K.DivId = D.DivId AND K.ContactId = " & fModId
		Case 2
			getSummary = "SELECT D_CreatedDate, " & doConCat(doConCat("UC.U_FirstName","' '"),"UC.U_LastName") & " AS CreatedBy, " & _
					"D_ModDate, " & doConCat(doConCat("UM.U_FirstName","' '"),"UM.U_LastName") & " AS ModBy, " & _
					"D_Notes, D_Attach, D_Events, D_Contacts, D_Sales, D_Serials, D_Projects, D_Tickets, D_Invoices, D_ShortDesc " & _
				"FROM CRM_Divisions D, ALL_Users UC, ALL_Users UM " & _
				"WHERE UC.UserId = D.D_CreatedBy " & _
					"AND UM.UserId = D.D_ModBy " & _
					"AND DivId = " & fModId
		Case 3
			getSummary = "SELECT S_CreatedDate, " & doConCat(doConCat("UC.U_FirstName","' '"),"UC.U_LastName") & " AS CreatedBy, " & _
					"S_ModDate, " & doConCat(doConCat("UM.U_FirstName","' '"),"UM.U_LastName") & " AS ModBy, " & _
					"S_Notes, S_Attach, S_Events, 0, 0, 0, 0, 0, 0, NULL " & _
				"FROM ((CRM_Sales S " & _
					"INNER JOIN ALL_Users UC ON UC.UserId = S.S_CreatedBy) " & _
					"INNER JOIN ALL_Users UM ON UM.UserId = S.S_ModBy) " & _
				"WHERE SaleId = " & fModId
		Case Else
			Call logError(1,1)
	End Select

End Function
%>