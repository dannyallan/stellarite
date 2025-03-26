<%
Function getSearch(fMod,fThird,fField,fValue,fMax)
	Dim fTemp, fArray, fMatch

	fMatch = sqlLike(fValue)

	Select Case fMod
		Case 1
			fArray = split(doConCat(doConCat("K_FirstName","' '"),"K_LastName") & "|K_LastName|K_FirstName|C_Client|D_Division|K_City|K_State|K_Country|K_ZIP|K_Email|K_Phone1","|")

			getSearch = " ContactId," & doConCat(doConCat("K_FirstName","' '"),"K_LastName") & ",C.C_Client,K.DivId," & fArray(fThird) & " " & _
				"FROM ((CRM_Contacts K " & _
					"INNER JOIN CRM_Divisions D ON K.DivId = D.DivId) " & _
					"INNER JOIN CRM_Clients C ON D.ClientId = C.ClientId) " & _
				"WHERE K_Status = 1 AND " & fArray(fField) & " LIKE " & fMatch & " ORDER BY " & fArray(fField)

		Case 2
			fArray = split("C_Client|D_Division|D_Account|" & doConCat(doConCat("U_FirstName","' '"),"U_LastName") & "|D_ShortDesc","|")

			getSearch = " D.DivId, C_Client, D_Division," & doConCat(doConCat("U_FirstName","' '"),"U_LastName") & "," & fArray(fThird) & " " & _
				"FROM ((CRM_Divisions D " & _
					"INNER JOIN CRM_Clients C ON D.ClientId = C.ClientId) " & _
					"LEFT JOIN ALL_Users U ON D.D_SalesRep = U.UserId) " & _
				"WHERE D.D_Status = 1 AND " & fArray(fField) & " LIKE " & fMatch & " ORDER BY " & fArray(fField)

		Case 3
			If fField = 3 Then fTemp = "INNER" else fTemp = "LEFT"
			fArray = split("SaleId|C_Client|D_Division|" & doConCat(doConCat("U_FirstName","' '"),"U_LastName"),"|")

			getSearch = " SaleId,C.C_Client," & doConCat(doConCat("U_FirstName","' '"),"U_LastName") & "," & fArray(fThird) & " " & _
				"FROM (((CRM_Sales S " & _
					"INNER JOIN CRM_Divisions D ON S.DivId = D.DivId) " & _
					"INNER JOIN CRM_Clients C ON D.ClientId = C.ClientId) " & _
					fTemp & " JOIN ALL_Users U ON S.S_SalesRep = U.UserId) " & _
				"WHERE S_Status = 1 AND " & fArray(fField) & " LIKE " & fMatch & " ORDER BY " & fArray(fField)

		Case 4
			fArray = split("P_Title|C_Client|D_Division|" & doConCat(doConCat("U_FirstName","' '"),"U_LastName") & "|P_ShortDesc","|")

			getSearch = " ProjectId, P_Title,C.C_Client," & doConCat(doConCat("U_FirstName","' '"),"U_LastName") & "," & fArray(fThird) & " " & _
				"FROM (((CRM_Projects P " & _
					"INNER JOIN CRM_Divisions D ON P.DivId = D.DivId) " & _
					"INNER JOIN CRM_Clients C ON D.ClientId = C.ClientId) " & _
					"LEFT JOIN ALL_Users U ON P.P_Owner = U.UserId) " & _
				"WHERE P_Status = 1 AND " & fArray(fField) & " LIKE " & fMatch & " ORDER BY " & fArray(fField)

		Case 5
			fArray = split("TicketId|C_Client|D_Division|" & doConCat(doConCat("K_FirstName","' '"),"K_LastName") & "|" & doConCat(doConCat("U_FirstName","' '"),"U_LastName") & "|O_Value|T_Description|T_Solution","|")

			getSearch = " TicketId,C.C_Client,O_Value," & fArray(fThird) & " " & _
				"FROM (((((CRM_Tickets T " & _
					"INNER JOIN CRM_Divisions D ON T.DivId = D.DivId) " & _
					"INNER JOIN CRM_Clients C ON D.ClientId = C.ClientId) " & _
					"INNER JOIN CRM_Contacts K ON T.ContactId = K.ContactId) " & _
					"INNER JOIN ALL_Users U ON T.T_Owner = U.UserId) " & _
					"LEFT JOIN ALL_Options O ON T.T_Priority = O.OptionId)" & _
				"WHERE T_Status = 1 AND " & fArray(fField) & " LIKE " & fMatch & " ORDER BY " & fArray(fField)

		Case 6
			fArray = split("BugId|" & doConCat(doConCat("U_FirstName","' '"),"U_LastName") & "|O_Value|B_Description|B_Solution","|")

			getSearch = " BugId," & doConCat(doConCat("U_FirstName","' '"),"U_LastName") & "," & fArray(fThird) & " " & _
				"FROM ((CRM_Bugs B " & _
					"INNER JOIN ALL_Users U ON B.B_Owner = U.UserId) " & _
					"LEFT JOIN ALL_Options O ON B.B_Priority = O.OptionId) " & _
				"WHERE B_Status = 1 AND " & fArray(fField) & " LIKE " & fMatch & " ORDER BY " & fArray(fField)

		Case 7
			If fField = 3 Then fTemp = "INNER" else fTemp = "LEFT"
			fArray = split("InvoiceId|C_Client|D_Division|" & doConCat(doConCat("K_FirstName","' '"),"K_LastName") & "|" & doConCat(doConCat("U_FirstName","' '"),"U_LastName"),"|")

			getSearch = " InvoiceId, C.C_Client," & fArray(fThird) & " " & _
				"FROM ((((CRM_Invoices I " & _
					"INNER JOIN CRM_Divisions D ON I.DivId = D.DivId) " & _
					"INNER JOIN CRM_Clients C ON D.ClientId = C.ClientId) " & _
					"INNER JOIN CRM_Contacts K ON I.ContactId = K.ContactId) " & _
					fTemp & " JOIN ALL_Users U ON I.I_Owner = U.UserId) " & _
				"WHERE I_Status = 1 AND " & fArray(fField) & " LIKE " & fMatch & " ORDER BY " & fArray(fField)

		Case 50
			fArray = split("E_Title|" & doConCat(doConCat("U_FirstName","' '"),"U_LastName") & "|E_StartTime","|")

			getSearch = " EventId,E_Module,E_ModuleId,E_Title," & doConCat(doConCat("U_FirstName","' '"),"U_LastName") & ",E_StartTime " & _
				"FROM (CRM_Events E INNER JOIN ALL_Users U ON E.E_Owner = U.UserId) " & _
				"WHERE E_Status = 1 AND " & fArray(fField) & " LIKE " & fMatch & " ORDER BY " & fArray(fField)

		Case 0
			fArray = split(doConCat(doConCat("U_FirstName","' '"),"U_LastName") & "|U_LastName|U_FirstName","|")

			getSearch = " UserId," & doConCat(doConCat("U_FirstName","' '"),"U_LastName") & ",U_Email," & fArray(fThird) & " " & _
				"FROM ALL_Users WHERE " & fArray(fField) & " LIKE " & fMatch & " AND U_Status=1 ORDER BY " & fArray(fField)
	End Select

	getSearch = getSelectTop(getSearch,fMax)
End Function
%>