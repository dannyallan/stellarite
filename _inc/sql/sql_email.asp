<%
Function getEmailSub(fUser,fMod,fModId)

	getEmailSub = "SELECT * FROM CRM_Email " & _
			"WHERE UserId = " & fUser & _
			" AND M_Module = " & fMod & _
			" AND M_ModuleId = " & fModId
End Function

Function insertEmailSub(fUser,fMod,fModId)

	insertEmailSub = "INSERT INTO CRM_Email (UserId,M_Module,M_ModuleId) VALUES (" & fUser & "," & fMod & "," & fModId & ")"

End Function

Function delEmailSub(fUser,fId,fMod,fModId)

	delEmailSub = "DELETE FROM CRM_Email WHERE UserId = " & fUser

	If fId = 0 Then
		delEmailSub = delEmailSub & " AND M_Module = " & fMod & " AND M_ModuleId = 0 "
	Else
		delEmailSub = delEmailSub & " AND MailId = " & fId
	End If

End Function

Function getSubscriptions(fId)
	getSubscriptions = "SELECT M.MailId, M.M_Module, " & _
				doConCat(doConCat(doConCat("'" & getIDS("IDS_Contact") & ": '","K.K_FirstName"),"' '"),"K.K_LastName") & " AS Mod1, " & _
				doConCat("'" & getIDS("IDS_Account") & ": '","C.C_Client") & " AS Mod2, " & _
				doConCat("'" & getIDS("IDS_Sale") & ": '","SC.C_Client") & " AS Mod3, " & _
				doConCat("'" & getIDS("IDS_Project") & ": '","P.P_Title") & " AS Mod4, " & _
				doConCat("'" & getIDS("IDS_TicketId") & ": '","T.TicketId") & " AS Mod5, " & _
				doConCat("'" & getIDS("IDS_BugId") & ": '","B.BugId") & "  AS Mod6, " & _
				doConCat("'" & getIDS("IDS_InvoiceId") & ": '","I.InvoiceId") & "  AS Mod7 " & _
			"FROM " & _
				"((((((((((CRM_Email M " & _
					"LEFT JOIN CRM_Contacts K ON M.M_ModuleId = K.ContactId) " & _
					"LEFT JOIN CRM_Divisions D ON M.M_ModuleId = D.DivId) " & _
					"LEFT JOIN CRM_Clients C on D.ClientId = C.ClientId) " & _
					"LEFT JOIN CRM_Sales S ON M.M_ModuleId = S.SaleId) " & _
					"LEFT JOIN CRM_Divisions SD on S.DivId = SD.DivId) " & _
					"LEFT JOIN CRM_Clients SC on SD.ClientId = SC.ClientId) " & _
					"LEFT JOIN CRM_Projects P ON M.M_ModuleId = P.ProjectId) " & _
					"LEFT JOIN CRM_Tickets T ON M.M_ModuleId = T.TicketId) " & _
					"LEFT JOIN CRM_Bugs B ON M.M_ModuleId = B.BugId) " & _
					"LEFT JOIN CRM_Invoices I ON M.M_ModuleId = I.InvoiceId) " & _
			"WHERE M.UserId = " & fId & " " & _
				"AND M.M_ModuleId > 0 " & _
			"ORDER BY 2 ASC, 1 ASC"
End Function
%>