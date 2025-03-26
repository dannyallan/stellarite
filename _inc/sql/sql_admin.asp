<%
Sub insertConfigValue(sVariable,sValue)

	If Len(getValue("F_Variable","ALL_Config","F_Variable=" & sqlText(sVariable),"")) > 0 Then
		objConn.Execute("UPDATE ALL_Config SET F_Value = " & sqlText(sValue) & " WHERE F_Variable = " & sqlText(sVariable))
	Else
		objConn.Execute("INSERT INTO ALL_Config (F_Value,F_Variable) VALUES (" & sqlText(sValue) & "," & sqlText(sVariable) & ")")
	End If

End Sub

Function getOptionGroups(sMod,iType)
	Select Case iType
		Case 0
			getOptionGroups = "SELECT OptGroupId, G_Name, G_Description FROM ALL_OptGroups WHERE G_Status = 1 AND G_Name <> 'IDS_Product' AND G_Module IN (" & sMod & ") ORDER BY G_Name"
		Case Else
			getOptionGroups = "SELECT OptGroupId, G_Name FROM ALL_OptGroups WHERE G_Status = 1 AND G_Name = 'IDS_Product'"
	End Select
End Function

Function insertOptionGroup(iMod,sName)
	insertOptionGroup = "INSERT INTO ALL_OptGroups (G_Name,G_Module,G_Required,G_Status) VALUES (" & sqlText(sName) & "," & iMod & ",0,1)"
End Function

Function delOptionGroup(nId)
	delOptionGroup = "UPDATE ALL_OptGroups SET G_Status = 0 WHERE G_Required = 0 AND OptGroupId = " & nId
End Function

Function getOptionGroupName(nId)
	getOptionGroupName = getValue("X_Name","ALL_CustomData","CustomId="&nId,"")
End Function

Function insertOptionValue(sGroup,sValue)
	insertOptionValue = "INSERT INTO ALL_Options (OptGroupId,O_Status,O_Value) VALUES (" & sGroup & ",1," & sqlText(sValue) & ")"
End Function

Function delOptionValue(iValue)
	delOptionValue = "UPDATE ALL_Options SET O_Status = 0 Where OptionId = " & iValue
End Function

Function getModuleList(iMod)
	If iMod = 1 Then
		getModuleList = "SELECT ContactId FROM CRM_Contacts WHERE K_Status = 1"
	Elseif iMod = 2 Then
		getModuleList = "SELECT DivId FROM CRM_Divisions WHERE D_Status = 1"
	Elseif iMod = 3 Then
		getModuleList = "SELECT SaleId FROM CRM_Sales WHERE S_Status = 1"
	Elseif iMod = 4 Then
		getModuleList = "SELECT ProjectId FROM CRM_Projects WHERE P_Status = 1"
	Elseif iMod = 5 Then
		getModuleList = "SELECT TicketId FROM CRM_Tickets WHERE T_Status = 1"
	Elseif iMod = 6 Then
		getModuleList = "SELECT BugId FROM CRM_Bugs WHERE B_Status = 1"
	Elseif iMod = 7 Then
		getModuleList = "SELECT InvoiceId FROM CRM_Invoices WHERE I_Status = 1"
	Elseif iMod = 8 Then
		getModuleList = "SELECT CatId FROM KB_Categories WHERE I_Status = 1"
	End If
End Function

Function updateModuleCount(iMod,nId)
	If iMod = 1 Then
		updateModuleCount = "UPDATE CRM_Contacts SET " & _
				"K_Notes = " & getValue("COUNT(*)","CRM_Notes","N_Status = 1 AND N_Module = 1 AND N_ModuleId = " & nId,0) & " " & _
				",K_Attach = " & getValue("COUNT(*)","CRM_Attach", "A_Status = 1 AND A_Module = 1 AND A_ModuleId = " & nId,0) & " " & _
				",K_Events = " & getValue("COUNT(*)","CRM_Attach", "A_Status = 1 AND A_Module = 1 AND A_ModuleId = " & nId,0) & " " & _
				",K_Sales = " & getValue("COUNT(*)","CRM_Sales", "S_Status = 1 AND ContactId = " & nId,0) & " " & _
				",K_Serials = " & getValue("COUNT(*)","CRM_Serialz", "Z_Status = 1 AND DivId = " & nId,0) & " " & _
				",K_Tickets = " & getValue("COUNT(*)","CRM_Tickets", "T_Status = 1 AND ContactId = " & nId,0) & " " & _
				",K_Invoices = " & getValue("COUNT(*)","CRM_Invoices", "I_Status = 1 AND DivId = " & nId,0) & " " & _
				"WHERE ContactId = "  & nId
	Elseif iMod = 2 Then
		updateModuleCount = "UPDATE CRM_Divisions SET " & _
				"D_Notes = " & getValue("COUNT(*)","CRM_Notes", "N_Status = 1 AND N_Module = 2 AND N_ModuleId = " & nId,0) & " " & _
				",D_Attach = " & getValue("COUNT(*)","CRM_Attach", "A_Status = 1 AND A_Module = 2 AND A_ModuleId = " & nId,0) & " " & _
				",D_Events = " & getValue("COUNT(*)","CRM_Events", "E_Status = 1 AND E_Module = 2 AND E_ModuleId = " & nId,0) & " " & _
				",D_Contacts = " & getValue("COUNT(*)","CRM_Contacts", "K_Status = 1 AND DivId = " & nId,0) & " " & _
				",D_Sales = " & getValue("COUNT(*)","CRM_Sales", "S_Status = 1 AND DivId = " & nId,0) & " " & _
				",D_Serials = " & getValue("COUNT(*)","CRM_Serialz", "Z_Status = 1 AND DivId = " & nId,0) & " " & _
				",D_Projects = " & getValue("COUNT(*)","CRM_Projects", "P_Status = 1 AND DivId = " & nId,0) & " " & _
				",D_Tickets = " & getValue("COUNT(*)","CRM_Tickets", "T_Status = 1 AND DivId = " & nId,0) & " " & _
				",D_Invoices = " & getValue("COUNT(*)","CRM_Invoices", "I_Status = 1 AND DivId = " & nId,0) & " " & _
				"WHERE DivId = "  & nId
	Elseif iMod = 3 Then
		updateModuleCount = "UPDATE CRM_Sales SET " & _
				"S_Notes = " & getValue("COUNT(*)","CRM_Notes", "N_Status = 1 AND N_Module = 3 AND N_ModuleId = " & nId,0) & " " & _
				",S_Attach = " & getValue("COUNT(*)","CRM_Attach", "A_Status = 1 AND A_Module = 3 AND A_ModuleId = " & nId,0) & " " & _
				",S_Events = " & getValue("COUNT(*)","CRM_Events", "E_Status = 1 AND E_Module = 3 AND E_ModuleId = " & nId,0) & " " & _
				"WHERE SaleId = "  & nId
	Elseif iMod = 4 Then
		updateModuleCount = "UPDATE CRM_Projects SET " & _
				"P_Notes = " & getValue("COUNT(*)","CRM_Notes", "N_Status = 1 AND N_Module = 4 AND N_ModuleId = " & nId,0) & " " & _
				",P_Attach = " & getValue("COUNT(*)","CRM_Attach", "A_Status = 1 AND A_Module = 4 AND A_ModuleId = " & nId,0) & " " & _
				",P_Events = " & getValue("COUNT(*)","CRM_Events", "E_Status = 1 AND E_Module = 4 AND E_ModuleId = " & nId,0) & " " & _
				"WHERE ProjectId = "  & nId
	Elseif iMod = 5 Then
		updateModuleCount = "UPDATE CRM_Tickets SET " & _
				"T_Notes = " & getValue("COUNT(*)","CRM_Notes", "N_Status = 1 AND N_Module = 5 AND N_ModuleId = " & nId,0) & " " & _
				",T_Attach = " & getValue("COUNT(*)","CRM_Attach", "A_Status = 1 AND A_Module = 5 AND A_ModuleId = " & nId,0) & " " & _
				",T_Events = " & getValue("COUNT(*)","CRM_Events", "E_Status = 1 AND E_Module = 5 AND E_ModuleId = " & nId,0) & " " & _
			"WHERE TicketId = "  & nId
	Elseif iMod = 6 Then
		updateModuleCount = "UPDATE CRM_Bugs SET " & _
				"B_Notes = " & getValue("COUNT(*)","CRM_Notes", "N_Status = 1 AND N_Module = 6 AND N_ModuleId = " & nId,0) & " " & _
				",B_Attach = " & getValue("COUNT(*)","CRM_Attach", "A_Status = 1 AND A_Module = 6 AND A_ModuleId = " & nId,0) & " " & _
				",B_Events = " & getValue("COUNT(*)","CRM_Events", "E_Status = 1 AND E_Module = 6 AND E_ModuleId = " & nId,0) & " " & _
				",B_Tickets = " & getValue("COUNT(*)","CRM_Tickets", "T_Status = 1 AND T_BugId = " & nId,0) & " " & _
				"WHERE BugId = "  & nId
	Elseif iMod = 7 Then
		updateModuleCount = "UPDATE CRM_Invoices SET " & _
				"I_Notes = " & getValue("COUNT(*)","CRM_Notes", "N_Status = 1 AND N_Module = 7 AND N_ModuleId = " & nId,0) & " " & _
				",I_Attach = " & getValue("COUNT(*)","CRM_Attach", "A_Status = 1 AND A_Module = 7 AND A_ModuleId = " & nId,0) & " " & _
				",I_Events = " & getValue("COUNT(*)","CRM_Events", "E_Status = 1 AND E_Module = 7 AND E_ModuleId = " & nId,0) & " " & _
				",I_Sales = " & getValue("COUNT(*)","CRM_Sales", "S_Status = 1 AND InvoiceId = " & nId,0) & " " & _
				",I_Serials = " & getValue("COUNT(*)","CRM_Serialz", "Z_Status = 1 AND InvoiceId = " & nId,0) & " " & _
				",I_Projects = " & getValue("COUNT(*)","CRM_Projects", "P_Status = 1 AND InvoiceId = " & nId,0) & " " & _
				"WHERE InvoiceId = "  & nId
	Elseif iMod = 8 Then
		updateModuleCount = "UPDATE KB_Categories SET " & _
				"I_Count = " & getValue("COUNT(*)","KB_Articles","H_Status = 1 AND CatId IN (" & getSubCats(nId) & ")",0) & _
				", I_Updated = " & sqlText(getValue("MAX(H_ModDate)","KB_Articles","H_Status = 1 AND CatId IN (" & getSubCats(nId) & ")",0)) & _
				" WHERE CatId = " & nId
	End If
End Function

Function getCustomFields(sPermissions)
	Dim sSQL, iCount

	sSQL = "SELECT CustomId, X_Status, X_Module, X_Name, X_Type, X_Length, X_Required, X_Order " & _
		"FROM ALL_CustomData " & _
		"WHERE X_Status = 1 AND X_Default = 0 AND ("

	For iCount = 1 to Len(sPermissions)
		If CByte(Mid(sPermissions,iCount,1)) >= 5 Then sSQL = sSQL & "(X_Module = " & iCount & ") OR "
	Next

	getCustomFields = Left(sSQL,Len(sSQL)-4) & ") ORDER BY X_Module, X_Name"
End Function

Function insertCustomField(iMod,iDataType,sName,iLength,bMand,iOrder)
	Dim iUnique, sTable, sField, sDataType, sLetter

	iUnique = getValue("CustomId","ALL_CustomData","1=1 ORDER BY CustomId DESC",0)+1
	Select Case iDataType
		Case 1,6,7
			sDataType = "varchar(" & iLength & ")"
		Case 2
			Select Case iLength
				Case 17
					sDataType = "tinyint"
				Case 2
					sDataType = "smallint"
				Case 3
					sDataType = "int"
				Case 5
					sDataType = "double"
				Case 6
					sDataType = "currency"
			End Select
		Case 3
			sDataType = "datetime"
		Case 4
			sDataType = "tinyint"
		Case 5
			objConn.Execute(insertOptionGroup(bytMod,strName))
			sDataType = "smallint"
	End Select

	sTable = getSqlInfo(iMod,1)
	sLetter = Right(sTable,1)
	sTable = Left(sTable,Len(sTable)-2)
	sField = sLetter & "_" & Left(Replace(sName," ",""),15) & iUnique

	objConn.Execute("ALTER TABLE " & sTable & " ADD " & sField & " " & sDataType & " NULL ")

	insertCustomField= "INSERT INTO ALL_CustomData " & _
				"(X_Module,X_Name,X_Field,X_Type,X_Length,X_Required,X_Order,X_Default,X_Status) VALUES " & _
				"(" & iMod & "," & sqlText(sName) & "," & sqlText(sLetter & "." & sField) & "," & iDataType & _
				"," & iLength & "," & bMand & "," & iOrder & ",0,1)"
End Function

Function delCustomField(nId)
	Dim sSQL

	sSQL = "SELECT G.OptGroupId FROM ALL_CustomData X " & _
			"INNER JOIN ALL_OptGroups G ON X.X_Name = G.G_Name " & _
			"WHERE X.X_Type = 5 AND X.X_Required = 0 AND X.CustomId = " & nId

	Set objFRS = objConn.Execute(sSQL)
	If not (objFRS.BOF and objFRS.EOF) Then objConn.Execute(delOptionGroup(objFRS.fields(0).value))

	delCustomField = "UPDATE ALL_CustomData SET X_Status = 0 WHERE CustomId = " & nId
End Function
%>