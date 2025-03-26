<%
Function getContact(fType,fId)

	Dim fPrevId, fNextId

	If fType = 1 Then
		fPrevId = doPrevNext(0,1,fId,0,0)
		fNextId = doPrevNext(1,1,fId,0,0)
	Else
		fPrevId = CLng(0)
		fNextId = CLng(0)
	End If

	getContact = "SELECT " & fPrevId & " AS PrevId," & fNextId & " AS NextId, K.*, " & _
			"C.C_Client, D.D_Division, " & _
			doConCat(doConCat("UC.U_FirstName","' '"),"UC.U_LastName") & " AS CreatedBy, " & _
			doConCat(doConCat("UM.U_FirstName","' '"),"UM.U_LastName") & " AS ModBy " & _
		"FROM ((((CRM_Contacts K " & _
			"INNER JOIN CRM_Divisions D ON K.DivId = D.DivId) " & _
			"INNER JOIN CRM_Clients C ON D.ClientId = C.ClientId) " & _
			"INNER JOIN ALL_Users UC ON K.K_CreatedBy = UC.UserId) " & _
			"INNER JOIN ALL_Users UM ON K.K_ModBy = UM.UserId) " & _
		"WHERE K_Status = 1 AND ContactId = " & fId

End Function

Sub delContact(fUser,fId)
	objConn.Execute("UPDATE CRM_Contacts SET K_Status = 0, K_ModBy="&fUser&", K_ModDate=" & Application("av_DateNow") & " WHERE ContactId = "&fId)
	objConn.Execute("UPDATE CRM_Divisions SET D_Contacts = D_Contacts-1 WHERE DivId = "&getValue("DivId","CRM_Contacts","ContactId="&fId,0))
	Call doNotification(1,fId)
End Sub

Sub updateContact(fUser,fId,fMod,fModId,fPrefix,fFirstName,FMiddleInitial,fLastName,fAddress1,fAddress2,fAddress3,fCity,fState,fCountry,fZIP,fEmail,fDept,fJobTitle,fPhone1,fExt1,fPhone2,fExt2,fFax,fClient,fDivision,fReportsTo,fAssistant,fNoEmail,fDoNotCall)
	Dim fSQL

	If fMod <> 2 Then fModId = getDivId(fUser,fDivision,getClientId(fClient))

	fSQL = "UPDATE CRM_Contacts SET K_Prefix = " & fPrefix & _
			",K_FirstName = " & sqlText(fFirstName) & _
			",K_MiddleInitial = " & sqlText(fMiddleInitial) & _
			",K_LastName = " & sqlText(fLastName) & _
			",K_Address1 = " & sqlText(fAddress1) & _
			",K_Address2 = " & sqlText(fAddress2) & _
			",K_Address3 = " & sqlText(fAddress3) & _
			",K_City = " & sqlText(fCity) & _
			",K_State = " & sqlText(fState) & _
			",K_Country = " & sqlText(fCountry) & _
			",K_ZIP = " & sqlText(fZIP) & _
			",K_Email = " & sqlText(fEmail) & _
			",K_EmailOptOut = " & fNoEmail & _
			",DivId = " & fModId & _
			",K_Dept = " & sqlText(fDept) & _
			",K_JobTitle = " & sqlText(fJobTitle) & _
			",K_Phone1 = " & fPhone1 & _
			",K_Ext1 = " & fExt1 & _
			",K_Phone2 = " & fPhone2 & _
			",K_Ext2 = " & fExt2 & _
			",K_Fax = " & fFax & _
			",K_DoNotCall = " & fDoNotCall & _
			",K_ReportsTo = " & fReportsTo & _
			",K_Assistant = " & fAssistant & _
			",K_ModBy = " & fUser & _
			",K_ModDate = " & Application("av_DateNow") & _
		" WHERE  ContactId = " & fId

	objConn.Execute(fSQL)
	Call doNotification(1,fId)
End Sub

Function insertContact(fUser,fId,fMod,fModId,fPrefix,fFirstName,FMiddleInitial,fLastName,fAddress1,fAddress2,fAddress3,fCity,fState,fCountry,fZIP,fEmail,fDept,fJobTitle,fPhone1,fExt1,fPhone2,fExt2,fFax,fClient,fDivision,fReportsTo,fAssistant,fNoEmail,fDoNotCall)
	Dim fSQL

	If fMod <> 2 Then fModId = getDivId(fUser,fDivision,getClientId(fClient))

	fSQL = "INSERT INTO CRM_Contacts (K_Prefix,K_FirstName,K_MiddleInitial,K_LastName" & _
		",K_Address1,K_Address2,K_Address3,K_City,K_State,K_Country,K_Zip,DivId,K_Dept" & _
		",K_JobTitle,K_Email,K_EmailOptOut,K_Phone1,K_Phone2,K_Ext1,K_Ext2,K_Fax,K_DoNotCall," & _
		"K_ReportsTo,K_Assistant,K_CreatedBy,K_CreatedDate,K_ModBy,K_ModDate,K_Status) VALUES (" & _
		fPrefix & "," & sqlText(fFirstName) & _
		"," & sqlText(fMiddleInitial) & "," & sqlText(fLastName) & _
		"," & sqlText(fAddress1) & "," & sqlText(fAddress2) & _
		"," & sqlText(fAddress3) & "," & sqlText(fCity) & _
		"," & sqlText(fState) & "," & sqlText(fCountry) & _
		"," & sqlText(fZIP) & "," & fModId & _
		"," & sqlText(fDept) & "," & sqlText(fJobTitle) & _
		"," & sqlText(fEmail) & "," & fNoEmail & "," & fPhone1 & _
		"," & fPhone2 & "," & fExt1 & _
		"," & fExt2 & "," & fFax & "," & fDoNotCall & _
		"," & fReportsTo & "," & fAssistant & _
		"," & fUser & "," & Application("av_DateNow") &  _
		"," & fUser & "," & Application("av_DateNow") &  ",1)"

	objConn.Execute(fSQL)
	insertContact = getLastInsert(fUser,1)
	objConn.Execute("UPDATE CRM_Divisions SET D_Contacts = D_Contacts+1 WHERE DivId = " & fModId)
End Function


Function getContactsByDiv(fDiv,fOrder,fSort)

	getContactsByDiv = "SELECT ContactId," & doConCat(doConCat("K_FirstName","' '"),"K_LastName") & ",K_Email,K_Phone1,K_Ext1 FROM CRM_Contacts WHERE K_Status = 1 AND DivId = " & fDiv

	If fOrder = "2" Then
		getContactsByDiv = getContactsByDiv & " ORDER BY K_Email " & fSort
	ElseIf fOrder = "3" Then
		getContactsByDiv = getContactsByDiv & " ORDER BY K_Phone1 " & fSort
	Else
		getContactsByDiv = getContactsByDiv & " ORDER BY K_FirstName " & fSort
	End If

End Function
%>