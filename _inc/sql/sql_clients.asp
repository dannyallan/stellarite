<%
Function getClient(fType,fId)

	Dim fPrevId, fNextId

	If fType = 1 Then
		fPrevId = doPrevNext(0,2,fId,0,0)
		fNextId = doPrevNext(1,2,fId,0,0)
	Else
		fPrevId = CLng(0)
		fNextId = CLng(0)
	End If

	getClient = "SELECT " & fPrevId & " AS PrevId," & fNextId & " AS NextId, D.*, " & _
			"C_Client, " & doConCat(doConCat("US.U_FirstName","' '"),"US.U_LastName") & " AS SalesRep, " & _
			doConCat(doConCat("UC.U_FirstName","' '"),"UC.U_LastName") & " AS CreatedBy, " & _
			doConCat(doConCat("UM.U_FirstName","' '"),"UM.U_LastName") & " AS ModBy " & _
		"FROM ((((CRM_Divisions D " & _
			"INNER JOIN CRM_Clients C ON D.ClientId = C.ClientId) " & _
			"LEFT JOIN ALL_Users US ON D.D_SalesRep = US.UserId) " & _
			"INNER JOIN ALL_Users UC ON D.D_CreatedBy = UC.UserId) " & _
			"INNER JOIN ALL_Users UM ON D.D_ModBy = UM.UserId) " & _
		"WHERE D_Status = 1 AND D.DivId = " & fId
End Function

Sub delClient(fUser,fId)
	objConn.Execute("UPDATE CRM_Divisions SET D_Status = 0, D_ModBy="&fUser&", D_ModDate=" & Application("av_DateNow") & " WHERE DivId = "&fId)
	Call doNotification(2,fId)
End Sub

Sub updateClient(fUser,fId,fRefAccount,fClient,fDivision,fAccount,fAccountType,fSalesRep,fRegion,fWebsite,fVertical,fProbFlag,fSize,fShortDesc)
	Dim fSQL, fClientId

	fClientId = getValue("ClientId","CRM_Divisions","DivId="&fId,0)

	objConn.Execute("UPDATE CRM_Clients SET C_Client = " & sqlText(fClient) & " WHERE ClientId = " & fClientId)

	fSQL = "UPDATE CRM_Divisions SET " & _
			"D_RefAccount = " & fRefAccount & _
			",ClientId = " & fClientId & _
			",D_Division = " & sqlText(fDivision) & _
			",D_Account = " & sqlText(fAccount) & _
			",D_AccountType = " & fAccountType & _
			",D_Website = " & sqlText(fWebsite) & _
			",D_Vertical = " & fVertical & _
			",D_Size = " & fSize & _
			",D_ProbFlag = " & fProbFlag & _
			",D_SalesRep = " & fSalesRep & _
			",D_Region = " & fRegion & _
			",D_ShortDesc = " & sqlText(fShortDesc) & _
			",D_ModBy = " & fUser & _
			",D_ModDate = " & Application("av_DateNow") & _
		" WHERE DivId = " & fId

	objConn.Execute(fSQL)
	Call doNotification(2,fId)
End Sub

Function insertClient(fUser,fId,fRefAccount,fClient,fDivision,fAccount,fAccountType,fSalesRep,fRegion,fWebsite,fVertical,fProbFlag,fSize,fShortDesc)
	Dim fSQL

	fSQL = "INSERT INTO CRM_Divisions (D_RefAccount,ClientId,D_Division,D_AccountType,D_Website" & _
			",D_SalesRep,D_Region,D_Vertical,D_Size,D_ProbFlag,D_Account,D_ShortDesc" & _
			",D_CreatedBy,D_CreatedDate,D_ModBy,D_ModDate,D_Status) VALUES (" & _
			fRefAccount & "," & getClientId(fClient) & "," & sqlText(fDivision) & _
			"," & fAccountType & "," & sqlText(fWebsite) & _
			"," & fSalesRep & "," &  fRegion & _
			"," & fVertical & "," &  fSize & "," & fProbFlag & _
			"," & sqlText(fAccount) & "," &  sqlText(fShortDesc) & _
			"," & fUser & "," & Application("av_DateNow") &  _
			"," & fUser & "," & Application("av_DateNow") &  ",1)"

	objConn.Execute(fSQL)
	insertClient = getLastInsert(fUser,2)
End Function
%>