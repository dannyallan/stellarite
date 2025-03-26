<%
Function getUsers(fOrder,fSort)
	Dim fSQL

	fSQL = "SELECT UserId, " & doConCat(doConCat("U_FirstName","' '"),"U_LastName") & ", U_Address1, U_Address2, U_Address3, " & _
			"U_City,U_State,U_Country,U_Zip,U_Email,U_Phone1,U_Ext1,U_Phone2,U_Ext2, " & _
			"U_UserName,U_Password,U_TimeZone,U_LastIP,U_LastAccess " & _
		"FROM ALL_Users WHERE U_Status = 1 AND UserId <> 1 "

	If fOrder = "2" Then
		getUsers = fSQL & " ORDER BY U_Email " & fSort
	ElseIf fOrder = "3" Then
		getUsers = fSQL & " ORDER BY U_LastAccess " & fSort
	Else
		getUsers = fSQL & " ORDER BY U_FirstName " & fSort
	End If

End Function

Function getUser(fId)
	getUser = "SELECT * FROM ALL_Users WHERE U_Status = 1 AND UserId = " & fId
End Function

Function delUser(fUser,fId)

	objConn.Execute("UPDATE ALL_Users SET U_Status=0, " & _
			"U_ModDate =" & Application("av_DateNow") & _
			" WHERE UserId = " & fId)
End Function

Sub updateUser(fId,fFName,fLName,fAdd1,fAdd2,fAdd3,fCity,fState,fCountry,fZIP,fTZ,fEmail,fPho1,fExt1,fPho2,fExt2,fUserName,fPass,fAdmin)
	Dim fSQL

	fSQL = "UPDATE ALL_Users SET " & _
			"U_FirstName = " & sqlText(fFName) & _
			",U_LastName = " & sqlText(fLName) & _
			",U_Address1 = " & sqlText(fAdd1) & _
			",U_Address2 = " & sqlText(fAdd2) & _
			",U_Address3 = " & sqlText(fAdd3) & _
			",U_City = " & sqlText(fCity) & _
			",U_State = " & sqlText(fState) & _
			",U_Country = " & sqlText(fCountry) & _
			",U_ZIP = " & sqlText(fZIP) & _
			",U_TimeZone = " & fTZ & _
			",U_Email = " & sqlText(fEmail) & _
			",U_Phone1 = " & fPho1 & _
			",U_Ext1 = " & fExt1 & _
			",U_Phone2 = " & fPho2 & _
			",U_Ext2 = " & fExt2

	' Do not update if the password remains MD5 hash for "password"
	If fAdmin and fPass <> "5f4dcc3b5aa765d61d8327deb882cf99" Then
		fSql = fSql & ",U_UserName = " & sqlText(fUserName) & _
				",U_Password = " & sqlText(fPass)
	End if

	fSql = fSql & ",U_ModDate = " & Application("av_DateNow") & " WHERE  UserId = " & fId

	objConn.Execute(fSQL)
End Sub

Function getDefUserPerm()
	getDefUserPerm = "SELECT U_Member,U_Permissions FROM ALL_Users WHERE UserId = 1"
End Function

Function insertUser(fFName,fLName,fAdd1,fAdd2,fAdd3,fCity,fState,fCountry,fZIP,fTZ,fEmail,fPho1,fExt1,fPho2,fExt2,fUserName,fPass,fMember,fPerm,fPort)
	Dim fSQL

	fSQL = "INSERT INTO ALL_Users (U_FirstName,U_LastName,U_Address1,U_Address2,U_Address3,U_City,U_State,U_Country,U_Zip" & _
		",U_TimeZone,U_Email,U_Member,U_Permissions,U_Portal,U_UserName,U_Password,U_CreatedDate,U_ModDate,U_Status" & _
		",U_Phone1,U_Phone2,U_Ext1,U_Ext2) VALUES (" & _
		sqlText(fFName) & "," & sqlText(fLName) & "," & sqlText(fAdd1) & _
		"," & sqlText(fAdd2) & "," & sqlText(fAdd3) & "," & sqlText(fCity) & _
		"," & sqlText(fState) & "," & sqlText(fCountry) & "," & sqlText(fZIP) & _
		"," & fTZ & "," & sqlText(fEmail) & "," & sqlText(fMember) & _
		"," & sqlText(fPerm) & "," & sqlText(fPort) & "," & sqlText(fUserName) & _
		"," & sqlText(fPass) & "," & Application("av_DateNow") & _
		"," & Application("av_DateNow") & ",1," & fPho1 & "," & fPho2 & _
		"," & fExt1 & "," & fExt2 & ")"

	objConn.Execute(fSQL)
	insertUser = getLastInsert(fUserName,"U")
End Function

Function getUserDetails(fId)
	getUserDetails = "SELECT " & doConCat(doConCat("U_FirstName","' '"),"U_LastName") & ",U_Password,U_Admin,U_ChangePassword,U_Member,U_Permissions,U_LoginAttempts " & _
			"FROM ALL_Users WHERE UserId = " & fId
End Function

Function updateUserPerm(fId,fMember,fPerm,fAdmin,fPass,fLock)
	Dim fSQL

	fSQL = "UPDATE ALL_Users SET " & _
				"U_Member = " & sqlText(fMember) & _
				",U_Permissions = " & sqlText(fPerm)

	If fAdmin <> "NULL" Then
		fSQL = fSQL & ",U_Admin = " & fAdmin & _
				",U_ChangePassword = " & fPass & _
				",U_LoginAttempts = " & fLock
	End If

	updateUserPerm = fSQL &    " WHERE UserId = " & fId

End Function

Function updateUserPass(fId,fPass)
	updateUserPass = "UPDATE ALL_Users SET U_Password = " & sqlText(fPass) & _
						",U_ChangePassword = 0" & _
						",U_ModDate =" & Application("av_DateNow") & _
					" WHERE UserId = " & fId
End Function
%>