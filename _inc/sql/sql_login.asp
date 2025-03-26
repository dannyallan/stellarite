<%
Function doLogin(fUserName,fPass,fIP,fLogins)
	Dim fArray

	Set objFRS = objConn.Execute("SELECT UserId, U_Password, " & doConCat(doConCat("U_FirstName","' '"),"U_LastName") & ", U_Admin, U_ChangePassword, U_Member, U_Permissions, U_LoginAttempts, U_Portal FROM ALL_Users WHERE U_Status = 1 AND U_UserName = " & sqlText(fUserName))

	If objFRS.BOF and objFRS.EOF Then
		doLogin = "SELECT 1"
	Else
		fArray = objFRS.GetRows()
		If fPass = fArray(1,0) and (fLogins > fArray(7,0) or fIP = "127.0.0.1") Then
			objConn.Execute("UPDATE ALL_Users SET U_LastAccess = " & Application("av_DateNow") & ", U_LastIP = " & sqlText(fIP) & ", U_LoginAttempts = 0 WHERE UserId =" & fArray(0,0))
			doLogin = "SELECT 0," & fArray(0,0) & "," & sqlText(fArray(2,0)) & "," & fArray(3,0) & "," & sqlText(fArray(4,0)) & "," & sqlText(fArray(5,0)) & "," & sqlText(fArray(6,0))& "," & sqlText(fArray(8,0))
		Else
			objConn.Execute("UPDATE ALL_Users SET U_LastAccess = " & Application("av_DateNow") & ", U_LastIP = " & sqlText(fIP) & " WHERE UserId =" & fArray(0,0))
			If CInt(fArray(7,0)) > fLogins Then
				If fArray(3,0) = 1 Then
					doLogin = "SELECT 2"
				Else
					doLogin = "SELECT 3"
				End If
			Elseif fPass <> fArray(1,0) Then
				doLogin = "SELECT 4"
			Else
				doLogin = "SELECT 5"
			End If
		End If
	End If
End Function
%>