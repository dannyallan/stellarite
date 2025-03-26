<%
Function getBug(fType,fId)

	Dim fPrevId, fNextId

	If fType = 1 Then
		fPrevId = doPrevNext(0,6,fId,0,0)
		fNextId = doPrevNext(1,6,fId,0,0)
	Else
		fPrevId = CLng(0)
		fNextId = CLng(0)
	End If

	getBug = "SELECT " & fPrevId & " AS PrevId," & fNextId & " AS NextId, B.*, " & _
			doConCat(doConCat("UO.U_FirstName","' '"),"UO.U_LastName") & " AS Owner, " & _
			doConCat(doConCat("UC.U_FirstName","' '"),"UC.U_LastName") & " AS CreatedBy, " & _
			doConCat(doConCat("UM.U_FirstName","' '"),"UM.U_LastName") & " AS ModBy " & _
		"FROM (((CRM_Bugs B " & _
			"INNER JOIN ALL_Users UO ON B.B_Owner = UO.UserId) " & _
			"INNER JOIN ALL_Users UC ON B.B_CreatedBy = UC.UserId) " & _
			"INNER JOIN ALL_Users UM ON B.B_ModBy = UM.UserId) " & _
		"WHERE B.B_Status = 1 AND B.BugId = " & fId

End Function

Sub delBug(fUser,fId)
	objConn.Execute("UPDATE CRM_Bugs SET B_Status = 0, B_ModBy="&fUser&", B_ModDate=" & Application("av_DateNow") & " WHERE BugId = "&fId)
	Call doNotification(6,fId)
End Sub

Sub updateBug(fUser,fId,fOwner,fHotIssue,fPriority,fBugType,fBugSource,fProduct,fBuild,fDescription,fSolution,fCause,fClosed,fCloseDate)
	Dim fSQL

	fSQL = "UPDATE CRM_Bugs SET " & _
			"B_Owner = " & fOwner & _
			",B_HotIssue = " & fHotIssue & _
			",B_Priority = " & fPriority & _
			",B_BugType = " & fBugType & _
			",B_BugSource = " & fBugSource & _
			",B_ProductId = " & fProduct & _
			",B_Build = " & sqlText(fBuild) & _
			",B_Description = " & sqlText(fDescription) & _
			",B_Solution = " & sqlText(fSolution) & _
			",B_Cause = " & fCause & _
			",B_Closed = " & fClosed & _
			",B_CloseDate = " & sqlDate(fCloseDate) & _
			",B_ModBy = " & fUser & _
			",B_ModDate = " & Application("av_DateNow") &  _
		" WHERE  BugId = " & fId

	objConn.Execute(fSQL)
	Call doNotification(6,fId)
End Sub


Function insertBug(fUser,fId,fOwner,fHotIssue,fPriority,fBugType,fBugSource,fProduct,fBuild,fDescription,fSolution,fCause,fClosed,fCloseDate)
	Dim fSQL

	fSQL = "INSERT INTO CRM_Bugs (B_Owner,B_HotIssue,B_Priority,B_BugType,B_BugSource" & _
		",B_ProductId,B_Build,B_Description,B_Solution,B_Cause,B_Closed,B_CloseDate" & _
		",B_CreatedBy,B_CreatedDate,B_ModBy,B_ModDate,B_Status) VALUES (" & _
		fOwner & "," & fHotIssue & _
		"," & fPriority & "," & fBugType & _
		"," & fBugSource & "," & fProduct & _
		"," & sqlText(fBuild) & "," & sqlText(fDescription) & _
		"," & sqlText(fSolution) & "," & fCause & _
		"," & fClosed & "," & sqlDate(fCloseDate) & _
		"," & fUser & "," & Application("av_DateNow") &  _
		"," & fUser & "," & Application("av_DateNow") &  ",1)"

	objConn.Execute(fSQL)
	insertBug = getLastInsert(fUser,6)
End Function
%>