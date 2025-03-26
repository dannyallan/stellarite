<%
Function getReport(fId)
	getReport = "SELECT R.* , " & doConCat(doConCat("UC.U_FirstName","' '"),"UC.U_LastName") & " AS CreatedBy, " & _
				doConCat(doConCat("UM.U_FirstName","' '"),"UM.U_LastName") & " AS ModBy, " & _
				doConCat(doConCat("UO.U_FirstName","' '"),"UO.U_LastName") & " AS Owner " & _
			"FROM (((CRM_Reports R INNER JOIN ALL_Users UC ON R.R_CreatedBy = UC.UserId) " & _
				"INNER JOIN ALL_Users UM ON R.R_ModBy = UM.UserId) " & _
				"INNER JOIN ALL_Users UO ON R.R_Owner = UO.UserId) " & _
			"WHERE R_Status = 1 " & _
			"AND ReportId = " & fId
End Function
%>