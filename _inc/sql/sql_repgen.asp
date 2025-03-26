<!--#include file="sql_reports.asp" -->
<%
Function genReport(iType,sFields,sFilters,iOrder,sOrderDir,iMax)

	Dim aFields, aFilters, aDefaults
	Dim sSqlFields, sSqlTables, sSqlFilters
	Dim iField, iCond, sValue, sOrder, sTemp
	Dim sTblLetter, sParens, sTblOptions
	Dim iCount, iThis, iPos1, iPos2

	aDefaults = Application("arr_Fields" & iType)
	aFields = Split(sFields,",")
	aFilters = Split(sFilters,"|")

	sSqlTables = getSqlInfo(iType,1)
	sTblLetter = Right(sSQLTables,1)

	'Build the SELECT fields
	'=======================
	For iCount = 0 to UBound(aFields)

		iThis = getArrayIndex(aFields(iCount),iType)
		sTemp = getSqlField(iThis,iType)

		Select Case aDefaults(3,iThis)
			Case 5
				sTemp = "O" & aFields(iCount) & ".O_Value"
				If iOrder = aFields(iCount) Then sOrder = sTemp
				sTblOptions = sTblOptions & " LEFT JOIN ALL_Options O" & aFields(iCount) & " ON " & aDefaults(2,iThis) & " = O" & aFields(iCount) & ".OptionId) "
				sParens = sParens & "("
		End Select

		If aDefaults(7,iCount) = 0 Then
			sSqlFields = sSqlFields & ", " & sTemp & " AS _" & Replace(aDefaults(1,iThis)," ","_")
		Else
			sSqlFields = sSqlFields & ", " & sTemp & " AS " & aDefaults(1,iThis)
		End If

	Next
	If sSqlFields <> "" Then sSqlFields = Mid(sSqlFields,2) & " " Else sSqlFields = " 'No Fields Selected' "


	'Decide the ORDER field
	'=======================
	If sOrder = "" Then
		iThis = getArrayIndex(iOrder,iType)
		sOrder = getSqlField(iThis,iType)
	End If
	If Instr(sOrder,Application("av_Concat")) > 0 Then sOrder = Left(sOrder,Instr(sOrder,Application("av_Concat"))-1)


	'Build the WHERE filters
	'=======================
	sSqlFilters = sTblLetter & "." & sTblLetter & "_Status = 1 "

	For iCount = 0 to UBound(aFilters)
		iPos1 = Instr(aFilters(iCount),"-")
		iPos2 = Instr(iPos1+1,aFilters(iCount),"-")

		iField = Left(aFilters(iCount),iPos1-1)
		iCond = Mid(aFilters(iCount),iPos1+1,iPos2-iPos1-1)
		sValue = Mid(aFilters(iCount),iPos2+1)

		iThis = getArrayIndex(iField,iType)
		sTemp = getSqlField(iThis,iType)
		If sTemp <> aDefaults(2,iThis) Then iThis = 1 Else iThis = aDefaults(3,iThis)
		If Instr(sTemp,"|") > 0 Then sTemp = Left(sTemp,Instr(sTemp,"|")-3)

		Select Case iCond
			Case 1,10
				sTemp = sTemp & "="
			Case 2,11
				sTemp = sTemp & "<>"
			Case 3,4
				sTemp = sTemp & " LIKE "
			Case 5
				sTemp = sTemp & " NOT LIKE "
			Case 6,9
				sTemp = sTemp & ">="
			Case 7,8
				sTemp = sTemp & "<="
			Case 12
				sTemp = sTemp & " IN "
		End Select

		Select Case iCond
			Case 3
				sValue = sValue & "%"
			Case 4,5
				sValue = "%" & sValue & "%"
		End Select

		Select Case iThis
			Case 1,6,7
				sSqlFilters = sSqlFilters & " AND " & sTemp & sqlText(sValue)
			Case 2,4
				sSqlFilters = sSqlFilters & " AND " & sTemp & sValue
			Case 3
				sSqlFilters = sSqlFilters & " AND " & sTemp & sqlDate(sValue)
			Case 5
				sSqlFilters = sSqlFilters & " AND " & sTemp & "(" & sValue & ")"
		End Select
	Next


	'Build the FROM tables
	'=====================
	sTemp = sSqlFields & sTblOptions & sSqlFilters & sOrder

	If iType <> 2 and (Instr(sTemp,"D.") > 0 or Instr(sTemp,"C_Client") > 0) Then sSqlTables = "(" & sSqlTables & " LEFT JOIN CRM_Divisions D ON " & sTblLetter & ".DivId = D.DivId) "
	If Instr(sTemp,"C_Client") > 0 Then sSqlTables = "(" & sSqlTables & " LEFT JOIN CRM_Clients C ON D.ClientId = C.ClientId) "
	If iType <> 1 and Instr(sTemp,"K.") > 0 Then sSqlTables = "(" & sSqlTables & " LEFT JOIN CRM_Contacts K ON " & sTblLetter & ".ContactId = K.ContactId) "
	If Instr(sTemp,"UC.") > 0 Then sSqlTables = "(" & sSqlTables & " LEFT JOIN ALL_Users UC ON " & sTblLetter & "." & sTblLetter & "_CreatedBy = UC.UserId) "
	If Instr(sTemp,"UM.") > 0 Then sSqlTables = "(" & sSqlTables & " LEFT JOIN ALL_Users UM ON " & sTblLetter & "." & sTblLetter & "_ModBy = UM.UserId) "
	If Instr(sTemp,"UO.") > 0 Then sSqlTables = "(" & sSqlTables & " LEFT JOIN ALL_Users UO ON " & sTblLetter & "." & sTblLetter & "_Owner = UO.UserId) "
	If Instr(sTemp,"US.") > 0 Then sSqlTables = "(" & sSqlTables & " LEFT JOIN ALL_Users US ON " & sTblLetter & "." & sTblLetter & "_SalesRep = US.UserId) "

	sSqlTables = sParens & sSqlTables & sTblOptions


	'Put the SQL query together
	'==========================
	genReport = sSqlFields & " FROM " & sSqlTables & " WHERE " & sSqlFilters & " ORDER BY " & sOrder & " " & sOrderDir

	genReport = getSelectTop(genReport,iMax)

'	Response.Write(genReport)
'	Call endResponse

End Function

Function getArrayIndex(iValue, iType)
	Dim iCount, aDefault

	aDefault = Application("arr_Fields" & iType)

	For iCount = 0 to UBound(aDefault,2)
		If CLng(aDefault(0,iCount)) = CLng(iValue) Then
			getArrayIndex = CLng(iCount)
			Exit Function
		End If
	Next
End Function

Function getSqlField(iIndex,iType)
	Dim sTemp, aDefault

	aDefault = Application("arr_Fields" & iType)

	sTemp = aDefault(1,iIndex)
	Select Case sTemp
		Case "IDS_Owner", "IDS_SalesRep", "IDS_CreatedBy", "IDS_ModifiedBy"
			getSqlField = doConCat(doConCat("U" & Mid(sTemp,5,1) & ".U_FirstName","' '"),"U" & Mid(sTemp,5,1) & ".U_LastName")
		Case "IDS_Contact"
			getSqlField = doConCat(doConCat(doConCat(doConCat("K.K_FirstName","' '"),"K.K_LastName"),"'|'"),"K.ContactId")
		Case "IDS_Account"
			getSqlField = doConCat(doConCat("C.C_Client","'|'"),"D.DivId")
		Case "IDS_Project"
			getSqlField = doConCat(doConCat("P.P_Title","'|'"),"P.ProjectId")
		Case "IDS_Event"
			getSqlField = doConCat(doConCat("E.E_Title","'|'"),"E.EventId")
		Case Else
			getSqlField = aDefault(2,iIndex)
	End Select
End Function

Function updateReport(sUser,nId,sReportName,iMod,sType,iPermissions,sOwner,ByVal sGenSql,aFields,sParams,sOrder,sOrderDir)

	Dim nOwnerId
	nOwnerId = getUserId(0,sOwner)

	updateReport = "UPDATE CRM_Reports SET " & _
			"R_Title = " & sqlText(sReportName) & _
			",R_Module = " & iMod & _
			",R_Type = " & sqlText(sType) & _
			",R_Permissions = " & iPermissions & _
			",R_Owner = " & nOwnerId & _
			",R_SQL = " & sqlText(sGenSql) & _
			",R_Fields = " & sqlText(aFields) & _
			",R_Params = " & sqlText(sParams) & _
			",R_Order = " & sOrder & _
			",R_OrderDir = " & sqlText(sOrderDir) & _
			",R_ModBy = " & sUser & _
			",R_ModDate = " & Application("av_DateNow") & _
			" WHERE  ReportId = " & nId
End Function

Function insertReport(sUser,sReportName,iMod,sType,iPermissions,sOwner,ByVal sGenSql,aFields,sParams,sOrder,sOrderDir)

	Dim nOwnerId
	nOwnerId = getUserId(0,sOwner)

	objConn.Execute("INSERT INTO CRM_Reports (R_Title,R_Module,R_Type,R_Permissions,R_Owner,R_SQL,R_Fields,R_Params" & _
			",R_Order,R_OrderDir,R_CreatedBy,R_CreatedDate,R_ModBy,R_ModDate,R_Status) VALUES (" & _
			sqlText(sReportName) & _
			"," & iMod & _
			"," & sqlText(sType) & _
			"," & iPermissions & _
			"," & nOwnerId & _
			"," & sqlText(sGenSql) & _
			"," & sqlText(aFields) & _
			"," & sqlText(sParams) & _
			"," & sOrder & "," & sqlText(sOrderDir) & _
			"," & sUser & "," & Application("av_DateNow") & _
			"," & sUser & "," & Application("av_DateNow") & ",1)")

	insertReport = getLastInsert(sUser,"R")
End Function

Function delReport(sUser,nId)
	delReport = "UPDATE CRM_Reports SET R_Status = 0, R_ModBy=" & sUser & ", R_ModDate="& Application("av_DateNow") & " WHERE ReportId = " & nId
End Function

%>