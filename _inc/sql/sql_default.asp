<%
Function doConCat(fVal1,fVal2)
	Dim sConcat
	sConcat = Application("av_Concat")
	If strDatabase = "Access" Then
		doConCat = fVal1 & sConcat & fVal2
	Elseif strDatabase = "MSSQL" Then
		If Right(UCase(fVal1),2) = "ID" Then fVal1 = "CAST(" & fVal1 & " AS nvarchar)"
		If Right(UCase(fVal2),2) = "ID" Then fVal2 = "CAST(" & fVal2 & " AS nvarchar)"
		doConCat = fVal1 & sConcat & fVal2
	Elseif strDatabase = "MySQL" Then
		doConCat = "CONCAT(" & fVal1 & "," & fVal2 & ")"
	Elseif strDatabase = "Oracle" Then
		doConCat = fVal1 & sConcat & fVal2
	End If
End Function

Function sqlText(ByVal fString)

	If Len(fString) > 0 Then
		fString = Replace(fString,"'","''")

		If strDatabase = "MySQL" Then
			fString = Replace(fString, "\", "\\")
		'    fString = Replace(fString, "\0", "\\0")
		'    fString = Replace(fString, "\'", "\\'")
		'    fString = Replace(fString, "\""", "\\""")
		'    fString = Replace(fString, "\b", "\\b")
		'    fString = Replace(fString, "\n", "\\n")
		'    fString = Replace(fString, "\r", "\\r")
		'    fString = Replace(fString, "\t", "\\t")
		'    fString = Replace(fString, "\z", "\\z")
		'    fString = Replace(fString, "\%", "\\%")
		'    fString = Replace(fString, "\_", "\\_")
		End If

		sqlText = "'" & fString & "'"
	Else
		sqlText = "NULL"
	End If
End Function

Function sqlLike(fString)
	sqlLike = sqlText(Replace(fString,"*","%"))
End Function

Function sqlName(fName)

	Select Case UCase(fName)
		Case "CURRENCY"
			fName = "Kurrency"
		Case "VALUE"
			fName = "Valu"
	End Select

	sqlName = Replace(fName," ","_")

End Function

Function sqlDate(fDate)
	If Len(fDate) <= 10 and IsDate(fDate) Then
		sqlDate = Application("av_DateDel") & Year(fDate) & "-" & Month(fDate) & "-" & Day(fDate) & Application("av_DateDel")
	Elseif IsDate(fDate) Then
		sqlDate = Application("av_DateDel") & Year(fDate) & "-" & Month(fDate) & "-" & Day(fDate) & " " & Hour(fDate) & ":" & Minute(fDate) & Application("av_DateDel")
	Else
		sqlDate = "NULL"
	End If
End Function

Function getValue(fField,fTables,fParams,fReturn)
	Dim fSQL
	Dim fRS

	fSQL = fField & " FROM " & fTables & " WHERE  " & fParams
	fSQL = getSelectTop(fSQL,1)

	Set fRS = objConn.Execute(fSQL)

	If (fRS.BOF and fRS.EOF) Then
		getValue = fReturn
	Elseif Len(fRS.fields(0).value & "") = 0 Then
		getValue = fReturn
	Else
		getValue = fRS.Fields(0).value
	End If

	fRS.Close
	Set fRS = Nothing
End Function

Function getSelectTop(fSQL,fCount)
	If fCount <> "" and IsNumeric(fCount) Then
		Select Case strDatabase
			Case "MySQL"
				getSelectTop = "SELECT " & fSQL & " LIMIT " & fCount
			Case Else
				getSelectTop = "SELECT TOP " & fCount & " " & fSQL
		End Select
	Else
		getSelectTop = "SELECT " & fSQL
	End If
End Function

Function getOptionValues(iGroup)
	getOptionValues = "SELECT OptionId, O_Value FROM ALL_Options WHERE O_Status = 1 "
	If iGroup > 0 Then
		getOptionValues = getOptionValues & " AND OptGroupId = " & iGroup & " ORDER BY O_Value"
	Else
		getOptionValues = getOptionValues & " ORDER BY OptionId"
	End If
End Function

Function getConfigValues()
	getConfigValues = "SELECT * FROM ALL_Config ORDER BY F_Variable"
End Function

Function getModuleFields(iMod)
	getModuleFields = "SELECT CustomId,X_Name,X_Field,X_Type,X_Length,X_Required,X_Order,X_Default FROM ALL_CustomData WHERE X_Status = 1 AND X_Name <> '' AND X_Module = " & CStr(iMod) & " ORDER BY X_Order"
End Function

Function getSubscriptionSQL(fMod,fModId,fHotIssue)
	Dim fSQL

	fSQL = "SELECT DISTINCT U.U_Email FROM CRM_Email M, ALL_Users U " & _
		"WHERE M.UserId = U.UserId " & _
		"AND M.M_Module = " & fMod & " AND M.M_ModuleId = " & fModId & " "

	If fHotIssue Then fSQL = fSQL & " OR M.M_ModuleId = 0"

	getSubscriptionSQL = fSQL
End Function

Function getSubCats(fCatId)

	Dim fAllCats, fThisCat, fMore

	fMore = True
	fAllCats = CStr(fCatId) & ","
	fThisCat = CStr(fCatId) & ","

	Do while fMore

		Set objFRS = objConn.Execute("SELECT CatId FROM KB_Categories WHERE I_ParentId IN (" & Left(fThisCat,Len(fThisCat)-1) & ")")
		fThisCat = ""

		If objFRS.BOF and objFRS.EOF Then
			fMore = False
		Else
			Do While Not objFRS.EOF
				fAllCats = fAllCats & objFRS.fields("CatId") & ","
				fThisCat = fThisCat & objFRS.fields("CatId") & ","
				objFRS.MoveNext
			Loop
		End If
	Loop

	getSubCats = Left(fAllCats,Len(fAllCats)-1)

End Function

%>