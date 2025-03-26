<%
Sub delCategory(fUser,fId)

	objConn.Execute("UPDATE KB_Categories SET I_STATUS = 0 WHERE CatId IN (" & getSubCats(fId) & ")")
	objConn.Execute("UPDATE KB_Articles SET H_STATUS = 0" & _
			", H_ModBy = " & fUser & _
			", H_ModDate = " & Application("av_DateNow")& _
			" WHERE CatId IN (" & getSubCats(fId) & ")")

End Sub

Sub updateCategory(fUser,fId,fParent,fName,fDescription)

	objConn.Execute("UPDATE KB_Categories SET " & _
			"I_Name = " & sqlText(fName) & _
			",I_ParentId = " & fParent & _
			",I_Description = " & sqlText(fDescription) & _
			" WHERE CatId = " & fId)
End Sub

Function insertCategory(fUser,fId,fParent,fName,fDescription)

	objConn.Execute("INSERT INTO KB_Categories (I_Name,I_ParentId,I_Description,I_Count,I_Status) VALUES (" & _
			sqlText(fName)& "," & fParent & "," & sqlText(fDescription) & ",0,1)")

	fId = getLastInsert(fUser,"C")
	insertCategory = fId
End Function

Function getCategory(fType,fCat)

	getCategory = "SELECT DISTINCT kc.CatId, kc.I_Name, kc.I_ParentId, kc.I_Description, kc.I_Updated, kc.I_Count " & _
			"FROM KB_Categories kc " & _
			"WHERE kc.I_Status = 1 "

	If fType = 1 Then getCategory = getCategory & " AND kc.CatId = " & fCat
	If fType = 2 Then getCategory = getCategory & " AND kc.I_ParentId = " & fCat

	getCategory = getCategory & " ORDER BY kc.I_Name "
End Function

Function getCategorySQL()

	getCategorySQL = "SELECT DISTINCT kc.CatId, kc.I_Name " & _
			"FROM KB_Categories kc " & _
			"WHERE kc.I_Status = 1 " & _
			"ORDER BY kc.I_Name"
End Function

Sub delArticle(fUser,fId,fCatId)

	objConn.Execute("UPDATE KB_Articles SET H_Status = 0 " & _
			",H_ModBy = " & fUser & _
			",H_ModDate = " & Application("av_DateNow") & _
			" WHERE ArticleId = " & fId)

	Call updateArticleCount(fCatId,"-")
End Sub

Function getArticle(fInc,fId)

	If fInc = 1 Then
		objConn.Execute("UPDATE KB_Articles SET H_Views = H_Views + 1 WHERE ArticleId = " & fId)
	End If

	getArticle = "SELECT CatId, H_Title, H_Keywords, H_Summary, H_Expire, H_Permissions, H_Views, H_RateCount, H_RateTotal, " & _
				doConCat(doConCat("UC.U_FirstName","' '"),"UC.U_LastName") & " AS CreatedBy, H_CreatedDate, " & _
				doConCat(doConCat("UM.U_FirstName","' '"),"UM.U_LastName") & " AS ModBy, H_ModDate, H_Info, H_Link " & _
			"FROM ((KB_Articles H INNER JOIN ALL_Users UC ON H.H_CreatedBy = UC.UserId) " & _
				"INNER JOIN ALL_Users UM ON H.H_ModBy = UM.UserId) " & _
			"WHERE H_Status = 1 AND ArticleId = " & fId
End Function

Function getArticleSearch(fTerms,fCat,fUsing,fWhere,fMax)

	Dim arrTerms

	getArticleSearch = "ArticleId, H_Title, H_ModDate, H_Summary, H_Expire, H_Views, H_RateCount, H_RateTotal " & _
			"FROM KB_Articles " & _
			"WHERE H_Status = 1 "

	If fCat <> 0 Then getArticleSearch = getArticleSearch & " AND CatId IN (" & fCat & ") "

	If fTerms <> "" Then

		arrTerms = Split(fTerms," ")

		If isArray(arrTerms) Then

			getArticleSearch = getArticleSearch & " AND ("

			Select Case fWhere
				Case 0
					For i = 0 to UBound(arrTerms)
						getArticleSearch = getArticleSearch & "(H_Title LIKE " & sqlLike(arrTerms(i)) & ") OR "
					Next
				Case 1
					For i = 0 to UBound(arrTerms)
						getArticleSearch = getArticleSearch & "(H_Title LIKE " & sqlLike("%" & arrTerms(i) & "%") & " OR H_Summary LIKE " & sqlLike("%" & arrTerms(i) & "%") & " OR H_Info LIKE " & sqlLike("%" & arrTerms(i) & "%") & ") OR "
					Next
				Case 2
					For i = 0 to UBound(arrTerms)
						getArticleSearch = getArticleSearch & "(ArticleId = " & CLng(arrTerms(i)) & ") OR "
					Next
			End Select

			getArticleSearch = Left(getArticleSearch,Len(getArticleSearch)-4) & ")"
		End If
	End If

	getArticleSearch = getSelectTop(getArticleSearch,fMax)

End Function

Sub updateArticle(fUser,fId,fCatId,fTitle,fKeywords,fSummary,fInfo,fLink,fExpiry,fPerm)
	Dim fField

	objConn.Execute("UPDATE KB_Articles SET " & _
			"CatID = " & fCatId & _
			",H_Title = " & sqlText(fTitle) & _
			",H_Keywords = " & sqlText(fKeywords) & _
			",H_Summary = " & sqlText(fSummary) & _
			",H_Info = " & sqlText(fInfo) & _
			",H_Link = " & sqlText(fLink) & _
			",H_Expire = " & sqlDate(fExpiry) & _
			",H_Permissions = " & fPerm & _
			",H_ModBy = " & fUser & _
			",H_ModDate = " & Application("av_DateNow") & _
			" WHERE ArticleId = " & fId)
End Sub

Function insertArticle(fUser,fId,fCatId,fTitle,fKeywords,fSummary,fInfo,fLink,fExpiry,fPerm)
	Dim fField

	objConn.Execute("INSERT INTO KB_Articles (CatId,H_Title,H_Keywords,H_Summary,H_Info,H_Link,H_Expire,H_Permissions," & _
			"H_Views,H_RateTotal,H_RateCount,H_CreatedBy,H_CreatedDate,H_ModBy,H_ModDate,H_Status) VALUES (" & _
			fCatId & "," & _
			sqlText(fTitle) & "," & _
			sqlText(fKeywords) & "," & _
			sqlText(fSummary) & "," & _
			sqlText(fInfo) & "," & _
			sqlText(fLink) & "," & _
			sqlDate(fExpiry) & "," & _
			fPerm & "," & _
			"0,0,0," & _
			fUser & "," & Application("av_DateNow") & "," & _
			fUser & "," & Application("av_DateNow") & ",1)")

	insertArticle = getLastInsert(fUser,8)
	Call updateArticleCount(fCatId,"+")
End Function

Sub updateArticleCount(fCatId,fChange)

Dim fIntLevel

	fIntLevel = fCatId

	Do while CInt(fIntlevel) <> 0
		objConn.Execute("UPDATE KB_Categories SET I_Count = I_Count " & fChange & " 1 WHERE CatId = " & fIntLevel)
		Set objFRS = objConn.Execute("SELECT I_ParentId FROM KB_Categories WHERE CatId = " & fIntLevel)
		If not (objFRS.BOF and objFRS.EOF) Then fIntLevel = objFRS.fields(0).value
	Loop
End Sub

Function rateArticle(fId,fRating)

	rateArticle = "UPDATE KB_Articles " & _
			"SET H_RateCount = H_RateCount + 1, " & _
			"H_RateTotal = H_RateTotal + " & fRating & " " & _
			"WHERE ArticleId = " & fId
End Function

Function getCrumbs(fCat)
	getCrumbs = "SELECT I_ParentId, I_Name FROM KB_Categories WHERE I_Status = 1 AND CatId = " & fCat
End Function
%>