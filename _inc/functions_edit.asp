<!--#include file="functions.asp" -->
<!--#include file="sql\sql_edit.asp" -->
<%
Sub setAppVar(fVariable,fValue)
	Application(fVariable) = fValue
	Call insertConfigValue(fVariable,fValue)
End Sub

Sub remAppVar(fVariable)
	Application.Contents.Remove(fVariable)
End Sub

Function getUserId(fMod,fVal)
	Dim fModName, fMsg

	fModName = getIDS("IDS_ModName" & fMod)

	If fVal = "" Then
		getUserId = "NULL"
	Else
		Set objFRS = objConn.Execute(getUserIdSQL(fVal))

		If objFRS.BOF and objFRS.EOF Then
			fMsg = "Sorry. " & fVal & " is not a valid user."
		Else
			If fMod = 0 Then
				getUserId = objFRS.fields(0).value
			Else
				If Mid(objFRS.fields(1).value,fMod,1) = "1" Then
					getUserId = objFRS.fields(0).value
				Else
					fMsg = "Sorry. " & fVal & " is not a member of " & fModName & "."
				End If
			End If
		End If

		If fMsg <> "" Then Call sendBack(fMsg)
	End If

End Function

Function getArray(fName,fSQL)
	Dim fNewName

	fNewName = Replace(fName,"IDS","ARR")

	If IsArray(Application(fNewName)) Then
		getArray = Application(fNewName)
	Else
		Set objFRS = objConn.Execute(fSQL)
		If not (objFRS.BOF and objFRS.EOF) Then getArray = objFRS.GetRows()

		If fName <> "NULL" Then
			Application.Lock
			Application(fNewName) = getArray
			Application.Unlock
		End If
	End If

End Function

Function getOptionDropDown(fWidth,fShowBlank,fListName,fOptGroup,fDefault)
	getOptionDropDown = getDropDown(fWidth,fShowBlank,fListName,getArray(fOptGroup,getOptionSQL(fOptGroup)),fDefault)
End Function

Function getInvoiceDropDown(fWidth,fShowBlank,fListName,fDivId,fDefault)
	getInvoiceDropDown = getDropDown(fWidth,fShowBlank,fListName,getArray("NULL",getInvoiceSQL(fDivId)),fDefault)
End Function

Function getCategoryDropDown(fWidth,fShowBlank,fListName,fDefault)
	getCategoryDropDown = getDropDown(fWidth,fShowBlank,fListName,getArray("IDS_Categories",getCategorySQL),fDefault)
End Function


Function getDropDown(fWidth,fShowBlank,fListName,fArray,fDefault)
	Dim fString, fInt

	If not isArray(fArray) then

		fString = "<select name=""" & fListName & """ id=""" & fListName & """ style=""width:" & fWidth & "px;"" class=""oNum"" onChange=""doChange();"">" & vbCrLf
		fString = fString & vbTab & "<option></option>" & vbCrLf
		fString = fString & vbTab & "</select>" & vbCrLf
	Else
		fString = "<select name=""" & fListName & """ id=""" & fListName & """ style=""width:" & fWidth & "px;"" class=""oNum"" onChange=""doChange();"">" & vbCrLf
		If fShowBlank Then fString = fString & vbTab & "<option></option>" & vbCrLf

		For fInt = 0 to UBound(fArray,2)

			fString = fString &  vbTab & "<option value=""" & fArray(0,fInt) & """ " & getDefault(0,fArray(0,fInt),fDefault) & "> " & showString(fArray(1,fInt)) & " </option>" & vbCrLf
		Next

		fString = fString & vbTab & "</select>" & vbCrLf
	End If

	getDropDown = fString
End Function

Function getModuleDropDown(fName,fDefault,fShowAll,fExtra)
	getModuleDropDown = "<select name=""" & fName & """ id=""" & fName & """ style=""width:190px;"" class=""oByte"" " & fExtra & " onChange=""doChange();"">" & vbCrLf

	If fShowAll Then getModuleDropDown = getModuleDropDown & "<option value=""0""" & getDefault(0,fDefault,0) & ">" & getIDS("IDS_ModAll") & "</option>" & vbCrLf
	If pContacts >= 5 Then getModuleDropDown = getModuleDropDown & "<option value=""1""" & getDefault(0,fDefault,1) & ">" & getIDS("IDS_Contacts") & "</option>" & vbCrLf
	If pClients >= 5 Then getModuleDropDown = getModuleDropDown & "<option value=""2""" & getDefault(0,fDefault,2) & ">" & getIDS("IDS_Accounts") & "</option>" & vbCrLf
	If pSales >= 5 Then getModuleDropDown = getModuleDropDown & "<option value=""3""" & getDefault(0,fDefault,3) & ">" & getIDS("IDS_Sales") & "</option>" & vbCrLf
	If pProjects >= 5 Then getModuleDropDown = getModuleDropDown & "<option value=""4""" & getDefault(0,fDefault,4) & ">" & getIDS("IDS_ModName4") & "</option>" & vbCrLf
	If pTickets >= 5 Then getModuleDropDown = getModuleDropDown & "<option value=""5""" & getDefault(0,fDefault,5) & ">" & getIDS("IDS_ModName5") & "</option>" & vbCrLf
	If pBugs >= 5 Then getModuleDropDown = getModuleDropDown & "<option value=""6""" & getDefault(0,fDefault,6) & ">" & getIDS("IDS_ModName6") & "</option>" & vbCrLf
	If pInvoices >= 5 Then getModuleDropDown = getModuleDropDown & "<option value=""7""" & getDefault(0,fDefault,7) & ">" & getIDS("IDS_ModName7") & "</option>" & vbCrLf

	getModuleDropDown = getModuleDropDown & "</select>" & vbCrLf
End Function

Function getPermissionsDropDown(fDefault,fMember)
	If fDefault = "" Then fDefault = 2

	getPermissionsDropDown = "<select name=""selPermissions"" id=""selPermissions"" class=""oByte"" style=""width:150px;"" onChange=""doChange();"">" & vbCrLf

	If fMember = 1 Then getPermissionsDropDown = getPermissionsDropDown & _
			vbTab & "<option value=1" & getDefault(0,1,fDefault) & ">" & getIDS("IDS_MembersOnly") & "</option>" & vbCrLf

	getPermissionsDropDown = getPermissionsDropDown & _
			vbTab & "<option value=2" & getDefault(0,2,fDefault) & ">" & getIDS("IDS_InternalView") & "</option>" & vbCrLf & _
			vbTab & "<option value=3" & getDefault(0,3,fDefault) & ">" & getIDS("IDS_Public") & "</option>" & vbCrLf & _
			"</select>" & vbCrLf
End Function

Sub showEditHeader(fTitle,fCreatedBy,fCreatedDate,fModBy,fModDate)

	Response.Write("<div id=""headerDiv"" class=""dvHeader"">" & vbCrLf)

	If fCreatedBy = "" and fCreatedDate = "" and fModBy = "" and fModDate = "" Then
		Response.Write("<table border=0 cellspacing=0 cellpadding=0 height=35 width=""100%"">" & vbCrLf & _
				"  <tr>" & vbCrLf & _
				"   <td valign=middle class=""tFont"">" & fTitle & "</td>" & vbCrLf & _
				"  </tr>" & vbCrLf & _
				"</table></div>" & vbCrLf & vbCrLf)
	Else
		Response.Write("<table border=0 cellspacing=0 cellpadding=0 height=35 width=""100%"">" & vbCrLf & _
				"  <tr>" & vbCrLf & _
				"   <td valign=middle><img src=""../images/" & strModImage & ".gif"" alt=""" & strModItem & """ width=32 height=32 hspace=10 align=absmiddle /><span class=""tFont"">" & showString(fTitle) & "</span></td>" & vbCrLf & _
				"   <td valign=middle align=right>" & vbCrLf & _
				"     <table border=0 cellspacing=0 cellpadding=0>" & vbCrLf & _
				"       <tr>" & vbCrLf & _
				"         <td class=""dFont"">" & getIDS("IDS_Created") & ":</td>" & vbCrLf & _
				"         <td class=""dFont"">&nbsp;&nbsp;" & showDate(0,fCreatedDate) & "</td>" & vbCrLf & _
				"         <td class=""dFont"">&nbsp;&nbsp;" & showString(fCreatedBy) & "</td>" & vbCrLf & _
				"       </tr>" & vbCrLf & _
				"       <tr>" & vbCrLf & _
				"         <td class=""dFont"">" & getIDS("IDS_Modified") & ":</td>" & vbCrLf & _
				"         <td class=""dFont"">&nbsp;&nbsp;" & showDate(0,fModDate) & "</td>" & vbCrLf & _
				"         <td class=""dFont"">&nbsp;&nbsp;" & showString(fModBy) & "</td>" & vbCrLf & _
				"      </tr>" & vbCrLf & _
				"     </table>" & vbCrLf & _
				"   </td>" & vbCrLf & _
				" </tr>" & vbCrLf & _
				"</table></div>" & vbCrLf & vbCrLf)
	End If

End Sub

Sub closeEdit()

	Dim i1, i2, fUrl, fTemp

	fUrl = showString(Session("LastPage"))
	i1 = Instr(LCase(fUrl),"?id=")

	If i1 <> 0 Then

		fTemp = Left(fUrl,i1)
		i2 = Instr(i1+1,fUrl,"&")

		If i2 = 0 then
			fUrl = fTemp & "id=" & lngRecordId
		Else
			fUrl = fTemp & "id=" & lngRecordId & "&" & Right(fUrl,len(fUrl)-i2)
		End If
	End If

	Call doRedirect(fUrl)
End Sub

Sub closeWindow(fOpener)

	Dim i1, i2, fTemp

	fOpener = showString(fOpener)
	i1 = Instr(LCase(fOpener),"?id=")

	If i1 <> 0 Then

		fTemp = left(fOpener,i1)
		i2 = Instr(i1+1,fOpener,"&")

		If i2 = 0 then
			fOpener = fTemp & "id=" & lngRecordId
		Else
			fOpener = fTemp & "id=" & lngRecordId & "&" & right(fOpener,len(fOpener)-i2)
		End If
	End If

	If fOpener = "" Then fOpener = "refresh"

	Response.Write("<html><head><title>" & getIDS("IDS_CloseWindow") & "</title>" & _
					"<script language=""JavaScript"" type=""text/javascript"" src=""" & strCRMURL & "common/js/crm.js""></script></head>" & _
					"<body onLoad=""closeWindow('" & fOpener & "');""></body></html>")
	Call endResponse()
End Sub

Function getCalendarScripts()
	getCalendarScripts = "<script language=""JavaScript"" type=""text/javascript"" src=""" & Application("av_CRMDir") & "common/calendar/default.js""></script>" & vbCrLf & _
						"<script language=""JavaScript"" type=""text/javascript"" src=""" & Application("av_CRMDir") & "common/calendar/calendar.js""></script>" & vbCrLf & _
						"<script language=""JavaScript"" type=""text/javascript"" src=""" & Application("av_CRMDir") & "common/calendar/lang/" & strLanguage & ".js""></script>" & vbCrLf & _
						"<link href=""" & Application("av_CRMDir") & "common/calendar/calendar.css"" rel=""stylesheet"" type=""text/css"" />"
End Function

Function getEditorScripts()
	getEditorScripts = "<script language=""JavaScript"" type=""text/javascript"" src=""" & Application("av_CRMDir") & "common/editor/default.js""></script>" & vbCrLf & _
						"<script language=""JavaScript"" type=""text/javascript"" src=""" & Application("av_CRMDir") & "common/editor/htmlarea.js""></script>" & vbCrLf & _
						"<script language=""JavaScript"" type=""text/javascript"" src=""" & Application("av_CRMDir") & "common/editor/dialog.js""></script>" & vbCrLf & _
						"<script language=""JavaScript"" type=""text/javascript"" src=""" & Application("av_CRMDir") & "common/editor/popupwin.js""></script>" & vbCrLf & _
						"<script language=""JavaScript"" type=""text/javascript"" src=""" & Application("av_CRMDir") & "common/editor/lang/" & strLanguage & ".js""></script>" & vbCrLf & _
						"<link href=""" & Application("av_CRMDir") & "common/editor/htmlarea.css"" rel=""stylesheet"" type=""text/css"" />"
End Function

Function getLastInsert(fUser,fType)

	Set objFRS = objConn.Execute(getLastInsertSQL(fUser,fType))

	If not (objFRS.BOF and objFRS.EOF) Then
		getLastInsert = objFRS.fields(0).value
	Else
		Call logError(0,1)
	End If
End Function

Sub saveCustomFields(iMod,lRecordId)
	Dim iCount, aDefault, sSQL, sTable, sField, sTemp, iMan

	aDefault = Application("arr_Fields" & iMod)

	For iCount = 0 to UBound(aDefault,2)
		If aDefault(7,iCount) = 0 Then
			If aDefault(5,iCount) = 0 Then iMan = -1
			Select Case aDefault(3,iCount)
				Case 1,6,7
					sField = "txt" & Replace(aDefault(1,iCount)," ","") & aDefault(0,iCount)
					sTemp = sqlText(valString(Request.Form(sField),aDefault(4,iCount),aDefault(5,iCount),0))
				Case 2
					sField = "txt" & Replace(aDefault(1,iCount)," ","") & aDefault(0,iCount)
					Select Case aDefault(4,iCount)
						Case 17
							sTemp = valNum(Request.Form(sField),0,iMan)
						Case 2
							sTemp = valNum(Request.Form(sField),0,iMan)
						Case 3
							sTemp = valNum(Request.Form(sField),0,iMan)
						Case 5
							sTemp = valNum(Request.Form(sField),0,iMan)
						Case 6
							sTemp = valNum(Request.Form(sField),0,iMan)
					End Select
				Case 3
					sField = "txt" & Replace(aDefault(1,iCount)," ","") & aDefault(0,iCount)
					sTemp = sqlDate(valDate(Request.Form(sField),aDefault(5,iCount)))
				Case 4
					sField = "chk" & Replace(aDefault(1,iCount)," ","") & aDefault(0,iCount)
					sTemp = valNum(Request.Form(sField),0,iMan)
				Case 5
					sField = "sel" & Replace(aDefault(1,iCount)," ","") & aDefault(0,iCount)
					sTemp = valNum(Request.Form(sField),2,iMan)
			End Select
			sSQL = sSQL & Mid(aDefault(2,iCount),3) & "=" & sTemp & ","
		End If
	Next

	If Len(sSql) > 0 Then
		sSql = Left(sSQL,Len(sSQL)-1)
		sTable = getSqlInfo(iMod,1)
		sTemp = getSqlInfo(iMod,2)
		objConn.Execute("UPDATE " & Left(sTable,Len(sTable)-2) & " SET " & sSQL & " WHERE " & sTemp & "=" & lRecordId)
	End If
End Sub

Function editCustomFields(iMod)
	Dim aDefault, iCount, sName, sField, sTemp, sValue,sMan

	aDefault = Application("arr_Fields" & iMod)

	For iCount = 0 to UBound(aDefault,2)
		If aDefault(7,iCount) = 0 Then
			If aDefault(5,iCount) = 1 Then sMan = "m" Else sMan = "o"
			If IsObject(objRS) Then
				If not (objRS.BOF and objRS.EOF) Then sValue = objRS.fields(Mid(aDefault(2,iCount),3)).value
			End If

			Select Case aDefault(3,iCount)
				Case 1,6
					sField = "txt" & Replace(aDefault(1,iCount)," ","") & aDefault(0,iCount)
					sTemp = getTextField(sField,sMan & "Text",sValue,20,aDefault(4,iCount),"")
				Case 2
					sField = "txt" & Replace(aDefault(1,iCount)," ","") & aDefault(0,iCount)
					Select Case aDefault(4,iCount)
						Case 17
							sTemp = getTextField(sField,sMan & "Bool",sValue,20,1,"")
						Case 2
							sTemp = getTextField(sField,sMan & "Byte",sValue,20,3,"")
						Case 3
							sTemp = getTextField(sField,sMan & "Int",sValue,20,5,"")
						Case 5
							sTemp = getTextField(sField,sMan & "Long",sValue,20,10,"")
						Case 6
							sTemp = getTextField(sField,sMan & "Currency",sValue,20,20,"")
					End Select
				Case 3
					sField = "txt" & Replace(aDefault(1,iCount)," ","") & aDefault(0,iCount)
					sTemp = getDateField(sField,sMan & "Date",sValue,aDefault(1,iCount))
				Case 4
					sField = "chk" & Replace(aDefault(1,iCount)," ","")
					sTemp = getCheckbox(sField,sValue,"")
				Case 5
					sField = "sel" & Replace(aDefault(1,iCount)," ","") & aDefault(0,iCount)
					sTemp = getOptionDropDown(150,False,sField,aDefault(1,iCount),sValue)
				Case 7
					sField = "txt" & Replace(aDefault(1,iCount)," ","") & aDefault(0,iCount)
					sTemp = getPassword(sField,sMan & "Text",sValue,20,aDefault(4,iCount),"")
			End Select
			sName = getLabel(aDefault(1,iCount),sField)
			editCustomFields = editCustomFields & "  <tr><td>" & sName & "</td><td>" & sTemp & "</td></tr>" & vbCrLf
		End If
	Next
End Function

Function getSqlInfo(iMod,iType)
	Dim sTable, sField

	Select Case iMod
		Case 1
			sTable = "CRM_Contacts K"
			sField = "ContactId"
		Case 2
			sTable = "CRM_Divisions D"
			sField = "DivId"
		Case 3
			sTable = "CRM_Sales S"
			sField = "SaleId"
		Case 4
			sTable = "CRM_Projects P"
			sField = "ProjectId"
		Case 5
			sTable = "CRM_Tickets T"
			sField = "TicketId"
		Case 6
			sTable = "CRM_Bugs B"
			sField = "BugId"
		Case 7
			sTable = "CRM_Invoices I"
			sField = "InvoiceId"
		Case 50
			sTable = "CRM_Events E"
			sField = "EventId"
	End Select

	Select Case iType
		Case 1
			getSqlInfo = sTable
		Case 2
			getSqlInfo = sField
	End Select
End Function

%>