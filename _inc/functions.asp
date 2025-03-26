<!--#include file="common.asp" -->
<!--#include file="header.asp" -->
<!--#include file="footer.asp" -->
<!--#include file="menu.asp" -->
<!--#include file="email.asp" -->
<!--#include file="errors.asp" -->
<!--#include file="sql\sql_default.asp" -->
<%
'##### Generic Functions ##############################################

Function getOrder(fO,fLO,fSO)

	If (CStr(fO) = CStr(fLO) and fSO = "ASC") or (CStr(fLO) = "" and fSO = "DESC") then
		getOrder = "DESC"
	Else
		getOrder = "ASC"
	End If

End Function

Function getDefault(fType,fVar1,fVar2)
	Dim fOut

	If fType = 1 Then fOut = " checked" Else fOut = " selected"

	If Len(fVar1) > 0 and Len(fVar2) > 0 Then
		If CStr(fVar1) = CStr(fVar2) Then getDefault = fOut Else getDefault = ""
	Else
		getDefault = ""
	End If
End Function

Function toggleRowColor(fVal)
	If fVal = "dRow1" Then toggleRowColor = "dRow2" Else toggleRowColor = "dRow1"
End Function

Function bigDigitNum(fNum,fVal)
	Dim fI, fTemp
	If Len(fVal) < CByte(fNum) Then
		For fI = 1 to (CByte(fNum) - Len(fVal))
			fTemp = fTemp & "0"
		Next
		bigDigitNum = fTemp & fVal
	Elseif IsNull(fVal) Then
		bigDigitNum = "NULL"
	Else
		bigDigitNum = fVal
	End If
End Function

Function showDuration(fTime1,fTime2)
	Dim fTemp
	If IsDate(fTime1) and IsDate(fTime2) Then
		fTime1 = CDate(fTime1)
		fTime2 = CDate(fTime2)
		If DateDiff("d",fTime1,fTime2) > 0 Then
			fTemp = DateDiff("d",fTime1,fTime2)
			If fTemp > 1 Then fTemp = fTemp & " days" Else fTemp = fTemp & " day"
		Elseif DateDiff("h",fTime1,fTime2) > 0 Then
			fTemp = DateDiff("h",fTime1,fTime2)
			If fTemp > 1 Then fTemp = fTemp & " hours" Else fTemp = fTemp & " hour"
		Else
			fTemp = DateDiff("n",fTime1,fTime2)
			If fTemp > 1 Then fTemp = fTemp & " minutes" Else fTemp = fTemp & " minute"
		End If
		showDuration = fTemp
	End If
End Function

Function showTrueFalse(fVal)
	showTrueFalse = getIDS("IDS_False")
	If fVal <> "" Then
		If CStr(fVal) = "1" Then showTrueFalse = getIDS("IDS_True")
	End If
End Function

Function showDate(fType,fDate)

	If IsDate(fDate) Then
		fDate = DateAdd("n",Session("TimeOffset"),CDate(fDate))

		showDate = Replace(Replace(Replace(Application("av_DateFormat"),"%Y",Year(fDate)),"%m",bigDigitNum(2,Month(fDate))),"%d",bigDigitNum(2,Day(fDate)))

		If fType >= 1 Then showDate = showDate & " " & bigDigitNum(2,Hour(fDate)) & ":" & bigDigitNum(2,Minute(fDate))
		If fType =  2 Then showDate = showDate & ":" & bigDigitNum(2,Second(fDate))
	Else
		showDate = fDate
	End If
End Function

Function showPhone(fVal)
	If Len(fVal) > 0 Then
		Select Case Len(fVal)
			Case 12
				showPhone = Left(fVal,2) & " " & Mid(fVal,3,2) & " " & Mid(fVal,5,4) & " " & Mid(fVal,9)
			Case 11
				showPhone = Left(fVal,1) & " (" & Mid(fVal,2,3) & ") " & Mid(fVal,5,3) & "-" & Mid(fVal,8)
			Case 10
				showPhone = "(" & Left(fVal,3) & ") " & Mid(fVal,4,3) & "-" & Mid(fVal,7)
			Case 7
				showPhone = Left(fVal,3) & "-" & Mid(fVal,4,4)
			Case Else
				showPhone = FormatNumber(fVal,0,0,0,0)
		End Select
	End If
End Function

Function showLink(fSec,fURL,fContext)
	Dim fTarget

	fContext = showString(fContext)

	If Len(fContext) > 0 and fContext <> "NULL" and fContext <> getIDS("IDS_Deleted") Then
		If getSecurity(fSec,1) Then
			If Instr(fContext,"http") > 0 Then fTarget = "_new" Else fTarget = "_top"
			showLink = "<a href=""" & fURL & """ target=""" & fTarget & """>" & showString(fContext) & "</a>"
		Else
			showLink = showString(fContext)
		End If
	Else
		showLink = "&nbsp;"
	End if
End Function

Function showEmail(fContext)
	fContext = showString(fContext)
	If Len(fContext) > 0 Then showEmail = "<a href=""mailto:" & fContext & """>" & showString(fContext) & "</a>"
End Function

Function showString(fText)
	If Len(fText) > 0 Then
		fText = Trim(fText)
		fText = Replace(fText,"<","&lt;")
		fText = Replace(fText,">","&gt;")
		fText = Replace(fText,"""","&#34;")
		fText = Replace(fText,"'","&#39;")
		showString = fText
	End If
End Function

Function showParagraph(fText)
	If Len(fText) > 0 Then
		fText = showString(fText)
		fText = Replace(fText, vbCrLf, "<br />" & vbCrLf)
		showParagraph = fText
	End If
End Function

Function showHTML(fText)
	'Must allow HTML formatting
	'Should probably try to stop XSS
	showHTML = fText
End Function

Function showCustomFields(iMod)
	Dim aDefault, iCount, sName, sField

	aDefault = Application("arr_Fields" & iMod)

	For iCount = 0 to UBound(aDefault,2)
		If aDefault(7,iCount) = 0 Then
			sName = aDefault(1,iCount)
			If IsObject(objRS) Then sField = objRS.fields(Mid(aDefault(2,iCount),3)).value
			Select Case aDefault(3,iCount)
				Case 1,2,6,7
					sField = showString(sField)
				Case 3
					sField = showDate(0,sField)
				Case 4
					sField = showTrueFalse(sField)
				Case 5
					sField = getAOS(sField)
			End Select
			Response.Write("  <tr><td class=""bFont"">" & sName & "</td><td class=""dFont"">" & showString(sField) & "</td></tr>" & vbCrLf)
		End If
	Next
End Function

Function trimString(fString,fLen)
	If Len(fString) > fLen and fLen > 3 Then
		fString = Left(fString,fLen-3) & "..."
	Elseif Len(fString) > fLen and fLen > 0 Then
		fString = Left(fString,fLen)
	Else
		fString = fString
	End If

	trimString = showString(fString)
End Function

Function getIDS(sName)
	If Left(sName,3) = "IDS" and Application(sName) <> "" Then
		getIDS = showString(Application(sName))
	Else
		getIDS = showString(sName)
	End If
End Function

Function getAOS(iVal)
	If iVal <> "" Then getAOS = showString(Application("ao_Option" & iVal))
End Function

Function getLabel(fText,fFor)
	getLabel = "<label class=""bFont"" for=""" & fFor & """ id=""lbl" & Mid(fFor,4) & """>" & showString(fText) & "</label>"
End Function

Function getTextField(fName,fClass,fValue,fSize,fMax,fExtra)
	getTextField = "<input type=""text"" name=""" & fName & """ id=""" & fName & """ class=""" & fClass & """ value=""" & showString(fValue) & """ size=""" & fSize & """ maxlength=""" & fMax & """ onChange=""doChange();"" " & fExtra & " />"
End Function

Function getFileField(fName,fClass,fValue,fSize,fMax,fExtra)
	getFileField = "<input type=""file"" name=""" & fName & """ id=""" & fName & """ class=""" & fClass & """ value=""" & showString(fValue) & """ size=""" & fSize & """ maxlength=""" & fMax & """ onChange=""doChange();"" " & fExtra & " />"
End Function

Function getDateField(fName,fClass,fValue,fAlt)
	getDateField = getTextField(fName,fClass,showDate(0,fValue),12,12,"") & " " & getIconImport(6,"showCalendar('" & fName & "');",fAlt)
End Function

Function getTextArea(fName,fClass,fValue,fWidth,fRows,fExtra)
	getTextArea = "<textarea name=""" & fName & """ id=""" & fName & """ class=""" & fClass & """ rows=""" & fRows & """ style=""width:" & fWidth & ";"" onChange=""doChange();"" " & fExtra & ">" & showString(fValue) & "</textarea>"
End Function

Function getPassword(fName,fClass,fValue,fSize,fMax,fExtra)
	getPassword = "<input type=""password"" name=""" & fName & """ id=""" & fName & """ class=""" & fClass & """ value=""" & fValue & """ size=""" & fSize & """ maxlength=""" & fMax & """ onChange=""doChange();"" " & fExtra & " />"
End Function

Function getCheckbox(fName,fDefault,fExtra)
	getCheckbox = "<input type=""checkbox"" name=""" & fName & """ id=""" & fName & """ value=""1""" & getDefault(1,1,fDefault) & " onChange=""doChange();"" " & fExtra & " />"
End Function

Function getRadio(fName,fValue,fDefault,fExtra)
	getRadio = "<input type=""radio"" name=""" & fName & """ id=""" & fName & fValue & """ value=""" & fValue & """ onChange=""doChange();""" & getDefault(1,fValue,fDefault) & " " & fExtra & " />"
End Function

Function getSubmit(fName,fValue,fSize,fAccess,fExtra)
	getSubmit = "<input type=""submit"" name=""" & fName & """ id=""" & fName & """ class=""button"" value=""" & showString(fValue) & """ style=""width:" & fSize & "px;"" accessKey=""" & fAccess & """ " & fExtra & " />"
End Function

Function getHidden(fName,fValue)
	getHidden = "<input type=""hidden"" name=""" & fName & """ id=""" & fName & """ value=""" & showString(fValue) & """ />" & vbCrLf
End Function

Sub sendBack(fMsg)
	Call closeConn()
	Response.Write("<html><head><script language=""JavaScript"" type=""text/javascript"" src=""" & Application("av_CRMDir") & "common/js/crm.js""></script>" & _
			"<title>" & getIDS("IDS_ErrorUnspecified") & "</title></head>" & _
			"<body onLoad=""sendBack('" & showString(Replace(fMsg,"'","\'")) & "');""></body></html>")
	Response.End
End Sub

Sub doRedirect(fDest)
	Call closeConn()

	Response.Clear
	Response.Buffer = True
	Response.Status = "302 Found"
	Response.AddHeader "Location",fDest

	fDest = Server.URLEncode(fDest)

	Response.Write("<html><head><title>302 Found</title>" & _
			"<meta http-equiv=""refresh"" content=""0;URL=" & fDest & """ /></head>" & vbCrLf & _
			"<body onLoad=""window.location='" & fDest & "'"">" & vbCrLf & _
			"This page has been moved.  Click <a href=""" & fDest & """>here</a> to continue." & vbCrLf & _
			"</body></html>")
	Response.End
End Sub

'##### Icon and Linking Functions #####################################

Function getEditURL(fType,fString)
	Dim fSymbol

	Select Case CStr(fType)
		Case "1"
			getEditURL = "sales/edit_contact.asp"
		Case "2"
			getEditURL = "sales/edit_client.asp"
		Case "3"
			getEditURL = "sales/edit_sale.asp"
		Case "4"
			getEditURL = "services/edit_project.asp"
		Case "5"
			getEditURL = "support/edit_ticket.asp"
		Case "6"
			getEditURL = "qa/edit_bug.asp"
		Case "7"
			getEditURL = "finance/edit_invoice.asp"
		Case "8"
			getEditURL = "kb/edit_article.asp"
		Case "50"
			getEditURL = "common/edit_event.asp"
		Case "A"
			getEditURL = "common/edit_attach.asp"
		Case "B"
			getEditURL = "common/edit_subscriptions.asp"
		Case "C"
			getEditURL = "kb/edit_category.asp"
		Case "N"
			getEditURL = "common/edit_notes.asp"
		Case "O"
			getEditURL = "admin/edit_portal.asp"
		Case "P"
			getEditURL = "admin/edit_permissions.asp"
		Case "S"
			getEditURL = "common/pop_search.asp"
		Case "U"
			getEditURL = "admin/edit_profile.asp"
		Case "W"
			getEditURL = "admin/edit_password.asp"
		Case "Z"
			getEditURL = "common/edit_product.asp"
	End Select

	getEditURL = Application("av_CRMDir") & getEditURL & fString

	Select Case CStr(fType)
		Case "1","3","4","5","6","7","50"
			If fString = "" Then fSymbol = "?" Else fSymbol = "&"
			getEditURL = getEditURL & fSymbol & "menu=" & bytMenu
	End Select

End Function

Function getSearchURL(fString)
	getSearchURL = "openWindow('" & getEditURL("S",fString) & "','sw_Search','500','400');"
End Function

Function getEditLink(fType,fString,fContext)
	getEditLink = "<a href=""" & getEditURL(fType,fString) & """>" & showString(fContext) & "</a>"
End Function

Function getPopLink(fType,fString,fContext)
	getPopLink = "<a href=""" & getEditURL(fType,fString) & """ onClick=""" & getSearchURL(fString) & "return false;"">" & showString(fContext) & "</a>"
End Function

Function getIconImport(ByVal fType,ByVal fAction,ByVal fAlt)
	Dim fImg, fLinkText
	Select Case fType
		Case 1
			fImg = "import2.gif"
			fLinkText = getIDS("IDS_Import")
		Case 2
			fImg = "edit2.gif"
			fLinkText = getIDS("IDS_Edit")
		Case 3
			'ONLY USED FOR NEW CONTACTS (Modules 2,6,7)
			fImg = "new2.gif"
			fLinkText = getIDS("IDS_New")
			fAction = "openWindow('" & fAction & "','sw_Contact','500','600');"
		Case 4
			fImg = "del2.gif"
			fLinkText = getIDS("IDS_Delete")
		Case 5
			fImg = "perm2.gif"
			fLinkText = getIDS("IDS_Permissions")
		Case 6
			fImg = "cal2.gif"
			fLinkText = getIDS("IDS_Import")
			fAction = "Javascript:" & fAction
	End Select

	If Left(fAction,10) = "openWindow" Then fAction = "Javascript:" & fAction
	fAlt = fLinkText & ": " & fAlt

	Select Case CStr(Application("av_Navigation"))
		Case "1"
			getIconImport = " <span class=""dFont"">[<a href=""" & fAction & """ title=""" & showString(fAlt) & """ onMouseOver=""(window.status='" & showString(fAlt) & "');return true;"" onMouseOut=""(window.status='');return true;"">" & showString(fLinkText) & "</a>]</span>"
		Case Else
			getIconImport = " <a href=""" & fAction & """ onMouseOver=""(window.status='" & showString(fAlt) & "');return true;"" onMouseOut=""(window.status='');return true;""><img src=""../images/" & fImg & """ alt=""" & showString(fAlt) & """ border=0 height=16 width=16 valign=absmiddle /></a>"
	End Select

End Function


Function getIconNew(fLink)
	getIconNew = getIcon(fLink,"N","new.gif",getIDS("IDS_New"))
End Function

Function getIconEdit(fLink)
	getIconEdit = getIcon(fLink,"O","edit.gif",getIDS("IDS_Edit"))
End Function

Function getIconExport(fLink)
	getIconExport = getIcon(fLink,"E","export.gif",getIDS("IDS_Export"))
End Function

Function getIconSave(fAction)
	getIconSave = getIcon("Javascript:confirmAction('" & fAction & "');","S","save.gif",getIDS("IDS_Save"))
End Function

Function getIconDelete()
	getIconDelete = getIcon("Javascript:confirmAction('del');","D","del.gif",getIDS("IDS_Delete"))
End Function

Function getIconCancel(fAction)
	Select Case fAction
		Case "back"
			fAction = "history.back();"
		Case "close"
			fAction = "closeWindow(null);"
		Case Else
			fAction = "document.location.href='" & fAction & "';"
	End Select
	getIconCancel = getIcon("Javascript:confirmAction('canc');" & fAction,"X","cancel.gif",getIDS("IDS_Cancel"))
End Function

Function getIconNext(fLink)
	getIconNext = getIcon(fLink,".","next.gif",getIDS("IDS_Next"))
End Function

Function getIconPrev(fLink)
	getIconPrev = getIcon(fLink,",","prev.gif",getIDS("IDS_Previous"))
End Function

Function getIconPrint(fLink)
	getIconPrint = getIcon(fLink,"P","print.gif",getIDS("IDS_Print"))
End Function

Function getIconSearch(fLink)
	getIconSearch = getIcon(fLink,"F","find.gif",getIDS("IDS_Search"))
End Function

Function getSpacer(fHeight,fWidth)
	getSpacer = vbTab & "<img src=""" & Application("av_CRMDir") & "images/spacer.gif"" alt="""" border=""0"" height=""" & fHeight & """ width=""" & fWidth & """ />" & vbCrLf
End Function

Function getIcon(fAction,fAccess,fIcon,fAlt)

	If Left(UCase(fAction),10) = "OPENWINDOW" Then
		fAction = fAction & "return false;"
	ElseIf Left(UCase(fAction),11) = "JAVASCRIPT:" Then
		fAction = Mid(fAction,12) & "return false;"
	Else
		fAction = "document.location.href='" & fAction & "';return false;"
	End If

	Select Case CStr(Application("av_Navigation"))
		Case "1"
			getIcon = vbTab & "<input type=""submit"" name=""btn" & Replace(fAlt," ","") & """ class=""button"" value=""" & showString(fAlt) & """ onClick=""" & fAction & """ accessKey=""" & fAccess & """ onMouseOver=""(window.status='" & showString(fAlt) & "');return true;"" onMouseOut=""(window.status='');return true;"" />" & vbCrLf
		Case Else
			getIcon = vbTab & "<input type=""image"" name=""img" & Replace(fAlt," ","") & """ src=""" & Application("av_CRMDir") & "images/" & fIcon & """ alt=""" & showString(fAlt) & """ border=""0"" height=""24"" width=""24"" hspace=""2"" onClick=""" & fAction & """ accessKey=""" & fAccess & """ onMouseOver=""(window.status='" & fAlt & "');return true;"" onMouseOut=""(window.status='');return true;"" />" & vbCrLf
	End Select
End Function

'##### Connection Functions ###########################################

Sub openConn()
	If objConn <> "" then
		Call closeConn()
		openConn
	Else
		Set objConn = Server.CreateObject("ADODB.Connection")
		objConn.Open strConnString
	End If
End Sub

Sub closeConn()
	If isObject(objRS) then
		objRS.close
		Set objRS = nothing
	End If

	If isObject(objFRS) then
		objFRS.close
		Set objFRS = nothing
	End If

	If isObject(objConn) then
		objConn.close
		Set objConn = nothing
	End If
End Sub

Sub updateMasterRecord(fType,fMod,fModId)
	Dim fSQL,aTables,aFields,aTableId,fMath

	If valNum(fModId,1,0) = 0 Then Exit Sub
	If valNum(fMod,1,0) = 0 Then Exit Sub
	If strDoAction = "new" Then fMath = "+" Else fMath="-"
	'On Error Resume Next

	aTables = split(" |Contacts|Divisions|Sales|Projects|Tickets|Bugs|Invoices","|")
	aFields = split(" |K_|D_|S_|P_|T_|B_|I_","|")
	aTableId = split(" |ContactId|DivId|SaleId|ProjectId|TicketId|BugId|InvoiceId","|")

	fSQL = "UPDATE CRM_" & aTables(fMod) & " SET " & _
		aFields(fMod) & "ModBy = " & lngUserId & "," & _
		aFields(fMod) & "ModDate = " & Application("av_DateNow") & _
		" WHERE " & aTableId(fMod) & " = " & fModId

	objConn.Execute(fSQL)

	If strDoAction <> "edit" Then

		fSQL = "UPDATE CRM_" & aTables(fMod) & " SET " & _
			aFields(fMod) & fType & " = " & _
			getValue(aFields(fMod)&fType,"CRM_"&aTables(fMod),aTableId(fMod)&"="&fModId,0) & fMath & "1" & _
			" WHERE " & aTableId(fMod) & " = " & fModId

		objConn.Execute(fSQL)
	End If

	Call doNotification(fMod,fModId)

End Sub

Sub endResponse()
	Call closeConn()
	Response.End
End Sub

Sub logError(fType,fShow)
	Select Case fType
		Case 1
			Session("ErrorMessage") = getIDS("IDS_ErrorBadRequest")
		Case 2
			Session("ErrorMessage") = getIDS("IDS_ErrorPermissions")
		Case 3
			Session("ErrorMessage") = getIDS("IDS_ErrorBadRecord")
		Case 4
			Session("ErrorMessage") = getIDS("IDS_MsgLockedOutUser")
		Case Else
			Session("ErrorMessage") = getIDS("IDS_ErrorUnspecified")
	End Select

	Call writeErrorLog()
	Session.Contents.Remove("ErrorMessage")

	If fShow = 1 Then
		Call closeConn()
		Server.Transfer Application("av_CRMDir") & "404.asp"
	End If
End Sub

'##### User Functions #################################################

Function getSecurity(fMod,fLevel)
	getSecurity = False
	Select Case fMod
		Case 1
			If pContacts >= fLevel Then getSecurity = True
		Case 2
			If pClients >= fLevel Then getSecurity = True
		Case 3
			If pSales >= fLevel Then getSecurity = True
		Case 4
			If pProjects >= fLevel Then getSecurity = True
		Case 5
			If pTickets >= fLevel Then getSecurity = True
		Case 6
			If pBugs >= fLevel Then getSecurity = True
		Case 7
			If pInvoices >= fLevel Then getSecurity = True
		Case 8
			If pArticles >= fLevel Then getSecurity = True
		Case Else
			If Instr(Session("Permissions"),"5") > 0 Then getSecurity = True
	End Select
End Function

Sub pageFunctions(fMod,fLevel)

	Dim fStop, fThisMod

	intMember = 2
	bytRealMod = fMod

	strDir = Application("av_CRMDir")
	lngRecordId = valNum(Request.QueryString("id"),3,-1)
	If lngRecordId = "NULL" Then lngRecordId = "" Else lngRecordId = Clng(lngRecordId)

	bytMod = valNum(Request.QueryString("m"),1,-1)
	If bytMod = "NULL" Then bytMod = CByte(fMod) Else bytMod = CByte(bytMod)

	Select Case fMod
		Case 0, 50
			fThisMod = bytMod
		Case Else
			fThisMod = fMod
	End Select

	lngModId = valNum(Request.QueryString("mid"),3,-1)
	If lngModId = "NULL" Then lngModId = lngRecordId Else lngModId = CLng(lngModId)

	SELECT Case fThisMod
		Case 90
			'Administration Module
			If pContacts < 5 and pClients < 5 and pSales < 5 and pSales < 5 and pProjects < 5 and pTickets < 5 and pBugs < 5 and pInvoices < 5 and pArticles < 5 and not blnAdmin Then fStop = True
			If blnAdmin Then intPerm = 5
			strDir = strDir & "admin/"
			strModItem = getIDS("IDS_Administration")
			strModName = getIDS("IDS_Administration")
		Case 1
			'Contacts Module
			If pContacts >= fLevel then
				intPerm = pContacts
				If mContacts or blnAdmin Then intMember = 1
				strDir = strDir & "sales/"
				strModImage = "contact"
				strModItem = getIDS("IDS_ModItem1")
				strModName = getIDS("IDS_ModName123")
			Else
				fStop = True
			End If
		Case 2
			'Clients Module
			If pClients >= fLevel then
				intPerm = pClients
				If mClients or blnAdmin Then intMember = 1
				strDir = strDir & "sales/"
				strModImage = "client"
				strModItem = getIDS("IDS_ModItem2")
				strModName = getIDS("IDS_ModName123")
			Else
				fStop = True
			End If
		Case 3
			'Sales Module
			If pSales >= fLevel then
				intPerm = pSales
				If mSales or blnAdmin Then intMember = 1
				strDir = strDir & "sales/"
				strModImage = "sale"
				strModItem = getIDS("IDS_ModItem3")
				strModName = getIDS("IDS_ModName123")
			Else
				fStop = True
			End If
		Case 4
			'Services Module
			If pProjects >= fLevel then
				intPerm = pProjects
				If mProjects or blnAdmin Then intMember = 1
				strDir = strDir & "services/"
				strModImage = "project"
				strModItem = getIDS("IDS_ModItem4")
				strModName = getIDS("IDS_ModName4")
			Else
				fStop = True
			End If
		Case 5
			'Support Module
			If pTickets >= fLevel then
				intPerm = pTickets
				If mTickets or blnAdmin Then intMember = 1
				strDir = strDir & "support/"
				strModImage = "ticket"
				strModItem = getIDS("IDS_ModItem5")
				strModName = getIDS("IDS_ModName5")
			Else
				fStop = True
			End If
		Case 6
			'QA Module
			If pBugs >= fLevel then
				intPerm = pBugs
				If mBugs or blnAdmin Then intMember = 1
				strDir = strDir & "qa/"
				strModImage = "bug"
				strModItem = getIDS("IDS_ModItem6")
				strModName = getIDS("IDS_ModName6")
			Else
				fStop = True
			End If
		Case 7
			'Finance
			If pInvoices >= fLevel then
				intPerm = pInvoices
				If mInvoices or blnAdmin Then intMember = 1
				strDir = strDir & "finance/"
				strModImage = "invoice"
				strModItem = getIDS("IDS_ModItem7")
				strModName = getIDS("IDS_ModName7")
			Else
				fStop = True
			End If
		Case 8
			'Knowledge Base
			If pArticles >= fLevel then
				intPerm = pArticles
				If mArticles or blnAdmin Then intMember = 1
				strDir = strDir & "kb/"
				strModImage = "article"
				strModItem = getIDS("IDS_ModItem8")
				strModName = getIDS("IDS_ModName8")
			Else
				fStop = True
			End If
		Case Else
			'Generic Module
			If (pContacts < fLevel) and (pClients < fLevel) and (pSales < fLevel) and (pProjects < fLevel) and (pTickets < fLevel) and (pBugs < fLevel) and (pArticles < fLevel) and (pInvoices < fLevel) Then fStop = True

	End Select


	If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

		If Instr(LCase(Request.ServerVariables("HTTP_REFERER")),LCase(Request.ServerVariables("HTTP_HOST"))) = 0 Then
			fStop = True
		End If

		If intMode > 0 Then
			Call sendBack(getIDS("IDS_MsgMaintDisabled"))
		End If

		bytMenu = valNum(Request.QueryString("menu"),1,0)

	End If

	If Instr(LCase(Request.ServerVariables("URL")),"upload_") = 0 Then
		strDoAction = valString(Request.Form("hdnAction"),10,0,0)
	End If

	If Request.ServerVariables("HTTP_REFERER") <> "" and Instr(Request.ServerVariables("HTTP_REFERER"),Request.ServerVariables("URL")) = 0 Then
		Session("LastPage") = Request.ServerVariables("HTTP_REFERER")
	End If

	If lngRecordId = "" then
		blnRS = False
		strAction = "new"
	Else
		blnRS = True
		strAction = "edit"
	End If

	If fStop Then Call logError(2,1)
End Sub

'##### Input Testing Functions #######################################

Function regExTest(fVal,fPattern)
	Dim oRegEx

	Set oRegEx           = New RegExp
	oRegEx.IgnoreCase    = True
	oRegEx.Global        = True
	oRegEx.Pattern       = fPattern

	regExTest = oRegEx.Test(fVal)

End Function

Function valNum(fVal,fLen,fType)
	If Len(fVal) > 0 and IsNumeric(fVal) Then                       'Valid Number

		If fLen = 0 and fVal = 0 or fVal = 1 Then                   'Boolean    (0 - 1)
			valNum = retNum(fVal,0)
		Elseif fLen = 1 and fVal >= 0 and fVal <= 255 Then          'Byte       (0 - 255)
			valNum = retNum(fVal,1)
		Elseif fLen = 2 and Abs(fVal) < 32767 Then                  'Integer    Abs(0 - 32,767)
			valNum = retNum(fVal,2)
		Elseif fLen = 3 and Abs(fVal) < 2147483647 Then             'Long       Abs(0 - 2,147,483,647)
			valNum = retNum(fVal,3)
		Elseif fLen = 4 and Len(fVal) <= 15 Then                    'Double     Phone numbers only
			valNum = retNum(fVal,4)
		Elseif fLen = 5 and Abs(fVal) < 922337203685477.5807 Then   'Decimal    Currency values
			valNum = retNum(fVal,5)
		Else
			Call logError(1,1)
		End If

	Elseif Len(fVal) = 0 and CStr(fType) = "-1" Then    'No number, Return NULL
		valNum = "NULL"
	Elseif Len(fVal) = 0 and CStr(fType) = "0" Then     'No number, Return 0
		valNum = retNum(0,fLen)
	Else                                                'Mandatory number missing
		Call logError(1,1)
	End If
End Function

Function valString(fVal,fLen,fMan,fType)
	If CStr(fMan) = "1" and Len(fVal) = 0 Then                          'Mandatory Value missing
		Call logError(1,1)
	Elseif not Len(fVal) = 0 and Len(fVal) > fLen and fLen > 0 Then     'Value exceeds field length
		Call logError(1,1)
	Elseif Len(fVal) = 0 Then                                           'Empty String
		valString = ""
	Else                                                                'Test string type
		Dim fPattern

		Select Case fType
			Case 1    'Email
				fPattern = "^[\w\d\.\%-]+@[\w\d\.\%-]+\.\w{2,4}$"
			Case 2    'Link
				fPattern = "^(https?|ftp|[A-Za-z]:\\|\\\\)[^<>()'""]+$"
			Case 3    'RGB Value
				fPattern = "^#[\dA-F]{6}$"
			Case 4  'Memo input
				'255 character scrolling text box allowing linefeed
			Case 5  'HTML input
				'Rich text input from HTMLArea field allowing linefeed
			Case Else
				'Free text field input
		End Select

		If regExTest(fVal,fPattern) Then
			valString = fVal
		Else
			Call logError(1,1)
		End If
	End If
End Function

Function valDate(fVal,fMan)
	If IsDate(fVal) Then                            'Valid Date
		valDate = fVal
	Elseif Len(fVal) > 0 Then                       'Invalid Date
		Call logError(1,1)
	Elseif CStr(fMan) = 1 and Len(fVal) = 0 Then    'Missing Mandatory Date
		Call logError(1,1)
	End If
End Function

Function retNum(fVal,fLen)
	Select Case fLen
		Case 0,1
			retNum = CByte(fVal)
		Case 2
			retNum = CInt(fVal)
		Case 3
			retNum = CLng(fVal)
		Case 4
			retNum = CDbl(fVal)
		Case 5
			retNum = CCur(fVal)
		Case Else
			Call logError(1,1)
	End Select
End Function
%>