<%
	Option Explicit
	Response.Buffer = True
	Session.Timeout = 10
	Server.ScriptTimeout = 300

%>
<!--#include file="config.asp" -->
<%

	'#########################################################################
	'Check if CRM is in maintenance mode
	'#########################################################################
	If intMode = 2 and Instr(LCase(strCRMURL & "default.asp"),LCase(Request.ServerVariables("SCRIPT_NAME"))) = 0 Then
		Call doRedirect(strCRMURL & "default.asp")
	End If

	'#########################################################################
	'Send the user back to the login screen if the Session has expired
	'#########################################################################

	If Session("UserId") = "" Then
		If Instr(LCase(strCRMURL & "default.asp"),LCase(Request.ServerVariables("SCRIPT_NAME"))) = 0 and Instr(LCase(strCRMURL & "admin/upgrade.asp"),LCase(Request.ServerVariables("SCRIPT_NAME"))) = 0 Then
			If Len(Request.QueryString) > 0 Then
				Call doRedirect(strCRMURL & "default.asp?dest=" & Server.URLEncode(Request.ServerVariables("URL")) & "?" & Server.URLEncode(Request.QueryString))
			Else
				Call doRedirect(strCRMURL & "default.asp?dest=" & Server.URLEncode(Request.ServerVariables("URL")))
			End If
		End If
'	Elseif (Session("screenH") < 500 or Session("screenW") < 700) Then
'		Call doRedirect(strCRMURL & "require.asp?prob=sr")
	End If

	'#########################################################################
	'Variables used on every page
	'#########################################################################

	Dim intDBVer       'as Integer      'Database version used for upgrade
	Dim objConn        'as Object       'Common DB connection on every page
	Dim objRS          'as Object       'Common RS used on every page
	Dim objFRS         'as Object       'Common RS used by functions
	Dim arrRS          'as Array        'RS is dumped into this array on most pages
	Dim blnRS          'as Boolean      'Boolean variable to determine if looking at an existing record
	Dim i, j           'as Integer      'Generic counters

	Dim strFullName    'as String       'String text of current user
	Dim lngUserId      'as Long         'UserId of current user
	Dim blnAdmin       'as Boolean      'Is the current user an admin
	Dim intMember      'as Integer      'Member of Module or not ...
	Dim intPerm        'as Integer      'Permission value of current page (0,1,2,3,4,5)

	Dim pContacts, pClients, pSales, pProjects, pTickets, pBugs, pArticles, pInvoices        'as Byte
	Dim mContacts, mClients, mSales, mProjects, mTickets, mBugs, mArticles, mInvoices        'as Boolean

	Dim lngRecordId    'as Long         'Current Record ID
	Dim lngPrevId      'as Long         'Previous Record ID
	Dim lngNextId      'as Long         'Next Record ID
	Dim bytModCount    'as Integer      'Number of Modules

	Dim strTab         'as Integer      'Numerical value of selected tab
	Dim strTabURL      'as String       'URL of loaded tab
	Dim strTabBuilder  'as String       'String containing all tab URLs

	Dim bytMod         'as Byte         'Current Module (1,2,3,4,5,6,7,8,9,50,90)
	Dim bytRealMod     'as Byte         'True module (if nested)
	Dim lngModId       'as Long         'Record ID in module
	Dim strModName     'as String       'Category Name
	Dim strModItem     'as String       'Name of current module
	Dim strModImage    'as String       'URL of module image

	Dim strTitle       'as String       'Title of page
	Dim strIncHead     'as String       'Add text to the page head
	Dim bytMenu        'as Byte         'Menu Showing = 1, No Menu = 0

	Dim strOpenerURL   'as String       'URL of the opening window
	Dim strReferURL    'as String       'Referring URL
	Dim strDir         'as String       'String value of CRM directory
	Dim strAction      'as String       'Action "to do" on this page
	Dim strDoAction    'as String       'Action requested on form submission

	Dim intScreenH     'as Integer      'Integer value of screen height
	Dim intScreenW     'as Integer      'Integer value of screen width
	Dim intStartTime   'as Integer      'Page generation time

	intDBVer      = 28
	intStartTime  = Timer
	lngPrevId     = CLng(0)
	lngNextId     = CLng(0)
	bytModCount   = 8

	'#########################################################################
	'Locally used functions called in this file only
	'#########################################################################

	Function getHexColor(fVal,fType)
		Dim str1, str2

		str1 = Left(fVal,1)
		str2 = Right(fVal,1)
		If Not IsNumeric(str1) Then str1 = getDecVal(UCase(str1))
		If Not IsNumeric(str2) Then str2 = getDecVal(UCase(str2))

		fVal = (str1 * 16) + str2

		If fType = 1 Then
			If CInt(fVal)+ 51 > 255 Then getHexColor = "FF" Else getHexColor = bigDigitNum(2,Hex(CInt(fVal)+51))
		Elseif fType = 2 Then
			If CInt(fVal)- 51 < 0 Then getHexColor = "00" Else getHexColor = bigDigitNum(2,Hex(CInt(fVal)-51))
		Else
			getHexColor = bigDigitNum(2,Hex(CInt(fVal)))
		End if
	End Function

	Function getDecVal(fVal)
		Select Case fVal
			Case "A"
				getDecVal = 10
			Case "B"
				getDecVal = 11
			Case "C"
				getDecVal = 12
			Case "D"
				getDecVal = 13
			Case "E"
				getDecVal = 14
			Case "F"
				getDecVal = 15
			Case Else
				getDecVal = 5
		End Select
	End Function

	'#########################################################################
	'Session Variables for each user
	'#########################################################################

	If Session("UserId") <> "" Then

		lngUserId    = Session("UserId")
		strFullName  = showString(Session("UserName"))
		intScreenH   = Session("screenH")
		intScreenW   = Session("screenW")

		pContacts    = CByte(valNum(Mid(Session("Permissions"),1,1),1,0))
		pClients     = CByte(valNum(Mid(Session("Permissions"),2,1),1,0))
		pSales       = CByte(valNum(Mid(Session("Permissions"),3,1),1,0))
		pProjects    = CByte(valNum(Mid(Session("Permissions"),4,1),1,0))
		pTickets     = CByte(valNum(Mid(Session("Permissions"),5,1),1,0))
		pBugs        = CByte(valNum(Mid(Session("Permissions"),6,1),1,0))
		pInvoices    = CByte(valNum(Mid(Session("Permissions"),7,1),1,0))
		pArticles    = CByte(valNum(Mid(Session("Permissions"),8,1),1,0))

		If Session("Admin") Then blnAdmin = True Else blnAdmin = False
		If valNum(Mid(Session("Member"),1,1),0,0) = 1 Then mContacts = True Else mContacts = False
		If valNum(Mid(Session("Member"),2,1),0,0) = 1 Then mClients = True Else mClients = False
		If valNum(Mid(Session("Member"),3,1),0,0) = 1 Then mSales = True Else mSales = False
		If valNum(Mid(Session("Member"),4,1),0,0) = 1 Then mProjects = True Else mProjects = False
		If valNum(Mid(Session("Member"),5,1),0,0) = 1 Then mTickets = True Else mTickets = False
		If valNum(Mid(Session("Member"),6,1),0,0) = 1 Then mBugs = True Else mBugs = False
		If valNum(Mid(Session("Member"),7,1),0,0) = 1 Then mInvoices = True Else mInvoices = False
		If valNum(Mid(Session("Member"),8,1),0,0) = 1 Then mArticles = True Else mArticles = False
	End If

	Call openConn()

	If Application("myCRMConfigLoaded") <> "Yes" Then

		'#################################################################
		'Retrieve and set Option variables from the database
		'#################################################################

		Set objFRS = objConn.Execute(getOptionValues(0))

		Do while not objFRS.EOF
			Application("ao_Option" & objFRS.fields("OptionId").value) = showString(objFRS.fields("O_Value").value)
			objFRS.MoveNext
		Loop

		'#################################################################
		'Retrieve and set Config variables from the database
		'#################################################################

		Set objFRS = objConn.Execute(getConfigValues)

		Application.Lock

		Do while not objFRS.EOF
			Application(objFRS.fields("F_Variable").value) = showString(objFRS.fields("F_Value").value)
			objFRS.MoveNext
		Loop

		If Application("av_MajorColor") = "" Then Application("av_MajorColor") = "#990000"
		If Application("av_MinorColor") = "" Then Application("av_MinorColor") = "#FAEBBE"
		If Application("av_LoginAttempts") = "" Then Application("av_LoginAttempts") = "5"
		If Application("av_Modules") = "" Then Application("av_Modules") = "11111111"

		Application("av_MajorColorLight")   = "#" & getHexColor(Mid(Application("av_MajorColor"),2,2),1) & getHexColor(Mid(Application("av_MajorColor"),4,2),1) & getHexColor(Right(Application("av_MajorColor"),2),1)
		Application("av_MajorColorDark")    = "#" & getHexColor(Mid(Application("av_MajorColor"),2,2),2) & getHexColor(Mid(Application("av_MajorColor"),4,2),2) & getHexColor(Right(Application("av_MajorColor"),2),2)
		Application("av_DateFormat")        = Replace(Replace(Replace(CStr(DateSerial(2000,10,30)),"2000","%Y"),"10","%m"),"30","%d")
		Application("av_CRMDir")            = Mid(strCRMURL,Instr(9,strCRMURL,"/"))

		For i = 1 to bytModCount
			If valNum(Mid(Application("av_Modules"),i,1),0,0) = 1 Then Application("av_Module" & i & "On") = True Else Application("av_Module" & i & "On") = False
		Next

		'#################################################################
		'Set the Database specific variables
		'#################################################################

		If strDatabase = "Access" Then
			Application("av_DateDel")     = "#"
			Application("av_DateNow")     = "NOW()"
			Application("av_Substring")   = "MID"
			Application("av_Concat")      = "&"
		Elseif strDatabase = "MSSQL" Then
			Application("av_DateDel")     = "'"
			Application("av_DateNow")     = "GETDATE()"
			Application("av_Substring")   = "SUBSTRING"
			Application("av_Concat")      = "+"
		Elseif strDatabase = "MySQL" Then
			Application("av_DateDel")     = "'"
			Application("av_DateNow")     = "NOW()"
			Application("av_Substring")   = "SUBSTRING"
			Application("av_Concat")      = "&"
		Elseif strDatabase = "Oracle" Then
			Application("av_DateDel")     = "'"
			Application("av_DateNow")     = "(SELECT SYSDATE FROM dual)"
			Application("av_Substring")   = "SUBSTR"
			Application("av_Concat")      = "||"
		End If

		'#################################################################
		'Retrieve the language textfile to set the internationalized data strings
		'#################################################################

		Dim objFSO            'as Object
		Dim objTextStream     'as Object
		Dim strStrings        'as String

		Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		Set objTextStream = objFSO.OpenTextFile(Server.MapPath(Application("av_CRMDir") & "_inc/strings-" & strLanguage & ".txt"),1)
		strStrings = objTextStream.ReadAll
		objTextStream.Close
		Set objFSO = Nothing

		If strStrings <> "" Then arrRS = split(strStrings,vbCrLf)

		For i = 0 to UBound(arrRS)
			Application(Left(arrRS(i),Instr(arrRS(i),",")-1)) = Mid(arrRS(i),Instr(arrRS(i),",")+1)
		Next

		'#################################################################
		'Upgrade the database if needed
		'#################################################################

		If intDBVer <> CInt(Application("av_DatabaseVersion")) Then
			Application("myCRMConfigLoaded") = "Yes"
			Application.Unlock
			Call doRedirect(strCRMURL & "admin/upgrade.asp")
		End If

		'#################################################################
		'Set the Module field data for reporting
		'#################################################################

		Set objFRS = objConn.Execute(getModuleFields(1))
		If not (objFRS.BOF and objFRS.EOF) Then Application("arr_Fields1") = objFRS.GetRows()

		Set objFRS = objConn.Execute(getModuleFields(2))
		If not (objFRS.BOF and objFRS.EOF) Then Application("arr_Fields2") = objFRS.GetRows()

		Set objFRS = objConn.Execute(getModuleFields(3))
		If not (objFRS.BOF and objFRS.EOF) Then Application("arr_Fields3") = objFRS.GetRows()

		Set objFRS = objConn.Execute(getModuleFields(4))
		If not (objFRS.BOF and objFRS.EOF) Then Application("arr_Fields4") = objFRS.GetRows()

		Set objFRS = objConn.Execute(getModuleFields(5))
		If not (objFRS.BOF and objFRS.EOF) Then Application("arr_Fields5") = objFRS.GetRows()

		Set objFRS = objConn.Execute(getModuleFields(6))
		If not (objFRS.BOF and objFRS.EOF) Then Application("arr_Fields6") = objFRS.GetRows()

		Set objFRS = objConn.Execute(getModuleFields(7))
		If not (objFRS.BOF and objFRS.EOF) Then Application("arr_Fields7") = objFRS.GetRows()

		Set objFRS = objConn.Execute(getModuleFields(50))
		If not (objFRS.BOF and objFRS.EOF) Then Application("arr_Fields50") = objFRS.GetRows()


		Application("myCRMConfigLoaded") = "Yes"
		Application.Unlock

	End If
%>