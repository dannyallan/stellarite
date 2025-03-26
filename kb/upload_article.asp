<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\upload.asp" -->
<!--#include file="..\_inc\sql\sql_kb.asp" -->
<%
	Server.ScriptTimeout = 1200

	Call pageFunctions(7,2)

	Dim lngEventId          'as Long
	Dim lngFileSize         'as Long
	Dim strArticleTitle     'as String
	Dim lngCatId            'as Long
	Dim strKeywords         'as String
	Dim strSummary          'as String
	Dim strLink             'as String
	Dim strResultMsg        'as String
	Dim datExpiry           'as Date
	Dim intPermissions      'as Integer
	Dim strDestinationPath  'as String
	Dim strDestinationURL   'as String
	Dim strLogFolder        'as String
	Dim strUploadSizeLimit  'as String
	Dim strUnique           'as String
	Dim objUpload           'as Object
	Dim objFile             'as Object
	Dim objFO               'as Object
	Dim strFileName         'as String
	Dim strContentType      'as String
	Dim strStatus           'as String
	Dim blnUpload           'as Boolean

	If Request.ServerVariables("REQUEST_METHOD") = "GET" or valNum(Request.QueryString("uid"),3,1) = 0 Then
		Call logError(1,1)
	End If

	strTitle             = getIDS("IDS_UploadInfo")
	strUnique            = getUniqueFolderName
	strDestinationPath   = Application("av_UploadPath") & strUnique
	strDestinationURL    = Application("av_UploadURL") & strUnique
	strLogFolder         = Application("av_UploadLog")
	strUploadSizeLimit   = Application("av_UploadLimit")
	blnUpload            = True
	strAction            = "new"

	If Right(strDestinationURL,1) <> "/" Then strDestinationURL = strDestinationURL & "/"
	If Right(strDestinationPath,1) <> "\" Then strDestinationPath = strDestinationPath & "\"

	Set objFO = Server.CreateObject("Scripting.FileSystemObject")
	If Not objFO.FolderExists(strDestinationPath) Then objFO.CreateFolder strDestinationPath
	Set objFO = Nothing

	Function getUniqueFolderName()
		Dim fUploadNumber

		Application.Lock
		If Application("fUploadNumber") = "" Then
			Application("fUploadNumber") = 1
		Else
			Application("fUploadNumber") = Application("fUploadNumber") + 1
		End If
		fUploadNumber = Application("fUploadNumber")
		Application.UnLock

		getUniqueFolderName = Right("0" & Year(Now), 2) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2) & "_" & fUploadNumber
	End Function

	Function getFileName(FullPath)
		Dim Pos, PosF
		PosF = 0
		For Pos = Len(FullPath) To 1 Step -1
			Select Case Mid(FullPath, Pos, 1)
				Case ":", "/", "\": PosF = Pos + 1: Pos = 0
			End Select
		Next
		If PosF = 0 Then PosF = 1
		getFileName = Mid(FullPath, PosF)
	End Function

	Select Case Application("av_UploadType")
		Case "0"
			Set objUpload = New ASPForm
			objUpload.UploadId = valNum(Request.QueryString("uid"),3,1)
			objUpload.SizeLimit = valNum(Application("av_UploadLimit"),3,1)
			objUpload.Files.Save strDestinationPath

			If objUpload.State > 10 Then
				blnUpload = False
				strResultMsg = objUpload.State & ": An Upload Error occurred."
			End If

			If objUpload.State = 13 Then strResultMsg = objUpload.State & ": Upload size " & objUpload.TotalBytes & "B exceeds the limit of " & objUpload.SizeLimit & "B."

			If blnUpload Then
				strArticleTitle = valString(objUpload("txtTitle"),100,0,0)
				lngCatId       = valNum(objUpload("selCategory"),2,-1)
				strKeywords    = valString(objUpload("txtKeywords"),40,0,0)
				strSummary     = valString(objUpload("txtSummary"),255,0,4)
				strLink        = getFileName(objUpload("filFile").FileName)
				intPermissions = valNum(objUpload("selPermissions"),1,1)
				datExpiry      = valDate(objUpload("txtExpiry"),1)
				strContentType = objUpload("filFile").ContentType
				lngFileSize    = objUpload("filFile").Length
				strStatus      = getIDS("IDS_Success")
			End If

		Case "1"
			Set objUpload = Server.CreateObject("Persits.Upload")
			'objUpload.SetMaxSize strUploadSizeLimit, True
			objUpload.Save(strDestinationPath)

			strArticleTitle    = valString(objUpload.Form("txtTitle"),100,0,0)
			lngCatId           = valNum(objUpload.Form("selCategory"),2,-1)
			strKeywords        = valString(objUpload.Form("txtKeywords"),40,0,0)
			strSummary         = valString(objUpload.Form("txtSummary"),255,0,4)
			strLink            = objUpload.Files.Item("filFile").FileName
			intPermissions     = valNum(objUpload.Form("selPermissions"),1,1)
			datExpiry          = valDate(objUpload.Form("txtExpiry"),1)
			strContentType     = objUpload.Files.Item("filFile").ContentType
			lngFileSize        = objUpload.Files.Item("filFile").Size
			strStatus          = getIDS.Form("IDS_Success")

		Case "2"
			Set objUpload = Server.CreateObject("AspSmartUpLoad.SmartUpLoad")
			objUpload.Upload
			objUpload.Save strDestinationPath

			strArticleTitle    = valString(objUpload.Form("txtTitle"),100,0,0)
			lngCatId           = valNum(objUpload.Form("selCategory"),2,-1)
			strKeywords        = valString(objUpload.Form("txtKeywords"),40,0,0)
			strSummary         = valString(objUpload.Form("txtSummary"),255,0,4)
			strLink            = objUpload.Files.Item("filFile").FileName
			intPermissions     = valNum(objUpload.Form("selPermissions"),1,1)
			datExpiry          = valDate(objUpload.Form("txtExpiry"),1)
			strContentType     = objUpload.Files.Item("filFile").ContentType
			lngFileSize        = objUpload.Files.Item("filFile").Size
			strStatus          = getIDS.Form("IDS_Success")

		Case "3"
			Set objUpload = Server.CreateObject("NET2DATABASE.AspFileUp")
			objUpload.Upload

			strArticleTitle    = valString(objUpload.Form("txtTitle"),100,0,0)
			lngCatId           = valNum(objUpload.Form("selCategory"),2,-1)
			strKeywords        = valString(objUpload.Form("txtKeywords"),40,0,0)
			strSummary         = valString(objUpload.Form("txtSummary"),255,0,4)
			strLink            = getFileName(objUpload.Filename("filFile"))
			intPermissions     = valNum(objUpload.Form("selPermissions"),1,1)
			datExpiry          = valDate(objUpload.Form("txtExpiry"),1)
			strContentType     = Upload.Form("filFile").ContentType("filFile")
			lngFileSize        = Upload.Form("filFile").FileLen("filFile")
			strStatus          = getIDS.Form("IDS_Success")

		Case "4"
			Set objUpload = Server.CreateObject("SoftArtisans.FileUp")
			objUpload.Path = strDestinationPath

			strArticleTitle    = valString(objUpload.Form("txtTitle"),100,0,0)
			lngCatId           = valNum(objUpload.Form("selCategory"),2,-1)
			strKeywords        = valString(objUpload.Form("txtKeywords"),40,0,0)
			strSummary         = valString(objUpload.Form("txtSummary"),255,0,4)
			strLink            = objUpload.Form("filFile").ShortFilename
			intPermissions     = valNum(objUpload.Form("selPermissions"),1,1)
			datExpiry          = valDate(objUpload.Form("txtExpiry"),1)
			strContentType     = objUpload.Form("filFile").ContentType
			lngFileSize        = objUpload.Form("filFile").TotalBytes
			strStatus          = getIDS.Form("IDS_Success")
	End Select


	If blnUpload Then
		strLink = strDestinationURL & strLink
		If blnRS Then
			Call updateArticle(lngUserId,lngRecordId,lngCatId,strArticleTitle,strKeywords,strSummary,"",strLink,datExpiry,intPermissions)
		Else
			lngRecordId = insertArticle(lngUserId,lngRecordId,lngCatId,strArticleTitle,strKeywords,strSummary,"",strLink,datExpiry,intPermissions)
		End If
	End If

	If IsObject(objUpload) Then Set objUpload = Nothing

	Call DisplayHeader(1)
	Call showEditHeader(strTitle,strFullName,Now,strFullName,Now)

%>
<div id="contentDiv" class="dvBorder" style="height:355px;"><br />

<span class="hFont"><% =getIDS("IDS_FileName") %>: </span><% =strFileName %><br />
<span class="hFont"><% =getIDS("IDS_FileContent") %>: </span><% =strContentType %><br />
<span class="hFont"><% =getIDS("IDS_FileSize") %>: </span><% =lngFileSize %><br />
<span class="hFont"><% =getIDS("IDS_Status") %>: </span><% =strStatus %><br />

</div>

<div id="footerDiv" class="dvFooter">
<%
	Response.Write(getIconNew(getEditURL("7","")))
	Response.Write(getIconCancel("default.asp"))
%>
</div>

<%
	Call DisplayFooter(1)
%>