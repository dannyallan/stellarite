<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\upload.asp" -->
<!--#include file="..\_inc\sql\sql_attachments.asp" -->
<%
	Server.ScriptTimeout = 1200

	Call pageFunctions(0,2)

	Dim lngEventId          'as Long
	Dim lngFileSize         'as Long
	Dim intDocType          'as Integer
	Dim intPermissions      'as Integer
	Dim intFiles            'as Integer
	Dim strUploadTitle      'as String
	Dim strInfo             'as String
	Dim strResultMsg        'as String
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

	If Request.ServerVariables("REQUEST_METHOD") = "GET" or valNum(bytMod,1,1) = 0 or valNum(lngModId,3,1) = 0 or valNum(Request.QueryString("uid"),3,1) = 0 Then
		Call logError(1,1)
	End If

	strTitle             = getIDS("IDS_UploadInfo")
	lngEventId           = valNum(Request.QueryString("eid"),3,0)
	strUnique            = getUniqueFolderName
	strDestinationPath   = Application("av_UploadPath") & strUnique
	strDestinationURL    = Application("av_UploadURL") & strUnique
	strLogFolder         = Application("av_UploadLog")
	strUploadSizeLimit   = Application("av_UploadLimit")
	blnUpload            = True

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
				intDocType     = valNum(objUpload("selDocType"),2,-1)
				intPermissions = valNum(objUpload("selPermissions"),1,1)
				strUploadTitle = valString(objUpload("txtTitle"),100,0,0)
				strInfo        = valString(objUpload("txtInfo"),255,0,0)
				intFiles       = valNum(objUpload("hdnFiles"),1,1)
			End If

		Case "1"
			Set objUpload = Server.CreateObject("Persits.Upload")
			'objUpload.SetMaxSize strUploadSizeLimit, True
			objUpload.Save(strDestinationPath)

			intDocType         = valNum(objUpload.Form("selDocType"),2,-1)
			intPermissions     = valNum(objUpload.Form("selPermissions"),1,1)
			strUploadTitle     = valString(objUpload.Form("txtTitle"),100,0,0)
			strInfo            = valString(objUpload.Form("txtInfo"),255,0,0)
			intFiles           = valNum(objUpload.Form("hdnFiles"),1,1)

		Case "2"
			Set objUpload = Server.CreateObject("AspSmartUpLoad.SmartUpLoad")
			objUpload.Upload
			objUpload.Save strDestinationPath

			intDocType         = valNum(objUpload.Form("selDocType"),2,-1)
			intPermissions     = valNum(objUpload.Form("selPermissions"),1,1)
			strUploadTitle     = valString(objUpload.Form("txtTitle"),100,0,0)
			strInfo            = valString(objUpload.Form("txtInfo"),255,0,0)
			intFiles           = valNum(objUpload.Form("hdnFiles"),1,1)

		Case "3"
			Set objUpload = Server.CreateObject("NET2DATABASE.AspFileUp")
			objUpload.Upload

			intDocType         = valNum(objUpload.Form("selDocType"),2,-1)
			intPermissions     = valNum(objUpload.Form("selPermissions"),1,1)
			strUploadTitle     = valString(objUpload.Form("txtTitle"),100,0,0)
			strInfo            = valString(objUpload.Form("txtInfo"),255,0,0)
			intFiles           = valNum(objUpload.Form("hdnFiles"),1,1)
		Case "4"
			Set objUpload = Server.CreateObject("SoftArtisans.FileUp")
			objUpload.Path = strDestinationPath

			intDocType         = valNum(objUpload.Form("selDocType"),2,-1)
			intPermissions     = valNum(objUpload.Form("selPermissions"),1,1)
			strUploadTitle     = valString(objUpload.Form("txtTitle"),100,0,0)
			strInfo            = valString(objUpload.Form("txtInfo"),255,0,0)
			intFiles           = valNum(objUpload.Form("hdnFiles"),1,1)
	End Select

	strAction = "new"

	If blnUpload Then lngRecordId = insertAttach(lngUserId,bytMod,lngModId,lngEventId,intDocType,intPermissions,strUploadTitle,strInfo)

	Select Case Application("av_UploadType")

		Case "0"
			If blnUpload Then
				For i = 1 to intFiles

					strFileName = getFileName(objUpload("file" & i).FileName)
					strContentType = objUpload("file" & i).ContentType
					lngFileSize = objUpload("file" & i).Length
					strStatus = getIDS("IDS_Success")

					strResultMsg = strResultMsg & "<span class=""hFont"">" & getIDS("IDS_FileName") & ": </span>" & strFileName & "<br />" & vbCrLf & _
							"<span class=""hFont"">" & getIDS("IDS_FileContent") & ": </span>" & strContentType & "<br />" & vbCrLf & _
							"<span class=""hFont"">" & getIDS("IDS_FileSize") & ": </span>" & lngFileSize & "<br />" & vbCrLf & _
							"<span class=""hFont"">" & getIDS("IDS_Status") & ": </span>" & strStatus & "<hr />" & vbCrLf & vbCrLf

					objConn.Execute(insertAttachLinks(lngRecordId,strDestinationUrl & strFileName))
				Next
			End If

		Case "1","2"
			For each objFile in objUpload.Files.Items

				strFileName = objFile.FileName
				strContentType = objFile.ContentType
				lngFileSize = objFile.Size
				strStatus = getIDS("IDS_Success")

				strResultMsg = strResultMsg & "<span class=""hFont"">" & getIDS("IDS_FileName") & ": </span>" & strFileName & "<br />" & vbCrLf & _
						"<span class=""hFont"">" & getIDS("IDS_FileContent") & ": </span>" & strContentType & "<br />" & vbCrLf & _
						"<span class=""hFont"">" & getIDS("IDS_FileSize") & ": </span>" & lngFileSize & "<br />" & vbCrLf & _
						"<span class=""hFont"">" & getIDS("IDS_Status") & ": </span>" & strStatus & "<hr />" & vbCrLf & vbCrLf

				objConn.Execute(insertAttachLinks(lngRecordId,strDestinationUrl & strFileName))
			Next

		Case "3","4"

			For i = 1 to intFiles

				Select Case Application("av_UploadType")
					Case "3"
						strFileName = getFileName(objUpload.Filename("file" & i))
						strContentType = objUpload.ContentType("file" & i)
						lngFileSize = objUpload.FileLen("file" & i)
						strStatus = getIDS("IDS_Success")
					Case "4"
						strFileName = objUpload.Form("file" & i).ShortFilename
						strContentType = objUpload.Form("file" & i).ContentType
						lngFileSize = objUpload.Form("file" & i).TotalBytes
						strStatus = getIDS("IDS_Success")
				End Select

				objUpload.SaveFile "file" & i, strDestinationPath & strFileName

				strResultMsg = strResultMsg & "<span class=""hFont"">" & getIDS("IDS_FileName") & ": </span>" & strFileName & "<br />" & vbCrLf & _
						"<span class=""hFont"">" & getIDS("IDS_FileContent") & ": </span>" & strContentType & "<br />" & vbCrLf & _
						"<span class=""hFont"">" & getIDS("IDS_FileSize") & ": </span>" & lngFileSize & "<br />" & vbCrLf & _
						"<span class=""hFont"">" & getIDS("IDS_Status") & ": </span>" & strStatus & "<hr />" & vbCrLf & vbCrLf

				objConn.Execute(insertAttachLinks(lngRecordId,strDestinationUrl & strFileName))
			Next
	End Select

	If IsObject(objUpload) Then Set objUpload = Nothing

	Call DisplayHeader(3)
	Call showEditHeader(strTitle,strFullName,Now,strFullName,Now)
%>
<div id="contentDiv" class="dvBorder" style="height:355px;"><br />

<table border=0 cellspacing=5 width="100%">
  <tr>
	 <td class="dFont"><% =strResultMsg %></td>
  </tr>
</table>

</div>

<div id="footerDiv" class="dvFooter">
<%
	Response.Write(getIconNew(getEditURL("A","?m="&bytMod&"&mid="&lngModId)))
	Response.Write(getIconCancel("i_attach.asp?m="&bytMod&"&mid="&lngModId))
%>
</div>

<script language="JavaScript" type="text/javascript">
	window.opener.location.reload();
</script>

<%
	Call DisplayFooter(3)
%>