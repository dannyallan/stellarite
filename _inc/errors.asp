<%
Sub writeErrorLog()
	Dim fError

	If Application("av_ErrorLog") <> "" or Application("av_ErrorEmail") <> "" Then

		If Err <> 0 Then
			fError = fError & "ASP error: " & Err & " " & Err.Description & vbCrLf
			fError = fError & "Source: " & Err.Source & vbCrLf
			Err.Clear
		End If

		If isObject(objConn) Then
			If objConn.Errors.Count > 0 then
				For each i in objConn.Errors
					If i.Number <> 0 Then fError = fError & "ERROR from ADO: " & i.Description & " (" & i.Number & ")" & vbCrLf
				Next
				objConn.Errors.Clear
			End If
		End If

		If Application("av_ErrorDetailed") = "1" Then

			fError = fError & vbCrLf & "-- Server Variables" & vbCrLf

			For Each eKey in Request.ServerVariables
				fError = fError & eKey & ": " & Request.ServerVariables(eKey) & vbCrLf
			Next
			fError = fError & "POST_DATA: " & Request.Form & vbCrLf

			fError = fError & vbCrLf & "-- Cookies" & vbCrLf

			For Each i in Session.Contents
				fError = fError & i & ": " & Session.Contents(i) & vbCrLf
			Next

			fError = fError & vbCrLf & "-- Time" & vbCrLf & Now & vbCrLf

		Else
			fError = fError & vbCrLf & "-- Specific Details" & vbCrLf

			fError = fError & "ERROR: " & Session("ErrorMessage") & vbCrLf
			fError = fError & "URL: " & Request.ServerVariables("URL") & vbCrLf
			fError = fError & "REFERER: " & Request.ServerVariables("HTTP_REFERER") & vbCrLf
			fError = fError & "REMOTE_IP: " & Request.ServerVariables("REMOTE_ADDR") & vbCrLf
			fError = fError & "REQUEST_METHOD: " & Request.ServerVariables("REQUEST_METHOD") & vbCrLf
			fError = fError & "QUERY_STRING: " & Request.ServerVariables("QUERY_STRING") & vbCrLf
			fError = fError & "POST_DATA: " & Request.Form & vbCrLf
			fError = fError & "USER: " & Session("UserName") & vbCrLf
			fError = fError & "TIME: " & Now & vbCrLf
		End If

		fError = fError & "---------------------------------------------------------" & vbCrLf

		If Application("av_ErrorLog") <> "" Then

			Dim eFSO,eFile,eKey
			Set eFSO = Server.CreateObject("Scripting.FileSystemObject")
			
			' This might not work if the Application variables are not loaded
			Set eFile = eFSO.OpenTextFile(Server.MapPath(Application("av_CRMDir") & "_errors.txt"),8,True)

			eFile.WriteLine(fError)

			eFile.Close
			Set eFile = Nothing
			Set eFSO = Nothing
		End If

		If Application("av_ErrorEmail") <> "" Then
			'Call doSendMail(Application("IDS_ErrorLog"),fError,Application("av_ErrorEmail"))
		End If

	End If

End Sub
%>