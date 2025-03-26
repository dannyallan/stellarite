<%
' Pure-ASP upload v. 2.04 (with progress bar)
' This software is a FreeWare with limited use.
' You can use this software to upload files with
' size up to 10MB for free.  If you want to upload
' bigger files, please register ScriptUtilities.
' http://www.pstruh.cz/help/scptutl/upload.asp

Const adTypeBinary = 1
Const adTypeText = 2

Const MaxLicensedLimit = &HA00000

Const xfsCompleted      = &H0
Const xfsNotPost        = &H1
Const xfsZeroLength     = &H2
Const xfsInProgress     = &H3
Const xfsNone           = &H5
Const xfsError          = &HA
Const xfsNoBoundary     = &HB
Const xfsUnknownType    = &HC
Const xfsSizeLimit      = &HD
Const xfsTimeOut        = &HE
Const xfsNoConnected    = &HF
Const xfsErrorBinaryRead= &H10

Class ASPForm
	Private m_ReadTime

	Public ChunkReadSize, BytesRead, TotalBytes, UploadID
	Public TempPath, MaxMemoryStorage, CharSet, FormType, SourceData, ReadTimeout

	Public Default Property Get Item(Key)
		Set Item = m_Items.Item(Key)
	End Property

	Public Property Get Items
		Read
		Set Items = m_Items
	End Property

	Public Property Get Files
		Read
		Set Files = m_Items.Files
	End Property

	Public Property Get Texts
		Read
		Set Texts = m_Items.Texts
	End Property

	Public Property Get NewUploadID
		Randomize
		NewUploadID = CLng(rnd * &H7FFFFFFF)
	End Property

	Public Property Get ReadTime
		If IsEmpty(m_ReadTime) Then
			If not IsEmpty(StartUploadTime) Then ReadTime = CLng((Now() - StartUploadTime) * 86400 * 1000)
		Else
			ReadTime = m_ReadTime
		End If
	End Property

	Public Property Get State
		If m_State = xfsNone Then Read
		State = m_State
	End Property

	Private Function CheckRequestProperties
		If UCase(Request.ServerVariables("REQUEST_METHOD")) <> "POST" Then
			m_State = xfsNotPost
			Exit Function
		End If

		Dim CT : CT = Request.ServerVariables("HTTP_CONTENT_TYPE")
		If Len(CT) = 0 Then CT = Request.ServerVariables("CONTENT_TYPE")

		 If LCase(Left(CT, 19)) <> "multipart/form-data" Then
			m_State = xfsUnknownType
			Exit Function
		End If

		Dim PosB : PosB = InStr(LCase(CT), "boundary=")

		If PosB = 0 Then
			m_State = xfsNoBoundary
			Exit Function
		End If

		If PosB > 0 Then Boundary = Mid(CT, PosB + 9)

		PosB = InStr(LCase(CT), "boundary=")
		If PosB > 0 Then
			PosB = InStr(Boundary, ",")
			If PosB > 0 Then Boundary = Left(Boundary, PosB - 1)

		End If

		On Error Resume Next
		TotalBytes = Request.TotalBytes

		If Err <> 0 Then
			TotalBytes = CLng(Request.ServerVariables("HTTP_Content_Length"))
			If Len(TotalBytes)=0 Then TotalBytes = CLng(Request.ServerVariables("CONTENT_LENGTH"))
		End If

		If TotalBytes = 0 Then
			m_State = xfsZeroLength
			Exit Function
		End If

		If IsInSizeLimit(TotalBytes) Then
			CheckRequestProperties = True
			m_State = xfsInProgress
		Else
			m_State = xfsSizeLimit
		End If

	End Function

	Public Sub Read()
		If m_State <> xfsNone Then Exit Sub

		If Not CheckRequestProperties Then
			WriteProgressInfo
			Exit Sub
		End If

		If IsEmpty(bSourceData) Then Set bSourceData = CreateObject("ADODB.Stream")
		bSourceData.Open
		bSourceData.Type = 1

		Dim DataPart, PartSize
		BytesRead = 0
		StartUploadTime = Now

		Do While BytesRead < TotalBytes

			PartSize = ChunkReadSize
			If PartSize + BytesRead > TotalBytes Then PartSize = TotalBytes - BytesRead
			DataPart = Request.BinaryRead(PartSize)
			BytesRead = BytesRead + PartSize

			bSourceData.Write DataPart

			WriteProgressInfo

			If Not Response.IsClientConnected Then
				m_State = xfsNoConnected
				Exit Sub
			End If
		Loop
		m_State = xfsCompleted

		ParseFormData
	End Sub

	Private Sub ParseFormData
		Dim Binary
		bSourceData.Position = 0
		Binary = bSourceData.Read

		m_Items.mpSeparateFields Binary, Boundary
	End Sub

	Public Function getForm(FormID)
		If IsEmpty(ProgressFile.UploadID) Then
			ProgressFile.UploadID = FormID
		End If

		Dim ProgressData

		ProgressData = ProgressFile

		If Len(ProgressData) > 0 Then
			If ProgressData = "DONE" Then
				ProgressFile.Done
				Err.Raise 1, "getForm", "Upload was done"
			Else
				ProgressData = Split (ProgressData, vbCrLf)
				If ubound(ProgressData) = 3 Then
					m_State = CLng(ProgressData(0))
					TotalBytes = CLng(ProgressData(1))
					BytesRead = CLng(ProgressData(2))
					m_ReadTime = CLng(ProgressData(3))
				End If
			End If
		End If
		Set getForm = Me
	End Function

	Private Sub WriteProgressInfo
		If UploadID > 0 Then
			If IsEmpty(ProgressFile.UploadID) Then
				ProgressFile.UploadID = UploadID
			End If

			Dim ProgressData, FileName
			ProgressData = m_State & vbCrLf & TotalBytes & vbCrLf & BytesRead & vbCrLf & ReadTime
			ProgressFile.Contents = ProgressData
		End If
	End Sub

	Private Sub Class_Initialize()
		ChunkReadSize = &H10000
		SizeLimit = &H100000

		BytesRead = 0
		m_State = xfsNone

		TotalBytes = Request.TotalBytes

		Set ProgressFile = New cProgressFile
		Set m_Items = New cFormFields
	End Sub

	Private Sub Class_Terminate()
		If UploadID > 0 Then
			ProgressFile.Contents = "DONE"
		End If
	End Sub

	Private Function IsInSizeLimit(TotalBytes)
		IsInSizeLimit = (m_SizeLimit = 0 or m_SizeLimit > TotalBytes) and (MaxLicensedLimit > TotalBytes)
	End Function

	Public Property Get SizeLimit
		SizeLimit = m_SizeLimit
	End Property

	Public Property Let SizeLimit(NewLimit)
		If NewLimit > MaxLicensedLimit Then
			Err.Raise 1, "ASPForm - SizeLimit", "This version of Pure-ASP upload is licensed with maximum limit of 10MB (" & MaxLicensedLimit & "B)"
			m_SizeLimit = MaxLicensedLimit
		Else
			m_SizeLimit = NewLimit
		End If
	End Property

	Public Boundary
	Private m_Items
	Private m_State
	Private m_SizeLimit
	Private bSourceData
	Private StartUploadTime , TempFiolder
	Private ProgressFile
End Class

Class cFormFields
	Dim m_Keys()
	Dim m_Items()
	Dim m_Count

	Public Default Property Get Item(ByVal Key)

		If vartype(Key) = vbInteger or vartype(Key) = vbLong Then

			If Key < 1 or Key > m_Count Then Err.raise "Index out of bounds"
			Set Item = m_Items(Key-1)
			Exit Property
		End If

		Dim Count
		Count = ItemCount(Key)
		Key = LCase(Key)

		If Count > 0 Then
			If Count>1 Then

				Dim OutItem, ItemCounter
				Set OutItem = New cFormFields
				ItemCounter = 0

				For ItemCounter = 0 To Ubound(m_Keys)
					If LCase(m_Keys(ItemCounter)) = Key Then OutItem.Add Key, m_Items(ItemCounter)
				Next

				Set Item = OutItem
			Else
				For ItemCounter = 0 To Ubound(m_Keys)
					If LCase(m_Keys(ItemCounter)) = Key Then Exit For
				Next

				If IsObject (m_Items(ItemCounter)) Then
					Set Item = m_Items(ItemCounter)
				Else
					Item = m_Items(ItemCounter)
				End If

			End If
		Else
			Set Item = New cFormField
		End If
	End Property

	Public Property Get MultiItem(ByVal Key)

		Dim Out: Set Out = New cFormFields
		Dim I, vItem
		Dim Count
		Count = ItemCount(Key)

		If Count = 1 Then

			Out.Add Key, Item(Key)

		Elseif Count > 1 Then

			For Each I In Item(Key).Items
				Out.Add Key, I
			Next
		End If

		Set MultiItem = Out
	End Property

	Public Property Get Value
		Dim I, V
		For Each I in m_Items
			V = V & ", " & I
		Next
		V = Mid(V, 3)
		Value = V
	End Property

	Public Property Get xA_NewEnum
		Set xA_NewEnum = m_Items
	End Property

	Public Property Get Items()
		Items = m_Items
	End Property

	Public Property Get Keys()
		Keys = m_Keys
	End Property

	Public Property Get Files
		Dim cItem, OutItem, ItemCounter
		Set OutItem = New cFormFields
		ItemCounter = 0
		If m_Count > 0 Then
			For ItemCounter = 0 To Ubound(m_Keys)
				Set cItem = m_Items(ItemCounter)
				If cItem.IsFile Then
					OutItem.Add m_Keys(ItemCounter), m_Items(ItemCounter)
				End If
			Next
		End If
		Set Files = OutItem
	End Property

	Public Property Get Texts
		Dim cItem, OutItem, ItemCounter
		Set OutItem = New cFormFields
		ItemCounter = 0

		For ItemCounter = 0 To Ubound(m_Keys)
			Set cItem = m_Items(ItemCounter)
			If Not cItem.IsFile Then
				OutItem.Add m_Keys(ItemCounter), m_Items(ItemCounter)
			End If
		Next
		Set Texts = OutItem
	End Property

	Public Sub Save(Path)
		Dim Item
		For Each Item In m_Items
			If Item.isFile Then
				Item.Save Path
			End If
		Next
	End Sub

	Public Property Get ItemCount(ByVal Key)

		Dim cKey, Counter
		Counter = 0
		Key = LCase(Key)
		For Each cKey In m_Keys
			If LCase(cKey) = Key Then Counter = Counter + 1
		Next
		ItemCount = Counter
	End Property

	Public Property Get Count()
		Count = m_Count
	End Property

	Public Sub Add(byval Key, Item)
		Key = "" & Key
		ReDim Preserve m_Items(m_Count)
		ReDim Preserve m_Keys(m_Count)
		m_Keys(m_Count) = Key
		Set m_Items(m_Count) = Item
		m_Count = m_Count + 1
	End Sub

	Private Sub Class_Initialize()
		Dim vHelp()

		On Error Resume Next
		m_Items = vHelp
		m_Keys = vHelp
		m_Count = 0
	End Sub

	Public Sub mpSeparateFields(Binary, ByVal Boundary)
		Dim PosOpenBoundary, PosCloseBoundary, PosEndOfHeader, isLastBoundary

		Boundary = "--" & Boundary
		Boundary = StringToBinary(Boundary)

		PosOpenBoundary = InStrB(Binary, Boundary)
		PosCloseBoundary = InStrB(PosOpenBoundary + LenB(Boundary), Binary, Boundary, 0)

		Do While (PosOpenBoundary > 0 And PosCloseBoundary > 0 And Not isLastBoundary)

			Dim HeaderContent, bFieldContent

			Dim Content_Disposition, FormFieldName, SourceFileName, Content_Type

			Dim TwoCharsAfterEndBoundary

			PosEndOfHeader = InStrB(PosOpenBoundary + Len(Boundary), Binary, StringToBinary(vbCrLf + vbCrLf))

			HeaderContent = MidB(Binary, PosOpenBoundary + LenB(Boundary) + 2, PosEndOfHeader - PosOpenBoundary - LenB(Boundary) - 2)

			bFieldContent = MidB(Binary, (PosEndOfHeader + 4), PosCloseBoundary - (PosEndOfHeader + 4) - 2)

			GetHeadFields BinaryToString(HeaderContent), FormFieldName, SourceFileName, Content_Disposition, Content_Type

			Dim Field
			Set Field = New cFormField

			Field.ByteArray = MultiByteToBinary(bFieldContent)

			Field.Name = FormFieldName
			Field.ContentDisposition = Content_Disposition
			If not IsEmpty(SourceFileName) Then
				Field.FilePath = SourceFileName
				Field.FileName = GetFileName(SourceFileName)
			Else
			End If
			Field.ContentType = Content_Type

			Add FormFieldName, Field

			TwoCharsAfterEndBoundary = BinaryToString(MidB(Binary, PosCloseBoundary + LenB(Boundary), 2))
			isLastBoundary = TwoCharsAfterEndBoundary = "--"

			If Not isLastBoundary Then
				PosOpenBoundary = PosCloseBoundary
				PosCloseBoundary = InStrB(PosOpenBoundary + LenB(Boundary), Binary, Boundary)
			End If
		Loop
	End Sub
End Class

Class cProgressFile
	Private fs
	Public TempFolder
	Public m_UploadID
	Public TempFileName

	Public Default Property Get Contents()
		Contents = GetFile(TempFileName)
	End Property

	Public Property Let Contents(inContents)
		WriteFile TempFileName, inContents
	End Property

	Public Sub Done
		FS.DeleteFile TempFileName
	End Sub

	Public Property Get UploadID()
		UploadID = m_UploadID
	End Property

	Public Property Let UploadID(inUploadID)
		If IsEmpty(FS) Then Set fs = CreateObject("Scripting.FileSystemObject")
		TempFolder = fs.GetSpecialFolder(2)

		m_UploadID = inUploadID
		TempFileName = TempFolder & "\pu" & m_UploadID & ".~tmp"

		Dim DateLastModified
		On Error Resume Next
		DateLastModified = fs.GetFile(TempFileName).DateLastModified
		on error goto 0
		If IsEmpty(DateLastModified) Then
		Elseif Now-DateLastModified > 1 Then
		FS.DeleteFile TempFileName
		End If
	End Property

	Private Function GetFile(Byref FileName)

		Dim InStream
		On Error Resume Next
		Set InStream = fs.OpenTextFile(FileName, 1)
		GetFile = InStream.ReadAll
		On Error Goto 0
	End Function

	Private Function WriteFile(Byref FileName, Byref Contents)

		Dim OutStream
		On Error Resume Next
		Set OutStream = fs.OpenTextFile(FileName, 2, True)
		OutStream.Write Contents
	End Function

	Private Sub Class_Initialize()
	End Sub
End Class

Class cFormField

	Public ContentDisposition, ContentType, FileName, FilePath, Name
	Public ByteArray

	Public CharSet, HexString, InProgress, SourceLength, RAWHeader, Index, ContentTransferEncoding

	Public Default Property Get String()

		String = BinaryToString(ByteArray)
	End Property

	Public Property Get IsFile()
		IsFile = not IsEmpty(FileName)
	End Property

	Public Property Get Length()
		Length = LenB(ByteArray)
	End Property

	Public Property Get Value()
		Set Value = Me
	End Property

	Public Sub Save(Path)
		If IsFile Then
			Dim fullFileName
			fullFileName = Path & "\" & FileName
			SaveAs fullFileName
		Else
			Err.raise "Text field " & Name & " does not have a file name"
		End If
	End Sub

	Public Sub SaveAs(newFileName)
		If Len(ByteArray)>0 Then SaveBinaryData newFileName, ByteArray
	End Sub

End Class

Function StringToBinary(String)
	Dim I, B
	For I=1 to Len(String)
		B = B & ChrB(Asc(Mid(String,I,1)))
	Next
	StringToBinary = B
End Function

Function BinaryToString(Binary)

	Dim TempString

	On Error Resume Next

	TempString = RSBinaryToString(Binary)
	If Len(TempString) <> LenB(Binary) Then

		TempString = MBBinaryToString(Binary)
		End If
	BinaryToString = TempString
End Function

Function MBBinaryToString(Binary)

	Dim cl1, cl2, cl3, pl1, pl2, pl3
	Dim L
	cl1 = 1
	cl2 = 1
	cl3 = 1
	L = LenB(Binary)

	Do While cl1<=L
		pl3 = pl3 & Chr(AscB(MidB(Binary,cl1,1)))
		cl1 = cl1 + 1
		cl3 = cl3 + 1
		If cl3>300 Then
			pl2 = pl2 & pl3
			pl3 = ""
			cl3 = 1
			cl2 = cl2 + 1
			If cl2>200 Then
				pl1 = pl1 & pl2
				pl2 = ""
				cl2 = 1
			End If
		End If
	Loop
	MBBinaryToString = pl1 & pl2 & pl3
End Function

Function RSBinaryToString(xBinary)

	Dim Binary

	If vartype(xBinary) = 8 Then Binary = MultiByteToBinary(xBinary) Else Binary = xBinary

	Dim RS, LBinary
	Const adLongVarChar = 201
	Set RS = CreateObject("ADODB.Recordset")
	LBinary = LenB(Binary)

	If LBinary>0 Then
		RS.Fields.Append "mBinary", adLongVarChar, LBinary
		RS.Open
		RS.AddNew
			RS("mBinary").AppendChunk Binary
		RS.Update
		RSBinaryToString = RS("mBinary")
	Else
		RSBinaryToString = ""
	End If
End Function

Function MultiByteToBinary(MultiByte)

	Dim RS, LMultiByte, Binary
	Const adLongVarBinary = 205
	Set RS = CreateObject("ADODB.Recordset")
	LMultiByte = LenB(MultiByte)
	If LMultiByte>0 Then
		RS.Fields.Append "mBinary", adLongVarBinary, LMultiByte
		RS.Open
		RS.AddNew
			RS("mBinary").AppendChunk MultiByte & ChrB(0)
		RS.Update
		Binary = RS("mBinary").GetChunk(LMultiByte)
	End If
	MultiByteToBinary = Binary
End Function

Function GetHeadFields(ByVal Head, Name, FileName, Content_Disposition, Content_Type)

	Name = (SeparateField(Head, "name=", ";"))

	If Left(Name, 1) = """" Then Name = Mid(Name, 2, Len(Name) - 2)

	FileName = (SeparateField(Head, "filename=", ";"))

	If Left(FileName, 1) = """" Then FileName = Mid(FileName, 2, Len(FileName) - 2)

	Content_Disposition = LTrim(SeparateField(Head, "content-disposition:", ";"))
	Content_Type = LTrim(SeparateField(Head, "content-type:", ";"))
End Function

Function SeparateField(From, ByVal sStart, ByVal sEnd)
	Dim PosB, PosE, sFrom
	sFrom = LCase(From)
	PosB = InStr(sFrom, sStart)
	If PosB > 0 Then
		PosB = PosB + Len(sStart)
		PosE = InStr(PosB, sFrom, sEnd)
		If PosE = 0 Then PosE = InStr(PosB, sFrom, vbCrLf)
		If PosE = 0 Then PosE = Len(sFrom) + 1
		SeparateField = Mid(From, PosB, PosE - PosB)
	Else
		SeparateField = Empty
	End If
End Function

Function SplitFileName(FullPath)
  Dim Pos, PosF
	PosF = 0
	For Pos = Len(FullPath) To 1 Step -1
		Select Case Mid(FullPath, Pos, 1)
			Case ":", "/", "\": PosF = Pos + 1: Pos = 0
		End Select
	Next
	If PosF = 0 Then PosF = 1
	SplitFileName = PosF
End Function

Function GetPath(FullPath)
	GetPath = left(FullPath, SplitFileName(FullPath)-1)
End Function

Function GetFileName(FullPath)
	GetFileName = Mid(FullPath, SplitFileName(FullPath))
End Function

Function RecurseMKDir(ByVal Path)
	Dim FS: Set FS = CreateObject("Scripting.FileSystemObject")

	Path = Replace(Path, "/", "\")
	If Right(Path, 1) <> "\" Then Path = Path & "\"
	Dim Pos, n
	Pos = 0: n = 0
	Pos = InStr(Pos + 1, Path, "\")
	Do While Pos > 0
		On Error Resume Next
		FS.CreateFolder Left(Path, Pos - 1)
		If Err = 0 Then n = n + 1
		Pos = InStr(Pos + 1, Path, "\")
	Loop
	RecurseMKDir = n
End Function

Function SaveBinaryData(FileName, ByteArray)
	SaveBinaryData = SaveBinaryDataStream(FileName, ByteArray)
End Function

Function SaveBinaryDataTextStream(FileName, ByteArray)
	  Dim FS : Set FS = CreateObject("Scripting.FileSystemObject")
	On Error Resume Next

	  Dim TextStream
	Set TextStream = FS.CreateTextFile(FileName)

	If Err = &H4c Then
		On error Goto 0
		RecurseMKDir GetPath(FileName)
		On Error Resume Next
		Set TextStream = FS.CreateTextFile(FileName)
	End If

	TextStream.Write BinaryToString(ByteArray)
	TextStream.Close

	Dim ErrMessage, ErrNumber
	ErrMessage = Err.Description
	ErrNumber = Err

	On Error Goto 0
	If ErrNumber<>0 Then Err.Raise ErrNumber, "SaveBinaryData", FileName & ":" & ErrMessage

End Function

Function SaveBinaryDataStream(FileName, ByteArray)
	Dim BinaryStream
	Set BinaryStream = CreateObject("ADODB.Stream")
	BinaryStream.Type = 1
	BinaryStream.Open
	BinaryStream.Write ByteArray
	On Error Resume Next

	BinaryStream.SaveToFile FileName, 2

	If Err = &Hbbc Then
		On error Goto 0
		RecurseMKDir GetPath(FileName)
		On Error Resume Next
		BinaryStream.SaveToFile FileName, 2
	End If
	Dim ErrMessage, ErrNumber

	ErrMessage = Err.Description
	ErrNumber = Err

	On Error Goto 0
	If ErrNumber<>0 Then Err.Raise ErrNumber, "SaveBinaryData", FileName & ":" & ErrMessage

End Function

Class cResponse
	Public Property Get IsClientConnected
		Randomize
		IsClientConnected = cbool(CLng(rnd * 4))
		IsClientConnected = True
	End Property
End Class

Class cRequest
	Private Readed
	Private BinaryStream

	Public Function ServerVariables(Name)
		select case UCase(Name)
			Case "CONTENT_TYPE":
			Case "HTTP_CONTENT_TYPE":
				ServerVariables = "multipart/form-data; boundary=---------------------------7d21960404e2"
			Case "CONTENT_LENGTH":
			Case "HTTP_CONTENT_LENGTH":
				ServerVariables = "" & TotalBytes
			Case "REQUEST_METHOD":
				ServerVariables = "POST"
		End Select
	End Function

	Public Function BinaryRead(ByRef Bytes)
		If Bytes <= 0 Then Exit Function

		If Readed + Bytes > TotalBytes Then Bytes = TotalBytes - Readed
		BinaryRead = BinaryStream.Read(Bytes)
	End Function

	Public Property Get TotalBytes
		TotalBytes = BinaryStream.Size
	End Property

	Private Sub Class_Initialize()
		Set BinaryStream = CreateObject("ADODB.Stream")
		BinaryStream.Type = 1
		BinaryStream.Open
		BinaryStream.LoadFromFile "C:\temp\upload.txt"
		BinaryStream.Position = 0
		Readed = 0
	End Sub
End Class

%>