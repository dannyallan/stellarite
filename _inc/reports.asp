<%

Dim sRptMsg           'as String         'Contains any messages
Dim iRptType          'as Integer        'Type of report
Dim iRptLoop          'as Integer        'Loop being analyzed
Dim iRptLoopCount     'as Integer        'Total number of loops
Dim sRptColor         'as String         'Color or the row

iRptLoop = 0
iRptLoopCount = 0

Sub execReport(sQuery,sRepType)
	Dim bStop, sSQL, sRow

	Select Case sRepType
		Case "print" : iRptType = 0
		Case "csv"   : iRptType = 2
		Case "xls"   : iRptType = 3
		Case Else    : iRptType = 1
	End Select

	On Error Resume Next
	iRptLoop = iRptLoop + 1
	iRptLoopCount = 0

	Call splitSQL(sQuery)

	Execute "iRptLoop" & iRptLoop & " = 0 : " & _
		"If Len(sSQL" & iRptLoop & ") <> 0 Then : " & _
			"If sSQL" & iRptLoop & " = ""?"" Then : " & _
				"Set objRS" & iRptLoop & " = objConn.OpenSchema(20) : " & _
			"Else : If sSQL" & iRptLoop & " = ""??"" Then : " & _
				"Set objRS" & iRptLoop & " = objConn.OpenSchema(4) : " & _
			"Else : " & _
				"Set objRS" & iRptLoop & " = objConn.Execute(sSQL" & iRptLoop & ") : " & _
			"End If : End If : " & _
			"sSql = sSQL" & iRptLoop & " : " & _
		"End If"

	If Err <> 0 or objConn.Errors.Count <> 0 Then
		bStop = True
	Elseif SQLis(sSQL,"ALTER") or SQLis(sSQL,"UPDATE") or SQLis(sSQL,"DELETE") or SQLis(sSQL,"DROP") or SQLis(sSQL,"CREATE") or SQLis(sSQL,"INSERT") Then
		bStop = True
	End If

	If iRptLoop = 1 and not bStop Then

		If not (objRS1.BOF and objRS1.EOF) Then
			arrRS = objRS1.GetRows()
			objRS1.MoveFirst

			If iRptType < 2 Then
				If CInt(UBound(arrRS,2)+1) < CInt(Application("av_MaxRecords")) Then
					Response.Write("<p class=""dFont"">" & Application("IDS_Total") & ": <b>" & uBound(arrRS,2)+1 & "</b> " & Application("IDS_Records") & ".</p>" & vbCrLf)
				Else
					Response.Write("<p class=""dFont"" style=""color:red;""><b>" & Replace(Application("IDS_MsgMaxRecords"),"[MAX]",Application("av_MaxRecords")) & "</b></p>" & vbCrLf)
				End If
			End If

			If iRptType = 1 Then Response.Write("</div><div id=""contentDiv"" class=""dvBorder"" style=""height:" & intScreenH-150 & "px;""><div id=""reportDiv"" class=""dvNoBorder"" style=""height:" & intScreenH-170 & "px;"">")
		End If
	End If

	If not bStop Then Execute "If objRS" & iRptLoop & ".EOF Then Call writeNoRecords() "

	If not bStop Then
		If iRptLoop=1 and iRptType=3 Then
			Response.Write("<html>" & vbCrLf & "<body><table style=""border-collapse:collapse;"">" & vbCrLf)
			If iRptLoopCount > 1 Then Call writeHeaders()
		End If

		Execute "Do while not objRS" & iRptLoop & ".EOF" & " : " & _
				"sRow = """" : " & _
				"iRptLoop" & iRptLoop & " = iRptLoop" & iRptLoop & " + 1 : " & _
				"iRptLoop" & iRptLoop & "Indent = iRptLoop" & iRptLoop-1 & "Indent : " & _
				"If (iRptLoop" & iRptLoop & "=1 or iRptLoop < iRptLoopCount) and iRepType < 2 Then : writeHeaders : End If : " & _
				"For each i in objRS" & iRptLoop & ".fields : " & _
					"If not (iRptLoopCount > 1 and UCase(Right(i.Name,2)) =""ID"") Then : " & _
						"iRptLoop" & iRptLoop & "Indent = iRptLoop" & iRptLoop & "Indent + 1 : " & _
						"sRow = sRow & getCell(i.Name,i.Value,i.ActualSize,i.Type) : " & _
					"End If : " & _
				"Next : " & _
				"Call writeLine(iRptLoop" & iRptLoop-1 & "Indent,sRow) : " & _
				"If iRptLoop < iRptLoopCount Then : Call execReport(sQuery,sRepType) : End If : " & _
				"objRS" & iRptLoop & ".MoveNext : " & _
				"If (iRptLoop = iRptLoopCount and objRS" & iRptLoop & ".EOF) and iRptType < 2 Then : Response.Write(""</table><br />"" & vbCrLf & vbCrLf) : End If : " & _
			"Loop : "

		If iRptType = 3 and objRS1.EOF Then Response.Write("</table>" & vbCrLf)
	End If
	iRptLoop = iRptLoop - 1

	Call showMsg()
	If iRptType = 1 and iRptLoop = 0 Then Response.Write("</div></div>" & vbCrLf)

End Sub


Function cleanSQL(ByVal sSQL)
	While Left(sSQL,1)=vbLF or Left(sSQL,1)=vbCR or Left(sSQL,1)=" "
		sSQL = Mid(sSQL,2)
	Wend

	While Right(sSQL,1)=vbLF or Right(sSQL,1)=vbCR or Right(sSQL,1)=" "
		sSQL = Left(sSQL,Len(sSQL)-1)
	Wend

	cleanSQL = Replace(sSQL,vbCrLf," ")
End Function


Sub splitSQL(ByVal sSQL)
	Dim iPos

	iPos = InStr(sSQL,";;")

	If iPos = 0 Then
		Call buildSQL(sSQL)
	Else
		Call buildSQL(Left(sSQL,iPos-1))
		Call splitSQL(LTrim(Mid(sSQL,iPos+2)))
	End If
End Sub


Sub buildSQL(ByVal sSQL)

	sSQL = cleanSQL(sSQL)

	If sSQL <> "" Then
		iRptLoopCount = iRptLoopCount + 1
		If Instr(sSQL,Chr(34)) = 0 Then sSQL = Chr(34) & Replace(sSQL,vbCrLf," ") & Chr(34)

		If iRptLoop = iRptLoopCount Then Execute "sSQL" & iRptLoopCount & " = " & sSQL
		If iRptLoop = iRptLoopCount Then Execute "iRptLoop" & iRptLoop & " = 0"
		If iRptLoop = iRptLoopCount Then Execute "iRptLoop" & iRptLoop & "Indent = iRptLoop" & iRptLoop-1 & "Indent"
	End If
End Sub


Function SQLis(ByVal sSQL,ByVal sWord)
	If LCase(Left(sSQL,Len(sWord))) = LCase(sWord) then
		addMsg "<center><b>" & sWord & " " & getIDS("IDS_Complete") & "!</b></center>"
		SQLis = True
	Else
		SQLis = False
	End if
End Function


Sub writeNoRecords()
	If iRptLoopCount = 1 Then
		Call writeHeaders()
		If iRptType <> 2 Then Response.Write("</table>" & vbCrLf)
	End If

	If iRptType <> 2 Then
		Call addMsg(getIDS("IDS_MsgNoResults"))
		Call showMsg()
	End If
End Sub


Sub writeHeaders()
	Dim sClass, sName, sTemp

	If iRptType = 3 Then
		sTemp = "<tr>"
	Elseif iRptType < 2 Then
		sTemp = "<table border=""0"" cellspacing=""0"" cellpadding=""2"" width=""100%"">" & vbCrLf
		If iRptLoopCount = 1 Then sTemp = sTemp & "<tr class=""hRow hScr"">" Else sTemp = sTemp & "<tr>"
	End If

	If iRptLoopCount = 1 Then
		For Each i in objRS1.fields
			Select Case Left(i.name,3)
				Case "IDS"
					sName = Application(i.name)
				Case Else
					sName = i.name
			End Select
			If iRptType <> 2 Then
				sTemp =  sTemp & "<th class=""hFont"" nowrap><b>" & sName & "</b></th>"
			Else
				sTemp = sTemp & sName & ","
			End If
		Next
		If iRptType = 2 Then
			sTemp = Left(sTemp,Len(sTemp)-1) & vbCrLf
		Else
			sTemp = sTemp & "</tr>" & vbCrLf
		End if
	End if
	Response.Write(sTemp)
End Sub


Function getCell(sName,sValue,iSize,iType)
	Dim sTemp

	If iRptType < 2 Then
		sTemp = "<td class=""dFont"" nowrap>"
	ElseIf iRptType = 3 Then
		sTemp = "<td nowrap>"
	End If

	If iSize = 0 or sValue = "|" Then
		If iRptType = 2 Then sTemp = sTemp & "null" Else sTemp = sTemp & "<font color=red>null</font>"
	Elseif Trim("" & sValue) = "" Then
		If iRptType = 2 Then sTemp = sTemp & "empty" Else sTemp = sTemp & "<font color=red>empty</font>"
	Else
		If iRptType = 1 Then
			If Instr(sValue,"|") > 0 Then
				Select Case sName
					Case sqlName(getIDS("IDS_Contact")), "IDS_Contact"
						sValue = showLink(1,"../sales/contact.asp?id=" & Mid(sValue,InstrRev(sValue,"|")+1),Left(sValue,InstrRev(sValue,"|")-1))
					Case sqlName(getIDS("IDS_Account")), "IDS_Account"
						sValue = showLink(2,"../sales/client.asp?id=" & Mid(sValue,InstrRev(sValue,"|")+1),Left(sValue,InstrRev(sValue,"|")-1))
					Case sqlName(getIDS("IDS_Project")), "IDS_Project"
						sValue = showLink(4,"../services/project.asp?id=" & Mid(sValue,InstrRev(sValue,"|")+1),Left(sValue,InstrRev(sValue,"|")-1))
					Case sqlName(getIDS("IDS_Event")), "IDS_Event"
						sValue = showLink(0,"../common/event.asp?id=" & Mid(sValue,InstrRev(sValue,"|")+1),Left(sValue,InstrRev(sValue,"|")-1))
				End Select
			Else
				Select Case sName
					Case sqlName(getIDS("IDS_Sale")), "IDS_Sale"
						sValue = showLink(3,"../sales/sale.asp?id=" & sValue,sValue)
					Case sqlName(getIDS("IDS_TicketId")), "IDS_Ticket", "IDS_TicketId"
						sValue = showLink(5,"../support/ticket.asp?id=" & sValue,sValue)
					Case sqlName(getIDS("IDS_BugId")), "IDS_Bug", "IDS_BugId"
						sValue = showLink(6,"../qa/bug.asp?id=" & sValue,sValue)
					Case sqlName(getIDS("IDS_InvoiceId")), "IDS_Invoice", "IDS_InvoiceId"
						sValue = showLink(7,"../finance/invoice.asp?id=" & sValue,sValue)
					Case sqlName(getIDS("IDS_Email")), "IDS_Email"
						sValue = showEmail(sValue)
				End Select
			End If

		Elseif Instr(sValue,"|") > 0 Then
			Select Case sName
				Case sqlName(getIDS("IDS_Contact")),"IDS_Contact",sqlName(getIDS("IDS_Account")),"IDS_Account",sqlName(getIDS("IDS_Sale")),"IDS_Sale",sqlName(getIDS("IDS_Project")),"IDS_Project",sqlName(getIDS("IDS_Event")),"IDS_Event"
					sValue = Left(sValue,InstrRev(sValue,"|")-1)
			End Select
		End If

		'Format Phone Numbers
		If sName = "IDS_Fax" or sName = sqlName(getIDS("IDS_Fax")) or Left(sName,9) = "IDS_Phone" or Left(sName,Len(getIDS("IDS_Phone"))) = sqlName(getIDS("IDS_Phone")) Then
			sValue = showPhone(sValue)
		End If

		'Format Dates
		If Len(sValue) >= 8 and IsDate(sValue) Then
			sValue = showDate(1,sValue)
		End If

		'Format Report Fields
		If iRptType = 2 and Instr(sValue,",") <> 0 Then
			sValue = Chr(34) & sValue & Chr(34)
		Elseif LCase(Left(sValue,4)) = "http" and iRptType > 0 Then
			sValue = "<a href=""" & sValue & """ target=""_top"">" & showString(sValue) & "</a>"
		Elseif NOT (iRptType = 2 or LCase(Left(sValue,7)) = "<a href") Then
			sValue = showString(sValue)
		End If

		sTemp = sTemp & sValue
	End if

	If iRptType = 2 Then
		getCell = sTemp & ","
	Else
		getCell = sTemp & "</td>"
	End If
End Function

Sub writeLine(nIndent,sData)
	Dim sTemp

	If iRptType > 1 Then
		If iRptType = 3 Then
			If iRptLoop-1 <> 0 Then sTemp = sTemp & " style=""mso-outline-level:" & iRptLoop-1 & ";display:none;""" Else sTemp = ""
			sTemp = "<tr" & sTemp & ">"
		End If

		For i = 1 to nIndent
			If iRptType = 3 Then sTemp = sTemp & "<td>&nbsp;</td>" Else sTemp = sTemp & ","
		Next

		If iRptType = 3 Then sTemp = sTemp & sData & "</tr>" Else sTemp = sTemp & Left(sData,Len(sData)-1)
	Else
		If iRptLoopCount = 1 or iRptLoop = iRptLoopCount Then
			sRptColor = toggleRowColor(sRptColor)
			sTemp = "<tr class=""" & sRptColor & """>" & sData & "</tr>"
		Elseif iRptLoopCount > 1 Then
			sTemp = "<tr class=""hRow hFont"">" & sData & "</tr>"
		Else
			sTemp = "<tr>" & sData & "</tr>"
		End If

		If iRptLoop < iRptLoopCount Then sTemp = sTemp & "</table>"
	End If
	Response.Write(sTemp & vbCrLf)
End Sub

Sub addErrMsg(sMsg)
	Call addMsg("<hr color=""red"" width=""70%"" size=3 />" & vbCrLf & sMsg & strSQL)
End Sub

Sub addMsg(sMsg)
	If sMsg = "" Then
		If Err <> 0 Then Call addErrMsg("ASP error: " & Err & " " & Err.Description & "<br />Source: " & Err.Source)
		Err = 0
			If objConn.Errors.Count > 0 then
				For Each i in objConn.Errors
					If i.Number <> 0 Then Call addErrMsg("ERROR from ADO: " & i.Description & " (" & i.Number & ")")
				Next
				objConn.Errors.Clear
			End If
	Else
		sRptMsg = sRptMsg & sMsg & "<br />"
	End if
End Sub

Sub showMsg()
	Call addMsg("")
	If sRptMsg <> "" then
		Response.Write("<center><table border=0 cellspacing=0 cellpadding=10 width=""100%""><tr class=""dRow3""><td class=""wFont"">" & sRptMsg & "</td></tr></table></center><br />")
		sRptMsg = ""
	End if
End Sub
%>