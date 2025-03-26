<!--#include file="sql\sql_lists.asp" -->
<%
Function getListName(fConst)

	Select Case fConst
		Case 0
			getListName = getIDS("IDS_My") & " " & getIDS("IDS_RecentlyUpdatedContacts")
		Case 1
			getListName = getIDS("IDS_My") & " " & getIDS("IDS_RecentlyUpdatedClients")
		Case 2
			getListName = getIDS("IDS_My") & " " & getIDS("IDS_RecentlyUpdatedSales")
		Case 3
			getListName = getIDS("IDS_My") & " " & getIDS("IDS_ProjectOpen")
		Case 4
			getListName = getIDS("IDS_My") & " " & getIDS("IDS_UpcomingEvents")
		Case 5
			getListName = getIDS("IDS_My") & " " & getIDS("IDS_TicketsOpen")
		Case 6
			getListName = getIDS("IDS_My") & " " & getIDS("IDS_BugsOpen")
		Case 7
			getListName = getIDS("IDS_My") & " " & getIDS("IDS_InvoicesOpen")
		Case 8
			getListName = getIDS("IDS_My") & " " & getIDS("IDS_Reports")
		Case 9
			getListName = getIDS("IDS_RecentlyUpdatedContacts")
		Case 10
			getListName = getIDS("IDS_RecentlyUpdatedClients")
		Case 11
			getListName = getIDS("IDS_RecentlyUpdatedSales")
		Case 12
			getListName = getIDS("IDS_ProjectOpen")
		Case 13
			getListName = getIDS("IDS_UpcomingClientEvents")
		Case 14
			getListName = getIDS("IDS_HotTickets")
		Case 15
			getListName = getIDS("IDS_TicketsOpen")
		Case 16
			getListName = getIDS("IDS_HotBugs")
		Case 17
			getListName = getIDS("IDS_BugsOpen")
		Case 18
			getListName = getIDS("IDS_InvoicesOverdue")
		Case 19
			getListName = getIDS("IDS_InvoicesOpen")
		Case 20
			getListName = getIDS("IDS_ReportsPublic")
		Case 21
			getListName = getIDS("IDS_ArticlesNew")
		Case 30
			getListName = getIDS("IDS_Articles")
		Case 31
			getListName = getIDS("IDS_Reports")
	End Select

End Function

Sub showList(fConst,fSize,fCount,fShowCols)

	Dim fColor,fSql,fRow,fTemp
	Dim fWidth1,fWidth2,fWidth3
	Dim fCol1,fCol2,fCol3
	Dim fVal1,fVal2,fVal3

	fCol1 = getListName(fConst)
	fCol2 = getIDS("IDS_Modified")
	fCol3 = ""

	If fSize < 100 Then fSize = 100

	Select Case fConst
		Case 0,9
			fSql = getContacts(fConst,lngUserId,fCount)
			fCol3 = getIDS("IDS_ModifiedBy")
		Case 1,10
			fSql = getClients(fConst,lngUserId,fCount)
			fCol3 = getIDS("IDS_ModifiedBy")
		Case 2,11
			fSql = getSales(fConst,lngUserId,fCount)
			fCol3 = getIDS("IDS_ModifiedBy")
		Case 3,12
			fSql = getProjects(fConst,lngUserId)
			fCol3 = getIDS("IDS_Owner")
		Case 4,13
			fSql = getEvents(fConst,lngUserId,fCount)
			fCol2 = getIDS("IDS_Owner")
			fCol3 = getIDS("IDS_StartTime")
		Case 5,14,15
			fSql = getTickets(fConst,lngUserId)
			fCol2 = getIDS("IDS_Account")
			fCol3 = getIDS("IDS_Owner")
		Case 6,16,17
			fSql = getBugs(fConst,lngUserId)
			fCol3 = getIDS("IDS_Owner")
		Case 7,18,19
			fSql = getInvoices(fConst,lngUserId)
			fCol2 = getIDS("IDS_Account")
			fCol3 = getIDS("IDS_Modified")
		Case 8,20
			fSql = getReports(fConst,lngUserId,"0,1,2,3,4,5,6,7,50,90",Session("Member"),Session("Permissions"))
			fCol2 = getIDS("IDS_Owner")
		Case 21
			fSQL = getArticles(lngUserId,fCount,"")
		Case 30
			fSQL = getArticles(lngUserId,"",fCount)
		Case 31
			fSQL = getReports(0,lngUserId,fCount,Session("Member"),Session("Permissions"))
	End Select

	Select Case fConst
		Case 4,9,10,11,12,16,17
			fTemp = fCol3
			fCol3 = fCol2
			fCol2 = fTemp
	End Select

	If fShowCols < 3 Then fCol3 = ""
	If fShowCols < 2 Then fCol2 = ""


	Select Case fShowCols
		Case 3
			fWidth1 = "33"
			fWidth2 = "33"
			fWidth3 = "33"
		Case 2
			fWidth1 = "60"
			fWidth2 = "40"
		Case Else
			fWidth1 = "100"
	End Select

	Response.Write("<div class=""dvNoBorder"" style=""height: " & fSize & "px;""><div class=""dvNoBorder"" style=""height: " & fSize-10 & "px;"">" & vbCrLf & _
			"  <table border=""0"" cellpadding=""2"" cellspacing=""0"" summary=""" & fCol1 & """ width=""100%"">" & vbCrLf & _
			"    <thead><tr class=""hRow hScr"">" & vbCrLf & _
			"      <th class=""hFont"" width=""" & fWidth1 & "%"">&nbsp;" & fCol1 & "</th>" & vbCrLf)

	If fCol2 <> "" Then Response.Write("        <th class=""hFont"" width=""" & fWidth2 & "%"">" & fCol2 & "</th>" & vbCrLf)
	If fCol3 <> "" Then Response.Write("        <th class=""hFont"" width=""" & fWidth3 & "%"">" & fCol3 & "</th>" & vbCrLf)

	Response.Write("    </tr></thead>" & vbCrLf & "    <tbody>" & vbCrLf)

	Set objFRS = objConn.Execute(fSql)
	If (objFRS.BOF and objFRS.EOF) Then
		Response.Write("<tr><td class=""dFont"" colspan=" & fShowCols & ">" & getIDS("IDS_NoneSpecified") & "<br /><br /></td></tr>" & vbCrLf)
	Else
		arrRS = objFRS.GetRows()

		For fRow = 0 to UBound(arrRS,2)
			fColor = toggleRowColor(fColor)

			Select Case fConst
				Case 0,9
					fVal1 = "<a href=""" & Application("av_CRMDir") & "sales/contact.asp?id=" & arrRS(0,fRow) & """>" & trimString(arrRS(1,fRow),25) & "</a>"
					fVal2 = showDate(0,arrRS(2,fRow))
					fVal3 = trimString(arrRS(3,fRow),25)
				Case 1,10
					fVal1 = "<a href=""" & Application("av_CRMDir") & "sales/client.asp?id=" & arrRS(0,fRow) & """>" & trimString(arrRS(1,fRow),25) & "</a>"
					fVal2 = showDate(0,arrRS(2,fRow))
					fVal3 = trimString(arrRS(3,fRow),25)
				Case 2,11
					fVal1 = "<a href=""" & Application("av_CRMDir") & "sales/sale.asp?id=" & arrRS(0,fRow) & """>" & trimString(arrRS(1,fRow),25) & "</a>"
					fVal2 = showDate(0,arrRS(2,fRow))
					fVal3 = trimString(arrRS(3,fRow),25)
				Case 3,12
					fVal1 = "<a href=""" & Application("av_CRMDir") & "services/project.asp?id=" & arrRS(0,fRow) & """>" & trimString(arrRS(1,fRow),25) & "</a>"
					fVal2 = showDate(0,arrRS(2,fRow))
					fVal3 = trimString(arrRS(3,fRow),25)
				Case 4,13
					fVal1 = "<a href=""" & Application("av_CRMDir") & "common/event.asp?id=" & arrRS(0,fRow) & "&m=" &arrRS(1,fRow) & "&mid=" & arrRS(2,fRow) & """>" & trimString(arrRS(3,fRow),25) & "</a>"
					fVal2 = trimString(arrRS(5,fRow),25)
					fVal3 = showDate(1,arrRS(4,fRow))
				Case 5,14,15
					fVal1 = "<a href=""" & Application("av_CRMDir") & "support/ticket.asp?id=" & arrRS(0,fRow) & """>" & bigDigitNum(7,arrRS(0,fRow)) & "</a>"
					fVal2 = showLink(2,Application("av_CRMDir") & "sales/client.asp?id=" & arrRS(1,fRow),trimString(arrRS(2,fRow),25))
					fVal3 = trimString(arrRS(4,fRow),25)
					If fConst = 15 and arrRS(3,fRow) = 1 Then fColor = "dRow3"
				Case 6,16,17
					fVal1 = "<a href=""" & Application("av_CRMDir") & "qa/bug.asp?id=" & arrRS(0,fRow) & """>" & bigDigitNum(7,arrRS(0,fRow)) & "</a>"
					fVal2 = showDate(1,arrRS(1,fRow))
					fVal3 = trimString(arrRS(3,fRow),25)
					If fConst = 17 and arrRS(2,fRow) = 1 Then fColor = "dRow3"
				Case 7,18,19
					fVal1 = "<a href=""" & Application("av_CRMDir") & "finance/invoice.asp?id=" & arrRS(0,fRow) & """>" & bigDigitNum(7,arrRS(0,fRow)) & "</a>"
					fVal2 = showLink(2,Application("av_CRMDir") & "sales/client.asp?id=" & arrRS(1,fRow),trimString(arrRS(2,fRow),25))
					fVal3 = showDate(1,arrRS(3,fRow))
					If fConst = 19 and DateDiff("d",Date,arrRS(4,fRow)) < 0 Then fColor = "dRow3"
				Case 8,20,31
					fVal1 = "<a href=""" & Application("av_CRMDir") & "reports/view_report.asp?id=" & arrRS(0,fRow) & "&m=" & arrRS(1,fRow) & "&type=" & arrRS(2,fRow) & """>" & trimString(arrRS(3,fRow),25) & "</a>"
					fVal2 = trimString(arrRS(4,fRow),25)
				Case 21,30
					fVal1 = "<a href=""" & Application("av_CRMDir") & "kb/article.asp?id=" & arrRS(0,fRow) & """>" & trimString(arrRS(1,fRow),30) & "</a>"
					fVal2 = showDate(0,arrRS(2,fRow))
			End Select

			Select Case fConst
				Case 4,9,10,11,12,16,17
					fTemp = fVal3
					fVal3 = fVal2
					fVal2 = fTemp
			End Select

			Response.Write("<tr class=""" & fColor & """><td class=""dFont"" width=""" & fWidth1 & "%"">" & fVal1 & "</td>")

			If fCol2 <> "" Then Response.Write("<td class=""dFont"" width=""" & fWidth2 & "%"">" & fVal2 & "</td>")
			If fCol3 <> "" Then Response.Write("<td class=""dFont"" width=""" & fWidth3 & "%"">" & fVal3 & "</td>")

			Response.Write("</tr>" & vbCrLf)
		Next
	End If
	Response.Write("    </tbody>" & vbCrLf & "  </table></div></div>" & vbCrLf)
End Sub

%>