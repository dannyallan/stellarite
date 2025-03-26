<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_calendar.asp" -->
<%
	Call pageFunctions(0,1)

	Dim datDate         'as Date        ' Date we're displaying calendar for
	Dim bytDisplay      'as Integer     ' Choose what to show on the calendar
	Dim blnBill         'as Boolean     ' Choose to show more information
	Dim blnOffsite      'as Boolean     ' Choose to show more information
	Dim blnMine         'as Bookean     ' Only show my events
	Dim intDIM          'as Integer     ' Days In Month
	Dim intDOW          'as Integer     ' Day Of Week that month starts on
	Dim intCurrent      'as Integer     ' Variable we use to hold current day of month as we write table
	Dim intPosition     'as Integer     ' Variable we use to hold current position in table
	Dim intStartYear    'as Integer     ' From RecordSet
	Dim intEndYear      'as Integer     ' From RecordSet
	Dim intStartMonth   'as Integer     ' From RecordSet
	Dim intEndMonth     'as Integer     ' From RecordSet
	Dim intStartDay     'as Integer     ' From RecordSet
	Dim intEndDay       'as Integer     ' From RecordSet
	Dim strTemp         'as String

	If IsDate(Request.QueryString("date")) Then
		datDate = valDate(Request.QueryString("date"),0)
	Elseif valNum(Request.QueryString("selMonth"),1,-1) <> "NULL" Then
		datDate = valDate(DateSerial(Request.QueryString("selYear"),Request.QueryString("selMonth"),"1"),0)
	Else
		datDate = Date
	End If

	strTitle = getIDS("IDS_Calendar") & " - " & MonthName(Month(datDate)) & " " & Year(datDate)
	bytDisplay = Request.QueryString("disp")
	blnBill = Request.QueryString("bill")
	blnOffsite = Request.QueryString("off")
	blnMine = Request.QueryString("my")

	If bytDisplay = "" Then bytDisplay = 1 Else bytDisplay = valNum(bytDisplay,1,0)
	If blnBill = "" Then blnBill = 1 Else blnBill = valNum(blnBill,0,0)
	If blnOffsite = "" Then blnOffsite = 1 Else blnOffsite = valNum(blnOffsite,0,0)
	If blnMine = "" Then blnMine = 1 Else blnMine = valNum(blnMine,0,0)

	intDIM = getDaysInMonth(Month(datDate), Year(datDate))
	intDOW = getWeekdayMonthStartsOn(Month(datDate), Year(datDate))

	Function getDaysInMonth(intMonth, intYear)

		Select Case intMonth
			Case 01, 03, 05, 07, 08, 10, 12
				getDaysInMonth = 31
			Case 04, 06, 09, 11
				getDaysInMonth = 30
			Case 02
				If IsDate("February 29, " & intYear) Then
					getDaysInMonth = 29
				Else
					getDaysInMonth = 28
				End If
		End Select

	End Function


	Function getWeekdayMonthStartsOn(intMonth, intYear)

		getWeekdayMonthStartsOn = WeekDay(DateSerial(intYear,intMonth,"1"))

	End Function


	Function subtractOneMonth(datDate)

		Dim intDay, intMonth, intYear

		intDay = Day(datDate)
		intMonth = Month(datDate)
		intYear = Year(datDate)

		If intMonth = 01 Then
			intMonth = 12
			intYear = intYear - 1
		Else
			intMonth = intMonth - 1
		End If

		If intDay > getDaysInMonth(intMonth, intYear) Then
			intDay = getDaysInMonth(intMonth, intYear)
		End If

		subtractOneMonth = DateSerial(intYear,intMonth,intDay)

	End Function


	Function addOneMonth(datDate)

		Dim intDay, intMonth, intYear

		intDay = Day(datDate)
		intMonth = Month(datDate)
		intYear = Year(datDate)

		If intMonth = 12 Then
			intMonth = 01
			intYear = intYear + 1
		Else
			intMonth = intMonth + 1
		End If

		If intDay > getDaysInMonth(intMonth, intYear) Then
			intDay = getDaysInMonth(intMonth, intYear)
		End if

		addOneMonth = DateSerial(intYear,intMonth,intDay)

	End Function

	Sub getMonths(intMonth)
		For i = 1 to 12
			Response.Write(vbTab & vbTab & "<option value=""" & i & """" & getDefault(0,intMonth,i) & ">" & MonthName(i) & "</option>" & vbCrLf)
		Next
	End Sub

	Sub getYears(intYear)
		For i = Year(CDate(showDate(0,Now)))-2 to Year(CDate(showDate(0,Now)))+2
			Response.Write(vbTab & vbTab & "<option value=""" & i & """" & getDefault(0,intYear,i) & ">" & i & "</option>" & vbCrLf)
		Next
	End Sub

	Set objRS = objConn.Execute(getCalendar(bytMod,blnBill,blnOffsite,blnMine,subtractOneMonth(datDate),addOneMonth(datDate),lngUserId))
	If not (objRS.BOF and objRS.EOF) Then arrRS = objRS.GetRows()

	If strModName = "" Then strModName = strTitle

	Call DisplayHeader(1)
%>

<div id="contentDiv" class="dvNoBorder">

<table border=0 cellspacing=0 cellpadding=10 width="100%">
  <tr>
	<td class="dFont" valign=top width="20%">
	  <form name="frmCalendar" method="get" action="calendar.asp">
	  <% =getHidden("m",bytMod) %>
	  <% =getHidden("disp",bytDisplay) %>
	  <% =getHidden("bill",blnBill) %>
	  <% =getHidden("off",blnOffsite) %>
	  <% =getHidden("my",blnMine) %>
	  <table border=0 cellpadding=10 width="100%">
		<tr class="hRow">
		  <td class="bFont" nowrap>
			<br />
<%
Response.Write(getRadio("rdoDisplay",1,bytDisplay,"onClick=""window.location.href='calendar.asp?date=" & datDate & "&m=" & bytMod & "&disp=1&bill=" & blnBill & "&off=" & blnOffsite & "&my=" & blnMine & "';""") & getLabel(getIDS("IDS_DisplayEventTitles"),"rdoDisplay1") & "<br />" & vbCrLf & _
		getRadio("rdoDisplay",2,bytDisplay,"onClick=""window.location.href='calendar.asp?date=" & datDate & "&m=" & bytMod & "&disp=2&bill=" & blnBill & "&off=" & blnOffsite & "&my=" & blnMine & "';""") & getLabel(getIDS("IDS_DisplayOwners"),"rdoDisplay2") & "<br />" & vbCrLf & _
		getRadio("rdoDisplay",3,bytDisplay,"onClick=""window.location.href='calendar.asp?date=" & datDate & "&m=" & bytMod & "&disp=3&bill=" & blnBill & "&off=" & blnOffsite & "&my=" & blnMine & "';""") & getLabel(getIDS("IDS_DisplayParent"),"rdoDisplay3") & "<br />" & vbCrLf)
%>
			<br />
		  </td>
		</tr>
	  </table>
	  <br />
	  <table border=0 cellpadding=10 width="100%">
		<tr class="hRow">
		  <td class="bFont" nowrap>
			<br />
			<% =getCheckbox("chkMine",blnMine,"onClick=""window.location.href='calendar.asp?date=" & datDate & "&m=" & bytMod & "&disp=" & bytDisplay & "&bill=" & blnBill & "&off=" & blnOffsite & "&my=" & Abs(blnMine-1) & "';""") %><% =getLabel(getIDS("IDS_DisplayMine"),"chkMine") %><br />
			<% =getCheckbox("chkBillable",blnBill,"onClick=""window.location.href='calendar.asp?date=" & datDate & "&m=" & bytMod & "&disp=" & bytDisplay & "&bill=" & Abs(blnBill-1) & "&off=" & blnOffsite & "&my=" & blnMine & "';""") %><% =getLabel(getIDS("IDS_DisplayBillable"),"chkBillable") %><br />
			<% =getCheckbox("chkOffsite",blnOffsite,"onClick=""window.location.href='calendar.asp?date=" & datDate & "&m=" & bytMod & "&disp=" & bytDisplay & "&bill=" & blnBill & "&off=" & Abs(blnOffsite-1) & "&my=" & blnMine & "';""") %><% =getLabel(getIDS("IDS_DisplayOffsite"),"chkOffsite") %><br />
			<br />
		  </td>
		</tr>
	  </table>
	  <br />
	  <% =getModuleDropDown("selCalendar",bytMod,True,"onChange=""window.location.href='calendar.asp?date=" & datDate & "&m='+document.forms[0].selCalendar.value+'&disp=" & bytDisplay & "&bill=" & blnBill & "&off=" & blnOffsite & "&my=" & blnMine & "';""") %><br /><br />
	  <table border=0 cellpadding=0 width="100%">
		<tr>
		  <td nowrap>
			<select name="selMonth" id="selMonth" class="oByte">
			<% Call getMonths(Month(datDate)) %>
			</select>
			<select name="selYear" id="selYear" class="oInt">
			<% Call getYears(Year(datDate)) %>
			</select>
			<% =getSubmit("btnGo",getIDS("IDS_Go"),30,"G","") %>
		  </td>
		</tr>
	  </table>
	  </form>
	</td>
	<td width="80%">
	  <table border="1" cellspacing="0" cellpadding="2" width="100%">
	  <tr>
		<td bgcolor="black" align="center" colspan="7">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="dFont" align="right" width="20%">
						<% If (Year(Date) - Year(datDate)) < 2 then %><b><a href="calendar.asp?date=<%= subtractOneMonth(datDate) %>&m=<% =bytMod %>&disp=<% =bytDisplay %>&bill=<% =blnBill %>&off=<% =blnOffsite %>&my=<% =blnMine %>" style="color: #FFFFFF">
											&lt;--
</a></b>
						<% End If %>
					</td>
					<td class="dFont" align="center" width="60%"><br /><font color="#FFFFFF"><b>
						<%= MonthName(Month(datDate)) & "  " & Year(datDate) %></b></font><br /><br />
					</td>
					<td class="dFont" width="20%">
						<% If (Year(datDate) - Year(Date)) < 2 then %><b><a href="calendar.asp?date=<%= addOneMonth(datDate) %>&m=<% =bytMod %>&disp=<% =bytDisplay %>&bill=<% =blnBill %>&off=<% =blnOffsite %>&my=<% =blnMine %>" style="color: #FFFFFF">
											--&gt;
</a></b>
						<% End If %>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr class="hRow">
		<th class="hFont" width="15%"><center><% =Left(WeekDayName(1,True),3) %></center></th>
		<th class="hFont" width="14%"><center><% =Left(WeekDayName(2,True),3) %></center></th>
		<th class="hFont" width="14%"><center><% =Left(WeekDayName(3,True),3) %></center></th>
		<th class="hFont" width="14%"><center><% =Left(WeekDayName(4,True),3) %></center></th>
		<th class="hFont" width="14%"><center><% =Left(WeekDayName(5,True),3) %></center></th>
		<th class="hFont" width="14%"><center><% =Left(WeekDayName(6,True),3) %></center></th>
		<th class="hFont" width="15%"><center><% =Left(WeekDayName(7,True),3) %></center></th>
	</tr>
	<%
	If intDOW <> 1 Then
		Response.Write(vbTab & "<tr>" & vbCrLf)
		intPosition = 1

		Do While intPosition < intDOW
			Response.Write(vbTab & vbTab & "<td class=""dFont"">&nbsp;</td>" & vbCrLf)
			intPosition = intPosition + 1
		Loop
	End If

	intCurrent = 1
	intPosition = intDOW


	Do While intCurrent <= intDIM

		If intPosition = 1 Then
			Response.Write(vbTab & "<tr class=""dFont"">" & vbCrLf)
		End If

		Response.Write(vbTab & vbTab & "<td class=""dFont"" align=left valign=top height=60><b>" & intCurrent & "</b>")

		If isArray(arrRS) Then

			For i = 0 to UBound(arrRS,2)

				intStartYear = Year(showDate(1,arrRS(2,i)))
				intEndYear = Year(showDate(1,arrRS(3,i)))
				intStartMonth = Month(showDate(1,arrRS(2,i)))
				intEndMonth = Month(showDate(1,arrRS(3,i)))
				intStartDay = Day(showDate(1,arrRS(2,i)))
				intEndDay = Day(showDate(1,arrRS(3,i)))

				If intStartYear <= Year(datDate) and Year(datDate) <= intEndYear Then
					If intStartMonth <= Month(datDate) and Month(datDate) <= intEndMonth Then

					'-- Displays the event.  Also allows for events which cross from one month into the next.  This will
					'-- not support events which cross from one year to the next, but this should not take place because
					'-- events should not cross through the Christmas holidays.

					If (intStartMonth = intEndMonth and intStartDay <= intCurrent and intCurrent <= intEndDay) or _
					   (intStartMonth <> intEndMonth and intStartMonth = Month(datDate) and intCurrent >= intStartDay) or _
					   (intStartMonth <> intEndMonth and intEndMonth = Month(datDate) and intCurrent <= intEndDay) or _
					   (intStartMonth <> intEndMonth and intStartMonth <> Month(datDate) and intEndMonth <> Month(datDate)) Then

						If bytDisplay = 1 Then
							strTemp = trimString(arrRS(1,i),15)
						Elseif bytDisplay = 2 Then
							strTemp = trimString(arrRS(4,i),15)
						Else
							Select Case bytMod
								Case 3,5,6,7
									strTemp = bigDigitNum(7,arrRS(7,i))
								Case Else
									strTemp = trimString(arrRS(7,i),15)
							End Select
						End If



						If strTemp = "" or IsNull(strTemp) Then strTemp = getIDS("IDS_Unspecified")

						'If intPosition <> 1 and intPosition <> 7 Then
							Response.Write("<br /><a href=""../common/event.asp?id=" & arrRS(0,i) & "&m=" & arrRS(5,i) & "&mid=" & arrRS(6,i) & """>" & strTemp & "</a>")
						'End If
					End If
					End If
				End If
			Next
		End If

		Response.Write("</td>" & vbCrLf)

		If intPosition = 7 Then
			Response.Write(vbTab & "</tr>" & vbCrLf)
			intPosition = 0
		End If

		intCurrent = intCurrent + 1
		intPosition = intPosition + 1
	Loop


	If intPosition <> 1 Then

		Do While intPosition <= 7
			Response.Write(vbTab & vbTab & "<td class=""dFont"">&nbsp;</td>" & vbCrLf)
			intPosition = intPosition + 1
		Loop

		Response.Write(vbTab & "</tr>" & vbCrLf)
	End If

	Response.Write("</table></td></tr></table></div>" & vbCrLf)

	Call DisplayFooter(1)
%>