<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_events.asp" -->
<%
	Call pageFunctions(50,2)

	Dim strEventType	'as String
	Dim intEventType	'as Integer
	Dim intPermissions	'as Integer
	Dim blnOnsite		'as Boolean
	Dim blnBillable		'as Boolean
	Dim strEventTitle	'as String
	Dim strOwner		'as String
	Dim datDate			'as Date
	Dim intSMonth		'as Integer
	Dim intSDay			'as Integer
	Dim intSYear		'as Integer
	Dim intSHour		'as Integer
	Dim intSMin			'as Integer
	Dim intEMonth		'as Integer
	Dim intEDay			'as Integer
	Dim intEYear		'as Integer
	Dim intEHour		'as Integer
	Dim intEMin			'as Integer
	Dim datStartDate	'as Date
	Dim datStartTime	'as Date
	Dim datEndDate		'as Date
	Dim datEndTime		'as Date
	Dim strCreatedBy	'as String
	Dim strModBy		'as String
	Dim datCreatedDate	'as Date
	Dim datModDate		'as Date

	strModName = Application("IDS_Event")

	Sub getMonths(iMonth)
		For i = 1 to 12
			Response.Write(vbTab & vbTab & "<option value=""" & i & """" & getDefault(0,iMonth,i) & ">" & MonthName(i) & "</option>" & vbCrLf)
		Next
	End Sub

	Sub getYears(iYear)
		For i = Year(CDate(showDate(0,Now)))-2 to Year(CDate(showDate(0,Now)))+2
			Response.Write(vbTab & vbTab & "<option value=""" & i & """" & getDefault(0,iYear,i) & ">" & i & "</option>" & vbCrLf)
		Next
	End Sub

	Sub getDays(iDay)
		For i = 1 to 31
			Response.Write(vbTab & vbTab & "<option value=""" & i & """" & getDefault(0,iDay,i) & ">" & i & "</option>" & vbCrLf)
		Next
	End Sub

	Sub getHours(iHour)
		For i = 0 to 23
			Response.Write(vbTab & vbTab & "<option value=""" & i & """" & getDefault(0,iHour,i) & ">" & i & "</option>" & vbCrLf)
		Next
	End Sub

	Sub getMins(iMin)
%>
		<option value="00"<% =getDefault(0,iMin,00) %>>00</option>
		<option value="15"<% =getDefault(0,iMin,15) %>>15</option>
		<option value="30"<% =getDefault(0,iMin,30) %>>30</option>
		<option value="45"<% =getDefault(0,iMin,45) %>>45</option>
<%
	End Sub

	strTitle = Application("IDS_Event")

	If strDoAction <> "" then

		bytMod = valNum(bytMod,1,0)
		lngModId = valNum(lngModId,3,0)
		blnOnsite = valNum(Request.Form("chkOnsite"),0,0)
		blnBillable = valNum(Request.Form("chkBillable"),0,0)
		strOwner = getUserId(0,valString(Request.Form("txtOwner"),255,1,0))
		intEventType = valNum(Request.Form("selEventType"),2,-1)
		intPermissions = valNum(Request.Form("selPermissions"),1,1)
		strEventTitle = valString(Request.Form("txtTitle"),40,1,0)

		datStartTime = TimeSerial(Request.Form("selSHour"),Request.Form("selSMin"),"00")
		datEndTime = TimeSerial(Request.Form("selEHour"),Request.Form("selEMin"),"00")

		If IsDate(datStartTime) and IsDate(datEndTime) Then
			datStartTime = DateAdd("n",0-Session("TimeOffset"),datStartTime)
			datEndTime = DateAdd("n",0-Session("TimeOffset"),datEndTime)

			datStartTime = valDate(DateSerial(Request.Form("selSYear"),Request.Form("selSMonth"),Request.Form("selSDay")) & " " & datStartTime,1)
			datEndTime = valDate(DateSerial(Request.Form("selEYear"),Request.Form("selEMonth"),Request.Form("selEDay")) & " " & datEndTime,1)
		End If

		If strDoAction = "del" and intPerm >= 4 Then

			Call delEvent(lngUserId,lngRecordId,bytMod,lngModId)

		Elseif strDoAction = "edit" and intPerm >= 3 then

			Call updateEvent(lngUserId,lngRecordId,bytMod,lngModId,strOwner,blnOnsite,blnBillable,intEventType,intPermissions,strEventTitle,datStartTime,datEndTime)

		ElseIf strDoAction = "new" Then

			lngRecordId = insertEvent(lngUserId,bytMod,lngModId,strOwner,blnOnsite,blnBillable,intEventType,intPermissions,strEventTitle,datStartTime,datEndTime)
		End If

		Call closeWindow(strOpenerURL)

	Else

		If blnRS and intPerm >= 3 Then

			Set objRS = objConn.Execute(getEvent(0,lngRecordId,bytMod,lngModId))

			If not (objRS.BOF and objRS.EOF) then
				intEventType = objRS.fields("E_EventType").value
				intPermissions = objRS.fields("E_Permissions").value
				strEventTitle = showString(objRS.fields("E_Title").value)
				strOwner = showString(objRS.fields("Owner").value)
				intSMonth = Month(objRS.fields("E_StartTime").value)
				intSDay = Day(objRS.fields("E_StartTime").value)
				intSYear = Year(objRS.fields("E_StartTime").value)
				intSHour = Hour(showDate(1,objRS.fields("E_StartTime").value))
				intSMin = Minute(showDate(1,objRS.fields("E_StartTime").value))
				intEMonth = Month(objRS.fields("E_EndTime").value)
				intEDay = Day(objRS.fields("E_EndTime").value)
				intEYear = Year(objRS.fields("E_EndTime").value)
				intEHour = Hour(showDate(1,objRS.fields("E_EndTime").value))
				intEMin = Minute(showDate(1,objRS.fields("E_EndTime").value))
				blnOnsite = valNum(objRS.fields("E_Onsite").value,0,0)
				blnBillable = valNum(objRS.fields("E_Billable").value,0,0)
				strCreatedBy = showString(objRS.fields("CreatedBy").value)
				datCreatedDate = objRS.fields("E_CreatedDate").value
				strModBy = showString(objRS.fields("ModBy").value)
				datModDate = objRS.fields("E_ModDate").value
			End If

			If intPermissions = 1 and intMember > 1 Then Call logError(2,1)

		Elseif blnRS Then
			Call doRedirect("pop_event.asp?m=" & bytMod & "&mid=" & lngModId)
		Else
			datDate = valDate(Request.QueryString("date"),0)
			If datDate = "" Then
				intSMonth = Month(Now)
				intSDay = Day(Now)
				intSYear = Year(Now)
			Else
				intSMonth = Month(CDate(datDate))
				intSDay = Day(CDate(datDate))
				intSYear = Year(CDate(datDate))
			End If
			intEMonth = intSMonth
			intEDay = intSDay
			intEYear = intSYear
			intSHour = 8
			intEHour = 17
			strOwner = strFullName
			strCreatedBy = strFullName
			datCreatedDate = Now
			strModBy = strFullName
			datModDate = Now
		End If
		strEventType = getOptionDropDown(150,False,"selEventType","Event Type",intEventType)
	End If

	Call DisplayHeader(3)
	Call showEditHeader(strTitle,strCreatedBy,datCreatedDate,strModBy,datModDate)
%>
<div id="contentDiv" class="dvBorder" style="height:335px;"><br>

<table border=0 cellspacing=5 width="100%">
<form name="frmEvent" method="post" action="pop_event.asp?id=<% =lngRecordId %>&m=<% =bytMod %>&mid=<% =lngModId %>">
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
<% =getHidden("hdnWinOpen","") %>
  <tr>
    <td><% =getLabel(Application("IDS_Title"),"txtTitle") %></td>
    <td colspan=3><% =getTextField("txtTitle","mText",strEventTitle,40,40,"") %></td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_Owner"),"txtOwner") %></td>
    <td colspan=3><% =getTextField("txtOwner","mText",strOwner,40,255,"") %>
    <a href="<% =newWindow("S","?m=0&rVal=txtOwner") %>"><img src="../images/import.gif" alt="<% =getImport("IDS_Owner") %>" border=0 height=16 width=16></a></td>
  </tr>
  <tr><td class="dfont" colspan=4>&nbsp;</td></tr>
  <tr>
    <td><% =getLabel(Application("IDS_Start"),"selSMonth") %></td>
    <td colspan=3>
	<select name="selSMonth" id="selSMonth" class="oByte" onChange="doChange();">
	<% Call getMonths(intSMonth) %>
	</select>
	<select name="selSDay" id="selSDay" class="oByte" onChange="doChange();">
	<% Call getDays(intSDay) %>
	</select>
	<select name="selSYear" id="selSYear" class="oInt" onChange="doChange();">
	<% Call getYears(intSYear) %>
	</select>&nbsp;&nbsp;
	<select name="selSHour" id="selSHour" class="oByte" onChange="doChange();">
	<% Call getHours(intSHour) %>
	</select><b>:</b>
	<select name="selSMin" id="selSMin" class="oByte" onChange="doChange();">
	<% Call getMins(intSMin) %>
	</select>
    </td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_End"),"selEMonth") %></td>
    <td colspan=3>
	<select name="selEMonth" id="selEMonth" class="oByte" onChange="doChange();">
	<% Call getMonths(intEMonth) %>
	</select>
	<select name="selEDay" id="selEDay" class="oByte" onChange="doChange();">
	<% Call getDays(intEDay) %>
	</select>
	<select name="selEYear" id="selEYear" class="oInt" onChange="doChange();">
	<% Call getYears(intEYear) %>
	</select>&nbsp;&nbsp;
	<select name="selEHour" id="selEHour" class="oByte" onChange="doChange();">
	<% Call getHours(intEHour) %>
	</select><b>:</b>
	<select name="selEMin" id="selEMin" class="oByte" onChange="doChange();">
	<% Call getMins(intEMin) %>
	</select>
    </td>
  </tr>
  <tr><td class="dfont" colspan=3>&nbsp;</td></tr>
  <tr>
    <td><% =getLabel(Application("IDS_Type"),"selEventType") %></td>
    <td colspan=3><% =strEventType %></td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_Permissions"),"selPermissions") %></td>
    <td colspan=3><% =getPermissionsDropDown(intPermissions,intMember) %></td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_Billable"),"chkBillable") %></td>
    <td colspan=3><% =getCheckbox("chkBillable",blnBillable,"") %></td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_Onsite"),"chkOnsite") %></td>
    <td colspan=3><% =getCheckbox("chkOnsite",blnOnsite,"") %></td>
 </tr>
</form>
</table>

</div>

<div id="footerDiv" class="dvFooter">
<%
	If blnRS Then
		Response.Write(getIconNew("pop_event.asp?m=" & bytMod & "&mid=" & lngModId))
	End If
	If blnRS and intPerm >= 4 Then
		Response.Write(getIconDelete())
	End If
	Response.Write(getIconSave(strAction))
	Response.Write(getIconCancel())
%>
</div>

<%
	Call DisplayFooter(3)
%>

