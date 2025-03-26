<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_reports.asp" -->
<!--#include file="..\_inc\currency.asp" -->
<!--#include file="..\_inc\states.asp" -->
<%
	Call pageFunctions(0,1)

	Dim strGenSQL		'as String
	Dim bytType			'as Byte
	Dim arrTemp			'as Array
	Dim arrParam(25,2)	'as Array
	Dim arrField(30,1)	'as Array
	Dim bytOrder		'as Byte
	Dim strOrderDir		'as String
	Dim intParamLimit	'as Integer
	Dim intFieldLimit	'as Integer
	Dim datCreatedBefore'as Date
	Dim datCreatedAfter	'as Date
	Dim strCreatedUser	'as String
	Dim lngCreatedId	'as Long
	Dim datModBefore	'as Date
	Dim datModAfter		'as Date
	Dim strModUser		'as String
	Dim strReportName	'as String
	Dim intPermissions	'as Integer
	Dim strOwner		'as String
	Dim strCreatedBy	'as String
	Dim datCreatedDate	'as Date
	Dim strModBy		'as String
	Dim datModDate		'as Date
	Dim blnChange		'as Boolean

	strTitle = Application("IDS_ReportNew")
	strDir = Application("av_CRMDir") & "reports/"
	strCatName = "Reports"

	blnChange = valNum(Request.Form("hdnChange"),0,0)
	bytType = valNum(Request.QueryString("type"),1,0)
	If bytType = "" Then bytType = bytMod

	Function getTrueFalse(fName,fDefault)
		getTrueFalse = "<select name=""" & fName & """ id=""" & fName & """ class=""oBool"">" & vbCrLf & _
			"<option></option>" & vbCrLf & _
			"<option value=""1""" & getDefault(0,fDefault,"1") & ">" & Application("IDS_True") & "</option>" & vbCrLf & _
			"<option value=""0""" & getDefault(0,fDefault,"0") & ">" & Application("IDS_False") & "</option></select>" & vbCrLf
	End Function

	If blnChange = 1 Then

		strGenSQL = Replace(Request.Form("txtSQL"),"Chr(34)",Chr(34))

		arrTemp = Split(Request.Form("txtParams"),"|")
		For i = 1 to UBound(arrTemp)
			arrParam(i,2) = arrTemp(i)
		Next

		arrTemp = Split(Request.Form("txtFields"),"|")
		For i = 1 to UBound(arrTemp)
			arrField(i,1) = arrTemp(i)
		Next

		datCreatedAfter = showDate(0,valDate(Request.Form("txtCreatedAfter"),0))
		datCreatedBefore = showDate(0,valDate(Request.Form("txtCreatedBefore"),0))
		strCreatedUser = valString(Request.Form("txtCreatedUser"),150,0,0)
		datModAfter = showDate(0,valDate(Request.Form("txtModAfter"),0))
		datModBefore = showDate(0,valDate(Request.Form("txtModBefore"),0))
		strModUser = valString(Request.Form("txtModUser"),150,0,0)
		bytOrder = valNum(Request.Form("selOrder"),1,0)
		strOrderDir = valString(Request.Form("selOrderDir"),4,0,0)
		strReportName = valString(Request.Form("txtReportName"),100,0,0)
		intPermissions = valNum(Request.Form("selPermissions"),1,1)
		strOwner = valString(Request.Form("txtOwner"),100,0,0)

	Elseif blnRS Then

		Set objRS = objConn.Execute(getReport(lngRecordId))

		If not (objRS.BOF and objRS.EOF) then
			strGenSQL = objRS.fields("R_SQL").value
			bytMod = objRS.fields("R_Module").value
			bytType = objRS.fields("R_Type").value

			If bytType = 90 and not blnAdmin Then

				Call logError(2,1)

			Elseif bytType <> 90 Then

				arrTemp = Split(showString(objRS.fields("R_Fields").value),"|")
				For i = 1 to UBound(arrTemp)
					arrField(i,1) = arrTemp(i)
				Next

				arrTemp = Split(showString(objRS.fields("R_Params").value),"|")
				For i = 1 to UBound(arrTemp)
					arrParam(i,2) = arrTemp(i)
				Next

			End If

			datCreatedAfter = showDate(0,objRS.fields("R_CreatedAfter").value)
			datCreatedBefore = showDate(0,objRS.fields("R_CreatedBefore").value)
			strCreatedUser = showString(objRS.fields("R_CreatedUser").value)
			datModAfter = showDate(0,objRS.fields("R_ModAfter").value)
			datModBefore = showDate(0,objRS.fields("R_ModBefore").value)
			strModUser = showString(objRS.fields("R_ModUser").value)
			bytOrder = objRS.fields("R_Order").value
			strOrderDir = objRS.fields("R_OrderDir").value
			strReportName = showString(objRS.fields("R_Title").value)
			intPermissions = objRS.fields("R_Permissions").value
			strOwner = showString(objRS.fields("Owner").value)
			strCreatedBy = showString(objRS.fields("CreatedBy").value)
			datCreatedDate = objRS.fields("R_CreatedDate").value
			strModBy = showString(objRS.fields("ModBy").value)
			datModDate = objRS.fields("R_ModDate").value
		Else
			Call logError(3,1)
		End If
	Else
		strOwner = strFullName
		strCreatedBy = strFullName
		datCreatedDate = Now
		strModBy = strFullName
		datModDate = Now
	End If

	Select Case bytType
		Case 1
			arrField(0,0) = Application("IDS_Contact")
			arrField(1,0) = Application("IDS_Account")
			arrField(2,0) = Application("IDS_Division")
			arrField(3,0) = Application("IDS_Region")
			arrField(4,0) = Application("IDS_Address")
			arrField(5,0) = Application("IDS_City")
			arrField(6,0) = Application("IDS_State")
			arrField(7,0) = Application("IDS_Country")
			arrField(8,0) = Application("IDS_ZIP")
			arrField(9,0) = Application("IDS_Department")
			arrField(10,0) = Application("IDS_JobTitle")
			arrField(11,0) = Application("IDS_Email")
			arrField(12,0) = Application("IDS_Phone") & " 1"
			arrField(13,0) = Application("IDS_Phone") & " 2"
			arrField(14,0) = Application("IDS_Fax")
			arrField(15,0) = Application("IDS_CreatedBy")
			arrField(16,0) = Application("IDS_Created")
			arrField(17,0) = Application("IDS_ModifiedBy")
			arrField(18,0) = Application("IDS_Modified")
			arrField(19,0) = Application("IDS_Total") & " " & Application("IDS_Notes")
			arrField(20,0) = Application("IDS_Total") & " " & Application("IDS_Attachments")
			arrField(21,0) = Application("IDS_Total") & " " & Application("IDS_Sales")
			arrField(22,0) = Application("IDS_Total") & " " & Application("IDS_Products")
			arrField(23,0) = Application("IDS_Total") & " " & Application("IDS_Tickets")
			arrParam(1,0) = Application("IDS_Account")
			arrParam(1,1) = getTextField("txtParam1","oText",arrParam(1,2),20,150,"") & "<a href=""" & newWindow("S","?m=2&rVal=txtParam1") & """><img src=""../images/import.gif"" alt=""" & getImport("IDS_Account") & """ border=0 height=16 width=16></a>"
			arrParam(2,0) = Application("IDS_SalesRep")
			arrParam(2,1) = getTextField("txtParam2","oText",arrParam(2,2),20,150,"") & "<a href=""" & newWindow("S","?m=0&rVal=txtParam2") & """><img src=""../images/import.gif"" alt=""" & getImport("IDS_SalesRep") & """ border=0 height=16 width=16></a>"
			arrParam(3,0) = Application("IDS_City")
			arrParam(3,1) = getTextField("txtParam3","oText",arrParam(3,2),20,100,"")
			arrParam(4,0) = Application("IDS_State")
			arrParam(4,1) = getTextField("txtParam4","oText",arrParam(4,2),20,100,"")
			arrParam(5,0) = Application("IDS_Country")
			arrParam(5,1) = getCountries(140,"txtParam5",arrParam(5,2))
			arrParam(6,0) = Application("IDS_ZIP")
			arrParam(6,1) = getTextField("txtParam6","oText",arrParam(6,2),10,10,"")
			arrParam(7,0) = Application("IDS_Region")
			arrParam(7,1) = getOptionDropDown(140,True,"txtParam7","Sales Region",arrParam(7,2))
			arrParam(8,0) = Application("IDS_JobTitle")
			arrParam(8,1) = getTextField("txtParam8","oText",arrParam(8,2),10,10,"")
			arrParam(9,0) = Application("IDS_Min") & " " & Application("IDS_Notes")
			arrParam(9,1) = getTextField("txtParam9","oInt",arrParam(9,2),5,5,"")
			arrParam(10,0) = Application("IDS_Max") & " " & Application("IDS_Notes")
			arrParam(10,1) = getTextField("txtParam10","oInt",arrParam(10,2),5,5,"")
			arrParam(11,0) = Application("IDS_Min") & " " & Application("IDS_Attachments")
			arrParam(11,1) = getTextField("txtParam11","oInt",arrParam(11,2),5,5,"")
			arrParam(12,0) = Application("IDS_Max") & " " & Application("IDS_Attachments")
			arrParam(12,1) = getTextField("txtParam12","oInt",arrParam(12,2),5,5,"")
			arrParam(13,0) = Application("IDS_Min") & " " & Application("IDS_Sales")
			arrParam(13,1) = getTextField("txtParam13","oInt",arrParam(13,2),5,5,"")
			arrParam(14,0) = Application("IDS_Max") & " " & Application("IDS_Sales")
			arrParam(14,1) = getTextField("txtParam14","oInt",arrParam(14,2),5,5,"")
			arrParam(15,0) = Application("IDS_Min") & " " & Application("IDS_Products")
			arrParam(15,1) = getTextField("txtParam15","oInt",arrParam(15,2),5,5,"")
			arrParam(16,0) = Application("IDS_Max") & " " & Application("IDS_Products")
			arrParam(16,1) = getTextField("txtParam16","oInt",arrParam(16,2),5,5,"")
			arrParam(17,0) = Application("IDS_Min") & " " & Application("IDS_Tickets")
			arrParam(17,1) = getTextField("txtParam17","oInt",arrParam(17,2),5,5,"")
			arrParam(18,0) = Application("IDS_Max") & " " & Application("IDS_Tickets")
			arrParam(18,1) = getTextField("txtParam18","oInt",arrParam(18,2),5,5,"")
			intParamLimit = 18
			intFieldLimit = 23
		Case 2
			arrField(0,0) = Application("IDS_Account")
			arrField(1,0) = Application("IDS_Division")
			arrField(2,0) = Application("IDS_SalesRep")
			arrField(3,0) = Application("IDS_Region")
			arrField(4,0) = Application("IDS_Website")
			arrField(5,0) = Application("IDS_AccountType")
			arrField(6,0) = Application("IDS_Reference")
			arrField(7,0) = Application("IDS_IndustrySector")
			arrField(8,0) = Application("IDS_ProblemFlag")
			arrField(9,0) = Application("IDS_CreatedBy")
			arrField(10,0) = Application("IDS_Created")
			arrField(11,0) = Application("IDS_ModifiedBy")
			arrField(12,0) = Application("IDS_Modified")
			arrField(13,0) = Application("IDS_Total") & " " & Application("IDS_Notes")
			arrField(14,0) = Application("IDS_Total") & " " & Application("IDS_Attachments")
			arrField(15,0) = Application("IDS_Total") & " " & Application("IDS_Contacts")
			arrField(16,0) = Application("IDS_Total") & " " & Application("IDS_Sales")
			arrField(17,0) = Application("IDS_Total") & " " & Application("IDS_Products")
			arrField(18,0) = Application("IDS_Total") & " " & Application("IDS_Projects")
			arrField(19,0) = Application("IDS_Total") & " " & Application("IDS_Tickets")
			arrParam(1,0) = Application("IDS_Account")
			arrParam(1,1) = getTextField("txtParam1","oText",arrParam(1,2),20,150,"") & "<a href=""" & newWindow("S","?m=2&rVal=txtParam1") & """><img src=""../images/import.gif"" alt=""" & getImport("IDS_Account") & """ border=0 height=16 width=16></a>"
			arrParam(2,0) = Application("IDS_SalesRep")
			arrParam(2,1) = getTextField("txtParam2","oText",arrParam(2,2),20,150,"") & "<a href=""" & newWindow("S","?m=0&rVal=txtParam2") & """><img src=""../images/import.gif"" alt=""" & getImport("IDS_SalesRep") & """ border=0 height=16 width=16></a>"
			arrParam(3,0) = Application("IDS_Region")
			arrParam(3,1) = getOptionDropDown(140,True,"txtParam3","Sales Region",arrParam(3,2))
			arrParam(4,0) = Application("IDS_AccountType")
			arrParam(4,1) = getOptionDropDown(140,True,"txtParam4","Account Type",arrParam(4,2))
			arrParam(5,0) = Application("IDS_Reference")
			arrParam(5,1) = getTrueFalse("txtParam5",arrParam(5,2))
			arrParam(6,0) = Application("IDS_IndustrySector")
			arrParam(6,1) = getOptionDropDown(140,True,"txtParam6","Sales Vertical",arrParam(6,2))
			arrParam(7,0) = Application("IDS_ProblemFlag")
			arrParam(7,1) = getTrueFalse("txtParam7",arrParam(7,2))
			arrParam(9,0) = Application("IDS_Min") & " " & Application("IDS_Notes")
			arrParam(9,1) = getTextField("txtParam9","oInt",arrParam(9,2),5,5,"")
			arrParam(10,0) = Application("IDS_Max") & " " & Application("IDS_Notes")
			arrParam(10,1) = getTextField("txtParam10","oInt",arrParam(10,2),5,5,"")
			arrParam(11,0) = Application("IDS_Min") & " " & Application("IDS_Attachments")
			arrParam(11,1) = getTextField("txtParam11","oInt",arrParam(11,2),5,5,"")
			arrParam(12,0) = Application("IDS_Max") & " " & Application("IDS_Attachments")
			arrParam(12,1) = getTextField("txtParam12","oInt",arrParam(12,2),5,5,"")
			arrParam(13,0) = Application("IDS_Min") & " " & Application("IDS_Contacts")
			arrParam(13,1) = getTextField("txtParam13","oInt",arrParam(13,2),5,5,"")
			arrParam(14,0) = Application("IDS_Max") & " " & Application("IDS_Contacts")
			arrParam(14,1) = getTextField("txtParam14","oInt",arrParam(14,2),5,5,"")
			arrParam(15,0) = Application("IDS_Min") & " " & Application("IDS_Sales")
			arrParam(15,1) = getTextField("txtParam15","oInt",arrParam(15,2),5,5,"")
			arrParam(16,0) = Application("IDS_Max") & " " & Application("IDS_Sales")
			arrParam(16,1) = getTextField("txtParam16","oInt",arrParam(16,2),5,5,"")
			arrParam(17,0) = Application("IDS_Min") & " " & Application("IDS_Products")
			arrParam(17,1) = getTextField("txtParam17","oInt",arrParam(17,2),5,5,"")
			arrParam(18,0) = Application("IDS_Max") & " " & Application("IDS_Products")
			arrParam(18,1) = getTextField("txtParam18","oInt",arrParam(18,2),5,5,"")
			arrParam(19,0) = Application("IDS_Min") & " " & Application("IDS_Projects")
			arrParam(19,1) = getTextField("txtParam19","oInt",arrParam(19,2),5,5,"")
			arrParam(20,0) = Application("IDS_Max") & " " & Application("IDS_Projects")
			arrParam(20,1) = getTextField("txtParam20","oInt",arrParam(20,2),5,5,"")
			arrParam(21,0) = Application("IDS_Min") & " " & Application("IDS_Tickets")
			arrParam(21,1) = getTextField("txtParam21","oInt",arrParam(21,2),5,5,"")
			arrParam(22,0) = Application("IDS_Max") & " " & Application("IDS_Tickets")
			arrParam(22,1) = getTextField("txtParam22","oInt",arrParam(22,2),5,5,"")
			intParamLimit = 22
			intFieldLimit = 19
		Case 3
			arrField(0,0) = Application("IDS_Sale")
			arrField(1,0) = Application("IDS_Account")
			arrField(2,0) = Application("IDS_Division")
			arrField(3,0) = Application("IDS_SalesRep")
			arrField(4,0) = Application("IDS_Region")
			arrField(5,0) = Application("IDS_IndustrySector")
			arrField(6,0) = Application("IDS_Phase")
			arrField(7,0) = Application("IDS_Pipeline")
			arrField(8,0) = Application("IDS_Invoice")
			arrField(9,0) = Application("IDS_Currency")
			arrField(10,0) = Application("IDS_Closed")
			arrField(11,0) = Application("IDS_CloseDate")
			arrField(12,0) = Application("IDS_SaleValue")
			arrField(13,0) = Application("IDS_CreatedBy")
			arrField(14,0) = Application("IDS_Created")
			arrField(15,0) = Application("IDS_ModifiedBy")
			arrField(16,0) = Application("IDS_Modified")
			arrField(17,0) = Application("IDS_Total") & " " & Application("IDS_Notes")
			arrField(18,0) = Application("IDS_Total") & " " & Application("IDS_Attachments")
			arrParam(1,0) = Application("IDS_Account")
			arrParam(1,1) = getTextField("txtParam1","oText",arrParam(1,2),20,150,"") & "<a href=""" & newWindow("S","?m=2&rVal=txtParam1") & """><img src=""../images/import.gif"" alt=""" & getImport("IDS_Account") & """ border=0 height=16 width=16></a>"
			arrParam(2,0) = Application("IDS_SalesRep")
			arrParam(2,1) = getTextField("txtParam2","oText",arrParam(2,2),20,100,"") & "<a href=""" & newWindow("S","?m=0&rVal=txtParam2") & """><img src=""../images/import.gif"" alt=""" & getImport("IDS_SalesRep") & """ border=0 height=16 width=16></a>"
			arrParam(3,0) = Application("IDS_Region")
			arrParam(3,1) = getOptionDropDown(140,True,"txtParam3","Sales Region",arrParam(3,2))
			arrParam(4,0) = Application("IDS_IndustrySector")
			arrParam(4,1) = getOptionDropDown(140,True,"txtParam4","Sales Vertical",arrParam(4,2))
			arrParam(5,0) = Application("IDS_Phase")
			arrParam(5,1) = getOptionDropDown(140,True,"txtParam5","Sales Phase",arrParam(5,2))
			arrParam(6,0) = Application("IDS_Currency")
			arrParam(6,1) = getCurrency(140,"txtParam6",arrParam(6,2))
			arrParam(7,0) = Application("IDS_Closed")
			arrParam(7,1) = getTrueFalse("txtParam7",arrParam(7,2))
			arrParam(9,0) = Application("IDS_Closed") & " " & Application("IDS_After")
			arrParam(9,1) = getTextField("txtParam9","oDate",arrParam(9,2),12,10,"") & "&nbsp;<a href=""Javascript:showCalendar('txtParam9');""><img src=""../images/cal.gif"" alt=""" & getImport("IDS_ClosedAfter") & """ border=0 height=16 width=16></a>"
			arrParam(10,0) = Application("IDS_Closed") & " " & Application("IDS_Before")
			arrParam(10,1) = getTextField("txtParam10","oDate",arrParam(10,2),12,10,"") & "&nbsp;<a href=""Javascript:showCalendar('txtParam10');""><img src=""../images/cal.gif"" alt=""" & getImport("IDS_ClosedBefore") & """ border=0 height=16 width=16></a>"
			arrParam(11,0) = Application("IDS_Min") & " " & Application("IDS_SaleValue")
			arrParam(11,1) = getTextField("txtParam11","oCurrency",arrParam(11,2),8,11,"")
			arrParam(12,0) = Application("IDS_Max") & " " & Application("IDS_SaleValue")
			arrParam(12,1) = getTextField("txtParam12","oCurrency",arrParam(12,2),8,11,"")
			arrParam(13,0) = Application("IDS_Min") & " " & Application("IDS_Pipeline")
			arrParam(13,1) = getTextField("txtParam13","oByte",arrParam(13,2),5,5,"") & "%"
			arrParam(14,0) = Application("IDS_Max") & " " & Application("IDS_Pipeline")
			arrParam(14,1) = getTextField("txtParam14","oByte",arrParam(14,2),5,5,"") & "%"
			arrParam(15,0) = Application("IDS_Min") & " " & Application("IDS_Notes")
			arrParam(15,1) = getTextField("txtParam15","oInt",arrParam(15,2),5,5,"")
			arrParam(16,0) = Application("IDS_Max") & " " & Application("IDS_Notes")
			arrParam(16,1) = getTextField("txtParam16","oInt",arrParam(16,2),5,5,"")
			arrParam(17,0) = Application("IDS_Min") & " " & Application("IDS_Attachments")
			arrParam(17,1) = getTextField("txtParam17","oInt",arrParam(17,2),5,5,"")
			arrParam(18,0) = Application("IDS_Max") & " " & Application("IDS_Attachments")
			arrParam(18,1) = getTextField("txtParam18","oInt",arrParam(18,2),5,5,"")
			intParamLimit = 18
			intFieldLimit = 18
		Case 4
			arrField(0,0) = Application("IDS_Project")
			arrField(1,0) = Application("IDS_Account")
			arrField(2,0) = Application("IDS_Division")
			arrField(3,0) = Application("IDS_Sale")
			arrField(4,0) = Application("IDS_Owner")
			arrField(5,0) = Application("IDS_DaysTotal")
			arrField(6,0) = Application("IDS_DaysOwed")
			arrField(7,0) = Application("IDS_Closed")
			arrField(8,0) = Application("IDS_CloseDate")
			arrField(9,0) = Application("IDS_CreatedBy")
			arrField(10,0) = Application("IDS_Created")
			arrField(11,0) = Application("IDS_ModifiedBy")
			arrField(12,0) = Application("IDS_Modified")
			arrField(13,0) = Application("IDS_Total") & " " & Application("IDS_Notes")
			arrField(14,0) = Application("IDS_Total") & " " & Application("IDS_Attachments")
			arrField(15,0) = Application("IDS_Total") & " " & Application("IDS_Events")
			arrParam(1,0) = Application("IDS_Project")
			arrParam(1,1) = getTextField("txtParam1","oText",arrParam(1,2),20,140,"") & "<a href=""" & newWindow("S","?m=4&rVal=txtParam1") & """><img src=""../images/import.gif"" alt=""" & getImport("IDS_Project") & """ border=0 height=16 width=16></a>"
			arrParam(2,0) = Application("IDS_Account")
			arrParam(2,1) = getTextField("txtParam2","oText",arrParam(2,2),20,140,"") & "<a href=""" & newWindow("S","?m=2&rVal=txtParam2") & """><img src=""../images/import.gif"" alt=""" & getImport("IDS_Account") & """ border=0 height=16 width=16></a>"
			arrParam(3,0) = Application("IDS_Owner")
			arrParam(3,1) = getTextField("txtParam3","oText",arrParam(3,2),20,140,"") & "<a href=""" & newWindow("S","?m=0&rVal=txtParam3") & """><img src=""../images/import.gif"" alt=""" & getImport("IDS_Owner") & """ border=0 height=16 width=16></a>"
			arrParam(4,0) = Application("IDS_Closed")
			arrParam(4,1) = getTrueFalse("txtParam4",arrParam(4,2))
			arrParam(5,0) = Application("IDS_Min") & " " & Application("IDS_DaysTotal")
			arrParam(5,1) = getTextField("txtParam5","oInt",arrParam(5,2),5,5,"")
			arrParam(6,0) = Application("IDS_Max") & " " & Application("IDS_DaysTotal")
			arrParam(6,1) = getTextField("txtParam6","oInt",arrParam(6,2),5,5,"")
			arrParam(7,0) = Application("IDS_Min") & " " & Application("IDS_DaysOwed")
			arrParam(7,1) = getTextField("txtParam7","oInt",arrParam(7,2),5,5,"")
			arrParam(8,0) = Application("IDS_Max") & " " & Application("IDS_DaysOwed")
			arrParam(8,1) = getTextField("txtParam8","oInt",arrParam(8,2),5,5,"")
			arrParam(9,0) = Application("IDS_Min") & " " & Application("IDS_Notes")
			arrParam(9,1) = getTextField("txtParam9","oInt",arrParam(9,2),5,5,"")
			arrParam(10,0) = Application("IDS_Max") & " " & Application("IDS_Notes")
			arrParam(10,1) = getTextField("txtParam10","oInt",arrParam(10,2),5,5,"")
			arrParam(11,0) = Application("IDS_Min") & " " & Application("IDS_Attachments")
			arrParam(11,1) = getTextField("txtParam11","oInt",arrParam(11,2),5,5,"")
			arrParam(12,0) = Application("IDS_Max") & " " & Application("IDS_Attachments")
			arrParam(12,1) = getTextField("txtParam12","oInt",arrParam(12,2),5,5,"")
			intParamLimit = 12
			intFieldLimit = 15
		Case 5
			arrField(0,0) = Application("IDS_TicketId")
			arrField(1,0) = Application("IDS_Account")
			arrField(2,0) = Application("IDS_Contact")
			arrField(3,0) = Application("IDS_Owner")
			arrField(4,0) = Application("IDS_HotIssue")
			arrField(5,0) = Application("IDS_Priority")
			arrField(6,0) = Application("IDS_TicketType")
			arrField(7,0) = Application("IDS_TicketSource")
			arrField(8,0) = Application("IDS_SupportType")
			arrField(9,0) = Application("IDS_Product")
			arrField(10,0) = Application("IDS_Version")
			arrField(11,0) = Application("IDS_BugId")
			arrField(12,0) = Application("IDS_Cause")
			arrField(13,0) = Application("IDS_Closed")
			arrField(14,0) = Application("IDS_CloseDate")
			arrField(15,0) = Application("IDS_CreatedBy")
			arrField(16,0) = Application("IDS_Created")
			arrField(17,0) = Application("IDS_ModifiedBy")
			arrField(18,0) = Application("IDS_Modified")
			arrField(19,0) = Application("IDS_Total") & " " & Application("IDS_Notes")
			arrField(20,0) = Application("IDS_Total") & " " & Application("IDS_Attachments")
			arrParam(1,0) = Application("IDS_TicketId")
			arrParam(1,1) = getTextField("txtParam1","oLong",arrParam(1,2),10,10,"") & "<a href=""" & newWindow("S","?m=5&rVal=txtParam1") & """><img src=""../images/import.gif"" alt=""" & getImport("IDS_TicketId") & """ border=0 height=16 width=16></a>"
			arrParam(2,0) = Application("IDS_Account")
			arrParam(2,1) = getTextField("txtParam2","oText",arrParam(2,2),20,140,"") & "<a href=""" & newWindow("S","?m=2&rVal=txtParam2") & """><img src=""../images/import.gif"" alt=""" & getImport("IDS_Account") & """ border=0 height=16 width=16></a>"
			arrParam(3,0) = Application("IDS_Contact")
			arrParam(3,1) = getTextField("txtParam3","oText",arrParam(3,2),20,140,"") & "<a href=""" & newWindow("S","?m=1&rVal=txtParam3") & """><img src=""../images/import.gif"" alt=""" & getImport("IDS_Contact") & """ border=0 height=16 width=16></a>"
			arrParam(4,0) = Application("IDS_Owner")
			arrParam(4,1) = getTextField("txtParam4","oText",arrParam(4,2),20,140,"") & "<a href=""" & newWindow("S","?m=0&rVal=txtParam4") & """><img src=""../images/import.gif"" alt=""" & getImport("IDS_Owner") & """ border=0 height=16 width=16></a>"
			arrParam(5,0) = Application("IDS_HotIssue")
			arrParam(5,1) = getTrueFalse("txtParam5",arrParam(5,2))
			arrParam(6,0) = Application("IDS_Priority")
			arrParam(6,1) = getOptionDropDown(140,True,"txtParam6","Priority",arrParam(6,2))
			arrParam(7,0) = Application("IDS_TicketType")
			arrParam(7,1) = getOptionDropDown(140,True,"txtParam7","Ticket Type",arrParam(7,2))
			arrParam(8,0) = Application("IDS_TicketSource")
			arrParam(8,1) = getOptionDropDown(140,True,"txtParam8","Ticket Source",arrParam(8,2))
			arrParam(9,0) = Application("IDS_SupportType")
			arrParam(9,1) = getOptionDropDown(140,True,"txtParam9","Support Type",arrParam(9,2))
			arrParam(10,0) = Application("IDS_Product")
			arrParam(10,1) = getProductDropDown(140,True,"txtParam10",arrParam(10,2))
			arrParam(11,0) = Application("IDS_Description")
			arrParam(11,1) = getTextField("txtParam11","oText",arrParam(11,2),40,255,"")
			arrParam(12,0) = Application("IDS_Solution")
			arrParam(12,1) = getTextField("txtParam12","oText",arrParam(12,2),40,255,"")
			arrParam(13,0) = Application("IDS_Cause")
			arrParam(13,1) = getOptionDropDown(140,True,"txtParam13","Ticket Cause",arrParam(13,2))
			arrParam(14,0) = Application("IDS_Closed")
			arrParam(14,1) = getTrueFalse("txtParam14",arrParam(14,2))
			arrParam(15,0) = Application("IDS_Min") & " " & Application("IDS_Notes")
			arrParam(15,1) = getTextField("txtParam15","oInt",arrParam(15,2),5,5,"")
			arrParam(16,0) = Application("IDS_Max") & " " & Application("IDS_Notes")
			arrParam(16,1) = getTextField("txtParam16","oInt",arrParam(16,2),5,5,"")
			arrParam(17,0) = Application("IDS_Min") & " " & Application("IDS_Attachments")
			arrParam(17,1) = getTextField("txtParam17","oInt",arrParam(17,2),5,5,"")
			arrParam(18,0) = Application("IDS_Max") & " " & Application("IDS_Attachments")
			arrParam(18,1) = getTextField("txtParam18","oInt",arrParam(18,2),5,5,"")
			intParamLimit = 18
			intFieldLimit = 20
		Case 6
			arrField(0,0) = Application("IDS_Bug")
			arrField(1,0) = Application("IDS_Owner")
			arrField(2,0) = Application("IDS_HotIssue")
			arrField(3,0) = Application("IDS_Priority")
			arrField(4,0) = Application("IDS_BugType")
			arrField(5,0) = Application("IDS_BugSource")
			arrField(6,0) = Application("IDS_Product")
			arrField(7,0) = Application("IDS_Version")
			arrField(8,0) = Application("IDS_Cause")
			arrField(9,0) = Application("IDS_Closed")
			arrField(10,0) = Application("IDS_CloseDate")
			arrField(11,0) = Application("IDS_CreatedBy")
			arrField(12,0) = Application("IDS_Created")
			arrField(13,0) = Application("IDS_ModifiedBy")
			arrField(14,0) = Application("IDS_Modified")
			arrField(15,0) = Application("IDS_Total") & " " & Application("IDS_Notes")
			arrField(16,0) = Application("IDS_Total") & " " & Application("IDS_Attachments")
			arrField(17,0) = Application("IDS_Total") & " " & Application("IDS_Tickets")
			arrParam(1,0) = Application("IDS_BugId")
			arrParam(1,1) = getTextField("txtParam1","oLong",arrParam(1,2),10,10,"") & "<a href=""" & newWindow("S","?m=6&rVal=txtParam1") & """><img src=""../images/import.gif"" alt=""" & getImport("IDS_BugId") & """ border=0 height=16 width=16></a>"
			arrParam(2,0) = Application("IDS_Owner")
			arrParam(2,1) = getTextField("txtParam2","oText",arrParam(2,2),20,140,"") & "<a href=""" & newWindow("S","?m=0&rVal=txtParam2") & """><img src=""../images/import.gif"" alt=""" & getImport("IDS_Owner") & """ border=0 height=16 width=16></a>"
			arrParam(3,0) = Application("IDS_HotIssue")
			arrParam(3,1) = getTrueFalse("txtParam3",arrParam(3,2))
			arrParam(4,0) = Application("IDS_Priority")
			arrParam(4,1) = getOptionDropDown(140,True,"txtParam4","Priority",arrParam(4,2))
			arrParam(5,0) = Application("IDS_BugType")
			arrParam(5,1) = getOptionDropDown(140,True,"txtParam5","Bug Type",arrParam(5,2))
			arrParam(6,0) = Application("IDS_BugSource")
			arrParam(6,1) = getOptionDropDown(140,True,"txtParam6","Bug Source",arrParam(6,2))
			arrParam(7,0) = Application("IDS_Description")
			arrParam(7,1) = getTextField("txtParam7","oText",arrParam(7,2),40,255,"")
			arrParam(8,0) = Application("IDS_Solution")
			arrParam(8,1) = getTextField("txtParam8","oText",arrParam(8,2),40,255,"")
			arrParam(9,0) = Application("IDS_Product")
			arrParam(9,1) = getProductDropDown(140,True,"txtParam9",arrParam(9,2))
			arrParam(10,0) = Application("IDS_Cause")
			arrParam(10,1) = getOptionDropDown(140,True,"txtParam10","Bug Cause",arrParam(10,2))
			arrParam(11,0) = Application("IDS_Closed")
			arrParam(11,1) = getTrueFalse("txtParam11",arrParam(11,2))
			arrParam(12,0) = ""
			arrParam(12,1) = ""
			arrParam(13,0) = Application("IDS_Min") & " " & Application("IDS_Notes")
			arrParam(13,1) = getTextField("txtParam13","oInt",arrParam(13,2),5,5,"")
			arrParam(14,0) = Application("IDS_Max") & " " & Application("IDS_Notes")
			arrParam(14,1) = getTextField("txtParam14","oInt",arrParam(14,2),5,5,"")
			arrParam(15,0) = Application("IDS_Min") & " " & Application("IDS_Attachments")
			arrParam(15,1) = getTextField("txtParam15","oInt",arrParam(15,2),5,5,"")
			arrParam(16,0) = Application("IDS_Max") & " " & Application("IDS_Attachments")
			arrParam(16,1) = getTextField("txtParam16","oInt",arrParam(16,2),5,5,"")
			intParamLimit = 16
			intFieldLimit = 17
		Case 7
			arrField(0,0) = Application("IDS_InvoiceId")
			arrField(1,0) = Application("IDS_Account")
			arrField(2,0) = Application("IDS_Contact")
			arrField(3,0) = Application("IDS_Owner")
			arrField(4,0) = Application("IDS_PurchaseOrder")
			arrField(5,0) = Application("IDS_InvoiceReceived")
			arrField(6,0) = Application("IDS_Type")
			arrField(7,0) = Application("IDS_Phase")
			arrField(8,0) = Application("IDS_Currency")
			arrField(9,0) = Application("IDS_Value")
			arrField(10,0) = Application("IDS_Tax")
			arrField(11,0) = Application("IDS_InvoiceDate")
			arrField(12,0) = Application("IDS_InvoiceDue")
			arrField(13,0) = Application("IDS_InvoicePaid")
			arrField(14,0) = Application("IDS_Closed")
			arrField(15,0) = Application("IDS_CreatedBy")
			arrField(16,0) = Application("IDS_Created")
			arrField(17,0) = Application("IDS_ModifiedBy")
			arrField(18,0) = Application("IDS_Modified")
			arrField(19,0) = Application("IDS_Total") & " " & Application("IDS_Notes")
			arrField(20,0) = Application("IDS_Total") & " " & Application("IDS_Attachments")
			arrField(21,0) = Application("IDS_Total") & " " & Application("IDS_Events")
			arrField(22,0) = Application("IDS_Total") & " " & Application("IDS_Sales")
			arrField(23,0) = Application("IDS_Total") & " " & Application("IDS_Products")
			arrField(24,0) = Application("IDS_Total") & " " & Application("IDS_Projects")
			arrParam(1,0) = Application("IDS_Account")
			arrParam(1,1) = getTextField("txtParam1","oText",arrParam(1,2),20,140,"") & "<a href=""" & newWindow("S","?m=2&rVal=txtParam1") & """><img src=""../images/import.gif"" alt=""" & getImport("IDS_Account") & """ border=0 height=16 width=16></a>"
			arrParam(2,0) = Application("IDS_Contact")
			arrParam(2,1) = getTextField("txtParam2","oText",arrParam(2,2),20,140,"") & "<a href=""" & newWindow("S","?m=1&rVal=txtParam2") & """><img src=""../images/import.gif"" alt=""" & getImport("IDS_Contact") & """ border=0 height=16 width=16></a>"
			arrParam(3,0) = Application("IDS_Owner")
			arrParam(3,1) = getTextField("txtParam3","oText",arrParam(3,2),20,140,"") & "<a href=""" & newWindow("S","?m=0&rVal=txtParam3") & """><img src=""../images/import.gif"" alt=""" & getImport("IDS_Owner") & """ border=0 height=16 width=16></a>"
			arrParam(4,0) = Application("IDS_InvoiceReceived")
			arrParam(4,1) = getTrueFalse("txtParam4",arrParam(4,2))
			arrParam(5,0) = Application("IDS_Type")
			arrParam(5,1) = getOptionDropDown(140,True,"txtParam5","Invoice Type",arrParam(5,2))
			arrParam(6,0) = Application("IDS_Phase")
			arrParam(6,1) = getOptionDropDown(140,True,"txtParam6","Invoice Phase",arrParam(6,2))
			arrParam(7,0) = Application("IDS_Closed")
			arrParam(7,1) = getTrueFalse("txtParam7",arrParam(7,2))
			arrParam(8,0) = Application("IDS_Currency")
			arrParam(8,1) = getCurrency(140,"txtParam8",arrParam(8,2))
			arrParam(9,0) = Application("IDS_Min") & " " & Application("IDS_Value")
			arrParam(9,1) = getTextField("txtParam9","oCurrency",arrParam(9,2),5,12,"")
			arrParam(10,0) = Application("IDS_Max") & " " & Application("IDS_Value")
			arrParam(10,1) = getTextField("txtParam10","oCurrency",arrParam(10,2),5,12,"")
			arrParam(11,0) = Application("IDS_InvoiceDate") & " " & Application("IDS_After")
			arrParam(11,1) = getTextField("txtParam11","oDate",arrParam(11,2),12,10,"") & "&nbsp;<a href=""Javascript:showCalendar('txtParam11');""><img src=""../images/cal.gif"" alt=""" & getImport("IDS_ClosedAfter") & """ border=0 height=16 width=16></a>"
			arrParam(12,0) = Application("IDS_InvoiceDate") & " " & Application("IDS_Before")
			arrParam(12,1) = getTextField("txtParam12","oDate",arrParam(12,2),12,10,"") & "&nbsp;<a href=""Javascript:showCalendar('txtParam12');""><img src=""../images/cal.gif"" alt=""" & getImport("IDS_ClosedBefore") & """ border=0 height=16 width=16></a>"
			arrParam(13,0) = Application("IDS_InvoiceDue") & " " & Application("IDS_After")
			arrParam(13,1) = getTextField("txtParam13","oDate",arrParam(13,2),12,10,"") & "&nbsp;<a href=""Javascript:showCalendar('txtParam13');""><img src=""../images/cal.gif"" alt=""" & getImport("IDS_ClosedAfter") & """ border=0 height=16 width=16></a>"
			arrParam(14,0) = Application("IDS_InvoiceDue") & " " & Application("IDS_Before")
			arrParam(14,1) = getTextField("txtParam14","oDate",arrParam(14,2),12,10,"") & "&nbsp;<a href=""Javascript:showCalendar('txtParam14');""><img src=""../images/cal.gif"" alt=""" & getImport("IDS_ClosedAfter") & """ border=0 height=16 width=16></a>"
			arrParam(15,0) = Application("IDS_InvoicePaid") & " " & Application("IDS_After")
			arrParam(15,1) = getTextField("txtParam15","oDate",arrParam(15,2),12,10,"") & "&nbsp;<a href=""Javascript:showCalendar('txtParam15');""><img src=""../images/cal.gif"" alt=""" & getImport("IDS_ClosedAfter") & """ border=0 height=16 width=16></a>"
			arrParam(16,0) = Application("IDS_InvoicePaid") & " " & Application("IDS_Before")
			arrParam(16,1) = getTextField("txtParam16","oDate",arrParam(16,2),12,10,"") & "&nbsp;<a href=""Javascript:showCalendar('txtParam16');""><img src=""../images/cal.gif"" alt=""" & getImport("IDS_ClosedAfter") & """ border=0 height=16 width=16></a>"
			arrParam(17,0) = Application("IDS_Min") & " " & Application("IDS_Notes")
			arrParam(17,1) = getTextField("txtParam17","oInt",arrParam(17,2),5,5,"")
			arrParam(18,0) = Application("IDS_Max") & " " & Application("IDS_Notes")
			arrParam(18,1) = getTextField("txtParam18","oInt",arrParam(18,2),5,5,"")
			arrParam(19,0) = Application("IDS_Min") & " " & Application("IDS_Attachments")
			arrParam(19,1) = getTextField("txtParam19","oInt",arrParam(19,2),5,5,"")
			arrParam(20,0) = Application("IDS_Max") & " " & Application("IDS_Attachments")
			arrParam(20,1) = getTextField("txtParam20","oInt",arrParam(20,2),5,5,"")
			intParamLimit = 20
			intFieldLimit = 24
		Case 50
			strTitle = Application("IDS_ReportNewEvent")
			arrField(0,0) = Application("IDS_Event")
			arrField(1,0) = Application("IDS_Owner")
			arrField(2,0) = Application("IDS_Onsite")
			arrField(3,0) = Application("IDS_Billable")
			arrField(4,0) = Application("IDS_EventType")
			arrField(5,0) = Application("IDS_StartTime")
			arrField(6,0) = Application("IDS_EndTime")
			arrField(7,0) = Application("IDS_CreatedBy")
			arrField(8,0) = Application("IDS_Created")
			arrField(9,0) = Application("IDS_ModifiedBy")
			arrField(10,0) = Application("IDS_Modified")
			arrParam(1,0) = Application("IDS_Owner")
			arrParam(1,1) = getTextField("txtParam1","oText",arrParam(1,2),20,140,"") & "<a href=""" & newWindow("S","?m=0&rVal=txtParam1") & """><img src=""../images/import.gif"" alt=""" & getImport("IDS_Owner") & """ border=0 height=16 width=16></a>"
			arrParam(2,0) = Application("IDS_Type")
			arrParam(2,1) = getOptionDropDown(140,True,"txtParam2","Event Type",arrParam(2,2))
			arrParam(3,0) = Application("IDS_Start") & " " & Application("IDS_After")
			arrParam(3,1) = getTextField("txtParam3","oDate",arrParam(3,2),10,10,"") & "&nbsp;<a href=""Javascript:showCalendar('txtParam3');""><img src=""../images/cal.gif"" alt=""" & getImport("IDS_StartAfter") & """ border=0 height=16 width=16></a>"
			arrParam(4,0) = Application("IDS_Start") & " " & Application("IDS_Before")
			arrParam(4,1) = getTextField("txtParam4","oDate",arrParam(4,2),10,10,"") & "&nbsp;<a href=""Javascript:showCalendar('txtParam4');""><img src=""../images/cal.gif"" alt=""" & getImport("IDS_StartBefore") & """ border=0 height=16 width=16></a>"
			arrParam(5,0) = Application("IDS_End") & " " & Application("IDS_After")
			arrParam(5,1) = getTextField("txtParam5","oDate",arrParam(5,2),10,10,"") & "&nbsp;<a href=""Javascript:showCalendar('txtParam5');""><img src=""../images/cal.gif"" alt=""" & getImport("IDS_EndBefore") & """ border=0 height=16 width=16></a>"
			arrParam(6,0) = Application("IDS_End") & " " & Application("IDS_Before")
			arrParam(6,1) = getTextField("txtParam6","oDate",arrParam(6,2),10,10,"") & "&nbsp;<a href=""Javascript:showCalendar('txtParam6');""><img src=""../images/cal.gif"" alt=""" & getImport("IDS_EndAfter") & """ border=0 height=16 width=16></a>"
			arrParam(7,0) = Application("IDS_Billable")
			arrParam(7,1) = getTrueFalse("txtParam7",arrParam(7,2))
			arrParam(8,0) = Application("IDS_Onsite")
			arrParam(8,1) = getTrueFalse("txtParam8",arrParam(8,2))
			intParamLimit = 8
			intFieldLimit = 10
		Case 90
			intPerm = 5
	End Select

	strIncHead = getCalendarScripts()

	Call DisplayHeader(1)
%>
<div id="headerDiv" class="dvBorder">

<table border=0 cellspacing=0 cellpadding=5 height=20 width="100%">
  <tr class="hrow">
    <td valign=bottom><span class="tfont"><% =strTitle %></span></td>
    <td valign=top align=right>
      <table border=0 cellspacing=0 cellpadding=0>
        <tr>
          <td class="dfont"><% =Application("IDS_Created") %>:</td>
          <td class="dfont">&nbsp;&nbsp;<% =showDate(0,datCreatedDate) %></td>
          <td class="dfont">&nbsp;&nbsp;<% =strCreatedBy %></td>
        </tr>
        <tr>
          <td class="dfont"><% =Application("IDS_Modified") %>:</td>
          <td class="dfont">&nbsp;&nbsp;<% =showDate(0,datModDate) %></td>
          <td class="dfont">&nbsp;&nbsp;<% =strModBy %></td>
        </tr>
      </table>
    </td>
  </tr>
</table>

</div>

<div id="contentDiv" class="dvBorder" style="height:<% =intScreenH-170 %>px;">

<%
	If bytType = 0 Then
		If pContacts >= 1 Then Response.Write("<li><a href=""create_report.asp?m=1&type=1"" class=""dfont"">" & Application("IDS_ReportNewContact") & "</a></li>" & vbCrLf)
		If pClients >= 1 Then Response.Write("<li><a href=""create_report.asp?m=2&type=2"" class=""dfont"">" & Application("IDS_ReportNewClient") & "</a></li>" & vbCrLf)
		If pSales >= 1 Then Response.Write("<li><a href=""create_report.asp?m=3&type=3"" class=""dfont"">" & Application("IDS_ReportNewSales") & "</a></li>" & vbCrLf)
		If pProjects >= 1 Then Response.Write("<li><a href=""create_report.asp?m=4&type=4"" class=""dfont"">" & Application("IDS_ReportNewProject") & "</a></li>" & vbCrLf & _
							"<li><a href=""create_report.asp?m=4&type=50"" class=""dfont"">" & Application("IDS_ReportNewEvent") & "</a></li>" & vbCrLf)
		If pTickets >= 1 Then Response.Write("<li><a href=""create_report.asp?m=5&type=5"" class=""dfont"">" & Application("IDS_ReportNewTicket") & "</a></li>" & vbCrLf)
		If pBugs >= 1 Then Response.Write("<li><a href=""create_report.asp?m=6&type=6"" class=""dfont"">" & Application("IDS_ReportNewBug") & "</a></li>" & vbCrLf)
		If pInvoices >= 1 Then Response.Write("<li><a href=""create_report.asp?m=7&type=7"" class=""dfont"">" & Application("IDS_ReportNewInvoice") & "</a></li>" & vbCrLf)
		Response.Write(	"<br><br>" & vbCrLf)
		If blnAdmin Then Response.Write("<li><a href=""create_report.asp?type=90"" class=""dfont"">" & Application("IDS_ReportNewCustom") & "</a></li>" & vbCrLf)
		Response.Write("</div>" & vbCrLf)
	Else
%>

<table border=0 cellpadding=5 width="100%">
<form name="frmReport" method="post" action="view_report.asp?id=<% =lngRecordId %>&m=<% =bytMod %>&type=<% =bytType %>">
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnWinOpen","") %>
<% =getHidden("hdnChange",blnChange) %>
<% =getHidden("hdnParams",intParamLimit) %>
<% =getHidden("hdnFields",intFieldLimit) %>
  <tr>
    <td rowspan=2 valign=top><% =getLabel(Application("IDS_ReportName"),"txtReportName") %></td>
    <td rowspan=2 valign=top><% =getTextField("txtReportName","mText",strReportName,20,100,"") %></td>
    <td><% =getLabel(Application("IDS_Owner"),"txtOwner") %></td>
    <td>
      <% If intPerm = 5 Then %>
      <% =getTextField("txtOwner","mText",strOwner,20,255,"") %>
      <a href="<% =newWindow("S","?m=0&rVal=txtOwner") %>"><img src="../images/import.gif" alt="<% =getImport("IDS_Owner") %>" border=0 height=16 width=16></a>
      <% Else %>
      <% =getTextField("txtOwner","dText",strOwner,20,20,"readonly=""readonly""") %>
      <% End If %>
    </td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_Permissions"),"selPermissions") %></td>
    <td>
      <select name="selPermissions" id="selPermissions" class="oByte" style="width:140px;">
        <option value="1"<% =getDefault(0,intPermissions,1) & ">" & Application("IDS_AccessPrivate") %></option>
        <% If intPerm >= 5 Then %>
        <option value="2"<% =getDefault(0,intPermissions,2) & ">" & Application("IDS_MembersOnly") %></option>
        <option value="3"<% =getDefault(0,intPermissions,3) & ">" & Application("IDS_InternalView") %></option>
        <% End If %>
      </select>
    </td>
  </tr>
  </tr>
  <tr><td colspan=4><hr></td></tr>
  <% If bytType = 90 Then %>
  <tr>
    <td colspan=4><% =getLabel(Application("IDS_CustomSQL"),"txtSQL") %><br>
    <% =getTextArea("txtSQL","oText",strGenSQL,"100%",8,"") %>
    <br><br><% =getLabel(Application("IDS_Module"),"selModule") %> &nbsp;&nbsp;&nbsp;&nbsp;
    <% =getModuleDropDown("selModule",bytMod,True,"") %></td>
  </tr>
  <% Else

    	For i = 1 to intParamLimit
    		If (i+1) Mod 2 = 0 Then Response.Write("  <tr>" & vbCrLf)

		Response.Write(vbTab & "<td width=""15%"">" & getLabel(arrParam(i,0),"txtParam" & i) & "</td>" & vbCrLf & _
    				vbTab & "<td width=""35%"" class=""dfont"">" & arrParam(i,1) & "</td>" & vbCrLf)

    		If (i+2) Mod 2 = 0 Then Response.Write("  </tr>" & vbCrLf)
    	Next
%>
  <tr><td colspan=4><hr></td></tr>
  <tr>
    <td><% =getLabel(Application("IDS_CreatedAfter"),"txtCreatedAfter") %></td>
    <td>
      <% =getTextField("txtCreatedAfter","oDate",datCreatedAfter,12,10,"") %>
      <a href="Javascript:showCalendar('txtCreatedAfter');"><img src="../images/cal.gif" alt="<% =getImport("IDS_CreatedAfter") %>" border=0 height=16 width=16></a>
    </td>
    <td><% =getLabel(Application("IDS_ModifiedAfter"),"txtModAfter") %></td>
    <td>
      <% =getTextField("txtModAfter","oDate",datModAfter,12,10,"") %>
      <a href="Javascript:showCalendar('txtModAfter');"><img src="../images/cal.gif" alt="<% =getImport("IDS_ModAfter") %>" border=0 height=16 width=16></a>
    </td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_CreatedBefore"),"txtCreatedBefore") %></td>
    <td>
      <% =getTextField("txtCreatedBefore","oDate",datCreatedBefore,12,10,"") %>
      <a href="Javascript:showCalendar('txtCreatedBefore');"><img src="../images/cal.gif" alt="<% =getImport("IDS_CreatedBefore") %>" border=0 height=16 width=16></a>
    </td>
    <td><% =getLabel(Application("IDS_ModifiedBefore"),"txtModBefore") %></td>
    <td>
      <% =getTextField("txtModBefore","oDate",datModBefore,12,10,"") %>
      <a href="Javascript:showCalendar('txtModBefore');"><img src="../images/cal.gif" alt="<% =getImport("IDS_CreatedBefore") %>" border=0 height=16 width=16></a>
    </td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_CreatedBy"),"txtCreatedUser") %></td>
    <td>
      <% =getTextField("txtCreatedUser","oText",strCreatedUser,20,150,"") %>
      <a href="<% =newWindow("S","?m=0&rVal=txtCreatedUser") %>"><img src="../images/import.gif" alt="<% =getImport("IDS_CreatedBy") %>" border=0 height=16 width=16></a>
    </td>
    <td><% =getLabel(Application("IDS_ModifiedBy"),"txtModUser") %></td>
    <td>
      <% =getTextField("txtModUser","oText",strModUser,20,150,"") %>
      <a href="<% =newWindow("S","?m=0&rVal=txtModUser") %>"><img src="../images/import.gif" alt="<% =getImport("IDS_ModBy") %>" border=0 height=16 width=16></a>
    </td>
  </tr>
  <tr><td colspan=4><hr></td></tr>
  <tr>
    <td colspan=4>
      <table border=0 width="100%">
        <tr><td colspan=4 class="dlabel"><% =Application("IDS_AdditionalFields") %></td></tr>
  <%
		For i = 0 to intFieldLimit
    		If (i+4) Mod 4 = 0 Then Response.Write("        <tr>" & vbCrLf)
			Response.Write(vbTab & vbTab & "<td class=""dfont""><input type=""checkbox"" name=""txtField" & i & """ id=""txtField" & i & """ value=""1""")
			If i = 0 Then Response.Write(" checked disabled") Else Response.Write(getDefault(1,arrField(i,1),"1"))
			Response.Write("><label for=""txtField" & i & """>" & arrField(i,0) & "</label></td>" & vbCrLf)
    		If (i+1) Mod 4 = 0 Then Response.Write("        </tr>" & vbCrLf)
    	Next
    	For i = 1 to 3 - intFieldLimit Mod 4
    		Response.Write(vbTab & vbTab & "<td class=""dfont"">&nbsp;</td>" & vbCrLf)
    		If (i+1) Mod 4 = 0 Then Response.Write("        </tr>" & vbCrLf)
    	Next
%>
      </table>
    </td>
  </tr>
  <tr><td colspan=4><hr></td></tr>
  <tr>
    <td><% =getLabel(Application("IDS_OrderBy"),"selOrder") %></td>
    <td colspan=3>
      <select name="selOrder" id="selOrder" class="oByte">

  <%
		For i = 0 to intFieldLimit
			Response.Write(vbTab & "        <option value=""" & i & """ " & getDefault(0,CStr(bytOrder),Cstr(i)) & ">" & arrField(i,0) & "</option>" & vbCrLf)
    	Next
%>
      </select>
      <select name="selOrderDir" id="selOrderDir" class="oText">
      	 <option value="ASC"<% =getDefault(0,"ASC",strOrderDir) %>>ASC</option>
      	 <option value="DESC"<% =getDefault(0,"DESC",strOrderDir) %>>DESC</option>
      </select>
    </td>
  </tr>
<%  End If %>
</form>
</table>

</div>


<div id="footerDiv" class="dvFooter">
<%
	Response.Write(getIcon("Javascript:confirmAction('gen');","V","view.gif",Application("IDS_ViewReport")))
	Response.Write(getIconNew("create_report.asp"))
	Response.Write(getIconCancel())
%>
</div>

<%
	End If

	Call DisplayFooter(1)
%>