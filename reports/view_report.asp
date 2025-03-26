<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\reports.asp" -->
<!--#include file="..\_inc\sql\sql_repgen.asp" -->
<%
	Call pageFunctions(0,1)

	Dim strSql              'as String
	Dim bytType             'as Byte
	Dim intPermissions      'as Integer
	Dim strOwner            'as String
	Dim strParams           'as String
	Dim strFields           'as String
	Dim bytOrder            'as Byte
	Dim strOrderDir         'as String
	Dim strCreatedBy        'as String
	Dim datCreatedDate      'as Date
	Dim strModBy            'as String
	Dim datModDate          'as Date
	Dim blnChange           'as Boolean

	blnChange = valNum(Request.Form("hdnChange"),0,0)
	bytType = valNum(Request.QueryString("type"),1,0)
	If bytType = "" Then bytType = bytMod

	If blnChange = 1 Then

		intPerm = 5

		strTitle = valString(Request.Form("txtReportName"),100,1,0)
		strOwner = valString(Request.Form("txtOwner"),100,1,0)
		intPermissions = valNum(Request.Form("selPermissions"),1,1)
		bytOrder = valNum(Request.Form("selOrder"),1,0)
		strOrderDir = valString(Request.Form("selOrderDir"),4,0,0)

		If bytType <> 90 Then

			strParams = valString(Request.Form("hdnParams"),-1,0,0)
			strFields = Replace(valString(Request.Form("selFields"),-1,0,0)," ","")

			If not regExTest(strFields,"^[\d,]+$") Then Call logError(1,1)

			strSQL = genReport(bytType,strFields,strParams,bytOrder,strOrderDir,Application("av_MaxRecords"))

		Else
			bytMod = valNum(Request.Form("selModule"),1,0)

			strSQL = valString(Request.Form("txtSQL"),-1,0,0)
			strSQL = Replace(strSQL,Chr(145),Chr(39))
			strSQL = Replace(strSQL,Chr(146),Chr(39))
		End If

		If strDoAction = "edit" and intPerm >= 3 Then

			blnRS = True
			blnChange = 0
			objConn.Execute(updateReport(lngUserId,lngRecordId,strTitle,bytMod,bytType,intPermissions, _
					strOwner,strSQL,strFields,strParams,bytOrder,strOrderDir))

		ElseIf strDoAction = "new" and intPerm >= 2 Then

			blnRS = True
			blnChange = 0
			lngRecordId = insertReport(lngUserId,strTitle,bytMod,bytType,intPermissions, _
					strOwner,strSQL,strFields,strParams,bytOrder,strOrderDir)
		End If

	Elseif lngRecordId <> "" Then

		blnRS = True
		Set objRS = objConn.Execute(getReport(lngRecordId))

		If (objRS.BOF and objRS.EOF) Then
			Call doRedirect("default.asp")
		Else
			strTitle = showString(objRS.fields("R_Title").value)
			strSQL = objRS.fields("R_SQL").value

			If blnAdmin or objRS.fields("R_ModBy").value = lngUserId or objRS.fields("R_Owner").value = lngUserId Then intPerm = 5
			If objRS.fields("R_Type").value = 90 and not blnAdmin Then intPerm = 1
		End If
	End If

	If strDoAction = "del" and intPerm >= 5 Then
		objConn.Execute(delReport(lngUserId,lngRecordId))
		Call doRedirect("default.asp")
	End If

	strModName = getIDS("IDS_Reports")
	strTitle = showString(strTitle)
	strDir = Application("av_CRMDir") & "reports/"

	If strSQL <> "" Then

		Select Case strDoAction
			Case "xls","csv"

				strTitle = Replace(strTitle,";","")

				Response.Buffer = False
				Response.AddHeader "content-disposition", "attachment; filename=" & strTitle & "." & strDoAction
				Response.ContentType = "text/" & strDoAction

			Case "print"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<html lang="<% =strLanguage %>">
<head>
<title><% =strTitle %></title>
<link href="../common/css.asp" rel="stylesheet" type="text/css" />
<script language="JavaScript" type="text/javascript">
	function printWin() {
		window.print();
		history.back();
	}
</script>
</head>

<body onLoad="printWin();">
<span class="tFont"><% =strTitle %></span><br />
<span class="bFont"><% =showDate(0,Now) %></span><hr />

<%        	Case Else
				Call DisplayHeader(1)
%>
<script language="JavaScript" type="text/javascript">
	function chooseFormat() {
		var retVal = makeMsgBox();
		if (retVal == 6) {
			confirmAction('csv');
		}
		else if (retVal == 7) {
			confirmAction('xls');
		}
	}
</script>

<script language="VBScript" type="text/vbscript">
	function makeMsgBox()
		makeMsgBox = MsgBox("The default export format is CSV. " & _
				vbCrLf & "Choose Yes for the the default format " & _
				vbCrLf & "or choose No for XLS format.",35,"Export Format")
	end function
</script>

<div id="headerDiv" class="dvBorder">

<form name="frmReport" method="post" action="view_report.asp?id=<% =lngRecordId %>&m=<% =bytMod %>&type=<% =bytType %>">
<table border=0 cellspacing=0 cellpadding=0 width="100%">
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange",blnChange) %>
<%
				If blnChange = 1 Then
					Response.Write(getHidden("txtReportName",strTitle) & _
									getHidden("selPermissions",intPermissions) & _
									getHidden("txtOwner",showString(strOwner)))

					If bytType <> 90 Then
						Response.Write(getHidden("selFields",showString(strFields)) & _
									getHidden("hdnParams",showString(strParams)) & _
									getHidden("selOrder",bytOrder) & _
									getHidden("selOrderDir",showString(strOrderDir)))
					Else
						Response.Write(getHidden("txtSQL",showString(Request.Form("txtSQL"))) & _
									getHidden("selModule",bytMod))
					End If
				End If
%>
  <tr class="hRow">
	<td class="tFont"><img src="../images/report.gif" alt="<% =getIDS("IDS_Report") %>" width=32 height=32 hspace=10 align=absmiddle /><% =strTitle %></td>
	<td align=right>
<%
				Response.Write(getIconExport("Javascript:chooseFormat();"))
				Response.Write(getIconNew("edit_report.asp"))

				If blnChange = 1 Then
					Response.Write(getIconEdit("Javascript:document.forms[0].action = 'edit_report.asp?id=" & lngRecordId & "&m=" & bytMod & "&type=" & bytType & "';document.forms[0].submit();"))
				Elseif intPerm = 5 Then
					Response.Write(getIconEdit("edit_report.asp?id=" & lngRecordId & "&m=" & bytMod & "&type=" & bytType))
				End If

				If intPerm = 5 Then
					If blnChange = 1 Then
						Response.Write(getIconSave(strAction))
					Else
						Response.Write(getSpacer(1,28))
					End If
					If blnRS Then
						Response.Write(getIconDelete())
					End If
				End If

				Response.Write(getIconPrint("Javascript:confirmAction('print');"))
%>
   </td>
  </tr>
</table>
</form>

<br />
<%
		End Select

		strSQL = Replace(strSQL,"Current User",strFullName)

		Call execReport(strSQL,strDoAction)
		Call closeConn()

		If strDoAction <> "csv" Then Response.Write("</body></html>")
	Else
		Call DisplayHeader(1)

		Response.Write("<div id=""contentDiv"" class=""dvBorder"" style=""height:" & intScreenH-150 & "px;"">" & getIDS("IDS_ErrorUnspecified") & "</div>" & vbCrLf)

		Call DisplayFooter(1)
	End if
%>
