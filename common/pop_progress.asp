<%    @EnableSessionState = False         %>
<!--    #INCLUDE FILE="..\_inc\upload.asp"    -->
<%
	Server.ScriptTimeout = 10

	Const RefreshTime = 1

	Dim Form
	Set Form = New ASPForm

	On Error Resume Next
	Set Form = Form.getForm(Request.QueryString("uid"))

	If Err = 0 Then

		On Error GoTo 0

		If Form.BytesRead > 0 Then
			Dim UpStateHTML
			UpStateHTML = FileStateInfo(Form)
		End If

		Response.CacheControl = "no-cache"
		Response.AddHeader "Pragma","no-cache"
		Response.AddHeader "Refresh", RefreshTime

		Dim PercBytesRead, PercentRest, BytesRead, TotalBytes
		Dim UploadTime, RestTime, TransferRate
		BytesRead = Form.BytesRead
		TotalBytes = Form.TotalBytes

		If TotalBytes > 0 Then
			PercBytesRead = int(100*BytesRead/TotalBytes)
			PercentRest = 100-PercBytesRead

			If Form.ReadTime > 0 Then TransferRate = BytesRead / Form.ReadTime
			If TransferRate > 0 Then RestTime = FormatTime((TotalBytes-BytesRead) / TransferRate)
			TransferRate = FormatSize(1000 * TransferRate)
		Else
			RestTime = "?"
			PercBytesRead = 0
			PercentRest = 100
			TransferRate = "?"
		End If

		Dim TDsread, TDsRemain
		TDsread = replace(space(0.5*PercBytesRead), " ", "<td bgcolor=""" & Application("av_MajorColor") & """>&nbsp;</td>")
		TDsRemain = replace(space(0.5*PercentRest), " ", "<td>&nbsp;</td>")

		UploadTime = FormatTime(Form.ReadTime)
		TotalBytes = FormatSize(TotalBytes)
		BytesRead = FormatSize(BytesRead)

		Function FormatTime(byval ms)
			ms = 0.001 * ms
			FormatTime = (ms \ 60) & ":" & right("0" & (ms mod 60),2) & "s"
		End Function

		Function FormatSize(byval Number)
			If isnumeric(Number) Then
				If Number > &H100000 Then'1M
					Number = FormatNumber (Number/&H100000,1) & "MB"
				Elseif Number > 1024 Then'1k
					Number = FormatNumber (Number/1024,1) & "kB"
				Else
					Number = FormatNumber (Number,0) & "B"
				End If
			End If
			FormatSize = Number
		End Function

		Function FileStateInfo(Form)
			On Error Resume Next
			Dim UpStateHTML, Field
			For Each Field in Form.Files
				UpStateHTML = UpStateHTML & "FieldName:" & Field.Name

				If Field.InProgress Then
					UpStateHTML = UpStateHTML & ", uploading: " & Field.FileName
				Elseif Field.Length > 0 Then
					UpStateHTML = UpStateHTML & ", received: " & Field.FileName & ", " & FormatSize(Field.Length)
				End If

				UpStateHTML = UpStateHTML & "<br />"
			Next
			FileStateInfo = UpStateHTML
		End Function
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<html lang="en">
<head>
 <title><% =PercBytesRead & "% " & getIDS("IDS_Complete") %></title>
 <link href="css.asp" rel="stylesheet" type="text/css" />
 <meta http-equiv="Page-Enter" content="revealTrans(Duration=0,Transition=6)" />
 <meta http-equiv="Refresh" content="<% =RefreshTime %>" />
</head>

<body bgcolor="<% =Application("av_MinorColor") %>" leftmargin=15 topmargin=4 rightmargin=4 bottommargin=4>

<br /><br />

<table cellpadding=0 cellspacing=0 border=0 width=100% >
 <tr>
	<td class="dFont"><span class="bFont">Uploading:</span> <% =TotalBytes %> to <% =Request.ServerVariables("HTTP_HOST") %> ...<br /></td>
  </tr>
</table>

<table cellpadding=0 cellspacing=0 border=0 width=100% >
  <tr>
	<% =TDsread %><% =TDsRemain %>
  </tr>
</table>

<table cellpadding=0 cellspacing=0 border=0>
  <tr>
	<td class="bFont">Progress</td>
	<td class="dFont">: <% =BytesRead %> of <%= TotalBytes %> (<% =PercBytesRead %>%) </td>
  </tr>
  <tr>
	<td class="bFont">Time </td>
	<td class="dFont">: <% =UploadTime %> (<% =TransferRate %>/s) </td>
  </tr>
  <tr>
	<td class="bFont">Time left</td>
	<td class="dFont">: <% =RestTime %> </td>
  </tr>
</table>

<br /><center>
<input type="button" value="<% =getIDS("IDS_Cancel") %>" OnClick="doCancel()" style="background-color:<% =Application("av_MajorColor") %>;color:white;cursor:hand;font-weight:bold" />
</center><br />

<% =UpStateHTML %>

<script language="JavaScript" type="text/javascript">
window.setTimeout('doRefresh()', <% =RefreshTime %>*2000);

function doRefresh() {
	window.location.href = window.location.href;
	window.setTimeout('refresh()', <% =RefreshTime %>*2000);
}
function doCancel() {
	var l = ''+opener.document.location;

	if (l.indexOf('Action=Cancel')<0) {
		l += (l.indexOf('?')<0 ? '?' : '&') + 'Action=Cancel'
	};

	opener.document.location = l;
	window.close();
}
</script>

</body>
</html>

<%    Else    %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<html lang="en">
 <head>
 <title><% =getIDS("IDS_Complete") %></title>
 <script language="JavaScript" type="text/javascript">window.close();</script>
 </head>
</html>
<%    End If    %>