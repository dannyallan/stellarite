<!--#include file="..\_inc\functions.asp" -->
<!--#include file="..\_inc\sql\sql_notes.asp" -->
<%
	Call pageFunctions(0,1)

	Dim lngEventId      'as Long
	Dim bytOrder        'as Byte
	Dim bytLastOrder    'as Byte
	Dim strSortOrder    'as String
	Dim strClass        'as String
	Dim blnLook         'as Boolean

	strTitle = getIDS("IDS_Notes")
	lngEventId = valNum(Request.QueryString("eid"),3,0)
	bytOrder = valNum(Request.QueryString("o"),1,0)
	bytLastOrder = valNum(Request.QueryString("lo"),1,0)
	strSortOrder = valString(Request.QueryString("so"),4,0,0)
	blnLook = True

	If bytMod = "" or lngModId = "" Then Call logError(3,1)
	If bytOrder <> 2 and bytOrder <> 3 Then bytOrder = 1
	strSortOrder = getOrder(bytOrder,bytLastOrder,strSortOrder)

	If strDoAction = "del" and intPerm >= 4 Then Call delNote(lngUserId,lngRecordId,bytMod,lngModId)

	Set objRS = objConn.Execute(getNotesBy(bytMod,lngModId,lngEventId,intMember,bytOrder,strSortOrder))
	If not (objRS.BOF and objRS.EOF) Then arrRS = objRS.GetRows()

	If isArray(arrRS) Then

		If blnRS Then
			i = 0
			Do while i <= UBound(arrRS,2) and blnLook
				If CLng(arrRS(0,i)) = lngRecordId Then
					blnLook = False
				Else
					lngPrevId = CLng(arrRS(0,i))
					i = i + 1
				End If
			Loop

			If blnLook Then
				Call doRedirect("i_notes.asp?m=" & bytMod & "&mid=" & lngModId & "&eid=" & lngEventId & "&o=" & bytOrder & "&so=" & strSortOrder)
			Else
				i = i + 1
				If i <= UBound(arrRS,2) Then lngNextId = CLng(arrRS(0,i))
				i = i - 1
			End If
		End If
	End If

	If blnRS and Instr(Request.QueryString,"export") > 0 Then
		Response.Buffer = False
		Response.AddHeader "content-disposition", "attachment; filename=" & arrRS(2,i) & ".doc"
		Response.ContentType = "application/msword"

		Response.Write("<html><body style=""font-size:88%;""><h1>" & showString(arrRS(2,i)) & "</h1>" & _
						"<b>" & getIDS("IDS_Created") & ": " & vbTab & showDate(0,arrRS(6,i)) & vbTab & showString(arrRS(5,i)) & "</b><br />" & _
						"<b>" & getIDS("IDS_Modified") & ": " & vbTab & showDate(0,arrRS(8,i)) & vbTab & showString(arrRS(7,i)) & "</b><br /><br />" & _
						"<p>" & showHTML(arrRS(3,i)) & "</p></body></html>")

		Call endResponse()
	End If

	Call DisplayHeader(2)

%>

<div id="headerDiv" class="dvNoBorder">

<form name="frmNote" method="post" action="i_notes.asp?id=<% =lngRecordId %>&m=<% =bytMod %>&mid=<% =lngModId %>&eid=<% =lngEventId %>&o=<% =bytOrder %>&so=<% =strSortOrder %>">
<table border=0 cellspacing=3 width="100%">
  <tr>
	<td>
<%
	If intPerm >= 2 Then
		Response.Write(getIconNew(getEditURL("N","?m=" & bytMod & "&mid=" & lngModId & "&eid=" & lngEventId)))
	End If

	If blnRS Then
		If intPerm >= 3 Then
			Response.Write(getIconEdit(getEditURL("N","?id=" & lngRecordId & "&m=" & bytMod & "&mid=" & lngModId & "&eid=" & lngEventId)))
		End If

		Response.Write(getIconExport("i_notes.asp?id=" & lngRecordId & "&m=" & bytMod & "&mid=" & lngModId & "&eid=" & lngEventId & "&export"))

		If intPerm >= 4 Then
			Response.Write(getIconDelete() & vbTab & getHidden("hdnAction",""))
		End If

		If CLng(lngPrevId) <> 0 Then
			Response.Write(getIconPrev("i_notes.asp?id=" & lngPrevId & "&m=" & bytMod & "&mid=" & lngModId & "&eid=" & lngEventId & "&o=" & bytOrder & "&so=" & strSortOrder))
		Else
			Response.Write(getSpacer(1,28))
		End If

		If CLng(lngNextId) <> 0 Then
			Response.Write(getIconNext("i_notes.asp?id=" & lngNextId & "&m=" & bytMod & "&mid=" & lngModId & "&eid=" & lngEventId & "&o=" & bytOrder & "&so=" & strSortOrder))
		End If
	End If

	Response.Write(vbTab & "</td>" & vbCrLf & vbTab & "<td align=right>" & vbCrLf)

	If isArray(arrRS) Then
		If blnRS Then
			Response.Write(getIcon("i_notes.asp?m=" & bytMod & "&mid=" & lngModId & "&eid=" & lngEventId & "&o=" & bytOrder & "&so=" & strSortOrder,"V","change.gif",getIDS("IDS_ChangeView")))
		Else
			Response.Write(getIcon("i_notes.asp?id=" & arrRS(0,0) & "&m=" & bytMod & "&mid=" & lngModId & "&eid=" & lngEventId & "&o=" & bytOrder & "&so=" & strSortOrder,"V","change.gif",getIDS("IDS_ChangeView")))
		End If
	End If
%>
	</td>
  </tr>
</table>
</form>
<%
	If blnRS then
		If isArray(arrRS) then
%>
<table border=0 cellspacing=0 cellpadding=0 width="100%">
  <tr class=hRow>
	<td class="tFont"><% =showString(arrRS(2,i)) %></td>
	<td valign=top align=right>
	  <table border=0 cellspacing=0 cellpadding=0>
		<tr>
		  <td class="dFont"><% =getIDS("IDS_Created") %>:</td>
		  <td class="dFont">&nbsp;&nbsp;<% =showDate(0,arrRS(6,i)) %></td>
		  <td class="dFont rIndent">&nbsp;&nbsp;<% =showString(arrRS(5,i)) %></td>
		</tr>
		<tr>
		  <td class="dFont"><% =getIDS("IDS_Modified") %>:</td>
		  <td class="dFont">&nbsp;&nbsp;<% =showDate(0,arrRS(8,i)) %></td>
		  <td class="dFont rIndent">&nbsp;&nbsp;<% =showString(arrRS(7,i)) %></td>
		</tr>
	  </table>
	</td>
  </tr>
</table>

</div>

<div id="contentDiv" class="dvNoBorder" style="height:<% =intScreenH-288 %>px;">
<table border=0 cellpadding=5>
  <tr>
	<td class="dFont">
	<br /><% =showHTML(arrRS(3,i)) %>
	</td>
  </tr>
</table>
</div>
<%        Else     %>
<table border=0 cellspacing=0 cellpadding=0 width="100%">
  <tr class="hRow">
	<td class="tFont"><% =getIDS("IDS_Deleted") %></td>
  </tr>
</table>
<%        End If
	Else
%>
<table border=0 cellspacing=0 cellpadding=0 width="100%">
  <tr class="hRow">
	<th class="hFont lIndent" width="22%"><a href="i_notes.asp?m=<% =bytMod %>&mid=<% =lngModId %>&eid=<% =lngEventId %>&o=1&lo=<% =bytOrder %>&so=<% =strSortOrder %>"><% =getIDS("IDS_Modified") %></a></th>
	<th class="hFont" width="33%"><a href="i_notes.asp?m=<% =bytMod %>&mid=<% =lngModId %>&eid=<% =lngEventId %>&o=2&lo=<% =bytOrder %>&so=<% =strSortOrder %>"><% =getIDS("IDS_Type") %></a></th>
	<th class="hFont" width="34%"><a href="i_notes.asp?m=<% =bytMod %>&mid=<% =lngModId %>&eid=<% =lngEventId %>&o=3&lo=<% =bytOrder %>&so=<% =strSortOrder %>"><% =getIDS("IDS_Author") %></a></th>
	<th class="hFont" width="10%">&nbsp;</td>
  </tr>
</table>

</div>

<div id="contentDiv" class="dvNoBorder" style="height:<% =intScreenH-276 %>px;">
<table border=0 cellspacing=0 cellpadding=0 width="100%">
<%
		If isArray(arrRS) then
			For i = 0 to UBound(arrRS,2)
				strClass = toggleRowColor(strClass)
				If IsNull(arrRS(2,i)) or arrRS(2,i) = "" Then arrRS(2,i) = "NULL"
%>
  <tr class="<%=strClass%>">
	<td class="dFont lIndent" width="22%"><% =showDate(0,arrRS(8,i)) %></td>
	<td class="dFont" width="33%">
	<a href="i_notes.asp?id=<% =arrRS(0,i) %>&m=<% =bytMod %>&mid=<% =lngModId %>&eid=<% =lngEventId %>&o=<% =bytOrder %>&so=<% =strSortOrder %>"><% =showString(arrRS(2,i)) %></a>
	</td>
	<td class="dFont" width="34%"><% =showString(arrRS(7,i)) %></td>
	<td class="dFont rIndent" width="10%" align=right>&nbsp;
<%
				If intPerm >= 3 Then
					Response.Write(vbTab & getIconImport(2,getEditURL("N","?id=" & arrRS(0,i) & "&m=" & bytMod & "&mid=" & lngModId & "&eid=" & lngEventId),showString(arrRS(2,i))) & vbCrLf)
				End If
%>
	</td>
  </tr>
<%            Next
		End If
%>
</table>
</div>
<%
	End If


	Call DisplayFooter(2)
%>

