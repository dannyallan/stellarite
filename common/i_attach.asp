<!--#include file="..\_inc\functions.asp" -->
<!--#include file="..\_inc\sql\sql_attachments.asp" -->
<%
	Call pageFunctions(0,1)

	Dim lngEventId      'as Long
	Dim bytOrder        'as Byte
	Dim bytLastOrder    'as Byte
	Dim strSortOrder    'as String
	Dim strClass        'as String
	Dim blnLook         'as Boolean
	Dim arrLinks        'as Array

	strTitle = getIDS("IDS_Attachments")
	lngEventId = valNum(Request.QueryString("eid"),3,0)
	bytOrder = valNum(Request.QueryString("o"),1,0)
	bytLastOrder = valNum(Request.QueryString("lo"),1,0)
	strSortOrder = valString(Request.QueryString("so"),4,0,0)
	blnLook = True

	If bytMod = "" or lngModId = "" Then Call logError(3,1)
	If bytOrder <> 2 and bytOrder <> 3 and bytOrder <> 4 Then bytOrder = 1
	strSortOrder = getOrder(bytOrder,bytLastOrder,strSortOrder)

	If strDoAction = "del" and intPerm >= 4 Then
		Call delAttach(lngUserId,lngRecordId,bytMod,lngModId)
	End If

	Set objRS = objConn.Execute(getAttachBy(bytMod,lngModId,lngEventId,intMember,bytOrder,strSortOrder))
	If not (objRS.BOF and objRS.EOF) Then arrRS = objRS.GetRows()

	If blnRS Then
		Set objRS = objConn.Execute(getAttachLinks(lngRecordId))
		If not (objRS.BOF and objRS.EOF) Then arrLinks = objRS.GetRows()
	End If

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
				Call doRedirect("i_attach.asp?m=" & bytMod & "&mid=" & lngModId & "&eid=" & lngEventId & "&o=" & bytOrder & "&so=" & strSortOrder)
			Else
				i = i + 1
				If i <= UBound(arrRS,2) Then lngNextId = CLng(arrRS(0,i))
				i = i - 1
			End If
		End If
	End If

	Call DisplayHeader(2)

%>

<div id="headerDiv" class="dvNoBorder">

<form name="frmAttach" method="post" action="i_attach.asp?id=<% =lngRecordId %>&m=<% =bytMod %>&mid=<% =lngModId %>&eid=<% =lngEventId %>&o=<% =bytOrder %>&so=<% =strSortOrder %>">
<table border=0 cellspacing=3 width="100%">
  <tr>
	<td>
<%
	If intPerm >= 2 Then
		Response.Write(getIconNew(getEditURL("A","?m=" & bytMod & "&mid=" & lngModId & "&eid=" & lngEventId)))
	End If

	If blnRS Then
		If intPerm >= 4 Then
			Response.Write(getIconDelete() & vbTab & getHidden("hdnAction","") & vbCrLf)
		End If

		If CLng(lngPrevId) <> 0 Then
			Response.Write(getIconPrev("i_attach.asp?id=" & lngPrevId & "&m=" & bytMod & "&mid=" & lngModId & "&eid=" & lngEventId & "&o=" & bytOrder & "&so=" & strSortOrder))
		Else
			Response.Write(getSpacer(1,28))
		End If

		If CLng(lngNextId) <> 0 Then
			Response.Write(getIconNext("i_attach.asp?id=" & lngNextId & "&m=" & bytMod & "&mid=" & lngModId & "&eid=" & lngEventId & "&o=" & bytOrder & "&so=" & strSortOrder))
		End If
	End If

	Response.Write(vbTab & "</td>" & vbCrLf & vbTab & "<td align=right>" & vbCrLf)

	If isArray(arrRS) Then
		If blnRS Then
			Response.Write(getIcon("i_attach.asp?m=" & bytMod & "&mid=" & lngModId & "&eid=" & lngEventId & "&o=" & bytOrder & "&so=" & strSortOrder,"V","change.gif",getIDS("IDS_ChangeView")))
		Else
			Response.Write(getIcon("i_attach.asp?id=" & arrRS(0,0) & "&m=" & bytMod & "&mid=" & lngModId & "&eid=" & lngEventId & "&o=" & bytOrder & "&so=" & strSortOrder,"V","change.gif",getIDS("IDS_ChangeView")))
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
			arrRS(1,i) = trimString(arrRS(1,i),30)
			If arrRS(1,i) = "" Then arrRS(1,i) = getIDS("IDS_Untitled")
%>
<table border=0 cellspacing=0 cellpadding=0 width="100%">
  <tr class="hRow">
	<td class="tFont"><% =showString(arrRS(1,i)) %></td>
	<td valign=top align=right>
	  <table border=0 cellspacing=0 cellpadding=0>
		<tr>
		  <td class="dFont"><% =getIDS("IDS_Created") %>:</td>
		  <td class="dFont">&nbsp;&nbsp;<% =showDate(0,arrRS(4,i)) %></td>
		  <td class="dFont rIndent">&nbsp;&nbsp;<% =showString(arrRS(5,i)) %></td>
		</tr>
		<tr>
		  <td class="dFont"><% =getIDS("IDS_Modified") %>:</td>
		  <td class="dFont">&nbsp;&nbsp;<% =showDate(0,arrRS(6,i)) %></td>
		  <td class="dFont rIndent">&nbsp;&nbsp;<% =showString(arrRS(7,i)) %></td>
		</tr>
	  </table>
	</td>
  </tr>
</table>

</div>

<div id="contentDiv" class="dvNoBorder" style="height:<% =intScreenH-288 %>px;">
<br />
<table border=0 cellpadding=5>
  <tr>
	<td valign=top class="bFont"><% =getIDS("IDS_Type") %></td>
	<td class="dFont"><% =showString(arrRS(2,i)) %></td>
  </tr>
  <tr>
	<td valign=top class="bFont"><% =getIDS("IDS_Description") %></td>
	<td class="dFont"><% =showParagraph(arrRS(3,i)) %></td>
  </tr>
  <tr>
	<td valign=top class="bFont"><% =getIDS("IDS_Attachments") %></td>
	<td class="dFont"><ul>
<%

	If isArray(arrLinks) Then
		For i = 0 to UBound(arrLinks,2)
			Response.Write("<li><a href=""" & showString(arrLinks(0,i)) & """ target=""_top"">" & showString(arrLinks(0,i)) & "</a></li>" & vbCrLf)
		Next
	End If
%>
	</ul></td>
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
	<th class="hFont lIndent" width="18%"><a href="i_attach.asp?m=<% =bytMod %>&mid=<% =lngModId %>&eid=<% =lngEventId %>&o=1&lo=<% =bytOrder %>&so=<% =strSortOrder %>"><% =getIDS("IDS_Modified") %></a></th>
	<th class="hFont" width="23%"><a href="i_attach.asp?m=<% =bytMod %>&mid=<% =lngModId %>&eid=<% =lngEventId %>&o=2&lo=<% =bytOrder %>&so=<% =strSortOrder %>"><% =getIDS("IDS_Type") %></a></th>
	<th class="hFont" width="35%"><a href="i_attach.asp?m=<% =bytMod %>&mid=<% =lngModId %>&eid=<% =lngEventId %>&o=3&lo=<% =bytOrder %>&so=<% =strSortOrder %>"><% =getIDS("IDS_Title") %></a></th>
	<th class="hFont" width="22%"><a href="i_attach.asp?m=<% =bytMod %>&mid=<% =lngModId %>&eid=<% =lngEventId %>&o=4&lo=<% =bytOrder %>&so=<% =strSortOrder %>"><% =getIDS("IDS_Author") %></a></th>
  </tr>
</table>

</div>

<div id="contentDiv" class="dvNoBorder" style="height:<% =intScreenH-276 %>px;">
<table border=0 cellspacing=0 cellpadding=0 width="100%">
<%
		If isArray(arrRS) Then
			For i = 0 to UBound(arrRS,2)
				strClass = toggleRowColor(strClass)

				arrRS(1,i) = trimString(arrRS(1,i),30)
				If arrRS(1,i) = "" Then arrRS(1,i) = getIDS("IDS_Untitled")
				If IsNull(arrRS(2,i)) or arrRS(2,i) = "" Then arrRS(2,i) = "NULL"
%>
  <tr class="<%=strClass%>">
	<td class="dFont lIndent" width="18%"><% =showDate(0,arrRS(6,i)) %></td>
	<td class="dFont" width="23%"><a href="i_attach.asp?id=<% =arrRS(0,i) %>&m=<% =bytMod %>&mid=<% =lngModId %>&eid=<% =lngEventId %>&o=<% =bytOrder %>&so=<% =strSortOrder %>"><% =showString(arrRS(2,i)) %></a></td>
	<td class="dFont" width="35%"><% =showString(arrRS(1,i)) %></td>
	<td class="dFont rIndent" width="22%"><% =showString(arrRS(7,i)) %></td>
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