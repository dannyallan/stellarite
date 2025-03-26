<!--#include file="..\_inc\functions.asp" -->
<!--#include file="..\_inc\sql\sql_users.asp" -->
<%
	Call pageFunctions(0,1)

	Dim bytOrder        'as Byte
	Dim bytLastOrder    'as Byte
	Dim strSortOrder    'as String
	Dim strClass        'as String
	Dim blnLook         'as Boolean

	If Application("av_HideUsers") = "1" and lngRecordId <> lngUserId and not (getSecurity(0,5) or blnAdmin) Then Call logError(2,1)

	bytOrder = valNum(Request.QueryString("o"),1,0)
	bytLastOrder = valNum(Request.QueryString("lo"),1,0)
	strSortOrder = valString(Request.QueryString("so"),4,0,0)
	blnLook = True

	If bytOrder <> 2 and bytOrder <> 3 Then bytOrder = 1
	strSortOrder = getOrder(bytOrder,bytLastOrder,strSortOrder)

	Set objRS = objConn.Execute(getUsers(bytOrder,strSortOrder))
	If not (objRS.BOF and objRS.EOF) Then arrRS = objRS.GetRows()

	If isArray(arrRS) Then

		If blnRS Then

			i = 0
			Do while i <= UBound(arrRS,2) and blnLook
				If CLng(arrRS(0,i)) = lngRecordId Then
					blnLook = False
					strTitle = arrRS(1,i)
				Else
					lngPrevId = arrRS(0,i)
					i = i + 1
				End If
			Loop

			If blnLook Then
				Call doRedirect("profile.asp?o=" & bytOrder & "&so=" & strSortOrder)
			Else
				i = i + 1
				If i <= UBound(arrRS,2) Then lngNextId = arrRS(0,i)
				i = i - 1
			End If
		Else
			strTitle = getIDS("IDS_UserProfiles")
		End If
	End If

	strModName = strTitle

	Call DisplayHeader(1)
%>

<div id="headerDiv" class="dvBorder">

<table border=0 cellspacing=0 cellpadding=0 width="100%">
<form name="frmProfile" method="post" action="profile.asp">
  <tr class="hRow">
	<td class="tFont"><img src="../images/contact.gif" alt="<% =getIDS("IDS_UserProfiles") %>" width=32 height=32 hspace=10 align=absmiddle /><% =showString(strTitle) %></td>
	<td align=right>
<%
	If blnRS and blnAdmin Then
		Response.Write(getHidden("hdnAction",""))
	End If

	If blnRS and (Application("av_HideUsers") <> "1" or (blnAdmin or getSecurity(0,5))) Then
		If CLng(lngPrevId) <> 0 Then
			Response.Write(getIconPrev("profile.asp?id=" & lngPrevId & "&o=" & bytOrder & "&so=" & strSortOrder))
		End If

		If CLng(lngNextId) <> 0 Then
			Response.Write(getIconNext("profile.asp?id=" & lngNextId & "&o=" & bytOrder & "&so=" & strSortOrder))
		Else
			Response.Write(getSpacer(1,28))
		End If
	End If

	If blnAdmin Then
		Response.Write(getIconNew(getEditURL("U","")))
	End If

	If blnRS Then
		If blnAdmin or lngRecordId = lngUserId Then
			Response.Write(getIconEdit(getEditURL("U","?id=" & lngRecordId)))
		Else
			Response.Write(getSpacer(1,28))
		End If

		If blnAdmin Then
			Response.Write(getIconDelete())
		End If

		If blnAdmin or getSecurity(0,5) Then
			Response.Write(getIcon(getEditURL("P","?id=" & lngRecordId),"R","perm.gif",getIDS("IDS_Permissions")))
		End If
	End If

	If isArray(arrRS) and (Application("av_HideUsers") <> "1" or (blnAdmin or getSecurity(0,5))) Then

		Response.Write(getIconSearch(getSearchURL("?m=0")))

		If blnRS Then
			Response.Write(getIcon("profile.asp?o=" & bytOrder & "&so=" & strSortOrder,"V","change.gif",getIDS("IDS_ChangeView")))
		Else
			Response.Write(getIcon("profile.asp?id=" & arrRS(0,0) & "&o=" & bytOrder & "&so=" & strSortOrder,"V","change.gif",getIDS("IDS_ChangeView")))
		End If
	End If
%>
	</td>
  </tr>
</form>
</table>

<br />

<%
	If blnRS then

		If isArray(arrRS) Then
%>


</div>

<div id="modDiv" class="dvBorder">

<table border=0 cellspacing=0 cellpadding=0 width="100%">
  <tr>
	<td width="50%" valign=top>
	  <table border=0>
		<tr>
		  <td class="bFont" valign=top><% =getIDS("IDS_Address") %></td>
		  <td class="dFont">
		  <% Response.Write(showString(arrRS(2,i)))
			 If arrRS(3,i) <> "" Then Response.Write("<br />" & showString(arrRS(3,i)))
			 If arrRS(4,i) <> "" Then Response.Write("<br />" & showString(arrRS(4,i)))
			 Response.Write("<br />")
			 If arrRS(5,i) <> "" Then Response.Write(showString(arrRS(5,i)))
			 If arrRS(5,i) <> ""  and arrRS(6,i) <> "" Then Response.Write(", ")
			 If arrRS(6,i) <> "" Then Response.Write(showString(arrRS(6,i)))
			 If arrRS(5,i) <> "" or arrRS(6,i) <> "" Then Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;")
			 Response.Write(showString(arrRS(8,i)) & "<br /><br />")
		  %>
		  </td>
		</tr>
		<tr>
		  <td class="bFont"><% =getIDS("IDS_Email") %></td>
		  <td class="dFont"><% =showEmail(arrRS(9,i)) %></td>
		</tr>
		<tr>
		  <td class="bFont"><% =getIDS("IDS_Phone") %> 1</td>
		  <td class="dFont">
		  <% Response.Write(showPhone(arrRS(10,i)))
			 If arrRS(11,i) <> "" Then Response.Write("&nbsp;&nbsp;&nbsp;x." & arrRS(11,i))
		  %>
		  </td>
		</tr>
		<tr>
		  <td class="bFont"><% =getIDS("IDS_Phone") %> 2</td>
		  <td class="dFont">
		  <% Response.Write(showPhone(arrRS(12,i)))
			 If arrRS(13,i) <> "" Then Response.Write("&nbsp;&nbsp;&nbsp;x." & arrRS(13,i))
		  %>
		  </td>
		</tr>
	  </table>
	</td>
	<td width="50%" valign=top>
	  <table border=0>
		<tr>
		  <td class="bFont"><% =getIDS("IDS_UserName") %></td>
		  <td class="dFont"><% =showString(arrRS(14,i)) %></td>
		</tr>
		<tr>
		  <td class="bFont"><% =getIDS("IDS_Password") %></td>
		  <td class="dFont">********</td>
		</tr>
		<tr>
		  <td class="bFont"><% =getIDS("IDS_TimeZone") %></td>
		  <td class="dFont"><% =showString(arrRS(16,i)) %></td>
		</tr>
		<tr>
		  <td class="bFont"><% =getIDS("IDS_LastIP") %></td>
		  <td class="dFont"><% =showString(arrRS(17,i)) %></td>
		</tr>
		<tr>
		  <td class="bFont"><% =getIDS("IDS_LastAccess") %></td>
		  <td class="dFont"><% =showDate(1,arrRS(18,i)) %></td>
		</tr>
	  </table>
	</td>
  </tr>
</table>

<%        End If
	Else
%>

</div>

<div id="contentDiv" class="dvBorder" style="height:<% =intScreenH-125 %>px;">

<table border=0 cellspacing=0 cellpadding=0 width="100%">
  <thead>
  <tr class="hRow" style="position: relative; top:expression(this.offsetParent.scrollTop-2); left: -1;">
    <th class="hFont" width="1%">&nbsp;</th>
	<th class="hFont" width="22%"><a href="profile.asp?o=1&lo=<% =bytOrder %>&so=<% =strSortOrder %>"><% =getIDS("IDS_NameFull") %></a></th>
	<th class="hFont" width="33%"><a href="profile.asp?o=2&lo=<% =bytOrder %>&so=<% =strSortOrder %>"><% =getIDS("IDS_Email") %></a></th>
	<th class="hFont" width="34%"><a href="profile.asp?o=3&lo=<% =bytOrder %>&so=<% =strSortOrder %>"><% =getIDS("IDS_LastAccess") %></a></th>
	<th class="hFont" width="10%">&nbsp;</td>
  </tr>
  </thead>
  <tbody style="max-height: <% =intScreenH-140 %>px; overflow: auto;">
<%
		If isArray(arrRS) Then
			For i = 0 to UBound(arrRS,2)
				strClass = toggleRowColor(strClass)
%>
  <tr class="<%=strClass%>">
	<td class="dFont" width="1%">&nbsp;</td>
	<td class="dFont" width="22%"><a href="profile.asp?id=<% =arrRS(0,i) %>&o=<% =bytOrder %>&so=<% =strSortOrder %>"><% =showString(arrRS(1,i)) %></a></td>
	<td class="dFont" width="32%"><% =showEmail(arrRS(9,i)) %></td>
	<td class="dFont" width="32%"><% =showDate(1,arrRS(18,i)) %></td>
	<td class="dFont" width="10%" align=right>&nbsp;
<%
				If blnAdmin or CLng(arrRS(0,i)) = lngUserId Then
					Response.Write(vbTab & getIconImport(2,getEditURL("U","?id=" & arrRS(0,i)),showString(arrRS(1,i))) & vbCrLf)
				End If
				If blnAdmin or getSecurity(0,5) Then
					Response.Write(vbTab & getIconImport(5,getEditURL("P","?id=" & arrRS(0,i)),showString(arrRS(1,i))) & vbCrLf)
				End If
%>
	</td>
  </tr>
<%            Next
		End If
%>
  </tbody>
</table>
</div>
<%
	End If

	Call DisplayFooter(1)
%>
