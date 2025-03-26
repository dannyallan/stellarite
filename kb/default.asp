<!--#include file="..\_inc\functions.asp" -->
<!--#include file="..\_inc\functions_list.asp" -->
<!--#include file="..\_inc\sql\sql_kb.asp" -->
<%
	Call pageFunctions(8,1)

	Dim intDivSize      'as Integer
	Dim intLevel        'as Integer
	Dim intHalf         'as Integer
	Dim strCrumbs       'as String
	Dim arrCategories   'as Array
	Dim arrArticles     'as Array

	strTitle = getIDS("IDS_ModName8")
	intDivSize = CInt(intScreenH-270)

	lngRecordId = valNum(lngRecordId,2,0)

	Set objRS = objConn.Execute(getCategory(2,lngRecordId))
	If not (objRS.BOF and objRS.EOF) Then arrCategories = objRS.GetRows()

	intLevel = lngRecordId
	Do While CInt(intLevel) <> 0
		Set objRS = objConn.Execute(getCrumbs(intLevel))
		If (objRS.BOF and objRS.EOF) Then
			strCrumbs = getIDS("IDS_Deleted")
			intLevel = 0
		Else
			strCrumbs = " &raquo; <a href=""default.asp?id=" & intLevel & """>" & showString(objRS.fields(1).value) & "</a>" & strCrumbs
			intLevel = objRS.fields(0).value
		End If
	Loop
	strCrumbs = "<a href=""default.asp"">" & getIDS("IDS_Home") & "</a>" & strCrumbs

	Call DisplayHeader(1)
%>

<div id="contentDiv" class="dvBorder">

	<div style="float:left;width:75%;">
	  <table width="100%" border=0 cellspacing=0 cellpadding=2>
		<tr class="hRow">
		  <th class="hFont">&nbsp;<% =strCrumbs %></th>
		  <th align=right><%
	If blnRS and (blnAdmin or intPerm >= 5) Then
		Response.Write(getIconImport(2,getEditURL("C","?id=" & lngRecordId),getIDS("IDS_Category")))
	End If
%></th>
		  </tr>
		<tr>
		  <td>
			<div class="dvNoBorder" style="height:150px;">
			  <table width="100%" border=0 cellspacing=0 cellpadding=2>
				<tr><td colspan=2 class="dFont">&nbsp;</td></tr>
<%

	If IsArray(arrCategories) Then
		intHalf = Fix(UBound(arrCategories,2)/2)
		For i = 0 to intHalf
			Response.Write("  <tr><td width=""50%"">&nbsp;<a class=""tFont"" href=""default.asp?id=" & arrCategories(0,i) & """>" & trimString(arrCategories(1,i),20) & "</a> <span class=""dFont"">(" & arrCategories(5,i) & ")</span></td>")
			If i+1+intHalf <= UBound(arrCategories,2) Then
				Response.Write("<td width=""50%"">&nbsp;<a class=""tFont"" href=""default.asp?id=" & arrCategories(0,i+1+intHalf) & """>" & trimString(arrCategories(1,i+1+intHalf),20) & "</a> <span class=""dFont"">(" & arrCategories(5,i+intHalf) & ")</span></td></tr>" & vbCrLf)
			Else
				Response.Write("<td class=""dFont"">&nbsp;</td></tr>" & vbCrLf)
			End If
		Next
	End If
%>
			  </table>
			  <br />
			</div>
		  </td>
		</tr>
	  </table>
<%
	If not blnRS Then
		Call showList(21,intDivSize,12,2)
	Else
		Call showList(30,intDivSize,lngRecordId,2)
	End If
%>
	</div>

	<div style="margin-left:75%;padding-left:10px;">
	  <div id="div13" class="dvRightMenu" style="height:<% =CInt((intScreenH-70)*0.3) %>px;">
		<div class="hRow hScr hFont"><% =getIDS("IDS_UsefulLinks") %></div>
		<div>
			<ul>
<%
	If blnAdmin or intPerm >= 5 Then Response.Write("<li>" & getEditLink("C","?cat=" & lngRecordId,getIDS("IDS_CategoryNew")) & "</li>" & vbCrLf)
	If blnAdmin or intPerm >= 2 Then Response.Write("<li>" & getEditLink(8,"?cat=" & lngRecordId,getIDS("IDS_ArticleNew")) & "</li>" & vbCrLf)
	Response.Write("<div class=""hr""></div>" & vbCrLf)
	Response.Write("<li><a href=""search.asp"">" & getIDS("IDS_ArticleSearch") & "</a></li>" & vbCrLf)
%>
			</ul>
		</div>
	  </div>
	</div>
</div>

<%
	Call DisplayFooter(1)
%>