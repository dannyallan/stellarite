<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_kb.asp" -->
<%
	Call pageFunctions(8,1)

	Dim strMsg          'as String
	Dim strTerms        'as String
	Dim lngCategory     'as Integer
	Dim strCategory     'as String
	Dim bytUsing        'as Byte
	Dim bytWhere        'as Byte
	Dim strWhere        'as String

	strTitle = getIDS("IDS_ArticleSearch")
	strMsg = getIDS("IDS_MsgEnterSearch")
	strTerms = valString(Request.QueryString("txtSearch"),255,0,0)

	If strTerms <> "" Then
		lngCategory = valNum(Request.QueryString("selCategory"),3,0)
		bytUsing = valNum(Request.QueryString("selUsing"),1,1)
		bytWhere = valNum(Request.QueryString("rdoWhere"),1,1)

		Set objRS = objConn.Execute(getArticleSearch(strTerms,lngCategory,bytUsing,bytWhere,Application("av_MaxRecords")))
		If not (objRS.BOF and objRS.EOF) Then
			arrRS = objRS.GetRows()
		Else
			bytWhere = 0
			strCategory = getCategoryDropDown(300,True,"selCategory",lngCategory)
			strMsg = getIDS("IDS_MsgNoResults")
		End If

	Else
		bytWhere = 0
		strCategory = getCategoryDropDown(300,True,"selCategory",lngCategory)
	End If

	If strTerms = "" Then strTerms = "*" Else strTerms = showString(strTerms)

	Call DisplayHeader(1)
%>

<div id="headerDiv" class="dvBorder">

<table border=0 cellspacing=0 cellpadding=0 width="100%">
  <tr class="hRow">
	<td class="tFont"><img src="../images/find.gif" alt="<% =getIDS("IDS_Search") %>" width=32 height=32 hspace=10 align=absmiddle /><% =getIDS("IDS_ArticleSearch") %></td>
  </tr>
</table>

</div>

<div id="contentDiv" class="dvBorder" style="height:<% =intScreenH-90 %>px;">

<%    If not isArray(arrRS) Then    %>

<form name="frmSearch" method="get" action="search.asp">
<table border=0 cellspacing=0 cellpadding=5 width="100%">
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
  <tr>
	<td colspan=2>&nbsp;</td>
  </tr>
  <tr>
	<td class="dFont" colspan=2><% =strMsg %></td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_SearchFor"),"txtSearch") %></td>
	<td><% =getTextField("txtSearch","mText",strTerms,50,255,"") %></td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_SearchGroup"),"selCategory") %></td>
	<td><% =strCategory %></td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_SearchUsing"),"selUsing") %></td>
	<td>
	  <select name="selUsing" id="selUsing" class="oByte" style="width:300px;">
		<option value="0"<% =getDefault(0,0,bytUsing) & ">" & getIDS("IDS_TermsAny") %></option>
		<option value="1"<% =getDefault(0,1,bytUsing) & ">" & getIDS("IDS_TermsAll") %></option>
		<option value="2"<% =getDefault(0,2,bytUsing) & ">" & getIDS("IDS_TermsExact") %></option>
	  </select>
	</td>
  </tr>
  <tr>
	<td class="bFont" valign=top><% =getIDS("IDS_SearchWhere") %></td>
	<td class="dFont">
	  <% =getRadio("rdoWhere",0,bytWhere,"") %><% =getLabel(getIDS("IDS_SearchTitles"),"rdoWhere0") %><br />
	  <% =getRadio("rdoWhere",1,bytWhere,"") %><% =getLabel(getIDS("IDS_SearchFullText"),"rdoWhere1") %><br />
	  <% =getRadio("rdoWhere",2,bytWhere,"") %><% =getLabel(getIDS("IDS_SearchID"),"rdoWhere2") %><br />
	</td>
  </tr>
  <tr>
	<td>&nbsp;</td>
	<td><% =getSubmit("btnSubmit",getIDS("IDS_Search"),100,"S","") %></td>
  </tr>
</table>
</form>

<%
	Else
		Response.Write("<br /><table border=0 cellspacing=0 cellpadding=2 width=""100%"">" & vbCrLf)
		For i = 0 to UBound(arrRS,2)
%>
  <tr class="dRow2">
	<td class="hFont" width="75%"><a href="article.asp?id=<% =arrRS(0,i) %>"><% =trimString(arrRS(1,i),40) %></a></td>
	<td class="dFont" width="25%"><span class="hFont"><% =getIDS("IDS_Updated") %></span> <% =showDate(0,arrRS(2,i)) %></td>
  </tr>
  <tr class="drow">
	<td class="dFont"><% =getIDS("IDS_KBHits") & " " & arrRS(5,i) %></td>
	<td class="dFont">
<%
	Response.Write(getIDS("IDS_KBReviews") & " " & bigDigitNum(3,arrRS(6,i)) & " &nbsp;&nbsp;&nbsp; ")

	If IsNumeric(arrRS(6,i)) and arrRS(6,i) > 0 Then
		Response.Write(getIDS("IDS_KBRating") & " " & FormatNumber(CInt(arrRS(7,i))/CInt(arrRS(6,i)),2))
	Else
		Response.Write(getIDS("IDS_KBRating") & " 0.00")
	End If
%></td>
  </tr>
  <tr class="dRow2">
	<td colspan=2><% =getSpacer(1,1) %></td>
  </tr>
  <tr class="drow">
	<td class="dFont" style="padding-left:15px;" colspan=2>&nbsp;<% =trimString(arrRS(3,i),250) %><p></td>
  </tr>
<%
		Next
		Response.Write("</table>" & vbCrLf)

	End If

	Response.Write("</div>" & vbCrLf & vbCrLf)

	Call DisplayFooter(1)
%>