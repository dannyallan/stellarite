<!--#include file="..\_inc\functions.asp" -->
<!--#include file="..\_inc\sql\sql_kb.asp" -->
<%
	Call pageFunctions(8,1)

	Dim intLevel        'as Integer
	Dim strCrumbs       'as String
	Dim strKeywords     'as String
	Dim lngViews        'as Long
	Dim lngRateCount    'as Long
	Dim lngRateTotal    'as Long
	Dim strSummary      'as String
	Dim strInfo         'as String
	Dim strLink         'as String
	Dim strType         'as String
	Dim datExpires      'as Date
	Dim datUpdated      'as Date
	Dim blnPrint        'as Boolean
	Dim blnExport        'as Boolean
	Dim bytRating       'as Byte

	lngRecordId = valNum(lngRecordId,3,1)

	If valString(Request.QueryString("print"),4,0,0) = "true" Then blnPrint = True Else blnPrint = False
	If valString(Request.QueryString("export"),4,0,0) = "true" Then blnExport = True Else blnExport = False

	If Request.Form.Count > 0 Then
		bytRating = valNum(Request.Form("selRating"),1,1)
		objConn.Execute(rateArticle(lngRecordId,bytRating))
	End If

	Set objRS = objConn.Execute(getArticle(1,lngRecordId))
	If not (objRS.BOF and objRS.EOF) then

		intLevel = objRS.fields("CatId").value
		strTitle = showString(objRS.fields("H_Title").value)
		strKeywords = showString(objRS.fields("H_Keywords").value)
		lngViews = objRS.fields("H_Views").value
		lngRateCount = objRS.fields("H_RateCount").value
		lngRateTotal = objRS.fields("H_RateTotal").value
		strSummary = objRS.fields("H_Summary").value
		strInfo = objRS.fields("H_Info").value
		strLink = objRS.fields("H_Link").value
		datExpires = showDate(0,objRS.fields("H_Expire").value)
		datUpdated = showDate(0,objRS.fields("H_ModDate").value)

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
	Else
		strTitle = getIDS("IDS_Deleted")
	End If

	If strLink <> "" Then
		strInfo = showLink(0,strLink,strLink)
		strType = getIDS("IDS_Link")
	Else
		strInfo = showHTML(strInfo)
		strType = getIDS("IDS_Description")
	End If

	If blnPrint or blnExport Then

		If blnExport Then
			Response.Buffer = False
			Response.AddHeader "content-disposition", "attachment; filename=" & strTitle & ".doc"
			Response.ContentType = "application/msword"
		End If

		Call closeConn()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<html lang="<% =strLanguage %>">
<head>
<title><% =strTitle %></title>
<%
		If blnPrint Then
%>
<script language="JavaScript" type="text/javascript">
	function printWin() {
		window.print();
		history.back();
	}
</script>
</head>

<body onLoad='printWin()'>
<%
		Else
%>
</head>

<body style="font-size: 88%">
<%
		End If
%>
<h1><% =strTitle %></h1>
<hr />
<h3><% =getIDS("IDS_Summary") %></h3><p><% =showParagraph(strSummary) %></p>
<h3><% =strType %></h3><p><% =strInfo %></p>

<table border=0>
  <tr><td><b><% =getIDS("IDS_ArticleId") %></b></td><td align="right">&nbsp;&nbsp;<% =bigDigitNum(10,lngRecordId) %></td></tr>
  <tr><td><b><% =getIDS("IDS_Updated") %></b></td><td align="right">&nbsp;&nbsp;<% =datUpdated %></td></tr>
  <tr><td><b><% =getIDS("IDS_KBHits") %></b></td><td align="right">&nbsp;&nbsp;<% =lngViews %></td></tr>
  <tr><td><b><% =getIDS("IDS_KBRating") %></b></td><td align="right">&nbsp;&nbsp;<% If valNum(lngRateCount,3,0) > 0 and valNum(lngRateTotal,3,0) > 0 Then Response.Write(FormatNumber((lngRateTotal/lngRateCount),2)) Else Response.Write("0.00") %></td></tr>
  <tr><td><b><% =getIDS("IDS_KBReviews") %></b></td><td align="right">&nbsp;&nbsp;<% =lngRateCount %></td></tr>
</table>
</body>
</html>
<%
	Else
		Call DisplayHeader(1)
%>

<div id="modDiv" class="dvBorder">

<table border=0 cellspacing=0 cellpadding=0 width="100%">
  <tr>
	<td width="100%" valign=top>
	  <table width="100%" border=0 cellspacing=0 cellpadding=2>
		<tr class="hRow">
	  <th class="hFont">&nbsp;<% =strCrumbs & " &raquo; " & strTitle %></th>
	  <th align=right>
<%
	If strType = getIDS("IDS_Description") Then
		Response.Write("<a href=""article.asp?id=" & lngRecordId & "&export=true""><img src=""../images/export2.gif"" alt=""" & getIDS("IDS_Export") & """ border=0 height=16 width=16 /></a>" & vbCrLf)
	End If
	If blnRS and (blnAdmin or intPerm >= 3) Then
		Response.Write(getIconImport(2,getEditURL(8,"?id=" & lngRecordId),strTitle))
	End If
%>
		<a href="article.asp?id=<% =lngRecordId %>&print=true"><img src="../images/print2.gif" alt="<% =getIDS("IDS_Print") %>" border=0 height=16 width=16 /></a>
	  </th>
	</tr>
	  </table>
	</td>
  </tr>
</table>

</div>

<div id="contentDiv" class="dvBorder" style="height:<% =intScreenH-100 %>px;">

<% If strTitle <> getIDS("IDS_Deleted") Then %>
<span class="tFont"><% =strTitle %></span>
<hr />
<p class="bFont"><% =getIDS("IDS_Summary") %></p>
<p class="dFont" style="padding-left:15px;"><% =showParagraph(strSummary) %></p>
<p class="bFont"><% =strType %></p>
<div class="dFont" style="padding-left:15px;">
<% Response.Write(strInfo) %>
</div><br />

<table border=0 class="hRow" width="25%">
  <tr><td class="bFont"><% =getIDS("IDS_ArticleId") %></td><td>&nbsp;</td><td class="dRow1" align="right"><span class="dFont"><% =bigDigitNum(10,lngRecordId) %></span></td></tr>
  <tr><td class="bFont"><% =getIDS("IDS_Updated") %></td><td>&nbsp;</td><td class="dRow1" align="right"><span class="dFont"><% =datUpdated %></span></td></tr>
  <tr><td class="bFont"><% =getIDS("IDS_KBHits") %></td><td>&nbsp;</td><td class="dRow1" align="right"><span class="dFont"><% =lngViews %></span></td></tr>
  <tr><td class="bFont"><% =getIDS("IDS_KBRating") %></td><td>&nbsp;</td><td class="dRow1" align="right"><span class="dFont"><% If valNum(lngRateCount,3,0) > 0 and valNum(lngRateTotal,3,0) > 0 Then Response.Write(FormatNumber((lngRateTotal/lngRateCount),2)) Else Response.Write("0.00") %></span></td></tr>
  <tr><td class="bFont"><% =getIDS("IDS_KBReviews") %></td><td>&nbsp;</td><td class="dRow1" align="right"><span class="dFont"><% =lngRateCount %></span></td></tr>
</table>

<% If Request.Form.Count = 0 Then %>
<form name="frmRate" method="post" action="article.asp?id=<% =lngRecordId %>">
  <p><% =getLabel(getIDS("IDS_KBRateArticle"),"selRating") %>&nbsp;&nbsp;&nbsp;
  <select name="selRating" id="selRating" class="mText" style="width:40px;">
	<option value="0">0</option>
	<option value="1">1</option>
	<option value="2">2</option>
	<option value="3">3</option>
	<option value="4">4</option>
	<option value="5">5</option>
  </select>
  <% =getSubmit("btnSubmit",getIDS("IDS_Go"),30,"S","") %>
  </p>
</form>
<% End If %>
<% End If %>
</div>

<%
		Call DisplayFooter(1)
	End If
%>