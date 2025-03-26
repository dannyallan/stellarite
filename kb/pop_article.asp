<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_kb.asp" -->
<%
	Call pageFunctions(8,2)

	Dim lngCatId		'as Long
	Dim strCategory		'as String
	Dim strArticleTitle	'as String
	Dim strKeywords		'as String
	Dim strSummary		'as String
	Dim strInfo			'as String
	Dim datExpiry		'as Date
	Dim intPermissions	'as Integer
	Dim strCreatedBy	'as String
	Dim strModBy		'as String
	Dim datModDate		'as Date
	Dim datCreatedDate	'as Date

	strTitle = Application("IDS_ArticleNew")
	lngCatId = valNum(Request.QueryString("cat"),3,0)

	If strDoAction <> "" then

		lngCatId = valNum(Request.Form("selCategory"),3,0)
		strArticleTitle = valString(Request.Form("txtTitle"),40,1,0)
		strKeywords = valString(Request.Form("txtKeywords"),255,0,0)
		strSummary = valString(Request.Form("txtSummary"),255,0,4)
		strInfo = valString(Request.Form("txtInfo"),-1,0,5)
		datExpiry = valDate(Request.Form("txtExpiry"),1)
		intPermissions = valNum(Request.Form("selPermissions"),1,1)

		If strDoAction = "del" and intPerm >= 4 Then

			Call delArticle(lngUserId,lngRecordId,lngCatId)

		Elseif strDoAction = "edit" and intPerm >= 3 Then

			Call updateArticle(lngUserId,lngRecordId,lngCatId,strArticleTitle,strKeywords,strSummary,strInfo,datExpiry,intPermissions)

		ElseIf strDoAction = "new" Then

			lngRecordId = insertArticle(lngUserId,lngRecordId,lngCatId,strArticleTitle,strKeywords,strSummary,strInfo,datExpiry,intPermissions)
		End If

		If Instr(strOpenerURL,"default.asp") > 0 Then strOpenerURL = ""
		Call closeWindow(strOpenerURL)
	Else

		If blnRS and intPerm >= 3 Then

			Set objRS = objConn.Execute(getArticle(0,lngRecordId))

			If not (objRS.BOF and objRS.EOF) then
				strArticleTitle = showString(objRS.fields("H_Title").value)
				lngCatId = objRS.fields("CatId").value
				strKeywords = showString(objRS.fields("H_Keywords").value)
				strSummary = showString(objRS.fields("H_Summary").value)
				strInfo = objRS.fields("H_Info").value
				datExpiry = showDate(0,objRS.fields("H_Expire").value)
				intPermissions = objRS.fields("H_Permissions").value
				strCreatedBy = showString(objRS.fields("CreatedBy").value)
				datCreatedDate = objRS.fields("H_CreatedDate").value
				strModBy = showString(objRS.fields("ModBy").value)
				datModDate = objRS.fields("H_ModDate").value
			End If

		Elseif blnRS Then

			Call doRedirect("pop_article.asp")
		Else
			datExpiry = showDate(0,DateAdd("yyyy",1,Date))
			strCreatedBy = strFullName
			datCreatedDate = Now
			strModBy = strFullName
			datModDate = Now
		End If
		strCategory = getCategoryDropDown(317,False,"selCategory",lngCatId)
	End If

	strIncHead = getEditorScripts() & vbCrLf & getCalendarScripts()

	Call DisplayHeader(3)
	Call showEditHeader(strTitle,strCreatedBy,datCreatedDate,strModBy,datModDate)
%>
<div id="contentDiv" class="dvBorder" style="height:505px;"><br>

<table border=0 cellspacing=5 width="100%">
<form name="frmArticle" method="post" action="pop_article.asp?id=<% =lngRecordId %>">
<% =getHidden("hdnWinOpen","") %>
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
  <tr>
    <td valign=top><% =getLabel(Application("IDS_Title"),"txtTitle") %></td>
    <td><% =getTextField("txtTitle","mText",strArticleTitle,50,40,"") %></td>
  </tr>
  <tr>
    <td valign=top><% =getLabel(Application("IDS_CategoryParent"),"selCategory") %></td>
    <td><% =strCategory %></td>
  </tr>
  <tr>
    <td valign=top><% =getLabel(Application("IDS_Keywords"),"txtKeywords") %></td>
    <td><% =getTextField("txtKeywords","oText",strKeywords,50,40,"") %></td>
  </tr>
  <tr>
    <td valign=top><% =getLabel(Application("IDS_Summary"),"txtSummary") %></td>
    <td><% =getTextArea("txtSummary","oMemo",strSummary,"100%",2,"") %></td>
  </tr>
  <tr>
    <td valign=top><% =getLabel(Application("IDS_Description"),"txtInfo") %></td>
    <td>
<% =getTextArea("txtInfo","oText",strInfo,"100%",18,"") %>
    </td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_Expiry"),"txtExpiry") %></td>
    <td>
      <% =getTextField("txtExpiry","mDate",datExpiry,19,255,"") %>
      <a href="Javascript:showCalendar('txtExpiry');"><img src="../images/cal.gif" alt="<% =getImport("IDS_Expiry") %>" border=0 height=16 width=16></a></td>
  </tr>
  <tr>
    <td><% =getLabel(Application("IDS_Permissions"),"selPermissions") %></td>
    <td><% =getPermissionsDropDown(intPermissions,intMember) %></td>
  </tr>
</form>
</table>

</div>

<div id="footerDiv" class="dvFooter">
<%
	If blnRS Then
		Response.Write(getIconNew("pop_article.asp?cat=" & lngCatId))
		If intPerm >= 4 Then Response.Write(getIconDelete())
	End If
	Response.Write(getIcon("Javascript:document.forms[0].onsubmit();confirmAction('" & strAction & "');","S","save.gif",Application("IDS_Save")))
	Response.Write(getIconCancel())
%>
</div>

<script language="Javascript">
	var editor = new HTMLArea("txtInfo");
	editor.generate();
</script>

<%
	Call DisplayFooter(3)
%>