<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_kb.asp" -->
<%
	Call pageFunctions(8,5)

	Dim lngParentId		'as Long
	Dim strParents		'as String
	Dim strCategory		'as String
	Dim strDescription	'as String
	Dim datUpdated		'as Date
	Dim strCreatedBy	'as String
	Dim datCreatedDate	'as Date
	Dim strModBy		'as String
	Dim datModDate		'as Date

	strTitle = Application("IDS_CategoryNew")

	If strDoAction <> "" Then

		lngParentId = valNum(Request.Form("selCategory"),3,0)
		strCategory = valString(Request.Form("txtCategory"),20,1,0)
		strDescription = valString(Request.Form("txtDescription"),255,0,4)

		If strDoAction = "del" Then

			Call delCategory(lngUserId,lngRecordId)
			Call remAppVar("Categories")

		ElseIf strDoAction = "edit" Then

			Call updateCategory(lngUserId,lngRecordId,lngParentId,strCategory,strDescription)

		ElseIf strDoAction = "new" Then

			lngRecordId = insertCategory(lngUserId,lngRecordId,lngParentId,strCategory,strDescription)
			Call remAppVar("Categories")
		End If
		Call closeWindow(strOpenerURL)
	Else

		If blnRS Then
			Set objRS = objConn.Execute(getCategory(1,lngRecordId))

			If not (objRS.BOF and objRS.EOF) then
				lngParentId = objRS.fields("I_ParentId").value
				strCategory = showString(objRS.fields("I_Name").value)
				strDescription = showString(objRS.fields("I_Description").value)
				datUpdated = showDate(0,objRS.fields("I_Updated").value)
			End If
		Else
			lngParentId = valNum(Request.QueryString("cat"),3,0)
			strCreatedBy = strFullName
			datCreatedDate = Now
			strModBy = strFullName
			datModDate = Now
		End If
		strParents = getCategoryDropDown(290,True,"selCategory",lngParentId)
	End If

	Call DisplayHeader(3)
	Call showEditHeader(strTitle,strCreatedBy,datCreatedDate,strModBy,datModDate)
%>
<div id="contentDiv" class="dvBorder" style="height:205px;"><br>

<table border=0 width="100%">
<form name="frmCategory" method="post" action="pop_category.asp?id=<% =lngRecordId %>">
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
<% =getHidden("hdnWinOpen","") %>
    <tr>
      <td><% =getLabel(Application("IDS_Name"),"txtCategory") %></td>
      <td><% =getTextField("txtCategory","mText",strCategory,45,20,"") %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_CategoryParent"),"selCategory") %></td>
      <td><% =strParents %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_Description"),"txtDescription") %></td>
      <td><% =getTextArea("txtDescription","oMemo",strDescription,"290px",7,"") %></td>
    </tr>
</form>
</table>

</div>

<div id="footerDiv" class="dvFooter">
<%
	If blnRS Then
		Response.Write(getIconNew("pop_category.asp"))
		If intPerm >= 4 Then Response.Write(getIconDelete())
	End If
	Response.Write(getIconSave(strAction))
	Response.Write(getIconCancel())
%>
</div>

<%
	Call DisplayFooter(3)
%>