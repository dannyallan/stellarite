<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_kb.asp" -->
<%
	Call pageFunctions(8,5)

	Dim lngParentId     'as Long
	Dim strParents      'as String
	Dim strCategory     'as String
	Dim strDescription  'as String
	Dim strCreatedBy    'as String
	Dim datCreatedDate  'as Date
	Dim strModBy        'as String
	Dim datModDate      'as Date

	strTitle = getIDS("IDS_CategoryNew")

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
		Call closeEdit()
	Else

		If blnRS Then
			Set objRS = objConn.Execute(getCategory(1,lngRecordId))

			If not (objRS.BOF and objRS.EOF) then
				lngParentId = objRS.fields("I_ParentId").value
				strCategory = objRS.fields("I_Name").value
				strDescription = objRS.fields("I_Description").value
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

	Call DisplayHeader(1)
	Call showEditHeader(strTitle,strCreatedBy,datCreatedDate,strModBy,datModDate)
%>
<form name="frmCategory" method="post" action="edit_category.asp?id=<% =lngRecordId %>">
<div id="contentDiv" class="dvBorder" style="height:<% =intScreenH-170 %>px;"><br />

<table border=0 width="100%">
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
	<tr>
	  <td><% =getLabel(getIDS("IDS_Name"),"txtCategory") %></td>
	  <td><% =getTextField("txtCategory","mText",strCategory,45,20,"") %></td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_CategoryParent"),"selCategory") %></td>
	  <td><% =strParents %></td>
	</tr>
	<tr>
	  <td><% =getLabel(getIDS("IDS_Description"),"txtDescription") %></td>
	  <td><% =getTextArea("txtDescription","oMemo",strDescription,"290px",7,"") %></td>
	</tr>
</table>

</div>

<div id="footerDiv" class="dvFooter">
<%
	If blnRS Then
		Response.Write(getIconNew(getEditURL("C","")))
		If intPerm >= 4 Then Response.Write(getIconDelete())
	End If
	Response.Write(getIconSave(strAction))
	Response.Write(getIconCancel("back"))
%>
</div>
</form>

<%
	Call DisplayFooter(1)
%>