<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_kb.asp" -->
<%
	Call pageFunctions(8,2)

	Dim lngCatId        'as Long
	Dim strCategory     'as String
	Dim strArticleTitle 'as String
	Dim strKeywords     'as String
	Dim strSummary      'as String
	Dim bytType			'as Byte
	Dim strInfo         'as String
	Dim strLink         'as String
	Dim datExpiry       'as Date
	Dim intPermissions  'as Integer
	Dim strCreatedBy    'as String
	Dim strModBy        'as String
	Dim datModDate      'as Date
	Dim datCreatedDate  'as Date
	Dim lngUploadId     'as Long

	strTitle = getIDS("IDS_ArticleNew")
	lngCatId = valNum(Request.QueryString("cat"),3,0)

	Call Randomize()
	lngUploadId = CLng(Rnd * &H7FFFFFFF)

	If strDoAction <> "" then

		lngCatId = valNum(Request.Form("selCategory"),3,1)

		Select Case strDoAction
			Case "del"
				If intPer >= 4 Then Call delArticle(lngUserId,lngRecordId,lngCatId)

			Case "new","edit"

				strArticleTitle = valString(Request.Form("txtTitle"),40,1,0)
				strKeywords = valString(Request.Form("txtKeywords"),40,0,0)
				strSummary = valString(Request.Form("txtSummary"),255,0,4)
				bytType = valNum(Request.Form("selType"),1,1)
				datExpiry = valDate(Request.Form("txtExpiry"),1)
				intPermissions = valNum(Request.Form("selPermissions"),1,1)

				Select Case bytType
					Case 0
						strInfo = valString(Request.Form("txtInfo"),-1,0,5)
					Case 1
						strLink = valString(Request.Form("txtLink"),-1,0,2)
				End Select

				If strDoAction = "edit" and intPerm >= 3 Then
					Call updateArticle(lngUserId,lngRecordId,lngCatId,strArticleTitle,strKeywords,strSummary,strInfo,strLink,datExpiry,intPermissions)
				ElseIf strDoAction = "new" Then
					lngRecordId = insertArticle(lngUserId,lngRecordId,lngCatId,strArticleTitle,strKeywords,strSummary,strInfo,strLink,datExpiry,intPermissions)
				End If

		End Select
		Call closeEdit()
	Else

		If blnRS and intPerm >= 3 Then

			Set objRS = objConn.Execute(getArticle(0,lngRecordId))

			If not (objRS.BOF and objRS.EOF) then
				strArticleTitle = objRS.fields("H_Title").value
				lngCatId = objRS.fields("CatId").value
				strKeywords = objRS.fields("H_Keywords").value
				strSummary = objRS.fields("H_Summary").value
				strInfo = objRS.fields("H_Info").value
				strLink = objRS.fields("H_Link").value
				datExpiry = objRS.fields("H_Expire").value
				intPermissions = objRS.fields("H_Permissions").value
				strCreatedBy = objRS.fields("CreatedBy").value
				datCreatedDate = objRS.fields("H_CreatedDate").value
				strModBy = objRS.fields("ModBy").value
				datModDate = objRS.fields("H_ModDate").value

				If strLink <> "" Then
					bytType = 1
				Else
					bytType = 0
				End If
			End If

		Elseif blnRS Then

			Call logError(2,1)
		Else
			datExpiry = DateAdd("yyyy",1,Date)
			bytType = 0
			strCreatedBy = strFullName
			datCreatedDate = Now
			strModBy = strFullName
			datModDate = Now
		End If
		strCategory = getCategoryDropDown(317,False,"selCategory",lngCatId)
	End If

	strIncHead = getEditorScripts() & vbCrLf & getCalendarScripts()

	Call DisplayHeader(1)
	Call showEditHeader(strTitle,strCreatedBy,datCreatedDate,strModBy,datModDate)
%>
<form name="frmArticle" method="post"">
<div id="contentDiv" class="dvBorder" style="height:<% =intScreenH-170 %>px;"><br />

<table border=0 cellspacing=5 cellpadding=0 width="100%">
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
  <tr>
	<td width=150><% =getLabel(getIDS("IDS_Title"),"txtTitle") %></td>
	<td><% =getTextField("txtTitle","mText",strArticleTitle,50,40,"") %></td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_CategoryParent"),"selCategory") %></td>
	<td><% =strCategory %></td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_Type"),"selType") %></td>
	<td>
	  <select name="selType" id="selType" class="oByte" style="width: 317px;" onChange="changeType(this.options[this.selectedIndex].value);doChange();">
	    <option value="0"<% =getDefault(0,bytType,0) & ">" & getIDS("IDS_Article") %></option>
	    <option value="1"<% =getDefault(0,bytType,1) & ">" & getIDS("IDS_Link") %></option>
	    <option value="2"<% =getDefault(0,bytType,2) & ">" & getIDS("IDS_File") %></option>
	  </select>
	</td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_Keywords"),"txtKeywords") %></td>
	<td><% =getTextField("txtKeywords","oText",strKeywords,50,40,"") %></td>
  </tr>
  <tr>
	<td valign=top><% =getLabel(getIDS("IDS_Summary"),"txtSummary") %></td>
	<td><% =getTextArea("txtSummary","oMemo",strSummary,"100%",2,"") %></td>
  </tr>
  <tr>
	<td valign=top><% =getLabel(getIDS("IDS_Description"),"txtInfo") %></td>
	<td>
	  <div id="divInfo" style="display:inline;">
<% =getTextArea("txtInfo","oText",strInfo,"100%",18,"") %>
      </div>
      <div id="divLink" style="display:none;">
<% =getTextField("txtLink","oLink",strLink,50,1000,"") %>
      </div>
      <div id="divFile" style="display:none;">
<% =getFileField("filFile","oLink",strLink,50,1000,"") %>
      </div>
	</td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_Expiry"),"txtExpiry") %></td>
	<td><% =getDateField("txtExpiry","mDate",datExpiry,getIDS("IDS_Expiry")) %></td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_Permissions"),"selPermissions") %></td>
	<td><% =getPermissionsDropDown(intPermissions,intMember) %></td>
  </tr>
</table>

</div>

<div id="footerDiv" class="dvFooter">
<%
	If blnRS Then
		Response.Write(getIconNew(getEditURL(8,"?cat=" & lngCatId)))
		If intPerm >= 4 Then Response.Write(getIconDelete())
	End If
	Response.Write(getIcon("Javascript:document.forms[0].onsubmit();doProgressBar();confirmAction('" & strAction & "');","S","save.gif",getIDS("IDS_Save")))
	Response.Write(getIconCancel("back"))
%>
</div>
</form>

<script language="JavaScript" type="text/javascript">

function changeType(iType) {
	var oInfo = getObject("divInfo");
	var oLink = getObject("divLink");
	var oFile = getObject("divFile");
	var oDesc = getObject("lblInfo");

	oInfo.style.display = "none";
	oLink.style.display = "none";
	oFile.style.display = "none";

	switch (parseInt(iType)) {
		case 0:
			oDesc.innerHTML = "<% =getIDS("IDS_Description") %>";
			oInfo.style.display = "inline";
			document.forms[0].action = "edit_article.asp?id=<% =lngRecordId %>";
			document.forms[0].encoding = "application/x-www-form-urlencoded";
			break;
		case 1:
			oDesc.innerHTML = "<% =getIDS("IDS_Link") %>";
			oLink.style.display = "inline";
			document.forms[0].action = "edit_article.asp?id=<% =lngRecordId %>";
			document.forms[0].encoding = "application/x-www-form-urlencoded";
			break;
		case 2:
			oDesc.innerHTML = "<% =getIDS("IDS_File") %>";
			oFile.style.display = "inline";
			document.forms[0].action = "upload_article.asp?id=<% =lngRecordId %>&uid=<% =lngUploadId %>";
			document.forms[0].encoding = "multipart/form-data";
			break;
	}
}

function doProgressBar() {
<% If Application("av_UploadType") = 0 and intMode = 0 Then %>
	if (document.forms[0].selType.value == 2) {
		window.open('../common/pop_progress.asp?uid=<%=lngUploadId%>','_blank','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=yes,width=350,height=200');
	}
<% End If %>
}

var editor = new HTMLArea("txtInfo");
editor.generate();
changeType(<% =bytType %>);

</script>

<%
	Call DisplayFooter(1)
%>