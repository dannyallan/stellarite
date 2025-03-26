<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_attachments.asp" -->
<%
	Call pageFunctions(0,2)

	Dim lngEventId		'as Long
	Dim strDocType		'as String
	Dim intDocType		'as Integer
	Dim bytPermissions	'as Byte
	Dim bytTotal		'as Byte
	Dim strAttachTitle	'as String
	Dim strInfo			'as String
	Dim strCreatedBy	'as String
	Dim strModBy		'as String
	Dim datModDate		'as Date
	Dim datCreatedDate	'as Date
	Dim lngUploadId		'as Long

	strTitle = Application("IDS_AttachmentInfo")
	lngEventId = valNum(Request.QueryString("eid"),3,0)

	Call Randomize()
	lngUploadId = CLng(Rnd * &H7FFFFFFF)

	If bytMod = "" or lngModId = "" Then Call logError(3,1)

	If strDoAction <> "" then

		intDocType = valNum(Request.Form("selDocType"),2,-1)
		bytPermissions = valNum(Request.Form("selPermissions"),1,1)
		strAttachTitle = valString(Request.Form("txtTitle"),40,0,0)
		strInfo = valString(Request.Form("txtInfo"),255,0,4)
		bytTotal = valNum(Request.Form("hdnFiles"),1,1)

		If strDoAction = "new" Then

			lngRecordId = insertAttach(lngUserId,bytMod,lngModId,lngEventId,intDocType,bytPermissions,strAttachTitle,strInfo)

			For i = 0 to bytTotal

				If valString(Request.Form("file" & i),-1,0,0) <> "" Then
					objConn.Execute(insertAttachLinks(lngRecordId,valString(Request.Form("file" & i),-1,0,2)))
				End If
			Next
		End If
		Call closeWindow(strOpenerURL)
	Else

		If blnRS and intPerm >= 3 Then

			Set objRS = objConn.Execute(getAttach(lngRecordId,intMember))

			If not (objRS.BOF and objRS.EOF) then
				intDocType = objRS.fields("A_DocType").value
				bytPermissions = objRS.fields("A_Permissions").value
				strAttachTitle = showString(objRS.fields("A_Title").value)
				strInfo = showString(objRS.fields("A_Info").value)
				strCreatedBy = showString(objRS.fields("CreatedBy").value)
				datCreatedDate = objRS.fields("A_CreatedDate").value
				strModBy = showString(objRS.fields("ModBy").value)
				datModDate = objRS.fields("A_ModDate").value
			End If

		Elseif blnRS Then
			Call doRedirect("pop_attach.asp?m=" & bytMod & "&mid=" & lngModId)
		Else
			strCreatedBy = strFullName
			datCreatedDate = Now
			strModBy = strFullName
			datModDate = Now
		End If
		strDocType = getOptionDropDown(150,False,"selDocType","Document Type",intDocType)
	End If

	Call DisplayHeader(3)
	Call showEditHeader(strTitle,strCreatedBy,datCreatedDate,strModBy,datModDate)
%>
<script>
var nfiles=1;
function doExpand() {
	nfiles++
	window.document.forms[0].hdnFiles.value = nfiles;
	files.insertAdjacentHTML('BeforeEnd','<br><input type="file" name="file'+nfiles+'" class="oLink" onChange="doChange();" size="57">');
}

function doLink(fType) {
	if (fType == "upload") {
<% If Application("av_UploadType") = 0 and intMode = 0 Then %>
		window.open('pop_progress.asp?uid=<%=lngUploadId%>','_blank','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=yes,width=350,height=200');
<% End If %>
		document.forms[0].action = "pop_upload.asp?id=<% =lngRecordId %>&m=<% =bytMod %>&mid=<% =lngModId %>&eid=<% =lngEventId %>&uid=<% =lngUploadId %>";
		document.forms[0].encoding = "multipart/form-data";
	}
	else {
		document.forms[0].action = "pop_attach.asp?id=<% =lngRecordId %>&m=<% =bytMod %>&mid=<% =lngModId %>&eid=<% =lngEventId %>";
	}
	confirmAction('<% =strAction %>');
}
</script>

<div id="contentDiv" class="dvBorder" style="height:355px;"><br>

<table border=0 cellspacing=5 width="100%">
<form name="frmAttach" method="post">
<% =getHidden("hdnWinOpen","") %>
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
<% =getHidden("hdnFiles","1") %>
  <% If Application("av_Upload") Then %>
  <tr>
    <td class="dfont"><p><% =Replace(Application("IDS_MsgFileSize"),"[SIZE]",Application("av_UploadLimit")/1000000) %></p></td>
  </tr>
  <% End If %>
</table>

<table border=0 cellspacing=5>
  <tr>
     <td class="dfont"><a href="Javascript:doExpand();"><% =Application("IDS_AddNewFile") %></a></td>
  </tr>
</table>

<div id=files style="width:98%;height:62px;overflow:auto;padding:0px">
<input type="file" name="file1" class="mLink" onChange="doChange();" size="57">
</div>

<table border=0 cellspacing=5>
  <tr>
    <td valign=top><% =getLabel(Application("IDS_Title"),"txtTitle") %></td>
    <td><% =getTextField("txtTitle","oText","",57,40,"") %></td>
  </tr>
  <tr>
    <td valign=top><% =getLabel(Application("IDS_Description"),"txtInfo") %></td>
    <td><% =getTextArea("txtInfo","oMemo",strInfo,"360px",3,"") %></td>
  </tr>
  <tr>
    <td valign=top><% =getLabel(Application("IDS_Type"),"selDocType") %></td>
    <td><% =strDocType %></td>
  </tr>
  <tr>
    <td valign=top><% =getLabel(Application("IDS_Permissions"),"selPermissions") %></td>
    <td><% =getPermissionsDropDown(bytPermissions,intMember) %></td>
  </tr>
</form>
</table>

</div>

<div id="footerDiv" class="dvFooter">
<%
	If not blnRS or (blnRS and intPerm >= 3) Then
		If Application("av_Upload") Then
			Response.Write(getIcon("Javascript:doLink('upload');","U","upload.gif",Application("IDS_Upload")))
		End If
		Response.Write(getIcon("Javascript:doLink('link');","L","link.gif",Application("IDS_CreateLink")))
	End If
	Response.Write(getIconCancel())
%>
</div>

<%
	Call DisplayFooter(3)
%>