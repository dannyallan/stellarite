<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_admin.asp" -->
<%
	Call pageFunctions(90,5)

	Dim bytType     'as Byte        'Type of Field
	Dim strName     'as String      'Name of Field
	Dim intSize     'as Integer     'Maximum data size
	Dim blnMand     'as Boolean     'Mandatory field
	Dim bytOrder    'as Byte        'Order of custom fields

	strTitle = getIDS("IDS_EditCustomFields")
	lngRecordId = valNum(lngRecordId,1,0)

	If strDoAction <> "" Then

		Select Case strDoAction
			Case "new"

				bytMod = valNum(Request.Form("selModule"),1,1)
				bytType = valNum(Request.Form("selType"),1,1)
				strName = valString(Request.Form("txtName"),40,1,0)
				blnMand = valNum(Request.Form("chkMand"),0,0)

				Select Case bytType
					Case 1,6,7
						intSize = valNum(Request.Form("txtSize"),1,1)
					Case 2
						intSize = valNum(Request.Form("selSize"),1,1)
					Case 3
						intSize = 7
					Case 4
						intSize = 17
					Case 5
						intSize = 2
				End Select

				'bytOrder = valNum(Request.Form("txtOrder"),1,0)
				bytOrder = UBound(Application("arr_Fields" & bytMod),2) + 10

				objConn.Execute(insertCustomField(bytMod,bytType,strName,intSize,blnMand,bytOrder))

			Case "edit"

				blnMand = valNum(Request.Form("chkMand"),0,0)
				bytOrder = valNum(Request.Form("txtOrder"),1,1)

				objConn.Execute(updateCustomField(lngRecordId,blnMand,bytOrder))

			Case "del"
				lngRecordId = valNum(Request.Form("hdnDelete"),3,0)
				objConn.Execute(delCustomField(lngRecordId))
		End Select

		Set objRS = objConn.Execute(getModuleFields(bytMod))
		If not (objRS.BOF and objRS.EOF) Then Application("arr_Fields" & bytMod) = objRS.GetRows()

		Call doRedirect("edit_custom.asp")

	Elseif blnRS Then
		strName = getOptionGroupName(lngRecordId)
	Else
		Set objRS = objConn.Execute(getCustomFields(Session("Permissions")))
		If not (objRS.BOF and objRS.EOF) Then arrRS = objRS.GetRows()
	End If

	Call DisplayHeader(1)
	Call showEditHeader(strTitle,"","","","")
%>
<form name="frmAdmin" method="post" action="edit_custom.asp">
<div id="contentDiv" class="dvBorder" style="height:330px;"><br />

<table border=0 cellspacing=0 cellpadding=2 width="100%">
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
<% =getHidden("hdnDelete","") %>
<%
	If blnRS Then

		Response.Write("  <tr><td>" & getLabel(getIDS("IDS_Name"),"txtName") & "</td>" & _
				"<td>" & getTextField("txtName","mText",strName,29,40,"") & "</td></tr>" & vbCrLf)

		If lngRecordId = 0 Then
			Response.Write("  <tr><td>" & getLabel(getIDS("IDS_Module"),"selModule") & "</td>" & _
					"<td>" & getModuleDropDown("selModule",bytMod,False,"") & "</td></tr>" & vbCrLf)

			Response.Write("  <tr><td>" & getLabel(getIDS("IDS_Type"),"selType") & "</td>" & _
					"<td><select name=""selType"" id=""selType"" style=""width:190px;"" class=""oByte"" onChange=""doSize(this.value);"">" & _
						"<option value=""1""" & getDefault(0,1,bytType) & ">" & getIDS("IDS_Text") & "</option>" & _
						"<option value=""2""" & getDefault(0,2,bytType) & ">" & getIDS("IDS_Number") & "</option>" & _
						"<option value=""3""" & getDefault(0,3,bytType) & ">" & getIDS("IDS_Date") & "</option>" & _
						"<option value=""4""" & getDefault(0,4,bytType) & ">" & getIDS("IDS_True") & " / " & getIDS("IDS_False") & "</option>" & _
						"<option value=""5""" & getDefault(0,5,bytType) & ">" & getIDS("IDS_DropDown") & "</option>" & _
						"<option value=""6""" & getDefault(0,6,bytType) & ">" & getIDS("IDS_Email") & "</option>" & _
						"<option value=""7""" & getDefault(0,7,bytType) & ">" & getIDS("IDS_Password") & "</option>" & _
					"</select></td></tr>" & vbCrLf)

			Response.Write("  <tr><td>" & getLabel(getIDS("IDS_Size"),"txtSize") & "</td>" & _
					"<td><div id=""divText"" style=""display:inline;"">" & getTextField("txtSize","mByte",intSize,4,10,"") & "</div>" & _
					"<div id=""divCmbo"" style=""display:none;""><select name=""selSize"" id=""selSize"" style=""width:190px;"" class=""oByte"">" & _
					"<option value=""17"">Tiny Integer</option>" & _
					"<option value=""2"">Small Integer</option>" & _
					"<option value=""3"">Integer</option>" & _
					"<option value=""5"">Double</option>" & _
					"<option value=""6"">Decimal</option>" & _
					"</select></div></td></tr>" & vbCrLf)

		End If

'		Response.Write("  <tr><td>" & getLabel(getIDS("IDS_Order"),"txtOrder") & "</td>" & _
'				"<td>" & getTextField("txtOrder","mByte",bytOrder,4,10,"") & "</td></tr>" & vbCrLf)

		Response.Write("  <tr><td>" & getLabel(getIDS("IDS_Mandatory"),"chkMand") & "</td>" & _
				"<td>" & getCheckbox("chkMand",blnMand,"") & "</td></tr>" & vbCrLf)
	Else
		Response.Write("  <tr><td colspan=3 class=""dFont""><a href=""edit_custom.asp?id=0"">" & getIDS("IDS_AddCustomField") & "</a><br /><br /></td></tr>" & _
				"<tr class=""hRow""><td class=""hFont"">" & getIDS("IDS_Name") & "</td><td colspan=2 class=""hFont"">" & getIDS("IDS_Module") & "</td></tr>" & vbCrLf)
		If isArray(arrRS) Then
			For i = 0 to UBound(arrRS,2)
				Response.Write("  <tr><td class=""dFont"">" & showString(arrRS(3,i)) & "</td>" & _
						"<td class=""dFont"">" & getIDS("IDS_ModItem" & arrRS(2,i)) & "</td>" & _
						"<td class=""dFont rIndent"" width=""10%"" align=""right"">" & getIconImport(4,"Javascript:document.forms[0].hdnDelete.value='" & arrRS(0,i) & "';confirmAction('del');",showString(arrRS(3,i))) & "</td></tr>" & vbCrLf)
			Next
		Else
			Response.Write("<tr><td colspan=3 class=""dFont"">" & getIDS("IDS_NoneSpecified") & "</td></tr>" & vbCrLf)
		End If
	End If
%>
</table>

</div>

<div id="footerDiv" class="dvFooter">
<%
	If blnRS Then
		If lngRecordId = 0 Then
			Response.Write(getIconSave("new"))
		Else
			Response.Write(getIconSave("edit"))
		End If
		Response.Write(getIconCancel("edit_custom.asp"))
	Else
		Response.Write(getIconCancel("default.asp"))
	End If
%>
</div>
</form>

<script language="JavaScript" type="text/javascript">

function doSize(iType) {
	var oText = getObject("divText");
	var oCmbo = getObject("divCmbo");
	var oSize = getObject("txtSize");

	oText.style.display = "none";
	oCmbo.style.display = "none";
	oSize.className = "oByte";

	switch (parseInt(iType)) {
		case 1: case 6: case 7:
			oText.style.display = "inline";
			oSize.className = "mByte";
			break;
		case 2:
			oCmbo.style.display = "inline";
			break;
	}
}

</script>

<%
	Call DisplayFooter(3)
%>