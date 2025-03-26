<!--#include file="..\_inc\functions.asp" -->
<%
	Call pageFunctions(90,5)

	strTitle = getIDS("IDS_Administration")

	Call DisplayHeader(1)

	Response.Write("<div id=""contentDiv"" class=""dvBorder"">" & vbCrLf & vbCrLf & _
			"<table border=0 cellspacing=0 cellpadding=5 width=""100%"">" & vbCrLf & _
			"<thead><tr class=""hRow"">" & vbCrLf & _
			"  <th class=""hFont"" width=""50%"" valign=top>" & getIDS("IDS_ModuleAdministration") & "</th>" & vbCrLf & _
			"  <th class=""hFont"" width=""50%"" valign=top>" & getIDS("IDS_CRMAdministration") & "</th>" & vbCrLf & _
			"</tr></thead>" & vbCrLf & _
			"<tbody><tr><td class=""dFont""><ul>")

	If pContacts >= 5 Then Response.Write("<li><a href=""edit_msg.asp?id=1"">" & getIDS("IDS_EditNoticeContact") & "</a></li>" & vbCrLf)
	If pClients >= 5 Then Response.Write("<li><a href=""edit_msg.asp?id=2"">" & getIDS("IDS_EditNoticeClient") & "</a></li>" & vbCrLf)
	If pSales >= 5 Then Response.Write("<li><a href=""edit_msg.asp?id=3"">" & getIDS("IDS_EditNoticeSales") & "</a></li>" & vbCrLf)
	If pProjects >= 5 Then Response.Write("<li><a href=""edit_msg.asp?id=4"">" & getIDS("IDS_EditNoticeServices") & "</a></li>" & vbCrLf)
	If pTickets >= 5 Then Response.Write("<li><a href=""edit_msg.asp?id=5"">" & getIDS("IDS_EditNoticeSupport") & "</a></li>" & vbCrLf)
	If pBugs >= 5 Then Response.Write("<li><a href=""edit_msg.asp?id=6"">" & getIDS("IDS_EditNoticeQA") & "</a></li>" & vbCrLf)
	If pInvoices >= 5 Then Response.Write("<li><a href=""edit_msg.asp?id=7"">" & getIDS("IDS_EditNoticeFinance") & "</a></li>" & vbCrLf)

	Response.Write("<div class=""hr""></div>" & vbCrLf & _
					"<li><a href=""edit_permissions.asp?id=1"">" & getIDS("IDS_EditDefaultPermissions") & "</a></li>" & vbCrLf & _
					"<div class=""hr""></div>" & vbCrLf & _
					"<li><a href=""edit_custom.asp"">" & getIDS("IDS_EditCustomFields") & "</a></li>" & vbCrLf & _
					"<li><a href=""edit_menus.asp"">" & getIDS("IDS_EditOptionValues") & "</a></li>" & vbCrLf & _
					"<div class=""hr""></div>" & vbCrLf)

	If blnAdmin or pTickets >=5 or pBugs >= 5 Then Response.Write("<li><a href=""edit_products.asp"">" & getIDS("IDS_EditProducts") & "</a></li>")

	Response.Write("</ul></td><td class=""dFont"">" & vbCrLf)

	If blnAdmin Then

		Response.Write("<ul><li><a href=""crm_info.asp"">" & getIDS("IDS_ServerInformation") & "</a></li>" & vbCrLf & _
				"<li><a href=""crm_variables.asp"">" & getIDS("IDS_ServerVariables") & "</a></li>" & vbCrLf & _
				"<div class=""hr""></div>" & vbCrLf & _
				"<li><a href=""edit_msg.asp?id=0"">" & getIDS("IDS_EditNoticeCRM")& "</a></li>" & vbCrLf & _
				"<div class=""hr""></div>" & vbCrLf & _
				"<li><a href=""edit_admin.asp"">" & getIDS("IDS_CRMAdministration") & "</a></li>" & vbCrLf & _
				"<li><a href=""edit_parameters.asp"">" & getIDS("IDS_CRMParameters") & "</a></li>" & vbCrLf & _
				"<li><a href=""edit_style.asp"">" & getIDS("IDS_CRMStylesheet") & "</a></li>" & vbCrLf & _
				"<li><a href=""edit_modules.asp"">" & getIDS("IDS_EnableModules") & "</a></li>" & vbCrLf & _
				"<li><a href=""edit_email.asp"">" & getIDS("IDS_EmailOptions") & "</a></li>" & vbCrLf & _
				"<li><a href=""edit_errorlog.asp"">" & getIDS("IDS_ErrorOptions") & "</a></li>" & vbCrLf & _
				"<li><a href=""edit_upload.asp"">" & getIDS("IDS_UploadOptions") & "</a></li>" & vbCrLf & _
			"</ul>" & vbCrLf)
	End If

	Response.Write("</td></tr></tbody></table></div>" & vbCrLf & vbCrLf)

	Call DisplayFooter(1)
%>
