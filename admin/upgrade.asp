<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_admin.asp" -->
<%
	Dim strMessage      'as String
	Dim lngProduct      'as Long
	Dim lngVersion      'as Long
	Dim strTemp         'as String

	strTitle = "Stellarite " & getIDS("IDS_Upgrade")

	Sub insertCustomData(sValues)
		objConn.Execute("INSERT INTO ALL_CustomData (X_Status,X_Module,X_Name,X_Field,X_Type,X_Length,X_Required,X_Order,X_Default) VALUES (" & sValues & ")")
	End Sub

	If Request.ServerVariables("REMOTE_ADDR") = "127.0.0.1" Then
		If Request.Form.Count > 0 or Session("Upgrade") = "YES" Then
			Session("Upgrade") = "YES"
			Select Case CInt(Application("av_DatabaseVersion"))
				Case 23

					objConn.Execute("ALTER TABLE ALL_Users ALTER COLUMN U_Permissions varchar(7)")
					objConn.Execute("ALTER TABLE ALL_Users ALTER COLUMN U_Member varchar(7)")
					objConn.Execute("ALTER TABLE CRM_Reports ADD R_Owner integer")
					objConn.Execute("ALTER TABLE CRM_Reports ALTER COLUMN R_Type tinyint")

					objConn.Execute("CREATE TABLE KB_Articles (ArticleId integer IDENTITY (1, 1) NOT NULL PRIMARY KEY, CatId integer NULL DEFAULT 0, H_CreatedBy smallint NULL DEFAULT 0, H_CreatedDate datetime NULL , H_Expire datetime NULL , H_Info LongText NULL , H_Keywords varchar(255) NULL , H_ModBy smallint NULL DEFAULT 0, H_ModDate datetime NULL , H_Permissions tinyint NULL DEFAULT 0, H_RateCount smallint NULL DEFAULT 0, H_RateTotal integer NULL DEFAULT 0, H_Status tinyint NULL DEFAULT 1, H_Summary varchar(255) NULL , H_Title varchar(40) NULL , H_Views smallint NULL DEFAULT 0)")
					objConn.Execute("CREATE TABLE KB_Categories (CatId integer IDENTITY (1, 1) NOT NULL PRIMARY KEY, I_Count integer NULL DEFAULT 0, I_Description varchar(255) NULL , I_Name varchar(20) NULL , I_ParentId integer NULL DEFAULT 0, I_Status tinyint NULL DEFAULT 1, I_Updated datetime NULL )")

					objConn.Execute("CREATE INDEX ArticleId ON KB_Articles(ArticleId)")
					objConn.Execute("CREATE INDEX CatId ON KB_Categories(CatId)")
					objConn.Execute("CREATE INDEX CatId ON KB_Articles(CatId)")

					Call setAppVar("av_DatabaseVersion",24)
				Case 24

					objConn.Execute("ALTER TABLE ALL_OptGroups DROP COLUMN G_Value ")
					objConn.Execute("ALTER TABLE ALL_OptGroups ADD G_Required tinyint, G_Status tinyint ")
					objConn.Execute("ALTER TABLE ALL_OptGroups ALTER COLUMN G_Description varchar(255) ")
					objConn.Execute("ALTER TABLE CRM_Contacts DROP COLUMN K_Type ")
					objConn.Execute("ALTER TABLE CRM_Contacts ADD K_Serials smallint, K_UserName varchar(20), K_Password varchar(35), K_Sales smallint, K_ReportsTo integer, K_Assistant integer ")
					objConn.Execute("ALTER TABLE CRM_Divisions DROP COLUMN D_Password, COLUMN D_Measure1, COLUMN D_Measure2, COLUMN D_Status1, COLUMN D_Status2 ")
					objConn.Execute("ALTER TABLE CRM_Divisions ADD D_Vertical smallint ")
					objConn.Execute("ALTER TABLE CRM_Sales DROP COLUMN S_Expiry, COLUMN S_DaysTotal, COLUMN S_SLA ")
					objConn.Execute("ALTER TABLE CRM_Sales ADD ContactId integer ")
					objConn.Execute("ALTER TABLE CRM_Sales ALTER COLUMN S_PO integer ")
					objConn.Execute("ALTER TABLE CRM_Serialz ADD ContactId integer, Z_Expiry datetime ")
					objConn.Execute("ALTER TABLE CRM_Tickets DROP COLUMN T_SLA ")

					objConn.Execute("INSERT INTO ALL_OptGroups(G_Name, G_Module, G_Description) VALUES ('Sales Vertical', 3, 'This specifies the industry sector which the client belongs to.') ")

					Call setAppVar("av_DatabaseVersion",25)
				Case 25

					objConn.Execute("ALTER TABLE ALL_Users ADD U_ChangePassword tinyint ")
					objConn.Execute("ALTER TABLE CRM_Bugs ADD B_Events smallint ")
					objConn.Execute("ALTER TABLE CRM_Contacts ADD K_Events smallint ")
					objConn.Execute("ALTER TABLE CRM_Divisions ADD D_Events smallint ")
					objConn.Execute("ALTER TABLE CRM_Sales ADD S_Events smallint ")
					objConn.Execute("ALTER TABLE CRM_Tickets ADD T_Events smallint ")

					objConn.Execute("CREATE TABLE ALL_CustomData (CustomId integer IDENTITY (1, 1) NOT NULL PRIMARY KEY, X_Length tinyint NULL DEFAULT 0, X_Module tinyint NULL DEFAULT 0, X_Name varchar(40) NULL , X_Order tinyint NULL DEFAULT 0, X_Required tinyint NULL DEFAULT 0, X_Status tinyint NULL DEFAULT 1, X_Type tinyint NULL DEFAULT 0) ")

					objConn.Execute("CREATE INDEX CustomId ON ALL_CustomData(CustomId) ")

					Call setAppVar("av_DatabaseVersion",26)
				Case 26

					objConn.Execute("INSERT INTO ALL_OptGroups(G_Name, G_Module, G_Description) VALUES ('Sales Phase', 3, 'This option shows the phases through which a sales opportunity is passing.') ")
					objConn.Execute("INSERT INTO ALL_OptGroups(G_Name, G_Module, G_Description) VALUES ('Invoice Phase', 7, 'This option shows the phases through which an invoice may be going.') ")
					objConn.Execute("INSERT INTO ALL_OptGroups(G_Name, G_Module, G_Description) VALUES ('Invoice Type', 7, 'This specifies the type of invoice which was either sent or received.') ")

					objConn.Execute("ALTER TABLE ALL_Users ALTER COLUMN U_Permissions varchar(8) ")
					objConn.Execute("ALTER TABLE ALL_Users ALTER COLUMN U_Member varchar(8) ")
					objConn.Execute("ALTER TABLE ALL_Users ALTER COLUMN U_Portal varchar(30) ")
					objConn.Execute("ALTER TABLE CRM_Contacts ADD K_Invoices smallint ")
					objConn.Execute("ALTER TABLE CRM_Divisions ADD D_Account varchar(25), D_Invoices smallint, D_ShortDesc varchar(255) ")
					objConn.Execute("DROP INDEX SaleId ON CRM_Projects")
					objConn.Execute("ALTER TABLE CRM_Projects DROP COLUMN SaleId ")
					objConn.Execute("ALTER TABLE CRM_Projects ADD InvoiceId integer, P_ShortDesc varchar(255) ")
					objConn.Execute("ALTER TABLE CRM_Reports ADD R_Portal tinyint ")
					objConn.Execute("ALTER TABLE CRM_Sales DROP COLUMN S_Serials, COLUMN S_PO ")
					objConn.Execute("ALTER TABLE CRM_Sales ADD InvoiceId integer, S_Currency varchar(3), S_Phase smallint ")
					objConn.Execute("DROP INDEX SaleId1 ON CRM_Serialz")
					objConn.Execute("ALTER TABLE CRM_Serialz DROP COLUMN SaleId ")
					objConn.Execute("ALTER TABLE CRM_Serialz ADD InvoiceId integer ")

					objConn.Execute("CREATE TABLE CRM_Invoices (InvoiceId integer IDENTITY (1, 1) NOT NULL PRIMARY KEY, ContactId integer NULL DEFAULT 0, DivId integer NULL DEFAULT 0, I_Attach smallint NULL DEFAULT 0, I_Closed tinyint NULL DEFAULT 0, I_CreatedBy integer NULL DEFAULT 0, I_CreatedDate datetime NULL , I_Currency varchar(3) NULL , I_DueDate datetime NULL , I_Events smallint NULL DEFAULT 0, I_ModBy integer NULL DEFAULT 0, I_ModDate datetime NULL , I_Notes smallint NULL DEFAULT 0, I_Owner integer NULL DEFAULT 0, I_PaidDate datetime NULL , I_PayInfo LongText NULL , I_Phase smallint NULL DEFAULT 0, I_Projects smallint NULL DEFAULT 0, I_PurchaseOrder varchar(25) NULL , I_Received tinyint NULL DEFAULT 0, I_Sales smallint NULL DEFAULT 0, I_SendDate datetime NULL , I_Serials smallint NULL DEFAULT 0, I_Status tinyint NULL DEFAULT 1, I_Tax double NULL DEFAULT 0, I_Type smallint NULL DEFAULT 0, I_Value double NULL DEFAULT 0) ")
					objConn.Execute("CREATE TABLE CRM_Portal (PortalId integer IDENTITY (1, 1) NOT NULL PRIMARY KEY, ReportId integer NULL DEFAULT 0, UserId integer NULL DEFAULT 0) ")

					objConn.Execute("CREATE INDEX InvoiceId ON CRM_Invoices(InvoiceId) ")
					objConn.Execute("CREATE INDEX PortalId ON CRM_Portal(PortalId) ")
					objConn.Execute("CREATE INDEX ReportId ON CRM_Portal(ReportId) ")

					Call setAppVar("av_DatabaseVersion",27)
				Case 27

					objConn.Execute("UPDATE CRM_Reports SET R_Fields = NULL, R_Params = NULL")

					objConn.Execute("UPDATE ALL_OptGroups SET G_Name = 'IDS_AccountType' WHERE G_Name = 'Account Type'")
					objConn.Execute("UPDATE ALL_OptGroups SET G_Name = 'IDS_BugCause' WHERE G_Name = 'Bug Cause'")
					objConn.Execute("UPDATE ALL_OptGroups SET G_Name = 'IDS_BugSource' WHERE G_Name = 'Bug Source'")
					objConn.Execute("UPDATE ALL_OptGroups SET G_Name = 'IDS_BugType' WHERE G_Name = 'Bug Type'")
					objConn.Execute("UPDATE ALL_OptGroups SET G_Name = 'IDS_NoteType' WHERE G_Name = 'Contact Type'")
					objConn.Execute("UPDATE ALL_OptGroups SET G_Name = 'IDS_AttachmentType' WHERE G_Name = 'Document Type'")
					objConn.Execute("UPDATE ALL_OptGroups SET G_Name = 'IDS_EventType' WHERE G_Name = 'Event Type'")
					objConn.Execute("UPDATE ALL_OptGroups SET G_Name = 'IDS_InvoicePhase' WHERE G_Name = 'Invoice Phase'")
					objConn.Execute("UPDATE ALL_OptGroups SET G_Name = 'IDS_InvoiceType' WHERE G_Name = 'Invoice Type'")
					objConn.Execute("UPDATE ALL_OptGroups SET G_Name = 'IDS_Priority' WHERE G_Name = 'Priority'")
					objConn.Execute("UPDATE ALL_OptGroups SET G_Name = 'IDS_ContactPrefix' WHERE G_Name = 'Prefix'")
					objConn.Execute("UPDATE ALL_OptGroups SET G_Name = 'IDS_SalesRegion' WHERE G_Name = 'Sales Region'")
					objConn.Execute("UPDATE ALL_OptGroups SET G_Name = 'IDS_IndustrySector' WHERE G_Name = 'Sales Vertical'")
					objConn.Execute("UPDATE ALL_OptGroups SET G_Name = 'IDS_SalesPhase' WHERE G_Name = 'Sales Phase'")
					objConn.Execute("UPDATE ALL_OptGroups SET G_Name = 'IDS_TicketSupport' WHERE G_Name = 'Support Type'")
					objConn.Execute("UPDATE ALL_OptGroups SET G_Name = 'IDS_TicketCause' WHERE G_Name = 'Ticket Cause'")
					objConn.Execute("UPDATE ALL_OptGroups SET G_Name = 'IDS_TicketSource' WHERE G_Name = 'Ticket Source'")
					objConn.Execute("UPDATE ALL_OptGroups SET G_Name = 'IDS_TicketType' WHERE G_Name = 'Ticket Type'")

					objConn.Execute("INSERT INTO ALL_OptGroups(G_Name, G_Module, G_Description, G_Required, G_Status) VALUES ('IDS_Product', 0, 'Products which the organization sells.',1,1)")
					lngProduct = getValue("OptGroupId","ALL_OptGroups","G_Name='IDS_Product'",0)

					objConn.Execute("ALTER TABLE ALL_CustomData ADD X_Field varchar(100), X_Default tinyint NULL DEFAULT 0")
					objConn.Execute("ALTER TABLE ALL_CustomData ALTER COLUMN X_Order tinyint")
					objConn.Execute("ALTER TABLE ALL_CustomData ALTER COLUMN X_Name varchar(100)")
					objConn.Execute("ALTER TABLE ALL_Users ALTER COLUMN U_Permissions varchar(10)")
					objConn.Execute("ALTER TABLE ALL_Users ALTER COLUMN U_Member varchar(10)")
					objConn.Execute("ALTER TABLE CRM_Contacts ADD K_DoNotCall tinyint, K_EmailOptOut tinyint")
					objConn.Execute("ALTER TABLE CRM_Divisions ADD D_Size smallint")
					objConn.Execute("ALTER TABLE KB_Articles ADD H_Link varchar(255)")

					objConn.Execute("ALTER TABLE CRM_Bugs ADD B_ProductId smallint")
					objConn.Execute("CREATE INDEX B_ProductId ON CRM_Bugs(B_ProductId)")

					objConn.Execute("ALTER TABLE CRM_Tickets ADD T_ProductId smallint")
					objConn.Execute("CREATE INDEX T_ProductId ON CRM_Tickets(T_ProductId)")

					objConn.Execute("ALTER TABLE CRM_Serialz ADD Z_ProductId smallint")
					objConn.Execute("CREATE INDEX Z_ProductId ON CRM_Serialz(Z_ProductId)")

					Set objRS = objConn.Execute("SELECT R.ProductId, R.R_Name, V.VersionId, V.V_Version FROM (CRM_Products R INNER JOIN CRM_ProdVersions V ON R.ProductId = V.ProductId) WHERE R_Status = 1 and V_Status = 1")
					If not (objRS.BOF and objRS.EOF) Then
						arrRS = objRS.GetRows()

						For i = 0 to UBound(arrRS)
							strTemp = arrRS(1,i) & " " & arrRS(3,i)
							objConn.Execute("INSERT INTO ALL_Options (OptGroupId, O_Value, O_Status) VALUES (" & lngProduct & "," & sqlText(strTemp) & ",1)")
							lngVersion = getValue("OptionId","ALL_Options","O_Value="&sqlText(strTemp),0)
							objConn.Execute("UPDATE CRM_Bugs SET B_ProductId = " & lngVersion & " WHERE VersionId = " & arrRS(2,i))
							objConn.Execute("UPDATE CRM_Tickets SET T_ProductId = " & lngVersion & " WHERE VersionId = " & arrRS(2,i))
							objConn.Execute("UPDATE CRM_Serialz SET Z_ProductId = " & lngVersion & " WHERE ProductId = " & arrRS(0,i))
						Next

					End If

					Set objRS = objConn.Execute("SELECT 'Need to drop last query from cache.'")

					objConn.Execute("DROP TABLE CRM_ProdVersions")
					objConn.Execute("UPDATE CRM_Serialz SET ProductId = NULL")
					objConn.Execute("DELETE FROM CRM_Products")

					objConn.Execute("DROP INDEX ProductId ON CRM_Bugs")
					objConn.Execute("ALTER TABLE CRM_Bugs DROP COLUMN ProductId")
					objConn.Execute("DROP INDEX VersionId ON CRM_Bugs")
					objConn.Execute("ALTER TABLE CRM_Bugs DROP COLUMN VersionId")

					objConn.Execute("DROP INDEX ProductId ON CRM_Tickets")
					objConn.Execute("ALTER TABLE CRM_Tickets DROP COLUMN ProductId")
					objConn.Execute("DROP INDEX VersionId ON CRM_Tickets")
					objConn.Execute("ALTER TABLE CRM_Tickets DROP COLUMN VersionId")

					objConn.Execute("ALTER TABLE CRM_Reports DROP COLUMN R_CreatedBefore")
					objConn.Execute("ALTER TABLE CRM_Reports DROP COLUMN R_CreatedAfter")
					objConn.Execute("ALTER TABLE CRM_Reports DROP COLUMN R_CreatedUser")
					objConn.Execute("ALTER TABLE CRM_Reports DROP COLUMN R_ModBefore")
					objConn.Execute("ALTER TABLE CRM_Reports DROP COLUMN R_ModAfter")
					objConn.Execute("ALTER TABLE CRM_Reports DROP COLUMN R_ModUser")

					Call insertCustomData("1,1,'IDS_Contact','K.ContactId',2,3,1,1,1")
					Call insertCustomData("1,1,'IDS_Account','K.DivId',2,3,1,2,-1")
					Call insertCustomData("1,1,NULL,'K.K_Status',4,17,0,3,-1")
					Call insertCustomData("1,1,'IDS_SalesRegion','D.D_Region',5,2,0,4,1")
					Call insertCustomData("1,1,'IDS_ContactPrefix','K.K_Prefix',5,2,0,5,1")
					Call insertCustomData("1,1,'IDS_NameFirst','K.K_FirstName',1,30,0,6,1")
					Call insertCustomData("1,1,'IDS_NameMI','K.K_MiddleInitial',1,1,0,7,1")
					Call insertCustomData("1,1,'IDS_NameLast','K.K_LastName',1,30,0,8,1")
					Call insertCustomData("1,1,NULL,'K.K_UserName',1,20,0,9,-1")
					Call insertCustomData("1,1,NULL,'K.K_Password',7,35,0,10,-1")
					Call insertCustomData("1,1,'IDS_Address1','K.K_Address1',1,60,0,11,1")
					Call insertCustomData("1,1,'IDS_Address2','K.K_Address2',1,60,0,12,-1")
					Call insertCustomData("1,1,'IDS_Address3','K.K_Address3',1,60,0,13,-1")
					Call insertCustomData("1,1,'IDS_City','K.K_City',1,20,0,14,1")
					Call insertCustomData("1,1,'IDS_State','K.K_State',1,2,0,15,1")
					Call insertCustomData("1,1,'IDS_Country','K.K_Country',1,20,0,16,1")
					Call insertCustomData("1,1,'IDS_ZIP','K.K_ZIP',1,7,0,17,1")
					Call insertCustomData("1,1,'IDS_Department','K.K_Dept',1,40,0,18,1")
					Call insertCustomData("1,1,'IDS_JobTitle','K.K_JobTitle',1,30,0,19,1")
					Call insertCustomData("1,1,'IDS_Email','K.K_Email',6,255,0,20,1")
					Call insertCustomData("1,1,'IDS_EmailOptOut','K.K_EmailOptOut',4,17,0,21,1")
					Call insertCustomData("1,1,'IDS_Phone1','K.K_Phone1',2,5,0,22,1")
					Call insertCustomData("1,1,'IDS_Ext1','K.K_Ext1',2,2,0,23,1")
					Call insertCustomData("1,1,'IDS_Phone2','K.K_Phone2',2,5,0,24,1")
					Call insertCustomData("1,1,'IDS_Ext2','K.K_Ext2',2,2,0,25,1")
					Call insertCustomData("1,1,'IDS_Fax','K.K_Fax',2,5,0,26,1")
					Call insertCustomData("1,1,'IDS_DoNotCall','K.K_DoNotCall',4,17,0,27,1")
					Call insertCustomData("1,1,NULL,'K.K_ReportsTo',2,3,0,28,-1")
					Call insertCustomData("1,1,NULL,'K.K_Assistant',2,3,0,29,-1")
					Call insertCustomData("1,1,'IDS_Notes','K.K_Notes',2,2,0,30,-1")
					Call insertCustomData("1,1,'IDS_Attachments','K.K_Attach',2,2,0,31,-1")
					Call insertCustomData("1,1,'IDS_Events','K.K_Events',2,2,0,32,-1")
					Call insertCustomData("1,1,'IDS_Sales','K.K_Sales',2,2,0,33,-1")
					Call insertCustomData("1,1,'IDS_Serials','K.K_Serials',2,2,0,34,-1")
					Call insertCustomData("1,1,'IDS_Tickets','K.K_Tickets',2,2,0,35,-1")
					Call insertCustomData("1,1,'IDS_Invoices','K.K_Invoices',2,2,0,36,-1")
					Call insertCustomData("1,1,'IDS_CreatedBy','K.K_CreatedBy',2,3,0,37,-1")
					Call insertCustomData("1,1,'IDS_Created','K.K_CreatedDate',3,7,0,38,-1")
					Call insertCustomData("1,1,'IDS_ModifiedBy','K.K_ModBy',2,3,0,39,-1")
					Call insertCustomData("1,1,'IDS_Modified','K.K_ModDate',3,7,0,40,-1")
					Call insertCustomData("1,2,'IDS_Account','D.DivId',2,3,1,1,1")
					Call insertCustomData("1,2,NULL,'D.ClientId',2,3,1,2,-1")
					Call insertCustomData("1,2,NULL,'D.D_Status',4,17,0,3,-1")
					Call insertCustomData("1,2,'IDS_Division','D.D_Division',1,40,0,4,1")
					Call insertCustomData("1,2,'IDS_AccountId','D.D_Account',1,25,0,5,1")
					Call insertCustomData("1,2,'IDS_Website','D.D_Website',1,255,0,6,1")
					Call insertCustomData("1,2,'IDS_SalesRep','D.D_SalesRep',2,3,0,7,1")
					Call insertCustomData("1,2,'IDS_SalesRegion','D.D_Region',5,2,0,8,1")
					Call insertCustomData("1,2,'IDS_AccountType','D.D_AccountType',5,2,0,9,1")
					Call insertCustomData("1,2,'IDS_IndustrySector','D.D_Vertical',5,2,0,10,1")
					Call insertCustomData("1,2,'IDS_AccountSize','D.D_Size',5,2,0,11,1")
					Call insertCustomData("1,2,'IDS_Reference','D.D_RefAccount',4,17,0,12,1")
					Call insertCustomData("1,2,'IDS_ProblemFlag','D.D_ProbFlag',4,17,0,13,1")
					Call insertCustomData("1,2,'IDS_Description','D.D_ShortDesc',1,255,0,14,1")
					Call insertCustomData("1,2,'IDS_Notes','D.D_Notes',2,2,0,15,-1")
					Call insertCustomData("1,2,'IDS_Attachments','D.D_Attach',2,2,0,16,-1")
					Call insertCustomData("1,2,'IDS_Events','D.D_Events',2,2,0,17,-1")
					Call insertCustomData("1,2,'IDS_Contacts','D.D_Contacts',2,2,0,18,-1")
					Call insertCustomData("1,2,'IDS_Sales','D.D_Sales',2,2,0,19,-1")
					Call insertCustomData("1,2,'IDS_Serials','D.D_Serials',2,2,0,20,-1")
					Call insertCustomData("1,2,'IDS_Projects','D.D_Projects',2,2,0,21,-1")
					Call insertCustomData("1,2,'IDS_Tickets','D.D_Tickets',2,2,0,22,-1")
					Call insertCustomData("1,2,'IDS_Invoices','D.D_Invoices',2,2,0,23,-1")
					Call insertCustomData("1,2,'IDS_CreatedBy','D.D_CreatedBy',2,3,0,24,-1")
					Call insertCustomData("1,2,'IDS_Created','D.D_CreatedDate',3,7,0,25,-1")
					Call insertCustomData("1,2,'IDS_ModifiedBy','D.D_ModBy',2,3,0,26,-1")
					Call insertCustomData("1,2,'IDS_Modified','D.D_ModDate',3,7,0,27,-1")
					Call insertCustomData("1,3,'IDS_Sale','S.SaleId',2,3,1,1,1")
					Call insertCustomData("1,3,'IDS_Contact','S.ContactId',2,3,1,2,1")
					Call insertCustomData("1,3,'IDS_Account','S.DivId',2,3,1,3,1")
					Call insertCustomData("1,3,'IDS_SalesRegion','D.D_Region',5,2,0,4,1")
					Call insertCustomData("1,3,'IDS_IndustrySector','D.D_Vertical',5,2,0,5,1")
					Call insertCustomData("1,3,'IDS_Invoice','S.InvoiceId',2,3,0,6,1")
					Call insertCustomData("1,3,NULL,'S.S_Status',4,17,0,7,-1")
					Call insertCustomData("1,3,'IDS_SalesPhase','S.S_Phase',5,2,0,8,1")
					Call insertCustomData("1,3,'IDS_Pipeline','S.S_Pipe',2,17,0,9,1")
					Call insertCustomData("1,3,'IDS_SalesRep','S.S_SalesRep',2,3,0,10,1")
					Call insertCustomData("1,3,'IDS_Currency','S.S_Currency',1,3,0,11,1")
					Call insertCustomData("1,3,'IDS_SaleValue','S.S_SaleValue',2,6,0,12,1")
					Call insertCustomData("1,3,'IDS_Closed','S.S_Closed',4,17,0,13,1")
					Call insertCustomData("1,3,'IDS_CloseDate','S.S_CloseDate',3,7,0,14,1")
					Call insertCustomData("1,3,'IDS_Notes','S.S_Notes',2,2,0,15,-1")
					Call insertCustomData("1,3,'IDS_Attachments','S.S_Attach',2,2,0,16,-1")
					Call insertCustomData("1,3,'IDS_Events','S.S_Events',2,2,0,17,-1")
					Call insertCustomData("1,3,'IDS_CreatedBy','S.S_CreatedBy',2,3,0,18,-1")
					Call insertCustomData("1,3,'IDS_Created','S.S_CreatedDate',3,7,0,19,-1")
					Call insertCustomData("1,3,'IDS_ModifiedBy','S.S_ModBy',2,3,0,20,-1")
					Call insertCustomData("1,3,'IDS_Modified','S.S_ModDate',3,7,0,21,-1")
					Call insertCustomData("1,4,'IDS_Project','P.ProjectId',2,3,1,1,1")
					Call insertCustomData("1,4,'IDS_Account','P.DivId',2,3,1,2,1")
					Call insertCustomData("1,4,'IDS_Invoice','P.InvoiceId',2,3,0,3,1")
					Call insertCustomData("1,4,NULL,'P.P_Status',4,17,0,4,-1")
					Call insertCustomData("1,4,'IDS_Owner','P.P_Owner',2,3,1,5,1")
					Call insertCustomData("1,4,NULL,'P.P_Title',1,40,0,6,-1")
					Call insertCustomData("1,4,'IDS_Description','P.P_ShortDesc',1,255,0,7,1")
					Call insertCustomData("1,4,'IDS_DaysTotal','P.P_DaysTotal',2,2,0,8,1")
					Call insertCustomData("1,4,'IDS_DaysOwed','P.P_DaysOwed',2,2,0,9,1")
					Call insertCustomData("1,4,'IDS_CreatedBy','P.P_CreatedBy',2,3,0,10,-1")
					Call insertCustomData("1,4,'IDS_Created','P.P_CreatedDate',3,7,0,11,-1")
					Call insertCustomData("1,4,'IDS_Notes','P.P_Notes',2,2,0,12,-1")
					Call insertCustomData("1,4,'IDS_Attachments','P.P_Attach',2,2,0,13,-1")
					Call insertCustomData("1,4,'IDS_Events','P.P_Events',2,2,0,14,-1")
					Call insertCustomData("1,4,'IDS_ModifiedBy','P.P_ModBy',2,3,0,15,-1")
					Call insertCustomData("1,4,'IDS_Modified','P.P_ModDate',3,7,0,16,-1")
					Call insertCustomData("1,4,'IDS_Closed','P.P_Closed',4,17,0,17,-1")
					Call insertCustomData("1,4,'IDS_CloseDate','P.P_CloseDate',3,7,0,18,-1")
					Call insertCustomData("1,5,'IDS_Ticket','T.TicketId',2,3,0,1,1")
					Call insertCustomData("1,5,'IDS_Account','T.DivId',2,3,0,2,1")
					Call insertCustomData("1,5,'IDS_Contact','T.ContactId',2,3,0,3,1")
					Call insertCustomData("1,5,NULL,'T.T_Status',4,17,0,4,1")
					Call insertCustomData("1,5,'IDS_Owner','T.T_Owner',2,3,0,5,1")
					Call insertCustomData("1,5,'IDS_HotIssue','T.T_HotIssue',4,17,0,6,1")
					Call insertCustomData("1,5,'IDS_Priority','T.T_Priority',5,2,0,7,1")
					Call insertCustomData("1,5,'IDS_TicketType','T.T_TicketType',5,2,0,8,1")
					Call insertCustomData("1,5,'IDS_TicketSource','T.T_TicketSource',5,2,0,9,1")
					Call insertCustomData("1,5,'IDS_TicketSupport','T.T_SupportType',5,2,0,10,1")
					Call insertCustomData("1,5,'IDS_Product','T.T_ProductId',5,2,0,11,1")
					Call insertCustomData("1,5,'IDS_Build','T.T_Build',1,10,0,12,1")
					Call insertCustomData("1,5,'IDS_Bug','T.T_BugId',2,3,0,13,1")
					Call insertCustomData("1,5,'IDS_Description','T.T_Description',1,255,0,14,1")
					Call insertCustomData("1,5,'IDS_Solution','T.T_Solution',1,255,0,15,1")
					Call insertCustomData("1,5,'IDS_TicketCause','T.T_Cause',5,2,0,16,1")
					Call insertCustomData("1,5,'IDS_Closed','T.T_Closed',4,17,0,17,1")
					Call insertCustomData("1,5,'IDS_CloseDate','T.T_CloseDate',3,7,0,18,1")
					Call insertCustomData("1,5,'IDS_Notes','T.T_Notes',2,2,0,19,1")
					Call insertCustomData("1,5,'IDS_Attachments','T.T_Attach',2,2,0,20,1")
					Call insertCustomData("1,5,'IDS_Events','T.T_Events',2,2,0,21,1")
					Call insertCustomData("1,5,'IDS_CreatedBy','T.T_CreatedBy',2,3,0,22,1")
					Call insertCustomData("1,5,'IDS_Created','T.T_Created',3,7,0,23,1")
					Call insertCustomData("1,5,'IDS_ModifiedBy','T.T_ModBy',2,3,0,24,1")
					Call insertCustomData("1,5,'IDS_Modified','T.T_ModDate',3,7,0,25,1")
					Call insertCustomData("1,6,'IDS_Bug','B.BugId',2,3,1,1,-1")
					Call insertCustomData("1,6,NULL,'B.B_Status',4,17,0,2,-1")
					Call insertCustomData("1,6,'IDS_Owner','B.B_Owner',2,3,1,3,1")
					Call insertCustomData("1,6,'IDS_HotIssue','B.B_HotIssue',4,17,0,4,1")
					Call insertCustomData("1,6,'IDS_Priority','B.B_Priority',5,2,0,5,1")
					Call insertCustomData("1,6,'IDS_BugType','B.B_BugType',5,2,0,6,1")
					Call insertCustomData("1,6,'IDS_BugSource','B.B_BugSource',5,2,0,7,1")
					Call insertCustomData("1,6,'IDS_Product','B.B_ProductId',5,2,0,8,1")
					Call insertCustomData("1,6,'IDS_Build','B.B_Build',1,10,0,9,1")
					Call insertCustomData("1,6,'IDS_Description','B.B_Description',1,255,0,10,1")
					Call insertCustomData("1,6,'IDS_Solution','B.B_Solution',1,255,0,11,1")
					Call insertCustomData("1,6,'IDS_BugCause','B.B_Cause',5,2,0,12,1")
					Call insertCustomData("1,6,'IDS_Notes','B.B_Notes',2,2,0,13,-1")
					Call insertCustomData("1,6,'IDS_Attachments','B.B_Attach',2,2,0,14,-1")
					Call insertCustomData("1,6,'IDS_Events','B.B_Events',2,2,0,15,-1")
					Call insertCustomData("1,6,'IDS_Tickets','B.B_Tickets',2,2,0,16,-1")
					Call insertCustomData("1,6,'IDS_CreatedBy','B.B_CreatedBy',2,3,0,17,-1")
					Call insertCustomData("1,6,'IDS_Created','B.B_CreatedDate',3,7,0,18,-1")
					Call insertCustomData("1,6,'IDS_ModifiedBy','B.B_ModBy',2,3,0,19,-1")
					Call insertCustomData("1,6,'IDS_Modified','B.B_ModDate',3,7,0,20,-1")
					Call insertCustomData("1,6,'IDS_Closed','B.B_Closed',4,17,0,21,1")
					Call insertCustomData("1,6,'IDS_CloseDate','B.B_CloseDate',3,7,0,22,1")
					Call insertCustomData("1,7,'IDS_Invoice','I.InvoiceId',2,3,1,1,1")
					Call insertCustomData("1,7,'IDS_Contact','I.ContactId',2,3,1,2,1")
					Call insertCustomData("1,7,'IDS_Account','I.DivId',2,3,1,3,1")
					Call insertCustomData("1,7,NULL,'I.I_Status',4,17,0,4,-1")
					Call insertCustomData("1,7,'IDS_Owner','I.I_Owner',2,3,1,5,1")
					Call insertCustomData("1,7,'IDS_PurchaseOrder','I.I_PurchaseOrder',1,25,0,6,1")
					Call insertCustomData("1,7,'IDS_InvoiceReceived','I.I_Received',4,17,0,7,1")
					Call insertCustomData("1,7,'IDS_InvoiceType','I.I_Type',5,2,0,8,1")
					Call insertCustomData("1,7,'IDS_Currency','I.I_Currency',1,3,0,9,1")
					Call insertCustomData("1,7,'IDS_Value','I.I_Value',2,5,0,10,1")
					Call insertCustomData("1,7,'IDS_Tax','I.I_Tax',2,5,0,11,1")
					Call insertCustomData("1,7,'IDS_InvoicePhase','I.I_Phase',5,2,0,12,1")
					Call insertCustomData("1,7,'IDS_InvoiceDetails','I.I_PayInfo',1,255,0,13,1")
					Call insertCustomData("1,7,'IDS_InvoiceSent','I.I_SendDate',3,7,0,14,1")
					Call insertCustomData("1,7,'IDS_InvoiceDue','I.I_DueDate',3,7,0,15,1")
					Call insertCustomData("1,7,'IDS_InvoicePaid','I.I_PaidDate',3,7,0,16,1")
					Call insertCustomData("1,7,'IDS_Closed','I.I_Closed',4,17,0,17,1")
					Call insertCustomData("1,7,'IDS_Notes','I.I_Notes',2,2,0,18,-1")
					Call insertCustomData("1,7,'IDS_Attachments','I.I_Attach',2,2,0,19,-1")
					Call insertCustomData("1,7,'IDS_Events','I.I_Events',2,2,0,20,-1")
					Call insertCustomData("1,7,'IDS_Sales','I.I_Sales',2,2,0,21,-1")
					Call insertCustomData("1,7,'IDS_Serials','I.I_Serials',2,2,0,22,-1")
					Call insertCustomData("1,7,'IDS_Projects','I.I_Projects',2,2,0,23,-1")
					Call insertCustomData("1,7,'IDS_Created','I.I_CreatedDate',3,7,0,24,-1")
					Call insertCustomData("1,7,'IDS_CreatedBy','I.I_CreatedBy',2,3,0,25,-1")
					Call insertCustomData("1,7,'IDS_Modified','I.I_ModDate',3,7,0,26,-1")
					Call insertCustomData("1,7,'IDS_ModifiedBy','I.I_ModBy',2,3,0,27,-1")
					Call insertCustomData("1,50,'IDS_Event','E.EventId',2,3,1,1,1")
					Call insertCustomData("1,50,NULL,'E.E_Module',2,17,1,2,-1")
					Call insertCustomData("1,50,NULL,'E.E_ModuleId',2,3,1,3,-1")
					Call insertCustomData("1,50,NULL,'E.E_Status',4,17,0,4,-1")
					Call insertCustomData("1,50,'IDS_Owner','E.E_Owner',2,3,1,5,1")
					Call insertCustomData("1,50,NULL,'E.E_Permissions',2,17,0,6,-1")
					Call insertCustomData("1,50,NULL,'E.E_Title',1,40,0,7,-1")
					Call insertCustomData("1,50,'IDS_Onsite','E.E_Onsite',4,17,0,8,1")
					Call insertCustomData("1,50,'IDS_Billable','E.E_Billable',4,17,0,9,1")
					Call insertCustomData("1,50,'IDS_EventType','E.E_EventType',5,2,0,10,1")
					Call insertCustomData("1,50,'IDS_StartTime','E.E_StartTime',3,7,0,11,1")
					Call insertCustomData("1,50,'IDS_EndTime','E.E_EndTime',3,7,0,12,1")
					Call insertCustomData("1,50,'IDS_CreatedBy','E.E_CreatedBy',2,3,0,13,-1")
					Call insertCustomData("1,50,'IDS_Created','E.E_CreatedDate',3,7,0,14,-1")
					Call insertCustomData("1,50,'IDS_ModifiedBy','E.E_ModBy',2,3,0,15,-1")
					Call insertCustomData("1,50,'IDS_Modified','E.E_ModDate',3,7,0,16,-1")

					Call setAppVar("av_DatabaseVersion",28)

					Application.Contents.RemoveAll()
					Call doRedirect("../default.asp")

				Case 28
					Application.Contents.RemoveAll()
					Call doRedirect("../default.asp")
				Case Else
					strMessage = getIDS("IDS_MsgUpgradeUnknown")
			End Select
		Else
			strMessage = Replace(Replace(getIDS("IDS_MsgUpgradeNeeded"),"[X]",Application("av_DatabaseVersion")),"[Y]",intDBVer)
		End If

		If strMessage = "" Then Call doRedirect("upgrade.asp")
	Else
		strMessage = getIDS("IDS_MsgUpgradeLocal")
	End If

	Call DisplayHeader(3)
%>

<table border=0 width="100%" height="90%"><tr><td valign="middle" align="center">

<form name="frmUpgrade" method="post" action="upgrade.asp">

<table border=0 cellspacing=10 cellpadding=0 width=300 height=150 class="hRow">
  <tr><td class="wFont"><% =strMessage %></td></tr>
  <tr><td align="center"><% =getSubmit("btnSubmit",getIDS("IDS_Upgrade"),100,"S","") %></td></tr>
</table>

</form>

</td></tr></table>

<%
	Call DisplayFooter(1)
%>