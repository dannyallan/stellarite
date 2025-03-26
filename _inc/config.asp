<%
	On Error Resume Next

	Dim strDatabase         'as String      'Database Type
	Dim strConnString       'as String      'DSN Connection String
	Dim strCRMURL           'as String      'URL of the CRM
	Dim strLanguage         'as String      'Language of the CRM
	Dim intMode             'as Integer     'CRM Maintenance


	'#########################################################################
	' Database Connection Types
	'#########################################################################
	'strDatabase = "MSSQL"
	'strConnString = "Driver={SQL Server};Server=(local);Database=Stellarite;UID=crm;PWD=crm;"

	strDatabase = "Access"
	strConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Temp\demo.mdb;"

	'strDatabase = "MySQL"
	'strConnString  = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=crm;UID=root;PWD=crm;"

	'strDatabase = "Oracle"
	'strConnString  = "Driver={Microsoft ODBC for Oracle};Server=OracleServer.Stellarite;UID=crm;PWD=crm"


	'#########################################################################
	' CRM URL
	'  -- Virtual directory can be used
	'  -- Must follow the directory with forward slash
	'#########################################################################
	strCRMURL    = "http://demo.stellarite.com/"


	'#########################################################################
	' Language Code used for the strings file
	'  -- Please edit the string files in the _inc directory
	'  -- You will also want to translate text from the following files
	'        -- \faq.asp
	'        -- \require.asp
	'        -- \common\js\crm.js
	'#########################################################################
	strLanguage     = "en"        'English
	'strLanguage    = "fr"        'French
	'strLanguage    = "sp"        'Spanish
	'strLanguage    = "de"        'German
	'strLanguage    = "it"        'Italian


	'#########################################################################
	' Locale ID
	' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/script56/html/vsmscLCID.asp
	'#########################################################################
	Session.LCID = 1033     '1033 - English (United States)
							'4105 - English (Canada)
							'2057 - English (United Kingdom)


	'#########################################################################
	' CRM Maintenance
	'#########################################################################
	intMode = 1             '0 - CRM is fully functioning
							'1 - CRM will not allow data changes
							'2 - CRM is down for maintenance
%>

