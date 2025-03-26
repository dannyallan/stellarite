<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_edit.asp" -->
<!--#include file="..\_inc\sql\sql_reports.asp" -->
<%
	Call pageFunctions(0,1)

	Dim strReportName   'as String      'Report Name
	Dim intPermissions  'as Integer     'Permissions on the Report
	Dim strOwner        'as String      'Owner
	Dim strCreatedBy    'as String      'Created By
	Dim datCreatedDate  'as Date        'Created Date
	Dim strModBy        'as String      'Modified By
	Dim datModDate      'as Date        'Modified Date
	Dim strGenSQL       'as String      'Generated SQL Query
	Dim bytType         'as Byte        'Type of Report
	Dim blnChange       'as Boolean     'Has the report been modified?
	Dim strFields       'as String      'String value of selected fields
	Dim arrFields       'as Array       'Array for Fields in report
	Dim strAvlFlds      'as String      'String for available fields
	Dim arrDefFlds      'as Array       'Array for default available fields
	Dim arrAvlFlds      'as Array       'Array for available fields
	Dim strParams       'as String      'String value for report filters
	Dim bytOrder        'as Byte        'Byte for selected column to order by
	Dim strOrderDir     'as String      'Ascending or Descending
	Dim strTemp         'as String      'Temporary string
	Dim arrTemp         'as Array       'Temporary array
	Dim strOptStr1      'as String      'String for building multi-select option fields
	Dim strOptVal1      'as String      'Array for building multi-select option fields
	Dim strOptStr2      'as String      'String for building multi-select option fields
	Dim strOptVal2      'as String      'Array for building multi-select option fields

	strTitle = getIDS("IDS_ReportNew")
	strDir = Application("av_CRMDir") & "reports/"
	strModName = getIDS("IDS_Reports")
	strModImage = "report"

	blnChange = valNum(Request.Form("hdnChange"),0,0)
	bytType = valNum(Request.QueryString("type"),1,0)
	If bytType = "" Then bytType = bytMod
	If blnAdmin Then intPerm = 5

	If blnChange = 1 Then

		strReportName = valString(Request.Form("txtReportName"),100,0,0)
		intPermissions = valNum(Request.Form("selPermissions"),1,1)
		strOwner = valString(Request.Form("txtOwner"),100,0,0)
		strGenSQL = Replace(Request.Form("txtSQL"),"Chr(34)",Chr(34))
		strParams = Request.Form("hdnParams")
		strFields = Request.Form("selFields")
		bytOrder = valNum(Request.Form("selOrder"),1,0)
		strOrderDir = valString(Request.Form("selOrderDir"),4,0,0)

	Elseif blnRS Then

		Set objRS = objConn.Execute(getReport(lngRecordId))

		If not (objRS.BOF and objRS.EOF) then

			strReportName = objRS.fields("R_Title").value
			intPermissions = objRS.fields("R_Permissions").value
			strOwner = objRS.fields("Owner").value
			bytMod = objRS.fields("R_Module").value
			bytType = objRS.fields("R_Type").value
			strGenSQL = objRS.fields("R_SQL").value
			strFields = objRS.fields("R_Fields").value
			strParams = objRS.fields("R_Params").value
			bytOrder = objRS.fields("R_Order").value
			strOrderDir = objRS.fields("R_OrderDir").value
			strCreatedBy = objRS.fields("CreatedBy").value
			datCreatedDate = objRS.fields("R_CreatedDate").value
			strModBy = objRS.fields("ModBy").value
			datModDate = objRS.fields("R_ModDate").value

			If bytType = 90 and not blnAdmin Then Call logError(2,1)

		Else
			Call logError(3,1)
		End If
	Else
		strOwner = strFullName
		strCreatedBy = strFullName
		datCreatedDate = Now
		strModBy = strFullName
		datModDate = Now
	End If

	strIncHead = "<script language=""JavaScript"" type=""text/javascript"" src=""" & Application("av_CRMDir") & "common/js/rpt.js""></script>" & vbCrLf & getCalendarScripts()

	Call DisplayHeader(1)
	Call showEditHeader(strTitle,strCreatedBy,datCreatedDate,strModBy,datModDate)

	If bytType <> 0 Then Response.Write("<form name=""frmReport"" method=""post"" action=""view_report.asp?id=" & lngRecordId & "&m=" & bytMod & "&type=" & bytType & """>" & vbCrLf)
	Response.Write("<div id=""contentDiv"" class=""dvBorder"" style=""height:" & intScreenH-170 & "px;"">" & vbCrLf)

	If bytType = 0 Then

		Response.Write("<ul>" & vbCrLf)
		If pContacts >= 1 Then Response.Write("<li class=""dFont""><a href=""edit_report.asp?m=1&type=1"">" & getIDS("IDS_ReportNewContact") & "</a></li>" & vbCrLf)
		If pClients >= 1 Then Response.Write("<li class=""dFont""><a href=""edit_report.asp?m=2&type=2"">" & getIDS("IDS_ReportNewClient") & "</a></li>" & vbCrLf)
		If pSales >= 1 Then Response.Write("<li class=""dFont""><a href=""edit_report.asp?m=3&type=3"">" & getIDS("IDS_ReportNewSales") & "</a></li>" & vbCrLf)
		If pProjects >= 1 Then Response.Write("<li class=""dFont""><a href=""edit_report.asp?m=4&type=4"">" & getIDS("IDS_ReportNewProject") & "</a></li>" & vbCrLf & _
							"<li class=""dFont""><a href=""edit_report.asp?m=4&type=50"">" & getIDS("IDS_ReportNewEvent") & "</a></li>" & vbCrLf)
		If pTickets >= 1 Then Response.Write("<li class=""dFont""><a href=""edit_report.asp?m=5&type=5"">" & getIDS("IDS_ReportNewTicket") & "</a></li>" & vbCrLf)
		If pBugs >= 1 Then Response.Write("<li class=""dFont""><a href=""edit_report.asp?m=6&type=6"">" & getIDS("IDS_ReportNewBug") & "</a></li>" & vbCrLf)
		If pInvoices >= 1 Then Response.Write("<li class=""dFont""><a href=""edit_report.asp?m=7&type=7"">" & getIDS("IDS_ReportNewInvoice") & "</a></li>" & vbCrLf)
		Response.Write(    "<br /><br />" & vbCrLf)
		If blnAdmin Then Response.Write("<li class=""dFont""><a href=""edit_report.asp?type=90"">" & getIDS("IDS_ReportNewCustom") & "</a></li>" & vbCrLf)
		Response.Write("</ul></div>" & vbCrLf)
	Else
%>

<table border=0 cellpadding=5 width="100%">
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnParams",strParams) %>
<% =getHidden("hdnChange",blnChange) %>
  <tr>
	<td><% =getLabel(getIDS("IDS_ReportName"),"txtReportName") %></td>
	<td><% =getTextField("txtReportName","mText",strReportName,20,100,"") %></td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_Owner"),"txtOwner") %></td>
	<td>
<%
		If intPerm = 5 Then
			Response.Write(getTextField("txtOwner","mText",strOwner,20,255,""))
			Response.Write(getIconImport(1,getSearchURL("?m=0&rVal=txtOwner"),getIDS("IDS_Owner")))
		Else
			Response.Write(getTextField("txtOwner","dText",strOwner,20,20,"readonly=""readonly"""))
		End If
%>
	</td>
  </tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_Permissions"),"selPermissions") %></td>
	<td>
	  <select name="selPermissions" id="selPermissions" class="oByte" style="width:140px;" onChange="doChange();">
		<option value="1"<% =getDefault(0,intPermissions,1) & ">" & getIDS("IDS_AccessPrivate") %></option>
		<% If intPerm >= 5 Then %>
		<option value="2"<% =getDefault(0,intPermissions,2) & ">" & getIDS("IDS_MembersOnly") %></option>
		<option value="3"<% =getDefault(0,intPermissions,3) & ">" & getIDS("IDS_InternalView") %></option>
		<% End If %>
	  </select>
	</td>
  </tr>
  </tr>
  <tr><td colspan=2><hr /></td></tr>
<%
		If bytType = 90 Then
%>
  <tr>
	<td colspan=2><% =getLabel(getIDS("IDS_CustomSQL"),"txtSQL") %><br />
	<% =getTextArea("txtSQL","oText",strGenSQL,"100%",8,"") %>
	<br /><br /><% =getLabel(getIDS("IDS_Module"),"selModule") %> &nbsp;&nbsp;&nbsp;&nbsp;
	<% =getModuleDropDown("selModule",bytMod,True,"") %></td>
  </tr>
<%
		Else
			arrFields = Application("arr_Fields" & CStr(bytType))
			If strFields = "" Then strFields = arrFields(0,0)
%>
  <tr>
    <td class="bFont" valign="top"><% =getIDS("IDS_ReportColumns") %></td>
    <td>
      <table border=0>
        <tr>
          <td>
            <% =getLabel(getIDS("IDS_FieldsAvailable"),"selAvailable") %><br />
            <select name="selAvailable" id="selAvailable" class="oByte" style="width: 200px;" multiple size=10 onDblClick="moveSelectedOptions(this,this.form.selFields);">
<%
			For i = 0 to UBound(arrFields,2)
				Response.Write("<option value=""" & arrFields(0,i) & """>" & getIDS(arrFields(1,i)) & "</option>" & vbCrLf)
			Next
%>
            </select>
          </td>
          <td>
<%
			Response.Write(getSubmit("btnAddAll",getIDS("IDS_AddAll"),80,"","onClick=""moveAllOptions(this.form.selAvailable,this.form.selFields);return false;""") & "<br />" & vbCrLf & _
					getSubmit("btnAdd",getIDS("IDS_Add"),80,"","onClick=""moveSelectedOptions(this.form.selAvailable,this.form.selFields);return false;""") & "<br />" & vbCrLf & _
					getSubmit("btnRem",getIDS("IDS_Remove"),80,"","onClick=""moveSelectedOptions(this.form.selFields,this.form.selAvailable);return false;""") & "<br />" & vbCrLf & _
					getSubmit("btnRemAll",getIDS("IDS_RemoveAll"),80,"","onClick=""moveAllOptions(this.form.selFields,this.form.selAvailable);return false;"""))
%>
          </td>
          <td>
            <% =getLabel(getIDS("IDS_FieldsShow"),"selFields") %><br />
            <select name="selFields" id="selFields" class="oByte" style="width: 200px;" multiple size=10 onDblClick="moveSelectedOptions(this,this.form.selAvailable);">

            </select>
          </td>
          <td>
<%
			Response.Write(getSubmit("btnMoveTop",getIDS("IDS_MoveTop"),80,"","onClick=""moveOptionTop(this.form.selFields);return false;""") & "<br />" & vbCrLf & _
					getSubmit("btnMoveUp",getIDS("IDS_MoveUp"),80,"","onClick=""moveOptionUp(this.form.selFields);return false;""") & "<br />" & vbCrLf & _
					getSubmit("btnMoveDown",getIDS("IDS_MoveDown"),80,"","onClick=""moveOptionDown(this.form.selFields);return false;""") & "<br />" & vbCrLf & _
					getSubmit("btnMoveBottom",getIDS("IDS_MoveBottom"),80,"","onClick=""moveOptionBottom(this.form.selFields);return false;"""))
%>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr><td colspan=2><hr /></td></tr>
  <tr>
    <td class="bFont" valign="top"><% =getIDS("IDS_ReportConditions") %></td>
    <td>
      <table border=0 width="100%">
        <tr>
          <td width="25%" valign="top">
            <% =getLabel(getIDS("IDS_ReportField"),"selField") %><br />
            <select name="selField" id="selField" class="oByte" style="width: 140px;" onChange="setConditions(this.selectedIndex-1,this.value);">
              <option></option>
<%
			For i = 0 to UBound(arrFields,2)
				Response.Write("<option value=""" & arrFields(0,i) & """>" & getIDS(arrFields(1,i)) & "</option>" & vbCrLf)
			Next
%>
            </select>
          </td>
          <td width="25%" valign="top">
            <% =getLabel(getIDS("IDS_ReportCondition"),"selCondition") %><br />
            <select name="selCondition" id="selCondition" class="oByte" style="width: 140px;">
              <option></option>
            </select>
          </td>
          <td width="25%" valign="top">
            <div id="divTxt" style="display:none;">
<%
			Response.Write(getLabel(getIDS("IDS_Value"),"txtValue") & "<br />" & getTextField("txtValue","oText","",20,20,"") & vbCrLf)

			Select Case CStr(Application("av_Navigation"))
				Case "1"
					Response.Write("<div id=""divImp"" style=""display:none;""><span class=""dFont"">[<a href="""" title=""" & getIDS("IDS_Import") & """ id=""lnkImp"" onMouseOver=""(window.status='" & getIDS("IDS_Import") & "');return true;"" onMouseOut=""(window.status='');return true;"">" & getIDS("IDS_Import") & "</a>]</span></div>" & vbCrLf)
				Case Else
					Response.Write("<div id=""divImp"" style=""display:none;""><a href="""" id=""lnkImp"" onMouseOver=""(window.status='" & getIDS("IDS_Import") & "');return true;"" onMouseOut=""(window.status='');return true;""><img src="""" name=""imgImp"" alt=""" & getIDS("IDS_Import") & """ border=0 height=16 width=16 valign=absmiddle /></a></div>" & vbCrLf)
			End Select
%>
            </div>
            <div id="divSel" style="display:none;">
              <% =getLabel(getIDS("IDS_Value"),"selOptions") %><br />
              <select name="selOptions" id="selOptions" class="oNum" multiple size=5 style="width: 140px;">
                <option></option>
              </select>
            </div>
          </td>
          <td width="25%" valign="top"><% =getSubmit("btnAdd",getIDS("IDS_Add"),40,"","onClick=""addFilter();return false;""") %></td>
        </tr>
        <tr>
          <td colspan="4"><div id="filterDiv" class="dFont"></div></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr><td colspan=2><hr /></td></tr>
  <tr>
	<td><% =getLabel(getIDS("IDS_OrderBy"),"selOrder") %></td>
	<td colspan=3>
	  <select name="selOrder" id="selOrder" class="oByte">
<%
			For i = 0 to UBound(arrFields,2)
				Response.Write("<option value=""" & arrFields(0,i) & """ " & getDefault(0,bytOrder,arrFields(0,i)) & ">" & getIDS(arrFields(1,i)) & "</option>" & vbCrLf)
			Next
%>
	  </select>
	  <select name="selOrderDir" id="selOrderDir" class="oText">
		   <option value="ASC"<% =getDefault(0,"ASC",strOrderDir) %>>ASC</option>
		   <option value="DESC"<% =getDefault(0,"DESC",strOrderDir) %>>DESC</option>
	  </select>
	</td>
  </tr>
<%
		End If
%>
</table>

</div>


<div id="footerDiv" class="dvFooter">
<%
	If bytType = 90 Then
		Response.Write(getIcon("Javascript:confirmAction('gen');","V","view.gif",getIDS("IDS_ViewReport")))
	Else
		Response.Write(getIcon("Javascript:selectAllOptions(document.frmReport.selFields);confirmAction('gen');","V","view.gif",getIDS("IDS_ViewReport")))
	End If
	Response.Write(getIconNew("edit_report.asp"))
	Response.Write(getIconCancel("default.asp"))
%>
</div>
</form>

<%
		If bytType <> 90 Then
%>
<script language="JavaScript" type="text/javascript">

function Tokenizer(sVal,sDlm) {
	var iBgn = 0;
	var iEnd = 0;
	var iCount = 0;
	while(sVal.length>iBgn) {
		iEnd = sVal.indexOf(sDlm,iBgn);
		if (iEnd == -1) { iEnd = sVal.length }
		this[iCount] = sVal.substring(iBgn,iEnd);
		iCount += 1;
		iBgn = iEnd + sDlm.length;
		this.length = iCount;
	}
}

function initColumns() {
	sFields = new Tokenizer("<% =strFields %>",",");
	for (var i=0; i<=sFields.length; i++) {
		for (var j=0; j<document.forms[0].selAvailable.options.length; j++) {
			if (sFields[i] == document.forms[0].selAvailable.options[j].value) {
				document.forms[0].selAvailable.options[j].selected = true;
				moveSelectedOptions(document.forms[0].selAvailable,document.forms[0].selFields);
				break;
			}
		}
	}
}

function setConditions(iIndex,iVal) {
<%
			strOptStr1 = ""
			strOptVal1 = ""
			For i = 0 to UBound(arrFields,2)
				Select Case arrFields(1,i)
					Case "IDS_Owner", "IDS_SalesRep", "IDS_CreatedBy", "IDS_ModifiedBy"
						strOptStr1 = strOptStr1 & "1,"
						strOptVal1 = strOptVal1 & """0"","
					Case "IDS_Contact"
						strOptStr1 = strOptStr1 & "1,"
						strOptVal1 = strOptVal1 & """1"","
					Case "IDS_Account"
						strOptStr1 = strOptStr1 & "1,"
						strOptVal1 = strOptVal1 & """2"","
					Case "IDS_Sale"
						strOptStr1 = strOptStr1 & "1,"
						strOptVal1 = strOptVal1 & """3"","
					Case "IDS_Project"
						strOptStr1 = strOptStr1 & "1,"
						strOptVal1 = strOptVal1 & """4"","
					Case "IDS_Event"
						strOptStr1 = strOptStr1 & "1,"
						strOptVal1 = strOptVal1 & """"","
					Case "IDS_Ticket"
						strOptStr1 = strOptStr1 & "1,"
						strOptVal1 = strOptVal1 & """5"","
					Case "IDS_Bug"
						strOptStr1 = strOptStr1 & "1,"
						strOptVal1 = strOptVal1 & """6"","
					Case "IDS_Invoice"
						strOptStr1 = strOptStr1 & "1,"
						strOptVal1 = strOptVal1 & """7"","
					Case Else
						strOptStr1 = strOptStr1 & arrFields(3,i) & ","
						strOptVal1 = strOptVal1 & """"","
				End Select
			Next
%>
	var aConds = new Array(<% =Left(strOptStr1,Len(strOptStr1)-1) %>);
	var aImport = new Array(<% =Left(strOptVal1,Len(strOptVal1)-1) %>);
	var aStrings = new Array("");
	var aValues = new Array("");

	var iCond = aConds[iIndex];

	var oTxtDiv = getObject("divTxt");
	var oImpDiv = getObject("divImp");
	var oImpLnk = getObject("imgImp");
	var oImpImg = getObject("lnkImp");
	var oSelDiv = getObject("divSel");

	oTxtDiv.style.display = "none";
	oImpDiv.style.display = "none";
	oSelDiv.style.display = "none";

	document.forms[0].selCondition.selectedIndex = -1;
	document.forms[0].selOptions.selectedIndex = -1;
	document.forms[0].selCondition.options.length = 0;
	document.forms[0].selOptions.options.length = 0;

	switch (iCond) {
		case 1: case 6:
			aStrings = new Array("<% =getIDS("IDS_ReportCondition1") %>","<% =getIDS("IDS_ReportCondition2") %>","<% =getIDS("IDS_ReportCondition3") %>","<% =getIDS("IDS_ReportCondition4") %>","<% =getIDS("IDS_ReportCondition5") %>");
			aValues = new Array("1","2","3","4","5");
			oTxtDiv.style.display = "inline";
			if (aImport[iIndex]!="") {
				oImpDiv.style.display = "inline";
				if (lnkImp != null) { lnkImp.href = "Javascript:openWindow('<% =Application("av_CRMDir") %>common/pop_search.asp?m=" + aImport[iIndex] + "&rVal=txtValue','sw_Search','500','400');"; }
				if (document.images.imgImp != null) { document.images.imgImp.src  = "<% =Application("av_CRMDir") %>images/import2.gif"; }
			}
			break;
		case 2:
			aStrings = new Array("<% =getIDS("IDS_ReportCondition1") %>","<% =getIDS("IDS_ReportCondition2") %>","<% =getIDS("IDS_ReportCondition6") %>","<% =getIDS("IDS_ReportCondition7") %>");
			aValues = new Array("1","2","6","7");
			oTxtDiv.style.display = "inline";
			break;
		case 3:
			aStrings = new Array("<% =getIDS("IDS_ReportCondition1") %>","<% =getIDS("IDS_ReportCondition2") %>","<% =getIDS("IDS_ReportCondition8") %>","<% =getIDS("IDS_ReportCondition9") %>");
			aValues = new Array("1","2","8","9");
			oTxtDiv.style.display = "inline";
			oImpDiv.style.display = "inline";
			if (lnkImp != null) { lnkImp.href = "Javascript:showCalendar('txtValue');"; }
			if (document.images.imgImp != null) { document.images.imgImp.src  = "<% =Application("av_CRMDir") %>images/cal2.gif"; }
			break;
		case 4:
			aStrings = new Array("<% =getIDS("IDS_ReportCondition10") %>","<% =getIDS("IDS_ReportCondition11") %>");
			aValues = new Array("10","11");
			break;
		case 5:
			switch (parseInt(iVal)) {
<%
			For i = 0 to UBound(arrFields,2)
				If arrFields(3,i) = 5 Then

					strOptStr1 = ""
					strOptVal1 = ""

					Response.Write(vbTab & vbTab & vbTab & vbTab & "case " & arrFields(0,i) & ":" & vbCrLf)

					arrTemp = getArray(arrFields(1,i),getOptionSQL(arrFields(1,i)))

					If isArray(arrTemp) Then
						For j = 0 to UBound(arrTemp,2)
							strOptStr1 = strOptStr1 & """" & arrTemp(1,j) & ""","
							strOptVal1 = strOptVal1 & arrTemp(0,j) & ","
						Next
						Response.Write(vbTab & vbTab & vbTab & vbTab & vbTab & "var aStrings = new Array(" & Left(strOptStr1,Len(strOptStr1)-1) & ");" & vbCrLf)
						Response.Write(vbTab & vbTab & vbTab & vbTab & vbTab & "var aValues = new Array(" & Left(strOptVal1,Len(strOptVal1)-1) & ");" & vbCrLf)

						strOptStr2 = strOptStr2 & strOptStr1
						strOptVal2 = strOptVal2 & strOptVal1
					End If

					Response.Write(vbTab & vbTab & vbTab & vbTab & vbTab & "break;" & vbCrLf)
				End If
			Next
			If strOptStr2 = "" Then
				strOptStr2 = """"","
				strOptVal2 = """"","
			End If

%>
			}
			for (i=0; i<aValues.length; i++) {
				document.forms[0].selOptions.options[i] = new Option(aStrings[i],aValues[i],false,false);
			}

			aStrings = new Array("<% =getIDS("IDS_ReportCondition12") %>");
			aValues = new Array("12");
			oSelDiv.style.display = "inline";
			break;
	}

	for (i=0; i<aValues.length; i++) {
		document.forms[0].selCondition.options[i] = new Option(aStrings[i],aValues[i],false,false);
	}
}

function getOptionValue(iVal) {
	var aStrings = new Array(<% =Left(strOptStr2,Len(strOptStr2)-1) %>);
	var aValues = new Array(<% =Left(strOptVal2,Len(strOptVal2)-1) %>);

	for (var i=0; i<aValues.length; i++) {
		if (iVal == aValues[i]) { return aStrings[i]; break; }
	}
}


function addFilter() {
	var oParams = getObject("hdnParams");
	var oField = getObject("selField");
	var oCond = getObject("selCondition");

	var sValue = '';
	var sTemp = '';

	if (oCond.value < 10) {
		var oValue = getObject("txtValue");
		sValue = oValue.value;
		oValue.value = "";
	} else if (oCond.value == 12) {
		var oValue = getObject("selOptions");
		for (var i=0; i<oValue.options.length; i++) {
			if (oValue.options[i].selected == true) { sValue += ',' + oValue.options[i].value; }
		}
		sValue = sValue.substring(1);
		oValue.options.length = 0;
	} else {
		sValue = "1";
	}

	if (oParams.value.length != 0) { sTemp = oParams.value + "|"; }
	sTemp = sTemp + oField.value + "-" + oCond.value + "-" + sValue;

	oParams.value = sTemp;

	oField.selectedIndex = -1;
	oCond.options.length = 0;
	divTxt.style.display = "none";
	divSel.style.display = "none";
	showFilters();
	doChange();
}

function remFilter(iVal) {
	var sTemp = '';
	var oParams = getObject("hdnParams");

	aFilters = new Tokenizer(oParams.value,"|");

	for (i=0; i<aFilters.length; i++) {
		if (i != iVal) {
			if (!((i==0)||((i==1)&&(iVal==0)))) { sTemp = sTemp + '|'; }
			sTemp = sTemp + aFilters[i];
		}
	}
	oParams.value = sTemp;
	showFilters();
	doChange();
}

function showFilters() {

	var oDiv = getObject("filterDiv");
	oDiv.innerHTML = "";

	aFilters = new Tokenizer(document.forms[0].hdnParams.value,"|");

<%
			strTemp = ""
			For i = 0 to UBound(arrFields,2)
				strTemp = strTemp & """" & Application(arrFields(1,i)) & ""","
			Next
			Response.Write(vbTab & "var aFields = new Array(" & Left(strTemp,Len(strTemp)-1) & ");" & vbCrLf)

			strTemp = ""
			For i = 0 to UBound(arrFields,2)
				strTemp = strTemp & arrFields(0,i) & ","
			Next
			Response.Write(vbTab & "var aValues = new Array(" & Left(strTemp,Len(strTemp)-1) & ");" & vbCrLf)

			strTemp = ""
			For i = 1 to 12
				strTemp = strTemp & """" & getIDS("IDS_ReportCondition" & i) & ""","
			Next
			Response.Write(vbTab & "var aConditions = new Array(" & Left(strTemp,Len(strTemp)-1) & ");" & vbCrLf)
%>
	for (var i=0; i<aFilters.length; i++) {
		var iPos1 = aFilters[i].indexOf("-");
		var iPos2 = aFilters[i].indexOf("-",iPos1+1);

		var iField = aFilters[i].substring(0,iPos1);
		var iCond = aFilters[i].substring(iPos1+1,iPos2);
		var sValue = aFilters[i].substring(iPos2+1);
		for (var j=0; j<aValues.length; j++) {
			if (iField == aValues[j]) {
				iField = j;
				break;
			}
		}

		oDiv.innerHTML += '\n<br />[<a href="Javascript:remFilter(' + i + ');"><% =getIDS("IDS_Delete") %></a>] <b>' + aFields[iField] + '</b> ' + aConditions[iCond-1];
		if (iCond < 10) {
			oDiv.innerHTML += ' <b>' + sValue + '</b>';
		} else if (iCond == 12) {
			aOptions = new Tokenizer(sValue,",");
			for (var j=0; j<aOptions.length; j++) {
				if (j==0) { oDiv.innerHTML += ': ' } else { oDiv.innerHTML += ', '; }
				oDiv.innerHTML += '<b>' + getOptionValue(aOptions[j]) + '</b>';
			}
		}
	}
}

initColumns();
showFilters();

</script>

<%
		End If

	End If

	Call DisplayFooter(1)
%>