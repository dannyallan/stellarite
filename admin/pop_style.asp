<!--#include file="..\_inc\functions_edit.asp" -->
<!--#include file="..\_inc\sql\sql_admin.asp" -->
<%
	Call pageFunctions(90,5)

	Dim strMajorColor	'as String
	Dim strMinorColor	'as String
	Dim strLinkStyle	'as String
	Dim strVisitedStyle	'as String
	Dim strHoverStyle	'as String
	Dim strPrintFont	'as String
	Dim strPrintSize	'as String
	Dim strLabelFont	'as String
	Dim strLabelSize	'as String
	Dim strLabelWeight	'as String
	Dim strLabelColor	'as String
	Dim strTitleFont	'as String
	Dim strTitleSize	'as String
	Dim strTitleWeight	'as String
	Dim strTitleColor	'as String
	Dim strHeaderFont	'as String
	Dim strHeaderSize	'as String
	Dim strHeaderWeight	'as String
	Dim strHeaderColor	'as String
	Dim strDefaultFont	'as String
	Dim strDefaultSize	'as String
	Dim strDefaultWeight	'as String
	Dim strDefaultColor	'as String
	Dim strDefaultAltBG	'as String
	Dim strDefaultHiBG	'as String
	Dim strFormEnabled	'as String
	Dim strFormDisabled	'as String
	Dim strFormMandatory	'as String

	strTitle = Application("IDS_CRMStyleSheet")

	Function chooseStyle(fName,fVal)
		chooseStyle = "<select name=""" & fName & """ id=""" & fName & """ class=""oText"" onChange=""doChange();"" style=""width:135"">" & vbCrLf & _
				vbTab & "<option value=""none""" & getDefault(0,fVal,"none") & ">none</option>" & vbCrLf & _
				vbTab & "<option value=""blink""" & getDefault(0,fVal,"blink") & ">blink</option>" & vbCrLf & _
				vbTab & "<option value=""line-through""" & getDefault(0,fVal,"line-through") & ">line-through</option>" & vbCrLf & _
				vbTab & "<option value=""overline""" & getDefault(0,fVal,"overline") & ">overline</option>" & vbCrLf & _
				vbTab & "<option value=""underline""" & getDefault(0,fVal,"underline") & ">underline</option>" & vbCrLf & _
				"</select>" & vbCrLf
	End Function

	Function chooseWeight(fName,fVal)
		chooseWeight ="<select name=""" & fName & """ id=""" & fName & """ class=""oText"" onChange=""doChange();"" style=""width:135"">" & vbCrLf & _
				vbTab & "<option value=""normal""" & getDefault(0,fVal,"normal") & ">normal</option>" & vbCrLf & _
				vbTab & "<option value=""bold""" & getDefault(0,fVal,"bold") & ">bold</option>" & vbCrLf & _
				vbTab & "<option value=""bolder""" & getDefault(0,fVal,"bolder") & ">bolder</option>" & vbCrLf & _
				vbTab & "<option value=""lighter""" & getDefault(0,fVal,"lighter") & ">lighter</option>" & vbCrLf & _
				"</select>" & vbCrLf
	End Function

	Function chooseColor(fIDS,fName,fVal)
		chooseColor = getTextField(fName,"mRGB",fVal,17,7,"") & _
				"<a href=""Javascript:Dialog('../common/editor/popups/select_color.html',function(color) {if (color) {document.forms[0]." & fName & ".value = '#' + color;}},document.forms[0]." & fName & ".value.substring(1));"">" & _
				"<img src=""../images/color.gif"" alt=""" & getImport(fIDS) & """ border=0 height=16 width=16></a>"
	End Function

	If strDoAction = "edit" Then

		Application.Lock

		Call setConfigValue("av_MajorColor",valString(Request.Form("txtMajorColor"),7,1,3))
		Call setConfigValue("av_MinorColor",valString(Request.Form("txtMinorColor"),7,1,3))
		Call setConfigValue("av_LinkStyle",valString(Request.Form("selLinkStyle"),15,1,0))
		Call setConfigValue("av_VisitedStyle",valString(Request.Form("selVisitedStyle"),15,1,0))
		Call setConfigValue("av_HoverStyle",valString(Request.Form("selHoverStyle"),15,1,0))
		Call setConfigValue("av_PrintFont",valString(Request.Form("txtPrintFont"),150,1,0))
		Call setConfigValue("av_PrintSize",valString(Request.Form("txtPrintSize"),5,1,0))
		Call setConfigValue("av_LabelFont",valString(Request.Form("txtLabelFont"),150,1,0))
		Call setConfigValue("av_LabelSize",valString(Request.Form("txtLabelSize"),5,1,0))
		Call setConfigValue("av_LabelWeight",valString(Request.Form("selLabelWeight"),7,1,0))
		Call setConfigValue("av_LabelColor",valString(Request.Form("txtLabelColor"),7,1,3))
		Call setConfigValue("av_TitleFont",valString(Request.Form("txtTitleFont"),150,1,0))
		Call setConfigValue("av_TitleSize",valString(Request.Form("txtTitleSize"),5,1,0))
		Call setConfigValue("av_TitleWeight",valString(Request.Form("selTitleWeight"),7,1,0))
		Call setConfigValue("av_TitleColor",valString(Request.Form("txtTitleColor"),7,1,3))
		Call setConfigValue("av_HeaderFont",valString(Request.Form("txtHeaderFont"),150,1,0))
		Call setConfigValue("av_HeaderSize",valString(Request.Form("txtHeaderSize"),5,1,0))
		Call setConfigValue("av_HeaderWeight",valString(Request.Form("selHeaderWeight"),7,1,0))
		Call setConfigValue("av_HeaderColor",valString(Request.Form("txtHeaderColor"),7,1,3))
		Call setConfigValue("av_DefaultFont",valString(Request.Form("txtDefaultFont"),150,1,0))
		Call setConfigValue("av_DefaultSize",valString(Request.Form("txtDefaultSize"),5,1,0))
		Call setConfigValue("av_DefaultWeight",valString(Request.Form("selDefaultWeight"),7,1,0))
		Call setConfigValue("av_DefaultColor",valString(Request.Form("txtDefaultColor"),7,1,3))
		Call setConfigValue("av_DefaultAltBG",valString(Request.Form("txtDefaultAltBG"),7,1,3))
		Call setConfigValue("av_DefaultHiBG",valString(Request.Form("txtDefaultHiBG"),7,1,3))
		Call setConfigValue("av_FormEnabled",valString(Request.Form("txtFormEnabled"),7,1,3))
		Call setConfigValue("av_FormDisabled",valString(Request.Form("txtFormDisabled"),7,1,3))
		Call setConfigValue("av_FormMandatory",valString(Request.Form("txtFormMandatory"),7,1,3))

		Application.Unlock

		Call closeWindow(strOpenerURL)
	Else
		strMajorColor = Application("av_MajorColor")
		strMinorColor = Application("av_MinorColor")
		strLinkStyle = Application("av_LinkStyle")
		strVisitedStyle = Application("av_VisitedStyle")
		strHoverStyle = Application("av_HoverStyle")
		strPrintFont = Application("av_PrintFont")
		strPrintSize = Application("av_PrintSize")
		strLabelFont = Application("av_LabelFont")
		strLabelSize = Application("av_LabelSize")
		strLabelWeight = Application("av_LabelWeight")
		strLabelColor = Application("av_LabelColor")
		strTitleFont = Application("av_TitleFont")
		strTitleSize = Application("av_TitleSize")
		strTitleWeight = Application("av_TitleWeight")
		strTitleColor = Application("av_TitleColor")
		strHeaderFont = Application("av_HeaderFont")
		strHeaderSize = Application("av_HeaderSize")
		strHeaderWeight = Application("av_HeaderWeight")
		strHeaderColor = Application("av_HeaderColor")
		strDefaultFont = Application("av_DefaultFont")
		strDefaultSize = Application("av_DefaultSize")
		strDefaultWeight = Application("av_DefaultWeight")
		strDefaultColor = Application("av_DefaultColor")
		strDefaultAltBG = Application("av_DefaultAltBG")
		strDefaultHiBG = Application("av_DefaultHiBG")
		strFormEnabled = Application("av_FormEnabled")
		strFormDisabled = Application("av_FormDisabled")
		strFormMandatory = Application("av_FormMandatory")
	End If

	strIncHead = "<script language=""Javascript"" src=""../common/editor/dialog.js""></script>" & vbCrLf & _
					"<script language=""Javascript"" src=""../common/editor/popupwin.js""></script>"

	Call DisplayHeader(3)
	Call ShowEditHeader(strTitle,"","","","")
%>

<div id="contentDiv" class="dvBorder" style="height:330px;"><br>

<table border=0 width="100%">
<form name="frmAdmin" method="post" action="pop_style.asp">
<% =getHidden("hdnAction","") %>
<% =getHidden("hdnChange","") %>
<% =getHidden("hdnWinOpen","") %>
    <tr>
      <td><% =getLabel(Application("IDS_CSS_MajorColor"),"txtMajorColor") %></td>
      <td><% =chooseColor("IDS_CSS_MajorColor","txtMajorColor",strMajorColor) %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_CSS_MinorColor"),"txtMinorColor") %></td>
      <td><% =chooseColor("IDS_CSS_MinorColor","txtMinorColor",strMinorColor) %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_CSS_LinkStyle"),"selLinkStyle") %></td>
      <td><% =chooseStyle("selLinkStyle",strLinkStyle) %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_CSS_VisitedLinkStyle"),"selVisitedStyle") %></td>
      <td><% =chooseStyle("selVisitedStyle",strVisitedStyle) %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_CSS_HoverLinkStyle"),"selHoverStyle") %></td>
      <td><% =chooseStyle("selHoverStyle",strHoverStyle) %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_CSS_PrintFont"),"txtPrintFont") %></td>
      <td><% =getTextField("txtPrintFont","mText",strPrintFont,20,150,"") %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_CSS_PrintSize"),"txtPrintSize") %></td>
      <td><% =getTextField("txtPrintSize","mText",strPrintSize,5,5,"") %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_CSS_LabelFont"),"txtLabelFont") %></td>
      <td><% =getTextField("txtLabelFont","mText",strLabelFont,20,150,"") %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_CSS_LabelSize"),"txtLabelSize") %></td>
      <td><% =getTextField("txtLabelSize","mText",strLabelSize,5,5,"") %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_CSS_LabelWeight"),"selLabelWeight") %></td>
      <td><% =chooseWeight("selLabelWeight",strLabelWeight) %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_CSS_LabelColor"),"txtLabelColor") %></td>
      <td><% =chooseColor("IDS_CSS_LabelColor","txtLabelColor",strLabelColor) %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_CSS_HeaderFont"),"txtHeaderFont") %></td>
      <td><% =getTextField("txtHeaderFont","mText",strHeaderFont,20,150,"") %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_CSS_HeaderSize"),"txtHeaderSize") %></td>
      <td><% =getTextField("txtHeaderSize","mText",strHeaderSize,5,5,"") %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_CSS_HeaderWeight"),"selHeaderWeight") %></td>
      <td><% =chooseWeight("selHeaderWeight",strHeaderWeight) %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_CSS_HeaderColor"),"txtHeaderColor") %></td>
      <td><% =chooseColor("IDS_CSS_HeaderColor","txtHeaderColor",strHeaderColor) %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_CSS_TitleFont"),"txtTitleFont") %></td>
      <td><% =getTextField("txtTitleFont","mText",strTitleFont,20,150,"") %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_CSS_TitleSize"),"txtTitleSize") %></td>
      <td><% =getTextField("txtTitleSize","mText",strTitleSize,5,5,"") %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_CSS_TitleWeight"),"selTitleWeight") %></td>
      <td><% =chooseWeight("selTitleWeight",strTitleWeight) %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_CSS_TitleColor"),"txtTitleColor") %></td>
      <td><% =chooseColor("IDS_CSS_TitleColor","txtTitleColor",strTitleColor) %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_CSS_DefaultFont"),"txtDefaultFont") %></td>
      <td><% =getTextField("txtDefaultFont","mText",strDefaultFont,20,150,"") %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_CSS_DefaultSize"),"txtDefaultSize") %></td>
      <td><% =getTextField("txtDefaultSize","mText",strDefaultSize,5,5,"") %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_CSS_DefaultWeight"),"selDefaultWeight") %></td>
      <td><% =chooseWeight("selDefaultWeight",strDefaultWeight) %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_CSS_DefaultColor"),"txtDefaultColor") %></td>
      <td><% =chooseColor("IDS_CSS_DefaultColor","txtDefaultColor",strDefaultColor) %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_CSS_AltBGColor"),"txtDefaultAltBG") %></td>
      <td><% =chooseColor("IDS_CSS_AltBGColor","txtDefaultAltBG",strDefaultAltBG) %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_CSS_HighlightBGColor"),"txtDefaultHiBG") %></td>
      <td><% =chooseColor("IDS_CSS_HighlightBGColor","txtDefaultHiBG",strDefaultHiBG) %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_CSS_EnabledFormColor"),"txtFormEnabled") %></td>
      <td><% =chooseColor("IDS_CSS_EnabledFormColor","txtFormEnabled",strFormEnabled) %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_CSS_DisabledFormColor"),"txtFormDisabled") %></td>
      <td><% =chooseColor("IDS_CSS_DisabledFormColor","txtFormDisabled",strFormDisabled) %></td>
    </tr>
    <tr>
      <td><% =getLabel(Application("IDS_CSS_MandatoryFormColor"),"txtFormMandatory") %></td>
     <td><% =chooseColor("IDS_CSS_MandatoryFormColor","txtFormMandatory",strFormMandatory) %></td>
    </tr>
</form>
</table>

</div>

<div id="footerDiv" class="dvFooter">
<%
	Response.Write(getIconSave("edit"))
	Response.Write(getIconCancel())
%>
</div>
<%
	Call DisplayFooter(3)
%>

