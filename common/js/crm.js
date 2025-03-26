// Global variables used to determine the version of the browser
var bIsCSS, bIsW3C, bIsIE4, bIsIE6CSS, bIsIEQuirks;
var bIsDetectionInitialized = false;

window.onresize = doResize;
window.onload = doPageLoad;

/*****************************************************************************
		Name:		initBrowserCheck
		Action:		initialize upon load to let all browsers establish content objects
		Params:
		Returns:
		Notes:		changes the state of the global variables
*****************************************************************************/
function initBrowserCheck()
{
	if (!bIsDetectionInitialized)
	{
	        bIsCSS = (document.body && document.body.style) ? true : false;
	        bIsW3C = (bIsCSS && document.getElementById) ? true : false;
	        bIsIE4 = (bIsCSS && document.all) ? true : false;
	        bIsIE6CSS = (bIsIE4 && document.compatMode && document.compatMode.indexOf("CSS1") >= 0) ? true : false;
	        bIsIEQuirks = (bIsIE4 && (!bIsIE6CSS));
    	}
	bIsDetectionInitialized = true;
}

/*****************************************************************************
		Name:		openWindow
		Action:		Opens new window
		Params:		URL, Window Name, Width, Height
		Returns:
		Notes:
*****************************************************************************/
function openWindow(sLocation,sName,nWidth,nHeight)
{
	var oWin = window.open(sLocation,sName,"status=0,resizable,width=" + nWidth + ",height=" + nHeight);
	if (oWin.opener == null)
		oWin.opener = self;
	oWin.focus();
	return;
}

/*****************************************************************************
		Name:		closeWindow
		Action:		Closes window if needed
		Params:		URL
		Returns:
		Notes:
*****************************************************************************/
function closeWindow(sLocation)
{
	if (window.opener != null) {
		if (sLocation == "refresh")
			window.opener.location.reload();
		else if (sLocation != null)
			window.opener.location.href = sLocation;
		window.opener.focus();
		window.opener = null;
		window.close();
	} else {
		if (sLocation == "refresh")
			window.location.reload();
		else if (sLocation != null)
			window.location.href = sLocation;
	}
}

/*****************************************************************************
		Name:		sendBack
		Action:		Sends the user back with message
		Params:
		Returns:
		Notes:
*****************************************************************************/
function sendBack(sMsg)
{
	alert(sMsg.replace(/[\'\+]/g, ''));
	history.back();
}

/*****************************************************************************
		Name:		doSetTabColor
		Action:		Sets the active tab color
		Params:		TabId, Frame location, Frame name
		Returns:
		Notes:
*****************************************************************************/
function doSetTabColor(sTabId,sDest,sName)
{
	if (document.forms[0].hdnClicked.value != sTabId) {
		sLastTab = getObject(document.forms[0].hdnClicked.value);
		sLastTab.className = 'dtab';
		sTab = getObject(sTabId);
		sTab.className = 'etab';
		document.forms[0].hdnClicked.value = sTabId;
		window.frames[0].document.location.href = sDest;
		window.frames[0].document.title = sName;
	}
}

/*****************************************************************************
		Name:		doChange
		Action:		Sets the hidden change field to notify user on cancellation
		Params:
		Returns:
		Notes:
*****************************************************************************/
function doChange()
{
	var oField = getObject("hdnChange");

	if(oField != null)
		oField.value = '1';
}

/*****************************************************************************
		Name:		doFocus
		Action:		Selects the contents of first form field
		Params:		None
		Returns:
		Notes:
*****************************************************************************/
function doFocus(sId)
{
	var oElement = getObject(sId);

	if(oElement && oElement.focus)
		oElement.focus();
	if(oElement && oElement.select)
		oElement.select();
}

/*****************************************************************************
		Name:		doWarning
		Action:		Focuses the attention on problematic form field
		Params:		sId
		Returns:
		Notes:
*****************************************************************************/
function doWarning(sId)
{
	var oElement = getObject(sId);

	if(oElement) {
		oElement.style.borderColor = "#FF0000";
		oElement.style.borderStyle = "solid";
		doFocus(sId);
	}
}

/*****************************************************************************
		Name:		dateComponents
		Action:
		Params:		Date Input, Format type
		Returns:	Array with date components
		Notes:
*****************************************************************************/
function dateComponents(sDate)
{
	var results = new Array();
	var datePat = /^(\d{1,4})(\/|-)(\d{1,2})\2(\d{1,4})$/;
	var matchArray = sDate.match(datePat);

	if (matchArray == null) return null;

	if (sDateFormat == "%Y/%m/%d"){
		results[0] = matchArray[3];
		results[1] = matchArray[4];
		results[2] = matchArray[1];
	}
	else if (sDateFormat == "%d/%m/%Y") {
		results[0] = matchArray[3];
		results[1] = matchArray[1];
		results[2] = matchArray[4];
	}
	else {
		results[0] = matchArray[1];
		results[1] = matchArray[3];
		results[2] = matchArray[4];
	}

	return results;
}

/*****************************************************************************
		Name:		confirmAction
		Action:		Client side check of input fields and submits form
		Params:		Action, Dynamic Parameter
		Returns:
		Notes:
*****************************************************************************/
function confirmAction(sAct)
{
	if (sAct == "canc") {
		var oElement = getObject("hdnChange");
		if (oElement != null && oElement.value == "1")
			if (!confirm("You have changed some values on this form.  Exit?"))
				return;

	}
	else if (sAct == "del") {
		if (confirm("Are you sure you wish to delete this record?"))
				submitAction(sAct);
	}
	else {
		var msg = "";
		var strng = "";
		var clss = "";
		var frmnm = "";
		var pass = "";

		for (i=document.forms[0].elements.length-1; i > 0; i--)	{

			strng = document.forms[0].elements[i].value;
			clss = document.forms[0].elements[i].className;
			frmnm = getFieldName(document.forms[0].elements[i].name);

			if (clss.substring(0,1) == "m" && strng == "") {
				msg += "You must fill in the mandatory field "+frmnm+"\n\n";
				doWarning(i);
			}
			else {
				if (clss.substring(1) == "Email" && strng != "") {
					var emailFilter=/^[\w\d\.\%-]+@[\w\d\.\%-]+\.\w{2,4}$/;
					if (!(emailFilter.test(strng))) {

						var illegalChars= /[^\w\d\.\%\-@]/;
						if (strng.match(illegalChars)) {
					  		msg += "The field "+frmnm+" can only contain alphanumeric\ncharacters and the following:  @.%-\n\n";
						} else {
							msg += "Please enter a valid email address for "+frmnm+"\n\n";
						}
						doWarning(i);
					}
				}
				else if (clss.substring(1) == "Link" && strng != "") {
					var linkFilter=/^(https?|ftp|[A-Za-z]:\\|\\\\)[^<>()''""]+$/;
					if (!(linkFilter.test(strng))) {

						var illegalChars= /[<>()''""]/;
						if (strng.match(illegalChars)) {
							msg += "The link "+frmnm+" may not contain the\nfollowing characters: '<>()\"\n\n";
						} else {
							msg += "The field "+frmnm+" does not appear to contain a valid link\n\n";
						}
						doWarning(i);
					}
				}
				else if (clss.substring(1) == "Phone" && strng != "") {
					var stripped = strng.replace(/[\(\)\.\-\ ]/g, '');
					var allowedValues=/[^\d]+/;

					if (stripped.length > 15) {
						msg += "The phone number "+frmnm+" can only contain a maximum of 15 digits\n\n"
						doWarning(i);
					}
					else if (allowedValues.test(stripped)) {
						msg += "The field "+frmnm+" may only contain\ndigits and the following characters:   ().-\n\n";
						doWarning(i);
					}
					else {
						window.document.forms[0].elements[i].value = stripped;
					}
				}
				else if (clss.substring(1) == "Num" && strng != "") {
					var allowedValues=/[\d\.]+/;

					if (!(allowedValues.test(strng))) {
						msg += "The field "+frmnm+" does not contain a valid number\n\n";
						doWarning(i);
					}
				}
				else if (clss.substring(1) == "Bool" && strng != "") {
					var allowedValues=/[^01]/;

					if (allowedValues.test(strng)) {
						msg += "The field "+frmnm+" can only contain a 0 or 1\n\n";
						doWarning(i);
					}
				}
				else if (clss.substring(1) == "Byte" && strng != "") {
					var allowedValues=/[\d]+/;

					if (!(allowedValues.test(strng))) {
						msg += "The field "+frmnm+" must contain a valid number\n\n";
						doWarning(i);
					}
					else if ((strng < 0) || (strng > 255)) {
						msg += "The field "+frmnm+" must contain an integer between 0 and 255\n\n";
						doWarning(i);
					}
				}
				else if (clss.substring(1) == "Int" && strng != "") {
					var allowedValues=/[\d]+/;

					if (!(allowedValues.test(strng))) {
						msg += "The field "+frmnm+" must contain a valid number\n\n";
						doWarning(i);
					}
					else if ((strng < -32767) || (strng > 32767)) {
						msg += "The field "+frmnm+" must contain an integer between -32,767 and 32,767\n\n";
						doWarning(i);
					}
				}
				else if (clss.substring(1) == "Long" && strng != "") {
					var allowedValues=/[\d]+/;

					if (!(allowedValues.test(strng))) {
						msg += "The field "+frmnm+" must contain a valid number\n\n";
						doWarning(i);
					}
					else if ((strng < -2147483647) || (strng > 2147483647)) {
						msg += "The field "+frmnm+" must contain an integer between -2,147,483,647 and 2,147,483,647\n\n";
						doWarning(i);
					}
				}
				else if (clss.substring(1) == "Currency" && strng != "") {
					var allowedValues=/[\d\.]+/;

					if (!(allowedValues.test(strng))) {
						msg += "The field "+frmnm+" must contain a valid number\n\n";
						doWarning(i);
					}
					else if ((strng < -922337203685477.5807) || (strng > 922337203685477.5807)) {
						msg += "The field "+frmnm+" is outside the allowed value\n\n";
						doWarning(i);
					}
				}
				else if (clss.substring(1) == "RGB" && strng != "") {
					var allowedValues=/^#[\dA-F]{6}$/;

					if (!(allowedValues.test(strng))) {
						msg += "The field "+frmnm+" must be a valid 6 character RGB value\npreceded by the pound symbol\n\n";
						doWarning(i);
					}
					else {
						window.document.forms[0].elements[i].value = strng;
					}
				}
				else if (clss.substring(1) == "Date" && strng != "") {
					var dateBits = dateComponents(strng);
					if (dateBits == null) {
						msg += "The date syntax of "+frmnm+" must be "+sDateFormat;
						doWarning(i);
					}
					else {
						month = dateBits[0];
						day = dateBits[1];
						year = dateBits[2];

						if ((month < 1 || month > 12) || (day < 1 || day > 31)) {
							msg += "The date value "+frmnm+" contains illegal values.\nThe date syntax should be "+sDateFormat+"\n\n";
							doWarning(i);
						}
						if ((month==4 || month==6 || month==9 || month==11) && day==31) {
							msg += "The date value "+frmnm+" contains illegal values.\nThe date syntax should be "+sDateFormat+"\n\n";
							doWarning(i);
						}
						if (month == 2) {
							var isleap = (year % 4 == 0 && (year % 100 != 0 || year % 400 == 0));
							if (day>29 || (day==29 && !isleap)) {
								msg += "The date value "+frmnm+" contains illegal values.\nThe date syntax should be "+sDateFormat+"\n\n";
								doWarning(i);
							}
						}
					}
				}
				else if (clss.substring(1) == "Memo") {
					if (strng.length > 255) {
						msg += "The memo field "+frmnm+" cannot exceed 255 characters.\n You current count is "+strng.length+" characters\n\n";
					}
				}
				else if (window.document.forms[0].elements[i].type == "password") {
					pass = "1";
				}
				else {
					window.document.forms[0].elements[i].value = strng;
				}
			}
		}
		if (msg != "") {
			if (pass != "") {
				for (i=document.forms[0].elements.length-1; i >= 0; i--)	{
					if (document.forms[0].elements[i].type == "password") {
						document.forms[0].elements[i].value = "";
					}
				}
			}
			alert(msg);
			return;
		}
		submitAction(sAct);
	}
}

/*****************************************************************************
		Name:		submitAction
		Action:		Sets the hidden form field with user action
		Params:		Action value
		Returns:
		Notes:
*****************************************************************************/
function submitAction(sAction)
{
	if (iMode == 0) {
		var oField = getObject("hdnAction");
		if(oField != null)
			oField.value = sAction;
		window.document.forms[0].submit();
	} else {
		alert("The CRM is currently in read-only mode.\nYou are not permitted to change form values.");
	}
	return;
}

/*****************************************************************************
		Name:		getFieldName
		Action:		Strips off the first three characters and enables friendly view
		Params:		Form Field Name
		Returns:	Friendly field name
		Notes:
*****************************************************************************/
function getFieldName(sField) {
	var sName = sField.substring(3);
	sName = sName.replace(/[\d]/g, '');
	sName = sName.replace(/([a-z])([A-Z])/g, '$1 $2');
	return sName;
}

/*****************************************************************************
		Name:		getObject
		Action:		Convert object name string or object reference into a valid element object reference
		Params:		obj - string with a name of an object or a reference to the object
		Returns:	reference to the object
		Notes:		Supports NN 6.2.3, IE5+, Mozilla 1.0+ (NN 7+)
*****************************************************************************/
function getObject(sId)
{
	initBrowserCheck();
	var oElement;
	if(typeof(sId) == "string") {
		if (bIsW3C) {
			oElement = document.getElementById(sId);
		}
		else if (bIsIE4) {
			oElement = document.all(sId);
		}
	} else {
		oElement = document.forms[0].elements[sId];
	}
	return oElement;
}

/*****************************************************************************
		Name:		getObjectHeight
		Action:		Retrieve the rendered height of an element
		Params:		String with a name of an object or a reference to the object
		Returns:	Height of the element
		Notes:		Supports NN 6.2.3, IE5+, Mozilla 1.0+ (NN 7+)
*****************************************************************************/
function getObjectHeight(obj)
{
	var oElem = getObject(obj);
	var nResult = 0;

	if (oElem.offsetHeight)
	{
		nResult = oElem.offsetHeight;
	}
	else if (oElem.clip && oElem.clip.height)
	{
		nResult = oElem.clip.height;
	}
	else if (oElem.style && oElem.style.pixelHeight)
	{
		nResult = oElem.style.pixelHeight;
	}
	return parseInt(nResult);
}

/*****************************************************************************
		Name:		getWindowHeight
		Action:		Return the available content height space in browser window
		Params:		None
		Returns:	Height in pixels
		Notes:		Supports NN 6.2.3, IE4+, Mozilla 1.0+ (NN 7+)
*****************************************************************************/
function getWindowHeight()
{
	initBrowserCheck();
	if (window.innerHeight)
	{
		return window.innerHeight;
	}
	else if (bIsIE6CSS)
	{
		return document.body.parentElement.clientHeight;
	}
	else if (document.body && document.body.clientHeight)
	{
		return document.body.clientHeight;
	}
	return 0;
}

/*****************************************************************************
		Name:		getWindowWidth
		Action:		Return the available content width in browser window
		Params:		None
		Returns:	Height in pixels
		Notes:		Supports NN 6.2.3, IE4+, Mozilla 1.0+ (NN 7+)
*****************************************************************************/
function getWindowWidth()
{
	initBrowserCheck();
	if (window.innerWidth)
	{
		return window.innerWidth;
	}
	else if (bIsIE6CSS)
	{
		return document.body.parentElement.clientWidth;
	}
	else if (document.body && document.body.clientWidth)
	{
		return document.body.clientWidth;
	}
	return 0;
}

/*****************************************************************************
		Name:		doPageLoad
		Action:		Handles resizing for all windows and focuses
				on first form field.
		Params:
		Returns:
		Notes:		delegates to doResize()
*****************************************************************************/
function doPageLoad()
{
	var nHeight = getWindowHeight();

	doResize();

	var nFHeight = getWindowHeight();
	if (nFHeight != nHeight) {
		doResize();
	}

	if (document.forms[0]) {
		for (i=0; i < document.forms[0].elements.length; i++) {
			if ((document.forms[0].elements[i].type != "hidden")
				&& (document.forms[0].elements[i].className != "dText")
				&& (document.forms[0].elements[i].type != "image")
				&& (document.forms[0].elements[i].type != "submit")) {
				doFocus(i);
				return;
			}
		}
	}
}

/*****************************************************************************
		Name:		doResize
		Action:		Handles resizing of content div
		Params:
		Returns:
		Notes:
*****************************************************************************/
function doResize()
{
	initBrowserCheck();

	var nHeight = getWindowHeight();
	var nHeightReserved = 0;

	var navDiv = getObject("navDiv");
	if(navDiv != null)
		nHeightReserved += getObjectHeight("navDiv");

	var headerDiv = getObject("headerDiv");
	if(headerDiv != null)
		nHeightReserved += getObjectHeight("headerDiv") + 0;

	var modDiv = getObject("modDiv");
	if(modDiv != null)
		nHeightReserved += getObjectHeight("modDiv") + 10;

	var footerDiv = getObject("footerDiv");
	if(footerDiv != null)
		nHeightReserved += getObjectHeight("footerDiv") + 40;

	var contentDiv = getObject("contentDiv");
	if(contentDiv != null) {
		if(contentDiv.className == "iBorder") {
			nHeightReserved += 10;
			contentDiv.style.width = (getWindowWidth() - 25) + "px";
		}
		nHeight = (nHeight - nHeightReserved);
		contentDiv.style.height = nHeight + "px";
	}
}