/*
The following errors were found when attempting to minify this file:
- Line 1675: Parse error. syntax error
- Line 1680: Parse error. syntax error
- Line 1681: Parse error. syntax error
- Line 1682: Parse error. 'try' without 'catch' or 'finally'
- Line 1684: Parse error. syntax error
- Line 1686: Parse error. syntax error
- Line 1688: Parse error. syntax error
- Line 1689: Parse error. missing ; before statement
- Line 1692: Parse error. syntax error
- Line 1694: Parse error. syntax error
- Line 1695: Parse error. syntax error
- Line 1696: Parse error. invalid return
- Line 1697: Parse error. syntax error
- Line 5013: Parse error. illegal character
- Line 5013: Parse error. missing ; before statement
- Line 5013: Parse error. missing ; before statement
- Line 5013: Parse error. unterminated regular expression literal
- Line 5015: Parse error. missing ; before statement
- Line 5016: Parse error. syntax error
- Line 5020: Parse error. syntax error
- Line 5022: Parse error. missing ; before statement
- Line 5025: Parse error. syntax error
- Line 5026: Parse error. syntax error
- Line 5032: Parse error. syntax error
- Line 5034: Parse error. missing ; before statement
- Line 5037: Parse error. syntax error
- Line 5038: Parse error. syntax error
- Line 5044: Parse error. syntax error
- Line 5046: Parse error. missing ; before statement
- Line 5049: Parse error. syntax error
- Line 5050: Parse error. syntax error
- Line 5053: Parse error. syntax error
- Line 5056: Parse error. missing ; before statement
- Line 5062: Parse error. syntax error
- Line 5064: Parse error. missing ; before statement
- Line 5069: Parse error. syntax error
- Line 5072: Parse error. missing ; before statement
- Line 5073: Parse error. syntax error
- Line 5092: Parse error. syntax error
- Line 5094: Parse error. missing ; before statement
- Line 5095: Parse error. missing ; before statement
- Line 5096: Parse error. syntax error
- Line 5097: Parse error. syntax error
- Line 5099: Parse error. missing ; before statement
- Line 5103: Parse error. syntax error
- Line 5109: Parse error. syntax error
- Line 5111: Parse error. missing ; before statement
- Line 5119: Parse error. invalid return
- Line 5145: Parse error. syntax error
- Line 5911: Parse error. missing ; before statement
- Line 5912: Parse error. missing ; before statement
- Line 5913: Parse error. missing ; before statement
- Line 5914: Parse error. missing ; before statement
- Line 5915: Parse error. missing ; before statement
- Line 5916: Parse error. missing ; before statement
- Line 7183: Parse error. missing ; before statement
- Line 7210: Parse error. syntax error
- Line 7237: Parse error. illegal character
- Line 7237: Parse error. illegal character
- Line 7237: Parse error. missing ) after argument list
- Line 7237: Parse error. missing ; before statement
- Line 7237: Parse error. illegal character
- Line 7237: Parse error. syntax error
- Line 7239: Parse error. syntax error
- Line 8198: Parse error. syntax error
- Line 8474: Parse error. missing ; before statement
- Line 10824: Parse error. syntax error
- Line 10834: Parse error. missing ) in parenthetical
- Line 10836: Parse error. syntax error
- Line 10839: Parse error. invalid return
- Line 10844: Parse error. invalid return
- Line 10857: Parse error. illegal character
- Line 10857: Parse error. missing ; before statement
- Line 10857: Parse error. illegal character
- Line 10860: Parse error. illegal character
- Line 10860: Parse error. missing ; before statement
- Line 10860: Parse error. illegal character
- Line 10865: Parse error. illegal character
- Line 10865: Parse error. missing ; before statement
- Line 10865: Parse error. illegal character
- Line 10868: Parse error. syntax error
- Line 10870: Parse error. syntax error
- Line 10871: Parse error. missing ; before statement
- Line 10873: Parse error. syntax error
- Line 10875: Parse error. syntax error
- Line 10876: Parse error. missing ; before statement
- Line 10913: Parse error. syntax error
- Line 10915: Parse error. syntax error
- Line 10928: Parse error. invalid return
- Line 10934: Parse error. invalid return
- Line 10948: Parse error. syntax error
- Line 10952: Parse error. invalid return
- Line 10953: Parse error. syntax error
- Line 10956: Parse error. invalid return
- Line 10962: Parse error. invalid return
- Line 10963: Parse error. syntax error
- Line 10965: Parse error. syntax error
- Line 10968: Parse error. invalid return
- Line 10973: Parse error. invalid return
- Line 10978: Parse error. syntax error
- Line 10979: Parse error. missing ; before statement
- Line 10981: Parse error. syntax error
- Line 10983: Parse error. missing ; before statement
- Line 10988: Parse error. syntax error
- Line 10991: Parse error. syntax error
- Line 11002: Parse error. invalid return
- Line 11005: Parse error. syntax error
- Line 11006: Parse error. missing ; before statement
- Line 11007: Parse error. syntax error
- Line 11010: Parse error. syntax error
- Line 11019: Parse error. invalid return
- Line 11020: Parse error. syntax error
- Line 11022: Parse error. syntax error
- Line 11025: Parse error. invalid return
- Line 11029: Parse error. invalid return
- Line 11031: Parse error. syntax error
- Line 11032: Parse error. syntax error
- Line 11033: Parse error. syntax error
- Line 11036: Parse error. syntax error
- Line 11037: Parse error. missing ; before statement
- Line 11040: Parse error. syntax error
- Line 11041: Parse error. syntax error
- Line 11043: Parse error. syntax error
- Line 11044: Parse error. syntax error
- Line 11047: Parse error. syntax error
- Line 11049: Parse error. syntax error
- Line 11051: Parse error. syntax error
- Line 11053: Parse error. syntax error
- Line 11055: Parse error. syntax error
- Line 11058: Parse error. syntax error
- Line 11064: Parse error. syntax error
- Line 11067: Parse error. missing ; before statement
- Line 11068: Parse error. missing ; before statement
- Line 11069: Parse error. syntax error
- Line 11073: Parse error. continue must be inside loop
- Line 11078: Parse error. invalid return
- Line 11079: Parse error. syntax error
- Line 11085: Parse error. syntax error
- Line 11086: Parse error. syntax error
- Line 11087: Parse error. invalid return
- Line 11087: Parse error. illegal character
*/
/// <reference name="MicrosoftAjax.js" />
/// <reference name="Util.js" />
/// <reference path="GlobalJSFunctionsDetail.js" />

//Callback global variables for Slideshow.
var callbackfnDocumentCenter = null;
var callbackfnSlideshowSave = null;
var $popUp_Slideshow = null;
var parentID_Slideshow = null;
var slideshowFolderDefault = '';
var numbersOnly = new RegExp('^[0-9]*$');
//Global variable for ThemeProperties modal
var $popup_ThemeProperties = null;
//Callback variables for Slideshow.
var SetCursorPosition = null;

// Prevents errors for Internet Explorer (modes prior to 9).
// They will occur unless the developer tools are open.
if (!window.console) {
	var noop = function () { };

	// Note: MSIE will override this once the dev. tools are open.
	window.console = {
		log: noop, clear: noop, warn: noop, error: noop, assert: noop,
		dir: noop, count: noop, profile: noop, profileEnd: noop,
		trace: noop, info: noop, memoryProfile: noop, memoryProfileEnd: noop,
		exception: noop, debug: noop, dirxml: noop, group: noop, groupEnd: noop,
		markTimeline: noop, time: noop, timeEnd: noop, groupCollapsed: noop
	};
}

// User-agent sniffing for functions that rely on the user-agent
// to determine version of Safari. Avoid using this variable and
// check for the presence of features instead (best-practice).
var isWebKit = (navigator.userAgent.toLowerCase().indexOf('webkit') > -1);

// Determines if an array contains an object.
// Uses == equality, case-sensitive. a containsExact() could do === equality.
if (!Array.prototype.contains) {
	Array.prototype.contains = function (item) {
		for (var i = 0; i < this.length; i++) {
			if (this[i] == item)
				return true;
		}
		return false;
	}
}

// Gets index of an item if it is present in the array, else returns -1.
if (!Array.prototype.indexOf) {
	Array.prototype.indexOf = function (item) {
		for (var i = 0; i < this.length; i++) {
			if (this[i] == item)
				return i;
		}
		return -1;
	}
}

// Implement ECMAScript 5 String.trim() and friends for browsers that
// don't have it (e.g. Firefox prior to 3.5, Internet Explorer 6, 7, 8).
if (!String.prototype.trim) {
	String.prototype.trim = function () {
		return this.replace(/^\s+/g, "").replace(/\s+$/g, "");
	}
	String.prototype.trimLeft = function () {
		return this.replace(/^\s+/g, "");
	}
	String.prototype.trimRight = function () {
		return this.replace(/\s+$/g, "");
	}
}

// Implement .NET-style left-padding method.
String.prototype.padLeft = function (desiredLength, paddingChar) {
	paddingChar += '';
	desiredLength = parseInt(desiredLength);

	if (paddingChar.length > 1)
		paddingChar = paddingChar.charAt(0);
	else if (paddingChar.length < 1)
		return this;

	if (this.length >= desiredLength)
		return this;
	else {
		var arr = new Array();
		var padNum = desiredLength - this.length;

		for (var i = 0; i < padNum; i++)
			arr.push(paddingChar);

		return arr.join('') + this;
	}
}

// Implement .NET-style right-padding method.
String.prototype.padRight = function (desiredLength, paddingChar) {
	paddingChar += '';
	desiredLength = parseInt(desiredLength);

	if (paddingChar.length > 1)
		paddingChar = paddingChar.charAt(0);
	else if (paddingChar.length < 1)
		return this;

	if (this.length >= desiredLength)
		return this;
	else {
		var arr = new Array();
		var padNum = desiredLength - this.length;

		for (var i = 0; i < padNum; i++)
			arr.push(paddingChar);

		return this + arr.join('');
	}
}

String.isNullOrEmpty = function (text) {
	return (text == null || text == '');
}

// Adds static method to regular expressions that escapes meta-characters.
if (!RegExp.metaEscape) {
	RegExp.metaEscape = function (text) {
		if (text == null)
			return null;
		else if (typeof (text) != "string")
			text += '';

		return text.replace(/(\\|\/|\<|\>|\:|\.|\*|\+|\?|\$|\[|\]|\(|\)|\{|\}|\||\&)/g, '\\$1');
	};
}

// Copies text data to clipboard in WebKit browsers (Chrome/Safari).
function toWebkitClipboard(text) {
	// Create element in DOM with text to copy.
	var tmp = document.createElement('div');
	tmp.textContent = text;
	document.body.appendChild(tmp);

	// Get current user selection, and remove selection current ranges.
	var curSelection = window.getSelection();
	curSelection.removeAllRanges();

	// Create new selection range with DOM element containing text to copy.
	var textRange = document.createRange();
	textRange.selectNode(tmp);
	curSelection.addRange(textRange);

	// Execute COPY DHTML command.
	document.execCommand("Copy");

	// Clean-up.
	document.body.removeChild(tmp);
}

// Copies value in stringVal to the clipboard, displaying the successMessage if hideAlert is false.
// NOTE: Replace this with something like toWebkitClipboard() at some point.
function toClipboardEx(stringVal, hideAlert, successMessage) {
	if (window.clipboardData) {		// Internet Explorer
		if (!window.clipboardData.setData("Text", stringVal) && !hideAlert)
			hideAlert = true;
	}
	else if ((window.WebKitPoint || !window.netscape) && !window.opera) {
		toWebkitClipboard(stringVal);
	}
	else if (window.netscape) {	// Mozilla Firefox and derivitives (Netscape, Seamonkey)...
		try {
			// Request full access to the XPCOM (Cross-Platform COM) API.
			netscape.security.PrivilegeManager.enablePrivilege('UniversalXPConnect');
		}
		catch (e) {
			// Unable to obtain access, user rejected or setting in about:config not set right.
			alert('Please type "about:config" in your browser and press enter. Type "signed.applets.codebase_principal_support" in Filter. Double click to change the value to "true". Then come back and click on the link again.\n\nIf you have already performed this action, make sure when you are asked whether to allow or deny the browser permission, that you are allowing it.');
			return;
		}

		// Create an instance of the clipboard class.
		var clipBoard = Components.classes['@mozilla.org/widget/clipboard;1'].createInstance(Components.interfaces.nsIClipboard);

		// Create an instance of the Transferable class (used to talk to the clipboard).
		var clipTrans = Components.classes['@mozilla.org/widget/transferable;1'].createInstance(Components.interfaces.nsITransferable);

		// Set clipboard format for text only.
		clipTrans.addDataFlavor('text/unicode');

		// Create XPCOM string, set data to copy of stringVal.
		var clipString = Components.classes["@mozilla.org/supports-string;1"].createInstance(Components.interfaces.nsISupportsString);
		var copyText = stringVal;
		clipString.data = copyText;

		// Note: This code may be bugged in some scenarios! 1 char does not always equal 2 bytes in UTF-16!
		// Set data to transfer to the clipboard (length * 2, since 1 char is usually 2 bytes in UTF-16).
		clipTrans.setTransferData("text/unicode", clipString, copyText.length * 2);

		// Transfer data to the global clipboard.
		clipBoard.setData(clipTrans, null, clipBoard.kGlobalClipboard);
	}

	if (!hideAlert)
		alert(successMessage);

	return true;
}

function toClipboard(stringVal, hideAlert) {
	return toClipboardEx(stringVal, hideAlert, "The link has been copied to your clipboard.");
}

function getClipboard() {
	if (window.clipboardData) {
		return window.clipboardData.getData('Text');
	} else if (window.netscape) {
		try {
			netscape.security.PrivilegeManager.enablePrivilege('UniversalXPConnect');
		}
		catch (e) {
			alert('Please type "about:config" in your browser and press enter. Type "signed.applets.codebase_principal_support" in Filter. Double click to change the value to "true". Then come back and click on the link again.\n\nIf you have already performed this action, make sure when you are asked whether to allow or deny the browser permission, that you are allowing it.');
			return;
		}

		// Create instances of the Clipboard and Transferable objects.
		var clipBoard = Components.classes['@mozilla.org/widget/clipboard;1'].createInstance(Components.interfaces.nsIClipboard);
		var clipTrans = Components.classes['@mozilla.org/widget/transferable;1'].createInstance(Components.interfaces.nsITransferable);

		// Get data from global clipboard, place in clipTrans object.
		clipTrans.addDataFlavor('text/unicode');
		clipBoard.getData(clipTrans, clipBoard.kGlobalClipboard);

		// Create objects to pass to getTransferData, which uses XPCOM String
		// and Integer classes instead of those normally used by JavaScript.
		var objStr = new Object(); var objNumBytes = new Object();
		try {
			clipTrans.getTransferData('text/unicode', objStr, objNumBytes);
		}
		catch (error) {
			return '';
		}

		// Query whichever interface is available, opting to use nsISupportsWString.
		if (objStr) {
			if (Components.interfaces.nsISupportsWString)
				objStr = objStr.value.QueryInterface(Components.interfaces.nsISupportsWString);
			else if (Components.interfaces.nsISupportsString)
				objStr = objStr.value.QueryInterface(Components.interfaces.nsISupportsString);
			else
				objStr = null;
		}

		// Note: This code may be bugged in some scenarios! 1 char does not always equal 2 bytes in UTF-16!
		// Get data out of the XPCOM string and into a normal JS string.
		if (objStr)
			return (objStr.data.substring(0, objNumBytes.value / 2));
	}
	return;
}

// Gets the computed style of an element (browser-independent).
var getCurrentStyle = (document.defaultView ?
	function (elem) {
		return document.defaultView.getComputedStyle(elem, '');
	} :
	function (elem) {
		return elem.currentStyle;
	}
);

// Gets children nodes of an HTML DOM element matching the tag name specified.
function getChildNodesByTag(domElement, tagName) {
	var cn = domElement.childNodes;
	var nodesPresent = 0, retVal = new Array();
	retVal.length = cn.length;

	tagName = tagName.toUpperCase();
	for (var cv = 0; cv < cn.length; cv++) {
		if (cn[cv].nodeType == 1 && cn[cv].nodeName == tagName)
			retVal[nodesPresent++] = cn[cv];
	}

	return retVal.slice(0, nodesPresent);
}

// Returns true on Safari browsers 1.3.4 and older (bad date object behavior in old Safari).
// Check does not work properly if user-agent has been modified by user.
var isSafariVersion13OrOlder = (isWebKit ?
	function () {
		var navAppVersion = navigator.appVersion;
		var phraseToFind = 'AppleWebKit/';
		var foundStartAt = navAppVersion.indexOf(phraseToFind) + phraseToFind.length;
		var foundEndAt = navAppVersion.indexOf(' ', foundStartAt + 1);
		return (parseInt(navAppVersion.substring(foundStartAt, foundEndAt)) < 320);
	} :
	function () {
		return false;
	}
);

// Returns true on Safari browsers 2.0.4 and older (textbox/button not stylable before 3.x).
// Check does not work properly if user-agent has been modified by user.
var isSafariVersion20OrOlder = (isWebKit ?
	function () {
		var navAppVersion = navigator.appVersion;
		var phraseToFind = 'AppleWebKit/';
		var foundStartAt = navAppVersion.indexOf(phraseToFind) + phraseToFind.length;
		var foundEndAt = navAppVersion.indexOf(' ', foundStartAt + 1);
		return (parseInt(navAppVersion.substring(foundStartAt, foundEndAt)) < 525);
	} :
	function () {
		return false;
	}
);

// Returns true if an event fired too soon after another event.
// Mechanism used on older Safari browsers to prevent
// double-fire problems (especially with key presses).
function safariEventRateLimitBlock() {
	if (isWebKit && isSafariVersion13OrOlder()) {
		if (safariRateLimited != 0)
			return true;
		else {
			safariRateLimited = setTimeout('safariRateLimited = 0;', 10);
			return false;
		}
	}
}

// Get coordinates relative to document (works out container issues).
// Wrote this so it works in quirks or standards mode.
function getDocumentCoordinates(element) {
	var htmlElem = document.body.parentNode;
	var bodyElem = document.body;
	var pos = { "x": 0, "y": 0 };
	pos.toString = function () { return this.x + ', ' + this.y; }
	while (element != null) {
		pos.x += element.offsetLeft;
		pos.y += element.offsetTop;
		switch (element.offsetParent) {
			case htmlElem:
			case bodyElem:
			case null:
				return pos;
		}
		element = element.offsetParent;
	}
}

// Attach an event handler to an object (browser-independent).
// First clause is W3C DOM method, second is DHTML (IE).
var addEvent = (window.addEventListener ?
	function (obj, evType, fn, useCapture) {
		try {
			obj.addEventListener(evType, fn, useCapture);
		} catch (e) { }
		return true;
	} :
	function (obj, evType, fn, useCapture) {
		try {
			return obj.attachEvent('on' + evType, fn);
		} catch (e) { }
	}
);

// Release an event handler from an object (browser-independent).
// First clause is W3C DOM method, second is DHTML (IE).
var removeEvent = (window.removeEventListener ?
	function (obj, evType, fn, useCapture) {
		try {
			obj.removeEventListener(evType, fn, useCapture);
		} catch (e) { }
		return true;
	} :
	function (obj, evType, fn, useCapture) {
		try {
			obj.detachEvent('on' + evType, fn);
		} catch (e) { }
		return true;
	}
);

// Stops a DOM event from propagating further than the current handler.
// Note: Returning false from event handlers in IE 6/7/8 does the same thing.
function stopEventPropagation(evObj) {
	if (evObj.preventDefault)
		evObj.preventDefault();

	// Calling cancelButton after stopPropagation
	// may negate the stopPropagation so do not do it. -KB
	if (evObj.stopPropagation)
		evObj.stopPropagation();
	else if (evObj.cancelBubble)
		evObj.cancelBubble();
}

// These functions assists in obscuring email addresses:
function mailTo(obj) {
	//alert('mailto:'+eval(obj.getAttribute('id')));
	obj.setAttribute('href', 'mailto:' + eval(obj.getAttribute('id')));
}

function js_mail(obj, Emails) {
	var mail_link = Emails[0];
	for (var email = 1; email < Emails.length; email++)
		if (Emails[email] != null && Emails[email] != '' && Emails[email].substr(1, 4) != 'href') mail_link = mail_link + Emails[email];

	obj.setAttribute('href', 'mailto:' + mail_link);
}

// Used to prevent right-click menu from appearing for some clients that wanted this ability.
function antiContextMenuHook() {
	// Note: Opera is not hookable.
	var showAlert = function () { alert('All images are protected by Copyright. Do not use without permission.'); }
	var mdClick = function (e) {
		if (!document.all) {
			if (e.button == 2 || e.button == 3)
				showAlert();
		}
		else if (event && event.button == 2)
			showAlert();
	}
	var cmClick = function (e) {
		if (navigator.userAgent.toLowerCase().indexOf('khtml') > -1) {
			// Safari, Konquerer
			if (e.preventDefault)
				e.preventDefault();
			showAlert();
		}
		if (e.stopPropagation)
			e.stopPropagation(); // Mozilla Firefox 2.0
		return false; // IE 6.0 and 7.0
	}
	document.onmousedown = mdClick;
	document.oncontextmenu = cmClick;
}

// Form validation functions:
function RegExValidate(expression, value, param) {
    var re = new RegExp(expression, param);
    if (value != '' && value != undefined) {
        if (value.match(re)) return true;
        else return false;
    }
    else return true;
}

// Used for validating emailaddress for special scenarios like continuous periods. - VB
function checkSpecialScenarios(s) {
    var bugchars = '!#$^&*()|}{[]?><~%:;/,=`"'; 
	var i;
	var lchar = "";
	// Search through string's characters one by one.
	// If character is not in bag.
	for (i = 0; i < s.length; i++) {
		// Check that current character isn't whitespace.
		var c = s.charAt(i);
		if (i > 0) lchar = s.charAt(i - 1)
		if (bugchars.indexOf(c) != -1 || (lchar == "." && c == ".")) return false;
	}
	return true;
}

// JavaScript version of CivicPlus.CMS.Site.Validation.IsValidEmailAddress();
function emailValidate(emailAddress) {
	var emailAddressTrimmed = TrimString(emailAddress + "");
	if (emailAddressTrimmed != "") {
		if (checkSpecialScenarios(emailAddressTrimmed) == false)
			return false;
		if (emailAddressTrimmed.replace(/[^@]/g, '').length > 1)
			return false;
		var parts = emailAddressTrimmed.splitRemoveEmpties('@');
		if (emailAddressTrimmed.substr(emailAddressTrimmed.length - 1) == '@')
			return false;
		if (parts.length == 2) {
			for (var i = 0; i < parts.length; i++) {
				if (((!RegExValidate('^[A-Z0-9]$', parts[i].substr(0, 1), 'i') || !RegExValidate('^[A-Z0-9]$', parts[i].substr(parts[i].length - 1), 'i')) && i == 1) || (!RegExValidate('^[A-Z0-9_%-\\.]+$', parts[i].substr(0, parts[i].length - 1), 'i')))
					return false;
			}
			var lastDotPos = parts[1].lastIndexOf('.');
			if (lastDotPos < 0 || parts[1].substr(lastDotPos).length < 3 || (!RegExValidate('^[A-Z]+$', parts[1].substr(lastDotPos + 1), 'i')))
				return false;
			else
				return true;
		}
		else
			return false;
	}
	else
		return false;
}

// Returns: 0 - success, 1 - illegal value, 2 - too large, 3 - too small, 4 - blank.
function intValidateWithRange(value, min, max) {
	if (value == '' || value == null) return 4;
	if (RegExValidate('^(-|)[0-9]*$', value, 'i')) {
		try { var pint = parseInt(value); } catch (ex) { return 1; }
		if (pint < min) return 3; else if (pint > max) return 2; return 0;
	} else
		return 1;
}

// Split function that remotes empty entries.
String.prototype.splitRemoveEmpties = function (separator, howmany) {
	var splitArr;
	var returnArr = new Array();

	if (arguments.length == 2)
		splitArr = this.split(separator, howmany);
	else
		splitArr = this.split(separator);

	for (var i = 0; i < splitArr.length; i++) {
		if (splitArr[i] != '')
			returnArr.push(splitArr[i]);
	}

	return returnArr;
}

// JavaScript equivalent of a function in GlobalFunctionsDetail.aspx.
// Do not use this function unless you have a good reason (e.g. textarea length on client must match size on server).
// KNOWN: Does not do entities properly, matches server behavior (where's the semi-colon?).
function SQLSafe(strInput) {
	return strInput.replace(/\'/g, "&#39").replace(/\"/g, "&#34");
}

// Returns true if the string is empty. Will blow up on NULLs.
function FieldIsEmpty(strInput) {
	if (strInput == undefined)
		return true;
	return TrimString(strInput).length == 0;
}

// Removes leading and trailing white space from a string. Will blow up on NULLs.
function TrimString(strInput) {
	return strInput.replace(/^\s+/g, "").replace(/\s+$/g, "");
}

// Returns true if the value is an integer value.
function isInteger(strInput) {
	var leadingZeros = calculateLeadingZeroStrings(strInput);

	return leadingZeros == strInput || (strInput == leadingZeros + parseInt(strInput).toString());
}

// Returns true if the value is a real number.
function isRealNumber(strInput) {
	var leadingZeros = calculateLeadingZeroStrings(strInput);

	return leadingZeros == strInput || (strInput == leadingZeros + parseFloat(strInput));
}

function calculateLeadingZeroStrings(strInput) {
	var leadingZeros = "";

	if (strInput != null) {
	for (var i = 0; i < strInput.length; i++) {
		if (strInput[i] == "0") {
			leadingZeros += "0";
		}
		else {
			break;
		}
	}
	}

	return leadingZeros;
}

// Date validator class.
function dateValidator() {
	this.firstValidDate = new Date('01/01/1753');
	this.lastValidDate = new Date('01/01/3000');
	this.strStartDateID = 'Start/Begin Date';
	this.strEndDateID = 'End/Stop/Expiration Date';
	this.strStartTimeID = 'Start Time';
	this.strEndTimeID = 'End Time';
	this.ysnRequireStartDateIfEndSpecified = false;
	this.ysnStartDateRequired = false;
	this.ysnEndDateRequired = false;
	this.ysnStartTimeRequired = false;
	this.ysnEndTimeRequired = false;
	this.ysnAllowEqualDates = false;
	this.ysnAllowTimeOnly = false;
	this.ysnCent = false; //Allow only four digit years
	var dtiEndDate;
	var dtiStartDate;
	var dtiStartTime;
	var dtiEndTime;
	this.timesAlreadyValidated = false;
	this.datesAlreadyValidated = false;

	this.setStartDate = function (date) {
		dtiStartDate = this.cleanDate(date);
	}

	this.setEndDate = function (date) {
		dtiEndDate = this.cleanDate(date);
	}

	this.setStartDateRequired = function (required) {
		this.ysnStartDateRequired = required;
	}

	this.setEndDateRequired = function (required) {
		this.ysnEndDateRequired = required;
	}

	this.setRequireStartDateIfEndSpecified = function (required) {
		this.ysnRequireStartDateIfEndSpecified = required;
	}

	this.cleanDate = function (date) {
		if (date) {
			if (typeof isMobileView != 'undefined' && isMobileView) {
				date = this.ChangeDateFormatForMobileView(date);
			}
			else if (RegExValidate('^([0-9-\\/]*?)([0-9]{2,4})$', date, 'i')) {
				var year = RegExp.$2;
				if (year.length == 2) {
					if (year >= 50) date = RegExp.$1 + "19" + year;
					else date = RegExp.$1 + "20" + year;
				}
				date = date.replace('\-', '/', 'g');
			}
		}
		return date;
	}

	this.ChangeDateFormatForMobileView = function (date) {
		var rxDatePattern = /^\d{4}\-\d{1,2}\-\d{1,2}$/; //Regex for yyyy-mm-dd format
		var match = date.match(rxDatePattern);
		if (match != null) {
			var dateArray = date.split('-');
			var dateFormat = getDateFormat().toLowerCase();
			if (dateFormat == "mm/dd/yyyy")
				date = dateArray[1] + '/' + dateArray[2] + '/' + dateArray[0];
			else if (dateFormat == "dd/mm/yyyy")
				date = dateArray[2] + '/' + dateArray[1] + '/' + dateArray[0];
		}
		return date;
	}

	this.dateValidate = function (dtiDate, ysnRequired, strID) {
		if (!ysnRequired && (dtiDate == null || dtiDate == '')) return true;
		else if (ysnRequired && (dtiDate == null || dtiDate == '')) {
			if (strID) this.error = strID + ' cannot be blank.';
			else this.error = ' cannot be blank';
			this.errorNumber = 1;
			return false;
		}
		else if (RegExValidate('^(1[0-2]|0?[1-9])(/|-)(0?[1-9]|[1-2][0-9]|3[0-1])\\2([0-9]{4}|[0-9]{2})$', dtiDate, 'i')) {
			var month = RegExp.$1;
			var day = RegExp.$3;
			var year = RegExp.$4;

			if (year.length == 2 && this.ysnCent == true) {
				if (strID) this.error = strID + ' requires a four digit year';
				else this.error = 'Please use a four digit year';
				return false;
			}
			if (year.length == 4 && (year > 3000 || year < 1753)) {
				this.error = dtiDate + '\nis outside of the date range.';
				this.errorNumber = 2;
				return false;
			}
			if (day == 31 && (month == 4 || month == 6 || month == 9 || month == 11)) {
				this.error = 'This month doesn\'t have 31 days';
				this.errorNumber = 3;
				return false;
			}
			if (day >= 30 && month == 2) {
				this.error = 'February doesn\'t have ' + day + ' days';
				this.errorNumber = 4;
				return false;
			}
			if (month == 2 && day == 29 && !(year % 4 == 0 && (year % 100 != 0 || year % 400 == 0))) {
				this.error = 'This is not a leap year\nFebruary doesn\'t have 29 days.';
				this.errorNumber = 5;
				return false;
			}
			return true;
		} else {
			if (strID) this.error = strID + ' is not a valid date format.\nPlease use ' + getDateFormat().toUpperCase() + '.';
			else this.error = dtiDate + '\n is an invalid date format.\nPlease use ' + getDateFormat().toUpperCase() + '.';
			this.errorNumber = 6;
			return false;
		}
		return false;
	}

	this.dateOrderValidate = function () {
		if (this.dateValidate(dtiStartDate, this.ysnStartDateRequired, this.strStartDateID) && this.dateValidate(dtiEndDate, this.ysnEndDateRequired, this.strEndDateID)) {
			if (this.ysnRequireStartDateIfEndSpecified && !dtiStartDate && dtiEndDate) {
				this.error = 'A Start Date must be specified if an End Date was entered.';
				this.errorNumber = 9;
				return false;
			}
			if (!dtiStartDate || !dtiEndDate) return true;
			else {
				var StartDate = new Date(dtiStartDate);
				var EndDate = new Date(dtiEndDate);
			}
			if (StartDate.getTime() < EndDate.getTime()) return true;
			else if (StartDate.getTime() == EndDate.getTime() && this.ysnAllowEqualDates == true) {
				this.ysnDatesAreEqual = true;
				return true;
			} else {
				this.error = 'The End Date must be after the Start Date.';
				this.errorNumber = 7;
				return false;
			}
		}
		else return false;
	}

	this.dateValidateNew = function (dtiDate, ysnRequired, strID) {
		var month;
		var day;
		var year;
		if (dtiDate === undefined)
			dtiDate = '';

        if (!ysnRequired && (dtiDate == '' || !dtiDate)) return true;

		if (ysnRequired && dtiDate == '') {
			if (strID) this.error = strID + ' is required';
			else this.error = ' is required';
			this.errorNumber = 1;
			return false;
		}

		var dateFormat = getDateFormat().toLowerCase();

		if (typeof isMobileView != 'undefined' && isMobileView) {
			dtiDate = this.ChangeDateFormatForMobileView(dtiDate);
		}

		if ((dateFormat == "mm/dd/yyyy") && (RegExValidate('^(1[0-2]|0?[1-9])(/|-)(0?[1-9]|[1-2][0-9]|3[0-1])\\2([0-9]{4}|[0-9]{2})$', dtiDate, 'i'))) {
			month = RegExp.$1;
			day = RegExp.$3;
			year = RegExp.$4;

			return this.validDateParts(month, day, year);
		}

		if ((dateFormat == "dd/mm/yyyy") && (RegExValidate('^(0?[1-9]|[1-2][0-9]|3[0-1])(/|-)(1[0-2]|0?[1-9])\\2([0-9]{4}|[0-9]{2})$', dtiDate, 'i'))) {
			month = RegExp.$3;
			day = RegExp.$1;
			year = RegExp.$4;

			return this.validDateParts(month, day, year);
		}

		if (strID) this.error = strID + ' is not a valid date format.\nPlease use ' + dateFormat.toUpperCase() + '.';
		else this.error = dtiDate + '\n is an invalid date format.\nPlease use ' + dateFormat.toUpperCase() + '.';
		this.errorNumber = 6;

		return false;
	}

	this.validDateParts = function (month, day, year) {
		if (year.length == 2 && this.ysnCent == true) {
			if (strID) this.error = strID + ' requires a four digit year';
			else this.error = 'Please use a four digit year';
			return false;
		}
		if (year.length == 4 && (year > 3000 || year < 1753)) {
			this.error = dtiDate + '\nis outside of the date range.';
			this.errorNumber = 2;
			return false;
		}
		if (day == 31 && (month == 4 || month == 6 || month == 9 || month == 11)) {
			this.error = 'This month doesn\'t have 31 days';
			this.errorNumber = 3;
			return false;
		}
		if (day >= 30 && month == 2) {
			this.error = 'February doesn\'t have ' + day + ' days';
			this.errorNumber = 4;
			return false;
		}
		if (month == 2 && day == 29 && !(year % 4 == 0 && (year % 100 != 0 || year % 400 == 0))) {
			this.error = 'This is not a leap year\nFebruary doesn\'t have 29 days.';
			this.errorNumber = 5;
			return false;
		}
		return true;
	}

	this.dateOrderValidateNew = function () {
		//If sending in US format run dateValidateNew first becuase it expects dates in daybeforemonth format
		if (this.datesAlreadyValidated || (this.dateValidateNew(dtiStartDate, this.ysnStartDateRequired, this.strStartDateID) && this.dateValidateNew(dtiEndDate, this.ysnEndDateRequired, this.strEndDateID))) {
			if (this.ysnRequireStartDateIfEndSpecified && !dtiStartDate && dtiEndDate) {
				this.error = 'A Start Date must be specified if an End Date was entered.'
				this.errorNumber = 9; return false;
			}
			if (!dtiStartDate || !dtiEndDate || dtiEndDate == "NaN/NaN/NaN" || dtiStartDate == "NaN/NaN/NaN") return true;
			else {
				//Expects dates in US format 
				var StartDate = new Date(dtiStartDate);
				var EndDate = new Date(dtiEndDate);
			}
			if (StartDate.getTime() < EndDate.getTime()) return true;
			else if (StartDate.getTime() == EndDate.getTime() && this.ysnAllowEqualDates == true) {
				this.ysnDatesAreEqual = true;
				return true;
			}
			else {
				this.error = 'The End Date must be after the Start Date.';
				this.errorNumber = 7; return false;
			}
		}
		else return false;
	}

	this.getStandardDate = function (dateText) {
		$.ajax({
			url: '/Utility/GetDate?dateText=' + dateText,
			async: false,
			type: 'GET',
			dataType: 'json',
			success: function (data) {
				dateText = data.date;
			},
			error: function (jqXHR, textStatus, errorThrown) {
			}
		});

		return dateText;
	}

	this.smallDateTimeMaxValueValidate = function () {
		var StartDate = new Date(dtiStartDate);
		var EndDate = new Date(dtiEndDate);
		var MaxDate = new Date('05/06/2079 23:59:59');
		var result = true;
		if (StartDate <= MaxDate && EndDate <= MaxDate) {
			result = true;
		}
		else if (StartDate > MaxDate) {
			this.error = "The Start Date must be less than '5/6/2079'.";
			this.errorNumber = 2;
			result = false;
		}
		else if (EndDate > MaxDate) {
			this.error = "The End Date must be less than '5/6/2079'.";
			this.errorNumber = 3;
			result = false;
		}
		return result;
	}

	this.format = function (format, date) {
		if (!date) date = new Date();
		var day = date.getDate();
		var month = date.getMonth() + 1;
		var year = date.getFullYear();
		return format.toLowerCase().replace(/dd/g, day).replace(/mm/g, month).replace(/y{1,4}/g, year)
	}

	this.timeValidate = function (ditTime, ysnRequired, strID) {
		if (!ysnRequired && ditTime == '') return true;
		else if (ysnRequired && ditTime == '') {
			if (strID) this.error = strID + ' is required';
			else this.error = 'Time is required';
			this.errorNumber = 1;
			return false;
		}
		else if (RegExValidate('^(1[0-2]|0?[1-9]):([0-5]?[0-9])(:([0-5][0-9]))?$', ditTime, 'i')) return true;
		else {
			if (strID) this.error = strID + ' is not a valid time format. Please use HH:MM.';
			else this.error = ' is not a valid time format. Please use HH:MM.';
			this.errorNumber = 8;
			return false;
		}
	}

	this.timeValidate24Hour = function (ditTime, ysnRequired, strID) {
		if (!ysnRequired && ditTime == '') return true;
		else if (ysnRequired && ditTime == '') {
			if (strID) this.error = strID + ' is required';
			else this.error = 'Time is required';
			this.errorNumber = 1;
			return false;
		}
		else if (RegExValidate('^([01]?[0-9]|2[0-3]):([0-5]?[0-9])(:([0-5][0-9]))?$', ditTime, 'i')) return true;
		else {
			if (strID) this.error = strID + ' is not valid.';
			else this.error = ' is not valid.';
			this.errorNumber = 8;
			return false;
		}
	}
	this.timeOrderValidate = function () {
		if (!this.ysnAllowTimeOnly && ((!dtiStartDate && this.dtiStartTime) || (!dtiEndDate && this.dtiEndTime))) {
			this.error = 'You only submited a time.\nPlease provide a day as well.';
			this.errorNumber = 10;
			this.startDateBlank = (!dtiStartDate && this.dtiStartTime);
			this.endDateBlank = (!dtiEndDate && this.dtiEndTime);
			return false;
		}
		if (this.timesAlreadyValidated || (this.timeValidate(this.dtiStartTime, this.ysnStartTimeRequired, this.strStartTimeID) && this.timeValidate(this.dtiEndTime, this.ysnEndTimeRequired, this.strEndTimeID))) {
			if (!this.ysnDatesAreEqual) return true;

			if (!this.dtiStartTime && !this.ysnStartTimeRequired) return true;
			if (!this.dtiEndTime && !this.ysnEndTimeRequired) return true;
			
			var dtiStartTime = this.convertTo24Hour(this.dtiStartTime, this.strStartAMPM, dtiStartDate);
			var dtiEndTime = this.convertTo24Hour(this.dtiEndTime, this.strEndAMPM, dtiEndDate);

			if (dtiStartTime.getTime() < dtiEndTime.getTime()) return true;
			else if (dtiStartTime.getTime() == dtiEndTime.getTime() && this.ysnAllowEqualTimes == true) {
				this.ysnTimesAreEqual = true;
				return true;
			}
			else {
				this.error = 'The End Time must be after the Start Time if the Start and End Dates are the same.'
				this.errorNumber = 9;
				return false;
			}
		}
		return false;
	}

	this.convertTo24Hour = function (time, AMPM, date) {
		if (!date) date = "1/1/70";

		var dtTime = time.indexOf(AMPM) == -1 ? new Date(date + " " + time + " " + AMPM) : new Date(date + " " + time);
		if (dtTime == 'Invalid Date') {
			var dayBeforeMonthOn = getDateFormat().toLocaleLowerCase() == "dd/mm/yyyy" ? true : false;
			if (dayBeforeMonthOn) {
				var dateArray = date.split('/');
				date = dateArray[1] + '/' + dateArray[0] + '/' + dateArray[2];
				dtTime = time.indexOf(AMPM) == -1 ? new Date(date + " " + time + " " + AMPM) : new Date(date + " " + time);
			}
		}
		return dtTime;
	}
}

// Helper function for inputAlert().
function inputEmailValidate(obj, required) {
	if (required == null || required == false)
		return (obj.value == '' || emailValidate(obj.value) == true)
	else
		return (obj.value != '' && emailValidate(obj.value) == true)
}

// Note: This function does not match standard validation behavior! Users are supposed to see
// a summary of problems with their inputs and not be alerted multiple times!
// -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
// Calls another validation routine, and determines success/failure on result of call (true/false).
// Displays error message on validation failure. Returns result of call.
function inputAlert(call, obj, required, errorMessage) {
	call = eval('input' + call + 'Validate');
	if (call(obj, required) == false) {
		if (errorMessage == null)
			errorMessage = 'Please enter a single valid email address without any extra body, subject, etc. information (ie user@domain.com).';
		obj.setAttribute('autocomplete', 'off');
		obj.focus();
		obj.setAttribute('autocomplete', 'on');
		alert(errorMessage);
		return false;
	}
	else
		return true;
}

// Takes array of required fields and checks whether they have a value or not
// If empty fields found, formats a javascript alert to inform the user and returns true
function checkRequiredFieldsEmpty(requiredFieldList) {
	var badList = new Array();
	for (var i = 0, len = requiredFieldList.length; i < len; i++) {
		var $this = requiredFieldList[i];
		var $fieldId = $('#' + $this);
		if (($fieldId.length > 0) && (FieldIsEmpty($fieldId.val()))) {
			var $label = $('label[for=' + $this + ']');
			if ($label.length > 0)
				badList.push($label.text().trim());
			else
				badList.push($this);
		}
	}
	if (badList.length > 0) {
		var msg = badList.join(" cannot be blank.\r\n");
		msg += " cannot be blank.";
		msg = msg.replace(/\*/g, '');
		alert(msg);
		return true;
	}
	return false;
}
// End Form Validation functions.

// Begin AJAX functions:
// See http://blogs.msdn.com/xmlteam/archive/2006/10/23/using-the-right-version-of-msxml-in-internet-explorer.aspx
// If you are curious why the XMLHTTP versions specified are and why others are not.
function makeErrorRequest(url, status) {
	var http_error_request = false;
	var origin = location.pathname;

	if (window.XMLHttpRequest) { // Mozilla, Safari, IE7
		http_error_request = new XMLHttpRequest();
		if (http_error_request.overrideMimeType)
			http_error_request.overrideMimeType("text/html");
	} else if (window.ActiveXObject) { // IE6
		try {
			http_error_request = new ActiveXObject("Msxml2.XMLHTTP.6.0");
		} catch (e) {
			try {
				http_error_request = new ActiveXObject("Msxml2.XMLHTTP"); // Version 3.0
			} catch (e) { }
		}
	}

	http_error_request.onreadystatechange = function () { }
	http_error_request.open("GET", "/AJAX-error.ashx?url=" + escape(url) + "&status=" + status + "&origin=" + escape(origin), true);
	http_error_request.send(null);
	//if(status == 403)
	//window.location = '/admin/AccessDenied.aspx?fromURL=' + window.location.pathname.substr(7);
}

// Makes AJAX request, additionally supporting POST data (makeHttpRequest doesn't).
// -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
// Note: It's YOUR responsibility as a developer to make sure data
// placed in formData is formatted as it should be. Form data should be
// in Query String format with the value portions escape()'d. -KB
function makeHttpRequestEx(URL, formData, callbackFunc, ysnReturnXML) {

	URL = window.location.origin + URL;

	if (formData != null) {
		$.ajax({
			type: 'POST',
			url: URL,
			data: formData,
			success: callbackFunc
		});
	}
	else {
		$.ajax({
			type: 'GET',
			url: URL,
			success: callbackFunc
		});
}
}

// NOTE: Do not use this in new code, use $.ajax() instead.
// Makes AJAX request. Used in n-menu code.
function makeHttpRequest(url, callback_function, return_xml) {
	var http_request = false;

	if (window.XMLHttpRequest) { // Mozilla, Safari, IE7
		http_request = new XMLHttpRequest();
		if (http_request.overrideMimeType && !return_xml)
			http_request.overrideMimeType("text/html");
	} else if (window.ActiveXObject) { // IE6
		try {
			http_request = new ActiveXObject("Msxml2.XMLHTTP.6.0");
		} catch (e) {
			try {
				http_request = new ActiveXObject("MSxml2.XMLHTTP"); // Version 3
			} catch (e) { }
		}
	}

	if (!http_request) {
		alert("Unfortunately, your browser doesn't support this feature.");
		return false;
	}

	http_request.onreadystatechange = function () {
		if (http_request.readyState == 4) {
			--ajax_call_counter;

			//try {
			if (http_request.status == 200) {
				if (return_xml)
					eval(callback_function + ', http_request.responseXML)');
				else {
					if (http_request.responseText.search(/<form action='error.asp' method=post>/i) < 0)
						eval(callback_function + ', http_request.responseText)');
					else
						makeErrorRequest(url, http_request.status);
				}
			}
			else if (http_request.status)
				makeErrorRequest(url, http_request.status);
			//} catch (e) {
			// If this happens, there is no way to notify anyone
			// of the previous error. It might be their connection died,
			// or failed in some other way.

			//alert('An error occured while reporting the previous error.');
			//var USER_EXPLAIN_ERR_EN = '\n\nTechnical Details:\n' + e.name + ' - ';
			// IE 6 & 7: No Line Number, No Filename
			//if (document.all && !window.opera) {
			//	alert(USER_EXPLAIN_ERR_EN + e.description + ' (line unavailable, code: ' + e.number + ')' + '\n\nJavaScript Callback:\n' + callback_function + ', <responseData>);');
			//}
			// Opera and Firefox 1.5+: Line Number + Filename
			//else {
			//	alert(USER_EXPLAIN_ERR_EN + e.message + '\n' + e.fileName + ', line ' + e.lineNumber + '\n\nJavaScript Callback:\n' + callback_function + ', <responseData>);');
			//}
			//}

			http_request = null;
		}
	}

	http_request.open("GET", url, true);
	http_request.send(null);
}
// Invoked on window.onunload for pages using synchronous AHAH.
// Catches switches to different pages and window closes.
// -=-=-=-=-
// Note: It is synchronous instead of asynchronous to ensure
// the operation completes before a close. This is SHAH.
//
function destroyAHAH(e) {
	var http_request;

	if (window.XMLHttpRequest) { // Mozilla, Safari, IE7
		http_request = new XMLHttpRequest();
		if (http_request.overrideMimeType) {
			http_request.overrideMimeType("text/html");
		}
	} else if (window.ActiveXObject) { // IE6
		try {
			http_request = new ActiveXObject("MSXML2.XMLHTTP.6.0");
		} catch (e) {
			try {
				http_request = new ActiveXObject("MSXML2.XMLHTTP"); // Version 3
			} catch (e) { }
		}
	}

	if (!http_request) return;
	http_request.open("GET", 'http://' + window.location.host + '/AJAX-NMenuLoader.ashx?clearCache=1', false);
	http_request.send(null);
	//alert(http_request.status + ': if 200, the session cache object was term\'d where needed be.');
}
// End AJAX functions.

//
//  slideShow class
//  vars:
//  slideShowId						constructor param for the id of the image that will have the Slide Show
//  slideShowLink					constructor param for the id of the link that will be around the image that has Slide Show, use '' for no link
//  slideShowSpeed					optional constructor param for the speed of the slide show, default = 5000
//  filterName						optional constructor param for the IE filter, default = blendTrans
//  filterAttr1						optional constructor param for the IE filter, default = duration=3
//  filterAttr2						optional constructor param for the IE filter
//  filterAttr3						optional constructor param for the IE filter
//  j								counter for the object
//  picArr							array of pictures
//  altArr							array of alt text
//  linkArr							array of links
//  methods:
//  addImage(n, img, alt, caption, lngCaption, link)	add an image with alt text and a link, n specifies which postion in the arrays
//  runSlideShow()					continually calls nextSlide to run the Slide Show
//  nextSlide()						applies the transition and advances the Slide Show image 1 spot in the arrays
//  mouseOver()						detects and loads that image to the Main Image Display & draws border around the slected image & removed border from the Other small image icons
function slideShow(slideShowId, slideShowLink, slideShowSpeed, filterName, filterAttr1, filterAttr2, filterAttr3, slideShowLinkTarget, layout, externalLink) {
	if (!slideShowLinkTarget || slideShowLinkTarget == null || slideShowLinkTarget == "undefined" || slideShowLinkTarget == "")
		slideShowLinkTarget = "_self";
	if (!layout || layout == null || layout == "undefined")
		layout = 1;

	this.slideShowId = document.getElementById(slideShowId);

	this.picArr = new Array();
	this.altArr = new Array();
	this.captionArr = new Array();
	this.lngCaptionArr = new Array();
	this.linkArr = new Array();
	this.linkTargetArr = new Array();
	var useStyle = false;
	var captionFilter = "";
	var cssRule = false;
	var pos = 0;

	//Get The filter for IE.//
	for (i = 0; i < document.styleSheets.length; i++) {
		var styles = document.styleSheets[i];
		do {
			if (styles.href && styles.href.indexOf("style.css") > -1) {
				if (styles.cssRules) {
					cssRule = styles.cssRules[pos];
				}
				else if (styles.rules) {
					cssRule = styles.rules[pos];
				}
			}
			if (cssRule && cssRule.selectorText != undefined && cssRule.selectorText.indexOf('slideCaption') != -1) {
				useStyle = true;
				captionFilter = cssRule.style.filter;
				break;
			}
			pos++;
		} while (cssRule && !useStyle)
		if (useStyle) break;
	}
	var useStyle = $('.slideCaption').attr('class') == 'slideCaption';

	// If junk data causes one slideshow to fail the others should not.
	if (this.slideShowId == null) {
		this.addImage = function (n, img, alt, link, linkTarget, caption, lngCaption, externalLink) { }
		this.runSlideShow = function () { }
		this.nextSlide = function () { }
		this.mouseOver = function (elem, imgPath, altText) { }
		return;
	}

	var tmpElem = document.createElement('a');
	if (slideShowLink && slideShowLink != '') {
		do {
			slideShowLink = (slideShowLink).replace("%38", "&");
		} while ((slideShowLink).indexOf("%38") != -1)
		tmpElem.href = slideShowLink;
		if (slideShowLinkTarget && slideShowLinkTarget != '')
			tmpElem.target = slideShowLinkTarget;

		var ssID = this.slideShowId.id.replace("cvpSlideShowImage", "");
		var capt = document.getElementById('slideCaption' + ssID);
		tmpElem.onclick = "";
		if (capt && capt.children && capt.children.length > 0)
			capt.children[0].removeAttribute("onclick");
	}
	else
		tmpElem.removeAttribute('href');

	this.j = 1;
	if (slideShowSpeed == null || slideShowSpeed == '')
		this.slideShowSpeed = 4000;
	else
		this.slideShowSpeed = slideShowSpeed * 1000;
	if (filterAttr1 == null || filterAttr1 == '')
		filterAttr1 = 2;
	if (filterName == null || filterName == '' || filterName.toLowerCase() == 'blendtrans')
		this.filterName = 'BlendTrans(duration=' + filterAttr1;
	else if (filterName.toLowerCase() != 'none')
		this.filterName = 'progid:DXImageTransform.Microsoft.' + filterName + ', duration=' + filterAttr1;
	if (filterAttr2 != null && filterAttr2 != '' && filterAttr2 != 'null')
		this.filterName = this.filterName + ',' + filterAttr2;
	if (filterAttr3 != null && filterAttr3 != '' && filterAttr3 != 'null')
		this.filterName = this.filterName + ',' + filterAttr3;
	if (filterName.toLowerCase() == 'none')
		this.filterName = filterName;
	else
		this.filterName = this.filterName + ')';

	switch (layout) {
		case 1: // Standard
			tmpElem.style.border = 'none';
			tmpElem.style.background = 'transparent';
			//tmpElem.style.border = 'medium none';
			//tmpElem.style.background = 'transparent none repeat scroll 0% 0%';
			tmpElem.id = slideShowId + '_link';
			this.slideShowLink = tmpElem;
			this.slideShowId.parentNode.insertBefore(tmpElem, this.slideShowId);
			this.slideShowId.parentNode.removeChild(this.slideShowId);
			this.slideShowLink.appendChild(this.slideShowId);
			break;
		case 5: // StandardFiveBottom
			//DO NOT COPY for this case - follow 12 for a better code.
			//tmpElem.style.border = 'none';
			//tmpElem.style.background = 'transparent';
			tmpElem.style.border = 'medium none';
			tmpElem.style.background = 'transparent none repeat scroll 0% 0%';
			tmpElem.id = slideShowId + '_link';
			this.slideShowLink = tmpElem;
			this.slideShowId.parentNode.insertBefore(tmpElem, this.slideShowId);
			//this.slideShowId.parentNode.removeChild(this.slideShowId);
			this.slideShowLink.appendChild(this.slideShowId);
			var tmpElemHTML = this.slideShowLink.innerHTML;

			var td = this.slideShowId.parentNode.parentNode;
			this.slideShowId.parentNode.parentNode.removeChild(this.slideShowId.parentNode);

			var newSection = document.createElement("div");
			newSection.setAttribute("id", "SubSection1");

			var cpSlideShowWrapper1 = document.createElement("div");
			cpSlideShowWrapper1.setAttribute("id", "cpSlideShowWrapper1");
			cpSlideShowWrapper1.style.padding = "5px";
			cpSlideShowWrapper1.style.background = "rgb(70, 43, 16) none repeat scroll 0% 0%";
			cpSlideShowWrapper1.style.position = "relative";
			cpSlideShowWrapper1.style.width = "416px";
			//cpSlideShowWrapper1.style.-moz-background-clip = "border";
			//cpSlideShowWrapper1.style.-moz-background-origin = "padding";
			//cpSlideShowWrapper1.style.-moz-background-inline-policy = "continuous";
			cpSlideShowWrapper1.style.height = "254px";

			var imageHead = document.createElement("div");
			imageHead.style.padding = "4px";
			imageHead.style.background = "rgb(241, 239, 234) none repeat scroll 0% 0%";
			imageHead.style.align = "left";
			imageHead.style.width = "406px";
			//-moz-background-clip: border;
			//-moz-background-origin: padding;
			//-moz-background-inline-policy: continuous;
			imageHead.style.float = "left";
			imageHead.style.height = "198px";
			imageHead.style.align = "left";

			imageHead.innerHTML = tmpElemHTML;
			cpSlideShowWrapper1.appendChild(imageHead);

			var imageRows = document.createElement("div");
			imageRows.setAttribute("id", "imageRows");
			imageRows.style.position = "relative";
			imageRows.style.align = "left";
			imageRows.style.margin = "5px 0px 0px 0px";
			imageRows.style.width = "416px";
			imageRows.style.float = "left";
			imageRows.style.height = "42px";
			imageRows.style.clear = "both";

			cpSlideShowWrapper1.appendChild(imageRows);
			newSection.appendChild(cpSlideShowWrapper1);
			td.appendChild(newSection);

			this.slideShowId = document.getElementById(slideShowId);
			if (isie) {
				this.slideShowId.style.height = "100%";
				this.slideShowId.style.width = "100%";
			}
			else {
				this.slideShowId.setAttribute("height", "100%");
				this.slideShowId.setAttribute("width", "100%");
			}
			break;
		case 12: //This case would handle the new 12 image slideshow without Rotation. & with mouseover
			tmpElem.id = slideShowId + '_link';
			this.slideShowLink = tmpElem;
			this.slideShowId.parentNode.insertBefore(tmpElem, this.slideShowId);
			this.slideShowLink.appendChild(this.slideShowId);
			var tmpElemHTML = this.slideShowLink.innerHTML;

			var td = this.slideShowId.parentNode.parentNode;

			// If the slideshow is in the middle of the page's content, we
			// have to keep track of its next sibling, so we put the 12-thumbnail
			// content in the new location.
			var nextSibling = this.slideShowLink.nextSibling;

			this.slideShowId.parentNode.parentNode.removeChild(this.slideShowId.parentNode);

			var newSection = document.createElement("div");
			newSection.setAttribute("id", "SubSectionFor" + slideShowId);

			var cpSlideShowWrapper1 = document.createElement("div");
			cpSlideShowWrapper1.setAttribute("id", "cpSlideShowWrapper" + slideShowId);

			//Remove these in-line style Height and Width after CSS changes
			cpSlideShowWrapper1.style.width = "441px";
			cpSlideShowWrapper1.style.height = "190px";
			cpSlideShowWrapper1.style.padding = "5px";
			cpSlideShowWrapper1.style.position = "relative";
			cpSlideShowWrapper1.style.clear = "both";

			var imageHead = document.createElement("div");
			imageHead.setAttribute("id", "imageHead");
			//Remove these in-line style Height and Width after CSS changes
			imageHead.style.width = "201px"; //185 px
			imageHead.style.height = "168px";
			imageHead.style.padding = "8px";
			imageHead.style.align = "left";
			imageHead.style.styleFloat = "left";
			imageHead.style.cssFloat = "left";
			imageHead.innerHTML = tmpElemHTML;
			cpSlideShowWrapper1.appendChild(imageHead);

			var imageRowsHeader = document.createElement("div");
			imageRowsHeader.setAttribute("id", "imageRowsHeader");
			imageRowsHeader.style.position = "relative";
			imageRowsHeader.style.align = "right";
			imageRowsHeader.style.styleFloat = "right";
			imageRowsHeader.style.cssFloat = "right";
			imageRowsHeader.style.width = "216px"; //247px
			imageRowsHeader.style.height = "18px";

			//Image Link
			var slideImageHeaderLink = document.createElement("a");
			slideImageHeaderLink.setAttribute("href", "/gallery.aspx");

			//Image Header
			var slideImageHeader = document.createElement("img");
			slideImageHeader.setAttribute("id", slideShowId + "Header");
			slideImageHeader.setAttribute("src", "/images/pages/ImageHeader.jpg");
			slideImageHeader.setAttribute("target", "_blank");
			slideImageHeader.style.width = "198px"; //247px
			slideImageHeader.style.height = "18px";
			slideImageHeader.setAttribute("border", "0");

			slideImageHeaderLink.appendChild(slideImageHeader);
			imageRowsHeader.appendChild(slideImageHeaderLink);

			cpSlideShowWrapper1.appendChild(imageRowsHeader);

			var imageRows = document.createElement("div");
			imageRows.setAttribute("id", "imageRows");
			imageRows.style.position = "relative";
			imageRows.style.align = "right";
			//imageRows.style.margin = "28px 0px 0px 0px";
			imageRows.style.styleFloat = "right";
			imageRows.style.cssFloat = "right";
			imageRows.style.width = "216px"; //247px
			imageRows.style.height = "145px";
			imageRows.style.padding = "4px";

			cpSlideShowWrapper1.appendChild(imageRows);
			newSection.appendChild(cpSlideShowWrapper1);

			var scriptTag = td.firstChild;
			while (scriptTag) {
				if (scriptTag.nodeType == 1 && scriptTag.tagName == 'SCRIPT')
					break;
				scriptTag = scriptTag.nextSibling;
			}
			if (nextSibling != null)
				td.insertBefore(newSection, nextSibling);
			else
				td.insertBefore(newSection, scriptTag);

			this.slideShowId = document.getElementById(slideShowId);
			//Remove these in-line style Height and Width after CSS changes
			if (isie) {
				this.slideShowId.style.height = "100%";
				this.slideShowId.style.width = "100%";
			}
			else {
				this.slideShowId.setAttribute("height", "100%");
				this.slideShowId.setAttribute("width", "100%");
			}
			break;
		default:
			break;
	}

	this.addImage = function (n, img, alt, link, linkTarget, caption, lngCaption, isExternal) {
		if (layout == 12 && n > 11)
			return; // Allow only upto 12 thumbnail images for Standard 12 Bottom layout

		if (!linkTarget || linkTarget == null || linkTarget == "undefined" || linkTarget == "")
			linkTarget = "_self";

		this.picArr[n] = new Image();
		this.picArr[n].src = img;
		do {
			alt = (alt).replace("%39", "'");
		} while ((alt).indexOf("%39") != -1)
		this.altArr[n] = alt;
		this.captionArr[n] = caption;
		this.lngCaptionArr[n] = lngCaption;

		do {
			link = (link).replace("%38", "&");
		} while ((link).indexOf("%38") != -1)
		this.linkArr[n] = link;

		this.linkTargetArr[n] = linkTarget;

		switch (layout) {
			case 5:
				var slideImage = document.createElement("img");
				slideImage.setAttribute("id", slideShowId + n);
				if (isie) {
					slideImage.style.height = "40";
					slideImage.style.width = "40";
				}
				else {
					slideImage.setAttribute("height", "40");
					slideImage.setAttribute("width", "40");
				}
				if (n == 0)
					slideImage.setAttribute("border", "1");
				else
					slideImage.setAttribute("border", "0");
				slideImage.style.borderStyle = "solid";
				slideImage.style.borderColor = "#FFFF00";
				slideImage.setAttribute("src", img);
				slideImage.setAttribute("alt", alt);
				slideImage.setAttribute("link", link);
				slideImage.setAttribute("target", linkTarget);
				slideImage.setAttribute("class", "photoGalleryImgArray");

				if (n == 0)
					slideImage.style.margin = "0px 0px 0px 0px";
				else
					slideImage.style.margin = "0px 0px 0px 6px";
				this.imageRows = document.getElementById("imageRows");
				this.imageRows.appendChild(slideImage);
				break;
			case 12: //This case would handle the new 12 image slideshow without Rotation. & with mouseover
				var slideImage = document.createElement("img");
				slideImage.setAttribute("id", slideShowId + n);
				//Remove these in-line style 'Padding' after CSS changes
				if (n == 0) {
					if (isie)
						slideImage.style.cssText = "margin:6px 6px; border: 2px solid transparent; background-color: lightblue; padding: 3px; height:30px; width:30px;";
					else
						slideImage.setAttribute('style', 'margin:6px 6px; border: 2px solid transparent; background-color: lightblue; padding: 3px; height:30px; width:30px;');
				}
				else {
					if (isie)
						slideImage.style.cssText = "margin:6px 6px; border: 2px solid transparent ; background-color: white; padding: 3px; height:30px; width:30px;";
					else
						slideImage.setAttribute('style', 'margin:6px 6px; border: 2px solid transparent ; background-color: white; padding: 3px;height:30px; width:30px;');
				}
				slideImage.setAttribute("src", img);
				slideImage.setAttribute("alt", alt);
				slideImage.setAttribute("link", link);
				slideImage.setAttribute("target", linkTarget);
				slideImage.setAttribute("class", "photoGalleryImgArray");
				var siObj = eval(slideShowId.replace("Image", ""));
				slideImage.onmouseover = function () { siObj.mouseOver(slideShowId + n, img, alt); };

				this.imageRows = document.getElementById("imageRows");
				this.imageRows.appendChild(slideImage);
				break;
			default:
				break;
		}
	}

	this.mouseOver = function (elem, imgPath, altText) {
		//Set the Main Display image to mouseOver Image
		this.slideShowId.setAttribute("src", imgPath);
		this.slideShowId.alt = altText;

		//Loop through all other images and set the border to 1
		this.ImageRows = document.getElementById("imageRows");
		this.ImageRows.Images = this.ImageRows.getElementsByTagName("img");
		var j = 0;
		while (j < this.ImageRows.Images.length) {
			if (this.ImageRows.Images[j] == document.getElementById(elem)) {
				if (isie)
					this.ImageRows.Images[j].style.cssText = "margin:6px 6px; border: 2px solid transparent; background-color: lightblue; padding: 3px; height:30px; width:30px;";
				else
					this.ImageRows.Images[j].setAttribute('style', 'margin:6px 6px; border: 2px solid transparent; background-color: lightblue; padding: 3px;height:30px; width:30px;');
			}
			else {
				if (isie)
					this.ImageRows.Images[j].style.cssText = "margin:6px 6px; border: 2px solid transparent; background-color:white; padding: 3px;height:30px; width:30px;";
				else
					this.ImageRows.Images[j].setAttribute('style', 'margin:6px 6px; border: 2px solid transparent; background-color:white; padding: 3px;height:30px; width:30px;');
			}
			j++;
		}
	}

	this.runSlideShow = function () {
		switch (layout) {
			case 12:
			case 5:
				//Do nothing, do not run a slideshow
				break;
			default:
				if (this.picArr.length > 1) {
					var self = this; 					//  reference to get around context loss of this during setTimeout()
					this.timeoutId = setTimeout(function () { self.nextSlide(); self.runSlideShow(); }, this.slideShowSpeed);
				}
				break;
		}
	}

	this.nextSlide = function () {
		var ssID = this.slideShowId.id.replace("cvpSlideShowImage", "");
		var capt = document.getElementById('slideCaption' + ssID);
		var lngCapt = document.getElementById('slideCaptionLng' + ssID);
		if (document.all && this.slideShowId.filters && this.slideShowId.filters.length > 0 && this.filterName.toLowerCase() != 'none') {
			this.slideShowId.style.filter = this.filterName;
			this.slideShowId.filters.item(0).apply();
			if (capt) {
				capt.style.filter = captionFilter + " " + this.filterName;
				capt.filters.item(capt.filters.length - 1).apply();
				if (lngCapt) {
					lngCapt.style.filter = captionFilter + " " + this.filterName;
					lngCapt.filters.item(lngCapt.filters.length - 1).apply();
				}
			}
		}
		this.slideShowId.setAttribute('src', this.picArr[this.j].src);
		this.slideShowId.setAttribute('alt', this.altArr[this.j]);
		this.slideShowId.setAttribute('title', this.altArr[this.j]);

		var height = '';
		var marginTop = useStyle ? "-61px" : "0px";
		var paddingTop = "3px";
		if (this.captionArr[this.j] == "") {
			height = "0px";
		}
		else if (this.lngCaptionArr[this.j] == "") {
			height = "31px";
			marginTop = useStyle ? "-35px" : "0px";
			paddingTop = useStyle ? "-4px" : "0px";
		}
		else
			height = useStyle ? "55px" : "";

		if (capt) {
			capt.innerHTML = useStyle ? this.captionArr[this.j] : "<strong>" + this.captionArr[this.j] + "</strong>";
			capt.style.marginTop = marginTop;
			capt.style.height = height;
			capt.style.paddingLeft = useStyle ? "15px" : "5px";
			capt.style.paddingRight = useStyle ? "10px" : "5px";
			capt.style.paddingTop = paddingTop;

			if (height != "0px" && lngCapt != null) {
				lngCapt.innerHTML = this.lngCaptionArr[this.j];
				lngCapt.style.marginTop = useStyle ? "-38px" : "0px";
				lngCapt.style.paddingLeft = useStyle ? "15px" : "5px";
				lngCapt.style.paddingRight = useStyle ? "10px" : "5px";
				lngCapt.style.paddingTop = paddingTop;
				lngCapt.style.overflow = 'hidden';
			}
			else if (lngCapt != null) {
				lngCapt.style.marginTop = "0px";
				lngCapt.innerHTML = "";
				lngCapt.height = "0px";
			}
		}

		var slideLink = this.linkArr[this.j]
		var slideLinkTarget = this.linkTargetArr[this.j]
		if (slideLink && slideLink != '') {
			tmpElem.href = slideLink;
			this.slideShowLink.removeAttribute("onclick");
			tmpElem.onclick = "";
			if (capt)
				capt.innerHTML = useStyle ? "<a class = 'slideCaptionHyperlink' href=" + slideLink + " target = " + slideLinkTarget + ">" + this.captionArr[this.j] + "</a>" :
                                             "<a class = 'Hyperlink' href=" + slideLink + " target = " + slideLinkTarget + ">" + this.captionArr[this.j] + "</a>";
			if (slideLinkTarget && slideLinkTarget != '')
				tmpElem.target = slideLinkTarget;
		}
		else {
			tmpElem.removeAttribute('href');
			if (capt)
				capt.removeAttribute('href');
		}
		if (document.all && this.slideShowId.filters && this.slideShowId.filters.length > 0 && filterName.toLowerCase() != 'none') {
			this.slideShowId.filters.item(0).play();
			if (capt) {
				capt.filters.item(capt.filters.length - 1).play();
				lngCapt.filters.item(lngCapt.filters.length - 1).play();
			}
		}

		var slideImageId, slideImage;
		if (layout == 5) {
			// Set the selected image's border
			if (this.j == 0)
				slideImageId = slideShowId + (this.picArr.length - 1);
			else
				slideImageId = slideShowId + (this.j - 1);
			slideImage = document.getElementById(slideImageId);
			slideImage.setAttribute("border", "0");

			slideImageId = slideShowId + this.j;
			slideImage = document.getElementById(slideImageId);
			slideImage.setAttribute("border", "1");
		}

		this.j++;
		if (this.j >= this.picArr.length)
			this.j = 0;
	}
}

function validateMultiFileUpload(el, useWhiteList) {
	try {
		const fileNames = [...el.files].map(file => file.name);
		const xhr = new XMLHttpRequest();
		xhr.open('POST', '/Utility/AreValidFileNames', false);
		xhr.setRequestHeader('X-Requested-With', 'XMLHttpRequest');
		const data = new FormData();
		fileNames.forEach(fileName => {
			data.append("paths[]", fileName);
		})

		xhr.send(data);
		const response = JSON.parse(xhr.responseText);
		const isValid = response.every(x => x === true);

		for (const fileName of fileNames) {
			if (!isAllowedFileExtension(fileName, useWhiteList) || !isValid) {
				return false;
			}
		}
		return true;
	} catch (e) {
		console.error(e);
		return false;
	}
}

//el is the pointer to the input type="file"
//example usage: <input type="file" size="38" name="txtFile" onChange="validateFileUpload(this);">
function validateFileUpload(el, useWhiteList) {
	var val = el.value;
	var cbox = document.getElementById(el.name + "_pdf");
	if (cbox) {
		cbox.disabled = true;
		cbox.checked = false;
	}
	if (val.length > 0) {
		var isValid = false;
		if (val.indexOf(':\\fakepath\\') > -1)
			val = val.substring(12);
		$.ajax({
			async: false,
			type: "POST",
			url: '/Utility/IsValidFileName',
			data: { path: val },
			dataType: "json",
			success: function (response) {
				if (typeof (response.d) != "undefined") {
					response = response.d;
				}
				if (response.result == "True")
					isValid = true;
			}
		});
		if (!isAllowedFileExtension(val, useWhiteList) || !isValid) {
			alert("The file you are trying to upload is not an allowed file type or has an invalid file name. These are the accepted characters: A-z 0-9 ~ ! ( ) - + = [ ] { } , . _");
			var tempElem = el.cloneNode(false);
			tempElem.value = "";
			el.parentNode.replaceChild(tempElem, el);
			el.focus();
			return false;
		}
		if (cbox) {
			cbox.disabled = !isAllowedFileExtensionPDF(val);
			if (cbox.disabled) cbox.checked = false;
		}
	}
	return true;
}

// Recommended you use isAllowedFileExtensionWhiteList() instead. Do not use this method
// where useWhiteList is false or null, that is not secure (uses blacklist rather than whitelist).
function isAllowedFileExtension(filename, useWhiteList) {
	if (filename + '' == '')
		return true;

	// Cut out the path portion (why not use lastIndexOf()?).
	while (filename.indexOf("/", 0) > 0)
		filename = filename.substr(filename.indexOf("/", 0) + 1, filename.length - filename.indexOf("/", 0))
	while (filename.indexOf("\\", 0) > 0)
		filename = filename.substr(filename.indexOf("\\", 0) + 1, filename.length - filename.indexOf("\\", 0))

	var lastPeriod = filename.indexOf(".", 0);
	var i = lastPeriod;

	while (lastPeriod > 0) {
		i = lastPeriod;
		lastPeriod = filename.indexOf(".", i + 1);
	}

	var fileExtension = (i > 0 ? filename.substr(i + 1, filename.length - i) : "");

	if (useWhiteList)
		return isAllowedFileExtensionWhiteList('.' + fileExtension)
	else {
		switch (fileExtension.toUpperCase()) {
			case "ASA":
			case "ASAX":
			case "ASBX":
			case "ASCX":
			case "ASP":
			case "ASPX":
			case "BAT":
			case "CAB":
			case "CF":
			case "CFM":
			case "CGI":
			case "COM":
			case "CONFIG":
			case "DLL":
			case "EXE":
			case "HTA":
			case "LHA":
			case "LHZ":
			case "MIM":
			case "PIF":
			case "PL":
			case "SYS":
			case "UUE":
			case "VBS":
			case "VXD":
			case "WEBINFO":
			case "WIZ":
			case "WSH":
				{
					return false;
					break;
				}
			default:
				{
					return true;
					break;
				}
		}
	}
}

// Makes sure a file being uploaded for PDF conversion is supported.
// Note: Server logic should additionally chop out paths and other evil things (e.g. Windows Device Names like 'CON', 'AUX', etc).
function isAllowedFileExtensionPDF(filename) {
	//cut out the path portion
	while (filename.indexOf("/", 0) > 0)
		filename = filename.substr(filename.indexOf("/", 0) + 1, filename.length - filename.indexOf("/", 0))
	while (filename.indexOf("\\", 0) > 0)
		filename = filename.substr(filename.indexOf("\\", 0) + 1, filename.length - filename.indexOf("\\", 0))

	var lastPeriod, i, fileExtension;
	lastPeriod = filename.indexOf(".", 0);
	i = lastPeriod;
	while (lastPeriod > 0) {
		i = lastPeriod;
		lastPeriod = filename.indexOf(".", i + 1);
	}

	if (i > 0)
		fileExtension = filename.substr(i + 1, filename.length - i).toUpperCase();
	else
		fileExtension = "";

	switch (fileExtension) {
		case "HTM":
		case "HTML":
		case "DOC":
		case "DOCX":
		case "XLS":
		case "XLSX":
		case "TXT":
			{
				return true;
				break;
			}
		default:
			{
				return false;
				break;
			}
	}
}

// returns true if it is a valid US or Canadian zip code.
// supports ZIP, ZIP+4 and Canadian zip code format
// * intCountryCode is a ISO 3166-1 numeric value.
function isZipCode(strInput, intCountryCode) {
	var countryCodeRe = {
		840: /^\d{5}([\-]\d{4})?$/, // USA
		124: /^[ABCEGHJ-NPRSTVXY]\d[ABCEGHJ-NPRSTV-Z][ -]?\d[ABCEGHJ-NPRSTV-Z]\d$/i, //Canada
		36: /^(?:(?:[2-8]\d|9[0-7]|0?[28]|0?9(?=09))(?:\d{2}))$/ //Australia
	};
	if (intCountryCode in countryCodeRe) {
		return countryCodeRe[intCountryCode].test(strInput);
	}
	return strInput.length > 0;
}

// returns name to assign postal/zip code.
function getZipCodeName(intCountryCode) {
	if (intCountryCode == 840) // USA
		return 'Zip';
	else// if (intCountryCode == 124) // Canada
		return 'Postal';
}

function updateAddressPlaceholdersForCanadinaClients(zipElementId, stateElementId) {
	document.getElementById(stateElementId).setAttribute('placeholder', 'Province');
	document.getElementById(stateElementId).style.width = '5em';
	document.getElementById(zipElementId).setAttribute('placeholder', 'Postal Code');
}

function getLowerCaseStateString() {
	if (intCountryCode == 124)
		return "province";
	else// if (intCountryCode == 840) <- neutral setting (en-US).
		return "state";
}

// Gets card type. Will return one of the following values (or an empty string for unknown/invalid cards):
// * Master = M
// * Visa = V
// * Discover = D
// * American Express = A
function getCreditCardType(cardNumber) {
	switch (cardNumber.substr(0, 2)) {
		case "50":
		case "51":
		case "52":
		case "53":
		case "54":
		case "55":
			return "M";
		case "65":
			return "D";
	}
	switch (cardNumber.substr(0, 1)) {
		case "4":
			return "V";
	}
	switch (cardNumber.substr(0, 4)) {
		case "6011":
			return "D";
	}
	switch (cardNumber.substr(0, 2)) {
		case "34":
		case "37":
			return "A";
	}

	// Card unknown/invalid.
	return '';
}

// Validates credit card numbers. cardType is an optional argument.
function isValidCreditCardNumber(cardNumber, cardType) {
	var isValid = false;
	var ccCheckRegExp = /[^\d ]/;

	isValid = !ccCheckRegExp.test(cardNumber);

	if (isValid) {
		// Get card type if not explicitly specified.
		if (cardType == null)
			cardType = getCreditCardType(cardNumber);

		// If unknown/unsupported/invalid card type, abort/fail.
		if (cardType == '')
			return false;

		var cardNumbersOnly = cardNumber.replace(/ /g, "");
		var cardNumberLength = cardNumbersOnly.length;
		var lengthIsValid = false;
		var patternIsValid = false;
		var patternRegExp;

		switch (cardType) {
			case "V":
				{ // visa
					lengthIsValid = (cardNumberLength == 13 || cardNumberLength == 16);
					patternRegExp = /^4/;
					break;
				}
			case "M":
				{ // mastercard
					lengthIsValid = (cardNumberLength == 16);
					patternRegExp = /^5[0-5]/;
					break;
				}
			case "A":
				{ // amercian express
					lengthIsValid = (cardNumberLength == 15);
					patternRegExp = /^3[4,7]/;
					break;
				}
			case "D":
				{ // discover
					lengthIsValid = (cardNumberLength == 16);
					patternRegExp = /^6011/;
					break;
				}
				/*case "J":
				{ // jcb
				lengthIsValid = (cardNumberLength == 15) || (cardNumberLength == 16);
				patternRegExp = /^[3,1800,2131]/;
				break;
				}
				case "C":
				{ // diner's club
				lengthIsValid = (cardNumberLength == 14);
				patternRegExp = /^3[0,6,8]/;
				break;
				}*/
			default:
				patternRegExp = /^$/;
		}

		patternIsValid = patternRegExp.test(cardNumbersOnly);
		isValid = patternIsValid && lengthIsValid;
	}

	if (isValid) {
		var numberProduct;
		var numberProductDigitIndex;
		var checkSumTotal = 0;

		for (digitCounter = cardNumberLength - 1; digitCounter >= 0; digitCounter--) {
			checkSumTotal += parseInt(cardNumbersOnly.charAt(digitCounter));
			digitCounter--;
			numberProduct = String((cardNumbersOnly.charAt(digitCounter) * 2));
			for (var productDigitCounter = 0; productDigitCounter < numberProduct.length; productDigitCounter++)
				checkSumTotal += parseInt(numberProduct.charAt(productDigitCounter));
		}

		isValid = (checkSumTotal % 10 == 0);
	}

	return isValid;
}

// Confirms data given is legal.
function isDate(varDateToCheck) {
	var objDate = new dateValidator();
	return objDate.dateValidate(varDateToCheck, true, 'Date');
}

// Confirms data given is legal.
function isDateNew(varDateToCheck) {
	var objDate = new dateValidator();
	return objDate.dateValidateNew(varDateToCheck, true, 'Date');
}

// TODO: JavaScript - Remove this function. Better ways to handle this.
// Date comparison function.
function dateDiff(dtiDateEntered, dtiDateToday) {
	var date1 = new Date();
	var date2 = new Date();
	var diff = new Date();
	// Initialize the first date
	var date1temp = new Date(dtiDateEntered);
	date1.setTime(date1temp.getTime());
	// Initialize the second date
	var date2temp = new Date(dtiDateToday);
	date2.setTime(date2temp.getTime());
	// sets difference date to difference of first date and second date
	diff.setTime(date1.getTime() - date2.getTime());
	timediff = diff.getTime();
	days = Math.floor(timediff / (86400000)); // 1000 * 60 * 60 * 24
	timediff -= days * (86400000);
	return days;
}

// .NET JavaScript debugging function.
// TO-DO: Add sections for 'add_' and 'remove_' event prefixes.
function debugDisplayMembers(obj) {
	var getters = new Array();
	var setters = new Array();
	var others = new Array();
	var sKey;

	for (var key in obj) {
		sKey = key + "";
		if (!((sKey.length > 0) && (sKey[0] == '_'))) {
			switch (sKey.substr(0, 4)) {
				case 'get_':
					getters.push(key);
					break;
				case 'set_':
					setters.push(key);
					break;
				default:
					others.push(key);
			}
		}
	}

	getters.sort();
	setters.sort();
	others.sort();

	alert(others.join(", ") + "\r\n\r\n" +
		getters.join(", ") + "\r\n\r\n" +
		setters.join(", "));
}

// BEGIN PSEUDO-FORMS CODE
// Do not call any of these methods directly in your application code.
// Pseudo-forms are transparent, you shouldn't need to use these methods.
addEvent(window, 'load', initializePseudoForms, false);

// Allows IE to recognize tag "subform". You cannot legally nest
// FORM tags inside FORM tags and at least Firefox pretends they're not
// in the DOM period. This is problematic, so the tag was renamed.
// Done for FORMS inside of FORMs trick to get around ASP.NET form limitations
// for content we cannot control (User Pages, Info Advanced HTML, etc).
document.createElement('subform');

// Initialize pseudo-forms. Called on window load.
function initializePseudoForms() {
	// Make sure we're in an HTML document (not XUL).
	if (document.body) {
		var subForms = document.body.getElementsByTagName('subform');
		for (var i = 0; i < subForms.length; i++)
			setupPseudoForm(subForms[i]);
	}
}

// Sets up pseudo-form.
function setupPseudoForm(pseudoFormElement) {
	// Re-routes form submission elements (no javascript used/necessary).
	// (those being input.type=submit, button.type=submit, and input.type=image)
	var rerouteElement = function (input, pseudoFormElement, newTagName, newType, copyInnerHTML) {
		var replacement = document.createElement(newTagName);
		replacement.setAttribute('type', newType);
		var onClickDest, onClickSrc = replacement.getAttribute('onclick');
		if (onClickSrc != null)
			replacement.removeAttribute('onclick');
		cloneAttributes(input, replacement);
		if (onClickSrc != null) {
			eval("window.tmp = function(event) { " + onClick + " }");
			onClickSrc = window.tmp;
			window.tmp = null;
			onClickDest = function (ev) {
				ev = (window.event ? window.event : ev);
				if (onClickSrc(ev) == false)
					return;
				pseudoFormElement._submitter = replacement;
				pseudoFormElement._submevent = ev;
				pseudoFormElement.submit();
			}
		}
		else {
			onClickDest = function (ev) {
				ev = (window.event ? window.event : ev);
				pseudoFormElement._submitter = replacement;
				pseudoFormElement._submevent = ev;
				pseudoFormElement.submit();
			}
		}

		// Return replacement routine.
		var fn = function () {
			var inputParent = input.parentNode;
			var inputRightAfter = input.nextSibling;
			var inputDisabled = input.disabled; // For Internet Explorer (it forgets).
			var inputHTML = (copyInnerHTML ? input.innerHTML : null);
			inputParent.removeChild(input); input = null;
			if (copyInnerHTML)
				replacement.innerHTML = inputHTML;
			inputParent.insertBefore(replacement, inputRightAfter);
			replacement.disabled = inputDisabled;
			addEvent(replacement, 'click', onClickDest, false);
		}
		return fn;
	}

	// Re-route submits (needs to be done early on).
	var inputs, input, inc, repls;
	repls = new Array();
	inputs = pseudoFormElement.getElementsByTagName('input');
	for (var i = 0; i < inputs.length; i++) {
		input = inputs[i];
		switch (input.type.toLowerCase()) {
			case 'submit':
				repls.push(rerouteElement(input, pseudoFormElement, 'input', 'button', false));
				break;
			case 'image':
				repls.push(rerouteElement(input, pseudoFormElement, 'img', 'image', false));
				break;
		}
	}
	inputs = pseudoFormElement.getElementsByTagName('button');
	for (var i = 0; i < inputs.length; i++) {
		input = inputs[i];
		switch (input.type.toLowerCase()) {
			case 'submit':
				repls.push(rerouteElement(input, pseudoFormElement, 'button', 'button', true));
				break;
		}
	}
	for (var i = 0; i < repls.length; i++)
		repls[i]();

	// Implement properties for pseudo-form.
	var name = pseudoFormElement.getAttribute('name');
	if (name && name != '') {
		document[name] = pseudoFormElement;
		document.forms[name] = pseudoFormElement;
	}
	pseudoFormElement.method = pseudoFormElement.getAttribute('method');
	if (pseudoFormElement.method == null || pseudoFormElement.method == '')
		pseudoFormElement.method = 'get';
	pseudoFormElement.enctype = pseudoFormElement.getAttribute('enctype');
	if (pseudoFormElement.enctype == null || pseudoFormElement.enctype == '')
		pseudoFormElement.enctype = 'application/x-www-form-urlencoded';
	pseudoFormElement.target = pseudoFormElement.getAttribute('target');
	pseudoFormElement.action = pseudoFormElement.getAttribute('action');
	pseudoFormElement.acceptCharset = pseudoFormElement.getAttribute('accept-charset');

	// Initialize.
	var elements = new Array(); // elements in pseudo-form (for form.elements).
	var simpleInputs = new Array(); // simple value inputs and textareas, no additional logic.
	var stateInputs = new Array(); // checkboxes and radiobuttons.
	var dropdowns = new Array(); // drop-downs and lists.
	pseudoFormElement.elements = elements;
	pseudoFormElement._simple = simpleInputs;
	pseudoFormElement._state = stateInputs;
	pseudoFormElement._dropdown = dropdowns;

	// Registers element for the pseudo-form.
	var registerElement = function (input) {
		// Add to elements array.
		elements.push(input);
		// Add as property.
		if (input.name != '' && input.name != null) {
			obj = pseudoFormElement[input.name];
			if (obj == null)
				pseudoFormElement[input.name] = input;
			else if (obj instanceof Array)
				obj.push(input);
			else {
				var arr = new Array();
				arr.push(obj);
				arr.push(input);
				pseudoFormElement[input.name] = arr;
			}
		}
	}

	// Add methods (can be overridden by form elements /w same names, just like real forms).
	pseudoFormElement.submit = function () { submitPseudoForm(this, false); }
	pseudoFormElement.reset = document.aspnetForm.reset; // HACK (we'll fix it if someone complains)

	// Register elements, and categorize them for form submission.
	inputs = pseudoFormElement.getElementsByTagName('input');
	for (var i = 0; i < inputs.length; i++) {
		input = inputs[i];
		switch (input.type.toLowerCase()) {
			case 'checkbox':
			case 'radio':
				stateInputs.push(input);
				registerElement(input);
				break;
			case 'text':
			case 'hidden':
			case 'password':
				simpleInputs.push(input);
				registerElement(input);
				break;
			default:
				registerElement(input);
		}
	}
	inputs = pseudoFormElement.getElementsByTagName('textarea');
	for (var i = 0; i < inputs.length; i++) {
		input = inputs[i];
		simpleInputs.push(input);
		registerElement(input);
	}
	inputs = pseudoFormElement.getElementsByTagName('select');
	for (var i = 0; i < inputs.length; i++) {
		input = inputs[i];
		dropdowns.push(input);
		registerElement(input);
	}
	inputs = pseudoFormElement.getElementsByTagName('button');
	for (var i = 0; i < inputs.length; i++)
		registerElement(inputs[i]);
}

// Copies attributes from one element to another. Copy should occur
// before destElement is added to the DOM, as some things become read-only.
// ID and CLASS need to be handled specially for compatiblity reasons.
// Warning: TYPE is not supported (blame IE)!
function cloneAttributes(srcElement, destElement) {
	var attr;
	for (var i = 0; i < srcElement.attributes.length; i++) {
		attr = srcElement.attributes[i];
		switch (attr.name.toLowerCase()) {
			case 'name':
				if ((attr.value + '').toLowerCase() == 'submit')
					destElement.setAttribute('name', 'submit_button');
				else
					destElement.setAttribute('name', attr.value);
				break;
			case 'class':
				destElement.className = attr.value;
				break;
			case 'id':
				destElement.id = attr.value;
				break;
			case 'type':
			case 'form':
				break;
			default:
						try { destElement.setAttribute(attr.name, attr.value); }
						catch (e) { }
		}
	}
}

// Called by pseudo-form when submit() is invoked.
function submitPseudoForm(pseudoFormElement, debugMode) {
	var inputs, input;
	var spawner = pseudoFormElement._submitter;
	var spawnerEvent = pseudoFormElement._submevent;
	pseudoFormElement._submitter = null;
	pseudoFormElement._submevent = null;

	// Destroy form if it already exists (possible if target not implicit/explicit "_self").
	if (pseudoFormElement._form != null)
		document.body.removeChild(pseudoFormElement._form);
	pseudoFormElement._form = null;

	// Fire pseudo-onsubmit event (remember that just like in real forms, only fires on non-J/S submits).
	var onsubmit = pseudoFormElement.getAttribute('onsubmit');
	if (onsubmit != null && onsubmit != '' && spawnerEvent) {
		eval("window.tmp = function(event) { " + onsubmit + " }");
		onsubmit = window.tmp;
		window.tmp = null;
		if (onsubmit(spawnerEvent) == false)
			return false;
	}

	// Create actual form to forward submission to.
	var newForm = document.createElement('form');
	pseudoFormElement._form = newForm;

	// Transfer some standard HTML/XHTML attributes.
	newForm.dir = pseudoFormElement.dir;
	newForm.lang = pseudoFormElement.lang;
	newForm.title = pseudoFormElement.title;

	// Setup actual form action, accept, accept-charset, encoding, method, and target.
	var action = pseudoFormElement.getAttribute('action');
	var method = pseudoFormElement.getAttribute('method');
	var target = pseudoFormElement.getAttribute('target');
	var enctype = pseudoFormElement.getAttribute('enctype');
	var accept = pseudoFormElement.getAttribute('accept');
	var acceptCharset = pseudoFormElement.getAttribute('accept-charset');
	if (action)
		newForm.action = action;
	if (method)
		newForm.method = method;
	if (target)
		newForm.target = target;
	if (enctype)
		newForm.enctype = enctype;
	if (accept)
		newForm.accept = accept;
	if (acceptCharset)
		newForm.acceptCharset = acceptCharset;

	// Method used to create hidden input copies.
	var newHiddenInput = function (name, value) {
		var ret = document.createElement('input');
		ret.type = 'hidden';
		ret.value = value;
		ret.name = name;
		newForm.appendChild(ret);
	}

	// Create element for spawner if one exists (submitted w/o JS, e.g. input.type=submit).
	if (spawner && spawner.name != null && spawner.name != '')
		newHiddenInput(spawner.name, spawner.value);

	// Forward text, textarea, hidden, and password input values.
	inputs = pseudoFormElement._simple;
	for (var i = 0; i < inputs.length; i++) {
		input = inputs[i];
		if ((!input.disabled) && (input.style.display != 'none'))
			newHiddenInput(input.name, input.value);
	}

	// Forward checkbox and radio-button values.
	inputs = pseudoFormElement._state;
	for (var i = 0; i < inputs.length; i++) {
		input = inputs[i];
		if ((!input.disabled) && (input.style.display != 'none')) {
			if (input.checked)
				newHiddenInput(input.name, input.value);
		}
	}

	// Forward drop-down/list selections.
	// Blame for text/value condition and necessity of options loop lies with IE.
	// 1) Specifically, IE 7  and IE 7 mode don't give the implicit value if it's left off.
	// 2) IE 8 should support retrieval of all values from the select.value prop, but doesn't.
	inputs = pseudoFormElement._dropdown;
	for (var i = 0; i < inputs.length; i++) {
		input = inputs[i];
		if ((!input.disabled) && (input.style.display != 'none')) {
			for (var j = 0; j < input.options.length; j++) {
				if (input.options[j].selected) {
					if (input.options[j].value == '')
						newHiddenInput(input.name, input.options[j].text);
					else
						newHiddenInput(input.name, input.options[j].value);
				}
			}
		}
	}

	document.body.appendChild(newForm);
	if (debugMode)
		alert(newForm.innerHTML);
	newForm.submit();
}
// END PSEUDO-FORMS CODE

function phoneStripFormatting(value) {
	/// <summary>Strips formatting characters (for display purposes only) from the NANP phone number given.</summary>
	/// <param name="value" type="String">Phone number.</summary>
	/// <returns type="String">Phone number with formatting cahracter removed.</returns>
	return (value + '').replace(/\(/g, '').replace(/\)/g, '').replace(/\s/g, '').replace(/\-/g, '');
}

function phoneReformat(value, delim) {
	/// <summary>Reformats a number for display purposes using delimiter specified.</summary>
	/// <param name="value" type="String">The phone number.</summary>
	/// <param name="delim" type="String">The delimiter used. If argument is not passed, a hyphen is used.</summary>
	/// <returns type="String">Reformatted phone number.</returns>
	value = phoneStripFormatting(value);

	if (delim == null)
		delim = '-';

	return value.substr(0, 3) + delim + value.substr(3, 3) + delim + value.substr(6)
}

function phoneObfuscate(value, secret, delim) {
	/// <summary>Reformats a number for display purposes, obfuscating the first 6 digits.</summary>
	/// <param name="value" type="String">The phone number.</summary>
	/// <param name="secret" type="String">The obfuscation character or string to use in place of a digit in the number. If argument is not pased, a asterisk is used.</summary>
	/// <param name="delim" type="String">The delimiter used. If argument is not passed, a hyphen is used.</summary>
	/// <returns type="String">Reformatted, obfuscated phone number.</returns>
	value = phoneStripFormatting(value);

	if (secret == null)
		secret = '*';

	if (delim == null)
		delim = '-';

	secret = secret + secret + secret + delim;
	return secret + secret + value.substr(6);
}

function phoneValidate(value, intCountryCode) {
	/// <summary>Validates that phone numbers appear to adhere to the NANP format (North American Numbering Plan).</summary>
	/// <remarks>
	/// Expects area code to be present, and no country code (e.g. 555-555-0123 is valid, 1-555-555-0123 is not).
	/// Does not validate that numbers are legal (e.g. area code 123 is invalid, area codes cannot begin with 1).
	/// </remarks>
	/// <param name="value" type="String">Phone number.</summary>
	/// <returns type="Boolean">Whether or not number passed a basic format validation.</returns>
	value = phoneStripFormatting(value);
	if (value.length == 10)
		return numbersOnly.test(value);
	else if (value.length == 8 && intCountryCode == 36) //For Australia
		return numbersOnly.test(value);
	else
		return false;
}

//gets the cookie, if it exists
function getCookieValue(cookiename) {
	var ck = document.cookie;
	var cn = cookiename + "=";
	var pos = ck.indexOf(cn);

	if (pos != -1) {
		var start = pos + cn.length;
		var end = ck.indexOf(";", start);
		if (end == -1) end = ck.length;
		var cookieValue = ck.substring(start, end);
		return (cookieValue);
	}

	return (null);
}

window['getCookieValue'] = getCookieValue;

function setCookieValue(cname, cvalue, exdays) {
    var d = new Date();
    d.setTime(d.getTime() + (exdays * 24 * 60 * 60 * 1000));
    var expires = "expires=" + d.toUTCString();
    document.cookie = cname + "=" + cvalue + "; " + expires + "; path=/";
}

//Gets no of days between 2 dates provided in a string format('mm/dd/yyyy')
function dayDifference(first, second) {
	return Math.floor((Date.parse(first) - Date.parse(second)) / 86400000);
}

// Returns whether or not file extension given is acceptable.
// Client-side version of FileSystemSafety.IsAllowedFileExtension().
function isAllowedFileExtensionWhiteList(ext) {

	// Check if the URL contains the string "Admin"
	const allowed = window.location.pathname.toLowerCase().startsWith("/admin") ? 1 : 0;

	if (!window.acceptedFileExtensions) {
		window.acceptedFileExtensions =
		{
			'.123': 1, '.3g2': 1, '.3gp': 1, '.3gp2': 1, '.7z': allowed, '.aac': 1, '.abw': 1, '.abx': 1, '.accdb': 1, '.accde': 1,
			'.accdr': 1, '.accdt': 1, '.accfl': 1, '.ace': allowed, '.ade': 1, '.adp': 1, '.aeh': 1, '.ai': 1, '.aif': 1, '.aiff': 1,
			'.amr': 1, '.apng': 1, '.arc': allowed, '.art': 1, '.asc': 1, '.asf': 1, '.asx': 1, '.atom': 1, '.au': 1, '.avi': 1,
			'.awt': 1, '.azw': 1, '.azw1': 1, '.b64': 1, '.backup': 1, '.bin': 1, '.bmp': 1, '.bz': 1, '.bz2': 1, '.bzip': 1,
			'.bzip2': allowed, '.cab': allowed, '.cad': 1, '.caf': 1, '.cal': 1, '.cgm': 1, '.crtx': 1, '.css': 1, '.csv': 1, '.cwk': 1, '.dap': 1, '.db': 1,
			'.dbf': 1, '.dcr': 1, '.dds': 1, '.dex': 1, '.dib': 1, '.dif': 1, '.diz': 1, '.dmg': 1, '.doc': 1, '.docm': 1,
			'.docx': 1, '.dot': 1, '.dotm': 1, '.dotx': 1, '.drw': 1, '.ds_store': 1, '.dv': 1, '.dvr-ms': 1, '.dwf': 1, '.dwg': 1,
			'.dxf': 1, '.emf': 1, '.eml': 1, '.emlx': 1, '.emz': 1, '.eot': 1, '.eps': 1, '.fbp': 1, '.fdp': 1, '.fhtml': 1,
			'.file': 1, '.flac': 1, '.flc': 1, '.fli': 1, '.flp': 1, '.flv': 1, '.fm': 1, '.fm2': 1, '.fm3': 1, '.fm5': 1,
			'.fmp': 1, '.fnt': 1, '.fon': 1, '.fp': 1, '.fp3': 1, '.fp5': 1, '.fp7': 1, '.fphtml': 1, '.fpt': 1, '.fpweb': 1,
			'.fpx': 1, '.gg': 1, '.gif': 1, '.gnumeric': 1, '.gra': 1, '.gsm': 1, '.gz': allowed, '.hdmov': 1, '.hdp': 1, '.hqx': 1,
			'.htm': 1, '.html': 1, '.ical': 1, '.icns': 1, '.ico': 1, '.ics': 1, '.ifo': 1, '.indd': 1, '.isf': 1, '.isu': 1,
			'.ivs': 1, '.jbf': 1, '.jfif': 1, '.jpe': 1, '.jpeg': 1, '.jpf': 1, '.jpg': 1, '.jxr': 1, '.key': 1, '.kml': 1, '.kmz': 1,
			'.knt': 1, '.kth': 1, '.lhz': 1, '.lit': 1, '.log': 1, '.lrc': 1, '.lrf': 1, '.lrx': 1, '.lst': 1, '.lyr': 1,
			'.m4a': 1, '.m4b': 1, '.m4p': 1, '.m4r': 1, '.mdb': 1, '.mde': 1, '.mht': 1, '.mhtml': 1, '.mid': 1, '.midi': 1,
			'.mim': 1, '.mix': 1, '.mng': 1, '.mobi': 1, '.mod': 1, '.moi': 1, '.mov': 1, '.movie': 1, '.mp3': 1, '.mp4': 1,
			'.mpa': 1, '.mpc': 1, '.mpe': 1, '.mpeg': 1, '.mpg': 1, '.mpv2': 1, '.msg': 1, '.mswmm': 1, '.mxd': 1, '.numbers': 1,
			'.odb': 1, '.odf': 1, '.odg': 1, '.ods': 1, '.odx': 1, '.ofm': 1, '.oft': 1, '.ogg': 1, '.ogm': 1, '.ogv': 1,
			'.one': 1, '.onepkg': 1, '.opx': 1, '.osis': 1, '.ost': 1, '.otf': 1, '.oth': 1, '.p65': 1, '.p7b': 1, '.pages': 1,
			'.pbm': 1, '.pcast': 1, '.pcf': 1, '.pcm': 1, '.pct': 1, '.pcx': 1, '.pdc': 1, '.pdd': 1, '.pdf': 1, '.pdp': 1,
			'.pfx': 1, '.pgf': 1, '.pic': 1, '.pict': 1, '.pkg': 1, '.pmd': 1, '.pmf': 1, '.png': 1, '.pot': 1, '.pothtml': 1,
			'.potm': 1, '.potx': 1, '.ppam': 1, '.pps': 1, '.ppsm': 1, '.ppsx': 1, '.ppt': 1, '.ppthtml': 1, '.pptm': 1, '.pptx': 1,
			'.prc': 1, '.ps': 1, '.psd': 1, '.psp': 1, '.pspimage': 1, '.pst': 1, '.pub': 1, '.pubhtml': 1, '.pubmhtml': 1, '.qbb': 1,
			'.qcp': 1, '.qt': 1, '.qxd': 1, '.qxp': 1, '.ra': 1, '.ram': 1, '.ramm': 1, '.rar': allowed, '.raw': 1, '.rax': 1,
			'.rm': 1, '.rmh': 1, '.rmi': 1, '.rmm': 1, '.rmvb': 1, '.rmx': 1, '.rp': 1, '.rss': 1, '.rt': 1, '.rtf': 1,
			'.rts': 1, '.rv': 1, '.sbx': 1, '.sdf': 1, '.sea': 1, '.shs': 1, '.sit': 1, '.sitx': 1, '.smil': 1, '.snapfireshow': 1,
			'.snp': 1, '.stc': 1, '.svg': 1, '.svgz': 1, '.swf': 1, '.sxc': 1, '.sxi': 1, '.tab': 1, '.tar': allowed, '.tex': 1,
			'.tga': 1, '.thmx': 1, '.tif': 1, '.tiff': 1, '.tpz': 1, '.tr': 1, '.trm': 1, '.tsv': 1, '.ttf': 1, '.txt': 1,
			'.uue': 1, '.vcf': 1, '.vob': 1, '.vrml': 1, '.vsc': 1, '.vsd': 1, '.wab': 1, '.wav': 1, '.wax': 1, '.wbk': 1,
			'.wbmp': 1, '.wdp': 1, '.webarchive': 1, '.webloc': 1, '.wk1': 1, '.wk3': 1, '.wm': 1, '.wma': 1, '.wmf': 1, '.wmmp': 1,
			'.wmv': 1, '.wmx': 1, '.wpd': 1, '.wps': 1, '.wri': 1, '.wvx': 1, '.xbm': 1, '.xcf': 1, '.xg0': 1, '.xg1': 1,
			'.xg2': 1, '.xg3': 1, '.xg4': 1, '.xht': 1, '.xhtm': 1, '.xhtml': 1, '.xif': 1, '.xlam': 1, '.xlb': 1, '.xlc': 1,
			'.xld': 1, '.xlk': 1, '.xlm': 1, '.xls': 1, '.xlsb': 1, '.xlshtml': 1, '.xlsm': 1, '.xlsmhtml': 1, '.xlsx': 1, '.xlt': 1,
			'.xlthtml': 1, '.xltm': 1, '.xltx': 1, '.xlv': 1, '.xlw': 1, '.xml': 1, '.xpi': 1, '.xps': 1, '.xsf': 1, '.xsn': 1,
			'.xtp': 1, '.zabw': 1, '.zip': allowed, '.zipx': allowed
		};
	}

	return (window.acceptedFileExtensions[(ext + '').toLowerCase()] === 1);
}

function handleServiceMethodError(error) {
	throw new Error('A server-side error occurred during the execution of an AJAX request (' +
		error.get_exceptionType() +
		')!\r\n\r\nAdditional Information:\r\n' +
		error.get_message() +
		'\r\n\r\nStack Trace:\r\n' +
		error.get_stackTrace().trim());
}

function toggleDisplay(elem) {
	elem.style.display =
		(elem.style.display == 'none' ? '' : 'none');
}

// BEGIN Old Character Counter
var APENDED_TEXT_LEN = typeof intCountryCode !== 'undefined' && intCountryCode === 840 && typeof FeatureToggles !== 'undefined' ? 9 : 0;
var MAX_MESSAGE_SIZE_SMS = 160 - APENDED_TEXT_LEN;
var trackLengths = new Array();

function registerCharCounter(input, outputElemID, maxMessageSize) {
	var outputElem = document.getElementById(outputElemID);
	var id = input.id;
	var tracker;

	tracker = new Object();
	tracker.oldLength = -1;
	tracker.outputElem = outputElem;
	tracker.input = input;
	trackLengths[id] = tracker;

	setInterval(
		'var tracker = trackLengths[\'' + id + '\'];\r\n' +
		'var newLength = (tracker.input.value + \'\').length;\r\n' +
		'if (newLength != tracker.oldLength) {\r\n' +
		'	if (newLength > ' + maxMessageSize + ')' +
		'		tracker.outputElem.innerHTML = (newLength - ' + maxMessageSize + ') + " Characters Over";\r\n' +
		'	else\r\n' +
		'		tracker.outputElem.innerHTML = (' + maxMessageSize + ' - newLength) + " Characters Remaining";\r\n' +
		'	\r\n' +
		'	tracker.oldLength = newLength;\r\n' +
		'}\r\n', 100);
}
// END Old Character Counter

function elemInsideOrEq(elemFirst, elemSecond) {
	/// <summary>Gets whether or not second element specified is inside or is equal to first element.</summary>
	/// <param name="elemFirst" type="domElement">First element.</param>
	/// <param name="elemSecond" type="domElement">Second element.</param>
	if (elemFirst == elemSecond)
		return true;

	if (elemFirst && elemSecond) {
		elemSecond = elemSecond.parentNode;

		while (elemSecond != null) {
			if (elemSecond == elemFirst)
				return true;

			elemSecond = elemSecond.parentNode;
		}
	}

	return false;
}

// Do not use in new code. Use jQuery with event.preventDefault().
function hookAnchorClick(anchor, fn) {
	/// <summary>Hooks an anchor onclick and prevents default HREF behavior.</summary>
	/// <param name="anchor" type="domElement">The DOM element to attach the event handler for.</param>
	/// <param name="fn" type="Function">Function to be called when anchor is clicked. Should have sender and event argument, in that order.</param>
	var clickHandler = function (event) {
		var sender;

		if (window.event) {
			event = window.event;
			sender = event.srcElement;
		}
		else
			sender = event.target;

		fn(sender, event);

		if (event.stopPropagation)
			event.stopPropagation();
		else
			event.cancelBubble = true;

		return false;
	}

	// HACK: MSIE/Opera onclick fails to operate correctly.
	if (isie || isie8 || isie7 || isie6 || (document.documentMode < 8) || window.opera)
		anchor.onmousedown = clickHandler;
	else
		anchor.onclick = clickHandler;
}

// For ClientCharacterCounter control.
var registeredCounters = null;
function registerCounter(assocElem, counterElem, fmtStandard, fmtOver, maxChars) {
	if (assocElem == null || counterElem == null || maxChars == null || fmtStandard == null || fmtOver == null)
		return;

	if (registeredCounters == null) {
		registeredCounters = [];

		setInterval(function () {
			for (var i = 0, l = registeredCounters.length; i < l; i++) {
				if (registeredCounters[i].counter && registeredCounters[i].input) {
					var len;
					if (typeof registeredCounters[i].input.value === 'undefined')
						len = registeredCounters[i].input.val().length;
					else
						len = registeredCounters[i].input.value.length;
					var prevLen = registeredCounters[i].prevLen;

					if (len == prevLen)
						continue;
					else {
						registeredCounters[i].prevLen = len;

						var max = registeredCounters[i].maxChars;
						var numToDisplay = max - len;

						var fmtToUse = (numToDisplay >= 0
							? registeredCounters[i].fmtStandard
							: registeredCounters[i].fmtOver);

						if (numToDisplay >= 0)
							registeredCounters[i].counter.className = 'counterUnderLimit';
						else
							registeredCounters[i].counter.className = 'counterOverLimit';

						numToDisplay = Math.abs(numToDisplay);

						registeredCounters[i].counter.innerHTML = fmtToUse
							.replace(/\{0\}/g, numToDisplay)
							.replace(/\{1\}/g, (numToDisplay == 1 ? '' : 's'));
					}
				}
			}
		}, 100);
	}

	registeredCounters.push({
		"fmtStandard": fmtStandard,
		"fmtOver": fmtOver,
		"counter": counterElem,
		"input": assocElem,
		"prevLen": -1,
		"maxChars": maxChars
	});
}

// Detects if nextElementSibling/previousElementSibling are present.
function detectElementSiblingSupport() {
	var p = document.createElement('div');
	var c1 = document.createElement('div');
	var c2 = document.createElement('div');

	p.appendChild(c1);
	p.appendChild(c2);

	return (c1.nextElementSibling != null);
}

// Gets next sibling element. Version for browsers missing nextElementSibling DOM property.
function nextElementSibling(elem) {
	elem = elem.nextSibling;

	while (elem != null) {
		if (elem.nodeType == 1)
			return elem;

		elem = elem.nextSibling;
	}

	return null;
}

// Gets previous sibling element. Version for browsers missing previousElementSibling DOM property.
function previousElementSibling(elem) {
	elem = elem.previousSibling;

	while (elem != null) {
		if (elem.nodeType == 1)
			return elem;

		elem = elem.previousSibling;
	}

	return null;
}

// Gets first child element. Version for browsers missing firstElementChild DOM property.
function firstElementChild(elem) {
	elem = elem.firstChild;

	while (elem != null) {
		if (elem.nodeType == 1)
			return elem;

		elem = elem.nextSibling;
	}

	return null;
}

if (detectElementSiblingSupport()) {
	// If browsers support built-in DOM properties, just use those.
	nextElementSibling = function (elem) { return elem.nextElementSibling; }
	previousElementSibling = function (elem) { return elem.previousElementSibling; }
	firstElementChild = function (elem) { return elem.firstElementChild; }
}

function haltEvent(event, windowEvent) {
	if (windowEvent & !window.opera)
		windowEvent.cancelBubble();
	else {
		if (event.stopPropagation)
			event.stopPropagation();

		if (event.preventDefault)
			event.preventDefault();
	}
}

// Gets absolute left offset.
function absoluteOffsetLeft(oElem) {
	var oLeft = 0;

	while (oElem != null) {
		oLeft += oElem.offsetLeft;
		oElem = oElem.offsetParent;
	}

	return oLeft;
}

function imposeMaxLength(e, Object, MaxLen) {
	//Truncate to maximum length
	if (Object.value.length > MaxLen)
		Object.value = Object.value.substring(0, MaxLen);

	//ALLOW:        BACKSPACE           DELETE                  LEFT ARROW              UP ARROW            RIGHT ARROW             DOWN ARROW
	if ((e.keyCode == '8') || (e.keyCode == '46') || (e.keyCode == '37') || (e.keyCode == '38') || (e.keyCode == '39') || (e.keyCode == '40'))
		return true;

	return (Object.value.length <= MaxLen);
}

function checkURL(val, showAlert) {
	val = TrimString(val);
	if (val != '' && (val.substr(0, 7) != 'http://' && val.substr(0, 8) != 'https://' && val.substr(0, 6) != 'ftp://' && val.substr(0, 1) != '/')) {
		if (arguments.length == 1 || showAlert)
			alert('All URLs must begin with http:// or https:// or ftp://.\nAll internal links must start with a /.');

		return false;
	}
	return true;
}

function toSQLSafe(formvar) {
	var s = String(formvar);
	s = (s.replace(/'/gi, "&#39")).replace(/"/gi, "&#34");
	return s;
}

function cancelFE(needsConfirm) {
	if (needsConfirm) {
		if (!confirm("You will discard any unsaved changes. Do you want to proceed?"))
			return;
	}
	window.parent.closeModalDialog('editItemBehavior');
	return false;
}

function getThumbNailPath(path, width, height) {
	var lastSlashPos = path.lastIndexOf('/');
	var filename = path.substr(lastSlashPos + 1);
	var directory = path.substr(0, lastSlashPos);
	var lastDotPos = filename.lastIndexOf('.');
	var filenameWithOutExtension = filename.substr(0, lastDotPos)
	var fileExtension = filename.substr(lastDotPos)
	return directory + "/ThumbNails/" + filenameWithOutExtension + "_" + width + "x" + height + "_thumb" + fileExtension;
}

function measureHtml(html, fontStyle) {
	var tmp = document.createElement('span');
	tmp.style.display = 'inline';
	tmp.style.position = 'absolute';
	tmp.style.top = '-1000px';
	tmp.style.left = '-1000px';
	tmp.style.font = fontStyle;
	tmp.innerHTML = html;

	document.body.appendChild(tmp);
	var retVal = { width: tmp.clientWidth, height: tmp.clientHeight };
	document.body.removeChild(tmp);

	return retVal;
}

function fileValidateError(errorType) {
	if (errorType == 1)
		alert("A file you tried to upload was not of a permitted type. Only *.jpeg, *.jpg, *.gif, *.png, *.bmp are allowed");
	else
		alert("You have an invalid character in your filename. These are the accepted characters: a-z 0-9 ~ !( ) - + = [ ] { } , . $");
}

function getNumericThousandsSeparator() {
	var dummy = new Number(11111111);
	var thousandSep = dummy.toLocaleString().replace(/1/g, '');
	return (thousandSep.length > 0 ? thousandSep.charAt(0) : ',');
}

function getNumericDecimalSeparator() {
	var dummy = new Number(1.1);
	var thousandSep = dummy.toLocaleString().replace(/1/g, '');
	return (thousandSep.length > 0 ? thousandSep.charAt(0) : '.');
}

function ChangeDateFormat(elem) {
	if (elem.value != "") {
		var Error = false;
		var dateSplit = (elem.value).split("/");

		if (dateSplit.length == 3) {
			if (dateSplit[0] <= 12 && dateSplit[1] <= 31) {
				var d = new Date(elem.value);
				var currentDate = new Date();

				var selectedDate = d.getDate();
				var selectedMonth = d.getMonth();
				var SelectedYear = d.getFullYear();

				if (isNaN(selectedDate))
					Error = true;
				else if (isNaN(selectedDate))
					Error = true;
				else if (isNaN(selectedDate))
					Error = true;
				if (Error == false) {
					if (selectedDate.toString().length == 1)
						selectedDate = '0' + selectedDate;

					selectedMonth++;

					if (selectedMonth.toString().length == 1)
						selectedMonth = '0' + selectedMonth;

					currentYear = currentDate.getFullYear();

					if (SelectedYear.toString().substr(0, 2) < 20)
						SelectedYear = currentYear;
					elem.value = selectedMonth + '/' + selectedDate + '/' + SelectedYear;
				}
			}
			else
				Error = true;
		}
		else
			Error = true;

		if (Error == true) {
			alert('Date should be in the format MM/DD/YYY');
			elem.value = "";
		}
	}
}

function ChangeDateFormatNew(elem) {
	return isDateNew(elem.value);
}

function isSilverlightInstalled() {
	var isSilverlightInstalled = false;
	try {
		//check on IE
		try {
			var slControl = new ActiveXObject('AgControl.AgControl');
			isSilverlightInstalled = true;
			isSilverlightInstalled = true;
		}
		catch (e) {
			//either not installed or not IE. Check Firefox.
			if (navigator.plugins["Silverlight Plug-In"]) {
				isSilverlightInstalled = true;
			}
		}
	}
	catch (e) {
		//we don't want to leak exceptions. However, you may want
		//to add exception tracking code here.
	}
	return isSilverlightInstalled;
}

function setModalClass(cssClass, behaviorID) {
	var modalContainer = document.getElementById(behaviorID._PopupControlID);

	if (modalContainer)
		modalContainer.className = 'modalContainer modalContainerCP ' + cssClass;
}

function setModalClassFE(cssClass, behaviorID) {
	var modalContainer = document.getElementById(behaviorID._PopupControlID);

	if (modalContainer)
		modalContainer.className = 'modalContainer ' + cssClass;
}

//------------------- Copy Link ---------------------//
function showCopyLinkWindow(relatedLink, absoluteAddress, behavior, ifrID) {
	if (behavior) {
		setModalClass('modalCopyLink', behavior);
		var iFrame = document.getElementById(ifrID);
		iFrame.src = '/common/admin/CopyLink.aspx?relatedLink=' + encodeURIComponent(relatedLink) + '&absoluteAddress=' + encodeURIComponent(absoluteAddress);
		iFrame.style.display = 'block';
		behavior.show();
	} else {
		openCpModal({
			title: 'Copy Link',
			className: 'modalCopyLink',
			isFrontEnd: false,
			useIframe: true,
			url: '/common/admin/CopyLink.aspx?relatedLink=' + encodeURIComponent(relatedLink) + '&absoluteAddress=' + encodeURIComponent(absoluteAddress)
		});
	}
}

function showCopyLinkWindowUsingAddress(absoluteAddress) {
	var relatedUrl = absoluteAddress.replace(window.location.host, '');
	relatedUrl = relatedUrl.replace(window.location.protocol + '//', '');

	var behavior = $find('copyLinkBehavior');
	if (behavior)
		showCopyLinkWindow(relatedUrl, absoluteAddress, behavior, 'copyLinkDialog');
	else
		showCopyLinkWindow(relatedUrl, absoluteAddress, $find('editItemBehavior'), 'liveEditDialog');
}

function showCopyLinkWindowUsingRelatedLink(relatedUrl) {
	var absoluteUrl = window.location.protocol + '//' + window.location.host + relatedUrl;

	var behavior = $find('copyLinkBehavior');
	if (behavior)
		showCopyLinkWindow(relatedUrl, absoluteUrl, behavior, 'copyLinkDialog');
	else
		showCopyLinkWindow(relatedUrl, absoluteUrl, $find('editItemBehavior'), 'liveEditDialog');
}

function closeCopyModal() {
	var behavior = $find('copyLinkBehavior');
	if (behavior)
		behavior.hide();
	else {
		behavior = $find('editItemBehavior');
		if (behavior)
			behavior.hide();
		else
			closeCpModal();
	}
}
//------------------- Copy Link End---------------------//

//Parses the query string for the specified paramter and returns the value
//Returns "" if the parameter is not found
function GetQueryStringParameter(name) {
	name = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
	var regexS = "[\\?&]" + name + "=([^&#]*)";
	var regex = new RegExp(regexS);
	var results = regex.exec(window.location.href);
	if (results == null)
		return "";
	else
		return results[1];
}

function setModalTitle(title, behaviorID) {
	if (behaviorID) {
		$('#' + behaviorID._PopupControlID).find('.modalTitle').text(title);
	}
	else {
		//if behaviorID is not provided, do the same as setModalTitle on FrontEnd.js
		var titleElem = document.getElementById('ctl00_LiveEditModalTitle');

		if (titleElem == null)
			titleElem = document.getElementById('ctl00_ct100_LiveEditModalTitle');

		if (titleElem == null)
			titleElem = document.getElementById('ctl00_ctl00_NotifyMeModalTitle');

		if (titleElem)
			titleElem.innerHTML = title;
	}
}

//------------------- Copy Link End---------------------//

function isNull(object, replacement) {
	return object == null ? replacement : object;
}

// Validates domain of email with service (ensures MX record present, DNS name valid, acceptable format).
// The callbacks object should contain the following methods:
// * badFormat - Email format invalid. Will also be thrown for null/empty email addresses.
// * mxMissing - MX record missing for the domain.
// * dnsError - A DNS error occurred (domain invalid, invalid characters in domain, etc).
// * success - The email appears to be valid.
function validateEmailDomain(emailAddress, callbacks) {
	$.ajax({
		type: "POST",
		url: "/Services/ValidateEmail.ashx",
		data: { email: emailAddress },
		success: function (data, textStatus, jqXHR) {
			switch (data.d) {
				case 1:
					callbacks.badFormat(emailAddress);
					break;
				case 2:
					callbacks.mxMissing(emailAddress);
					break;
				case 3:
					callbacks.dnsError(emailAddress);
					break;
				case 0:
				default:
					callbacks.success(emailAddress);
					break;
			}
		}
	});
}

//Called when opening action menu.
function openPopUp(id, event) {
	var $popUp = $(document.getElementById(id));

	if ($popUp.hasClass('popped')) {
		//do not do anything if popup is already open.
		return false;
	}

	//Hide all popups already open if nesessary.
	$('.popUp').fadeOut(200).removeClass('popped');
	$('.popUpParent').removeClass('popped');

	//Open popup
	$popUp.fadeIn(200);
	$popUp.addClass('popped');
	$popUp.parents('.popUpParent').addClass('popped');

	//prevent triggering any other event.
	if (window.event)
		window.event.cancelBubble = true;
	else
		event.stopPropagation();
}

function SignIn() {
	window.location.href = '/MyAccount?from=url&url=' + window.location.pathname + window.location.search;
}

function SignOut() {
	window.location.href = '/publiclogin.aspx?out=true&txtRedirect=' + window.location.pathname.substring(1) + window.location.search;
}

function removeIEParagraphs(editor) {
	editor.attachEventHandler("onkeyup", function (e) {
		if (e.keyCode == 8 || e.keyCode == 46 || e.keyCode == 13) {
			var result = '';
			var re = new RegExp("<[P]>&nbsp;<\\/[P]>", "gi");

			var m = re.exec(editor.get_html(true));

			if (m == null) {
				return;
			} else {
				result = editor.get_html(true).replace(re, '<BR>');
				editor.set_html(result);
			}
		}
	});
}

//To toggle the slide show and image manger menu items based on if the img is a from a slide show or not.
function setMenuItems(editor) {
	var contextMenu = editor.getContextMenuByTagName("IMG");
	if (!contextMenu) return;

	//Attach to the context menu show event, obtain editor selection and determine which tools to show
	contextMenu.add_show(function () {
		//Must be an image, the context menu would not fire
		var sel = editor.getSelectedElement();

		//An array of menu item objects
		var menuItems = contextMenu.get_items();
		var menuItem = null;

		//reset to show all and then disable the menu item based on currently selection
		for (i = 0; i < menuItems.length; i++) {
			menuItems[i].get_element().style.display = "";
			menuItems[i].get_element().parentNode.style.display = "";
		} //end for

		if (sel.getAttribute("slideshowimage")) {
			menuItem = menuItems[menuItems.length - 3];
			menuItem.get_element().style.display = "none";
			menuItem.get_element().parentNode.style.display = "none";

			menuItem = menuItems[menuItems.length - 2];
			menuItem.get_element().style.display = "none";
			menuItem.get_element().parentNode.style.display = "none";
		} else {
			menuItem = menuItems[menuItems.length - 1];
			menuItem.get_element().style.display = "none";
			menuItem.get_element().parentNode.style.display = "none";
		}
	});
}

function integersOnly(e, input) {
	var charCode = (e.which) ? e.which : e.keyCode;

	var keyChar = String.fromCharCode(charCode);
	var currentValue = $(input).val();
	var allowNegativeSign = !$(input).hasClass("positive") && currentValue.indexOf('-') == -1 && (currentValue.length == 0 || currentValue == window.getSelection().toString());
	if (keyChar == '-' && allowNegativeSign)
		return true;

	if (charCode > 31 && (charCode < 48 || charCode > 57))
		return false;

	return true;
}

function decimalOnly(e, input, skipSelected) {
	skipSelected = (skipSelected == undefined || skipSelected == null) ? false : skipSelected;

	var charCode = (e.which) ? e.which : e.keyCode;

	var keyChar = String.fromCharCode(charCode);
	var currentValue = $(input).val();
	var hasDecimalPoint = currentValue.indexOf('.') != -1;
	if ((keyChar == '.' && !hasDecimalPoint))
		return true;

	if (!skipSelected && isTextSelected(input))
		return true;

	return integersOnly(e, input);
}

function isTextSelected(input) {
	if (typeof input.selectionStart == "number") {
		return input.selectionStart == 0 && input.selectionEnd == input.value.length;
	} else if (typeof document.selection != "undefined") {
		input.focus();
		return document.selection.createRange().text == input.value;
	}
}

function validFileName(e) {
	var charCode = (e.which) ? e.which : e.keyCode;
	var keyChar = String.fromCharCode(charCode);
	var invalidCharacters = "\\/:*?\"<>|',";

	return invalidCharacters.indexOf(keyChar) < 0;
}

function isValidURL(value) {
	return /^(https?:\/\/|ftp:\/\/|\/).+/i.test(value);
}
function isValidAbsoluteURL(value) {
	try {
		var parsedUrl = new URL(value, window.location.href);
	} catch (e) {
		return false;
	}
	return /^(https?|ftp):\/\//i.test(value) && /^(https?|ftp):$/.test(parsedUrl.protocol) && parsedUrl.host != '';
}

function validateLink(link) {
	if (link != '') {
		if ((link.indexOf('http://') == -1) &&
		 (link.indexOf('https://') == -1) &&
		 (link.substr(0, 1) != '/')) {
			return false;
		}
	}
	return true;
}

// Check if string is a whole number(digits only).
function isWholeNumber(val) {
	return String(val).search(/^(([1-9]{1}\d{0,2},(\d{3},)*\d{3})|([1-9]{1}\d{0,}))$/) != -1;
}

// Checks that an input string is a decimal number
function isDecimalOrZero(val) {
	isDecimal_re = /^\d{1,5}(\.\d{1,2})?$/;
	return String(val).search(isDecimal_re) != -1;
}

//Allows login to CPC
function RebuildCSS(menu) {
	if (confirm('This will take some time\nand the site may look funny during the process.\nAre you sure?')) {
		var frmARC = document.forms.frmAdminRebuldCSS;
		frmARC.ysnNeedInclude.value = '1';
		frmARC.intMenu.value = menu;
		frmARC.submit();
	}
}

//Allows login to CPC
function adminGetHelp(helpUrl, newWindow) {
	document.cpconnect.action = document.cpconnect.action.split('?')[0] + '?from=url&url=' + encodeURIComponent(helpUrl);
	cpconnecttLogin();
}

function connectCPC() {
	document.cpconnect.action = document.cpconnect.action.split('?')[0];
	var ifr = document.createElement('iframe');
	ifr.name = "submitCPCFrame";
	ifr.id = "submitCPCFrame";
	ifr.style.display = 'none';
	var frm = $(document.cpconnect).clone().get(0);
	frm.name = "iframeCPCConnect"
	frm.target = 'submitCPCFrame';
	ifr.appendChild(frm);
	document.body.appendChild(ifr);
	frm.submit();
}

function showFeatureNotAvailablePopup(hypothesisID) {
	var data = {
		id: hypothesisID
	};

	$('.emailPage').removeClass('emailPage');

	$.ajax({
		url: '/FeatureNotAvailable/Popup',
		type: 'GET',
		data: {
			hypothesisID: hypothesisID
		},
		cache: false,
		success: function (response) {
			if (response.action == 'modal') {
				openCpModal({
					title: 'Feature Not Yet Available',
					className: 'featureNotAvailableModal moduleContentNew modalContainer',
					htmlContent: response.html,
					autoAdjust: true,
					isFrontEnd: true,
					elementID: 'featureNotAvailableModal'
				});
			} else {
				window.location = '/FeatureNotAvailable?hypothesisID='
                    + hypothesisID
                    + '&referrerUrl=' + window.location;
			}
		},
		beforeSend: function () {
			ajaxPostBackStart('Loading');
		},
		complete: function () {
			ajaxPostBackEnd();
		}
	});
}

function getDateFormat() {
	if (!gDateFormat) {
		$.ajax({
			url: "/GetDateFormat",
			type: "GET",
			success: function (response) {
				if (response.DateFormat)
					gDateFormat = response.DateFormat;
			},
			error: function (xhr, textStatus, exception) {
				gDateFormat = "MM/dd/yyyy";
			},
			async: false
		});
	}

	return gDateFormat == null ? "MM/dd/yyyy" : gDateFormat;	

}

function getdDateFormat() {
	var dateFormat;

	$.ajax({
		url: '/Utility/GetdDateFormat',
		async: false,
		type: 'GET',
		dataType: 'json',
		success: function (data) {
			dateFormat = data.dateFormat;
		},
		error: function (jqXHR, textStatus, errorThrown) {
			dateFormat = "mm/dd/yyyy";
		}
	});

	return dateFormat;
}

function getDateFromInput(datePickerID, keepDayBeforeMonth) {
	var datePicker = $(datePickerID);
	var date;
	var dateFormat = keepDayBeforeMonth ? getDateFormat() : "mm/dd/yyyy";
	//See if datepicker if so becuase of UK format we cannot just grab the value
	if (!isNaN(datePicker.data('tDatePicker')) && datePicker.data('tDatePicker') && datePicker.data('tDatePicker').value()) {
	    var dateValTmp = new dateValidator();
		date = dateValTmp.format(dateFormat,new Date(datePicker.data('tDatePicker').value()));
	} else {
		if (datePicker.val()) {
			var dayBeforeMonthOn = getDateFormat().toLocaleLowerCase() == "dd/mm/yyyy" ? true : false;
			//Check if valid date before trying to format            			
			var dateString = checkDateFormatReturnUSString(datePicker.val(), dayBeforeMonthOn);
			var dateVal = new dateValidator();
			if (dateVal.dateValidateNew(dateString, false)) {
				date = dateVal.format(dateFormat,new Date(checkDateFormatReturnUSString(dateString, false)));
			} else {
				//Send the date string formatted eventhough invalid date. Allow date validation to return approiate error
				if (keepDayBeforeMonth) {
					date = dateString;
				} else {
					date = checkDateFormatReturnUSString(dateString, true);
				}
			}
		}
	}
	return date;
}

function formatUkDateReturnUSString(dateStr, isForce) {
	var retVal;
	if (isForce || getDateFormat().toLocaleLowerCase() == "dd/mm/yyyy") {
		dateStr = dateStr.replace(/\-/g, '/');
		dateStr = dateStr.split("/");
		retVal = dateStr[1] + "/" + dateStr[0] + "/" + dateStr[2];
	} else {
		retVal = dateStr;
	}
	return retVal;
}

function formatRFC3339DateReturnString(dateStr, keepDayBeforeMonth) {
	var retVal;
	var dateStrMod = dateStr.replace(/\-/g, '/');
	dateStrMod = dateStr.split("/");
	if (dateStrMod.length == 3 && dateStrMod[0].length == 4) {
		if (keepDayBeforeMonth && getDateFormat().toLowerCase() == "dd/mm/yyyy") {
			retVal = dateStrMod[2] + "/" + dateStrMod[1] + "/" + dateStrMod[0];
        } else {
			retVal = dateStrMod[1] + "/" + dateStrMod[2] + "/" + dateStrMod[0];
		}
	} else {
		retVal = dateStr;
	}
	return retVal;
}

function checkDateFormatReturnUSString(dateStr, keepDayBeforeMonth) {
	var retVal;
	if (dateStr) {
		var splitDate = dateStr;
		dateStr = dateStr.replace(/\-/g, '/');
		splitDate = dateStr.split("/");
		if (splitDate.length == 3) {
			if (splitDate[0].length == 4) {
				retVal = formatRFC3339DateReturnString(dateStr, keepDayBeforeMonth);
			} else if (keepDayBeforeMonth) {
				retVal = dateStr;
			} else {
				var dateVal = new dateValidator();
				retVal = formatUkDateReturnUSString(dateVal.cleanDate(dateStr));
			}
		} else {
			retVal = dateStr;
		}
	}
	return retVal;
}

//IE no longer displays as IE by design code to determin if it is IE11
function isIE11() {
	var trident = !!navigator.userAgent.match(/Trident\/7.0/);
	var net = !!navigator.userAgent.match(/.NET4.0E/);
	return trident && net;
}

function isValidUrl(url) {
	return url.match(/^(ht|f)tps?:\/\/[a-z0-9-\.]+\.[a-z]{2,4}\/?([^\s<>\#%"\,\{\}\\|\\\^\[\]`]+)?$/);
}

function isValidAbsoluteUrl(value) {
	return isValidAbsoluteURL(value);
}

//Replaces the given tag with the tag to replace inside the element for all occurences of the tag within the element.
//$element - The element to perform the replace to
//tagToReplace - Tag to be replaced replace provided as string, ex: "i", "b", "strong" etc
//tagToReplaceWith - Tag to replace with, provided as string, ex: "i", "b", "strong" etc
function replaceTag($element, tagToReplace, tagToReplaceWith) {
    var $tagElemsToReplace = $element.find(tagToReplace);
    $tagElemsToReplace.each(function () {
        var $tagElem = $(this);
        var tagElemeHtml = $tagElem.html();
        $tagElem.replaceWith('<' + tagToReplaceWith + '>' + tagElemeHtml + '</' + tagToReplaceWith + '>');        
    });
}

window.addEventListener("load", function () {
	if (MutationObserver != null) {
		if ($('.contentContainerOld').length > 0 && ($('.spellchecker').length > 0 || $('[id^="DialogOpener"]').length > 0)) {
			var observer = new MutationObserver(function moveTelerikModalOverlay(mutations) {
				mutations.forEach(function (mutation) {
					if (!mutation.addedNodes) return;

					for (var i = 0; i < mutation.addedNodes.length; i++) {
						// do things to your newly added nodes here
						var node = mutation.addedNodes[i];

						if (node.classList != null && node.classList.contains("TelerikModalOverlay")) {
							$('.contentContainerOld').after(node);
						}
					}
				});
			});

			observer.observe(document.getElementsByClassName('contentContainerOld')[0], {
				childList: true,
				subtree: true,
				attributes: false,
				characterData: false
			});
		}
	}
});

if (!Array.prototype.includes) {
	Array.prototype.includes = function (searchElement) {
		'use strict';
		if (this == null) {
			throw new TypeError('Array.prototype.includes called on null or undefined');
		}

		var O = Object(this);
		var len = parseInt(O.length, 10) || 0;
		if (len === 0) {
			return false;
		}
		var n = parseInt(arguments[1], 10) || 0;
		var k;
		if (n >= 0) {
			k = n;
		} else {
			k = len + n;
			if (k < 0) { k = 0; }
		}
		var currentElement;
		while (k < len) {
			currentElement = O[k];
			if (searchElement === currentElement ||
			   (searchElement !== searchElement && currentElement !== currentElement)) { // NaN !== NaN
				return true;
			}
			k++;
		}
		return false;
	};
}

function isMsie() {
	ua = navigator.userAgent;
	/* MSIE used to detect old browsers and Trident used to newer ones*/
	var is_ie = ua.indexOf("MSIE ") > -1 || ua.indexOf("Trident/") > -1;

	return is_ie;
}

function isMozilla() {
	return navigator.userAgent.toLowerCase().indexOf('firefox') > -1;
}

function browserVersion() {
	ua = navigator.userAgent;
	var browserVersion = null;
	var verOffset = null;
	if (isMsie()) {
		verOffset = ua.indexOf("rv:");
		browserVersion = ua.slice(verOffset + 3).slice(0, 4);
	}
	else if (isMozilla()) {
		verOffset = ua.indexOf("Firefox")
		browserVersion=  ua.substring(verOffset + 8)
	}
	return browserVersion;
}

function setDatepickerHighlightedDate(date, datePickerId) {
	var dateFormat = getdDateFormat();
	var momentDate = moment(date, dateFormat.toUpperCase()).format('MM/DD/YYYY');
	if (momentDate.toLowerCase() !== "invalid date") {
		$('#' + datePickerId).pickadate('picker').set('highlight', date, { format: dateFormat });
	}
}
;
/// <reference path="../../Common/Controls/jquery-ui-1.14.2/jquery-ui.js" />
if (!window.Core) {
	window.Core = new CoreNamespace();
}

function AJAX(url, type, json, callBack, showLoadingMessage, contentType, forceCallBack, suppressErrorMessage) {
	contentType = contentType || 'application/x-www-form-urlencoded; charset=UTF-8';
	showLoadingMessage = showLoadingMessage == null ? true : showLoadingMessage;
	forceCallBack = forceCallBack == null ? false : forceCallBack;
	if (typeof (suppressErrorMessage) === 'undefined') suppressErrorMessage = false;
	$.ajax({
		url: url,
		type: type,
		contentType: contentType,
		data: json,
		success: function (response) {
			if (response.ErrorMessage && !forceCallBack) {
				if (response.ErrorMessage.indexOf('You do not have permission to perform this action.') != -1 && !isUserLoggedIn()) { // 079B4F43 
					alert('You are not logged in. Please login to continue.');
					location.href = '/myaccount?from=url&url=' + window.location.pathname;
				} else {
					alert(response.ErrorMessage);
				}
			}
			else if (response.RedirectURL && !forceCallBack)
				window.location.href = response.RedirectURL;
			else if (callBack)
				callBack(response);
		},
		beforeSend: function () {
			if (showLoadingMessage) {
				ajaxPostBackStart('Loading');
			}
		},
		complete: function () {
			if (showLoadingMessage) {
				ajaxPostBackEnd();
			}
		},
		error: function (xhr, textStatus, exception) {
			if (!suppressErrorMessage) {
				alert('Error: ' + xhr.statusText + '\nStatus: ' + xhr.status);
			}
		}
	});
}

function isUserLoggedIn() {
	var result = true; // if there is an error lets not show the login message.
	$.ajaxSetup({ async: false, ignoreEditor: true });
	$.post('/Utility/IsUserLoggedIn', function (data) {
		result = data;
	});
	$.ajaxSetup({ async: true, ignoreEditor: false });
	return result;
}

///Method to pause a javascript execution for x seconds, works like sleep(). VB
function sleep(seconds) {
	var max_sec = new Date().getTime();
	while (new Date() < max_sec + seconds * 1000) { }
	return true;
}

if (!String.prototype.format) {
	String.prototype.format = function () {
		var args = arguments;
		return this.replace(/{(\d+)}/g, function (match, number) {
			return typeof args[number] != 'undefined'
				? args[number]
				: match
				;
		});
	};
}

if (!String.prototype.isNullOrWhiteSpace) {
	String.prototype.isNullOrWhiteSpace = function () {
		return this == null || this.trim() == '';
	};
}

function isValidMoney(value) {
	var isValid = false;
	if (value && value != '')
		isValid = /^\$?\s*\d+(,\d{3})*(\.\d{1,2})?$/.test(value);
	return isValid;
};

if (typeof String.prototype.startsWith != 'function') {
	String.prototype.startsWith = function (str) {
		return this.slice(0, str.length) == str;
	};
}

String.prototype.format = String.prototype.f = function () {
	var s = this,
		i = arguments.length;

	while (i--) {
		s = s.replace(new RegExp('\\{' + i + '\\}', 'gm'), arguments[i]);
	}
	return s;
};

//Creates a javascript namespace, if one already exists, just returns that namespace.
function createNamespace(namespaceString) {
	var parts = namespaceString.split('.'),
		parent = window,
		currentPart = '',
		nsLength = parts.length;

	for (var i = 0, length = nsLength; i < length; i++) {
		currentPart = parts[i];
		parent[currentPart] = parent[currentPart] || {};
		parent = parent[currentPart];
	}
	return parent;
}
$(function () {

	// Ensures a dependency exists. If it doesn't acquire it.
	function ensure(dependencyExists, dependencyUrl, onceItExistsCallback) {
		if (dependencyExists) {
			onceItExistsCallback();
			return;
		}
		$.getScript(dependencyUrl, onceItExistsCallback);
	}

	var ajaxControlToolkit = window.AjaxControlToolkit;
	// This is hack override in order to fix problems with Ajax Toolkit Drag and Drop in newer browsers. It substitutes in jQuery Draggable if it is available.
	// A more proper fix is ideal, but extremely expensive as it would require replacing over a hundred instances of invocations of the old modal.
	// There isn't a clean way to integrate in the new modal system. JE
	// https://system.netsuite.com/app/crm/support/issuedb/issue.nl?id=3809568&l=T
	if (ajaxControlToolkit && ajaxControlToolkit.ModalPopupBehavior) {
		ajaxControlToolkit.ModalPopupBehavior.prototype._attachPopup = function () {
			/// <summary>
			/// Attach the event handlers for the popup
			/// </summary>

			if (this._DropShadow && !this._dropShadowBehavior) {
				this._dropShadowBehavior = $create(ajaxControlToolkit.DropShadowBehavior, {}, null, null, this._popupElement);
			}

			var _this = this;
			// Drag and Drop Replacement. Acquire any dependencies on the fly.
			ensure($.ui, "/Common/Controls/jquery-ui-1.14.2/jquery-ui.min.js", function () {
				$("#" + _this._PopupControlID).draggable({ handle: '#' + _this._PopupDragHandleControlID });
				$("#" + _this._PopupDragHandleControlID).css('cursor', 'move');
			});

			$addHandler(window, 'resize', this._resizeHandler);
			$addHandler(window, 'scroll', this._scrollHandler);
			this._windowHandlersAttached = true;
		}
	}
});

String.prototype.replaceAll = function (search, replacement) {
	var target = this;
	return target.replace(new RegExp(search, 'g'), replacement);
};

function executeViewMapPostClickActions($element) {
	//Hide the view map element post click(to avoid clicking repeatedly)
	$element.hide('fast');

	//Collect analytics around map clicks
	var url = '/CollectGoogleMapClicks/Execute';
	var moduleId = $element.data('moduleid');
	$.ajax({
		url: url,
		type: 'POST',
		data: {
			url: window.location.href,
			moduleId: moduleId
		},
		success: function (json) {
			//send data to telemetry
			appInsights.trackEvent("googlemapclicked", json);
		},
		error: function (xhr, textStatus, exception) {
			alert("Error: " + xhr.statusText + "\nStatus: " + xhr.status);
		}
	});
}

function importModal(options) {
	if (window.ImportModalInstance === undefined) {
		var script = document.createElement("script");
		script.onload = function () {
			window.ImportModalInstance.init(options);
		};
		script.src = "/Assets/Scripts/ImportModal.js";
		document.head.appendChild(script);
	}
	else
		window.ImportModalInstance.init(options);
}

function exportModal(options) {
	if (window.ExportModalInstance === undefined) {
		var script = document.createElement("script");
		script.onload = function () {
			window.ExportModalInstance.init(options);
		};
		script.src = "/Assets/Scripts/ExportModal.js";
		document.head.appendChild(script);
	}
	else
		window.ExportModalInstance.init(options);
}

document.addEventListener("DOMContentLoaded", function () {
	var modal = document.querySelector(".importModal");
	if (modal !== null) {
		try {
			if (!modal.classList.contains("processed")) {
				var options = modal.dataset.importconfig;
				if (options) {
					modal.addEventListener("click",
						function (e) {
							e.preventDefault();
							importModal(JSON.parse(options));
						});
					//Avoid further processing if the event gets triggered multiple times during page load
					modal.classList.add("processed");
				}
			}
		} catch (err) {
			console.error("Unable to initialize import modal");
		}
	}
	var modalExport = document.querySelector(".exportModal");
	if (modalExport !== null) {
		try {
			if (!modalExport.classList.contains("processed")) {
				var options = modalExport.dataset.exportconfig;
				if (options) {
					modalExport.addEventListener("click",
						function (e) {
							e.preventDefault();
							exportModal(JSON.parse(options));
						});
					//Avoid further processing if the event gets triggered multiple times during page load
					modalExport.classList.add("processed");
				}
			}
		} catch (err) {
			console.error("Unable to initialize export modal");
		}
	}
});
$(document).ready(function () {
	if (window.isResponsiveEnabled && !window.pageHandleResponsive) {
		function executeRWDResizeEvents(event) {
			var listener = function (e) {
				if ($('#hdnPageID').length == 0) {
					var breakpointWidth = 600,
					viewportWidth = $(window).width(),
					currentCookieValue = getCookieValue('viewportWidth');
					currentCookieValue = currentCookieValue == null || currentCookieValue == '' ? 0 : Number(currentCookieValue);
					document.cookie = 'viewportWidth=' + viewportWidth + '; path=/';
					if ((viewportWidth < breakpointWidth && currentCookieValue >= breakpointWidth) ||
					(viewportWidth >= breakpointWidth && currentCookieValue < breakpointWidth)) {
						window.location.reload();
					}
				}
			};

			if (window.addEventListener) {
				window.addEventListener(event, listener, false);
			} else if (window.attachEvent) {
				window.attachEvent("on" + event, listener);
			}
		}

		executeRWDResizeEvents('resize');
	}
	
	// Always reset the viewport width now that the document is fully loaded and we can make it more accurate.
	// (Sometimes on load it comes back much wider than it should)	
	document.cookie = 'viewportWidth=' + $(window).width() + '; path=/';
});
;
function Animations() {
	const animationsJsonEl = $('.activeWidgetSkinComponentsOnPageJson');
	const animationsJson = animationsJsonEl.html();
	if (animationsJson)
		this.components = JSON.parse(animationsJson);
	this.selectors = [
		'', '.widgetHeader', '.widgetBody', '.widgetItem', '.widgetTitle', '.widgetMeta', '.widgetBody .bullets',
		'a:link', '.readOn:link', '.widgetViewAll:link', '.widgetRSS:link', '.widgetFooter', '.cpTabs',
		'.cpTabs > li > a:link', '.cpTabPanel',
		'.separated .half:after, .separated .third:after, .separated .fourth:after, .separated .fifth:after, .separated.cpGrid .col:after ', '.miniCalendarHeader', '.miniCalendar > table', '.miniCalendar th', '.miniCalendar td > span', '.miniCalendar td > a', '.miniCalendar .today > span', '.miniCalendar .not > span', '.miniCalendar'
	];
	animationsJsonEl.remove();
}

Animations.prototype.Apply = function () {
	var self = this;
	self.addClassToContainers();
	self.bindContainerTriggers();
	self.addClassAndTriggerToSkinComponents();
};

/// Gets the css class to be applied from the hidden element already present on the page and applies the css class to the Container(Structural/Content)
Animations.prototype.addClassToContainers = function () {
	var $hiddenAnimationElements = $('.hiddenAnimations');
	if ($hiddenAnimationElements != null) {
		$hiddenAnimationElements.each(function () {
			//Get hidden element
			var $hiddenAnimationElement = $(this);
			//Get css class from it
			var cssClass = $hiddenAnimationElement.data('animationclass');
			var startCssClass = cssClass + '_Start';
			var triggerData = $hiddenAnimationElement.data('trigger');
			//Apply that class to parent container, depending on container type
			//For Strctural container
			if ($hiddenAnimationElement.parent('[data-cprole=structuralContainer]') != null) {
				$hiddenAnimationElement.parent('[data-cprole=structuralContainer]').addClass(startCssClass);
				$hiddenAnimationElement.parent('[data-cprole=structuralContainer]').addClass(triggerData);
				setTimeout(function () {
						$hiddenAnimationElement.parent('[data-cprole=structuralContainer]').addClass(cssClass);
						$hiddenAnimationElement.parent('[data-cprole=structuralContainer]').removeClass(startCssClass);
					},
					10);
			}
			//For Content Container
			if ($hiddenAnimationElement.parent('[data-cprole=contentContainer]') != null) {
				$hiddenAnimationElement.parent('[data-cprole=contentContainer]').addClass(startCssClass);
				$hiddenAnimationElement.parent('[data-cprole=contentContainer]').addClass(triggerData);
				setTimeout(function () {
						$hiddenAnimationElement.parent('[data-cprole=contentContainer]').addClass(cssClass);
						$hiddenAnimationElement.parent('[data-cprole=contentContainer]').removeClass(startCssClass);
					},
					10);
			}
		});
	}
};

//Apply animation class to widget skin components on page
Animations.prototype.addClassAndTriggerToSkinComponents = function () {
	var self = this;
	var $widgetsHavingSkinsApplied = $("[class*='skin']");

	if ($widgetsHavingSkinsApplied != null) {
		//array of unique skin ids
		var $uniqueSkinIds = [];
		$widgetsHavingSkinsApplied.each(function () {
			//Get skin class
			var skinClass = $(this).attr('class').match(/.skin\d+/g)[0];
			//Parse skinId off of skin class
			var skinId = parseInt(skinClass.match(/\d+/)[0]);
			//Insert if not there in array
			if ($.inArray(skinId, $uniqueSkinIds) === -1)
				$uniqueSkinIds.push(skinId);
		});

		$.each($uniqueSkinIds, function (index, value) {
			var skinId = value;
			for (var i = 0; i < self.selectors.length; i++) {
				//Find all widget skincomponents on the page using this skin.
				var $skinComponentsOnPageUsingThisSkin = $('.skin' + skinId + ' ' + self.selectors[i] + '');
				var animationClass = self.classAndTrigger(skinId, i).className;
				var animationTrigger = self.classAndTrigger(skinId, i).triggerName;
				var scrollOffset = self.classAndTrigger(skinId, i).scrollOffset;
				if (animationClass !== undefined) {
					//Apply animation class
					$skinComponentsOnPageUsingThisSkin.addClass(animationClass + '_Start').addClass(animationClass).removeClass(animationClass + '_Start');
				}
				if (animationTrigger !== undefined && animationTrigger !== 'hover') {
					//Apply animation trigger
					$skinComponentsOnPageUsingThisSkin.addClass(animationTrigger + 'AnimationTrigger');
					$skinComponentsOnPageUsingThisSkin.data("scrolloffset", scrollOffset);
				}
			}
		});
	}
};

///Get animation class given the component and skin Id
Animations.prototype.classAndTrigger = function (skinId, componentTypeId) {
	var self = this;
	var components = self.components;
	var classAndTrigger = {};
	if (components && components.length > 0) {
		for (var i = 0; i < components.length; i++) {
			if (components[i].WidgetSkinID === skinId && components[i].ComponentType === componentTypeId) {
				classAndTrigger.className = components[i].AnimationClass;
				classAndTrigger.triggerName = components[i].TriggerNameLowerCase;
				classAndTrigger.scrollOffset = components[i].ScrollOffset;
				return classAndTrigger;
			}
		}
	}
	return classAndTrigger;
};

/// Gets the css class to be applied from the hidden element already present on the page and applies the css class to the Nav container(Main/Side)
Animations.prototype.applyAnimationClassToNavContainers = function ($olSubMenu, menuType) {
	var mainMenu = 2;
	var sideMenu = 1;
	var cssClass;
	var triggerData;
	var $mainNavHiddenElement = $('#mainNav').children('[class*="hiddenAnimations"]');
	var $secondaryNavHiddenElement = $('#secondaryNav').children('[class*="hiddenAnimations"]');

	if ($('#mainNav').length > 0 && menuType === mainMenu) {
		cssClass = $mainNavHiddenElement.data('animationclass');
		triggerData = $mainNavHiddenElement.data('trigger');
		$olSubMenu.addClass(cssClass);
		setTimeout(function () {
				$olSubMenu.addClass(cssClass);
				$olSubMenu.addClass('animation-triggered');
				$olSubMenu.removeClass(cssClass + '_Start');
			},
			10);
	};
	if ($('#secondaryNav').length > 0 && menuType === sideMenu) {
		cssClass = $secondaryNavHiddenElement.data('animationclass');
		triggerData = $secondaryNavHiddenElement.data('trigger');
		$olSubMenu.addClass(cssClass);
		setTimeout(function () {
			$olSubMenu.addClass(cssClass);
			$olSubMenu.addClass('animation-triggered');
			$olSubMenu.removeClass(cssClass + '_Start');
		},
			10);
	};
};

Animations.prototype.applyAnimationClassToMegaMenuContainer = function ($megaMenuContainer, timeOut) {
	var cssClass;
	var $megaMenuHiddenElement = $('.megaMenuContainer').children('[class*="hiddenAnimations"]');

	$megaMenuContainer.mouseleave(function () {
		if (!$('.megaMenuContainer').children('.pinned').length > 0) {
			$(this).removeClass('animation-triggered');
		}
	});

	if ($('.megaMenuContainer').length > 0) {
		cssClass = $megaMenuHiddenElement.data('animationclass');
		setTimeout(function () {
				$megaMenuContainer.addClass('animation-triggered');
				$megaMenuContainer.removeClass(cssClass + '_Start');
			},
			timeOut);
	};
};

/// Gets the css class to be applied from the hidden element already present on the page and applies the css class to the Nav container(Main/Side)
Animations.prototype.applyInitial = function ($olSubMenu, menuType) {
	var mainMenu = 2;
	var sideMenu = 1;
	var cssClass;
	var triggerData;
	var $mainNavHiddenElement = $('#mainNav').children('[class*="hiddenAnimations"]');
	var $secondaryNavHiddenElement = $('#secondaryNav').children('[class*="hiddenAnimations"]');
	if ($('#mainNav').length > 0 && menuType === mainMenu) {
		cssClass = $mainNavHiddenElement.data('animationclass');
		triggerData = $mainNavHiddenElement.data('trigger');
		$olSubMenu.addClass(cssClass + '_Start').addClass(triggerData);
	};
	if ($('#secondaryNav').length > 0 && menuType === sideMenu) {
		cssClass = $secondaryNavHiddenElement.data('animationclass');
		triggerData = $secondaryNavHiddenElement.data('trigger');
		$olSubMenu.addClass(cssClass + '_Start').addClass(triggerData);
	};
};

Animations.prototype.bindContainerTriggers = function () {
	$(document)
		.delegate('[class*="clickAnimationTrigger"]',
			'click',
			function () {
				$(this).addClass('animation-triggered');
			});

	$(window).scroll(function () {
		$('.scrollAnimationTrigger').not('.cpCarousel .widgetItem.scrollAnimationTrigger').each(function () {
			var scrollOffsetPercentage = $(this).data('scrolloffset') / 100;
			if (isNaN(scrollOffsetPercentage))
				scrollOffsetPercentage = $(this).children().first().data('scrolloffset') / 100;
			if (scrollOffsetPercentage != 0 && ($(window).scrollTop() + ($(window).height() * (1 - scrollOffsetPercentage))) >
				$(this).offset().top && ($(this).offset().top + $(this).height()) > $(window).scrollTop())
				$(this).addClass('animation-triggered');
		});

		$('.cpCarousel .widgetItem.scrollAnimationTrigger').each(
			function () {
				$(this).addClass('animation-triggered');
			}
		);
	});
};
function StickyStructuralContainers() {
	//Add sticky class to structuralStructuralContainer
	var stickyContainers = $('div.stickyInnerDiv.sticky');
	if (stickyContainers.length > 0) {
		stickyContainers.each(function () {
			$(this).parent().addClass('stickyStructuralContainer stickySticky');
		});
		$.when(window.Pages && window.Pages.rwdReady).done(function() {
			var $stickyContainer = $(".stickyStructuralContainer.stickySticky");
			var pos = $stickyContainer.css('position');
			var isStickyCurrently = pos === 'fixed' || pos === 'sticky';
			var divToolbars = $("#divToolbars")[0];
			if (divToolbars && isStickyCurrently) {
				$stickyContainer.css('top',
					(divToolbars.getBoundingClientRect().height + parseInt($('#divToolbars').css('margin-top'))) + "px");
			}
		});
	}

	//Add sticky class to structural container inside of stickyStructuralContainer
	var structuralContainersUnderSticky = $('div.structuralUnderStickyInnerDiv');
	if (structuralContainersUnderSticky.length > 0) {
		structuralContainersUnderSticky.each(function () {
			//To prevent adding classes to main structural container
			if ($(this).find('div.structuralUnderStickyInnerDiv').length == 0) {
				$(this).parent().addClass($(this).attr('class'));
			}
		});
	}
}

StickyStructuralContainers.prototype.init = function () {
	var self = this;
	var stickyContainers = $('div.stickyInnerDiv.sticky');
	var paddingCalc = $("#divToolbars").outerHeight() - parseInt($("#divToolbars").css('padding-top'));
	if (stickyContainers.length > 0) {
		var bottomOfSticky = $('.stickyStructuralContainer.stickySticky').outerHeight();
		$('#bodyWrapper').addClass('noTransition');
		$('.stickyStructuralContainer.stickySticky').next().css('padding-top', bottomOfSticky + 'px');

		if (stickyContainers.find('.anchorOffset')) {
			document.documentElement.style.setProperty('--header-offset', bottomOfSticky + $('#divToolbars').outerHeight() + 10 + 'px');
		}
	}

	//Adjusted / re - calculate the padding for sibling of sticky structural container.
	window.Pages.onResizeHandlers = window.Pages.onResizeHandlers || [];
	window.Pages.onResizeHandlers.push(function () {
		$('.stickySticky').each(function () {
			var padding = $(this).outerHeight() - 1 + 'px';
			var self = this;
			if ($('.stickySticky').next().attr('data-cprole') == "banner") {
				$('.stickySticky').next().find('[id^="structuralContainer"]').first().css('padding-top', padding);
			} else {
				window.setTimeout(function () {
					$(self).next().css('padding-top', padding);
				}, 500);
			}
		});
	});

	//set up scroll events
	var stickyCollapsedContainers = $('div.stickyInnerDiv.stickyCollapsedWhenScrolled');
	if (stickyCollapsedContainers.length > 0 && !window.isMobileBrowserIncludingTablets) {
		stickyCollapsedContainers.each(function () {
			$(this).parent().addClass('stickyStructuralContainer stickyStickyCollapsedWhenScrolled');
			if ($(this).hasClass('stickyFadeAnimation')) {
				$(this).parent().addClass('stickyStickyFadeAnimation');
			} else if ($(this).hasClass('stickySlideDownAnimation')) {
				$(this).parent().addClass('stickyStickySlideDownAnimation');
			}
		});

		var bottomOfStickyCollapsed = $('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled').offset().top +
			$('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled').outerHeight() -
			$("#divToolbars").height();
		var hideWhenExpandedHeights = self.getAllHeights($('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled .structuralUnderStickyInnerDiv.hideWhenExpanded'));
		var hideWhenCollapsedHeights = self.getAllHeights($('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled .structuralUnderStickyInnerDiv.hideWhenCollapsed'));

		var scrollDownTrigger = self.getStickyScrollBreakpoint(bottomOfStickyCollapsed, hideWhenExpandedHeights, hideWhenCollapsedHeights);

		self.showHideStickyContainers(true, scrollDownTrigger, paddingCalc);

		self.bindScrollForSticky(bottomOfStickyCollapsed,
			hideWhenExpandedHeights,
			hideWhenCollapsedHeights);

		if (stickyCollapsedContainers.find('.anchorOffset')) {
			document.documentElement.style.setProperty('--header-offset', bottomOfStickyCollapsed + $('#divToolbars').outerHeight() + 10 + 'px');
		}
	}
	else {
		stickyContainers = $('div.stickyInnerDiv');
		if (stickyContainers.length > 0 && stickyContainers.find('.structuralUnderStickyInnerDiv.showInMobile').length > 0) {
			stickyContainers.find('.structuralUnderStickyInnerDiv').not('.showInMobile')
				.each(function () {
					$(this).hide();
				});
		}
	}
}

StickyStructuralContainers.prototype.bindScrollForSticky = function (bottomOfStickyCollapsed, hideWhenExpandedHeights, hideWhenCollapsedHeights) {
	var self = this;
	var scrollDownTrigger = self.getStickyScrollBreakpoint(bottomOfStickyCollapsed, hideWhenExpandedHeights, hideWhenCollapsedHeights);
	var paddingCalc = $("#divToolbars").outerHeight() - parseInt($("#divToolbars").css('padding-top'));

	$(window)
		.scroll(function () {
			self.showHideStickyContainers(false, scrollDownTrigger, paddingCalc);
		});
}

StickyStructuralContainers.prototype.showHideStickyContainers = function (initialLoad, scrollDownTrigger, paddingCalc) {
	var $stickyStructuralContainer = $('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled');
	var wS = $(window).scrollTop();

	if (wS >= scrollDownTrigger + paddingCalc) {
		$stickyStructuralContainer.addClass('stickyCollapsed');

		$(".stickyStructuralContainer.stickyStickyCollapsedWhenScrolled")
			.css('top', $("#divToolbars").outerHeight() + "px");

		//Structrual overrides content container so show/hide if content container doesn't have a structural container around it
		$('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled .hideWhenExpanded')
			.each(function () {
				if ($(this).parents('.structuralUnderStickyInnerDiv') === 0) {
					$(this).show();
				}
			});

		$('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled .hideWhenCollapsed')
			.each(function () {
				if ($(this).parents('.structuralUnderStickyInnerDiv') === 0) {
					$(this).hide();
				}
			});

		//Structrual overrides content container so show/hide structural containers and any content containers under it that would be opposite behavior
		$('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled .structuralUnderStickyInnerDiv.hideWhenCollapsed').hide();
		$('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled .structuralUnderStickyInnerDiv.hideWhenExpanded').show();
		$('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled .hideWhenExpanded').find('.hideWhenCollapsed').show();
		$('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled .hideWhenCollapsed').find('.hideWhenExpanded').hide();

		//Animate the slide down
		$('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled.stickyStickyFadeAnimation.stickyCollapsed')
			.addClass('stickyAnimFade');
		$('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled.stickyStickySlideDownAnimation.stickyCollapsed')
			.addClass('stickyAnimSlideDown');

		if ($('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled.stickyCollapsed').length > 0 &&
			!$('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled.stickyCollapsed').hasClass('stickyAnimComplete')
		) {
			setTimeout(function () {
				$('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled.stickyCollapsed')
					.addClass('stickyAnimComplete');
			},
				0);

			$('#stickyPlaceholder').css('height', scrollDownTrigger + paddingCalc + 'px');
			if (!window.isRemoveSetHeights)
				SetHeights();
		}

	} else {
		if (($('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled.stickyCollapsed').length > 0 &&
			$('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled.stickyCollapsed').hasClass('stickyAnimComplete'))
			|| initialLoad) {
			$(".stickyStructuralContainer.stickyStickyCollapsedWhenScrolled").css('top', 'auto');

			//Structrual overrides content container so show/hide if content container doesn't have a structural container around it
			$('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled .hideWhenExpanded')
				.each(function () {
					if ($(this).parents('.structuralUnderStickyInnerDiv') === 0) {
						$(this).hide();
					}
				});

			$('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled .hideWhenCollapsed')
				.each(function () {
					if ($(this).parents('.structuralUnderStickyInnerDiv') === 0) {
						$(this).show();
					}
				});

			//Structrual overrides content container so show/hide structural containers and any content containers under it that would be opposite behavior
			$('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled .structuralUnderStickyInnerDiv.hideWhenExpanded').hide();
			$('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled .structuralUnderStickyInnerDiv.hideWhenCollapsed').show();
			$('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled .hideWhenExpanded').find('.hideWhenCollapsed').hide();
			$('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled .hideWhenCollapsed').find('.hideWhenExpanded').show();

			$('#stickyPlaceholder').css('height', '0');

			$('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled.stickyStickyFadeAnimation.stickyCollapsed')
				.removeClass('stickyAnimFade');
			$('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled.stickyStickySlideDownAnimation.stickyCollapsed')
				.removeClass('stickyAnimSlideDown');
			$stickyStructuralContainer.removeClass('stickyCollapsed stickyAnimComplete');
			if (!window.isRemoveSetHeights)
				SetHeights();
		}
	}
}

StickyStructuralContainers.prototype.getAllHeights = function (elements) {
	var heights = 0;
	if (elements && elements.length > 0) {
		elements.each(function () {
			if ($(this).parents('.structuralUnderStickyInnerDiv.hideWhenExpanded').length === 0
				&& $(this).parents('.structuralUnderStickyInnerDiv.hideWhenCollapsed').length === 0)
				heights += $(this).height();
		});
	}
	return heights;
}


StickyStructuralContainers.prototype.getStickyScrollBreakpoint = function (bottomOfStickyCollapsed, hideWhenExpandedHeights, hideWhenCollapsedHeights) {
	var scrollDownTrigger;
	var $stickyStructuralContainer = $('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled');
	var toolBarAndStickyHeight = $stickyStructuralContainer.outerHeight() -
		$("#divToolbars").height();

	//Logic for when there are structural containers
	if ($('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled .structuralUnderStickyInnerDiv.hideWhenCollapsed').length > 0
		|| $('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled .structuralUnderStickyInnerDiv.hideWhenExpanded').length > 0) {
		//If page loaded and already collapsed and has both show and hide containers
		if ($('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled.stickyCollapsed .structuralUnderStickyInnerDiv.hideWhenCollapsed').length > 0 &&
			$('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled.stickyCollapsed .structuralUnderStickyInnerDiv.hideWhenExpanded').length > 0) {
			scrollDownTrigger = hideWhenCollapsedHeights + toolBarAndStickyHeight - hideWhenExpandedHeights;
			//If page loaded and NOT collapsed and has both show and hide containers
		} else if ($('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled .structuralUnderStickyInnerDiv.hideWhenCollapsed').length > 0 &&
			$('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled .structuralUnderStickyInnerDiv.hideWhenExpanded').length > 0) {
			scrollDownTrigger = toolBarAndStickyHeight;
			//If page loaded and collapsed and has only visible containers when collapsed
		} else if (
			$('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled.stickyCollapsed .structuralUnderStickyInnerDiv.hideWhenExpanded').length > 0) {
			scrollDownTrigger = toolBarAndStickyHeight - hideWhenExpandedHeights;
			//If page loaded and collapsed and has only visible containers when expanded
		} else if (
			$('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled.stickyCollapsed .structuralUnderStickyInnerDiv.hideWhenCollapsed').length > 0) {
			scrollDownTrigger = hideWhenCollapsedHeights + toolBarAndStickyHeight;
			//If page loaded and NOT collapsed and has only one style of visible container (collapsed or non)
		} else if ($('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled .structuralUnderStickyInnerDiv.hideWhenCollapsed').length > 0 ||
			$('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled .structuralUnderStickyInnerDiv.hideWhenExpanded').length > 0) {
			scrollDownTrigger = toolBarAndStickyHeight;
		} else {
			scrollDownTrigger = toolBarAndStickyHeight;
		}
		//Logic for when there are only content containers. Same scenarios as strucrual
	} else {
		if ($('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled.stickyCollapsed .hideWhenCollapsed').length > 0 &&
			$('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled.stickyCollapsed .hideWhenExpanded').length > 0) {
			scrollDownTrigger = hideWhenCollapsedHeights + toolBarAndStickyHeight - hideWhenExpandedHeights;
		} else if ($('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled .hideWhenCollapsed').length > 0 &&
			$('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled .hideWhenExpanded').length > 0) {
			scrollDownTrigger = toolBarAndStickyHeight;
		} else if (
			$('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled.stickyCollapsed .hideWhenExpanded').length > 0) {
			scrollDownTrigger = toolBarAndStickyHeight - hideWhenExpandedHeights;
		} else if (
			$('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled.stickyCollapsed .hideWhenCollapsed').length > 0) {
			scrollDownTrigger = hideWhenCollapsedHeights + toolBarAndStickyHeight;
		} else if ($('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled .hideWhenCollapsed').length > 0 ||
			$('.stickyStructuralContainer.stickyStickyCollapsedWhenScrolled .hideWhenExpanded').length > 0) {
			scrollDownTrigger = toolBarAndStickyHeight;
		} else {
			scrollDownTrigger = toolBarAndStickyHeight;
		}
	}

	return scrollDownTrigger;
};
var commonLocalization = null;
var userAgent = window.navigator.userAgent;
var iOSSafari = userAgent.match(/iP(ad|hone)/i) && userAgent.match(/WebKit/i) && !userAgent.match(/CriOS/i);
var hashRegex = /^(#cc[A-z0-9]{8}-[A-z0-9]{4}-[A-z0-9]{4}-[A-z0-9]{4}-[A-z0-9]{12})$/;
var stickyStructuralContainers = null;

if (window.location.hash && hashRegex.test(window.location.hash)) {
	var targetHash = window.location.hash;
	window.location.hash = '';
	scroll(0, 0);
	setTimeout(function () {
		scroll(0, 0);
	}, 1);
}

if (window.isResponsiveEnabled) {
	loadMediaFramework();
}

function loadMediaFramework() {
	Modernizr.load([
		//test old browsers to use polyfills
		{
			test: window.addEventListener,
			nope: "/Assets/Scripts/RWD/eventListener.js"
		},
		{
			test: window.matchMedia,
			nope: "/Assets/Scripts/RWD/media.match.js"
		},
		{
			test: Modernizr.touch,
			yep: "/Assets/Scripts/RWD/quo.debug.js"
		},
		//Load Media Framework after polyfills if needed
		getAbsoluteUrl("/Assets/Scripts/RWD/MediaFramework.js")
	]);
}

(function () {
	var timer = 4000;
	var incrementTimer = function () { timer = timer * 1.25; };

	var timeOutFunction = function () {
		setTimeout(timeOutFunction, timer);
		setBannerHeights();
		incrementTimer();
	};
	timeOutFunction();
}());

$(document).ready(function () {

	if (!window.isResponsiveEnabled) {
		//Setup execution of events when resizing
		var timeout = 250; //time out to debounce event
		window.Pages = window.Pages || {};
		window.Pages.onResizeHandlers = window.Pages.onResizeHandlers || [];

		if (typeof updateMegaMenusOnBrowserResize === 'function')
			window.Pages.onResizeHandlers.push(updateMegaMenusOnBrowserResize);

		if ($('#hdnWidgetManager').length == 0 && typeof rearrangeFlyoutsOnResizeEvents === 'function')
			window.Pages.onResizeHandlers.push(rearrangeFlyoutsOnResizeEvents);

		function executeResizeEvents(event) {
			var timer;
			var listener = function (e) {
				//Execute the handlers for resize events.
				timer && clearTimeout(timer);
				timer = setTimeout(function () {
					for (var i = 0; i < window.Pages.onResizeHandlers.length; i++) {
						if (typeof (window.Pages.onResizeHandlers[i]) === 'function') {
							window.Pages.onResizeHandlers[i]();
						}
					}
				}, timeout);
			};

			if (window.addEventListener) {
				window.addEventListener(event, listener, false);
			} else if (window.attachEvent) {
				window.attachEvent("on" + event, listener);
			}
		}

		executeResizeEvents('resize');
		executeResizeEvents('orientationChange');
	}

	if (window.isResponsiveEnabled) {
		//Inserts data-label on each table cell in froala editor. Needed for responsive view
		$.each($.find('.fr-view table'), function (tableIndex, tableValue) {
			$(tableValue).parent().addClass('responsiveEditor');
			if($(tableValue).find('thead th').length === 1) {
				$(this).addClass('singleHeaderTable');
			}
			$.each($(tableValue).find('thead th'), function (headerIndex, headerValue) {
				$.each($(tableValue).find('tbody tr'), function (rowIndex, rowValue) {
					$($(rowValue).find('td')[headerIndex]).attr('data-label', headerValue.textContent);
				});
			});
		});
	}

	$('#divToolbars').css('overflow-x', 'visible');
	$('#divToolbars').css('overflow-y', 'visible');

	browserFailOvers();

	var GetCookie = function (cname) {
		var name = cname + "=";
		var ca = document.cookie.split(';');
		for (var i = 0; i < ca.length; i++) {
			var c = ca[i];
			while (c.charAt(0) == ' ') c = c.substring(1);
			if (c.indexOf(name) == 0) return c.substring(name.length, c.length);
		}
		return "";
	}

	var IsLiveEditOn = function () {
		return GetCookie("enableLiveEdit") === "true";
	};

	if (IsLiveEditOn() && isUserLoggedIn()) {
		$('body').addClass('liveEditOn');
	}

	setTimeout(function () {
		$('.mobile.header .mobileMenu').css('top', $('#divToolbars .mobileAlertBar').outerHeight(true));
	}, 10);

	$('div.mobile.header > a.mobileMenu').unbind().click(function (e) {
		e.preventDefault();
		$('#menuPanel').css('top', $('#divToolbars').outerHeight(true));
		$('#menuPanel').toggle('slide');
		window.scrollTo(0, 0);
	});

	// Set Heights really fast at first then after 4 seconds slow it down.
	if (typeof SetHeights === 'function') {

		var stretchContainers = (typeof window.dynamicStretchContainers === 'function');
		var setHeightInterval = setInterval(function () {
			if (!window.isRemoveSetHeights)
				SetHeights();
			if (stretchContainers) window.dynamicStretchContainers();
		}, 50);

		setTimeout(function () {
			clearInterval(setHeightInterval);
			if (IsLiveEditOn()) {
				setInterval(function () {
					if (!window.isRemoveSetHeights)
						SetHeights();
					if (stretchContainers) window.dynamicStretchContainers();
				}, 2000);
			}
		}, 4000);

	}

	//Do not execute for WidgetManager, it has a slightly different logic thats handled in WidgetManager code itself.
	if ($('#hdnWidgetManager').length == 0) {
		if (typeof rearrangeFlyouts === 'function') {
			setTimeout(function () {
				setInterval(function () {
					rearrangeFlyouts(false);
				}, 2000);
			}, 2000);
		}
	}

	window.setTimeout(function () { moveMegaMenusInDOM({ direction: $('#megaMenuPopupDirection').val() }); }, 50);

	frontEnd_removeAttributesFromGraphicLinks();

	if ($('.viewArchiveLink').length > 0) {
		$('body').addClass('cpViewArchiveOn');
	}

	//workaround to make legacy LE modal to be centered correctly
	$("body").append($("#ctl00_ctl00_MainContent_ctl00_liveEditPopupWindow")).append($("#editItemBehavior_backgroundElement"));

	//Workaround to adjust the toolbar dropdowns dynamically based on the .cpToolbar height + some other initialization stuff for Layout.
	Core.Layout.toggleFeatureColumn();

	Core.Layout.attachTabbedWidgetTabHandlers();

	//when clicking on error details links
	$("a.errorDetails").on("click.errorDetails", function (e) {
		e.preventDefault();
		$(this).next("div").slideToggle('fast');
	});

	$("img.imageHover").on("mouseenter.imageHover", function () {
		$(this).attr('src', $(this).data("hover"));
	}).on("mouseleave.imageHover", function () {
		$(this).attr('src', $(this).data("image"));
	});

	// Setup handlers for My Account toolbar
	var $Favorites = $('#favoritesList');
	var $UserMenu = $('.cpToolbar .userMenu .popOut');
	$Favorites.hide();
	$UserMenu.hide();

	$('a[href="#favoritesList"]').click(function (e) {
		e.preventDefault();
		if ($Favorites.hasClass('open')) {
			$Favorites.slideUp(200);
		} else {
			slideToggle($Favorites, true, 200);
		}
		$Favorites.toggleClass('open');
	});

	$('.cpToolbar .popOutContainer > a').click(function (e) {
		if ($UserMenu.hasClass('open')) {
			$UserMenu.slideUp(200);
		} else {
			$UserMenu.slideDown(200);
		}
		$UserMenu.toggleClass('open');
	});

	//Authentication
	$('.requireSignIn').on("keypress", function () {
		$(this).click();
	});

	$('.requireSignIn').on("mouseup", function (event) {
		event.preventDefault();
		$(this).click(function (event) { event.preventDefault(); });
		var $self = $(this);
		$.ajax({
			url: '/Saml/IsSamlLoginEnabled?loginPage=LoginPageMyAccount',
			type: 'GET',
			success: function (response) { RequireSignInSuccess(response, $self) },
			async: false
		});
	});

	function RequireSignInSuccess(response, $elem) {
		if (response.IsSamlLoginEnabled) {
			window.location = '/Admin/Saml/LogonRequest?RelayState=' + window.location.pathname.substring(1);
		}
		else {
			var popupBasedAuthenticationJs = new PopupBasedAuthentication();
			popupBasedAuthenticationJs.requireLoggedIn($elem.attr('href'), '');
		}
	}

	flipHandlesIfNecessary();
	$(window).scroll(function () {
		$('.flipHandles').removeClass('flipHandles');
	});

	if (location.search.toLowerCase().indexOf('preview=yes') > -1 || location.search.toLowerCase().indexOf('print=yes') > -1) {
		$('body').addClass('printPreview');
	}
	else {
		$('body').removeClass('printPreview');
	}

	$('#viewFullWebsiteOnMobile').unbind('click').click(function () {
		var curDate = new Date();
		curDate.setHours(curDate.getHours() + 2);
		var curDateInUTCStringFormat = curDate.toUTCString();
		document.cookie = "forceMobileOff=true; expires=" + curDateInUTCStringFormat + "; path=/";
	});

	$('#backToMobileView').unbind('click').click(function () {
		//expire the cookie
		document.cookie = "forceMobileOff=; expires=Thu, 01 Jan 1970 00:00:00 UTC; path=/";
	});

	if ($("iframe").length > 0) {
		$.getScript('/Assets/Scripts/attrchange.js', function () {
			$.getScript('/Assets/Scripts/attrchange_ext.js', function () {
				$("iframe").attrchange({ callback: function (event) { if (!window.isRemoveSetHeights) SetHeights(); } });
			});
		});
	}

	//Apply classes to parent containers and initialize the stickyStructuralContainers class
	if (typeof (StickyStructuralContainers) === 'function') {
		stickyStructuralContainers = new StickyStructuralContainers();
	}

	var liveEditEnabled = getCookieValue("enableLiveEdit") === "true";
	if (typeof (Animations) === 'function') {
		var animations = new Animations();
		//If live edit enabled, give it a delay for animation classes to apply properly.
		if (liveEditEnabled) {
			setTimeout(function () {
				animations.Apply();
			},
				100);
		} else {
			//Apply animation classes and triggers
			animations.Apply();
		}
	}

	//Fixes image cache problem on iOS safari
	if (/iPhone|iPad|iPod/i.test(navigator.userAgent)) {
		$(".graphicButtonLink").removeAttr("onmouseover");
	}

	$('.cp-modificationProhibited').on("click", function () {
		if ($(this).hasClass('cp-modificationProhibited')) {
			showWidgetDisabledPopUp();
		}
	});

	$('.pageCopyEditWidgetModal').on("click", function () {
		savePageCopy();
	});

	$('.closeEditWidgetModal').on("click", function () {
		closeModal();
	});

	$('.viewMap, .viewMapAsLink').on("click", function () {
		executeViewMapPostClickActions($(this));
	});

	// Mobile skip link handler
	if (window.FeatureToggles && window.FeatureToggles.isActive("CMS.JqueryUpgrade.UpgradeTo371"))
        initializeSkipLinkHandlers();
});

$(window).on("load", function () {
	if (stickyStructuralContainers) {
		stickyStructuralContainers.init();
		// Fire resize handlers to recalculate any height adjustments required after sticky container setup.
		var resizeHandlers = window.Pages && window.Pages.onResizeHandlers || [];
		resizeHandlers.forEach(function (handler) { handler(); });
	}
	$('[class*="pageloadAnimationTrigger"]').each(
		function () {
			$(this).addClass('animation-triggered');
		});

	setTimeout(function () {
		//Fix for blank target elements with rel attributes to avoid phishing attack.
		var blankTargetElements = $("[target='_blank']");
		$.each(blankTargetElements, function (index, element) {
			$(element).attr("rel", "noopener");
		});
	}, 175);

	$('body').find('table').removeAttr('border');
	$('body').find('table').removeAttr('cellspacing');
	$('body').find('table').removeAttr('cellpadding');
});

function flipHandlesIfNecessary() {
	$('.widget').unbind('mouseover').mouseover(function (e) {
		var $this = $(this);
		var $innerCol = $this.parents('div.inner.col');
		if ($innerCol != null && !$innerCol.hasClass('flipHandles') && $('#ysnEnableLiveEdit').is(':checked')) {
			var $removeHandle = $innerCol.find('.handle.remove');
			var $toolbarHeight = $('.cpToolbars').outerHeight(true);
			var $removeHandleHeight = $removeHandle.outerHeight();
			if ($removeHandle.length > 0) {
				var $removeHandleTop = $removeHandle.offset().top - $(document).scrollTop();
				if (parseFloat($toolbarHeight - $removeHandleTop) >= parseFloat($removeHandleHeight / 3))
					$innerCol.addClass('flipHandles');
				else
					$innerCol.removeClass('flipHandles');
			}
		}
	});
}

function reinitCarousels() {
	if (window.carouselsToInit && window.carouselsToInit != undefined) {
		for (var i = 0; i < window.carouselsToInit.length; i++) {
			carouselsToInit[i]();
		}
	}
}

function SetHeights() {
	if ($('body').hasClass('mobile') || $('.pageContent.dragging').length > 1 || !window.Core.performSetHeights) {
		return;
	}

	var originalScrollPosition = $('body').scrollTop();
	if (!originalScrollPosition) {
		originalScrollPosition = $('html').scrollTop();
	}

	// Grow the heights and set them in place.
	$('.inner.col, .outer.col, .resize, .indicator.vertical, div.widgetSpacer').css('height', 'auto');

	var adjustColumnHeights = function () {
		if (typeof CustomAdjustHeights === 'function') {
			CustomAdjustHeights();
			return;
		}
		$('.row.nest.wide').each(function () {
			var height = $(this).outerHeight();
			if (height < 15) {
				height = 15;
			}
			$(this).children('.inner.col, .resize, .indicator.vertical').css('height', height + "px");
		});

		$('.row.outer.wide:not(.autoWidths)').each(function () {
			var height = $(this).outerHeight();
			if (height < 15) {
				height = 15;
			}
			$(this).children('.outer.col, .resize, .indicator.vertical, .row.nest.first.last > .inner.col')
				.css('height', height + "px");
			$(this).find('.outer.col > .row.nest.first.last > .inner.col').css('height', height + "px");

		});
	};

	adjustColumnHeights();

	var autoWidthRowsHandler = function ($this) {
		var height = 15;
		$this.children('.col').each(function () {
			var columnHeight = $(this).outerHeight();
			if (columnHeight > height) {
				height = columnHeight;
			}
		});
		$this.find('.indicator.vertical').height(height);
	};

	$('.outer.row.autoWidths.wide').each(function () {
		autoWidthRowsHandler($(this));

		$(this).find('.row').each(function () {
			autoWidthRowsHandler($(this));
		});
	});

	if (!window.isResponsiveEnabled || getCookieValue("enableLiveEdit") === "true") {
		$('div.widgetSpacer').each(function () {
			var $spacer = $(this);
			var $innerCol = $spacer.parent(),
				$innerRow = $innerCol.parent();

			if ($innerRow.hasClass("first") && $innerRow.hasClass("last")) {
				var outerRowHeight = $innerRow.parent().parent().height(),
					innerColPadding = $innerCol.css("padding-top").replace("px", ""),
					spacerHeight = outerRowHeight - (2 * parseInt(innerColPadding, 10));
				$spacer.height(spacerHeight);
				$innerCol.height(spacerHeight);
			} else {
				var innerColHeight = $innerCol.innerHeight();
				if (innerColHeight > 30)
					$spacer.height(innerColHeight);
				else
					$spacer.height(30); //default spacer height
			}
		});
	}

	adjustColumnHeights();

	$('.outer.col').height('auto');
	if (!iOSSafari) {
		$('html,body').scrollTop(originalScrollPosition);
	}
	setBannerHeights();
	reinitCarousels();
}

function frontEnd_removeAttributesFromGraphicLinks() {
	$('.widgetGraphicLinks li.GraphicLinks > a > img').removeAttr('height').removeAttr('width');
}

function closeModalDialog(behaviourID) {
	var behavior = $find(behaviourID);

	if (behavior)
		behavior.hide();
	else
		closeCpModal();
}

function SetMegaMenuToOpenAbove() {
	moveMegaMenusInDOM({ direction: '1' });
}
function SetMegaMenuToOpenBelow() {
	moveMegaMenusInDOM({ direction: '0' });
}

function changeModalDialogHeight(height, showDetails) {
	var ifr = document.getElementById('liveEditDialog');
	if (ifr)
		ifr.style.height = height + 'px';
	if (showDetails != '')
		document.cookie = "showAddDetails=" + showDetails + ";";
}

function setModalClassForEditItemBehavior(cssClass, showDetails) {
	var behaviorID = $find('editItemBehavior');
	var modalContainer = document.getElementById(behaviorID._PopupControlID);
	if (modalContainer)
		modalContainer.className = 'modalContainer modalContainerCP ' + cssClass;
	if (showDetails != '')
		document.cookie = "showAddDetails=" + showDetails + ";";
}

function loadStyleSheet(href, id) {
	var styleSheet = document.createElement('link');
	styleSheet.type = 'text/css';
	styleSheet.rel = 'stylesheet';
	styleSheet.setAttribute('href', href);
	if (id) {
		styleSheet.setAttribute('id', id);
	}
	var head = document.getElementsByTagName('head')[0];
	head.appendChild(styleSheet);
}

function enableLiveEdit(checked) {

	var $liveEditTabs = $('#liveEditTabs');
	if (checked) {
		document.cookie = "enableLiveEdit=true; expires=12/31/2999; path=/";
		$('.handle').show();
		$('#bodyWrapper').removeClass('noTransition');
		$('#LiveEditCSS').prop('disabled', false);
		var versionTitle = $('#versionHeadLine').text();
		$('#versionHeadLine').html('<a id="versionHeadLineOptionsHandle" href="#" class="handle options" title="Version Options">Version Options</a>');
		$('#versionHeadLine').append($('<span contenteditable="true">').attr('data-versiontitle', versionTitle).text(versionTitle));
		$('#versionHeadLineOptionsHandle').unbind().click(function () {
			versionActions(ContentCollectionInstance.VersionID, 'edit', { "pageID": ContentCollectionInstance.PageID, "versionID": ContentCollectionInstance.VersionID });
		});
		if ($('#404Content').length > 0 || $('#pagePageID').val() === '') {
			location.reload();
		} else if ($liveEditTabs.length > 0) {
			redrawContextualInnerToolbar(window.ContentCollectionInstance.PageID, window.ContentCollectionInstance.VersionID);
			window.ContentCollectionInstance.reloadContent();
			InitializeContainers();
			setTimeout(window.ContentCollectionInstance.refreshNavContainers, 10);
			refreshMegaMenusForLiveEdit();
			showLiveEditElement($liveEditTabs);
			if (liveEditTabs.currentState !== 'closed') {
				liveEditTabs.setBodyPadding();
			}
		} else {
			$.ajax({
				url: '/LiveEditMarkup',
				type: 'GET',
				cache: false,
				async: false,
				success: function (response) {
					var $body = $('body');
					$body.append(response).addClass('liveEditOn');
					loadStyleSheet('/Areas/Pages/Assets/Styles/Home.css');
					loadStyleSheet('/Assets/AdminRefresh/LiveEdit/Styles/LiveEdit.css', 'LiveEditCSS');
					if (!window.ContentCollectionInstance) {
						window.ContentCollectionInstance = new ContentCollectionClass();
						window.ContentCollectionInstance.init();
					}

					if (!window.ResizeData) {
						window.ResizeData = new ResizeClass();
					}
					redrawContextualInnerToolbar(window.ContentCollectionInstance.PageID, window.ContentCollectionInstance.VersionID);
					var moduleID = $('#pageModuleID').val();
					if (moduleID == '') {
						moduleID = 76;
					} else {
						moduleID = Number(moduleID);
					}
					renderLiveEditTabs(moduleID);

					window.ContentCollectionInstance.reloadContent();
					InitializeContainers();
					setTimeout(window.ContentCollectionInstance.refreshNavContainers, 10);
					refreshMegaMenusForLiveEdit();
				},
				error: function (xhr, textStatus, exception) {
					alert("Error: " + xhr.statusText + "\nStatus: " + xhr.status);
				}
			});
		}
		triggerStickyStructuralInitLogic();
		//no responsive when live edit enabled, so default to wide view
		$('.widget.narrow').each(function () {
			var $this = $(this);
			$this.removeClass('narrow').addClass('wide');
		});
		$('.tabbedWidgetNarrow.cpTabs:not(.accordion)').hide();
		$('.tabbedWidget.cpTabs').show();
		Core.Layout.attachTabbedWidgetTabHandlers();
		if ($('script:contains("audioeye.com")').length >= 1) {
			$('script:contains("audioeye.com")').remove();
		}
	}
	else {
		document.cookie = "enableLiveEdit=false; expires=12/31/2999; path=/";
		$('#liveEditMenu').hide();
		setTimeout(window.ContentCollectionInstance.refreshNavContainers, 10);
		refreshMegaMenusForLiveEdit();
		var publishedVersion = $liveEditTabs.length > 0 && $('#hdnVersionStatus').val() === '40';
		var isPublishedVersionAndContainsTabbedWidget = publishedVersion && $('.widgetTabbed').length > 0;
		if (!publishedVersion || isPublishedVersionAndContainsTabbedWidget)
			location.reload();

		$('body.liveEditOn').removeClass('liveEditOn').addClass('liveEditOff');
		window.ContentCollectionInstance.reloadContent(null, null, true);
		loadMediaFramework();
		$('#bodyWrapper').attr('style', '');
		$('#versionHeadLineOptionsHandle').remove();
		$('#versionHeadLine').text($('#versionHeadLine').text());
		$('a.editCustomHtml').hide();
		hideLiveEditElement($liveEditTabs);
		hideLiveEditElement($('#innerToolbar'));
		triggerStickyStructuralInitLogic();
		setTimeout(function () {
			$('#LiveEditCSS').prop('disabled', true);
		}, 666);
		//For some reason, loadMediaFramework doesnt load MediaFramework.js on disabling live edit, this forces it to download it again hence fixing a responsive issue.
		if (window.cpMedia) {
			window.cpMedia = undefined;
			var scriptSrc = '';
			var script = document.createElement('script');
			script.async = false;
			script.type = 'text/javascript';
			scriptSrc = '/Assets/Scripts/RWD/MediaFramework.js';
			script.src = scriptSrc;
			document.body.appendChild(script);
		}
		Core.Layout.attachTabbedWidgetTabHandlers();
		$.ajax({
			url: '/IncludeScript/AudioEyeScript/',
			type: 'GET',
			cache: false,
			async: false,
			success: function (response) {
				if (response) {
					var scriptText = $(response)[2].innerText;
					var s = document.createElement("script");
					s.type = "text/javascript";
					s.text = scriptText;
					document.body.appendChild(s);
				}
			},
			error: function (xhr, textStatus, exception) {
				alert("Error: " + xhr.statusText + "\nStatus: " + xhr.status);
			}
		});
	}

	//If IE, simply perform location.reload in order to execute _rootfrontend file again to render the IE meta tag according to settings. This tag cannot be 
	//added dynamically, so a reload is required for (https://civicplus.tpondemand.com/entity/29964).
	if (window.isie) {
		location.reload();
		return;
	}

}

function refreshMegaMenusForLiveEdit() {
	setTimeout(function () {
		var megaMenuContentContainers = $('[data-displaymegamenu="True"]')
			.map(function () {
				$('#mainNavMegaMenu' + $(this).data('pageid') + ' div.pageContent').html('Loading...');
				return $('#mainNavMegaMenu' + $(this).data('pageid') + ' div.pageContent').data('containerid');
			}).toArray();
		window.Pages.megaMenuLoaded.resolve();
		reloadSharedContentContainers(megaMenuContentContainers);
	}, 200);
}

function showHideUnpublishedItems(checked) {
	if (checked) {
		document.cookie = "showAll=true; expires=12/31/2999 23:59:59; path=/";
		document.getElementById('ysnLiveEditShowMyActionItems').checked = true;
		document.getElementById('ysnLiveEditShowMyActionItems').disabled = true;
	}
}

function showHideUnpublishedItems(checked) {
	var actionItemsCheckbox = document.getElementById('ysnLiveEditShowMyActionItems');
	document.cookie = `showAll=${checked}; expires=12/31/2999 23:59:59; path=/`;

	if (actionItemsCheckbox) {
		actionItemsCheckbox.checked = checked;
		actionItemsCheckbox.disabled = checked;
		showHideActionItems(checked);
	}
}

function showHideActionItems(checked) {
	if (checked)
		document.cookie = "showActionItems=true; expires=12/31/2999 23:59:59; path=/";
	else
		document.cookie = "showActionItems=false; expires=12/31/2999 23:59:59; path=/";

	if (typeof redrawContent == 'function')
		redrawContent();
	else
		location.reload(true);
}

function showHideLiveEditControl(checked) {
	if (checked)
		document.cookie = "showLiveEditControls=true; expires=12/31/2999 23:59:59;";
	else
		document.cookie = "showLiveEditControls=false; expires=12/31/2999 23:59:59;";

	//if (typeof redrawContent == 'function' && !$.browser.msie)
	//	redrawContent();
	//else
	location.reload(true);
}

function disableCustomHtmlRendering(checked) {
	if (checked)
		document.cookie = "disableCustomHtmlRendering=true; expires=12/31/2999 23:59:59; path=/";
	else
		document.cookie = "disableCustomHtmlRendering=false; expires=12/31/2999 23:59:59; path=/";

	location.reload(true);
}

// Print preview.
function pPreview(mode) {
	var printStr = 'PREVIEW=YES';
	if (mode == 1)
		printStr = 'PRINT=YES';
	var url = document.URL.split("#")[0];
	window.open(url + (url.indexOf("?") == -1 ? '?' : '&') + printStr, '_blank');
}

function Navigate(txtAction) {
	if (txtAction.lastIndexOf("&") == txtAction.length - 1)
		txtAction = txtAction.slice(0, -1);
	document.frmNavigate.action = txtAction;
	location.href = txtAction;
}

// todo: rename to openModalDialog, to match closeModalDialog
function openGenericModalDialog(height, className, title, src) {
	document.getElementById('ctl00_LiveEditModalTitle').innerHTML = title;

	behavior = $find('editItemBehavior');

	if (behavior) {
		var ifr = document.getElementById('liveEditDialog');
		var target = document.getElementById('ctl00_ctl00_MainContent_ctl00_liveEditPopupWindow');

		ifr.src = src;

		if (className)
			target.className = 'modalContainer ' + className;
		if (height)
			ifr.style.height = height;

		ifr.style.display = 'block';

		behavior.show();
	}
}

function expandCollapseCategory(catID) {
	if (commonLocalization == null) {
		commonLocalization = GetJson('/Localization/Index');
	}

	if (document.getElementById(catID).style.display == 'none') {
		$('#cat' + catID + ' > a').attr('title', commonLocalization.CollapseThisCategory);
		document.getElementById(catID).style.display = 'block';
		document.getElementById('a_' + catID).innerHTML = '&#9660;';
	}
	else {
		$('#cat' + catID + ' > a').attr('title', commonLocalization.ExpandThisCategory);
		document.getElementById(catID).style.display = 'none';
		document.getElementById('a_' + catID).innerHTML = '&#9658;';
	}
}

function slideToggle(element, bShow, duration) {
	var $el = $(element)
		, visible = $el.is(":visible");

	// if the bShow isn't present, get the current visibility and reverse it
	if (arguments.length == 1) bShow = !visible;

	// if the current visiblilty is the same as the requested state, cancel
	if (bShow == visible) return false;

	$.each($el, function (i, e) {
		var $e = $(e);
		var visible = $e.is(":visible");
		var height = $e.show().height();
		if (!$e.data("originalHeight")) {
			$e.data("originalHeight", height);
		};
		if (!visible) $e.hide().css({ height: 0 });
	});

	// expand the knowledge (instead of slideDown/Up, use custom animation which applies fix)
	if (bShow) {
		$.each($el, function (i, e) {
			$e = $(e);
			$e.show().animate({ height: $e.data("originalHeight") }, { duration: duration });
		});
	} else {
		$el.animate({ height: 0 }, {
			duration: duration, complete: function () {
				$el.hide();
			}
		});
	}
	$el.queue(function () { $el.height('auto'); $(this).dequeue(); });
}


$(window).on("load", function () {
	removeEmptyFeaturedAreas();
	//Run these once here since window.load invokes after images are loaded as well in the DOM. This fixes one of the issues where the landscape version of the page on ipad etc has images all messed up
	//since SetHeights run before the images are loaded in the DOM
	if (window.isResponsiveEnabled) {
		if (typeof SetHeights === 'function' && !window.isRemoveSetHeights)
			SetHeights();
		if (typeof dynamicStretchContainers === 'function')
			dynamicStretchContainers();
	}
});

function GetSignIn() {
	$.get('Services/ForgotPassword.ashx?Username=' + encodeURIComponent(userName), function (data) {
		if (data == -1)
			alert($("#Localization_AccountDoesNotExist").val());
		else
			alert($("#Localization_PasswordReminder").val());
	});
}

function PerformSignIn() {
	var userName = $('#UserName').val();
	var password = $('#Password').val();
	var rememberMe = $('#RememberMe').val();
	var emailAddress = $('#UserName').val();
	var url;
	$.post('/MyAccount/SignInJson?' +
		'UserName=' + userName +
		'&Password=' + password +
		'&RememberMe=' + rememberMe +
		'&EmailAddress=' + emailAddress
		, function (data) {
			if (data.status != 1)
				alert(data.msg);
			else {
				if (data.HasAcceptedTerms == true) {
					url = getParameterByName('url');
					url = url == '' ? window.location.href : url;
					try {
						var redirectUrl = new URL(url, window.location.href);
						if (redirectUrl.protocol === window.location.protocol && redirectUrl.host === window.location.host)
							window.location.href = redirectUrl.href;
					} catch (e) { }
				}
				else {
					window.location.href = '/MyAccount/Terms';
				}
			}
		});
}

function SignIn() {
	window.location.href = '/MyAccount?from=url&url=' + window.location.pathname + window.location.search;
}

function SignOut() {
		var url = '/publiclogin.aspx?out=true&txtRedirect=' + window.location.pathname.substring(1) + window.location.search;
		var dn = getParameterByName('dn');
		if (dn != undefined) {
			if (!dn.toLowerCase().startsWith('http'))
				dn = 'http://' + dn;
			try {
				var dnUrl = new URL(dn);
				if (dnUrl.protocol === 'http:' || dnUrl.protocol === 'https:') {
					url = dn + url;
				}
			} catch (e) { }
		}

		window.location.href = url;
}

function getParameterByName(name) {
	var match = RegExp('[?&]' + name + '=([^&]*)').exec(window.location.search);
	return match && decodeURIComponent(match[1].replace(/\+/g, ' '));
}

//Share current page url using common mvc modal (only for mvc modules)
function openShareEmailPageModal(subject) {
	var emailSubject = (subject && subject != '') ? subject : document.title;
	openCpModal({
		type: 'POST',
		data: {
			subject: emailSubject
		},
		title: 'Email This Page',
		className: 'emailPage',
		url: '/EmailPage/Home/EmailPageModal'
	});
}

function openShareEmailPageModalWithReferrel(subject, urlReferrer) {
	var emailSubject = (subject && subject != '') ? subject : document.title;
	openCpModal({
		type: 'POST',
		data: {
			subject: emailSubject,
			specificUrlReferrer: urlReferrer
		},
		title: 'Email This Page',
		className: 'emailPage',
		url: '/EmailPage/Home/EmailPageModalWithSpecificUrl',
		elementID: 'mvcModal'
	});
}

//Share current page url redirecting to email page (for legacy pages do not support new mvc modal)
function openShareEmailPage(subject) {
	var emailSubject = (subject && subject != '') ? subject : document.title;
	window.location.href = '/EmailPage/?subject=' + escape(emailSubject);
}

//Layout Banners
function renderBanner(bannerPlaceHolderID, bannerSlotID, bannerOption, bannerImage, imageFolder, imageTag) {
	var $bannerSlot = $('#' + bannerSlotID);
	var $bannerPlaceHolder = $('#' + bannerPlaceHolderID);
	var images = bannerOption.BannerImages;
	const sequenceId = Date.now() % (images.length || 1);
	bannerImage = bannerOption.BannerMode === 2 ? images[sequenceId] : bannerImage;
	//render external banner
	$bannerPlaceHolder.replaceWith(createBannerExternalMarkup($bannerSlot, bannerOption, bannerImage, imageFolder, imageTag));
	var stickyContainers = $('div.stickyInnerDiv.sticky');
	if (stickyContainers.length > 0) {
		$('#bodyWrapper .bannerObject.external').css('top', '0');
	}
	setTimeout(function () {
		if (!window.isRemoveSetHeights)
			SetHeights();
	}, 5500);
}

function centerImage(parentWidth, parentHeight, imageWidth, imageHeight) {
	var left = 0;
	//Adjust width to view port proportionally
	var proportionRatio = 1;
	if (imageWidth > parentWidth) {
		proportionRatio = (imageWidth - parentWidth) / imageWidth;
		imageHeight -= imageHeight * proportionRatio;
	} else if (imageWidth < parentWidth) {
		proportionRatio = (parentWidth - imageWidth) / imageWidth;
		imageHeight += imageHeight * proportionRatio;
	}
	imageWidth = parentWidth;

	if (imageHeight < parentHeight) {
		//Readjust Height to view port proportionally
		proportionRatio = (parentHeight - imageHeight) / imageHeight;
		imageWidth += imageWidth * proportionRatio; imageHeight = parentHeight;
		left = -1 * (imageWidth - parentWidth) / 2; //move left to center image
	}

	return { imageWidth: imageWidth, imageHeight: imageHeight, left: left };
}

function createBannerExternalMarkup($bannerSlot, bannerOption, bannerImage, imageFolder, imageTag) {
	if (!window.isRemoveSetHeights)
		SetHeights();
	if (typeof window.dynamicStretchContainers === 'function') window.dynamicStretchContainers();

	var parentHeight = $bannerSlot.parent().outerHeight(true);
	var windowHeight = $(window).height();
	var imageHeight = bannerImage.Height;
	var isBackgroundBanner = $bannerSlot.parent().attr("id") == "bodyWrapper" || $bannerSlot.parent().parent().attr("id") == "bodyWrapper";
	var left = 0;

	if (isBackgroundBanner && windowHeight > parentHeight) {
		parentHeight = windowHeight;
	}

	var imageWidth = bannerImage.Width;
	if (bannerOption.ImageScale) {
		var parentWidth = $bannerSlot.outerWidth(true);
		if (isBackgroundBanner) {
			parentHeight = windowHeight;
		}
		var centerValues = centerImage(parentWidth, parentHeight, imageWidth, imageHeight);
		imageWidth = centerValues.imageWidth;
		parentHeight = centerValues.imageHeight;
		left = centerValues.left;
	}

	if (bannerImage.FileName.startsWith('/')) {
		imageFolder = '';
	}

	var $img = (!imageTag)
		? $('<img class="bannerObject" src="' + imageFolder + bannerImage.FileName + '" alt="">')
		: $(imageTag);
	$img.attr("width", imageWidth + "px");
	$img.css("position", "absolute");

	switch (bannerOption.ImageAlignment) {
		case 1:
		case "Left":
			$img.css("left", "0");
			$img.css("margin", "0");
			break;
		case 2:
		case "Center":
			$img.css("left", "50%");
			$img.css("margin", "0 0 0 -" + Math.round(bannerImage.Width / 2) + "px");
			break;
		case 3:
		case "Right":
			$img.css("right", "0px");
			$img.css("margin", "0");
			break;
	}

	var position = 'position: ' + (bannerOption.ImageScroll ? 'absolute;' : 'fixed;');
	var $bodyWrapper = $("#bodyWrapper");
	if (!bannerOption.ImageScroll) {
		left += parseInt($bodyWrapper.css("padding-left"), 10); //workaround for fixed position
	}

	var html = '<div class="bannerObject external" id="bannerDiv' + $bannerSlot.attr('id') + '" style="min-height: ' + parentHeight + 'px ;top: 0px; left: ' + left + 'px; right: 0; ' + position + ' overflow: hidden; z-index: -999999; padding: 0px;"></div>';

	var $html = $(html).append($img);
	$html.attr('data-image-scale', bannerOption.ImageScale);

	var oldScreenHeight = screen.height;
	var oldScreenWidth = screen.width;
	//bind windows resize event if needed
	$(window).unbind('resize.banner' + bannerOption.SlotName);
	if (bannerOption.ImageScale) {
		$(window).bind('resize.banner' + bannerOption.SlotName, function () {
			if (screen.height !== oldScreenHeight && screen.width !== oldScreenWidth) {
				oldScreenHeight = screen.height;
				oldScreenWidth = screen.width;
				resize($bannerSlot);
			}
		});
	}
	///private function that takes care of resizing when image scale
	var resize = function ($bannerSlotContainer) {
		var parentHeight = $bannerSlotContainer.parent().outerHeight(true);
		var windowHeight = $(window).height();
		var left = 0;

		if (isBackgroundBanner && windowHeight > parentHeight) {
			parentHeight = windowHeight;
		}

		var imageWidth = bannerImage.Width;
		if (bannerOption.ImageScale) {
			var parentWidth = $bannerSlotContainer.outerWidth(true);
			if (isBackgroundBanner) {
				parentHeight = windowHeight;
			}
			var centerValues = centerImage(parentWidth, parentHeight, imageWidth, imageHeight);
			imageWidth = centerValues.imageWidth;
			parentHeight = centerValues.imageHeight;
			left = centerValues.left;
		}

		if (!bannerOption.ImageScroll) {
			left += parseInt($bodyWrapper.css("padding-left"), 10); //workaround for fixed position
		}

		$html.css("left", left);
		$html.find('img.bannerObject').width(imageWidth);
		$html.css("min-height", parentHeight + "px");
	};

	window.Pages = window.Pages || {};
	window.Pages.Banners = window.Pages.Banners || {};
	$(window.Pages.Banners).bind('ExternalBannersResize', function () { resize($bannerSlot); }); // resize external banners

	if (jQuery.resize) {
		$bannerSlot.children(":not(.bannerObject, script)").unbind("resize").bind("resize", function () { $(window.Pages.Banners).trigger('ExternalBannersResize'); });
	}

	if (bannerOption.ImageScale) {
		window.Pages.onResizeHandlers = window.Pages.onResizeHandlers || [];
		window.Pages.onResizeHandlers.push(function () {
			$(window.Pages.Banners).trigger('ExternalBannersResize');
		});
	}

	setTimeout(function () {
		if (!window.isRemoveSetHeights)
			SetHeights();
	}, 8000);
	return $html;
};

function setBannerHeights() {
	$('div.bannerObject.external').each(function () {
		var $this = $(this);
		var isBackgroundBanner = $this.parent().attr("id") == "bodyWrapper" || $this.parent().parent().attr("id") == "bodyWrapper";
		var parentHeight = parentHeight = $this.parent().outerHeight(true);
		var windowHeight = $(window).height();

		if (isBackgroundBanner && windowHeight > parentHeight) {
			parentHeight = windowHeight;
		}

		if ($this.attr('data-image-scale') === "true") {
			$this.css('min-height', $this.find('img').height());
		} else {
			$this.css('min-height', parentHeight);
		}
	});
}

//Creates slideshow markup for external banners.
function createBannerExternalMarkupForSlideshow($bannerSlot, bannerOption, imageFolder, bannerImagesJSON, bannerMode, rotatingSequence) {
	if (!window.isRemoveSetHeights)
		SetHeights();
	if (typeof window.dynamicStretchContainers === 'function') window.dynamicStretchContainers();

	var parentHeight = $bannerSlot.parent().outerHeight(true);
	var parentWidth = $bannerSlot.outerWidth(true);
	var windowHeight = $(window).height();
	var isBackgroundBanner = $bannerSlot.parent().attr("id") == "bodyWrapper" || $bannerSlot.parent().parent().attr("id") == "bodyWrapper";
	var greaterImageWidth = 0, greaterImageHeight = 0;
	var left = 0;

	if (isBackgroundBanner && windowHeight > parentHeight) {
		parentHeight = windowHeight;
	}

	var img = '';
	for (var i = 0; i < bannerImagesJSON.length; i++) {
		var bannerImage = bannerImagesJSON[i];
		var imageWidth = bannerImagesJSON[i].Width;
		if (imageWidth > greaterImageWidth)
			greaterImageWidth = imageWidth;
		if (bannerImagesJSON[i].Height > greaterImageHeight)
			greaterImageHeight = bannerImagesJSON[i].Height;
		var alignmentCss = "";
		switch (bannerOption.ImageAlignment) {
			case 1:
			case "Left":
				alignmentCss += 'left: 0;  margin: 0;';
				break;
			case 2:
			case "Center":
				alignmentCss += ' left: 50%; margin: 0 0 0 -' + Math.round(bannerImage.Width / 2) + 'px;';
				break;
			case 3:
			case "Right":
				alignmentCss += ' right: 0px; margin: 0;';
				break;
		}
		if (bannerOption.ImageScale) {
			imageWidth = parentWidth;
		}

		var tempImageFolder = imageFolder;
		if (bannerImagesJSON[i].FileName.startsWith('/')) {
			tempImageFolder = '';
		}

		if (bannerMode === 'Slideshow') {
			if (i > 0) {
				img += '<img class="bannerObject" width="' + imageWidth + 'px" style="position: absolute; display: none;' + alignmentCss + '" src="data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7" data-is-delayed="true" data-delayed-image="' + tempImageFolder + bannerImagesJSON[i].FileName + '" alt="">';
			} else {
				img += '<img class="bannerObject" width="' + imageWidth + 'px" style="position: absolute;' + alignmentCss + '" src="' + tempImageFolder + bannerImagesJSON[i].FileName + '" data-is-delayed="false" alt="">';
			}
		}
		else if (bannerMode === 'Rotating') {
			if (rotatingSequence < 0) {
				rotatingSequence = 0;
			}
			var rotSeqLessOne = rotatingSequence === 0 ? bannerImagesJSON.length - 1 : rotatingSequence - 1;
			var rotSeqPlusOne = rotatingSequence === bannerImagesJSON.length - 1 ? 0 : rotatingSequence + 1;
			if (bannerOption.SlideshowControlsPosition === 0 || bannerOption.SlideshowControlsPosition === 'None') {
				if (i === rotatingSequence) {
					img += '<img class="bannerObject" width="' + imageWidth + 'px" style="position: absolute;' + alignmentCss + '" src="' + tempImageFolder + bannerImagesJSON[i].FileName + '" data-is-delayed="false" alt="">';
				}
			} else {
				if (i === rotSeqLessOne || i === rotatingSequence || i === rotSeqPlusOne) {
					img += '<img class="bannerObject" width="' + imageWidth + 'px" style="position: absolute;' + alignmentCss + '" src="' + tempImageFolder + bannerImagesJSON[i].FileName + '" data-is-delayed="false" alt="">';
				} else {
					img += '<img class="bannerObject" width="' + imageWidth + 'px" style="position: absolute;' + alignmentCss + '" src="data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7" data-is-delayed="true" data-delayed-image="' + tempImageFolder + bannerImagesJSON[i].FileName + '" alt="">';
				}
			}
		}

	}

	var position = 'position: ' + (bannerOption.ImageScroll ? 'absolute;' : 'fixed;');
	var $bodyWrapper = $("#bodyWrapper");
	if (!bannerOption.ImageScroll) {
		left += parseInt($bodyWrapper.css("padding-left"), 10); //workaround for fixed position
	}
	var html = '<div class="bannerObject external" id="bannerDiv' + $bannerSlot.attr('id') + '" style="min-height: ' + parentHeight + 'px ;top: 0; left: ' + left + 'px; right: 0; ' + position + ' overflow: hidden; z-index: -999999; padding: 0px;">'
		+ img + '</div>';

	var $html = $(html);
	$html.attr('data-image-scale', bannerOption.ImageScale);

	///private function that takes care of resizing when image scale
	var resize = function ($bannerSlotContainer) {
		var parentHeight = $bannerSlotContainer.parent().outerHeight(true);
		var windowHeight = $(window).height();
		var left = 0;

		if (isBackgroundBanner && windowHeight > parentHeight) {
			parentHeight = windowHeight;
		}

		var imageWidth = greaterImageWidth;
		if (bannerOption.ImageScale) {
			var parentWidth = $bannerSlotContainer.outerWidth(true);
			if (isBackgroundBanner) {
				parentHeight = windowHeight;
			}
			var centerValues = centerImage(parentWidth, parentHeight, imageWidth, greaterImageHeight);
			imageWidth = centerValues.imageWidth;
			parentHeight = centerValues.imageHeight;
			left = centerValues.left;
		}

		if (!bannerOption.ImageScroll) {
			left += parseInt($bodyWrapper.css("padding-left"), 10); //workaround for fixed position
		}

		$html.css("left", left);
		$html.find('img.bannerObject').width(imageWidth);
		$html.css("min-height", parentHeight + "px");
	};

	var displayArrows = function ($bannerSlotContainer) {
		var $bannerSlideshowControls = $('#bannerSlideshowControls_' + $bannerSlotContainer[0].id);
		var windowWidth = $(window).width();
		if ($bannerSlideshowControls.length > 0 && windowWidth < 416) {
			$bannerSlideshowControls.hide();
			$html.removeClass('fixedBanner');
		} else if ($bannerSlideshowControls.length > 0 && windowWidth >= 416 && !$bannerSlideshowControls.hasClass("none")) {
			$bannerSlideshowControls.show().removeAttr('style');
			$html.addClass('fixedBanner');
		}
	}

	//bind windows resize event if needed
	$(window).unbind('resize.banner' + bannerOption.SlotName);
	if (bannerOption.ImageScale) {
		window.Pages.onResizeHandlers = window.Pages.onResizeHandlers || [];
		window.Pages.onResizeHandlers.push(function () {
			resize($bannerSlot);
		});
		resize($bannerSlot);
	}

	window.Pages.onResizeHandlers = window.Pages.onResizeHandlers || [];
	window.Pages.onResizeHandlers.push(function () {
		displayArrows($bannerSlot);
	});

	var $bannerSlideshowControls = $('#bannerSlideshowControls_' + $bannerSlot[0].id);
	if ($bannerSlideshowControls.length > 0) {
		$html.addClass('fixedBanner');
	}

	return $html;
}

//Render Slideshow for External banners.
function renderExternalBannerSlideshow(bannerSlotID, bannerOption, imageFolder, bannerMode) {
	var bannerImagesJSON = bannerOption.BannerImages;

	if (!bannerOption.SlideShowSlideTiming) {
		bannerOption.SlideShowSlideTiming = 1;
	}
	if (!bannerOption.SlideShowTransitionTiming) {
		bannerOption.SlideShowTransitionTiming = 1;
	}
	
	const rotatingSequence = bannerMode === 'Rotating' ? Date.now() % (bannerImagesJSON.length || 1) : 0;

	var slideTransitionTiming = parseFloat(bannerOption.SlideShowSlideTiming) * 1000;
	var transitionTiming = parseFloat(bannerOption.SlideShowTransitionTiming) * 1000;

	var $bannerSlot = $('#' + bannerSlotID);
	var imagesHtml = createBannerExternalMarkupForSlideshow($bannerSlot, bannerOption, imageFolder, bannerImagesJSON, bannerMode, rotatingSequence);

	var $imagesHtml = $(imagesHtml);
	$imagesHtml.insertAfter($('#bannerImagesJSON' + bannerOption.BannerOptionID));

	var images = null;
	images = $('div#bannerDiv' + bannerSlotID).find('img.bannerObject');

	if (bannerOption.SlideshowTransition == '1') // 1 corresponds to "none"
	{
		transitionTiming = 0; // A 0 transition time is the same as no transition.
	}

	//Moves banner controls to the body level of the DOM to avoid layout issues with z-index
	if ($('.bannerSlideshowControlsNew.centered.container').length > 0) {
		var $banner = $('#' + bannerSlotID);
		$banner.prepend($('#bannerSlideshowControls_' + bannerSlotID));
	}
	else {
		$(document.body).append($('#bannerSlideshowControls_' + bannerSlotID));
	}

	var ss = new SimpleSlideshow(images, transitionTiming, slideTransitionTiming, bannerMode, parseInt(rotatingSequence));

	if (window.FeatureToggles.isActive("CMS.DesignCenter.EnablePauseAndPlayForBannerSlideshow")) {
		$('#bannerSlideshowPlayPauseBtn_' + bannerSlotID).click(function () {
			ss.togglePause();
		});
	}
}

function SimpleSlideshow($images, transitionTime, slideTime, bannerMode, rotatingSequence) {
	"use strict";
	var slideshowRef = this;
	var $imageToBeShown = null;
	var $currentImage = null;
	var $nextImage = null;
	var $previousImage = null;
	//Setting the imageToShowIndex = -1 make sure that the 1st image showing in the slideshow is the 1st image in the slideshow.
	var imageToShowIndex = -1;
	var nextImageIndex = -1;
	var previousImageIndex = 0;
	var paused = false;
	var numberOfImages = $images.length;
	var onLoad = false;
	//Setting imageToShowIndex as the next rotating sequence if slideshow enhancements are enabled and option is refreshing
	if (bannerMode === 'Rotating' && numberOfImages > 1) {
		imageToShowIndex = rotatingSequence - 1;
		onLoad = true;
	}

	var DEBUG = false;
	var diag = function (msg) {
		if (DEBUG) {
			console.log('[Simple Slideshow] ' + msg);
		}
	};

	slideshowRef.autoFadeNext = function () {
		if (!paused && bannerMode == 'Slideshow') {
			slideshowRef.goToNext();
		}
	};

	slideshowRef.goToNext = function () {
		imageToShowIndex = imageToShowIndex >= (numberOfImages - 1) ? 0 : imageToShowIndex + 1;
		previousImageIndex = imageToShowIndex == 0 ? numberOfImages - 1 : imageToShowIndex - 1;
		nextImageIndex = imageToShowIndex < numberOfImages - 1 ? imageToShowIndex + 1 : 0;

		$imageToBeShown = $($images[imageToShowIndex]);
		$nextImage = $($images[nextImageIndex]);
		if ($nextImage.data('is-delayed') == true) {
			$nextImage.attr('src', $nextImage.data('delayed-image'));
			(new Image()).src = $nextImage.data('delayed-image');
			$nextImage.data('is-delayed', 'false');
			$nextImage.data('delayed-image', '');
		}
		$currentImage = $($images[previousImageIndex]);
		diag('Automatic transition next. Image To Show Index: ' +
			imageToShowIndex +
			', Previous Image Index: ' +
			previousImageIndex);
		slideshowRef.sendBehindFront($imageToBeShown).show();

		if (!onLoad) {

			if (numberOfImages == 1) {
				slideshowRef.sendToFront($currentImage).show();
			}
			else if (numberOfImages > 1) {
				slideshowRef.sendToFront($currentImage)
					.show()
					.fadeOut(transitionTime,
						function () {
							diag('Fade out.');
							setTimeout(slideshowRef.autoFadeNext, slideTime);
						});
			}
		}
	};

	slideshowRef.goToPrevious = function () {
		imageToShowIndex = imageToShowIndex == 0 ? numberOfImages - 1 : imageToShowIndex - 1;
		previousImageIndex = imageToShowIndex == 0 ? numberOfImages - 1 : imageToShowIndex - 1;
		nextImageIndex = imageToShowIndex < numberOfImages - 1 ? imageToShowIndex + 1 : 0;

		$imageToBeShown = $($images[imageToShowIndex]);
		$previousImage = $($images[previousImageIndex]);
		if ($previousImage.data('is-delayed') == true) {
			$previousImage.attr('src', $previousImage.data('delayed-image'));
			(new Image()).src = $previousImage.data('delayed-image');
			$previousImage.data('is-delayed', 'false');
			$previousImage.data('delayed-image', '');
		}
		$currentImage = $($images[nextImageIndex]);

		diag('Automatic transition next. Image To Show Index: ' +
			imageToShowIndex +
			', Previous Image Index: ' +
			nextImageIndex);
		slideshowRef.sendBehindFront($imageToBeShown).show();
		slideshowRef.sendToFront($currentImage)
			.show()
			.fadeOut(transitionTime,
				function () {
					diag('Fade out.');
					setTimeout(slideshowRef.autoFadeNext, slideTime);
				});

	};

	this.pause = function () {
		paused = true;
		togglePlayPauseButton();
	}

	this.play = function () {
		paused = false;
		togglePlayPauseButton();
		slideshowRef.autoFadeNext();
	};

	this.togglePause = function () {
		if (paused) {
			this.play();
		} else {
			this.pause();
		}
	};

	function togglePlayPauseButton() {
		if (paused) {
			$('#slideshowPlayIcon').show();
			$('#slideshowPauseIcon').hide();
			$('.bannerSlideshowPlayPause-btn').attr('aria-label', 'Play Slideshow');
		} else {
			$('#slideshowPlayIcon').hide();
			$('#slideshowPauseIcon').show();
			$('.bannerSlideshowPlayPause-btn').attr('aria-label', 'Pause Slideshow');
		}
	}

	this.sendToFront = function ($image) {
		$image.css('z-index', '3');
		return $image;
	};

	this.sendBehindFront = function ($image) {
		$image.css('z-index', '2');
		return $image;
	};

	window.setTimeout(function () {
		slideshowRef.autoFadeNext();
	}, 500);

	if (bannerMode == 'Rotating' && numberOfImages > 1) {
		this.goToNext();
	}

	var disabledNext = false;
	var disabledPrevious = false;
	$('a[id*="arrowNext_"],button[id*="arrowNext_"]').click(function (e) {
		if (disabledNext) {
			e.preventDefault();
			return false;
		}
		slideshowRef.pause();
		slideshowRef.goToNext();
		disabledNext = true;
		setTimeout(function () {
			disabledNext = false;
		}, 1000);
		e.preventDefault();
	});

	$('a[id*="arrowPrev_"],button[id*="arrowPrev_"]').click(function (e) {
		if (disabledPrevious) {
			e.preventDefault();
			return false;
		}
		slideshowRef.pause();
		slideshowRef.goToPrevious();
		disabledPrevious = true;
		setTimeout(function () {
			disabledPrevious = false;
		}, 1000);
		e.preventDefault();
	});
	onLoad = false;
}

//External banner for videos
function renderBannerVideo(bannerPlaceHolderID, bannerSlotID, bannerVideo, videoTag) {
	if (bannerVideo == null)
		return;
	var $bannerSlot = $('#' + bannerSlotID);
	var $bannerPlaceHolder = $('#' + bannerPlaceHolderID);
	var isBackgroundBanner = $bannerSlot.parent().attr("id") == "bodyWrapper" || $bannerSlot.parent().parent().attr("id") == "bodyWrapper";

	var html = '<div class="bannerObject external" id="bannerDiv' + $bannerSlot.attr('id') + '" style="top: 0; left: 0px; right: 0; position: absolute; overflow: hidden; z-index: -999999; padding: 0px;"></div>';
	var $html = $(html);

	// Inserting html elements before resizing because they need to be on DOM beforehand.
	$bannerPlaceHolder.after($html);
	// Workaround. videotag needs to be added with jquery html() function in order to prevent issues with cache JQuery version has
	$bannerPlaceHolder.next().html(videoTag);

	///private function that takes care of resizing for video
	var resize = function ($bannerSlotContainer) {
		var parentHeight = $bannerSlotContainer.parent().outerHeight(true);
		var windowHeight = $(window).height();
		var left = 0;

		if (isBackgroundBanner && windowHeight > parentHeight) {
			parentHeight = windowHeight;
		}

		var videoHeight = bannerVideo.VideoHeight;
		var videoWidth = bannerVideo.VideoWidth;
		//by default video will always scale
		var parentWidth = $bannerSlotContainer.outerWidth(true);
		if (isBackgroundBanner) {
			parentHeight = windowHeight;
		}
		var centerValues = centerImage(parentWidth, parentHeight, videoWidth, videoHeight);
		videoWidth = centerValues.imageWidth;
		parentHeight = centerValues.imageHeight;
		left = centerValues.left;

		var $html = $bannerPlaceHolder.next();
		$html.css("left", left);
		$html.find('video').css("width", videoWidth + "px");
		$html.find('img').css("width", videoWidth + "px");
		$html.css("min-height", parentHeight + "px");
	};

	//bind windows resize event if needed
	$(window).unbind('resize.banner' + bannerSlotID);
	window.Pages.onResizeHandlers = window.Pages.onResizeHandlers || [];
	window.Pages.onResizeHandlers.push(function () {
		resize($bannerSlot);
	});
	//adjust resize first time
	resize($bannerSlot);

	// Pause/Play button event binding

	if (window.FeatureToggles.isActive("CMS.DesignCenter.EnablePauseAndPlayForBannerVideo") && $('button[id*="bannerVideoToggleBtn_"]').length) {
		$('button[id*="bannerVideoToggleBtn_"]').click(function (e) {
			e.preventDefault();
			toggleVideoPlayPause($bannerSlot)
		});
	}

	if (window.FeatureToggles.isActive("CMS.DesignCenter.EnablePauseAndPlayForBannerVideo")){
		$('button[id*="bannerVideoToggleBtn_"]').on('keydown', function (e) {
			if (e.key === 'Enter' || e.key === ' ') {
				e.preventDefault();
				toggleVideoPlayPause($bannerSlot);
			}
		});
	}
}

function toggleVideoPlayPause($bannerSlot) {
	let playIcon = $("#videoPlayIcon");
	let pauseIcon = $("#videoPauseIcon");
	let iconTitle = $("#videoIconTitle");
	let button = $(".bannerVideoToggle-btn");
	let $bannerObject = $bannerSlot.children('.bannerObject');
	let video = $bannerObject.find('video')[0];
	if (video) {
		if (video.paused) {
			video.play();
			playIcon.hide();
			pauseIcon.show();
			button.attr("aria-label", "Pause Video");
			iconTitle.text("Pause Video");
		} 
		else {
			video.pause();
			playIcon.show();
			pauseIcon.hide();
			button.attr("aria-label", "Play Video");
			iconTitle.text("Play Video");
		}
	}
}

//removes empty feature areas
function removeEmptyFeaturedAreas() {
	$('div[id^="featureArea"]').each(function () {
		if ($(this).html().trim() == "") {
			$(this).parent().css("display", "none");
		}
	});

	$('div[id^="featureContent"]').each(function () {
		if ($(this).html().trim() == "") {
			$(this).css("display", "none");
		}
	});
}

function setWidgetServiceHovering($widgetHeader, bindKeyPress, additionalLogic) {
	if (window.isMobileBrowserIncludingTablets) {
		setWidgetServiceClicking($widgetHeader, bindKeyPress, additionalLogic);
		return;
	}
	var $widgetBody = $widgetHeader.next();
	var $parent = $widgetHeader.parent();
	var timeOutID = null;
	var hideDelay = 400;
	var showDelay = 300;
	$widgetHeader.bind("mouseenter.services", function () {
		clearTimeout(timeOutID);
		timeOutID = setTimeout(function () {
			if (!$widgetBody.is(':visible'))
				adjustFlyoutPosition($parent, $widgetBody);
			$widgetBody.show();
			additionalLogic();
		}, showDelay);
	});
	$widgetHeader.bind("mouseleave.services", function () {
		clearTimeout(timeOutID);
		timeOutID = setTimeout(function () {
			$widgetBody.hide();
			additionalLogic();
		}, hideDelay);
	});
	$widgetBody.children().bind("mouseenter.services", function () {
		clearTimeout(timeOutID);
		timeOutID = setTimeout(function () {
			if (!$widgetBody.is(':visible'))
				adjustFlyoutPosition($parent, $widgetBody);
			$widgetBody.show();
			additionalLogic();
		}, showDelay);
	});
	$widgetBody.children().bind("mouseleave.services", function () {
		clearTimeout(timeOutID);
		timeOutID = setTimeout(function () {
			$widgetBody.hide();
			additionalLogic();
		}, hideDelay);
	});

	if (bindKeyPress) {
		$widgetHeader.unbind("keypress").bind("keypress", function (e) {
			if (e.which === 13) {
				$widgetHeader.trigger("mouseenter.services");
			}
		});
	}
}

function setWidgetServiceClicking($widgetHeader, bindKeyPress, additionalLogic) {
	var $widgetBody = $widgetHeader.next();
	var $parent = $widgetHeader.parent();

	var click = function (e) {
		e.stopPropagation();
		e.preventDefault();
		//If the Flyout is for a megamenu, add a class of forMegaMenu.
		if ($widgetHeader.parents('.megaMenu').length > 0)
			$widgetBody.addClass('forMegaMenu');
		if (!$widgetBody.is(':visible'))
			adjustFlyoutPosition($parent, $widgetBody);
		$widgetBody.toggle();
		if (!window.isRemoveSetHeights)
			SetHeights();

		additionalLogic();
	};

	$widgetHeader.unbind("click").bind("click", click);

	if (bindKeyPress) {
		$widgetHeader.unbind("keypress").bind("keypress", function (e) {
			if (e.which === 13) {
				click(e);
			}
		});
	}
}

function setWidgetSearchClicking($widgetHeader) {
	var $widgetBody = $widgetHeader.next();
	$widgetHeader.unbind("click").bind("click", function (e) {
		e.stopPropagation();
		e.preventDefault();
		$widgetBody.toggle();
		if (!window.isRemoveSetHeights)
			SetHeights()
	});
}

//removes empty feature content areas for IE7
function removeEmptyFeaturedAreas() {
	$('div[class^="html5aside"]').each(function () {
		if ($(this).html().trim() === "") {
			$(this).remove();
		}
	});
}


LayoutClass.prototype.debouncedStyleSheetUpdate = (function () {
	var writeStylesTimeouts = {};
	var DEBUG = false;

	var diag = DEBUG && console && console.log ?
		function (msg) {
			console.log('[Debounced Stylesheet Update] ' + msg);
		}
		: function () { };


	// Original Function We are Debouncing to update stylesheet.
	var updateStylesheet = function (styleSheetID, styles, isEditor) {
		diag('Debounced Call');
		var styleSheet = document.getElementById(styleSheetID);
		isEditor = isEditor === undefined ? false : isEditor;
		if (!isEditor) {
			//Include misc styles from Theme properties.
			if ($('#hdnThemePropertiesMiscStyles').length > 0 && $('#hdnThemePropertiesMiscStyles').html() != '')
				styles += $('#hdnThemePropertiesMiscStyles').html();
		}
		if (styleSheet == null) {
			styleSheet = document.createElement('style');
			styleSheet.id = styleSheetID;
			styleSheet.setAttribute('type', 'text/css');
			var head = document.getElementsByTagName('head')[0];
			head.appendChild(styleSheet);
		}
		if (styleSheet.styleSheet) {
			styleSheet.styleSheet.cssText = styles; //IE only
		} else {
			styleSheet.innerHTML = styles;
		}
	};

	return function (styleSheetID, styles) {
		diag('Bounce Call');
		if (writeStylesTimeouts[styleSheetID]) {
			clearTimeout(writeStylesTimeouts[styleSheetID]);
		}
		writeStylesTimeouts[styleSheetID] = setTimeout(function () { updateStylesheet(styleSheetID, styles); }, 100);
	};
}());

LayoutClass.prototype.updateStylesheet = function (styleSheetID, styles) {
	this.debouncedStyleSheetUpdate(styleSheetID, styles);
};

//Usage:  Core.Layout.toggleFeatureColumn();
LayoutClass.prototype.toggleFeatureColumn = function () {
	if ($('#pageShowFeatureColumn').val() == 'True') {
		$('body').addClass('withFeatureColumn');
		$('#featureColumn').show();
	} else {
		$('body').removeClass('withFeatureColumn');
		$('#featureColumn').hide();
	}

	if (typeof SetHeights == 'function' && !window.isRemoveSetHeights) {
		SetHeights();
	}
};

//Usage:  Core.Layout.stretchContainers(containerSelectors);
LayoutClass.prototype.stretchContainers = function (containerSelectors) {
	this.unstretchContainers(containerSelectors);
	$(containerSelectors).each(function () {
		var $container = $(this);
		var height = $container.outerHeight();
		$container.siblings().each(function () {
			var $sibling = $(this);
			if ($sibling.css('float') != 'none') {
				var siblingHeight = $sibling.outerHeight();
				height = siblingHeight > height ? siblingHeight : height;
			}
		});
		if (height >= ($container.outerHeight())) {

			var desiredContainerHeight = height;
			var $innermostWrapper = $container;

			//The paddings of the container itself needs to be considered as well since we allow for paddings on the structural containers too. 
			//This fixes the height issues we have been having with several sites(neighboring containers dont match height) Vishal
			var $wrappers = $innermostWrapper;

			while ($innermostWrapper.children('.inner').length) {
				$wrappers = $wrappers.add($innermostWrapper.children('.inner'));
				$innermostWrapper = $innermostWrapper.children('.inner');
			}

			var paddingHeight = 0;

			$wrappers.each(function () {
				$this = $(this);
				paddingHeight += parseFloat($this.css('paddingTop')) + parseFloat($this.css('paddingBottom')) + parseFloat($this.css('borderBottomWidth')) + parseFloat($this.css('borderTopWidth'));
			});

			var calculatedHeight = desiredContainerHeight - paddingHeight;
			var heightString = calculatedHeight.toString() + 'px';
			$innermostWrapper.css('min-height', heightString);
		}
	});

	setTimeout(function () {
		if (!window.isRemoveSetHeights)
			SetHeights();
	}, 8000);
};

//Usage:  Core.Layout.unstretchContainers(containerSelectors);
LayoutClass.prototype.unstretchContainers = function (containerSelectors) {
	$(containerSelectors).each(function () {
		var $container = $(this);
		$container.css('min-height', '0');
		$('.inner', $container).css('min-height', '0');
	});
};

//Usage:  Core.Layout.dynamicJavascript(javaScriptCode);
LayoutClass.prototype.dynamicJavascript = function (javaScriptCode) {
	var code = javaScriptCode.replace(/\\'/g, "'").replace(/\\\"/g, "\"").replace(/\\\t/g, "\t").replace(/\\\r/g, "\r").replace(/\\\n/g, "\n").replace(/\\\f/g, "\f").replace(/\\\\/g, "\\");
	var myFucn = new Function(code);
	myFucn();
};

//Usage:  Core.Layout.setColumnWidth($widget);
LayoutClass.prototype.setColumnWidth = function ($widget) {
	var width = $('input.auto.update.widgetOption.common[name=WidgetWidth]').val();
	$widget.closest('div.outer.col').css('width', width ? width + '%' : "");
};

function LayoutClass() {
	this.filter = null;
}

if (!window.Core.Layout) {
	window.Core.Layout = new LayoutClass();
}
// Handles polling login status after ajax requests lazily, and notifying the user if they should be notified.
//  The whole intention of this block of code is to alert the user they aren't logged in after executing an AJAX request 
//  without colliding with other AJAX requests or affecting user experience. This is only done if the user is initially logged in.
$(function () {
	var DEBUG = false;

	var diag = function (msg) {
		if (DEBUG) {
			console.log("[Send User Home If Logged Out] " + msg);
		}
	};

	// We want to be able to start/load this at a specific time (after all of the initial ajax calls complete) so I wrapped it in a function.
	var isLoggedInModule = function () {
		"use strict";
		diag('Initializing Send User Home If Logged Out module.');
		var ns = this;
		var notLoggedInMessage = 'You are no longer logged in. Please login to continue.';
		var loginCheckPending = false;

		ns.initiallyLoggedInWithLiveEdit = $('[href="#liveEditMenu"]').hasClass('on');

		if (!ns.initiallyLoggedInWithLiveEdit) {
			diag('Initialization Stopped - Not initially logged in with Live Edit On.');
			return; // We don't need to notify them they aren't logged in, if they aren't logged in initially.
		}

		$.ajaxSetup({
			beforeSend: function () {
				ns.isAjaxOccuring = true;
			},
			complete: function () {
				diag('Ajax Call Completed (hook) - ' + (loginCheckPending ? 'Login Check is Pending.' : 'Login Check is not already Pending.'));
				ns.isAjaxOccuring = false;
				// If an ajax call is already occuring wait until it is finished. We don't want to affect the user experience with normal AJAX calls.
				if (!loginCheckPending) {
					lazyExecute(
						alertIfUserNotLoggedIn,
						function () { return !ns.isAjaxOccuring; },
						2000
					);
				}
			}
		});

		// Executes a function based on a condition lazily, checking a given predicate at a set interval. 
		var lazyExecute = function (fn, executeWhenPredicate, pollingTime) {
			diag('Lazy Execution Triggered.');
			loginCheckPending = true;
			var intervalExecuter = null;
			var internalHandler = function () {
				diag('Attempting to execute login check.');
				if (executeWhenPredicate()) {
					fn();
					clearInterval(intervalExecuter);
					diag('Login check is no longer pending. Login check invoked.');
					loginCheckPending = false;
				} else {
					diag('AJAX request in progress. Trying again in a short bit.');
				}
			};

			intervalExecuter = setInterval(internalHandler, pollingTime);
		}

		var alertIfUserNotLoggedIn = function () {
			if (!isUserLoggedIn()) {
				alert('You are no longer logged in. Please login to continue.');
				location.href = '/myaccount?from=url&url=' + window.location.pathname;
			}
		}
	}

	// We don't need to start checking if the user is logged in until after everything is probably loaded. I went with 10 seconds
	setTimeout(isLoggedInModule, 6000);
});

//Begin Authentication
if (!window.location.origin) {
	window.location.origin = window.location.protocol + "//" + window.location.host;
}

function popupSignIn(redirectUrl, parameterSet, jsCallback) {
	var popupBasedAuthenticationJs = new PopupBasedAuthentication();

	popupBasedAuthenticationJs.setupJsCallback(jsCallback);
	popupBasedAuthenticationJs.requireLoggedIn(redirectUrl, parameterSet);
};

function internetExplorerMessagingHandlerForPopup(data) {
	var popupBasedAuthenticationJs = new PopupBasedAuthentication();
	popupBasedAuthenticationJs.internetExplorerMessageHandling(data);
};

var internetExplorerVersion = (function () {
	var undef;

	var agent = navigator.userAgent;
	var reg = /MSIE\s?(\d+)(?:\.(\d+))?/i;
	var matches = agent.match(reg);
	if (matches != null) {
		return matches[1];
	}
	return undef;
}());
//End authentication

/********************************************************************************
Logic for moving mega menus in DOM and updating their attributes.
/*******************************************************************************/
function moveMegaMenusInDOM(data) {
	//Grab all the flyouts that are not moved yet and are not tooltips and not in WidgetManager
	var $megaMenus = $('.megaMenuContainer').not('.moved');
	var $megaMenusAlreadyMoved = $('.megaMenuContainer.moved');

	var parentID;
	var marginLeft, marginTop;
	$megaMenus.each(function (e) {
		$(this).appendTo('body');
		$(this).addClass('moved').addClass('megaMenu');
		$(this).css('position', 'absolute');
		$(this).css('z-index', '100');
		updateMegaMenuAttributes($(this), data);
	});
	$megaMenusAlreadyMoved.each(function (e) {
		updateMegaMenuAttributes($(this), data);
	});
}

function updateMegaMenusOnBrowserResize() {
	var $megaMenus = $('.megaMenuContainer');
	$megaMenus.each(function () {
		updateMegaMenuAttributes($(this));
	});
}

function updateMegaMenuAttributes($this, data) {
	//Enums
	var MegaMenuWidthReferenceEnum = { MainNavigation: 0, SitesContent: 1 };
	var MegaMenuPopupDirectionEnum = { Below: 0, Above: 1 };
	var widthReference = $('#megaMenuWidthReference').val();
	var direction = $('#megaMenuPopupDirection').val();

	//If data has been provded, use it!.
	if (data != null) {
		if (data.widthReference != null)
			widthReference = data.widthReference;
		if (data.direction != null)
			direction = data.direction;
	}

	//Get the main nav.
	var $mainNav = $('.siteNav.mainNav');
	//Get the closest parent having a siteWrap class.
	var $siteWrap = $mainNav.closest('.siteWrap');

	var $referenceElem = null;
	if (widthReference == MegaMenuWidthReferenceEnum.MainNavigation)
		$referenceElem = $mainNav;
	else
		$referenceElem = $siteWrap;

	if ($referenceElem != null && $referenceElem.length > 0) {
		//Set left position of megamenu
		$this.css('left', $referenceElem.offset().left + parseFloat($referenceElem.css('padding-left')));
		//Set width of megamenu
		$this.css('width', $referenceElem.width());
	}
	if (direction == null || direction == MegaMenuPopupDirectionEnum.Below) {
		//Set top position of megamenu. 
		//NOTE: Even if the widthReference is set to SitesContent, the top position is calculated always using siteNav.mainNav since thats where the megamenu needs to be displayed. 
		//The siteWrap is some scenarios contains all the content pushing the megamnus to show even below the footer. Vishal
		$this.css('top', $mainNav.offset().top + $mainNav.height());
	}
	else {
		//Set top position of megamenu
		$this.css('top', '-' + ($this.height() - $('#divToolbars').height()) + 'px');
	}
}

function adjustMegaMenuDisplay(widthReference, direction) {
	var $megaMenu = $('#sampleMegaMenu');
	updateMegaMenuAttributes($megaMenu, { widthReference: widthReference, direction: direction });
}
/********************************************************************************
Logic for moving mega menus in DOM and updating their attributes.
/*******************************************************************************/

/***********************************************************************************************************
Common editor rendering functions used in multiple widgets(editor and InfoAdvanced so far)
/**********************************************************************************************************/
function renderSlideshowIfApplicable($element) {
	var $slideshow = $element.find('.cpSlideshowImages');
	$slideshow.each(function (e) {
		var $this = $(this);
		var imgWidth = $this.attr('data-width');
		var className = $this.attr('class');
		var id = $this.attr('id');
		var html = getSlideshowContent(id);

		$this.replaceWith(html);

		// Froala specific data
		//Add froala related classes and image attributes only after html is replaced.
		if (className.indexOf('fr-') > -1) {
			var inlineStyles = $this.attr("style");
			var $images = $('.cpSlideshowJS' + id + ' img[data-delayed-image]');
			$images.each(function (index) {
				var $image = $(this);
				$image.attr("style", inlineStyles);
			});

			// In-line and block classes
			if (this.classList.contains('fr-dib'))
				$('.cpSlideshowJS' + id).addClass('fr-dib');
			else if (this.classList.contains('fr-dii'))
				$('.cpSlideshowJS' + id).addClass('fr-dii');
			// Alignment classes
			if (this.classList.contains('fr-fil'))
				$('.cpSlideshowJS' + id).addClass('fr-fil');
			else if (this.classList.contains('fr-fir'))
				$('.cpSlideshowJS' + id).addClass('fr-fir');
		}

		//If the slideshow has an alignment set in the editor, that alignment needs to be applied to the ol surrounding the images in slideshow.
		if (className.indexOf('float') > 0) {
			var applyClass = '';
			switch (className) {
				case 'cpSlideshowImages floatRight':
					applyClass = 'floatRight';
					break;
				case 'cpSlideshowImages floatLeft':
					applyClass = 'floatLeft';
					break;
				case 'cpSlideshowImages floatCenter':
					applyClass = 'floatCenter';
					break;
				case 'cpSlideshowImages floatJustify':
					applyClass = 'floatJustify';
					break;
			}
			$('.cpSlideshowJS' + id).addClass(applyClass);
		}

		// Changing the container size for an inline slideshow
		var $container = $('#divSlideshow' + id);
		var hasDii = $container.hasClass('fr-dii');
		var hasFil = $container.hasClass('fr-fil');
		var hasFir = $container.hasClass('fr-fir');

		if (hasDii && (hasFil || hasFir || (!hasFil && !hasFir))) {
			$container.css('width', imgWidth);
			$container.css('margin-right', '5px 5px 5px 0px');
		}
	});
}

function getSlideshowContent(slideshowID) {
	var slideshowHtml = '';
	$.ajax({
		url: location.protocol + '//' + window.location.hostname + '/Slideshow/Render/' + slideshowID,
		type: 'POST',
		cache: false,
		async: false,
		success: function (response) {
			slideshowHtml = response;
		},
		error: function (xhr, textStatus, exception) {
			alert("Error: " + xhr.statusText + "\nStatus: " + xhr.status);
		}
	});
	return slideshowHtml;
}
/*****************************************************************************************************************
Common editor rendering functions used in multiple widgets(editor and InfoAdvanced so far) Ends
/****************************************************************************************************************/

///Gets absolute Url when given a relative Url. Useful for certain scenarios where http call is made under https hood and it fails.
function getAbsoluteUrl(relativeUrl) {
	var myAccount = $('#MyAccountLoginPage').val();

	if (myAccount != null && myAccount.toLowerCase() == "true") {
		var basePath = location.protocol + '//' + location.hostname + (typeof (location.port) !== 'undefined' && location.port != '' && location.port != '80' ? ':' + location.port : '');
		return basePath + relativeUrl;
	}

	return relativeUrl;
}

function fireMediaQueriesIfNeeded($parent) {
	if (window.cpMedia) {
		fireMediaElementQueriesOfChildren($parent);
	}
}

function loadMaps() {
	downloadGoogleMapsCustomScript();
}

function googleMapsCustomScriptDownloaded() {
	$('.googleMap').each(function (index, item) {
		new CivicPlus.Maps().LoadMap(item);
	});
}

function downloadGoogleMapsCustomScript() {
	if (typeof (CivicPlus) !== 'undefined')
		CivicPlus.Maps = {};
	var script = document.createElement('script');
	script.type = 'text/javascript';
	if (script.readyState) {  //IE
		script.onreadystatechange = function () {
			if (script.readyState === "loaded" || script.readyState === "complete") {
				script.onreadystatechange = null;
				if (window.googleMapsMode !== 2)//dynamic but lazy loaded
					googleMapsCustomScriptDownloaded();
			}
		};
	} else {  //Others
		script.onload = function () {
			if (window.googleMapsMode !== 2)//dynamic but lazy loaded
				googleMapsCustomScriptDownloaded();
		};
	}

	script.src = '/Assets/Scripts/GoogleMaps.js';
	document.getElementsByTagName("head")[0].appendChild(script);
}

function loadGoogleMapsScript(mode, callback) {
	var script;
	window.googleMapsMode = mode;
	if ($('div.widgetStaffDirectory').length > 0 || $('.DirectoryNormalText .mapWrap .googleMap').length > 0) {
		if ($('#GoogleMapsJS').length == 0) {
			script = document.createElement('script');
			var key = document.getElementById('GoogleMapsKey').value;
			script.id = 'GoogleMapsJS';
			script.type = 'text/javascript';
			if (key != null) {
				script.src = 'https://maps.googleapis.com/maps/api/js?v=3.exp&key=' + encodeURIComponent(key) + '&callback=';
			}
			else {
				script.src = 'https://maps.googleapis.com/maps/api/js?v=3.exp&callback=';
			}
			script.src += mode !== 2 ? callback : 'downloadGoogleMapsCustomScript';
			document.body.appendChild(script);
		} else {
			if (window.google && window.google.maps) {
				loadMaps();
			}
		}
	}
}
$(window).on("load", function () {
	if (typeof targetHash !== "undefined" && targetHash) {
		smoothScroll(targetHash);
	}

	$("a[href*='#']:not(a[href^='/#'], [href='#'])").click(function () {
		var $this = $(this)[0];
		if (hashRegex.test($this.hash)) {
			if ($this.target !== '_blank') {
				if (location.pathname
					.replace(/^\//, '') ===
					this.pathname.replace(/^\//, '') &&
					location.hostname === this.hostname) {
					return smoothScroll(this.hash);
				}
			}
		}
	});
	if (window.location && window.location.hash && $(window.location.hash.replace('/', '\\/')).offset()) {
		setTimeout(function () {
			$('html, body').animate({
				scrollTop: $(window.location.hash).offset().top - calculateTotalFixedHeight()
			}, 300);
		}, 2000);
	}
});
//Calculates the total fixed length
function calculateTotalFixedHeight() {
	var totalHeight = 0;

	if ($('.fixedTop_ts').height()) {
		totalHeight += $('.fixedTop_ts').height();
	}
	if ($('#divToolbars').height()) {
		totalHeight += $('#divToolbars').height();
	}
	if ($('div.cp-Toolbars').height()) {
		totalHeight += $('div.cp-Toolbars').height();
	}
	if ($('.stickyStructuralContainer').height()) {
		totalHeight += $('.stickyStructuralContainer').height();
	}

	return totalHeight;
}

function smoothScroll(hash) {
	var target = $(hash);
	target = target.length ? target : $('[name=' + hash.slice(1) + ']');
	if (!target.length) {
		return false;
	}
	var speed = target[0].getAttribute('data-smoothscrolling') ? 1333 : 1;
	var cpToolbarsHeight = $('div.cp-Toolbars').height() || 0;
	$('html, body').animate({
		scrollTop: target.offset().top - ($('#divToolbars').height() + cpToolbarsHeight)
	}, speed);
	window.location.hash = ''; // Fixes chrome focus bug
	window.location.hash = hash;
	return false;
}

function hideLiveEditElement(target) {
	target.addClass('hide');
	setTimeout(function () {
		target.hide();
	}, 666);
}

function showLiveEditElement(target) {
	target.show();
	setTimeout(function () {
		target.removeClass('hide');
	}, 1);
}

function redrawContextualInnerToolbar(pageID, versionID) {
	if (versionID > 0) {
		var url = '/Pages/Toolbar/Contextual?pageID=' + pageID + '&versionID=' + versionID;
		$.ajax({
			url: url,
			type: 'GET',
			cache: false,
			success: function (response) {
				$('#informativeToolBar').remove();
				var $toolbar = $('#innerToolbar');
				if ($toolbar.length > 0) {
					$toolbar.replaceWith(response);
				} else {
					$('#divToolbars').append(response);
				}
				var toolbarHeight = $('#divToolbars').height() + 'px';
				$('#bodyWrapper').css('padding-top', toolbarHeight);
				if (window.liveEditTabs != null && typeof (window.liveEditTabs.updateMaxHeightTabs) == "function")
					window.liveEditTabs.updateMaxHeightTabs();
				InitializeToolbar();
				if (window.Toolbars && window.Toolbars.ContextualToolbar) {
					$(window.Toolbars.ContextualToolbar).trigger('Loaded');
				}
				$('.widgetSpacer').each(function () {
					var $this = $(this);
					if (!$this.hasClass('adminWrap')) {
						$this.addClass('adminWrap');
					}
				});
			},
			error: function (xhr, textStatus, exception) {
				if (xhr.status !== 0 && xhr.readyState !== 0)
					alert('Error: ' + xhr.statusText + '\nStatus: ' + xhr.status);
			}
		});
	}
}

function InitializeToolbar() {
	$('#lnkVersionSaveAsCopy', '#ulVersionToolBar').unbind('click').click(function (e) {
		e.preventDefault();
		if (!$(this).hasClass('inactive')) {
			SaveAsCopy();
		}
	});
	$('#lnkVersionCancel', '#ulVersionToolBar').unbind('click').click(function (e) {
		e.preventDefault();
		if (!$(this).hasClass('inactive')) {
			AbandonWorkingCopy();
		}
	});
	$('#layoutPageSelector', "#LayoutToolBar").unbind('change').change(function () {
		var pageId = parseInt($(this).val(), 10);
		if (!isNaN(pageId))
			window.location.href = "/" + pageId;
	});
}

function triggerStickyStructuralInitLogic() {
	setTimeout(function () {
		new StickyStructuralContainers().init();
	}, 666);
}

//Function to toggle for widget link.
function toggleWidgetsLink(isEnabled) {
	var widgetLink = $('#innerLnkwidgets');

	if (widgetLink) {
		if (isEnabled) {
			widgetLink.prop('disabled', false).removeClass('cp-modificationProhibited');
			widgetLink[0].href = '#liveEditTab_widgets';
			addClassToWidgets(true);
		}
		else {
			widgetLink.prop('disabled', true).addClass('cp-modificationProhibited');
			widgetLink[0].href = '#';
			addClassToWidgets(false);
		}
	}
}

//Function to show/hide widget disabled popup.
function showWidgetDisabledPopUp() {
	var options = {
		title: 'Widget Edit',
		elementID: 'widgetDisabledPopup',
		autoAdjust: false,
		className: 'adminWrap accountActivation widgetPopUp popup-position',
		isFrontEnd: false,
		htmlContent: '<div class="modalScreen socialSignIn"><p>To modify this content, you\'ll have to create a copy of the page first.</p><p><a class="widget-disable socialButton twitter toTwitterPageSelect pageCopyEditWidgetModal cursor-pointer"><span class="socialButton-override">Copy Page</span></a><a class="widget-disable socialButton twitter toTwitterPageSelect closeEditWidgetModal cursor-pointer" style="padding-left:20px;"><span class="socialButton-override">Cancel</span></a></p></div>',
		draggable: true
	}

	openCpModal(options);
}

//Function to save page copy.
function savePageCopy() {
	SaveAsCopy();
	closeModal();
}

//Function to close the modal
function closeModal() {
	hideModal('.modalContainer.modalContainerCP');
}

//Function to add class to widgets
function addClassToWidgets(isEnabled) {
	var pageContent = $('.pageContent div');
	$.each(pageContent, function (index, item) {
		if (isEnabled) {
			$(item).removeClass('cp-modificationProhibited');
		}
		else {
			$(item).addClass('cp-modificationProhibited');
		}
	});
}

function initializeSkipLinkHandlers() {
	var skipLinks = document.querySelectorAll('#skipToContentLinks a');
	if (!skipLinks.length)
		return;

	skipLinks.forEach(function (link) {
		link.addEventListener('touchstart', handleSkipLinkActivation, { passive: false });
		link.addEventListener('click', handleSkipLinkActivation);
		link.addEventListener('keydown', function (e) {
			if (e.key === 'Enter' || e.key === ' ') {
				handleSkipLinkActivation.call(link, e);
			}
		});
	});
}

function handleSkipLinkActivation(e) {
	e.preventDefault();
	e.stopPropagation();

	// Prevent duplicate execution on mobile (touchstart + click)
	if (e.type === 'click' && e.isTrusted === false) {
		return;
	}

	// handle rapid touches (300ms threshold)
	var now = Date.now();
	var lastActivation = this.dataset.lastActivation || 0;
	if (now - lastActivation < 300) {
		return;
	}

	this.dataset.lastActivation = now;
	var href = this.getAttribute('href');
	if (!href || href[0] !== '#' || href.length < 2)
		return;

	var target = document.querySelector(href);
	if (!target) return;

	var previousTabIndex = target.getAttribute('tabindex');
	if (previousTabIndex === null) {
		target.setAttribute('tabindex', '-1');
	}

	void document.body.offsetHeight;

	// Get toolbar elements
	var $divToolbars = $('#divToolbars');
	var $cpToolbars = $('div.cp-Toolbars');

	var divToolbarsHeight = $divToolbars.length && $divToolbars.is(':visible') ? ($divToolbars.outerHeight(true) || 0) : 0;
	var cpToolbarsHeight = $cpToolbars.length && $cpToolbars.is(':visible') ? ($cpToolbars.outerHeight(true) || 0) : 0;

	var toolbarHeight = divToolbarsHeight + cpToolbarsHeight;
	var $target = $(target);
	var targetOffsetTop = $target.offset().top;
	var targetTop = targetOffsetTop - toolbarHeight;

	$('html, body').animate({scrollTop: targetTop}, 300, function () {
        try {
            target.focus({ preventScroll: true });
		} catch (focusErr) {
            target.focus();
        }
	});
}
;
//Usage:  Core.Layout.attachTabbedWidgetTabHandlers();
LayoutClass.prototype.attachTabbedWidgetTabHandlers = function () {
	window.Pages.tabChangeHandlers = window.Pages.tabChangeHandlers || [];

	storePanelScripts();

	// Hiding all panels, then showing the first one
	$('.tabbedWidget.cpTabPanel').hide();
	$('.tabbedWidget.cpTabPanel.showing').show();

	$('.tabbedWidget.cpTabs, .tabbedWidgetNarrow.cpTabs').each(function (e) {
		var $tabs = $(this);
		var activateTab = function ($tab) {
			$tab.trigger('click');
		};

		// Keyboard navigation for tabs (WCAG 2.1 AA)
		$tabs.find('li:not(.newTab) > a[role="tab"]').unbind('keydown.tabWidget').bind('keydown.tabWidget', function (e) {
			var $currentTab = $(this);
			var $tabList = $currentTab.closest('ol');
			var $allTabs = $tabList.find('li:not(.newTab) > a[role="tab"]');
			var currentIndex = $allTabs.index($currentTab);
			var newIndex = -1;

			switch (e.key) {
				case 'ArrowRight':
				case 'Right':
					e.preventDefault();
					newIndex = (currentIndex + 1) % $allTabs.length;
					break;
				case 'ArrowLeft':
				case 'Left':
					e.preventDefault();
					newIndex = (currentIndex - 1 + $allTabs.length) % $allTabs.length;
					break;
				case 'Home':
					e.preventDefault();
					newIndex = 0;
					break;
				case 'End':
					e.preventDefault();
					newIndex = $allTabs.length - 1;
					break;
				case 'Enter':
				case ' ':
					e.preventDefault();
					$currentTab.click();
					return;
				default:
					return;
			}

			if (newIndex >= 0) {
				// Roving tabindex: set all tabs to -1, focused tab to 0
				$allTabs.attr('tabindex', '-1');
				var $newTab = $allTabs.eq(newIndex);
				$newTab.attr('tabindex', '0').focus();
				activateTab($newTab);
			}
		});

		// Keyboard navigation for accordion buttons
		$tabs.find('li:not(.newTab) > a[role="button"]').unbind('keydown.tabWidget').bind('keydown.tabWidget', function (e) {
			if (e.key === 'Enter' || e.key === ' ') {
				e.preventDefault();
				$(this).click();
			}
		});

		// When a tab is clicked
		$tabs.find('li:not(.newTab) > a').unbind('click').click(function (e) {
			e.preventDefault();
			var $selectedTab = $(this);
			var selectedTabHref = attrOrProps($selectedTab, 'href');
			var hashTabOffset = selectedTabHref.indexOf('#');
			var target = selectedTabHref.slice(hashTabOffset);
			var anchor = this;
			var $target = $('div' + target);
			var $targetli = $selectedTab.parent();
			var isNarrow = $selectedTab.parent().parent().hasClass('tabbedWidgetNarrow');
			var isAccordion = $selectedTab.parent().parent().hasClass('accordion');
			var $parentTab = $selectedTab.parent().parent().parent();
			var tabAriaAttribute = isAccordion ? "aria-expanded" : "aria-selected";

			var showTab = function () {
				$target.addClass('showing').slideDown(function () {
					doAfterSwitchingTabs($target);
					if (document.body.classList.contains("narrow")) {
						anchor.scrollIntoView({behavior: "smooth", block: "start", inline: "nearest"});
					}
				}).css({'overflow': ''});
			}
			// Checking to see if it's showing already
			if (!$target.hasClass('showing') || isAccordion === true) {
				// Making the current tab inactive, and the new one active
				var $widget = $selectedTab.parents('.widget.widgetTabbed');
				
				if (isAccordion === false) {
					if ($widget.hasClass('narrow')) {
						$selectedTab.parent().parent().siblings('ol').find('li.active').removeClass('active');
						$selectedTab.parent().parent().siblings('ol').find('a[' + tabAriaAttribute + '="true"]').attr(tabAriaAttribute, "false");
						// Update roving tabindex on sibling tab lists
						$selectedTab.parent().parent().siblings('ol').find('a[role="tab"]').attr('tabindex', '-1');
					} else {
						$parentTab.children().first().find('li.active').removeClass('active');
						$parentTab.children().first().find('a[' + tabAriaAttribute + '="true"]').attr(tabAriaAttribute, "false");
						// Update roving tabindex on the primary tab list
						$parentTab.children().first().find('a[role="tab"]').attr('tabindex', '-1');
					}
				}

				var tabButton = $widget.find('a[href="' + target + '"]');

				if (isNarrow && isAccordion === false) {
					tabButton.attr(tabAriaAttribute, "true");
					tabButton.attr('tabindex', '0');
					tabButton.parent().addClass('active');
				}
				// Hiding the current panel, showing the new one
				if (!isNarrow) {
					tabButton.attr(tabAriaAttribute, "true");
					tabButton.attr('tabindex', '0');
					tabButton.parent().addClass('active');
					$target.siblings('.showing').removeClass('showing').hide();
					$target.addClass('showing').show();
					doAfterSwitchingTabs($target);
				} else {
					if(isAccordion === false)
					{
						$target.siblings('.showing').removeClass('showing').slideUp();
						
						showTab();
					}
					else
					{
						if ($targetli.hasClass('active')) { 
							$targetli.removeClass('active');
							$targetli.find("a").attr(tabAriaAttribute, "false");
						} else {
							$targetli.addClass('active');
							$targetli.find("a").attr(tabAriaAttribute, "true");
						}
						$target.hasClass('showing') ? $target.removeClass('showing').slideUp() : showTab();
					}
				}

				var $arr = ($target).contents().find('.videoPlayer');
				if ($arr.length > 0) {
					($target).contents().find('.videoPlayer').children('iframe').trigger('onload');
				}
			}
			//Checking any on tab change handlers were added
			if (window.Pages.tabChangeHandlers.length > 0) {
				window.Pages.tabChangeHandlers.forEach(function (car) { car(); });
			}
		});
	});

	$('.nextAction.tabName').unbind('click').click(function (e) {
		e.preventDefault();
	});

	$.fn.hasData = function (prop) {
		return this.filter(
			function () { return $(this).data(prop) != null; }
		);
	};

	var updateChildWidgetClasses = function ($target) {
		// Apply wide/narrow classes based on parent tabbed widget state
		var liveEditEnabled = getCookieValue("enableLiveEdit") === "true";
		if (liveEditEnabled)
			return;

		var $parentTabbedWidget = $target.closest('.widget.widgetTabbed');
		var isParentNarrow = $parentTabbedWidget.hasClass('narrow');
		var $childWidgets = $target.find('.widget').not('[data-elementqueryclasses]');

		if (isParentNarrow) {
			$childWidgets.addClass('narrow').removeClass('wide');
			$target.find('[data-event-list-row="true"]').addClass('narrow').removeClass('wide');
		} else {
			$childWidgets.addClass('wide').removeClass('narrow');
			$target.find('[data-event-list-row="true"]').addClass('wide').removeClass('narrow');
		}
	};

	var doAfterSwitchingTabs = function ($target) {
		fireMediaQueriesIfNeeded($target);

		updateChildWidgetClasses($target);

		if (!window.isRemoveSetHeights)
			SetHeights();

		reloadPanelScripts($target[0]);
	};
};

///Matches the rest of the tab's height with the highest tab in a tabbed widget.
function setTabbedWidgetsTabHeight(tabWrapperID) {
	if (tabWrapperID == null || typeof tabWrapperID == 'undefined') {
		$('.tabbedWidget.cpTabs').not('.left,.right').each(function (e) {
			var $this = $(this);
			$this.find('li').each(function () {
				var $thisLi = $(this);
				if ($thisLi.css('height') != '0px')
					$thisLi.find('a').css('height', $thisLi.css('height'));
			})
		});
	}
	else {
		var $wrapper = $('#' + tabWrapperID);
		$wrapper.find('ol.tabbedWidget.cpTabs').not('.left,.right').find('li').each(function () {
			var $thisLi = $(this);
			if (!$thisLi.is(':hidden')) {
				if ($thisLi.css('height') != '0px')
					$thisLi.find('a').css('height', $thisLi.css('height'));
			}
		});
	}
}

function storePanelScripts() {
	var panels = document.querySelectorAll('.tabbedWidget.cpTabPanel');
	for (var i = 0; i < panels.length; i++) {
		var panel = panels[i];
		if (panel.style.display === 'none') {
			var containers = panel.querySelectorAll('[id^="customHtml"]');
			for (var j = 0; j < containers.length; j++) {
				var container = containers[j];
				if (!container.getAttribute('data-original-html')) {
					var scripts = container.getElementsByTagName('script');
					if (scripts.length > 0) {
						container.setAttribute('data-original-html', container.innerHTML);
					}
				}
			}
		}
	}
}

function reloadPanelScripts(panelElement) {
	if (!panelElement)
		return;

	var containers = panelElement.querySelectorAll('[id^="customHtml"]');
	if (!containers || containers.length === 0)
		return;

	for (var i = 0; i < containers.length; i++) {
		var container = containers[i];
		var originalHtml = container.getAttribute('data-original-html');
		if (!originalHtml)
			continue;

		container.innerHTML = originalHtml;

		var scripts = container.getElementsByTagName('script');
		var scriptsArray = Array.prototype.slice.call(scripts);
		for (var j = 0; j < scriptsArray.length; j++) {
			var oldScript = scriptsArray[j];
			var newScript = document.createElement('script');
			if (oldScript.src) {
				newScript.src = oldScript.src;
			} else {
				newScript.textContent = oldScript.textContent || oldScript.innerText;
			}
			oldScript.parentNode.replaceChild(newScript, oldScript);
		}
	}
}
;
/**** Functions to handle Main/Left Menu on Front-end. *****/
//Global variables
var MAIN_MENU = 2;
var SIDE_MENU = 1;
var menuManager = new MenuManager();
var accessibleMenuToggleEnabled = window.FeatureToggles.isActive("CMS.Pages.AccessibleMainMenuWithIndicators");
//setup RWD menu
window.Pages = window.Pages || {};
$.when(window.Pages.rwdReady, menuManager.mainMenuReady).done(function () {
	menuManager.initRWD();
	var isLoggedBackendUser = $('#helpMenu').length == 1;
	var mql = matchMedia(window.mainMenuMediaQuery);
	var isFirstTimeRun = true;
	var executeOnce = function () {
		fireAllMediaElementQueries();
		if (typeof rearrangeFlyouts === 'function')
			rearrangeFlyouts(false);
		window.Pages.onResizeHandlersExecute();
		window.Pages.toolBarDeferred.resolve();
	};
	if (!mql.matches) {
		executeOnce();
		isFirstTimeRun = false;
	}
	cpMedia.register(window.mainMenuMediaQuery, {
		setup: function () {
			if (isLoggedBackendUser)
				$('div.cpToolbar.user').removeClass('public').addClass('admin adminWrap');
		},
		match: function () {
			$('body').addClass('narrow').removeClass('wide');
			if (isLoggedBackendUser)
				$('div.cpToolbar.user').removeClass('admin adminWrap').addClass('public');
			menuManager.moveToMobile();
			resetBodyPadding && resetBodyPadding();
			if (isFirstTimeRun) {
				executeOnce();
			}
			isFirstTimeRun = false;
		},
		unmatch: function () {
			$('body').addClass('wide').removeClass('narrow');
			if (isLoggedBackendUser)
				$('div.cpToolbar.user').removeClass('public').addClass('admin adminWrap');
			menuManager.moveToDesktop();
		}
	});
});

//Defines menu class to handle Main/Left Menu on Front-end.
function MenuManager() {
	//object variables
	this.$divListSubMenus = null; //hide div where submenus are be placed.
	//this.$containerSubMenus = $('#containerSubMenus'); ; //no hidden div to hold the submenus of main menu being displayed. Workaround to solve z-index issues. could be removed along with its related code after redo layout project!
	this.menuTimeoutID = null;
	this.mainMenuMaxSubMenuLevels = 0;
	this.sideMenuMaxSubMenuLevels = 0;
	this.popdownDelaySideMenu = 500;
	this.popdownDelayTopMenu = 500;
	this.popupDelaySideMenu = 300;
	this.popupDelayTopMenu = 300;
	this.isMainMenuEditable = false;
	this.isMainMenuBeingDragged = false;
	this.isSideMenuEditable = false;
	this.isSideMenuBeingDragged = false;
	this.sideMenuHasCustomLinks = false;
	this.bouncing = true;
	this.initialized = false;
	this.setupDraggable = false;
	this.mobileMainNav = false;
	this.mainMenuInit = false;
	this.mainMenuReady = $.Deferred();
	this.$mainNavHolderMobile = null;
	this.$mainNav = null;
	this.$mainNavMenu = null;
	this.rwdInitNavOpenButton = false;
	this.mainMenuTextResizer = false;
	this.mainMenuTextResizerRatio = 0.5;
	this.isAdjustingMainMenuItemsWidth = false;

	//Initialize menu manger after document was load
	this.init = function () {
		var self = this; // Reference to current object
		if (!self.initialized) {
			self.initialized = true;
			this.resetSubMenus();

			// Hide all submenus after clicking any menu item
			$('li.subMenuItem > a.navMenuItem, li.topMenuItem > a.navMainItem').on("click", function () {
				menuManager.hideAllSubMenus(MAIN_MENU);
				menuManager.hideAllSubMenus(SIDE_MENU);
			});

			// Hide submenu/s when the Escape key is pressed
			document.addEventListener('keydown', function (event) {
				if (event.key === "Escape") {
					// Only handle if focus is inside a menu or submenu
					if (!document.activeElement.closest('#secondaryNav, #mainNav, .mainNavMegaMenu, .secondarySubMenu, .mainSubMenu')) {
						return;
					}

					let menuButton = null;
					const previouslyFocusedElement = document.activeElement;
					
					const subMenuContainer = previouslyFocusedElement.closest("ol[id^='ParentMainItem']");
					if (subMenuContainer) {
						// Get the parent menu item id (e.g., MainItem2151 for ParentMainItem2151)
						const parentItemId = subMenuContainer.id.replace('Parent', '');
						const parentMenuItem = document.getElementById(parentItemId);
						if (parentMenuItem) {

							var liElement = parentMenuItem;
							const currentLi = previouslyFocusedElement.closest('li');
							const currentLiAnchor = currentLi ? currentLi.querySelector('a') : null;
							if (currentLi && currentLi['ariaHasPopup'] === "true" && currentLiAnchor.getAttribute('aria-expanded') === "true") {
								liElement = currentLi;
							}

							menuManager.hideSubMenu($(liElement));

							// Find the anchor inside the parent menu item
							const anchor = liElement.querySelector('a');
							menuButton = anchor;
						}
					} else {
						// Logic for mega menus
						const menuContainer = previouslyFocusedElement.closest("aside[id^='mainNavMegaMenu'], ol[id^='ParentMainItem']");
						if (menuContainer) {
							const itemID = menuContainer.id.match(/\d+$/)?.[0];
							menuButton = findMainNavButton(itemID);
						}

						// Hide all submenus
						menuManager.hideAllSubMenus(MAIN_MENU);
						menuManager.hideAllSubMenus(SIDE_MENU);
					}

					// Return focus to the parent menu item if applicable
					if (menuButton) {
						menuButton.focus();
						self.toggleIndicator($(menuButton));
						return;
					} else {
						self.toggleIndicator($(previouslyFocusedElement));
					}
				}
			});

			/*
				Recursively traverses the menu hierarchy to find the main navigation button corresponding to a given menu item ID.

				@param {string} itemID - The ID of the menu item to find the button for.
				@returns {HTMLElement|null} - The main navigation button element, or null if not found.
			*/
			function findMainNavButton(itemID) {
				const menuItem = document.querySelector(`#MainItem${itemID}`);
				const parent = menuItem && menuItem.closest("ol[id^='ParentMainItem']");
				const parentID = parent && parent.id.match(/\d+$/)?.[0];

				if (parentID) {
					return findMainNavButton(parentID);
				} else {
					return menuItem.querySelector('a');
				}
			}
		}
		if (self.setupDraggable)
			self.attachDragAndDropHandlers();
	};

	this.initRWD = function () {
		var self = this; // Reference to current object
		self.$mainNavHolderMobile = $("#mainNavHolderMobile");
		self.$mainNav = $("#mainNav");
		self.$mainNavMenu = $("#mainNavMenu");
		if (self.isAdjustingMainMenuItemsWidth) {
			self.resetMainItemsWidth("#mainNavMenu");
			self.adjustMainItemsWidth("#mainNavMenu");
		}
		if (self.mainMenuTextResizer)
			self.$mainNavMenu.responsiveMenuText({ ratio: self.mainMenuTextResizerRatio });
	};

	this.resetSubMenus = function () {
		var self = this; // Reference to current object
		if (self.$divListSubMenus != null)
			self.$divListSubMenus.remove();
		var $divListSubMenus = $('<div id="divListSubMenus" style="display:block; visibility:hidden; position:absolute; left:-9999px; top:-9999px;" class="siteNav"></div>');
		//Trick to get width from hidden elements needed to calculate bouncing. If above styles creates problems use hack from http://stackoverflow.com/questions/1472303/jquery-get-width-of-element-when-not-visible-display-none
		var docFrag = $(document.createDocumentFragment());
		docFrag[0].appendChild($divListSubMenus[0].cloneNode(true));
		self.$divListSubMenus = docFrag.children();
		$divListSubMenus.remove();
	};

	/*var $body = $("body");
	this.isNarrowSite = function () {
		//return $body.hasClass("narrow");
		return true;
	};*/
}

MenuManager.prototype.moveToMobile = function () {
	var menusToClose = document.getElementsByClassName("megaMenuContainer");
	for (i = 0; i < menusToClose.length; i++) {
		menusToClose[i].style.display = "none";
	}

	var self = this; // Reference to current object
	if (!self.mobileMainNav) {
		self.mobileMainNav = true;
		if (self.mainMenuTextResizer && self.$mainNavMenu.data('responsiveMenuText'))
			self.$mainNavMenu.responsiveMenuText('dispose');
		if ($("#nav-open-btn").length === 0) {
			$('#divToolbars').append($('<a class="nav-btn" id="nav-open-btn" href="#nav" aria-haspopup="true" role="menu">Menu</a>'));
			self.applyResponsiveMenuSettings();
			self.initNavOpenButton();
		}
		self.$mainNav.before($('<div id="mainNavOriginalPosition" style="display:none;"></div>'));
		self.$mainNavHolderMobile.append(self.$mainNav);
		self.$mainNavMenu.wrap('<div id="rwd-menu" class="rwd-menuwrapper">');
		self.$mainNavMenu.attr('aria-hidden', 'true');
		var $menu = $('#rwd-menu');

		$menu.rwdmenu();

		$menu.onclick = function () { };

		if (self.isAdjustingMainMenuItemsWidth)
			self.resetMainItemsWidth("#mainNavMenu");
	}
};

MenuManager.prototype.applyResponsiveMenuSettings = function () {
	var signInBar = $('.cpToolbar.user.public');
	var adminToolbar = $('.cp-Admin-Legacy-Toolbar');
	var alertBar = $('.alertToolbar');
	var navButton = $("#nav-open-btn");
	var fullWidth = window.fullWidthResponsiveMenu;
	var fixedMenu = window.fixedResponsiveMenu;

	var responsiveMenuButtonHeight = window.responsiveMenuHeight !== '' ? window.responsiveMenuHeight : "34";
	var responsiveMenuButtonHeightWithPx = "";
	if (responsiveMenuButtonHeight.indexOf('px') === -1) responsiveMenuButtonHeightWithPx = responsiveMenuButtonHeight + 'px';

	if (responsiveMenuButtonHeight > 34) //dont assign any inline styles if height < 34
	{
		responsiveMenuButtonHeight -= (signInBar.length && 8);
		var responsiveMenuButtonHeightNavbtnWithpx = "";
		if (responsiveMenuButtonHeight.toString().indexOf('px') === -1) responsiveMenuButtonHeightNavbtnWithpx = responsiveMenuButtonHeight + 'px';

		// Set the button height
		if (responsiveMenuButtonHeight !== "px") navButton.css("height", responsiveMenuButtonHeightNavbtnWithpx).css("line-height", responsiveMenuButtonHeightNavbtnWithpx);
	}

	var responsiveMenuFontSize = window.responsiveMenuFontSize;
	if (responsiveMenuFontSize.indexOf('px') === -1) responsiveMenuFontSize += 'px';
	// Set the button font-size
	if (responsiveMenuFontSize !== 'px') navButton.css('font-size', responsiveMenuFontSize);

	if ((fullWidth || alertBar.length) && !signInBar.length) {
		// Apply styles for the full-width button
		navButton.css("right", "0");
		navButton.css("bottom", "auto");
	}

	// Special case: Forces full-width in order to avoid empty gap on right side of hamburger menu - see bug https://civicplus.tpondemand.com/entity/77390-an-extra-spacing-appears-below-the
	if (!fixedMenu && !fullWidth && !alertBar.length && !signInBar.length) {
		// Apply styles for the full-width button
		navButton.css("right", "0");
		navButton.css("bottom", "auto");
	}
	if (adminToolbar && adminToolbar.length != 0) {
		navButton.css("margin-top", "71px");
	}
	if (signInBar.length) {
		// Fix sign in bar height
		signInBar.css("height", responsiveMenuButtonHeightWithPx);

		// Fix sign in button height
		signInBar.find(".button").css("height", responsiveMenuButtonHeightWithPx).css("line-height", responsiveMenuButtonHeightWithPx);

		if (window.fixedResponsiveMenu && $('#nav-open-btn').length) {
			$('.cpToolbar.user.public').css("z-index" , "101");
		}
	}

	if (fixedMenu) {
		// Move the menu button and menu to #outer-wrap (outside of transform)
		$("#nav-open-btn, #mainNavHolderMobile").appendTo("#outer-wrap");

		// Make fixed
		navButton.css("position", "fixed");

		// Prevent overlapping in the menu
		$("#mainNavHolderMobile").css("position", "fixed").css("top", responsiveMenuButtonHeightWithPx).css("left", "0").css("right", "0").css("bottom", "0");

		// Remove main nav transform, fix z-index issues
		$("<style>.js-navopen #mainNav { transform: none; } .mainNavHolderMobile{ z-index: -1;}.js-navopen .mainNavHolderMobile{z-index: 1;}</style>").appendTo("body");

		// Prevent overflow of the menu
		$("#mainNav").css("overflow-y", "auto");

		if (signInBar.length) {
			// Move the sign in bar to #outer-wrap (outside of transform)
			signInBar.appendTo("#outer-wrap");
			signInBar.css("top", "0").css("left", "0").css("right", "0").css("position", "fixed");
		}

		if (fullWidth || signInBar.length || alertBar.length) {
			// Add padding-top to #outer-wrap (prevent the button/sign in bar from overlapping stuff)
			$("#outer-wrap").css("padding-top", responsiveMenuButtonHeightWithPx);
		}
	}
};

MenuManager.prototype.moveToDesktop = function () {
	var self = this; // Reference to current object
	if (self.mobileMainNav) {
		self.mobileMainNav = false;
		$("#mainNavOriginalPosition").replaceWith(self.$mainNav);
		self.$mainNavMenu.unwrap().rwdmenu('destroy');
		if (self.isAdjustingMainMenuItemsWidth)
			self.adjustMainItemsWidth("#mainNavMenu");
		if (self.mainMenuTextResizer)
			self.$mainNavMenu.responsiveMenuText({ ratio: self.mainMenuTextResizerRatio });
		self.$mainNavMenu.attr('aria-hidden', 'false');
		self.$mainNav.css("overflow-y", "");
	}
};

MenuManager.prototype.initNavOpenButton = function () {
	var self = this; // Reference to current object
	if (self.rwdInitNavOpenButton)
		return;
	self.rwdInitNavOpenButton = true;

	var trim = function (str) {
		return str.trim ? str.trim() : str.replace(/^\s+|\s+$/g, '');
	};

	var hasClass = function (el, cn) {
		return (' ' + el.className + ' ').indexOf(' ' + cn + ' ') !== -1;
	};

	var addClass = function (el, cn) {
		if (!hasClass(el, cn)) {
			el.className = (el.className === '') ? cn : el.className + ' ' + cn;
		}
	};

	var removeClass = function (el, cn) {
		el.className = trim((' ' + el.className + ' ').replace(' ' + cn + ' ', ' '));
	};

	var hasParent = function (el, id) {
		if (el) {
			do {
				if (el.id === id) {
					return true;
				}
				if (el.nodeType === 9) {
					break;
				}
			}
			while ((el = el.parentNode));
		}
		return false;
	};

	var doc = document.documentElement;

	var transform_prop = window.Modernizr.prefixed('transform'),
		transition_prop = window.Modernizr.prefixed('transition'),
		transition_end = (function () {
			var props = {
				'WebkitTransition': 'webkitTransitionEnd',
				'MozTransition': 'transitionend',
				'OTransition': 'oTransitionEnd otransitionend',
				'msTransition': 'MSTransitionEnd',
				'transition': 'transitionend'
			};
			return props.hasOwnProperty(transition_prop) ? props[transition_prop] : false;
		})();

	var inner = $("#inner-wrap")[0];
	var nav_open = false;
	var nav_class = 'js-navopen';

	var closeNavEnd = function (e) {
		if (e && e.target === inner) {
			document.removeEventListener(transition_end, closeNavEnd, false);
		}
		nav_open = false;
	};

	var closeNav = function () {
		if (nav_open) {
			// close navigation after transition or immediately
			var duration = (transition_end && transition_prop) ? parseFloat(window.getComputedStyle(inner, '')[transition_prop + 'Duration']) : 0;
			if (duration > 0) {
				document.addEventListener(transition_end, closeNavEnd, false);
			} else {
				closeNavEnd(null);
			}
		}
		removeClass(doc, nav_class);
		$('ol#mainNavMenu').find('li.rwd-close').remove();
		$('ol#mainNavMenu').attr('aria-hidden', 'true');
		nav_open = false;
	};

	var openNav = function () {
		if (nav_open) {
			return;
		}
		addClass(doc, nav_class);
		$('ol#mainNavMenu').attr('aria-hidden', 'false');
		var $firstItem = $('ol#mainNavMenu').find('li > a').first();
		if (!$firstItem.hasClass('backMenuLink') && !$firstItem.hasClass('closeMenuLink')) {
			var $close = $('<li class="rwd-close"><a href="#" style="z-index:2;" class="navMenuItem closeMenuLink" aria-haspopup="false" role="menuitem">Close Menu</a></li>');
			$('#mainNavMenu').prepend($close);
			$firstItem = $('ol#mainNavMenu').find('li > a').first();
			$firstItem.bind('click', function (e) {
				e.preventDefault();
				closeNav();
			});
		}
		$firstItem.focus();
		nav_open = true;
	};

	var toggleNav = function (e) {
		if (nav_open && hasClass(doc, nav_class)) {
			closeNav();
		} else {
			openNav();
		}
		if (e) {
			e.preventDefault();
		}
	};

	// open nav with main "nav" button
	document.getElementById('nav-open-btn').addEventListener('click', toggleNav, false);

	/* close nav with main "close" button
	document.getElementById('nav-close-btn').addEventListener('click', toggleNav, false);
	*/

	// close nav by touching the partial off-screen content
	document.addEventListener('click', function (e) {
			if (nav_open && !hasParent(e.target, 'mainNav') && !hasClass(e.target, 'nav-btn')) {
				e.preventDefault();
				closeNav();
			}
		},
		true);
};

//Configure a menu. This function should be call only inside $(document).ready
MenuManager.prototype.setupMenu = function (menuID, menuContainerID, menuType, setupDraggable, urlToGetHiddenMenus) {
	var self = this; // Reference to current object
	self.setupDraggable = setupDraggable;

	//Initialize menu manger after DOM elements are loaded
	self.init();

	$('#' + menuID).data('menu', { type: menuType, level: 1, parentID: null });

	//mouse over on menu items from first level
	$("li.topMenuItem", '#' + menuID).mouseover(function (e) {
		e.stopPropagation();
	});

	//mouse enter on menu items from first level
	$("li.topMenuItem", '#' + menuID).mouseenter(function () {
		var $li = $(this);
		self.menuMouseOver($li);
		
		//return false; //do not bubble mouse over events on other nested li tags
	});



	//mouse out on menu items from first level
	$("li.topMenuItem", '#' + menuID).mouseleave(function () {
		var $li = $(this);
		self.menuMouseOut($li);
	});

	$.ajax({
		url: urlToGetHiddenMenus,
		type: 'get',
		async: true, //let the page continue loading/processing meanwhile submenus are loaded
		cache: false,
		dataType: 'html',
		success: function (response) {
			if (response.ErrorMessage)
				alert(response.ErrorMessage);
			else {
				//Append submenus into divListSubMenus and bind menuMouseOut function on mouseleave event for all <li>.
				//create temporary container.
				var tempContainerID = 'dummyDivTemp' + menuID;
				var $tempContainer = $('<div id="' + tempContainerID + '" style="display:block;"></div>').append(response);

				//set up hasChildren class on top level items
				$('#' + menuID + ' li').each(function () {
					var $this = $(this);
					var id = $this.attr('id'); //.replace(/item|leftItem/i, '');
					if (id != null) {
						var $olSubMenu = $tempContainer.find('#Parent' + id).find('li').not(".addPage");
						if ($olSubMenu.length > 0) {
							var $targetElement = $this.children('a').length ? $this.children('a') : $this;
							$targetElement.attr('aria-haspopup', 'true');
							$targetElement.attr('aria-expanded', 'false');
							//for secondary menus, set up withChildren class on top level items
							if (menuType === SIDE_MENU && $this.data("displayMainItemSubMenuIndicators") == "True") {
								$this.children('a.secondaryNavItem').addClass('withChildren');
							}
						}
					}
				});

				//This needs to be executed once to bind draggable to main and side menu's root elements.
				if (setupDraggable) {
					//init drag and drop for menu items
					$('#' + menuID).find("li").each(function () {
						//set as droppable all submenu items on main menu
						self.setupDroppableElement($(this), menuType);
					});

					$('#' + menuID).find(".grippy").each(function () {
						//setup grip icons as draggable elements for main menu sub items.
						self.setupDraggableElement($(this), menuType, '#' + menuContainerID);
					});
				}

				//move submenus to main submenus container.
				self.$divListSubMenus.append($tempContainer.children()/*$('#' + tempContainerID + ' ol')*/);
				$tempContainer.remove();
			}
		},
		error: function (xhr, textStatus, exception) {
			if (xhr.status != 0)
				alert("Error retrieving sub-menus for menuID '" + menuID + "' : " + xhr.statusText + "\ntextStatus: " + xhr.status);
		}
	});
};

//Handles mouse over event on menu items
MenuManager.prototype.menuMouseOver = function ($li) {
	var self = this; // Reference to current object
	if (self.mobileMainNav || ($('.pageContent.cpGrid.dragging').length > 0) || ($('.pageContent.cpGrid.pinned').length > 0)) {
		return;
	}
	this.resetMouseOvers();
	$li.addClass('mouseover');
	$li.parentsUntil('#mainNavMenu').children('a').addClass('mouseover');
	clearTimeout(self.menuTimeoutID);

	var $parent = $li.parent();
	var menuData = $parent.data('menu');

	
	//bail when is currently dragging an item to a different type of menu only if there is not custom links on side menu
	if (!self.sideMenuHasCustomLinks && self.isMainMenuBeingDragged && menuData.type === SIDE_MENU)
		return;
	if (!self.sideMenuHasCustomLinks && self.isSideMenuBeingDragged && menuData.type === MAIN_MENU)
		return;

	if ((menuData.type === MAIN_MENU && self.isMainMenuEditable && (self.isMainMenuBeingDragged || self.isSideMenuBeingDragged) && $li.hasClass('liDropabble') && !$li.hasClass('beingDragged'))
		|| (menuData.type === SIDE_MENU && self.isSideMenuEditable && (self.isMainMenuBeingDragged || self.isSideMenuBeingDragged) && $li.hasClass('liDropabble') && !$li.hasClass('beingDragged'))) {
		//set elements settings when element is over
		$li.addClass('hover');
	}

	$('#mainNav').css('z-index', '10');

	//hide if needed submenus from siblings
	if ($li.hasClass('topMenuItem')) {
		//for top level menu items hide all the submenus before displaying another level
		self.hideAllSubMenus(menuData.type, $li);
	} else {
		//for all other submenus which are naturally inside DOM elements
		$parent.children().each(function () {
			if ($(this) != $li)
				self.hideSubMenu($(this));
		});
	}

	//When opening different type of menu, hide submenus from the other types menu
	if (menuData.type === MAIN_MENU) {
		self.hideAllSubMenus(SIDE_MENU);
	} else if (menuData.type === SIDE_MENU) {
		self.hideAllSubMenus(MAIN_MENU);
	}

	// As the sub-menu is displayed, we let the screen reader know the menu has expanded.
	var $anchor = $li.find('a');
	if ($anchor.length && $anchor.attr('aria-haspopup') === 'true') {
		$anchor.attr('aria-expanded', 'true');
		self.toggleIndicator($anchor);
	}
	else if ($anchor.length && $anchor.attr('aria-haspopup') != 'true') {
		return;
	}
	

	if (!$li.hasClass('beingDragged')) {
		var id = $li.attr('id'); //.replace(/item|leftItem/i, '');
		var pageID = $li.data('pageid');
		var displayMegaMenu = $li.data('displaymegamenu');
		var delay = $li.hasClass('topMenuItem') ? self.popupDelayTopMenu : self.popupDelaySideMenu;

		//This is to pick the list with with the most sub menus and remove the rest. Prevents issue with not displaying all submenus when there is more than one of the same
		//submenu being rendered, since the first rendered sub menu will be used for all submenu duplicates even though it may not have all the submenus included.
		//https://civicplus.tpondemand.com/entity/22126
		var maxSubMenus = 0;
		var maxResult;
		if (self.$divListSubMenus.find('#Parent' + id).length > 1) {
			$(self.$divListSubMenus.find('#Parent' + id)).each(function () {
				if ($(this).children().length > maxSubMenus) {
					maxSubMenus = $(this).children().length;
					maxResult = this;
				}
				$(this).addClass('removal');
			});
		}
		$(maxResult).removeClass('removal');
		self.$divListSubMenus.remove(".removal");
		self.$divListSubMenus.children().append(maxResult);

		var $olSubMenu = self.$divListSubMenus.find('#Parent' + id);
		$('.megaMenuContainer.hover').hide().removeClass('hover');
		if (displayMegaMenu == 'True') {
			if (menuData.type === MAIN_MENU) {
				//Hack to get the heights of the megaMenu elements correctly, in state of display none, incorrect height is being returned. Vishal
				var $megaMenuContainer = $('.mainNavMegaMenu.megaMenuContainer.pageID_' + pageID);
				$megaMenuContainer.css('visibility', 'hidden');
				$megaMenuContainer.css("overflow", "hidden");
				$megaMenuContainer.css("height", "0");
				$megaMenuContainer.css("padding", "0");
				$megaMenuContainer.show();


				//Apply menu animation's start state
				var applyAnimation = function ($megaMenuContainer, timeout) {
					if (typeof (Animations) === 'function') {
						var animations = new Animations();
						animations.applyAnimationClassToMegaMenuContainer($megaMenuContainer, timeout);
					}
				}

				var timeout = delay ? delay : 0;
				self.menuTimeoutID = window.setTimeout(function () {
					updateMegaMenuAttributes($megaMenuContainer);
					fireMediaQueriesIfNeeded($megaMenuContainer);
					$megaMenuContainer.show();
					$megaMenuContainer.css("visibility", "");
					$megaMenuContainer.css("overflow", "");
					$megaMenuContainer.css("height", "");
					$megaMenuContainer.css("padding", "");
					applyAnimation($megaMenuContainer, timeout);
					if (!window.isRemoveSetHeights)
						SetHeights();
				}, timeout);
			}
			else {
				delay ? self.menuTimeoutID = setTimeout(function () {
					$('.secondaryNavMegaMenu.megaMenuContainer.pageID_' + pageID).show();
					if (!window.isRemoveSetHeights)
						SetHeights();
				}, delay) : $('.secondaryNavMegaMenu.megaMenuContainer.pageID_' + pageID).show();
			}
		}

		//For side menu, there isnt a mega menu, so render regular sub menu irrespective of displayMegaMenu setting.
		if (displayMegaMenu != 'True' || menuData.type === SIDE_MENU) {
			if ($olSubMenu.length > 1) {
				//remove duplicates. they might have been added by different menu (main/side)
				$olSubMenu.slice(1).remove();
				$olSubMenu = self.$divListSubMenus.find('#Parent' + id);
				$olSubMenu.removeClass('takenIntoAccount');
			}
			if ($olSubMenu.length == 1) {
				$olSubMenu.removeClass('takenIntoAccount');
				$olSubMenu.data('menu', { type: menuData.type, level: menuData.level + 1, parentID: $parent.attr('id') });
				var allowedToShow = ($olSubMenu.data('menu').type === MAIN_MENU) ? $olSubMenu.data('menu').level <= self.mainMenuMaxSubMenuLevels : $olSubMenu.data('menu').level <= self.sideMenuMaxSubMenuLevels;
				if (allowedToShow) {
					delay ? self.menuTimeoutID = setTimeout(function () { self.showSubMenu($li, $olSubMenu); }, delay) : self.showSubMenu($li, $olSubMenu);
				}
			}
		}

		if (displayMegaMenu == 'True' && !window.isRemoveSetHeights) {
			SetHeights();
		}
	}
};

//Display submenu items
MenuManager.prototype.showSubMenu = function ($li, $olSubMenu) {
	var self = this; // Reference to current object
	var menuType = $olSubMenu.data('menu').type;
	var animations = null;

	//Apply menu animation's start state
	if (typeof (Animations) === 'function') {
		animations = new Animations();
		animations.applyInitial($olSubMenu, menuType);
	}

	//Bind mouseenter for second level submenus
	//Also set the tab index for each <li>
	$olSubMenu.children().each(function () {
		if (self.type != "hidden") {
			var $input = $(this);
		}
		//mouse over on menu items from second and subsequent levels.
		$(this).bind("mouseenter.menu", function () {
			var $li = $(this);
			self.menuMouseOver($li);
			return false; //do not bubble mouse over events on other nested li tags
		});
	});

	//Bind mouseleave for second level submenus
	$olSubMenu.children().each(function () {
		//mouse over on menu items from second and subsequent levels.
		$(this).bind("mouseleave.menu", function () {
			var $li = $(this);
			if (!$olSubMenu.is(":hover")) {
				self.menuMouseOut($li);
			}
		});
	});

	var bouncing = self.bouncing &&
		((menuType === MAIN_MENU && $olSubMenu.data('menu').level > 2)
			|| (menuType === SIDE_MENU));

	var $parentOl = $li.parent();
	var parentLeftLi = $li.offset().left;

	self.moveSubMenuInDOM($olSubMenu, menuType);

	if (bouncing) {
		//logic to bounce submenu to be opened if needed.
		var subMenuWidth = $olSubMenu.outerWidth(); //get width from submenu which is hidden (required to add extra styles on divListSubMenus)

		var isOpeningToRight = !$parentOl.hasClass('flipLeft');
		if (isOpeningToRight) {
			var winWidth = $(window).width();
			var parentWidthLi = $parentOl.outerWidth();

			if ((winWidth - parentLeftLi - parentWidthLi) <= (subMenuWidth + 3)) {
				$olSubMenu.addClass('flipLeft');
			}
		} else {
			if ((parentLeftLi) <= (subMenuWidth + 3)) {
				$olSubMenu.addClass('flipRight');
			} else {
				$olSubMenu.addClass('flipLeft');
			}
		}
	}
	self.positionSubMenuInDOM($olSubMenu, $li, menuType);

	if (self.setupDraggable) {
		//init drag and drop for menu items
		$olSubMenu.find("li").each(function () {
			//set as droppable all submenu items on main menu
			self.setupDroppableElement($(this), menuType);
		});

		$olSubMenu.find(".grippy").each(function () {
			//setup grip icons as draggable elements for main menu sub items.
			self.setupDraggableElement($(this), menuType, menuType == MAIN_MENU ? '#mainNav' : '#secondaryNav');
		});
	}

	//If Animation manager toggled on, implement the menu animations
	if (animations != null) {
		animations.applyAnimationClassToNavContainers($olSubMenu, menuType);
	}
};

//Handles mouse out event on menu items
MenuManager.prototype.menuMouseOut = function ($li) {
	var self = this; // Reference to current object

	if (self.mobileMainNav) return;
	var parentMenuID = $li.parents('ol').attr('id') || '';
	var menuType = $li.attr('id') != null ? ($li.attr('id').match(/MainItem/) != null ? MAIN_MENU : SIDE_MENU) : (($li.hasClass('secondaryNavItem') || parentMenuID.match(/SideItem/) != null) ? SIDE_MENU : MAIN_MENU);
	//log("popdownDelaySideMenu" + popdownDelaySideMenu);
	var delay = $li.hasClass('topMenuItem') ? self.popdownDelayTopMenu : self.popdownDelaySideMenu;
	if (self.isMainMenuBeingDragged || self.isSideMenuBeingDragged) {
		//reset elements settings
		$li.removeClass('hover');
	}

	$('#mainNav').css('z-index', '');

	clearTimeout(self.menuTimeoutID);
	self.menuTimeoutID = setTimeout(function () {
		/*if ($li.parent().data('menu').type == MAIN_MENU)
		self.hideAllSubMenus(MAIN_MENU); //workaround related with containerSubMenus
		else*/
		$('#bodyWrapper').animate({ 'padding-bottom': 0 }, 'fast');
		$('#bodyWrapper').data('maxPaddingBottom', '0');
		$('.takenIntoAccount').removeClass('takenIntoAccount');
		if (!$li.hasClass('topMenuItem'))
			self.hideAllSubMenus(menuType, $li);
		else
			self.hideSubMenu($li);
	}, delay);
};

MenuManager.prototype.hideAllSubMenus = function (menuType, $currentLi) {
	var self = this; // Reference to current object
	switch (menuType) {
		case SIDE_MENU: //hide sub menus for left menu
			$("li.topMenuItem", ".secondaryNav ol.rootNavMenu").each(function () {
				if ($currentLi && $(this).attr('id') === $currentLi.attr('id'))
					return true; //continue. don't hide current top menu item opened if provided
				self.hideSubMenu($(this));
			});
			break;
		case MAIN_MENU: //hide sub menus for main menu
			/*var openMenuID = self.$containerSubMenus.find('ol:first').attr('id');
			if ($currentLi && openMenuID === ('Parent' + $currentLi.attr('id')))
			return true; //continue. don't hide current top menu item opened if provided
			self.hideSubMenu(self.$containerSubMenus);*/
			$("li.topMenuItem", "#mainNavMenu").each(function () {
				if ($currentLi && $(this).attr('id') === $currentLi.attr('id'))
					return true; //continue. don't hide current top menu item opened if provided
				self.hideSubMenu($(this));
			});
			break;
		default: //hide submenus for all type of menus
			//self.hideSubMenu(self.$containerSubMenus);
			$("li.topMenuItem", "#mainNavMenu").each(function () {
				self.hideSubMenu($(this));
			});
			$("li.topMenuItem", "#sideNavMenu").each(function () {
				self.hideSubMenu($(this));
			});
			break;
	}
};

MenuManager.prototype.isSubMenuTrigger = function ($elem) {
	return $elem.closest(".rwd-menu").length === 0 && $elem.attr("aria-haspopup") === "true";
};

MenuManager.prototype.toggleSubMenu = function ($elem) {
	var expanded = $elem.attr("aria-expanded") === "true";
	var $li = $elem.closest("li");
	if (expanded) {
		menuManager.hideSubMenu($li);
	} else {
		menuManager.menuMouseOver($li);
	}
};

MenuManager.prototype.getFirstSubmenuItem = function ($elem) {
	var $parentLi = $elem.parent('li');
	var pageId = $parentLi.data('pageid');
	var displayMegaMenu = $parentLi.data('displaymegamenu');
	
	if (displayMegaMenu === 'True') {
		var $megaMenuContainer = $('.mainNavMegaMenu.megaMenuContainer.pageID_' + pageId);
		if ($megaMenuContainer.length > 0) {
			return $megaMenuContainer.find('a[role="menuitem"]').first();
		}
	}
	
	var firstAnchorId = "pagelink" + pageId;
	return $('#' + firstAnchorId);
};

MenuManager.prototype.setFocusToFirstSubmenuitem = function ($elem) {
	var $firstAnchor = menuManager.getFirstSubmenuItem($elem);
	if ($firstAnchor.length) {
		$firstAnchor[0].focus();
		return true;
	}
	return false;
};

MenuManager.prototype.setFocusToLastTopLevelItem = function () {
	// Close any open submenus when tabbing out, but do not move focus.
	// Let the browser preserve the natural tab order so Shift+Tab re-enters
	// the menubar on the first top-level item rather than the last focused item.
	if (accessibleMenuToggleEnabled) {
		menuManager.hideAllSubMenus(MAIN_MENU);
	}
};

MenuManager.prototype.getLastSubmenuItem = function ($elem) {
	var $parentLi = $elem.parent('li');
	var pageId = $parentLi.data('pageid');
	var displayMegaMenu = $parentLi.data('displaymegamenu');
	
	if (displayMegaMenu === 'True' || displayMegaMenu === true) {
		var $megaMenuContainer = $('.mainNavMegaMenu.megaMenuContainer.pageID_' + pageId);
		if ($megaMenuContainer.length > 0) {
			return $megaMenuContainer.find('a[role="menuitem"]').last();
		}
	}
	
	var parentOlId = $parentLi.attr('id');
	var $submenu = $('#Parent' + parentOlId);
	if ($submenu.length > 0) {
		return $submenu.find('a[role="menuitem"]').last();
	}
	return $();
};

MenuManager.prototype.setFocusToLastSubmenuitem = function ($elem) {
	var $lastAnchor = menuManager.getLastSubmenuItem($elem);
	if ($lastAnchor.length) {
		$lastAnchor[0].focus();
		return true;
	}
	return false;
};

MenuManager.prototype.toggleIndicator = function ($anchor) {
	if (window.FeatureToggles.isActive("CMS.Pages.AccessibleMainMenuWithIndicators")) {
		if ($anchor.attr('aria-expanded') === 'true') {
			if ($anchor.hasClass("arrowDown"))
				$anchor.removeClass("arrowDown").addClass("arrowUp");
			else if ($anchor.hasClass("arrowRight"))
				$anchor.removeClass("arrowRight").addClass("arrowDownSubMenu");
		} else {
			if ($anchor.hasClass("arrowUp"))
				$anchor.removeClass("arrowUp").addClass("arrowDown");
			else if ($anchor.hasClass("arrowDownSubMenu"))
				$anchor.removeClass("arrowDownSubMenu").addClass("arrowRight");
		}
	}
}

MenuManager.prototype.hideSubMenu = function ($li) {
	var self = this; // Reference to current object

	// As the sub-menu is closed, we let the screen reader know that the menu is no longer expanded.
	var $anchor = $li.find('a');
	if ($anchor.length && $anchor.attr('aria-haspopup') === 'true') {
		$anchor.attr('aria-expanded', 'false');
		self.toggleIndicator($anchor);
	}

	var displayMegaMenu = $li.data('displaymegamenu');
	var mainMenu = $li.parents('nav').hasClass('mainNav') | $li.parents('ol').hasClass('mainSubMenu');
	//When a menu item is set to display as mega menu, we still display regular sub menu for secondary nav, since it doesnt have a megamenu. So while hiding, hide the regular submenu for side navigation.
	if (displayMegaMenu != 'True' || !mainMenu) {
		if ($('#mainNavMenu').find('.navMenu').length == 0) {
			$('#mainNavMenu li a.mouseover').removeClass('mouseover');
		}
		var id = $li.attr('id');
		var $olSubMenu = $('#Parent' + id);
		if ($olSubMenu.length == 1) {
			//moves sub menu to hidden div
			$olSubMenu.appendTo(self.$divListSubMenus);
			//reset data of the menu
			$olSubMenu.removeData('menu');
			if (self.bouncing) {
				//remove bouncing classes
				$olSubMenu.removeClass('flipLeft flipRight');
			}
			$olSubMenu.css({ position: '', top: '', left: '', zIndex: '' });
			//moves recursively its children if needed
			$olSubMenu.children().each(function () {
				self.hideSubMenu($(this));
				$(this).unbind("mouseenter.menu");
			});

			//Remove animation trigger class
			$olSubMenu.removeClass('animation-triggered');
		}
	} else {
		var pageID = $li.data('pageid');
		var megaMenuContainer = $('.mainNavMegaMenu.megaMenuContainer.pageID_' + pageID);
		if (!megaMenuContainer.hasClass('hover') && $('.megaMenuContainer .handle.pin.pinned').length === 0) {
			//$('#megaMenu').hide();
			megaMenuContainer.hide();
			$('#mainNav').css('z-index', '');
			$('.megaMenuContainer').css("display", "none");
			$('.megaMenuContainer').removeClass('animation-triggered');
		}
		if (!$('.secondaryNavMegaMenu.megaMenuContainer.pageID_' + pageID).hasClass('hover')) {
			$('.secondaryNavMegaMenu.megaMenuContainer.pageID_' + pageID).hide();
		}
	}
};

//Supports legacy code to set delays when opening/close submenus
MenuManager.prototype.setMOMMode = function (mommode, menuType) {
	//support legacy code
	if (menuType === MAIN_MENU) {
		switch (mommode) {
			case 0:
				self.popdownDelayTopMenu = 3;
				self.popupDelayTopMenu = false;
				self.mainMenuMaxSubMenuLevels = 0;
				break;
			case 2:
				self.popdownDelayTopMenu = 500;
				self.popupDelayTopMenu = 500;
				break;
			case 3:
				self.popdownDelayTopMenu = 500;
				self.popupDelayTopMenu = false;
				break;
		}
	} else if (menuType === SIDE_MENU) {
		switch (mommode) {
			case 0:
				self.popdownDelaySideMenu = 3;
				self.popupDelaySideMenu = false;
				self.sideMenuMaxSubMenuLevels = 0;
				break;
			case 2:
				self.popdownDelaySideMenu = 500;
				self.popupDelaySideMenu = 500;
				break;
			case 3:
				self.popdownDelaySideMenu = 500;
				self.popupDelaySideMenu = false;
				break;
		}
	}
};

MenuManager.prototype.adjustMainItemsWidth = function (menuSelector) {
	this.isAdjustingMainMenuItemsWidth = true;
	return adjustMainItemsWidth(menuSelector);
};

MenuManager.prototype.isMegaMenuItem = function ($elem) {
	var $parent = $elem.parent('li');
	if ($parent != null && $parent.data('displaymegamenu') == 'True') {
		return true;
	}

	var $megaMenuContainer = $elem.closest('.mainNavMegaMenu');
	if ($megaMenuContainer.length > 0) {
		return true;
	}

	return false;
};

MenuManager.prototype.isFirstSubmenuItem = function ($elem) {
	var $submenu = $elem.closest('.mainSubMenu, .mainNavMegaMenu');
	if ($submenu.length === 0) return false;
	
	var $allItems = $submenu.find('a[role="menuitem"]');
	return $allItems.index($elem) === 0;
};

MenuManager.prototype.isLastSubmenuItem = function ($elem) {
	var $submenu = $elem.closest('.mainSubMenu, .mainNavMegaMenu');
	if ($submenu.length === 0) return false;
	
	var $allItems = $submenu.find('a[role="menuitem"]');
	return $allItems.index($elem) === $allItems.length - 1;
};

MenuManager.prototype.focusPrevSubmenuItem = function ($elem) {
	var $prev = $elem.parent('li').prev('li').find('a:first');
	if ($prev.length > 0) {
		$prev.focus();
	}
};

MenuManager.prototype.focusNextSubmenuItem = function ($elem) {
	var $next = $elem.parent('li').next('li').find('a:first');
	if ($next.length > 0) {
		$next.focus();
	}
};

MenuManager.prototype.findTopLevelParentId = function ($elem) {
	// For mega menus
	var $megaMenu = $elem.closest('.mainNavMegaMenu');
	if ($megaMenu.length > 0) {
		var pageId = $megaMenu.attr('id').replace('mainNavMegaMenu', '');
		return 'MainItem' + pageId;
	}

	// For regular submenus - find the submenu that contains the current element
	var $submenu = $elem.closest('ol[id^="ParentMainItem"]');
	if ($submenu.length === 0) {
		return null;
	}

	// Get the ID of the menu item that OWNS this submenu
	var ownerId = $submenu.attr('id').replace('ParentMainItem', ''); // e.g., "27" from "ParentMainItem27"

	// Recursive helper to find the top-level parent
	var findTopLevel = function (id) {
		var $item = document.querySelector('#MainItem' + id);
		if (!$item) return null;

		var $parentSubmenu = $item.closest("ol[id^='ParentMainItem']");

		if (!$parentSubmenu) {
			// This item is not inside any submenu, so it's top-level
			return 'MainItem' + id;
		}

		// Get the parent's owner ID and recurse
		var parentOwnerId = $parentSubmenu.id.match(/\d+$/)?.[0];
		return parentOwnerId ? findTopLevel(parentOwnerId) : 'MainItem' + id;
	};

	return findTopLevel(ownerId);
};

MenuManager.prototype.getSubmenuParentId = function ($elem) {
	var $submenu = $elem.closest('.mainSubMenu');
	if ($submenu.length > 0) {
		// For regular submenus, extract parent ID from submenu's id (e.g., "ParentMainItem123")
		return $submenu.attr('id').replace('Parent', '');
	}
	
	var $megaMenu = $elem.closest('.mainNavMegaMenu');
	if ($megaMenu.length > 0) {
		// For mega menus, extract from mega menu container id (e.g., "mainNavMegaMenu123")
		return 'MainItem' + $megaMenu.attr('id').replace('mainNavMegaMenu', '');
	}
	
	return null;
};

MenuManager.prototype.handleUpKeyOnTopLevelItem = function ($elem) {
	if (!accessibleMenuToggleEnabled || !this.isSubMenuTrigger($elem)) {
		return false;
	}
	
	var isExpanded = $elem.attr('aria-expanded') === 'true';
	if (!isExpanded) {
		this.toggleSubMenu($elem);
	}
	
	var self = this;
	setTimeout(function () {
		self.setFocusToLastSubmenuitem($elem);
	}, isExpanded ? 0 : 300);
	
	return true;
};

MenuManager.prototype.handleDownKeyOnTopLevelItem = function ($elem) {
	if (!accessibleMenuToggleEnabled || !this.isSubMenuTrigger($elem)) {
		return false;
	}
	
	var isExpanded = $elem.attr('aria-expanded') === 'true';
	if (!isExpanded) {
		this.toggleSubMenu($elem);
	}
	
	var self = this;
	setTimeout(function () {
		self.setFocusToFirstSubmenuitem($elem);
	}, isExpanded ? 0 : 300);
	
	return true;
};

MenuManager.prototype.handleUpKeyInSubmenu = function ($elem, options) {
	options = options || {};

	// Existing behavior for non-toggle mode	
	if (!accessibleMenuToggleEnabled) {
		if (this.isFirstSubmenuItem($elem) && options.onFirstItem) {
			options.onFirstItem($elem);
		} else {
			this.focusPrevSubmenuItem($elem);
		}
		return;
	}

	// New behavior with toggle enabled
	if (this.isFirstSubmenuItem($elem)) {
		var parentId = this.getSubmenuParentId($elem);
		if (parentId) {
			this.setFocusToLastSubmenuitem($('#' + parentId).find('a:first'));
		}
	} else {
		this.focusPrevSubmenuItem($elem);
	}
};

MenuManager.prototype.handleDownKeyInSubmenu = function ($elem) {
	// Existing behavior for non-toggle mode	
	if (!accessibleMenuToggleEnabled) {
		this.focusNextSubmenuItem($elem);
		return;
	}

	// New behavior with toggle enabled
	if (this.isLastSubmenuItem($elem)) {
		var parentId = this.getSubmenuParentId($elem);
		if (parentId) {
			this.setFocusToFirstSubmenuitem($('#' + parentId).find('a:first'));
		}
	} else {
		this.focusNextSubmenuItem($elem);
	}
};

MenuManager.prototype.getAdjacentTopLevelItem = function ($currentLi, direction) {
	var $targetLi = direction === 'prev' ? $currentLi.prev('li.topMenuItem') : $currentLi.next('li.topMenuItem');

	if ($targetLi.length === 0) {
		var $siblings = $currentLi.parent().children('li.topMenuItem');
		$targetLi = direction === 'prev' ? $siblings.last() : $siblings.first();
	}

	return $targetLi;
};

MenuManager.prototype.handleLeftKeyOnTopLevelItem = function ($anchor, openSubmenu) {
	var $currentLi = $anchor.parent('li.topMenuItem');
	var $prevLi = this.getAdjacentTopLevelItem($currentLi, 'prev');
	var $targetAnchor = $prevLi.find('a:first');
	$targetAnchor.focus();

	if (openSubmenu && this.isSubMenuTrigger($targetAnchor)) {
		clearTimeout(this.menuTimeoutID);
		this.openSubmenuImmediately($prevLi);
	}
};

MenuManager.prototype.handleRightKeyOnTopLevelItem = function ($anchor, openSubmenu) {
	var $currentLi = $anchor.parent('li.topMenuItem');
	var $nextLi = this.getAdjacentTopLevelItem($currentLi, 'next');
	var $targetAnchor = $nextLi.find('a:first');
	$targetAnchor.focus();

	if (openSubmenu && this.isSubMenuTrigger($targetAnchor)) {
		clearTimeout(this.menuTimeoutID);
		this.openSubmenuImmediately($nextLi);
	}
};

MenuManager.prototype.handleLeftKeyInSubmenu = function ($elem, shouldOpenSubmenu) {
	var self = this;
	// Check if current item has a parent submenu item
	var $currentSubmenu = $elem.closest('ol.mainSubMenu');
	if ($currentSubmenu.length === 0) {
		return;
	}

	var parentItemId = $currentSubmenu.attr('id').replace('Parent', '');
	var $parentItem = $('#' + parentItemId);
	var $parentAnchor = $parentItem.find('a:first');

	// Check if the parent item is a submenu item
	var $parentSubmenu = $parentItem.closest('ol.mainSubMenu');
	var parentIsSubmenuItem = $parentSubmenu.length > 0;

	if (parentIsSubmenuItem) {
		this.hideSubMenu($parentItem);
		$parentAnchor.focus();
		return;
	}

	var parentId = this.getSubmenuParentId($elem);
	if (!parentId)
		return;

	var $parentLi = $('#' + parentId);

	this.hideSubMenu($parentLi);

	var $prevLi = $parentLi.prev('li.topMenuItem');
	if ($prevLi.length === 0) {
		$prevLi = $parentLi.parent().children('li.topMenuItem').last();
	}

	var $prevAnchor = $prevLi.find('a:first');
	$prevAnchor.focus();

	if (shouldOpenSubmenu && this.isSubMenuTrigger($prevAnchor)) {
		this.openSubmenuImmediately($prevLi);
		var isExpanded = $prevAnchor.attr('aria-expanded') === 'true';

		if (isExpanded) {
			setTimeout(function () {
				self.setFocusToFirstSubmenuitem($prevAnchor);
			}, 300);
		}
	}
};

MenuManager.prototype.handleRightKeyInSubmenu = function ($elem, shouldOpenSubmenu) {
	var self = this;

	// Check if current item has child submenu
	if (this.isSubMenuTrigger($elem)) {
		var isExpanded = $elem.attr('aria-expanded') === 'true';
		if (!isExpanded) {
			var $parentLi = $elem.parent('li');
			this.openSubmenuImmediately($parentLi);
			setTimeout(function () {
				self.setFocusToFirstSubmenuitem($elem);
			}, 300);
			return;
		}
	}

	var topLevelParentId = this.findTopLevelParentId($elem);
	if (!topLevelParentId) {
		return;
	}

	this.hideAllSubMenus(MAIN_MENU);

	var $topLevelParent = $('#' + topLevelParentId);
	var $nextLi = $topLevelParent.next('li.topMenuItem');
	if ($nextLi.length === 0) {
		$nextLi = $('#mainNavMenu > li.topMenuItem').first();
	}

	var $nextAnchor = $nextLi.children('a:first');
	if ($nextAnchor.length > 0) {
		$nextAnchor.focus();
		if (shouldOpenSubmenu) {
			// Check if target is a mega menu OR regular submenu
			var isMegaMenu = $nextLi.attr('data-displaymegamenu') === 'True';
			var hasSubmenu = this.isSubMenuTrigger($nextAnchor);

			if (isMegaMenu || hasSubmenu) {
				self.menuMouseOver($nextLi);

				setTimeout(function () {
					self.setFocusToFirstSubmenuitem($nextAnchor);
				}, 300);
			}
		}
	}
};

MenuManager.prototype.openSubmenuImmediately = function ($li) {
	var self = this;
	var savedDelay = self.popupDelayTopMenu;
	self.popupDelayTopMenu = 0; // Disable delay temporarily
	self.menuMouseOver($li);
	self.popupDelayTopMenu = savedDelay; // Restore delay
};

MenuManager.prototype.shouldOpenSubmenuOnNavigation = function ($elem) {
	// Find if current element is in a submenu
	var $currentSubmenu = $elem.closest('ol.mainSubMenu');
	if ($currentSubmenu.length === 0) {
		return false;
	}

	// Get the parent menu item of this submenu
	var parentItemId = $currentSubmenu.attr('id').replace('Parent', '');
	var $parentItem = $('#' + parentItemId);

	// Check if parent is a top-level item or nested submenu item
	var $parentSubmenu = $parentItem.closest('ol.mainSubMenu');

	if ($parentSubmenu.length === 0) {
		// Parent is top-level,so we should open next submenu
		return true;
	} else {
		// Parent is in a nested submenu, check its aria-expanded
		var $parentAnchor = $parentItem.find('a:first');
		return $parentAnchor.attr('aria-expanded') === 'true';
	}
};

MenuManager.prototype.isTopLevelMenuItem = function ($elem) {
	return $elem.hasClass('navMainItem') || $elem.closest('li').hasClass('topMenuItem');
};

MenuManager.prototype.setActiveTopLevelItem = function ($elem) {
	if (!accessibleMenuToggleEnabled || !this.isTopLevelMenuItem($elem)) {
		return;
	}

	// Roving tabindex pattern
	let $topLevelAnchor = $elem.hasClass('navMainItem') ? $elem : $elem.closest('li.topMenuItem').find('a.navMainItem').first();
	if ($topLevelAnchor.length === 0) {
		return;
	}
	
	$('#mainNavMenu > li.topMenuItem > a.navMainItem').attr('tabindex', '-1');
	$topLevelAnchor.attr('tabindex', '0');
};

MenuManager.prototype.handleUpKeyOnAnyTopLevelItem = function ($elem) {
	if (this.isMegaMenuItem($elem)) {
		return false;
	}

	return this.handleUpKeyOnTopLevelItem($elem);
};

MenuManager.prototype.handleLeftKeyOnAnyTopLevelItem = function ($elem) {
	if (accessibleMenuToggleEnabled) {
		var shouldOpenSubmenu = $elem.attr('aria-expanded') === 'true';
		this.handleLeftKeyOnTopLevelItem($elem, shouldOpenSubmenu);
		return;
	}

	this.navigateToSiblingTopLevelItem($elem, 'prev');
};

MenuManager.prototype.handleRightKeyOnAnyTopLevelItem = function ($elem) {
	if (accessibleMenuToggleEnabled) {
		var shouldOpenSubmenu = $elem.attr('aria-expanded') === 'true';
		this.handleRightKeyOnTopLevelItem($elem, shouldOpenSubmenu);
		return;
	}

	this.navigateToSiblingTopLevelItem($elem, 'next');
};

MenuManager.prototype.navigateToSiblingTopLevelItem = function ($elem, direction) {
	var $currentLi = $elem.parent('li.topMenuItem');
	var $targetLi = this.getAdjacentTopLevelItem($currentLi, direction);
	$targetLi.find('a:first').focus();
};

function adjustMainItemsWidth(menuSelector) {
	var widths = [];
	var $items = [];
	$(menuSelector).data('autowidth', true).find(".topMenuItem").each(function () {
		widths.push($(this).outerWidth());
		$items.push($(this));
	});
	if (widths.length > 0) {
		var totalWidth = 0;
		var i;
		for (i = 0; i < widths.length; i++) {
			totalWidth += widths[i];
		}
		var menuWidth = $(menuSelector).outerWidth();
		var availableWidth = menuWidth - totalWidth;
		var paddingPerItem = availableWidth / $items.length;
		for (i = 0; i < $items.length; i++) {
			var width = (100 * (widths[i] + paddingPerItem) / menuWidth);
			//This is to account for the anomoly of the last menu item wrapping for one client (Anoka County) only in Mac Safari
			//A site-specific solution was sought but none found that was acceptable
			if (navigator.userAgent.indexOf("Safari") > -1 && navigator.userAgent.indexOf('Chrome') == -1 && i == $items.length - 1) {
				width = width - 1;
			}
			$items[i].outerWidth(width + "%");
		}
	}
}

MenuManager.prototype.resetMainItemsWidth = function (menuSelector) {
	$(menuSelector).find(".topMenuItem").each(function () {
		$(this).css("width", "auto");
	});
};

function log(text) {
	console.log(text);
}

/**** End-Functions to handle Main/Left Menu on Front-end. *****/

/***** Drag and Drop functions to handle main/side menus  ********/

//Handles when an element was dropped
MenuManager.prototype.dragHandler = function ($element, dragInfo) {
	var self = this; // Reference to current object

	//var miLeave = 0;
	var menuItemLeave = function ($menuItemOver) {
		$menuItemOver.removeClass('hover');
		//miLeave++;
		//log("Leave event " + (miLeave) + ", liID=" + $menuItemOver.attr('id') + ", class=" + $menuItemOver.attr('class'));
		//$('#log3').val("Leave event " + (miLeave) + ", liID=" + $menuItemOver.attr('id') + ", class=" + $menuItemOver.attr('class'));
	};

	if ($element.length != 0 && !$element.hasClass('thingToDrag')) { //workaround to avoid executing drag handler when mouse pointer is over drag helper
		//resetElementOver();
		//dragInfo.currentElementOver = { $element: null, $li: null, isAddPage: null, $ol: null };
		dragInfo.resetCurrentElementOver();
		// This element will hold info to keep track of the current element that the mouse pointer is over.
		dragInfo.currentElementOver.$element = $element;
		//$('#log1').val('element  ID:' + $element.attr('id') + ', class=' + $element.attr('class') + ", tag=" + $element.get(0).nodeName + "html=" + $element.html());

		if ($element.hasClass('inside')) {
			//when mouse is over any DOM element inside an LI MenuItem
			var $li = $element.closest('li');

			dragInfo.currentElementOver.$li = $li;
		} else if ($element.hasClass('liDropabble')) {
			//when mouse is over LI form field element
			dragInfo.currentElementOver.$li = $element;
		}

		if (dragInfo.currentElementOver.$li != null && dragInfo.currentElementOver.$li.hasClass('liDropabble')) {
			//Over a valid droppable li
			dragInfo.currentElementOver.$ol = dragInfo.currentElementOver.$li.parent();
			//$('#log2').val("inside.. . currentLI=" + dragInfo.currentElementOver.$li.attr('id'));
		}

		// Logic to determines if mouse pointer is out of a menu item.
		if (dragInfo.$prevMenuItemOver != null && (dragInfo.currentElementOver.$li == null || dragInfo.currentElementOver.$li[0] != dragInfo.$prevMenuItemOver[0])) {
			//fire menu item leave event
			menuItemLeave(dragInfo.$prevMenuItemOver);
		}
		dragInfo.$prevMenuItemOver = dragInfo.currentElementOver.$li;

		//mouse is over a valid form field to resize.??
	}
};

///Handles when an element was dropped
MenuManager.prototype.dropHandler = function (dragInfo) {
	var self = this; // Reference to current object

	var validPosition = false;
	if (dragInfo.currentElementOver != null && dragInfo.currentElementOver.$li != null && dragInfo.currentElementOver.$li.hasClass('liDropabble') && !dragInfo.currentElementOver.$li.hasClass('beingDragged')) {
		var $liTarget = dragInfo.currentElementOver.$li;
		$liTarget.removeClass('hover');

		if (self.sideMenuHasCustomLinks || dragInfo.$draggedElement.data('drag').menuType === $liTarget.data('drag').menuType) {
			validPosition = true;
			//var msgSource = "ID=" + dragInfo.$draggedElement.data('drag').itemID + ",parentID=" + dragInfo.$draggedElement.data('drag').parentID + ",prevItemID=" + dragInfo.$draggedElement.data('drag').prevItemID + ",isAddPage=" + dragInfo.$draggedElement.data('drag').isAddPage;
			//var msgTarget = "ID=" + $liTarget.data('drag').itemID + ",parentID=" + $liTarget.data('drag').parentID + ",prevItemID=" + $liTarget.data('drag').prevItemID + ",isAddPage=" + $liTarget.data('drag').isAddPage;
			//alert("dropped source=" + msgSource + "\nOn target=" + msgTarget);
			var url = '';
			var addingPage = false;
			if (dragInfo.$draggedElement.data('drag').isAddPage) {
				//Creating a new page in a specific menu position
				addingPage = true;
				url = "/Pages/Page/Add";
			} else {
				//Moving an existing menu item on different position
				url = "/Pages/PageStatus/Move/" + dragInfo.$draggedElement.data('drag').itemID;
			}
			//logic to build position where menu item was dropped
			url += "?parentID=" + $liTarget.data('drag').parentID;
			if ($liTarget.data('drag').isAddPage) {
				if ($liTarget.data('drag').prevItemID != null)
					url += "&refPageID={1}&position=Next".replace("{1}", $liTarget.data('drag').prevItemID);
				else
					url += "&position=Child";
			} else {
				url += "&refPageID={1}&position=Previous".replace("{1}", $liTarget.data('drag').itemID);
				if ($liTarget.hasClass('topMenuItem') && $liTarget.data('drag').menuType == MAIN_MENU) {
					//Check whether or not is possible to copy layout from top sibling page
					var message = '';
					$.ajax({
						url: '/Layout/CopyAssociation/CheckTopPages?copyLayoutFrom={0}'
							.replace('{0}', $liTarget.data('drag').itemID),
						type: 'POST',
						async: false, //Needs to be a sync request
						success: function (response) {
							if (response.ErrorMessage) {
								alert(response.ErrorMessage);
							} else if (response.Message) {
								message = response.Message;
							}
						},
						error: function (xhr, textStatus, exception) {
							alert("Error: " + xhr.statusText + "\nStatus: " + xhr.status);
						}
					});
					if (message.length > 0 && confirm(message)) {
						url += "&copyLayoutFrom=" + $liTarget.data('drag').itemID;
					}
				}
			}

			if (addingPage) {
				window.location.href = url;
			} else {
				$.ajax({
					url: url,
					type: 'POST',
					success: function (response) {
						if (response.ErrorMessage) {
							alert(response.ErrorMessage);
							ajaxPostBackEnd();
						} else if (response.RedirectURL && response.RedirectURL !== location.href) {
							location.href = response.RedirectURL;
						} else {
							location.reload(true);
						}
					},
					beforeSend: function () {
						ajaxPostBackStart('Loading');
					},
					error: function (xhr, textStatus, exception) {
						alert("Error: " + xhr.statusText + "\nStatus: " + xhr.status);
						ajaxPostBackEnd();
					}
				});
			}
		}
	}
};

MenuManager.prototype.setupDraggableElement = function ($element, menuType, containerSelectorToAppend) {
	var self = this; // Reference to current object

	//Unbind the event before binding it again. Both solutions below work.
	$element.unbind('draginit dragstart drag dragend');
	//$element.drag('destroy');
	// Prevents IE from selecting text while starting a drag.
	$element.bind('selectstart', function (event) {
		event.preventDefault();
		return false;
	});

	var dragInfo = null;

	var relativeContainerOffSet = { top: 0, left: 0 };
	var draggedRelativeElementOffset = { top: 0, left: 0 };
	var isAddPageElem = $element.hasClass('addPage');

	//jquery event drag 2.0 plugin.
	$element.drag("start", function () {
		if (self.mobileMainNav) return false;
		self.isMainMenuBeingDragged = (menuType == MAIN_MENU);
		self.isSideMenuBeingDragged = (menuType == SIDE_MENU);

		var $draggedElement = null;
		if ($(this).hasClass('grippy'))
			$draggedElement = $(this).parent('li'); //page menu items
		else if ($(this).hasClass('addPage'))
			$draggedElement = $(this); //Add Page menu items

		$draggedElement.addClass('beingDragged');

		/*if ($draggedElement.hasClass('topMenuItem') && (menuType == MAIN_MENU)) {
		//workaround to hide top level menu items on main menu. (hideMenu don't work since submenus are not naturally on DOM)
		self.hideAllSubMenus(MAIN_MENU);
		}
		else*/
		self.hideSubMenu($draggedElement);

		dragInfo = {
			$draggedElement: $draggedElement,
			currentElementOver: { $element: null, $li: null, isAddPage: null, $ol: null },
			$prevMenuItemOver: null,
			resetCurrentElementOver: function () {
				this.currentElementOver.$element = null;
				this.currentElementOver.$li = null;
				this.currentElementOver.$ol = null;
				this.currentElementOver.isAddPage = null;
			}
		};

		//If container for thing to drag has relative position, gets its offset to subtract coordinates while dragging.
		relativeContainerOffSet = ($(containerSelectorToAppend).css("position") == "relative") ? $(containerSelectorToAppend).offset() : { top: 0, left: 0 };

		var hasChildren = $draggedElement.hasClass('hasChildren');
		var menuText = $draggedElement.children('a.navMenuItem, a.navMainItem').text();
		if (menuText == '') {
			menuText = $draggedElement.children('div.accordionNavItem').children('a.navMenuItem').text();
		}
		var $dragHelper = $('<a/>', {
			'class': 'thingToDrag navMenuItem' + (hasChildren ? ' withChildren' : '')
		}).css('z-index', 1002).text(menuText);
		return $dragHelper.appendTo(containerSelectorToAppend);
	}).drag(function (e, dd) {
		//if draggable element has relative position, gets the offset where the mouse was clicked to add coordinates while dragging.
		draggedRelativeElementOffset = (isAddPageElem) ? { top: dd.startY - dd.originalY, left: dd.startX - dd.originalX } : { top: 0, left: 0 };

		$(dd.proxy).css({
			top: (dd.offsetY + draggedRelativeElementOffset.top - relativeContainerOffSet.top) + 15,
			left: (dd.offsetX + draggedRelativeElementOffset.left - relativeContainerOffSet.left) - 10
		});
		//$('#log4').val('dd.offsetX=' + dd.offsetX + ',dd.offsetY=' + dd.offsetY + ',e.clientX=' + e.clientX + ',e.clientY=' + e.clientY + ',e.pageX=' + e.pageX + ',e.pageY=' + e.pageY);
		var $elementOver = $(document.elementFromPoint(e.originalEvent.clientX, e.originalEvent.clientY));
		// Execute drag handler
		if (dragInfo != null)
			self.dragHandler($elementOver, dragInfo);
	}, { relative: false })
		.drag("end", function (ev, dd) {
			if (self.mobileMainNav) return false;
			$(dd.proxy).remove();

			if (dragInfo != null) {
				self.dropHandler(dragInfo);
				dragInfo.$draggedElement.removeClass('beingDragged');
				dragInfo.$draggedElement = null;
				dragInfo.$prevMenuItemOver = null;
				dragInfo = null;
			}
			self.isMainMenuBeingDragged = false;
			self.isSideMenuBeingDragged = false;
		});
};

MenuManager.prototype.setupDroppableElement = function ($element, menuType) {
	var self = this; // Reference to current object
	if (!self.mobileMainNav && $element.find('div.indicator').length == 1 && !$element.hasClass('liDropabble')) {
		$element.addClass('liDropabble');

		//initialize drag jquery data for droppable elements
		var isAddPage = $element.hasClass('addPage');
		var itemID = !isAddPage ? $element.attr('id').replace(/MainItem|SideItem/i, '') : '';
		var prevItemID = ($element.prev().length == 1) ? $element.prev().attr('id').replace(/MainItem|SideItem/i, '') : null;
		var parentID = $element.parent().attr('id').replace(/ParentMainItem|ParentSideItem/i, '');
		if (isNaN(parentID)) //for top level menu items
			parentID = $element.parent().data('parent');

		$element.data('drag', { itemID: itemID, parentID: parentID, prevItemID: prevItemID, isAddPage: isAddPage, menuType: menuType });

		/// Setups recursively all children DOM elements with class needed by drag algorithm to recognize which element mouse pointer is
		var setUpChildren = function ($children) {
			if ($children.length > 0) {
				$children.each(function () {
					$children.addClass('inside');
					setUpChildren($(this).children());
				});
			}
		};

		setUpChildren($element.children());
	}
};

MenuManager.prototype.setupDroppableAccordionElement = function ($element) {
	var self = this;
	if (!self.mobileMainNav && $element.find('div.indicator').length > 0 && !$element.hasClass('liDropabble')) {
		$element.addClass('liDropabble');

		//initialize drag jquery data for droppable elements
		var isAddPage = $element.hasClass('addPage');
		var itemID = !isAddPage ? $element.attr('id').replace(/MainItem|SideItem/i, '') : '';
		var prevItemID = ($element.prev().length == 1) ? $element.prev().attr('id').replace(/MainItem|SideItem/i, '') : null;
		var $parent = $element.parents('li:first');
		if ($parent.length == 0) {
			$parent = $element.parents('ol:first');
		}
		var parentID = $parent.attr('id').replace(/ParentMainItem|ParentSideItem/i, '');
		if (isNaN(parentID)) //for top level menu items
			parentID = $element.parent().data('parent');

		$element.data('drag', { itemID: itemID, parentID: parentID, prevItemID: prevItemID, isAddPage: isAddPage, menuType: SIDE_MENU });

		/// Setups recursively all children DOM elements with class needed by drag algorithm to recognize which element mouse pointer is
		var setUpChildren = function ($children) {
			$children.each(function () {
				$children.addClass('inside');
				setUpChildren($(this).children());
			});
		};

		setUpChildren($element.children());
	}

	$element.unbind('mouseenter').mouseenter(function (e) {
		e.stopPropagation();
		if (self.isSideMenuBeingDragged) {
			$element.addClass('hover');
		}
	});

	$element.unbind('mouseleave').mouseleave(function (e) {
		$element.removeClass('hover');
	});
};

///Moves the submenu to be displayed at the bottom of the DOM.
MenuManager.prototype.moveSubMenuInDOM = function ($submenu, menuType) {

	//Remove any old .pop elements from the DOM to reload the new one.
	if ($('body').find($submenu).length > 0)
		$('body').find($submenu).remove();

	//$submenu is the ol, $parent is the li.
	$submenu.addClass('subMenuMoved').appendTo('body');

	var className = menuType == MAIN_MENU ? 'mainSubMenu' : 'secondarySubMenu';
	$submenu.addClass(className);
	$submenu.addClass('moved');

	$submenu.css('position', 'absolute');
	$submenu.css('z-index', '100');
}

///Updates the position of submenus in the DOM.
MenuManager.prototype.positionSubMenuInDOM = function ($submenu, $li, menuType) {
	var topMainNavigationLink = menuType == MAIN_MENU && $li.hasClass('topMenuItem');
	var self = this;
	self.updateSubMenuPositionAttributes($submenu, topMainNavigationLink, menuType);
	//If after the submenu is positioned, it has a class of flipLeft, remove it since that causes some display issues on the front-end. Vishal
	if ($submenu.hasClass('flipLeft'))
		$submenu.removeClass('flipLeft').addClass('flippedLeft');
};

MenuManager.prototype.updateSubMenuPositionAttributes = function ($submenu, topMainNavigationLink, menuType) {
	//Get the main nav.
	var parentID = $submenu.attr('id').replace(/Parent/, '');
	var $parent = $('#' + parentID);

	$submenu.removeClass('flippedLeft');

	if ($('#bodyWrapper').data('maxPaddingBottom') == undefined)
		$('#bodyWrapper').data('maxPaddingBottom', 0);

	//Reset bottom-padding on body wrapper to allow accurate calculation
	$('#bodyWrapper').css('padding-bottom', $('#bodyWrapper').data('maxPaddingBottom'));

	var docHeight = $(document).height();
	var menuVisibleSpaceAvailable = docHeight - $parent.offset().top;
	var submenuHeight = $submenu.height();

	//If the parent submenu opened in towards left, let all the subsequent submenus open in the same flow.
	if ($parent.length > 0 && $parent.parents('ol.subMenuMoved').hasClass('flippedLeft'))
		$submenu.addClass('flipLeft');

	if (menuVisibleSpaceAvailable < submenuHeight && !$submenu.hasClass('takenIntoAccount')) {
		//20 is just a default safe padding apart from the calculations.
		$('#bodyWrapper').css('padding-bottom', (submenuHeight - menuVisibleSpaceAvailable + parseFloat($('#bodyWrapper').data('maxPaddingBottom')) + 20));
		$('#bodyWrapper').data('maxPaddingBottom', $('#bodyWrapper').css('padding-bottom'));
		$submenu.addClass('takenIntoAccount');
		if (!$submenu.hasClass('safePaddingApplied')) {
			//Add a padding-bottom of 10px to the submenu to be more presentable
			$submenu.css('padding-bottom', '10px');
			$submenu.addClass('safePaddingApplied');
		}
	}
	if (menuType == MAIN_MENU) {
		//Set top position of main submenu
		$submenu.css('top', $parent.offset().top + (topMainNavigationLink ? $parent.height() : 0));

		if ($submenu.hasClass('flipLeft')) {
			var width = $submenu.find('li:first').width();
			//Set left position of side submenu
			$submenu.css('left', $parent.offset().left - width);
		}
		else {
			//Set left position of main submenu
			$submenu.css('left', $parent.offset().left + (topMainNavigationLink ? 0 : $parent.width()));
		}
	}
	else {
		//Set top position of side submenu
		$submenu.css('top', $parent.offset().top);

		if ($submenu.hasClass('flipLeft')) {
			var width = $submenu.find('li:first').width();
			//Set left position of side submenu
			$submenu.css('left', $parent.offset().left - width);
		}
		else {
			//Set left position of side submenu
			$submenu.css('left', $parent.offset().left + $parent.width());
		}
	}
};
/***** End - Drag and Drop functions to handle main/side menus  ********/

/***** Start - Mega Menu methods  ********/
MenuManager.prototype.pinMegaMenuOpen = function ($this) {
	var $megaMenuGrid = $this.parents('.megaMenuContainer').children('.pageContent.cpGrid');
	if ($megaMenuGrid.is(':visible')) {
		$('#mainNav').css('z-index', '10');
		$megaMenuGrid.addClass('pinned');
		$megaMenuGrid.find('.handle.pin').addClass('pinned');
		if (typeof PageEditLocalization != 'undefined') {
			$megaMenuGrid.find('.handle.pin').attr('title', pageEditLocalization.UnpinMegaMenu);
		}
	}
};

MenuManager.prototype.resetMouseOvers = function () {
	$('#mainNav').css('z-index', '').find('*').removeClass('mouseover');
	$('#secondaryNav').find('*').removeClass('mouseover');
};

/// Setup
MenuManager.prototype.initializeMegaMenus = function () {
	var self = this;
	if (typeof window.DesignCenter === 'undefined' || typeof window.DesignCenter.themeJSON === 'undefined') { // If we aren't in design center....
		$('#megaMenu,.megaMenuContainer').unbind('mouseenter').mouseenter(function (e) {
			e.stopPropagation();
			var $megaMenuContainer = $(this).attr('id') == 'megaMenu' ? $('.megaMenuContainer:visible') : $(this);
			var pageID = $megaMenuContainer.attr('id').match(/\d+$/)[0];
			$('#mainNav [data-pageid=' + pageID + ']  a').addClass('mouseover');

			$megaMenuContainer.addClass('hover');
		});

		$('.megaMenuContainer').unbind('mouseleave').mouseleave(function (e) {
			var $megaMenuContainer = $(this);
			var visible = $(this).is(':visible');
			if (visible) {
				if ($megaMenuContainer.hasClass('optionOpen') && !$('.cpPopOver.common').is(':visible')) {
					$megaMenuContainer.removeClass('optionOpen');
				}
				if (!$megaMenuContainer.hasClass('handleHover') && !$megaMenuContainer.hasClass('optionOpen') && !$megaMenuContainer.find('.pageContent.cpGrid').hasClass('dragging') && !$megaMenuContainer.find('.pageContent.cpGrid').hasClass('pinned')) {
					$megaMenuContainer.hide().removeClass('hover');
					//$('#megaMenu').hide();
					$('.megaMenu').hide();

					//While mousing out of mega menu, if it contains any flyout widgets, hide them.
					$('.flyOut.moved.forMegaMenu').hide();

					self.resetMouseOvers();
				}
			}
		});

		$('.megaMenuContainer .handle.remove, .megaMenuContainer .handle.move, .megaMenuContainer .handle.inheritance').unbind('mouseenter').mouseenter(function (e) {
			e.stopPropagation();
			$(this).parents('.megaMenuContainer').addClass('handleHover');
		});

		$('.megaMenuContainer .handle.remove, .megaMenuContainer .handle.move, .megaMenuContainer .handle.inheritance').unbind('mouseleave').mouseleave(function (e) {
			$(this).parents('.megaMenuContainer').removeClass('handleHover');
		});

		$('.megaMenuContainer .handle.pin').unbind('click').click(function (e) {
			$(this).toggleClass('pinned');
			$(this).parent().toggleClass('pinned');
			if (typeof PageEditLocalization != 'undefined') {
				$(this).attr('title', $(this).hasClass('pinned') ? pageEditLocalization.UnpinMegaMenu : pageEditLocalization.PinMegaMenu);
			}
		});
	}
};

MenuManager.prototype.attachDragAndDropHandlers = function () {
	//Make mega menu items draggable
	var $currentDraggables = $('.megaMenuContainer').find('li.level1.ui-draggable, li.level2.ui-draggable');
	if ($currentDraggables.length) {
		$currentDraggables.draggable('destroy');
	}

	$('.megaMenuContainer').find('li.level1, li.level2').draggable({
		helper: function (event) {
			var html = '<ol class="semeanticList" style="white-space: nowrap">';
			html += $('<div>').append($(this).clone()).html() + '</ol>';
			return $(html);
		},
		start: function (event, ui) {
			$('.megaMenuContainer .pageContent').addClass('megaMenudragging');
		},
		stop: function (event, ui) {
			$('.megaMenuContainer .pageContent').removeClass('megaMenudragging');
		},
		cursorAt: { top: 5, right: 5 },
		cursor: 'move',
		handle: '.grippy'
	});

	//Attach drop target handlers
	var $currentDroppables = $('.megaMenuTarget.ui-droppable', '.megaMenuContainer.ui-droppable');
	if ($currentDroppables.length) {
		$currentDroppables.droppable('destroy');
	}
	$('.megaMenuTarget', '.megaMenuContainer').droppable({
		over: function (event, ui) {
			$(this).parent().addClass('hover');
		},
		out: function (event, ui) {
			$(this).parent().removeClass('hover');
		},
		drop: function (event, ui) {
			var $target = $(this),
				targetPageID = $target.data('pageid'),
				targetPosition = $target.data('position'),
				parentID = $target.data('parentid'),
				pageID = $(ui.helper).children('li').data('pageid');

			$target.parent().removeClass('hover');

			var url = "/Pages/PageStatus/Move/" + pageID;
			url += "?parentID=" + parentID;
			url += "&refPageID={0}&position={1}".replace("{0}", targetPageID).replace("{1}", targetPosition);

			$.ajax({
				url: url,
				type: 'POST',
				success: function (response) {
					if (response.ErrorMessage) {
						alert(response.ErrorMessage);
						ajaxPostBackEnd();
					} else {
						location.href = response.RedirectURL;
					}
				},
				beforeSend: function () {
					ajaxPostBackStart('Loading');
				},
				error: function (xhr, textStatus, exception) {
					alert("Error: " + xhr.statusText + "\nStatus: " + xhr.status);
					ajaxPostBackEnd();
				}
			});
		},
		tolerance: 'pointer'
	});
};

/***** End - Mega Menu methods  ********/

/***** Start - Class/functions to handle custom links ********/
function MenuCustomLinksManager(confirmCustomLinksDeletion, pageID, moduleID) {
	this.pageID = pageID;
	this.moduleID = moduleID;
	this.checkOrphanFeatureLinks = function () {
		if (confirmCustomLinksDeletion && confirmCustomLinksDeletion.length > 0) {
			return confirm(confirmCustomLinksDeletion);
		}
		return true;
	};

	this.updateSideMenu = function (actionUrl) {
		$.ajax({
			url: actionUrl,
			type: 'post',
			cache: false,
			success: function (response) {
				if (response.ErrorMessage) {
					alert(response.ErrorMessage);
					ajaxPostBackEnd();
				} else if (response.RedirectURL)
					window.location.href = response.RedirectURL;
				else {
					//update side menu UI.
					window.location.reload(true);
				}
			},
			beforeSend: function () {
				ajaxPostBackStart('Loading');
			},
			error: function (xhr, textStatus, exception) {
				ajaxPostBackEnd();
				alert("Error choosing default side menu: " + xhr.statusText + "\nStatus: " + xhr.status);
			}
		});
	};
}

/// open modal to chose page to associate with current page/module as custom link
MenuCustomLinksManager.prototype.chosePage = function () {
	var self = this; // Reference to current object
	if (self.checkOrphanFeatureLinks()) {
		openCpModal({
			title: 'Choose Page',
			className: 'notContent choosePageForLinks',
			isFrontEnd: false,
			url: '/Pages/CustomLinks/ChoosePage/?pageID=' + self.pageID + '&moduleID=' + self.moduleID
		});
	}
};

/// Call back to execute after choosing a page on modal.
MenuCustomLinksManager.prototype.chosePageExecute = function (chosenPageID) {
	var self = this; // Reference to current object
	closeCpModal();
	//Update side nav.
	self.updateSideMenu('/Pages/CustomLinks/ChangeSideMenu/?pageID=' + self.pageID + '&moduleID=' + self.moduleID + '&chosenPageID=' + chosenPageID);
};

/// Creates a new empty root feature link for specific page/module
MenuCustomLinksManager.prototype.createNewCustomLink = function () {
	var self = this; // Reference to current object
	if (self.checkOrphanFeatureLinks()) {
		self.updateSideMenu('/Pages/CustomLinks/CreateCustomFeatureLinks/?pageID=' + self.pageID + '&moduleID=' + self.moduleID);
	}
};

/// Make default side menu for current page/module (remove custom link previously assigned).
MenuCustomLinksManager.prototype.choseDefaultSideMenu = function () {
	var self = this; // Reference to current object
	if (self.checkOrphanFeatureLinks()) {
		self.updateSideMenu('/Pages/CustomLinks/ChoseDefaultSideMenu/?pageID=' + self.pageID + '&moduleID=' + self.moduleID);
	}
};

/***** End - Class/functions to handle feature links on Secondary menus ********/

/***** Start - Class/functions to handle Main Menu Options ********/
function MainMenuOptions(pageID, mainMenuRootID) {
	var wasCheckBoxChanged = function ($checkbox) {
		if ($checkbox.is(":checked"))
			return $checkbox.data('original') != 'True';
		else
			return $checkbox.data('original') == 'True';
	};

	var closeMainMenuPopover = function ($handler, doNotResetCheckboxes) {
		$handler.data("cpPopover").hide();
		if (doNotResetCheckboxes)
			return;

		//reset checkbox selections
		$('div.menuOptions input[type=checkbox]').each(function () {
			if ($(this).data("original") == "True")
				$(this).prop("checked", true);
			else
				$(this).prop("checked", false);
		});
	};

	var initializeMainMenuPopup = function ($handler) {
		$('#cancelMainNavOptions', "#modalMainNavOptions")
			.unbind('click.mainMenuOptions').bind("click.mainMenuOptions", function (e) {
			e.preventDefault();
			closeMainMenuPopover($handler, false);
		});

		$('#saveMainNavOptions', "#modalMainNavOptions")
			.unbind('click.mainMenuOptions').bind("click.mainMenuOptions", function (e) {
			e.preventDefault();

			var useThisPageAsMainMenu = null;
			var displayMegaMenu = [];

			$('div.menuOptions input.DisplayMegaMenu').each(function () {
				if (wasCheckBoxChanged($(this)))
					displayMegaMenu.push({ PageID: $(this).data('pageid'), DisplayMegaMenu: $(this).is(':checked') });
			});

			if ($('#ysnUseThisPagesSubpages', 'div.menuOptions').length == 1 && wasCheckBoxChanged($('#ysnUseThisPagesSubpages'), 'div.menuOptions'))
				useThisPageAsMainMenu = $('#ysnUseThisPagesSubpages', 'div.menuOptions').is(':checked');

			var data = JSON.stringify({ pageID: pageID, versionID: $('#hdnVersionID').val(), useThisPageAsMainMenu: useThisPageAsMainMenu, displayMegaMenu: displayMegaMenu, structureID: $('#hdnStructureID').val() });

			AJAX("/Pages/Menu/MainMenuOptionsSave", 'POST', data, function (response) {
				window.location.reload(true);
			}, true, 'application/json');

			closeMainMenuPopover($handler, true);
		});
	};

	var $handlerMainMenu = $('#mainNavMenuHandleOptions');
	if ($('#hdnIsLayoutPage').val() === "True" || $('#hdnVersionID').length === 0 || $('#hdnVersionID').val() == "0") {
		$handlerMainMenu.remove();
	} else {
		$handlerMainMenu.cpPopover({
			url: '/Pages/Menu/MainMenuOptions',
			type: 'POST',
			contentType: 'application/json',
			data: JSON.stringify({ pageID: pageID, mainMenuRootID: mainMenuRootID }),
			className: 'options adminWrap admin menuOptions',
			onInit: function () {
				initializeMainMenuPopup($handlerMainMenu);
			},
			onClose: function (e) {
				e.preventDefault();
				closeMainMenuPopover($handlerMainMenu, false);
			}
		});
	}
}

/***** End - Class/functions to handle Main Menu Options ********/

function showHideAccordionMenuForSecondaryNav($elem) {
	var secondaryMenu = 1;
	var hasAccordionSubmenu = $elem.next('a.openAccordionNav').length > 0;
	if (hasAccordionSubmenu)
		$elem = $elem.next('a.openAccordionNav');
	var $container = $elem.parent().siblings('.accordionNavContainer');
	var $secondaryNavHiddenElement = $('#secondaryNav').children('[class*="hiddenAnimations"]');
	var animations = null;
	//Apply initial animation state before the sliding happens to open the accordion submenu
	if (!$elem.hasClass('accordionNavOpened')) {
		if (typeof (Animations) === 'function') {
			animations = new Animations();
			animations.applyInitial($container, secondaryMenu);
		}
	} else {//reset the container by removing animation classes.
		if ($secondaryNavHiddenElement.length > 0 && $secondaryNavHiddenElement.data('animationClass') !== undefined)
			$container.removeClass($secondaryNavHiddenElement.data('animationclass'));
		$container.removeClass('animation-triggered');
	}

	$elem.toggleClass('accordionNavOpened');
	$elem.attr('aria-expanded', $elem.attr('aria-expanded') === 'true' ? 'false' : 'true');
	$container.css('overflow', 'hidden');
	$container.slideToggle("333ms");

	if ($container.attr('aria-hidden') === 'true') {
		$container.attr('aria-hidden', 'false');
	} else {
		$container.attr('aria-hidden', 'true');
	}

	setTimeout(function () {
		$container.css('overflow', '');
		if ($elem.hasClass('accordionNavOpened')) {
			//If Animation manager toggled on, implement the menu animations
			if (animations != null)
				animations.applyAnimationClassToNavContainers($container, 1);
		}

	}, 400);

	if (typeof dynamicStretchContainers === 'function') {
		setTimeout(dynamicStretchContainers, 400);
	}
}

$(document).ready(function () {
	//Mega menus handlers
	$('.megaMenuContainer')
		.hide()
		.removeClass('hidden')
		.find('.handle.inheritance').remove();
	menuManager.initializeMegaMenus();
});
;
/*Strategy pattern
AccessibleMenu: Abstract class with default implementation for bindHandlers, initialize
MainAccessibleMenu: Concrete implementation of Accessible menu
SecondaryAccessibleMenu: Concrete implementation of Accessible menu
Note: AccordionAccessibleMenu has been moved to AccessibleAccordionMenu.js*/
var stack = [];
var AccessibleMenu = function () {
	this.menuItem = "";
};
var accessibleMenuToggleEnabled = window.FeatureToggles.isActive("CMS.Pages.AccessibleMainMenuWithIndicators");

AccessibleMenu.prototype = {
	initialize: function (mainMenuItem, accordionMenu) {
		if (mainMenuItem)
			this.setStrategy(new MainAccessibleMenu());
		else {
			this.setStrategy(new SecondaryAccessibleMenu());
		}
	},
	setStrategy: function (menuItem) {
		this.menuItem = menuItem;
	},
	bindHandlers: function ($elem) {
		var self = this;

		$elem.off('keydown.Menu').on('keydown.Menu', function (e) {
			if (e.which == 37) {//left
				self.left($elem);
				e.preventDefault();
			}
			if (e.which == 38) {//up
				self.up($elem);
				e.preventDefault();
			}
			if (e.which == 39) {//right
				self.right($elem);
				e.preventDefault();
			}
			if (e.which == 40) {//down
				self.down($elem);
				e.preventDefault();
			}
			if (e.which === 13 || e.which === 32) {//enter or space
				if (self.enter($elem)) {
					e.preventDefault();
				}
			}
			if (e.which === 9) {//tab
				if (self.tab($elem)) {
					e.preventDefault();
				}
			}
		});
	},
	left: function ($elem) {
		this.menuItem.left($elem);
	},
	up: function ($elem) {
		this.menuItem.up($elem);
	},
	right: function ($elem) {
		this.menuItem.right($elem);
	},
	down: function ($elem) {
		this.menuItem.down($elem);
	},
	enter: function ($elem) {
		return this.menuItem.enter($elem);
	},
	tab: function ($elem) {
		return this.menuItem.tab($elem);
	}
};

//MainAccessibleMenu concrete class
var MainAccessibleMenu = function () {

	// Get parent menu item from submenu
	var getParentMenuItem = function ($elem) {
		var parentOlId = $elem.parents('ol').attr('id');
		var parentMenuItemId = parentOlId.slice('Parent'.length);
		return $('#' + parentMenuItemId).find('a:first');
	};

	// Get first anchor in child submenu
	var getFirstChildSubmenuAnchor = function ($elem) {
		var $parentLi = $elem.parent('li');
		var $parentId = "Parent" + $parentLi.attr('id');
		return $('ol[id=' + $parentId + '] a:first');
	};

	this.left = function ($elem) {
		if (menuManager.isTopLevelMenuItem($elem)) {
			menuManager.handleLeftKeyOnAnyTopLevelItem($elem);
			return;
		}

		// Submenu item handling
		if (!accessibleMenuToggleEnabled) {
			getParentMenuItem($elem).focus();
			return;
		}

		var shouldOpenSubmenu = menuManager.shouldOpenSubmenuOnNavigation($elem);
		menuManager.handleLeftKeyInSubmenu($elem, shouldOpenSubmenu);
	};

	this.up = function ($elem) {
		if (menuManager.isTopLevelMenuItem($elem)) {
			if (menuManager.isMegaMenuItem($elem)) {
				return;
			}
			menuManager.handleUpKeyOnTopLevelItem($elem);
			return;
		}

		// Submenu item handling - use MenuManager handler
		menuManager.handleUpKeyInSubmenu($elem, {
			onFirstItem: function ($elem) {
				getParentMenuItem($elem).focus();
			}
		});
	};

	this.right = function ($elem) {
		if (menuManager.isTopLevelMenuItem($elem)) {
			menuManager.handleRightKeyOnAnyTopLevelItem($elem);
			return;
		}

		// Submenu item handling
		if (!accessibleMenuToggleEnabled) {
			getFirstChildSubmenuAnchor($elem).focus();
			return;
		}

		var shouldOpenSubmenu = menuManager.shouldOpenSubmenuOnNavigation($elem);
		menuManager.handleRightKeyInSubmenu($elem, shouldOpenSubmenu);
	};

	this.down = function ($elem) {
		if (menuManager.isTopLevelMenuItem($elem)) {
			if (menuManager.isMegaMenuItem($elem)) {
				return;
			}

			if (!menuManager.handleDownKeyOnTopLevelItem($elem)) {
				getFirstChildSubmenuAnchor($elem).focus();
			}
			return;
		}

		// Submenu item handling - use MenuManager handler
		menuManager.handleDownKeyInSubmenu($elem);
	};

	this.enter = function ($elem) {
		if (!menuManager.isSubMenuTrigger($elem)) {
			$elem[0].click();
			return true;
		}

		if (accessibleMenuToggleEnabled) {
			menuManager.toggleSubMenu($elem);
			menuManager.toggleIndicator($elem);
			setTimeout(function () {
				menuManager.setFocusToFirstSubmenuitem($elem);
			}, 300);
			return true;
		} else if ($elem.hasClass("mainNavItem")) {
			menuManager.toggleSubMenu($elem);
			return true;
		}

		return false;
	};

    this.tab = function ($elem) {
        menuManager.setFocusToLastTopLevelItem();

        // Return false to allow default tab behavior (focus moves to next focusable element outside menu)
        return false;
	};
};

//SecondaryAccessibleMenu concrete class
var SecondaryAccessibleMenu = function () {
	this.left = function ($elem) {
		var parentOlId = $elem.parents('ol').attr('id');
		var parentMenuItemId = parentOlId.slice('Parent'.length);
		$('#' + parentMenuItemId).find('a:first').focus();
	};

	this.up = function ($elem) {
		var $prev = $elem.parent('li').prev('li').find('a:first');
		if ($prev != null)
			$prev.focus();
	};

	this.right = function ($elem) {
		var $parentLi = $elem.parent('li');
		var $parentId = "Parent" + $parentLi.attr('id');
		var $firstAnchorOfSubmenu = $('ol[id=' + $parentId + '] a:first');
		$firstAnchorOfSubmenu.focus();
	};

	this.down = function ($elem) {
		var $next = $elem.parent('li').next('li').find('a:first');
		if ($next != null)
			$next.focus();
	};

	this.enter = function ($elem) {
		return false;
	};

	this.tab = function ($elem) {
		return false;
	};
};

//Self invoked function
(function () {
	var isClick = false;
	$("a.navMainItem, a.navMenuItem").on("mousedown", function () {
		//Set isClick true if clicked on secondary nav menu item.
		isClick = true;
	});

	function menuHandler(e) {
		var $this = $(this);

		var focusSource = $this.data('focus-source');
		if (focusSource) {
			$this.removeData('focus-source');
			if (accessibleMenuToggleEnabled) {
				menuManager.hideSubMenu($this.parent('li'));
			}
		}

		var isAccordionToggle = $this.hasClass('openAccordionNav');
		var isAccordionNavItem = $this.parent('div').hasClass('accordionNavItem');
		var isInAccordionNav = $this.closest('.accordionNav').length > 0;

		if (isAccordionToggle || isAccordionNavItem || isInAccordionNav) {
			return;
		}

		var accessibleMenu = new AccessibleMenu();
		var mainMenuItem = $this.parents('nav').hasClass('mainNav') | $this.parents('ol').hasClass('mainSubMenu');

		accessibleMenu.initialize(mainMenuItem, false);

		if (!mainMenuItem) {
			if (!accessibleMenuToggleEnabled)
				menuManager.menuMouseOver($this.parent('li'));
		} else {
			if (accessibleMenuToggleEnabled)
				menuManager.setActiveTopLevelItem($this);
			if (!accessibleMenuToggleEnabled || !mainMenuItem)
				menuManager.menuMouseOver($this.parent('li'));

			window.lastMegaMenuOpen = $this.parent('li');
		}
		accessibleMenu.bindHandlers($this);
	}

	$(document).on("focusin", "a.navMainItem, a.navMenuItem", menuHandler);

	//hide last main menu item on tabbing through the menu item.
	$(".siteWrap").on("focusin", function (e) {
		menuManager.hideAllSubMenus(MAIN_MENU);
	});
}());
;
/*Strategy pattern
AccessibleMegaMenu: Abstract class with default implementation for bindHandlers, initialize
FormattedMegaMenu: Concrete implementation of Accessible mega menu*/
var AccessibleMegaMenu = function () {
	this.megaMenuItem = "";
	this.depth = 1;
	this.cols = 1;
};

AccessibleMegaMenu.prototype = {
	initialize: function (format) {
		this.setStrategy(new FormatedMegaMenu());
	},
	setStrategy: function (megaMenuItem) {
		this.megaMenuItem = megaMenuItem;
	},
	bindHandlers: function ($elem) {
		var self = this;
		var LEFT = 37;
		var UP = 38;
		var RIGHT = 39;
		var DOWN = 40;		
        var SPACE = 32;
        var TAB = 9;
		$elem.off('keydown.MegaMenu').on('keydown.MegaMenu', function (e) {
			if (e.which == LEFT) {
				self.left($elem);
				e.preventDefault();
			}
			if (e.which == UP) {
				self.up($elem);
				e.preventDefault();
			}
			if (e.which == RIGHT) {
				self.right($elem);
				e.preventDefault();
			}
			if (e.which == DOWN) {
				self.down($elem);
				e.preventDefault();
			}
			if (e.which == SPACE) {
				self.space($elem)
				e.preventDefault();
            }
            if (e.which == TAB) {
                if (self.tab($elem)) {
                    e.preventDefault();
                }
            }
		});
	},
	exitMegaMenu: function ($elem) {
		this.megaMenuItem.exitMegaMenu($elem);
	},
	left: function ($elem) {
		this.megaMenuItem.left($elem);
	},
	up: function ($elem) {
		this.megaMenuItem.up($elem);
	},
	right: function ($elem) {
		this.megaMenuItem.right($elem);
	},
	down: function ($elem) {
		this.megaMenuItem.down($elem);
	},
	space: function ($elem) {
		return this.megaMenuItem.space($elem);
    },
    tab: function ($elem) {
        return this.megaMenuItem.tab($elem);
    }
};

//Format 1,2,3,4 concrete implementation
var FormatedMegaMenu = function () {
	// Navigate to root mega menu item
	var focusRootMegaMenuItem = function ($elem) {
		var closestMenu = $elem.closest('.mainNavMegaMenu');
		if (closestMenu.length > 0) {
			var rootItem = closestMenu.attr('id');
			if (rootItem && rootItem.startsWith('mainNavMegaMenu')) {
				rootItem = rootItem.slice('mainNavMegaMenu'.length);
				if (!isNaN(parseInt(rootItem))) {
					var $rootAnchor = $('#MainItem' + rootItem + ' a');
					if ($rootAnchor.length > 0) {
						$rootAnchor.focus();
						return true;
					}
				}
			}
		}
		return false;
	};

	// Check if element is in mega menu header
	var isMegaMenuHeader = function ($elem) {
		return $elem.closest("header").length > 0;
	};

	var getWidgetRoot = function ($elem) {
		return accessibleMenuToggleEnabled ? $elem.closest('.widgetPages') : $elem.closest('section');
	};

	// Focus element if it exists
	var focusIfExists = function ($target) {
		if ($target && $target.length > 0) {
			$target.focus();
			return true;
		}
		return false;
	};

	// Find previous widget item for up navigation
	var findPreviousWidgetItem = function ($elem) {
		var $prevLi = $elem.closest('li').prev('li');
		if (!$prevLi.hasClass('widgetItem')) {
			return $prevLi.find('a:first');
		}
		
		// If previous has children, go to last child, otherwise go to previous item
		if ($prevLi.children('ol').length > 0) {
			return $prevLi.children('ol').children('li:last').find('a:first');
		}
		return $prevLi.find('a:first');
	};

	// Find next widget item for down navigation
	var findNextWidgetItem = function ($elem, isWidgetItemElement) {
		if (isWidgetItemElement) {
			// If has children, go to first child, otherwise go to next sibling
			var $children = $elem.closest('li').children('ol');
			if ($children.length > 0) {
				return $children.children('li:first').children('a');
			}
			return $elem.closest('li').next('li').find('a:first');
		}
		
		// Try next sibling first
		var $target = $elem.closest('li').next('li').find('a:first');
		// If no sibling, try next widget item
		if ($target.length === 0) {
			$target = $elem.parents('li.widgetItem').next('li').find('a:first');
		}
		return $target;
	};

	// Check if element is at a specific position in the mega menu
	var isItemAtPosition = function ($elem, position) {
		var $megaMenuContainer = $elem.closest('.mainNavMegaMenu');
		if ($megaMenuContainer.length === 0)
			return false;

		// Find the link at the specified position (first or last)
		var $targetLink = $megaMenuContainer.find('a[role="menuitem"]')[position]();
		return $targetLink.length > 0 && $targetLink[0] === $elem[0];
	};

	var isFirstItemInFirstColumn = function ($elem) {
		return isItemAtPosition($elem, 'first');
	};

	var isLastItemInLastColumn = function ($elem) {
		return isItemAtPosition($elem, 'last');
	};

	// Navigate to adjacent top-level menu item (mega menu or normal menu)
	var navigateToAdjacentMegaMenu = function ($elem, direction) {
		var topLevelParentId = menuManager.findTopLevelParentId($elem);
		if (!topLevelParentId) return false;

		var $currentTopLevelItem = $('#' + topLevelParentId);
		if ($currentTopLevelItem.length === 0) return false;

		var menuDirection = direction === 'left' ? 'prev' : 'next';
		var $targetTopLevelItem = menuManager.getAdjacentTopLevelItem($currentTopLevelItem, menuDirection);
		if ($targetTopLevelItem.length === 0) return false;

		var $targetAnchor = $targetTopLevelItem.find('a:first');
		if ($targetAnchor.length === 0) return false;

		// Close current mega menu
		var $currentMegaMenu = $elem.closest('.mainNavMegaMenu');
		if ($currentMegaMenu.length > 0) {
			$currentMegaMenu.removeClass('hover');
			menuManager.menuMouseOut($currentTopLevelItem);
		}

		var isMegaMenu = $targetTopLevelItem.attr('data-displaymegamenu') === 'True';
		menuManager.menuMouseOver($targetTopLevelItem);

		setTimeout(function () {
			if (isMegaMenu || menuManager.isSubMenuTrigger($targetAnchor)) {
				menuManager.setFocusToFirstSubmenuitem($targetAnchor);
			}
		}, 300);

		return true;
	};

	this.left = function ($elem) {
		if (menuManager.isTopLevelMenuItem($elem)) {
			menuManager.handleLeftKeyOnAnyTopLevelItem($elem);
			return;
		}

		// Check if we're at the first item in the first column
		if (accessibleMenuToggleEnabled && isFirstItemInFirstColumn($elem)) {
			navigateToAdjacentMegaMenu($elem, 'left');
			return;
		}

		// Navigate within mega menu columns
		var $prev = $elem.closest('div.col').prev().find('li.widgetItem:first').find('a:first');
		if ($prev.length > 0) {
			$prev.focus();
		} else {
			focusRootMegaMenuItem($elem);
		}
	};

	this.up = function ($elem) {
		if (menuManager.isTopLevelMenuItem($elem)) {
			if (!menuManager.handleUpKeyOnTopLevelItem($elem)) {
				var id = $elem.closest('li').attr('id');
				$('#mainNavMegaMenu' + id).removeClass('hover');
				menuManager.menuMouseOut($elem.closest('li'));
			}
			return;
		}

		// Check if this is the first item in a widget list
		var isWidgetItemElement = $elem.closest('li').hasClass('widgetItem');
		var isFirstInWidgetList = isWidgetItemElement && 
			$elem[0] == $elem.closest('li').parent('ol').children('li:first').find('a:first')[0];

		if (isFirstInWidgetList) {
			if (accessibleMenuToggleEnabled) {
				focusRootMegaMenuItem($elem);
			} else {
				var $headerLink = getWidgetRoot($elem).find("header a");
				focusIfExists($headerLink);
			}
			return;
		}

		// Check if first item in parent list
		var $parentOl = $elem.closest('li').parent('ol');
		if ($parentOl.length > 0 && $parentOl.children('li:first').children('a')[0] == $elem[0]) {
			var $parentLink = $parentOl.parent('li').find('a:first');
			focusIfExists($parentLink);
			return;
		}

		// Navigate to previous widget item
		var $target = findPreviousWidgetItem($elem);
		focusIfExists($target);
	};

	this.right = function ($elem) {
		if (menuManager.isTopLevelMenuItem($elem)) {
			menuManager.handleRightKeyOnAnyTopLevelItem($elem);
			return;
		}

		// Check if we're at the last item in the last column
		if (accessibleMenuToggleEnabled && isLastItemInLastColumn($elem)) {
			navigateToAdjacentMegaMenu($elem, 'right');
			return;
		}

		// In mega menu content - navigate to next column
		var $next = $elem.closest('div.col').next().find('li.widgetItem:first').find('a:first');
		if ($next.length > 0) {
			$next.focus();
		}
	};

	this.down = function ($elem) {
		// Handle top-level menu items
		if ($elem.hasClass('navMainItem')) {
			if (!menuManager.handleDownKeyOnTopLevelItem($elem)) {
				// Legacy behavior - find first link in mega menu
				var $parentLi = $elem.parent('li');
				var $parentId = "mainNavMegaMenu" + $parentLi.attr('data-pageid');
				var $firstLink = accessibleMenuToggleEnabled
					? $('aside[id=' + $parentId + '] a[role="menuitem"]').first()
					: $('aside[id=' + $parentId + '] header a[role="menuitem"]').first();
				focusIfExists($firstLink);
			}
			return;
		}

		// Navigate down in mega menu content
		var isWidgetItemElement = $elem.closest('li').hasClass('widgetItem');
		var $target = findNextWidgetItem($elem, isWidgetItemElement);
		focusIfExists($target);
	};

	this.exitMegaMenu = function ($elem) {
		var asideContainerId = $elem.parents('.megaMenuContainer.megaMenu').attr('id').slice('mainNavMegaMenu'.length);
		var $li = $('#MainItem' + asideContainerId);
		$elem.parents('.megaMenuContainer.megaMenu').removeClass('hover');
		menuManager.menuMouseOut($li);
		$elem.focus();
	};
	this.space = function ($elem) {
		if (accessibleMenuToggleEnabled && !menuManager.isSubMenuTrigger($elem)) {
			$elem[0].click();
			return true;
		}
    };
    this.tab = function ($elem) {
        menuManager.setFocusToLastTopLevelItem();

        // Return false to allow default tab behavior (focus moves to next focusable element outside menu)
        return false;
    };
};

//Self invloked function
(function () {
	$(document).on('focusin', '.megaMenuContainer div.widgetTitle, .megaMenuContainer header a, #mainNav a.navMainItem , li.megaMenuItem a[role]', function (e) {
		var $this = $(this);
		var execute = true;
		var topMenuItem = $this.hasClass('navMainItem');
		if (topMenuItem) {
			var $parent = $this.parent('li');
			if ($parent != null) {
				//Logic for flyout menu is in AccessibleMenu.js, so do not execute the logic below for the same
				if ($parent.data('displaymegamenu') == 'False')
					execute = false;
			}
		}
		if (execute) {
			var depth = $this.parents('.widgetPages').find('.depth').html();
			var cols = $this.parents('.widgetPages').find('.cols').html();
			var accessibleMegaMenu = new AccessibleMegaMenu();
			var mainMenuItem = true;
			//If top menu item(and has megamenu which is already checked above), replicate the logic for "mouse near" to load the mega menus
			if (topMenuItem) {
				var megaMenuContentContainers = $('[data-displaymegamenu="True"]')
					.map(function () {
						$('#mainNavMegaMenu' + $(this).data('pageid') + ' div.pageContent').html('Loading...');
						return $('#mainNavMegaMenu' + $(this).data('pageid') + ' div.pageContent').data('containerid');
					}).toArray();
				reloadSharedContentContainers(megaMenuContentContainers);
			}
			accessibleMegaMenu.initialize(depth, cols);
			accessibleMegaMenu.bindHandlers($this);
		}
	});
}());
;
/*Strategy pattern
AccessibleAccordionMenu: Self-contained module for handling accordion menu keyboard navigation
Moved from AccessibleMenu.js for better maintainability
*/

var AccessibleAccordionMenu = function () {
	this.accordionMenuItem = "";
};

//Static field for accordion menu navigation stack
AccessibleAccordionMenu.accordionMenuStack = [];

AccessibleAccordionMenu.prototype = {
	initialize: function () {
		this.setStrategy(new AccordionMenu());
	},
	setStrategy: function (accordionMenuItem) {
		this.accordionMenuItem = accordionMenuItem;
	},
	bindHandlers: function ($elem) {
		var self = this;
		var LEFT = 37;
		var UP = 38;
		var RIGHT = 39;
		var DOWN = 40;
		var ENTER = 13;
		var SPACE = 32;
		var TAB = 9;

		$elem.off('keydown.AccordionMenu').on('keydown.AccordionMenu', function (e) {
			if (e.which === LEFT) {
				self.left($elem);
				e.preventDefault();
			}
			if (e.which === UP) {
				self.up($elem);
				e.preventDefault();
			}
			if (e.which === RIGHT) {
				self.right($elem);
				e.preventDefault();
			}
			if (e.which === DOWN) {
				self.down($elem);
				e.preventDefault();
			}
			if (e.which === ENTER || e.which === SPACE) {
				if (self.enter($elem)) {
					e.preventDefault();
				}
			}
			if (e.which === TAB) {
				if (self.tab($elem)) {
					e.preventDefault();
				}
			}
		});
	},
	left: function ($elem) {
		this.accordionMenuItem.left($elem);
	},
	up: function ($elem) {
		this.accordionMenuItem.up($elem);
	},
	right: function ($elem) {
		this.accordionMenuItem.right($elem);
	},
	down: function ($elem) {
		this.accordionMenuItem.down($elem);
	},
	enter: function ($elem) {
		return this.accordionMenuItem.enter($elem);
	},
	tab: function ($elem) {
		return this.accordionMenuItem.tab($elem);
	}
};

//AccordionMenu concrete implementation
var AccordionMenu = function () {

	var getAdjacentAccordionTopLevelItem = function ($currentLi, direction) {
		var $targetLi = direction === 'prev' ? $currentLi.prev('.topMenuItem') : $currentLi.next('.topMenuItem');

		// If no adjacent sibling found, wrap around
		if ($targetLi.length === 0) {
			var $siblings = $currentLi.parent().children('.topMenuItem');
			$targetLi = direction === 'prev' ? $siblings.last() : $siblings.first();
		}

		return $targetLi;
	};

	var isMainLevelAccordionItem = function ($elem) {
		return $elem.hasClass('navMainItem') && $elem.hasClass('secondaryNavItem');
	};

	var isTopLevelToggle = function ($elem) {
		return $elem.hasClass('accordionSecondaryNavItem');
	};

	var isNestedToggle = function ($elem) {
		return $elem.hasClass('accordionNavMenuItem');
	};

	var isSubmenuLink = function ($elem) {
		return $elem.hasClass('navMenuItem');
	};

	var focusElement = function ($elem) {
		if ($elem && $elem.length > 0) {
			$elem.focus();
			return true;
		}
		return false;
	};

    var hasOpenedSubmenuInCurrentContext = function ($elem) {
        // Check if the currently focused main-level item has an opened submenu directly under it
        if (!isMainLevelAccordionItem($elem)) {
            return false;
        }
        
        var $parentTopMenuItem = $elem.closest('.topMenuItem');
        if ($parentTopMenuItem.length > 0) {
            var $openedSubmenu = $parentTopMenuItem.children('.accordionNavContainer:visible');
            return $openedSubmenu.length > 0;
        }
        return false;
    };

    var getParentToggleButton = function ($parentSubmenu) {
        var $parentNavItemDiv = $parentSubmenu.siblings('.accordionNavItem');
        return $parentNavItemDiv.find('a.openAccordionNav');
    };

    var navigateToAdjacentSibling = function ($currentMenuItem, direction, shouldPushToStack, $currentElem) {
        var $siblingLi = direction === 'next' 
            ? $currentMenuItem.next('li[role=menuitem]')
            : $currentMenuItem.prev('li[role=menuitem]');

        if ($siblingLi.length > 0) {
            if (shouldPushToStack && $currentElem) {
                AccessibleAccordionMenu.accordionMenuStack.push($currentElem);
            }
            focusElement($siblingLi.find('a:first'));
            return true;
        }
        return false;
    };

    // Helper: Navigate to parent level when no sibling exists
    var navigateToParentLevel = function ($currentMenuItem, direction, shouldPushToStack, $currentElem) {
        var $parentSubmenu = $currentMenuItem.parent('.accordionNavContainer');
        if ($parentSubmenu.length === 0) {
            return false;
        }

        var $grandParentMenuItem = $parentSubmenu.parent('[role=menuitem]');
        if ($grandParentMenuItem.length === 0) {
            return false;
        }

        if (direction === 'next') {
            // Moving down - go to parent's next sibling
            var $parentNextSibling = $grandParentMenuItem.next('li[role=menuitem]');
            if ($parentNextSibling.length > 0) {
                if (shouldPushToStack && $currentElem) {
                    AccessibleAccordionMenu.accordionMenuStack.push($currentElem);
                }
                focusElement($parentNextSibling.find('a:first'));
                return true;
            }

            // No parent sibling - move to top-level's next item
            var $topMenuItem = $grandParentMenuItem.closest('.topMenuItem');
            if ($topMenuItem.length > 0) {
                var $nextTopMenuItem = getAdjacentAccordionTopLevelItem($topMenuItem, 'next');
                if (shouldPushToStack && $currentElem) {
                    AccessibleAccordionMenu.accordionMenuStack.push($currentElem);
                }
                focusElement($nextTopMenuItem.find('a.navMainItem:first'));
                return true;
            }
        } else {
            // Moving up - go to parent's toggle button
            var $parentToggleButton = getParentToggleButton($parentSubmenu);
            if ($parentToggleButton.length > 0 && $parentToggleButton.hasClass('accordionNavOpened')) {
                focusElement($parentToggleButton);
                return true;
            }
        }

        return false;
    };

    var closeSiblingSubmenus = function ($currentMenuItem) {
        var $siblings = $currentMenuItem.siblings('li[role=menuitem]');
        
        $siblings.each(function() {
            var $siblingToggle = $(this).find('> .accordionNavItem > a.openAccordionNav');
            if ($siblingToggle.length > 0 && $siblingToggle.hasClass('accordionNavOpened')) {
                $siblingToggle[0].click();
            }
        });
    };

    // Top-level toggle (accordionSecondaryNavItem) arrow handling
    var handleTopLevelToggleUp = function ($elem) {
        var $currentTopMenuItem = $elem.closest('.topMenuItem');

        // If submenu is open, close it first
        if ($elem.hasClass('accordionNavOpened')) {
            $elem[0].click();
        }

        var $prevTopMenuItem = getAdjacentAccordionTopLevelItem($currentTopMenuItem, 'prev');
        focusElement($prevTopMenuItem.find('a.navMainItem:first'));
    };

    var handleTopLevelToggleDown = function ($elem) {
        var $parentMenuItem = $elem.closest('[role=menuitem]');

        // If submenu is open, move INTO the submenu (don't close)
        if ($elem.hasClass('accordionNavOpened')) {
            var $submenu = $parentMenuItem.children('.accordionNavContainer');
            if ($submenu.length > 0 && $submenu.is(':visible')) {
                var $firstSubmenuItem = $submenu.children('[role=menuitem]:first').find('a:first');
                if ($firstSubmenuItem.length > 0) {
                    AccessibleAccordionMenu.accordionMenuStack.push($elem);
                    focusElement($firstSubmenuItem);
                    return;
                }
            }
        }

        // Submenu not open or no children: go to next top-level
        var $nextTopMenuItem = getAdjacentAccordionTopLevelItem($parentMenuItem, 'next');
        focusElement($nextTopMenuItem.find('a.navMainItem:first'));
    };

    // Nested toggle (accordionNavMenuItem) arrow handling
    var handleNestedToggleUp = function ($elem) {
        var $parentMenuItem = $elem.closest('[role=menuitem]');

        // If submenu is open, close it first
        if ($elem.hasClass('accordionNavOpened')) {
            $elem[0].click();
        }

        // Try previous sibling first; if none, go to parent level
        if (navigateToAdjacentSibling($parentMenuItem, 'prev', false, null)) {
            return;
        }
        navigateToParentLevel($parentMenuItem, 'prev', false, null);
    };

    var handleNestedToggleDown = function ($elem) {
        var $parentMenuItem = $elem.closest('[role=menuitem]');

        // If submenu is open, move INTO the submenu (don't close)
        if ($elem.hasClass('accordionNavOpened')) {
            var $submenu = $parentMenuItem.children('.accordionNavContainer');
            if ($submenu.length > 0 && $submenu.is(':visible')) {
                var $firstSubmenuItem = $submenu.children('[role=menuitem]:first').find('a:first');
                if ($firstSubmenuItem.length > 0) {
                    AccessibleAccordionMenu.accordionMenuStack.push($elem);
                    focusElement($firstSubmenuItem);
                    return;
                }
            }
        }

        // Submenu not open - try next sibling; else parent level
        if (navigateToAdjacentSibling($parentMenuItem, 'next', false, null)) {
            return;
        }
        navigateToParentLevel($parentMenuItem, 'next', false, null);
    };

    // Submenu link (navMenuItem) arrow handling
    var handleSubmenuItemUp = function ($elem) {
        var $currentMenuItem = $elem.closest('[role=menuitem]');
        var $parentSubmenu = $currentMenuItem.parent('.accordionNavContainer');

        if ($parentSubmenu.length === 0) return;

        // If this is first item, go to its toggle
        var $firstItem = $parentSubmenu.children('[role=menuitem]:first');
        if ($firstItem.length > 0 && $currentMenuItem[0] === $firstItem[0]) {
            var $parentToggleButton = getParentToggleButton($parentSubmenu);
            if ($parentToggleButton.length > 0 && $parentToggleButton.hasClass('accordionNavOpened')) {
                focusElement($parentToggleButton);
                return;
            }
        }

        // Otherwise go to previous sibling
        navigateToAdjacentSibling($currentMenuItem, 'prev', false, null);
    };

    var handleSubmenuItemDown = function ($elem) {
        var $currentMenuItem = $elem.closest('[role=menuitem]');

        // Try next sibling first
        if (navigateToAdjacentSibling($currentMenuItem, 'next', true, $elem)) {
            return;
        }

        // No sibling - go to parent level
        navigateToParentLevel($currentMenuItem, 'next', true, $elem);
    };

    this.left = function ($elem) {
        //do nothing
    };

    this.up = function ($elem) {
        // Top-level toggle
        if (isTopLevelToggle($elem)) {
            handleTopLevelToggleUp($elem);
            return;
        }

        // Nested toggle
        if (isNestedToggle($elem)) {
            handleNestedToggleUp($elem);
            return;
        }

        // Submenu link
        if (isSubmenuLink($elem)) {
            handleSubmenuItemUp($elem);
            return;
        }

        // Main-level text links (navMainItem)
        if (isMainLevelAccordionItem($elem)) {
            if (AccessibleAccordionMenu.accordionMenuStack.length > 0) {
                var $popped = AccessibleAccordionMenu.accordionMenuStack.pop();
                if ($popped) {
                    $popped.focus();
                }
                return;
            }
            var $currentTopMenuItem = $elem.parents('.topMenuItem:first');
            var $prevTopMenuItem = getAdjacentAccordionTopLevelItem($currentTopMenuItem, 'prev');
            focusElement($prevTopMenuItem.find('a.navMainItem:first'));
            return;
        }

        // Fallback to stack for any other element
        var $poppedFallback = AccessibleAccordionMenu.accordionMenuStack.pop();
        if ($poppedFallback) {
            $poppedFallback.focus();
        }
    };

    this.right = function ($elem) {
        //do nothing
    };

    this.down = function ($elem) {
        // Top-level toggle
        if (isTopLevelToggle($elem)) {
            handleTopLevelToggleDown($elem);
            return;
        }

        // Nested toggle
        if (isNestedToggle($elem)) {
            handleNestedToggleDown($elem);
            return;
        }

        // Submenu link
        if (isSubmenuLink($elem)) {
            handleSubmenuItemDown($elem);
            return;
        }

        // Main-level text links (navMainItem) when no submenu is open
        if (isMainLevelAccordionItem($elem) && !hasOpenedSubmenuInCurrentContext($elem)) {
            var $currentTopMenuItem = $elem.parents('.topMenuItem:first');
            var $nextTopMenuItem = getAdjacentAccordionTopLevelItem($currentTopMenuItem, 'next');
            focusElement($nextTopMenuItem.find('a.navMainItem:first'));
            return;
        }

        // Original behavior for nested/other items
        AccessibleAccordionMenu.accordionMenuStack.push($elem);
        var hasAccordionSubmenu = $elem.next('a.openAccordionNav').length > 0;
        var accordionSubmenuAlreadyOpen = $elem.next('a.openAccordionNav').hasClass('accordionNavOpened');
        var $nextElementToFocus = '';

        if (hasAccordionSubmenu && !accordionSubmenuAlreadyOpen) {
            $nextElementToFocus = $elem.parents('.topMenuItem').next('li').find('a:first');
            if ($nextElementToFocus.length > 0) {
                $nextElementToFocus.focus();
            }
            return;
        }

        if (hasAccordionSubmenu) {
            $nextElementToFocus = $elem.parents('[role=menuitem]:first').children('ol.accordionNavContainer').find('a:first');
            if ($nextElementToFocus.length > 0) {
                $nextElementToFocus.focus();
            }
        } else {
            var $nextSiblingLi = $elem.parents('[role=menuitem]').next('li:first');
            if ($nextSiblingLi && $nextSiblingLi.length > 0) {
                $nextElementToFocus = $nextSiblingLi.find('a:first');
                if ($nextElementToFocus.length > 0) {
                    $nextElementToFocus.focus();
                }
            } else {
                $nextElementToFocus = $elem.parents('.topMenuItem').next('li').find('a:first');
                if ($nextElementToFocus.length > 0) {
                    $nextElementToFocus.focus();
                }
            }
        }
    };

    this.enter = function ($elem) {
        // Focus is on accordion toggle button (+/- or arrow)
        if ($elem.hasClass('openAccordionNav')) {
            var $parentMenuItem = $elem.closest('[role=menuitem]');
            
            // If toggle is currently closed (will be opened), close sibling submenus first
            if (!$elem.hasClass('accordionNavOpened')) {
                closeSiblingSubmenus($parentMenuItem);
            }
            
            $elem[0].click();
            return true;
        }

        if (isMainLevelAccordionItem($elem)) {
            $elem[0].click();
            return true;
        }

        if (isSubmenuLink($elem)) {
            $elem[0].click();
            return true;
        }

        return false;
    };

    this.tab = function ($elem) {
        return false;
    };
};

//Self invoked function
(function () {
	function accordionMenuHandler(e) {
		var $this = $(this);

		var accessibleAccordionMenu = new AccessibleAccordionMenu();
		var $topmostAccordionMenuItem = $this.parents('nav').find('a:first');

		if ($this[0] === $topmostAccordionMenuItem[0]) {
			AccessibleAccordionMenu.accordionMenuStack = [];
		}

		accessibleAccordionMenu.initialize();
		accessibleAccordionMenu.bindHandlers($this);
	}

	$(document).on("focusin", ".accordionNav a.navMainItem, .accordionNav a.navMenuItem, .accordionNav a.openAccordionNav", accordionMenuHandler);
}());
var AccessibleAnchors = function () {
	
};

AccessibleAnchors.prototype =
{
	configure: function (config) {
		var self = this;
		self.OPEN_IN_NEW_WINDOW_TEXT = ' Opens in new window';
		if (config.type === 'popup') {
			this.bindPopupHandlers(self);
		}
		if (config.type === 'extend') {
			this.bindExtendHandlers(self);
		}
		if (config.type === 'both') {
			this.bindPopupHandlers(self);
			this.bindExtendHandlers(self);
		}
		$(document).ready(function () {
			var anchors_TargetBlank = $('a[target=_blank]');
			$.each(anchors_TargetBlank, function () {
				var $this = $(this);
				self.setAriaLabelAndAlt($this);
			});
		});
	},
	bindExtendHandlers: function (self) {
		$('a[target=_blank]').on("focusin mouseenter", function (e) {
			var $this = $(this);
			self.setAriaLabelAndAlt($this);
		});		
	},
	setAriaLabelAndAlt: function ($this) {
		//If anchor has an image inside it
		if (this.hasImage($this)) {
			var $img = $this.children('img');
			if ($img.length == 0)
				$img = $this.children('span').children('img');
			var alt = $img.attr('alt');
			if (alt) {
				if (alt.indexOf(this.OPEN_IN_NEW_WINDOW_TEXT) === -1)
					$img.attr('alt', alt + this.OPEN_IN_NEW_WINDOW_TEXT);
			}
			else {
				$img.attr('alt', $img.attr('title') + this.OPEN_IN_NEW_WINDOW_TEXT);
			}
			this.setAriaLabel($this, $img.attr('alt'));
		}
		else {
			this.setAriaLabel($this, $this.text() + this.OPEN_IN_NEW_WINDOW_TEXT);
		}
	},
	bindPopupHandlers: function (self) {
		$('a[target=_blank]').on("focusin mouseenter", function (e) {
			var $this = $(this);
			$this.trigger('focusout');
			$this.append($('<span aria-hidden="false" aria-label="opens in new window" title="opens in new window" class="newWindowIndicator"></span>'));
		});	

		$('a[target=_blank]').on("focusout mouseleave", function (e) {
			$('.newWindowIndicator').remove();
		});		
	},
	setAriaLabel: function($anchor, ariaLabelText){
		var ariaLabel = $anchor.attr('aria-label');
		if (ariaLabel) {
			if (ariaLabel.indexOf(this.OPEN_IN_NEW_WINDOW_TEXT) == -1)
				$anchor.attr('aria-label', ariaLabelText);
		}
		else {
			$anchor.attr('aria-label', ariaLabelText);
		}
	},
	hasImage: function ($anchor) {
		var hasImage = false;
		if ($anchor.children('img') != null && $anchor.children('img').length > 0)
			return true;
		else if ($anchor.children('span') != null && $anchor.children('span').length > 0) {
			return $anchor.children('span').children('img') != null && $anchor.children('span').children('img').length > 0;
		}
	}

};

//Self invloked function
(function () {
	//Add a warning span to body for anchors that open in new window
	$(
		new AccessibleAnchors().configure({ type: 'extend' })
	);
}());



;
/* File Created: February 11, 2013 by Vishal Bhatia*/

function rearrangeFlyouts(clearDOM) {
    if (clearDOM) {
        //Remove any old .pop elements from the DOM to reload the new one.
        $('.flyOut.moved').remove();
    }
    //Grab all the flyouts that are not moved yet and are not tooltips and not in WidgetManager
    var $flyOutDivs = $('.flyOut').not('.moved').not('.helpTip').not('.widgetManager');
    var parentID;
    var marginLeft, marginTop;
    $flyOutDivs.each(function (e) {
        parentID = $(this).parent('.flyOutParent').attr('id');

        //If the element does not have a flyOutParent, display it on the center of the screen.
        if (parentID == undefined) {
            marginLeft = '-' + $(this).outerWidth() / 2 + 'px';
            marginTop = '-' + $(this).outerHeight() / 2 + 'px';

            $(this).css('left', '50%');
            $(this).css('margin-left', marginLeft);
            $(this).css('top', '50%');
            $(this).css('margin-top', marginTop);

            $(this).appendTo('body');
            $(this).addClass('moved');
            //TODO: Confirm from Kyle and remove 2 lines below since they are being done through css.
            $(this).css('position', 'absolute');
            $(this).css('z-index', '100');
        }
        else {
            //Remove any pre existing copies of this element from the DOM.
        	$('[data-parentID=' + parentID + ']').remove();
			//Remove margin styles if the element has a flyOutParent.
            $(this).css({ 'margin-left': '', 'margin-top': '' });
            $(this).appendTo('body');
            $(this).addClass('moved');
            //TODO: Confirm from Kyle and remove 2 lines below since they are being done through css.
            $(this).css('position', 'absolute');
            $(this).css('z-index', '100');
            $(this).attr('data-parentID', parentID);
            adjustFlyoutPosition($('#' + parentID), $(this));
        }
    });
}

///Executes when the browser is being resized or the screen orientation changes(portrait/landscape).
function rearrangeFlyoutsOnResizeEvents() {
	//Grab all the flyouts that are not moved yet and are not tooltips and not in WidgetManager
	var $flyOutDivs = $('.flyOut').not('.helpTip').not('.widgetManager').not('.mainNavMegaMenu.megaMenuContainer');;
	var parentID;
	var marginLeft, marginTop;
	$flyOutDivs.each(function (e) {
		parentID = $(this).data('parentid');
		//If the element does not have a flyOutParent, display it on the center of the screen.
		if (parentID == undefined) {
			marginLeft = '-' + $(this).outerWidth() / 2 + 'px';
			marginTop = '-' + $(this).outerHeight() / 2 + 'px';
			$(this).css('margin-left', marginLeft);
			$(this).css('margin-top', marginTop);
		}
		else {			
			adjustFlyoutPosition($('#' + parentID), $(this));
		}
	});
}

///Checks if any ancestor of a flyout element is fixed positioned.
function anyAncestorFixedPositioned($parent) {
	var anyAncestorFixedPosition = false;
	if ($parent.css('position') === 'fixed')
		anyAncestorFixedPosition = true;
	if (!anyAncestorFixedPosition) {
		$parent.parents().each(function () {
			if (!anyAncestorFixedPosition) {
				if ($(this).css('position') === 'fixed')
					anyAncestorFixedPosition = true;
			}
		});
	}
	return anyAncestorFixedPosition;
}

function adjustFlyoutPosition($parent, $this) {
    var $parentElem = $parent;    
    if ($parentElem != null) {        
        var parentWidth = $parent.outerWidth(false);
        var parentHeight = $parent.outerHeight(false);
        var offset = $parent.offset();        
        if (offset != null) {
            var offsetLeft = offset.left;
            //Take into account scroll height if the document has been scrolled, to compute the top distance from viewport. VB
            var offsetTop = $this.css('position') === 'fixed' ? offset.top - $(document).scrollTop() : offset.top; // Take into account if the item is fixed or absolutely positioned.                        
                        
			if (anyAncestorFixedPositioned($parent))
				$this.css('position', 'fixed');

            var offsetRight;
            var offsetBottom;

            if ($this.hasClass('left') || $this.hasClass('leftCorner')) {
                offsetLeft = offsetLeft - $this.outerWidth(false);
                $this.css('left', offsetLeft);
            }

            if ($this.hasClass('right') || $this.hasClass('rightCorner')) {
                offsetLeft = offsetLeft + parentWidth;
                $this.css('left', offsetLeft);
            }

            if ($this.hasClass('top') || $this.hasClass('topCorner')) {
                offsetBottom = offsetTop - $this.outerHeight(false);
                $this.css('top', offsetBottom);
            }

            if ($this.hasClass('bottom') || $this.hasClass('bottomCorner')) {
                offsetTop = offsetTop + parentHeight;
                $this.css('top', offsetTop);                
            }

            if ($this.hasClass('leftEdge')) {
                $this.css('left', offsetLeft);

                if ($this.hasClass('rightEdge')) {
                    $this.css('width', parentWidth);
                }                
            }

            if ($this.hasClass('rightEdge')) {
                offsetLeft = offsetLeft + parentWidth - $this.outerWidth(false);
                $this.css('left', offsetLeft);
            }

            if ($this.hasClass('topEdge')) {
                $this.css('top', offsetTop);

                if ($this.hasClass('bottomEdge')) {
                    $this.css('height', parentHeight);
                }
            }

            if ($this.hasClass('bottomEdge')) {
                offsetTop = offsetTop + parentHeight - $this.outerHeight(false);                
                $this.css('top', offsetTop);
            }

            if ($this.hasClass('center') && ($this.hasClass('top') || $this.hasClass('bottom'))) {
                offsetLeft = offsetLeft + (parentWidth / 2) - ($this.outerWidth(false) / 2);
                $this.css('left', offsetLeft);
            }

            if ($this.hasClass('center') && ($this.hasClass('left') || $this.hasClass('right'))) {
                offsetTop = offsetTop + (parentHeight / 2) - ($this.outerHeight(false) / 2);
                $this.css('top', offsetTop);
            }
        }
    }
};
// Note some of this functionality is in Menu.js as well

// Deferred Loading for mega menus
$(function () {
	"use strict";

	var thisRef = {};

	thisRef.initialized = false;

	var LOAD_TRIGGER_DISTANCE = 200; // in px
	var $megaMenus = $('[data-displaymegamenu="True"]');
	function isNear(element, distance, event) {
		var left = element.offset().left - distance,
			top = element.offset().top - distance,
			right = left + element.width() + 2 * distance,
			bottom = top + element.height() + 2 * distance,
			x = event.pageX,
			y = event.pageY;

		return (x > left && x < right && y > top && y < bottom);
	};

	$('body:not(.mobile)').bind('mousemove.megaMenuLazyLoad', function (e) {
		$megaMenus.each(function (index) {
			var mouseIsNear = isNear($(this), LOAD_TRIGGER_DISTANCE, e);
			if (mouseIsNear) {
				if (!thisRef.initialized) {
					$('body').unbind('.megaMenuLazyLoad');
					thisRef.initialized = true;
					var megaMenuContentContainers = $('[data-displaymegamenu="True"]')
					.map(function () {
						$('#mainNavMegaMenu' + $(this).data('pageid') + ' div.pageContent').html('Loading...');
						return $('#mainNavMegaMenu' + $(this).data('pageid') + ' div.pageContent').data('containerid');
					}).toArray();

					window.Pages.megaMenuLoaded.resolve();
					reloadSharedContentContainers(megaMenuContentContainers);
				}
			} 
		});
	});
});

var reloadSharedContentContainers = function (containerIDs) {
	if (containerIDs.length > 0) {
		// Reload all mega menus?
		var url = getAbsoluteUrl('/Content/Load?contentCollectionID=' + $('#hdnContentCollectionID').val() + "&StructureID=" + $('#hdnStructureID').val() + "&themeID=" + $('#hdnThemeID').val());
		if ($('body').hasClass('preview'))
			url += "&preview=true";
		$.ajax({
			url: url,
			type: 'POST',
			data: JSON.stringify({ "contentContainerIDs": containerIDs }),
			contentType: 'application/json',
			cache: false,
			async: true,
			ignoreEditor: true, //prevent for executing global editor event
			error: function (xhr, textStatus, exception) {
				if (xhr.status != 0) {
					alert("Error: " + xhr.statusText + "\nStatus: " + xhr.status);
				}
			},
			success: function (response) {
				if (response.PageModified !== undefined && response.PageModified === true && response.PageModifiedMessage) {
					if (confirm(response.PageModifiedMessage)) {
						document.cookie = "refresh=true; max-age=5";
						location.reload();
					}
				}
				else if (response.ErrorMessage)
					alert(response.ErrorMessage);
				else {
					//Replace content containers HTML
					var root = document.createElement("div");
					root.innerHTML = response;
					$(root).find('#fb-root').remove();
					$('head').append($(root).find('.bundledStyles'));
					var allChilds = root.childNodes;
					for (var i = 0; i < containerIDs.length; i++) {
						var contentContainerID = containerIDs[i];
						//Grab Content Container HTML from response
						var contentContainerHtml = null;
						for (var j = 0; j < allChilds.length; j++) {
							if (allChilds[j].id && allChilds[j].id == 'cc' + contentContainerID) {
								contentContainerHtml = allChilds[j].outerHTML;
								break;
							}
						}
						if (contentContainerHtml != null) {
							var $contentContainer = $('#cc' + contentContainerID);
							$contentContainer.empty().replaceWith(contentContainerHtml);
						} else {
							//ContentContainer Html was not found. This might happen when we have a html version and switch (Refresh containers) to a linked version on the same page.
						}
					}

					if (window.cpMedia) {
						$('.mainNavMegaMenu.megaMenuContainer').each(function () {
							applyElementQueryToChildren($(this));
						});
					}

					//TODO: unbind/destroy events/drag/drop on elements before initialize containers
					if (typeof InitializeContainers == 'function' && getCookieValue("enableLiveEdit") === "true") {
						InitializeContainers({ hidePopover: false });
						$('.mainNavMegaMenu.megaMenuContainer').find('.pinned').removeClass('pinned');
						setTimeout(function () { // It is unfortunate to use  timeout here, but there is a very difficult to reproduce race condition in FF where it is still getting unpinned by something
							$('.mainNavMegaMenu.megaMenuContainer').find('.pinned').removeClass('pinned');
						},300);
					}
					if (menuManager && typeof menuManager.attachDragAndDropHandlers == 'function') {
						menuManager.attachDragAndDropHandlers();
					}

					//Execute this function or else the click event wont be registered for tabbed widgets inside megamenus.
					Core.Layout.attachTabbedWidgetTabHandlers();
				}
			}
		});
	}
};

// Handle touchstart event for touch devices. First click will load the megamenus if they haven't been loaded and show the mega menu. Second click will navigate to the menu link itself.
$('[data-displaymegamenu="True"]').bind('touchstart', function (e) {
	e.stopPropagation(); // prevent this event from bubbling up and triggering the handler on document below!
	if (($(this).closest('.mainNavHolderMobile')[0] !== undefined) || ($('body:not(.mobile)').length === 0)) { // Return if mobile menu is being shown
		return;
	}
	if (window.lastMegaMenuOpen === this) { // Used to do the second click navigation if the same menu is clicked twice.
		return;
	}
	e.preventDefault();
	if (!window.megaMenuInitializedForTouch) {
		$('body').unbind('.megaMenuLazyLoad');
		window.megaMenuInitializedForTouch = true;
		var megaMenuContentContainers = $('[data-displaymegamenu="True"]')
		.map(function () {
			$('#mainNavMegaMenu' + $(this).data('pageid') + ' div.pageContent').html('Loading...');
			return $('#mainNavMegaMenu' + $(this).data('pageid') + ' div.pageContent').data('containerid');
		}).toArray();

		window.Pages.megaMenuLoaded.resolve();
		reloadSharedContentContainers(megaMenuContentContainers);
	}

	menuManager.menuMouseOver($(this)); 
	window.lastMegaMenuOpen = this;
});

$(document).bind('touchstart click', function (e) {
	if (window.lastMegaMenuOpen === undefined || window.lastMegaMenuOpen === this || !isResolvedOrState(window.Pages.megaMenuLoaded)) {
		return;
	}
	//Skip menumouseout method if user slide up the expanded mega menu to scroll
	if (!e.target.id.startsWith("mainNavMegaMenu")) {
		if (!$(e.target).parents("aside[id^='mainNavMegaMenu']").length) {
			menuManager.menuMouseOut($(window.lastMegaMenuOpen));
		}
	}
});
(function () {
	const focusableSelector = 'a:not([disabled]):not([aria-hidden="true"]),input:not([disabled]):not([type="hidden"])';
	const toggleEl = (target, toggleable) => {
		const display = target.getAttribute("aria-expanded") === "false";
		toggleable.style.display = display ? "block" : "none";
		target.setAttribute("aria-expanded", display);

		if (display) {
			target.classList.add("active");
		} else {
			target.classList.remove("active");
		}
	};

	const repositionPopover = (target) => {
		const elId = target.getAttribute("aria-controls");
		if (!elId) {
			return;
		}

		const popover = document.getElementById(elId);
		if (!popover) {
			return;
		}

		const popoverWidth = popover.offsetWidth;
		const rect = target.getBoundingClientRect();
		const targetHeight = target.offsetHeight;
		const distanceFromRight = window.innerWidth - rect.left;
		const parent = target.parentElement;
		const parentWidth = parent.offsetWidth;

		if (distanceFromRight < popoverWidth && parentWidth < popoverWidth) {
			popover.classList.add("flip");
			const shift = parentWidth - popoverWidth;
			popover.style.left = `${shift}px`;
		} else {
			popover.classList.remove("flip");
			popover.style.left = `0px`;
		}

		const parentRect = parent.getBoundingClientRect();
		const distanceFromParentBottom = parentRect.bottom - rect.bottom;
		popover.style.bottom = `${distanceFromParentBottom + targetHeight}px`;
		popover.style.zIndex = 1;
		popover.scrollIntoView({ behavior: "smooth", block: "center", inline: "nearest"});
	};

	const handleCpNavKeydown = (event, focusables) => {
		if (focusables.length < 1) {
			return;
		}

		const currentFocusedIndex = focusables.findIndex(element => element === document.activeElement);
		let nextFocusIndex = -1;
		switch (event.key) {
			case 'ArrowLeft':
			case 'ArrowUp':
				event.preventDefault();
				if (currentFocusedIndex > 0) {
					nextFocusIndex = currentFocusedIndex - 1;
				} else {
					nextFocusIndex = focusables.length - 1;
				}
				break;
			case 'ArrowRight':
			case 'ArrowDown':
				event.preventDefault();
				if (currentFocusedIndex < focusables.length - 1) {
					nextFocusIndex = currentFocusedIndex + 1;
				} else {
					nextFocusIndex = 0;
				}
				break;
			case 'Home':
				event.preventDefault();
				nextFocusIndex = 0;
				break;
			case 'End':
				event.preventDefault();
				nextFocusIndex = focusables.length - 1;
				break;
			default:
				break;
		}

		if (nextFocusIndex > -1) {
			focusables[currentFocusedIndex].setAttribute('tabindex', '-1');
			focusables[nextFocusIndex].setAttribute('tabindex', '0');
			focusables[nextFocusIndex].focus();
		}
	};

	const handleCpNav2DKeydown = (e, focusables) => {
		const map = {
			"ArrowLeft": { pos: "left", sign: 1, finder: "findLast", ctrl: false },
			"ArrowRight": { pos: "left", sign: -1, finder: "find", ctrl: false },
			"ArrowUp": { pos: "top", sign: 1, finder: "findLast", ctrl: false },
			"ArrowDown": { pos: "top", sign: -1, finder: "find", ctrl: false },
			"Home": { pos: "left", sign: 1, finder: "find", ctrl: e.ctrlKey },
			"End": { pos: "left", sign: -1, finder: "findLast", ctrl: e.ctrlKey }
		};

		const settings = map[e.key];

		if (!settings) {
			return;
		}

		const current = e.target.closest(focusableSelector);

		if (!current) {
			return;
		}

		e.preventDefault();

		const pos = settings.pos;
		const offsetConst = pos === "top" ? "left" : "top";
		const sign = settings.sign;
		const ctrlKey = settings.ctrl;
		const rectCurrent = current.getBoundingClientRect();
		const tolerance = pos === "top" ? current.offsetHeight / 3 : current.offsetWidth / 3;

		const finder = settings.finder;

		const elFound = focusables[finder](el => {
			const rect = el.getBoundingClientRect();
			const currentDistance = sign * (rectCurrent[pos] - rect[pos]);
			const offset = Math.abs(rectCurrent[offsetConst] - rect[offsetConst]);
			return (offset < tolerance && currentDistance > 0) || ctrlKey;
		});

		if (!elFound) {
			return;
		}

		current.setAttribute("tabindex", "-1");
		elFound.setAttribute("tabindex", "0");
		elFound.focus();
		return elFound;
	}

	const handleEl = (target) => {
		const elId = target.getAttribute("aria-controls");
		if (!elId) {
			return;
		}

		const toggleable = document.getElementById(elId);
		if (!toggleable) {
			return;
		}

		toggleEl(target, toggleable);

		const handleFocusout = (e) => {
			if (toggleable.contains(e.relatedTarget) || target.getAttribute("aria-expanded") !== "true") {
				return;
			}

			if (toggleable.contains(e.currentTarget) && target.contains(e.relatedTarget)) {
				// Exclude the case where the focus is moving from popup back to button (e.g. when closing the currently focused popup with a left-mouse click on the button)
				// The popup will be closed within the keydown handleEl(); failure to exclude this case will cause a bug where toggleEl() is called twice, making it difficult to close the popup.
				// https://civicplus.tpondemand.com/entity/182413-211-calendar-month-view-remove-menu
				return;
			}

			toggleEl(target, toggleable);
		};

		if (!target["data-cp-toggle"]) {
			target["data-cp-toggle"] = true;
			toggleable.addEventListener("focusout", handleFocusout);
			target.addEventListener("focusout", handleFocusout);
		}

		const closeActionEl = toggleable.querySelector(".close");

		if (!closeActionEl) {
			return toggleable;
		}

		const handleClose = (e) => {
			if (e.type === "click" || (e.type === "keydown" && (e.key === "Enter" || e.key === " "))) {
				e.preventDefault();
				toggleEl(target, toggleable);
				target.focus();
			}
		};

		if (!closeActionEl["data-cp-toggle"]) {
			closeActionEl["data-cp-toggle"] = true;
			closeActionEl.addEventListener("click", handleClose);
			closeActionEl.addEventListener("keydown", handleClose);
		}

		return toggleable;
	};

	const handleDropdown = (target, func) => {
		const toggleable = handleEl(target);
		if (!toggleable) {
			return;
		}
		const focusables = Array.from(toggleable.querySelectorAll(focusableSelector));
		if (focusables.length === 0) {
			return;
		}
		focusables.forEach( (item, index) => {
			item.setAttribute("tabindex", -!!index|0);
		});
		if (!target["data-cp-toggle-navkeys"]) {
			target["data-cp-toggle-navkeys"] = true;
			target.addEventListener("keydown", e => {
				if (e.key === "ArrowDown" && target.getAttribute("aria-expanded") === "true") {
					focusables[0].focus();
				}
			});
			toggleable.addEventListener("keydown", e => {
				func(e, focusables);
			});
		}
	}

	const handleDropdown1D = (target) => {
		return handleDropdown(target, handleCpNavKeydown);
	}

	const handleDropdown2D = (target) => {
		return handleDropdown(target, handleCpNav2DKeydown);
	}

	const handlePopover = (target) => {
		handleEl(target);
		repositionPopover(target);
	};

	/**
	 * From the given toggle button's aria-controls elements returns active and total panel count
	 * @target  {Element} foo [description]
	 */
	const getPanels = (target) => {
		const elIds = target.getAttribute("aria-controls");
		const panelIds = elIds.split(' ');
		let activePanels = 0;
		let totalPanels = 0;
		for (const panelId of panelIds) {
			const panel = document.getElementById(panelId);
			if (!panel) {
				continue
			}
			totalPanels += 1
			activePanels += Number(!!( panel.offsetWidth || panel.offsetHeight || panel.getClientRects().length ))
		}

		return [activePanels, totalPanels];
	}

	/**
	 * Get total height of the given element's children
	 * @collapsable  {Element} foo [description]
	 */
	const collapsableHeight = (collapsable) => {
		const totalHeight = Array.from(collapsable.children).reduce((acc, child) => acc + child.offsetHeight, 0);
		return `calc(${totalHeight}px + 2px)`
	}

	const handleCollapsar = (target) => {
		const elIds = target.getAttribute("aria-controls");
		if (!elIds) {
			return;
		}

		const collapsables = elIds.split(' ').filter(x => document.getElementById(x));

		const expand = target.getAttribute("aria-expanded") === "false";

		for (const collapsableId of collapsables) {

			const collapsable = document.getElementById(collapsableId);
			collapsable.style.overflow = "hidden";
			collapsable.style.transition = "all 300ms ease-in";

			// get all the buttons those are responsible for the given panel
			const controllers = document.querySelectorAll(`[aria-controls^="${collapsableId}"],[aria-controls*=" ${collapsableId} "],[aria-controls$="${collapsableId}"]`);

			if (expand) {
				collapsable.style.height = "0";
				collapsable.style.display = "block";
				collapsable.style.height = collapsableHeight(collapsable);
				let nextFocus = null;
				for (const controller of controllers) {
					const [activePanels, totalPanels] = getPanels(controller);
					if (activePanels !== totalPanels) {
						continue;
					}
					nextFocus = document.getElementById(controller.dataset.cpToggleFocus);
					controller.setAttribute("aria-expanded", "true");
					controller.classList.add("active");
					if (controller.dataset.cpToggleActiveTitle) {
						controller.innerHTML = controller.dataset.cpToggleActiveTitle;
					}
				}

				window.setTimeout(() => {
					collapsable.style.removeProperty("height");
					collapsable.style.removeProperty("overflow");
					nextFocus && nextFocus.focus();
				}, 300);
				continue;
			}

			collapsable.style.height = collapsableHeight(collapsable);
			collapsable.style.overflow = "hidden";
			window.setTimeout(() => {
				collapsable.style.height = "0";
			}, 0);

			window.setTimeout(function () {
				collapsable.style.display = "none";
				for (const controller of controllers) {
					const [activePanels, _] = getPanels(controller);
					if (activePanels !== 0) {
						continue;
					}
					controller.setAttribute("aria-expanded", "false");
					controller.classList.remove("active");
					if (controller.dataset.cpToggleInactiveTitle) {
						controller.innerHTML = controller.dataset.cpToggleInactiveTitle;
					}
				}
			}, 300);
		}
	};

	const handleCpToggles = (e) => {
		if (e.type === "click" || (e.type === "keydown" && (e.key === "Enter" || e.key === " "))) {
			const target = e.target.closest("[data-cp-toggle]");
			if (!target) {
				return;
			}
			e.preventDefault();
			// target.focus(); // focus element so that focusout is fired to close other open menus.

			const actionDescriptor = target.dataset.cpToggle;
			switch (actionDescriptor) {
				case "popover":
					handlePopover(target);
					break;
				case "collapse":
					handleCollapsar(target);
					break;
				case "dropdown":
					handleDropdown1D(target);
					break
				case "dropdown-2d":
					handleDropdown2D(target);
					break
				default:
					handleEl(target);
			}
		}

		const escTargets = new Set(["dropdown", "dropdown-2d"]);

		if (e.type === "keydown" && e.key === "Escape") {
			const toggles = document.querySelectorAll('[data-cp-toggle][aria-expanded="true"]');
			for (const target of toggles) {
				const toggleType = target.dataset.cpToggle;
				if (escTargets.has(toggleType)) {
					const elId = target.getAttribute("aria-controls");
					if (!elId) {
						continue;
					}

					const toggleable = document.getElementById(elId);
					if (!toggleable) {
						continue;
					}

					const refocus = toggleable.contains(document.activeElement);

					toggleEl(target, toggleable);

					if (refocus) {
						target.focus();
					}
				}
			}
		}
	};

	document.addEventListener("click", handleCpToggles);
	document.addEventListener("keydown", handleCpToggles);
})();
;
/* Responsive Menu Font Sizing
 * Joshua Enfield
 *  -------------
 * This resizes menus responsively. It takes into account the original font size.
 * If auto widths are on font sizes are adjusted based on the size left in menu items. 
 * If auto widths are off font sizes are adjusted based on the size left in the main nav.
 * ---------------
 * Common Usage: $('menuOLSelector').responsiveMenuText();
 * jQuery Usage expects markup in the form of ol > li > a
 */
(function ($, window, undefined) {

	'use strict';

	var DEBUG = false;
	var diag = function (msg) {
		if (DEBUG) {
			window.console.log('[ Responsive Menus ] ' + msg);
		}
	};
	
	$.MenuTextResizer = function (options, element) {
		this.$mainNavMenu = $(element);
		this.$topLevelMenuItems = this.$mainNavMenu.find('a');
		this._init(options);
	};

	// the options
	$.MenuTextResizer.defaults = {
		ratio: 0.5,
		timeout: 120,
		numberOfFontSizes: 20,
		tolerance: 0.00
	};

	$.MenuTextResizer.prototype = {
		_init: function (options) {
			// options
			this.options = $.extend(true, {}, $.MenuTextResizer.defaults, options);

			// local variables
			var self = this;
			this.timer = null;
			this.autoCreatedWrapper = false;
			createInnerWrappersIfWrappersDoNotExist();
			var desktopFontSize = parseFloat(this.$topLevelMenuItems.css('font-size')); // Initial Font Size. When at maximum size font-size should be this.
			var numberOfFontSizes = this.options.numberOfFontSizes;
			var fontSizes = getFontSizes();
			var currentFontIndex = fontSizes.length - 1;
			var tolerance = this.options.tolerance;
			var autoWidths = this.$mainNavMenu.data('autowidth') == true;

			var ratioToTriggerResize = autoWidths ? { menuRatio: 1.00, itemRatio: self.options.ratio } : { menuRatio: self.options.ratio, itemRatio: 1.00 };
			//var ratioToTriggerResize = { menuRatio: 0.68, itemRatio: 1.00 };
			//var ratioToTriggerResize = { menuRatio: 1.00, itemRatio: 0.70 };

			if (autoWidths)
				diag('Auto Widths Detected.');
			else {
				diag('Auto Widths Turned Off.');
			}

			diag('Font Sizes: ' + fontSizes);

			// Gets list of font sizes to resize amongst

			function getFontSizes () {
				var fontSizes = [];

				// Linear Scale
				for (var i = 0; i < numberOfFontSizes; i++) {
					fontSizes[i] = Math.floor(desktopFontSize / numberOfFontSizes * (i + 1));
				}

				// Exponential Decay
				/*
				fontSizes[numberOfFontSizes - 1] = desktopFontSize;
				for (var i = (numberOfFontSizes-2); i > 0; i--) {
					fontSizes[i] = Math.floor(0.95 * fontSizes[i+1]);
				}*/

				return fontSizes;
			}

			// Increases font size.
			var increaseFontSize = function () {
				var wasFontSizeIncreased = false;

				if (currentFontIndex < (fontSizes.length - 1)) {
					currentFontIndex++;
					wasFontSizeIncreased = true;
					// ratioToTriggerResize.itemRatio -= 0.1;
				}

				self.$topLevelMenuItems.css('font-size', fontSizes[currentFontIndex]);
				diag('INCREASE IN SIZE TRIGGERED (sizeIndex: ' + currentFontIndex + (currentFontIndex == (numberOfFontSizes - 1) ? '(MAX)' : '') + ' ,' + fontSizes[currentFontIndex] + ')');
				return wasFontSizeIncreased;
			};

			// Decreases font size.
			var decreaseFontSize = function () {
				var wasFontSizeDecreased = false;

				if (currentFontIndex > 0) {
					currentFontIndex--;
					wasFontSizeDecreased = true;
					//  ratioToTriggerResize.itemRatio += 0.1;
				}

				self.$topLevelMenuItems.css('font-size', fontSizes[currentFontIndex]);

				diag('DECREASE IN SIZE TRIGGERED (sizeIndex: ' + currentFontIndex + ',' + fontSizes[currentFontIndex] + ')');
				return wasFontSizeDecreased;
			};

			// Wraps menu Items with a span if spans are not already wrapping them.
			function createInnerWrappersIfWrappersDoNotExist () {
				if (self.$topLevelMenuItems.find('span').length === 0) {
					self.$topLevelMenuItems.removeAttr('style').wrapInner('<span/>').find('span').css('display', 'inline');
					self.autoCreatedWrapper = true;
				}
			}

			// Get the ratios of space taken to space available for menu items and the menu itself
			function getRatio () {
				var getItemRatio = function ($item) { return $item.find('span').width() / $item.width(); };
				var itemRatio = getItemRatio(self.$topLevelMenuItems.first());

				var widthConsumedByMenuItems = 0;
				self.$topLevelMenuItems.each(function () {
					widthConsumedByMenuItems += $(this).outerWidth(true);

					var currentItemRatio = getItemRatio($(this));
					if (currentItemRatio > itemRatio) {
						itemRatio = currentItemRatio;
					}

				});

				var menuRatio = widthConsumedByMenuItems / self.$mainNavMenu.width();
				var result = {
					menuRatio: menuRatio.toFixed(2),
					itemRatio: itemRatio.toFixed(2)
				};

				return result;
			}

			// Conditions
			var predicates = (function () {
				// Should return true when an increase in size should occur.
				var shouldIncreaseInSize = function () {
					var currentRatio = getRatio();
					diag(JSON.stringify(currentRatio));
					if (autoWidths) {
						return currentRatio.itemRatio < (ratioToTriggerResize.itemRatio - tolerance);
					} else {
						return currentRatio.menuRatio < (ratioToTriggerResize.menuRatio - tolerance);
					}
				};

				// Should return true when a decrease in size should occur.
				var shouldDescreaseInSize = function () {
					var currentRatio = getRatio();
					diag(JSON.stringify(currentRatio));
					if (autoWidths) {
						return currentRatio.itemRatio > (ratioToTriggerResize.itemRatio - tolerance);
					} else {
						return currentRatio.menuRatio > (ratioToTriggerResize.menuRatio - tolerance);
					}
				};

				return {
					shouldIncreaseInSize: shouldIncreaseInSize,
					shouldDescreaseInSize: shouldDescreaseInSize
				};

			}());

			// Function that actually adjusts menu font sizes.
			this.adjustMenuFontSize = function () {
				while (predicates.shouldIncreaseInSize() && increaseFontSize()) ;
				while (predicates.shouldDescreaseInSize() && decreaseFontSize()) ;
			};

			this._initDefault();
		},
		_initDefault : function () {
			// Setup and bind menu resizing logic
			var self = this;
			var timeout = self.options.timeout || 100; //debouncing
			$(window).bind('resize.responsiveMenuText', function () {
				self.timer && clearTimeout(self.timer);
				
				self.timer = setTimeout(function () {
					self.adjustMenuFontSize();
				}, timeout);
			});
			self.adjustMenuFontSize();
		},
		dispose: function () {
			var self = this;
			clearTimeout(self.timer);
			$(window).unbind('.responsiveMenuText');
			if (self.autoCreatedWrapper) {
				self.$topLevelMenuItems.find('span').contents().unwrap();
				setTimeout(function () { self.$topLevelMenuItems.removeAttr('style'); }, 200);
			}
		}
	};

	/* jQuery Plugin Support
	   Expected Markup structure: ol > li > a
	   Menu Text is dynamically wrapped by a span if it isn't already.
	*/
	$.fn.responsiveMenuText = function (options) {
		if (typeof options === 'string') {
			var args = Array.prototype.slice.call(arguments, 1);
			this.each(function () {
				var instance = $.data(this, 'responsiveMenuText');
				if (!instance) {
					diag("cannot call methods on responsiveMenuText prior to initialization; " +
					"attempted to call method '" + options + "'");
					return;
				}
				if (!$.isFunction(instance[options]) || options.charAt(0) === "_") {
					diag("no such method '" + options + "' for responsiveMenuText instance");
					return;
				}
				instance[options].apply(instance, args);
			});
		}
		else {
			this.each(function () {
				var instance = $.data(this, 'responsiveMenuText');
				if (instance) {
					instance._init(options);
				}
				else {
					instance = $.data(this, 'responsiveMenuText', new $.MenuTextResizer(options, this));
				}
			});
		}

		return this;
	};

})(jQuery, window);
/**
 * Based on  jquery.rwdmenu.js v1.0.1
 * http://www.codrops.com
 *
 */
; (function ($, window, undefined) {

	'use strict';

	// global
	var Modernizr = window.Modernizr, $body = $('body');

	$.RWDMenu = function (options, element) {
		this.$el = $(element);
		this._init(options);
	};

	// the options
	$.RWDMenu.defaults = {
		// classes for the animation effects
		animationClasses: { classin: 'rwd-animate-in', classout: 'rwd-animate-out' },
		// callback: click a link that has a sub menu
		// el is the link element (li); name is the level name
		onLevelClick: function (el, name) { return false; },
		// callback: click a link that does not have a sub menu
		// el is the link element (li); ev is the event obj
		onLinkClick: function (el, ev) { return false; }
	};

	$.RWDMenu.prototype = {
		_init: function (options) {

			// options
			this.options = $.extend(true, {}, $.RWDMenu.defaults, options);
			// cache some elements and initialize some variables
			this._config();

			var animEndEventNames = {
				'WebkitAnimation': 'webkitAnimationEnd',
				'OAnimation': 'oAnimationEnd',
				'msAnimation': 'MSAnimationEnd',
				'animation': 'animationend'
			},
				transEndEventNames = {
					'WebkitTransition': 'webkitTransitionEnd',
					'MozTransition': 'transitionend',
					'OTransition': 'oTransitionEnd',
					'msTransition': 'MSTransitionEnd',
					'transition': 'transitionend'
				};
			// animation end event name
			this.animEndEventName = animEndEventNames[Modernizr.prefixed('animation')] + '.rwdmenu';
			// transition end event name
			this.transEndEventName = transEndEventNames[Modernizr.prefixed('transition')] + '.rwdmenu',
			// support for css animations and css transitions
			this.supportAnimations = Modernizr.cssanimations,
			this.supportTransitions = Modernizr.csstransitions;

			var self = this;
			this.clickMenuItem = function (event) {
				//if (!menuManager.mobileMainNav) return true;

				event.stopPropagation();

				var $li = $(this),
					$submenu = $li.children('ol.rwd-submenu');

				if ($submenu.length == 0) {
					var id = $li.attr('id'); //.replace(/item|leftItem/i, '');
					var $olSubMenu = menuManager.$divListSubMenus.find('#Parent' + id);
					if ($olSubMenu.length > 1) {
						//remove duplicates. they might have been added by different menu (main/side)
						$olSubMenu.slice(1).remove();
						$olSubMenu = menuManager.$divListSubMenus.find('#Parent' + id);
					}
					if ($olSubMenu.length == 1) {
						var $back = $('<li class="rwd-back"><a href="#" style="z-index:2;" class="navMenuItem backMenuLink" aria-haspopup="true" role="menuitem">Back</a></li>');
						$back.bind('click.rwdmenu', self.clickBack);
						$olSubMenu.prepend($back);
						$olSubMenu.addClass('rwd-submenu');
						$olSubMenu.find('li:not(.rwd-back)').bind('click.rwdmenu', self.clickMenuItem);
						$li.append($olSubMenu);
						$submenu = $olSubMenu;
					}
				}

				if ($submenu.length > 0) {

					var $flyin = $submenu.clone().css('opacity', 0).insertAfter(self.$menu),
						onAnimationEndFn = function () {
							self.$menu.unbind(self.animEndEventName).removeClass(self.options.animationClasses.classout).addClass('rwd-subview');
							$('.rwd-subviewopen').children('ol').attr('aria-hidden', 'true');
							$li.addClass('rwd-subviewopen').parents('.rwd-subviewopen:first').removeClass('rwd-subviewopen').addClass('rwd-subview');
							$li.children('ol').attr('aria-hidden', 'false');
							$flyin.remove();
							$li.find('.openThePageLink').focus();
						};

					setTimeout(function () {
						$flyin.addClass(self.options.animationClasses.classin);
						self.$menu.addClass(self.options.animationClasses.classout);
						if (self.supportAnimations) {
							self.$menu.bind(self.animEndEventName, onAnimationEndFn);
						} else {
							onAnimationEndFn.call();
						}

						self.options.onLevelClick($li, $li.children('a:first').text());
					});

					return false;

				} else {
					self.options.onLinkClick($li, event);
				}

			};

			this.clickBack = function (event) {

				var $this = $(this),
					$submenu = $this.parents('ol.rwd-submenu:first'),
					$item = $submenu.parent(),
					$flyin = $submenu.clone().insertAfter(self.$menu);

				var onAnimationEndFn = function () {
					self.$menu.unbind(self.animEndEventName).removeClass(self.options.animationClasses.classin);
					$flyin.remove();
				};

				setTimeout(function () {
					$flyin.addClass(self.options.animationClasses.classout);
					self.$menu.addClass(self.options.animationClasses.classin);
					if (self.supportAnimations) {
						self.$menu.bind(self.animEndEventName, onAnimationEndFn);
					} else {
						onAnimationEndFn.call();
					}

					$item.removeClass('rwd-subviewopen').children('ol').attr('aria-hidden', 'true');

					var $subview = $this.parents('.rwd-subview:first');
					if ($subview.is('li')) {
						$subview.addClass('rwd-subviewopen').children('ol').attr('aria-hidden', 'false');
					}
					$subview.removeClass('rwd-subview');
					$('ol#mainNavMenu').find('li > a').first().focus();
				});

				return false;
			};

			this._initEvents();

		},
		_config: function () {
			this.$menu = this.$el;// this.$el.children('ol.rwd-menu');
			this.$menuitems = this.$menu.find('li.topMenuItem');
			this.$menu.addClass('rwd-menu');
		},
		_initEvents: function () {

			var self = this;

			this.$menuitems.bind('click.rwdmenu', self.clickMenuItem);
		},
		destroy: function () {
			//Remove classes, other events, not needed lis (back, to parent) and place back into hidden div.
			this.$menu.removeClass('rwd-menu rwd-subview');
			this.$menu.find('li').removeClass('rwd-subviewopen rwd-subview').unbind('.rwdmenu');
			this.$menu.find('li.rwd-back').remove();
			this.$menu.find('li.toParent').remove();
			this.$menu.find('ol').removeClass('rwd-submenu').each(function () {
				menuManager.$divListSubMenus.append($(this));
			});
		}
	};

	var logError = function (message) {
		if (window.console) {
			window.console.error(message);
		}
	};

	$.fn.rwdmenu = function (options) {
		if (typeof options === 'string') {
			var args = Array.prototype.slice.call(arguments, 1);
			this.each(function () {
				var instance = $.data(this, 'rwdmenu');
				if (!instance) {
					logError("cannot call methods on rwdmenu prior to initialization; " +
					"attempted to call method '" + options + "'");
					return;
				}
				if (!$.isFunction(instance[options]) || options.charAt(0) === "_") {
					logError("no such method '" + options + "' for rwdmenu instance");
					return;
				}
				instance[options].apply(instance, args);
			});
		}
		else {
			this.each(function () {
				var instance = $.data(this, 'rwdmenu');
				if (instance) {
					instance._init();
				}
				else {
					instance = $.data(this, 'rwdmenu', new $.RWDMenu(options, this));
				}
			});
		}
		return this;
	};

})(jQuery, window);
/*global jQuery:false, alert:false */

/*
 * Default text - jQuery plugin for html5 dragging files from desktop to browser
 *
 * Author: Weixi Yen
 *
 * Email: [Firstname][Lastname]@gmail.com
 *
 * Copyright (c) 2010 Resopollution
 *
 * Licensed under the MIT license:
 *   http://www.opensource.org/licenses/mit-license.php
 *
 * Project home:
 *   http://www.github.com/weixiyen/jquery-filedrop
 *
 * Version:  0.1.0
 *
 * Features:
 *      Allows sending of extra parameters with file.
 *      Works with Firefox 3.6+
 *      Future-compliant with HTML5 spec (will work with Webkit browsers and IE9)
 * Usage:
 *  See README at project homepage
 *
 */
;(function($) {

  var default_opts = {
        fallback_id: '',
        fallback_dropzoneClick : true,
        url: '',
        refresh: 1000,
        paramname: 'userfile',
        requestType: 'POST',    // just in case you want to use another HTTP verb
        allowedfileextensions:[],
        allowedfiletypes:[],
        maxfiles: 25,           // Ignored if queuefiles is set > 0
        maxfilesize: 1,         // MB file size limit
        queuefiles: 0,          // Max files before queueing (for large volume uploads)
        queuewait: 200,         // Queue wait time if full
        data: {},
        headers: {},
        drop: empty,
        dragStart: empty,
        dragEnter: empty,
        dragOver: empty,
        dragLeave: empty,
        docEnter: empty,
        docOver: empty,
        docLeave: empty,
        beforeEach: empty,
        afterAll: empty,
        rename: empty,
        error: function(err, file, i, status) {
          alert(err);
        },
        uploadStarted: empty,
        uploadFinished: empty,
        progressUpdated: empty,
        globalProgressUpdated: empty,
        speedUpdated: empty
      },
      errors = ["BrowserNotSupported", "TooManyFiles", "FileTooLarge", "FileTypeNotAllowed", "NotFound", "NotReadable", "AbortError", "ReadError", "FileExtensionNotAllowed"];

  $.fn.filedrop = function(options) {
    var opts = $.extend({}, default_opts, options),
        global_progress = [],
        doc_leave_timer, stop_loop = false,
        files_count = 0,
        files;

    if ( opts.fallback_id )
    {
      if ( opts.fallback_dropzoneClick === true )
      {
        $('#' + opts.fallback_id).css({
          display: 'none',
          width: 0,
          height: 0
        });
      }




      if ( opts.fallback_dropzoneClick === true )
      {
        if ( this.find('#' + opts.fallback_id).length > 0 )
        {
          throw "Fallback element ["+opts.fallback_id+"] cannot be inside dropzone, unless option fallback_dropzoneClick is false";
        }
        else
        {
          this.on('click', function(e){
            $('#' + opts.fallback_id).trigger(e);
          });
        }
      }

      $('#' + opts.fallback_id).change(function(e) {
        opts.drop(e);
        files = e.target.files;
        files_count = files.length;
        upload();
      });
    }

    this.on('drop', drop).on('dragstart', opts.dragStart).on('dragenter', dragEnter).on('dragover', dragOver).on('dragleave', dragLeave);
    $(document).on('drop', docDrop).on('dragenter', docEnter).on('dragover', docOver).on('dragleave', docLeave);

    function drop(e) {
      if( opts.drop.call(this, e) === false ) return false;
      if(!e.originalEvent.dataTransfer)
        return;
      files = e.originalEvent.dataTransfer.files;
      if (files === null || files === undefined || files.length === 0) {
        opts.error(errors[0]);
        return false;
      }
      files_count = files.length;
      upload();
      e.preventDefault();
      return false;
    }

    function getBuilder(filename, filedata, mime, boundary) {
      var dashdash = '--',
          crlf = '\r\n',
          builder = '',
          paramname = opts.paramname;

      if (opts.data) {
        var params = $.param(opts.data).replace(/\+/g, '%20').split(/&/);

        $.each(params, function() {
          var pair = this.split("=", 2),
              name = decodeURIComponent(pair[0]),
              val  = decodeURIComponent(pair[1]);

          if (pair.length !== 2) {
            return;
          }

          builder += dashdash;
          builder += boundary;
          builder += crlf;
          builder += 'Content-Disposition: form-data; name="' + name + '"';
          builder += crlf;
          builder += crlf;
          builder += val;
          builder += crlf;
        });
      }

      if (jQuery.isFunction(paramname)){
        paramname = paramname(filename);
      }

      builder += dashdash;
      builder += boundary;
      builder += crlf;
      builder += 'Content-Disposition: form-data; name="' + (paramname||"") + '"';
      builder += '; filename="' + encodeURIComponent(filename) + '"';
      builder += crlf;

      builder += 'Content-Type: ' + mime;
      builder += crlf;
      builder += crlf;

      builder += filedata;
      builder += crlf;

      builder += dashdash;
      builder += boundary;
      builder += dashdash;
      builder += crlf;
      return builder;
    }

    function progress(e) {
      if (e.lengthComputable) {
        var percentage = Math.round((e.loaded * 100) / e.total);
        if (this.currentProgress !== percentage) {

          this.currentProgress = percentage;
          opts.progressUpdated(this.index, this.file, this.currentProgress);

          global_progress[this.global_progress_index] = this.currentProgress;
          globalProgress();

          var elapsed = new Date().getTime();
          var diffTime = elapsed - this.currentStart;
          if (diffTime >= opts.refresh) {
            var diffData = e.loaded - this.startData;
            var speed = diffData / diffTime; // KB per second
            opts.speedUpdated(this.index, this.file, speed);
            this.startData = e.loaded;
            this.currentStart = elapsed;
          }
        }
      }
    }

    function globalProgress() {
      if (global_progress.length === 0) {
        return;
      }

      var total = 0, index;
      for (index in global_progress) {
        if(global_progress.hasOwnProperty(index)) {
          total = total + global_progress[index];
        }
      }

      opts.globalProgressUpdated(Math.round(total / global_progress.length));
    }

    // Respond to an upload
    function upload() {
      stop_loop = false;

      if (!files) {
        opts.error(errors[0]);
        return false;
      }

      if (opts.allowedfiletypes.push && opts.allowedfiletypes.length) {
        for(var fileIndex = files.length;fileIndex--;) {
          if(!files[fileIndex].type || $.inArray(files[fileIndex].type, opts.allowedfiletypes) < 0) {
            opts.error(errors[3], files[fileIndex]);
            return false;
          }
        }
      }

      if (opts.allowedfileextensions.push && opts.allowedfileextensions.length) {
        for(var fileIndex = files.length;fileIndex--;) {
          var allowedextension = false;
          for (i=0;i<opts.allowedfileextensions.length;i++){
            if (files[fileIndex].name.substr(files[fileIndex].name.length-opts.allowedfileextensions[i].length).toLowerCase()
                == opts.allowedfileextensions[i].toLowerCase()
            ) {
              allowedextension = true;
            }
          }
          if (!allowedextension){
            opts.error(errors[8], files[fileIndex]);
            return false;
          }
        }
      }

      var filesDone = 0,
          filesRejected = 0;

      if (files_count > opts.maxfiles && opts.queuefiles === 0) {
        opts.error(errors[1]);
        return false;
      }

      // Define queues to manage upload process
      var workQueue = [];
      var processingQueue = [];
      var doneQueue = [];

      // Add everything to the workQueue
      for (var i = 0; i < files_count; i++) {
        workQueue.push(i);
      }

      // Helper function to enable pause of processing to wait
      // for in process queue to complete
      var pause = function(timeout) {
        setTimeout(process, timeout);
        return;
      };

      // Process an upload, recursive
      var process = function() {

        var fileIndex;

        if (stop_loop) {
          return false;
        }

        // Check to see if are in queue mode
        if (opts.queuefiles > 0 && processingQueue.length >= opts.queuefiles) {
          return pause(opts.queuewait);
        } else {
          // Take first thing off work queue
          fileIndex = workQueue[0];
          workQueue.splice(0, 1);

          // Add to processing queue
          processingQueue.push(fileIndex);
        }

        try {
          if (beforeEach(files[fileIndex]) !== false) {
            if (fileIndex === files_count) {
              return;
            }
            var reader = new FileReader(),
                max_file_size = 1048576 * opts.maxfilesize;

            reader.index = fileIndex;
            if (files[fileIndex].size > max_file_size) {
              opts.error(errors[2], files[fileIndex], fileIndex);
              // Remove from queue
              processingQueue.forEach(function(value, key) {
                if (value === fileIndex) {
                  processingQueue.splice(key, 1);
                }
              });
              filesRejected++;
              return true;
            }

            reader.onerror = function(e) {
              switch(e.target.error.code) {
                case e.target.error.NOT_FOUND_ERR:
                  opts.error(errors[4]);
                  return false;
                case e.target.error.NOT_READABLE_ERR:
                  opts.error(errors[5]);
                  return false;
                case e.target.error.ABORT_ERR:
                  opts.error(errors[6]);
                  return false;
                default:
                  opts.error(errors[7]);
                  return false;
              };
            };

            reader.onloadend = !opts.beforeSend ? send : function (e) {
              opts.beforeSend(files[fileIndex], fileIndex, function () { send(e); });
            };

            reader.readAsDataURL(files[fileIndex]);

          } else {
            filesRejected++;
          }
        } catch (err) {
          // Remove from queue
          processingQueue.forEach(function(value, key) {
            if (value === fileIndex) {
              processingQueue.splice(key, 1);
            }
          });
          opts.error(errors[0]);
          return false;
        }

        // If we still have work to do,
        if (workQueue.length > 0) {
          process();
        }
      };

      var send = function(e) {

        var fileIndex = (e.srcElement || e.target).index;

        // Sometimes the index is not attached to the
        // event object. Find it by size. Hack for sure.
        if (e.target.index === undefined) {
          e.target.index = getIndexBySize(e.total);
        }

        var xhr = new XMLHttpRequest(),
            upload = xhr.upload,
            file = files[e.target.index],
            index = e.target.index,
            start_time = new Date().getTime(),
            boundary = '------multipartformboundary' + (new Date()).getTime(),
            global_progress_index = global_progress.length,
            builder,
            newName = rename(file.name),
            mime = file.type;

        if (opts.withCredentials) {
          xhr.withCredentials = opts.withCredentials;
        }

        var encodedString = e.target.result.split(',')[1];
        var data = encodedString === undefined ? '' : atob(encodedString);
        if (typeof newName === "string") {
          builder = getBuilder(newName, data, mime, boundary);
        } else {
          builder = getBuilder(file.name, data, mime, boundary);
        }

        upload.index = index;
        upload.file = file;
        upload.downloadStartTime = start_time;
        upload.currentStart = start_time;
        upload.currentProgress = 0;
        upload.global_progress_index = global_progress_index;
        upload.startData = 0;
        upload.addEventListener("progress", progress, false);

        // Allow url to be a method
        if (jQuery.isFunction(opts.url)) {
          xhr.open(opts.requestType, opts.url(upload), true);
        } else {
          xhr.open(opts.requestType, opts.url, true);
        }

        xhr.setRequestHeader('content-type', 'multipart/form-data; boundary=' + boundary);
        xhr.setRequestHeader("X-Requested-With", "XMLHttpRequest");

        // Add headers
        $.each(opts.headers, function(k, v) {
          xhr.setRequestHeader(k, v);
        });

        if(!xhr.sendAsBinary){
          xhr.sendAsBinary = function(datastr) {
            function byteValue(x) {
              return x.charCodeAt(0) & 0xff;
            }
            var ords = Array.prototype.map.call(datastr, byteValue);
            var ui8a = new Uint8Array(ords);
            this.send(ui8a.buffer);
          }
        }

        xhr.sendAsBinary(builder);

        global_progress[global_progress_index] = 0;
        globalProgress();

        opts.uploadStarted(index, file, files_count);

        xhr.onload = function() {
          var serverResponse = null;

          if (xhr.responseText) {
            try {
              serverResponse = jQuery.parseJSON(xhr.responseText);
            }
            catch (e) {
              serverResponse = xhr.responseText;
            }
          }

          var now = new Date().getTime(),
              timeDiff = now - start_time,
              result = opts.uploadFinished(index, file, serverResponse, timeDiff, xhr);
          filesDone++;

          // Remove from processing queue
          processingQueue.forEach(function(value, key) {
            if (value === fileIndex) {
              processingQueue.splice(key, 1);
            }
          });

          // Add to donequeue
          doneQueue.push(fileIndex);

          // Make sure the global progress is updated
          global_progress[global_progress_index] = 100;
          globalProgress();

          if (filesDone === (files_count - filesRejected)) {
            afterAll();
          }
          if (result === false) {
            stop_loop = true;
          }


          // Pass any errors to the error option
          if (xhr.status < 200 || xhr.status > 299) {
            opts.error(xhr.statusText, file, fileIndex, xhr.status);
          }
        };
      };

      // Initiate the processing loop
      process();
    }

    function getIndexBySize(size) {
      for (var i = 0; i < files_count; i++) {
        if (files[i].size === size) {
          return i;
        }
      }

      return undefined;
    }

    function rename(name) {
      return opts.rename(name);
    }

    function beforeEach(file) {
      return opts.beforeEach(file);
    }

    function afterAll() {
      return opts.afterAll();
    }

    function dragEnter(e) {
      clearTimeout(doc_leave_timer);
      e.preventDefault();
      opts.dragEnter.call(this, e);
    }

    function dragOver(e) {
      clearTimeout(doc_leave_timer);
      e.preventDefault();
      opts.docOver.call(this, e);
      opts.dragOver.call(this, e);
    }

    function dragLeave(e) {
      clearTimeout(doc_leave_timer);
      opts.dragLeave.call(this, e);
      e.stopPropagation();
    }

    function docDrop(e) {
      e.preventDefault();
      opts.docLeave.call(this, e);
      return false;
    }

    function docEnter(e) {
      clearTimeout(doc_leave_timer);
      e.preventDefault();
      opts.docEnter.call(this, e);
      return false;
    }

    function docOver(e) {
      clearTimeout(doc_leave_timer);
      e.preventDefault();
      opts.docOver.call(this, e);
      return false;
    }

    function docLeave(e) {
      doc_leave_timer = setTimeout((function(_this) {
        return function() {
          opts.docLeave.call(_this, e);
        };
      })(this), 200);
    }

    return this;
  };

  function empty() {}

  try {
    if (XMLHttpRequest.prototype.sendAsBinary) {
      return;
    }
    XMLHttpRequest.prototype.sendAsBinary = function(datastr) {
      function byteValue(x) {
        return x.charCodeAt(0) & 0xff;
      }
      var ords = Array.prototype.map.call(datastr, byteValue);
      var ui8a = new Uint8Array(ords);

      // Not pretty: Chrome 22 deprecated sending ArrayBuffer, moving instead
      // to sending ArrayBufferView.  Sadly, no proper way to detect this
      // functionality has been discovered.  Happily, Chrome 22 also introduced
      // the base ArrayBufferView class, not present in Chrome 21.
      if ('ArrayBufferView' in window)
        this.send(ui8a);
      else
        this.send(ui8a.buffer);
    };
  } catch (e) {}

})(jQuery);
;
/*!
 * Javascript Cookie v1.5.1
 * https://github.com/js-cookie/js-cookie
 *
 * Copyright 2006, 2014 Klaus Hartl
 * Released under the MIT license
 */
(function(n){var i,r,t;if(typeof define=="function"&&define.amd)define(["jquery"],n);else if(typeof exports=="object"){try{i=require("jquery")}catch(u){}module.exports=n(i)}else r=window.Cookies,t=window.Cookies=n(window.jQuery),t.noConflict=function(){return window.Cookies=r,t}})(function(n){function i(n){return t.raw?n:encodeURIComponent(n)}function o(n){return t.raw?n:decodeURIComponent(n)}function s(n){return i(t.json?JSON.stringify(n):String(n))}function h(n){n.indexOf('"')===0&&(n=n.slice(1,-1).replace(/\\"/g,'"').replace(/\\\\/g,"\\"));try{return n=decodeURIComponent(n.replace(e," ")),t.json?JSON.parse(n):n}catch(i){}}function r(n,i){var r=t.raw?n:h(n);return f(i)?i(r):r}function u(){for(var n,t,i=0,r={};i<arguments.length;i++){t=arguments[i];for(n in t)r[n]=t[n]}return r}function f(n){return Object.prototype.toString.call(n)==="[object Function]"}var e=/\+/g,t=function(n,e,h){var y,l;if(arguments.length>1&&!f(e))return h=u(t.defaults,h),typeof h.expires=="number"&&(y=h.expires,l=h.expires=new Date,l.setMilliseconds(l.getMilliseconds()+y*864e5)),document.cookie=[i(n),"=",s(e),h.expires?"; expires="+h.expires.toUTCString():"",h.path?"; path="+h.path:"",h.domain?"; domain="+h.domain:"",h.secure?"; secure":""].join("");for(var a=n?undefined:{},p=document.cookie?document.cookie.split("; "):[],v=0,k=p.length;v<k;v++){var w=p[v].split("="),b=o(w.shift()),c=w.join("=");if(n===b){a=r(c,e);break}n||(c=r(c))===undefined||(a[b]=c)}return a};return t.get=t.set=t,t.defaults={},t.remove=function(n,i){return t(n,"",u(i,{expires:-1})),!t(n)},n&&(n.cookie=t,n.removeCookie=t.remove),t});
/*! jQuery UI - v1.14.2 - 2026-01-29
* https://jqueryui.com
* Includes: widget.js, position.js, data.js, disable-selection.js, effect.js, effects/effect-blind.js, effects/effect-bounce.js, effects/effect-clip.js, effects/effect-drop.js, effects/effect-explode.js, effects/effect-fade.js, effects/effect-fold.js, effects/effect-highlight.js, effects/effect-puff.js, effects/effect-pulsate.js, effects/effect-scale.js, effects/effect-shake.js, effects/effect-size.js, effects/effect-slide.js, effects/effect-transfer.js, focusable.js, form-reset-mixin.js, jquery-patch.js, keycode.js, labels.js, scroll-parent.js, tabbable.js, unique-id.js, widgets/accordion.js, widgets/autocomplete.js, widgets/button.js, widgets/checkboxradio.js, widgets/controlgroup.js, widgets/datepicker.js, widgets/dialog.js, widgets/draggable.js, widgets/droppable.js, widgets/menu.js, widgets/mouse.js, widgets/progressbar.js, widgets/resizable.js, widgets/selectable.js, widgets/selectmenu.js, widgets/slider.js, widgets/sortable.js, widgets/spinner.js, widgets/tabs.js, widgets/tooltip.js
* Copyright OpenJS Foundation and other contributors; Licensed MIT */

!function(t){"use strict";"function"==typeof define&&define.amd?define(["jquery"],t):t(jQuery)}(function(t){"use strict";t.ui=t.ui||{},t.ui.version="1.14.2";var e,i,s,n,o,a=0,r=Array.prototype.hasOwnProperty,l=Array.prototype.slice;h=t.cleanData,t.cleanData=function(e){var i,s,n;for(n=0;null!=(s=e[n]);n++)(i=t._data(s,"events"))&&i.remove&&t(s).triggerHandler("remove");h(e)},t.widget=function(e,i,s){var n,o,a,r={},l=e.split(".")[0];if("__proto__"===(e=e.split(".")[1])||"constructor"===e)return t.error("Invalid widget name: "+e);var h=l+"-"+e;return s||(s=i,i=t.Widget),Array.isArray(s)&&(s=t.extend.apply(null,[{}].concat(s))),t.expr.pseudos[h.toLowerCase()]=function(e){return!!t.data(e,h)},t[l]=t[l]||{},n=t[l][e],o=t[l][e]=function(t,e){if(!this||!this._createWidget)return new o(t,e);arguments.length&&this._createWidget(t,e)},t.extend(o,n,{version:s.version,_proto:t.extend({},s),_childConstructors:[]}),(a=new i).options=t.widget.extend({},a.options),t.each(s,function(t,e){if("function"!=typeof e){r[t]=e;return}r[t]=function(){function s(){return i.prototype[t].apply(this,arguments)}function n(e){return i.prototype[t].apply(this,e)}return function(){var t,i=this._super,o=this._superApply;return this._super=s,this._superApply=n,t=e.apply(this,arguments),this._super=i,this._superApply=o,t}}()}),o.prototype=t.widget.extend(a,{widgetEventPrefix:n&&a.widgetEventPrefix||e},r,{constructor:o,namespace:l,widgetName:e,widgetFullName:h}),n?(t.each(n._childConstructors,function(e,i){var s=i.prototype;t.widget(s.namespace+"."+s.widgetName,o,i._proto)}),delete n._childConstructors):i._childConstructors.push(o),t.widget.bridge(e,o),o},t.widget.extend=function(e){for(var i,s,n=l.call(arguments,1),o=0,a=n.length;o<a;o++)for(i in n[o])s=n[o][i],r.call(n[o],i)&&void 0!==s&&(t.isPlainObject(s)?e[i]=t.isPlainObject(e[i])?t.widget.extend({},e[i],s):t.widget.extend({},s):e[i]=s);return e},t.widget.bridge=function(e,i){var s=i.prototype.widgetFullName||e;t.fn[e]=function(n){var o="string"==typeof n,a=l.call(arguments,1),r=this;return o?this.length||"instance"!==n?this.each(function(){var i,o=t.data(this,s);return"instance"===n?(r=o,!1):o?"function"!=typeof o[n]||"_"===n.charAt(0)?t.error("no such method '"+n+"' for "+e+" widget instance"):(i=o[n].apply(o,a))!==o&&void 0!==i?(r=i&&i.jquery?r.pushStack(i.get()):i,!1):void 0:t.error("cannot call methods on "+e+" prior to initialization; attempted to call method '"+n+"'")}):r=void 0:(a.length&&(n=t.widget.extend.apply(null,[n].concat(a))),this.each(function(){var e=t.data(this,s);e?(e.option(n||{}),e._init&&e._init()):t.data(this,s,new i(n,this))})),r}},t.Widget=function(){},t.Widget._childConstructors=[],t.Widget.prototype={widgetName:"widget",widgetEventPrefix:"",defaultElement:"<div>",options:{classes:{},disabled:!1,create:null},_createWidget:function(e,i){i=t(i||this.defaultElement||this)[0],this.element=t(i),this.uuid=a++,this.eventNamespace="."+this.widgetName+this.uuid,this.bindings=t(),this.hoverable=t(),this.focusable=t(),this.classesElementLookup={},i!==this&&(t.data(i,this.widgetFullName,this),this._on(!0,this.element,{remove:function(t){t.target===i&&this.destroy()}}),this.document=t(i.style?i.ownerDocument:i.document||i),this.window=t(this.document[0].defaultView||this.document[0].parentWindow)),this.options=t.widget.extend({},this.options,this._getCreateOptions(),e),this._create(),this.options.disabled&&this._setOptionDisabled(this.options.disabled),this._trigger("create",null,this._getCreateEventData()),this._init()},_getCreateOptions:function(){return{}},_getCreateEventData:t.noop,_create:t.noop,_init:t.noop,destroy:function(){var e=this;this._destroy(),t.each(this.classesElementLookup,function(t,i){e._removeClass(i,t)}),this.element.off(this.eventNamespace).removeData(this.widgetFullName),this.widget().off(this.eventNamespace).removeAttr("aria-disabled"),this.bindings.off(this.eventNamespace)},_destroy:t.noop,widget:function(){return this.element},option:function(e,i){var s,n,o,a=e;if(0==arguments.length)return t.widget.extend({},this.options);if("string"==typeof e){if(a={},e=(s=e.split(".")).shift(),s.length){for(o=0,n=a[e]=t.widget.extend({},this.options[e]);o<s.length-1;o++)n[s[o]]=n[s[o]]||{},n=n[s[o]];if(e=s.pop(),1==arguments.length)return void 0===n[e]?null:n[e];n[e]=i}else{if(1==arguments.length)return void 0===this.options[e]?null:this.options[e];a[e]=i}}return this._setOptions(a),this},_setOptions:function(t){var e;for(e in t)this._setOption(e,t[e]);return this},_setOption:function(t,e){return"classes"===t&&this._setOptionClasses(e),this.options[t]=e,"disabled"===t&&this._setOptionDisabled(e),this},_setOptionClasses:function(e){var i,s,n;for(i in e)n=this.classesElementLookup[i],e[i]!==this.options.classes[i]&&n&&n.length&&(s=t(n.get()),this._removeClass(n,i),s.addClass(this._classes({element:s,keys:i,classes:e,add:!0})))},_setOptionDisabled:function(t){this._toggleClass(this.widget(),this.widgetFullName+"-disabled",null,!!t),t&&(this._removeClass(this.hoverable,null,"ui-state-hover"),this._removeClass(this.focusable,null,"ui-state-focus"))},enable:function(){return this._setOptions({disabled:!1})},disable:function(){return this._setOptions({disabled:!0})},_classes:function(e){var i=[],s=this;function n(n,o){var a,r;for(r=0;r<n.length;r++)a=s.classesElementLookup[n[r]]||t(),e.add?(function(){var i=[];e.element.each(function(e,n){t.map(s.classesElementLookup,function(t){return t}).some(function(t){return t.is(n)})||i.push(n)}),s._on(t(i),{remove:"_untrackClassesElement"})}(),a=t(t.uniqueSort(a.get().concat(e.element.get())))):a=t(a.not(e.element).get()),s.classesElementLookup[n[r]]=a,i.push(n[r]),o&&e.classes[n[r]]&&i.push(e.classes[n[r]])}return(e=t.extend({element:this.element,classes:this.options.classes||{}},e)).keys&&n(e.keys.match(/\S+/g)||[],!0),e.extra&&n(e.extra.match(/\S+/g)||[]),i.join(" ")},_untrackClassesElement:function(e){var i=this;t.each(i.classesElementLookup,function(s,n){-1!==t.inArray(e.target,n)&&(i.classesElementLookup[s]=t(n.not(e.target).get()))}),this._off(t(e.target))},_removeClass:function(t,e,i){return this._toggleClass(t,e,i,!1)},_addClass:function(t,e,i){return this._toggleClass(t,e,i,!0)},_toggleClass:function(t,e,i,s){s="boolean"==typeof s?s:i;var n="string"==typeof t||null===t,o={extra:n?e:i,keys:n?t:e,element:n?this.element:t,add:s};return o.element.toggleClass(this._classes(o),s),this},_on:function(e,i,s){var n,o=this;"boolean"!=typeof e&&(s=i,i=e,e=!1),s?(i=n=t(i),this.bindings=this.bindings.add(i)):(s=i,i=this.element,n=this.widget()),t.each(s,function(s,a){function r(){if(!(!e&&(!0===o.options.disabled||t(this).hasClass("ui-state-disabled"))))return("string"==typeof a?o[a]:a).apply(o,arguments)}"string"!=typeof a&&(r.guid=a.guid=a.guid||r.guid||t.guid++);var l=s.match(/^([\w:-]*)\s*(.*)$/),h=l[1]+o.eventNamespace,c=l[2];c?n.on(h,c,r):i.on(h,r)})},_off:function(e,i){i=(i||"").split(" ").join(this.eventNamespace+" ")+this.eventNamespace,e.off(i),this.bindings=t(this.bindings.not(e).get()),this.focusable=t(this.focusable.not(e).get()),this.hoverable=t(this.hoverable.not(e).get())},_delay:function(t,e){var i=this;return setTimeout(function(){return("string"==typeof t?i[t]:t).apply(i,arguments)},e||0)},_hoverable:function(e){this.hoverable=this.hoverable.add(e),this._on(e,{mouseenter:function(e){this._addClass(t(e.currentTarget),null,"ui-state-hover")},mouseleave:function(e){this._removeClass(t(e.currentTarget),null,"ui-state-hover")}})},_focusable:function(e){this.focusable=this.focusable.add(e),this._on(e,{focusin:function(e){this._addClass(t(e.currentTarget),null,"ui-state-focus")},focusout:function(e){this._removeClass(t(e.currentTarget),null,"ui-state-focus")}})},_trigger:function(e,i,s){var n,o,a=this.options[e];if(s=s||{},(i=t.Event(i)).type=(e===this.widgetEventPrefix?e:this.widgetEventPrefix+e).toLowerCase(),i.target=this.element[0],o=i.originalEvent)for(n in o)n in i||(i[n]=o[n]);return this.element.trigger(i,s),!("function"==typeof a&&!1===a.apply(this.element[0],[i].concat(s))||i.isDefaultPrevented())}},t.each({show:"fadeIn",hide:"fadeOut"},function(e,i){t.Widget.prototype["_"+e]=function(s,n,o){"string"==typeof n&&(n={effect:n});var a,r=n?!0===n||"number"==typeof n?i:n.effect||i:e;"number"==typeof(n=n||{})?n={duration:n}:!0===n&&(n={}),a=!t.isEmptyObject(n),n.complete=o,n.delay&&s.delay(n.delay),a&&t.effects&&t.effects.effect[r]?s[e](n):r!==e&&s[r]?s[r](n.duration,n.easing,o):s.queue(function(i){t(this)[e](),o&&o.call(s[0]),i()})}}),t.widget,function(){var e,i=Math.max,s=Math.abs,n=/left|center|right/,o=/top|center|bottom/,a=/[\+\-]\d+(\.[\d]+)?%?/,r=/^\w+/,l=/%$/,h=t.fn.position;function c(t,e,i){return[parseFloat(t[0])*(l.test(t[0])?e/100:1),parseFloat(t[1])*(l.test(t[1])?i/100:1)]}function u(e,i){return parseInt(t.css(e,i),10)||0}function d(t){return null!=t&&t===t.window}t.position={scrollbarWidth:function(){if(void 0!==e)return e;var i,s,n=t("<div style='display:block;position:absolute;width:200px;height:200px;overflow:hidden;'><div style='height:300px;width:auto;'></div></div>"),o=n.children()[0];return t("body").append(n),i=o.offsetWidth,n.css("overflow","scroll"),i===(s=o.offsetWidth)&&(s=n[0].clientWidth),n.remove(),e=i-s},getScrollInfo:function(e){var i=e.isWindow||e.isDocument?"":e.element.css("overflow-x"),s=e.isWindow||e.isDocument?"":e.element.css("overflow-y"),n="scroll"===i||"auto"===i&&e.width<e.element[0].scrollWidth;return{width:"scroll"===s||"auto"===s&&e.height<e.element[0].scrollHeight?t.position.scrollbarWidth():0,height:n?t.position.scrollbarWidth():0}},getWithinInfo:function(e){var i=t(e||window),s=d(i[0]),n=!!i[0]&&9===i[0].nodeType;return{element:i,isWindow:s,isDocument:n,offset:s||n?{left:0,top:0}:t(e).offset(),scrollLeft:i.scrollLeft(),scrollTop:i.scrollTop(),width:i.outerWidth(),height:i.outerHeight()}}},t.fn.position=function(e){if(!e||!e.of)return h.apply(this,arguments);var l,p,f,g,m,_,v,b="string"==typeof(e=t.extend({},e)).of?t(document).find(e.of):t(e.of),y=t.position.getWithinInfo(e.within),w=t.position.getScrollInfo(y),x=(e.collision||"flip").split(" "),k={};return v=9===(l=b[0]).nodeType?{width:b.width(),height:b.height(),offset:{top:0,left:0}}:d(l)?{width:b.width(),height:b.height(),offset:{top:b.scrollTop(),left:b.scrollLeft()}}:l.preventDefault?{width:0,height:0,offset:{top:l.pageY,left:l.pageX}}:{width:b.outerWidth(),height:b.outerHeight(),offset:b.offset()},b[0].preventDefault&&(e.at="left top"),f=v.width,g=v.height,m=v.offset,_=t.extend({},m),t.each(["my","at"],function(){var t,i,s=(e[this]||"").split(" ");1===s.length&&(s=n.test(s[0])?s.concat(["center"]):o.test(s[0])?["center"].concat(s):["center","center"]),s[0]=n.test(s[0])?s[0]:"center",s[1]=o.test(s[1])?s[1]:"center",t=a.exec(s[0]),i=a.exec(s[1]),k[this]=[t?t[0]:0,i?i[0]:0],e[this]=[r.exec(s[0])[0],r.exec(s[1])[0]]}),1===x.length&&(x[1]=x[0]),"right"===e.at[0]?_.left+=f:"center"===e.at[0]&&(_.left+=f/2),"bottom"===e.at[1]?_.top+=g:"center"===e.at[1]&&(_.top+=g/2),p=c(k.at,f,g),_.left+=p[0],_.top+=p[1],this.each(function(){var n,o,a=t(this),r=a.outerWidth(),l=a.outerHeight(),h=u(this,"marginLeft"),d=u(this,"marginTop"),v=r+h+u(this,"marginRight")+w.width,C=l+d+u(this,"marginBottom")+w.height,D=t.extend({},_),I=c(k.my,a.outerWidth(),a.outerHeight());"right"===e.my[0]?D.left-=r:"center"===e.my[0]&&(D.left-=r/2),"bottom"===e.my[1]?D.top-=l:"center"===e.my[1]&&(D.top-=l/2),D.left+=I[0],D.top+=I[1],n={marginLeft:h,marginTop:d},t.each(["left","top"],function(i,s){t.ui.position[x[i]]&&t.ui.position[x[i]][s](D,{targetWidth:f,targetHeight:g,elemWidth:r,elemHeight:l,collisionPosition:n,collisionWidth:v,collisionHeight:C,offset:[p[0]+I[0],p[1]+I[1]],my:e.my,at:e.at,within:y,elem:a})}),e.using&&(o=function(t){var n=m.left-D.left,o=n+f-r,h=m.top-D.top,c=h+g-l,u={target:{element:b,left:m.left,top:m.top,width:f,height:g},element:{element:a,left:D.left,top:D.top,width:r,height:l},horizontal:o<0?"left":n>0?"right":"center",vertical:c<0?"top":h>0?"bottom":"middle"};f<r&&s(n+o)<f&&(u.horizontal="center"),g<l&&s(h+c)<g&&(u.vertical="middle"),i(s(n),s(o))>i(s(h),s(c))?u.important="horizontal":u.important="vertical",e.using.call(this,t,u)}),a.offset(t.extend(D,{using:o}))})},t.ui.position={fit:{left:function(t,e){var s,n=e.within,o=n.isWindow?n.scrollLeft:n.offset.left,a=n.width,r=t.left-e.collisionPosition.marginLeft,l=o-r,h=r+e.collisionWidth-a-o;e.collisionWidth>a?l>0&&h<=0?(s=t.left+l+e.collisionWidth-a-o,t.left+=l-s):h>0&&l<=0?t.left=o:l>h?t.left=o+a-e.collisionWidth:t.left=o:l>0?t.left+=l:h>0?t.left-=h:t.left=i(t.left-r,t.left)},top:function(t,e){var s,n=e.within,o=n.isWindow?n.scrollTop:n.offset.top,a=e.within.height,r=t.top-e.collisionPosition.marginTop,l=o-r,h=r+e.collisionHeight-a-o;e.collisionHeight>a?l>0&&h<=0?(s=t.top+l+e.collisionHeight-a-o,t.top+=l-s):h>0&&l<=0?t.top=o:l>h?t.top=o+a-e.collisionHeight:t.top=o:l>0?t.top+=l:h>0?t.top-=h:t.top=i(t.top-r,t.top)}},flip:{left:function(t,e){var i,n,o=e.within,a=o.offset.left+o.scrollLeft,r=o.width,l=o.isWindow?o.scrollLeft:o.offset.left,h=t.left-e.collisionPosition.marginLeft,c=h-l,u=h+e.collisionWidth-r-l,d="left"===e.my[0]?-e.elemWidth:"right"===e.my[0]?e.elemWidth:0,p="left"===e.at[0]?e.targetWidth:"right"===e.at[0]?-e.targetWidth:0,f=-2*e.offset[0];c<0?((i=t.left+d+p+f+e.collisionWidth-r-a)<0||i<s(c))&&(t.left+=d+p+f):u>0&&((n=t.left-e.collisionPosition.marginLeft+d+p+f-l)>0||s(n)<u)&&(t.left+=d+p+f)},top:function(t,e){var i,n,o=e.within,a=o.offset.top+o.scrollTop,r=o.height,l=o.isWindow?o.scrollTop:o.offset.top,h=t.top-e.collisionPosition.marginTop,c=h-l,u=h+e.collisionHeight-r-l,d="top"===e.my[1]?-e.elemHeight:"bottom"===e.my[1]?e.elemHeight:0,p="top"===e.at[1]?e.targetHeight:"bottom"===e.at[1]?-e.targetHeight:0,f=-2*e.offset[1];c<0?((n=t.top+d+p+f+e.collisionHeight-r-a)<0||n<s(c))&&(t.top+=d+p+f):u>0&&((i=t.top-e.collisionPosition.marginTop+d+p+f-l)>0||s(i)<u)&&(t.top+=d+p+f)}},flipfit:{left:function(){t.ui.position.flip.left.apply(this,arguments),t.ui.position.fit.left.apply(this,arguments)},top:function(){t.ui.position.flip.top.apply(this,arguments),t.ui.position.fit.top.apply(this,arguments)}}}}(),t.ui.position,t.extend(t.expr.pseudos,{data:t.expr.createPseudo(function(e){return function(i){return!!t.data(i,e)}})}),t.fn.extend({disableSelection:(c="onselectstart"in document.createElement("div")?"selectstart":"mousedown",function(){return this.on(c+".ui-disableSelection",function(t){t.preventDefault()})}),enableSelection:function(){return this.off(".ui-disableSelection")}});var h,c,u,d={},p=d.toString,f=/^([\-+])=\s*(\d+\.?\d*)/,g=[{re:/rgba?\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*(?:,\s*(\d?(?:\.\d+)?)\s*)?\)/,parse:function(t){return[t[1],t[2],t[3],t[4]]}},{re:/rgba?\(\s*(\d+(?:\.\d+)?)\%\s*,\s*(\d+(?:\.\d+)?)\%\s*,\s*(\d+(?:\.\d+)?)\%\s*(?:,\s*(\d?(?:\.\d+)?)\s*)?\)/,parse:function(t){return[2.55*t[1],2.55*t[2],2.55*t[3],t[4]]}},{re:/#([a-f0-9]{2})([a-f0-9]{2})([a-f0-9]{2})([a-f0-9]{2})?/,parse:function(t){return[parseInt(t[1],16),parseInt(t[2],16),parseInt(t[3],16),t[4]?(parseInt(t[4],16)/255).toFixed(2):1]}},{re:/#([a-f0-9])([a-f0-9])([a-f0-9])([a-f0-9])?/,parse:function(t){return[parseInt(t[1]+t[1],16),parseInt(t[2]+t[2],16),parseInt(t[3]+t[3],16),t[4]?(parseInt(t[4]+t[4],16)/255).toFixed(2):1]}},{re:/hsla?\(\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\%\s*,\s*(\d+(?:\.\d+)?)\%\s*(?:,\s*(\d?(?:\.\d+)?)\s*)?\)/,space:"hsla",parse:function(t){return[t[1],t[2]/100,t[3]/100,t[4]]}}],m=t.Color=function(e,i,s,n){return new t.Color.fn.parse(e,i,s,n)},_={rgba:{props:{red:{idx:0,type:"byte"},green:{idx:1,type:"byte"},blue:{idx:2,type:"byte"}}},hsla:{props:{hue:{idx:0,type:"degrees"},saturation:{idx:1,type:"percent"},lightness:{idx:2,type:"percent"}}}},v={byte:{floor:!0,max:255},percent:{max:1},degrees:{mod:360,floor:!0}},b=t.each;function y(t){return null==t?t+"":"object"==typeof t?d[p.call(t)]||"object":typeof t}function w(t,e,i){var s=v[e.type]||{};return null==t?i||!e.def?null:e.def:(t=s.floor?~~t:parseFloat(t),s.mod)?(t+s.mod)%s.mod:Math.min(s.max,Math.max(0,t))}function x(e){var i=m(),s=i._rgba=[];return(e=e.toLowerCase(),b(g,function(t,n){var o,a=n.re.exec(e),r=a&&n.parse(a),l=n.space||"rgba";if(r)return o=i[l](r),i[_[l].cache]=o[_[l].cache],s=i._rgba=o._rgba,!1}),s.length)?("0,0,0,0"===s.join()&&t.extend(s,u.transparent),i):u[e]}function k(t,e,i){return 6*(i=(i+1)%1)<1?t+(e-t)*i*6:2*i<1?e:3*i<2?t+(e-t)*(2/3-i)*6:t}b(_,function(t,e){e.cache="_"+t,e.props.alpha={idx:3,type:"percent",def:1}}),t.each("Boolean Number String Function Array Date RegExp Object Error Symbol".split(" "),function(t,e){d["[object "+e+"]"]=e.toLowerCase()}),m.fn=t.extend(m.prototype,{parse:function(e,i,s,n){if(void 0===e)return this._rgba=[null,null,null,null],this;(e.jquery||e.nodeType)&&(e=t(e).css(i),i=void 0);var o=this,a=y(e),r=this._rgba=[];return(void 0!==i&&(e=[e,i,s,n],a="array"),"string"===a)?this.parse(x(e)||u._default):"array"===a?(b(_.rgba.props,function(t,i){r[i.idx]=w(e[i.idx],i)}),this):"object"===a?(e instanceof m?b(_,function(t,i){e[i.cache]&&(o[i.cache]=e[i.cache].slice())}):b(_,function(i,s){var n=s.cache;b(s.props,function(t,i){if(!o[n]&&s.to){if("alpha"===t||null==e[t])return;o[n]=s.to(o._rgba)}o[n][i.idx]=w(e[t],i,!0)}),o[n]&&0>t.inArray(null,o[n].slice(0,3))&&(null==o[n][3]&&(o[n][3]=1),s.from&&(o._rgba=s.from(o[n])))}),this):void 0},is:function(t){var e=m(t),i=!0,s=this;return b(_,function(t,n){var o,a=e[n.cache];return a&&(o=s[n.cache]||n.to&&n.to(s._rgba)||[],b(n.props,function(t,e){if(null!=a[e.idx])return i=a[e.idx]===o[e.idx]})),i}),i},_space:function(){var t=[],e=this;return b(_,function(i,s){e[s.cache]&&t.push(i)}),t.pop()},transition:function(t,e){var i=m(t),s=i._space(),n=_[s],o=0===this.alpha()?m("transparent"):this,a=o[n.cache]||n.to(o._rgba),r=a.slice();return i=i[n.cache],b(n.props,function(t,s){var n=s.idx,o=a[n],l=i[n],h=v[s.type]||{};null!==l&&(null===o?r[n]=l:(h.mod&&(l-o>h.mod/2?o+=h.mod:o-l>h.mod/2&&(o-=h.mod)),r[n]=w((l-o)*e+o,s)))}),this[s](r)},blend:function(e){if(1===this._rgba[3])return this;var i=this._rgba.slice(),s=i.pop(),n=m(e)._rgba;return m(t.map(i,function(t,e){return(1-s)*n[e]+s*t}))},toRgbaString:function(){var e="rgba(",i=t.map(this._rgba,function(t,e){return null!=t?t:+(e>2)});return 1===i[3]&&(i.pop(),e="rgb("),e+i.join(", ")+")"},toHslaString:function(){var e="hsla(",i=t.map(this.hsla(),function(t,e){return null==t&&(t=+(e>2)),e&&e<3&&(t=Math.round(100*t)+"%"),t});return 1===i[3]&&(i.pop(),e="hsl("),e+i.join(", ")+")"},toHexString:function(e){var i=this._rgba.slice(),s=i.pop();return e&&i.push(~~(255*s)),"#"+t.map(i,function(t){return("0"+(t||0).toString(16)).substr(-2)}).join("")},toString:function(){return this.toRgbaString()}}),m.fn.parse.prototype=m.fn,_.hsla.to=function(t){if(null==t[0]||null==t[1]||null==t[2])return[null,null,null,t[3]];var e,i=t[0]/255,s=t[1]/255,n=t[2]/255,o=t[3],a=Math.max(i,s,n),r=Math.min(i,s,n),l=a-r,h=a+r,c=.5*h;return[Math.round(r===a?0:i===a?60*(s-n)/l+360:s===a?60*(n-i)/l+120:60*(i-s)/l+240)%360,0===l?0:c<=.5?l/h:l/(2-h),c,null==o?1:o]},_.hsla.from=function(t){if(null==t[0]||null==t[1]||null==t[2])return[null,null,null,t[3]];var e=t[0]/360,i=t[1],s=t[2],n=t[3],o=s<=.5?s*(1+i):s+i-s*i,a=2*s-o;return[Math.round(255*k(a,o,e+1/3)),Math.round(255*k(a,o,e)),Math.round(255*k(a,o,e-1/3)),n]},b(_,function(t,e){var i=e.props,s=e.cache,n=e.to,o=e.from;m.fn[t]=function(t){if(n&&!this[s]&&(this[s]=n(this._rgba)),void 0===t)return this[s].slice();var e,a=y(t),r="array"===a||"object"===a?t:arguments,l=this[s].slice();return(b(i,function(t,e){var i=r["object"===a?t:e.idx];null==i&&(i=l[e.idx]),l[e.idx]=w(i,e)}),o)?((e=m(o(l)))[s]=l,e):m(l)},b(i,function(e,i){!m.fn[e]&&(m.fn[e]=function(s){var n,o,a,r,l=y(s);return(o=(n=this[r="alpha"===e?this._hsla?"hsla":"rgba":t]())[i.idx],"undefined"===l)?o:("function"===l&&(l=y(s=s.call(this,o))),null==s&&i.empty)?this:("string"===l&&(a=f.exec(s))&&(s=o+parseFloat(a[2])*("+"===a[1]?1:-1)),n[i.idx]=s,this[r](n))})})}),m.hook=function(e){b(e.split(" "),function(e,i){t.cssHooks[i]={set:function(t,e){var s;"transparent"!==e&&("string"!==y(e)||(s=x(e)))&&(e=(e=m(s||e)).toRgbaString()),t.style[i]=e}},t.fx.step[i]=function(e){e.colorInit||(e.start=m(e.elem,i),e.end=m(e.end),e.colorInit=!0),t.cssHooks[i].set(e.elem,e.start.transition(e.end,e.pos))}})},m.hook("backgroundColor borderBottomColor borderLeftColor borderRightColor borderTopColor color columnRuleColor outlineColor textDecorationColor textEmphasisColor"),t.cssHooks.borderColor={expand:function(t){var e={};return b(["Top","Right","Bottom","Left"],function(i,s){e["border"+s+"Color"]=t}),e}},u=t.Color.names={aqua:"#00ffff",black:"#000000",blue:"#0000ff",fuchsia:"#ff00ff",gray:"#808080",green:"#008000",lime:"#00ff00",maroon:"#800000",navy:"#000080",olive:"#808000",purple:"#800080",red:"#ff0000",silver:"#c0c0c0",teal:"#008080",white:"#ffffff",yellow:"#ffff00",transparent:[null,null,null,0],_default:"#ffffff"};var C="ui-effects-",D="ui-effects-style",I="ui-effects-animated";t.effects={effect:{}},function(){var e,i,s,n=["add","remove","toggle"],o={border:1,borderBottom:1,borderColor:1,borderLeft:1,borderRight:1,borderTop:1,borderWidth:1,margin:1,padding:1};function a(t){var e,i,s=t.ownerDocument.defaultView.getComputedStyle(t),n={};i=s.length;while(i--)e=s[i],"string"==typeof s[e]&&(n[e.replace(/-([\da-z])/gi,function(t,e){return e.toUpperCase()})]=s[e]);return n}t.each(["borderLeftStyle","borderRightStyle","borderBottomStyle","borderTopStyle"],function(e,i){t.fx.step[i]=function(e){("none"===e.end||e.setAttr)&&(1!==e.pos||e.setAttr)||(t.style(e.elem,i,e.end),e.setAttr=!0)}}),t.effects.animateClass=function(e,i,s,r){var l=t.speed(i,s,r);return this.queue(function(){var i,s=t(this),r=s.attr("class")||"",h=l.children?s.find("*").addBack():s;h=h.map(function(){return{el:t(this),start:a(this)}}),(i=function(){t.each(n,function(t,i){e[i]&&s[i+"Class"](e[i])})})(),h=h.map(function(){return this.end=a(this.el[0]),this.diff=function(e,i){var s,n,a={};for(s in i)n=i[s],e[s]===n||o[s]||!t.fx.step[s]&&isNaN(parseFloat(n))||(a[s]=n);return a}(this.start,this.end),this}),s.attr("class",r),h=h.map(function(){var e=this,i=t.Deferred(),s=t.extend({},l,{queue:!1,complete:function(){i.resolve(e)}});return this.el.animate(this.diff,s),i.promise()}),t.when.apply(t,h.get()).done(function(){i(),t.each(arguments,function(){var e=this.el;t.each(this.diff,function(t){e.css(t,"")})}),l.complete.call(s[0])})})},t.fn.extend({addClass:(e=t.fn.addClass,function(i,s,n,o){return s?t.effects.animateClass.call(this,{add:i},s,n,o):e.apply(this,arguments)}),removeClass:(i=t.fn.removeClass,function(e,s,n,o){return arguments.length>1?t.effects.animateClass.call(this,{remove:e},s,n,o):i.apply(this,arguments)}),toggleClass:(s=t.fn.toggleClass,function(e,i,n,o,a){return"boolean"!=typeof i&&void 0!==i?t.effects.animateClass.call(this,{toggle:e},i,n,o):n?t.effects.animateClass.call(this,i?{add:e}:{remove:e},n,o,a):s.apply(this,arguments)}),switchClass:function(e,i,s,n,o){return t.effects.animateClass.call(this,{add:i,remove:e},s,n,o)}})}(),function(){var e,i,s,n;function o(e,i,s,n){return t.isPlainObject(e)&&(i=e,e=e.effect),e={effect:e},null==i&&(i={}),"function"==typeof i&&(n=i,s=null,i={}),("number"==typeof i||t.fx.speeds[i])&&(n=s,s=i,i={}),"function"==typeof s&&(n=s,s=null),i&&t.extend(e,i),s=s||i.duration,e.duration=t.fx.off?0:"number"==typeof s?s:s in t.fx.speeds?t.fx.speeds[s]:t.fx.speeds._default,e.complete=n||i.complete,e}function a(e){return!e||"number"==typeof e||!!t.fx.speeds[e]||"string"==typeof e&&!t.effects.effect[e]||"function"==typeof e||"object"==typeof e&&!e.effect}function r(t,e){var i=e.outerWidth(),s=e.outerHeight(),n=/^rect\((-?\d*\.?\d*px|-?\d+%|auto),?\s*(-?\d*\.?\d*px|-?\d+%|auto),?\s*(-?\d*\.?\d*px|-?\d+%|auto),?\s*(-?\d*\.?\d*px|-?\d+%|auto)\)$/.exec(t)||["",0,i,s,0];return{top:parseFloat(n[1])||0,right:"auto"===n[2]?i:parseFloat(n[2]),bottom:"auto"===n[3]?s:parseFloat(n[3]),left:parseFloat(n[4])||0}}t.expr&&t.expr.pseudos&&t.expr.pseudos.animated&&(t.expr.pseudos.animated=(e=t.expr.pseudos.animated,function(i){return!!t(i).data(I)||e(i)})),!0===t.uiBackCompat&&t.extend(t.effects,{save:function(t,e){for(var i=0,s=e.length;i<s;i++)null!==e[i]&&t.data(C+e[i],t[0].style[e[i]])},restore:function(t,e){for(var i,s=0,n=e.length;s<n;s++)null!==e[s]&&(i=t.data(C+e[s]),t.css(e[s],i))},setMode:function(t,e){return"toggle"===e&&(e=t.is(":hidden")?"show":"hide"),e},createWrapper:function(e){if(e.parent().is(".ui-effects-wrapper"))return e.parent();var i={width:e.outerWidth(!0),height:e.outerHeight(!0),float:e.css("float")},s=t("<div></div>").addClass("ui-effects-wrapper").css({fontSize:"100%",background:"transparent",border:"none",margin:0,padding:0}),n={width:e.width(),height:e.height()},o=document.activeElement;try{o.id}catch(t){o=document.body}return e.wrap(s),(e[0]===o||t.contains(e[0],o))&&t(o).trigger("focus"),s=e.parent(),"static"===e.css("position")?(s.css({position:"relative"}),e.css({position:"relative"})):(t.extend(i,{position:e.css("position"),zIndex:e.css("z-index")}),t.each(["top","left","bottom","right"],function(t,s){i[s]=e.css(s),isNaN(parseInt(i[s],10))&&(i[s]="auto")}),e.css({position:"relative",top:0,left:0,right:"auto",bottom:"auto"})),e.css(n),s.css(i).show()},removeWrapper:function(e){var i=document.activeElement;return e.parent().is(".ui-effects-wrapper")&&(e.parent().replaceWith(e),(e[0]===i||t.contains(e[0],i))&&t(i).trigger("focus")),e}}),t.extend(t.effects,{version:"1.14.2",define:function(e,i,s){return s||(s=i,i="effect"),t.effects.effect[e]=s,t.effects.effect[e].mode=i,s},scaledDimensions:function(t,e,i){if(0===e)return{height:0,width:0,outerHeight:0,outerWidth:0};var s="horizontal"!==i?(e||100)/100:1,n="vertical"!==i?(e||100)/100:1;return{height:t.height()*n,width:t.width()*s,outerHeight:t.outerHeight()*n,outerWidth:t.outerWidth()*s}},clipToBox:function(t){return{width:t.clip.right-t.clip.left,height:t.clip.bottom-t.clip.top,left:t.clip.left,top:t.clip.top}},unshift:function(t,e,i){var s=t.queue();e>1&&s.splice.apply(s,[1,0].concat(s.splice(e,i))),t.dequeue()},saveStyle:function(t){t.data(D,t[0].style.cssText)},restoreStyle:function(t){t[0].style.cssText=t.data(D)||"",t.removeData(D)},mode:function(t,e){var i=t.is(":hidden");return"toggle"===e&&(e=i?"show":"hide"),(i?"hide"===e:"show"===e)&&(e="none"),e},getBaseline:function(t,e){var i,s;switch(t[0]){case"top":i=0;break;case"middle":i=.5;break;case"bottom":i=1;break;default:i=t[0]/e.height}switch(t[1]){case"left":s=0;break;case"center":s=.5;break;case"right":s=1;break;default:s=t[1]/e.width}return{x:s,y:i}},createPlaceholder:function(e){var i,s=e.css("position"),n=e.position();return e.css({marginTop:e.css("marginTop"),marginBottom:e.css("marginBottom"),marginLeft:e.css("marginLeft"),marginRight:e.css("marginRight")}).outerWidth(e.outerWidth()).outerHeight(e.outerHeight()),/^(static|relative)/.test(s)&&(s="absolute",i=t("<"+e[0].nodeName+">").insertAfter(e).css({display:/^(inline|ruby)/.test(e.css("display"))?"inline-block":"block",visibility:"hidden",marginTop:e.css("marginTop"),marginBottom:e.css("marginBottom"),marginLeft:e.css("marginLeft"),marginRight:e.css("marginRight"),float:e.css("float")}).outerWidth(e.outerWidth()).outerHeight(e.outerHeight()).addClass("ui-effects-placeholder"),e.data(C+"placeholder",i)),e.css({position:s,left:n.left,top:n.top}),i},removePlaceholder:function(t){var e=C+"placeholder",i=t.data(e);i&&(i.remove(),t.removeData(e))},cleanUp:function(e){t.effects.restoreStyle(e),t.effects.removePlaceholder(e)},setTransition:function(e,i,s,n){return n=n||{},t.each(i,function(t,i){var o=e.cssUnit(i);o[0]>0&&(n[i]=o[0]*s+o[1])}),n}}),t.fn.extend({effect:function(){var e=o.apply(this,arguments),i=t.effects.effect[e.effect],s=i.mode,n=e.queue,a=n||"fx",r=e.complete,l=e.mode,h=[],c=function(e){var i=t(this),n=t.effects.mode(i,l)||s;i.data(I,!0),h.push(n),s&&("show"===n||n===s&&"hide"===n)&&i.show(),s&&"none"===n||t.effects.saveStyle(i),"function"==typeof e&&e()};if(t.fx.off||!i)return l?this[l](e.duration,r):this.each(function(){r&&r.call(this)});function u(n){var o=t(this);function a(){"function"==typeof r&&r.call(o[0]),"function"==typeof n&&n()}e.mode=h.shift(),!0!==t.uiBackCompat||s?"none"===e.mode?(o[l](),a()):i.call(o[0],e,function(){o.removeData(I),t.effects.cleanUp(o),"hide"===e.mode&&o.hide(),a()}):(o.is(":hidden")?"hide"===l:"show"===l)?(o[l](),a()):i.call(o[0],e,a)}return!1===n?this.each(c).each(u):this.queue(a,c).queue(a,u)},show:(i=t.fn.show,function(t){if(a(t))return i.apply(this,arguments);var e=o.apply(this,arguments);return e.mode="show",this.effect.call(this,e)}),hide:(s=t.fn.hide,function(t){if(a(t))return s.apply(this,arguments);var e=o.apply(this,arguments);return e.mode="hide",this.effect.call(this,e)}),toggle:(n=t.fn.toggle,function(t){if(a(t)||"boolean"==typeof t)return n.apply(this,arguments);var e=o.apply(this,arguments);return e.mode="toggle",this.effect.call(this,e)}),cssUnit:function(e){var i=this.css(e),s=[];return t.each(["em","px","%","pt"],function(t,e){i.indexOf(e)>0&&(s=[parseFloat(i),e])}),s},cssClip:function(t){return t?this.css("clip","rect("+t.top+"px "+t.right+"px "+t.bottom+"px "+t.left+"px)"):r(this.css("clip"),this)},transfer:function(e,i){var s=t(this),n=t(e.to),o="fixed"===n.css("position"),a=t("body"),r=o?a.scrollTop():0,l=o?a.scrollLeft():0,h=n.offset(),c={top:h.top-r,left:h.left-l,height:n.innerHeight(),width:n.innerWidth()},u=s.offset(),d=t("<div class='ui-effects-transfer'></div>");d.appendTo("body").addClass(e.className).css({top:u.top-r,left:u.left-l,height:s.innerHeight(),width:s.innerWidth(),position:o?"fixed":"absolute"}).animate(c,e.duration,e.easing,function(){d.remove(),"function"==typeof i&&i()})}}),t.fx.step.clip=function(e){e.clipInit||(e.start=t(e.elem).cssClip(),"string"==typeof e.end&&(e.end=r(e.end,e.elem)),e.clipInit=!0),t(e.elem).cssClip({top:e.pos*(e.end.top-e.start.top)+e.start.top,right:e.pos*(e.end.right-e.start.right)+e.start.right,bottom:e.pos*(e.end.bottom-e.start.bottom)+e.start.bottom,left:e.pos*(e.end.left-e.start.left)+e.start.left})}}(),e={},t.each(["Quad","Cubic","Quart","Quint","Expo"],function(t,i){e[i]=function(e){return Math.pow(e,t+2)}}),t.extend(e,{Sine:function(t){return 1-Math.cos(t*Math.PI/2)},Circ:function(t){return 1-Math.sqrt(1-t*t)},Elastic:function(t){return 0===t||1===t?t:-Math.pow(2,8*(t-1))*Math.sin(((t-1)*80-7.5)*Math.PI/15)},Back:function(t){return t*t*(3*t-2)},Bounce:function(t){var e,i=4;while(t<((e=Math.pow(2,--i))-1)/11);return 1/Math.pow(4,3-i)-7.5625*Math.pow((3*e-2)/22-t,2)}}),t.each(e,function(e,i){t.easing["easeIn"+e]=i,t.easing["easeOut"+e]=function(t){return 1-i(1-t)},t.easing["easeInOut"+e]=function(t){return t<.5?i(2*t)/2:1-i(-2*t+2)/2}});var n=t.effects;t.effects.define("blind","hide",function(e,i){var s={up:["bottom","top"],vertical:["bottom","top"],down:["top","bottom"],left:["right","left"],horizontal:["right","left"],right:["left","right"]},n=t(this),o=e.direction||"up",a=n.cssClip(),r={clip:t.extend({},a)},l=t.effects.createPlaceholder(n);r.clip[s[o][0]]=r.clip[s[o][1]],"show"===e.mode&&(n.cssClip(r.clip),l&&l.css(t.effects.clipToBox(r)),r.clip=a),l&&l.animate(t.effects.clipToBox(r),e.duration,e.easing),n.animate(r,{queue:!1,duration:e.duration,easing:e.easing,complete:i})}),t.effects.define("bounce",function(e,i){var s,n,o,a=t(this),r=e.mode,l="hide"===r,h="show"===r,c=e.direction||"up",u=e.distance,d=e.times||5,p=2*d+(h||l?1:0),f=e.duration/p,g=e.easing,m="up"===c||"down"===c?"top":"left",_="up"===c||"left"===c,v=0,b=a.queue().length;for(t.effects.createPlaceholder(a),o=a.css(m),u||(u=a["top"===m?"outerHeight":"outerWidth"]()/3),h&&((n={opacity:1})[m]=o,a.css("opacity",0).css(m,_?-(2*u):2*u).animate(n,f,g)),l&&(u/=Math.pow(2,d-1)),(n={})[m]=o;v<d;v++)(s={})[m]=(_?"-=":"+=")+u,a.animate(s,f,g).animate(n,f,g),u=l?2*u:u/2;l&&((s={opacity:0})[m]=(_?"-=":"+=")+u,a.animate(s,f,g)),a.queue(i),t.effects.unshift(a,b,p+1)}),t.effects.define("clip","hide",function(e,i){var s,n={},o=t(this),a=e.direction||"vertical",r="both"===a,l=r||"horizontal"===a,h=r||"vertical"===a;s=o.cssClip(),n.clip={top:h?(s.bottom-s.top)/2:s.top,right:l?(s.right-s.left)/2:s.right,bottom:h?(s.bottom-s.top)/2:s.bottom,left:l?(s.right-s.left)/2:s.left},t.effects.createPlaceholder(o),"show"===e.mode&&(o.cssClip(n.clip),n.clip=s),o.animate(n,{queue:!1,duration:e.duration,easing:e.easing,complete:i})}),t.effects.define("drop","hide",function(e,i){var s,n=t(this),o=e.mode,a=e.direction||"left",r="up"===a||"down"===a?"top":"left",l="up"===a||"left"===a?"-=":"+=",h={opacity:0};t.effects.createPlaceholder(n),s=e.distance||n["top"===r?"outerHeight":"outerWidth"](!0)/2,h[r]=l+s,"show"===o&&(n.css(h),h[r]=("+="===l?"-=":"+=")+s,h.opacity=1),n.animate(h,{queue:!1,duration:e.duration,easing:e.easing,complete:i})}),t.effects.define("explode","hide",function(e,i){var s,n,o,a,r,l,h=e.pieces?Math.round(Math.sqrt(e.pieces)):3,c=t(this),u="show"===e.mode,d=c.show().css("visibility","hidden").offset(),p=Math.ceil(c.outerWidth()/h),f=Math.ceil(c.outerHeight()/h),g=[];function m(){g.push(this),g.length===h*h&&(c.css({visibility:"visible"}),t(g).remove(),i())}for(s=0;s<h;s++)for(n=0,a=d.top+s*f,l=s-(h-1)/2;n<h;n++)o=d.left+n*p,r=n-(h-1)/2,c.clone().appendTo("body").wrap("<div></div>").css({position:"absolute",visibility:"visible",left:-n*p,top:-s*f}).parent().addClass("ui-effects-explode").css({position:"absolute",overflow:"hidden",width:p,height:f,left:o+(u?r*p:0),top:a+(u?l*f:0),opacity:+!u}).animate({left:o+(u?0:r*p),top:a+(u?0:l*f),opacity:+!!u},e.duration||500,e.easing,m)}),t.effects.define("fade","toggle",function(e,i){var s="show"===e.mode;t(this).css("opacity",+!s).animate({opacity:+!!s},{queue:!1,duration:e.duration,easing:e.easing,complete:i})}),t.effects.define("fold","hide",function(e,i){var s=t(this),n=e.mode,o=e.size||15,a=/([0-9]+)%/.exec(o),r=e.horizFirst?["right","bottom"]:["bottom","right"],l=e.duration/2,h=t.effects.createPlaceholder(s),c=s.cssClip(),u={clip:t.extend({},c)},d={clip:t.extend({},c)},p=[c[r[0]],c[r[1]]],f=s.queue().length;a&&(o=parseInt(a[1],10)/100*p[+("hide"!==n)]),u.clip[r[0]]=o,d.clip[r[0]]=o,d.clip[r[1]]=0,"show"===n&&(s.cssClip(d.clip),h&&h.css(t.effects.clipToBox(d)),d.clip=c),s.queue(function(i){h&&h.animate(t.effects.clipToBox(u),l,e.easing).animate(t.effects.clipToBox(d),l,e.easing),i()}).animate(u,l,e.easing).animate(d,l,e.easing).queue(i),t.effects.unshift(s,f,4)}),t.effects.define("highlight","show",function(e,i){var s=t(this),n={backgroundColor:s.css("backgroundColor")};"hide"===e.mode&&(n.opacity=0),t.effects.saveStyle(s),s.css({backgroundImage:"none",backgroundColor:e.color||"#ffff99"}).animate(n,{queue:!1,duration:e.duration,easing:e.easing,complete:i})}),t.effects.define("size",function(e,i){var s,n,o,a=t(this),r=["fontSize"],l=["borderTopWidth","borderBottomWidth","paddingTop","paddingBottom"],h=["borderLeftWidth","borderRightWidth","paddingLeft","paddingRight"],c=e.mode,u="effect"!==c,d=e.scale||"both",p=e.origin||["middle","center"],f=a.css("position"),g=a.position(),m=t.effects.scaledDimensions(a),_=e.from||m,v=e.to||t.effects.scaledDimensions(a,0);t.effects.createPlaceholder(a),"show"===c&&(o=_,_=v,v=o),n={from:{y:_.height/m.height,x:_.width/m.width},to:{y:v.height/m.height,x:v.width/m.width}},("box"===d||"both"===d)&&(n.from.y!==n.to.y&&(_=t.effects.setTransition(a,l,n.from.y,_),v=t.effects.setTransition(a,l,n.to.y,v)),n.from.x!==n.to.x&&(_=t.effects.setTransition(a,h,n.from.x,_),v=t.effects.setTransition(a,h,n.to.x,v))),("content"===d||"both"===d)&&n.from.y!==n.to.y&&(_=t.effects.setTransition(a,r,n.from.y,_),v=t.effects.setTransition(a,r,n.to.y,v)),p&&(s=t.effects.getBaseline(p,m),_.top=(m.outerHeight-_.outerHeight)*s.y+g.top,_.left=(m.outerWidth-_.outerWidth)*s.x+g.left,v.top=(m.outerHeight-v.outerHeight)*s.y+g.top,v.left=(m.outerWidth-v.outerWidth)*s.x+g.left),delete _.outerHeight,delete _.outerWidth,a.css(_),("content"===d||"both"===d)&&(l=l.concat(["marginTop","marginBottom"]).concat(r),h=h.concat(["marginLeft","marginRight"]),a.find("*[width]").each(function(){var i=t(this),s=t.effects.scaledDimensions(i),o={height:s.height*n.from.y,width:s.width*n.from.x,outerHeight:s.outerHeight*n.from.y,outerWidth:s.outerWidth*n.from.x},a={height:s.height*n.to.y,width:s.width*n.to.x,outerHeight:s.height*n.to.y,outerWidth:s.width*n.to.x};n.from.y!==n.to.y&&(o=t.effects.setTransition(i,l,n.from.y,o),a=t.effects.setTransition(i,l,n.to.y,a)),n.from.x!==n.to.x&&(o=t.effects.setTransition(i,h,n.from.x,o),a=t.effects.setTransition(i,h,n.to.x,a)),u&&t.effects.saveStyle(i),i.css(o),i.animate(a,e.duration,e.easing,function(){u&&t.effects.restoreStyle(i)})})),a.animate(v,{queue:!1,duration:e.duration,easing:e.easing,complete:function(){var e=a.offset();0===v.opacity&&a.css("opacity",_.opacity),u||(a.css("position","static"===f?"relative":f).offset(e),t.effects.saveStyle(a)),i()}})}),t.effects.define("scale",function(e,i){var s=t(this),n=e.mode,o=parseInt(e.percent,10)||(0===parseInt(e.percent,10)?0:100*("effect"===n)),a=t.extend(!0,{from:t.effects.scaledDimensions(s),to:t.effects.scaledDimensions(s,o,e.direction||"both"),origin:e.origin||["middle","center"]},e);e.fade&&(a.from.opacity=1,a.to.opacity=0),t.effects.effect.size.call(this,a,i)}),t.effects.define("puff","hide",function(e,i){var s=t.extend(!0,{},e,{fade:!0,percent:parseInt(e.percent,10)||150});t.effects.effect.scale.call(this,s,i)}),t.effects.define("pulsate","show",function(e,i){var s=t(this),n=e.mode,o="show"===n,a=2*(e.times||5)+ +!!(o||"hide"===n),r=e.duration/a,l=0,h=1,c=s.queue().length;for((o||!s.is(":visible"))&&(s.css("opacity",0).show(),l=1);h<a;h++)s.animate({opacity:l},r,e.easing),l=1-l;s.animate({opacity:l},r,e.easing),s.queue(i),t.effects.unshift(s,c,a+1)}),t.effects.define("shake",function(e,i){var s=1,n=t(this),o=e.direction||"left",a=e.distance||20,r=e.times||3,l=2*r+1,h=Math.round(e.duration/l),c="up"===o||"down"===o?"top":"left",u="up"===o||"left"===o,d={},p={},f={},g=n.queue().length;for(t.effects.createPlaceholder(n),d[c]=(u?"-=":"+=")+a,p[c]=(u?"+=":"-=")+2*a,f[c]=(u?"-=":"+=")+2*a,n.animate(d,h,e.easing);s<r;s++)n.animate(p,h,e.easing).animate(f,h,e.easing);n.animate(p,h,e.easing).animate(d,h/2,e.easing).queue(i),t.effects.unshift(n,g,l+1)}),t.effects.define("slide","show",function(e,i){var s,n,o=t(this),a={up:["bottom","top"],down:["top","bottom"],left:["right","left"],right:["left","right"]},r=e.mode,l=e.direction||"left",h="up"===l||"down"===l?"top":"left",c=e.distance||o["top"===h?"outerHeight":"outerWidth"](!0),u={};t.effects.createPlaceholder(o),s=o.cssClip(),n=o.position()[h],u[h]=("up"===l||"left"===l?-1:1)*c+n,u.clip=o.cssClip(),u.clip[a[l][1]]=u.clip[a[l][0]],"show"===r&&(o.cssClip(u.clip),o.css(h,u[h]),u.clip=s,u[h]=n),o.animate(u,{queue:!1,duration:e.duration,easing:e.easing,complete:i})}),!0===t.uiBackCompat&&t.effects.define("transfer",function(e,i){t(this).transfer(e,i)}),t.ui.focusable=function(e,i){var s,n,o,a,r,l=e.nodeName.toLowerCase();return"area"===l?(n=(s=e.parentNode).name,!!e.href&&!!n&&"map"===s.nodeName.toLowerCase()&&(o=t("img[usemap='#"+n+"']")).length>0&&o.is(":visible")):(/^(input|select|textarea|button|object)$/.test(l)?(a=!e.disabled)&&(r=t(e).closest("fieldset")[0])&&(a=!r.disabled):a="a"===l&&e.href||i,a&&t(e).is(":visible")&&"visible"===t(e).css("visibility"))},t.extend(t.expr.pseudos,{focusable:function(e){return t.ui.focusable(e,null!=t.attr(e,"tabindex"))}}),t.ui.focusable,t.ui.formResetMixin={_formResetHandler:function(){var e=t(this);setTimeout(function(){var i=e.data("ui-form-reset-instances");t.each(i,function(){this.refresh()})})},_bindFormResetHandler:function(){if(this.form=t(this.element.prop("form")),this.form.length){var e=this.form.data("ui-form-reset-instances")||[];e.length||this.form.on("reset.ui-form-reset",this._formResetHandler),e.push(this),this.form.data("ui-form-reset-instances",e)}},_unbindFormResetHandler:function(){if(this.form.length){var e=this.form.data("ui-form-reset-instances");e.splice(t.inArray(this,e),1),e.length?this.form.data("ui-form-reset-instances",e):this.form.removeData("ui-form-reset-instances").off("reset.ui-form-reset")}}},t.escapeSelector||(t.escapeSelector=function(t){return CSS.escape(t+"")}),t.fn.even&&t.fn.odd||t.fn.extend({even:function(){return this.filter(function(t){return t%2==0})},odd:function(){return this.filter(function(t){return t%2==1})}}),t.ui.keyCode={BACKSPACE:8,COMMA:188,DELETE:46,DOWN:40,END:35,ENTER:13,ESCAPE:27,HOME:36,LEFT:37,PAGE_DOWN:34,PAGE_UP:33,PERIOD:190,RIGHT:39,SPACE:32,TAB:9,UP:38},t.fn.labels=function(){var t,e,i,s,n;return this.length?this[0].labels&&this[0].labels.length?this.pushStack(this[0].labels):(s=this.eq(0).parents("label"),(i=this.attr("id"))&&(n=(t=this.eq(0).parents().last()).add(t.length?t.siblings():this.siblings()),e="label[for='"+CSS.escape(i)+"']",s=s.add(n.find(e).addBack(e))),this.pushStack(s)):this.pushStack([])},t.fn.scrollParent=function(e){var i=this.css("position"),s="absolute"===i,n=e?/(auto|scroll|hidden)/:/(auto|scroll)/,o=this.parents().filter(function(){var e=t(this);return(!s||"static"!==e.css("position"))&&n.test(e.css("overflow")+e.css("overflow-y")+e.css("overflow-x"))}).eq(0);return"fixed"!==i&&o.length?o:t(this[0].ownerDocument||document)},t.extend(t.expr.pseudos,{tabbable:function(e){var i=t.attr(e,"tabindex"),s=null!=i;return(!s||i>=0)&&t.ui.focusable(e,s)}}),t.fn.extend({uniqueId:(i=0,function(){return this.each(function(){this.id||(this.id="ui-id-"+ ++i)})}),removeUniqueId:function(){return this.each(function(){/^ui-id-\d+$/.test(this.id)&&t(this).removeAttr("id")})}}),t.widget("ui.accordion",{version:"1.14.2",options:{active:0,animate:{},classes:{"ui-accordion-header":"ui-corner-top","ui-accordion-header-collapsed":"ui-corner-all","ui-accordion-content":"ui-corner-bottom"},collapsible:!1,event:"click",header:function(t){return t.find("> li > :first-child").add(t.find("> :not(li)").filter(function(t){return t%2==0}))},heightStyle:"auto",icons:{activeHeader:"ui-icon-triangle-1-s",header:"ui-icon-triangle-1-e"},activate:null,beforeActivate:null},hideProps:{borderTopWidth:"hide",borderBottomWidth:"hide",paddingTop:"hide",paddingBottom:"hide",height:"hide"},showProps:{borderTopWidth:"show",borderBottomWidth:"show",paddingTop:"show",paddingBottom:"show",height:"show"},_create:function(){var e=this.options;this.prevShow=this.prevHide=t(),this._addClass("ui-accordion","ui-widget ui-helper-reset"),this.element.attr("role","tablist"),e.collapsible||!1!==e.active&&null!=e.active||(e.active=0),this._processPanels(),e.active<0&&(e.active+=this.headers.length),this._refresh()},_getCreateEventData:function(){return{header:this.active,panel:this.active.length?this.active.next():t()}},_createIcons:function(){var e,i,s=this.options.icons;s&&(e=t("<span>"),this._addClass(e,"ui-accordion-header-icon","ui-icon "+s.header),e.prependTo(this.headers),i=this.active.children(".ui-accordion-header-icon"),this._removeClass(i,s.header)._addClass(i,null,s.activeHeader)._addClass(this.headers,"ui-accordion-icons"))},_destroyIcons:function(){this._removeClass(this.headers,"ui-accordion-icons"),this.headers.children(".ui-accordion-header-icon").remove()},_destroy:function(){var t;this.element.removeAttr("role"),this.headers.removeAttr("role aria-expanded aria-selected aria-controls tabIndex").removeUniqueId(),this._destroyIcons(),t=this.headers.next().css("display","").removeAttr("role aria-hidden aria-labelledby").removeUniqueId(),"content"!==this.options.heightStyle&&t.css("height","")},_setOption:function(t,e){if("active"===t){this._activate(e);return}"event"===t&&(this.options.event&&this._off(this.headers,this.options.event),this._setupEvents(e)),this._super(t,e),"collapsible"!==t||e||!1!==this.options.active||this._activate(0),"icons"===t&&(this._destroyIcons(),e&&this._createIcons())},_setOptionDisabled:function(t){this._super(t),this.element.attr("aria-disabled",t),this._toggleClass(null,"ui-state-disabled",!!t)},_keydown:function(e){if(!e.altKey&&!e.ctrlKey){var i=t.ui.keyCode,s=this.headers.length,n=this.headers.index(e.target),o=!1;switch(e.keyCode){case i.RIGHT:case i.DOWN:o=this.headers[(n+1)%s];break;case i.LEFT:case i.UP:o=this.headers[(n-1+s)%s];break;case i.SPACE:case i.ENTER:this._eventHandler(e);break;case i.HOME:o=this.headers[0];break;case i.END:o=this.headers[s-1]}o&&(t(e.target).attr("tabIndex",-1),t(o).attr("tabIndex",0),t(o).trigger("focus"),e.preventDefault())}},_panelKeyDown:function(e){e.keyCode===t.ui.keyCode.UP&&e.ctrlKey&&t(e.currentTarget).prev().trigger("focus")},refresh:function(){var e=this.options;this._processPanels(),(!1!==e.active||!0!==e.collapsible)&&this.headers.length?!1===e.active?this._activate(0):this.active.length&&!t.contains(this.element[0],this.active[0])?this.headers.length===this.headers.find(".ui-state-disabled").length?(e.active=!1,this.active=t()):this._activate(Math.max(0,e.active-1)):e.active=this.headers.index(this.active):(e.active=!1,this.active=t()),this._destroyIcons(),this._refresh()},_processPanels:function(){var t=this.headers,e=this.panels;"function"==typeof this.options.header?this.headers=this.options.header(this.element):this.headers=this.element.find(this.options.header),this._addClass(this.headers,"ui-accordion-header ui-accordion-header-collapsed","ui-state-default"),this.panels=this.headers.next().filter(":not(.ui-accordion-content-active)").hide(),this._addClass(this.panels,"ui-accordion-content","ui-helper-reset ui-widget-content"),e&&(this._off(t.not(this.headers)),this._off(e.not(this.panels)))},_refresh:function(){var e,i=this.options,s=i.heightStyle,n=this.element.parent();this.active=this._findActive(i.active),this._addClass(this.active,"ui-accordion-header-active","ui-state-active")._removeClass(this.active,"ui-accordion-header-collapsed"),this._addClass(this.active.next(),"ui-accordion-content-active"),this.active.next().show(),this.headers.attr("role","tab").each(function(){var e=t(this),i=e.uniqueId().attr("id"),s=e.next(),n=s.uniqueId().attr("id");e.attr("aria-controls",n),s.attr("aria-labelledby",i)}).next().attr("role","tabpanel"),this.headers.not(this.active).attr({"aria-selected":"false","aria-expanded":"false",tabIndex:-1}).next().attr({"aria-hidden":"true"}).hide(),this.active.length?this.active.attr({"aria-selected":"true","aria-expanded":"true",tabIndex:0}).next().attr({"aria-hidden":"false"}):this.headers.eq(0).attr("tabIndex",0),this._createIcons(),this._setupEvents(i.event),"fill"===s?(e=n.height(),this.element.siblings(":visible").each(function(){var i=t(this),s=i.css("position");"absolute"!==s&&"fixed"!==s&&(e-=i.outerHeight(!0))}),this.headers.each(function(){e-=t(this).outerHeight(!0)}),this.headers.next().each(function(){t(this).height(Math.max(0,e-t(this).innerHeight()+t(this).height()))}).css("overflow","auto")):"auto"===s&&(e=0,this.headers.next().each(function(){var i=t(this).is(":visible");i||t(this).show(),e=Math.max(e,t(this).css("height","").height()),i||t(this).hide()}).height(e))},_activate:function(e){var i=this._findActive(e)[0];i!==this.active[0]&&(i=i||this.active[0],this._eventHandler({target:i,currentTarget:i,preventDefault:t.noop}))},_findActive:function(e){return"number"==typeof e?this.headers.eq(e):t()},_setupEvents:function(e){var i={keydown:"_keydown"};e&&t.each(e.split(" "),function(t,e){i[e]="_eventHandler"}),this._off(this.headers.add(this.headers.next())),this._on(this.headers,i),this._on(this.headers.next(),{keydown:"_panelKeyDown"}),this._hoverable(this.headers),this._focusable(this.headers)},_eventHandler:function(e){var i,s,n=this.options,o=this.active,a=t(e.currentTarget),r=a[0]===o[0],l=r&&n.collapsible,h=l?t():a.next(),c=o.next(),u={oldHeader:o,oldPanel:c,newHeader:l?t():a,newPanel:h};e.preventDefault(),(!r||n.collapsible)&&!1!==this._trigger("beforeActivate",e,u)&&(n.active=!l&&this.headers.index(a),this.active=r?t():a,this._toggle(u),this._removeClass(o,"ui-accordion-header-active","ui-state-active"),n.icons&&(i=o.children(".ui-accordion-header-icon"),this._removeClass(i,null,n.icons.activeHeader)._addClass(i,null,n.icons.header)),r||(this._removeClass(a,"ui-accordion-header-collapsed")._addClass(a,"ui-accordion-header-active","ui-state-active"),n.icons&&(s=a.children(".ui-accordion-header-icon"),this._removeClass(s,null,n.icons.header)._addClass(s,null,n.icons.activeHeader)),this._addClass(a.next(),"ui-accordion-content-active")))},_toggle:function(e){var i=e.newPanel,s=this.prevShow.length?this.prevShow:e.oldPanel;this.prevShow.add(this.prevHide).stop(!0,!0),this.prevShow=i,this.prevHide=s,this.options.animate?this._animate(i,s,e):(s.hide(),i.show(),this._toggleComplete(e)),s.attr({"aria-hidden":"true"}),s.prev().attr({"aria-selected":"false","aria-expanded":"false"}),i.length&&s.length?s.prev().attr({tabIndex:-1,"aria-expanded":"false"}):i.length&&this.headers.filter(function(){return 0===parseInt(t(this).attr("tabIndex"),10)}).attr("tabIndex",-1),i.attr("aria-hidden","false").prev().attr({"aria-selected":"true","aria-expanded":"true",tabIndex:0})},_animate:function(t,e,i){var s,n,o,a=this,r=0,l=t.css("box-sizing"),h=t.length&&(!e.length||t.index()<e.index()),c=this.options.animate||{},u=h&&c.down||c,d=function(){a._toggleComplete(i)};return("number"==typeof u&&(o=u),"string"==typeof u&&(n=u),n=n||u.easing||c.easing,o=o||u.duration||c.duration,e.length)?t.length?void(s=t.show().outerHeight(),e.animate(this.hideProps,{duration:o,easing:n,step:function(t,e){e.now=Math.round(t)}}),t.hide().animate(this.showProps,{duration:o,easing:n,complete:d,step:function(t,i){i.now=Math.round(t),"height"!==i.prop?"content-box"===l&&(r+=i.now):"content"!==a.options.heightStyle&&(i.now=Math.round(s-e.outerHeight()-r),r=0)}})):e.animate(this.hideProps,o,n,d):t.animate(this.showProps,o,n,d)},_toggleComplete:function(t){var e=t.oldPanel,i=e.prev();this._removeClass(e,"ui-accordion-content-active"),this._removeClass(i,"ui-accordion-header-active")._addClass(i,"ui-accordion-header-collapsed"),this._trigger("activate",null,t)}}),t.widget("ui.menu",{version:"1.14.2",defaultElement:"<ul>",delay:300,options:{icons:{submenu:"ui-icon-caret-1-e"},items:"> *",menus:"ul",position:{my:"left top",at:"right top"},role:"menu",blur:null,focus:null,select:null},_create:function(){this.activeMenu=this.element,this.mouseHandled=!1,this.lastMousePosition={x:null,y:null},this.element.uniqueId().attr({role:this.options.role,tabIndex:0}),this._addClass("ui-menu","ui-widget ui-widget-content"),this._on({"mousedown .ui-menu-item":function(t){t.preventDefault(),this._activateItem(t)},"click .ui-menu-item":function(e){var i=t(e.target),s=t(this.document[0].activeElement);!this.mouseHandled&&i.not(".ui-state-disabled").length&&(this.select(e),e.isPropagationStopped()||(this.mouseHandled=!0),i.has(".ui-menu").length?this.expand(e):!this.element.is(":focus")&&s.closest(".ui-menu").length&&(this.element.trigger("focus",[!0]),this.active&&1===this.active.parents(".ui-menu").length&&clearTimeout(this.timer)))},"mouseenter .ui-menu-item":"_activateItem","mousemove .ui-menu-item":"_activateItem",mouseleave:"collapseAll","mouseleave .ui-menu":"collapseAll",focus:function(t,e){var i=this.active||this._menuItems().first();e||this.focus(t,i)},blur:function(e){this._delay(function(){t.contains(this.element[0],this.document[0].activeElement)||this.collapseAll(e)})},keydown:"_keydown"}),this.refresh(),this._on(this.document,{click:function(t){this._closeOnDocumentClick(t)&&this.collapseAll(t,!0),this.mouseHandled=!1}})},_activateItem:function(e){if(!this.previousFilter&&(e.clientX!==this.lastMousePosition.x||e.clientY!==this.lastMousePosition.y)){this.lastMousePosition={x:e.clientX,y:e.clientY};var i=t(e.target).closest(".ui-menu-item"),s=t(e.currentTarget);!(i[0]!==s[0]||s.is(".ui-state-active"))&&(this._removeClass(s.siblings().children(".ui-state-active"),null,"ui-state-active"),this.focus(e,s))}},_destroy:function(){var e=this.element.find(".ui-menu-item").removeAttr("role aria-disabled").children(".ui-menu-item-wrapper").removeUniqueId().removeAttr("tabIndex role aria-haspopup");this.element.removeAttr("aria-activedescendant").find(".ui-menu").addBack().removeAttr("role aria-labelledby aria-expanded aria-hidden aria-disabled tabIndex").removeUniqueId().show(),e.children().each(function(){var e=t(this);e.data("ui-menu-submenu-caret")&&e.remove()})},_keydown:function(e){var i,s,n,o,a=!0;switch(e.keyCode){case t.ui.keyCode.PAGE_UP:this.previousPage(e);break;case t.ui.keyCode.PAGE_DOWN:this.nextPage(e);break;case t.ui.keyCode.HOME:this._move("first","first",e);break;case t.ui.keyCode.END:this._move("last","last",e);break;case t.ui.keyCode.UP:this.previous(e);break;case t.ui.keyCode.DOWN:this.next(e);break;case t.ui.keyCode.LEFT:this.collapse(e);break;case t.ui.keyCode.RIGHT:this.active&&!this.active.is(".ui-state-disabled")&&this.expand(e);break;case t.ui.keyCode.ENTER:case t.ui.keyCode.SPACE:this._activate(e);break;case t.ui.keyCode.ESCAPE:this.collapse(e);break;default:a=!1,s=this.previousFilter||"",o=!1,n=e.keyCode>=96&&e.keyCode<=105?(e.keyCode-96).toString():String.fromCharCode(e.keyCode),clearTimeout(this.filterTimer),n===s?o=!0:n=s+n,i=this._filterMenuItems(n),(i=o&&-1!==i.index(this.active.next())?this.active.nextAll(".ui-menu-item"):i).length||(n=String.fromCharCode(e.keyCode),i=this._filterMenuItems(n)),i.length?(this.focus(e,i),this.previousFilter=n,this.filterTimer=this._delay(function(){delete this.previousFilter},1e3)):delete this.previousFilter}a&&e.preventDefault()},_activate:function(t){this.active&&!this.active.is(".ui-state-disabled")&&(this.active.children("[aria-haspopup='true']").length?this.expand(t):this.select(t))},refresh:function(){var e,i,s,n,o=this,a=this.options.icons.submenu,r=this.element.find(this.options.menus);this._toggleClass("ui-menu-icons",null,!!this.element.find(".ui-icon").length),i=r.filter(":not(.ui-menu)").hide().attr({role:this.options.role,"aria-hidden":"true","aria-expanded":"false"}).each(function(){var e=t(this),i=e.prev(),s=t("<span>").data("ui-menu-submenu-caret",!0);o._addClass(s,"ui-menu-icon","ui-icon "+a),i.attr("aria-haspopup","true").prepend(s),e.attr("aria-labelledby",i.attr("id"))}),this._addClass(i,"ui-menu","ui-widget ui-widget-content ui-front"),(e=r.add(this.element).find(this.options.items)).not(".ui-menu-item").each(function(){var e=t(this);o._isDivider(e)&&o._addClass(e,"ui-menu-divider","ui-widget-content")}),n=(s=e.not(".ui-menu-item, .ui-menu-divider")).children().not(".ui-menu").uniqueId().attr({tabIndex:-1,role:this._itemRole()}),this._addClass(s,"ui-menu-item")._addClass(n,"ui-menu-item-wrapper"),e.filter(".ui-state-disabled").attr("aria-disabled","true"),this.active&&!t.contains(this.element[0],this.active[0])&&this.blur()},_itemRole:function(){return({menu:"menuitem",listbox:"option"})[this.options.role]},_setOption:function(t,e){if("icons"===t){var i=this.element.find(".ui-menu-icon");this._removeClass(i,null,this.options.icons.submenu)._addClass(i,null,e.submenu)}this._super(t,e)},_setOptionDisabled:function(t){this._super(t),this.element.attr("aria-disabled",String(t)),this._toggleClass(null,"ui-state-disabled",!!t)},focus:function(t,e){var i,s,n;this.blur(t,t&&"focus"===t.type),this._scrollIntoView(e),this.active=e.first(),s=this.active.children(".ui-menu-item-wrapper"),this._addClass(s,null,"ui-state-active"),this.options.role&&this.element.attr("aria-activedescendant",s.attr("id")),n=this.active.parent().closest(".ui-menu-item").children(".ui-menu-item-wrapper"),this._addClass(n,null,"ui-state-active"),t&&"keydown"===t.type?this._close():this.timer=this._delay(function(){this._close()},this.delay),(i=e.children(".ui-menu")).length&&t&&/^mouse/.test(t.type)&&this._startOpening(i),this.activeMenu=e.parent(),this._trigger("focus",t,{item:e})},_scrollIntoView:function(e){var i,s,n,o,a,r;this._hasScroll()&&(i=parseFloat(t.css(this.activeMenu[0],"borderTopWidth"))||0,s=parseFloat(t.css(this.activeMenu[0],"paddingTop"))||0,n=e.offset().top-this.activeMenu.offset().top-i-s,o=this.activeMenu.scrollTop(),a=this.activeMenu.height(),r=e.outerHeight(),n<0?this.activeMenu.scrollTop(o+n):n+r>a&&this.activeMenu.scrollTop(o+n-a+r))},blur:function(t,e){e||clearTimeout(this.timer),this.active&&(this._removeClass(this.active.children(".ui-menu-item-wrapper"),null,"ui-state-active"),this._trigger("blur",t,{item:this.active}),this.active=null)},_startOpening:function(t){clearTimeout(this.timer),"true"===t.attr("aria-hidden")&&(this.timer=this._delay(function(){this._close(),this._open(t)},this.delay))},_open:function(e){var i=t.extend({of:this.active},this.options.position);clearTimeout(this.timer),this.element.find(".ui-menu").not(e.parents(".ui-menu")).hide().attr("aria-hidden","true"),e.show().removeAttr("aria-hidden").attr("aria-expanded","true").position(i)},collapseAll:function(e,i){clearTimeout(this.timer),this.timer=this._delay(function(){var s=i?this.element:t(e&&e.target).closest(this.element.find(".ui-menu"));s.length||(s=this.element),this._close(s),this.blur(e),this._removeClass(s.find(".ui-state-active"),null,"ui-state-active"),this.activeMenu=s},i?0:this.delay)},_close:function(t){t||(t=this.active?this.active.parent():this.element),t.find(".ui-menu").hide().attr("aria-hidden","true").attr("aria-expanded","false")},_closeOnDocumentClick:function(e){return!t(e.target).closest(".ui-menu").length},_isDivider:function(t){return!/[^\-\u2014\u2013\s]/.test(t.text())},collapse:function(t){var e=this.active&&this.active.parent().closest(".ui-menu-item",this.element);e&&e.length&&(this._close(),this.focus(t,e))},expand:function(t){var e=this.active&&this._menuItems(this.active.children(".ui-menu")).first();e&&e.length&&(this._open(e.parent()),this._delay(function(){this.focus(t,e)}))},next:function(t){this._move("next","first",t)},previous:function(t){this._move("prev","last",t)},isFirstItem:function(){return this.active&&!this.active.prevAll(".ui-menu-item").length},isLastItem:function(){return this.active&&!this.active.nextAll(".ui-menu-item").length},_menuItems:function(t){return(t||this.element).find(this.options.items).filter(".ui-menu-item")},_move:function(t,e,i){var s;this.active&&(s="first"===t||"last"===t?this.active["first"===t?"prevAll":"nextAll"](".ui-menu-item").last():this.active[t+"All"](".ui-menu-item").first()),s&&s.length&&this.active||(s=this._menuItems(this.activeMenu)[e]()),this.focus(i,s)},nextPage:function(e){var i,s,n;if(!this.active){this.next(e);return}!this.isLastItem()&&(this._hasScroll()?(s=this.active.offset().top,n=this.element.innerHeight(),0===t.fn.jquery.indexOf("3.2.")&&(n+=this.element[0].offsetHeight-this.element.outerHeight()),this.active.nextAll(".ui-menu-item").each(function(){return(i=t(this)).offset().top-s-n<0}),this.focus(e,i)):this.focus(e,this._menuItems(this.activeMenu)[this.active?"last":"first"]()))},previousPage:function(e){var i,s,n;if(!this.active){this.next(e);return}!this.isFirstItem()&&(this._hasScroll()?(s=this.active.offset().top,n=this.element.innerHeight(),0===t.fn.jquery.indexOf("3.2.")&&(n+=this.element[0].offsetHeight-this.element.outerHeight()),this.active.prevAll(".ui-menu-item").each(function(){return(i=t(this)).offset().top-s+n>0}),this.focus(e,i)):this.focus(e,this._menuItems(this.activeMenu).first()))},_hasScroll:function(){return this.element.outerHeight()<this.element.prop("scrollHeight")},select:function(e){this.active=this.active||t(e.target).closest(".ui-menu-item");var i={item:this.active};this.active.has(".ui-menu").length||this.collapseAll(e,!0),this._trigger("select",e,i)},_filterMenuItems:function(e){var i=RegExp("^"+e.replace(/[\-\[\]{}()*+?.,\\\^$|#\s]/g,"\\$&"),"i");return this.activeMenu.find(this.options.items).filter(".ui-menu-item").filter(function(){return i.test(String.prototype.trim.call(t(this).children(".ui-menu-item-wrapper").text()))})}}),t.widget("ui.autocomplete",{version:"1.14.2",defaultElement:"<input>",options:{appendTo:null,autoFocus:!1,delay:300,minLength:1,position:{my:"left top",at:"left bottom",collision:"none"},source:null,change:null,close:null,focus:null,open:null,response:null,search:null,select:null},requestIndex:0,pending:0,liveRegionTimer:null,_create:function(){var e,i,s,n=this.element[0].nodeName.toLowerCase(),o="textarea"===n,a="input"===n;this.isMultiLine=o||!a&&"true"===this.element.prop("contentEditable"),this.valueMethod=this.element[o||a?"val":"text"],this.isNewMenu=!0,this._addClass("ui-autocomplete-input"),this.element.attr("autocomplete","off"),this._on(this.element,{keydown:function(n){if(this.element.prop("readOnly")){e=!0,s=!0,i=!0;return}e=!1,s=!1,i=!1;var o=t.ui.keyCode;switch(n.keyCode){case o.PAGE_UP:e=!0,this._move("previousPage",n);break;case o.PAGE_DOWN:e=!0,this._move("nextPage",n);break;case o.UP:e=!0,this._keyEvent("previous",n);break;case o.DOWN:e=!0,this._keyEvent("next",n);break;case o.ENTER:this.menu.active&&(e=!0,n.preventDefault(),this.menu.select(n));break;case o.TAB:this.menu.active&&this.menu.select(n);break;case o.ESCAPE:this.menu.element.is(":visible")&&(this.isMultiLine||this._value(this.term),this.close(n),n.preventDefault());break;default:i=!0,this._searchTimeout(n)}},keypress:function(s){if(e){e=!1,(!this.isMultiLine||this.menu.element.is(":visible"))&&s.preventDefault();return}if(!i){var n=t.ui.keyCode;switch(s.keyCode){case n.PAGE_UP:this._move("previousPage",s);break;case n.PAGE_DOWN:this._move("nextPage",s);break;case n.UP:this._keyEvent("previous",s);break;case n.DOWN:this._keyEvent("next",s)}}},input:function(t){if(s){s=!1,t.preventDefault();return}this._searchTimeout(t)},focus:function(){this.selectedItem=null,this.previous=this._value()},blur:function(t){clearTimeout(this.searching),this.close(t),this._change(t)}}),this._initSource(),this.menu=t("<ul>").appendTo(this._appendTo()).menu({role:null}).hide().menu("instance"),this._addClass(this.menu.element,"ui-autocomplete","ui-front"),this._on(this.menu.element,{mousedown:function(t){t.preventDefault()},menufocus:function(e,i){var s,n;if(this.isNewMenu&&(this.isNewMenu=!1,e.originalEvent&&/^mouse/.test(e.originalEvent.type))){this.menu.blur(),this.document.one("mousemove",function(){t(e.target).trigger(e.originalEvent)});return}n=i.item.data("ui-autocomplete-item"),!1!==this._trigger("focus",e,{item:n})&&e.originalEvent&&/^key/.test(e.originalEvent.type)&&this._value(n.value),(s=i.item.attr("aria-label")||n.value)&&String.prototype.trim.call(s).length&&(clearTimeout(this.liveRegionTimer),this.liveRegionTimer=this._delay(function(){this.liveRegion.html(t("<div>").text(s))},100))},menuselect:function(t,e){var i=e.item.data("ui-autocomplete-item"),s=this.previous;this.element[0]!==this.document[0].activeElement&&(this.element.trigger("focus"),this.previous=s),!1!==this._trigger("select",t,{item:i})&&this._value(i.value),this.term=this._value(),this.close(t),this.selectedItem=i}}),this.liveRegion=t("<div>",{role:"status","aria-live":"assertive","aria-relevant":"additions"}).appendTo(this.document[0].body),this._addClass(this.liveRegion,null,"ui-helper-hidden-accessible"),this._on(this.window,{beforeunload:function(){this.element.removeAttr("autocomplete")}})},_destroy:function(){clearTimeout(this.searching),this.element.removeAttr("autocomplete"),this.menu.element.remove(),this.liveRegion.remove()},_setOption:function(t,e){this._super(t,e),"source"===t&&this._initSource(),"appendTo"===t&&this.menu.element.appendTo(this._appendTo()),"disabled"===t&&e&&this.xhr&&this.xhr.abort()},_isEventTargetInWidget:function(e){var i=this.menu.element[0];return e.target===this.element[0]||e.target===i||t.contains(i,e.target)},_closeOnClickOutside:function(t){this._isEventTargetInWidget(t)||this.close()},_appendTo:function(){var e=this.options.appendTo;return e&&(e=e.jquery||e.nodeType?t(e):this.document.find(e).eq(0)),e&&e[0]||(e=this.element.closest(".ui-front, dialog")),e.length||(e=this.document[0].body),e},_initSource:function(){var e,i,s=this;Array.isArray(this.options.source)?(e=this.options.source,this.source=function(i,s){s(t.ui.autocomplete.filter(e,i.term))}):"string"==typeof this.options.source?(i=this.options.source,this.source=function(e,n){s.xhr&&s.xhr.abort(),s.xhr=t.ajax({url:i,data:e,dataType:"json",success:function(t){n(t)},error:function(){n([])}})}):this.source=this.options.source},_searchTimeout:function(t){clearTimeout(this.searching),this.searching=this._delay(function(){var e=this.term===this._value(),i=this.menu.element.is(":visible"),s=t.altKey||t.ctrlKey||t.metaKey||t.shiftKey;e&&(!e||i||s)||(this.selectedItem=null,this.search(null,t))},this.options.delay)},search:function(t,e){return(t=null!=t?t:this._value(),this.term=this._value(),t.length<this.options.minLength)?this.close(e):!1!==this._trigger("search",e)?this._search(t):void 0},_search:function(t){this.pending++,this._addClass("ui-autocomplete-loading"),this.cancelSearch=!1,this.source({term:t},this._response())},_response:function(){var t=++this.requestIndex;return(function(e){t===this.requestIndex&&this.__response(e),this.pending--,this.pending||this._removeClass("ui-autocomplete-loading")}).bind(this)},__response:function(t){t&&(t=this._normalize(t)),this._trigger("response",null,{content:t}),!this.options.disabled&&t&&t.length&&!this.cancelSearch?(this._suggest(t),this._trigger("open")):this._close()},close:function(t){this.cancelSearch=!0,this._close(t)},_close:function(t){this._off(this.document,"mousedown"),this.menu.element.is(":visible")&&(this.menu.element.hide(),this.menu.blur(),this.isNewMenu=!0,this._trigger("close",t))},_change:function(t){this.previous!==this._value()&&this._trigger("change",t,{item:this.selectedItem})},_normalize:function(e){return e.length&&e[0].label&&e[0].value?e:t.map(e,function(e){return"string"==typeof e?{label:e,value:e}:t.extend({},e,{label:e.label||e.value,value:e.value||e.label})})},_suggest:function(e){var i=this.menu.element.empty();this._renderMenu(i,e),this.isNewMenu=!0,this.menu.refresh(),i.show(),this._resizeMenu(),i.position(t.extend({of:this.element},this.options.position)),this.options.autoFocus&&this.menu.next(),this._on(this.document,{mousedown:"_closeOnClickOutside"})},_resizeMenu:function(){var t=this.menu.element;t.outerWidth(Math.max(t.width("").outerWidth()+1,this.element.outerWidth()))},_renderMenu:function(e,i){var s=this;t.each(i,function(t,i){s._renderItemData(e,i)})},_renderItemData:function(t,e){return this._renderItem(t,e).data("ui-autocomplete-item",e)},_renderItem:function(e,i){return t("<li>").append(t("<div>").text(i.label)).appendTo(e)},_move:function(t,e){if(!this.menu.element.is(":visible")){this.search(null,e);return}if(this.menu.isFirstItem()&&/^previous/.test(t)||this.menu.isLastItem()&&/^next/.test(t)){this.isMultiLine||this._value(this.term),this.menu.blur();return}this.menu[t](e)},widget:function(){return this.menu.element},_value:function(){return this.valueMethod.apply(this.element,arguments)},_keyEvent:function(t,e){(!this.isMultiLine||this.menu.element.is(":visible"))&&(this._move(t,e),e.preventDefault())}}),t.extend(t.ui.autocomplete,{escapeRegex:function(t){return t.replace(/[\-\[\]{}()*+?.,\\\^$|#\s]/g,"\\$&")},filter:function(e,i){var s=RegExp(t.ui.autocomplete.escapeRegex(i),"i");return t.grep(e,function(t){return s.test(t.label||t.value||t)})}}),t.widget("ui.autocomplete",t.ui.autocomplete,{options:{messages:{noResults:"No search results.",results:function(t){return t+(t>1?" results are":" result is")+" available, use up and down arrow keys to navigate."}}},__response:function(e){var i;this._superApply(arguments),!this.options.disabled&&!this.cancelSearch&&(i=e&&e.length?this.options.messages.results(e.length):this.options.messages.noResults,clearTimeout(this.liveRegionTimer),this.liveRegionTimer=this._delay(function(){this.liveRegion.html(t("<div>").text(i))},100))}}),t.ui.autocomplete;var T=/ui-corner-([a-z]){2,6}/g;function P(){this._curInst=null,this._keyEvent=!1,this._disabledInputs=[],this._datepickerShowing=!1,this._inDialog=!1,this._mainDivId="ui-datepicker-div",this._inlineClass="ui-datepicker-inline",this._appendClass="ui-datepicker-append",this._triggerClass="ui-datepicker-trigger",this._dialogClass="ui-datepicker-dialog",this._disableClass="ui-datepicker-disabled",this._unselectableClass="ui-datepicker-unselectable",this._currentClass="ui-datepicker-current-day",this._dayOverClass="ui-datepicker-days-cell-over",this.regional=[],this.regional[""]={closeText:"Done",prevText:"Prev",nextText:"Next",currentText:"Today",monthNames:["January","February","March","April","May","June","July","August","September","October","November","December"],monthNamesShort:["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"],dayNames:["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"],dayNamesShort:["Sun","Mon","Tue","Wed","Thu","Fri","Sat"],dayNamesMin:["Su","Mo","Tu","We","Th","Fr","Sa"],weekHeader:"Wk",dateFormat:"mm/dd/yy",firstDay:0,isRTL:!1,showMonthAfterYear:!1,yearSuffix:"",selectMonthLabel:"Select month",selectYearLabel:"Select year"},this._defaults={showOn:"focus",showAnim:"fadeIn",showOptions:{},defaultDate:null,appendText:"",buttonText:"...",buttonImage:"",buttonImageOnly:!1,hideIfNoPrevNext:!1,navigationAsDateFormat:!1,gotoCurrent:!1,changeMonth:!1,changeYear:!1,yearRange:"c-10:c+10",showOtherMonths:!1,selectOtherMonths:!1,showWeek:!1,calculateWeek:this.iso8601Week,shortYearCutoff:"+10",minDate:null,maxDate:null,duration:"fast",beforeShowDay:null,beforeShow:null,onSelect:null,onChangeMonthYear:null,onClose:null,onUpdateDatepicker:null,numberOfMonths:1,showCurrentAtPos:0,stepMonths:1,stepBigMonths:12,altField:"",altFormat:"",constrainInput:!0,showButtonPanel:!1,autoSize:!1,disabled:!1},t.extend(this._defaults,this.regional[""]),this.regional.en=t.extend(!0,{},this.regional[""]),this.regional["en-US"]=t.extend(!0,{},this.regional.en),this.dpDiv=M(t("<div id='"+this._mainDivId+"' class='ui-datepicker ui-widget ui-widget-content ui-helper-clearfix ui-corner-all'></div>"))}function M(e){var i="button, .ui-datepicker-prev, .ui-datepicker-next, .ui-datepicker-calendar td a";return e.on("mouseout",i,function(){t(this).removeClass("ui-state-hover"),-1!==this.className.indexOf("ui-datepicker-prev")&&t(this).removeClass("ui-datepicker-prev-hover"),-1!==this.className.indexOf("ui-datepicker-next")&&t(this).removeClass("ui-datepicker-next-hover")}).on("mouseover",i,S)}function S(){t.datepicker._isDisabledDatepicker(o.inline?o.dpDiv.parent()[0]:o.input[0])||(t(this).parents(".ui-datepicker-calendar").find("a").removeClass("ui-state-hover"),t(this).addClass("ui-state-hover"),-1!==this.className.indexOf("ui-datepicker-prev")&&t(this).addClass("ui-datepicker-prev-hover"),-1!==this.className.indexOf("ui-datepicker-next")&&t(this).addClass("ui-datepicker-next-hover"))}function H(e,i){for(var s in t.extend(e,i),i)null==i[s]&&(e[s]=i[s]);return e}t.widget("ui.controlgroup",{version:"1.14.2",defaultElement:"<div>",options:{direction:"horizontal",disabled:null,onlyVisible:!0,items:{button:"input[type=button], input[type=submit], input[type=reset], button, a",controlgroupLabel:".ui-controlgroup-label",checkboxradio:"input[type='checkbox'], input[type='radio']",selectmenu:"select",spinner:".ui-spinner-input"}},_create:function(){this._enhance()},_enhance:function(){this.element.attr("role","toolbar"),this.refresh()},_destroy:function(){this._callChildMethod("destroy"),this.childWidgets.removeData("ui-controlgroup-data"),this.element.removeAttr("role"),this.options.items.controlgroupLabel&&this.element.find(this.options.items.controlgroupLabel).find(".ui-controlgroup-label-contents").contents().unwrap()},_initWidgets:function(){var e=this,i=[];t.each(this.options.items,function(s,n){var o,a={};if(n){if("controlgroupLabel"===s){(o=e.element.find(n)).each(function(){var e=t(this);!e.children(".ui-controlgroup-label-contents").length&&e.contents().wrapAll("<span class='ui-controlgroup-label-contents'></span>")}),e._addClass(o,null,"ui-widget ui-widget-content ui-state-default"),i=i.concat(o.get());return}t.fn[s]&&(a=e["_"+s+"Options"]?e["_"+s+"Options"]("middle"):{classes:{}},e.element.find(n).each(function(){var n=t(this),o=n[s]("instance"),r=t.widget.extend({},a);if("button"!==s||!n.parent(".ui-spinner").length){o||(o=n[s]()[s]("instance")),o&&(r.classes=e._resolveClassesValues(r.classes,o)),n[s](r);var l=n[s]("widget");t.data(l[0],"ui-controlgroup-data",o||n[s]("instance")),i.push(l[0])}}))}}),this.childWidgets=t(t.uniqueSort(i)),this._addClass(this.childWidgets,"ui-controlgroup-item")},_callChildMethod:function(e){this.childWidgets.each(function(){var i=t(this).data("ui-controlgroup-data");i&&i[e]&&i[e]()})},_updateCornerClass:function(t,e){var i=this._buildSimpleOptions(e,"label").classes.label;this._removeClass(t,null,"ui-corner-top ui-corner-bottom ui-corner-left ui-corner-right ui-corner-all"),this._addClass(t,null,i)},_buildSimpleOptions:function(t,e){var i="vertical"===this.options.direction,s={classes:{}};return s.classes[e]=({middle:"",first:"ui-corner-"+(i?"top":"left"),last:"ui-corner-"+(i?"bottom":"right"),only:"ui-corner-all"})[t],s},_spinnerOptions:function(t){var e=this._buildSimpleOptions(t,"ui-spinner");return e.classes["ui-spinner-up"]="",e.classes["ui-spinner-down"]="",e},_buttonOptions:function(t){return this._buildSimpleOptions(t,"ui-button")},_checkboxradioOptions:function(t){return this._buildSimpleOptions(t,"ui-checkboxradio-label")},_selectmenuOptions:function(t){var e="vertical"===this.options.direction;return{width:!!e&&"auto",classes:({middle:{"ui-selectmenu-button-open":"","ui-selectmenu-button-closed":""},first:{"ui-selectmenu-button-open":"ui-corner-"+(e?"top":"tl"),"ui-selectmenu-button-closed":"ui-corner-"+(e?"top":"left")},last:{"ui-selectmenu-button-open":e?"":"ui-corner-tr","ui-selectmenu-button-closed":"ui-corner-"+(e?"bottom":"right")},only:{"ui-selectmenu-button-open":"ui-corner-top","ui-selectmenu-button-closed":"ui-corner-all"}})[t]}},_resolveClassesValues:function(e,i){var s={};return t.each(e,function(t){var n=i.options.classes[t]||"";n=String.prototype.trim.call(n.replace(T,"")),s[t]=(n+" "+e[t]).replace(/\s+/g," ")}),s},_setOption:function(t,e){if("direction"===t&&this._removeClass("ui-controlgroup-"+this.options.direction),this._super(t,e),"disabled"===t){this._callChildMethod(e?"disable":"enable");return}this.refresh()},refresh:function(){var e,i=this;this._addClass("ui-controlgroup ui-controlgroup-"+this.options.direction),"horizontal"===this.options.direction&&this._addClass(null,"ui-helper-clearfix"),this._initWidgets(),e=this.childWidgets,this.options.onlyVisible&&(e=e.filter(":visible")),e.length&&(t.each(["first","last"],function(t,s){var n=e[s]().data("ui-controlgroup-data");if(n&&i["_"+n.widgetName+"Options"]){var o=i["_"+n.widgetName+"Options"](1===e.length?"only":s);o.classes=i._resolveClassesValues(o.classes,n),n.element[n.widgetName](o)}else i._updateCornerClass(e[s](),s)}),this._callChildMethod("refresh"))}}),t.widget("ui.checkboxradio",[t.ui.formResetMixin,{version:"1.14.2",options:{disabled:null,label:null,icon:!0,classes:{"ui-checkboxradio-label":"ui-corner-all","ui-checkboxradio-icon":"ui-corner-all"}},_getCreateOptions:function(){var e,i,s,n=this._super()||{};return this._readType(),i=this.element.labels(),this.label=t(i[i.length-1]),this.label.length||t.error("No label found for checkboxradio widget"),this.originalLabel="",(s=this.label.contents().not(this.element[0])).length&&(this.originalLabel+=s.clone().wrapAll("<div></div>").parent().html()),this.originalLabel&&(n.label=this.originalLabel),null!=(e=this.element[0].disabled)&&(n.disabled=e),n},_create:function(){var t=this.element[0].checked;this._bindFormResetHandler(),null==this.options.disabled&&(this.options.disabled=this.element[0].disabled),this._setOption("disabled",this.options.disabled),this._addClass("ui-checkboxradio","ui-helper-hidden-accessible"),this._addClass(this.label,"ui-checkboxradio-label","ui-button ui-widget"),"radio"===this.type&&this._addClass(this.label,"ui-checkboxradio-radio-label"),this.options.label&&this.options.label!==this.originalLabel?this._updateLabel():this.originalLabel&&(this.options.label=this.originalLabel),this._enhance(),t&&this._addClass(this.label,"ui-checkboxradio-checked","ui-state-active"),this._on({change:"_toggleClasses",focus:function(){this._addClass(this.label,null,"ui-state-focus ui-visual-focus")},blur:function(){this._removeClass(this.label,null,"ui-state-focus ui-visual-focus")}})},_readType:function(){var e=this.element[0].nodeName.toLowerCase();this.type=this.element[0].type,"input"===e&&/radio|checkbox/.test(this.type)||t.error("Can't create checkboxradio on element.nodeName="+e+" and element.type="+this.type)},_enhance:function(){this._updateIcon(this.element[0].checked)},widget:function(){return this.label},_getRadioGroup:function(){var e,i=this.element[0].name,s="input[name='"+CSS.escape(i)+"']";return i?(this.form.length?t(this.form[0].elements).filter(s):t(s).filter(function(){return 0===t(t(this).prop("form")).length})).not(this.element):t([])},_toggleClasses:function(){var e=this.element[0].checked;this._toggleClass(this.label,"ui-checkboxradio-checked","ui-state-active",e),this.options.icon&&"checkbox"===this.type&&this._toggleClass(this.icon,null,"ui-icon-check ui-state-checked",e)._toggleClass(this.icon,null,"ui-icon-blank",!e),"radio"===this.type&&this._getRadioGroup().each(function(){var e=t(this).checkboxradio("instance");e&&e._removeClass(e.label,"ui-checkboxradio-checked","ui-state-active")})},_destroy:function(){this._unbindFormResetHandler(),this.icon&&(this.icon.remove(),this.iconSpace.remove())},_setOption:function(t,e){if("label"!==t||e){if(this._super(t,e),"disabled"===t){this._toggleClass(this.label,null,"ui-state-disabled",e),this.element[0].disabled=e;return}this.refresh()}},_updateIcon:function(e){var i="ui-icon ui-icon-background ";this.options.icon?(this.icon||(this.icon=t("<span>"),this.iconSpace=t("<span> </span>"),this._addClass(this.iconSpace,"ui-checkboxradio-icon-space")),"checkbox"===this.type?(i+=e?"ui-icon-check ui-state-checked":"ui-icon-blank",this._removeClass(this.icon,null,e?"ui-icon-blank":"ui-icon-check")):i+="ui-icon-blank",this._addClass(this.icon,"ui-checkboxradio-icon",i),e||this._removeClass(this.icon,null,"ui-icon-check ui-state-checked"),this.icon.prependTo(this.label).after(this.iconSpace)):void 0!==this.icon&&(this.icon.remove(),this.iconSpace.remove(),delete this.icon)},_updateLabel:function(){var t=this.label.contents().not(this.element[0]);this.icon&&(t=t.not(this.icon[0])),this.iconSpace&&(t=t.not(this.iconSpace[0])),t.remove(),this.label.append(this.options.label)},refresh:function(){var t=this.element[0].checked,e=this.element[0].disabled;this._updateIcon(t),this._toggleClass(this.label,"ui-checkboxradio-checked","ui-state-active",t),null!==this.options.label&&this._updateLabel(),e!==this.options.disabled&&this._setOptions({disabled:e})}}]),t.ui.checkboxradio,t.widget("ui.button",{version:"1.14.2",defaultElement:"<button>",options:{classes:{"ui-button":"ui-corner-all"},disabled:null,icon:null,iconPosition:"beginning",label:null,showLabel:!0},_getCreateOptions:function(){var t,e=this._super()||{};return this.isInput=this.element.is("input"),null!=(t=this.element[0].disabled)&&(e.disabled=t),this.originalLabel=this.isInput?this.element.val():this.element.html(),this.originalLabel&&(e.label=this.originalLabel),e},_create:function(){!this.option.showLabel&!this.options.icon&&(this.options.showLabel=!0),null==this.options.disabled&&(this.options.disabled=this.element[0].disabled||!1),this.hasTitle=!!this.element.attr("title"),this.options.label&&this.options.label!==this.originalLabel&&(this.isInput?this.element.val(this.options.label):this.element.html(this.options.label)),this._addClass("ui-button","ui-widget"),this._setOption("disabled",this.options.disabled),this._enhance(),this.element.is("a")&&this._on({keyup:function(e){e.keyCode===t.ui.keyCode.SPACE&&(e.preventDefault(),this.element[0].click?this.element[0].click():this.element.trigger("click"))}})},_enhance:function(){this.element.is("button")||this.element.attr("role","button"),this.options.icon&&(this._updateIcon("icon",this.options.icon),this._updateTooltip())},_updateTooltip:function(){this.title=this.element.attr("title"),this.options.showLabel||this.title||this.element.attr("title",this.options.label)},_updateIcon:function(e,i){var s="iconPosition"!==e,n=s?this.options.iconPosition:i,o="top"===n||"bottom"===n;this.icon?s&&this._removeClass(this.icon,null,this.options.icon):(this.icon=t("<span>"),this._addClass(this.icon,"ui-button-icon","ui-icon"),this.options.showLabel||this._addClass("ui-button-icon-only")),s&&this._addClass(this.icon,null,i),this._attachIcon(n),o?(this._addClass(this.icon,null,"ui-widget-icon-block"),this.iconSpace&&this.iconSpace.remove()):(this.iconSpace||(this.iconSpace=t("<span> </span>"),this._addClass(this.iconSpace,"ui-button-icon-space")),this._removeClass(this.icon,null,"ui-wiget-icon-block"),this._attachIconSpace(n))},_destroy:function(){this.element.removeAttr("role"),this.icon&&this.icon.remove(),this.iconSpace&&this.iconSpace.remove(),this.hasTitle||this.element.removeAttr("title")},_attachIconSpace:function(t){this.icon[/^(?:end|bottom)/.test(t)?"before":"after"](this.iconSpace)},_attachIcon:function(t){this.element[/^(?:end|bottom)/.test(t)?"append":"prepend"](this.icon)},_setOptions:function(t){var e=void 0===t.showLabel?this.options.showLabel:t.showLabel,i=void 0===t.icon?this.options.icon:t.icon;e||i||(t.showLabel=!0),this._super(t)},_setOption:function(t,e){"icon"===t&&(e?this._updateIcon(t,e):this.icon&&(this.icon.remove(),this.iconSpace&&this.iconSpace.remove())),"iconPosition"===t&&this._updateIcon(t,e),"showLabel"===t&&(this._toggleClass("ui-button-icon-only",null,!e),this._updateTooltip()),"label"===t&&(this.isInput?this.element.val(e):(this.element.html(e),this.icon&&(this._attachIcon(this.options.iconPosition),this._attachIconSpace(this.options.iconPosition)))),this._super(t,e),"disabled"===t&&(this._toggleClass(null,"ui-state-disabled",e),this.element[0].disabled=e,e&&this.element.trigger("blur"))},refresh:function(){var t=this.element.is("input, button")?this.element[0].disabled:this.element.hasClass("ui-button-disabled");t!==this.options.disabled&&this._setOptions({disabled:t}),this._updateTooltip()}}),!0===t.uiBackCompat&&(t.widget("ui.button",t.ui.button,{options:{text:!0,icons:{primary:null,secondary:null}},_create:function(){this.options.showLabel&&!this.options.text&&(this.options.showLabel=this.options.text),!this.options.showLabel&&this.options.text&&(this.options.text=this.options.showLabel),!this.options.icon&&(this.options.icons.primary||this.options.icons.secondary)?this.options.icons.primary?this.options.icon=this.options.icons.primary:(this.options.icon=this.options.icons.secondary,this.options.iconPosition="end"):this.options.icon&&(this.options.icons.primary=this.options.icon),this._super()},_setOption:function(t,e){if("text"===t){this._super("showLabel",e);return}"showLabel"===t&&(this.options.text=e),"icon"===t&&(this.options.icons.primary=e),"icons"===t&&(e.primary?(this._super("icon",e.primary),this._super("iconPosition","beginning")):e.secondary&&(this._super("icon",e.secondary),this._super("iconPosition","end"))),this._superApply(arguments)}}),t.fn.button=(s=t.fn.button,function(e){var i="string"==typeof e,n=Array.prototype.slice.call(arguments,1),o=this;return i?this.length||"instance"!==e?this.each(function(){var i,s=t(this).attr("type"),a=t.data(this,"ui-"+("checkbox"!==s&&"radio"!==s?"button":"checkboxradio"));return"instance"===e?(o=a,!1):a?"function"!=typeof a[e]||"_"===e.charAt(0)?t.error("no such method '"+e+"' for button widget instance"):(i=a[e].apply(a,n))!==a&&void 0!==i?(o=i&&i.jquery?o.pushStack(i.get()):i,!1):void 0:t.error("cannot call methods on button prior to initialization; attempted to call method '"+e+"'")}):o=void 0:(n.length&&(e=t.widget.extend.apply(null,[e].concat(n))),this.each(function(){var i=t(this).attr("type"),n="checkbox"!==i&&"radio"!==i?"button":"checkboxradio",o=t.data(this,"ui-"+n);if(o)o.option(e||{}),o._init&&o._init();else{if("button"===n){s.call(t(this),e);return}t(this).checkboxradio(t.extend({icon:!1},e))}})),o}),t.fn.buttonset=function(){return(t.ui.controlgroup||t.error("Controlgroup widget missing"),"option"===arguments[0]&&"items"===arguments[1]&&arguments[2])?this.controlgroup.apply(this,[arguments[0],"items.button",arguments[2]]):"option"===arguments[0]&&"items"===arguments[1]?this.controlgroup.apply(this,[arguments[0],"items.button"]):("object"==typeof arguments[0]&&arguments[0].items&&(arguments[0].items={button:arguments[0].items}),this.controlgroup.apply(this,arguments))}),t.ui.button,t.extend(t.ui,{datepicker:{version:"1.14.2"}}),t.extend(P.prototype,{markerClassName:"hasDatepicker",maxRows:4,_widgetDatepicker:function(){return this.dpDiv},setDefaults:function(t){return H(this._defaults,t||{}),this},_attachDatepicker:function(e,i){var s,n,o;n="div"===(s=e.nodeName.toLowerCase())||"span"===s,e.id||(this.uuid+=1,e.id="dp"+this.uuid),(o=this._newInst(t(e),n)).settings=t.extend({},i||{}),"input"===s?this._connectDatepicker(e,o):n&&this._inlineDatepicker(e,o)},_newInst:function(e,i){return{id:e[0].id.replace(/([^A-Za-z0-9_\-])/g,"\\\\$1"),input:e,selectedDay:0,selectedMonth:0,selectedYear:0,drawMonth:0,drawYear:0,inline:i,dpDiv:i?M(t("<div class='"+this._inlineClass+" ui-datepicker ui-widget ui-widget-content ui-helper-clearfix ui-corner-all'></div>")):this.dpDiv}},_connectDatepicker:function(e,i){var s=t(e);i.append=t([]),i.trigger=t([]),!s.hasClass(this.markerClassName)&&(this._attachments(s,i),s.addClass(this.markerClassName).on("keydown",this._doKeyDown).on("keypress",this._doKeyPress).on("keyup",this._doKeyUp),this._autoSize(i),t.data(e,"datepicker",i),i.settings.disabled&&this._disableDatepicker(e))},_attachments:function(e,i){var s,n,o,a=this._get(i,"appendText"),r=this._get(i,"isRTL");i.append&&i.append.remove(),a&&(i.append=t("<span>").addClass(this._appendClass).text(a),e[r?"before":"after"](i.append)),e.off("focus",this._showDatepicker),i.trigger&&i.trigger.remove(),("focus"===(s=this._get(i,"showOn"))||"both"===s)&&e.on("focus",this._showDatepicker),("button"===s||"both"===s)&&(n=this._get(i,"buttonText"),o=this._get(i,"buttonImage"),this._get(i,"buttonImageOnly")?i.trigger=t("<img>").addClass(this._triggerClass).attr({src:o,alt:n,title:n}):(i.trigger=t("<button type='button'>").addClass(this._triggerClass),o?i.trigger.html(t("<img>").attr({src:o,alt:n,title:n})):i.trigger.text(n)),e[r?"before":"after"](i.trigger),i.trigger.on("click",function(){return t.datepicker._datepickerShowing&&t.datepicker._lastInput===e[0]?t.datepicker._hideDatepicker():(t.datepicker._datepickerShowing&&t.datepicker._lastInput!==e[0]&&t.datepicker._hideDatepicker(),t.datepicker._showDatepicker(e[0])),!1}))},_autoSize:function(t){if(this._get(t,"autoSize")&&!t.inline){var e,i,s,n,o=new Date(2009,11,20),a=this._get(t,"dateFormat");a.match(/[DM]/)&&(e=function(t){for(n=0,i=0,s=0;n<t.length;n++)t[n].length>i&&(i=t[n].length,s=n);return s},o.setMonth(e(this._get(t,a.match(/MM/)?"monthNames":"monthNamesShort"))),o.setDate(e(this._get(t,a.match(/DD/)?"dayNames":"dayNamesShort"))+20-o.getDay())),t.input.attr("size",this._formatDate(t,o).length)}},_inlineDatepicker:function(e,i){var s=t(e);!s.hasClass(this.markerClassName)&&(s.addClass(this.markerClassName).append(i.dpDiv),t.data(e,"datepicker",i),this._setDate(i,this._getDefaultDate(i),!0),this._updateDatepicker(i),this._updateAlternate(i),i.settings.disabled&&this._disableDatepicker(e),i.dpDiv.css("display","block"))},_dialogDatepicker:function(e,i,s,n,o){var a,r,l,h,c,u=this._dialogInst;return u||(this.uuid+=1,a="dp"+this.uuid,this._dialogInput=t("<input type='text' id='"+a+"' style='position: absolute; top: -100px; width: 0px;'/>"),this._dialogInput.on("keydown",this._doKeyDown),t("body").append(this._dialogInput),(u=this._dialogInst=this._newInst(this._dialogInput,!1)).settings={},t.data(this._dialogInput[0],"datepicker",u)),H(u.settings,n||{}),i=i&&i.constructor===Date?this._formatDate(u,i):i,this._dialogInput.val(i),this._pos=o?o.length?o:[o.pageX,o.pageY]:null,this._pos||(r=document.documentElement.clientWidth,l=document.documentElement.clientHeight,h=document.documentElement.scrollLeft||document.body.scrollLeft,c=document.documentElement.scrollTop||document.body.scrollTop,this._pos=[r/2-100+h,l/2-150+c]),this._dialogInput.css("left",this._pos[0]+20+"px").css("top",this._pos[1]+"px"),u.settings.onSelect=s,this._inDialog=!0,this.dpDiv.addClass(this._dialogClass),this._showDatepicker(this._dialogInput[0]),t.blockUI&&t.blockUI(this.dpDiv),t.data(this._dialogInput[0],"datepicker",u),this},_destroyDatepicker:function(e){var i,s=t(e),n=t.data(e,"datepicker");s.hasClass(this.markerClassName)&&(i=e.nodeName.toLowerCase(),t.removeData(e,"datepicker"),"input"===i?(n.append.remove(),n.trigger.remove(),s.removeClass(this.markerClassName).off("focus",this._showDatepicker).off("keydown",this._doKeyDown).off("keypress",this._doKeyPress).off("keyup",this._doKeyUp)):("div"===i||"span"===i)&&s.removeClass(this.markerClassName).empty(),t.datepicker._hideDatepicker(),o===n&&(o=null,this._curInst=null))},_enableDatepicker:function(e){var i,s,n=t(e),o=t.data(e,"datepicker");n.hasClass(this.markerClassName)&&("input"===(i=e.nodeName.toLowerCase())?(e.disabled=!1,o.trigger.filter("button").each(function(){this.disabled=!1}).end().filter("img").css({opacity:"1.0",cursor:""})):("div"===i||"span"===i)&&((s=n.children("."+this._inlineClass)).children().removeClass("ui-state-disabled"),s.find("select.ui-datepicker-month, select.ui-datepicker-year").prop("disabled",!1)),this._disabledInputs=t.map(this._disabledInputs,function(t){return t===e?null:t}))},_disableDatepicker:function(e){var i,s,n=t(e),o=t.data(e,"datepicker");n.hasClass(this.markerClassName)&&("input"===(i=e.nodeName.toLowerCase())?(e.disabled=!0,o.trigger.filter("button").each(function(){this.disabled=!0}).end().filter("img").css({opacity:"0.5",cursor:"default"})):("div"===i||"span"===i)&&((s=n.children("."+this._inlineClass)).children().addClass("ui-state-disabled"),s.find("select.ui-datepicker-month, select.ui-datepicker-year").prop("disabled",!0)),this._disabledInputs=t.map(this._disabledInputs,function(t){return t===e?null:t}),this._disabledInputs[this._disabledInputs.length]=e)},_isDisabledDatepicker:function(t){if(!t)return!1;for(var e=0;e<this._disabledInputs.length;e++)if(this._disabledInputs[e]===t)return!0;return!1},_getInst:function(e){try{return t.data(e,"datepicker")}catch(t){throw"Missing instance data for this datepicker"}},_optionDatepicker:function(e,i,s){var n,o,a,r,l=this._getInst(e);if(2==arguments.length&&"string"==typeof i)return"defaults"===i?t.extend({},t.datepicker._defaults):l?"all"===i?t.extend({},l.settings):this._get(l,i):null;n=i||{},"string"==typeof i&&((n={})[i]=s),l&&(this._curInst===l&&this._hideDatepicker(),o=this._getDateDatepicker(e,!0),a=this._getMinMaxDate(l,"min"),r=this._getMinMaxDate(l,"max"),H(l.settings,n),null!==a&&void 0!==n.dateFormat&&void 0===n.minDate&&(l.settings.minDate=this._formatDate(l,a)),null!==r&&void 0!==n.dateFormat&&void 0===n.maxDate&&(l.settings.maxDate=this._formatDate(l,r)),"disabled"in n&&(n.disabled?this._disableDatepicker(e):this._enableDatepicker(e)),this._attachments(t(e),l),this._autoSize(l),this._setDate(l,o),this._updateAlternate(l),this._updateDatepicker(l))},_changeDatepicker:function(t,e,i){this._optionDatepicker(t,e,i)},_refreshDatepicker:function(t){var e=this._getInst(t);e&&this._updateDatepicker(e)},_setDateDatepicker:function(t,e){var i=this._getInst(t);i&&(this._setDate(i,e),this._updateDatepicker(i),this._updateAlternate(i))},_getDateDatepicker:function(t,e){var i=this._getInst(t);return i&&!i.inline&&this._setDateFromField(i,e),i?this._getDate(i):null},_doKeyDown:function(e){var i,s,n,o=t.datepicker._getInst(e.target),a=!0,r=o.dpDiv.is(".ui-datepicker-rtl");if(o._keyEvent=!0,t.datepicker._datepickerShowing)switch(e.keyCode){case 9:t.datepicker._hideDatepicker(),a=!1;break;case 13:return(n=t("td."+t.datepicker._dayOverClass+":not(."+t.datepicker._currentClass+")",o.dpDiv))[0]&&t.datepicker._selectDay(e.target,o.selectedMonth,o.selectedYear,n[0]),(i=t.datepicker._get(o,"onSelect"))?(s=t.datepicker._formatDate(o),i.apply(o.input?o.input[0]:null,[s,o])):t.datepicker._hideDatepicker(),!1;case 27:t.datepicker._hideDatepicker();break;case 33:t.datepicker._adjustDate(e.target,e.ctrlKey?-t.datepicker._get(o,"stepBigMonths"):-t.datepicker._get(o,"stepMonths"),"M");break;case 34:t.datepicker._adjustDate(e.target,e.ctrlKey?+t.datepicker._get(o,"stepBigMonths"):+t.datepicker._get(o,"stepMonths"),"M");break;case 35:(e.ctrlKey||e.metaKey)&&t.datepicker._clearDate(e.target),a=e.ctrlKey||e.metaKey;break;case 36:(e.ctrlKey||e.metaKey)&&t.datepicker._gotoToday(e.target),a=e.ctrlKey||e.metaKey;break;case 37:(e.ctrlKey||e.metaKey)&&t.datepicker._adjustDate(e.target,r?1:-1,"D"),a=e.ctrlKey||e.metaKey,e.originalEvent.altKey&&t.datepicker._adjustDate(e.target,e.ctrlKey?-t.datepicker._get(o,"stepBigMonths"):-t.datepicker._get(o,"stepMonths"),"M");break;case 38:(e.ctrlKey||e.metaKey)&&t.datepicker._adjustDate(e.target,-7,"D"),a=e.ctrlKey||e.metaKey;break;case 39:(e.ctrlKey||e.metaKey)&&t.datepicker._adjustDate(e.target,r?-1:1,"D"),a=e.ctrlKey||e.metaKey,e.originalEvent.altKey&&t.datepicker._adjustDate(e.target,e.ctrlKey?+t.datepicker._get(o,"stepBigMonths"):+t.datepicker._get(o,"stepMonths"),"M");break;case 40:(e.ctrlKey||e.metaKey)&&t.datepicker._adjustDate(e.target,7,"D"),a=e.ctrlKey||e.metaKey;break;default:a=!1}else 36===e.keyCode&&e.ctrlKey?t.datepicker._showDatepicker(this):a=!1;a&&(e.preventDefault(),e.stopPropagation())},_doKeyPress:function(e){var i,s,n=t.datepicker._getInst(e.target);if(t.datepicker._get(n,"constrainInput"))return i=t.datepicker._possibleChars(t.datepicker._get(n,"dateFormat")),s=String.fromCharCode(null==e.charCode?e.keyCode:e.charCode),e.ctrlKey||e.metaKey||s<" "||!i||i.indexOf(s)>-1},_doKeyUp:function(e){var i=t.datepicker._getInst(e.target);if(i.input.val()!==i.lastVal)try{t.datepicker.parseDate(t.datepicker._get(i,"dateFormat"),i.input?i.input.val():null,t.datepicker._getFormatConfig(i))&&(t.datepicker._setDateFromField(i),t.datepicker._updateAlternate(i),t.datepicker._updateDatepicker(i))}catch(t){}return!0},_showDatepicker:function(e){var i,s,n,o,a,r,l;if("input"!==(e=e.target||e).nodeName.toLowerCase()&&(e=t("input",e.parentNode)[0]),!t.datepicker._isDisabledDatepicker(e)&&t.datepicker._lastInput!==e)i=t.datepicker._getInst(e),t.datepicker._curInst&&t.datepicker._curInst!==i&&(t.datepicker._curInst.dpDiv.stop(!0,!0),i&&t.datepicker._datepickerShowing&&t.datepicker._hideDatepicker(t.datepicker._curInst.input[0])),!1!==(n=(s=t.datepicker._get(i,"beforeShow"))?s.apply(e,[e,i]):{})&&(H(i.settings,n),i.lastVal=null,t.datepicker._lastInput=e,t.datepicker._setDateFromField(i),t.datepicker._inDialog&&(e.value=""),t.datepicker._pos||(t.datepicker._pos=t.datepicker._findPos(e),t.datepicker._pos[1]+=e.offsetHeight),o=!1,t(e).parents().each(function(){return!(o|="fixed"===t(this).css("position"))}),a={left:t.datepicker._pos[0],top:t.datepicker._pos[1]},t.datepicker._pos=null,i.dpDiv.empty(),i.dpDiv.css({position:"absolute",display:"block",top:"-1000px"}),t.datepicker._updateDatepicker(i),a=t.datepicker._checkOffset(i,a,o),i.dpDiv.css({position:t.datepicker._inDialog&&t.blockUI?"static":o?"fixed":"absolute",display:"none",left:a.left+"px",top:a.top+"px"}),i.inline||(r=t.datepicker._get(i,"showAnim"),l=t.datepicker._get(i,"duration"),i.dpDiv.css("z-index",function(t){var e,i;while(t.length&&t[0]!==document){if(("absolute"===(e=t.css("position"))||"relative"===e||"fixed"===e)&&!isNaN(i=parseInt(t.css("zIndex"),10))&&0!==i)return i;t=t.parent()}return 0}(t(e))+1),t.datepicker._datepickerShowing=!0,t.effects&&t.effects.effect[r]?i.dpDiv.show(r,t.datepicker._get(i,"showOptions"),l):i.dpDiv[r||"show"](r?l:null),t.datepicker._shouldFocusInput(i)&&i.input.trigger("focus"),t.datepicker._curInst=i))},_updateDatepicker:function(e){this.maxRows=4,o=e,e.dpDiv.empty().append(this._generateHTML(e)),this._attachHandlers(e);var i,s=this._getNumberOfMonths(e),n=s[1],a=e.dpDiv.find("."+this._dayOverClass+" a"),r=t.datepicker._get(e,"onUpdateDatepicker");a.length>0&&S.apply(a.get(0)),e.dpDiv.removeClass("ui-datepicker-multi-2 ui-datepicker-multi-3 ui-datepicker-multi-4").width(""),n>1&&e.dpDiv.addClass("ui-datepicker-multi-"+n).css("width",17*n+"em"),e.dpDiv[(1!==s[0]||1!==s[1]?"add":"remove")+"Class"]("ui-datepicker-multi"),e.dpDiv[(this._get(e,"isRTL")?"add":"remove")+"Class"]("ui-datepicker-rtl"),e===t.datepicker._curInst&&t.datepicker._datepickerShowing&&t.datepicker._shouldFocusInput(e)&&e.input.trigger("focus"),e.yearshtml&&(i=e.yearshtml,setTimeout(function(){i===e.yearshtml&&e.yearshtml&&e.dpDiv.find("select.ui-datepicker-year").first().replaceWith(e.yearshtml),i=e.yearshtml=null},0)),r&&r.apply(e.input?e.input[0]:null,[e])},_shouldFocusInput:function(t){return t.input&&t.input.is(":visible")&&!t.input.is(":disabled")},_checkOffset:function(e,i,s){var n=e.dpDiv.outerWidth(),o=e.dpDiv.outerHeight(),a=e.input?e.input.outerWidth():0,r=e.input?e.input.outerHeight():0,l=document.documentElement.clientWidth+(s?0:t(document).scrollLeft()),h=document.documentElement.clientHeight+(s?0:t(document).scrollTop());return i.left-=this._get(e,"isRTL")?n-a:0,i.left-=s&&i.left===e.input.offset().left?t(document).scrollLeft():0,i.top-=s&&i.top===e.input.offset().top+r?t(document).scrollTop():0,i.left-=Math.min(i.left,i.left+n>l&&l>n?Math.abs(i.left+n-l):0),i.top-=Math.min(i.top,i.top+o>h&&h>o?Math.abs(o+r):0),i},_findPos:function(e){var i,s=this._getInst(e),n=this._get(s,"isRTL");while(e&&("hidden"===e.type||1!==e.nodeType||t.expr.pseudos.hidden(e)))e=e[n?"previousSibling":"nextSibling"];return[(i=t(e).offset()).left,i.top]},_hideDatepicker:function(e){var i,s,n,o,a=this._curInst;a&&(!e||a===t.data(e,"datepicker"))&&this._datepickerShowing&&(i=this._get(a,"showAnim"),s=this._get(a,"duration"),n=function(){t.datepicker._tidyDialog(a)},t.effects&&t.effects.effect[i]?a.dpDiv.hide(i,t.datepicker._get(a,"showOptions"),s,n):a.dpDiv["slideDown"===i?"slideUp":"fadeIn"===i?"fadeOut":"hide"](i?s:null,n),i||n(),this._datepickerShowing=!1,(o=this._get(a,"onClose"))&&o.apply(a.input?a.input[0]:null,[a.input?a.input.val():"",a]),this._lastInput=null,this._inDialog&&(this._dialogInput.css({position:"absolute",left:"0",top:"-100px"}),t.blockUI&&(t.unblockUI(),t("body").append(this.dpDiv))),this._inDialog=!1)},_tidyDialog:function(t){t.dpDiv.removeClass(this._dialogClass).off(".ui-datepicker-calendar")},_checkExternalClick:function(e){if(t.datepicker._curInst){var i=t(e.target),s=t.datepicker._getInst(i[0]);(i[0].id!==t.datepicker._mainDivId&&0===i.parents("#"+t.datepicker._mainDivId).length&&!i.hasClass(t.datepicker.markerClassName)&&!i.closest("."+t.datepicker._triggerClass).length&&t.datepicker._datepickerShowing&&!(t.datepicker._inDialog&&t.blockUI)||i.hasClass(t.datepicker.markerClassName)&&t.datepicker._curInst!==s)&&t.datepicker._hideDatepicker()}},_adjustDate:function(e,i,s){var n=t(e),o=this._getInst(n[0]);!this._isDisabledDatepicker(n[0])&&(this._adjustInstDate(o,i,s),this._updateDatepicker(o))},_gotoToday:function(e){var i,s=t(e),n=this._getInst(s[0]);this._get(n,"gotoCurrent")&&n.currentDay?(n.selectedDay=n.currentDay,n.drawMonth=n.selectedMonth=n.currentMonth,n.drawYear=n.selectedYear=n.currentYear):(n.selectedDay=(i=new Date).getDate(),n.drawMonth=n.selectedMonth=i.getMonth(),n.drawYear=n.selectedYear=i.getFullYear()),this._notifyChange(n),this._adjustDate(s)},_selectMonthYear:function(e,i,s){var n=t(e),o=this._getInst(n[0]);o["selected"+("M"===s?"Month":"Year")]=o["draw"+("M"===s?"Month":"Year")]=parseInt(i.options[i.selectedIndex].value,10),this._notifyChange(o),this._adjustDate(n)},_selectDay:function(e,i,s,n){var o,a=t(e);!(t(n).hasClass(this._unselectableClass)||this._isDisabledDatepicker(a[0]))&&((o=this._getInst(a[0])).selectedDay=o.currentDay=parseInt(t("a",n).attr("data-date")),o.selectedMonth=o.currentMonth=i,o.selectedYear=o.currentYear=s,this._selectDate(e,this._formatDate(o,o.currentDay,o.currentMonth,o.currentYear)))},_clearDate:function(e){var i=t(e);this._selectDate(i,"")},_selectDate:function(e,i){var s,n=t(e),o=this._getInst(n[0]);i=null!=i?i:this._formatDate(o),o.input&&o.input.val(i),this._updateAlternate(o),(s=this._get(o,"onSelect"))?s.apply(o.input?o.input[0]:null,[i,o]):o.input&&o.input.trigger("change"),o.inline?this._updateDatepicker(o):(this._hideDatepicker(),this._lastInput=o.input[0],"object"!=typeof o.input[0]&&o.input.trigger("focus"),this._lastInput=null)},_updateAlternate:function(e){var i,s,n,o=this._get(e,"altField");o&&(i=this._get(e,"altFormat")||this._get(e,"dateFormat"),s=this._getDate(e),n=this.formatDate(i,s,this._getFormatConfig(e)),t(document).find(o).val(n))},noWeekends:function(t){var e=t.getDay();return[e>0&&e<6,""]},iso8601Week:function(t){var e,i=new Date(t.getTime());return i.setDate(i.getDate()+4-(i.getDay()||7)),e=i.getTime(),i.setMonth(0),i.setDate(1),Math.floor(Math.round((e-i)/864e5)/7)+1},parseDate:function(e,i,s){if(null==e||null==i)throw"Invalid arguments";if(""===(i="object"==typeof i?i.toString():i+""))return null;var n,o,a,r,l=0,h=(s?s.shortYearCutoff:null)||this._defaults.shortYearCutoff,c="string"!=typeof h?h:new Date().getFullYear()%100+parseInt(h,10),u=(s?s.dayNamesShort:null)||this._defaults.dayNamesShort,d=(s?s.dayNames:null)||this._defaults.dayNames,p=(s?s.monthNamesShort:null)||this._defaults.monthNamesShort,f=(s?s.monthNames:null)||this._defaults.monthNames,g=-1,m=-1,_=-1,v=-1,b=!1,y=function(t){var i=n+1<e.length&&e.charAt(n+1)===t;return i&&n++,i},w=function(t){var e=y(t),s="@"===t?14:"!"===t?20:"y"===t&&e?4:"o"===t?3:2,n=RegExp("^\\d{"+("y"===t?s:1)+","+s+"}"),o=i.substring(l).match(n);if(!o)throw"Missing number at position "+l;return l+=o[0].length,parseInt(o[0],10)},x=function(e,s,n){var o=-1,a=t.map(y(e)?n:s,function(t,e){return[[e,t]]}).sort(function(t,e){return-(t[1].length-e[1].length)});if(t.each(a,function(t,e){var s=e[1];if(i.substr(l,s.length).toLowerCase()===s.toLowerCase())return o=e[0],l+=s.length,!1}),-1!==o)return o+1;throw"Unknown name at position "+l},k=function(){if(i.charAt(l)!==e.charAt(n))throw"Unexpected literal at position "+l;l++};for(n=0;n<e.length;n++)if(b)"'"!==e.charAt(n)||y("'")?k():b=!1;else switch(e.charAt(n)){case"d":_=w("d");break;case"D":x("D",u,d);break;case"o":v=w("o");break;case"m":m=w("m");break;case"M":m=x("M",p,f);break;case"y":g=w("y");break;case"@":g=(r=new Date(w("@"))).getFullYear(),m=r.getMonth()+1,_=r.getDate();break;case"!":g=(r=new Date((w("!")-this._ticksTo1970)/1e4)).getFullYear(),m=r.getMonth()+1,_=r.getDate();break;case"'":y("'")?k():b=!0;break;default:k()}if(l<i.length&&(a=i.substr(l),!/^\s+/.test(a)))throw"Extra/unparsed characters found in date: "+a;if(-1===g?g=new Date().getFullYear():g<100&&(g+=new Date().getFullYear()-new Date().getFullYear()%100+(g<=c?0:-100)),v>-1){m=1,_=v;do{if(_<=(o=this._getDaysInMonth(g,m-1)))break;m++,_-=o}while(!0)}if((r=this._daylightSavingAdjust(new Date(g,m-1,_))).getFullYear()!==g||r.getMonth()+1!==m||r.getDate()!==_)throw"Invalid date";return r},ATOM:"yy-mm-dd",COOKIE:"D, dd M yy",ISO_8601:"yy-mm-dd",RFC_822:"D, d M y",RFC_850:"DD, dd-M-y",RFC_1036:"D, d M y",RFC_1123:"D, d M yy",RFC_2822:"D, d M yy",RSS:"D, d M y",TICKS:"!",TIMESTAMP:"@",W3C:"yy-mm-dd",_ticksTo1970:(718685+Math.floor(492.5)-Math.floor(19.7)+Math.floor(4.925))*864e9,formatDate:function(t,e,i){if(!e)return"";var s,n=(i?i.dayNamesShort:null)||this._defaults.dayNamesShort,o=(i?i.dayNames:null)||this._defaults.dayNames,a=(i?i.monthNamesShort:null)||this._defaults.monthNamesShort,r=(i?i.monthNames:null)||this._defaults.monthNames,l=function(e){var i=s+1<t.length&&t.charAt(s+1)===e;return i&&s++,i},h=function(t,e,i){var s=""+e;if(l(t))while(s.length<i)s="0"+s;return s},c=function(t,e,i,s){return l(t)?s[e]:i[e]},u="",d=!1;if(e)for(s=0;s<t.length;s++)if(d)"'"!==t.charAt(s)||l("'")?u+=t.charAt(s):d=!1;else switch(t.charAt(s)){case"d":u+=h("d",e.getDate(),2);break;case"D":u+=c("D",e.getDay(),n,o);break;case"o":u+=h("o",Math.round((new Date(e.getFullYear(),e.getMonth(),e.getDate()).getTime()-new Date(e.getFullYear(),0,0).getTime())/864e5),3);break;case"m":u+=h("m",e.getMonth()+1,2);break;case"M":u+=c("M",e.getMonth(),a,r);break;case"y":u+=l("y")?e.getFullYear():(e.getFullYear()%100<10?"0":"")+e.getFullYear()%100;break;case"@":u+=e.getTime();break;case"!":u+=1e4*e.getTime()+this._ticksTo1970;break;case"'":l("'")?u+="'":d=!0;break;default:u+=t.charAt(s)}return u},_possibleChars:function(t){var e,i="",s=!1,n=function(i){var s=e+1<t.length&&t.charAt(e+1)===i;return s&&e++,s};for(e=0;e<t.length;e++)if(s)"'"!==t.charAt(e)||n("'")?i+=t.charAt(e):s=!1;else switch(t.charAt(e)){case"d":case"m":case"y":case"@":i+="0123456789";break;case"D":case"M":return null;case"'":n("'")?i+="'":s=!0;break;default:i+=t.charAt(e)}return i},_get:function(t,e){return void 0!==t.settings[e]?t.settings[e]:this._defaults[e]},_setDateFromField:function(t,e){if(t.input.val()!==t.lastVal){var i=this._get(t,"dateFormat"),s=t.lastVal=t.input?t.input.val():null,n=this._getDefaultDate(t),o=n,a=this._getFormatConfig(t);try{o=this.parseDate(i,s,a)||n}catch(t){s=e?"":s}t.selectedDay=o.getDate(),t.drawMonth=t.selectedMonth=o.getMonth(),t.drawYear=t.selectedYear=o.getFullYear(),t.currentDay=s?o.getDate():0,t.currentMonth=s?o.getMonth():0,t.currentYear=s?o.getFullYear():0,this._adjustInstDate(t)}},_getDefaultDate:function(t){return this._restrictMinMax(t,this._determineDate(t,this._get(t,"defaultDate"),new Date))},_determineDate:function(e,i,s){var n,o=null==i||""===i?s:"string"==typeof i?function(i){try{return t.datepicker.parseDate(t.datepicker._get(e,"dateFormat"),i,t.datepicker._getFormatConfig(e))}catch(t){}var s=(i.toLowerCase().match(/^c/)?t.datepicker._getDate(e):null)||new Date,n=s.getFullYear(),o=s.getMonth(),a=s.getDate(),r=/([+\-]?[0-9]+)\s*(d|D|w|W|m|M|y|Y)?/g,l=r.exec(i);while(l){switch(l[2]||"d"){case"d":case"D":a+=parseInt(l[1],10);break;case"w":case"W":a+=7*parseInt(l[1],10);break;case"m":case"M":o+=parseInt(l[1],10),a=Math.min(a,t.datepicker._getDaysInMonth(n,o));break;case"y":case"Y":n+=parseInt(l[1],10),a=Math.min(a,t.datepicker._getDaysInMonth(n,o))}l=r.exec(i)}return new Date(n,o,a)}(i):"number"==typeof i?isNaN(i)?s:((n=new Date).setDate(n.getDate()+i),n):new Date(i.getTime());return(o=o&&"Invalid Date"===o.toString()?s:o)&&(o.setHours(0),o.setMinutes(0),o.setSeconds(0),o.setMilliseconds(0)),this._daylightSavingAdjust(o)},_daylightSavingAdjust:function(t){return t?(t.setHours(t.getHours()>12?t.getHours()+2:0),t):null},_setDate:function(t,e,i){var s=!e,n=t.selectedMonth,o=t.selectedYear,a=this._restrictMinMax(t,this._determineDate(t,e,new Date));t.selectedDay=t.currentDay=a.getDate(),t.drawMonth=t.selectedMonth=t.currentMonth=a.getMonth(),t.drawYear=t.selectedYear=t.currentYear=a.getFullYear(),n===t.selectedMonth&&o===t.selectedYear||i||this._notifyChange(t),this._adjustInstDate(t),t.input&&t.input.val(s?"":this._formatDate(t))},_getDate:function(t){return!t.currentYear||t.input&&""===t.input.val()?null:this._daylightSavingAdjust(new Date(t.currentYear,t.currentMonth,t.currentDay))},_attachHandlers:function(e){var i=this._get(e,"stepMonths"),s="#"+e.id.replace(/\\\\/g,"\\");e.dpDiv.find("[data-handler]").map(function(){t(this).on(this.getAttribute("data-event"),{prev:function(){t.datepicker._adjustDate(s,-i,"M")},next:function(){t.datepicker._adjustDate(s,+i,"M")},hide:function(){t.datepicker._hideDatepicker()},today:function(){t.datepicker._gotoToday(s)},selectDay:function(){return t.datepicker._selectDay(s,+this.getAttribute("data-month"),+this.getAttribute("data-year"),this),!1},selectMonth:function(){return t.datepicker._selectMonthYear(s,this,"M"),!1},selectYear:function(){return t.datepicker._selectMonthYear(s,this,"Y"),!1}}[this.getAttribute("data-handler")])})},_generateHTML:function(e){var i,s,n,o,a,r,l,h,c,u,d,p,f,g,m,_,v,b,y,w,x,k,C,D,I,T,P,M,S,H,z,O,A,E,N,W,F,L,R,Y=new Date,B=this._daylightSavingAdjust(new Date(Y.getFullYear(),Y.getMonth(),Y.getDate())),j=this._get(e,"isRTL"),q=this._get(e,"showButtonPanel"),K=this._get(e,"hideIfNoPrevNext"),U=this._get(e,"navigationAsDateFormat"),V=this._getNumberOfMonths(e),X=this._get(e,"showCurrentAtPos"),$=this._get(e,"stepMonths"),G=1!==V[0]||1!==V[1],Q=this._daylightSavingAdjust(e.currentDay?new Date(e.currentYear,e.currentMonth,e.currentDay):new Date(9999,9,9)),J=this._getMinMaxDate(e,"min"),Z=this._getMinMaxDate(e,"max"),tt=e.drawMonth-X,te=e.drawYear;if(tt<0&&(tt+=12,te--),Z){i=this._daylightSavingAdjust(new Date(Z.getFullYear(),Z.getMonth()-V[0]*V[1]+1,Z.getDate())),i=J&&i<J?J:i;while(this._daylightSavingAdjust(new Date(te,tt,1))>i)--tt<0&&(tt=11,te--)}for(e.drawMonth=tt,e.drawYear=te,s=this._get(e,"prevText"),s=U?this.formatDate(s,this._daylightSavingAdjust(new Date(te,tt-$,1)),this._getFormatConfig(e)):s,n=this._canAdjustMonth(e,-1,te,tt)?t("<a>").attr({class:"ui-datepicker-prev ui-corner-all","data-handler":"prev","data-event":"click",title:s}).append(t("<span>").addClass("ui-icon ui-icon-circle-triangle-"+(j?"e":"w")).text(s))[0].outerHTML:K?"":t("<a>").attr({class:"ui-datepicker-prev ui-corner-all ui-state-disabled",title:s}).append(t("<span>").addClass("ui-icon ui-icon-circle-triangle-"+(j?"e":"w")).text(s))[0].outerHTML,o=this._get(e,"nextText"),o=U?this.formatDate(o,this._daylightSavingAdjust(new Date(te,tt+$,1)),this._getFormatConfig(e)):o,a=this._canAdjustMonth(e,1,te,tt)?t("<a>").attr({class:"ui-datepicker-next ui-corner-all","data-handler":"next","data-event":"click",title:o}).append(t("<span>").addClass("ui-icon ui-icon-circle-triangle-"+(j?"w":"e")).text(o))[0].outerHTML:K?"":t("<a>").attr({class:"ui-datepicker-next ui-corner-all ui-state-disabled",title:o}).append(t("<span>").attr("class","ui-icon ui-icon-circle-triangle-"+(j?"w":"e")).text(o))[0].outerHTML,r=this._get(e,"currentText"),l=this._get(e,"gotoCurrent")&&e.currentDay?Q:B,r=U?this.formatDate(r,l,this._getFormatConfig(e)):r,h="",e.inline||(h=t("<button>").attr({type:"button",class:"ui-datepicker-close ui-state-default ui-priority-primary ui-corner-all","data-handler":"hide","data-event":"click"}).text(this._get(e,"closeText"))[0].outerHTML),c="",q&&(c=t("<div class='ui-datepicker-buttonpane ui-widget-content'>").append(j?h:"").append(this._isInRange(e,l)?t("<button>").attr({type:"button",class:"ui-datepicker-current ui-state-default ui-priority-secondary ui-corner-all","data-handler":"today","data-event":"click"}).text(r):"").append(j?"":h)[0].outerHTML),u=isNaN(u=parseInt(this._get(e,"firstDay"),10))?0:u,d=this._get(e,"showWeek"),p=this._get(e,"dayNames"),f=this._get(e,"dayNamesMin"),g=this._get(e,"monthNames"),m=this._get(e,"monthNamesShort"),_=this._get(e,"beforeShowDay"),v=this._get(e,"showOtherMonths"),b=this._get(e,"selectOtherMonths"),y=this._getDefaultDate(e),w="",k=0;k<V[0];k++){for(D=0,C="",this.maxRows=4;D<V[1];D++){if(I=this._daylightSavingAdjust(new Date(te,tt,e.selectedDay)),T=" ui-corner-all",P="",G){if(P+="<div class='ui-datepicker-group",V[1]>1)switch(D){case 0:P+=" ui-datepicker-group-first",T=" ui-corner-"+(j?"right":"left");break;case V[1]-1:P+=" ui-datepicker-group-last",T=" ui-corner-"+(j?"left":"right");break;default:P+=" ui-datepicker-group-middle",T=""}P+="'>"}for(P+="<div class='ui-datepicker-header ui-widget-header ui-helper-clearfix"+T+"'>"+(/all|left/.test(T)&&0===k?j?a:n:"")+(/all|right/.test(T)&&0===k?j?n:a:"")+this._generateMonthYearHeader(e,tt,te,J,Z,k>0||D>0,g,m)+"</div><table class='ui-datepicker-calendar'><thead><tr>",M=d?"<th class='ui-datepicker-week-col'>"+this._get(e,"weekHeader")+"</th>":"",x=0;x<7;x++)S=(x+u)%7,M+="<th scope='col'"+((x+u+6)%7>=5?" class='ui-datepicker-week-end'":"")+"><span title='"+p[S]+"'>"+f[S]+"</span></th>";for(P+=M+"</tr></thead><tbody>",H=this._getDaysInMonth(te,tt),te===e.selectedYear&&tt===e.selectedMonth&&(e.selectedDay=Math.min(e.selectedDay,H)),O=Math.ceil(((z=(this._getFirstDayOfMonth(te,tt)-u+7)%7)+H)/7),A=G&&this.maxRows>O?this.maxRows:O,this.maxRows=A,E=this._daylightSavingAdjust(new Date(te,tt,1-z)),N=0;N<A;N++){for(P+="<tr>",W=d?"<td class='ui-datepicker-week-col'>"+this._get(e,"calculateWeek")(E)+"</td>":"",x=0;x<7;x++)F=_?_.apply(e.input?e.input[0]:null,[E]):[!0,""],R=(L=E.getMonth()!==tt)&&!b||!F[0]||J&&E<J||Z&&E>Z,W+="<td class='"+((x+u+6)%7>=5?" ui-datepicker-week-end":"")+(L?" ui-datepicker-other-month":"")+(E.getTime()===I.getTime()&&tt===e.selectedMonth&&e._keyEvent||y.getTime()===E.getTime()&&y.getTime()===I.getTime()?" "+this._dayOverClass:"")+(R?" "+this._unselectableClass+" ui-state-disabled":"")+(L&&!v?"":" "+F[1]+(E.getTime()===Q.getTime()?" "+this._currentClass:"")+(E.getTime()===B.getTime()?" ui-datepicker-today":""))+"'"+((!L||v)&&F[2]?" title='"+F[2].replace(/'/g,"&#39;")+"'":"")+(R?"":" data-handler='selectDay' data-event='click' data-month='"+E.getMonth()+"' data-year='"+E.getFullYear()+"'")+">"+(L&&!v?"&#xa0;":R?"<span class='ui-state-default'>"+E.getDate()+"</span>":"<a class='ui-state-default"+(E.getTime()===B.getTime()?" ui-state-highlight":"")+(E.getTime()===Q.getTime()?" ui-state-active":"")+(L?" ui-priority-secondary":"")+"' href='#' aria-current='"+(E.getTime()===Q.getTime()?"true":"false")+"' data-date='"+E.getDate()+"'>"+E.getDate()+"</a>")+"</td>",E.setDate(E.getDate()+1),E=this._daylightSavingAdjust(E);P+=W+"</tr>"}++tt>11&&(tt=0,te++),P+="</tbody></table>"+(G?"</div>"+(V[0]>0&&D===V[1]-1?"<div class='ui-datepicker-row-break'></div>":""):""),C+=P}w+=C}return w+=c,e._keyEvent=!1,w},_generateMonthYearHeader:function(t,e,i,s,n,o,a,r){var l,h,c,u,d,p,f,g,m=this._get(t,"changeMonth"),_=this._get(t,"changeYear"),v=this._get(t,"showMonthAfterYear"),b=this._get(t,"selectMonthLabel"),y=this._get(t,"selectYearLabel"),w="<div class='ui-datepicker-title'>",x="";if(o||!m)x+="<span class='ui-datepicker-month'>"+a[e]+"</span>";else{for(l=s&&s.getFullYear()===i,h=n&&n.getFullYear()===i,x+="<select class='ui-datepicker-month' aria-label='"+b+"' data-handler='selectMonth' data-event='change'>",c=0;c<12;c++)(!l||c>=s.getMonth())&&(!h||c<=n.getMonth())&&(x+="<option value='"+c+"'"+(c===e?" selected='selected'":"")+">"+r[c]+"</option>");x+="</select>"}if(v||(w+=x+(o||!(m&&_)?"&#xa0;":"")),!t.yearshtml){if(t.yearshtml="",o||!_)w+="<span class='ui-datepicker-year'>"+i+"</span>";else{for(u=this._get(t,"yearRange").split(":"),d=new Date().getFullYear(),g=Math.max(f=(p=function(t){var e=t.match(/c[+\-].*/)?i+parseInt(t.substring(1),10):t.match(/[+\-].*/)?d+parseInt(t,10):parseInt(t,10);return isNaN(e)?d:e})(u[0]),p(u[1]||"")),f=s?Math.max(f,s.getFullYear()):f,g=n?Math.min(g,n.getFullYear()):g,t.yearshtml+="<select class='ui-datepicker-year' aria-label='"+y+"' data-handler='selectYear' data-event='change'>";f<=g;f++)t.yearshtml+="<option value='"+f+"'"+(f===i?" selected='selected'":"")+">"+f+"</option>";t.yearshtml+="</select>",w+=t.yearshtml,t.yearshtml=null}}return w+=this._get(t,"yearSuffix"),v&&(w+=(o||!(m&&_)?"&#xa0;":"")+x),w+="</div>"},_adjustInstDate:function(t,e,i){var s=t.selectedYear+("Y"===i?e:0),n=t.selectedMonth+("M"===i?e:0),o=Math.min(t.selectedDay,this._getDaysInMonth(s,n))+("D"===i?e:0),a=this._restrictMinMax(t,this._daylightSavingAdjust(new Date(s,n,o)));t.selectedDay=a.getDate(),t.drawMonth=t.selectedMonth=a.getMonth(),t.drawYear=t.selectedYear=a.getFullYear(),("M"===i||"Y"===i)&&this._notifyChange(t)},_restrictMinMax:function(t,e){var i=this._getMinMaxDate(t,"min"),s=this._getMinMaxDate(t,"max"),n=i&&e<i?i:e;return s&&n>s?s:n},_notifyChange:function(t){var e=this._get(t,"onChangeMonthYear");e&&e.apply(t.input?t.input[0]:null,[t.selectedYear,t.selectedMonth+1,t])},_getNumberOfMonths:function(t){var e=this._get(t,"numberOfMonths");return null==e?[1,1]:"number"==typeof e?[1,e]:e},_getMinMaxDate:function(t,e){return this._determineDate(t,this._get(t,e+"Date"),null)},_getDaysInMonth:function(t,e){return 32-this._daylightSavingAdjust(new Date(t,e,32)).getDate()},_getFirstDayOfMonth:function(t,e){return new Date(t,e,1).getDay()},_canAdjustMonth:function(t,e,i,s){var n=this._getNumberOfMonths(t),o=this._daylightSavingAdjust(new Date(i,s+(e<0?e:n[0]*n[1]),1));return e<0&&o.setDate(this._getDaysInMonth(o.getFullYear(),o.getMonth())),this._isInRange(t,o)},_isInRange:function(t,e){var i,s,n=this._getMinMaxDate(t,"min"),o=this._getMinMaxDate(t,"max"),a=null,r=null,l=this._get(t,"yearRange");return l&&(i=l.split(":"),s=new Date().getFullYear(),a=parseInt(i[0],10),r=parseInt(i[1],10),i[0].match(/[+\-].*/)&&(a+=s),i[1].match(/[+\-].*/)&&(r+=s)),(!n||e.getTime()>=n.getTime())&&(!o||e.getTime()<=o.getTime())&&(!a||e.getFullYear()>=a)&&(!r||e.getFullYear()<=r)},_getFormatConfig:function(t){var e=this._get(t,"shortYearCutoff");return{shortYearCutoff:e="string"!=typeof e?e:new Date().getFullYear()%100+parseInt(e,10),dayNamesShort:this._get(t,"dayNamesShort"),dayNames:this._get(t,"dayNames"),monthNamesShort:this._get(t,"monthNamesShort"),monthNames:this._get(t,"monthNames")}},_formatDate:function(t,e,i,s){e||(t.currentDay=t.selectedDay,t.currentMonth=t.selectedMonth,t.currentYear=t.selectedYear);var n=e?"object"==typeof e?e:this._daylightSavingAdjust(new Date(s,i,e)):this._daylightSavingAdjust(new Date(t.currentYear,t.currentMonth,t.currentDay));return this.formatDate(this._get(t,"dateFormat"),n,this._getFormatConfig(t))}}),t.fn.datepicker=function(e){if(!this.length)return this;t.datepicker.initialized||(t(document).on("mousedown",t.datepicker._checkExternalClick),t.datepicker.initialized=!0),0===t("#"+t.datepicker._mainDivId).length&&t("body").append(t.datepicker.dpDiv);var i=Array.prototype.slice.call(arguments,1);return"string"==typeof e&&("isDisabled"===e||"getDate"===e||"widget"===e)||"option"===e&&2==arguments.length&&"string"==typeof arguments[1]?t.datepicker["_"+e+"Datepicker"].apply(t.datepicker,[this[0]].concat(i)):this.each(function(){"string"==typeof e?t.datepicker["_"+e+"Datepicker"].apply(t.datepicker,[this].concat(i)):t.datepicker._attachDatepicker(this,e)})},t.datepicker=new P,t.datepicker.initialized=!1,t.datepicker.uuid=new Date().getTime(),t.datepicker.version="1.14.2",t.datepicker;var z=!1;function O(t){return function(){var e=this.element.val();t.apply(this,arguments),this._refresh(),e!==this.element.val()&&this._trigger("change")}}t(document).on("mouseup",function(){z=!1}),t.widget("ui.mouse",{version:"1.14.2",options:{cancel:"input, textarea, button, select, option",distance:1,delay:0},_mouseInit:function(){var e=this;this.element.on("mousedown."+this.widgetName,function(t){return e._mouseDown(t)}).on("click."+this.widgetName,function(i){if(!0===t.data(i.target,e.widgetName+".preventClickEvent"))return t.removeData(i.target,e.widgetName+".preventClickEvent"),i.stopImmediatePropagation(),!1}),this.started=!1},_mouseDestroy:function(){this.element.off("."+this.widgetName),this._mouseMoveDelegate&&this.document.off("mousemove."+this.widgetName,this._mouseMoveDelegate).off("mouseup."+this.widgetName,this._mouseUpDelegate)},_mouseDown:function(e){if(!z){this._mouseMoved=!1,this._mouseStarted&&this._mouseUp(e),this._mouseDownEvent=e;var i=this,s=1===e.which,n="string"==typeof this.options.cancel&&t(e.target).closest(this.options.cancel).length;return!(s&&!n&&this._mouseCapture(e))||((this.mouseDelayMet=!this.options.delay,this.mouseDelayMet||(this._mouseDelayTimer=setTimeout(function(){i.mouseDelayMet=!0},this.options.delay)),this._mouseDistanceMet(e)&&this._mouseDelayMet(e)&&(this._mouseStarted=!1!==this._mouseStart(e),!this._mouseStarted))?e.preventDefault():(!0===t.data(e.target,this.widgetName+".preventClickEvent")&&t.removeData(e.target,this.widgetName+".preventClickEvent"),this._mouseMoveDelegate=function(t){return i._mouseMove(t)},this._mouseUpDelegate=function(t){return i._mouseUp(t)},this.document.on("mousemove."+this.widgetName,this._mouseMoveDelegate).on("mouseup."+this.widgetName,this._mouseUpDelegate),e.preventDefault(),z=!0),!0)}},_mouseMove:function(t){if(this._mouseMoved&&!t.which){if(t.originalEvent.altKey||t.originalEvent.ctrlKey||t.originalEvent.metaKey||t.originalEvent.shiftKey)this.ignoreMissingWhich=!0;else if(!this.ignoreMissingWhich)return this._mouseUp(t)}return((t.which||t.button)&&(this._mouseMoved=!0),this._mouseStarted)?(this._mouseDrag(t),t.preventDefault()):(this._mouseDistanceMet(t)&&this._mouseDelayMet(t)&&(this._mouseStarted=!1!==this._mouseStart(this._mouseDownEvent,t),this._mouseStarted?this._mouseDrag(t):this._mouseUp(t)),!this._mouseStarted)},_mouseUp:function(e){this.document.off("mousemove."+this.widgetName,this._mouseMoveDelegate).off("mouseup."+this.widgetName,this._mouseUpDelegate),this._mouseStarted&&(this._mouseStarted=!1,e.target===this._mouseDownEvent.target&&t.data(e.target,this.widgetName+".preventClickEvent",!0),this._mouseStop(e)),this._mouseDelayTimer&&(clearTimeout(this._mouseDelayTimer),delete this._mouseDelayTimer),this.ignoreMissingWhich=!1,z=!1,e.preventDefault()},_mouseDistanceMet:function(t){return Math.max(Math.abs(this._mouseDownEvent.pageX-t.pageX),Math.abs(this._mouseDownEvent.pageY-t.pageY))>=this.options.distance},_mouseDelayMet:function(){return this.mouseDelayMet},_mouseStart:function(){},_mouseDrag:function(){},_mouseStop:function(){},_mouseCapture:function(){return!0}}),t.ui.plugin={add:function(e,i,s){var n,o=t.ui[e].prototype;for(n in s)o.plugins[n]=o.plugins[n]||[],o.plugins[n].push([i,s[n]])},call:function(t,e,i,s){var n,o=t.plugins[e];if(o&&(s||t.element[0].parentNode&&11!==t.element[0].parentNode.nodeType))for(n=0;n<o.length;n++)t.options[o[n][0]]&&o[n][1].apply(t.element,i)}},t.widget("ui.draggable",t.ui.mouse,{version:"1.14.2",widgetEventPrefix:"drag",options:{addClasses:!0,appendTo:"parent",axis:!1,connectToSortable:!1,containment:!1,cursor:"auto",cursorAt:!1,grid:!1,handle:!1,helper:"original",iframeFix:!1,opacity:!1,refreshPositions:!1,revert:!1,revertDuration:500,scope:"default",scroll:!0,scrollSensitivity:20,scrollSpeed:20,snap:!1,snapMode:"both",snapTolerance:20,stack:!1,zIndex:!1,drag:null,start:null,stop:null},_create:function(){"original"===this.options.helper&&this._setPositionRelative(),this.options.addClasses&&this._addClass("ui-draggable"),this._setHandleClassName(),this._mouseInit()},_setOption:function(t,e){this._super(t,e),"handle"===t&&(this._removeHandleClassName(),this._setHandleClassName())},_destroy:function(){if((this.helper||this.element).is(".ui-draggable-dragging")){this.destroyOnClear=!0;return}this._removeHandleClassName(),this._mouseDestroy()},_mouseCapture:function(e){var i=this.options;return!this.helper&&!i.disabled&&!(t(e.target).closest(".ui-resizable-handle").length>0)&&(this.handle=this._getHandle(e),!!this.handle&&(this._blurActiveElement(e),this._blockFrames(!0===i.iframeFix?"iframe":i.iframeFix),!0))},_blockFrames:function(e){this.iframeBlocks=this.document.find(e).map(function(){var e=t(this);return t("<div>").css("position","absolute").appendTo(e.parent()).outerWidth(e.outerWidth()).outerHeight(e.outerHeight()).offset(e.offset())[0]})},_unblockFrames:function(){this.iframeBlocks&&(this.iframeBlocks.remove(),delete this.iframeBlocks)},_blurActiveElement:function(e){var i=this.document[0].activeElement;!t(e.target).closest(i).length&&t(i).trigger("blur")},_mouseStart:function(e){var i=this.options;return(this.helper=this._createHelper(e),this._addClass(this.helper,"ui-draggable-dragging"),this._cacheHelperProportions(),t.ui.ddmanager&&(t.ui.ddmanager.current=this),this._cacheMargins(),this.cssPosition=this.helper.css("position"),this.scrollParent=this.helper.scrollParent(!0),this.offsetParent=this.helper.offsetParent(),this.hasFixedAncestor=this.helper.parents().filter(function(){return"fixed"===t(this).css("position")}).length>0,this.positionAbs=this.element.offset(),this._refreshOffsets(e),this.originalPosition=this.position=this._generatePosition(e,!1),this.originalPageX=e.pageX,this.originalPageY=e.pageY,i.cursorAt&&this._adjustOffsetFromHelper(i.cursorAt),this._setContainment(),!1===this._trigger("start",e))?(this._clear(),!1):(this._cacheHelperProportions(),t.ui.ddmanager&&!i.dropBehaviour&&t.ui.ddmanager.prepareOffsets(this,e),this._mouseDrag(e,!0),t.ui.ddmanager&&t.ui.ddmanager.dragStart(this,e),!0)},_refreshOffsets:function(t){this.offset={top:this.positionAbs.top-this.margins.top,left:this.positionAbs.left-this.margins.left,scroll:!1,parent:this._getParentOffset(),relative:this._getRelativeOffset()},this.offset.click={left:t.pageX-this.offset.left,top:t.pageY-this.offset.top}},_mouseDrag:function(e,i){if(this.hasFixedAncestor&&(this.offset.parent=this._getParentOffset()),this.position=this._generatePosition(e,!0),this.positionAbs=this._convertPositionTo("absolute"),!i){var s=this._uiHash();if(!1===this._trigger("drag",e,s))return this._mouseUp(new t.Event("mouseup",e)),!1;this.position=s.position}return this.helper[0].style.left=this.position.left+"px",this.helper[0].style.top=this.position.top+"px",t.ui.ddmanager&&t.ui.ddmanager.drag(this,e),!1},_mouseStop:function(e){var i=this,s=!1;return t.ui.ddmanager&&!this.options.dropBehaviour&&(s=t.ui.ddmanager.drop(this,e)),this.dropped&&(s=this.dropped,this.dropped=!1),"invalid"===this.options.revert&&!s||"valid"===this.options.revert&&s||!0===this.options.revert||"function"==typeof this.options.revert&&this.options.revert.call(this.element,s)?t(this.helper).animate(this.originalPosition,parseInt(this.options.revertDuration,10),function(){!1!==i._trigger("stop",e)&&i._clear()}):!1!==this._trigger("stop",e)&&this._clear(),!1},_mouseUp:function(e){return this._unblockFrames(),t.ui.ddmanager&&t.ui.ddmanager.dragStop(this,e),this.handleElement.is(e.target)&&this.element.trigger("focus"),t.ui.mouse.prototype._mouseUp.call(this,e)},cancel:function(){return this.helper.is(".ui-draggable-dragging")?this._mouseUp(new t.Event("mouseup",{target:this.element[0]})):this._clear(),this},_getHandle:function(e){return!this.options.handle||!!t(e.target).closest(this.element.find(this.options.handle)).length},_setHandleClassName:function(){this.handleElement=this.options.handle?this.element.find(this.options.handle):this.element,this._addClass(this.handleElement,"ui-draggable-handle")},_removeHandleClassName:function(){this._removeClass(this.handleElement,"ui-draggable-handle")},_createHelper:function(e){var i=this.options,s="function"==typeof i.helper,n=s?t(i.helper.apply(this.element[0],[e])):"clone"===i.helper?this.element.clone().removeAttr("id"):this.element;return n.parents("body").length||n.appendTo("parent"===i.appendTo?this.element[0].parentNode:i.appendTo),s&&n[0]===this.element[0]&&this._setPositionRelative(),n[0]===this.element[0]||/(fixed|absolute)/.test(n.css("position"))||n.css("position","absolute"),n},_setPositionRelative:function(){/^(?:r|a|f)/.test(this.element.css("position"))||(this.element[0].style.position="relative")},_adjustOffsetFromHelper:function(t){"string"==typeof t&&(t=t.split(" ")),Array.isArray(t)&&(t={left:+t[0],top:+t[1]||0}),"left"in t&&(this.offset.click.left=t.left+this.margins.left),"right"in t&&(this.offset.click.left=this.helperProportions.width-t.right+this.margins.left),"top"in t&&(this.offset.click.top=t.top+this.margins.top),"bottom"in t&&(this.offset.click.top=this.helperProportions.height-t.bottom+this.margins.top)},_isRootNode:function(t){return/(html|body)/i.test(t.tagName)||t===this.document[0]},_getParentOffset:function(){var e=this.offsetParent.offset(),i=this.document[0];return"absolute"===this.cssPosition&&this.scrollParent[0]!==i&&t.contains(this.scrollParent[0],this.offsetParent[0])&&(e.left+=this.scrollParent.scrollLeft(),e.top+=this.scrollParent.scrollTop()),this._isRootNode(this.offsetParent[0])&&(e={top:0,left:0}),{top:e.top+(parseInt(this.offsetParent.css("borderTopWidth"),10)||0),left:e.left+(parseInt(this.offsetParent.css("borderLeftWidth"),10)||0)}},_getRelativeOffset:function(){if("relative"!==this.cssPosition)return{top:0,left:0};var t=this.element.position(),e=this._isRootNode(this.scrollParent[0]);return{top:t.top-(parseInt(this.helper.css("top"),10)||0)+(e?0:this.scrollParent.scrollTop()),left:t.left-(parseInt(this.helper.css("left"),10)||0)+(e?0:this.scrollParent.scrollLeft())}},_cacheMargins:function(){this.margins={left:parseInt(this.element.css("marginLeft"),10)||0,top:parseInt(this.element.css("marginTop"),10)||0,right:parseInt(this.element.css("marginRight"),10)||0,bottom:parseInt(this.element.css("marginBottom"),10)||0}},_cacheHelperProportions:function(){this.helperProportions={width:this.helper.outerWidth(),height:this.helper.outerHeight()}},_setContainment:function(){var e,i,s,n=this.options,o=this.document[0];if(this.relativeContainer=null,!n.containment){this.containment=null;return}if("window"===n.containment){this.containment=[t(window).scrollLeft()-this.offset.relative.left-this.offset.parent.left,t(window).scrollTop()-this.offset.relative.top-this.offset.parent.top,t(window).scrollLeft()+t(window).width()-this.helperProportions.width-this.margins.left,t(window).scrollTop()+(t(window).height()||o.body.parentNode.scrollHeight)-this.helperProportions.height-this.margins.top];return}if("document"===n.containment){this.containment=[0,0,t(o).width()-this.helperProportions.width-this.margins.left,(t(o).height()||o.body.parentNode.scrollHeight)-this.helperProportions.height-this.margins.top];return}if(n.containment.constructor===Array){this.containment=n.containment;return}"parent"===n.containment&&(n.containment=this.helper[0].parentNode),(s=(i=t(n.containment))[0])&&(e=/(scroll|auto)/.test(i.css("overflow")),this.containment=[(parseInt(i.css("borderLeftWidth"),10)||0)+(parseInt(i.css("paddingLeft"),10)||0),(parseInt(i.css("borderTopWidth"),10)||0)+(parseInt(i.css("paddingTop"),10)||0),(e?Math.max(s.scrollWidth,s.offsetWidth):s.offsetWidth)-(parseInt(i.css("borderRightWidth"),10)||0)-(parseInt(i.css("paddingRight"),10)||0)-this.helperProportions.width-this.margins.left-this.margins.right,(e?Math.max(s.scrollHeight,s.offsetHeight):s.offsetHeight)-(parseInt(i.css("borderBottomWidth"),10)||0)-(parseInt(i.css("paddingBottom"),10)||0)-this.helperProportions.height-this.margins.top-this.margins.bottom],this.relativeContainer=i)},_convertPositionTo:function(t,e){e||(e=this.position);var i="absolute"===t?1:-1,s=this._isRootNode(this.scrollParent[0]);return{top:e.top+this.offset.relative.top*i+this.offset.parent.top*i-("fixed"===this.cssPosition?-this.offset.scroll.top:s?0:this.offset.scroll.top)*i,left:e.left+this.offset.relative.left*i+this.offset.parent.left*i-("fixed"===this.cssPosition?-this.offset.scroll.left:s?0:this.offset.scroll.left)*i}},_generatePosition:function(t,e){var i,s,n,o,a=this.options,r=this._isRootNode(this.scrollParent[0]),l=t.pageX,h=t.pageY;return r&&this.offset.scroll||(this.offset.scroll={top:this.scrollParent.scrollTop(),left:this.scrollParent.scrollLeft()}),e&&(this.containment&&(this.relativeContainer?(s=this.relativeContainer.offset(),i=[this.containment[0]+s.left,this.containment[1]+s.top,this.containment[2]+s.left,this.containment[3]+s.top]):i=this.containment,t.pageX-this.offset.click.left<i[0]&&(l=i[0]+this.offset.click.left),t.pageY-this.offset.click.top<i[1]&&(h=i[1]+this.offset.click.top),t.pageX-this.offset.click.left>i[2]&&(l=i[2]+this.offset.click.left),t.pageY-this.offset.click.top>i[3]&&(h=i[3]+this.offset.click.top)),a.grid&&(n=a.grid[1]?this.originalPageY+Math.round((h-this.originalPageY)/a.grid[1])*a.grid[1]:this.originalPageY,h=i?n-this.offset.click.top>=i[1]||n-this.offset.click.top>i[3]?n:n-this.offset.click.top>=i[1]?n-a.grid[1]:n+a.grid[1]:n,o=a.grid[0]?this.originalPageX+Math.round((l-this.originalPageX)/a.grid[0])*a.grid[0]:this.originalPageX,l=i?o-this.offset.click.left>=i[0]||o-this.offset.click.left>i[2]?o:o-this.offset.click.left>=i[0]?o-a.grid[0]:o+a.grid[0]:o),"y"===a.axis&&(l=this.originalPageX),"x"===a.axis&&(h=this.originalPageY)),{top:h-this.offset.click.top-this.offset.relative.top-this.offset.parent.top+("fixed"===this.cssPosition?-this.offset.scroll.top:r?0:this.offset.scroll.top),left:l-this.offset.click.left-this.offset.relative.left-this.offset.parent.left+("fixed"===this.cssPosition?-this.offset.scroll.left:r?0:this.offset.scroll.left)}},_clear:function(){this._removeClass(this.helper,"ui-draggable-dragging"),this.helper[0]===this.element[0]||this.cancelHelperRemoval||this.helper.remove(),this.helper=null,this.cancelHelperRemoval=!1,this.destroyOnClear&&this.destroy()},_trigger:function(e,i,s){return s=s||this._uiHash(),t.ui.plugin.call(this,e,[i,s,this],!0),/^(drag|start|stop)/.test(e)&&(this.positionAbs=this._convertPositionTo("absolute"),s.offset=this.positionAbs),t.Widget.prototype._trigger.call(this,e,i,s)},plugins:{},_uiHash:function(){return{helper:this.helper,position:this.position,originalPosition:this.originalPosition,offset:this.positionAbs}}}),t.ui.plugin.add("draggable","connectToSortable",{start:function(e,i,s){var n=t.extend({},i,{item:s.element});s.sortables=[],t(s.options.connectToSortable).each(function(){var i=t(this).sortable("instance");i&&!i.options.disabled&&(s.sortables.push(i),i.refreshPositions(),i._trigger("activate",e,n))})},stop:function(e,i,s){var n=t.extend({},i,{item:s.element});s.cancelHelperRemoval=!1,t.each(s.sortables,function(){this.isOver?(this.isOver=0,s.cancelHelperRemoval=!0,this.cancelHelperRemoval=!1,this._storedCSS={position:this.placeholder.css("position"),top:this.placeholder.css("top"),left:this.placeholder.css("left")},this._mouseStop(e),this.options.helper=this.options._helper):(this.cancelHelperRemoval=!0,this._trigger("deactivate",e,n))})},drag:function(e,i,s){t.each(s.sortables,function(){var n=!1,o=this;o.positionAbs=s.positionAbs,o.helperProportions=s.helperProportions,o.offset.click=s.offset.click,o._intersectsWith(o.containerCache)&&(n=!0,t.each(s.sortables,function(){return this.positionAbs=s.positionAbs,this.helperProportions=s.helperProportions,this.offset.click=s.offset.click,this!==o&&this._intersectsWith(this.containerCache)&&t.contains(o.element[0],this.element[0])&&(n=!1),n})),n?(o.isOver||(o.isOver=1,s._parent=i.helper.parent(),o.currentItem=i.helper.appendTo(o.element).data("ui-sortable-item",!0),o.options._helper=o.options.helper,o.options.helper=function(){return i.helper[0]},e.target=o.currentItem[0],o._mouseCapture(e,!0),o._mouseStart(e,!0,!0),o.offset.click.top=s.offset.click.top,o.offset.click.left=s.offset.click.left,o.offset.parent.left-=s.offset.parent.left-o.offset.parent.left,o.offset.parent.top-=s.offset.parent.top-o.offset.parent.top,s._trigger("toSortable",e),s.dropped=o.element,t.each(s.sortables,function(){this.refreshPositions()}),s.currentItem=s.element,o.fromOutside=s),o.currentItem&&(o._mouseDrag(e),i.position=o.position)):o.isOver&&(o.isOver=0,o.cancelHelperRemoval=!0,o.options._revert=o.options.revert,o.options.revert=!1,o._trigger("out",e,o._uiHash(o)),o._mouseStop(e,!0),o.options.revert=o.options._revert,o.options.helper=o.options._helper,o.placeholder&&o.placeholder.remove(),i.helper.appendTo(s._parent),s._refreshOffsets(e),i.position=s._generatePosition(e,!0),s._trigger("fromSortable",e),s.dropped=!1,t.each(s.sortables,function(){this.refreshPositions()}))})}}),t.ui.plugin.add("draggable","cursor",{start:function(e,i,s){var n=t("body"),o=s.options;n.css("cursor")&&(o._cursor=n.css("cursor")),n.css("cursor",o.cursor)},stop:function(e,i,s){var n=s.options;n._cursor&&t("body").css("cursor",n._cursor)}}),t.ui.plugin.add("draggable","opacity",{start:function(e,i,s){var n=t(i.helper),o=s.options;n.css("opacity")&&(o._opacity=n.css("opacity")),n.css("opacity",o.opacity)},stop:function(e,i,s){var n=s.options;n._opacity&&t(i.helper).css("opacity",n._opacity)}}),t.ui.plugin.add("draggable","scroll",{start:function(t,e,i){i.scrollParentNotHidden||(i.scrollParentNotHidden=i.helper.scrollParent(!1)),i.scrollParentNotHidden[0]!==i.document[0]&&"HTML"!==i.scrollParentNotHidden[0].tagName&&(i.overflowOffset=i.scrollParentNotHidden.offset())},drag:function(e,i,s){var n=s.options,o=!1,a=s.scrollParentNotHidden[0],r=s.document[0];a!==r&&"HTML"!==a.tagName?((!n.axis||"x"!==n.axis)&&(s.overflowOffset.top+a.offsetHeight-e.pageY<n.scrollSensitivity?a.scrollTop=o=a.scrollTop+n.scrollSpeed:e.pageY-s.overflowOffset.top<n.scrollSensitivity&&(a.scrollTop=o=a.scrollTop-n.scrollSpeed)),(!n.axis||"y"!==n.axis)&&(s.overflowOffset.left+a.offsetWidth-e.pageX<n.scrollSensitivity?a.scrollLeft=o=a.scrollLeft+n.scrollSpeed:e.pageX-s.overflowOffset.left<n.scrollSensitivity&&(a.scrollLeft=o=a.scrollLeft-n.scrollSpeed))):((!n.axis||"x"!==n.axis)&&(e.pageY-t(r).scrollTop()<n.scrollSensitivity?o=t(r).scrollTop(t(r).scrollTop()-n.scrollSpeed):t(window).height()-(e.pageY-t(r).scrollTop())<n.scrollSensitivity&&(o=t(r).scrollTop(t(r).scrollTop()+n.scrollSpeed))),(!n.axis||"y"!==n.axis)&&(e.pageX-t(r).scrollLeft()<n.scrollSensitivity?o=t(r).scrollLeft(t(r).scrollLeft()-n.scrollSpeed):t(window).width()-(e.pageX-t(r).scrollLeft())<n.scrollSensitivity&&(o=t(r).scrollLeft(t(r).scrollLeft()+n.scrollSpeed)))),!1!==o&&t.ui.ddmanager&&!n.dropBehaviour&&t.ui.ddmanager.prepareOffsets(s,e)}}),t.ui.plugin.add("draggable","snap",{start:function(e,i,s){var n=s.options;s.snapElements=[],t(n.snap.constructor!==String?n.snap.items||":data(ui-draggable)":n.snap).each(function(){var e=t(this),i=e.offset();this!==s.element[0]&&s.snapElements.push({item:this,width:e.outerWidth(),height:e.outerHeight(),top:i.top,left:i.left})})},drag:function(e,i,s){var n,o,a,r,l,h,c,u,d,p,f=s.options,g=f.snapTolerance,m=i.offset.left,_=m+s.helperProportions.width,v=i.offset.top,b=v+s.helperProportions.height;for(d=s.snapElements.length-1;d>=0;d--){if(h=(l=s.snapElements[d].left-s.margins.left)+s.snapElements[d].width,u=(c=s.snapElements[d].top-s.margins.top)+s.snapElements[d].height,_<l-g||m>h+g||b<c-g||v>u+g||!t.contains(s.snapElements[d].item.ownerDocument,s.snapElements[d].item)){s.snapElements[d].snapping&&s.options.snap.release&&s.options.snap.release.call(s.element,e,t.extend(s._uiHash(),{snapItem:s.snapElements[d].item})),s.snapElements[d].snapping=!1;continue}"inner"!==f.snapMode&&(n=Math.abs(c-b)<=g,o=Math.abs(u-v)<=g,a=Math.abs(l-_)<=g,r=Math.abs(h-m)<=g,n&&(i.position.top=s._convertPositionTo("relative",{top:c-s.helperProportions.height,left:0}).top),o&&(i.position.top=s._convertPositionTo("relative",{top:u,left:0}).top),a&&(i.position.left=s._convertPositionTo("relative",{top:0,left:l-s.helperProportions.width}).left),r&&(i.position.left=s._convertPositionTo("relative",{top:0,left:h}).left)),p=n||o||a||r,"outer"!==f.snapMode&&(n=Math.abs(c-v)<=g,o=Math.abs(u-b)<=g,a=Math.abs(l-m)<=g,r=Math.abs(h-_)<=g,n&&(i.position.top=s._convertPositionTo("relative",{top:c,left:0}).top),o&&(i.position.top=s._convertPositionTo("relative",{top:u-s.helperProportions.height,left:0}).top),a&&(i.position.left=s._convertPositionTo("relative",{top:0,left:l}).left),r&&(i.position.left=s._convertPositionTo("relative",{top:0,left:h-s.helperProportions.width}).left)),!s.snapElements[d].snapping&&(n||o||a||r||p)&&s.options.snap.snap&&s.options.snap.snap.call(s.element,e,t.extend(s._uiHash(),{snapItem:s.snapElements[d].item})),s.snapElements[d].snapping=n||o||a||r||p}}}),t.ui.plugin.add("draggable","stack",{start:function(e,i,s){var n,o=s.options,a=t.makeArray(t(o.stack)).sort(function(e,i){return(parseInt(t(e).css("zIndex"),10)||0)-(parseInt(t(i).css("zIndex"),10)||0)});a.length&&(n=parseInt(t(a[0]).css("zIndex"),10)||0,t(a).each(function(e){t(this).css("zIndex",n+e)}),this.css("zIndex",n+a.length))}}),t.ui.plugin.add("draggable","zIndex",{start:function(e,i,s){var n=t(i.helper),o=s.options;n.css("zIndex")&&(o._zIndex=n.css("zIndex")),n.css("zIndex",o.zIndex)},stop:function(e,i,s){var n=s.options;n._zIndex&&t(i.helper).css("zIndex",n._zIndex)}}),t.ui.draggable,t.widget("ui.resizable",t.ui.mouse,{version:"1.14.2",widgetEventPrefix:"resize",options:{alsoResize:!1,animate:!1,animateDuration:"slow",animateEasing:"swing",aspectRatio:!1,autoHide:!1,classes:{"ui-resizable-se":"ui-icon ui-icon-gripsmall-diagonal-se"},containment:!1,ghost:!1,grid:!1,handles:"e,s,se",helper:!1,maxHeight:null,maxWidth:null,minHeight:10,minWidth:10,zIndex:90,resize:null,start:null,stop:null},_num:function(t){return parseFloat(t)||0},_isNumber:function(t){return!isNaN(parseFloat(t))},_hasScroll:function(e,i){var s,n=!1,o=t(e).css("overflow");if("hidden"===o)return!1;if("scroll"===o||e[s=i&&"left"===i?"scrollLeft":"scrollTop"]>0)return!0;try{e[s]=1,n=e[s]>0,e[s]=0}catch(t){}return n},_create:function(){var e,i=this.options,s=this;this._addClass("ui-resizable"),t.extend(this,{_aspectRatio:!!i.aspectRatio,aspectRatio:i.aspectRatio,originalElement:this.element,_proportionallyResizeElements:[],_helper:i.helper||i.ghost||i.animate?i.helper||"ui-resizable-helper":null}),this.element[0].nodeName.match(/^(canvas|textarea|input|select|button|img)$/i)&&(this.element.wrap(t("<div class='ui-wrapper'></div>").css({overflow:"hidden",position:this.element.css("position"),width:this.element.outerWidth(),height:this.element.outerHeight(),top:this.element.css("top"),left:this.element.css("left")})),this.element=this.element.parent().data("ui-resizable",this.element.resizable("instance")),this.elementIsWrapper=!0,e={marginTop:this.originalElement.css("marginTop"),marginRight:this.originalElement.css("marginRight"),marginBottom:this.originalElement.css("marginBottom"),marginLeft:this.originalElement.css("marginLeft")},this.element.css(e),this.originalResizeStyle=this.originalElement.css("resize"),this.originalElement.css("resize","none"),this._proportionallyResizeElements.push(this.originalElement.css({position:"static",zoom:1,display:"block"})),this._proportionallyResize()),this._setupHandles(),i.autoHide&&t(this.element).on("mouseenter",function(){!i.disabled&&(s._removeClass("ui-resizable-autohide"),s._handles.show())}).on("mouseleave",function(){!i.disabled&&(s.resizing||(s._addClass("ui-resizable-autohide"),s._handles.hide()))}),this._mouseInit()},_destroy:function(){this._mouseDestroy(),this._addedHandles.remove();var e,i=function(e){t(e).removeData("resizable").removeData("ui-resizable").off(".resizable")};return this.elementIsWrapper&&(i(this.element),e=this.element,this.originalElement.css({position:e.css("position"),width:e.outerWidth(),height:e.outerHeight(),top:e.css("top"),left:e.css("left")}).insertAfter(e),e.remove()),this.originalElement.css("resize",this.originalResizeStyle),i(this.originalElement),this},_setOption:function(t,e){switch(this._super(t,e),t){case"handles":this._removeHandles(),this._setupHandles();break;case"aspectRatio":this._aspectRatio=!!e}},_setupHandles:function(){var e,i,s,n,o,a=this.options,r=this;if(this.handles=a.handles||(t(".ui-resizable-handle",this.element).length?{n:".ui-resizable-n",e:".ui-resizable-e",s:".ui-resizable-s",w:".ui-resizable-w",se:".ui-resizable-se",sw:".ui-resizable-sw",ne:".ui-resizable-ne",nw:".ui-resizable-nw"}:"e,s,se"),this._handles=t(),this._addedHandles=t(),this.handles.constructor===String)for("all"===this.handles&&(this.handles="n,e,s,w,se,sw,ne,nw"),s=this.handles.split(","),this.handles={},i=0;i<s.length;i++)n="ui-resizable-"+(e=String.prototype.trim.call(s[i])),o=t("<div>"),this._addClass(o,"ui-resizable-handle "+n),o.css({zIndex:a.zIndex}),this.handles[e]=".ui-resizable-"+e,this.element.children(this.handles[e]).length||(this.element.append(o),this._addedHandles=this._addedHandles.add(o));this._renderAxis=function(e){var i,s,n,o;for(i in e=e||this.element,this.handles)this.handles[i].constructor===String?this.handles[i]=this.element.children(this.handles[i]).first().show():(this.handles[i].jquery||this.handles[i].nodeType)&&(this.handles[i]=t(this.handles[i]),this._on(this.handles[i],{mousedown:r._mouseDown})),this.elementIsWrapper&&this.originalElement[0].nodeName.match(/^(textarea|input|select|button)$/i)&&(s=t(this.handles[i],this.element),o=/sw|ne|nw|se|n|s/.test(i)?s.outerHeight():s.outerWidth(),n=["padding",/ne|nw|n/.test(i)?"Top":/se|sw|s/.test(i)?"Bottom":/^e$/.test(i)?"Right":"Left"].join(""),e.css(n,o),this._proportionallyResize()),this._handles=this._handles.add(this.handles[i])},this._renderAxis(this.element),this._handles=this._handles.add(this.element.find(".ui-resizable-handle")),this._handles.disableSelection(),this._handles.on("mouseover",function(){r.resizing||(this.className&&(o=this.className.match(/ui-resizable-(se|sw|ne|nw|n|e|s|w)/i)),r.axis=o&&o[1]?o[1]:"se")}),a.autoHide&&(this._handles.hide(),this._addClass("ui-resizable-autohide"))},_removeHandles:function(){this._addedHandles.remove()},_mouseCapture:function(e){var i,s,n=!1;for(i in this.handles)((s=t(this.handles[i])[0])===e.target||t.contains(s,e.target))&&(n=!0);return!this.options.disabled&&n},_mouseStart:function(e){var i,s,n,o,a=this.options,r=this.element;return this.resizing=!0,this._renderProxy(),i=this._num(this.helper.css("left")),s=this._num(this.helper.css("top")),a.containment&&(i+=t(a.containment).scrollLeft()||0,s+=t(a.containment).scrollTop()||0),this.offset=this.helper.offset(),this.position={left:i,top:s},this._helper||(o=this._calculateAdjustedElementDimensions(r)),this.size=this._helper?{width:this.helper.width(),height:this.helper.height()}:{width:o.width,height:o.height},this.originalSize=this._helper?{width:r.outerWidth(),height:r.outerHeight()}:{width:o.width,height:o.height},this.sizeDiff={width:r.outerWidth()-r.width(),height:r.outerHeight()-r.height()},this.originalPosition={left:i,top:s},this.originalMousePosition={left:e.pageX,top:e.pageY},this.aspectRatio="number"==typeof a.aspectRatio?a.aspectRatio:this.originalSize.width/this.originalSize.height||1,n=t(".ui-resizable-"+this.axis).css("cursor"),t("body").css("cursor","auto"===n?this.axis+"-resize":n),this._addClass("ui-resizable-resizing"),this._propagate("start",e),!0},_mouseDrag:function(e){var i,s,n=this.originalMousePosition,o=this.axis,a=e.pageX-n.left||0,r=e.pageY-n.top||0,l=this._change[o];return this._updatePrevProperties(),!!l&&(i=l.apply(this,[e,a,r]),this._updateVirtualBoundaries(e.shiftKey),(this._aspectRatio||e.shiftKey)&&(i=this._updateRatio(i,e)),i=this._respectSize(i,e),this._updateCache(i),this._propagate("resize",e),s=this._applyChanges(),!this._helper&&this._proportionallyResizeElements.length&&this._proportionallyResize(),t.isEmptyObject(s)||(this._updatePrevProperties(),this._trigger("resize",e,this.ui()),this._applyChanges()),!1)},_mouseStop:function(e){this.resizing=!1;var i,s,n,o,a,r,l,h=this.options;return this._helper&&(n=(s=(i=this._proportionallyResizeElements).length&&/textarea/i.test(i[0].nodeName))&&this._hasScroll(i[0],"left")?0:this.sizeDiff.height,o=s?0:this.sizeDiff.width,a={width:this.helper.width()-o,height:this.helper.height()-n},r=parseFloat(this.element.css("left"))+(this.position.left-this.originalPosition.left)||null,l=parseFloat(this.element.css("top"))+(this.position.top-this.originalPosition.top)||null,h.animate||this.element.css(t.extend(a,{top:l,left:r})),this.helper.height(this.size.height),this.helper.width(this.size.width),this._helper&&!h.animate&&this._proportionallyResize()),t("body").css("cursor","auto"),this._removeClass("ui-resizable-resizing"),this._propagate("stop",e),this._helper&&this.helper.remove(),!1},_updatePrevProperties:function(){this.prevPosition={top:this.position.top,left:this.position.left},this.prevSize={width:this.size.width,height:this.size.height}},_applyChanges:function(){var t={};return this.position.top!==this.prevPosition.top&&(t.top=this.position.top+"px"),this.position.left!==this.prevPosition.left&&(t.left=this.position.left+"px"),this.helper.css(t),this.size.width!==this.prevSize.width&&(t.width=this.size.width+"px",this.helper.width(t.width)),this.size.height!==this.prevSize.height&&(t.height=this.size.height+"px",this.helper.height(t.height)),t},_updateVirtualBoundaries:function(t){var e,i,s,n,o,a=this.options;o={minWidth:this._isNumber(a.minWidth)?a.minWidth:0,maxWidth:this._isNumber(a.maxWidth)?a.maxWidth:1/0,minHeight:this._isNumber(a.minHeight)?a.minHeight:0,maxHeight:this._isNumber(a.maxHeight)?a.maxHeight:1/0},(this._aspectRatio||t)&&(e=o.minHeight*this.aspectRatio,s=o.minWidth/this.aspectRatio,i=o.maxHeight*this.aspectRatio,n=o.maxWidth/this.aspectRatio,e>o.minWidth&&(o.minWidth=e),s>o.minHeight&&(o.minHeight=s),i<o.maxWidth&&(o.maxWidth=i),n<o.maxHeight&&(o.maxHeight=n)),this._vBoundaries=o},_updateCache:function(t){this.offset=this.helper.offset(),this._isNumber(t.left)&&(this.position.left=t.left),this._isNumber(t.top)&&(this.position.top=t.top),this._isNumber(t.height)&&(this.size.height=t.height),this._isNumber(t.width)&&(this.size.width=t.width)},_updateRatio:function(t){var e=this.position,i=this.size,s=this.axis;return this._isNumber(t.height)?t.width=t.height*this.aspectRatio:this._isNumber(t.width)&&(t.height=t.width/this.aspectRatio),"sw"===s&&(t.left=e.left+(i.width-t.width),t.top=null),"nw"===s&&(t.top=e.top+(i.height-t.height),t.left=e.left+(i.width-t.width)),t},_respectSize:function(t){var e=this._vBoundaries,i=this.axis,s=this._isNumber(t.width)&&e.maxWidth&&e.maxWidth<t.width,n=this._isNumber(t.height)&&e.maxHeight&&e.maxHeight<t.height,o=this._isNumber(t.width)&&e.minWidth&&e.minWidth>t.width,a=this._isNumber(t.height)&&e.minHeight&&e.minHeight>t.height,r=this.originalPosition.left+this.originalSize.width,l=this.originalPosition.top+this.originalSize.height,h=/sw|nw|w/.test(i),c=/nw|ne|n/.test(i);return o&&(t.width=e.minWidth),a&&(t.height=e.minHeight),s&&(t.width=e.maxWidth),n&&(t.height=e.maxHeight),o&&h&&(t.left=r-e.minWidth),s&&h&&(t.left=r-e.maxWidth),a&&c&&(t.top=l-e.minHeight),n&&c&&(t.top=l-e.maxHeight),t.width||t.height||t.left||!t.top?t.width||t.height||t.top||!t.left||(t.left=null):t.top=null,t},_getPaddingPlusBorderDimensions:function(t){for(var e=0,i=[],s=[t.css("borderTopWidth"),t.css("borderRightWidth"),t.css("borderBottomWidth"),t.css("borderLeftWidth")],n=[t.css("paddingTop"),t.css("paddingRight"),t.css("paddingBottom"),t.css("paddingLeft")];e<4;e++)i[e]=parseFloat(s[e])||0,i[e]+=parseFloat(n[e])||0;return{height:i[0]+i[2],width:i[1]+i[3]}},_calculateAdjustedElementDimensions:function(t){var e,i,s,n=t.get(0);return"content-box"===t.css("box-sizing")&&(this._hasScroll(n)||this._hasScroll(n,"left"))?(e=parseFloat(n.style.width),i=parseFloat(n.style.height),s=this._getPaddingPlusBorderDimensions(t),e=isNaN(e)?this._getElementTheoreticalSize(t,s,"width"):e,{height:i=isNaN(i)?this._getElementTheoreticalSize(t,s,"height"):i,width:e}):{height:parseFloat(t.css("height")),width:parseFloat(t.css("width"))}},_getElementTheoreticalSize:function(t,e,i){return Math.max(0,Math.ceil(t.get(0)["offset"+i[0].toUpperCase()+i.slice(1)]-e[i]-.5))||0},_proportionallyResize:function(){if(this._proportionallyResizeElements.length)for(var t,e=0,i=this.helper||this.element;e<this._proportionallyResizeElements.length;e++)t=this._proportionallyResizeElements[e],this.outerDimensions||(this.outerDimensions=this._getPaddingPlusBorderDimensions(t)),t.css({height:i.height()-this.outerDimensions.height||0,width:i.width()-this.outerDimensions.width||0})},_renderProxy:function(){var e=this.element,i=this.options;this.elementOffset=e.offset(),this._helper?(this.helper=this.helper||t("<div></div>").css({overflow:"hidden"}),this._addClass(this.helper,this._helper),this.helper.css({width:this.element.outerWidth(),height:this.element.outerHeight(),position:"absolute",left:this.elementOffset.left+"px",top:this.elementOffset.top+"px",zIndex:++i.zIndex}),this.helper.appendTo("body").disableSelection()):this.helper=this.element},_change:{e:function(t,e){return{width:this.originalSize.width+e}},w:function(t,e){var i=this.originalSize;return{left:this.originalPosition.left+e,width:i.width-e}},n:function(t,e,i){var s=this.originalSize;return{top:this.originalPosition.top+i,height:s.height-i}},s:function(t,e,i){return{height:this.originalSize.height+i}},se:function(e,i,s){return t.extend(this._change.s.apply(this,arguments),this._change.e.apply(this,[e,i,s]))},sw:function(e,i,s){return t.extend(this._change.s.apply(this,arguments),this._change.w.apply(this,[e,i,s]))},ne:function(e,i,s){return t.extend(this._change.n.apply(this,arguments),this._change.e.apply(this,[e,i,s]))},nw:function(e,i,s){return t.extend(this._change.n.apply(this,arguments),this._change.w.apply(this,[e,i,s]))}},_propagate:function(e,i){t.ui.plugin.call(this,e,[i,this.ui()]),"resize"!==e&&this._trigger(e,i,this.ui())},plugins:{},ui:function(){return{originalElement:this.originalElement,element:this.element,helper:this.helper,position:this.position,size:this.size,originalSize:this.originalSize,originalPosition:this.originalPosition}}}),t.ui.plugin.add("resizable","animate",{stop:function(e){var i=t(this).resizable("instance"),s=i.options,n=i._proportionallyResizeElements,o=n.length&&/textarea/i.test(n[0].nodeName),a=o&&i._hasScroll(n[0],"left")?0:i.sizeDiff.height,r=o?0:i.sizeDiff.width,l={width:i.size.width-r,height:i.size.height-a},h=parseFloat(i.element.css("left"))+(i.position.left-i.originalPosition.left)||null,c=parseFloat(i.element.css("top"))+(i.position.top-i.originalPosition.top)||null;i.element.animate(t.extend(l,c&&h?{top:c,left:h}:{}),{duration:s.animateDuration,easing:s.animateEasing,step:function(){var s={width:parseFloat(i.element.css("width")),height:parseFloat(i.element.css("height")),top:parseFloat(i.element.css("top")),left:parseFloat(i.element.css("left"))};n&&n.length&&t(n[0]).css({width:s.width,height:s.height}),i._updateCache(s),i._propagate("resize",e)}})}}),t.ui.plugin.add("resizable","containment",{start:function(){var e,i,s,n,o,a,r,l=t(this).resizable("instance"),h=l.options,c=l.element,u=h.containment,d=u instanceof t?u.get(0):/parent/.test(u)?c.parent().get(0):u;d&&(l.containerElement=t(d),/document/.test(u)||u===document?(l.containerOffset={left:0,top:0},l.containerPosition={left:0,top:0},l.parentData={element:t(document),left:0,top:0,width:t(document).width(),height:t(document).height()||document.body.parentNode.scrollHeight}):(e=t(d),i=[],t(["Top","Right","Left","Bottom"]).each(function(t,s){i[t]=l._num(e.css("padding"+s))}),l.containerOffset=e.offset(),l.containerPosition=e.position(),l.containerSize={height:e.innerHeight()-i[3],width:e.innerWidth()-i[1]},s=l.containerOffset,n=l.containerSize.height,o=l.containerSize.width,a=l._hasScroll(d,"left")?d.scrollWidth:o,r=l._hasScroll(d)?d.scrollHeight:n,l.parentData={element:d,left:s.left,top:s.top,width:a,height:r}))},resize:function(e){var i,s,n,o,a=t(this).resizable("instance"),r=a.options,l=a.containerOffset,h=a.position,c=a._aspectRatio||e.shiftKey,u={top:0,left:0},d=a.containerElement,p=!0;d[0]!==document&&/static/.test(d.css("position"))&&(u=l),h.left<(a._helper?l.left:0)&&(a.size.width=a.size.width+(a._helper?a.position.left-l.left:a.position.left-u.left),c&&(a.size.height=a.size.width/a.aspectRatio,p=!1),a.position.left=r.helper?l.left:0),h.top<(a._helper?l.top:0)&&(a.size.height=a.size.height+(a._helper?a.position.top-l.top:a.position.top),c&&(a.size.width=a.size.height*a.aspectRatio,p=!1),a.position.top=a._helper?l.top:0),n=a.containerElement.get(0)===a.element.parent().get(0),o=/relative|absolute/.test(a.containerElement.css("position")),n&&o?(a.offset.left=a.parentData.left+a.position.left,a.offset.top=a.parentData.top+a.position.top):(a.offset.left=a.element.offset().left,a.offset.top=a.element.offset().top),i=Math.abs(a.sizeDiff.width+(a._helper?a.offset.left-u.left:a.offset.left-l.left)),s=Math.abs(a.sizeDiff.height+(a._helper?a.offset.top-u.top:a.offset.top-l.top)),i+a.size.width>=a.parentData.width&&(a.size.width=a.parentData.width-i,c&&(a.size.height=a.size.width/a.aspectRatio,p=!1)),s+a.size.height>=a.parentData.height&&(a.size.height=a.parentData.height-s,c&&(a.size.width=a.size.height*a.aspectRatio,p=!1)),p||(a.position.left=a.prevPosition.left,a.position.top=a.prevPosition.top,a.size.width=a.prevSize.width,a.size.height=a.prevSize.height)},stop:function(){var e=t(this).resizable("instance"),i=e.options,s=e.containerOffset,n=e.containerPosition,o=e.containerElement,a=t(e.helper),r=a.offset(),l=a.outerWidth()-e.sizeDiff.width,h=a.outerHeight()-e.sizeDiff.height;e._helper&&!i.animate&&/relative/.test(o.css("position"))&&t(this).css({left:r.left-n.left-s.left,width:l,height:h}),e._helper&&!i.animate&&/static/.test(o.css("position"))&&t(this).css({left:r.left-n.left-s.left,width:l,height:h})}}),t.ui.plugin.add("resizable","alsoResize",{start:function(){var e=t(this).resizable("instance");t(e.options.alsoResize).each(function(){var i=t(this),s=e._calculateAdjustedElementDimensions(i);i.data("ui-resizable-alsoresize",{width:s.width,height:s.height,left:parseFloat(i.css("left")),top:parseFloat(i.css("top"))})})},resize:function(e,i){var s=t(this).resizable("instance"),n=s.options,o=s.originalSize,a=s.originalPosition,r={height:s.size.height-o.height||0,width:s.size.width-o.width||0,top:s.position.top-a.top||0,left:s.position.left-a.left||0};t(n.alsoResize).each(function(){var e=t(this),s=t(this).data("ui-resizable-alsoresize"),n={},o=e.parents(i.originalElement[0]).length?["width","height"]:["width","height","top","left"];t.each(o,function(t,e){var i=(s[e]||0)+(r[e]||0);i&&i>=0&&(n[e]=i||null)}),e.css(n)})},stop:function(){t(this).removeData("ui-resizable-alsoresize")}}),t.ui.plugin.add("resizable","ghost",{start:function(){var e=t(this).resizable("instance"),i=e.size;e.ghost=e.originalElement.clone(),e.ghost.css({opacity:.25,display:"block",position:"relative",height:i.height,width:i.width,margin:0,left:0,top:0}),e._addClass(e.ghost,"ui-resizable-ghost"),!0===t.uiBackCompat&&"string"==typeof e.options.ghost&&e.ghost.addClass(this.options.ghost),e.ghost.appendTo(e.helper)},resize:function(){var e=t(this).resizable("instance");e.ghost&&e.ghost.css({position:"relative",height:e.size.height,width:e.size.width})},stop:function(){var e=t(this).resizable("instance");e.ghost&&e.helper&&e.helper.get(0).removeChild(e.ghost.get(0))}}),t.ui.plugin.add("resizable","grid",{resize:function(){var e,i=t(this).resizable("instance"),s=i.options,n=i.size,o=i.originalSize,a=i.originalPosition,r=i.axis,l="number"==typeof s.grid?[s.grid,s.grid]:s.grid,h=l[0]||1,c=l[1]||1,u=Math.round((n.width-o.width)/h)*h,d=Math.round((n.height-o.height)/c)*c,p=o.width+u,f=o.height+d,g=s.maxWidth&&s.maxWidth<p,m=s.maxHeight&&s.maxHeight<f,_=s.minWidth&&s.minWidth>p,v=s.minHeight&&s.minHeight>f;s.grid=l,_&&(p+=h),v&&(f+=c),g&&(p-=h),m&&(f-=c),/^(se|s|e)$/.test(r)?(i.size.width=p,i.size.height=f):/^(ne)$/.test(r)?(i.size.width=p,i.size.height=f,i.position.top=a.top-d):/^(sw)$/.test(r)?(i.size.width=p,i.size.height=f,i.position.left=a.left-u):((f-c<=0||p-h<=0)&&(e=i._getPaddingPlusBorderDimensions(this)),f-c>0?(i.size.height=f,i.position.top=a.top-d):(f=c-e.height,i.size.height=f,i.position.top=a.top+o.height-f),p-h>0?(i.size.width=p,i.position.left=a.left-u):(p=h-e.width,i.size.width=p,i.position.left=a.left+o.width-p))}}),t.ui.resizable,t.widget("ui.dialog",{version:"1.14.2",options:{appendTo:"body",autoOpen:!0,buttons:[],classes:{"ui-dialog":"ui-corner-all","ui-dialog-titlebar":"ui-corner-all"},closeOnEscape:!0,closeText:"Close",draggable:!0,hide:null,height:"auto",maxHeight:null,maxWidth:null,minHeight:150,minWidth:150,modal:!1,position:{my:"center",at:"center",of:window,collision:"fit",using:function(e){var i=t(this).css(e).offset().top;i<0&&t(this).css("top",e.top-i)}},resizable:!0,show:null,title:null,uiDialogTitleHeadingLevel:0,width:300,beforeClose:null,close:null,drag:null,dragStart:null,dragStop:null,focus:null,open:null,resize:null,resizeStart:null,resizeStop:null},sizeRelatedOptions:{buttons:!0,height:!0,maxHeight:!0,maxWidth:!0,minHeight:!0,minWidth:!0,width:!0},resizableRelatedOptions:{maxHeight:!0,maxWidth:!0,minHeight:!0,minWidth:!0},_create:function(){this.originalCss={display:this.element[0].style.display,width:this.element[0].style.width,minHeight:this.element[0].style.minHeight,maxHeight:this.element[0].style.maxHeight,height:this.element[0].style.height},this.originalPosition={parent:this.element.parent(),index:this.element.parent().children().index(this.element)},this.originalTitle=this.element.attr("title"),null==this.options.title&&null!=this.originalTitle&&(this.options.title=this.originalTitle),this.options.disabled&&(this.options.disabled=!1),this._createWrapper(),this.element.show().removeAttr("title").appendTo(this.uiDialog),this._addClass("ui-dialog-content","ui-widget-content"),this._createTitlebar(),this._createButtonPane(),this.options.draggable&&t.fn.draggable&&this._makeDraggable(),this.options.resizable&&t.fn.resizable&&this._makeResizable(),this._isOpen=!1,this._trackFocus()},_init:function(){this.options.autoOpen&&this.open()},_appendTo:function(){var e=this.options.appendTo;return e&&(e.jquery||e.nodeType)?t(e):this.document.find(e||"body").eq(0)},_destroy:function(){var t,e=this.originalPosition;this._untrackInstance(),this._destroyOverlay(),this.element.removeUniqueId().css(this.originalCss).detach(),this.uiDialog.remove(),this.originalTitle&&this.element.attr("title",this.originalTitle),(t=e.parent.children().eq(e.index)).length&&t[0]!==this.element[0]?t.before(this.element):e.parent.append(this.element)},widget:function(){return this.uiDialog},disable:t.noop,enable:t.noop,close:function(e){var i=this;this._isOpen&&!1!==this._trigger("beforeClose",e)&&(this._isOpen=!1,this._focusedElement=null,this._destroyOverlay(),this._untrackInstance(),this.opener.filter(":focusable").trigger("focus").length||t(this.document[0].activeElement).trigger("blur"),this._hide(this.uiDialog,this.options.hide,function(){i._trigger("close",e)}))},isOpen:function(){return this._isOpen},moveToTop:function(){this._moveToTop()},_moveToTop:function(e,i){var s=!1,n=this.uiDialog.siblings(".ui-front:visible").map(function(){return+t(this).css("z-index")}).get(),o=Math.max.apply(null,n);return o>=+this.uiDialog.css("z-index")&&(this.uiDialog.css("z-index",o+1),s=!0),s&&!i&&this._trigger("focus",e),s},open:function(){var e=this;if(this._isOpen){this._moveToTop()&&this._focusTabbable();return}this._isOpen=!0,this.opener=t(this.document[0].activeElement),this._size(),this._position(),this._createOverlay(),this._moveToTop(null,!0),this.overlay&&this.overlay.css("z-index",this.uiDialog.css("z-index")-1),this._show(this.uiDialog,this.options.show,function(){e._focusTabbable(),e._trigger("focus")}),this._makeFocusTarget(),this._trigger("open")},_focusTabbable:function(){var t=this._focusedElement;t||(t=this.element.find("[autofocus]")),t.length||(t=this.element.find(":tabbable")),t.length||(t=this.uiDialogButtonPane.find(":tabbable")),t.length||(t=this.uiDialogTitlebarClose.filter(":tabbable")),t.length||(t=this.uiDialog),t.eq(0).trigger("focus")},_restoreTabbableFocus:function(){var e=this.document[0].activeElement;this.uiDialog[0]===e||t.contains(this.uiDialog[0],e)||this._focusTabbable()},_keepFocus:function(t){t.preventDefault(),this._restoreTabbableFocus()},_createWrapper:function(){this.uiDialog=t("<div>").hide().attr({tabIndex:-1,role:"dialog","aria-modal":this.options.modal?"true":null}).appendTo(this._appendTo()),this._addClass(this.uiDialog,"ui-dialog","ui-widget ui-widget-content ui-front"),this._on(this.uiDialog,{keydown:function(e){if(this.options.closeOnEscape&&!e.isDefaultPrevented()&&e.keyCode&&e.keyCode===t.ui.keyCode.ESCAPE){e.preventDefault(),this.close(e);return}if(!(e.keyCode!==t.ui.keyCode.TAB||e.isDefaultPrevented())){var i=this.uiDialog.find(":tabbable"),s=i.first(),n=i.last();e.target!==n[0]&&e.target!==this.uiDialog[0]||e.shiftKey?(e.target===s[0]||e.target===this.uiDialog[0])&&e.shiftKey&&(this._delay(function(){n.trigger("focus")}),e.preventDefault()):(this._delay(function(){s.trigger("focus")}),e.preventDefault())}},mousedown:function(t){this._moveToTop(t)&&this._focusTabbable()}}),this.element.find("[aria-describedby]").length||this.uiDialog.attr({"aria-describedby":this.element.uniqueId().attr("id")})},_createTitlebar:function(){var e;this.uiDialogTitlebar=t("<div>"),this._addClass(this.uiDialogTitlebar,"ui-dialog-titlebar","ui-widget-header ui-helper-clearfix"),this._on(this.uiDialogTitlebar,{mousedown:function(e){t(e.target).closest(".ui-dialog-titlebar-close")||this.uiDialog.trigger("focus")}}),this.uiDialogTitlebarClose=t("<button type='button'></button>").button({label:t("<a>").text(this.options.closeText).html(),icon:"ui-icon-closethick",showLabel:!1}).appendTo(this.uiDialogTitlebar),this._addClass(this.uiDialogTitlebarClose,"ui-dialog-titlebar-close"),this._on(this.uiDialogTitlebarClose,{click:function(t){t.preventDefault(),this.close(t)}}),e=t("<"+(Number.isInteger(this.options.uiDialogTitleHeadingLevel)&&this.options.uiDialogTitleHeadingLevel>0&&this.options.uiDialogTitleHeadingLevel<=6?"h"+this.options.uiDialogTitleHeadingLevel:"span")+">").uniqueId().prependTo(this.uiDialogTitlebar),this._addClass(e,"ui-dialog-title"),this._title(e),this.uiDialogTitlebar.prependTo(this.uiDialog),this.uiDialog.attr({"aria-labelledby":e.attr("id")})},_title:function(t){this.options.title?t.text(this.options.title):t.html("&#160;")},_createButtonPane:function(){this.uiDialogButtonPane=t("<div>"),this._addClass(this.uiDialogButtonPane,"ui-dialog-buttonpane","ui-widget-content ui-helper-clearfix"),this.uiButtonSet=t("<div>").appendTo(this.uiDialogButtonPane),this._addClass(this.uiButtonSet,"ui-dialog-buttonset"),this._createButtons()},_createButtons:function(){var e=this,i=this.options.buttons;if(this.uiDialogButtonPane.remove(),this.uiButtonSet.empty(),t.isEmptyObject(i)||Array.isArray(i)&&!i.length){this._removeClass(this.uiDialog,"ui-dialog-buttons");return}t.each(i,function(i,s){var n,o;s="function"==typeof s?{click:s,text:i}:s,n=(s=t.extend({type:"button"},s)).click,o={icon:s.icon,iconPosition:s.iconPosition,showLabel:s.showLabel,icons:s.icons,text:s.text},delete s.click,delete s.icon,delete s.iconPosition,delete s.showLabel,delete s.icons,"boolean"==typeof s.text&&delete s.text,t("<button></button>",s).button(o).appendTo(e.uiButtonSet).on("click",function(){n.apply(e.element[0],arguments)})}),this._addClass(this.uiDialog,"ui-dialog-buttons"),this.uiDialogButtonPane.appendTo(this.uiDialog)},_makeDraggable:function(){var e=this,i=this.options;function s(t){return{position:t.position,offset:t.offset}}this.uiDialog.draggable({cancel:".ui-dialog-content, .ui-dialog-titlebar-close",handle:".ui-dialog-titlebar",containment:"document",start:function(i,n){e._addClass(t(this),"ui-dialog-dragging"),e._blockFrames(),e._trigger("dragStart",i,s(n))},drag:function(t,i){e._trigger("drag",t,s(i))},stop:function(n,o){var a=o.offset.left-e.document.scrollLeft(),r=o.offset.top-e.document.scrollTop();i.position={my:"left top",at:"left"+(a>=0?"+":"")+a+" top"+(r>=0?"+":"")+r,of:e.window},e._removeClass(t(this),"ui-dialog-dragging"),e._unblockFrames(),e._trigger("dragStop",n,s(o))}})},_makeResizable:function(){var e=this,i=this.options,s=i.resizable,n=this.uiDialog.css("position");function o(t){return{originalPosition:t.originalPosition,originalSize:t.originalSize,position:t.position,size:t.size}}this.uiDialog.resizable({cancel:".ui-dialog-content",containment:"document",alsoResize:this.element,maxWidth:i.maxWidth,maxHeight:i.maxHeight,minWidth:i.minWidth,minHeight:this._minHeight(),handles:"string"==typeof s?s:"n,e,s,w,se,sw,ne,nw",start:function(i,s){e._addClass(t(this),"ui-dialog-resizing"),e._blockFrames(),e._trigger("resizeStart",i,o(s))},resize:function(t,i){e._trigger("resize",t,o(i))},stop:function(s,n){var a=e.uiDialog.offset(),r=a.left-e.document.scrollLeft(),l=a.top-e.document.scrollTop();i.height=e.uiDialog.height(),i.width=e.uiDialog.width(),i.position={my:"left top",at:"left"+(r>=0?"+":"")+r+" top"+(l>=0?"+":"")+l,of:e.window},e._removeClass(t(this),"ui-dialog-resizing"),e._unblockFrames(),e._trigger("resizeStop",s,o(n))}}).css("position",n)},_trackFocus:function(){this._on(this.widget(),{focusin:function(e){this._makeFocusTarget(),this._focusedElement=t(e.target)}})},_makeFocusTarget:function(){this._untrackInstance(),this._trackingInstances().unshift(this)},_untrackInstance:function(){var e=this._trackingInstances(),i=t.inArray(this,e);-1!==i&&e.splice(i,1)},_trackingInstances:function(){var t=this.document.data("ui-dialog-instances");return t||(t=[],this.document.data("ui-dialog-instances",t)),t},_minHeight:function(){var t=this.options;return"auto"===t.height?t.minHeight:Math.min(t.minHeight,t.height)},_position:function(){var t=this.uiDialog.is(":visible");t||this.uiDialog.show(),this.uiDialog.position(this.options.position),t||this.uiDialog.hide()},_setOptions:function(e){var i=this,s=!1,n={};t.each(e,function(t,e){i._setOption(t,e),t in i.sizeRelatedOptions&&(s=!0),t in i.resizableRelatedOptions&&(n[t]=e)}),s&&(this._size(),this._position()),this.uiDialog.is(":data(ui-resizable)")&&this.uiDialog.resizable("option",n)},_setOption:function(e,i){var s,n,o=this.uiDialog;"disabled"!==e&&(this._super(e,i),"appendTo"===e&&this.uiDialog.appendTo(this._appendTo()),"buttons"===e&&this._createButtons(),"closeText"===e&&this.uiDialogTitlebarClose.button({label:t("<a>").text(""+this.options.closeText).html()}),"draggable"===e&&((s=o.is(":data(ui-draggable)"))&&!i&&o.draggable("destroy"),!s&&i&&this._makeDraggable()),"position"===e&&this._position(),"resizable"!==e||((n=o.is(":data(ui-resizable)"))&&!i&&o.resizable("destroy"),n&&"string"==typeof i&&o.resizable("option","handles",i),n||!1===i||this._makeResizable()),"title"===e&&this._title(this.uiDialogTitlebar.find(".ui-dialog-title")),"modal"===e&&o.attr("aria-modal",i?"true":null))},_size:function(){var t,e,i,s=this.options;this.element.show().css({width:"auto",minHeight:0,maxHeight:"none",height:0}),s.minWidth>s.width&&(s.width=s.minWidth),t=this.uiDialog.css({height:"auto",width:s.width}).outerHeight(),e=Math.max(0,s.minHeight-t),i="number"==typeof s.maxHeight?Math.max(0,s.maxHeight-t):"none","auto"===s.height?this.element.css({minHeight:e,maxHeight:i,height:"auto"}):this.element.height(Math.max(0,s.height-t)),this.uiDialog.is(":data(ui-resizable)")&&this.uiDialog.resizable("option","minHeight",this._minHeight())},_blockFrames:function(){this.iframeBlocks=this.document.find("iframe").map(function(){var e=t(this);return t("<div>").css({position:"absolute",width:e.outerWidth(),height:e.outerHeight()}).appendTo(e.parent()).offset(e.offset())[0]})},_unblockFrames:function(){this.iframeBlocks&&(this.iframeBlocks.remove(),delete this.iframeBlocks)},_allowInteraction:function(e){return!!t(e.target).closest(".ui-dialog").length||!!t(e.target).closest(".ui-datepicker").length},_createOverlay:function(){if(this.options.modal){var e=!0;this._delay(function(){e=!1}),this.document.data("ui-dialog-overlays")||this.document.on("focusin.ui-dialog",(function(t){if(!e){var i=this._trackingInstances()[0];i._allowInteraction(t)||(t.preventDefault(),i._focusTabbable())}}).bind(this)),this.overlay=t("<div>").appendTo(this._appendTo()),this._addClass(this.overlay,null,"ui-widget-overlay ui-front"),this._on(this.overlay,{mousedown:"_keepFocus"}),this.document.data("ui-dialog-overlays",(this.document.data("ui-dialog-overlays")||0)+1)}},_destroyOverlay:function(){if(this.options.modal&&this.overlay){var t=this.document.data("ui-dialog-overlays")-1;t?this.document.data("ui-dialog-overlays",t):(this.document.off("focusin.ui-dialog"),this.document.removeData("ui-dialog-overlays")),this.overlay.remove(),this.overlay=null}}}),!0===t.uiBackCompat&&t.widget("ui.dialog",t.ui.dialog,{options:{dialogClass:""},_createWrapper:function(){this._super(),this.uiDialog.addClass(this.options.dialogClass)},_setOption:function(t,e){"dialogClass"===t&&this.uiDialog.removeClass(this.options.dialogClass).addClass(e),this._superApply(arguments)}}),t.ui.dialog,t.widget("ui.droppable",{version:"1.14.2",widgetEventPrefix:"drop",options:{accept:"*",addClasses:!0,greedy:!1,scope:"default",tolerance:"intersect",activate:null,deactivate:null,drop:null,out:null,over:null},_create:function(){var t,e=this.options,i=e.accept;this.isover=!1,this.isout=!0,this.accept="function"==typeof i?i:function(t){return t.is(i)},this.proportions=function(){if(!arguments.length)return t||(t={width:this.element[0].offsetWidth,height:this.element[0].offsetHeight});t=arguments[0]},this._addToManager(e.scope),e.addClasses&&this._addClass("ui-droppable")},_addToManager:function(e){t.ui.ddmanager.droppables[e]=t.ui.ddmanager.droppables[e]||[],t.ui.ddmanager.droppables[e].push(this)},_splice:function(t){for(var e=0;e<t.length;e++)t[e]===this&&t.splice(e,1)},_destroy:function(){var e=t.ui.ddmanager.droppables[this.options.scope];this._splice(e)},_setOption:function(e,i){if("accept"===e)this.accept="function"==typeof i?i:function(t){return t.is(i)};else if("scope"===e){var s=t.ui.ddmanager.droppables[this.options.scope];this._splice(s),this._addToManager(i)}this._super(e,i)},_activate:function(e){var i=t.ui.ddmanager.current;this._addActiveClass(),i&&this._trigger("activate",e,this.ui(i))},_deactivate:function(e){var i=t.ui.ddmanager.current;this._removeActiveClass(),i&&this._trigger("deactivate",e,this.ui(i))},_over:function(e){var i=t.ui.ddmanager.current;i&&(i.currentItem||i.element)[0]!==this.element[0]&&this.accept.call(this.element[0],i.currentItem||i.element)&&(this._addHoverClass(),this._trigger("over",e,this.ui(i)))},_out:function(e){var i=t.ui.ddmanager.current;i&&(i.currentItem||i.element)[0]!==this.element[0]&&this.accept.call(this.element[0],i.currentItem||i.element)&&(this._removeHoverClass(),this._trigger("out",e,this.ui(i)))},_drop:function(e,i){var s=i||t.ui.ddmanager.current,n=!1;return!!s&&(s.currentItem||s.element)[0]!==this.element[0]&&(this.element.find(":data(ui-droppable)").not(".ui-draggable-dragging").each(function(){var i=t(this).droppable("instance");if(i.options.greedy&&!i.options.disabled&&i.options.scope===s.options.scope&&i.accept.call(i.element[0],s.currentItem||s.element)&&t.ui.intersect(s,t.extend(i,{offset:i.element.offset()}),i.options.tolerance,e))return n=!0,!1}),!n&&!!this.accept.call(this.element[0],s.currentItem||s.element)&&(this._removeActiveClass(),this._removeHoverClass(),this._trigger("drop",e,this.ui(s)),this.element))},ui:function(t){return{draggable:t.currentItem||t.element,helper:t.helper,position:t.position,offset:t.positionAbs}},_addHoverClass:function(){this._addClass("ui-droppable-hover")},_removeHoverClass:function(){this._removeClass("ui-droppable-hover")},_addActiveClass:function(){this._addClass("ui-droppable-active")},_removeActiveClass:function(){this._removeClass("ui-droppable-active")}}),t.ui.intersect=function(t,e,i,s){if(!e.offset)return!1;var n,o,a,r,l=(t.positionAbs||t.position.absolute).left+t.margins.left,h=(t.positionAbs||t.position.absolute).top+t.margins.top,c=l+t.helperProportions.width,u=h+t.helperProportions.height,d=e.offset.left,p=e.offset.top,f=d+e.proportions().width,g=p+e.proportions().height;switch(i){case"fit":return d<=l&&c<=f&&p<=h&&u<=g;case"intersect":return d<l+t.helperProportions.width/2&&c-t.helperProportions.width/2<f&&p<h+t.helperProportions.height/2&&u-t.helperProportions.height/2<g;case"pointer":return n=s.pageY,o=e.proportions().height,n>=p&&n<p+o&&(a=s.pageX,r=e.proportions().width,a>=d&&a<d+r);case"touch":return(h>=p&&h<=g||u>=p&&u<=g||h<p&&u>g)&&(l>=d&&l<=f||c>=d&&c<=f||l<d&&c>f);default:return!1}},t.ui.ddmanager={current:null,droppables:{default:[]},prepareOffsets:function(e,i){var s,n,o=t.ui.ddmanager.droppables[e.options.scope]||[],a=i?i.type:null,r=(e.currentItem||e.element).find(":data(ui-droppable)").addBack();t:for(s=0;s<o.length;s++){if(!o[s].options.disabled&&(!e||o[s].accept.call(o[s].element[0],e.currentItem||e.element))){for(n=0;n<r.length;n++)if(r[n]===o[s].element[0]){o[s].proportions().height=0;continue t}o[s].visible="none"!==o[s].element.css("display"),o[s].visible&&("mousedown"===a&&o[s]._activate.call(o[s],i),o[s].offset=o[s].element.offset(),o[s].proportions({width:o[s].element[0].offsetWidth,height:o[s].element[0].offsetHeight}))}}},drop:function(e,i){var s=!1;return t.each((t.ui.ddmanager.droppables[e.options.scope]||[]).slice(),function(){this.options&&(!this.options.disabled&&this.visible&&t.ui.intersect(e,this,this.options.tolerance,i)&&(s=this._drop.call(this,i)||s),!this.options.disabled&&this.visible&&this.accept.call(this.element[0],e.currentItem||e.element)&&(this.isout=!0,this.isover=!1,this._deactivate.call(this,i)))}),s},dragStart:function(e,i){e.element.parentsUntil("body").on("scroll.droppable",function(){e.options.refreshPositions||t.ui.ddmanager.prepareOffsets(e,i)})},drag:function(e,i){e.options.refreshPositions&&t.ui.ddmanager.prepareOffsets(e,i),t.each(t.ui.ddmanager.droppables[e.options.scope]||[],function(){if(!this.options.disabled&&!this.greedyChild&&this.visible){var s,n,o,a=t.ui.intersect(e,this,this.options.tolerance,i),r=!a&&this.isover?"isout":a&&!this.isover?"isover":null;r&&(this.options.greedy&&(n=this.options.scope,(o=this.element.parents(":data(ui-droppable)").filter(function(){return t(this).droppable("instance").options.scope===n})).length&&((s=t(o[0]).droppable("instance")).greedyChild="isover"===r)),s&&"isover"===r&&(s.isover=!1,s.isout=!0,s._out.call(s,i)),this[r]=!0,this["isout"===r?"isover":"isout"]=!1,this["isover"===r?"_over":"_out"].call(this,i),s&&"isout"===r&&(s.isout=!1,s.isover=!0,s._over.call(s,i)))}})},dragStop:function(e,i){e.element.parentsUntil("body").off("scroll.droppable"),e.options.refreshPositions||t.ui.ddmanager.prepareOffsets(e,i)}},!0===t.uiBackCompat&&t.widget("ui.droppable",t.ui.droppable,{options:{hoverClass:!1,activeClass:!1},_addActiveClass:function(){this._super(),this.options.activeClass&&this.element.addClass(this.options.activeClass)},_removeActiveClass:function(){this._super(),this.options.activeClass&&this.element.removeClass(this.options.activeClass)},_addHoverClass:function(){this._super(),this.options.hoverClass&&this.element.addClass(this.options.hoverClass)},_removeHoverClass:function(){this._super(),this.options.hoverClass&&this.element.removeClass(this.options.hoverClass)}}),t.ui.droppable,t.widget("ui.progressbar",{version:"1.14.2",options:{classes:{"ui-progressbar":"ui-corner-all","ui-progressbar-value":"ui-corner-left","ui-progressbar-complete":"ui-corner-right"},max:100,value:0,change:null,complete:null},min:0,_create:function(){this.oldValue=this.options.value=this._constrainedValue(),this.element.attr({role:"progressbar","aria-valuemin":this.min}),this._addClass("ui-progressbar","ui-widget ui-widget-content"),this.valueDiv=t("<div>").appendTo(this.element),this._addClass(this.valueDiv,"ui-progressbar-value","ui-widget-header"),this._refreshValue()},_destroy:function(){this.element.removeAttr("role aria-valuemin aria-valuemax aria-valuenow"),this.valueDiv.remove()},value:function(t){if(void 0===t)return this.options.value;this.options.value=this._constrainedValue(t),this._refreshValue()},_constrainedValue:function(t){return void 0===t&&(t=this.options.value),this.indeterminate=!1===t,"number"!=typeof t&&(t=0),!this.indeterminate&&Math.min(this.options.max,Math.max(this.min,t))},_setOptions:function(t){var e=t.value;delete t.value,this._super(t),this.options.value=this._constrainedValue(e),this._refreshValue()},_setOption:function(t,e){"max"===t&&(e=Math.max(this.min,e)),this._super(t,e)},_setOptionDisabled:function(t){this._super(t),this.element.attr("aria-disabled",t),this._toggleClass(null,"ui-state-disabled",!!t)},_percentage:function(){return this.indeterminate?100:100*(this.options.value-this.min)/(this.options.max-this.min)},_refreshValue:function(){var e=this.options.value,i=this._percentage();this.valueDiv.toggle(this.indeterminate||e>this.min).width(i.toFixed(0)+"%"),this._toggleClass(this.valueDiv,"ui-progressbar-complete",null,e===this.options.max)._toggleClass("ui-progressbar-indeterminate",null,this.indeterminate),this.indeterminate?(this.element.removeAttr("aria-valuenow"),this.overlayDiv||(this.overlayDiv=t("<div>").appendTo(this.valueDiv),this._addClass(this.overlayDiv,"ui-progressbar-overlay"))):(this.element.attr({"aria-valuemax":this.options.max,"aria-valuenow":e}),this.overlayDiv&&(this.overlayDiv.remove(),this.overlayDiv=null)),this.oldValue!==e&&(this.oldValue=e,this._trigger("change")),e===this.options.max&&this._trigger("complete")}}),t.widget("ui.selectable",t.ui.mouse,{version:"1.14.2",options:{appendTo:"body",autoRefresh:!0,distance:0,filter:"*",tolerance:"touch",selected:null,selecting:null,start:null,stop:null,unselected:null,unselecting:null},_create:function(){var e=this;this._addClass("ui-selectable"),this.dragged=!1,this.refresh=function(){e.elementPos=t(e.element[0]).offset(),e.selectees=t(e.options.filter,e.element[0]),e._addClass(e.selectees,"ui-selectee"),e.selectees.each(function(){var i=t(this),s=i.offset(),n={left:s.left-e.elementPos.left,top:s.top-e.elementPos.top};t.data(this,"selectable-item",{element:this,$element:i,left:n.left,top:n.top,right:n.left+i.outerWidth(),bottom:n.top+i.outerHeight(),startselected:!1,selected:i.hasClass("ui-selected"),selecting:i.hasClass("ui-selecting"),unselecting:i.hasClass("ui-unselecting")})})},this.refresh(),this._mouseInit(),this.helper=t("<div>"),this._addClass(this.helper,"ui-selectable-helper")},_destroy:function(){this.selectees.removeData("selectable-item"),this._mouseDestroy()},_mouseStart:function(e){var i=this,s=this.options;this.opos=[e.pageX,e.pageY],this.elementPos=t(this.element[0]).offset(),!this.options.disabled&&(this.selectees=t(s.filter,this.element[0]),this._trigger("start",e),t(s.appendTo).append(this.helper),this.helper.css({left:e.pageX,top:e.pageY,width:0,height:0}),s.autoRefresh&&this.refresh(),this.selectees.filter(".ui-selected").each(function(){var s=t.data(this,"selectable-item");s.startselected=!0,e.metaKey||e.ctrlKey||(i._removeClass(s.$element,"ui-selected"),s.selected=!1,i._addClass(s.$element,"ui-unselecting"),s.unselecting=!0,i._trigger("unselecting",e,{unselecting:s.element}))}),t(e.target).parents().addBack().each(function(){var s,n=t.data(this,"selectable-item");if(n)return s=!e.metaKey&&!e.ctrlKey||!n.$element.hasClass("ui-selected"),i._removeClass(n.$element,s?"ui-unselecting":"ui-selected")._addClass(n.$element,s?"ui-selecting":"ui-unselecting"),n.unselecting=!s,n.selecting=s,n.selected=s,s?i._trigger("selecting",e,{selecting:n.element}):i._trigger("unselecting",e,{unselecting:n.element}),!1}))},_mouseDrag:function(e){if(this.dragged=!0,!this.options.disabled){var i,s=this,n=this.options,o=this.opos[0],a=this.opos[1],r=e.pageX,l=e.pageY;return o>r&&(i=r,r=o,o=i),a>l&&(i=l,l=a,a=i),this.helper.css({left:o,top:a,width:r-o,height:l-a}),this.selectees.each(function(){var i=t.data(this,"selectable-item"),h=!1,c={};i&&i.element!==s.element[0]&&(c.left=i.left+s.elementPos.left,c.right=i.right+s.elementPos.left,c.top=i.top+s.elementPos.top,c.bottom=i.bottom+s.elementPos.top,"touch"===n.tolerance?h=!(c.left>r||c.right<o||c.top>l||c.bottom<a):"fit"===n.tolerance&&(h=c.left>o&&c.right<r&&c.top>a&&c.bottom<l),h?(i.selected&&(s._removeClass(i.$element,"ui-selected"),i.selected=!1),i.unselecting&&(s._removeClass(i.$element,"ui-unselecting"),i.unselecting=!1),i.selecting||(s._addClass(i.$element,"ui-selecting"),i.selecting=!0,s._trigger("selecting",e,{selecting:i.element}))):(i.selecting&&((e.metaKey||e.ctrlKey)&&i.startselected?(s._removeClass(i.$element,"ui-selecting"),i.selecting=!1,s._addClass(i.$element,"ui-selected"),i.selected=!0):(s._removeClass(i.$element,"ui-selecting"),i.selecting=!1,i.startselected&&(s._addClass(i.$element,"ui-unselecting"),i.unselecting=!0),s._trigger("unselecting",e,{unselecting:i.element}))),!i.selected||e.metaKey||e.ctrlKey||i.startselected||(s._removeClass(i.$element,"ui-selected"),i.selected=!1,s._addClass(i.$element,"ui-unselecting"),i.unselecting=!0,s._trigger("unselecting",e,{unselecting:i.element}))))}),!1}},_mouseStop:function(e){var i=this;return this.dragged=!1,t(".ui-unselecting",this.element[0]).each(function(){var s=t.data(this,"selectable-item");i._removeClass(s.$element,"ui-unselecting"),s.unselecting=!1,s.startselected=!1,i._trigger("unselected",e,{unselected:s.element})}),t(".ui-selecting",this.element[0]).each(function(){var s=t.data(this,"selectable-item");i._removeClass(s.$element,"ui-selecting")._addClass(s.$element,"ui-selected"),s.selecting=!1,s.selected=!0,s.startselected=!0,i._trigger("selected",e,{selected:s.element})}),this._trigger("stop",e),this.helper.remove(),!1}}),t.widget("ui.selectmenu",[t.ui.formResetMixin,{version:"1.14.2",defaultElement:"<select>",options:{appendTo:null,classes:{"ui-selectmenu-button-open":"ui-corner-top","ui-selectmenu-button-closed":"ui-corner-all"},disabled:null,icons:{button:"ui-icon-triangle-1-s"},position:{my:"left top",at:"left bottom",collision:"none"},width:!1,change:null,close:null,focus:null,open:null,select:null},_create:function(){var e=this.element.uniqueId().attr("id");this.ids={element:e,button:e+"-button",menu:e+"-menu"},this._drawButton(),this._drawMenu(),this._bindFormResetHandler(),this._rendered=!1,this.menuItems=t()},_drawButton:function(){var e,i=this,s=this._parseOption(this.element.find("option:selected"),this.element[0].selectedIndex);this.labels=this.element.labels().attr("for",this.ids.button),this._on(this.labels,{click:function(t){this.button.trigger("focus"),t.preventDefault()}}),this.element.hide(),this.button=t("<span>",{tabindex:this.options.disabled?-1:0,id:this.ids.button,role:"combobox","aria-expanded":"false","aria-autocomplete":"list","aria-owns":this.ids.menu,"aria-haspopup":"true",title:this.element.attr("title")}).insertAfter(this.element),this._addClass(this.button,"ui-selectmenu-button ui-selectmenu-button-closed","ui-button ui-widget"),e=t("<span>").appendTo(this.button),this._addClass(e,"ui-selectmenu-icon","ui-icon "+this.options.icons.button),this.buttonItem=this._renderButtonItem(s).appendTo(this.button),!1!==this.options.width&&this._resizeButton(),this._on(this.button,this._buttonEvents),this.button.one("focusin",function(){i._rendered||i._refreshMenu()})},_drawMenu:function(){var e=this;this.menu=t("<ul>",{"aria-hidden":"true","aria-labelledby":this.ids.button,id:this.ids.menu}),this.menuWrap=t("<div>").append(this.menu),this._addClass(this.menuWrap,"ui-selectmenu-menu","ui-front"),this.menuWrap.appendTo(this._appendTo()),this.menuInstance=this.menu.menu({classes:{"ui-menu":"ui-corner-bottom"},role:"listbox",select:function(t,i){t.preventDefault(),e._select(i.item.data("ui-selectmenu-item"),t)},focus:function(t,i){var s=i.item.data("ui-selectmenu-item");null==e.focusIndex||s.index===e.focusIndex||(e._trigger("focus",t,{item:s}),e.isOpen||e._select(s,t)),e.focusIndex=s.index,e.button.attr("aria-activedescendant",e.menuItems.eq(s.index).attr("id"))}}).menu("instance"),this.menuInstance._off(this.menu,"mouseleave"),this.menuInstance._closeOnDocumentClick=function(){return!1},this.menuInstance._isDivider=function(){return!1}},refresh:function(){this._refreshMenu(),this.buttonItem.replaceWith(this.buttonItem=this._renderButtonItem(this._getSelectedItem().data("ui-selectmenu-item")||{})),null===this.options.width&&this._resizeButton()},_refreshMenu:function(){var t,e=this.element.find("option");this.menu.empty(),this._parseOptions(e),this._renderMenu(this.menu,this.items),this.menuInstance.refresh(),this.menuItems=this.menu.find("li").not(".ui-selectmenu-optgroup").find(".ui-menu-item-wrapper"),this._rendered=!0,e.length&&(t=this._getSelectedItem(),this.menuInstance.focus(null,t),this._setAria(t.data("ui-selectmenu-item")),this._setOption("disabled",this.element.prop("disabled")))},open:function(t){if(!this.options.disabled)this._rendered?(this._removeClass(this.menu.find(".ui-state-active"),null,"ui-state-active"),this.menuInstance.focus(null,this._getSelectedItem())):this._refreshMenu(),this.menuItems.length&&(this.isOpen=!0,this._toggleAttr(),this._resizeMenu(),this._position(),this._on(this.document,this._documentClick),this._trigger("open",t))},_position:function(){this.menuWrap.position(t.extend({of:this.button},this.options.position))},close:function(t){this.isOpen&&(this.isOpen=!1,this._toggleAttr(),this.range=null,this._off(this.document),this._trigger("close",t))},widget:function(){return this.button},menuWidget:function(){return this.menu},_renderButtonItem:function(e){var i=t("<span>");return this._setText(i,e.label),this._addClass(i,"ui-selectmenu-text"),i},_renderMenu:function(e,i){var s=this,n="";t.each(i,function(i,o){var a;o.optgroup!==n&&(a=t("<li>",{text:o.optgroup}),s._addClass(a,"ui-selectmenu-optgroup","ui-menu-divider"+(o.element.parent("optgroup").prop("disabled")?" ui-state-disabled":"")),a.appendTo(e),n=o.optgroup),s._renderItemData(e,o)})},_renderItemData:function(t,e){return this._renderItem(t,e).data("ui-selectmenu-item",e)},_renderItem:function(e,i){var s=t("<li>"),n=t("<div>",{title:i.element.attr("title")});return i.disabled&&this._addClass(s,null,"ui-state-disabled"),i.hidden?s.prop("hidden",!0):this._setText(n,i.label),s.append(n).appendTo(e)},_setText:function(t,e){e?t.text(e):t.html("&#160;")},_move:function(t,e){var i,s,n=".ui-menu-item";this.isOpen?i=this.menuItems.eq(this.focusIndex).parent("li"):(i=this.menuItems.eq(this.element[0].selectedIndex).parent("li"),n+=":not(.ui-state-disabled)"),(s="first"===t||"last"===t?i["first"===t?"prevAll":"nextAll"](n).eq(-1):i[t+"All"](n).eq(0)).length&&this.menuInstance.focus(e,s)},_getSelectedItem:function(){return this.menuItems.eq(this.element[0].selectedIndex).parent("li")},_toggle:function(t){this[this.isOpen?"close":"open"](t)},_setSelection:function(){var t;this.range&&((t=window.getSelection()).removeAllRanges(),t.addRange(this.range))},_documentClick:{mousedown:function(e){this.isOpen&&(t(e.target).closest(".ui-selectmenu-menu, #"+CSS.escape(this.ids.button)).length||this.close(e))}},_buttonEvents:{mousedown:function(){var t=window.getSelection();t.rangeCount&&(this.range=t.getRangeAt(0))},click:function(t){this._setSelection(),this._toggle(t)},keydown:function(e){var i=!0;switch(e.keyCode){case t.ui.keyCode.TAB:case t.ui.keyCode.ESCAPE:this.close(e),i=!1;break;case t.ui.keyCode.ENTER:this.isOpen&&this._selectFocusedItem(e);break;case t.ui.keyCode.UP:e.altKey?this._toggle(e):this._move("prev",e);break;case t.ui.keyCode.DOWN:e.altKey?this._toggle(e):this._move("next",e);break;case t.ui.keyCode.SPACE:this.isOpen?this._selectFocusedItem(e):this._toggle(e);break;case t.ui.keyCode.LEFT:this._move("prev",e);break;case t.ui.keyCode.RIGHT:this._move("next",e);break;case t.ui.keyCode.HOME:case t.ui.keyCode.PAGE_UP:this._move("first",e);break;case t.ui.keyCode.END:case t.ui.keyCode.PAGE_DOWN:this._move("last",e);break;default:this.menu.trigger(e),i=!1}i&&e.preventDefault()}},_selectFocusedItem:function(t){var e=this.menuItems.eq(this.focusIndex).parent("li");e.hasClass("ui-state-disabled")||this._select(e.data("ui-selectmenu-item"),t)},_select:function(t,e){var i=this.element[0].selectedIndex;this.element[0].selectedIndex=t.index,this.buttonItem.replaceWith(this.buttonItem=this._renderButtonItem(t)),this._setAria(t),this._trigger("select",e,{item:t}),t.index!==i&&this._trigger("change",e,{item:t}),this.close(e)},_setAria:function(t){var e=this.menuItems.eq(t.index).attr("id");this.button.attr({"aria-labelledby":e,"aria-activedescendant":e}),this.menu.attr("aria-activedescendant",e)},_setOption:function(t,e){if("icons"===t){var i=this.button.find("span.ui-icon");this._removeClass(i,null,this.options.icons.button)._addClass(i,null,e.button)}this._super(t,e),"appendTo"===t&&this.menuWrap.appendTo(this._appendTo()),"width"===t&&this._resizeButton()},_setOptionDisabled:function(t){this._super(t),this.menuInstance.option("disabled",t),this.button.attr("aria-disabled",t),this._toggleClass(this.button,null,"ui-state-disabled",t),this.element.prop("disabled",t),t?(this.button.attr("tabindex",-1),this.close()):this.button.attr("tabindex",0)},_appendTo:function(){var e=this.options.appendTo;return e&&(e=e.jquery||e.nodeType?t(e):this.document.find(e).eq(0)),e&&e[0]||(e=this.element.closest(".ui-front, dialog")),e.length||(e=this.document[0].body),e},_toggleAttr:function(){this.button.attr("aria-expanded",this.isOpen),this._removeClass(this.button,"ui-selectmenu-button-"+(this.isOpen?"closed":"open"))._addClass(this.button,"ui-selectmenu-button-"+(this.isOpen?"open":"closed"))._toggleClass(this.menuWrap,"ui-selectmenu-open",null,this.isOpen),this.menu.attr("aria-hidden",!this.isOpen)},_resizeButton:function(){var t=this.options.width;if(!1===t){this.button.css("width","");return}null===t&&(t=this.element.show().outerWidth(),this.element.hide()),this.button.outerWidth(t)},_resizeMenu:function(){this.menu.outerWidth(Math.max(this.button.outerWidth(),this.menu.width("").outerWidth()))},_getCreateOptions:function(){var t=this._super();return t.disabled=this.element.prop("disabled"),t},_parseOptions:function(e){var i=this,s=[];e.each(function(e,n){s.push(i._parseOption(t(n),e))}),this.items=s},_parseOption:function(t,e){var i=t.parent("optgroup");return{element:t,index:e,value:t.val(),label:t.text(),hidden:i.prop("hidden")||t.prop("hidden"),optgroup:i.attr("label")||"",disabled:i.prop("disabled")||t.prop("disabled")}},_destroy:function(){this._unbindFormResetHandler(),this.menuWrap.remove(),this.button.remove(),this.element.show(),this.element.removeUniqueId(),this.labels.attr("for",this.ids.element)}}]),t.widget("ui.slider",t.ui.mouse,{version:"1.14.2",widgetEventPrefix:"slide",options:{animate:!1,classes:{"ui-slider":"ui-corner-all","ui-slider-handle":"ui-corner-all","ui-slider-range":"ui-corner-all ui-widget-header"},distance:0,max:100,min:0,orientation:"horizontal",range:!1,step:1,value:0,values:null,change:null,slide:null,start:null,stop:null},numPages:5,_create:function(){this._keySliding=!1,this._mouseSliding=!1,this._animateOff=!0,this._handleIndex=null,this._detectOrientation(),this._mouseInit(),this._calculateNewMax(),this._addClass("ui-slider ui-slider-"+this.orientation,"ui-widget ui-widget-content"),this._refresh(),this._animateOff=!1},_refresh:function(){this._createRange(),this._createHandles(),this._setupEvents(),this._refreshValue()},_createHandles:function(){var e,i,s=this.options,n=this.element.find(".ui-slider-handle"),o=[];for(i=s.values&&s.values.length||1,n.length>i&&(n.slice(i).remove(),n=n.slice(0,i)),e=n.length;e<i;e++)o.push("<span tabindex='0'></span>");this.handles=n.add(t(o.join("")).appendTo(this.element)),this._addClass(this.handles,"ui-slider-handle","ui-state-default"),this.handle=this.handles.eq(0),this.handles.each(function(e){t(this).data("ui-slider-handle-index",e).attr("tabIndex",0)})},_createRange:function(){var e=this.options;e.range?(!0===e.range&&(e.values?e.values.length&&2!==e.values.length?e.values=[e.values[0],e.values[0]]:Array.isArray(e.values)&&(e.values=e.values.slice(0)):e.values=[this._valueMin(),this._valueMin()]),this.range&&this.range.length?(this._removeClass(this.range,"ui-slider-range-min ui-slider-range-max"),this.range.css({left:"",bottom:""})):(this.range=t("<div>").appendTo(this.element),this._addClass(this.range,"ui-slider-range")),("min"===e.range||"max"===e.range)&&this._addClass(this.range,"ui-slider-range-"+e.range)):(this.range&&this.range.remove(),this.range=null)},_setupEvents:function(){this._off(this.handles),this._on(this.handles,this._handleEvents),this._hoverable(this.handles),this._focusable(this.handles)},_destroy:function(){this.handles.remove(),this.range&&this.range.remove(),this._mouseDestroy()},_mouseCapture:function(e){var i,s,n,o,a,r,l,h=this,c=this.options;return!c.disabled&&(this.elementSize={width:this.element.outerWidth(),height:this.element.outerHeight()},this.elementOffset=this.element.offset(),i={x:e.pageX,y:e.pageY},s=this._normValueFromMouse(i),n=this._valueMax()-this._valueMin()+1,this.handles.each(function(e){var i=Math.abs(s-h.values(e));(n>i||n===i&&(e===h._lastChangedValue||h.values(e)===c.min))&&(n=i,o=t(this),a=e)}),!1!==this._start(e,a)&&(this._mouseSliding=!0,this._handleIndex=a,this._addClass(o,null,"ui-state-active"),o.trigger("focus"),r=o.offset(),l=!t(e.target).parents().addBack().is(".ui-slider-handle"),this._clickOffset=l?{left:0,top:0}:{left:e.pageX-r.left-o.width()/2,top:e.pageY-r.top-o.height()/2-(parseInt(o.css("borderTopWidth"),10)||0)-(parseInt(o.css("borderBottomWidth"),10)||0)+(parseInt(o.css("marginTop"),10)||0)},this.handles.hasClass("ui-state-hover")||this._slide(e,a,s),this._animateOff=!0,!0))},_mouseStart:function(){return!0},_mouseDrag:function(t){var e={x:t.pageX,y:t.pageY},i=this._normValueFromMouse(e);return this._slide(t,this._handleIndex,i),!1},_mouseStop:function(t){return this._removeClass(this.handles,null,"ui-state-active"),this._mouseSliding=!1,this._stop(t,this._handleIndex),this._change(t,this._handleIndex),this._handleIndex=null,this._clickOffset=null,this._animateOff=!1,!1},_detectOrientation:function(){this.orientation="vertical"===this.options.orientation?"vertical":"horizontal"},_normValueFromMouse:function(t){var e,i,s,n,o;return"horizontal"===this.orientation?(e=this.elementSize.width,i=t.x-this.elementOffset.left-(this._clickOffset?this._clickOffset.left:0)):(e=this.elementSize.height,i=t.y-this.elementOffset.top-(this._clickOffset?this._clickOffset.top:0)),(s=i/e)>1&&(s=1),s<0&&(s=0),"vertical"===this.orientation&&(s=1-s),n=this._valueMax()-this._valueMin(),o=this._valueMin()+s*n,this._trimAlignValue(o)},_uiHash:function(t,e,i){var s={handle:this.handles[t],handleIndex:t,value:void 0!==e?e:this.value()};return this._hasMultipleValues()&&(s.value=void 0!==e?e:this.values(t),s.values=i||this.values()),s},_hasMultipleValues:function(){return this.options.values&&this.options.values.length},_start:function(t,e){return this._trigger("start",t,this._uiHash(e))},_slide:function(t,e,i){var s,n=this.value(),o=this.values();this._hasMultipleValues()&&(s=this.values(+!e),n=this.values(e),2===this.options.values.length&&!0===this.options.range&&(i=0===e?Math.min(s,i):Math.max(s,i)),o[e]=i),i!==n&&!1!==this._trigger("slide",t,this._uiHash(e,i,o))&&(this._hasMultipleValues()?this.values(e,i):this.value(i))},_stop:function(t,e){this._trigger("stop",t,this._uiHash(e))},_change:function(t,e){this._keySliding||this._mouseSliding||(this._lastChangedValue=e,this._trigger("change",t,this._uiHash(e)))},value:function(t){if(arguments.length){this.options.value=this._trimAlignValue(t),this._refreshValue(),this._change(null,0);return}return this._value()},values:function(t,e){var i,s,n;if(arguments.length>1){this.options.values[t]=this._trimAlignValue(e),this._refreshValue(),this._change(null,t);return}if(!arguments.length)return this._values();if(!Array.isArray(arguments[0]))return this._hasMultipleValues()?this._values(t):this.value();for(n=0,i=this.options.values,s=arguments[0];n<i.length;n+=1)i[n]=this._trimAlignValue(s[n]),this._change(null,n);this._refreshValue()},_setOption:function(t,e){var i,s=0;switch("range"===t&&!0===this.options.range&&("min"===e?(this.options.value=this._values(0),this.options.values=null):"max"===e&&(this.options.value=this._values(this.options.values.length-1),this.options.values=null)),Array.isArray(this.options.values)&&(s=this.options.values.length),this._super(t,e),t){case"orientation":this._detectOrientation(),this._removeClass("ui-slider-horizontal ui-slider-vertical")._addClass("ui-slider-"+this.orientation),this._refreshValue(),this.options.range&&this._refreshRange(e),this.handles.css("horizontal"===e?"bottom":"left","");break;case"value":this._animateOff=!0,this._refreshValue(),this._change(null,0),this._animateOff=!1;break;case"values":for(this._animateOff=!0,this._refreshValue(),i=s-1;i>=0;i--)this._change(null,i);this._animateOff=!1;break;case"step":case"min":case"max":this._animateOff=!0,this._calculateNewMax(),this._refreshValue(),this._animateOff=!1;break;case"range":this._animateOff=!0,this._refresh(),this._animateOff=!1}},_setOptionDisabled:function(t){this._super(t),this._toggleClass(null,"ui-state-disabled",!!t)},_value:function(){var t=this.options.value;return this._trimAlignValue(t)},_values:function(t){var e,i,s;if(arguments.length)return e=this.options.values[t],e=this._trimAlignValue(e);if(!this._hasMultipleValues())return[];for(s=0,i=this.options.values.slice();s<i.length;s+=1)i[s]=this._trimAlignValue(i[s]);return i},_trimAlignValue:function(t){if(t<=this._valueMin())return this._valueMin();if(t>=this._valueMax())return this._valueMax();var e=this.options.step>0?this.options.step:1,i=(t-this._valueMin())%e,s=t-i;return 2*Math.abs(i)>=e&&(s+=i>0?e:-e),parseFloat(s.toFixed(5))},_calculateNewMax:function(){var t=this.options.max,e=this._valueMin(),i=this.options.step;(t=Math.round((t-e)/i)*i+e)>this.options.max&&(t-=i),this.max=parseFloat(t.toFixed(this._precision()))},_precision:function(){var t=this._precisionOf(this.options.step);return null!==this.options.min&&(t=Math.max(t,this._precisionOf(this.options.min))),t},_precisionOf:function(t){var e=t.toString(),i=e.indexOf(".");return -1===i?0:e.length-i-1},_valueMin:function(){return this.options.min},_valueMax:function(){return this.max},_refreshRange:function(t){"vertical"===t&&this.range.css({width:"",left:""}),"horizontal"===t&&this.range.css({height:"",bottom:""})},_refreshValue:function(){var e,i,s,n,o,a=this.options.range,r=this.options,l=this,h=!this._animateOff&&r.animate,c={};this._hasMultipleValues()?this.handles.each(function(s){i=(l.values(s)-l._valueMin())/(l._valueMax()-l._valueMin())*100,c["horizontal"===l.orientation?"left":"bottom"]=i+"%",t(this).stop(1,1)[h?"animate":"css"](c,r.animate),!0===l.options.range&&("horizontal"===l.orientation?(0===s&&l.range.stop(1,1)[h?"animate":"css"]({left:i+"%"},r.animate),1===s&&l.range[h?"animate":"css"]({width:i-e+"%"},{queue:!1,duration:r.animate})):(0===s&&l.range.stop(1,1)[h?"animate":"css"]({bottom:i+"%"},r.animate),1===s&&l.range[h?"animate":"css"]({height:i-e+"%"},{queue:!1,duration:r.animate}))),e=i}):(s=this.value(),n=this._valueMin(),i=(o=this._valueMax())!==n?(s-n)/(o-n)*100:0,c["horizontal"===this.orientation?"left":"bottom"]=i+"%",this.handle.stop(1,1)[h?"animate":"css"](c,r.animate),"min"===a&&"horizontal"===this.orientation&&this.range.stop(1,1)[h?"animate":"css"]({width:i+"%"},r.animate),"max"===a&&"horizontal"===this.orientation&&this.range.stop(1,1)[h?"animate":"css"]({width:100-i+"%"},r.animate),"min"===a&&"vertical"===this.orientation&&this.range.stop(1,1)[h?"animate":"css"]({height:i+"%"},r.animate),"max"===a&&"vertical"===this.orientation&&this.range.stop(1,1)[h?"animate":"css"]({height:100-i+"%"},r.animate))},_handleEvents:{keydown:function(e){var i,s,n,o=t(e.target).data("ui-slider-handle-index");switch(e.keyCode){case t.ui.keyCode.HOME:case t.ui.keyCode.END:case t.ui.keyCode.PAGE_UP:case t.ui.keyCode.PAGE_DOWN:case t.ui.keyCode.UP:case t.ui.keyCode.RIGHT:case t.ui.keyCode.DOWN:case t.ui.keyCode.LEFT:if(e.preventDefault(),!this._keySliding&&(this._keySliding=!0,this._addClass(t(e.target),null,"ui-state-active"),!1===this._start(e,o)))return}switch(n=this.options.step,i=s=this._hasMultipleValues()?this.values(o):this.value(),e.keyCode){case t.ui.keyCode.HOME:s=this._valueMin();break;case t.ui.keyCode.END:s=this._valueMax();break;case t.ui.keyCode.PAGE_UP:s=this._trimAlignValue(i+(this._valueMax()-this._valueMin())/this.numPages);break;case t.ui.keyCode.PAGE_DOWN:s=this._trimAlignValue(i-(this._valueMax()-this._valueMin())/this.numPages);break;case t.ui.keyCode.UP:case t.ui.keyCode.RIGHT:if(i===this._valueMax())return;s=this._trimAlignValue(i+n);break;case t.ui.keyCode.DOWN:case t.ui.keyCode.LEFT:if(i===this._valueMin())return;s=this._trimAlignValue(i-n)}this._slide(e,o,s)},keyup:function(e){var i=t(e.target).data("ui-slider-handle-index");this._keySliding&&(this._keySliding=!1,this._stop(e,i),this._change(e,i),this._removeClass(t(e.target),null,"ui-state-active"))}}}),t.widget("ui.sortable",t.ui.mouse,{version:"1.14.2",widgetEventPrefix:"sort",ready:!1,options:{appendTo:"parent",axis:!1,connectWith:!1,containment:!1,cursor:"auto",cursorAt:!1,dropOnEmpty:!0,forcePlaceholderSize:!1,forceHelperSize:!1,grid:!1,handle:!1,helper:"original",items:"> *",opacity:!1,placeholder:!1,revert:!1,scroll:!0,scrollSensitivity:20,scrollSpeed:20,scope:"default",tolerance:"intersect",zIndex:1e3,activate:null,beforeStop:null,change:null,deactivate:null,out:null,over:null,receive:null,remove:null,sort:null,start:null,stop:null,update:null},_isOverAxis:function(t,e,i){return t>=e&&t<e+i},_isFloating:function(t){return/left|right/.test(t.css("float"))||/inline|table-cell/.test(t.css("display"))},_create:function(){this.containerCache={},this._addClass("ui-sortable"),this.refresh(),this.offset=this.element.offset(),this._mouseInit(),this._setHandleClassName(),this.ready=!0},_setOption:function(t,e){this._super(t,e),"handle"===t&&this._setHandleClassName()},_setHandleClassName:function(){var e=this;this._removeClass(this.element.find(".ui-sortable-handle"),"ui-sortable-handle"),t.each(this.items,function(){e._addClass(this.instance.options.handle?this.item.find(this.instance.options.handle):this.item,"ui-sortable-handle")})},_destroy:function(){this._mouseDestroy();for(var t=this.items.length-1;t>=0;t--)this.items[t].item.removeData(this.widgetName+"-item");return this},_mouseCapture:function(e,i){var s=null,n=!1,o=this;return!this.reverting&&!this.options.disabled&&"static"!==this.options.type&&(this._refreshItems(e),t(e.target).parents().each(function(){if(t.data(this,o.widgetName+"-item")===o)return s=t(this),!1}),t.data(e.target,o.widgetName+"-item")===o&&(s=t(e.target)),!!s&&(!this.options.handle||!!i||(t(this.options.handle,s).find("*").addBack().each(function(){this===e.target&&(n=!0)}),!!n))&&(this.currentItem=s,this._removeCurrentsFromItems(),!0))},_mouseStart:function(e,i,s){var n,o,a=this.options;if(this.currentContainer=this,this.refreshPositions(),this.appendTo=t("parent"!==a.appendTo?a.appendTo:this.currentItem.parent()),this.helper=this._createHelper(e),this._cacheHelperProportions(),this._cacheMargins(),this.offset=this.currentItem.offset(),this.offset={top:this.offset.top-this.margins.top,left:this.offset.left-this.margins.left},t.extend(this.offset,{click:{left:e.pageX-this.offset.left,top:e.pageY-this.offset.top},relative:this._getRelativeOffset()}),this.helper.css("position","absolute"),this.cssPosition=this.helper.css("position"),a.cursorAt&&this._adjustOffsetFromHelper(a.cursorAt),this.domPosition={prev:this.currentItem.prev()[0],parent:this.currentItem.parent()[0]},this.helper[0]!==this.currentItem[0]&&this.currentItem.hide(),this._createPlaceholder(),this.scrollParent=this.placeholder.scrollParent(),t.extend(this.offset,{parent:this._getParentOffset()}),a.containment&&this._setContainment(),a.cursor&&"auto"!==a.cursor&&(o=this.document.find("body"),this._storedStylesheet=t("<style>*{ cursor: "+a.cursor+" !important; }</style>").appendTo(o)),a.zIndex&&(this.helper.css("zIndex")&&(this._storedZIndex=this.helper.css("zIndex")),this.helper.css("zIndex",a.zIndex)),a.opacity&&(this.helper.css("opacity")&&(this._storedOpacity=this.helper.css("opacity")),this.helper.css("opacity",a.opacity)),this.scrollParent[0]!==this.document[0]&&"HTML"!==this.scrollParent[0].tagName&&(this.overflowOffset=this.scrollParent.offset()),this._trigger("start",e,this._uiHash()),this._preserveHelperProportions||this._cacheHelperProportions(),!s)for(n=this.containers.length-1;n>=0;n--)this.containers[n]._trigger("activate",e,this._uiHash(this));return t.ui.ddmanager&&(t.ui.ddmanager.current=this),t.ui.ddmanager&&!a.dropBehaviour&&t.ui.ddmanager.prepareOffsets(this,e),this.dragging=!0,this._addClass(this.helper,"ui-sortable-helper"),this.helper.parent().is(this.appendTo)||(this.helper.detach().appendTo(this.appendTo),this.offset.parent=this._getParentOffset()),this.position=this.originalPosition=this._generatePosition(e),this.originalPageX=e.pageX,this.originalPageY=e.pageY,this.lastPositionAbs=this.positionAbs=this._convertPositionTo("absolute"),this._mouseDrag(e),!0},_scroll:function(t){var e=this.options,i=!1;return this.scrollParent[0]!==this.document[0]&&"HTML"!==this.scrollParent[0].tagName?(this.overflowOffset.top+this.scrollParent[0].offsetHeight-t.pageY<e.scrollSensitivity?this.scrollParent[0].scrollTop=i=this.scrollParent[0].scrollTop+e.scrollSpeed:t.pageY-this.overflowOffset.top<e.scrollSensitivity&&(this.scrollParent[0].scrollTop=i=this.scrollParent[0].scrollTop-e.scrollSpeed),this.overflowOffset.left+this.scrollParent[0].offsetWidth-t.pageX<e.scrollSensitivity?this.scrollParent[0].scrollLeft=i=this.scrollParent[0].scrollLeft+e.scrollSpeed:t.pageX-this.overflowOffset.left<e.scrollSensitivity&&(this.scrollParent[0].scrollLeft=i=this.scrollParent[0].scrollLeft-e.scrollSpeed)):(t.pageY-this.document.scrollTop()<e.scrollSensitivity?i=this.document.scrollTop(this.document.scrollTop()-e.scrollSpeed):this.window.height()-(t.pageY-this.document.scrollTop())<e.scrollSensitivity&&(i=this.document.scrollTop(this.document.scrollTop()+e.scrollSpeed)),t.pageX-this.document.scrollLeft()<e.scrollSensitivity?i=this.document.scrollLeft(this.document.scrollLeft()-e.scrollSpeed):this.window.width()-(t.pageX-this.document.scrollLeft())<e.scrollSensitivity&&(i=this.document.scrollLeft(this.document.scrollLeft()+e.scrollSpeed))),i},_mouseDrag:function(e){var i,s,n,o,a=this.options;for(this.position=this._generatePosition(e),this.positionAbs=this._convertPositionTo("absolute"),this.options.axis&&"y"===this.options.axis||(this.helper[0].style.left=this.position.left+"px"),this.options.axis&&"x"===this.options.axis||(this.helper[0].style.top=this.position.top+"px"),a.scroll&&!1!==this._scroll(e)&&(this._refreshItemPositions(!0),t.ui.ddmanager&&!a.dropBehaviour&&t.ui.ddmanager.prepareOffsets(this,e)),this.dragDirection={vertical:this._getDragVerticalDirection(),horizontal:this._getDragHorizontalDirection()},i=this.items.length-1;i>=0;i--)if(n=(s=this.items[i]).item[0],(o=this._intersectsWithPointer(s))&&s.instance===this.currentContainer&&n!==this.currentItem[0]&&this.placeholder[1===o?"next":"prev"]()[0]!==n&&!t.contains(this.placeholder[0],n)&&("semi-dynamic"!==this.options.type||!t.contains(this.element[0],n))){if(this.direction=1===o?"down":"up","pointer"===this.options.tolerance||this._intersectsWithSides(s))this._rearrange(e,s);else break;this._trigger("change",e,this._uiHash());break}return this._contactContainers(e),t.ui.ddmanager&&t.ui.ddmanager.drag(this,e),this._trigger("sort",e,this._uiHash()),this.lastPositionAbs=this.positionAbs,!1},_mouseStop:function(e,i){if(e){if(t.ui.ddmanager&&!this.options.dropBehaviour&&t.ui.ddmanager.drop(this,e),this.options.revert){var s=this,n=this.placeholder.offset(),o=this.options.axis,a={};o&&"x"!==o||(a.left=n.left-this.offset.parent.left-this.margins.left+(this.offsetParent[0]===this.document[0].body?0:this.offsetParent[0].scrollLeft)),o&&"y"!==o||(a.top=n.top-this.offset.parent.top-this.margins.top+(this.offsetParent[0]===this.document[0].body?0:this.offsetParent[0].scrollTop)),this.reverting=!0,t(this.helper).animate(a,parseInt(this.options.revert,10)||500,function(){s._clear(e)})}else this._clear(e,i);return!1}},cancel:function(){if(this.dragging){this._mouseUp(new t.Event("mouseup",{target:null})),"original"===this.options.helper?(this.currentItem.css(this._storedCSS),this._removeClass(this.currentItem,"ui-sortable-helper")):this.currentItem.show();for(var e=this.containers.length-1;e>=0;e--)this.containers[e]._trigger("deactivate",null,this._uiHash(this)),this.containers[e].containerCache.over&&(this.containers[e]._trigger("out",null,this._uiHash(this)),this.containers[e].containerCache.over=0)}return this.placeholder&&(this.placeholder[0].parentNode&&this.placeholder[0].parentNode.removeChild(this.placeholder[0]),"original"!==this.options.helper&&this.helper&&this.helper[0].parentNode&&this.helper.remove(),t.extend(this,{helper:null,dragging:!1,reverting:!1,_noFinalSort:null}),this.domPosition.prev?t(this.domPosition.prev).after(this.currentItem):t(this.domPosition.parent).prepend(this.currentItem)),this},serialize:function(e){var i=this._getItemsAsjQuery(e&&e.connected),s=[];return e=e||{},t(i).each(function(){var i=(t(e.item||this).attr(e.attribute||"id")||"").match(e.expression||/(.+)[\-=_](.+)/);i&&s.push((e.key||i[1]+"[]")+"="+(e.key&&e.expression?i[1]:i[2]))}),!s.length&&e.key&&s.push(e.key+"="),s.join("&")},toArray:function(e){var i=this._getItemsAsjQuery(e&&e.connected),s=[];return e=e||{},i.each(function(){s.push(t(e.item||this).attr(e.attribute||"id")||"")}),s},_intersectsWith:function(t){var e=this.positionAbs.left,i=e+this.helperProportions.width,s=this.positionAbs.top,n=s+this.helperProportions.height,o=t.left,a=o+t.width,r=t.top,l=r+t.height,h=this.offset.click.top,c=this.offset.click.left,u="x"===this.options.axis||s+h>r&&s+h<l,d="y"===this.options.axis||e+c>o&&e+c<a;return"pointer"===this.options.tolerance||this.options.forcePointerForContainers||"pointer"!==this.options.tolerance&&this.helperProportions[this.floating?"width":"height"]>t[this.floating?"width":"height"]?u&&d:o<e+this.helperProportions.width/2&&i-this.helperProportions.width/2<a&&r<s+this.helperProportions.height/2&&n-this.helperProportions.height/2<l},_intersectsWithPointer:function(t){var e,i,s="x"===this.options.axis||this._isOverAxis(this.positionAbs.top+this.offset.click.top,t.top,t.height),n="y"===this.options.axis||this._isOverAxis(this.positionAbs.left+this.offset.click.left,t.left,t.width);return!!s&&!!n&&(e=this.dragDirection.vertical,i=this.dragDirection.horizontal,this.floating?"right"===i||"down"===e?2:1:e&&("down"===e?2:1))},_intersectsWithSides:function(t){var e=this._isOverAxis(this.positionAbs.top+this.offset.click.top,t.top+t.height/2,t.height),i=this._isOverAxis(this.positionAbs.left+this.offset.click.left,t.left+t.width/2,t.width),s=this.dragDirection.vertical,n=this.dragDirection.horizontal;return this.floating&&n?"right"===n&&i||"left"===n&&!i:s&&("down"===s&&e||"up"===s&&!e)},_getDragVerticalDirection:function(){var t=this.positionAbs.top-this.lastPositionAbs.top;return 0!==t&&(t>0?"down":"up")},_getDragHorizontalDirection:function(){var t=this.positionAbs.left-this.lastPositionAbs.left;return 0!==t&&(t>0?"right":"left")},refresh:function(t){return this._refreshItems(t),this._setHandleClassName(),this.refreshPositions(),this},_connectWith:function(){var t=this.options;return t.connectWith.constructor===String?[t.connectWith]:t.connectWith},_getItemsAsjQuery:function(e){var i,s,n,o,a=[],r=[],l=this._connectWith();if(l&&e)for(i=l.length-1;i>=0;i--)for(s=(n=t(l[i],this.document[0])).length-1;s>=0;s--)(o=t.data(n[s],this.widgetFullName))&&o!==this&&!o.options.disabled&&r.push(["function"==typeof o.options.items?o.options.items.call(o.element):t(o.options.items,o.element).not(".ui-sortable-helper").not(".ui-sortable-placeholder"),o]);function h(){a.push(this)}for(r.push(["function"==typeof this.options.items?this.options.items.call(this.element,null,{options:this.options,item:this.currentItem}):t(this.options.items,this.element).not(".ui-sortable-helper").not(".ui-sortable-placeholder"),this]),i=r.length-1;i>=0;i--)r[i][0].each(h);return t(a)},_removeCurrentsFromItems:function(){var e=this.currentItem.find(":data("+this.widgetName+"-item)");this.items=t.grep(this.items,function(t){for(var i=0;i<e.length;i++)if(e[i]===t.item[0])return!1;return!0})},_refreshItems:function(e){this.items=[],this.containers=[this];var i,s,n,o,a,r,l,h,c=this.items,u=[["function"==typeof this.options.items?this.options.items.call(this.element[0],e,{item:this.currentItem}):t(this.options.items,this.element),this]],d=this._connectWith();if(d&&this.ready)for(i=d.length-1;i>=0;i--)for(s=(n=t(d[i],this.document[0])).length-1;s>=0;s--)(o=t.data(n[s],this.widgetFullName))&&o!==this&&!o.options.disabled&&(u.push(["function"==typeof o.options.items?o.options.items.call(o.element[0],e,{item:this.currentItem}):t(o.options.items,o.element),o]),this.containers.push(o));for(i=u.length-1;i>=0;i--)for(s=0,a=u[i][1],h=(r=u[i][0]).length;s<h;s++)(l=t(r[s])).data(this.widgetName+"-item",a),c.push({item:l,instance:a,width:0,height:0,left:0,top:0})},_refreshItemPositions:function(e){var i,s,n,o;for(i=this.items.length-1;i>=0;i--)s=this.items[i],(!this.currentContainer||s.instance===this.currentContainer||s.item[0]===this.currentItem[0])&&(n=this.options.toleranceElement?t(this.options.toleranceElement,s.item):s.item,e||(s.width=n.outerWidth(),s.height=n.outerHeight()),s.left=(o=n.offset()).left,s.top=o.top)},refreshPositions:function(t){var e,i;if(this.floating=!!this.items.length&&("x"===this.options.axis||this._isFloating(this.items[0].item)),this.offsetParent&&this.helper&&(this.offset.parent=this._getParentOffset()),this._refreshItemPositions(t),this.options.custom&&this.options.custom.refreshContainers)this.options.custom.refreshContainers.call(this);else for(e=this.containers.length-1;e>=0;e--)i=this.containers[e].element.offset(),this.containers[e].containerCache.left=i.left,this.containers[e].containerCache.top=i.top,this.containers[e].containerCache.width=this.containers[e].element.outerWidth(),this.containers[e].containerCache.height=this.containers[e].element.outerHeight();return this},_createPlaceholder:function(e){var i,s,n=(e=e||this).options;n.placeholder&&n.placeholder.constructor!==String||(i=n.placeholder,s=e.currentItem[0].nodeName.toLowerCase(),n.placeholder={element:function(){var n=t("<"+s+">",e.document[0]);return e._addClass(n,"ui-sortable-placeholder",i||e.currentItem[0].className)._removeClass(n,"ui-sortable-helper"),"tbody"===s?e._createTrPlaceholder(e.currentItem.find("tr").eq(0),t("<tr>",e.document[0]).appendTo(n)):"tr"===s?e._createTrPlaceholder(e.currentItem,n):"img"===s&&n.attr("src",e.currentItem.attr("src")),i||n.css("visibility","hidden"),n},update:function(t,o){(!i||n.forcePlaceholderSize)&&((!o.height()||n.forcePlaceholderSize&&("tbody"===s||"tr"===s))&&o.height(e.currentItem.innerHeight()-parseInt(e.currentItem.css("paddingTop")||0,10)-parseInt(e.currentItem.css("paddingBottom")||0,10)),o.width()||o.width(e.currentItem.innerWidth()-parseInt(e.currentItem.css("paddingLeft")||0,10)-parseInt(e.currentItem.css("paddingRight")||0,10)))}}),e.placeholder=t(n.placeholder.element.call(e.element,e.currentItem)),e.currentItem.after(e.placeholder),n.placeholder.update(e,e.placeholder)},_createTrPlaceholder:function(e,i){var s=this;e.children().each(function(){t("<td>&#160;</td>",s.document[0]).attr("colspan",t(this).attr("colspan")||1).appendTo(i)})},_contactContainers:function(e){var i,s,n,o,a,r,l,h,c,u,d=null,p=null;for(i=this.containers.length-1;i>=0;i--)if(!t.contains(this.currentItem[0],this.containers[i].element[0])){if(this._intersectsWith(this.containers[i].containerCache)){if(d&&t.contains(this.containers[i].element[0],d.element[0]))continue;d=this.containers[i],p=i}else this.containers[i].containerCache.over&&(this.containers[i]._trigger("out",e,this._uiHash(this)),this.containers[i].containerCache.over=0)}if(d){if(1===this.containers.length)this.containers[p].containerCache.over||(this.containers[p]._trigger("over",e,this._uiHash(this)),this.containers[p].containerCache.over=1);else{for(n=1e4,o=null,a=(c=d.floating||this._isFloating(this.currentItem))?"left":"top",r=c?"width":"height",u=c?"pageX":"pageY",s=this.items.length-1;s>=0;s--)t.contains(this.containers[p].element[0],this.items[s].item[0])&&this.items[s].item[0]!==this.currentItem[0]&&(l=this.items[s].item.offset()[a],h=!1,e[u]-l>this.items[s][r]/2&&(h=!0),Math.abs(e[u]-l)<n&&(n=Math.abs(e[u]-l),o=this.items[s],this.direction=h?"up":"down"));if(!o&&!this.options.dropOnEmpty)return;if(this.currentContainer===this.containers[p]){this.currentContainer.containerCache.over||(this.containers[p]._trigger("over",e,this._uiHash()),this.currentContainer.containerCache.over=1);return}o?this._rearrange(e,o,null,!0):this._rearrange(e,null,this.containers[p].element,!0),this._trigger("change",e,this._uiHash()),this.containers[p]._trigger("change",e,this._uiHash(this)),this.currentContainer=this.containers[p],this.options.placeholder.update(this.currentContainer,this.placeholder),this.scrollParent=this.placeholder.scrollParent(),this.scrollParent[0]!==this.document[0]&&"HTML"!==this.scrollParent[0].tagName&&(this.overflowOffset=this.scrollParent.offset()),this.containers[p]._trigger("over",e,this._uiHash(this)),this.containers[p].containerCache.over=1}}},_createHelper:function(e){var i=this.options,s="function"==typeof i.helper?t(i.helper.apply(this.element[0],[e,this.currentItem])):"clone"===i.helper?this.currentItem.clone():this.currentItem;return s.parents("body").length||this.appendTo[0].appendChild(s[0]),s[0]===this.currentItem[0]&&(this._storedCSS={width:this.currentItem[0].style.width,height:this.currentItem[0].style.height,position:this.currentItem.css("position"),top:this.currentItem.css("top"),left:this.currentItem.css("left")}),(!s[0].style.width||i.forceHelperSize)&&s.width(this.currentItem.width()),(!s[0].style.height||i.forceHelperSize)&&s.height(this.currentItem.height()),s},_adjustOffsetFromHelper:function(t){"string"==typeof t&&(t=t.split(" ")),Array.isArray(t)&&(t={left:+t[0],top:+t[1]||0}),"left"in t&&(this.offset.click.left=t.left+this.margins.left),"right"in t&&(this.offset.click.left=this.helperProportions.width-t.right+this.margins.left),"top"in t&&(this.offset.click.top=t.top+this.margins.top),"bottom"in t&&(this.offset.click.top=this.helperProportions.height-t.bottom+this.margins.top)},_getParentOffset:function(){this.offsetParent=this.helper.offsetParent();var e=this.offsetParent.offset();return"absolute"===this.cssPosition&&this.scrollParent[0]!==this.document[0]&&t.contains(this.scrollParent[0],this.offsetParent[0])&&(e.left+=this.scrollParent.scrollLeft(),e.top+=this.scrollParent.scrollTop()),this.offsetParent[0]===this.document[0].body&&(e={top:0,left:0}),{top:e.top+(parseInt(this.offsetParent.css("borderTopWidth"),10)||0),left:e.left+(parseInt(this.offsetParent.css("borderLeftWidth"),10)||0)}},_getRelativeOffset:function(){if("relative"!==this.cssPosition)return{top:0,left:0};var t=this.currentItem.position();return{top:t.top-(parseInt(this.helper.css("top"),10)||0)+this.scrollParent.scrollTop(),left:t.left-(parseInt(this.helper.css("left"),10)||0)+this.scrollParent.scrollLeft()}},_cacheMargins:function(){this.margins={left:parseInt(this.currentItem.css("marginLeft"),10)||0,top:parseInt(this.currentItem.css("marginTop"),10)||0}},_cacheHelperProportions:function(){this.helperProportions={width:this.helper.outerWidth(),height:this.helper.outerHeight()}},_setContainment:function(){var e,i,s,n=this.options;"parent"===n.containment&&(n.containment=this.helper[0].parentNode),("document"===n.containment||"window"===n.containment)&&(this.containment=[0-this.offset.relative.left-this.offset.parent.left,0-this.offset.relative.top-this.offset.parent.top,"document"===n.containment?this.document.width():this.window.width()-this.helperProportions.width-this.margins.left,("document"===n.containment?this.document.height()||document.body.parentNode.scrollHeight:this.window.height()||this.document[0].body.parentNode.scrollHeight)-this.helperProportions.height-this.margins.top]),/^(document|window|parent)$/.test(n.containment)||(e=t(n.containment)[0],i=t(n.containment).offset(),s="hidden"!==t(e).css("overflow"),this.containment=[i.left+(parseInt(t(e).css("borderLeftWidth"),10)||0)+(parseInt(t(e).css("paddingLeft"),10)||0)-this.margins.left,i.top+(parseInt(t(e).css("borderTopWidth"),10)||0)+(parseInt(t(e).css("paddingTop"),10)||0)-this.margins.top,i.left+(s?Math.max(e.scrollWidth,e.offsetWidth):e.offsetWidth)-(parseInt(t(e).css("borderLeftWidth"),10)||0)-(parseInt(t(e).css("paddingRight"),10)||0)-this.helperProportions.width-this.margins.left,i.top+(s?Math.max(e.scrollHeight,e.offsetHeight):e.offsetHeight)-(parseInt(t(e).css("borderTopWidth"),10)||0)-(parseInt(t(e).css("paddingBottom"),10)||0)-this.helperProportions.height-this.margins.top])},_convertPositionTo:function(e,i){i||(i=this.position);var s="absolute"===e?1:-1,n="absolute"!==this.cssPosition||this.scrollParent[0]!==this.document[0]&&t.contains(this.scrollParent[0],this.offsetParent[0])?this.scrollParent:this.offsetParent,o=/(html|body)/i.test(n[0].tagName);return{top:i.top+this.offset.relative.top*s+this.offset.parent.top*s-("fixed"===this.cssPosition?-this.scrollParent.scrollTop():o?0:n.scrollTop())*s,left:i.left+this.offset.relative.left*s+this.offset.parent.left*s-("fixed"===this.cssPosition?-this.scrollParent.scrollLeft():o?0:n.scrollLeft())*s}},_generatePosition:function(e){var i,s,n=this.options,o=e.pageX,a=e.pageY,r="absolute"!==this.cssPosition||this.scrollParent[0]!==this.document[0]&&t.contains(this.scrollParent[0],this.offsetParent[0])?this.scrollParent:this.offsetParent,l=/(html|body)/i.test(r[0].tagName);return"relative"===this.cssPosition&&(this.scrollParent[0]===this.document[0]||this.scrollParent[0]===this.offsetParent[0])&&(this.offset.relative=this._getRelativeOffset()),this.originalPosition&&(this.containment&&(e.pageX-this.offset.click.left<this.containment[0]&&(o=this.containment[0]+this.offset.click.left),e.pageY-this.offset.click.top<this.containment[1]&&(a=this.containment[1]+this.offset.click.top),e.pageX-this.offset.click.left>this.containment[2]&&(o=this.containment[2]+this.offset.click.left),e.pageY-this.offset.click.top>this.containment[3]&&(a=this.containment[3]+this.offset.click.top)),n.grid&&(i=this.originalPageY+Math.round((a-this.originalPageY)/n.grid[1])*n.grid[1],a=this.containment?i-this.offset.click.top>=this.containment[1]&&i-this.offset.click.top<=this.containment[3]?i:i-this.offset.click.top>=this.containment[1]?i-n.grid[1]:i+n.grid[1]:i,s=this.originalPageX+Math.round((o-this.originalPageX)/n.grid[0])*n.grid[0],o=this.containment?s-this.offset.click.left>=this.containment[0]&&s-this.offset.click.left<=this.containment[2]?s:s-this.offset.click.left>=this.containment[0]?s-n.grid[0]:s+n.grid[0]:s)),{top:a-this.offset.click.top-this.offset.relative.top-this.offset.parent.top+("fixed"===this.cssPosition?-this.scrollParent.scrollTop():l?0:r.scrollTop()),left:o-this.offset.click.left-this.offset.relative.left-this.offset.parent.left+("fixed"===this.cssPosition?-this.scrollParent.scrollLeft():l?0:r.scrollLeft())}},_rearrange:function(t,e,i,s){i?i[0].appendChild(this.placeholder[0]):e.item[0].parentNode.insertBefore(this.placeholder[0],"down"===this.direction?e.item[0]:e.item[0].nextSibling),this.counter=this.counter?++this.counter:1;var n=this.counter;this._delay(function(){n===this.counter&&this.refreshPositions(!s)})},_clear:function(t,e){this.reverting=!1;var i,s=[];if(!this._noFinalSort&&this.currentItem.parent().length&&this.placeholder.before(this.currentItem),this._noFinalSort=null,this.helper[0]===this.currentItem[0]){for(i in this._storedCSS)("auto"===this._storedCSS[i]||"static"===this._storedCSS[i])&&(this._storedCSS[i]="");this.currentItem.css(this._storedCSS),this._removeClass(this.currentItem,"ui-sortable-helper")}else this.currentItem.show();function n(t,e,i){return function(s){i._trigger(t,s,e._uiHash(e))}}for(this.fromOutside&&!e&&s.push(function(t){this._trigger("receive",t,this._uiHash(this.fromOutside))}),(this.fromOutside||this.domPosition.prev!==this.currentItem.prev().not(".ui-sortable-helper")[0]||this.domPosition.parent!==this.currentItem.parent()[0])&&!e&&s.push(function(t){this._trigger("update",t,this._uiHash())}),this===this.currentContainer||e||(s.push(function(t){this._trigger("remove",t,this._uiHash())}),s.push((function(t){return function(e){t._trigger("receive",e,this._uiHash(this))}}).call(this,this.currentContainer)),s.push((function(t){return function(e){t._trigger("update",e,this._uiHash(this))}}).call(this,this.currentContainer))),i=this.containers.length-1;i>=0;i--)e||s.push(n("deactivate",this,this.containers[i])),this.containers[i].containerCache.over&&(s.push(n("out",this,this.containers[i])),this.containers[i].containerCache.over=0);if(this._storedStylesheet&&(this._storedStylesheet.remove(),this._storedStylesheet=null),this._storedOpacity&&this.helper.css("opacity",this._storedOpacity),this._storedZIndex&&this.helper.css("zIndex","auto"===this._storedZIndex?"":this._storedZIndex),this.dragging=!1,e||this._trigger("beforeStop",t,this._uiHash()),this.placeholder[0].parentNode.removeChild(this.placeholder[0]),this.cancelHelperRemoval||(this.helper[0]!==this.currentItem[0]&&this.helper.remove(),this.helper=null),!e){for(i=0;i<s.length;i++)s[i].call(this,t);this._trigger("stop",t,this._uiHash())}return this.fromOutside=!1,!this.cancelHelperRemoval},_trigger:function(){!1===t.Widget.prototype._trigger.apply(this,arguments)&&this.cancel()},_uiHash:function(e){var i=e||this;return{helper:i.helper,placeholder:i.placeholder||t([]),position:i.position,originalPosition:i.originalPosition,offset:i.positionAbs,item:i.currentItem,sender:e?e.element:null}}}),t.widget("ui.spinner",{version:"1.14.2",defaultElement:"<input>",widgetEventPrefix:"spin",options:{classes:{"ui-spinner":"ui-corner-all","ui-spinner-down":"ui-corner-br","ui-spinner-up":"ui-corner-tr"},culture:null,icons:{down:"ui-icon-triangle-1-s",up:"ui-icon-triangle-1-n"},incremental:!0,max:null,min:null,numberFormat:null,page:10,step:1,change:null,spin:null,start:null,stop:null},_create:function(){this._setOption("max",this.options.max),this._setOption("min",this.options.min),this._setOption("step",this.options.step),""!==this.value()&&this._value(this.element.val(),!0),this._draw(),this._on(this._events),this._refresh(),this._on(this.window,{beforeunload:function(){this.element.removeAttr("autocomplete")}})},_getCreateOptions:function(){var e=this._super(),i=this.element;return t.each(["min","max","step"],function(t,s){var n=i.attr(s);null!=n&&n.length&&(e[s]=n)}),e},_events:{keydown:function(t){this._start(t)&&this._keydown(t)&&t.preventDefault()},keyup:"_stop",focus:function(){this.previous=this.element.val()},blur:function(t){this._stop(),this._refresh(),this.previous!==this.element.val()&&this._trigger("change",t)},wheel:function(t){var e=this.document[0].activeElement,i=this.element[0]===e,s=t.deltaY||t.originalEvent&&t.originalEvent.deltaY;if(i&&s){if(!this.spinning&&!this._start(t))return!1;this._spin((s>0?-1:1)*this.options.step,t),clearTimeout(this.mousewheelTimer),this.mousewheelTimer=this._delay(function(){this.spinning&&this._stop(t)},100),t.preventDefault()}},mousewheel:function(e,i){if(e.isTrigger){var s=t.Event(e);return s.type="wheel",i&&(s.deltaY=-i),this._events.wheel.call(this,s)}},"mousedown .ui-spinner-button":function(e){var i;i=this.element[0]===this.document[0].activeElement?this.previous:this.element.val(),e.preventDefault(),(function(){this.element[0]!==this.document[0].activeElement&&(this.element.trigger("focus"),this.previous=i)}).call(this),!1!==this._start(e)&&this._repeat(null,t(e.currentTarget).hasClass("ui-spinner-up")?1:-1,e)},"mouseup .ui-spinner-button":"_stop","mouseenter .ui-spinner-button":function(e){if(t(e.currentTarget).hasClass("ui-state-active")){if(!1===this._start(e))return!1;this._repeat(null,t(e.currentTarget).hasClass("ui-spinner-up")?1:-1,e)}},"mouseleave .ui-spinner-button":"_stop"},_enhance:function(){this.uiSpinner=this.element.attr("autocomplete","off").wrap("<span>").parent().append("<a></a><a></a>")},_draw:function(){this._enhance(),this._addClass(this.uiSpinner,"ui-spinner","ui-widget ui-widget-content"),this._addClass("ui-spinner-input"),this.element.attr("role","spinbutton"),this.buttons=this.uiSpinner.children("a").attr("tabIndex",-1).attr("aria-hidden",!0).button({classes:{"ui-button":""}}),this._removeClass(this.buttons,"ui-corner-all"),this._addClass(this.buttons.first(),"ui-spinner-button ui-spinner-up"),this._addClass(this.buttons.last(),"ui-spinner-button ui-spinner-down"),this.buttons.first().button({icon:this.options.icons.up,showLabel:!1}),this.buttons.last().button({icon:this.options.icons.down,showLabel:!1}),this.buttons.height()>Math.ceil(.5*this.uiSpinner.height())&&this.uiSpinner.height()>0&&this.uiSpinner.height(this.uiSpinner.height())},_keydown:function(e){var i=this.options,s=t.ui.keyCode;switch(e.keyCode){case s.UP:return this._repeat(null,1,e),!0;case s.DOWN:return this._repeat(null,-1,e),!0;case s.PAGE_UP:return this._repeat(null,i.page,e),!0;case s.PAGE_DOWN:return this._repeat(null,-i.page,e),!0}return!1},_start:function(t){return(!!this.spinning||!1!==this._trigger("start",t))&&(this.counter||(this.counter=1),this.spinning=!0,!0)},_repeat:function(t,e,i){t=t||500,clearTimeout(this.timer),this.timer=this._delay(function(){this._repeat(40,e,i)},t),this._spin(e*this.options.step,i)},_spin:function(t,e){var i=this.value()||0;this.counter||(this.counter=1),i=this._adjustValue(i+t*this._increment(this.counter)),(!this.spinning||!1!==this._trigger("spin",e,{value:i}))&&(this._value(i),this.counter++)},_increment:function(t){var e=this.options.incremental;return e?"function"==typeof e?e(t):Math.floor(t*t*t/5e4-t*t/500+17*t/200+1):1},_precision:function(){var t=this._precisionOf(this.options.step);return null!==this.options.min&&(t=Math.max(t,this._precisionOf(this.options.min))),t},_precisionOf:function(t){var e=t.toString(),i=e.indexOf(".");return -1===i?0:e.length-i-1},_adjustValue:function(t){var e,i,s=this.options;return(i=Math.round((i=t-(e=null!==s.min?s.min:0))/s.step)*s.step,t=parseFloat((t=e+i).toFixed(this._precision())),null!==s.max&&t>s.max)?s.max:null!==s.min&&t<s.min?s.min:t},_stop:function(t){this.spinning&&(clearTimeout(this.timer),clearTimeout(this.mousewheelTimer),this.counter=0,this.spinning=!1,this._trigger("stop",t))},_setOption:function(t,e){var i,s,n;if("culture"===t||"numberFormat"===t){i=this._parse(this.element.val()),this.options[t]=e,this.element.val(this._format(i));return}("max"===t||"min"===t||"step"===t)&&"string"==typeof e&&(e=this._parse(e)),"icons"===t&&(s=this.buttons.first().find(".ui-icon"),this._removeClass(s,null,this.options.icons.up),this._addClass(s,null,e.up),n=this.buttons.last().find(".ui-icon"),this._removeClass(n,null,this.options.icons.down),this._addClass(n,null,e.down)),this._super(t,e)},_setOptionDisabled:function(t){this._super(t),this._toggleClass(this.uiSpinner,null,"ui-state-disabled",!!t),this.element.prop("disabled",!!t),this.buttons.button(t?"disable":"enable")},_setOptions:O(function(t){this._super(t)}),_parse:function(t){return"string"==typeof t&&""!==t&&(t=window.Globalize&&this.options.numberFormat?Globalize.parseFloat(t,10,this.options.culture):+t),""===t||isNaN(t)?null:t},_format:function(t){return""===t?"":window.Globalize&&this.options.numberFormat?Globalize.format(t,this.options.numberFormat,this.options.culture):t},_refresh:function(){this.element.attr({"aria-valuemin":this.options.min,"aria-valuemax":this.options.max,"aria-valuenow":this._parse(this.element.val())})},isValid:function(){var t=this.value();return null!==t&&t===this._adjustValue(t)},_value:function(t,e){var i;""!==t&&null!==(i=this._parse(t))&&(e||(i=this._adjustValue(i)),t=this._format(i)),this.element.val(t),this._refresh()},_destroy:function(){this.element.prop("disabled",!1).removeAttr("autocomplete role aria-valuemin aria-valuemax aria-valuenow"),this.uiSpinner.replaceWith(this.element)},stepUp:O(function(t){this._stepUp(t)}),_stepUp:function(t){this._start()&&(this._spin((t||1)*this.options.step),this._stop())},stepDown:O(function(t){this._stepDown(t)}),_stepDown:function(t){this._start()&&(this._spin(-((t||1)*this.options.step)),this._stop())},pageUp:O(function(t){this._stepUp((t||1)*this.options.page)}),pageDown:O(function(t){this._stepDown((t||1)*this.options.page)}),value:function(t){if(!arguments.length)return this._parse(this.element.val());O(this._value).call(this,t)},widget:function(){return this.uiSpinner}}),!0===t.uiBackCompat&&t.widget("ui.spinner",t.ui.spinner,{_enhance:function(){this.uiSpinner=this.element.attr("autocomplete","off").wrap(this._uiSpinnerHtml()).parent().append(this._buttonHtml())},_uiSpinnerHtml:function(){return"<span>"},_buttonHtml:function(){return"<a></a><a></a>"}}),t.ui.spinner,t.widget("ui.tabs",{version:"1.14.2",delay:300,options:{active:null,classes:{"ui-tabs":"ui-corner-all","ui-tabs-nav":"ui-corner-all","ui-tabs-panel":"ui-corner-bottom","ui-tabs-tab":"ui-corner-top"},collapsible:!1,event:"click",heightStyle:"content",hide:null,show:null,activate:null,beforeActivate:null,beforeLoad:null,load:null},_isLocal:function(t){var e=new URL(t.href),i=new URL(location.href);return t.hash.length>1&&e.origin===i.origin&&e.pathname===i.pathname&&e.search===i.search},_create:function(){var e=this,i=this.options;this.running=!1,this._addClass("ui-tabs","ui-widget ui-widget-content"),this._toggleClass("ui-tabs-collapsible",null,i.collapsible),this._processTabs(),i.active=this._initialActive(),Array.isArray(i.disabled)&&(i.disabled=t.uniqueSort(i.disabled.concat(t.map(this.tabs.filter(".ui-state-disabled"),function(t){return e.tabs.index(t)}))).sort()),!1!==this.options.active&&this.anchors.length?this.active=this._findActive(i.active):this.active=t(),this._refresh(),this.active.length&&this.load(i.active)},_initialActive:function(){var e=this.options.active,i=this.options.collapsible,s=location.hash.substring(1),n=decodeURIComponent(s);return null===e&&(s&&(this.tabs.each(function(i,n){if(t(n).attr("aria-controls")===s)return e=i,!1}),null===e&&this.tabs.each(function(i,s){if(t(s).attr("aria-controls")===n)return e=i,!1})),null===e&&(e=this.tabs.index(this.tabs.filter(".ui-tabs-active"))),(null===e||-1===e)&&(e=!!this.tabs.length&&0)),!1!==e&&-1===(e=this.tabs.index(this.tabs.eq(e)))&&(e=!i&&0),!i&&!1===e&&this.anchors.length&&(e=0),e},_getCreateEventData:function(){return{tab:this.active,panel:this.active.length?this._getPanelForTab(this.active):t()}},_tabKeydown:function(e){var i=t(this.document[0].activeElement).closest("li"),s=this.tabs.index(i),n=!0;if(!this._handlePageNav(e)){switch(e.keyCode){case t.ui.keyCode.RIGHT:case t.ui.keyCode.DOWN:s++;break;case t.ui.keyCode.UP:case t.ui.keyCode.LEFT:n=!1,s--;break;case t.ui.keyCode.END:s=this.anchors.length-1;break;case t.ui.keyCode.HOME:s=0;break;case t.ui.keyCode.SPACE:e.preventDefault(),clearTimeout(this.activating),this._activate(s);return;case t.ui.keyCode.ENTER:e.preventDefault(),clearTimeout(this.activating),this._activate(s!==this.options.active&&s);return;default:return}e.preventDefault(),clearTimeout(this.activating),s=this._focusNextTab(s,n),e.ctrlKey||e.metaKey||(i.attr("aria-selected","false"),this.tabs.eq(s).attr("aria-selected","true"),this.activating=this._delay(function(){this.option("active",s)},this.delay))}},_panelKeydown:function(e){!this._handlePageNav(e)&&e.ctrlKey&&e.keyCode===t.ui.keyCode.UP&&(e.preventDefault(),this.active.trigger("focus"))},_handlePageNav:function(e){return e.altKey&&e.keyCode===t.ui.keyCode.PAGE_UP?(this._activate(this._focusNextTab(this.options.active-1,!1)),!0):e.altKey&&e.keyCode===t.ui.keyCode.PAGE_DOWN?(this._activate(this._focusNextTab(this.options.active+1,!0)),!0):void 0},_findNextTab:function(e,i){var s=this.tabs.length-1;while(-1!==t.inArray((e>s&&(e=0),e<0&&(e=s),e),this.options.disabled))e=i?e+1:e-1;return e},_focusNextTab:function(t,e){return t=this._findNextTab(t,e),this.tabs.eq(t).trigger("focus"),t},_setOption:function(t,e){if("active"===t){this._activate(e);return}this._super(t,e),"collapsible"!==t||(this._toggleClass("ui-tabs-collapsible",null,e),e||!1!==this.options.active||this._activate(0)),"event"===t&&this._setupEvents(e),"heightStyle"===t&&this._setupHeightStyle(e)},refresh:function(){var e=this.options,i=this.tablist.children(":has(a[href])");e.disabled=t.map(i.filter(".ui-state-disabled"),function(t){return i.index(t)}),this._processTabs(),!1!==e.active&&this.anchors.length?this.active.length&&!t.contains(this.tablist[0],this.active[0])?this.tabs.length===e.disabled.length?(e.active=!1,this.active=t()):this._activate(this._findNextTab(Math.max(0,e.active-1),!1)):e.active=this.tabs.index(this.active):(e.active=!1,this.active=t()),this._refresh()},_refresh:function(){this._setOptionDisabled(this.options.disabled),this._setupEvents(this.options.event),this._setupHeightStyle(this.options.heightStyle),this.tabs.not(this.active).attr({"aria-selected":"false","aria-expanded":"false",tabIndex:-1}),this.panels.not(this._getPanelForTab(this.active)).hide().attr({"aria-hidden":"true"}),this.active.length?(this.active.attr({"aria-selected":"true","aria-expanded":"true",tabIndex:0}),this._addClass(this.active,"ui-tabs-active","ui-state-active"),this._getPanelForTab(this.active).show().attr({"aria-hidden":"false"})):this.tabs.eq(0).attr("tabIndex",0)},_processTabs:function(){var e=this,i=this.tabs,s=this.anchors,n=this.panels;this.tablist=this._getList().attr("role","tablist"),this._addClass(this.tablist,"ui-tabs-nav","ui-helper-reset ui-helper-clearfix ui-widget-header"),this.tablist.on("mousedown"+this.eventNamespace,"> li",function(e){t(this).is(".ui-state-disabled")&&e.preventDefault()}),this.tabs=this.tablist.find("> li:has(a[href])").attr({role:"tab",tabIndex:-1}),this._addClass(this.tabs,"ui-tabs-tab","ui-state-default"),this.anchors=this.tabs.map(function(){return t("a",this)[0]}).attr({tabIndex:-1}),this._addClass(this.anchors,"ui-tabs-anchor"),this.panels=t(),this.anchors.each(function(i,s){var n,o,a,r=t(s).uniqueId().attr("id"),l=t(s).closest("li"),h=l.attr("aria-controls");e._isLocal(s)?(a=(n=s.hash).substring(1),(o=e.element.find("#"+CSS.escape(a))).length||(a=decodeURIComponent(a),o=e.element.find("#"+CSS.escape(a)))):(n="#"+(a=l.attr("aria-controls")||t({}).uniqueId()[0].id),(o=e.element.find(n)).length||(o=e._createPanel(a)).insertAfter(e.panels[i-1]||e.tablist),o.attr("aria-live","polite")),o.length&&(e.panels=e.panels.add(o)),h&&l.data("ui-tabs-aria-controls",h),l.attr({"aria-controls":a,"aria-labelledby":r}),o.attr("aria-labelledby",r)}),this.panels.attr("role","tabpanel"),this._addClass(this.panels,"ui-tabs-panel","ui-widget-content"),i&&(this._off(i.not(this.tabs)),this._off(s.not(this.anchors)),this._off(n.not(this.panels)))},_getList:function(){return this.tablist||this.element.find("ol, ul").eq(0)},_createPanel:function(e){return t("<div>").attr("id",e).data("ui-tabs-destroy",!0)},_setOptionDisabled:function(e){var i,s,n;for(Array.isArray(e)&&(e.length?e.length===this.anchors.length&&(e=!0):e=!1),n=0;s=this.tabs[n];n++)i=t(s),!0===e||-1!==t.inArray(n,e)?(i.attr("aria-disabled","true"),this._addClass(i,null,"ui-state-disabled")):(i.removeAttr("aria-disabled"),this._removeClass(i,null,"ui-state-disabled"));this.options.disabled=e,this._toggleClass(this.widget(),this.widgetFullName+"-disabled",null,!0===e)},_setupEvents:function(e){var i={};e&&t.each(e.split(" "),function(t,e){i[e]="_eventHandler"}),this._off(this.anchors.add(this.tabs).add(this.panels)),this._on(!0,this.anchors,{click:function(t){t.preventDefault()}}),this._on(this.anchors,i),this._on(this.tabs,{keydown:"_tabKeydown"}),this._on(this.panels,{keydown:"_panelKeydown"}),this._focusable(this.tabs),this._hoverable(this.tabs)},_setupHeightStyle:function(e){var i,s=this.element.parent();"fill"===e?(i=s.height()-(this.element.outerHeight()-this.element.height()),this.element.siblings(":visible").each(function(){var e=t(this),s=e.css("position");"absolute"!==s&&"fixed"!==s&&(i-=e.outerHeight(!0))}),this.element.children().not(this.panels).each(function(){i-=t(this).outerHeight(!0)}),this.panels.each(function(){t(this).height(Math.max(0,i-t(this).innerHeight()+t(this).height()))}).css("overflow","auto")):"auto"===e&&(i=0,this.panels.each(function(){i=Math.max(i,t(this).height("").height())}).height(i))},_eventHandler:function(e){var i=this.options,s=this.active,n=t(e.currentTarget).closest("li"),o=n[0]===s[0],a=o&&i.collapsible,r=a?t():this._getPanelForTab(n),l=s.length?this._getPanelForTab(s):t(),h={oldTab:s,oldPanel:l,newTab:a?t():n,newPanel:r};e.preventDefault(),!(n.hasClass("ui-state-disabled")||n.hasClass("ui-tabs-loading"))&&!this.running&&(!o||i.collapsible)&&!1!==this._trigger("beforeActivate",e,h)&&(i.active=!a&&this.tabs.index(n),this.active=o?t():n,this.xhr&&this.xhr.abort(),l.length||r.length||t.error("jQuery UI Tabs: Mismatching fragment identifier."),r.length&&this.load(this.tabs.index(n),e),this._toggle(e,h))},_toggle:function(e,i){var s=this,n=i.newPanel,o=i.oldPanel;function a(){s.running=!1,s._trigger("activate",e,i)}function r(){s._addClass(i.newTab.closest("li"),"ui-tabs-active","ui-state-active"),n.length&&s.options.show?s._show(n,s.options.show,a):(n.show(),a())}this.running=!0,o.length&&this.options.hide?this._hide(o,this.options.hide,function(){s._removeClass(i.oldTab.closest("li"),"ui-tabs-active","ui-state-active"),r()}):(this._removeClass(i.oldTab.closest("li"),"ui-tabs-active","ui-state-active"),o.hide(),r()),o.attr("aria-hidden","true"),i.oldTab.attr({"aria-selected":"false","aria-expanded":"false"}),n.length&&o.length?i.oldTab.attr("tabIndex",-1):n.length&&this.tabs.filter(function(){return 0===t(this).attr("tabIndex")}).attr("tabIndex",-1),n.attr("aria-hidden","false"),i.newTab.attr({"aria-selected":"true","aria-expanded":"true",tabIndex:0})},_activate:function(e){var i,s=this._findActive(e);s[0]!==this.active[0]&&(s.length||(s=this.active),i=s.find(".ui-tabs-anchor")[0],this._eventHandler({target:i,currentTarget:i,preventDefault:t.noop}))},_findActive:function(e){return!1===e?t():this.tabs.eq(e)},_getIndex:function(t){return"string"==typeof t&&(t=this.anchors.index(this.anchors.filter("[href$='"+CSS.escape(t)+"']"))),t},_destroy:function(){this.xhr&&this.xhr.abort(),this.tablist.removeAttr("role").off(this.eventNamespace),this.anchors.removeAttr("role tabIndex").removeUniqueId(),this.tabs.add(this.panels).each(function(){t.data(this,"ui-tabs-destroy")?t(this).remove():t(this).removeAttr("role tabIndex aria-live aria-busy aria-selected aria-labelledby aria-hidden aria-expanded")}),this.tabs.each(function(){var e=t(this),i=e.data("ui-tabs-aria-controls");i?e.attr("aria-controls",i).removeData("ui-tabs-aria-controls"):e.removeAttr("aria-controls")}),this.panels.show(),"content"!==this.options.heightStyle&&this.panels.css("height","")},enable:function(e){var i=this.options.disabled;!1!==i&&(void 0===e?i=!1:(e=this._getIndex(e),i=Array.isArray(i)?t.map(i,function(t){return t!==e?t:null}):t.map(this.tabs,function(t,i){return i!==e?i:null})),this._setOptionDisabled(i))},disable:function(e){var i=this.options.disabled;if(!0!==i){if(void 0===e)i=!0;else{if(e=this._getIndex(e),-1!==t.inArray(e,i))return;i=Array.isArray(i)?t.merge([e],i).sort():[e]}this._setOptionDisabled(i)}},load:function(e,i){e=this._getIndex(e);var s=this,n=this.tabs.eq(e),o=n.find(".ui-tabs-anchor"),a=this._getPanelForTab(n),r={tab:n,panel:a},l=function(t,e){"abort"===e&&s.panels.stop(!1,!0),s._removeClass(n,"ui-tabs-loading"),a.removeAttr("aria-busy"),t===s.xhr&&delete s.xhr};!this._isLocal(o[0])&&(this.xhr=t.ajax(this._ajaxSettings(o,i,r)),"canceled"!==this.xhr.statusText&&(this._addClass(n,"ui-tabs-loading"),a.attr("aria-busy","true"),this.xhr.done(function(t,e,n){a.html(t),s._trigger("load",i,r),l(n,e)}).fail(function(t,e){l(t,e)})))},_ajaxSettings:function(e,i,s){var n=this;return{url:e.attr("href"),beforeSend:function(e,o){return n._trigger("beforeLoad",i,t.extend({jqXHR:e,ajaxSettings:o},s))}}},_getPanelForTab:function(e){var i=t(e).attr("aria-controls");return this.element.find("#"+CSS.escape(i))}}),!0===t.uiBackCompat&&t.widget("ui.tabs",t.ui.tabs,{_processTabs:function(){this._superApply(arguments),this._addClass(this.tabs,"ui-tab")}}),t.ui.tabs,t.widget("ui.tooltip",{version:"1.14.2",options:{classes:{"ui-tooltip":"ui-corner-all ui-widget-shadow"},content:function(){var e=t(this).attr("title");return t("<a>").text(e).html()},hide:!0,items:"[title]:not([disabled])",position:{my:"left top+15",at:"left bottom",collision:"flipfit flip"},show:!0,track:!1,close:null,open:null},_addDescribedBy:function(t,e){var i=(t.attr("aria-describedby")||"").split(/\s+/);i.push(e),t.data("ui-tooltip-id",e).attr("aria-describedby",String.prototype.trim.call(i.join(" ")))},_removeDescribedBy:function(e){var i=e.data("ui-tooltip-id"),s=(e.attr("aria-describedby")||"").split(/\s+/),n=t.inArray(i,s);-1!==n&&s.splice(n,1),e.removeData("ui-tooltip-id"),(s=String.prototype.trim.call(s.join(" ")))?e.attr("aria-describedby",s):e.removeAttr("aria-describedby")},_create:function(){this._on({mouseover:"open",focusin:"open"}),this.tooltips={},this.parents={},this.liveRegion=t("<div>").attr({role:"log","aria-live":"assertive","aria-relevant":"additions"}).appendTo(this.document[0].body),this._addClass(this.liveRegion,null,"ui-helper-hidden-accessible"),this.disabledTitles=t([])},_setOption:function(e,i){var s=this;this._super(e,i),"content"===e&&t.each(this.tooltips,function(t,e){s._updateContent(e.element)})},_setOptionDisabled:function(t){this[t?"_disable":"_enable"]()},_disable:function(){var e=this;t.each(this.tooltips,function(i,s){var n=t.Event("blur");n.target=n.currentTarget=s.element[0],e.close(n,!0)}),this.disabledTitles=this.disabledTitles.add(this.element.find(this.options.items).addBack().filter(function(){var e=t(this);if(e.is("[title]"))return e.data("ui-tooltip-title",e.attr("title")).removeAttr("title")}))},_enable:function(){this.disabledTitles.each(function(){var e=t(this);e.data("ui-tooltip-title")&&e.attr("title",e.data("ui-tooltip-title"))}),this.disabledTitles=t([])},open:function(e){var i=this,s=t(e?e.target:this.element).closest(this.options.items);!(!s.length||s.data("ui-tooltip-id"))&&(s.attr("title")&&s.data("ui-tooltip-title",s.attr("title")),s.data("ui-tooltip-open",!0),e&&"mouseover"===e.type&&s.parents().each(function(){var e,s=t(this);s.data("ui-tooltip-open")&&((e=t.Event("blur")).target=e.currentTarget=this,i.close(e,!0)),s.attr("title")&&(s.uniqueId(),i.parents[this.id]={element:this,title:s.attr("title")},s.attr("title",""))}),this._registerCloseHandlers(e,s),this._updateContent(s,e))},_updateContent:function(t,e){var i,s=this.options.content,n=this,o=e?e.type:null;if("string"==typeof s||s.nodeType||s.jquery)return this._open(e,t,s);(i=s.call(t[0],function(i){t.data("ui-tooltip-open")&&(e&&(e.type=o),n._open(e,t,i))}))&&this._open(e,t,i)},_open:function(e,i,s){var n,o,a,r,l=t.extend({},this.options.position);if(s){if(n=this._find(i)){n.tooltip.find(".ui-tooltip-content").html(s);return}i.is("[title]")&&(e&&"mouseover"===e.type?i.attr("title",""):i.removeAttr("title")),o=(n=this._tooltip(i)).tooltip,this._addDescribedBy(i,o.attr("id")),o.find(".ui-tooltip-content").html(s),this.liveRegion.children().hide(),(r=t("<div>").html(o.find(".ui-tooltip-content").html())).removeAttr("name").find("[name]").removeAttr("name"),r.removeAttr("id").find("[id]").removeAttr("id"),r.appendTo(this.liveRegion),this.options.track&&e&&/^mouse/.test(e.type)?(this._on(this.document,{mousemove:h}),h(e)):o.position(t.extend({of:i},this.options.position)),o.hide(),this._show(o,this.options.show),this.options.track&&this.options.show&&this.options.show.delay&&(a=this.delayedShow=setInterval(function(){o.is(":visible")&&(h(l.of),clearInterval(a))},13)),this._trigger("open",e,{tooltip:o})}function h(t){l.of=t,!o.is(":hidden")&&o.position(l)}},_registerCloseHandlers:function(e,i){var s={keyup:function(e){if(e.keyCode===t.ui.keyCode.ESCAPE){var s=t.Event(e);s.currentTarget=i[0],this.close(s,!0)}}};i[0]!==this.element[0]&&(s.remove=function(){var t=this._find(i);t&&this._removeTooltip(t.tooltip)}),e&&"mouseover"!==e.type||(s.mouseleave="close"),e&&"focusin"!==e.type||(s.focusout="close"),this._on(!0,i,s)},close:function(e){var i,s=this,n=t(e?e.currentTarget:this.element),o=this._find(n);if(!o){n.removeData("ui-tooltip-open");return}i=o.tooltip,!o.closing&&(clearInterval(this.delayedShow),n.data("ui-tooltip-title")&&!n.attr("title")&&n.attr("title",n.data("ui-tooltip-title")),this._removeDescribedBy(n),o.hiding=!0,i.stop(!0),this._hide(i,this.options.hide,function(){s._removeTooltip(t(this))}),n.removeData("ui-tooltip-open"),this._off(n,"mouseleave focusout keyup"),n[0]!==this.element[0]&&this._off(n,"remove"),this._off(this.document,"mousemove"),e&&"mouseleave"===e.type&&t.each(this.parents,function(e,i){t(i.element).attr("title",i.title),delete s.parents[e]}),o.closing=!0,this._trigger("close",e,{tooltip:i}),o.hiding||(o.closing=!1))},_tooltip:function(e){var i=t("<div>").attr("role","tooltip"),s=t("<div>").appendTo(i),n=i.uniqueId().attr("id");return this._addClass(s,"ui-tooltip-content"),this._addClass(i,"ui-tooltip","ui-widget ui-widget-content"),i.appendTo(this._appendTo(e)),this.tooltips[n]={element:e,tooltip:i}},_find:function(t){var e=t.data("ui-tooltip-id");return e?this.tooltips[e]:null},_removeTooltip:function(t){clearInterval(this.delayedShow),t.remove(),delete this.tooltips[t.attr("id")]},_appendTo:function(t){var e=t.closest(".ui-front, dialog");return e.length||(e=this.document[0].body),e},_destroy:function(){var e=this;t.each(this.tooltips,function(i,s){var n=t.Event("blur"),o=s.element;n.target=n.currentTarget=o[0],e.close(n,!0),t("#"+i).remove(),o.data("ui-tooltip-title")&&(o.attr("title")||o.attr("title",o.data("ui-tooltip-title")),o.removeData("ui-tooltip-title"))}),this.liveRegion.remove()}}),!0===t.uiBackCompat&&t.widget("ui.tooltip",t.ui.tooltip,{options:{tooltipClass:null},_tooltip:function(){var t=this._superApply(arguments);return this.options.tooltipClass&&t.tooltip.addClass(this.options.tooltipClass),t}}),t.ui.tooltip});
$(document).ready(function () {
    var $body = $('body');
    var expiryDate = new Date('12/31/2999');    
    var size = Math.floor($('body').css('font-size').replace('px', ''));    

    //Append a class of cpTextResizeOn on body to get right styles for text resizer.
    $body.addClass('cpTextResizeOn');

    var factor = 1;
    if ($.cookie('sizeFactor') == null || $.cookie('sizeFactor') == '0') {
        //Do nothing.
    }
    else {
        factor = parseFloat($.cookie('sizeFactor'));
        size = size + factor;
        $body.css('font-size', '' + size + 'px');
}

    $('.cpTextResizer').find('a').click(function (e) {
        e.preventDefault();
        resizePageText(size, $(this));
    });
});

//Resize the text based on cookie value set.
function resizePageText(originalFontSize, $anchor) {
    var $body = $('body');
    var expiryDate = new Date('12/31/2999');
    var currentFontSize = $body.css('font-size');
    var bigOrSmall = $anchor.attr('id');
    var multiplier = 1;
    var fontSize;
    if (bigOrSmall == "makeTextSmaller") {        
        $.cookie('sizeFactor', parseInt($.cookie('sizeFactor') == null ? 0 : $.cookie('sizeFactor')) - 1, { path: '/', expires: expiryDate });
        fontSize = parseInt(currentFontSize.replace('px', '')) - 1;
    } else if (bigOrSmall == "makeTextBigger") {        
        $.cookie('sizeFactor', parseInt($.cookie('sizeFactor') == null ? 0 : $.cookie('sizeFactor')) + 1, { path: '/', expires: expiryDate });
        fontSize = parseInt(currentFontSize.replace('px', '')) +1;
    }    
     $body.css('font-size', '' + fontSize + 'px');

     if (!window.isRemoveSetHeights)
     	SetHeights();//Adjust heights of the containers.
};
var PopupBasedAuthentication = function () {
    var self = this;
    var popupWindow;
    var url;
    var paramSet;
    var jsCallback;
    var popupOrigin;

    self.requireLoggedIn = function (redirectUri, parameterSet) {
        if (redirectUri == null || redirectUri == '') {
            redirectUri = (window.location.pathname + window.location.search).split('#')[0];
        }
        self.setupParameters(redirectUri, parameterSet);
        self.checkLoggedIn();
    };

    self.setupParameters = function (redirectUri, parameterSet) {
        url = redirectUri;
        paramSet = parameterSet;
    };

    self.setupJsCallback = function (callback) {
    	jsCallback = callback;
    };

    self.checkLoggedIn = function () {
        $.ajax({
            url: '/Authentication/IsLoggedIn',
            type: 'GET',
            success: self.processCheckLoggedInResponse,
            async: false
        })
    };

    self.processCheckLoggedInResponse = function (loggedInResponse) {
    	if (loggedInResponse.loggedIn == true) {
    		if (jsCallback == null) {
    			window.location.href = url;
    		}
    		else {
    			jsCallback();
    		}
    	}
    	else if (loggedInResponse.UseIdentityServer) {
    		window.location.href = '/MyAccount?from=url&url=' + encodeURIComponent(url);
	    } else {
		    self.authenticate(loggedInResponse.domain);
	    }
    };

    self.authenticate = function (domain) {
        var popupUrl = document.createElement('a');
        popupUrl.href = 'https://' + domain;
        popupOrigin = popupUrl.protocol + '//' + popupUrl.host;
        if (internetExplorerVersion == undefined) {
            self.setupEventToListenFromChildWindow();
        }

        self.disableScreen();

        var width = 500;
        var height = 420;

        var left = (screen.width / 2) - (width / 2);
        var top = (screen.height / 2) - (height / 2);
		
        if (navigator.appName == 'Microsoft Internet Explorer' || !!(navigator.userAgent.match(/Trident/) || navigator.userAgent.match(/rv:11/)) || (typeof $.browser !== "undefined" && $.browser.msie == 1)) {
			window.location = "/MyAccount?from=url&url=" + window.location.pathname + window.location.search;
			return;
		}

        popupWindow = window.open(
                popupOrigin + '/Authentication/Popup?'
				+ 'redirectUri=' + encodeURIComponent(url)
				+ '&parameterSet=' + paramSet
				+ self.getUsernameFromCookie()
                + '&callingHost=' + window.location.protocol + '//' + window.location.hostname,
            'SignIn',
            'height=420,width=500,toolbar=no,menubar=no,left=' + left + ',top=' + top);

        self.checkBlockedPopup(popupWindow);
        self.attachOnCloseEvent(popupWindow);
    };

    self.getUsernameFromCookie = function () {
        var usernameCookieValue = cookieGet("CP");
        if (usernameCookieValue != null && usernameCookieValue.indexOf('uname=') == 0) {
            return '&prePopulatedUsername=' + usernameCookieValue.substring(6);
        } else {
            return '';
        }
    };

    self.attachOnCloseEvent = function (win) {
        var intervalId;
        function checkWindow() {
            if (win && win.closed) {
                window.clearInterval(intervalId);
                self.enableScreen();
            }
        }
        intervalId = window.setInterval(checkWindow, 500);
    };

    self.disableScreen = function () {
        $('#editItemBehavior_backgroundElement').show();
    };

    self.enableScreen = function () {
        $('#editItemBehavior_backgroundElement').hide();
    };

    self.checkBlockedPopup = function (popup) {
        setTimeout(function () {
        	if (!popup || popup.closed || popup.closed == 'undefined' || (popup.outerHeight && popup.outerHeight === 0)) {
                alert('Popup blocked');
                self.enableScreen();
            }
        }, 25);
    };

    self.setupEventToListenFromChildWindow = function () {
        window.addEventListener('message', function (msg) {
            if (msg.source === popupWindow && msg.origin === popupOrigin && msg.data.Type == 'Redirect') {
                self.handleResponseFromChild(msg.data);
            }
        }, true);
    };

    self.internetExplorerMessageHandling = function (data) {
        self.handleResponseFromChild(data);
    };

    self.handleResponseFromChild = function (data) {
        if (data.Type == 'Redirect') {
            var redirectUrl = document.createElement('a');
            redirectUrl.href = data.Url.replace(/&amp;/g, '&');
            if (redirectUrl.protocol === window.location.protocol && redirectUrl.host === window.location.host) {
                window.location.href = redirectUrl.href;
            }
        }
    };
};
;
/**
 * http://www.openjs.com/scripts/events/keyboard_shortcuts/
 * Version : 2.01.B
 * By Binny V A
 * License : BSD
 */

var pressCount = 0; //this is used to track button presses for a specific table widget IE hack

shortcut = {
	'all_shortcuts':{},//All the shortcuts are stored in this array
	'add': function(shortcut_combination,callback,opt) {
		//Provide a set of default options
		var default_options = {
			'type':'keydown',
			'propagate':false,
			'disable_in_input':false,
			'target':document,
			'keycode':false
		}
		if(!opt) opt = default_options;
		else {
			for(var dfo in default_options) {
				if(typeof opt[dfo] == 'undefined') opt[dfo] = default_options[dfo];
			}
		}

		var ele = opt.target;
		if(typeof opt.target == 'string') ele = document.getElementById(opt.target);
		var ths = this;
		shortcut_combination = shortcut_combination.toLowerCase();

		//The function to be called at keypress
		var func = function(e) {
			e = e || window.event;
			var character = null;
			
		    //this is an IE only hack that will fix double (<p>) vs single line (<br>) breaks in the table widget (<div contenteditable=true)
			if (detectIE() && e.keyCode == 13 && e.target.tagName == 'DIV' && e.target.contentEditable == "true") {
			    if ((pressCount) % 4 == 0) {
			        var node = document.createElement('br')
			        var positionNode = document.createTextNode('\u200B'); //used as a fake node to advance the cursor
			        node.appendChild(positionNode);
			        var selection = window.getSelection();
			        var range = selection.getRangeAt(0);
			        range.insertNode(node);
			        range.setStartAfter(node);
			        range.setEndAfter(node);
			        selection.removeAllRanges();
			        selection.addRange(range);
			    }
			    else {
			        e.preventDefault();
			    }

			    pressCount++;
			    return false;
			}

			if(opt['disable_in_input']) { //Don't enable shortcut keys in Input, Textarea fields
				var element;
				if(e.target) element=e.target;
				else if(e.srcElement) element=e.srcElement;
				if(element.nodeType==3) element=element.parentNode;

				if(element.tagName == 'INPUT' || element.tagName == 'TEXTAREA') return;
			}
	
			//Find Which key is pressed
			if (e.keyCode) code = e.keyCode;
			else if (e.which) code = e.which;
			if (typeof (code) !== "undefined")
			{
				character = String.fromCharCode(code).toLowerCase();
				if (code == 188) character = ","; //If the user presses , when the type is onkeydown
				if (code == 190) character = "."; //If the user presses , when the type is onkeydown
			}

			var keys = shortcut_combination.split("+");
			//Key Pressed - counts the number of valid keypresses - if it is same as the number of keys, the shortcut function is invoked
			var kp = 0;
			
			//Work around for stupid Shift key bug created by using lowercase - as a result the shift+num combination was broken
			var shift_nums = {
				"`":"~",
				"1":"!",
				"2":"@",
				"3":"#",
				"4":"$",
				"5":"%",
				"6":"^",
				"7":"&",
				"8":"*",
				"9":"(",
				"0":")",
				"-":"_",
				"=":"+",
				";":":",
				"'":"\"",
				",":"<",
				".":">",
				"/":"?",
				"\\":"|"
			}
			//Special Keys - and their codes
			var special_keys = {
				'esc':27,
				'escape':27,
				'tab':9,
				'space':32,
				'return':13,
				'enter':13,
				'backspace':8,
	
				'scrolllock':145,
				'scroll_lock':145,
				'scroll':145,
				'capslock':20,
				'caps_lock':20,
				'caps':20,
				'numlock':144,
				'num_lock':144,
				'num':144,
				
				'pause':19,
				'break':19,
				
				'insert':45,
				'home':36,
				'delete':46,
				'end':35,
				
				'pageup':33,
				'page_up':33,
				'pu':33,
	
				'pagedown':34,
				'page_down':34,
				'pd':34,
	
				'left':37,
				'up':38,
				'right':39,
				'down':40,
	
				'f1':112,
				'f2':113,
				'f3':114,
				'f4':115,
				'f5':116,
				'f6':117,
				'f7':118,
				'f8':119,
				'f9':120,
				'f10':121,
				'f11':122,
				'f12':123
			}
	
			var modifiers = { 
				shift: { wanted:false, pressed:false},
				ctrl : { wanted:false, pressed:false},
				alt  : { wanted:false, pressed:false},
				meta : { wanted:false, pressed:false}	//Meta is Mac specific
			};
                        
			if(e.ctrlKey)	modifiers.ctrl.pressed = true;
			if(e.shiftKey)	modifiers.shift.pressed = true;
			if(e.altKey)	modifiers.alt.pressed = true;
			if(e.metaKey)   modifiers.meta.pressed = true;
                        
			for(var i=0; k=keys[i],i<keys.length; i++) {
				//Modifiers
				if(k == 'ctrl' || k == 'control') {
					kp++;
					modifiers.ctrl.wanted = true;

				} else if(k == 'shift') {
					kp++;
					modifiers.shift.wanted = true;

				} else if(k == 'alt') {
					kp++;
					modifiers.alt.wanted = true;
				} else if(k == 'meta') {
					kp++;
					modifiers.meta.wanted = true;
				} else if(k.length > 1) { //If it is a special key
					if(special_keys[k] == code) kp++;
					
				} else if(opt['keycode']) {
					if(opt['keycode'] == code) kp++;

				} else { //The special keys did not match
					if(character == k) kp++;
					else {
						if(shift_nums[character] && e.shiftKey) { //Stupid Shift key bug created by using lowercase
							character = shift_nums[character]; 
							if(character == k) kp++;
						}
					}
				}
			}
			
			if(kp == keys.length && 
						modifiers.ctrl.pressed == modifiers.ctrl.wanted &&
						modifiers.shift.pressed == modifiers.shift.wanted &&
						modifiers.alt.pressed == modifiers.alt.wanted &&
						modifiers.meta.pressed == modifiers.meta.wanted) {
				callback(e);
	
				if(!opt['propagate']) { //Stop the event
					//e.cancelBubble is supported by IE - this will kill the bubbling process.
					e.cancelBubble = true;
					e.returnValue = false;
	
					//e.stopPropagation works in Firefox.
					if (e.stopPropagation) {
						e.stopPropagation();
						e.preventDefault();
					}
					return false;
				}
			}
		}
		this.all_shortcuts[shortcut_combination] = {
			'callback':func, 
			'target':ele, 
			'event': opt['type']
		};
		//Attach the function with the event
		if(ele.addEventListener) ele.addEventListener(opt['type'], func, false);
		else if(ele.attachEvent) ele.attachEvent('on'+opt['type'], func);
		else ele['on'+opt['type']] = func;
	},

	//Remove the shortcut - just specify the shortcut and I will remove the binding
	'remove':function(shortcut_combination) {
		shortcut_combination = shortcut_combination.toLowerCase();
		var binding = this.all_shortcuts[shortcut_combination];
		delete(this.all_shortcuts[shortcut_combination])
		if(!binding) return;
		var type = binding['event'];
		var ele = binding['target'];
		var callback = binding['callback'];

		if(ele.detachEvent) ele.detachEvent('on'+type, callback);
		else if(ele.removeEventListener) ele.removeEventListener(type, callback, false);
		else ele['on'+type] = false;
	}
}

/**
 * detect IE
 * returns version of IE or false, if browser is not Internet Explorer
 */
function detectIE() {
    var ua = window.navigator.userAgent;

    var msie = ua.indexOf('MSIE ');
    if (msie > 0) {
        // IE 10 or older => return version number
        return parseInt(ua.substring(msie + 5, ua.indexOf('.', msie)), 10);
    }

    var trident = ua.indexOf('Trident/');
    if (trident > 0) {
        // IE 11 => return version number
        var rv = ua.indexOf('rv:');
        return parseInt(ua.substring(rv + 3, ua.indexOf('.', rv)), 10);
    }

    var edge = ua.indexOf('Edge/');
    if (edge > 0) {
        // Edge (IE 12+) => return version number
        return parseInt(ua.substring(edge + 5, ua.indexOf('.', edge)), 10);
    }

    // other browser
    return false;
}
;
// http://www.codinghorror.com/blog/2006/02/standard-browser-keyboard-shortcuts.html
// NOTE: Ctrl+Shift+T is used by browsers to reopen the last closed tab
// NOTE: Ctrl+Shift+B is used by several browsers for various functions

//Toggle live edit
shortcut.add("Ctrl+Shift+L", function () {
	enableLiveEdit($.cookie("enableLiveEdit") === "false");
});

//Toggle set heights
shortcut.add("Ctrl+Shift+S", function () {
	window.Core.performSetHeights = !window.Core.performSetHeights;
});

//Toggle bundling
shortcut.add("Ctrl+Shift+A", function () {
	if (window.location.href.indexOf('bundle=on') > -1) {
		window.location.href = window.location.href.replace('bundle=on', 'bundle=off');
	} else if (window.location.href.indexOf('bundle=off') > -1) {
		window.location.href = window.location.href.replace('bundle=off', 'bundle=on');
	} else {
		var value = $('script[src="/Assets/Scripts/Shared.js"]').length == 0 ? 'off' : 'on';
		var connector = window.location.href.indexOf('?') > -1 ? '&' : '?';
		window.location.href = window.location.href + connector + 'bundle=' + value;
	}
});

//Toggle mobile
shortcut.add("Ctrl+Shift+M", function () {
	if (window.location.href.indexOf('mobile=on') > -1) {
		window.location.href = window.location.href.replace('mobile=on', 'mobile=off');
	} else if (window.location.href.indexOf('mobile=off') > -1) {
		window.location.href = window.location.href.replace('mobile=off', 'mobile=on');
	} else {
		var value = $('body').hasClass('mobile') ? 'off' : 'on';
		var connector = window.location.href.indexOf('?') > -1 ? '&' : '?';
		window.location.href = window.location.href + connector + 'mobile=' + value;
	}
});
;
var needRegisterH5Exts = true;
var isAlreadyInitialized = false;

// Only register the placeholder attachment code once, regardless of include count.
// File shouldn't be included more than once, but it could happen.
if (needRegisterH5Exts) {
	$(document).ready(function () {
		// Register placeholder shims, if needed.
		if (needPlaceHoldersShim() && !isAlreadyInitialized) {
		    isAlreadyInitialized = true;

		    // Not all INPUTs/TEXTAREAs are automatically hooked into this,
			// since we have older implementations of placeholders and
			// overridding them would cause wierd behavior. Plus, it
			// would not make sense to hook to image/checkbox/radio/etc.

			// The attribute cp5ph (CivicPlus HTML5 PlaceHolder) is automatically placed on inputs
			// with placeholders through the CPHtml5TextInput() HTML extension.
			registerPlaceHoldersShim($('input[cp5ph]'));
			registerPlaceHoldersShim($('textarea[cp5ph]'));
		}
	});

	needRegisterH5Exts = false;
}

// Registers placeholders on fields inside a specific context.
function registerPlaceHoldersShimOnContainer($container) {
	registerPlaceHoldersShim($('input[cp5ph]', $container));
	registerPlaceHoldersShim($('textarea[cp5ph]', $container));
}

// Unregisters placeholders on fields inside a specific context.
function unregisterPlaceHoldersShimOnContainer($container) {
	unregisterPlaceHolderShim($('input[cp5ph]', $container));
	unregisterPlaceHolderShim($('textarea[cp5ph]', $container));
}

// Determines if placeholder shim is necessary.
function needPlaceHoldersShim() {
	var tmpElem = document.createElement('input');
	return !('placeholder' in tmpElem);
}

// Added due to a bug with IE 10's IE 7 document standards mode.  The elem.getAttribute('placeholder')
// always returns NULL if IE 10 in IE 7 document standards mode, but elem.attributes['placeholder'] 
// works.
function getPlaceholder(elem) {
    var placeholder = elem.getAttribute('placeholder')
    if (placeholder == null) {
        placeholder = elem.attributes['placeholder'];
        if (placeholder == null)
            return '';
        else
            return placeholder.value;
    }
    return placeholder;
}

// Unregisters placeholder shim hooks. Destroys event hooks and spawned placeholder element.
function unregisterPlaceHolderShim($elemsToHack) {
	// Get hooks.
	var hookData = $elemsToHack.data('cp5ph_hooks');

	if (hookData) {
		// Destroy all the hooks present.
		var attachedMSIE = hookData.ie == null ? [] : hookData.ie;
		var attachedSTD = hookData.std == null ? [] : hookData.std;
		var attachedJQR = hookData.jqr == null ? [] : hookData.jqr;
		var spawnedPHs = hookData.spawns == null ? [] : hookData.spawns;
		
		// Remove MSIE style events.
		for (var i = 0, len = attachedMSIE.length; i < len; i++) {
			if (attachedMSIE[i].element)
				attachedMSIE[i].element.detachEvent(attachedMSIE[i].event, attachedMSIE[i].method);
		}

		// Remove DOM standard (STD) events.
		for (var i = 0, len = attachedSTD.length; i < len; i++) {
			if (attachedSTD[i].element)
				attachedSTD[i].element.removeEventListener(attachedSTD[i].event, attachedSTD[i].method, false);
		}

		// Remove jQuery bindings.
		for (var i = 0, len = attachedJQR.length; i < len; i++) {
			if (attachedJQR[i].instance)
				attachedJQR[i].instance.unbind(attachedJQR[i].event, attachedJQR[i].method);
		}
		
		// Destroy placeholder spawned elements.
		for (var i = 0, len = spawnedPHs.length; i < len; i++) {
			if (spawnedPHs[i]) {
				spawnedPHs[i].innerHTML = "";
				spawnedPHs[i].style.display = "none";

				if (spawnedPHs[i].parentNode)
					spawnedPHs[i].parentNode.removeChild(spawnedPHs[i]);
			}
		}

		// Destroy hook data now that hooks have been cleaned.
		$elemsToHack.removeData('cp5ph_hooks');
	}
}

// Registers placeholder shims for jQuery element group specified.
// The shim is by no means perfect (won't catch things like attribute selectors causing changes).

function registerPlaceHoldersShim($elemsToHack) {
	// Track standard, MSIE, and jQuery event assignments so they can be cancelled/destroyed.
	// Also track any spawned placeholder elements since they too need to be cleaned up.
	var attachedMSIE = [];
	var attachedSTD = [];
	var attachedJQR = [];
	var spawnedPHs = [];

	// Hooks programmatic changes to the value property for browsers without 
	// placeholder support. jQuery did not seem to have a way to do this.
	var hookProgrammaticValueChange = function(elem, callback) {
		// Note: Browsers are buggy at this time with DOMControlValueChanged.

		// Handle Firefox Prior to 4 (which has built-in placeholder support).
		if (elem.__defineSetter__) {
			var oldSetter = elem.__lookupSetter__('value');

			if (oldSetter) {
				elem.__defineSetter__('value', function(v) {
					oldSetter.call(elem, v);
					callback();
				});
			}
		}

		// Handle IE, even up to IE 9.
		if (elem.attachEvent) {
			attachedMSIE.push({ element: elem, event: 'onpropertychange', method: callback });
			elem.attachEvent('onpropertychange', callback);
		}
	};

	var displayAttrMatch = new RegExp('^(class|value|className|placeholder|id|style|style\\..*)$');

	// Hooks display changes for an element (placeholder may need to update itself).
	var hookDisplayUpdates = function(elem, callback) {
		if (elem.addEventListener) {
			var domAttrModifiedHook = function(event) {
				// Prevents an infinite loop due to bubbling.
				event.stopPropagation();

				// Another necessary escape hatch.
				if (event.target.className == '_ph_shim')
					return;

				if (displayAttrMatch.test(event.attrName))
					callback(event.attrName);
			};

			attachedSTD.push({ element: elem, event: 'DOMAttrModified', method: domAttrModifiedHook });
			elem.addEventListener('DOMAttrModified', domAttrModifiedHook, false);
		} else if (elem.attachEvent) {
			var propertyChangeHook = function(event) {
				if (displayAttrMatch.test(event.propertyName))
					callback(event.propertyName);
			};

			attachedMSIE.push({ element: elem, event: 'onpropertychange', method: propertyChangeHook });
			elem.attachEvent('onpropertychange', propertyChangeHook);
		}
	};

	// Hooks display changes for an element or its parents.
	var hookDisplayUpdatesAndParents = function(elem, callback) {
		var node = elem;

		while (node && node != document.body) {
			hookDisplayUpdates(node, callback);

			node = node.parentNode;
		}
	};

	// Debouncer used to prevent the same placeholder updates from ocuring over and over again within a short time (30 ms).
	var debouncer = [];
	var isBouncy = function(element) {
		var now = function() {
			return new Date().getTime();
		};
		var getIndexOf = function(elem) {
			for (var i = 0; i < debouncer.length; i++) {
				if (debouncer[i].object === element) {
					return i;
				}
			}
			return -1;
		};
		var index = getIndexOf(element);
		if (index === -1) {
			var obj = { object: element, recentlyOccured: true };
			debouncer.push(obj);
			setTimeout(function () {
				obj.recentlyOccured = false;
			}
			,30);
			return false;
		} else if (debouncer[index].recentlyOccured) {
			return true;
		} else {
			debouncer[index].recentlyOccured = true;
			setTimeout(function () {
				debouncer[index].recentlyOccured = false;
			}
			, 30);
			return false;
		}
	};

	// Updates position/style of pseudo-element that acts as placeholder message.
	// Called when necessary to ensure placeholder looks right (e.g. styles changed).
	var updatePlaceHolderMessageElement = function (phElem, elem) {
		var $elem = $(elem);
		if (isBouncy(elem)) {
			return;
		}

		// TODO: Does $elem.is() not work if the element is not in the document?
		if (elem.parentNode != null && !elem._focus && $elem.is(':visible'))
			phElem.style.display = (elem.value == '' || elem.value == null ? 'block' : 'none');
		else {
			phElem.style.display = 'none';
			return;
		}

		// Determine where placeholder element should go.
		var shouldParent = $elem.offsetParent().get(0);

		// Placeholder element cannot go inside table rows, use cell instead.
		if (shouldParent.tagName == 'TR') {
			shouldParent = elem;

			while (shouldParent.tagName != 'TD')
				shouldParent = shouldParent.parentNode;
		}

		var phChanged = (phElem.getAttribute('_ph_val') != getPlaceholder(elem));

		if ((shouldParent != phElem.parentNode) || phChanged) {
			// Remove from old parent (or existing, if phChanged).
			// Necessary for phChanged since element may be inside a table,
			// and trying to set innerHTML in a table with IE causes trouble.
			phElem.parentNode.removeChild(phElem);

			// If placeholder changed, update innerHTML.
			if (phChanged)
			    $(phElem).text(getPlaceholder(elem));

			// Ensure pseudo-placeholder is placed in same offsetParent
			// that element resides within (positioning), and add it to
			// the end (resolves some possible z-index issues).
			shouldParent.appendChild(phElem);
		}

		phElem.style.width = elem.clientWidth + 'px';
		phElem.style.height = elem.clientHeight + 'px';
		phElem.style.verticalAlign = elem.style.verticalAlign;
		phElem.style.lineHeight = elem.style.lineHeight;
		phElem.style.textAlign = elem.style.textAlign;
		phElem.style.fontSize = elem.style.fontSize;
		phElem.style.fontFamily = elem.style.fontFamily;

		var compStyle = (window.getComputedStyle ? window.getComputedStyle(elem, null) : elem.currentStyle);
		var elemPaddingLeft = parseInt(compStyle.paddingLeft.replace(/px$/g, ''), 10);
		var elemPaddingTop = parseInt(compStyle.paddingTop.replace(/px$/g, ''), 10);
		var pos = $elem.position();

		phElem.style.top = (pos.top + elem.clientTop + elemPaddingTop - 2) + 'px';
		phElem.style.left = (pos.left + elem.clientLeft + elemPaddingLeft) + 'px';
	}

	// Creates pseudo-element that acts as placeholder message.
	var createPlaceHolderMessageElement = function(elem) {
		// Create pseudo-placeholder that will appear over the top of the control.
		var phElem = document.createElement('div');

		// Attach to element for reference.
		elem._ph_shim = phElem;

		// Ensure pseudo-placeholder is placed in same offsetParent
		// that element resides within (positioning), and add it to
		// the end (resolves some possible z-index issues).
		var phValue = getPlaceholder(elem);
		$(phElem).text(phValue == null ? '' : phValue);
		phElem.setAttribute('_ph_val', 'phValue');

		phElem.setAttribute('id', 'ph_' + elem.id);
		phElem.setAttribute('name', 'ph_' + elem.id);

		var shouldParent = $(elem).offsetParent().get(0);
		shouldParent.appendChild(phElem);

		// Configure the pseudo-placeholder.
		phElem.className = '_ph_shim';
		phElem.style.position = 'absolute';
		phElem.style.color = '#aaa'; // Placeholder Text Color.
		phElem.style.backgroundColor = 'transparent';
		phElem.style.cssFloat = 'none';

		// Update appearance.
		updatePlaceHolderMessageElement(phElem, elem);

		// Set up events for pseudo-placeholder and element.
		var $phElem = $(phElem);

		var mouseDownHook = function(event) {
			phElem.style.display = 'none';

			// Mouse down is passed to the input control
			// the placeholder is shown on top of.
			$(elem).trigger('mousedown', event);

			// Hack: IE needs setTimeout to actually do this properly for TEXTAREA.
			setTimeout(function() {
				try {
					elem.focus()
				} catch(err) {
					//When the text box is readonly then the above focus code would throw error
					//so that error is only handled only for Center, if u happen to get this error please do ur implemntation below
					//NOTE: here we have checked only for /formcenter and so if there is any other place which uses readonly textbox and trys to focus also will get the same alert - This is known bug 
					//for DEtails talk to Akila or Robin - AK

					//**Please note the use of lowercase here**
					var url = (location.href).toLowerCase();

					if (url.indexOf("/admin/formcenter") > -1) {
						alert('You have selected the field input. Please select the field label or instructions.');
					}
				}
			}, 15);
		};

		$phElem.mousedown(mouseDownHook);

		attachedJQR.push({ instance: $phElem, event: 'mousedown', method: mouseDownHook });

		return phElem;
	};

	// Registers a shim placeholder for an element (private helper method).
	var registerPlaceHolder = function (elem) {
		// Create jQuery wrapper.
		var $elem = $(elem);

		// Create placeholder.
		var phElem = createPlaceHolderMessageElement(elem);

		// Track creation.
		spawnedPHs.push(phElem);

		var focusHook = function () {
			// Older firefox does not support document.activeElement.
			elem._focus = true;

			phElem.style.display = 'none';
		};

		var blurHook = function () {
			// Older firefox does not support document.activeElement.
			elem._focus = false;

			if ((elem.value == '' || elem.value == null) && $(elem).is(':visible'))
				phElem.style.display = 'block';
		};

		$elem.focus(focusHook);
		$elem.blur(blurHook);

		attachedJQR.push({ instance: $elem, event: 'focus', method: focusHook });
		attachedJQR.push({ instance: $elem, event: 'blur', method: blurHook });

		// Ensure placeholder state is updated when the value
		// is changed programmatically.
		hookProgrammaticValueChange(elem, function () {
			if (elem._focus)
				phElem.style.display = 'none';
			else {
				if ((elem.value == '' || elem.value == null) && $(elem).is(':visible'))
					phElem.style.display = 'block';
				else
					phElem.style.display = 'none';
			}
		});

		// Ensure placeholder UI updates correctly when display changes are made.
		hookDisplayUpdatesAndParents(elem, function () {
			updatePlaceHolderMessageElement(phElem, elem);
		});
	};

	// Attach shim to each wrapped element passed to the method.
	$elemsToHack.each(function () {
		registerPlaceHolder(this);
	});

	// Keep track of hooks so they can be destroyed.
	$elemsToHack.data('cp5ph_hooks', {
		ie: attachedMSIE,
		std: attachedSTD,
		jqr: attachedJQR,
		spawns: spawnedPHs 
	});
}


// Unregisters placeholder shim hooks. Destroys event hooks and spawned placeholder element.
function unregisterPlaceHoldersShimForEach($elemsToHack) {
	// Get hooks.
	$elemsToHack.each(function () {
		var hookData = $(this).data('cp5ph_hooks');

		if (hookData) {
			// Destroy all the hooks present.
			var attachedMSIE = hookData.ie == null ? [] : hookData.ie;
			var attachedSTD = hookData.std == null ? [] : hookData.std;
			var attachedJQR = hookData.jqr == null ? [] : hookData.jqr;
			var spawnedPHs = hookData.spawns == null ? [] : hookData.spawns;

			// Remove MSIE style events.
			for (var i = 0, len = attachedMSIE.length; i < len; i++) {
				if (attachedMSIE[i].element)
					attachedMSIE[i].element.detachEvent(attachedMSIE[i].event, attachedMSIE[i].method);
			}

			// Remove DOM standard (STD) events.
			for (var i = 0, len = attachedSTD.length; i < len; i++) {
				if (attachedSTD[i].element)
					attachedSTD[i].element.removeEventListener(attachedSTD[i].event, attachedSTD[i].method, false);
			}

			// Remove jQuery bindings.
			for (var i = 0, len = attachedJQR.length; i < len; i++) {
				if (attachedJQR[i].instance)
					attachedJQR[i].instance.unbind(attachedJQR[i].event, attachedJQR[i].method);
			}

			// Destroy placeholder spawned elements.
			for (var i = 0, len = spawnedPHs.length; i < len; i++) {
				if (spawnedPHs[i]) {
					spawnedPHs[i].innerHTML = "";
					spawnedPHs[i].style.display = "none";

					if (spawnedPHs[i].parentNode)
						spawnedPHs[i].parentNode.removeChild(spawnedPHs[i]);
				}
			}

			// Destroy hook data now that hooks have been cleaned.
			$(this).removeData('cp5ph_hooks');
		}
	});
}

// Registers placeholder shims for jQuery element group specified.
// The shim is by no means perfect (won't catch things like attribute selectors causing changes).
function registerPlaceHoldersShimForEach($elemsToHack) {
	// Track standard, MSIE, and jQuery event assignments so they can be cancelled/destroyed.
	// Also track any spawned placeholder elements since they too need to be cleaned up.
	var attachedMSIE = [];
	var attachedSTD = [];
	var attachedJQR = [];
	var spawnedPHs = [];

	// Hooks programmatic changes to the value property for browsers without 
	// placeholder support. jQuery did not seem to have a way to do this.
	var hookProgrammaticValueChange = function (elem, callback) {
		// Note: Browsers are buggy at this time with DOMControlValueChanged.

		// Handle Firefox Prior to 4 (which has built-in placeholder support).
		if (elem.__defineSetter__) {
			var oldSetter = elem.__lookupSetter__('value');

			if (oldSetter) {
				elem.__defineSetter__('value', function (v) {
					oldSetter.call(elem, v);
					callback();
				});
			}
		}

		// Handle IE, even up to IE 9.
		if (elem.attachEvent) {
			attachedMSIE.push({ element: elem, event: 'onpropertychange', method: callback });
			elem.attachEvent('onpropertychange', callback);
		}
	};

	var displayAttrMatch = new RegExp('^(class|value|className|placeholder|id|style|style\\..*)$');

	// Hooks display changes for an element (placeholder may need to update itself).
	var hookDisplayUpdates = function (elem, callback) {
		if (elem.addEventListener) {
			var domAttrModifiedHook = function (event) {
				// Prevents an infinite loop due to bubbling.
				event.stopPropagation();

				// Another necessary escape hatch.
				if (event.target.className == '_ph_shim')
					return;

				if (displayAttrMatch.test(event.attrName))
					callback(event.attrName);
			};

			attachedSTD.push({ element: elem, event: 'DOMAttrModified', method: domAttrModifiedHook });
			elem.addEventListener('DOMAttrModified', domAttrModifiedHook, false);
		}
		else if (elem.attachEvent) {
			var propertyChangeHook = function (event) {
				if (displayAttrMatch.test(event.propertyName))
					callback(event.propertyName);
			};

			attachedMSIE.push({ element: elem, event: 'onpropertychange', method: propertyChangeHook });
			elem.attachEvent('onpropertychange', propertyChangeHook);
		}
	};

	// Hooks display changes for an element or its parents.
	var hookDisplayUpdatesAndParents = function (elem, callback) {
		var node = elem;

		while (node && node != document.body) {
			hookDisplayUpdates(node, callback);

			node = node.parentNode;
		}
	};

	// Updates position/style of pseudo-element that acts as placeholder message.
	// Called when necessary to ensure placeholder looks right (e.g. styles changed).
	var updatePlaceHolderMessageElement = function (phElem, elem) {
		var $elem = $(elem);

		// TODO: Does $elem.is() not work if the element is not in the document?
		if (elem.parentNode != null && !elem._focus && $elem.is(':visible'))
			phElem.style.display = (elem.value == '' || elem.value == null ? 'block' : 'none');
		else {
			phElem.style.display = 'none';
			return;
		}

		// Determine where placeholder element should go.
		var shouldParent = $elem.offsetParent().get(0);

		// Placeholder element cannot go inside table rows, use cell instead.
		if (shouldParent.tagName == 'TR') {
			shouldParent = elem;

			while (shouldParent.tagName != 'TD')
				shouldParent = shouldParent.parentNode;
		}

		var phChanged = (phElem.getAttribute('_ph_val') != getPlaceholder(elem));

		if ((shouldParent != phElem.parentNode) || phChanged) {
			// Remove from old parent (or existing, if phChanged).
			// Necessary for phChanged since element may be inside a table,
			// and trying to set innerHTML in a table with IE causes trouble.
			phElem.parentNode.removeChild(phElem);

			// If placeholder changed, update innerHTML.
			if (phChanged)
				$(phElem).text(getPlaceholder(elem));

			// Ensure pseudo-placeholder is placed in same offsetParent
			// that element resides within (positioning), and add it to
			// the end (resolves some possible z-index issues).
			shouldParent.appendChild(phElem);
		}

		phElem.style.width = elem.clientWidth + 'px';
		phElem.style.height = elem.clientHeight + 'px';
		phElem.style.verticalAlign = elem.style.verticalAlign;
		phElem.style.lineHeight = elem.style.lineHeight;
		phElem.style.textAlign = elem.style.textAlign;
		phElem.style.fontSize = elem.style.fontSize;
		phElem.style.fontFamily = elem.style.fontFamily;

		var compStyle = (window.getComputedStyle ? window.getComputedStyle(elem, null) : elem.currentStyle);
		var elemPaddingLeft = parseInt(compStyle.paddingLeft.replace(/px$/g, ''), 10);
		var elemPaddingTop = parseInt(compStyle.paddingTop.replace(/px$/g, ''), 10);
		var pos = $elem.position();

		phElem.style.top = (pos.top + elem.clientTop + elemPaddingTop - 2) + 'px';
		phElem.style.left = (pos.left + elem.clientLeft + elemPaddingLeft) + 'px';
	}

	// Creates pseudo-element that acts as placeholder message.
	var createPlaceHolderMessageElement = function (elem) {
		// Create pseudo-placeholder that will appear over the top of the control.
		var phElem = document.createElement('div');

		// Attach to element for reference.
		elem._ph_shim = phElem;

		// Ensure pseudo-placeholder is placed in same offsetParent
		// that element resides within (positioning), and add it to
		// the end (resolves some possible z-index issues).
		var phValue = getPlaceholder(elem);
		$(phElem).text(phValue == null ? '' : phValue);
		phElem.setAttribute('_ph_val', 'phValue');

		phElem.setAttribute('id', 'ph_' + elem.id);
		phElem.setAttribute('name', 'ph_' + elem.id);

		var shouldParent = $(elem).offsetParent().get(0);
		shouldParent.appendChild(phElem);

		// Configure the pseudo-placeholder.
		phElem.className = '_ph_shim';
		phElem.style.position = 'absolute';
		phElem.style.color = '#aaa'; // Placeholder Text Color.
		phElem.style.backgroundColor = 'transparent';
		phElem.style.cssFloat = 'none';

		// Update appearance.
		updatePlaceHolderMessageElement(phElem, elem);

		// Set up events for pseudo-placeholder and element.
		var $phElem = $(phElem);

		var mouseDownHook = function (event) {
			phElem.style.display = 'none';

			// Mouse down is passed to the input control
			// the placeholder is shown on top of.
			$(elem).trigger('mousedown', event);

			// Hack: IE needs setTimeout to actually do this properly for TEXTAREA.
			setTimeout(function () {
				try { elem.focus() } catch (err) {
					//When the text box is readonly then the above focus code would throw error
					//so that error is only handled only for Center, if u happen to get this error please do ur implemntation below
					//NOTE: here we have checked only for /formcenter and so if there is any other place which uses readonly textbox and trys to focus also will get the same alert - This is known bug 
					//for DEtails talk to Akila or Robin - AK

					//**Please note the use of lowercase here**
					var url = (location.href).toLowerCase();

					if (url.indexOf("/admin/formcenter") > -1) {
						alert('You have selected the field input. Please select the field label or instructions.');
					}
				}
			}, 15);
		};

		$phElem.mousedown(mouseDownHook);

		attachedJQR.push({ instance: $phElem, event: 'mousedown', method: mouseDownHook });

		return phElem;
	}

	// Registers a shim placeholder for an element (private helper method).
	var registerPlaceHolder = function (elem) {
		// Create jQuery wrapper.
		var $elem = $(elem);

		// Create placeholder.
		var phElem = createPlaceHolderMessageElement(elem);

		// Track creation.
		spawnedPHs.push(phElem);

		var focusHook = function () {
			// Older firefox does not support document.activeElement.
			elem._focus = true;

			phElem.style.display = 'none';
		};

		var blurHook = function () {
			// Older firefox does not support document.activeElement.
			elem._focus = false;

			if ((elem.value == '' || elem.value == null) && $(elem).is(':visible'))
				phElem.style.display = 'block';
		};

		$elem.focus(focusHook);
		$elem.blur(blurHook);

		attachedJQR.push({ instance: $elem, event: 'focus', method: focusHook });
		attachedJQR.push({ instance: $elem, event: 'blur', method: blurHook });

		// Ensure placeholder state is updated when the value
		// is changed programmatically.
		hookProgrammaticValueChange(elem, function () {
			if (elem._focus)
				phElem.style.display = 'none';
			else {
				if ((elem.value == '' || elem.value == null) && $(elem).is(':visible'))
					phElem.style.display = 'block';
				else
					phElem.style.display = 'none';
			}
		});

		// Ensure placeholder UI updates correctly when display changes are made.
		hookDisplayUpdatesAndParents(elem, function () {
			updatePlaceHolderMessageElement(phElem, elem);
		});
	};

	// Attach shim to each wrapped element passed to the method.
	$elemsToHack.each(function () {
		registerPlaceHolder(this);
	});

	// Keep track of hooks so they can be destroyed.
	$elemsToHack.each(function() {
		$(this).data('cp5ph_hooks', {
			ie: attachedMSIE,
			std: attachedSTD,
			jqr: attachedJQR,
			spawns: spawnedPHs
			});
	});
}
;

