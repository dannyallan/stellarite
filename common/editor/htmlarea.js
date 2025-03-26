// htmlArea v3.0 - Copyright (c) 2002-2003, interactivetools.com, inc.
// All rights reserved.

function HTMLArea(textarea, config) {
	if (HTMLArea.checkSupportedBrowser()) {
		if (typeof config == "undefined") {
			this.config = new HTMLArea.Config();
		} else {
			this.config = config;
		}
		this._htmlArea = null;
		this._textArea = textarea;
		this._editMode = "wysiwyg";
		this._timerToolbar = null;
		this._timerUndo = null;
		this._undoQueue = new Array(this.config.undoSteps);
		this._undoPos = -1;
		this._customUndo = false;
		this._mdoc = document;
		this.doctype = '';
	}
};

HTMLArea.RE_tagName = /(<\/|<)\s*([^ \t\n>]+)/ig;
HTMLArea.RE_doctype = /(<!doctype((.|\n)*?)>)\n?/i;
HTMLArea.RE_head    = /<head>((.|\n)*?)<\/head>/i;
HTMLArea.RE_body    = /<body>((.|\n)*?)<\/body>/i;

HTMLArea.Config = function () {
	this.version = "3.0";
	this.width = "auto";
	this.height = "auto";
	this.statusBar = true;
	this.undoSteps = 20;
	this.undoTimeout = 500;
	this.sizeIncludesToolbar = true;
	this.fullPage = false;
	this.killWordOnPaste = false;
	this.baseURL = document.baseURI || document.URL;
	if (this.baseURL && this.baseURL.match(/(.*)\/([^\/]+)/))
		this.baseURL = RegExp.$1 + "/";

	this.imgURL = "images/";
	this.popupURL = "popups/";
	this.pageStyle = "body { background-color: #fff; font-family: verdana,arial,sans-serif; font-size:8pt; }";

	this.toolbar = [
		[ "fontname", "space",
		"fontsize", "space",
		"formatblock", "space",
		"bold", "italic", "underline", "separator",
		"forecolor", "hilitecolor"],

		[ "copy", "cut", "paste", "separator",
		"justifyleft", "justifycenter", "justifyright", "justifyfull", "separator",
		"orderedlist", "unorderedlist", "outdent", "indent", "separator",
		"horizontalrule", "createlink", "insertimage", "inserttable", "separator",
		"htmlmode"]
	];



	this.fontname = {
		"Arial":			'arial,helvetica,sans-serif',
		"Courier New":		'courier new,courier,monospace',
		"Georgia":			'georgia,times new roman,times,serif',
		"Tahoma":			'tahoma,arial,helvetica,sans-serif',
		"Times New Roman":	'times new roman,times,serif',
		"Verdana":			'verdana,arial,helvetica,sans-serif',
		"impact":			'impact',
		"WingDings":		'wingdings'
	};

	this.fontsize = {
		"1 (8 pt)":  "1",
		"2 (10 pt)": "2",
		"3 (12 pt)": "3",
		"4 (14 pt)": "4",
		"5 (18 pt)": "5",
		"6 (24 pt)": "6",
		"7 (36 pt)": "7"
	};

	this.formatblock = {
		"Heading 1": "h1",
		"Heading 2": "h2",
		"Heading 3": "h3",
		"Heading 4": "h4",
		"Heading 5": "h5",
		"Heading 6": "h6",
		"Normal": "p",
		"Address": "address",
		"Formatted": "pre"
	};

	this.customSelects = {};

	function cut_copy_paste(e, cmd, obj) {
		e.execCommand(cmd);
	};

	this.btnList = {
		bold: [ "Bold", "ed_format_bold.gif", false, function(e) {e.execCommand("bold");} ],
		italic: [ "Italic", "ed_format_italic.gif", false, function(e) {e.execCommand("italic");} ],
		underline: [ "Underline", "ed_format_underline.gif", false, function(e) {e.execCommand("underline");} ],
		strikethrough: [ "Strikethrough", "ed_format_strike.gif", false, function(e) {e.execCommand("strikethrough");} ],
		subscript: [ "Subscript", "ed_format_sub.gif", false, function(e) {e.execCommand("subscript");} ],
		superscript: [ "Superscript", "ed_format_sup.gif", false, function(e) {e.execCommand("superscript");} ],
		justifyleft: [ "Justify Left", "ed_align_left.gif", false, function(e) {e.execCommand("justifyleft");} ],
		justifycenter: [ "Justify Center", "ed_align_center.gif", false, function(e) {e.execCommand("justifycenter");} ],
		justifyright: [ "Justify Right", "ed_align_right.gif", false, function(e) {e.execCommand("justifyright");} ],
		justifyfull: [ "Justify Full", "ed_align_justify.gif", false, function(e) {e.execCommand("justifyfull");} ],
		orderedlist: [ "Ordered List", "ed_list_num.gif", false, function(e) {e.execCommand("orderedlist");} ],
		unorderedlist: [ "Bulleted List", "ed_list_bullet.gif", false, function(e) {e.execCommand("unorderedlist");} ],
		outdent: [ "Decrease Indent", "ed_indent_less.gif", false, function(e) {e.execCommand("outdent");} ],
		indent: [ "Increase Indent", "ed_indent_more.gif", false, function(e) {e.execCommand("indent");} ],
		forecolor: [ "Font Color", "ed_color_fg.gif", false, function(e) {e.execCommand("forecolor");} ],
		hilitecolor: [ "Background Color", "ed_color_bg.gif", false, function(e) {e.execCommand("hilitecolor");} ],
		horizontalrule: [ "Horizontal Rule", "ed_hr.gif", false, function(e) {e.execCommand("horizontalrule");} ],
		createlink: [ "Insert Web Link", "ed_link.gif", false, function(e) {e.execCommand("createlink", true);} ],
		insertimage: [ "Insert/Modify Image", "ed_image.gif", false, function(e) {e.execCommand("insertimage");} ],
		inserttable: [ "Insert Table", "insert_table.gif", false, function(e) {e.execCommand("inserttable");} ],
		htmlmode: [ "Toggle HTML Source", "ed_html.gif", true, function(e) {e.execCommand("htmlmode");} ],
		popupeditor: [ "Enlarge Editor", "fullscreen_maximize.gif", true, function(e) {e.execCommand("popupeditor");} ],
		undo: [ "Undoes your last action", "ed_undo.gif", false, function(e) {e.execCommand("undo");} ],
		redo: [ "Redoes your last action", "ed_redo.gif", false, function(e) {e.execCommand("redo");} ],
		cut: [ "Cut selection", "ed_cut.gif", false, cut_copy_paste ],
		copy: [ "Copy selection", "ed_copy.gif", false, cut_copy_paste ],
		paste: [ "Paste from clipboard", "ed_paste.gif", false, cut_copy_paste ],
		lefttoright: [ "Direction left to right", "ed_left_to_right.gif", false, function(e) {e.execCommand("lefttoright");} ],
		righttoleft: [ "Direction right to left", "ed_right_to_left.gif", false, function(e) {e.execCommand("righttoleft");} ]
	};
	for (var i in this.btnList) {
		var btn = this.btnList[i];
		btn[1] = _editor_url + this.imgURL + btn[1];
		if (typeof HTMLArea.I18N.tooltips[i] != "undefined") {
			btn[0] = HTMLArea.I18N.tooltips[i];
		}
	}
};

HTMLArea.Config.prototype.registerButton = function(id, tooltip, image, textMode, action, context) {
	var the_id;
	if (typeof id == "string") {
		the_id = id;
	} else if (typeof id == "object") {
		the_id = id.id;
	} else {
		alert("ERROR [HTMLArea.Config::registerButton]:\ninvalid arguments");
		return false;
	}
	if (typeof this.customSelects[the_id] != "undefined") {
	}
	if (typeof this.btnList[the_id] != "undefined") {
	}
	switch (typeof id) {
	    case "string": this.btnList[id] = [ tooltip, image, textMode, action, context ]; break;
	    case "object": this.btnList[id.id] = [ id.tooltip, id.image, id.textMode, id.action, id.context ]; break;
	}
};

HTMLArea.Config.prototype.registerDropdown = function(object) {
	if (typeof this.customSelects[object.id] != "undefined") {
	}
	if (typeof this.btnList[object.id] != "undefined") {
	}
	this.customSelects[object.id] = object;
};

HTMLArea.Config.prototype.hideSomeButtons = function(remove) {
	var toolbar = this.toolbar;
	for (var i in toolbar) {
		var line = toolbar[i];
		for (var j = line.length; --j >= 0; ) {
			if (remove.indexOf(" " + line[j] + " ") >= 0) {
				var len = 1;
				if (/separator|space/.test(line[j + 1])) {
					len = 2;
				}
				line.splice(j, len);
			}
		}
	}
};

HTMLArea.replaceAll = function(config) {
	var tas = document.getElementsByTagName("textarea");
	for (var i = tas.length; i > 0; (new HTMLArea(tas[--i], config)).generate());
};

HTMLArea.replace = function(id, config) {
	var ta = HTMLArea.getElementById("textarea", id);
	return ta ? (new HTMLArea(ta, config)).generate() : null;
};

HTMLArea.prototype._createToolbar = function () {
	var editor = this;

	var toolbar = document.createElement("div");
	this._toolbar = toolbar;
	toolbar.className = "toolbar";
	toolbar.unselectable = "1";
	var tb_row = null;
	var tb_objects = new Object();
	this._toolbarObjects = tb_objects;

	function newLine() {
		var table = document.createElement("table");
		table.border = "0px";
		table.cellSpacing = "0px";
		table.cellPadding = "0px";
		toolbar.appendChild(table);
		var tb_body = document.createElement("tbody");
		table.appendChild(tb_body);
		tb_row = document.createElement("tr");
		tb_body.appendChild(tb_row);
	};
	newLine();

	function setButtonStatus(id, newval) {
		var oldval = this[id];
		var el = this.element;
		if (oldval != newval) {
			switch (id) {
			    case "enabled":
				if (newval) {
					HTMLArea._removeClass(el, "buttonDisabled");
					el.disabled = false;
				} else {
					HTMLArea._addClass(el, "buttonDisabled");
					el.disabled = true;
				}
				break;
			    case "active":
				if (newval) {
					HTMLArea._addClass(el, "buttonPressed");
				} else {
					HTMLArea._removeClass(el, "buttonPressed");
				}
				break;
			}
			this[id] = newval;
		}
	};

	function createSelect(txt) {
		var options = null;
		var el = null;
		var cmd = null;
		var customSelects = editor.config.customSelects;
		var context = null;
		switch (txt) {
		    case "fontsize":
		    case "fontname":
		    case "formatblock":
			options = editor.config[txt];
			cmd = txt;
			break;
		    default:
			cmd = txt;
			var dropdown = customSelects[cmd];
			if (typeof dropdown != "undefined") {
				options = dropdown.options;
				context = dropdown.context;
			} else {
				alert("ERROR [createSelect]:\nCan't find the requested dropdown definition");
			}
			break;
		}
		if (options) {
			el = document.createElement("select");
			var obj = {
				name	: txt,
				element : el,
				enabled : true,
				text	: false,
				cmd	: cmd,
				state	: setButtonStatus,
				context : context
			};
			tb_objects[txt] = obj;
			for (var i in options) {
				var op = document.createElement("option");
				op.appendChild(document.createTextNode(i));
				op.value = options[i];
				el.appendChild(op);
			}
			HTMLArea._addEvent(el, "change", function () {
				editor._comboSelected(el, txt);
			});
		}
		return el;
	};

	function createButton(txt) {
		var el = null;
		var btn = null;
		switch (txt) {
		    case "separator":
			el = document.createElement("div");
			el.className = "separator";
			break;
		    case "space":
			el = document.createElement("div");
			el.className = "space";
			break;
		    case "linebreak":
			newLine();
			return false;
		    case "textindicator":
			el = document.createElement("div");
			el.appendChild(document.createTextNode("A"));
			el.className = "indicator";
			el.title = HTMLArea.I18N.tooltips.textindicator;
			var obj = {
				name	: txt,
				element : el,
				enabled : true,
				active	: false,
				text	: false,
				cmd	: "textindicator",
				state	: setButtonStatus
			};
			tb_objects[txt] = obj;
			break;
		    default:
			btn = editor.config.btnList[txt];
		}
		if (!el && btn) {
			el = document.createElement("div");
			el.title = btn[0];
			el.className = "button";
			var obj = {
				name	: txt,
				element : el,
				enabled : true,
				active	: false,
				text	: btn[2],
				cmd	: btn[3],
				state	: setButtonStatus,
				context : btn[4] || null
			};
			tb_objects[txt] = obj;
			HTMLArea._addEvent(el, "mouseover", function () {
				if (obj.enabled) {
					HTMLArea._addClass(el, "buttonHover");
				}
			});
			HTMLArea._addEvent(el, "mouseout", function () {
				if (obj.enabled) with (HTMLArea) {
					_removeClass(el, "buttonHover");
					_removeClass(el, "buttonActive");
					(obj.active) && _addClass(el, "buttonPressed");
				}
			});
			HTMLArea._addEvent(el, "mousedown", function (ev) {
				if (obj.enabled) with (HTMLArea) {
					_addClass(el, "buttonActive");
					_removeClass(el, "buttonPressed");
					_stopEvent(is_ie ? window.event : ev);
				}
			});
			HTMLArea._addEvent(el, "click", function (ev) {
				if (obj.enabled) with (HTMLArea) {
					_removeClass(el, "buttonActive");
					_removeClass(el, "buttonHover");
					obj.cmd(editor, obj.name, obj);
					_stopEvent(is_ie ? window.event : ev);
				}
			});
			var img = document.createElement("img");
			img.src = btn[1];
			img.style.width = "18px";
			img.style.height = "18px";
			el.appendChild(img);
		} else if (!el) {
			el = createSelect(txt);
		}
		if (el) {
			var tb_cell = document.createElement("td");
			tb_row.appendChild(tb_cell);
			tb_cell.appendChild(el);
		} else {
			alert("FIXME: Unknown toolbar item: " + txt);
		}
		return el;
	};

	var first = true;
	for (var i in this.config.toolbar) {
		if (!first) {
			createButton("linebreak");
		} else {
			first = false;
		}
		var group = this.config.toolbar[i];
		for (var j in group) {
			var code = group[j];
			if (/^([IT])\[(.*?)\]/.test(code)) {
				var l7ed = RegExp.$1 == "I";
				var label = RegExp.$2;
				if (l7ed) {
					label = HTMLArea.I18N.custom[label];
				}
				var tb_cell = document.createElement("td");
				tb_row.appendChild(tb_cell);
				tb_cell.className = "label";
				tb_cell.innerHTML = label;
			} else {
				createButton(code);
			}
		}
	}

	this._htmlArea.appendChild(toolbar);
};

HTMLArea.prototype._createStatusBar = function() {
	var statusbar = document.createElement("div");
	statusbar.className = "statusBar";
	this._htmlArea.appendChild(statusbar);
	this._statusBar = statusbar;
	div = document.createElement("span");
	div.className = "statusBarTree";
	div.innerHTML = HTMLArea.I18N.msg["Path"] + ": ";
	this._statusBarTree = div;
	this._statusBar.appendChild(div);
	if (!this.config.statusBar) {
		statusbar.style.display = "none";
	}
};

HTMLArea.prototype.generate = function () {
	var editor = this;
	var textarea = this._textArea;
	if (typeof textarea == "string") {
		this._textArea = textarea = HTMLArea.getElementById("textarea", textarea);
	}
	this._ta_size = {
		w: textarea.offsetWidth,
		h: textarea.offsetHeight
	};
	textarea.style.display = "none";

	var htmlarea = document.createElement("div");
	htmlarea.className = "oHTML";
	this._htmlArea = htmlarea;

	textarea.parentNode.insertBefore(htmlarea, textarea);

	if (textarea.form) {
		var f = textarea.form;
		if (typeof f.onsubmit == "function") {
			var funcref = f.onsubmit;
			if (typeof f.__msh_prevOnSubmit == "undefined") {
				f.__msh_prevOnSubmit = [];
			}
			f.__msh_prevOnSubmit.push(funcref);
		}
		f.onsubmit = function() {
			editor._textArea.value = editor.getHTML();
			var a = this.__msh_prevOnSubmit;
			if (typeof a != "undefined") {
				for (var i in a) {
					a[i]();
				}
			}
		};
	}

	window.onunload = function() {
		editor._textArea.value = editor.getHTML();
	};

	this._createToolbar();

	var iframe = document.createElement("iframe");
	htmlarea.appendChild(iframe);

	this._iframe = iframe;

	this._createStatusBar();

	if (!HTMLArea.is_ie) {
		iframe.style.borderWidth = "1px";
	}

	var height = (this.config.height == "auto" ? (this._ta_size.h + "px") : this.config.height);
	height = parseInt(height);
	var width = (this.config.width == "auto" ? (this._ta_size.w + "px") : this.config.width);
	width = parseInt(width);

	if (!HTMLArea.is_ie) {
		height -= 2;
		width -= 2;
	}

	iframe.style.width = width + "px";
	if (this.config.sizeIncludesToolbar) {
		height -= this._toolbar.offsetHeight;
		height -= this._statusBar.offsetHeight;
	}
	if (height < 0) {
		height = 0;
	}
	iframe.style.height = height + "px";

	textarea.style.width = iframe.style.width;
 	textarea.style.height = iframe.style.height;

	function initIframe() {
		var doc = editor._iframe.contentWindow.document;
		if (!doc) {
			if (HTMLArea.is_gecko) {
				setTimeout(initIframe, 100);
				return false;
			} else {
				alert("ERROR: IFRAME can't be initialized.");
			}
		}
		if (HTMLArea.is_gecko) {
			doc.designMode = "on";
		}
		editor._doc = doc;
		if (!editor.config.fullPage) {
			doc.open();
			var html = "<html>\n";
			html += "<head>\n";
			if (editor.config.baseURL)
				html += '<base href="' + editor.config.baseURL + '" />';
			html += "<style> html,body { border: 0px; } " +
				editor.config.pageStyle + "</style>\n";
			html += "</head>\n";
			html += "<body>\n";
			html += editor._textArea.value;
			html += "</body>\n";
			html += "</html>";
			doc.write(html);
			doc.close();
		} else {
			var html = editor._textArea.value;
			if (html.match(HTMLArea.RE_doctype)) {
				editor.setDoctype(RegExp.$1);
				html = html.replace(HTMLArea.RE_doctype, "");
			}
			doc.open();
			doc.write(html);
			doc.close();
		}

		if (HTMLArea.is_ie) {
			doc.body.contentEditable = true;
		}

		editor.focusEditor();

		HTMLArea._addEvents
			(doc, ["keydown", "keypress", "mousedown", "mouseup", "drag"],
			 function (event) {
				 return editor._editorEvent(HTMLArea.is_ie ? editor._iframe.contentWindow.event : event);
			 });

		setTimeout(function() {
			editor.updateToolbar();
		}, 250);

		if (typeof editor.onGenerate == "function")
			editor.onGenerate();
	};
	setTimeout(initIframe, 100);
};

HTMLArea.prototype.setMode = function(mode) {
	if (typeof mode == "undefined") {
		mode = ((this._editMode == "textmode") ? "wysiwyg" : "textmode");
	}
	switch (mode) {
	    case "textmode":
		this._textArea.value = this.getHTML();
		this._iframe.style.display = "none";
		this._textArea.style.display = "block";
		if (this.config.statusBar) {
			this._statusBar.innerHTML = HTMLArea.I18N.msg["TEXT_MODE"];
		}
		break;
	    case "wysiwyg":
		if (HTMLArea.is_gecko) {
			try {
				this._doc.designMode = "off";
			} catch(e) {};
		}
		if (!this.config.fullPage)
			this._doc.body.innerHTML = this.getHTML();
		else
			this.setFullHTML(this.getHTML());
		this._iframe.style.display = "block";
		this._textArea.style.display = "none";
		if (HTMLArea.is_gecko) {
			try {
				this._doc.designMode = "on";
			} catch(e) {};
		}
		if (this.config.statusBar) {
			this._statusBar.innerHTML = '';
			this._statusBar.appendChild(document.createTextNode(HTMLArea.I18N.msg["Path"] + ": "));
			this._statusBar.appendChild(this._statusBarTree);
		}
		break;
	    default:
		alert("Mode <" + mode + "> not defined!");
		return false;
	}
	this._editMode = mode;
	this.focusEditor();
};

HTMLArea.prototype.setFullHTML = function(html) {
	var save_multiline = RegExp.multiline;
	RegExp.multiline = true;
	if (html.match(HTMLArea.RE_doctype)) {
		this.setDoctype(RegExp.$1);
		html = html.replace(HTMLArea.RE_doctype, "");
	}
	RegExp.multiline = save_multiline;
	if (!HTMLArea.is_ie) {
		if (html.match(HTMLArea.RE_head))
			this._doc.getElementsByTagName("head")[0].innerHTML = RegExp.$1;
		if (html.match(HTMLArea.RE_body))
			this._doc.getElementsByTagName("body")[0].innerHTML = RegExp.$1;
	} else {
		var html_re = /<html>((.|\n)*?)<\/html>/i;
		html = html.replace(html_re, "$1");
		this._doc.open();
		this._doc.write(html);
		this._doc.close();
		this._doc.body.contentEditable = true;
		return true;
	}
};

HTMLArea.prototype._wordClean = function() {
	var D = this.getInnerHTML();
	if (D.indexOf('class=Mso') >= 0) {

		D = D.replace(/\r\n/g, ' ').
			replace(/\n/g, ' ').
			replace(/\r/g, ' ').
			replace(/\&nbsp\;/g,' ');

		D = D.replace(/ class=[^\s|>]*/gi,'').
			replace(/ style=\"[^>]*\"/gi,'').
			replace(/ align=[^\s|>]*/gi,'');

		D = D.replace(/<b [^>]*>/gi,'<b>').
			replace(/<i [^>]*>/gi,'<i>').
			replace(/<li [^>]*>/gi,'<li>').
			replace(/<ul [^>]*>/gi,'<ul>');

		D = D.replace(/<b>/gi,'<strong>').
			replace(/<\/b>/gi,'</strong>');

		D = D.replace(/<em>/gi,'<i>').
			replace(/<\/em>/gi,'</i>');

		D = D.replace(/<\?xml:[^>]*>/g, '').
			replace(/<\/?st1:[^>]*>/g,'').
			replace(/<\/?[a-z]\:[^>]*>/g,'').
			replace(/<\/?font[^>]*>/gi,'').
			replace(/<\/?span[^>]*>/gi,' ').
			replace(/<\/?div[^>]*>/gi,' ').
			replace(/<\/?pre[^>]*>/gi,' ').
			replace(/<\/?h[1-6][^>]*>/gi,' ');

		oldlen = D.length + 1;
		while(oldlen > D.length) {
			oldlen = D.length;
			D = D.replace(/<([a-z][a-z]*)> *<\/\1>/gi,' ').
				replace(/<([a-z][a-z]*)> *<([a-z][^>]*)> *<\/\1>/gi,'<$2>');
		}
		D = D.replace(/<([a-z][a-z]*)><\1>/gi,'<$1>').
			replace(/<\/([a-z][a-z]*)><\/\1>/gi,'<\/$1>');

		D = D.replace(/  */gi,' ');

		this.setHTML(D);
		this.updateToolbar();
	}
};

HTMLArea.prototype.forceRedraw = function() {
	this._doc.body.style.visibility = "hidden";
	this._doc.body.style.visibility = "visible";
};

HTMLArea.prototype.focusEditor = function() {
	switch (this._editMode) {
	    case "wysiwyg" : this._iframe.contentWindow.focus(); break;
	    case "textmode": this._textArea.focus(); break;
	    default	   : alert("ERROR: mode " + this._editMode + " is not defined");
	}
	return this._doc;
};

HTMLArea.prototype._undoTakeSnapshot = function() {
	++this._undoPos;
	if (this._undoPos >= this.config.undoSteps) {
		this._undoQueue.shift();
		--this._undoPos;
	}
	var take = true;
	var txt = this.getInnerHTML();
	if (this._undoPos > 0)
		take = (this._undoQueue[this._undoPos - 1] != txt);
	if (take) {
		this._undoQueue[this._undoPos] = txt;
	} else {
		this._undoPos--;
	}
};

HTMLArea.prototype.undo = function() {
	if (this._undoPos > 0) {
		var txt = this._undoQueue[--this._undoPos];
		if (txt) this.setHTML(txt);
		else ++this._undoPos;
	}
};

HTMLArea.prototype.redo = function() {
	if (this._undoPos < this._undoQueue.length - 1) {
		var txt = this._undoQueue[++this._undoPos];
		if (txt) this.setHTML(txt);
		else --this._undoPos;
	}
};

HTMLArea.prototype.updateToolbar = function(noStatus) {
	var doc = this._doc;
	var text = (this._editMode == "textmode");
	var ancestors = null;
	if (!text) {
		ancestors = this.getAllAncestors();
		if (this.config.statusBar && !noStatus) {
			this._statusBarTree.innerHTML = HTMLArea.I18N.msg["Path"] + ": ";
			for (var i = ancestors.length; --i >= 0;) {
				var el = ancestors[i];
				if (!el) {
					continue;
				}
				var a = document.createElement("a");
				a.href = "#";
				a.el = el;
				a.editor = this;
				a.onclick = function() {
					this.blur();
					this.editor.selectNodeContents(this.el);
					this.editor.updateToolbar(true);
					return false;
				};
				a.oncontextmenu = function() {
					this.blur();
					var info = "Inline style:\n\n";
					info += this.el.style.cssText.split(/;\s*/).join(";\n");
					alert(info);
					return false;
				};
				var txt = el.tagName.toLowerCase();
				a.title = el.style.cssText;
				if (el.id) {
					txt += "#" + el.id;
				}
				if (el.className) {
					txt += "." + el.className;
				}
				a.appendChild(document.createTextNode(txt));
				this._statusBarTree.appendChild(a);
				if (i != 0) {
					this._statusBarTree.appendChild(document.createTextNode(String.fromCharCode(0xbb)));
				}
			}
		}
	}
	for (var i in this._toolbarObjects) {
		var btn = this._toolbarObjects[i];
		var cmd = i;
		var inContext = true;
		if (btn.context && !text) {
			inContext = false;
			var context = btn.context;
			var attrs = [];
			if (/(.*)\[(.*?)\]/.test(context)) {
				context = RegExp.$1;
				attrs = RegExp.$2.split(",");
			}
			context = context.toLowerCase();
			var match = (context == "*");
			for (var k in ancestors) {
				if (!ancestors[k]) {
					continue;
				}
				if (match || (ancestors[k].tagName.toLowerCase() == context)) {
					inContext = true;
					for (var ka in attrs) {
						if (!eval("ancestors[k]." + attrs[ka])) {
							inContext = false;
							break;
						}
					}
					if (inContext) {
						break;
					}
				}
			}
		}
		btn.state("enabled", (!text || btn.text) && inContext);
		if (typeof cmd == "function") {
			continue;
		}
		var dropdown = this.config.customSelects[cmd];
		if ((!text || btn.text) && (typeof dropdown != "undefined")) {
			dropdown.refresh(this);
			continue;
		}
		switch (cmd) {
		    case "fontname":
		    case "fontsize":
		    case "formatblock":
			if (!text) try {
				var value = ("" + doc.queryCommandValue(cmd)).toLowerCase();
				if (!value) {
					break;
				}
				var options = this.config[cmd];
				var k = 0;

				for (var j in options) {
					if ((j.toLowerCase() == value) ||
					    (options[j].substr(0, value.length).toLowerCase() == value)) {
						btn.element.selectedIndex = k;
						break;
					}
					++k;
				}
			} catch(e) {};
			break;
		    case "textindicator":
			if (!text) {
				try {with (btn.element.style) {
					backgroundColor = HTMLArea._makeColor(
						doc.queryCommandValue(HTMLArea.is_ie ? "backcolor" : "hilitecolor"));
					if (/transparent/i.test(backgroundColor)) {
						backgroundColor = HTMLArea._makeColor(doc.queryCommandValue("backcolor"));
					}
					color = HTMLArea._makeColor(doc.queryCommandValue("forecolor"));
					fontFamily = doc.queryCommandValue("fontname");
					fontWeight = doc.queryCommandState("bold") ? "bold" : "normal";
					fontStyle = doc.queryCommandState("italic") ? "italic" : "normal";
				}} catch (e) {
				}
			}
			break;
		    case "htmlmode": btn.state("active", text); break;
		    case "lefttoright":
		    case "righttoleft":
			var el = this.getParentElement();
			while (el && !HTMLArea.isBlockElement(el))
				el = el.parentNode;
			if (el)
				btn.state("active", (el.style.direction == ((cmd == "righttoleft") ? "rtl" : "ltr")));
			break;
		    default:
			try {
				btn.state("active", (!text && doc.queryCommandState(cmd)));
			} catch (e) {}
		}
	}

	if (this._customUndo && !this._timerUndo) {
		this._undoTakeSnapshot();
		var editor = this;
		this._timerUndo = setTimeout(function() {
			editor._timerUndo = null;
		}, this.config.undoTimeout);
	}
};

HTMLArea.prototype.insertNodeAtSelection = function(toBeInserted) {
	if (!HTMLArea.is_ie) {
		var sel = this._getSelection();
		var range = this._createRange(sel);

		sel.removeAllRanges();
		range.deleteContents();
		var node = range.startContainer;
		var pos = range.startOffset;
		switch (node.nodeType) {
		    case 3:
			if (toBeInserted.nodeType == 3) {
				node.insertData(pos, toBeInserted.data);
				range = this._createRange();
				range.setEnd(node, pos + toBeInserted.length);
				range.setStart(node, pos + toBeInserted.length);
				sel.addRange(range);
			} else {
				node = node.splitText(pos);
				var selnode = toBeInserted;
				if (toBeInserted.nodeType == 11) {
					selnode = selnode.firstChild;
				}
				node.parentNode.insertBefore(toBeInserted, node);
				this.selectNodeContents(selnode);
				this.updateToolbar();
			}
			break;
		    case 1:
			var selnode = toBeInserted;
			if (toBeInserted.nodeType == 11) {
				selnode = selnode.firstChild;
			}
			node.insertBefore(toBeInserted, node.childNodes[pos]);
			this.selectNodeContents(selnode);
			this.updateToolbar();
			break;
		}
	} else {
		return null;
	}
};

HTMLArea.prototype.getParentElement = function() {
	var sel = this._getSelection();
	var range = this._createRange(sel);
	if (HTMLArea.is_ie) {
		switch (sel.type) {
		    case "Text":
		    case "None":
			return range.parentElement();
		    case "Control":
			return range.item(0);
		    default:
			return this._doc.body;
		}
	} else try {
		var p = range.commonAncestorContainer;
		if (!range.collapsed && range.startContainer == range.endContainer &&
		    range.startOffset - range.endOffset <= 1 && range.startContainer.hasChildNodes())
			p = range.startContainer.childNodes[range.startOffset];
		while (p.nodeType == 3) {
			p = p.parentNode;
		}
		return p;
	} catch (e) {
		return null;
	}
};

HTMLArea.prototype.getAllAncestors = function() {
	var p = this.getParentElement();
	var a = [];
	while (p && (p.nodeType == 1) && (p.tagName.toLowerCase() != 'body')) {
		a.push(p);
		p = p.parentNode;
	}
	a.push(this._doc.body);
	return a;
};

HTMLArea.prototype.selectNodeContents = function(node, pos) {
	this.focusEditor();
	this.forceRedraw();
	var range;
	var collapsed = (typeof pos != "undefined");
	if (HTMLArea.is_ie) {
		range = this._doc.body.createTextRange();
		range.moveToElementText(node);
		(collapsed) && range.collapse(pos);
		range.select();
	} else {
		var sel = this._getSelection();
		range = this._doc.createRange();
		range.selectNodeContents(node);
		(collapsed) && range.collapse(pos);
		sel.removeAllRanges();
		sel.addRange(range);
	}
};

HTMLArea.prototype.insertHTML = function(html) {
	var sel = this._getSelection();
	var range = this._createRange(sel);
	if (HTMLArea.is_ie) {
		range.pasteHTML(html);
	} else {
		var fragment = this._doc.createDocumentFragment();
		var div = this._doc.createElement("div");
		div.innerHTML = html;
		while (div.firstChild) {
			fragment.appendChild(div.firstChild);
		}
		var node = this.insertNodeAtSelection(fragment);
	}
};

HTMLArea.prototype.surroundHTML = function(startTag, endTag) {
	var html = this.getSelectedHTML();
	this.insertHTML(startTag + html + endTag);
};

HTMLArea.prototype.getSelectedHTML = function() {
	var sel = this._getSelection();
	var range = this._createRange(sel);
	var existing = null;
	if (HTMLArea.is_ie) {
		existing = range.htmlText;
	} else {
		existing = HTMLArea.getHTML(range.cloneContents(), false, this);
	}
	return existing;
};

HTMLArea.prototype.hasSelectedText = function() {
	return this.getSelectedHTML() != '';
};

HTMLArea.prototype._createLink = function(link) {
	var editor = this;
	var outparam = null;
	if (typeof link == "undefined") {
		link = this.getParentElement();
		if (link && !/^a$/i.test(link.tagName))
			link = null;
	}
	if (link) outparam = {
		f_href   : HTMLArea.is_ie ? editor.stripBaseURL(link.href) : link.getAttribute("href"),
		f_title  : link.title,
		f_target : link.target
	};
	this._popupDialog("link.html", function(param) {
		if (!param)
			return false;
		var a = link;
		if (!a) {
			editor._doc.execCommand("createlink", false, param.f_href);
			a = editor.getParentElement();
			var sel = editor._getSelection();
			var range = editor._createRange(sel);
			if (!HTMLArea.is_ie) {
				a = range.startContainer;
				if (!/^a$/i.test(a.tagName))
					a = a.nextSibling;
			}
		} else a.href = param.f_href.trim();
		if (!/^a$/i.test(a.tagName))
			return false;
		a.target = param.f_target.trim();
		a.title = param.f_title.trim();
		editor.selectNodeContents(a);
		editor.updateToolbar();
	}, outparam);
};

HTMLArea.prototype._insertImage = function(image) {
	var editor = this;
	var outparam = null;
	if (typeof image == "undefined") {
		image = this.getParentElement();
		if (image && !/^img$/i.test(image.tagName))
			image = null;
	}
	if (image) outparam = {
		f_url    : HTMLArea.is_ie ? editor.stripBaseURL(image.src) : image.getAttribute("src"),
		f_alt    : image.alt,
		f_border : image.border,
		f_align  : image.align,
		f_vert   : image.vspace,
		f_horiz  : image.hspace
	};
	this._popupDialog("insert_image.html", function(param) {
		if (!param) {
			return false;
		}
		var img = image;
		if (!img) {
			var sel = editor._getSelection();
			var range = editor._createRange(sel);
			editor._doc.execCommand("insertimage", false, param.f_url);
			if (HTMLArea.is_ie) {
				img = range.parentElement();
				if (img.tagName.toLowerCase() != "img") {
					img = img.previousSibling;
				}
			} else {
				img = range.startContainer.previousSibling;
			}
		} else {
			img.src = param.f_url;
		}
		for (field in param) {
			var value = param[field];
			switch (field) {
			    case "f_alt"    : img.alt	 = value; break;
			    case "f_border" : img.border = parseInt(value || "0"); break;
			    case "f_align"  : img.align	 = value; break;
			    case "f_vert"   : img.vspace = parseInt(value || "0"); break;
			    case "f_horiz"  : img.hspace = parseInt(value || "0"); break;
			}
		}
	}, outparam);
};

HTMLArea.prototype._insertTable = function() {
	var sel = this._getSelection();
	var range = this._createRange(sel);
	var editor = this;
	this._popupDialog("insert_table.html", function(param) {
		if (!param) {
			return false;
		}
		var doc = editor._doc;
		var table = doc.createElement("table");
		for (var field in param) {
			var value = param[field];
			if (!value) {
				continue;
			}
			switch (field) {
			    case "f_width"   : table.style.width = value + param["f_unit"]; break;
			    case "f_align"   : table.align	 = value; break;
			    case "f_border"  : table.border	 = parseInt(value); break;
			    case "f_spacing" : table.cellspacing = parseInt(value); break;
			    case "f_padding" : table.cellpadding = parseInt(value); break;
			}
		}
		var tbody = doc.createElement("tbody");
		table.appendChild(tbody);
		for (var i = 0; i < param["f_rows"]; ++i) {
			var tr = doc.createElement("tr");
			tbody.appendChild(tr);
			for (var j = 0; j < param["f_cols"]; ++j) {
				var td = doc.createElement("td");
				tr.appendChild(td);
				(HTMLArea.is_gecko) && td.appendChild(doc.createElement("br"));
			}
		}
		if (HTMLArea.is_ie) {
			range.pasteHTML(table.outerHTML);
		} else {
			editor.insertNodeAtSelection(table);
		}
		return true;
	}, null);
};

HTMLArea.prototype._comboSelected = function(el, txt) {
	this.focusEditor();
	var value = el.options[el.selectedIndex].value;
	switch (txt) {
	    case "fontname":
	    case "fontsize": this.execCommand(txt, false, value); break;
	    case "formatblock":
		(HTMLArea.is_ie) && (value = "<" + value + ">");
		this.execCommand(txt, false, value);
		break;
	    default:
		var dropdown = this.config.customSelects[txt];
		if (typeof dropdown != "undefined") {
			dropdown.action(this);
		} else {
			alert("FIXME: combo box " + txt + " not implemented");
		}
	}
};

HTMLArea.prototype.execCommand = function(cmdID, UI, param) {
	var editor = this;
	this.focusEditor();
	cmdID = cmdID.toLowerCase();
	switch (cmdID) {
	    case "htmlmode" : this.setMode(); break;
	    case "hilitecolor":
		(HTMLArea.is_ie) && (cmdID = "backcolor");
	    case "forecolor":
		this._popupDialog("select_color.html", function(color) {
			if (color) {
				editor._doc.execCommand(cmdID, false, "#" + color);
			}
		}, HTMLArea._colorToRgb(this._doc.queryCommandValue(cmdID)));
		break;
	    case "createlink":
		this._createLink();
		break;
	    case "popupeditor":

		HTMLArea._object = this;
		if (HTMLArea.is_ie) {
			{
				window.open(this.popupURL("fullscreen.html"), "ha_fullscreen",
					    "toolbar=no,location=no,directories=no,status=no,menubar=no," +
					    "scrollbars=no,resizable=yes,width=640,height=480");
			}
		} else {
			window.open(this.popupURL("fullscreen.html"), "ha_fullscreen",
				    "toolbar=no,menubar=no,personalbar=no,width=640,height=480," +
				    "scrollbars=no,resizable=yes");
		}
		break;
	    case "undo":
	    case "redo":
		if (this._customUndo)
			this[cmdID]();
		else
			this._doc.execCommand(cmdID, UI, param);
		break;
	    case "inserttable": this._insertTable(); break;
	    case "insertimage": this._insertImage(); break;
	    case "killword": this._wordClean(); break;
	    case "cut":
	    case "copy":
	    case "paste":
		try {
			if (this.config.killWordOnPaste)
				this._wordClean();
			this._doc.execCommand(cmdID, UI, param);
		} catch (e) {
			if (HTMLArea.is_gecko) {
				if (confirm("Unprivileged scripts cannot access Cut/Copy/Paste programatically " +
					    "for security reasons.  Click OK to see a technical note at mozilla.org " +
					    "which shows you how to allow a script to access the clipboard."))
					window.open("http://mozilla.org/editor/midasdemo/securityprefs.html");
			}
		}
		break;
	    case "lefttoright":
	    case "righttoleft":
		var dir = (cmdID == "righttoleft") ? "rtl" : "ltr";
		var el = this.getParentElement();
		while (el && !HTMLArea.isBlockElement(el))
			el = el.parentNode;
		if (el) {
			if (el.style.direction == dir)
				el.style.direction = "";
			else
				el.style.direction = dir;
		}
		break;
	    default: this._doc.execCommand(cmdID, UI, param);
	}
	this.updateToolbar();
	return false;
};

HTMLArea.prototype._editorEvent = function(ev) {
	var editor = this;
	var keyEvent = (HTMLArea.is_ie && ev.type == "keydown") || (ev.type == "keypress");
	if (keyEvent && ev.ctrlKey) {
		var sel = null;
		var range = null;
		var key = String.fromCharCode(HTMLArea.is_ie ? ev.keyCode : ev.charCode).toLowerCase();
		var cmd = null;
		var value = null;
		switch (key) {
		    case 'a':
			if (!HTMLArea.is_ie) {
				sel = this._getSelection();
				sel.removeAllRanges();
				range = this._createRange();
				range.selectNodeContents(this._doc.body);
				sel.addRange(range);
				HTMLArea._stopEvent(ev);
			}
			break;

		    case 'b': cmd = "bold"; break;
		    case 'i': cmd = "italic"; break;
		    case 'u': cmd = "underline"; break;
		    case 's': cmd = "strikethrough"; break;
		    case 'l': cmd = "justifyleft"; break;
		    case 'e': cmd = "justifycenter"; break;
		    case 'r': cmd = "justifyright"; break;
		    case 'j': cmd = "justifyfull"; break;
		    case 'z': cmd = "undo"; break;
		    case 'y': cmd = "redo"; break;
		    case 'v': cmd = "paste"; break;

		    case '0': cmd = "killword"; break;

		    case '1':
		    case '2':
		    case '3':
		    case '4':
		    case '5':
		    case '6':
			cmd = "formatblock";
			value = "h" + key;
			if (HTMLArea.is_ie) {
				value = "<" + value + ">";
			}
			break;
		}
		if (cmd) {
			this.execCommand(cmd, false, value);
			HTMLArea._stopEvent(ev);
		}
	}

	if (editor._timerToolbar) {
		clearTimeout(editor._timerToolbar);
	}
	editor._timerToolbar = setTimeout(function() {
		editor.updateToolbar();
		editor._timerToolbar = null;
	}, 50);
};

HTMLArea.prototype.getHTML = function() {
	switch (this._editMode) {
	    case "wysiwyg"  :
		if (!this.config.fullPage) {
			return HTMLArea.getHTML(this._doc.body, false, this);
		} else
			return this.doctype + "\n" + HTMLArea.getHTML(this._doc.documentElement, true, this);
	    case "textmode" : return this._textArea.value;
	    default	    : alert("Mode <" + mode + "> not defined!");
	}
	return false;
};

HTMLArea.prototype.getInnerHTML = function() {
	switch (this._editMode) {
	    case "wysiwyg"  :
		if (!this.config.fullPage)
			return this._doc.body.innerHTML;
		else
			return this.doctype + "\n" + this._doc.documentElement.innerHTML;
	    case "textmode" : return this._textArea.value;
	    default	    : alert("Mode <" + mode + "> not defined!");
	}
	return false;
};

HTMLArea.prototype.setHTML = function(html) {
	switch (this._editMode) {
	    case "wysiwyg"  :
		if (!this.config.fullPage)
			this._doc.body.innerHTML = html;
		else
			this._doc.body.innerHTML = html;
		break;
	    case "textmode" : this._textArea.value = html; break;
	    default	    : alert("Mode <" + mode + "> not defined!");
	}
	return false;
};

HTMLArea.prototype.setDoctype = function(doctype) {
	this.doctype = doctype;
};

HTMLArea.agt = navigator.userAgent.toLowerCase();
HTMLArea.is_ie	   = ((HTMLArea.agt.indexOf("msie") != -1) && (HTMLArea.agt.indexOf("opera") == -1));
HTMLArea.is_opera  = (HTMLArea.agt.indexOf("opera") != -1);
HTMLArea.is_mac	   = (HTMLArea.agt.indexOf("mac") != -1);
HTMLArea.is_mac_ie = (HTMLArea.is_ie && HTMLArea.is_mac);
HTMLArea.is_win_ie = (HTMLArea.is_ie && !HTMLArea.is_mac);
HTMLArea.is_gecko  = (navigator.product == "Gecko");

HTMLArea._object = null;

HTMLArea.cloneObject = function(obj) {
	var newObj = new Object;

	if (obj.constructor.toString().indexOf("function Array(") == 1) {
		newObj = obj.constructor();
	}

	if (obj.constructor.toString().indexOf("function Function(") == 1) {
		newObj = obj;
	} else for (var n in obj) {
		var node = obj[n];
		if (typeof node == 'object') { newObj[n] = HTMLArea.cloneObject(node); }
		else                         { newObj[n] = node; }
	}

	return newObj;
};

HTMLArea.checkSupportedBrowser = function() {
	if (HTMLArea.is_gecko) {
		if (navigator.productSub < 20021201) {
			alert("You need at least Mozilla-1.3 Alpha.\n" +
			      "Sorry, your Gecko is not supported.");
			return false;
		}
		if (navigator.productSub < 20030210) {
			alert("Mozilla < 1.3 Beta is not supported!\n" +
			      "I'll try, though, but it might not work.");
		}
	}
	return HTMLArea.is_gecko || HTMLArea.is_ie;
};

HTMLArea.prototype._getSelection = function() {
	if (HTMLArea.is_ie) {
		return this._doc.selection;
	} else {
		return this._iframe.contentWindow.getSelection();
	}
};

HTMLArea.prototype._createRange = function(sel) {
	if (HTMLArea.is_ie) {
		return sel.createRange();
	} else {
		this.focusEditor();
		if (typeof sel != "undefined") {
			try {
				return sel.getRangeAt(0);
			} catch(e) {
				return this._doc.createRange();
			}
		} else {
			return this._doc.createRange();
		}
	}
};

HTMLArea._addEvent = function(el, evname, func) {
	if (HTMLArea.is_ie) {
		el.attachEvent("on" + evname, func);
	} else {
		el.addEventListener(evname, func, true);
	}
};

HTMLArea._addEvents = function(el, evs, func) {
	for (var i in evs) {
		HTMLArea._addEvent(el, evs[i], func);
	}
};

HTMLArea._removeEvent = function(el, evname, func) {
	if (HTMLArea.is_ie) {
		el.detachEvent("on" + evname, func);
	} else {
		el.removeEventListener(evname, func, true);
	}
};

HTMLArea._removeEvents = function(el, evs, func) {
	for (var i in evs) {
		HTMLArea._removeEvent(el, evs[i], func);
	}
};

HTMLArea._stopEvent = function(ev) {
	if (HTMLArea.is_ie) {
		ev.cancelBubble = true;
		ev.returnValue = false;
	} else {
		ev.preventDefault();
		ev.stopPropagation();
	}
};

HTMLArea._removeClass = function(el, className) {
	if (!(el && el.className)) {
		return;
	}
	var cls = el.className.split(" ");
	var ar = new Array();
	for (var i = cls.length; i > 0;) {
		if (cls[--i] != className) {
			ar[ar.length] = cls[i];
		}
	}
	el.className = ar.join(" ");
};

HTMLArea._addClass = function(el, className) {
	HTMLArea._removeClass(el, className);
	el.className += " " + className;
};

HTMLArea._hasClass = function(el, className) {
	if (!(el && el.className)) {
		return false;
	}
	var cls = el.className.split(" ");
	for (var i = cls.length; i > 0;) {
		if (cls[--i] == className) {
			return true;
		}
	}
	return false;
};

HTMLArea.isBlockElement = function(el) {
	var blockTags = " body form textarea fieldset ul ol dl li div " +
		"p h1 h2 h3 h4 h5 h6 quote pre table thead " +
		"tbody tfoot tr td iframe address ";
	return (blockTags.indexOf(" " + el.tagName.toLowerCase() + " ") != -1);
};

HTMLArea.needsClosingTag = function(el) {
	var closingTags = " head script style div span tr td tbody table em strong font a title ";
	return (closingTags.indexOf(" " + el.tagName.toLowerCase() + " ") != -1);
};

HTMLArea.htmlEncode = function(str) {
	str = str.replace(/&/ig, "&amp;");
	str = str.replace(/</ig, "&lt;");
	str = str.replace(/>/ig, "&gt;");
	str = str.replace(/\x22/ig, "&quot;");
	return str;
};

HTMLArea.getHTML = function(root, outputRoot, editor) {
	var html = "";
	switch (root.nodeType) {
	    case 1:
	    case 11:
		var closed;
		var i;
		var root_tag = (root.nodeType == 1) ? root.tagName.toLowerCase() : '';
		if (HTMLArea.is_ie && root_tag == "head") {
			if (outputRoot)
				html += "<head>";
			var save_multiline = RegExp.multiline;
			RegExp.multiline = true;
			var txt = root.innerHTML.replace(HTMLArea.RE_tagName, function(str, p1, p2) {
				return p1 + p2.toLowerCase();
			});
			RegExp.multiline = save_multiline;
			html += txt;
			if (outputRoot)
				html += "</head>";
			break;
		} else if (outputRoot) {
			closed = (!(root.hasChildNodes() || HTMLArea.needsClosingTag(root)));
			html = "<" + root.tagName.toLowerCase();
			var attrs = root.attributes;
			for (i = 0; i < attrs.length; ++i) {
				var a = attrs.item(i);
				if (!a.specified) {
					continue;
				}
				var name = a.nodeName.toLowerCase();
				if (/_moz|contenteditable|_msh/.test(name)) {
					continue;
				}
				var value;
				if (name != "style") {
					if (typeof root[a.nodeName] != "undefined" && name != "href" && name != "src") {
						value = root[a.nodeName];
					} else {
						value = a.nodeValue;
						if (HTMLArea.is_ie && (name == "href" || name == "src")) {
							value = editor.stripBaseURL(value);
						}
					}
				} else {
					value = root.style.cssText;
				}
				if (/(_moz|^$)/.test(value)) {
					continue;
				}
				html += " " + name + '="' + value + '"';
			}
			html += closed ? " />" : ">";
		}
		for (i = root.firstChild; i; i = i.nextSibling) {
			html += HTMLArea.getHTML(i, true, editor);
		}
		if (outputRoot && !closed) {
			html += "</" + root.tagName.toLowerCase() + ">";
		}
		break;
	    case 3:
		if ( !root.previousSibling && !root.nextSibling && root.data.match(/^\s*$/i) ) html = '&nbsp;';
		else html = HTMLArea.htmlEncode(root.data);
		break;
	    case 8:
		html = "<!--" + root.data + "-->";
		break;
	}
	return html;
};

HTMLArea.prototype.stripBaseURL = function(string) {
	var baseurl = this.config.baseURL;

	baseurl = baseurl.replace(/[^\/]+$/, '');
	var basere = new RegExp(baseurl);
	string = string.replace(basere, "");

	baseurl = baseurl.replace(/^(https?:\/\/[^\/]+)(.*)$/, '$1');
	basere = new RegExp(baseurl);
	return string.replace(basere, "");
};

String.prototype.trim = function() {
	a = this.replace(/^\s+/, '');
	return a.replace(/\s+$/, '');
};

HTMLArea._makeColor = function(v) {
	if (typeof v != "number") {
		return v;
	}
	var r = v & 0xFF;
	var g = (v >> 8) & 0xFF;
	var b = (v >> 16) & 0xFF;
	return "rgb(" + r + "," + g + "," + b + ")";
};

HTMLArea._colorToRgb = function(v) {
	if (!v)
		return '';

	function hex(d) {
		return (d < 16) ? ("0" + d.toString(16)) : d.toString(16);
	};

	if (typeof v == "number") {
		var r = v & 0xFF;
		var g = (v >> 8) & 0xFF;
		var b = (v >> 16) & 0xFF;
		return "#" + hex(r) + hex(g) + hex(b);
	}

	if (v.substr(0, 3) == "rgb") {
		var re = /rgb\s*\(\s*([0-9]+)\s*,\s*([0-9]+)\s*,\s*([0-9]+)\s*\)/;
		if (v.match(re)) {
			var r = parseInt(RegExp.$1);
			var g = parseInt(RegExp.$2);
			var b = parseInt(RegExp.$3);
			return "#" + hex(r) + hex(g) + hex(b);
		}
		return null;
	}

	if (v.substr(0, 1) == "#") {
		return v;
	}

	return null;
};

HTMLArea.prototype._popupDialog = function(url, action, init) {
	Dialog(this.popupURL(url), action, init);
};

HTMLArea.prototype.imgURL = function(file, plugin) {
	return _editor_url + file;
};

HTMLArea.prototype.popupURL = function(file) {
	return _editor_url + this.config.popupURL + file;
};

HTMLArea.getElementById = function(tag, id) {
	var el, i, objs = document.getElementsByTagName(tag);
	for (i = objs.length; --i >= 0 && (el = objs[i]);)
		if (el.id == id)
			return el;
	return null;
};