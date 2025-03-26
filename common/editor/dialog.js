// htmlArea v3.0 - Copyright (c) 2002-2003, interactivetools.com, inc.
// All rights reserved.

function Dialog(url, action, init) {
	if (typeof init == "undefined") {
		init = window;
	}
	Dialog._geckoOpenModal(url, action, init);
};

Dialog._parentEvent = function(ev) {
	if (Dialog._modal && !Dialog._modal.closed) {
		Dialog._modal.focus();
		HTMLArea._stopEvent(ev);
	}
};

Dialog._return = null;
Dialog._modal = null;
Dialog._arguments = null;

Dialog._geckoOpenModal = function(url, action, init) {
	var dlg = window.open(url, "hadialog",
			      "toolbar=no,menubar=no,personalbar=no,width=10,height=10," +
			      "scrollbars=no,resizable=yes");
	Dialog._modal = dlg;
	Dialog._arguments = init;

	function capwin(w) {
		try {
			HTMLArea._addEvent(w, "click", Dialog._parentEvent);
			HTMLArea._addEvent(w, "mousedown", Dialog._parentEvent);
			HTMLArea._addEvent(w, "focus", Dialog._parentEvent);
		} catch(e) {}
	};
	function relwin(w) {
		try {
			HTMLArea._removeEvent(w, "click", Dialog._parentEvent);
			HTMLArea._removeEvent(w, "mousedown", Dialog._parentEvent);
			HTMLArea._removeEvent(w, "focus", Dialog._parentEvent);
		} catch(e) {}
	};
	capwin(window);

	for (var i = 0; i < window.frames.length; capwin(window.frames[i++]));
	Dialog._return = function (val) {
		if (val && action) {
			action(val);
		}
		relwin(window);
		for (var i = 0; i < window.frames.length; relwin(window.frames[i++]));
		Dialog._modal = null;
	};
};
