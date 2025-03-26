// ===================================================================
// Author: Matt Kruse <matt@mattkruse.com>
// WWW: http://www.mattkruse.com/
// ===================================================================

function hasOptions(obj) {
	if (obj!=null && obj.options!=null) { return true; }
	return false;
}

function selectAllOptions(obj) {
	if (!hasOptions(obj)) { return; }
	for (var i=0; i<obj.options.length; i++) {
		obj.options[i].selected = true;
	}
}

function moveAllOptions(from,to) {
	selectAllOptions(from);
	moveSelectedOptions(from,to);
}

function moveSelectedOptions(from,to) {
	if (!hasOptions(from)) { return; }
	for (var i=0; i<from.options.length; i++) {
		var o = from.options[i];
		if (o.selected) {
			if (!hasOptions(to)) { var index = 0; } else { var index=to.options.length; }
			to.options[index] = new Option( o.text, o.value, false, false);
		}
	}
	for (var i=(from.options.length-1); i>=0; i--) {
		var o = from.options[i];
		if (o.selected) {
			from.options[i] = null;
		}
	}
	from.selectedIndex = -1;
	to.selectedIndex = -1;
	doChange();
}

function moveOptionTop(obj) {
	if (!hasOptions(obj)) { return; }

	var selIndex = obj.selectedIndex;
	obj.selectedIndex = -1;

	for (i=selIndex-1; i>=0; i--) {
		obj.options[i].selected = true;
	}
	moveOptionDown(obj);
	obj.selectedIndex = 0;
}

function moveOptionUp(obj) {
	if (!hasOptions(obj)) { return; }
	for (i=0; i<obj.options.length; i++) {
		if (obj.options[i].selected) {
			if (i != 0 && !obj.options[i-1].selected) {
				swapOptions(obj,i,i-1);
				obj.options[i-1].selected = true;
			}
		}
	}
}

function moveOptionBottom(obj) {
	if (!hasOptions(obj)) { return; }

	var selIndex = obj.selectedIndex;
	obj.selectedIndex = -1;

	for (i=selIndex+1; i<obj.options.length; i++) {
		obj.options[i].selected = true;
	}
	moveOptionUp(obj);
	obj.selectedIndex = obj.options.length-1;
}

function moveOptionDown(obj) {
	if (!hasOptions(obj)) { return; }
	for (i=obj.options.length-1; i>=0; i--) {
		if (obj.options[i].selected) {
			if (i != (obj.options.length-1) && ! obj.options[i+1].selected) {
				swapOptions(obj,i,i+1);
				obj.options[i+1].selected = true;
			}
		}
	}
}

function swapOptions(obj,i,j) {
	var o = obj.options;
	var i_selected = o[i].selected;
	var j_selected = o[j].selected;
	var temp = new Option(o[i].text, o[i].value, o[i].defaultSelected, o[i].selected);
	var temp2= new Option(o[j].text, o[j].value, o[j].defaultSelected, o[j].selected);
	o[i] = temp2;
	o[j] = temp;
	o[i].selected = j_selected;
	o[j].selected = i_selected;
	doChange();
}