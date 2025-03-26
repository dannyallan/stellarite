// Stellarite load script for DHTML Calendar 0.9.6

function showCalendar(id) {
	var el = document.getElementById(id);
	if (calendar != null) {
		calendar.hide();
	} else {
		var cal = new Calendar(0, null, selected, closeHandler);

		cal.weekNumbers = false;
		cal.setDateFormat(sDateFormat);
		cal.yearStep = 1;

	    calendar = cal;
	    cal.create();
	}
	calendar.parseDate(el.value);
	calendar.sel = el;
	calendar.showAtElement(el, "Tr");
	return;
}

function selected(cal, date) {
	cal.sel.value = date;
	if (cal.dateClicked)
		cal.callCloseHandler();
}

function closeHandler(cal) {
	cal.hide();
	calendar = null;
}

function getToday() {
	var dt_today = new Date();
	var st_return = sDateFormat;

	st_return = st_return.replace(/\%Y/g, dt_today.getFullYear());
	st_return = st_return.replace(/\%m/g, dt_today.getMonth()+1);
	st_return = st_return.replace(/\%d/g, dt_today.getDate());

	return st_return;
}