var daterangepickersYMD = [];
var obj_calwindow

function DateRangePickerYMD(obj_sourceform, obj_target, titlebartext, SubmitSourceFormOnClose) {
//function DatePickerYMD_EA(obj_sourceform, obj_target, titlebartext, SubmitSourceFormOnClose) {
	//this.popup = pm_popup2;
	this.popup = pm_popupYMDDateRange
	this.popupclose = pm_popupcloseYMDDateRange;
	this.CloseCalendar = pm_CloseCalendar;
	this.target = obj_target;
	this.obj_sourceform = obj_sourceform;
	this.id = daterangepickersYMD.length;
	daterangepickersYMD[this.id] = this;
	this.titlebartext = titlebartext
	this.SubmitSourceFormOnClose = SubmitSourceFormOnClose	
}

function pm_CloseCalendar()  {
	obj_calwindow.close();
	if (this.SubmitSourceFormOnClose == 1 )  {
		this.obj_sourceform.hdAction.value = "clientselectionchanged"
		this.obj_sourceform.submit();
	}
}

function pm_popupYMDDateRange()  {
	obj_calwindow = window.open('DateRangePickerYMD.aspx?id=' + this.id, 'Calendar', 'width=515,height=300,status=no,resizable=no,top=200,left=200,dependent=yes,alwaysRaised=yes')
	obj_calwindow.opener = window;
	obj_calwindow.focus();
}

function pm_popupcloseYMDDateRange() {
	try {
		if (obj_calwindow != undefined)  {
			obj_calwindow.close()
		}						
	}
	catch (everything) {   }
}