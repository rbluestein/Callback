<%@ Page Language="VB" AutoEventWireup="false" CodeFile="DateRangePickerYMD.aspx.vb" Inherits="DateRangePickerYMD" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<title>Select Date</title>
	<%--	<link title="BVIStyle" href="JSControls.css" type="text/css" rel="stylesheet" />--%>
        <link href="css/JSControls.css" rel="stylesheet" />
	</head>
	<body onload="fnDateRangePickerYMD_Load()" onunload="fnClose()">
		<form id="form2" method="post" runat="server">
			<input type='hidden' id='hdCurMonth' /><input type='hidden' id='hdCurYear' /><input type='hidden' id='hdInitial' value='1' />
			<input type='hidden' id='hdFromDay' /><input type='hidden' id='hdToDay' /> <input type='hidden' id='hdFromMonth' /><input type='hidden' id='hdFromYear' /><input type='hidden' id='hdToMonth' /><input type='hidden' id='hdToYear' />
			<table cellpadding="0" style="LEFT: 10px; POSITION: absolute; TOP: 0px">
				<tr>
					<td style="FONT: bold 13px Verdana">From Date</td>
				</tr>
			</table>
			<table class="MPtable" cellpadding="0" style="LEFT: 10px; POSITION: absolute; TOP: 20px">
				<tr>
					<td class="MPtd"><a href="javascript:fnSelectYear(0, -1)">Prev Year</a></td>
					<td class="MPtd"><a href="javascript:fnSelectYear(0, 1)">Next Year</a></td>
				</tr>
			</table>
			<table class="MPtable" cellpadding="0" style="LEFT: 10px; POSITION: absolute; TOP: 43px">
				<tr>
					<td class="MPtd"><a href="javascript:fnSelectMonth('0','0')">Jan</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('0','1')">Feb</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('0','2')">Mar</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('0','3')">Apr</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('0','4')">May</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('0','5')">Jun</a></td>
				</tr>
				<tr>
					<td class="MPtd"><a href="javascript:fnSelectMonth('0','6')">Jul</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('0','7')">Aug</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('0','8')">Sep</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('0','9')">Oct</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('0','10')">Nov</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('0','11')">Dec</a></td>
				</tr>
			</table>
			<table class="MPtable" cellpadding="0" style="LEFT: 10px; POSITION: absolute; TOP: 86px">
				<tr>
					<td class="MPtd">Sun</td>
					<td class="MPtd">Mon</td>
					<td class="MPtd">Tue</td>
					<td class="MPtd">Wed</td>
					<td class="MPtd">Thu</td>
					<td class="MPtd">Fri</td>
					<td class="MPtd">Sat</td>
				</tr>
				<tr>
					<td id="F1" class="MPtd">&nbsp;</td>
					<td id="F2" class="MPtd">&nbsp;</td>
					<td id="F3" class="MPtd">&nbsp;</td>
					<td id="F4" class="MPtd">&nbsp;</td>
					<td id="F5" class="MPtd">&nbsp;</td>
					<td id="F6" class="MPtd">&nbsp;</td>
					<td id="F7" class="MPtd">&nbsp;</td>
				</tr>
				<tr>
					<td id="F8" class="MPtd">&nbsp;</td>
					<td id="F9" class="MPtd">&nbsp;</td>
					<td id="F10" class="MPtd">&nbsp;</td>
					<td id="F11" class="MPtd">&nbsp;</td>
					<td id="F12" class="MPtd">&nbsp;</td>
					<td id="F13" class="MPtd">&nbsp;</td>
					<td id="F14" class="MPtd">&nbsp;</td>
				</tr>
				<tr>
					<td id="F15" class="MPtd">&nbsp;</td>
					<td id="F16" class="MPtd">&nbsp;</td>
					<td id="F17" class="MPtd">&nbsp;</td>
					<td id="F18" class="MPtd">&nbsp;</td>
					<td id="F19" class="MPtd">&nbsp;</td>
					<td id="F20" class="MPtd">&nbsp;</td>
					<td id="F21" class="MPtd">&nbsp;</td>
				</tr>
				<tr>
					<td id="F22" class="MPtd">&nbsp;</td>
					<td id="F23" class="MPtd">&nbsp;</td>
					<td id="F24" class="MPtd">&nbsp;</td>
					<td id="F25" class="MPtd">&nbsp;</td>
					<td id="F26" class="MPtd">&nbsp;</td>
					<td id="F27" class="MPtd">&nbsp;</td>
					<td id="F28" class="MPtd">&nbsp;</td>
				</tr>
				<tr>
					<td id="F29" class="MPtd">&nbsp;</td>
					<td id="F30" class="MPtd">&nbsp;</td>
					<td id="F31" class="MPtd">&nbsp;</td>
					<td id="F32" class="MPtd">&nbsp;</td>
					<td id="F33" class="MPtd">&nbsp;</td>
					<td id="F34" class="MPtd">&nbsp;</td>
					<td id="F35" class="MPtd">&nbsp;</td>
				</tr>
				<tr id="FromDateLastRow">
					<td id="F36" class="MPtd">&nbsp;</td>
					<td id="F37" class="MPtd">&nbsp;</td>
					<td id="F38" class="MPtd">&nbsp;</td>
					<td id="F39" class="MPtd">&nbsp;</td>
					<td id="F40" class="MPtd">&nbsp;</td>
					<td id="F41" class="MPtd">&nbsp;</td>
					<td id="F42" class="MPtd">&nbsp;</td>
				</tr>
			</table>
			<table cellpadding="2" cellspacing="0" style="LEFT: 10px; POSITION: absolute; TOP: 222px">
				<tr>
					<td><input type="text" id="txtFromMonthDayYear" name="txtFromMonthDayYear" readonly></td>
				</tr>
			</table>
			<table class="MPtable" cellpadding="0" style="LEFT: 10px; POSITION: absolute; TOP: 249px">
				<tr>
					<td class="MPtd"><a href="javascript:fnClear(0)">Clear</a></td>
				</tr>
			</table>
			<table cellpadding="0" style="LEFT: 265px; POSITION: absolute; TOP: 0px">
				<tr>
					<td style="FONT: bold 13px Verdana">To Date</td>
				</tr>
			</table>
			<table class="MPtable" cellpadding="0" style="LEFT: 265px; POSITION: absolute; TOP: 20px">
				<tr>
					<td class="MPtd"><a href="javascript:fnSelectYear(1, -1)">Prev Year</a></td>
					<td class="MPtd"><a href="javascript:fnSelectYear(1, 1)">Next Year</a></td>
				</tr>
			</table>
			<table class="MPtable" cellpadding="0" style="LEFT: 265px; POSITION: absolute; TOP: 43px">
				<tr>
					<td class="MPtd"><a href="javascript:fnSelectMonth('1','0')">Jan</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('1','1')">Feb</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('1','2')">Mar</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('1','3')">Apr</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('1','4')">May</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('1','5')">Jun</a></td>
				</tr>
				<tr>
					<td class="MPtd"><a href="javascript:fnSelectMonth('1','6')">Jul</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('1','7')">Aug</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('1','8')">Sep</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('1','9')">Oct</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('1','10')">Nov</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('1','11')">Dec</a></td>
				</tr>
			</table>
			<table class="MPtable" cellpadding="0" style="LEFT: 265px; POSITION: absolute; TOP: 86px">
				<tr>
					<td class="MPtd">Sun</td>
					<td class="MPtd">Mon</td>
					<td class="MPtd">Tue</td>
					<td class="MPtd">Wed</td>
					<td class="MPtd">Thu</td>
					<td class="MPtd">Fri</td>
					<td class="MPtd">Sat</td>
				</tr>
				<tr>
					<td id="T1" class="MPtd">&nbsp;</td>
					<td id="T2" class="MPtd">&nbsp;</td>
					<td id="T3" class="MPtd">&nbsp;</td>
					<td id="T4" class="MPtd">&nbsp;</td>
					<td id="T5" class="MPtd">&nbsp;</td>
					<td id="T6" class="MPtd">&nbsp;</td>
					<td id="T7" class="MPtd">&nbsp;</td>
				</tr>
				<tr>
					<td id="T8" class="MPtd">&nbsp;</td>
					<td id="T9" class="MPtd">&nbsp;</td>
					<td id="T10" class="MPtd">&nbsp;</td>
					<td id="T11" class="MPtd">&nbsp;</td>
					<td id="T12" class="MPtd">&nbsp;</td>
					<td id="T13" class="MPtd">&nbsp;</td>
					<td id="T14" class="MPtd">&nbsp;</td>
				</tr>
				<tr>
					<td id="T15" class="MPtd">&nbsp;</td>
					<td id="T16" class="MPtd">&nbsp;</td>
					<td id="T17" class="MPtd">&nbsp;</td>
					<td id="T18" class="MPtd">&nbsp;</td>
					<td id="T19" class="MPtd">&nbsp;</td>
					<td id="T20" class="MPtd">&nbsp;</td>
					<td id="T21" class="MPtd">&nbsp;</td>
				</tr>
				<tr>
					<td id="T22" class="MPtd">&nbsp;</td>
					<td id="T23" class="MPtd">&nbsp;</td>
					<td id="T24" class="MPtd">&nbsp;</td>
					<td id="T25" class="MPtd">&nbsp;</td>
					<td id="T26" class="MPtd">&nbsp;</td>
					<td id="T27" class="MPtd">&nbsp;</td>
					<td id="T28" class="MPtd">&nbsp;</td>
				</tr>
				<tr>
					<td id="T29" class="MPtd">&nbsp;</td>
					<td id="T30" class="MPtd">&nbsp;</td>
					<td id="T31" class="MPtd">&nbsp;</td>
					<td id="T32" class="MPtd">&nbsp;</td>
					<td id="T33" class="MPtd">&nbsp;</td>
					<td id="T34" class="MPtd">&nbsp;</td>
					<td id="T35" class="MPtd">&nbsp;</td>
				</tr>
				<tr id="ToDateLastRow">
					<td id="T36" class="MPtd">&nbsp;</td>
					<td id="T37" class="MPtd">&nbsp;</td>
					<td id="T38" class="MPtd">&nbsp;</td>
					<td id="T39" class="MPtd">&nbsp;</td>
					<td id="T40" class="MPtd">&nbsp;</td>
					<td id="T41" class="MPtd">&nbsp;</td>
					<td id="T42" class="MPtd">&nbsp;</td>
				</tr>
			</table>
			<table cellpadding="0" style="LEFT: 265px; POSITION: absolute; TOP: 222px">
				<tr>
					<td><input type="text" id="txtToMonthDayYear" name="txtToMonthDayYear" readonly="readonly" /></td>
				</tr>
			</table>
			<table class="MPtable" cellpadding="0" style="LEFT: 265px; POSITION: absolute; TOP: 249px">
				<tr>
					<td class="MPtd"><a href="javascript:fnClear(1)">Clear</a></td>
				</tr>
			</table>
			<table class="MPtable" cellpadding="0" style="LEFT: 10px; POSITION: absolute; TOP: 272px">
				<tr>
					<td class="MPtd"><a href="javascript:fnClose()">Done</a></td>
				</tr>
			</table>
			<script type="text/javascript">
							var re_id = new RegExp('id=(\\d+)');
							var num_id = (re_id.exec(String(window.location))  ? new Number(RegExp.$1) : 0);
							var obj_caller = (window.opener ? window.opener.daterangepickersYMD[num_id] : null);
							var ARR_MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];		
						
							function fnDateRangePickerYMD_Load()  {
								fnDateRangePickerYMD_Init()
							}
							
							function fnDateRangePickerYMD_Init()  {	
									
								document.title = obj_caller.titlebartext					
								//var ARR_MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];		
								//  var ShowDate =  String(ARR_MONTHS[form2.hdFromMonth.value]) + '&nbsp;' + String(form2.hdFromYear.value)			 not used					
								var ArrValues = obj_caller.target.value.split("|")										
								if (ArrValues[0].length == 0) {
									var d = new Date()
									form2.hdFromYear.value = d.getFullYear()		
									form2.hdFromMonth.value = d.getMonth()
									form2.hdFromDay.value = d.getDate()			
									form2.hdToYear.value = d.getFullYear()			
									form2.hdToMonth.value = d.getMonth()	
									form2.hdToDay.value =  d.getDate()
								} else {									
									form2.hdFromYear.value = ArrValues[0]		
									form2.hdFromMonth.value = ArrValues[1] - 1
									form2.hdFromDay.value = ArrValues[2]							
									form2.hdToYear.value = ArrValues[3]	
									form2.hdToMonth.value = ArrValues[4]	- 1
									form2.hdToDay.value = ArrValues[5]																							
								}							
								form2.hdInitial.value = "0"	
								fnAdjustCalendar(0)	
								fnAdjustCalendar(1)											
								fnShowDate(0)
								fnShowDate(1)					
							}
								
							// Methods			
							
							
							// NOTES //
							// <body onunload="fnClose()">
							// <td class="MPtd"><a href="javascript:fnClose()">Done</a></td>	
								
								function fnClose()  {
									fnSetSourceControl()
									window.close();
								}			

								function fnShowDate(vSelect)  {
									if (vSelect == 0) {
										if (form2.hdFromMonth.value == "" && form2.hdFromYear.value == "") {
											form2.txtFromMonthDayYear.value = ""
										} else {
											//form2.txtFromMonthDayYear.value  = ARR_MONTHS[form2.hdFromMonth.value] + ' ' + form2.hdFromYear.value
											if (form2.hdFromDay.value == "")  {
												form2.hdFromDay.value = 1
											}	
											if (form2.hdFromMonth.value == "NaN")  {
												form2.txtFromMonthDayYear.value = ""
											} else {
												form2.txtFromMonthDayYear.value  = ARR_MONTHS[form2.hdFromMonth.value] + ' ' + form2.hdFromDay.value + ' ' + form2.hdFromYear.value
											}	
										}
									}
									
									if (vSelect == 1) {
										if (form2.hdToMonth.value == "" && form2.hdToYear.value == "") {
											form2.txtToMonthDayYear.value = ""
										} else {
											if (form2.hdToDay.value == "")  {
												form2.hdToDay.value = 1
											}		
											if (form2.hdToMonth.value == "NaN")  {
												form2.txtToMonthDayYear.value = ""
											} else {						
												form2.txtToMonthDayYear.value  = ARR_MONTHS[form2.hdToMonth.value] + ' ' + form2.hdToDay.value + ' ' + form2.hdToYear.value	
											}
										}
									}					
								
								}	
																		
								function fnSelectMonth(vSelect, vValue)  {
									fnAdjustYear(vSelect)			
									if (vSelect == 0)  {				
										form2.hdFromMonth.value = vValue
									} else {
										form2.hdToMonth.value = vValue				
									}			
									fnShowDate(vSelect)
									fnAdjustCalendar(vSelect)		
									//fnSetSourceControl()
								}
											
								function fnSelectYear(vSelect, vValue)  {
									fnAdjustMonth(vSelect)				
									if (vSelect == 0) {	
										if (form2.hdFromYear.value == '')  {
											fnAdjustYear(0)			
										}  else {
											form2.hdFromYear.value = parseInt(form2.hdFromYear.value) + vValue										
										}				
									}	
									if (vSelect == 1) {	
										if (form2.hdToYear.value == '')  {
											fnAdjustYear(1)			
										}  else {
											form2.hdToYear.value = parseInt(form2.hdToYear.value) + vValue										
										}			
									}						
									fnShowDate(vSelect)	
									fnAdjustCalendar(vSelect)				
									//fnSetSourceControl()			
								}
								
								function fnSelectDay(vSelect, vValue) {
									if (vSelect == 0)  {				
										form2.hdFromDay.value = vValue
									} else {
										form2.hdToDay.value = vValue				
									}	
									fnShowDate(vSelect)
								}	
								
									
								
				
								
								function fnClear(vSelect)  {
									if (vSelect == 0)  {		
										form2.hdFromMonth.value = ''
										form2.hdFromYear.value = ''	
									} else {		
										form2.hdToMonth.value = ''
										form2.hdToYear.value = ''	
									}
									fnShowDate(vSelect)		
									fnAdjustCalendar(vSelect)				
								}	
								
								function fnAdjustMonth(vSelect) 	{
									if (vSelect == 0  && form2.hdFromMonth.value == "")	{
										form2.hdFromMonth.value = 0
									} 
									if (vSelect == 1 && form2.hdToMonth.value == '')	{
										form2.hdToMonth.value = 0
									}												
								}
								
								function fnAdjustYear(vSelect)  {
									if (vSelect == 0 && form2.hdFromYear.value == '') {
											var d = new Date()
											form2.hdFromYear.value = d.getFullYear()
									}	
									if  (vSelect == 1 && form2.hdToYear.value == '') 	{
											var d = new Date()
											form2.hdToYear.value = d.getFullYear()				
									}			
								}			
																																			
									function fnSetSourceControl() {
									try {
										obj_caller.target.value = form2.hdFromYear.value + '|' + (parseInt(form2.hdFromMonth.value) + 1) + '|' + form2.hdFromDay.value + '|' + form2.hdToYear.value + '|' + (parseInt(form2.hdToMonth.value) + 1) + '|' + form2.hdToDay.value						
										eval(obj_caller).CloseCalendar()		
									}
									catch (everything) {   }				
								}
															
												
								function fnAdjustCalendar(vSelect)  {
									var d
									var StartingDayOfWeek
									var LastDayOfMonth
									var ElementLetter
									var Month
									var Year
									var DateLastRow
									var ThisDay = 1
									
									if (vSelect == 0) {
										ElementLetter = "F"
										DateLastRow = "FromDateLastRow"			
										Month = form2.hdFromMonth.value
										Year = form2.hdFromYear.value		
									} else {
										ElementLetter = "T"
										DateLastRow = "ToDateLastRow"
										Month = form2.hdToMonth.value
										Year = form2.hdToYear.value						
									}
									
									d = new Date(Year, Month, 1)
									StartingDayOfWeek = d.getDay()
									StartingDayOfWeek++
									LastDayOfMonth = daysInMonth(Month, Year) 
									for (var CellNum=1;CellNum<43;CellNum++)  {
										ElementName = ElementLetter + CellNum
										if (CellNum < StartingDayOfWeek || ThisDay > LastDayOfMonth)  {
											document.getElementById(ElementName).innerHTML = "&nbsp;"	
										} else {
											document.getElementById(ElementName).innerHTML = "<a href='javascript:fnSelectDay(" + vSelect + ", " + ThisDay + ")'>" + ThisDay + "</a>"		
											ThisDay++		
										}
									}
									
									if ((StartingDayOfWeek + LastDayOfMonth) < 37) {
										document.getElementById(DateLastRow).style.display='none';
									} else {
										document.getElementById(DateLastRow).style.display='inline';
									}		
								}					
								
								function daysInMonth(iMonth, iYear)  {
  									return 32 - new Date(iYear, iMonth, 32).getDate();
								}						
																					
			</script>
		</form>
	</body>
</html>
