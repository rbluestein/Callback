function OnLoadEvents() {
    var EmpID
    //DoBrowser()

    if (document.getElementById("hdJumpToRecInd").value == "1") {
        EmpID = document.getElementById("hdJumpToRecID").value
        window.location.hash = '# + EmpID + '
    }
}

function OnUnloadEvents() {
    if (CallbackMaintainPage) {
        CallbackMaintainPage.close()
    }
    if (CallbackAttemptPage) {
        CallbackAttemptPage.close()
    }
}

// ___User, site, buildnum table	
try {
    document.write("<table style='background:#eeeedd;PADDING-LEFT: 12px;FONT: 8pt Arial, Helvetica, sans-serif; POSITION: absolute;TOP: 14px;left: 14px' cellSpacing='0' cellPadding='0' width='125' border='0'><tr><td width='30'>User:</td><td>" + form1.hdLoggedInUserID.value + "</td></tr><tr><td>Site:</td><td WRAP=HARD>" + form1.hdDBHost.value + "</td></tr><tr><td>Build:</td><td>" + form1.hdBuildNum.value + "</td></tr></table>")
} catch (everything) { }

//// ___ AltLogin
//try {
//    var statement
//    statement = "<table style=\"background:#eeeedd;PADDING-LEFT: 12px;FONT: 8pt Arial, Helvetica, sans-serif; POSITION: absolute;TOP: 14px;left: 170px\" "
//    statement = statement + "cellSpacing=\"0\" cellPadding=\"0\" width=\"250\" border=\"0\">"
//    statement = statement + "<colgroup><col width=\"40\"/><col width=\"190\"/></colgroup>"
//    statement = statement + "<tr><td align=\"center\" colspan=\"2\">App not cooperating?</td></tr>"
//    statement = statement + "<tr><td>UserID:</td>"
//    //statement = statement + "<td><input type=\"text\" id=\"AltLogin\" name=\"Altlogin\" onkeypress=\"EnterKeySubmit(event)\" />&nbsp;&nbsp;"
//    statement = statement + "<td><input type=\"text\" id=\"txtAltLogin\" name=\"txtAltlogin\" />&nbsp;&nbsp;"
//    statement = statement + "<input type=\"button\" "
//    // statement = statement + "onclick=\"document.getElementById('hdAction').value='AltLogin'; document.getElementsByTagName('form')[0].submit()\"/>"
//    statement = statement + "onclick=\"fnAltLogin()\"/>"
//    statement = statement + "</td></tr></table>"
//    document.write(statement)
//} catch (everything) { }

// ___ DateRangePicker, var CallbackMaintainPage, var CallbackAttemptPage, Feature
try {
    // var DateRangePickerYMD = new DateRangePickerYMD(document.forms['form1'], document.getElementById("hdDateRange"), "Enter Date Range", 0);
    var CallbackMaintainPage
    var CallbackAttemptPage
    var CallbackFeatures = "location=no, toolbar=no, menubar=no; status=no, scrollbars=no, resizable=no, left=10, top=10, screenX= (screen.width - 416) / 2, screenY=10, height=432, width=562"
    var CallbackAttemptFeatures = "location=no, toolbar=no, menubar=no; status=no, scrollbars=no, resizable=no, left=10, top=10, screenX= (screen.width - 416) / 2, screenY=10, height=294, width=562"
} catch (everything) { }

function NewRecord() {
    Process("NewRecord")
}

function ExistingRequirement(CallbackID) {
    Process("ExistingRequirement", CallbackID)
}

function NewAttempt(CallbackID) {
    Process("NewAttempt", CallbackID)
}

function ExistingCallbackAttempt(CallbackAttemptID, CallbackID) {
    Process("ExistingCallbackAttempt", CallbackID, CallbackAttemptID)
}

function ReturnFromMaintain() {
    if (IsNumeric(document.getElementById("hdCheckedOutCallbackID").value)) {
        CallbackMaintainPage = null
        CheckInRecord()
    }
    ApplyFilter()
}

function ReturnFromAttempt() {
    if (IsNumeric(document.getElementById("hdCheckedOutCallbackID").value)) {
        CallbackAttemptPage = null
        CheckInRecord()
    }
    ApplyFilter()
}


// ___ Delete CallbackMaintain record
function DeleteCallback(CallbackID) {
    Process("DeleteCallback")
    // CheckoutRecord(CallbackID, 0, "requirement", "delete")
}

// ___ Delete CallbackAttempt record
function DeleteCallbackAttempt(CallbackAttemptID, CallbackID) {
    Process("DeleteCallbackAttempt", CallbackID, CallbackAttemptID)
    // CheckoutRecord(CallbackID, CallbackAttemptID, "existingattempt", "delete")
}

function Process(RequestAction, CallbackID, CallbackAttemptID) {
    var DisallowMultMaintain
    var DisallowMultAttempt
    var NotifyIndex
    var CheckOut
    var CheckIn
    var ApplyFilterr
    var MaintOpen = "0"
    var AttemptOpen = "0"

    if (CallbackMaintainPage) {
        MaintOpen = "1"
    }
    if (CallbackAttemptPage) {
        AttemptOpen = "1"
    }

    switch (RequestAction) {
        case "NewRecord":
            if (CallbackMaintainPage) {
                alert(AnotherMaintain)
            } else {
                CallbackMaintainPage = window.open("CallbackMaintain.aspx?CallType=New&CallbackID=''", "Maintain", CallbackFeatures)
                MaintOpen = "1"
            }
            break;

        case "ExistingRequirement":
            CheckoutRecordIfAvailable(CallbackID, 0, "requirement", "maint", "existing", MaintOpen, AttemptOpen)
            break;

        case "ReturnFromMaintain":
            MaintOpen = "0"
            if (IsNumeric(document.getElementById("hdCheckedOutCallbackID").value)) {
                CheckInRecord()
            }
            CallbackMaintainPage = null
            ApplyFilter()
            break;

        case "NewAttempt":
            CheckoutRecordIfAvailable(CallbackID, 0, "newattempt", "maint", "new", MaintOpen, AttemptOpen)
            break;

        case "ExistingCallbackAttempt":
            CheckoutRecordIfAvailable(CallbackID, CallbackAttemptID, "existingattempt", "maint", "existing", MaintOpen, AttemptOpen)
            break;

        case "ReturnFromAttempt":
            AttemptOpen = "0"
            if (IsNumeric(document.getElementById("hdCheckedOutCallbackID").value)) {
                CheckInRecord()
            }
            CallbackAttemptPage = null
            ApplyFilter()
            break;

        case "DeleteCallback":    // !!!!!!!!!!!!! need testing
            CheckoutRecordIfAvailable(CallbackID, 0, "requirement", "delete", "existing", MaintOpen, AttemptOpen)
            break;

        case "DeleteCallbackAttempt":   // !!!!!!!!!!!!! need testing
            CheckoutRecordIfAvailable(CallbackID, CallbackAttemptID, "existingattempt", "delete", "existing", MaintOpen, AttemptOpen)
            break;
    }
}