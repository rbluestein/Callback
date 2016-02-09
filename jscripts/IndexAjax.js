function fnAltLogin() {
    PageMethods.AuthenticateEnroller(document.getElementById("txtAltLogin").value, fnAltLogin_Succeeded, fnAltLogin_Failed)
}
function fnAltLogin_Failed(result, userContext, methodName) {
    alert("A connection error occurred. Please try again.")
}
function fnAltLogin_Succeeded(result, userContext, methodName) {
    var box = result.split("|")

    // 0: IsAuthenticated
    // 1: NotFound
    // 2: ConnectionError
    // 3: InsufficentRights
    // 4: Message

    if (box[1] == 1 || box[2] == 1 || box[3] == 1) {
        alert(box[4])
    } else {
        document.getElementById("hdAction").value = "AltLoginAuthenticated"
        document.getElementsByTagName("form")[0].submit()
    }
}

function CheckoutRecordIfAvailable(CallbackID, CallbackAttemptID, PageCode, MaintDelete, NewExisting, MaintOpen, AttemptOpen) {
    var LoggedInUserID = document.getElementById("hdLoggedInUserID").value
    PageMethods.CheckoutRecordIfAvailable(LoggedInUserID, CallbackID, CallbackAttemptID, PageCode, MaintDelete, NewExisting, MaintOpen, AttemptOpen, CheckoutRecord_Succeeded, CheckoutRecord_Failed)
}

function CheckoutRecord_Failed(result, userContext, methodName) {
    alert("Index: " + result._message)
}

function CheckoutRecord_Succeeded(result, userContext, methodName) {
    var CallbackFeatures = "location=no, toolbar=no, menubar=no; status=no, scrollbars=no, resizable=no, left=10, top=10, screenX= (screen.width - 416) / 2, screenY=10, height=432, width=562"
    var CallbackAttemptFeatures = "location=no, toolbar=no, menubar=no; status=no, scrollbars=no, resizable=no, left=10, top=10, screenX= (screen.width - 416) / 2, screenY=10, height=294, width=562"
    var AnotherMaintain = "Another Callback page is open. Please close it. Save your changes if any. You may need to wait 30 seconds. Use the close button, not the red X button."
    var AnotherAttempt = "Another Attempt page is open. Please close it. Save your changes if any. You may need to wait 30 seconds. Use the close button, not the red X button."
    var OKToOpenCallbackAttempt = 0
    var RecordLockOverride = 0
    var box = result.split("|")

    //   0: CallbackID
    //   1: IsAlreadyCheckedOut
    //   2: CheckoutUser
    //   3: ErrorInd
    //   4: ErrorMessage
    //   5: PageCode
    //   6: NewCheckoutSuccessfulInd      
    //   7: CallbackAttemptID
    //   8: MaintDelete
    //   9: CheckoutUserID
    //  10: NewExisting
    //  11: MaintOpen
    //  12: AttemptOpen

    // ___ Method returns error
    if (box[3] == "1") {
        alert(box[4])
        return;
    }

    // ___ Override lock to allow one maintain and one attempt page
    if (box[1] == "1") {
        if (box[8] == "maint") {
            if (CallbackMaintainPage && (box[5] == "newattempt" || box[5] == "existingattempt")) {
                box[1] = 0
            }
            else if (CallbackAttemptPage && box[5] == "requirement") {
                box[1] = 0
            }
        }
    }

    // ___ Record is currently locked
    if (box[1] == "1") {
        alert("Record is in use by " + box[2])
        return;
    }

    // ___ Register checkout
    document.getElementById("hdCheckedOutCallbackID").value = box[0]

    // ___ Navigate to page
    if (box[8] == "maint") {
        if (box[5] == "requirement") {

            //if (CallbackMaintainPage) {
            //    var OKToProceed = confirm("Another Callback page is open in this browser. By clicking OK, you will cause this page to close this page and lose all data in it. Do you wish to proceed with this action?")
            //    if (OKToProceed == true) {
            //        CallbackMaintainPage = window.open("CallbackMaintain.aspx?CallType=Existing&CallbackID=" + box[0], "Maintain", CallbackFeatures)
            //    }
            //} else {
            //    CallbackMaintainPage = window.open("CallbackMaintain.aspx?CallType=Existing&CallbackID=" + box[0], "Maintain", CallbackFeatures)
            //}

            if (box[11]=="1") {
                alert(AnotherMaintain)
            } else {
                window.open("CallbackMaintain.aspx?CallType=" + box[10] + "&CallbackID=" + box[0], "Maintain", CallbackFeatures)
                MaintOpen = "1"
            }


        } else if (box[5] == "newattempt" || box[5] == "existingattempt") {

            if (box[12]=="1") {
                //var OKToProceed = confirm("Another callback attempt page is open in this browser. By clicking OK, you will cause this page to close this page and lose all data in it. Do you wish to proceed with this action?")
                //if (OKToProceed == true) {
                //    OKToOpenCallbackAttempt = 1
                //}
                alert(AnotherAttempt)
            } else {
                OKToOpenCallbackAttempt = 1
            }

            if (OKToOpenCallbackAttempt == 1) {
                if (box[5] == "newattempt") {
                    //CallbackAttemptPage = window.open("CallbackAttempt.aspx?CallType=New&CallbackID=" + parseInt(box[0]), "CallbackAttempt", CallbackAttemptFeatures)
                    window.open("CallbackAttempt.aspx?CallType=New&CallbackID=" + parseInt(box[0]), "CallbackAttempt", CallbackAttemptFeatures)
                    box[12] ="1"
                } else if (box[5] == "existingattempt") {
                    CallbackAttemptPage = window.open("CallbackAttempt.aspx?CallType=Existing&CallbackID=" + parseInt(box[0]) + "&CallbackAttemptID=" + parseInt(box[7]), "CallbackAttempt", CallbackAttemptFeatures)
                }
            }

        }


    } else {
        var OKToDelete = confirm("Are you sure you wish to delete this record?")
        if (OKToDelete == true) {
            if (box[5] == "requirement") {
                document.getElementById("hdAction").value = "DeleteCallback"
                document.getElementById("hdCallbackID").value = box[0]
            } else {
                document.getElementById("hdAction").value = "DeleteCallbackAttempt"
                document.getElementById("hdCallbackAttemptID").value = box[7]
            }
            //form1.submit()	
            document.getElementsByTagName("form")[0].submit()
        }
    }
}

function CheckInRecord() {
    PageMethods.CheckInRecord(document.getElementById("hdCheckedOutCallbackID").value, CheckInRecord_Succeeded, CheckInRecord_Failed)
}
function CheckInRecord_Failed(result, userContext, methodName) {
    alert("Index: " + result._message)
}
function CheckInRecord_Succeeded(result, userContext, methodName) {
}
