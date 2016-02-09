// ___ DATAGRID METHODS					
function GetDateRange() {
    vIn = "DateRangePickerYMD"
    eval(vIn).popup()
}

function Sort(vField) {
    document.getElementById("hdAction").value = "Sort"
    document.getElementById("hdSortField").value = vField
    document.getElementsByTagName("form")[0].submit()
}

function ToggleShowFilter() {
    document.getElementById("hdFilterShowHideToggle").value = 1
    document.getElementById("hdAction").value = "ApplyFilter"
    document.getElementsByTagName("form")[0].submit()
}

function ApplyFilter() {
    document.getElementById("hdAction").value = "ApplyFilter"
    document.getElementsByTagName("form")[0].submit()
}

function SubmitOnEnterKey(e) {
    var keypressevent = e ? e : window.event
    if (keypressevent.keyCode == 13) {
        form1.hdAction.value = "ApplyFilter"
        document.getElementsByTagName("form")[0].submit()
    }
}

function ShowHideSubTable(vRequestOpenClose, vActiveSubTableRecID) {
    document.getElementById("hdAction").value = vRequestOpenClose
    document.getElementById("hdCallbackID").value = vActiveSubTableRecID
    document.getElementById("hdActiveSubTableRecID").value = vActiveSubTableRecID
    document.getElementsByTagName("form")[0].submit()
}