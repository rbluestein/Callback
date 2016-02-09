<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Index.aspx.vb" Inherits="Index" %>

<!DOCTYPE html>

<html>
<head runat="server">
    <title runat="server" id="PageCaption"></title>
    <link href="css/BVI.css" rel="stylesheet" />   
    <link href="css/JSControls.css" rel="stylesheet" />
    <script src="jscripts/JSLibrary.js"></script>
    <script src="jscripts/DateRangePickerYMD.js" type="text/javascript"></script>
    <script src="jscripts/IndexMain1.js"></script>
    <script src="jscripts/IndexMain2.js"></script>
    <script src="jscripts/IndexDG.js"></script>  
    <script src="jscripts/IndexAjax.js"></script>
</head>
<body onload="OnLoadEvents()" onunload="OnUnloadEvents()">
    <form id="form1" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods="True">
        </asp:ScriptManager>          
        <div>
			<input type="hidden" id="hdAction" name="hdAction"/> <input type="hidden" id="hdSortField" name="hdSortField" />
			<input type="hidden" id="hdCallbackID" name="hdCallbackID" /><input type="hidden" id="hdCheckedOutCallbackID" /><input type="hidden" id="hdCallbackAttemptID" name="hdCallbackAttemptID" />
			<input type="hidden" id="hdFilterShowHideToggle" value="0" name="hdFilterShowHideToggle" /><input type="hidden" id="hdCallbackMaintainSave" />
			<input type='hidden' id='hdActiveSubTableRecID' name='hdActiveSubTableRecID' />
            <input type="button" id="btnCallbackMaintainNotifier" onclick="ReturnFromMaintain()" style="position:absolute;top:50px;left:-50px;z-index:-1" />
            <input type="button" id="btnCallbackAttemptNotifier"  onclick="ReturnFromAttempt()"  style="position:absolute;top:90px;left:-50px;z-index:-1" />			
				
            <asp:Literal ID="litAltLogin" runat="server"></asp:Literal>

			<%--<table class="PrimaryTbl" style="POSITION: absolute; TOP: 14px; LEFT: 14px" cellspacing="0"--%>
			<table class="PrimaryTbl" style="width:950px; POSITION: absolute; TOP: 60px; LEFT: 14px" cellspacing="0"
				cellpadding="0" border="0">	
				<colgroup style="width:100px""></colgroup>
				<colgroup style="width:300px"></colgroup>				
				<tr>
					<td class="PrimaryTblTitle" colspan="2"><asp:Literal ID="litHeading" runat="server" EnableViewState="False"></asp:Literal>
                    </td>
				</tr>
             
				<tr>
					<td class="cellseparator" colspan="2">&nbsp;</td>
				</tr>
				<tr>
					<td colspan="2"><asp:literal id="litDG" runat="server" enableviewstate="False"></asp:literal>
                    </td>
				</tr>
							
			</table>
			<asp:label id="lblCurrentRights" runat="server"></asp:label><asp:literal id="litEnviro" runat="server" enableviewstate="False"></asp:literal>
			<asp:literal id="litMessage" runat="server" enableviewstate="False"></asp:literal><asp:literal id="litFilterHiddens" runat="server" enableviewstate="False"></asp:literal>
        </div>	
		<script  type="text/javascript">		
			    try {
			        document.write("<table style='background:#eeeedd;PADDING-LEFT: 12px;FONT: 8pt Arial, Helvetica, sans-serif; POSITION: absolute;TOP: 14px;left: 14px' cellSpacing='0' cellPadding='0' width='125' border='0'><tr><td width='30'>User:</td><td>" + form1.hdLoggedInUserID.value + "</td></tr><tr><td>Site:</td><td WRAP=HARD>" + form1.hdDBHost.value + "</td></tr><tr><td>Build:</td><td>" + form1.hdBuildNum.value + "</td></tr></table>")
			    } catch (everything) { }
			   // document.write("<table style='background:#eeeedd;PADDING-LEFT: 12px;FONT: 8pt Arial, Helvetica, sans-serif; POSITION: absolute;TOP: 14px;left: 100px' cellSpacing='0' cellPadding='0' width='125' border='0'><colgroup><col width='50'/><col width='35'/></colgroup><tr><td colspan='2'>App not cooperating?</td><td>UserID:</td></tr><tr><td>Site:</td><td WRAP=HARD>" + form1.hdDBHost.value + "</td></tr><tr><td>Build:</td><td>" + form1.hdBuildNum.value + "</td></tr></table>")
        
            

			    var DateRangePickerYMD = new DateRangePickerYMD(document.forms['form1'], document.getElementById("hdDateRange"), "Enter Date Range", 0);

            </script>       	        	
    </form>
</body>
</html>

