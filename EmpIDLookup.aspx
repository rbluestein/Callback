<%@ Page Language="VB" AutoEventWireup="false" CodeFile="EmpIDLookup.aspx.vb" Inherits="EmpIDLookup" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Untitled Page</title>
    <link href="css/BVI.css" rel="stylesheet" />
</head>
<body>
    <form id="form2" runat="server">
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods="True">
    </asp:ScriptManager>
    
    <div>
        <asp:Literal ID="litHiddens" runat="server" EnableViewState="False"></asp:Literal>
			<table class="PrimaryTbl" style="POSITION: absolute; TOP: 0px; LEFT: 0px" cellspacing="0"
				cellpadding="0" width="400" border="0">
				<tr>
					<td class="PrimaryTblTitle" colspan="2">EmpID Lookup</td>
				</tr>
				<tr>
					<td class="CellSeparator" colspan="2"></td>
				</tr>
				<tr>
					<td><asp:Literal ID="litDG" runat="server" EnableViewState="False"></asp:Literal></td>
				</tr>
				<tr>
					<td colspan="2"></td>
				</tr>

			</table>        
    </div>
        <script type="text/javascript">
            function EmpIDClick(EmpID) {
                var ClientID = document.getElementById ("hdClientID").value
                PageMethods.EmpIDClick(EmpID, ClientID, EmpIDClick_Succeeded, EmpIDClick_Failed)  			        
            }
            
            function EmpIDClick_Failed(result, userContext, methodName) {
                   alert("EmpIDLookup EmpIDClick_Failed: " + result._message)
            }            
          
            function EmpIDClick_Succeeded(result, userContext, methodName) {
                // 0: ErrorInd
                // 1: ErrorMessage
                // 2: EmpID
                // 3: TicketExistsInd 
                           
                var Box = result.split("|") 
                 
                // ___ Method returns error
                if (Box[0] == "1") {
                    alert(Box[1])
                    return;
                } 
                
                                  
                if (Box[3] == 1) {
                    alert("A ticket already exists for this employee.")
                } else {                          
                    window.opener.document.getElementById("txtEmpID").value = Box[2]
                    window.opener.document.getElementById("btnEmpIDLookupNotifier").click()
                    window.close()
                }                
                                                 
            }            
                            
        </script>
    
    </form>
</body>
</html>