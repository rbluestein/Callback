<%@ Page Language="VB" AutoEventWireup="false" CodeFile="CallbackAttempt.aspx.vb" Inherits="CallbackAttempt" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title runat="server" id="PageCaption"></title>    
    <link href="css/BVI.css" rel="stylesheet" type="text/css" />
    <link href="css/DCMenu.css" rel="stylesheet" type="text/css" />    
    <link href="css/JSControls.css" rel="stylesheet" type="text/css" />    
    <script src="jscripts/DCMenu.js" type="text/javascript"></script> 
    <script src="jscripts/JSLibrary.js" type="text/javascript"></script>  
</head>
<body onload="OnLoadEvents()">
   <form id="form2" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods="True">
        </asp:ScriptManager>
        <asp:Literal ID="litHiddens" runat="server"></asp:Literal><asp:Label ID="lblCurrentRights" runat="server"></asp:Label>
        <input type="hidden" id="hdAction" name="hdAction" /><input type="hidden" id="hdSubAction" name="hdSubAction" />
        <asp:Literal ID="litEnviro" runat="server" EnableViewState="False"></asp:Literal>
    
        <%--    PAGE HEADING--%>
        <table cellpadding="0" cellspacing="0" border="0" style="position:absolute;left:22px;top:2px;width:517px; z-index:-1">
            <tr>
                <td width="140px">
                          <table cellpadding="0" cellspacing="0" border="0" width="140px">
                                <tr valign="bottom"><td><asp:Label ID="lblTicketNumber" runat="server" CssClass="headstyle1"></asp:Label></td></tr>
                                <tr valign="bottom"><td><asp:Label ID="lblStatus" runat="server" CssClass="headstyle1"></asp:Label></td></tr>             
                          </table>
                </td>
                <td width="217px">
                        <table width="217px">
                            <tr valign="middle"><td class="headstyle2" align="center">Callback Attempt</td></tr>
                         </table>
                </td>
                <td width="160px">
                        <table width="160px">
                            <tr valign="middle">
                                <td align="right">
                                    <asp:Image ID="Image1" runat="server" ImageUrl="~/img/Telephone_93_77.jpg" 
                                        Height="55px" Width="75px" />
                                </td>
                             </tr>
                         </table>            
                </td>
            </tr>
        </table>
        

         	    <table class="PrimaryTbl" cellspacing="0" cellpadding="0" border="0"  style="position:absolute;left:14px;top:54px;width:512;border-style:solid;border-color:#1E547E;border-width:1px;margin-left:14px; padding-left:10px;background: #eeeedd;" >

        	            <colgroup><col width="152px" /><col width="336px" /></colgroup>
   
             	    <tr><td colspan="2" height="18px"></td></tr>  
             	    
             	   <tr>
					    <td>Date:</td><td>
                        <asp:TextBox ID="txtDate" runat="server" 
                            CssClass="txtsize"></asp:TextBox>
                        </td>
				    </tr>		
				        	   			    
				    <tr>
					    <td>Emp Name:</td><td>
                        <asp:TextBox ID="txtEmpName" 
                            runat="server" CssClass="txtsize" MaxLength="30"></asp:TextBox>
                        </td>
				    </tr>	
				    
			        <tr>
					    <td>Client:</td><td>
                        <asp:TextBox ID="txtClient" runat="server" CssClass="txtsize"></asp:TextBox>                        
                        </td>
				    </tr>		
				    
  			        <tr>
					    <td>State:</td><td>
                        <asp:TextBox ID="txtState" runat="server" CssClass="txtsize"></asp:TextBox>
                        </td>
				    </tr>    				    				    				    
				    
                    <tr><td colspan="2" height="10px"></td></tr>    
				    	    
				    <tr>
					    <td>Enroller:</td><td>
                        <asp:DropDownList ID="ddEnroller" runat="server" CssClass="txtsize">
                        </asp:DropDownList>
                        <asp:TextBox ID="txtEnroller" runat="server" CssClass="txtsize"></asp:TextBox>
                        </td>
				    </tr>	

                  	    	    
				    <tr>
					    <td>Action:</td><td>
                        <asp:DropDownList ID="ddAction" runat="server" CssClass="txtsize">
                        </asp:DropDownList>
                        <asp:TextBox ID="txtAction" runat="server" CssClass="txtsize"></asp:TextBox>
                        </td>
				    </tr>
				    
				    <tr>
					    <td>Left message with:</td><td>
                        <asp:TextBox ID="txtLeftMessageWith" 
                            runat="server" CssClass="txtsize" MaxLength="30"></asp:TextBox>
                        </td>
				    </tr>					    					    
	
                 <tr><td colspan="2" height="10px"></td></tr>   
            
				    					    			        			    			 	    				    				  		
			    </table> 
 
             			  		        			    
			   
			        <%--LITSAVECANCELRETURN  --%>     
                <asp:Literal ID="litSaveCancelReturn" runat="server" EnableViewState="False"></asp:Literal>	
                
    
                <%--LITMESSAGE  --%>   
                <asp:Literal ID="litMessage" runat="server" EnableViewState="False"></asp:Literal>
				
		
		        <%--FUNCTIONS  --%> 	    
			    <script type="text/javascript">
			    
			        function OnLoadEvents() {
			        
			        	// Notify parent then close
			           if (document.getElementById("hdNotifyParentThenClose").value == "1")  {
			                window.opener.document.getElementById ("btnCallbackAttemptNotifier").click()
			                window.close();
			                return;
			           }
			           	
	                    // Refresh the checkout
	                    RefreshCheckout()
	                    setInterval("RefreshCheckout()", 20000)			           
			                 			           
			        }


   			        function RefreshCheckout() {
   			                var CallbackID = document.getElementById ("hdCallbackID").value
                            var LoggedInUserID = document.getElementById("hdLoggedInUserID").value
                            PageMethods.RefreshCheckout(LoggedInUserID, CallbackID, RefreshCheckout_Succeeded, RefreshCheckout_Failed)  			        
   			        }
  
                    function RefreshCheckout_Failed(result, userContext, methodName) {
                           // alert("Refresh checkout failed")
                           alert("CallbackAttempt: " + result._message)
                    }
                    
                    function RefreshCheckout_Succeeded(result, userContext, methodName) {
                            // good
                    }  
	     
			      		        		        
			        function Save() { 
			            var savetype         		       		          
                        if (document.getElementById("hdExistingRecordInd").value == "1") {
                            savetype = "SaveExisting"
                        } else {
                            savetype = "SaveNew"
                        }                           
                         document.getElementById ("hdAction").value = savetype                      
                         document.getElementsByTagName("form")[0].submit()
                    }                                             			        
	
			        function Cancel() {
			                document.getElementById ("hdAction").value = "CancelChanges"
			                document.getElementsByTagName("form")[0].submit()
			        }
			        
			        function Close() {          		       		        
			            //document.getElementById ("hdAction").value = "Return"
			            document.getElementsByTagName("form")[0].submit()
			            window.opener.document.getElementById ("btnCallbackAttemptNotifier").click()
			            window.close()
			        }			        			        
			        
			    </script> 						    			    
        </form>
    </body>
</html>
