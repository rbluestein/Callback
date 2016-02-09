<%@ Page Language="VB" AutoEventWireup="false" CodeFile="CallbackMaintain.aspx.vb" Inherits="CallbackMaintain" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title runat="server" id="PageCaption"></title>    
    <link href="css/BVI.css" rel="stylesheet" type="text/css" />
    <link href="css/DCMenu.css" rel="stylesheet" type="text/css" />    
    <link href="css/JSControls.css" rel="stylesheet" type="text/css" />    
    <script src="jscripts/DCMenu.js" type="text/javascript"></script> 
    <script src="jscripts/JSLibrary.js" type="text/javascript"></script>  
     
</head>
<body onload="OnLoadEvents()" onunload="CloseChildren()">
    <form id="form2" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods="True">
        </asp:ScriptManager>
    
        <asp:Literal ID="litHiddens" runat="server"></asp:Literal><asp:Label ID="lblCurrentRights" runat="server"></asp:Label>
        <input type="hidden" id="hdAction" name="hdAction" /><input type="hidden" id="hdSubAction" name="hdSubAction" />
<%--        <input type="button" id="xxbtnEmpIDLookupNotifier" onclick="EmpIDInserted()" style="position:absolute;top:50px;left:-50px;z-index:-1" />
	    <input type="button" id="btnCallbackMaintainNotifier" />--%>
        <asp:Button ID="btnEmpIDLookupNotifier" runat="server" Text="Button" UseSubmitBehavior="False" style="position:absolute;top:50px;left:-50px;z-index:-1" />

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
                            <tr valign="middle"><td class="headstyle2" align="center">Callback Requirement</td></tr>
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
        

        
        
        <%--  DC MENU--%>
        <script type="text/javascript">             
                //var DCMenu = new DCMenu("CallbackSetup", "52px", "14px", "99px", "CBV|CBE",  "PUB|OTHER", "javascript:fnTabChg", "0");
                var DCMenu = new DCMenu("CallbackSetup", "52px", "14px", "99px", "CBV|CBE",  "PUB|OTHER", "javascript:fnTabChg", document.getElementById ("hdSelectedTab").value);                
   
 // var animal = DCMenu.GetAnimals(80)
 //  alert(animal)

                
                
                
                function fnTabChg(vIn) {
                    var SelectedTab = DCMenu.GetMenuItems()[vIn].id
                    if (vIn != document.getElementById("hdSelectedTab").value) {
                        //document.getElementById("hdSelectedTab").value = vIn
                        document.getElementById("hdSelectedTab").value = SelectedTab
                        document.getElementById("hdAction").value = "ChangeTab"
                        document.getElementsByTagName("form")[0].submit()
                    }
                }	
        </script>  
 

   

        
        

    <%--PANEL #0: HOME  --%>      
        <asp:Panel ID="pnHome" runat="server" Height="300px" Width="501px"  
        CssClass="CallbackMaintainPanel">  
         	    <table class="PrimaryTbl" cellspacing="0" cellpadding="0" border="0" width="501px">
        	            <colgroup><col width="160px" /><col width="328px" /></colgroup>
         	    
					  <tr>
				        <td colspan="2">Enter identifying information and click Emp ID Lookup for Aces clients. Enter identifying information as completely as possible, including Emp ID, for non-Aces clients.</td>
                    </tr>
                    <tr><td colspan="2" height="10px"></td></tr>    
  			        <tr>
					    <td>State:</td><td>
                        <asp:DropDownList ID="ddState" runat="server" CssClass="txtsize">
                        </asp:DropDownList>
                        <asp:TextBox ID="txtState" runat="server" CssClass="txtsize"></asp:TextBox>
                        </td>
				    </tr>                      	    	
				    <tr>
					    <td>Client:</td><td>
                        <asp:DropDownList ID="ddClientID" runat="server" CssClass="txtsize">
                        </asp:DropDownList>
                        <asp:TextBox ID="txtClientID" runat="server" CssClass="txtsize"></asp:TextBox>                        
                        </td>
				    </tr>			    
				    <tr>
					    <td>Emp Last Name:</td><td>
                        <asp:TextBox ID="txtEmpLastName" 
                            runat="server" CssClass="txtsize" MaxLength="30"></asp:TextBox>
                        </td>
				    </tr>	
				    <tr>
					    <td>Emp First Name:</td><td>
                        <asp:TextBox ID="txtEmpFirstName" runat="server" CssClass="txtsize" 
                            MaxLength="15"></asp:TextBox>
                        </td>
				    </tr>													

				  
    				
				    <tr>
					    <td>Emp MI:</td><td>
                        <asp:TextBox ID="txtEmpMI" runat="server" CssClass="txtsize" MaxLength="1"></asp:TextBox>
                        </td>
				    </tr>	
                 <tr><td colspan="2" height="10px"></td></tr>   
                
				    <tr>
					    <td><asp:HyperLink ID="lnkLookupEmpID" runat="server" CssClass="bvianchor">Emp ID Lookup:</asp:HyperLink></td>
					    <td>
                        <asp:TextBox ID="txtEmpID" runat="server" CssClass="txtsize"></asp:TextBox>&nbsp;&nbsp;
                            <asp:HyperLink ID="lnkEmpIDClear" runat="server" CssClass="bvianchor">Clear</asp:HyperLink>
                        </td>
				    </tr>
				    <tr><td colspan="2" height="18px"></td></tr>
				    <tr>
				        <td>Num Empl Calls</td>
				        <td>
                            <asp:TextBox ID="txtNumEmployeeCalls" runat="server" CssClass="txtsize" 
                                MaxLength="2"></asp:TextBox>
                        </td>
				    </tr>

				    <tr>
				        <td>Diagnostic</td>
				        <td>
                            <asp:TextBox ID="txtDiagnostic" runat="server" CssClass="txtsize" 
                                MaxLength="900"></asp:TextBox>
                        </td>
				    </tr>				    					    			        			    		

	 	    				    				  		
			    </table> 
        </asp:Panel>  
 
 
         <%--PANEL #1: SETUP  --%>
         <asp:Panel ID="pnSetup" runat="server" Height="300px" Width="500px" CssClass="CallbackMaintainPanel">
            
            
         	    <table class="PrimaryTbl" cellspacing="0" cellpadding="0" border="0" width="500px">
          	            <colgroup><col width="160px" /><col width="328px" /></colgroup>
          	        <tr><td colspan="2" height="18px"></td></tr>
    							
				    <tr>
					    <td>Ticket #:</td><td>
                        <asp:TextBox ID="txtTicketNumber" runat="server" 
                            CssClass="txtsize"></asp:TextBox>
                        </td>
				    </tr>    							
				    <tr>
					    <td>Date:</td><td><asp:TextBox ID="txtDate" runat="server" 
                            CssClass="txtsize"></asp:TextBox>
                        </td>
				    </tr>
				    <tr>
					    <td>Overflow Agent:</td><td>
                        <asp:DropDownList ID="ddOverflowAgent" runat="server" CssClass="txtsize">
                        </asp:DropDownList>
                        <asp:TextBox ID="txtOverflowAgent" runat="server" CssClass="txtsize"></asp:TextBox>
                        </td>
				    </tr>	
				    
				    <tr>
					    <td>Enrollment Window:</td><td>
                        <asp:DropDownList ID="ddEnrollWin" runat="server" CssClass="txtsize">
                        </asp:DropDownList>
                        <asp:TextBox ID="txtEnrollWin" runat="server" CssClass="txtsize"></asp:TextBox>
                        </td>
				    </tr>	
	
				    <tr>
					    <td>Priority Tag:</td><td>
                        <asp:DropDownList ID="ddPriorityTagInd" runat="server" CssClass="txtsize">
                        </asp:DropDownList>
                        <asp:TextBox ID="txtPriorityTagInd" runat="server" CssClass="txtsize"></asp:TextBox>
                        </td>
				    </tr>		
				    <tr>
					    <td>Call Type:</td><td>
                        <asp:DropDownList ID="ddCallPurpose" runat="server" CssClass="txtsize">
                        </asp:DropDownList>
                        <asp:TextBox ID="txtCallPurpose" runat="server" CssClass="txtsize"></asp:TextBox>
                        </td>
				    </tr>	
				    <tr>
					    <td>Prefer Spanish:</td><td>
                        <asp:DropDownList ID="ddPreferSpanishInd" runat="server" CssClass="txtsize">
                        </asp:DropDownList>
                        <asp:TextBox ID="txtPreferSpanishInd" runat="server" CssClass="txtsize"></asp:TextBox>
                        </td>
				    </tr>						
			    </table> 		   			    
           </asp:Panel>  
           
                                 
            <%--PANEL #2: CONTACT INFO  --%>
            <asp:Panel ID="pnContact" runat="server" Height="300px" Width="501px" CssClass="CallbackMaintainPanel">
            
            
         	    <table class="PrimaryTbl" cellspacing="0" cellpadding="0" border="0" width="501px">
          	            <colgroup><col width="160px" /><col width="328px" /></colgroup>
	                    <tr><td height="18px" colspan="2"></td></tr>
				    <tr>
				        <td>Phone #1:</td>			    
				        <td><asp:TextBox ID="txtContactPhoneNumber1" runat="server" CssClass="txtsize"  MaxLength="14"></asp:TextBox>
                            &nbsp;&nbsp;&nbsp;Ext:&nbsp;&nbsp;<asp:TextBox ID="txtContactPhoneNumber1Ext" 
                                runat="server" CssClass="txtsize" 
                                MaxLength="10"></asp:TextBox>
                        </td>
				    </tr>
				    <tr>
				        <td>Type:</td> 
				        <td><asp:TextBox ID="txtContactType1" runat="server" CssClass="txtsize" MaxLength="30"></asp:TextBox></td>
				    </tr>
		            <tr>
                        <td>Best Time:</td>
					    <td><asp:TextBox ID="txtContactBestTime1" runat="server" CssClass="txtsize" MaxLength="30"></asp:TextBox></td>				    					    
				    </tr>				    
						
		
				    
				    <tr>				    
				        <td style="height:4px" colspan="2"></td>
				    </tr>
				    <tr>    
				        <td style="background-color:#74A2E3; height:1px" colspan="2"></td>
				    </tr>
				    <tr>
				        <td style="height:4px" colspan="2"></td>
				    </tr>		
				    
				    <tr>
				        <td>Phone #2:</td>
					    <td><asp:TextBox ID="txtContactPhoneNumber2" runat="server" CssClass="txtsize" 
                                MaxLength="14"></asp:TextBox>
                             &nbsp;&nbsp;&nbsp;Ext:&nbsp;&nbsp;<asp:TextBox ID="txtContactPhoneNumber2Ext" runat="server" CssClass="txtsize" 
                                MaxLength="10"></asp:TextBox>                               
                        </td>
                    </tr>
                    <tr>					    				        
				        <td>Type:</td>
				        <td><asp:TextBox ID="txtContactType2" runat="server" CssClass="txtsize" 
                                MaxLength="30"></asp:TextBox></td>
				    </tr>							
				    <tr>
                        <td>Best Time:</td>
					    <td><asp:TextBox ID="txtContactBestTime2" runat="server" CssClass="txtsize" 
                                MaxLength="30"></asp:TextBox></td>				    					  

  
				    </tr>
				    
				    <tr>				    
				        <td style="height:4px" colspan="2"></td>
				    </tr>
				    <tr>    
				        <td style="background-color:#74A2E3; height:1px" colspan="2"></td>
				    </tr>
				    <tr>
				        <td style="height:4px" colspan="2"></td>
				    </tr>	
				    
				    <tr>
				        <td>Phone #3:</td>
					    <td><asp:TextBox ID="txtContactPhoneNumber3" runat="server" CssClass="txtsize" 
                                MaxLength="14"></asp:TextBox>
                             &nbsp;&nbsp;&nbsp;Ext:&nbsp;&nbsp;<asp:TextBox ID="txtContactPhoneNumber3Ext" runat="server" CssClass="txtsize" 
                                MaxLength="10"></asp:TextBox>                                       
                     </td>
                    </tr>
                    <tr>					    				        
				        <td>Type:</td>
				        <td><asp:TextBox ID="txtContactType3" runat="server" CssClass="txtsize" 
                                MaxLength="30"></asp:TextBox></td>
				    </tr>							
				    <tr>
                        <td>Best Time:</td>
					    <td><asp:TextBox ID="txtContactBestTime3" runat="server" CssClass="txtsize" 
                                MaxLength="30"></asp:TextBox></td>				    					  

  
				    </tr>		    			 	    				    			

	  		
			    </table> 
           </asp:Panel>    
                  
           
            <%--PANEL #3: AUTHORIZED PERSON  --%>
            <asp:Panel ID="pnAuthorize" runat="server" Height="300px" Width="501px" CssClass="CallbackMaintainPanel">  
         	    <table class="PrimaryTbl" cellspacing="0" cellpadding="0" border="0" width="501px">
        	            <colgroup><col width="160px" /><col width="328px" /></colgroup>
					  <tr>
				        <td colspan="2">Name and relationship of the person authorized to perform the enrollment if other than the employee.</td>
                    </tr>
                    

                    <tr><td colspan="2" height="10px"></td></tr>    
                    <tr>					    				        
				        <td>Name</td>
				        <td><asp:TextBox ID="txtAuthPerson" runat="server" CssClass="txtsize" 
                                MaxLength="75"></asp:TextBox></td>
				    </tr>							
				    <tr>
                        <td>Relationship</td>
					    <td><asp:TextBox ID="txtAuthRelationship" runat="server" CssClass="txtsize" 
                                MaxLength="80"></asp:TextBox></td>				    					  

  
				    </tr>    			    			 	    				    		

		  		
			    </table> 
           </asp:Panel>  
           
                      
            <%--PANEL #4: NOTES  --%>
            <asp:Panel ID="pnNotes" runat="server" Height="300px" Width="501px"  CssClass="CallbackMaintainPanel">  
         	    <table class="PrimaryTbl" cellspacing="0" cellpadding="0" border="0" width="501px">
                    <tr>
				        <td>
                            <asp:TextBox ID="txtNotes" runat="server" MaxLength="1000" Rows="14" 
                                TextMode="MultiLine"></asp:TextBox>
                        </td>
                    </tr>    			    			 	    				    				

  		
			    </table> 
           </asp:Panel>                       
                               
                
               <%--PANEL #5: ESCALATE  --%>   
            <asp:Panel ID="pnEscalate" runat="server" Height="300px" Width="501px"  
        CssClass="CallbackMaintainPanel">  
         	    <table class="PrimaryTbl" cellspacing="0" cellpadding="0" border="0" width="501px">
        	            <colgroup><col width="160px" /><col width="328px" /></colgroup>
	
	
				    <tr>
				        <td>Email escalation notice to designated enroller, copy to supervisors.</td>
                    </tr>
                   <tr><td height="10px"></td></tr>    
							
				    <tr>
                        <td>Enroller&nbsp;&nbsp;&nbsp;&nbsp;<asp:DropDownList ID="ddDesignatedEnroller" runat="server" CssClass="txtsize">
                            </asp:DropDownList>
                        </td>				    					    
				    </tr>
				    <tr><td>&nbsp;</td></tr>				    
                    <tr><td>Additional comments</td></tr>     

				    <tr>
				    <td>
				        <asp:TextBox ID="txtEmailComments" runat="server" CssClass="unselected" 
                            MaxLength="500" Rows="5" TextMode="MultiLine" Width="485px"></asp:TextBox>
				    </td>
				    </tr>
				     <tr>					    				        
				        <td>&nbsp;</td>                  
				    </tr> 
				    <tr>					    				        
				        <td>
                            <asp:HyperLink ID="lnkSendEmail" runat="server" CssClass="bvianchor">Send Email</asp:HyperLink>
                        </td>                  
				    </tr> 
				    <tr>					    				        
				        <td id="EmailResults"></td>                  
				    </tr> 				    			    			 	    							

		    				  		
			    </table>
			        <%--<a style="position:absolute;left:24px;top:300px" class="pairedbuttons2">Send</a>--%>
           </asp:Panel>  
           
			             			    			   
			        <%--LITSAVECANCELRETURN  --%>     
                <asp:Literal ID="litSaveCancelReturn" runat="server" EnableViewState="False"></asp:Literal>	
                                		                         

                <%--LITMESSAGE  --%>   
                <asp:Literal ID="litMessage" runat="server" EnableViewState="False"></asp:Literal>	


  
          <script type="text/javascript">
   			        var EmpIDLookupPage   
   			        var Init = 0

   			                
			        function OnLoadEvents() {   
			               // Notify parent then close
			               if (document.getElementById("hdNotifyParentThenClose").value == 1)  {
			                    window.opener.document.getElementById ("btnCallbackMaintainNotifier").click()
			                    window.close();
			                    return;
			               }			        
			        
			                // Refresh the checkout
			               if (IsNumeric(document.getElementById ("hdCallbackID").value)) {
			                        RefreshCheckout()
			                        setInterval("RefreshCheckout()", 20000)
			               }
			          
			                // Show VIP alert			        
			        	    if (document.getElementById("hdVIPInd").value == "1") {
			                    alert("Enrollee is VIP employee.")
			                }
			                
			                
			                               
			                // Set focus                			                
                            switch (document.getElementById ("hdSelectedTab").value) {
                                case "Home":
			                            if (document.getElementById("txtEmpLastName").getAttribute("readonly") == false) {
			                                document.getElementById("txtEmpLastName").focus()
			                            }                                
                                        break;                                      
                                case "Setup":
 			                            break;                                                                
                                case "Contact":
                                        if (Init == 0) {
                                            Init = 1
                                        } else {       			                    
			                                if (document.getElementById("txtContactPhoneNumber1").getAttribute("readonly") == false) {
			                                    document.getElementById("txtContactPhoneNumber1").focus()
			                                }
			                            }			                            		                            
			                            break;                                                                                              
                                case "Authorize":
 			                            if (document.getElementById("txtAuthPerson").getAttribute("readonly") == false) {
			                                document.getElementById("txtAuthPerson").focus()
			                            }
			                            break;                                                               
                                case "Notes":
 			                            if (document.getElementById("txtNotes").getAttribute("readonly") == false) {
			                                document.getElementById("txtNotes").focus()
			                            }
			                            break;                                                                                              
                                case "Escalate":
			                            if (document.getElementById("txtEmailComments").getAttribute("readonly") == false) {
			                                document.getElementById("txtEmailComments").focus()
			                            }
			                            break;	                                
                            }
                            
                           
//			                // Set focus
//			                switch(document.getElementById("hdSelectedTab").value) {
//			                        case "0":			                
//			                            if (document.getElementById("txtEmpLastName").getAttribute("readonly") == false) {
//			                                document.getElementById("txtEmpLastName").focus()
//			                            }
//			                            break;		                  		                    		               
//			                        case "2":			                			                    
//			                            if (document.getElementById("txtContactPhoneNumber1").getAttribute("readonly") == false) {
//			                                document.getElementById("txtContactPhoneNumber1").focus()
//			                            }
//			                            break;			                    
//			                        case "3":			                
//			                            if (document.getElementById("txtAuthPerson").getAttribute("readonly") == false) {
//			                                document.getElementById("txtAuthPerson").focus()
//			                            }
//			                            break;			                    			                  
//			                        case "4":			                
//			                            if (document.getElementById("txtNotes").getAttribute("readonly") == false) {
//			                                document.getElementById("txtNotes").focus()
//			                            }
//			                            break;
//			                        case "5":			                
//			                            if (document.getElementById("txtEmailComments").getAttribute("readonly") == false) {
//			                                document.getElementById("txtEmailComments").focus()
//			                            }
//			                            break;					                                        
//                            }
			       } 
			       
			       
   			        function RefreshCheckout() {
   			                var CallbackID = document.getElementById ("hdCallbackID").value
                            var LoggedInUserID = document.getElementById("hdLoggedInUserID").value
                            PageMethods.RefreshCheckout(LoggedInUserID, CallbackID, RefreshCheckout_Succeeded, RefreshCheckout_Failed)  			        
   			        }
  
                    function RefreshCheckout_Failed(result, userContext, methodName) {
                            //alert("Refresh checkout failed")
                           alert("CallbackMaintain: " + result._message)                            
                    }
                    
                    function RefreshCheckout_Succeeded(result, userContext, methodName) {
                            // good
                    }  			       
			       
			        function Cancel() {
                        var OKToCancel = confirm("This action deletes changes made under all of the tabs\nand restores the values existing at the time of the last save.\nDo you wish to proceed with this action?")
                        if (OKToCancel == true) {
			                document.getElementById ("hdAction").value = "CancelChanges"
			                document.getElementsByTagName("form")[0].submit()
			            }                          	
			        }
			       
			       
			        //function ClientIDChanged() {
			        //        document.getElementById ("hdClientID").value = document.getElementById ("ddClientID").value
					//        if (EmpIDLookupPage) {
	                //            EmpIDLookupPage.close()
	                //        }			             
              //}

			        function ClientIDChanged() {
			            try {
			                document.getElementById("hdClientID").value = document.getElementById("ddClientID").value
			            } catch (everything) { }
			        }
			        
			        
					function CloseChildren()  {
					        if (EmpIDLookupPage) {
	                            EmpIDLookupPage.close()
	                        }	
					}	
					
					
//		                function fnCloseChild()  {
//				            try {
//						            DatePickerYMD_EA.popupclose()
//				            }						
//				            catch (everything) {   }
//			            }

						
			        function EmpIDClear() {
			                document.getElementById("hdAction").value = "Other"
			                document.getElementById("hdSubAction").value = "EmpIDClear"
			                document.getElementsByTagName("form")[0].submit()
			        } 
			        			       		      
                    			        
			        function EmpIDInserted() {
			                if (document .getElementById ("txtEmpID").value.length > 0) {			        
			                    document.getElementById("hdAction").value = "Other"
			                    document .getElementById ("hdSubAction").value = "EmpIDInserted"			          
  
			                    //document.getElementById("hdSelectedTab").value = "0"
			                    document.getElementsByTagName("form")[0].submit()
			                }
			        }	
			        
			        
			        function HandleEnrollWinChange() {
			                document.getElementById("hdEnrollWinCode").value = document.getElementById("ddEnrollWin").value		        
			        }
			        
			        
			        function LookupEmpID() {		          
			                 // if (document.getElementById("hdAcesClients").value.toLowerCase().indexOf(document.getElementById("ddClientID").value.toLowerCase()) == -1) {	
			                  if (  (document.getElementById("hdAcesClients").value.toLowerCase().indexOf(document.getElementById("ddClientID").value.toLowerCase()) == -1) || (document.getElementById("ddClientID").value =="")       ) {				                  	           
			                        alert("Use lookup for Aces clients only.")   			          
			                  } else {
			                        if (document.getElementById("txtEmpLastName").value.length == "0" || document.getElementById("ddClientID").value == "0") {
			                                alert("ClientID and at least part of last name required.")
			                        } else {
			                            var Querystring = "EmpIDLookup.aspx?ClientID=" +  document.getElementById("ddClientID").value + "&EmpLastName=" + trim(document.getElementById("txtEmpLastName").value) + "&EmpFirstName=" + trim(document.getElementById("txtEmpFirstName").value) + "&EmpMI=" + trim(document.getElementById("txtEmpMI").value + "&State=" + document.getElementById("ddState").value)
			                            var Features = "location=no, toolbar=no, menubar=no; status=no, scrollbars=yes, resizable=no, left=10, top=10, screenX= (screen.width - 416) / 2, screenY=10, height=500, width=416"
			                            EmpIDLookupPage = window.open (Querystring, "EmpIDLookup", Features) 	
                                    }                       
                              }                         
			        }			        				
								        		           
			        function Save() {
                           var Proceed = 0
                           var AcesInd = 0
                           var EmpIDInd = 0	
                           
                           if (document.getElementById("hdAcesClients").value.toLowerCase().indexOf(document.getElementById("hdClientID").value.toLowerCase()) > -1) {
                                AcesInd = 1
                            }
                            
                            if (document.getElementById("hdEmpIDInd").value == "1") {
                                EmpIDInd = 1
                            } 
                        
                            if (AcesInd == 0) {
                                Proceed = 1
                            } 
                            else if (AcesInd == 1 && EmpIDInd == 1) {
                                Proceed = 1
                            } 
                            else if (AcesInd == 1 && EmpIDInd == 0) {
                                        var text = "You are trying to save an Aces record without an Emp ID. \nThis will cause the record to be saved as a trouble call. Click \nOK if you want to proceed with the save. Click Cancel to \nreturn to edit the record."             
                                        var response = confirm(text)                                         
                                        if (response == true) {
                                            Proceed = 1                                                     
                                        }                           
                            } 
                        
                            if (Proceed == 1) {              
                                    if (document.getElementById("hdExistingRecordInd").value == "1") {
                                        vIn = "SaveExisting"
                                    } else {
                                        vIn = "SaveNew"
                                    }                           
                                     document.getElementById ("hdAction").value = vIn                      
                                     document.getElementsByTagName("form")[0].submit()
                            }                                             
                    }     
   
    
			        function Close() {
			           window.opener.document.getElementById ("btnCallbackAttemptNotifier").click()
			           window.close()
			        }       
			        
//			        function OldNext() {
//			            var TabNum
//			            TabNum = parseInt(document.getElementById("hdSelectedTab").value)
//			            if (TabNum < 5) {
//			                TabNum = TabNum + 1
//			            }			                 
//			            fnTabChg(TabNum) 
//			        }
			        
			        function Next() {			       			     	        
			            var SelectedTab = document.getElementById("hdSelectedTab").value
			            var TabNum = DCMenu.GetTabNum(SelectedTab) + 1
			            if (TabNum < DCMenu.MenuItems.length) {
			                fnTabChg(TabNum)
			            }			            
			        }
			        
//			        function OldPrevious() {
//			            var TabNum
//			            TabNum = parseInt(document.getElementById("hdSelectedTab").value)
//			            if (TabNum > 0) {
//			                TabNum = TabNum - 1
//			            }
//			            fnTabChg(TabNum)
//			        }

//			        function OldPrevious() {
//			            var TabNum
//			            TabNum = parseInt(document.getElementById("hdSelectedTab").value)
//			            if (TabNum > 0) {
//			                TabNum = TabNum - 1
//			            }
//			            fnTabChg(TabNum)
//			        }
			        
			      
			        function Previous() {
			            var SelectedTab = document.getElementById("hdSelectedTab").value			        
			            var TabNum = DCMenu.GetTabNum(SelectedTab) - 1
			            if (TabNum > -1) {
			                fnTabChg(TabNum)	
			            }
			        }  
			       
                  
                    
                    function fnHandleClientChange() {
                        var optn
                        var dd = document.getElementById("ddEnrollWin") 
                        dd.innerHTML = ""
                                                             
                        var ClientID = document.getElementById("ddClientID").value                       
                            if (ClientID == "0") {
                                optn = document.createElement("OPTION")
                                optn.text = "Client Required"
                                optn.value = "0"
                                dd.options.add(optn)                  
                            } else {
                                optn = document.createElement("OPTION")
                                optn.text = ""
                                optn.value = "0"
                                dd.options.add(optn)
                                PageMethods.SetEnrollWinCodes(ClientID, SetEnrollWinCodes_Succeeded, OnFailed)                        
                            }                    

                    }
                    
                    function SetEnrollWinCodes_Succeeded(result, userContext, methodName) {
                        var dd = document.getElementById("ddEnrollWin")      
                        var Box1 = result.split("|")      
                        for (var i = 0 ; i < Box1.length ; i++ ) {
                           var Box2 = Box1[i].split("~")                                                           
                           var optn = document.createElement("OPTION")
                           optn.text = Box2[1]
                           optn.value = Box2[0]
                           dd.options.add(optn)                                                   
                        }                                                                                        
                    }	        
                    
                    
                    function EscalateCallback() {                                                        
                        if (document.getElementById("hdOverflowAgent").value != "0" && document.getElementById("ddDesignatedEnroller").value != "0") {
                            PageMethods.EscalateCallback(document.getElementById("hdTicketNumber").value, document.getElementById("hdOverflowAgent").value, document.getElementById("ddDesignatedEnroller").value, document.getElementById("hdPriorityTag").value, trim(document.getElementById("txtEmailComments").value), EscalateCallback_Succeeded, OnFailed) 
                        } else {
                            alert("Overflow agent and designated enroller must be provided.")
                        }
                    }
                    
                    function EscalateCallback_Succeeded(result, userContext, methodName) {
                        var dd = document.getElementById("ddDesignatedEnroller")
                        var Box = dd[dd.selectedIndex].text.split(",")                        
                        DesignatedEnroller = trim(Box[1]) + " " + trim(Box[0])          
                        alert("Email sent to " + DesignatedEnroller + " " + result + " Eastern Standard Time.")
                    }
                    
                    
                                    
                    function OnFailed(result, userContext, methodName) {
                            alert("Failed " + methodName)
                    }
                    
                    // NOTES
                    function NotesMax(txtbox) {
                        if (txtbox.value.length > 1000) {
                            txtbox.value = txtbox.value.substring(0, 1000)			                    
                        }                     
                    }
                    

                    

                    
                    
                    // PHONE NUMBER FUNCTIONS
                    function PhoneBoxChange(txtbox) {
                        //var NewText
                                                
                        if (txtbox.value.length > 10) {
                            //txtbox.value = txtbox.value.substring(0, txtbox.value.length-1)	
                            txtbox.value = txtbox.value.substring(0, 10)			                    
                        }  
                        NumericOnly(txtbox)  
                    }


//                    function PhoneBoxFocus(txtbox) {
//                        NumericOnly(txtbox) 
//                        var NewText
//                        NewText = NumericOnly(txtbox.value)
//                        if (NewText != txtbox.value) {
//                            txtbox.value = NewText
//                        }
//                    } 

                    function PhoneBoxBlur(txtbox) {
                        var NewText = ""
                        var Working
                        
                        NumericOnly(txtbox)
                        
   		                if (txtbox.value.length > 10) {
		                    txtbox.value = txtbox.value.substring(0, 10)			                    
		                }                         
                        
                       if (txtbox.value.length == 10) {
                            Working = txtbox.value
                            NewText = "(" + Working.substring(0, 3) + ") " + Working.substring(3, 6)+ "-" + Working.substring(6, 10)
                            txtbox.value = NewText
                        }
                    }
                    
                         
        </script>
       
				    		                
        </form>
    </body>
</html>

