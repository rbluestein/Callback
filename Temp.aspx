<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Temp.aspx.vb" Inherits="Temp" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <link href="css/BVI.css" rel="stylesheet" />
    <link href="css/DCMenu.css" rel="stylesheet" />
    <link href="css/JSControls.css" rel="stylesheet" />
    <script src="jscripts/DCMenu.js"></script>
    <script src="jscripts/JSLibrary.js"></script>

</head>
<body>
    <form id="form1" runat="server">
    <div>
hello
<a href="javascript:Save()"  style="position: absolute; left:24px; top: 392px;"  class="pairedbuttons4">Save&nbsp;&nbsp;</a>
<a href="javascript:Cancel()"  style="position: absolute; left:136px; top: 392px;"  class="pairedbuttons4">Cancel&nbsp;&nbsp;</a>
<a href="javascript:Close()"  style="position: absolute; left:248px; top: 392px;"  class="pairedbuttons4">Close&nbsp;&nbsp;</a>
<a href="javascript:Previous()" style="position: absolute;left: 440px;top:388px;" class="previousbutton"></a>
<a href="javascript:Next()" style="position: absolute;left: 480px;top:388px;" class="nextbutton"></a>	
greetings    
    </div>
        <asp:TextBox ID="TextBox1" runat="server"></asp:TextBox>
        <asp:TextBox ID="TextBox2" runat="server"></asp:TextBox>
    </form>
</body>
</html>
