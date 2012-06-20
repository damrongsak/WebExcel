<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Default.aspx.vb" Inherits="WebExcel._Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Test Excel</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
        <asp:Button ID="Button1" runat="server" Text="Export To Excel" />
        <asp:Button ID="Button2" runat="server" Text="Get Excel Data" />
        <asp:Button ID="Button3" runat="server" Text="Create Excel File" />
        <asp:CheckBox ID="FillWithStrings" runat="server" Text="Convert to string" />
    
    </div>
    </form>
</body>
</html>
