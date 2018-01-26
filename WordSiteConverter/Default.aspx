<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="WordSiteConverter.Default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
           <asp:FileUpload ID="FileUpload1" runat="server" /><br />
         <asp:Label ID="lbl_error" runat="server" Text=""></asp:Label>
        <asp:Button runat="server" ID="btnConvert" Text="Convet" OnClick="btnConvert_Click" />

         <asp:Button runat="server" ID="btnDispose" Text="Dispose" OnClick="btnDispose_Click" />

    </div>
    </form>
</body>
</html>
