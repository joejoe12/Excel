<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="index.aspx.cs" Inherits="Excel.index" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
     <asp:FileUpload ID="FileUpload1" runat="server" AllowMultiple="false"     />
        <asp:Button Text="Load" runat="server" ID="btnLoad" OnClick="btnLoad_Click" />
        <asp:GridView runat="server" ID="grd1" AutoGenerateColumns="true" >       </asp:GridView>
    </div>
    </form>
</body>
</html>
