<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="WebForm1.aspx.cs" Inherits="ExportExcelToSQL.WebForm1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>Import Excel Data</title>
</head>
<body>
    <form id="form1" runat="server">
    
        <!-- ADD A FILE UPLOAD CONTROL AND A BUTTON TO EXECUTE. -->
        <div style="font:14px Verdana">
        
            Select a file to upload: 
                <asp:FileUpload ID="FileUpload" Width="450px" runat="server" />
            <p>
                <asp:Button ID="Button1" runat="server" OnClick="ImportFromExcel" Text="ImportToDatabase" />
            </p>
            <p><asp:Label id="lblConfirm" runat="server"></asp:Label></p>
                
        </div>
        
    </form>
</body>
</html>
