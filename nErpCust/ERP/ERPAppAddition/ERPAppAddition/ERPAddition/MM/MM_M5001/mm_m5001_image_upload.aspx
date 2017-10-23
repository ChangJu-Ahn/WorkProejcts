<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="mm_m5001_image_upload.aspx.cs" Inherits="ERPAppAddition.ERPAddition.MM.MM_M5001.mm_m5001_image_upload" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
        <asp:FileUpload ID="fuControl" runat="server" />
&nbsp;<asp:Button ID="btn_img_upload" runat="server" Text="Button" 
            onclick="btn_img_upload_Click" />
    
    </div>
    </form>
</body>
</html>
