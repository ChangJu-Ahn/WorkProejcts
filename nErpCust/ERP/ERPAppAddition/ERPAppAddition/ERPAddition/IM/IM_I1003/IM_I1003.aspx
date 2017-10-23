<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="IM_I1003.aspx.cs" Inherits="ERPAppAddition.ERPAddition.IM.IM_I1003.IM_I1003" %>

<%@ Register Assembly="FarPoint.Web.Spread, Version=6.0.3505.2008, Culture=neutral, PublicKeyToken=327c3516b1b18457"
    Namespace="FarPoint.Web.Spread" TagPrefix="FarPoint" %>


<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <script src="http://code.jquery.com/jquery-latest.js"></script>
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <FarPoint:FpSpread ID="FpSpread1" runat="server" BorderColor="#A0A0A0" BorderStyle="Solid" BorderWidth="1px" Height="200" Width="400">
            <CommandBar BackColor="#F6F6F6" ButtonFaceColor="Control" ButtonHighlightColor="ControlLightLight" ButtonShadowColor="ControlDark"></CommandBar>
            <Sheets>
                <FarPoint:SheetView SheetName="Sheet1"></FarPoint:SheetView>
            </Sheets>
        </FarPoint:FpSpread>
        </div>
    </form>
</body>
</html>
