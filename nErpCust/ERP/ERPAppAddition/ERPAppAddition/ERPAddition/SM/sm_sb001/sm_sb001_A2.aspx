<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="sm_sb001_A2.aspx.cs" Inherits="ERPAppAddition.ERPAddition.SM.sm_sb001.sm_sb001_A2" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <script>
        function setIFrameHeight(obj) {
            if (obj.contentDocument) {
                obj.height = obj.contentDocument.body.offsetHeight + 40;
            } else {
                obj.height = obj.contentWindow.document.body.scrollHeight + 40;
            }
        }
</script>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <iframe name=iname width="100%" scrolling=no frameborder=0 src="sm_sb001_A02.aspx" onLoad="setIFrameHeight(this)"></iframe>
    </div>
    </form>
</body>
</html>
