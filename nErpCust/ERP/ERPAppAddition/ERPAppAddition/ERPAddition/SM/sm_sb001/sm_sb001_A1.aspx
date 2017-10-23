<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="sm_sb001_A1.aspx.cs" Inherits="ERPAppAddition.ERPAddition.SM.sm_sb001.sm_sb001_A1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<script>
        function setIFrameHeight(obj) {
            if (obj.contentDocument) {
                //alert('요기1');
                if (obj.contentDocument.body.offsetHeight < 400)
                {
                    obj.height = obj.contentDocument.body.offsetHeight + 400;
                }
                else
                {
                    obj.height = obj.contentDocument.body.offsetHeight + 50;
                }
            } else {
                if (obj.contentWindow.document.body.scrollHeight < 400)
                {
                    obj.height = obj.contentWindow.document.body.scrollHeight + 400;
                }
                else {
                    obj.height = obj.contentWindow.document.body.scrollHeight + 50;
                }
                
            }
        }
</script>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    <iframe name=iname width="100%" scrolling=no src="sm_sb001_A01.aspx?userid=<%=userid %>"  frameborder=0 onLoad="setIFrameHeight(this)"></iframe>
        
    </div>
    </form>
</body>
</html>
