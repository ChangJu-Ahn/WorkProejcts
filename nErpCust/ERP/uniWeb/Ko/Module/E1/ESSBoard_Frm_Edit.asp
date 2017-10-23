<% Option Explicit %>
<!-- #Include file="../../inc/incServer.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc"  -->
<!-- #Include file="../../inc/Adovbs.inc"  -->
<!-- #Include file="../../inc/incServerAdoDb.asp" -->
<!-- #Include file="../../inc/incSvrFuncSims.inc" -->
<!-- #Include file="../../inc/incSvrVarSims.inc"  -->
<!--#include file="ESSBoard_functions.asp"-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/ccm.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCommFunc.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEvent.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/adoQuery.vbs"></SCRIPT>
<Script Language="VBScript">
Sub Form_Load()
    Err.Clear                                                                       '☜: Clear err status
	On Error Resume Next
    Call SetToolBar("00000")
End Sub
Sub Form_UnLoad()
	On Error Resume Next
End Sub
</Script>
<html>
<head>
<TITLE><%=gLogoName%>-공지사항</TITLE>
<% 
  dim table:table = "ESS_Board"
  Dim page : page = Request("page")
  Dim seq : seq = Request("seq") 

  if seq = "" then Response.Redirect "ESSBoard_list.asp"

  Dim part 
  Dim  id, subject, ipAddr, readcount, inputDate, content, content_o,name  
  part =  right(table, len(table)-instr(table,"_"))
  
  'Dim userid, secureLevel	
'  userid = gEmpNo
'  secureLevel = Request.Cookies("N")("SecureLevel")
  'if secureLevel = "" then secureLevel = 0
  
	'if userid = "" and int(secureLevel) < 2 then
'		Response.Redirect "ESSBoard_list.asp?table=" & table
'	end if
  Call SubOpenDB(lgObjConn)  
  lgStrSQL = "SELECT seq, id,  subject, ipAddr, inputDate,  readCount, content,name  from " & table
  lgStrSQL = lgStrSQL  & " WHERE  seq  in ( " & seq & ") order by seq desc "
  If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = true Then
	if lgObjRs.BOF or lgObjRs.EOF then 
		Response.Write "<Script>"
		Response.Write " location.href='ESSBoard_list.asp"
		Response.Write "</Script>"
		Response.End
    else
  		seq = lgObjRs("seq")
		id = lgObjRs("id")
		subject = lgObjRs("subject")
		subject = Tag2Text(subject)	
		ipAddr = lgObjRs("ipAddr")
		inputDate = lgObjRs("inputDate")
		readCount = lgObjRs("readCount")
		content = lgObjRs("content")
		content = Tag2Text(content)
		name = lgObjRs("name")
   End IF 
 End IF 
'  if userid <> id and int(secureLevel) < 2 then
'	Response.Redirect "Content.asp?" & Request.ServerVariables("QUERY_STRING")
'	Response.End
 ' end if
	
%>  

<LINK REL="stylesheet" TYPE="Text/css" href="../../inc/ESS_board.css">
<meta http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
<script language="javascript" src="js/mouseMoveOnButton.js"></script>
<script language="javascript">
function PostDate()
{
	document.frmInsert.submit();
}

function mouseOverOnButton(obj)
{
	obj.style.backgroundColor = "#b6c9d9";
	obj.style.border = "1 solid black"; 
}


function mouseOutOnButton(obj)
{
	obj.style.backgroundColor = "#dddddd";
	obj.style.border = "1 solid slategray"; 
}


function mouseOverOnButton2(obj)
{
	obj.style.backgroundColor = "#b6c9d9";
	obj.style.border = "1 solid black"; 
}


function mouseOutOnButton2(obj)
{
	obj.style.backgroundColor = "#dddddd";
	obj.style.border = "#dddddd"; 
}
</script>
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="javascript:document.frmInsert.subject.focus();">
<table bgcolor="white" cellpadding="0" cellspacing="0" height="100%" width="750"><tr><td align="center" valign="top">

<form name="frmInsert" Method="post" action="ESSBoard_Edit.asp">
<input type="hidden" name="seq" value="<%=seq%>">
<input type="hidden" name="page" value="<%=page%>">
<br>
<table height="30" bgcolor="#EFEFEF" width="650">
	<tr>
		<td align="center">
			작성자 : <%=name%> <span style="width:30"></span> <%=UNIDateClientFormat(inputdate)%><br>
		</td>
	</tr>
</table>

<table cellspacing="1" bgcolor="#99a9bc" width="650">
	<tr>
		<td width="100" align="center" style="color:black" bgcolor='#d0d6e4'>제목</td>
		<td bgcolor="white" style="padding:0">
			<input name="subject" style="width:350" style="border:1 solid white" value="<%=Subject%>">
		</td>
		<td width="100" style="padding:0">
			<button onClick="javascript:PostDate();" style="background-color:#dddddd;width:100%; height:25; border:1 solid buttonface" class="verdana" accessKey="s" onmouseover="javascript:mouseOverOnButton2(this);" onmouseout="javascript:mouseOutOnButton2(this);">
				<u>S</u>ave</button>
		</td>
	</tr>
</table>
<br>
<table cellpadding="1" cellspacing="0" bgcolor="white" width="650">
	<tr>
		<td style="padding:1">
			<textarea name="content" wrap="hard" style="font-family:돋움; width:100%; height:200; border:1 solid silver; background-image: url('images/line.gif')"><%=Content%></textarea>
		</td>
	</tr>
</table>
</form>

</td></tr></table>
</body>
