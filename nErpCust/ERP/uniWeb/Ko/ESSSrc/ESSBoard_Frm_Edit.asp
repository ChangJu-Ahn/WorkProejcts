<% Option Explicit %>
<!-- #Include file="../ESSinc/incServer.asp"  -->
<!-- #Include file="../ESSinc/lgsvrvariables.inc"  -->
<!-- #Include file="../ESSinc/Adovbs.inc"  -->
<!-- #Include file="../ESSinc/incServerAdoDb.asp" -->
<!-- #Include file="../ESSinc/incSvrFuncSims.inc" -->
<!-- #Include file="../ESSinc/incSvrVarSims.inc"  -->
<!--#include file="ESSBoard_functions.asp"-->
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/ccm.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incCookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incCommFunc.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incEvent.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/adoQuery.vbs"></SCRIPT>
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

<LINK REL="stylesheet" TYPE="Text/css" href="../ESSinc/common.css">
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

<form name="frmInsert" Method="post" action="ESSBoard_Edit.asp">
	<input type="hidden" name="seq" value="<%=seq%>">
	<input type="hidden" name="page" value="<%=page%>">
	
<table width=732 border="0" cellspacing="0" cellpadding="0">
	<tr>
	    <td height="10"></td>
	</tr>
	<tr>
	    <TD valign="top">
		<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="DDDDDD">
		    <tr> 
		      <td width="100" class="ctrow03">제목</td>
		      <td width="587" colspan=3 class="ctrow04" style='text-align:left;'>
			  <input name="subject" class=form01 maxlength="100" size=98 value="<%=subject%>"></td>
		    </tr>
		    <tr> 
		      <td class="ctrow01">작성자</td>
		      <td class="ctrow02" style='text-align:left;'><%=name%></td>
		      <td class="ctrow01">날짜</td>
		      <td class="ctrow02" style='text-align:left;'><%=UNIDateClientFormat(inputdate)%></td>
		    </tr>
			<tr valign="top"> 
			  <td colspan="4" bgcolor="F7F7F7">
				<textarea name="content" wrap="hard" style="font-family:돋움; width:100%; height:270; border:1 solid silver;padding-top:5px;padding-bottom:5px;padding-left:5px;padding-right:5px;FONT-SIZE: 9pt;"><%=Content%></textarea>
			  </td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
	    <td height="10"></td>
	</tr>
	<tr>
		<td CLASS="ctrow06" align=right>
		<a href="ESSBoard_list.asp"><img src="../ESSimage/button_09.gif" alt="리스트보기" border="0" onMouseOver="javascript:this.src='../ESSimage/button_r_09.gif';" onMouseOut="javascript:this.src='../ESSimage/button_09.gif';"></a><span style="width:10"></span>
		<img src="../ESSimage/button_02.gif" alt="저장" border="0" onclick = "javascript:PostDate();" onMouseOver="javascript:this.src='../ESSimage/button_r_02.gif';" onMouseOut="javascript:this.src='../ESSimage/button_02.gif';"><span style="width:10"></span>
		</td>
	</tr>
</table>
</form>
</body>
</html>
