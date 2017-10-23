<%	Option Explicit%>
<!-- #Include file="../ESSinc/incServer.asp"  -->
<!-- #Include file="../ESSinc/lgsvrvariables.inc"  -->
<!-- #Include file="../ESSinc/Adovbs.inc"  -->
<!-- #Include file="../ESSinc/incServerAdoDb.asp" -->
<!-- #Include file="../ESSinc/incSvrFuncSims.inc" -->
<!-- #Include file="../ESSinc/incSvrVarSims.inc"  -->
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
<%
	dim table:table = "ESS_Board"
	Dim page : page = request("page")
	Dim part 
	part =  right(table, len(table)-instr(table,"_"))
  
'로긴하지 않은 사용자 거르기 
	Dim userid : userid = gEmpNo
	if userid = "" then
		Response.Redirect "ESSBoard_list.asp"
	end if
	if gUsrNm="" and gEmpno="unierp" then
		gUsrNm="admin"
	end if
	dim subject:subject = Request.Form("Subject")
	if Subject <> "" then 
		Subject = replace(Subject, chr(34) & chr(34), "&#34;")
	End if

	dim content:content = Request.Form("content")
	if content <> "" then 
		content = Tag2Text(content)
		content = replace(content, chr(13) &chr(10), chr(13) &chr(10) & "<br>")
		content = replace(content, "<br><br>", chr(13) &chr(10) & "<p>&nbsp;</p>")
	end if
	
%>
<html>
<head>
<LINK REL="stylesheet" TYPE="Text/css" href="../ESSinc/common.css">
<meta http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
<script language="javascript" src="js/mouseMoveOnButton.js"></script>
<script language="javascript">
var statusForm;
function PostDate()
{
   // 상태 확인 
   if (statusForm) {
		alert("서버로 자료 전송 중입니다.\r\r잠시 기다려 주세요.");
		return false;
   }
        
	var val = document.frmInsert.subject.value;
	if (CheckStr(val, " ", "")==0) 
    {
      alert("제목을 입력해 주세요");
      document.frmInsert.subject.value= "";
      document.frmInsert.subject.focus();
      return;
    }
    
	var val = document.frmInsert.content.value;
	//var strEnterCode = String.fromCharCode(13, 10);
	//CheckStr(val, strEnterCode, "");
	if (CheckStr(val, " ", "")==0) 
    {
      alert("내용을 입력해 주세요");
      document.frmInsert.content.value= "";
      document.frmInsert.content.focus();
      return;
    }

	statusForm = true;
	document.frmInsert.submit();
}


function CheckStr(strOriginal, strFind, strChange){
    var position, strOri_Length;
    position = strOriginal.indexOf(strFind);  
    
    while (position != -1){
      strOriginal = strOriginal.replace(strFind, strChange);
      position    = strOriginal.indexOf(strFind);
    }
  
    strOri_Length = strOriginal.length;
    return strOri_Length;
}
 
</script>
</head>
<base language="javascript">
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="javascript:document.frmInsert.subject.focus();">

<form name="frmInsert" Method="post" action="ESSBoard_Insert.asp">
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
		      <td class="ctrow02" style='text-align:left;'><%=gUsrNm%></td>
		      <td class="ctrow01">날짜</td>
		      <td class="ctrow02" style='text-align:left;'><%=UNIDateClientFormat(GetSvrDate)%></td>
		    </tr>
			<tr valign="top"> 
			  <td colspan="4" bgcolor="F7F7F7">
				<textarea name="content" wrap="hard" style="font-family:돋움; width:100%; height:270; border:1 solid silver;padding-top:5px;padding-bottom:5px;padding-left:5px;padding-right:5px;FONT-SIZE: 9pt;"></textarea>
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
