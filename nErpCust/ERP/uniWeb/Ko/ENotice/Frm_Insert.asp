<%	Option Explicit%>
<!-- #Include file="../inc/incServer.asp"  -->
<!-- #Include file="../inc/lgsvrvariables.inc"  -->
<!-- #Include file="../inc/Adovbs.inc"  -->
<!-- #Include file="../inc/incServerAdoDb.asp" -->
<!-- #Include file="../inc/incSvrFuncSims.inc" -->
<!-- #Include file="../inc/incSvrVarSims.inc"  -->
<!--#include file="Title.asp"-->
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/ccm.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/incCookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/incCommFunc.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/incEvent.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/adoQuery.vbs"></SCRIPT>
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
	dim table:table = "EIS_Board"
	Dim page : page = request("page")
	Dim part 
	part =  right(table, len(table)-instr(table,"_"))
	
	Dim seq : seq = Request("seq") 
	Dim from_where : from_where = request("from_where")
	Dim SearchPart :	SearchPart = Request("SearchPart")
	Dim SearchStr :	SearchStr = Request("SearchStr")	
  
	Dim to_where
	if from_where="s" then
			to_where = "Search.asp?SearchPart=" & SearchPart & "&amp;SearchStr=" & SearchStr & "&amp;"
	else
			to_where = "List.asp?"
	end if  
	
  
'로긴하지 않은 사용자 거르기 
	Dim userid : userid = gUsrId
	if userid = "" then
		Response.Redirect "List.asp"
	end if
	if gUsrNm="" and gUsrId="unierp" then
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
<LINK REL="stylesheet" TYPE="Text/css" href="../inc/EIS_Board.css">
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
<title>Taeyo Board</title>

</head>
<base language="javascript">
<body leftmargin="20" topmargin="0" marginwidth="0" marginheight="0" onload="javascript:document.frmInsert.subject.focus();">
<table bgcolor="white" cellpadding="0" cellspacing="0" height="100%" width="100%">
 <tr>
  <td  align="center" valign="top">
	<form name="frmInsert" Method="post" action="Insert.asp">
		<input type="hidden" name="page" value="<%=page%>">		
		<table cellpadding="1" bgcolor="white" width="100%">
			<tr>
				<td>
				<div id="divTitle" ><% gotoTitle "INSERT"  %>&nbsp;&nbsp;</div>		    
				</td>
		    </tr>
		    <tr>
				<td>		  
					<table cellspacing="1" bgcolor="#dddddd" width="100%" >   
						<tr align=center > 
						  <td class="ctrow03" width=15% >작성자&nbsp;&nbsp;&nbsp;&nbsp;</td>
						  <td class="ctrow02" width=90% align=left> <%=gUsrNm%></td>
						</tr>
						 <tr height=10> 
						  <td width="15%" class="ctrow03" align=center>제목&nbsp;&nbsp;&nbsp;&nbsp;</td>
						  <td width="90%" class="ctrow02"  style="background:F5F5EE"><input name="subject"  size="120" style="border:1 solid F5F5EE; background:F5F5EE" value="<%=Subject%>">
						  </td>
						</tr>	
					</table>
				</td>
			</tr>		
		</table>	
		<table cellpadding="1" cellspacing="0" bgcolor="white" width="100%">
			<tr>
				<td style="padding:1">
					<textarea name="content" wrap="hard" style="font-family:돋움; width:100%; height:300; border:1 solid silver; background-image: url('../image/EIS/enotice/line.gif')"></textarea>
				</td>
			</tr>
		</table>
	</form>
	</td>
</tr>
</table>	
</body>
</html>
