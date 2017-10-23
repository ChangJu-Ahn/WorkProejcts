<% Option Explicit %>
<!-- #Include file="../inc/incServer.asp"  -->
<!-- #Include file="../inc/lgsvrvariables.inc"  -->
<!-- #Include file="../inc/Adovbs.inc"  -->
<!-- #Include file="../inc/incServerAdoDb.asp" -->
<!-- #Include file="../inc/incSvrFuncSims.inc" -->
<!-- #Include file="../inc/incSvrVarSims.inc"  -->
<!--#include file="Functions.asp"-->
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
<html>
<head>
<TITLE><%=gLogoName%>-공지사항</TITLE>
<% 
  dim table:table = "EIS_Board"
  Dim page : page = Request("page")
  Dim seq : seq = Request("seq") 
  Dim userId : userId=Request("userId")
    
  Dim from_where : from_where = request("from_where")
  Dim SearchPart :	SearchPart = Request("SearchPart")
  Dim SearchStr :	SearchStr = Request("SearchStr")	
  
  Dim to_where
  if from_where="s" then
		to_where = "Search.asp?SearchPart=" & SearchPart & "&amp;SearchStr=" & SearchStr & "&amp;"
  else
		to_where = "List.asp?"
  end if  

  if seq = "" then Response.Redirect "List.asp"

  Dim part 
  Dim  id, subject, ipAddr, readcount, inputDate, content, content_o,name  
  part =  right(table, len(table)-instr(table,"_"))
  
  'Dim userid, secureLevel	
'  userid = gEmpNo
'  secureLevel = Request.Cookies("N")("SecureLevel")
  'if secureLevel = "" then secureLevel = 0
  
	'if userid = "" and int(secureLevel) < 2 then
'		Response.Redirect "List.asp?table=" & table
'	end if
  Call SubOpenDB(lgObjConn)  
  lgStrSQL = "SELECT seq, id,  subject, ipAddr, inputDate,  readCount, content,name  from " & table
  lgStrSQL = lgStrSQL  & " WHERE  seq  in ( " & seq & ") order by seq desc "
  If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = true Then
	if lgObjRs.BOF or lgObjRs.EOF then 
		Response.Write "<Script>"
		Response.Write " location.href='List.asp"
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

<LINK REL="stylesheet" TYPE="Text/css" href="../inc/EIS_Board.css">
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

<body leftmargin="20" topmargin="0" marginwidth="0" marginheight="0" onload="javascript:document.frmInsert.subject.focus();">
<table bgcolor="white" cellpadding="0" cellspacing="0" height="100%" width="100%">
 <tr>
  <td align="center" valign="top">

	<form name="frmInsert" Method="post" action="Edit.asp">
		<input type="hidden" name="seq" value="<%=seq%>">
		<input type="hidden" name="page" value="<%=page%>">
		<table   width="100%" border=0 cellpadding=0 cellspacing=0>
		    <tr>
				<td>
				<div id="divTitle" ><% gotoTitle "MOD"%>&nbsp;&nbsp;</div>		    
				</td>
		    </tr>
			<tr>
				<td>		  
					<table cellspacing="1" bgcolor="#dddddd" width="100%" >   
						<tr align=center > 
						  <td class="ctrow03" width=15%>작성자&nbsp;&nbsp;&nbsp;&nbsp;</td>
						  <td class="ctrow02" width=23%> <%=name%></td>
						  <td class="ctrow03" width=10% >날짜&nbsp;&nbsp;&nbsp;&nbsp;</td>
						  <td class="ctrow02" width=23%><%=UNIDateClientFormat(inputDate)%></td>
						  <td class="ctrow03" width=10%>조회수&nbsp;&nbsp;</td>
						  <td class="ctrow02" width=23%><%=readCount%></td>
						</tr>
						 <tr height=10> 
						  <td width="10%" class="ctrow03" align=center>제목&nbsp;&nbsp;&nbsp;&nbsp;</td>
						  <td width="90%" class="ctrow02" colspan=5 style="background:F5F5EE"><input name="subject"  size="120" style="border:1 solid F5F5EE; background:F5F5EE" value="<%=Subject%>">
						  </td>
						</tr>	
					</table>
				</td>
			</tr>		
			<tr>
				<td style="padding:1">
					<textarea name="content" wrap="hard" style="font-family:돋움; width:100%; height:350; border:1 solid silver; background-image: url('../image/EIS/enotice/line.gif')"><%=Content%></textarea>
				</td>
			</tr>
		</table>
	</form>
  </td>
 </tr>
</table>
</body>
