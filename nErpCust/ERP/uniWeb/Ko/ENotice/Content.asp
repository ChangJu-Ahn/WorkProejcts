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
  Dim seqs : seqs = request("seqs") 
  Dim page : page = request("page")
  Dim from_where : from_where = request("from_where")
  Dim SearchPart :	SearchPart = Request("SearchPart")
  dim SearchStr :	SearchStr = Request("SearchStr")	

  Dim to_where
  if from_where="s" then
		to_where = "Search.asp?SearchPart=" & SearchPart & "&amp;SearchStr=" & SearchStr & "&amp;"
  else
		to_where = "List.asp?"
  end if
  if seqs = "" then Response.Redirect to_where

  Dim userid	
  userid = gUsrId
  Call SubOpenDB(lgObjConn)  	
  Dim seq, id, subject, ipAddr, readcount, inputDate, content, content_o,Uid
  lgStrSQL = "SELECT seq, id,  subject, ipAddr, inputDate,  readCount, content,name  from " & table
  lgStrSQL = lgStrSQL  & " WHERE  seq  in ( " & seqs & ") order by seq desc "
  If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = true Then
	if lgObjRs.BOF or lgObjRs.EOF then 
	  Response.Write "<Script>"
	  Response.Write " alert('데이터가 존재하지 않습니다');"
	  Response.Write " location.href='List.asp'"
	  Response.Write "</Script>"
	  Response.End
	end if
  End IF  
  
%>
<LINK REL="stylesheet" TYPE="Text/css" href="../inc/EIS_Board.css">
<meta http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
<script language="javascript">
<!--
	function deleteIt(seq)
	{
			var bool = confirm("정말로 삭제하시겠습니까?");
			if (bool){
				location.href = "Delete.asp?page=<%=page%>&seq=" + seq ;
			}		
	}
	function goEdit(seq,page)
	{
		location.href = "Frm_Edit.asp?page=" + page +"&seq=" + seq;
	}
//-->
</script>

</head>

<body leftmargin="20" topmargin="0" marginwidth="0" marginheight="0">

<table bgcolor="white" cellpadding="0" cellspacing="0" height="100%" width="100%">
<tr><td  align="center" valign="top">

<%
  Dim name
	seq = lgObjRs("seq")
	id = lgObjRs("id")
	subject = lgObjRs("subject")
	if Subject <> ""  then subject = Tag2Text(Subject)	
	ipAddr = lgObjRs("ipAddr")
	inputDate = lgObjRs("inputDate")
	readCount = lgObjRs("readCount")
	content = lgObjRs("content")
	id  = lgObjRs("id")
	name = lgObjRs("name")
    if content <> "" then 
		content = Tag2Text(content)
		content_o = content
		content = replace(content, chr(13) &chr(10), chr(13) &chr(10) & "<br>")
		content = replace(content, "<br><br>", chr(13) &chr(10) & "<p>&nbsp;</p>")
	end if
%>

<table cellpadding="0" cellspacing="0" width="100%" align="center">
    <tr>
		<td>
		<div id="divTitle" ><% gotoTitle "CON"%>&nbsp;&nbsp;</div>		    
		</td>
    </tr>
    <TR>
	  <td >
		<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="DDDDDD" >
		    <tr > 
		      <td width="15%" class="ctrow03" align=center>제목&nbsp;&nbsp;&nbsp;&nbsp;</td>
		      <td width="90%" class="ctrow02" colspan=5>&nbsp;<%=subject%></td>
		    </tr>
		    <tr align=center> 
		      <td class="ctrow03" width=10%>작성자&nbsp;&nbsp;&nbsp;&nbsp;</td>
		      <td class="ctrow02" width=23%> <%=name%></td>
		      <td class="ctrow03" width=10% >날짜&nbsp;&nbsp;&nbsp;&nbsp;</td>
		      <td class="ctrow02" width=23%><%=UNIDateClientFormat(inputDate)%></td>
		      <td class="ctrow03" width=10%>조회수&nbsp;&nbsp;&nbsp;&nbsp;</td>
		      <td class="ctrow02" width=23%><%=readCount%></td>
		    </tr>
		    
		    <tr> 
		      <td height="350" colspan="6" valign="top" bgcolor="F5F5EE" style="padding:0;padding-top:10;padding-left:10;border-bottom:1 solid #F5F5EE">
				<pre><%=content%></pre></td>
		    </tr>
		  </table>
		 </td>    
    </TR>    	
</table>
</form>
<br>
</td></tr>
</table>
</body>
</html>


