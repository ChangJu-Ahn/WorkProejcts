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
  Dim seqs : seqs = request("seqs") 
  Dim page : page = request("page")
  Dim from_where : from_where = request("from_where")
  Dim SearchPart :	SearchPart = Request("SearchPart")
  dim SearchStr :	SearchStr = Request("SearchStr")	

  Dim to_where
  if from_where="s" then
		to_where = "ESSBoard_SearchResult.asp?SearchPart=" & SearchPart & "&amp;SearchStr=" & SearchStr & "&amp;"
  else
		to_where = "ESSBoard_list.asp?"
  end if
  if seqs = "" then Response.Redirect to_where

  Dim userid	
  userid = gEmpNo
  Call SubOpenDB(lgObjConn)  	
  Dim seq, id, subject, ipAddr, readcount, inputDate, content, content_o,Uid
  lgStrSQL = "SELECT seq, id,  subject, ipAddr, inputDate,  readCount, content,name  from " & table
  lgStrSQL = lgStrSQL  & " WHERE  seq  in ( " & seqs & ") order by seq desc "
  If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = true Then
	if lgObjRs.BOF or lgObjRs.EOF then 
	  Response.Write "<Script>"
	  Response.Write " alert('데이터가 존재하지 않습니다');"
	  Response.Write " location.href='ESSBoard_list.asp'"
	  Response.Write "</Script>"
	  Response.End
	end if
  End IF  
  
%>
<LINK REL="stylesheet" TYPE="Text/css" href="../../inc/ESS_board.css">
<meta http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
<script language="javascript">
<!--
	function deleteIt(seq)
	{
			var bool = confirm("정말로 삭제하시겠습니까?");
			if (bool){
				location.href = "ESSBoard_Delete.asp?page=<%=page%>&seq=" + seq ;
			}		
	}
	function goEdit(seq,page)
	{
		location.href = "ESSBoard_Frm_Edit.asp?page=" + page +"&seq=" + seq;
	}
//-->
</script>

</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<table bgcolor="white" cellpadding="0" cellspacing="0" height="100%" width="750">
<tr><td  align="center" valign="top">

<br>
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

<table cellpadding="5" cellspacing="0" width="700" align="center">

	<tr  width="650" >
		<td bgcolor="white" style="padding:2" align="right" valign="bottom" colspan="2">
			<%=UNIDateClientFormat(inputDate)%>&nbsp; <span style="width:20"></span>조회수 : <%=readCount%>
		</td>
	</tr>
	<tr width="650">
		<td bgcolor="#d0d6e4" style="padding:5;padding-left:30;border-bottom:1 solid #99a9bc;border-top:1 solid #99a9bc;font-weight:bold;color:#2e4287" colspan="2">
			&nbsp; 제목 : 	<%=subject%></td>
	</tr>
	<tr width="650">
		<td bgcolor="#eeeeee" align="right" width="100%">
		 작성자 : <%=name%> <!--(ip:<%=ipAddr%>)--> </td>
	</tr>
	<tr width="650">
		<td  bgcolor="#f5f5f5" style="padding:0;padding-left:0;border-bottom:1 solid #99a9bc" colspan="2">
			<pre><%=content%></pre>
		</td>
	</tr>
	<tr HEIGHT="26" width="650">
		<td bgcolor="white" HEIGHT="26" style="padding:5;padding-left:30;border-bottom:1 solid #99a9bc" colspan="2" >
			<p ALIGN="right">
					<a href="<%=to_where%>seq=<%=seq%>&amp;page=<%=page%>"><img src="../../../CShared/image/uniSIMS/print1.jpg" alt="리스트보기" border="0"WIDTH="26" HEIGHT="26"></a><span style="width:10"></span>
					<% if id = userid and  gProAuth=0 then %>								
					<img src="../../../CShared/image/uniSIMS/save1.jpg" alt="수정" border="0" WIDTH="26" HEIGHT="26" onclick = "javascript:goEdit('<%=seq%>','<%=page%>')"><span style="width:10"></span>
					<a href="javascript:deleteIt('<%=seq%>');">
						<img src="../../../CShared/image/uniSIMS/del1.jpg" alt="삭제" border="0" WIDTH="26" HEIGHT="26"></a><span style="width:10"></span>
					<% end if%>
			</p>
		</td>
	</tr>
</table>
</form>
<br>
</td></tr>
</table>
</body>
</html>


