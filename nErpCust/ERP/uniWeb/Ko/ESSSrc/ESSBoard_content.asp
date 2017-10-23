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
<LINK REL="stylesheet" TYPE="Text/css" href="../ESSinc/common.css">
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

<table width=717 border="0" cellspacing="0" cellpadding="0">
<tr>
    <td height="10"></td>
</tr>
<tr>
    <TD valign="top">
	<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="DDDDDD">
	    <tr> 
	      <td width="100" class="ctrow03">제목</td>
	      <td width="587" class="ctrow04" style='text-align:left;'><%=subject%></td>
	    </tr>
	    <tr> 
	      <td class="ctrow03">작성자</td>
	      <td class="ctrow04" style='text-align:left;'><%=name%></td>
	    </tr>
	    <tr> 
	      <td class="ctrow03">날짜</td>
	      <td class="ctrow04" style='text-align:left;'><%=UNIDateClientFormat(inputDate)%></td>
	    </tr>
	    <tr> 
	      <td class="ctrow03">조회수</td>
	      <td class="ctrow04" style='text-align:left;'><%=readCount%></td>
	    </tr>
		<tr valign="top"> 
		  <td height="120" colspan="2" bgcolor="F7F7F7" style='padding-top:5px;padding-bottom:5px;padding-left:5px;padding-right:5px;FONT-FAMILY:"돋움";FONT-SIZE: 9pt;COLOR: #5d5d5d;'><pre><%=content%></pre></td>
		</tr>
	</table>
	</td>
</tr>
<tr>
    <td height="10"></td>
</tr>
<tr>
	<TD CLASS="ctrow06" align=right height=30>
		<a href="<%=to_where%>seq=<%=seq%>&amp;page=<%=page%>"><img src="../ESSimage/button_09.gif" alt="리스트보기" border="0" onMouseOver="javascript:this.src='../ESSimage/button_r_09.gif';" onMouseOut="javascript:this.src='../ESSimage/button_09.gif';"></a><span style="width:10"></span>
		<% if gProAuth=0 then%>	
			<a href="javascript:deleteIt('<%=seq%>');">
				<img src="../ESSimage/button_10.gif" alt="삭제" border="0" onMouseOver="javascript:this.src='../ESSimage/button_r_10.gif';" onMouseOut="javascript:this.src='../ESSimage/button_10.gif';"></a><span style="width:10"></span>
			<img src="../ESSimage/button_06.gif" alt="수정" border="0" onclick = "javascript:goEdit('<%=seq%>','<%=page%>')" onMouseOver="javascript:this.src='../ESSimage/button_r_06.gif';this.style.cursor='hand';" onMouseOut="javascript:this.src='../ESSimage/button_06.gif';"><span style="width:10"></span>
		<%end if%>
	</TD>
</tr>
</table>
</body>
</html>
