 <LINK REL="stylesheet" TYPE="Text/css" href="../inc/common.css">
<link href="../inc/top.css" rel="stylesheet" type="text/css">

 <%sub gotoTitle(strName)
	Response.Write "<TABLE BORDER =0  WIDTH=100% CELLPADDING=0 CELLSPACING=1>" & vbcrlf
	Response.Write "    <TR>" & vbcrlf
	Response.Write "        <td width='61' height=36><img src='../../CShared/EISImage/Main/title_icon.gif'></td>        " & vbcrlf
	Response.Write "        <td width='150'  class=title>공지사항</td>                "& vbcrlf
	Response.Write "	    <TD ALIGN=RIGHT VALIGN=BOTTOM>&nbsp;"& vbcrlf
	'-----------button----s------------------------------------------------------------

	Response.Write "					<table BORDER = 0 CELLPADDING=0 CELLSPACING=1>" & vbcrlf
	Response.Write "                     <tr> " & vbcrlf
	Response.Write "                       <td></td>" & vbcrlf
		
	if gProAuth=0 then				
		select case strName
		case "LIST"    '목록 
			Response.Write "<td width=110>&nbsp;</td> " & vbcrlf 
			Response.Write "<td align='right' > " & vbcrlf 
			Response.Write " <span style='width:2'></span>			" & vbcrlf 
			Response.Write " <img align='absmiddle' style='cursor:hand' id='imgNew' src='../image/EIS/enotice/button_12.gif' " 
			 if userid <> "" then 
				 Response.Write " onClick=javascript:location.href='Frm_Insert.asp?page=" & page & "'  " 
			 else 
				 Response.Write " onClick='javascript:alert('로긴을 하셔야 합니다');' "
			 end if
			 Response.Write " alt='새글작성' onMouseOver=javascript:this.src='../image/EIS/enotice/button_r_12.gif'  " 
			 Response.write " onMouseOut=javascript:this.src='../image/EIS/enotice/button_12.gif'>" & vbcrlf	
			 Response.Write " </td>			" & vbcrlf 

		case "CON"	'내용보기 
			
		Response.Write " <td width=110>&nbsp;</td> " & vbcrlf 
		Response.Write "<td align='right' > " & vbcrlf 
		Response.Write " <span style='width:2'></span>			" & vbcrlf 
		Response.Write " <a href=" & to_where & " seq=" & seq & " &amp;page=" & page & ">"
		Response.Write " <img src='../image/EIS/enotice/bu_list.gif' alt='리스트' border='0' " 
		Response.Write "  onMouseOver=javascript:this.src='../image/EIS/enotice/bu_r_list.gif'  "
		Response.write "  onMouseOut=javascript:this.src='../image/EIS/enotice/bu_list.gif' ></a>"		
		if id = userid    then
			Response.Write " <A href='#'><img src='../image/EIS/enotice/bu_modi.gif'  alt='수정' border='0' "
			Response.Write " onClick='javascript:goEdit(" & seq & "," & page & ")' "
			Response.Write "  onMouseOver=javascript:this.src='../image/EIS/enotice/bu_r_modi.gif'  "
			Response.write "  onMouseOut=javascript:this.src='../image/EIS/enotice/bu_modi.gif' > "
			Response.Write " </a><span style='width:5'></span>" & vbcrlf 
			Response.Write " <a href='javascript:deleteIt(" & seq & ")'>"
			Response.Write " <img src='../image/EIS/enotice/bu_del.gif'  alt='삭제' border='0'  "
			Response.Write "  onMouseOver=javascript:this.src='../image/EIS/enotice/bu_r_del.gif'  "
			Response.write "  onMouseOut=javascript:this.src='../image/EIS/enotice/bu_del.gif' > "
			Response.Write " </a><span style='width:5'></span> " & vbcrlf 
		end if
		Response.Write "</td>"
		
		
		case "INSERT","MOD"	'새글작성,수정 
			Response.Write " <td width=110>&nbsp;</td> " & vbcrlf 
			Response.Write " <td align='right' > " & vbcrlf 
			Response.Write " <span style='width:2'></span>			" & vbcrlf 			
			Response.Write " <a href=" & to_where & " seq=" & seq & " &amp;page=" & page & "><img src='../image/EIS/enotice/bu_list.gif'                              alt='리스트' border='0'  onMouseOver=javascript:this.src='../image/EIS/enotice/bu_r_list.gif'   onMouseOut=javascript:this.src='../image/EIS/enotice/bu_list.gif' ></a>"		
			Response.Write " <A href='#'                                                    ><img src='../image/EIS/enotice/bu_save.gif' onClick='javascript:PostDate();'  alt='저장' border=0 onMouseOver=javascript:this.src='../image/EIS/enotice/bu_r_save.gif'  "
			Response.write "  onMouseOut=javascript:this.src='../image/EIS/enotice/bu_save.gif' ></a>"		& vbcrlf			
			Response.Write " </td>"			& vbcrlf 
		end select
	end if
	
	Response.Write "                    </tr>" & vbcrlf
	Response.Write "                   </table>"  & vbcrlf
	'-----------button-----e-----------------------------------------------------------
	Response.Write "	    </TD>" & vbcrlf
	Response.Write "      </TR>                 " & vbcrlf
	Response.Write "      <tr>" & vbcrlf
	Response.Write "		<td colspan=3 height='2' background='../../CShared/EISImage/Main/path_line.gif'></td>					" & vbcrlf
	Response.Write "	</tr>" & vbcrlf
	Response.Write "</TABLE>"  & vbcrlf
 end sub%>


