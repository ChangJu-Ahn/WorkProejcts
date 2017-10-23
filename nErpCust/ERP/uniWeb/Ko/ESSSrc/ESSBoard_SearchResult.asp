  <%	Option Explicit%>
<!-- #Include file="../ESSinc/incServer.asp"  -->
<!-- #Include file="../ESSinc/lgsvrvariables.inc"  -->
<!-- #Include file="../ESSinc/Adovbs.inc"  -->
<!-- #Include file="../ESSinc/incServerAdoDb.asp" -->
<!-- #Include file="../ESSinc/incSvrFuncSims.inc" -->
<!-- #Include file="../ESSinc/incSvrVarSims.inc"  -->
<!--#include file="ESSBoard_module_Gotopage.asp"-->
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
	parent.document.All("nextprev").style.VISIBILITY = "hidden"
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
	Dim part: part =  right(table, len(table)-instr(table,"_"))

	Dim page : page = request("page")
	if page = "" then page = 1
	page = int(page)

	Dim SearchPart, SearchPart_o, SearchStr, SearchString
	SearchPart = Request("SearchPart")
	SearchPart_o = SearchPart
	SearchStr = Request("SearchStr")
	if len(SearchStr) > 0 then SearchStr = replace(SearchStr,"'", "''")
	
	SearchString = Split(SearchStr, "and")
	
	Dim pageSize : pageSize = 8
	Dim recordCount
	
	Call SubOpenDB(lgObjConn)  

	Dim pageCount
	lgStrSQL = "Select count(seq) from " & table & " Where 1=1 "
	for i=0 to Ubound(SearchString)
		lgStrSQL = lgStrSQL & " and " & SearchPart & " LIKE  " & FilterVar("%" & SearchString(i) & "%", "''", "S") & " "
		if SearchPart = "name" then
			lgStrSQL = lgStrSQL & " or id LIKE  " & FilterVar("%" & SearchString(i) & "%", "''", "S") & " "
		end if
	next
	if	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") =false then
		pageCount = 0	
	else
		pageCount=int((lgObjRs(0)-1)/pageSize)+1
		Call SubCloseDB(lgObjConn)
	end if

	lgStrSQL = "Select  seq, id, subject, inputDate, readCount,name from " & table & " Where 1=1 "
	for i=0 to Ubound(SearchString)
		lgStrSQL = lgStrSQL & " and " & SearchPart & " LIKE  " & FilterVar("%" & SearchString(i) & "%", "''", "S") & " "
		if SearchPart = "name" then
			lgStrSQL = lgStrSQL & " or id LIKE  " & FilterVar("%" & SearchString(i) & "%", "''", "S") & " "
		end if
	next
	lgStrSQL = lgStrSQL & " ORDER BY SEQ DESC"

	Call SubOpenDB(lgObjConn)  
	
	Set lgObjRs = Server.CreateObject("ADODB.Recordset")
	lgObjRs.Open lgStrSQL, lgObjConn, adOpenStatic, adLockReadOnly, adCmdText

	lgObjRs.PageSize = 5
			
%>	
<meta http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
<LINK REL="stylesheet" TYPE="Text/css" href="../ESSinc/common.css">
<script language="javascript">
var CheckedItems="";

function mouseOnTD(seq, bool)
{
	var oTD = eval("document.all.listXP" + seq);
	var len = oTD.length;
	var borderStyle = "1 solid slategray";
	
	if (bool){
		for(var i =0; i < len ; i++){
			oTD[i].style.borderTop = borderStyle;
			oTD[i].style.borderBottom = borderStyle;
			oTD[i].style.cursor = "default";
		}
		oTD[0].style.borderLeft = borderStyle;
		oTD[0].style.backgroundColor = "#b6c9d9";
		oTD[len-1].style.borderRight = borderStyle;
	}else{
		for(var i =0; i < len; i++){
			oTD[i].style.border = "";
		}
		oTD[0].style.backgroundColor = "";
	}
}

function mouseOverOnInfo(obj, bool)
{

	if (bool)
		obj.style.backgroundColor="#dddddd";
	else
		obj.style.backgroundColor="#ffffff";

}

function go2Contnet(seq)
{	
	location.href = "ESSBoard_listContent.asp?page=<%=page%>&seqs=" + seq+"&from_where=s&SearchPart=<%=SearchPart%>&SearchStr=<%=SearchStr%>";
}

function changeSearchPart(o)
{
	var oSearchImg = document.all.searchImg;
	for(var i=0; i < oSearchImg.length; i++)
	{
		oSearchImg[i].src = "../../CShared/ESSimage/s_" + oSearchImg[i].key + "_off.gif";
	}
	
	o.src = "../../CShared/ESSimage/s_" + o.key + "_on.gif";
	document.frmlist.searchPart.value= o.key;
}	

function initSearchPart()
{
	var oSearchImg = document.all.searchImg;
	for(var i=0; i < oSearchImg.length; i++)
	{
		if(oSearchImg[i].key == "<%=SearchPart_o%>")
			oSearchImg[i].src = "../../CShared/ESSimage/s_" + oSearchImg[i].key + "_on.gif";
		else
			oSearchImg[i].src = "../../CShared/ESSimage/s_" + oSearchImg[i].key + "_off.gif";
	}
	
	document.frmlist.searchPart.value= "<%=SearchPart_o%>";
}

function submit_searchFrom()
{
	var val = document.frmlist.searchStr.value;
        
	document.frmlist.submit();
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
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="javascript:initSearchPart();">
<form name="frmlist" method="post" action="ESSBoard_SearchResult.asp">
  <table cellpadding="0" cellspacing="0" height="100%" width="732">
	<tr> 
	  <td height=10></td>
	</tr>
	<tr>
	  <td align="center" valign="top">
		<table id="tblBottomBar" width="732" HEIGHT="34" cellpadding="0" cellspacing="1" bgcolor="#DDDDDD">	
			<tr>
				<td height=34 bgcolor=F1F1F1>
	              <table border=0 cellspacing=1 cellpadding=0>
	                  <tr> 
	                    <td width=18>&nbsp;</td>
						<td><input name="searchStr" class="form01" style=width:200px></td>
							<input type="hidden" name="searchPart" size="10" value="subject">
	                    <td width=5>&nbsp;</td>
						<td><img src="../ESSimage/button_01.gif" style="cursor:hand" align="absmiddle" onClick="javascript:submit_searchFrom();" onMouseOver="javascript:this.src='../ESSimage/button_r_01.gif';" onMouseOut="javascript:this.src='../ESSimage/button_01.gif';">			
	                    <td width=10>&nbsp;</td>
						<td><img id="searchImg" key="subject" style="cursor:hand" src="../../CShared/ESSimage/s_subject_on.gif" onClick="javascript:changeSearchPart(this);">
							<span style="width:5"></span>
						<td><img id="searchImg" key="name" style="cursor:hand" src="../../CShared/ESSimage/s_name_off.gif" onClick="javascript:changeSearchPart(this);">
							<span style="width:5"></span>
						<td><img id="searchImg" key="content" style="cursor:hand" src="../../CShared/ESSimage/s_content_off.gif" onClick="javascript:changeSearchPart(this);">
	                  </tr>
	              </table>
				</td>
			</tr>	
		</table>
	  </td>
	</tr>
	<tr> 
	  <td height=10></td>
	</tr>
	<tr>
	  <td align="center" valign="top">
		<table border=0 cellpadding="0" cellspacing="1" bgcolor="DDDDDD" width=732 height=280 onSelectStart="javascript:return false;">	
			<tr>
				<td CLASS="listitle02" width="460" height=30 background="../../CShared/ESSimage/list_title_bg.gif">제목</td>
				<td CLASS="listitle02" width="90" background="../../CShared/ESSimage/list_title_bg.gif">작성자</td>
				<td CLASS="listitle02" width="100" background="../../CShared/ESSimage/list_title_bg.gif">작성날짜</td>
				<td CLASS="listitle02" width="80" background="../../CShared/ESSimage/list_title_bg.gif">조회수</td>
			</tr>

		<%
		if	Not(lgObjRs.BOF and lgObjRs.EOF) then	
			lgObjRs.AbsolutePage = page

			Dim i, bgcolor, seq, id, subject, inputDate, readCount,name
			i = 1
			Do until lgObjRs.EOF or i > pageSize
				if (i mod 2) = 0 then 
					bgcolor="white"
				else
					bgcolor="white"
				end if
				
				seq = lgObjRs("seq")
				id = lgObjRs("id")
				subject = lgObjRs("subject")
				name =  lgObjRs("name")
				if Subject <> "" then Subject = Tag2Text(Subject)
				if len(Subject) > 30 then Subject = nLeft(Subject,60) 
				
				inputDate = FormaTDateTime(lgObjRs(3), 2) 
				Readcount = lgObjRs(4)
				%>
				<tr bgcolor="<%=bgcolor%>" height=24>
					<td bgcolor="FFFFFF" class=listrow01 width="460" height=24 id="listXP<%=seq%>" onmouseover="javascript:mouseOnTD('<%=seq%>',true);" onmouseout="javascript:mouseOnTD('<%=seq%>',false);" onClick="javascript:go2Contnet('<%=seq%>');" >
						<%=Subject%></td>
					<td class=listrow01 width="90"  ALIGN="center" id="listXP<%=seq%>" onmouseover="javascript:mouseOnTD('<%=seq%>',true);" onmouseout="javascript:mouseOnTD('<%=seq%>',false);" onClick="javascript:go2Contnet('<%=seq%>');">
						<%=name%></td>
					<td class=listrow01 width="100" ALIGN="center" id="listXP<%=seq%>" align="center" onmouseover="javascript:mouseOnTD('<%=seq%>',true);" onmouseout="javascript:mouseOnTD('<%=seq%>',false);"  onClick="javascript:go2Contnet('<%=seq%>');">
						<%=inpuTDate%></td>
					<td class=listrow01 width="80" ALIGN="center" id="listXP<%=seq%>" onmouseover="javascript:mouseOnTD('<%=seq%>',true);" onmouseout="javascript:mouseOnTD('<%=seq%>',false);"  onClick="javascript:go2Contnet('<%=seq%>');"  >
						<%=Readcount%></td>
				</tr>
				<%
				lgObjRs.MoveNext
				i = i + 1
			Loop

			Set lgObjRs = nothing

			Dim irestBlankCount
			for irestBlankCount=0 to pagesize-i
			%>
			    <tr bgcolor="<%=bgcolor%>" height=26> 
			      <td class=listrow01 height=26>&nbsp;</td>
			      <td class=listrow01>&nbsp;</td>
			      <td class=listrow01>&nbsp;</td>
			      <td class=listrow01>&nbsp;</td>
			    </tr>
			<% 
			next
		else 
			%>
			<tr>
				<td bgcolor="FFFFFF" colspan="5" align="center" style="font-family:'돋움';">
					검색된 결과가 없습니다.<span style="width:100"></span>
				</td>
			</tr>	
		<%End if %>
		</table>
		<table border=0 cellpadding="0" cellspacing="1" width="732" onSelectStart="javascript:return false;">	
		  <tr> 
		  	<td height=10></td>
		  </tr> 
		  <tr> 
	        <td align=center>
		  	<div id="divPaging" class=ftmvpage><% GoToPageDirectly page, pagecount %>&nbsp;&nbsp;</div>
		    </td>
		  </tr>
		</table>
	  </td>
    </tr>
  </table>		
</form>	
</body>
</html>
<%
Function nLeft(str,strcut)
    Dim bytesize, nLeft_count
    bytesize = 0

    For nLeft_count = 1 to len(str)
        if asc(mid(str,nLeft_count,1)) > 0 then '한글값은 0보다 작다 
            bytesize = bytesize + 1 '한글이 아닌경우 1Byte
        else
            bytesize = bytesize + 2 '한글인 경우 2Byte
        end if
        if strcut >= bytesize then nLeft = nLeft & mid(str,nLeft_count,1)  
            '끊고싶은 길이(Byte)만큼 
    Next
 
	if  nLeft <> "" then
		if len(str) > len(nLeft) then nLeft= left(nLeft,len(nLeft)-2) & "..."
      '문자열이 짤렸을 경우 뒤에 ...을 붙여줌 
    end if
End Function
%>
