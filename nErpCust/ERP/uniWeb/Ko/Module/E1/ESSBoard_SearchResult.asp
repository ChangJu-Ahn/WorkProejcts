  <%	Option Explicit%>
<!-- #Include file="../../inc/incServer.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc"  -->
<!-- #Include file="../../inc/Adovbs.inc"  -->
<!-- #Include file="../../inc/incServerAdoDb.asp" -->
<!-- #Include file="../../inc/incSvrFuncSims.inc" -->
<!-- #Include file="../../inc/incSvrVarSims.inc"  -->
<!--#include file="ESSBoard_module_Gotopage.asp"-->
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
	
	Dim pageSize : pageSize = 7
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
<LINK REL="stylesheet" TYPE="Text/css" href="../../inc/ESS_board.css">
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
		oSearchImg[i].src = "../../../CShared/image/uniSIMS/s_" + oSearchImg[i].key + "_off.gif";
	}
	
	o.src = "../../../CShared/image/uniSIMS/s_" + o.key + "_on.gif";
	document.frmlist.searchPart.value= o.key;
}	

function initSearchPart()
{
	var oSearchImg = document.all.searchImg;
	for(var i=0; i < oSearchImg.length; i++)
	{
		if(oSearchImg[i].key == "<%=SearchPart_o%>")
			oSearchImg[i].src = "../../../CShared/image/uniSIMS/s_" + oSearchImg[i].key + "_on.gif";
		else
			oSearchImg[i].src = "../../../CShared/image/uniSIMS/s_" + oSearchImg[i].key + "_off.gif";
	}
	
	document.frmlist.searchPart.value= "<%=SearchPart_o%>";
}

function submit_searchFrom()
{
	var val = document.frmlist.searchStr.value;
	if (CheckStr(val, " ", "")==0) 
    {
      alert("검색할 단어를 입력해 주세요");
      document.frmlist.searchStr.value="";
      document.frmlist.searchStr.focus();
      return;
    }
        
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
<table bgcolor="white" cellpadding="0" cellspacing="0" height="100%" width="770">
<tr><td style="border-right:0 solid silver" align="center" valign="top">
<table cellpadding="0" cellspacing="0" bgcolor="white" width="700" style="margin-top:3" onSelectStart="javascript:return false;">	
<br>
	<tr width=100%>
		<td style="padding:10;padding-left:40;border:1 solid silver" colspan="5">
			- 검색 시간이 오래 걸릴 수 있습니다. 구체적으로 검색하시기 바랍니다.
		</td>
	</tr >
	<tr bgcolor="#3a6ea5" height="3"  width=100%>
		<td width="430" align="center" style="color:#eeeeee"></td>
		<td width="90" align="center" style="color:#eeeeee"></td>
		<td width="100" align="center" style="color:#eeeeee"></td>
		<td width="80" align="center" style="color:#eeeeee"></td>
	</tr>
			<tr bgcolor="#3a6ea5" height="23">
				<td CLASS="TDFAMILY_TITLE" ><img src="../../../CShared/image/uniSIMS/blank.gif" width="20" height="10">제목</td>
				<td CLASS="TDFAMILY_TITLE"    ALIGN="center">작성자</td>
				<td CLASS="TDFAMILY_TITLE"  ALIGN="center">작성날짜</td>
				<td CLASS="TDFAMILY_TITLE"  ALIGN="right">조회수<img src="../../../CShared/image/uniSIMS/blank.gif" width="10" height="10"></td>
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
		<tr bgcolor="<%=bgcolor%>" height="22">
		</tr>			
		<tr bgcolor="#F0F8FF" height="22">
			<td id="listXP<%=seq%>" onmouseover="javascript:mouseOnTD('<%=seq%>',true)" onmouseout="javascript:mouseOnTD('<%=seq%>',false)" onClick="javascript:go2Contnet('<%=seq%>');" >
				<img src="../../../CShared/image/uniSIMS/blank.gif" width="10" height="10"> <%=Subject%>
			</td>
			<td id="listXP<%=seq%>" ALIGN="center" onmouseover="javascript:mouseOnTD('<%=seq%>',true)" onmouseout="javascript:mouseOnTD('<%=seq%>',false)" onClick="javascript:go2Contnet('<%=seq%>');">
				<%=name%>
			</td>
			<td id="listXP<%=seq%>" ALIGN="center" onmouseover="javascript:mouseOnTD('<%=seq%>',true)" onmouseout="javascript:mouseOnTD('<%=seq%>',false)" onClick="javascript:go2Contnet('<%=seq%>');">
				<%=inpuTDate%>
			</td>
			<td id="listXP<%=seq%>" align="right" onmouseover="javascript:mouseOnTD('<%=seq%>',true)" onmouseout="javascript:mouseOnTD('<%=seq%>',false)" onClick="javascript:go2Contnet('<%=seq%>');">
				<%=Readcount%><span style="width:20"></span>
			</td>
		</tr>
	<%
			lgObjRs.MoveNext
			i = i + 1
		Loop
		
		else
	%>
	<tr>
		<td colspan="5" height="100" align="center" style="border:1 solid silver">
			검색된 결과가 없습니다<span style="width:100"></span>
		</td>
	</tr>	
	<%	End if %>
</table>
</td></tr>
<tr><td style="border-right:0 solid silver" align="center" valign="top">
	<table id="tblBottomBar"  HEIGHT="26" cellpadding="0" cellspacing="1" bgcolor="#99a9bc" width="700" style="margin-top:7">	
		<tr bgcolor="white" HEIGHT="26">
			<td bgcolor="white"  width=50% style="padding-left:15;padding-right:15;padding-bottom:;padding-top:5">
				<input name="searchStr" class="verdana" value style="width:110; height:18; padding:2; border:1 solid slategray">  
				<input type="hidden" name="searchPart" size="10" value="subject">
				<img src="../../../CShared/image/uniSIMS/ret1.jpg" align="absmiddle" WIDTH="26" HEIGHT="26" onClick="javascript:submit_searchFrom();">			
				<span style="width:10"></span>			
				<img id="searchImg" key="subject" style="cursor:hand" src="../../../CShared/image/uniSIMS/s_subject_on.gif" onClick="javascript:changeSearchPart(this);" WIDTH="51" HEIGHT="6">
				<span style="width:5"></span>
				<img id="searchImg" key="name" style="cursor:hand" src="../../../CShared/image/uniSIMS/s_name_off.gif" onClick="javascript:changeSearchPart(this);" WIDTH="33" HEIGHT="6">
				<span style="width:5"></span>
				<img id="searchImg" key="content" style="cursor:hand" src="../../../CShared/image/uniSIMS/s_content_off.gif" onClick="javascript:changeSearchPart(this);" WIDTH="51" HEIGHT="6">
			</td>
			<td   width=45% HEIGHT="28" >
			<div id="divPaging" align="right"><% GotoPageInSearchResult page, pagecount, searchPart, searchStr %>&nbsp;&nbsp;</div>
			</td>
			<td   width=28 HEIGHT="28" style="cursor:hand" align="center">
				<span style="width:2"></span>			
				<a href="ESSBoard_list.asp?seq=<%=seq%>&amp;page=<%=page%>"><img src="../../../CShared/image/uniSIMS/print1.jpg" alt="리스트보기" border="0"WIDTH="26" HEIGHT="26" ></a>								
			</td>			
		</tr>	
	</table>
</td></tr></table>		
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
