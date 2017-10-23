  <%	Option Explicit  %>
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

	Dim userid : userid = gEmpNo

	Dim page : page = request("page")

	if page = "" then page = 1
	page = int(page)

	Dim pageSize : pageSize = 7
	Dim recordCount, recentCount

    Call SubOpenDB(lgObjConn)  	
    lgStrSQL = "Select count(seq) seq_count from " & table

	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = true Then
		recordCount	   = lgObjRs("seq_count")
		if recordCount=0 then
		  recordCount	   = 1
		end if
	End IF  

    lgStrSQL = " SELECT count(*) datediff_count FROM " & table & " WHERE DATEDIFF(mi, inputDate, getdate()) <1440  "
	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = true Then
		recentCount	   = lgObjRs("datediff_count")
	End IF  

	Dim n_from,n_to,s_desc,s_asc,n_range, n_limit
	Dim pageCount : pageCount=int((recordCount-1)/pageSize)+1
	if page > pageCount then page = pageCount
 
	//계층형 로직을 위한 부분 by Cassatt	
	s_desc="desc"
	s_asc="asc"
  
	n_from = (page-1) * pageSize
	n_to = n_from + pageSize

	if n_to > recordCount then n_to = recordCount
  
	if page > int(pageCount/2)+1 then
        Dim t : t=n_to
		n_to=recordCount-n_from
		n_from=recordCount-t
		s_desc="asc"	
		s_asc="desc"
	end if

	n_range= n_to			'StartPos
	n_limit= n_to-n_from  
   
	lgStrSQL = "SELECT seq, id, subject, inputdate, readcount,  name  FROM " & table 
	lgStrSQL = lgStrSQL  & " WHERE seq IN ( SELECT TOP  " & n_limit & " seq  from  ( SELECT TOP  " & n_range
	lgStrSQL = lgStrSQL  & "  seq from " & table & " ORDER BY seq " & s_desc  
	lgStrSQL = lgStrSQL  & " ) as a  ORDER BY seq  asc ) ORDER BY seq desc"
'Response.Write 	"**lgStrSQL:"  & lgStrSQL

	if FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X")=true then 
		if lgObjRs.BOF and lgObjRs.EOF then Response.Redirect "ESSBoard_Frm_Insert.asp"
	end if
	
%>

<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="javascript">
var CheckedItems="";

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
function DoChecking(seq){
	eval("document.frmlist.img" + seq + ".click()");
}

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
		oTD[0].style.backgroundColor = "white";
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
	location.href = "ESSBoard_listContent.asp?page=<%=page%>&seqs=" + seq;
}

function showSearchForm(obj){

	document.search_form.searchStr.focus();
		
}

function changeSearchPart(o)
{
	var oSearchImg = document.all.searchImg;
	
	for(var i=0; i < oSearchImg.length; i++)
	{
		oSearchImg[i].src = "../../../CShared/image/uniSIMS/s_" + oSearchImg[i].key + "_off.gif";
	}
	
	o.src = "../../../CShared/image/uniSIMS/s_" + o.key + "_on.gif";
	document.search_form.searchPart.value= o.key;
}	

function submit_searchFrom()
{
	var val = document.search_form.searchStr.value;
	if (CheckStr(val, " ", "")==0) 
    {
      alert("검색할 단어를 입력해 주세요");
      document.search_form.searchStr.value="";
      document.search_form.searchStr.focus();
      return;
    }
        
	document.search_form.submit();
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
<LINK REL="stylesheet" TYPE="Text/css" href="../../inc/ESS_board.css">
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<div id="divBody" style="display:none">
<table bgcolor="white" cellpadding="0" cellspacing="0" height="100%" width="770">
	<tr><td align="center" valign="top">
	<table cellpadding="0" cellspacing="0" BGCOLOR="#5A9CD6" width="100%"  onSelectStart="javascript:return false;" style="margin-top:0">
		<tr height="30">
			<td width="100%" align="center" style="padding-left:;border-bottom:1 solid gray;color:#2E4287;FONT-WEIGHT: bolder;FONT-SIZE: 10pt;" background="../../../CShared/image/uniSIMS/skyback_03.gif" >
			공지사항</td>
		</tr>
	</table>
	<form name="frmlist" method="post" action="ESSBoard_SearchResult.asp">


	<table cellpadding="0" cellspacing="0" bgcolor="white" width="700" style="margin-top:3" onSelectStart="javascript:return false;">	
		<table cellpadding="0" cellspacing="1"  bgcolor="#99a9bc" width="700" style="margin-top:3" onSelectStart="javascript:return false;">	
			<tr bgcolor="#3a6ea5" height="23">
				<td CLASS="TDFAMILY_TITLE" width="430"><img src="../../../CShared/image/uniSIMS/blank.gif" width="20" height="10">제목</td>
				<td CLASS="TDFAMILY_TITLE"  width="90"  ALIGN="center">작성자</td>
				<td CLASS="TDFAMILY_TITLE"  width="100" ALIGN="center">작성날짜</td>
				<td CLASS="TDFAMILY_TITLE" width="80" ALIGN="right">조회수<img src="../../../CShared/image/uniSIMS/blank.gif" width="10" height="10"></td>
			</tr>
		</table>
		<table cellpadding="0" cellspacing="0"  bgcolor="white" width="700" style="margin-top:3" onSelectStart="javascript:return false;" >				
	<%
	
		Dim i, bgcolor, seq, oid, id, subject, inpuTDate, Readcount, origInInputDate,name
		i = 1

		Do until lgObjRs.EOF
			if (i mod 2) = 0 then 
				bgcolor="white"
			else
				bgcolor="#F0F8FF"
			end if
			
			seq = lgObjRs(0)
			id = lgObjRs(1) : oid = id
			
			subject = lgObjRs(2)
			if Subject <> "" then Subject = Tag2Text(Subject)
			if len(Subject) > 30 then Subject = nLeft(Subject,60) 
			
			origInInputDate = lgObjRs(3)
			inpuTDate = UNIDateClientFormat(origInInputDate)
			Readcount = lgObjRs(4)
			name = lgObjRs(5)
			if len(id) > 5 then id = nLeft(id,10) 
	%>
		<tr bgcolor="<%=bgcolor%>" height="22">

			<td  width="430" id="listXP<%=seq%>" onmouseover="javascript:mouseOnTD('<%=seq%>',true);" onmouseout="javascript:mouseOnTD('<%=seq%>',false);" onClick="javascript:go2Contnet('<%=seq%>');" >
				<img src="../../../CShared/image/uniSIMS/blank.gif" width="20" height="10">
					<%=Subject%>
			</td>
			<td  width="90"  ALIGN="center" id="listXP<%=seq%>" onmouseover="javascript:mouseOnTD('<%=seq%>',true);" onmouseout="javascript:mouseOnTD('<%=seq%>',false);" onClick="javascript:go2Contnet('<%=seq%>');">
				<%=name%>	</td>
			<td width="100" ALIGN="center" id="listXP<%=seq%>" align="center" onmouseover="javascript:mouseOnTD('<%=seq%>',true);" onmouseout="javascript:mouseOnTD('<%=seq%>',false);"  onClick="javascript:go2Contnet('<%=seq%>');">
				<%=inpuTDate%>
			</td>
			<td width="80" id="listXP<%=seq%>" align="right" onmouseover="javascript:mouseOnTD('<%=seq%>',true);" onmouseout="javascript:mouseOnTD('<%=seq%>',false);"  onClick="javascript:go2Contnet('<%=seq%>');"  >
				<%=Readcount%>&nbsp;&nbsp;
			</td>
		</tr>	
	<%
	
			lgObjRs.MoveNext
			i = i + 1
		Loop
		
		Set lgObjRs = nothing

		Dim irestBlankCount
		for irestBlankCount=0 to pagesize-i
	%>
		<tr height="22">
			<td bgcolor="white" width="20&quot;">&nbsp;</td>
			<td colspan="4">&nbsp;</td>
		</tr>
	<% next %>	
		</table>
</form>		

<form name="search_form" method="post" action="ESSBoard_SearchResult.asp">
	<table id="tblBottomBar"  HEIGHT="26" cellpadding="0" cellspacing="1" bgcolor="#99a9bc" width="700" style="margin-top:7">	
		<tr bgcolor="white" HEIGHT="26">
			<td bgcolor="white" width=50% style="padding-left:15;padding-right:15;padding-bottom:;padding-top:5">
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
			<td align="right" width=45% HEIGHT="28">
				<div id="divPaging"><% GoToPageDirectly page, pagecount %>&nbsp;&nbsp;</div>
			</td>
<% if gProAuth=0 then%>					
			<td align="center" width=28 HEIGHT="28"  style="cursor:hand">
				<span style="width:2"></span>			
				<img align="absmiddle" border=0 id="imgNew" src="../../../CShared/image/uniSIMS/add1.jpg" WIDTH="26" HEIGHT="26"  <% if userid <> "" then %> onClick="javascript:location.href='ESSBoard_Frm_Insert.asp?page=<%=page%>';" <% else %> onClick="javascript:alert('로긴을 하셔야 합니다');" <%end if%> alt="새글작성">			
			</td>			
<%end if%>						
		</tr>	
	</table>
</form>		
	</table>
</table>	
</div>
</body>
</html>

<script language="javascript">
	document.all.divBody.style.display="";			//전제 body를 묶은 div를 나타나게 한다.
</script>
