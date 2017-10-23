<%	Option Explicit  %>
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

	Dim userid : userid = gEmpNo

	Dim page : page = request("page")

	if page = "" then page = 1
	page = int(page)

	Dim pageSize : pageSize = 8
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
		oSearchImg[i].src = "../../CShared/ESSimage/s_" + oSearchImg[i].key + "_off.gif";
	}
	
	o.src = "../../CShared/ESSimage/s_" + o.key + "_on.gif";
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
<LINK REL="stylesheet" TYPE="Text/css" href="../ESSinc/common.css">
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<div id="divBody" style="display:none">
<table cellpadding="0" cellspacing="0" height="100%" width="731" height=100% border=0>
    <tr> 
      <td height=10></td>
    </tr>
	<tr height=10>
	  <td align="center" valign="top">
		<form name="search_form" method="post" action="ESSBoard_SearchResult.asp">
			<table id="tblBottomBar" width="731" HEIGHT="34" cellpadding="0" cellspacing="1" bgcolor="#DDDDDD" border=0>	
				<tr>
					<td height=30 bgcolor=F1F1F1>
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
							<% if gProAuth=0 then%>					
                            <td width=110>&nbsp;</td>
							<td align="right">
								<img align="absmiddle" style="cursor:hand" id="imgNew" src="../ESSimage/button_12.gif" <% if userid <> "" then %> onClick="javascript:location.href='ESSBoard_Frm_Insert.asp?page=<%=page%>';" <% else %> onClick="javascript:alert('로긴을 하셔야 합니다');" <%end if%> alt="새글작성" onMouseOver="javascript:this.src='../ESSimage/button_r_12.gif';" onMouseOut="javascript:this.src='../ESSimage/button_12.gif';">			
							</td>			
							<%end if%>						
                          </tr>
                      </table>
					</td>
				</tr>	
			</table>
		</form>		
	  </td>		
	</tr>
	<tr>
	  <td align="center" valign="top">
		<form name="frmlist" method="post" action="ESSBoard_SearchResult.asp">
		  <table border=0 cellpadding="0" cellspacing="0" width="731">	
		  <tr>
		  <td>
		  <table border=0 cellpadding="0" cellspacing="1" bgcolor="DDDDDD" width="731" height=270 onSelectStart="javascript:return false;">	
			<tr>
		        <td height=26 colspan=4 background=../../CShared/ESSimage/list_title_bg.gif>
					<table width=100% border=0 cellspacing=1 cellpadding=0>
		            <tr>
						<td CLASS="listitle02" width="460" height=26>제목</td>
						<td CLASS="listitle02" width="90">작성자</td>
						<td CLASS="listitle02" width="100">작성날짜</td>
						<td CLASS="listitle02" width="80">조회수</td>
					</tr>
					</table>
				</td>
			</tr>
		<%
	
			Dim i, bgcolor, seq, oid, id, subject, inpuTDate, Readcount, origInInputDate,name
			i = 1

			Do until lgObjRs.EOF
				if (i mod 2) = 0 then 
					bgcolor="F8F8F8"
				else
					bgcolor="#FFFFFF"
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
			<tr bgcolor="<%=bgcolor%>">
				<td class=listrow01 width="460" id="listXP<%=seq%>" onmouseover="javascript:mouseOnTD('<%=seq%>',true);" onmouseout="javascript:mouseOnTD('<%=seq%>',false);" onClick="javascript:go2Contnet('<%=seq%>');" >
					<%=Subject%></td>
				<td class=listrow01 width="90"  ALIGN="center" id="listXP<%=seq%>" onmouseover="javascript:mouseOnTD('<%=seq%>',true);" onmouseout="javascript:mouseOnTD('<%=seq%>',false);" onClick="javascript:go2Contnet('<%=seq%>');">
					<%=name%></td>
				<td class=listrow01 width="100" ALIGN="center" id="listXP<%=seq%>" align="center" onmouseover="javascript:mouseOnTD('<%=seq%>',true);" onmouseout="javascript:mouseOnTD('<%=seq%>',false);"  onClick="javascript:go2Contnet('<%=seq%>');">
					<%=inpuTDate%></td>
				<td class=listrow01 width="80" id="listXP<%=seq%>" align="center"  onmouseover="javascript:mouseOnTD('<%=seq%>',true);" onmouseout="javascript:mouseOnTD('<%=seq%>',false);"  onClick="javascript:go2Contnet('<%=seq%>');"  >
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
		    <tr bgcolor="#FFFFFF"> 
		      <td class=listrow01 width="460">&nbsp;</td>
		      <td class=listrow01 width="90">&nbsp;</td>
		      <td class=listrow01 width="100">&nbsp;</td>
		      <td class=listrow01 width="80">&nbsp;</td>
		    </tr>
		    <%next %>	
		  </table>
		  </td>		
		  </tr>
		 </table>
		</form>
	  </td>		
	</tr>
	<tr height=5>
	<td>
	  <table border=0 cellpadding="0" cellspacing="1" width="731" onSelectStart="javascript:return false;">	
	    <tr>
          <td align=center>
			<div id="divPaging" class=ftmvpage><% GoToPageDirectly page, pagecount %>&nbsp;&nbsp;</div>
	      </td>
	    </tr>
	  </table>
	 </td>		
	 </tr>
</table>	
</div>
</body>
</html>

<script language="javascript">
	document.all.divBody.style.display="";			//전제 body를 묶은 div를 나타나게 한다.
</script>
