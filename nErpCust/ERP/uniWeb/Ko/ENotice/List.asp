<%	Option Explicit  %>
<!-- #Include file="../inc/incServer.asp"  -->
<!-- #Include file="../inc/lgsvrvariables.inc"  -->
<!-- #Include file="../inc/Adovbs.inc"  -->
<!-- #Include file="../inc/incServerAdoDb.asp" -->
<!-- #Include file="../inc/incSvrFuncSims.inc" -->
<!-- #Include file="../inc/incSvrVarSims.inc"  -->
<!--#include file="Gotopage.asp"-->
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

	Dim userid : userid =gUsrId

	Dim page : page = request("page")

	if page = "" then page = 1
	page = int(page)

	Dim pageSize : pageSize = 12
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
 
	'//계층형 로직을 위한 부분 by Cassatt	
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

	if FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X")=true then 
		if lgObjRs.BOF and lgObjRs.EOF then Response.Redirect "Frm_Insert.asp"
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
	location.href = "ListContent.asp?page=<%=page%>&seqs=" + seq;
}

function showSearchForm(obj){

	document.search_form.searchStr.focus();
		
}

function changeSearchPart(o)
{
	var oSearchImg = document.all.searchImg;
	
	for(var i=0; i < oSearchImg.length; i++)
	{
		oSearchImg[i].src = "../../CShared/EISImage/ENotice/s_" + oSearchImg[i].key + "_off.gif";
	}
	
	o.src = "../../CShared/EISImage/ENotice/s_" + o.key + "_on.gif";
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
<body leftmargin="20" topmargin="0" marginwidth="0" marginheight="0">
<div id="divBody" style="display:none">
<table cellpadding="0" cellspacing="0" height="100%" width="100%" border=0>
    <tr> 
      <td height=10></td>
    </tr>
    <tr>
		<td>
		<div id="divTitle" ><% gotoTitle "LIST"%></div> 
		</td>
    </tr>
	<tr>
	  <td align="center" valign="top">
			<table id="tblBottomBar" width="100%" HEIGHT="34" cellpadding="0" cellspacing="1" bgcolor="E6E4CD" border=0>	
				<tr>
					<td height=34 bgcolor=F2F1E4>
                      <table border=0 cellspacing=1 cellpadding=0>
                        <form name="search_form" method="post" action="Search.asp">
                          <tr> 
                            <td width=18>&nbsp;</td>
							<td><input name="searchStr" class="form01" style=width:200px></td>
								<input type="hidden" name="searchPart" size="10" value="subject">
                            <td width=5>&nbsp;</td>
							<td>
							<img src="../image/EIS/enotice/bu_05.gif"   style="cursor:hand" onClick="javascript:submit_searchFrom()" onMouseOver="javascript:this.src='../image/EIS/enotice/bu_r_05.gif';" onMouseOut="javascript:this.src='../image/EIS/enotice/bu_05.gif';"></td>
                            <td width=10>&nbsp;</td>
							<td valign="bottom"><img id="searchImg" key="subject" style="cursor:hand" src="../../CShared/EISImage/ENotice/s_subject_on.gif" onClick="javascript:changeSearchPart(this);">
								<span style="width:5"></span>
							<td valign="bottom"><img id="searchImg" key="name" style="cursor:hand" src="../../CShared/EISImage/ENotice/s_name_off.gif" onClick="javascript:changeSearchPart(this);">
								<span style="width:5"></span>
							<td valign="bottom"><img id="searchImg" key="content" style="cursor:hand" src="../../CShared/EISImage/ENotice/s_content_off.gif" onClick="javascript:changeSearchPart(this);">						
                          </tr>
                        </form>		
                      </table>
					</td>
				</tr>					
			</table>			
	  </td>		
	</tr>
	<tr>
	  <td align="center" valign="top">
		<form name="frmlist" method="post" action="Search.asp">
		  <table border=0 cellpadding="0" cellspacing="0" width="100%">	
		  <tr>
		  <td>
		  <table border=0 cellpadding="0" cellspacing="1" bgcolor="#dddddd" width="100%" height=290 onSelectStart="javascript:return false;">	
			<tr>
		        <td height=28 colspan=5 background=../../CShared/EISImage/ENotice/list_title_bg.gif>
					<table width=100% border=0 cellspacing=1 cellpadding=0>
		            <tr>
						<td CLASS="listitle02" width="7%" height=28>번호</td>
						<td CLASS="listitle02" width="60%" height=28>제목</td>
						<td CLASS="listitle02" width="13%">작성자</td>
						<td CLASS="listitle02" width="10%">작성날짜</td>
						<td CLASS="listitle02" width="10%">조회수</td>
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
				<td class=listrow01 width="7%" ><%=i%></td>
				<td   width="60%" id="listXP<%=seq%>" onmouseover="javascript:mouseOnTD('<%=seq%>',true);" onmouseout="javascript:mouseOnTD('<%=seq%>',false);" onClick="javascript:go2Contnet('<%=seq%>');" >
					&nbsp;<%=Subject%></td>
				<td class=listrow01 width="13%"  ALIGN="center" id="listXP<%=seq%>" onmouseover="javascript:mouseOnTD('<%=seq%>',true);" onmouseout="javascript:mouseOnTD('<%=seq%>',false);" onClick="javascript:go2Contnet('<%=seq%>');">
					<%=name%></td>
				<td class=listrow01 width="10%" ALIGN="center" id="listXP<%=seq%>" align="center" onmouseover="javascript:mouseOnTD('<%=seq%>',true);" onmouseout="javascript:mouseOnTD('<%=seq%>',false);"  onClick="javascript:go2Contnet('<%=seq%>');">
					<%=inpuTDate%></td>
				<td class=listrow01 width="10%" id="listXP<%=seq%>" align="right" onmouseover="javascript:mouseOnTD('<%=seq%>',true);" onmouseout="javascript:mouseOnTD('<%=seq%>',false);"  onClick="javascript:go2Contnet('<%=seq%>');"  >
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
		    <tr bgcolor="<%=bgcolor%>"> 
			  <td class=listrow01>&nbsp;</td>
		      <td class=listrow01>&nbsp;</td>
		      <td class=listrow01>&nbsp;</td>
		      <td class=listrow01>&nbsp;</td>
		      <td class=listrow01>&nbsp;</td>		      
		    </tr>
		  <%next %>	
		  </table>
		  </td>		
		</tr>
		<tr>
		<td>
		  <table border=0 cellpadding="0" cellspacing="1" width="100%" onSelectStart="javascript:return false;">	
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
	  </td>		
	</tr>
</table>	
</div>
</body>
</html>

<script language="javascript">
	document.all.divBody.style.display="";			//전제 body를 묶은 div를 나타나게 한다.
</script>
