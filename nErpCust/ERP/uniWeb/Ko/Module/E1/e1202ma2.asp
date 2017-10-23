<%@ LANGUAGE="VBSCRIPT" %>
<% Response.Expires = -1%>
<!--

======================================================================================================
*  1. Module Name          : Human Resources
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strTitle")%></TITLE>
<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../../inc/IncServer.asp"  -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" href="../../inc/CommStyleSheet.css">

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/ccm.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCommFunc.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEvent.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/adoQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<!-- #Include file="../../inc/incSvrVarSims.inc"  -->
<!-- #Include file="../../inc/incSvrFuncSims.inc" -->
<%
	dim lgEmpn_no
	dim lgPay_Yymm
	dim lgProv_Type
	
	lgEmp_no = request("Emp_no")
	lgPay_Yymm = request("Pay_Yymm")
	lgProv_Type = request("Prov_Type")

%>
<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance


'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID      = "e1202mb2.asp"						           '☆: Biz Logic ASP Name
Const C_SHEETMAXROWS = 10	                                      '☜: Visble row
Const C_SHEETMAXCOLS = 10
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
'--------------------------------------------------------------------------------------------------------
'  Constants for SpreadSheet #1
'--------------------------------------------------------------------------------------------------------

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" --> 
<!-- #Include file="../../inc/incGrid.inc" -->
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim Grid1

Dim gDecimal_day
Dim gDecimal_time
'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================
'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
    lgKeyStream = "<%=lgEmp_no%>" & gColSep
    lgKeyStream = lgKeyStream & "<%=lgPay_Yymm%>" & gColSep
    lgKeyStream = lgKeyStream & "<%=lgProv_Type%>" & gColSep
   
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim strTemp
    Dim lgstrData
    strTemp = "<%=lgProv_Type%>"
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
     Call CommonQueryRs(" MINOR_NM "," B_MINOR ", " MAJOR_CD = " & FilterVar("H0040", "''", "S") & " AND MINOR_CD =  " & FilterVar(strTemp , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    if  Replace(lgF0, Chr(11), "") = "" then
        lgstrData = ""
    else
        lgstrData = Replace(lgF0, Chr(11), "")
    end if

		document.frm1.Prov_nm.value = lgstrData
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub
'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitGrid()
End Sub


'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	<% Call loadInfTB19029(gCurrency, "Q", "H") %>

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
'    parent.document.All("nextprev").style.VISIBILITY = "hidden"
    Call SetToolBar("0000")    
    Call LoadInfTB19029()
    call get_decimal()
    Call InitComboBox()
 '   Call LayerShowHide(0)

'    Call InitGrid()
'    Call LockField(Document)
    Call DbQuery()
    
    
End Sub
'========================================================================================
' Function Name : Form_UnLoad
' Function Desc : 페이지 전환이나 화면이 닫힐 경우 실행해야 될 로직 처리 
'========================================================================================
Sub Form_UnLoad()
	On Error Resume Next
    Set Grid1 = Nothing
End Sub

Function DbQuery()
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG
	Call LayerShowHide(1)
    Call MakeKeyStream("Q")

    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
'    strVal = BIZ_PGM_ID & "?Emp_no=<%=lgEmp_no%>"                     '☜: Query
'    strVal = strVal     & "&Pay_Yymm=<%=lgPay_Yymm%>"                   '☜: Query Key
'    strVal = strVal     & "&Prov_Type=<%=lgProv_Type%>"                   '☜: Query Key
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
 
    DbQuery = True                                                               '☜: Processing is NG
 
End Function

'========================================================================================
' Function Name : fncGoBack
' Function Desc : 이전페이지로 이동한다.
'========================================================================================
Function fncGoBack()
	dim strVal
	strVal = "e1202ma1.asp?strTitle=월급여&year=<%=left(lgPay_Yymm,4)%>"
	
 document.location = strVal

End Function


Function DbQueryOk()
    Err.Clear                                                                    '☜: Clear err status

End Function

Function DbQueryFail()
    Err.Clear
   
End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()
	                                                         '☜: Processing is NG
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
	
End Function



Function GetRow(pRow)
	
End Function

Sub SubSum()                      '// MB에서 데이터를 넘겨받는다.
       
End Sub



Function DoubleGetRow(pRow)
   
End Function

Sub MouseRow(pRow)
	
End Sub


'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================

Sub Query_OnClick()
    Call DbQuery(1)
End Sub

Sub Print_onClick()
    Call SubPrint(MyBizASP)
End Sub
'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================%>
Function FncBtnPrint() 

	Dim strUrl
	Dim lngPos
	Dim intCnt
	Dim StrEbrFile
	dim prov_tp
	
	prov_tp = "<%=lgProv_Type%>"
	if prov_tp = "1" then
		StrEbrFile = "h6016oa1.ebr"
	else
		StrEbrFile = "h6016oa1_1.ebr"
	end if
	

'    On Error Resume Next                                                    '☜: Protect system from crashing

    <%'--출력조건을 지정하는 부분 수정 %>
	strUrl = strUrl & "Pay_Yymm|<%=lgPay_Yymm%>"
	strUrl = strUrl & "|Prov_Type|<%=lgProv_Type%>"
	strUrl = strUrl & "|Pay_cd|" & "%"
	strUrl = strUrl & "|Fr_Dept_cd|1" 
	strUrl = strUrl & "|To_Dept_cd|zzzzzzzzz"
	strUrl = strUrl & "|Emp_no|<%=lgEmp_no%>"
	strUrl = strUrl & "|gDecimal_day|" & gDecimal_day
	strUrl = strUrl & "|gDecimal_time|" & gDecimal_time
		<%'--출력조건을 지정하는 부분 수정 - 끝 %>

	call FncEBRPrint(EBAction , StrEbrFile , strUrl)

'msgbox "print"

End Function

Sub GRID_PAGE_OnChange()
End Sub

Sub DELETE_OnClick()
    Call Grid1.DeleteClick()
End Sub

Sub CANCEL_OnClick()
    Call Grid1.CancelClick()
End Sub

Sub get_decimal()
    Dim intRetCd
    
	gDecimal_day = 0
	gDecimal_time = 0

	IntRetCd = CommonQueryRs(" DECI_PLACE "," HDA041T "," ATTEND_TYPE = " & FilterVar("1", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	If IntRetCd = True Then
	    gDecimal_day  = Trim(Replace(lgF0,Chr(11),""))
	End If

	IntRetCd = CommonQueryRs(" DECI_PLACE "," HDA041T "," ATTEND_TYPE = " & FilterVar("2", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	If IntRetCd = True Then
	    gDecimal_time  = Trim(Replace(lgF0,Chr(11),""))
	End If

End Sub

</SCRIPT>
 <!-- #Include file="../../inc/uniSimsClassID.inc" --> 

<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->
</HEAD>

	

<BODY topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">

    <TABLE cellSpacing=0 cellPadding=0 width=700 border=0 bgcolor=#ffffff>
        <TR><TD WIDTH="100%" HEIGHT=0><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=0 FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME></TD></TR>
    </TABLE>
<br>
    <TABLE cellSpacing=0 cellPadding=0 width=600 border=0 bgcolor='#ffffff' align = center>
        <TR height=30 bgcolor='#ffffff' align = center  width=600>
            <TD bgcolor='#d0d6e4'valign=middle align='right'  width='250'>
            		<font size='2' color='#29499C'><b><%=left(lgPay_Yymm,4)%>&nbsp;년 &nbsp;<%=right(lgPay_Yymm,2)%>&nbsp;월<b></font>
            </TD>
            <TD bgcolor='#d0d6e4'valign=middle align='right'>
								<INPUT NAME="prov_nm"  class=Base1 style='text-align:right;color:#29499C;FONT-WEIGHT: bolder;' SiZE=7 readonly>            </TD>
            <TD bgcolor='#d0d6e4'valign=middle align='left'>
            		<font size='2' color='#29499C'><b>지급명세서<b></font>
            </TD>
        </TR>
    </TABLE>
<BR>
    <TABLE cellSpacing=1 cellPadding=1 border=0 bgcolor='#ffffff'  align = center  width=600>
        <TR height=20 bgcolor='#E9EDF9' align = center>
            <TD width=150 bgcolor='#d0d6e4'><font color='#29499C'><b>성명</b></font></TD>
            <TD width=150>
            <INPUT NAME="name"  class=Base1 style=';text-align:center;color:#29499C' SiZE=15 readonly>
            </TD>
            <TD width=150 bgcolor='#d0d6e4'><font color='#29499C'><b>사번</b></font></TD>
            <TD width=150>
            <INPUT class=Base1 NAME="emp_no" style='text-align:center;color:#29499C' SiZE=20 readonly>
            </TD>
        </TR>
        <TR height=20 bgcolor='#E9EDF9' align = center>
            <TD width=150 bgcolor='#d0d6e4'><font color='#29499C'><b>직급</b></font></TD>
            <TD width=150>
            <INPUT class=Base1 NAME="grade" style='text-align:center;color:#29499C' SiZE=15 readonly>
            </TD>
            <TD width=150 bgcolor='#d0d6e4'><font color='#29499C'><b>부서</b></font></TD>
            <TD width=150>
            <INPUT class=Base1 NAME="dept_cd" style='text-align:center;color:#29499C' SiZE=20 readonly>
            </TD>
        </TR>
    </TABLE>
    <br>
    <TABLE cellSpacing=1 cellPadding=0 border=0 bgcolor='#ffffff'  align = center width=600>
        <TR height=20 bgcolor='#E9EDF9' align = center>
            <TD width=150 bgcolor='#d0d6e4'><font color='#29499C'><b>근무일수</b></font></TD>
            <TD width=150>
                 <INPUT class=Base2 NAME="work_day" style='text-align:right;color:#29499C' SiZE=15 readonly>&nbsp;&nbsp;
						</TD>
            <TD width=150 bgcolor='#d0d6e4'>
                    <INPUT class=Base1 NAME="work_nm1" SiZE=15 style='text-align:center;color=#29499C' readonly>
             </TD>
            <TD width=150>
                    <INPUT class=Base1  NAME="work_hh1" SiZE=10 style='text-align:right;color:#29499C' readonly>
                    <INPUT class=Base1  NAME="work_mm1" SiZE=6 style='text-align:right;color:#29499C' readonly>
            </TD>
        </TR>
        <TR height=20 bgcolor='#E9EDF9' align = center>
            <TD width=150 bgcolor='#d0d6e4'>&nbsp;</TD>
            <TD width=150>&nbsp;</TD>
            <TD width=150 bgcolor='#d0d6e4'>
                    <INPUT class=Base2 NAME="work_nm2" SiZE=15 style='text-align:center;color:#29499C' readonly>
             </TD>
            <TD width=150>
                    <INPUT class=Base2  NAME="work_hh2" SiZE=10 style='text-align:right;color:#29499C' readonly>
                    <INPUT class=Base2  NAME="work_mm2" SiZE=6 style='text-align:right;color:#29499C' readonly>
            </TD>
        </TR>
        <TR height=20 bgcolor='#E9EDF9' align = center>
            <TD width=150 bgcolor='#d0d6e4'>&nbsp;</TD>
            <TD width=150>&nbsp;</TD>
            <TD width=150 bgcolor='#d0d6e4'>
                    <INPUT class=Base1 NAME="work_nm3" SiZE=15 style='text-align:center;color:#29499C' readonly>
             </TD>
            <TD width=150>
                    <INPUT class=Base1  NAME="work_hh3" SiZE=10 style='text-align:right;color:#29499C' readonly>
                    <INPUT class=Base1  NAME="work_mm3" SiZE=6 style='text-align:right;color:#29499C' readonly>
            </TD>
        </TR>
        <TR height=20 bgcolor='#E9EDF9' align = center>
            <TD width=150 bgcolor='#d0d6e4'>&nbsp;</TD>
            <TD width=150>&nbsp;</TD>
            <TD width=150 bgcolor='#d0d6e4'>
                    <INPUT class=Base1 NAME="work_nm4" SiZE=15 style='text-align:center;color:#29499C' readonly>
             </TD>
            <TD width=150>
                    <INPUT class=Base1  NAME="work_hh4" SiZE=10 style='text-align:right;color:#29499C' readonly>
                    <INPUT class=Base1  NAME="work_mm4" SiZE=6 style='text-align:right;color:#29499C' readonly>
            </TD>
        </TR>
	  </TABLE>    
    <br>
    <TABLE cellSpacing=1 cellPadding=1 border=0 bgcolor='#ffffff'  align = center width='600'>
        <TR height=20 bgcolor='#d0d6e4' align = center>
            <TD width=300><font color='#29499C'><b>지급내역</b></font></TD>
            <TD width=300><font color='#29499C'><b>공제내역</b></font></TD>
        </TR>
        <TR bgcolor='#E9EDF9' align = center  valign=top>
            <TD width=300>
            		<table width=290  cellSpacing=0 cellPadding=0 >
          				<%For i = 0 to 11%>
            			<tr>
            				<td width=60%>
        							&nbsp;<INPUT class=Base1 NAME="pay_nm<%=i%>" SiZE=20 style='text-align:left;color:#29499C' readonly>
            				</td>
            				<td width=40%>&nbsp;
        							<INPUT class=Base1 NAME="pay_amt<%=i%>" SiZE=17 style='text-align:right;color:#29499C' readonly>
        						</td>
            			</tr>
            			<%Next%>
            		</table>            </TD>
            <TD width=300>
            		<table width=290  cellSpacing=0 cellPadding=0>
          				<%For i = 0 to 11%>
            			<tr>
            				<td width=60%>
        							&nbsp;<INPUT class=Base1 NAME="sub_nm<%=i%>" SiZE=20 style='text-align:left;color:#29499C' readonly>
            				</td>
            				<td width=40%>&nbsp;
        							<INPUT class=Base1 NAME="sub_amt<%=i%>" SiZE=17 style='text-align:right;color:#29499C' readonly>
        						</td>
            			</tr>
            			<%Next%>
            		</table>
            </TD>
        </TR>
    </TABLE>
    <TABLE cellSpacing=1 cellPadding=1 border=0 bgcolor='#ffffff' width=600 align = center>
        <TR height=20 bgcolor='#ffffff' width=600>
     				<td bgcolor='#d0d6e4' width = 30% align=center><font color='#29499C'><b>지급총액</b></font></td>
     				<td align = 'left' bgcolor='#E9EDF9' width = 20%>&nbsp;&nbsp;&nbsp;
     				   <INPUT NAME="pay_tot" class=Base1 SiZE=18 style='text-align:right;color:#29499C' readonly>&nbsp;
     				</td>
     				<td bgcolor='#d0d6e4' width = 30% align=center><font color='#29499C'><b>공제총액</b></font></td>
     				<td align = 'left' bgcolor='#E9EDF9' width = 20%>&nbsp;&nbsp;&nbsp;
     				   <INPUT NAME="sub_tot" class=Base1 SiZE=18 style='text-align:right;color:#29499C'  readonly>&nbsp;
     				</td>
        </TR>
    </TABLE>
     <TABLE cellSpacing=1 cellPadding=1 border=0 bgcolor='#ffffff'  align = center  width=600>
         <TR height=20 bgcolor='#E9EDF9' align = center width=600>
             <TD width=200 bgcolor='#d0d6e4'><font color='#29499C'><b>실지급액</b></font></TD>
             <TD width=400  align = right>
 		 				   <INPUT class=Base1 NAME="real" SiZE=25 style='text-align:right;color:#29499C'  readonly>&nbsp;
             </TD>
         </TR>
     </TABLE>
	  <br>
    <TABLE border=0  align = center>
        <TR align = center>
        <td>
    			<INPUT style="WIDTH: 150px; HEIGHT: 20px" TYPE=button NAME=printprev VALUE="출력" OnClick="Vbscript: call FncBtnPrint()">
    			<INPUT style="WIDTH: 150px; HEIGHT: 20px" TYPE=button NAME=printprev VALUE="돌아가기" OnClick=":: call fncGoBack()">
    		</td>
    		</tr>
    </table>

<br>
<br>
    <INPUT TYPE=hidden NAME="txtMode">
    <INPUT TYPE=hidden NAME="txtKeyStream">
    <INPUT TYPE=hidden NAME="txtUpdtUserId">
    <INPUT TYPE=hidden NAME="txtInsrtUserId">
    <INPUT TYPE=hidden NAME="txtFlgMode">
    <INPUT TYPE=hidden NAME="txtPrevNext">
    
    <INPUT TYPE=hidden NAME=GRID_TOTPAGES >
    <INPUT TYPE=hidden NAME=GRID_PAGE value=1 >
</FORM>	
<FORM NAME="EBAction" TARGET = "MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname">
	<INPUT TYPE="HIDDEN" NAME="dbname">
	<INPUT TYPE="HIDDEN" NAME="filename">
	<INPUT TYPE="HIDDEN" NAME="condvar">
	<INPUT TYPE="HIDDEN" NAME="date">
</FORM></BODY>
</HTML>
