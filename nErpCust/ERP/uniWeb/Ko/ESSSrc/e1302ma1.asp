<%@ LANGUAGE="VBSCRIPT" %>
<% Response.Expires = -1%>
<!--
'======================================================================================================
*  1. Module Name          : Human Resources
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strTitle")%></TITLE>


<!-- #Include file="../ESSinc/IncServer.asp"  -->

<LINK REL="stylesheet" TYPE="Text/css" href="../ESSinc/common.css">


<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/ccm.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incCookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incCommFunc.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incEvent.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/adoQuery.vbs"></SCRIPT>
<!-- #Include file="../ESSinc/incSvrVarSims.inc"  -->
<!-- #Include file="../ESSinc/incSvrFuncSims.inc" -->
<%
    Dim RYear
    Dim Emp_no
    Emp_no = Trim(Request("txtEmp_no"))
    RYear = Trim(Request("txtYear"))
%>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID      = "e1302mb1.asp"						           '☆: Biz Logic ASP Name

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================

<!-- #Include file="../ESSinc/lgvariables.inc" --> 

'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim RYear
Dim Emp_no
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029(gCurrency, "Q", "H") %>
End Sub

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   
    if  pOpt = "Q" then
        lgKeyStream = Trim(parent.txtEmp_no.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(parent.txtinternal_cd.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(Replace(frm1.cboYear.Value,"-","")) & gColSep
    else
        lgKeyStream = Trim(frm1.txtEmp_no.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(parent.txtinternal_cd.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(Replace(frm1.cboYear.Value,"-","")) & gColSep
    end if

End Sub        

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
Dim lgYear,i
dim strSQL, IntRetCd
    strSQL = " org_cd = " & FilterVar("1", "''", "S") & " AND pay_gubun = " & FilterVar("Z", "''", "S") & " AND PAY_TYPE = " & FilterVar("*", "''", "S") & " "
    IntRetCD = CommonQueryRs(" year(close_dt) close_year "," hda270t ", strSQL,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    If  IntRetCd = true then
		lgYear = Replace(lgF0, Chr(11), "") 
	end if
	
	For i=lgYear To lgYear-10 step -1
		Call SetCombo(frm1.cboYear, i, i)
	next

	frm1.cboYear.value = CStr(lgYear)
    
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status

    RYear = "<%=RYear%>"
    Emp_no = "<%=Emp_no%>"
    
    parent.document.All("nextprev").style.VISIBILITY = "hidden"

    Call LayerShowHide(0)

    Call SetToolBar("10000")
    Call InitComboBox()

    Call LockField(Document)

    Call DbQuery(1)

End Sub
'========================================================================================
' Function Name : Form_UnLoad
' Function Desc : 페이지 전환이나 화면이 닫힐 경우 실행해야 될 로직 처리 
'========================================================================================
Sub Form_UnLoad()
	On Error Resume Next
 	Set gActiveElement = Nothing
    Set Grid1 = Nothing
End Sub

Function DbQuery(ppage)
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG
    
    Call ClearField(document,2)
    Call LayerShowHide(1)
    Call MakeKeyStream("Q")

    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
    
    DbQuery = True                                                               '☜: Processing is NG
End Function

Function DbQueryOk()
    Err.Clear                                                                    '☜: Clear err status
End Function

Function DbQueryFail()
    Err.Clear
    Call ClearField(Document,2)                                                                    '☜: Clear err status

End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
		
	DbSave = False														         '☜: Processing is NG
		
	Call LayerShowHide(1)

	With Frm1
		.txtMode.value        = "UID_M0002"                                        '☜: Save
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
    DbSave  = True                                                               '☜: Processing is NG
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
    Call DbQuery()
End Function


'========================================================================================================
' Name : FncNext
' Desc : developer describe this line Called by MainNext in Common.vbs
'========================================================================================================
Function FncNext() 
    Dim strVal
    Dim IntRetCD
    Dim strSQL

    FncNext = False                                                              '☜: Processing is OK
    Err.Clear                                                                    '☜: Clear err status

    Call MakeKeyStream("N")
    Call LayerShowHide(1)

    strVal = BIZ_PGM_ID & "?txtMode="          & "UID_M0001"                     '☜: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key
    strVal = strVal     & "&txtPrevNext="      & "N"	                         '☜: Direction

	Call RunMyBizASP(MyBizASP, strVal)										     '☜: Run Biz 

    FncNext = True                                                               '☜: Processing is OK
	
End Function

Sub SubPrint(objFrame)
    Set objActiveEl = document.activeElement
    objFrame.focus()
    objFrame.print()
    objActiveEl.focus
    Set objActiveEl = nothing
End Sub

'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================

Sub Query_OnClick()
    Call DbQuery()
End Sub

Sub Print_onClick()
    Call SubPrint(MyBizASP)
End Sub


Sub GRID_PAGE_OnChange()
End Sub

Sub DELETE_OnClick()
    Call Grid1.DeleteClick()
End Sub

Sub CANCEL_OnClick()
    Call Grid1.CancelClick()
End Sub

</SCRIPT>

<!-- #Include file="../ESSinc/uniSimsClassID.inc" --> 

</HEAD>

<BODY topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
    <TABLE cellSpacing=0 cellPadding=0 border=0>
        <TR>
            <TD valign="top">
                <TABLE width="100%" cellSpacing=0 cellPadding=0 border=0>
                    <TR>
                       <td height="10"></td>
                    </TR>
                    <TR>
                        <td><table width="100%" border="0" cellspacing="1" cellpadding="0" bgcolor="DDDDDD">
                            <tr> 
								<td width="80" height="27" bgcolor="D4E5E8" class="base1">사번</td>
								<td width="85" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtEmp_no" MAXLENGTH=13 SiZE=12 readonly></td>
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">성명</td>
								<td width="80" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtName" MAXLENGTH=20 SiZE=16  readonly></td>
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">직위</td>
								<td width="80" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtroll_pstn" MAXLENGTH=20 SiZE=16  readonly></td>
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">부서</td>
								<td width="153" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtDept_nm" MAXLENGTH=25 SiZE=22  readonly></td>
                            </tr>
                            <tr> 
								<td width="80" height="30" bgcolor="D4E5E8" class=base1 valign=middle>정산년도
								</td>
								<td width="85" bgcolor="FFFFF" align=center>
								    <SELECT Name="cboYear" tabindex=-1 class=base2>
								    </SELECT>
								</td>
								<td bgcolor="FFFFFF" class="base1">&nbsp;</td>
								<td bgcolor="FFFFFF" class="base2">&nbsp;</td>
								<td bgcolor="FFFFFF" class="base1">&nbsp;</td>
								<td bgcolor="FFFFFF" class="base2">&nbsp;</td>
								<td bgcolor="FFFFFF" class="base1">&nbsp;</td>
								<td bgcolor="FFFFFF" class="base2">&nbsp;</td>
                            </tr>
                            </table>
                        </td>
                    </TR>
                    <TR>
                       <td height="10"></td>
                    </TR>

                    <TR>
                        <TD>
                            <TABLE cellSpacing=0 cellPadding=0 width=100% border=0>
								<TR>
								    <TD class="ftgray">&nbsp;
										<img src="../../CShared/ESSimage/icon_07.gif" width="12" height="11"><strong><font color="#014A73">소득사항</font></strong></td>                               
									<TD>
								</TR>
								<tr> 
								    <td height="2"></td>
								</tr>
								<TR><TD>
									<TABLE  border="0" cellSpacing=1 cellPadding=0 width="100%" bgcolor="#DDDDDD">
		                        		<TR>
		                        			<TD CLASS=ctrow03 width=40% valign=middle align=center>구분</TD>
		                        			<TD CLASS=ctrow03 width=15% valign=middle align=center>급여</TD>
		                        			<TD CLASS=ctrow03 width=15% valign=middle align=center>상여</TD>
		                        			<TD CLASS=ctrow03 width=15% valign=middle align=center>인정상여</TD>
		                        			<TD CLASS=ctrow03 width=15% valign=middle align=center>합계</TD>
		                        		</TR>
		                        		<TR>
		                        		    <TD CLASS=ctrow03 width=40%>현근무지근로소득수입금액</TD>
		                        		    <TD CLASS=ctrow04 width=15%><INPUT CLASS=form02 NAME="txtNew_pay_tot_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT></TD>
		                        		    <TD CLASS=ctrow04 width=15%><INPUT CLASS=form02 NAME="txtNew_bonus_tot_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT></TD>
		                        		    <TD CLASS=ctrow04 width=15%><INPUT CLASS=form02 NAME="txtafter_bonus_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT></TD>
		                        		    <TD CLASS=ctrow04 width=15%><INPUT CLASS=form02 NAME="txta_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT></TD>
		                        		</TR>
		                        		<TR>
		                        			<TD CLASS=ctrow03 width=40%>전근무지근로소득수입금액</TD>
		                        			<TD CLASS=ctrow04 width=15%><INPUT CLASS=form02 NAME="txtpay_tot_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT></TD>
		                        			<TD CLASS=ctrow04 width=15%><INPUT CLASS=form02 NAME="txtbonus_tot_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT></TD>
		                        			<TD CLASS=ctrow04 width=15%><INPUT CLASS=form02 NAME="txtold_after_bonus_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT></TD>
		                        			<TD CLASS=ctrow04 width=15%><INPUT CLASS=form02 NAME="txtb_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT></TD>
		                        		</TR>
		                        		<TR>
		                        			<TD CLASS=ctrow03 width=40%>근로소득수입금액</TD>
		                        			<TD CLASS=ctrow04 width=15%></TD>
		                        			<TD CLASS=ctrow04 width=15%></TD>
		                        			<TD CLASS=ctrow04 width=15%></TD>
		                        			<TD CLASS=ctrow04 width=15%><INPUT CLASS=form02 NAME="txtincome_tot_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT></TD>
		                        		</TR>
		                        		<TR>
		                        			<TD CLASS=ctrow03 width=40%>근로소득공제</TD>
		                        			<TD CLASS=ctrow04 width=15%></TD>
		                        			<TD CLASS=ctrow04 width=15%></TD>
		                        			<TD CLASS=ctrow04 width=15%></TD>
		                        			<TD CLASS=ctrow04 width=15%><INPUT CLASS=form02 NAME="txtincome_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT></TD>
		                        		</TR>
		                        		<TR>
		                        			<TD CLASS=ctrow03 width=40%>근로소득금액</TD>
		                        			<TD CLASS=ctrow04 width=15%></TD>
		                        			<TD CLASS=ctrow04 width=15%></TD>
		                        			<TD CLASS=ctrow04 width=15%></TD>
		                        			<TD CLASS=ctrow04 width=15%><INPUT CLASS=form02 NAME="txthfa050t_income_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly readonly></INPUT></TD>
		                        		</TR>
		                        	</TABLE>
                                </TD></TR>
        
								<tr> 
								    <td height="5"></td>
								</tr>
								<TR>
								    <TD class="ftgray">&nbsp;
										<img src="../../CShared/ESSimage/icon_07.gif" width="12" height="11"><strong><font color="#014A73">인적공제</font></strong></td>                               
									<TD>
								</TR>
								<tr> 
								    <td height="2"></td>
								</tr>
								<TR><TD>
									<TABLE  border="0" cellSpacing=1 cellPadding=0 width="100%" bgcolor="#DDDDDD">
		                        		<TR>
		                        		     <TD CLASS=ctrow03 width="34%" colspan="2" align="middle">공제항목</TD>
		                        		     <TD CLASS=ctrow03 width="08%" align="middle">공제사항</TD>
		                        		     <TD CLASS=ctrow03 width="08%" align="middle">정산결과</TD>
		                        		     <TD CLASS=ctrow03 width="34%" colspan="2" align="middle">공제항목</TD>
		                        		     <TD CLASS=ctrow03 width="08%" align="middle">공제사항</TD>
		                        		     <TD CLASS=ctrow03 width="08%" align="middle">정산결과</TD>
		                        		</TR>
		                        		<TR>
		                        		     <TD CLASS=ctrow03 rowspan="4">기본공제</TD>    
		                        		     <TD CLASS=ctrow03 >본인공제</TD>       
		                        		     <TD CLASS="ctrow04" colspan="2">
		                        				<INPUT CLASS="form02" NAME="txtper_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS=ctrow03 rowspan="5">추가공제</TD>       
		                        		     <TD CLASS=ctrow03>장애인수</TD>       
		                        		     <TD CLASS="ctrow04">
		                        				<INPUT CLASS="form02" NAME="txtparia_cnt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="ctrow04">
		                        				<INPUT CLASS="form02" NAME="txtparia_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR>       
		                        		<TR>
		                        		     <TD CLASS=ctrow03>배우자(Y/N)</TD>       
		                        		     <TD CLASS="ctrow04">
		                        				<INPUT CLASS="form02" NAME="txtspouse" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="ctrow04">
		                        				<INPUT CLASS="form02" NAME="txtspouse_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS=ctrow03>경로우대수(65세이상)</TD>      
		                        		     <TD CLASS="ctrow04">
		                        				<INPUT CLASS="form02" NAME="txtold_cnt1" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="ctrow04" rowspan="2">
		                        				<INPUT CLASS="form02" NAME="txtold_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR> 
		                        		<TR>   
		                        		     <TD CLASS=ctrow03 >부양자(여55,남60세이상)</TD>     
		                        		     <TD CLASS="ctrow04" width="11%">
		                        				<INPUT CLASS="form02" NAME="txtsupp_old_cnt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="ctrow04" rowspan="2">
		                        				<INPUT CLASS="form02" NAME="txtsupp_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS=ctrow03>경로우대수(70세이상)</TD>      
		                        		     <TD CLASS="ctrow04">
		                        				<INPUT CLASS="form02" NAME="txtold_cnt2" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR>   
		                        		<TR>   
		                        		     <TD CLASS=ctrow03 >부양자(20세이하/초과장애인)</TD>      
		                        		     <TD CLASS="ctrow04" >
		                        				<INPUT CLASS="form02" NAME="txtsupp_young_cnt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS=ctrow03 >부녀자세대주여부(Y/N)</TD>      
		                        		     <TD CLASS="ctrow04" >
		                        				<INPUT CLASS="form02" NAME="txtlady" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="ctrow04" >
		                        				<INPUT CLASS="form02" NAME="txtlady_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     
		                        		</TR>   
		                        		<TR>   
		                        		     <TD CLASS=ctrow03 colspan="2">소수공제자추가공제</TD>
		                        		     <TD CLASS="ctrow04" colspan="2">
		                        				<INPUT CLASS="form02" NAME="txtsmall_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS=ctrow03 >자녀양육수(6세이하)</TD>
		                        		     <TD CLASS="ctrow04" >
		                        				<INPUT CLASS="form02" NAME="txtchl_rear" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="ctrow04" >
		                        				<INPUT CLASS="form02" NAME="txtchl_rear_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     
		                        		</TR>
		                        		<TR>   
		                        		     <TD CLASS=ctrow03 colspan="2">인적공제계</TD>
		                        		     <TD CLASS="ctrow04" colspan="2">
		                        				<INPUT CLASS="form02" NAME="txtd_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS=ctrow03 colspan="2">&nbsp;</TD>
		                        		     <TD CLASS="ctrow04" colspan="2">&nbsp;</TD>
		                        		     
		                        		</TR>
		                        	</TABLE>
                                </TD></TR>
        
 								<tr> 
								    <td height="5"></td>
								</tr>
								<TR>
								    <TD class="ftgray">&nbsp;
										<img src="../../CShared/ESSimage/icon_07.gif" width="12" height="11"><strong><font color="#014A73">특별세액공제</font></strong></td>                               
									<TD>
								</TR>
								<tr> 
								    <td height="2"></td>
								</tr>
								<TR><TD>
									<TABLE  border="0" cellSpacing=1 cellPadding=0 width="100%" bgcolor="#DDDDDD">
		                        		<TR>
		                        		     <TD CLASS=ctrow03  width="45%" colspan="3" align="middle">공제항목</TD>
		                        		     <TD CLASS=ctrow03  width="08%" align="middle">공제사항</TD>
		                        		     <TD CLASS=ctrow03  width="08%" align="middle">정산결과</TD>
		                        		     <TD CLASS=ctrow03  width="23%" colspan="2" align="middle">공제항목</TD>
		                        		     <TD CLASS=ctrow03  width="08%" align="middle">공제사항</TD>
		                        		     <TD CLASS=ctrow03  width="08%" align="middle">정산결과</TD>
		                        		</TR>
		                        		<TR>
		                        		     <TD CLASS=ctrow03  width="4%" rowspan="10">특별공제</TD>
		                        		     <TD CLASS=ctrow03  width="4%" rowspan="4">보험료</TD>
		                        		     <TD CLASS=ctrow03  width="12%">의료보험료</TD>
		                        		     <TD CLASS="ctrow04">
		                        				<INPUT CLASS="form02" NAME="txtinsur_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="ctrow04">
		                        				<INPUT CLASS="form02" NAME="txtmed_insur_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS=ctrow03 rowspan="3">주택자금</TD>
		                        		     <TD CLASS=ctrow03>주택저축/차입금상환액</TD>
		                        		     <TD CLASS="ctrow04">
		                        				<INPUT CLASS="form02" NAME="txthfa030t_house_fund_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="ctrow04" rowspan="3">
		                        				<INPUT CLASS="form02" NAME="txthfa050t_house_fund_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR>
		                        		<TR>
		                        		     <TD CLASS=ctrow03>고용보험료</TD>
		                        		     <TD CLASS="ctrow04" width="12%">
		                        				<INPUT CLASS="form02" NAME="txthfa030t_emp_insur_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="ctrow04">
		                        				<INPUT CLASS="form02" NAME="txthfa050t_emp_insur_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS=ctrow03>장기주택저당차입금이자상환액(15년미만)</TD>
		                        		     <TD CLASS="ctrow04" width="12%">
		                        				<INPUT CLASS="form02" NAME="txtlong_house_loan_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>		                        		     
		                        		</TR>
		                        		<TR>
		                        		     <TD CLASS=ctrow03>기타보장성보험료</TD>
		                        		     <TD CLASS="ctrow04" width="12%">
		                        				<INPUT CLASS="form02" NAME="txthfa030t_other_insur_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="ctrow04">
		                        				<INPUT CLASS="form02" NAME="txthfa050t_other_insur_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS=ctrow03>장기주택저당차입금이자상환액(15년이상)</TD>
		                        		     <TD CLASS="ctrow04" width="12%">
		                        				<INPUT CLASS="form02" NAME="txtlong_house_loan_amt1" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>	
		                        		</TR>
		                        		<TR>
		                        		     <TD CLASS=ctrow03>장애자전용보험료</TD>
		                        		     <TD CLASS="ctrow04">
		                        				<INPUT CLASS="form02" NAME="txthfa030t_disabled_insur_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="ctrow04">
		                        				<INPUT CLASS="form02" NAME="txthfa050t_disabled_insur_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS=ctrow03 rowspan="7">기부금</TD>
		                        		     <TD CLASS=ctrow03 >법정기부금</TD>
		                        		     <TD CLASS="ctrow04">
		                        				<INPUT CLASS="form02" NAME="txtlegal_contr_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="ctrow04" rowspan="7">
		                        				<INPUT CLASS="form02" NAME="txtcontr_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>			                        		     
		                        		</TR>
		                        		<TR>
		                        		     <TD CLASS=ctrow03  rowspan="2">의료비</TD>
		                        		     <TD CLASS=ctrow03 >일반의료비</TD>
		                        		     <TD CLASS="ctrow04" >
		                        				<INPUT CLASS="form02" NAME="txttot_med_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="ctrow04" rowspan="2">
		                        				<INPUT CLASS="form02" NAME="txtmed_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS=ctrow03 >정치자금기부금</TD>
		                        		     <TD CLASS="ctrow04" >
		                        				<INPUT CLASS="form02" NAME="txtPoli_contr_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR>
		                        		<TR>
		                        		     <TD CLASS=ctrow03 >본인/경로자/장애인의료비</TD>
		                        		     <TD CLASS="ctrow04" >
		                        				<INPUT CLASS="form02" NAME="txtspeci_med_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS=ctrow03 >특례기부금(100%)</TD>
		                        		     <TD CLASS="ctrow04">
		                        				<INPUT CLASS="form02" NAME="txtTaxLaw_contr_amt2" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR>
		                        		<TR>
			                        		     <TD CLASS=ctrow03  rowspan="3">교육비</TD>
		                        		     <TD CLASS=ctrow03  width="10%">본인교육비</TD>
		                        		     <TD CLASS="ctrow04" width="12%">
		                        				<INPUT CLASS="form02" NAME="txtper_edu_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="ctrow04" rowspan="3">
		                        				<INPUT CLASS="form02" NAME="txtedu_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS=ctrow03  >특례기부금(50%)</TD>
		                        		     <TD CLASS="ctrow04"  >		                        		     
		                        				<INPUT CLASS="form02" NAME="txtTaxLaw_contr_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>		                        		     
		                        		     </TD>
		                        		<TR>
		                        		     <TD CLASS=ctrow03  >가족교육비</TD>
		                        		     <TD CLASS="ctrow04" >
		                        				<INPUT CLASS="form02" NAME="txtedu_sum_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS=ctrow03  >우리사주조합기부금</TD>
		                        		     <TD CLASS="ctrow04"  >		                        		     
		                        				<INPUT CLASS="form02" NAME="txtOurstock_contr_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>		                        		     
		                        		     </TD>
		                        		</TR>
		                        		<TR>
		                        		     <TD CLASS=ctrow03  >장애인특수교육비</TD>
		                        		     <TD CLASS="ctrow04" >
		                        				<INPUT CLASS="form02" NAME="txtDisabled_edu_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS=ctrow03  >지정기부금</TD>
		                        		     <TD CLASS="ctrow04" >
		                        				<INPUT CLASS="form02" NAME="txtapp_contr_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR>
		                        		<TR>
		                        		     <TD CLASS=ctrow03  colspan="2">결혼장례비</TD>
		                        		     <TD CLASS="ctrow04" >
		                        				<INPUT CLASS="form02" NAME="hfa030t_Ceremony_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="ctrow04" > 
		                        				<INPUT CLASS="form02" NAME="hfa050t_Ceremony_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>		                        		     
		                        		     </TD>
		                        		     <TD CLASS=ctrow03>노동조합비</TD>
		                        		     <TD CLASS="ctrow04">
		                        				<INPUT CLASS="form02" NAME="txtpriv_contr_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR>
		                        		<TR>
		                        		     <TD CLASS=ctrow03  colspan="3"  >계 또는 표준공제</TD>
		                        		     <TD CLASS="ctrow04" colspan="6">
		                        				<INPUT CLASS="form02" NAME="txtstd_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR>		                        				                        		
		                        		<TR>
		                        		     <TD CLASS=ctrow03 colspan="3">개인연금저축액</TD>
		                        		     <TD CLASS="ctrow04">
		                        				<INPUT CLASS="form02" NAME="txthfa030t_indiv_anu_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="ctrow04">
		                        				<INPUT CLASS="form02" NAME="txthfa050t_indiv_anu_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS=ctrow03  colspan="3">우리사주출연금</TD>
		                        		     <TD CLASS="ctrow04" >
		                        				<INPUT CLASS="form02" NAME="txtOur_stock_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR>
		                        		<TR>
		                        		     <TD CLASS=ctrow03  colspan="3">투자소득공제</TD>
		                        		     <TD CLASS="ctrow04" ></TD>		                        		     
		                        		     <TD CLASS="ctrow04" >
		                        				<INPUT CLASS="form02" NAME="txtinvest_sub_sum_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>

		                        		     <TD CLASS=ctrow03  colspan="3">카드소득공제</TD>
		                        		     <TD CLASS="ctrow04"  >
		                        				<INPUT CLASS="form02" NAME="txtcard_sub_sum_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR>
		                        		<TR>
											 <TD CLASS=ctrow03  colspan="3" >연금보험료공제</TD>
		                        		     <TD CLASS="ctrow04">
		                        				<INPUT CLASS="form02" NAME="txthfa030t_National_pension_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="ctrow04" >
		                        				<INPUT CLASS="form02" NAME="txthfa050t_National_pension_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>		                        		
		                        		     <TD CLASS=ctrow03  colspan="3">외국인교육비/임차료</TD>
		                        		     <TD CLASS="ctrow04" >
		                        				<INPUT CLASS="form02" NAME="txtFore_edu_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR>
                      					<TR>
											<TD CLASS=ctrow03  colspan="3" >퇴직연금소득공제</TD>
		                        		     <TD CLASS="ctrow04">
		                        				<INPUT CLASS="form02" NAME="txthfa030t_Retire_pension" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="ctrow04" >
		                        				<INPUT CLASS="form02" NAME="txthfa050t_Retire_pension" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>			                        		     
		                        		     <TD CLASS=ctrow03 colspan="3">소득공제계</TD>
		                        		     <TD CLASS="ctrow04" >
		                        				<INPUT CLASS="form02" NAME="txtsum_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR>		                        		
		                        		<TR>
		                        		     <TD CLASS=ctrow03 colspan="3">소득과세표준</TD>
		                        		     <TD CLASS="ctrow04" colspan="2">
		                        				<INPUT CLASS="form02" NAME="txttax_std_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS=ctrow03  colspan="3">산출세액</TD>
		                        		     <TD CLASS="ctrow04"  >
		                        				<INPUT CLASS="form02" NAME="txtcalu_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR>
		                        		<TR>
		                        		     <TD CLASS=ctrow03  colspan="2" rowspan="4">세액공제</TD>
		                        		     <TD CLASS=ctrow03  >근로소득</TD>
		                        		     <TD CLASS="ctrow04"  colspan="2">
		                        				<INPUT CLASS="form02" NAME="txtincome_tax_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS=ctrow03 colspan="3">주택자금차입금이자상환액</TD>
		                        		     <TD CLASS="ctrow04" >
		                        				<INPUT CLASS="form02" NAME="txthouse_repay_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR>
		                        		<TR>
		                        		     <TD CLASS=ctrow03  >외국납부세액공제</TD>
		                        		     <TD CLASS="ctrow04"  colspan="2">
		                        				<INPUT CLASS="form02" NAME="txtFore_pay_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>		                        		
		                        		     <TD CLASS=ctrow03  colspan="3">을근납세조합공제</TD>
		                        		     <TD CLASS="ctrow04" >
		                        				<INPUT CLASS="form02" NAME="txtTax_Union_Ded" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR>
		                        		<TR>
		                        		     <TD CLASS=ctrow03  >정치자금기부금세액공제</TD>
		                        		     <TD CLASS="ctrow04"  colspan="2">
		                        				<INPUT CLASS="form02" NAME="txtPolicontr_tax_sub_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>		                        		
		                        		
		                        		     <TD CLASS=ctrow03   colspan="3">세액공제계</TD>
		                        		     <TD CLASS="ctrow04" >
		                        				<INPUT CLASS="form02" NAME="txttax_sub_sum_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR>		                        		
		                        	</TABLE>
                                </TD></TR>

 								<tr> 
								    <td height="5"></td>
								</tr>
								<TR>
								    <TD class="ftgray">&nbsp;
										<img src="../../CShared/ESSimage/icon_07.gif" width="12" height="11"><strong><font color="#014A73">결정세액/차감징수세액</font></strong></td>                               
									<TD>
								</TR>
								<tr> 
								    <td height="2"></td>
								</tr>
								<TR><TD>
									<TABLE  border="0" cellSpacing=1 cellPadding=0 width="100%" bgcolor="#DDDDDD">
		                        		 <TR>
		                        		     <TD CLASS=ctrow03 width=25% valign="middle" align=center>구분</TD>
		                        		     <TD CLASS=ctrow03 width=18% valign="middle" align=center>소득세</TD>
		                        		     <TD CLASS=ctrow03 width=18% valign="middle" align=center>주민세</TD>
		                        		     <TD CLASS=ctrow03 width=18% valign="middle" align=center>농특세</TD>
		                        		     <TD CLASS=ctrow03 width=18% valign="middle" align=center>계</TD>
		                        		</TR>
		                        		<TR>
		                        		     <TD CLASS=ctrow03 >정산세액</TD>
		                        		     <TD CLASS="ctrow04">
		                        				<INPUT CLASS="form02" NAME="txtdec_income_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="ctrow04">
		                        				<INPUT CLASS="form02" NAME="txtdec_res_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="ctrow04">
		                        				<INPUT CLASS="form02" NAME="txtdec_farm_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="ctrow04">
		                        				<INPUT CLASS="form02" NAME="txtdec_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR>
		                        		<TR>
		                        		     <TD CLASS=ctrow03 >현근무지징수세액</TD>  
		                        		     <TD CLASS="ctrow04">
		                        				<INPUT CLASS="form02" NAME="txtnew_income_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="ctrow04">
		                        				<INPUT CLASS="form02" NAME="txtnew_res_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="ctrow04">
		                        				<INPUT CLASS="form02" NAME="txtnew_farm_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="ctrow04">
		                        				<INPUT CLASS="form02" NAME="txtincome_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR>
		                        		<TR>
		                        		     <TD CLASS=ctrow03 >종전근무지세액</TD>
		                        		     <TD CLASS="ctrow04">
		                        				<INPUT CLASS="form02" NAME="txtold_income_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="ctrow04">
		                        				<INPUT CLASS="form02" NAME="txtold_res_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="ctrow04">
		                        				<INPUT CLASS="form02" NAME="txtold_farm_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="ctrow04">
		                        				<INPUT CLASS="form02" NAME="txtold_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR>
		                        		<TR>
		                        		     <TD CLASS=ctrow03  width="30%">징수해야할세액</TD>
		                        		     <TD CLASS="ctrow04" width="15%">
		                        				<INPUT CLASS="form02" NAME="txtincome_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="ctrow04" width="15%">
		                        				<INPUT CLASS="form02" NAME="txtres_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="ctrow04" width="15%">
		                        				<INPUT CLASS="form02" NAME="txtfarm_tax_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		     <TD CLASS="ctrow04" width="15%">
		                        				<INPUT CLASS="form02" NAME="txtf_amt" TYPE="Text" MAXLENGTH=30 SiZE=12 style='TEXT-ALIGN: right;' readonly></INPUT>
		                        		     </TD>
		                        		</TR>

		                        	</TABLE>
                                </TD></TR>
                            </TABLE>
                        </TD>
                    </TR>
                </TABLE>
            </TD>
        </TR>
    </TABLE>

    <TABLE cellSpacing=2 cellPadding=0 width=700 border=0 bgcolor=#ffffff>
        <TR><TD WIDTH="100%" HEIGHT=0><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=0 FRAMEBORDER=0 SCROLLING=yes noresize framespacing=0></IFRAME></TD></TR>
    </TABLE>
</FORM>	

</BODY>
</HTML>
