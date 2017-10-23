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
<!-- #Include file="../../inc/incSvrVarSims.inc"  -->
<!-- #Include file="../../inc/incSvrFuncSims.inc" -->

<Script Language="VBScript">
Option Explicit 
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID      = "e1301mb1.asp"						           '☆: Biz Logic ASP Name

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" --> 

'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim Grid1
Dim Emp_no

'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029(gCurrency, "I", "H") %>
End Sub

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   
    if  pOpt = "Q" then
        lgKeyStream = Trim(parent.txtEmp_no.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(parent.txtinternal_cd.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(Replace(frm1.txtYear.Value,"-","")) & gColSep
    else

        lgKeyStream = Trim(frm1.txtEmp_no.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(parent.txtinternal_cd.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(Replace(frm1.txtYear.Value,"-","")) & gColSep
    end if

End Sub        
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx
    dim strSQL, IntRetCD
    
	iCodeArr = ""
	iNameArr = ""

    strSQL = " org_cd = " & FilterVar("1", "''", "S") & " AND pay_gubun = " & FilterVar("Z", "''", "S") & " AND PAY_TYPE = " & FilterVar("*", "''", "S") & " "
    IntRetCD = CommonQueryRs(" year(close_dt) close_year "," hda270t ", strSQL,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    If  IntRetCd = true then
		iDx = Replace(lgF0, Chr(11), "") +1
	end if
	iCodeArr = cdbl(idx) & Chr(11) & iCodeArr
	iNameArr = cdbl(idx) & Chr(11) & iNameArr
	   
    Call SetCombo2(frm1.txtYear, iCodeArr, iNameArr, Chr(11))

End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status

	parent.document.All("nextprev").style.VISIBILITY = "hidden"

	call LoadInfTB19029()
    Call InitComboBox()

    Call LayerShowHide(0)

    Call SetToolBar("10010")

    frm1.txtEmp_no.value = parent.txtEmp_no.Value
    
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
    
    If frm1.txtYear.value = "" then
		Call DisplayMsgBox("800094","X","X","X")
		Exit Function
    End if
    
    If len(frm1.txtYear.value)<>4 then
		Call DisplayMsgBox("800094","X","X","X")
		Exit Function
    End if
    
    Call ClearField(document,2)
    Call LayerShowHide(1)
    if ppage = 1 then
        Call MakeKeyStream("Q")
    else
        Call MakeKeyStream("S")        
    end if

    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
    
    DbQuery = True                                                               '☜: Processing is NG
End Function

Function DbQueryOk()
    Err.Clear                                                                    '☜: Clear err status
	lgIntFlgMode      = OPMD_UMODE                                              '⊙: Indicates that current mode is Create mode
    'Call Grid1.ShowData(frm1,1)
'    Call ElementVisible(window.parent.document.all("RunQuery"), 1)
    Call SetToolBar("10010")

End Function

Function DbQueryFail()
    Err.Clear
	lgIntFlgMode = ""

    Call ClearField(Document,2)
    Call SetToolBar("10000")

End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
		
	DbSave = False														         '☜: Processing is NG
   
    If frm1.txtYear.value = "" then
        Call DisplayMsgBox("800094","X","X","X")
		Exit Function
    End if
    
    If len(frm1.txtYear.value)<>4 then
        Call DisplayMsgBox("800094","X","X","X")
		Exit Function
    End if
   
	if ChkField(Document, "2") then
		exit function
	end if
	
	Call LayerShowHide(1)
	Call MakeKeyStream("S")
	
	With Frm1
		.txtMode.value        = "UID_M0002"                                        '☜: Save
		.txtFlgMode.value     = lgIntFlgMode
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
    Call DbQuery(2)
End Function

Function DbSaveFail()
	Call DisplayMsgBox("990024","X","X","X")
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
    Call DbQuery(1)
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

<!-- #Include file="../../inc/uniSimsClassID.inc" -->

</HEAD>

<BODY topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
    <TABLE cellSpacing=0 cellPadding=0 width=749 border=0>
        <TR>
            <TD width=13></TD>
            <TD>
                <TABLE cellSpacing=0 cellPadding=0 width=722 border=0 bgcolor=#ffffff>
                    <TR height=26 valign=middle>
                        <TD class=base1>사번:<INPUT class=base1 NAME="txtEmp_no" MAXLENGTH=13 SiZE=12 tag=14></TD>
                        <TD class=base1>성명:<INPUT class=base1 NAME="txtName" MAXLENGTH=20 SiZE=10  tag=14></TD>
                        <TD class=base1>직위:<INPUT class=base1 NAME="txtroll_pstn" MAXLENGTH=20 SiZE=10  tag=14></TD>
                        <TD class=base1>부서:<INPUT class=base1 NAME="txtDept_nm" MAXLENGTH=25 SiZE=15  tag=14></TD>
                    </TR>   
                    <TR height=26 valign=middle>
                        <TD class=base1>정산연도:<SELECT NAME="txtYear" ALT="정산연도" STYLE="WIDTH: 100px" TAG="12"></SELECT></TD>
                        <TD></TD>
		            	<TD class=base1></TD>
		            	<TD></TD>
                    </TR>

                    <TR>
                        <TD colspan=4>
                            <TABLE cellSpacing=1 cellPadding=0 width=100% border=0 bgcolor=#ffffff>
                                <TR><TD>
		                        	<FIELDSET><LEGEND ALIGN="LEFT">소득공제</LEGEND>
		                        	<TABLE  border="0" cellSpacing=1 cellPadding=0 width="100%">
		                        		<TR>
                                    	    <TD CLASS="TDFAMILY_TITLE" >배우자</TD>
	                                       	<TD CLASS="TDFAMILY">
		                        		        <INPUT CLASS="SINPUTTEST_STYLE" TYPE="CHECKBOX" NAME="rdoSpouse_t" ID="rdoPhantomType1" tag="24" STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid; BACKGROUND-COLOR: #E9EDF9"</INPUT>
		                        		    </TD>
                                    	    <TD CLASS="TDFAMILY_TITLE" >부녀자</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        		        <INPUT CLASS="SINPUTTEST_STYLE" TYPE="CHECKBOX" NAME="rdoLady_t" ID="rdoPhantomType2" tag="24" STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid; BACKGROUND-COLOR: #E9EDF9"></INPUT>
		                        		    </TD>
		                        		</TR>
		                        		<TR>
		                        		    <TD CLASS="TDFAMILY_TITLE" >부양자(노)</TD>
                                        	<TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtSupp_old_cnt_t" ALT="부양자(노)" TYPE="Text" MAXLENGTH=5 SiZE=10 tag="24" style='TEXT-ALIGN: right;'></INPUT>
                                        	</TD>
		                        		    <TD CLASS="TDFAMILY_TITLE" >부양자(소)</TD>
                                        	<TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtSupp_young_cnt_t" ALT="부양자(소)" TYPE="Text" MAXLENGTH=5 SiZE=10 tag="24" style='TEXT-ALIGN: right;'></INPUT>
                                        	</TD>
		                        		</TR>
		                        		<TR>
		                        		    <TD CLASS="TDFAMILY_TITLE" >경로자(65이상)</TD>
                                        	<TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtOld_cnt_t1" ALT="경로자" TYPE="Text" MAXLENGTH=5 SiZE=10 tag="24" style='TEXT-ALIGN: right;'></INPUT>
                                        	</TD>
		                        		    <TD CLASS="TDFAMILY_TITLE" >경로자(70이상)</TD>
                                        	<TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtOld_cnt_t2" ALT="경로자" TYPE="Text" MAXLENGTH=5 SiZE=10 tag="24" style='TEXT-ALIGN: right;'></INPUT>
                                        	</TD>
		                        		</TR>
		                        		<TR>
		                        		    <TD CLASS="TDFAMILY_TITLE" >자녀양육수</TD>
                                        	<TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtChl_rear_inwon_t" ALT="자녀양육수" TYPE="Text" MAXLENGTH=5 SiZE=10 tag="24" style='TEXT-ALIGN: right;'></INPUT>
                                        	</TD>
		                        		    <TD CLASS="TDFAMILY_TITLE" >장애인</TD>
                                        	<TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtParia_cnt_t" ALT="장애인" TYPE="Text" MAXLENGTH=5 SiZE=10 tag="24" style='TEXT-ALIGN: right;'></INPUT>
                                        	</TD>
		                        		</TR>
		                        	</TABLE>
		                        	</FIELDSET>
                                </TD></TR>
        
                                <TR><TD>
		                        	<FIELDSET>
		                        	<TABLE  border="0"  cellSpacing=1 cellPadding=0 width="100%">
		                        		<TR>
		                        		    <TD CLASS="TDFAMILY_TITLE" >기 타 보 험</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtOther_insur_amt" ALT="기타보험" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="22FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		    <TD CLASS="TDFAMILY_TITLE" >건 강 보 험</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtMed_insur_amt" ALT="건강보험" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="24FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		</TR>
		                        		<TR>
		                        		    <TD CLASS="TDFAMILY_TITLE" >고 용 보 험</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtEmp_insur_amt" ALT="고용보험" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="24FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		    <TD CLASS="TDFAMILY_TITLE" >국 민 연 금</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtNational_pension_amt" ALT="국민연금" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="24FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		</TR>
		                        		<TR>
		                        		    <TD CLASS="TDFAMILY_TITLE" >장애인전용보험</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtDisabled_insur_amt" ALT="장애인전용보험" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="22FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		    <TD CLASS="TDFAMILY_TITLE" >&nbsp;</TD>
		                        		    <TD CLASS="TDFAMILY">&nbsp;</TD>
		                        		</TR>
		                        	</TABLE>
		                        	</FIELDSET>
                                </TD></TR>
        
                                <TR><TD>
		                        	<FIELDSET>
		                        	<TABLE  border="0" cellSpacing=1 cellPadding=0 width="100%">
		                        		<TR>
		                        		    <TD CLASS="TDFAMILY_TITLE" >본인교육비</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtPer_edu_amt" ALT="본인교육비" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="24FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		    <TD CLASS="TDFAMILY_TITLE" >장애인특수교육비</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtDisabled_edu_amt" ALT="장애인특수교육비" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="24FU" style='TEXT-ALIGN: right;'></INPUT>
		                        		    </TD>
		                        		</TR>
		                        		<TR>
		                        		    <TD CLASS="TDFAMILY_TITLE" >초중고교육비</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtFam_edu_amt" ALT="초중고교육비" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="24FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		    <TD CLASS="TDFAMILY_TITLE" >자녀수</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtFam_edu_cnt" ALT="자녀수" TYPE="Text" MAXLENGTH=5 SiZE=10 tag="24" style='TEXT-ALIGN: right;'></INPUT> 명
		                        			</TD>
		                        		</TR>
		                        		<TR>
		                        		    <TD CLASS="TDFAMILY_TITLE" >유치원교육비</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtKind_edu_amt" ALT="유치원교육비" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="24FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		    <TD CLASS="TDFAMILY_TITLE" >자녀수</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtKind_edu_cnt" ALT="자녀수" TYPE="Text" MAXLENGTH=5 SiZE=10 tag="24" style='TEXT-ALIGN: right;'></INPUT> 명
		                        			</TD>
		                        		</TR>
		                        		<TR>
		                        		    <TD CLASS="TDFAMILY_TITLE" >대학교육비</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtUniv_edu_amt" ALT="대학교육비" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="24FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		    <TD CLASS="TDFAMILY_TITLE" >자녀수</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtUniv_edu_cnt" ALT="자녀수" TYPE="Text" MAXLENGTH=5 SiZE=10 tag="24" style='TEXT-ALIGN: right;'></INPUT> 명
		                        			</TD>
		                        		</TR>
		                        	</TABLE>
		                        	</FIELDSET>
                                </TD></TR>

                                <TR><TD>
		                        	<FIELDSET>
		                        	<TABLE  border="0" cellSpacing=1 cellPadding=0 width="100%">
		                        		<TR>
		                        		    <TD CLASS="TDFAMILY_TITLE" >일반의료비</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtTot_med_amt" ALT="일반의료비" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="24FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		    <TD CLASS="TDFAMILY_TITLE" >본인/경로자/장애인의료비</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtSpeci_med_amt" ALT="경로장애의료비" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="24FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		</TR>
		                        	</TABLE>
		                        	</FIELDSET>
                                </TD></TR>
                                 <TR><TD>
		                        	<FIELDSET>
		                        	<TABLE  border="0" cellSpacing=1 cellPadding=0 width="100%">
		                        		<TR>
		                        		    <TD CLASS="TDFAMILY_TITLE" >법정기부금</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtLegal_contr_amt" ALT="법정기부금" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="24FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		    <TD CLASS="TDFAMILY_TITLE" >정치자금기부금(04/3/11 이전)</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtPoli_contr_amt1" ALT="정치자금기부금1" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="24FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		</TR>
		                        		<TR>
		                        		    <TD CLASS="TDFAMILY_TITLE" >정치자금기부금(04/3/12 이후)</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtPoli_contr_amt2" ALT="정치자금기부금2" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="24FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		    <TD CLASS="TDFAMILY_TITLE" >특례기부금</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtTaxLaw_contr_amt" ALT="특례기부금" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="24FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		</TR>		                        		
		                        		<TR>
		                        		    <TD CLASS="TDFAMILY_TITLE" >우리사주조합기부금</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtOurstock_contr_amt" ALT="우리사주조합기부금" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="24FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		    <TD CLASS="TDFAMILY_TITLE" >지정기부금</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtApp_contr_amt" ALT="지정기부금" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="24FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		</TR>	
		                        		<TR>
		                        		    <TD CLASS="TDFAMILY_TITLE" >노동조합비</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtPriv_contr_amt" ALT="노동조합비" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="24FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		    <TD CLASS="TDFAMILY_TITLE" >&nbsp;</TD>
		                        		    <TD CLASS="TDFAMILY">&nbsp;</TD>
		                        		</TR>			                        		
		                        	</TABLE>
		                        	</FIELDSET>
                                </TD></TR>    
                                <TR><TD>
		                        	<FIELDSET>
		                        	<TABLE  border="0" cellSpacing=1 cellPadding=0 width="100%">
		                        		<TR>
		                        		    <TD CLASS="TDFAMILY_TITLE" >주택저축/차입금상환액</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtHouse_fund_amt" ALT="주택저축/차입금상환액" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="22FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		    <TD CLASS="TDFAMILY_TITLE" >장기주택저당차입금이자상환액(15년미만)</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtLong_house_loan_amt" ALT="장기주택저당차입금이자상환액1" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="22FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		</TR>
		                        		<TR>
		                        		    <TD CLASS="TDFAMILY_TITLE" >장기주택저당차입금이자상환액(15년이상)</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtLong_house_loan_amt1" ALT="장기주택저당차입금이자상환액2" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="22FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		    <TD CLASS="TDFAMILY_TITLE" >결혼/장례/이사비</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        		        <INPUT CLASS="SINPUTTEST_STYLE" NAME="txtCeremony_cnt" ALT="횟수" TYPE="Text" MAXLENGTH=3 SiZE=3 tag="22" style='TEXT-ALIGN: right;'></INPUT>회 
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtCeremony_amt" ALT="결혼장례비" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="24FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		</TR>		                        		
		                        		<TR>
		                        		    <TD CLASS="TDFAMILY_TITLE" >외국인교육비/임차료</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtFore_edu_amt" ALT="외국인교육비/임차료" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="22FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		    <TD CLASS="TDFAMILY_TITLE" >우리사주출연금</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtOur_stock_amt" ALT="우리사주" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="22FU" style='TEXT-ALIGN: right;'></INPUT>
		                        		    </TD>
		                        		</TR>
		                        		<TR>
		                        		    <TD CLASS="TDFAMILY_TITLE" >개인연금(2000년이전)</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtIndiv_anu_amt" ALT="개인연금(2000년이전)" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="22FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		    <TD CLASS="TDFAMILY_TITLE" >연금저축(2001년이후)</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtIndiv_anu2_amt" ALT="연금저축(2001년이후)" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="22FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		</TR>
		                        		<TR>
		                        		    <TD CLASS="TDFAMILY_TITLE" >1999.8.31이전출자금</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtInvest_sub_amt" ALT="1999.8.31이전출자금" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="22FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		    <TD CLASS="TDFAMILY_TITLE" >1999.8.31이후출자금</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtVenture_sub_amt" ALT="1999.8.31이후출자금" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="22FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		</TR>
		                        		<TR>
		                        		    <TD CLASS="TDFAMILY_TITLE" >2001.12.31이후출자금</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtinvest2_sub_amt" ALT="2001.12.31이후출자금" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="22FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		    <TD CLASS="TDFAMILY_TITLE" ></TD>
		                        		    <TD CLASS="TDFAMILY">
		                        			</TD>
		                        		</TR>		                        		
		                        	</TABLE>
		                        	</FIELDSET>
                                </TD></TR>
        
                                <TR><TD>
		                        	<FIELDSET>
		                        	<TABLE  border="0" cellSpacing=1 cellPadding=0 width="100%">
		                        		<TR>
		                        			<TD CLASS="TDFAMILY_TITLE" >신용카드사용금액</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtCard_use_amt" ALT="카드사용금액" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="22FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		    <TD CLASS="TDFAMILY_TITLE" >직불카드사용금액</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtCard2_use_amt" ALT="기타소득" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="22FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		</TR>
		                        		<TR>
		                        		    <TD CLASS="TDFAMILY_TITLE" >외국소득</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtFore_income_amt" ALT="외국소득" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="22FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		    <TD CLASS="TDFAMILY_TITLE" >인정상여</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtAfter_bonus_amt" ALT="인정상여" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="22FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		</TR>
		                        		<TR>
		                        		    <TD CLASS="TDFAMILY_TITLE" >기타소득</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtOther_income_amt" ALT="기타소득" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="22FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		    <TD CLASS="TDFAMILY_TITLE" ></TD>
		                        		    <TD CLASS="TDFAMILY">
		                        			</TD>
		                        		</TR>		                        		
		                        	</TABLE>
		                        	</FIELDSET>
                                </TD></TR>
 
                                <TR><TD>
		                        	<FIELDSET>
		                        	<TABLE  border="0" cellSpacing=1 cellPadding=0 width="100%">
		                        		<TR>
		                        		    <TD CLASS="TDFAMILY_TITLE" >주택차입금이자상환액</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtHouse_repay_amt" ALT="주택차입금이자상환액" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="22FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		    <TD CLASS="TDFAMILY_TITLE" >&nbsp;</TD>
		                        		    <TD CLASS="TDFAMILY">&nbsp;</TD>
		                        		</TR>
		                        		<TR>
		                        		    <TD CLASS="TDFAMILY_TITLE" >외국납부세액</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtFore_pay_amt" ALT="외국납부세액" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="22FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		    <TD CLASS="TDFAMILY_TITLE" >현근무지재형기금</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtSave_tax_sub_amt" ALT="현근무지재형기금" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="22FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		</TR>
		                        		<TR>
		                        		    <TD CLASS="TDFAMILY_TITLE" >소득세법</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtIncome_redu_amt" ALT="소득세법" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="22FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		    <TD CLASS="TDFAMILY_TITLE" >조감법</TD>
		                        		    <TD CLASS="TDFAMILY">
		                        				<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtTaxes_redu_amt" ALT="조감법" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="22FU" style='TEXT-ALIGN: right;'></INPUT>
		                        			</TD>
		                        		</TR>
		                        	</TABLE>
		                        	</FIELDSET>
                                </TD></TR>
                            </TABLE>
                        </TD>
                    </TR>
                </TABLE>
            </TD>
            <TD width=14></TD>
        </TR>
    </TABLE>

    <TABLE cellSpacing=2 cellPadding=0 width=700 border=0 bgcolor=#ffffff>
        <TR><TD WIDTH="100%" HEIGHT=0><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=0 FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME></TD></TR>
    </TABLE>
    <INPUT TYPE=HIDDEN NAME="txtMode">
    <INPUT TYPE=HIDDEN NAME="txtKeyStream">
    <INPUT TYPE=HIDDEN NAME="txtUpdtUserId">
    <INPUT TYPE=HIDDEN NAME="txtInsrtUserId">
    <INPUT TYPE=HIDDEN NAME="txtFlgMode">
    <INPUT TYPE=HIDDEN NAME="txtPrevNext">
</FORM>	

</BODY>
</HTML>
