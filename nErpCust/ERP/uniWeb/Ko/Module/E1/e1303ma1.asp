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
Option Explicit                                                        '☜: indicates that All variables must be declared in advance
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================

Const BIZ_PGM_ID      = "e1303mb1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_ID1     = "e1303ma2.asp"						           '☆: Biz Logic ASP Name
'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" --> 
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029(gCurrency, "I", "H") %>
End Sub

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx
    dim strSQL, IntRetCd

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
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   
    if  pOpt = "Q" then
        lgKeyStream = Trim(parent.txtEmp_no.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(parent.txtinternal_cd.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(frm1.txtYear.Value) & gColSep
        lgKeyStream = lgKeyStream & "Q" & gColSep
    Elseif pOpt = "P" Then
        lgKeyStream = Trim(parent.txtEmp_no.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(parent.txtinternal_cd.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(frm1.txtYear.Value) & gColSep
        lgKeyStream = lgKeyStream & "P" & gColSep
    else
        lgKeyStream = Trim(frm1.txtEmp_no.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(parent.txtinternal_cd.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(frm1.txtYear.Value) & gColSep
    end if
End Sub        

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status

	Call LoadInfTB19029()
    Call InitComboBox()
    parent.document.All("nextprev").style.VISIBILITY = "hidden"

    Call LayerShowHide(0)

    Call SetToolBar("00000")

    Call LockField(Document)

    Call DbQueryEmp(1)

End Sub
'========================================================================================
' Function Name : Form_UnLoad
' Function Desc : 페이지 전환이나 화면이 닫힐 경우 실행해야 될 로직 처리 
'========================================================================================
Sub Form_UnLoad()
	On Error Resume Next
 	'Set gActiveElement = Nothing
    'Set Grid1 = Nothing
End Sub
'========================================================================================
' Function Name : Form_UnLoad
' Function Desc : 페이지 전환이나 화면이 닫힐 경우 실행해야 될 로직 처리 
'========================================================================================
Function dodata(strdo)
	On Error Resume Next
	dim IntRetCD
	
'    IntRetCD = CommonQueryRs(" * "," HFA031T "," EMP_NO = '" & parent.txtEmp_no.Value & "' AND YEAR_YY = '" &  frm1.txtYear.Value &"'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

'	if  IntRetCD = true then
		if strdo="1" then
			IntRetCD = DisplayMsgBox("900018",parent.VB_YES_NO,"X","X")
	
			If IntRetCD = vbNo Then
				Exit Function
			End If
			
			call dbquery("2")
		else 
			call dbquery("1")		
		end if
'	else
'	    Call DisplayMsgBox("800186","X","X","X")		
'	end if



End Function

Function DbQueryEmp(ppage)
    Dim strVal

    Err.Clear                                                                    '☜: Clear err status

    DbQueryEmp = False                                                              '☜: Processing is NG
    'If Grid1.ChkChange() Then Exit Function
    Call ClearField(document,2)
    Call LayerShowHide(1)
    Call MakeKeyStream("Q")

    strVal = BIZ_PGM_ID & "?txtMode="      & "UID_M0001"                   '☜: Query
    strVal = strVal     & "&txtKeyStream=" & lgKeyStream                   '☜: Query Key
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run biz logic

    DbQueryEmp = True                                                               '☜: Processing is NG
End Function


Function DbQueryOk()

    Dim strVal
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
Function DbQuery(ppage)
	Dim strVal
	Dim strDate
    Err.Clear                                                                    '☜: Clear err status
		
	DbQuery = False														         '☜: Processing is NG
		
	if ChkField(Document, "2") then
		exit function
	end if

	if ppage = "2" Then

		Call MakeKeyStream("P")

		strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                     '☜: Query
		strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
		
		Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
	Else
		
		Call MakeKeyStream("Q")
		
		With Frm1

	        if  Num_chk(.PAY_TAX_AMT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .PAY_TAX_AMT.focus()
	            exit function
	        end if

	        if  Num_chk(.BONUS_TAX_AMT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .BONUS_TAX_AMT.focus()
	            exit function
	        end if

	        if  Num_chk(.AFTER_BONUS_AMT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .AFTER_BONUS_AMT.focus()
	            exit function
	        end if

	        if  Num_chk(.BEFORE_INCOME_TAX_AMT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .BEFORE_INCOME_TAX_AMT.focus()
	            exit function
	        end if


	        if  Num_chk(.BEFORE_RES_TAX_AMT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .BEFORE_RES_TAX_AMT.focus()
	            exit function
	        end if
	    
	    
	        if  Num_chk(.OLD_SUPP_CNT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .OLD_SUPP_CNT.focus()
	            exit function
	        end if

	        if  Num_chk(.YOUNG_SUPP_CNT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .YOUNG_SUPP_CNT.focus()
	            exit function
	        end if

	        if  Num_chk(.OLD_CNT1.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .OLD_CNT1.focus()
	            exit function
	        end if

	        if  Num_chk(.OLD_CNT2.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .OLD_CNT2.focus()
	            exit function
	        end if
	        
	        if  Num_chk(.PARIA_CNT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .PARIA_CNT.focus()
	            exit function
	        end if


	        if  Num_chk(.CHL_REAR_CNT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .CHL_REAR_CNT.focus()
	            exit function
	        end if

	        if  Num_chk(.MED_INSUR.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .MED_INSUR.focus()
	            exit function
	        end if

	        if  Num_chk(.EMP_INSUR.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .EMP_INSUR.focus()
	            exit function
	        end if

	        if  Num_chk(.OTHER_INSUR.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .OTHER_INSUR.focus()
	            exit function
	        end if

	        if  Num_chk(.MED_SPPORT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .MED_SPPORT.focus()
	            exit function
	        end if

	        if  Num_chk(.SPECI_MED.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .SPECI_MED.focus()
	            exit function
	        end if

	        if  Num_chk(.PER_EDU.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .PER_EDU.focus()
	            exit function
	        end if

	        if  Num_chk(.FAMILY1_AMT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .FAMILY1_AMT.focus()
	            exit function
	        end if

	        if  Num_chk(.FAMILY2_AMT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .FAMILY2_AMT.focus()
	            exit function
	        end if

	        if  Num_chk(.FAMILY3_AMT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .FAMILY3_AMT.focus()
	            exit function
	        end if

	        if  Num_chk(.FAMILY4_AMT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .FAMILY4_AMT.focus()
	            exit function
	        end if

	        if  Num_chk(.FAMILY1_CNT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .FAMILY1_AMT.focus()
	            exit function
	        end if

	        if  Num_chk(.FAMILY2_CNT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .FAMILY2_AMT.focus()
	            exit function
	        end if

	        if  Num_chk(.FAMILY3_CNT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .FAMILY3_AMT.focus()
	            exit function
	        end if

	        if  Num_chk(.FAMILY4_CNT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .FAMILY4_AMT.focus()
	            exit function
	        end if
	        
	        if  Num_chk(.HOUSE_FUND.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .HOUSE_FUND.focus()
	            exit function
	        end if

	        if  Num_chk(.LONG_HOUSE_LOAN_AMT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .LONG_HOUSE_LOAN_AMT.focus()
	            exit function
	        end if

	        if  Num_chk(.txtLegal_contr_amt.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .txtLegal_contr_amt.focus()
	            exit function
	        end if

	        if  Num_chk(.txtApp_contr_amt.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .txtApp_contr_amt.focus()
	            exit function
	        end if

	        if  Num_chk(.INDIV_ANU2.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .INDIV_ANU2.focus()
	            exit function
	        end if

	        if  Num_chk(.NATIONAL_PENSION_AMT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .NATIONAL_PENSION_AMT.focus()
	            exit function
	        end if

	        if  Num_chk(.INVEST_SUB_AMT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .INVEST_SUB_AMT.focus()
	            exit function
	        end if

	        if  Num_chk(.VENTURE_SUB_AMT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .VENTURE_SUB_AMT.focus()
	            exit function
	        end if

	        if  Num_chk(.CARD_USE_AMT.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .CARD_USE_AMT.focus()
	            exit function
	        end if

	        if  Num_chk(.HOUSE_REPAY.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .HOUSE_REPAY.focus()
	            exit function
	        end if

	        if  Num_chk(.FORE_INCOME.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .FORE_INCOME.focus()
	            exit function
	        end if

	        if  Num_chk(.FORE_PAY.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .FORE_PAY.focus()
	            exit function
	        end if

	        if  Num_chk(.INCOME_REDU.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .INCOME_REDU.focus()
	            exit function
	        end if

	        if  Num_chk(.TAXES_REDU.value, strDate) = True then
	            
	        else
	            Call DisplayMsgBox("800094","X","X","X")
	            .TAXES_REDU.focus()
	            exit function
	        end if

	Call LayerShowHide(1)
		'------ Developer Coding part (End )   -------------------------------------------------------------- 

			.txtMode.value        = "UID_M0002"                                        '☜: Save
	'		.txtFlgMode.value     = lgIntFlgMode
	        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
		End With
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)
    End if 
	
    DbQuery  = True                                                               '☜: Processing is NG
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    Dim strVal

    strVal = BIZ_PGM_ID1 & "?txtEmp_no=" & frm1.txtEmp_no.value
    strVal = strVal & "&txtYear=" & frm1.txtYear.value

    document.location = strVal
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

Sub txtYear_OnChange()
    Call DbQueryEmp(1)
End Sub

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

Sub FncPrintPrev()
	Dim strDate
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

	With Frm1

        if  Date_chk(.year_yy.value & "0101", strDate) = True then
            '.year_yy.value = strDate
        else
            Call DisplayMsgBox("800094","X","X","X")
            .year_yy.focus()
            exit sub
        end if

        if  Date_chk(.Bas_dt.value, strDate) = True then
            .Bas_dt.value = strDate
        else
            Call DisplayMsgBox("800094","X","X","X")
            .Bas_dt.focus()
            exit sub
        end if
    End With

    Call LayerShowHide(1)
    Call MakeKeyStream("Q")

    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

End Sub

</SCRIPT>

<!-- #Include file="../../inc/uniSimsClassID.inc" -->

</HEAD>

<BODY topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
    <TABLE cellSpacing=0 cellPadding=0 border=0 bgcolor=#ffffff width=743>
        <TR height=26 valign=middle>
            <TD class=base1>사번:<INPUT class=base1 NAME="txtEmp_no" MAXLENGTH=13 SiZE=12 tag=14></TD>
            <TD class=base1>성명:<INPUT class=base1 NAME="txtName" MAXLENGTH=20 SiZE=10  tag=14></TD>
            <TD class=base1>직위:<INPUT class=base1 NAME="txtroll_pstn" MAXLENGTH=20 SiZE=10  tag=14></TD>
            <TD class=base1>부서:<INPUT class=base1 NAME="txtDept_nm" MAXLENGTH=25 SiZE=15  tag=14></TD>
        </TR>
        <TR height=24 valign=middle>
	    	<TD class=base1>정산연도:<SELECT NAME="txtYear" ALT="정산연도" STYLE="WIDTH: 100px" TAG="12"></SELECT></TD>
            <TD></TD>
	    	<TD class=base1></TD>
	    	<TD></TD>
        </TR>
        <TR>
            <TD colspan=4>
                <TABLE cellSpacing=1 cellPadding=0 border=0 bgcolor=#ffffff>
                    <TR><TD CLASS=TDFAMILY5 colspan=4>
	                	<FIELDSET><LEGEND ALIGN="LEFT">기본사항</LEGEND>
	                	<TABLE  border="0" cellSpacing=1 cellPadding=0 width="100%">
                            <TR>
	            	        	<TD CLASS=TDFAMILY_TITLE5 >과세대상급여</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="PAY_TAX_AMT" ALT="과세대상급여"  TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22FU" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>
	            	        	<TD CLASS=TDFAMILY_TITLE5 >과세대상상여</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="BONUS_TAX_AMT" ALT="과세대상급여" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22FU" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>
                            </TR>
                            <TR>
	            	        	<TD CLASS=TDFAMILY_TITLE5 >인정상여</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="AFTER_BONUS_AMT" ALT="인정상여" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22FU" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>
	            	        	<TD CLASS=TDFAMILY_TITLE5>국외근로소득</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="FORE_INCOME"  ALT="국외근로소득" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22FU" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>	            	        	
                            </TR>
                            <TR>
	            	        	<TD CLASS=TDFAMILY_TITLE5 >기타소득</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left  >
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="other_income" ALT="기타소득" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22FU" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>
	            	        	<TD CLASS=TDFAMILY_TITLE5 >외국인근로자분리과세적용여부</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left >
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="txtForeign_separate_tax_yn"  TYPE=CHECKBOX STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid; BACKGROUND-COLOR: #E9EDF9" TAG="22" >
	            	        	</TD>	            	        	
                            </TR>                            
                            <TR>
	            	        	<TD CLASS=TDFAMILY_TITLE5 >기납부소득세액</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="BEFORE_INCOME_TAX_AMT" ALT="기납부소득세액" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22FU" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>                            
	            	        	<TD CLASS=TDFAMILY_TITLE5 >기납부주민세액</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="BEFORE_RES_TAX_AMT" ALT="기납부주민세액" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22FU" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>
                            </TR>
	                	</TABLE></FIELDSET>
                    </TD></TR>

                    <TR><TD CLASS=TDFAMILY5 colspan=4>
	                	<FIELDSET><LEGEND ALIGN="LEFT">기본공제</LEGEND>
	                	<TABLE  border="0" cellSpacing=1 cellPadding=0 width="100%">
                            <TR>
	            	        	<TD CLASS=TDFAMILY_TITLE5 >배우자공제</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left colspan=3>
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" TYPE=CHECKBOX STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid; BACKGROUND-COLOR: #E9EDF9" TAG="22" NAME="SPOUSE">
	            	        	</TD>
                            </TR>
                            <TR>
	            	        	<TD CLASS=TDFAMILY_TITLE5 >부양자(소)</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="YOUNG_SUPP_CNT" TYPE="Text" MAXLENGTH=1 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right' ID="Text1">
	            	        	</TD>
	            	        	<TD CLASS=TDFAMILY_TITLE5 >부양자(노)</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="OLD_SUPP_CNT" TYPE="Text" MAXLENGTH=1 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>
                            </TR>
	                	</TABLE></FIELDSET>
                    </TD></TR>

                    <TR><TD CLASS=TDFAMILY5 colspan=4>
	                	<FIELDSET><LEGEND ALIGN="LEFT">추가공제</LEGEND>
	                	<TABLE  border="0" cellSpacing=1 cellPadding=0 width="100%">
                            <TR>
	            	        	<TD CLASS=TDFAMILY_TITLE5>경로우대공제(65세이상)</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="OLD_CNT1" ALT="경로우대공제" TYPE="Text" MAXLENGTH=1 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>
	            	        	<TD CLASS=TDFAMILY_TITLE5>장애자공제</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="PARIA_CNT" ALT="장애자공제" TYPE="Text" MAXLENGTH=1 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>
                            </TR>
                            <TR>
	            	        	<TD CLASS=TDFAMILY_TITLE5>경로우대공제(70세이상)</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="OLD_CNT2" ALT="경로우대공제" TYPE="Text" MAXLENGTH=1 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>
	            	        	<TD CLASS=TDFAMILY_TITLE5>부녀자공제</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" TYPE=CHECKBOX STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid; BACKGROUND-COLOR: #E9EDF9" TAG="22" NAME="LADY">
	            	        	</TD>
                            </TR>
                            <TR>
	            	        	<TD CLASS=TDFAMILY_TITLE5>자녀양육</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="CHL_REAR_CNT" ALT="자녀양육" TYPE="Text" MAXLENGTH=1 SiZE=14 tag="22" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>
	            	        	<TD CLASS=TDFAMILY_TITLE5>&nbsp;</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>&nbsp;</TD>
                            </TR>                            
	                	</TABLE></FIELDSET>
                    </TD></TR>

                    <TR><TD CLASS=TDFAMILY5 colspan=4>
	                	<FIELDSET><LEGEND ALIGN="LEFT">특별공제</LEGEND>
	                	<TABLE  border="0" cellSpacing=1 cellPadding=0 width="100%">
                            <TR>
	            	        	<TD CLASS=TDFAMILY_TITLE5>국민건강보험료</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="MED_INSUR" ALT="국민건강보험료" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22FU" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>
	            	        	<TD CLASS=TDFAMILY_TITLE5>고용보험료</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="EMP_INSUR" ALT="고용보험료" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22FU" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>
                            </TR>
                            <TR>
	            	        	<TD CLASS=TDFAMILY_TITLE5>보장성보험료</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="OTHER_INSUR" ALT="보장성보험료" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22FU" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>
	            	        	<TD CLASS=TDFAMILY_TITLE5>장애자전용보험료</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="Disabled_INSUR" ALT="장애자전용보험료" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22FU" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>
                            </TR>
                            <TR>
	            	        	<TD CLASS=TDFAMILY_TITLE5>일반의료비</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="MED_SPPORT" ALT="일반의료비" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22FU" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>
	            	        	<TD CLASS=TDFAMILY_TITLE5>본인/경로자/장애인의료비</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="SPECI_MED" ALT="장애경로의료비" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22FU" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>
                            </TR>
                            <TR>
	            	        	<TD CLASS=TDFAMILY_TITLE5>본인교육비</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="PER_EDU" ALT="본인교육비" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22FU" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>
	            	        	<TD CLASS=TDFAMILY_TITLE5>유치원교육비</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        		<INPUT CLASS="SINPUTTEST_STYLE" NAME="FAMILY1_CNT" ALT="명" TYPE="Text" MAXLENGTH=3 SiZE=3 tag="22" style='TEXT-ALIGN: right;'></INPUT>명 
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="FAMILY1_AMT" ALT="초중고교육비"  TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22FU" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>
                            </TR>
                            <TR>
	            	        	<TD CLASS=TDFAMILY_TITLE5>초중고교육비</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        		<INPUT CLASS="SINPUTTEST_STYLE" NAME="FAMILY2_CNT" ALT="명" TYPE="Text" MAXLENGTH=3 SiZE=3 tag="22" style='TEXT-ALIGN: right;'></INPUT>명 
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="FAMILY2_AMT" ALT="가족2교육비" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22FU" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>
	            	        	<TD CLASS=TDFAMILY_TITLE5>대학교육비</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        		<INPUT CLASS="SINPUTTEST_STYLE" NAME="FAMILY3_CNT" ALT="명" TYPE="Text" MAXLENGTH=3 SiZE=3 tag="22" style='TEXT-ALIGN: right;'></INPUT>명 
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="FAMILY3_AMT" ALT="가족3교육비" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22FU" STYLE='TEXT-ALIGN: right'></TR>
                            <TR>
	            	        	<TD CLASS=TDFAMILY_TITLE5>장애인특수교육비</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        		<INPUT CLASS="SINPUTTEST_STYLE" NAME="FAMILY4_CNT" ALT="명" TYPE="Text" MAXLENGTH=3 SiZE=3 tag="22" style='TEXT-ALIGN: right;'></INPUT>명 
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME='FAMILY4_AMT' ALT="가족4교육비" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22FU" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>
	            	        	<TD CLASS=TDFAMILY_TITLE5>결혼장례비</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
		                        	<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtCeremony_cnt" ALT="횟수" TYPE="Text" MAXLENGTH=3 SiZE=3 tag="22" style='TEXT-ALIGN: right;'></INPUT>회 
		                        	<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtCeremony_amt" ALT="결혼장례비" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="24FU" style='TEXT-ALIGN: right;'></INPUT>
	            	        	</TD>	            	        	
                            </TR>
		                    <TR>
		                        <TD CLASS="TDFAMILY_TITLE5" >법정기부금</TD>
		                        <TD CLASS="TDFAMILY5">
		                    		<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtLegal_contr_amt" ALT="법정기부금" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="22FU" style='TEXT-ALIGN: right;'></INPUT>
		                    	</TD>
		                        <TD CLASS="TDFAMILY_TITLE5" >정치자금기부금(04/3/11 이전)</TD>
		                        <TD CLASS="TDFAMILY5">
		                    		<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtPoli_contr_amt1" ALT="정치자금기부금1" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="22FU" style='TEXT-ALIGN: right;'></INPUT>
		                    	</TD>
		                    </TR>
		                    <TR>
		                        <TD CLASS="TDFAMILY_TITLE5" >정치자금기부금(04/3/12 이전)</TD>
		                        <TD CLASS="TDFAMILY5">
		                    		<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtPoli_contr_amt2" ALT="정치자금기부금2" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="22FU" style='TEXT-ALIGN: right;'></INPUT>
		                    	</TD>
		                        <TD CLASS="TDFAMILY_TITLE5" >특례기부금</TD>
		                        <TD CLASS="TDFAMILY5">
		                    		<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtTaxLaw_contr_amt" ALT="특례기부금" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="22FU" style='TEXT-ALIGN: right;'></INPUT>
		                    	</TD>
		                    </TR>		                        		
		                    <TR>
		                        <TD CLASS="TDFAMILY_TITLE5" >우리사주조합기부금</TD>
		                        <TD CLASS="TDFAMILY5">
		                    		<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtOurstock_contr_amt" ALT="우리사주조합기부금" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="22FU" style='TEXT-ALIGN: right;'></INPUT>
		                    	</TD>
		                        <TD CLASS="TDFAMILY_TITLE5" >지정기부금</TD>
		                        <TD CLASS="TDFAMILY5">
		                    		<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtApp_contr_amt" ALT="지정기부금" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="22FU" style='TEXT-ALIGN: right;'></INPUT>
		                    	</TD>
		                    </TR>	
		                    <TR>
		                        <TD CLASS="TDFAMILY_TITLE5" >노동조합비</TD>
		                        <TD CLASS="TDFAMILY5">
		                    		<INPUT CLASS="SINPUTTEST_STYLE" NAME="txtPriv_contr_amt" ALT="노동조합비" TYPE="Text" MAXLENGTH=14 SiZE=20 tag="22FU" style='TEXT-ALIGN: right;'></INPUT>
		                    	</TD>
	            	        	<TD CLASS=TDFAMILY_TITLE5>주택저축/차입금상환액</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="HOUSE_FUND" ALT="주택저축/차입금상환액" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22FU" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>
		                    </TR>	
		                    <TR>
	            	        	<TD CLASS=TDFAMILY_TITLE5>장기주택저당차입금이자상환액(15년미만)</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="LONG_HOUSE_LOAN_AMT" ALT="장기주택저당차입금이자상환액" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22FU" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>
	            	        	<TD CLASS=TDFAMILY_TITLE5>장기주택저당차입금이자상환액(15년이상)</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="LONG_HOUSE_LOAN_AMT1" ALT="장기주택저당차입금이자상환액1" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22FU" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>
		                    </TR>		                    
	                	</TABLE></FIELDSET>
                    </TD></TR>

                    <TR><TD CLASS=TDFAMILY5 colspan=4>
	                	<FIELDSET><LEGEND ALIGN="LEFT">기타소득공제</LEGEND>
	                	<TABLE  border="0" cellSpacing=1 cellPadding=0 width="100%">
                            <TR>
	            	        	<TD CLASS=TDFAMILY_TITLE5>개인연금불입액(2000년이전)</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="INDIV_ANU" ALT="개인연금불입액(2000년이전)" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22FU" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>
	            	        	<TD CLASS=TDFAMILY_TITLE5>개인연금불입액(2001년이후)</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="INDIV_ANU2" ALT="개인연금불입액(2001년이후)" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22FU" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>
	            	        	
                            </TR>
                            <TR>
	            	        	<TD CLASS=TDFAMILY_TITLE5>국민연금불입액</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="NATIONAL_PENSION_AMT" ALT="국민연금불입액" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22FU" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>
	            	        	<TD CLASS=TDFAMILY_TITLE5>우리사주출연금</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="txtOur_stock_amt" ALT="우리사주출연금" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22FU" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>
	            	        	
                            </TR>
                            <TR>
	            	        	<TD CLASS=TDFAMILY_TITLE5>투자조합출자액공제율15%</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="txtinvest2_sub_amt"  ALT="투자조합출자액공제율30%" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22FU" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>                            
	            	        	<TD CLASS=TDFAMILY_TITLE5>투자조합출자액공제율20%</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="INVEST_SUB_AMT"  ALT="투자조합출자액공제율20%" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22FU" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>

							</TR>	                            
                            <TR>
	            	        	<TD CLASS=TDFAMILY_TITLE5>투자조합출자액공제율30%</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="VENTURE_SUB_AMT"  ALT="투자조합출자액공제율30%" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22FU" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>
	            	        	<TD CLASS=TDFAMILY_TITLE5>외국인근로자교육비</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="FORE_EDU_AMT" ALT="외국인근로자교육비" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22FU" STYLE='TEXT-ALIGN: right'>	            	        	
	            	        	</TD>
                            </TR>
                            <TR>
	            	        	<TD CLASS=TDFAMILY_TITLE5>신용카드사용액</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="CARD_USE_AMT"  ALT="신용카드사용액" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22FU" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>
	            	        	<TD CLASS=TDFAMILY_TITLE5>직불카드사용액</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left >
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="CARD2_USE_AMT"  ALT="직불카드사용액" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22FU" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>
                            </TR>                            
	                	</TABLE></FIELDSET>
                    </TD></TR>


                    <TR><TD CLASS=TDFAMILY5 colspan=4>
	                	<FIELDSET><LEGEND ALIGN="LEFT">세액공제 및 세액감면</LEGEND>
	                	<TABLE  border="0" cellSpacing=1 cellPadding=0 width="100%">
                             <TR>
	            	        	<TD CLASS=TDFAMILY_TITLE5>미분양주택차입금이자상환액</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left >
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="HOUSE_REPAY"  ALT="미분양주택차입금이자상환액" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22FU" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>                            
	            	        	<TD CLASS=TDFAMILY_TITLE5>외국납부세액</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="FORE_PAY"  ALT="외국납부세액" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22FU" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>
                            </TR>                            
                            <TR>
	            	        	<TD CLASS=TDFAMILY_TITLE5>감면세액 소득세법</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="INCOME_REDU"  ALT="감면세액 소득세법" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22FU" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>
	            	        	<TD CLASS=TDFAMILY_TITLE5>감면세액 조감법</TD>
	            	        	<TD CLASS=TDFAMILY5 align=left>
	            	        	    <INPUT CLASS="SINPUTTEST_STYLE" NAME="TAXES_REDU"  ALT="감면세액 조감법" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="22FU" STYLE='TEXT-ALIGN: right'>
	            	        	</TD>
                            </TR>
	                	</TABLE></FIELDSET>
                    </TD></TR>
					<TR valign=middle height=50>
					    <TD colspan=4 align=center>
	            			<INPUT style="WIDTH: 150px; HEIGHT: 20px" TYPE=button NAME=printprev2 VALUE="기초자료생성" OnClick="vbscript: call dodata('1')">
	            			<INPUT style="WIDTH: 150px; HEIGHT: 20px" TYPE=button NAME=printprev VALUE="계산" OnClick="vbscript: call dodata('2')">
					    </TD>
					</TR>
                </TABLE>
            </TD></TR>
        </TABLE>


    <TABLE cellSpacing=0 cellPadding=0 width=700 border=0 bgcolor=#ffffff>
        <TR><TD WIDTH="100%" HEIGHT=0><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=0 FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME></TD></TR>
    </TABLE>

    <INPUT TYPE=HIDDEN NAME="txtMode">
    <INPUT TYPE=HIDDEN NAME="txtKeyStream">
    <INPUT TYPE=HIDDEN NAME="txtUpdtUserId">
    <INPUT TYPE=HIDDEN NAME="txtInsrtUserId">
    <INPUT TYPE=HIDDEN NAME="txtFlgMode">
    <INPUT TYPE=HIDDEN NAME="txtPrevNext">
    <INPUT TYPE=HIDDEN NAME="txtres_no">
    <INPUT TYPE=HIDDEN NAME="txtdomi">
    <INPUT TYPE=HIDDEN NAME="txtaddr">
    <INPUT TYPE=HIDDEN NAME="txtentr_dt">
    <INPUT TYPE=HIDDEN NAME="txtretire_dt">

</FORM>	
</BODY>
</HTML>

