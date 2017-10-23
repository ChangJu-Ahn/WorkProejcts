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
<!-- #Include file="../../inc/incServer.asp" -->
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
<%

    Dim txtBas_dt
    Dim txtpresent_dt
    Dim Emp_no,txtBackPgmId
    Emp_no = Trim(Request("txtEmp_no"))
    txtBas_dt = Trim(Request("txtBas_dt"))
	txtBackPgmId = Trim(Request("txtBackPgmId"))
%>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance


'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID      = "e1304mb2.asp"						           '☆: Biz Logic ASP Name
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
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================


Dim Bas_dt
Dim Present_dt
Dim Emp_no

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
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
    if  pOpt = "Q" then
        lgKeyStream = Trim(frm1.txtEmp_no.Value) & gColSep       'You Must append one character(gColSep)
        lgKeyStream = lgKeyStream & "" & gColSep        
    else
        lgKeyStream = Trim(frm1.txtEmp_no.Value) & gColSep
        lgKeyStream = lgKeyStream & "" & gColSep
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
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
    'frm1.txtUID.focus() 

    Bas_dt = "<%=txtBas_dt%>"
   ' Present_dt = "<%=txtpresent_dt%>"
    
    Call InitComboBox()
    'Call LockField(Document)	
'    Call ElementVisible(window.parent.document.all("RunQuery"), 0)
    Call LayerShowHide(0)

    Call SetToolBar("00000")

    frm1.txtEmp_no.value = parent.txtEmp_no.Value
  
    if  Bas_dt = "" then
        frm1.txtRetire_dt.value = parent.txtBas_dt.Value
    else
        frm1.txtRetire_dt.value = Bas_dt
    end if

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
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
    
    DbQuery = True                                                               '☜: Processing is NG
End Function

Function DbQueryOk()
    Err.Clear                                                                    '☜: Clear err status
	lgIntFlgMode      = OPMD_UMODE                                              '⊙: Indicates that current mode is Create mode
    'Call Grid1.ShowData(frm1,1)

	frm1.txtRetire_dt.value = Bas_dt
'	frm1.txtRetire_yyyy.value = Bas_dt
   
'    Call ElementVisible(window.parent.document.all("RunQuery"), 1)

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
	
	
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	With Frm1
		.txtMode.value        = "UID_M0002"                                        '☜: Save
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
	End With
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
    DbSave  = True                                                               '☜: Processing is NG
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call DbQuery(1)
End Function


'========================================================================================================
' Name : FncNext
' Desc : developer describe this line Called by MainNext in Common.vbs
'========================================================================================================
Function FncNext() 
    Dim strVal
    Dim IntRetCD

    FncNext = False                                                              '☜: Processing is OK
    Err.Clear                                                                    '☜: Clear err status
    
    if parent.txtDEPT_AUTH.value = "N" then
        Call DisplayMsgBox("800454","X","X","X")
        return
    end if

    Call MakeKeyStream("N")

    Call ClearField(Document,2)                                                                    '☜: Clear err status
    
    'Call SetDefaultVal
    'Call InitVariables														     '⊙: Initializes local global variables

    Call LayerShowHide(1)

    strVal = BIZ_PGM_ID & "?txtMode="          & "UID_M0001"                     '☜: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key
    strVal = strVal     & "&txtPrevNext="      & "N"	                         '☜: Direction
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 

	Call RunMyBizASP(MyBizASP, strVal)										     '☜: Run Biz 

    FncNext = True                                                               '☜: Processing is OK
	
End Function

'========================================================================================================
' Name : FncPrev
' Desc : developer describe this line Called by MainNext in Common.vbs
'========================================================================================================
Function FncPrev() 
    Dim strVal
    Dim IntRetCD

    FncPrev = False                                                              '☜: Processing is OK
    Err.Clear                                                                    '☜: Clear err status

    if parent.txtDEPT_AUTH.value = "N" then
        Call DisplayMsgBox("800454","X","X","X")
        return
    end if

    Call MakeKeyStream("N")

    Call ClearField(Document,2)                                                                    '☜: Clear err status
    
    'Call SetDefaultVal
    'Call InitVariables														     '⊙: Initializes local global variables

    Call LayerShowHide(1)

    strVal = BIZ_PGM_ID & "?txtMode="          & "UID_M0001"                     '☜: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key
    strVal = strVal     & "&txtPrevNext="      & "P"	                         '☜: Direction
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 

	Call RunMyBizASP(MyBizASP, strVal)										     '☜: Run Biz 

    FncPrev = True                                                               '☜: Processing is OK
	
End Function

Sub SubPrint(objFrame)
    Set objActiveEl = document.activeElement
    objFrame.focus()
    objFrame.print()
    objActiveEl.focus
    Set objActiveEl = nothing
End Sub


'========================================================================================================
' Name : goBackForm
' Desc : 
'========================================================================================================
Function goBackForm1() 
    goBackForm1 = False                                                              '☜: Processing is OK
    Err.Clear																		 '☜: Clear err status
	document.location = "<%=txtBackPgmId%>"
    goBackForm1 = True                                                               '☜: Processing is OK
End Function


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

<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->

</HEAD>

<BODY topmargin=0 leftmargin=0 marginheight=0 marginwidth=0 TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
    <TABLE cellSpacing=0 cellPadding=0 width=749 border=0>
        <TR>
            <TD width=13></TD>
            <TD>
                <TABLE cellSpacing=0 cellPadding=0 width=722 border=0 bgcolor=#ffffff>
                    <TR height=26 valign=middle>
                        <TD class=base1>사번:<INPUT class=base1 NAME="txtEmp_no" MAXLENGTH=13 SiZE=12 ></TD>
                        <TD class=base1>성명:<INPUT class=base1 NAME="txtName" MAXLENGTH=20 SiZE=10 ></TD>
                        <TD class=base1>직위:<INPUT class=base1 NAME="txtroll_pstn" MAXLENGTH=20 SiZE=10 ></TD>
                        <TD class=base1>부서:<INPUT class=base1 NAME="txtDept_nm" MAXLENGTH=25 SiZE=15 ></TD>
                    </TR>   
                    <TR>
                        <TD colspan=4>
                            <TABLE cellSpacing=1 cellPadding=0 width=100% border=0 bgcolor=#ffffff>
                                <TR><TD CLASS=TDFAMILY5 colspan=4>
		                        	<FIELDSET><LEGEND ALIGN="LEFT">기준일자</LEGEND>
		                        	<TABLE  border="0" cellSpacing=1 cellPadding=0 width="100%">
                                        <TR>
							                <TD CLASS=TDFAMILY_TITLE5>당사입사일</TD>
							                <TD CLASS=TDFAMILY5><INPUT CLASS="SINPUTTEST_STYLE" NAME="txtEntr_dt" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="24"  style='TEXT-ALIGN: center;'></TD>
	                        	        	<TD CLASS=TDFAMILY_TITLE5>정산시작일</TD>
	                        	        	<TD CLASS=TDFAMILY5><INPUT CLASS="SINPUTTEST_STYLE" NAME="txtRetire_yyyy" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="24"  style='TEXT-ALIGN: center;'></TD>
                                        </TR>
						                <TR>
							                <TD CLASS=TDFAMILY_TITLE5>예상퇴사일</TD>
							                <TD CLASS=TDFAMILY5><INPUT CLASS="SINPUTTEST_STYLE" NAME="txtRetire_dt" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="24" style='TEXT-ALIGN: center;'></TD>
							                <TD CLASS=TDFAMILY_TITLE5></TD>
							                <TD CLASS=TDFAMILY5></TD>
	                   		            </TR>
		                        	</TABLE></FIELDSET>
                                </TD></TR>

                                <TR><TD CLASS=TDFAMILY5 colspan=4>
		                        	<FIELDSET><LEGEND ALIGN="LEFT">퇴직금(평균임금 X 근속개월 / 12)</LEGEND>
		                        	<TABLE  border="0" cellSpacing=1 cellPadding=0 width="100%">
						                <TR>
							                <TD CLASS=TDFAMILY_TITLE5>평균임금</TD>
							                <TD CLASS=TDFAMILY5><INPUT CLASS="SINPUTTEST_STYLE" NAME="txtAvr_wages_amt" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="24" style='TEXT-ALIGN: right;'></TD>
							                <TD CLASS=TDFAMILY_TITLE5>근속개월</TD>
							                <TD CLASS=TDFAMILY5><INPUT CLASS="SINPUTTEST_STYLE" NAME="txtTot_duty_mm" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="24" style='TEXT-ALIGN: right;'></TD>	                   		    </TR>
	                   		            </TR>
						                <TR>
              				                <TD CLASS=TDFAMILY_TITLE5 NOWRAP>퇴직금</TD>
	                   		            	<TD CLASS=TDFAMILY5><INPUT CLASS="SINPUTTEST_STYLE" NAME="txtTot_prov_amt" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="24" style='TEXT-ALIGN: right;'></TD>
							                <TD CLASS=TDFAMILY_TITLE5></TD>
							                <TD CLASS=TDFAMILY5></TD>
	                   		            </TR>
		                        	</TABLE></FIELDSET>
                                </TD></TR>

                                <TR><TD CLASS=TDFAMILY5 colspan=4>
		                        	<FIELDSET><LEGEND ALIGN="LEFT">퇴직정산결과</LEGEND>
		                        	<TABLE  border="0" cellSpacing=1 cellPadding=0 width="100%">
						                <TR>
              				                <TD CLASS=TDFAMILY_TITLE5 >퇴직소득금액</TD>
	                   		            	<TD CLASS=TDFAMILY5 ><INPUT CLASS="SINPUTTEST_STYLE" NAME="txtIncome_amt" TYPE="Text" MAXLENGTH=14 SiZE=14 tag="24" style='TEXT-ALIGN: right;'></TD>
              				                <TD CLASS=TDFAMILY5 ></TD>
	                   		            	<TD CLASS=TDFAMILY5 ></TD>
	                   		            </TR>
		                        	</TABLE></FIELDSET>
                                </TD></TR>
								<TR>
								    <TD colspan=4 align=center>
										<INPUT style="WIDTH: 150px; HEIGHT: 20px" TYPE=button NAME=gobackForm VALUE="돌아가기" OnClick = "vbscript: call goBackForm1()">
								    </TD>
								</TR>
                            </TABLE>
                        </TD>
                    </TR>
                </TABLE>
            </TD>
            <TD width=14></TD>
        </TR>
    </TABLE>
    <TABLE cellSpacing=2 cellPadding=0 width=700 border=0 bgcolor=#ffffff>
        <TR><TD WIDTH="100%" HEIGHT=0><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=0 FRAMEBORDER=0 SCROLLING=yes noresize framespacing=0></IFRAME></TD></TR>
    </TABLE>
</FORM>	

</BODY>
</HTML>
