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

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance


'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID      = "e1201mb1.asp"						           '☆: Biz Logic ASP Name
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" --> 

'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================


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
    if  pOpt = "Q" then
'        if  Trim(parent.txtEmp_no2.Value) = "" then
            lgKeyStream = Trim(parent.txtEmp_no.Value) & gColSep
'        else
'            lgKeyStream = Trim(parent.txtEmp_no2.Value) & gColSep
'        end if
        lgKeyStream = lgKeyStream & Trim(parent.txtinternal_cd.Value) & gColSep
        lgKeyStream = lgKeyStream & "" & gColSep
    else
        lgKeyStream = Trim(frm1.txtEmp_no.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(parent.txtinternal_cd.Value) & gColSep
    end if
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
End Sub
'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitGrid()
    Set Grid1 = New Grid
    Grid1.MaxCols = 4+1
    Grid1.SheetMaxrows = 3
    Set Grid1.Source = document.frm1
End Sub

'========================================================================================================
' Name : LoadInfTB19029
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Private Sub LoadInfTB19029()

<!--#Include file="../../ComAsp/LoadInfTB19029.asp"-->

<%Call loadInfTB19029(gCurrency,"Q","H")%>

End Sub
'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear           
'    if  parent.txtDEPT_AUTH.value = "Y" then
'        parent.document.All("nextprev").style.VISIBILITY = "visible"
'        Call SetToolBar("10000")    
'    else
        parent.document.All("nextprev").style.VISIBILITY = "hidden"
        Call SetToolBar("00000")    
'    end if

    Call LayerShowHide(0)
    Call LockField(Document)
	Call LoadInfTB19029()
    Call DbQuery(1)
End Sub
'========================================================================================
' Function Name : Form_UnLoad
' Function Desc : 페이지 전환이나 화면이 닫힐 경우 실행해야 될 로직 처리 
'========================================================================================
Sub Form_UnLoad()
	On Error Resume Next
    Set Grid1 = Nothing
End Sub

Function DbQuery(ppage)

    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG
    'If Grid1.ChkChange() Then Exit Function
    'Call ClearField(document,2)
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
'		.txtFlgMode.value     = lgIntFlgMode
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

    'Call InitVariables

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call DbQuery(1)
End Function


Sub SubPrint(objFrame)
    Set objActiveEl = document.activeElement
    objFrame.focus()
    objFrame.print()
    objActiveEl.focus
    Set objActiveEl = nothing
End Sub

Function GetRow(pRow)
	GetRow=False
    Grid1.ActiveRow = pRow
    If Mid(document.activeElement.getAttribute("tag"),3,1) = "1" Then
	    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	    	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End If
	GetRow=True
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

'==========================================================================================
'   Event Name : Radio OnClick()
'   Event Desc : Radio Button Click시 lgBlnFlgChgValue 처리 / Value
'==========================================================================================
Sub rdoUnionFlag1_OnClick()
	lgBlnFlgChgValue = True	
End Sub

Sub rdoUnionFlag2_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub rdoPressFlag1_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub rdoPressFlag2_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub rdoOverseaFlag1_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub rdoOverseaFlag2_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub rdoResFlag1_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub rdoResFlag2_OnClick()
	lgBlnFlgChgValue = True
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
<FORM NAME=frm1 target=MyBizASP METHOD="POST">
    <TABLE cellSpacing=0 cellPadding=0 width=770 border=0>
        <TR>
            <TD width=13></TD>
            <TD>
                <TABLE cellSpacing=0 cellPadding=0 width=743 border=0 bgcolor=#ffffff>
                    <TR height=36 valign=center>
                        <TD class=base1>사번:<INPUT class=base1 NAME="txtEmp_no" MAXLENGTH=13 SiZE=12 tag=14></TD>
                        <TD class=base1>성명:<INPUT class=base1 NAME="txtName" MAXLENGTH=20 SiZE=10  tag=14></TD>
                        <TD class=base1>직위:<INPUT class=base1 NAME="txtroll_pstn" MAXLENGTH=20 SiZE=10  tag=14></TD>
                        <TD class=base1>부서:<INPUT class=base1 NAME="txtDept_nm" MAXLENGTH=25 SiZE=15  tag=14></TD>
                    </TR>
                    <TR>
                        <TD colspan=4>
                            <TABLE cellSpacing=1 cellPadding=0 width=100% border=0 bgcolor=#ffffff>
                                <TR>
				                	<TD CLASS="TDFAMILY_TITLE">급여구분</TD>
				                	<TD CLASS="TDFAMILY"><INPUT CLASS="SINPUTTEST_STYLE" TYPE="Text" size=15 name=txtpay_cd tag="24"></TD>
				                	<TD CLASS="TDFAMILY_TITLE">연봉(연봉직)</TD>
				                	<TD CLASS="TDFAMILY"><INPUT CLASS="SINPUTTEST_STYLE" TYPE="Text" size=15 name=txtAnnualSal tag="24" style="TEXT-ALIGN: right"></TD>
				                </TR>
				                <TR>
				                	<TD CLASS="TDFAMILY_TITLE">기본급(연봉)</TD>
				                	<TD CLASS="TDFAMILY"><INPUT CLASS="SINPUTTEST_STYLE" TYPE="Text" size=15 name=txtsalary tag="24" style="TEXT-ALIGN: right"></TD>
				                	<TD CLASS="TDFAMILY_TITLE">상여기준금(연봉)</TD>
				                	<TD CLASS="TDFAMILY"><INPUT CLASS="SINPUTTEST_STYLE" TYPE="Text" size=15 name=txtBonusSalary tag="24" style="TEXT-ALIGN: right"></TD>
				                </TR>
				                <TR>
				                	<TD CLASS="TDFAMILY_TITLE">연장비과세적용구분</TD>
				                	<TD CLASS="TDFAMILY"><INPUT CLASS="SINPUTTEST_STYLE" TYPE="Text" ALT="세액구분" size=15 name=txttax_cd tag="24"></TD>
				                	<TD CLASS="TDFAMILY_TITLE">은행/계좌번호</TD>
				                	<TD CLASS="TDFAMILY"><INPUT CLASS="SINPUTTEST_STYLE" TYPE="Text" size=10 NAME="txtBankNm" tag="24">
				                							   <INPUT CLASS="SINPUTTEST_STYLE" TYPE="Text" size=20 NAME="txtAccntNo" tag="24"></TD>
				                </TR>
				                <TR>
				                	<TD CLASS="TDFAMILY_TITLE">거주구분</TD>
				                	<TD CLASS="TDFAMILY">
				                		<INPUT TYPE="RADIO" CLASS="SINPUTTEST_STYLE" NAME="rdoResFlag" TAG="24" VALUE="Y" ID="rdoResFlag1" STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid; BACKGROUND-COLOR: #E9EDF9"><LABEL FOR="rdoResFlag1">거주자</LABEL>&nbsp;&nbsp;&nbsp;
								        <INPUT TYPE="RADIO" CLASS="SINPUTTEST_STYLE" NAME="rdoResFlag" TAG="24" VALUE="N" ID="rdoResFlag2" STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid; BACKGROUND-COLOR: #E9EDF9"><LABEL FOR="rdoResFlag2">비거주자</LABEL>			
                                    </TD>
				                	<TD CLASS="TDFAMILY_TITLE">기자구분</TD>
				                	<TD CLASS="TDFAMILY">
				                		<INPUT TYPE="RADIO" CLASS="SINPUTTEST_STYLE" NAME="rdoPressFlag" TAG="24" VALUE="Y" ID="rdoPressFlag1" STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid; BACKGROUND-COLOR: #E9EDF9"><LABEL FOR="rdoPressFlag1">기자</LABEL>&nbsp;&nbsp;&nbsp;
								        <INPUT TYPE="RADIO" CLASS="SINPUTTEST_STYLE" NAME="rdoPressFlag" TAG="24" VALUE="N" ID="rdoPressFlag2" STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid; BACKGROUND-COLOR: #E9EDF9"><LABEL FOR="rdoPressFlag2">비기자</LABEL>							                	
                                    </TD>
				                </TR>
				                <TR>
				                	<TD CLASS="TDFAMILY_TITLE">국외근로자구분</TD>
				                	<TD CLASS="TDFAMILY">
				                		<INPUT TYPE="RADIO" CLASS="SINPUTTEST_STYLE" NAME="rdoOverseaFlag" TAG="24" VALUE="Y" ID="rdoOverseaFlag1" STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid; BACKGROUND-COLOR: #E9EDF9"><LABEL FOR="rdoOverseaFlag1">국외근로자</LABEL>&nbsp;&nbsp;&nbsp;
								        <INPUT TYPE="RADIO" CLASS="SINPUTTEST_STYLE" NAME="rdoOverseaFlag" TAG="24" VALUE="N" ID="rdoOverseaFlag2" STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid; BACKGROUND-COLOR: #E9EDF9"><LABEL FOR="rdoOverseaFlag2">국내근로자</LABEL>
                                    </TD>
				                	<TD CLASS="TDFAMILY_TITLE">노조구분</TD>
				                	<TD CLASS="TDFAMILY">
				                		<INPUT TYPE="RADIO" CLASS="SINPUTTEST_STYLE" NAME="rdoUnionFlag" TAG="24" VALUE="Y" ID="rdoUnionFlag1" STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid; BACKGROUND-COLOR: #E9EDF9"><LABEL FOR="rdoUnionFlag1">노조원</LABEL>&nbsp;&nbsp;&nbsp;
								        <INPUT TYPE="RADIO" CLASS="SINPUTTEST_STYLE" NAME="rdoUnionFlag" TAG="24" VALUE="N" ID="rdoUnionFlag2" STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid; BACKGROUND-COLOR: #E9EDF9"><LABEL FOR="rdoUnionFlag2">비노조원</LABEL>
				                	</TD>
				                </TR>
				                <TR>
				                	<TD CLASS="TDFAMILY_TITLE"></TD>
				                	<TD CLASS="TDFAMILY"><INPUT CLASS="SINPUTTEST_STYLE" TYPE=CHECKBOX STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid; BACKGROUND-COLOR: #E9EDF9" TAG="24" NAME="chkPayFlg">
				                	                    임금지급대상여부</TD>
				                	<TD CLASS="TDFAMILY_TITLE"></TD>
				                	<TD CLASS="TDFAMILY"><INPUT CLASS="SINPUTTEST_STYLE" TYPE=CHECKBOX STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid; BACKGROUND-COLOR: #E9EDF9" TAG="24" NAME="chkEmpInsurFlg">
				                	                    고용보험여부</TD>
				                </TR>
				                <TR>
				                	<TD CLASS="TDFAMILY_TITLE"></TD>
				                	<TD CLASS="TDFAMILY"><INPUT CLASS="SINPUTTEST_STYLE" TYPE=CHECKBOX STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid; BACKGROUND-COLOR: #E9EDF9" TAG="24" NAME="chkYearFlg">
				                	                     연월차지급대상</TD>
				                	<TD CLASS="TDFAMILY_TITLE"></TD>
				                	<TD CLASS="TDFAMILY"><INPUT CLASS="SINPUTTEST_STYLE" TYPE=CHECKBOX STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid; BACKGROUND-COLOR: #E9EDF9" TAG="24" NAME="chkRetireFlg">
				                	                    퇴직금지급대상</TD>
				                </TR>
				                <TR>
				                	<TD CLASS="TDFAMILY_TITLE"></TD>
				                	<TD CLASS="TDFAMILY"><INPUT CLASS="SINPUTTEST_STYLE" TYPE=CHECKBOX STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid; BACKGROUND-COLOR: #E9EDF9" TAG="24" NAME="chkTaxFlg">
				                	                    세액계산대상</TD>
				                	<TD CLASS="TDFAMILY_TITLE"></TD>
				                	<TD CLASS="TDFAMILY"><INPUT CLASS="SINPUTTEST_STYLE" TYPE=CHECKBOX STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid; BACKGROUND-COLOR: #E9EDF9" TAG="24" NAME="chkYearTaxFlg">
				                	                        연말정산신고대상</TD>
				                </TR>
				                <TR>
				                	<TD CLASS="TDFAMILY_TITLE">부양자(노)</TD>
				                	<TD CLASS="TDFAMILY"><INPUT CLASS="SINPUTTEST_STYLE" TYPE="Text" size=15 name=txtOld tag="24" style="TEXT-ALIGN: right"></TD>
				                	<TD CLASS="TDFAMILY_TITLE">부양자(소)</TD>
				                	<TD CLASS="TDFAMILY"><INPUT CLASS="SINPUTTEST_STYLE" TYPE="Text" size=15 name=txtYoung tag="24" style="TEXT-ALIGN: right"></TD>
				                </TR>
				                <TR>
				                	<TD CLASS="TDFAMILY_TITLE">장애자</TD>
				                	<TD CLASS="TDFAMILY"><INPUT CLASS="SINPUTTEST_STYLE" TYPE="Text" size=15 name=txtParia tag="24" style="TEXT-ALIGN: right"></TD>
				                	<TD CLASS="TDFAMILY_TITLE">경로자</TD>
				                	<TD CLASS="TDFAMILY"><INPUT CLASS="SINPUTTEST_STYLE" TYPE="Text" size=15 name=txtOldCnt tag="24" style="TEXT-ALIGN: right"></TD>
				                </TR>
				                <TR>
				                	<TD CLASS="TDFAMILY_TITLE">자녀양육수</TD>
				                	<TD CLASS="TDFAMILY"><INPUT CLASS="SINPUTTEST_STYLE" TYPE="Text" size=15 name=txtChild tag="24" style="TEXT-ALIGN: right"></TD>
				                	<TD CLASS="TDFAMILY_TITLE"></TD>
				                	<TD CLASS="TDFAMILY"><INPUT CLASS="SINPUTTEST_STYLE" TYPE=CHECKBOX STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid; BACKGROUND-COLOR: #E9EDF9" TAG="24" NAME="chkSpouseFlg">배우자
				                	                     <INPUT CLASS="SINPUTTEST_STYLE" TYPE=CHECKBOX STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid; BACKGROUND-COLOR: #E9EDF9" TAG="24" NAME="chkLadyFlg">부녀자</TD>
				                </TR>
                            </TABLE>
                        </TD>
                    </TR>
                </TABLE>
            </TD>
            <TD width=14></TD>
        </TR>
    </TABLE>

    <TABLE cellSpacing=0 cellPadding=0 width=700 border=0 bgcolor=#ffffff>
        <TR><TD WIDTH="100%" HEIGHT=0><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=0 FRAMEBORDER=0 SCROLLING=yes noresize framespacing=0></IFRAME></TD></TR>
    </TABLE>

    <INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
</FORM>	

</BODY>
</HTML>
