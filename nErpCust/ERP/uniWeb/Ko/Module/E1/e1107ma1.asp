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

Const BIZ_PGM_ID      = "e1107mb1.asp"						           '☆: Biz Logic ASP Name
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" --> 

'========================================================================================================
'=                       4.3 Common variables
'========================================================================================================


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

        if  Trim(parent.txtEmp_no2.Value) = "" then
            lgKeyStream = Trim(parent.txtEmp_no.Value) & gColSep
        else
            lgKeyStream = Trim(parent.txtEmp_no2.Value) & gColSep
        end if
        lgKeyStream = lgKeyStream & Trim(parent.txtinternal_cd.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(parent.txtdept_nm.value) & gColSep        
        lgKeyStream = lgKeyStream & "" & gColSep
    else
        lgKeyStream = Trim(frm1.txtEmp_no.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(parent.txtinternal_cd.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(parent.txtdept_nm.value) & gColSep        
    end if
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
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
  Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0019", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    iCodeArr = lgF0
    iNameArr = lgF1
    Call SetCombo2(frm1.txtMil_type, iCodeArr, iNameArr, Chr(11))    

  Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0020", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    iCodeArr = lgF0
    iNameArr = lgF1
    Call SetCombo2(frm1.txtMil_kind, iCodeArr, iNameArr, Chr(11))    

  Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0021", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    iCodeArr = lgF0
    iNameArr = lgF1
    Call SetCombo2(frm1.txtMil_grade, iCodeArr, iNameArr, Chr(11))    

  Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0022", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    iCodeArr = lgF0
    iNameArr = lgF1
    Call SetCombo2(frm1.txtMil_branch, iCodeArr, iNameArr, Chr(11))    



	'------ Developer Coding part (End )   -------------------------------------------------------------- 
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
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
    if  parent.txtDEPT_AUTH.value = "Y" then
        parent.document.All("nextprev").style.VISIBILITY = "visible"
        Call SetToolBar("10010")    
    else
        parent.document.All("nextprev").style.VISIBILITY = "hidden"
        Call SetToolBar("00010")    
    end if

    Call InitComboBox()

    Call LayerShowHide(0)
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

'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	
	frm1.txtEmp_no.value = ""
	frm1.txtName.value = ""
	frm1.txtroll_pstn.value = ""
	frm1.txtDept_nm.value = ""
	
    Call ClearField(document,2)
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    FncNew = True																 '☜: Processing is OK
End Function

Function DbQuery(ppage)

    Dim strVal,IntRetCD
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG
    
    Call ClearField(document,2)
    Call LayerShowHide(1)
    Call MakeKeyStream("Q")

    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                   '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is NG

End Function

Function DbQueryOk()
    Err.Clear                                                                    '☜: Clear err status
    'Call Grid1.ShowData(frm1,1)
'    Call ElementVisible(window.parent.document.all("RunQuery"), 1)
    if  parent.txtDEPT_AUTH.value = "Y" then
        parent.document.All("nextprev").style.VISIBILITY = "visible"
        if  frm1.txtEmp_no.value = parent.txtEmp_no.Value then
            Call SetToolBar("10010")
        else
            Call SetToolBar("10000")
        end if
    else
        parent.document.All("nextprev").style.VISIBILITY = "hidden"
        Call SetToolBar("00010")
    end if

End Function

Function DbQueryFail()
    Err.Clear
    if  parent.txtDEPT_AUTH.value = "Y" then
        parent.document.All("nextprev").style.VISIBILITY = "visible"
        if  frm1.txtEmp_no.value = parent.txtEmp_no.Value then
            Call SetToolBar("10010")
        else
            Call SetToolBar("10000")
        end if
    else
        parent.document.All("nextprev").style.VISIBILITY = "hidden"
        Call SetToolBar("00010")
    end if

    Call ClearField(Document,2)                                                                    '☜: Clear err status

End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()
	Dim strVal
	Dim strDate
    Err.Clear                                                                    '☜: Clear err status

    if frm1.txtMil_start.value = "" then
    elseif Date_chk(frm1.txtMil_start.value, strDate) = True then
        frm1.txtMil_start.value = strDate
    else
        Call DisplayMsgBox("800094","X","X","X")
        frm1.txtMil_start.focus()
        exit function
    end if

    if frm1.txtMil_end.value = "" then
    elseif Date_chk(frm1.txtMil_end.value, strDate) = True then
        frm1.txtMil_end.value = strDate
    else
        Call DisplayMsgBox("800094","X","X","X")
        frm1.txtMil_end.focus()
        exit function
    end if

	With Frm1	
	    if chkField(Document,"2") then
	    	exit function
	    end if
	    if  uniConvDateToYYYYMMDD(frm1.txtMil_start.value,gDateFormat,"-") >= uniConvDateToYYYYMMDD(frm1.txtMil_end.value,gDateFormat,"-") then
            Call DisplayMsgBox("800002","X","X","X")
           .txtMil_start.focus()
            exit function
        end if
	End With
    
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

Function txtEmp_no2_Onchange()
        Call DbQuery(1)	
End Function
'========================================================================================================
' Name : OpenSItemDC()
' Desc : developer describe this line 
'========================================================================================================
Function OpenSItemDC(strCode, iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
	    Case 1
            arrParam(0) = "H0019"
	    Case 2
            arrParam(0) = "H0020"
	    Case 3
            arrParam(0) = "H0021"
        Case 4
            arrParam(0) = "H0022"
	End Select
	arrParam(1) = strCode			' Code Condition
	arrParam(2) = "1"
	
	arrRet = window.showModalDialog("E1CodePopa1.asp", Array(arrParam), _
		"dialogWidth=538px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
	    Select Case iWhere
	        Case 1
	            frm1.txtMil_kind.value = arrRet(0)
	            frm1.txtMil_kind_nm.value = arrRet(1)
	        Case 2
	            frm1.txtMil_kind.value = arrRet(0)
	            frm1.txtMil_kind_nm.value = arrRet(1)
	        Case 3
	            frm1.txtMil_kind.value = arrRet(0)
	            frm1.txtMil_kind_nm.value = arrRet(1)
            Case 4
	            frm1.txtMil_kind.value = arrRet(0)
	            frm1.txtMil_kind_nm.value = arrRet(1)
	    End Select
	End If	

End Function
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
    <TABLE cellSpacing=0 cellPadding=0 border=0 width=770>
        <TR>
            <TD width=13></TD>
            <TD>
                <TABLE cellSpacing=0 cellPadding=0 border=0 bgcolor=#ffffff width=743>
                    <TR height=36 valign=center>
                        <TD class=base1>사번:<INPUT class=base1 NAME="txtEmp_no" MAXLENGTH=13 SiZE=12 tag=14></TD>
                        <TD class=base1>성명:<INPUT class=base1 NAME="txtName" MAXLENGTH=20 SiZE=10  tag=14></TD>
                        <TD class=base1>직위:<INPUT class=base1 NAME="txtroll_pstn" MAXLENGTH=20 SiZE=10  tag=14></TD>
                        <TD class=base1>부서:<INPUT class=base1 NAME="txtDept_nm" MAXLENGTH=25 SiZE=15  tag=14></TD>
                    </TR>
                    <TR>
                        <TD colspan=4>
                            <TABLE cellSpacing=1 cellPadding=0 border=0 bgcolor=#ffffff width=100%>
                                <TR>
		                            <TD CLASS="TDFAMILY_TITLE" NOWRAP>병역구분</TD>
		                            <TD CLASS="TDFAMILY">
		                                <SELECT NAME="txtMil_type" STYLE="WIDTH: 120px" TAG="22"><OPTION VALUE=""></OPTION></SELECT>
		                            </TD>
		                            <TD CLASS="TDFAMILY_TITLE" NOWRAP>병역군별</TD>
		                            <TD CLASS="TDFAMILY">
		                                <SELECT NAME="txtMil_kind" STYLE="WIDTH: 120px" TAG="22"><OPTION VALUE=""></OPTION></SELECT>
	                                </TD>
                                </TR>
                                <TR>
		                            <TD CLASS="TDFAMILY_TITLE" NOWRAP>복무기간</TD>
		                            <TD CLASS="TDFAMILY">
		                                <INPUT CLASS="SINPUTTEST_STYLE" NAME="txtMil_start" ALT="복무기간1" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="22" ondblclick="VBScript:Call OpenCalendar('txtMil_start',3)">&nbsp;~&nbsp;
		                                <INPUT CLASS="SINPUTTEST_STYLE" NAME="txtMil_End" ALT="복무기간2" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="22" ondblclick="VBScript:Call OpenCalendar('txtMil_End',3)">
                                    </TD>      
		                            <TD CLASS="TDFAMILY_TITLE" NOWRAP>병역등급</TD>
		                            <TD CLASS="TDFAMILY">
		                                <SELECT NAME="txtMil_grade" STYLE="WIDTH: 120px" TAG="22"><OPTION VALUE=""></OPTION></SELECT>
		                            </TD>
		                        </TR>
		                        <TR>
		                            <TD CLASS="TDFAMILY_TITLE" NOWRAP>병역병과</TD>
		                            <TD CLASS="TDFAMILY">
		                                <SELECT NAME="txtMil_branch" STYLE="WIDTH: 120px" TAG="22"><OPTION VALUE=""></OPTION></SELECT>
		                            </TD>
		                            <TD CLASS="TDFAMILY_TITLE" NOWRAP>군번</TD>
		                            <TD CLASS="TDFAMILY"><INPUT CLASS="SINPUTTEST_STYLE" NAME="txtMil_no" ALT="군번" TYPE="Text" MAXLENGTH=10 SiZE=12 tag="22XU"></TD>
                                </TR>
		                        <TR height=180>
		                            <TD CLASS="TDFAMILY_TITLE" NOWRAP></TD>
		                            <TD CLASS="TDFAMILY"></TD>
		                            <TD CLASS="TDFAMILY_TITLE" NOWRAP></TD>
		                            <TD CLASS="TDFAMILY"></TD>
                                </TR>
		                        <TR height=5>
		                            <TD></TD>
		                            <TD></TD>
		                            <TD></TD>
		                            <TD></TD>
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

    <INPUT TYPE=HIDDEN NAME="txtMode">
    <INPUT TYPE=HIDDEN NAME="txtKeyStream">
    <INPUT TYPE=HIDDEN NAME="txtUpdtUserId">
    <INPUT TYPE=HIDDEN NAME="txtInsrtUserId">
    <INPUT TYPE=HIDDEN NAME="txtFlgMode">
    <INPUT TYPE=HIDDEN NAME="txtPrevNext">

</FORM>	

</BODY>
</HTML>
