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
<!-- #Include file="../../inc/incServer.asp"  -->

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
    Dim dilig_strt_dt, dilig_end_dt
    Dim dilig_cd

    dilig_strt_dt = Trim(Request("dilig_strt_dt"))
    dilig_end_dt  = Trim(Request("dilig_end_dt"))
    dilig_cd      = Trim(Request("Dilig_cd"))
   
%>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance


'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID      = "e1601mb1.asp"						           '☆: Biz Logic ASP Name
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" --> 


Dim isOpenPop
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================


Dim dilig_strt_dt,dilig_end_dt
Dim Dilig_cd
Dim gQuery_YN
Dim StartDate
dim fDiligAuth,fAuthCheck
<%
StartDate	= GetSvrDate
%>


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
    else     
            lgKeyStream = Trim(frm1.txtEmp_no.Value) & gColSep   
    end if            
        lgKeyStream = lgKeyStream & Trim(parent.txtinternal_cd.Value) & gColSep
        lgKeyStream = lgKeyStream & UniConvDateAToB(Trim(frm1.txtDilig_strt_dt.value),gDateFormat, gServerDateFormat) & gColSep
        lgKeyStream = lgKeyStream & UniConvDateAToB(Trim(frm1.txtDilig_end_dt.value),gDateFormat, gServerDateFormat) & gColSep
        lgKeyStream = lgKeyStream & Trim(frm1.txtDilig_cd.value) & gColSep
        lgKeyStream = lgKeyStream & Trim(fDiligAuth) & gColSep        
        lgKeyStream = lgKeyStream & Trim(fAuthCheck) & gColSep     
        lgKeyStream = lgKeyStream & Trim(parent.txtEmp_no.Value) & gColSep          
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
	lgKeyStream = replace(lgKeyStream, "'", "''")
End Sub        
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx
    
	 Call CommonQueryRs(" dilig_cd, dilig_nm "," hca010t ", " day_time = " & FilterVar("1", "''", "S") & " " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    Call SetCombo2(frm1.txtdilig_cd, iCodeArr, iNameArr,Chr(11))
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

    dilig_strt_dt = "<%=dilig_strt_dt%>"
	dilig_end_dt  = "<%=dilig_end_dt%>"  
    Dilig_cd = "<%=dilig_cd%>"

    lgIntFlgMode = OPMD_CMODE   'insert mode
    gQuery_YN = ""              
   
    call FncGetDiligAuth(fDiligAuth,fAuthCheck)
'msgbox   fDiligAuth & "*** " &  fAuthCheck
    If Replace(fDiligAuth,Chr(11),"") = "" Then
        parent.document.All("nextprev").style.VISIBILITY = "hidden"
    Else
        parent.document.All("nextprev").style.VISIBILITY = "visible"
    End If

    Call InitComboBox()
    Call LayerShowHide(0)

	Call parent.Click_OpenFrame(Replace(UCase(BIZ_PGM_ID),"MB","MA"))

    if  dilig_strt_dt <> "" then
        frm1.txtdilig_strt_dt.Value = dilig_strt_dt
		frm1.txtdilig_end_dt.Value  = dilig_end_dt
    else 
		frm1.txtdilig_strt_dt.Value = UniConvDateAToB("<%=StartDate%>",gServerDateFormat,gDateFormat)
		frm1.txtdilig_end_dt.Value  = frm1.txtdilig_strt_dt.Value
    end if
	if parent.txtName2.value = "" then
		parent.txtEmp_no2.Value = parent.txtemp_no.value 
	end if
    if  Dilig_cd <> "" then
        Call SetToolBar("01110")
        frm1.txtdilig_cd.value = Dilig_cd
        Call DbQuery(1)
    else    
        Call SetToolBar("01010")
        Call DbQueryEmp(1)
    end if
    Call LockField(Document)


    
End Sub
'========================================================================================
' Function Name : Window_onUnLoad
' Function Desc : 페이지 전환이나 화면이 닫힐 경우 실행해야 될 로직 처리 
'========================================================================================
Sub Form_unLoad()
End Sub

Function DbQuery(ppage)

    Dim strVal
    Dim iret
    Err.Clear                                                                    '☜: Clear err status
	
    DbQuery = False                                                              '☜: Processing is NG

    If Not chkFieldLength(Document, "1") Then									         '☜: This function check required field
		Exit Function
	end if    
    Call LayerShowHide(1)
    
    Call MakeKeyStream("Q")
    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                   '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is NG
End Function

'========================================================================================
' Function Name : DbQueryEmp
' Function Desc : 
'========================================================================================
Function DbQueryEmp(ppage)

    Dim strDate
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQueryEmp = False                                                              '☜: Processing is NG
    Call LayerShowHide(1)
    Call MakeKeyStream("Q")
    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                   '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQueryEmp = True                                                               '☜: Processing is NG
End Function


Function DbQueryOk()
    Err.Clear                                                                    '☜: Clear err status
    
    frm1.txtAppyn.value = "<%=Trim(Request("txtapp_yn"))%>"
    If gQuery_YN = "Y" Then                    '근태코드 아이템 체인지가 일어난경우.
		lgIntFlgMode = OPMD_CMODE              'create mode
    Else                                       '일일근태 현황에서 update를 하려고 들어 왔을때 
		lgIntFlgMode = OPMD_UMODE              'update mode
		ProtectTag(frm1.txtDilig_cd)
		ProtectTag(frm1.txtDilig_STRT_dt)
		ProtectTag(frm1.txtDilig_END_dt)

		if frm1.txtAppyn.value = "Y" or frm1.txtAppyn.value = "R" then
			ProtectTag(frm1.txtRemark)
			ProtectTag(frm1.txtApp_emp_no)
			Call SetToolBar("01000")
		else
			Call SetToolBar("01110")
		end if
		
    End if 
	
    frm1.txtDilig_cd.disabled = true
	gQuery_YN = ""
	
End Function

Function DbQueryFail()
    Err.Clear
    lgIntFlgMode = OPMD_CMODE                'insert mode
End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()

    Dim strDate
	Dim strVal
	Dim strAppyn

	On Error Resume Next
    Err.Clear                                                                    '☜: Clear err status

	DbSave = False														         '☜: Processing is NG

	'----반려처리된 사항에 대해서는 수정이 이루어질 수 없다.
	strAppyn = frm1.txtAppyn.value 
	If strAppyn = "R" Then
		Call DisplayMsgBox("800477","X","X","X")
        exit function
	End IF
		

	With Frm1
        if .txtEmp_no.value = "" then
            Call DisplayMsgBox("800006","X","X","X")
            .txtDilig_cd.focus()
            exit function
        end if

        if .txtEmp_no.value = .txtApp_emp_no.value  then
            Call DisplayMsgBox("800476","X","X","X")
		        .txtapp_name.value      = ""
		        .txtApp_emp_no.focus()
            exit function
        end if

        if .txtDilig_cd.value = "" then
            Call DisplayMsgBox("800094","X","X","X")
            .txtDilig_cd.focus()
            exit function
        end if

        if  Date_chk(.txtDilig_strt_dt.value, strDate) = True then
            .txtDilig_strt_dt.value = strDate
        else
            Call DisplayMsgBox("800094","X","X","X")
            .txtDilig_strt_dt.focus()
            exit function
        end if

        if  Date_chk(.txtDilig_end_dt.value, strDate) = True then
            .txtDilig_end_dt.value = strDate
        else
            Call DisplayMsgBox("800094","X","X","X")
            .txtDilig_end_dt.focus()
            exit function
        end if

		If CompareDateByFormat(.txtDilig_strt_dt.value,.txtDilig_end_dt.value,.txtDilig_strt_dt.Alt,.txtDilig_end_dt.Alt,"970025", gDateFormat, gComDateType,TRUE) = False Then
			frm1.txtDilig_strt_dt.focus
            exit function
		END IF

		if txtApp_emp_no_check() = False then
            exit function
        end if

	End With
	
	If Not chkFieldLength(Document, "2") Then									         '☜: This function check required field
       Exit Function
    End If
    
    Call MakeKeyStream("C")

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
	gQuery_YN = ""	    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call DbQuery(1)

End Function

'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
		
	DbDelete = False			                                                 '☜: Processing is NG
		
	Call LayerShowHide(1)

	With Frm1
		.txtMode.value        = "UID_M0003"                                        '☜: Save
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
	End With

	'------ Developer Coding part (End )   -------------------------------------------------------------- 

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

	DbDelete = True                                                              '⊙: Processing is NG
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	dilig_strt_dt = ""
	Call FncNew()	
End Function

'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	lgIntFlgMode = OPMD_CMODE              'create mode

    Call ClearField(document,2)
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    Call LockField(Document)	
    Call SetToolBar("01010")

	frm1.txtdilig_strt_dt.value = UniConvDateAToB("<%=StartDate%>",gServerDateFormat,gDateFormat)
	frm1.txtdilig_end_dt.value  = frm1.txtdilig_strt_dt.value

    frm1.txtDilig_cd.focus()    
    frm1.txtAppyn.value = ""
    frm1.txtApp_emp_no.value = ""
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    FncNew = True																 '☜: Processing is OK
    
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
'   Event Name : txtApp_emp_no_check()            '<==승인자 이름가져오기 
'   Event Desc :
'========================================================================================================
Function txtApp_emp_no_check()
    On Error Resume Next
    Err.Clear
    
    Dim iDx
    Dim IntRetCd

    txtApp_emp_no_check = False    
    IF frm1.txtApp_emp_no.value = "" THEN
        frm1.txtApp_name.value = ""
        Call DisplayMsgBox("800094","X","X","X")
        frm1.txtApp_emp_no.focus()        
    ELSE

        IntRetCd = CommonQueryRs(" NAME "," HAA010T "," EMP_NO =  " & FilterVar(frm1.txtApp_emp_no.value , "''", "S") & " and retire_dt is null",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
        If IntRetCd = false then
			Call DisplayMsgBox("970000","X","사번","X")
            frm1.txtApp_name.value = ""
            frm1.txtApp_emp_no.focus
        ELSE    
            frm1.txtApp_name.value = Trim(Replace(lgF0,Chr(11),"")) 
            txtApp_emp_no_check = true
        END IF
    END IF 
End Function
'========================================================================================================
'   Event Name : CheckLimit()            '<==승인자 권한체크 
'   Event Desc :
'========================================================================================================

'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================
'========================================================================================================
' Name : OpenEmp()
' Desc : developer describe this line 
'========================================================================================================
Function OpenEmp(pEmpNo)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True or  frm1.txtAppyn.value = "Y"  or  frm1.txtAppyn.value = "R" Then Exit Function

	IsOpenPop = True

	arrParam(0) = frm1.txtApp_Emp_no.value			' Code Condition
	arrParam(1) = ""								' Name Cindition
    arrParam(2) = "APPROVAL_CODE"					'lgUsrIntCd
	arrRet = window.showModalDialog("E1EmpPopa3.asp", Array(arrParam), _
		"dialogWidth=540px; dialogHeight=385px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
	    frm1.txtApp_emp_no.value = arrRet(0)
	    frm1.txtApp_name.value = arrRet(1)
	End If	
			
End Function
'========================================================================================================
' Name : FncGetDiligAuth()
' Desc : developer describe this line 
'========================================================================================================
Function FncGetDiligAuth(fDiligAuth,fAuthCheck)
    fDiligAuth = ""
    fAuthCheck = ""
    Call CommonQueryRs(" internal_cd,internal_auth "," e11090t "," emp_no =  " & FilterVar(parent.txtEmp_no.Value, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    fDiligAuth = replace(lgF0,chr(11),chr(12))
    fDiligAuth = replace(fDiligAuth," ","")    
    fAuthCheck = replace(lgF1,chr(11),chr(12))
    fAuthCheck = replace(fAuthCheck," ","")      
End Function


Sub txtDilig_cd_OnChange()
    gQuery_YN = "Y"                                                    ' 아이템 체인지를 인식한다.전역변수 
    
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
Function txtEmp_no2_Onchange()
        Call DbQuery(1)	
End Function
Function Document_onKeyDown()
Dim CuEvObj,KeyCode
	Set CuEvObj = window.event.srcElement		
	KeyCode = window.event.keycode
	Select Case KeyCode
		Case 13 'enter key
	End Select		
	Document_onKeyDown	= True	
End Function

Function txtApp_emp_no_onKeyDown()
Dim CuEvObj,KeyCode,IntRetCd
	Set CuEvObj = window.event.srcElement		
	KeyCode = window.event.keycode
	Select Case KeyCode
		Case 13 'enter key
	End Select		
	txtApp_emp_no_onKeyDown	= True	
End Function
'========================================================================================================
'   Event Name : txtApp_emp_no_Onchange()            '<==사번으로 이름가져오기 
'   Event Desc :
'========================================================================================================
Function txtApp_emp_no_Onchange()
    On Error Resume Next
    Err.Clear
    
    Dim iDx
    Dim IntRetCd
    Dim strEmp_no

	frm1.txtApp_emp_no.value=Trim(frm1.txtApp_emp_no.value)
    IF Trim(frm1.txtApp_emp_no.value) = "" THEN
        frm1.txtApp_name.value = ""
        txtApp_emp_no_Onchange = true
    ELSE
		strEmp_no = Trim(frm1.txtApp_emp_no.value)
        IntRetCd = CommonQueryRs(" NAME "," HAA010T "," EMP_NO =  " & FilterVar(strEmp_no , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
        If IntRetCd = false then
            frm1.txtApp_name.value = ""
            frm1.txtApp_emp_no.focus
        ELSE    
            frm1.txtApp_name.value = Trim(Replace(ConvSPChars(lgF0),Chr(11),""))   '사번에 해당하는 이름 
            txtApp_emp_no_Onchange = true
        END IF
    END IF 
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
    <TABLE cellSpacing=0 cellPadding=0 width=770 border=0>
        <TR>
            <TD width=13></TD>
            <TD>
                <TABLE cellSpacing=0 cellPadding=0 width=743 border=0 bgcolor=#ffffff>
                    <TR height=45 valign=center>
                        <TD class=base1>사번:<INPUT class=base1 NAME="txtEmp_no" MAXLENGTH=13 SiZE=12 TAG=14></TD>
                        <TD class=base1>성명:<INPUT class=base1 NAME="txtName" MAXLENGTH=20 SiZE=10 TAG=14></TD>
                        <TD class=base1>직위:<INPUT class=base1 NAME="txtroll_pstn" MAXLENGTH=20 SiZE=10 TAG=14></TD>
                        <TD class=base1>부서:<INPUT class=base1 NAME="txtDept_nm" MAXLENGTH=25 SiZE=20 TAG=14></TD>
                    </TR>
                    <TR>
                        <TD colspan=4>
                            <TABLE cellSpacing=1 cellPadding=0 width=100% border=0 bgcolor=#ffffff>
                                <TR>
		                            <TD CLASS="TDFAMILY_TITLE" NOWRAP>근태</TD>
		                            <TD CLASS="TDFAMILY2" COLSPAN=3><SELECT NAME="txtDilig_cd" ALT="근태코드" STYLE="WIDTH: 150px" TAG="22"><OPTION VALUE=""></OPTION></SELECT>
		                            </TD>
                                </TR>
                                <TR>
		                            <TD CLASS="TDFAMILY_TITLE" NOWRAP>근태기간</TD>
		                            <TD CLASS="TDFAMILY2" COLSPAN=3>
		                                <INPUT CLASS="SINPUTTEST_STYLE" ID="txtDilig_STRT_dt" NAME="txtDilig_STRT_dt" TYPE="Text" MAXLENGTH=10 SiZE=10 alt="근태시작일" tag="22D" ondblclick="VBScript:Call OpenCalendar('txtDilig_STRT_dt',3)">&nbsp;~
		                                <INPUT CLASS="SINPUTTEST_STYLE" ID="txtDilig_END_dt" NAME="txtDilig_END_dt" TYPE="Text" MAXLENGTH=10 SiZE=10 alt="근태종료일" tag="22D" ondblclick="VBScript:Call OpenCalendar('txtDilig_END_dt',3)">
		                            </TD>
                                </TR>
                                <TR>
		                            <TD CLASS="TDFAMILY_TITLE" NOWRAP>사유</TD>
		                            <TD CLASS="TDFAMILY2" COLSPAN=3><INPUT CLASS="SINPUTTEST_STYLE" NAME="txtRemark" ALT="사유" TYPE="Text" MAXLENGTH=39 SiZE=40 tag="22">
		                            </TD>
                                </TR>
		                        <TR>
		                            <TD CLASS="TDFAMILY_TITLE" NOWRAP>승인자</TD>
		                            <TD CLASS="TDFAMILY2" colspan=3><INPUT CLASS="SINPUTTEST_STYLE" NAME="txtApp_emp_no" ALT="승인사번" TYPE="Text" MAXLENGTH=13 SiZE=13 tag="22"><IMG SRC="../../../Cshared/Image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="VBScript:Call OpenEmp(frm1.txtApp_emp_no.value)">
		                                    <INPUT CLASS="SINPUTTEST_STYLE" NAME="txtApp_name" TYPE="Text" MAXLENGTH=20 SiZE=20 tag="24">
		                            </TD>
                                </TR>
		                        <TR height=150>
		                            <TD CLASS="TDFAMILY" NOWRAP colspan=4></TD>
                                </TR>
		                        <TR height=10>
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

    <TABLE cellSpacing=2 cellPadding=0 width=700 border=0 bgcolor=#ffffff>
        <TR><TD WIDTH="100%" HEIGHT=0><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=0 FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME></TD></TR>
    </TABLE>

    <INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">

    <INPUT TYPE=HIDDEN NAME="txtAppyn"    TAG="24">
</FORM>	

</BODY>
</HTML>
