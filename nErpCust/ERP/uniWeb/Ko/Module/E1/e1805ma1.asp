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
<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance


'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID      = "e1805mb1.asp"						           '☆: Biz Logic ASP Name
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

Dim IsOpenPop
Dim emp_no,Dept_cd
Dim lgEnterChk

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
emp_no = "<%=Trim(Request("emp_no"))%>"
Dept_cd = "<%=Trim(Request("Dept_cd"))%>"

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
        lgKeyStream = Trim(frm1.txtEmp_no1.Value) & gColSep       'You Must append one character(gColSep)
        lgKeyStream = lgKeyStream & UCase(Trim(frm1.txtDept_cd.Value)) & gColSep       'You Must append one character(gColSep)
        lgKeyStream = lgKeyStream & Trim(parent.txtinternal_cd.Value) & gColSep
        lgKeyStream = lgKeyStream & UCase(Trim(Dept_cd)) & gColSep       'You Must append one character(gColSep)
        lgKeyStream = lgKeyStream & UCase(Trim(frm1.txtuse_ynv.Value)) & gColSep
        lgKeyStream = lgKeyStream & gEmpNo & gColSep                                
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr

    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
    parent.document.All("nextprev").style.VISIBILITY = "hidden"

    lgIntFlgMode = OPMD_CMODE   'insert mode
	lgEnterChk = false
	
    Call InitComboBox()
    Call LockField(Document)	
    Call LayerShowHide(0)
	Call SetToolBar("01110")    
    frm1.txtuse_yn2.checked = true	
    if  emp_no <> "" then
        frm1.txtemp_no1.value = emp_no
        frm1.txtdept_cd.value = dept_cd
        Call parent.Click_OpenFrame(Replace(UCase(BIZ_PGM_ID),"MB","MA"))
        Call DbQuery(1)
	    ProtectTag(frm1.txtemp_no1)
    End If
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
	if lgEnterChk = false then

		Call LayerShowHide(1)
		Call MakeKeyStream("Q")

		strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                   '☜: Query
		strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
		Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
	end if
    DbQuery = True                                                               '☜: Processing is NG
End Function

Function DbQueryOk()
    Err.Clear                                                                    '☜: Clear err status
    If emp_no = "" Then
        lgIntFlgMode = OPMD_CMODE   'create mode
    Else
        lgIntFlgMode = OPMD_UMODE   'update mode
    End If
       ProtectTag(frm1.txtemp_no1)
       ProtectTag(frm1.txtDept_cd)
       Call SetToolbar("01110")
    lgIntFlgMode = OPMD_UMODE       

    if  frm1.txtuse_ynv.value = "Y" then
        frm1.txtuse_yn1.checked = true
        frm1.txtuse_yn2.checked = false
    else
        frm1.txtuse_yn2.checked = true
        frm1.txtuse_yn1.checked = false
    end if
    
End Function

Function DbQueryFail()
    Err.Clear
    lgIntFlgMode = OPMD_CMODE   'insert mode
End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()
	Dim strVal
	Dim strDate
	Dim lgStrSQL
	Dim lgObjConn
    Err.Clear                                                                    '☜: Clear err status

	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	With Frm1
	
        if  Trim(frm1.txtemp_no1.value = "") then
            Call  DisplayMsgBox("800094","X","X","X")
            frm1.txtemp_no1.focus()
            exit function
        end if
        if  Trim(frm1.txtDept_cd.value) = "" then
            Call  DisplayMsgBox("800094","X","X","X")
            frm1.txtDept_cd.focus()
            exit function
        end if
        
        if  frm1.txtname1.value = "" then
			Call DisplayMsgBox("970000", "x", "사번", "x")
            frm1.txtemp_no1.focus()
            exit function
        end if
        if  frm1.txtDept_nm.value = "" then
			Call DisplayMsgBox("970000", "x", "부서코드", "x")
            frm1.txtDept_cd.focus()
            exit function
        end if
        if  frm1.txtuse_yn1.checked = true then
            frm1.txtuse_ynv.value = "Y"
        else
            frm1.txtuse_ynv.value = "N"
        end if
	End With
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

	DbSave = False														         '☜: Processing is NG
		
	Call LayerShowHide(1)
    Call MakeKeyStream("Q")

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
    Call DbQuery(1)

End Function

'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
	Dim strVal
	Dim IntRetCD,where_stm, IntOkCd
	
    Err.Clear                                                                    '☜: Clear err status

	With Frm1

        if  Trim(frm1.txtemp_no1.value = "") then
            Call  DisplayMsgBox("800094","X","X","X")
            frm1.txtemp_no1.focus()
            exit function
        end if
        if  Trim(frm1.txtDept_cd.value) = "" then
            Call  DisplayMsgBox("800094","X","X","X")
            frm1.txtDept_cd.focus()
            exit function
        end if
        
        if  frm1.txtname1.value = "" then
			Call DisplayMsgBox("970000", "x", "사번", "x")
            frm1.txtemp_no1.focus()
            exit function
        end if
        if  frm1.txtDept_nm.value = "" then
			Call DisplayMsgBox("970000", "x", "부서코드", "x")
            frm1.txtDept_cd.focus()
            exit function
        end if

	End With
    
    where_stm = " emp_no = " & FilterVar(frm1.txtemp_no1.value, "''", "S")                              ' 사번char(10)
    where_stm = where_stm & " and   dept_cd = " & FilterVar(frm1.txtDept_cd.value, "''", "S") 
    
    IntOkCd = CommonQueryRs(" emp_no "," E11090T ",where_stm,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    if IntOkCd=false or lgF0=null then
			Call DisplayMsgBox("970000", "x", "근태담당자", "x")
			frm1.txtname1.value = ""
			frm1.txtDept_nm.value = ""
			exit function
    end if

	IntRetCD =  DisplayMsgBox("900003", VB_YES_NO,"x","x")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If

	DbDelete = False			                                                 '☜: Processing is NG

	Call LayerShowHide(1)
    Call MakeKeyStream("Q")
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

	'Call InitVariables()

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call FncNew()	
	'Frm1.txtFlgMode.value = OPMD_CMODE

	'Call DbQuery(1)
End Function

'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    Call ClearField(document,2)
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Call LockField(Document)
    Call SetToolbar("01110")
    emp_no = ""
    Dept_cd = ""
    'ReleaseTag(frm1.txtemp_no1)                                                   '⊙: Initializes local global variables
    frm1.txtemp_no1.focus()    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	lgIntFlgMode= OPMD_CMODE    
    FncNew = True																 '☜: Processing is OK
End Function

Sub SubPrint(objFrame)
    Set objActiveEl = document.activeElement
    objFrame.focus()
    objFrame.print()
    objActiveEl.focus
    Set objActiveEl = nothing
End Sub

'========================================================================================================
'========================================================================================================
' Name : OpenDept()
' Desc : developer describe this line 
'========================================================================================================
Function OpenDept()
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True  or lgIntFlgMode = OPMD_UMODE Then Exit Function

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtDept_cd.value)			    ' Code Condition
	arrParam(1) = ""'frm1.txtName1.value			' Name Cindition
'   arrParam(2) = Trim(parent.txtinternal_cd.Value) 'lgUsrIntCd
	
	arrRet = window.showModalDialog("E1DeptPopa1.asp", Array(arrParam), _
		"dialogWidth=600px; dialogHeight=385px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
	    frm1.txtDept_cd.value = Trim(arrRet(0))
	    frm1.txtDept_nm.value = Trim(arrRet(1))
	End If	
			
End Function
'========================================================================================================
' Name : OpenEmp()
' Desc : developer describe this line 
'========================================================================================================
Function OpenEmp(pEmpNo)
	Dim arrRet
	Dim arrParam(2)
'msgbox lgIntFlgMode & ":"& OPMD_UMODE
	If IsOpenPop = True   or lgIntFlgMode = OPMD_UMODE  Then Exit Function

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtEmp_no1.value)			' Code Condition
	arrParam(1) = ""'frm1.txtName1.value			' Name Cindition
    arrParam(2) = Trim(parent.txtinternal_cd.Value)'lgUsrIntCd
	
	arrRet = window.showModalDialog("E1EmpPopa4.asp", Array(arrParam), _
		"dialogWidth=615px; dialogHeight=375px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
	    frm1.txtemp_no1.value = Trim(arrRet(0))
	    frm1.txtname1.value = Trim(arrRet(1))
	End If	
			
End Function
'========================================================================================================
'   Event Name : txtemp_no1_Onchange()            '<==사번으로 이름가져오기 
'   Event Desc :
'========================================================================================================
Function txtemp_no1_Onchange()
    On Error Resume Next
    Err.Clear
    
    Dim iDx
    Dim IntRetCd
    Dim strEmp_no

	frm1.txtemp_no1.value=Trim(frm1.txtemp_no1.value)
    IF Trim(frm1.txtemp_no1.value) = "" THEN
        frm1.txtname1.value = ""
        txtemp_no1_Onchange = true
    ELSE
		strEmp_no = Trim(frm1.txtemp_no1.value)
        IntRetCd = CommonQueryRs(" NAME "," HAA010T "," EMP_NO =  " & FilterVar(strEmp_no , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
        If IntRetCd = false then
            frm1.txtname1.value = ""
            frm1.txtemp_no1.focus
        ELSE    
            frm1.txtname1.value = Trim(Replace(ConvSPChars(lgF0),Chr(11),""))   '사번에 해당하는 이름 
            txtemp_no1_Onchange = true
        END IF
    END IF 
End Function


Function txtDept_cd_Onchange()
    Dim iKey1
    On Error Resume Next
    Err.Clear
    frm1.txtDept_cd.value=UCase(Trim(frm1.txtDept_cd.value))
    iKey1 = " dept_cd = " & FilterVar(frm1.txtDept_cd.value, "''", "S") 
    iKey1 = iKey1 & " AND org_change_dt = (select max(org_change_dt) from b_acct_dept where org_change_dt<=getdate())"
			
	frm1.txtDept_nm.value = ""
	IntRetCd = CommonQueryRs(" dept_nm "," b_acct_dept ",iKey1,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 		
	if IntRetCd  then
		frm1.txtDept_nm.value= left(lgF0,len(lgF0)-1)
	else
		frm1.txtDept_cd.focus()
	end if
End Function

Function txtDept_cd_onKeyDown()
Dim CuEvObj,KeyCode,IntRetCd,iKey1
	Set CuEvObj = window.event.srcElement		
	KeyCode = window.event.keycode
	Select Case KeyCode
		Case 13 'enter key
			lgEnterChk = true
			call txtDept_cd_Onchange()
	End Select		
	txtDept_cd_onKeyDown	= true	
End Function

Function txtemp_no1_onKeyDown()
Dim CuEvObj,KeyCode,IntRetCd
	Set CuEvObj = window.event.srcElement		
	KeyCode = window.event.keycode
	Select Case KeyCode
		Case 13 'enter key
			lgEnterChk = True
			call txtemp_no1_Onchange()
	End Select		
	txtemp_no1_onKeyDown	= True	
End Function

'===========================================================,=============================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================

<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->
</SCRIPT>
<!-- #Include file="../../inc/uniSimsClassID.inc" -->

</HEAD>

<BODY topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
    <TABLE cellSpacing=0 cellPadding=0 width=754 border=0>
        <TR>
            <TD width=13></TD>
            <TD>
                <TABLE cellSpacing=0 cellPadding=0 width=727 border=0 bgcolor=#ffffff>
                    <TR height=25 valign=center>
                        <TD colspan=4></TD>
                    </TR>
                    <TR>
                        <TD colspan=4>
                            <TABLE cellSpacing=1 cellPadding=0 width=100% border=0 bgcolor=#ffffff>
		                        <TR>
		                            <TD CLASS="TDFAMILY_TITLE" NOWRAP>근태담당자</TD>
		                            <TD CLASS="TDFAMILY2" colspan=3><INPUT CLASS="SINPUTTEST_STYLE" NAME="txtemp_no1" TYPE="Text" MAXLENGTH=13 SiZE=13 tag="22"><IMG SRC="../../../Cshared/Image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="VBScript:Call OpenEmp(frm1.txtemp_no1.value)">
		                                    <INPUT CLASS="SINPUTTEST_STYLE" NAME="txtname1" TYPE="Text" MAXLENGTH=20 SiZE=20 tag="24">
		                            </TD>
                                </TR>
		                        <TR>
		                            <TD CLASS="TDFAMILY_TITLE" NOWRAP>부서</TD>
		                            <TD CLASS="TDFAMILY2" colspan=3><INPUT CLASS="SINPUTTEST_STYLE" NAME="txtDept_cd" TYPE="Text" MAXLENGTH=13 SiZE=13 tag="22XXXU"><IMG SRC="../../../Cshared/Image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="VBScript:Call OpenDept()">
		                                    <INPUT CLASS="SINPUTTEST_STYLE" NAME="txtDept_nm" TYPE="Text" MAXLENGTH=20 SiZE=20 tag="24">
		                                    <INPUT CLASS="SINPUTTEST_STYLE" NAME="txtInternal_cd" TYPE="Hidden" MAXLENGTH=20 SiZE=20 tag="24">
		                            </TD>
                                </TR>
		                        <TR>
		                            <TD CLASS="TDFAMILY_TITLE" NOWRAP>하위부서권한</TD>
		                            <TD CLASS="TDFAMILY2" colspan=3>
		                           	    <INPUT TYPE="RADIO" NAME="txtuse_yn" tag="22" CHECKED ID="txtuse_yn1" VALUE=1 STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid; BACKGROUND-COLOR: #E9EDF9"><LABEL FOR="txtuse_yn1">사용</LABEL>
    					                <INPUT TYPE="RADIO" NAME="txtuse_yn" tag="22" ID="txtuse_yn2" VALUE=2 STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid; BACKGROUND-COLOR: #E9EDF9"><LABEL FOR="txtuse_yn2">미사용</LABEL>
                                        <INPUT TYPE=HIDDEN NAME="txtuse_ynv">
                                    </TD>      
                                </TR>                                
		                        <TR height=110>
		                            <TD CLASS="TDFAMILY_TITLE" NOWRAP></TD>
		                            <TD CLASS="TDFAMILY2" colspan=3></TD>
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

    <TABLE cellSpacing=0 cellPadding=0 width=700 border=0 bgcolor=#ffffff>
        <TR><TD HEIGHT=0><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=0 FRAMEBORDER=0 SCROLLING=yes noresize framespacing=0></IFRAME></TD></TR>
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
