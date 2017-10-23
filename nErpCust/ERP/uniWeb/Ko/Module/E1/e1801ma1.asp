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
    Dim emp_no,updateok
    emp_no = Trim(Request("emp_no"))
    updateok = Trim(Request("updateok"))

%>


<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance


'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID      = "e1801mb1.asp"						           '☆: Biz Logic ASP Name
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
Dim emp_no
dim updateok

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
        lgKeyStream = Trim(parent.txtEmp_no.Value) & gColSep       'You Must append one character(gColSep)
        lgKeyStream = lgKeyStream & "" & gColSep
        lgKeyStream = lgKeyStream & Trim(frm1.txtEmp_no1.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(frm1.txtEmp_no1.Value) & gColSep          
    else
        lgKeyStream = Trim(parent.txtEmp_no.Value) & gColSep       'You Must append one character(gColSep)
        lgKeyStream = lgKeyStream & "" & gColSep
        lgKeyStream = lgKeyStream & Trim(frm1.txtEmp_no1.Value) & gColSep 
        lgKeyStream = lgKeyStream & Trim(frm1.txtEmp_no1.Value) & gColSep         
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

    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
  Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0120", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    Call SetCombo2(frm1.txtpro_auth, iCodeArr, iNameArr, Chr(11))    

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status

    parent.document.All("nextprev").style.VISIBILITY = "hidden"
    frm1.txtemp_no1.focus() 

    emp_no = "<%=emp_no%>"
	updateok= "<%=updateok%>"
    lgIntFlgMode = OPMD_CMODE   'insert mode

    Call InitComboBox()
    Call LockField(Document)	

    Call LayerShowHide(0)

    Call SetToolBar("01110")


    if  emp_no <> "" then
        frm1.txtemp_no1.value = emp_no
        Call parent.Click_OpenFrame(Replace(UCase(BIZ_PGM_ID),"MB","MA"))
        Call DbQuery(1)
    end if
     ProtectTag(frm1.txtuser_id)
'    ProtectTag(frm1.txtemp_no1)

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

    lgIntFlgMode = OPMD_UMODE   'update mode

    if  frm1.txtdept_authv.value = "Y" then
        frm1.txtdept_auth1.checked = true
        frm1.txtdept_auth2.checked = false
    else
        frm1.txtdept_auth2.checked = true
        frm1.txtdept_auth1.checked = false
    end if


    if  frm1.txtuse_ynv.value = "Y" then
        frm1.txtuse_yn1.checked = true
        frm1.txtuse_yn2.checked = false
    else
        frm1.txtuse_yn2.checked = true
        frm1.txtuse_yn1.checked = false
    end if
		frm1.txtpassword.focus()
		ProtectTag(frm1.txtemp_no1)
		ProtectTag(frm1.txtuser_id)
End Function

Function DbQueryFail()
    Err.Clear
    Call ClearField(Document,2)                                                                    '☜: Clear err status
'    Call ElementVisible(window.parent.document.all("RunQuery"), 1)

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
	
		if .txtemp_no1.value = "" then
		   Call  DisplayMsgBox("800478","X","X","X")
		   .txtemp_no1.focus()
		   exit function
		end if
        if  frm1.txtpro_auth.value = "" then
            Call  DisplayMsgBox("800094","X","X","X")
            frm1.txtpro_auth.focus()
            exit function
        end if

        if  frm1.txtdept_auth1.checked = true then
            frm1.txtdept_authv.value = "Y"
        else
            frm1.txtdept_authv.value = "N"
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
    if lgIntFlgMode <> OPMD_UMODE then
		if emp_no_check()=false then
			Call LayerShowHide(0)
			Exit Function
		end if
	end if
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
    'Call InitVariables
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call DbQuery(1)
End Function

'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
	Dim strVal
	Dim IntRetCD
	
    Err.Clear                                                                    '☜: Clear err status

	
	if  frm1.txtemp_no1.value = "" then
       Call  DisplayMsgBox("800478","X","X","X")
       frm1.txtemp_no1.focus()
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
    Call LockField(Document)
 
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    Call SetToolbar("0111")
    frm1.txtdept_auth1.checked = true
    frm1.txtuse_yn1.checked = true
    frm1.txtemp_no1.focus()    
     ProtectTag(frm1.txtuser_id)    
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

'========================================================================================================
'========================================================================================================
' Name : OpenEmp()
' Desc : developer describe this line 
'========================================================================================================
Function OpenEmp(pEmpNo)
	Dim arrRet
	Dim arrParam(2)
	
	if lgIntFlgMode = OPMD_UMODE then
		Exit Function
	end if

	If IsOpenPop = True Then Exit Function
	If frm1.txtemp_no1.className = "ted" Then Exit Function
	IsOpenPop = True

	arrParam(0) = frm1.txtEmp_no1.value			' Code Condition
	arrParam(1) = ""'frm1.txtName1.value			' Name Cindition
    arrParam(2) = Trim(parent.txtinternal_cd.Value)'lgUsrIntCd
	
	arrRet = window.showModalDialog("E1EmpPopa2.asp", Array(arrParam), _
		"dialogWidth=615px; dialogHeight=410px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
	    frm1.txtemp_no1.value = arrRet(0)
	    frm1.txtname1.value = arrRet(1)
	    frm1.txtuser_id.value = arrRet(2)
		frm1.txtpassword.focus()
'		ProtectTag(frm1.txtemp_no1)
'		ProtectTag(frm1.txtuser_id)
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
Function Document_onKeyDown()
Dim CuEvObj,KeyCode
	Set CuEvObj = window.event.srcElement		
	KeyCode = window.event.keycode
	Select Case KeyCode
		Case 13 'enter key
	End Select		
	Document_onKeyDown	= True	
End Function

Function txtemp_no1_onKeyDown()
Dim CuEvObj,KeyCode
	Set CuEvObj = window.event.srcElement		
	KeyCode = window.event.keycode
	Select Case KeyCode
		Case 13 'enter key
			if emp_no_check() then
				txtemp_no1_onKeyDown	= True
			else 
				txtemp_no1_onKeyDown	= False
			end if
	End Select		
	txtemp_no1_onKeyDown	= True	
End Function

Function emp_no_check()
	dim strEmp_no,IntRetCd,strUID,IntRetCd2
	dim tem_name1,tem_user_id
	
 	strEmp_no = Trim(frm1.txtemp_no1.value)
 	strUID = Trim(frm1.txtuser_id.value)
    IntRetCd = CommonQueryRs(" NAME,res_no "," HAA010T "," EMP_NO =  " & FilterVar(strEmp_no , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
	tem_name1 = Trim(Replace(ConvSPChars(lgF0),Chr(11),""))
	tem_user_id=Trim(Replace(ConvSPChars(lgF1),Chr(11),""))    
	
    If IntRetCd = false then
		Call DisplayMsgbox("800048","X","X","X")
		frm1.txtuser_id.value=""
        frm1.txtname1.value = ""
        frm1.txtemp_no1.focus
        emp_no_check = false

    ELSE    
	    IntRetCd2 = CommonQueryRs(" emp_no "," E11002T "," EMP_NO =  " & FilterVar(strEmp_no , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)  
		if IntRetCd2=true then
			Call DisplayMsgbox("800473","X","X","X")
			frm1.txtuser_id.value=""
			frm1.txtname1.value = ""
			frm1.txtemp_no1.focus			
			emp_no_check = false
		else
			frm1.txtname1.value = tem_name1
			frm1.txtuser_id.value= tem_user_id
			emp_no_check = true
        end if
    END IF
End Function



'========================================================================================================
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
		                            <TD CLASS="TDFAMILY_TITLE" NOWRAP>사번</TD>
		                            <TD CLASS="TDFAMILY2" colspan=3><INPUT CLASS="SINPUTTEST_STYLE" NAME="txtemp_no1" TYPE="Text" MAXLENGTH=13 SiZE=13 tag="22"><IMG SRC="../../../Cshared/Image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="VBScript:Call OpenEmp(frm1.txtemp_no1.value)">
		                                    <INPUT CLASS="SINPUTTEST_STYLE" NAME="txtname1" TYPE="Text" MAXLENGTH=20 SiZE=20 tag="24">
		                            </TD>
                                </TR>
                                <TR>
		                            <TD CLASS="TDFAMILY_TITLE" NOWRAP>사용자ID</TD>
		                            <TD CLASS="TDFAMILY2" colspan=3>
		                                <INPUT CLASS="SINPUTTEST_STYLE" NAME="txtuser_id" TYPE="Text" MAXLENGTH=13 SiZE=13 tag="22">
		                            </TD>
                                </TR>
                                <TR>
		                            <TD CLASS="TDFAMILY_TITLE" NOWRAP>패스워드</TD>
		                            <TD CLASS="TDFAMILY2" colspan=3><INPUT NAME="txtpassword" TYPE="password" MAXLENGTH=10 SiZE=10 tag="22">
		                            </TD>
                                </TR>
                                <TR>
		                            <TD CLASS="TDFAMILY_TITLE" NOWRAP>레벨</TD>
		                            <TD CLASS="TDFAMILY2" colspan=3>
		                                <SELECT NAME="txtpro_auth" STYLE="WIDTH: 100px" TAG="22"><OPTION VALUE=""></OPTION></SELECT>
		                                사용가능한 프로그램의 그룹을 지정합니다.
                                    </TD>      
		                        </TR>
                                <TR>
		                            <TD CLASS="TDFAMILY_TITLE" NOWRAP>자료권한</TD>
		                            <TD CLASS="TDFAMILY2" colspan=3>
		                           	    <INPUT TYPE="RADIO" NAME="txtdept_auth" tag="22" CHECKED ID="txtdept_auth1" VALUE='Y' STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid; BACKGROUND-COLOR: #E9EDF9"><LABEL FOR="txtdept_auth1">사용</LABEL>
    					                <INPUT TYPE="RADIO" NAME="txtdept_auth" tag="22" ID="txtdept_auth2" VALUE='N' STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid; BACKGROUND-COLOR: #E9EDF9"><LABEL FOR="txtdept_auth2">미사용</LABEL>
                                        <INPUT TYPE=HIDDEN NAME="txtdept_authv">
                                    </TD>      
		                        </TR>
                                <TR>
		                            <TD CLASS="TDFAMILY_TITLE" NOWRAP>사용여부</TD>
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

    <TABLE cellSpacing=0 cellPadding=0 width=700 HEIGHT=0 border=0 bgcolor=#ffffff>
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
