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
    Dim dilig_dt
    Dim dilig_cd
	DIM dilig_hour,dilig_min
	dim emp_no
    dilig_dt = Trim(Request("dilig_dt"))
    dilig_cd      = Trim(Request("Dilig_cd"))
    dilig_hour      = Trim(Request("dilig_hour"))
    dilig_min      = Trim(Request("dilig_min"))
	emp_no		 = Trim(Request("emp_no"))
%>

<Script Language="VBScript">
Option Explicit                                                        '��: indicates that All variables must be declared in advance


'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID      = "e1605mb1.asp"						           '��: Biz Logic ASP Name
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


Dim dilig_dt
Dim Dilig_cd
Dim gQuery_YN
dim fDiligAuth,fAuthCheck
Dim StartDate


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
    
	lgKeyStream = lgKeyStream & Trim(fDiligAuth) & gColSep        
    lgKeyStream = lgKeyStream & Trim(fAuthCheck) & gColSep     
    
    lgKeyStream = lgKeyStream & Trim(frm1.txtDilig_cd.value) & gColSep        
    lgKeyStream = lgKeyStream & UniConvDateAToB(Trim(frm1.txtDilig_dt.value),gDateFormat, gServerDateFormat) & gColSep
	if Trim(frm1.txtDilig_hour.value) = "" then
		lgKeyStream = lgKeyStream & "0" & gColSep 
	else 
		lgKeyStream = lgKeyStream & frm1.txtDilig_hour.value & gColSep  
	end if
	
	if Trim(frm1.txtDilig_min.value) = "" then
		lgKeyStream = lgKeyStream & "0" & gColSep        
	else 
		lgKeyStream = lgKeyStream & frm1.txtDilig_min.value & gColSep      
	end if      
    lgKeyStream = lgKeyStream & Trim(parent.txtEmp_no.Value) & gColSep          
'msgbox lgKeyStream              	
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
    Dim iDx,i
    
    For i=0 To 23
    	Call SetCombo(frm1.txtDilig_hour, i, i)
    Next
    
    For i=0 To 59
    	Call SetCombo(frm1.txtDilig_min, i, i)
    Next

	Call CommonQueryRs(" dilig_cd, dilig_nm "," hca010t ", " dilig_type=2 " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
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
    Err.Clear                                                                       '��: Clear err status
    dim dilig_hour,dilig_min,emp_no

	StartDate	= Date()
    dilig_dt = "<%=dilig_dt%>"
    Dilig_cd = "<%=dilig_cd%>"
	dilig_hour = "<%=dilig_hour%>"
	dilig_min = "<%=dilig_min%>"
	emp_no = "<%=emp_no%>"
    lgIntFlgMode = OPMD_CMODE   'insert mode
    gQuery_YN = ""              
    call FncGetDiligAuth(fDiligAuth,fAuthCheck)

    If Replace(fDiligAuth,Chr(11),"") = "" Then
        parent.document.All("nextprev").style.VISIBILITY = "hidden"
    Else
        parent.document.All("nextprev").style.VISIBILITY = "visible"
    End If

	if emp_no  <> "" then
		 frm1.txtEmp_no.Value =emp_no  
	end if

    Call InitComboBox()
    Call LayerShowHide(0)

	Call parent.Click_OpenFrame(Replace(UCase(BIZ_PGM_ID),"MB","MA"))
	if parent.txtName2.value = "" then
		parent.txtEmp_no2.Value = parent.txtemp_no.value 
	end if
    if  dilig_dt <> "" then
        frm1.txtDilig_dt.Value = dilig_dt
        frm1.txtDilig_hour.Value = dilig_hour
        frm1.txtDilig_min.Value = dilig_min
    else 
		frm1.txtDilig_dt.Value = UniConvDateAToB(StartDate,gServerDateFormat,gDateFormat)
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
    frm1.txtAppyn.value = "<%=Trim(Request("txtapp_yn"))%>"
End Sub
'========================================================================================
' Function Name : Window_onUnLoad
' Function Desc : ������ ��ȯ�̳� ȭ���� ���� ��� �����ؾ� �� ���� ó�� 
'========================================================================================
Sub Form_unLoad()
End Sub

Function DbQuery(ppage)

    Dim strVal
    Dim iret
    Err.Clear                                                                    '��: Clear err status

    DbQuery = False                                                              '��: Processing is NG
    'If Grid1.ChkChange() Then Exit Function
    'Call ClearField(document,2)
    If Not chkFieldLength(Document, "1") Then									         '��: This function check required field
		Exit Function
	end if    
    Call LayerShowHide(1)
    
    Call MakeKeyStream("Q")
    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                   '��: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '��: Query Key
    Call RunMyBizASP(MyBizASP, strVal)                                           '��:  Run biz logic

    DbQuery = True                                                               '��: Processing is NG
End Function

'========================================================================================
' Function Name : DbQueryEmp
' Function Desc : 
'========================================================================================
Function DbQueryEmp(ppage)

    Dim strDate
    Dim strVal
    Err.Clear                                                                    '��: Clear err status
    DbQueryEmp = False                                                              '��: Processing is NG
    Call LayerShowHide(1)
    Call MakeKeyStream("Q")
    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                   '��: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '��: Query Key
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call RunMyBizASP(MyBizASP, strVal)                                           '��:  Run biz logic

    DbQueryEmp = True                                                               '��: Processing is NG
End Function


Function DbQueryOk()
    Err.Clear                                                                    '��: Clear err status
  
    If gQuery_YN = "Y" Then                    '�����ڵ� ������ ü������ �Ͼ���.
		lgIntFlgMode = OPMD_CMODE              'create mode
    Else                                       '���ϱ��� ��Ȳ���� update�� �Ϸ��� ��� ������ 
		lgIntFlgMode = OPMD_UMODE              'update mode
		ProtectTag(frm1.txtDilig_cd)
		ProtectTag(frm1.txtDilig_dt)
		
		if frm1.txtAppyn.value = "Y" or frm1.txtAppyn.value = "R" then
			ProtectTag(frm1.txtDilig_hour)
			ProtectTag(frm1.txtDilig_min)		
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
	dim datechk
	On Error Resume Next
    Err.Clear                                                                    '��: Clear err status
	
	DbSave = False	
	'----�ݷ�ó���� ���׿� ���ؼ��� ������ �̷���� �� ����.
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
        if  Date_chk(.txtDilig_dt.value, strDate) = True then
            .txtDilig_dt.value = strDate
        else
            Call DisplayMsgBox("800094","X","X","X")
            .txtDilig_dt.focus()
            exit function
        end if

        if .txtDilig_hour.value = "" then
            Call DisplayMsgBox("800094","X","X","X")
            .txtDilig_hour.focus()
            exit function
        end if
		if Trim(.txtDilig_min.value)="" then
			.txtDilig_min.value = 0
		end if        
        if .txtDilig_cd.value = "" then
            Call DisplayMsgBox("800094","X","X","X")
            .txtDilig_cd.focus()
            exit function
        end if        
        if .txtApp_emp_no.value = "" then
            Call DisplayMsgBox("800094","X","X","X")
            .txtApp_emp_no.focus()
            exit function
        end if		
	 '----����ڿ� �����ڰ� ������ ����.
        if .txtEmp_no.value = .txtApp_emp_no.value  then 
            Call DisplayMsgBox("800476","X","X","X")
		        .txtapp_emp_no.value    = ""
		        .txtapp_name.value      = ""
		        .txtApp_emp_no.focus()
            exit function
        end if
	'-------�����ڰ� �����ϴ� ������� 
	if txtApp_emp_no_Onchange() = False then
		Exit Function
	end if
	
	End With    
	If Not chkFieldLength(Document, "2") Then									         '��: This function check required field
       Exit Function
    End If

    Call MakeKeyStream("C")
	Call LayerShowHide(1)

	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	With Frm1
		.txtMode.value        = "UID_M0002"                                        '��: Save
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '��: Save Key
	End With
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
	
    DbSave  = True                                                               '��: Processing is NG

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
    Err.Clear                                                                    '��: Clear err status
		
	DbDelete = False			                                                 '��: Processing is NG
		
	Call LayerShowHide(1)

	With Frm1
		.txtMode.value        = "UID_M0003"                                        '��: Save
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '��: Save Key
	End With

	'------ Developer Coding part (End )   -------------------------------------------------------------- 

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

	DbDelete = True                                                              '��: Processing is NG
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call FncNew()	
End Function

'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

	lgIntFlgMode = OPMD_CMODE              'create mode

    Call ClearField(document,2)
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    Call LockField(Document)	
	
    Call SetToolBar("01010")

	frm1.txtDilig_dt.value = UniConvDateAToB(Date(),gServerDateFormat,gDateFormat)

    frm1.txtDilig_cd.focus()    
    frm1.txtAppyn.value = ""
    frm1.txtApp_emp_no.value = ""
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    FncNew = True																 '��: Processing is OK
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
'   Event Name : txtApp_emp_no_Onchange()            '<==������ �̸��������� 
'   Event Desc :
'========================================================================================================
Function txtApp_emp_no_Onchange()
    On Error Resume Next
    Err.Clear
    
    Dim iDx
    Dim IntRetCd
    Dim strEmp_no
    
    IF frm1.txtApp_emp_no.value = "" THEN
        frm1.txtApp_name.value = ""
        txtApp_emp_no_Onchange = true
    ELSE
		strEmp_no = frm1.txtApp_emp_no.value 

        IntRetCd = CommonQueryRs(" NAME "," HAA010T "," EMP_NO =  " & FilterVar(frm1.txtApp_emp_no.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
        If IntRetCd = false then
'			Call DisplayMsgbox("800048","X","X","X")	'���޳����ڵ忡 ��ϵ��� ���� �ڵ��Դϴ�.
			Call DisplayMsgBox("970000","X","���","X")

            frm1.txtApp_name.value = ""
            frm1.txtApp_emp_no.focus
        ELSE    
            frm1.txtApp_name.value = Trim(Replace(lgF0,Chr(11),""))   '�����ڵ� 
            txtApp_emp_no_Onchange = true
        END IF
    END IF 
End Function
'========================================================================================================
'   Event Name : CheckLimit()            '<==������ ����üũ 
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
    gQuery_YN = "Y"                                                    ' ������ ü������ �ν��Ѵ�.�������� 
    
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
'   Event Name : txtApp_emp_no_Onchange()            '<==������� �̸��������� 
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
            frm1.txtApp_name.value = Trim(Replace(ConvSPChars(lgF0),Chr(11),""))   '����� �ش��ϴ� �̸� 
            txtApp_emp_no_Onchange = true
        END IF
    END IF 
End Function
</SCRIPT>
<!-- #Include file="../../inc/uniSimsClassID.inc" --> 


<!--
'########################################################################################################
'#						6. TAG ��																		#
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
                        <TD class=base1>���:<INPUT class=base1 NAME="txtEmp_no" MAXLENGTH=13 SiZE=12 TAG=14></TD>
                        <TD class=base1>����:<INPUT class=base1 NAME="txtName" MAXLENGTH=20 SiZE=10 TAG=14></TD>
                        <TD class=base1>����:<INPUT class=base1 NAME="txtroll_pstn" MAXLENGTH=20 SiZE=10 TAG=14></TD>
                        <TD class=base1>�μ�:<INPUT class=base1 NAME="txtDept_nm" MAXLENGTH=25 SiZE=20 TAG=14></TD>
                    </TR>
                    <TR>
                        <TD colspan=4>
                            <TABLE cellSpacing=1 cellPadding=0 width=100% border=0 bgcolor=#ffffff>
                                <TR>
		                            <TD CLASS="TDFAMILY_TITLE" NOWRAP>�ٹ���¥</TD>
		                            <TD CLASS="TDFAMILY2" COLSPAN=3>
										<INPUT CLASS="SINPUTTEST_STYLE" ID="txtDilig_dt" NAME="txtDilig_dt" TYPE="Text" MAXLENGTH=10 SiZE=10 alt="�ٹ���¥" tag="22D" ondblclick="VBScript:Call OpenCalendar('txtDilig_dt',3)">&nbsp;
		                            </TD>
                                </TR>  
                                <TR>
		                            <TD CLASS="TDFAMILY_TITLE" NOWRAP>�ٹ�����</TD>
		                            <TD CLASS="TDFAMILY2" COLSPAN=3><SELECT NAME="txtDilig_cd" ALT="�ٹ�����" STYLE="WIDTH: 150px" TAG="22"><OPTION VALUE=""></OPTION></SELECT>
		                            </TD>
                                </TR>
                                <TR>
		                            <TD CLASS="TDFAMILY_TITLE" NOWRAP>�ٹ��ð�</TD>
		                            <TD CLASS="TDFAMILY2" COLSPAN=3>
										<SELECT NAME="txtDilig_hour" ALT="�ð�" STYLE="WIDTH: 50px" TAG="22"><OPTION VALUE=""></OPTION></SELECT>�ð�
										<SELECT NAME="txtDilig_min" ALT="��" STYLE="WIDTH: 50px" TAG="22"><OPTION VALUE=""></OPTION></SELECT>��
		                            </TD>
                                </TR>                              
  
                                <TR>
		                            <TD CLASS="TDFAMILY_TITLE" NOWRAP>����</TD>
		                            <TD CLASS="TDFAMILY2" COLSPAN=3><INPUT CLASS="SINPUTTEST_STYLE" NAME="txtRemark" ALT="����" TYPE="Text" MAXLENGTH=39 SiZE=40 tag="22">
		                            </TD>
                                </TR>
		                        <TR>
		                            <TD CLASS="TDFAMILY_TITLE" NOWRAP>������</TD>
		                            <TD CLASS="TDFAMILY2" colspan=3><INPUT CLASS="SINPUTTEST_STYLE" NAME="txtApp_emp_no" ALT="���λ��" TYPE="Text" MAXLENGTH=13 SiZE=13 tag="22"><IMG SRC="../../../Cshared/Image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="VBScript:Call OpenEmp(frm1.txtApp_emp_no.value)">
		                                    <INPUT CLASS="SINPUTTEST_STYLE" NAME="txtApp_name" TYPE="Text" MAXLENGTH=20 SiZE=20 tag="24">
		                            </TD>
                                </TR>
		                        <TR height=50>
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
    <INPUT TYPE=HIDDEN NAME="txtAppyn"		 TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtDilig_STRT_dt"		 TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtDilig_END_dt"		 TAG="24">    

</FORM>	

</BODY>
</HTML>
