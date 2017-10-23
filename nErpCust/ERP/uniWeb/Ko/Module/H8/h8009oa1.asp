<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
'*  1. Module Name          : 인사/급여관리 
'*  2. Function Name        : 소급분관리 
'*  3. Program ID           : h8009oa1
'*  4. Program Name         : 소급급상여대장출력 
'*  5. Program Desc         : 소급급상여대장출력 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2001/05/27
'*  8. Modified date(Last)  : 2003/06/13
'*  9. Modifier (First)     : Shin Kwang-Ho
'* 10. Modifier (Last)      : Lee SiNa
'* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop          
Dim lsInternal_cd

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	Dim strYear
	Dim strMonth
	Dim strDay

	Call ExtractDateFrom("<%=GetsvrDate%>",Parent.gServerDateFormat , Parent.gServerDateType ,strYear,strMonth,strDay)
	
    frm1.txtFr_pay_yymm.focus 
		
	frm1.txtFr_pay_yymm.Year = strYear 		'년월 default value setting
	frm1.txtFr_pay_yymm.Month = strMonth 

	frm1.txtTo_pay_yymm.Year = strYear 		'년월 default value setting
	frm1.txtTo_pay_yymm.Month = strMonth 

End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "H", "NOCOOKIE", "OA") %>
End Sub


'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx
    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0003", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    Call SetCombo2(frm1.cboOcpt_type,iCodeArr, iNameArr,Chr(11))
    
    Call CommonQueryRs("MINOR_CD, MINOR_NM ","B_MINOR","MAJOR_CD = " & FilterVar("H0005", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    Call SetCombo2(frm1.cboPay_cd, iCodeArr, iNameArr, Chr(11))
    
    Call CommonQueryRs("MINOR_CD, MINOR_NM ","B_MINOR","MAJOR_CD = " & FilterVar("H0035", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    Call SetCombo2(frm1.cboSect_cd, iCodeArr, iNameArr, Chr(11))
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.FormatDate(frm1.txtFr_pay_yymm, Parent.gDateFormat, 2)
	Call ggoOper.FormatDate(frm1.txtTo_pay_yymm, Parent.gDateFormat, 2)
    Call InitVariables
    Call FuncGetAuth(gStrRequestMenuID, Parent.gUsrID, lgUsrIntCd)                 ' 자료권한:lgUsrIntCd ("%", "1%")
    Call SetDefaultVal
    Call InitComboBox
    Call SetToolbar("1000000000000111")
        
End Sub
'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub


'========================================================================================================
' Name : FncQuery
' Desc : developer describe this line Called by MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If txtEmp_no_Onchange() Then        'enter key 로 조회시 사원을 check후 해당사항 없으면 query종료...
        Exit Function
    End if
    
    If txtFr_dept_cd_Onchange()  Then        'enter key 로 조회시 부서코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if
    
    If txtTo_dept_cd_Onchange()  Then        'enter key 로 조회시 부서코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if
    
    FncQuery = True                                                              '☜: Processing is OK

End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    FncDelete = True                                                             '☜: Processing is OK
End Function
'========================================================================================================
' Name : FncCancel
' Desc : developer describe this line Called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
	On Error Resume Next                                                        '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                      '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncExcel
' Desc : developer describe this line Called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
	Call Parent.FncExport(Parent.C_SINGLE)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind(Parent.C_SINGLE, False)
End Function

'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

	FncExit = False

	FncExit = True
End Function

'========================================================================================================
' Name : OpenEmptName()
' Desc : developer describe this line 
'========================================================================================================

Function OpenEmptName(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then 'TextBox(Condition)
		arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	    arrParam(1) = ""'frm1.txtName.value			' Name Cindition
	Else 'spread

	End If
	arrParam(2) = lgUsrIntCd                    ' 자료권한 Condition  

	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtEmp_no.focus
		Exit Function
	Else
		Call SubSetCondEmp(arrRet, iWhere)
	End If	
			
End Function

'======================================================================================================
'	Name : SubSetCondEmp()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SubSetCondEmp(Byval arrRet, Byval iWhere)
	With frm1
		If iWhere = 0 Then
			.txtEmp_no.value = arrRet(0)
			.txtName.value = arrRet(1)
			.txtEmp_no.focus
		End If
	End With
End Sub

'========================================================================================================
' Name : OpenDept
' Desc : 부서 POPUP
'========================================================================================================
Function OpenDept(iWhere)
	Dim arrRet
	Dim arrParam(2)
    Dim strBasDt 
    
	strBasDt = UNIGetLastDay(frm1.txtFr_pay_yymm.Text,Parent.gDateFormatYYYYMM)
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then
		arrParam(0) = frm1.txtFr_dept_cd.value			            '  Code Condition
	ElseIf iWhere = 1 Then
		arrParam(0) = frm1.txtTo_dept_cd.value			            ' Code Condition
	End If
	
    arrParam(1) = strBasDt
	arrParam(2) = lgUsrIntCd                    ' 자료권한 Condition  

	arrRet = window.showModalDialog(HRAskPRAspName("DeptPopupDt"), Array(window.parent, arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
		     Case "0"
               frm1.txtFr_dept_cd.focus
             Case "1"  
               frm1.txtTo_dept_cd.focus
        End Select	
		Exit Function
	Else
		Call SetDept(arrRet, iWhere)
	End If	
			
End Function

'------------------------------------------  SetDept()  ------------------------------------------------
'	Name : SetDept()
'	Description : Dept Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetDept(Byval arrRet, Byval iWhere)
		
	With frm1
		Select Case iWhere
		     Case "0"
               .txtFr_dept_cd.value = arrRet(0)
               .txtFr_dept_nm.value = arrRet(1)
               .txtFr_internal_cd.value = arrRet(2)
               .txtFr_dept_cd.focus
             Case "1"  
               .txtTo_dept_cd.value = arrRet(0)
               .txtTo_dept_nm.value = arrRet(1) 
               .txtTo_internal_cd.value = arrRet(2) 
               .txtTo_dept_cd.focus
        End Select
	End With
End Function       		

'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================%>

Function FncBtnPrint() 

	Dim strUrl
	Dim lngPos
	Dim intCnt
    Dim StrEbrFile
    Dim strFrDept
    Dim strToDept
    Dim strMin
    Dim strMax
    Dim strBasDt

	dim fr_pay_yymm, to_pay_yymm, emp_no, pay_cd, sect_cd, ocpt_type, fr_dept_cd, to_dept_cd 
	Dim ObjName

	strBasDt = UNIGetLastDay(frm1.txtFr_pay_yymm.Text,Parent.gDateFormatYYYYMM)
    strBasDt = UNIConvDateCompanyToDB(strBasDt, Parent.gDateFormat)

    Call FuncGetTermDept(lgUsrIntCd,strBasDt,strMin,strMax)     '로그인한 사원의 부서권한 최소 ,최대를 가지고온다.  

    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
		Call BtnDisabled(0)

		Exit Function
    End If
 
    If txtEmp_no_Onchange() Then        'enter key 로 조회시 사원을 check후 해당사항 없으면 query종료...
        Exit Function
    End if
    
    If txtFr_dept_cd_Onchange()  Then        'enter key 로 조회시 부서코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if
    
    If txtTo_dept_cd_Onchange()  Then        'enter key 로 조회시 부서코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if
      
    strFrDept = Trim(frm1.txtFr_internal_cd.value)
    strToDept = Trim(frm1.txtTo_internal_cd.value)
    If strFrDept = "" AND strToDept ="" Then       
    Else
        If strFrDept = "" then
            strFrDept = strMin
        End if
        If strToDept = "" then
            strToDept = strMax
        ElseIf strFrDept > strToDept then
	        Call DisplayMsgbox("800153","X","X","X")	'시작부서는 종료부서보다 작아야 합니다.
            frm1.txtFr_dept_cd.focus
            Set gActiveElement = document.activeElement
			
			Call BtnDisabled(0)

            Exit Function
        End if 
    End if 
    
    If CompareDateByFormat(frm1.txtFr_pay_yymm,frm1.txtTo_pay_yymm,frm1.txtFr_pay_yymm.Alt,frm1.txtTo_pay_yymm.Alt,"970025",Parent.gDateFormatYYYYMM,Parent.gComDateType,True) = False Then
        frm1.txtFr_pay_yymm.focus
        Set gActiveElement = document.activeElement

        Call BtnDisabled(0)
        Exit Function
    End if 
	Call BtnDisabled(1)
	
	StrEbrFile = "h8009oa1"
	
    fr_pay_yymm = frm1.txtFr_pay_yymm.year & Right("0" & frm1.txtFr_pay_yymm.month , 2)	
    to_pay_yymm = frm1.txtTo_pay_yymm.year & Right("0" & frm1.txtTo_pay_yymm.month , 2)	
	emp_no = frm1.txtEmp_no.value
	ocpt_type = frm1.cboOcpt_type.value
	pay_cd = frm1.cboPay_cd.value
	sect_cd = frm1.cboSect_cd.value
	fr_dept_cd = frm1.txtFr_internal_cd.value
	to_dept_cd = frm1.txtTo_internal_cd.value
	
	if emp_no = "" then
		emp_no = "%"
		frm1.txtName.value = ""
	End if	

	if ocpt_type = "" then
		ocpt_type = "%"
	End if	

	if pay_cd = "" then
		pay_cd = "%"
	End if	
	
	if sect_cd = "" then
		sect_cd = "%"
	End if	
	
	if fr_dept_cd = "" then
		fr_dept_cd = strMin
		frm1.txtFr_dept_nm.value = ""
	End if	
	
	if to_dept_cd = "" then
		to_dept_cd = strMax
		frm1.txtTo_dept_nm.value = ""
	End if 
	
	strUrl = "fr_pay_yymm|" & fr_pay_yymm
	strUrl = strUrl & "|to_pay_yymm|" & to_pay_yymm
	strUrl = strUrl & "|emp_no|" & emp_no
	strUrl = strUrl & "|pay_cd|" & pay_cd
	strUrl = strUrl & "|sect_cd|" & sect_cd
	strUrl = strUrl & "|ocpt_type|" & ocpt_type 
	strUrl = strUrl & "|fr_dept_cd|" & fr_dept_cd
	strUrl = strUrl & "|to_dept_cd|" & to_dept_cd
	strUrl = strUrl & "|dt|" & strBasDt

    ObjName = AskEBDocumentName(StrEbrFile, "ebr")
   
 	call FncEBRPrint(EBAction , ObjName , strUrl)

End Function

'========================================================================================
' Function Name : FncBtnPreview()
' Function Desc : This function is related to Preview Button
'========================================================================================

Function FncBtnPreview()

    Dim strFrDept
    Dim strToDept
    Dim strMin
    Dim strMax
    Dim strBasDt
    Dim ObjName
	
	dim strUrl
	dim arrParam, arrField, arrHeader
    Dim StrEbrFile
		
	dim fr_pay_yymm, to_pay_yymm, emp_no, pay_cd, sect_cd, ocpt_type, fr_dept_cd, to_dept_cd 

	strBasDt = UNIGetLastDay(frm1.txtFr_pay_yymm.Text,Parent.gDateFormatYYYYMM)
    strBasDt = UNIConvDateCompanyToDB(strBasDt, Parent.gDateFormat)

    Call FuncGetTermDept(lgUsrIntCd,strBasDt,strMin,strMax)     '로그인한 사원의 부서권한 최소 ,최대를 가지고온다.  

    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
		Call BtnDisabled(0)

		Exit Function
    End If
    If txtEmp_no_Onchange() Then        'enter key 로 조회시 사원을 check후 해당사항 없으면 query종료...
        Exit Function
    End if
    
    If txtFr_dept_cd_Onchange()  Then        'enter key 로 조회시 부서코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if
    
    If txtTo_dept_cd_Onchange()  Then        'enter key 로 조회시 부서코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if
    
    strFrDept = Trim(frm1.txtFr_internal_cd.value)
    strToDept = Trim(frm1.txtTo_internal_cd.value)
    If strFrDept = "" AND strToDept ="" Then       
    Else
        If strFrDept = "" then
            strFrDept = strMin
        End if
        If strToDept = "" then
            strToDept = strMax
        ElseIf strFrDept > strToDept then
	        Call DisplayMsgbox("800153","X","X","X")	'시작부서는 종료부서보다 작아야 합니다.
            frm1.txtFr_dept_cd.focus
            Set gActiveElement = document.activeElement

			Call BtnDisabled(0)

            Exit Function
        End if 
    End if 
    
    If CompareDateByFormat(frm1.txtFr_pay_yymm,frm1.txtTo_pay_yymm,frm1.txtFr_pay_yymm.Alt,frm1.txtTo_pay_yymm.Alt,"970025",Parent.gDateFormatYYYYMM,Parent.gComDateType,True) = False Then
        frm1.txtFr_pay_yymm.focus
        Set gActiveElement = document.activeElement

        Call BtnDisabled(0)
        Exit Function
    End if 
	Call BtnDisabled(1)
	
	StrEbrFile = "h8009oa1"
	
    fr_pay_yymm = frm1.txtFr_pay_yymm.year & Right("0" & frm1.txtFr_pay_yymm.month , 2)	
    to_pay_yymm = frm1.txtTo_pay_yymm.year & Right("0" & frm1.txtTo_pay_yymm.month , 2)	
	
	emp_no = frm1.txtEmp_no.value
	ocpt_type = frm1.cboOcpt_type.value
	pay_cd = frm1.cboPay_cd.value
	sect_cd = frm1.cboSect_cd.value
	fr_dept_cd = frm1.txtFr_internal_cd.value
	to_dept_cd = frm1.txtTo_internal_cd.value

	if emp_no = "" then
		emp_no = "%"
		frm1.txtName.value = ""
	End if	

	if ocpt_type = "" then
		ocpt_type = "%"
	End if	

	if pay_cd = "" then
		pay_cd = "%"
	End if	
	
	if sect_cd = "" then
		sect_cd = "%"
	End if	
	
	if fr_dept_cd = "" then
		fr_dept_cd = strMin
		frm1.txtFr_dept_nm.value = ""
	End if	
	
	if to_dept_cd = "" then
		to_dept_cd = strMax
		frm1.txtTo_dept_nm.value = ""
	End if 
					
	strUrl = "fr_pay_yymm|" & fr_pay_yymm
	strUrl = strUrl & "|to_pay_yymm|" & to_pay_yymm
	strUrl = strUrl & "|emp_no|" & emp_no
	strUrl = strUrl & "|pay_cd|" & pay_cd
	strUrl = strUrl & "|sect_cd|" & sect_cd
	strUrl = strUrl & "|ocpt_type|" & ocpt_type 
	strUrl = strUrl & "|fr_dept_cd|" & fr_dept_cd
	strUrl = strUrl & "|to_dept_cd|" & to_dept_cd
        strUrl = strUrl & "|dt|" & strBasDt

    ObjName = AskEBDocumentName(StrEbrFile, "ebr")

	call FncEBRPreview(ObjName , strUrl)

End Function

Function FncPrint() 
    Call parent.FncPrint()
End Function

Function FncFind() 
	Call parent.FncFind(Parent.C_SINGLE,False)
End Function

'========================================================================================================
'   Event Name : txtEmp_no_change             '<==인사마스터에 있는 사원인지 확인 
'   Event Desc :
'========================================================================================================
Function txtEmp_no_Onchange()
    Dim IntRetCd
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd
    Dim strVal
    
    If frm1.txtEmp_no.value = "" Then
		frm1.txtName.value = ""
    Else
	    IntRetCd = FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	    if  IntRetCd < 0 then
	        if  IntRetCd = -1 then
    			Call DisplayMsgbox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            else
                Call DisplayMsgbox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            end if
			frm1.txtName.value = ""
            frm1.txtEmp_no.focus
            Set gActiveElement = document.ActiveElement
            txtEmp_no_Onchange = true
            Exit Function      
        Else
            frm1.txtName.value = strName
        End if 
    End if  
    
End Function 

'========================================================================================================
'   Event Name : txtFr_dept_cd_change
'   Event Desc :
'========================================================================================================
Function txtFr_dept_cd_Onchange()
  
    Dim IntRetCd
    Dim strDept_nm
    Dim strBasDt 

	strBasDt = UNIGetLastDay(frm1.txtFr_pay_yymm.Text, Parent.gDateFormatYYYYMM)            '해당년월의 마지막 날을 가지고 온다.

    If frm1.txtFr_dept_cd.value = "" Then
		frm1.txtFr_dept_nm.value = ""
		frm1.txtFr_internal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtFr_dept_cd.value,UNIConvDate(strBasDt),lgUsrIntCd,strDept_nm,lsInternal_cd)
        
        if  IntRetCd < 0 then
            if  IntRetCd = -1 then
                Call DisplayMsgBox("800012", "x","x","x")   '부서코드정보에 등록되지 않은 코드입니다.
            else
                Call DisplayMsgBox("800455", "x","x","x")   ' 자료권한이 없습니다.
            end if
		    frm1.txtFr_dept_nm.value = ""
		    frm1.txtFr_internal_cd.value = ""
            lsInternal_cd = ""
            frm1.txtFr_dept_cd.focus
            Set gActiveElement = document.ActiveElement
            txtFr_dept_cd_Onchange = True
            Exit Function      
        else
            frm1.txtFr_dept_nm.value = strDept_nm
            frm1.txtFr_internal_cd.value = lsInternal_cd
        end if
        
    End if
End Function

'========================================================================================================
'   Event Name : txtTo_dept_cd_Onchange
'   Event Desc :
'========================================================================================================
Function txtTo_dept_cd_Onchange()
    Dim IntRetCd
    Dim strDept_nm
    Dim strBasDt 

	strBasDt = UNIGetLastDay(frm1.txtFr_pay_yymm.Text, Parent.gDateFormatYYYYMM)            '해당년월의 마지막 날을 가지고 온다.

    If frm1.txtTo_dept_cd.value = "" Then
		frm1.txtTo_dept_nm.value = ""
		frm1.txtTo_internal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtTo_dept_cd.value,UNIConvDate(strBasDt),lgUsrIntCd,strDept_nm,lsInternal_cd)
        
        if  IntRetCd < 0 then
            if  IntRetCd = -1 then
                Call DisplayMsgBox("800012", "x","x","x")   '부서코드정보에 등록되지 않은 코드입니다.
            else
                Call DisplayMsgBox("800455", "x","x","x")   ' 자료권한이 없습니다.
            end if
		    frm1.txtTo_dept_nm.value = ""
		    frm1.txtTo_internal_cd.value = ""
            lsInternal_cd = ""
            frm1.txtTo_dept_cd.focus
            Set gActiveElement = document.ActiveElement
            txtTo_dept_cd_Onchange = True
            Exit Function      
        else
            frm1.txtTo_dept_nm.value = strDept_nm
            frm1.txtTo_internal_cd.value = lsInternal_cd
        end if
    End if          
    
End Function

'========================================================================================================
' Name : txtBas_dt_DblClick
' Desc : 달력 Popup을 호출 
'========================================================================================================
Sub txtFr_pay_yymm_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtFr_pay_yymm.Action = 7
		frm1.txtFr_pay_yymm.focus
	End If
End Sub

Sub txtTo_pay_yymm_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtTo_pay_yymm.Action = 7
		frm1.txtTo_pay_yymm.focus
	End If
End Sub

'=======================================================================================================
'   Event Name : txtFr_year_Keypress(Key)
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub txtFr_pay_yymm_Keypress(Key)
    If Key = 13 Then
        MainQuery()
    End If
End Sub

Sub txtTo_pay_yymm_Keypress(Key)
    If Key = 13 Then
        MainQuery()
    End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	

<SCRIPT LANGUAGE="JavaScript">
<!-- Hide script from old browsers
function setCookie(name, value, expire)
{
	document.cookie = name + "=" + escape(value)
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
	document.cookie = name + "=" + escape(value)
		+ "; path=/EasyBaseWeb/bin"
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
	document.cookie = name + "=" + escape(value)
		+ "; path=/EasyBaseWeb/lib"
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
}

setCookie("client", "-1", null)
setCookie("owner", "admin", null)
setCookie("identity", "admin", null)
-->
</SCRIPT>

</HEAD>

<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTAB"><font color=white>소급급/상여대장출력</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* HEIGHT="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT=20>
						<TABLE <%=LR_SPACE_TYPE_60%>>   
							<TR>
								<TD CLASS="TD5" NOWRAP>해당년월</TD>
								<TD CLASS=TD6 NOWRAP><OBJECT classid=<%=gCLSIDFPDT%> ID=txtFr_pay_yymm NAME="txtFr_pay_yymm" CLASS=FPDTYYYYMM tag="12X1" ALT="시작년월" Title="FPDATETIME"></OBJECT>&nbsp;~&nbsp;
													 <OBJECT classid=<%=gCLSIDFPDT%> ID=txtTo_pay_yymm NAME="txtTo_pay_yymm" CLASS=FPDTYYYYMM tag="12X1" ALT="종료년월" Title="FPDATETIME"></OBJECT></td>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>직종</TD>
								<TD CLASS=TD6 NOWRAP><SELECT ID=cboOcpt_type NAME="cboOcpt_type" ALT="직종구분" STYLE="WIDTH: 120px" tag="11"><OPTION Value=""></OPTION></SELECT></TD>
								
							</TR>

							<TR>
								<TD CLASS="TD5" NOWRAP>급여구분</TD>
	                        	<TD CLASS="TD6" NOWRAP><SELECT ID=cboPay_cd NAME="cboPay_cd" ALT="급여구분" STYLE="WIDTH: 120px" tag="11"><OPTION Value=""></OPTION></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>대상자</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" ID=txtEmp_no NAME="txtEmp_no" SIZE=13 MAXLENGTH=13 tag="11XXXU" ALT="대상자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenEmptName (0)">
								                       <INPUT TYPE="Text" ID=txtName NAME="txtName" SIZE=20 MAXLENGTH=30 tag="14XXXU" ALT="대상자코드명"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>부서코드</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" ID=txtFr_dept_cd NAME="txtFr_dept_cd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="시작부서코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(0)">
								                       <INPUT TYPE="Text" ID=txtFr_dept_nm NAME="txtFr_dept_nm" SIZE=20 MAXLENGTH=40 tag="14XXXU" ALT="부서코드명">&nbsp;~&nbsp;
								                       <INPUT ID=txtFr_internal_cd NAME="txtFr_internal_cd" ALT="내부부서코드" TYPE="hidden" SiZE=7 MAXLENGTH=30 tag="14XXXU"></TD>
							</TR>			
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" ID=txtTo_dept_cd NAME="txtTo_dept_cd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="종료부서코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(1)">
								                       <INPUT TYPE="Text" ID=txtTo_dept_nm NAME="txtTo_dept_nm" SIZE=20 MAXLENGTH=40 tag="14XXXU" ALT="부서코드명">&nbsp;
								                       <INPUT ID=txtTo_internal_cd NAME="txtTo_internal_cd" ALT="내부부서코드" TYPE="hidden" SiZE=7 MAXLENGTH=30 tag="14XXXU"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>근무구역</TD>
	                        	<TD CLASS="TD6" NOWRAP><SELECT ID=cboSect_cd Name="cboSect_cd" ALT="근무구역" STYLE="WIDTH: 120px" tag="11"><OPTION Value=""></OPTION></SELECT></TD>
							</TR>
    					</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD>
		    <TABLE <%=LR_SPACE_TYPE_30%>>
		        <TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
                         <BUTTON NAME="btnRun"   CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;
                         <BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>인쇄</BUTTON>
		            </TD>
					<TD WIDTH=* ALIGN="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME="EBAction" TARGET = "MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname">
	<INPUT TYPE="HIDDEN" NAME="dbname">
	<INPUT TYPE="HIDDEN" NAME="filename">
	<INPUT TYPE="HIDDEN" NAME="condvar">
	<INPUT TYPE="HIDDEN" NAME="date">
</FORM>
</BODY>
</HTML>


