<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
'*  1. Module Name          : �λ�/�޿����� 
'*  2. Function Name        : ��/�󿩰����������� 
'*  3. Program ID           : h5104oa1
'*  4. Program Name         : ������������� 
'*  5. Program Desc         : ������������� 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2001/05/27
'*  8. Modified date(Last)  : 2003/06/11
'*  9. Modifier (First)     : Shin Kwang-Ho
'* 10. Modifier (Last)      : Lee SiNa
'* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncServer.asp" -->

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
Dim lsInternal_cd

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
    lgIntFlgMode =  Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
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

	Call  ExtractDateFrom("<%=GetsvrDate%>", Parent.gServerDateFormat ,  Parent.gServerDateType ,strYear,strMonth,strDay)
	
	frm1.txtSub_dt.Focus
		
	frm1.txtSub_dt.Year = strYear 		'��� default value setting
	frm1.txtSub_dt.Month = strMonth
	frm1.txtSub_dt.Day = strDay 
	
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
    
    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0040", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    Call  SetCombo2(frm1.cboSub_type,iCodeArr, iNameArr,Chr(11))
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
    Call  ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field

	Call  ggoOper.FormatField(Document, "1", ggStrIntegeralPart,  ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call  ggoOper.FormatDate(frm1.txtSub_dt,  Parent.gDateFormat, 2)		'��� 

    Call InitVariables  
    Call  FuncGetAuth(gStrRequestMenuID,  Parent.gUsrID, lgUsrIntCd)                                ' �ڷ����:lgUsrIntCd ("%", "1%")
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
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    
    FncDelete = True                                                             '��: Processing is OK
End Function

'========================================================================================================
' Name : FncCancel
' Desc : developer describe this line Called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
	On Error Resume Next                                                        '��: Protect system from crashing
End Function

'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                      '��: Protect system from crashing
End Function

'========================================================================================================
' Name : FncExcel
' Desc : developer describe this line Called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
	Call Parent.FncExport( Parent.C_SINGLE)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind( Parent.C_SINGLE, False)
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

'======================================================================================================
'	Name : OpenCode()
'	Description : Code PopUp at vspdData
'=======================================================================================================
Function OpenCode(Byval strCode, Byval iWhere, ByVal Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
		Case "SUB_CD_POP"
	        arrParam(0) = "�����ڵ� �˾�"			        ' �˾� ��Ī 
	    	arrParam(1) = "HDA010T"							    ' TABLE ��Ī 
	    	arrParam(2) = frm1.txtSub_cd.value     				' Code Condition
	    	arrParam(3) = ""'frm1.txtSub_cd_nm.value 					' Name Cindition
	    	arrParam(4) = " PAY_CD = " & FilterVar("*", "''", "S") & "  AND CODE_TYPE = " & FilterVar("2", "''", "S") & ""	' Where Condition
	    	arrParam(5) = "�����ڵ�"  			            ' TextBox ��Ī 
	
	    	arrField(0) = "ALLOW_CD"						   	' Field��(0)
	    	arrField(1) = "ALLOW_NM"    				  		' Field��(1)
	    	arrField(2) = ""    				        		' Field��(2)
    
	    	arrHeader(0) = "�����ڵ�"		    			' Header��(0)
	    	arrHeader(1) = "�����ڵ��"	       		        ' Header��(1)
	    	arrHeader(2) = ""   	    						' Header��(1)		
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtSub_cd.focus
		Exit Function
	Else
		Call SetCode(arrRet, iWhere)
	End If	

End Function

'======================================================================================================
'	Name : SetCode()
'	Description : Code PopUp���� Return�Ǵ� �� setting
'=======================================================================================================
Function SetCode(Byval arrRet, Byval iWhere)

	With frm1

		Select Case iWhere
			Case "SUB_CD_POP"
		        .txtSub_cd.value = arrRet(0) 
		    	.txtSub_cd_nm.value = arrRet(1) 
		    	.txtSub_cd.focus
        End Select

	End With

End Function

'========================================================================================================
' Name : OpenDept
' Desc : �μ� POPUP
'========================================================================================================
Function OpenDept(iWhere)
	Dim arrRet
	Dim arrParam(2)
    Dim rDate
	Dim strBasDt
	Dim strSub_dt

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then
		arrParam(0) = frm1.txtFr_dept_cd.value			            '  Code Condition
	ElseIf iWhere = 1 Then
		arrParam(0) = frm1.txtTo_dept_cd.value			            ' Code Condition
	End If
	

	strBasDt =  UniConvYYYYMMDDToDate( Parent.gDateFormat,frm1.txtSub_dt.Year,Right("0" & frm1.txtSub_dt.Month,2),frm1.txtSub_dt.Day)
	strBasDt =  UNIGetLastDay (strBasDt, Parent.gDateFormat)
	
    arrParam(1) = strBasDt
	arrParam(2) = lgUsrIntCd                    ' �ڷ���� Condition  

	arrRet = window.showModalDialog(HRAskPRAspName("DeptPopupDt"), Array(window.parent,arrParam), _
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
'	Description : Dept Popup���� Return�Ǵ� �� setting
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
	End If
	arrParam(2) = lgUsrIntCd                    ' �ڷ���� Condition  

	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent,arrParam), _
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
'	Description : Item Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Sub SubSetCondEmp(Byval arrRet, Byval iWhere)
	With frm1
		If iWhere = 0 Then 'TextBox(Condition)
			.txtEmp_no.value = arrRet(0)
			.txtName.value = arrRet(1)
			.txtEmp_no.focus
		End If
	End With
End Sub		

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
    Dim rDate
	Dim strYear
	Dim strMonth
	Dim strSub_dt
	Dim strMonthLastDate

	strYear = frm1.txtSub_dt.year
    strMonth = frm1.txtSub_dt.month
    
    If len(strMonth) = 1 then
		strMonth = "0" & strMonth
	End if

	strSub_dt = strYear & strMonth

    rDate =  UNIGetLastDay(frm1.txtSub_dt.Text,  Parent.gDateFormatYYYYMM)
	strMonthLastDate =  UNIConvDate(rDate)				'Client PC Date Format�� DB Format���� ���� 

	Call  FuncGetTermDept(lgUsrIntCd,strMonthLastDate,strMin,strMax)     '�α����� ����� �μ����� �ּ� ,�ִ븦 ������´�.  
    
    If Not chkField(Document, "1") Then									<%'��: This function check indispensable field%>
		Call BtnDisabled(0)

		Exit Function
    End If
    
	if  txtSub_cd_Onchange() then
		Exit Function
    End If    
	if 	 txtEmp_no_Onchange() then
		Exit Function
    End If
	if 	txtTo_dept_cd_Onchange() then
		Exit Function
    End If
	if 	 txtFr_dept_cd_Onchange() then
		Exit Function
    End If

    strFrDept = Trim(frm1.txtFr_internal_cd.value)
    strToDept = Trim(frm1.txtTo_internal_cd.value)
    If strFrDept = "" AND strToDept = "" Then       
    Else
        If strFrDept = "" then
            strFrDept = strMin
        End if
        If strToDept = "" then
            strToDept = strMax
        ElseIf strFrDept > strToDept then
	        Call  DisplayMsgBox("800153","X","X","X")	'���ۺμ��� ����μ����� �۾ƾ� �մϴ�.
            frm1.txtFr_dept_cd.focus
            Set gActiveElement = document.activeElement

			Call BtnDisabled(0)

            Exit Function
        End if 
    End if 
	Call BtnDisabled(1)

	dim sub_type, sub_cd, emp_no, sub_yymm, fr_dept_cd, to_dept_cd 
    Dim ObjName

	StrEbrFile = "h5104oa1"
	
	sub_type = frm1.cboSub_type.value
	sub_cd = frm1.txtSub_cd.value
	emp_no = frm1.txtEmp_no.value
	sub_yymm = strSub_dt
	fr_dept_cd = frm1.txtFr_internal_cd.value
	to_dept_cd = frm1.txtTo_internal_cd.value

	if emp_no = "" then
		emp_no = "%"
		frm1.txtName.value = ""
	End if	

	if fr_dept_cd = "" then
		fr_dept_cd = strMin
		frm1.txtFr_dept_nm.value = ""
	End if	
	
	if to_dept_cd = "" then
		to_dept_cd = strMax
		frm1.txtTo_dept_nm.value = ""
	End if 
    
    lngPos = 0
	
	strUrl = "sub_type|" & sub_type
	strUrl = strUrl & "|sub_cd|" & sub_cd
	strUrl = strUrl & "|emp_no|" & emp_no
	strUrl = strUrl & "|sub_yymm|" & sub_yymm 
	strUrl = strUrl & "|fr_dept_cd|" & fr_dept_cd
	strUrl = strUrl & "|to_dept_cd|" & to_dept_cd
   
	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
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
    Dim rDate
	Dim strYear
	Dim strMonth
	Dim strSub_dt
	Dim strMonthLastDate

	strYear = frm1.txtSub_dt.year
    strMonth = Right("0" & frm1.txtSub_dt.month,2)

	strSub_dt = strYear & strMonth

    rDate =  UNIGetLastDay(frm1.txtSub_dt.Text,  Parent.gDateFormatYYYYMM)
	strMonthLastDate =  UNIConvDate(rDate)				'Client PC Date Format�� DB Format���� ���� 

    Call  FuncGetTermDept(lgUsrIntCd,strMonthLastDate,strMin,strMax)     '�α����� ����� �μ����� �ּ� ,�ִ븦 ������´�.  
    
    If Not chkField(Document, "1") Then									<%'��: This function check indispensable field%>

		Call BtnDisabled(0)
		Exit Function
    End If

	if  txtSub_cd_Onchange() then
		Exit Function
    End If    
	if 	 txtEmp_no_Onchange() then
		Exit Function
    End If
	if 	txtTo_dept_cd_Onchange() then
		Exit Function
    End If
	if 	 txtFr_dept_cd_Onchange() then
		Exit Function
    End If
    	
    strFrDept = Trim(frm1.txtFr_internal_cd.value)
    strToDept = Trim(frm1.txtTo_internal_cd.value)
    If strFrDept = "" AND strToDept = "" Then       
    Else
        If strFrDept = "" then
            strFrDept = strMin
        End if
        If strToDept = "" then
            strToDept = strMax
        ElseIf strFrDept > strToDept then
	        Call  DisplayMsgBox("800153","X","X","X")	'���ۺμ��� ����μ����� �۾ƾ� �մϴ�.
            frm1.txtFr_dept_cd.focus
            Set gActiveElement = document.activeElement
			Call BtnDisabled(0)
            Exit Function
        End if 
    End if 
	Call BtnDisabled(1)
	
	dim strUrl
	dim arrParam, arrField, arrHeader
    Dim StrEbrFile
    Dim ObjName
		
	dim sub_type, sub_cd, emp_no, sub_yymm, fr_dept_cd, to_dept_cd 
	
	StrEbrFile = "h5104oa1"
	
	sub_type = frm1.cboSub_type.value
	sub_cd = frm1.txtSub_cd.value
	emp_no = frm1.txtEmp_no.value
	sub_yymm = strSub_dt
	fr_dept_cd = frm1.txtFr_internal_cd.value
	to_dept_cd = frm1.txtTo_internal_cd.value

	if emp_no = "" then
		emp_no = "%"
		frm1.txtName.value = ""
	End if	

	if fr_dept_cd = "" then
		fr_dept_cd = strMin
		frm1.txtFr_dept_nm.value = ""
	End if	
	
	if to_dept_cd = "" then
		to_dept_cd = strMax
		frm1.txtTo_dept_nm.value = ""
	End if 
					
	strUrl = "sub_type|" & sub_type
	strUrl = strUrl & "|sub_cd|" & sub_cd
	strUrl = strUrl & "|emp_no|" & emp_no
	strUrl = strUrl & "|sub_yymm|" & sub_yymm 
	strUrl = strUrl & "|fr_dept_cd|" & fr_dept_cd
	strUrl = strUrl & "|to_dept_cd|" & to_dept_cd

	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	call FncEBRPreview(ObjName , strUrl)

End Function

Function FncPrint() 
    Call parent.FncPrint()
End Function

Function FncFind() 
	Call parent.FncFind( Parent.C_SINGLE,False)
End Function

'========================================================================================================
'   Event Name : txtSub_cd_Onchange
'   Event Desc :
'========================================================================================================
Function txtSub_cd_Onchange()
    Dim IntRetCd
    If frm1.txtSub_cd.value = "" Then
		frm1.txtSub_cd_nm.value = ""
    Else
        IntRetCd =  CommonQueryRs(" ALLOW_NM "," HDA010T "," ALLOW_CD =  " & FilterVar(frm1.txtSub_cd.value , "''", "S") & " AND PAY_CD = " & FilterVar("*", "''", "S") & "  AND CODE_TYPE = " & FilterVar("2", "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
        If IntRetCd = false then
			Call  DisplayMsgBox("800176","X","X","X")	'��ϵ��� ���� �����ڵ� �Դϴ�.
			 frm1.txtSub_cd_nm.value = ""
             frm1.txtSub_cd.focus
            Set gActiveElement = document.ActiveElement 
            txtSub_cd_Onchange = true
            Exit Function      
        Else
			frm1.txtSub_cd_nm.value = Trim(Replace(lgF0,Chr(11),""))
        End if 
    End if  
    
End Function 

'========================================================================================================
'   Event Name : txtEmp_no_change             '<==�λ縶���Ϳ� �ִ� ������� Ȯ�� 
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
	    IntRetCd =  FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	    if  IntRetCd < 0 then
	        if  IntRetCd = -1 then
    			Call  DisplayMsgBox("800048","X","X","X")	'�ش����� �������� �ʽ��ϴ�.
            else
                Call  DisplayMsgBox("800454","X","X","X")	'�ڷῡ ���� ������ �����ϴ�.
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
    Dim rDate
	Dim strBasDt
	Dim strSub_dt
	Dim strMonthLastDate
        
    strBasDt =  UniConvYYYYMMDDToDate( Parent.gDateFormat,frm1.txtSub_dt.Year,Right("0" & frm1.txtSub_dt.Month,2),frm1.txtSub_dt.Day)
	strBasDt =  UNIGetLastDay (strBasDt, Parent.gDateFormat)

	strMonthLastDate =  UNIConvDate(strBasDt)				'Client PC Date Format�� DB Format���� ���� 

    If RTrim(frm1.txtFr_dept_cd.value) = "" Then
		frm1.txtFr_dept_nm.value = ""
		frm1.txtFr_internal_cd.value = ""
    Else
        IntRetCd =  FuncDeptName(frm1.txtFr_dept_cd.value,strMonthLastDate,lgUsrIntCd,strDept_nm,lsInternal_cd)
        
        if  IntRetCd < 0 then
            if  IntRetCd = -1 then
                Call  DisplayMsgBox("800012", "x","x","x")   '�μ��ڵ������� ��ϵ��� ���� �ڵ��Դϴ�.
            else
                Call  DisplayMsgBox("800455", "x","x","x")   ' �ڷ������ �����ϴ�.
            end if
		    frm1.txtFr_dept_nm.value = ""
		    frm1.txtFr_internal_cd.value = ""
            lsInternal_cd = ""
            frm1.txtFr_dept_cd.focus
            Set gActiveElement = document.ActiveElement
            txtFr_dept_cd_Onchange = true
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
    Dim rDate
	Dim strBasDt
	Dim strSub_dt
	Dim strMonthLastDate
    
	strBasDt =  UniConvYYYYMMDDToDate( Parent.gDateFormat,frm1.txtSub_dt.Year,Right("0" & frm1.txtSub_dt.Month,2),frm1.txtSub_dt.Day)
	strBasDt =  UNIGetLastDay (strBasDt, Parent.gDateFormat)

	strMonthLastDate =  UNIConvDate(strBasDt)				'Client PC Date Format�� DB Format���� ���� 

    If RTrim(frm1.txtTo_dept_cd.value) = "" Then
		frm1.txtTo_dept_nm.value = ""
		frm1.txtTo_internal_cd.value = ""
    Else
        IntRetCd =  FuncDeptName(frm1.txtTo_dept_cd.value,strMonthLastDate,lgUsrIntCd,strDept_nm,lsInternal_cd)
        
        if  IntRetCd < 0 then
            if  IntRetCd = -1 then
                Call  DisplayMsgBox("800012", "x","x","x")   '�μ��ڵ������� ��ϵ��� ���� �ڵ��Դϴ�.
            else
                Call  DisplayMsgBox("800455", "x","x","x")   ' �ڷ������ �����ϴ�.
            end if
		    frm1.txtTo_dept_nm.value = ""
		    frm1.txtTo_internal_cd.value = ""
            lsInternal_cd = ""
            frm1.txtTo_dept_cd.focus
            Set gActiveElement = document.ActiveElement
            txtTo_dept_cd_Onchange = true
            Exit Function      
        else
            frm1.txtTo_dept_nm.value = strDept_nm
            frm1.txtTo_internal_cd.value = lsInternal_cd
        end if
    End if          
    
End Function

'========================================================================================================
' Name : txtSub_dt_DblClick
' Desc : �޷� Popup�� ȣ�� 
'========================================================================================================
Sub txtSub_dt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtSub_dt.Action = 7
		frm1.txtSub_dt.focus	
	End If
End Sub

'=======================================================================================================
'   Event Name : txtFr_year_Keypress(Key)
'   Event Desc : enter key down�ÿ� ��ȸ�Ѵ�.
'=======================================================================================================
Sub txtSub_dt_Keypress(Key)
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTAB"><font color=white>�������������</font></td>
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
								<TD CLASS="TD5" NOWRAP>�ش���</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h5104oa1_txtSub_dt_txtSub_dt.js'></script></TD>
							</TR>
							<TR>	
								<TD CLASS=TD5 NOWRAP>��������</TD>
								<TD CLASS=TD6 NOWRAP><SELECT ID=cboSub_type NAME="cboSub_type" ALT="��������"  STYLE="WIDTH: 100px" TAG="12N"><OPTION VALUE=""></OPTION></SELECT></TD>
							</TR>
							
							<TR>
								<TD CLASS=TD5 NOWRAP>�����ڵ�</TD>
								<TD CLASS=TD6 NOWRAP><INPUT ID=txtSub_cd NAME="txtSub_cd" ALT="�����ڵ�" TYPE="Text" MAXLENGTH="3" SIZE=5  tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBP_CD" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenCode('x', 'SUB_CD_POP', 'x')">
								                     <INPUT ID=txtSub_cd_nm NAME="txtSub_cd_nm" TYPE="Text" MAXLENGTH="20" SIZE=15  tag="14XXXU"></TD>	
							</TR>	
							<TR>
								<TD CLASS="TD5" NOWRAP>�����</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" ID = txtEmp_no NAME="txtEmp_no"  SIZE=13 MAXLENGTH=13 tag="11XXXU" ALT="�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenEmptName (0)">
								                       <INPUT TYPE="Text" ID=txtName NAME="txtName" SIZE=20 MAXLENGTH=30  tag="14XXXU" ALT="������ڵ��"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>�μ��ڵ�</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" ID = txtFr_dept_cd NAME="txtFr_dept_cd" SIZE=10 MAXLENGTH=10  tag="11XXXU" ALT="���ۺμ��ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(0)">
								                       <INPUT TYPE="Text" ID=txtFr_dept_nm NAME="txtFr_dept_nm" SIZE=20 MAXLENGTH=40  tag="14XXXU" ALT="�μ��ڵ��">&nbsp;~
								                       <INPUT ID=txtFr_internal_cd NAME="txtFr_internal_cd" ALT="���κμ��ڵ�" TYPE="hidden" SiZE=7 MAXLENGTH=30  tag="14XXXU"></TD>
							</TR>			
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" ID = "txtTo_dept_cd" NAME="txtTo_dept_cd" SIZE=10 MAXLENGTH=10  tag="11XXXU" ALT="����μ��ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(1)">
								                       <INPUT TYPE="Text" ID=txtTo_dept_nm NAME="txtTo_dept_nm" SIZE=20 MAXLENGTH=40  tag="14XXXU" ALT="�μ��ڵ��">
								                       <INPUT ID=txtTo_internal_cd NAME="txtTo_internal_cd" ALT="���κμ��ڵ�" TYPE="hidden" SiZE=7 MAXLENGTH=30  tag="14XXXU"></TD>
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
                         <BUTTON NAME="btnRun"   CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>�̸�����</BUTTON>&nbsp;
                         <BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>�μ�</BUTTON>
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


