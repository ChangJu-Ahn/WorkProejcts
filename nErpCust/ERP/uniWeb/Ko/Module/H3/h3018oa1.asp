<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
'*  1. Module Name          : �λ�/�޿����� 
'*  2. Function Name        : �ٹ��̷°��� 
'*  3. Program ID           : h3018oa1
'*  4. Program Name         : ���������� 
'*  5. Program Desc         : ���������� 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2001/05/27
'*  8. Modified date(Last)  : 2003/06/10
'*  9. Modifier (First)     : Shin Kwang-Ho/mok yong bin
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

<SCRIPT LANGUAGE="VBsCRIPT"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

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

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
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
	
	Call ExtractDateFrom("<%=GetsvrDate%>", parent.gServerDateFormat , parent.gServerDateType ,strYear,strMonth,strDay)
	
	frm1.txtFr_year.Year = strYear 		 '����� default value setting
	frm1.txtTo_year.Year = strYear 		 '����� default value setting

End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "H","NOCOOKIE","OA") %>
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.FormatDate(frm1.txtFr_year, parent.gDateFormat, 3)	
	Call ggoOper.FormatDate(frm1.txtTo_year, parent.gDateFormat, 3)	
    Call InitVariables
    Call FuncGetAuth(gStrRequestMenuID, parent.gUsrID, lgUsrIntCd)                                ' �ڷ����:lgUsrIntCd ("%", "1%")
    Call SetDefaultVal
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
    
    FncQuery = False															 '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    
    
    If txtFr_emp_no_Onchange() Then        'enter key �� ��ȸ�� ����� check�� �ش���� ������ query����...
        Exit Function
    End if
    
    If txtTo_emp_no_Onchange() Then        'enter key �� ��ȸ�� ����� check�� �ش���� ������ query����...
        Exit Function
    End if

    If txtFr_dept_cd_Onchange() Then        'enter key �� ��ȸ�� �μ��ڵ带 check�� �ش���� ������ query����...
        Exit Function
    End if
    
    If txtTo_dept_cd_Onchange() Then        'enter key �� ��ȸ�� �μ��ڵ带 check�� �ش���� ������ query����...
        Exit Function
    End if

    FncQuery = True                                                              '��: Processing is OK

End Function
	
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
' Name : FncExcel
' Desc : developer describe this line Called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
	Call Parent.FncExport(parent.C_SINGLE)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind(parent.C_SINGLE, False)
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
		arrParam(0) = frm1.txtFr_emp_no.value			' Code Condition
	    arrParam(1) = ""'frm1.txtFr_name.value			' Name Cindition
	End If
	arrParam(2) = lgUsrIntCd                    ' �ڷ���� Condition  

	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtFr_emp_no.focus
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
			.txtFr_emp_no.value = arrRet(0)
			.txtFr_name.value = arrRet(1)
			.txtFr_emp_no.focus
		End If
	End With
End Sub


'========================================================================================================
' Name : OpenEmptName2()
' Desc : developer describe this line 
'========================================================================================================

Function OpenEmptName2(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then 'TextBox(Condition)
		arrParam(0) = frm1.txtTo_emp_no.value			' Code Condition
	    arrParam(1) = ""'frm1.txtTo_name.value			' Name Cindition
	End If
	arrParam(2) = lgUsrIntCd                    ' �ڷ���� Condition  

	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtTo_emp_no.focus
		Exit Function
	Else
		Call SubSetCondEmp2(arrRet, iWhere)
	End If	
			
End Function

'======================================================================================================
'	Name : SubSetCondEmp2()
'	Description : Item Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Sub SubSetCondEmp2(Byval arrRet, Byval iWhere)
	With frm1
		If iWhere = 0 Then 'TextBox(Condition)
			.txtTo_emp_no.value = arrRet(0)
			.txtTo_name.value = arrRet(1)
			.txtTo_emp_no.focus
		End If
	End With
End Sub
		
'========================================================================================================
' Name : OpenDept
' Desc : �μ� POPUP
'========================================================================================================
Function OpenDept(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then
		arrParam(0) = frm1.txtFr_dept_cd.value			            '  Code Condition
	ElseIf iWhere = 1 Then
		arrParam(0) = frm1.txtTo_dept_cd.value			            ' Code Condition
	End If
    arrParam(1) = ""
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

'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================%>
Function FncBtnPrint() 

	Dim strUrl
	Dim StrEbrFile
    Dim strFrDept
    Dim strToDept
    Dim strtxtFrYear
    Dim strtxtToYear
    Dim strFrEmpNo
    Dim strToEmpNo
    Dim strMin
    Dim strMax
	dim ObjName
	   
    Call FuncGetTermDept(lgUsrIntCd,"",strMin,strMax)
    
    
    If Not chkField(Document, "1") Then									<%'��: This function check indispensable field%>
	
	   Exit Function
    End If
    
    strtxtFrYear = Trim(frm1.txtFr_year.Text)
    strtxtToYear = Trim(frm1.txtTo_year.Text)
    
    If strtxtFrYear = "" AND strtxtToYear = "" Then       
    Else
        If strtxtFrYear = "" then
            strtxtFrYear = "1950"
        End if
        If strtxtToYear = "" then
            strtxtToYear = "2100"
        ElseIf Cdbl(strtxtFrYear) > Cdbl(strtxtToYear) then
	        Call DisplayMsgBox("800055","X","X","X")	'���۳⵵�� ����⵵���� �۾ƾ� �մϴ�.
            frm1.txtFr_year.focus
            Set gActiveElement = document.activeElement
			Exit Function
        End if 
    End if 
    
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
	        Call DisplayMsgBox("800153","X","X","X")	'���ۺμ��� ����μ����� �۾ƾ� �մϴ�.
            frm1.txtFr_dept_cd.focus
            Set gActiveElement = document.activeElement
	        Exit Function
        End if 
    End if 
    
    strFrEmpNo = Trim(frm1.txtFr_emp_no.value)
    strToEmpNo = Trim(frm1.txtTo_emp_no.value)
    If strFrEmpNo = "" AND strToEmpNo = "" Then       
    Else
        If strFrEmpNo = "" then
            strFrEmpNo =  "0"
        End if
        If strToEmpNo = "" then
            strToEmpNo = "ZZZZZZZZZZZZZ"
        ElseIf strFrEmpNo > strToEmpNo then
	        Call DisplayMsgBox("800056","X","X","X")	'���ۻ���� ���������� �۾ƾ� �մϴ�.
            frm1.txtFr_emp_no.focus
            Set gActiveElement = document.activeElement
	        Exit Function
        End if 
    End if 

	dim stand_dt, fr_year, to_year, fr_emp_no, to_emp_no, fr_dept_cd, to_dept_cd
	
	StrEbrFile = "h3018oa1"
	
	stand_dt = UniConvYYYYMMDDToDate(parent.gAPDateFormat, frm1.txtTo_year.Year, "12", "31")
	fr_year = frm1.txtFr_year.Text	 
	to_year = frm1.txtTo_year.Text	 
	fr_emp_no = frm1.txtFr_emp_no.value
	to_emp_no = frm1.txtTo_emp_no.value
	fr_dept_cd = frm1.txtFr_internal_cd.value
	to_dept_cd = frm1.txtTo_internal_cd.value

	if fr_emp_no = "" then
		fr_emp_no = "0"
		frm1.txtFr_name.value = ""
	End if	

	if to_emp_no = "" then
		to_emp_no = "ZZZZZZZZZZZZZ"
		frm1.txtTo_name.value = ""
	End if	

	if fr_dept_cd = "" then
		fr_dept_cd = strMin
		frm1.txtFr_dept_nm.value = ""
	End if	
	
	if to_dept_cd = "" then
		to_dept_cd = strMax
		frm1.txtTo_dept_nm.value = ""
	End if		

	strUrl =          "stand_dt|"    & stand_dt
	strUrl = strUrl & "|fr_year|"    & fr_year
	strUrl = strUrl & "|to_year|"    & to_year	
	strUrl = strUrl & "|fr_emp_no|" & fr_emp_no
	strUrl = strUrl & "|to_emp_no|" & to_emp_no
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
    Dim strtxtFrYear
    Dim strtxtToYear
    Dim strFrEmpNo
    Dim strToEmpNo
    Dim strMin
    Dim strMax

	dim strUrl
	dim arrParam, arrField, arrHeader
    Dim StrEbrFile
	dim ObjName		
	dim stand_dt, fr_year, to_year, fr_emp_no, to_emp_no, fr_dept_cd, to_dept_cd
   
    Call FuncGetTermDept(lgUsrIntCd,"",strMin,strMax)

    If Not chkField(Document, "1") Then									<%'��: This function check indispensable field%>
	
	   Exit Function
    End If
    
    strtxtFrYear = Trim(frm1.txtFr_year.Text)
    strtxtToYear = Trim(frm1.txtTo_year.Text)
    If strtxtFrYear = "" AND strtxtToYear = "" Then       
    Else
        If strtxtFrYear = "" then
            strtxtFrYear = "1950"
        End if
        If strtxtToYear = "" then
            strtxtToYear = "2100"
        ElseIf Cdbl(strtxtFrYear) > Cdbl(strtxtToYear) then
	        Call DisplayMsgBox("800055","X","X","X")	'���۳⵵�� ����⵵���� �۾ƾ� �մϴ�.
            frm1.txtFr_year.focus
            Set gActiveElement = document.activeElement
			Exit Function
        End if 
    End if 
    
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
	        Call DisplayMsgBox("800153","X","X","X")	'���ۺμ��� ����μ����� �۾ƾ� �մϴ�.
            frm1.txtFr_dept_cd.focus
            Set gActiveElement = document.activeElement
			
			Exit Function
        End if 
    End if 
        
    strFrEmpNo = Trim(frm1.txtFr_emp_no.value)
    strToEmpNo = Trim(frm1.txtTo_emp_no.value)
    If strFrEmpNo = "" AND strToEmpNo = "" Then       
    Else
        If strFrEmpNo = "" then
            strFrEmpNo = "0"
        End if
        If strToEmpNo = "" then
            strToEmpNo = "ZZZZZZZZZZZZZ"
        ElseIf strFrEmpNo > strToEmpNo then
	        Call DisplayMsgBox("800056","X","X","X")	'���ۻ���� ���������� �۾ƾ� �մϴ�.
            frm1.txtFr_emp_no.focus
            Set gActiveElement = document.activeElement
			Exit Function
        End if 
    End if 	
	
	StrEbrFile = "h3018oa1"
		
	stand_dt = UniConvYYYYMMDDToDate(parent.gAPDateFormat, frm1.txtTo_year.Year, "12", "31")
	fr_year = frm1.txtFr_year.Text	 
	to_year = frm1.txtTo_year.Text	 
	fr_emp_no = frm1.txtFr_emp_no.value
	to_emp_no = frm1.txtTo_emp_no.value
	fr_dept_cd = frm1.txtFr_internal_cd.value
	to_dept_cd = frm1.txtTo_internal_cd.value

	if fr_emp_no = "" then
		fr_emp_no = "0"
		frm1.txtFr_name.value = ""
	End if	

	if to_emp_no = "" then
		to_emp_no = "ZZZZZZZZZZZZZ"
		frm1.txtTo_name.value = ""
	End if	

	if fr_dept_cd = "" then
		fr_dept_cd = strMin
		frm1.txtFr_dept_nm.value = ""
	End if	
	
	if to_dept_cd = "" then
		to_dept_cd = strMax
		frm1.txtTo_dept_nm.value = ""
	End if
	
	strUrl =          "stand_dt|"    & stand_dt
	strUrl = strUrl & "|fr_year|"    & fr_year
	strUrl = strUrl & "|to_year|"    & to_year	
	strUrl = strUrl & "|fr_emp_no|" & fr_emp_no
	strUrl = strUrl & "|to_emp_no|" & to_emp_no
	strUrl = strUrl & "|fr_dept_cd|" & fr_dept_cd
	strUrl = strUrl & "|to_dept_cd|" & to_dept_cd

	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	call FncEBRPreview(ObjName , strUrl)

End Function

'========================================================================================================
'   Event Name : txtFr_emp_no_change             '<==�λ縶���Ϳ� �ִ� ������� Ȯ�� 
'   Event Desc :
'========================================================================================================
Function txtFr_emp_no_Onchange()
    
    Dim IntRetCd
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd
    Dim strVal
    
    If frm1.txtFr_emp_no.value = "" Then
		frm1.txtFr_name.value = ""
    Else
	    IntRetCd = FuncGetEmpInf2(frm1.txtFr_emp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	    if  IntRetCd < 0 then
	        if  IntRetCd = -1 then
    			Call DisplayMsgBox("800048","X","X","X")	'�ش����� �������� �ʽ��ϴ�.
            else
                Call DisplayMsgBox("800454","X","X","X")	'�ڷῡ ���� ������ �����ϴ�.
            end if
			frm1.txtFr_name.value = ""
            frm1.txtFr_emp_no.focus
            Set gActiveElement = document.ActiveElement
            txtFr_emp_no_Onchange = true
            Exit Function      
        Else
            frm1.txtFr_name.value = strName
        End if 
    End if  
    
End Function 

'========================================================================================================
'   Event Name : txtFr_emp_no_change             '<==�λ縶���Ϳ� �ִ� ������� Ȯ�� 
'   Event Desc :
'========================================================================================================
Function txtTo_emp_no_Onchange()
    
    Dim IntRetCd
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd
    Dim strVal
    
    If frm1.txtTo_emp_no.value = "" Then
		frm1.txtTo_name.value = ""
    Else
	    IntRetCd = FuncGetEmpInf2(frm1.txtTo_emp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	    if  IntRetCd < 0 then
	        if  IntRetCd = -1 then
    			Call DisplayMsgBox("800048","X","X","X")	'�ش����� �������� �ʽ��ϴ�.
            else
                Call DisplayMsgBox("800454","X","X","X")	'�ڷῡ ���� ������ �����ϴ�.
            end if
			frm1.txtTo_name.value = ""
            frm1.txtTo_emp_no.focus
            Set gActiveElement = document.ActiveElement
            txtTo_emp_no_Onchange = true
            Exit Function      
        Else
            frm1.txtTo_name.value = strName
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

    If frm1.txtFr_dept_cd.value = "" Then
		frm1.txtFr_dept_nm.value = ""
		frm1.txtFr_internal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtFr_dept_cd.value,"",lgUsrIntCd,strDept_nm,lsInternal_cd)
        
        if  IntRetCd < 0 then
            if  IntRetCd = -1 then
                Call DisplayMsgBox("800012", "x","x","x")   '�μ��ڵ������� ��ϵ��� ���� �ڵ��Դϴ�.
            else
                Call DisplayMsgBox("800455", "x","x","x")   ' �ڷ������ �����ϴ�.
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
    
    If frm1.txtTo_dept_cd.value = "" Then
		frm1.txtTo_dept_nm.value = ""
		frm1.txtTo_internal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtTo_dept_cd.value,"",lgUsrIntCd,strDept_nm,lsInternal_cd)
        
        if  IntRetCd < 0 then
            if  IntRetCd = -1 then
                Call DisplayMsgBox("800012", "x","x","x")   '�μ��ڵ������� ��ϵ��� ���� �ڵ��Դϴ�.
            else
                Call DisplayMsgBox("800455", "x","x","x")   ' �ڷ������ �����ϴ�.
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
' Name : txtFr_year_DblClick(Button)
' Desc : �޷� Popup�� ȣ�� 
'========================================================================================================
Sub txtFr_year_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M") 	
		frm1.txtFr_year.Action = 7
		frm1.txtFr_year.focus
	End If
End Sub

Sub txtTo_year_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M") 	
		frm1.txtTo_year.Action = 7
		frm1.txtTo_year.focus
	End If
End Sub

'=======================================================================================================
'   Event Name : txtFr_year_Keypress(Key)
'   Event Desc : enter key down�ÿ� ��ȸ�Ѵ�.
'=======================================================================================================
Sub txtFr_year_Keypress(Key)
    If Key = 13 Then
        MainQuery()
    End If
End Sub

Sub txtTo_year_Keypress(Key)
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTAB"><font color=white>����������</font></td>
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
								<TD CLASS=TD5  NOWRAP>��³⵵</TD>
								<TD CLASS=TD6  NOWRAP><script language =javascript src='./js/h3018oa1_fpDateTime1_txtFr_year.js'></script>&nbsp;����</TD>
							</TR>
							<TR>
								<TD CLASS=TD5  NOWRAP></TD>
								<TD CLASS=TD6  NOWRAP><script language =javascript src='./js/h3018oa1_fpDateTime2_txtTo_year.js'></script>&nbsp;����</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>�μ��ڵ�</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" ID = "txtFr_dept_cd" NAME="txtFr_dept_cd" SIZE=10 MAXLENGTH=10  tag="11XXXU" ALT="���ۺμ��ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(0)">
								                       <INPUT TYPE="Text" NAME="txtFr_dept_nm" SIZE=20 MAXLENGTH=40  tag="14XXXU" ALT="�μ��ڵ��">&nbsp;����&nbsp;
								                       <INPUT NAME="txtFr_internal_cd" ALT="���κμ��ڵ�" TYPE="hidden" SiZE=7 MAXLENGTH=30  tag="14XXXU"></TD>
							</TR>			
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" ID = "txtTo_dept_cd" NAME="txtTo_dept_cd" SIZE=10 MAXLENGTH=10  tag="11XXXU" ALT="����μ��ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(1)">
								                       <INPUT TYPE="Text" NAME="txtTo_dept_nm" SIZE=20 MAXLENGTH=40  tag="14XXXU" ALT="�μ��ڵ��">&nbsp;����
								                       <INPUT NAME="txtTo_internal_cd" ALT="���κμ��ڵ�" TYPE="hidden" SiZE=7 MAXLENGTH=30  tag="14XXXU"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>�����</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" ID = "txtFr_emp_no" NAME="txtFr_emp_no"  SIZE=13 MAXLENGTH=13 tag="11XXXU" ALT="�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenEmptName(0)">
								                       <INPUT TYPE="Text" NAME="txtFr_name" SIZE=20 MAXLENGTH=30  tag="14XXXU" ALT="���۴�����ڵ��">&nbsp;����</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" ID = "txtTo_emp_no" NAME="txtTo_emp_no"  SIZE=13 MAXLENGTH=13 tag="11XXXU" ALT="�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenEmptName2(0)">
								                       <INPUT TYPE="Text" NAME="txtTo_name" SIZE=20 MAXLENGTH=30  tag="14XXXU" ALT="���������ڵ��">&nbsp;����</TD>
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


