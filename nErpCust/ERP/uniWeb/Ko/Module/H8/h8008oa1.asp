<%@ LANGUAGE="VBSCRIPT" %>
<!--
'======================================================================================================
'*  1. Module Name          : 인사/급여관리 
'*  2. Function Name        : 소급분관리 
'*  3. Program ID           : h8008oa1
'*  4. Program Name         : 월별소급급/상여출력 
'*  5. Program Desc         : 월별소급급/상여출력 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2001/05/27
'*  8. Modified date(Last)  : 2003/06/13
'*  9. Modifier (First)     : Shin Kwang-Ho
'* 10. Modifier (Last)      : Lee SiNa
'* 11. Comment              :
'=======================================================================================================-->
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
Dim lgOldRow

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
	
	frm1.txtPay_yymm.Focus
		
	frm1.txtPay_yymm.Year = strYear 		'년월 default value setting
	frm1.txtPay_yymm.Month = strMonth 

End Sub
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx
    
    ' 직종 
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0003", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr =  lgF0
    iNameArr =  lgF1
    Call SetCombo2(frm1.cboOcpt_type, iCodeArr, iNameArr, Chr(11))
    
    Call CommonQueryRs("MINOR_CD, MINOR_NM ","B_MINOR","MAJOR_CD = " & FilterVar("H0005", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    Call SetCombo2(frm1.cboPay_cd, iCodeArr, iNameArr, Chr(11))
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
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.FormatDate(frm1.txtPay_yymm, Parent.gDateFormat, 2)
   
    Call InitVariables 
    Call InitComboBox
    
    Call FuncGetAuth(gStrRequestMenuID , Parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%") 
    
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
' Name : OpenEmpName()
' Desc : developer describe this line(grid외에서 사용) 
'========================================================================================================
Function OpenEmpName(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

    If  iWhere = 0 Then
	    arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	    arrParam(1) = ""'frm1.txtName.value			' Name Cindition
    Else
	    arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	    arrParam(1) = ""'frm1.txtName.value		' Name Cindition
	End If
    
    arrParam(1) = ""
    arrParam(2) = lgUsrIntcd
    
	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtEmp_no.focus
		Exit Function
	Else
		Call SetEmpName(arrRet)
	End If	
			
End Function

'======================================================================================================
'	Name : SetEmpName()
'	Description : Item Popup에서 Return되는 값 setting(grid외에서 사용)
'=======================================================================================================
Sub SetEmpName(arrRet)
	With frm1
		.txtEmp_no.value = arrRet(0)
		.txtName.value = arrRet(1)
		Call ggoOper.ClearField(Document, "2")					 '☜: Clear Contents  Field
		Set gActiveElement = document.ActiveElement
		.txtEmp_no.focus
		lgBlnFlgChgValue = False
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
    
	strBasDt = UNIGetLastDay(frm1.txtPay_yymm.Text,Parent.gDateFormatYYYYMM)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then
		arrParam(0) = frm1.txtFr_dept_cd.value			            '  Code Condition
	ElseIf iWhere = 1 Then
		arrParam(0) = frm1.txtTo_dept_cd.value			            ' Code Condition
	End If
	
    arrParam(1) = strBasDt
	arrParam(2) = lgUsrIntcd
	
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
'	Name : OpenCode()
'	Description : Code PopUp at vspdData
'=======================================================================================================
Function OpenCode(Byval strCode, Byval iWhere, ByVal Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
	    Case "SECT_CD_POP"
	        arrParam(0) = "근무구역 팝업"			        ' 팝업 명칭 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = frm1.txtSect_cd.value     			' Code Condition
	    	arrParam(3) = ""'frm1.txtSect_cd_nm.value  				' Name Cindition
	    	arrParam(4) = "major_cd = " & FilterVar("H0035", "''", "S") & ""	               	' Where Condition
	    	arrParam(5) = "근무구역코드"  			            ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"						   	' Field명(0)
	    	arrField(1) = "minor_nm"    				  		' Field명(1)
	    	arrField(2) = ""    				        		' Field명(2)
    
	    	arrHeader(0) = "근무구역코드"	   		    			' Header명(0)
	    	arrHeader(1) = "근무구역코드명"	          		        ' Header명(1)
	    	arrHeader(2) = ""           	    						' Header명(1)

		Case "OCPT_TYPE_POP"
	        arrParam(0) = "직종코드 팝업"			        ' 팝업 명칭 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = frm1.cboOcpt_type.value     			' Code Condition
	    	arrParam(3) = ""'frm1.cboOcpt_type.value				' Name Cindition
	    	arrParam(4) = "major_cd = " & FilterVar("H0003", "''", "S") & ""	               	' Where Condition
	    	arrParam(5) = "{직종코드}}"  			            ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"						   	' Field명(0)
	    	arrField(1) = "minor_nm"    				  		' Field명(1)
	    	arrField(2) = ""    				        		' Field명(2)
    
	    	arrHeader(0) = "직종코드"	   		    			' Header명(0)
	    	arrHeader(1) = "직종코드명"	          		        ' Header명(1)
	    	arrHeader(2) = ""	             						' Header명(1)		
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
		    Case "SECT_CD_POP"
		    	frm1.txtSect_cd.focus
			Case "SECT_CD_POP"
		    	frm1.cboOcpt_type.focus
        End Select	
		Exit Function
	Else
		Call SetCode(arrRet, iWhere)
	End If	

End Function

'======================================================================================================
'	Name : SetCode()
'	Description : Code PopUp에서 Return되는 값 setting
'=======================================================================================================
Function SetCode(Byval arrRet, Byval iWhere)

	With frm1

		Select Case iWhere
		    Case "SECT_CD_POP"
		        .txtSect_cd.value = arrRet(0) 
		    	.txtSect_cd_nm.value = arrRet(1) 
		    	.txtSect_cd.focus
        
			Case "SECT_CD_POP"
		        .cboOcpt_type.value = arrRet(0) 
		    	.cboOcpt_type_nm.value = arrRet(1) 
		    	.cboOcpt_type.focus
        End Select

	End With

End Function

'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================
Function FncBtnPrint() 
	Dim strUrl
	Dim lngPos
	Dim intCnt
    Dim StrEbrFile
    Dim ObjName
    	
    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
       Exit Function
    End If

	Call BtnDisabled(1)

	Dim pay_yymm, emp_no, pay_cd, sect_cd, ocpt_type, fr_dept_cd, to_dept_cd , rFrDept ,rToDept ,IntRetCd
	
	StrEbrFile = "h8008oa1"
	
    pay_yymm = frm1.txtPay_yymm.year & Right("0" & frm1.txtPay_yymm.month , 2)	
	emp_no = frm1.txtEmp_no.value
	ocpt_type = frm1.cboOcpt_type.value
	pay_cd = frm1.cboPay_cd.value
	sect_cd = frm1.txtSect_cd.value
	fr_dept_cd = frm1.txtFr_internal_cd.value
	to_dept_cd = frm1.txtTo_internal_cd.value

	if emp_no = "" then
		emp_no = "%"
		frm1.txtName.value = ""
	End if	

	if ocpt_type = "" then
		ocpt_type = "%"
		frm1.cboOcpt_type.value = ""
	End if	

	if pay_cd = "" then
		pay_cd = "%"
		frm1.cbopay_cd.value = ""
	End if	
	
	if sect_cd = "" then
		sect_cd = "%"
		frm1.txtSect_cd_nm.value = ""
	End if	
	
    If  txtEmp_no_Onchange() then
		Call BtnDisabled(0)
		frm1.txtEmp_no.focus
        Exit Function
    End If
    If  txtSect_cd_OnChange()  then
		Call BtnDisabled(0)
		frm1.txtSect_cd.focus
        Exit Function
    End If
    If  txtFr_Dept_cd_Onchange() then
		Call BtnDisabled(0)
		frm1.txtFr_Dept_cd.focus
        Exit Function
    End If
    If  txtTo_Dept_cd_Onchange()  then
		Call BtnDisabled(0)
		frm1.txtTo_dept_cd.focus
        Exit Function
    End If
    
    Fr_dept_cd = frm1.txtFr_internal_cd.value
    To_dept_cd = frm1.txtTo_internal_cd.value
    
    If fr_dept_cd = "" then
        IntRetCd = FuncGetTermDept(lgUsrIntCd ,"", rFrDept ,rToDept)
		frm1.txtFr_internal_cd.value = rFrDept
		frm1.txtFr_dept_nm.value = ""
	End If	
	
	If to_dept_cd = "" then
        IntRetCd = FuncGetTermDept(lgUsrIntCd ,"", rFrDept ,rToDept)
		frm1.txtTo_internal_cd.value = rToDept
		frm1.txtTo_dept_nm.value = ""
	End If  
    
    Fr_dept_cd = frm1.txtFr_internal_cd.value
    To_dept_cd = frm1.txtTo_internal_cd.value
    
    If (Fr_dept_cd<> "") AND (To_dept_cd<>"") Then       
        If Fr_dept_cd > To_dept_cd then
	        Call DisplayMsgbox("800153","X","X","X")	'시작부서코드는 종료부서코드보다 작아야합니다.
            frm1.txtFr_dept_cd.focus
            Set gActiveElement = document.activeElement

			Call BtnDisabled(0)

            Exit Function
        End IF 
    END IF   
	
	strUrl = "pay_yymm|" & pay_yymm
	strUrl = strUrl & "|emp_no|" & emp_no
	strUrl = strUrl & "|pay_cd|" & pay_cd
	strUrl = strUrl & "|sect_cd|" & sect_cd
	strUrl = strUrl & "|ocpt_type|" & ocpt_type 
	strUrl = strUrl & "|fr_dept_cd|" & fr_dept_cd
	strUrl = strUrl & "|to_dept_cd|" & to_dept_cd

    ObjName = AskEBDocumentName(StrEbrFile, "ebr")
   
 	call FncEBRPrint(EBAction , ObjName , strUrl)

End Function
'========================================================================================
' Function Name : FncBtnPreview()
' Function Desc : This function is related to Preview Button
'========================================================================================
Function FncBtnPreview()
	
	Dim strUrl
	Dim arrParam, arrField, arrHeader
    Dim StrEbrFile, ObjName
		
	Dim pay_yymm, emp_no, pay_cd, sect_cd, ocpt_type, fr_dept_cd, to_dept_cd , rFrDept ,rToDept ,IntRetCd
    	
    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
       Exit Function
    End If
	
	StrEbrFile = "h8008oa1"
	
    pay_yymm = frm1.txtPay_yymm.year & Right("0" & frm1.txtPay_yymm.month , 2)	
	emp_no = frm1.txtEmp_no.value
	ocpt_type = frm1.cboOcpt_type.value
	pay_cd = frm1.cboPay_cd.value
	sect_cd = frm1.txtSect_cd.value
	fr_dept_cd = frm1.txtFr_internal_cd.value
	to_dept_cd = frm1.txtTo_internal_cd.value

	if emp_no = "" then
		emp_no = "%"
		frm1.txtName.value = ""
	End if	

	if ocpt_type = "" then
		ocpt_type = "%"
		frm1.cboOcpt_type.value = ""
	End if	

	if pay_cd = "" then
		pay_cd = "%"
		frm1.cbopay_cd.value = ""
	End if	
	
	if sect_cd = "" then
		sect_cd = "%"
		frm1.txtSect_cd_nm.value = ""
	End if	
	
    If  txtEmp_no_Onchange() then
		Call BtnDisabled(0)
		frm1.txtEmp_no.focus
        Exit Function
    End If
    If  txtSect_cd_OnChange()  then
		Call BtnDisabled(0)
		frm1.txtSect_cd.focus
        Exit Function
    End If
    If  txtFr_Dept_cd_Onchange() then
		Call BtnDisabled(0)
		frm1.txtFr_Dept_cd.focus
        Exit Function
    End If
    If  txtTo_Dept_cd_Onchange()  then
		Call BtnDisabled(0)
		frm1.txtTo_dept_cd.focus
        Exit Function
    End If
    
    Fr_dept_cd = frm1.txtFr_internal_cd.value
    To_dept_cd = frm1.txtTo_internal_cd.value
    
    If fr_dept_cd = "" then
        IntRetCd = FuncGetTermDept(lgUsrIntCd ,"", rFrDept ,rToDept)
		frm1.txtFr_internal_cd.value = rFrDept
		frm1.txtFr_dept_nm.value = ""
	End If	
	
	If to_dept_cd = "" then
        IntRetCd = FuncGetTermDept(lgUsrIntCd ,"", rFrDept ,rToDept)
		frm1.txtTo_internal_cd.value = rToDept
		frm1.txtTo_dept_nm.value = ""
	End If  
    
    Fr_dept_cd = frm1.txtFr_internal_cd.value
    To_dept_cd = frm1.txtTo_internal_cd.value
    
    If (Fr_dept_cd<> "") AND (To_dept_cd<>"") Then       
        If Fr_dept_cd > To_dept_cd then
	        Call DisplayMsgbox("800153","X","X","X")	'시작부서코드는 종료부서코드보다 작아야합니다.
            frm1.txtFr_dept_cd.focus
            Set gActiveElement = document.activeElement
            Exit Function
        End IF 
    END IF   
				
	strUrl = "pay_yymm|" & pay_yymm
	strUrl = strUrl & "|emp_no|" & emp_no
	strUrl = strUrl & "|pay_cd|" & pay_cd
	strUrl = strUrl & "|sect_cd|" & sect_cd
	strUrl = strUrl & "|ocpt_type|" & ocpt_type 
	strUrl = strUrl & "|fr_dept_cd|" & fr_dept_cd
	strUrl = strUrl & "|to_dept_cd|" & to_dept_cd

    ObjName = AskEBDocumentName(StrEbrFile, "ebr")

	call FncEBRPreview(ObjName , strUrl)

End Function
'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                      '☜: Protect system from crashing
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
	FncExit = True
End Function


'========================================================================================================
'   Event Name : txtEmp_no_Onchange           
'   Event Desc :
'========================================================================================================s
Function txtEmp_no_Onchange()
    Dim IntRetCd
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd

    If  frm1.txtEmp_no.value = "" Then
		frm1.txtName.value = ""
    Else
	    IntRetCd = FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	                
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgbox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            Else
                Call DisplayMsgbox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
			frm1.txtName.value = ""
            Frm1.txtEmp_no.focus 
            Set gActiveElement = document.ActiveElement
			txtEmp_no_Onchange = true
        Else
			frm1.txtName.value = strName
        End if 
    End if  
End Function

'======================================================================================================
'   Event Name : txtSect_cd_OnChange
'   Event Desc : 근무구역코드 에러체크 
'=======================================================================================================
Function txtSect_cd_OnChange()
        Dim iDx
        Dim IntRetCd
        
        If frm1.txtSect_cd.value = "" Then
            frm1.txtSect_cd_nm.value = ""
        ELSE
            IntRetCd = CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0035", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtSect_cd.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
            IF IntRetCd = False THEN
				Call DisplayMsgBox("970000","X","근무구역코드","X")
    	        frm1.txtSect_cd_nm.value = ""
    	        frm1.txtSect_cd.focus
    	        txtSect_cd_OnChange = true
				Exit Function     	        
    	    Else
    	        frm1.txtSect_cd_nm.value = Trim(Replace(lgF0, Chr(11), ""))
    	        frm1.txtSect_cd.focus
    	    End If
        End If
    	    
End Function
'========================================================================================================
'   Event Name : txtFr_dept_cd_Onchange
'   Event Desc :
'========================================================================================================
Function txtFr_dept_cd_Onchange()
    Dim IntRetCd,Dept_Nm,Internal_cd
    Dim rDate
	
	rDate = UNIGetLastDay(frm1.txtPay_yymm.Text, Parent.gDateFormatYYYYMM)            '해당년월의 마지막 날을 가지고 온다.

    If frm1.txtFr_dept_cd.value = "" Then
		frm1.txtFr_dept_nm.value = ""
		frm1.txtFr_internal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtFr_dept_cd.value , UNIConvDate(rDate), lgUsrIntCd,Dept_Nm, Internal_cd)
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgbox("800098","X","X","X")	'부서정보코드에 등록되지 않은 코드입니다.
            Else
                Call DisplayMsgbox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
		    frm1.txtFr_dept_nm.value = ""
		    frm1.txtFr_internal_cd.value = ""
            frm1.txtFr_dept_cd.focus
            Set gActiveElement = document.ActiveElement 
            txtFr_dept_cd_Onchange = true
            Exit Function      
        Else
			frm1.txtFr_dept_nm.value = Dept_Nm
		    frm1.txtFr_internal_cd.value = Internal_cd
        End if 
    End if  
    
End Function

'========================================================================================================
'   Event Name : txtTo_dept_cd_Onchange
'   Event Desc :
'========================================================================================================
Function txtTo_dept_cd_Onchange()
    Dim IntRetCd,Dept_Nm,Internal_cd
    Dim rDate
	
	rDate = UNIGetLastDay(frm1.txtPay_yymm.Text, Parent.gDateFormatYYYYMM)            '해당년월의 마지막 날을 가지고 온다.
	
    If frm1.txtTo_dept_cd.value = "" Then
		frm1.txtTo_dept_nm.value = ""
		frm1.txtTo_internal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtTo_dept_cd.value , UNIConvDate(rDate), lgUsrIntCd,Dept_Nm , Internal_cd)
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgbox("800098","X","X","X")	'부서정보코드에 등록되지 않은 코드입니다.
            Else
                Call DisplayMsgbox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
		    frm1.txtTo_dept_nm.value = ""
		    frm1.txtTo_internal_cd.value = ""
            frm1.txtTo_dept_cd.focus
            Set gActiveElement = document.ActiveElement 
            txtTo_dept_cd_Onchange = true
            Exit Function      
        Else
			frm1.txtTo_dept_nm.value = Dept_Nm
		    frm1.txtTo_internal_cd.value = Internal_cd
        End if 
    End if  
    
End Function

'=======================================================================================================
'   Event Name : txt________Keypress
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub txtEmp_no_Keypress(Key)
    If Key = 13 Then
        MainQuery()
    End If
End Sub     

Sub txtProv_Type_Keypress(Key)
    If Key = 13 Then
        MainQuery()
    End If
End Sub     

Sub txtFr_dept_cd_Keypress(Key)
    If Key = 13 Then
        MainQuery()
    End If
End Sub

Sub txtto_dept_cd_Keypress(Key)
    If Key = 13 Then
        MainQuery()
    End If
End Sub
'========================================================================================================
' Name : txtPay_yymm_DblClick
' Desc : 달력 Popup을 호출 
'========================================================================================================
Sub txtPay_yymm_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtPay_yymm.Action = 7
		frm1.txtPay_yymm.focus
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
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTAB"><font color=white>월별소급급/상여출력</font></td>
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
								<TD CLASS=TD5 NOWRAP>해당년월</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h8008oa1_txtPay_yymm_txtPay_yymm.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>직종</TD>
						  		<TD CLASS=TD6 NOWRAP><SELECT ID=cboOcpt_type NAME="cboOcpt_type" ALT="직종" STYLE="WIDTH: 100px" TAG="11"><OPTION VALUE=""></OPTION></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>급여구분</TD>
	                        	<TD CLASS="TD6" NOWRAP><SELECT ID=cboPay_cd NAME="cboPay_cd" ALT="급여구분" STYLE="WIDTH: 100px" tag="11"><OPTION Value=""></OPTION></SELECT></TD>
							</TR>			
							<TR>
								<TD CLASS=TD5 NOWRAP>대상자</TD>
								<TD CLASS=TD6 NOWRAP><INPUT ID=txtEmp_no NAME="txtEmp_no" ALT="대상자" TYPE="Text" MAXLENGTH="13" SIZE=13 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBP_CD" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenEmpName (1)">
								                     <INPUT ID=txtName NAME="txtName" TYPE="Text" MAXLENGTH="30" SIZE=20 tag="14XXXU"></TD>	
						    </TR>
							<TR>			
							    <TD CLASS=TD5 NOWRAP>부서</TD>
								<TD CLASS=TD6 NOWRAP><INPUT ID=txtFr_dept_cd NAME="txtFr_dept_cd" ALT="부서코드" TYPE="Text" SiZE=10 MAXLENGTH=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(0)">
			                                         <INPUT ID=txtFr_dept_nm NAME="txtFr_dept_nm" ALT="부서코드명" TYPE="Text" SiZE=20 MAXLENGTH=40 tag="14XXXU">&nbsp;~
		                                             <INPUT ID=txtFr_internal_cd NAME="txtFr_internal_cd" ALT="내부부서코드" TYPE="HIDDEN" SiZE=7 MAXLENGTH=7 tag="14XXXU">
		                    </TR>
		                    <TR>    
		                        <TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP><INPUT ID=txtto_dept_cd NAME="txtto_dept_cd" ALT="부서코드" TYPE="Text" SIZE=10 MAXLENGTH="10"  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnname" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(1)">
							                         <INPUT ID=txtto_dept_nm NAME="txtto_dept_nm" ALT="부서코드명" TYPE="Text" SIZE=20 MAXLENGTH="40"  tag="14XXXU">
    			                                     <INPUT ID=txtTo_internal_cd NAME="txtTo_internal_cd" ALT="내부부서코드" TYPE="HIDDEN" SiZE=7 MAXLENGTH=7 tag="14XXXU"></TD>
    			            </TR> 
							<TR>
								<TD CLASS=TD5 NOWRAP>근무구역</TD>
								<TD CLASS=TD6 NOWRAP><INPUT ID=txtSect_cd NAME="txtSect_cd" ALT="근무구역" TYPE="Text" SiZE=10 MAXLENGTH=7 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenCode('x', 'SECT_CD_POP', 'x')">
								                     <INPUT ID=txtSect_cd_nm NAME="txtSect_cd_nm" ALT="근무구역" TYPE="Text" SiZE=20 MAXLENGTH=20 tag="14XXXU"></td>
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
		                <BUTTON NAME="btnPreview" CLASS="CLSSBTN" onclick="VBScript:FncBtnPreview()">미리보기</BUTTON>&nbsp;
		                <BUTTON NAME="btnPrint"   CLASS="CLSSBTN" OnClick="VBScript:FncBtnPrint()">인쇄</BUTTON>

		            </TD>
					<TD WIDTH=* ALIGN="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=20><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
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



