<%@ LANGUAGE="VBSCRIPT" %>
<!--======================================================================================================
*  1. Module Name          : 인사/급여관리 
*  2. Function Name        : 급여관리 
*  3. Program ID           : h6020ma1
*  4. Program Name         : 은행이체파일생성 
*  5. Program Desc         : 은행이체파일생성 
*  6. Comproxy List        : +
*  7. Modified date(First) : 2001/05/27
*  8. Modified date(Last)  : 2003/06/13
*  9. Modifier (First)     : Shin Kwang-Ho
* 10. Modifier (Last)      : Lee SiNa
* 11. Comment              :
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

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit
'==============================================================================================
'=                       4.3 Common variables 
'==============================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->				           '☆: Biz Logic ASP Name

'==============================================================================================
'							1.2.3 Global Variable값 정의  
'==============================================================================================
Const BIZ_PGM_ID      = "h6020mb1.asp"		

Dim IsOpenPop
Dim lgOldRow    
'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
End Sub
'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()
	Dim strYear
	Dim strMonth
	Dim strDay
	
	Call ExtractDateFrom("<%=GetsvrDate%>",Parent.gServerDateFormat , Parent.gServerDateType ,strYear,strMonth,strDay)
	
	frm1.txtPay_yymm_dt.Focus			'년월 default value setting
	
	frm1.txtPay_yymm_dt.Year = strYear 
	frm1.txtPay_yymm_dt.Month = strMonth 

	frm1.txtYy_mm_dd_dt.Year = strYear 		 '년월일 default value setting
	frm1.txtYy_mm_dd_dt.Month = strMonth 
	frm1.txtYy_mm_dd_dt.Day = strDay
	
End Sub
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iNameArr1,iCodeArr1
        
    Call CommonQueryRs("MINOR_NM,MINOR_CD","B_MINOR","MAJOR_CD = " & FilterVar("H0005", "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iNameArr1 = lgF0
    iCodeArr1 = lgF1
    Call SetCombo2(frm1.txtcboPay_cd,iCodeArr1, iNameArr1,Chr(11))            ''''''''DB에서 불러 condition에서        
End Sub
'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "H", "NOCOOKIE", "BA") %>
End Sub
'========================================================================================================
' Name : OpenDept
' Desc : 부서 POPUP
'========================================================================================================
Function OpenDept(iWhere)
	Dim arrRet
	Dim arrParam(2)
    Dim strBasDt 
    Dim strBasDtAdd
    	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then
		arrParam(0) = frm1.txtFr_dept_cd.value			            '  Code Condition
	ElseIf iWhere = 1 Then
		arrParam(0) = frm1.txtTo_dept_cd.value			            ' Code Condition
	End If
	
    arrParam(1) = ""
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
'	Name : OpenPopUp
'	Description : Code PopUp at vspdData
'=======================================================================================================
Function OpenPopUp(Byval strCode, Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	IsOpenPop = True
		
	Select Case iWhere
	    Case "PROV_TYPE"
			arrParam(0) = "지급구분 팝업"			        ' 팝업 명칭 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = frm1.txtProv_type.value     			' Code Condition
	    	arrParam(3) = ""'frm1.txtProv_type_nm.value									' Name Cindition
	    	arrParam(4) = "major_cd = " & FilterVar("H0040", "''", "S") & ""	   		    	' Where Condition
	    	arrParam(5) = "지급코드"  			            ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"						    ' Field명(0)
	    	arrField(1) = "minor_nm"    					  	' Field명(1)
	    	arrField(2) = ""    				        		' Field명(2)
    
	    	arrHeader(0) = "지급코드"	   		    	    ' Header명(0)
	    	arrHeader(1) = "지급코드명"        		        ' Header명(1)
	    	arrHeader(2) = ""	    							' Header명(1)
	   Case "SECT_CD"
			arrParam(0) = "근무구역 팝업"			        ' 팝업 명칭 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = frm1.txtSect_cd.value     			' Code Condition
	    	arrParam(3) = ""'frm1.txtSect_cd_nm.value									' Name Cindition
	    	arrParam(4) = "major_cd = " & FilterVar("H0035", "''", "S") & ""	   		    	' Where Condition
	    	arrParam(5) = "근무구역코드"  			            ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"						    ' Field명(0)
	    	arrField(1) = "minor_nm"    					  	' Field명(1)
	    	arrField(2) = ""    				        		' Field명(2)
    
	    	arrHeader(0) = "근무구역코드"	   		    	    ' Header명(0)
	    	arrHeader(1) = "근무구역코드명"        		        ' Header명(1)
	    	arrHeader(2) = ""	    							' Header명(1)
	    Case "OCPT_TYPE"
	        arrParam(0) = "직종 팝업"			        ' 팝업 명칭 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = frm1.txtOcpt_type.value     			' Code Condition
	    	arrParam(3) = ""'frm1.txtOcpt_type_nm.value									' Name Cindition
	    	arrParam(4) = "major_cd = " & FilterVar("H0003", "''", "S") & ""	   		    	' Where Condition
	    	arrParam(5) = "근무구역코드"  			            ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"						    ' Field명(0)
	    	arrField(1) = "minor_nm"    					  	' Field명(1)
	    	arrField(2) = ""    				        		' Field명(2)
    
	    	arrHeader(0) = "직종코드"	   		    	    ' Header명(0)
	    	arrHeader(1) = "직종코드명"        		        ' Header명(1)
	    	arrHeader(2) = ""	    							' Header명(1)	    
	End Select	    
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		    "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
		    Case "PROV_TYPE"
				frm1.txtProv_type.focus
		    Case "SECT_CD"
				frm1.txtSect_cd.focus
		    Case "OCPT_TYPE"
				frm1.txtOcpt_type.focus
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
		    Case "PROV_TYPE"
				.txtProv_type.value = arrRet(0) 
				.txtProv_type_nm.value  = arrRet(1) 
				.txtProv_Oldtype.value = arrRet(0) 
				.txtProv_type.focus
		    Case "SECT_CD"
				.txtSect_cd.value = arrRet(0) 
				.txtSect_cd_nm.value  = arrRet(1)
				.txtSect_Oldcd.value = arrRet(0) 
				.txtSect_cd.focus
		    Case "OCPT_TYPE"
		        .txtOcpt_type.value = arrRet(0) 
				.txtOcpt_type_nm.value  = arrRet(1)
				.txtOcpt_Oldtype.value = arrRet(0) 
				.txtOcpt_type.focus
        End Select
	End With

End Function
'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

	Call ggoOper.FormatDate(frm1.txtPay_yymm_dt, Parent.gDateFormat, 2)   '년월 
    Call ggoOper.FormatDate(frm1.txtYy_mm_dd_dt, Parent.gDateFormat, 1)             '년월일 
	
	Call FuncGetAuth(gStrRequestMenuID , Parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")
	
    Call InitVariables
    Call InitComboBox                                        
    Call SetDefaultVal
    Call SetToolbar("1000000000000111")    
     
End Sub
'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub
'======================================================================================================
'   Event Name : txtYyyymm_DblClick
'   Event Desc : 달력 Popup을 호출 
'=======================================================================================================
Sub txtPay_yymm_dt_DblClick(Button) 
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtPay_yymm_dt.Action = 7
		frm1.txtPay_yymm_dt.focus
	End If
End Sub
'======================================================================================================
'   Event Name : txtYyyymm_DblClick
'   Event Desc : 달력 Popup을 호출 
'=======================================================================================================
Sub txtYy_mm_dd_dt_DblClick(Button) 
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtYy_mm_dd_dt.Action = 7
		frm1.txtYy_mm_dd_dt.focus
	End If
End Sub
'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 

  With frm1
	
    If  txtProv_type_Onchange() Then   
        Exit Function
    End if
    
    If  txtSect_cd_OnChange() Then
        Exit Function
    End if	
    If  txtOcpt_type_Onchange() Then
        Exit Function
    End if	    	
    If  txtFr_dept_cd_Onchange() Then      'enter key 로 조회시 시작부서코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if
    
    If  txtTo_dept_cd_Onchange() Then     'enter key 로 조회시 종료부서코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if
 
    Dim Fr_dept_cd , To_dept_cd, rFrDept ,rToDept ,IntRetCd
    
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
    
    If (Fr_dept_cd= "") AND (To_dept_cd="") Then       
    Else
        If Fr_dept_cd > To_dept_cd then
	        Call DisplayMsgBox("800153","X","X","X")	'시작부서코드는 종료부서코드보다 작아야합니다.
            frm1.txtfr_dept_cd.focus()
            Set gActiveElement = document.activeElement
        End IF 
        
    END IF   
    End with
End Function

'========================================================================================
' Function Name : txtProv_type_OnChange()
' Function Desc : 
'========================================================================================
Function txtProv_type_OnChange()
    Dim iDx
    Dim IntRetCd
    
    IF frm1.txtProv_type.value = "" THEN
        frm1.txtProv_type_nm.value = ""
    ELSE
        IntRetCd = CommonQueryRs(" minor_nm "," b_minor "," major_cd = " & FilterVar("H0040", "''", "S") & " and minor_cd =  " & FilterVar(frm1.txtProv_type.value , "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false then
			Call DisplayMsgBox("800140","X","X","X")	'지급내역코드에 등록되지 않은 코드입니다.
            frm1.txtProv_type_nm.value = ""
            frm1.txtProv_type.focus
            txtProv_type_Onchange = true
            Exit Function
        ELSE    
            frm1.txtProv_type_nm.value = Trim(Replace(lgF0,Chr(11),""))   '수당코드 
        END IF
    END IF 
End Function
'========================================================================================
' Function Name : txtSect_cd_OnChange()
' Function Desc : 
'========================================================================================
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
	        Set gActiveElement = document.ActiveElement
	        txtSect_cd_OnChange = true
	        Exit Function
	    Else
	        frm1.txtSect_cd_nm.value = Trim(Replace(lgF0, Chr(11), ""))

	    End If
    End If
    	    
End Function
'========================================================================================
' Function Name : txtProv_type_OnChange()
' Function Desc : 
'========================================================================================
Function txtOcpt_type_OnChange()
   Dim iDx
    Dim IntRetCd
        
    If frm1.txtOcpt_type.value = "" Then
        frm1.txtOcpt_type_nm.value = ""
    ELSE
        IntRetCd = CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0003", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtOcpt_type.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
        IF IntRetCd = False THEN
            Call DisplayMsgBox("970000","X","직종코드","X")
	        frm1.txtOcpt_type_nm.value = ""
	        Set gActiveElement = document.ActiveElement
	        txtOcpt_type_OnChange = true
	        Exit Function
	    Else
	        frm1.txtOcpt_type_nm.value = Trim(Replace(lgF0, Chr(11), ""))
	    End If
    End If
End Function

'========================================================================================================
'   Event Name : txtFr_dept_cd_change
'   Event Desc :
'========================================================================================================
Function txtFr_dept_cd_Onchange()
    Dim IntRetCd,Dept_Nm,Internal_cd
    
    If frm1.txtFr_dept_cd.value = "" Then
		frm1.txtFr_dept_nm.value = ""
		frm1.txtFr_internal_cd.value = ""        
    Else
        IntRetCd = FuncDeptName(frm1.txtFr_dept_cd.value , "" , lgUsrIntCd,Dept_Nm , Internal_cd)
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgBox("800098","X","X","X")	'부서정보코드에 등록되지 않은 코드입니다.
            Else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
		    frm1.txtFr_dept_nm.value = ""
		    frm1.txtFr_internal_cd.value = ""
            Set gActiveElement = document.ActiveElement 
		    txtFr_dept_cd_Onchange = True
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
    
    If frm1.txtTo_dept_cd.value = "" Then
		frm1.txtTo_dept_nm.value = ""
		frm1.txtTo_internal_cd.value = ""        
    Else
        IntRetCd = FuncDeptName(frm1.txtTo_dept_cd.value , "" , lgUsrIntCd,Dept_Nm , Internal_cd)
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgBox("800098","X","X","X")	'부서정보코드에 등록되지 않은 코드입니다.
            Else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
		    frm1.txtTo_dept_nm.value = ""
		    frm1.txtTo_internal_cd.value = ""
            Set gActiveElement = document.ActiveElement 
 		    txtTo_dept_cd_Onchange = True
            Exit Function      
        Else
			frm1.txtTo_dept_nm.value = Dept_Nm
		    frm1.txtTo_internal_cd.value = Internal_cd      
        End if 
    End if  
    
End Function

'========================================================================================
' Function Name : btnAction_OnClick()
' Function Desc : This function is related to Preview Button
'========================================================================================
Function btnAction_OnClick()
    Dim Fr_dept_cd, To_dept_cd, rFrDept ,rToDept
	dim IntRetCd
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If

    Dim strVal, RetFlag, pay_yymm, prov_type, gigup_type1, strDate 
	
	pay_yymm  = frm1.txtPay_yymm_dt.Year & Right("0" & frm1.txtPay_yymm_dt.Month, 2)
	strDate   = frm1.txtYy_mm_dd_dt.Year & Right("0" & frm1.txtYy_mm_dd_dt.Month, 2) & Right("0" & frm1.txtYy_mm_dd_dt.Day, 2) 
	prov_type = frm1.txtProv_type.value

	If frm1.txtGigup_type(0).checked Then 
		gigup_type1 = "1"
	Elseif frm1.txtGigup_type(1).checked Then 
		gigup_type1 = "2"
	Else 
		gigup_type1 = "3" 
	End if	
		
	If  txtProv_type_Onchange() Then   
	    Exit Function
	End if
		    
	If  txtSect_cd_OnChange() Then
	    Exit Function
	End if	
	If  txtOcpt_type_Onchange() Then
	    Exit Function
	End if	    	
	If  txtFr_dept_cd_Onchange() Then      'enter key 로 조회시 시작부서코드를 check후 해당사항 없으면 query종료...
	    Exit Function
	End if
		    
	If  txtTo_dept_cd_Onchange() Then     'enter key 로 조회시 종료부서코드를 check후 해당사항 없으면 query종료...
	    Exit Function
	End if

	
	With frm1
       Fr_dept_cd = .txtFr_internal_cd.value
        To_dept_cd = .txtTo_internal_cd.value
        
        If Fr_dept_cd = "" then
            IntRetCd = FuncGetTermDept(lgUsrIntCd ,"", rFrDept ,rToDept)
            .txtFr_internal_cd.value = rFrDept
            .txtFr_dept_nm.value = ""
        End If	
            
        If To_dept_cd = "" then
            IntRetCd = FuncGetTermDept(lgUsrIntCd ,"", rFrDept ,rToDept)
            .txtTo_internal_cd.value = rToDept
            .txtTo_dept_nm.value = ""
        End If  

        Fr_dept_cd = .txtFr_internal_cd.value
        To_dept_cd = .txtTo_internal_cd.value

        If (Fr_dept_cd<> "") AND (To_dept_cd<>"") Then       
    
            If Fr_dept_cd > To_dept_cd then
	            Call DisplayMsgBox("800153","X","X","X")	'시작부서코드는 종료부서코드보다 작아야합니다.
                .txtFr_dept_cd.focus()
                Set gActiveElement = document.activeElement
                Exit Function
            End IF 
            
        END IF   

	    RetFlag = DisplayMsgBox("900018", Parent.VB_YES_NO,"x","x")   '☜ 바뀐부분	

	    If RetFlag = VBNO Then
	    	Exit Function
	    End IF

	    If LayerShowHide(1) = False Then
	       Exit Function
	    End If		
	    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001							'☜: 비지니스 처리 ASP의 상태 	    	    
		strVal = strVal & "&txtPay_yymm_dt=" & Trim(pay_yymm)
		strVal = strVal & "&txtProv_type=" & Trim(.txtProv_type.value)
		strVal = strVal & "&txtYy_mm_dd_dt=" & Trim(strDate)
		strVal = strVal & "&txtcboPay_cd=" & Trim(.txtcboPay_cd.value)
		strVal = strVal & "&txtSect_cd=" & Trim(.txtSect_cd.value)
		strVal = strVal & "&txtOcpt_type=" & Trim(.txtOcpt_type.value)
		strVal = strVal & "&txtFr_dept_cd=" & Trim(.txtFr_Internal_cd.value)
		strVal = strVal & "&txtTo_dept_cd=" & Trim(.txtTo_Internal_cd.value)
		strVal = strVal & "&txtGigup_type=" & Trim(gigup_type1)
		strVal = strVal & "&txtStand_amt=" & Trim(.txtStand_amt.Text)				

		Call RunMyBizASP(MyBizASP, strVal)

    End With    
End Function

Function subVatDiskOK(ByVal pFileName) 
Dim strVal
    Err.Clear                                                               '☜: Protect system from crashing
    If pFileName <> "" Then
	    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0002							'☜: 비지니스 처리 ASP의 상태 
	    strVal = strVal & "&txtFileName=" & pFileName							'☆: 조회 조건 데이타	
	    Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	End If
End Function


Sub txtProv_type_Keypress(Key)
Dim pflag
    If Key = 13 Then
        Call FncQuery()
    End If
End Sub     

Sub txtSect_cd_Keypress(Key)
Dim pflag
    If Key = 13 Then
        Call FncQuery()
    End If
End Sub     

Sub txtOcpt_type_Keypress(Key)
Dim pflag
    If Key = 13 Then
        Call FncQuery()
    End If
End Sub     

Sub txtFr_dept_cd_Keypress(Key)
Dim pflag
    If Key = 13 Then
        Call FncQuery()
    End If
End Sub     

Sub txtto_dept_cd_Keypress(Key)
Dim pflag
    If Key = 13 Then
        Call FncQuery()
    End If
End Sub     


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
End Function
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	

<SCRIPT LANGUAGE="JavaScript">
<!-- Hide script from old browsers

function setCookie(name, value, expire)
{
}
-->
</SCRIPT>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB2">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>은행이체파일생성</font></td>
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
					    <FIELDSET CLASS="CLSFLD" STYLE="HEIGHT:100%">
						<TABLE <%=LR_SPACE_TYPE_40%>>   
						    <TR>
								<TD CLASS=TD5  NOWRAP>해당년월</TD>
								<TD CLASS=TD6  NOWRAP><script language =javascript src='./js/h6020ma1_txtPay_yymm_dt_txtPay_yymm_dt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>지급구분</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE="Text" ID = "txtProv_type" NAME="txtProv_type" SIZE=7 MAXLENGTH=1 tag="12XXXU" ALT="지급구분코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtProv_type.value, 'PROV_TYPE')"> 
								<INPUT TYPE="Text" ID=txtProv_type_nm NAME="txtProv_type_nm" SIZE=20 MAXLENGTH=30 tag="14XXXU" ALT="지급구분코드명">
								<INPUT TYPE="HIDDEN" ID=txtProv_Oldtype NAME="txtProv_Oldtype" SIZE=20 MAXLENGTH=30 tag="14XXXU" ALT="지급구분코드명"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5>입금희망일</TD>
								<TD CLASS=TD6><script language =javascript src='./js/h6020ma1_txtYy_mm_dd_dt_txtYy_mm_dd_dt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>급여구분</TD>
	                            <TD CLASS=TD6 NOWRAP><SELECT Name="txtcboPay_cd" ALT="급여구분" CLASS=cboNormal tag="11"><OPTION Value=""></OPTION></SELECT></TD>
							</TR>			
							<TR>
							    <TD CLASS=TD5 NOWRAP>근무구역</TD>
							    <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT ID="txtSect_cd" NAME="txtSect_cd" SIZE=10  tag="11XXXU" ALT="근무구역코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtSect_cd.Value, 'SECT_CD')">
							                           <INPUT TYPE=TEXT ID="txtSect_cd_nm" NAME="txtSect_cd_nm" SIZE=15  tag="14XXXU" ALT="근무구역명">
							                           <INPUT TYPE="HIDDEN" ID="txtSect_Oldcd" NAME="txtSect_Oldcd" SIZE=20 MAXLENGTH=30 tag="14XXXU" ALT="근무구역코드"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>직종</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT ID="txtOcpt_type" NAME="txtOcpt_type" SIZE=10 MAXLENGTH=10 ALT="직종" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtOcpt_type.Value, 'OCPT_TYPE')">
								                     <INPUT TYPE=TEXT ID="txtOcpt_type_nm" NAME="txtOcpt_type_nm" SIZE=15  tag="14XXXU" ALT="직종">
							                         <INPUT TYPE="HIDDEN" ID="txtOcpt_Oldtype" NAME="txtOcpt_Oldtype" SIZE=20 MAXLENGTH=30 tag="14XXXU" ALT="직종"></TD></TD>
							</TR>			
							<TR>
							    <TD CLASS=TD5 NOWRAP>부서</TD>
								<TD CLASS=TD6 NOWRAP><INPUT ID="txtFr_dept_cd" NAME="txtFr_dept_cd" ALT="부서코드" TYPE="Text" SiZE="10" MAXLENGTH="10" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(0)">
			                            <INPUT ID="txtFr_dept_nm" NAME="txtFr_dept_nm" ALT="부서코드명" TYPE="Text" SiZE="20" MAXLENGTH="40" tag="14XXXU">
		                                <INPUT ID="txtFr_Internal_cd" NAME="txtFr_Internal_cd" ALT="내부부서코드" TYPE="HIDDEN" SiZE="7" MAXLENGTH="7" tag="14XXXU">&nbsp;~</TD>
							</TR>			
							<TR>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP><INPUT ID="txtto_dept_cd" NAME="txtto_dept_cd" ALT="부서코드" TYPE="Text" SIZE="10" MAXLENGTH="10" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnname" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(1)">
							            <INPUT ID="txtto_dept_nm" NAME="txtto_dept_nm" ALT="부서코드명" TYPE="Text"SIZE="20" MAXLENGTH="40" tag="14XXXU">
							            <INPUT ID="txtTo_Internal_cd" NAME="txtTo_Internal_cd" ALT="내부부서코드" TYPE="HIDDEN" SiZE="7" MAXLENGTH="7" tag="14XXXU"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>지급방식</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=txtGigup_type ID=Rb_tot tag="12"><LABEL FOR=Rb_tot>기준금액 제외한 은행이체</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=txtGigup_type ID=Rb_dur tag="12"><LABEL FOR=Rb_dur>기준금액 미만 금액만 은행이체</LABEL></TD>
							</TR>			
							<TR>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=txtGigup_type ID=Rb_dept Checked tag="12"><LABEL FOR=Rb_dept>모든 금액 은행이체</LABEL></TD>
							</TR>
							
	    					<TR>
              				    <TD CLASS=TD5 NOWRAP>기준금액</TD>
	                   			<TD CLASS=TD6><script language =javascript src='./js/h6020ma1_txtStand_amt_txtStand_amt.js'></script></TD>
	                   			
	                   	    </TR>		    		
    					</TABLE>
						</FIELDSET>
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
		                <BUTTON NAME="btnAction" CLASS="CLSMBTN" >실행</BUTTON>
		            </TD>
					<TD WIDTH=* ALIGN="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
