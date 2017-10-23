<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Loan
'*  3. Program ID           : f4202ma1
'*  4. Program Name         : 은행기초차입금등록 
'*  5. Program Desc         : Register of Loan Master
'*  6. Comproxy List        : FL0061, FL0069
'*  7. Modified date(First) : 2002.05.20
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Oh Soo Min
'* 10. Modifier (Last)      : Ahn Do hyun
'* 11. Comment              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--
#########################################################################################################
												1. 선 언 부 
##########################################################################################################
******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->							<!-- '⊙: 화면처리ASP에서 서버작업이 필요한 경우  -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAMain.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAEvent.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit                                                            '☜: indicates that All variables must be declared in advance 

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Const BIZ_PGM_ID = "f4204mb1.asp"			 '☆: 비지니스 로직 ASP명 
Const BIZ_PGM_ID2 = "f4204mb2.asp"


'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 

Dim tempStrFg

Dim lgLoanNo
Dim lgNextNo						'☜: 화면이 Single/SingleMulti 인경우만 해당 
Dim lgPrevNo						' ""
Dim lgtempStrFg
'==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= 
'----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 
<!-- #Include file="../../inc/lgvariables.inc" -->	
'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 

Dim IsOpenPop

<%
dim dtToday
dtToday = GetSvrDate
%>
 '#########################################################################################################
'												2. Function부 
'
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'######################################################################################################### 
 '==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE                                               '⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                                                '⊙: Indicates that no value changed
    lgIntGrpCount = 0                                                       '⊙: Initializes Group View Size
    frm1.hOrgChangeId.value = parent.gChangeOrgId
    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False														'☆: 사용자 변수 초기화 
End Sub

'==========================================  2.1.1 InitVariablesForCopy()  ======================================
'	Name : InitVariablesForCopy()
'	Description : The variables will be initialized when the copy button is clicked.
'========================================================================================================= 
''FINE_20030725_HC_Copy기능_START
Sub InitVariablesForCopy()
	With frm1
		.txtLoanNo.value = ""
		lgLoanNo = ""

		.txtXchrate.Text = 0
		.txtLoanLocAmt.text = 0
		.txtBasIntPayLocAmt.text = 0
		.txtBasRdpLocAmt.text = 0
		.txtRdpAmt.text = 0
		.txtRdpLocAmt.text = 0
		.txtIntPayAmt.text = 0
		.txtIntPayLocAmt.text = 0
		.txtLoanBalAmt.text = 0
		.txtLoanBalLocAmt.text = 0
		.txtBasIntPayLocAmt.text = 0
		.txtBasRdpLocAmt.text = 0
		
		.cboRdpClsFg.value = "N"
	End With
End Sub
''FINE_20030725_HC_Copy기능_END

'=============================== 2.1.2 LoadInfTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub


 '******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'********************************************************************************************************* 
 '==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtLoanDt.text = UNIConvDateAToB("<%=dtToday%>",parent.gServerDateFormat,gDateFormat)
	frm1.txtBasicLoanDt.text = UNIConvDateAToB("<%=dtToday%>",parent.gServerDateFormat,gDateFormat)

	With frm1
		.Rb_IntVotl1.Checked = True			
		.hRb_Cur1.value = "1"								
		.htxtLcNo.value = ""
		.hClsRoFg.value = "N"
		.txtPrRdpUnitAmt.value = "0"
		.txtPrRdpUnitLocAmt.value = "0"
		.htxtLoanPlcType.value = "BK"
		.txtDocCur.value = parent.gCurrency
	End With
	
	lgBlnFlgChgValue = False
End Sub
'========================================================================================================
' Name : CookiePage()
' Desc : Write or Read cookie value 
'========================================================================================================
Sub CookiePage(Kubun)
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'--------------------------------------------------------------
' ComboBox 초기화 
'-------------------------------------------------------------- 
Sub InitComboBox()

	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1040", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboPrRdpCond ,lgF0  ,lgF1  ,Chr(11))
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1020", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboLoanFg ,lgF0  ,lgF1  ,Chr(11))
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1030", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboIntPayStnd ,lgF0  ,lgF1  ,Chr(11))
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1090", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboIntBaseMthd ,lgF0  ,lgF1  ,Chr(11))
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F2020", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboRdpClsFg ,lgF0  ,lgF1  ,Chr(11))

End Sub
'******************************************  2.3 Operation 처리함수  *************************************
'	기능: Operation 처리부분 
'	설명: Tab처리, Reference등을 행한다. 
'********************************************************************************************************* 

 '******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 

 '========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : 중복되어 있는 PopUp을 재정의, 재정의가 필요한 경우는 반드시 CommonPopUp.vbs 와 
'				  ManufactPopUp.vbs 에서 Copy하여 재정의한다.
'========================================================================================================= 
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	frm1.hOrgChangeId.value = parent.gChangeOrgId

	Select Case iWhere
		Case 0		'통화코드 
			arrParam(0) = frm1.txtDocCur.Alt								' 팝업 명칭 
			arrParam(1) = "B_CURRENCY"	 									' TABLE 명칭 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = ""												' Where Condition
			arrParam(5) = frm1.txtDocCur.Alt								' 조건필드의 라벨 명칭 

		    arrField(0) = "CURRENCY"										' Field명(0)
		    arrField(1) = "CURRENCY_DESC"									' Field명(1)
   
		    arrHeader(0) = "통화코드"									' Header명(0)
			arrHeader(1) = "통화코드명"									' Header명(1)
			
		Case 2		'부서코드 
			arrParam(0) = strCode		            '  Code Condition
		   	arrParam(1) = frm1.txtBasicLoanDt.Text
			arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  
			arrParam(3) = "F"									' 결의일자 상태 Condition  
		
		Case 5		'차입처							
			lgtempStrFg = "B"		
			arrParam(0) = frm1.txtBankLoanCd.Alt
			arrParam(1) = "B_BANK A"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = frm1.txtBankLoanCd.Alt
	
			arrField(0) = "A.BANK_CD" 
			arrField(1) = "A.BANK_NM"
				    
			arrHeader(0) = "은행코드"
			arrHeader(1) = "은행명"			
	
		Case 6		'차입용도 
			arrParam(0) = frm1.txtLoanType.Alt
			arrParam(1) = "B_MINOR A"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = "A.MAJOR_CD = " & FilterVar("F1000", "''", "S") & " "
			arrParam(5) = frm1.txtLoanType.Alt
	
			arrField(0) = "A.MINOR_CD"
			arrField(1) = "A.MINOR_NM"
			    
			arrHeader(0) = frm1.txtLoanType.Alt
			arrHeader(1) = frm1.txtLoanTypeNm.Alt
			
		Case 7
			If frm1.txtLoanAcctCd.className = "protected" Then Exit Function    
			
			arrParam(0) = "차입금계정팝업"								' 팝업 명칭 
			arrParam(1) = "A_ACCT	A,A_ACCT_GP 	B, A_JNL_ACCT_ASSN	C	"				' TABLE 명칭 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = "C.TRANS_TYPE = " & FilterVar("FI001", "''", "S") & "  " 			' Where Condition
			arrParam(4) = arrParam(4) & " AND 	C.ACCT_CD = A.ACCT_CD AND A.GP_CD = B.GP_CD "
			arrParam(4) = arrParam(4) & " AND 	A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND 	DR_CR_FG = " & FilterVar("CR", "''", "S") & "  " 
			arrParam(4) = arrParam(4) & " AND JNL_CD = " & FilterVar(frm1.cboLoanFg.Value, "''", "S") 		
			arrParam(5) = frm1.txtLoanAcctCd.Alt							' 조건필드의 라벨 명칭 
							
			arrField(0) = "A.ACCT_CD"									' Field명(0)
			arrField(1) = "A.ACCT_NM"									' Field명(1)
			arrField(2) = "B.GP_CD"										' Field명(2)
			arrField(3) = "B.GP_NM"					 					' Field명(3)
			
			arrHeader(0) = frm1.txtLoanAcctCd.Alt									' Header명(0)
			arrHeader(1) = frm1.txtLoanAcctNm.Alt								' Header명(1)
			arrHeader(2) = "그룹코드"									' Header명(2)
			arrHeader(3) = "그룹명"										' Header명(3)		
		Case 8
			If frm1.txtIntAcctCd.className = "protected" Then Exit Function    
						
				arrParam(0) = "이자계정팝업"								' 팝업 명칭 
				arrParam(1) = "A_ACCT	A,A_ACCT_GP 	B, A_JNL_ACCT_ASSN	C	"				' TABLE 명칭 
				arrParam(2) = strCode											' Code Condition
				arrParam(3) = ""												' Name Cindition
				arrParam(4) = "C.TRANS_TYPE = " & FilterVar("FI002", "''", "S") & "  " 			' Where Condition
				arrParam(4) = arrParam(4) & " AND 	C.ACCT_CD = A.ACCT_CD AND A.GP_CD = B.GP_CD "
				arrParam(4) = arrParam(4) & " AND 	A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND 	DR_CR_FG = " & FilterVar("DR", "''", "S") & "  " 
				arrParam(4) = arrParam(4) & " AND JNL_CD = " & FilterVar(frm1.cboIntPayStnd.Value, "''", "S") 	
				arrParam(5) = frm1.txtIntAcctCd.Alt							' 조건필드의 라벨 명칭		

				arrField(0) = "A.Acct_CD"									' Field명(0)
				arrField(1) = "A.Acct_NM"									' Field명(1)
				arrField(2) = "B.GP_CD"										' Field명(2)
				arrField(3) = "B.GP_NM"										' Field명(3)
			
				arrHeader(0) = frm1.txtIntAcctCd.Alt									' Header명(0)
				arrHeader(1) = frm1.txtIntAcctNm.Alt								' Header명(1)
				arrHeader(2) = "그룹코드"									' Header명(2)
				arrHeader(3) = "그룹명"										' Header명(3)				
		Case Else
			Exit Function
	End Select
    
	IsOpenPop = True

	Select Case iWhere
		Case 2
			arrRet = window.showModalDialog("../../comasp/DeptPopupDtA2.asp", Array(window.parent,arrParam), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		Case 3, 4
			arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
					 "dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		Case Else
			arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
					 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End Select
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
		Case 0	'통화 
			frm1.txtDocCur.focus
		Case 2	'부서 
			frm1.txtDeptCd.focus
		Case 5	'차입은행 
			frm1.txtBankLoanCd.focus			
		Case 6	'차입용도 
			frm1.txtLoanType.focus

		Case 7
			frm1.txtLoanAcctCd.focus
		Case 8
			frm1.txtIntAcctCd.focus
		End Select
		Exit Function
	Else
		Select Case iWhere
		Case 0	'통화 
			frm1.txtDocCur.value = arrRet(0)
			
			If parent.gCurrency = UCase(Trim(frm1.txtDocCur.value)) Then
				frm1.txtXchrate.Text = "1"
			Else
				Call FncCalcRate
			End If
			
			call txtDocCur_OnChange()
			frm1.txtDocCur.focus
		Case 2	'부서 
			frm1.txtBasicLoanDt.text = arrRet(3)
            frm1.txtDeptCd.value = arrRet(0)
            frm1.txtDeptNm.value = arrRet(1)
			call txtDeptCd_OnChange()  
			frm1.txtDeptCd.focus
		Case 5	'차입은행 
			frm1.txtBankLoanCd.value = arrRet(0)
			frm1.txtBankLoanNm.value = arrRet(1)
			frm1.txtBankLoanCd.focus			
		Case 6	'차입용도 
			frm1.txtLoanType.value = arrRet(0)
			frm1.txtLoanTypeNm.value = arrRet(1)
			frm1.txtLoanType.focus

		Case 7
			frm1.txtLoanAcctCd.value = arrRet(0)
			frm1.txtLoanAcctNm.value = arrRet(1)
			frm1.txtLoanAcctCd.focus
		Case 8
			frm1.txtIntAcctCd.value = arrRet(0)
			frm1.txtIntAcctNm.value = arrRet(1)
			frm1.txtIntAcctCd.focus
		End Select
	End If
	
	lgBlnFlgChgValue = True
End Function

'============================================================
'부서코드 팝업 
'============================================================
Function OpenPopupDept(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function
	If frm1.txtDeptCd.className = parent.UCN_PROTECTED Then Exit Function
	
	arrParam(0) = strCode				'부서코드 
	arrParam(1) = frm1.txtBasicLoanDt.Text	'날짜(Default:현재일)
	arrParam(2) = "1"					'부서권한(lgUsrIntCd)
	
	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/DeptPopupDt.asp", Array(window.parent,arrParam), _
			"dialogWidth=500px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDeptCd.focus
		Exit Function
	End If

	frm1.txtDeptCd.value = arrRet(0)
	frm1.txtDeptNm.value = arrRet(1)
	frm1.txtDeptCd.focus
	
	lgBlnFlgChgValue = True
End Function

'============================================================
'차입금번호 팝업 
'============================================================
Function OpenPopupLoan()
	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(3)	

	If IsOpenPop = True Then Exit Function
	iCalledAspName = AskPRAspName("f4202ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "f4202ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True
	arrRet = window.showModalDialog(iCalledAspName & "?PGM=" & gStrRequestMenuID , Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False
	
	If arrRet(0) = ""  Then			
		frm1.txtLoanNo.focus
		Exit Function
	Else		
		frm1.txtLoanNo.value = arrRet(0)
		frm1.txtLoanNo.focus
	End If
End Function

'============================================================
'회계전표 팝업 
'============================================================
Function OpenPopupGL()

	Dim arrRet
	Dim arrParam(1)	

	If IsOpenPop = True Then Exit Function
	
	arrParam(0) = Trim(frm1.txtGlNo.value)	'회계전표번호 
	arrParam(1) = ""						'Reference번호 

	IsOpenPop = True
   
	arrRet = window.showModalDialog("../../ComAsp/a5120ra1.asp", Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
End Function

'============================================================
'결의전표 팝업 
'============================================================
Function OpenPopupTempGL()

	Dim arrRet
	Dim arrParam(1)	

	If IsOpenPop = True Then Exit Function
	
	arrParam(0) = Trim(frm1.txtTempGlNo.value)	'결의전표번호 
	arrParam(1) = ""							'Reference번호 

	IsOpenPop = True
   
	arrRet = window.showModalDialog("../../ComAsp/a5130ra1.asp", Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
End Function

 '=========================================================================================================
'	Name : FncCalcRate()
'	Description : lookup exchange rate 
'========================================================================================================= 
Function FncCalcRate()
End Function

Function fncGetAPamt()
End Function

 '++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
'========================================================================================================
'	Desc : Cookie Setting
'========================================================================================================
Function CookiePage(ByVal Kubun)

'	Const CookieSplit = 4877						'Cookie Split String : CookiePage Function Use
	Dim strTemp

	Select Case Kubun		
	Case "FORM_LOAD"
		strTemp = ReadCookie("LOAN_NO")
		Call WriteCookie("LOAN_NO", "")
		
		If strTemp = "" then Exit Function
					
		frm1.txtLoanNo.value = strTemp
				
		If Err.number <> 0 Then
			Err.Clear
			Call WriteCookie("LOAN_NO", "")
			Exit Function 
		End If
				
		Call FncQuery()
	
	Case Else
		Exit Function
	End Select
End Function	


'#########################################################################################################
'												3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'##########################################################################################################
 '******************************************  3.1 Window 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 
'======================================================================================================
'   Event Name : Radio_Cur 
'   Event Desc : 차입처 변경 
'=======================================================================================================
'=======================================================================================================
'   Event Name : _DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtLoanDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtLoanDt.Action = 7
		Call SetFocusToDocument("M")
		Frm1.txtLoanDt.Focus
		        
    End If
End Sub

Sub txtBasicLoanDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtBasicLoanDt.Action = 7
		Call SetFocusToDocument("M")
		Frm1.txtBasicLoanDt.Focus        
    End If
End Sub

Sub txtDueDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDueDt.Action = 7
		Call SetFocusToDocument("M")
		Frm1.txtDueDt.Focus        
    End If
End Sub

Sub txt1StIntDueDt_DblClick(Button)
    If Button = 1 Then
        frm1.txt1StIntDueDt.Action = 7
		Call SetFocusToDocument("M")
		Frm1.txt1StIntDueDt.Focus        
    End If
End Sub

Sub txt1StPrRdpDt_DblClick(Button)
    If Button = 1 Then
        frm1.txt1StPrRdpDt.Action = 7
		Call SetFocusToDocument("M")
		Frm1.txt1StPrRdpDt.Focus        
    End If
End Sub

'=======================================================================================================
'   Event Name : _Change()
'   Event Desc : 원금상환방법별 Set Protected/Required Fields
'=======================================================================================================
Sub cboPrRdpCond_OnChange()
	 '최초원금상환일, 원금상환일, 상환주기, 원금상환액 
	Select Case frm1.cboPrRdpCond.value
	Case "EQ"	'균등상환 
		Call ggoOper.SetReqAttr(frm1.txt1StPrRdpDt, "N")	'N:Required, Q:Protected, D:Default
		Call ggoOper.SetReqAttr(frm1.txtPrRdpPerd, "N")
		Call ggoOper.SetReqAttr(frm1.txtPrRdpUnitAmt, "D")
		Call ggoOper.SetReqAttr(frm1.txtPrRdpUnitLocAmt, "D")
	Case "EX",""	'만기상환 
		frm1.txt1StPrRdpDt.Text = ""
		'frm1.txtPrPayDt.Text    = ""
		frm1.txtPrRdpPerd.Text  = ""
		Call ggoOper.SetReqAttr(frm1.txt1StPrRdpDt, "Q")	'N:Required, Q:Protected, D:Default
		Call ggoOper.SetReqAttr(frm1.txtPrRdpPerd, "Q")
		Call ggoOper.SetReqAttr(frm1.txtPrRdpUnitAmt, "Q")
		Call ggoOper.SetReqAttr(frm1.txtPrRdpUnitLocAmt, "Q")
	Case Else
	End Select
End Sub

'=======================================================================================================
'   Event Name : _Change()
'   Event Desc : 이자지급형태별 Set Protected/Required Fields
'=======================================================================================================
Sub cboIntPayStnd_Change()
	 '최초이자지급일	 	 
	Select Case frm1.cboIntPayStnd.value
	Case "AI"	'선급 
		frm1.txt1StIntDueDT.Text = ""	
		Call ggoOper.SetReqAttr(frm1.txt1StIntDueDT, "Q")	'N:Required, Q:Protected, D:Default		
	Case "DI"	'후급		
		Call ggoOper.SetReqAttr(frm1.txt1StIntDueDT, "N")	'N:Required, Q:Protected, D:Default		
	Case Else
		Call ggoOper.SetReqAttr(frm1.txt1StIntDueDT, "Q")	'N:Required, Q:Protected, D:Default		
	End Select
	frm1.hRdpSprdFg.value = "N"
End Sub

'=======================================================================================================
'   Event Name : _Change()
'   Event Desc : 이자지급형태별 Set Protected/Required Fields
'=======================================================================================================
Sub cboIntPayStnd_OnChange()
	Call cboIntPayStnd_Change()

	frm1.txtIntAcctCd.value = ""
	frm1.txtIntAcctNm.value = ""
End Sub

'=======================================================================================================
'   Event Name : _Change()
'   Event Desc : 
'=======================================================================================================
Function Radio1_onChange									'이자율변동성(확정)	
	lgBlnFlgChgValue = True
End Function

Function Radio2_onChange									'이자율변동성(변동)	
	lgBlnFlgChgValue = True
End Function

Function Radio3_onChange	
	frm1.txtIntPayPerd.Text = ""
	lgBlnFlgChgValue = True
End Function

Function Radio6_onChange									
	lgBlnFlgChgValue = True
	frm1.hRdpSprdFg.value = "N"
End Function

Function Radio7_onChange									'만기일포함여부 
	lgBlnFlgChgValue = True
	frm1.hRdpSprdFg.value = "N"
End Function

Function Radio8_onChange									
	lgBlnFlgChgValue = True
	frm1.hRdpSprdFg.value = "N"
End Function

Sub txtLoanAcctCd_OnChange()
	frm1.txtLoanAcctNm.value = ""
	lgBlnFlgChgValue = True	
End Sub 

Sub txtIntAcctCd_OnChange()
	frm1.txtIntAcctNm.value = ""
	lgBlnFlgChgValue = True	
End Sub 

'==========================================================================================
'   Event Name : cboLoanFg_OnChange
'   Event Desc : 
'==========================================================================================
Sub cboLoanFg_OnChange()
	frm1.txtLoanAcctCd.value = ""
	frm1.txtLoanAcctNm.value = ""
	lgBlnFlgChgValue = True
End Sub 

Sub Type_itemChange()
    lgBlnFlgChgValue = True
End Sub

'==========================================================================================
'   Event Name : txtDeptCd_Change
'   Event Desc : 
'==========================================================================================
Sub txtDeptCd_OnChange()
        
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii, jj

	If Trim(frm1.txtDeptCd.value = "") Then		Exit sub
	If Trim(frm1.txtBasicLoanDt.Text = "") Then		Exit sub
    lgBlnFlgChgValue = True

		'----------------------------------------------------------------------------------------
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtBasicLoanDt.Text, gDateFormat,""), "''", "S") & "))"			

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
		Else 
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)
					
			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))
				frm1.hOrgChangeId.value = Trim(arrVal2(2))
			Next	
			
		End If
		'----------------------------------------------------------------------------------------
End Sub

'==========================================================================================
'   Event Name : txtLoanDt_Change
'   Event Desc : 
'==========================================================================================
Sub txtLoanDt_Change()
    Call FncCalcRate()
    lgBlnFlgChgValue = True
End Sub

Sub txtBasicLoanDt_Change()

	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii, jj
	Dim arrVal1, arrVal2

	If Trim(frm1.txtDeptCd.value) <> "" and Trim(frm1.txtBasicLoanDt.Text <> "") Then
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtBasicLoanDt.Text, gDateFormat,""), "''", "S") & "))"			

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
		Else 
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)
						
			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))
				If Trim(arrVal2(2)) <> Trim(frm1.hOrgChangeId.value) Then
					frm1.txtDeptCd.value = ""
					frm1.txtDeptNm.value = ""
					frm1.hOrgChangeId.value = Trim(arrVal2(2))
				End If
			Next
			
		End If
	End If

    lgBlnFlgChgValue = True
End Sub

Sub txtDueDt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtBankLoanCd_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txt1StIntDueDt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txt1StPrRdpDt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtXchRate_Change()
'	frm1.txtLoanLocAmt.Text ="0"
	lgBlnFlgChgValue = True
End Sub

Sub txtLoanAmt_Change()
'	frm1.txtLoanLocAmt.Text ="0"
	lgBlnFlgChgValue = True
End Sub

Sub txtLoanLocAmt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtLoanLocAmt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtBasRdpAmt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtBasIntPayAmt_Change()
	lgBlnFlgChgValue = True
End Sub	

Sub txtPrRdpPerd_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtIntRate_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtIntPayPerd_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtIntPayAmt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtDocCur_OnChange()
'    lgBlnFlgChgValue = True
    IF CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
		Call CurFormatNumericOCX()
	END IF	    

 
End Sub


'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029																'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field
        
    Call InitVariables																'⊙: Initializes local global variables
    Call InitComboBox
	
	Call ggoOper.FormatNumber(frm1.txtPrRdpPerd, "99", "1", False)					'상환주기 
	Call ggoOper.FormatNumber(frm1.txtIntPayPerd, "99", "1", False)					'이자지급주기 
	Call ggoOper.FormatNumber(frm1.txtIntRate, "99.999999", "0", False, 6)			'이자율		
	
	Call FncNew()
	Call CookiePage("FORM_LOAD")

	Call SetDefaultVal	
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

 '#########################################################################################################
'												4. Common Function부 
'	기능: Common Function
'	설명: 환율처리함수, VAT 처리 함수 
'######################################################################################################### 


 '#########################################################################################################
'												5. Interface부 
'	기능: Interface
'	설명: 각각의 Toolbar에 대한 처리를 행한다. 
'	      Toolbar의 위치순서대로 기술하는 것으로 한다. 
'	<< 공통변수 정의 부분 >>
' 	공통변수 : Global Variables는 아니지만 각각의 Sub나 Function에서 자주 사용하는 변수로 변수명은 
'				통일하도록 한다.
' 	1. 공통컨트롤을 Call하는 변수 
'    	   ADF (ADS, ADC, ADF는 그대로 사용)
'    	   - ADF는 Set하고 사용한 뒤 바로 Nothing 하도록 한다.
' 	2. 공통컨트롤에서 Return된 값을 받는 변수 
'    		strRetMsg
'######################################################################################################### 

 '*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'	설명 : Fnc함수명 으로 시작하는 모든 Function
'********************************************************************************************************* 
'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD     
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

	'-----------------------
	'Check previous data area
	'------------------------ 
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
    
	'-----------------------
	'Check condition area
	'----------------------- 
	If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
		Exit Function
	End If
    
	'-----------------------
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    Call InitVariables															'⊙: Initializes local global variables
    
    Call FncSetToolBar("New")
  '-----------------------
    'Query function call area
    '----------------------- 
	frm1.hCommand.value = "LOOKUP"
    Call DbQuery
   
    Set gActiveElement = document.ActiveElement   
    FncQuery = True                                                               '☜: Processing is OK    
        
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          '⊙: Processing is NG
    
  '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"X","X")           '⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If       
    
    Call ggoOper.ClearField(Document, "1")                                      '⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                      '⊙: Clear Contents  Field
    
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
    
    Call ggoOper.FormatNumber(frm1.txtPrRdpPerd, "99999", "1", False)					'상환주기 
	Call ggoOper.FormatNumber(frm1.txtIntPayPerd, "99999", "1", False)					'이자지급주기 
	Call ggoOper.FormatNumber(frm1.txtIntRate, "99.999999", "0", False, 6)			'이자율		
	
    Call ggoOper.LockField(Document, "N")                                       '⊙: Lock  Suitable  Field

    frm1.cboRdpClsFg.value = "N"
    Call cboPrRdpCond_OnChange()
    Call cboIntPayStnd_OnChange()
    
	Call FncSetToolBar("New")
	Call SetDefaultVal()
	Call InitVariables															'⊙: Initializes local global variables
	
	frm1.txtLoanNo.focus
	Set gActiveElement = document.activeElement
	
    FncNew = True																'⊙: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete()
	Dim intRetCD
	    
    FncDelete = False														'⊙: Processing is NG
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                
        Exit Function
    End If
    
	IF Trim(frm1.txtLoanNo.value) <> Trim(lgLoanNo) Then
		Call DisplayMsgBox("900002","X","X","X")                                
		Exit Function
	End If
    
  '-----------------------
    'Delete function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900003",parent.VB_YES_NO,"X","X")
    IF IntRetCD = vbNO Then
		Exit Function
	End IF	
    
    Call DbDelete															'☜: Delete db data
    
    FncDelete = True                                                        '⊙: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
    Dim IntRetCD 

    FncSave = False                                                         '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    
    '-----------------------
    'Precheck area
    '-----------------------
	' key data is changed
    If lgIntFlgMode = parent.OPMD_UMODE Then
		IF Trim(frm1.txtLoanNo.value) <> Trim(lgLoanNo) Then
			Call DisplayMsgBox("900002","X","X","X")                                
			Exit Function
		End If
    End If

    If lgBlnFlgChgValue = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")                    '⊙: No data changed!!
        Exit Function
    End If    
    
  '-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "2") Then                             '⊙: Check contents area
       Exit Function
    End If
    
    If frm1.txtDocCur.value =  parent.gCurrency Then
		frm1.txtXchRate.text = 1
    End If 
  '-----------------------
    'Save function call area
    '-----------------------
    CAll DbSave				                                                '☜: Save db data
    
    FncSave = True                                                          '⊙: Processing is OK

End Function

'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
''FINE_20030725_HC_Copy기능_START
	Call InitVariablesForCopy()

	lgBlnFlgChgValue = True
	lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
''FINE_20030725_HC_Copy기능_END
End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
' Keep : This Function is related to multi but single
'========================================================================================================
Function FncCancel() 
    FncCancel = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCancel = True                                                             '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
' Keep : This Function is related to multi but single
'========================================================================================================
Function FncInsertRow()
    FncInsertRow = False														 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncInsertRow = True                                                          '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
' Keep : This Function is related to multi but single
'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows

    FncDeleteRow = False														 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncDeleteRow = True                                                          '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
'========================================================================================================
Function FncPrint()
    FncPrint = False	                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call Parent.FncPrint()                                                       '☜: Protect system from crashing
    FncPrint = True	                                                             '☜: Processing is OK
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
    Dim IntRetCD     
    
    FncPrev = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

   '-----------------------
    'Query First
    '------------------------ 
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
		Call DisplayMsgBox("900002","X","X","X")
		Exit Function
    End If
    
   '-----------------------
    'Check previous data area
    '------------------------ 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
  '-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    
  '-----------------------
    'Erase contents area
    '----------------------- 
    Call InitVariables															'⊙: Initializes local global variables
    
  '-----------------------
    'Query function call area
    '----------------------- 
	frm1.hCommand.value = "PREV"
    Call DbQuery																'☜: Query db data

    Set gActiveElement = document.ActiveElement   	
    FncPrev = True                                                               '☜: Processing is OK   
    
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    Dim IntRetCD     
    
    FncNext = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

   '-----------------------
    'Query First
    '------------------------ 
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
		Call DisplayMsgBox("900002","X","X","X")
		Exit Function
    End If
    
   '-----------------------
    'Check previous data area
    '------------------------ 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
  '-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    
  '-----------------------
    'Erase contents area
    '----------------------- 
    Call InitVariables															'⊙: Initializes local global variables
    
  '-----------------------
    'Query function call area
    '----------------------- 
	frm1.hCommand.value = "NEXT"
    Call DbQuery																'☜: Query db data    

    Set gActiveElement = document.ActiveElement   	
    FncNext = True                                                               '☜: Processing is OK
    
End Function

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
    FncExcel = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncExport(Parent.C_SINGLE)

    FncExcel = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncFind(parent.C_SINGLE, True)

    FncFind = True                                                               '☜: Processing is OK
End Function


'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

    FncExit = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")	                 '⊙: Data is changed.  Do you want to exit? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If
    FncExit = True                                                               '☜: Processing is OK

End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow() 
     On Error Resume Next                                                   '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	Call parent.FncPrint()
End Function

'*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'********************************************************************************************************* 

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
	Call LayerShowHide(1)
    Err.Clear                                                               '☜: Protect system from crashing
    
    DbDelete = False														'⊙: Processing is NG
    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtLoanNo="   & Trim(lgLoanNo)		'☜: 삭제 조건 데이타 
    
    
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	
    DbDelete = True                                                         '⊙: Processing is NG

End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================
Function DbDeleteOk()														'☆: 삭제 성공후 실행 로직 
	Call FncNew()
End Function

'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()

	With frm1
		ggoOper.FormatFieldByObjectOfCur .txtLoanAmt,	  .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec		
		ggoOper.FormatFieldByObjectOfCur .txtLoanAmt,	  .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtBasRdpAmt,	  .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtBasIntPayAmt,.txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec		
		ggoOper.FormatFieldByObjectOfCur .txtRdpAmt,	  .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtIntPayAmt,   .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtLoanBalAmt,  .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
	End With
	
End Sub
'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    Dim strVal
        
    Call LayerShowHide(1)                                                        '☜: Show Processing Message
    Err.Clear																	'☜: Protect system from crashing
    
    DbQuery = False																'⊙: Processing is NG
    
    Call DisableToolBar(parent.TBC_QUERY)                                               '☜: Disable Query Button Of ToolBar    
        
    strVal = BIZ_PGM_ID & "?txtMode		=" & parent.UID_M0001						'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtCommand		=" & Trim(frm1.hCommand.value)		'☆: 조회 조건 데이타 
    strVal = strVal & "&txtLoanNo		=" & Trim(frm1.txtLoanNo.value)  	'☆: 조회 조건 데이타 
	strVal = strVal & "&txtLoanPlcType	=" & "BK"						 	'☆: 조회 조건 데이타 
    strVal = strVal & "&txtLoanBasicFg	=" & "LT"							'☆: 조회 조건 데이타 
	
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동	
    
    DbQuery = True   
    Set gActiveElement = document.ActiveElement   
    
End Function
  
'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 
	
    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field

    Call cboPrRdpCond_OnChange()
	Call cboIntPayStnd_Change

    Call CurFormatNumericOCX()        

    Call InitVariables

	lgLoanNo = frm1.txtLoanNo.value
    
    lgIntFlgMode = parent.OPMD_UMODE
    tempstrfg  = frm1.txtStrFg.Value												'⊙: Indicates that current mode is Update mode

    Call FncSetToolBar("Query")
    
    frm1.txtLoanNo.focus
    Set gActiveElement = document.activeElement 
    
End Function

'========================================================================================
' Function Name : DBSave
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'========================================================================================
Function DbSave() 
	Call LayerShowHide(1)
    Err.Clear																'☜: Protect system from crashing

	DbSave = False															'⊙: Processing is NG        

	With frm1
		.txtMode.value = parent.UID_M0002											'☜: 비지니스 처리 ASP 의 상태 
		.txtFlgMode.value = lgIntFlgMode
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID
		.txtstrFg.value = tempstrFg
		.txtloanbasicFg.value = "LT"
		.htxtLoanPlcType.value = "BK"
		
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)
	End With

    DbSave = True                                                           '⊙: Processing is NG

End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk(byval pLoanNo)															'☆: 저장 성공후 실행 로직	      
    '-----------------------
    'Reset variables area
    '-----------------------
     Select Case lgIntFlgMode
		Case parent.OPMD_CMODE
			' 신규시 FncQuery를 위하여 RectNo를 넘겨줌 
			frm1.txtLoanNo.value = pLoanNo
    End Select 
    
    Call InitVariables
    Call MainQuery
    
End Function

'==========================================================
'툴바버튼 세팅 
'==========================================================
Function FncSetToolBar(Cond)
	Select Case UCase(Cond)
	Case "NEW"
		Call SetToolbar("1110100000001111")
	Case "QUERY"
''FINE_20030725_HC_Copy기능_START
		Call SetToolbar("1111100011111111")
''FINE_20030725_HC_Copy기능_END
	End Select
End Function

'==========================================================
'상환전개버튼 클릭 
'==========================================================
Function FnButtonExec()
	Dim intRetCD
    Dim strVal

    If lgIntFlgMode <> parent.OPMD_UMODE Then				'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")	'조회를 먼저 하십시오.
        Exit Function
    End If
    
	IF Trim(frm1.txtLoanNo.value) <> Trim(lgLoanNo) Then
		Call DisplayMsgBox("900002","X","X","X")                                
		Exit Function
	End If
    
	'-----------------------
	'Check condition area
	'----------------------- 
	If Not chkField(Document, "1") Then					'⊙: This function check indispensable field
		Exit Function
	End If
    
	'-----------------------
	'Check previous data area
	'------------------------ 
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")	'데이터가 변경되었습니다. 계속하시겠습니까?
		If IntRetCD = vbNo Then
			Exit Function
		End If
	Else
		IntRetCD = DisplayMsgBox("900018",parent.VB_YES_NO,"X","X")	'작업을 수행하시겠습니까?
		If IntRetCD = vbNO Then
			Exit Function
		End If
	End If
    
    Call LayerShowHide(1)
    
    Err.Clear                                                               '☜: Protect system from crashing
    
    strVal = BIZ_PGM_ID & "?txtMode=" & "PAFG400"							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtLoanNo=" & Trim(lgLoanNo)  		    '☆: 조회 조건 데이타 
    strVal = strVal & "&txtDateFr=" & Trim(frm1.txtLoanDt.text) 
    strVal = strVal & "&txtDateTo=" & Trim(frm1.txtLoanDt.text)
    
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	
End Function

'***************************************************************************************************************

Const TAB1 = 1																		'☜: Tab의 위치 
Const TAB2 = 2
Const TAB3 = 3

Function ClickTab1()
	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)
	gSelframeFlg = TAB1
End Function


</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc"  -->	
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
// -->
</SCRIPT>

</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>은행기초차입금등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT><A HREF="VBSCRIPT:OpenPopupTempGL()">결의전표</A> |
										    <A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100% COLSPAN=2>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>차입금번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtLoanNo" SIZE="18" MAXLENGTH="18" tag="12XXXU" ALT="차입금번호" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoanNo" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopupLoan()"></TD>
									<TD CLASS="TD6" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD6" NOWRAP>&nbsp;</TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>차입내역</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtLoanNm" SIZE="40" MAXLENGTH="40" tag="22X" ALT="차입금내역"></TD>
									<TD CLASS="TD5" NOWRAP>부서</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDeptCd" ALT="부서코드" Size= "10" MAXLENGTH="10"  tag="22XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopUp(frm1.txtDeptCd.value, 2)">
														   <INPUT NAME="txtDeptNm" ALT="부서명" SIZE = "20" STYLE="TEXT-ALIGN: left" tag="24X"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>차입일</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f4204ma1_fpLoanDt_txtLoanDt.js'></script></TD>
									<TD CLASS="TD5" NOWRAP>기초등록일</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f4204ma1_fpBasicLoanDt_txtBasicLoanDt.js'></script></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>장단기구분</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboLoanFg" ALT="장단기구분" STYLE="WIDTH: 135px" tag="22X" OnClick ="vbscript:Type_itemChange()"><OPTION VALUE=""></OPTION></SELECT></TD>
									<TD CLASS="TD5" NOWRAP>차입금계정</TD>												
								    <TD CLASS="TD6" NOWRAP><INPUT NAME="txtLoanAcctCd" ALT="차입금계정" SIZE="10" MAXLENGTH="20"  tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoanAcct" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtLoanAcctCd.value, 7)">
														   <INPUT NAME="txtLoanAcctNm" ALT="차입금계정명" SIZE="20" tag="24X"></TD>									
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>차입용도</TD>												
								    <TD CLASS="TD6" NOWRAP><INPUT NAME="txtLoanType" ALT="차입용도" SIZE="10" MAXLENGTH="2" tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoanType" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtLoanType.value, 6)">
														   <INPUT NAME="txtLoanTypeNm" ALT="차입용도명" STYLE="TEXT-ALIGN: Left" tag="24X"></TD>
									<TD CLASS="TD5" NOWRAP>차입은행</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBankLoanCd" SIZE="10" MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="22X" ALT="차입은행"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankLoanCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopUp(frm1.txtBankLoanCd.value, 5)">
									                       <INPUT TYPE=TEXT NAME="txtBankLoanNm" ALT="차입은행명" SIZE=20 tag="24X"></TD>																		
								</TR>		
								<TR>
									<TD CLASS="TD5" NOWRAP>거래통화|환율</TD>
								    <TD CLASS="TD6" NOWRAP><INPUT NAME="txtDocCur" ALT="거래통화" SIZE = "10" MAXLENGTH="3"  tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCurCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtDocCur.value, 0)">&nbsp;
															<script language =javascript src='./js/f4204ma1_OBJECT5_txtXchRate.js'></script></TD>
									<TD CLASS="TD5" NOWRAP>차입금액|자국</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f4204ma1_fpLoanAmt_txtLoanAmt.js'></script>&nbsp;
														   <script language =javascript src='./js/f4204ma1_fpLoanLocAmt_txtLoanLocAmt.js'></script></TD>									
								</TR>	
								<TR>
									<TD CLASS="TD5" NOWRAP>이자지급기초액|자국</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f4204ma1_fpBasIntPayAmt_txtBasIntPayAmt.js'></script>&nbsp;
														   <script language =javascript src='./js/f4204ma1_fpBasIntPayLocAmt_txtBasIntPayLocAmt.js'></script></TD>
									<TD CLASS="TD5" NOWRAP>원금상환기초액|자국</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f4204ma1_fpBasRdpAmt_txtBasRdpAmt.js'></script>&nbsp;
														   <script language =javascript src='./js/f4204ma1_fpBasRdpLocAmt_txtBasRdpLocAmt.js'></script></TD>									
								</TR>
							<TR>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>원금상환방법</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboPrRdpCond" ALT="원금상환방법" STYLE="WIDTH: 135px" tag="22X" OnClick ="vbscript:Type_itemChange()"><OPTION VALUE=""></OPTION></SELECT></TD>
									<TD CLASS="TD5" NOWRAP>최초원금상환일</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f4204ma1_fp1StPrRdpDt_txt1StPrRdpDt.js'></script></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>원금상환주기</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f4204ma1_fpPrRdpPerd_txtPrRdpPerd.js'></script>개월</TD>
									<TD CLASS="TD5" NOWRAP>상환만기일</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f4204ma1_fpDueDt_txtDueDt.js'></script></TD>
								</TR>								
								<TR>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>이자율변동성</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio_IntVotl ID=Rb_IntVotl1 Checked tag = 2 value="X" onclick=radio1_onchange()><LABEL FOR=Rb_IntVotl1>확정</LABEL>&nbsp;
														   <INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio_IntVotl ID=Rb_IntVotl2 tag = 2 value="F" onclick=radio2_onchange()><LABEL FOR=Rb_IntVotl2>변동</LABEL>&nbsp;</TD>
									<TD CLASS="TD5" NOWRAP>이자율</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f4204ma1_OBJECT5_txtIntRate.js'></script>&nbsp;%/년</TD>
								</TR>								
								<TR>
									<TD CLASS="TD5" NOWRAP>이자지급형태</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboIntPayStnd" ALT="이자지급형태" STYLE="WIDTH: 135px" tag="22X" OnClick ="vbscript:Type_itemChange()"><OPTION VALUE=""></OPTION></SELECT></TD>
									<TD CLASS="TD5" NOWRAP>이자계정</TD>
								    <TD CLASS="TD6" NOWRAP><INPUT NAME="txtIntAcctCd" ALT="이자계정" SIZE="10" MAXLENGTH="20"  tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnIntAcct" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtIntAcctCd.value, 8)">
																		  <INPUT NAME="txtIntAcctNm" ALT="이자계정명" SIZE="20" tag="24X"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>이자지급주기</TD>
								    <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f4204ma1_fpIntPayPerd_txtIntPayPerd.js'></script>&nbsp;/&nbsp;<SELECT NAME="cboIntBaseMthd" ALT="이자계산방법" STYLE="WIDTH: 135px" tag="22X" OnClick ="vbscript:Type_itemChange()"><OPTION VALUE=""></OPTION></SELECT></TD>
								    <TD CLASS="TD5" NOWRAP>최초이자지급일</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f4204ma1_fp1StIntDueDT_txt1StIntDueDT.js'></script></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>사용자필드1</TD>									
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtUserFld1" SIZE="40" MAXLENGTH="50" tag="21X" ALT="사용자필드"></TD>									
									<TD CLASS="TD5" NOWRAP>사용자필드2</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtUserFld2"  SIZE="40" MAXLENGTH="50" tag="21X" ALT="사용자필드2"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>비고</TD>
									<TD CLASS="TD6" COLSPAN=3 NOWRAP><INPUT TYPE=TEXT NAME="txtLoanDesc" SIZE="80" MAXLENGTH="128" tag="21X" ALT="비고"></TD>													  									
								</TR>
								<TR>
								</TR>
								<TR>	
									<TD CLASS="TD5" NOWRAP>원금상환총액|자국</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f4204ma1_fpPrRdpAmt_txtRdpAmt.js'></script>&nbsp;
														   <script language =javascript src='./js/f4204ma1_fpPrRdpLocAmt_txtRdpLocAmt.js'></script></TD>									
									<TD CLASS="TD5" NOWRAP>이자지급총액|자국</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f4204ma1_fpIntPayAmt_txtIntPayAmt.js'></script>&nbsp;
														   <script language =javascript src='./js/f4204ma1_fpIntPayLocAmt_txtIntPayLocAmt.js'></script></TD>																				
								</TR>								
								<TR>									
									<TD CLASS="TD5" NOWRAP>차입잔액|자국</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f4204ma1_fpLoanBalAmt_txtLoanBalAmt.js'></script>&nbsp;
														   <script language =javascript src='./js/f4204ma1_fpLoanBalLocAmt_txtLoanBalLocAmt.js'></script></TD>									
									<TD CLASS="TD5" NOWRAP>상환완료여부</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboRdpClsFg" ALT="상환완료여부" STYLE="WIDTH: 135px" tag="24X"><OPTION VALUE=""> </OPTION></SELECT></TD>                     									
								</TR>								
								<TR>
								</TR>																
								<INPUT TYPE=hidden CLASS="Radio" NAME=Radio_Cur ID=hRb_Cur1  onclick=radio9_onchange()><LABEL FOR=Rb_Cur1>
								<INPUT TYPE=hidden CLASS="Radio" NAME=Radio_Cur ID=hRb_Cur2 onclick=radio10_onchange()><LABEL FOR=Rb_Cur2>																
								<INPUT TYPE=hidden NAME="htxtLcNo" tag="24">
								<INPUT TYPE=hidden NAME="htxtLoanPlcType" tag="24">
								<INPUT TYPE=hidden NAME="hClsRoFg" tag="24">								
								<INPUT TYPE=hidden NAME="htxtStIntPayAmt" tag="24">
								<INPUT TYPE=hidden NAME="htxtStIntPayLocAmt" tag="24">																						
								<INPUT TYPE=hidden NAME="hRdpSprdFg" tag="24">								
								<INPUT TYPE=hidden NAME="txtTempGlNo" tag="24">		
								<INPUT TYPE=hidden NAME="txtGlNo" tag="24">			
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
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<BUTTON NAME="btnExec" CLASS="CLSMBTN" OnClick="VBScript:Call FnButtonExec()" Flag=1>상환전개</BUTTON>&nbsp;
					</TD>
					<TD WIDTH=* ALIGN=RIGHT>
					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=hidden NAME="txtMode" tag="24">
<INPUT TYPE=hidden NAME="hCommand" tag="24">
<INPUT TYPE=hidden NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=hidden NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=hidden NAME="txtMaxRows" tag="24">
<INPUT TYPE=hidden NAME="txtFlgMode" tag="24">
<INPUT TYPE=hidden NAME="hOrgChangeId" tag="24">
<INPUT TYPE=hidden NAME="txtstrFg" tag="24">
<INPUT TYPE=hidden NAME="txtPrRdpUnitAmt" tag="24">
<INPUT TYPE=hidden NAME="txtPrRdpUnitLocAmt" tag="24">
<INPUT TYPE=hidden NAME="txtLoanBasicFg" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
