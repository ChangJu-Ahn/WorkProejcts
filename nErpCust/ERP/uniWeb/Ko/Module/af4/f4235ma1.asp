<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Long Loan
'*  3. Program ID           : f4235ma1
'*  4. Program Name         : 거래처차입금만기연장 
'*  5. Program Desc         : Register of Loan Master
'*  6. Comproxy List        : FL0069, FL0061
'*  7. Modified date(First) : 2002/03/29
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Oh, Soo Min
'* 10. Modifier (Last)      : 
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
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<!--'==========================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	  SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>


Option Explicit                                                             '☜: indicates that All variables must be declared in advance 

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
<!-- #Include file="../../inc/lgvariables.inc" -->	
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Const BIZ_PGM_ID = "f4235mb1.asp"			 '☆: 비지니스 로직 ASP명 
Const BIZ_PGM_ID2 = "f4235mb3.asp"

Const JUMP_PGM_ID_LOAN_CHG = "f4231ma1"		 '차입금변경등록 
Const JUMP_PGM_ID_LOAN_REP = "f4250ma1"		 '차입금상환등록 
 											
 '==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 
Dim lgtempStrFg

Dim lgLoanRoNo
Dim lgNextNo						'☜: 화면이 Single/SingleMulti 인경우만 해당 
Dim lgPrevNo						' ""
Dim PGM
Dim strDiffDate    
Dim strDiffYr
Dim strDiffMnth

'==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= 
'----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 

Dim IsOpenPop

dim dtToday
dtToday = "<%=GetSvrDate%>"

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 
 
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
    lgIntFlgMode = Parent.OPMD_CMODE                                               '⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                                                '⊙: Indicates that no value changed
    lgIntGrpCount = 0                                                       '⊙: Initializes Group View Size
    frm1.hOrgChangeId.value = parent.gChangeOrgId

    IsOpenPop = False														'☆: 사용자 변수 초기화 
End Sub

'=============================== 2.1.2 LoadInfTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

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
	With frm1
		.txtLoanRoDt.text = UniConvDateAToB("<%=dtToday%>",Parent.gServerDateFormat,Parent.gDateFormat)			'차입일 
		.txtDueDt.text = UniConvDateAToB("<%=dtToday%>",Parent.gServerDateFormat,Parent.gDateFormat)				'상환만기일 
		
		.Rb_IntVotl1.Checked = True	
		.Rb_IntStart1.Checked = True	
		.Rb_IntEnd2.Checked = True	
		.hRb_Cur1.value = "1"			
		.htxtPrRdpUnitAmt.value = "0"
		.htxtPrRdpUnitLocAmt.value = "0"	

		.txtDocCur.value = parent.gCurrency
	End With

	lgBlnFlgChgValue = False
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
	Dim arrParam(8), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	frm1.hOrgChangeId.value = Parent.gChangeOrgId

	Select Case iWhere
		Case 2
			arrParam(0) = strCode		            '  Code Condition
		   	arrParam(1) = frm1.txtLoanRoDt.Text
			arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  
			arrParam(3) = "F"									' 결의일자 상태 Condition  

			' 권한관리 추가 
			arrParam(5) = lgAuthBizAreaCd
			arrParam(6) = lgInternalCd
			arrParam(7) = lgSubInternalCd
			arrParam(8) = lgAuthUsrID
			
		Case 3			
			If frm1.txtBankCd.className = Parent.UCN_PROTECTED Then Exit Function		
			
			arrParam(0) = frm1.txtBankCd.Alt										' 팝업 명칭 
			arrParam(1) = "B_BANK A, B_BANK_ACCT B, F_DPST C"						' TABLE 명칭 
			arrParam(2) = strCode													' Code Condition
			arrParam(3) = ""														' Name Cindition
			arrParam(4) = "A.BANK_CD = B.BANK_CD "									' Where Condition			
			arrParam(4) = arrParam(4) & "AND B.BANK_ACCT_NO = C.BANK_ACCT_NO "
			arrParam(4) = arrParam(4) & "AND A.BANK_CD = C.BANK_CD "												
		   'arrParam(4) = arrParam(4) & "AND C.DOC_CUR = Parent.gCurrency "		
			
			arrParam(5) = frm1.txtBankCd.Alt										' 조건필드의 라벨 명칭 

			arrField(0) = "A.BANK_CD"						' Field명(0)
			arrField(1) = "A.BANK_NM"						' Field명(1)
			arrField(2) = "B.BANK_ACCT_NO"					' Field명(2)
'			arrField(3) = "C.DOC_CUR"						' Field명(3)
    
			arrHeader(0) = "은행코드"					' Header명(0)
			arrHeader(1) = "은행명"						' Header명(1)
			arrHeader(2) = "계좌번호"					' Header명(2)
'			arrHeader(3) = "거래통화"					' Header명(3)										
		Case 4			
			If frm1.txtBankAcctNo.className = Parent.UCN_PROTECTED Then Exit Function		
			
			arrParam(0) = frm1.txtBankAcctNo.Alt								' 팝업 명칭 
			arrParam(1) = "B_BANK A, B_BANK_ACCT B, F_DPST C"				' TABLE 명칭 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = "A.BANK_CD = B.BANK_CD "							' Where Condition
			arrParam(4) = arrParam(4) & "AND B.BANK_ACCT_NO = C.BANK_ACCT_NO "	
			arrParam(4) = arrParam(4) & "AND A.BANK_CD = C.BANK_CD "												
		   'arrParam(4) = arrParam(4) & "AND C.DOC_CUR = Parent.gCurrency "		
					
			arrParam(5) = frm1.txtBankAcctNo.Alt								' 조건필드의 라벨 명칭 

			arrField(0) = "B.BANK_ACCT_NO"					' Field명(0)
			arrField(1) = "A.BANK_CD"						' Field명(1)
			arrField(2) = "A.BANK_NM"						' Field명(2)
'			arrField(3) = "C.DOC_CUR"						' Field명(3)
    
			arrHeader(0) = "계좌번호"					' Header명(0)
			arrHeader(1) = "은행코드"					' Header명(1)
			arrHeader(2) = "은행명"						' Header명(2)			
'			arrHeader(3) = "거래통화"					' Header명(3)										
		Case 5		'차입처 
			If frm1.txtBpRoCd.className = Parent.UCN_PROTECTED Then Exit Function
			lgtempStrFg = "B"
			arrParam(0) = frm1.txtBpRoCd.Alt
			arrParam(1) = "B_BIZ_PARTNER A"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = frm1.txtBpRoCd.Alt
	
			arrField(0) = "A.BP_CD" 
			arrField(1) = "A.BP_NM"
				    
			arrHeader(0) = "거래처코드"
			arrHeader(1) = "거래처명"			
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
		Case 7		'출금유형 
			If frm1.txtRcptType.className = Parent.UCN_PROTECTED Then Exit Function

			arrParam(0) = frm1.txtRcptType.Alt
			arrParam(1) = "B_MINOR A, B_CONFIGURATION B"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = "A.MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  " 
			arrParam(4) = arrParam(4) & " AND A.MAJOR_CD = B.MAJOR_CD AND A.MINOR_CD = B.MINOR_CD AND B.SEQ_NO = 4 AND B.REFERENCE IN ( " & FilterVar("DP", "''", "S") & " ," & FilterVar("CS", "''", "S") & " ," & FilterVar("CK", "''", "S") & " ," & FilterVar("FO", "''", "S") & " ) "
			arrParam(5) = frm1.txtRcptType.Alt
	
			arrField(0) = "A.MINOR_CD"
			arrField(1) = "A.MINOR_NM"
			    
			arrHeader(0) = frm1.txtRcptType.Alt
			arrHeader(1) = frm1.txtRcptTypeNm.Alt
		Case 8
			If frm1.txtLoanAcctCd.className = "protected" Then Exit Function    
			
			arrParam(0) = "만기연장계정팝업"								' 팝업 명칭 
			arrParam(1) = "A_ACCT	A,A_ACCT_GP 	B, A_JNL_ACCT_ASSN	C	"				' TABLE 명칭 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = "C.TRANS_TYPE = " & FilterVar("FI007", "''", "S") & "  " 			' Where Condition
			arrParam(4) = arrParam(4) & " AND 	C.ACCT_CD = A.ACCT_CD AND A.GP_CD = B.GP_CD "
			arrParam(4) = arrParam(4) & " AND 	A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND 	DR_CR_FG = " & FilterVar("CR", "''", "S") & "  " 
			arrParam(4) = arrParam(4) & " AND JNL_CD = " & FilterVar(frm1.cboLoanFg.Value & "_RO", "''", "S") 
			arrParam(5) = frm1.txtLoanAcctCd.Alt							' 조건필드의 라벨 명칭 

			arrField(0) = "A.ACCT_CD"									' Field명(0)
			arrField(1) = "A.ACCT_NM"									' Field명(1)
			arrField(2) = "B.GP_CD"										' Field명(2)
			arrField(3) = "B.GP_NM"					 					' Field명(3)
			
			arrHeader(0) = frm1.txtLoanAcctCd.Alt							' Header명(0)
			arrHeader(1) = frm1.txtLoanAcctNm.Alt						' Header명(1)
			arrHeader(2) = "그룹코드"									' Header명(2)
			arrHeader(3) = "그룹명"										' Header명(3)									
		Case 9
			If frm1.txtChargeAcctCd.className = "protected" Then Exit Function    
			
			arrParam(0) = "부대비용계정팝업"								' 팝업 명칭 
			arrParam(1) = "A_ACCT	A,A_ACCT_GP 	B, A_JNL_ACCT_ASSN	C	"				' TABLE 명칭 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""													' Name Cindition
			arrParam(4) = "C.TRANS_TYPE = " & FilterVar("FI007", "''", "S") & "  " 					' Where Condition
			arrParam(4) = arrParam(4) & " AND 	C.ACCT_CD = A.ACCT_CD AND A.GP_CD = B.GP_CD "
			arrParam(4) = arrParam(4) & " AND 	A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND 	DR_CR_FG = " & FilterVar("DR", "''", "S") & "  " 
			arrParam(4) = arrParam(4) & " AND JNL_CD = " & FilterVar("BC", "''", "S") & "  " 
			arrParam(5) = frm1.txtChargeAcctCd.Alt							' 조건필드의 라벨 명칭 

			arrField(0) = "A.Acct_CD"									' Field명(0)
			arrField(1) = "A.Acct_NM"									' Field명(1)
			arrField(2) = "B.GP_CD"										' Field명(2)
			arrField(3) = "B.GP_NM"										' Field명(3)
			
			arrHeader(0) = frm1.txtChargeAcctCd.Alt									' Header명(0)
			arrHeader(1) = frm1.txtChargeAcctNm.Alt								' Header명(1)
			arrHeader(2) = "그룹코드"									' Header명(2)
			arrHeader(3) = "그룹명"										' Header명(3)						
		Case 10
			If frm1.txtIntAcctCd.className = "protected" Then Exit Function    
			
			arrParam(0) = "이자계정팝업"								' 팝업 명칭 
			arrParam(1) = "A_ACCT	A,A_ACCT_GP 	B, A_JNL_ACCT_ASSN	C	"				' TABLE 명칭 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""													' Name Cindition
			arrParam(4) = "C.TRANS_TYPE = " & FilterVar("FI002", "''", "S") & "  " 					' Where Condition
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
		Case 11
			If frm1.txtRcptAcctCd.className = "protected" Then Exit Function    
			
			arrParam(0) = "출금계정팝업"								' 팝업 명칭 
			arrParam(1) = "A_ACCT	A,A_ACCT_GP 	B, A_JNL_ACCT_ASSN	C	"				' TABLE 명칭 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = "C.TRANS_TYPE = " & FilterVar("FI007", "''", "S") & "  " 			' Where Condition
			arrParam(4) = arrParam(4) & " AND 	C.ACCT_CD = A.ACCT_CD AND A.GP_CD = B.GP_CD "
			arrParam(4) = arrParam(4) & " AND 	A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND 	DR_CR_FG = " & FilterVar("CR", "''", "S") & "  " 
			arrParam(4) = arrParam(4) & " AND JNL_CD = " & FilterVar(frm1.txtRcptType.Value, "''", "S") 			
			arrParam(5) = frm1.txtRcptAcctCd.Alt							' 조건필드의 라벨 명칭 

			arrField(0) = "A.Acct_CD"									' Field명(0)
			arrField(1) = "A.Acct_NM"									' Field명(1)
			arrField(2) = "B.GP_CD"										' Field명(2)
			arrField(3) = "B.GP_NM"										' Field명(3)
			
			arrHeader(0) = frm1.txtRcptAcctCd.Alt									' Header명(0)
			arrHeader(1) = frm1.txtRcptAcctNm.Alt								' Header명(1)
			arrHeader(2) = "그룹코드"									' Header명(2)
			arrHeader(3) = "그룹명"										' Header명(3)
		Case 12
			If frm1.txtBPAcctCd.className = "protected" Then Exit Function    
			
			arrParam(0) = "수수료계정팝업"								' 팝업 명칭 
			arrParam(1) = "A_ACCT	A,A_ACCT_GP 	B, A_JNL_ACCT_ASSN	C	"				' TABLE 명칭 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""													' Name Cindition
			arrParam(4) = "C.TRANS_TYPE = " & FilterVar("FI007", "''", "S") & "  " 					' Where Condition
			arrParam(4) = arrParam(4) & " AND 	C.ACCT_CD = A.ACCT_CD AND A.GP_CD = B.GP_CD "
			arrParam(4) = arrParam(4) & " AND 	A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND 	DR_CR_FG = " & FilterVar("DR", "''", "S") & "  " 
			arrParam(4) = arrParam(4) & " AND JNL_CD = " & FilterVar("BP", "''", "S") & "  " 
			arrParam(5) = frm1.txtBPAcctCd.Alt							' 조건필드의 라벨 명칭 

			arrField(0) = "A.Acct_CD"									' Field명(0)
			arrField(1) = "A.Acct_NM"									' Field명(1)
			arrField(2) = "B.GP_CD"										' Field명(2)
			arrField(3) = "B.GP_NM"										' Field명(3)
			
			arrHeader(0) = frm1.txtBPAcctCd.Alt									' Header명(0)
			arrHeader(1) = frm1.txtBPAcctNm.Alt								' Header명(1)
			arrHeader(2) = "그룹코드"									' Header명(2)
			arrHeader(3) = "그룹명"	

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
		frm1.txtLoanRoNo.focus
		Exit Function
	Else
		Select Case iWhere

		Case 2	'부서 
            frm1.txtDeptCd.value = arrRet(0)
            frm1.txtDeptNm.value = arrRet(1)
			frm1.txtLoanRoDt.text = arrRet(3)
			call txtDeptCd_OnChange()
			frm1.txtDeptCd.focus
		Case 3	'은행 
			frm1.txtBankCd.value	= arrRet(0)
			frm1.txtBankNm.value	= arrRet(1)
			frm1.txtBankAcctNo.value  = arrRet(2)			
			frm1.txtBankCd.focus
		Case 4	'계좌번호 
			frm1.txtBankAcctNo.value  = arrRet(0)
			frm1.txtBankCd.value	= arrRet(1)
			frm1.txtBankNm.value	= arrRet(2)		
			frm1.txtBankAcctNo.focus
		Case 5	'차입은행 
			frm1.txtBpRoCd.value = arrRet(0)
			frm1.txtBpRoNm.value = arrRet(1)
			frm1.txtBpRoCd.focus
		Case 6	'차입용도 
			frm1.txtLoanType.value = arrRet(0)
			frm1.txtLoanTypeNm.value = arrRet(1)
			frm1.txtLoanType.focus
		Case 7	'출금유형 
			frm1.txtRcptType.value = arrRet(0)
			frm1.txtRcptTypeNm.value = arrRet(1)
			Call txtRcptType_OnChange
			frm1.txtRcptType.focus
		Case 8  '만기연장계정코드 
			frm1.txtLoanAcctCd.value = arrRet(0)
			frm1.txtLoanAcctNm.value = arrRet(1)
			frm1.txtLoanAcctCd.focus
		Case 9
			frm1.txtChargeAcctCd.value = arrRet(0)
			frm1.txtChargeAcctNm.value = arrRet(1)
			frm1.txtChargeAcctCd.focus
		Case 10
			frm1.txtIntAcctCd.value = arrRet(0)
			frm1.txtIntAcctNm.value = arrRet(1)
			frm1.txtIntAcctCd.focus
		Case 11
			frm1.txtRcptAcctCd.value = arrRet(0)
			frm1.txtRcptAcctNm.value = arrRet(1)
			frm1.txtRcptAcctCd.focus
		Case 12'수수료계정코드 
			frm1.txtBPAcctCd.value = arrRet(0)
			frm1.txtBPAcctNm.value = arrRet(1)
			frm1.txtBPAcctCd.focus

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
	If frm1.txtDeptCd.className = Parent.UCN_PROTECTED Then Exit Function
	
	arrParam(0) = strCode						'부서코드 
	arrParam(1) = frm1.txtLoanRoDt.Text			'날짜(Default:현재일)
	arrParam(2) = "1"							'부서권한(lgUsrIntCd)
	
	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/DeptPopupDt.asp", Array(arrParam), _
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
'만기연장번호 팝업 
'============================================================
Function OpenPopupRoLoan()
	Dim arrRet
	Dim arrParam(8)	
	Dim iCalledAspName

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	
	iCalledAspName = AskPRAspName("f4205ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "f4205ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
    
	arrRet = window.showModalDialog(iCalledAspName & "?PGM=" & gStrRequestMenuID , Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = ""  Then			
		frm1.txtLoanRoNo.focus
		Exit Function
	Else		
		frm1.txtLoanRoNo.value = arrRet(0)
	End If
	
	frm1.txtLoanRoNo.focus
End Function

'============================================================
'차입금번호 팝업(default값 setting)
'============================================================
Function OpenPopupLoan()
	Dim arrRet
	Dim arrParam(8)	
	Dim iCalledAspName

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	
	iCalledAspName = AskPRAspName("f4234ra1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "f4234ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
    
	arrRet = window.showModalDialog(iCalledAspName & "?PGM=" & gStrRequestMenuID , Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	With frm1
		If arrRet(0) = ""  Then			
			.txtLoanRoNo.focus
			Exit Function
		Else
			Call ggoOper.ClearField(Document, "A")																									'⊙: Clear Contents  Field
			Call ggoOper.ClearField(Document, "N")																									'⊙: Clear Contents  Field
			.txtLoanNo.value = arrRet(0)																											'차입금번호 
			.txtBpRoCd.value = arrRet(33)																											'차입거래처코드 
			.txtBpRoNm.value = arrRet(34)																											'차입거래처명		
			.txtdeptCd.value = arrRet(13)																											'부서코드 
			.txtdeptNm.value = arrRet(14)																											'부서코드명					
			
			.txtLoanAmt.Text = arrRet(9)																											'Org차입금액			
			.txtLoanLocAmt.Text = arrRet(10)																										'Org차입금액(자국)			
			
			'Hidden field는 not serverformat -> UNIFormatNumber 필요 
			.txtIntPayAmt.Text = UNIFormatNumber(arrRet(28),ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)										'Org이자지급액						
			.txtIntPayLocAmt.Text = UNIFormatNumber(arrRet(31),ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)									'Org이자지급액(자국)			
			
			.txtLoanBalAmt.Text = arrRet(11)																										'Org 차입잔액 
			.txtLoanBalLocAmt.Text = arrRet(12)																										'Org 차입잔액(자국)
			
			.txtRdpAmt.Text = UNIFormatNumber(arrRet(29),ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)											'Org 원금상환액			
			.txtRdpLocAmt.Text = UNIFormatNumber(arrRet(32),ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)										'Org 원금상환액(자국)						
			.txtLoanRoDt.text = arrRet(7)																											'Rollover Date
			.cboLoanFg.value = arrRet(2)																											'차입금구분 
			.txtDueDt.text = arrRet(7)																												'Org 만기일 
			
			.txtLoanType.value = arrRet(4)																											'차입용도 
			.txtLoanTypeNm.value = arrRet(5)																										'차입용도명 
			.txtDocCur.value	= arrRet(8)																											'통화						
			.txtXchRate.Text	= UNIFormatNumber(arrRet(30),ggExchRate.DecPoint,-2,0,ggExchRate.RndPolicy,ggExchRate.RndUnit)						'이자율 
			
'			.txtLoanRoAmt.Text = UNIFormatNumber(UNICDbl(arrRet(9)) - arrRet(29),ggAmtOfMoney.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)			'Rollover 차입금액									
'			.txtLoanRoLocAmt.Text = UNIFormatNumber(UNICDbl(arrRet(10)) - arrRet(32),ggAmtOfMoney.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)		'Rollover 차입금액(자국)
			
			.txtLoanRoAmt.Text = UNIFormatNumber(UNICDbl(arrRet(9)) - UNICDbl(UNIFormatNumber(arrRet(29),ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)	),ggAmtOfMoney.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)			'Rollover 차입금액	
			.txtLoanRoLocAmt.Text = UNIFormatNumber(UNICDbl(arrRet(10)) - UNICDbl(UNIFormatNumber(arrRet(32),ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)	),ggAmtOfMoney.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)		'Rollover 차입금액(자국)
			
			.cboPrRdpCond.value = arrRet(18)																										'원금상환방법(CS/DP/CK)
			Call cboPrRdpCond_OnChange()
						
			If arrRet(23) = "X" Then																												'이율변동성 
				.Rb_IntVotl1.Checked = True
			Else 
				.Rb_IntVotl2.Checked = True
			End If
			
			.txtIntPayPerd.Text = arrRet(19)																										'이자지급주기			
			.txtIntRate.Text = UNIFormatNumber(arrRet(17),ggExchRate.DecPoint,-2,0,ggExchRate.RndPolicy,ggExchRate.RndUnit)							'이자율 
			.cboIntPayStnd.value = arrRet(21)																										'이자지급형태 
			Call cboIntPayStnd_OnChange()
			
			.cboIntBaseMthd.value = arrRet(22)																										'이자지급계산기준 
			
			If arrRet(24) = "YY" Then																												'일수계산방법(YY/YN/NY/NN)
				.Rb_IntStart1.Checked = True				
				.Rb_IntEnd1.Checked = True				
			ElseIf arrRet(24) = "YN" Then		
				.Rb_IntStart1.Checked = True
				.Rb_IntEnd2.Checked = True
			ElseIf arrRet(24) = "NY" Then		
				.Rb_IntStart2.Checked = True
				.Rb_IntEnd1.Checked = True
			Else 
				.Rb_IntStart2.Checked = True
				.Rb_IntEnd2.Checked = True
			End If 								
			
			.htxtOrgLoanRcptType.value = arrRet(25)												'출금유형 
			.htxtOrgLoanRcptAcctCd.value = arrRet(35)											'출금계정코드						
			.htxtOrgLoanBankAcctNo.value = arrRet(27)											'입금계좌번호						
			.htxtOrgLoanBankCd.value = arrRet(26)												'입금은행 

			lgIntFlgMode = Parent.OPMD_CMODE                                               '⊙: Indicates that current mode is Create mode
		End If
	End With
	
	Call CurFormatNumericOCX()	
	Call FncSetToolBar("REF")
	
	frm1.txtLoanRoNo.focus
End Function

'============================================================
'회계전표 팝업 
'============================================================
Function OpenPopupGL()

	Dim arrRet
	Dim arrParam(8)	
	Dim iCalledAspName
	
	iCalledAspName = AskPRAspName("a5120ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function
	
	arrParam(0) = Trim(frm1.htxtGlNo.value)	'회계전표번호 
	arrParam(1) = ""						'Reference번호 

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID

	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	frm1.txtLoanRoNo.focus
End Function

'============================================================
'결의전표 팝업 
'============================================================
Function OpenPopupTempGL()

	Dim arrRet
	Dim arrParam(8)	
	Dim iCalledAspName
	
	iCalledAspName = AskPRAspName("a5130ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function
	
	arrParam(0) = Trim(frm1.htxtTempGlNo.value)	'결의전표번호 
	arrParam(1) = ""							'Reference번호 

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID

	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	frm1.txtLoanRoNo.focus
	
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
					
		frm1.txtLoanRoNo.value = strTemp
				
		If Err.number <> 0 Then
			Err.Clear
			Call WriteCookie("LOAN_NO", "")
			Exit Function 
		End If
				
		Call MainQuery()
	
	Case JUMP_PGM_ID_LOAN_CHG
		Call WriteCookie("LOAN_NO", frm1.txtLoanRoNo.value)
	
	Case JUMP_PGM_ID_LOAN_REP
		Call WriteCookie("LOAN_NO", frm1.txtLoanRoNo.value)
	
	Case Else
		Exit Function
	End Select
End Function	

'========================================================================================================
'	Desc : 화면이동 
'========================================================================================================
Function PgmJumpChk(strPgmId)
	Dim IntRetCD
	
	'-----------------------
	'Check previous data area
	'------------------------ 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call CookiePage(strPgmId)
    Call PgmJump(strPgmId)
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
'=======================================================================================================
'   Event Name : _DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtDueDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDueDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtDueDt.Focus
    End If
End Sub

Sub txt1StIntDueDt_DblClick(Button)
    If Button = 1 Then
        frm1.txt1StIntDueDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txt1StIntDueDt.Focus
    End If
End Sub

Sub txt1StPrRdpDt_DblClick(Button)
    If Button = 1 Then
        frm1.txt1StPrRdpDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txt1StPrRdpDt.Focus
    End If
End Sub

Sub txtLoanRoDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtLoanRoDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtLoanRoDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : cboIntPayStnd_OnChange()
'   Event Desc : 이자지급형태별 Set Protected/Required Fields
'=======================================================================================================
Sub cboIntPayStnd_Change()
	'이자지급형태 
	Select Case frm1.cboIntPayStnd.value
	Case "AI"	'선급		
		frm1.txt1StIntDueDT.Text = ""
		Call ggoOper.SetReqAttr(frm1.txt1StIntDueDT, "Q")	'N:Required, Q:Protected, D:Default
		Call ggoOper.SetReqAttr(frm1.txtStIntPayAmt, "D")
		Call ggoOper.SetReqAttr(frm1.txtStIntPayLocAmt, "D")
		Call ggoOper.SetReqAttr(frm1.txtRcptType, "N")
		Call ggoOper.SetReqAttr(frm1.txtRcptAcctCd, "N")
	Case "DI"	'후급 
		frm1.txtStIntPayAmt.Text = "0"
		frm1.txtStIntPayLocAmt.Text = "0"
		Call ggoOper.SetReqAttr(frm1.txt1StIntDueDT, "N")	'N:Required, Q:Protected, D:Default
		Call ggoOper.SetReqAttr(frm1.txtStIntPayAmt, "Q")
		Call ggoOper.SetReqAttr(frm1.txtStIntPayLocAmt, "Q")
	Case Else
		frm1.txt1StIntDueDT.Text = ""
		frm1.txtStIntPayAmt.Text = "0"
		frm1.txtStIntPayLocAmt.Text = "0"
		Call ggoOper.SetReqAttr(frm1.txt1StIntDueDT, "Q")	'N:Required, Q:Protected, D:Default
		Call ggoOper.SetReqAttr(frm1.txtStIntPayAmt, "Q")
		Call ggoOper.SetReqAttr(frm1.txtStIntPayLocAmt, "Q")
	End Select
	Call txtChargeAmt_change()
	Call txtBPAmt_change()
	frm1.hRdpSprdFg.value = "N"		
End Sub

Sub cboIntPayStnd_OnChange()
	Call cboIntPayStnd_Change()

	frm1.txtIntAcctCd.Value			= ""
	frm1.txtIntAcctNm.Value			= ""
End Sub

'=======================================================================================================
'   Event Name : _Change()
'   Event Desc : 원금상환방법별 Set Protected/Required Fields
'=======================================================================================================
Sub cboPrRdpCond_OnChange()
	 '최초원금상환일, 원금상환일, 상환주기, 원금상환액 
	Select Case frm1.cboPrRdpCond.value
	Case "EQ"		'균등상환 
		Call ggoOper.SetReqAttr(frm1.txt1StPrRdpDt, "N")	'N:Required, Q:Protected, D:Default		
		Call ggoOper.SetReqAttr(frm1.txtPrRdpPerd, "N")	
		Call ggoOper.SetReqAttr(frm1.htxtPrRdpUnitAmt, "D")
		Call ggoOper.SetReqAttr(frm1.htxtPrRdpUnitLocAmt, "D")		
				
	Case "EX",""	'만기상환 
		frm1.txt1StPrRdpDt.Text = ""		
		frm1.txtPrRdpPerd.Text  = ""
		Call ggoOper.SetReqAttr(frm1.txt1StPrRdpDt, "Q")	'N:Required, Q:Protected, D:Default		
		Call ggoOper.SetReqAttr(frm1.txtPrRdpPerd, "Q")	
		Call ggoOper.SetReqAttr(frm1.htxtPrRdpUnitAmt, "Q")
		Call ggoOper.SetReqAttr(frm1.htxtPrRdpUnitLocAmt, "Q")
	Case Else
	End Select
End Sub

'=======================================================================================================
'   Event Name : _Change()
'   Event Desc : 차입처 선택시 clear
'=======================================================================================================

Function Radio1_onChange()									'환율변동여부 
	lgBlnFlgChgValue = True
'	frm1.hRdpSprdFg.value = "N"
End Function

Function Radio2_onChange()
	lgBlnFlgChgValue = True
'	frm1.hRdpSprdFg.value = "N"
End Function

Function Radio5_onChange									'시작일포함여부 
	lgBlnFlgChgValue = True
	frm1.hRdpSprdFg.value = "N"
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

Sub txtRcptAcctCd_OnChange()
	frm1.txtRcptAcctNm.value = ""
	lgBlnFlgChgValue = True	
End Sub 

Sub txtChargeAcctCd_OnChange()
	frm1.txtChargeAcctNm.value = ""
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

'=======================================================================================================
'   Event Desc : 출금유형별 Set Protected/Required Fields
'=======================================================================================================
Sub txtRcptType_OnChange()

	Dim strval
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	
    strval = frm1.txtRcptType.value
            
    IF CommonQueryRs( "REFERENCE" , "B_CONFIGURATION  " , "MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND MINOR_CD =  " & FilterVar(strval , "''", "S") & " AND SEQ_NO = 4 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then         
		   Select Case UCase(lgF0)
				Case "CS" & Chr(11)					
					frm1.txtBankCd.value = ""
					frm1.txtBankNm.value = ""
					frm1.txtBankAcctNo.value = ""
					Call ggoOper.SetReqAttr(frm1.txtBankCd, "Q")
					Call ggoOper.SetReqAttr(frm1.txtBankAcctNo, "Q")									
										
				Case "DP" & Chr(11)			' 예적금 
					Call ggoOper.SetReqAttr(frm1.txtBankCd, "N")
					Call ggoOper.SetReqAttr(frm1.txtBankAcctNo, "N")
				Case "NO" & Chr(11)
					frm1.txtBankCd.value = ""
					frm1.txtBankNm.value = ""
					frm1.txtBankAcctNo.value = ""
					Call ggoOper.SetReqAttr(frm1.txtBankCd, "Q")
					Call ggoOper.SetReqAttr(frm1.txtBankAcctNo, "Q")
			
				Case Else
					frm1.txtBankCd.value = ""
					frm1.txtBankNm.value = ""
					frm1.txtBankAcctNo.value = ""
					Call ggoOper.SetReqAttr(frm1.txtBankCd, "Q")
					Call ggoOper.SetReqAttr(frm1.txtBankAcctNo, "Q")
			End Select
	else
		frm1.txtBankCd.value = ""
		frm1.txtBankNm.value = ""
		frm1.txtBankAcctNo.value = ""		
		Call ggoOper.SetReqAttr(frm1.txtBankCd, "Q")
		Call ggoOper.SetReqAttr(frm1.txtBankAcctNo, "Q")		
	End If
	
	frm1.txtRcptAcctCd.value = ""
	frm1.txtRcptAcctNm.value = ""  
	
End Sub

Sub txtRcptType_Change()
	Dim strval
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	
    strval = frm1.txtRcptType.value
            
    IF CommonQueryRs( "REFERENCE" , "B_CONFIGURATION  " , "MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND MINOR_CD =  " & FilterVar(strval , "''", "S") & " AND SEQ_NO = 4 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then         
		   Select Case UCase(lgF0)
				Case "CS" & Chr(11)					
					frm1.txtBankCd.value = ""
					frm1.txtBankNm.value = ""
					frm1.txtBankAcctNo.value = ""
					Call ggoOper.SetReqAttr(frm1.txtBankCd, "Q")
					Call ggoOper.SetReqAttr(frm1.txtBankAcctNo, "Q")									
										
				Case "DP" & Chr(11)			' 예적금 
					Call ggoOper.SetReqAttr(frm1.txtBankCd, "N")
					Call ggoOper.SetReqAttr(frm1.txtBankAcctNo, "N")
				Case "NO" & Chr(11)
					frm1.txtBankCd.value = ""
					frm1.txtBankNm.value = ""
					frm1.txtBankAcctNo.value = ""
					Call ggoOper.SetReqAttr(frm1.txtBankCd, "Q")
					Call ggoOper.SetReqAttr(frm1.txtBankAcctNo, "Q")
			
				Case Else
					frm1.txtBankCd.value = ""
					frm1.txtBankNm.value = ""
					frm1.txtBankAcctNo.value = ""
					Call ggoOper.SetReqAttr(frm1.txtBankCd, "Q")
					Call ggoOper.SetReqAttr(frm1.txtBankAcctNo, "Q")
			End Select
	else
		frm1.txtBankCd.value = ""
		frm1.txtBankNm.value = ""
		frm1.txtBankAcctNo.value = ""		
		Call ggoOper.SetReqAttr(frm1.txtBankCd, "Q")
		Call ggoOper.SetReqAttr(frm1.txtBankAcctNo, "Q")		
	End If	
End Sub

Sub Type_itemChange()
    lgBlnFlgChgValue = True
End Sub

Sub txtBpRoCd_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtDueDt_Change()
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
	If Trim(frm1.txtLoanRoDt.Text = "") Then	Exit sub
    lgBlnFlgChgValue = True

		'----------------------------------------------------------------------------------------
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtLoanRoDt.Text, Parent.gDateFormat,""), "''", "S") & "))"			

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

Sub txtLoanRoDt_Change()
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii, jj
	Dim arrVal1, arrVal2

	If Trim(frm1.txtDeptCd.value) <> "" and Trim(frm1.txtLoanRoDt.Text <> "") Then
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtLoanRoDt.Text, Parent.gDateFormat,""), "''", "S") & "))"			
	End If
    lgBlnFlgChgValue = True
End Sub

Sub txtPrRdpPerd_Change()
	lgBlnFlgChgValue = True
End Sub 

Sub txt1StIntDueDt_Change()
    lgBlnFlgChgValue = True   
End Sub

Sub txt1StPrRdpDt_Change() 
    lgBlnFlgChgValue = True
'    Call txt1StPrRdpDt_OnChange()
End Sub

Sub txtXchRate_Change()
	frm1.txtLoanRoLocAmt.Text="0"
	lgBlnFlgChgValue = True
End Sub

Sub txtLoanAmt_Change()
	frm1.txtLoanRoLocAmt.Text="0"
	lgBlnFlgChgValue = True
End Sub

Sub txtStIntPayAmt_Change()
	lgBlnFlgChgValue = True

	If UNICDbl(frm1.txtStIntPayAmt.Text) > 0 Then
		Call ggoOper.SetReqAttr(frm1.txtRcptType, "N")
		Call ggoOper.SetReqAttr(frm1.txtRcptAcctCd, "N")
	Else
		If UNICDbl(frm1.txtChargeAmt.Text) > 0 Or UNICDbl(frm1.txtBpAmt.Text) > 0 Then
			Call ggoOper.SetReqAttr(frm1.txtRcptType, "N")
			Call ggoOper.SetReqAttr(frm1.txtRcptAcctCd, "N")
		Else
			frm1.txtRcptType.value = ""
			frm1.txtRcptTypeNm.value = ""
			frm1.txtRcptAcctCd.value = ""
			frm1.txtRcptAcctNm.value = ""
			frm1.txtBankCd.value = ""
			frm1.txtBankNm.value = ""
			frm1.txtBankAcctNo.value = ""
			Call ggoOper.SetReqAttr(frm1.txtRcptType, "Q")
			Call ggoOper.SetReqAttr(frm1.txtRcptAcctCd, "Q")
			Call ggoOper.SetReqAttr(frm1.txtBankCd, "Q")
			Call ggoOper.SetReqAttr(frm1.txtBankAcctNo, "Q")
		End If
	End If
End Sub

Sub txtStIntPayLocAmt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtPrRdpUnitAmt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtPrRdpUnitLocAmt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtIntRate_Change()
	lgBlnFlgChgValue = True
	frm1.hRdpSprdFg.Value = "N"
End Sub

Sub txtIntPayPerd_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtChargeAmt_Change()
	lgBlnFlgChgValue = True

	If UNICDbl(frm1.txtChargeAmt.Text) > 0 Then
		Call ggoOper.SetReqAttr(frm1.txtChargeAcctCd, "N")
		Call ggoOper.SetReqAttr(frm1.txtRcptType, "N")
		Call ggoOper.SetReqAttr(frm1.txtRcptAcctCd, "N")
	Else
		frm1.txtChargeAcctCd.value = ""
		frm1.txtChargeAcctNm.value = ""
		Call ggoOper.SetReqAttr(frm1.txtChargeAcctCd, "Q")
		If UNICDbl(frm1.txtStIntPayAmt.Text) > 0 Or UNICDbl(frm1.txtBpAmt.Text) > 0 Then
			Call ggoOper.SetReqAttr(frm1.txtRcptType, "N")
			Call ggoOper.SetReqAttr(frm1.txtRcptAcctCd, "N")
		Else
			frm1.txtRcptType.value = ""
			frm1.txtRcptTypeNm.value = ""
			frm1.txtRcptAcctCd.value = ""
			frm1.txtRcptAcctNm.value = ""
			frm1.txtBankCd.value = ""
			frm1.txtBankNm.value = ""
			frm1.txtBankAcctNo.value = ""
			Call ggoOper.SetReqAttr(frm1.txtRcptType, "Q")
			Call ggoOper.SetReqAttr(frm1.txtRcptAcctCd, "Q")
			Call ggoOper.SetReqAttr(frm1.txtBankCd, "Q")
			Call ggoOper.SetReqAttr(frm1.txtBankAcctNo, "Q")
		End If
	End If
End Sub

Sub txtChargeLocAmt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtBpAmt_Change()
	lgBlnFlgChgValue = True
	
	If UNICDbl(frm1.txtBPAmt.Text) > 0 Then
		Call ggoOper.SetReqAttr(frm1.txtBPAcctCd, "N")
		Call ggoOper.SetReqAttr(frm1.txtRcptType, "N")
		Call ggoOper.SetReqAttr(frm1.txtRcptAcctCd, "N")
	Else
		frm1.txtBPAcctCd.value = ""
		frm1.txtBPAcctNm.value = ""
		Call ggoOper.SetReqAttr(frm1.txtBPAcctCd, "Q")
		If UNICDbl(frm1.txtStIntPayAmt.Text) > 0 Or UNICDbl(frm1.txtChargeAmt.Text) > 0 Then
			Call ggoOper.SetReqAttr(frm1.txtRcptType, "N")
			Call ggoOper.SetReqAttr(frm1.txtRcptAcctCd, "N")
		Else
			frm1.txtRcptType.value = ""
			frm1.txtRcptTypeNm.value = ""
			frm1.txtRcptAcctCd.value = ""
			frm1.txtRcptAcctNm.value = ""
			frm1.txtBankCd.value = ""
			frm1.txtBankNm.value = ""
			frm1.txtBankAcctNo.value = ""
			Call ggoOper.SetReqAttr(frm1.txtRcptType, "Q")
			Call ggoOper.SetReqAttr(frm1.txtRcptAcctCd, "Q")
			Call ggoOper.SetReqAttr(frm1.txtBankCd, "Q")
			Call ggoOper.SetReqAttr(frm1.txtBankAcctNo, "Q")
		End If
	End If
End Sub

Sub txtBPLocAmt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtBPAcctCd_onChange()
	lgBlnFlgChgValue = True
End Sub
Sub txtIntRdpAmt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtIntRdpLocAmt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtDocCur_OnChange()
'    lgBlnFlgChgValue = True
    IF CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
		Call CurFormatNumericOCX()
	END IF	    
	
'	Call FncCalcRate()

'	If frm1.txtDocCur.value <> Parent.gCurrency Then
'		frm1.txtXchRate.Text = "0"
'	ElseIf frm1.txtDocCur.value = Parent.gCurrency Then
'		frm1.txtXchRate.Text = "1"
'	End If  
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029																'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    
    Call InitVariables																'⊙: Initializes local global variables
    Call InitComboBox
	
	'ggoOper.FormatNumber(Obj, Max, Min, Separator(True), DecimalPlace(0), DecimalPoint(.), Separator(,))
	Call ggoOper.FormatNumber(frm1.txtPrRdpPerd, "999", "1", False)					'상환주기 
	Call ggoOper.FormatNumber(frm1.txtPrRdpPerd, "99999", "1", False)				'원금상환주기	
	Call ggoOper.FormatNumber(frm1.txtIntPayPerd, "99999", "1", False)				'이자지급주기 
	
	Call FncNew()
	Call CookiePage("FORM_LOAD")

	Call SetDefaultVal

	' 권한관리 추가 
	Dim xmlDoc
	
	Call GetDataAuthXML(parent.gUsrID, gStrRequestMenuID, xmlDoc) 
	
	' 사업장 
	lgAuthBizAreaCd	= xmlDoc.selectSingleNode("/root/data/data_biz_area_cd").Text
	lgAuthBizAreaNm	= xmlDoc.selectSingleNode("/root/data/data_biz_area_nm").Text

	' 내부부서 
	lgInternalCd	= xmlDoc.selectSingleNode("/root/data/data_internal_cd").Text
	lgDeptCd		= xmlDoc.selectSingleNode("/root/data/data_dept_cd").Text
	lgDeptNm		= xmlDoc.selectSingleNode("/root/data/data_dept_nm").Text
	
	' 내부부서(하위포함)
	lgSubInternalCd	= xmlDoc.selectSingleNode("/root/data/data_sub_internal_cd").Text
	lgSubDeptCd		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_cd").Text
	lgSubDeptNm		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_nm").Text
	
	' 개인 
	lgAuthUsrID		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_id").Text
	lgAuthUsrNm		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_nm").Text
	
	Set xmlDoc = Nothing
	
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
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")
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
    Call DbQuery																'☜: Query db data
           
    FncQuery = True																'⊙: Processing is OK        
    
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
       IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"X","X")           '⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "A")                                      '⊙: Clear Condition Field
    
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

	Call ggoOper.FormatNumber(frm1.txtPrRdpPerd, "999", "1", False)					'상환주기 
	Call ggoOper.FormatNumber(frm1.txtPrRdpPerd, "99999", "1", False)				'원금상환주기	
	Call ggoOper.FormatNumber(frm1.txtIntPayPerd, "99999", "1", False)				'이자지급주기 
	
    Call ggoOper.LockField(Document, "N")                                       '⊙: Lock  Suitable  Field
	Call cboIntPayStnd_OnChange()
    frm1.cboRdpClsFg.value = "N"    
    Call txtRcptType_OnChange()
    Call cboPrRdpCond_OnChange()
    
    Call SetDefaultVal
    Call InitVariables						'⊙: Initializes local global variables

    Call FncSetToolBar("New")
		
	frm1.txtLoanRoNo.focus
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
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                
        Exit Function
    End If
    
	IF Trim(frm1.txtLoanRoNo.value) <> Trim(lgLoanRoNo) Then
		Call DisplayMsgBox("900002","X","X","X")                                
		Exit Function
	End If
    
  '-----------------------
    'Delete function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900003",Parent.VB_YES_NO,"X","X")
    IF IntRetCD = vbNO Then
		Exit Function
	End IF	
    
    Call DbDelete															'☜: Delete db data
    
    FncDelete = True                                                        '⊙: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to save Button of Main ToolBar
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
		IF Trim(frm1.txtLoanRoNo.value) <> Trim(lgLoanRoNo) Then
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
    If Not chkField(Document, "2") Then									  '⊙: Check contents area
       Exit Function
    End If
    
    If frm1.txtDocCur.value =  parent.gCurrency Then
		frm1.txtXchRate.text = 1
    End If 
    '-----------------------
    'Save function call area
    '-----------------------
    CAll DBSave				                                                '☜: Save db data
    
    FncSave = True                                                          '⊙: Processing is OK

End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
    On Error Resume Next                                                    '☜: Protect system from crashing
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
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      'Check if there is retrived data
		Call DisplayMsgBox("900002","X","X","X")
		Exit Function
    End If
    
   '-----------------------
    'Check previous data area
    '------------------------ 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")
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
           
    FncPrev = True																'⊙: Processing is OK        
    
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
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      'Check if there is retrived data
		Call DisplayMsgBox("900002","X","X","X")
		Exit Function
    End If
    
   '-----------------------
    'Check previous data area
    '------------------------ 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")
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
           
    FncNext = True																'⊙: Processing is OK        
    
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_SINGLE)												'☜: 화면 유형 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE, False)                                         '☜:화면 유형, Tab 유무 
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	
	FncExit = False
	
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")			'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    FncExit = True
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
    
    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtLoanRoNo=" & Trim(lgLoanRoNo)		'☜: 삭제 조건 데이타 
    strVal = strVal & "&txtLoanNo="   & Trim(frm1.txtLoanNo.value)			'☜: 삭제 조건 데이타 

	' 권한관리 추가 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 
   
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	
    DbDelete = True                                                         '⊙: Processing is NG

End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일 때 수행 
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

		ggoOper.FormatFieldByObjectOfCur .txtLoanAmt,	  .txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtIntPayAmt,	  .txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec		
		ggoOper.FormatFieldByObjectOfCur .txtLoanBalAmt,  .txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtLoanRoAmt,	  .txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec		
		ggoOper.FormatFieldByObjectOfCur .txtRdpAmt,	  .txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtStIntPayAmt, .txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtChargeAmt,	  .txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtBPAmt,		  .txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtTotPrRdpRoAmt,.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtIntPayRoAmt,  .txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtLoanBalRoAmt, .txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec		
	
	End With

End Sub
'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    Dim strVal
        
    Call LayerShowHide(1)
    
    Err.Clear                                                               '☜: Protect system from crashing
    
    DbQuery = False                                                         '⊙: Processing is NG
    
    strVal = BIZ_PGM_ID2 & "?txtMode	=" & Parent.UID_M0001						'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtCommand		=" & Trim(frm1.hCommand.value)
    strVal = strVal & "&txtLoanNo		=" & Trim(frm1.txtLoanRoNo.value)  	'☆: 조회 조건 데이타 
	strVal = strVal & "&txtLoanPlcType	=" & "BP"						 	'☆: 조회 조건 데이타 
    strVal = strVal & "&txtLoanBasicFg	=" & "LR"							'☆: 조회 조건 데이타 

	' 권한관리 추가 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	
    DbQuery = True   
                                                           '⊙: Processing is NG

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 
	
	Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
	Call cboIntPayStnd_Change()
    Call cboPrRdpCond_OnChange()    
'    Call txtRcptType_OnChange
    Call txtRcptType_Change
    Call CurFormatNumericOCX()   
    
    Call InitVariables

	lgLoanRoNo = frm1.txtLoanRoNo.value
        
    lgIntFlgMode = Parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
	lgtempstrfg  = frm1.txtStrFg.Value
    
    Call FncSetToolBar("Query")
    
    frm1.txtLoanRoNo.focus
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
		.txtMode.value = Parent.UID_M0002											'☜: 비지니스 처리 ASP 의 상태 
		.txtFlgMode.value = lgIntFlgMode
		.txtUpdtUserId.value = Parent.gUsrID
		.txtInsrtUserId.value = Parent.gUsrID
		.txtstrFg.value = lgtempStrFg
		.htxtLoanPlcType.value = "BP"
		.txtloanbasicFg.value = "LR"
		.txtLoanTerm.value = strDiffDate

		'권한관리추가 start
		.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
		.txthInternalCd.value =  lgInternalCd
		.txthSubInternalCd.value = lgSubInternalCd
		.txthAuthUsrID.value = lgAuthUsrID		
		'권한관리추가 end
				
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
		Case Parent.OPMD_CMODE
			frm1.txtLoanRoNo.value = pLoanNo
						
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
		Call SetToolbar("1110000000001111")
	Case "QUERY"
		Call SetToolbar("1111100011011111")
	Case "REF"
		Call SetToolbar("1110100000001111")
	End Select
End Function

'==========================================================
'상환전개버튼 클릭 
'==========================================================
Function FnButtonExec()
	Dim intRetCD
    Dim strVal

    If lgIntFlgMode <> Parent.OPMD_UMODE Then				'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")	'조회를 먼저 하십시오.
        Exit Function
    End If
    
	IF Trim(frm1.txtLoanRoNo.value) <> Trim(lgLoanRoNo) Then
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
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"X","X")	'데이터가 변경되었습니다. 계속하시겠습니까?
		If IntRetCD = vbNo Then
			Exit Function
		End If
	Else
		IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X")	'작업을 수행하시겠습니까?
		If IntRetCD = vbNO Then
			Exit Function
		End If
	End If
    
    Call LayerShowHide(1)
    
    Err.Clear                                                               '☜: Protect system from crashing
    
    strVal = BIZ_PGM_ID & "?txtMode=" & "PAFG400"							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtLoanNo=" & Trim(lgLoanRoNo)  	    '☆: 조회 조건 데이타 
    strVal = strVal & "&txtDateFr=" & Trim(frm1.txtLoanRoDt.text) 
    strVal = strVal & "&txtDateTo=" & Trim(frm1.txtLoanRoDt.text)
    
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
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
// -->
</SCRIPT>

</HEAD>
<BODY TABINDEX="-1" SCROLL="auto">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH=* ALIGN=RIGHT><A HREF="VBSCRIPT:OpenPopupLoan()">차입금참조</A> |
											<A HREF="VBSCRIPT:OpenPopupTempGL()">결의전표</A> |
											<A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A>
					</TD>
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
									<TD CLASS="TD5" NOWRAP>만기연장번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtLoanRoNo" SIZE="18" MAXLENGTH="18" tag="12XXXU" ALT="만기연장번호" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoanNo" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopupRoLoan()"></TD>
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
					<TD HEIGHT=100% valign=top>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>차입금번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtLoanNo" SIZE="18" MAXLENGTH="18" tag="24X" ALT="차입금번호" ></TD>
									<TD CLASS="TD5" NOWRAP>차입거래처</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBpRoCd" SIZE="10" MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="24X" ALT="차입거래처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankLoanCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopUp(frm1.txtBpRoCd.value, 5)">
									                       <INPUT TYPE=TEXT NAME="txtBpRoNm" ALT="차입거래처명" SIZE=20 tag="24X"></TD>									
								</TR>							
								<TR>								     
									<TD CLASS="TD5" NOWRAP>차입금액|자국</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpLoanAmt name=txtLoanAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="차입금액" tag="24X2Z"></OBJECT>');</SCRIPT>&nbsp;
														   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpLoanLocAmt name=txtLoanLocAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="차입금액(자국)" tag="24X2Z"></OBJECT>');</SCRIPT></TD>
									<TD CLASS="TD5" NOWRAP>이자지급액|자국</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpIntPayAmt name=txtIntPayAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="이자지급액" tag="24X2Z"></OBJECT>');</SCRIPT>&nbsp;
														   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpIntPayLocAmt name=txtIntPayLocAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="이자지급액(자국)" tag="24X2Z"></OBJECT>');</SCRIPT></TD>
								</TR>								
								<TR>							
									<TD CLASS="TD5" NOWRAP>차입잔액|자국</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpIntRdpAmt name=txtLoanBalAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="상환잔액" tag="24X2Z"></OBJECT>');</SCRIPT>&nbsp;
														   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpIntRdpLocAmt name=txtLoanBalLocAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="상환잔액(자국)" tag="24X2Z"></OBJECT>');</SCRIPT></TD>
									<TD CLASS="TD5" NOWRAP>원금상환액|자국</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpPrRdpAmt name=txtRdpAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="원금상환액" tag="24X2Z"></OBJECT>');</SCRIPT>&nbsp;
														   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpPrRdpLocAmt name=txtRdpLocAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="원금상환액(자국)" tag="24X2Z"></OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>만기연장내역</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtLoanRoNm" SIZE="38" MAXLENGTH="40" tag="22X" ALT="만기연장내역"></TD>
									<TD CLASS="TD5" NOWRAP>부서</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDeptCd" ALT="부서코드" Size= "10" MAXLENGTH="10"  tag="22X" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopUp(frm1.txtDeptCd.value, 2)">
														   <INPUT NAME="txtDeptNm" ALT="부서명" SIZE = "20" STYLE="TEXT-ALIGN: left" tag="24X"></TD>
								</TR>								
								<TR>
									<TD CLASS="TD5" NOWRAP>만기연장일</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpLoanRoDt name=txtLoanRoDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="22X1" ALT="만기연장일"></OBJECT>');</SCRIPT></TD>
									<TD CLASS="TD5" NOWRAP>상환만기일</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDueDt name=txtDueDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="22X1" ALT="상환만기일"></OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>장단기구분</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboLoanFg" ALT="장단기구분" STYLE="WIDTH: 135px" tag="22X" OnClick ="vbscript:Type_itemChange()"><OPTION VALUE=""></OPTION></SELECT></TD>
									<TD CLASS="TD5" NOWRAP>만기연장계정</TD>												
								    <TD CLASS="TD6" NOWRAP><INPUT NAME="txtLoanAcctCd" ALT="만기연장계정" SIZE="10" MAXLENGTH="20"  tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoanAcct" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtLoanAcctCd.value, 8)">
														   <INPUT NAME="txtLoanAcctNm" ALT="만기연장계정명" SIZE="20" tag="24X"></TD>									
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>차입용도</TD>												
								    <TD CLASS="TD6" NOWRAP><INPUT NAME="txtLoanType" ALT="차입용도" SIZE="10" MAXLENGTH="2"  tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoanType" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtLoanType.value, 6)">
														   <INPUT NAME="txtLoanTypeNm" ALT="차입용도명" STYLE="TEXT-ALIGN: Left" tag="24X"></TD>
									<TD CLASS="TD5" NOWRAP>거래통화|환율</TD>
								    <TD CLASS="TD6" NOWRAP><INPUT NAME="txtDocCur" ALT="거래통화" SIZE = "10" MAXLENGTH="3"  tag="24XXXU">&nbsp;
															<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpXRate name=txtXchRate align="top" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="환율" tag="24X5Z"></OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>원금상환방법</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboPrRdpCond" ALT="원금상환방법" STYLE="WIDTH: 135px" tag="22X" OnClick ="vbscript:Type_itemChange()"><OPTION VALUE=""></OPTION></SELECT></TD>
									<TD CLASS="TD5" NOWRAP>만기연장액|자국</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpLoanRoAmt name=txtLoanRoAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="차입금액" tag="24X2Z"></OBJECT>');</SCRIPT>&nbsp;
														   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpLoanRoLocAmt name=txtLoanRoLocAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="차입금액(자국)" tag="24X2Z"></OBJECT>');</SCRIPT></TD>																							 									
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>최초원금상환일</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fp1StPrRdpDt name=txt1StPrRdpDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="22X1" ALT="최초원금상환예정일"></OBJECT>');</SCRIPT></TD>
									<TD CLASS="TD5" NOWRAP>원금상환주기</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpPrRdpPerd Name=txtPrRdpPerd CLASS=FPDS65 title=FPDOUBLESINGLE ALT="원금상환주기" tag="22X"></OBJECT>');</SCRIPT>개월</TD>
								</TR>
								<TR>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>이자율변동성</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio_IntVotl ID=Rb_IntVotl1 Checked tag = 2 value="X" onclick=radio1_onchange()><LABEL FOR=Rb_IntVotl1>확정</LABEL>&nbsp;
														   <INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio_IntVotl ID=Rb_IntVotl2 tag = 2 value="F" onclick=radio2_onchange()><LABEL FOR=Rb_IntVotl2>변동</LABEL>&nbsp;</TD>
									<TD CLASS="TD5" NOWRAP>이자율</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=OBJECT5 Name=txtIntRate CLASS=FPDS90 title=FPDOUBLESINGLE ALT="이자율" tag="22X5Z"></OBJECT>');</SCRIPT>&nbsp;%&nbsp;/&nbsp;년</TD>
								</TR>								
								<TR>
									<TD CLASS="TD5" NOWRAP>이자지급형태</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboIntPayStnd" ALT="이자지급형태" STYLE="WIDTH: 135px" tag="22X" ><OPTION VALUE=""></OPTION></SELECT></TD>
									<TD CLASS="TD5" NOWRAP>이자계정</TD>
								    <TD CLASS="TD6" NOWRAP><INPUT NAME="txtIntAcctCd" ALT="이자계정" SIZE="10" MAXLENGTH="20"  tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnIntAcct" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtIntAcctCd.value, 10)">
														   <INPUT NAME="txtIntAcctNm" ALT="이자계정명" SIZE="20" tag="24X"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>최초이자지급일</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fp1StIntDueDT name=txt1StIntDueDT CLASS=FPDTYYYYMMDD title=FPDATETIME tag="22X1" ALT="최초이자지급일"></OBJECT>');</SCRIPT></TD>									
									<TD CLASS="TD5" NOWRAP>이자지급주기</TD>
								    <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpIntPayPerd Name=txtIntPayPerd align="top" CLASS=FPDS40 title=FPDOUBLESINGLE ALT="이자지급주기" tag="22X"></OBJECT>');</SCRIPT>&nbsp;/&nbsp;<SELECT NAME="cboIntBaseMthd" ALT="이자계산방법" STYLE="WIDTH: 135px" tag="22X" OnClick ="vbscript:Type_itemChange()"><OPTION VALUE=""></OPTION></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>시작일포함여부</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio_IntStart ID=Rb_IntStart1 Checked tag = 2 value="Y" onclick=radio5_onchange()><LABEL FOR=Rb_IntStart1>포함</LABEL>&nbsp;
														   <INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio_IntStart ID=Rb_IntStart2 tag = 2 value="N" onclick=radio6_onchange()><LABEL FOR=Rb_IntStart2>미포함</LABEL>&nbsp;</TD>														   
									<TD CLASS="TD5" NOWRAP>
									<TD CLASS="TD6" NOWRAP>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>만기일포함여부</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio_IntEnd ID=Rb_IntEnd1 tag = 2 value="Y" onclick=radio7_onchange()><LABEL FOR=Rb_IntEnd1>포함</LABEL>&nbsp;
														   <INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio_IntEnd ID=Rb_IntEnd2 Checked tag = 2 value="N" onclick=radio8_onchange()><LABEL FOR=Rb_IntEnd2>미포함</LABEL>&nbsp;</TD>														   
									<TD CLASS="TD5" NOWRAP>선급이자액|자국</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpStIntPayAmt name=txtStIntPayAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="선급이자액" tag="24X2Z"></OBJECT>');</SCRIPT>&nbsp;
														   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpStIntPayLocAmt name=txtStIntPayLocAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="선급이자액" tag="24X2Z"></OBJECT>');</SCRIPT></TD>									
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>부대비용|자국</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpChargeAmt name=txtChargeAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="부대비용" tag="21X2Z"></OBJECT>');</SCRIPT>&nbsp;
														   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpChargeLocAmt name=txtChargeLocAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="부대비용" tag="21X2Z"></OBJECT>');</SCRIPT></TD>									
									<TD CLASS="TD5" NOWRAP>부대비용계정</TD>
								    <TD CLASS="TD6" NOWRAP><INPUT NAME="txtChargeAcctCd" ALT="부대비용계정" SIZE="10" MAXLENGTH="20"  tag="24X"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnChargeAcct" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtChargeAcctCd.value, 9)">
														   <INPUT NAME="txtChargeAcctNm" ALT="부대비용계정명" SIZE="20" tag="24X"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>수수료|자국</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpBPAmt name=txtBPAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="수수료" tag="21X2Z"></OBJECT>');</SCRIPT>&nbsp;
														   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpBPLocAmt name=txtBPLocAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="수수료" tag="21X2Z"></OBJECT>');</SCRIPT></TD>									
									<TD CLASS="TD5" NOWRAP>수수료계정</TD>
								    <TD CLASS="TD6" NOWRAP><INPUT NAME="txtBPAcctCd" ALT="수수료계정" SIZE="10" MAXLENGTH="20"  tag="24X"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBPAcct" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBPAcctCd.value, 12)">
														   <INPUT NAME="txtBPAcctNm" ALT="수수료계정명" SIZE="20" tag="24X"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>출금유형</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtRcptType" ALT="출금유형" SIZE="10" MAXLENGTH="2"  tag="24XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRcptType" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtRcptType.value, 7)">
														   <INPUT NAME="txtRcptTypeNm" ALT="출금유형명" STYLE="TEXT-ALIGN: Left" tag="24X"></TD>									          
									<TD CLASS="TD5" NOWRAP>출금계정</TD>
								    <TD CLASS="TD6" NOWRAP><INPUT NAME="txtRcptAcctCd" ALT="출금계정" SIZE="10" MAXLENGTH="20"  tag="24X"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRcptAcct" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtRcptAcctCd.value, 11)">
														   <INPUT NAME="txtRcptAcctNm" ALT="출금계정명" SIZE="20" tag="24X"></TD>
								</TR>
								<TR>
 								    <TD CLASS="TD5" NOWRAP>출금계좌번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtBankAcctNo" ALT="출금계좌번호" SIZE="18" MAXLENGTH="30"  tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankAcct" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopUp(frm1.txtBankAcctNo.value, 4)"></TD> 								    
									<TD CLASS="TD5" NOWRAP>출금은행</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtBankCd" ALT="출금은행" SIZE="10" MAXLENGTH="10"  tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBankCd.value, 3)">
														   <INPUT NAME="txtBankNm" ALT="출금은행명" STYLE="TEXT-ALIGN: Left" tag="24X"></TD>							
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
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpPrRdpRoAmt name=txtTotPrRdpRoAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="원금상환총액" tag="24X2Z"></OBJECT>');</SCRIPT>&nbsp;
														   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpPrRdpRoLocAmt name=txtTotPrRdpRoLocAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="원금상환총액" tag="24X2Z"></OBJECT>');</SCRIPT></TD>
									<TD CLASS="TD5" NOWRAP>이자지급총액|자국</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpIntPayRoAmt name=txtIntPayRoAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="이자지급총액" tag="24X2Z"></OBJECT>');</SCRIPT>&nbsp;
														   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpIntPayRoLocAmt name=txtIntPayRoLocAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="이자지급총액" tag="24X2Z"></OBJECT>');</SCRIPT></TD>									
								</TR>
								<TR>									
									<TD CLASS="TD5" NOWRAP>차입잔액|자국</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpIntRdpRoAmt name=txtLoanBalRoAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="상환잔액" tag="24X2Z"></OBJECT>');</SCRIPT>&nbsp;
														   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpIntRdpRoLocAmt name=txtLoanBalRoLocAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="상환잔액" tag="24X2Z"></OBJECT>');</SCRIPT></TD>
									<TD CLASS="TD5" NOWRAP>상환완료여부</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboRdpClsFg" ALT="상환완료여부" STYLE="WIDTH: 135px" tag="24X" OnClick ="vbscript:Type_itemChange()"><OPTION VALUE=""> </OPTION></SELECT></TD>									
								</TR>
								<TR>
								</TR>								
								<INPUT TYPE=hidden CLASS="Radio" NAME=Radio_Cur ID=hRb_Cur1  onclick=radio9_onchange()><LABEL FOR=Rb_Cur1>
								<INPUT TYPE=hidden CLASS="Radio" NAME=Radio_Cur ID=hRb_Cur2 onclick=radio10_onchange()><LABEL FOR=Rb_Cur2>																							
								<INPUT TYPE=hidden NAME="htxtTempGlNo" tag="24">
								<INPUT TYPE=hidden NAME="htxtGlNo" tag="24">
								<INPUT TYPE=hidden NAME="htxtLoanPlcType" tag="24">								
								<INPUT TYPE=hidden NAME="htxtPrRdpUnitAmt" tag="24">
								<INPUT TYPE=hidden NAME="htxtPrRdpUnitLocAmt" tag="24">								
								<INPUT TYPE=hidden NAME="hRdpSprdFg" tag="24">
								<INPUT TYPE=hidden NAME="hClsRoFg" tag="24">								    																								
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
<INPUT TYPE=hidden NAME="txtloanbasicFg" tag="24">
<INPUT TYPE=hidden NAME="txtLoanTerm" tag="24">
<INPUT TYPE=hidden NAME="htxtOrgLoanRcptType" tag="24">
<INPUT TYPE=hidden NAME="htxtOrgLoanRcptAcctCd" tag="24">
<INPUT TYPE=hidden NAME="htxtOrgLoanBankAcctNo" tag="24">
<INPUT TYPE=hidden NAME="htxtOrgLoanBankCd" tag="24">
<INPUT TYPE=hidden NAME="txthAuthBizAreaCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthSubInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthUsrID"		tag="24" Tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

