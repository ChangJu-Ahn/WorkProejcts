
<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Deposit
'*  3. Program ID           : f3101ma1
'*  4. Program Name         : 예적금등록 
'*  5. Program Desc         : Register of Deposit Master
'*  6. Comproxy List        : FD0011, FD0019
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2001.03.05
'*  9. Modifier (First)     : Kim, Jong Hwan
'* 10. Modifier (Last)      : Kim, Hee Jung
'* 11. Comment              : 2001.05.31 Song,MunGil 조직변경반영/자국필드추가내용반영 
'======================================================================================================= -->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'##############################################################################################################
'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->					<!-- '⊙: 화면처리ASP에서 서버작업이 필요한 경우  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->				
<!--'==========================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!-- '☆: 해당 위치에 따라 달라짐, 상대 경로  -->

<!--'==========================================  1.1.2 공통 Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                              '☜: indicates that All variables must be declared in advance 

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
<!-- #Include file="../../inc/lgvariables.inc" -->	
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================

' 선언전 실행할 내용 Coding


Const BIZ_PGM_ID = "f3101mb1.asp"											 '☆: 비지니스 로직 ASP명 
Const BIZ_PGM_ID2 = "f3101mb2.asp"	

Const JUMP_PGM_ID_BANK_REP = "b1310ma1"										 '☆: Jump Page to 은행정보등록 

 '==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 
Dim lgBlnFlgConChg				'☜: Condition 변경 Flag

Dim lgNextNo						'☜: 화면이 Single/SingleMulti 인경우만 해당 
Dim lgPrevNo						' ""

'==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= 
 '----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 

 '++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 
Dim lgCurName()					'☆ : 개별 화면당 필요한 로칼 전역 변수 
Dim cboOldVal          
Dim IsOpenPop          

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
    lgIntFlgMode = parent.OPMD_CMODE   '⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False    '⊙: Indicates that no value changed
    lgIntGrpCount = 0           '⊙: Initializes Group View Size
    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False			'☆: 사용자 변수 초기화 

End Sub

'=============================== 2.1.2 LoadInfTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE" , "MA") %>
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

    if frm1.cboTransSts.length > 0 then
       frm1.cboTransSts.selectedindex = 0
    end if
	frm1.hTemp.value = ""
	frm1.txtStartDt.text = UniConvDateAToB("<%=GetSvrDate%>" ,Parent.gServerDateFormat,Parent.gDateFormat) 						
    frm1.hOrgChangeId.value = Parent.gChangeOrgId	
	frm1.txtDocCur.value	= Parent.gCurrency
	frm1.txtXchRate.text	= 1

	frm1.hTemp.value = ""

End Sub

 '==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
		
	Dim arrData
	
	'예적금구분 
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F3011", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboDpstFg ,lgF0  ,lgF1  ,Chr(11))
	
	'예적금유형 
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F3012", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboDpstType ,lgF0  ,lgF1  ,Chr(11))
	
	'거래상태 
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F3014", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboTransSts ,lgF0  ,lgF1  ,Chr(11))
	
	'계좌유형 
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F3013", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboBankAcctFg ,lgF0  ,lgF1  ,Chr(11))
End Sub


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
 '++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++
 '------------------------------------------  OpenRefBankAcctNo()  ----------------------------------------
'	Name : OpenRefBankAcctNo()
'	Description : 은행계좌참조 
'--------------------------------------------------------------------------------------------------------- 
'Function OpenRefBankAcctNo()
'	Dim arrRet
'	Dim arrParam(6), arrField(6), arrHeader(6)
	
'	If IsOpenPop = True Then Exit Function
	
'	arrParam(0) = "은행계좌참조"										' 팝업 명칭 
'	arrParam(1) = "B_BANK A, B_BANK_ACCT B, B_MINOR C, B_MINOR D "			' TABLE 명칭 
'	arrParam(2) = ""														' Code Condition
'	arrParam(3) = ""														' Name Cindition
'	arrParam(4) = "A.BANK_CD = B.BANK_CD AND (B.BP_CD IS NULL OR B.BP_CD = ' ') "
'	arrParam(4) = arrParam(4) & "AND C.MAJOR_CD = 'F3011' AND C.MINOR_CD = B.BANK_ACCT_TYPE "	
'	arrParam(4) = arrParam(4) & "AND D.MAJOR_CD = 'F3012' AND D.MINOR_CD = B.DPST_TYPE "		' Where Condition	
'	arrParam(5) = "은행코드"											' 조건필드의 라벨 명칭 

'	arrField(0) = "A.BANK_CD"								' Field명(0)
'	arrField(1) = "A.BANK_NM"								' Field명(1)
'	arrField(2) = "B.BANK_ACCT_NO"							' Field명(2)
'	arrField(3) = "C.MINOR_NM"								' Field명(3)
'   arrField(4) = "D.MINOR_NM"								' Field명(4)
'   arrField(5) = "HH" & Parent.gColSep & "C.MINOR_CD"		' Field명(5) - Hidden
'	arrField(6) = "HH" & Parent.gColSep & "D.MINOR_CD"		' Field명(6) - Hidden
    
'	arrHeader(0) = "은행코드"							' Header명(0)
'	arrHeader(1) = "은행명"								' Header명(1)
'	arrHeader(2) = "계좌번호"							' Header명(2)
'	arrHeader(3) = "예적금구분"							' Header명(3)
'	arrHeader(4) = "예적금유형"							' Header명(4)	
	
'	IsOpenPop = True
	
'	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
'			"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

'	IsOpenPop = False
	
'	If arrRet(0) = "" Then
'		Exit Function
'	Else
'		frm1.txtBankCd.value		= arrRet(0)
'		frm1.txtBankNM.value		= arrRet(1)
'		frm1.txtBankAcctNo.value	= arrRet(2)
'		frm1.cboDpstFg.Value		= arrRet(5)
'		frm1.cboDpstType.Value		= arrRet(6)
				
'	End If
	
'	Call cboDpstFg_Change()
	
'	frm1.txtBankAcctNo.focus
	
'End Function

Function OpenRefBankAcctNo(ByVal iOpt1, Byval iOpt2)
	Dim arrRet
	Dim arrParam(11)	                           '권한관리 추가 (3 -> 4)
	Dim IntRetCD	
	Dim iCalledAspName
	
	iCalledAspName = AskPRAspName("f3101ra1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "f3101ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
'	arrParam(4)	= lgAuthorityFlag              '권한관리 추가	

   arrParam(5) = iOpt2
   arrParam(6) = iOpt1

	' 권한관리 추가 
	arrParam(8) = lgAuthBizAreaCd
	arrParam(9) = lgInternalCd
	arrParam(10) = lgSubInternalCd
	arrParam(11) = lgAuthUsrID

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=660px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
 
	If arrRet(0) = ""  Then
		frm1.txtBankCd.focus			
		Exit Function
	Else
		frm1.txtBankCd.value		= arrRet(0)
		frm1.txtBankNM.value		= arrRet(1)
		frm1.txtBankAcctNo.value	= arrRet(2)
		frm1.cboDpstFg.Value		= arrRet(3)
		frm1.cboDpstType.Value		= arrRet(4)
	End If
	
	Call cboDpstFg_Change()

End Function

 '------------------------------------------  OpenPopUp()  -------------------------------------------------
'	Name : OpenPopUp()
'	Description : PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	Select Case iWhere

		Case 5
			arrParam(0) = "거래통화 팝업"			' 팝업 명칭 
			arrParam(1) = "B_CURRENCY"		 			' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "거래통화"					' 조건필드의 라벨 명칭 

			arrField(0) = "CURRENCY"					' Field명(0)
			arrField(1) = "CURRENCY_DESC"				' Field명(0)
    
			arrHeader(0) = "거래통화"				' Header명(0)
			arrHeader(1) = "거래통화명"				' Header명(0)

		Case Else
			Exit Function
	End Select
    
	IsOpenPop = True
	
	Select Case iWhere
	Case 0, 3
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Case Else
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End Select
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function

Function OpenPopupDept(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("DeptPopupDt")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "DeptPopupDt", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = strCode						'부서코드 
	arrParam(1) = frm1.txtStartDt.Text			'날짜(Default:현재일)
	arrParam(2) = "1"							'부서권한(lgUsrIntCd)
'	If lgIntFlgMode = parent.OPMD_UMODE then
'		arrParam(3) = "T"									' 결의일자 상태 Condition  
'	Else
'		arrParam(3) = "F"									' 결의일자 상태 Condition  
'	End If

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	
	IsOpenPop = True

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
			"dialogWidth=500px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	End If
	
	frm1.txtDeptCd.value = arrRet(0)
	frm1.txtDeptNm.value = arrRet(1)
	Call txtDeptCD_Change()
	frm1.txtDeptCD.focus
	
	lgBlnFlgChgValue = True
End Function

 '==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'========================================================================================================= 
 '++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 
'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPopUp(Byval arrRet, Byval iWhere)

	With frm1
		Select Case iWhere
			Case 5		' 거래통화 
				.txtDocCur.value = arrRet(0)
				
				If Parent.gCurrency = UCase(Trim(frm1.txtDocCur.value)) Then
					frm1.txtXchRate.Text = "1"
				Else
					Call FncCalcRate
				End If
				
				call txtDocCur_OnChange()				
				lgBlnFlgChgValue = True	
				.txtDocCur.focus
		End Select

	End With
End Function

 '=========================================================================================================
'	Name : FncCalcRate()
'	Description : lookup exchange rate 
'========================================================================================================= 
Function FncCalcRate()
    Dim strXrate
    Dim strVal
    
    Err.Clear   
	
	FncCalcRate = False
	
	If Trim(frm1.txtDocCur.value) = "" then
		frm1.txtXchRate.Text = ""
	ElseIf Trim(frm1.txtDocCur.value) = Parent.gCurrency Then
		frm1.txtXchRate.Text  = "1"
	Else
		strVal = BIZ_PGM_ID2 & "?txtMode=" & "XRate"	        
 		strVal = strVal & "&txtLocCurr=" & Parent.gCurrency
 		strVal = strVal & "&txtToCurr=" & Trim(frm1.txtDocCur.value)
 	 	 	
		If frm1.txtStartDt.Text = "" Then		   
			strVal = strVal & "&txtAppDt=" & Trim("1900-01-01")
		Else
			strVal = strVal & "&txtAppDt=" & Trim(frm1.txtStartDt.Text) '☆: 조회 조건 데이타 
		End If	
			
 		Call RunMyBizASP(MyBizASP, strVal) 
 	End if
 	 	
 	FncCalcRate = True
 	
End Function

'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
Function CookiePage(ByVal Kubun)

	Dim strTemp

	Select Case Kubun		
	Case "FORM_LOAD"
		strTemp = ReadCookie("BANK_CD")
		Call WriteCookie("BANK_CD", "")
		
		If strTemp = "" then Exit Function
					
		frm1.txtBankCd.value = strTemp
				
		If Err.number <> 0 Then
			Err.Clear
			Call WriteCookie("BANK_CD", "")
			Exit Function 
		End If
				
		Call MainQuery()
	
	Case JUMP_PGM_ID_BANK_REP
		Call WriteCookie("BANK_CD", frm1.txtBankCd.value)

	Case Else
		Exit Function
	End Select
End Function	

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
 '==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()

    Call InitVariables							'⊙: Initializes local global variables
    Call LoadInfTB19029							'⊙: Load table , B_numeric_format
	
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)    
	Call ggoOper.LockField(Document, "N")		'⊙: Lock  Suitable  Field
	  
    '----------  Coding part  -------------------------------------------------------------
	Call FncSetToolBar("New")  
	Call InitComboBox
    Call SetDefaultVal
	
	Call ggoOper.FormatNumber(frm1.txtPaymDt, "31", "0", False)					'적금납입일 
	Call ggoOper.FormatNumber(frm1.txtPaymPeriod, "99", "0", False)				'납입주기 
	Call ggoOper.FormatNumber(frm1.txtPaymCnt, "9999", "0", True)				'불입횟수 
	Call ggoOper.FormatNumber(frm1.txtTotPaymCnt, "9999", "0", True)			'총불입횟수 

	Call CookiePage("FORM_LOAD") 
	frm1.txtBankCd.focus
	
	lgBlnFlgChgValue = False


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

 '**************************  3.2 HTML Form Element & Object Event처리  **********************************
'	Document의 TAG에서 발생 하는 Event 처리	
'	Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나 
'	Event간 충돌을 고려하여 작성한다.
'********************************************************************************************************* 

 '******************************  3.2.1 Object Tag 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 
 '-----------------------------------------------------------------------------------------------------
'	Name : SetXchRate()
'	Description : lookup exchange rate 
'--------------------------------------------------------------------------------------------------------- 
Function SetXchRate()
    Dim strXrate
    Dim strVal
    
    Err.Clear   
	
	SetXchRate = False
	
	If Trim(frm1.txtDocCur.value) = "" Then
		frm1.txtXchRate.Text = ""
	Else
		strVal = BIZ_PGM_ID2 & "?txtMode="   & "XchRate"	        
 		strVal = strVal & "&txtLocCur=" & Parent.gCurrency
 		strVal = strVal & "&txtDocCur=" & Trim(frm1.txtDocCur.value)
 	 	
		If Trim(frm1.txtStartDt.Text) = "" Then		   
			Call DisplayMsgBox("700110","X","X","X")
			'Msgbox "거래시작일을 입력하세요."
			Exit Function
		Else
			strVal = strVal & "&txtAppDt=" & Trim(frm1.txtStartDt.Text) '☆: 조회 조건 데이타 
		End If	
    
 		Call RunMyBizASP(MyBizASP, strVal) 	
 	End If

 	SetXchRate = True

End Function
 '-----------------------------------------------------------------------------------------------------
'	Name : SetCnclXchRate()
'	Description : lookup exchange rate 
'--------------------------------------------------------------------------------------------------------- 
Function SetCnclXchRate()
Dim strXrate
Dim strVal
    
    Err.Clear   
	
	SetCnclXchRate = False
	
	If Trim(frm1.txtDocCur.value) = "" Then
		frm1.txtCnclXchRate.Text = ""
	Else
		strVal = BIZ_PGM_ID2 & "?txtMode="   & "CnclXchRate"	        
 		strVal = strVal & "&txtLocCur=" & Parent.gCurrency
 		strVal = strVal & "&txtDocCur=" & Trim(frm1.txtDocCur.value)
 	 	
		If Trim(frm1.txtCnclDt.Text) = "" Then		   
			Call DisplayMsgBox("700111","X","X","X")
			'Msgbox "해약일자를 입력하세요."
			Exit Function
		Else
			strVal = strVal & "&txtAppDt=" & Trim(frm1.txtCnclDt.Text) '☆: 조회 조건 데이타 
		End If	
    
 		Call RunMyBizASP(MyBizASP, strVal) 	
 	End If

 	SetCnclXchRate = True
 	
End Function
 '-----------------------------------------------------------------------------------------------------
'	Name : Amt's fields'  event
'	Description : 
'--------------------------------------------------------------------------------------------------------- 

Sub txtStartDt_onblur()
 
End Sub

Sub txtCnclDt_onblur()
End Sub

'=======================================================================================================
'   Event Name : txtStartDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtStartDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtStartDt.Action = 7
    End If
End Sub

Sub txtEndDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtEndDt.Action = 7
    End If
End Sub

Sub txtCnclDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtCnclDt.Action = 7
    End If
End Sub

'=======================================================================================================
'   Event Name : txtStartDt_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtStartDt_Change()
	
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii
	Dim arrVal1, arrVal2


	If Trim(frm1.txtDeptCd.value) <> "" and Trim(frm1.txtStartDt.Text <> "") Then
	
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtStartDt.Text, gDateFormat,""), "''", "S") & "))"

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
		Else 
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
						
			For ii = 0 to Ubound(arrVal1,1) - 1
				arrVal2 = Split(arrVal1(ii), chr(11))
				If Trim(arrVal2(2)) <> Trim(frm1.hOrgChangeId.value) Then
					frm1.txtDeptCd.value = ""
					frm1.txtDeptNm.value = ""
					frm1.hOrgChangeId.value = Trim(arrVal2(2))
				End If
			Next
		End If
	End If
	
	Call FncCalcRate()
    lgBlnFlgChgValue = True
End Sub

Sub txtEndDt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtCnclDt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtBankRate_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtXchRate_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtPaymDt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtPaymPeriod_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtTotPaymCnt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtContractAmt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtContractLocAmt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtPaymAmt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtCnclXchRate_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtCnclCapitalAmt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtCnclCapLocAmt_Change()
    lgBlnFlgChgValue = True
End Sub


Sub txtCnclIntRate_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtCnclIntAmt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtCnclIntLocAmt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtCnclAmt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtCnclLocAmt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtPaymLocAmt_Change()
    lgBlnFlgChgValue = True
End Sub



Sub txtDeptCD_Change()

    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii
	If Trim(frm1.txtDeptCd.value) = "" and Trim(frm1.txtStartDt.Text = "") Then		Exit Sub

		'----------------------------------------------------------------------------------------
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtStartDt.Text, gDateFormat,""), "''", "S") & "))"			

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
		Else 
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
					
			For ii = 0 to Ubound(arrVal1,1) - 1
				arrVal2 = Split(arrVal1(ii), chr(11))
				frm1.hOrgChangeId.value = Trim(arrVal2(2))
			Next	
			
		End If
		'----------------------------------------------------------------------------------------

     lgBlnFlgChgValue = True
End Sub

Sub txtBankCd_OnChange()
End Sub

Sub txtBankAcctNo_OnChange()
End Sub

Sub txtDocCur_OnChange()
    IF CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
		Call CurFormatNumericOCX()
	END IF	    
	
End Sub

Sub Type_itemChange()
	lgBlnFlgChgValue = True
End Sub

'=====================================================
'예적금구분 변경시 
'=======================================================
Sub cboDpstFg_Change()

	Select Case Trim(frm1.cboDpstFg.value)
	Case "SV", ""

    	Call ggoOper.SetReqAttr(frm1.txtEndDt, "Q")			'만기일 
		Call ggoOper.SetReqAttr(frm1.txtPaymDt, "Q")		'납입일 
		Call ggoOper.SetReqAttr(frm1.txtPaymPeriod, "Q")	'납입주기 
		Call ggoOper.SetReqAttr(frm1.txtContractAmt, "Q")	'계약금액 
		Call ggoOper.SetReqAttr(frm1.txtContractLocAmt, "Q")'계약금액(자국)
		Call ggoOper.SetReqAttr(frm1.txtPaymAmt, "Q")		'월납입금액 
		Call ggoOper.SetReqAttr(frm1.txtPaymLocAmt, "Q")	'월납입금액(자국)

	Case Else
		Call ggoOper.SetReqAttr(frm1.txtEndDt, "D")			'만기일 
		Call ggoOper.SetReqAttr(frm1.txtPaymDt, "D")		'납입일 
		Call ggoOper.SetReqAttr(frm1.txtPaymPeriod, "D")	'납입주기 
		Call ggoOper.SetReqAttr(frm1.txtContractAmt, "D")	'계약금액 
		Call ggoOper.SetReqAttr(frm1.txtContractLocAmt, "D")'계약금액(자국)
		Call ggoOper.SetReqAttr(frm1.txtPaymAmt, "D")		'월납입금액 
		Call ggoOper.SetReqAttr(frm1.txtPaymLocAmt, "D")	'월납입금액(자국)

	End Select

End Sub

'=====================================================
'거래상태 변경시 
'=======================================================
Sub cboTransSts_Change()
	Select Case Trim(frm1.cboTransSts.value)
	Case "TR", ""
		frm1.txtCnclDt.Text         = ""
		frm1.txtCnclXchRate.Text    = ""
		frm1.txtCnclCapitalAmt.Text = ""
		frm1.txtCnclCapLocAmt.Text  = ""
		frm1.txtCnclIntRate.Text    = ""
		frm1.txtCnclIntAmt.Text     = ""
		frm1.txtCnclIntLocAmt.Text  = ""
		frm1.txtCnclAmt.Text        = ""
		
		Call ggoOper.SetReqAttr(frm1.txtCnclDt, "Q")			'해약일 
		Call ggoOper.SetReqAttr(frm1.txtCnclXchRate, "Q")		'해약환율 
		Call ggoOper.SetReqAttr(frm1.txtCnclCapitalAmt, "Q")	'해약원금 
		Call ggoOper.SetReqAttr(frm1.txtCnclCapLocAmt, "Q")		'해약원금(자국)
		Call ggoOper.SetReqAttr(frm1.txtCnclIntRate, "Q")		'해약이율 
		Call ggoOper.SetReqAttr(frm1.txtCnclIntAmt, "Q")		'해약이자 
		Call ggoOper.SetReqAttr(frm1.txtCnclIntLocAmt, "Q")		'해약이자(자국)

	Case Else
		Call ggoOper.SetReqAttr(frm1.txtCnclDt, "D")			'해약일 
		Call ggoOper.SetReqAttr(frm1.txtCnclXchRate, "D")		'해약환율 
		Call ggoOper.SetReqAttr(frm1.txtCnclCapitalAmt, "D")	'해약원금 
		Call ggoOper.SetReqAttr(frm1.txtCnclCapLocAmt, "D")		'해약원금(자국)
		Call ggoOper.SetReqAttr(frm1.txtCnclIntRate, "D")		'해약이율 
		Call ggoOper.SetReqAttr(frm1.txtCnclIntAmt, "D")		'해약이자 
		Call ggoOper.SetReqAttr(frm1.txtCnclIntLocAmt, "D")		'해약이자(자국)

	End Select
	
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
    '----------------------- 
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")				'☜ 바뀐부분 
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
    
      '-----------------------
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field    
    Call SetDefaultVal
    Call InitVariables	
    
  '-----------------------
    'Check condition area
    '----------------------- 
    If Not ChkField(Document, "1") Then		'⊙: This function check indispensable field
       Exit Function
    End If
    
    Call FncSetToolBar("New")
  '-----------------------
    'Query function call area
    '----------------------- 
    Call DbQuery									'☜: Query db data
       
    FncQuery = True									'⊙: Processing is OK
        
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False      '⊙: Processing is NG
    
  '-----------------------
    'Check previous data area
    '-----------------------
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"X","X") '☜ 바뀐부분 
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
    
  '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------    
    Call ggoOper.ClearField(Document, "A")  '⊙: Clear Condition/Contents(All) Field    
    
    Call InitVariables						'⊙: Initializes local global variables
    
    
    call txtDocCur_OnChange()
	Call cboTransSts_Change
	
	Call SetDefaultVal
		
	Call FncSetToolBar("New")

    lgBlnFlgChgValue = False
	frm1.txtBankCd.focus 
    
    FncNew = True							'⊙: Processing is OK

End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
    Dim IntRetCD 
    
    FncDelete = False														'⊙: Processing is NG
    
  '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002","X","X","X")  '☜ 바뀐부분 
        Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"X","X")  '☜ 바뀐부분 
    If IntRetCD = vbNo Then
        Exit Function
    End If
    
  '-----------------------
    'Delete function call area
    '-----------------------
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
	If lgBlnFlgChgValue = False Then
		IntRetCD = DisplayMsgBox("900001","X","X","X")  '☜ 바뀐부분 
		Exit Function
	End If
	    
    '-----------------------
    'Check content area
    '-----------------------
    If Not ChkField(Document, "1")     then                              '⊙: Check contents area
       Exit Function
    End If

    If Not ChkField(Document, "2") Then                             '⊙: Check contents area
       Exit Function
    End If
    
    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave				                                                '☜: Save db data
    
    FncSave = True                                                          '⊙: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
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
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	Call parent.FncPrint()
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
    Call parent.FncFind(Parent.C_SINGLE , False)                                     '☜:화면 유형, Tab 유무 
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")   '☜ 바뀐부분 
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
    Err.Clear                                                               '☜: Protect system from crashing
    
    DbDelete = False														'⊙: Processing is NG
    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtBankCd=" & EnCoding(Trim(frm1.txtBankCd.value))
    strVal = strVal & "&txtBankAcctNo=" & EnCoding(Trim(frm1.txtBankAcctNo.value))	'☜: 삭제 조건 데이타 

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

		ggoOper.FormatFieldByObjectOfCur .txtAmt,			.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtPaymAmt,		.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtContractAmt,	.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtCnclCapitalAmt,.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtCnclIntAmt,	.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtCnclAmt,       .txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		
	End With

End Sub

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 
Dim strVal
    
    Err.Clear																		'☜: Protect system from crashing
    
    DbQuery = False																	 '⊙: Processing is NG
    
	Call LayerShowHide(1)
	
	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001									'☜: 비지니스 처리 ASP의 상태 
	strVal = strVal & "&txtBankCd=" & Trim(frm1.txtBankCd.value)
	strVal = strVal & "&txtBankAcctNo=" & Trim(frm1.txtBankAcctNo.value)			'☆: 조회 조건 데이타 

	' 권한관리 추가 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 

	Call RunMyBizASP(MyBizASP, strVal)												'☜: 비지니스 ASP 를 가동 
	
    DbQuery = True																	'⊙: Processing is NG

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================

Function DbQueryOk()							'☆: 조회 성공후 실행로직 
    '-----------------------
    'Reset variables area
    '-----------------------
    Call ggoOper.LockField(Document, "Q")		'⊙: This function lock the suitable field

    Call cboTransSts_Change
   
	Call FncSetToolBar("Query")
    call txtDocCur_OnChange()
    	
    lgIntFlgMode = parent.OPMD_UMODE					'⊙: Indicates that current mode is Update mode
	lgBlnFlgChgValue = False
	
	frm1.txtBankCd.focus 
	Set gActiveElement = document.activeElement 
End Function


'========================================================================================
' Function Name : DBSave
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'========================================================================================

Function DbSave() 
Dim strVal

    Err.Clear																'☜: Protect system from crashing

	DbSave = False															'⊙: Processing is NG

	Call LayerShowHide(1)

	With frm1
		.txtMode.value = Parent.UID_M0002											'☜: 비지니스 처리 ASP 의 상태 
		.txtFlgMode.value = lgIntFlgMode
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

Function DbSaveOk()			'☆: 저장 성공후 실행 로직 
    Call InitVariables
	Call MainQuery()
End Function

'==========================================================
'툴바버튼 세팅 
'==========================================================
Function FncSetToolBar(Cond)
	Select Case UCase(Cond)
	Case "NEW"
		Call SetToolbar("1110100000001111")
	Case "QUERY"
		Call SetToolbar("1111100000011111")
	End Select
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag부 
'#########################################################################################################  -->
<BODY TABINDEX="-1" SCROLL="no">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenRefBankAcctNo(frm1.hTemp.value,1)">은행계좌참조</A>
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
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>은행</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtBankCd" NAME="txtBankCd" SIZE=10 MAXLENGTH=10   tag="12XXXU" ALT="은행코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenRefBankAcctNo(frm1.txtBankCd.Value, 2)">&nbsp;
														 <INPUT CLASS="clstxt" TYPE=TEXT ID="txtBankNM" NAME="txtBankNM" SIZE=20 MAXLENGTH=30  STYLE="TEXT-ALIGN: left" tag="24X" ALT="은행명"></TD>
									<TD CLASS=TD5 NOWRAP>계좌번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE="Text" ID="txtBankAcctNo" NAME="txtBankAcctNo" SIZE=18 MAXLENGTH=30  tag="12XXXU" ALT="계좌번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankAcctNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenRefBankAcctNo(frm1.txtBankAcctNo.Value, 3)"></TD>
								</TR>									
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>									
								<TR>
									<TD CLASS=TD5 NOWRAP>예적금구분</TD>
									<TD CLASS=TD6 NOWRAP><SELECT ID="cboDpstFg" NAME="cboDpstFg" ALT="예적금구분" STYLE="WIDTH: 132px" tag="14X" OnClick ="vbscript:Type_itemChange()" OnChange="vbscript:Call cboDpstFg_Change()" ><OPTION VALUE="" selected></OPTION></SELECT></TD>
									<TD CLASS=TD5 NOWRAP>예적금유형</TD>
									<TD CLASS=TD6 NOWRAP><SELECT ID="cboDpstType" NAME="cboDpstType" ALT="예적금유형" STYLE="WIDTH: 132px" tag="14X" OnClick ="vbscript:Type_itemChange()"><OPTION VALUE="" selected></OPTION></SELECT></TD>
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
								<TD CLASS=TD5 NOWRAP>부서</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtDeptCD" NAME="txtDeptCD" SIZE=10 MAXLENGTH=10  tag="22XXXU" ALT="부서"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" tag="23X" ONCLICK="vbscript:Call OpenPopupDept(frm1.txtDeptCD.Value, 1)">&nbsp;<INPUT CLASS="clstxt" TYPE=TEXT ID="txtDeptNm" NAME="txtDeptNm" SIZE=20 MAXLENGTH=40 STYLE="TEXT-ALIGN: left" tag="24X" ALT="부서"></TD>
								<TD CLASS=TD5 NOWRAP>거래시작일</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpStartDt name=txtStartDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="거래시작일" tag="22X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
							</TR>							
							<TR>
								<TD CLASS=TD5 NOWRAP>거래상태</TD>
								<TD CLASS=TD6 NOWRAP><SELECT ID="cboTransSts" NAME="cboTransSts" ALT="거래상태" STYLE="WIDTH: 132px" tag="22X" OnClick ="vbscript:Type_itemChange()" OnChange="vbscript:Call cboTransSts_Change()"><!--OPTION VALUE="" selected></OPTION--></SELECT></TD>
								<TD CLASS=TD5 NOWRAP>계좌유형</TD>
								<TD CLASS=TD6 NOWRAP><SELECT ID="cboBankAcctFg" NAME="cboBankAcctFg" ALT="계좌유형" STYLE="WIDTH: 132px" tag="2XX" OnClick ="vbscript:Type_itemChange()"><OPTION VALUE="" selected></OPTION></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>이율</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpBankRate name=txtBankRate style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE ALT="이율" tag="21X5" VIEWASTEXT></OBJECT>');</SCRIPT>&nbsp;%</TD>							
								<TD CLASS=TD5 NOWRAP>가입사유</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDpstNm" SIZE="35" MAXLENGTH="40" tag="21X" ALT="가입사유"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>통화</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtDocCur" NAME="txtDocCur" SIZE=15 MAXLENGTH=3  tag="22XXXU" ALT="통화"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCur" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtDocCur.Value, 5)"></TD>
								<TD CLASS=TD5 NOWRAP>환율</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpXchRate name=txtXchRate style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 132px" title=FPDOUBLESINGLE ALT="환율" tag="21X5Z" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>잔액</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 id=fpAmt name=txtAmt title=FPDOUBLESINGLE ALT="잔액" tag="24X2Z" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>잔액(자국)</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 id=fpLocAmt name=txtLocAmt title=FPDOUBLESINGLE ALT="잔액(자국)" tag="24X2Z" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR></TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>만기일</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpEndDt name=txtEndDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="만기일" tag="21X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>납입일</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpPaymDt name=txtPaymDt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 50px" title=FPDOUBLESINGLE ALT="납입일" tag="21X" VIEWASTEXT></OBJECT>');</SCRIPT>&nbsp;일</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>납입주기</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpPaymPeriod name=txtPaymPeriod style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 50px" title=FPDOUBLESINGLE ALT="납입주기" tag="21X" VIEWASTEXT></OBJECT>');</SCRIPT>&nbsp;개월</TD>
								<TD CLASS=TD5 NOWRAP>불입횟수</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpPaymCnt name=txtPaymCnt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 50px" title=FPDOUBLESINGLE ALT="불입횟수" tag="24X" VIEWASTEXT></OBJECT>');</SCRIPT>&nbsp;/&nbsp;
											  <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpTotPaymCnt name=txtTotPaymCnt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 50px" title=FPDOUBLESINGLE ALT="총불입횟수" tag="24X" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>월납입액</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 id=fpPaymAmt name=txtPaymAmt title=FPDOUBLESINGLE ALT="월납입액" tag="21X2Z" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>월납입액(자국)</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 id=fpPaymLocAmt name=txtPaymLocAmt title=FPDOUBLESINGLE ALT="월납입액(자국)" tag="21X2Z" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
							</TR>							
							<TR>
								<TD CLASS=TD5 NOWRAP>계약금액</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 id=fpContractAmt name=txtContractAmt title=FPDOUBLESINGLE ALT="계약금액" tag="21X2Z" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>계약금액(자국)</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 id=fpContractLocAmt name=txtContractLocAmt title=FPDOUBLESINGLE ALT="계약금액(자국)" tag="21X2Z" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR></TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>해약일자</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpCnclDt name=txtCnclDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="해약일자" tag="21X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>해약시환율</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpCnclXchRate name=txtCnclXchRate style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 132px" title=FPDOUBLESINGLE ALT="해약시환율" tag="24X5Z" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>해약시이자율</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpCnclIntRate name=txtCnclIntRate style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 132px" title=FPDOUBLESINGLE ALT="해약시이자율" tag="24X5Z" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
								<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>해약시원금</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 id=fpCnclCapitalAmt name=txtCnclCapitalAmt title=FPDOUBLESINGLE ALT="해약시원금" tag="24X2Z" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>해약시원금(자국)</TD>	
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 id=fpCnclCapLocAmt name=txtCnclCapLocAmt title=FPDOUBLESINGLE ALT="해약시원금(자국)" tag="24X2Z" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
							</TR>							
							<TR>
								<TD CLASS=TD5 NOWRAP>해약시이자</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpCnclIntAmt name=txtCnclIntAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 132px" title=FPDOUBLESINGLE ALT="해약시이자" tag="24X2Z"></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>해약시이자(자국)</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpCnclIntLocAmt name=txtCnclIntLocAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 132px" title=FPDOUBLESINGLE ALT="해약시이자(자국)" tag="24X2Z" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>해약금액</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 id=fpCnclAmt name=txtCnclAmt title=FPDOUBLESINGLE ALT="해약금액" tag="24X2Z"></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>해약금액(자국)</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 id=fpCnclLocAmt name=txtCnclLocAmt title=FPDOUBLESINGLE ALT="해약금액(자국)" tag="24X2Z" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>적요</TD>
								<TD CLASS="TD6" NOWRAP	COLSPAN=3><INPUT TYPE=TEXT NAME="txtDpstDesc" SIZE="70" MAXLENGTH="128" tag="21X" ALT="예적금적요"></TD>
							</TR>
<!--							<% Call SubFillRemBodyTD5656(1) %> -->
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=* ALIGN=RIGHT>
						<A HREF="VBSCRIPT:PgmJumpChk(JUMP_PGM_ID_BANK_REP)">은행정보등록</a>
					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<TEXTAREA CLASS="hidden" NAME="txtSpread"	tag="2"  Tabindex="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode"			tag="2"  Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"		tag="2"  Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"		tag="2"  Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId"	tag="2"  Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows"		tag="2"  Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="horgchangeid"		tag="2"  Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="hTemp"				tag="2"  Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthBizAreaCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthSubInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthUsrID"		tag="24" Tabindex="-1">
</FORM>
</BODY>
</HTML>

