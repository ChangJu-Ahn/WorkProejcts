<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Loan
'*  3. Program ID           : f4223ma1
'*  4. Program Name         : 차입금상환계획변경 
'*  5. Program Desc         : Register of Loan Repay
'*  6. Comproxy List        : FL0081, FL0088
'*  7. Modified date(First) : 2002/04/26
'*  8. Modified date(Last)  : 2003/05/19
'*  9. Modifier (First)     : Oh, Soo Min
'* 10. Modifier (Last)      : 
'* 11. Comment              :
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
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->								<!-- '⊙: 화면처리ASP에서 서버작업이 필요한 경우  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"     SRC="../ag/AcctCtrl.vbs">           </SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'☜: indicates that All variables must be declared in advance
<!-- #Include file="../../inc/lgvariables.inc" --> 
Const gIsShowLocal = "Y"
<%
Const gIsShowLocal = "Y"
%>
'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'⊙: 비지니스 로직 ASP명 
Const BIZ_PGM_ID =  "f4223mb1_ko441.asp"								'☆: 비지니스 로직 ASP명 >> air
Const BIZ_PGM_ID3 = "f4223mb3_ko441.asp"								'☆: 비지니스 로직 ASP명 >> air

Const JUMP_PGM_ID_LOAN_ENTRY = "f4203ma1"						'차입금등록 
Const COM_BIZ_EXCHRATE_ID = "../../inc/GetExchRate.asp"			'환율정보 비지니스 로직 ASP명 

'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
'⊙: Grid Columns

Dim C_PAY_PLAN_DT
Dim C_PAY_DT		
Dim C_PAY_OBJ_CD	
Dim C_PAY_OBJ_NM	
Dim C_PAY_PLAN_AMT
Dim C_PAY_PLAN_LOC_AMT
Dim C_PAY_AMT			
Dim C_PAY_LOC_AMT		
Dim C_RESL_FG_CD		
Dim C_RESL_FG_NM		
Dim C_DOC_CUR			
Dim C_DOC_CUR_PB		
Dim C_XCH_RATE		
Dim C_FLT_CONV_FG		
Dim C_LOAN_DESC		
Dim C_H_PAY_PLAN_DT	
Dim C_H_PAY_CHG_AMT	
Dim C_H_PAY_PLAN_AMT	
Dim C_COL_END			

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1.변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 

'Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
'Dim lgIntGrpCount              ' Group View Size를 조사할 변수 
'Dim lgIntFlgMode               ' Variable is for Operation Status

'Dim lgStrPrevKey

 '==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= 
 '----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 
Dim IsOpenPop
'Dim lgSortKey

 '++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 
Dim strPayObjCd1,strPayObjCd2
Dim strPayObjNm1,strPayObjNm2
Dim TotNewPrPlanAmt, TotOldPrPlanAmt

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

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgStrPrevKeyIndex = ""                                      '⊙: initializes Previous Key Index
    lgSortKey         = 1                                       '⊙: initializes sort direction

    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    
    lgSortKey = 1
End Sub

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
	C_PAY_PLAN_DT			= 1
	C_PAY_DT				= 2
	C_PAY_OBJ_CD			= 3
	C_PAY_OBJ_NM			= 4
	C_PAY_PLAN_AMT			= 5
	C_PAY_PLAN_LOC_AMT		= 6
	C_PAY_AMT				= 7
	C_PAY_LOC_AMT			= 8 
	C_RESL_FG_CD			= 9
	C_RESL_FG_NM			= 10
	C_DOC_CUR				= 11
	C_DOC_CUR_PB			= 12
	C_XCH_RATE				= 13
	C_FLT_CONV_FG			= 14
	C_LOAN_DESC				= 15
	C_H_PAY_PLAN_DT			= 16
	C_H_PAY_CHG_AMT			= 17
	C_H_PAY_PLAN_AMT		= 18
	C_COL_END				= 19
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
    Call ggoOper.ClearField(Document, "1")        '⊙: Condition field clear
'    frm1.txtDocCur.value	= gCurrency
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub  LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>
	<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ==============
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()

	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021206",,parent.gAllowDragDropSpread    
	
	With frm1.vspdData
		.MaxCols = C_COL_END
		
		.Col = .MaxCols				'☜: 공통콘트롤 사용 Hidden Column
		.ColHidden = True
		
		.ColsFrozen = C_PAY_DT

		.MaxRows = 0

		.ReDraw = False
		
		Call GetSpreadColumnPos("A")
		
		'ggoSpread.Spreadinit
		ggoSpread.SSSetDate   C_PAY_PLAN_DT,		"지급예정일자"		,15, 2,parent.gDateFormat 
		ggoSpread.SSSetDate   C_PAY_DT,				"지급일자"			,15, 2,parent.gDateFormat		
		ggoSpread.SSSetCombo  C_PAY_OBJ_CD,			"상환대상"			,30   
		ggoSpread.SSSetCombo  C_PAY_OBJ_NM,			"상환대상"			,17  
		ggoSpread.SSSetFloat  C_PAY_PLAN_AMT,		"지급예정액"		, 20, "A"  , ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
		ggoSpread.SSSetFloat  C_PAY_PLAN_LOC_AMT,	"지급예정액(자국)"	, 20, parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"								
		ggoSpread.SSSetFloat  C_PAY_AMT,			"지급액"			, 20, "A"  , ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
		ggoSpread.SSSetFloat  C_PAY_LOC_AMT,		"지급액(자국)"		, 20, parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"				
		ggoSpread.SSSetCombo  C_RESL_FG_CD,			"완료여부",10   
		ggoSpread.SSSetCombo  C_RESL_FG_NM,			"완료여부",10  
		ggoSpread.SSSetEdit   C_DOC_CUR,			"통화", 5, , ,3,2
		ggoSpread.SSSetButton C_DOC_CUR_PB		
		ggoSpread.SSSetFloat  C_XCH_RATE,			"상환환율"			, 10, parent.ggExchRateNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit   C_FLT_CONV_FG,		"유동성전환여부"	, 5, , , 5
		ggoSpread.SSSetEdit   C_LOAN_DESC,			"변경내역"			, 30, , , 128
		ggoSpread.SSSetDate   C_H_PAY_PLAN_DT,		"지급예정일자"		,15, 2,parent.gDateFormat 
		ggoSpread.SSSetFloat  C_H_PAY_PLAN_AMT,		"지급예정액"		, 20, "A"  , ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
		ggoSpread.SSSetFloat  C_H_PAY_CHG_AMT,		"지급예정액"		, 20, "A"  , ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
		.ReDraw = True
		
		Call ggoSpread.SSSetColHidden(C_PAY_OBJ_CD ,C_PAY_OBJ_CD	,True)
		Call ggoSpread.SSSetColHidden(C_RESL_FG_CD ,C_RESL_FG_CD	,True)
		Call ggoSpread.SSSetColHidden(C_DOC_CUR ,C_DOC_CUR	,True)
		Call ggoSpread.SSSetColHidden(C_DOC_CUR_PB ,C_DOC_CUR_PB	,True)
		Call ggoSpread.SSSetColHidden(C_XCH_RATE ,C_XCH_RATE	,True)
		Call ggoSpread.SSSetColHidden(C_FLT_CONV_FG ,C_FLT_CONV_FG	,True)
		Call ggoSpread.SSSetColHidden(C_H_PAY_PLAN_DT ,C_H_PAY_PLAN_DT	,True)
		Call ggoSpread.SSSetColHidden(C_H_PAY_PLAN_AMT ,C_H_PAY_PLAN_AMT	,True)
		Call ggoSpread.SSSetColHidden(C_H_PAY_CHG_AMT ,C_H_PAY_CHG_AMT	,True)
		
		Call SetSpreadLock

    End With
    
End Sub

'================================== 2.2.4 SetSpreadLock() =============================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
	Dim RowCnt
	
	ggoSpread.Source = frm1.vspdData
	
	With frm1				
		.vspdData.ReDraw = False			    
		For RowCnt = 1 To .vspdData.MaxRows								
			
			.vspdData.Col = C_RESL_FG_CD
			.vspdData.Row = RowCnt					
			If .vspdData.text = "Y" Then			    
			    				
				ggoSpread.SpreadLock	C_PAY_PLAN_DT    ,RowCnt, RowCnt
				ggoSpread.SpreadLock	C_PAY_DT         ,RowCnt, RowCnt
				ggoSpread.SpreadLock	C_PAY_OBJ_NM	 ,RowCnt, RowCnt
				ggoSpread.SpreadLock	C_PAY_PLAN_AMT	 ,RowCnt, RowCnt
				ggoSpread.SpreadLock	C_PAY_PLAN_LOC_AMT,RowCnt, RowCnt
				ggoSpread.SpreadLock	C_PAY_AMT		 ,RowCnt, RowCnt
				ggoSpread.SpreadLock	C_PAY_LOC_AMT	 ,RowCnt, RowCnt
				ggoSpread.SpreadLock	C_RESL_FG_NM	 ,RowCnt, RowCnt
				ggoSpread.SpreadLock	C_LOAN_DESC		 ,RowCnt, RowCnt	
				ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1      			
							
			Else			
				.vspdData.Col = C_FLT_CONV_FG
				If .vspdData.text = "CV" Then	
									
					ggoSpread.SpreadLock	C_PAY_PLAN_DT    ,RowCnt, RowCnt
					ggoSpread.SpreadLock	C_PAY_DT         ,RowCnt, RowCnt
					ggoSpread.SpreadLock	C_PAY_OBJ_NM	 ,RowCnt, RowCnt
					ggoSpread.SpreadLock	C_PAY_PLAN_AMT	 ,RowCnt, RowCnt
					ggoSpread.SpreadLock	C_PAY_PLAN_LOC_AMT,RowCnt, RowCnt
					ggoSpread.SpreadLock	C_PAY_AMT		 ,RowCnt, RowCnt
					ggoSpread.SpreadLock	C_PAY_LOC_AMT	 ,RowCnt, RowCnt
					ggoSpread.SpreadLock	C_RESL_FG_NM	 ,RowCnt, RowCnt
					ggoSpread.SpreadLock	C_LOAN_DESC		 ,RowCnt, RowCnt
					ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1      
				Else	
					ggoSpread.SpreadUnLock	C_PAY_PLAN_DT    ,RowCnt, RowCnt		
					ggoSpread.SSSetRequired	C_PAY_PLAN_DT    ,RowCnt, RowCnt		
					ggoSpread.SpreadLock	C_PAY_DT         ,RowCnt, RowCnt		
					ggoSpread.SpreadLock	C_PAY_OBJ_NM	 ,RowCnt, RowCnt
					ggoSpread.SpreadUnLock	C_PAY_PLAN_AMT   ,RowCnt, RowCnt		
					ggoSpread.SSSetRequired C_PAY_PLAN_AMT	 ,RowCnt, RowCnt					
					'지급예정액(자국)Locking해제 >> air | ggoSpread.SpreadLock    C_PAY_PLAN_LOC_AMT,RowCnt, RowCnt
					ggoSpread.SpreadUnLock  C_PAY_PLAN_LOC_AMT,RowCnt, RowCnt
					ggoSpread.SSSetRequired C_PAY_PLAN_LOC_AMT,RowCnt, RowCnt
					ggoSpread.SpreadLock	C_PAY_AMT		 ,RowCnt, RowCnt
					ggoSpread.SpreadLock	C_PAY_LOC_AMT	 ,RowCnt, RowCnt
					ggoSpread.SpreadLock	C_RESL_FG_NM	 ,RowCnt, RowCnt								
					ggoSpread.SpreadUnLock  C_LOAN_DESC		 ,RowCnt, RowCnt	
					ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1      
				
				End If
			End If 		
		Next
		
		.vspdData.ReDraw = True	
	End With		
	
		If UCase(frm1.txtDocCur.value) = UCase(parent.gCurrency) Then
			ggoSpread.SpreadLock C_XCH_RATE,	-1, C_XCH_RATE
		Else
			ggoSpread.SpreadUnLock C_XCH_RATE,		-1, C_XCH_RATE
		End If	
	
End Sub

'================================== 2.2.5 SetSpreadColor() ============================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal lRow)	
	ggoSpread.Source = frm1.vspdData
	
	With frm1.vspdData
		.col = C_RESL_FG_CD
		.text = "N"
		.col = C_RESL_FG_NM
		.text = "미상환"
	End With		
    
    With frm1
		.vspdData.ReDraw = False

		ggoSpread.SSSetRequired		C_PAY_PLAN_DT,		lRow, lRow
		ggoSpread.SSSetProtected	C_PAY_DT,			lRow, lRow		
		ggoSpread.SSSetRequired		C_PAY_OBJ_CD,		lRow, lRow		
		ggoSpread.SSSetRequired		C_PAY_OBJ_NM,		lRow, lRow		
		ggoSpread.SSSetRequired		C_PAY_PLAN_AMT,		lRow, lRow
		'지급예정액(자국)Locking해제 >> air | ggoSpread.SSSetProtected	C_PAY_PLAN_LOC_AMT,	lRow, lRow	
		ggoSpread.SSSetRequired		C_PAY_PLAN_LOC_AMT,	lRow, lRow						
		ggoSpread.SSSetProtected	C_PAY_AMT,			lRow, lRow		
		ggoSpread.SSSetProtected	C_PAY_LOC_AMT,		lRow, lRow

		ggoSpread.SSSetProtected	C_RESL_FG_CD,		lRow, lRow		
		ggoSpread.SSSetProtected	C_RESL_FG_NM,		lRow, lRow
		ggoSpread.SpreadUnLock		C_LOAN_DESC,		lRow, lRow

		.vspdData.ReDraw = True    
    End With    
    
	If UCase(frm1.txtDocCur.value) = UCase(parent.gCurrency) Then
		ggoSpread.SpreadLock C_XCH_RATE,	lRow, C_XCH_RATE,		lRow
	Else
		ggoSpread.SpreadUnLock C_XCH_RATE,		lRow, C_XCH_RATE,		lRow
	End If

End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
       
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            
			C_PAY_PLAN_DT			= iCurColumnPos(1)
			C_PAY_DT				= iCurColumnPos(2)
			C_PAY_OBJ_CD			= iCurColumnPos(3)
			C_PAY_OBJ_NM			= iCurColumnPos(4)
			C_PAY_PLAN_AMT			= iCurColumnPos(5)
			C_PAY_PLAN_LOC_AMT		= iCurColumnPos(6)
			C_PAY_AMT				= iCurColumnPos(7)
			C_PAY_LOC_AMT			= iCurColumnPos(8) 
			C_RESL_FG_CD			= iCurColumnPos(9)
			C_RESL_FG_NM			= iCurColumnPos(10)
			C_DOC_CUR				= iCurColumnPos(11)
			C_DOC_CUR_PB			= iCurColumnPos(12)
			C_XCH_RATE				= iCurColumnPos(13)
			C_FLT_CONV_FG			= iCurColumnPos(14)
			C_LOAN_DESC				= iCurColumnPos(15)
			C_H_PAY_PLAN_DT			= iCurColumnPos(16)
			C_H_PAY_CHG_AMT			= iCurColumnPos(17)
			C_H_PAY_PLAN_AMT		= iCurColumnPos(18)
			C_COL_END				= iCurColumnPos(19)
    End Select    
    
End Sub


'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
 '------------------------------------------ OpenLoanNo() -------------------------------------------------
'	Name : OpenLoanNo()
'	Description : Loan Number PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopup(Byval strCode,byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function
	
	Select Case iWhere
    Case 1
       	arrParam(0) = "통화코드팝업"	
		arrParam(1) = "B_CURRENCY"				
		arrParam(2) = strCode
		arrParam(3) = "" 
		arrParam(4) = ""			
		arrParam(5) = "통화코드"			
	
		arrField(0) = "CURRENCY"	
		arrField(1) = "CURRENCY_DESC"	
    
		arrHeader(0) = "통화코드"		
		arrHeader(1) = "통화명"		
    
    Case Else
		Exit Function
    End Select    
	
	IsOpenPop = True
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtLoanNo.focus
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function

 '++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 
'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPopUp(Byval arrRet, Byval iWhere)
	Dim Row
	
	With frm1
		Select Case iWhere
		Case 1
		    .vspdData.Col = C_DOC_CUR
			.vspdData.Text = arrRet(0)
			Row = .vspdData.ActiveRow
			Call vspdData_Change(.vspdData.Col,.vspdData.Row )	
		End Select
	End With

End Function

'================================================================
'차입금번호 팝업 
'================================================================
Function OpenPopupLoan()
	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(3)	

	If IsOpenPop = True Then Exit Function
	iCalledAspName = AskPRAspName("f4232ra1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "f4232ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True
    
	arrRet = window.showModalDialog(iCalledAspName & "?PGM=" & gStrRequestMenuID , Array(window.parent,arrParam), _
		     "dialogWidth=850px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = ""  Then
		frm1.txtLoanNo.focus
		Exit Function
	Else
		frm1.txtLoanNo.value = arrRet(0)
		frm1.txtLoanNm.value = arrRet(1)
	End If
	
	frm1.txtLoanNo.focus
End Function

'============================================================
'회계전표 팝업 
'============================================================
Function OpenPopupGL()

	Dim arrRet
	Dim arrParam(1)	

	If IsOpenPop = True Then Exit Function

	With frm1.vspdData
		If .ActiveRow > 0 Then
			.Row = .ActiveRow
			.Col = C_GL_NO
			arrParam(0) = Trim(.Text)	'회계전표번호 
			arrParam(1) = ""			'Reference번호 
		Else
			Call DisplayMsgBox("900025","X","X","X")
			Exit Function
		End If
	End With

	IsOpenPop = True
   
	arrRet = window.showModalDialog("../../ComAsp/a5120ra1.asp", Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
End Function

'=========================================== InitComboBox ================================================
'	Name : InitComboBox
'	Description : 
'=========================================================================================================== 
Sub InitComboBox()
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1020", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboLoanFg ,lgF0  ,lgF1  ,Chr(11))    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F2020", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    lgF0 = replace(lgF0, Chr(11), vbTab)
	ggoSpread.SetCombo lgF0, C_RESL_FG_CD
	lgF1 = replace(lgF1, Chr(11), vbTab)
	ggoSpread.SetCombo lgF1, C_RESL_FG_NM
End Sub
'==========================================================================================
'   Event Name : InitData()
'   Event Desc : Combo 변경 이벤트 
'==========================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		For intRow = 1 To .MaxRows
			.Row = intRow
			.Col = C_PAY_OBJ_CD
			intIndex = .value
			.Col = C_PAY_OBJ_NM
			.value = intIndex
		
			.Col = C_RESL_FG_CD
			intIndex = .value
			.Col = C_RESL_FG_NM
			.value = intIndex		
		Next
	End With
		
End Sub


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
				
		Call MainQuery()
	
	Case JUMP_PGM_ID_LOAN_ENTRY
		Call WriteCookie("LOAN_NO", frm1.txtLoanNo.value)
	
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
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")			'⊙: "Will you destory previous data"
		if IntRetCD = vbNo Then
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

    Call LoadInfTB19029                           '⊙: Load table , B_numeric_format

	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	        
    Call InitSpreadSheet                          '⊙: Setup the Spread Sheet
    Call InitVariables                            '⊙: Initializes local global Variables
    Call InitComboBox
    
    Call SetDefaultVal
    'Call CookiePage("FORM_LOAD")
    '----------  Coding part  -------------------------------------------------------------
	Call FncNew()
	Call FncSetToolBar("New")
    
    frm1.txtLoanNo.focus 
	Set gActiveElement = document.activeElement
	
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
Sub vspdData_Click(ByVal Col, ByVal Row)   
 
    
    Call SetPopupMenuItemInf("1101111111")
    gMouseClickStatus = "SPC"	'Split 상태코드 
    Set gActiveSpdSheet = frm1.vspdData

	If Row = 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort
			lgSortKey = 2
		Else
			ggoSpread.SSSort ,lgSortKey
			lgSortKey = 1
		End If 
	Else
		frm1.vspdData.Row = Row
	End If
    
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
	
	If Row<=0 then
		Exit Sub
	End If
	If Frm1.vspdData.MaxRows =0 then
		Exit Sub
	End if
	
End Sub


'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub
'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub


'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    If Col <= C_PAY_PLAN_DT Or NewCol <= C_PAY_PLAN_DT Then
        Cancel = True
        Exit Sub
    End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
	Dim strval, Oldamt, Newamt, strval2
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	Dim i 
	
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

	With frm1
		.vspdData.Redraw = False
	
		Select Case Col
			Case C_RESL_FG_CD, C_RESL_FG_NM
				.vspdData.Col = C_RESL_FG_CD
				.vspdData.Row = Row
				strval = .vspdData.text 
				IF CommonQueryRs( "MINOR_CD" , "B_MINOR  " , "MAJOR_CD = " & FilterVar("F2020", "''", "S") & "  AND MINOR_CD =  " & FilterVar(strval , "''", "S") & "", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then 
					Select Case UCase(lgF0)
						Case "Y" & Chr(11)			' 상환(Repay) - Not Change(Locking)				
							ggoSpread.SpreadLock	 C_PAY_PLAN_DT, Row, Row									
							ggoSpread.SpreadLock	 C_PAY_DT, Row, Row		
							ggoSpread.SpreadLock	 C_PAY_OBJ_NM, Row, Row	
							ggoSpread.SpreadLock	 C_PAY_OBJ_CD, Row, Row	
							ggoSpread.SpreadLock	 C_PAY_PLAN_AMT, Row, Row
							ggoSpread.SpreadLock	 C_PAY_PLAN_LOC_AMT, Row, Row	'>> air															
							ggoSpread.SpreadLock	 C_PAY_AMT, Row, Row									
							
							ggoSpread.SpreadLock	 C_RESL_FG_NM, Row, Row			
							ggoSpread.SpreadLock	 C_RESL_FG_CD, Row, Row
							ggoSpread.SpreadLock	 C_LOAN_DESC, Row, Row						
							
						Case "N" & Chr(11)			' 미상환(Plan)
							ggoSpread.SSSetRequired C_PAY_PLAN_DT, Row, Row
							.vspdData.COL = C_PAY_PLAN_DT
							ggoSpread.SpreadLock	C_PAY_DT, Row, Row	
							.vspdData.COL = C_PAY_DT
							ggoSpread.SSSetRequired	C_PAY_OBJ_NM, Row, Row
							.vspdData.COL = C_PAY_OBJ_NM
							ggoSpread.SSSetRequired	C_PAY_OBJ_CD, Row, Row
							.vspdData.COL = C_PAY_OBJ_CD
							ggoSpread.SSSetRequired C_PAY_PLAN_AMT, Row, Row
							ggoSpread.SSSetRequired C_PAY_PLAN_LOC_AMT, Row, Row	'>> air
							.vspdData.COL = C_PAY_PLAN_AMT
							ggoSpread.SpreadLock	C_PAY_AMT, Row, Row
							.vspdData.COL = C_PAY_AMT
							ggoSpread.SpreadLock	C_RESL_FG_NM, Row, Row
							.vspdData.COL = C_RESL_FG_NM
							ggoSpread.SpreadLock	C_RESL_FG_CD, Row, Row
							.vspdData.COL = C_RESL_FG_CD
							ggoSpread.SpreadUnLock	C_LOAN_DESC, Row, Row
							.vspdData.COL = C_LOAN_DESC																							
					End Select				
				End If	
			Case C_PAY_PLAN_AMT										
			'계획금액 변경시, 예정금액에 sum
				.vspdData.Row = row
				.vspdData.Col = C_PAY_OBJ_CD			
				strval2 = .vspdData.text			'상환대상 
					
				.vspdData.Col = C_PAY_PLAN_AMT				
				Newamt = UniCdbl(.vspdData.text)

				.vspdData.Col = C_H_PAY_CHG_AMT
				Oldamt = UniCdbl(.vspdData.text)
						

				If (strVal2 = "SL" or strVal2 = "LL" or strVal2 = "SN" or strVal2 = "LN" ) Then
					.txtTotPrPlanAmt.text = UNIFormatNumberByCurrecny(UniCdbl(.txtTotPrPlanAmt.text) - (Oldamt - Newamt), frm1.txtDocCur.value, Parent.ggAmtOfMoneyNo)
					.vspdData.text = UNIFormatNumberByCurrecny(Newamt, frm1.txtDocCur.value, Parent.ggAmtOfMoneyNo)
				ElseIf strVal2 = "DI" or strVal2 = "AI" Then			
					.txtTotIntPlanAmt.text = UNIFormatNumberByCurrecny(UniCdbl(.txtTotIntPlanAmt.text) - (Oldamt - Newamt), frm1.txtDocCur.value, Parent.ggAmtOfMoneyNo)
					.vspdData.text = UNIFormatNumberByCurrecny(Newamt, frm1.txtDocCur.value, Parent.ggAmtOfMoneyNo)
				End If
			Case C_PAY_OBJ_NM										
				.vspdData.Row = row
				.vspdData.Col = C_PAY_OBJ_CD

				strval2 = .vspdData.text			'상환대상 
					
				.vspdData.Col = C_PAY_PLAN_AMT				
				Newamt = UniCdbl(.vspdData.text)

				.vspdData.Col = C_H_PAY_CHG_AMT
				Oldamt = UniCdbl(.vspdData.text)
						
				If (strVal2 = "SL" or strVal2 = "LL" or strVal2 = "SN" or strVal2 = "LN" ) Then
					.txtTotPrPlanAmt.text = UNIFormatNumberByCurrecny(UniCdbl(.txtTotPrPlanAmt.text) + Newamt, frm1.txtDocCur.value, Parent.ggAmtOfMoneyNo)
					.txtTotIntPlanAmt.text = UNIFormatNumberByCurrecny(UniCdbl(.txtTotIntPlanAmt.text) - Oldamt, frm1.txtDocCur.value, Parent.ggAmtOfMoneyNo)
					.vspdData.text = UNIFormatNumberByCurrecny(Newamt, frm1.txtDocCur.value, Parent.ggAmtOfMoneyNo)
				ElseIf strVal2 = "DI" or strVal2 = "AI" Then			
					.txtTotIntPlanAmt.text = UNIFormatNumberByCurrecny(UniCdbl(.txtTotIntPlanAmt.text) + Newamt, frm1.txtDocCur.value, Parent.ggAmtOfMoneyNo)
					.txtTotPrPlanAmt.text = UNIFormatNumberByCurrecny(UniCdbl(.txtTotPrPlanAmt.text) - Oldamt, frm1.txtDocCur.value, Parent.ggAmtOfMoneyNo)
					.vspdData.text = UNIFormatNumberByCurrecny(Newamt, frm1.txtDocCur.value, Parent.ggAmtOfMoneyNo)
				End If
		End Select

		.vspdData.Redraw = True
	End With
End Sub
'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)
	Dim strTemp

	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		
    End With
End Sub

Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    With frm1.vspdData 
		If Row >= NewRow Then
			Exit Sub
		End If
    End With
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
 '----------  Coding part  -------------------------------------------------------------   
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	'☜: 재쿼리 체크 
		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			Call DbQuery
		End If
    End if
        
End Sub

'==========================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	 '----------  Coding part  -------------------------------------------------------------   
	 ' 이 Template 화면에서는 없는 로직임, 콤보(Name)가 변경되면 콤보(Code, Hidden)도 변경시켜주는 로직 
	With frm1.vspdData
		.Row = Row
    
		Select Case Col			
			Case C_PAY_OBJ_NM	
				.Col = C_PAY_OBJ_NM
				intIndex = .Value
				.Col = C_PAY_OBJ_CD
				.Value = intIndex							
		End Select
	End With
End Sub

'Sub txtLoanNo_onChange()
'	frm1.txtLoanNm.value = ""
'End Sub

Sub txtDocCur_OnChange()
	
    If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then                     
		Call CurFormatNumericOCX()
		Call CurFormatNumSprSheet()
	End If
		    
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
    
    FncQuery = False          '⊙: Processing is NG
    Err.Clear                 '☜: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")			'⊙: "Will you destory previous data"
		if IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	
    '-----------------------
    'Check condition area
    '-----------------------
	If Not chkField(Document, "1") Then	'⊙: This function check indispensable field
       Exit Function
    End If
    
	Call ggoOper.ClearField(Document, "2")
		    
    '-----------------------
    'Erase contents area
    '-----------------------
	frm1.vspdData.maxrows = 0 
    Call InitVariables							  '⊙: Initializes local global variables
	
	Call FncSetToolBar("New")
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery																'☜: Query db data
       
    FncQuery = True																'⊙: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
	Dim IntRetCD 
    
    FncNew = False                  '⊙: Processing is NG
    Err.Clear                       '☜: Protect system from crashing
    'On Error Resume Next           '☜: Protect system from crashing
    
    '-----------------------
    'Check previous data area
    '-----------------------
    ' 변경된 내용이 있는지 확인한다.
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015",parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
	
    Call ggoOper.ClearField(Document, "1")     '⊙: Clear Condition Field
	Call ggoOper.LockField(Document, "N")      '⊙: Lock  Suitable  Field
	frm1.vspddata.maxrows = 0
    'Call InitVariables                         '⊙: Initializes local global variables
    'Call InitComboBox    
    Call SetDefaultVal
    
    Call FncSetToolBar("New")
    
    'SetGridFocus
    FncNew = True                              '⊙: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
	Dim IntRetCD 
    
    FncDelete = False            '⊙: Processing is NG
    Err.Clear                    '☜: Protect system from crashing
    'On Error Resume Next        '☜: Protect system from crashing
    
    '-----------------------
    'Precheck area
    '-----------------------
    ' Update 상태인지를 확인한다.
    If lgIntFlgMode <> parent.OPMD_UMODE Then        'Check if there is retrived data
        Call DisplayMsgbox("900002","X","X","X")
        Exit Function
    End If
    
    '-----------------------
    'Delete function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")		            '⊙: "Will you destory previous data"
	
	If IntRetCD = vbNo Then
		Exit Function
	End If
    
    If DbDelete = False Then											  '☜: Delete db data
       Exit Function                        
    End If
    
    '-----------------------
    'Erase condition area
    '-----------------------
	Call ggoOper.ClearField(Document, "1")								  '⊙: Clear Condition Field
    FncDelete = True													 '⊙: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
	Dim IntRetCD 
    
    FncSave = False            '⊙: Processing is NG
    Err.Clear                  '☜: Protect system from crashing
        
    ggoSpread.Source = frm1.vspdData                          '⊙: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = False Then                   '⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","X","X","X")        '⊙: Display Message(There is no changed data.)
        Exit Function
    End If
    

    If UniCdbl(frm1.txtTotPrPlanAmt.text) > UniCdbl(frm1.txtLoanBalAmt.text) Then  
		Call DisplayMsgBox("141155","X","X","X")
		Exit Function
	End If
    
    ggoSpread.Source = frm1.vspdData                          '⊙: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck         Then              '⊙: Check required field(Multi area)
       Exit Function
    End If
    
    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave				                                                  '☜: Save db data

	 FncSave = True                                                           '⊙: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	If frm1.vspdData.MaxRows < 1 Then Exit Function
	
	With frm1
		.vspdData.ReDraw = False
	
		ggoSpread.Source = .vspdData
		ggoSpread.CopyRow

		Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData,.vspdData.ActiveRow,.vspdData.ActiveRow,.txtDocCur.value,C_PAY_PLAN_AMT,   "A" ,"I","X","X")
		Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData,.vspdData.ActiveRow,.vspdData.ActiveRow,.txtDocCur.value,C_PAY_AMT,   "A" ,"I","X","X")
		
		.vspdData.Col = C_PAY_DT
		.vspdData.Text = ""
		
		Call SetSpreadColor(frm1.vspdData.ActiveRow)

		.vspdData.ReDraw = True
	End With
	
	frm1.vspdData.Focus
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel()
	Dim Row
	
	If frm1.vspdData.MaxRows < 1 Then Exit Function
	
	With frm1
		.vspdData.Redraw = False
		
		ggoSpread.Source = frm1.vspdData	
		ggoSpread.EditUndo

		Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData,.vspdData.ActiveRow,.vspdData.ActiveRow,.txtDocCur.value,C_PAY_PLAN_AMT,   "A" ,"I","X","X")
		Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData,.vspdData.ActiveRow,.vspdData.ActiveRow,.txtDocCur.value,C_PAY_AMT,   "A" ,"I","X","X")
		
		Call InitData

		Row = .vspdData.ActiveRow		
'		ggoSpread.SpreadUnLock	C_BANK_CD,		Row, C_BANK_CD,		Row
'		ggoSpread.SpreadUnLock	C_BANK_PB,		Row, C_BANK_PB,		Row
'		ggoSpread.SpreadUnLock	C_BANK_ACCT,	Row, C_BANK_ACCT,	Row
'		ggoSpread.SpreadUnLock	C_BANK_ACCT_PB,	Row, C_BANK_ACCT_PB,Row
		
		.vspdData.Redraw = True
	End With
End Function
'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow
    Dim imRow2
    Dim iCurRowPos
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    FncInsertRow = False                                                         '☜: Processing is NG

    If IsNumeric(Trim(pvRowCnt)) then
        imRow = CInt(pvRowCnt)
    Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then
            Exit Function
        End If
    End If
	
	With frm1
		iCurRowPos = .vspdData.ActiveRow + 1

        .vspdData.ReDraw = False
        .vspdData.focus
        for imRow2 = 1 to imRow 
            ggoSpread.Source = .vspdData
            ggoSpread.InsertRow ,1

			Call SetSpreadColor(.vspdData.ActiveRow)
        Next

		Call ReFormatSpreadCellByCellByCurrency2(.vspdData,iCurRowPos, iCurRowPos + imRow,.txtDocCur.value,C_PAY_PLAN_AMT,   "A" ,"I","X","X")
		Call ReFormatSpreadCellByCellByCurrency2(.vspdData,iCurRowPos, iCurRowPos + imRow,.txtDocCur.value,C_PAY_AMT,   "A" ,"I","X","X")

        .vspdData.ReDraw = True
    End With
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
    Dim lDelRows
    FncDeleteRow = False                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If Frm1.vspdData.MaxRows < 1 then
       Exit function
	End if	
	
    With Frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncDeleteRow = True         
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
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
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)												'☜: 화면 유형 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                         '☜:화면 유형, Tab 유무 
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()
	
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
    
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
	Call SetSpreadLock
	Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData,-1,-1,frm1.txtDocCur.value,C_PAY_PLAN_AMT,"A" ,"I","X","X")
	Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData,-1,-1,frm1.txtDocCur.value,C_PAY_AMT,"A" ,"I","X","X")
End Sub

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")                '데이타가 변경되었습니다. 종료 하시겠습니까?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function


'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()									
	With frm1		
		ggoOper.FormatFieldByObjectOfCur .txtLoanAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec		
		ggoOper.FormatFieldByObjectOfCur .txtLoanBalAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec		
		ggoOper.FormatFieldByObjectOfCur .txtTotPrRdpAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec		
		ggoOper.FormatFieldByObjectOfCur .txtTotIntPayAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec		
		ggoOper.FormatFieldByObjectOfCur .txtTotPrPlanAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec		
		ggoOper.FormatFieldByObjectOfCur .txtTotIntPlanAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec  		
	End With
End Sub

'===================================== CurFormatNumSprSheet()  ======================================
'	Name : CurFormatNumSprSheet()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet()	
	Dim ii 
	With frm1
		For ii = 1 To .vspdData.MaxRows 
			Call FixDecimalPlaceByCurrency2(frm1.vspdData,ii,.txtDocCur.value,C_PAY_PLAN_AMT,"A" ,"X","X")
			Call FixDecimalPlaceByCurrency2(frm1.vspdData,ii,.txtDocCur.value,C_PAY_AMT,"A" ,"X","X")
			Call FixDecimalPlaceByCurrency2(frm1.vspdData,ii,.txtDocCur.value,C_H_PAY_CHG_AMT,"A" ,"X","X")
			Call FixDecimalPlaceByCurrency2(frm1.vspdData,ii,.txtDocCur.value,C_H_PAY_PLAN_AMT,"A" ,"X","X")
      	Next
      	
       Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData,1,-1,.txtDocCur.value,C_PAY_PLAN_AMT,"A" ,"I","X","X")         
       Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData,1,-1,.txtDocCur.value,C_PAY_AMT,"A" ,"I","X","X")         
       Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData,1,-1,.txtDocCur.value,C_H_PAY_CHG_AMT,"A" ,"I","X","X")         
       Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData,1,-1,.txtDocCur.value,C_H_PAY_PLAN_AMT,"A" ,"I","X","X")         
	End With
End Sub  

'*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'********************************************************************************************************* 
'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Dim strVal

	Call LayerShowHide(1)
    
    DbQuery = False
    Err.Clear                '☜: Protect system from crashing
    
    With frm1
        
		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID3 & "?txtMode="	& parent.UID_M0001
			strVal = strVal & "&txtLoanNo="		& Trim(.htxtLoanNo.value)	'조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey="	& lgStrPrevKey
			strVal = strVal & "&txtMaxRows="	& .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_ID3 & "?txtMode="	& parent.UID_M0001
			strVal = strVal & "&txtLoanNo="		& Trim(.txtLoanNo.value)	'조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey="	& lgStrPrevKey
			strVal = strVal & "&txtMaxRows="	& .vspdData.MaxRows
		End If
    
	    Call RunMyBizASP(MyBizASP, strVal)		'☜: 비지니스 ASP 를 가동 
    End With
    
    DbQuery = True
	
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()			'☆: 조회 성공후 실행로직	
			
	Call InitData()
	Call SetSpreadLock		
    
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE	'⊙: Indicates that current mode is Update mode
    
    Call ggoOper.LockField(Document, "Q")	'⊙: This function lock the suitable field    
	Call FncSetToolBar("Query")
	
	Call txtDocCur_OnChange()	             											'⊙: Initializes local global variables ()														'⊙: Initializes local global variables 		

	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.focus
	End If
	
	Set gActiveElement = document.activeElement 
End Function

'========================================================================================
' Function Name : DbSave()
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 
	Dim lRow        
	Dim lGrpCnt     
	Dim strVal, strDel, iColSep, iRowSep
	
	Call LayerShowHide(1)
	
    DbSave = False				'⊙: Processing is NG
    'On Error Resume Next		'☜: Protect system from crashing


	With frm1
		.txtMode.value			= Parent.UID_M0002
		.txtUpdtUserId.value	= Parent.gUsrID
		.txtInsrtUserId.value	= Parent.gUsrID
		
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
		strVal = ""
		strDel = ""
		iColSep = Parent.gColSep
		iRowSep = Parent.gRowSep
    
		'-----------------------
		'Data manipulate area
		'-----------------------
		For lRow = 1 To .vspdData.MaxRows
    
		    .vspdData.Row = lRow
		    .vspdData.Col = 0
			
		    Select Case .vspdData.Text
		    
				Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag							'☜: 입력, 수정 
				
					If .vspdData.Text = ggoSpread.InsertFlag Then
						strVal = strVal & "C" & iColSep & lRow & iColSep				'☜: C=Create
					Else
						strVal = strVal & "U" & iColSep & lRow & iColSep				'☜: U=Update
					End If
					
					.vspdData.Col = C_PAY_PLAN_DT
					If Trim(.vspdData.Text) = "" Then
						strVal = strVal & Trim(.vspdData.Text) & iColSep
					Else 
						strVal = strVal & UNIConvDate(Trim(.vspdData.Text)) & iColSep
					End If
'msgbox Trim(.vspdData.Text)					
					.vspdData.Col = C_PAY_DT
					If Trim(.vspdData.Text) = "" Then
						strVal = strVal & Trim(.vspdData.Text) & iColSep
					Else 
						strVal = strVal & UNIConvDate(Trim(.vspdData.Text)) & iColSep
					End If
					.vspdData.Col = C_PAY_OBJ_CD
					strVal = strVal & Trim(.vspdData.Text) & iColSep		            
					.vspdData.Col = C_PAY_PLAN_AMT
'msgbox "C_PAY_PLAN_AMT - " & CStr(UNICdbl(Trim(.vspdData.Text)))					
					strVal = strVal & UNICdbl(Trim(.vspdData.Text)) & iColSep
					.vspdData.Col = C_PAY_PLAN_LOC_AMT
'msgbox "C_PAY_PLAN_LOC_AMT - " & CStr(UNICdbl(Trim(.vspdData.Text)))
					strVal = strVal & UNICdbl(Trim(.vspdData.Text)) & iColSep		            		            
					.vspdData.Col = C_RESL_FG_CD
					strVal = strVal & Trim(.vspdData.Text) & iColSep		            
					.vspdData.Col = C_DOC_CUR
					strVal = strVal & Trim(.vspdData.Text) & iColSep
					.vspdData.Col = C_XCH_RATE
					strVal = strVal & UNICdbl(Trim(.vspdData.Text)) & iColSep		             	   	            
					.vspdData.Col = C_FLT_CONV_FG
					strVal = strVal & Trim(.vspdData.Text) & iColSep					
					.vspdData.Col = C_LOAN_DESC
					strVal = strVal & Trim(.vspdData.Text) & iColSep
					.vspdData.Col = C_H_PAY_PLAN_DT
					If Trim(.vspdData.Text) = "" Then
						strVal = strVal & Trim(.vspdData.Text) & iColSep
					Else 
						strVal = strVal & UNIConvDate(Trim(.vspdData.Text)) & iColSep
					End If
					strVal = strVal & Trim(.htxtLoanNo.value) & iRowSep					
										            
		            
		            lGrpCnt = lGrpCnt + 1		                  
		          
		        Case ggoSpread.DeleteFlag												'☜: 삭제 

					strVal = strVal & "D" & iColSep & lRow & iColSep					'☜: U=Delete
				    .vspdData.Col = C_PAY_PLAN_DT
					If Trim(.vspdData.Text) = "" Then
						strVal = strVal & Trim(.vspdData.Text) & iColSep
					Else 
						strVal = strVal & UNIConvDate(Trim(.vspdData.Text)) & iColSep
					End If
					.vspdData.Col = C_PAY_OBJ_CD	'3
					strVal = strVal & Trim(.vspdData.Text) & iColSep				
					strVal = strVal & Trim(.htxtLoanNo.value) & iColSep	'4			
					.vspdData.Col = C_H_PAY_PLAN_DT	'5
					strVal = strVal & UNIConvDate(Trim(.vspdData.Text)) & iRowSep			

		            lGrpCnt = lGrpCnt + 1

		    End Select
			            
		Next
	
		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strVal
		
		 Call ExecMyBizASP(frm1, BIZ_PGM_ID)		'☜: 비지니스 ASP 를 가동 
	
	End With

    DbSave = True                           '⊙: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
    ggoSpread.SSDeleteFlag 1 
	
	Call InitVariables
    Call MainQuery()
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete()
	On Error Resume Next
End Function


'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------

'==========================================================
'툴바버튼 세팅 
'==========================================================
Function FncSetToolBar(Cond)
	Select Case UCase(Cond)
	Case "NEW"
		Call SetToolbar("1100110100101111")
	Case "QUERY"
		Call SetToolbar("1100111100111111")
	End Select
End Function

'=======================================================================================================
'   Event Name : SetGridFocus
'   Event Desc :
'=======================================================================================================    
Sub SetGridFocus()	
   
	Frm1.vspdData.Row = 1
	Frm1.vspdData.Col = 1
	Frm1.vspdData.Action = 1
		
End Sub

'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc"  -->	
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag부 
'#########################################################################################################  -->

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSLTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td NOWRAP background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTAB"><font color=white>차입금상환계획변경(KO441)</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
<!--					<TD WIDTH=* ALIGN=RIGHT><A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A></TD> -->
					<TD>&nbsp;</TD>
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
									<TD CLASS="TD5" NOWRAP>차입금번호</TD>
									<TD CLASS="TD6" NOWRAP Colspan=3><INPUT NAME="txtLoanNo" MAXLENGTH="18" SIZE=15  ALT ="차입금번호" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoanNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopupLoan()">
																     <INPUT NAME="txtLoanNm" SIZE=40 STYLE="TEXT-ALIGN: left" ALT ="차입금내역" tag="24"></td>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
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
								<TD CLASS="TD5" NOWRAP>차입일</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtLoanDt" ALT="차입일" SIZE = "10" MAXLENGTH="10" STYLE="TEXT-ALIGN: Center" tag="24X1"></TD>
								<TD CLASS="TD5" NOWRAP>상환만기일</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDueDt" ALT="상환만기일" SIZE = "10" MAXLENGTH="10" STYLE="TEXT-ALIGN: Center" tag="24X1"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>장단기구분</TD>
								<TD CLASS="TD6" NOWRAP><SELECT NAME="cboLoanFg" ALT="장단기구분" STYLE="WIDTH: 135px" tag="24X" OnClick ="vbscript:Type_itemChange()"><OPTION VALUE=""></OPTION></SELECT></TD>
								<TD CLASS="TD5" NOWRAP>차입금계정</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtLoanAcctCd" ALT="차입금계정" SIZE="10" MAXLENGTH="20"  tag="24XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoanAcct" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtLoanAcctCd.value, 8)">
													  <INPUT NAME="txtLoanAcctNm" ALT="차입금계정명" SIZE="20" tag="24X"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>차입용도</TD>												
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtLoanType" ALT="차입용도" SIZE="10" MAXLENGTH="2"  tag="24XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoanType" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtLoanType.value, 6)">
													   <INPUT NAME="txtLoanTypeNm" ALT="차입용도명" STYLE="TEXT-ALIGN: Left" tag="24X"></TD>
								<TD CLASS="TD5" NOWRAP>이자율</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f4223ma1_fpIntRate_txtIntRate.js'></script>&nbsp; %</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>거래통화</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDocCur" ALT="거래통화" SIZE = "10" MAXLENGTH="3"  tag="24XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCurCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtDocCur.value, 0)"></TD>
								<TD CLASS="TD5" NOWRAP>환율</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f4223ma1_fpXRate_txtXchRate.js'></script></TD>
							</TR>
							<TR>									
								<TD CLASS="TD5" NOWRAP>차입금액|자국</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f4223ma1_fpLoanAmt_txtLoanAmt.js'></script>&nbsp;
													   <script language =javascript src='./js/f4223ma1_fpLoanLocAmt_txtLoanLocAmt.js'></script></TD>
								<TD CLASS="TD5" NOWRAP>차입잔액|자국</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f4223ma1_fpLoanBalAmt_txtLoanBalAmt.js'></script>&nbsp;
													   <script language =javascript src='./js/f4223ma1_fpLoanBalLocAmt_txtLoanBalLocAmt.js'></script></TD>
							</TR>							
							<TR>
								<TD CLASS="TD5" NOWRAP>원금상환총액|자국</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f4223ma1_fpPrRdpUnitAmt_txtTotPrRdpAmt.js'></script>&nbsp;	
													   <script language =javascript src='./js/f4223ma1_fpPrRdpUnitLocAmt_txtTotPrRdpLocAmt.js'></script></TD>
								<TD CLASS="TD5" NOWRAP>이자지급총액|자국</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f4223ma1_fpIntPayAmt_txtTotIntPayAmt.js'></script>&nbsp;
													   <script language =javascript src='./js/f4223ma1_fpIntPayLocAmt_txtTotIntPayLocAmt.js'></script></TD>
							</TR>							
							<TR>
								<TD CLASS="TD5" NOWRAP>원금상환예정총액|자국</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f4223ma1_fpPrRdpUnitAmt_txtTotPrPlanAmt.js'></script>&nbsp;
													   <script language =javascript src='./js/f4223ma1_fpPrRdpUnitLocAmt_txtTotPrPlanLocAmt.js'></script></TD>
								<TD CLASS="TD5" NOWRAP>이자지급예정총액|자국</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f4223ma1_fpIntPayPlamAmt_txtTotIntPlanAmt.js'></script>&nbsp;
													   <script language =javascript src='./js/f4223ma1_fpIntPayPlanLocAmt_txtTotIntPlanLocAmt.js'></script></TD>
							</TR>							
							<TR>
								<TD WIDTH="100%" HEIGHT="100%" COLSPAN=4>
									<script language =javascript src='./js/f4223ma1_vaSpread1_vspdData.js'></script>
								</TD>
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
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS=hidden NAME=txtSpread Cols=0 Rows=0 tag="24"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode" tag="24">
<INPUT TYPE=hidden NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=hidden NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=hidden NAME="txtMaxRows" tag="24">
<INPUT TYPE=hidden NAME="htxtLoanNo" tag="24">
<INPUT TYPE=hidden NAME="htxtLoanFgNm" tag="24">
<INPUT TYPE=hidden NAME="htxtLoanType" tag="24">
<INPUT TYPE=hidden NAME="htxtIntPayStnd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
