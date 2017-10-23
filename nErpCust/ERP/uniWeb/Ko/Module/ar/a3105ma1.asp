<%@ LANGUAGE="VBSCRIPT" %>

<!--
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : 채권관리 
'*  3. Program ID           : A3105ma1
'*  4. Program Name         : 입금등록및 채권반제 
'*  5. Program Desc         : 
'*  6. Comproxy List        : +B21011 (Manager)
'                             +B21019 (조회용)
'*  7. Modified date(First) : 2001/02/22
'*  8. Modified date(Last)  : 2003/08/18
'*  9. Modifier (First)     : Chang Sung Hee
'* 10. Modifier (Last)      : Jeong Yong Kyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
 -->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################
'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->

<!--'==========================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="javascript"	SRC="../../inc/TabScript.js">			</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../ag/AcctCtrl.vbs">				</SCRIPT>
<SCRIPT LANGUAGE=vbscript>

Option Explicit																	'☜: indicates that All variables must be declared in advance

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Const BIZ_PGM_QUERY_ID = "a3105mb1.asp"												'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID = "a3105mb2.asp"												'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_DEL_ID =  "a3105mb3.asp"

'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Const COM_BIZ_EXCHRATE_ID = "../../inc/GetExchRate.asp"								'☆: 환율정보 비지니스 로직 ASP명 

Const TAB1 = 1																		'☜: Tab의 위치 
Const TAB2 = 2

Dim C_ArNo 
Dim C_ArDt 
Dim C_ArDueDt 
Dim C_Ar_DocCur
Dim C_ArAmt 
Dim C_ArRemAmt 
Dim C_ArClsAmt 
Dim C_ArClsLocAmt 
Dim C_ArDcAmt 
Dim C_ArDcLocAmt 
Dim C_ArClsDesc 
Dim C_ArAcctCd 
Dim C_AcctNmAr 
Dim C_BizCd 
Dim C_BizNm 

Dim C_ItemSeq 
Dim C_AcctCd 
Dim C_AcctPB 
Dim C_AcctNm 
Dim C_DcAmt 
Dim C_DcLocAmt 

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim  lgStrPrevKey1
Dim  lgStrPrevKeyDtl
Dim  lgStrPrevKey2
Dim  lgStrPrevKey3

Dim  strMode

 '==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= 
 '----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 
Dim  IsOpenPop          
Dim  lgRetFlag
Dim  lgQueryOk
Dim  gSelframeFlg

Dim  lgCurrRow

Dim  dtToday
dtToday = "<%=GetSvrDate%>"

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.1 Common Group -1
' Description : This part declares 1st common function group
'=======================================================================================================
'*******************************************************************************************************



'======================================================================================================
' Name : initSpreadPosVariables()
' Description : 그리드(스프래드) 컬럼 관련 변수 초기화 
'=======================================================================================================
Sub initSpreadPosVariables(ByVal pvSpdNo)
	Select Case UCase(Trim(pvSpdNo))
		Case "A"
			C_ArNo        = 1
			C_ArDt        = 2
			C_ArDueDt     = 3
			C_Ar_DocCur   = 4
			C_ArAmt       = 5 
			C_ArRemAmt    = 6
			C_ArClsAmt    = 7
			C_ArClsLocAmt = 8 
			C_ArDcAmt     = 9
			C_ArDcLocAmt  = 10
			C_ArClsDesc   = 11
			C_ArAcctCd    = 12
			C_AcctNmAr    = 13
			C_BizCd       = 14
			C_BizNm       = 15
		Case "B"
			C_ItemSeq     = 1
			C_AcctCd      = 2
			C_AcctPB      = 3
			C_AcctNm      = 4
			C_DcAmt       = 5
			C_DcLocAmt    = 6
	End Select			
End Sub

'=========================================================================================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub  InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed

    lgIntGrpCount = 0                           'initializes Group View Size
    lgStrPrevKey = ""                            'initializes Previous Key
    lgStrPrevKey1 = ""
    lgStrPrevKeyDtl = 0                         'initializes Previous Key
    lgLngCurRows = 0  
'    gSelframeFlg = TAB1
	frm1.hOrgChangeId.value = parent.gChangeOrgId        
    lgSortKey = 1
    lgQueryOk = False    
End Sub

 '******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'********************************************************************************************************* 
Sub  SetDefaultVal()	
	frm1.txtRcptDt.text  = UniConvDateAToB(dtToday, parent.gServerDateFormat,parent.gDateFormat)
	
	frm1.txtDocCur.value = parent.gCurrency
	frm1.txtDeptCd.value = parent.gDepart
	frm1.txtXchRate.text = 1

    lgBlnFlgChgValue = False
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

'========================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub  InitSpreadSheet(ByVal pvSpdNo)
	Call initSpreadPosVariables(pvSpdNo)
	With frm1
		Select Case UCase(Trim(pvSpdNo))
			Case "A"
				.vspdData1.MaxCols = C_BizNm + 1    
				.vspdData1.Col =.vspdData1.MaxCols
				.vspdData1.ColHidden = True
	
				ggoSpread.Source = .vspdData1
				.vspdData1.Redraw = False	
			
				.vspdData1.MaxRows = 0
				ggoSpread.SpreadInit "V20021128",,parent.gAllowDragDropSpread 

				Call GetSpreadColumnPos(pvSpdNo)

				ggoSpread.SSSetEdit  C_ArNo       , "채권번호"      , 20, 3	'1
				ggoSpread.SSSetDate  C_ArDt       , "채권일자"      , 10, 2, parent.gDateFormat  
				ggoSpread.SSSetDate  C_ArDueDt    , "만기일자"      , 10, 2, parent.gDateFormat  	
				ggoSpread.SSSetEdit	 C_Ar_DocCur  , "거래통화"      , 10, 3   	 									
				ggoSpread.SSSetFloat C_ArAmt      , "채권액"        , 15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_ArRemAmt   , "채권잔액"      , 15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_ArClsAmt   , "반제금액"      , 15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_ArClsLocAmt, "반제금액(자국)", 15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_ArDcAmt    , "할인금액"      , 15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_ArDcLocAmt , "할인금액(자국)", 15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetEdit  C_ArClsDesc  , "비고"          , 20, 3	'7       
				ggoSpread.SSSetEdit  C_ArAcctCd   , "계정코드"      , 20, 3	'2
				ggoSpread.SSSetEdit  C_AcctNmAr   , "계정코드명"    , 20, 3	'3    
				ggoSpread.SSSetEdit  C_BizCd      , "사업장"        , 15, 3	'6
				ggoSpread.SSSetEdit  C_BizNm      , "사업장명"      , 20, 3	'7    
   
				.vspdData1.Redraw = True
			Case "B"
				.vspdData.MaxCols = C_DcLocAmt + 1 												'☜: 최대 Columns의 항상 1개 증가시킴 
				.vspdData.Col = .vspdData.MaxCols													'공통콘트롤 사용 Hidden Column
				.vspdData.ColHidden = True    
    
				ggoSpread.Source = .vspdData
				.vspdData.Redraw = False	    
    
				.vspdData.MaxRows = 0
				Call AppendNumberPlace("6","3","0")

				ggoSpread.SpreadInit "V20021128",,parent.gAllowDragDropSpread 
			
				Call GetSpreadColumnPos(pvSpdNo)

				ggoSpread.SSSetFloat  C_ItemSeq , "NO"            ,  6,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,2,,,"0","999"    
				ggoSpread.SSSetEdit	  C_AcctCd  , "계정코드"      , 20, ,,20, 2
				ggoSpread.SSSetButton C_AcctPB
				ggoSpread.SSSetEdit	  C_AcctNm  , "계정코드명"    , 50,,,20,2
				ggoSpread.SSSetFloat  C_DcAmt   , "할인금액"      , 20, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat  C_DcLocAmt, "할인금액(자국)", 20, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			
				Call ggoSpread.MakePairsColumn(C_AcctCd,C_AcctPB)
			
				.vspdData.Redraw = True
		End Select								
    End With
    
	Call SetSpreadLock(pvSpdNo)
End Sub

'========================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub  SetSpreadLock(ByVal pvSpdNo)
	Dim objSpread
	Dim C_MAX_1 , C_MAX_2

    With frm1    
		Select Case UCase(Trim(pvSpdNo))
			Case "A"
				C_MAX_1 = frm1.vspddata1.MaxCols
				ggoSpread.source = .vspdData1
				    
				.vspdData1.ReDraw = False
				ggoSpread.SpreadLock C_ArNo    ,-1, C_ArNo
				ggoSpread.SpreadLock C_ArAcctCd,-1, C_ArAcctCd
				ggoSpread.SpreadLock C_Ar_DocCur,-1, C_Ar_DocCur
				ggoSpread.SpreadLock C_AcctNmAr,-1, C_AcctNmAr
				ggoSpread.SpreadLock C_BizCd   ,-1, C_BizCd
				ggoSpread.SpreadLock C_BizNm   ,-1, C_BizNm
				ggoSpread.SpreadLock C_ArDt    ,-1, C_ArDt
				ggoSpread.SpreadLock C_ArDueDt ,-1, C_ArDueDt    
				ggoSpread.SpreadLock C_ArAmt   ,-1, C_ArAmt
				ggoSpread.SpreadLock C_ArRemAmt,-1, C_ArRemAmt    
			
				ggoSpread.SSSetRequired  C_ArClsAmt, -1, -1 		
				ggoSpread.SSSetProtected C_MAX_1   , -1, -1 		
						
				.vspdData1.ReDraw = True   
			Case "B"
				C_MAX_2 = frm1.vspddata.MaxCols			
				ggoSpread.Source = .vspdData
				.vspdData.ReDraw = False		
			
				ggoSpread.SpreadLock C_ItemSeq, -1, C_ItemSeq, -1
				ggoSpread.SpreadLock C_AcctCd , -1, C_AcctCd , -1
				ggoSpread.SpreadLock C_AcctPB , -1, C_AcctPB , -1
				ggoSpread.SpreadLock C_AcctNm , -1, C_AcctNm , -1
   
				ggoSpread.SSSetRequired  C_DcAmt, -1, -1 		   
				ggoSpread.SSSetProtected C_MAX_2, -1, -1 						
   				.vspdData.ReDraw = True
		End Select   			
    End With   
End Sub

'========================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub  SetSpreadColor(ByVal pvStartRow , ByVal pvEndRow)
	With frm1.vspdData
		.Redraw = False
		ggoSpread.Source = frm1.vspdData			
		ggoSpread.SSSetProtected C_ItemSeq, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_AcctCd , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_AcctNm , pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_DcAmt  , pvStartRow, pvEndRow   
		.Col = 2											'컬럼의 절대 위치로 이동 
		.Row = .ActiveRow
		.Action = 0                         
		.EditMode = True		
		.Redraw = True		
    End With		
End Sub

'======================================================================================================
' Function Name : SetSpread2Colorar
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub  SetSpread2ColorAr()
	Dim i

    With frm1
		ggoSpread.Source = .vspdData2
		.vspdData2.ReDraw = False	 
	
		For i = 1 To .vspdData2.MaxRows
			ggoSpread.SSSetProtected C_DtlSeq   , i, i
			ggoSpread.SSSetProtected C_CtrlCd   , i, i
			ggoSpread.SSSetProtected C_CtrlNm   , i, i			
			ggoSpread.SSSetProtected C_CtrlValNm, i, i
			.vspdData2.Row = i
			.vspdData2.Col = C_DrFg

			If (.vspdData2.text = "Y")  Or (.vspdData2.text = "DC") Or (.vspdData2.text = "D") Then
				ggoSpread.SSSetRequired C_CtrlVal, i, i	' 
			End If
		Next
		.vspdData2.ReDraw = True
    End With
End Sub

'======================================================================================================
' Function Name : GetSpreadColumnPos
' Function Desc : This method call saved columnorder
'=======================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData1

			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)		

			C_ArNo        = iCurColumnPos(1)
			C_ArDt        = iCurColumnPos(2)
			C_ArDueDt     = iCurColumnPos(3)
			C_Ar_DocCur   = iCurColumnPos(4)
			C_ArAmt       = iCurColumnPos(5) 
			C_ArRemAmt    = iCurColumnPos(6)
			C_ArClsAmt    = iCurColumnPos(7)
			C_ArClsLocAmt = iCurColumnPos(8)
			C_ArDcAmt     = iCurColumnPos(9)			
			C_ArDcLocAmt  = iCurColumnPos(10)
			C_ArClsDesc   = iCurColumnPos(11)
			C_ArAcctCd    = iCurColumnPos(12)
			C_AcctNmAr    = iCurColumnPos(13)
			C_BizCd       = iCurColumnPos(14)
			C_BizNm       = iCurColumnPos(15)		
		Case "B"
			ggoSpread.Source = frm1.vspdData

			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)		
						
			C_ItemSeq     = iCurColumnPos(1)
			C_AcctCd      = iCurColumnPos(2)
			C_AcctPB      = iCurColumnPos(3)
			C_AcctNm      = iCurColumnPos(4)
			C_DcAmt       = iCurColumnPos(5)
			C_DcLocAmt    = iCurColumnPos(6)
	End select
End Sub

'======================================================================================================
' Function Name : OpenPopupGL
' Function Desc : This method Open The Popup window for GL
'=======================================================================================================
Function OpenPopupGL()
	Dim arrRet
	Dim arrParam(1)	
	Dim iCalledAspName
	iCalledAspName = AskPRAspName("a5120ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function
	
	arrParam(0) = Trim(frm1.txtGlNo.value)							'회계전표번호 
	arrParam(1) = ""												'Reference번호 

	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function

'======================================================================================================
' Function Name : OpenPopupTempGL
' Function Desc : This method Open The Popup window for TempGL
'=======================================================================================================
Function OpenPopuptempGL()
	Dim arrRet
	Dim arrParam(1)	
	Dim iCalledAspName
	
	iCalledAspName = AskPRAspName("a5130ra1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function
	
	arrParam(0) = Trim(frm1.txtTempGlNo.value)						'회계전표번호 
	arrParam(1) = ""												'Reference번호 

	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function
'------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function

	If frm1.txtBpCd.className = "protected" Then Exit Function

	IsOpenPop = True

	arrParam(0) = strCode									'Code Condition
   	arrParam(1) = ""										'채권과 연계(거래처 유무)
	arrParam(2) = ""										'FrDt
	arrParam(3) = ""										'ToDt
	arrParam(4) = "B"										'B :매출 S: 매입 T: 전체 
	arrParam(5) = "PAYER"									'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtBpCd.focus
		Exit Function
	Else
		frm1.txtBpCd.value=arrRet(0)
		frm1.txtBpNm.value= arrRet(1)
		frm1.txtBpCd.focus
	End If	
End Function

'======================================================================================================
'   Function Name : OpenPopUp()
'   Function Desc : 
'=======================================================================================================
Function  OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iArrParam(8)
	Dim strCd
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function

	Select Case iWhere
		Case 0

		Case 1
			If frm1.txtBpCd.className = "protected" Then Exit Function
			
			arrParam(0) = "거래처팝업"
			arrParam(1) = "B_BIZ_PARTNER"				
			arrParam(2) = Trim(frm1.txtBpCd.Value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "거래처"
	
			arrField(0) = "BP_CD"
			arrField(1) = "BP_NM"
    
			arrHeader(0) = "거래처"
			arrHeader(1) = "거래처명"									' Header명(1)
		Case 3		
			If frm1.txtDocCur.className = "protected" Then Exit Function
			
			arrParam(0) = "거래통화팝업"								' 팝업 명칭 
			arrParam(1) = "B_CURRENCY"										' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtDocCur.Value)						' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = ""												' Where Condition
			arrParam(5) = "거래통화"			
	
			arrField(0) = "CURRENCY"										' Field명(0)
			arrField(1) = "CURRENCY_DESC"									' Field명(1)
    
			arrHeader(0) = "거래통화"									' Header명(0)
			arrHeader(1) = "거래통화명"
		Case 4
			arrParam(0) = "계정코드팝업"								' 팝업 명칭 
			arrParam(1) = "A_Acct, A_ACCT_GP" 								' TABLE 명칭 
			arrParam(2) = Trim(strCode)										' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = "A_ACCT.GP_CD=A_ACCT_GP.GP_CD AND A_ACCT.DEL_FG <> " & FilterVar("Y", "''", "S") & " "												' Where Condition
			arrParam(5) = "계정코드"									' 조건필드의 라벨 명칭 

			arrField(0) = "A_ACCT.Acct_CD"									' Field명(0)
			arrField(1) = "A_ACCT.Acct_NM"									' Field명(1)
    		arrField(2) = "A_ACCT_GP.GP_CD"									' Field명(2)
			arrField(3) = "A_ACCT_GP.GP_NM"									' Field명(3)
			
			arrHeader(0) = "계정코드"									' Header명(0)
			arrHeader(1) = "계정코드명"									' Header명(1)
			arrHeader(2) = "그룹코드"									' Header명(2)
			arrHeader(3) = "그룹명"										' Header명(3)
		Case 5	
			If frm1.txtBankCd.className = "protected" Then Exit Function
		
			arrParam(0) = "은행팝업"
			arrParam(1) = "F_DPST, B_BANK"				
			arrParam(2) = Trim(frm1.txtBankCd.Value)
			arrParam(3) = ""
			arrParam(4) = "F_DPST.BANK_CD = B_BANK.BANK_CD"
			arrParam(5) = "은행"			
	
			arrField(0) = "F_DPST.BANK_CD"	
			arrField(1) = "B_BANK.BANK_NM"	
    
			arrHeader(0) = "은행"		
			arrHeader(1) = "은행명"	
		Case 6
			If frm1.txtBankAcct.className = "protected" Then Exit Function
			
			arrParam(0) = "계좌번호팝업"
			arrParam(1) = "F_DPST, B_BANK"				
			arrParam(2) = Trim(frm1.txtBankAcct.Value)
			arrParam(3) = ""
			
			If Trim(frm1.txtBankCd.Value) = "" Then
				strCd = "F_DPST.BANK_CD = B_BANK.BANK_CD "
			Else
				strCd = "F_DPST.BANK_CD = B_BANK.BANK_CD AND  F_DPST.BANK_CD =  " & FilterVar(frm1.txtBankCd.Value, "''", "S") & " "	
			End If		
			
			arrParam(4) = strCd
			arrParam(5) = "계좌번호"			
			
		    arrField(0) = "F_DPST.BANK_ACCT_NO"	
		    arrField(1) = "F_DPST.BANK_CD"	
		    arrField(2) = "B_BANK.BANK_NM"	
		    
		    arrHeader(0) = "계좌번호"		
		    arrHeader(1) = "은행"	
		    arrHeader(2) = "은행명"				
		Case 7
			If frm1.txtCheckCd.className = "protected" Then Exit Function
			Dim strWhere 
			
			strWhere = "MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND B_CONFIGURATION.SEQ_NO = 3 AND  B_CONFIGURATION.REFERENCE = " & FilterVar("PR", "''", "S") & "  "
			strWhere = strWhere & "AND  MINOR_CD= " & FilterVar(UCase(frm1.txtInputType.value), "''", "S") & ""

			If CommonQueryRs( "MINOR_CD" , "B_CONFIGURATION" , strWhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then 
				Select Case UCase(Trim(Left(lgF0, Len(lgF0)-1)))
					Case "NR"
						arrParam(0) = "어음번호팝업"													' 팝업 명칭 
						arrParam(1) = "f_note a,b_biz_partner b, b_bank c"									' TABLE 명칭 
						arrParam(2) = Trim(frm1.txtCheckCd.Value)											' Code Condition
						arrParam(3) = ""		
						arrParam(4) = "a.note_sts = " & FilterVar("BG", "''", "S") & "  AND a.note_fg = " & FilterVar("D1", "''", "S") & "  And a.bp_cd = b.bp_cd And a.bank_cd *= c.bank_cd"					' Where Condition										
						arrParam(5) = "어음번호"														' 조건필드의 라벨 명칭 
							
						arrHeader(0) = "어음번호"														' Header명(0)
						arrHeader(1) = "금액"															' Header명(1)
						arrHeader(2) = "발행일"															' Header명(1)     
						arrHeader(3) = "거래처"															' Header명(1)
						arrHeader(4) = "은행"															' Header명(1)						
							
						arrField(0) = "Note_no"																' Field명(0)
						arrField(1) =  "F2" & parent.gColSep & "a.Note_amt"									' Field명(1)
						arrField(2) =  "DD" & parent.gColSep & "a.Issue_dt"									' Field명(2)
						arrField(3) = "b.bp_nm"
						arrField(4) = "c.bank_nm"         						
					Case "CR"
						arrParam(0) = "수취구매카드 팝업"											' 팝업 명칭 
						arrParam(1) = "f_note a,b_biz_partner b, b_bank c , b_card_co d "					' TABLE 명칭 
						arrParam(2) = Trim(frm1.txtCheckCd.Value)											' Code Condition
						arrParam(3) = ""						
						arrParam(4) = "a.note_sts = " & FilterVar("BG", "''", "S") & "  AND a.note_fg = " & FilterVar("CR", "''", "S") & "  And a.bp_cd = b.bp_cd And a.bank_cd *= c.bank_cd and a.card_co_cd *= d.card_co_cd "		' Where Condition   						
						arrParam(5) = "수취구매카드번호"												' 조건필드의 라벨 명칭						               
							
						arrHeader(0) = "수취구매카드번호"												' Header명(0)
						arrHeader(1) = "금액"															' Header명(1)
						arrHeader(2) = "발행일"															' Header명(1)     
						arrHeader(3) = "거래처"															' Header명(1)
						arrHeader(4) = "카드사"															' Header명(1)						
							
						arrField(0) = "Note_no"																' Field명(0)
						arrField(1) =  "F2" & parent.gColSep & "a.Note_amt"									' Field명(1)
						arrField(2) =  "DD" & parent.gColSep & "a.Issue_dt"									' Field명(2)
						arrField(3) = "b.bp_nm"
						arrField(4) = "d.card_co_nm"         						
					Case Else
						Call DisplayMsgBox("141327", "X", "X", "X")
						Exit Function
				End Select		
			ENd if					
		Case 8   
			If frm1.txtInputType.className = "protected" Then Exit Function

			If frm1.txtDocCur.value <> "" Then
				If UCase(Trim(frm1.txtDocCur.value)) = parent.gCurrency Then
					arrParam(0) = "입금유형"														' 팝업 명칭						
					arrParam(1) = "B_MINOR A,B_CONFIGURATION B"
					arrParam(2) = Trim(frm1.txtInputType.value)											' Code Condition
					arrParam(3) = ""																	' Name Cindition
					arrParam(4) = "A.MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND A.MAJOR_CD = B.MAJOR_CD AND A.MINOR_CD = B.MINOR_CD   " _
					   			  & "AND B.SEQ_NO = 3 AND B.REFERENCE = " & FilterVar("PR", "''", "S") & " "							' Where Condition
					arrParam(5) = "입금유형"														' TextBox 명칭 
		
					arrField(0) = "A.MINOR_CD"													' Field명(0)
					arrField(1) = "A.MINOR_NM"													' Field명(1)
					arrField(2) = "B.REFERENCE"															' Field명(1)					
	    
					arrHeader(0) = "입금유형"														' Header명(0)
					arrHeader(1) = "입금유형명"														' Header명(1)		
				Else
					arrParam(0) = "입금유형"														' 팝업 명칭						
					arrParam(1) = "B_MINOR A,B_CONFIGURATION B "
					arrParam(2) = Trim(frm1.txtInputType.value)											' Code Condition
					arrParam(3) = ""																	' Name Cindition
					arrParam(4) = "A.MINOR_CD = B.MINOR_CD and A.MAJOR_CD = " & FilterVar("A1006", "''", "S") & " " _
								& " and B.SEQ_NO = 3 and B.REFERENCE = " & FilterVar("PR", "''", "S") & " " _
								& " And A.minor_cd Not in ( Select  minor_cd  from b_configuration " _ 
								& " where major_cd=" & FilterVar("a1006", "''", "S") & "  and seq_no=4 and reference=" & FilterVar("NO", "''", "S") & " ) "			' Where Condition								
					arrParam(5) = "입금유형"														' TextBox 명칭 
		
					arrField(0) = "A.MINOR_CD"													' Field명(0)
					arrField(1) = "A.MINOR_NM"													' Field명(1)
					arrField(2) = "B.REFERENCE"															' Field명(1)					
	    
					arrHeader(0) = "입금유형"														' Header명(0)
					arrHeader(1) = "입금유형명"														' Header명(1)		
				End If
			Else
				arrParam(0) = "입금유형"															' 팝업 명칭						
				arrParam(1) = "B_MINOR A,B_CONFIGURATION B"
				arrParam(2) = Trim(frm1.txtInputType.value)												' Code Condition
				arrParam(3) = ""																		' Name Cindition
				arrParam(4) = "A.MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND A.MAJOR_CD = B.MAJOR_CD AND A.MINOR_CD = B.MINOR_CD   " _
				   			  & "AND B.SEQ_NO = 3 AND B.REFERENCE = " & FilterVar("PR", "''", "S") & " "								' Where Condition					
				arrParam(5) = "입금유형"															' TextBox 명칭 
		
				arrField(0) = "A.MINOR_CD"														' Field명(0)
				arrField(1) = "A.MINOR_NM"														' Field명(1)
				arrField(2) = "B.REFERENCE"																' Field명(1)				
	    
				arrHeader(0) = "입금유형"															' Header명(0)
				arrHeader(1) = "입금유형명"															' Header명(1)									
			End If				
		Case 9	'입금계정코드 
			arrParam(0) = "계정코드팝업"								' 팝업 명칭 
			arrParam(1) = "A_Acct A, A_ACCT_GP B, A_JNL_ACCT_ASSN C"				' TABLE 명칭 
			arrParam(2) = ""												' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = "A.GP_CD=B.GP_CD AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND A.Acct_cd=C.Acct_CD" & _
							" and C.trans_type = " & FilterVar("ar002", "''", "S") & "  and C.jnl_cd =  " & FilterVar(frm1.txtInputType.Value, "''", "S") & "  "	' Where Condition
			arrParam(5) = "계정코드"									' 조건필드의 라벨 명칭 

			arrField(0) = "A.Acct_CD"									' Field명(0)
			arrField(1) = "A.Acct_NM"									' Field명(1)
    		arrField(2) = "B.GP_CD"										' Field명(2)
			arrField(3) = "B.GP_NM"										' Field명(3)
			
			arrHeader(0) = "계정코드"									' Header명(0)
			arrHeader(1) = "계정코드명"									' Header명(1)
			arrHeader(2) = "그룹코드"									' Header명(2)
			arrHeader(3) = "그룹명"										' Header명(3)						
		Case Else
			Exit Function
	End Select				

	IsOpenPop = True

	If iwhere = 0 Then	
		iCalledAspName = AskPRAspName("a3105ra1")

		' 권한관리 추가 
		iarrParam(5) = lgAuthBizAreaCd
		iarrParam(6) = lgInternalCd
		iarrParam(7) = lgSubInternalCd
		iarrParam(8) = lgAuthUsrID

		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a3105ra1", "X")
			IsOpenPop = False
			Exit Function
		End If
					
		arrRet = window.showModalDialog(iCalledAspName,Array(window.parent,iarrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Else
		arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")			
	End If
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then	    
		Call EscPopup(iWhere)
		Exit Function
	Else
		Call SetPopup(arrRet, iWhere)
	End If
End Function
'======================================================================================================
'   Function Name : EscPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function EscPopup(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0		
				.txtRcptNo.focus				
			Case 1	
				.txtBpCd.focus	
			Case 3
				.txtDocCur.focus	
			Case 4
				Call SetActiveCell(frm1.vspdData,C_AcctCd,frm1.vspdData.ActiveRow ,"M","X","X")
			Case 5
				.txtBankCd.focus			    		
			Case 6
				.txtBankAcct.focus		
			Case 7	
				.txtCheckCd.focus		
			Case 8
				.txtInputType.focus	
			Case 9
				.txtAcctCd.focus					
		End Select				
	End With
	
End Function

'======================================================================================================
'   Function Name : SetPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function SetPopup(Byval arrRet,Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0		
				.txtRcptNo.value = arrRet(0)
				.txtRcptNo.focus				
			Case 1	
				.txtBpCd.value = arrRet(0)		
				.txtBpNm.value = arrRet(1)
				.txtBpCd.focus	
			Case 3
				.txtDocCur.value = arrRet(0)		
				Call txtDocCur_OnChange()
				.txtDocCur.focus	
			Case 4
				.vspdData.Col = C_AcctCd
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_AcctNm
				.vspdData.Text = arrRet(1)
			
				Call vspdData_Change(C_AcctCd, frm1.vspdData.activerow )	 ' 변경이 일어났다고 알려줌 
				Call SetActiveCell(frm1.vspdData,C_AcctCd,frm1.vspdData.ActiveRow ,"M","X","X")
			Case 5
				.txtBankCd.value = arrRet(0)		
				.txtBankNm.value = arrRet(1)	
				.txtBankCd.focus			    		
			Case 6
				.txtBankAcct.value = arrRet(0)		
				.txtBankCd.value = arrRet(1)		
				.txtBankNm.value = arrRet(2)
				.txtBankAcct.focus		
			Case 7	
				.txtCheckCd.value = arrRet(0)	
				.txtCheckCd.focus		
			Case 8
				.txtInputType.value = arrRet(0)		
				.txtInputTypeNm.value = arrRet(1)				
				Call txtInputType_OnChange()
				.txtInputType.focus	
			Case 9
				.txtAcctCd.value = arrRet(0)		
				.txtAcctnm.value = arrRet(1)
				.txtAcctCd.focus					
		End Select				
	End With
	If iwhere  <> 0 Then
		lgBlnFlgChgValue = True
	End If	
End Function

'------------------------------------------  OpenDept()  ---------------------------------------
'	Name : OpenDept()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenDept(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(8)

	If IsOpenPop = True Then Exit Function
	If frm1.txtDeptCd.className = "protected" Then Exit Function
			
	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = frm1.txtRcptDt.Text
	arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  
	arrParam(3) = "F"									' 결의일자 상태 Condition
	
	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID	

	arrRet = window.showModalDialog("../../comasp/DeptPopupDtA2.asp", Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtDeptCd.focus
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
				.txtDeptCd.value = arrRet(0)
				.txtDeptNm.value = arrRet(1)
				.txtRcptDt.text = arrRet(3)
				call txtDeptCd_OnBlur()  
				frm1.txtDeptCd.focus
	    End Select
	End With
End Function 

'============================================================================================
'	Name : OpenRefOpenAr()
'	Description : Ref 화면을 call한다. 
'============================================================================================
Function OpenRefOpenAr()
	Dim arrRet
	Dim arrParam(11)	
	Dim iCalledAspName
	
	iCalledAspName = AskPRAspName("a3106ra6")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a3106ra6", "X")
		IsOpenPop = False
		Exit Function
	End If

	If gSelframeFlg <> TAB1 Then Exit Function		 		
	If IsOpenPop = True Then Exit Function
   
	IsOpenPop = True

	If frm1.vspdData1.MaxRows = 0 Then frm1.hArDocCur.value	= ""

	arrParam(0) = frm1.txtBpCd.value				' 검색조건이 있을경우 파라미터 
	arrParam(1) = frm1.txtBpNm.value			
	arrParam(2) = frm1.hArDocCur.value
    arrParam(3) = "M"
    arrParam(6) = frm1.txtRcptDt.text
    arrParam(7) = frm1.txtRcptDt.Alt
    
    ' 권한관리 추가 
	arrParam(8) = lgAuthBizAreaCd
	arrParam(9) = lgInternalCd
	arrParam(10) = lgSubInternalCd
	arrParam(11) = lgAuthUsrID
        
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0, 0) = "" Then		
		Exit Function
	Else		
		Call SetRefOpenAr(arrRet)
	End If
End Function

'=========================================================================================================
'	Name : SetRefOpenAr()
'	Description : OpenAp Popup에서 Return되는 값 setting
'=========================================================================================================
Function SetRefOpenAr(Byval arrRet)
	Dim intRtnCnt, strData
	Dim TempRow, I
	Dim j	
	DIM X
	Dim sFindFg
	
	With frm1
		.vspdData1.focus		
		ggoSpread.Source = .vspdData1
		.vspdData1.ReDraw = False	
	
		TempRow = .vspdData1.MaxRows												'☜: 현재까지의 MaxRows

		For I = TempRow To TempRow + Ubound(arrRet, 1) 
			sFindFg	= "N"
			For x = 1 To TempRow
				.vspdData1.Row = x
				.vspdData1.Col = C_ArNo				
				If "" & UCase(Trim(.vspdData1.Text)) = "" & UCase(Trim(arrRet(I - TempRow, 0))) Then
					sFindFg	= "Y"
				End If
			Next
			If 	sFindFg	= "N" Then
				.vspdData1.MaxRows = .vspdData1.MaxRows + 1
				.vspdData1.Row = I + 1				
				.vspdData1.Col = 0
				.vspdData1.Text = ggoSpread.InsertFlag

				.vspdData1.Col = C_ArNo        												
				.vspdData1.text = arrRet(I - TempRow, 0)
				.vspdData1.Col = C_ArDt       												
				.vspdData1.text = arrRet(I - TempRow, 1)
				.vspdData1.Col = C_ArDueDt         											
				.vspdData1.text = arrRet(I - TempRow, 2)
				.vspdData1.Col = C_Ar_DocCur
				.vspdData1.text = arrRet(I - TempRow, 14)				
				.vspdData1.Col = C_ArAmt        											
				.vspdData1.text = arrRet(I - TempRow, 3)
				.vspdData1.Col = C_ArRemAmt         										
				.vspdData1.text = arrRet(I - TempRow, 4)
				.vspdData1.Col = C_ArClsAmt
				.vspdData1.text = arrRet(I - TempRow, 6)
				.vspdData1.Col = C_ArAcctCd
				.vspdData1.text = arrRet(I - TempRow, 7)
				.vspdData1.Col = C_AcctNmAr
				.vspdData1.text = arrRet(I - TempRow, 8)
				.vspdData1.Col = C_BizCd
				.vspdData1.text = arrRet(I - TempRow, 9)
				.vspdData1.Col = C_BizNm
				.vspdData1.text = arrRet(I - TempRow, 10)
				.vspdData1.Col = C_ArClsDesc
				.vspdData1.text = arrRet(I - TempRow, 13)
			End If	
		Next	
	
		If Trim(.txtBpCd.Value) = "" Then
			.txtbpCd.Value = arrRet(0, 11)
			.txtbpNm.Value = arrRet(0, 12)		
		End If
		
		If Trim(.txtBpCd.value) <> "" Then					
			Call ggoOper.SetReqAttr(frm1.txtBpCd,   "Q")		
		Else			
			Call ggoOper.SetReqAttr(frm1.txtBpCd,   "N")		
		End If
	
		.hArDocCur.Value = arrRet(0, 14)
		.txtDocCur.value = arrRet(0, 14) '20051201 추가 
		If .txtDocCur.Value <> "" Then			
			If UCase(Trim(.txtDocCur.Value)) = UCase(Trim(arrRet(0, 14))) Then
			Else
				.txtRcptAmt.Text	= "0"
				.txtRcptLocAmt.Text = "0"
			End If
		End If

		Call CurFormatNumSprSheet()		
		Call ggoOper.SetReqAttr(frm1.txtRcptDt,   "Q")	
		Call DoSum()
		Call SetSpreadLock("A")
		
		.vspdData1.ReDraw = True
    End With
End Function

'======================================================================================================
'	기능: Tab Click
'	설명: Tab Click시 필요한 기능을 수행한다.
'=======================================================================================================
Function ClickTab1()
	If lgIntFlgMode <> parent.OPMD_UMODE Then
	    Call SetToolBar("1110100100001111")										'⊙: 버튼 툴바 제어 
	Else                 
	    Call SetToolBar("1111101100001111")										'⊙: 버튼 툴바 제어 
	End If
	
	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)														'~~~ 첫번째 Tab 
	gSelframeFlg = TAB1
End Function

Function ClickTab2()
	If gSelframeFlg = TAB2 Then Exit Function
	Call changeTabs(TAB2)														'~~~ 두번째 Tab 
	gSelframeFlg = TAB2
	
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetToolBar("1110111100001111")
	Else                 
		Call SetToolBar("1111111100001111")
	End If	
End Function

'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.2 Common Group-2
' Description : This part declares 2nd common function group
'=======================================================================================================
'*******************************************************************************************************



'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub  Form_Load()
    Call LoadInfTB19029()  
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, _
							parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)    
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, _
							parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)	'⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")										'Lock  Suitable  Field
    Call InitSpreadSheet("A")	    															'Setup the Spread sheet
    Call InitSpreadSheet("B")	    															'Setup the Spread sheet    
	Call InitCtrlSpread()
	Call InitCtrlHSpread()

    Call txtInputType_onChange()
    Call InitVariables()																	'Initializes local global variables
    Call SetDefaultVal()
	Call ClickTab1()
    frm1.txtRcptNo.focus
	gIsTab     = "Y" 
	gTabMaxCnt = 2
	
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

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function  FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False																'⊙: Processing is NG
    
    Err.Clear																		'☜: Protect system from crashing
	'-----------------------
    'Check condition area
    '-----------------------    
    If Not chkField(Document, "1") Then												'⊙: This function check indispensable field
       Exit Function
    End If
    
    '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")			'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field
	ggoSpread.Source = frm1.vspdData	:	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData1	:	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2	:	ggoSpread.ClearSpreadData		
	ggoSpread.Source = frm1.vspdData3	:	ggoSpread.ClearSpreadData
	
    Call ClickTab1()
    Call InitVariables()

    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery()																	'☜: Query db data
       
    FncQuery = True																	'⊙: Processing is OK
		
	Set gActiveElement = document.activeElement    
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function  FncNew() 
	Dim IntRetCD 
	Dim var1, var2, var3
	    
    FncNew = False                                                          
    
    ggoSpread.Source = frm1.vspdData1
    var1 = ggoSpread.SSCheckChange
	ggoSpread.Source = frm1.vspdData
    var2 = ggoSpread.SSCheckChange
    
	'-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Or var1 = True Or var2 = True Or var3 = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"X","X")               
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")                                       '⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                       '⊙: Clear Contents  Field
    Call ggoOper.LockField(Document, "N")                                        '⊙: Lock  Suitable  Field
    
    Call txtDocCur_OnChange()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData1
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData		
	ggoSpread.Source = frm1.vspdData3
	ggoSpread.ClearSpreadData
    
    Call txtInputType_onChange()
    Call InitVariables()																	'⊙: Initializes local global variables
    Call ClickTab1()
    Call SetDefaultVal()
    Call txtDocCur_OnChange()
	lgBlnFlgChgValue = False
			
    FncNew = True																			'⊙: Processing is OK
		
	Set gActiveElement = document.activeElement    
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function  FncDelete() 
    Dim IntRetCD 
    
    FncDelete = False																		'⊙: Processing is NG
    
    On Error Resume Next																	'☜: Protect system from crashing    
    Err.Clear
    
    '-----------------------
    'Precheck area
    '-----------------------    
    If lgIntFlgMode <> parent.OPMD_UMODE Then												'Check if there is retrived data
        IntRetCD = DisplayMsgBox("900002","X","X","X")                                       
        Exit Function
    End If
        
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")							'Will you destory previous data"
    
	If IntRetCD = vbNo Then
		Exit Function
	End If
	
    '-----------------------
    'Delete function call area
    '-----------------------    
    If DbDelete = False Then																'☜: Delete db data
		Exit Function               
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------    
    FncDelete = True																		'⊙: Processing is OK
		
	Set gActiveElement = document.activeElement    
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function  FncSave() 
	Dim IntRetCD 
    Dim var1,var2, var3
	
    FncSave = False                                                         
    
    Err.Clear                                                               
    
    ggoSpread.Source = frm1.vspdData1
    var1 = ggoSpread.SSCheckChange

    ggoSpread.Source = frm1.vspdData
    var2 = ggoSpread.SSCheckChange

    ggoSpread.Source = frm1.vspdData2
    var3 = ggoSpread.SSCheckChange

    If lgBlnFlgChgValue = False And var1 = False And var2 = False And var3 = False Then		'⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","X","X","X")										'⊙: Display Message(There is no changed data.)
		Exit Function
    End If

	If Not chkField(Document, "2") Then														'⊙: Check required field(Single area)
		Exit Function
    End If    

	ggoSpread.Source = frm1.vspdData1
    If Not ggoSpread.SSDefaultCheck Then
		Call ClickTab1()																	'⊙: Check contents area
		Exit Function
    End If
    
    ggoSpread.Source = frm1.vspddata
    If Not ggoSpread.SSDefaultCheck Then													'⊙: Check contents area
		Call ClickTab2()
		Exit Function
    End If
    
    If Not chkAllcDate Then
		Exit Function
    End If

    If chkInputType= False Then
		Exit Function
    End If

    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave()																			'☜: Save db data

    FncSave = True																			'⊙: Processing is OK
    		
	Set gActiveElement = document.activeElement    
End Function

'=======================================================================================================
' Function Name : chkInputType
' Function Desc : 
'========================================================================================================
Function chkInputType()
	Dim intI
	Dim IntRetCD
	
	chkInputType = True

	If CommonQueryRs("REFERENCE" , "B_CONFIGURATION " , "MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND MINOR_CD = " & FilterVar(frm1.txtInputType.value, "''", "S") & " AND SEQ_NO = 4 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then 
		Select Case UCase(Trim(Left(lgF0, Len(lgF0)-1)))
			Case "NO" 	
				If UCase(Trim(frm1.txtDocCur.value)) <> UCase(parent.gCurrency) Then		
					IntRetCD = DisplayMsgBox("111620","X","X","X")
					frm1.txtInputType.value = ""
					frm1.txtInputTypeNm.value = ""					
					frm1.txtAcctCd.value = ""
					frm1.txtAcctNm.value = ""										
					frm1.txtInputType.focus
					Call ggoOper.SetReqAttr(frm1.txtBankCd, "Q")				
					Call ggoOper.SetReqAttr(frm1.txtBankAcct, "Q")

					chkInputType = False
				End If					
			Case Else
		End Select
	End If	
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function  FncCopy() 
	If frm1.vspdData1.MaxRows < 1 Then Exit Function 
	
	frm1.vspdData1.ReDraw = False
	
    ggoSpread.Source = frm1.vspdData1	
    ggoSpread.CopyRow
    Call SetSpreadColor(frm1.vspdData1.ActiveRow, frm1.vspdData1.ActiveRow)
    
	frm1.vspdData1.ReDraw = True
		
	Set gActiveElement = document.activeElement    
	
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function  FncCancel() 
   Dim i
   
   If gSelframeFlg = TAB1 Then
		If frm1.vspdData1.MaxRows < 1 Then Exit Function
		
		With frm1.vspdData1
		    .Row = .ActiveRow
		    .Col = 0
		    
		    ggoSpread.Source = frm1.vspdData1
		    ggoSpread.EditUndo
			Call DoSum()
			If frm1.vspdData1.MaxRows < 1 Then 
				Call ggoOper.SetReqAttr(frm1.txtRcptDt,   "N")
				Exit Function
			End if					
		    .Row = .ActiveRow
		    .Col = 0		    
			
			For i = .MaxRows to 0 Step -1 
				.Row= i
				.Col =0			
				If Trim(frm1.vspddata1.text) = ggoSpread.InsertFlag Then 
					Call ggoOper.SetReqAttr(frm1.txtRcptDt,   "Q")
					Exit Function
				End if
				
				Call ggoOper.SetReqAttr(frm1.txtRcptDt,   "N")
			Next
		
		End With   
	Else
		If frm1.vspdData.MaxRows < 1 Then Exit Function
		
		With frm1.vspdData
		    .Row = .ActiveRow
		    .Col = 0
		    If .Text = ggoSpread.InsertFlag Then
				.Col = C_AcctCd
				If Len(Trim(.text)) > 0 Then 
					.Col = C_ItemSeq
					DeleteHSheet(.Text)
				End If		
		    End If
   
		    ggoSpread.Source = frm1.vspdData	
		    ggoSpread.EditUndo
			Call DoSum()
			If frm1.vspdData.MaxRows < 1 Then Exit Function
			
			.Row = .ActiveRow
			.Col = 0		    
			
			If .Row = 0 Then Exit Function
			
		    If .Text = ggoSpread.InsertFlag Then
				.Col = C_AcctCd
				If Len(Trim(.text)) > 0 Then 
					.Col = C_ItemSeq
					frm1.hItemSeq.value = .Text
					ggoSpread.Source = frm1.vspdData2
					ggoSpread.ClearSpreadData
					Call DbQuery3(.ActiveRow)
				End If
		    Else
		        .Col = C_ItemSeq
		        frm1.hItemSeq.value = .Text
				ggoSpread.Source = frm1.vspdData2
				ggoSpread.ClearSpreadData
		        Call DbQuery2(.ActiveRow)
		    End If
		End With
	End If    
		
	Set gActiveElement = document.activeElement    
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function  FncInsertRow(ByVal pvRowCnt) 
	Dim imRow
	Dim ii
	Dim iCurRowPos

    If gSelframeFlg <> TAB2 Then
		Call ClickTab2()															'sstData.Tab = 1
    End If
    
	FncInsertRow = False															'☜: Processing is NG	   

	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = Cint(pvRowCnt)
	Else
	    imRow = AskSpdSheetAddRowCount()
    
		If imRow = "" Then
		    Exit Function
		End If
	End If		
        
    With frm1.vspdData
		iCurRowPos = .ActiveRow	
		.ReDraw = False		    
		ggoSpread.Source = frm1.vspdData
		
		ggoSpread.InsertRow ,imRow

		For ii = .ActiveRow To  .ActiveRow + imRow - 1
			Call MaxSpreadVal(frm1.vspdData, C_ItemSeq , ii)
		Next        

		.Col = 1																	' 컬럼의 절대 위치로 이동 
		.Row = ii - 1
		.Action = 0		
		.ReDraw = True

		Call SetSpreadColor(iCurRowPos + 1, iCurRowPos + imRow)
	End With
	
    If Err.number = 0 Then
		FncInsertRow = True                                                          '☜: Processing is OK
    End If  		

	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData
			
	Set gActiveElement = document.activeElement    
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function  FncDeleteRow() 
    Dim lDelRows
	
	If gSelframeFlg = TAB1 Then
		if frm1.vspdData1.MaxRows < 1 Then Exit Function
		ggoSpread.Source = frm1.vspdData1
	else
		if frm1.vspdData.MaxRows < 1 Then Exit Function
		ggoSpread.Source = frm1.vspdData
	end if	
	
    lDelRows = ggoSpread.DeleteRow
    Call DoSum()
    		
	Set gActiveElement = document.activeElement    
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function  FncPrint() 
    On Error Resume Next                                                    '☜: Protect system from crashing
    parent.FncPrint()
		
	Set gActiveElement = document.activeElement    
End Function

'=======================================================================================================
' Function Name : FncFind
' Function Desc : 화면 속성, Tab유무 
'========================================================================================================
Function  FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI , True)                          
		
	Set gActiveElement = document.activeElement    
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function  FncExcel() 
	Call FncExport(parent.C_SINGLEMULTI)
		
	Set gActiveElement = document.activeElement    
	
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
		
	Set gActiveElement = document.activeElement    
End Sub

'=======================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function  FncExit()
	Dim IntRetCD
	Dim var1,var2, var3
	
	FncExit = False

	ggoSpread.Source = frm1.vspdData1
    var1 = ggoSpread.SSCheckChange

    ggoSpread.Source = frm1.vspdData
    var2 = ggoSpread.SSCheckChange

    ggoSpread.Source = frm1.vspdData2
    var3 = ggoSpread.SSCheckChange
    
    If lgBlnFlgChgValue = True or var1 = True or var2 = True or var3 = True Then  '⊙: Check If data is chaged
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")                '데이타가 변경되었습니다. 종료 하시겠습니까?
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
    FncExit = True
    		
	Set gActiveElement = document.activeElement    
End Function

'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.3 Common Group - 3
' Description : This part declares 3rd common function group
'=======================================================================================================
'*******************************************************************************************************



'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function  DbQuery() 
	Dim strVal

    Call LayerShowHide(1)
    
    DbQuery = False
    
    Err.Clear																	'☜: Protect system from crashing

    With frm1 
		strVal = BIZ_PGM_QUERY_ID & "?txtRcptNo=" & Trim(.txtRcptNo.value)		'☜: 
		strVal = strVal & "&txtMode=" & parent.UID_M0001
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey						'☆: 조회 조건 데이타    

		' 권한관리 추가 
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd				' 사업장 
		strVal = strVal & "&lgInternalCd="		& lgInternalCd					' 내부부서 
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd				' 내부부서(하위포함)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID					' 개인 

		Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    End With
    
    DbQuery = True    
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()
    If frm1.vspdData.MaxRows > 0 Then
        frm1.vspdData.Row = 1
        frm1.vspdData.Col = C_ItemSeq
        frm1.hItemSeq.Value = frm1.vspdData.Text 
        Call DbQuery2(1)
	End If

    lgIntFlgMode = parent.OPMD_UMODE											'⊙: Indicates that current mode is Update mode    

    Call LayerShowHide(0)

    Call txtInputType_onChange()
	Call ClickTab1()
	Call SetSpreadLock("A")
	Call SetSpreadLock("B")	

	lgQueryok = True 

	Call DoSum()
	Call CurFormatNumSprSheet()
	Call txtDocCur_OnChange()
	Call txtDeptCd_OnBlur()

	frm1.txtRcptNo.focus
	lgBlnFlgChgValue = False
	lgQueryOk = False
End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data Save
'========================================================================================
Function  DbSave()     
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
	Dim strVal
	DIm lngRows
	
	Call LayerShowHide(1)
	
    DbSave = False																				'⊙: Processing is NG
    
    On Error Resume Next																		'☜: Protect system from crashing
    Err.Clear
    
	lgRetFlag = False
	With frm1
		.txtMode.value = parent.UID_M0002														'☜: 저장 상태 
		.txtMode.value = lgIntFlgMode															'☜: 신규입력/수정 상태 
    
		strMode = frm1.txtMode.value
    	'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
		strVal = ""    
		'-----------------------
		'Data manipulate area
		'-----------------------
		For lRow = 1 To .vspdData1.MaxRows
		    .vspdData1.Row = lRow
		    .vspdData1.Col = 0
		    Select Case .vspdData1.Text
		        Case ggoSpread.DeleteFlag														'☜: 삭제 
					
				Case Else	
					strVal = strVal & "C" & parent.gColSep  									'☜: C=Create, Row위치 정보 
		            .vspdData1.Col = C_ArNo														'1
		            strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep
		            .vspdData1.Col = C_ArAcctCd
		            strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep
		            .vspdData1.Col = C_ArDt
		            strVal = strVal & UniConvDate(Trim(.vspdData1.Text)) & parent.gColSep
		            .vspdData1.Col = C_Ar_DocCur
		            strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep		            
		            .vspdData1.Col = C_ArClsAmt
		            strVal = strVal & Trim(UNIConvNum(.vspdData1.Text,0)) & parent.gColSep
		            .vspdData1.Col = C_ArClsLocAmt		            
		            strVal = strVal & Trim(UNIConvNum(.vspdData1.Text,0)) & parent.gColSep
		            .vspdData1.Col = C_ArDcAmt
		            strVal = strVal & Trim(UNIConvNum(.vspdData1.Text,0)) & parent.gColSep
		            .vspdData1.Col = C_ArDcLocAmt		            
		            strVal = strVal & Trim(UNIConvNum(.vspdData1.Text,0)) & parent.gColSep
		            .vspdData1.Col = C_ArClsDesc		            
		            strVal = strVal & Trim(.vspdData1.Text) & parent.gRowSep              		            
		            
		            lGrpCnt = lGrpCnt + 1            
		    End Select
		Next
			
		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strVal
		
		lGrpCnt = 1
		strVal = ""	

		ggoSpread.Source = frm1.vspdData
		With frm1.vspdData
			For lngRows = 1 To .MaxRows
				.Row = lngRows
				.Col = 0
				Select Case .Text
					Case ggoSpread.DeleteFlag

					Case Else
						strVal = strVal & "C" & parent.gColSep & lngRows & parent.gColSep		'C=Create, Sheet가 2개 이므로 구별 
						.Col = C_ItemSeq	'1
						strVal = strVal & Trim(.Text) & parent.gColSep
						.Col = C_DcAmt		'2
						strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gColSep
						.Col = C_DcLocAmt		'3
						strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gColSep
						.Col = C_AcctCd		'4
						strVal = strVal & Trim(.Text) & parent.gRowSep	
						        
						lGrpCnt = lGrpCnt + 1
				End Select							        
			Next
		End With
	
		frm1.txtMaxRows1.value = lGrpCnt-1														'Spread Sheet의 변경된 최대갯수 
		frm1.txtSpread1.value  = strVal															'Spread Sheet 내용을 저장    
					
		lGrpCnt = 1
		strVal = ""
		
		With frm1.vspdData3  
			For lngRows = 1 To .MaxRows
				.Row = lngRows
				.Col = 0
				Select Case .Text
					Case ggoSpread.DeleteFlag

					Case Else
						strVal = strVal & "C" & parent.gColSep & lngRows & parent.gColSep		'C=Create, Sheet가 2개 이므로 구별 
					'이곳에서는 컬럼변수를 사용하지 않고 절대위치를 지정해서 스트링을 만들어야한다.
					'왜냐하면 저장될 관리항목들이 frm1.vspdData3에 세팅될 때 절대위치로 이동하기 때문이다.
						.Col =  1                           
						strVal = strVal & Trim(.Text) & parent.gColSep
						.Col =  2
						strVal = strVal & Trim(.Text) & parent.gColSep
						.Col =  3
						strVal = strVal & Trim(.Text) & parent.gColSep
						.Col =  5
						strVal = strVal & Trim(.Text) & parent.gRowSep
						lGrpCnt = lGrpCnt + 1		
				End Select
			Next
		End With

		frm1.txtMaxRows3.value = lGrpCnt-1														'Spread Sheet의 변경된 최대갯수 
		frm1.txtSpread3.value  = strVal															'Spread Sheet 내용을 저장			

		'권한관리추가 start
		frm1.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
		frm1.txthInternalCd.value =  lgInternalCd
		frm1.txthSubInternalCd.value = lgSubInternalCd
		frm1.txthAuthUsrID.value = lgAuthUsrID		
		'권한관리추가 end

		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)												'☜: 비지니스 ASP 를 가동 
	End With	

    DbSave = True																				'⊙: Processing is NG
End Function

'======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'======================================================================================================
Function DbSaveOk()																				'☆: 저장 성공후 실행 로직 

	Call LayerShowHide(0)
    Call ggoOper.ClearField(Document, "2")														'⊙: Clear Contents  Field
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData1
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData3
	ggoSpread.ClearSpreadData
	
    Call InitVariables()    

    Call DBQuery()	
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function  DbDelete() 
	Dim strVal 

    Call LayerShowHide(1)
    
    Err.Clear
    
	DbDelete = False																			'⊙: Processing is NG
    
    With frm1
		.txtMode.value = parent.UID_M0003
	End With
    strMode = frm1.txtMode.value
    
    strVal = BIZ_PGM_DEL_ID & "?txtMode=" & parent.UID_M0003
    strVal = strVal & "&txtRcptNo=" & Trim(frm1.txtAllcNo.value)								'☜: 삭제 조건 데이타 
	' 권한관리 추가 
	strVal = strVal & "&txthAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
	strVal = strVal & "&txthInternalCd="	& lgInternalCd				' 내부부서 
	strVal = strVal & "&txthSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
	strVal = strVal & "&txthAuthUsrID="		& lgAuthUsrID				' 개인    
    
	Call RunMyBizASP(MyBizASP, strVal)															'☜: 비지니스 ASP 를 가동 
    
    DbDelete = True																				'⊙: Processing is NG
End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================
Function DbDeleteOk()																			'☆: 삭제 성공후 실행 로직	
	Call LayerShowHide(0)
	Call ggoOper.ClearField(Document, "1")														'⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")														'⊙: Clear Contents  Field
    Call ggoOper.LockField(Document, "N")														'⊙: Lock  Suitable  Field
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData1
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData3
	ggoSpread.ClearSpreadData
    
    Call txtInputType_onChange()
    Call InitVariables()															   '⊙: Initializes local global variables
    Call ClickTab1()    
    Call SetDefaultVal()
End Function




'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.4 User-defined Method Part
' Description : This part declares user-defined method
'=======================================================================================================
'*******************************************************************************************************




'=======================================================================================================
' Function Name : DbQuery2																				
' Function Desc : This function is data query and display												
'=======================================================================================================
Function DbQuery2(ByVal Row)
	Dim strVal	
	Dim lngRows
		
	Dim strSelect
	Dim strFrom
	Dim strWhere 	
	
	Dim strTableid
	Dim strColid
	Dim strColNm	
	Dim strMajorCd	
	Dim strNmwhere
	Dim i,Indx1
	Dim arrVal,arrTemp
	
	Err.Clear
	
	With frm1
	    .vspdData.Row = Row
	    .vspdData.Col = C_ItemSeq
	    .hItemSeq.Value = .vspdData.Text

	    If Trim(.hItemSeq.Value) = "" Then
	        Exit Function
	    End If
	    
        If CopyFromData(.hItemSeq.Value) = True Then
			Call SetSpread2ColorAr()
            Exit Function
        End If
    	
		Call LayerShowHide(1)
	
		DbQuery2 = False
		
		.vspdData.Row = Row
		.vspdData.Col = C_ItemSeq
		
		strSelect =				" C.DTL_SEQ,  A.CTRL_CD, A.CTRL_NM , C.CTRL_VAL, '',"
		strSelect = strSelect & " CASE  WHEN A.ColM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  THEN " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  END , D.SEQ, LTrim(ISNULL(A.TBL_ID,'')), LTrim(ISNULL(A.DATA_COLM_ID,'')), "
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM,'')),  LTrim(ISNULL(A.ColM_DATA_TYPE,'')), LTrim(ISNULL(A.DATA_LEN,'')), "
		strSelect = strSelect & " CASE WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("DC", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  THEN " & FilterVar("D", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("C", "''", "S") & "  "
		strSelect = strSelect & " END	, " & .hItemSeq.Value & ",  "
		strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD,'')), CHAR(8) "
    		
		strFrom = " A_CTRL_ITEM	A (NOLOCK), A_ACCT_CTRL_ASSN B (NOLOCK), A_RCPT_DC_DTL C (NOLOCK), A_RCPT_DC D (NOLOCK) "
		
		strWhere =			  " D.ALLC_NO =  " & FilterVar(UCase(.txtALLCNo.value), "''", "S") & "  "
		strWhere = strWhere & " AND D.SEQ = " & .hItemSeq.Value & " "
		strWhere = strWhere & " AND D.ALLC_NO  =  C.ALLC_NO  "
		strWhere = strWhere & " AND D.SEQ  =  C.SEQ "
		strWhere = strWhere & "	AND D.ACCT_CD = B.ACCT_CD "
		strWhere = strWhere & " AND C.CTRL_CD = B.CTRL_CD "
		strWhere = strWhere & " AND D.ACCT_CD = B.ACCT_CD "
		strWhere = strWhere & " AND B.CTRL_CD = A.CTRL_CD "
		strWhere = strWhere & " ORDER BY C.DTL_SEQ "
			
		frm1.vspdData2.ReDraw = False
			
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then   
			ggoSpread.Source = frm1.vspdData2			
			arrTemp =  Split(lgF2By2,Chr(12))
			For Indx1 = 0 To Ubound(arrTemp) - 1
				arrTemp(indx1) = Replace(arrTemp(indx1), Chr(8), indx1 + 1)
			Next
			lgF2By2 = Join(arrTemp,Chr(12))
			ggoSpread.SSShowData lgF2By2
		
			For lngRows = 1 To frm1.vspdData2.MaxRows
				frm1.vspddata2.Row = lngRows	
				frm1.vspdData2.Col = C_Tableid 
				If Trim(frm1.vspddata2.text) <> "" Then
					frm1.vspdData2.Col = C_Tableid
					strTableid = frm1.vspddata2.text
					frm1.vspdData2.Col = C_Colid
					strColid = frm1.vspddata2.text
					frm1.vspdData2.Col = C_ColNm
					strColNm = frm1.vspddata2.text	
					frm1.vspdData2.Col = C_MajorCd					
					strMajorCd = frm1.vspddata2.text	
					
					frm1.vspdData2.Col = C_CtrlVal
					
					strNmwhere = strColid & " =   " & FilterVar(frm1.vspddata2.text, "''", "S") & "  " 
					
					If Trim(strMajorCd) <> "" Then
						strNmwhere = strNmwhere & " AND MAJOR_CD =  " & FilterVar(strMajorCd, "''", "S") & "  "
					End If
					
					If CommonQueryRs( strColNm , strTableid ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   
						frm1.vspdData2.Col = C_CtrlValNm
						arrVal = Split(lgF0, Chr(11))  
						frm1.vspddata2.text = arrVal(0)
					End If
				End If								
				
				strVal = strVal & Chr(11) & .hItemSeq.Value 
				frm1.vspdData2.Col = C_DtlSeq  
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_CtrlCd   
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_CtrlNm   
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_CtrlVal 
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_CtrlPB   
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_CtrlValNm 
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_Seq 
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_Tableid 
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_Colid 
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_ColNm 
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_Datatype 
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_DataLen 
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_DRFg 
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_MajorCd 
				strVal = strVal & Chr(11) & .vspddata2.text
				frm1.vspdData2.Col = C_MajorCd+1 				
				.vspdData2.Text = lngRows
				strVal = strVal & Chr(11) & .vspddata2.text
				strVal = strVal & Chr(11) & Chr(12)									
			Next					
			ggoSpread.Source = frm1.vspdData3
			ggoSpread.SSShowData strVal	
		End If 		
		
'		Call CopyFromData(.hItemSeq.value)
		Call SetSpread2ColorAr()
	End With
	
	Call LayerShowHide(0)
	
	frm1.vspdData2.ReDraw = True
	
	DbQuery2 = True
End Function

'=======================================================================================================
' Function Name : DbQueryOk2
' Function Desc : DbQuery2가 성공적일 경우 MyBizASP 에서 호출되는 Function
'=======================================================================================================
Function  DbQueryOk2()
    Call SetSpread2ColorAr()
End Function

'=======================================================================================================
'   Function Name : chkAllcDate
'   Function Desc : 
'=======================================================================================================
Function chkAllcDate()
	Dim intI
	
	chkAllcDate = True
	With frm1
		For intI = 1 To .vspdData1.MaxRows
			.vspdData1.Row = intI
			.vspdData1.Col = C_ArDt
			If CompareDateByFormat(.vspdData1.Text,.txtRcptDt.Text,"채권일자",.txtRcptDt.Alt, _
		    	               "970025",.txtRcptDt.UserDefinedFormat,parent.gComDateType, true) = False Then
			   .txtRcptDt.focus
			   chkAllcDate = False
			   Exit Function
			End If
		Next
	End With
End Function

'=======================================================================================================
'   Function Name : chkAllcDate
'   Function Desc : 
'=======================================================================================================
Sub DoSum()
	Dim dblTotArAmt
	Dim dblTotArRemAmt
	Dim dblTotArClsAmt 
	Dim dblTotArClsLocAmt 
	Dim dblTotArDcAmt
	Dim dblTotArDcLocAmt

	With frm1.vspdData1
		dblTotArAmt      = FncSumSheet1(frm1.vspdData1,C_ArAmt     , 1, .MaxRows, False, -1, -1, "V")
		dblTotArRemAmt   = FncSumSheet1(frm1.vspdData1,C_ArRemAmt  , 1, .MaxRows, False, -1, -1, "V")
		dblTotArClsAmt   = FncSumSheet1(frm1.vspdData1,C_ArClsAmt  , 1, .MaxRows, False, -1, -1, "V")
		dblTotArDcAmt    = FncSumSheet1(frm1.vspdData1,C_ArDcAmt   , 1, .MaxRows, False, -1, -1, "V")
		dblTotArDcLocAmt = FncSumSheet1(frm1.vspdData1,C_ArDcLocAmt, 1, .MaxRows, False, -1, -1, "V")
		
		If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
			If lgQueryOk = False Then
				If UCase(Trim(frm1.hArDocCur.Value)) = UCase(Trim(frm1.txtDocCur.Value)) Then
					frm1.txtRcptAmt.text = UNIConvNumPCToCompanyByCurrency(dblTotArClsAmt,frm1.txtDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
				Else
'					frm1.txtRcptAmt.text = "0"			
				End If
			End If				
		End If	
		
		frm1.txtTotArAmt.text	 = UNIConvNumPCToCompanyByCurrency(dblTotArAmt,frm1.hArDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
		frm1.txtTotArRemAmt.text = UNIConvNumPCToCompanyByCurrency(dblTotArRemAmt,frm1.hArDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
		frm1.txtDcAmt.text	     = UNIConvNumPCToCompanyByCurrency(dblTotArDcAmt,frm1.hArDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
		frm1.txtTotDcAmt.text	 = UNIConvNumPCToCompanyByCurrency(dblTotArDcAmt,frm1.hArDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
					
		frm1.txtDcLocAmt.text    = UNIConvNumPCToCompanyByCurrency(dblTotArDcLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")
        frm1.txtTotDcLocAmt.text = UNIConvNumPCToCompanyByCurrency(dblTotArDcLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")
	End With
End Sub    
    
'==========================================================================================
'   Event Name : txtDocCur_OnChange
'   Event Desc : 
'==========================================================================================
Sub txtDocCur_OnChange()
	Dim iRef

    lgBlnFlgChgValue = True
    
	If CommonQueryRs( "reference" , "b_configuration" , " major_cd=" & FilterVar("a1006", "''", "S") & "  and minor_cd =  " & FilterVar(frm1.txtinputtype.value , "''", "S") & " and seq_no=4 " , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
		iRef = split(lgF0,Chr(11))
		If lgQueryOk = False Then
			If UCase(Trim(frm1.txtDocCur.value)) <> UCase(parent.gCurrency) Then 
				If iRef(0) = "NO" Then
					frm1.txtInputType.value = ""
					frm1.txtInputTypeNm.value = ""
					frm1.txtAcctCd.value = ""
					frm1.txtAcctNm.value = ""					
					frm1.txtCheckCd.value = ""
					Call ggoOper.SetReqAttr(frm1.txtCheckCd,  "Q")
				End If
			End If			
		End If
	Else
		frm1.txtInputType.value = ""
		Call ggoOper.SetReqAttr(frm1.txtCheckCd,  "Q")
	End If

	If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        									
		Call CurFormatNumericOCX()
'		Call CurFormatNumSprSheet()		
		Call DoSum()
	End If
	
    If lgQueryok <> True then
		If UCase(parent.gCurrency) <> UCase(Trim(frm1.txtDocCur.value)) Then
			frm1.txtXchRate.Text = "0"
		Else
			frm1.txtXchRate.Text = "1"			
		End If		
	End If
End Sub

'==========================================================================================
'   Event Name : txtRcptAmt_Change
'   Event Desc : 
'==========================================================================================
Sub txtRcptAmt_Change()
    lgBlnFlgChgValue = True	
End sub

'==========================================================================================
'   Event Name : txtXchRate_Change
'   Event Desc : 
'==========================================================================================
Sub txtXchRate_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtInputType_Change()
'   Event Desc :  
'=======================================================================================================
Sub  txtInputType_onChange()
	Dim IntRetCD 
	
    lgBlnFlgChgValue = True
    
	If CommonQueryRs( "REFERENCE" , "B_CONFIGURATION " , "MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND MINOR_CD =  " & FilterVar(frm1.txtInputType.value, "''", "S") & "  AND SEQ_NO = 4 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then 
		Select Case UCase(Trim(Left(lgF0, Len(lgF0)-1)))
			Case "CS" 
				frm1.txtCheckCd.value   = ""
				frm1.txtBankCd.value   = ""
				frm1.txtBankAcct.value   = ""
				spnNoteInfo.innerHTML =  "어음번호"
				Call ggoOper.SetReqAttr(frm1.txtBankCd,   "Q")
				Call ggoOper.SetReqAttr(frm1.txtBankAcct, "Q")
				Call ggoOper.SetReqAttr(frm1.txtCheckCd,   "Q")
			Case "DP" 		' 예적금 
				frm1.txtCheckCd.value   = ""
				spnNoteInfo.innerHTML =  "어음번호"
				Call ggoOper.SetReqAttr(frm1.txtBankCd,   "N")
				Call ggoOper.SetReqAttr(frm1.txtBankAcct, "N")
				Call ggoOper.SetReqAttr(frm1.txtCheckCd,   "Q")
			Case "NO"
				If UCase(Trim(frm1.txtInputType.value)) = "CR" Then
					spnNoteInfo.innerHTML =  "수취구매카드번호"
				Else
					spnNoteInfo.innerHTML =  "어음번호"
				End If
				
				If UCase(Trim(frm1.txtDocCur.value)) = parent.gCurrency Then
					frm1.txtBankCd.value   = ""
					frm1.txtBankAcct.value   = ""				
					Call ggoOper.SetReqAttr(frm1.txtBankCd,   "Q")
					Call ggoOper.SetReqAttr(frm1.txtBankAcct, "Q")
					Call ggoOper.SetReqAttr(frm1.txtCheckCd,  "N")	
				Else
					IntRetCD = DisplayMsgBox("111620","X","X","X")  
					frm1.txtInputType.value = ""
					frm1.txtInputTypeNm.value = ""					
					Call ggoOper.SetReqAttr(frm1.txtCheckCd,  "Q")						
					Exit Sub
				End If					
			Case Else
				frm1.txtCheckCd.value  = ""
				frm1.txtBankCd.value   = ""
				frm1.txtBankAcct.value = ""	
				spnNoteInfo.innerHTML =  "어음번호"
				Call ggoOper.SetReqAttr(frm1.txtBankCd,   "Q")
				Call ggoOper.SetReqAttr(frm1.txtBankAcct, "Q")
				Call ggoOper.SetReqAttr(frm1.txtCheckCd,   "Q")
		End Select
	End If
	
	If frm1.txtInputType.value = "" Then
		Call ggoOper.SetReqAttr(frm1.txtBankCd,   "Q")
		Call ggoOper.SetReqAttr(frm1.txtBankAcct, "Q")	
		Call ggoOper.SetReqAttr(frm1.txtCheckCd,   "Q")	
	End If
	
	frm1.txtAcctCd.value = "" :	frm1.txtAcctnm.value = ""	
End Sub

'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()
	With frm1
		' 입금금액 
		ggoOper.FormatFieldByObjectOfCur .txtRcptAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 할인금액 
		ggoOper.FormatFieldByObjectOfCur .txtXchRate, .txtDocCur.value, parent.ggExchRateNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec		
	End With
End Sub

'===================================== CurFormatNumSprSheet()  ======================================
'	Name : CurFormatNumSprSheet()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet()
	With frm1
		ggoSpread.Source = frm1.vspdData1
		' 채권액 
		ggoSpread.SSSetFloatByCellOfCur C_ArAmt,-1, .hArDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		' 채권잔액 
		ggoSpread.SSSetFloatByCellOfCur C_ArRemAmt,-1, .hArDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		' 반제금액 
		ggoSpread.SSSetFloatByCellOfCur C_ArClsAmt,-1, .hArDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		' 할인금액 
		ggoSpread.SSSetFloatByCellOfCur C_ArDcAmt,-1, .hArDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		
		ggoSpread.Source = frm1.vspdData
		' 할인금액 
		ggoSpread.SSSetFloatByCellOfCur C_DcAmt,-1, .hArDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		
		ggoOper.FormatFieldByObjectOfCur .txtDcAmt, .hArDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 채권액 
		ggoOper.FormatFieldByObjectOfCur .txtTotArAmt, .hArDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 채권잔액 
		ggoOper.FormatFieldByObjectOfCur .txtTotArRemAmt, .hArDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 할인금액 
		ggoOper.FormatFieldByObjectOfCur .txtTotDcAmt, .hArDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec		
	End With
End Sub

'====================================================================================================
'	Name : XchLocRate()
'	Description : 환율이 변경되는 Factor 가 변했을 때 수정되는 Local Amt. Setting
'====================================================================================================
Sub XchLocRate()
	Dim ii

	With frm1
		For ii = 1 To .vspdData1.MaxRows 
			.vspdData1.Row = ii	
			.vspdData1.Col = C_ArClsLocAmt	
			.vspdData1.Text = ""   
			.vspdData1.Col = C_ArDcLocAmt	
			.vspdData1.Text = ""    	
			ggoSpread.Source = .vspdData1
			ggoSpread.UpdateRow ii
		Next	
						
		For ii = 1 To .vspdData.MaxRows 
			.vspdData.Row = ii	
			.vspdData.Col = C_DcLocAmt	
			.vspdData.Text = ""    		
			ggoSpread.Source = .vspdData
			ggoSpread.UpdateRow ii
		Next
		.txtDcLocAmt.text="0"
		.txtTotDcLocAmt.text="0"
	End With
End Sub


'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.5 Spread Popup method 
' Description : This part declares spread popup method
'=======================================================================================================
'*******************************************************************************************************




'===================================== PopSaveSpreadColumnInf()  ======================================
' Name : PopSaveSpreadColumnInf()
' Description : 이동한 컬럼의 정보를 저장 
'====================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'===================================== PopRestoreSpreadColumnInf()  ======================================
' Name : PopRestoreSpreadColumnInf()
' Description : 컬럼의 순서정보를 복원함 
'====================================================================================================
Sub  PopRestoreSpreadColumnInf()
	Dim indx

	ggoSpread.Source = gActiveSpdSheet
	Select Case Trim(UCase(gActiveSpdSheet.Name))
		Case "VSPDDATA" 
			Call PrevspdDataRestore(gActiveSpdSheet)
			Call ggoSpread.RestoreSpreadInf()
			Call InitSpreadSheet("B")
			Call ggoSpread.ReOrderingSpreadData()
			Call SetSpreadLock("B")
			Call SetSpread2ColorAr()									
		Case "VSPDDATA1" 
'			Call PrevspdDataRestore(gActiveSpdSheet)
			Call ggoSpread.RestoreSpreadInf()
			Call InitSpreadSheet("A")
			Call ggoSpread.ReOrderingSpreadData()			
			Call SetSpreadLock("A")
		Case "VSPDDATA2"
			Call PrevspdData2Restore(gActiveSpdSheet)   
			Call ggoSpread.RestoreSpreadInf()
			Call InitCtrlSpread()			'관리항목 그리드 초기화 
			Call ggoSpread.ReOrderingSpreadData()
			Call SetSpread2ColorAr()  
	End Select
	
	If frm1.vspdData2.MaxRows <= 0 Then
		Call DbQuery2(frm1.vspdData.ActiveRow)
	End If		
End Sub

'===================================== PrevspdDataRestore()  ========================================
' Name : PrevspdDataRestore()
' Description : 그리드 복원시 관리항목 복원 
'====================================================================================================
Sub PrevspdDataRestore(pActiveSheetName)
	Dim indx, indx1

	For indx = 0 To frm1.vspdData.MaxRows
        frm1.vspdData.Row    = indx
        frm1.vspdData.Col    = 0
		
		If frm1.vspdData.Text <> "" Then
			Select Case frm1.vspdData.Text			
				Case ggoSpread.InsertFlag					
					frm1.vspdData.Col = C_ItemSeq					
					Call DeleteHsheet(frm1.vspdData.Text)					
				Case ggoSpread.UpdateFlag		
					For indx1 = 0 To frm1.vspdData3.MaxRows					
						frm1.vspdData3.Row = indx1
						frm1.vspdData3.Col = 0
						Select Case frm1.vspdData3.Text 
							Case ggoSpread.UpdateFlag
								frm1.vspdData.Col = C_ItemSeq
								frm1.vspdData3.Col = 1					
								If UCase(Trim(frm1.vspdData.Text)) = UCase(Trim(frm1.vspdData3.Text)) Then
									Call DeleteHsheet(frm1.vspdData.Text)										
									Call FncRestoreDbQuery2(indx, frm1.vspdData.ActiveRow, frm1.txtAllcNo.Value)
								End If
						End Select
					Next
				Case ggoSpread.DeleteFlag
					Call fncRestoreDbQuery2(indx, frm1.vspdData.ActiveRow, frm1.txtAllcNo.Value)
			End Select
		End If
	Next
	
	ggoSpread.Source = pActiveSheetName
End Sub

'===================================== PrevspdDataRestore()  ========================================
' Name : PrevspdData2Restore()
' Description : 그리드 복원시 관리항목 복원 
'====================================================================================================
Sub PrevspdData2Restore(pActiveSheetName)
	Dim indx, indx1

	For indx = 0 to frm1.vspdData2.MaxRows
        frm1.vspdData2.Row    = indx
        frm1.vspdData2.Col    = 0

		If frm1.vspdData2.Text <> "" Then
			Select Case frm1.vspdData2.Text
				Case ggoSpread.InsertFlag
					frm1.vspdData2.Col = C_HItemSeq
					For indx1 = 0 to frm1.vspdData.MaxRows
						frm1.vspdData.Row = indx1
						frm1.vspdData.Col = C_ItemSeq
						If frm1.vspdData.Text = frm1.vspdData2.Text Then
							Call DeleteHsheet(frm1.vspdData.Text)
							ggoSpread.Source = frm1.vspdData	
					        ggoSpread.EditUndo							
						End If
					Next
				Case ggoSpread.UpdateFlag
					frm1.vspdData2.Col = C_HItemSeq
					For indx1 = 0 to frm1.vspdData.MaxRows
						frm1.vspdData.Row = indx1
						frm1.vspdData.Col = C_ItemSeq
						If frm1.vspdData.Text = frm1.vspdData2.Text Then
							Call DeleteHsheet(frm1.vspdData.Text)
							ggoSpread.Source = frm1.vspdData
							ggoSpread.EditUndo
							Call fncRestoreDbQuery2(indx1, frm1.vspdData.ActiveRow, frm1.txtAllcNo.Value) 
						End If
					Next
				Case ggoSpread.DeleteFlag

			End Select
		End If
	Next
	ggoSpread.Source = pActiveSheetName
End Sub

'========================================================================================================
' Name : fncRestoreDbQuery2																				
' Desc : This function is data query and display												
'========================================================================================================
Function fncRestoreDbQuery2(Row, CurrRow, Byval pInvalue1)
	Dim strItemSeq
	Dim strSelect, strFrom, strWhere
	Dim arrTempRow, arrTempCol
	Dim Indx1
	Dim strTableid, strColid, strColNm, strMajorCd
	Dim strNmwhere
	Dim arrVal
	Dim strVal

	On Error Resume Next
	Err.Clear

	fncRestoreDbQuery2 = False

	Call DisableToolBar(parent.TBC_QUERY)
	Call LayerShowHide(1)
	With frm1
		.vspdData.row = Row
	    .vspdData.col = C_ItemSeq
		strItemSeq    = .vspdData.Text

	    If Trim(strItemSeq) = "" Then
	        Exit Function
	    End If
    	
		Call LayerShowHide(1)
	
		DbQuery2 = False
		
		.vspdData.Row = Row
		.vspdData.Col = C_ItemSeq
		
		strSelect =				" C.DTL_SEQ,  A.CTRL_CD, A.CTRL_NM , C.CTRL_VAL, '',"
		strSelect = strSelect & " CASE  WHEN A.ColM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  THEN " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  END , D.SEQ, LTrim(ISNULL(A.TBL_ID,'')), LTrim(ISNULL(A.DATA_COLM_ID,'')), "
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM,'')),  LTrim(ISNULL(A.ColM_DATA_TYPE,'')), LTrim(ISNULL(A.DATA_LEN,'')), "
		strSelect = strSelect & " CASE WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("DC", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  THEN " & FilterVar("D", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("C", "''", "S") & "  "
		strSelect = strSelect & " END	, " & strItemSeq & ",  "
		strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD,'')), CHAR(8) "
    		
		strFrom = " A_CTRL_ITEM	A (NOLOCK), A_ACCT_CTRL_ASSN B (NOLOCK), A_RCPT_DC_DTL C (NOLOCK), A_RCPT_DC D (NOLOCK) "
		
		strWhere =			  " D.ALLC_NO =  " & FilterVar(UCase(.txtALLCNo.value), "''", "S") & "  "
		strWhere = strWhere & " AND D.SEQ = " & strItemSeq & " "
		strWhere = strWhere & " AND D.ALLC_NO  =  C.ALLC_NO  "
		strWhere = strWhere & " AND D.SEQ  =  C.SEQ "
		strWhere = strWhere & "	AND D.ACCT_CD = B.ACCT_CD "
		strWhere = strWhere & " AND C.CTRL_CD = B.CTRL_CD "
		strWhere = strWhere & " AND D.ACCT_CD = B.ACCT_CD "
		strWhere = strWhere & " AND B.CTRL_CD = A.CTRL_CD "
		strWhere = strWhere & " ORDER BY C.DTL_SEQ "
				
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then
			arrTempRow =  Split(lgF2By2, Chr(12))
			For Indx1 = 0 To Ubound(arrTempRow) - 1
				arrTempCol = split(arrTempRow(indx1), Chr(11))
				If Trim(arrTempCol(8)) <> "" Then
					strTableid = arrTempCol(8)
					strColid   = arrTempCol(9)
					strColNm   = arrTempCol(10)
					strMajorCd = arrTempCol(15)
					
					strNmwhere = strColid & " =   " & FilterVar(arrTempCol(C_CtrlVal), "''", "S") & "  " 

					If Trim(strMajorCd) <> "" Then
						strNmwhere = strNmwhere & " AND MAJOR_CD =  " & FilterVar(strMajorCd, "''", "S") & "  "
					End If

					If CommonQueryRs( strColNm , strTableid ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
						arrVal = Split(lgF0, Chr(11))
						arrTempCol(6) = arrVal(0)
					End If
				End If

				strVal = strVal & Chr(11) & strItemSeq
				strVal = strVal & Chr(11) & arrTempCol(1)
				strVal = strVal & Chr(11) & arrTempCol(2)
				strVal = strVal & Chr(11) & arrTempCol(3)
				strVal = strVal & Chr(11) & arrTempCol(4)
				strVal = strVal & Chr(11) & arrTempCol(5)
				strVal = strVal & Chr(11) & arrTempCol(6)
				strVal = strVal & Chr(11) & arrTempCol(7)
				strVal = strVal & Chr(11) & arrTempCol(8)
				strVal = strVal & Chr(11) & arrTempCol(9)
				strVal = strVal & Chr(11) & arrTempCol(10)
				strVal = strVal & Chr(11) & arrTempCol(11)
				strVal = strVal & Chr(11) & arrTempCol(12)
				strVal = strVal & Chr(11) & arrTempCol(13)
				strVal = strVal & Chr(11) & arrTempCol(15)
				strVal = strVal & Chr(11) & Indx1 + 1
				strVal = strVal & Chr(11) & Chr(12)
			Next
			ggoSpread.Source = .vspdData3
			ggoSpread.SSShowData strVal	
		End If 		

		If Row = CurrRow Then
			Call CopyFromData (strItemSeq)
		End If

		Call LayerShowHide(0)
		Call RestoreToolBar()
	End With

	If Err.number = 0 Then
		fncRestoreDbQuery2 = True
	End If
End Function


'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.6 Spread OCX Tag Event
' Description : This part declares Spread OCX Tag Event
'=======================================================================================================
'*******************************************************************************************************



'=======================================================================================================
'   Event Name : vspdData1_onfocus
'   Event Desc :
'=======================================================================================================
Sub  vspdData_onfocus()
   If lgIntFlgMode <> parent.OPMD_UMODE Then    
        Call SetToolBar("1110111100001111")                                     '버튼 툴바 제어 
    Else                 
        Call SetToolBar("1111111100001111")                                     '버튼 툴바 제어 
    End If    
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopUpMenuItemInf("1101111111")
	    
    gMouseClickStatus = "SP2C" 'Split 상태코드 
	Set gActiveSpdSheet = frm1.vspdData
	
	If frm1.vspdData.MaxRows = 0 then
	    Exit Sub
	End if
	
	If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If 
		Exit sub   
    End If

	If Col <> C_AcctCd then
	    Exit Sub
    End If

	ggoSpread.Source = frm1.vspdData
	frm1.vspdData.Row = frm1.vspdData.ActiveRow	

 	frm1.vspdData.Col = C_AcctCd
	
    If Len(frm1.vspdData.Text) > 0 Then

	Else
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.ClearSpreadData
	End If		
End Sub

'==========================================================================================
'   Event Name : vspdData1_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)
	Call SetPopUpMenuItemInf("0000111111")

    gMouseClickStatus = "SPC" 'Split 상태코드 
	Set gActiveSpdSheet = frm1.vspdData1
	
	If frm1.vspdData1.MaxRows = 0 then
	    Exit Sub
	End if
	
	If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData1
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If 
		Exit sub   
    End If
End Sub

'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub

'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData1_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'======================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 상세내역 그리드의 (멀티)컬럼의 너비를 조절하는 경우 
'=======================================================================================================
Sub  vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.Source = frm1.vspdData
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'======================================================================================================
'   Event Name : vspdData1_ColWidthChange
'   Event Desc : 상세내역 그리드의 (멀티)컬럼의 너비를 조절하는 경우 
'=======================================================================================================
Sub  vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.Source = frm1.vspdData1
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'======================================================================================================
'   Event Name :vspdData_DblClick
'   Event Desc :
'=======================================================================================================
Sub  vspdData_DblClick( ByVal Col , ByVal Row )
    If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If
	
	If Row <=0 Then
		Exit Sub			
	End If		
End Sub

'======================================================================================================
'   Event Name :vspdData_DblClick
'   Event Desc :
'=======================================================================================================
Sub  vspdData1_DblClick(ByVal Col , ByVal Row)
    If frm1.vspdData1.MaxRows = 0 Then
		Exit Sub
	End If
	
	If Row <=0 Then
		Exit Sub					
	End If		
End Sub

'======================================================================================================
'   Event Name :vspddata_ScriptDragDropBlock
'   Event Desc :
'=======================================================================================================
Sub  vspddata_ScriptDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
	ggoSpread.Source = frm1.vspdData 
	Call ggoSpread.SpreadDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
	Call GetSpreadColumnPos("B")
End Sub

'======================================================================================================
'   Event Name :vspddata_ScriptDragDropBlock
'   Event Desc :
'=======================================================================================================
Sub  vspddata1_ScriptDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
	ggoSpread.Source = frm1.vspdData1 
	Call ggoSpread.SpreadDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
	Call GetSpreadColumnPos("A")
End Sub

'=======================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc :
'=======================================================================================================
Sub  vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If Row <> NewRow And NewRow > 0 Then
        With frm1        
            .vspdData.Row = NewRow
            .vspddata1.Col = C_ArNo
                        
            .vspdData.Col = C_ItemSeq
            .hItemSeq.value = .vspdData.Text
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.ClearSpreadData
        End With
        
        frm1.vspdData.Col = 0
        If frm1.vspdData.Text = ggoSpread.DeleteFlag Then
			Exit Sub
        End if
        lgCurrRow = NewRow
        Call DbQuery2(lgCurrRow)
    End If
End Sub

'======================================================================================================
'   Event Name : vspdData1_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub  vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1

	With frm1.vspdData
        ggoSpread.Source = frm1.vspdData
       
        If Row > 0 And Col = C_AcctPB Then
            .Col = Col - 1
            .Row = Row
            Call OpenPopup(.Text, 4)
        End If    
    End With
End Sub

'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'=======================================================================================================
Sub  vspdData_Change(ByVal Col, ByVal Row )
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    
    frm1.vspdData.Row = Row
    frm1.vspdData.Col = 0
    
    Select Case Col
		Case  C_AcctCD
			If frm1.vspdData.Text = ggoSpread.InsertFlag Then
			    frm1.vspdData.Col = C_ItemSeq
			    frm1.hItemSeq.value = frm1.vspdData.Text
			    frm1.vspdData.Col = C_AcctCd
			    If Len(frm1.vspdData.Text) > 0 Then
					frm1.vspdData.Row = Row
					frm1.vspdData.Col = C_ItemSeq	   	
					DeleteHsheet frm1.vspdData.Text
			        Call DbQuery3(Row)
					Call SetSpread2ColorAR()
			    End If    
			End If 
		Case C_DcAmt	
			frm1.vspdData.col = C_DcLocAmt
			frm1.vspdData.Text = ""
'			Call DoSum()
	End Select
End Sub

'======================================================================================================
'   Event Name : vspdData1_Change
'   Event Desc :
'=======================================================================================================
Sub  vspdData1_Change(ByVal Col, ByVal Row)
	Dim ArAmt
	Dim ClsAmt
	Dim DcAmt
	Dim dblTotDcAmt
	Dim dblTotClsAmt
	
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.UpdateRow Row
    
    frm1.vspdData1.Row = Row
    frm1.vspdData1.Col = 0             

    Select Case Col
		Case C_ArClsAmt
			frm1.vspdData1.Col = C_ArAmt
			ArAmt = frm1.vspdData1.Text
			frm1.vspdData1.Col = C_ArClsAmt
			ClsAmt = UniCdbl(frm1.vspdData1.Text)
			If (UNICDbl(ArAmt) > 0 And UNICDbl(ClsAmt) < 0) Or (UNICDbl(ArAmt) < 0 And UNICDbl(ClsAmt) > 0) Then
				frm1.vspdData1.Col = C_ArClsAmt
				frm1.vspdData1.Text = UNIConvNumPCToCompanyByCurrency(ClsAmt * (-1),frm1.hArDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
				
			End If
			
			dblTotClsAmt = FncSumSheet1(frm1.vspdData1,C_ArClsAmt , 1, frm1.vspdData1.MaxRows, False, -1, -1, "V")			
			
			If UCase(Trim(frm1.hArDocCur.Value)) = UCase(Trim(frm1.txtDocCur.Value)) Then
				frm1.txtRcptAmt.text = UNIConvNumPCToCompanyByCurrency(dblTotClsAmt,frm1.txtDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
			End If
			
			If UCase(Trim(frm1.hArDocCur.Value)) <> UCase(parent.gCurrency) Then			
				frm1.vspdData1.Col = C_ArClsLocAmt
				frm1.vspdData1.Row = frm1.vspdData1.ActiveRow				
				frm1.vspdData1.Text = ""
			End If			
		Case C_ArDcAmt
			frm1.vspdData1.Col = C_ArAmt
			ArAmt = frm1.vspdData1.Text
			frm1.vspdData1.Col = C_ArDcAmt
			DcAmt = UNICDbl(frm1.vspdData1.Text)

			If (UNICDbl(ArAmt) > 0 And UNICDbl(DcAmt) < 0) Or (UNICDbl(ArAmt) < 0 And UNICDbl(DcAmt) > 0) Then
				frm1.vspdData1.Col = C_ArDcAmt
				frm1.vspdData1.Text =UNIConvNumPCToCompanyByCurrency(DcAmt * (-1),frm1.hArDocCur.value,parent.ggAmtOfMoneyNo, "X", "X") 
			End If		

			dblTotDcAmt = FncSumSheet1(frm1.vspdData1,C_ArDcAmt , 1, frm1.vspdData1.MaxRows, False, -1, -1, "V")
			
'			If UCase(Trim(frm1.hArDocCur.Value)) = UCase(Trim(frm1.txtDocCur.Value)) Then			
			frm1.txtDcAmt.text = UNIConvNumPCToCompanyByCurrency(dblTotDcAmt ,frm1.hArDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
'			End If
			
			If UCase(Trim(frm1.hArDocCur.Value)) <> UCase(parent.gCurrency) Then
				frm1.vspdData1.Col = C_ArDcLocAmt
				frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
				frm1.vspdData1.Text = ""							
			End If
			
			frm1.txtTotDcAmt.text = UNIConvNumPCToCompanyByCurrency(dblTotDcAmt ,frm1.hArDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")			
		Case C_ArAmt, C_ArRemAmt, C_ArClsLocAmt, C_ArDcLocAmt
			Call DoSum()
	End Select
End Sub

'======================================================================================================
'   Event Name : vspdData1_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
Sub  vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    Dim intListGrvCnt 
    Dim LngLastRow    
    Dim LngMaxRow     
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
End Sub

'======================================================================================================
'   Event Name : vspdData1_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
Sub  vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    Dim intListGrvCnt 
    Dim LngLastRow    
    Dim LngMaxRow     
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
End Sub

'==========================================================================================
'   Event Name : txtDeptCd_OnBlur
'   Event Desc : 
'==========================================================================================
Sub txtDeptCd_OnBlur()
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii, jj

	If Trim(frm1.txtRcptDt.Text = "") Then    
		Exit sub
    End If
    
    lgBlnFlgChgValue = True
	
	If TRim(frm1.txtDeptCd.value) <>"" Then
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtRcptDt.Text, gDateFormat,""), "''", "S") & "))"			
	
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
	
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
			frm1.txtDeptCd.focus
		Else 
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)
					
			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))			
				frm1.hOrgChangeId.value = Trim(arrVal2(2))
			Next	
		End If
	End If	
End Sub

'==========================================================================================
'   Event Name : txtRcptDt_onBlur
'   Event Desc : 
'==========================================================================================
Sub txtRcptDt_onBlur()
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii, jj
	Dim arrVal1, arrVal2
 
  	lgBlnFlgChgValue = True
	With frm1
		If LTrim(RTrim(.txtDeptCd.value)) <> "" and Trim(.txtRcptDt.Text <> "") Then
			strSelect	=			 " Distinct org_change_id "    		
			strFrom		=			 " b_acct_dept(NOLOCK) "		
			strWhere	=			 " dept_Cd =  " & FilterVar(LTrim(RTrim(.txtDeptCd.value)), "''", "S") 
			strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
			strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
			strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(.txtRcptDt.Text, gDateFormat,""), "''", "S") & "))"			
	
			IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
					
			If IntRetCD = False  OR Trim(Replace(lgF0,Chr(11),"")) <> Trim(.hOrgChangeId.value) Then
				.txtDeptCd.value = ""
				.txtDeptNm.value = ""
				.hOrgChangeId.value = ""
				.txtDeptCd.focus
			End If
		End If
	End With
End Sub




'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.7 Date-Numeric OCX Tag Event
' Description : This part declares HTML Tag Event
'=======================================================================================================
'*******************************************************************************************************




'=======================================================================================================
'   Event Name : txtRcptDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtRcptDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtRcptDt.Action = 7
        Call txtRcptDt_onBlur()
        Call SetFocusToDocument("M")
		Frm1.txtRcptDt.Focus 
	End If
End Sub

'=======================================================================================================
'   Event Name : txtRcptDt_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtRcptDt_Change()
    lgBlnFlgChgValue = True
End Sub



'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.8 HTML Tag Event
' Description : This part declares HTML Tag Event
'=======================================================================================================
'*******************************************************************************************************
'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/uni2KCM.inc"  --> 
</HEAD>

<!--'======================================================================================================
'       					6. Tag부 
'	기능: Tag부분 설정 
'======================================================================================================= -->
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
			    <TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/seltab_up_left.gif" width=9></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>반제정보</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/seltab_up_right.gif" width=10></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/tab_up_left.gif" width=9></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>할인상세정보</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/tab_up_right.gif" width=10></td>
						    </TR>
						</TABLE>
					</TD>	
					<TD WIDTH=* align=right><A HREF="VBSCRIPT:OpenPopupTempGL()">결의전표</A>&nbsp;|&nbsp;<A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A>&nbsp;|&nbsp;<a href="vbscript:OpenRefOpenAr()">채권발생정보</A></TD>								
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH="100%" CLASS="Tab11">		
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH="100%">
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>입금번호</TD>
									<TD CLASS="TD656" NOWRAP><INPUT NAME="txtRcptNo" MAXLENGTH=18 ALT="입금번호" STYLE="TEXT-ALIGN: Left" tag ="12XXXU"><IMG align=top name=btnCalType src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON" onclick="vbscript: Call OpenPopup(frm1.txtRcptNo.value,0)"></TD>								
								</TR>						
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>		
					<TD WIDTH="100%">
					
					
					<DIV ID="TabDiv"  SCROLL="no">
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>입금일</TD>
								<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtRcptDt" CLASS=FPDTYYYYMMDD tag="22" Title="FPDATETIME" ALT="입금일" id=fpDateTime1></OBJECT>');</SCRIPT>
								</TD>
								<TD CLASS=TD5 NOWRAP>수금처</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE="Text" NAME="txtBpCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="23NXXU" ALT="수금처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBp(frm1.txtBpCd.value,1)"> <INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24" ALT="거래처명"></TD>
								<TD ></TD>
								<TD ></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>부서</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDeptCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag=22NXXU" ALT="부서"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(frm1.txtDeptCd.Value, 0)"> <INPUT TYPE=TEXT NAME="txtDeptNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="24" ALT="부서명"></TD>
								<TD CLASS=TD5 NOWRAP>입금유형</TD>
								<TD CLASS=TD6 nowrap><INPUT TYPE=TEXT NAME="txtInputType" ALT="입금유형" style="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=5 tag="22NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" style="HEIGHT: 21px; WIDTH: 16px" NAME="btnPayMethod" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtInputType.value, 8)">
													   <INPUT TYPE=TEXT NAME="txtInputTypeNm" ALT="입금유형" style="HEIGHT: 20px; WIDTH: 150px" tag="24X" ></TD>								
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>은행</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE="Text" NAME="txtBankCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="21NXXU" ALT="은행"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBankCd.value,5)"> <INPUT TYPE=TEXT NAME="txtBankNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24" ALT="은행명"></TD>											
								<TD CLASS=TD5 NOWRAP><span id="spnNoteInfo">어음번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCheckCd" SIZE=30 MAXLENGTH=30 STYLE="TEXT-ALIGN: Left" tag="21XXU" ALT="어음번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtCheckCd.value,7)"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>계좌번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT  TYPE=TEXT NAME="txtBankAcct" SIZE=30 MAXLENGTH=30 STYLE="TEXT-ALIGN: Left" tag="21XXXU" ALT="계좌번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBankAcct.value,6)"></TD>
								<TD CLASS=TD5 NOWRAP>입금계정코드</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtAcctCd" ALT="계정코드" MAXLENGTH="20" SIZE=10 STYLE="TEXT-ALIGN: Left" tag="22NXXU"><IMG align=top name=btnCalType onclick="vbscript:CALL OpenPopUp(frm1.txtAcctCd.value,9)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON"> 
													 <INPUT NAME="txtAcctnm" ALT="계정코드명" MAXLENGTH="20"  tag  ="24"></TD>								
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>결의전표번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTempGlNo" SIZE=19 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24XXXU" ALT="결의전표번호"> </TD>																						
								<TD CLASS=TD5 NOWRAP>전표번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtGlNo" SIZE=19 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24XXXU" ALT="전표번호"> </TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>거래통화</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDocCur" SIZE=10 MAXLENGTH=3 tag="23NXXU" STYLE="TEXT-ALIGN: Left" ALT="거래통화"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCur" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript: CALL OpenPopup(frm1.txtDocCur.value,3)"></TD>
								<TD CLASS=TD5 NOWRAP>환율</TD>
								<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtXchRate" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="환율" tag="21X5Z" ></OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>입금금액</TD>
								<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtRcptAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="입금금액" tag="22X2" ></OBJECT>');</SCRIPT>											
								</TD>
								<TD CLASS=TD5 NOWRAP>입금금액(자국통화)</TD>
								<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtRcptLocAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="입금금액(자국통화)" tag="24X2" ></OBJECT>');</SCRIPT>											
								</TD>								
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>할인금액</TD>
								<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtDcAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="할인금액" tag="24X2" ></OBJECT>');</SCRIPT>																							
								</TD>
								<TD CLASS=TD5 NOWRAP>할인금액(자국통화)</TD>
								<TD CLASS=TD6 NOWRAP COLSPAN="4">
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtDcLocAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="할인금액(자국통화)" tag="24X2" ></OBJECT>');</SCRIPT>																							
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>비고</TD>
								<TD CLASS="TD6" NOWRAP COLSPAN=3><INPUT TYPE=TEXT NAME="txtRcptDesc" SIZE=90 MAXLENGTH=128 tag="21XXX" ALT="적요"></TD>
							</TR>						
												
							<TR HEIGHT="100%">
								<TD WIDTH="100%" COLSPAN="4">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" name=vspdData1 width="100%" TITLE="SPREAD" tag="2"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>											
							</TR>						
							<TR>
								<TD  COLSPAN="4">
									<TABLE <%=LR_SPACE_TYPE_60%>>
										<TR>
											<TD class=TD5 NOWRAP>채권액</TD>
											<TD class=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTotArAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="채권액" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>
											<TD class=TD5 STYLE="WIDTH : 0px;"></TD>
											<TD class=TD5 NOWRAP>채권잔액</TD>
											<TD class=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTotArRemAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="채권잔액" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>									
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>						
					</DIV>
					
					
					
					<DIV ID="TabDiv" SCROLL=no>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR HEIGHT="60%">
								<TD WIDTH="100%" COLSPAN="4">
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" name=vspdData width="100%" tag="2" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD COLSPAN=4>
									<TABLE <%=LR_SPACE_TYPE_20%>>
										<TR>							
								<TD class=TD5 NOWRAP>할인금액</TD>
								<TD class=TD6 NOWRAP>									
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTotDcAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="할인금액" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>
								<TD class=TD5 STYLE="WIDTH : 0px;"></TD>
								<TD class=TD5 NOWRAP>할인금액(자국)</TD>
								<TD class=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTotDcLocAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="할인금액(자국)" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>
										</TR>
									</TABLE>
								</TD>									
							</TR>
						    <TR HEIGHT="40%">
								<TD WIDTH="100%" COLSPAN="4">
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" name=vspdData2 width="100%" tag="2" TITLE="SPREAD" > <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT></TD>
							</TR>
						</TABLE>		
					</DIV>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
</TR>
</TABLE>
<TEXTAREA Class=hidden name=txtSpread		tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA Class=hidden name=txtSpread1		tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA Class=hidden name=txtSpread2		tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA Class=hidden name=txtSpread3		tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode"			tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtUpdtUserId"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtInsrtUserId"	tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows1"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows2"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows3"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtFlgMode"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="htxtRcptNo"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtAllcNo"			tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hItemSeq"			tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hAcctCd"			tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hOrgChangeId"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hArDocCur"			tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txthAuthBizAreaCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthSubInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthUsrID"		tag="24" Tabindex="-1">
<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT TYPE=hidden CLASSID=<%=gCLSIDFPSPD%> HEIGHT=0 name=vspdData3 width="100%" tag="2" TABINDEX="-1"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      

