<%@ LANGUAGE="VBSCRIPT" %>

<!--
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : bank Register
'*  3. Program ID           : a3117ma1.asp
'*  4. Program Name         : (-)채무/입금반제 
'*  5. Program Desc         :
'*  6. Comproxy List        : ap001m
'*  7. Modified date(First) : 2000/03/30
'*  8. Modified date(Last)  : 2000/03/30
'*  9. Modifier (First)     : You So Eun
'* 10. Modifier (Last)      : You So Eun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☆) Means that "must change"
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
'********************************************************************************************************* ! -->
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
<SCRIPT LANGUAGE="VBScript"     SRC="../ag/AcctCtrl.vbs">				</SCRIPT>
<SCRIPT LANGUAGE=vbscript>

Option Explicit																		'☜: indicates that All variables must be declared in advance

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Const BIZ_PGM_QRY_ID = "a3117mb1.asp"												'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID = "a3117mb2.asp"												'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_DEL_ID =  "a3117mb3.asp"
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
Const COM_BIZ_EXCHRATE_ID = "../../inc/GetExchRate.asp"								'☆: 환율정보 비지니스 로직 ASP명 

Const TAB1 = 1																		'☜: Tab의 위치 
Const TAB2 = 2

Dim C_ApNo 
Dim C_AcctCd 
Dim C_AcctNm 
Dim C_BizCd 
Dim C_BizNm 
Dim C_ApDt 
Dim C_ApDueDt 
Dim C_DocCur
Dim C_ApAmt 
Dim C_ApRemAmt 
Dim C_ApClsAmt 
Dim C_ApClsLocAmt 
Dim C_ApClsDesc 


Dim  lgStrPrevKey1
Dim  lgStrPrevKeyDtl
Dim  lgStrPrevKey2
Dim  lgStrPrevKey3
Dim  lgRetFlag																		'Popup
Dim  lgCurrRow

Dim  strMode
Dim  intItemCnt					
Dim  IsOpenPop	
Dim  gSelframeFlg

<%
Dim dtToday
dtToday = GetSvrDate
%>

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm

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
Sub initSpreadPosVariables()
	C_ApNo = 1
	C_AcctCd = 2
	C_AcctNm = 3							
	C_BizCd = 4
	C_BizNm = 5
	C_ApDt = 6
	C_ApDueDt = 7
	C_DocCur = 8
	C_ApAmt = 9
	C_ApRemAmt = 10
	C_ApClsAmt = 11
	C_ApClsLocAmt = 12
	C_ApClsDesc = 13
End Sub

'======================================================================================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=======================================================================================================
Sub  InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE												'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False														'Indicates that no value changed
    lgIntGrpCount = 0																'initializes Group View Size
        
    lgStrPrevKey = ""																'initializes Previous Key
    lgStrPrevKey1 = ""
    lgStrPrevKeyDtl = 0																'initializes Previous Key
    lgLngCurRows = 0																'initializes Deleted Rows Count
End Sub

'======================================================================================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=======================================================================================================
Sub  SetDefaultVal()
	frm1.txtAllcDt.text =  UniConvDateAToB("<%=dtToday%>", parent.gServerDateFormat,gDateFormat)
	lgBlnFlgChgValue = False
End Sub

'======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub  LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
End Sub

'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub  InitSpreadSheet()
    Call initSpreadPosVariables()
    
    With frm1.vspdData
    
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadInit "V20021128",,parent.gAllowDragDropSpread 

		.Redraw = False

		.MaxCols = C_ApClsDesc + 1												'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols															'공통콘트롤 사용 Hidden Column
		.ColHidden = True    
		.MaxRows = 0
		    
		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit  C_ApNo       , "채무번호"      , 20, 3		'1
		ggoSpread.SSSetEdit  C_AcctCd     , "계정코드"      , 20, 3	'2
		ggoSpread.SSSetEdit  C_AcctNm     , "계정코드명"    , 20, 3	'3    
		ggoSpread.SSSetEdit  C_BizCd      , "사업장"        , 15,,,10,2	'6
		ggoSpread.SSSetEdit  C_BizNm      , "사업장명"      , 20, 3	'7    
		ggoSpread.SSSetDate  C_ApDt       , "채무일자"      , 10, 2, gDateFormat  
		ggoSpread.SSSetDate  C_ApDueDt    , "만기일자"      , 10, 2, gDateFormat  
		ggoSpread.SSSetEdit  C_DocCur     , "거래통화"      ,  8, 3	'10
		ggoSpread.SSSetFloat C_ApAmt      , "채무액"        , 15, "A"  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat C_ApRemAmt   , "채무잔액"      , 15, "A"  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat C_ApClsAmt   , "반제금액"      , 15, "A"  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat C_ApClsLocAmt, "반제금액(자국)", 15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
    	ggoSpread.SSSetEdit  C_ApClsDesc  , "비고"          , 20, 3	'7   		
    	
		.Redraw = True     	
    End With
   
    Call SetSpreadLock()
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub  SetSpreadLock()
    With frm1.vspdData
		ggoSpread.Source = frm1.vspdData    
		.ReDraw = False

		ggoSpread.SpreadLock C_ApNo    ,-1, C_ApNo
		ggoSpread.SpreadLock C_AcctCd  ,-1, C_AcctCd
		ggoSpread.SpreadLock C_AcctNm  ,-1, C_AcctNm
		ggoSpread.SpreadLock C_BizCd   ,-1, C_BizCd
		ggoSpread.SpreadLock C_BizNm   ,-1, C_BizNm
		ggoSpread.SpreadLock C_ApDt    ,-1, C_ApDt
		ggoSpread.SpreadLock C_ApDueDt ,-1, C_ApDueDt
		ggoSpread.SpreadLock C_DocCur  ,-1, C_DocCur
		ggoSpread.SpreadLock C_ApAmt   ,-1, C_ApAmt
		ggoSpread.SpreadLock C_ApRemAmt,-1, C_ApRemAmt    
		
		ggoSpread.SSSetRequired C_ApClsAmt,-1, -1  		
		
		.ReDraw = True   
    End With
    
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub  SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		.Redraw = False
				
		ggoSpread.SSSetRequired C_ApClsAmt, pvStartRow, pvEndRow
    
		.Col = 1		'절대위치 
		.Row = .ActiveRow
		.Action = 0                         'parent.SS_ACTION_ACTIVE_CELL
		.EditMode = True
		.Redraw = True
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
			ggoSpread.Source = frm1.vspdData

			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)		

			C_ApNo		= iCurColumnPos(1)
			C_AcctCd	= iCurColumnPos(2)
			C_AcctNm	= iCurColumnPos(3)							
			C_BizCd		= iCurColumnPos(4)
			C_BizNm		= iCurColumnPos(5)
			C_ApDt		= iCurColumnPos(6)
			C_ApDueDt	= iCurColumnPos(7)
			C_DocCur	= iCurColumnPos(8)
			C_ApAmt		= iCurColumnPos(9)
			C_ApRemAmt	= iCurColumnPos(10)
			C_ApClsAmt	= iCurColumnPos(11)
			C_ApClsLocAmt = iCurColumnPos(12)
			C_ApClsDesc = iCurColumnPos(13)
	End Select
End Sub

'========================================================================================================= 
'	Name : OpenOpenRefOpenAp()
'	Description : Ref 화면을 call한다. 
'========================================================================================================= 
Function OpenRefOpenAp()
	Dim arrRet
	Dim arrParam(11)	
	Dim iCalledAspName
	
	iCalledAspName = AskPRAspName("a3113ra1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a3113ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = frm1.txtBpCd.value				' 검색조건이 있을경우 파라미터 
	arrParam(1) = frm1.txtBpNm.value			
'	arrParam(2) = frm1.txtDocCur.value					
	arrParam(5) = frm1.txtAllcDt.text
    arrParam(6) = frm1.txtAllcDt.Alt
    
	' 권한관리 추가 
	arrParam(8) = lgAuthBizAreaCd
	arrParam(9) = lgInternalCd
	arrParam(10) = lgSubInternalCd
	arrParam(11) = lgAuthUsrID
        
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0, 0) = "" Then		
		Exit Function
	Else		
		Call SetRefOpenAp(arrRet)
	End If
End Function

'========================================================================================================= 
'	Name : SetRefOpenAp()
'	Description : OpenAp Popup에서 Return되는 값 setting
'========================================================================================================= 
Function SetRefOpenAp(Byval arrRet)
	Dim intRtnCnt, strData
	Dim TempRow, I
	Dim j	
	DIM X
	Dim sFindFg
	
	With frm1.vspdData
		.focus		
		ggoSpread.Source = frm1.vspdData
		.ReDraw = False	
	
		TempRow = .MaxRows												'☜: 현재까지의 MaxRows

		For I = TempRow to TempRow + Ubound(arrRet, 1)
			sFindFg	= "N"
			For x = 1 to TempRow
				.Row = x
				.Col = C_ApNo				
				If "" & UCase(Trim(.Text)) = "" & UCase(Trim(arrRet(I - TempRow, 0))) Then
					sFindFg	= "Y"
				End If
			Next			
			If 	sFindFg	= "N" Then
				.MaxRows = .MaxRows + 1
				.Row = I + 1				
				.Col = 0
				.Text = ggoSpread.InsertFlag

				.Col = C_ApNo												
				.text = arrRet(I - TempRow, 0)				
				.Col = C_AcctCd												
				.text = arrRet(I - TempRow, 1)
				.Col = C_AcctNm												
				.text = arrRet(I - TempRow, 2)				
				.Col = C_BizCd												
				.text = arrRet(I - TempRow, 7)				
				.Col = C_BizNm												
				.text = arrRet(I - TempRow, 8)				
				.Col = C_ApDt 												
				.text = arrRet(I - TempRow, 9)				
				.Col = C_ApDueDt 												
				.text = arrRet(I - TempRow, 10)				
				.Col = C_DocCur												
				.text = arrRet(I - TempRow, 11)
				Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,i+1,i+1,C_DocCur, C_ApAmt,"A" ,"I","X","X")         		
				Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,i+1,i+1,C_DocCur, C_ApRemAmt,"A" ,"I","X","X")         		
				Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,i+1,i+1,C_DocCur, C_ApClsAmt,"A" ,"I","X","X")    						
				.Col = C_ApAmt												
				.text = arrRet(I - TempRow, 13)				
				.Col = C_ApRemAmt 												
				.text = arrRet(I - TempRow, 15)				
				.Col = C_ApClsAmt 												
				.text = arrRet(I - TempRow, 15)	
				.Col = C_ApClsDesc
				.text = arrRet(I - TempRow, 21)							
			End If	
		Next	
		
		frm1.txtDocCur.Value = arrRet(0, 11)				
		frm1.txtbpCd.Value = arrRet(0, 3)				
		frm1.txtbpNm.Value = arrRet(0, 4)				
		
		Call ggoOper.SetReqAttr(frm1.txtALLCDt,   "Q")
		
		ggoSpread.SpreadUnlock   C_ApNo  , TempRow + 1, C_AcctCd , .MaxRows				'⊙: Unlock 컬럼 
		ggoSpread.ssSetProtected C_ApNo  , TempRow + 1, .MaxRows
		ggoSpread.ssSetProtected C_AcctCd, TempRow + 1, .MaxRows						'⊙: Protected	
		
		ggoSpread.SSSetRequired  C_ApClsAmt, TempRow + 1, .MaxRows

		Call txtDocCur_OnChange()		
		.ReDraw = True
    End With
End Function

'========================================================================================================= 
'	Name : OpenRefRcptNo()
'	Description : 
'========================================================================================================= 
Function OpenRefRcptNo()
	Dim arrRet
	Dim arrParam(11)
	Dim iCalledAspName
	
	iCalledAspName = AskPRAspName("a3107ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a3107ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If lgIntFlgMode = parent.OPMD_UMODE Then Exit Function
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = frm1.txtBpCd.value				' 검색조건이 있을경우 파라미터 
	arrParam(1) = frm1.txtBpNm.value			
	arrParam(2) = frm1.txtDocCur.value					
	arrParam(3) = "S"
	arrParam(6) = frm1.txtAllcDt.text
    arrParam(7) = frm1.txtAllcDt.Alt
    
	' 권한관리 추가 
	arrParam(8) = lgAuthBizAreaCd
	arrParam(9) = lgInternalCd
	arrParam(10) = lgSubInternalCd
	arrParam(11) = lgAuthUsrID    
    
	arrRet = window.showModalDialog(iCalledAspName & "?PGM=" & gStrRequestMenuID, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then		
		Exit Function
	Else		
		Call SetRefRcptNo(arrRet)
	End If
End Function

'========================================================================================================= 
'	Name : SetRefRcptNo()
'	Description : Popup에서 Return되는 값 setting
'========================================================================================================= 
Function  SetRefRcptNo(Byval arrRet)
	With frm1
		.txtRcptNo.Value			= arrRet(0)		'C_RcptNo = 1
		.txtRcptDt.text				= arrRet(5)		'C_RcptDt = 8
		.txtBizCd.Value				= arrRet(3)		'C_CostCd = 6	
		.txtBizNm.Value				= arrRet(4)		'C_CostNm = 7	
		.txtBpCd.Value				= arrRet(9)		'C_BizCd = 4
		.txtBpNm.Value				= arrRet(10)	'C_BizNm = 5
		.txtDocCur.value			= arrRet(11)	'C_DocCur = 9		
		.txtBalAmt.Text				= arrRet(7)		'C_RcptAmt = 10
		.txtBalLocAmt.Text			= arrRet(8)		'C_RcptLocAmt = 11
		.txtDeptCd.value			= arrRet(12)	'C_DeptCd = 12
'		.txtDesc.value              = arrRet(13)
		
		.txtAllcNo.value			= ""
		.txtGlNo.value				= ""			
    End With
End Function

'======================================================================================================
' Function Name : OpenPopupGL
' Function Desc : 
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
	
	arrParam(0) = Trim(frm1.txtGlNo.value)	'회계전표번호 
	arrParam(1) = ""						'Reference번호 

	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function

'======================================================================================================
' Function Name : OpenPopupTempGL
' Function Desc : 
'=======================================================================================================
Function OpenPopupTempGL()
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
	
	arrParam(0) = Trim(frm1.txtTempGlNo.value)	'결의전표번호 
	arrParam(1) = ""							'Reference번호 

	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function

'======================================================================================================
'   Function Name : OpenPopUp()
'   Function Desc : 
'=======================================================================================================
Function  OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strCd
	Dim arrParamAdo(8)
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	Select Case iWhere
		Case 1
			arrParam(0) = "거래처팝업"
			arrParam(1) = "B_BIZ_PARTNER"				
			arrParam(2) = Trim(frm1.txtBpCd.Value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "거래처"			
	
			arrField(0) = "BP_CD"	
			arrField(1) = "BP_NM"	
    
			arrHeader(0) = "거래처"		
			arrHeader(1) = "거래처명"	    						' Header명(1)			
		Case 2
			arrParam(0) = "부서팝업"			' 팝업 명칭 
			arrParam(1) = "B_Acct_Dept"					' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtBizCd.Value)		' Code Condition
			arrParam(3) = ""								' Name Cindition
			arrParam(4) = "ORG_CHANGE_ID =  " & FilterVar(parent.gChangeOrgId, "''", "S") & " "			' Where Condition
			arrParam(5) = "부서"			
	
			arrField(0) = "Dept_CD"							' Field명(0)
			arrField(1) = "Dept_NM"							' Field명(1)
    
			arrHeader(0) = "부서"						' Header명(0)
			arrHeader(1) = "부서명"						' Header명(1)    					
		Case 3		
			arrParam(0) = "거래통화팝업"				' 팝업 명칭 
			arrParam(1) = "B_CURRENCY"						' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtDocCur.Value)		' Code Condition
			arrParam(3) = ""								' Name Cindition
			arrParam(4) = ""								' Where Condition
			arrParam(5) = "거래통화"			
	
			arrField(0) = "CURRENCY"							' Field명(0)
			arrField(1) = "CURRENCY_DESC"						' Field명(1)
    
			arrHeader(0) = "거래통화"					' Header명(0)
			arrHeader(1) = "거래통화명"
		Case 4
			arrParam(0) = "계정코드팝업"								' 팝업 명칭 
			arrParam(1) = "A_Acct, A_ACCT_GP" 											' TABLE 명칭 
			arrParam(2) = Trim(strCode)											' Code Condition
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
			arrParam(0) = "은행팝업"
			arrParam(1) = "B_BANK"				
			arrParam(2) = Trim(frm1.txtBankCd.Value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "은행"			
	
			arrField(0) = "BANK_CD"	
			arrField(1) = "BANK_NM"	
    
			arrHeader(0) = "은행"		
			arrHeader(1) = "은행명"	
		Case 6
			arrParam(0) = "계좌번호팝업"
			arrParam(1) = "B_BANK, B_BANK_ACCT"				
			arrParam(2) = Trim(frm1.txtBankAcct.Value)
			arrParam(3) = ""
			
			If Trim(frm1.txtBankCd.Value) = "" Then
				strCd = "B_BANK.BANK_CD = B_BANK_ACCT.BANK_CD "
			Else
				strCd = "B_BANK.BANK_CD = B_BANK_ACCT.BANK_CD AND  B_BANK_ACCT.BANK_CD =  " & FilterVar(frm1.txtBankCd.Value, "''", "S") & " "	
			End If		
			
			arrParam(4) = strCd
			arrParam(5) = "계좌번호"			
			
		    arrField(0) = "B_BANK_ACCT.BANK_ACCT_NO"	
		    arrField(1) = "B_BANK.BANK_CD"	
		    arrField(2) = "B_BANK.BANK_NM"	
		    
		    arrHeader(0) = "계좌번호"		
		    arrHeader(1) = "은행"	
		    arrHeader(2) = "은행명"	
		Case 7
			arrParam(0) = "어음번호팝업"
			arrParam(1) = "F_NOTE"				
			arrParam(2) = Trim(frm1.txtCheckCd.Value)
			arrParam(3) = ""			
			
			arrParam(4) = ""
			arrParam(5) = "어음번호"			
			
		    arrField(0) = "NOTE_NO"	
		    
		    arrHeader(0) = "어음번호"				    
	End Select				
		
	If iwhere = 0 Then	
		Dim iCalledAspName
	
		iCalledAspName = AskPRAspName("a3117ra1")
	
		' 권한관리 추가 
		arrParamAdo(5) = lgAuthBizAreaCd
		arrParamAdo(6) = lgInternalCd
		arrParamAdo(7) = lgSubInternalCd
		arrParamAdo(8) = lgAuthUsrID	
	
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a3117ra1", "X")
			IsOpenPop = False
			Exit Function
		End If

		arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParamAdo),_
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
				.txtAllcNo.focus
			Case 1	
				.txtBpCd.focus
			Case 2
				.txtBizCd.focus
			Case 3
				.txtDocCur.focus
			Case 4

			Case 5
				.txtBankCd.focus			    		
			Case 6
				.txtBankAcct.focus	
			Case 7	
				.txtCheckCd.focus		
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
				.txtAllcNo.value = arrRet(0)
				.txtAllcNo.focus
			Case 1	
				.txtBpCd.value = arrRet(0)		
				.txtBpNm.value = arrRet(1)
				.txtBpCd.focus
			Case 2
				.txtBizCd.value = arrRet(0)		
				.txtBizNm.value = arrRet(1)
				.txtBizCd.focus
			Case 3
				.txtDocCur.value = arrRet(0)		
				
				Call txtDocCur_OnChange()
				.txtDocCur.focus
			Case 4

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
		End Select				
	End With
	If iwhere  <> 0 Then
		lgBlnFlgChgValue = True
	End If	
End Function


'======================================================================================================
'	기능: Tab Click
'	설명: Tab Click시 필요한 기능을 수행한다.
'=======================================================================================================
Function ClickTab1()

	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)	 '~~~ 첫번째 Tab 
	gSelframeFlg = TAB1
	
	If lgIntFlgMode <> parent.OPMD_UMODE Then
	    Call SetToolbar("1110101100001111")										'⊙: 버튼 툴바 제어 
	Else				 
	    Call SetToolbar("1111101100001111")										'⊙: 버튼 툴바 제어 
	End If
	    
End Function

Function ClickTab2()

	If gSelframeFlg = TAB2 Then Exit Function
	Call changeTabs(TAB2)	 '~~~ 첫번째 Tab 
	gSelframeFlg = TAB2

End Function




'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.2 Common Group-2
' Description : This part declares 2nd common function group
'=======================================================================================================
'*******************************************************************************************************




'======================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 
'                 함수를 Call하는 부분 
'=======================================================================================================
Sub  Form_Load()
    Call LoadInfTB19029()																'Load table , B_numeric_format
    
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, _
							gDateFormat, parent.gComNum1000, parent.gComNumDec)    
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, _
							gDateFormat, parent.gComNum1000, parent.gComNumDec)
                         
    Call ggoOper.LockField(Document, "N")										'Lock  Suitable  Field
    Call InitSpreadSheet()																'Setup the Spread sheet
    Call InitVariables()																'Initializes local global variables
    Call SetDefaultVal()    
    
    Call SetToolbar("1110101100001111")										'버튼 툴바 제어	
	frm1.txtAllcNo.focus

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

'======================================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'=======================================================================================================
Sub  Form_QueryUnload(Cancel , UnloadMode )
    
End Sub

'======================================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'=======================================================================================================
Function  FncQuery() 
    Dim IntRetCD 
    Dim var1
    
    FncQuery = False                                                        
    
    Err.Clear                                                               
	'-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then													'This function check indispensable field
       Exit Function
    End If
    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    var1 = ggoSpread.SSCheckChange
    
    If lgBlnFlgChgValue = True Or var1 = True  Then		
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")	    
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'Clear Contents  Field
    Call InitVariables()																'Initializes local global variables
    frm1.vspdData.MaxRows = 0    
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery()																		'☜: Query db data
           
    FncQuery = True																
	
	Set gActiveElement = document.activeElement    
	
End Function

'=======================================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================================
Function  FncNew() 
    Dim IntRetCD 
	Dim var1
	    
    FncNew = False                                                          
    
    ggoSpread.Source = frm1.vspdData
    var1 = ggoSpread.SSCheckChange
	'-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Or var1 = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"X","X")               
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	'-----------------------
    'Erase condition area
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "1")                                         '⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                  'Clear Condition Field
    Call ggoOper.LockField(Document, "N")                                   'Lock  Suitable  Field
    Call InitVariables()
    Call SetDefaultVal()
	Call txtDocCur_OnChange()
    Call DisableRefPop()
    frm1.vspdData.MaxRows = 0    
    
    frm1.txtAllcNo.Value = ""
    frm1.txtAllcNo.focus

    lgBlnFlgChgValue = False    
    
	FncNew = True  
	
	Set gActiveElement = document.activeElement    
	
                                                      
End Function

'=======================================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================================
Function  FncDelete() 
    Dim IntRetCD
    
    FncDelete = False                                                      
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then										'Check if there is retrived data
        IntRetCD = DisplayMsgBox("900002","X","X","X")                               
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")					'Will you destory previous data"
	If IntRetCD = vbNo Then
		Exit Function
	End If
    '-----------------------
    'Delete function call area
    '-----------------------
    If DbDelete = False Then														'☜: Delete db data
		Exit Function																'☜:
    End If					
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------    													
    
    FncDelete = True                                                        
	
	Set gActiveElement = document.activeElement    
	
End Function

'=======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================================
Function  FncSave() 
    Dim IntRetCD 
	Dim var1
	
    FncSave = False                                                         
    
    Err.Clear                                                               
      
    ggoSpread.Source = frm1.vspdData
    var1 = ggoSpread.SSCheckChange

    If lgBlnFlgChgValue = False And var1 = False  Then								'⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","X","X","X")								'⊙: Display Message(There is no changed data.)
		Exit Function
    End If
  
  	If Not chkField(Document, "2") Then												'⊙: Check required field(Single area)
		Exit Function
    End If
    
	ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then										'⊙: Check contents area
		Exit Function
    End If

    If Not chkAllcDate() Then
		Exit Function
    End If  
    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave				                                             '☜: Save db data
    
    FncSave = True                                                       
	
	Set gActiveElement = document.activeElement    
	
End Function

'=======================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function  FncCopy() 
	Dim  IntRetCD
	 
	If frm1.vspdData.Maxrows < 1 Then Exit Function 
	
	frm1.vspdData.ReDraw = False
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")	'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	
	With frm1
		.vspdData.ReDraw = False
	
		ggoSpread.Source = frm1.vspdData	
		ggoSpread.CopyRow
		Call SetSpreadColor(frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow)
    
		.vspdData.ReDraw = True
	End With
	
	Set gActiveElement = document.activeElement    
		
End Function

'=======================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function  FncCancel() 
	Dim i
    if frm1.vspdData.Maxrows < 1 Then Exit Function

	With frm1.vspddata
	    .Row = .ActiveRow
	    .Col = 0
		    
	    ggoSpread.Source = frm1.vspddata
	    ggoSpread.EditUndo
		Call Dosum()
		
		If frm1.vspdData.MaxRows < 1 Then 
			Call ggoOper.SetReqAttr(frm1.txtALLCDt,   "N")
			Exit Function
		End if					

	    .Row = .ActiveRow
	    .Col = 0	
		
		For i = .MaxRows to 0 Step -1 
			.Row= i
			.Col =0			
			If Trim(frm1.vspddata1.text) = ggoSpread.InsertFlag Then 
				Call ggoOper.SetReqAttr(frm1.txtALLCDt,   "Q")
				Exit Function
			End if
				
			Call ggoOper.SetReqAttr(frm1.txtALLCDt,   "N")
		Next
			    	    
	End With   
	
	Set gActiveElement = document.activeElement    
	
End Function

'=======================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function  FncInsertRow() 

End Function

'=======================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function  FncDeleteRow() 
    Dim lDelRows 
	
	If frm1.vspdData.Maxrows < 1 Then Exit Function
	ggoSpread.Source = frm1.vspdData 
    lDelRows = ggoSpread.DeleteRow

	Call DoSum()
	
	Set gActiveElement = document.activeElement    
		
End Function

'=======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================================
Function  FncPrint() 
    On Error Resume Next                                               
    parent.FncPrint()
    	
	Set gActiveElement = document.activeElement    
	
End Function

'=======================================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================================
Function  FncPrev() 
    On Error Resume Next   
                                                
End Function

'=======================================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================================
Function  FncNext() 
    On Error Resume Next                                               
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
	Call parent.FncExport(parent.C_SINGLEMULTI)
		
	Set gActiveElement = document.activeElement    
	
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function FncSplitColumn()
	Dim ACol
	Dim ARow
	Dim iRet
	Dim iColumnLimit
	
	iColumnLimit = 5
	
	ACol = frm1.vspdData.ActiveCol
	ARow = frm1.vspdData.ActiveRow
	
	If ACol > iColumnLimit Then
		Frm1.vspdData.Col = iColumnLimit : frm1.vspdData.Row = 0  	 	 	 	 		
		  iRet = DisplayMsgBox("900030", "X", Trim(frm1.Vspddata.text), "X")
		Exit Function
	End If
	
	frm1.vspdData.ScrollBars = parent.SS_SCROLLBAR_NONE
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SSSetSplit(ACol)
	
	frm1.vspdData.Col = ACol
	frm1.vspdData.Row = ARow
	frm1.vspdData.Action = 0
	frm1.vspdData.ScrollBars = parent.SS_SCROLLBAR_BOTH
End Function

'=======================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function  FncExit()
	Dim IntRetCD
	Dim var1
	
	FncExit = False

	ggoSpread.Source = frm1.vspdData
    var1 = ggoSpread.SSCheckChange

    If lgBlnFlgChgValue = True or var1 = True Then  '⊙: Check If data is chaged
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




'=======================================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================================
Function  DbDelete() 
    DbDelete = False	
    													
    Call LayerShowHide(1)
    
    Dim strVal
    
    strVal = BIZ_PGM_DEL_ID & "?txtlgMode=" & parent.UID_M0003
    strVal = strVal & "&txtAllcNo=" & Trim(frm1.txtAllcNo.value)					'☜: 삭제 조건 데이타 
	' 권한관리 추가 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 

    
	Call RunMyBizASP(MyBizASP, strVal)												'☜: 비지니스 ASP 를 가동 
    
    DbDelete = True                                                         
End Function

'=======================================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================================
Function DbDeleteOk()																'삭제 성공후 실행 로직 
	Call ggoOper.ClearField(Document, "1")                                   '⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")									'Clear Condition Field
    Call ggoOper.LockField(Document, "N")									'Lock  Suitable  Field

    frm1.vspdData.MaxRows = 0    

    Call InitVariables																'Initializes local global variables
    Call SetDefaultVal
	Call DisableRefPop()        
    frm1.txtAllcNo.Value = ""
    frm1.txtAllcNo.focus
End Function

'=======================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================================
Function  DbQuery() 
    DbQuery = False                                                             
    Call LayerShowHide(1)
    
    Dim strVal
    
    With frm1
		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001					
			strVal = strVal & "&txtAllcNo=" & Trim(.htxtAllcNo.value)				'조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
			strVal = strVal & "&lgStrPrevKeyDtl=" & lgStrPrevKeyDtl
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001					
			strVal = strVal & "&txtAllcNo=" & Trim(.txtAllcNo.value)				'조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
			strVal = strVal & "&lgStrPrevKeyDtl=" & lgStrPrevKeyDtl
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If
    End With

	' 권한관리 추가 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 

	Call RunMyBizASP(MyBizASP, strVal)												'☜: 비지니스 ASP 를 가동 

    DbQuery = True                                                              
End Function

'=======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function
'=======================================================================================================
Function  DbQueryOk()
	If frm1.vspdData.MaxRows > 0 Then
		Call SetSpreadLock()  
	End If 

    Call ggoOper.LockField(Document, "Q")									'This function lock the suitable field
    Call SetToolbar("1111101100001111")										'버튼 툴바 제어        
	Call InitVariables()						 
	
	Call DoSum()
	Call txtDocCur_OnChange()	
	lgIntFlgMode = parent.OPMD_UMODE												'Indicates that current mode is Update mode
	Call DisableRefPop()
    lgBlnFlgChgValue = False    	
End Function

'=======================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================================
Function  DbSave() 
    Dim lngRows 
    Dim lGrpcnt
    Dim strVal 
    Dim strDel

    DbSave = False                                                          
    Call LayerShowHide(1)

    On Error Resume Next                                                   
	Err.Clear 

	frm1.txtFlgMode.value = lgIntFlgMode									
    '-----------------------
    'Data manipulate area
    '-----------------------
    ' Data 연결 규칙 
    ' 0: Sheet명, 1: Flag , 2: Row위치, 3~N: 각 데이타 

    lGrpCnt = 1
    strVal = ""
    strDel = ""
    
    ggoSpread.Source = frm1.vspdData

	With frm1.vspdData
		For lngRows = 1 To .MaxRows
		    .Row = lngRows
			.Col = 0
			Select Case .Text
				Case ggoSpread.DeleteFlag

				Case Else
					strVal = strVal & "C" & parent.gColSep  					'☜: C=Create, Row위치 정보 
			        .Col = C_ApNo	
			        strVal = strVal & Trim(.Text) & parent.gColSep
			        .Col = C_AcctCd
			        strVal = strVal & Trim(.Text) & parent.gColSep
			        .Col = C_ApDt
			        strVal = strVal & UniConvDate(Trim(.Text)) & parent.gColSep
			        .Col = C_DocCur
			        strVal = strVal & Trim(.Text) & parent.gColSep
			        .Col = C_ApClsAmt
			        strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gColSep
			        .Col = C_ApClsLocAmt		            
			        strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gColSep
			        .Col = C_ApClsDesc		            
			        strVal = strVal & Trim(.Text) & parent.gRowSep		        
			            
			        lGrpCnt = lGrpCnt + 1	
			End Select	
		Next
	End With	

	With frm1
		.txtMaxRows.value = lGrpCnt-1											'Spread Sheet의 변경된 최대갯수 
		.txtSpread.value =  strDel & strVal										'Spread Sheet 내용을 저장 

		'권한관리추가 start
		.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
		.txthInternalCd.value =  lgInternalCd
		.txthSubInternalCd.value = lgSubInternalCd
		.txthAuthUsrID.value = lgAuthUsrID		
		'권한관리추가 end
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)									'저장 비지니스 ASP 를 가동 
        
    DbSave = True                                                           
End Function

'=======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================================
Function  DbSaveOk(ByVal AllcNo)												'☆: 저장 성공후 실행 로직 
    ggoSpread.SSDeleteFlag 1
    
    If lgIntFlgMode = parent.OPMD_CMODE Then
		  frm1.txtAllcNo.value = AllcNo
	End If	  
	
	Call ggoOper.ClearField(Document, "2")								'Clear Contents  Field
    Call InitVariables															'Initializes local global variables
    frm1.vspdData.MaxRows = 0    
   
	Call DbQuery()
End Function

'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.4 User-defined Method Part
' Description : This part declares user-defined method
'=======================================================================================================
'*******************************************************************************************************

'===================================== DisableRefPop()  =======================================
'	Name : DisableRefPop()
'	Description :
'====================================================================================================
Sub DisableRefPop()
	IF lgIntFlgMode = parent.OPMD_UMODE Then
		RefPop.innerHTML="<font color=""#777777"">입금정보</font>"
	ELse 
		RefPop.innerHTML="<A href=""vbscript:OpenRefRcptNo()"">입금정보</A>"
	End if

END sub


'=======================================================================================================
' Function Name : chkAllcDate
' Function Desc : 
'========================================================================================================
Function chkAllcDate()
	Dim intI
	
	chkAllcDate = True
	With frm1
		For intI = 1 To .vspdData.Maxrows
			.vspdData.Row = intI
			.vspdData.Col = C_ApDt		

			If CompareDateByFormat(.vspdData.Text,.txtAllcDt.Text,"채무일자",.txtAllcDt.Alt, _
		    	               "970025",.txtAllcDt.UserDefinedFormat,parent.gComDateType, true) = False Then
			   .txtAllcDt.focus
			   chkAllcDate = False
			   Exit Function
			End If

			If CompareDateByFormat(.vspdData.Text,.txtRcptDt.Text,"채무일자",.txtRcptDt.Alt, _
		    	               "970025",.txtRcptDt.UserDefinedFormat,parent.gComDateType, true) = False Then
			   '.txtRcptDt.focus
			   chkAllcDate = False
			   Exit Function
			End If
		Next
	End With
End Function

'======================================================================================================
'   Name : DoSum()
'   Desc : Sum sheet Data
'=======================================================================================================
Sub DoSum()
	Dim dblToApAmt			'채무액	변수 
	Dim dblToApRemAmt		'채무잔액 변수 

	With frm1.vspdData
		dblToApAmt = FncSumSheet1(frm1.vspdData,C_ApAmt, 1, .MaxRows, False, -1, -1, "V")
		dblToApRemAmt = FncSumSheet1(frm1.vspdData,C_ApRemAmt, 1, .MaxRows, false, -1, -1, "V")
		
		If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
			frm1.txtTotApAmt.text	= UNIConvNumPCToCompanyByCurrency(dblToApAmt,frm1.txtDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
			frm1.txtTotApRemAmt.text	= UNIConvNumPCToCompanyByCurrency(dblToApRemAmt,frm1.txtDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")			
		End If	
	End With	
End Sub

'==========================================================================================
'   Event Name : txtDocCur_OnChange
'   Event Desc : 
'==========================================================================================
Sub txtDocCur_OnChange()
    lgBlnFlgChgValue = True
    If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
		Call CurFormatNumericOCX()
		Call CurFormatNumSprSheet()
	End If	    
End Sub

'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()
	With frm1
		' 입금잔액 
		ggoOper.FormatFieldByObjectOfCur .txtBalAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 반제금액 
		ggoOper.FormatFieldByObjectOfCur .txtClsAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 채무액 
		ggoOper.FormatFieldByObjectOfCur .txtTotApAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 채무잔액 
		ggoOper.FormatFieldByObjectOfCur .txtTotApRemAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
	End With
End Sub

'===================================== CurFormatNumSprSheet()  ======================================
'	Name : CurFormatNumSprSheet()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet()
'	With frm1
'		ggoSpread.Source = frm1.vspdData
'		' 채무액 
'		ggoSpread.SSSetFloatByCellOfCur C_ApAmt,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
'		' 채무잔액 
'		ggoSpread.SSSetFloatByCellOfCur C_ApRemAmt,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
'		' 반제금액 
'		ggoSpread.SSSetFloatByCellOfCur C_ApClsAmt,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
'	End With
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
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.RestoreSpreadInf()
	Call InitSpreadSheet()
	Call ggoSpread.ReOrderingSpreadData()
End Sub




'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.6 Spread OCX Tag Event
' Description : This part declares Spread OCX Tag Event
'=======================================================================================================
'*******************************************************************************************************




'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopUpMenuItemInf("0101111111")
    gMouseClickStatus = "SPC"							'Split 상태코드 
 
	Set gActiveSpdSheet = frm1.vspdData
	
	If frm1.vspdData.Maxrows = 0 then
	    Exit Sub
	End if

	If Row <= 0 Then
		Exit Sub
	End If		
End Sub

'======================================================================================================
'   Event Name :vspddata_EditChange
'   Event Desc :
'=======================================================================================================
Sub  vspdData_EditChange(ByVal Col , ByVal Row )

End Sub

'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'=======================================================================================================
Sub  vspdData_Change(ByVal Col, ByVal Row )
	Dim ApAmt
	Dim ClsAmt

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    
    frm1.vspdData.Row = Row
    frm1.vspdData.Col = 0   

    Select Case Col
		Case C_ApClsAmt
			frm1.vspdData.Col = C_ApAmt
			ApAmt = frm1.vspdData.Text
			frm1.vspdData.Col = C_ApClsAmt
			ClsAmt = frm1.vspdData.Text

			If (UNICDbl(ApAmt) > 0 And parent.UNICDbl(ClsAmt) < 0) Or (UNICDbl(ApAmt) < 0 And parent.UNICDbl(ClsAmt) > 0) Then
				frm1.vspdData.Col = C_ApClsAmt
				frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(UNICDbl(frm1.vspdData.Text) * (-1),frm1.txtDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
				
			End If
			frm1.vspdData.Col  = C_ApClsLocAmt			
			frm1.vspdData.text = "" 
	End Select
	
	Call DoSum()
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
'   Event Name :vspddata_DblClick
'   Event Desc :
'=======================================================================================================
Sub  vspddata_DblClick(ByVal Col,ByVal Row)
	If Row <=0 Then
		Exit Sub				
	End If		

    If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If
End Sub

'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'======================================================================================================
'   Event Name :vspddata_ScriptDragDropBlock
'   Event Desc :
'=======================================================================================================
Sub  vspddata_ScriptDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
	ggoSpread.Source = frm1.vspdData 
	Call ggoSpread.SpreadDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
	Call GetSpreadColumnPos("A")
End Sub

'======================================================================================================
'   Event Name :vspddata_KeyPress
'   Event Desc :
'=======================================================================================================
Sub  vspddata_KeyPress(KeyAscii )
     
End Sub




'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.7 Date-Numeric OCX Tag Event
' Description : This part declares HTML Tag Event
'=======================================================================================================
'*******************************************************************************************************




'=======================================================================================================
'   Event Name : txtAllcDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtAllcDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtAllcDt.Action = 7                        
        Call SetFocusToDocument("M")
		Frm1.txtAllcDt.Focus     
    End If
End Sub

'=======================================================================================================
'   Event Name : txtAllcDt_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtAllcDt_Change()
    lgBlnFlgChgValue = True
End Sub

'===================================== XchLocRate()  ======================================
'	Name : XchLocRate()
'	Description : 환율이 변경되는 Factor 가 변했을 때 수정되는 Local Amt. Setting
'====================================================================================================
Sub XchLocRate()
	Dim ii

	With frm1
		For ii = 1 To .vspdData.MaxRows 
			.vspdData.Row = ii	
			.vspdData.Col = C_ApClsLocAmt	
			.vspdData.Text = ""  
			ggoSpread.Source = .vspdData
			ggoSpread.UpdateRow ii			  		
		Next	
	End With
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  --> 
</HEAD>
<!--
 '#########################################################################################################
'       					6. Tag부 
'######################################################################################################### 
 -->

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD	WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 >
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/seltab_up_left.gif" width=9></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>(-)채무/입금반제</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/seltab_up_right.gif" width=10></td>
						    </TR>
						</TABLE>
					</TD>								
					<TD WIDTH=* align=right><A HREF="VBSCRIPT:OpenPopupTempGL()">결의전표</A>&nbsp;|&nbsp;<A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A>&nbsp;|&nbsp;<Span id="RefPop"><a href="vbscript:OpenRefRcptNo()">입금정보</A></Span>&nbsp;|&nbsp;<A href="vbscript:OpenRefOpenAp()">채무발생정보</A></TD>								
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
									<TD CLASS="TD5" NOWRAP>반제번호</TD>
									<TD CLASS="TD656" NOWRAP><INPUT NAME="txtAllcNo" ALT="반제번호" MAXLENGTH=18 STYLE="TEXT-ALIGN: Left" tag ="12XXXU"><IMG align=top name=btnCalType src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON" onclick="vbscript: Call OpenPopup(frm1.txtAllcNo.value,0)"></TD>								
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
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>입금번호</TD>
								<TD CLASS=TD6 NOWRAP>
								<INPUT TYPE="Text" NAME="txtRcptNo" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="24" ALT="입금번호"></TD>
								<TD CLASS=TD5 NOWRAP>입금일</TD>
								<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtRcptDt" CLASS=FPDTYYYYMMDD tag="24" Title="FPDATETIME" ALT="입금일" id=fpDateTime1></OBJECT>');</SCRIPT></TD>
							</TR>												
							<TR>
								<TD CLASS=TD5 NOWRAP>반제일</TD>
								<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtAllcDt" CLASS=FPDTYYYYMMDD tag="22" Title="FPDATETIME" ALT="반제일" id=fpDateTime1></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>거래처</TD>
								<TD CLASS=TD6 NOWRAP>
								<INPUT TYPE="Text" NAME="txtBpCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="24" ALT="거래처"> <INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24" ALT="거래처명"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>사업장</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBizCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag=24" ALT="사업장"> 
								<INPUT TYPE=TEXT NAME="txtBizNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="24" ALT="사업장명">
									<INPUT TYPE=hidden NAME="txtDeptCd"></TD>
								<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
								<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>결의전표번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTempGlNo" SIZE=19 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24XXXU" ALT="결의전표번호"> </TD>																						
								<TD CLASS="TD5" NOWRAP>전표번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtGlNo" SIZE=19 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24XXXU" ALT="전표번호"></TD>								
							</TR>	
							<TR>
								<TD CLASS=TD5 NOWRAP>거래통화</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDocCur" SIZE=10 MAXLENGTH=4 tag="24" STYLE="TEXT-ALIGN: left" ALT="거래통화"></TD>
								<TD CLASS=TD5 NOWRAP>환율</TD>
								<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtXchRate" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="환율" tag="24X5Z" ></OBJECT>');</SCRIPT></TD>											
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>입금잔액</TD>
								<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtBalAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="입금잔액" tag="24X2" ></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>입금잔액(자국통화)</TD>
								<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtBalLocAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="입금잔액(자국통화)" tag="24X2" ></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>반제금액</TD>
								<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtClsAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="반제금액" tag="24X2" ></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>반제금액(자국통화)</TD>
								<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtClsLocAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="반제금액(자국통화)" tag="24X2" ></OBJECT>');</SCRIPT></TD>
							</TR>	
							<TR>
								<TD CLASS=TD5 NOWRAP>비고</TD>
								<TD CLASS=TD656 NOWRAP COLSPAN=3><INPUT TYPE=TEXT NAME="txtDesc" SIZE=80 MAXLENGTH=128 tag="21XXX" ALT="비고"></TD>								
							</TR>																		
							<TR HEIGHT="100%">
								<TD WIDTH="100%" COLSPAN="4">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" name=vspdData width="100%" tag="2" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>											
							</TR>

							<TR>
								<TD  COLSPAN="4">
									<TABLE <%=LR_SPACE_TYPE_60%>>
										<TR>														
								<TD CLASS=TDT NOWRAP>채무액</TD>
								<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTotApAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="채무액" tag="24X2" ></OBJECT>');</SCRIPT></TD>
								<TD class=TDT STYLE="WIDTH : 0px;"></TD>
								<TD CLASS=TDT NOWRAP>채무잔액</TD>
								<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTotApRemAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="채무잔액" tag="24X2" ></OBJECT>');</SCRIPT></TD>								
										</TR>
									</TABLE>
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
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>	
</TABLE>
<TEXTAREA class=hidden name=txtSpread		tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA class=hidden name=txtSpread1		tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA class=hidden name=txtSpread2		tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA class=hidden name=txtSpread3		tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode"			tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtUpdtUserId"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtInsrtUserId"	tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows1"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtFlgMode"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="htxtAllcNo"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hItemSeq"			tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hAcctCd"			tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows3"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txthAuthBizAreaCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthSubInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthUsrID"		tag="24" Tabindex="-1">
<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT=0 name=vspdData3 width="100%" tag="2" TABINDEX="-1"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
