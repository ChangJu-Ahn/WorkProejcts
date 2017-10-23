<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>


<!--
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : bank Register
'*  3. Program ID           : a3111ma.asp
'*  4. Program Name         : 채권반제(가수금)
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
<!--'=======================================================================================================
'												1. 선 언 부 
'=======================================================================================================
'=======================================================================================================
'                                               1.1 Inc 선언   
'	기능: Inc. Include
'======================================================================================================= -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="javascript"   SRC="../../inc/TabScript.js">			</SCRIPT>
<SCRIPT LANGUAGE="VBScript"     SRC="../ag/AcctCtrl.vbs">				</SCRIPT>
<SCRIPT LANGUAGE=vbscript>

Option Explicit																	'☜: indicates that All variables must be declared in advance

'=======================================================================================================
'                                               1.2 Global 변수/상수 선언  
'	.Constant는 반드시 대문자 표기.
'	.변수 표준에 따름. prefix로 g를 사용함.
'	.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=======================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

Const BIZ_PGM_QRY_ID = "a3111mb1.asp"												'☆: Head Query 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID = "a3111mb2.asp"												'☆: Save 비지니스 로직 ASP명 
Const BIZ_PGM_DEL_ID = "a3111mb3.asp"

'Const COM_BIZ_EXCHRATE_ID = "../../inc/GetExchRate.asp"								'☆: 환율정보 비지니스 로직 ASP명 

Const TAB1 = 1																		'☜: Tab의 위치 
Const TAB2 = 2

Dim C_RcptNo 
Dim C_Rcpt_AcctCd 
Dim C_Rcpt_AcctNm 
Dim C_Rcpt_BizCd 
Dim C_Rcpt_BizNm 
Dim C_RcptDt 
Dim C_Rcpt_DocCur
Dim C_RcptAmt 
Dim C_BalAmt 
Dim C_AllcAmt 
Dim C_AllcLocAmt 
Dim C_AllcDesc 

Dim C_ArNo 
Dim C_Ar_AcctCd 
Dim C_Ar_AcctNm 
Dim C_Ar_BizCd 
Dim C_Ar_BizNm 
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

Dim C_ItemSeq 
Dim C_AcctCd 
Dim C_AcctPB 
Dim C_AcctNm 
Dim C_DcAmt 
Dim C_DcLocAmt 


Dim  lgStrPrevKey1
Dim  lgStrPrevKeyDtl
Dim  lgStrPrevKey2
Dim  lgStrPrevKey3

Dim  IsOpenPop	                'Popup
Dim  gSelframeFlg
Dim  lgVspdNo

Dim  lgCurrRow

Dim dtToday
dtToday = "<%=GetSvrDate%>"


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
			C_RcptNo      = 1
			C_Rcpt_AcctCd = 2
			C_Rcpt_AcctNm = 3
			C_Rcpt_BizCd  = 4
			C_Rcpt_BizNm  = 5
			C_RcptDt      = 6
			C_RcptAmt     = 7
			C_BalAmt      = 8
			C_AllcAmt     = 9
			C_AllcLocAmt  = 10
			C_AllcDesc    = 11
		Case "B"			
			C_ArNo        = 1
			C_Ar_AcctCd   = 2
			C_Ar_AcctNm   = 3
			C_Ar_BizCd    = 4
			C_Ar_BizNm    = 5
			C_ArDt        = 6
			C_ArDueDt     = 7
			C_Ar_DocCur   = 8
			C_ArAmt       = 9
			C_ArRemAmt    = 10
			C_ArClsAmt    = 11
			C_ArClsLocAmt = 12
			C_ArDcAmt     = 13
			C_ArDcLocAmt  = 14
			C_ArClsDesc   = 15
		Case "C"			
			C_ItemSeq     = 1
			C_AcctCd      = 2
			C_AcctPB      = 3
			C_AcctNm      = 4
			C_DcAmt       = 5
			C_DcLocAmt    = 6
	End Select
End Sub

'======================================================================================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=======================================================================================================
Sub  InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
        
    lgStrPrevKey = ""                            'initializes Previous Key
    lgStrPrevKey1 = ""
    lgStrPrevKeyDtl = 0                         'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
	frm1.hOrgChangeId.value= parent.gChangeOrgId
    lgSortKey  = 1
End Sub

'======================================================================================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=======================================================================================================
Sub  SetDefaultVal()
	frm1.txtAllcDt.text =  UniConvDateAToB(dtToday, parent.gServerDateFormat,gDateFormat)
	lgBlnFlgChgValue = False 
	frm1.txtDocCur.value= parent.gcurrency
	frm1.hArDocCur.value =parent.gcurrency
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
Sub  InitSpreadSheet(ByVal pvSpdNo)
    Call initSpreadPosVariables(pvSpdNo)
    
    Select Case UCase(Trim(pvSpdNo))
		Case "A"
			With frm1.vspdData1    
				ggoSpread.Source = frm1.vspdData1
				ggoSpread.SpreadInit "V20021128",,parent.gAllowDragDropSpread 

				.Redraw = False    
				
				.MaxCols = C_AllcDesc + 1 												'☜: 최대 Columns의 항상 1개 증가시킴 
				.Col = .MaxCols															'공통콘트롤 사용 Hidden Column
				.ColHidden = True    
				.MaxRows = 0
    
				Call GetSpreadColumnPos(pvSpdNo)
				    
				ggoSpread.SSSetEdit	 C_RcptNo     , "입금번호"      , 20, 3
				ggoSpread.SSSetEdit	 C_Rcpt_AcctCd, "계정코드"      , 20, 3    
				ggoSpread.SSSetEdit	 C_Rcpt_AcctNm, "계정코드명"    , 20, 3
				ggoSpread.SSSetEdit	 C_Rcpt_BizCd , "사업장"        , 10, 3    
				ggoSpread.SSSetEdit	 C_Rcpt_BizNm , "사업장명"      , 20, 3       
				ggoSpread.SSSetDate	 C_RcptDt     , "입금일자"      , 10, 2, gDateFormat     
				ggoSpread.SSSetFloat C_RcptAmt    , "입금액"        , 15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_BalAmt     , "입금잔액"      , 15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_AllcAmt    , "반제금액"      , 15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_AllcLocAmt , "반제금액(자국)", 15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetEdit	 C_AllcDesc   , "비고"          , 20, 3        
   
				.Redraw = True 
			End With		
		Case "B"
			With frm1.vspdData10    
				ggoSpread.Source = frm1.vspdData10
				ggoSpread.SpreadInit "V20021128",,parent.gAllowDragDropSpread    
   
				.Redraw = False    
				   
			    .MaxCols = C_ArClsDesc + 1 												'☜: 최대 Columns의 항상 1개 증가시킴 
			    .Col = .MaxCols															'공통콘트롤 사용 Hidden Column
			    .ColHidden = True    
				.MaxRows = 0

				Call GetSpreadColumnPos(pvSpdNo)
				
				ggoSpread.SSSetEdit	 C_ArNo       , "채권번호"      , 20, 3
				ggoSpread.SSSetEdit	 C_Ar_AcctCd  , "계정코드"      , 20, 3
				ggoSpread.SSSetEdit	 C_Ar_AcctNm  , "계정코드명"    , 20, 3
				ggoSpread.SSSetEdit	 C_Ar_BizCd   , "사업장"        , 10, 3    
				ggoSpread.SSSetEdit	 C_Ar_BizNm   , "사업장명"      , 20, 3   
				ggoSpread.SSSetDate	 C_ArDt       , "채권일자"      , 10, 2, gDateFormat 
				ggoSpread.SSSetDate	 C_ArDueDt    , "만기일자"      , 10, 2, gDateFormat    
				ggoSpread.SSSetEdit	 C_Ar_DocCur  , "거래통화"      , 10, 3				
				ggoSpread.SSSetFloat C_ArAmt      , "채권액"        , 15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_ArRemAmt   , "채권잔액"      , 15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_ArClsAmt   , "반제금액"      , 15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_ArClsLocAmt, "반제금액(자국)", 15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_ArDcAmt    , "할인금액"      , 15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_ArDcLocAmt , "할인금액(자국)", 15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetEdit	 C_ArClsDesc  , "비고"          , 20, 3       
    
				Call ggoSpread.SSSetColHidden(C_ArDueDt,C_ArDueDt,True)    
				
				.Redraw = True 
			End With		
		Case "C"
			With frm1.vspdData    
				ggoSpread.Source = frm1.vspdData
				ggoSpread.SpreadInit "V20021128",,parent.gAllowDragDropSpread    
   
				.Redraw = False        
    
				.MaxCols = C_DcLocAmt + 1 										'☜: 최대 Columns의 항상 1개 증가시킴 
				.Col = .MaxCols										'공통콘트롤 사용 Hidden Column
				.ColHidden = True       
				.MaxRows = 0		
    
				Call GetSpreadColumnPos(pvSpdNo)
			    Call AppendNumberPlace("6","3","0")
			    
				ggoSpread.SSSetFloat  C_ItemSeq , "NO"            , 6,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,2,,,"0","999"    
				ggoSpread.SSSetEdit	  C_AcctCd  , "계정코드"      ,20, ,,20, 2
				ggoSpread.SSSetButton C_AcctPB
				ggoSpread.SSSetEdit	  C_AcctNm  , "계정코드명"    ,50,,,20,2
				ggoSpread.SSSetFloat  C_DcAmt   , "할인금액"      ,20, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat  C_DcLocAmt, "할인금액(자국)",20, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				
				Call ggoSpread.MakePairsColumn(C_AcctCd,C_AcctPB)
				
				.Redraw = True            
			End With
	End Select			
    
    Call SetSpreadLock(pvSpdNo)
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub  SetSpreadLock(ByVal pvSpdNo)
    Dim objSpread

	With frm1    
		Select Case UCase(Trim(pvSpdNo))
			Case "A"
				ggoSpread.Source = .vspdData1
				.vspddata1.Redraw = False        
    
				ggoSpread.SpreadLock C_RcptNo     , -1, C_RcptNo
				ggoSpread.SpreadLock C_Rcpt_AcctCd, -1, C_Rcpt_AcctCd
				ggoSpread.SpreadLock C_Rcpt_AcctNm, -1, C_Rcpt_AcctNm
				ggoSpread.SpreadLock C_Rcpt_BizCd , -1, C_Rcpt_BizCd
				ggoSpread.SpreadLock C_Rcpt_BizNm , -1, C_Rcpt_BizNm
				ggoSpread.SpreadLock C_RcptDt     , -1, C_RcptDt
				ggoSpread.SpreadLock C_RcptAmt    , -1, C_RcptAmt
				ggoSpread.SpreadLock C_BalAmt     , -1, C_BalAmt	

				ggoSpread.SSSetRequired C_AllcAmt, -1, -1
					 
				.vspddata1.Redraw = True        	
			Case "B"	
				ggoSpread.Source = .vspdData10
				.vspddata10.Redraw = False            
    
				ggoSpread.SpreadLock C_ArNo       , -1, C_ArNo
				ggoSpread.SpreadLock C_Ar_AcctCd  , -1, C_Ar_AcctCd
				ggoSpread.SpreadLock C_Ar_AcctNm  , -1, C_Ar_AcctNm
				ggoSpread.SpreadLock C_Ar_BizCd   , -1, C_Ar_BizCd
				ggoSpread.SpreadLock C_Ar_BizNm   , -1, C_Ar_BizNm
				ggoSpread.SpreadLock C_ArDt       , -1, C_ArDt
				ggoSpread.SpreadLock C_ArDueDt    , -1, C_ArDueDt
				ggoSpread.SpreadLock C_Ar_DocCur  , -1, C_Ar_DocCur
				ggoSpread.SpreadLock C_ArAmt      , -1, C_ArAmt
				ggoSpread.SpreadLock C_ArRemAmt   , -1, C_ArRemAmt
	
				ggoSpread.SSSetRequired C_ArClsAmt, -1, -1
	
				.vspddata1.Redraw = True
			Case "C"			
				ggoSpread.Source = .vspdData
				.vspddata.Redraw = False        
    
				ggoSpread.SpreadLock C_ItemSeq, -1, C_ItemSeq
				ggoSpread.SpreadLock C_AcctCd , -1, C_AcctCd
				ggoSpread.SpreadLock C_AcctPB , -1, C_AcctPB
				ggoSpread.SpreadLock C_AcctNm , -1, C_AcctNm
			
				ggoSpread.SSSetRequired  C_DcAmt, -1, -1 		   
  
				.vspddata.Redraw = True
		End Select
    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
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
' Function Name : SetSpread2ColorAr()
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

			C_RcptNo      = iCurColumnPos(1)
			C_Rcpt_AcctCd = iCurColumnPos(2)
			C_Rcpt_AcctNm = iCurColumnPos(3)
			C_Rcpt_BizCd  = iCurColumnPos(4)
			C_Rcpt_BizNm  = iCurColumnPos(5)
			C_RcptDt      = iCurColumnPos(6)
			C_RcptAmt     = iCurColumnPos(7)
			C_BalAmt      = iCurColumnPos(8)
			C_AllcAmt     = iCurColumnPos(9)
			C_AllcLocAmt  = iCurColumnPos(10)
			C_AllcDesc    = iCurColumnPos(11)
		Case "B"
			ggoSpread.Source = frm1.vspdData10

			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)							
			
			C_ArNo        = iCurColumnPos(1)
			C_Ar_AcctCd   = iCurColumnPos(2)
			C_Ar_AcctNm   = iCurColumnPos(3)
			C_Ar_BizCd    = iCurColumnPos(4)
			C_Ar_BizNm    = iCurColumnPos(5)
			C_ArDt        = iCurColumnPos(6)
			C_ArDueDt     = iCurColumnPos(7)
			C_Ar_DocCur   = iCurColumnPos(8)			
			C_ArAmt       = iCurColumnPos(9)
			C_ArRemAmt    = iCurColumnPos(10)
			C_ArClsAmt    = iCurColumnPos(11)
			C_ArClsLocAmt = iCurColumnPos(12)
			C_ArDcAmt     = iCurColumnPos(13)
			C_ArDcLocAmt  = iCurColumnPos(14)
			C_ArClsDesc   = iCurColumnPos(15)
		Case "C"
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

'=========================================================================================================
'	Name : OpenRefOpenAr()
'	Description : Ref 화면을 call한다. 
'========================================================================================================= 
Function OpenRefOpenAr()
	Dim arrRet
	Dim arrParam(8)	
	Dim iCalledAspName
	
	iCalledAspName = AskPRAspName("a3106ra5")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a3106ra5", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	If gSelframeFlg <> TAB1 Then Exit Function		 	
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	If frm1.vspdData10.MaxRows = 0 Then frm1.hArDocCur.value = ""

	arrParam(0) = frm1.txtBpCd.value				' 검색조건이 있을경우 파라미터 
	arrParam(1) = frm1.txtBpNm.value			
	arrParam(2) = frm1.hArDocCur.value					
	arrParam(3) = "M"	
	arrParam(6) = frm1.txtAllcDt.text
    arrParam(7) = frm1.txtAllcDt.Alt
        				
	arrRet = window.showModalDialog(iCalledAspName, Array(window.Parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0, 0) = "" Then		
		Exit Function
	Else		
		Call SetRefOpenAr(arrRet)
	End If
End Function

'=========================================================================================================
'	Name : OpenRefRcptNo()
'	Description : Ref 화면을 call한다. 
'========================================================================================================= 
Function OpenRefRcptNo()
	Dim arrRet
	Dim arrParam(7)
	Dim iCalledAspName
	
	iCalledAspName = AskPRAspName("a3107ra1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a3107ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	If gSelframeFlg <> TAB1 Then Exit Function		 	
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	If frm1.vspdData1.MaxRows = 0 Then frm1.txtDocCur.value = ""

	arrParam(0) = frm1.txtBpCd.value				' 검색조건이 있을경우 파라미터 
	arrParam(1) = frm1.txtBpNm.value			
	arrParam(2) = frm1.txtDocCur.value	
	arrParam(3) = "M"	
	arrParam(6) = frm1.txtAllcDt.text
    arrParam(7) = frm1.txtAllcDt.Alt								

	arrRet = window.showModalDialog(iCalledAspName & "?PGM=" & gStrRequestMenuID, Array(window.Parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0,0) = "" Then		
		Exit Function
	Else		
		Call SetRefRcptNo(arrRet)
	End If
End Function


'=========================================================================================================
'	Name : OpenPopupGL()
'	Description : 
'=========================================================================================================
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
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.Parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function

'=========================================================================================================
'	Name : OpenPopuptempGL()
'	Description :
'=========================================================================================================
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
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.Parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
End Function

'=========================================================================================================
'	Name : SetRefRcptNo()
'	Description : OpenAp Popup에서 Return되는 값 setting
'=========================================================================================================
Function  SetRefRcptNo(Byval arrRet)
	Dim intRtnCnt, strData
	Dim TempRow, I
	Dim j	
	DIM X
	Dim sFindFg
	Dim tempamt		

	With frm1
		.vspdData1.focus		
		ggoSpread.Source = .vspdData1
		.vspdData1.ReDraw = False	
	
		TempRow = .vspdData1.MaxRows														'☜: 현재까지의 MaxRows

		For I = TempRow To TempRow + Ubound(arrRet, 1) 
			sFindFg	= "N"
			For x = 1 to TempRow
				.vspdData1.Row = x
				.vspdData1.Col = C_RcptNo				
				If .vspdData1.Text = arrRet(I - TempRow, 0) Then
					sFindFg	= "Y"
				End If
			Next			
			If 	sFindFg	= "N" Then
				.vspdData1.MaxRows = .vspdData1.MaxRows + 1
				.vspdData1.Row = I + 1				
				.vspdData1.Col = 0
				.vspdData1.Text = ggoSpread.InsertFlag
				
				.vspdData1.Col = C_RcptNo													
				.vspdData1.text = arrRet(I - TempRow, 0)												
				.vspdData1.Col = C_Rcpt_AcctCd 												
				.vspdData1.text = arrRet(I - TempRow, 1)												
				.vspdData1.Col = C_Rcpt_AcctNm												
				.vspdData1.text = arrRet(I - TempRow, 2)												
				.vspdData1.Col = C_Rcpt_BizCd  												
				.vspdData1.text = arrRet(I - TempRow, 3)												
				.vspdData1.Col = C_Rcpt_BizNm 												
				.vspdData1.text = arrRet(I - TempRow, 4)												
				.vspdData1.Col = C_RcptDt      												
				.vspdData1.text = arrRet(I - TempRow, 5)												
				.vspdData1.Col = C_RcptAmt     												
				.vspdData1.text = arrRet(I - TempRow, 6)												
				.vspdData1.Col = C_BalAmt      												
				.vspdData1.text = arrRet(I - TempRow, 7)
				.vspdData1.Col = C_AllcAmt      												
				.vspdData1.text = arrRet(I - TempRow, 7)
				.vspdData1.Col = C_AllcDesc
				.vspdData1.text = arrRet(I - TempRow, 13)											
			End If	
		Next	

		.txtBpCd.value   = arrRet(0, 9)
		.txtBpNm.value   = arrRet(0, 10)
		.txtDocCur.value = arrRet(0, 11)
		
		If Trim(.hArDocCur.value) = "" Then frm1.hArDocCur.value = Trim(arrRet(0, 11))
		
		'Call ggoOper.SetReqAttr(frm1.txtAllcDt,   "Q")	
		
		ggoSpread.SpreadUnlock   C_RcptNo     , TempRow + 1, C_Rcpt_AcctCd , .vspdData1.MaxRows				'⊙: Unlock 컬럼 
		ggoSpread.ssSetProtected C_RcptNo     , TempRow + 1, .vspdData1.MaxRows
		ggoSpread.ssSetProtected C_Rcpt_AcctCd, TempRow + 1, .vspdData1.MaxRows								'⊙: Protected
		ggoSpread.SSSetRequired  C_AllcAmt    , TempRow + 1, .vspdData1.MaxRows

		Call DoSum()   		
		Call txtDocCur_OnChange()
		.vspdData1.ReDraw = True
    End With
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
	Dim tempamt

	With frm1
		.vspdData10.focus		
		ggoSpread.Source = .vspdData10
		.vspdData10.ReDraw = False	
	
		TempRow = .vspdData10.MaxRows																	'☜: 현재까지의 MaxRows

		For I = TempRow to TempRow + Ubound(arrRet, 1) 
			sFindFg	= "N"
			For x = 1 to TempRow
				.vspdData10.Row = x
				.vspdData10.Col = C_ArNo				
				If .vspdData10.Text = arrRet(I - TempRow, 0) Then
					sFindFg	= "Y"
				End If
			Next
			If 	sFindFg	= "N" Then
				.vspdData10.MaxRows = .vspdData10.MaxRows + 1
				.vspdData10.Row = I + 1				
				.vspdData10.Col = 0
				.vspdData10.Text = ggoSpread.InsertFlag
				
				.vspdData10.Col = C_ArNo        											
				.vspdData10.text = arrRet(I - TempRow, 0)
				.vspdData10.Col = C_Ar_AcctCd           									
				.vspdData10.text = arrRet(I - TempRow, 1)				
				.vspdData10.Col = C_Ar_AcctNm           									
				.vspdData10.text = arrRet(I - TempRow, 2)				
				.vspdData10.Col = C_Ar_BizCd            									
				.vspdData10.text = arrRet(I - TempRow, 3)				
				.vspdData10.Col = C_Ar_BizNm            									
				.vspdData10.text = arrRet(I - TempRow, 4)				
				.vspdData10.Col = C_ArDt                									
				.vspdData10.text = arrRet(I - TempRow, 5)				
				.vspdData10.Col = C_ArDueDt   
				.vspdData10.text = arrRet(I - TempRow, 6)
				.vspdData10.Col = C_Ar_DocCur   
				.vspdData10.text = UCase(arrRet(I - TempRow, 14))
				.vspdData10.Col = C_ArAmt               									
				.vspdData10.text = arrRet(I - TempRow, 7)				
				.vspdData10.Col = C_ArRemAmt            									
				.vspdData10.text = arrRet(I - TempRow, 8)				
				.vspdData10.Col = C_ArClsAmt            									
				.vspdData10.text = arrRet(I - TempRow, 10)	
				.vspdData10.Col = C_ArClsDesc
				.vspdData10.text = arrRet(I - TempRow, 13)
			End If		
		Next	

		If Trim(.txtBpCd.Value) = "" Then
			.txtbpCd.Value = arrRet(0, 11)	
			.txtbpNm.Value = arrRet(0, 12)			
		End If
		
		.hArDocCur.Value = arrRet(0, 14)
		.txtDocCur.value = arrRet(0, 14) '20051201 추가 
		
'		If .txtDocCur.value = "" Then .txtDocCur.value = UCase(arrRet(0, 13))
		
		'Call ggoOper.SetReqAttr(frm1.txtALLCDt,   "Q")	
		
		ggoSpread.SpreadUnlock   C_ArNo     , TempRow + 1, C_Ar_AcctCd, .vspdData10.MaxRows				'⊙: Unlock 컬럼 
		ggoSpread.ssSetProtected C_ArNo     , TempRow + 1, .vspdData10.MaxRows
		ggoSpread.ssSetProtected C_Ar_AcctCd, TempRow + 1, .vspdData10.MaxRows							'⊙: Protected
		ggoSpread.SSSetRequired  C_ArClsAmt , TempRow + 1, .vspdData10.MaxRows

		Call DoSum()   		
		Call txtDocCur_OnChange()
		.vspdData10.ReDraw = True
    End With
End Function

'======================================================================================================
'   Function Name : OpenPopUp()
'   Function Desc : 
'=======================================================================================================
Function  OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function
	
	Select Case iWhere
		Case 0

		Case 1
			If frm1.txtBpCd.className = "protected" Then Exit Function
			arrParam(0) = "거래처팝업"						' 팝업 명칭 
			arrParam(1) = "b_biz_partner"						' TABLE 명칭 
			arrParam(2) = strCode						 	    ' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = ""									' Where Condition
			arrParam(5) = "거래처"			
	
			arrField(0) = "BP_CD"								' Field명(0)
			arrField(1) = "BP_NM"								' Field명(1)
    
			arrHeader(0) = "거래처"							' Header명(0)
			arrHeader(1) = "거래처명"						' Header명(1)			
		Case 3
			If frm1.txtDocCur.className = "protected" Then Exit Function
			arrParam(0) = "거래통화팝업"					' 팝업 명칭 
			arrParam(1) = "b_currency"							' TABLE 명칭 
			arrParam(2) = strCode						 	    ' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = ""									' Where Condition
			arrParam(5) = "거래통화" 			
	
			arrField(0) = "CURRENCY"							' Field명(0)
			arrField(1) = "CURRENCY_DESC"						' Field명(1)
    
			arrHeader(0) = "거래통화"						' Header명(0)
			arrHeader(1) = "거래통화명"						' Header명(1)
		Case 4
			arrParam(0) = "계정코드팝업"								' 팝업 명칭 
			arrParam(1) = "A_Acct, A_ACCT_GP" 								' TABLE 명칭 
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
	End Select				
	
	IsOpenPop = True
	
	If iwhere = 0 Then	
		iCalledAspName = AskPRAspName("a3111ra1")
	
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a3111ra1", "X")
			IsOpenPop = False
			Exit Function
		End If
					
		arrRet = window.showModalDialog(iCalledAspName,array(window.parent,arrParam), _
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
			Case 3
				.txtDocCur.focus
			Case 4
				Call SetActiveCell(frm1.vspdData,C_AcctCd,frm1.vspdData.ActiveRow ,"M","X","X")
		End Select				
	End With
	IF iwhere  <> 0 Then
		lgBlnFlgChgValue = True
	End if	
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
			Case 3
				.txtDocCur.value = arrRet(0)		
				
				Call txtDocCur_OnChange()									' insert by Kim Sang Joong
				.txtDocCur.focus
			Case 4
				.vspdData.Col = C_AcctCd
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_AcctNm
				.vspdData.Text = arrRet(1)
			
				Call vspdData_Change(C_AcctCd, frm1.vspdData.activerow )	 ' 변경이 읽어났다고 알려줌 
				Call SetActiveCell(frm1.vspdData,C_AcctCd,frm1.vspdData.ActiveRow ,"M","X","X")
		End Select				
	End With
	IF iwhere  <> 0 Then
		lgBlnFlgChgValue = True
	End if	
End Function
'------------------------------------------  OpenDept()  ---------------------------------------
'	Name : OpenDept()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 

Function OpenDept(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(3)

	If IsOpenPop = True Then Exit Function
	If frm1.txtDeptCd.className = "protected" Then Exit Function
			
	IsOpenPop = True

	arrParam(0) = strCode		            '  Code Condition
   	arrParam(1) = frm1.txtAllcDt.Text
	arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  
	arrParam(3) = "F"									' 결의일자 상태 Condition  
	
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
				.txtAllcDt.text = arrRet(3)
				call txtDeptCd_OnBlur()  
				.txtDeptCd.focus
	    End Select
	End With
End Function 

'======================================================================================================
'	기능: Tab Click
'	설명: Tab Click시 필요한 기능을 수행한다.
'=======================================================================================================
Function ClickTab1()
	If lgIntFlgMode <> parent.OPMD_UMODE Then
	    Call SetToolBar("1110101100001111")										'⊙: 버튼 툴바 제어 
	Else    
	    Call SetToolBar("1111101100001111")										'⊙: 버튼 툴바 제어 
	End If
	
	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)	 '~~~ 첫번째 Tab 
	gSelframeFlg = TAB1
End Function

Function ClickTab2()
	If gSelframeFlg = TAB2 Then Exit Function
	Call changeTabs(TAB2)	 '~~~ 첫번째 Tab 
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



'======================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 
'                 함수를 Call하는 부분 
'=======================================================================================================
Sub  Form_Load()
    Call LoadInfTB19029()                                                         'Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, _
							gDateFormat, parent.gComNum1000, parent.gComNumDec)    
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, _
							gDateFormat, parent.gComNum1000, parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                  'Lock  Suitable  Field
    
    Call InitSpreadSheet("A")                                                     'Setup the Spread sheet
    Call InitSpreadSheet("B")                                                     'Setup the Spread sheet
    Call InitSpreadSheet("C")                                                     'Setup the Spread sheet        
	Call InitCtrlSpread()
	Call InitCtrlHSpread()	
    Call InitVariables()                                                          'Initializes local global variables
    Call SetDefaultVal()
    Call ClickTab1()
    
    frm1.txtAllcNo.focus
    
	gIsTab     = "Y" 
	gTabMaxCnt = 2  
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
    Dim var1, var2,var3, var4
    
    FncQuery = False                                                        
    
    Err.Clear                                                               
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'This function check indispensable field
       Exit Function
    End If
    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData1
    var1 = ggoSpread.SSCheckChange
    ggoSpread.Source = frm1.vspdData10
    var2 = ggoSpread.SSCheckChange
    ggoSpread.Source = frm1.vspdData
    var3 = ggoSpread.SSCheckChange
    ggoSpread.Source = frm1.vspdData2
    var4 = ggoSpread.SSCheckChange
    
    If lgBlnFlgChgValue = True Or var1 = True Or var2 = True Or var3 = True Or var4 = True Then		
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")	    
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ClickTab1()
    Call ggoOper.ClearField(Document, "2")						'Clear Contents  Field
    Call ggoOper.LockField(Document, "N") 
    Call InitVariables()												'Initializes local global variables
    																
	ggoSpread.Source = frm1.vspdData	:	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData1	:	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData10	:	ggoSpread.ClearSpreadData	
	ggoSpread.Source = frm1.vspdData2	:	ggoSpread.ClearSpreadData		
	ggoSpread.Source = frm1.vspdData3	:	ggoSpread.ClearSpreadData
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery()														'☜: Query db data
           
    FncQuery = True																
	
	Set gActiveElement = document.activeElement    
	
End Function

'=======================================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================================
Function  FncNew() 
    Dim IntRetCD 
	Dim var1, var2, var3, var4
	    
    FncNew = False                                                          
    
    ggoSpread.Source = frm1.vspdData1
    var1 = ggoSpread.SSCheckChange
    ggoSpread.Source = frm1.vspdData10
    var2 = ggoSpread.SSCheckChange
	ggoSpread.Source = frm1.vspdData
    var3 = ggoSpread.SSCheckChange
	ggoSpread.Source = frm1.vspdData2
    var4 = ggoSpread.SSCheckChange
	'-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Or var1 = True Or var2 = True Or var3 = True Or var4 = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"X","X")               
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	'-----------------------
    'Erase condition area
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "1")									'Clear Condition Field
    Call ggoOper.ClearField(Document, "2")									'Clear Condition Field
    Call ggoOper.LockField(Document, "N")									'Lock  Suitable  Field
    Call InitVariables()															'Initializes local global variables
    Call SetDefaultVal()
    Call txtDocCur_OnChange()														' insert by Kim Sang Joong
		                                                   
	ggoSpread.Source = frm1.vspdData	:	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData1	:	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData10	:	ggoSpread.ClearSpreadData	
	ggoSpread.Source = frm1.vspdData2	:	ggoSpread.ClearSpreadData		
	ggoSpread.Source = frm1.vspdData3	:	ggoSpread.ClearSpreadData
    
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
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")		            'Will you destory previous data"
	If IntRetCD = vbNo Then
		Exit Function
	End If
    '-----------------------
    'Delete function call area
    '-----------------------
    Call DbDelete() 																'☜: Delete db data
    
    FncDelete = True   
    	
	Set gActiveElement = document.activeElement    
	                                                    
End Function

'=======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================================
Function  FncSave() 
    Dim IntRetCD 
	Dim var1,var2, var3, var4
	
    FncSave = False                                                         
    
    Err.Clear                                                               
	    
    ggoSpread.Source = frm1.vspdData1
    var1 = ggoSpread.SSCheckChange

    ggoSpread.Source = frm1.vspdData10
    var2 = ggoSpread.SSCheckChange

    ggoSpread.Source = frm1.vspdData
    var3 = ggoSpread.SSCheckChange

    ggoSpread.Source = frm1.vspdData2
    var4 = ggoSpread.SSCheckChange

    If lgBlnFlgChgValue = False And var1 = False And var2 = False And var3 = False And var4 = False Then  '⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","X","X","X")								'⊙: Display Message(There is no changed data.)
		Exit Function		
    End If
    
    If Not chkField(Document, "2") Then												'⊙: Check required field(Single area)
       Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData1
    If Not ggoSpread.SSDefaultCheck Then	
		Call ClickTab1()									'⊙: Check contents area
       Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData10
    If Not ggoSpread.SSDefaultCheck Then
		Call ClickTab1()										'⊙: Check contents area
       Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then										'⊙: Check contents area
       Call ClickTab2()
       Exit Function
    End If

	If Not chkAllcDate() Then
		Exit Function
    End If
    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave()																	'☜: Save db data
    
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
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")				'⊙: "Will you destory previous data"
		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	
	With frm1
		.vspdData1.ReDraw = False
	
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
   If gSelframeFlg = TAB1 Then
   
		If lgVspdNo = 1 Then
		
			If frm1.vspdData1.Maxrows < 1 Then 	Exit Function
			
			With frm1.vspdData1
				.Row = .ActiveRow
				.Col = 0
			    ggoSpread.Source = frm1.vspdData1
			    ggoSpread.EditUndo
			    Call DoSum()

			End With   
		ELse
			If frm1.vspdData10.Maxrows < 1 Then Exit Function
			
			With frm1.vspdData10
				.Row = .ActiveRow
				.Col = 0
			    ggoSpread.Source = frm1.vspdData10
			    ggoSpread.EditUndo
			    Call DoSum()
			End With   
		End If
		
		If frm1.vspdData10.Maxrows < 1  AND frm1.vspdData1.Maxrows < 1 Then 
			Call ggoOper.SetReqAttr(frm1.txtALLCDt,   "N")
		End if	
	Else
		If frm1.vspdData.Maxrows < 1 Then Exit Function
		
		With frm1.vspdData
		    .Row = .ActiveRow
		    .Col = 0	    
		    
		    If .Text = ggoSpread.InsertFlag Then
		        .Col = C_AcctCd
				If Len(Trim(.text)) > 0 Then  
					.Col = C_ItemSeq		        
					DeleteHSheet(.Text)
				End IF
		    End if
   
		    ggoSpread.Source = frm1.vspdData	
		    ggoSpread.EditUndo
			
			If frm1.vspdData.Maxrows < 1 Then Exit Function
			
		    .Row = .ActiveRow
		    .Col = 0		    

			If .Row = 0  Then Exit Function
			
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

'=======================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function  FncInsertRow(ByVal pvRowCnt) 
	Dim imRow
	Dim ii
	Dim iCurRowPos
	
    On Error Resume Next															'☜: If process fails
    Err.Clear																		'☜: Clear error stat	

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
			Call MaxSpreadVal(frm1.vspdData, C_ItemSeq, ii)
		Next  
		.Col = 1																	' 컬럼의 절대 위치로 이동      
		.Row = 	ii - 1
		.Action = 0

        Call SetSpreadColor(iCurRowPos + 1, iCurRowPos + imRow)
        .ReDraw = True
    End With
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
    ggoSpread.Source = frm1.vspdData2										
	ggoSpread.ClearSpreadData		
	
    Call ggoOper.LockField(Document, "Q")                              'This function lock the suitable field

End Function

'=======================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function  FncDeleteRow() 
    Dim lDelRows 
	If gSelframeFlg = TAB1 Then
		IF lgVspdNo = 1 Then
			If frm1.vspdData1.Maxrows < 1 Then Exit Function
			ggoSpread.Source = frm1.vspdData1
		Else
			If frm1.vspdData10.Maxrows < 1 Then Exit Function
			ggoSpread.Source = frm1.vspdData10
		End If	
	Else
		If frm1.vspdData.Maxrows < 1 Then Exit Function
		ggoSpread.Source = frm1.vspdData
	End If	
	
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
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
End Sub

'=======================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function  FncExit()
	Dim IntRetCD
	Dim var1,var2, var3, var4
	
	FncExit = False

	ggoSpread.Source = frm1.vspdData1
    var1 = ggoSpread.SSCheckChange

    ggoSpread.Source = frm1.vspdData10
    var2 = ggoSpread.SSCheckChange

    ggoSpread.Source = frm1.vspdData
    var3 = ggoSpread.SSCheckChange

    ggoSpread.Source = frm1.vspdData2
    var4 = ggoSpread.SSCheckChange
    
    If lgBlnFlgChgValue = True or var1 = True or var2 = True or var3 = True or var4 = True Then  '⊙: Check If data is chaged
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
    
    Dim strVal
    
    strVal = BIZ_PGM_DEL_ID & "?txtMode=" & parent.UID_M0003
    strVal = strVal & "&txtAllcNo=" & Trim(frm1.txtAllcNo.value)			'☜: 삭제 조건 데이타 
    
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    
    DbDelete = True                                                         
End Function

'=======================================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================================
Function DbDeleteOk()												        '삭제 성공후 실행 로직 
	Call ggoOper.ClearField(Document, "1")                           'Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                           'Clear Condition Field
    Call ggoOper.LockField(Document, "N")                            'Lock  Suitable  Field

	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData1
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData10
	ggoSpread.ClearSpreadData	
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData		
	ggoSpread.Source = frm1.vspdData3
	ggoSpread.ClearSpreadData

    Call InitVariables()                                                     'Initializes local global variables
    Call Clicktab1()    
    Call SetDefaultVal()    
    
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
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001					'☜: 
			strVal = strVal & "&txtAllcNo=" & Trim(.htxtAllcNo.value)				'조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
			strVal = strVal & "&lgStrPrevKeyDtl=" & lgStrPrevKeyDtl
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001					'☜: 
			strVal = strVal & "&txtAllcNo=" & Trim(.txtAllcNo.value)				'조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
			strVal = strVal & "&lgStrPrevKeyDtl=" & lgStrPrevKeyDtl
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If
    End With

	Call RunMyBizASP(MyBizASP, strVal)										    '☜: 비지니스 ASP 를 가동 

    DbQuery = True                                                              
End Function

'=======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function
'=======================================================================================================
Function  DbQueryOk()
	With frm1
	    Call SetSpreadLock("A")
	    Call SetSpreadLock("B") 	     
	    Call SetSpreadLock("C") 	    
        '-----------------------
        'Reset variables area
        '-----------------------
        If .vspdData.MaxRows > 0 Then
            .vspdData.Row = 1
            .vspdData.Col = C_ItemSeq
            .hItemSeq.Value = .vspdData.Text 
            Call DbQuery2(1)
        End If
    
    End With    
    
    Call ggoOper.LockField(Document, "Q")									'This function lock the suitable field    
    lgIntFlgMode = parent.OPMD_UMODE												'Indicates that current mode is Update mode
    Call ClickTab1()
    Call DoSum()
    Call txtDocCur_OnChange()
    Call txtDeptCd_OnBlur()
    lgBlnFlgChgValue = False    
End Function

'=======================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================================
Function  DbSave() 
    Dim pAP010M 
    Dim lngRows 
    Dim lGrpcnt
    DIM strVal 
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
    
    ggoSpread.Source = frm1.vspdData1
	With frm1.vspdData1
		For lngRows = 1 To .MaxRows
		    .Row = lngRows
			.Col = 0
				
			Select Case .Text
				Case ggoSpread.DeleteFlag

				Case Else
					.Row = lngRows
					.Col = 0

					strVal = strVal & "C" & parent.gColSep & lngRows & parent.gColSep				'C=Create, Sheet가 2개 이므로 구별 
					.Col = C_RcptNo	'1
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = C_RcptDt		'2
					strVal = strVal & Trim(UniConvDate(.Text)) & parent.gColSep
					.Col = C_AllcAmt		'3
					strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gColSep
					.Col = C_AllcLocAmt		'4
					strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gColSep	
					.Col = C_AllcDesc		'5
					strVal = strVal & Trim(.Text) & parent.gRowSep					
					        
					lGrpCnt = lGrpCnt + 1
			End Select				
		Next
	End With	
	
	frm1.txtMaxRows.value = lGrpCnt-1										'Spread Sheet의 변경된 최대갯수 
	frm1.txtSpread.value =  strDel & strVal									'Spread Sheet 내용을 저장 

	lGrpCnt = 1
    strVal = ""
    strDel = ""
    
	ggoSpread.Source = frm1.vspdData10
	With frm1.vspdData10
		For lngRows = 1 To .MaxRows
    
			.Row = lngRows
			.Col = 0
				
			Select Case .Text
				Case ggoSpread.DeleteFlag

				Case Else	
					strVal = strVal & "C" & parent.gColSep & lngRows & parent.gColSep				'C=Create, Sheet가 2개 이므로 구별 
					
					.Col = C_ArNo				'2
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = C_ArDt				'3
					strVal = strVal & Trim(UniConvDate(.Text)) & parent.gColSep
					.Col = C_Ar_AcctCd			'4
					strVal = strVal & Trim(.Text) & parent.gColSep				        
					.Col = C_Ar_DocCur			'5
					strVal = strVal & Trim(.Text) & parent.gColSep				        
					.Col = C_ArClsAmt			'6
					strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gColSep	
					.Col = C_ArClsLocAmt		'7
					strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gColSep	
					.Col = C_ArDcAmt			'8
					strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gColSep	
					.Col = C_ArDcLocAmt			'9
					strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gColSep
					.Col = C_ArClsDesc			'10
					strVal = strVal & Trim(.Text) & parent.gRowSep				        
									
					lGrpCnt = lGrpCnt + 1
			End Select		        
			
		Next
	End With
	
	frm1.txtMaxRows0.value = lGrpCnt-1										'Spread Sheet의 변경된 최대갯수 
	frm1.txtSpread1.value =  strDel & strVal									'Spread Sheet 내용을 저장 

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
					strVal = strVal & "C" & parent.gColSep & lngRows & parent.gColSep				'C=Create, Sheet가 2개 이므로 구별 
			
					.Col = C_ItemSeq	'1
					strVal = strVal & Trim(.Text) & parent.gColSep
					            
					.Col = C_DcAmt		'2
					strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gColSep

					.Col = C_DcLocAmt		'3
					strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gColSep
					        
					.Col = C_AcctCd		'4
					strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gRowSep	
					        
					lGrpCnt = lGrpCnt + 1
			End Select							        
		Next
	End With
	
	frm1.txtMaxRows1.value = lGrpCnt-1														'Spread Sheet의 변경된 최대갯수 
	frm1.txtSpread2.value =  strDel & strVal												'Spread Sheet 내용을 저장    
				
    lGrpCnt = 1
    strVal = ""
    strDel = ""    

    With frm1.vspdData3   
		For lngRows = 1 To .MaxRows
			.Row = lngRows
			.Col = 0
				
			Select Case .Text
				Case ggoSpread.DeleteFlag

				Case Else
					strVal = strVal & "C" & parent.gColSep & lngRows & parent.gColSep		'C=Create, Sheet가 2개 이므로 구별 

					.Col = 1
					strVal = strVal & Trim(.Text) & parent.gColSep
					            
					.Col =  2
					strVal = strVal & Trim(.Text) & parent.gColSep

					.Col =  3
					strVal = strVal & Trim(.Text) & parent.gColSep
					        
					.Col = 5
					strVal = strVal & Trim(.Text) & parent.gRowSep  
   					        
					lGrpCnt = lGrpCnt + 1		
			End Select
		Next
	End With
	
    frm1.txtMaxRows3.value = lGrpCnt-1														'Spread Sheet의 변경된 최대갯수 
	frm1.txtSpread3.value =  strDel & strVal												'Spread Sheet 내용을 저장 
	
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)												'저장 비지니스 ASP 를 가동 
        
    DbSave = True                                                           
End Function

'=======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================================
Function  DbSaveOk(ByVal AllcNo)															'☆: 저장 성공후 실행 로직 
    ggoSpread.SSDeleteFlag 1
    
    If lgIntFlgMode = parent.OPMD_CMODE Then
		  frm1.txtAllcNo.value = AllcNo
	End If	  
	
	frm1.txtAllcNo.focus
	Call ClickTab1()
    Call ggoOper.ClearField(Document, "2")											'Clear Contents  Field
    Call ggoOper.LockField(Document, "N") 
    Call InitVariables()																	'Initializes local global variables
    
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData1
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData10
	ggoSpread.ClearSpreadData	
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData		
	ggoSpread.Source = frm1.vspdData3
	ggoSpread.ClearSpreadData		    
    Call DBQuery()
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
		strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD,'')),CHAR(8) "
    		
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
			
			For lngRows = 1 To frm1.vspdData2.Maxrows
				frm1.vspddata2.Row = lngRows	
				frm1.vspddata2.Col = C_Tableid 
				If Trim(frm1.vspddata2.text) <> "" Then
					frm1.vspddata2.Col = C_Tableid
					strTableid = frm1.vspddata2.text
					frm1.vspddata2.Col = C_Colid
					strColid = frm1.vspddata2.text
					frm1.vspddata2.Col = C_ColNm
					strColNm = frm1.vspddata2.text	
					frm1.vspddata2.Col = C_MajorCd					
					strMajorCd = frm1.vspddata2.text	
					
					frm1.vspddata2.Col = C_CtrlVal

					strNmwhere = strColid & " =   " & FilterVar(frm1.vspddata2.text, "''", "S") & "  " 
					
					If Trim(strMajorCd) <> "" Then
						strNmwhere = strNmwhere & " AND MAJOR_CD =  " & FilterVar(strMajorCd, "''", "S") & "  "
					End If				 
					
					If CommonQueryRs( strColNm , strTableid ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   
						frm1.vspddata2.Col = C_CtrlValNm
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
    
	lgBlnFlgChgValue = False            
End Function

'=======================================================================================================
' Function Name : chkAllcDate()
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function chkAllcDate()
	Dim intI
	
	chkAllcDate = True
	With frm1
		For intI = 1 To .vspdData10.Maxrows
			.vspdData10.Row = intI
			.vspdData10.Col = C_ArDt

			If CompareDateByFormat(.vspdData10.Text,.txtAllcDt.Text,"채권일자",.txtAllcDt.Alt, _
		    	               "970025",.txtAllcDt.UserDefinedFormat,parent.gComDateType, true) = False Then
			   .txtAllcDt.focus
			   chkAllcDate = False
			   Exit Function
			End If
		Next
	End With
End Function

'=======================================================================================================
' Function Name : DoSum()																			
' Function Desc : Sum Sheet Data
'=======================================================================================================
Sub DoSum()
	Dim dblTotRcptAmt
	Dim dblTotBalAmt
	Dim dblTotAllcAmt
	Dim dblTotAllcLocAmt
	
	Dim dblTotArAmt
	Dim dblTotArRemAmt
	Dim dblTotArClsAmt
	Dim dblTotArClsLocAmt
	Dim dblTotArDcAmt
	Dim dblTotArDcLocAmt
	
	
	With frm1.vspdData1
		dblTotRcptAmt    = FncSumSheet1(frm1.vspdData1,C_RcptAmt   , 1, .MaxRows, False, -1, -1, "V")
		dblTotBalAmt     = FncSumSheet1(frm1.vspdData1,C_BalAmt    , 1, .MaxRows, False, -1, -1, "V")
		dblTotAllcAmt    = FncSumSheet1(frm1.vspdData1,C_AllcAmt   , 1, .MaxRows, False, -1, -1, "V")
		dblTotAllcLocAmt = FncSumSheet1(frm1.vspdData1,C_AllcLocAmt, 1, .MaxRows, False, -1, -1, "V")
		
		If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
			frm1.txtTotRcptAmt.text	= UNIConvNumPCToCompanyByCurrency(dblTotRcptAmt,frm1.txtDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
			frm1.txtTotBalAmt.text	= UNIConvNumPCToCompanyByCurrency(dblTotBalAmt,frm1.txtDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
			frm1.txtTotAllcAmt.text	= UNIConvNumPCToCompanyByCurrency(dblTotAllcAmt,frm1.txtDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
		End If	
		frm1.txtTotAllcLocAmt.text = UNIConvNumPCToCompanyByCurrency(dblTotAllcLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")				
	End With

	With frm1.vspdData10
		dblTotArAmt = FncSumSheet1(frm1.vspdData10,C_ArAmt, 1, .MaxRows, False, -1, -1, "V")
		dblTotArRemAmt = FncSumSheet1(frm1.vspdData10,C_ArRemAmt, 1, .MaxRows, False, -1, -1, "V")
		dblTotArClsAmt = FncSumSheet1(frm1.vspdData10,C_ArClsAmt, 1, .MaxRows, False, -1, -1, "V")
		dblTotArClsLocAmt = FncSumSheet1(frm1.vspdData10,C_ArClsLocAmt, 1, .MaxRows, False, -1, -1, "V")
		dblTotArDcAmt = FncSumSheet1(frm1.vspdData10,C_ArDcAmt, 1, .MaxRows, False, -1, -1, "V")
		dblTotArDcLocAmt = FncSumSheet1(frm1.vspdData10,C_ArDcLocAmt, 1, .MaxRows, False, -1, -1, "V")
		
		If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.hArDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
			frm1.txtTotArAmt.text	 = UNIConvNumPCToCompanyByCurrency(dblTotArAmt,frm1.hArDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
			frm1.txtTotArRemAmt.text = UNIConvNumPCToCompanyByCurrency(dblTotArRemAmt,frm1.hArDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
			frm1.txtTotArClsAmt.text = UNIConvNumPCToCompanyByCurrency(dblTotArClsAmt,frm1.hArDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
			frm1.txtTotArDcAmt.text	 = UNIConvNumPCToCompanyByCurrency(dblTotArDcAmt,frm1.hArDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
		End If	
        frm1.txtTotArClsLocAmt.text = UNIConvNumPCToCompanyByCurrency(dblTotArClsLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")
        frm1.txtTotArDcLocAmt.text  = UNIConvNumPCToCompanyByCurrency(dblTotArDcLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")                
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
		' 입금액 
		ggoOper.FormatFieldByObjectOfCur .txtTotRcptAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 입금잔액 
		ggoOper.FormatFieldByObjectOfCur .txtTotBalAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 반제금액 
		ggoOper.FormatFieldByObjectOfCur .txtTotAllcAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		
		' 채권액 
		ggoOper.FormatFieldByObjectOfCur .txtTotArAmt, .hArDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 채권잔액 
		ggoOper.FormatFieldByObjectOfCur .txtTotArRemAmt, .hArDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 채권반제 
		ggoOper.FormatFieldByObjectOfCur .txtTotArClsAmt, .hArDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 할인금액 
		ggoOper.FormatFieldByObjectOfCur .txtTotArDcAmt, .hArDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
	End With
End Sub

'===================================== CurFormatNumSprSheet()  ======================================
'	Name : CurFormatNumSprSheet()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet()
	With frm1
		ggoSpread.Source = frm1.vspdData
		' 할인금액 
		ggoSpread.SSSetFloatByCellOfCur C_DcAmt,-1, .hArDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		

		ggoSpread.Source = frm1.vspdData1
		' 입금액 
		ggoSpread.SSSetFloatByCellOfCur C_RcptAmt,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		' 입금잔액 
		ggoSpread.SSSetFloatByCellOfCur C_BalAmt,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		' 반제금액 
		ggoSpread.SSSetFloatByCellOfCur C_AllcAmt,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec

		ggoSpread.Source = frm1.vspdData10
		' 채권액 
		ggoSpread.SSSetFloatByCellOfCur C_ArAmt,-1, .hArDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		' 채권잔액 
		ggoSpread.SSSetFloatByCellOfCur C_ArRemAmt,-1, .hArDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		' 반제금액 
		ggoSpread.SSSetFloatByCellOfCur C_ArClsAmt,-1, .hArDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		' 할인금액 
		ggoSpread.SSSetFloatByCellOfCur C_ArDcAmt,-1, .hArDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
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
			.vspdData1.Col = C_AllcLocAmt	
			.vspdData1.Text = ""    
			ggoSpread.Source = .vspdData1
			ggoSpread.UpdateRow ii		
		Next	
			
		For ii = 1 To .vspdData10.MaxRows 
			.vspdData10.Row = ii	
			.vspdData10.Col = C_ArClsLocAmt	
			.vspdData10.Text = ""    		
			.vspdData10.Row = ii	
			.vspdData10.Col = C_ArDcLocAmt	
			.vspdData10.Text = ""    		
			ggoSpread.Source = .vspdData10
			ggoSpread.UpdateRow ii				
		Next	
		
		.txtTotAllcLocAmt.text="0"
		.TxtTotArClsLocAmt.text="0"
			
		For ii = 1 To .vspdData.MaxRows 
			.vspdData.Row = ii	
			.vspdData.Col = C_DcLocAmt	
			.vspdData.Text = ""    		
			ggoSpread.Source = .vspdData
			ggoSpread.UpdateRow ii				
		Next
			
		.TxtTotArDcLocAmt.Text = "0"
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

	Select Case Trim(UCase(gActiveSpdSheet.Name))
		Case "VSPDDATA" 
			ggoSpread.Source = frm1.vspdData		
			Call PrevspdDataRestore(gActiveSpdSheet)
			Call ggoSpread.RestoreSpreadInf()
			Call InitSpreadSheet("C")
			Call ggoSpread.ReOrderingSpreadData()
			Call SetSpreadLock("C")
			Call SetSpread2ColorAr()									
		Case "VSPDDATA1" 
			ggoSpread.Source = frm1.vspdData1		
			Call ggoSpread.RestoreSpreadInf()
			Call InitSpreadSheet("A")
			Call ggoSpread.ReOrderingSpreadData()
			Call SetSpreadLock("A")	
			Call ggoOper.SetReqAttr(frm1.txtALLCDt,   "N")		
		Case "VSPDDATA10" 
			ggoSpread.Source = frm1.vspdData10
			Call ggoSpread.RestoreSpreadInf()
			Call InitSpreadSheet("B")
			Call ggoSpread.ReOrderingSpreadData()									
			Call SetSpreadLock("B")			
			Call ggoOper.SetReqAttr(frm1.txtALLCDt,   "N")
		Case "VSPDDATA2"
			ggoSpread.Source = frm1.vspdData2				
			Call PrevspdData2Restore(gActiveSpdSheet)   
			Call ggoSpread.RestoreSpreadInf()
			Call InitCtrlSpread()			'관리항목 그리드 초기화 
			Call ggoSpread.ReOrderingSpreadData()
			Call SetSpread2ColorAr()  
	End Select
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
		strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD,'')),CHAR(8) "
    		
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
		Call SetToolBar("1110111100001111")           
    Else                 
        Call SetToolBar("1111111100001111")           
    End If    
End Sub

'=======================================================================================================
'   Event Name : vspdData1_onfocus
'   Event Desc :
'=======================================================================================================
Sub  vspdData1_onfocus()
    lgVspdNo = 1
End Sub

'=======================================================================================================
'   Event Name : vspdData10_onfocus
'   Event Desc :
'=======================================================================================================
Sub  vspdData10_onfocus()
    lgVspdNo = 0
End Sub

'=======================================================================================================
'   Event Name : vspdData2_onfocus
'   Event Desc :
'=======================================================================================================
Sub  vspdData2_onfocus()
    
End Sub

'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub  vspdData_Click(ByVal Col, ByVal Row)

	Call SetPopUpMenuItemInf("1101111111")
	
    gMouseClickStatus = "SP1C"									'Split 상태코드 
 
	Set gActiveSpdSheet = frm1.vspdData
	
	If frm1.vspdData.Maxrows = 0 Then
	    Exit Sub
	End if

	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col							'Ascending Sort
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col,lgSortKey				'Descending Sort
			lgSortKey = 1
		End If																
		Exit Sub
	End If		

	ggoSpread.Source = frm1.vspdData
	frm1.vspdData.Row = frm1.vspdData.ActiveRow	

 	frm1.vspdData.Col = C_AcctCd
	
    If Len(frm1.vspdData.Text) > 0 Then
	Else
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.ClearSpreadData		
	End if	
End Sub

'==========================================================================================
'   Event Name : vspdData1_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)
    Call SetPopUpMenuItemInf("0000111111")
    
    gMouseClickStatus = "SPC"									'Split 상태코드 
 
	Set gActiveSpdSheet = frm1.vspdData1
	
	If frm1.vspdData.Maxrows = 0 Then
	    Exit Sub
	End if

	If Row <= 0 Then
		Exit Sub
	End If		
End Sub

'==========================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData10_Click(ByVal Col, ByVal Row)
	Call SetPopUpMenuItemInf("0000111111")    

    gMouseClickStatus = "SP2C"									'Split 상태코드 
 
	Set gActiveSpdSheet = frm1.vspdData10
	
	If frm1.vspdData.Maxrows = 0 Then
	    Exit Sub
	End if

	If Row <= 0 Then
		Exit Sub
	End If		
End Sub

'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP1C" Then
		gMouseClickStatus = "SP1CR"
	End If
End Sub

Sub vspdData1_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

Sub vspdData10_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
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
'   Event Name : vspdData10_ColWidthChange
'   Event Desc : 상세내역 그리드의 (멀티)컬럼의 너비를 조절하는 경우 
'=======================================================================================================
Sub  vspdData10_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.Source = frm1.vspdData10
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'=======================================================================================================
'   Event Name : vspdData1_ScriptLeaveCell
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
'   Event Name : vspdData_Change
'   Event Desc :
'=======================================================================================================
Sub  vspdData_Change(ByVal Col, ByVal Row )
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    
    frm1.vspdData.Row = Row
    frm1.vspdData.Col = 0    
    Select case Col
		Case   C_AcctCD
			If frm1.vspddata.Text = ggoSpread.InsertFlag Then
			    frm1.vspddata.Col = C_ItemSeq
			    frm1.hItemSeq.value = frm1.vspddata.Text
			    frm1.vspddata.Col = C_AcctCd
			    
			    If Len(frm1.vspddata.Text) > 0 Then
					frm1.vspddata.Row = Row
					frm1.vspddata.Col = C_ItemSeq	   	
					DeleteHsheet frm1.vspddata.Text
			        Call DbQuery3(Row)
			        Call SetSpread2ColorAr()
			    End If    
			End If 
		Case C_DcAmt
			frm1.vspddata.col = C_DcLocAmt
			frm1.vspdData.text=""
			
	End Select
End Sub

'======================================================================================================
'   Event Name : vspdData1_Change
'   Event Desc :
'=======================================================================================================
Sub  vspdData1_Change(ByVal Col, ByVal Row )
	Dim RcptAmt
	Dim AllcAmt

    ggoSpread.Source = frm1.vspdData1
    ggoSpread.UpdateRow Row

    frm1.vspdData2.Row = Row
    frm1.vspdData2.Col = 0

	Select Case Col
		Case C_AllcAmt
			frm1.vspdData1.Col = C_RcptAmt
			RcptAmt = frm1.vspdData1.Text
			frm1.vspdData1.Col = C_AllcAmt
			AllcAmt = UNICDbl(frm1.vspdData1.Text)
			frm1.vspdData1.Col = C_AllcLocAmt
			frm1.vspddata1.text = ""

			If (UNICDbl(RcptAmt) > 0 And parent.UNICDbl(AllcAmt) < 0) Or (UNICDbl(RcptAmt) < 0 And parent.UNICDbl(AllcAmt) > 0) Then
				frm1.vspdData1.Col = C_AllcAmt
				
				frm1.vspdData1.Text  = UNIConvNumPCToCompanyByCurrency(AllcAmt * (-1),frm1.txtDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
			End If
			Call DoSum()   
		Case C_AllcLocAmt
			Call DoSum()   
	End Select
End Sub

'======================================================================================================
'   Event Name : vspdData10_Change
'   Event Desc :
'=======================================================================================================
Sub  vspdData10_Change(ByVal Col, ByVal Row )
	Dim ArAmt
	Dim ClsAmt
	Dim DcAmt
	Dim dblTotDcAmt

    ggoSpread.Source = frm1.vspdData10
    ggoSpread.UpdateRow Row

    frm1.vspdData10.Row = Row
    frm1.vspdData10.Col = 0             

	Select Case Col
		Case C_ArClsAmt
			frm1.vspdData10.Col = C_ArAmt
			ArAmt = frm1.vspdData10.Text
			frm1.vspdData10.Col = C_ArClsAmt
			ClsAmt = UNICDbl(frm1.vspdData10.Text)
			frm1.vspdData10.col=C_ArClsLocAmt
			frm1.vspdData10.text="" 		

			If (UNICDbl(ArAmt) > 0 And parent.UNICDbl(ClsAmt) < 0) Or (UNICDbl(ArAmt) < 0 And parent.UNICDbl(ClsAmt) > 0) Then
				frm1.vspdData10.Col = C_ArClsAmt
				frm1.vspdData10.Text  = UNIConvNumPCToCompanyByCurrency(ClsAmt * (-1),frm1.hArDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
			
			End If
			Call DoSum()
		Case C_ArDcAmt
			frm1.vspdData10.Col = C_ArAmt
			ArAmt = frm1.vspdData10.Text
			frm1.vspdData10.Col = C_ArDcAmt
			DcAmt = UNICDbl(frm1.vspdData10.Text)
			frm1.vspdData10.col=C_ArDcLocAmt
			frm1.vspdData10.text="" 													

			If (UNICDbl(ArAmt) > 0 And parent.UNICDbl(DcAmt) < 0) Or (UNICDbl(ArAmt) < 0 And parent.UNICDbl(DcAmt) > 0) Then
				frm1.vspdData10.Col = C_ArDcAmt
				frm1.vspdData10.Text = UNIConvNumPCToCompanyByCurrency(DcAmt * (-1),frm1.hArDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
			End If		    
			
			dblTotDcAmt = FncSumSheet1(frm1.vspdData10,C_ArDcAmt , 1, frm1.vspdData10.MaxRows, False, -1, -1, "V")
			frm1.TxtTotArDcAmt.text = UNIConvNumPCToCompanyByCurrency(dblTotDcAmt ,frm1.hArDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")

			If UCase(Trim(frm1.hArDocCur.Value)) <> UCase(parent.gCurrency) Then
				frm1.vspdData1.Col = C_ArDcLocAmt
				frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
				frm1.vspdData1.Text = ""							
			End If
			Call DoSum()
	   Case C_ArClsLocAmt, C_ArDcLocAmt
			Call DoSum()
    End Select
End Sub

'======================================================================================================
'   Event Name :vspdData_DblClick
'   Event Desc :
'=======================================================================================================
Sub  vspdData_DblClick( ByVal Col , ByVal Row )
	If Row <=0 Then
		Exit Sub				
	End If		

    If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If
End Sub

'======================================================================================================
'   Event Name :vspdData1_DblClick
'   Event Desc :
'=======================================================================================================
Sub  vspdData1_DblClick( ByVal Col , ByVal Row )
	If Row <=0 Then
		Exit Sub				
	End If		

    If frm1.vspdData1.MaxRows = 0 Then
		Exit Sub
	End If
End Sub

'======================================================================================================
'   Event Name :vspdData_DblClick
'   Event Desc :
'=======================================================================================================
Sub  vspdData10_DblClick( ByVal Col , ByVal Row )
	If Row <=0 Then
		Exit Sub				
	End If		

    If frm1.vspdData10.MaxRows = 0 Then
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
	Call GetSpreadColumnPos("C")
End Sub

'======================================================================================================
'   Event Name :vspddata1_ScriptDragDropBlock
'   Event Desc :
'=======================================================================================================
Sub  vspddata1_ScriptDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
	ggoSpread.Source = frm1.vspdData1 
	Call ggoSpread.SpreadDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
	Call GetSpreadColumnPos("A")
End Sub

'======================================================================================================
'   Event Name :vspddata10_ScriptDragDropBlock
'   Event Desc :
'=======================================================================================================
Sub  vspddata10_ScriptDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
	ggoSpread.Source = frm1.vspdData10
	Call ggoSpread.SpreadDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
	Call GetSpreadColumnPos("B")
End Sub

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
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
'   Event Name :vspdData1_EditChange
'   Event Desc :
'=======================================================================================================
Sub  vspdData1_EditChange(ByVal Col , ByVal Row )
                
End Sub

'======================================================================================================
'   Event Name :vspdData1_KeyPress
'   Event Desc :
'=======================================================================================================
Sub  vspdData1_KeyPress(KeyAscii )
     
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




'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.7 Date-Numeric OCX Tag Event
' Description : This part declares HTML Tag Event
'=======================================================================================================
'*******************************************************************************************************
'==========================================================================================
'   Event Name : txtDeptCd_OnBlur
'   Event Desc : 
'==========================================================================================

Sub txtDeptCd_OnBlur()
        
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii, jj

	If Trim(frm1.txtAllcDt.Text = "") Then    
		Exit sub
    End If
    lgBlnFlgChgValue = True
	
	If TRim(frm1.txtDeptCd.value) <>"" Then
		'----------------------------------------------------------------------------------------
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtAllcDt.Text, gDateFormat,""), "''", "S") & "))"			
		
	
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
	End if
		'----------------------------------------------------------------------------------------
End Sub

'==========================================================================================
'   Event Name : txtAllcDt_onBlur
'   Event Desc : 
'==========================================================================================
Sub txtAllcDt_onBlur()
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii, jj
	Dim arrVal1, arrVal2
 
  	lgBlnFlgChgValue = True
	With frm1
	
		If LTrim(RTrim(.txtDeptCd.value)) <> "" and Trim(.txtAllcDt.Text <> "") Then
			'----------------------------------------------------------------------------------------
				strSelect	=			 " Distinct org_change_id "    		
				strFrom		=			 " b_acct_dept(NOLOCK) "		
				strWhere	=			 " dept_Cd =  " & FilterVar(LTrim(RTrim(.txtDeptCd.value)), "''", "S") 
				strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
				strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
				strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(.txtAllcDt.Text, gDateFormat,""), "''", "S") & "))"			
	
			IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
					
			If IntRetCD = False  OR Trim(Replace(lgF0,Chr(11),"")) <> Trim(.hOrgChangeId.value) Then
					.txtDeptCd.value = ""
					.txtDeptNm.value = ""
					.hOrgChangeId.value = ""
					.txtDeptCd.focus
			End if
		End If
	End With
'----------------------------------------------------------------------------------------
End Sub




'=======================================================================================================
'   Event Name : txtAllcDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtAllcDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtAllcDt.Action = 7
        Call txtAllcDt_onBlur()
		Call SetFocusToDocument("M")
		frm1.txtAllcDt.Focus
    End If
End Sub



</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  --> 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>채권반제(가수금)</font></td>
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
					<TD WIDTH=* align=right><A HREF="VBSCRIPT:OpenPopupTempGL()">결의전표</A>&nbsp;|&nbsp;<A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A>&nbsp;|&nbsp;<a href="vbscript:OpenRefRcptNo()">가수금정보</A>&nbsp;|&nbsp;<a href="vbscript:OpenRefOpenAr()">채권발생정보</A></TD>								
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
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtAllcNo" ALT="반제번호" MAXLENGTH=18 STYLE="TEXT-ALIGN: Left" tag ="12XXXU"><IMG align=top name=btnCalType src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON" onclick="vbscript: Call OpenPopup(frm1.txtAllcNo.value,0)"></TD>
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
					<TD WIDTH="100%">
					
					
					<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; OVERFLOW: auto; WIDTH: 100%" SCROLL="no">
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>거래처</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtBpCd" ALT="거래처" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN: Left" tag="24NXXU"><IMG align=top name=btnCalType onclick="vbscript:Call OpenPopup(frm1.txtBpCd.value,1)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON">&nbsp;<INPUT  NAME="txtBpNm" SIZE="20" tag = "24" ></TD>								
								<TD CLASS="TD5" NOWRAP>거래통화</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDocCur" ALT="거래통화" MAXLENGTH="3" SIZE=10 STYLE="TEXT-ALIGN: Left" tag ="24NXXU"><IMG align=top name=btnCalType onclick="vbscript:Call OpenPopup(frm1.txtDocCur.value,3)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON"></TD>
							</TR>							
							<TR>
								<TD CLASS="TD5" NOWRAP>부서</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDeptCd" ALT="부서" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN: Left" tag="22NXXU"><IMG align=top name=btnCalType onclick="vbscript:Call OpenDept(frm1.txtDeptCd.Value, 0)"" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON">&nbsp;<INPUT  NAME="txtDeptNm" SIZE="20" tag = "24" ></TD>								
								<TD CLASS="TD5" NOWRAP>반제일</TD>
								<TD CLASS="TD6" NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtAllcDt" CLASS=FPDTYYYYMMDD tag="22" Title="FPDATETIME" ALT="반제일" ></OBJECT>');</SCRIPT></TD>												
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>결의전표번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTempGlNo" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24XXXU" ALT="결의전표번호"> </TD>																						
								<TD CLASS="TD5" NOWRAP>전표번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtGlNo" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24XXXU" ALT="전표번호"></TD>								
							</TR>	
							<TR>
								<TD CLASS="TD5" NOWRAP>비고</TD>
								<TD CLASS="TD6" NOWRAP COLSPAN=3><INPUT TYPE=TEXT NAME="txtRcptDesc" SIZE=90 MAXLENGTH=128 tag="21XXX" ALT="비고"></TD>
							</TR>													
								<TR HEIGHT="100%">
								<TD COLSPAN=2>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" name=vspdData1 width="100%" TITLE="SPREAD" tag="23"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
								</TD>		
								<TD COLSPAN=2>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" name=vspdData10 width="100%" TITLE="SPREAD" tag="23"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD COLSPAN=4>
									<TABLE <%=LR_SPACE_TYPE_60%>>
										<TR>
											<TD CLASS=TDT NOWRAP></TD>
											<TD CLASS=TDT NOWRAP>입금액</TD>
											<TD CLASS=TDT NOWRAP></TD>
											<TD CLASS=TDT NOWRAP >반제금액</TD>
											<TD CLASS=TDT NOWRAP></TD>
											<TD CLASS=TDT NOWRAP>채권액</TD>
											<TD CLASS=TDT NOWRAP></TD>
											<TD CLASS=TDT NOWRAP>채권반제</TD>
										</TR>
										<TR>														
											<TD CLASS=TDT NOWRAP COLSPAN=2><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTotRcptAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="입금반제" tag="24X2" id=OBJECT5></OBJECT>');</SCRIPT></TD>
											<TD CLASS=TDT NOWRAP COLSPAN=2><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTotAllcAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="반제금액" tag="24X2" id=OBJECT3></OBJECT>');</SCRIPT></TD>
											<TD CLASS=TDT NOWRAP COLSPAN=2><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="TxtTotArAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="채권액" tag="24X2" id=OBJECT7></OBJECT>');</SCRIPT></TD>
											<TD CLASS=TDT NOWRAP COLSPAN=2><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="TxtTotArClsAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="반제금액" tag="24X2" id=OBJECT9></OBJECT>');</SCRIPT></TD>
										</TR>
													
										<TR>
											<TD CLASS=TDT NOWRAP></TD>
											<TD CLASS=TDT NOWRAP>입금잔액</TD>
											<TD CLASS=TDT NOWRAP></TD>
											<TD CLASS=TDT NOWRAP>반제금액(자국통화)</TD>
											<TD CLASS=TDT NOWRAP></TD>
											<TD CLASS=TDT NOWRAP>채권잔액</TD>
											<TD CLASS=TDT NOWRAP></TD>
											<TD CLASS=TDT NOWRAP>채권반제(자국통화)</TD>
										</TR>
										<TR>
											<TD CLASS=TDT NOWRAP COLSPAN=2><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTotBalAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="입금잔액" tag="24X2" id=OBJECT6></OBJECT>');</SCRIPT></TD>
											<TD CLASS=TDT NOWRAP COLSPAN=2><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTotAllcLocAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="반제금액(자국)" tag="24X2" id=OBJECT4></OBJECT>');</SCRIPT></TD>
											<TD CLASS=TDT NOWRAP COLSPAN=2><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="TxtTotArRemAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="채권잔액" tag="24X2" id=OBJECT8></OBJECT>');</SCRIPT></TD>
											<TD CLASS=TDT NOWRAP COLSPAN=2><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="TxtTotArClsLocAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="반제금액(자국)" tag="24X2" id=OBJECT10></OBJECT>');</SCRIPT></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>	
						</TABLE>
					</DIV>
					
					
					
					<DIV ID="TabDiv" SCROLL=no>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR HEIGHT="60%">
								<TD WIDTH="100%" COLSPAN=4>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" name=vspdData width="100%" tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD COLSPAN=4>
									<TABLE <%=LR_SPACE_TYPE_20%>>
										<TR>
											<TD class=TD5 NOWRAP>할인금액</TD>
											<TD class=TD6 NOWRAP>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="TxtTotArDcAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="할인금액" tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>
											<TD class=TD5 STYLE="WIDTH : 0px;"></TD>
											<TD class=TD5 NOWRAP>할인금액(자국)</TD>											
											<TD class=TD6 NOWRAP>											
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="TxtTotArDcLocAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="할인금액(자국)" tag="24X2" id=OBJECT2></OBJECT>');</SCRIPT></TD>
											
										</TR>
									</TABLE>
								</TD>
							</TR>
						    <TR HEIGHT="40%">
								<TD WIDTH="100%" COLSPAN=4>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" name=vspdData2 width="100%" tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT></TD>
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
<TEXTAREA Class=hidden name=txtSpread    tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA Class=hidden name=txtSpread1   tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA Class=hidden name=txtSpread2   tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA Class=hidden name=txtSpread3   tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode"		 tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtUpdtUserId"  tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows"	 tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows0"    tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows1"    tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtFlgMode"     tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="htxtAllcNo"     tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hItemSeq"       tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hAcctCd"		 tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows3"	 tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hOrgChangeId"	 tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hArDocCur"		 tag="24" TABINDEX="-1">

<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT TYPE=hidden CLASSID=<%=gCLSIDFPSPD%> HEIGHT=0 name=vspdData3 width="100%" tag="2" TABINDEX="-1"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
