<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>


<!--
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : 
'*  3. Program ID           : a4106ma1.asp
'*  4. Program Name         : 채무반제(선금급)
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
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="javascript"   SRC="../../inc/TabScript.js">			</SCRIPT>
<SCRIPT LANGUAGE="VBScript"     SRC="../ag/AcctCtrl.vbs">				</SCRIPT>
<SCRIPT LANGUAGE=vbscript>

Option Explicit																	'☜: indicates that All variables must be declared in advance
<!-- #Include file="../../inc/lgvariables.inc" -->	
'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Const BIZ_PGM_QRY_ID = "a4106mb1.asp"												'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID = "a4106mb2.asp"												'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_DEL_ID =  "a4106mb3.asp"
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================

Const TAB1 = 1																		'☜: Tab의 위치 
Const TAB2 = 2

Dim C_ApNo 
Dim C_ApDt 
Dim C_ApDueDt 
Dim C_DocCur 
Dim C_ApAmt 
Dim C_ApRemAmt 
Dim C_ApClsAmt 
Dim C_ApClsLocAmt 
Dim C_ApDcAmt 
Dim C_ApDcLocAmt 
Dim C_ApClsDesc
Dim C_ApAcctCd 
Dim C_AcctNmap 
Dim C_BizCd 
Dim C_BizNm 

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
Dim  lgRetFlag	                'Popup
Dim  lgCurrRow
Dim  lgQueryOk

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
			C_ApNo = 1
			C_ApDt = 2
			C_ApDueDt = 3
			C_DocCur = 4
			C_ApAmt = 5
			C_ApRemAmt = 6
			C_ApClsAmt = 7
			C_ApClsLocAmt = 8
			C_ApDcAmt = 9
			C_ApDcLocAmt = 10
			C_ApClsDesc = 11
			C_ApAcctCd = 12
			C_AcctNmap = 13							
			C_BizCd = 14
			C_BizNm = 15
		Case "B"
			C_ItemSeq = 1																	
			C_AcctCd = 2
			C_AcctPB = 3
			C_AcctNm = 4
			C_DcAmt = 5
			C_DcLocAmt = 6
	End Select				
End Sub

'======================================================================================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=======================================================================================================
Sub  InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE								'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False										'Indicates that no value changed
    lgIntGrpCount = 0												'initializes Group View Size
        
    lgStrPrevKey = ""												'initializes Previous Key
    lgStrPrevKey1 = ""
    lgStrPrevKeyDtl = 0												'initializes Previous Key
    lgLngCurRows = 0												'initializes Deleted Rows Count
    
    lgSortKey = 1
    lgQueryOk = False
End Sub

'======================================================================================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=======================================================================================================
Sub  SetDefaultVal()
	frm1.txtAllcDt.text =  UniConvDateAToB("<%=dtToday%>", parent.gServerDateFormat,gDateFormat)
	frm1.txtDocCur.value    = parent.gCurrency
	frm1.hApDocCur.value    = parent.gCurrency
	frm1.txtBalAmt.text		= "0"
	frm1.txtBalLocAmt.text	= "0"
	frm1.txtClsLocAmt.text	= "0"
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
Sub  InitSpreadSheet(ByVal pvSpdNo)
    Call initSpreadPosVariables(pvSpdNo)

	Select Case UCase(Trim(pvSpdNo))
		Case "A"
			With frm1.vspddata1
				ggoSpread.Source = frm1.vspdData1
				ggoSpread.SpreadInit "V20021127",,parent.gAllowDragDropSpread 

				.Redraw = False	

				.MaxCols = C_BizNm + 1   
				.Col = .MaxCols
				.ColHidden = True
				.MaxRows = 0
	
				Call GetSpreadColumnPos(pvSpdNo)	 
 
				ggoSpread.SSSetEdit  C_ApNo       , "채무번호"      , 18, 3		'1
				ggoSpread.SSSetDate  C_ApDt       , "채무일자"      , 10, 2, gDateFormat  
				ggoSpread.SSSetDate  C_ApDueDt    , "만기일자"      , 10, 2, gDateFormat    
				ggoSpread.SSSetEdit  C_DocCur     , "거래통화"      ,  8, 3'10  
				ggoSpread.SSSetFloat C_ApAmt      , "채무액"        , 15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_ApRemAmt   , "채무잔액"      , 15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_ApClsAmt   , "반제금액"      , 15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_ApClsLocAmt, "반제금액(자국)", 15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_ApDcAmt    , "할인금액",15   , parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_ApDcLocAmt , "할인금액(자국)", 15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetEdit  C_ApClsDesc  , "비고"          , 20, 3	'2
				ggoSpread.SSSetEdit  C_ApAcctCd   , "계정코드"      , 10, 3	'2
				ggoSpread.SSSetEdit  C_AcctNmAp   , "계정코드명"    , 15, 3	'3    
				ggoSpread.SSSetEdit  C_BizCd      , "사업장"        , 10, 3	'6
				ggoSpread.SSSetEdit  C_BizNm      , "사업장명"      , 15, 3	'7
    
				.Redraw = True 
			End With    
		Case "B"    
			With frm1.vspddata
				ggoSpread.Source = frm1.vspdData
				ggoSpread.SpreadInit "V20021128",,parent.gAllowDragDropSpread 

				.Redraw = False	

				.MaxCols = C_DcLocAmt + 1 												'☜: 최대 Columns의 항상 1개 증가시킴 
				.Col = .MaxCols															'공통콘트롤 사용 Hidden Column
				.ColHidden = True       
				.MaxRows = 0		
    
				Call GetSpreadColumnPos(pvSpdNo)    
				Call AppendNumberPlace("6","3","0")
    
				ggoSpread.SSSetFloat  C_ItemSeq , "NO"            ,  6,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,2,,,"0","999"    
				ggoSpread.SSSetEdit	  C_AcctCd  , "계정코드"      , 20, ,,20, 2
				ggoSpread.SSSetButton C_AcctPB
				ggoSpread.SSSetEdit	  C_AcctNm  , "계정코드명"    , 50,,,20,2
				ggoSpread.SSSetFloat  C_DcAmt   , "할인금액"      , 20, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat  C_DcLocAmt, "할인금액(자국)", 20, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				
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
Sub  SetSpreadLock(pvSpdNo)
    With frm1
		Select Case UCase(Trim(pvSpdNo))
			Case "A"
				ggoSpread.Source = .vspddata1
				.vspddata1.ReDraw = False

				ggoSpread.SpreadLock C_ApNo    ,-1, C_ApNo
				ggoSpread.SpreadLock C_ApAcctCd,-1, C_ApAcctCd
				ggoSpread.SpreadLock C_AcctNmAp,-1, C_AcctNmAp
				ggoSpread.SpreadLock C_BizCd   ,-1, C_BizCd
				ggoSpread.SpreadLock C_BizNm   ,-1, C_BizNm
				ggoSpread.SpreadLock C_DocCur  ,-1, C_DocCur    
				ggoSpread.SpreadLock C_ApDt    ,-1, C_ApDt
				ggoSpread.SpreadLock C_ApDueDt ,-1, C_ApDueDt    
				ggoSpread.SpreadLock C_ApAmt   ,-1, C_ApAmt
				ggoSpread.SpreadLock C_ApRemAmt,-1, C_ApRemAmt    
    
				ggoSpread.SSSetRequired C_ApClsAmt, -1, -1

				.vspddata1.ReDraw = True   
			Case "B"
				ggoSpread.Source = .vspddata
				.vspddata.Redraw = False    
    
				ggoSpread.SpreadLock C_ItemSeq, -1, C_ItemSeq
				ggoSpread.SpreadLock C_AcctCd , -1, C_AcctCd 
				ggoSpread.SpreadLock C_AcctPB , -1, C_AcctPB 
				ggoSpread.SpreadLock C_AcctNm , -1, C_AcctNm 
		    
				ggoSpread.SSSetRequired C_DcAmt, -1, -1        
			
				.vspddata1.ReDraw = True   		
		End Select			
    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub  SetSpreadColor(ByVal pvStartRow , ByVal pvEndRow)							'행추가, 행복사 후 추가된 그리드 세팅 
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
' Function Name : SetSpread2ColorAp
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub  SetSpread2ColorAp()
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
			
			C_ApNo		  = iCurColumnPos(1)
			C_ApDt        = iCurColumnPos(2)
			C_ApDueDt     = iCurColumnPos(3)
			C_DocCur      = iCurColumnPos(4)
			C_ApAmt       = iCurColumnPos(5)
			C_ApRemAmt    = iCurColumnPos(6)
			C_ApClsAmt	  = iCurColumnPos(7)
			C_ApClsLocAmt = iCurColumnPos(8)
			C_ApDcAmt     = iCurColumnPos(9)
			C_ApDcLocAmt  = iCurColumnPos(10)
			C_ApClsDesc   = iCurColumnPos(11)
			C_ApAcctCd    = iCurColumnPos(12)
			C_AcctNmap    = iCurColumnPos(13)							
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

'========================================================================================================= 
'	Name : OpenRefOpenAp()
'	Description : Ref 화면을 call한다. 
'========================================================================================================= 
Function OpenRefOpenAp()
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName
	Dim IntRetCD

	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("A4105RA5")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "A4105RA5", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	If gSelframeFlg <> TAB1 Then Exit Function		 
	
	IsOpenPop = True
	
	If frm1.vspdData1.MaxRows = 0 Then frm1.hApDocCur.value	= ""	

	arrParam(0) = frm1.txtBpCd.value				' 검색조건이 있을경우 파라미터 
	arrParam(1) = frm1.txtBpNm.value			
	arrParam(2) = frm1.hApDocCur.value	
	arrParam(3) = frm1.txtAllcDt.text			
	arrParam(4) = frm1.txtAllcDt.alt				

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	
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
'	Name : openpopupgl()
'	Description : 
'========================================================================================================= 
Function OpenPopupGL()
	Dim arrRet
	Dim arrParam(1)	
	Dim iCalledAspName
	Dim IntRetCD

	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("A5120RA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "A5120RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrParam(0) = Trim(frm1.txtGlNo.value)	'회계전표번호 
	arrParam(1) = ""						'Reference번호 

	IsOpenPop = True
   
    arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function

'========================================================================================================= 
'	Name : openTempglpopup
'	Description :결의전표  POP-UP
'========================================================================================================= 
Function OpenPopuptempGL()
	Dim arrRet
	Dim arrParam(1)	
Dim iCalledAspName
	Dim IntRetCD

	If IsOpenPop = True Then Exit Function
	
	arrParam(0) = Trim(frm1.txtTempGlNo.value)	'회계전표번호 
	arrParam(1) = ""						'Reference번호 
	
	iCalledAspName = AskPRAspName("A5130RA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "A5130RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True
    arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		     	
	IsOpenPop = False
End Function

'========================================================================================================= 
'	Name : SetRefOpenAp()
'	Description : OpenAp Popup에서 Return되는 값 setting
'========================================================================================================= 
Function SetRefOpenAp(Byval arrRet)
	Dim intRtnCnt, strData
	Dim TempRow, I
	Dim X
	Dim sFindFg
	
	With frm1
		.vspddata1.focus		
		ggoSpread.Source = .vspddata1
		.vspddata1.ReDraw = False	
	
		TempRow = .vspddata1.MaxRows												'☜: 현재까지의 MaxRows

		For I = TempRow to TempRow + Ubound(arrRet, 1)
			sFindFg	= "N"
			For x = 1 to TempRow
				.vspddata1.Row = x
				.vspddata1.Col = C_ApNo				
				If "" & UCase(Trim(.vspddata1.Text)) = "" & UCase(Trim(arrRet(I - TempRow, 0))) Then
					sFindFg	= "Y"
				End If
			Next			
			If 	sFindFg	= "N" Then
				.vspddata1.MaxRows = .vspddata1.MaxRows + 1
				.vspddata1.Row = I + 1				
				.vspddata1.Col = 0
				.vspddata1.Text = ggoSpread.InsertFlag

				.vspddata1.Col = C_ApNo												
				.vspddata1.text = arrRet(I - TempRow, 0)				
				.vspddata1.Col = C_ApDt 												
				.vspddata1.text = arrRet(I - TempRow, 1)				
				.vspddata1.Col = C_ApDueDt 												
				.vspddata1.text = arrRet(I - TempRow, 2)
				.vspddata1.Col = C_ApAmt 												
				.vspddata1.text = arrRet(I - TempRow, 3)
				.vspddata1.Col = C_ApRemAmt 												
				.vspddata1.text = arrRet(I - TempRow, 4)				
				.vspddata1.Col = C_ApClsAmt
				.vspddata1.text = arrRet(I - TempRow, 6)
				.vspddata1.Col = C_ApAcctCd
				.vspddata1.text = arrRet(I - TempRow, 7)
				.vspddata1.Col = C_AcctNmap
				.vspddata1.text = arrRet(I - TempRow, 8)				
				.vspddata1.Col = C_BizCd
				.vspddata1.text = arrRet(I - TempRow, 9)				
				.vspddata1.Col = C_BizNm
				.vspddata1.text = arrRet(I - TempRow, 10)				
				.vspddata1.Col = C_DocCur
				.vspddata1.text = arrRet(I - TempRow, 11)	
				.vspddata1.Col = C_ApClsDesc				
				.vspddata1.text = arrRet(I - TempRow, 14)								
			End If	
		Next	
		
		.hApDocCur.Value = arrRet(0, 11)				
		'.txtDocCur.value = arrRet(0, 11) '20051201 추가 
		.txtbpCd.Value = UCase(arrRet(0, 12))
		.txtbpNm.Value = arrRet(0, 13)				
		
		If .txtBpCd.value <> "" Then					
			Call ggoOper.SetReqAttr(.txtBpCd,   "Q")		
		Else			
			Call ggoOper.SetReqAttr(.txtBpCd,   "N")		
		End If
	
		If .txtDocCur.value <> "" Then					
			If UCase(Trim(.txtDocCur.Value)) = UCase(Trim(arrRet(0, 11))) Then

			Else
				.txtClsAmt.Text	   = "0"
				.txtClsLocAmt.Text = "0"
			End If
		End If			
		
	
		Call txtDocCur_OnChange()
	
		'Call ggoOper.SetReqAttr(frm1.txtAllcDt,   "Q")
		ggoSpread.SpreadUnlock   C_ApNo    , TempRow + 1, C_ApDt , .vspddata1.MaxRows				
		ggoSpread.ssSetProtected C_ApNo    , TempRow + 1, .vspddata1.MaxRows
		ggoSpread.ssSetProtected C_ApDt    , TempRow + 1, .vspddata1.MaxRows				    
		ggoSpread.SSSetRequired  C_ApClsAmt, TempRow + 1, .vspddata1.MaxRows
		
		Call DoSum()
	
		.vspddata1.ReDraw = True
    End With
End Function

'========================================================================================================= 
'	Name : OpenRefPrePaymNo()
'	Description : Ref 화면을 call한다. 
'========================================================================================================= 
Function OpenRefPrePaymNo()
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName
	Dim IntRetCD

	
	If gSelframeFlg <> TAB1 Then Exit Function		 
	IF lgIntFlgMode = parent.OPMD_UMODE THEN Exit Function
	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("A4107RA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "A4107RA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = frm1.txtBpCd.value				' 검색조건이 있을경우 파라미터 
	arrParam(1) = frm1.txtBpNm.value			
	arrParam(2) = frm1.txtDocCur.value					
    arrParam(3) = frm1.txtAllcDt.text			
	arrParam(4) = frm1.txtAllcDt.alt	

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then		
		Exit Function
	Else		
		Call SetRefPrePaymNo(arrRet)
	End If
	
End Function

'========================================================================================================= 
'	Name : SetRefPrePaymNo()
'	Description : OpenAp Popup에서 Return되는 값 setting
'========================================================================================================= 
Function  SetRefPrePaymNo(Byval arrRet)
	lgBlnFlgChgValue = True
	With frm1
		.txtPPNo.Value				= arrRet(0)		' C_PpNo = 1	
		.txtPPDt.text				= arrRet(7)		'C_PpDt = 8
		.txtDeptCd.Value			= arrRet(5)		'C_BizCd = 6	
		.txtDeptNm.Value		    = arrRet(6)		'C_BizNm = 7	
		.txtBizCd.Value				= arrRet(13)	'C_BizCd = 14
		.txtBizNm.Value				= arrRet(14)	' C_BizNm = 15
		.txtBpCd.Value				= arrRet(3)		'C_BpCd = 4
		.txtBpNm.Value				= arrRet(4)		'C_BpNm = 5
		.txtBankAcct.Value			= arrRet(15)	'C_BankAcctNo = 16	
		.txtCheckNo.Value			= arrRet(16)	'C_CheckNo = 17
		.txtDocCur.value			= arrRet(8)		'C_DocCur = 9		
		.txtBalAmt.Text			    = arrRet(11)	'C_PpRemAmt = 12
		.txtBalLocAmt.Text			= arrRet(12)	'C_PpRemLocAmt = 13
		
		.txtAllcNo.value			= ""
		.txtGlNo.value				= ""			
	End With
	
	Call txtDocCur_OnChange()
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
	Dim IntRetCD

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	Select Case iWhere

		Case 4
			arrParam(0) = "계정코드팝업"										' 팝업 명칭 
			arrParam(1) = "A_ACCT,A_ACCT_GP"										' TABLE 명칭 
			arrParam(2) = Trim(strCode)							 					' Code Condition
			arrParam(3) = ""														' Name Cindition
			arrParam(4) = "A_ACCT.GP_CD = A_ACCT_GP.GP_CD AND A_ACCT.DEL_FG <> " & FilterVar("Y", "''", "S") & " "	' Where Condition
			arrParam(5) = "계정코드"			
	
			arrField(0) = "A_ACCT.ACCT_CD"											' Field명(0)
			arrField(1) = "A_ACCT.ACCT_NM"											' Field명(1)
			arrField(2) = "A_ACCT_GP.GP_CD"											' Field명(2)
			arrField(3) = "A_ACCT_GP.GP_NM"											' Field명(3)
    
			arrHeader(0) = "계정코드"											' Header명(0)
			arrHeader(1) = "계정코드명"											' Header명(1)				
			arrHeader(2) = "그룹코드"											' Header명(2)
			arrHeader(3) = "그룹명"												' Header명(3)				
	End Select				

	iCalledAspName = AskPRAspName("A4106RA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "A4106RA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	' 권한관리 추가 
	iArrParam(5) = lgAuthBizAreaCd
	iArrParam(6) = lgInternalCd
	iArrParam(7) = lgSubInternalCd
	iArrParam(8) = lgAuthUsrID
		
	If iwhere = 0 Then	
		arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, iArrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")				
	Else
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")			
	End If
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then	    
		Call EScPopup(iWhere)
		Exit Function
	Else
		Call SetPopup(arrRet, iWhere)
	End If
End Function
'======================================================================================================
'   Function Name : SetPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function EScPopup(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0		
				.txtAllcNo.focus

			Case 4
				Call SetActiveCell(.vspdData,C_AcctCd,.vspdData.ActiveRow ,"M","X","X")
		End Select				
	End With
	If iwhere  <> 0 Then
		lgBlnFlgChgValue = True
	End If	
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

			Case 4
				.vspddata.Col = C_AcctCd
				.vspddata.Text = arrRet(0)
				.vspddata.Col = C_AcctNm
				.vspddata.Text = arrRet(1)
			
				Call vspddata_Change(C_AcctCd, frm1.vspddata.activerow )						' 변경이 일어났다고 알려줌 
				Call SetActiveCell(.vspdData,C_AcctCd,.vspdData.ActiveRow ,"M","X","X")
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
	If lgIntFlgMode <> parent.OPMD_UMODE Then
	    Call SetToolbar("1110100100001111")										'⊙: 버튼 툴바 제어 
	Else				 
	    Call SetToolbar("1111101100001111")										'⊙: 버튼 툴바 제어 
	End If

	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)	 '~~~ 첫번째 Tab 
	gSelframeFlg = TAB1
End Function

Function ClickTab2()
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetToolBar("1110111100001111")
	Else                   
		Call SetToolBar("1111111100001111")
	End If
	
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
    Call LoadInfTB19029()																		'Load table , B_numeric_format

    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, _
							gDateFormat, parent.gComNum1000, parent.gComNumDec)    
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, _
							gDateFormat, parent.gComNum1000, parent.gComNumDec)
                             
    Call ggoOper.LockField(Document, "N")												'Lock  Suitable  Field
    Call InitSpreadSheet("A")																		'Setup the Spread sheet
    Call InitSpreadSheet("B")																		'Setup the Spread sheet    
	Call InitCtrlSpread()
	Call InitCtrlHSpread()	    
    Call InitVariables()																		'Initializes local global variables
    Call SetDefaultVal()

    Call ClickTab1()
    
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
    Dim var1, var2, var3
    
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
    ggoSpread.Source = frm1.vspddata1
    var1 = ggoSpread.SSCheckChange
    ggoSpread.Source = frm1.vspddata
    var2 = ggoSpread.SSCheckChange
    ggoSpread.Source = frm1.vspdData2
    var3 = ggoSpread.SSCheckChange
    
    If lgBlnFlgChgValue = True Or var1 = True Or var2 = True Or var3 = True Then
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

	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData1
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData		
	ggoSpread.Source = frm1.vspdData3
	ggoSpread.ClearSpreadData

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
    Dim var1, var2, var3
    
    FncNew = False                                                          
	'-----------------------
    'Check previous data area
    '----------------------- 
    ggoSpread.Source = frm1.vspddata1		:    var1 = ggoSpread.SSCheckChange
    ggoSpread.Source = frm1.vspddata		:	 var2 = ggoSpread.SSCheckChange
    ggoSpread.Source = frm1.vspdData2		:    var3 = ggoSpread.SSCheckChange
    
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
    Call ggoOper.ClearField(Document, "1")											'⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")											'Clear Condition Field
    Call ggoOper.LockField(Document, "N")											'Lock  Suitable  Field
    Call InitVariables()																	'Initializes local global variables
    Call SetDefaultVal()    
    Call txtDocCur_OnChange()
    Call DisableRefPop()
	ggoSpread.Source = frm1.vspdData		:	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData1		:	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2		:	ggoSpread.ClearSpreadData		
	ggoSpread.Source = frm1.vspdData3		:	ggoSpread.ClearSpreadData	
    
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
    If DbDelete = False Then														'☜: Delete db data
		Exit Function                                                        
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
    Dim var1, var2, var3
    
    FncSave = False                                                         

    On Error Resume Next																	'☜: Protect system from crashing
    Err.Clear																		
	'-----------------------
    'Check previous data area
    '----------------------- 
    ggoSpread.Source = frm1.vspddata1
    var1 = ggoSpread.SSCheckChange
    ggoSpread.Source = frm1.vspddata
    var2 = ggoSpread.SSCheckChange
    ggoSpread.Source = frm1.vspdData2
    var3 = ggoSpread.SSCheckChange
    '-----------------------
    'Precheck area
    '-----------------------
    If lgBlnFlgChgValue = False And var1 = False And var2 = False And var2 = False Then		'⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","X","X","X")										'⊙: Display Message(There is no changed data.)
        Exit Function
    End If
    '-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "2") Then														'Check contents area
		Exit Function
    End If
    
    ggoSpread.Source = frm1.vspddata1
    If Not ggoSpread.SSDefaultCheck Then
    	Call ClickTab1()												'⊙: Check contents area
		Exit Function
    End If 
    
    ggoSpread.Source = frm1.vspddata
    If Not ggoSpread.SSDefaultCheck Then												'⊙: Check contents area
		Call ClickTab2()
		Exit Function
    End If 
    
    If Not chkAllcDate Then
		Exit Function
    End If  
    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave()																			'☜: Save db data
    
    FncSave = True                                                       
	    		
	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function  FncCopy() 
	Dim  IntRetCD
	 
	If frm1.vspddata1.Maxrows < 1 Then Exit Function 
	frm1.vspddata1.ReDraw = False
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")						'⊙: "Will you destory previous data"
		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
	With frm1
		.vspddata1.ReDraw = False
	
		ggoSpread.Source = frm1.vspddata1	
		ggoSpread.CopyRow
		Call SetSpreadColor(frm1.vspddata1.ActiveRow, frm1.vspddata1.ActiveRow)
    
		.vspddata1.ReDraw = True
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
    
		If frm1.vspddata1.Maxrows < 1 Then Exit Function
		With frm1.vspddata1
			.Row = .ActiveRow
			.Col = 0

			ggoSpread.Source = frm1.vspddata1
			ggoSpread.EditUndo
			Call DoSum()
			
			If frm1.vspdData1.MaxRows < 1 Then 
				Call ggoOper.SetReqAttr(frm1.txtAllcDt,   "N")
				Exit Function
			End if					
				
			.Row = .ActiveRow
			.Col = 0	
			
			For i = .MaxRows to 0 Step -1 
				.Row= i
				.Col =0			
			Next
		End With   
	Else
		If frm1.vspddata.Maxrows < 1 Then Exit Function
		
		With frm1.vspddata
		    .Row = .ActiveRow
		    .Col = 0
		    If .Text = ggoSpread.InsertFlag Then
		        .Col = C_AcctCd
				If Len(Trim(.text)) > 0 Then  
					.Col = C_ItemSeq		        
					DeleteHSheet(.Text)
				End If
		    End if
   
		    ggoSpread.Source = frm1.vspddata	
		    ggoSpread.EditUndo

			If frm1.vspddata.Maxrows < 1 Then Exit Function
			
		    .Row = .ActiveRow
		    .Col = 0		    
			If .Row = 0 Then
				Exit Function
			Else	
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
				End if
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

	ggoSpread.Source = frm1.vspddata2
	ggoSpread.ClearSpreadData
		
    Call ggoOper.LockField(Document, "Q")                              'This function lock the suitable field

	    		
	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function  FncDeleteRow() 
    Dim lDelRows
    Dim  DelItemSeq
    
	If gSelframeFlg = TAB1 Then
		If frm1.vspddata1.Maxrows < 1 Then Exit Function
		ggoSpread.Source = frm1.vspddata1
		
		lDelRows = ggoSpread.DeleteRow		
		Call DoSum()
	Else
		If frm1.vspddata.Maxrows < 1 Then Exit Function
			
		With frm1.vspddata 
		    .Row = .ActiveRow
			.Col = C_ItemSeq 
		    DelItemSeq = .Text
		    
			ggoSpread.Source = frm1.vspddata 

			lDelRows = ggoSpread.DeleteRow
		End With

		ggoSpread.Source = frm1.vspdData2
		ggoSpread.ClearSpreadData
		Call DeleteHsheet(DelItemSeq)
	End If	  
	    		
	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================================
Function  FncPrint() 
    On Error Resume Next                                               
    Call parent.FncPrint()
	    		
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
End Sub

'=======================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function  FncExit()
	Dim IntRetCD
	Dim var1, var2, var3
	
	FncExit = False
	
	ggoSpread.Source = frm1.vspddata1
    var1 = ggoSpread.SSCheckChange
    ggoSpread.Source = frm1.vspddata
    var2 = ggoSpread.SSCheckChange
    ggoSpread.Source = frm1.vspdData2
    var3 = ggoSpread.SSCheckChange
    
    If lgBlnFlgChgValue = True Or var1 = True Or var2 = True Or var3 = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")			'⊙: "Will you destory previous data"
	
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
    Dim strVal
    
    DbDelete = False	
    
    Call LayerShowHide(1)    													
    
    strVal = BIZ_PGM_DEL_ID & "?txtMode=" & parent.UID_M0003
    strVal = strVal & "&txtAllcNo=" & Trim(frm1.txtAllcNo.value)				'☜: 삭제 조건 데이타 

	' 권한관리 추가 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 
    
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    
    DbDelete = True                                                         
End Function

'=======================================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================================
Function DbDeleteOk()												        '삭제 성공후 실행 로직 
	Call ggoOper.ClearField(Document, "1")                           '⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")							'Clear Condition Field
    Call ggoOper.LockField(Document, "N")                            'Lock  Suitable  Field

	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData1
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData		
	ggoSpread.Source = frm1.vspdData3
	ggoSpread.ClearSpreadData

    Call InitVariables()                                                      'Initializes local global variables
    Call ClickTab1()
    Call SetDefaultVal()
    Call DisableRefPop()
    frm1.txtAllcNo.Value = ""
    frm1.txtAllcNo.focus
End Function

'=======================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================================
Function  DbQuery() 
    Dim strVal
    
    DbQuery = False                                                             
    Call LayerShowHide(1)
    
    with frm1
		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001					'☜: 
			strVal = strVal & "&txtAllcNo=" & Trim(.htxtAllcNo.value)				'조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
			strVal = strVal & "&lgStrPrevKeyDtl=" & lgStrPrevKeyDtl
			strVal = strVal & "&txtMaxRows=" & .vspddata.MaxRows
		Else
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001					'☜: 
			strVal = strVal & "&txtAllcNo=" & Trim(.txtAllcNo.value)				'조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
			strVal = strVal & "&lgStrPrevKeyDtl=" & lgStrPrevKeyDtl
			strVal = strVal & "&txtMaxRows=" & .vspddata.MaxRows
		End If
    End With

	' 권한관리 추가 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 

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
        '-----------------------
        'Reset variables area
        '-----------------------
        If .vspddata.MaxRows > 0 Then        
            .vspddata.Row = 1
            .vspddata.Col = C_ItemSeq
            .hItemSeq.Value = .vspddata.Text 
            Call DbQuery2(1)       
        End If
        lgIntFlgMode = parent.OPMD_UMODE												'Indicates that current mode is Update mode        
        Call ggoOper.LockField(Document, "Q")									'This function lock the suitable field
        call ClickTab1()
    End With

	lgQueryOk = True

	Call DoSum()
	Call txtDocCur_OnChange()
	Call DisableRefPop()
	frm1.txtAllcNo.focus
    lgBlnFlgChgValue = FALSE
    
    If Frm1.vspdData1.MaxRows > 0 Then         
		Frm1.vspdData1.Focus 
    End If
    
    lgQueryOk = False
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
    
    ggoSpread.Source = frm1.vspddata1
	With frm1.vspddata1
		For lngRows = 1 To .MaxRows
		    .Row = lngRows
			.Col = 0
				
			Select Case .Text
				Case ggoSpread.DeleteFlag

				Case Else
					strVal = strVal & "C" & parent.gColSep  					'☜: C=Create, Row위치 정보 
			        .Col = C_ApNo								'1
			        strVal = strVal & Trim(.Text) & parent.gColSep
			        .Col = C_ApAcctCd
			        strVal = strVal & Trim(.Text) & parent.gColSep            		                    		        
			        .Col = C_ApDt
			        strVal = strVal & UNIConvDate(Trim(.Text)) & parent.gColSep		        
			        .Col = C_DocCur
			        strVal = strVal & Trim(.Text) & parent.gColSep
			        .Col = C_ApClsAmt
			        strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
			        .Col = C_ApClsLocAmt		            
			        strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
			        .Col = C_ApDcAmt
			        strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
			        .Col = C_ApDcLocAmt		            
			        strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep 
			        .Col = C_ApClsDesc		            
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

	ggoSpread.Source = frm1.vspddata
	With frm1.vspddata
		For lngRows = 1 To .MaxRows
			.Row = lngRows
			.Col = 0
			Select Case .Text
				Case ggoSpread.DeleteFlag

				Case Else
					strVal = strVal & "C" & parent.gColSep & lngRows & parent.gColSep				'C=Create, Sheet가 2개 이므로 구별 
					.Col = C_ItemSeq	'2
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = C_DcAmt		'3
					strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
					.Col = C_DcLocAmt		'4
					strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
					.Col = C_AcctCd		'5
					strVal = strVal & Trim(.Text) & parent.gRowSep	
					        
					lGrpCnt = lGrpCnt + 1
			End Select							        
		Next
	End With
	
	frm1.txtMaxRows1.value = lGrpCnt-1											'Spread Sheet의 변경된 최대갯수 
	frm1.txtSpread2.value =  strDel & strVal									'Spread Sheet 내용을 저장    
				
    lGrpCnt = 1
    strVal = ""
    strDel = ""    

	ggoSpread.Source = frm1.vspdData3
    With frm1.vspdData3   
		For lngRows = 1 To .MaxRows
			.Row = lngRows
			.Col = 0
				
			Select Case .Text
				Case ggoSpread.DeleteFlag

				Case Else
					strVal = strVal & "C" & parent.gColSep & lngRows & parent.gColSep				'C=Create, Sheet가 2개 이므로 구별 
					.Col = 1
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col =  2
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = 5
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = 3
					strVal = strVal & Trim(.Text) & parent.gRowSep  
   					        
					lGrpCnt = lGrpCnt + 1					
			End Select
		Next
	End With
	
	With frm1
		.txtMaxRows3.value = lGrpCnt-1											'Spread Sheet의 변경된 최대갯수 
		.txtSpread3.value =  strDel & strVal									'Spread Sheet 내용을 저장 

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
    Call InitVariables()
    																			'Initializes local global variables
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData1
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData		
	ggoSpread.Source = frm1.vspdData3
	ggoSpread.ClearSpreadData

    frm1.txtAllcNo.focus	
    Call DbQuery()
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
			Call SetSpread2ColorAp()
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
    		
		strFrom = " A_CTRL_ITEM	A (NOLOCK), A_ACCT_CTRL_ASSN B (NOLOCK), A_PAYM_DC_DTL C (NOLOCK), A_PAYM_DC D (NOLOCK) "
		
		strWhere =			  " D.PAYM_NO =  " & FilterVar(UCase(.txtALLCNo.value), "''", "S") & "  "
		strWhere = strWhere & " AND D.SEQ = " & .hItemSeq.Value & " "
		strWhere = strWhere & " AND D.PAYM_NO  =  C.PAYM_NO  "
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
		Call SetSpread2ColorAp()
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
    Call SetSpread2ColorAp()
   
    lgBlnFlgChgValue = False    
End Function

'===================================== DisableRefPop()  =======================================
'	Name : DisableRefPop()
'	Description :
'====================================================================================================
Sub DisableRefPop()
	IF lgIntFlgMode = parent.OPMD_UMODE Then
		RefPop.innerHTML="<font color=""#777777"">선급금정보</font>"
	ELse 
		RefPop.innerHTML="<A href=""vbscript:OpenRefPrePaymNo()"">선급금정보</A>"
	End if

END sub


'=======================================================================================================
'   Function Name : chkAllcDate
'   Function Desc : 
'=======================================================================================================
Function chkAllcDate()
	Dim intI
	
	chkAllcDate = True
	With frm1
		For intI = 1 To .vspdData1.Maxrows
			.vspdData1.Row = intI
			.vspdData1.Col = C_ApDt
			If CompareDateByFormat(.vspdData1.Text,.txtAllcDt.Text,"채무일자",.txtAllcDt.Alt, _
		    	               "970025",.txtAllcDt.UserDefinedFormat,parent.gComDateType, true) = False Then
			   .txtAllcDt.focus
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
	Dim dblToApAmt			'채무액 
	Dim dblToApRemAmt		'채무잔액 
	Dim dblToApClsAmt		'반제금액 
	Dim dblToApDcAmt		'할인금액 
	Dim dblToApDcLocAmt		'할인금액(자국)

	With frm1
		dblToApAmt = FncSumSheet1(.vspddata1,C_ApAmt, 1, .vspddata1.MaxRows, False, -1, -1, "V")
		dblToApRemAmt = FncSumSheet1(.vspddata1,C_ApRemAmt, 1, .vspddata1.MaxRows, False, -1, -1, "V")
		dblToApClsAmt = FncSumSheet1(.vspddata1,C_ApClsAmt, 1, .vspddata1.MaxRows, False, -1, -1, "V")
		dblToApDcAmt = FncSumSheet1(.vspddata1,C_ApDcAmt, 1, .vspddata1.MaxRows, False, -1, -1, "V")
		dblToApDcLocAmt = FncSumSheet1(.vspddata1,C_ApDcLocAmt, 1, .vspddata1.MaxRows, False, -1, -1, "V")

		If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
			If lgQueryOk = False Then
				If UCase(Trim(.hApDocCur.Value)) = UCase(Trim(.txtDocCur.Value)) Then
					.txtClsAmt.text	= UNIConvNumPCToCompanyByCurrency(dblToApClsAmt,.txtDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
				Else
'					.txtClsAmt.text = "0"			
				End If
			End If				
		End If	
		
		.txtTotApAmt.text	 = UNIConvNumPCToCompanyByCurrency(dblToApAmt,.hApDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
		.txtTotApRemAmt.text = UNIConvNumPCToCompanyByCurrency(dblToApRemAmt,.hApDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
		.txtDcAmt.text	     = UNIConvNumPCToCompanyByCurrency(dblToApDcAmt,.hApDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
		.txtDcAmt2.text	     = UNIConvNumPCToCompanyByCurrency(dblToApDcAmt,.hApDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
					
		.txtDcLocAmt.text  = UNIConvNumPCToCompanyByCurrency(dblToApDcLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")
		.txtDcLocAmt2.text = UNIConvNumPCToCompanyByCurrency(dblToApDcLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")
	End With	
End Sub

'====================================================================================================
'   Event Name : txtDocCur_OnChange
'   Event Desc : 
'====================================================================================================
Sub txtDocCur_OnChange()
    lgBlnFlgChgValue = True
    If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
		Call CurFormatNumericOCX()
		Call CurFormatNumSprSheet()
		Call DoSum()
	End If
End Sub

Sub txtClsAmt_Change()
    lgBlnFlgChgValue = True	
End sub

'====================================================================================================
'	Name : CurFormatNumericOCX()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()
	With frm1
		' 선급금잔액 
		ggoOper.FormatFieldByObjectOfCur .txtBalAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 반제금액 
		ggoOper.FormatFieldByObjectOfCur .txtClsAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		
		ggoOper.FormatFieldByObjectOfCur .txtDcAmt, .hApDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 채무액 
		ggoOper.FormatFieldByObjectOfCur .txtTotApAmt, .hApDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 채무잔액 
		ggoOper.FormatFieldByObjectOfCur .txtTotApRemAmt, .hApDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 할인금액 
		ggoOper.FormatFieldByObjectOfCur .txtDcAmt2, .hApDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec		

	
	End With
End Sub

'====================================================================================================
'	Name : CurFormatNumSprSheet()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet()
	With frm1
		ggoSpread.Source = frm1.vspdData
		' 할인금액 
		ggoSpread.SSSetFloatByCellOfCur C_DcAmt,-1, .hApDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		
		ggoSpread.Source = frm1.vspdData1
		' 채무액 
		ggoSpread.SSSetFloatByCellOfCur C_ApAmt,-1, .hApDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		' 채무잔액 
		ggoSpread.SSSetFloatByCellOfCur C_ApRemAmt,-1, .hApDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		' 반제금액 
		ggoSpread.SSSetFloatByCellOfCur C_ApClsAmt,-1, .hApDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		' 할인금액 
		ggoSpread.SSSetFloatByCellOfCur C_ApDcAmt,-1, .hApDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec		
		
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
			Call SetSpread2ColorAp()									
		Case "VSPDDATA1" 
			Call ggoSpread.RestoreSpreadInf()
			Call InitSpreadSheet("A")
			Call ggoSpread.ReOrderingSpreadData()			
			Call SetSpreadLock("A")
		Case "VSPDDATA2"
			Call PrevspdData2Restore(gActiveSpdSheet)   
			Call ggoSpread.RestoreSpreadInf()
			Call InitCtrlSpread()			'관리항목 그리드 초기화 
			Call ggoSpread.ReOrderingSpreadData()
			Call SetSpread2ColorAp()  
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
    		
		strFrom = " A_CTRL_ITEM	A (NOLOCK), A_ACCT_CTRL_ASSN B (NOLOCK), A_PAYM_DC_DTL C (NOLOCK), A_PAYM_DC D (NOLOCK) "
		
		strWhere =			  " D.PAYM_NO =  " & FilterVar(UCase(.txtALLCNo.value), "''", "S") & "  "
		strWhere = strWhere & " AND D.SEQ = " & strItemSeq & " "
		strWhere = strWhere & " AND D.PAYM_NO  =  C.PAYM_NO  "
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
'   Event Name : vspddata1_onfocus
'   Event Desc :
'=======================================================================================================
Sub  vspddata_onfocus()
    If lgIntFlgMode <> parent.OPMD_UMODE Then    
        Call SetToolbar("1110111100001111")                                     '버튼 툴바 제어 
    Else                 
        Call SetToolbar("1111111100001111")                                     '버튼 툴바 제어 
    End If  
End Sub

'=======================================================================================================
'   Event Name : vspddata2_onfocus
'   Event Desc :
'=======================================================================================================
Sub  vspddata2_onfocus()

End Sub

'==========================================================================================
'   Event Name : vspddata1_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspddata1_Click(ByVal Col, ByVal Row)
	Call SetPopUpMenuItemInf("0101111111")

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
'   Event Name : vspddata2_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspddata_Click(ByVal Col, ByVal Row)
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
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspddata1_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

Sub vspddata_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP1C" Then
		gMouseClickStatus = "SP1CR"
	End If
End Sub

'=======================================================================================================
'   Event Name : vspddata1_LeaveCell
'   Event Desc :
'=======================================================================================================
Sub  vspddata_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If Row <> NewRow And NewRow > 0 Then
        With frm1        
            .vspddata.Row = NewRow
            .vspddata1.Col = C_ApNo
            
            .vspddata.Col = C_ItemSeq
            .hItemSeq.value = .vspddata.Text
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.ClearSpreadData
        End With
        
        frm1.vspddata.Col = 0
        If frm1.vspddata.Text = ggoSpread.DeleteFlag Then
			Exit Sub
        End if
        
        lgCurrRow = NewRow
        
        Call DbQuery2(lgCurrRow)
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
'   Event Name : vspddata1_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub  vspddata_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1

	With frm1.vspddata
        ggoSpread.Source = frm1.vspddata
       
        If Row > 0 And Col = C_AcctPB Then
            .Col = Col - 1
            .Row = Row
            
            Call OpenPopup(.Text, 4)
        End If    
    End With
End Sub

'======================================================================================================
'   Event Name :vspddata1_EditChange
'   Event Desc :
'=======================================================================================================
Sub  vspddata1_EditChange(ByVal Col , ByVal Row )
                
End Sub

'======================================================================================================
'   Event Name : vspddata_Change
'   Event Desc :
'=======================================================================================================
Sub  vspddata1_Change(ByVal Col, ByVal Row )
	Dim ApAmt
	Dim ClsAmt
	Dim DcAmt
	Dim dblTotClsAmt
	Dim dblTotDcAmt

	With frm1
		ggoSpread.Source = .vspddata1
		ggoSpread.UpdateRow Row
    
		.vspddata1.Row = Row
		.vspddata1.Col = Col
    
		Select Case Col
			Case C_ApClsAmt
				.vspdData1.Col = C_ApAmt
				ApAmt = .vspdData1.Text
				.vspdData1.Col = C_ApClsAmt
				ClsAmt = .vspdData1.Text

				If (UNICDbl(ApAmt) > 0 And UNICDbl(ClsAmt) < 0) Or (UNICDbl(ApAmt) < 0 And UNICDbl(ClsAmt) > 0) Then
					.vspdData1.Col = C_ApClsAmt
					.vspdData1.Text = UNIConvNumPCToCompanyByCurrency(UNICDbl(.vspdData1.Text) * (-1),.hApDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
				End If
				
				dblTotClsAmt = FncSumSheet1(.vspdData1,C_ApClsAmt , 1, .vspdData1.MaxRows, False, -1, -1, "V")
							
				If UCase(Trim(.hApDocCur.Value)) = UCase(Trim(.txtDocCur.Value)) Then
					.txtClsAmt.text = UNIConvNumPCToCompanyByCurrency(dblTotClsAmt,.txtDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
				End If
				
				If UCase(Trim(.hApDocCur.Value)) <> UCase(parent.gCurrency) Then			
					.vspdData1.Col = C_ApClsLocAmt
					.vspdData1.Row = .vspdData1.ActiveRow				
					.vspdData1.Text = ""
				End If			
			Case C_ApDcAmt
				.vspdData1.Col = C_ApAmt
				ApAmt = .vspdData1.Text
				.vspdData1.Col = C_ApDcAmt
				DcAmt = .vspdData1.Text									

				If (UNICDbl(ApAmt) > 0 And UNICDbl(DcAmt) < 0) Or (UNICDbl(ApAmt) < 0 And UNICDbl(DcAmt) > 0) Then
					.vspdData1.Col = C_ApDcAmt
					.vspdData1.Text = UNIConvNumPCToCompanyByCurrency(UNICDbl(.vspdData1.Text) * (-1),.hApDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
				End If		  
				
				dblTotDcAmt = FncSumSheet1(.vspdData1,C_ApDcAmt , 1, .vspdData1.MaxRows, False, -1, -1, "V")
				.txtDcAmt.text = UNIConvNumPCToCompanyByCurrency(dblTotDcAmt ,.hApDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
				
				If UCase(Trim(.hApDocCur.Value)) <> UCase(parent.gCurrency) Then
					.vspdData1.Col = C_ApDcLocAmt
					.vspdData1.Row = .vspdData1.ActiveRow
					.vspdData1.Text = ""							
				End If
				
				.txtDcAmt2.text = UNIConvNumPCToCompanyByCurrency(dblTotDcAmt ,.hApDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")			
			Case  C_ApClsLocAmt, C_ApDcLocAmt
				Call DoSum()
		End Select
	End With
End Sub

'======================================================================================================
'   Event Name : vspddata_Change
'   Event Desc :
'=======================================================================================================
Sub  vspddata_Change(ByVal Col, ByVal Row )
    ggoSpread.Source = frm1.vspddata
    ggoSpread.UpdateRow Row
    
    frm1.vspddata.Row = Row
    frm1.vspddata.Col = 0
    
    Select Case Col
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
'					Call SetSpread2ColorAR()
			    End If    
			End If
		Case C_DcAmt
			frm1.vspdData.Col  = C_DcLocAmt
			frm1.vspdData.text = ""			
'			Call DoSum()						 
	End Select
End Sub

'======================================================================================================
'   Event Name :vspddata1_DblClick
'   Event Desc :
'=======================================================================================================
Sub  vspddata_DblClick( ByVal Col , ByVal Row )
	If Row <=0 Then
		Exit Sub				
	End If		

    If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If
End Sub

'======================================================================================================
'   Event Name :vspddata1_DblClick
'   Event Desc :
'=======================================================================================================
Sub  vspddata1_DblClick( ByVal Col , ByVal Row )
	If Row <=0 Then
		Exit Sub				
	End If		

    If frm1.vspdData1.MaxRows = 0 Then
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

'======================================================================================================
'   Event Name :vspddata1_KeyPress
'   Event Desc :
'=======================================================================================================
Sub  vspddata1_KeyPress(KeyAscii )
     
End Sub

'======================================================================================================
'   Event Name : vspddata1_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
Sub  vspddata_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
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
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
			    <TR>
					<TD	WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/seltab_up_left.gif" width=9></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>채무반제(선급금)</font></td>
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
					<TD WIDTH=* align=right><A HREF="VBSCRIPT:OpenPopupTempGL()">결의전표</A>&nbsp;|&nbsp;<A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A>&nbsp;|&nbsp;<Span id="RefPop"><a href="vbscript:OpenRefPrePaymNo()">선급금정보</A></Span>&nbsp;|&nbsp;<a href="vbscript:OpenRefOpenAp()">채무발생정보</A></TD>								
					<TD WIDTH=10>&nbsp;</TD>					
				</TR>				
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH="100%" CLASS="Tab11">		
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD  <%=HEIGHT_TYPE_02%> WIDTH="100%"></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH="100%">
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>반제번호</TD>
									<TD CLASS="TD656" NOWRAP><INPUT NAME="txtAllcNo" ALT="반제번호" MAXLENGTH=18 STYLE="TEXT-ALIGN: left" tag ="12XXXU"><IMG align=top name=btnCalType src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON" onclick="vbscript: Call OpenPopup(frm1.txtAllcNo.value,0)"></TD>								
								</TR>						
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH="100%"></TD>
				</TR>			
				<TR>		
					<TD WIDTH="100%" HEIGHT=* VALIGN=TOP>					
						<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; OVERFLOW: auto; WIDTH: 100%" SCROLL="no">
							<TABLE <%=LR_SPACE_TYPE_60%>>																		
								<TR>
									<TD CLASS=TD5 NOWRAP>선급금번호</TD>
									<TD CLASS=TD6 NOWRAP>
									<INPUT TYPE="Text" NAME="txtPPNo" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="24" ALT="선급금번호"></TD>
									<TD CLASS=TD5 NOWRAP>출금일</TD>
									<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtPPDt" CLASS=FPDTYYYYMMDD tag="24" Title="FPDATETIME" ALT="출금일"></OBJECT>');</SCRIPT>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>반제일</TD>
									<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtAllcDt" CLASS=FPDTYYYYMMDD tag="22" Title="FPDATETIME" ALT="반제일"></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS=TD5 NOWRAP>지급처</TD>
									<TD CLASS=TD6 NOWRAP>
									<INPUT TYPE="Text" NAME="txtBpCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="24" ALT="지급처"> <INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24" ALT="지급처명"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>부서</TD>
									<TD CLASS=TD6 NOWRAP>
									<INPUT TYPE=TEXT NAME="txtDeptCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag=24" ALT="부서"> <INPUT TYPE=TEXT NAME="txtDeptNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="24" ALT="부서명"></TD>
									<TD CLASS=TD5 NOWRAP>사업장</TD>
									<TD CLASS=TD6 NOWRAP COLSPAN=2>
									<INPUT TYPE=TEXT NAME="txtBizCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag=24" ALT="사업장"> <INPUT TYPE=TEXT NAME="txtBizNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="24" ALT="사업장명"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>계좌번호</TD>
									<TD CLASS=TD6 NOWRAP>
									<INPUT  TYPE=TEXT NAME="txtBankAcct" SIZE=30 MAXLENGTH=30 STYLE="TEXT-ALIGN: left" tag="24" ALT="계좌번호"></TD>
									<TD CLASS=TD5 NOWRAP>어음번호</TD>
									<TD CLASS=TD6 NOWRAP>
									<INPUT TYPE=TEXT NAME="txtCheckNo" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="24" ALT="어음번호"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>거래통화</TD>
									<TD CLASS=TD6 NOWRAP>
									<INPUT TYPE=TEXT NAME="txtDocCur" SIZE=10 MAXLENGTH=4 tag="24" STYLE="TEXT-ALIGN: left" ALT="거래통화"></TD>
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
									<TD CLASS=TD5 NOWRAP>선급금잔액</TD>
									<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtBalAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="선급금잔액" tag="24X2"></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS=TD5 NOWRAP>선급금잔액(자국통화)</TD>
									<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtBalLocAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="선급금잔액(자국통화)" tag="24X2"></OBJECT>');</SCRIPT>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>반제금액</TD>
									<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtClsAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="반제금액" tag="22X2"></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS=TD5 NOWRAP>반제금액(자국통화)</TD>
									<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtClsLocAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="반제금액(자국통화)" tag="24X2"></OBJECT>');</SCRIPT>
									</TD> 
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>할인금액</TD>
									<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtDcAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="할인금액" tag="24X2" ></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS=TD5 NOWRAP>할인금액(자국통화)</TD>
									<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtDcLocAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="할인금액(자국통화)" tag="24X2" ></OBJECT>');</SCRIPT>
									</TD>
								</TR>								
								<TR>
									<TD CLASS="TD5" NOWRAP>비고</TD>
									<TD CLASS="TD6" NOWRAP COLSPAN=3><INPUT TYPE=TEXT NAME="txtAllcDesc" SIZE=90 MAXLENGTH=128 tag="21XXX" ALT="비고"></TD>
								</TR>						
								<TR>								
									<TD HEIGHT="100%" WIDTH="100%" COLSPAN=6>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" tag="2" TITLE="SPREAD" name=vspdData1 width="100%" > <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
									</TD>								
								</TR>
								<TR>
									<TD  COLSPAN="4">
									    <TABLE <%=LR_SPACE_TYPE_60%>>
									        <TR>									
												<TD CLASS=TD5 NOWRAP>채무액</TD>
												<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTotApAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="채무액" tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>
												<TD class=TD5 STYLE="WIDTH : 0px;"></TD>									
												<TD CLASS=TD5 NOWRAP>채무잔액</TD>
												<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTotApRemAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="채무잔액" tag="24X2" id=OBJECT2></OBJECT>');</SCRIPT></TD>
									  	    </TR>
									    </TABLE>
								    </TD>											
								</TR>
							</TABLE>
						</DIV>


						<DIV ID="TabDiv"  SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%>>				
								<TR HEIGHT="60%">
									<TD WIDTH="100%" Colspan=4>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" tag="2" TITLE="SPREAD" name=vspdData width="100%" > <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
									</TD>
								</TR>
								<TR>
									<TD COLSPAN=4>
									    <TABLE <%=LR_SPACE_TYPE_20%>>
										    <TR>								
												<TD CLASS=TD5 NOWRAP>할인금액</TD>
												<TD CLASS=TD6 NOWRAP>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtDcAmt2" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="할인금액" tag="24X2" ></OBJECT>');</SCRIPT></TD>
												<TD class=TD5 STYLE="WIDTH : 0px;"></TD>																			
												<TD CLASS=TD5 NOWRAP>할인금액(자국)</TD>
												<TD CLASS=TD6 NOWRAP>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtDcLocAmt2" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="할인금액(자국)" tag="24X2" ></OBJECT>');</SCRIPT></TD>
											</TR>
										</TABLE>
									</TD>											
								</TR>
								<TR HEIGHT="40%">
									<TD WIDTH="100%" Colspan=4>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" tag="2" TITLE="SPREAD" name=vspdData2 width="100%" > <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
									</TD>
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
			<IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>	
</TABLE>
<TEXTAREA Class=hidden name=txtSpread	 tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA Class=hidden name=txtSpread1	 tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA Class=hidden name=txtSpread2	 tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA Class=hidden name=txtSpread3	 tag="24" TABINDEX="-1"></TEXTAREA>
<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT TYPE=hidden CLASSID=<%=gCLSIDFPSPD%> HEIGHT=0 name=vspdData3 tag="2" width="100%" TABINDEX="-1" id=OBJECT3> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
<INPUT TYPE=hidden NAME="txtMode"		 tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtUpdtUserId"  tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows" 	 tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows1"	 tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtFlgMode"	 tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="htxtAllcNo"	 tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hItemSeq"		 tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hAcctCd"		 tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows3"	 tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hApDocCur"		 tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txthAuthBizAreaCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthSubInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthUsrID"		tag="24" Tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>

