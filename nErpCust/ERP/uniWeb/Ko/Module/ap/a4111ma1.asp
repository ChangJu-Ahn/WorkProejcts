<%@ LANGUAGE="VBSCRIPT"%>

<!--
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : bank Register
'*  3. Program ID           : a4111ma.asp
'*  4. Program Name         : 채무/채권 상계 
'*  5. Program Desc         :
'*  6. Comproxy List        : ap001m
'*  7. Modified date(First) : 2000/03/30
'*  8. Modified date(Last)  : 2003/08/20
'*  9. Modifier (First)     : You So Eun
'* 10. Modifier (Last)      : Jeong Yong Kyun
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
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"     SRC="../ag/AcctCtrl.vbs">				</SCRIPT>
<SCRIPT LANGUAGE=vbscript>

Option Explicit																		'☜: indicates that All variables must be declared in advance

'=======================================================================================================
'                                               1.2 Global 변수/상수 선언  
'	.Constant는 반드시 대문자 표기.
'	.변수 표준에 따름. prefix로 g를 사용함.
'	.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=======================================================================================================
'@PGM_ID
<!-- #Include file="../../inc/lgvariables.inc" -->	
Const BIZ_PGM_QRY_ID  = "a4111mb1.asp"												'☆: Head Query 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID = "a4111mb2.asp"												'☆: Save 비지니스 로직 ASP명 
Const BIZ_PGM_DEL_ID  = "a4111mb3.asp"

Const TAB1 = 1																		'☜: Tab의 위치 
Const TAB2 = 2

Dim C_ApNo 
Dim C_AcctCd 
Dim C_AcctNm 
Dim C_ApDt 
Dim C_ApDueDt 
Dim C_ApAmt 
Dim C_ApRemAmt 
Dim C_ApClsAmt 
Dim C_ApClsLocAmt 
Dim C_ApClsDesc 

Dim C_ArNo 
Dim C_Ar_AcctCd 
Dim C_Ar_AcctNm 
Dim C_ArDt 
Dim C_ArDueDt 
Dim C_ArAmt 
Dim C_ArRemAmt 
Dim C_ArClsAmt 
Dim C_ArClsLocAmt 
Dim C_ArClsDesc 


Dim  lgStrPrevKeyDtl
Dim  lgStrPrevKey2
Dim  lgStrPrevKey3
Dim  lgCurrRow

Dim  intItemCnt					
Dim  IsOpenPop	                'Popup
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
			C_AcctCd = 2
			C_AcctNm = 3						
			C_ApDt = 4
			C_ApDueDt = 5
			C_ApAmt = 6
			C_ApRemAmt = 7
			C_ApClsAmt = 8
			C_ApClsLocAmt = 9
			C_ApClsDesc = 10
		Case "B"		
			C_ArNo = 1
			C_Ar_AcctCd = 2
			C_Ar_AcctNm = 3
			C_ArDt = 4
			C_ArDueDt = 5
			C_ArAmt = 6
			C_ArRemAmt = 7
			C_ArClsAmt = 8
			C_ArClsLocAmt = 9
			C_ArClsDesc = 10							
	End Select			
End Sub

'======================================================================================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=======================================================================================================
Sub  InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE							'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False									'Indicates that no value changed
    lgIntGrpCount = 0											'initializes Group View Size
        
    lgStrPrevKey = 0											'initializes Previous Key
    lgStrPrevKeyDtl = 0											'initializes Previous Key
    lgLngCurRows = 0											'initializes Deleted Rows Count
    gSelframeFlg = Tab1	
End Sub

'======================================================================================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=======================================================================================================
Sub  SetDefaultVal()
	frm1.txtAllcDt.text =  UniConvDateAToB("<%=dtToday%>", parent.gServerDateFormat,parent.gDateFormat)
	frm1.txtDocCur.value	= parent.gCurrency
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
        
    With frm1
		Select Case UCase(Trim(pvSpdNo))
			Case "A"
			
				ggoSpread.Source = .vspdData
				ggoSpread.SpreadInit "V20021127",,parent.gAllowDragDropSpread 

				.vspdData.Redraw = False		
				.vspdData.MaxCols = C_ApClsDesc + 1										'☜: 최대 Columns의 항상 1개 증가시킴 
				.vspdData.Col = .vspdData.MaxCols													'공통콘트롤 사용 Hidden Column
				.vspdData.ColHidden = True
				.vspdData.MaxRows = 0		    
  
				Call GetSpreadColumnPos(pvSpdNo)        
        
				ggoSpread.SSSetEdit  C_ApNo       , "채무번호"      ,20, 3		'1
				ggoSpread.SSSetEdit  C_AcctCd     , "계정"          ,20, 3	'2
				ggoSpread.SSSetEdit  C_AcctNm     , "계정명"        ,20, 3	'3    
				ggoSpread.SSSetDate  C_ApDt       , "채무일자"      ,10, 2, parent.gDateFormat  
				ggoSpread.SSSetDate  C_ApDueDt    , "만기일자"      ,10, 2, parent.gDateFormat      
				ggoSpread.SSSetFloat C_ApAmt      , "채무액"        ,15, parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_ApRemAmt   , "채무잔액"      ,15, parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_ApClsAmt   , "반제금액"      ,15, parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_ApClsLocAmt, "반제금액(자국)",15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetEdit  C_ApClsDesc  , "비고"          ,20, 3		'1    
		
				.vspdData.Redraw = True 
			Case "B"  

				ggoSpread.Source = .vspdData1
				ggoSpread.SpreadInit "V20021127",,parent.gAllowDragDropSpread 

				.vspdData1.Redraw = False		
				.vspdData1.MaxCols = C_ArClsDesc + 1										'☜: 최대 Columns의 항상 1개 증가시킴 
				.vspdData1.Col = .vspdData1.MaxCols													'공통콘트롤 사용 Hidden Column
				.vspdData1.ColHidden = True
				.vspdData1.MaxRows = 0		    
  
				Call GetSpreadColumnPos(pvSpdNo)            
    
				ggoSpread.SSSetEdit	 C_ArNo       ,"채권번호"      , 20, 3
				ggoSpread.SSSetEdit	 C_Ar_AcctCd  ,"계정"          , 20, 3    
				ggoSpread.SSSetEdit	 C_Ar_AcctNm  ,"계정명"        , 20, 3
				ggoSpread.SSSetDate	 C_ArDt       ,"채권일자"      , 10, 2, parent.gDateFormat  
				ggoSpread.SSSetDate	 C_ArDueDt    ,"채권만기일"    , 10, 2, parent.gDateFormat      
				ggoSpread.SSSetFloat C_ArAmt      ,"채권액"        , 15, parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_ArRemAmt   ,"채권잔액"      , 15, parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_ArClsAmt   ,"반제금액"      , 15, parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_ArClsLocAmt,"반제금액(자국)", 15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec  
				ggoSpread.SSSetEdit	 C_ArClsDesc  ,"비고"          , 20, 3    
	
				.vspdData1.Redraw = True 
		End Select				
    End With			
    
    Call SetSpreadLock(pvSpdNo)
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub  SetSpreadLock(ByVal pvSpdNo)
    With frm1
		Select Case UCase(Trim(pvSpdNo))
			Case "A"    
				ggoSpread.Source = .vspddata
				.vspddata1.ReDraw = False    

				ggoSpread.SpreadLock C_ApNo    , -1, C_ApNo    , -1
				ggoSpread.SpreadLock C_AcctCd  , -1, C_AcctCd  , -1
				ggoSpread.SpreadLock C_AcctNm  , -1, C_AcctNm  , -1
				ggoSpread.SpreadLock C_ApDt    , -1, C_ApDt    , -1
				ggoSpread.SpreadLock C_ApDueDt , -1, C_ApDueDt , -1
				ggoSpread.SpreadLock C_ApAmt   , -1, C_ApAmt   , -1
				ggoSpread.SpreadLock C_ApRemAmt, -1, C_ApRemAmt, -1

				ggoSpread.SSSetRequired  C_ArClsAmt, -1, -1

				.vspddata1.ReDraw = True   
			Case "B"
				ggoSpread.Source = .vspddata1
				.vspddata.Redraw = False   
				            
				ggoSpread.SpreadLock C_ArNo     , -1, C_ArNo     , -1
				ggoSpread.SpreadLock C_Ar_AcctCd, -1, C_Ar_AcctCd, -1
				ggoSpread.SpreadLock C_Ar_AcctNm, -1, C_Ar_AcctNm, -1
				ggoSpread.SpreadLock C_ArDt     , -1, C_ArDt     , -1
				ggoSpread.SpreadLock C_ArDueDt  , -1, C_ArDueDt  , -1
				ggoSpread.SpreadLock C_ArAmt    , -1, C_ArAmt    , -1
				ggoSpread.SpreadLock C_ArRemAmt , -1, C_ArRemAmt , -1                    

				ggoSpread.SSSetRequired  C_ArClsAmt, -1, -1

				.vspddata1.ReDraw = True   				    
		End Select				
    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub  SetSpreadColor()

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
			
			C_ApNo = iCurColumnPos(1)
			C_AcctCd = iCurColumnPos(2)
			C_AcctNm = iCurColumnPos(3)							
			C_ApDt = iCurColumnPos(4)
			C_ApDueDt = iCurColumnPos(5)
			C_ApAmt = iCurColumnPos(6)
			C_ApRemAmt = iCurColumnPos(7)
			C_ApClsAmt = iCurColumnPos(8)
			C_ApClsLocAmt = iCurColumnPos(9)
			C_ApClsDesc = iCurColumnPos(10)
		Case "B"
			ggoSpread.Source = frm1.vspdData1

			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)							

			C_ArNo = iCurColumnPos(1)
			C_Ar_AcctCd = iCurColumnPos(2)
			C_Ar_AcctNm = iCurColumnPos(3)
			C_ArDt = iCurColumnPos(4)
			C_ArDueDt = iCurColumnPos(5)
			C_ArAmt = iCurColumnPos(6)
			C_ArRemAmt = iCurColumnPos(7)
			C_ArClsAmt = iCurColumnPos(8)
			C_ArClsLocAmt = iCurColumnPos(9)
			C_ArClsDesc = iCurColumnPos(10)				
	End Select
End Sub

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

'======================================================================================================
'	Name : OpenRefOpenAr()
'	Description : Ref 화면을 call한다. 
'========================================================================================================= 
Function OpenRefOpenAr()
	Dim arrRet
	Dim arrParam(11)	
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function 
	
	iCalledAspName = AskPRAspName("A3106RA5")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "A3106RA5", "X")
		IsOpenPop = False
		Exit Function
	End If
  
	IsOpenPop = True

	arrParam(0) = frm1.txtBpCd.value				' 검색조건이 있을경우 파라미터 
	arrParam(1) = frm1.txtBpNm.value			
	arrParam(2) = frm1.txtDocCur.value		
	arrParam(3) = "M"
	arrParam(6) = frm1.txtAllcDt.text			
	arrParam(7) = frm1.txtAllcDt.alt
	
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
		Call SetRefOpenAr(arrRet)
	End If
End Function

'======================================================================================================
'	Name : OpenRefOpenAp()
'	Description : Ref 화면을 call한다. 
'========================================================================================================= 
Function OpenRefOpenAp()
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("A4105RA6")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "A4105RA6", "X")
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
	
	If arrRet(0, 0) = "" Then		
		Exit Function
	Else		
		Call SetRefOpenAp(arrRet)
	End If
End Function

'======================================================================================================
'	Name : SetRefOpenAr()
'	Description : OpenAp Popup에서 Return되는 값 setting
'======================================================================================================
Function SetRefOpenAr(Byval arrRet)
	Dim intRtnCnt, strData
	Dim TempRow, I
	DIM X
	Dim sFindFg
	
	With frm1.vspdData1
		.focus
		ggoSpread.Source = frm1.vspdData1
		.ReDraw = False	
	
		TempRow = .MaxRows												'☜: 현재까지의 MaxRows

        For I = TempRow To TempRow + Ubound(arrRet, 1) 
			sFindFg	= "N"
			For x = 1 to TempRow
				.Row = x
				.Col = C_ArNo				
				If .Text = arrRet(I - TempRow, 0) Then
					sFindFg	= "Y"
				End If
			Next
			If 	sFindFg	= "N" Then	
				.MaxRows = .MaxRows + 1
				.Row = I + 1				
				.Col = 0
				.Text = ggoSpread.InsertFlag
				
				.Col = C_ArNo												
				.text = arrRet(I - TempRow, 0)								
				.Col = C_Ar_AcctCd 												
				.text = arrRet(I - TempRow, 1)							
				.Col = C_Ar_AcctNm 												
				.text = arrRet(I - TempRow, 2)															
				.Col = C_ArDt 												
				.text = arrRet(I - TempRow, 5)											
				.Col = C_ArDueDt 												
				.text = arrRet(I - TempRow, 6)											
				.Col = C_ArAmt 												
				.text = arrRet(I - TempRow, 7)											
				.Col = C_ArRemAmt 												
				.text = arrRet(I - TempRow, 8)							
				.Col = C_ArClsAmt 												
				.text = arrRet(I - TempRow, 10)							
				.Col = C_ArClsDesc
				.text = arrRet(I - TempRow, 13)											
			End If	
		Next	
		
		frm1.txtDocCur.Value = arrRet(0, 14 )				
		frm1.txtbpCd.Value = arrRet(0, 11)				
		frm1.txtbpNm.Value = arrRet(0, 12)				
		
		ggoSpread.SpreadLock C_ArNo     , -1, C_ArNo     , -1
        ggoSpread.SpreadLock C_Ar_AcctCd, -1, C_Ar_AcctCd, -1
        ggoSpread.SpreadLock C_Ar_AcctNm, -1, C_Ar_AcctNm, -1
        ggoSpread.SpreadLock C_ArDt     , -1, C_ArDt     , -1
        ggoSpread.SpreadLock C_ArDueDt  , -1, C_ArDueDt  , -1
        ggoSpread.SpreadLock C_ArAmt    , -1, C_ArAmt    , -1
        ggoSpread.SpreadLock C_ArRemAmt , -1, C_ArRemAmt , -1                    
        
		ggoSpread.SSSetRequired  C_ArClsAmt, -1, -1
		Call ggoOper.SetReqAttr(frm1.txtALLCDt,   "Q")
		
		If frm1.txtBpCd.value <> "" Then					
			Call ggoOper.SetReqAttr(frm1.txtBpCd,   "Q")		
		Else			
			Call ggoOper.SetReqAttr(frm1.txtBpCd,   "N")		
		End If
	
		If frm1.txtDocCur.value <> "" Then					
			Call ggoOper.SetReqAttr(frm1.txtDocCur,   "Q")		
		Else			
			Call ggoOper.SetReqAttr(frm1.txtDocCur,   "N")		
		End If

		Call txtDocCur_OnChange()
		
		gSelframeFlg = Tab2
		Set gActiveSpdSheet = frm1.vspdData1
		.ReDraw = True
		
		Call DoSum1()
    End With
End Function

'======================================================================================================
'	Name : SetRefOpenAp()
'	Description : OpenAp Popup에서 Return되는 값 setting
'======================================================================================================
Function SetRefOpenAp(Byval arrRet)
	Dim intRtnCnt, strData
	Dim TempRow, I
	DIM X
	Dim sFindFg
	
	With frm1.vspdData
		.focus		
		ggoSpread.Source = frm1.vspdData
		.ReDraw = False	
	
		TempRow = .MaxRows												'☜: 현재까지의 MaxRows

		For I = TempRow To TempRow + Ubound(arrRet, 1) 
			sFindFg	= "N"
			For x = 1 To TempRow
				.Row = x
				.Col = C_ApNo				
				If .Text = arrRet(I - TempRow, 0) Then
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
				.Col = C_ApDt 												
				.text = arrRet(I - TempRow, 5)											
				.Col = C_ApDueDt 												
				.text = arrRet(I - TempRow, 6)											
				.Col = C_ApAmt 												
				.text = arrRet(I - TempRow, 8)											
				.Col = C_ApRemAmt 												
				.text = arrRet(I - TempRow, 9)	
				.Col = C_ApClsAmt 												
				.text = arrRet(I - TempRow, 11)											
 				.Col = C_ApClsDesc
 				.text = arrRet(I - TempRow, 14)								
			End If	
		Next	
		
		frm1.txtDocCur.Value = arrRet(0,7)				
		frm1.txtbpCd.Value = arrRet(0, 12)				
		frm1.txtbpNm.Value = arrRet(0, 13)				
		
        ggoSpread.SpreadLock C_ApNo    , -1, C_ApNo    , -1
        ggoSpread.SpreadLock C_AcctCd  , -1, C_AcctCd  , -1
        ggoSpread.SpreadLock C_AcctNm  , -1, C_AcctNm  , -1
        ggoSpread.SpreadLock C_ApDt    , -1, C_ApDt    , -1
        ggoSpread.SpreadLock C_ApDueDt , -1, C_ApDueDt , -1
        ggoSpread.SpreadLock C_ApAmt   , -1, C_ApAmt   , -1
        ggoSpread.SpreadLock C_ApRemAmt, -1, C_ApRemAmt, -1
        
		ggoSpread.SSSetRequired  C_ApClsAmt, -1, -1        
		
		Call ggoOper.SetReqAttr(frm1.txtALLCDt,   "Q")
		If frm1.txtBpCd.value <> "" Then					
			Call ggoOper.SetReqAttr(frm1.txtBpCd,   "Q")		
		Else			
			Call ggoOper.SetReqAttr(frm1.txtBpCd,   "N")		
		End If
	
		If frm1.txtDocCur.value <> "" Then					
			Call ggoOper.SetReqAttr(frm1.txtDocCur,   "Q")		
		Else			
			Call ggoOper.SetReqAttr(frm1.txtDocCur,   "N")		
		End If
		
		Call txtDocCur_OnChange()
		
		gSelframeFlg = Tab1
		Set gActiveSpdSheet = frm1.vspdData		
		.ReDraw = True
		
		Call DoSum()		
    End With
End Function
'------------------------------------------  OpenDept()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function
	
	If frm1.txtBpCd.className = "protected" Then Exit Function	
	
	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = ""							' 채권과 연계(거래처 유무)
	arrParam(2) = ""								'FrDt
	arrParam(3) = ""								'ToDt
	arrParam(4) = "S"							'B :매출 S: 매입 T: 전체 
	arrParam(5) = ""									'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 	
	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then	 
		Call EscPopup(iWhere)   
		Exit Function
	Else
		Call SetPopup(arrRet, iWhere)
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
	Dim iCalledAspName
	Dim IntRetCD

	If IsOpenPop = True Then Exit Function
	
	Select Case iWhere
		Case 0
			If frm1.txtClearNo.className = "protected" Then Exit Function			
		Case 1
			If frm1.txtBpCd.className = "protected" Then Exit Function			
			arrParam(0) = "거래처팝업"						' 팝업 명칭 
			arrParam(1) = "b_biz_partner"						' TABLE 명칭 
			arrParam(2) = strCode						 	    ' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = ""									' Where Condition
			arrParam(5) = "거래처"			
	
			arrField(0) = "BP_CD"								' Field명(0)
			arrField(1) = "BP_FULL_NM"							' Field명(1)
    
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
	End Select				

	iCalledAspName = AskPRAspName("A4111RA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "A4111RA1", "X")
		IsOpenPop = False
		Exit Function
	End If		

	' 권한관리 추가 
	iArrParam(5) = lgAuthBizAreaCd
	iArrParam(6) = lgInternalCd
	iArrParam(7) = lgSubInternalCd
	iArrParam(8) = lgAuthUsrID
	
	IsOpenPop = True	
	
	If iwhere = 0 Then	
		arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, iArrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")				
	Else
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
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
'   Function Name : SetPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function EscPopup(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0		
				.txtClearNo.focus
			Case 1
				.txtBpCd.focus
	    	Case 3
    			.txtDocCur.focus		    		
		End Select				
	End With
	IF iwhere  <> 0 Then
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
				.txtClearNo.value = arrRet(0)							
				.txtClearNo.focus
			Case 1
				.txtBpCd.value = arrRet(0)
				.txtBpNm.value = arrRet(1)
				.txtBpCd.focus
	    	Case 3
    			.txtDocCur.value = arrRet(0)	
    			Call txtDocCur_OnChange()	
    			.txtDocCur.focus		    		
		End Select				
	End With
	IF iwhere  <> 0 Then
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

	arrParam(0) = strCode		            '  Code Condition
   	arrParam(1) = frm1.txtAllcDt.Text
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
				.txtAllcDt.text = arrRet(3)
				Call txtDeptCd_Onblur()  
				.txtDeptCd.focus
	    End Select
	End With
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
							parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)    
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, _
							parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
                         
    Call ggoOper.LockField(Document, "N")                                       'Lock  Suitable  Field
    Call InitSpreadSheet("A")                                                        'Setup the Spread sheet
    Call InitSpreadSheet("B")                                                        'Setup the Spread sheet    
    Call InitVariables()                                                          'Initializes local global variables    
    Call SetToolbar("1110101100001111")										    '버튼 툴바 제어 

	frm1.txtClearNo.focus
    				 
	Call SetDefaultVal()
	
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
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'=======================================================================================================
Function  FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear
    
	'-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then											'This function check indispensable field
		Exit Function
    End If
    
    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then		
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'Clear Contents  Field
    Call InitVariables()														'Initializes local global variables
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData1
	ggoSpread.ClearSpreadData
    
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery()																'☜: Query db data
           
    FncQuery = True																
	    		
	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================================
Function  FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          

	'-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"X","X")               
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

	'-----------------------
    'Erase condition area
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "2")                                  'Clear Condition Field
    Call ggoOper.LockField(Document, "N")                                   'Lock  Suitable  Field
    Call InitVariables()															'Initializes local global variables
    Call SetDefaultVal()    
    Call txtDocCur_OnChange()
    
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData1
	ggoSpread.ClearSpreadData
    
    frm1.txtClearNo.Value = ""
    frm1.txtClearNo.focus

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
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
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
    Call DbDelete()															'☜: Delete db data

    FncDelete = True                                                        

	Set gActiveElement = document.activeElement    
End Function

'=======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================================
Function  FncSave() 
    Dim IntRetCD 
    Dim var1, var2
    
    FncSave = False                                                         
    
    On Error Resume Next                                                   
    Err.Clear
    
    '-----------------------
    'Precheck area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    var1 = ggoSpread.SSCheckChange
    
    ggoSpread.Source = frm1.vspdData1
    var2 = ggoSpread.SSCheckChange
    
    If lgBlnFlgChgValue = False And var1 = False And var2 = False  Then				'⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","X","X","X")								'⊙: Display Message(There is no changed data.)
        Exit Function
    End If
	
	If Not chkField(Document, "2") Then												'⊙: Check required field(Single area)
		Exit Function
    End If
    
	'-----------------------
    'Check content area
    '----------------------- 
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then										'⊙: Check contents area
		Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData1
    If Not ggoSpread.SSDefaultCheck Then										'⊙: Check contents area
		Exit Function
    End If
    
    If Not chkAllcDate Then
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
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function  FncCancel() 
	Select Case UCase(Trim(gActiveSpdSheet.Name))
		Case "VSPDDATA"
			If frm1.vspdData.Maxrows < 1 Then Exit Function
		 	ggoSpread.Source = frm1.vspdData
		 	 ggoSpread.EditUndo 
		Case "VSPDDATA1"					 	
			If frm1.vspdData1.Maxrows < 1 Then Exit Function
		 	ggoSpread.Source = frm1.vspdData1
		 	ggoSpread.EditUndo 
	End Select

	If frm1.vspdData.Maxrows < 1  And frm1.vspdData1.Maxrows < 1 Then 
		Call ggoOper.SetReqAttr(frm1.txtALLCDt,   "N")
	End If	

	Call DoSum()
	Call DoSum1()
		    		
	Set gActiveElement = document.activeElement    
End Function

'=======================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function  FncDeleteRow() 
    Dim lDelRows 
    
	Select Case UCase(Trim(gActiveSpdSheet.Name))
		Case "VSPDDATA"
			If frm1.vspdData.Maxrows < 1 Then Exit Function
		 	ggoSpread.Source = frm1.vspdData
		Case "VSPDDATA1"					 	
			If frm1.vspdData1.Maxrows < 1 Then Exit Function
		 	ggoSpread.Source = frm1.vspdData1
	End Select
	
    lDelRows = ggoSpread.DeleteRow
    
	Call DoSum()
	Call DoSum1()
		    		
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
	Dim var1, var2
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData
    var1 = ggoSpread.SSCheckChange
    
    ggoSpread.Source = frm1.vspdData1
    var2 = ggoSpread.SSCheckChange

    If lgBlnFlgChgValue = True Or var1 = True Or var2 = True Then
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
    strVal = strVal & "&txtClearNo=" & Trim(frm1.txtClearNo.value)				'☜: 삭제 조건 데이타 

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
	Call ggoOper.ClearField(Document, "1")                                  'Clear Condition Field
	Call ggoOper.ClearField(Document, "2")                                  'Clear Condition Field
    Call ggoOper.LockField(Document, "N")                                   'Lock  Suitable  Field

	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData1
	ggoSpread.ClearSpreadData

    Call InitVariables()                                                      'Initializes local global variables
    Call SetDefaultVal()
    
    frm1.txtClearNo.Value = ""
    frm1.txtClearNo.focus
End Function

'=======================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================================
Function  DbQuery() 
    Dim strVal
    
    DbQuery = False                                                             
    Call LayerShowHide(1)
    
    With frm1
		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001					'☜: 
			strVal = strVal & "&txtClearNo=" & Trim(.htxtClearNo.value)				'조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&lgStrPrevKeyDtl=" & lgStrPrevKeyDtl
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001					'☜: 
			strVal = strVal & "&txtClearNo=" & Trim(.txtClearNo.value)				'조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&lgStrPrevKeyDtl=" & lgStrPrevKeyDtl
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
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
	Call SetSpreadLock("A")
	Call SetSpreadLock("B")	

    lgIntFlgMode = parent.OPMD_UMODE												'Indicates that current mode is Update mode

    Call ggoOper.LockField(Document, "Q")									'This function lock the suitable field
    Call SetToolbar("1111101100001111")										'버튼 툴바 제어       

	Call DoSum()
	Call DoSum1()
	Call txtDocCur_OnChange()
	call txtDeptCd_Onblur()  
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
					strVal = strVal & "C" & parent.gColSep  						'☜: C=Create, Row위치 정보 
				    .Col = C_ApNo								'1
				    strVal = strVal & Trim(.Text) & parent.gColSep
				    .Col = C_AcctCd
				    strVal = strVal & Trim(.Text) & parent.gColSep
				    .Col = C_ApDt
				    strVal = strVal & Trim(UniConvDate(.Text)) & parent.gColSep		        
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

	frm1.txtMaxRows.value = lGrpCnt-1												'Spread Sheet의 변경된 최대갯수 
	frm1.txtSpread.value =   strVal													'Spread Sheet 내용을 저장 

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
					strVal = strVal & "C" & parent.gColSep  						'C=Create, Sheet가 2개 이므로 구별 
					.Col = C_ArNo	'1
					strVal = strVal & Trim(.Text) & parent.gColSep					            
					.Col = C_ArDt		'2
					strVal = strVal & Trim(UniConvDate(.Text)) & parent.gColSep
					.Col = C_Ar_AcctCd		'3
					strVal = strVal & Trim(.Text) & parent.gColSep					        
					.Col = C_ArDueDt		'4
					strVal = strVal & Trim(UniConvDate(.Text)) & parent.gColSep		
					.Col = C_ArClsAmt		'4
					strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gColSep		
					.Col = C_ArClsLocAmt		'4
					strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gColSep	
					.Col = C_ArClsDesc		'4
					strVal = strVal & Trim(.Text) & parent.gRowSep						
						        
					lGrpCnt = lGrpCnt + 1
			End Select						
		Next
	End With

	With frm1
		.txtMaxRows1.value = lGrpCnt-1												'Spread Sheet의 변경된 최대갯수 
		.txtSpread1.value =  strVal													'Spread Sheet 내용을 저장 

		'권한관리추가 start
		.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
		.txthInternalCd.value =  lgInternalCd
		.txthSubInternalCd.value = lgSubInternalCd
		.txthAuthUsrID.value = lgAuthUsrID		
		'권한관리추가 end
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										'저장 비지니스 ASP 를 가동 
        
    DbSave = True                                                           
End Function

'=======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================================
Function  DbSaveOk(ByVal ClearNo)													'☆: 저장 성공후 실행 로직 
    ggoSpread.SSDeleteFlag 1
    
    If lgIntFlgMode = parent.OPMD_CMODE Then
		  frm1.txtClearNo.value = ClearNo
	End If	  
	
	Call ggoOper.ClearField(Document, "2")											'Clear Contents  Field
    Call InitVariables()															'Initializes local global variables
    
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData1
	ggoSpread.ClearSpreadData
    
	Call DbQuery()	
End Function




'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.4 User-defined Method Part
' Description : This part declares user-defined method
'=======================================================================================================
'*******************************************************************************************************





'======================================================================================================
'   Name : DoSum()
'   Desc : Sum sheet Data(AP)
'=======================================================================================================
Sub DoSum()		
	Dim dblToApAmt			
	Dim dblToApRemAmt		
	Dim dblToApClsAmt		
	Dim dblToApClsLocAmt	

	With frm1.vspdData
		dblToApAmt = FncSumSheet1(frm1.vspdData,C_ApAmt, 1, .MaxRows, False, -1, -1, "V")
		dblToApRemAmt = FncSumSheet1(frm1.vspdData,C_ApRemAmt, 1, .MaxRows, False, -1, -1, "V")
		dblToApClsAmt = FncSumSheet1(frm1.vspdData,C_ApClsAmt, 1, .MaxRows, False, -1, -1, "V")
		dblToApClsLocAmt = FncSumSheet1(frm1.vspdData,C_ApClsLocAmt, 1, .MaxRows, False, -1, -1, "V")
	
		If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
			frm1.txtTotApAmt.text	= UNIConvNumPCToCompanyByCurrency(dblToApAmt,frm1.txtDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
			frm1.txtTotApRemAmt.text	= UNIConvNumPCToCompanyByCurrency(dblToApRemAmt,frm1.txtDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
			frm1.txtTotApClsAmt.text	= UNIConvNumPCToCompanyByCurrency(dblToApClsAmt,frm1.txtDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
		End If	
        frm1.txtTotApClsLocAmt.text = UNIConvNumPCToCompanyByCurrency(dblToApClsLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")
	End With	
End Sub 

'======================================================================================================
'   Name : DoSum()
'   Desc : Sum sheet Data(AR)
'=======================================================================================================
Sub DoSum1()
	Dim dblToArAmt			
	Dim dblToArRemAmt		
	Dim dblToArClsAmt		
	Dim dblToArClsLocAmt	

	With frm1.vspdData1
		dblToArAmt = FncSumSheet1(frm1.vspdData1,C_ArAmt, 1, .MaxRows, False, -1, -1, "V")
		dblToArRemAmt = FncSumSheet1(frm1.vspdData1,C_ArRemAmt, 1, .MaxRows, False, -1, -1, "V")
		dblToArClsAmt = FncSumSheet1(frm1.vspdData1,C_ArClsAmt, 1, .MaxRows, False, -1, -1, "V")
		dblToArClsLocAmt = FncSumSheet1(frm1.vspdData1,C_ArClsLocAmt, 1, .MaxRows, False, -1, -1, "V")
		
		If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
			frm1.txtTotArAmt.text	= UNIConvNumPCToCompanyByCurrency(dblToArAmt,frm1.txtDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
			frm1.txtTotArRemAmt.text	= UNIConvNumPCToCompanyByCurrency(dblToArRemAmt,frm1.txtDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
			frm1.txtTotArClsAmt.text	= UNIConvNumPCToCompanyByCurrency(dblToArClsAmt,frm1.txtDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
		End If	
        frm1.txtTotArClsLocAmt.text = UNIConvNumPCToCompanyByCurrency(dblToArClsLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")
	End With	
End Sub

'=======================================================================================================
'   Function Name : chkAllcDate
'   Function Desc : 
'=======================================================================================================
Function chkAllcDate()
	Dim intI
	
	chkAllcDate = True
	With frm1
		For intI = 1 To .vspdData.MaxRows
			.vspdData.Row = intI
			.vspdData.Col = C_ApDt
			If CompareDateByFormat(.vspdData.Text,.txtAllcDt.Text,"채무일자",.txtAllcDt.Alt, _
		    	               "970025",.txtAllcDt.UserDefinedFormat,parent.gComDateType, true) = False Then
			   .txtAllcDt.focus
			   chkAllcDate = False
			   Exit Function
			End If
		Next
		
		For intI = 1 To .vspdData1.MaxRows
			.vspdData1.Row = intI
			.vspdData1.Col = C_ArDt
			If CompareDateByFormat(.vspdData1.Text,.txtAllcDt.Text,"채권일자",.txtAllcDt.Alt, _
		    	               "970025",.txtAllcDt.UserDefinedFormat,parent.gComDateType, true) = False Then
			   .txtAllcDt.focus
			   chkAllcDate = False
			   Exit Function
			End If
		Next
	End With
End Function

'======================================================================================================
'   Name : txtDocCur_OnChange()
'   Desc : 
'=======================================================================================================
Sub txtDocCur_OnChange()
    lgBlnFlgChgValue = True
    If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
		Call CurFormatNumericOCX()
		Call CurFormatNumSprSheet()
	End If	    
End Sub

'======================================================================================================
'	Name : CurFormatNumericOCX()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX
'======================================================================================================
Sub CurFormatNumericOCX()
	With frm1
		' 채무액 
		ggoOper.FormatFieldByObjectOfCur .txtTotApAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 채무잔액 
		ggoOper.FormatFieldByObjectOfCur .txtTotApRemAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 반제금액(채무)
		ggoOper.FormatFieldByObjectOfCur .txtTotApClsAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 채권액 
		ggoOper.FormatFieldByObjectOfCur .txtTotArAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 채권잔액 
		ggoOper.FormatFieldByObjectOfCur .txtTotArRemAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 반제금액(채권)
		ggoOper.FormatFieldByObjectOfCur .txtTotArClsAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
	End With
End Sub

'======================================================================================================
'	Name : CurFormatNumSprSheet()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric Spread Sheet
'======================================================================================================
Sub CurFormatNumSprSheet()
	With frm1
		' 채무 
		ggoSpread.Source = frm1.vspdData
		' 채무액 
		ggoSpread.SSSetFloatByCellOfCur C_ApAmt,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		' 채무잔액 
		ggoSpread.SSSetFloatByCellOfCur C_ApRemAmt,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		' 반제금액 
		ggoSpread.SSSetFloatByCellOfCur C_ApClsAmt,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		
		' 채권 
		ggoSpread.Source = frm1.vspdData1
		' 채권액 
		ggoSpread.SSSetFloatByCellOfCur C_ArAmt,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		' 채권잔액 
		ggoSpread.SSSetFloatByCellOfCur C_ArRemAmt,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		' 반제금액 
		ggoSpread.SSSetFloatByCellOfCur C_ArClsAmt,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
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
	Select Case Trim(UCase(gActiveSpdSheet.Name))
		Case "VSPDDATA" 
			ggoSpread.Source = frm1.vspdData
			Call ggoSpread.RestoreSpreadInf()
			Call InitSpreadSheet("A")
			Call ggoSpread.ReOrderingSpreadData()
			Call ggoOper.SetReqAttr(frm1.txtALLCDt,   "N")
		Case "VSPDDATA1" 
			ggoSpread.Source = frm1.vspdData1
			Call ggoSpread.RestoreSpreadInf()
			Call InitSpreadSheet("B")
			Call ggoSpread.ReOrderingSpreadData()
			Call ggoOper.SetReqAttr(frm1.txtALLCDt,   "N")
	End Select
End Sub





'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.6 Spread OCX Tag Event
' Description : This part declares Spread OCX Tag Event
'=======================================================================================================
'*******************************************************************************************************




'==========================================================================================
'   Event Name : vspddata1_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspddata1_Click(ByVal Col, ByVal Row)
    gMouseClickStatus = "SPC"									'Split 상태코드 
 
	Set gActiveSpdSheet = frm1.vspdData1
	Call SetPopUpMenuItemInf("0101111111")
	
	If frm1.vspdData.Maxrows = 0 then
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
    gMouseClickStatus = "SP1C"									'Split 상태코드 
 
	Set gActiveSpdSheet = frm1.vspdData
	Call SetPopUpMenuItemInf("0101111111")	
	
	If frm1.vspdData.Maxrows = 0 then
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
		frm1.vspdData2.Maxrows = 0
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
'   Event Name :vspdData_Change
'   Event Desc :
'=======================================================================================================
Sub  vspdData_Change(ByVal Col, ByVal Row )
	Dim ApAmt
	Dim ClsAmt

	lgBlnFlgChgValue = True
	
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

			If (UNICDbl(ApAmt) > 0 And UNICDbl(ClsAmt) < 0) Or (UNICDbl(ApAmt) < 0 And UNICDbl(ClsAmt) > 0) Then
				frm1.vspdData.Col = C_ApClsAmt
				frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(UNICDbl(frm1.vspdData.Text) * (-1),frm1.txtDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
			End If
			
			frm1.vspdData.col  = C_ApClsLocAmt
			frm1.vspdData.text = ""
			Call Dosum()			
	End Select	
End Sub

'======================================================================================================
'   Event Name :vspdData1_Change
'   Event Desc :
'=======================================================================================================
Sub  vspdData1_Change(ByVal Col, ByVal Row )
	Dim ArAmt
	Dim ClsAmt

	lgBlnFlgChgValue = True
	
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.UpdateRow Row
    
    frm1.vspdData1.Row = Row
    frm1.vspdData1.Col = 0             

    Select Case Col
		Case C_ArClsAmt
			frm1.vspdData1.Col = C_ArAmt
			ArAmt = frm1.vspdData1.Text
			frm1.vspdData1.Col = C_ArClsAmt
			ClsAmt = frm1.vspdData1.Text

			If (UNICDbl(ArAmt) > 0 And UNICDbl(ClsAmt) < 0) Or (UNICDbl(ArAmt) < 0 And UNICDbl(ClsAmt) > 0) Then
				frm1.vspdData1.Col = C_ArClsAmt
				frm1.vspdData1.Text = UNIConvNumPCToCompanyByCurrency(UNICDbl(frm1.vspdData1.Text) * (-1),frm1.txtDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
				
			End If
			frm1.vspdData1.Col  = C_ArClsLocAmt
			frm1.vspdData1.text = ""			
			Call Dosum1()
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
	Call GetSpreadColumnPos("A")
End Sub

'======================================================================================================
'   Event Name :vspddata_ScriptDragDropBlock
'   Event Desc :
'=======================================================================================================
Sub  vspddata1_ScriptDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
	ggoSpread.Source = frm1.vspdData1 
	Call ggoSpread.SpreadDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
	Call GetSpreadColumnPos("B")
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
'==========================================================================================
'   Event Name : txtDeptCd_Onblur
'   Event Desc : 
'==========================================================================================
Sub txtDeptCd_Onblur()
        
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii, jj

	If Trim(frm1.txtAllcDt.Text = "") Then    
		Exit sub
    End If

    lgBlnFlgChgValue = True
	
	If Trim(frm1.txtDeptCd.value) <> "" Then
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
	End If	
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
End Sub

'=======================================================================================================
'   Event Name : txtAllcDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtAllcDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtAllcDt.Action = 7        
        Call SetFocusToDocument("M")
		Frm1.txtAllcDt.Focus         
        Call txtAllcDt_onBlur()
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 >
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/seltab_up_left.gif" width=9></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>채무/채권상계</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/seltab_up_right.gif" width=10></td>
						    </TR>
						</TABLE>
					</TD>					
					<TD WIDTH=* align=right><A HREF="VBSCRIPT:OpenPopupTempGL()">결의전표</A>&nbsp;|&nbsp;<A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A>&nbsp;|&nbsp;<a href="vbscript:OpenRefOpenAp()">채무발생정보</A>&nbsp;|&nbsp;<a href="vbscript:OpenRefOpenAr()">채권발생정보</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>			
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH="100%" CLASS="Tab11">		
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD  <%=HEIGHT_TYPE_02%> WIDTH="100%" ></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH="100%">
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>상계번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtClearNo" ALT="상계번호" MAXLENGTH=18 STYLE="TEXT-ALIGN: left" tag ="12XXXU"><IMG align=top name=btnCalType src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON" onclick="vbscript:CALL OpenPopUp(frm1.txtClearNo.Value,0)"></TD>
									<TD CLASS="TD6" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>																
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH="100%" ></TD>
				</TR>
				<TR>		
					<TD WIDTH="100%" HEIGHT=* VALIGN=TOP >
						<TABLE <%=LR_SPACE_TYPE_60%>>											
							<TR>
								<TD CLASS=TD5 NOWRAP>거래처</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBpCd" ALT="거래처" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN: left" tag="23NXXU"><IMG align=top name=btnCalType onclick="vbscript:CALL OpenBp(frm1.txtBpCd.Value, 1)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON">&nbsp;<INPUT  NAME="txtBpNm" SIZE="20" tag = "24" ></TD>								
								<TD CLASS=TD5 NOWRAP>거래통화</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDocCur" ALT="거래통화" MAXLENGTH="3" SIZE=10 STYLE="TEXT-ALIGN: Left" tag ="23XXXU"><IMG align=top name=btnCalType onclick="vbscript:CALL OpenPopUp(frm1.txtDocCur.Value,3)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON"></TD>
							</TR>							
							<TR>
								<TD CLASS=TD5 NOWRAP>상계부서</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDeptCd" ALT="상계부서" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN: left" tag="22NXXU"><IMG align=top name=btnCalType onclick="vbscript:CALL OpenDept(frm1.txtDeptCd.Value, 0)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON">&nbsp;<INPUT  NAME="txtDeptNm" SIZE="20" tag = "24" ></TD>								
								<TD CLASS=TD5 NOWRAP>상계일</TD>
								<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtAllcDt" CLASS=FPDTYYYYMMDD tag="22" Title="FPDATETIME" ALT="상계일" id=fpDateTime1></OBJECT>');</SCRIPT>
								</TD>												
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>결의전표번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTempGlNo" SIZE=19 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24XXXU" ALT="결의전표번호"> </TD>																						
								<TD CLASS=TD5 NOWRAP>전표번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtGlNo" SIZE=19 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24XXXU" ALT="전표번호"></TD>								
							</TR>								
							<TR>
								<TD CLASS=TD5 NOWRAP>비고</TD>
								<TD CLASS=TD656 NOWRAP COLSPAN=3><INPUT TYPE=TEXT NAME="txtDesc" SIZE=80 MAXLENGTH=128 tag="21XXX" ALT="비고"></TD>								
							</TR>							
							<TR HEIGHT="100%">
								<TD colspan =2>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" tag="2" TITLE="SPREAD" name=vspdData width="100%" > <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>	
								<TD colspan =2>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" tag="2" TITLE="SPREAD" name=vspdData1 width="100%" > <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>							
								<TD COLSPAN=4>
									<TABLE <%=LR_SPACE_TYPE_60%>>	
										<TR>
											<TD CLASS=TDT NOWRAP></TD>
											<TD CLASS=TDT NOWRAP>채무액</TD>
											<TD CLASS=TDT NOWRAP></TD>
											<TD CLASS=TDT NOWRAP>채무잔액</TD>	
											<TD CLASS=TDT NOWRAP></TD>
											<TD CLASS=TDT NOWRAP>채권액</TD>	
											<TD CLASS=TDT NOWRAP></TD>
											<TD CLASS=TDT NOWRAP>채권잔액</TD>																																	
										</TR>							
										<TR>
											<TD CLASS=TDT NOWRAP COLSPAN=2><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTotApAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="채무액" tag="24X2" id=OBJECT5></OBJECT>');</SCRIPT></TD>
											<TD CLASS=TDT NOWRAP COLSPAN=2><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTotApRemAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="채무잔액" tag="24X2" ></OBJECT>');</SCRIPT></TD>		
											<TD CLASS=TDT NOWRAP COLSPAN=2><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTotArAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="채권잔액" tag="24X2" id=OBJECT4></OBJECT>');</SCRIPT></TD>									
											<TD CLASS=TDT NOWRAP COLSPAN=2><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTotArRemAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="채권잔액" tag="24X2" id=OBJECT6></OBJECT>');</SCRIPT></TD>																											
										</TR>			

										<TR>	
											<TD CLASS=TDT NOWRAP></TD>										
											<TD CLASS=TDT NOWRAP>반제금액</TD>		
											<TD CLASS=TDT NOWRAP></TD>
											<TD CLASS=TDT NOWRAP>반제금액(자국)</TD>	
											<TD CLASS=TDT NOWRAP></TD>
											<TD CLASS=TDT NOWRAP>반제금액</TD>	
											<TD CLASS=TDT NOWRAP></TD>
											<TD CLASS=TDT NOWRAP>반제금액(자국)</TD>												
										</TR>															
										<TR>
											<TD CLASS=TDT NOWRAP COLSPAN=2><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTotApClsAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="반제금액" tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>								
											<TD CLASS=TDT NOWRAP COLSPAN=2><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTotApClsLocAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="반제금액(자국)" tag="24X2" id=OBJECT2></OBJECT>');</SCRIPT></TD>									
											<TD CLASS=TDT NOWRAP COLSPAN=2><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTotArClsAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="반제금액" tag="24X2" id=OBJECT3></OBJECT>');</SCRIPT></TD>									
											<TD CLASS=TDT NOWRAP COLSPAN=2><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTotArClsLocAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="반제금액(자국)" tag="24X2" id=OBJECT2></OBJECT>');</SCRIPT></TD>																				
										</TR>	
									</TABLE>
								<TD>
							</TR>
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
			<IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>		
</TABLE>
<TEXTAREA class=hidden name=txtSpread tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA class=hidden name=txtSpread1 tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="htxtClearNo" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hItemSeq" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hAcctCd" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows1" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hOrgChangeId"	tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txthAuthBizAreaCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthSubInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthUsrID"		tag="24" Tabindex="-1">
<PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

