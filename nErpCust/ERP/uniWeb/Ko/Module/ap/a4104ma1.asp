
<%@ LANGUAGE="VBSCRIPT" %>

<!--
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : bank Register
'*  3. Program ID           : a4101ma
'*  4. Program Name         : 출금등록 
'*  5. Program Desc         :
'*  6. Complus List         : 
'*  7. Modified date(First) : 2000/03/30
'*  8. Modified date(Last)  : 2002/11/14
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
<!-- '#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################
'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="javascript"   SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ag/AcctCtrl.vbs">					  </SCRIPT>
<SCRIPT LANGUAGE=vbscript>

Option Explicit																		'☜: indicates that All variables must be declared in advance
<!-- #Include file="../../inc/lgvariables.inc" -->	
'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Const BIZ_PGM_QRY_ID = "a4104mb1.asp"												'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID = "a4104mb2.asp"												'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_DEL_ID =  "a4104mb3.asp"
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
Dim C_AcctNmAp
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

Dim  strMode
Dim  IsOpenPop	
Dim  gSelframeFlg


Dim	lgFormLoad
Dim	lgQueryOk					' Queryok여부 (loc_amt =0 check)
Dim lgstartfnc

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




'========================================================================================================= 
' Name : initSpreadPosVariables()
' Description : 그리드(스프래드) 컬럼 관련 변수 초기화 
'========================================================================================================= 
Sub initSpreadPosVariables(ByVal pvSpdNo)
	Select Case UCase(Trim(pvSpdNo))
		Case "A"
			C_ApNo        = 1
			C_ApDt        = 2
			C_ApDueDt     = 3
			C_DocCur      = 4
			C_ApAmt       = 5
			C_ApRemAmt    = 6
			C_ApClsAmt    = 7
			C_ApClsLocAmt = 8
			C_ApDcAmt     = 9
			C_ApDcLocAmt  = 10
			C_ApClsDesc   = 11
			C_ApAcctCd    = 12
			C_AcctNmAp    = 13							
			C_BizCd       = 14
			C_BizNm       = 15
		Case "B"
			C_ItemSeq = 1
			C_AcctCd = 2
			C_AcctPB = 3
			C_AcctNm = 4
			C_DcAmt = 5
			C_DcLocAmt = 6
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
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    
	lgstartfnc = False
	lgFormLoad = True
	lgQueryOk  = False    
    lgSortKey  = 1
End Sub

'========================================================================================================= 
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub  SetDefaultVal()
	frm1.txtAllcDt.text     = UniConvDateAToB("<%=dtToday%>", parent.gServerDateFormat,parent.gDateFormat)
	frm1.txtDocCur.value    = parent.gCurrency
	frm1.hApDocCur.value    = parent.gCurrency
	frm1.txtXchRate.text    = 1
	frm1.hOrgChangeId.value = parent.gChangeOrgId
	frm1.txtDeptCd.value = parent.gDepart
	frm1.txtPaymLocAmt.text	= "0"
	lgBlnFlgChgValue = False
End Sub

'========================================================================================================= 
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================================= 
Sub  LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
End Sub

'========================================================================================================= 
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================= 
Sub  InitSpreadSheet(ByVal pvSpdNo)
    Call initSpreadPosVariables(pvSpdNo)  

	With frm1  
		Select Case UCase(Trim(pvSpdNo))
			Case "A"
				ggoSpread.Source = .vspddata1
				ggoSpread.Spreadinit "V20021127",,parent.gAllowDragDropSpread    
	
				frm1.vspddata1.ReDraw = False

				.vspddata1.MaxCols = C_BizNm + 1   
				.vspddata1.Col = .vspddata1.MaxCols
				.vspddata1.ColHidden = True
				.vspddata1.MaxRows = 0
	
				Call GetSpreadColumnPos(pvSpdNo)

				ggoSpread.SSSetEdit  C_ApNo,		"채무번호", 20,,,18,2		'1
				ggoSpread.SSSetEdit  C_DocCur,		"거래통화", 8 ,3'10  
				ggoSpread.SSSetFloat C_ApAmt,		"채무액", 15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec													
				ggoSpread.SSSetFloat C_ApRemAmt,	"채무잔액", 15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_ApClsAmt,	"반제금액",15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_ApClsLocAmt,	"반제금액(자국)",15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_ApDcAmt,		"할인금액",15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_ApDcLocAmt,	"할인금액(자국)",15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetEdit  C_ApClsDesc,	"비고", 20,3	'2 
				ggoSpread.SSSetEdit  C_ApAcctCd,	"계정코드", 20,3	'2
				ggoSpread.SSSetEdit  C_AcctNmAp,	"계정코드명", 20,3	'3    
				ggoSpread.SSSetEdit  C_BizCd,		"사업장", 10,3	'6
				ggoSpread.SSSetEdit  C_BizNm,		"사업장명", 20,3	'7    
				ggoSpread.SSSetDate  C_ApDt,		"채무일자",10, 2, parent.gDateFormat  
				ggoSpread.SSSetDate  C_ApDueDt,		"만기일자", 10, 2, parent.gDateFormat    
    
				frm1.vspddata1.ReDraw = True        
			Case "B"
				ggoSpread.Source = .vspddata
				ggoSpread.Spreadinit "V20021127",,parent.gAllowDragDropSpread    
		    
				frm1.vspddata.ReDraw = False	    

				.vspddata.MaxCols = C_DcLocAmt + 1 												'☜: 최대 Columns의 항상 1개 증가시킴 
				.vspddata.Col = .vspddata.MaxCols													'공통콘트롤 사용 Hidden Column
				.vspddata.ColHidden = True       
				.vspddata.MaxRows = 0

 				Call AppendNumberPlace("6","3","0")
				Call GetSpreadColumnPos(pvSpdNo)

				ggoSpread.SSSetFloat  C_ItemSeq , "NO"            ,  6,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,2,,,"0","999"    
				ggoSpread.SSSetEdit	  C_AcctCd  , "계정코드"      , 17, ,,20,2
				ggoSpread.SSSetButton C_AcctPB
				ggoSpread.SSSetEdit	  C_AcctNm  , "계정코드명"    , 50,,,20,2
				ggoSpread.SSSetFloat  C_DcAmt   , "할인금액", 20  , parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat  C_DcLocAmt, "할인금액(자국)", 20, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

				Call ggoSpread.MakePairsColumn(C_AcctCd,C_AcctPB)
		    
				frm1.vspddata.ReDraw = True        
		End Select			
	End With
    
    Call SetSpreadLock(pvSpdNo)
End Sub

'========================================================================================================= 
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================================= 
Sub  SetSpreadLock(ByVal pvSpdNo)
    With frm1
		Select Case UCase(Trim(pvSpdNo))
			Case "A"
				ggoSpread.Source = .vspdData1
				.vspdData1.Redraw = False

				ggoSpread.SpreadLock C_ApNo,-1, C_ApNo
				ggoSpread.SpreadLock C_ApAcctCd,-1, C_ApAcctCd
				ggoSpread.SpreadLock C_AcctNmAp,-1, C_AcctNmAp
				ggoSpread.SpreadLock C_BizCd,-1, C_BizCd
				ggoSpread.SpreadLock C_BizNm,-1, C_BizNm
				ggoSpread.SpreadLock C_DocCur,-1, C_DocCur    
				ggoSpread.SpreadLock C_ApDt,-1, C_ApDt
				ggoSpread.SpreadLock C_ApDueDt,-1, C_ApDueDt    
				ggoSpread.SpreadLock C_ApAmt,-1, C_ApAmt
				ggoSpread.SpreadLock C_ApRemAmt,-1, C_ApRemAmt 
				    
				ggoSpread.SSSetRequired C_ApClsAmt, -1, -1
	
				.vspdData1.Redraw = True
			Case "B"
				ggoSpread.Source = .vspdData
				.vspdData.Redraw = False

		 		ggoSpread.SpreadLock C_ItemSeq ,-1, C_ItemSeq
   				ggoSpread.SpreadLock C_AcctCd  ,-1, C_AcctCd
   				ggoSpread.SpreadLock C_AcctPB  ,-1, C_AcctPB
   				ggoSpread.SpreadLock C_AcctNm  ,-1, C_AcctNm   						
'				ggoSpread.SpreadLock C_DcLocAmt,-1, C_DcLocAmt 
					
				ggoSpread.SSSetRequired C_DcAmt,-1, -1
				.vspdData1.Redraw = True
		End Select
	End With
End Sub

'========================================================================================================= 
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================================= 
Sub  SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1    
		ggoSpread.source = .vspddata
		.vspdData.Redraw = False			

		ggoSpread.SSSetProtected C_ItemSeq , pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_AcctCd  , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_AcctNm  , pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_DcAmt   , pvStartRow, pvEndRow  

		.vspdData.Redraw = True			
    End With   
End Sub

'========================================================================================================= 
' Function Name : SetSpread2ColorAP
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================================= 
Sub  SetSpread2ColorAP()
	Dim i

    With frm1
		ggoSpread.Source = .vspdData2
		.vspdData2.ReDraw = False	 
	
		For i = 1 To .vspddata2.maxrows
			ggoSpread.SSSetProtected C_DtlSeq, i, i
			ggoSpread.SSSetProtected C_CtrlCd, i, i
			ggoSpread.SSSetProtected C_CtrlNm, i, i
			ggoSpread.SSSetRequired  C_CtrlVal, i, i			
			ggoSpread.SSSetProtected C_CtrlValNm, i, i
			.vspddata2.Row = i
			.vspddata2.Col = C_DrFg
			If (.vspddata2.text = "Y")  Or (.vspddata2.text = "DC") Or (.vspddata2.text = "C") Then
				ggoSpread.SSSetRequired C_CtrlVal, i, i	' 
			End if
		Next
		.vspdData2.ReDraw = True
    End With
End Sub
 
'========================================================================================================= 
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================================= 
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_ApNo = iCurColumnPos(1)
            C_ApDt = iCurColumnPos(2)
            C_ApDueDt = iCurColumnPos(3)
            C_DocCur = iCurColumnPos(4)
            C_ApAmt = iCurColumnPos(5)
            C_ApRemAmt = iCurColumnPos(6)
            C_ApClsAmt = iCurColumnPos(7)
            C_ApClsLocAmt = iCurColumnPos(8)
            C_ApDcAmt = iCurColumnPos(9)
            C_ApDcLocAmt = iCurColumnPos(10)
            C_ApClsDesc = iCurColumnPos(11)
            C_ApAcctCd = iCurColumnPos(12)
            C_AcctNmAp = iCurColumnPos(13)							
            C_BizCd = iCurColumnPos(14)
            C_BizNm = iCurColumnPos(15)
	   Case "B"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_ItemSeq = iCurColumnPos(1)
            C_AcctCd = iCurColumnPos(2)
            C_AcctPB = iCurColumnPos(3)
            C_AcctNm = iCurColumnPos(4)
            C_DcAmt = iCurColumnPos(5)
            C_DcLocAmt = iCurColumnPos(6)
    End Select    
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

	If gSelframeFlg <> TAB1 Then Exit Function		 
	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("A4105RA5")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "A4105RA5", "X")
		IsOpenPop = False
		Exit Function
	End If
	
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
'	Name : openglpopup
'	Description : 회계전표 POP-UP
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
	
	arrParam(0) = Trim(frm1.txtGlNo.value)									'회계전표번호 
	arrParam(1) = ""														'Reference번호 

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
	
	iCalledAspName = AskPRAspName("A5130RA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "A5130RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrParam(0) = Trim(frm1.txtTempGlNo.value)								'회계전표번호 
	arrParam(1) = ""														'Reference번호 

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
	Dim j	
	DIM X
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

				.vspddata1.Col = C_ApNo:				
				.vspddata1.text = arrRet(I - TempRow, 0)
		    	.vspddata1.Col = C_ApDt:				
				.vspddata1.text = arrRet(I - TempRow, 1)				
				.vspddata1.Col = C_ApDueDt:				
				.vspddata1.text = arrRet(I - TempRow, 2)				
				.vspddata1.Col = C_ApAmt:				
				.vspddata1.text = arrRet(I - TempRow, 3)
				.vspddata1.Col = C_ApRemAmt:				
				.vspddata1.text = arrRet(I - TempRow, 4)
				.vspddata1.Col = C_ApClsAmt:				
				.vspddata1.text = arrRet(I - TempRow, 6)				
				.vspddata1.Col = C_ApAcctCd
				.vspddata1.text = arrRet(I - TempRow, 7)				
				.vspddata1.Col = C_AcctNmAp
				.vspddata1.text = arrRet(I - TempRow, 8)				
				.vspddata1.Col = C_BizCd
				.vspddata1.text = arrRet(I - TempRow, 9)				
				.vspddata1.Col = C_BizNm
				.vspddata1.text = arrRet(I - TempRow, 10)				
				.vspddata1.Col = C_DocCur
				.vspddata1.text = arrRet(I - TempRow, 11)
				.vspddata1.Col = C_ApClsDesc:				
				.vspddata1.text = arrRet(I - TempRow, 14)				
			End If	
		Next	
		
		.hApDocCur.Value = arrRet(0, 11)
		.txtDocCur.value = arrRet(0, 11) '20051201 추가 
		.txtbpCd.Value = arrRet(0, 12)				
		.txtbpNm.Value = arrRet(0, 13)						
		
		If .txtDocCur.Value <> "" Then			
			If UCase(Trim(.txtDocCur.Value)) = UCase(Trim(arrRet(0, 11))) Then
			Else
				.txtPaymAmt.Text	= "0"
				.txtPaymLocAmt.Text = "0"
			End If
		End If	

		Call CurFormatNumSprSheet()		
		Call CurFormatNumericOCX()
		
		Call ggoOper.SetReqAttr(frm1.txtAllcDt,   "Q")
		ggoSpread.SpreadUnlock   C_ApNo    , TempRow + 1, C_ApDt, .vspddata1.MaxRows				'⊙: Unlock 컬럼 
		ggoSpread.ssSetProtected C_ApNo    , TempRow + 1, .vspddata1.MaxRows
		ggoSpread.ssSetProtected C_ApDt    , TempRow + 1, .vspddata1.MaxRows				'⊙: Protected
		ggoSpread.SSSetRequired  C_ApClsAmt, TempRow + 1, .vspddata1.MaxRows

		Call DoSum()
		.vspddata1.ReDraw = True
    End With
End Function

 '------------------------------------------  OpenDept()  ---------------------------------------
'	Name : OpenDept()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenDept(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(8)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
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
				call txtDeptCd_Onblur()  
				.txtDeptCd.focus
        End Select
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
	
	If iWhere = 1 Then
		if UCase(frm1.txtBpCd.className) = "PROTECTED" Then Exit Function
	End if
	
	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = ""							' 채권과 연계(거래처 유무)
	arrParam(2) = ""								'FrDt
	arrParam(3) = ""								'ToDt
	arrParam(4) = "S"							'B :매출 S: 매입 T: 전체 
	arrParam(5) = "PAYTO"									'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 	
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
	Dim strCd	
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	Select Case iWhere
		Case 0
			If frm1.txtAllcNo.className = "protected" Then Exit Function
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
			arrHeader(1) = "거래처명"	
		Case 3		
			If frm1.txtDocCur.className = "protected" Then Exit Function
			
			arrParam(0) = "거래통화팝업"													' 팝업 명칭 
			arrParam(1) = "B_CURRENCY"															' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtDocCur.Value)											' Code Condition
			arrParam(3) = ""																	' Name Cindition
			arrParam(4) = ""																	' Where Condition
			arrParam(5) = "거래통화"			
		
			arrField(0) = "CURRENCY"															' Field명(0)
			arrField(1) = "CURRENCY_DESC"														' Field명(1)
    
			arrHeader(0) = "거래통화"														' Header명(0)
			arrHeader(1) = "거래통화명"
		Case 4
			arrParam(0) = "계정코드팝업"													' 팝업 명칭 
			arrParam(1) = "A_ACCT,A_ACCT_GP"													' TABLE 명칭 
			arrParam(2) = Trim(strCode)							 								' Code Condition
			arrParam(3) = ""																	' Name Cindition
			arrParam(4) = "A_ACCT.GP_CD = A_ACCT_GP.GP_CD AND A_ACCT.DEL_FG <> " & FilterVar("Y", "''", "S") & " "				' Where Condition
			arrParam(5) = "계정코드"			
	
			arrField(0) = "A_ACCT.ACCT_CD"														' Field명(0)
			arrField(1) = "A_ACCT.ACCT_NM"														' Field명(1)
			arrField(2) = "A_ACCT_GP.GP_CD"														' Field명(2)
			arrField(3) = "A_ACCT_GP.GP_NM"														' Field명(3)
    
			arrHeader(0) = "계정코드"														' Header명(0)
			arrHeader(1) = "계정코드명"														' Header명(1)				
			arrHeader(2) = "그룹코드"														' Header명(2)
			arrHeader(3) = "그룹명"															' Header명(3)				
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
			If frm1.txtBankCd.className = "protected" Then Exit Function
			
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
			
			DIm strWhere
			
			If frm1.txtCheckCd.className = "protected" Then Exit Function
						
			arrParam(0) = "어음번호팝업"													' 팝업 명칭 
			arrParam(1) = "f_note a,b_biz_partner b, b_bank c"									' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtCheckCd.Value)											' Code Condition
			arrParam(3) = ""	
																			' Name Condition
			If UCase(Trim(frm1.txtDocCur.value)) = parent.gCurrency Then
				strWhere = "B_MINOR.MINOR_CD = B_CONFIGURATION.MINOR_CD And B_MINOR.MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  " _ 
								& "AND B_CONFIGURATION.SEQ_NO = 2 AND B_CONFIGURATION.REFERENCE = " & FilterVar("PP", "''", "S") & " "	_ 
								& " AND B_CONFIGURATION.MINOR_CD =  " & FilterVar(UCase(frm1.txtInputType.value), "''", "S") & ""
			ELse
				strWhere = "B_MINOR.MINOR_CD = B_CONFIGURATION.MINOR_CD and B_MINOR.MAJOR_CD = " & FilterVar("A1006", "''", "S") & " " _
								& " and B_CONFIGURATION.SEQ_NO = 2 and B_CONFIGURATION.REFERENCE = " & FilterVar("PP", "''", "S") & " " _
								& " And B_minor.minor_cd Not in ( Select  minor_cd  from b_configuration " _ 
								& " where major_cd=" & FilterVar("a1006", "''", "S") & "  and seq_no=4 and reference=" & FilterVar("NO", "''", "S") & " ) " _ 
								& " AND B_CONFIGURATION.MINOR_CD =  " & FilterVar(UCase(frm1.txtInputType.value), "''", "S") & ""
			End if
			
			If CommonQueryRs( " B_MINOR.MINOR_CD" , "B_CONFIGURATION ,  B_MINOR   " , strWhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then 
				
				Select Case UCase(Trim(Left(lgF0, Len(lgF0)-1)))
					Case "NP"
						'지급어음 
						arrParam(4) = "a.note_sts = " & FilterVar("BG", "''", "S") & "  AND a.note_fg = " & FilterVar("D3", "''", "S") & "  and a.bp_cd = b.bp_cd and a.bank_cd = c.bank_cd"					' Where Condition
						arrParam(5) = "어음번호"
							
						arrField(4) = "c.bank_nm"    	    					
					
						arrHeader(0) = "어음번호"													' Header명(0)' 조건필드의 라벨 명칭				
						arrHeader(4) = "은행"														' Header명(1)								
					Case "CP"  
						'지불구매카드 
						arrParam(0) = "지불구매카드번호팝업"										' 팝업 명칭				
						arrParam(1) = "f_note a,b_biz_partner b, b_bank c, b_card_co d "						' TABLE 명칭				
						arrParam(4) = "a.note_sts = " & FilterVar("BG", "''", "S") & "  AND a.note_fg = " & FilterVar("CP", "''", "S") & "  and a.bp_cd = b.bp_cd and a.bank_cd *= c.bank_cd and a.card_co_cd*=d.card_co_cd "						' Where Condition				
						arrParam(5) = "지불구매카드번호"					    					' 조건필드의 라벨 명칭				
					
						arrField(4) = " d.card_co_nm "    	    
									
						arrHeader(0) = "지불구매카드번호"											' Header명(0)				
						arrHeader(4) = "카드사"						
					Case "NE" ' Header명(1)								
						'배서어음 
						arrParam(4) = "a.note_sts = " & FilterVar("ED", "''", "S") & "  AND a.note_fg = " & FilterVar("D1", "''", "S") & "  and a.bp_cd = b.bp_cd and a.bank_cd = c.bank_cd"					' Where Condition
						arrParam(5) = "어음번호"					    							' 조건필드의 라벨 명칭					
					
						arrField(4) = "c.bank_nm"    	    				
					
						arrHeader(0) = "어음번호"													' Header명(0)				
						arrHeader(4) = "은행"														' Header명(1)								
									' Header명(1)								
					Case Else
						arrParam(4) = "((a.note_sts = " & FilterVar("ED", "''", "S") & "  AND a.note_fg = " & FilterVar("D1", "''", "S") & " ) or (a.note_sts = " & FilterVar("BG", "''", "S") & "  AND a.note_fg = " & FilterVar("D3", "''", "S") & " )) " 
						arrParam(4) = arrParam(4) & " and a.bp_cd = b.bp_cd and a.bank_cd = c.bank_cd"	
						arrParam(5) = "어음번호"					    							' 조건필드의 라벨 명칭				
					
						arrField(4) = "c.bank_nm"    	    				
					
						arrHeader(0) = "어음번호"													' Header명(0)				
						arrHeader(4) = "은행"														' Header명(1)								
				End Select 
			
			ENd if
			
			arrField(0) = "a.Note_no"															' Field명(0)
			arrField(1) =  "F2" & parent.gColSep & "a.Note_amt"									' Field명(1)
			arrField(2) =  "DD" & parent.gColSep & "a.Issue_dt"									' Field명(2)
			arrField(3) = "b.bp_nm"

	
			arrHeader(1) = "금액"															' Header명(1)
			arrHeader(2) = "발행일"															' Header명(1)	    
			arrHeader(3) = "거래처"															' Header명(1)

		Case 8 
			If frm1.txtInputType.className = "protected" Then Exit Function
			
			If frm1.txtDocCur.value <> "" Then
				If UCase(Trim(frm1.txtDocCur.value)) = parent.gCurrency Then
					arrParam(0) = "지급유형"														' 팝업 명칭						
					arrParam(1) = "B_MINOR,B_CONFIGURATION "
					arrParam(2) = Trim(frm1.txtInputType.value)											' Code Condition
					arrParam(3) = ""																	' Name Cindition
					arrParam(4) = "B_MINOR.MINOR_CD = B_CONFIGURATION.MINOR_CD And B_MINOR.MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  " _
								& "AND B_CONFIGURATION.SEQ_NO = 2 AND B_CONFIGURATION.REFERENCE = " & FilterVar("PP", "''", "S") & " "	' Where Condition								
					arrParam(5) = "지급유형"														' TextBox 명칭 
		
					arrField(0) = "B_MINOR.MINOR_CD"													' Field명(0)
					arrField(1) = "B_MINOR.MINOR_NM"													' Field명(1)
	    
					arrHeader(0) = "지급유형"														' Header명(0)
					arrHeader(1) = "지급유형명"														' Header명(1)		
				Else
					arrParam(0) = "지급유형"														' 팝업 명칭						
					arrParam(1) = "B_MINOR,B_CONFIGURATION "
					arrParam(2) = Trim(frm1.txtInputType.value)											' Code Condition
					arrParam(3) = ""																	' Name Cindition
					arrParam(4) = "B_MINOR.MINOR_CD = B_CONFIGURATION.MINOR_CD and B_MINOR.MAJOR_CD = " & FilterVar("A1006", "''", "S") & " " _
								& " and B_CONFIGURATION.SEQ_NO = 2 and B_CONFIGURATION.REFERENCE = " & FilterVar("PP", "''", "S") & " " _
								& " And B_minor.minor_cd Not in ( Select  minor_cd  from b_configuration " _ 
								& " where major_cd=" & FilterVar("a1006", "''", "S") & "  and seq_no=4 and reference=" & FilterVar("NO", "''", "S") & " ) "			' Where Condition								
					arrParam(5) = "지급유형"														' TextBox 명칭 
		
					arrField(0) = "B_MINOR.MINOR_CD"													' Field명(0)
					arrField(1) = "B_MINOR.MINOR_NM"													' Field명(1)
	    
					arrHeader(0) = "지급유형"														' Header명(0)
					arrHeader(1) = "지급유형명"														' Header명(1)		
				End If
			Else
				arrParam(0) = "지급유형"														' 팝업 명칭						
				arrParam(1) = "B_MINOR,B_CONFIGURATION "
				arrParam(2) = Trim(frm1.txtInputType.value)											' Code Condition
				arrParam(3) = ""																	' Name Cindition
				arrParam(4) = "B_MINOR.MINOR_CD = B_CONFIGURATION.MINOR_CD And B_MINOR.MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  " _
							& "AND B_CONFIGURATION.SEQ_NO = 2 AND B_CONFIGURATION.REFERENCE = " & FilterVar("PP", "''", "S") & " "	' Where Condition								
				arrParam(5) = "지급유형"														' TextBox 명칭 
		
				arrField(0) = "B_MINOR.MINOR_CD"													' Field명(0)
				arrField(1) = "B_MINOR.MINOR_NM"													' Field명(1)
	    
				arrHeader(0) = "지급유형"														' Header명(0)
				arrHeader(1) = "지급유형명"														' Header명(1)									
			End If				
		Case 9	'출금계정코드 
			If frm1.txtAcctCd.className = "protected" Then Exit Function    
			
			arrParam(0) = "계정코드팝업"													' 팝업 명칭 
			arrParam(1) = "A_Acct A, A_ACCT_GP B, A_JNL_ACCT_ASSN C"							' TABLE 명칭 
			arrParam(2) = ""																	' Code Condition
			arrParam(3) = ""																	' Name Cindition
			arrParam(4) = "A.GP_CD=B.GP_CD AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND A.Acct_cd=C.Acct_CD" & _
							" and C.trans_type = " & FilterVar("ap001", "''", "S") & "  and C.jnl_cd = " & FilterVar(frm1.txtInputType.Value, "''", "S")	' Where Condition
			arrParam(5) = "계정코드"														' 조건필드의 라벨 명칭 

			arrField(0) = "A.Acct_CD"															' Field명(0)
			arrField(1) = "A.Acct_NM"															' Field명(1)
    		arrField(2) = "B.GP_CD"																' Field명(2)
			arrField(3) = "B.GP_NM"																' Field명(3)
			
			arrHeader(0) = "계정코드"														' Header명(0)
			arrHeader(1) = "계정코드명"														' Header명(1)
			arrHeader(2) = "그룹코드"														' Header명(2)
			arrHeader(3) = "그룹명"															' Header명(3)		
		Case Else				    
			Exit Function
	End Select	
	
	IsOpenPop = True

	If iwhere = 0 Then
		' 권한관리 추가 
		iArrParam(5) = lgAuthBizAreaCd
		iArrParam(6) = lgInternalCd
		iArrParam(7) = lgSubInternalCd
		iArrParam(8) = lgAuthUsrID	

		iCalledAspName = AskPRAspName("A4104RA1")

		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "A4104RA1", "X")
			IsOpenPop = False'
			Exit Function
		End If	

		arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, iArrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")				
	Else
		arrRet = window.showModalDialog("../../comasp/AdoCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
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
				.txtAllcNo.focus
			Case 1	
				.txtBpCd.focus
			Case 3
				.txtDocCur.focus
			Case 4
'				Call SetActiveCell(.vspdData,C_AcctCd,.vspdData.ActiveRow ,"M","X","X")
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
				
				Call txtDocCur_OnChange()
				.txtDocCur.focus
			Case 4
				.vspddata.Col = C_AcctCd
				.vspddata.Text = arrRet(0)
				.vspddata.Col = C_AcctNm
				.vspddata.Text = arrRet(1)
			
				Call vspddata_Change(C_AcctCd, frm1.vspddata.activerow )	 ' 변경이 일어났다고 알려줌 
'				Call SetActiveCell(.vspdData,C_AcctCd,.vspdData.ActiveRow ,"M","X","X")
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
	IF iwhere  <> 0 Then
		lgBlnFlgChgValue = True
	End if	
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
	ELSE                 
		Call SetToolBar("1111111100001111")
	END IF	
	
	If gSelframeFlg = TAB2 Then Exit Function

	Call changeTabs(TAB2)	 '~~~ 두번째 Tab 
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
    Call LoadInfTB19029()																	'Load table , B_numeric_format

    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, _
							parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)   
							 
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, _
							parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
                         
    Call ggoOper.LockField(Document, "N")													'Lock  Suitable  Field
   
    Call InitSpreadSheet("A")																'Setup the Spread sheet
    Call InitSpreadSheet("B")																'Setup the Spread sheet                    
	Call InitCtrlSpread()
	Call InitCtrlHSpread()	    
    call txtInputType_OnChange()
    Call InitVariables()																	'Initializes local global variables
    Call SetDefaultVal()    
    Call ClickTab1()

    gIsTab     = "Y" 
	gTabMaxCnt = 2  	
	
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
    lgstartfnc = True    
    
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
    Call ggoOper.LockField(Document, "N") 
    call txtInputType_OnChange()
    Call InitVariables()															'Initializes local global variables
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
    Call DbQuery()																'☜: Query db data
           
    FncQuery = True																
	lgstartfnc = False	  
	    		
'	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================================
Function  FncNew() 
    Dim IntRetCD 
    Dim var1, var2, var3
    
    FncNew = False   
    lgstartfnc = True     
                                                           
	'-----------------------
    'Check previous data area
    '----------------------- 
    ggoSpread.Source = frm1.vspddata1		:    var1 = ggoSpread.SSCheckChange
    ggoSpread.Source = frm1.vspddata		:    var2 = ggoSpread.SSCheckChange
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
    Call ggoOper.ClearField(Document, "1")													'⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")													'Clear Condition Field
    Call ggoOper.LockField(Document, "N")													'Lock  Suitable  Field
    Call txtInputType_OnChange()
    Call InitVariables()
    Call SetDefaultVal()    
    
    Call txtDocCur_OnChange()																' insert by Kim Sang Joong    
    
	ggoSpread.Source = frm1.vspdData		:	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData1		:	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2		:	ggoSpread.ClearSpreadData		
	ggoSpread.Source = frm1.vspdData3		:	ggoSpread.ClearSpreadData
    
    frm1.txtAllcNo.Value = ""
    frm1.txtAllcNo.focus
    
    lgBlnFlgChgValue = False    

    FncNew = True  
    lgFormLoad = True																		' tempgldt read
    lgstartfnc = False    
	    		
'	Set gActiveElement = document.activeElement    
                                                            
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
		Exit Function																		'☜:
    End If					
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    FncDelete = True                                                        
	    		
'	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================================
Function  FncSave() 
    Dim IntRetCD 
    Dim var1, var2, var3
    
    FncSave = False                                                         

    Err.Clear																				'☜: Protect system from crashing
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
    If Not ggoSpread.SSDefaultCheck Then													'⊙: Check contents area
		Call ClickTab1()
		Exit Function
    End If 
    
    ggoSpread.Source = frm1.vspddata
    If Not ggoSpread.SSDefaultCheck Then
    	Call ClickTab2()													'⊙: Check contents area
		Exit Function
    End If     

    If Not chkAllcDate() Then
		Exit Function
    End If  
    
    If chkInputType= False Then
		Exit Function
    End If      
    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave()																			'☜: Save db data   
    
    FncSave = True                                                       
	    		
'	Set gActiveElement = document.activeElement    

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
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")	'⊙: "Will you destory previous data"
		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
	With frm1
		.vspddata1.ReDraw = False
	
		ggoSpread.Source = frm1.vspddata1	
		ggoSpread.CopyRow
		Call MaxSpreadVal(frm1.vspdData, C_ItemSeq , frm1.vspdData.ActiveRow)
		Call SetSpreadColor(frm1.vspddata1.ActiveRow, frm1.vspddata1.ActiveRow)
    
		.vspddata1.ReDraw = True
	End With
	    		
'	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function  FncCancel() 
	DIm i
	If gSelframeFlg = TAB1 Then
	
		if frm1.vspddata1.Maxrows < 1 Then Exit Function
		
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
				If Trim(frm1.vspddata1.text) = ggoSpread.InsertFlag Then 
					Call ggoOper.SetReqAttr(frm1.txtALLCDt,   "Q")
					Exit Function
				End if
				
				Call ggoOper.SetReqAttr(frm1.txtALLCDt,   "N")
			Next
				    	    
		End With   
	Else
		if frm1.vspddata.Maxrows < 1 Then Exit Function
		
		With frm1.vspddata
		    .Row = .ActiveRow
		    .Col = 0
		    If .Text = ggoSpread.InsertFlag Then
		        .Col = C_AcctCd		        
				If Len(Trim(.text)) > 0 Then  
					.Col = C_ItemSeq		        					
					DeleteHSheet(.Text)
				End If
		    End If
   
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
				End If
			End If		    
		End With
	End If  
	    		
'	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function  FncInsertRow(ByVal pvRowcnt) 
	Dim ii
	Dim imRow
	Dim iCurRowPos
	
    If gSelframeFlg <> TAB2 Then
		Call ClickTab2()		'sstData.Tab = 1
    End If
    
    If IsNumeric(Trim(pvRowcnt)) Then 
       imRow  = Cint(pvRowcnt)
    Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then
		    Exit Function
		End If
    End If
   
    With frm1.vspddata
		iCurRowPos = .ActiveRow
        .ReDraw = False
		ggoSpread.Source = frm1.vspddata
		ggoSpread.InsertRow ,imRow
		
		For ii = .ActiveRow To .ActiveRow + imRow - 1
			Call MaxSpreadVal(frm1.vspdData, C_ItemSeq , ii)
	    Next
	    .Col = 2																	' 컬럼의 절대 위치로 이동      
		.Row = 	ii - 1
		.Action = 0

		Call SetSpreadColor(iCurRowPos + 1, iCurRowPos + imRow)		
		.ReDraw = True		
    End With
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData
	    		
'	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function  FncDeleteRow() 
    Dim lDelRows
	Dim iDelRowCnt, i
    Dim DelItemSeq
    
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
	    		
'	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================================
Function  FncPrint() 
    On Error Resume Next                                               
    parent.FncPrint()
    	    		
'	Set gActiveElement = document.activeElement    

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
	    		
'	Set gActiveElement = document.activeElement    

End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function  FncExcel() 
	Call parent.FncExport(parent.C_SINGLEMULTI)
	    		
'	Set gActiveElement = document.activeElement    

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
	    		
'	Set gActiveElement = document.activeElement    

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
	Call ggoOper.ClearField(Document, "1")                                         '⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                  'Clear Condition Field
    Call ggoOper.LockField(Document, "N")                                   'Lock  Suitable  Field
    Call txtInputType_OnChange()
    Call InitVariables()                                                    'Initializes local global variables
    Call SetDefaultVal()                                                    
    
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData1
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData		
	ggoSpread.Source = frm1.vspdData3
	ggoSpread.ClearSpreadData
    
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
			strVal = strVal & "&txtMaxRows=" & .vspddata.MaxRows
			strVal = strVal & "&txtMaxRows1=" & .vspddata1.MaxRows
		Else
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001					'☜: 
			strVal = strVal & "&txtAllcNo=" & Trim(.txtAllcNo.value)				'조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
			strVal = strVal & "&lgStrPrevKeyDtl=" & lgStrPrevKeyDtl
			strVal = strVal & "&txtMaxRows=" & .vspddata.MaxRows
			strVal = strVal & "&txtMaxRows1=" & .vspddata1.MaxRows
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
	Dim strTemp

	lgQueryOk= True	

	With frm1
        '-----------------------
        'Reset variables area
        '-----------------------        
        If .vspddata.MaxRows > 0 Then       
            .vspddata.Row = 1
            .vspddata.Col = C_ItemSeq
            .hItemSeq.Value = .vspddata.Text 
            Call DbQuery2(1)
        End If

        Call ggoOper.LockField(Document, "Q")									'This function lock the suitable field
        Call ggoOper.SetReqAttr(frm1.txtDocCur,   "N")
        Call txtInputType_OnChange()
        lgIntFlgMode = parent.OPMD_UMODE										'Indicates that current mode is Update mode
        Call ClickTab1()
        frm1.txtAllcNo.focus
    End With

	Call DoSum()
	
	strTemp = frm1.txtXchRate.Text
	Call txtDocCur_OnChange()	
	frm1.txtXchRate.Text = strTemp
	
	If Frm1.vspdData1.MaxRows > 0 Then         
'		Frm1.vspdData1.Focus 
    End If
    call txtDeptCd_Onblur()  
	lgBlnFlgChgValue = False    
	lgQueryOk= False
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
					.Col = C_ItemSeq	'1
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = C_DcAmt		'2
					strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
					.Col = C_DcLocAmt		'3
					strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
					.Col = C_AcctCd		'4
					strVal = strVal & Trim(.Text) & parent.gRowSep	
					        
					lGrpCnt = lGrpCnt + 1
			End Select							        
		Next
	End With
	
	frm1.txtMaxRows1.value = lGrpCnt-1												'Spread Sheet의 변경된 최대갯수 
	frm1.txtSpread2.value =  strDel & strVal										'Spread Sheet 내용을 저장    
				
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
					strVal = strVal & "C" & parent.gColSep & lngRows & parent.gColSep				'C=Create, Sheet가 2개 이므로 구별 
					.Col =  1 'C_Seq	
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
	
	With frm1
		.txtMaxRows3.value = lGrpCnt-1												'Spread Sheet의 변경된 최대갯수 
		.txtSpread3.value =  strDel & strVal										'Spread Sheet 내용을 저장 

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
Function  DbSaveOk(ByVal AllcNo)													'☆: 저장 성공후 실행 로직 
    ggoSpread.SSDeleteFlag 1
    
    If lgIntFlgMode = parent.OPMD_CMODE Then
		frm1.txtAllcNo.value = AllcNo
	End If	  
   
	Call ggoOper.ClearField(Document, "2")											'Clear Contents  Field
    Call ggoOper.LockField(Document, "N") 
    call txtInputType_OnChange()
    Call InitVariables()															'Initializes local global variables
	
	ggoSpread.Source = frm1.vspdData		:	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData1		:	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2		:	ggoSpread.ClearSpreadData		
	ggoSpread.Source = frm1.vspdData3		:	ggoSpread.ClearSpreadData
    frm1.txtAllcNo.focus
    Call Dbquery()    
End Function




'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.4 User-defined Method Part
' Description : This part declares user-defined method
'=======================================================================================================
'*******************************************************************************************************

'=======================================================================================================
' Function Name : chkAllcDate
' Function Desc : 
'========================================================================================================
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
					IntRetCD = DisplayMsgBox("111524","X","X","X")
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



'======================================================================================================
'   Name : DoSum()
'   Desc : Sum sheet Data
'=======================================================================================================
Sub DoSum()	'출금등록 
	Dim dblToApAmt1														'채무액 
	Dim dblToApRemAmt1													'채무잔액 
	Dim dblToApClsAmt1													'반제금액 
	Dim dblToApDcAmt1													'할인금액 
	Dim dblToApDcLocAmt1												'할인금액(자국)

	With frm1.vspddata1
		dblToApAmt1 = FncSumSheet1(frm1.vspddata1,C_ApAmt, 1, .MaxRows, False, -1, -1, "V")
		dblToApRemAmt1 = FncSumSheet1(frm1.vspddata1,C_ApRemAmt, 1, .MaxRows, False, -1, -1, "V")
		dblToApClsAmt1 = FncSumSheet1(frm1.vspddata1,C_ApClsAmt, 1, .MaxRows, False, -1, -1, "V")
		dblToApDcAmt1 = FncSumSheet1(frm1.vspddata1,C_ApDcAmt, 1, .MaxRows, False, -1, -1, "V")
		dblToApDcLocAmt1 = FncSumSheet1(frm1.vspddata1,C_ApDcLocAmt, 1, .MaxRows, False, -1, -1, "V")
	
		If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
			If lgQueryOk = False Then
				If UCase(Trim(frm1.hApDocCur.Value)) = UCase(Trim(frm1.txtDocCur.Value)) Then
					frm1.txtPaymAmt.text = UNIConvNumPCToCompanyByCurrency(dblToApClsAmt1,frm1.txtDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
				Else
'					frm1.txtPaymAmt.text = "0"			
				End If
			End If				
		End If			

		frm1.txtTotApAmt1.text	  = UNIConvNumPCToCompanyByCurrency(dblToApAmt1,frm1.hApDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
		frm1.txtTotApRemAmt1.text = UNIConvNumPCToCompanyByCurrency(dblToApRemAmt1,frm1.hApDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
		frm1.txtDcAmt.text	      = UNIConvNumPCToCompanyByCurrency(dblToApDcAmt1,frm1.hApDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
		frm1.txtDcAmt2.text	      = UNIConvNumPCToCompanyByCurrency(dblToApDcAmt1,frm1.hApDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")

		frm1.txtDcLocAmt.text  = UNIConvNumPCToCompanyByCurrency(dblToApDcLocAmt1,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")
		frm1.txtDcLocAmt2.text = UNIConvNumPCToCompanyByCurrency(dblToApDcLocAmt1,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")
	End With	
End Sub


'=======================================================================================================
'   Event Name : txtInputType_OnChange()
'   Event Desc :  
'=======================================================================================================
Sub txtInputType_OnChange()
	Dim IntRetCD
	
    lgBlnFlgChgValue = True
	Call ggoOper.SetReqAttr(frm1.txtBankCd,   "Q")
	Call ggoOper.SetReqAttr(frm1.txtBankAcct, "Q")    
	Call ggoOper.SetReqAttr(frm1.txtCheckCd,  "Q")	
	
	If CommonQueryRs( "REFERENCE" , "B_CONFIGURATION " , "MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND MINOR_CD =  " & FilterVar(frm1.txtInputType.value, "''", "S") & "  AND SEQ_NO = 4 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then 
		Select Case UCase(Trim(Left(lgF0, Len(lgF0)-1)))
			Case "CS" 
				frm1.txtCheckCd.value   = ""
				frm1.txtBankCd.value   = ""
				frm1.txtBankAcct.value   = ""
				spnNoteInfo.innerHTML =  "어음번호"				
				Call ggoOper.SetReqAttr(frm1.txtBankCd,   "Q")
				Call ggoOper.SetReqAttr(frm1.txtBankAcct, "Q")
				Call ggoOper.SetReqAttr(frm1.txtCheckCd,  "Q")
			Case "DP"			' 예적금 
				spnNoteInfo.innerHTML =  "어음번호"			
				frm1.txtCheckCd.value   = ""
				Call ggoOper.SetReqAttr(frm1.txtBankCd,   "N")
				Call ggoOper.SetReqAttr(frm1.txtBankAcct, "N")
				Call ggoOper.SetReqAttr(frm1.txtCheckCd,  "Q")
			Case "NO"
				If UCase(Trim(frm1.txtInputType.value)) = "CP" Then
					spnNoteInfo.innerHTML =  "지불구매카드번호"
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
					IntRetCD = DisplayMsgBox("111524","X","X","X")  
					frm1.txtInputType.value = ""
					frm1.txtInputTypeNm.value = ""					
					Call ggoOper.SetReqAttr(frm1.txtCheckCd,  "Q")						
					Exit Sub
				End If					
			Case Else
				frm1.txtCheckCd.value   = ""
				frm1.txtBankCd.value   = ""
				frm1.txtBankAcct.value   = ""
				spnNoteInfo.innerHTML =  "어음번호"						
				Call ggoOper.SetReqAttr(frm1.txtBankCd,   "Q")
				Call ggoOper.SetReqAttr(frm1.txtBankAcct, "Q")
				Call ggoOper.SetReqAttr(frm1.txtCheckCd,  "Q")
		End Select
	End If
	
	If frm1.txtInputType.value = "" Then
		Call ggoOper.SetReqAttr(frm1.txtBankCd,   "Q")
		Call ggoOper.SetReqAttr(frm1.txtBankAcct, "Q")
		Call ggoOper.SetReqAttr(frm1.txtCheckCd,  "Q")	
	End If
	
	frm1.txtAcctCd.value = "" :	frm1.txtAcctnm.value = ""
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
		Call CurFormatNumSprSheet()		
		Call DoSum()
	End If

    If lgQueryOk = False Then
		If UCase(Trim(frm1.txtDocCur.value)) <> UCase(parent.gCurrency) Then 
			frm1.txtXchRate.Text = "0"
		Else
			frm1.txtXchRate.Text = "1"
		End If
	End If	
End Sub

'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()
	With frm1
		' 출금액 
		ggoOper.FormatFieldByObjectOfCur .txtPaymAmt,	.txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtDcAmt,		.hApDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtDcAmt2,	.hApDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtTotApAmt1, .hApDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtTotApRemAmt1, .hApDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
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
		ggoSpread.SSSetFloatByCellOfCur C_DcAmt,-1,		.hApDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		
		ggoSpread.Source = frm1.vspdData1
		' 채무액 
		ggoSpread.SSSetFloatByCellOfCur C_ApAmt,-1,		.hApDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		' 채무잔액	
		ggoSpread.SSSetFloatByCellOfCur C_ApRemAmt,-1, .hApDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		' 반제금액 
		ggoSpread.SSSetFloatByCellOfCur C_ApClsAmt,-1, .hApDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		' 할인금액 
		ggoSpread.SSSetFloatByCellOfCur C_ApDcAmt,-1,	.hApDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec

	End With
End Sub

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
	    .vspdData.Col = C_itemSeq
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
		strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD,'')),CHAR(8) "
    		
		strFrom = " A_CTRL_ITEM	A (NOLOCK), A_ACCT_CTRL_ASSN B (NOLOCK), A_PAYM_DC_DTL C (NOLOCK), A_PAYM_DC D (NOLOCK) "
		
		strWhere =			  " D.PAYM_NO = " & FilterVar(UCase(.txtALLCNo.value), "''", "S")
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
					
					strNmwhere = strColid & " =   " & FilterVar(frm1.vspddata2.text , "''", "S") & " " 
					
					If Trim(strMajorCd) <> "" Then
						strNmwhere = strNmwhere & " AND MAJOR_CD =  " & FilterVar(strMajorCd , "''", "S") & " "
					End If				 
					
					If CommonQueryRs( strColNm , strTableid ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   
						frm1.vspddata2.Col = C_CtrlValNm
						arrVal = Split(lgF0, Chr(11))  
						frm1.vspddata2.text = arrVal(0)
					End If
				End If							
				
				strVal = strVal & Chr(11) & .hItemSeq.Value 
				
				frm1.vspddata2.Col = C_DtlSeq
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.Col = C_CtrlCd
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.Col = C_CtrlNm
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.Col = C_CtrlVal
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.Col = C_CtrlPB
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.Col = C_CtrlValNm
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.Col = C_Seq
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.Col = C_Tableid
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.Col = C_Colid
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.Col = C_ColNm
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.Col = C_Datatype
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.Col = C_DataLen
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.Col = C_DRFg
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.Col = C_MajorCd
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspdData2.Col = C_MajorCd+1 				
				.vspdData2.Text = lngRows
				strVal = strVal & Chr(11) & frm1.vspddata2.text								
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
    Call SetSpread2ColorAP()
End Function




'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.5 Spread Popup method 
' Description : This part declares spread popup method
'=======================================================================================================
'*******************************************************************************************************




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
		strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD,'')),CHAR(8) "
    		
		strFrom = " A_CTRL_ITEM	A (NOLOCK), A_ACCT_CTRL_ASSN B (NOLOCK), A_PAYM_DC_DTL C (NOLOCK), A_PAYM_DC D (NOLOCK) "
		
		strWhere =			  " D.PAYM_NO = " & FilterVar(UCase(.txtALLCNo.value), "''", "S")
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
'   Event Name : vspddata_onfocus
'   Event Desc :
'=======================================================================================================
Sub  vspddata_onfocus()

End Sub

'=======================================================================================================
'   Event Name : vspddata_LeaveCell
'   Event Desc :
'=======================================================================================================
Sub  vspddata_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If Row <> NewRow And NewRow > 0 Then
        With frm1        
            .vspddata.Row = NewRow
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
'   Event Name : vspddata_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub  vspddata_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("1101111111")
    
    gMouseClickStatus = "SPC"	'Split 상태코드 
    Set gActiveSpdSheet = frm1.vspdData
    	
	If Row <= 0 Then
        ggoSpread.Source = frm1.vspddata
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If 
		Exit Sub   
    End If

    If Col <> C_AcctCd then
	    Exit Sub
    End if

	ggoSpread.Source = frm1.vspddata
	frm1.vspddata.Row = frm1.vspddata.ActiveRow	

 	frm1.vspddata.Col = C_AcctCd
	
    If Len(frm1.vspddata.Text) > 0 Then

	Else
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.ClearSpreadData		
	End If	
End Sub

Sub  vspddata1_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("0101111111")
    
    gMouseClickStatus = "SPC"	'Split 상태코드 
    Set gActiveSpdSheet = frm1.vspdData1
    	
	If Row <= 0 Then
        ggoSpread.Source = frm1.vspddata1
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If 
		Exit Sub   
    End If

    If Col <> C_AcctCd Then
	    Exit Sub
    End If

	ggoSpread.Source = frm1.vspddata1
	frm1.vspddata1.Row = frm1.vspddata1.ActiveRow	

 	frm1.vspddata1.Col = C_AcctCd
	
    If Len(frm1.vspddata1.Text) > 0 Then

	Else
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.ClearSpreadData		
	End if	
End Sub

'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

	If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub


Sub vspdData1_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

	If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
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
'   Event Name :vspddata_EditChange
'   Event Desc :
'=======================================================================================================
Sub  vspdData_EditChange(ByVal Col , ByVal Row )

End Sub

'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'=======================================================================================================
Sub  vspdData1_Change(ByVal Col, ByVal Row )
	Dim ApAmt
	Dim ApClsAmt
	Dim ApDcAmt
	Dim dblTotClsAmt
	Dim dblTotDcAmt
	
	lgBlnFlgChgValue = True
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.UpdateRow Row
    
    With frm1.vspdData1
		.Row = Row
		.Col = C_ApAmt
		ApAmt = .Text
		
		Select Case Col
			Case C_ApClsAmt
				.Col = C_ApClsAmt
				ApClsAmt = .Text
				
				If (UNICDbl(ApAmt) > 0 And parent.UNICDbl(ApClsAmt) < 0) Or (UNICDbl(ApAmt) < 0 And parent.UNICDbl(ApClsAmt) > 0) then
					.Text = UNIConvNumPCToCompanyByCurrency(UNICDbl(.Text) * (-1),frm1.hApDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
					
				End If
				
				dblTotClsAmt = FncSumSheet1(frm1.vspdData1,C_ApClsAmt , 1, .MaxRows, False, -1, -1, "V")			
			
				If UCase(Trim(frm1.hApDocCur.Value)) = UCase(Trim(frm1.txtDocCur.Value)) Then
					frm1.txtPaymAmt.text = UNIConvNumPCToCompanyByCurrency(dblTotClsAmt,frm1.txtDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
				End If
			
				If UCase(Trim(frm1.hApDocCur.Value)) <> UCase(parent.gCurrency) Then			
					.Col = C_ApClsLocAmt
					.Row = .ActiveRow				
					.Text = ""
				End If
			Case C_ApDcAmt
				.Col = C_ApDcAmt
				ApDcAmt = .Text
				
				If (UNICDbl(ApAmt) > 0 And parent.UNICDbl(ApDcAmt) < 0) Or (UNICDbl(ApAmt) < 0 And parent.UNICDbl(ApDcAmt) > 0) then
					.Text = UNIConvNumPCToCompanyByCurrency(UNICDbl(.Text) * (-1),frm1.hApDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
				End If
				
				dblTotDcAmt = FncSumSheet1(frm1.vspdData1,C_ApDcAmt , 1, .MaxRows, False, -1, -1, "V")
				frm1.txtDcAmt.text = UNIConvNumPCToCompanyByCurrency(dblTotDcAmt ,frm1.hApDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")

				If UCase(Trim(frm1.hApDocCur.Value)) <> UCase(parent.gCurrency) Then
					frm1.vspdData1.Col = C_ApDcLocAmt
					frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
					frm1.vspdData1.Text = ""							
				End If
			
				frm1.txtDcAmt2.text = UNIConvNumPCToCompanyByCurrency(dblTotDcAmt ,frm1.hApDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")			
			Case C_ApClsLocAmt, C_ApDcLocAmt
				Call DoSum()
		End Select
	End With
End Sub


Sub  vspdData_Change(ByVal Col, ByVal Row )
	lgBlnFlgChgValue = True                    'Indicates that no value changed

    ggoSpread.Source = frm1.vspddata
    ggoSpread.UpdateRow Row
    
    frm1.vspddata.Row = Row
    frm1.vspddata.Col = 0
    
	If Col = C_AcctCD and frm1.vspddata.Text = ggoSpread.InsertFlag Then
		frm1.vspddata.Col = C_ItemSeq
		frm1.hItemSeq.value = frm1.vspddata.Text
		frm1.vspddata.Col = C_AcctCd
			
		If Len(frm1.vspddata.Text) > 0 Then
			frm1.vspddata.Row = Row
			frm1.vspddata.Col = C_ItemSeq	   	
			DeleteHsheet frm1.vspddata.Text
			Call DbQuery3(Row)
			Call SetSpread2ColorAP()
		End If    
	End If 
	
	Select Case Col
		Case C_DcAmt
			frm1.vspdData.col=C_DcLocAmt
			frm1.vspdData.text=""
	End Select	
End Sub

'======================================================================================================
'   Event Name :vspddata_DblClick
'   Event Desc :
'=======================================================================================================
Sub  vspddata_DblClick( ByVal Col , ByVal Row )
    Dim iColumnName
    
    If Row <= 0 Then
		Exit Sub
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
    End If       
End Sub

Sub  vspddata1_DblClick( ByVal Col , ByVal Row )
    Dim iColumnName
   
    If Row <= 0 Then
		Exit Sub
    End If
    
    If frm1.vspdData1.MaxRows = 0 Then
		Exit Sub
    End If       
End Sub

'======================================================================================================
'   Event Name :vspddata_KeyPress
'   Event Desc :
'=======================================================================================================
Sub  vspddata_KeyPress(KeyAscii )
     
End Sub

'======================================================================================================
'   Event Name : vspdData_TopLeftChange
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

'======================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 상세내역 그리드의 (멀티)컬럼의 너비를 조절하는 경우 
'=======================================================================================================
Sub  vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.Source = frm1.vspdData
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

Sub  vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.Source = frm1.vspdData1
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")
End Sub

Sub vspdData1_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
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
    Call txtAllcDt_OnBlur()
End Sub

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
	End If
		'----------------------------------------------------------------------------------------

End Sub

'=======================================================================================================
'   Event Name : txtAllcDt_onblur()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtAllcDt_onblur()
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii, jj
	Dim arrVal1, arrVal2
	
	frm1.txtXchRate.Text = 0
   If lgstartfnc = False Then
		If lgFormLoad = True Then
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
		End If
	End If
	
	Call XchLocRate()
End Sub

'===================================== CurFormatNumSprSheet()  ======================================
'	Name : XchLocRate()
'	Description : 환율이 변경되는 Factor 가 변했을 때 수정되는 Local Amt. Setting
'====================================================================================================
Sub XchLocRate()
	Dim ii

	With frm1
		For ii = 1 To .vspdData1.MaxRows 
			.vspdData1.Row = ii	
			.vspdData1.Col = C_ApClsLocAmt	
			.vspdData1.Text = ""    		
			.vspdData1.Row = ii	
			.vspdData1.Col = C_ApDcLocAmt	
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
	End With
	
	Call DoSum()
End Sub

Sub  txtXchRate_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtPaymAmt_Change()
    lgBlnFlgChgValue = True
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
					<TD WIDTH=10>&nbsp;</TD>						
					<TD CLASS="CLSMTABP">				    
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">							
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/seltab_up_left.gif" width=9></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>출금등록</font></td>
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
					<TD WIDTH=* align=right><A HREF="VBSCRIPT:OpenPopupTempGL()">결의전표</A>&nbsp;|&nbsp;<A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A>&nbsp;|&nbsp;<A href="vbscript:OpenRefOpenAp()">채무발생정보</A></TD>								
					<TD WIDTH=*>&nbsp;</TD>												
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
									<TD CLASS="TD5" NOWRAP>출금번호</TD>
									<TD CLASS="TD656" NOWRAP><INPUT NAME="txtAllcNo" ALT="출금번호" MAXLENGTH=18 STYLE="TEXT-ALIGN: left" tag ="12XXXU"><IMG align=top name=btnCalType src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON" onclick="vbscript: Call OpenPopup(frm1.txtAllcNo.value,0)"></TD>								
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
							<TABLE <%=LR_SPACE_TYPE_60%> border=0>																				
								<TR>
									<TD CLASS=TD5 NOWRAP>출금일</TD>
									<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtAllcDt" CLASS=FPDTYYYYMMDD tag="22" Title="FPDATETIME" ALT="출금일" id=fpDateTime1></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS=TD5 NOWRAP>지급처</TD>
									<TD CLASS=TD6 NOWRAP colspan=2>
										<INPUT TYPE="Text" NAME="txtBpCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="23NXXU" ALT="지급처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:CALL OpenBp(frm1.txtBpCd.Value, 1)">
										<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24" ALT="지급처명">
									 </TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>부서</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE=TEXT NAME="txtDeptCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="22NXXU" ALT="부서"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(frm1.txtDeptCd.Value, 0)"">
										<INPUT TYPE=TEXT NAME="txtDeptNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="24" ALT="부서명">
									</TD>									
									<TD CLASS=TD5 NOWRAP>지급유형</TD>
									<TD CLASS="TD6" nowrap>
									<INPUT TYPE=TEXT NAME="txtInputType" ALT="지급유형" style="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=5 tag="22NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" style="HEIGHT: 21px; WIDTH: 16px" NAME="btnPayMethod" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtInputType.value, 8)">
													   <INPUT TYPE=TEXT NAME="txtInputTypeNm" ALT="지급유형" style="HEIGHT: 20px; WIDTH: 150px" tag="24X" ></TD>																	   
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>은행</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE="Text" NAME="txtBankCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="21NXXU" ALT="은행"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBankCd.value,5)">
										<INPUT TYPE=TEXT NAME="txtBankNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24" ALT="은행명">
									</TD>																				
									<TD CLASS=TD5 NOWRAP><span id="spnNoteInfo">어음번호</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE=TEXT NAME="txtCheckCd" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="21XXXU" ALT="어음번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtCheckCd.value,7)">
									</TD>
								</TR>																	
								<TR>	
									<TD CLASS=TD5 NOWRAP>계좌번호</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT  TYPE=TEXT NAME="txtBankAcct" SIZE=30 MAXLENGTH=30 STYLE="TEXT-ALIGN: left" tag="21XXXU" ALT="계좌번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBankAcct.value,6)">
									</TD>
									<TD CLASS=TD5 NOWRAP>출금계정코드</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtAcctCd" ALT="계정코드" MAXLENGTH="20" SIZE=10 STYLE="TEXT-ALIGN: Left" tag="22NXXU"><IMG align=top name=btnCalType onclick="vbscript:CALL OpenPopUp(frm1.txtAcctCd.value,9)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON"> 
													 <INPUT NAME="txtAcctnm" ALT="계정코드명" MAXLENGTH="20"  tag  ="24"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>결의전표번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTempGlNo" SIZE=19 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24XXXU" ALT="결의전표번호"> </TD>																						
									<TD CLASS="TD5" NOWRAP>전표번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtGlNo" SIZE=19 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24XXXU" ALT="전표번호"></TD>								
								</TR>	
								<TR>
									<TD CLASS=TD5 NOWRAP>거래통화</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE=TEXT NAME="txtDocCur" SIZE=10 MAXLENGTH=4 tag="23NXXU" STYLE="TEXT-ALIGN: left" ALT="거래통화"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCur" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript: CALL OpenPopup(frm1.txtDocCur.value,3)">
									</TD>
									<TD CLASS=TD5 NOWRAP>환율</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtXchRate" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="환율" tag="21X5Z" ></OBJECT>');</SCRIPT>
									</TD>											
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>출금액</TD>
									<TD CLASS=TD6 NOWRAP> 
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtPaymAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="출금액" tag="22X2"></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS=TD5 NOWRAP>출금액(자국통화)</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtPaymLocAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="출금액(자국통화)" tag="24X2" ></OBJECT>');</SCRIPT>
										<!--<INPUT NAME="cbSetPaymLocAmt" TYPE=CHECKBOX CLASS="RADIO" TAG="11X" ID="cbSetPaymLocAmt" VALUE = "Y">-->
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
									<TD CLASS="TD6" NOWRAP COLSPAN=3><INPUT TYPE=TEXT NAME="txtPaymDesc" SIZE=90 MAXLENGTH=128 tag="21XXX" ALT="비고"></TD>
								</TR>						
								<TR>
									<TD WIDTH="100%" HEIGHT="100%" Colspan="4">									
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" tag="2" TITLE="SPREAD" name=vspddata1 width="100%" > <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>										
									</TD>
								</TR>
								<TR>
									<TD  COLSPAN="4">
									    <TABLE <%=LR_SPACE_TYPE_60%>>
									        <TR>											
												<TD CLASS=TD5 NOWRAP >채무액</TD>
												<TD CLASS=TD6 NOWRAP ><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTotApAmt1" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="채무액" tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>									
												<TD class=TD5 STYLE="WIDTH : 0px;"></TD>
												<TD CLASS=TD5 NOWRAP >채무잔액</TD>
												<TD CLASS=TD6 NOWRAP ><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTotApRemAmt1" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="채무잔액" tag="24X2" ></OBJECT>');</SCRIPT></TD>
									  	    </TR>
									    </TABLE>
								    </TD>									
								</TR>
							</TABLE>
						</DIV>


						<DIV ID="TabDiv"  SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%>>				
								<TR HEIGHT="60%">
									<TD WIDTH="100%" colspan="12">
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" tag="2" TITLE="SPREAD" name=vspddata width="100%" > <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
									</TD>
								</TR>
								<TR>
									<TD COLSPAN=4>
									    <TABLE <%=LR_SPACE_TYPE_20%>>
										    <TR>									
									<TD CLASS=TD5 NOWRAP >할인금액</TD>
									<TD CLASS=TD6 NOWRAP >
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtDcAmt2" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="할인금액" tag="24X2" ></OBJECT>');</SCRIPT></TD>
									<TD class=TD5 STYLE="WIDTH : 0px;"></TD>										
									<TD CLASS=TD5 NOWRAP >할인금액(자국)</TD>
									<TD CLASS=TD6 NOWRAP >
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtDcLocAmt2" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="할인금액(자국)" tag="24X2" ></OBJECT>');</SCRIPT></TD>
											</TR>
										</TABLE>
									</TD>													
								</TR>
							    <TR>
									<TD HEIGHT="100%" WIDTH="100%" COLSPAN=4>
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
		<TD <%=HEIGHT_TYPE01%>></TD>
	</TR>		
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
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
<INPUT TYPE=hidden NAME="txtFlgMode"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="htxtAllcNo"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hItemSeq"			tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hAcctCd"			tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows3"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hOrgChangeId"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hApDocCur"			tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txthAuthBizAreaCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthSubInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthUsrID"		tag="24" Tabindex="-1">
<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT TYPE=hidden CLASSID=<%=gCLSIDFPSPD%> HEIGHT=0 name=vspdData3 tag="2" width="100%" TABINDEX="-1"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

