<%@ LANGUAGE="VBSCRIPT"%>
<%Response.Expires = -1%>
<!--'**********************************************************************************************
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID           : a7102ma1
'*  4. Program Name         : 고정자산취득내역등록 
'*  5. Program Desc         : 고정자산별 취득내역을 등록,수정,삭제,조회 
'*  6. Comproxy List        : +As0021
'                             +As0029
'                             +B1a028ListMinorCode
'*  7. Modified date(First) : 2000/03/30	
'*  8. Modified date(Last)  : 2001/05/19
'*  9. Modifier (First)     : 김희정 
'* 10. Modifier (Last)      : 김희정 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/003/30 : ..........
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->

<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAMain.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAEvent.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/Cookie.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js">   </SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit 
<!-- #Include file="../../inc/lgvariables.inc" -->	

'==========================================================================================================
Const gIsShowLocal = "Y"
<%
Const gIsShowLocal = "Y"
%>

Const BIZ_PGM_QRY_ID  = "a7102mb1.asp"												'☆: Head Query 비지니스 로직 ASP명 
Const Biz_PGM_QRY_ID2 = "a7102mb4.asp"
Const BIZ_PGM_DEL_ID  = "a7102mb3.asp"												'☆: Delete 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID = "a7102mb2.asp"												'☆: Save 비지니스 로직 ASP명 

Const COM_BIZ_EXCHRATE_ID = "../../inc/GetExchRate.asp"								'☆: 환율정보 비지니스 로직 ASP명 

Const TAB1 = 1																		'☜: Tab의 위치 
Const TAB2 = 2

'''자산master
Dim C_Deptcd
Dim C_DeptPop
Dim C_DeptNm
Dim C_AcctCd
Dim C_AcctPop
Dim C_AcctNm
Dim C_AsstNo
Dim C_AsstNm
Dim C_AcqAmt
Dim C_AcqLocAmt
Dim C_AcqQty
Dim C_ResAmt
Dim C_RefNo
Dim C_Desc


'''취득상세내역 
Dim C_Seq
Dim C_RcptType								            'Spread Sheet 의 Columns 인덱스 
Dim C_RcptTypePopup
Dim C_RcptTypeNm								            'Spread Sheet 의 Columns 인덱스 
Dim C_Amt
Dim C_LocAmt
Dim C_BankAcct
Dim C_BankAcctPopup
Dim C_NoteNo
Dim C_NoteNoPopup

Const C_SHEETMAXROWS_i  = 10


Const C_SHEETMAXROWS_m = 30
'========================================================================================================= 
'DIM lgBlnFlgChgValue           ' Variable is for Dirty flag
'Dim lgIntGrpCount              ' Group View Size를 조사할 변수 
'Dim lgIntFlgMode               ' Variable is for Operation Status

Dim lgStrPrevKey_i,lgStrPrevKey_m
'========================================================================================================= 
Dim ihGridCnt                     'hidden Grid Row Count
Dim intItemCnt                    'hidden Grid Row Count
Dim lgstrConffg  
Dim dblXchRate		              'Exchange Rate 를 가지고 오기 
Dim IsOpenPop						' Popup
Dim gSelframeFlg

'========================================================================================================= 
Sub initSpreadPosVariables(ByVal pvSpdNo)
	Select Case UCase(Trim(pvSpdNo))
		Case "A"	
			C_Deptcd		= 1
			C_DeptPop		= 2
			C_DeptNm		= 3
			C_AcctCd		= 4
			C_AcctPop		= 5
			C_AcctNm		= 6
			C_AsstNo		= 7
			C_AsstNm		= 8
			C_AcqAmt		= 9
			C_AcqLocAmt		= 10
			C_AcqQty		= 11
			C_ResAmt		= 12
			C_RefNo			= 13
			C_Desc			= 14

		Case "B"
			C_Seq				= 1
			C_RcptType			= 2									            'Spread Sheet 의 Columns 인덱스 
			C_RcptTypePopup		= 3
			C_RcptTypeNm		= 4								            'Spread Sheet 의 Columns 인덱스 
			C_Amt				= 5
			C_LocAmt			= 6
			C_BankAcct			= 7
			C_BankAcctPopup		= 8
			C_NoteNo			= 9
			C_NoteNoPopup		= 10
	End Select
End Sub




'========================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE
    lgIntGrpCount = 0

    lgStrPrevKey_i = ""                           'initializes Previous Key
    lgStrPrevKey_m = ""
    lgLngCurRows = 0                            'initializes Deleted Rows Count

	lgBlnFlgChgValue = False                    'Indicates that no value changed	
End Sub


'========================================================================================================= 
Sub SetDefaultVal()
	
<%
Dim svrDate
svrDate = GetSvrDate
%>

	if frm1.cboAcqFg.length > 0 then
		frm1.cboAcqFg.selectedIndex = 0
	end if

	frm1.txtAcqDt.text    = UniConvDateAToB("<%=svrDate%>", parent.gServerDateFormat,gDateFormat)	
	frm1.txtGLDt.text     = UniConvDateAToB("<%=svrDate%>", parent.gServerDateFormat,gDateFormat)	
	frm1.txtApDueDt.text  = UniConvDateAToB("<%=svrDate%>", parent.gServerDateFormat,gDateFormat)

	'frm1.txtIssuedDt.text  = UniConvDateAToB("<%=svrDate%>", parent.gServerDateFormat,gDateFormat)		

	frm1.txtDocCur.value	= parent.gCurrency
	frm1.txtXchRate.value	= 1
	
	lgBlnFlgChgValue = False
	
End Sub

'========================================================================================
Sub LoadInfTB19029()
		<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
    	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
		<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub  InitSpreadSheet(ByVal pvSpdNo)
    Call initSpreadPosVariables(pvSpdNo)
   
    Select Case UCase(Trim(pvSpdNo))
		Case "A"
			With frm1.vspdData '자산master	'A
			    ggoSpread.Source = frm1.vspdData
				ggoSpread.Spreadinit "V20060301",,parent.gAllowDragDropSpread 
				
				.ReDraw = False  
			    .MaxCols = C_Desc +1	' 마지막 상수명 사용 
				.Col = .MaxCols			'☜: 공통콘트롤 사용 Hidden Column
				.ColHidden = True
			    .MaxRows = 0

				Call GetSpreadColumnPos("A")

				'Col, Header, ColWidth, HAlign, FloatMax, FloatMin, FloatSeparator, FloatSepChar, FloatDecimalPlaces, FloatDeciamlChar

				ggoSpread.SSSetEdit		C_DeptCd,  "부서코드", 10,0,-1, 10, 2
				ggoSpread.SSSetButton	C_DeptPop
				ggoSpread.SSSetEdit		C_DeptNm,  "부서명",   20

				ggoSpread.SSSetEdit		C_AcctCd,  "계정코드", 10,0,-1, 18, 2
				ggoSpread.SSSetButton	C_AcctPop
				ggoSpread.SSSetEdit		C_AcctNm,  "계정명",   30
				ggoSpread.SSSetEdit		C_AsstNo, "자산번호", 18,0,-1, 18, 2
			    ggoSpread.SSSetEdit		C_AsstNm, "자산명",   30,0,-1, 40, 2
			    
				ggoSpread.SSSetFloat    C_AcqAmt,   "취득금액",      15, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat    C_AcqLocAmt,"취득금액(자국)",15, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				If gIsShowLocal = "N" Then
					'.vspdData.Col		= C_AcqLocAmt
					'.vspdData.ColHidden = True
					Call ggoSpread.SSSetColHidden(C_AcqLocAmt,C_AcqLocAmt,True)
				End If
				Call AppendNumberPlace("6","11","0")
'			    ggoSpread.SSSetFloat    C_AcqQty,   "취득수량",      15,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			    ggoSpread.SSSetFloat    C_AcqQty,   "취득수량",      15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat    C_ResAmt,"잔존가액(자국)",15, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetEdit		C_RefNo, "참조번호", 30,0,-1, 30, 2
				ggoSpread.SSSetEdit		C_Desc,  "적요",     30,0,-1, 128, 2
				
				
				Call ggoSpread.MakePairsColumn(C_DeptCd,C_DeptPop,"1")
				Call ggoSpread.MakePairsColumn(C_AcctCd,C_AcctPop,"1")
				.ReDraw = True				
			End With

		Case "B"
		
			With frm1.vspdData2
				ggoSpread.Source	 = frm1.vspdData2	      
				ggoSpread.Spreadinit "V20021103",,parent.gAllowDragDropSpread  
				.ReDraw = False 
				
				.MaxCols   = C_NoteNoPopup + 1 												'☜: 최대 Columns의 항상 1개 증가시킴 
			'	.Col		= C_RcptType
			'   .ColHidden = True
				.Col		 = .MaxCols													'공통콘트롤 사용 Hidden Column
				.ColHidden = True  

			    Call GetSpreadColumnPos("B")
			    
			    ggoSpread.SSSetEdit	  C_Seq,       "순번",        5, 2, -1, 5
				ggoSpread.SSSetEdit  C_RcptType,  "지급유형"       ,10,,,5,2
				ggoSpread.SSSetButton C_RcptTypePopup
				ggoSpread.SSSetEdit  C_RcptTypeNm,"지급유형명"     ,16

				ggoSpread.SSSetFloat  C_Amt,       "금액",       19, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat  C_LocAmt,    "금액(자국)", 19, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				
				If gIsShowLocal = "N" Then
					'.Col		= C_LocAmt
					'.ColHidden = True
					Call ggoSpread.SSSetColHidden(C_LocAmt,C_LocAmt,True)
				End If
					
				ggoSpread.SSSetEdit	  C_BankAcct,  "예적금코드",   25, 0, -1, 30,2
				ggoSpread.SSSetButton C_BankAcctPopup
				ggoSpread.SSSetEdit   C_NoteNo,    "어음번호",     25, 0, -1, 30,2
				ggoSpread.SSSetButton C_NoteNoPopup
				
				Call ggoSpread.MakePairsColumn(C_RcptType,C_RcptTypePopup,"1")
				Call ggoSpread.MakePairsColumn(C_BankAcct,C_BankAcctPopup,"1")
				Call ggoSpread.MakePairsColumn(C_NoteNo,C_NoteNoPopup,"1")
				'Call ggoSpread.SSSetColHidden(C_RcptType,C_RcptType,True)
				'Call InitComboBox_rcpt()
				.ReDraw = True	
			End With
	End Select

	Call SetSpreadLock(pvSpdNo)

End Sub


'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)

	Select Case UCase(Trim(pvSpdNo))
		Case "A"
			With frm1.vspdData
				ggoSpread.Source = frm1.vspdData
				.ReDraw = True
				ggoSpread.SpreadLock C_DeptNm, -1, C_DeptNm
				ggoSpread.SpreadLock C_AcctNm, -1, C_AcctNm

				ggoSpread.SSSetProtected C_Desc +1, -1,C_Desc +1
				
				If lgIntFlgMode = parent.OPMD_UMODE Then
					ggoSpread.SpreadLock C_AsstNo,   -1, C_DeptNm
				End If

				.ReDraw = True
			End With
		Case "B"
			With frm1.vspdData2

				ggoSpread.Source = frm1.vspdData2
				.ReDraw = False
				ggoSpread.SpreadLock C_Seq,		   -1, C_Seq
				ggoSpread.SSSetProtected C_NoteNoPopup +1, -1,-1

				.ReDraw = True
			End With
		End Select

End Sub



'========================================================================================
Sub SetSpreadColor_Item(ByVal pvStarRow, ByVal pvEndRow)

    With frm1.vspdData2 

		ggoSpread.Source = frm1.vspdData2

		.ReDraw = False

		ggoSpread.SSSetProtected C_Seq, pvStarRow, pvEndRow
		ggoSpread.SSSetRequired	 C_Amt, pvStarRow, pvEndRow
		ggoSpread.SSSetRequired	 C_RcptType, pvStarRow, pvEndRow
		ggoSpread.SSSetProtected C_RcptTypeNm, pvStarRow, pvEndRow

		.ReDraw = True

    End With
End Sub

Sub SetSpreadColor_Master(ByVal pvStarRow, ByVal pvEndRow, ByVal lock_fg)
    
    With frm1.vspdData
    ggoSpread.Source = frm1.vspdData

    .Redraw = False

    ggoSpread.SSSetRequired  C_Deptcd, pvStarRow, pvEndRow
    ggoSpread.SSSetProtected C_DeptNm, pvStarRow, pvEndRow

	ggoSpread.SSSetRequired  C_AcctCd, pvStarRow, pvEndRow
    ggoSpread.SSSetProtected C_AcctNm, pvStarRow, pvEndRow 

    ggoSpread.SSSetRequired  C_AsstNm, pvStarRow, pvEndRow

    ggoSpread.SSSetRequired  C_AcqAmt, pvStarRow, pvEndRow

    ggoSpread.SSSetRequired  C_AcqQty, pvStarRow, pvEndRow

	ggoSpread.SSSetRequired  C_ResAmt, pvStarRow, pvEndRow
	if lock_fg = "query" then
		ggoSpread.SSSetProtected C_AsstNo, pvStarRow, pvEndRow
	end if

    .Redraw = True     

    End With

End Sub


Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_Deptcd		= iCurColumnPos(1)
			C_DeptPop		= iCurColumnPos(2)
			C_DeptNm		= iCurColumnPos(3)
			C_AcctCd		= iCurColumnPos(4)
			C_AcctPop		= iCurColumnPos(5)
			C_AcctNm		= iCurColumnPos(6)
			C_AsstNo		= iCurColumnPos(7)
			C_AsstNm		= iCurColumnPos(8)
			C_AcqAmt		= iCurColumnPos(9)
			C_AcqLocAmt	= iCurColumnPos(10)
			C_AcqQty		= iCurColumnPos(11)
			C_ResAmt		= iCurColumnPos(12)
			C_RefNo		= iCurColumnPos(13)
			C_Desc		= iCurColumnPos(14)
		Case "B"
			ggoSpread.Source = frm1.vspdData2
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_Seq				= iCurColumnPos(1)
			C_RcptType			= iCurColumnPos(2)
			C_RcptTypePopup		= iCurColumnPos(3)
			C_RcptTypeNm		= iCurColumnPos(4)
			C_Amt				= iCurColumnPos(5)
			C_LocAmt			= iCurColumnPos(6)
			C_BankAcct			= iCurColumnPos(7)
			C_BankAcctPopup		= iCurColumnPos(8)
			C_NoteNo			= iCurColumnPos(9)
			C_NoteNoPopup		= iCurColumnPos(10)
	End Select
End Sub

'========================================================================================================= 
Sub InitComboBox_acqfg()

	'------ Developer Coding part (Start ) --------------------------------------------------------------
	Dim IntRetCD1
	Dim intMaxRow, intLoopCnt
	Dim ArrTmpF0, ArrTmpF1

	On error resume next

	IntRetCD1= CommonQueryRs("MINOR_CD,MINOR_NM","B_MINOR","(MAJOR_CD = " & FilterVar("A2005", "''", "S") & " and MINOR_CD<>" & FilterVar("03", "''", "S") & ")",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	ArrTmpF0 = split(lgF0,parent.gColSep)
	ArrTmpF1 = split(lgF1,parent.gColSep)

	intMaxRow = ubound(ArrTmpF0)

	If intRetCD1 <> False Then
		for intLoopCnt = 0 to intMaxRow - 1
			Call SetCombo(frm1.cboAcqFg, ArrTmpF0(intLoopCnt), ArrTmpF1(intLoopCnt))
		next
	End If		
End Sub

 '==========================================  2.2.6 InitComboBox_rcpt()  =======================================
'	Name : InitComboBox_rcpt()
'	Description : Combo Display
'========================================================================================================= 

'Sub InitComboBox_rcpt()
'    Dim IntRetCD1
'    On Error Resume Next
'    IntRetCD1 = CommonQueryRs("A.MINOR_CD,A.MINOR_NM", "B_MINOR A, B_CONFIGURATION B", _
'                        "(A.MINOR_CD = B.MINOR_CD AND A.MAJOR_CD = B.MAJOR_CD) AND (A.MAJOR_CD = 'A1006') AND A.MINOR_CD NOT IN ( 'NR', 'PP', 'CR', 'PR') AND B.SEQ_NO = 4 ", _
'                             lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) 'A1006
'    If IntRetCD1 <> False Then
'        ggoSpread.Source = frm1.vspddata2
'        ggoSpread.SetCombo Replace(lgF0, Parent.gColSep, vbTab), C_RcptType
'        ggoSpread.SetCombo Replace(lgF1, Parent.gColSep, vbTab), C_RcptTypeNm
'    End If
'End Sub
'
'

 '==========================================  2.3.1 Tab Click 처리  =================================================
'	기능: Tab Click
'	설명: Tab Click시 필요한 기능을 수행한다.
'=================================================================================================================== 
 '----------------  ClickTab1(): Header Tab처리 부분 (Header Tab이 있는 경우만 사용)  ---------------------------- 
Function ClickTab1()
	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)	 '~~~ 첫번째 Tab 
	gSelframeFlg = TAB1
End Function

Function ClickTab2()
	Dim IntRetCD
	If gSelframeFlg = TAB2 Then Exit Function
	Call changeTabs(TAB2)	 '~~~ 첫번째 Tab 
	gSelframeFlg = TAB2
End Function

'=======================================================================================================
'   Event Name : txtDueDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtAcqDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtAcqDt.Action = 7
    End If
End Sub

Sub txtApDueDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtApDueDt.Action = 7
    End If
End Sub

Sub txtIssuedDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssuedDt.Action = 7
    End If
End Sub

Sub txtGlDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtGlDt.Action = 7
    End If
End Sub

'======================================================================================================
'   Function Name : OpenAcqNoInfo()
'   Function Desc : 
'=======================================================================================================
Function OpenAcqNoInfo()
	Dim arrRet
	Dim arrParam(3)	
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True	

	iCalledAspName = AskPRAspName("a7102ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a7102ra1" ,"x")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetAcqNoInfo(arrRet)
	End If

End Function

'======================================================================================================
'   Function Name : SetAcqNoInfo(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function SetAcqNoInfo(Byval arrRet)

	With frm1
		.txtAcqNo.value  = arrRet(0)
		.txtAcqNo.focus
	End With

End Function

'===========================================================================
' Function Name : OpenAcctCd
' Function Desc : OpenAcctCd Reference Popup
'===========================================================================
Function OpenAcctCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim field_fg

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "자산계정코드 팝업"			' 팝업 명칭 
	arrParam(1) = "a_Asset_acct  a, a_acct  b"					' TABLE 명칭 
	arrParam(2) = Trim(frm1.vspdData.text)	        ' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = "a.acct_cd = b.acct_cd"							' Where Condition
	arrParam(5) = "계정코드"		    	' 조건필드의 라벨 명칭 

    arrField(0) = "a.ACCT_CD"						' Field명(0)
	arrField(1) = "b.ACCT_NM"						' Field명(1)

    arrHeader(0) = "계정코드"				' Header명(0)
	arrHeader(1) = "계정코드명"				' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")	

	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		field_fg = "AcctCd"
		Call SetReturnVal(arrRet,field_fg)
	End If
End Function

'===========================================================================
' Function Name : OpenDept
' Function Desc : OpenDeptCode Reference Popup
'===========================================================================

Function OpenDept()

	Dim arrRet
	Dim arrParam(3)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtDeptCd.value) 'strCode		            '  Code Condition
   	arrParam(1) = frm1.txtGLDt.Text

	arrRet = window.showModalDialog("../../comasp/DeptPopupDtA2.asp", Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetReturnVal(arrRet,"DeptCd")
	End If	
End Function


'===========================================================================
' Function Name : OpenDeptCd (called from multi grid)
' Function Desc : OpenDeptCode Reference Popup
'===========================================================================
Function OpenDeptCd(strcode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
    dim field_fg

	If RTrim(LTrim(frm1.txtDeptCd.value)) <> "" 	Then
		'jsk scr 238 20030826 전체 부서 조회도로록 수정 
		arrParam(0) = "부서 팝업"	
		arrParam(1) = "B_ACCT_DEPT A, B_COST_CENTER B"
		arrParam(2) = Trim(strCode)
		arrParam(3) = "" 
		arrParam(4) = "A.ORG_CHANGE_ID = " & FilterVar(frm1.hOrgChangeId.value, "''", "S")
		arrParam(4) = arrParam(4) & " AND A.COST_CD = B.COST_CD AND B.BIZ_AREA_CD IN ( SELECT B.BIZ_AREA_CD"
		arrParam(4) = arrParam(4) & " FROM B_ACCT_DEPT A, B_COST_CENTER B "
		arrParam(4) = arrParam(4) & " WHERE A.COST_CD = B.COST_CD AND A.ORG_CHANGE_ID = " & FilterVar(frm1.hOrgChangeId.value, "''", "S") & ")"
		arrParam(5) = "부서코드"
		arrField(0) = "A.DEPT_CD"
		arrField(1) = "A.DEPT_Nm"
		arrField(2) = "B.BIZ_AREA_CD"

		arrHeader(0) = "부서코드"
		arrHeader(1) = "부서코드명"
		arrHeader(2) = "사업장코드"
	Else
		arrParam(0) = "부서 팝업"	
		arrParam(1) = "B_ACCT_DEPT A"
		arrParam(2) = Trim(strCode)
		arrParam(3) = "" 
		arrParam(4) = "A.ORG_CHANGE_ID = (select distinct org_change_id"
		arrParam(4) = arrParam(4) & " from b_acct_dept where org_change_dt = (select max(org_change_dt)"
		arrParam(4) = arrParam(4) & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtGLDt.Text, gDateFormat,""), "''", "S") & "))"

		arrParam(5) = "부서코드"
		arrField(0) = "A.DEPT_CD"
		arrField(1) = "A.DEPT_Nm"
		arrHeader(0) = "부서코드"
		arrHeader(1) = "부서코드명"
	End IF

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		field_fg = "DeptCd_grid"
		Call SetReturnVal(arrRet,field_fg)
	End If
End Function

'===========================================================================
' Function Name : OpenCurrency()
' Function Desc : OpenCurrency Reference Popup
'===========================================================================
Function OpenCurrency()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
    dim field_fg

	If IsOpenPop = True Or UCase(frm1.txtDocCur.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "거래통화 팝업"	
	arrParam(1) = "B_CURRENCY"
	arrParam(2) = Trim(frm1.txtDocCur.Value)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "거래통화"

    arrField(0) = "CURRENCY"
    arrField(1) = "CURRENCY_DESC"

    arrHeader(0) = "거래통화"		
    arrHeader(1) = "거래통화명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		field_fg = "Currency"
		Call SetReturnVal(arrRet,field_fg)
	End If
End Function

'===========================================================================
' Function Name : OpenBpCd()
' Function Desc : OpenBpCd Reference Popup
'===========================================================================
Function OpenBpCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
    dim field_fg

	If IsOpenPop = True Or UCase(frm1.txtBpCd.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "거래처 팝업"	
	arrParam(1) = "B_BIZ_PARTNER"
	arrParam(2) = Trim(frm1.txtBpCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "거래처 코드"

    arrField(0) = "BP_CD"
    arrField(1) = "BP_NM"
    arrField(2) = "BP_RGST_NO"

    arrHeader(0) = "거래처 코드"
    arrHeader(1) = "거래처 명"
    arrHeader(2) = "사업자등록번호"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=650px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		field_fg = "BpCd"
		Call SetReturnVal(arrRet,field_fg)
	End If
End Function

'------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function
		
	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = ""							' 채권과 연계(거래처 유무)
	arrParam(2) = ""								'FrDt
	arrParam(3) = ""								'ToDt
	arrParam(4) = "T"							'B :매출 S: 매입 T: 전체 
	arrParam(5) = ""									'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 	
	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtBpCd.focus
		Exit Function
	Else
		Call SetBpCd(arrRet)
		lgBlnFlgChgValue = True
	End If
		
End Function
'========================================================================================
Function SetBpCd(byval arrRet)
	frm1.txtBpCd.focus
	frm1.txtBpCd.Value    = arrRet(0)		
	frm1.txtBpNm.Value    = arrRet(1)		
	lgBlnFlgChgValue = True
End Function

'===========================================================================
' Function Name : OpenApAcct()
' Function Desc : OpenApAcct Reference Popup
'===========================================================================
Function OpenApAcct()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
    dim field_fg

	If IsOpenPop = True Or UCase(frm1.txtBpCd.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "미지급금계정 팝업"	
	arrParam(1) = "a_jnl_acct_assn a, a_acct b"
	arrParam(2) = Trim(frm1.txtApAcctCd.Value)
	arrParam(3) = ""
	arrParam(4) = "A.trans_type = " & FilterVar("AS001", "''", "S") & " and A.Acct_cd = B.Acct_cd and Jnl_cd = " & FilterVar("AP", "''", "S") & ""
	arrParam(5) = "미지급금계정 코드"

    arrField(0) = "a.acct_cd"
    arrField(1) = "b.acct_nm"

    arrHeader(0) = "미지급금계정 코드"
    arrHeader(1) = "미지급금계정명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=470px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		field_fg = "ApCd"
		Call SetReturnVal(arrRet,field_fg)
	End If
End Function
'========================================================================================
Function OpenCardNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
    dim field_fg

	If IsOpenPop = True Or UCase(frm1.txtCardNo.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "신용카드번호"	
	arrParam(1) = "b_credit_card"
	arrParam(2) = Trim(frm1.txtCardNo.Value)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "신용카드번호"

    arrField(0) = "Credit_no"
    arrField(1) = "Credit_nm"

    arrHeader(0) = "신용카드번호"		
    arrHeader(1) = "신용카드명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		field_fg = "CreditNo"
		Call SetReturnVal(arrRet,field_fg)
	End If
End Function


'=======================================================================================================
'	Name : OpenNoteNo()
'	Description : Note No PopUp
'=======================================================================================================
Function OpenNoteNo(byVal strCode, byVal strCard)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim field_fg
	If IsOpenPop = True  Then Exit Function

	IsOpenPop = True	

IF UCase(strCard) = "CP"	Then

	arrParam(0) = "지불구매카드 팝업"				        ' 팝업 명칭 
	arrParam(1) = "f_note a,b_biz_partner b, b_bank c, b_card_co d"		' TABLE 명칭 
	arrParam(2) = ""								' Code Condition
	arrParam(3) = ""								' Name Cindition			
	arrParam(4) = "a.note_sts = " & FilterVar("BG", "''", "S") & " AND a.note_fg = " & FilterVar("CP", "''", "S") & " and a.bp_cd = b.bp_cd  "			
	arrParam(4) = arrParam(4) & " and a.bank_cd *= c.bank_cd and a.card_co_cd *= d.card_co_cd "
	arrParam(5) = "구매카드번호"						' 조건필드의 라벨 명칭 

	arrField(0) = "a.Note_no"					' Field명(0)
	arrField(1) = "F2" & parent.gColSep & "a.Note_amt"		' Field명(1)
	arrField(2) = "DD" & parent.gColSep & "a.Issue_dt"		' Field명(2)
	arrField(3) = "b.bp_nm"					' Field명(3)
	arrField(4) = "d.card_co_nm"    	    			' Field명(4)

	arrHeader(0) = "구매카드번호"				' Header명(0)
	arrHeader(1) = "금액"				' Header명(1)
	arrHeader(2) = "발행일"				' Header명(2)
	arrHeader(3) = "거래처"				' Header명(3)
	arrHeader(4) = "카드사"				' Header명(4)



ElseIf UCase(strCard) = "NP"	Then
	arrParam(0) = "지급어음번호 팝업"	
	arrParam(1) = "F_NOTE A,B_BANK B,B_BIZ_PARTNER C"
	arrParam(2) = strCode
	arrParam(3) = ""
	
	arrParam(4) = "A.NOTE_STS = " & FilterVar("BG", "''", "S") & " AND A.NOTE_FG = " & FilterVar("D3", "''", "S") & " AND A.BP_CD = C.BP_CD AND A.BANK_CD = B.BANK_CD"
	arrParam(5) = "지급어음번호"
	
    arrField(0) = "A.NOTE_NO"
    arrField(1) = "F2" & parent.gColSep & "A.NOTE_AMT"
    arrField(2) = "C.BP_NM"	    
    arrField(3) = "DD" & parent.gColSep & "A.ISSUE_DT"
    arrField(4) = "DD" & parent.gColSep & "A.DUE_DT"
    arrField(5) = "B.BANK_NM"
    
    arrHeader(0) = "지급어음번호"
    arrHeader(1) = "어음금액"
	arrHeader(2) = "거래처"
	arrHeader(3) = "발행일"
	arrHeader(4) = "만기일"
	arrHeader(5) = "은행"

Else
	arrParam(0) = "배서어음번호 팝업"	
	arrParam(1) = "F_NOTE A,B_BANK B,B_BIZ_PARTNER C"
	arrParam(2) = strCode
	arrParam(3) = ""

	arrParam(4) = "A.NOTE_STS = " & FilterVar("ED", "''", "S") & " AND A.NOTE_FG = " & FilterVar("D1", "''", "S") & " AND A.BP_CD = C.BP_CD AND A.BANK_CD = B.BANK_CD"
	arrParam(5) = "배서어음번호"

    arrField(0) = "A.NOTE_NO"
    arrField(1) = "F2" & parent.gColSep & "A.NOTE_AMT"
    arrField(2) = "C.BP_NM"
    arrField(3) = "DD" & parent.gColSep & "A.ISSUE_DT"
    arrField(4) = "DD" & parent.gColSep & "A.DUE_DT"
    arrField(5) = "B.BANK_NM"

    arrHeader(0) = "배서어음번호"
    arrHeader(1) = "어음금액"
	arrHeader(2) = "거래처"
	arrHeader(3) = "발행일"
	arrHeader(4) = "만기일"
	arrHeader(5) = "은행"

End If

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=700px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		field_fg = "NoteNo"
		Call SetReturnVal(arrRet,field_fg)
	End If

End Function

'=======================================================================================================
'	Name : OpenBankAcct()
'	Description : Bank Account No PopUp
'=======================================================================================================
Function OpenBankAcct(byVal strCode , byVal strCard)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim field_fg

	If IsOpenPop = True  Then Exit Function

	IF UCase(strCard) = "DF"	Then

		arrParam(0) = "예적금코드 팝업"
		arrParam(1) = "B_BANK A, B_BANK_ACCT B, F_DPST C"
		arrParam(2) = strCode
		arrParam(3) = ""
		arrParam(4) = "A.BANK_CD = B.BANK_CD "
		arrParam(4) = arrParam(4) & "AND A.BANK_CD = C.BANK_CD "
		arrParam(4) = arrParam(4) & "AND B.BANK_ACCT_NO = C.BANK_ACCT_NO "
		arrParam(5) = "은행코드"

		arrField(0) = "A.BANK_NM"
		arrField(1) = "B.BANK_ACCT_NO"

		arrHeader(0) = "은행명"
		arrHeader(1) = "예적금코드"
	Else
		arrParam(0) = "외화예적금코드 팝업"
		arrParam(1) = "B_BANK A, B_BANK_ACCT B, F_DPST C"
		arrParam(2) = strCode
		arrParam(3) = ""
		arrParam(4) = "A.BANK_CD = B.BANK_CD "
		arrParam(4) = arrParam(4) & "AND A.BANK_CD = C.BANK_CD "
		arrParam(4) = arrParam(4) & "AND B.BANK_ACCT_NO = C.BANK_ACCT_NO "
		arrParam(5) = "은행코드"

		arrField(0) = "A.BANK_NM"
		arrField(1) = "B.BANK_ACCT_NO"

		arrHeader(0) = "은행명"
		arrHeader(1) = "예적금코드"

	End If

	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		field_fg = "BankAcct"
		Call SetReturnVal(arrRet,field_fg)
	End If

End Function


'--------------------------------------------------------------------------------------------------------- 
Function SetReturnVal(byval arrRet,byval field_fg)
	With frm1	
		Select case field_fg
		
			case "DeptCd"
				.txtGLDt				= arrRet(3)
				.txtDeptCd.value        = arrRet(0)
				.txtDeptNm.value 		= arrRet(1)
				Call txtDeptCd_OnChange()

			case "BpCd"
				.txtBpCd.Value			= arrRet(0)
				.txtBpNm.Value			= arrRet(1)

			case "ApCd"
				.txtApAcctCd.Value		= arrRet(0)
				.txtApAcctNm.Value		= arrRet(1)

			case "Currency"
				.txtDocCur.Value		= arrRet(0)
				call txtDocCur_onChange()

			case "BankAcct"
				.vspdData2.Col			= C_BankAcct
				.vspdData2.Text			= arrRet(1)

			case "NoteNo"
				.vspdData2.Col			= C_NoteNo
				.vspdData2.Text			= arrRet(0)
				.vspdData2.Col			= C_Amt	
				.vspdData2.Text			= arrRet(1)
				.vspdData2.Col			= C_LocAmt
				.vspdData2.Text			= arrRet(1)

			case "DeptCd_grid"
				.vspdData.Col			= C_DeptCd
				.vspdData.Text			= arrRet(0)
				.vspdData.Col			= C_DeptNm
				.vspdData.Text			= arrRet(1)

			case "AcctCd"
				.vspdData.Col			= C_AcctCd
				.vspdData.Text			= arrRet(0)
				.vspdData.Col			= C_AcctNm
				.vspdData.Text			= arrRet(1)
			
			case "CreditNo"
				.txtCardNo.Value		= arrRet(0)
				.txtCardNm.Value		= arrRet(1)		
				
				

		End select	

		lgBlnFlgChgValue = True
	End With
End Function

Function OpenVatType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Then Exit Function

   	arrHeader(0) = "부가세유형"
	arrHeader(1) = "부가세명"
	arrHeader(2) = "부가세Rate"

	arrField(0) = "B_Minor.MINOR_CD"
	arrField(1) = "B_Minor.MINOR_NM"
	arrField(2) = "F2" & parent.gColSep & "b_configuration.REFERENCE"

	arrParam(0) = "부가세유형"
	arrParam(1) = "B_Minor,b_configuration"
	arrParam(2) = Trim(frm1.txtVatType.value)
	'arrParam(3) = Trim(frm1.txtPayMethNM.Value)

	arrParam(4) = "B_Minor.Major_Cd=" & FilterVar("B9001", "''", "S") & " and B_Minor.minor_cd = b_configuration.minor_cd and " & _
	              " b_configuration.SEQ_NO=1 AND b_configuration.major_cd = B_Minor.Major_Cd"	 
	arrParam(5) = "부가세유형"						' TextBox 명칭 

	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetVatType(arrRet)
	End If
End Function

'=======================================================================================================
Function SetVatType(byval arrRet)
	frm1.txtVatType.Value		= arrRet(0)
	frm1.txtVatTypeNm.Value     = arrRet(1)
	frm1.txtVatRate.Value		= arrRet(2)
	lgBlnFlgChgValue = True
	Call txtVatType_OnChange
End Function


'==========================================================================================
Sub txtVatType_OnChange()

	Dim AmtValue
	'lgBlnFlgChgValue = True
	lgBlnFlgChgValue = True

	If Trim(frm1.txtVatAmt.text) = "" then
		AmtValue = 0
	else
		AmtValue = UNICDbl(frm1.txtVatAmt.text)
	end if

	If Trim(frm1.txtVatType.Value) <> "" or AmtValue > 0 Then
		ggoOper.SetReqAttr frm1.txtVatAmt, "N"    '부가세금액 
		ggoOper.SetReqAttr frm1.txtVatType, "N"    '부가세타입 
'		ggoOper.SetReqAttr frm1.txtReportAreaCd, "N"
	Else
		ggoOper.SetReqAttr frm1.txtVatAmt, "D"    '부가세금액 D
		ggoOper.SetReqAttr frm1.txtVatType, "D"    '부가세타입 D
'		ggoOper.SetReqAttr frm1.txtReportAreaCd, "D"
	End If

End Sub


Sub  txtReportAreaCd_OnChange()
	lgBlnFlgChgValue = True

	If UNIConvNum(frm1.txtVatAmt.Text,0) <> 0 Or Trim(frm1.txtVatType.value) <> ""  Then
		ggoOper.SetReqAttr frm1.txtVatAmt, "N"    '부가세금액 
		ggoOper.SetReqAttr frm1.txtVatType, "N"    '부가세타입 
'		ggoOper.SetReqAttr frm1.txtReportAreaCd, "N"
	Else
		ggoOper.SetReqAttr frm1.txtVatAmt, "D"    '부가세금액 D
		ggoOper.SetReqAttr frm1.txtVatType, "D"    '부가세타입 D
'		ggoOper.SetReqAttr frm1.txtReportAreaCd, "D"
	End If
End Sub
'=======================================================================================================
'Description : 결의전표 생성내역 팝업 
'=======================================================================================================
Function OpenPopupTempGL()

	Dim arrRet
	Dim arrParam(1)	
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	iCalledAspName = AskPRAspName("a5130ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1" ,"x")
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = Trim(frm1.txtTempGlNo.value)	'결의전표번호 
	arrParam(1) = ""							'Reference번호 

	IsOpenPop = True

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
End Function
'=======================================================================================================
'Description : 회계전표 생성내역 팝업 
'=======================================================================================================
Function OpenPopupGL()

	Dim arrRet
	Dim arrParam(1)	
	Dim iCalledAspName

	iCalledAspName = AskPRAspName("a5120ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1" ,"x")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function

	arrParam(0) = Trim(frm1.txtGlNo.value)	'회계전표번호 
	arrParam(1) = ""						'Reference번호 

	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/a5120ra1.asp", Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

End Function


'===========================================================================
' Function Name : OpenReportAreaCd
' Function Desc : OpenReportAreaCd Reference Popup
'===========================================================================
Function OpenReportAreaCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	IsOpenPop = True

	arrParam(0) = "신고사업장 팝업"
	arrParam(1) = "B_TAX_BIZ_AREA"
	arrParam(2) = Trim(frm1.txtReportAreaCd.value)
	arrParam(3) = "" 
	arrParam(4) = ""
	arrParam(5) = "신고사업장"
	
    arrField(0) = "TAX_BIZ_AREA_CD"	
    arrField(1) = "TAX_BIZ_AREA_NM"
    
    arrHeader(0) = "신고사업장코드"
    arrHeader(1) = "신고사업장명"
        
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetReportArea(arrRet)
	End If	
End Function

'=======================================================================================================
'	Name : SetReportArea()
'	Description : Bp Cd Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetReportArea(byval arrRet)
	frm1.txtReportAreaCd.Value		= arrRet(0)
	frm1.txtReportAreaNm.Value		= arrRet(1)
	lgBlnFlgChgValue = True
End Function


Function  OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim arrParamAdo(3)

	If IsOpenPop = True Then Exit Function	
	
	Select Case iWhere
		Case 6    
			arrParam(0) = "입금유형"								' 팝업 명칭 
		 
			arrParam(1) = "B_MINOR A, B_CONFIGURATION B"
			arrParam(2) = Trim(frm1.vspdData2.text)
			arrParam(3) = ""											' Name Condition
			arrParam(4) = "(A.MINOR_CD = B.MINOR_CD AND A.MAJOR_CD = B.MAJOR_CD) AND (A.MAJOR_CD = " & FilterVar("A1006", "''", "S") & ") " _
			   & " AND A.MINOR_CD NOT IN ( " & FilterVar("NR", "''", "S") & ", " & FilterVar("PP", "''", "S") & ", " & FilterVar("CR", "''", "S") & ", " & FilterVar("PR", "''", "S") & ") AND B.SEQ_NO = 4 " ' Where Condition        
			arrParam(5) = "입금유형"								' TextBox 명칭 
	 
			arrField(0) = "A.MINOR_CD"							' Field명(0)
			arrField(1) = "A.MINOR_NM"							' Field명(1)
			  
			arrHeader(0) = "입금유형"								' Header명(0)
			arrHeader(1) = "입금유형명"								' Header명(1) 
	End Select
	
	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")			

	IsOpenPop = False

	Call GridSetFocus(iWhere)
	If arrRet(0) = "" Then	    
		Exit Function
	Else
		Call SetPopup(arrRet, iWhere)
	End If

End Function
'=======================================================================================================
Function GridsetFocus(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 6
				Call SetActiveCell(.vspdData2,C_Rcpttype,.vspdData2.ActiveRow ,"M","X","X")
		END Select
	End With
End Function
'======================================================================================================
'   Function Name : SetPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function SetPopup(Byval arrRet,Byval iWhere)
	With frm1
	Select Case iWhere
		Case 6
			.vspdData2.Col = C_RcptType
			.vspdData2.Text = arrRet(0)
			.vspdData2.Col = C_RcptTypeNm
			.vspdData2.Text = arrRet(1)
			Call vspdData2_Change(C_RcptType, frm1.vspdData2.Row)				 ' 변경이 읽어났다고 알려줌 		
	END Select
	End With
	lgBlnFlgChgValue = True
End Function


'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format

	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)

    frm1.txtAcqAmt.AllowNull =false
    frm1.txtAcqLocAmt.AllowNull =false
    frm1.txtApAmt.AllowNull =false
    frm1.txtApLocAmt.AllowNull =false
    frm1.txtVatAmt.AllowNull =false
    frm1.txtVatLocAmt.AllowNull =false

    Call ggoOper.LockField(Document, "N") 
    Call InitSpreadSheet("A")
    Call InitSpreadSheet("B")

    Call InitVariables
    Call SetDefaultVal

	Call SetToolBar("1110110100101111")
	Call InitComboBox_acqfg

	gSelframeFlg = TAB1
	frm1.txtAcqNo.focus

End Sub

'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

Function SetProctedField(Byval pAcqFg)	
	Select case pAcqFg
		Case "01"   ''신규취득 
			If gIsShowLocal <> "N" Then			         
				ggoOper.SetReqAttr frm1.fpDoubleSingle5, "D"       '''부가세 금액 
				ggoOper.SetReqAttr frm1.fpDoubleSingle6, "D"       '''부가세 금액 자국 
				ggoOper.SetReqAttr frm1.fpDoubleSingle11, "D"	   '''부가세		
				ggoOper.SetReqAttr frm1.txtVatType,		"D"				'''부가세타입 

				ggoOper.SetReqAttr frm1.txtReportAreaCd,		"D"			'''부가세신고사업장 
				ggoOper.SetReqAttr frm1.fpDateTime4,	"D"				'''계산서계산일 

			End If

			ggoOper.SetReqAttr frm1.fpDateTime2,	 "D"       '''fpDateTime2
			ggoOper.SetReqAttr frm1.txtBpCd,		 "N"      '''거래처 필수		

			Call txtVatType_OnChange		'로드시 부가세 필수 구분 


		Case "02"   ''무상기부 
			ggoOper.SetReqAttr frm1.fpDoubleSingle5, "Q"       '''미지급금 금액 
			'ggoOper.SetReqAttr frm1.fpDoubleSingle6, "Q"       '''미지급금 금액 자국 
			ggoOper.SetReqAttr frm1.fpDateTime2,	 "Q"       '''fpDateTime2

			If gIsShowLocal <> "N" Then			         
				ggoOper.SetReqAttr frm1.fpDoubleSingle11, "Q"	   '''부가세		
				ggoOper.SetReqAttr frm1.txtVatType,		"Q"				'''부가세타입 

				ggoOper.SetReqAttr frm1.txtReportAreaCd,	"Q"			'''부가세신고사업장 
				ggoOper.SetReqAttr frm1.fpDateTime4,	"Q"				'''계산서계산일 

			End If

			ggoOper.SetReqAttr frm1.txtBpCd,		 "D"      '''거래처 Optional		

		Case "03"   ''기초자산 
			ggoOper.SetReqAttr frm1.fpDoubleSingle5, "Q"       '''미지급금 금액 
			'ggoOper.SetReqAttr frm1.fpDoubleSingle6, "Q"       '''미지급금 금액 자국 
			ggoOper.SetReqAttr frm1.fpDateTime2,	 "Q"       '''fpDateTime2

			If gIsShowLocal <> "N" Then			         
				ggoOper.SetReqAttr frm1.fpDoubleSingle11, "Q"	   '''부가세		
				ggoOper.SetReqAttr frm1.txtVatType,		"Q"			'''부가세타입 
			End If

			ggoOper.SetReqAttr frm1.txtBpCd,		 "Q"      '''거래처 Optional
	End Select
End Function

Function cboAcqFg_onChange()
	Dim varAcqFg	
	varAcqFg = frm1.cboAcqFg.value 	

	if frm1.cboAcqFg.value = "03" then
		'frm1.vspdData2.MaxRows = 0
		ggoSpread.Source = frm1.vspdData2
		ggospread.ClearSpreadData		'Buffer Clear
	end if
	Call SetProctedField(varAcqFg)	

End Function

'==========================================================================================
'   Event Name : txtDocCur_OnChange
'   Event Desc : 
'==========================================================================================
Sub txtDocCur_OnChange()

    lgBlnFlgChgValue = True									'수정/변동된 내역이 있음을 Setting
    If UCase(frm1.txtDocCur.value) <> parent.gCurrency Then		' 거래통화하고 Company 통화가 다를때 환율을 0으로 셋팅 
		frm1.txtXchRate.text	= 0                         ' 디폴트값인 1이 들어가 있으면 환율이 입력된 것으로 판단하여 
								                                        ' 환율정보를 읽지 않고 입력된 값으로 계산. 
	Else 
		frm1.txtXchRate.text	= 1
	End If

	IF CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then	

		Call CurFormatNumericOCX()
		Call CurFormatNumSprSheet()
	End If
End Sub

function txtApAmt_onBlur()

	lgBlnFlgChgValue = True

	if UNICDbl(frm1.txtAcqAmt.text) <> 0 then
''''		frm1.fpDateTime2.Enabled = True
		ggoOper.SetReqAttr frm1.fpDateTime2, "N"	   '''fpDateTime2
	else
'''''	    frm1.fpDateTime2.Enabled = False
		ggoOper.SetReqAttr frm1.fpDateTime2, "Q"	   '''fpDateTime2
	end if
end function

function txtApLocAmt_onBlur()
	lgBlnFlgChgValue = True

	if UNICDbl(frm1.txtAcqAmt.text) <> 0 then
		ggoOper.SetReqAttr frm1.txtApDueDt, "N"	   '''fpDateTime2
	else
		ggoOper.SetReqAttr frm1.txtApDueDt, "Q"	   '''fpDateTime2
	end if	
end function

'function txtVatAmt_change()	'onBlur
'	lgBlnFlgChgValue = True
'end function

'function txtVatLocAmt_change()	'onBlur
'	lgBlnFlgChgValue = True
'end function


'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopUpMenuItemInf("1101111111")
    gMouseClickStatus = "SPC"	'Split 상태코드 
    Set gActiveSpdSheet = frm1.vspdData

	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort
			lgSortKey = 2
		Else
			ggoSpread.SSSort ,lgSortKey
			lgSortKey = 1
		End If
	End If
End Sub


'==========================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)
	Call SetPopUpMenuItemInf("1101111111")
    gMouseClickStatus = "SP2C"	'Split 상태코드 

    Set gActiveSpdSheet = frm1.vspdData2

	If Row = 0 Then
		ggoSpread.Source = frm1.vspdData2
		If lgSortKey = 1 Then
			ggoSpread.SSSort
			lgSortKey = 2
		Else
			ggoSpread.SSSort ,lgSortKey
			lgSortKey = 1
		End If
	End If
End Sub


'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal Col1, ByVal Col2)		' 7) 컬럼 width 변경 이벤트 핸들러 
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(Col1,Col2)

End Sub

Sub vspdData2_ColWidthChange(ByVal Col1, ByVal Col2)		' 7) 컬럼 width 변경 이벤트 핸들러 
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(Col1,Col2)

End Sub


'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				'8) 컬럼 title 변경 
    Dim iColumnName

	If Row<=0 then
		Exit Sub
	End If
	If Frm1.vspdData.MaxRows =0 then
		Exit Sub
	End if
End Sub

Sub vspdData2_DblClick(ByVal Col, ByVal Row)				'8) 컬럼 title 변경 
    Dim iColumnName

	If Row<=0 then
		Exit Sub
	End If
	If Frm1.vspdData.MaxRows =0 then
		Exit Sub
	End if
End Sub
'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================

Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

Sub vspdData2_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc : 
'==========================================================================================

Sub vspdData_Change(ByVal Col, ByVal Row)

    Call CheckMinNumSpread(frm1.vspdData, Col, Row)  

	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
	
End Sub

Sub vspdData2_Change(ByVal Col, ByVal Row)
	Dim intIndex
	Dim varData

	With frm1.vspdData2
	
		.Row = Row

		frm1.vspdData2.ReDraw = False
		Select Case Col
			Case  C_RcptType
				.Col = Col
				intIndex = .Value
				.Col = C_RcptType
				.Value = intIndex
				varData = .text
				If Trim(varData) <> "" Then 
					IF CommonQueryRs( " A.MINOR_CD,A.MINOR_NM " , "B_MINOR A, B_CONFIGURATION B  " , "  (A.MINOR_CD = B.MINOR_CD AND A.MAJOR_CD = B.MAJOR_CD) AND A.MAJOR_CD = " & FilterVar("A1006", "''", "S") & " AND A.MINOR_CD =  " & FilterVar(varData , "''", "S") & " AND SEQ_NO = 4  AND A.MINOR_CD NOT IN ( " & FilterVar("NR", "''", "S") & ", " & FilterVar("PP", "''", "S") & ", " & FilterVar("CR", "''", "S") & ", " & FilterVar("PR", "''", "S") & ")", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
						Select Case UCase(lgF0)
							Case "DP" & Chr(11)			' 예적금 
								.Row  = Row
								.Col  = C_NoteNo
								.Text = ""
							Case "NO" & Chr(11)
								.Row  = Row
								.Col  = C_BankAcct
								.Text = ""
							Case Else
								.Row  = Row
								.Col  = C_NoteNo
								.Text = ""

								.Row  = Row
								.Col  = C_BankAcct
								.Text = ""
						End Select
						.Col  = C_RcptTypeNm
						.Text = Replace(lgF1, Chr(11), "")
					Else
						Call DisplayMsgBox("179051", "X", "X" ,"x")
						.Col  = C_RcptType
						.Text = ""
						.Col  = C_RcptTypeNm
						.Text = ""
						Call SetActiveCell(frm1.vspdData2,C_RcptType,frm1.vspdData2.ActiveRow ,"M","X","X")
					End if
				End if

				'.Col  = C_Amt
				'.Text = ""
				'.Col  = C_LocAmt
				'.Text = ""

				call subVspdSettingChange(Row,varData)
		End Select
	End With

	frm1.vspdData2.ReDraw = True	


	Call CheckMinNumSpread(frm1.vspdData2, Col, Row)

	lgBlnFlgChgValue = True
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.UpdateRow Row
	
End Sub


'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'==========================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    With frm1.vspdData 
  If Row >= NewRow Then
      Exit Sub
  End If
    End With
End Sub

Sub vspdData2_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    With frm1.vspdData2 
  If Row >= NewRow Then
      Exit Sub
  End If
    End With
End Sub

'==========================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'==========================================================================================

Sub vspdData_ScriptDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
    Call GetSpreadColumnPos("A")
End Sub

Sub vspdData2_ScriptDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
    Call GetSpreadColumnPos("B")
End Sub
 '==========================================================================================
   '   Event Name : vspdData_ComboSelChange
   '   Event Desc : Combo 변경 이벤트 
   '========================================================================================== 
Sub vspdData2_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	Dim varData

	With frm1.vspdData2
	
		.Row = Row
    
		Select Case Col
			Case  C_RcptTypeNm
				.Col = Col
				intIndex = .Value
				.Col = C_RcptType
				.Value = intIndex
				varData = .text
		End Select
	End With	

	frm1.vspdData2.ReDraw = False		

	 IF CommonQueryRs( "REFERENCE" , "B_CONFIGURATION  " , "MAJOR_CD = " & FilterVar("A1006", "''", "S") & " AND MINOR_CD =  " & FilterVar(varData , "''", "S") & " AND SEQ_NO = 4 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then         
		   Select Case UCase(lgF0)					
				Case "DP" & Chr(11)			' 예적금 
					frm1.vspdData2.Row  = Row
					frm1.vspdData2.Col  = C_NoteNo
					frm1.vspdData2.Text = ""
				Case "NO" & Chr(11)											
					frm1.vspdData2.Row  = Row
					frm1.vspdData2.Col  = C_BankAcct
					frm1.vspdData2.Text = ""			
				Case Else
					frm1.vspdData2.Row  = Row
					frm1.vspdData2.Col  = C_NoteNo
					frm1.vspdData2.Text = ""			
							
					frm1.vspdData2.Row  = Row
					frm1.vspdData2.Col  = C_BankAcct
					frm1.vspdData2.Text = ""				
			End Select			

	end if
		
'	frm1.vspdData2.Col  = C_Amt
'	frm1.vspdData2.Text = ""	
'	frm1.vspdData2.Col  = C_LocAmt
'	frm1.vspdData2.Text = ""	
				
	call subVspdSettingChange(Row,varData)

	frm1.vspdData2.ReDraw = True	

End Sub

 '==========================================================================================
'   Sub Procedure Name : subVspdSettingChange
'   Sub Procedure Desc : 
'==========================================================================================

Sub subVspdSettingChange(ByVal lRow, Byval varData)	

	ggoSpread.Source = frm1.vspdData2
	
	IF CommonQueryRs( "REFERENCE" , "B_CONFIGURATION  " , "MAJOR_CD = " & FilterVar("A1006", "''", "S") & " AND MINOR_CD =  " & FilterVar(varData , "''", "S") & " AND SEQ_NO = 4 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then         
		   Select Case UCase(lgF0)				
				Case "DP" & Chr(11)			' 예적금 
					ggoSpread.SSSetRequired	C_BankAcct,		 lRow, lRow			
					ggoSpread.SpreadUnLock  C_BankAcct,      lRow, C_BankAcct
					ggoSpread.SpreadUnLock  C_BankAcctPopUp, lRow, C_BankAcctPopUp

					'ggoSpread.SSSetEdit	C_BankAcct, "예적금코드", 25, 0, lRow, 30    
		
					ggoSpread.SSSetRequired	C_BankAcct,      lRow, lRow	
												
					ggoSpread.SpreadLock     C_NoteNo,		 lRow, C_NoteNo,lRow   '어음번호 protect
					ggoSpread.SSSetProtected C_NoteNo,       lRow, lRow						
					ggoSpread.SpreadLock     C_NoteNoPopup,  lRow, C_NoteNoPopup,lRow          

	
				Case "NO" & Chr(11)				
					ggoSpread.SpreadUnLock   C_NoteNo,        lRow, C_NoteNo,       lRow
					ggoSpread.SpreadUnLock   C_NoteNoPopup,   lRow, C_NoteNoPopup,  lRow
					 
					ggoSpread.SpreadLock     C_BankAcct,      lRow, C_BankAcct,     lRow   
					ggoSpread.SpreadLock     C_BankAcctPopup, lRow, C_BankAcctPopup,lRow
		
					ggoSpread.SSSetProtected C_BankAcct,      lRow, lRow								
		
					'ggoSpread.SSSetEdit      C_NoteNo, "어음번호", 25, 0, lRow, 30	
					ggoSpread.SSSetRequired  C_NoteNo,        lRow, lRow
		
				Case Else									
					ggoSpread.SpreadLock     C_BankAcct,      lRow, C_BankAcct,     lRow   			
					ggoSpread.SpreadLock     C_BankAcctPopup, lRow, C_BankAcctPopup,lRow
							
					ggoSpread.SSSetProtected C_BankAcct,      lRow, lRow							
		
					ggoSpread.SpreadLock     C_NoteNo,        lRow, C_NoteNo,     lRow
					ggoSpread.SpreadLock     C_NoteNoPopup,   lRow, C_NoteNoPopup,lRow		
		
					ggoSpread.SSSetProtected C_NoteNo,        lRow, lRow													
			End Select		
		
	end if
	
End Sub	

Sub fncGetVspdLocalAmt(ByVal Col, Byval Row)
	Dim strVal
    Dim varAmt
    Dim strDoc
    Dim varXrate
	
	Err.Clear

	fncGetVspdLocalAmt = false
	
	frm1.vspdData.Col = Col
	varAmt = UNICDbl(frm1.vspdData.Text)
	
	frm1.vspdData.Col = "C_DocCurr"
	varDoc = frm1.vspdData.Text 
	
	frm1.vspdData.Col = "C_Xrate"
	varXrate = frm1.vspdData.Text 
	
	
	strVal = BIZ_PGM_ID2 & "?txtMode=" & "LocAmt"
	strVal = strVal & "&txtAmtFg=" & "VspdAcqAmt"
	strVal = strVal & "&txtLocCurr=" & parent.gCurrency
 	strVal = strVal & "&txtToCurr=" & varDoc
 	strVal = strVal & "&txtAcqAmt=" & varXrate
 	strVal = strVal & "&txtFromAmt=" & varAmt 	 	
	
	strVal = strVal & "&txtAppDt=" & UniConvDateToYYYYMMDD(frm1.txtAcqDt.text, gDateFormat, parent.gServerDateType) '☆: 조회 조건 데이타 
	    
	Call RunMyBizASP(MyBizASP, strVal) 	
 	
End Sub

 '-----------------------------------------------------------------------------------------------------
'	Name : fncGetLocalAmt()
'	Description : Get local amt for each field's amt
'--------------------------------------------------------------------------------------------------------- 
Function fncGetLocalAmt(byval Amt_fg)
	Dim strVal

	Err.Clear

	fncGetLocalAmt = false
	
	if Trim(frm1.fpDoubleSingle1.Value) = 0 then
		frm1.fpDoubleSingle2.Value = o		
	else
		strVal = BIZ_PGM_ID2 & "?txtMode="   & "LocAmt"	        
		strVal = strVal & "&txtAmtFg=" & Amt_fg
 		strVal = strVal & "&txtLocCurr=" & parent.gCurrency
 		strVal = strVal & "&txtToCurr=" & Trim(frm1.txtCurr.value)
 		
 		Select case Amt_fg
 	 	case "AcqAmt"
 	 		strVal = strVal & "&txtFromAmt=" & Trim(frm1.fpDoubleSingle1.Value) 
 	 	case "PaymAmt1" 
 	 		strVal = strVal & "&txtFromAmt=" & Trim(frm1.fpDoubleSingle3.Value) 
 	 	case "PaymAmt2" 
 	 		strVal = strVal & "&txtFromAmt=" & Trim(frm1.fpDoubleSingle5.Value)  	 	
 	 	case "PaymAmt3" 
 	 		strVal = strVal & "&txtFromAmt=" & Trim(frm1.fpDoubleSingle7.Value)  		
 	 	case "ApAmt" 
 	 		strVal = strVal & "&txtFromAmt=" & Trim(frm1.fpDoubleSingle9.Value)  	 		 		
 		End Select				 	 		
 		
		If frm1.txtAcqdt.text = "" Then		   
			strVal = strVal & "&txtAppDt=" & UNIDateClientFormat("parent.gServerBaseDate")
		Else
			strVal = strVal & "&txtAppDt=" & UniConvDateToYYYYMMDD(frm1.txtAcqDt.text, gDateFormat, parent.gServerDateType) '☆: 조회 조건 데이타 
		End If	
    
 		Call RunMyBizASP(MyBizASP, strVal) 	
 	end if

 	fncGetLocalAmt = True	

End Function

'==========================================================================================
'   Event Name :vspdData_KeyPress
'   Event Desc :
'==========================================================================================

Sub vspdData2_KeyPress(index , KeyAscii )
     lgBlnFlgChgValue = True                                                 '⊙: Indicates that value changed
End Sub
Sub vspdData_KeyPress(index , KeyAscii )
     lgBlnFlgChgValue = True                                                 '⊙: Indicates that value changed
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    Dim LngLastRow    
    Dim LngMaxRow     
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
     '----------  Coding part  ------------------------------------------------------------- 
    LngLastRow = NewTop + 30
    LngMaxRow = frm1.vspdData2.MaxRows
    
    If LngLastRow = frm1.vspdData2.MaxRows Then
        Call DbQuery
    End If    
End Sub

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    Dim LngLastRow    
    Dim LngMaxRow     
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
     '----------  Coding part  ------------------------------------------------------------- 
    LngLastRow = NewTop + 30
    LngMaxRow = frm1.vspdData.MaxRows
    
    If LngLastRow = frm1.vspdData.MaxRows Then
        Call DbQuery2
    End If    
End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 
'==========================================================================================

Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)
Dim strTemp
Dim intPos1
Dim strCard

	With frm1.vspdData2 

		If Row > 0 then
			IF Row > 0 And Col = C_BankAcctPopup Then
				.Col = C_BankAcct
				.Row = Row
				strTemp = Trim(.text)

				.col = C_RcptType
				strCard = .text

				Call OpenBankAcct(strTemp, strCard)

			Elseif Row > 0 And Col = C_NoteNoPopup Then
				.Col = C_NoteNo
				.Row = Row
				strTemp = Trim(.text)
				.col = C_RcptType
				strCard = .text

				Call OpenNoteNo(strTemp, strCard)
			Elseif Row > 0 And Col = C_RcptTypePopup Then
				.Col = C_RcptType
				.Row = Row
				Call OpenPopup(.Text, 6)
			End If
		End If

	End With
End Sub

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)
Dim strTemp
Dim intPos1

	With frm1.vspdData

		If Row > 0 then
			if  Col = C_AcctPop Then
				.Col = C_AcctCd
				.Row = Row

				Call OpenAcctCd()

			elseif Col = C_DeptPop Then
				.Col = C_DeptCd
				.Row = Row

				Call OpenDeptCd(.text)
			end if
		End If
	End With

End Sub

Sub txtBpCd_onblur()
	if frm1.txtBpCd.value = "" then
		frm1.txtBpNm.value = ""
	end if
End Sub
Sub txtDeptCd_onblur()
	if frm1.txtDeptCd.value = "" then
		frm1.txtDeptNm.value = ""
	end if
End Sub

	

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 
    Dim IntRetCD 
    Dim var_i, var_m

    FncQuery = False


    Err.Clear 
    '-----------------------
    'Check previous data area
    '-----------------------

    ggoSpread.Source = frm1.vspdData2
    var_i = ggoSpread.SSCheckChange

    ggoSpread.Source = frm1.vspdData
    var_m = ggoSpread.SSCheckChange

    If lgBlnFlgChgValue = True Or var_i = True or var_m = True    Then    
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X") '☜ 바뀐부분 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

	call ClickTab1()

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field

    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData

    ggoSpread.Source = frm1.vspdData2
	ggospread.ClearSpreadData
        
	Call SetDefaultVal

    Call InitVariables															'⊙: Initializes local global variables

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery																'☜: Query db data

    if frm1.vspddata.maxrows = 0 then	
       frm1.txtAcqNo.value = ""
    end if

    FncQuery = True	
    
    													'⊙: Processing is OK
	   
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
    If lgBlnFlgChgValue = True  Or ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"X","X") '☜ 바뀐부분 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")                                         '⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")


    Call ggoOper.LockField(Document, "N")

    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData

    ggoSpread.Source = frm1.vspdData2
	ggospread.ClearSpreadData
    
    Call InitVariables

    'Call InitSpreadSheet("A")		'스프레드시트 초기화 막음 
    'Call InitSpreadSheet("B")		'스프레드시트 초기화 막음 

	Call ClickTab1		'sstData.Tab = 1

    Call SetToolBar("1110110100101111")


    Call SetDefaultVal

	call txtDocCur_OnChangeASP()   

    
    FncNew = True                                                           '⊙: Processing is OK
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncDelete() 

    Dim IntRetCD 

    FncDelete = False                                                      '⊙: Processing is NG
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
		IntRetCD = DisplayMsgBox("900002","X","X","X")  '☜ 바뀐부분 
        Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")  '☜ 바뀐부분 
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
    Dim var_i,var_m
    Dim varApDueDt
    Dim varAcqDt
    Dim varGLDt
    'Dim varIssuedDt
    
    FncSave = False                                                         '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    On Error Resume Next                                                    '☜: Protect system from crashing
    
    '-----------------------
    'Precheck area
    '-----------------------    
	''if frm1.vspdData2.MaxRows < 1 then
	''	IntRetCD = DisplayMsgBox("117298","X","X","X")  ''유형별 지출내역을 입력하십시오.
	''	Exit Function
	''end if
	if frm1.vspdData.MaxRows < 1 then
		IntRetCD = DisplayMsgBox("117294","X","X","X")  ''자산세부내역을 입력하십시오.
		Exit Function
	end if
		
    ggoSpread.Source = frm1.vspdData2
    var_i = ggoSpread.SSCheckChange

    ggoSpread.Source = frm1.vspdData
    var_m = ggoSpread.SSCheckChange

    If lgBlnFlgChgValue = False and var_i = False and var_m = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")  '☜ 바뀐부분 
        'Call MsgBox("No data changed!!", vbInformation)
        Exit Function
    End If
    
    '-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "2") then                                   '⊙: Check contents area
       Exit Function
    End If

    ggoSpread.Source = frm1.vspdData2
    If Not ggoSpread.SSDefaultCheck Then	
		Exit Function
    End if
	    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then	
		Exit Function
    End if            
  
	if IsNull(frm1.txtApDuedt.text) then
		frm1.txtApDuedt.text = ""		
	end if
	
	if IsNull(frm1.txtIssuedDt.text) then
		frm1.txtIssuedDt.text = ""
	end if
	
'	if frm1.txtIssuedDt.text = "" then
'		frm1.txtissuedDt.text = frm1.txtAcqDt.text
'	end if

	varAcqDt   = UniConvDateToYYYYMMDD(frm1.txtAcqDt.Text, gDateFormat,"")
	varApDueDt = UniConvDateToYYYYMMDD(frm1.txtApDueDt.Text, gDateFormat,"")
	varGLDt    = UniConvDateToYYYYMMDD(frm1.txtGLDt.Text, gDateFormat,"")
	varIssuedDt= UniConvDateToYYYYMMDD(frm1.txtIssuedDt.Text, gDateFormat,"")

    If UNICDbl(frm1.txtApAmt.text) > 0 Then
    	If varApDueDt = "" or varApDueDT <= "19000101" Then		
			Call DisplayMsgBox("117292", "X", "X", "X")
			Exit Function
			End if
	End if		
	
	If CompareDateByFormat(frm1.txtAcqDt.text,frm1.txtGLDt.text,frm1.txtAcqDt.Alt,frm1.txtGLDt.Alt, _
        	               "970025",frm1.txtAcqDt.UserDefinedFormat,parent.gComDateType, true) = False Then
	   frm1.txtAcqDt.focus
	   Exit Function
	End If

    CAll DbSave				                                                
    
    FncSave = True                                                          
    
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 
    Dim IntRetCD
    
    If gSelframeFlg = TAB1 Then	 
		frm1.vspdData.ReDraw = False

		if frm1.vspdData.MaxRows < 1 then Exit Function

		ggoSpread.Source = frm1.vspdData	
		ggoSpread.CopyRow
		SetSpreadColor_Master frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow, "insert"

		frm1.vspdData.Col  = C_AsstNo
		frm1.vspdData.Text = ""
		frm1.vspdData.Col  = C_AsstNm
		frm1.vspdData.Text = ""

		frm1.vspdData.ReDraw = True

	Elseif  gSelframeFlg = TAB2 Then
		frm1.vspdData2.ReDraw = False

		if frm1.vspdData2.MaxRows < 1 then Exit Function

		ggoSpread.Source = frm1.vspdData2
		ggoSpread.CopyRow
		SetSpreadColor_Item frm1.vspdData2.ActiveRow, frm1.vspdData2.ActiveRow

    	frm1.vspdData2.Col = C_RcptType

		call subVspdSettingChange(frm1.vspdData2.ActiveRow,frm1.vspdData2.Text)

'    	frm1.vspdData2.Col = C_RcptType
'		frm1.vspdData2.Text = ""
'		frm1.vspdData2.Col = C_RcptTypeNm
'		frm1.vspdData2.Text = ""

		MaxSpreadVal frm1.vspdData2.ActiveRow
		frm1.vspdData2.ReDraw = True
	End if

End Function


'========================================================================================
' Function Name : FncPaste
' Function Desc : This function is related to Paste Button of Main ToolBar
'========================================================================================
Function FncPaste() 
     ggoSpread.SpreadPaste
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 

	If  gSelframeFlg = TAB1 Then  'Acq Item 단 
	    if frm1.vspdData.MaxRows < 1 then	 Exit Function
		ggoSpread.Source = frm1.vspdData
		
    ElseIf gSelframeFlg = TAB2 Then  'Master단 
		ggoSpread.Source = frm1.vspdData2
	End If	
	
	ggoSpread.EditUndo                                                  '☜: Protect system from crashing

End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow(Byval pvRowCnt) 
	Dim varMaxRow
	Dim strDoc
	Dim varXrate


	Dim imRow
	FncInsertRow = False

	if IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
	else
		imRow = AskSpdSheetAddRowcount()

	If ImRow="" then
		Exit Function
	End If
	End If
		
    if gSelframeFlg = TAB1 Then

		with frm1
			varMaxRow = .vspdData.MaxRows 
			''''''''lgBlnFlgChgValue = True                            'Indicates that value changed
			.vspdData.focus
		
			ggoSpread.Source = .vspdData
			.vspdData.ReDraw = False
		
			ggoSpread.InsertRow ,imRow
			.vspdData.ReDraw = True
		
			SetSpreadColor_Master .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1,"insert"				
	
		end with
         
     ElseIf   gSelframeFlg = TAB2 Then        '''' Acq Item
     
		with frm1
		
			IF frm1.cboAcqFg.value = "03" then '기초자산인 경우 
				exit Function
			End If
			.vspdData2.focus
			
			varMaxRow = .vspdData2.MaxRows		
		
			ggoSpread.Source = .vspdData2
			.vspdData2.ReDraw = False
		
			ggoSpread.InsertRow ,imRow
			.vspdData2.ReDraw = True
		
			SetSpreadColor_Item .vspdData2.ActiveRow, .vspdData2.ActiveRow + imRow - 1			
			MaxSpreadVal .vspdData2.ActiveRow				

		end with
	END if
'	Call ggoOper.LockField(Document, "Q")                                           '⊙: This function lock the suitable field

End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
    Dim lDelRows 
    Dim lTempRows 

    ''If gSelframeFlg <> TAB2 Then
	''	Call ClickTab2		'sstData.Tab = 1
    ''End If
	if gSelframeFlg = TAB1 Then	
		frm1.vspdData.focus
    	ggoSpread.Source = frm1.vspdData
		if frm1.vspdData.MaxRows < 1 then Exit Function
	
		lDelRows = ggoSpread.DeleteRow    
    
    Elseif gSelframeFlg = TAB2 Then
		frm1.vspdData2.focus
		ggoSpread.Source = frm1.vspdData2
		
		if frm1.vspdData2.MaxRows < 1 then Exit Function		
		lDelRows = ggoSpread.DeleteRow		
		
    End if    
End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
	Parent.fncPrint()    
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
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLEMULTI)												'☜: 화면 유형 
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 화면 속성, Tab유무 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI , True)                                                    '☜: Protect system from crashing
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
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
	Dim IntRetCD
	FncExit = False	
		
	If lgBlnFlgChgValue = True then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")   '☜ 바뀐부분 
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 종료 하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then
		
			Exit Function
		End If
    End If
    FncExit = True
End Function


'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================

Function DbDelete() 

    DbDelete = False														'⊙: Processing is NG
    
    Dim strVal
        
    strVal = BIZ_PGM_DEL_ID & "?txtMode=" & parent.UID_M0003
    strVal = strVal & "&txtAcqNo=" & Trim(frm1.txtAcqNo.value)			'☜: 삭제 조건 데이타 
    strVal = strVal & "&cboAcqFg=" & Trim(frm1.cboAcqFg.value)				'☜: 삭제 조건 데이타 
    strVal = strVal & "&txtGlNo="   & Trim(frm1.txtGlNo.value)				'☜: 삭제 조건 데이타 
    strVal = strVal & "&txtApNo=" & Trim(frm1.txtApNo.value)				'☜: 삭제 조건 데이타 
    
    Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    
    DbDelete = True                                                         '⊙: Processing is NG

End Function


'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================

Function DbDeleteOk()												'☆: 삭제 성공후 실행 로직 
	Call FncNew()
End Function


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 
    Dim strVal
    
    DbQuery = False                                                         '⊙: Processing is NG
	Err.Clear
	    
	Call LayerShowHide(1)	

	If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001							'☜: 
		strVal = strVal & "&txtAcqNo="        & Trim(frm1.htxtAcqNo.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&lgStrPrevKey_i=" & lgStrPrevKey_i
		strVal = strVal & "&txtMaxRows_i="    & frm1.vspdData.MaxRows
		strVal = strVal & "&txtMaxRows_m="    & frm1.vspdData2.MaxRows
	Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001							'☜: 
		strVal = strVal & "&txtAcqNo="        & Trim(frm1.txtAcqNo.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&lgStrPrevKey_i=" & lgStrPrevKey_i
		strVal = strVal & "&txtMaxRows_i="    & frm1.vspdData.MaxRows
		strVal = strVal & "&txtMaxRows_m="    & frm1.vspdData2.MaxRows
	End If
    
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
 	
    DbQuery = True                                                          '⊙: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()													'☆: 조회 성공후 실행로직	

	Dim varData
	Dim iRow	
			
    lgIntFlgMode = parent.OPMD_UMODE													'⊙: Indicates that current mode is Update mode
    
    Call ggoOper.LockField(Document, "Q")										'⊙: This function lock the suitable field    	
	Call SetToolBar("1111111100111111")	'1111111100011111

	
	Call SetSpreadColor_Item(-1 ,-1)
	
	With frm1				
	
		.vspdData.Redraw = False
	
'		For iRow = 0 To frm1.vspdData.MaxRows
			Call SetSpreadColor_Master(-1, -1,"query")		
'		Next
		
		.vspdData.Redraw = True			
	End With
	
	With frm1				
	
		.vspdData2.Redraw = False
	
		For iRow = 0 To frm1.vspdData2.MaxRows
	
			.vspdData2.Col = C_RcptType		
			.vspdData2.Row = iRow
			
			varData = frm1.vspdData2.text
			Call subVspdSettingChange(iRow,varData)   ''''Rcpt Type별 입력필수 필드 표시 
		Next
		
		.vspdData2.Redraw = True			
	End With

	
	Call txtDocCur_OnChangeASP()
	'Call txtVatType_OnChange		'로드시 부가세 필수 구분 (cboAcqFg_onChange()에 포함되므로 주석)
	
	lgBlnFlgChgValue = False
	
	call cboAcqFg_onChange()	'부가세 유형 및 금액의 쿼리시 무상기증일경우 disable시킴 
	'varAcqFg = frm1.cboAcqFg.value
	'Call SetProctedField(varAcqFg)	
	lgBlnFlgChgValue = False

	
	'SetGridFocus()				
	'SetGridFocus2()

'조회되는것이 기초자산이면 조회할수 없도록.... 조건이 "" 인 이유는 value값이 03를 가질수 없고 ""값을 가지고 온다.

'	IF frm1.cboAcqFg.value = "" Then
'		Call DisplayMsgBox("117217", "X", "X", "X")
'		Fncnew()
'	Else
'		Call dbquery2
'	End If


End Function

 '========================================================================================
'    Function Name : InitData()
'    Function Desc : 
'   ======================================================================================== 
Sub InitData()
	Dim intRow
	Dim intIndex 
	dim temp
	
	With frm1.vspdData2
	
		For intRow = 1 To .MaxRows			
			.Row  = intRow			
			.Col	 = C_RcptType			
			intIndex = .Value

			.Col     = C_RcptTypeNm
			.Value   = intIndex
		Next	
	End With	

'추가부분.
call cboAcqFg_onChange()

End Sub

'========================================================================================
' Function Name : DbQuery2
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery2() 
    
    DbQuery2 = False                                                         '⊙: Processing is NG
	Err.Clear
	    
	Call LayerShowHide(1)
	
    Dim strVal

	If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID2 & "?txtMode=" & parent.UID_M0001							'☜: 
		strVal = strVal & "&txtAcqNo="        & Trim(frm1.htxtAcqNo.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&lgStrPrevKey_m=" & lgStrPrevKey_m
		strVal = strVal & "&txtMaxRows_m="    & frm1.vspdData.MaxRows
	Else
		strVal = BIZ_PGM_QRY_ID2 & "?txtMode=" & parent.UID_M0001							'☜: 
		strVal = strVal & "&txtAcqNo="        & Trim(frm1.txtAcqNo.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&lgStrPrevKey_m=" & lgStrPrevKey_m
		strVal = strVal & "&txtMaxRows_m="    & frm1.vspdData.MaxRows
	End If
    
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

    DbQuery2 = True                                                          '⊙: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbQueryOk2
' Function Desc : DbQuery2가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk2()													'☆: 조회 성공후 실행로직 
	Dim varAcqFg	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE													'⊙: Indicates that current mode is Update mode

    Call ggoOper.LockField(Document, "Q")										'⊙: This function lock the suitable field    

	Call SetToolBar("1101111100111111")	'1111111100011111
	
	lgBlnFlgChgValue = False
	
	Call SetSpreadColor_Master(-1, -1, "query")
	
	varAcqFg = frm1.cboAcqFg.value 	
	Call SetProctedField(varAcqFg)	
End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 

    Dim pAs0021 'As New As0021ManageSvr
    Dim IntRows 
    Dim IntCols 
    Dim vbIntRet 
    Dim lStartRow 
    Dim lEndRow 
    Dim boolCheck 
    Dim lGrpcnt 
	Dim strVal, strDel
	Dim ApAmt, PayAmt
	
    DbSave = False                                                          '⊙: Processing is NG    

	Call LayerShowHide(1)
	
	With frm1
		.txtMode.value    = parent.UID_M0002										'☜: 저장 상태 
		.txtFlgMode.value = lgIntFlgMode									'☜: 신규입력/수정 상태			
	End With
	
    '-----------------------
    'Data manipulate area
    '-----------------------
    ' Data 연결 규칙 
    ' 0: Sheet명, 1: Flag , 2: Row위치, 3~N: 각 데이타 

    lGrpCnt = 1    
	strVal = ""
	strDel = ""
    
    '-----------------------------
    '   Acq item Part
    '-----------------------------
    With frm1.vspdData2
	    
    For IntRows = 1 To .MaxRows
    
		.Row = IntRows
		.Col = 0		

		If .Text = ggoSpread.DeleteFlag Then
			strDel = strDel & "D" & parent.gColSep 		'D=Delete
		ElseIf .Text = ggoSpread.UpdateFlag Then
			strVal = strVal & "U" & parent.gColSep 		'U=Update
		Else
			strVal = strVal & "C" & parent.gColSep 		'C=Create
		End If		
		
		.Col = 0
		
		Select Case .Text		    
		        
		    Case ggoSpread.DeleteFlag

				.Col = C_Seq
				strDel = strDel & Trim(.Text) & parent.gRowSep				    '마지막 데이타는 Row 분리기호를 넣는다 
					
				lGrpcnt = lGrpcnt + 1            
		    
		    Case Else
		        .Col = C_Seq	
		        strVal = strVal & Trim(.Text) & parent.gColSep

		        .Col = C_RcptType   
		        strVal = strVal & Trim(.Text) & parent.gColSep

				.Col = C_Amt		

		        strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
		        
		        .Col = C_LocAmt	
		        strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
   		        
		        .Col = C_NoteNo		
		        strVal = strVal & Trim(.Text) & parent.gColSep

		        .Col = C_BankAcct			
		        strVal = strVal & Trim(.Text) & parent.gColSep & parent.gRowSep			        		        
		           		        
		        lGrpCnt = lGrpCnt + 1

		End Select

    Next

	End With
	
    frm1.txtMaxRows_i.value  = lGrpCnt-1										'Spread Sheet의 변경된 최대갯수 
	frm1.txtSpread_i.value   = strDel & strVal								'Spread Sheet 내용을 저장 

    lGrpCnt = 1    
	strVal = ""
	strDel = ""

	With frm1.vspdData
		
	For IntRows = 1 To .MaxRows
    
		.Row = IntRows
		.Col = 0

		If .Text = ggoSpread.DeleteFlag Then
			.Col = C_AcctCd  
			strDel = strDel & Trim(.Text) & parent.gColSep							'0
			strDel = strDel & "D" & parent.gColSep & frm1.hORGCHANGEID.value & parent.gColSep		'12		D=Delete
		ElseIf .Text = ggoSpread.InsertFlag Then
			.Col = C_AcctCd  
			strVal = strVal & Trim(.Text) & parent.gColSep							'0
			strVal = strVal & "C" & parent.gColSep & frm1.hORGCHANGEID.value & parent.gColSep		'12		U=Update			
		Else
			.Col = C_AcctCd  
			strVal = strVal & Trim(.Text) & parent.gColSep							'0
			strVal = strVal & "U" & parent.gColSep & frm1.hORGCHANGEID.value & parent.gColSep		'12		C=Create
		End If
		
		.Col = 0

		Select Case  .Text 
			Case ggoSpread.DeleteFlag			
				

				.Col = C_DeptCd													'3	A073_IG1_I3_dept_cd
				strDel = strDel & Trim(.Text) & parent.gColSep

				strDel = strDel & Trim(.Text) & parent.gRowSep				'⊙: 마지막 데이타는 Row 분리기호를 넣는다 
				
				
				lGrpCnt = lGrpCnt + 1
            Case Else 
				.Col = C_DeptCd													'3	A073_IG1_I3_dept_cd
				strVal = strVal & Trim(.Text) & parent.gColSep

				.Col = C_AsstNo													'4	A073_IG1_I4_asst_no
				strVal = strVal & Trim(.Text) & parent.gColSep		
				
				.Col = C_AsstNm													'5	A073_IG1_I4_asst_nm
				strVal = strVal & Trim(.Text) & parent.gColSep		

				.Col = C_AcqAmt													'6	A073_IG1_I4_acq_amt
				strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep

				.Col = C_AcqLocAmt												'7	A073_IG1_I4_acq_loc_amt
				strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep

				.Col = C_AcqQty													'8	A073_IG1_I4_acq_qty
				strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
				
				.Col = C_ResAmt													'9	A073_IG1_I4_res_amt
				strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep

				.Col = C_RefNo													'10	A073_IG1_I4_ref_no
				strVal = strVal & Trim(.Text) & parent.gColSep

				.Col = C_Desc													'11	A073_IG1_I4_asset_desc
				strVal = strVal & Trim(.Text) & parent.gColSep

				strVal = strVal & parent.gColSep	'12	A073_IG1_I4_reg_dt
				strVal = strVal & parent.gColSep	'13	A073_IG1_I4_spec
				strVal = strVal & parent.gColSep	'14	A073_IG1_I4_doc_cur
				strVal = strVal & parent.gColSep	'15	A073_IG1_I4_xch_rate
				strVal = strVal & parent.gColSep	'16	A073_IG1_I4_inv_qty
				strVal = strVal & parent.gColSep	'17	A073_IG1_I4_tax_dur_yrs
				strVal = strVal & parent.gColSep	'18	A073_IG1_I4_cas_dur_yrs
				strVal = strVal & parent.gColSep	'19	A073_IG1_I4_tax_end_l_term_cpt_tot_amt
				strVal = strVal & parent.gColSep	'20	A073_IG1_I4_cas_end_l_term_cpt_tot_amt
				strVal = strVal & parent.gColSep	'21	A073_IG1_I4_tax_end_l_term_depr_tot_amt
				strVal = strVal & parent.gColSep	'22	A073_IG1_I4_cas_end_l_term_depr_tot_amt
				strVal = strVal & parent.gColSep	'23	A073_IG1_I4_tax_end_l_term_bal_amt
				strVal = strVal & parent.gColSep	'24	A073_IG1_I4_cas_end_l_term_bal_amt
				strVal = strVal & parent.gColSep	'25	A073_IG1_I4_tax_depr_sts
				strVal = strVal & parent.gColSep	'26	A073_IG1_I4_cas_depr_sts
				strVal = strVal & parent.gColSep	'27	A073_IG1_I4_tax_depr_end_yyyymm
				strVal = strVal & parent.gColSep	'28	A073_IG1_I4_cas_depr_end_yyyymm
				strVal = strVal & parent.gColSep	'29	A073_IG1_I4_start_depr_yymm
				strVal = strVal & parent.gColSep	'30	A073_IG1_I4_tax_dur_mnth
				strVal = strVal & parent.gColSep	'31	A073_IG1_I4_cas_dur_mnth
				
				strVal = strVal & parent.gRowSep

				lGrpCnt = lGrpCnt + 1

		End Select

    Next

	End With

	frm1.txtMaxRows_m.value  = lGrpCnt-1										'☜: Spread Sheet의 변경된 최대갯수 
	frm1.txtSpread_m.value = strDel & strVal									'☜: Spread Sheet 내용을 저장 



	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)								'☜: 저장 비지니스 ASP 를 가동 
        
    DbSave = True                                                           ' ⊙: Processing is OK
    
End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================

Function DbSaveOk()													'☆: 저장 성공후 실행 로직 

	Dim iAcq_no
	
	iAcq_no = frm1.txtAcqNo.value
   
    Call InitVariables	
    Call ClickTab1()    
    
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")										'⊙: Clear Contents  Field
    'Call ggoOper.ClearField(Document, "2")
    
    'frm1.vspdData.MaxRows  = 0
    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
    'frm1.vspdData2.MaxRows = 0
    ggoSpread.Source = frm1.vspdData2
	ggospread.ClearSpreadData		'Buffer Clear
        
    'call initspreadsheet		'스프레드시트 초기화 막음 

    Call initspreadsheet("A")		'스프레드시트 초기화 막음 
    Call initspreadsheet("B")		'스프레드시트 초기화 막음 
    
'    Call SetSpreadLock("I","",-1)
 '   Call SetSpreadLock("M","insert",-1)
	'''''Call InitComboBox_acqfg
	'Call InitComboBox_rcpt	

	frm1.txtAcqNo.value = iAcq_no
	    	
	call dbquery()

End Function


'========================================================================================
' Function Name : MaxSpreadVal
' Function Desc : 
'========================================================================================

Function MaxSpreadVal(byval Row)
  Dim iRows
  Dim MaxValue  
  Dim tmpVal

	MAxValue = 0

		with frm1
			For iRows = 1 to  .vspdData2.MaxRows
				.vspddata2.row = iRows
		        .vspddata2.col = C_Seq

				if .vspdData2.Text = "" then 
					tmpVal = 0
				else
  					tmpVal = cdbl(.vspdData2.value)
				end if

				if tmpval > MaxValue   then
					MaxValue = cdbl(tmpVal)
				end if
			Next

			MaxValue = MaxValue + 1

			.vspddata2.row = row
			.vspddata2.col = C_Seq
			.vspddata2.text = MaxValue
		end with
		
end Function

'==========================================================================================

Sub txtDeptCd_OnChange()
        
    Dim strSelect
	Dim strFrom
	Dim strWhere 	
    Dim IntRetCD 
	Dim arrVal1
	Dim arrVal2
	Dim ii
	Dim jj

	If Trim(frm1.txtGLDt.Text = "") Then    
		Exit sub
    End If
    lgBlnFlgChgValue = True

		'----------------------------------------------------------------------------------------
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd =  " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.Value)), "''", "S") 
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtGLDt.Text, parent.gDateFormat,""), "''", "S") & "))"			
		
	
		
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			If lgIntFlgMode <> parent.OPMD_UMODE Then
				IntRetCD = DisplayMsgBox("124600","X","X","X")  
			End If			
			frm1.txtDeptCd.Value = ""
			frm1.txtDeptNm.Value = ""
			frm1.hOrgChangeId.Value = ""
		Else 
		
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)
					
			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))			
				frm1.hOrgChangeId.Value = Trim(arrVal2(2))
			Next	
			
		End If
	
		'----------------------------------------------------------------------------------------

End Sub
'=======================================================================================================
Sub txtGLDt_onBlur()
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii
	Dim arrVal1, arrVal2
	Dim IntRows
	If Trim(frm1.txtDeptCd.value) <> "" and Trim(frm1.txtGLDt.Text <> "") Then
		strSelect	=			 " Distinct org_change_id "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtGLDt.Text, parent.gDateFormat,""), "''", "S") & "))"			

		IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
					
		If IntRetCD = False  OR Trim(Replace(lgF0,Chr(11),"")) <> Trim(frm1.hOrgChangeId.value) Then
			'IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
			frm1.txtDeptCd.focus
		End If
	End If

    lgBlnFlgChgValue = True
End Sub


Sub txtApDueDt_Change()
       lgBlnFlgChgValue = true
End Sub
Sub txtAcqDt_Change()
    lgBlnFlgChgValue = true
End Sub

Sub txtVatAmt_Change() 'onblur
   lgBlnFlgChgValue = true
End Sub
Sub txtVatLocAmt_Change()	'onblur
   lgBlnFlgChgValue = true
End Sub
Sub txtXchRate_Change()	'onblur
   lgBlnFlgChgValue = true
End Sub

Sub cboAcqFg_onblur()	'onblur
   lgBlnFlgChgValue = true
End Sub

Sub txtVatRate_Change()	'부가세율 
	lgBlnFlgChgValue = true
End Sub

Sub txtIssuedDt_Change()
       lgBlnFlgChgValue = true
End Sub

Sub txtReportAreaCd_Change()	'신고사업장 
	lgBlnFlgChgValue = true
End Sub

'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX 
'====================================================================================================
Sub CurFormatNumericOCX()

	With frm1
	'해당되는 금액이 있는 조건부 필드에 대하여 각각 처리 
		'총취득금액 
		ggoOper.FormatFieldByObjectOfCur .txtAcqAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		'미지급금액 
		ggoOper.FormatFieldByObjectOfCur .txtApAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		'부가세금액 
		ggoOper.FormatFieldByObjectOfCur .txtVatAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
	End With
End Sub

'===================================== CurFormatNumSprSheet()  ======================================
'	Name : CurFormatNumSprSheet()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet()

	With frm1
		ggoSpread.Source = frm1.vspdData
	'해당되는 금액이 있는 Grid에 대하여 각각 처리 
		'자산취득내역등록TAB-취득금액 
		ggoSpread.SSSetFloatByCellOfCur C_AcqAmt,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec		
		'출금내역TAB-금액 
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.SSSetFloatByCellOfCur C_Amt,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec		
	End With
End Sub  

'==========================================================================================
'   Event Name : txtDocCur_OnChangeASP
'   Event Desc : 
'==========================================================================================
Sub txtDocCur_OnChangeASP()
 
    IF CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							

		Call CurFormatNumericOCX()
		Call CurFormatNumSprSheet()

	END IF	    
End Sub


'=======================================================================================================
'   Event Name : SetGridFocus
'   Event Desc :
'=======================================================================================================    
Sub SetGridFocus()	
    
	Frm1.vspdData.Row = 1
	Frm1.vspdData.Col = 1
	Frm1.vspdData.Action = 1	

End Sub
Sub SetGridFocus2()	
    
	Frm1.vspdData2.Row = 1
	Frm1.vspdData2.Col = 1
	Frm1.vspdData2.Action = 1	

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
	Dim indx
	Dim iRow
	Dim IntRetCD
	Dim varData
	
	Select Case Trim(UCase(gActiveSpdSheet.Name))
		Case "VSPDDATA" 
			ggoSpread.Source = frm1.vspdData
			Call ggoSpread.RestoreSpreadInf()
			Call InitSpreadSheet("A")
			Call SetSpreadColor_Master(-1, -1, "query")

			Call ggoSpread.ReOrderingSpreadData()

		Case "VSPDDATA2"
			ggoSpread.Source = frm1.vspdData2
			Call ggoSpread.RestoreSpreadInf()
			Call InitSpreadSheet("B")
			Call ggoSpread.ReOrderingSpreadData()
			'Call initData()
			Call SetSpreadColor_Item(-1 ,-1)
				
			frm1.vspdData2.Redraw = False
			For iRow = 0 To frm1.vspdData2.MaxRows
	
				frm1.vspdData2.Col = C_RcptType		
				frm1.vspdData2.Row = iRow
				
				varData = frm1.vspdData2.text

				Call subVspdSettingChange(iRow,varData )   ''''Rcpt Type별 입력필수 필드 표시 
			Next
			frm1.vspdData2.Redraw = True
			
	End Select

End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
</HEAD>
<!--'#########################################################################################################
'       					6. Tag부 
'######################################################################################################### -->
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>고정자산취득내역등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>출금내역</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT><A HREF="VBSCRIPT:OpenPopupTempGL()">결의전표</A>&nbsp;|&nbsp;<A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A></TD>					
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<DIV ID="TabDiv" STYLE="FlOAT: left; HEIGHT:100%; OVERFLOW:auto; WIDTH:100%;" SCROLL=no>
				<TABLE <%=LR_SPACE_TYPE_20%>>
					<TR>
						<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
					</TR>
					<TR>
						<TD HEIGHT=20 WIDTH=100%>
							<FIELDSET CLASS="CLSFLD">
								<TABLE <%=LR_SPACE_TYPE_40%>>
									<TR>
										<TD CLASS="TD5" NOWRAP>취득번호</TD>
										<TD CLASS="TD6"><INPUT NAME="txtAcqNo" TYPE="Text" MAXLENGTH=18 tag="12XXXU" ALT="취득번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcqNo1" ALIGN=Top TYPE="BUTTON" ONCLICK="vbscript:OpenAcqNoInfo"></TD>
										<TD CLASS="TD6"></TD>
										<TD CLASS="TD6"></TD>
									</TR>
								</TABLE>
							</FIELDSET>
						</TD>
					</TR>
					<TR>
						<TD <%=HEIGHT_TYPE_03%>></TD>
					</TR>
					<TR HEIGHT=100%>
						<TD>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
								    <TD CLASS="TD5" NOWRAP>취득일자</TD>																							    
									<TD CLASS="TD6" NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 name=txtAcqDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="22" ALT="취득일자"> </OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS="TD5" NOWRAP>전표일자</TD>																							    
									<TD CLASS="TD6" NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime3 name=txtGLDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="22" ALT="전표생성일자"> </OBJECT>');</SCRIPT>
									</TD>
								</TR>	
								<TR>
									<TD CLASS="TD5" NOWRAP>취득부서</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtDeptCd" SIZE=10 MAXLENGTH=10 tag="22XXXU" ALT="취득부서"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcqNo1" ALIGN=Top TYPE="BUTTON" ONCLICK="vbscript:OpenDept()">&nbsp;<INPUT TYPE=TEXT NAME="txtDeptNm" SIZE=22 tag="24"  alt = "부서명"></TD>												
									<TD CLASS="TD5"></TD>
									<TD CLASS="TD6"></TD>
								</TR>	
								<TR>
									<TD CLASS=TD5 NOWRAP>취득구분</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboAcqFg" STYLE="Width:150px;" tag="22" ALT="취득구분"><!--OPTION VALUE=""></OPTION--></SELECT></TD>
									<TD CLASS=TD5 NOWRAP>거래처</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBpCd" ALT="거래처" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" ALIGN=Top TYPE="BUTTON" ONCLICK="vbscript:call OpenBp(frm1.txtBpCd.value,1)"> <INPUT NAME="txtBpNm" TYPE="Text" SIZE = 22 tag="24"></TD>
								</TR>
<%	If gIsShowLocal <> "N" Then	%>								
								<TR>
									<TD CLASS=TD5 NOWRAP>거래통화</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDocCur" ALT="거래통화" TYPE="Text" MAXLENGTH=3 SIZE=10 tag="22XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCur" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenCurrency()"></TD>
									<TD CLASS=TD5 NOWRAP>환율</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle9 name="txtXchRate" style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 150px" title="FPDOUBLESINGLE" ALT="환율" tag="21X5Z"> </OBJECT>');</SCRIPT>&nbsp;
	                                </TD>
								</TR>
<%	ELSE %>
<INPUT TYPE=HIDDEN NAME="txtDocCur"><INPUT TYPE=HIDDEN NAME="txtXchRate">
<%	End If %>								
								<TR>
									<TD CLASS=TD5 NOWRAP>총취득금액</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 name=txtAcqAmt style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 150px" title="FPDOUBLESINGLE" ALT="총취득금액" tag="24X2"> </OBJECT>');</SCRIPT>&nbsp;																				
	                                </TD>
<%	If gIsShowLocal <> "N" Then	%>
									<TD CLASS=TD5 NOWRAP>총취득금액(자국)</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtAcqLocAmt style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 150px" title="FPDOUBLESINGLE" ALT="총취득금액(자국)" tag="24X2"> </OBJECT>');</SCRIPT>&nbsp;
	                                </TD>
<%	ELSE %>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=HIDDEN NAME="txtAcqLocAmt"></TD>
<%	End If %>
								</TR>	
								<TR>
									<TD CLASS="TD5" NOWRAP>미지급금계정</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtApAcctCd" SIZE=10 MAXLENGTH=20 tag="21XXXU" ALT="미지급금계정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnApAcctCd" ALIGN=Top TYPE="BUTTON" ONCLICK="vbscript:OpenApAcct()">&nbsp;<INPUT TYPE=TEXT NAME="txtApAcctNm" SIZE=22 tag="24"  alt = "미지급금계정명"></TD>
									<TD CLASS=TD5 NOWRAP>신용카드번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCardNo" SIZE=20 MAXLENGTH=20 tag="21XXXU" ALT="신용카드번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCardNo" ALIGN=Top TYPE="BUTTON" ONCLICK="vbscript:OpenCardNo()">&nbsp;<INPUT TYPE=TEXT NAME="txtCardNm" SIZE=22 tag="24"  alt = "신용카드명"></TD>
								</TR>	
								<TR>                    
									<TD CLASS=TD5 NOWRAP>미지급금액</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle7 name=txtApAmt style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 150px" title="FPDOUBLESINGLE" ALT="미지급금액" tag="24X2"> </OBJECT>');</SCRIPT>&nbsp;
	                                </TD>
<%	If gIsShowLocal <> "N" Then	%>
									<TD CLASS=TD5 NOWRAP>미지급금액(자국)</TD>
									<TD CLASS=TD6 NOWRAP>									
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle8 name=txtApLocAmt style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 150px" title="FPDOUBLESINGLE" ALT="미지급금액(자국)" tag="24X2"> </OBJECT>');</SCRIPT>&nbsp;
 									</TD>
<%	ELSE %>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=HIDDEN NAME="txtApLocAmt"></TD>
<%	End If %>
								</TR>									
								<TR>	
									<TD CLASS="TD5" NOWRAP>미지급금 만기일자</TD>																							    
									<TD CLASS="TD6" NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtApDueDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="21X1" ALT="미지급금 만기일자"> </OBJECT>');</SCRIPT>											    
									</TD>															
									<TD CLASS="TD5" NOWRAP>미지급금 번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtApNo" SIZE=20 MAXLENGTH=18 tag="24" ALT="미지급금 번호"></TD>								
								</TR>	
<%	If gIsShowLocal <> "N" Then	%>
								<TR>
									<TD CLASS="TD5" NOWRAP>부가세유형</TD>
								    <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtVatType" SIZE=10 MAXLENGTH=10 tag="21XXXU" ALT="부가세유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnVatType" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenVatType()">&nbsp;<INPUT TYPE=TEXT NAME="txtVatTypeNm" SIZE=20 tag="24" ALT="부가세유형"></TD>
									<TD CLASS="TD5" NOWRAP>부가세율</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtVatRate style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 80px" title=FPDOUBLESINGLE ALT="부가세율" tag="21X5Z" id=fpDoubleSingle11></OBJECT>');</SCRIPT>&nbsp;%</TD>
								</TR>
								<TR>                    
									<TD CLASS=TD5 NOWRAP>부가세금액</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 name=txtVatAmt style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 150px" title="FPDOUBLESINGLE" ALT="부가세금액" tag="21X2" onblur="vbscript:txtVatType_OnChange()"> </OBJECT>');</SCRIPT>&nbsp;
	                                </TD>
									<TD CLASS=TD5 NOWRAP>부가세금액(자국)</TD>
									<TD CLASS=TD6 NOWRAP>									
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle6 name=txtVatLocAmt style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 150px" title="FPDOUBLESINGLE" ALT="부가세금액(자국)" tag="21X2"> </OBJECT>');</SCRIPT> &nbsp;
 									</TD>
								</TR>

<%	ELSE %>
<INPUT TYPE=HIDDEN NAME="txtVatAmt"><INPUT TYPE=HIDDEN NAME="txtVatLocAmt"><INPUT TYPE=HIDDEN NAME="txtVatType"><INPUT TYPE=HIDDEN NAME="txtVatTypeNm"><INPUT TYPE=HIDDEN NAME="txtVatRate">
<%	End If %>
								<TR>
									<TD CLASS="TD5" NOWRAP>신고사업장</TD>
								    <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtReportAreaCd" SIZE=10 MAXLENGTH=10 tag="21XXXU" ALT="신고사업장코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnReportAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenReportAreaCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtReportAreaNm" SIZE=20 tag="24" ALT="신고사업장명"></TD>
									<TD CLASS="TD5" NOWRAP>계산서발행일</TD>																							    
									<TD CLASS="TD6" NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime4 name=txtIssuedDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="21" ALT="전표생성일자"> </OBJECT>');</SCRIPT>											    
									</TD>
								</TR>

								<TR>							
									<TD CLASS="TD5" NOWRAP>회계전표번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtGLNo" SIZE=20 MAXLENGTH=18  tag="24" ALT="전표번호"></TD>
									<TD CLASS="TD5" NOWRAP>결의전표번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtTempGLNo" ALT="결의전표번호" TYPE="Text" MAXLENGTH=18 SIZE=25 tag="24" ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>적요</TD>
									<TD CLASS="TD656" NOWRAP COLSPAN=3><INPUT TYPE=TEXT NAME="txtDesc" SIZE=90 MAXLENGTH=128 tag="2X" ALT="적요"></TD>
									<!--<TD CLASS=TD5 NOWRAP>
									<TD CLASS=TD6 NOWRAP>  -->
								</TR>																
								<TR HEIGHT=100%>
									<TD WIDTH="100%" HEIGHT=100% COLSPAN=4>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="2" TITLE="SPREAD" id=fpSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
									</TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
				</TABLE>
			</DIV>
			<!-- 두번째 탭 내용  -->
			<DIV ID="TabDiv" STYLE="DISPLAY: none;" SCROLL=no>
				<TABLE <%=LR_SPACE_TYPE_20%>>
					<TR>
						<TD HEIGHT="100%" NOWRAP>
							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" tag="2" name=vspdData2 width="100%" id=fpSpread2> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
						</TD>
					</TR>
				</TABLE>

			</DIV>			
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>	
	<TR HEIGHT=10>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread_m" tag="24"></TEXTAREA><% '업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<TEXTAREA CLASS="hidden" NAME="txtSpread_i" tag="24"></TEXTAREA><% '업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>

<INPUT TYPE=HIDDEN NAME="htxtAcqNo"    tag="24">
<INPUT TYPE=HIDDEN NAME="txtMode"      tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows_m" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows_i" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"   tag="24">
<INPUT TYPE=HIDDEN NAME="hORGCHANGEID"   tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
