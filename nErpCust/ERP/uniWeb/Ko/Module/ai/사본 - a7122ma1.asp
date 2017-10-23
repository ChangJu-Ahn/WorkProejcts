<%@ LANGUAGE="VBSCRIPT"%>
<%Response.Expires = -1%>
<!--'**********************************************************************************************
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID           : a7122ma1
'*  4. Program Name         : 기초자산취득내역등록 
'*  5. Program Desc         : 고정자산별 취득내역을 등록,수정,삭제,조회 
'*  6. Comproxy List        : +As0021
'                             +As0029
'                             +B1a028ListMinorCode
'*  7. Modified date(First) : 2002/01/29
'*  8. Modified date(Last)  : 2002/01/29
'*  9. Modifier (First)     : 김호영 
'* 10. Modifier (Last)      : 김호영 
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

<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js">   </SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit                                                             '☜: indicates that All variables must be declared in advance
<!-- #Include file="../../inc/lgvariables.inc" -->	
'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Const BIZ_PGM_QRY_ID  = "a7122mb1.asp"												'☆: Head Query 비지니스 로직 ASP명 
'Const Biz_PGM_QRY_ID2 = "a7122mb4.asp"
Const BIZ_PGM_QRY_ID2  = "a7122mb1.asp"												'☆: Head Query 비지니스 로직 ASP명 

Const BIZ_PGM_DEL_ID  = "a7122mb3.asp"												'☆: Delete 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID = "a7122mb2.asp"												'☆: Save 비지니스 로직 ASP명 

Const gIsShowLocal = "Y"
<%
Const gIsShowLocal = "Y"
%>

Const COM_BIZ_EXCHRATE_ID = "../../inc/GetExchRate.asp"								'☆: 환율정보 비지니스 로직 ASP명 

Const TAB1 = 1																		'☜: Tab의 위치 


'''자산master
Dim C_AcqDt
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
Dim C_DeprLocAmt
Dim C_InvQty
Dim C_ResAmt
Dim C_DeprFrDt
Dim C_DurYrs
Dim C_DeprstsCd
Dim C_DeprstsPop
Dim C_Deprsts
Dim C_RefNo
Dim C_Desc

Const C_SHEETMAXROWS_m = 30


'========================================================================================================= 
'DIM lgBlnFlgChgValue           ' Variable is for Dirty flag
'Dim lgIntGrpCount              ' Group View Size를 조사할 변수 
'Dim lgIntFlgMode               ' Variable is for Operation Status

Dim lgStrPrevKey_i,lgStrPrevKey_m
'Dim lgLngCurRows

'========================================================================================================= 

Dim ihGridCnt                     'hidden Grid Row Count
Dim intItemCnt                    'hidden Grid Row Count
Dim lgstrConffg  
Dim dblXchRate		              'Exchange Rate 를 가지고 오기 
Dim IsOpenPop						' Popup
Dim gSelframeFlg

Dim lgFirstDeprYYYYMM
Function SetYYYYMMDt
	Dim intRetCD1
	Dim strYear, strMonth, strDay
	IntRetCD1 = CommonQueryRs("Top 1 FIRST_DEPR_YYYYMM","B_COMPANY","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If intRetCD1 <> False Then
		lgFirstDeprYYYYMM = Replace(lgF0,Chr(11),"")
	End If

	Call ExtractDateFrom(lgFirstDeprYYYYMM, "YYYYMM", "", strYear, strMonth, strDay)
	lgFirstDeprYYYYMM = UniConvYYYYMMDDToDate(parent.gAPDateFormat, strYear, strMonth, "01")

End Function
'========================================================================================================= 
Sub initSpreadPosVariables()
	 C_AcqDt		= 1
	 C_Deptcd		= 2
	 C_DeptPop		= 3
	 C_DeptNm		= 4
	 C_AcctCd		= 5
	 C_AcctPop		= 6
	 C_AcctNm		= 7
	 C_AsstNo		= 8
	 C_AsstNm		= 9
	 C_AcqAmt		= 10
	 C_AcqLocAmt	= 11
	 C_DeprLocAmt	= 12
	 C_InvQty		= 13
	 C_ResAmt		= 14
	 C_DeprFrDt	    = 15
	 C_DurYrs		= 16
	 C_DeprstsCd	= 17
	 C_DeprstsPop	= 18
	 C_Deprsts		= 19
	 C_RefNo		= 20
	 C_Desc		    = 21
End Sub

'========================================================================================
' Function Name : InitVariables
' Function Desc : This method initializes general variables
'========================================================================================

Sub InitVariables()
    
    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
'    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey_i = ""                           'initializes Previous Key
    lgStrPrevKey_m = ""
    lgLngCurRows = 0                            'initializes Deleted Rows Count
	lgBlnFlgChgValue = False                    'Indicates that no value changed	
End Sub

'========================================================================================================= 
Sub SetDefaultVal()
    Dim svrDate
    
	frm1.txtGLDt.text     = UniConvDateAToB("<%=GetSvrDate%>",parent.gServerDateFormat,gDateFormat)
	frm1.hOrgChangeId.value = parent.gChangeOrgId
	frm1.cboAcqFg.value		= "03"	
	frm1.txtDocCur.value	= parent.gCurrency
	frm1.txtXchRate.value	= 1

	
	lgBlnFlgChgValue = False
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
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

Sub InitSpreadSheet()
	
	Call InitSpreadPosVariables()
        
    Dim sList
    Dim strMaskYM
	
	strMaskYM = parent.gDateFormatYYYYMM
	
	strMaskYM = Replace(strMaskYM,"YYYY"      ,"9999")
	strMaskYM = Replace(strMaskYM,"YY"        ,"99")
	strMaskYM = Replace(strMaskYM,"MM"        ,"99")
	strMaskYM = Replace(strMaskYM,parent.gComDateType,"X")
    
    With frm1   
        
    ''''자산master
    ggoSpread.Source = .vspdData
    ggoSpread.Spreadinit "V20021104",,parent.gAllowDragDropSpread  
    .vspdData.ReDraw = False 
    
    .vspdData.MaxCols = C_Desc +1
   
	.vspdData.Col = .vspdData.MaxCols
	.vspdData.ColHidden = True

    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData
	
	Call GetSpreadColumnPos("A")

	ggoSpread.SSSetDate		C_AcqDt,  "취득일자",   10,2,  parent.gDateFormat  
	ggoSpread.SSSetEdit		C_DeptCd,  "부서코드",   10, , , 10, 2
	ggoSpread.SSSetButton	C_DeptPop
	ggoSpread.SSSetEdit		C_DeptNm,  "부서명",   20

	ggoSpread.SSSetEdit		C_AcctCd,  "계정코드", 10, , , 18, 2
	ggoSpread.SSSetButton	C_AcctPop
	ggoSpread.SSSetEdit		C_AcctNm,  "계정명",   30
	ggoSpread.SSSetEdit		C_AsstNo, "자산번호", 30
    ggoSpread.SSSetEdit		C_AsstNm, "자산명",   30
    
	ggoSpread.SSSetFloat    C_AcqAmt,   "취득금액",      15, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
	ggoSpread.SSSetFloat    C_AcqLocAmt,"취득금액(자국)",15, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
	ggoSpread.SSSetFloat    C_DeprLocAmt,"상각누계",15, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

	If gIsShowLocal = "N" Then
		Call ggoSpread.SSSetColHidden(C_AcqLocAmt,C_AcqLocAmt,True)
	End If

	Call AppendNumberPlace("6","11","0")
    ggoSpread.SSSetFloat    C_InvQty,    "재고수량",      15,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''<<<<<<<<<<<<<<<<<<<<<<<<<<<<
	ggoSpread.SSSetFloat    C_ResAmt,"잔존가액(자국)",15, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
    ggoSpread.SSSetMask		C_DeprFrDt,  "감가상각시작년월",   15,2, strMaskYM    
    ggoSpread.SSSetFloat    C_DurYrs,    "내용년수",      15,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z","0","60"
	ggoSpread.SSSetEdit		C_DeprstsCd,  "상각상태코드", 12, , , 2, 2
	ggoSpread.SSSetButton	C_DeprstsPop
	ggoSpread.SSSetEdit		C_Deprsts,  "상각상태명",   15
	ggoSpread.SSSetEdit		C_RefNo, "참조번호", 20
	ggoSpread.SSSetEdit		C_Desc,  "적요",     30
	
	Call ggoSpread.MakePairsColumn(C_DeptCd,C_DeptPop,"1")
	Call ggoSpread.MakePairsColumn(C_AcctCd,C_AcctPop,"1")
	Call ggoSpread.MakePairsColumn(C_DeprstsCd,C_DeprstsPop,"1")

    .vspdData.ReDraw = True
    End With

End Sub


'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadLock(byVal gird_fg, byVal lock_fg, byVal iRow)
    With frm1
		ggoSpread.Source = .vspdData
		.vspdData.ReDraw = False
		if Lock_fg = "insert" then		
			ggoSpread.SpreadLock C_DeptNm, iRow, C_DeptNm
			ggoSpread.SpreadLock C_AcctNm, iRow, C_AcctNm
			ggoSpread.SpreadLock C_DeprFrDt, iRow, C_DeprFrDt
		else
			ggoSpread.SpreadLock C_DeptNm,   iRow, C_DeptNm
			ggoSpread.SpreadLock C_AcctNm,   iRow, C_AcctNm
			ggoSpread.SpreadLock C_AsstNo,   iRow, C_DeptNm
			ggoSpread.SpreadLock C_DeprFrDt, iRow, C_DeprFrDt
		end if
		.vspdData.ReDraw = True
	End With    
End Sub


'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor_Master(ByVal pvStarRow, ByVal pvEndRow, ByVal lock_fg)
    
    With frm1.vspdData     
      
    ggoSpread.Source = frm1.vspdData
    
    .Redraw = False    	

    ggoSpread.SSSetRequired  C_AcqDt, pvStarRow, pvEndRow
    ggoSpread.SSSetRequired  C_Deptcd, pvStarRow, pvEndRow
    ggoSpread.SSSetProtected C_DeptNm, pvStarRow, pvEndRow
    
    ggoSpread.SSSetRequired  C_AcctCd, pvStarRow, pvEndRow
    ggoSpread.SSSetProtected C_AcctNm, pvStarRow, pvEndRow
    
'    ggoSpread.SSSetRequired  C_AsstNm,pvStarRow, pvEndRow
    
    ggoSpread.SSSetRequired  C_AcqAmt, pvStarRow, pvEndRow
    
    ggoSpread.SSSetRequired  C_DeprLocAmt, pvStarRow, pvEndRow  
    ggoSpread.SSSetRequired  C_InvQty, pvStarRow, pvEndRow
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''<<<<<<<<<<<<
    ggoSpread.SSSetRequired  C_ResAmt, pvStarRow, pvEndRow
    ggoSpread.SSSetRequired  C_DurYrs, pvStarRow, pvEndRow
    ggoSpread.SSSetProtected  C_DeprFrDt, pvStarRow, pvEndRow	
    ggoSpread.SSSetRequired  C_DeprStsCd, pvStarRow, pvEndRow
    ggoSpread.SSSetProtected C_DeprSts, pvStarRow, pvEndRow

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
		C_AcqDt		    = iCurColumnPos(1)
		C_Deptcd		= iCurColumnPos(2)
		C_DeptPop		= iCurColumnPos(3)
		C_DeptNm		= iCurColumnPos(4)
		C_AcctCd		= iCurColumnPos(5)
		C_AcctPop		= iCurColumnPos(6)
		C_AcctNm		= iCurColumnPos(7)
		C_AsstNo		= iCurColumnPos(8)
		C_AsstNm		= iCurColumnPos(9)
		C_AcqAmt		= iCurColumnPos(10)
		C_AcqLocAmt	    = iCurColumnPos(11)
		C_DeprLocAmt	= iCurColumnPos(12)
		C_InvQty		= iCurColumnPos(13)
		C_ResAmt		= iCurColumnPos(14)
		C_DeprFrDt	    = iCurColumnPos(15)
		C_DurYrs		= iCurColumnPos(16)
		C_DeprstsCd	    = iCurColumnPos(17)
		C_DeprstsPop    = iCurColumnPos(18)
		C_Deprsts		= iCurColumnPos(19)
		C_RefNo		    = iCurColumnPos(20)
		C_Desc		    = iCurColumnPos(21)
	End Select
End Sub

 '==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox_acqfg()

	'01신규취득	02무상기부 03기초자산	

	'------ Developer Coding part (Start ) --------------------------------------------------------------
	Dim IntRetCD1
	Dim intMaxRow, intLoopCnt
	Dim ArrTmpF0, ArrTmpF1
	
	On error resume next

	IntRetCD1= CommonQueryRs("MINOR_CD,MINOR_NM","B_MINOR","(MAJOR_CD = " & FilterVar("A2005", "''", "S") & " )",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	ArrTmpF0 = split(lgF0,chr(11))
	ArrTmpF1 = split(lgF1,chr(11))
	
	intMaxRow = ubound(ArrTmpF0)
	
	If intRetCD1 <> False Then
		for intLoopCnt = 0 to intMaxRow - 1
			Call SetCombo(frm1.cboAcqFg, ArrTmpF0(intLoopCnt), ArrTmpF1(intLoopCnt))
		next
	End If
	'------ Developer Coding part (End )   --------------------------------------------------------------
	
End Sub


 '==========================================  2.2.6 InitComboBox_rcpt()  =======================================
'	Name : InitComboBox_Depr()
'	Description : Combo Display
'========================================================================================================= 

Sub InitComboBox_Depr()

	'01상각진행중 02상각완료 03비상각	

	'------ Developer Coding part (Start ) --------------------------------------------------------------
	Dim IntRetCD1
		
	On error resume next

	IntRetCD1 = CommonQueryRs("MINOR_NM, MINOR_CD","B_MINOR","(MAJOR_CD = " & FilterVar("A2004", "''", "S") & " )  ORDER BY MINOR_CD",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		
	If intRetCD1 <> False Then
		ggoSpread.Source = frm1.vspddata
		ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_DeprstsCd
		ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_Deprsts
	End If
	'------ Developer Coding part (End )   --------------------------------------------------------------
	
end sub

Sub PopSaveSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.SaveSpreadColumnInf()
End Sub

Sub  PopRestoreSpreadColumnInf()	
	Dim varData
	Dim iRow
	
	ggoSpread.Source = frm1.vspdData
	Call ggoSpread.RestoreSpreadInf()
	Call InitSpreadSheet()
	Call ggoSpread.ReOrderingSpreadData()
	Call SetSpreadColor_Master(-1,-1,"query")

	With frm1

		.vspdData.Redraw = False

		For iRow = 0 To frm1.vspdData.MaxRows
			.vspdData.Col = C_DeprstsCd
			.vspdData.Row = iRow
			varData = frm1.vspdData.text
			
			Call subVspdSettingChange(iRow,varData)   ''''Rcpt Type별 입력필수 필드 표시 
		Next

		.vspdData.Redraw = True
	End With	

End Sub

 '==========================================  2.3.1 Tab Click 처리  =================================================
'	기능: Tab Click
'	설명: Tab Click시 필요한 기능을 수행한다.
'=================================================================================================================== 
 '----------------  ClickTab1(): Header Tab처리 부분 (Header Tab이 있는 경우만 사용)  ---------------------------- 
Function ClickTab1()
End Function

Function ClickTab2()
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


Sub txtGlDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtGlDt.Action = 7
        Call txtGlDt_onBlur()
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


	iCalledAspName = AskPRAspName("a7102ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a7102ra1" ,"x")
		IsOpenPop = False
		Exit Function
	End If
	IsOpenPop = True

	arrRet = window.showModalDialog(iCalledAspName & "?PID=" & gStrRequestMenuID, Array(window.parent,arrParam), _
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
' Function Name : OpenAcctCd
' Function Desc : OpenAcctCd Reference Popup
'===========================================================================
Function OpenDeprstsCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim field_fg

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "상각상태 팝업"			' 팝업 명칭 
	arrParam(1) = "B_MINOR"					' TABLE 명칭 
	arrParam(2) = Trim(frm1.vspdData.text)	        ' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = "(MAJOR_CD = " & FilterVar("A2004", "''", "S") & " ) And MINOR_CD <> " & FilterVar("03", "''", "S") & "  "							' Where Condition
	arrParam(5) = "상각상태코드"		    	' 조건필드의 라벨 명칭 

    arrField(0) = "MINOR_CD"						' Field명(0)
	arrField(1) = "MINOR_NM"						' Field명(1)

    arrHeader(0) = "상각상태코드"				' Header명(0)
	arrHeader(1) = "상각상태명"				' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")	

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		field_fg = "DeprstsCd"
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
    dim field_fg   	
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	if 	frm1.txtAcqDt.Text="" then '2006.10
		call DisplayMsgBox("971012","X", "취득일자","X")
		frm1.txtAcqDt.focus()
		exit function
	end if

	iCalledAspName = AskPRAspName("DeptPopupDtA2")
    
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "DeptPopupDtA2", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = Trim(frm1.txtDeptCd.value)						'부서코드 
	arrParam(1) = frm1.txtAcqDt.Text					'날짜(Default:현재일)
	arrParam(2) = lgUsrIntCd							'부서권한(lgUsrIntCd)
	arrParam(3) = "F"
	IsOpenPop = True

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		field_fg = "DeptCd"
		Call SetReturnVal(arrRet,field_fg)
	End If
End Function
'===========================================================================
' Function Name : OpenDeptCd (called from multi grid)
' Function Desc : OpenDeptCode Reference Popup
'===========================================================================
'Function OpenDeptCd()
'	Dim arrRet
'	Dim arrParam(3)
'    dim field_fg   	
'	Dim iCalledAspName
'
'	If IsOpenPop = True Then Exit Function
'	
'	iCalledAspName = AskPRAspName("DeptPopupDtA2")
'
'	If Trim(iCalledAspName) = "" Then
'		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "DeptPopupDtA2", "X")
'		IsOpenPop = False
'		Exit Function
'	End If
'
'	arrParam(0) = Trim(frm1.vspdData.text)						'부서코드 
'	arrParam(1) = frm1.txtGLDt.Text					'날짜(Default:현재일)
'	arrParam(2) = lgUsrIntCd							'부서권한(lgUsrIntCd)
'	arrParam(3) = "F"
'	IsOpenPop = True
'
'	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
'			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
'		
'	IsOpenPop = False
'
'	If arrRet(0) = "" Then
'		Exit Function
'	Else
'		field_fg = "DeptCd_grid"
'		Call SetReturnVal(arrRet,field_fg)
'	End If	
'End Function


Function OpenDeptCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
    Dim field_fg 
	Dim IntRetCD

	If IsOpenPop = True Then Exit Function

	IF Trim(frm1.txtDeptCd.value) = "" Then 
		IsOpenPop = False
		IntRetCD = DisplayMsgBox("124600","X","X","X")            '⊙: Display Message(There is no changed data.)
		frm1.txtDeptCd.focus
		Exit Function
	End If
	
	if 	frm1.txtAcqDt.Text="" then '2006.10
		call DisplayMsgBox("971012","X", "취득일자","X")
		frm1.txtAcqDt.focus()
		exit function
	end if
	
	IsOpenPop = True

	arrParam(0) = "부서 팝업"	
	arrParam(1) = "  ( SELECT DEPT_CD ,DEPT_NM FROM B_ACCT_DEPT "
	arrParam(1) = 	arrParam(1) & " WHERE COST_CD IN ( "
	arrParam(1) = 	arrParam(1) & " SELECT COST_CD "
	arrParam(1) = 	arrParam(1) & " FROM B_COST_CENTER "
	arrParam(1) = 	arrParam(1) & " WHERE BIZ_AREA_CD=(select Distinct C.BIZ_AREA_CD from B_ACCT_DEPT A, B_COST_CENTER B,B_BIZ_AREA C "
	arrParam(1) = 	arrParam(1) & " WHERE A.DEPT_CD = B.DEPT_CD "
	arrParam(1) = 	arrParam(1) & " AND B.BIZ_AREA_CD = C.BIZ_AREA_CD "
	arrParam(1) = 	arrParam(1) & " AND A.DEPT_CD = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S") 
	arrParam(1) = 	arrParam(1) & "		)"
	arrParam(1) = 	arrParam(1) & " ) "
	arrParam(1) = 	arrParam(1) & " AND ORG_CHANGE_ID =(select distinct org_change_id "
	arrParam(1) = 	arrParam(1) & " from b_acct_dept where org_change_dt = ( select max(org_change_dt) "
	arrParam(1) = 	arrParam(1) & " from b_acct_dept where org_change_dt <=  " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtAcqDt.Text, gDateFormat,""), "''", "S") & "))) A "
	
	
	arrParam(2) = Trim(frm1.vspdData.text)
	arrParam(3) = "" 
	'arrParam(4) = "A.ORG_CHANGE_ID = '" & parent.gChangeOrgId & "'"			
	arrParam(5) = "부서코드"			
	
    arrField(0) = "A.DEPT_CD"	
    arrField(1) = "A.DEPT_Nm"
    
    arrHeader(0) = "부서코드"
    arrHeader(1) = "부서코드명"

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

'=======================================================================================================
'	Name : OpenNoteNo()
'	Description : Note No PopUp
'=======================================================================================================
Function OpenNoteNo(byVal strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim field_fg
	If IsOpenPop = True  Then Exit Function

	IsOpenPop = True

	arrParam(0) = "지급어음번호 팝업"	
	arrParam(1) = "F_NOTE A,B_BANK B,B_BIZ_PARTNER C"
	arrParam(2) = strCode
	arrParam(3) = ""

	arrParam(4) = "A.NOTE_STS = " & FilterVar("BG", "''", "S") & "  AND A.NOTE_FG = " & FilterVar("D3", "''", "S") & "  AND A.BP_CD = C.BP_CD AND A.BANK_CD = B.BANK_CD"
	arrParam(5) = "지급어음번호"

    arrField(0) = "A.NOTE_NO"
    arrField(1) = "F2" & parent.gColSep & "CONVERT(VARCHAR(15),A.NOTE_AMT)"
    arrField(2) = "C.BP_NM"
    arrField(3) = "DD" & parent.gColSep & "Convert(varchar(40),A.ISSUE_DT)"
    arrField(4) = "DD" & parent.gColSep & "Convert(varchar(40),A.DUE_DT)"
    arrField(5) = "B.BANK_NM"

    arrHeader(0) = "지급어음번호"
    arrHeader(1) = "어음금액"
	arrHeader(2) = "거래처"
	arrHeader(3) = "발행일"
	arrHeader(4) = "만기일"
	arrHeader(5) = "은행"

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
Function OpenBankAcct(byVal strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim field_fg

	If IsOpenPop = True  Then Exit Function

	IsOpenPop = True

	arrParam(0) = "예적금코드 팝업"	' 팝업 명칭 
	arrParam(1) = "B_BANK A, F_DPST B"			' TABLE 명칭 
	arrParam(2) = strCode						' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = "A.BANK_CD = B.BANK_CD "		' Where Condition
	arrParam(5) = "은행코드"				' 조건필드의 라벨 명칭 

	arrField(0) = "A.BANK_NM"					' Field명(1)
	arrField(1) = "B.BANK_ACCT_NO"				' Field명(2)

	arrHeader(0) = "은행명"						' Header명(1)
	arrHeader(1) = "예적금코드"

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

 '------------------------------------------  SetReturnVal()  --------------------------------------------------
'	Name : SetReturnVal()
'	Description : Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetReturnVal(byval arrRet,byval field_fg)
	With frm1	
		Select case field_fg
			case "DeptCd"
				.txtDeptCd.value        = Trim(arrRet(0))
				.txtDeptNm.value 		= arrRet(1)
				.txtGLDt.Text			= arrRet(3)	
				Call txtDeptCd_OnBlur()	
			case "BpCd"
				.txtBpCd.Value			= Trim(arrRet(0))
				.txtBpNm.Value			= arrRet(1)

			case "Currency"
				.txtDocCur.Value		= arrRet(0)
				call txtDocCur_onChange()

			case "DeptCd_grid"
				.vspdData.Col			= C_DeptCd
				.vspdData.Text			= Trim(arrRet(0))
				.vspdData.Col			= C_DeptNm
				.vspdData.Text			= arrRet(1)
				'.txtGLDt.Text			= arrRet(3)
				Call txtGridDeptCd_OnBlur()
				call vspdData_Change(C_DeptCd, frm1.vspddata.activerow)

			case "AcctCd"
				.vspdData.Col			= C_AcctCd
				.vspdData.Text			= Trim(arrRet(0))
				.vspdData.Col			= C_AcctNm
				.vspdData.Text			= arrRet(1)
				call vspdData_Change(C_AcctCd, frm1.vspddata.activerow)
			case "DeprstsCd"
				.vspdData.Col			= C_DeprstsCd
				.vspdData.Text			= Trim(arrRet(0))
				.vspdData.Col			= C_Deprsts
				.vspdData.Text			= arrRet(1)
				call vspdData_Change(C_Deprsts, frm1.vspddata.activerow)

		End select	

		lgBlnFlgChgValue = True

	End With

End Function


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

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

End Function


'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
	
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
    frm1.txtAcqAmt.AllowNull =false
    Call ggoOper.LockField(Document, "N")
    Call InitSpreadSheet
    Call SetSpreadLock("I","",-1)
    Call SetSpreadLock("M","insert",-1)
	Call InitComboBox_acqfg
    Call InitVariables
    Call SetDefaultVal
	Call SetYYYYMMDt
	Call SetToolBar("1110010000001111")   '1110110100101111     

	'gSelframeFlg = TAB1
	frm1.txtAcqNo.focus
	
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub


'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
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
	End If
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

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal Col1, ByVal Col2)		' 7) 컬럼 width 변경 이벤트 핸들러 
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(Col1,Col2)

End Sub

'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================

Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub


Sub vspdData_ScriptDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc : 
'==========================================================================================

Sub vspdData_Change(ByVal Col, ByVal Row)
	Dim tmpDrCrFG
	Dim IntRetCD
	Dim TempExchRate
	Dim TempAmt
   'Call CheckMinNumSpread(frm1.vspdData, Col, Row)  

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

    frm1.vspdData.Row = Row   

    Select Case Col
	    Case   C_DeptCd
			frm1.vspdData.Col = C_DeptCd
			Call DeptCd_underChange(frm1.vspdData.text)

	    Case   C_AcctCd
			frm1.vspdData.Col = C_AcctCd
			Call AcctCd_underChange(frm1.vspdData.text)

		Case   C_DeprstsCd
			frm1.vspdData.Col = C_DeprstsCd
			Call DeprstsCd_underChange(frm1.vspdData.text)
    End Select
End Sub
Sub vspdData_ScriptLeaveCell(Col, Row, NewCol, NewRow, Cancel)

    If frm1.vspdData.MaxRows = 0 Then							'no data일 경우 vspdData_LeaveCell no 실행 
       Exit Sub													'tab이동시에 잘못된 140318 message 방지 
    End If
    
    With frm1.vspdData
		'If Col <> NewCol And NewCol > 0 Then
		If  NewCol > 0 Then
			If Col = C_DeprFrDt Then
				.Row = Row
				.Col = Col
			
				If .Text <> "" Then
                    If CheckDateFormat(.Text, parent.gDateFormatYYYYMM) = False  Then
						Call DisplayMsgBox("140318","X","X","X")	'년월을 올바로 입력하세요.
						.Text = ""
					End If
				End If
			End If
		
		End If
		
    End With

End Sub


 '==========================================================================================
'   Sub Procedure Name : subVspdSettingChange
'   Sub Procedure Desc : 
'==========================================================================================

Sub subVspdSettingChange(ByVal lRow, Byval varData)	
	
	if varData = "03" then 
		ggoSpread.SSSetProtected C_DeprLocAmt,       lRow, lRow	
		ggoSpread.SSSetProtected C_ResAmt,       lRow, lRow	
		ggoSpread.SSSetProtected C_DurYrs,       lRow, lRow	
		ggoSpread.SSSetProtected C_DeprstsCd,       lRow, lRow	
		ggoSpread.SSSetProtected C_DeprstsPop,       lRow, lRow	

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
	varAmt = frm1.vspdData.Text

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

Sub vspdData_KeyPress(index , KeyAscii )
     lgBlnFlgChgValue = True                                                 '⊙: Indicates that value changed
End Sub


'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 
'==========================================================================================

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)
	Dim strTemp
	Dim intPos1

	With frm1.vspdData
		If Row > 0 then
			if  Col = C_AcctPop Then
				.Col = C_AcctCd
				.Row = Row
				Call OpenAcctCd()
			ElseIf Col = C_DeptPop Then
				.Col = C_DeptCd
				.Row = Row
				Call OpenDeptCd()
			ElseIf Col =  C_DeprstsPop Then
				.Col = C_DeprstsCd
				.Row = Row
				Call OpenDeprstsCd()
			end if
		End If
	Call SetActiveCell(frm1.vspdData,Col -1 ,.ActiveRow ,"M","X","X")
	End With

End Sub

Sub txtBpCd_onblur()
	if frm1.txtBpCd.value = "" then
		frm1.txtBpNm.value = ""
	end if				
End Sub

'==========================================================================================
'   Event Name : txtDeptCd_OnBlur
'   Event Desc : 
'==========================================================================================
Sub AcctCd_underChange(Byval strCode)

    Dim strSelect
	Dim strFrom
	Dim strWhere
    Dim IntRetCD 
	Dim arrVal1

    If Trim(frm1.txtGLDt.Text = "") Then
		Exit sub
    End If
    lgBlnFlgChgValue = True

	strSelect	=			 "  A.DEPR_PROC_FG,C.ACCT_NM "
	strFrom		=			 " A_ASSET_DEPR_METHOD A (NOLOCK),A_ASSET_ACCT B (NOLOCK), A_ACCT C (NOLOCK) "
	strWhere	=			 " A.DEPR_MTHD = B.DEPR_MTHD"
	strWhere	= strWhere & "  AND B.ACCT_CD = C.ACCT_CD"
	strWhere	= strWhere & " 	AND B.ACCT_CD =  " & FilterVar(LTrim(RTrim(strCode)), "''", "S")

	'msgbox "select " & strSelect & " From " & strFrom & " where " & strWhere 

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
		IntRetCD = DisplayMsgBox("110100","X","X","X")  
		frm1.vspdData.Col = C_AcctCd
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.text = ""
		frm1.vspdData.Col = C_AcctNm
		frm1.vspdData.Row = frm1.vspdData.ActiveRow	
		frm1.vspdData.text = ""
		Call C_AcctNm_changeMode(frm1.vspdData.ActiveRow,"")
	Else
		arrVal1 = Split(lgF2By2, Chr(11))
		frm1.vspdData.Col = C_AcctNm
		frm1.vspdData.Row = frm1.vspdData.ActiveRow	
		frm1.vspdData.text = arrVal1(2)
		Call C_AcctNm_changeMode(frm1.vspdData.ActiveRow,arrVal1(1))
		frm1.vspdData.focus
	End If
	
End Sub

Sub DeprstsCd_underChange(Byval strCode)

    Dim strSelect
	Dim strFrom
	Dim strWhere
    Dim IntRetCD 
	Dim arrVal1


    lgBlnFlgChgValue = True

	strSelect	=			 "  A.MINOR_CD,A.MINOR_NM "
	strFrom		=			 " B_MINOR A (NOLOCK) "
	strWhere	=			 " MAJOR_CD = " & FilterVar("A2004", "''", "S") & " "
	strWhere	= strWhere & " 	AND A.MINOR_CD <> " & FilterVar("03", "''", "S") & "  " '비상각제외 
	strWhere	= strWhere & " 	AND A.MINOR_CD =  " & FilterVar(LTrim(RTrim(strCode)), "''", "S")
	'msgbox "select " & strSelect & " From " & strFrom & " where " & strWhere 

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
		IntRetCD = DisplayMsgBox("117423","X","X","X")  
		frm1.vspdData.Col = C_DeprstsCd
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.text = ""
		frm1.vspdData.Col = C_Deprsts
		frm1.vspdData.Row = frm1.vspdData.ActiveRow	
		frm1.vspdData.text = ""
		frm1.vspdData.focus
	Else
		arrVal1 = Split(lgF2By2, Chr(11))
		frm1.vspdData.Col = C_Deprsts
		frm1.vspdData.Row = frm1.vspdData.ActiveRow	
		frm1.vspdData.text = arrVal1(2)
		frm1.vspdData.focus
	End If
	
End Sub

Function C_AcctNm_changeMode(ByVal lRow,Byval strwhere)
	Dim IntRetCD1
	Dim ArrTmpF0, ArrTmpF1

	With frm1.vspdData     
		.Redraw = False    	
			Select Case Trim(strwhere)
			Case  "N"  
				.Row  = lRow
				.Col  = C_DurYrs
				.Text = "0"
				.Row  = lRow
				.Col  = C_DeprLocAmt
				.Text = "0"
				.Row  = lRow
				.Col  = C_ResAmt
				.Text = "0"

				.Row  = lRow
				IntRetCD1 = CommonQueryRs("Top 1 MINOR_CD,MINOR_NM","B_MINOR","(MAJOR_CD = " & FilterVar("A2004", "''", "S") & "  and MINOR_CD = " & FilterVar("03", "''", "S") & " )",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				ArrTmpF0 = split(lgF0,chr(11))
				ArrTmpF1 = split(lgF1,chr(11))

				.Col  = C_DeprstsCd
				.Text = ArrTmpF0(0)
				.Col  = C_Deprsts
				.Text = ArrTmpF1(0)
				ggoSpread.SSSetProtected C_DurYrs,       lRow, lRow	
				ggoSpread.SSSetProtected C_DeprLocAmt,       lRow, lRow	
				ggoSpread.SSSetProtected C_ResAmt,       lRow, lRow	
				ggoSpread.SSSetProtected C_DeprstsCd,       lRow, lRow	
				ggoSpread.SSSetProtected C_DeprstsPop,       lRow, lRow	

			Case Else
				.Row  = lRow
				.Col  = C_DurYrs
				.Lock = False

				.Row  = lRow
				.Col  = C_DeprLocAmt
				.Lock = False

				.Row  = lRow
				.Col  = C_ResAmt
				.Lock = False

				.Row  = lRow
				.Col  = C_Deprsts
				.Value = ""

				.Col  = C_DeprstsCd
				.Lock = False
				.Value = ""

				ggoSpread.SSSetRequired C_DurYrs,       lRow, lRow	
				ggoSpread.SSSetRequired C_DeprLocAmt,       lRow, lRow
				ggoSpread.SSSetRequired C_ResAmt,       lRow, lRow
				ggoSpread.SSSetRequired C_DeprstsCd,       lRow, lRow	
				ggoSpread.SpreadUnLock C_DeprstsPop,       lRow,C_DeprstsPop, lRow	

			End Select
		.Redraw = True
	End With
End Function
'==========================================================================================
'   Event Name : txtDeptCd_OnBlur
'   Event Desc : 
'==========================================================================================
Sub DeptCd_underChange(Byval strCode)

    Dim strSelect
	Dim strFrom
	Dim strWhere
    Dim IntRetCD 

    If Trim(frm1.txtAcqDt.Text = "") Then
		Exit sub
    End If
    lgBlnFlgChgValue = True

	strSelect	=			 " dept_cd, org_change_id, internal_cd "
	strFrom		=			 " b_acct_dept(NOLOCK) "
	strWhere	=  " COST_CD IN ( "
	strWhere	= strWhere & " 	SELECT COST_CD "
	strWhere	= strWhere & " 	FROM B_COST_CENTER "
	strWhere	= strWhere & " 	WHERE BIZ_AREA_CD=(select Distinct C.BIZ_AREA_CD from B_ACCT_DEPT A, B_COST_CENTER B,B_BIZ_AREA C "
	strWhere	= strWhere & " 		WHERE A.DEPT_CD = B.DEPT_CD "
	strWhere	= strWhere & " 		AND B.BIZ_AREA_CD = C.BIZ_AREA_CD "
	strWhere	= strWhere & " 		AND A.DEPT_CD = " & FilterVar(LTrim(RTrim(strCode)), "''", "S")
	strWhere	= strWhere & " 		AND A.ORG_CHANGE_ID =(select distinct org_change_id "
	strWhere	= strWhere & " 			 from b_acct_dept where org_change_dt = ( select max(org_change_dt) "
	strWhere	= strWhere & " 				 from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtAcqDt.Text, parent.gDateFormat,""), "''", "S") & "))"
	strWhere	= strWhere & " 			) "
	strWhere	= strWhere & " 		) "
	strWhere	= strWhere & " 		AND ORG_CHANGE_ID =(select distinct org_change_id "
	strWhere	= strWhere & " 			 from b_acct_dept where org_change_dt = ( select max(org_change_dt) "
	strWhere	= strWhere & " 				 from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtAcqDt.Text, parent.gDateFormat,""), "''", "S") & ")) "

	'msgbox "select " & strSelect & " From " & strFrom & " where " & strWhere 

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
		IntRetCD = DisplayMsgBox("124600","X","X","X")  
		frm1.vspdData.Col = C_deptcd
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.text = ""
		frm1.vspdData.Col = C_deptnm
		frm1.vspdData.Row = frm1.vspdData.ActiveRow	
		frm1.vspdData.text = ""

	End If
End Sub

Sub txtDeptCd_OnBlur()
        
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii, jj
	Dim IntRows

	If Trim(frm1.txtAcqDt.Text = "") Then    
		Exit sub
    End If
    
	if frm1.txtDeptCd.value = "" then
		For IntRows = 1 To frm1.vspdData.MaxRows
			If lgIntFlgMode <> parent.OPMD_CMODE	Then
				ggoSpread.UpdateRow IntRows
			End If
			frm1.vspdData.Row = IntRows
			frm1.vspdData.Col = C_DeptCd
			frm1.vspdData.text = ""
			frm1.vspdData.Col = C_DeptNm
			frm1.vspdData.text = ""
		Next
		frm1.txtDeptNm.value = ""
	end if	
    lgBlnFlgChgValue = True

	If TRim(frm1.txtDeptCd.value) <>"" Then
		'----------------------------------------------------------------------------------------
		strSelect	=			 " dept_cd, org_change_id, internal_cd "
		strFrom		=			 " b_acct_dept(NOLOCK) "
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtAcqDt.Text, gDateFormat,""), "''", "S") & "))"			

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then

			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
				For IntRows = 1 To frm1.vspdData.MaxRows
					If lgIntFlgMode <> parent.OPMD_CMODE	Then
						ggoSpread.UpdateRow IntRows
					End If
					frm1.vspdData.Row = IntRows
					frm1.vspdData.Col = C_DeptCd
					frm1.vspdData.text = ""
					frm1.vspdData.Col = C_DeptNm
					frm1.vspdData.text = ""
				Next
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
'   Event Name : txtGridDeptCd_OnBlur
'   Event Desc : 
'==========================================================================================

Sub txtGridDeptCd_OnBlur()
        
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii, jj

	If Trim(frm1.txtAcqDt.Text = "") Then
		Exit sub
    End If

    frm1.vspdData.Col			= C_DeptCd
	frm1.vspdData.row			= frm1.vspdData.ActiveRow

	if frm1.vspdData.text = "" then
		frm1.vspdData.Col			= C_DeptNM
		frm1.vspdData.text = ""
	end if	
    lgBlnFlgChgValue = True

	frm1.vspdData.Col			= C_DeptCd

	If TRim(frm1.txtDeptCd.value) <>"" Then
		'----------------------------------------------------------------------------------------
		
		strSelect	=			 " dept_cd, org_change_id, internal_cd "
		strFrom		=			 " b_acct_dept(NOLOCK) "
		strWhere	=  " COST_CD IN ( "
		strWhere	= strWhere & " 	SELECT COST_CD "
		strWhere	= strWhere & " 	FROM B_COST_CENTER "
		strWhere	= strWhere & " 	WHERE BIZ_AREA_CD=(select Distinct C.BIZ_AREA_CD from B_ACCT_DEPT A, B_COST_CENTER B,B_BIZ_AREA C "
		strWhere	= strWhere & " 		WHERE A.DEPT_CD = B.DEPT_CD "
		strWhere	= strWhere & " 		AND B.BIZ_AREA_CD = C.BIZ_AREA_CD "
		strWhere	= strWhere & " 		AND A.DEPT_CD = " & FilterVar(LTrim(RTrim(frm1.vspdData.text)), "''", "S")
		strWhere	= strWhere & " 		AND A.ORG_CHANGE_ID =(select distinct org_change_id "
		strWhere	= strWhere & " 			 from b_acct_dept where org_change_dt = ( select max(org_change_dt) "
		strWhere	= strWhere & " 				 from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtAcqDt.Text, parent.gDateFormat,""), "''", "S") & "))"
		strWhere	= strWhere & " 			) "
		strWhere	= strWhere & " 		) "
		strWhere	= strWhere & " 		AND ORG_CHANGE_ID =(select distinct org_change_id "
		strWhere	= strWhere & " 			 from b_acct_dept where org_change_dt = ( select max(org_change_dt) "
		strWhere	= strWhere & " 				 from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtAcqDt.Text, parent.gDateFormat,""), "''", "S") & ")) "

	'msgbox "select " & strSelect & " From " & strFrom & " where " & strWhere 

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then

			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.vspdData.Col	= C_DeptCd	:			frm1.vspdData.text = ""
			frm1.vspdData.Col	= C_DeptNM	:			frm1.vspdData.text = ""
			frm1.hOrgChangeId.value = ""

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
	

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 
    Dim IntRetCD 
    Dim var_m
    
    FncQuery = False

    Err.Clear                                                               '☜: Protect system from crashing
    ggoSpread.Source = frm1.vspdData
    var_m = ggoSpread.SSCheckChange
    
    If lgBlnFlgChgValue = True or var_m = True    Then    
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
	ggospread.ClearSpreadData		'Buffer Clear

    Call InitVariables															'⊙: Initializes local global variables

    Call SetSpreadLock("I","",-1)
    Call SetSpreadLock("M","query",-1)
    

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
    Call ggoOper.ClearField(Document, "2")                                         '⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear

    Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field
    
        
    Call InitVariables                                                      '⊙: Initializes local global variables    
    
    Call SetSpreadLock("I","",-1)
    Call SetSpreadLock("M","insert",-1)
	
	Call ClickTab1		'sstData.Tab = 1
   
	call txtDocCur_OnChangeASP()             

    Call SetToolBar("1110010000001111")    

    Call SetDefaultVal     


    
    FncNew = True                                                           '⊙: Processing is OK
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncDelete() 

    Dim IntRetCD 

    FncDelete = False                                                      '⊙: Processing is NG
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
		IntRetCD = DisplayMsgBox("900002","X","X","X")  '☜ 바뀐부분 
        Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")  '☜ 바뀐부분 
    If IntRetCD = vbNo Then

        Exit Function
    End If

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
    Dim varDLDt
    
    FncSave = False                                                         '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    On Error Resume Next                                                    '☜: Protect system from crashing
    
    '-----------------------
    'Precheck area
    '-----------------------    
	if frm1.vspdData.MaxRows < 1 then
		IntRetCD = DisplayMsgBox("117294","X","X","X")  ''자산세부내역을 입력하십시오.
		Exit Function
	end if
		
    ggoSpread.Source = frm1.vspdData
    var_m = ggoSpread.SSCheckChange

    If lgBlnFlgChgValue = False  and var_m = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")  '☜ 바뀐부분 
        Exit Function
    End If
    
    '-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "2") then                                   '⊙: Check contents area
       Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then	
		Exit Function
    End if            

	varAcqDt   = UniConvDateToYYYYMMDD(frm1.txtAcqDt.Text, gDateFormat,"")
	varGLDt    = UniConvDateToYYYYMMDD(frm1.txtAcqDt.Text, gDateFormat,"")
	
	If CompareDateByFormat(frm1.txtAcqDt.text,frm1.txtAcqDt.text,frm1.txtAcqDt.Alt,frm1.txtAcqDt.Alt, _
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
    
		frm1.vspdData.ReDraw = False

		if frm1.vspdData.MaxRows < 1 then Exit Function
	
		ggoSpread.Source = frm1.vspdData	
		ggoSpread.CopyRow
		SetSpreadColor_Master frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow, "insert"
    
		frm1.vspdData.Col  = C_AsstNo
		frm1.vspdData.Text = ""
		frm1.vspdData.Col  = C_AsstNm
		frm1.vspdData.Text = ""
		frm1.vspdData.Row  = frm1.vspdData.ActiveRow
		frm1.vspdData.Col  = C_AcctCd

		Call AcctCd_underChange(frm1.vspdData.text)
		frm1.vspdData.ReDraw = True
	
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
	
	if frm1.vspdData.MaxRows < 1 then	 Exit Function
	ggoSpread.Source = frm1.vspdData
		
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
	Dim imRow,iRow
	if IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
	else
		imRow = AskSpdSheetAddRowcount()

	If ImRow="" then
		Exit Function
	End If
	End If

	On Error Resume Next

		with frm1
			varMaxRow = .vspdData.MaxRows 

			.vspdData.focus

			ggoSpread.Source = .vspdData
			.vspdData.ReDraw = False

			ggoSpread.InsertRow ,imRow

			SetSpreadColor_Master .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1,"insert"
			For iRow = .vspdData.ActiveRow to .vspdData.ActiveRow + imRow - 1
				.vspdData.Row = iRow
				.vspdData.Col  = C_DeprFrDt
				.vspdData.Text = UNIMonthClientFormat(lgFirstDeprYYYYMM)
			Next

			.vspdData.ReDraw = True

		end with
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If 
    Set gActiveElement = document.ActiveElement  
    If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetToolBar("1110110100101111")   
	End If

End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
    Dim lDelRows 
    Dim lTempRows 

		frm1.vspdData.focus
    	ggoSpread.Source = frm1.vspdData
		if frm1.vspdData.MaxRows < 1 then Exit Function

		lDelRows = ggoSpread.DeleteRow    

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

    DbDelete = False														'⊙: Processing is NG

    Dim strVal

    strVal = BIZ_PGM_DEL_ID & "?txtMode=" & parent.UID_M0003
    strVal = strVal & "&txtAcqNo=" & Trim(frm1.txtAcqNo.value)			'☜: 삭제 조건 데이타 
    strVal = strVal & "&cboAcqFg=" & Trim(frm1.cboAcqFg.value)				'☜: 삭제 조건 데이타 

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
	Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001							'☜: 
		strVal = strVal & "&txtAcqNo="        & Trim(frm1.txtAcqNo.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&lgStrPrevKey_i=" & lgStrPrevKey_i
		strVal = strVal & "&txtMaxRows_i="    & frm1.vspdData.MaxRows
	End If
    
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	
    DbQuery = True                                                          '⊙: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()
	Dim varData
	Dim iRow

    lgIntFlgMode = parent.OPMD_UMODE

    Call ggoOper.LockField(Document, "Q")

	With frm1
		Call SetSpreadColor_Master(-1,-1,"query")
		.vspdData.Redraw = False
		For iRow = 0 To frm1.vspdData.MaxRows
			.vspdData.Col = C_DeprstsCd
			.vspdData.Row = iRow
			varData = Trim(frm1.vspdData.text)
			Call subVspdSettingChange(iRow,varData)   ''''Rcpt Type별 입력필수 필드 표시 
		Next
		.vspdData.Redraw = True
	End With

	call txtDocCur_OnChangeASP()  
	Call txtDeptCd_OnBlur()
	Call txtGridDeptCd_OnBlur()
	lgBlnFlgChgValue = False
	Call SetToolBar("111111110011111")
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
	Dim IntRetCD
	Dim strYear,strMonth,strDay
	
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
			.Col = C_AcctCd													'0
			strDel = strDel & Trim(.Text) & parent.gColSep
			strDel = strDel & "D" & parent.gColSep & parent.gChangeOrgId & parent.gColSep		'12		D=Delete
		ElseIf .Text = ggoSpread.UpdateFlag Then
			.Col = C_AcctCd  
			strVal = strVal & Trim(.Text) & parent.gColSep							'0
			strVal = strVal & "U" & parent.gColSep & parent.gChangeOrgId & parent.gColSep		'12		U=Update			
		Else
			.Col = C_AcctCd  
			strVal = strVal & Trim(.Text) & parent.gColSep							'0
			strVal = strVal & "C" & parent.gColSep & parent.gChangeOrgId & parent.gColSep		'12		C=Create
		End If	

		.Col = 0
		
		Select Case  .Text 

			Case ggoSpread.DeleteFlag

				.Col = C_DeptCd													'3	A073_IG1_I3_dept_cd
				strDel = strDel & Trim(.Text) & parent.gColSep

				.Col = C_AsstNo
				strDel = strDel & Trim(.Text) &parent.gRowSep		'⊙: 마지막 데이타는 Row 분리기호를 넣는다 

				lGrpCnt = lGrpCnt + 1
            
            Case Else 

				.Col = C_DeptCd										'3	부서코드 
				strVal = strVal & Trim(.Text) & parent.gColSep

				.Col = C_AsstNo										'4	자산번호 
				strVal = strVal & Trim(.Text) & parent.gColSep		
				
				.Col = C_AsstNm										'5	자산명 
				strVal = strVal & Trim(.Text) & parent.gColSep		

				.Col = C_AcqAmt										'6	취득금액 
				strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep

				.Col = C_AcqLocAmt									'7	취득금액(자국)
				strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep

				.Col = C_InvQty										'8	재고수량 
				strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep

If Trim(.Text) <> 0 then											' 재고수량이 0 이면 ERROR 체크 
Else 
   		IntRetCD = DisplayMsgBox("117215","X","X","X")  
   		Call SetToolBar("1111111100011111")
   		Call LayerShowHide(0)	
   		Exit Function
End If 

				.Col = C_ResAmt										'9	잔존가액(자국) 
				strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep

				.Col = C_RefNo										'10	참조번호	A073_IG1_I4_ref_no
				strVal = strVal & Trim(.Text) & parent.gColSep

				.Col = C_Desc										'11	적요	A073_IG1_I4_asset_desc
				strVal = strVal & Trim(.Text) & parent.gColSep	

				.Col = C_AcqDt									'							'12	A073_IG1_I4_reg_dt
				strVal = strVal & Trim(UniConvDateToYYYYMMDD(.Text,Parent.gDateFormat,"-")) & parent.gColSep
				strVal = strVal & parent.gColSep							'13	A073_IG1_I4_spec
				strVal = strVal & parent.gColSep							'14	A073_IG1_I4_doc_cur
				strVal = strVal & parent.gColSep							'15	A073_IG1_I4_xch_rate
				strVal = strVal & parent.gColSep							'16	A073_IG1_I4_inv_qty
				strVal = strVal & parent.gColSep							'17	A073_IG1_I4_tax_dur_yrs

				.Col = C_DurYrs										'18	내용년수	A073_IG1_I4_cas_dur_yrs
				strVal = strVal & Trim(.Text) & parent.gColSep

				strVal = strVal & parent.gColSep							'19	A073_IG1_I4_tax_end_l_term_cpt_tot_amt
				strVal = strVal & parent.gColSep							'20	A073_IG1_I4_cas_end_l_term_cpt_tot_amt
				strVal = strVal & parent.gColSep							'21	A073_IG1_I4_tax_end_l_term_depr_tot_amt

				.Col = C_DeprLocAmt									'22	상각누계	A073_IG1_I4_cas_end_l_term_depr_tot_amt
				strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep

				strVal = strVal & parent.gColSep							'23	A073_IG1_I4_tax_end_l_term_bal_amt
				strVal = strVal & parent.gColSep							'24	A073_IG1_I4_cas_end_l_term_bal_amt
				strVal = strVal & parent.gColSep							'25	A073_IG1_I4_tax_depr_sts

				.Col = C_DeprStsCd									'26	상각상태코드	A073_IG1_I4_cas_depr_sts
				strVal = strVal & Trim(.Text) & parent.gColSep

				strVal = strVal & parent.gColSep							'27	A073_IG1_I4_tax_depr_end_yyyymm
				strVal = strVal & parent.gColSep							'28	A073_IG1_I4_cas_depr_end_yyyymm
  
				.Col = C_DeprFrDt									'29	감가상각시작년월	A073_IG1_I4_start_depr_yymm
				Call ExtractDateFrom(.Text,parent.gDateFormatYYYYMM,parent.gComDateType,strYear,strMonth,strDay)
				strVal = strVal & strYear & strMonth & parent.gColSep	
				'strVal = strVal & replace(Trim(.Text),"-","") & parent.gColSep	

				strVal = strVal & parent.gColSep							'30	A073_IG1_I4_tax_dur_mnth
				strVal = strVal & parent.gColSep							'31	A073_IG1_I4_cas_dur_mnth


				strVal = strVal & parent.gRowSep

				lGrpCnt = lGrpCnt + 1

		End Select

    Next

	End With

	frm1.txtMaxRows_m.value  = lGrpCnt-1					'☜: Spread Sheet의 변경된 최대갯수 
	frm1.txtSpread_m.value = strDel & strVal				'☜: Spread Sheet 내용을 저장 

	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)				'☜: 저장 비지니스 ASP 를 가동 

    DbSave = True                                           ' ⊙: Processing is OK

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
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear

    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear

    Call SetSpreadLock("I","",-1)
    Call SetSpreadLock("M","insert",-1)


	frm1.txtAcqNo.value = iAcq_no

	call dbquery()
End Function


'========================================================================================
' Function Name : MaxSpreadVal
' Function Desc : 
'========================================================================================
Function MaxSpreadVal(byval Row)

end Function


'==========================================================================================
'   Event Name : txtGLDt_onBlur
'   Event Desc : 
'==========================================================================================
Sub txtGLDt_onBlur()
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii, jj
	Dim arrVal1, arrVal2
	Dim IntRows

     ggoSpread.Source = frm1.vspdData

	'If lgIntFlgMode = parent.OPMD_CMODE	Then
		lgBlnFlgChgValue = True
	With frm1
	
			If LTrim(RTrim(.txtDeptCd.value)) <> "" and Trim(.txtAcqDt.Text <> "") Then
				'----------------------------------------------------------------------------------------
					strSelect	=			 " Distinct org_change_id "    		
					strFrom		=			 " b_acct_dept(NOLOCK) "		
					strWhere	=			 " dept_Cd =  " & FilterVar(LTrim(RTrim(.txtDeptCd.value)), "''", "S") 
					strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
					strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
					strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(.txtAcqDt.Text, gDateFormat,""), "''", "S") & "))"			

				IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
					'msgbox "select " & strselect & " from " & strfrom & " where " & strwhere
				If IntRetCD = False  OR Trim(Replace(lgF0,Chr(11),"")) <> Trim(.hOrgChangeId.value) Then
						.txtDeptCd.value = ""
						.txtDeptNm.value = ""
						.hOrgChangeId.value = ""
						For IntRows = 1 To .vspdData.MaxRows
							If lgIntFlgMode <> parent.OPMD_CMODE	Then
								ggoSpread.UpdateRow IntRows
							End If
							.vspdData.Row = IntRows
							.vspdData.Col = C_DeptCd
							.vspdData.text = ""
							.vspdData.Col = C_DeptNm
							.vspdData.text = ""

						Next
						.txtDeptCd.focus
				End if
			End If
		End With
	'----------------------------------------------------------------------------------------
	'End If


End Sub

Sub txtAcqDt_Change()
    lgBlnFlgChgValue = true
End Sub

Sub txtVatAmt_Change()	'onblur
   lgBlnFlgChgValue = true
End Sub
Sub txtVatLocAmt_Change()	'onblur
   lgBlnFlgChgValue = true
End Sub
Sub txtXchRate_Change()	'onblur
   lgBlnFlgChgValue = true
End Sub

Sub cboAcqFg_Change()	'onblur
   lgBlnFlgChgValue = true
End Sub

'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX 
'====================================================================================================
Sub CurFormatNumericOCX()

	With frm1
		ggoOper.FormatFieldByObjectOfCur .txtAcqAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
	End With
End Sub

'===================================== CurFormatNumSprSheet()  ======================================
'	Name : CurFormatNumSprSheet()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet()

	With frm1
		ggoSpread.Source = frm1.vspdData

		ggoSpread.SSSetFloatByCellOfCur C_AcqAmt,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec		
	End With
End Sub  


'==========================================================================================
'   Event Name : txtDocCur_OnChange
'   Event Desc : 
'==========================================================================================
Sub txtDocCur_OnChange()
    lgBlnFlgChgValue = True '수정/변동된 내역이 있음을 Setting
    
	If UCase(frm1.txtDocCur.value) <> parent.gCurrency Then               ' 거래통화하고 Company 통화가 다를때 환율을 0으로 셋팅 
		frm1.txtXchRate.text	= 0                         ' 디폴트값인 1이 들어가 있으면 환율이 입력된 것으로 판단하여 
	Else 
		frm1.txtXchRate.text	= 1
	End If				                                    ' 환율정보를 읽지 않고 입력된 값으로 계산. 
    
    IF CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
		Call CurFormatNumericOCX()		
		Call CurFormatNumSprSheet()
	END IF	    
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>기초자산Master등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
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
						<TD <%=HEIGHT_TYPE_03%></TD>
					</TR>
					<TR HEIGHT=100%>
						<TD>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
								    <TD CLASS="TD5" NOWRAP>취득일자</TD>
									<TD CLASS="TD6" NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 name=txtAcqDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="22" ALT="취득일자"> </OBJECT>');</SCRIPT>
									</TD>
<%	If gIsShowLocal <> "N" Then	%>
									<TD CLASS="TD5" NOWRAP>전표일자</TD>
									<TD CLASS="TD6" NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime3 name=txtGLDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="22" ALT="전표생성일자"> </OBJECT>');</SCRIPT>
<%	ELSE %>
									<TD CLASS="TD5" NOWRAP>전표일자</TD>
									<TD CLASS="TD6" NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime3 name=txtGLDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="21" ALT="전표생성일자"> </OBJECT>');</SCRIPT>
<%	End If %>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>취득부서</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtDeptCd" SIZE=10 MAXLENGTH=10 tag="22XXXU" ALT="취득부서"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcqNo1" ALIGN=Top TYPE="BUTTON" ONCLICK="vbscript:OpenDept()">&nbsp;<INPUT TYPE=TEXT NAME="txtDeptNm" SIZE=22 tag="24"  alt = "부서명"></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>취득구분</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboAcqFg" STYLE="Width:150px;" tag="24" ALT="취득구분"><OPTION VALUE=""></OPTION></SELECT></TD>
									<TD CLASS=TD5 NOWRAP>거래처</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBpCd" ALT="거래처" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" ALIGN=Top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBp(frm1.txtBpCd.value,1)"> <INPUT NAME="txtBpNm" TYPE="Text" SIZE = 22 tag="24"></TD>
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
								</TR>
									
<%	ELSE %>
									<TD CLASS=TD5 NOWRAP>
									<TD CLASS=TD6 NOWRAP>
								
<INPUT TYPE=HIDDEN NAME="txtAcqLocAmt">
<%	End If %>					
								<TR>
									<TD CLASS="TD5" NOWRAP>회계전표번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtGLNo" SIZE=20 MAXLENGTH=18  tag="24" ALT="전표번호"></TD>
									<TD CLASS="TD5" NOWRAP>결의전표번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtTempGLNo" ALT="결의전표번호" TYPE="Text" MAXLENGTH=18 SIZE=25 tag="24" ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>적요</TD>
									<TD CLASS="TD656" NOWRAP COLSPAN=3><INPUT TYPE=TEXT NAME="txtDesc" SIZE=90 MAXLENGTH=128 tag="2X" ALT="적요"></TD>

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

<TEXTAREA CLASS="hidden" NAME="txtSpread_m" tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread_i" tag="24" TABINDEX="-1"></TEXTAREA>

<INPUT TYPE=HIDDEN NAME="htxtAcqNo"    tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMode"      tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows_m" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows_i" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"   tag="24" TABINDEX="-1">
<INPUT	TYPE=hidden	 NAME="hOrgChangeId"	tag="24" TABINDEX="-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
