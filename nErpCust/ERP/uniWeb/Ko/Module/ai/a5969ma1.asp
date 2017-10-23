<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : 월차 결산 
'*  2. Function Name        :
'*  3. Program ID           : A5969MA1
'*  4. Program Name         : A5969MA1
'*  5. Program Desc         : 보험정보 기초등록 
'*  6. Component List       :
'*  7. Modified date(First) : 2003/01/09
'*  8. Modified date(Last)  : 2003/05/30
'*  9. Modifier (First)     : leenamyo
'* 10. Modifier (Last)      : Jung Sung Ki
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>	
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">	</SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance

'========================================================================================================

Const BIZ_PGM_ID   = "a5969mb1.asp"
Const BIZ_PGM_ID1  = "a5969mb2.asp"
Const COOKIE_SPLIT =  4877	                                                        'Cookie Split String

'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->
'========================================================================================================
Dim IsOpenPop
Dim IsOpenPopDept
Dim UserPrevNext

'========================================================================================================
<%
	Dim StratDate	

	StratDate = GetSvrDate
%>
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE
	lgBlnFlgChgValue  = False
	lgIntGrpCount     = 0
    lgStrPrevKey      = ""
	lgStrPrevKeyIndex = 0 
    lgSortKey         = 1
End Sub

'========================================================================================================
Sub SetDefaultVal()
	Dim strYear,strMonth,strDay
	
	frm1.fpDateTime.Text = UniConvDateAToB("<%=StratDate%>",parent.gServerDateFormat,parent.gDateFormat)
	
    Call ggoOper.FormatDate(frm1.txtFromDt, parent.gDateFormat, 2)
    Call ggoOper.FormatDate(frm1.txtToDt, parent.gDateFormat, 2)    
    lgBlnFlgChgValue = False
End Sub

'========================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
End Sub

'========================================================================================================
' Name : MakeKeyStream
' Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
    Dim CNT_FROM_DT, CNT_TO_DT, FROM_DT, TO_DT, TEMP_GL_DT
    Dim strYear,strMonth,strDay
    Dim tstr

    Call ExtractDateFrom(frm1.txtCntFrom,parent.gDateFormat,parent.gComDateType,strYear,strMonth,strDay)
	CNT_FROM_DT = strYear & strMonth & strDay
	Call ExtractDateFrom(frm1.txtCntTo,parent.gDateFormat,parent.gComDateType,strYear,strMonth,strDay)
	CNT_TO_DT = strYear & strMonth & strDay
	Call ExtractDateFrom(frm1.txtFromDt,frm1.txtFromDt.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
	FROM_DT = strYear & strMonth
	Call ExtractDateFrom(frm1.txtToDt,frm1.txtToDt.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
	TO_DT = strYear & strMonth
	Call ExtractDateFrom(frm1.fpDateTime,parent.gDateFormat,parent.gComDateType,strYear,strMonth,strDay)
	TEMP_GL_DT = strYear & strMonth & strDay

	Select Case pOpt
		Case "Q"
			lgKeyStream = Frm1.txtInsuerCd.Value & parent.gColSep       'You Must append one character(parent.gColSep)
		Case "S"
            lgKeyStream = Trim(CNT_FROM_DT) & parent.gColSep       'You Must append one character(parent.gColSep)
            lgKeyStream = lgKeyStream & Trim(CNT_TO_DT)        & parent.gColSep 
            lgKeyStream = lgKeyStream & Trim(FROM_DT)        & parent.gColSep
            lgKeyStream = lgKeyStream & Trim(TO_DT)        & parent.gColSep
            lgKeyStream = lgKeyStream & Trim(TEMP_GL_DT)        & parent.gColSep
       Case "D" 
			lgKeyStream = Trim(frm1.txtInsuerCd1.value)        & parent.gColSep
	End Select
End Sub

'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX 
'====================================================================================================
Sub CurFormatNumericOCX()
	With frm1
	'해당되는 금액이 있는 Data 필드에 대하여 각각 처리 
		'환율 
		ggoOper.FormatFieldByObjectOfCur .txtExRate , .txtTradeCur.value, parent.ggExchRateNo  , gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		'계약금액 
		ggoOper.FormatFieldByObjectOfCur .txtCntAmt , .txtTradeCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		'보험료 
		ggoOper.FormatFieldByObjectOfCur .txtAmt    , .txtTradeCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
	End With
End Sub
'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 

	With frm1
		.txtExRate.text	        = 1
		.txtCntAmt.text	        = 0
		.txtLocCntAmt.text	    = 0
		.txtAmt.text		    = 0
		.txtLocAmt.text	        = 0
		.cboPrivatePublic.value = "N"
	End With
	
	frm1.txtInsuerCd.focus	
	lgBlnFlgChgValue = False
End Sub

'========================================================================================================
Sub Form_Load()
    Err.Clear

	Call LoadInfTB19029 
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")

	Call SetDefaultVal
	Call InitData()
	Call InitVariables
	Call SetToolbar("1111100000111111")                                                     '☆: Developer must customize
	'//전표요청이 있을시 주석 풀기 
	'Call chgBtnDisable(1)
End Sub

'========================================================================================================
Function FncQuery()
    Dim IntRetCD

    FncQuery = False
    Err.Clear

	If Not chkField(Document, "1") Then
		Exit Function
    End If

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X") '☜ "데이타가 변경되었습니다. 조회하시겠습니까?"
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If	

	Call ggoOper.ClearField(Document, "2")
    Call ggoOper.LockField(Document , "N")
    Call InitData()
	Call SetDefaultVal()
	'//전표요청이 있을시 주석 풀기 
    'Call chgBtnDisable(1)

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    Call InitVariables()
    Call MakeKeyStream("Q")

	'------ Developer Coding part (End )   --------------------------------------------------------------
    If DbQuery = False Then
		Exit Function
    End If

    Set gActiveElement = document.ActiveElement
    FncQuery = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
Function FncNew()
    Dim IntRetCD 

    FncNew = False
    Err.Clear

    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"x","x")
       If IntRetCD = vbNo Then
			Exit Function
       End If
    End If

    Call ggoOper.ClearField(Document, "1")
    Call ggoOper.ClearField(Document, "2")
    Call ggoOper.LockField(Document , "N")
	Call SetToolbar("1111100000101111")
    Call SetDefaultVal()
    Call InitData
    Call InitVariables
    call txtTradeCur_Onchange()
    '//전표요청이 있을시 주석 풀기 
    'Call chgBtnDisable(1)

    Set gActiveElement = document.ActiveElement
    FncNew = True
End Function
	
'========================================================================================================
Function FncDelete()
    Dim intRetCD

    FncDelete = False
    Err.Clear

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                            '☜: Please do Display first. 
        Call DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"x","x")
	If IntRetCD = vbNo Then
		Exit Function
	End If

    Call MakeKeyStream("D")

    If DbDelete = False Then
		Exit Function
    End If

    Set gActiveElement = document.ActiveElement
    FncDelete = True
End Function

'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    Dim pFromDt, pToDt, pDateTime
    Dim strYear, strMonth, strDay
    Dim FrDt
	Dim strSelect, strFrom, strWhere

    FncSave = False
    Err.Clear

    If lgBlnFlgChgValue = False Then 
        IntRetCD = DisplayMsgBox("900001","x","x","x")
        Exit Function
    End If

    Call deptCheck()

    If Not chkField(Document, "2") Then
		Exit Function
    End If

	If ValidDateCheck(frm1.txtCntFrom, frm1.txtCntTo)=False Then				 '☜ : 시장월은 종료월보다 작아야합니다.
	    frm1.txtCntFrom.Text = ""
        frm1.txtCntFrom.focus
        Exit Function
    End If 
	
	If CompareDateByFormat(frm1.txtFromDt.Text, frm1.txtToDt.Text,frm1.txtFromDt.Alt,frm1.txtToDt.Alt, "970024", frm1.txtFromDt.UserDefinedFormat, parent.gComDateType, True)=False Then  '☜ : 시장월은 종료월보다 작아야합니다. 
	    frm1.txtFromDt.Text = ""
        frm1.txtFromDt.focus
        Exit Function
    End If 

    Call ExtractDateFrom(frm1.txtFromDt,frm1.txtFromDt.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
	pFromDt = strYear & strMonth
	Call ExtractDateFrom(frm1.txtToDt,frm1.txtToDt.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
	pToDt = strYear & strMonth
	Call ExtractDateFrom(frm1.fpDateTime,frm1.fpDateTime.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
	pDateTime = strYear & strMonth	

    If UNICDbl(pFromDt) > UNICDbl(pDateTime) or UNICDbl(pToDt) < UNICDbl(pDateTime) Then		'☜ : 전표일자는 적용기간 사이에 있어야 합니다.(수정요망)
		Call DisplayMsgBox("am0030","X","전표일자","적용기간")
		frm1.fpDateTime.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	End If

	If UNICDbl(frm1.txtExRate.text) = 0 Then
		Call DisplayMsgBox("121500","X","X","X")                         '☜ : 숫자영 
		frm1.txtExRate.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End If

	If UNICDbl(frm1.txtCntAmt.text) = 0 Then
		Call DisplayMsgBox("189306","X","X","X")                         '☜ : 숫자영 
		frm1.txtCntAmt.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End If

	If UNICDbl(frm1.txtAmt.text) = 0 Then
		Call DisplayMsgBox("189306","X","X","X")                         '☜ : 숫자영 
		frm1.txtAmt.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End If

	If UNICDbl(frm1.txtCntAmt.text) < UNICDbl(frm1.txtAmt.text) Then
		Call DisplayMsgBox("970023","X",frm1.txtCntAmt.Alt,frm1.txtAmt.Alt)                         '☜ : 숫자영 
		Exit Function
	End If

	If UNICDbl(frm1.txtLocCntAmt.text) < UNICDbl(frm1.txtLocAmt.text) Then
		Call DisplayMsgBox("970023","X",frm1.txtLocCntAmt.Alt,frm1.txtLocAmt.Alt)                         '☜ : 숫자영 
		Exit Function
	End If

	If UNICDbl(frm1.txtLocCntAmt.text) = 0 Or Trim(frm1.txtLocCntAmt.text) = "" Then
		frm1.txtLocCntAmt.text = UNICDbl(parent.UNICDbl(frm1.txtExRate.text) * UNICDbl(frm1.txtCntAmt.text))
	End If

	If UNICDbl(frm1.txtLocAmt.text) = 0 Or Trim(frm1.txtLocAmt.text) = "" Then
		frm1.txtLocAmt.text = UNICDbl(parent.UNICDbl(frm1.txtExRate.text) * UNICDbl(frm1.txtAmt.text))
	End If

	FrDt = UniConvDateToYYYYMMDD(frm1.fpDateTime.Text,parent.gDateFormat,"")   '//parent.UNIConvDate(frm1.txtInDt.Text)
	strSelect = strSelect & "isnull(case t.loc_cur when " & FilterVar(UCase(frm1.txtTradeCur.value), "''", "S") & " Then 1 "
	strSelect = strSelect & " 	Else  Case t.xch_rate_fg " 
	strSelect = strSelect & " 		  When " & FilterVar("M", "''", "S") & "  Then ( SELECT isnull(STD_RATE,0) "
	strSelect = strSelect & "				FROM	b_monthly_exchange_rate (nolock) "
	strSelect = strSelect & "				WHERE	apprl_yrmnth 	= CONVERT (varchar(06), " & FilterVar(FrDt,"" & FilterVar("29991231", "''", "S") & " ", "S") & ", 112) "
	strSelect = strSelect & "				and	from_currency	= " & FilterVar(UCase(frm1.txtTradeCur.value), "''", "S") 
	strSelect = strSelect & "				and	to_currency	= t.loc_cur ) "
	strSelect = strSelect & "		  Else (	SELECT  isnull(STD_RATE,0) "
	strSelect = strSelect & "				FROM	b_daily_exchange_rate (nolock) "
	strSelect = strSelect & "				WHERE	apprl_dt		= " & FilterVar(FrDt,"" & FilterVar("29991231", "''", "S") & " ", "S")
	strSelect = strSelect & "				and	from_currency	= " & FilterVar(UCase(frm1.txtTradeCur.value), "''", "S") 
	strSelect = strSelect & "				and	to_currency	= t.loc_cur ) "
	strSelect = strSelect & "		  End  "
	strSelect = strSelect & "End,0) as xch_rate "
	strFrom  = " (SELECT isnull(XCH_RATE_FG,'') as xch_rate_fg, loc_cur  from  b_company) t "
	strWhere = ""
	IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If IntRetCD = false or Trim(Replace(lgF0,Chr(11),""))="0" or Trim(Replace(lgF0,Chr(11),"")) = ""  Then
		Call DisplayMsgBox("am0023","X","X","X")
		Exit Function
	End If	

    Call MakeKeyStream("S")

    If DbSave = False Then                                                       '☜: Query db data
		Exit Function
    End If
    
    Set gActiveElement = document.ActiveElement
    FncSave = True
End Function

'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False
    Err.Clear
	
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    Call ggoOper.ClearField(Document, "1")
    Call ggoOper.LockField(Document, "N")
    
   	lgIntFlgMode = parent.OPMD_CMODE
    frm1.txtInsuerCd.value = ""
	frm1.txtInsuerCd1.value = ""
	frm1.txtInsuerNm.value = ""
	frm1.txtInsuerNm1.value = ""

    Set gActiveElement = document.ActiveElement   
    FncCopy = True
End Function

'========================================================================================================
Function FncCancel() 
    FncCancel = False
    Err.Clear
    Set gActiveElement = document.ActiveElement
    FncCancel = True
End Function

'========================================================================================================
Function FncPrint()
	Parent.fncPrint()
End Function

'========================================================================================================
Function FncExcel() 
    FncExcel = False
    Err.Clear
	Call Parent.FncExport(parent.C_SINGLE)
    FncExcel = True
End Function

'========================================================================================================
Function FncFind() 
    FncFind = False
    Err.Clear
	Call Parent.FncFind(parent.C_SINGLE, True)
    FncFind = True
End Function

'========================================================================================================
Function FncExit()
	Dim IntRetCD

    FncExit = False
    Err.Clear

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
    FncExit = True
End Function

'========================================================================================================
Function DbQuery()
    Dim strVal

    Err.Clear																		'☜: Clear err status
    DbQuery = False																	'☜: Processing is NG
    
    If LayerShowHide(1) = False Then
	    Exit Function
	End If

    strVal = BIZ_PGM_ID & "?txtMode="          & parent.UID_M0001					'☜: Query
    strVal = strVal     & "&txtPrevNext="      & ""									'☜: Direction
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream						'☜: Query Key
    strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex
    strVal = strVal		& "&lgCurrency="	   & frm1.txtTradeCur.value

	Call RunMyBizASP(MyBizASP, strVal)												'☜:  Run biz logic
    DbQuery = True																	'☜: Processing is OK
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
Function DbSave()
	Dim lGrpcnt 
	Dim strVal, strDel
	Dim IntRows

    Err.Clear

	DbSave = False
	lGrpCnt =0

	If layerShowHide(1) = False Then
	    Exit Function
	End If

	With Frm1
		.txtMode.value        = parent.UID_M0002									'☜: Save
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream											'☜: Save Key
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
    DbSave  = True                                                                  '☜: Processing is NG
End Function

'========================================================================================================
Function DbDelete()
	Dim strVal
    Err.Clear

	DbDelete = False

	If LayerShowHide(1) = False Then
		Exit Function
	End If

    strVal = BIZ_PGM_ID & "?txtMode="          & parent.UID_M0003                   '☜: DELETE
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream 

	Call RunMyBizASP(MyBizASP, strVal)

	DbDelete = True
End Function

'========================================================================================================
Sub DbQueryOk()
	Dim iRow,intIndex
	Dim varData

	lgIntFlgMode      = parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode
	
	Call ggoOper.LockField(Document, "N")
	Call ggoOper.SetReqAttr(frm1.txtInsuerCd1, "Q")
	Call ggoOper.SetReqAttr(frm1.txtDept2, "N")	
	Call ggoOper.SetReqAttr(frm1.txtRefNo, "D")	
	Call SetToolbar("111110000011111")                                                     '☆: Developer must customize
	
	'//Button 조작 
	'//전표요청이 있을시 주석 풀기 
	'If trim(frm1.txtInsuerCd1.value) = "" and trim(frm1.txtTempGlNo.value) = "" and trim(frm1.txtGlNo.value) = "" Then
	'	Call chgBtnDisable(1)
	'ElseIF trim(frm1.txtInsuerCd1.value) <> "" and trim(frm1.txtTempGlNo.value) = "" and trim(frm1.txtGlNo.value) = "" Then
	'	Call chgBtnDisable(2)
	'ElseIF trim(frm1.txtInsuerCd1.value) <> "" and trim(frm1.txtTempGlNo.value) <> "" or trim(frm1.txtGlNo.value) <> "" Then
	'	Call chgBtnDisable(3)
	'End If	
	
    Set gActiveElement = document.ActiveElement  
    lgBlnFlgChgValue = False  
End Sub

'========================================================================================================
Sub DbSaveOk()
	On Error Resume Next
	Err.Clear 

	frm1.txtInsuerCd.value = frm1.txtInsuerCd1.value
    frm1.txtInsuerNm.value = frm1.txtInsuerNm1.value

    Set gActiveElement = document.ActiveElement

    Call MakeKeyStream("Q")
    Call DbQuery()
End Sub

'========================================================================================================
Sub DbDeleteOk()
	Call InitVariables()
	Call FncNew()	
End Sub

'------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function
	
	If UCase(frm1.txtCustomCd.className) = "PROTECTED" Then Exit Function
	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = ""									' 채권과 연계(거래처 유무)
	arrParam(2) = ""									'FrDt
	arrParam(3) = ""									'ToDt
	arrParam(4) = "T"									'B :매출 S: 매입 T: 전체 
	arrParam(5) = ""									'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 	
	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
        frm1.txtCustomCd.focus
		Exit Function
	Else
		Call SetConSItemDC(arrRet, iWhere)
		lgBlnFlgChgValue = True
	End If
End Function

'========================================================================================================
Function BtnPopUp(Inobj,iRequired)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If BtnPopupDisabled(Inobj) = False Then Exit Function

	Select Case iRequired
		Case 0
		    arrParam(1) = "A_INSURE"
			arrParam(2) = Trim(frm1.txtInsuerCd.value)
			arrParam(3) = ""
			arrParam(4) = "ISNULL(TEMP_GL_NO,'') = '' AND ISNULL(GL_NO, '') = ''"
			arrParam(5) = "보험코드"

		    arrField(0) = "INSURE_CD"
		    arrField(1) = "INSURE_NM"

		    arrHeader(0) = "보험코드"
		    arrHeader(1) = "보험명"
		Case 1
			arrParam(1) = "B_MINOR"
			arrParam(2) = Trim(frm1.txtInsuerTp.value)
			arrParam(3) = ""
			arrParam(4) = "MAJOR_CD=" & FilterVar("A1030", "''", "S") & " "
			arrParam(5) = "보험종류"

			arrField(0) = "MINOR_CD"
			arrField(1) = "MINOR_NM"

		    arrHeader(0) = "보험종류코드"
		    arrHeader(1) = "보험종류"
		Case 2
			arrParam(1) = "B_CURRENCY"
			arrParam(2) = Trim(frm1.txtTradeCur.value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "거래통화"

			arrField(0) = "CURRENCY"
			arrField(1) = "CURRENCY_DESC"

		    arrHeader(0) = "거래통화코드"
		    arrHeader(1) = "거래통화"
		Case 3
			arrParam(1) = "B_BIZ_PARTNER"
			arrParam(2) = Trim(frm1.txtCustomCd.Value)
			arrParam(3) = ""
			arrParam(4) = "BP_TYPE <= " & FilterVar("CS", "''", "S") & " "
			arrParam(5) = "거래처"

			arrField(0) = "BP_CD"
			arrField(1) = "BP_NM"

			arrHeader(0) = "거래처"
			arrHeader(1) = "거래처명"
		Case 6
			arrParam(1) = "B_ACCT_DEPT A, B_COST_CENTER B"
			arrParam(2) = Trim(frm1.txtDept1.value)
			arrParam(3) = ""
			arrParam(4) = "A.ORG_CHANGE_ID = (SELECT CUR_ORG_CHANGE_ID FROM B_COMPANY) AND A.COST_CD = B.COST_CD"
			arrParam(5) = "발의부서"

			arrField(0) = "A.DEPT_CD"
			arrField(1) = "A.DEPT_NM"
			arrField(2) = "B.BIZ_AREA_CD"
			arrField(3) = "A.ORG_CHANGE_ID"

		    arrHeader(0) = "발의부서코드"
		    arrHeader(1) = "발의부서"	
		    arrHeader(2) = "사업장코드"
		    arrHeader(3) = "조직개편코드"
		Case 7
			arrParam(1) = "B_ACCT_DEPT A, B_COST_CENTER B"
			arrParam(2) = Trim(frm1.txtDept2.value)
			arrParam(3) = ""
			arrParam(4) = "A.COST_CD = B.COST_CD AND A.ORG_CHANGE_ID = (SELECT CUR_ORG_CHANGE_ID from B_COMPANY) AND B.BIZ_AREA_CD =  " & FilterVar(frm1.txtCostCd.value, "''", "S") 
			arrParam(5) = "귀속부서"

			arrField(0) = "A.DEPT_CD"
			arrField(1) = "A.DEPT_NM"

			arrHeader(0) = "귀속부서코드"
			arrHeader(1) = "귀속부서"	
		Case 8
			arrParam(1) = "A_ACCT"
			arrParam(2) = frm1.txtInsureAcct.value
			arrParam(3) = ""
			arrParam(4) = "DEL_FG <> " & FilterVar("Y", "''", "S") & " "
			arrParam(5) = "보험료계정"

			arrField(0) = "ACCT_CD"
			arrField(1) = "ACCT_NM"

			arrHeader(0) = "보험료계정코드"
			arrHeader(1) = "보험료계정명"
	End Select

	arrParam(0) = arrParam(5)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		With frm1
			Select Case iRequired
				Case 0
					.txtInsuerCd.focus
				Case 1
					.txtInsuerTp.focus
				Case 2
					.txtTradeCur.focus
				Case 3
					.txtCustomCd.focus
				Case 6
					.txtDept1.focus
				Case 7
					.txtDept2.focus
				Case 8
					.txtInsureAcct.focus
			End Select
		End With
		Exit Function
	Else
		Call SetConSItemDC(arrRet, iRequired)
	End If
End Function

'========================================================================================================
Function OpenDept(Byval strCode, iWhere)
	Dim arrRet
	Dim arrParam(5)
	Dim iCalledAspName
	Dim IntRetCD
	Dim Inobj

	If IsOpenPop = True Then Exit Function

	If iWhere = "1" Then
		Set Inobj = frm1.txtDept1
	Else
		Set Inobj = frm1.txtDept2 
	End If

	If BtnPopupDisabled(Inobj) = False Then Exit Function
		
	iCalledAspName = AskPRAspName("DeptPopupDt3")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "DeptPopupDt3", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = strCode		            '  Code Condition
   	arrParam(1) = frm1.fpDateTime.Text
	arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  

	' T : protected F: 필수 
	If lgIntFlgMode = Parent.OPMD_UMODE then
		arrParam(3) = "T"									' 결의일자 상태 Condition  
	Else
		arrParam(3) = "F"									' 결의일자 상태 Condition  
	End If

	arrParam(4) = iWhere
	arrParam(5) = Trim(frm1.txtCostCd.value)

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Select Case iWhere
			Case 1
				frm1.txtDept1.focus
			Case 2
				frm1.txtDept2.focus
		End Select
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
			Case 1
				.txtDept1.value = arrRet(0)
				.txtDept1Nm.value = arrRet(1)
				.txtCostCd.value = arrRet(2)
				.txtOrgChId.value = arrRet(3)
				.txtInternalCd1.value = arrRet(4)
				.fpDateTime.text = arrRet(5)
				.txtDept2.value = ""
				.txtDept2Nm.value = ""
				Call txtDept1_OnChange()
			Case 2
				.txtDept2.focus
				.txtDept2.value = arrRet(0)
				.txtDept2Nm.value = arrRet(1)
				.txtInternalCd2.value = arrRet(4)
				.fpDateTime.text = arrRet(5)
				Call txtDept2_OnChange()  
		End Select
	End With
End Function

'========================================================================================================
'	Name : SetBizArea()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function SetConSItemDC(Byval arrRet, Byval iRequired)
	With frm1
		Select Case iRequired
			Case 0
				.txtInsuerCd.focus
				.txtInsuerCd.value = arrRet(0)
				.txtInsuerNm.value = arrRet(1)
			Case 1
				.txtInsuerTp.focus
				.txtInsuerTp.value = arrRet(0) 
				.txtInsuerTpNm.value = arrRet(1)
			Case 2
				.txtTradeCur.focus
				.txtTradeCur.value = arrRet(0) 
				.txtTradeCurNm.value = arrRet(1)
				Call FncRate("CURR")
				Call txtTradeCur_Onchange()
			Case 3
				.txtCustomCd.focus
				.txtCustomCd.value = arrRet(0)
				.txtCustomNm.value = arrRet(1)
'				Call FncCurrency
			Case 6
				.txtDept1.value = arrRet(0)
				.txtDept1Nm.value = arrRet(1)
				.txtCostCd.value = arrRet(2)
				.txtOrgChId.value = arrRet(3)
				.txtDept2.value = ""
				.txtDept2Nm.value = ""
				Call ggoOper.SetReqAttr(frm1.txtDept2, "N")
				frm1.txtDept2.focus
			Case 7
				.txtDept2.focus
				.txtDept2.value = arrRet(0)
				.txtDept2Nm.value = arrRet(1)
			Case 8
				.txtInsureAcct.focus
				.txtInsureAcct.value = arrRet(0)
				.txtInsureAcctNm.value = arrRet(1)
		End Select
	End With
	
	If iRequired <> 0 Then
		lgBlnFlgChgValue = True
	End If
End Function

'========================================================================================================
' Name : txtDept1_Onchange
' Desc : developer describe this line
'========================================================================================================
Function txtDept1_Onchange()
    Dim strSelect,strFrom,strWhere
    Dim IntRetCD
	Dim arrVal1,arrVal2
	Dim ii,jj

	With Frm1
		If .txtDept1.value = "" Then
			.txtDept1.value       = ""
			.txtDept1Nm.value     = ""
			.txtDept2.value       = ""
			.txtDept2Nm.value     = ""
			.txtInternalCd1.value = ""
			.txtInternalCd2.value = ""
			.txtCostCd.value      = ""
			.txtOrgChId.value     = ""
			Call ggoOper.SetReqAttr(.txtDept2, "Q")
			.txtDept1.focus
			Set gActiveElement    = document.activeElement
			lgBlnFlgChgValue      = True 
			Exit Function
		End If

		If Trim(.fpDateTime.Text = "") Then    
			Exit Function
		End If
		lgBlnFlgChgValue = True

		strSelect	=			 " a.dept_cd,a.dept_nm, a.org_change_id, a.internal_cd, b.biz_area_cd "
		strFrom		=			 " b_acct_dept a, b_cost_center b "
		strWhere	= " a.cost_cd = b.cost_cd "
		strWhere	= strWhere & " and a.dept_Cd = " & FilterVar(LTrim(RTrim(.txtDept1.value)), "''", "S")
		strWhere	= strWhere & " and a.org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(.fpDateTime.Text, parent.gDateFormat,""), "''", "S") & "))"			

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			.txtDept1.value       = ""
			.txtDept1Nm.value     = ""
			.txtCostCd.value      = ""
			.txtOrgChId.value     = ""
			.txtInternalCd1.value = ""
			.txtInternalCd2.value = ""
			.txtDept2.value       = ""
			.txtDept2Nm.value     = ""
			Call ggoOper.SetReqAttr(.txtDept2, "Q")
			.txtDept1.focus
			Set gActiveElement = document.activeElement  
		Else 
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
			jj = Ubound(arrVal1,1)

			For ii = 0 To jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))
				frm1.txtDept1Nm.value     = Trim(arrVal2(2))
			    frm1.txtOrgChId.value     = Trim(arrVal2(3))
			    frm1.txtInternalCd1.value = Trim(arrVal2(4))
			    frm1.txtCostCd.value      = Trim(arrVal2(5))
			    frm1.txtDept2.value       = ""
			    frm1.txtDept2Nm.value     = ""
			    frm1.txtDept2.focus
			    Call ggoOper.SetReqAttr(frm1.txtDept2, "N")
			Next
		End If
	End With

	lgBlnFlgChgValue = True   
End Function 

'========================================================================================================
' Desc : developer describe this line
'========================================================================================================
Function txtDept2_Onchange()
   Dim strSelect
	Dim strFrom
	Dim strWhere
    Dim IntRetCD
	Dim arrVal1
	Dim arrVal2
	Dim ii
	Dim jj

	With Frm1
		If .txtDept2.value = "" Then
			.txtDept2.value   = ""
			.txtDept2Nm.value = ""
			.txtDept2.focus
			lgBlnFlgChgValue  = True 
			Exit Function
		End If

		If Trim(.fpDateTime.Text = "") Then    
			Exit Function
		End If
	    lgBlnFlgChgValue = True

		strSelect	=			 " a.dept_cd,a.dept_nm "
		strFrom		=			 " b_acct_dept a, b_cost_center b "
		strWhere	= " a.cost_cd = b.cost_cd " 	 
		strWhere	= strWhere & " and a.dept_Cd = " & FilterVar(LTrim(RTrim(.txtDept2.value)), "''", "S")
		strWhere	= strWhere & " and b.biz_area_cd = " & FilterVar(LTrim(RTrim(.txtCostCd.value)), "''", "S")
		strWhere	= strWhere & " and a.org_change_id = (select distinct org_change_id "
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(.fpDateTime.Text, parent.gDateFormat,""), "''", "S") & "))"

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			.txtDept2.value    = ""
			.txtDept2Nm.value  = ""
			.txtDept2.focus
			Set gActiveElement = document.activeElement  
		Else 
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
			jj = Ubound(arrVal1,1)
			For ii = 0 to jj - 1
				arrVal2               = Split(arrVal1(ii), chr(11))
				frm1.txtDept2.value   = Trim(arrVal2(1))
			    frm1.txtDept2Nm.value = Trim(arrVal2(2))
			    frm1.txtCntAmt.focus
			Next
		End If
	End With
	
	lgBlnFlgChgValue = True
End Function 

'========================================================================================================
' Name : txtInsuerTp_Onchange
' Desc : developer describe this line
'========================================================================================================
Function txtInsuerTp_Onchange()
    Dim IntRetCd

    If frm1.txtInsuerTp.value = "" Then
		frm1.txtInsuerTp.value   = ""
		frm1.txtInsuerTpNm.value = ""
		frm1.txtInsuerTp.focus
    Else
        IntRetCD= CommonQueryRs(" MINOR_NM "," B_MINOR","  MAJOR_CD = " & FilterVar("A1030", "''", "S") & "  AND MINOR_CD = " & FilterVar(frm1.txtInsuerTp.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
		If IntRetCD=False And Trim(frm1.txtInsuerTp.value) <> "" Then
		    Call DisplayMsgBox("800054","X","X","X")                         '☜ : 등록되지 않은 코드입니다.
		    frm1.txtInsuerTp.value   = ""
		    frm1.txtInsuerTpNm.value = ""
		    frm1.txtInsuerTp.focus
		    Set gActiveElement       = document.activeElement
		Else
		    frm1.txtInsuerTpNm.value = Trim(Replace(lgF0,Chr(11),""))
		End If
    End If
    
	lgBlnFlgChgValue = True
End Function 

'========================================================================================================
' Name : txtCustomCd_Onchange
' Desc : developer describe this line
'========================================================================================================
Function txtCustomCd_Onchange()
    Dim IntRetCd

    If frm1.txtCustomCd.value = "" Then
		frm1.txtCustomNm.value = ""
    Else
        IntRetCD= CommonQueryRs(" BP_NM "," B_BIZ_PARTNER ","  BP_TYPE <= " & FilterVar("CS", "''", "S") & "  AND  BP_CD = " & FilterVar(frm1.txtCustomCd.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
		If IntRetCD=False And Trim(frm1.txtCustomCd.value) <> "" Then
		    Call DisplayMsgBox("800054","X","X","X")                         '☜ : 등록되지 않은 코드입니다.
		    frm1.txtCustomCd.value   = ""
		    frm1.txtCustomNm.value   = ""
		    frm1.txtTradeCur.value   = ""
			frm1.txtTradeCurNm.value = ""
			frm1.txtExRate.text      = 1
		    frm1.txtCustomCd.focus
		    Set gActiveElement       = document.activeElement  
		Else
		    frm1.txtCustomNm.value   = Trim(Replace(lgF0,Chr(11),""))
		    Call FncCurrency
		End If
    End If
    
	lgBlnFlgChgValue = True
End Function 

'========================================================================================================
' Name : txtTradeCur_Onchange
' Desc : developer describe this line
'========================================================================================================
Function txtTradeCur_Onchange()
    Dim IntRetCd
    
	If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtTradeCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
		Call CurFormatNumericOCX()
	End If	
		
	If frm1.txtTradeCur.value = "" Then
		frm1.txtTradeCur.value = ""
		frm1.txtTradeCurNm.value=""
		frm1.txtTradeCur.focus
	Else
	    IntRetCD= CommonQueryRs(" CURRENCY_DESC "," B_CURRENCY "," CURRENCY = " & FilterVar(frm1.txtTradeCur.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		If IntRetCD=False And Trim(frm1.txtTradeCur.value) <> "" Then
		    Call DisplayMsgBox("800054","X","X","X")                         '☜ : 등록되지 않은 코드입니다.
		    frm1.txtTradeCur.value   = ""
		    frm1.txtTradeCurNm.value = ""
		    frm1.txtTradeCur.focus
		    Set gActiveElement       = document.activeElement
		Else
		    frm1.txtTradeCurNm.value = Trim(Replace(lgF0,Chr(11),""))
		    Call FncRate("CURR")
		End If
	End If

	lgBlnFlgChgValue = True
End Function 

'========================================================================================================
' Name : txtInsureAcct_Onchange
' Desc : developer describe this line
'========================================================================================================
Function txtInsureAcct_Onchange()
    Dim IntRetCd

	If frm1.txtInsureAcct.value = "" Then
		frm1.txtInsureAcct.value   = ""
		frm1.txtInsureAcctNm.value = ""
		frm1.txtInsureAcct.focus
	Else
	    IntRetCD= CommonQueryRs(" ACCT_CD,ACCT_NM "," A_ACCT "," ACCT_CD = " & FilterVar(frm1.txtInsureAcct.value, "''", "S") & " and DEL_FG <> " & FilterVar("Y", "''", "S") & " " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		If IntRetCD=False And Trim(frm1.txtInsureAcct.value) <> "" Then
		    Call DisplayMsgBox("800054","X","X","X")                         '☜ : 등록되지 않은 코드입니다.
		    frm1.txtInsureAcct.value   = ""
		    frm1.txtInsureAcctNm.value = ""
		    frm1.txtInsureAcct.focus
		    Set gActiveElement         = document.activeElement
		Else
		    frm1.txtInsureAcct.value   = Trim(Replace(lgF0,Chr(11),""))
		    frm1.txtInsureAcctNm.value = Trim(Replace(lgF1,Chr(11),""))
		End If
	End If
	
	lgBlnFlgChgValue = True
End Function 

'========================================================================================================
Function txtCntAmt_Change()
	frm1.txtLocCntAmt.text = 0
	lgBlnFlgChgValue = True
End Function

'========================================================================================================
Function txtAmt_Change()
	frm1.txtLocAmt.text = 0
	lgBlnFlgChgValue = True
End Function

'========================================================================================================
' Name : txtCntAmt_onChange
' Desc : developer describe this line
'========================================================================================================
Function txtExRate_Change()
	Dim iRows
	Dim iVal

	frm1.txtLocAmt.text = 0
	frm1.txtLocCntAmt.text = 0
	lgBlnFlgChgValue = True
End Function

'========================================================================================================
' Name : cboPrivatePublic_onChange
' Desc : developer describe this line
'========================================================================================================
Function cboPrivatePublic_onChange()
  	lgBlnFlgChgValue = True
End Function

'========================================================================================================
' Name : cboCloseYesNo_onChange
' Desc : developer describe this line
'========================================================================================================
Function cboCloseYesNo_onChange()
  	lgBlnFlgChgValue = True
End Function

'========================================================================================================
' Name : cboCloseYesNo_onChange
' Desc : developer describe this line
'========================================================================================================
Function txtRefNo_onChange()
  	lgBlnFlgChgValue = True
End Function

'========================================================================================================
' Name : cboPrivatePublic_onChange
' Desc : developer describe this line
'========================================================================================================
Function txtLocAmt_Change()
  	lgBlnFlgChgValue = True
End Function

'========================================================================================================
' Name : cboCloseYesNo_onChange
' Desc : developer describe this line
'========================================================================================================
Function txtLocCntAmt_Change()
  	lgBlnFlgChgValue = True
End Function

'========================================================================================================
Sub fpDateTime_DblClick(Button)
	If Button = 1 Then
		frm1.fpDateTime.Action = 7
		Call SetFocusToDocument("M")
		frm1.fpDateTime.Focus
		lgBlnFlgChgValue = True
		Call deptCheck()
		Call FncRate("CURR")
	End If
End Sub

Sub txtCntFrom_DblClick(Button)
	If Button = 1 Then
		frm1.txtCntFrom.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtCntFrom.Focus
		lgBlnFlgChgValue = True
	End If
End Sub

Sub txtCntTo_DblClick(Button)
	If Button = 1 Then
		frm1.txtCntTo.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtCntTo.Focus
		lgBlnFlgChgValue = True
	End If
End Sub

Sub txtFromDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtFromDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFromDt.Focus
		lgBlnFlgChgValue = True
	End If
End Sub

Sub txtToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtToDt.Focus
		lgBlnFlgChgValue = True
	End If
End Sub

Sub fpDateTime_change()
	If CheckDateFormat(frm1.fpDateTime.Text,parent.gDateFormat) Then
		If Trim(frm1.txtTradeCur.value) <> "" Then
'			Call FncRate("CURR")
		End IF
	End If
End Sub

Function BtnPopupDisabled(Inobj) 
	If UCase(Inobj.className) = UCase("protected") Then 
		IsOpenPop = False
		BtnPopupDisabled = False
	Else
		BtnPopupDisabled = True
	End If
End Function

'========================================================================================================
'   Event Name : txtWork_dt_KeyPress
'   Event Desc :
'========================================================================================================
Sub txtCntFrom_KeyPress(Key)
    If key = 13 Then
        Call MainQuery
	End If
End Sub

'=======================================================================================================
'   Event Name : fpDateTime_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub fpDateTime_Change()
    lgBlnFlgChgValue = True
End Sub
'=======================================================================================================
Sub txtCntFrom_Change()
    lgBlnFlgChgValue = True
End Sub
'=======================================================================================================
Sub txtCntTo_Change()
    lgBlnFlgChgValue = True
End Sub
'=======================================================================================================
Sub txtFromDt_Change()
    lgBlnFlgChgValue = True
End Sub
'=======================================================================================================
Sub txtToDt_Change()
    lgBlnFlgChgValue = True
End Sub

'========================================================================================================
' Name : FncCurrency
' Desc : 통화를 Select
'========================================================================================================
Function FncCurrency()
	Dim IntRetCd, IntRetCD1

	Call FncCurrencyNm
	
	IntRetCD= CommonQueryRs(" CURRENCY "," B_BIZ_PARTNER ","  BP_TYPE <= " & FilterVar("CS", "''", "S") & "  AND BP_CD =  " & FilterVar(frm1.txtCustomCd.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If IntRetCD <> False AND Trim(Replace(lgF0,Chr(11),"")) <> "" Then
		frm1.txtTradeCur.value = Trim(Replace(lgF0,Chr(11),""))
		Call FncRate("CURR")
	End If
End Function

'========================================================================================================
' Name : FncCurrencyNm
' Desc : 통화의 이름을 Select
'========================================================================================================
Function FncCurrencyNm()
	Dim IntRetCd

	IntRetCD= CommonQueryRs(" CURRENCY_DESC "," B_CURRENCY "," CURRENCY = " & FilterVar(frm1.txtTradeCur.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	If IntRetCD <> False Then
		frm1.txtTradeCurNm.value = Trim(Replace(lgF0,Chr(11),""))
	End If
End Function

'========================================================================================================
' Name : FncRate
' Desc : 환율을 Select
'========================================================================================================
Function FncRate(InType)
	Dim IntRetCd,FrDt, Rate
	Dim strSelect, strFrom, strWhere

	If Trim(frm1.fpDateTime.Text) = "" Then
		frm1.fpDateTime.Text = UniConvDateAToB("<%=StratDate%>",parent.gServerDateFormat,parent.gDateFormat)
		Exit Function
	End IF

	FrDt = UniConvDateToYYYYMMDD(frm1.fpDateTime.Text,parent.gDateFormat,"")  'parent.UNIConvDate(frm1.fpDateTime.Text)
	Select Case UCase(InType)
		Case "DEFAULT"
			frm1.txtExRate.text = 1
			Call ggoOper.SetReqAttr(frm1.txtLocCntAmt, "Q")
			Call ggoOper.SetReqAttr(frm1.txtLocAmt, "Q")
			Call ggoOper.SetReqAttr(frm1.txtExRate, "Q")
		Case Else
			strSelect = strSelect & "isnull(case t.loc_cur when " & FilterVar(UCase(frm1.txtTradeCur.value), "''", "S") & " Then 1 "
			strSelect = strSelect & " 	Else  Case t.xch_rate_fg " 
			strSelect = strSelect & " 		  When " & FilterVar("M", "''", "S") & "  Then ( SELECT isnull(STD_RATE,0) "
			strSelect = strSelect & "				FROM	b_monthly_exchange_rate (nolock) "
			strSelect = strSelect & "				WHERE	apprl_yrmnth 	= CONVERT (varchar(06), " & FilterVar(FrDt,"" & FilterVar("29991231", "''", "S") & " ", "S") & ", 112) "
			strSelect = strSelect & "				and	from_currency	= " & FilterVar(UCase(frm1.txtTradeCur.value), "''", "S") 
			strSelect = strSelect & "				and	to_currency	= t.loc_cur ) "
			strSelect = strSelect & "		  Else (	SELECT  isnull(STD_RATE,0) "
			strSelect = strSelect & "				FROM	b_daily_exchange_rate (nolock) "
			strSelect = strSelect & "				WHERE	apprl_dt		= " & FilterVar(FrDt,"" & FilterVar("29991231", "''", "S") & " ", "S")
			strSelect = strSelect & "				and	from_currency	= " & FilterVar(UCase(frm1.txtTradeCur.value), "''", "S") 
			strSelect = strSelect & "				and	to_currency	= t.loc_cur ) "
			strSelect = strSelect & "		  End  "
			strSelect = strSelect & "End,0) as xch_rate "
			strFrom  = " (SELECT isnull(XCH_RATE_FG,'') as xch_rate_fg, loc_cur  from  b_company) t "    
			strWhere = ""

			IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			
			If IntRetCD <> False Then
				frm1.txtExRate.text = Trim(Replace(lgF0,Chr(11),""))
				Call ggoOper.SetReqAttr(frm1.txtLocCntAmt, "D")
				Call ggoOper.SetReqAttr(frm1.txtLocAmt, "D")
				Call ggoOper.SetReqAttr(frm1.txtExRate, "N")
			Else
				IntRetCD= CommonQueryRs(" LOC_CUR "," B_COMPANY "," CO_CD LIKE " & FilterVar("%", "''", "S") & " " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				If IntRetCD <> False And Trim(Replace(lgF0,Chr(11),"")) = Trim(UCase(frm1.txtTradeCur.value)) Then
					frm1.txtExRate.text = 1
					Call ggoOper.SetReqAttr(frm1.txtLocCntAmt, "Q")
					Call ggoOper.SetReqAttr(frm1.txtLocAmt, "Q")
					Call ggoOper.SetReqAttr(frm1.txtExRate, "Q")
				Else
					Call DisplayMsgBox("am0023","X","X","X")                         '☜ : 환율정보가 없습니다.
					frm1.txtExRate.text = 1
					Call ggoOper.SetReqAttr(frm1.txtLocCntAmt, "D")
					Call ggoOper.SetReqAttr(frm1.txtLocAmt, "D")
					Call ggoOper.SetReqAttr(frm1.txtExRate, "N")
					frm1.txtExRate.focus
					Set gActiveElement = document.activeElement
				End If
			End If
	End Select
End Function

Function fpDateTime_onblur()
	Call deptCheck()
	Call FncRate("CURR")
End Function

'========================================================================================================
' Name : deptCheck
' Desc : 부서체크 
'========================================================================================================
Function deptCheck()
	Dim strSelect
	Dim strFrom
	Dim strWhere
    Dim IntRetCD

	With Frm1
		If .txtDept1.value = "" Then
			.txtDept1.value       = ""
			.txtDept1Nm.value     = ""
			.txtCostCd.value      = ""
			.txtOrgChId.value     = ""
			.txtInternalCd1.value = ""
			.txtInternalCd2.value = ""
			.txtDept2.value       = ""
			.txtDept2Nm.value     = ""
			Call ggoOper.SetReqAttr(.txtDept2, "Q")
			.txtDept1.focus
			Set gActiveElement    = document.activeElement
			lgBlnFlgChgValue      = True 
			Exit Function
		End If

		If Trim(.fpDateTime.Text = "") Then
			Exit Function
		End If

		strSelect	=			 " distinct org_change_id "
		strFrom		=			 " b_acct_dept "
		strWhere	=			 " org_change_id = (select distinct org_change_id "
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(.fpDateTime.Text, parent.gDateFormat,""), "''", "S") & "))"

		IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

		If IntRetCD = False Or Trim(Replace(lgF0,Chr(11),"")) <> .txtOrgChId.value  Then
			.txtDept1.value       = ""
			.txtDept1Nm.value     = ""
			.txtCostCd.value      = ""
			.txtOrgChId.value     = ""
			.txtInternalCd1.value = ""
			.txtInternalCd2.value = ""
			.txtDept2.value       = ""
			.txtDept2Nm.value     = ""
			lgBlnFlgChgValue      = True
			Call ggoOper.SetReqAttr(.txtDept2, "Q")
			Exit Function
		End If
	End With
End Function

'========================================================================================================
' Name : chgBtnDisable
' Desc : 버튼변경 
'========================================================================================================
Sub chgBtnDisable(Gubun)
	Select Case Gubun
		Case 1		'//버튼 둘다 비활성 
			frm1.btnConf.disabled	=	True
			frm1.btnUnCon.disabled	=	True
		Case 2		'//확정버튼 활성화, 취소버튼비활성화 
			frm1.btnConf.disabled	=	False
			frm1.btnUnCon.disabled	=	True
		Case 3		'//확정버튼 비활성화, 취소버튼 활성화 
			frm1.btnConf.disabled	=	True
			frm1.btnUnCon.disabled	=	False
	End Select
End Sub

'========================================================================================================
' Name : fnBttnConf
' Desc : 전표작업 
'========================================================================================================
Sub fnBttnConf(Gubun)
	Dim IntRetCD
	Dim strVal
	Dim strYear,strMonth,strDay, txtGlDt
	Err.Clear 

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X") '☜ "데이타가 변경되었습니다. 계속 하시겠습니까?"
		If IntRetCD = vbNo Then
			Exit Sub
		End If
	End If	

	If LayerShowHide(1) = False Then
		Exit Sub
	End If

	Call ExtractDateFrom(frm1.fpDateTime,parent.gDateFormat,parent.gComDateType,strYear,strMonth,strDay)
	txtGlDt = strYear & strMonth & strDay
    
    strVal = BIZ_PGM_ID1 & "?txtMode="          & Gubun
    strVal = strVal     & "&txtInsuerCd1="      & Trim(frm1.txtInsuerCd1.value)
    strVal = strVal     & "&txtGlDt="			& Trim(txtGlDt)

	Call RunMyBizASP(MyBizASP, strVal)
    Set gActiveElement = document.ActiveElement
End Sub


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
</HEAD>

<BODY SCROLL="NO" TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">

<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 >
							<TR>
								<td background="../../image/table/seltab_up_bg.gif"><img src="../../image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>보험정보 기초등록</font></td>
								<td background="../../image/table/seltab_up_bg.gif" align="right"><img src="../../image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT></TD>
					<TD WIDTH=*>&nbsp;</TD>
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
									<TD CLASS=TD5 NOWRAP>보험코드</TD>
									<TD CLASS=TD656 NOWRAP><INPUT NAME="txtInsuerCd" TYPE=TEXT SIZE=20 MAXLENGTH="25" TAG="12XXXU" ALT="보험코드"><IMG SRC="../../image/btnPopup.gif" NAME="btnInsuerCd" align=top TYPE="BUTTON" ONCLICK ="vbscript:Call BtnPopUp(frm1.btnInsuerCd,0)">&nbsp;<INPUT TYPE=TEXT NAME="txtInsuerNm"  SIZE="30" MAXLENGTH="30" TAG="24"></TD></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>보험코드</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtInsuerCd1" TYPE=TEXT SIZE="20" MAXLENGTH="20"   TAG="23XXXU" ALT="보험코드"></TD>
								<TD CLASS=TD5 NOWRAP>보험명</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtInsuerNm1" TYPE=TEXT SIZE="30" MAXLENGTH="30"   TAG="23XXXX" ALT="보험명"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>보험종류</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtInsuerTp" TYPE=TEXT SIZE=10  MAXLENGTH="2" TAG="23XXXU" ALT="보험종류"><IMG SRC="../../image/btnPopup.gif" NAME="btnInsuerTp" align=top TYPE="BUTTON" ONCLICK ="vbscript:Call BtnPopUp(frm1.txtInsuerTp,1)">&nbsp;<INPUT TYPE=TEXT NAME="txtInsuerTpNm"  SIZE="20" MAXLENGTH="50" TAG="24"></TD>
								<TD CLASS=TD5 NOWRAP>보험료계정</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtInsureAcct" SIZE="10"  MAXLENGTH="20" TAG="23XXXU" ALT="보험료계정" OnChange="txtInsureAcct_Onchange()"><IMG SRC="../../image/btnPopup.gif" NAME="btnInsureAcct" align=top TYPE="BUTTON" ONCLICK ="vbscript:Call BtnPopUp(frm1.txtInsureAcct,8)">&nbsp;<INPUT TYPE=TEXT NAME="txtInsureAcctNm" SIZE="20" MAXLENGTH="50" TAG="24"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>거래처</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCustomCd" TYPE=TEXT SIZE=10 MAXLENGTH="10" TAG="23XXXU" ALT="거래처"><IMG SRC="../../image/btnPopup.gif" NAME="btnCustomCd" align=top TYPE="BUTTON" ONCLICK ="vbscript:Call OpenBp(frm1.txtCustomCd.value,3)">&nbsp;<INPUT TYPE=TEXT NAME="txtCustomNm"  SIZE="20" MAXLENGTH="50" TAG="24"></TD>
								<TD CLASS=TD5 NOWRAP>전표일자</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5969ma1_fpDateTime1_fpDateTime.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>거래통화</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTradeCur" SIZE="10"  MAXLENGTH="3" TAG="23XXXU" ALT="거래통화"><IMG SRC="../../image/btnPopup.gif" NAME="btnTradeCur" align=top TYPE="BUTTON" ONCLICK ="vbscript:Call BtnPopUp(frm1.txtTradeCur,2)">&nbsp;<INPUT TYPE=TEXT NAME="txtTradeCurNm" SIZE="20" MAXLENGTH="50" TAG="24"></TD>
								<TD CLASS=TD5 NOWRAP>환율</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5969ma1_txtExRate1_txtExRate.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>발의부서</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDept1" TYPE=TEXT SIZE=10  MAXLENGTH="10" TAG="23XXXU" ALT="발의부서"><IMG SRC="../../image/btnPopup.gif" NAME="btnDept1" align=top TYPE="BUTTON" ONCLICK ="vbscript:Call OpenDept(frm1.txtDept1.value,1)">&nbsp;<INPUT TYPE=TEXT NAME="txtDept1Nm"  SIZE="20" MAXLENGTH="50" TAG="24"></TD>
								<TD CLASS=TD5 NOWRAP>귀속부서</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDept2" TYPE=TEXT SIZE=10  MAXLENGTH="10" TAG="24XXXU" ALT="귀속부서"><IMG SRC="../../image/btnPopup.gif" NAME="btnDept2" align=top TYPE="BUTTON" ONCLICK ="vbscript:Call OpenDept(frm1.txtDept2.value,2)">&nbsp;<INPUT TYPE=TEXT NAME="txtDept2Nm"  SIZE="20" MAXLENGTH="50" TAG="24"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>계약금액</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5969ma1_txtCntAmt1_txtCntAmt.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>계약금액(자국)</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5969ma1_txtLocCntAmt1_txtLocCntAmt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>보험료</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5969ma1_txtAmt1_txtAmt.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>보험료(자국)</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5969ma1_txtLocAmt1_txtLocAmt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>계약기간</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5969ma1_txtCntFrom1_txtCntFrom.js'></script>~
									<script language =javascript src='./js/a5969ma1_txtCntTo1_txtCntTo.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>적용기간(년월)</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5969ma1_txtFromDt1_txtFromDt.js'></script>~
									<script language =javascript src='./js/a5969ma1_txtToDt1_txtToDt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>계약완료</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboPrivatePublic" TAG="23" ALT="완료여부"><OPTION VALUE="N">N</OPTION><OPTION VALUE="Y">Y</OPTION></SELECT></TD>
								<TD CLASS=TD5 NOWRAP>관리번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRefNo" TYPE=TEXT SIZE="30" MAXLENGTH="20"   TAG="21XXXU" ALT="관리번호"></TD>
							</TR>
							<!--요청이 있을시 주석풀기<TR>
								<TD CLASS=TD5 NOWRAP>결의전표번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTempGlNo" TYPE=TEXT SIZE="30" MAXLENGTH="18"   TAG="24XXXU" ALT="결의전표번호"></TD>
								<TD CLASS=TD5 NOWRAP>회계전표번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtGlNo" TYPE=TEXT SIZE="30" MAXLENGTH="18"   TAG="24XXXU" ALT="회계전표번호"></TD>
							</TR>-->
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<!--요청이 있을시 주석풀기 
	<TR HEIGHT="20">
		<TD HEIGHT=20 WIDTH="100%">
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR HEIGHT=20>
						<TD CLASS=TD6 NOWRAP></TD>
						<TD CLASS=TD6 NOWRAP></TD>
						<TD CLASS=TD6 NOWRAP><BUTTON NAME="btnConf" CLASS="CLSMBTN" OnClick="VBScript:Call fnBttnConf('C')">전표확정</BUTTON>&nbsp<BUTTON NAME="btnUnCon" CLASS="CLSMBTN" OnClick="VBScript:Call fnBttnConf('D')">전표취소</BUTTON></TD>
						<TD CLASS=TD6 NOWRAP></TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>-->
	<TR >
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtCHKMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtPrevNext" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtCostCd" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtOrgChId" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="UserPrevNext" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInternalCd1" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInternalCd2" tag="24" TABINDEX="-1">


</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm" TABINDEX="-1"></IFRAME>
</DIV>
</BODY>
</HTML>
