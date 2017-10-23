<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : 월차 결산 
'*  2. Function Name        :
'*  3. Program ID           : A5970MA1
'*  4. Program Name         : A5970MA1
'*  5. Program Desc         : 유가증권 기초등록 
'*  6. Component List       :
'*  7. Modified date(First) : 2003/01/09
'*  8. Modified date(Last)  : 2003/10/20
'*  9. Modifier (First)     : leenamyo
'* 10. Modifier (Last)      : Jeong Yong Kyun
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
Const BIZ_PGM_ID    = "a5970mb1.asp"
Const BIZ_PGM_ID1    = "a5970mb2.asp"
Const BIZ_PGM_JUMP_ID   = "a5959ma1"									'사업부별 손익비교(컴퍼니 메뉴에 등록된 명)
Const COOKIE_SPLIT  =  4877	                                                        'Cookie Split String

'========================================================================================================

Const gIsShowLocal = "Y"
<%
Const gIsShowLocal = "Y"
%>

'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->

'========================================================================================================
Dim lgIsOpenPop
Dim IsOpenPop
Dim lgRecordPage
Dim UserPrevNext

<%
StartDate	= GetSvrDate                                               'Get Server DB Date
%>
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE
	lgBlnFlgChgValue  = False
	lgIntGrpCount     = 0
    lgStrPrevKey      = ""
    lgStrPrevKeyIndex = 0
    lgSortKey         = 1
    lgLngCurRows = 0                                            'initializes Deleted Rows Count
End Sub

'========================================================================================================
Sub SetDefaultVal()
	Call ggoOper.FormatDate(frm1.txtBillDt, parent.gDateFormat, 1)
	Call ggoOper.FormatDate(frm1.txtPubDt, parent.gDateFormat, 1)
	Call ggoOper.FormatDate(frm1.txtExpireDt, parent.gDateFormat, 1)
	Call ggoOper.FormatDate(frm1.txtInDt, parent.gDateFormat, 1)
	Call ggoOper.FormatDate(frm1.txtOutDt, parent.gDateFormat, 1)
    Call ggoOper.SetReqAttr(frm1.txtDept2,"Q")
    Call ggoOper.SetReqAttr(frm1.txtOutDt,"Q")
    frm1.txtBillDt.text	= UniConvDateAToB("<%=StartDate%>",parent.gServerDateFormat,parent.gDateFormat)
    frm1.txtInDt.text	= UniConvDateAToB("<%=StartDate%>",parent.gServerDateFormat,parent.gDateFormat)

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
    Dim txtBillDt, txtPubDt, txtExpireDt, txtInDt, txtOutDt
    Dim strYear, strMonth, strDay

	Call ExtractDateFrom(frm1.txtBillDt.Text,frm1.txtBillDt.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
    txtBillDt = strYear  & strMonth  & strDay

    Call ExtractDateFrom(frm1.txtPubDt.Text,frm1.txtPubDt.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
    txtPubDt = strYear & strMonth & strDay

    Call ExtractDateFrom(frm1.txtExpireDt.Text,frm1.txtExpireDt.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
    txtExpireDt = strYear & strMonth & strDay

    Call ExtractDateFrom(frm1.txtInDt.Text,frm1.txtInDt.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
    txtInDt = strYear & strMonth & strDay

    Call ExtractDateFrom(frm1.txtOutDt.Text,frm1.txtOutDt.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
    txtOutDt = strYear & strMonth & strDay

	Select Case pOpt
		Case "S"
			lgKeyStream = Trim(frm1.txtSecuCode1.value) & parent.gColSep
			lgKeyStream = lgKeyStream & Trim(frm1.txtSecuNm1.value) & parent.gColSep
			lgKeyStream = lgKeyStream & Trim(frm1.txtSecuType.value) & parent.gColSep
			lgKeyStream = lgKeyStream & Trim(frm1.txtDept1Area.value) & parent.gColSep
			lgKeyStream = lgKeyStream & Trim(frm1.txtDept1OrgId.value) & parent.gColSep
			lgKeyStream = lgKeyStream & Trim(frm1.txtDept1.value) & parent.gColSep
			lgKeyStream = lgKeyStream & Trim(frm1.txtDept2.value) & parent.gColSep
			lgKeyStream = lgKeyStream & txtPubDt & parent.gColSep
			lgKeyStream = lgKeyStream & txtExpireDt & parent.gColSep
			lgKeyStream = lgKeyStream & txtInDt & parent.gColSep
			lgKeyStream = lgKeyStream & Trim(frm1.txtCust1.value) & parent.gColSep
			lgKeyStream = lgKeyStream & Trim(frm1.txtCust2.value) & parent.gColSep
			lgKeyStream = lgKeyStream & Trim(frm1.txtTradeCur.value) & parent.gColSep
			lgKeyStream = lgKeyStream & Trim(frm1.txtXchRate.text) & parent.gColSep
			lgKeyStream = lgKeyStream & Trim(frm1.txtBuyAmt.text) & parent.gColSep
			If UNICDbl(frm1.txtLocBuyAmt.text) = 0 or Trim(frm1.txtLocBuyAmt.text) = "" Then
			    lgKeyStream = lgKeyStream & UNICDbl(frm1.txtXchRate.text) * UNICDbl(frm1.txtBuyAmt.text) & parent.gColSep
			Else
			    lgKeyStream = lgKeyStream & Trim(frm1.txtLocBuyAmt.text) & parent.gColSep
			End If
			lgKeyStream = lgKeyStream & Trim(frm1.txtCalRate.text) & parent.gColSep
			lgKeyStream = lgKeyStream & Trim(frm1.selCalYn.value) & parent.gColSep
			lgKeyStream = lgKeyStream & Trim(frm1.selEndYn.value) & parent.gColSep
			lgKeyStream = lgKeyStream & Trim(frm1.txtPriceAmt.text) & parent.gColSep
			If UNICDbl(frm1.txtLocPriceAmt.text) = 0 or Trim(frm1.txtLocPriceAmt.text) = "" Then
			    lgKeyStream = lgKeyStream & UNICDbl(frm1.txtXchRate.text) * UNICDbl(frm1.txtPriceAmt.text) & parent.gColSep
			Else
			    lgKeyStream = lgKeyStream & Trim(frm1.txtLocPriceAmt.text) & parent.gColSep
			End If
			lgKeyStream = lgKeyStream & Trim(frm1.txtCnt.Text) & parent.gColSep
			lgKeyStream = lgKeyStream & txtOutDt & parent.gColSep
			lgKeyStream = lgKeyStream & Trim(frm1.selComYn.value) & parent.gColSep
			lgKeyStream = lgKeyStream & Trim(frm1.txtRefNo.value) & parent.gColSep
			lgKeyStream = lgKeyStream & txtBillDt & parent.gColSep
		Case "D" 
			lgKeyStream = Trim(frm1.txtSecuCode1.value)        & parent.gColSep
	End Select    
End Sub

'========================================================================================================
Sub InitComboBox()
	Dim iCodeArr
	Dim iNameArr
	Dim i, isize
	Dim IntRetCD1
	
	On Error Resume Next
	Err.Clear 
	
	i = 0
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1080", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = Split(lgF0,Chr(11))
    iNameArr = Split(lgF1,Chr(11))

    If isArray(iCodeArr) Then
        Do While Not isNull(iCodeArr(i))
            i = i + 1
            If iCodeArr(i) = "" Then
                Exit Do
            End If
        Loop
        isize = i
        frm1.selEndYn.length = isize

        For i = 0 to isize-1
            frm1.selEndYn.options(i).value = iCodeArr(i)
            frm1.selEndYn.options(i).text	= iNameArr(i)
        Next
    End If
End Sub

'========================================================================================================
Sub InitData()
	frm1.txtXchRate.text = 1
	frm1.txtBuyAmt.text = 0
	frm1.txtLocBuyAmt.text = 0
	frm1.txtPriceAmt.text = 0
	frm1.txtLocPriceAmt.text = 0
	frm1.txtCnt.text = 0
	frm1.txtCalRate.text = 0
    frm1.selCalYn.value = "Y"
    frm1.selComYn.value = "N"
    frm1.selEndYn.value = "O"
    frm1.txtSecuCode.focus
    lgBlnFlgChgValue = False
End Sub

'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX 
'====================================================================================================
Sub CurFormatNumericOCX()
	With frm1
		'해당되는 금액,환율이 있는 Data 필드에 대하여 각각 처리 
		ggoOper.FormatFieldByObjectOfCur .txtXchRate, .txtTradeCur.value, parent.ggExchRateNo  , gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtCalRate, .txtTradeCur.value, parent.ggExchRateNo  , gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtBuyAmt,  .txtTradeCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtPriceAmt,.txtTradeCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
	End With
End Sub

'========================================================================================================
Sub Form_Load()
    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format

	Call AppendNumberPlace("6","12","0")	
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                                           '⊙: Lock  Suitable  Field

    Call InitVariables                                                              '⊙: Setup the Spread sheet
	Call InitComboBox
	Call InitData
	Call SetDefaultVal
	Call SetToolbar("1111100000111111")                                             '☆: Developer must customize
	'//전표요청이 있을시 주석 풀기 
	'Call chgBtnDisable(1)
End Sub

'========================================================================================================
Function FncQuery()
    Dim IntRetCD
    FncQuery = False																'☜: Processing is NG
    Err.Clear																		'☜: Clear err status

	If Not chkField(Document, "1") Then												'☜: This function check required field
		Exit Function
    End If

    If lgBlnFlgChgValue = True  Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X") '☜ "데이타가 변경되었습니다. 조회하시겠습니까?"
	
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If	

    Call ggoOper.ClearField(Document, "2")											'☜: Clear Contents  Field
    Call ggoOper.LockField(Document , "N")											'☜: Lock  Field
	Call InitData
 	Call SetDefaultVal()
    Call InitVariables
	Call InitComboBox
	'//전표요청이 있을시 주석 풀기 
    'Call chgBtnDisable(1)
    If DbQuery = False Then
		Exit Function
    End If																			'☜: Query db data
	
    Set gActiveElement = document.ActiveElement
    FncQuery = True																	'☜: Processing is OK
End Function

'========================================================================================================
Function FncNew()
    Dim IntRetCD

    FncNew = False																	'☜: Processing is NG
    Err.Clear																		'☜: Clear err status

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"x","x")					'☜: Data is changed.  Do you want to make it new?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    Call ggoOper.ClearField(Document, "1")											'☜: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")											'☜: Clear Contents  Field
    Call ggoOper.LockField(Document , "N")											'☜: Lock  Field
	Call InitVariables                                                              '⊙: Setup the Spread sheet
	Call InitData
	Call InitComboBox
	Call SetDefaultVal
	Call SetToolbar("1111100000101111")
	'//전표요청이 있을시 주석 풀기 
    'Call chgBtnDisable(1)
    Call txtTradeCur_OnChange()
    Set gActiveElement = document.ActiveElement
    FncNew = True
End Function

'========================================================================================================
Function FncDelete()
    Dim intRetCD

    FncDelete = False																'☜: Processing is NG
    Err.Clear																		'☜: Clear err status

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                       '☜: Please do Display first.
        Call DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"x","x")                    '☜: Do you want to delete?
	If IntRetCD = vbNo Then
		Exit Function
	End If

    If Not chkField(Document, "1") Then												'⊙: This function check indispensable field
       Exit Function
    End If

	Call MakeKeyStream("D")
    If DbDelete = False Then														'☜: Query db data
		Exit Function
    End If

    Set gActiveElement = document.ActiveElement
    FncDelete = True																'☜: Processing is OK
End Function

'========================================================================================================
Function FncSave()
	Dim IntRetCD
	Dim FrDt
	Dim strSelect, strFrom, strWhere
	
    FncSave = False																	'☜: Processing is NG
    Err.Clear																		'☜: Clear err status

    Call deptCheck()
    
    If Not chkField(Document, "2") Then
		Exit Function
    End If
	
	If frm1.selCalYn.value = "Y" Then
        If CompareDateByFormat(frm1.txtBillDt.text,frm1.txtExpireDt.text,frm1.txtBillDt.Alt,frm1.txtExpireDt.Alt,"970023",frm1.txtBillDt.UserDefinedFormat,parent.gComDateType,True) = False Then
		    Exit Function
        End If
    End If
	
	If CompareDateByFormat(frm1.txtInDt.Text, frm1.txtBillDt.Text,frm1.txtInDt.Alt,frm1.txtBillDt.Alt, "970025", frm1.txtInDt.UserDefinedFormat, parent.gComDateType, True)=False Then  '☜ : 시장월은 종료월보다 작아야합니다. 
        frm1.txtBillDt.focus
        Set gActiveElement = document.ActiveElement
        Exit Function
    End if 
	
    If frm1.selComYn.value = "Y" Then
        If CompareDateByFormat(frm1.txtBillDt.text,frm1.txtOutDt.text,frm1.txtBillDt.Alt,frm1.txtOutDt.Alt,"970023",frm1.txtBillDt.UserDefinedFormat,parent.gComDateType,True) = False Then
		    Exit Function
        End If
    End If
	
	If UNICDbl(frm1.txtXchRate.text) = 0 Then
		Call DisplayMsgBox("800443","X",frm1.txtXchRate.alt,"0")					'☜ : 숫자영 '//환율 
		frm1.txtXchRate.focus
		Set gActiveElement = document.activeElement  	
		Exit Function
	End If	
	
	If UNICDbl(frm1.txtBuyAmt.text) = 0 Then
		Call DisplayMsgBox("800443","X",frm1.txtBuyAmt.alt,"0")                     '☜ : 숫자영'//취득 
		frm1.txtBuyAmt.focus
		Set gActiveElement = document.activeElement  	
		Exit Function
	End If	
	
	If UNICDbl(frm1.txtPriceAmt.text) = 0 Then
		Call DisplayMsgBox("800443","X",frm1.txtPriceAmt.alt,"0")                   '☜ : 숫자영'//액면 
		frm1.txtPriceAmt.focus
		Set gActiveElement = document.activeElement  	
		Exit Function
	End If	
	
	If UNICDbl(frm1.txtCnt.Text) = 0 Then
		Call DisplayMsgBox("800443","X",frm1.txtCnt.alt,"0")                        '☜ : 숫자영'//매수 
		frm1.txtCnt.focus
		Set gActiveElement = document.activeElement  	
		Exit Function
	End If	
	
	If UNICDbl(frm1.txtCalRate.text) = 0 Then
		Call DisplayMsgBox("141157","X","X","X")									'☜ : 숫자영'//이자율 
		frm1.txtCalRate.focus
		Set gActiveElement = document.activeElement  	
		Exit Function
	End If	
	
	FrDt = UniConvDateToYYYYMMDD(frm1.txtInDt.Text,parent.gDateFormat,"")   '//parent.UNIConvDate(frm1.txtInDt.Text)
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

    If DbSave = False Then
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
	frm1.txtSecuCode.value = ""
	frm1.txtSecuCode1.value = ""
	frm1.txtSecuNm.value = ""
	frm1.txtSecuNm1.value = ""

    Set gActiveElement = document.ActiveElement   
    FncCopy = True
End Function

'==========================================================================================
Function FncPrint()
    FncPrint = False
    Err.Clear
	Call Parent.FncPrint()
    FncPrint = True
End Function

'==========================================================================================
Function FncExcel()
    FncExcel = False
    Err.Clear
	Call Parent.FncExport(parent.C_SINGLE)
    FncExcel = True
End Function

'==========================================================================================
Function FncFind()
    FncFind = False
    Err.Clear
	Call Parent.FncFind(parent.C_SINGLE, True)
    FncFind = True
End Function

'==========================================================================================
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

'==========================================================================================
Function DbQuery()
	Dim strVal

    Err.Clear
    DbQuery = False

	If LayerShowHide(1) = False then
	    Exit Function
	End If

    With frm1
		strVal = BIZ_PGM_ID & "?txtMode="          & parent.UID_M0001
		strVal = strVal     & "&txtPrevNext="      & ""
		strVal = strVal     & "&txtSecuCode="      & Trim(frm1.txtSecuCode.value)
		strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex
		strVal = strVal		& "&lgCurrency="	   & frm1.txtTradeCur.value
    End With
	
	Call RunMyBizASP(MyBizASP, strVal)

    DbQuery = True
     Set gActiveElement = document.ActiveElement   
End Function

'==========================================================================================
Function DbSave()
	Dim lGrpcnt 
	Dim strVal, strDel
	Dim IntRows

    Err.Clear
    DbSave = False
   	lGrpCnt =0

    If LayerShowHide(1) = False Then
		Exit Function
	End If

  	With frm1
		.txtMode.value        = parent.UID_M0002								'☜: Delete
        .txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                     '☜: Save Key
        .txtUpdtUserId.value  = parent.gUsrID
        .txtInsrtUserId.value = parent.gUsrID
  	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    DbSave = True

    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================================
Function DbDelete()
	Dim strVal

    Err.Clear
	DbDelete = False

	If LayerShowHide(1) = False Then
		Exit Function
	End If

    strVal = BIZ_PGM_ID & "?txtMode="          & parent.UID_M0003
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream

	Call RunMyBizASP(MyBizASP, strVal)
	DbDelete = True
End Function

'========================================================================================================
Sub DbQueryOk()
	dim iRow
	Dim varData
	Dim intIndex
	
	lgIntFlgMode      = parent.OPMD_UMODE

    If Trim(frm1.txtSecuCode1.value) <> "" Then
        Call ggoOper.SetReqAttr(frm1.txtSecuCode1,"Q")
        Call ggoOper.SetReqAttr(frm1.txtDept2,"N")
    End If
   
		
	Call ggoOper.LockField(Document, "N")
	Call ggoOper.SetReqAttr(frm1.txtSecuCode1,"Q")
	selCalYn_Change()
	selComYn_Change()
	'//Button 조작 
	'//전표요청이 있을시 주석 풀기 
	'If trim(frm1.txtSecuCode1.value) = "" and trim(frm1.txtTempGlNo.value) = "" and trim(frm1.txtGlNo.value) = "" Then
	'	Call chgBtnDisable(1)
	'ElseIF trim(frm1.txtSecuCode1.value) <> "" and trim(frm1.txtTempGlNo.value) = "" and trim(frm1.txtGlNo.value) = "" Then
	'	Call chgBtnDisable(2)
	'ElseIF trim(frm1.txtSecuCode1.value) <> "" and trim(frm1.txtTempGlNo.value) <> "" or trim(frm1.txtGlNo.value) <> "" Then
	'	Call chgBtnDisable(3)
	'End If	
	
    Set gActiveElement = document.ActiveElement  
	Call SetToolbar("111110000011111")                                          '☆: Developer must customize

	lgBlnFlgChgValue = False       
End Sub

'========================================================================================================
Sub DbSaveOk()
	On Error Resume Next
	
    Call InitComboBox() 
	
	Set gActiveElement = document.ActiveElement   
	frm1.txtSecuCode.value = frm1.txtSecuCode1.value
	Call DbQuery()
End Sub

'========================================================================================================
Sub DbDeleteOk()
	Call InitVariables()
	Call FncNew()
End Sub

'========================================================================================================
Function btnSecuCodeOnClick()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "유가증권"		    									' 팝업 명칭 
	arrParam(1) = "A_SECURITY"													' TABLE 명칭 
	arrParam(2) = frm1.txtSecuCode.value										' Code Condition
	arrParam(3) = "" 		            										' Name Condition
	arrParam(4) = "ISNULL(TEMP_GL_NO,'') = '' AND ISNULL(GL_NO, '') = ''"					
	arrParam(5) = "유가증권"

    arrField(0) = "SECURITY_CD"	     											' Field명(1)
    arrField(1) = "SECURITY_NM"													' Field명(0)


    arrHeader(0) = "유가증권코드"			    							' Header명(0)
    arrHeader(1) = "유가증권명"												' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=470px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
        frm1.txtSecuCode.focus
		Exit Function
	Else
		Call SetSecuCode(arrRet)
	End If
End Function

'======================================================================================================
'	Name : SetSecuCode()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetSecuCode(Byval arrRet)
    With frm1
        .txtSecuCode.focus
        .txtSecuCode.value  = arrRet(0)
        .txtSecuNm.value    = arrRet(1)
    End With
End Function

'======================================================================================================
'	Name : btnSecuTypeOnClick()
'	Description : Major PopUp
'=======================================================================================================
Function btnSecuTypeOnClick(Inobj)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	If BtnPopupDisabled(Inobj) = False Then Exit Function	

	IsOpenPop = True

	arrParam(0) = "유가증권종류"		    								' 팝업 명칭 
	arrParam(1) = "B_MINOR"														' TABLE 명칭 
	arrParam(2) = frm1.txtSecuType.value										' Code Condition
	arrParam(3) = "" 		            										' Name Cindition
	arrParam(4) = " MAJOR_CD = " & FilterVar("A1031", "''", "S") & "  "										' Where Condition
	arrParam(5) = "유가증권종류"

    arrField(0) = "MINOR_CD"	     											' Field명(1)
    arrField(1) = "MINOR_NM"													' Field명(0)


    arrHeader(0) = "유가증권종류코드"			    						' Header명(0)
    arrHeader(1) = "유가증권종류명"											' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=470px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
        frm1.txtSecuType.focus
		Exit Function
	Else
		Call SetSecuType(arrRet)
	End If
End Function

'======================================================================================================
'	Name : SetSecuType()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetSecuType(Byval arrRet)
    With frm1
        .txtSecuType.focus
        .txtSecuType.value   = arrRet(0)
        .txtSecuTypeNm.value = arrRet(1)
    End With
    
    lgBlnFlgChgValue = True 
End Function

'======================================================================================================
'	Name : btnTradeCurOnClick()
'	Description : Major PopUp
'=======================================================================================================
Function btnTradeCurOnClick(Inobj)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	If BtnPopupDisabled(Inobj) = False Then Exit Function	

	IsOpenPop = True

	arrParam(0) = "거래통화"		    										' 팝업 명칭 
	arrParam(1) = "B_CURRENCY"														' TABLE 명칭 
	arrParam(2) = frm1.txtTradeCur.value											' Code Condition
	arrParam(3) = "" 		            											' Name Cindition
	arrParam(4) = ""																' Where Condition
	arrParam(5) = "거래통화"

    arrField(0) = "CURRENCY"	     												' Field명(1)
    arrField(1) = "CURRENCY_DESC"													' Field명(0)


    arrHeader(0) = "통화코드"			    									' Header명(0)
    arrHeader(1) = "통화명"														' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=470px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtTradeCur.focus
		Exit Function
	Else
		Call SetTradeCur(arrRet)
	End If
End Function

'======================================================================================================
'	Name : SetTradeCur()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetTradeCur(Byval arrRet)
    With frm1
        .txtTradeCur.focus
        .txtTradeCur.value  = arrRet(0)
        .txtTradeCurNm.value    = arrRet(1)
    End With
    Call txtTradeCur_OnChange()
End Function

'======================================================================================================
'	Name : OpenDept
'	Description : 
'=======================================================================================================
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

	arrParam(0) = strCode														' Code Condition
   	arrParam(1) = frm1.txtBillDt.Text
	arrParam(2) = lgUsrIntCd													' 자료권한 Condition  

	' T : protected F: 필수 
	If lgIntFlgMode = Parent.OPMD_UMODE then
		arrParam(3) = "T"														' 결의일자 상태 Condition  
	Else
		arrParam(3) = "F"														' 결의일자 상태 Condition  
	End If
	
	arrParam(4) = iWhere
	arrParam(5) = Trim(frm1.txtDept1Area.value)
	
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
				.txtDept1Area.value = arrRet(2)
				.txtDept1OrgId.value = arrRet(3)
				.txtInternalCd1.value = arrRet(4)
				.txtBillDt.text = arrRet(5)
				.txtDept2.value = ""
				.txtDept2Nm.value = ""
				Call txtDept1_OnChange()  
			Case 2
				.txtDept2.value = arrRet(0)
				.txtDept2Nm.value = arrRet(1)
				.txtInternalCd2.value = arrRet(4)
				.txtBillDt.text = arrRet(5)
				Call txtDept2_OnChange()
		End Select
	End With
End Function

'------------------------------------------  OpenBp()  ---------------------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function
	
	if UCase(frm1.txtCust1.className) = "PROTECTED" Then Exit Function
	IsOpenPop = True

	arrParam(0) = strCode														'Code Condition
   	arrParam(1) = ""															'채권과 연계(거래처 유무)
	arrParam(2) = ""															'FrDt
	arrParam(3) = ""															'ToDt
	arrParam(4) = "T"															'B :매출 S: 매입 T: 전체 
	arrParam(5) = ""															'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 	
	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
        frm1.txtCust1.focus
		Exit Function
	Else
		Call SetCust1(arrRet)
		lgBlnFlgChgValue = True
	End If
End Function

'======================================================================================================
'	Name : btnCust1OnClick()
'	Description : Major PopUp
'=======================================================================================================
Function btnCust1OnClick(Inobj)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	If BtnPopupDisabled(Inobj) = False Then Exit Function	

	IsOpenPop = True

	arrParam(0) = "발행거래처"		    										' 팝업 명칭 

    arrParam(1) = "B_BIZ_PARTNER"
    arrParam(2) = Trim(frm1.txtCust1.value)
    arrParam(3) = ""
    arrParam(4) = ""
    arrParam(5) = "발행거래처코드"

    arrField(0) = "BP_CD"
    arrField(1) = "BP_NM"

    arrHeader(0) = "거래처코드"
    arrHeader(1) = "거래처명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=470px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
        frm1.txtCust1.focus
		Exit Function
	Else
		Call SetCust1(arrRet)
	End If
End Function

'=======================================================================================================
'	Name : SetCust1()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetCust1(Byval arrRet)
    With frm1
        .txtCust1.focus
        .txtCust1.value   = arrRet(0)
        .txtCust1Nm.value = arrRet(1)
    End With
    Call txtCust1_Change()
End Function

'=======================================================================================================
'	Name : btnCust2OnClick()
'	Description : Major PopUp
'=======================================================================================================
Function btnCust2OnClick(Inobj)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	If BtnPopupDisabled(Inobj) = False Then Exit Function	

	IsOpenPop = True

	arrParam(0) = "주식보관회사"		    										' 팝업 명칭 

    arrParam(1) = "B_BIZ_PARTNER"
    arrParam(2) = Trim(frm1.txtCust2.value)
    arrParam(3) = ""
    arrParam(4) = ""
    arrParam(5) = "주식보관회사코드"

    arrField(0) = "BP_CD"
    arrField(1) = "BP_NM"

    arrHeader(0) = "주식보관회사코드"
    arrHeader(1) = "주식보관회사명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=470px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
        frm1.txtCust2.focus
		Exit Function
	Else
		Call SetCust2(arrRet)
	End If
End Function

'=======================================================================================================
'	Name : SetCust1()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetCust2(Byval arrRet)
    With frm1
        .txtCust2.focus
        .txtCust2.value  = arrRet(0)
        .txtCust2Nm.value    = arrRet(1)
    End With
    lgBlnFlgChgValue = True 
End Function

Function BtnPopupDisabled(Inobj) 
	If UCase(Inobj.className) = UCase("protected") Then 
		IsOpenPop = False
		BtnPopupDisabled = False
	Else
		BtnPopupDisabled = True
	End If
End Function

'=======================================================================================================
'	Name : OpenAcctPopup()
'	Description : Major PopUp
'=======================================================================================================
Function OpenAcctPopup(iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	Select Case iwhere
		Case 0
			arrParam(0) = "이자수익계정"
			arrParam(1) = "A_ACCT"
			arrParam(2) = Trim(frm1.txtAcct1.value)
			arrParam(3) = ""
			arrParam(4) = "DEL_FG <> " & FilterVar("Y", "''", "S") & " "
			arrParam(5) = "미수수익계정"

			arrField(0) = "ACCT_CD"
			arrField(1) = "ACCT_NM"
			
			arrHeader(0) = "미수수익계정코드"
			arrHeader(1) = "미수수익계정명"
		Case 1
			arrParam(0) = "이자수익계정"
			arrParam(1) = "A_ACCT"
			arrParam(2) = Trim(frm1.txtAcct2.value)
			arrParam(3) = ""
			arrParam(4) = "DEL_FG <> " & FilterVar("Y", "''", "S") & " "
			arrParam(5) = "이자수익계정"

			arrField(0) = "ACCT_CD"
			arrField(1) = "ACCT_NM"

			arrHeader(0) = "이자수익계정코드"
			arrHeader(1) = "이자수익계정명"
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Select Case iwhere
			Case 0
				frm1.txtAcct1.focus
			Case 1
				frm1.txtAcct2.focus
		End Select
		Exit Function
	Else
		Call SetAcctPopup(arrRet, iWhere)
	End If	
End Function

'=======================================================================================================
'	Name : SetAcctPopup()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetAcctPopup(Byval arrRet, Byval iwhere)
	With frm1
		Select Case iwhere
			Case 0
				.txtAcct1.focus
				.txtAcct1.value = arrRet(0)
				.txtAcctNm1.value = arrRet(1)
			Case 1
				.txtAcct2.focus
				.txtAcct2.value = arrRet(0)
				.txtAcctNm2.value = arrRet(1)
		End Select
	End With
	
	lgBlnFlgChgValue = True	
End Function

'========================================================================================================
' Name : txtInsureAcct_Onchange
' Desc : developer describe this line
'========================================================================================================
Function txtAcct1_Onchange()
    Dim IntRetCd
	
	If  frm1.txtAcct1.value = "" Then
		frm1.txtAcct1.value = ""
		frm1.txtAcctNM1.value=""
		frm1.txtAcct1.focus
	Else
	    IntRetCD= CommonQueryRs(" ACCT_CD,ACCT_NM "," A_ACCT "," ACCT_CD = " & FilterVar(frm1.txtAcct1.value, "''", "S") & " and DEL_FG <> " & FilterVar("Y", "''", "S") & " " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 		    
			If IntRetCD=False And Trim(frm1.txtAcct1.value)<>"" Then
			    Call DisplayMsgBox("800054","X","X","X")                         '☜ : 등록되지 않은 코드입니다.
			    frm1.txtAcct1.value=""
			    frm1.txttxtAcctNm1.value=""
			    frm1.txtAcct1.focus
			    Set gActiveElement = document.activeElement  
			Else
			    frm1.txtAcct1.value=Trim(Replace(lgF0,Chr(11),""))
			    frm1.txtAcctNm1.value=Trim(Replace(lgF1,Chr(11),""))
			End If
	End If
	
	lgBlnFlgChgValue = True   
End Function 

'========================================================================================================
' Name : txtInsureAcct_Onchange
' Desc : developer describe this line
'========================================================================================================
Function txtAcct2_Onchange()
    Dim IntRetCd
	
	If  frm1.txtAcct2.value = "" Then
		frm1.txtAcct2.value = ""
		frm1.txtAcctNm2.value=""
		frm1.txtAcct2.focus
	Else
	    IntRetCD= CommonQueryRs(" ACCT_CD,ACCT_NM "," A_ACCT "," ACCT_CD = " & FilterVar(frm1.txtAcct1.value, "''", "S") & " and DEL_FG <> " & FilterVar("Y", "''", "S") & " " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 		    
			If IntRetCD=False And Trim(frm1.txtAcct2.value)<>"" Then
			    Call DisplayMsgBox("800054","X","X","X")                         '☜ : 등록되지 않은 코드입니다.
			    frm1.txtAcct2.value=""
			    frm1.txttxtAcctNm2.value=""
			    frm1.txtAcct2.focus
			    Set gActiveElement = document.activeElement  
			Else
			    frm1.txtAcct2.value=Trim(Replace(lgF0,Chr(11),""))
			    frm1.txtAcctNm2.value=Trim(Replace(lgF1,Chr(11),""))
			End If
	End If
	
	lgBlnFlgChgValue = True   
End Function 

'========================================================================================================
'   Event Name : txtBillDt_DbClick
'   Event Desc :
'========================================================================================================
Sub txtBillDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtBillDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtBillDt.Focus
	End If
	
	Call deptCheck()
	lgBlnFlgChgValue = True 
End Sub

'========================================================================================================
'   Event Name : txtPubDt_DbClick
'   Event Desc :
'========================================================================================================
Sub txtPubDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtPubDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtPubDt.Focus
	End If
	
	lgBlnFlgChgValue = True 
End Sub

'========================================================================================================
'   Event Name : txtExpireDt_DbClick
'   Event Desc :
'========================================================================================================
Sub txtExpireDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtExpireDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtExpireDt.Focus
	End If
	lgBlnFlgChgValue = True 
End Sub

'========================================================================================================
'   Event Name : txtInDt_DbClick
'   Event Desc :
'========================================================================================================
Sub txtInDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtInDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtInDt.Focus
	End If
	Call txtTradeCur_OnChange()
	lgBlnFlgChgValue = True 
End Sub

Sub txtInDt_Change()
    If frm1.txtInDt.text <> "" AND frm1.txtCust1.value <> "" Then
        Call txtTradeCur_OnChange()
    End If
End Sub

'========================================================================================================
'   Event Name : txtOutDt_DbClick
'   Event Desc :
'========================================================================================================
Sub txtOutDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtOutDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtOutDt.Focus
	End If
	lgBlnFlgChgValue = True 
End Sub
'=======================================================================================================
'   Event Name : fpDateTime_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtBillDt_Change()
    lgBlnFlgChgValue = True
End Sub
'=======================================================================================================
Sub txtPubDt_Change()
    lgBlnFlgChgValue = True
End Sub
'=======================================================================================================
Sub txtExpireDt_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
Sub txtOutDt_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
Sub txtXchRate_Change()
	Dim iRows
	Dim iVal
	
	frm1.txtLocBuyAmt.text = 0
	frm1.txtLocPriceAmt.text = 0
    lgBlnFlgChgValue = True   
End Sub

Sub txtBuyAmt_Change()
	frm1.txtLocBuyAmt.text = 0
	lgBlnFlgChgValue = True   
End Sub

Sub txtPriceAmt_Change()
    frm1.txtLocPriceAmt.text = 0
    lgBlnFlgChgValue = True   
End Sub

Sub selCalYn_Change()
    If frm1.selCalYn.value = "Y" Then
        Call ggoOper.SetReqAttr(frm1.txtCalRate,"N")
        Call ggoOper.SetReqAttr(frm1.selEndYn,"N")
        Call ggoOper.SetReqAttr(frm1.txtExpireDt,"N")
    Else
        Call ggoOper.SetReqAttr(frm1.txtCalRate,"Q")
        Call ggoOper.SetReqAttr(frm1.selEndYn,"Q")
        Call ggoOper.SetReqAttr(frm1.txtExpireDt,"Q")
    End If

	lgBlnFlgChgValue = True 
End Sub

Sub selComYn_Change()
    If frm1.selComYn.value = "Y" Then
        Call ggoOper.SetReqAttr(frm1.txtOutDt,"N")
    Else
        Call ggoOper.SetReqAttr(frm1.txtOutDt,"Q")
    End If
    
    lgBlnFlgChgValue = True 
End Sub

Sub txtSecuType_Change()
    Dim var1
    If Trim(frm1.txtSecuType.value) <> "" Then
		Call CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("A1031", "''", "S") & "  AND MINOR_CD =  " & FilterVar(frm1.txtSecuType.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		var1 = Replace(lgF0, Chr(11), "")

		If var1 = "" Then
		    Call DisplayMsgBox("970000","X",frm1.txtSecuType.alt,"X")
		    frm1.txtSecuType.value = ""
		    frm1.txtSecuTypeNm.value = ""
		    frm1.txtSecuType.focus
		    Set gActiveElement = document.activeElement
		Else
		    frm1.txtSecuTypeNm.value = var1
		End If
	Else
		frm1.txtSecuType.value = ""
		frm1.txtSecuTypeNm.value = ""
	End If	
    
    lgBlnFlgChgValue = True 
End Sub

'========================================================================================================
' Name : txtDept1_Onchange
' Desc : developer describe this line
'========================================================================================================
Function txtDept1_Onchange()
    Dim strSelect
	Dim strFrom
	Dim strWhere 	
    Dim IntRetCD 
	Dim arrVal1
	Dim arrVal2
	Dim ii
	Dim jj
	
	With Frm1
		If .txtDept1.value = "" Then
			.txtDept1.value = ""
			.txtDept1Nm.value=""
			.txtDept2.value=""
			.txtDept2Nm.value=""
			.txtDept1Area.value=""
			.txtDept1OrgId.value = ""
			Call ggoOper.SetReqAttr(.txtDept2, "Q")
			.txtDept1.focus
			Set gActiveElement = document.activeElement
			lgBlnFlgChgValue = True 
			Exit Function
		End If
    
		If Trim(.txtBillDt.Text = "") Then    
			Exit Function
		End If
		
		lgBlnFlgChgValue = True

		strSelect	=			 " a.dept_cd,a.dept_nm, a.org_change_id, a.internal_cd, b.biz_area_cd "    		
		strFrom		=			 " b_acct_dept a, b_cost_center b "		
		strWhere	= " a.cost_cd = b.cost_cd " 	 
		strWhere	= strWhere & " and a.dept_Cd = " & FilterVar(LTrim(RTrim(.txtDept1.value)), "''", "S")
		strWhere	= strWhere & " and a.org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(.txtBillDt.Text, parent.gDateFormat,""), "''", "S") & "))"			
		
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			.txtDept1.value = ""
			.txtDept1Nm.value = ""
			.txtDept1Area.value = ""
			.txtDept1OrgId.value = ""
			.txtInternalCd1.value = ""
			.txtInternalCd2.value = ""
			.txtDept2.value = ""
			.txtDept2Nm.value = ""
			Call ggoOper.SetReqAttr(.txtDept2, "Q")
			.txtDept1.focus
			Set gActiveElement = document.activeElement  
		Else 
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)
						
			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))			
				.txtDept1Nm.value=Trim(arrVal2(2))
			    .txtDept1OrgId.value =Trim(arrVal2(3))
			    .txtInternalCd1.value =Trim(arrVal2(4))
			    .txtDept1Area.value =Trim(arrVal2(5))
			    .txtDept2.value=""
			    .txtDept2Nm.value=""
			    .txtDept2.focus
			    Call ggoOper.SetReqAttr(.txtDept2, "N")
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
			.txtDept2.value=""
			.txtDept2Nm.value=""
			.txtDept2.focus
			lgBlnFlgChgValue = True 
			Exit Function
		End If
    
		If Trim(.txtBillDt.Text = "") Then    
			Exit Function
		End If
		
		lgBlnFlgChgValue = True

		strSelect	=			 " a.dept_cd,a.dept_nm "    		
		strFrom		=			 " b_acct_dept a, b_cost_center b "		
		strWhere	= " a.cost_cd = b.cost_cd " 	 
		strWhere	= strWhere & " and a.dept_Cd = " & FilterVar(LTrim(RTrim(.txtDept2.value)), "''", "S")
		strWhere	= strWhere & " and b.biz_area_cd = " & FilterVar(LTrim(RTrim(.txtDept1Area.value)), "''", "S")
		strWhere	= strWhere & " and a.org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(.txtBillDt.Text, parent.gDateFormat,""), "''", "S") & "))"			
			
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			.txtDept2.value = ""
			.txtDept2Nm.value = ""
			.txtDept2.focus
			Set gActiveElement = document.activeElement  
		Else 
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)
						
			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))			
				.txtDept2.value=Trim(arrVal2(1))
			    .txtDept2Nm.value=Trim(arrVal2(2))
			    .txtCust2.focus
			Next	
		End If
	End With

	lgBlnFlgChgValue = True   
End Function 

Sub txtCust1_Change()
    Dim var1, var2

	With frm1
		If Trim(.txtCust1.value) <> "" Then
			Call CommonQueryRs(" BP_NM, CURRENCY "," B_BIZ_PARTNER "," BP_CD =  " & FilterVar(.txtCust1.value , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			var1 = Replace(lgF0, Chr(11), "")
			var2 = Replace(lgF1, Chr(11), "")

			If var1 = "" Then
			    Call DisplayMsgBox("126100","X","X","X")
			    .txtCust1.value = ""
			    .txtCust1Nm.value = ""
			    .txtCust1.focus
			    Set gActiveElement = document.activeElement
			Else
			    .txtCust1Nm.value = var1
'			    .txtTradeCur.value = Trim(var2)
'			    Call txtTradeCur_OnChange()
			End If
		Else
			.txtCust1.value = ""
			.txtCust1Nm.value = ""	
		End If	
	End With
	
    lgBlnFlgChgValue = True 
End Sub

Sub txtCust2_Change()
    Dim var1
    
    With frm1
		If Trim(.txtCust2.value) <> "" Then
			Call CommonQueryRs(" BP_NM "," B_BIZ_PARTNER "," BP_CD =  " & FilterVar(.txtCust2.value , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			var1 = Replace(lgF0, Chr(11), "")

			If var1 = "" Then
			    Call DisplayMsgBox("126100","X","X","X")
			    .txtCust2.value   = ""
			    .txtCust2Nm.value = ""
			    .txtCust2.focus
			    Set gActiveElement = document.activeElement
			Else
			    .txtCust2Nm.value = var1
			End If
		Else
			.txtCust2.value   = ""
			.txtCust2Nm.value = ""	
		End If	
	End With

    lgBlnFlgChgValue = True 
End Sub

Sub txtTradeCur_OnChange()
    Dim var1, var2
	Dim FrDt
	Dim IntRetCD, strSelect, strFrom, strWhere
	
	If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtTradeCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
		Call CurFormatNumericOCX()
	End If	    
	
	If Trim(frm1.txtInDt.Text) = "" Then
		frm1.txtInDt.Text = UniConvDateAToB("<%=StartDate%>",parent.gServerDateFormat,parent.gDateFormat)
		Exit Sub
	End If

	FrDt = UniConvDateToYYYYMMDD(frm1.txtInDt.Text,parent.gDateFormat,"")   '//parent.UNIConvDate(frm1.txtInDt.Text)
	
	If Trim(frm1.txtTradeCur.value) <> "" Then
	    Call CommonQueryRs(" CURRENCY_DESC "," B_CURRENCY "," CURRENCY =  " & FilterVar(frm1.txtTradeCur.value , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	    var1 = Replace(lgF0, Chr(11), "")

	    If var1 = "" Then
	        frm1.txtXchRate.text = 1
	        Set gActiveElement = document.activeElement
	    Else
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
	       
			If IntRetCD = True Then
				var2 = Trim(Replace(lgF0,Chr(11),""))
		    End If
		    
	        frm1.txtTradeCurNm.value = Trim(var1)
	        
	        If Trim(var2) = "" Then
	            var2 = 1
	        End If
	        
	        frm1.txtXchRate.text = UNIConvNumPCToCompanyByCurrency(Trim(var2), frm1.txtTradeCur.value,parent.ggExchRateNo, "X", "X")
	    End If
	Else
		frm1.txtTradeCur.value = ""
		frm1.txtTradeCurNm.value = ""
	End If
	
    lgBlnFlgChgValue = True 
End Sub

Function txtBillDt_onblur()
	Call deptCheck()
End Function


Function txtInDt_onblur()
	Call txtTradeCur_OnChange()
End Function

'//////////////////////////check 작업중/////////////////////////////
Function deptCheck()
	Dim strSelect
	Dim strFrom
	Dim strWhere 	
    Dim IntRetCD 
	
	With Frm1
		If .txtDept1.value = "" Then
			.txtDept1.value = ""
			.txtDept1Nm.value = ""
			.txtDept1Area.value = ""
			.txtDept1OrgId.value = ""
			.txtInternalCd1.value = ""
			.txtInternalCd2.value = ""
			.txtDept2.value = ""
			.txtDept2Nm.value = ""
			Call ggoOper.SetReqAttr(.txtDept2, "Q")
			.txtDept1.focus
			Set gActiveElement = document.activeElement
			lgBlnFlgChgValue = True 
			Exit Function
		End If
    
		If Trim(.txtBillDt.Text = "") Then    
			Exit Function
		End If

		strSelect	=			 " distinct org_change_id "    		
		strFrom		=			 " b_acct_dept "		
		strWhere	=			 " org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(.txtBillDt.Text, parent.gDateFormat,""), "''", "S") & "))"			
			
		IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 			
	
		If IntRetCD = False or Trim(Replace(lgF0,Chr(11),"")) <> .txtDept1OrgId.value  Then
			.txtDept1.value = ""
			.txtDept1Nm.value = ""
			.txtDept1Area.value = ""
			.txtDept1OrgId.value = ""
			.txtInternalCd1.value = ""
			.txtInternalCd2.value = ""
			.txtDept2.value = ""
			.txtDept2Nm.value = ""
			Call ggoOper.SetReqAttr(.txtDept2, "Q")
			lgBlnFlgChgValue = True
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

	Err.Clear                                                                    '☜: Clear err status
	   	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X") '☜ "데이타가 변경되었습니다. 계속 하시겠습니까?"
		If IntRetCD = vbNo Then
			Exit Sub
		End If
	End If	
	
	If   LayerShowHide(1) = False Then
	     Exit Sub
	End If

	Call ExtractDateFrom(frm1.txtBillDt,parent.gDateFormat,parent.gComDateType,strYear,strMonth,strDay)
	txtGlDt = strYear & strMonth & strDay
    
    strVal = BIZ_PGM_ID1 & "?txtMode="          & Gubun                       '☜: Query
    strVal = strVal     & "&txtSecuCode1="      & Trim(frm1.txtSecuCode1.value)
    strVal = strVal     & "&txtGlDt="			& Trim(txtGlDt)
    
	Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
    Set gActiveElement = document.ActiveElement   
End Sub


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
</HEAD>

<BODY SCROLL="No" TABINDEX="-1">
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../image/table/seltab_up_bg.gif"><img src="../../image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>유가증권정보 기초등록</font></td>
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
									<TD CLASS=TD5 NOWRAP>유가증권</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSecuCode" SIZE="20" MAXLENGTH="20" TAG="12XXXU" ALT="유가증권 코드" ><IMG SRC="../../image/btnPopup.gif" NAME="btnSecuCode" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:btnSecuCodeOnClick()">&nbsp;<INPUT NAME="txtSecuNm" TYPE=TEXT SIZE="30" MAXLENGTH="30"   TAG="24XXXU" ALT="증권명칭"></TD>
                                    <TD CLASS=TDT NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
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
								<TD CLASS=TD5 NOWRAP>증권코드</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSecuCode1" TYPE=TEXT SIZE="20" MAXLENGTH="20"   TAG="23XXXU" ALT="증권코드"></TD>
								<TD CLASS=TD5 NOWRAP>증권명칭</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSecuNm1" TYPE=TEXT SIZE="30" MAXLENGTH="30"   TAG="23XXXX" ALT="증권명칭"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>증권종류</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSecuType" TYPE=TEXT SIZE=10  MAXLENGTH="20" TAG="23XXXU" ALT="증권종류" OnChange="txtSecuType_Change()"><IMG SRC="../../image/btnPopup.gif" NAME="btnSecuType" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnSecuTypeOnClick(frm1.txtSecuType)">&nbsp;<INPUT TYPE=TEXT NAME="txtSecuTypeNm"  SIZE="20" MAXLENGTH="50" TAG="24"></TD>
								<TD CLASS=TD5 NOWRAP>전표일자</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5970ma1_txtBillDt_txtBillDt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>발행거래처</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCust1" TYPE=TEXT SIZE=10  MAXLENGTH="10" TAG="23XXXU" ALT="발행거래처" OnChange="txtCust1_Change()"><IMG SRC="../../image/btnPopup.gif" NAME="btnCust1" align=top TYPE="BUTTON" ONCLICK ="vbscript:Call OpenBp(frm1.txtCust1.value, 1)">&nbsp;<INPUT TYPE=TEXT NAME="txtCust1Nm"  SIZE="20" MAXLENGTH="50" TAG="24"></TD>
								<TD CLASS=TD5 NOWRAP>매수</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5970ma1_txtCnt_txtCnt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>거래통화</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTradeCur" SIZE="10"  MAXLENGTH="3" TAG="23XXXU" ALT="거래통화" ><IMG SRC="../../image/btnPopup.gif" NAME="btnTradeCur" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnTradeCurOnClick(frm1.txtTradeCur)">&nbsp;<INPUT TYPE=TEXT NAME="txtTradeCurNm" SIZE="20" MAXLENGTH="50" TAG="24"></TD>
								<TD CLASS=TD5 NOWRAP>환율</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5970ma1_txtXchRate_txtXchRate.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>취득금액</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5970ma1_txtBuyAmt_txtBuyAmt.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>취득금액(자국)</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5970ma1_txtLocBuyAmt_txtLocBuyAmt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>액면금액</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5970ma1_txtPriceAmt_txtPriceAmt.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>액면금액(자국)</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5970ma1_txtLocPriceAmt_txtLocPriceAmt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>발의부서</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDept1" TYPE=TEXT SIZE=10  MAXLENGTH="10" TAG="23XXXU" ALT="발의부서"><IMG SRC="../../image/btnPopup.gif" NAME="btnDept1" align=top TYPE="BUTTON" ONCLICK ="vbscript:Call OpenDept(frm1.txtDept1.value,1)">&nbsp;<INPUT TYPE=TEXT NAME="txtDept1Nm"  SIZE="20" MAXLENGTH="50" TAG="24"></TD>
								<TD CLASS=TD5 NOWRAP>귀속부서</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDept2" TYPE=TEXT SIZE=10  MAXLENGTH="10" TAG="23XXXU" ALT="귀속부서"><IMG SRC="../../image/btnPopup.gif" NAME="btnDept2" align=top TYPE="BUTTON" ONCLICK ="vbscript:Call OpenDept(frm1.txtDept2.value,2)">&nbsp;<INPUT TYPE=TEXT NAME="txtDept2Nm"  SIZE="20" MAXLENGTH="50" TAG="24"></TD>
								
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>주식보관회사</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCust2" TYPE=TEXT SIZE=10  MAXLENGTH="20" TAG="25XXXU" ALT="주식보관회사" OnChange="txtCust2_Change()"><IMG SRC="../../image/btnPopup.gif" NAME="btnCust2" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnCust2OnClick(frm1.txtCust2)" >&nbsp;<INPUT TYPE=TEXT NAME="txtCust2Nm"  SIZE="20" MAXLENGTH="50" TAG="24"></TD>
								<TD CLASS=TD5 NOWRAP>발행일자</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5970ma1_txtPubDt_txtPubDt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>미수수익계정</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAcct1" SIZE="10"  MAXLENGTH="20" TAG="23XXXU" ALT="미수수익계정"><IMG SRC="../../image/btnPopup.gif" NAME="btnAcct1" align=top TYPE="BUTTON" ONCLICK ="vbscript:Call OpenAcctPopup(0)">&nbsp;<INPUT TYPE=TEXT NAME="txtAcctNm1" SIZE="20" MAXLENGTH="50" TAG="24"></TD>
								<TD CLASS=TD5 NOWRAP>이자수익계정</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAcct2" SIZE="10"  MAXLENGTH="20" TAG="23XXXU" ALT="이자수익계정"><IMG SRC="../../image/btnPopup.gif" NAME="btnAcct2" align=top TYPE="BUTTON" ONCLICK ="vbscript:Call OpenAcctPopup(1)">&nbsp;<INPUT TYPE=TEXT NAME="txtAcctNm2" SIZE="20" MAXLENGTH="50" TAG="24"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>만기여부</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="selComYn" TAG="23XXXU" ALT="완료여부" OnChange="selComYn_Change()"><OPTION VALUE="Y">Y</OPTION><OPTION VALUE="N">N</OPTION></SELECT></TD>
								<TD CLASS=TD5 NOWRAP>이자율</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5970ma1_txtCalRate_txtCalRate.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>이자계산</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="selCalYn" TAG="23XXXU" ALT="이자계산" OnChange="selCalYn_Change()"><OPTION VALUE="Y">계산</OPTION><OPTION VALUE="N">미계산</OPTION></SELECT></TD>
								<TD CLASS=TD5 NOWRAP>만기일자</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5970ma1_txtExpireDt_txtExpireDt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>양편구분</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="selEndYn" TAG="23XXXU" ALT="양편구분"></SELECT></TD>
								<TD CLASS=TD5 NOWRAP>취득일자</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5970ma1_txtInDt_txtInDt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>관리번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRefNo" TYPE=TEXT SIZE="20" MAXLENGTH="20"   TAG="25XXXU" ALT="관리번호"></TD>
								<TD CLASS=TD5 NOWRAP>처분일자</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5970ma1_txtOutDt_txtOutDt.js'></script></TD>
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
		<TD WIDTH=100% HEIGHT=<%=BIZSIZE%>> <IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"    TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtPrevNext" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtDept1Area" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtDept1OrgId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInternalCd1" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInternalCd2" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm" TABINDEX="-1"></IFRAME>
</DIV>
</BODY>
</HTML>
