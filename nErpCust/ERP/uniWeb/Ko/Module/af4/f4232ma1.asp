<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : LOAN TERM CHANGE
*  2. Function Name        : F4232MA1
*  3. Program ID           : F4232MA1
*  4. Program Name         : 유동성 전환 
*  5. Program Desc         : 
*  6. Comproxy List        :
*  7. Modified date(First) : 2002/03/30
*  8. Modified date(Last)  : 2002/05/19
*  9. Modifier (First)     : JANG YOON KI
* 10. Modifier (Last)      : Ahn do hyun
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->
<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->


<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../ag/AcctCtrl.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance
'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID = "F4232MB1.asp"                                      'Biz Logic ASP 
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Dim C_Choice
Dim C_CHG_DT
Dim C_Loan_No
Dim C_Loan_Nm
Dim C_PLAN_ACCT_CD
Dim C_PLAN_ACCT_BT
Dim C_PLAN_ACCT_NM
Dim C_LOAN_ACCT_CD
Dim C_LOAN_ACCT_NM
Dim C_Pay_Plan_Dt
Dim C_Doc_Cur
Dim C_Xch_Rate
Dim C_PLAN_AMT
Dim C_PLAN_LOC_AMT
Dim C_Loan_Dt
Dim C_Due_Dt
Dim C_Loan_Int_Rate
Dim C_PAY_OBJ

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop          
Dim BaseDate,LastDate
Dim strSvrDate
Dim lstxtPlanAmtSum
Dim StartDate, EndDate
Dim lgIsOpenPop

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

    StartDate	= "<%=GetSvrDate%>"                                               'Get Server DB Date
	EndDate		= StartDate

'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = Parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgStrPrevKeyIndex = ""                                      '⊙: initializes Previous Key Index
    lgSortKey         = 1                                       '⊙: initializes sort direction
		
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
	C_Choice			=  1	'선택 
	C_CHG_DT			=  2	'전환일자 
	C_Loan_No			=  3	'차입금번호 
	C_Loan_Nm			=  4	'차입내역 
	C_PLAN_ACCT_CD		=  5	'유동성계정 
	C_PLAN_ACCT_BT		=  6	'유동성계정 
	C_PLAN_ACCT_NM		=  7
	C_LOAN_ACCT_CD		=  8	'차입금계정 
	C_LOAN_ACCT_NM		=  9
	C_Pay_Plan_Dt		= 10	'상환예정일자 
	C_Doc_Cur			= 11 	'통화 
	C_Xch_Rate			= 12	'환율 
	C_PLAN_AMT			= 13	'차입잔액 
	C_PLAN_LOC_AMT		= 14	'차입잔액(자국)
	C_Loan_Dt			= 15	'차입일자 
	C_Due_Dt			= 16	'만기일자 
	C_Loan_Int_Rate		= 17	'이자율 
	C_PAY_OBJ			= 18	'장단기 구분 


End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	DIm strYear, strMonth, strDay
	Dim frDt, toDt

	strSvrDate = "<%=GetSvrDate%>"
	Call ExtractDateFrom(strSvrDate, Parent.gServerDateFormat, Parent.gServerDateType, strYear,strMonth,strDay)
		
	frDt = UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, strDay)	
	toDt = UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear+1, "12", "31")
   
	frm1.txtDateFr.Text = frDt   	
	frm1.txtDateTo.Text = toDt
	frm1.txtChgDt.Text  = frDt
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
' Para : 1. Currency
'        2. I(Input),Q(Query),P(Print),B(Bacth)
'        3. "*" is for Common module
'           "A" is for Accounting
'           "I" is for Inventory
'           ...
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
End Sub

'========================================================================================================
' Name : CookiePage()
' Desc : Write or Read cookie value 
'========================================================================================================
Sub CookiePage(Kubun)
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : MakeKeyStream
' Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
  
	Dim strYYYYMMDD_fr
	Dim strYYYYMMDD_to
	Dim strYYYYMMDD_Chg
	Dim strFCAcctNm

	strYYYYMMDD_fr  = UNIConvDate(frm1.txtDateFr.text)
	strYYYYMMDD_to  = UNIConvDate(frm1.txtDateTo.text)
	strYYYYMMDD_Chg = UNIConvDate(frm1.txtChgDt.text)

	lgKeyStream = strYYYYMMDD_fr & Parent.gColSep       'You Must append one character(gColSep)
	lgKeyStream = lgKeyStream & strYYYYMMDD_to & Parent.gColSep       'You Must append one character(gColSep)
	lgKeyStream = lgKeyStream & strYYYYMMDD_Chg & Parent.gColSep       'You Must append one character(gColSep)
	lgKeyStream = lgKeyStream & Parent.gColSep
	lgKeyStream = lgKeyStream & UCase(Trim(frm1.txtLoan_NO.value)) & Parent.gColSep 
	lgKeyStream = lgKeyStream & UCase(Trim(frm1.txtFCAcctCd.value)) & Parent.gColSep
	lgKeyStream = lgKeyStream & Parent.gColSep		'frm1.txtFCAcctNm.value
	lgKeyStream = lgKeyStream & frm1.cboLoanFg.value & Parent.gColSep
	lgKeyStream = lgKeyStream & frm1.txtBizAreaCd.value & Parent.gColSep
	lgKeyStream = lgKeyStream & frm1.txtBizAreaCd1.value & Parent.gColSep		
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        

	
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1020", "''", "S") & "  and MINOR_CD IN (" & FilterVar("LL", "''", "S") & " ," & FilterVar("LN", "''", "S") & " ) ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboLoanFg ,lgF0  ,lgF1  ,Chr(11))
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset ComboBox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex
	
	ggoSpread.Source = frm1.vspdData
	With frm1.vspdData
		For intRow = 1 to .MaxRows
			.Col = C_Choice
			.Row = intRow	:	.text= "1"
		Next
	End With
	
End Sub


'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()
    ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20021207",,parent.gAllowDragDropSpread    

	With frm1.vspdData
	
       .MaxCols   = C_PAY_OBJ + 1                                                  ' ☜:☜: Add 1 to Maxcols
       .MaxRows = 0                                                                  ' ☜: Clear spreadsheet data 

	   .ReDraw = false
		
       Call GetSpreadColumnPos("A")

								'ColumnPosition		Header					Width  Align(0:L,1:R,2:C)  Row   Length  CharCase(0:L,1:N,2:U)
       ggoSpread.SSSetCheck		C_Choice			,"선택"				,8    ,2                   ,     ,15     
       ggoSpread.SSSetDate		C_CHG_DT			,"전환일자"			,10    ,2                  ,Parent.gDateFormat   ,-1
       ggoSpread.SSSetEdit		C_Loan_No			,"차입금번호"       ,13    ,0                  ,     ,18     ,2
       ggoSpread.SSSetEdit		C_Loan_Nm			,"차입내역"			,15    ,0                  ,     ,50     ,2
       ggoSpread.SSSetEdit		C_PLAN_ACCT_CD		,"유동성계정"       ,13    ,0                  ,     ,15     ,2
       ggoSpread.SSSetButton	C_PLAN_ACCT_BT
       ggoSpread.SSSetEdit		C_PLAN_ACCT_NM		,"유동성계정명"		,15    ,0                  ,     ,50     ,2
       ggoSpread.SSSetEdit		C_LOAN_ACCT_CD		,"차입금계정"       ,13    ,0                  ,     ,15     ,2
       ggoSpread.SSSetEdit		C_LOAN_ACCT_NM		,"차입금계정명"		,15    ,0                  ,     ,50     ,2
       ggoSpread.SSSetDate		C_Pay_Plan_Dt		,"상환예정일자"		,12    ,2                  ,Parent.gDateFormat   ,-1
       ggoSpread.SSSetEdit		C_Doc_Cur			,"통화"				,10    ,					,     ,15     ,2
       ggoSpread.SSSetFloat		C_Xch_Rate			,"환율"				,11    ,Parent.ggExchRateNo	,ggStrIntegeralPart	,ggStrDeciPointPart	,Parent.gComNum1000	,Parent.gComNumDec
       ggoSpread.SSSetFloat		C_PLAN_AMT			,"상환예정액"		,15    , "A",ggStrIntegeralPart,ggStrDeciPointPart	,Parent.gComNum1000	,Parent.gComNumDec
       ggoSpread.SSSetFloat		C_PLAN_LOC_AMT		,"상환예정액(자국)"	,15    ,parent.ggAmtOfMoneyNo,ggStrIntegeralPart,ggStrDeciPointPart	,Parent.gComNum1000	,Parent.gComNumDec	
       ggoSpread.SSSetDate		C_Loan_Dt			,"차입일자"			,15    ,2                  ,Parent.gDateFormat   ,-1
       ggoSpread.SSSetDate		C_Due_Dt			,"상환만기일자"		,15    ,2                  ,Parent.gDateFormat   ,-1
       ggoSpread.SSSetFloat		C_Loan_Int_Rate		,"이자율"			,8     ,Parent.ggExchRateNo	,ggStrIntegeralPart	,ggStrDeciPointPart	,Parent.gComNum1000	,Parent.gComNumDec
       ggoSpread.SSSetEdit		C_PAY_OBJ			,"상환번호"			,15    ,0                  ,     ,15     ,2          

	   .ReDraw = true

       Call ggoSpread.MakePairsColumn(C_PLAN_ACCT_CD,C_PLAN_ACCT_BT)
       
	   Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
       Call ggoSpread.SSSetColHidden(C_Loan_Int_Rate ,C_Loan_Int_Rate	,True)
       Call ggoSpread.SSSetColHidden(C_PAY_OBJ ,C_PAY_OBJ	,True)

       Call SetSpreadLock 
    
    End With
    
End Sub

'======================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect cell in spread sheet
'======================================================================================================
Sub SetSpreadLock()
   
	ggoSpread.Source = frm1.vspdData

	With frm1
		.vspdData.ReDraw = False 
                                 'Col-1          Row-1
		ggoSpread.SpreadLock    C_Loan_No			, -1 
		ggoSpread.SpreadUnLock	C_PLAN_ACCT_CD		, -1
		ggoSpread.SpreadLock	C_PLAN_ACCT_NM		, -1
		ggoSpread.SSSetRequired	C_PLAN_ACCT_CD		, -1,	-1
	    .vspdData.ReDraw = True

    End With    
   
End Sub

'======================================================================================================
' Name : SetSpreadColor
' Desc : This method set color and protect cell by cell in spread sheet
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
End Sub

'======================================================================================================
' Name : SubSetErrPos
' Desc : This method set focus to position of error
'      : This method is called in MB area
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,Parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> Parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub


'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
       
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_Choice				= iCurColumnPos(1)
			C_CHG_DT				= iCurColumnPos(2)
			C_Loan_No				= iCurColumnPos(3)    
			C_Loan_Nm				= iCurColumnPos(4)
			C_PLAN_ACCT_CD			= iCurColumnPos(5)
			C_PLAN_ACCT_BT			= iCurColumnPos(6)
			C_PLAN_ACCT_NM			= iCurColumnPos(7)
			C_LOAN_ACCT_CD			= iCurColumnPos(8)
			C_LOAN_ACCT_NM			= iCurColumnPos(9)
			C_Pay_Plan_Dt			= iCurColumnPos(10)
			C_Doc_Cur				= iCurColumnPos(11)
			C_Xch_Rate				= iCurColumnPos(12)
			C_PLAN_AMT				= iCurColumnPos(13)
			C_PLAN_LOC_AMT			= iCurColumnPos(14)
			C_Loan_Dt			    = iCurColumnPos(15)
			C_Due_Dt				= iCurColumnPos(16)
			C_Loan_Int_Rate			= iCurColumnPos(17)
			C_PAY_OBJ				= iCurColumnPos(18)

    End Select    
End Sub

'----------------------------------------  OpenBizAreaCd()  -------------------------------------------------
'	Name : BizAreaCd()
'	Description : Business PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizAreaCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "사업장 팝업"				' 팝업 명칭 
	arrParam(1) = "B_BIZ_AREA"					' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtBizAreaCd.Value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
	' 권한관리 추가 
	If lgAuthBizAreaCd <> "" Then
		arrParam(4) = " BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
	Else
		arrParam(4) = ""
	End If
	arrParam(5) = "사업장 코드"			

    arrField(0) = "BIZ_AREA_CD"					' Field명(0)
    arrField(1) = "BIZ_AREA_NM"					' Field명(1)

    arrHeader(0) = "사업장코드"				' Header명(0)
	arrHeader(1) = "사업장명"				' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
'		frm1.txtBizAreaCd.focus
		Exit Function
	Else
		frm1.txtBizAreaCd.Value	= arrRet(0)
		frm1.txtBizAreaNm.Value	= arrRet(1)
	End If
End Function


'----------------------------------------  OpenBizAreaCd()  -------------------------------------------------
'	Name : BizAreaCd()
'	Description : Business PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizAreaCd1()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "사업장 팝업"				' 팝업 명칭 
	arrParam(1) = "B_BIZ_AREA"					' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtBizAreaCd1.Value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
	' 권한관리 추가 
	If lgAuthBizAreaCd <> "" Then
		arrParam(4) = " BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
	Else
		arrParam(4) = ""
	End If
	arrParam(5) = "사업장 코드"			

    arrField(0) = "BIZ_AREA_CD"					' Field명(0)
    arrField(1) = "BIZ_AREA_NM"					' Field명(1)

    arrHeader(0) = "사업장코드"				' Header명(0)
	arrHeader(1) = "사업장명"				' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
'		frm1.txtBizAreaCd.focus
		Exit Function
	Else
		frm1.txtBizAreaCd1.Value = arrRet(0)
		frm1.txtBizAreaNm1.Value = arrRet(1)
	End If
End Function

'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : Form_Load
' Desc : This sub is called from window_OnLoad in Common.vbs automatically
'========================================================================================================
Sub Form_Load()
    Err.Clear                                                                        '☜: Clear err status
    
	Call LoadInfTB19029                                                              '☜: Load table , B_numeric_format
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											 '⊙: Lock Field

    Call InitSpreadSheet                                                             'Setup the Spread sheet

	Call InitVariables
    Call SetDefaultVal
    Call InitComboBox

	frm1.txtDateFr.focus
	Call SetToolbar("1100100000011111")                                              '☆: Developer must customize
	
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
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub
	
'========================================================================================================
' Name : Form_QueryUnload
' Desc : This sub is called from window_Unonload in Common.vbs automatically
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	
	If CompareDateByFormat(frm1.txtDateFr.text,frm1.txtDateTo.text,frm1.txtDateFr.Alt,frm1.txtDateTo.Alt, _
        	               "970025",frm1.txtDateFr.UserDefinedFormat,parent.gComDateType, true) = False Then	   
	   frm1.txtDateFr.focus
	   Set gActiveElement = document.ActiveElement
	   
		Exit Function
	End If
	
    ggoSpread.Source = Frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
    															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

    Call InitVariables                                                               '⊙: Initializes local global variables
    frm1.txtPlanAmtSum.Text = 0
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    If DbQuery("MQ") = False Then                                                      '☜: Query db data
       Exit Function
    End If

    Set gActiveElement = document.ActiveElement   
    FncQuery = True                                                              '☜: Processing is OK
End Function
	
'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew()

    Dim IntRetCD 
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'In Multi, You need not to implement this area
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncNew = True																 '☜: Processing is OK
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'In Multi, You need not to implement this area
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement       
    FncDelete = True                                                             '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD
    Dim strPayObj
    Dim lRow
    
    FncSave = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then                                      '☜:match pointer
        IntRetCD = DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
       Exit Function
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	With Frm1.vspdData
		For lRow = 1 To .MaxRows
			.Row = lRow
			.Col = C_PAY_OBJ
			strPayObj = Trim(.Text)
			.Col = C_Choice
			If .Text = "1" Then        
				.Col = C_PLAN_ACCT_CD
				If Trim(.Text) <> "" Then
					If CommonQueryRs("A.ACCT_CD,A.DEL_FG","A_Acct A, A_ACCT_GP B, A_JNL_ACCT_ASSN C","A.GP_CD=B.GP_CD AND A.Acct_cd=C.Acct_CD and C.trans_type = " & FilterVar("FI003", "''", "S") & "  and C.jnl_cd = 'C" & Right(strPayObj,1) & _
							"' AND A.ACCT_CD = " & FilterVar(.Text, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
						Call DisplayMsgBox("141167","x","x","x")
						Exit Function
					Else
						If lgF1 = "Y" & chr(11) Then
							Call DisplayMsgBox("110104","x","x","x")
						End If
					End If
				End If
			End If
		Next
	End With

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If DbSave = False Then                                                       '☜: Query db data
       Exit Function
    End If

    Set gActiveElement = document.ActiveElement   
    FncSave = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
    
    ggoSpread.Source = Frm1.vspdData
	With Frm1.VspdData
         .ReDraw = False
		 If .ActiveRow > 0 Then
            ggoSpread.CopyRow
			SetSpreadColor .ActiveRow
    
            .ReDraw = True
		    .Focus
		 End If
	End With
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	' Clear key field
	'---------------------------------------------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCopy = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
    Dim iDx
    FncCancel = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = Frm1.vspdData	
    ggoSpread.EditUndo  
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,C_DOC_CUR,C_PLAN_AMT,   "A" ,"I","X","X")
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCancel = False                                                            '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    
End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()

End Function

'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
'========================================================================================================
Function FncPrint()
    FncPrint = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()                                                       '☜: Protect system from crashing
    FncPrint = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncPrev
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev() 
    FncPrev = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   	
    FncPrev = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncNext
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncNext() 
    FncNext = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   	
    FncNext = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
    FncExcel = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncExport(parent.C_MULTI)

    FncExcel = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncFind(parent.C_MULTI, True)

    FncFind = True                                                               '☜: Processing is OK
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
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
'    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,-1,-1,C_DOC_CUR,C_PLAN_AMT,"A" ,"I","X","X")
End Sub

'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

    FncExit = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")		             '⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
'========================================================================================================
'                        5.3 Common Method-3
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : DbQuery
' Desc : This sub is called by FncQuery
'========================================================================================================
Function DbQuery(pDirect)

	Dim strVal
	
    Err.Clear                                                                    '☜: Clear err status
    On Error Resume Next
    
    DbQuery = False                                                              '☜: Processing is NG

    Call DisableToolBar(parent.TBC_QUERY)                                               '☜: Disable Query Button Of ToolBar
    Call LayerShowHide(1)                                                        '☜: Show Processing Message

    Call MakeKeyStream(pDirect)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream               '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex         '☜: Next key tag
		' 권한관리 추가 
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
		strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인        
        
    End With
	
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is OK

    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
' Name : DbSave
' Desc : This sub is called by FncSave
'========================================================================================================
Function DbSave()
		
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel, iColSep, iRowSep

    On Error Resume Next
    DbSave = False                                                               '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    Call DisableToolBar(parent.TBC_SAVE)                                                '☜: Disable Save Button Of ToolBar
    Call LayerShowHide(1)                                                        '☜: Show Processing Message
		
    Frm1.txtMode.value        = Parent.UID_M0002                                        '☜: Delete
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	
    ggoSpread.Source = frm1.vspdData

    strVal = ""
    strDel = ""
    lGrpCnt = 1
	iColSep = Parent.gColSep
	iRowSep = Parent.gRowSep

	With Frm1
		For lRow = 1 To .vspdData.MaxRows
			
			
			.vspdData.Row = lRow
			.vspdData.Col = C_Choice

			If .vspdData.Text = "1" Then        
													  strVal = strVal & "C"								& iColSep
													  strVal = strVal & lRow							& iColSep                                                    
				.vspdData.Col = C_CHG_DT			: strVal = strVal & UNIConvDate(Trim(.vspdData.Text))     & iColSep	' 2 전환날짜 
				.vspdData.Col = C_LOAN_NO			: strVal = strVal & Trim(.vspdData.Text)			& iColSep	' 3 차입금번호 
				.vspdData.Col = C_PLAN_ACCT_CD		: strVal = strVal & Trim(.vspdData.Text)			& iColSep	' 4 유동성계정 
				.vspdData.Col = C_LOAN_ACCT_CD		: strVal = strVal & Trim(.vspdData.Text)			& iColSep	' 5 차입금계정 
				.vspdData.Col = C_Pay_Plan_Dt		: strVal = strVal & UNIConvDate(Trim(.vspdData.Text))     & iColSep  ' 6 상환예정일자(이자는 지급, 원금은 상환)
				.vspdData.Col = C_Doc_Cur			: strVal = strVal & Trim(.vspdData.Text)			& iColSep  ' 7 통화 
				.vspdData.Col = C_XCH_RATE			: strVal = strVal & UNICdbl(Trim(.vspdData.Text))     & iColSep  ' 8 환율 
				.vspdData.Col = C_PLAN_AMT			: strVal = strVal & UNICdbl(Trim(.vspdData.Text))     & iColSep  ' 9 변환금액 
				.vspdData.Col = C_PLAN_LOC_AMT		: strVal = strVal & UNICdbl(Trim(.vspdData.Text))     & iColSep  '10 변환금액(자국)
				.vspdData.Col = C_DUE_DT			: strVal = strVal & UNIConvDate(Trim(.vspdData.Text))     & iColSep  '11 만기일				
				.vspdData.Col = C_Loan_Int_Rate		: strVal = strVal & UNICdbl(Trim(.vspdData.Text))     & iColSep  '12 이자율 
				.vspdData.Col = C_PAY_OBJ			: strVal = strVal & Trim(.vspdData.Text)			& iRowSep  '13 장단기 구분 
				
				lGrpCnt = lGrpCnt + 1
            
			End If
		Next

		.txtMaxRows.value     = lGrpCnt-1	
		.txtSpread.value      = strDel & strVal
	   
		'권한관리추가 start
		.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
		.txthInternalCd.value =  lgInternalCd
		.txthSubInternalCd.value = lgSubInternalCd
		.txthAuthUsrID.value = lgAuthUsrID		
		'권한관리추가 end	   

	End With
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    DbSave = True                                                                '☜: Processing is OK

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : DbDelete
' Desc : This sub is called by FncDelete
'========================================================================================================
Function DbDelete()
    Err.Clear                                                                    '☜: Clear err status
    DbDelete = False                                                             '☜: Processing is NG
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	'In Multi, You need not to implement this area

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    DbDelete = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Sub DbQueryOk()
	
	Call SetSpreadLock 
    lgIntFlgMode = Parent.OPMD_UMODE    
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	If frm1.vspdData.MaxRows > 0 Then
	    Call InitData()
		frm1.vspdData.focus
		Call SetToolbar("1100100100011111")                                              '☆: Developer must customize
	Else
		frm1.txtDateFr.focus
	End If
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call ggoOper.LockField(Document, "Q")

    Set gActiveElement = document.ActiveElement   
End Sub
	
'========================================================================================================
' Name : DbSaveOk
' Desc : Called by MB Area when save operation is successful
'========================================================================================================
Sub DbSaveOk()
    Call InitVariables															     '⊙: Initializes local global variables
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
    Call ggoOper.ClearField(Document, "2")										     '⊙: Clear Contents  Field
    
	Call SetToolbar("1100100000011111")                                              '☆: Developer must customize
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If DbQuery("MQ") = False Then
       Call RestoreToolBar()
       Exit Sub
    End if
    Set gActiveElement = document.ActiveElement   
End Sub
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Sub DbDeleteOk()
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : OpenReference
' Desc : developer describe this line 
'========================================================================================================
Function OpenReference()
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 		
	MsgBox "You need to code this part.",,gLogoName 
	'------ Developer Coding part (End)    -------------------------------------------------------------- 
End Function

'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================
' Name : OpenPopup()
' Desc : developer describe this line 
'========================================================================================================
Function OpenPopup(Byval strCode,Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function
	
	If frm1.txtFCAcctCd.className = "protected" Then Exit Function    
			
	arrParam(0) = "유동성계정팝업"								' 팝업 명칭 
	arrParam(1) = "A_Acct A, A_ACCT_GP B, A_JNL_ACCT_ASSN C"				' TABLE 명칭 
	arrParam(2) = strCode											' Code Condition
	arrParam(3) = ""												' Name Cindition
	If iWhere = 1 Then
		arrParam(4) = "A.GP_CD=B.GP_CD AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND A.Acct_cd=C.Acct_CD and C.trans_type = " & FilterVar("FI003", "''", "S") & "  and C.jnl_cd = 'C" & Right(Trim(frm1.cboLoanFg.Value),1) & "'"
	ElseIf iWhere = 2 Then
		frm1.vspdData.Col  = C_PAY_OBJ
		arrParam(4) = "A.GP_CD=B.GP_CD AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND A.Acct_cd=C.Acct_CD and C.trans_type = " & FilterVar("FI003", "''", "S") & "  and C.jnl_cd = 'C" & Right(Trim(frm1.vspdData.Text),1) & "'"
	End If		
	
	arrParam(5) = frm1.txtFCAcctCd.Alt							' 조건필드의 라벨 명칭 

	arrField(0) = "A.Acct_CD"									' Field명(0)
	arrField(1) = "A.Acct_NM"									' Field명(1)
	arrField(2) = "B.GP_CD"										' Field명(2)
	arrField(3) = "B.GP_NM"										' Field명(3)
			
	arrHeader(0) = frm1.txtFCAcctCd.Alt									' Header명(0)
	arrHeader(1) = frm1.txtFCAcctNm.Alt								' Header명(1)
	arrHeader(2) = "그룹코드"									' Header명(2)
	arrHeader(3) = "그룹명"										' Header명(3)						
	
	IsOpenPop = True
	
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDateFr.focus
		Exit Function
	Else
		With Frm1
			Select Case iWhere
				Case 1
					.txtFCAcctCd.value = arrRet(0)
					.txtFCAcctNm.value = arrRet(1)
					.txtFCAcctCd.focus
				Case 2
					.vspdData.Col  = C_PLAN_ACCT_CD
					.vspdData.Text = arrRet(0)
					.vspdData.Col  = C_PLAN_ACCT_NM
					.vspdData.Text = arrRet(1)
			End Select			
		End With
	End If	

End Function

 '------------------------------------------ OpenPopupLoan() -------------------------------------------------
'	Name : OpenPopupLoan()
'	Description : Loan Number PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopupLoan()
	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(8)	

	If IsOpenPop = True Then Exit Function
	iCalledAspName = AskPRAspName("f4232ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "f4232ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
    
	arrRet = window.showModalDialog(iCalledAspName & "?PGM=" & gStrRequestMenuID , Array(window.parent,arrParam), _
		     "dialogWidth=850px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = ""  Then			
		frm1.txtLoan_No.focus
		Exit Function
	Else		
		frm1.txtLoan_No.value = arrRet(0)
		frm1.txtLoan_Nm.value = arrRet(1)
	End If
	
	frm1.txtLoan_No.focus
End Function

'==========================================================================================
'   Event Name : cboLoanFg_onChange
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub cboLoanFg_onChange()
	frm1.txtFCAcctCd.value = ""
	frm1.txtFCAcctNm.value = ""
End Sub

'==========================================================================================
'   Event Name : txtFCAcctCd_onChange
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub txtFCAcctCd_onChange()
	If Trim(frm1.txtFCAcctCd.value) = "" Then
		frm1.txtFCAcctNm.value = ""
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	lstxtPlanAmtSum = 0
	
    With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		
		If Row > 0 And Col = C_PLAN_ACCT_BT Then
			.Col = Col
			.Row = Row
			Call OpenPopup(.Text, 2)
		Else

			.Row = Row
			.Col = C_Choice
		
			ggoSpread.Source = frm1.vspdData
		
			If .Text = "Y" Then
				If ButtonDown = 0 Then
					ggoSpread.UpdateRow Row
				Else
					ggoSpread.SSDeleteFlag Row,Row
				End If
			Else
				If ButtonDown = 1 Then
					ggoSpread.UpdateRow Row
					.col = C_PLAN_LOC_AMT
					lstxtPlanAmtSum = UNIFormatNumber(UNICDbl(frm1.txtPlanAmtSum.Text) + UNICDbl(.Text),ggAmtOfMoney.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
					frm1.txtPlanAmtSum.Text = lstxtPlanAmtSum
				Else
					ggoSpread.SSDeleteFlag Row,Row				
					.col = C_PLAN_LOC_AMT
					lstxtPlanAmtSum = UNIFormatNumber(UNICDbl(frm1.txtPlanAmtSum.Text) - UNICDbl(.Text),ggAmtOfMoney.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
					frm1.txtPlanAmtSum.Text = lstxtPlanAmtSum
				End If		
			End If
		End If
	End With
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("0001111111")
    gMouseClickStatus = "SPC"
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
    Else
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    	frm1.vspdData.Row = Row
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    End If
    
End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

	With frm1.vspdData
		If Col = C_PLAN_ACCT_CD Then
			.Row = Row
			.Col = C_PLAN_ACCT_NM
			.Text = ""
		End If
	End With

End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(OldLeft , OldTop , NewLeft , NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKeyIndex <> "" Then                         
           If DbQuery("MN") = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
        End If
    End if
End Sub


'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
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
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub


'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc :
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub 

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    If Col <= C_Choice Or NewCol <= C_Choice Then
        Cancel = True
        Exit Sub
    End If
    
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub
  
'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
'   Event Name : txtDateFr_DblClick
'   Event Desc :
'========================================================================================================
Sub txtDateFr_DblClick(Button)
	If Button = 1 Then
		frm1.txtDateFr.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtDateFr.focus
	End If
End Sub

'========================================================================================================
'   Event Name : txtDateTo_DblClick
'   Event Desc :
'========================================================================================================
Sub txtDateTo_DblClick(Button)
	If Button = 1 Then
		frm1.txtDateto.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtDateto.focus
	End If
End Sub

'========================================================================================================
'   Event Name : txtChgDt_DblClick
'   Event Desc :
'========================================================================================================
Sub txtChgDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtChgDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtChgDt.focus
	End If
End Sub

'========================================================================================================
'   Event Name : txtDateFr_KeyPress
'   Event Desc :
'========================================================================================================
Sub txtDateFr_KeyPress(Key)
    If key = 13 Then
		frm1.txtDateTo.Focus
        Call MainQuery
	End If
End Sub

'========================================================================================================
'   Event Name : txtDateTo_KeyPress
'   Event Desc :
'========================================================================================================
Sub txtDateTo_KeyPress(Key)
    If key = 13 Then
		frm1.txtDateFr.Focus
        Call MainQuery
	End If
End Sub

'========================================================================================================
'   Event Name : txtChgDt_KeyPress
'   Event Desc :
'========================================================================================================
Sub txtChgDt_KeyPress(Key)
    If key = 13 Then
        Call MainQuery
	End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->
<BODY TABINDEX="-1" SCROLL="no">
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
					<TD CLASS="CLSLTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>								
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="15" height="23"></td>
						    </TR>
						</TABLE>
					</TD>					
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
									<TD CLASS=TD5 NOWRAP>상환예정일자</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateFr name=txtDateFr CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="조회시작일" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTo name=txtDateTo CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="조회종료일" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS=TD5 NOWRAP>장단기구분</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboLoanFg" ALT="장단기구분" style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 135px" TAG="12xxxU"></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>차입금번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtLOAN_NO" MAXLENGTH="18" SIZE=15  ALT ="차입금번호" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoanNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopupLoan()">
														   <INPUT NAME="txtLOAN_NM" MAXLENGTH="40" SIZE=20  ALT ="차입금내역" tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>사업장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="사업장코드" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenBizAreaCd()"> <INPUT TYPE="Text" NAME="txtBizAreaNm" SIZE=25 tag="24X" ALT="사업장명">&nbsp;~</TD>
								</TR>
								<TR>						   
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBizAreaCd1" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="사업장코드" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenBizAreaCd1()"> <INPUT TYPE="Text" NAME="txtBizAreaNm1" SIZE=25 tag="24X" ALT="사업장명"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;전환일자  </TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtChgDt name=txtChgDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="조회시작일" tag="11X1" VIEWASTEXT></OBJECT>');</SCRIPT>										
									</TD>
									<TD CLASS=TD5 NOWRAP>유동성계정</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtFCAcctCd" MAXLENGTH="18" SIZE=15  ALT ="유동성계정" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnFCAcct" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtFCAcctCd.value, 1)">
														   <INPUT NAME="txtFCAcctNm" MAXLENGTH="40" SIZE=20  ALT ="유동성계정명" tag="14"></TD>
								</TR>	
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
							<TABLE <%=LR_SPACE_TYPE_30%>>
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
								<TD CLASS=TDT>
								<TD CLASS=TD6>
								<TD CLASS=TD5>유동성전환총액(자국)</TD>
								<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name="txtPlanAmtSum" CLASS=FPDS140 title="FPDOUBLESINGLE" ALT="상환예정누계액" tag="24X2"> </OBJECT>');</SCRIPT>&nbsp;
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread"	tag="24" Tabindex="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows"		TAG="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtMode"			TAG="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"		TAG="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="htxtDateFr"		TAG="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="htxtDateTo"		TAG="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="htxtChgDt"			TAG="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthBizAreaCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthSubInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthUsrID"		tag="24" Tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
