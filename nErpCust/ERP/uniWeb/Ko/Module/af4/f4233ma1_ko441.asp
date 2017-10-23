<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : LOAN TERM CHANGE
*  2. Function Name        : F4233MA1
*  3. Program ID           : F4233MA1
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
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<Script Language="VBScript">

Option Explicit                                                        '☜: indicates that All variables must be declared in advance
'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID = "F4233MB1_KO441.asp"                                      'Biz Logic ASP 
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================

Dim C_Choice
Dim C_DelChoice
Dim C_FLT_CONV_DT
Dim C_Loan_No
Dim C_Loan_Nm
Dim C_PLAN_ACCT_CD
Dim C_PLAN_ACCT_NM
Dim C_LOAN_ACCT_CD
Dim C_LOAN_ACCT_NM
Dim C_PAY_PLAN_DT
Dim C_DOC_CUR
Dim C_XCH_RATE
Dim C_PLAN_AMT
Dim C_PLAN_LOC_AMT
Dim C_LOAN_DT
Dim C_DUE_DT
Dim C_TEMP_GL_NO
Dim C_GL_NO
Dim C_REF_NO
Dim C_REF_SEQ
Dim C_PAY_OBJ
Dim C_RDP_CLS_FG
Dim C_RESL_FG
Dim C_CONF_FG


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

    StartDate	= <%=GetSvrDate%>                                               'Get Server DB Date
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
End Sub

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
    Dim i 
	i = 1

	C_Choice				= i : i = i + 1	'선택 
	C_DelChoice				= i : i = i + 1	'전표삭제여부(ko441) 
	C_FLT_CONV_DT			= i : i = i + 1	'전화일자 
	C_Loan_No				= i : i = i + 1	'차입금번호 
	C_Loan_Nm				= i : i = i + 1	'차입내역 
	C_PLAN_ACCT_CD			= i : i = i + 1	'유동성계정 
	C_PLAN_ACCT_NM			= i : i = i + 1	'유동성계정명 
	C_LOAN_ACCT_CD			= i : i = i + 1	'차입계정 
	C_LOAN_ACCT_NM			= i : i = i + 1	'차입계정명 
	C_PAY_PLAN_DT			= i : i = i + 1	'상환예정일자 
	C_DOC_CUR				= i : i = i + 1	'통화 
	C_XCH_RATE				= i : i = i + 1	'환율 
	C_PLAN_AMT				= i : i = i + 1	'상환예정액 
	C_PLAN_LOC_AMT			= i : i = i + 1	'상환예정액(자국)
	C_LOAN_DT				= i : i = i + 1 	'차입일자 
	C_DUE_DT				= i : i = i + 1	'만기일자 
	C_TEMP_GL_NO			= i : i = i + 1	'결의 전표번호 
	C_GL_NO					= i : i = i + 1	'전표번호 
	C_REF_NO				= i : i = i + 1	'LOAN_NO 참조값 
	C_REF_SEQ				= i : i = i + 1	'SEQ	참조값 
	C_PAY_OBJ				= i : i = i + 1	'장단기 원금 OR 이자 구분(EX : LL, SL, LN, SN)
	C_RDP_CLS_FG			= i : i = i + 1	'상환완료여부 
	C_RESL_FG				= i : i = i + 1	'상환여부 
	C_CONF_FG				= i : i = i + 1	'확정여부 
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
	
  	frm1.txtRpyDateFr.Text = frDt   
	frm1.txtRpyDateTo.Text = toDt
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>
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
	Dim RpyYYYYMMDD_fr
	Dim RpyYYYYMMDD_to
	Dim ChgYYYYMMDD_fr
	Dim ChgYYYYMMDD_to
	Dim txtLoan_NO

	Dim strYear,strMonth,strDay
	
	RpyYYYYMMDD_fr = UNIConvDate(frm1.txtRpyDateFr.text)
	RpyYYYYMMDD_to = UNIConvDate(frm1.txtRpyDateTo.text)

	ChgYYYYMMDD_fr = ""
	ChgYYYYMMDD_to = ""
	txtLoan_NO = ""
	If Trim(frm1.txtChgDateFr.text) <> "" Then ChgYYYYMMDD_fr = UNIConvDate(frm1.txtChgDateFr.text)
	If Trim(frm1.txtChgDateTo.text) <> "" Then ChgYYYYMMDD_to = UNIConvDate(frm1.txtChgDateTo.text)
	If Trim(frm1.txtLoan_NO.value) <> "" Then txtLoan_NO = UCase(Trim(frm1.txtLoan_NO.value))
	
	lgKeyStream = RpyYYYYMMDD_fr & Parent.gColSep       'You Must append one character(Parent.gColSep)
	lgKeyStream = lgKeyStream & RpyYYYYMMDD_to & Parent.gColSep       
	lgKeyStream = lgKeyStream & ChgYYYYMMDD_fr & Parent.gColSep 
	lgKeyStream = lgKeyStream & ChgYYYYMMDD_to & Parent.gColSep 
	lgKeyStream = lgKeyStream & txtLoan_NO & Parent.gColSep 
	lgKeyStream = lgKeyStream & Parent.gColSep
	lgKeyStream = lgKeyStream & Parent.gColSep
	lgKeyStream = lgKeyStream & Parent.gColSep
	lgKeyStream = lgKeyStream & frm1.txtBizAreaCd.value & Parent.gColSep
	lgKeyStream = lgKeyStream & frm1.txtBizAreaCd1.value & Parent.gColSep			
End Sub        

	
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset ComboBox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex	
End Sub


'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()
    ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20021206",,parent.gAllowDragDropSpread    

	With frm1.vspdData
	
       .MaxCols		= C_CONF_FG + 1    
       .MaxRows = 0

	   .ReDraw = false
		Call GetSpreadColumnPos("A")
       
								'ColumnPosition		Header					Width  Align(0:L,1:R,2:C)  Row   Length  CharCase(0:L,1:N,2:U)
       ggoSpread.SSSetCheck		C_Choice			,"선택"				,7    ,2                   ,     ,15     
       ggoSpread.SSSetCheck		C_DelChoice			,"전표삭제"			,7    ,2                   ,     ,15     
       ggoSpread.SSSetDate		C_FLT_CONV_DT		,"전환일자"			,10    ,2                  ,Parent.gDateFormat   ,-1
       ggoSpread.SSSetEdit		C_Loan_No			,"차입금번호"       ,13    ,0                  ,     ,18     ,2
       ggoSpread.SSSetEdit		C_Loan_Nm			,"차입내역"			,15    ,0                  ,     ,15     ,2
       ggoSpread.SSSetEdit		C_PLAN_ACCT_CD		,"유동성계정"       ,13    ,0                  ,     ,15     ,2
       ggoSpread.SSSetEdit		C_PLAN_ACCT_NM		,"유동성계정명"		,15    ,0                  ,     ,50     ,2
       ggoSpread.SSSetEdit		C_LOAN_ACCT_CD		,"차입금계정"       ,13    ,0                  ,     ,15     ,2
       ggoSpread.SSSetEdit		C_LOAN_ACCT_NM		,"차입금계정명"		,15    ,0                  ,     ,50     ,2
       ggoSpread.SSSetDate		C_Pay_Plan_Dt		,"상환예정일자"		,12    ,2                  ,Parent.gDateFormat   ,-1
       ggoSpread.SSSetEdit		C_Doc_Cur			,"통화"				,10    ,                   ,     ,15     ,2
       ggoSpread.SSSetFloat		C_Xch_Rate			,"환율"				,11    ,Parent.ggExchRateNo		,ggStrIntegeralPart		,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec	,1      ,True  ,"P"	,"0"		
       ggoSpread.SSSetFloat		C_PLAN_AMT			,"상환예정액"		,15    , "A"	,ggStrIntegeralPart		,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec	,1      ,True  ,"P"	,"0"		
       ggoSpread.SSSetFloat		C_PLAN_LOC_AMT		,"상환예정액(자국)"	,15    ,Parent.ggAmtOfMoneyNo	,ggStrIntegeralPart		,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec	,1      ,True  ,"P"	,"0"		
       ggoSpread.SSSetDate		C_Loan_Dt			,"차입일자"			,12    ,2                  ,Parent.gDateFormat   ,-1
       ggoSpread.SSSetDate		C_Due_Dt			,"상환만기일자"		,12    ,2                  ,Parent.gDateFormat   ,-1
       ggoSpread.SSSetEdit		C_TEMP_GL_NO		,"결의전표번호"		,15		, , , 18
       ggoSpread.SSSetEdit		C_GL_NO				,"회계전표번호"		,15		, , , 18
       ggoSpread.SSSetEdit		C_REF_NO			,"차입금참조번호"	,12    ,0                  ,     ,15     ,2
       ggoSpread.SSSetEdit		C_REF_SEQ			,"순번참조"			,15    ,0                  ,     ,15     ,2
       ggoSpread.SSSetEdit		C_PAY_OBJ			,"상환대상"			,15    ,0                  ,     ,15     ,2
       ggoSpread.SSSetEdit		C_RDP_CLS_FG		,"상환완료여부"		,15    ,0                  ,     ,15     ,2
       ggoSpread.SSSetEdit		C_RESL_FG			,"상환여부"			,15    ,0                  ,     ,15     ,2
       ggoSpread.SSSetEdit		C_CONF_FG			,"전표확정여부"		,15    ,0                  ,     ,15     ,2
	   .ReDraw = true

	   Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
       Call ggoSpread.SSSetColHidden(C_REF_NO ,C_REF_NO	,True)
       Call ggoSpread.SSSetColHidden(C_REF_SEQ ,C_REF_SEQ	,True)
       Call ggoSpread.SSSetColHidden(C_PAY_OBJ ,C_PAY_OBJ	,True)
       Call ggoSpread.SSSetColHidden(C_RDP_CLS_FG ,C_RDP_CLS_FG	,True)
       Call ggoSpread.SSSetColHidden(C_RESL_FG ,C_RESL_FG	,True)
       Call ggoSpread.SSSetColHidden(C_CONF_FG ,C_CONF_FG	,True)

       Call SetSpreadLock

    End With
    
End Sub

'======================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect cell in spread sheet
'======================================================================================================
Sub SetSpreadLock()

	Dim RowCnt
	
	ggoSpread.Source = frm1.vspdData

	With frm1
		.vspdData.ReDraw = False			    
		For RowCnt = 1 To .vspdData.MaxRows
			
			.vspdData.Col = C_RDP_CLS_FG
			.vspdData.Row = RowCnt

			If .vspdData.text = "Y" Then
										'Col-1		Row-1		Col-2		Row-2
				ggoSpread.SpreadLock	C_Choice	, RowCnt	, C_CONF_FG	, RowCnt
			Else
				.vspdData.Col = C_RESL_FG
				If .vspdData.text = "Y" Then
					ggoSpread.SpreadLock	C_Choice	, RowCnt	, C_CONF_FG	, RowCnt
				Else
					.vspdData.Col = C_CONF_FG
					If .vspdData.text = "C" Then
						ggoSpread.SpreadLock	C_FLT_CONV_DT	, RowCnt	, C_CONF_FG	, RowCnt
					Else
						ggoSpread.SpreadLock	C_FLT_CONV_DT	, RowCnt	, C_CONF_FG	, RowCnt
					End If
				End If
			End If

		Next

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
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> parent.UC_PROTECTED Then
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
	Dim i 

	i = 1
    
    Select Case UCase(pvSpdNo)
       Case "A"

            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_Choice				= iCurColumnPos(i) : i = i + 1
			C_DelChoice				= iCurColumnPos(i) : i = i + 1
			C_FLT_CONV_DT			= iCurColumnPos(i) : i = i + 1
			C_Loan_No				= iCurColumnPos(i) : i = i + 1    
			C_Loan_Nm				= iCurColumnPos(i) : i = i + 1
			C_PLAN_ACCT_CD			= iCurColumnPos(i) : i = i + 1
			C_PLAN_ACCT_NM			= iCurColumnPos(i) : i = i + 1
			C_LOAN_ACCT_CD			= iCurColumnPos(i) : i = i + 1
			C_LOAN_ACCT_NM			= iCurColumnPos(i) : i = i + 1
			C_PAY_PLAN_DT			= iCurColumnPos(i) : i = i + 1
			C_Doc_Cur				= iCurColumnPos(i) : i = i + 1
			C_Xch_Rate				= iCurColumnPos(i) : i = i + 1
			C_PLAN_AMT				= iCurColumnPos(i) : i = i + 1
			C_PLAN_LOC_AMT			= iCurColumnPos(i) : i = i + 1
			C_Loan_Dt			    = iCurColumnPos(i) : i = i + 1
			C_Due_Dt				= iCurColumnPos(i) : i = i + 1
			C_TEMP_GL_NO			= iCurColumnPos(i) : i = i + 1
			C_GL_NO					= iCurColumnPos(i) : i = i + 1
			C_REF_NO				= iCurColumnPos(i) : i = i + 1
			C_REF_SEQ				= iCurColumnPos(i) : i = i + 1
			C_PAY_OBJ				= iCurColumnPos(i) : i = i + 1
			C_RDP_CLS_FG			= iCurColumnPos(i) : i = i + 1
			C_RESL_FG				= iCurColumnPos(i) : i = i + 1
			C_CONF_FG				= iCurColumnPos(i) : i = i + 1

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

	frm1.txtRpyDateFr.focus
	Call FncSetToolBar("New")                                              '☆: Developer must customize
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call InitComboBox
	Call CookiePage (0)                                                              '☜: Check Cookie
	
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
	
	If CompareDateByFormat(frm1.fpRpyDateFr.text,frm1.fpRpyDateTo.text,frm1.fpRpyDateFr.Alt,frm1.fpRpyDateTo.Alt, _
        	               "970025",frm1.fpRpyDateFr.UserDefinedFormat,Parent.gComDateType, true) = False Then	   
	   frm1.fpRpyDateFr.focus
	   Set gActiveElement = document.ActiveElement
	   
		Exit Function
	End If
	
	If CompareDateByFormat(frm1.txtChgDateFr.text,frm1.txtChgDateTo.text,frm1.txtChgDateFr.Alt,frm1.txtChgDateTo.Alt, _
        	               "970025",frm1.txtChgDateFr.UserDefinedFormat,Parent.gComDateType, true) = False Then	   
	   frm1.txtChgDateFr.focus
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

    Call InitVariables                                                           '⊙: Initializes local global variables
    frm1.txtPlanAmtSum.Text = "0"
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
Function FncInsertRow()
   
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
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
	Call SetSpreadLock
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
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")		             '⊙: Data is changed.  Do you want to exit? 
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

    Call DisableToolBar(TBC_QUERY)                                               '☜: Disable Query Button Of ToolBar
    Call LayerShowHide(1)                                                        '☜: Show Processing Message
	
    Call MakeKeyStream(pDirect)

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
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

    Call DisableToolBar(TBC_SAVE)                                                '☜: Disable Save Button Of ToolBar
    Call LayerShowHide(1)                                                        '☜: Show Processing Message
		
    Frm1.txtMode.value        = parent.UID_M0002                                        '☜: Delete
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
           .vspdData.Col = 1
           
           If .vspdData.Value = 1 Then        

				.vspdData.Col = C_DelChoice
				If  .vspdData.Value = 1 then 
													  strVal = strVal & "D"                      & iColSep
				Else										
													  strVal = strVal & "R"                      & iColSep
				End if										
													  strVal = strVal & lRow                     & iColSep
				.vspdData.Col = C_LOAN_NO			: strVal = strVal & Trim(.vspdData.Text)     & iColSep	'2 차입금번호 
				.vspdData.Col = C_PAY_PLAN_DT		: strVal = strVal & UNIConvDAte(Trim(.vspdData.Text))     & iColSep  '3 상환예정일자(이자는 지급, 원금은 상환)
				.vspdData.Col = C_REF_NO			: strVal = strVal & Trim(.vspdData.Text)     & iColSep  '4 LOAN_NO 참조값 
				.vspdData.Col = C_REF_SEQ			: strVal = strVal & Trim(.vspdData.Text)     & iColSep  '5 SEQ 참조값 
				.vspdData.Col = C_PAY_OBJ			: strVal = strVal & Trim(.vspdData.Text)     & iRowSep  '6 이자 OR 원금 장단기 구분 
			
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
	Call SetToolbar("1100100000011111")                                              '☆: Developer must customize
	Frm1.vspdData.Focus
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call InitData()
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

'==========================================================
'툴바버튼 세팅 
'==========================================================
Function FncSetToolBar(Cond)
	Select Case UCase(Cond)
	Case "NEW"
		Call SetToolbar("1100110100101111")
	Case "QUERY"
		Call SetToolbar("1100111100111111")
	End Select
End Function

'=======================================================================================================
'   Event Name : SetGridFocus
'   Event Desc :
'=======================================================================================================    
Sub SetGridFocus()	
   
	Frm1.vspdData.Row = 1
	Frm1.vspdData.Col = 1
	Frm1.vspdData.Action = 1
		
End Sub


'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
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

'============================================================
'회계전표 팝업 
'============================================================
Function OpenPopupGL()

	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(1)	

	If IsOpenPop = True Then Exit Function
	
	With frm1.vspdData
		If .ActiveRow > 0 Then
			.Row = .ActiveRow
			.Col = C_GL_NO
			arrParam(0) = Trim(.Text)	'회계전표번호 
			arrParam(1) = ""			'Reference번호 
		Else
			Call DisplayMsgBox("900025","X","X","X")
			Exit Function
		End If
	End With
	
	iCalledAspName = AskPRAspName("a5120ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	frm1.txtRpyDateFr.focus
	
End Function

'============================================================
'결의전표 팝업 
'============================================================
Function OpenPopupTempGL()

	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(1)	

	If IsOpenPop = True Then Exit Function
	
	With frm1.vspdData
		If .ActiveRow > 0 Then
			.Row = .ActiveRow
			.Col = C_TEMP_GL_NO
			arrParam(0) = Trim(.Text)	'회계전표번호 
			arrParam(1) = ""			'Reference번호 
		Else
			Call DisplayMsgBox("900025","X","X","X")
			Exit Function
		End If
	End With
	
	iCalledAspName = AskPRAspName("a5130ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	frm1.txtRpyDateFr.focus
	
End Function
'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	lstxtPlanAmtSum = 0
	
    With frm1.vspdData
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
				'.col = C_CHG_DT		:.TEXT = UNIDateClientFormat(strSvrDate)
				.col = C_PLAN_LOC_AMT
				lstxtPlanAmtSum = UNIFormatNumber(UNICDbl(frm1.txtPlanAmtSum.Text) + UNICDbl(.Text),ggAmtOfMoney.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
				frm1.txtPlanAmtSum.Text = lstxtPlanAmtSum
			Else
				ggoSpread.SSDeleteFlag Row,Row
				'.col = C_CHG_DT		:.TEXT = ""				
				.col = C_PLAN_LOC_AMT
				lstxtPlanAmtSum = UNIFormatNumber(UNICDbl(frm1.txtPlanAmtSum.Text) - UNICDbl(.Text),ggAmtOfMoney.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
				frm1.txtPlanAmtSum.Text = lstxtPlanAmtSum				
			End If			
		End If
	End With
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

'========================================================================================================
'   Event Name : txtRpyDateFr_DblClick
'   Event Desc :
'========================================================================================================
Sub txtRpyDateFr_DblClick(Button)
	If Button = 1 Then
		frm1.txtRpyDateFr.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtRpyDateFr.Focus
	End If
End Sub

'========================================================================================================
'   Event Name : txtRpyDateTo_DblClick
'   Event Desc :
'========================================================================================================
Sub txtRpyDateTo_DblClick(Button)
	If Button = 1 Then
		frm1.txtRpyDateTo.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtRpyDateTo.Focus
	End If
End Sub
'========================================================================================================
'   Event Name : txtChgDateFr_DblClick
'   Event Desc :
'========================================================================================================
Sub txtChgDateFr_DblClick(Button)
	If Button = 1 Then
		frm1.txtChgDateFr.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtChgDateFr.Focus
	End If
End Sub
'========================================================================================================
'   Event Name : txtChgDateTo_DblClick
'   Event Desc :
'========================================================================================================
Sub txtChgDateTo_DblClick(Button)
	If Button = 1 Then
		frm1.txtChgDateTo.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtChgDateTo.Focus
	End If
End Sub
'========================================================================================================
'   Event Name : txtRpyDateFr_KeyPress
'   Event Desc :
'========================================================================================================
Sub txtRpyDateFr_KeyPress(Key)
    If key = 13 Then
		frm1.txtRpyDateTo.Focus
        Call MainQuery
	End If
End Sub

'========================================================================================================
'   Event Name : txtRpyDateTo_KeyPress
'   Event Desc :
'========================================================================================================
Sub txtRpyDateTo_KeyPress(Key)
    If key = 13 Then
		frm1.txtRpyDateFr.Focus
        Call MainQuery
	End If
End Sub


'========================================================================================================
'   Event Name : txtChgDateFr_KeyPress
'   Event Desc :
'========================================================================================================
Sub txtChgDateFr_KeyPress(Key)
    If key = 13 Then
		frm1.txtChgDateTo.Focus
        Call MainQuery
	End If
End Sub


'========================================================================================================
'   Event Name : txtChgDateTo_KeyPress
'   Event Desc :
'========================================================================================================
Sub txtChgDateTo_KeyPress(Key)
    If key = 13 Then
		frm1.txtChgDateFr.Focus
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>					
					<TD WIDTH=*>
						<TABLE CELLSPACING=0 CELLPADDING=0 align=right>
							<TR>
								<td>
									<A HREF="VBSCRIPT:OpenPopupTempGL()">결의전표</a> |
									<A HREF="VBSCRIPT:OpenPopupGL()">회계전표</a>
								</td>
						    </TR>
						</TABLE>
					</TD>
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
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpRpyDateFr name=txtRpyDateFr CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="상환시작일" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpRpyDateTo name=txtRpyDateTo CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="상환종료일" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS="TD5" NOWRAP>차입금번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtLOAN_NO" MAXLENGTH="18" SIZE=15  ALT ="차입금번호" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoanNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopupLoan()">
														   <INPUT NAME="txtLOAN_NM" MAXLENGTH="40" SIZE=20  ALT ="차입금내역" tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>전환일자</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpChgDateFr name=txtChgDateFr CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="전환시작일" tag="11X1" VIEWASTEXT></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpChgDateTo name=txtChgDateTo CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="전환종료일" tag="11X1" VIEWASTEXT></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS="TD5" NOWRAP>사업장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="사업장코드" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenBizAreaCd()"> <INPUT TYPE="Text" NAME="txtBizAreaNm" SIZE=25 tag="24X" ALT="사업장명">&nbsp;~</TD>
								</TR>
								<TR>																		
									<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
									<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBizAreaCd1" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="사업장코드" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenBizAreaCd1()"> <INPUT TYPE="Text" NAME="txtBizAreaNm1" SIZE=25 tag="24X" ALT="사업장명"></TD>									
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
								<TD CLASS=TD5>유동성전환취소총액(자국)</TD>
								<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name="txtPlanAmtSum" CLASS=FPDS140 title="FPDOUBLESINGLE" ALT="상환예정누계액" tag="34X2"> </OBJECT>');</SCRIPT>&nbsp;
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
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread"	tag="24" Tabindex="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows"		TAG="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtMode"			TAG="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"		TAG="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="htxtRpyDateFr"		TAG="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="htxtRpyDateTo"		TAG="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="htxtChgDateFr"		TAG="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="htxtChgDateTo"		TAG="24" Tabindex="-1">
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
