<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Accounting
*  2. Function Name        : Treasury - Loan
*  3. Program ID           : f4240ma1
*  4. Program Name         : 선급이자 월결산 
*  5. Program Desc         : 
*  6. Comproxy List        :
*  7. Modified date(First) : 2002/03/30
*  8. Modified date(Last)  : 2001/03/30
*  9. Modifier (First)     : Hwang Eun Hee
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
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance
'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID = "f4240mb1_ko441.asp"                                      'Biz Logic ASP 
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Dim C_CHOICE_FG         
Dim C_INT_FG
Dim C_SEQ				
Dim C_LOAN_NO
Dim C_LOAN_BT
Dim C_LOAN_NM
Dim C_PAY_NO
Dim C_PAY_BT
Dim C_INT_EXP_ACCT_CD	
Dim C_INT_EXP_ACCT_BT 
Dim C_INT_EXP_ACCT_NM 
Dim C_INT_CLS_AMT		
Dim C_INT_CLS_LOC_AMT		
Dim C_INT_CLS_PLAN_AMT  
Dim C_INT_CLS_PLAN_LOC_AMT 
Dim C_DOC_CUR			
Dim C_XCH_RATE		
Dim C_ADV_INT_ACCT_CD	
Dim C_ADV_INT_ACCT_NM	
Dim C_INT_CLS_DT		
Dim C_INT_PAY_DT		
Dim C_INT_PAY_AMT		
Dim C_INT_PAY_LOC_AMT	
Dim C_INT_RATE		
Dim C_LOAN_DT			
Dim C_DUE_DT			
Dim C_TEMP_GL_NO		
Dim C_GL_NO			
Dim C_CLS_FG			
Dim C_CONF_FG			

'네패스 추가항목...kbs...2009.08.28
Dim	C_LAST_PAY_DT

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop          
Dim lgIsOpenPop

Dim lgKeyPos                                                '☜: Key위치                               
Dim lgKeyPosVal                                             '☜: Key위치 Value   
Dim GL_Number													'☜: gl_no 값  	

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
	C_CHOICE_FG				=  1  
	C_INT_FG				=  2
	C_SEQ					=  3
	C_LOAN_NO				=  4
	C_LOAN_BT				=  5
	C_LOAN_NM				=  6
	C_PAY_NO				=  7
	C_PAY_BT				=  8
	C_INT_EXP_ACCT_CD		=  9
	C_INT_EXP_ACCT_BT		= 10
	C_INT_EXP_ACCT_NM		= 11
	C_INT_CLS_AMT			= 12
	C_INT_CLS_LOC_AMT		= 13
	C_INT_CLS_PLAN_AMT		= 14
	C_INT_CLS_PLAN_LOC_AMT	= 15
	C_DOC_CUR				= 16
	C_XCH_RATE				= 17
	C_ADV_INT_ACCT_CD		= 18
	C_ADV_INT_ACCT_NM		= 19
	C_INT_CLS_DT			= 20
	C_INT_PAY_DT			= 21
	C_INT_PAY_AMT			= 22
	C_INT_PAY_LOC_AMT		= 23
	C_INT_RATE				= 24
	C_LOAN_DT				= 25
	C_DUE_DT				= 26
	C_TEMP_GL_NO			= 27
	C_GL_NO					= 28
	C_CLS_FG				= 29
	C_CONF_FG				= 30

	'네패스 추가항목...kbs...2009.08.28
	C_LAST_PAY_DT			= 31

End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Dim LastDt				 	
	Dim BaseDate
	DIm strYear, strMonth, strDay
	dim frdt
	
	
	LastDt     = UNIGetLastDay ("<%=GetSvrDate%>",Parent.gServerDateFormat) 
	frm1.txtBaseDt.Text  = UniConvDateAToB(LastDt, Parent.gServerDateFormat, Parent.gDateFormat)
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
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
   DIm FirstDT
   Dim LastDT
   DIm strYear,strMonth,strDay
   
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Call ExtractDateFrom(frm1.txtBaseDt.Text, frm1.txtBaseDt.UserDefinedFormat, Parent.gComDateType, strYear,strMonth,strDay)
	FirstDT = UniConvYYYYMMDDToDate(Parent.gServerDateFormat, strYear, strMonth, "01")	
	LastDT = UNIGetLastDay (UNIConvDate(FirstDT),Parent.gServerDateFormat) 

	lgKeyStream = LastDt & Parent.gColSep & FirstDT & Parent.gColSep
	lgKeyStream = lgKeyStream  & Trim(frm1.txtLOAN_NO.value) & Parent.gColSep
	lgKeyStream = lgKeyStream  & Trim(frm1.txtIntExpAcctCd.value) & Parent.gColSep
	If frm1.rdoGiYes.checked = true Then
		lgKeyStream = lgKeyStream & "Y" & Parent.gColSep
	ELSE
		lgKeyStream = lgKeyStream & "N" & Parent.gColSep
	End If
	lgKeyStream = lgKeyStream & frm1.txtBizAreaCd.value & Parent.gColSep
	lgKeyStream = lgKeyStream & frm1.txtBizAreaCd1.value & Parent.gColSep	

	'네패스...결산일자를 기준으로 금액 계산...임미희과장 요청...2009.08.28...kbs
	lgKeyStream = lgKeyStream & frm1.txtBaseDt.Text & Parent.gColSep	

	   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    ggoSpread.Source = frm1.vspdData
                       'Data        Seperator            Column position 
'    ggoSpread.SetCombo "결산" & vbTab & "미결산" , C_INT_FG  
    
	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("F3015", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_CLS_FG
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_INT_FG

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : vspdData_ComboSelChange
' Desc : ComboBox에 값 변경시 처리 
'========================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	
	With frm1.vspdData
		
		.Row = Row
    
        Select Case Col
            Case C_INT_FG
                .Col = Col
                intIndex = .Value        '  COMBO의 VALUE값 
				.Col = C_CLS_FG      '  CODE값란으로 이동 
				.Value = intIndex        '  CODE란의 값은 COMBO의 VALUE값이된다.
				
		End Select
	End With

   	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset ComboBox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 

	With frm1.vspdData
		If frm1.rdoGiNo.checked = true then
			For intRow = 1 To .MaxRows			
				.Row = intRow
				.Col = C_CHOICE_FG	:	.text= "1"
			Next
		End If
	End With		
End Sub




'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()
    ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20030820",,parent.gAllowDragDropSpread    

	With frm1.vspdData
	
	'네패스 추가항목...kbs...2009.08.28
       '.MaxCols		= C_CONF_FG + 1    
	.MaxCols = C_LAST_PAY_DT + 1    

       .MaxRows = 0
	   .ReDraw = false
		
		Call AppendNumberPlace("6","4","6")
		Call GetSpreadColumnPos("A")   
			   
		                      'ColumnPosition			Header					Width	Align(0:L,1:R,2:C)  Row   Length  CharCase(0:L,1:N,2:U)
		ggoSpread.SSSetCheck   C_CHOICE_FG				,"선택"				,6		,					, "", true, -1
		ggoSpread.SSSetCombo   C_INT_FG					,"결산여부"		,10     ,2                  ,False         ,-1
		ggoSpread.SSSetFloat   C_SEQ					,"번호"				,8		,"6"           ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,1      ,True  ,"Z" ,"1"      ,"134x30"
		ggoSpread.SSSetEdit    C_LOAN_NO				,"차입번호"			,13		,		,	,20
		ggoSpread.SSSetButton  C_LOAN_BT
		ggoSpread.SSSetEdit    C_LOAN_NM				,"차입내역"			,15		,		,	,30		,1
		ggoSpread.SSSetEdit    C_PAY_NO					,"상환번호"			,13		,		,	,20
		ggoSpread.SSSetButton  C_PAY_BT
		ggoSpread.SSSetEdit    C_INT_EXP_ACCT_CD		,"이자비용계정"		,13		,		,	, 20
		ggoSpread.SSSetButton  C_INT_EXP_ACCT_BT
		ggoSpread.SSSetEdit    C_INT_EXP_ACCT_NM		,"이자비용계정명"	,18		,					,		, 20     ,2
		ggoSpread.SSSetFloat   C_INT_CLS_AMT			,"결산금액"		,15		,"A"					,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat   C_INT_CLS_LOC_AMT		,"결산금액(자국)"	,15		,Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat   C_INT_CLS_PLAN_AMT		,"결산계획금액"		,15		,"A"					,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat   C_INT_CLS_PLAN_LOC_AMT	,"결산계획금액(자국)",15	,Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetEdit    C_DOC_CUR				,"거래통화"			,12		,		,	,15		,2
		ggoSpread.SSSetFloat   C_XCH_RATE				,"환율"				,10		,Parent.ggExchRateNo	,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit    C_ADV_INT_ACCT_CD		,"선급이자계정"		,13		,		,	, 20
		ggoSpread.SSSetEdit    C_ADV_INT_ACCT_NM		,"선급이자계정명"	,18		,       ,	, 20     ,2
		ggoSpread.SSSetDate    C_INT_CLS_DT				,"결산일자"		,12		,2						,Parent.gDateFormat   ,-1
		ggoSpread.SSSetDate    C_INT_PAY_DT				,"이자지급일자"		,12		,2						,Parent.gDateFormat   ,-1
		ggoSpread.SSSetFloat   C_INT_PAY_AMT		    ,"이자지급액"		,15		,"A"					,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat   C_INT_PAY_LOC_AMT		,"이자지급액(자국)",15		,Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec

		'네패스 추가항목...kbs...2009.08.28
		'ggoSpread.SSSetFloat   C_INT_RATE				,"이자율"			,10		,Parent.ggExchRateNo	,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat   C_INT_RATE				,"이자율"			,10		,"6"	,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"

		ggoSpread.SSSetDate    C_LOAN_DT				,"차입일자"			,12		,2						,Parent.gDateFormat   ,-1
		ggoSpread.SSSetDate    C_DUE_DT					,"상환만기일자"		,12		,2						,Parent.gDateFormat   ,-1
		ggoSpread.SSSetEdit    C_TEMP_GL_NO				,"결의전표번호"		,15		,		,	,18
		ggoSpread.SSSetEdit    C_GL_NO					,"회계전표번호"		,15		,		,	,18
		ggoSpread.SSSetEdit    C_CLS_FG					,"결산FG"			,12		,		,	,15		,2
		ggoSpread.SSSetEdit    C_CONF_FG				,"승인FG"			,12		,		,	,15		,2

		'네패스 추가항목...kbs...2009.08.28
		ggoSpread.SSSetDate    C_LAST_PAY_DT			,"최종이자상환일자"     ,14		,2                  ,Parent.gDateFormat   ,-1

		.ReDraw = true
		  
		Call ggoSpread.MakePairsColumn(C_INT_EXP_ACCT_CD,C_INT_EXP_ACCT_BT)
		Call ggoSpread.MakePairsColumn(C_LOAN_NO,C_LOAN_BT)
		Call ggoSpread.MakePairsColumn(C_PAY_NO,C_PAY_BT)

		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_SEQ ,C_SEQ	,True)
		Call ggoSpread.SSSetColHidden(C_CLS_FG ,C_CLS_FG	,True)
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
			
			.vspdData.Col = C_CLS_FG
			.vspdData.Row = RowCnt

			If .vspdData.text = "N" Then
										'Col-1				Row-1		Col-2		Row-2

				'네패스 추가항목...kbs...2009.08.28
				'ggoSpread.SpreadLock	C_INT_FG			, RowCnt	, C_CONF_FG		, RowCnt
				ggoSpread.SpreadLock	C_INT_FG			, RowCnt	, C_LAST_PAY_DT		, RowCnt

				ggoSpread.SpreadUnLock	C_CHOICE_FG			, RowCnt	, C_CHOICE_FG	, RowCnt
				ggoSpread.SpreadUnLock	C_INT_EXP_ACCT_CD	, RowCnt	, C_INT_EXP_ACCT_BT	, RowCnt
				ggoSpread.SSSetRequired	C_INT_EXP_ACCT_CD   , RowCnt	, RowCnt
				ggoSpread.SpreadUnLock	C_INT_CLS_AMT		, RowCnt	, C_INT_CLS_AMT	, RowCnt
				ggoSpread.SSSetRequired	C_INT_CLS_AMT		, RowCnt	, RowCnt
				ggoSpread.SpreadUnLock	C_INT_CLS_LOC_AMT	, RowCnt	, C_INT_CLS_LOC_AMT	, RowCnt
			Else
				.vspdData.Col = C_CONF_FG
				If Trim(.vspdData.text) = "C" Then

					'네패스 추가항목...kbs...2009.08.28
					'ggoSpread.SpreadLock	C_CHOICE_FG	, RowCnt	, C_CLS_FG	, RowCnt
					ggoSpread.SpreadLock	C_CHOICE_FG	, RowCnt	, C_LAST_PAY_DT , RowCnt

				Else

					'네패스 추가항목...kbs...2009.08.28
					'ggoSpread.SpreadLock	C_CHOICE_FG + 1	, RowCnt	, C_CLS_FG	, RowCnt
					ggoSpread.SpreadLock	C_CHOICE_FG + 1	, RowCnt	, C_LAST_PAY_DT , RowCnt

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
	ggoSpread.Source = frm1.vspdData
	With frm1
		.vspdData.ReDraw = False
		ggoSpread.SpreadLock	-1			, pvStartRow	, -1			, pvEndRow
		ggoSpread.SpreadUnLock	C_CHOICE_FG		, pvStartRow	, C_CHOICE_FG	, pvEndRow
		ggoSpread.SpreadUnLock	C_INT_EXP_ACCT_CD		, pvStartRow	, C_INT_EXP_ACCT_BT	, pvEndRow
		ggoSpread.SpreadUnLock	C_INT_CLS_AMT		, pvStartRow	, C_INT_CLS_AMT	, pvEndRow
		ggoSpread.SpreadUnLock	C_INT_CLS_LOC_AMT	, pvStartRow	, C_INT_CLS_LOC_AMT	, pvEndRow
		ggoSpread.SpreadUnLock	C_LOAN_NO	, pvStartRow	, C_LOAN_BT	, pvEndRow
		ggoSpread.SpreadUnLock	C_PAY_NO	, pvStartRow	, C_PAY_BT	, pvEndRow

		ggoSpread.SSSetRequired	C_INT_EXP_ACCT_CD   , pvStartRow	, pvEndRow
		ggoSpread.SSSetRequired	C_INT_CLS_AMT   , pvStartRow	, pvEndRow
		ggoSpread.SSSetRequired	C_LOAN_NO   , pvStartRow	, pvEndRow
		
		.vspdData.ReDraw = True	
	End With
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
            
			C_CHOICE_FG				= iCurColumnPos(1)
			C_INT_FG				= iCurColumnPos(2)
			C_SEQ					= iCurColumnPos(3)    
			C_LOAN_NO				= iCurColumnPos(4)
			C_LOAN_BT				= iCurColumnPos(5)
			C_LOAN_NM				= iCurColumnPos(6)
			C_PAY_NO				= iCurColumnPos(7)
			C_PAY_BT				= iCurColumnPos(8)
			C_INT_EXP_ACCT_CD		= iCurColumnPos(9)
			C_INT_EXP_ACCT_BT		= iCurColumnPos(10)
			C_INT_EXP_ACCT_NM		= iCurColumnPos(11)
			C_INT_CLS_AMT			= iCurColumnPos(12)
			C_INT_CLS_LOC_AMT		= iCurColumnPos(13)
			C_INT_CLS_PLAN_AMT		= iCurColumnPos(14)
			C_INT_CLS_PLAN_LOC_AMT	= iCurColumnPos(15)
			C_DOC_CUR				= iCurColumnPos(16)
			C_XCH_RATE				= iCurColumnPos(17)
			C_ADV_INT_ACCT_CD		= iCurColumnPos(18)
			C_ADV_INT_ACCT_NM		= iCurColumnPos(19)
			C_INT_CLS_DT		    = iCurColumnPos(20)
			C_INT_PAY_DT		    = iCurColumnPos(21)
			C_INT_PAY_AMT		    = iCurColumnPos(22)
			C_INT_PAY_LOC_AMT		= iCurColumnPos(23)
			C_INT_RATE				= iCurColumnPos(24)
			C_LOAN_DT				= iCurColumnPos(25)
			C_DUE_DT				= iCurColumnPos(26)
			C_TEMP_GL_NO		    = iCurColumnPos(27)
			C_GL_NO					= iCurColumnPos(28)
			C_CLS_FG				= iCurColumnPos(29)
			C_CONF_FG				= iCurColumnPos(30)

			'네패스 추가항목...kbs...2009.08.28
			C_LAST_PAY_DT					= iCurColumnPos(31)

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
	arrParam(4) = ""							' Where Condition
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
	arrParam(4) = ""							' Where Condition
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
	frm1.btnClsProcess.disabled = true
	frm1.btnClsCancel.disabled = true

	Call AppendNumberPlace("7","15","2")
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")		
            
    Call InitSpreadSheet                                                             'Setup the Spread sheet

	Call InitVariables
    Call SetDefaultVal
    Call InitComboBox

	frm1.txtBaseDt.focus
	Call SetToolbar("1100000000011111")                                              '☆: Developer must customize
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
    
	Err.Clear                                                                    '☜: Clear err status
	FncQuery = False															 '☜: Processing is NG

    ggoSpread.Source = Frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
    
    If frm1.txtLOAN_NO.value = "" then							'☜: Clear Contents  Field
		frm1.txtLOAN_NM.value = ""
    End If															
    
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

    Call InitVariables                                                           '⊙: Initializes local global variables
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
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'In Multi, You need not to implement this area
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'In Multi, You need not to implement this area
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Function

'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    Dim lRow
    Dim strPayno, strLoanno, strPayamt, strPayClsDT, strPayClsAMT
    
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

'	With Frm1.vspdData
'		For lRow = 1 To .MaxRows
'			.Row = lRow
'	 		.Col = C_CLS_FG						'☜: Update(Yes Check)
'			If trim(.value) = "N" Then
'				.Col = 0
'				Select Case .Text
'				    Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag 
'						.Row = lRow
'	 					.Col = C_PAY_NO		:		strPayno = Trim(.Text)
'	 					.Col = C_LOAN_NO	:		strLoanno = Trim(.Text)
'	 					.Col = C_INT_PAY_AMT :		strPayamt = Trim(.Text)
'	 					.Col = C_INT_CLS_DT :		strPayClsDT = Trim(.Text)
'	 					.Col = C_INT_CLS_PLAN_AMT :	strPayClsAMT = Trim(.Text)
'
'						If strPayno <> "" Then
'							Call CommonQueryRs("sum(A.int_cls_plan_amt)","f_ln_mon_adv_int A"," A.LOAN_NO = " & FilterVar(strLoanno, "''", "S") & _
'															" AND A.PAY_NO = " & FilterVar(strPayno, "''", "S") & _
'															" AND A.INT_CLS_PLAN_DT <> " & FilterVar(strPayClsDT, "''", "S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
'							If UNICdbl(replace(lgF0, chr(11), "")) + UNICdbl(strPayClsAMT)  > UNICdbl(strPayamt) Then
'								Call DisplayMsgBox("141171","x","x","x")
'								Exit Function
'							End If
'						End If
'				    
'				    Case ggoSpread.DeleteFlag
'				End Select
'
'			End If
'		Next
'	End With

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
			SetSpreadColor .ActiveRow, .ActiveRow
    
            .ReDraw = True
		    .Focus
		 End If
	End With

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	' Clear key field
	'---------------------------------------------------------------------------------------------------- 
	With Frm1.VspdData
			.row = .ActiveRow
			.col = C_INT_FG
			.value = "N"
			.col = C_CLS_FG
			.Text = "N"
 	End With

	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,C_DOC_CUR,C_INT_CLS_AMT,"A" ,"I","X","X")
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,C_DOC_CUR,C_INT_CLS_PLAN_AMT,"A" ,"I","X","X")
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,C_DOC_CUR,C_INT_PAY_AMT,"A" ,"I","X","X")
	
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
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,C_DOC_CUR,C_INT_CLS_AMT,"A" ,"I","X","X")
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,C_DOC_CUR,C_INT_CLS_PLAN_AMT,"A" ,"I","X","X")
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,C_DOC_CUR,C_INT_PAY_AMT,"A" ,"I","X","X")
	Call SetSpreadLock
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCancel = False                                                            '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow
    Dim ii
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    FncInsertRow = False                                                         '☜: Processing is NG

	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then
			Exit Function
		End If	
	End If

	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow  .vspdData.ActiveRow, imRow
		SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
	'------ Developer Coding part (Start ) --------------------------------------------------------------

		for ii = frm1.vspdData.ActiveRow to frm1.vspdData.ActiveRow + imRow - 1
			.vspdData.row = ii
			.vspdData.col = C_INT_FG
			.vspdData.value = "N"
			.vspdData.col = C_CLS_FG
			.vspdData.Text = "N"
			.vspdData.col = C_INT_CLS_DT
			.vspdData.Text = frm1.txtBaseDt.text
		Next

        .vspdData.ReDraw = True
    End With

	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,C_DOC_CUR,C_INT_CLS_AMT,"A" ,"I","X","X")
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,C_DOC_CUR,C_INT_CLS_PLAN_AMT,"A" ,"I","X","X")
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,C_DOC_CUR,C_INT_PAY_AMT,"A" ,"I","X","X")
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    
    Set gActiveElement = document.ActiveElement  
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows
    DIM iRow
    DIM ii
    FncDeleteRow = False                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If Frm1.vspdData.MaxRows < 1 then
       Exit function
	End if
	
    With Frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	ggoSpread.Source = frm1.vspdData 
	iRow = frm1.vspdData.ActiveRow
	ggoSpread.SpreadLock	-1		, iRow	, -1	, iRow + lDelRows - 1
	for ii = iRow to iRow + lDelRows - 1
		frm1.vspdData.Col = C_CHOICE_FG
		frm1.vspdData.Row = ii
		frm1.vspdData.Text = "0"
	NEXT

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncDeleteRow = True                                                          '☜: Processing is OK
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

	Call Parent.FncExport(Parent.C_MULTI)

    FncExcel = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncFind(Parent.C_MULTI, True)

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
'	Call InitData()
	Call SetSpreadLock
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,-1,-1,frm1.txtDocCur.value,C_INT_CLS_AMT,"A" ,"I","X","X")
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,-1,-1,frm1.txtDocCur.value,C_INT_CLS_PLAN_AMT,"A" ,"I","X","X")
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,-1,-1,frm1.txtDocCur.value,C_INT_PAY_AMT,"A" ,"I","X","X")
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

    Call DisableToolBar(Parent.TBC_QUERY)                                               '☜: Disable Query Button Of ToolBar
    Call LayerShowHide(1)                                                        '☜: Show Processing Message

    Call MakeKeyStream(pDirect)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    If frm1.rdoGiYes.Checked = True THen
		frm1.btnClsProcess.disabled = true
    Else
		frm1.btnClsCancel.disabled = true
    ENd if
   
			With Frm1
				strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001						         
			    strVal = strVal     & "&txtKeyStream="       & lgKeyStream               '☜: Query Key
			    strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
			    strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex         '☜: Next key tag
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
	Dim Rowchecked

    DbSave = False                                                               '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    Call DisableToolBar(Parent.TBC_SAVE)                                                '☜: Disable Save Button Of ToolBar
    Call LayerShowHide(1)                                                        '☜: Show Processing Message
		
    Frm1.txtMode.value        = Parent.UID_M0002                                        '☜: Delete
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	
    ggoSpread.Source = frm1.vspdData

    strVal = ""
    strDel = ""
    lGrpCnt = 1
	iColSep = Parent.gColSep
	iRowSep = Parent.gRowSep

	With Frm1.vspdData
       For lRow = 1 To .MaxRows
			.Row = lRow
			.Col = C_CHOICE_FG
			Rowchecked = Trim(.Text)
			If Rowchecked = "" Then 
				Rowchecked = 0
			End If
    
			.Row = lRow
	 		.Col = C_CLS_FG						'☜: Update(Yes Check)

			If Trim(.value) = "N" Then
	           .Col = 0
				Select Case .Text
				    Case ggoSpread.InsertFlag
				 											  strVal = strVal & "CREATE"						& iColSep
				 											  strVal = strVal & Rowchecked						& iColSep
				 											  strVal = strVal & UNIConvDate(Trim(frm1.txtBaseDt.text))	& iColSep
				 			.Col = C_LOAN_NO				: strVal = strVal & Trim(.Text)      & iColSep
				 			.Col = C_SEQ					: strVal = strVal & UNICDbl(Trim(.Text))     & iColSep
				 			.Col = C_INT_CLS_AMT			: strVal = strVal & UNICDbl(Trim(.Text))     & iColSep
				 			.Col = C_INT_CLS_LOC_AMT		: strVal = strVal & UNICDbl(Trim(.Text))     & iColSep
				 			.Col = C_INT_EXP_ACCT_CD		: strVal = strVal & Trim(.Text)     & iColSep
				 			.Col = C_PAY_NO					: strVal = strVal & Trim(.Text)		& iColSep
				 			.Col = C_INT_CLS_DT				: strVal = strVal & UNIConvDate(Trim(.Text))		& iColSep
				 			.Col = C_INT_PAY_DT
				 					IF Trim(.Text) <> "" THEN
				 						strVal = strVal & UNIConvDate(Trim(.Text))		& iColSep
				 					ELSE
				 						strVal = strVal & iColSep
				 					END IF
				 			.Col = C_INT_PAY_AMT			: strVal = strVal & UNICDbl(Trim(.Text))		& iColSep
				 			.Col = C_INT_PAY_LOC_AMT		: strVal = strVal & UNICDbl(Trim(.Text))		& iRowSep
				                    
				 			lGrpCnt = lGrpCnt + 1
				    Case ggoSpread.UpdateFlag 
															  strVal = strVal & "C"                       & iColSep
															  strVal = strVal & Rowchecked                & iColSep
															  strVal = strVal & UNIConvDate(Trim(frm1.txtBaseDt.text))	& iColSep
							.Col = C_LOAN_NO				: strVal = strVal & Trim(.Text)				& iColSep
							.Col = C_SEQ					: strVal = strVal & UNICDbl(Trim(.Text))	& iColSep
							.Col = C_INT_CLS_AMT			: strVal = strVal & UNICDbl(Trim(.Text))    & iColSep
							.Col = C_INT_CLS_LOC_AMT		: strVal = strVal & UNICDbl(Trim(.Text))    & iColSep
							.Col = C_INT_EXP_ACCT_CD		: strVal = strVal & Trim(.Text)				& iColSep
				 			.Col = C_PAY_NO					: strVal = strVal & Trim(.Text)				& iRowSep

				 			lGrpCnt = lGrpCnt + 1
				    Case ggoSpread.DeleteFlag
				 											  strVal = strVal & "DELETE"                    & iColSep
				 											  strVal = strVal & lRow						& iColSep
				 											  strVal = strVal & UNIConvDate(Trim(frm1.txtBaseDt.text))	& iColSep
				 			.Col = C_LOAN_NO				: strVal = strVal & Trim(.Text)					& iColSep
				 			.Col = C_SEQ					: strVal = strVal & UNICDbl(Trim(.Text))		& iColSep
				 			.Col = C_INT_CLS_AMT			: strVal = strVal & UNICDbl(Trim(.Text))		& iColSep
				 			.Col = C_INT_CLS_LOC_AMT		: strVal = strVal & UNICDbl(Trim(.Text))		& iColSep
				 			.Col = C_INT_EXP_ACCT_CD		: strVal = strVal & Trim(.Text)     & iColSep
				 			.Col = C_PAY_NO					: strVal = strVal & Trim(.Text)		& iRowSep
										
				 			lGrpCnt = lGrpCnt + 1
				End Select
			Else
				If Rowchecked = "1" Then
													  strVal = strVal & "D"                & iColSep
													  strVal = strVal & lRow                     & iColSep
													  strVal = strVal & UNIConvDate(Trim(frm1.txtBaseDt.text))	& iColSep
					.Col = C_LOAN_NO				: strVal = strVal & Trim(.Text)      & iColSep
					.Col = C_SEQ					: strVal = strVal & UNICDbl(Trim(.Text))     & iColSep
					.Col = C_INT_CLS_AMT			: strVal = strVal & 0						  & iColSep
					.Col = C_INT_CLS_LOC_AMT		: strVal = strVal & 0						  & iColSep
					.Col = C_INT_EXP_ACCT_CD		: strVal = strVal & Trim(.Text)     & iRowSep
						    
					lGrpCnt = lGrpCnt + 1
				End If
			End If
       Next

	   frm1.txtMaxRows.value     = lGrpCnt-1	
	   frm1.txtSpread.value      = strVal

	End With
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    DbSave = True                                                                '☜: Processing is OK

    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
' Name : MonClsProcess
' Desc : 
'========================================================================================================
Function MonClsProcess()

    Dim IntRetCD 
    Dim lRow
    Dim lGrpCnt
    Dim strVal, iColSep, iRowSep
    Dim strIsData
    
    MonClsProcess = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    strVal = ""
    strIsData = False
	iColSep = Parent.gColSep
	iRowSep = Parent.gRowSep
    
	With Frm1.vspdData
       For lRow = 1 To .MaxRows
			.Row = lRow
			.Col = C_CHOICE_FG
			
			If Trim(.Text) = "1" Then
				strIsData = True
				.Col = C_CLS_FG
				If Trim(.value) = "N" Then
					.Col = C_INT_EXP_ACCT_CD
					If Trim(.Text) = "" Then
						Call DisplayMsgBox("970021","x",frm1.txtIntExpAcctCd.Alt,"x")
						Exit Function
					Else
						If CommonQueryRs("A.ACCT_CD","A_Acct A, A_ACCT_GP B, A_JNL_ACCT_ASSN C","A.GP_CD=B.GP_CD AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND A.Acct_cd=C.Acct_CD and C.trans_type = " & FilterVar("FI005", "''", "S") & "  and C.jnl_cd = " & FilterVar("PI", "''", "S") & " " & _
								" AND A.ACCT_CD = " & FilterVar(.Text, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
							Call DisplayMsgBox("971001","x",frm1.txtIntExpAcctCd.Alt,"x")
							Exit Function
						End If
					End If
				End If
			End If
		Next

		If strIsData = False Then
'			Call DisplayMsgBox("900025","x","x","x")
			CALL DBSAVEOK()
			Exit Function
		End If
		
	End With
		
    Call DisableToolBar(Parent.TBC_SAVE)                                                '☜: Disable Save Button Of ToolBar
    Call LayerShowHide(1)                                                        '☜: Show Processing Message
    frm1.btnClsProcess.disabled = true
    frm1.btnClsCancel.disabled = true
		
    Frm1.txtMode.value        = Parent.UID_M0002                                        '☜: Delete

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	With Frm1.vspdData
       For lRow = 1 To .MaxRows
			.Row = lRow
			.Col = C_CHOICE_FG

			If Trim(.Text) = "1" Then
				.Col = C_CLS_FG
				If Trim(.value) = "N" Then
													  strVal = strVal & "C"                       & iColSep
													  strVal = strVal & lRow                      & iColSep
													  strVal = strVal & UNIConvDate(Trim(frm1.txtBaseDt.text))	& iColSep
					.Col = C_LOAN_NO				: strVal = strVal & Trim(.Text)      & iColSep
					.Col = C_SEQ					: strVal = strVal & UNICDbl(Trim(.Text))	& iColSep
					.Col = C_INT_CLS_AMT			: strVal = strVal & UNICDbl(Trim(.Text))     & iColSep
					.Col = C_INT_CLS_LOC_AMT		: strVal = strVal & UNICDbl(Trim(.Text))     & iColSep
					.Col = C_INT_EXP_ACCT_CD		: strVal = strVal & Trim(.Text)     & iRowSep
				        
					lGrpCnt = lGrpCnt + 1
				Else
													  strVal = strVal & "D"                & iColSep
													  strVal = strVal & lRow                     & iColSep
													  strVal = strVal & UNIConvDate(Trim(frm1.txtBaseDt.text))	& iColSep
					.Col = C_LOAN_NO				: strVal = strVal & Trim(.Text)      & iColSep
					.Col = C_SEQ					: strVal = strVal & UNICDbl(Trim(.Text))     & iColSep
					.Col = C_INT_CLS_AMT			: strVal = strVal & 0						  & iColSep
					.Col = C_INT_CLS_LOC_AMT		: strVal = strVal & 0						  & iColSep
					.Col = C_INT_EXP_ACCT_CD		: strVal = strVal & Trim(.Text)     & iRowSep
					    
					lGrpCnt = lGrpCnt + 1
				End If
			End If
       Next
	   frm1.txtMaxRows.value     = lGrpCnt-1	
	   frm1.txtSpread.value      = strVal

	End With
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    Set gActiveElement = document.ActiveElement   
    MonClsProcess = True                                                              '☜: Processing is OK

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
	
	Frm1.vspdData.Focus
	
    If frm1.rdoGiYes.Checked = True THen
		Call SetToolbar("1100100100011111")                                              '☆: Developer must customize
		frm1.btnClsCancel.disabled = false
    Else
		Call SetToolbar("1100111100111111")                                              '☆: Developer must customize
		frm1.btnClsProcess.disabled = false
		call InitData()
    ENd if

	'------ Developer Coding part (End )   -------------------------------------------------------------- 	
	Call ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement   
End Sub
			

'========================================================================================================
' Name : DbSaveOk
' Desc : Called by MB Area when save operation is successful
'========================================================================================================
Sub DbSaveOk()
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	Call ggoOper.ClearField(Document, "2")
    
	Call SetToolbar("1100111100111111")												 '☆: Developer must customize
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If DbQuery("MR") = False Then
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
	MsgBox "You need to code this part.",,Parent.gLogoName 
	'------ Developer Coding part (End)    -------------------------------------------------------------- 
End Function

'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 
'========================================================================================================
' Name : OpenPopup()
' Desc : developer describe this line 
'========================================================================================================
Function OpenPopup(Byval strCode,Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function
	
	Select Case iWhere
	Case 1,2
		If frm1.txtIntExpAcctCd.className = "protected" Then Exit Function    
				
		arrParam(0) = "이자비용계정팝업"								' 팝업 명칭 
		arrParam(1) = "A_Acct A, A_ACCT_GP B, A_JNL_ACCT_ASSN C"				' TABLE 명칭 
		arrParam(2) = strCode											' Code Condition
		arrParam(3) = ""												' Name Cindition
		arrParam(4) = "A.GP_CD=B.GP_CD AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND A.Acct_cd=C.Acct_CD" & _
			" AND C.trans_type = " & FilterVar("FI005", "''", "S") & "  AND C.jnl_cd = " & FilterVar("PI", "''", "S") & " "
		arrParam(5) = frm1.txtIntExpAcctCd.Alt							' 조건필드의 라벨 명칭 

		arrField(0) = "A.Acct_CD"									' Field명(0)
		arrField(1) = "A.Acct_NM"									' Field명(1)
		arrField(2) = "B.GP_CD"										' Field명(2)
		arrField(3) = "B.GP_NM"										' Field명(3)
				
		arrHeader(0) = frm1.txtIntExpAcctCd.Alt									' Header명(0)
		arrHeader(1) = frm1.txtIntExpAcctNm.Alt								' Header명(1)
		arrHeader(2) = "그룹코드"									' Header명(2)
		arrHeader(3) = "그룹명"										' Header명(3)

	Case 3
		arrParam(0) = "차입금번호팝업"								' 팝업 명칭 
		arrParam(1) = "F_LN_INFO B "				' TABLE 명칭 
		arrParam(2) = strCode											' Code Condition
		arrParam(3) = ""												' Name Cindition
		arrParam(4) = "B.INT_PAY_STND = " & FilterVar("AI", "''", "S") & "  AND B.CONF_FG IN (" & FilterVar("C", "''", "S") & " ," & FilterVar("E", "''", "S") & " ) AND B.LOAN_DT <= " & FilterVar(UNICONVDATE(frm1.txtBaseDt.TEXT), "''", "S") & _
					" AND B.LOAN_BASIC_DT <= " & FilterVar(UNICONVDATE(frm1.txtBaseDt.TEXT), "''", "S")
		arrParam(5) = frm1.txtLOAN_NO.Alt							' 조건필드의 라벨 명칭 

		arrField(0) = "B.LOAN_NO"									' Field명(0)
		arrField(1) = "B.LOAN_NM"									' Field명(1)
		arrField(2) = "CASE WHEN B.LOAN_PLC_TYPE = " & FilterVar("BK", "''", "S") & "  THEN B.LOAN_BANK_CD ELSE B.BP_CD END"
		arrField(3) = "B.DOC_CUR"
		arrField(4) = "F5" & parent.gColSep & "B.XCH_RATE"
		arrField(5) = "DD" & parent.gColSep & "B.LOAN_DT"	'"B.INT_ACCT_CD"
		arrField(6) = "DD" & parent.gColSep & "B.DUE_DT"										' Field명(3)
				
		arrHeader(0) = frm1.txtLOAN_NO.Alt
		arrHeader(1) = frm1.txtLOAN_NM.Alt
		arrHeader(2) = "차입처"
		arrHeader(3) = "거래통화"
		arrHeader(4) = "환율"
		arrHeader(5) = "차입일"
		arrHeader(6) = "만기일"

	Case 4
		arrParam(0) = "상환번호팝업"								' 팝업 명칭 
		arrParam(1) = "F_LN_REPAY_ITEM A, F_LN_INFO B"				' TABLE 명칭 
		arrParam(2) = strCode											' Code Condition
		arrParam(3) = ""												' Name Cindition
		arrParam(4) = "B.LOAN_NO = A.LOAN_NO AND A.PAY_OBJ = " & FilterVar("AI", "''", "S") & "  AND A.CONF_FG IN (" & FilterVar("C", "''", "S") & " ," & FilterVar("E", "''", "S") & " ) AND B.CONF_FG IN (" & FilterVar("C", "''", "S") & " ," & FilterVar("E", "''", "S") & " )"
		frm1.vspdData.Col  = C_LOAN_NO
		If Trim(frm1.vspdData.Text) <> "" Then
			arrParam(4) = arrParam(4) & " AND A.LOAN_NO = " & Filtervar(Trim(frm1.vspdData.Text), "''", "S")
		End If
		arrParam(5) = "상환번호"							' 조건필드의 라벨 명칭 

		arrField(0) = "A.PAY_NO"									' Field명(0)
		arrField(1) = "DD" & parent.gColSep & "A.PAY_DT"									' Field명(1)
		arrField(2) = "F2" & parent.gColSep & "A.PAY_AMT"										' Field명(2)
		arrField(3) = "F2" & parent.gColSep & "A.PAY_LOC_AMT"										' Field명(3)
'		arrField(4) = "A.LOAN_NO"
				
		arrHeader(0) = "상환번호"									' Header명(0)
		arrHeader(1) = "이자지급일자"								' Header명(1)
		arrHeader(2) = "이자지급액"									' Header명(2)
		arrHeader(3) = "이자지급액(자국)"										' Header명(3)
'		arrHeader(4) = frm1.txtLOAN_NO.Alt

	End Select
	
	IsOpenPop = True

	Select Case iWhere
	Case 1,2
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Case 3,4
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			 "dialogWidth=620px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End Select

	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
			Case 1
				frm1.txtIntExpAcctCd.focus
			Case 2,3,4
				frm1.vspdData.focus
		End Select
				
		Exit Function
	Else
		With Frm1
			Select Case iWhere
				Case 1
					.txtIntExpAcctCd.value = arrRet(0)
					.txtIntExpAcctNm.value = arrRet(1)
					.txtIntExpAcctCd.focus
				Case 2
					.vspdData.Col  = C_INT_EXP_ACCT_CD
					.vspdData.Text = arrRet(0)
					.vspdData.Col  = C_INT_EXP_ACCT_NM
					.vspdData.Text = arrRet(1)

				Case 3
					.vspdData.Col  = C_LOAN_NO
					.vspdData.Text = arrRet(0)
					.vspdData.Col  = C_LOAN_NM
					.vspdData.Text = arrRet(1)
					.vspdData.Col  = C_DOC_CUR
					.vspdData.Text = arrRet(3)
					.vspdData.Col  = C_XCH_RATE
					.vspdData.Text = arrRet(4)
					.vspdData.Col  = C_LOAN_DT
					.vspdData.Text = arrRet(5)
					.vspdData.Col  = C_DUE_DT
					.vspdData.Text = arrRet(6)

				Case 4
'					.vspdData.Col  = C_LOAN_NO
'					.vspdData.Text = arrRet(4)
					.vspdData.Col  = C_PAY_NO
					.vspdData.Text = arrRet(0)
					.vspdData.Col  = C_INT_PAY_DT
					.vspdData.Text = arrRet(1)
					.vspdData.Col  = C_INT_PAY_AMT
					.vspdData.Text = arrRet(2)
					.vspdData.Col  = C_INT_PAY_LOC_AMT
					.vspdData.Text = arrRet(3)
			End Select
		End With

	End If	

End Function

 '------------------------------------------ OpenLoanNo() -------------------------------------------------
'	Name : OpenLoanNo()
'	Description : Loan Number PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopupLoan()
	Dim arrRet
	Dim arrParam(3)	
	Dim iCalledAspName
	
	iCalledAspName = AskPRAspName("f4240ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "f4240ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	arrRet = window.showModalDialog(iCalledAspName & "?PGM=" & gStrRequestMenuID, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False
	
	If arrRet(0) = ""  Then
		frm1.txtLOAN_NO.focus
		Exit Function
	Else		
		frm1.txtLOAN_NO.value = arrRet(0)
		frm1.txtLOAN_NM.value = arrRet(1)
		frm1.txtLOAN_NO.focus
	End If
End Function

'============================================================
'회계전표 팝업 
'============================================================
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

	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	frm1.txtBaseDt.focus
End Function

'============================================================
'결의전표 팝업 
'============================================================
Function OpenPopupTempGL()

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
	End With									'Reference번호 

	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	frm1.txtBaseDt.focus
End Function
'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'==========================================================================================
'   Event Name : txtIntExpAcctCd_onChange
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub txtIntExpAcctCd_onChange()
	If Trim(frm1.txtIntExpAcctCd.value) = "" Then
		frm1.txtIntExpAcctNm.value = ""
	End If
End Sub

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

    With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		
		.Col = Col
		.Row = Row
		If Row > 0 Then
			Select Case Col
			Case C_INT_EXP_ACCT_BT
				Call OpenPopup(.Text, 2)
			Case C_LOAN_BT
				Call OpenPopup(.Text, 3)
			Case C_PAY_BT
				Call OpenPopup(.Text, 4)
			End Select
			.row = row
			.col = 0

			If col = C_CHOICE_FG and ggoSpread.InsertFlag <> Trim(.text) and ggoSpread.DeleteFlag <> Trim(.text) then
				.Row = Row
				.Col = C_CHOICE_FG

				If .Text = "Y" Then
					If ButtonDown = 0 Then
						ggoSpread.UpdateRow Row
					Else
						ggoSpread.SSDeleteFlag Row,Row
					End If
				Else
					If ButtonDown = 1 Then
						ggoSpread.UpdateRow Row
					Else
						ggoSpread.SSDeleteFlag Row,Row
					End If			
				End If
			end if
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
    If Col <= C_CHOICE_FG Or NewCol <= C_CHOICE_FG Then
        Cancel = True
        Exit Sub
    End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData

	With frm1.vspdData

		.Row = Row
		Select Case Col
		Case C_INT_EXP_ACCT_CD
			.Col = C_INT_EXP_ACCT_NM
			.Text = ""
		Case C_LOAN_NO
			.Col = C_LOAN_NM	: .Text = ""
			.Col = C_PAY_NO		: .Text = ""
			.Col = C_INT_PAY_DT	: .Text = ""
			.Col = C_INT_PAY_AMT		: .Text = ""
			.Col = C_INT_PAY_LOC_AMT	: .Text = ""

		End Select

	End With
		

End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("1111111111")

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

Sub fpBaseDt_DblClick(Button)
	if Button = 1 then
		frm1.fpBaseDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.fpBaseDt.Focus
	End if
End Sub


Sub fpBaseDt_KeyPress(Key)
    If key = 13 Then
        Call MainQuery
	End If
End Sub

Function AcctApply()
	Dim lRow
	With Frm1.vspdData
       For lRow = 1 To .MaxRows
			.Row = lRow

			.Col = C_CHOICE_FG
			If Trim(.Text) = "1" Then
				.Col = C_CLS_FG
				If Trim(.value) = "N" Then
					.Col = C_INT_EXP_ACCT_CD
					If Trim(.Text) = "" Then
						.Col = C_INT_EXP_ACCT_CD
						.Text = Trim(frm1.txtIntExpAcctCd.value)
						.Col = C_INT_EXP_ACCT_NM
						.Text = Trim(frm1.txtIntExpAcctNM.value)
					End If
				End If
			End If
		Next
		
	End With

End Function

Function Rowcancel()
	Dim lRow

	With Frm1.vspdData
		For lRow = 1 To .MaxRows
			.Row = lRow
			.COL = 0
			IF Trim(.TEXT) = ggoSpread.UPDATEFlag OR Trim(.TEXT) = ggoSpread.INSERTFlag THEN
				.Col = C_CHOICE_FG
				.Text = "0"
				IF Trim(.TEXT) = ggoSpread.UPDATEFlag THEN
					ggoSpread.SSDeleteFlag lRow,lRow
				END IF
			END IF
		Next
	End With
End Function

Function Rowselect()
	Dim lRow

	With Frm1.vspdData
		For lRow = 1 To .MaxRows
			.Row = lRow
			.COL = 0
			IF Trim(.TEXT) <> ggoSpread.DELETEFlag THEN
				.Col = C_CHOICE_FG
				If .Lock = False Then
					.Col = C_CHOICE_FG
					.Text = "1"
					ggoSpread.UpdateRow lRow
				End If
			END IF
		Next
	End With
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>선급이자월결산</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						  </TABLE>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH=* ALIGN=RIGHT>
					 			<A HREF="VBSCRIPT:OpenPopupTempGL()">결의전표</A>&nbsp;|
					 			<A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A></TD> 
					<TD WIDTH=10>&nbsp;</TD>
					</TD>
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
						<FIELDSET>
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>결산일자</TD>
									<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/f4240ma1_fpBaseDt_txtBaseDt.js'></script>
									<TD CLASS="TD5" NOWRAP>차입금번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtLOAN_NO" MAXLENGTH="18" SIZE=15  ALT ="차입금번호" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoanNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopupLoan()">
														   <INPUT NAME="txtLOAN_NM" MAXLENGTH="40" SIZE=20  ALT ="차입금내역" tag="14"></TD>
								</TR>
								<TR>	
									<TD CLASS=TD5 NOWRAP>결산여부</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE=radio CLASS = "RADIO" NAME="rdoGiFlag" id="rdoGiYes" VALUE="Y" tag = "11" >
											<LABEL FOR="rdoGiYes">결산</LABEL>&nbsp;&nbsp;	
										<INPUT TYPE=radio CLASS = "RADIO" NAME="rdoGiFlag" id="rdoGiNo" VALUE="N" tag = "11" CHECKED>
										<LABEL FOR="rdoGiNo">미결산</LABEL></TD>
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
					<TD HEIGHT=20 WIDTH="100%">
						<FIELDSET>
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>이자비용계정</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtIntExpAcctCd" MAXLENGTH="18" SIZE=15  ALT ="이자비용계정" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAdvInt" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtIntExpAcctCd.value, 1)">
														   <INPUT NAME="txtIntExpAcctNm" MAXLENGTH="40" SIZE=20  ALT ="이자비용계정명" tag="14"> &nbsp;
														   <BUTTON NAME="btnApply" style="height:20px" CLASS="CLSSBTN" ONCLICK="vbscript:AcctApply()">적용</BUTTON></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP><BUTTON NAME="btncancel" style="height:20px" CLASS="CLSSBTN" ONCLICK="vbscript:Rowselect()">전체선택</BUTTON>
														<BUTTON NAME="btnselect" style="height:20px" CLASS="CLSSBTN" ONCLICK="vbscript:Rowcancel()">전체취소</BUTTON></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH="100%" HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/f4240ma1_vaSpread1_vspdData.js'></script>
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
									<TD CLASS="TD5" NOWRAP>결산총액|자국</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/f4240ma1_fpDoubleSingle11_txtAlcSum.js'></script>&nbsp;
										<script language =javascript src='./js/f4240ma1_fpDoubleSingle12_txtAlcLocSum.js'></script>&nbsp;
	                                </TD>
									<TD CLASS="TD5" NOWRAP>결산계획총액|자국</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/f4240ma1_fpDoubleSingle21_txtPlanSum.js'></script>&nbsp;
										<script language =javascript src='./js/f4240ma1_fpDoubleSingle22_txtPlanLocSum.js'></script>&nbsp;
	                                </TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20 style="display:none;">
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnClsProcess" CLASS="CLSSBTN" ONCLICK="vbscript:MonClsProcess()">결산</BUTTON>&nbsp;
						<BUTTON NAME="btnClsCancel" CLASS="CLSSBTN" ONCLICK="vbscript:MonClsProcess()">결산취소</BUTTON>&nbsp;
					</TD>
					<TD WIDTH=* ALIGN=RIGHT></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows"      TAG="24">
<INPUT TYPE=HIDDEN NAME="txtMode"         TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"    TAG="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

