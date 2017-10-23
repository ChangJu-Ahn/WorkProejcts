<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--
======================================================================================================
*  1. Module Name          : Account
*  2. Function Name        : 
*  3. Program ID           : A5406ma1 
*  4. Program Name         : 미결반제(신용카드)
*  5. Program Desc         : Multi Sample
*  6. Comproxy List        :
*  7. ModIfied date(First) : 2002/11/05
*  8. ModIfied date(Last)  : 2003/08/04
*  9. ModIfier (First)     :
* 10. ModIfier (Last)      : Jeong Yong Kyun
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
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="javascript"	SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../ag/AcctCtrl.vbs">							</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../ag/AcctCtrl2.vbs">							</SCRIPT>

<SCRIPT LANGUAGE="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID = "A5407MB1.asp"			'☆: 비지니스 로직 ASP명 

Const TAB1 = 1																		'☜: Tab의 위치
Const TAB2 = 2

Const C_MENU_NEW_TAB1 = "1110000000011111"
Const C_MENU_NEW_TAB2 = "1110010000011111"
Const C_MENU_CRT_TAB1 =	"1110100100111111"
Const C_MENU_CRT_TAB2 =	"1110111100111111"
Const C_MENU_UPD_TAB1 =	"1111000000011111"
Const C_MENU_UPD_TAB2 =	"1111000000011111"

Const C_CONDFIELD = 0
Const C_VSPDDATA1 = 1
Const C_vspdData2 = 2
Const C_VSPDDATA3 = 3
Const C_VSPDDATA4 = 4
Const C_VSPDDATA5 = 5
Const C_VSPDDATA6 = 6

Const C_GLINPUTTYPE = "OC"

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
'Const C_MaxKey            = 17                                    '☆☆☆☆: Max key value
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
'☜:--------Spreadsheet #1-----------------------------------------------------------------------------   
'@Grid_Column
Dim C_ItemSeq1
Dim C_CardNo1
Dim C_UnSettCd1
Dim C_User1
Dim C_CardCo1
Dim C_GL_NO1
Dim C_GL_DT1
Dim C_DOC_CUR1
Dim C_XCH_RATE1
Dim C_OPEN_AMT1
Dim C_OPEN_DOC_AMT1
Dim C_DEPT_CD1
Dim C_DEPT_NM1
Dim C_ACCT_CD1
Dim C_ACCT_NM1
Dim C_DR_CR_FG1
Dim C_DR_CR_NM1
Dim C_OpenGlItemSeq
Dim C_InternalCd
Dim C_Costcd
Dim C_OrgchangeId

'☜:--------Spreadsheet #2-----------------------------------------------------------------------------   
Dim C_ItemSeq	    																	'☆: Spread Sheet 의 Columns 인덱스 
Dim C_deptcd_2
Dim C_deptPopup_2	
Dim C_deptnm_2	   	
Dim C_AcctCd		
Dim C_AcctPopup      
Dim C_AcctNm            
Dim C_DrCrFg      
Dim C_DrCrNm_2 
Dim C_DocCur_2 
Dim C_DocCurPopup_2
Dim C_ExchRate_2    
Dim C_ItemAmt_2	
Dim C_ItemLocAmt_2		
Dim C_ItemDesc_2     
Dim C_IsLAmtChange_2  
Dim C_OpenGlNo_2	
Dim C_OpenGlItemSeq_2	
Dim C_MgntFg_2		
Dim C_AcctCd2_2	

'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim gSelframeFlg

Dim lgcurrrow
Dim IsOpenPop  
Dim lgBlnExecDelete
Dim lgCurrentTabFg
Dim lgIntMaxItemSeq
Dim lgPreToolBarTab1
Dim lgPreToolBarTab2
Dim lgQueryOk

Dim BaseDate, LastDate, FirstDate
Dim FromDateOfDB, ToDateOfDB

' 권한관리 추가
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인
                                                 
   BaseDate     = "<%=GetSvrDate%>"																	'Get DB Server Date

   LastDate     = UNIGetLastDay (BaseDate,parent.gServerDateFormat)                                 'Last  day of this month
   FirstDate    = UNIGetFirstDay(BaseDate,parent.gServerDateFormat)                                 'First day of this month

   FromDateOfDB = UNIDateAdd("m", -20, BaseDate,parent.gServerDateFormat)
   ToDateOfDB   = UNIDateAdd("m",  40, BaseDate,parent.gServerDateFormat)
 
   FromDateOfDB  = UniConvDateAToB(FromDateOfDB ,parent.gServerDateFormat,parent.gDateFormat)       'Convert DB date type To Company
   ToDateOfDB    = UniConvDateAToB(ToDateOfDB   ,parent.gServerDateFormat,parent.gDateFormat)       'Convert DB date type To Company

'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================
Sub initSpreadPosVariables(ByVal pvSpdNo)
	Select Case  UCase(Trim(pvSpdNo))
		Case  "A"
            C_ItemSeq1        = 1
			C_CardNo1	      = 2
			C_UnSettCd1	      = 3
			C_User1		      = 4
			C_CardCo1         = 5
			C_GL_NO1	      = 6
			C_GL_DT1	      = 7
			C_DOC_CUR1	      = 8
			C_XCH_RATE1	      = 9
			C_OPEN_DOC_AMT1	  = 10	
			C_OPEN_AMT1	      = 11
			C_DEPT_CD1	      = 12
			C_DEPT_NM1	      = 13
			C_ACCT_CD1	      = 14
			C_ACCT_NM1	      = 15
			C_DR_CR_FG1	      = 16
			C_DR_CR_NM1		  = 17
			C_OpenGlItemSeq   = 18
			C_internalcd	  = 19
			C_costcd		  = 20
			C_orgchangeid	  = 21
		Case "B"
			C_ItemSeq		  = 1 																	'☆: Spread Sheet 의 Columns 인덱스 
			C_deptcd_2		  = 2
			C_deptPopup_2	  = 3
			C_deptnm_2		  = 4
			C_AcctCd		  = 5
			C_AcctPopup   	  = 6
			C_AcctNm      	  = 7
			C_DrCrFg		  = 8
			C_DrCrNm_2		  = 9
			C_DocCur_2		  = 10
			C_DocCurPopup_2   = 11
			C_ExchRate_2	  = 12
			C_ItemAmt_2		  = 13
			C_ItemLocAmt_2	  = 14	
			C_ItemDesc_2	  = 15
			C_IsLAmtChange_2  = 16
			C_OpenGlNo_2	  = 17
			C_OpenGlItemSeq_2 = 18
			C_MgntFg_2		  = 19
			C_AcctCd2_2		  = 20
	End Select
End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
    lgBlnFlgChgValue = False										'Indicates that no value changed
    lgIntFlgMode     = parent.OPMD_CMODE							'Indicates that current mode is Create mode
    lgIntGrpCount = 0												'initializes Group View Size
    lgStrPrevKey = ""												'initializes Previous Key
    lgLngCurRows = 0  
    lgIntMaxItemSeq = 0
    lgCurrentTabFg = TAB1  
    gSelframeFlg = TAB1    
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
    frm1.txtClsDt.text = UniConvDateAToB("<%=GetSvrDate%>" ,Parent.gServerDateFormat,Parent.gDateFormat)
    
    Call ggoOper.ClearField(Document, "1")							'⊙: Condition field clear
    
    frm1.cboGlType.value = "03"
    frm1.cboGlInputType.value = C_GLINPUTTYPE
	frm1.txtDeptCd.value	= Parent.gDepart
	frm1.hOrgChangeId.value = Parent.gChangeOrgId

	Call GetCheckAcct()
	
	frm1.txtClsNo.focus
	lgBlnFlgChgValue = False										'Indicates that no value changed
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number Format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
	<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
End Sub

'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub  InitSpreadSheet(ByVal pvSpdNo)
	Call initSpreadPosVariables(pvSpdNo)

	With frm1
		Select Case UCase(Trim(pvSpdNo))
			Case  "A"
				.vspdData.MaxCols = C_orgchangeid + 1 
				.vspdData.Col = .vspdData.MaxCols				'☜: 공통콘트롤 사용 Hidden Column  
				.vspdData.ColHidden = True
					
				ggoSpread.Source = .vspdData
				ggoSpread.Spreadinit "V20021127",,parent.gAllowDragDropSpread  
				ggoSpread.ClearSpreadData   
				.vspdData.ReDraw = False

				Call GetSpreadColumnPos(pvSpdNo)

				ggoSpread.SSSetFloat C_ItemSeq1     , " "                 ,  4, "6", ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
				ggoSpread.SSSetEdit  C_CardNo1      , "카드번호"      , 15, , , 30
				ggoSpread.SSSetEdit  C_UnSettCd1    , "미결코드2"     , 15,3
				ggoSpread.SSSetEdit  C_User1        , "사용자"        , 12,3	'
				ggoSpread.SSSetEdit  C_CardCo1      , "카드사"        , 12, , , 30
				ggoSpread.SSSetEdit  C_GL_NO1       , "전표번호"      , 10,3
				ggoSpread.SSSetDate  C_GL_DT1       , "발생일"		  , 12,2                  ,Parent.gDateFormat   
				ggoSpread.SSSetEdit  C_DOC_CUR1     , "거래통화"      , 10,3	'
				ggoSpread.SSSetFloat C_XCH_RATE1    , "환율"          , 15, parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_OPEN_DOC_AMT1, "발생금액"      , 15, parent.ggAmTofMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_OPEN_AMT1    , "발생금액(자국)", 15, parent.ggAmTofMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetEdit  C_DEPT_CD1     , "부서코드"      , 10,3	'
				ggoSpread.SSSetEdit  C_DEPT_NM1     , "부서명"        , 10,3	'
				ggoSpread.SSSetEdit  C_ACCT_CD1     , "계정코드"      , 10,3	'
				ggoSpread.SSSetEdit  C_ACCT_NM1     , "계정명"        , 10,3	'
				ggoSpread.SSSetCombo C_DR_CR_FG1    , ""                  ,  8
				ggoSpread.SSSetCombo C_DR_CR_NM1    , "차대구분"      , 10,3	'
				ggoSpread.SSSetEdit  C_OpenGlItemSeq, "순서"          , 10,3	'			
				ggoSpread.SSSetEdit  C_InternalCd,    "내부부서코드",     10
				ggoSpread.SSSetEdit  C_Costcd,		  "코스트센터",       10
				ggoSpread.SSSetEdit  C_OrgchangeId,   "조직변경ID",       10
				
				Call ggoSpread.SSSetColHidden(C_DR_CR_FG1,C_DR_CR_FG1,True)								'차대구분 Hidden Column		
				Call ggoSpread.SSSetColHidden(C_ItemSeq1  ,C_ItemSeq1	,True)							'순서
				Call ggoSpread.SSSetColHidden(C_OpenGlItemSeq  ,C_OpenGlItemSeq	,True)                  '전표순서    				                      
				Call ggoSpread.SSSetColHidden(C_InternalCd,C_InternalCd,True)								'차대구분 Hidden Column		
				Call ggoSpread.SSSetColHidden(C_Costcd  ,C_Costcd	,True)							'순서
				Call ggoSpread.SSSetColHidden(C_OrgchangeId  ,C_OrgchangeId	,True)                  '전표순서    				                      
								
				

				Call SetSpreadLock ("I", 0, 1, "" )
			Case  "B"	
				.vspdData4.MaxCols = C_AcctCd2_2 + 1 												'☜: 최대 Columns의 항상 1개 증가시킴
				.vspdData4.Col = .vspdData4.MaxCols													
				.vspdData4.ColHidden = True															'공통콘트롤 사용 Hidden Column

				ggoSpread.Source = .vspdData4
				ggoSpread.Spreadinit "V20030324",,parent.gAllowDragDropSpread 
				ggoSpread.ClearSpreadData 
					
				.vspdData4.ReDraw = False   
				
				Call GetSpreadColumnPos(pvSpdNo)
				Call AppEndNumberPlace("6","3","0")
				
				ggoSpread.SSSetFloat  C_ItemSeq       ," "                   ,  4,"6", ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
				ggoSpread.SSSetEdit   C_deptcd_2      ,"부서코드"        , 10, , , 10
				ggoSpread.SSSetButTon C_deptPopup_2
				ggoSpread.SSSetEdit   C_deptnm_2      ,"부서명"          , 17, , , 30
				ggoSpread.SSSetEdit   C_AcctCd        ,"계정코드"        , 15, , , 18
				ggoSpread.SSSetButTon C_AcctPopup   
				ggoSpread.SSSetEdit   C_AcctNm        ,"계정코드명"      , 20, , , 30
				ggoSpread.SSSetCombo  C_DrCrFg        ," "                   ,  8
				ggoSpread.SSSetCombo  C_DrCrNm_2      ,"차대구분"        , 10
				ggoSpread.SSSetEdit   C_DocCur_2      ,"거래통화"        , 10, , , 10, 2
				ggoSpread.SSSetButTon C_DocCurPopup_2
				ggoSpread.SSSetFloat  C_ExchRate_2    ,"환율"            , 15, Parent.ggAmTofMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
				ggoSpread.SSSetFloat  C_ItemAmt_2     ,"금액"            , 15, Parent.ggAmTofMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
				ggoSpread.SSSetFloat  C_ItemLocAmt_2  ,"금액(자국)"      , 15, Parent.ggAmTofMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
				ggoSpread.SSSetEdit   C_ItemDesc_2    ,"비고"            , 30, , , 128
				ggoSpread.SSSetEdit   C_IsLAmtChange_2," "                   , 10		

				ggoSpread.SSSetEdit  C_OpenGlNo_2     ,"미결전표번호"    , 10
				ggoSpread.SSSetEdit  C_OpenGlItemSeq_2,"미결전표항목번호", 10								'C_MgntFg
				ggoSpread.SSSetEdit  C_MgntFg_2       ,"미결계정여부"    , 10
				ggoSpread.SSSetEdit  C_AcctCd2_2      ,"계정코드2"       , 10
				
				call ggoSpread.MakePairsColumn(C_deptcd_2,	C_deptPopup_2)
				call ggoSpread.MakePairsColumn(C_AcctCd,	C_AcctPopup)
						
				Call ggoSpread.SSSetColHidden(C_ItemSeq			,C_ItemSeq			,True)						'☜: 차대구분 Hidden Column
				Call ggoSpread.SSSetColHidden(C_DrCrFg			,C_DrCrFg			,True)						'☜: 차대구분 Hidden Column
				Call ggoSpread.SSSetColHidden(C_IsLAmtChange_2	,C_IsLAmtChange_2	,True)						'☜: 사용자가 Local 금액을 직접입력하였는지 
				Call ggoSpread.SSSetColHidden(C_OpenGlNo_2		,C_OpenGlNo_2		,True)
				Call ggoSpread.SSSetColHidden(C_OpenGlItemSeq_2	,C_OpenGlItemSeq_2	,True)
				Call ggoSpread.SSSetColHidden(C_MgntFg_2		,C_MgntFg_2			,True)
				Call ggoSpread.SSSetColHidden(C_AcctCd2_2		,C_AcctCd2_2		,True)                      '☜: 차대구분 Hidden Column  

				.vspdData4.ReDraw = True

				Call SetSpread4Lock ("I", 0, 1, "" )
		End Select
    End With
End Sub

'=======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadLock(Byval stsFg, Byval Index, ByVal lRow  , ByVal lRow2 )
    Dim objSpread
    
    With frm1
		ggoSpread.Source = .vspdData
		Set objSpread = .vspdData
		lRow2 = objSpread.MaxRows
		objSpread.Redraw = False
		Select Case  Index
			Case  0
				ggoSpread.SpreadLock C_ItemSeq1			,lRow,	C_ItemSeq1    
				ggoSpread.SpreadLock C_CardNo1			,lRow,	C_CardNo1    
				ggoSpread.SpreadLock C_UnSettCd1		,lRow,	C_UnSettCd1  
		        ggoSpread.SpreadLock C_User1			,lRow,	C_User1    
				ggoSpread.SpreadLock C_CardCo1			,lRow,	C_CardCo1    
		        ggoSpread.SpreadLock C_GL_NO1			,lRow,	C_GL_NO1  
				ggoSpread.SpreadLock C_GL_DT1			,lRow,	C_GL_DT1 
				ggoSpread.SpreadLock C_DOC_CUR1			,lRow,	C_DOC_CUR1 
				ggoSpread.SpreadLock C_XCH_RATE1        ,lRow,	C_XCH_RATE1 
				ggoSpread.SpreadLock C_OPEN_DOC_AMT1    ,lRow,	C_OPEN_DOC_AMT1 
				ggoSpread.SpreadLock C_OPEN_AMT1        ,lRow,	C_OPEN_AMT1 
				ggoSpread.SpreadLock C_DEPT_CD1			,lRow,	C_DEPT_CD1 
				ggoSpread.SpreadLock C_DEPT_NM1			,lRow,	C_DEPT_NM1 
				ggoSpread.SpreadLock C_ACCT_CD1			,lRow,	C_ACCT_CD1 
				ggoSpread.SpreadLock C_ACCT_NM1			,lRow,	C_ACCT_NM1 
'				ggoSpread.SpreadLock C_DR_CR_FG1        ,lRow,	C_DR_CR_FG1
				ggoSpread.SpreadLock C_DR_CR_NM1        ,lRow,	C_DR_CR_NM1 
				ggoSpread.SpreadLock C_OpenGlItemSeq    ,lRow,	C_OpenGlItemSeq   
				ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
			Case  1
				ggoSpread.SpreadLock C_ItemSeq1, lRow, C_OpenGlItemSeq	'Item Grid 전체 Lock설정 
				ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
		End Select

		objSpread.Redraw = True
		Set objSpread = Nothing
    End With
End Sub

'=======================================================================================================
' Function Name : SetSpread2Lock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpread2Lock(Byval stsFg, Byval Index, ByVal lRow  , ByVal lRow2 )
	Dim objSpread
    
    With frm1
		ggoSpread.Source = .vspdData2
		Set objSpread = .vspdData2

		lRow2 = objSpread.MaxRows
		objSpread.Redraw = False
	
		Select Case Index
			Case  0			
				ggoSpread.SSSetProtected	.vspdData5.MaxCols,-1,-1
			Case  1
				ggoSpread.SpreadLock 1, lRow, objSpread.MaxCols, lRow2	
				ggoSpread.SSSetProtected	.vspdData5.MaxCols,-1,-1
		End Select

		objSpread.Redraw = True
		Set objSpread = Nothing
    End With
End Sub

'=======================================================================================================
' Function Name : SetSpread4Lock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpread4Lock(Byval stsFg, Byval Index, ByVal lRow  , ByVal lRow2 )
    Dim objSpread
    
    With frm1
		ggoSpread.Source = .vspdData4
		Set objSpread = .vspdData4
		lRow2 = objSpread.MaxRows
		objSpread.Redraw = False
		Select Case    Index
			Case 0			
				ggoSpread.SpreadLock C_AcctNm		,lRow	,C_AcctNm
				ggoSpread.SpreadLock C_AcctPopup	,lRow	,C_AcctPopup
		        ggoSpread.SpreadLock C_AcctCd		,lRow	,C_AcctCd
		        ggoSpread.SpreadLock C_deptnm_2		,lRow	,C_deptnm_2
				ggoSpread.SSSetProtected	.vspdData4.MaxCols,-1,-1
			Case 1
				ggoSpread.SpreadLock C_ItemSeq		,lRow	,C_AcctCd2_2	'Item Grid 전체 Lock설정 
				ggoSpread.SSSetProtected	.vspdData4.MaxCols,-1,-1
		End Select

		objSpread.Redraw = True
		Set objSpread = Nothing
    End With
End Sub

'=======================================================================================================
' Function Name : SetSpread5Lock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpread5Lock(Byval stsFg, Byval Index, ByVal lRow  , ByVal lRow2 )
	Dim objSpread
    
    With frm1
		ggoSpread.Source = .vspdData5
		Set objSpread = .vspdData5
		lRow2 = objSpread.MaxRows
		objSpread.Redraw = False
	
		Select Case Index
			Case  0
				ggoSpread.SpreadLock 1, lRow, objSpread.MaxCols, lRow2				
				ggoSpread.SSSetProtected	.vspdData5.MaxCols,-1,-1			
			Case  1
'				ggoSpread.SpreadLock 1, lRow, objSpread.MaxCols, lRow2	
				ggoSpread.SSSetProtected	.vspdData5.MaxCols,-1,-1
		End Select

		objSpread.Redraw = True
		Set objSpread = Nothing
    End With
End Sub

'================================== 2.2.5 SetSpreadColor() ==============================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(Byval stsFg, Byval Index, ByVal pvStartRow, ByVal pvEndRow)
	If  pvEndRow = "" Then pvEndRow = pvStartRow
	
    With frm1
        If Index = "0" Then
			ggoSpread.Source = .vspddata
			.vspdData.ReDraw = False

			ggoSpread.SSSetProtected	C_ItemSeq1,			pvStartRow, pvEndRow
			ggoSpread.SSSetProtected	C_CardNo1,			pvStartRow, pvEndRow
			ggoSpread.SSSetProtected	C_UnSettCd1,		pvStartRow, pvEndRow
			ggoSpread.SSSetProtected	C_User1,			pvStartRow, pvEndRow
            ggoSpread.SSSetProtected	C_CardCo1,			pvStartRow, pvEndRow 
    		ggoSpread.SSSetProtected	C_GL_NO1,			pvStartRow, pvEndRow
			ggoSpread.SSSetProtected	C_GL_DT1,			pvStartRow, pvEndRow
			ggoSpread.SSSetProtected	C_DOC_CUR1,			pvStartRow, pvEndRow
			ggoSpread.SSSetProtected	C_XCH_RATE1,		pvStartRow, pvEndRow
			ggoSpread.SSSetProtected	C_OPEN_AMT1,		pvStartRow, pvEndRow
			ggoSpread.SSSetProtected	C_OPEN_DOC_AMT1,	pvStartRow, pvEndRow
			
			ggoSpread.SSSetProtected	C_DEPT_CD1,			pvStartRow, pvEndRow
			ggoSpread.SSSetProtected	C_DEPT_NM1,			pvStartRow, pvEndRow
			ggoSpread.SSSetProtected	C_ACCT_CD1,			pvStartRow, pvEndRow
			ggoSpread.SSSetProtected	C_ACCT_NM1,			pvStartRow, pvEndRow
			ggoSpread.SSSetProtected	C_DR_CR_NM1,		pvStartRow, pvEndRow
			ggoSpread.SSSetProtected	C_OpenGlItemSeq,	pvStartRow, pvEndRow
			
			.vspdData.ReDraw = True
		End If		
    End With
End Sub

'================================== 2.2.5 SetSpreadColor2() ==============================
' Function Name : SetSpreadColor2
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor2(Byval stsFg, Byval Index, ByVal pvStartRow,ByVal pvEndRow)
    With frm1
		If  pvEndRow = "" Then	pvEndRow = pvStartRow

		ggoSpread.Source = .vspdData4	
		.vspdData4.ReDraw = False
		ggoSpread.SSSetProtected	C_ItemSeq   ,pvStartRow	,pvEndRow	'
		ggoSpread.SSSetProtected	C_deptnm_2  ,pvStartRow	,pvEndRow
		ggoSpread.SSSetProtected	C_AcctNm    ,pvStartRow	,pvEndRow   ' 계정코드명	 
		ggoSpread.SSSetProtected	C_ExchRate_2,pvStartRow	,pvEndRow
		ggoSpread.SSSetRequired		C_deptcd_2  ,pvStartRow	,pvEndRow   ' 부서코드		
		ggoSpread.SSSetRequired		C_AcctCd    ,pvStartRow	,pvEndRow	' 계정코드 
		ggoSpread.SSSetRequired		C_DrCrNm_2  ,pvStartRow	,pvEndRow	' 차대구분
		ggoSpread.SSSetRequired		C_DocCur_2  ,pvStartRow ,pvEndRow	' 통화						
		ggoSpread.SSSetRequired		C_ItemAmt_2 ,pvStartRow	,pvEndRow	' 금액	

		.vspdData4.ReDraw = True
    End With
End Sub

'=============================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'============================================================================================================
Sub InitComboBox()
	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("A1013", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)	
	Call SetCombo2(frm1.cboGlType ,lgF0  ,lgF1  ,Chr(11))

	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("A1001", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	Call SetCombo2(frm1.cboGlInputType ,lgF0  ,lgF1  ,Chr(11))
End Sub

'=============================================  2.2.6 InitComboBoxGrid()  =======================================
'	Name : InitComboBoxGrid()
'	Description : Combo Display
'============================================================================================================
Function InitComboBoxGrid(ByVal pvSpdNo)
	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("A1012", "''", "S") & "  order by minor_cd desc ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	
	lgF0 = "" & chr(11) & lgF0
	lgF1 = "" & chr(11) & lgF1
	
	Select Case UCase(Trim(pvSpdNo))
		Case "A"
			ggoSpread.Source = frm1.vspdData
			ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_DR_CR_FG1
			ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_DR_CR_NM1
		Case "B"
			ggoSpread.Source = frm1.vspdData4
			ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_DrCrFg
			ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_DrCrNm_2
	End Select
End Function

'------------------------------------------  회계전표pop-up  ---------------------------------------------
'	Name : openglpopup
'	Description : 회계전표 POP-UP
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopupGL()
	Dim iCalledAspName
	Dim IntRetCD
	Dim arrRet
	Dim arrParam(1)	

	If IsOpenPop = True Then Exit Function

	iCalledAspName = AskPRAspName("a5120ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INForMATION, "a5120ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrParam(0) = Trim(frm1.txtGlNo.value)					'회계전표번호
	arrParam(1) = ""										'Reference번호

	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function

'------------------------------------------  결의전표 pop-up  ---------------------------------------------
'	Name : openTempglpopup
'	Description :결의전표  POP-UP
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopuptempGL()
	Dim iCalledAspName
	Dim IntRetCD
	Dim arrRet
	Dim arrParam(1)	

	If IsOpenPop = True Then Exit Function
	iCalledAspName = AskPRAspName("a5130ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INForMATION, "a5130ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrParam(0) = Trim(frm1.txtTempGlNo.value)				'회계전표번호
	arrParam(1) = ""										'Reference번호

	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function

'------------------------------------------  OpenPopupOA()  ------------------------------------------------
'	Name : OpenPopupOA()
'	Description : 
'---------------------------------------------------------------------------------------------__------------
Function OpenPopupOA()
	Dim arrRet	
	Dim iStrParm
	Dim iStrParm2
	Dim arrParam(8)
	Dim ii
	Dim IntRetCD
	Dim iCalledAspName

	If lgIntFlgMode = Parent.OPMD_UMODE Then Exit Function								'수정추가시 지우기 
	
	iCalledAspName = AskPRAspName("a5407ra2")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INForMATION, "a5407ra2", "X")
		IsOpenPop = False
		Exit Function
	End If

	With frm1		
	
		if .vspdData.MaxRows > 130 then
			
			call DisplayMsgBox("AU1000", "X", "X", "X")
			exit function
		end if
	
		For ii = 1 To .vspdData.MaxRows
			.vspdData.Row = ii
			.vspdData.Col = C_Gl_NO1
			iStrParm = iStrParm & .vspdData.Text	& Parent.gColSep
			.vspdData.Col = C_OpenGlItemSeq
			iStrParm = iStrParm & .vspdData.Text	& Parent.gRowSep
		Next
	End With
	
	iStrParm2 = frm1.txtClsDt.text

	' 권한관리 추가
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,iStrParm,iStrParm2,arrParam), _
		     "dialogWidth=900px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0,0) = ""  Then			
		Exit Function
	Else		
		Call SeTopenPopupOA(arrRet)
	End If
End Function

'------------------------------------------  SeTopenPopupOA()  ------------------------------------------------
'	Name : SeTopenPopupOA()
'	Description : 
'--------------------------------------------------------------------------------------------------------------
Sub SeTopenPopupOA(ByVal arrRet)
	Dim ii
	Dim jj
	Dim tempstr
	
	If lgCurrentTabFg = TAB2 Then
		Call ChangeTabs(TAB1)
		lgCurrentTabFg = TAB1
	End If
		
	With frm1
		ggoSpread.Source	= frm1.vspdData
		.vspdData.ReDraw	= False

		For ii = 0 To UBound(arrRet,1)
			ggoSpread.InsertRow .vspdData.MaxRows
			lgIntMaxItemSeq = lgIntMaxItemSeq + 1			
			.vspdData.Row 	= .vspdData.MaxRows
			For jj = 1 To C_orgchangeid  + 1
				.vspdData.Col = jj
				Select Case  jj
					Case C_ItemSeq1
						.vspdData.value = lgIntMaxItemSeq
					Case C_CardNo1
						.vspdData.value = arrRet(ii,0)
					Case C_UnSettCd1
						.vspdData.value = arrRet(ii,1)
					Case C_User1
						.vspdData.value = arrRet(ii,2)
					Case C_CardCo1
						.vspdData.value = arrRet(ii,3)	
					Case C_GL_NO1
						.vspdData.value = arrRet(ii,4)
					Case C_GL_DT1
						.vspdData.text  = arrRet(ii,5)
					Case C_DOC_CUR1
						.vspdData.value = arrRet(ii,6)
					Case C_XCH_RATE1
						.vspdData.value = UNICDbl(arrRet(ii,7))
					Case C_OPEN_DOC_AMT1
						.vspdData.value = UNICDbl(arrRet(ii,8))
					Case C_OPEN_AMT1
						.vspdData.value = UNICDbl(arrRet(ii,9))
					Case C_DEPT_CD1
						.vspdData.value = arrRet(ii,10)
					Case C_DEPT_NM1
						.vspdData.value = arrRet(ii,11)
					Case C_ACCT_CD1
						.vspdData.value = arrRet(ii,12)
					Case C_ACCT_NM1
						.vspdData.value = arrRet(ii,13)
					Case C_DR_CR_FG1
						tempstr = arrRet(ii,14)
						If tempstr = "DR" Then			'DR
							.vspdData.value = 2
						ElseIf tempstr = "CR" Then		'CR
							.vspdData.value = 1
						End If
					Case C_DR_CR_NM1
						tempstr = arrRet(ii,14)  
						If tempstr = "DR" Then			'DR
							.vspdData.Value = 2
						ElseIf tempstr = "CR" Then		'CR
							.vspdData.Value = 1
						End If
					Case C_OpenGlItemSeq
						.vspdData.value = arrRet(ii,17)   'lgIntMaxItemSeq
				
					Case C_internalcd
						.vspdData.value = arrRet(ii,18)   'lgIntMaxItemSeq

					Case C_costcd
						.vspdData.value = arrRet(ii,19)   'lgIntMaxItemSeq

					Case C_orgchangeid
						.vspdData.value = arrRet(ii,20)   'lgIntMaxItemSeq
				
					Case Else
						.vspdData.value = ""
				End Select			
			Next
			
			Call vspdData_Change(C_OPEN_DOC_AMT1,frm1.vspdData.ActiveRow)
			Call SetSpreadColor(1, "0", frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow)
		Next		
		
		Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,-1,-1,C_DOC_CUR1, C_XCH_RATE1,"D" ,"I","X","X")
		Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,-1,-1,C_DOC_CUR1, C_OPEN_DOC_AMT1,"A" ,"I","X","X")
		
		.vspdData.ReDraw = True
		.vspdData4.ReDraw = False
		If lgCurrentTabFg = TAB1 Then
		    
		    For ii=1 To .vspdData.MaxRows
				Call DbQuery2(ii, C_VSPDDATA1)
			Next
		End If
		.vspdData4.ReDraw = True
	
		.txtTempGlNo.value = ""
		.txtGlNo.value = ""
	End With
	
    Call SetToolBar(C_MENU_CRT_TAB1)
    lgPreToolBarTab1 = C_MENU_CRT_TAB1
End Sub

'======================================================================================================
'   Function Name : OpenClsPopUp()
'   Function Desc : 
'=======================================================================================================
Function  OpenClsPopUp(Byval strCode, Byval iWhere)
	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(8)
	Dim strCd	
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("a5407ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INForMATION, "a5407ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	' 권한관리 추가
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID

	IsOpenPop = True
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam),  "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")	

	IsOpenPop = False
	
	If arrRet(0) = "" Then	    
		Exit Function
	Else
		frm1.txtClsNo.value = arrRet(0)
		frm1.txtClsNo.focus						
	End If
End Function

'======================================================================================================
'   Function Name : OpenPopUp()
'   Function Desc : 
'=======================================================================================================
Function OpenPopUp(Byval pStrCode, Byval pIntPopUp, ByVal pIntVspdData)
	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)
	Dim iArrStrRet				'권한관리 추가   							  
	
	If IsOpenPop = True Then Exit Function

	Select Case   pIntPopUp
		Case   C_AcctPopup      													' Header명(1)			
			iArrParam(0) = "계정코드팝업"								' 팝업 명칭 
			iArrParam(1) = "A_Acct, A_ACCT_GP" 								' TABLE 명칭 
			iArrParam(2) = pStrCode											' Code Condition
			iArrParam(3) = ""												' Name Cindition
			iArrParam(4) = "A_ACCT.GP_CD=A_ACCT_GP.GP_CD AND A_ACCT.DEL_FG <> " & FilterVar("Y", "''", "S") & "  "'미결계정은 반제계정으로 사용하지 못하게 함 AND ISNULL(A_ACCT.MGNT_FG,'') <> 'Y' "												' Where Condition
			iArrParam(5) = "계정코드"									' 조건필드의 라벨 명칭 

			iArrField(0) = "A_ACCT.Acct_CD"									' Field명(0)
			iArrField(1) = "A_ACCT.Acct_NM"									' Field명(1)
    		iArrField(2) = "A_ACCT_GP.GP_CD"								' Field명(2)
			iArrField(3) = "A_ACCT_GP.GP_NM"								' Field명(3)
			
			iArrHeader(0) = "계정코드"									' Header명(0)
			iArrHeader(1) = "계정코드명"									' Header명(1)
			iArrHeader(2) = "그룹코드"									' Header명(2)
			iArrHeader(3) = "그룹명"
		Case   C_DocCurPopup_2
			iArrParam(0) = "통화코드 팝업"								' 팝업 명칭			
			iArrParam(1) = "B_Currency"	    								' TABLE 명칭 
			iArrParam(2) = pStrCode											' Code Condition
			iArrParam(3) = ""												' Name Cindition
			iArrParam(4) = ""												' Where Condition
			iArrParam(5) = "통화코드"									' 조건필드의 라벨 명칭 

			iArrField(0) = "Currency"	    								' Field명(0)
			iArrField(1) = "Currency_desc"	    							' Field명(1)
    
			iArrHeader(0) = "통화코드"									' Header명(0)
			iArrHeader(1) = "통화코드명"									' Header명(3)
	End Select
    
	IsOpenPop = True
	iArrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
									 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False
	
	If iArrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopUp(iArrRet, pIntPopUp, pIntVspdData)
	End If	
End Function

'-----------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPopUp(Byval arrRet, Byval pIntPopUp, ByVal iVspdData)
	Dim iObjSpread
	
	If iVspdData = C_VSPDDATA1 Then
		Set iObjSpread = frm1.vspdData
	ElseIf iVspdData = C_vspdData4 Then
		Set iObjSpread = frm1.vspdData4
	End If

	With frm1	
		Select Case   pIntPopUp				
			Case   C_AcctPopup      
				iObjSpread.Row = iObjSpread.ActiveRow 
				iObjSpread.Col  = C_AcctCd
				iObjSpread.Text = arrRet(0)
				iObjSpread.Col  = C_AcctNm            
				iObjSpread.Text = arrRet(1)
				
				If iVspdData = C_VSPDDATA1 Then

				ElseIf iVspdData = C_vspdData4 Then					
					Call vspdData4_Change(C_AcctCd, iObjSpread.Activerow)
				End If
			Case	C_DocCurPopup_2 
				iObjSpread.Row = iObjSpread.ActiveRow 
				iObjSpread.Col  = C_DocCur_2 
				iObjSpread.Text = arrRet(0)
				Call vspdData4_Change(C_DocCur_2, iObjSpread.Activerow)
		End Select
	End With
End Function

'-----------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpEndept(Byval pStrCode)
	Dim iCalledAspName
	Dim iArrRet
	Dim iArrParam(8)

	If IsOpenPop = True Then Exit Function
	If frm1.txtDeptCd.className = parent.UCN_PROTECTED Then Exit Function

	iCalledAspName = AskPRAspName("DeptPopupDtA2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INForMATION, "DeptPopupDtA2", "X")
		IsOpenPop = False
		Exit Function
	End If

	iArrParam(0) = pStrCode									'  Code Condition
   	iArrParam(1) = frm1.txtClsDt.Text
	iArrParam(2) = lgUsrIntCd								' 자료권한 Condition  

	If lgIntFlgMode = Parent.OPMD_UMODE Then
		iArrParam(3) = "T"									' 결의일자 상태 Condition  
	Else
		iArrParam(3) = "F"									' 결의일자 상태 Condition  
	End If

	' 권한관리 추가
	iArrParam(5) = lgAuthBizAreaCd
	iArrParam(6) = lgInternalCd
	iArrParam(7) = lgSubInternalCd
	iArrParam(8) = lgAuthUsrID
	
	IsOpenPop = True
	
	iArrRet = window.showModalDialog(iCalledAspName, Array(window.parent, iArrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If iArrRet(0) = "" Then
		Exit Function
	Else
		Call SetDept(iArrRet, C_CONDFIELD)
	End If	
End Function

'-----------------------------------------  OpenUnderDept()  --------------------------------------------------
'	Name : OpenUnderDept()
'	Description : 
'---------------------------------------------------------------------------------------------------------
Function OpenUnderDept(Byval pStrCode, ByVal pIntVspdData)
	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)
    Dim field_fg   	

	IsOpenPop = True
	If RTrim(LTrim(frm1.txtDeptCd.value)) <> "" 	Then
		iArrParam(0) = "부서 팝업"	
		iArrParam(1) = "B_ACCT_DEPT A, B_COST_CENTER B"				
		iArrParam(2) = Trim(pStrCode)
		iArrParam(3) = "" 
		iArrParam(4) = "A.ORG_CHANGE_ID =  " & FilterVar(frm1.hOrgChangeId.value , "''", "S") & ""
		iArrParam(4) = iArrParam(4) & " AND A.COST_CD = B.COST_CD AND B.BIZ_AREA_CD = ( SELECT B.BIZ_AREA_CD"
		iArrParam(4) = iArrParam(4) & " FROM B_ACCT_DEPT A, B_COST_CENTER B WHERE A.DEPT_CD =  " & FilterVar(frm1.txtDeptCd.value , "''", "S") & ""
		iArrParam(4) = iArrParam(4) & " AND A.COST_CD = B.COST_CD AND A.ORG_CHANGE_ID =  " & FilterVar(frm1.hOrgChangeId.value , "''", "S") & ")"
		iArrParam(5) = "부서코드"			
		iArrField(0) = "A.DEPT_CD"	
		iArrField(1) = "A.DEPT_Nm"
		iArrField(2) = "B.BIZ_AREA_CD"
		iArrHeader(0) = "부서코드"		
		iArrHeader(1) = "부서코드명"
		iArrHeader(2) = "사업장코드"	
	Else
		iArrParam(0) = "부서 팝업"	
		iArrParam(1) = "B_ACCT_DEPT A"				
		iArrParam(2) = Trim(pStrCode)
		iArrParam(3) = "" 
		iArrParam(4) = "A.ORG_CHANGE_ID = (select distinct org_change_id"
		iArrParam(4) = iArrParam(4) & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		iArrParam(4) = iArrParam(4) & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtClsDt.Text, Parent.gDateFormat,""), "''", "S") & "))"			
		iArrParam(5) = "부서코드"			
		iArrField(0) = "A.DEPT_CD"	
		iArrField(1) = "A.DEPT_Nm"
		iArrHeader(0) = "부서코드"		
		iArrHeader(1) = "부서코드명"
	End If

	iArrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If iArrRet(0) = "" Then
		Exit Function
	Else
		Call SetDept(iArrRet, pIntVspdData)
	End If	
End Function

'------------------------------------------  SetDept()  ------------------------------------------------
'	Name : SetDept()
'	Description : Dept Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetDept(Byval pArrRet, ByVal pIntVspdData)
	Dim iObjSpread
	
	If pIntVspdData = C_VSPDDATA1 Then
		Set iObjSpread = frm1.vspdData
	ElseIf pIntVspdData = C_vspdData4 Then
		Set iObjSpread = frm1.vspdData4
	End If
		
	With frm1
		If  pIntVspdData = C_CONDFIELD Then
			.txtDeptCd.value = pArrRet(0)
			.txtDeptNm.value = pArrRet(1)
			.txtInternalCd.value = pArrRet(2)
  			 If lgQueryOk <> True Then
			 	.txtClsDt.text = pArrRet(3)
			 End If           
			 Call txtDeptCd_OnChange() 
		Else
			iObjSpread.Row = iObjSpread.ActiveRow 
				
			ggoSpread.Source = iObjSpread
			ggoSpread.UpdateRow iObjSpread.ActiveRow 
				
			iObjSpread.Col  = C_deptcd_2
			iObjSpread.Text = pArrRet(0)
			iObjSpread.Col  = C_deptnm_2
			iObjSpread.Text = pArrRet(1)
 		
			Call deptCd_underChange(pArrRet(0))
		End If				
	End With
End Function     

'======================================================================================================
'	기능: Tab Click
'	설명: Tab Click시 필요한 기능을 수행한다.
'=======================================================================================================
Function ClickTab1()
	If lgCurrentTabFg = TAB1 Then Exit Function

	Call ChangeTabs(TAB1)	 '

	lgCurrentTabFg = TAB1

	If lgPreToolBarTab1 <> "" Then
		Call SetToolBar(lgPreToolBarTab1)
		Exit Function
	End If

	If lgIntFlgMode <> Parent.OPMD_UMODE Then
	   Call SetToolBar(C_MENU_NEW_TAB1)
	End If
End Function

Function ClickTab2()
	If lgCurrentTabFg = TAB2 Then Exit Function
	
	Call ChangeTabs(TAB2)	 '~~~ 첫번째 Tab 
	lgCurrentTabFg = TAB2

	If lgPreToolBarTab2 <> "" Then
		Call SetToolBar(lgPreToolBarTab2)
		Exit Function
	End If
	
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		If frm1.vspdData4.MaxRows = 0 Then
			Call SetToolBar(C_MENU_NEW_TAB2)
		Else
			Call SetToolBar(C_MENU_CRT_TAB2)
		End If		
	End If	
End Function

'========================================================================================================
'	Desc : 입출금 화면에 따른 Grid의 Protect변환 
'========================================================================================================
Sub CboGLType_ProtectGrid(Byval GlType)
	ggoSpread.Source = frm1.vspdData
	Select Case  GlType		
		Case "01"			
			ggoSpread.SSSetProtected C_DR_CR_FG1, 1, frm1.vspddata.maxrows	' 차대구분 
			ggoSpread.SSSetProtected C_DR_CR_NM1, 1, frm1.vspddata.maxrows	' 차대구분 
		Case "02"			
			ggoSpread.SSSetProtected C_DR_CR_FG1, 1, frm1.vspddata.maxrows	' 차대구분 
			ggoSpread.SSSetProtected C_DR_CR_NM1, 1, frm1.vspddata.maxrows	' 차대구분 
		Case "03"			
			ggoSpread.SpreadUnLock   C_DR_CR_FG1, 1, C_DrCrNm, frm1.vspddata.maxrows
			ggoSpread.SSSetRequired  C_DR_CR_FG1, 1, frm1.vspddata.maxrows	' 차대구분 
			ggoSpread.SSSetRequired  C_DR_CR_NM1, 1, frm1.vspddata.maxrows	' 차대구분 
	End Select 				
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

			C_ItemSeq1      = iCurColumnPos(1)
            C_CardNo1	    = iCurColumnPos(2) 
            C_UnSettCd1	    = iCurColumnPos(3) 
            C_User1		    = iCurColumnPos(4)
            C_CardCo1	    = iCurColumnPos(5)		
            C_GL_NO1	    = iCurColumnPos(6) 
            C_GL_DT1	    = iCurColumnPos(7) 
            C_DOC_CUR1	    = iCurColumnPos(8) 
            C_XCH_RATE1	    = iCurColumnPos(9) 
            C_OPEN_DOC_AMT1	= iCurColumnPos(10) 
            C_OPEN_AMT1	    = iCurColumnPos(11) 
            C_DEPT_CD1	    = iCurColumnPos(12) 
            C_DEPT_NM1	    = iCurColumnPos(13) 
            C_ACCT_CD1	    = iCurColumnPos(14) 
            C_ACCT_NM1	    = iCurColumnPos(15) 
            C_DR_CR_FG1	    = iCurColumnPos(16) 
            C_DR_CR_NM1	    = iCurColumnPos(17) 
            C_OpenGlItemSeq = iCurColumnPos(18)
       Case "B"
            ggoSpread.Source = frm1.vspdData4
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
	
            C_ItemSeq		  = iCurColumnPos(1)
            C_deptcd_2		  = iCurColumnPos(2)
            C_deptPopup_2	  = iCurColumnPos(3)
            C_deptnm_2		  = iCurColumnPos(4)
            C_AcctCd		  = iCurColumnPos(5)
            C_AcctPopup   	  = iCurColumnPos(6)
            C_AcctNm      	  = iCurColumnPos(7)
            C_DrCrFg		  = iCurColumnPos(8)
            C_DrCrNm_2	 	  = iCurColumnPos(9)
            C_DocCur_2		  = iCurColumnPos(10)
            C_DocCurPopup_2   = iCurColumnPos(11)
            C_ExchRate_2	  = iCurColumnPos(12)
            C_ItemAmt_2		  = iCurColumnPos(13)
            C_ItemLocAmt_2	  = iCurColumnPos(14)
            C_ItemDesc_2	  = iCurColumnPos(15)
            C_IsLAmtChange_2  = iCurColumnPos(16)
            C_OpenGlNo_2	  = iCurColumnPos(17)
            C_OpenGlItemSeq_2 = iCurColumnPos(18)
            C_MgntFg_2		  = iCurColumnPos(19)
            C_AcctCd2_2		  = iCurColumnPos(20)
    End Select    
End Sub

'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : Form_Load
' Desc : This sub is called from window_OnLoad in Common.vbs auTomatically
'========================================================================================================
Sub Form_Load()
   	Call LoadInfTB19029															'⊙: Load table , B_numeric_Format
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)    
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)    
    Call ggoOper.LockField(Document, "N")										'⊙: Lock  Suitable  Field

	Call InitVariables															'⊙: Initializes local global variables
    Call InitSpreadSheet("A")
    Call InitSpreadSheet("B")

	Call InitCtrlSpread()
	Call InitCtrlHSpread()

	Call InitCtrlSpread2()
	Call InitCtrlHSpread2()

	Call InitComboBox()
	Call InitComboBoxGrid("A")     
	Call InitComboBoxGrid("B")     
	Call SetDefaultVal
    Call SetToolbar("1100000000001111")											'⊙: 버튼 툴바 제어

	gIsTab			 = "Y" 
	gTabMaxCnt       = 2         
	lgBlnFlgChgValue = False			
    frm1.txtClsNo.focus 
    
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
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 

    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = Frm1.vspdData
	ggoSpread.ClearSpreadData
    ggoSpread.Source = Frm1.vspdData2
	ggoSpread.ClearSpreadData
    ggoSpread.Source = Frm1.vspdData3
	ggoSpread.ClearSpreadData		
    ggoSpread.Source = Frm1.vspdData4
	ggoSpread.ClearSpreadData
    ggoSpread.Source = Frm1.vspdData5
	ggoSpread.ClearSpreadData
    ggoSpread.Source = Frm1.vspdData6
	ggoSpread.ClearSpreadData	

    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want To display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

	Call ggoOper.LockField(Document, "Q")	

    If gSelframeFlg = TAB1 Then
		ggoSpread.Source = Frm1.vspdData
		If Not chkField(Document, "1") Then									      '⊙: This function check indispensable field
			Exit Function
		End If		
		If DbQuery() = False Then                                                      '☜: Query db data
			Exit Function
		End If
	Else
		ggoSpread.Source = Frm1.vspdData4
		If Not chkField(Document, "1") Then									      '⊙: This function check indispensable field
			Exit Function
		End If		
		If DbQuery() = False Then                                                      '☜: Query db data
		   Exit Function
		End If
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
	Dim var1, var2
	    
    On Error Resume Next														'☜: Protect system from crashing
    Err.Clear																	'☜: Protect system from crashing

    FncNew = False                                                          
   
    ggoSpread.Source = frm1.vspdData
    var1 = ggoSpread.SSCheckChange
    ggoSpread.Source = frm1.vspdData4
    var2 = ggoSpread.SSCheckChange

	'-----------------------
    'Check previous data area
    '----------------------- 
    If (lgBlnFlgChgValue = True Or var1 = True Or var2 = True) And lgBlnExecDelete <> True Then
        IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

	lgBlnExecDelete = False
	
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")										'⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    Call ggoOper.LockField(Document, "N")										'⊙: Lock  Suitable  Field
	Call InitComboBoxGrid("A")     
	Call InitComboBoxGrid("B")     

    frm1.txtClsNo.focus        
    
    Call ClickTab1()
    Call SetToolbar(C_MENU_NEW_TAB1)
    lgPreToolBarTab1 = 	C_MENU_NEW_TAB1		 		    
    lgPreToolBarTab2 = 	C_MENU_NEW_TAB2
    
	Call ggoOper.SetReqAttr(frm1.txtDeptCd,"N")
	Call ggoOper.SetReqAttr(frm1.txtClsDt, "N")
	Call ggoOper.SetReqAttr(frm1.txtdesc,  "D")	
    
    ggoSpread.Source = Frm1.vspdData
	ggoSpread.ClearSpreadData
    ggoSpread.Source = Frm1.vspdData2
	ggoSpread.ClearSpreadData
    ggoSpread.Source = Frm1.vspdData3
	ggoSpread.ClearSpreadData		
    ggoSpread.Source = Frm1.vspdData4
	ggoSpread.ClearSpreadData
    ggoSpread.Source = Frm1.vspdData5
	ggoSpread.ClearSpreadData
    ggoSpread.Source = Frm1.vspdData6
	ggoSpread.ClearSpreadData	
        
	SetGridFocus()
    
	Call SetDefaultVal    
	Call InitVariables															'⊙: Initializes local global variables	
	Call txtDocCur_OnChange()
		
    lgBlnFlgChgValue = False

    FncNew = True																'⊙: Processing is OK
    lgQueryOk = False
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
	Dim IntRetCD                                                    '☜: Processing is OK
    FncDelete = False                                                      
   
    lgBlnExecDelete = True

    '-----------------------
    '	Precheck area
    '-----------------------
    ' Update 상태인지를 확인한다.
    ggoSpread.Source = frm1.vspdData4    

    If ggoSpread.SSCheckChange = False Then									'변경된 부분이 없을경우 
		intRetCd = DisplayMsgBox("900003", Parent.VB_YES_NO, "X", "X")				'삭제하시겠습니까?
		If intRetCd = VBNO Then
			Exit Function
		End If
	Else
		IntRetCD = DisplayMsgBox("900038", Parent.VB_YES_NO, "X", "X")			    '데이타가 변경되었습니다. 조회하시겠습니까?
    	If IntRetCD = vbNo Then    		
      		Exit Function
    	End If
    End If

    '-----------------------
    'Delete function call area
    '-----------------------
    If  DbDelete = False Then														'☜: Delete db data
    	Exit Function
    End If

    FncDelete = True 
End Function

'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
	Dim IntRetCD 
	
	On Error Resume Next
	Err.Clear                                                                    '☜: Clear err status
	    
	FncSave = False                                                              '☜: Processing is NG
    
	ggoSpread.Source = frm1.vspdData
	If Not ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
		Exit Function
	End If

	ggoSpread.Source = frm1.vspdData4
	If Not ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
		Exit Function
	End If

	If CheckSpread6 = False Then
		IntRetCD = DisplayMsgBox("110420", "X", "X", "X")							'필수입력 check!!
        Exit Function
    End If
    
	If DbSave = False Then                                                       '☜: Query db data
		Exit Function
	End If

	Set gActiveElement = document.ActiveElement   
	FncSave = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure To clear primary key area
'========================================================================================================
Function FncCopy() 
	Dim  IntRetCD
	 
	frm1.vspdData4.ReDraw = False	
	If frm1.vspdData4.MaxRows < 1 Then Exit Function

    ggoSpread.Source = frm1.vspdData4	
    ggoSpread.CopyRow
    
    frm1.vspdData4.Col   = C_ItemSeq
	frm1.vspdData4.Row   = frm1.vspdData4.ActiveRow
	lgIntMaxItemSeq      = lgIntMaxItemSeq + 1
	frm1.vspdData4.value = lgIntMaxItemSeq
	
    Call SetSpread2Color("I",0, frm1.vspdData4.ActiveRow, frm1.vspdData4.ActiveRow)
    Call SetSumItem()
    
	frm1.vspdData4.ReDraw = True
	Call vspdData4_Change(C_AcctCd, frm1.vspdData4.activerow)    
End Function

'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related To Cancel ButTon of Main ToolBar
'========================================================================================================
Function FncCancel() 
	Dim iItemSeq
	
	If lgCurrentTabFg = TAB1 Then
		If frm1.vspdData.MaxRows < 1 Then Exit Function	

		With frm1.vspdData
		    .Row = .ActiveRow
		    .Col = 0		    
		    If .Text = ggoSpread.InsertFlag Then
				.Col = C_ACCT_CD1
				If len(Trim(.text)) > 0 Then 
					.Col = C_ItemSeq1
					Call DeleteHSheet(.Text)
				End If	
		    End If       

		    ggoSpread.Source = frm1.vspdData	
		    ggoSpread.EditUndo

			Call SetSumItem()

			If .MaxRows = 0 Then
				Call SetToolbar(C_MENU_NEW_TAB1)
				lgPreToolBarTab1 = C_MENU_NEW_TAB1
				Exit Function
			End If

		    .Row = .ActiveRow
		    .Col = 0
		    
			If .Row = 0 Then 			
				Exit Function
			End If
		End With
	ElseIf lgCurrentTabFg = TAB2 Then
		If frm1.vspdData4.MaxRows < 1 Then Exit Function	

		With frm1.vspdData4
		    .Row = .ActiveRow
		    .Col = 0		    
		    If .Text = ggoSpread.InsertFlag Then
				.Col = C_AcctCd
				If len(Trim(.text)) > 0 Then 
					.Col = C_ItemSeq
					Call DeleteHSheet2(.Text)
				End If	
		    End If       
			
		    ggoSpread.Source = frm1.vspdData4	
		    ggoSpread.EditUndo

			Call SetSumItem()
			If .MaxRows = 0 Then
				Call SetToolbar(C_MENU_NEW_TAB2)
				lgPreToolBarTab1 = C_MENU_NEW_TAB2
				Exit Function
			End If

		    .Row = .ActiveRow
		    .Col = 0
		    
			If .Row = 0 Then 			
				Exit Function
			End If
			
		    If .Text = ggoSpread.InsertFlag Then            
			    .Col = C_AcctCd
		        If Len(.Text) > 0 Then
					.Col = C_ItemSeq
					frm1.hItemSeq.value = .Text
		            frm1.vspdData5.MaxRows = 0
			        Call DbQuery4(.ActiveRow)
		        End If
		    Else
				.Col = C_ItemSeq
		        frm1.hItemSeq.value = .Text
		        frm1.vspdData5.MaxRows = 0
			    Call DbQuery2(.ActiveRow, C_vspdData4)
		    End If		    
		End With		
    End If
End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
	Dim iRow
    Dim imRow
    Dim imRow2

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
   '-----------------------
    'Check content area
    '----------------------- 
    If Not chkField(Document, "2") Then 
		Call ClickTab1()
        Exit Function
    End If

    If IsNumeric(Trim(pvRowCnt)) Then
        imRow = CInt(pvRowCnt)
    Else
        imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then
            Exit Function
        End If
    End If
    
	With frm1
		.vspdData4.focus
		ggoSpread.Source = .vspdData4	
		.vspdData4.ReDraw = False
		ggoSpread.InsertRow ,imRow
		
		Call SetSpreadColor2("I",0, frm1.vspdData4.ActiveRow, frm1.vspdData4.ActiveRow + imRow - 1)
		
		For iRow = .vspdData4.ActiveRow To .vspdData4.ActiveRow + imRow - 1
			'귀속부서를 default로 뿌려준다.
    	    .vspdData4.Row = iRow
			.vspdData4.col		= C_deptcd_2
			.vspdData4.value	= UCase(.txtDeptCd.value)
			
			.vspdData4.col		= C_deptnm_2
			.vspdData4.value	= .txtDeptNm.value		
			
			.vspdData4.col		= C_ItemDesc_2     
			.vspdData4.value	= .txtDesc.value
			
			.vspdData4.col		= C_DocCur_2 
			.vspdData4.value	= parent.gcurrency
			
			.vspdData4.col		= C_ExchRate_2    
			.vspdData4.value	= "1"		
			
			'입금전표이면 (01) '	'cr'을 넣어준다.
			If  frm1.cboGlType.value = "01" Then
				.vspdData4.col = C_DrCrNm_2
				.vspdData4.value	= 1					
				.vspdData4.col = C_DrCrFg
				.vspdData4.value	= 1					
			ElseIf frm1.cboGlType.value = "02" Then		
				.vspdData4.col = C_DrCrNm_2
				.vspdData4.value	= 2				
				.vspdData4.col = C_DrCrFg
				.vspdData4.value	= 2			
			End If
			.vspdData4.Col = C_ItemSeq
			lgIntMaxItemSeq = lgIntMaxItemSeq + 1
			.vspdData4.value = lgIntMaxItemSeq
			.vspdData4.ReDraw = True
        Next

		Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData4,-1,-1,C_DocCur_2, C_ExchRate_2,"D" ,"I","X","X")
		Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData4,-1,-1,C_DocCur_2, C_ItemAmt_2,"A" ,"I","X","X")

		.vspdData5.MaxRows =  0		        
        Call SetToolBar(C_MENU_CRT_TAB2)
        lgPreToolBarTab2 = C_MENU_CRT_TAB2
    End With

    Set gActiveElement = document.ActiveElement 
End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow() 
	Dim lDelRows
	Dim iDelRowCnt, i
    Dim DelItemSeq

	If lgCurrentTabFg = TAB1 Then
		With frm1.vspdData 
			ggoSpread.Source = frm1.vspdData 
			.Row = .ActiveRow
			.Col = 0 		
			If frm1.vspdData.MaxRows < 1 Or .Text = ggoSpread.InsertFlag Then Exit Function
			
			.Col = 1 
			DelItemSeq = .Text    	
			lDelRows = ggoSpread.DeleteRow    
		End With
		
		Call DeleteHsheet(DelItemSeq)
    ElseIf lgCurrentTabFg = TAB2 Then
		With frm1.vspdData4 
			ggoSpread.Source = frm1.vspdData4 
			.Row = .ActiveRow
			.Col = 0 		
			If frm1.vspdData4.MaxRows < 1 Or .Text = ggoSpread.InsertFlag Then Exit Function
			
			.Col = 1 
			DelItemSeq = .Text    	
			lDelRows = ggoSpread.DeleteRow    
		End With
		
		Call DeleteHsheet2(DelItemSeq)
	End If
    
    Call SetSumItem()
End Function

'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
'========================================================================================================
Function FncPrint()
    FncPrint = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncPrint()                                                       '☜: Protect system from crashing
    FncPrint = True                                                              '☜: Processing is OK
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

'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    FncExit = True                                                            '☜: Processing is OK
End Function

'======================================================================================================
' Function Name : SetSumItem
' Function Desc :
'=======================================================================================================
Function SetSumItem()
    Dim DblTotDrAmt 
    DIm DblTotLocDrAmt
    Dim DblTotCrAmt 
    DIm DblTotLocCrAmt        
    Dim lngRows 
    
    DblTotDrAmt		= 0
	DblTotLocDrAmt	= 0
    DblTotCrAmt		= 0 
    DblTotLocCrAmt	= 0

	ggoSpread.Source = frm1.vspdData
	
    With frm1.vspdData 
		If .MaxRows > 0 Then    
	        For lngRows = 1 To .MaxRows
	            .Row = lngRows
                .Col = 0
                If .text <> ggoSpread.DeleteFlag Then
		            .col = C_DR_CR_FG1
		            If .text = "DR" Then		
			            .Col = C_OPEN_DOC_AMT1	'6
			            If .Text = "" Then
			                DblTotDrAmt = UNICDbl(DblTotDrAmt) + 0
			            Else
			                DblTotDrAmt = UNICDbl(DblTotDrAmt) + UNICDbl(.Text)
			            End If
			            			            
			            .Col = C_OPEN_AMT1	'7
			            If .Text = "" Then
			                DblTotLocDrAmt = UNICDbl(DblTotLocDrAmt) + 0
			            Else
			                DblTotLocDrAmt = UNICDbl(DblTotLocDrAmt) + UNICDbl(.Text)
			            End If
		            ElseIf .text = "CR" Then
			            .Col = C_OPEN_DOC_AMT1 	'6
			            If .Text = "" Then
			                DblTotCrAmt = UNICDbl(DblTotCrAmt) + 0
			            Else
			                DblTotCrAmt = UNICDbl(DblTotCrAmt) + UNICDbl(.Text)
			            End If
			            
			            .Col = C_OPEN_AMT1	'7
			            If .Text = "" Then
			                DblTotLocCrAmt = UNICDbl(DblTotLocCrAmt) + 0
			            Else
			                DblTotLocCrAmt = UNICDbl(DblTotLocCrAmt) + UNICDbl(.Text)
			            End If
					End If	
				End If	            
	        Next 
		End If                
	End With
	
	With frm1.vspdData4 
		If .MaxRows > 0 Then    
	        For lngRows = 1 To .MaxRows
	            .Row = lngRows
                .Col = 0
                If .text <> ggoSpread.DeleteFlag Then
		            .col = C_DrCrFg      
		            If .text = "DR" Then		
			            .Col = C_ItemAmt_2	'6
			            If .Text = "" Then
			                DblTotDrAmt = UNICDbl(DblTotDrAmt) + 0
			            Else
			                DblTotDrAmt = UNICDbl(DblTotDrAmt) + UNICDbl(.Text)
			            End If
			            			            
			            .Col = C_ItemLocAmt_2	'7
			            If .Text = "" Then
			                DblTotLocDrAmt = UNICDbl(DblTotLocDrAmt) + 0
			            Else
			                DblTotLocDrAmt = UNICDbl(DblTotLocDrAmt) + UNICDbl(.Text)
			            End If
		            ElseIf .text = "CR" Then
			            .Col = C_ItemAmt_2	'6
			            If .Text = "" Then
			                DblTotCrAmt = UNICDbl(DblTotCrAmt) + 0
			            Else
			                DblTotCrAmt = UNICDbl(DblTotCrAmt) + UNICDbl(.Text)
			            End If
			            
			            .Col = C_ItemLocAmt_2	'7
			            If .Text = "" Then
			                DblTotLocCrAmt = UNICDbl(DblTotLocCrAmt) + 0
			            Else
			                DblTotLocCrAmt = UNICDbl(DblTotLocCrAmt) + UNICDbl(.Text)
			            End If
					End If	
				End If	            
	        Next 
		End If                
	End With

    frm1.txtDrLocAmt.text = UNIConvNumPCToCompanyByCurrency(DblTotLocDrAmt,Parent.gCurrency,Parent.ggAmTofMoneyNo, Parent.gLocRndPolicyNo, "X")
    frm1.txtCrLocAmt.text = UNIConvNumPCToCompanyByCurrency(DblTotLocCrAmt,Parent.gCurrency,Parent.ggAmTofMoneyNo, Parent.gLocRndPolicyNo, "X")
    frm1.txtDrLocAmt2.text = UNIConvNumPCToCompanyByCurrency(DblTotLocDrAmt,Parent.gCurrency,Parent.ggAmTofMoneyNo, Parent.gLocRndPolicyNo, "X")
    frm1.txtCrLocAmt2.text = UNIConvNumPCToCompanyByCurrency(DblTotLocCrAmt,Parent.gCurrency,Parent.ggAmTofMoneyNo, Parent.gLocRndPolicyNo, "X")
      
	If frm1.cboGlType.value = "01" Then
		frm1.txtDrLocAmt.text	= frm1.txtCrLocAmt.text
		frm1.txtDrLocAmt2.text	= frm1.txtCrLocAmt2.text
	ElseIf frm1.cboGlType.value = "02" Then
		frm1.txtCrLocAmt.text	= frm1.txtDrLocAmt.text
		frm1.txtCrLocAmt2.text	= frm1.txtDrLocAmt2.text
	End If
End Function

'========================================================================================
' Function Name : FncBtnCalc
' Function Desc : This function calculate local amt from amt of multi
'========================================================================================
Function FncBtnCalc() 
	Dim ii
	Dim tempAmt, tempLocAmt, tempExch, TempSep, tempDoc
	Dim strSelect, strFrom, strWhere
	Dim arrTemp
	Dim strDate
	Dim strExchFg
	Dim IntRetCD
	Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6		

	With frm1
		strSelect	= "b.minor_cd"
		strFrom		= "b_company a, b_minor b"
		strWhere	= "b.major_cd = " & FilterVar("a1004", "''", "S") & "  and	a.xch_rate_fg = b.minor_cd"
		If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then				
			arrTemp = Split(lgF0, chr(11))
			strExchFg =  arrTemp(0)
		End If	

		strDate = UniConvDateToYYYYMMDD(frm1.txtClsDt.text,parent.gDateFormat,"")

		If .vspdData.MaxRows <> 0 Then
			For ii = 1 To .vspdData.MaxRows
				.vspdData.Row	=	ii
				.vspdData.Col	=	C_DOC_CUR1			
				tempDoc			=	UCase(Trim(.vspdData.text))
				.vspdData.Col	=	C_OPEN_DOC_AMT1
				tempAmt			=	UNICDbl(.vspdData.text)
				.vspdData.Col	=	C_XCH_RATE1
				tempExch		=	UNICDbl(.vspdData.text)

				If tempDoc	<> "" and tempDoc <> parent.gCurrency Then
					If UCase(strExchFg) <> "D" Then 	' Fixed Exchange Rate
						strDate = Mid(strDate, 1, 6)
						strSelect	= "multi_divide"
						strFrom		= "b_monthly_exchange_rate (noLock) "
						strWhere	= "from_currency =  " & FilterVar(tempDoc , "''", "S") & ""
						strWhere	= strWhere & " And To_currency   =  " & FilterVar(parent.gCurrency , "''", "S") & ""
						strWhere	= strWhere & " And apprl_yrmnth  =  " & FilterVar(strDate , "''", "S") & ""

						If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then				
							arrTemp = Split(lgF0, chr(11))
							TempSep =  arrTemp(0)
						Else
							IntRetCD = DisplayMsgBox("121600", "X", "X", "X")
						End If
					Else					' Floating Exchange Rate
						strSelect	= "Top 1 multi_divide"
						strFrom		= "b_daily_exchange_rate (noLock) "
						strWhere	= "from_currency =  " & FilterVar(tempDoc , "''", "S") & ""
						strWhere	= strWhere & " And To_currency   =  " & FilterVar(parent.gCurrency , "''", "S") & ""
						strWhere	= strWhere & " And apprl_dt  <= convert(char(21), " & FilterVar(strDate, "''", "S") & ", 20) order by apprl_dt"

						If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then				
							arrTemp = Split(lgF0, chr(11))
							TempSep =  arrTemp(0)
						Else
							IntRetCD = DisplayMsgBox("121500", "X", "X", "X")
						End If
					End If
					If RTrim(LTrim(TempSep)) <> "/" Then
						tempLocAmt		=	tempAmt * TempExch
					Else
						tempLocAmt		=	tempAmt / TempExch
					End If
					.vspdData.Col	=	C_OPEN_AMT1
					.vspdData.text	=	tempLocAmt	'UNIConvNumPCToCompanyByCurrency(tempLocAmt,parent.gCurrency,parent.ggAmTofMoneyNo, parent.gLocRndPolicyNo, "X")

				ElseIf tempDoc = parent.gCurrency Then
					.vspdData.Col	=	C_OPEN_AMT1
					.vspdData.text	=	tempAmt		'UNIConvNumPCToCompanyByCurrency(tempAmt,parent.gCurrency,parent.ggAmTofMoneyNo, parent.gLocRndPolicyNo, "X")
				End If
			Next		
		End If

		If .vspdData4.MaxRows <> 0 Then
			For ii = 1 To .vspdData4.MaxRows
				.vspdData4.Row	=	ii
				.vspdData4.Col	=	C_DocCur_2			
				tempDoc			=	UCase(Trim(.vspdData4.text))
				.vspdData4.Col	=	C_ItemAmt_2
				tempAmt			=	UNICDbl(.vspdData4.text)
				.vspdData4.Col	=	C_ExchRate_2
				tempExch		=	UNICDbl(.vspdData4.text)

				If tempDoc	<> "" and tempDoc <> parent.gCurrency Then
					If UCase(strExchFg) <> "D" Then 	' Fixed Exchange Rate
						strDate = Mid(strDate, 1, 6)
						strSelect	= "multi_divide"
						strFrom		= "b_monthly_exchange_rate (noLock) "
						strWhere	= "from_currency =  " & FilterVar(tempDoc , "''", "S") & ""
						strWhere	= strWhere & " And To_currency   =  " & FilterVar(parent.gCurrency , "''", "S") & ""
						strWhere	= strWhere & " And apprl_yrmnth  =  " & FilterVar(strDate , "''", "S") & ""

						If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then				
							arrTemp = Split(lgF0, chr(11))
							TempSep =  arrTemp(0)
						Else
							IntRetCD = DisplayMsgBox("121600", "X", "X", "X")
						End If
					Else					' Floating Exchange Rate
						strSelect	= "Top 1 multi_divide"
						strFrom		= "b_daily_exchange_rate (noLock) "
						strWhere	= "from_currency =  " & FilterVar(tempDoc , "''", "S") & ""
						strWhere	= strWhere & " And To_currency   =  " & FilterVar(parent.gCurrency , "''", "S") & ""
						strWhere	= strWhere & " And apprl_dt  <= convert(char(21), " & FilterVar(strDate, "''", "S") & ", 20) order by apprl_dt"

						If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then				
							arrTemp = Split(lgF0, chr(11))
							TempSep =  arrTemp(0)
						Else
							IntRetCD = DisplayMsgBox("121500", "X", "X", "X")
						End If
					End If
					If RTrim(LTrim(TempSep)) <> "/" Then
						tempLocAmt		=	tempAmt * TempExch
					Else
						tempLocAmt		=	tempAmt / TempExch
					End If
					.vspdData4.Col	=	C_ItemLocAmt_2
					.vspdData4.text	=	tempLocAmt'UNIConvNumPCToCompanyByCurrency(tempLocAmt,parent.gCurrency,parent.ggAmTofMoneyNo, parent.gLocRndPolicyNo, "X")

				ElseIf tempDoc = parent.gCurrency Then
					.vspdData4.Col	=	C_ItemLocAmt_2
					.vspdData4.text	=	tempAmt'UNIConvNumPCToCompanyByCurrency(tempAmt,parent.gCurrency,parent.ggAmTofMoneyNo, parent.gLocRndPolicyNo, "X")
				End If
			Next		
		End If
	End With
	Call SetSumItem	
End Function

'==========================================================================================
'   Event Name : DocCur_OnChange
'   Event Desc : 
'==========================================================================================
Sub DocCur_OnChange(byVal strDocCur, byVal Row)
    lgBlnFlgChgValue = True
	If Trim(strDocCur) = parent.gCurrency Then
		frm1.vspdData4.Col  = C_ExchRate_2
		frm1.vspdData4.Text = "1"
	Else
		Call FindExchRate(UniConvDateToYYYYMMDD(frm1.txtClsDt.text,parent.gDateFormat,""), UCase(Trim(strDocCur)),Row)
	End If
	Call SetSumItem()
End Sub

'=======================================================================================================
'   Function Name : FindExchRate
'   Function Desc : 1.날짜, Row를 입력받아 날짜에 해당하는 환율정보를 읽어온다.
'=======================================================================================================
Function FindExchRate(Byval strDate, Byval FromCurrency,Byval Row )
	Dim strSelect, strFrom, strWhere
	Dim arrTemp
	Dim strExchFg
	Dim strExchRate
	Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
	Dim IntRetCD		

	strSelect	= "b.minor_cd"
	strFrom		= "b_company a, b_minor b"
	strWhere	= "b.major_cd = " & FilterVar("a1004", "''", "S") & "  and	a.xch_rate_fg = b.minor_cd"
	If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then				
		arrTemp = Split(lgF0, chr(11))
		strExchFg =  arrTemp(0)
	End If

	If UCase(strExchFg) <> "D" Then 	' Fixed Exchange Rate
		strDate = Mid(strDate, 1, 6)
		strSelect	= "std_rate"
		strFrom		= "b_monthly_exchange_rate (noLock) "
		strWhere	= "from_currency =  " & FilterVar(FromCurrency , "''", "S") & ""
		strWhere	= strWhere & " And To_currency   =  " & FilterVar(parent.gCurrency , "''", "S") & ""
		strWhere	= strWhere & " And apprl_yrmnth  =  " & FilterVar(strDate , "''", "S") & ""

		If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then				
			arrTemp = Split(lgF0, chr(11))
			strExchRate =  arrTemp(0)
			frm1.vspdData4.row  = Row
			frm1.vspdData4.Col  = C_ExchRate_2
			frm1.vspdData4.Text = UNIConvNumPCToCompanyByCurrency(strExchRate, parent.gCurrency, parent.ggExchRateNo, parent.gLocRndPolicyNo, "X")
		Else
			IntRetCD = DisplayMsgBox("121600", "X", "X", "X")
		End If
	Else					' Floating Exchange Rate
		strSelect	= "Top 1 std_rate"
		strFrom		= "b_daily_exchange_rate (noLock) "
		strWhere	= "from_currency =  " & FilterVar(FromCurrency , "''", "S") & ""
		strWhere	= strWhere & " And To_currency   =  " & FilterVar(parent.gCurrency , "''", "S") & ""
		strWhere	= strWhere & " And apprl_dt  <= convert(char(21), " & FilterVar(strDate, "''", "S") & ", 20) order by apprl_dt desc"

		If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then				
			arrTemp = Split(lgF0, chr(11))
			strExchRate =  arrTemp(0)
			frm1.vspdData4.row  = Row
			frm1.vspdData4.Col  = C_ExchRate_2
			frm1.vspdData4.Text = UNIConvNumPCToCompanyByCurrency(strExchRate, parent.gCurrency, parent.ggExchRateNo, parent.gLocRndPolicyNo, "X")
		Else
			IntRetCD = DisplayMsgBox("121500", "X", "X", "X")
		End If
	End If
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
Function DbQuery()
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim StrNextKey      
    
    Call LayerShowHide(1)
    DbQuery = False
    
    Err.Clear                                                               '☜: Protect system from crashing

	Dim strVal

    With frm1
	    If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtClsNo="		& Trim(.txtClsNo.value)
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001 
			strVal = strVal & "&txtClsNo="		& Trim(.txtClsNo.value)
		End If

		' 권한관리 추가
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장
		strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인

		Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동
    End With
    
    DbQuery = True  
End Function

'=======================================================================================================
' Function Name : DbQuery2																				
' Function Desc : This function is data query and display												
'=======================================================================================================
Function DbQuery2(ByVal Row, ByVal pVspdData)
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
	Dim i
	Dim arrVal
	
	Dim iStrTempGlNo
	Dim iStrTempGlItemSeq
	Dim iStrOpenGlNo
	Dim iStrOpenGlItemSeq
	Dim iStrGlNo
	Dim intItemCnt
	Dim iStrItemSeq
	Dim iObjSpread1
	Dim iObjSpread2
	Dim iObjSpread3

	If pVspdData = C_VSPDDATA1 Then
		Set iObjSpread1 = frm1.vspdData
		Set iObjSpread2 = frm1.vspdData2
		Set iObjSpread3 = frm1.vspdData3
	ElseIf pVspdData = C_vspdData4 Then
		Set iObjSpread1 = frm1.vspdData4
		Set iObjSpread2 = frm1.vspdData5
		Set iObjSpread3 = frm1.vspdData6
	End If

	With frm1
		If pVspdData = C_VSPDDATA1 Then
			iStrGlNo		  = frm1.txtGlNo.value				'나중에 .htxtTempGlNo.value로 바꾸장 
			iStrTempGlNo	  = frm1.txtTempGlNo.value			'나중에 .htxtTempGlNo.value로 바꾸장 
		    iObjSpread1.Row	  = Row
		    
		    iObjSpread1.Col	  = C_ItemSeq1   
		    iStrTempGlItemSeq = Trim(iObjSpread1.Text)
		    
		    iObjSpread1.Col	  = C_GL_NO1  
		    iStrOpenGlNo	  = iObjSpread1.Text
		    
		    iObjSpread1.Col	  = C_OpenGlItemSeq  
		    iStrOpenGlItemSeq = iObjSpread1.Text
		ElseIf pVspdData = C_vspdData4 Then
			iStrGlNo		  = frm1.txtGlNo.value				'나중에 .htxtTempGlNo.value로 바꾸장 
			iStrTempGlNo	  = frm1.txtTempGlNo.value			'나중에 .htxtTempGlNo.value로 바꾸장 
					    
		    iObjSpread1.Row	  = Row

		    iObjSpread1.Col	  = C_ItemSeq
		    iStrTempGlItemSeq = Trim(iObjSpread1.Text)

		    iObjSpread1.Col	  = C_OpenGlNo_2
		    iStrOpenGlNo	  = iObjSpread1.Text
		    
		    iObjSpread1.Col	  = C_OpenGlItemSeq_2
		    iStrOpenGlItemSeq = iObjSpread1.Text
		End If

		iObjSpread2.ReDraw = False

		If pVspdData = C_VSPDDATA1 Then
			If CopyFromData(iStrTempGlItemSeq) = True Then
				Exit Function
			End If
		ElseIf pVspdData = C_VSPDDATA4 Then
			If CopyFromData2(iStrTempGlItemSeq) = True Then	
				If lgIntFlgMode = Parent.OPMD_CMODE Then
					Call SetSpread5Lock("Q","0","1","")
'					Call SetSpread4Color2()
				Else
					Call SetSpread5Lock("Q","0","1","")
'					Call SetSpread4Color2()				
'					Call CtrlSpreadLock("","",1,-1)
				End If				
				Exit Function
			End If
		End If
			
		Call LayerShowHide(1)
	
		DbQuery2 = False
		iObjSpread1.Row = Row

		If pVspdData = C_VSPDDATA1 Then
			iObjSpread1.Col = C_ItemSeq1
		Else			
			iObjSpread1.Col = C_ItemSeq
		End If
		
		If pVspdData = C_VSPDDATA1 Then
			If iStrOpenGlNo <> "" And iStrOpenGlItemSeq <> "" Then
				strSelect =				" C.DTL_SEQ, "		
				strSelect = strSelect & " A.CTRL_CD, "
				strSelect = strSelect & " A.CTRL_NM, "
				strSelect = strSelect & " C.CTRL_VAL, "
				strSelect = strSelect & " '', "		
				strSelect = strSelect & " Case     WHEN A.COLM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  Then " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  End, "
				strSelect = strSelect &	  iStrOpenGlItemSeq  & ", "
				strSelect = strSelect & " LTrim(ISNULL(A.TBL_ID,'')), "		
				strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_ID,'')), "
				strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM,'')), "		
				strSelect = strSelect & " LTrim(ISNULL(A.COLM_DATA_TYPE,'')), "
				strSelect = strSelect & " LTrim(ISNULL(A.DATA_LEN,'')), "
				strSelect = strSelect & " Case   	WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  Then " & FilterVar("DC", "''", "S") & "  "
				strSelect = strSelect & "		WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  Then " & FilterVar("D", "''", "S") & "  "
				strSelect = strSelect & "		WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  Then " & FilterVar("C", "''", "S") & "  End, " & iStrOpenGlItemSeq & ","		
				strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD,'')), CHAR(8) "
    	
				strFrom	=			" A_CTRL_ITEM 	 A (NOLOCK), "
				strFrom = strFrom & " A_ACCT_CTRL_ASSN   	B (NOLOCK), "
				strFrom = strFrom & " A_GL_DTL      			C (NOLOCK), "
				strFrom = strFrom & " A_GL_ITEM   			D (NOLOCK)	"
						
				strWhere =			  " D.GL_NO = " & FilterVar(UCase   (Trim(iStrOpenGlNo)), "''", "S")   
				strWhere = strWhere & " AND D.ITEM_SEQ	= " & iStrOpenGlItemSeq & " "
				strWhere = strWhere & " AND D.GL_NO		=  C.GL_NO  "
				strWhere = strWhere & " AND D.ITEM_SEQ  =  C.ITEM_SEQ "
				strWhere = strWhere & "	AND D.ACCT_CD	*= B.ACCT_CD "
				strWhere = strWhere & " AND C.CTRL_CD	*= B.CTRL_CD "		
				strWhere = strWhere & " AND C.CTRL_CD	= A.CTRL_CD "
				strWhere = strWhere & " ORDER BY C.DTL_SEQ "	
			Else
				strSelect =				" C.DTL_SEQ, "		
				strSelect = strSelect & " A.CTRL_CD, "
				strSelect = strSelect & " A.CTRL_NM, "
				strSelect = strSelect & " C.CTRL_VAL, "
				strSelect = strSelect & " '', "		
				strSelect = strSelect & " Case     WHEN A.COLM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  Then " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  End, "
				strSelect = strSelect &	  iStrTempGlItemSeq  & ", "
				strSelect = strSelect & " LTrim(ISNULL(A.TBL_ID,'')), "		
				strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_ID,'')), "
				strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM,'')), "		
				strSelect = strSelect & " LTrim(ISNULL(A.COLM_DATA_TYPE,'')), "
				strSelect = strSelect & " LTrim(ISNULL(A.DATA_LEN,'')), "
				strSelect = strSelect & " Case   	WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  Then " & FilterVar("DC", "''", "S") & "  "
				strSelect = strSelect & "		WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  Then " & FilterVar("D", "''", "S") & "  "
				strSelect = strSelect & "		WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  Then " & FilterVar("C", "''", "S") & "  End, " & iStrTempGlItemSeq & ","		
				strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD,'')), CHAR(8) "
    	
				strFrom	=			" A_CTRL_ITEM        A (NOLOCK), "
				strFrom = strFrom & " A_ACCT_CTRL_ASSN   	B (NOLOCK), "
				strFrom = strFrom & " A_TEMP_GL_DTL   		C (NOLOCK), "
				strFrom = strFrom & " A_TEMP_GL_ITEM      	D (NOLOCK)	"
						
				strWhere =			  " D.TEMP_GL_NO = " & FilterVar(UCase(iStrTempGlNo), "''", "S")   
				strWhere = strWhere & " AND D.ITEM_SEQ		= " & iStrTempGlItemSeq & " "
				strWhere = strWhere & " AND D.TEMP_GL_NO	=  C.TEMP_GL_NO  "
				strWhere = strWhere & " AND D.ITEM_SEQ		=  C.ITEM_SEQ "
				strWhere = strWhere & "	AND D.ACCT_CD		*= B.ACCT_CD "
				strWhere = strWhere & " AND C.CTRL_CD		*= B.CTRL_CD "		
				strWhere = strWhere & " AND C.CTRL_CD		= A.CTRL_CD "
				strWhere = strWhere & " ORDER BY C.DTL_SEQ "
			End If	
		Else
			If iStrGlNo <> ""  Then
				strSelect =				" C.DTL_SEQ, "		
				strSelect = strSelect & " A.CTRL_CD, "
				strSelect = strSelect & " A.CTRL_NM, "
				strSelect = strSelect & " C.CTRL_VAL, "
				strSelect = strSelect & " '', "		
				strSelect = strSelect & " Case     WHEN A.COLM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  Then " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  End, "
				strSelect = strSelect &	  iStrOpenGlItemSeq  & ", "
				strSelect = strSelect & " LTrim(ISNULL(A.TBL_ID,'')), "		
				strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_ID,'')), "
				strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM,'')), "		
				strSelect = strSelect & " LTrim(ISNULL(A.COLM_DATA_TYPE,'')), "
				strSelect = strSelect & " LTrim(ISNULL(A.DATA_LEN,'')), "
				strSelect = strSelect & " Case   	WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  Then " & FilterVar("DC", "''", "S") & "  "
				strSelect = strSelect & "		WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  Then " & FilterVar("D", "''", "S") & "  "
				strSelect = strSelect & "		WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  Then " & FilterVar("C", "''", "S") & "  End, " & iStrOpenGlItemSeq & ","		
				strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD,'')), CHAR(8) "
    	
				strFrom	=			" A_CTRL_ITEM 	 A (NOLOCK), "
				strFrom = strFrom & " A_ACCT_CTRL_ASSN   	B (NOLOCK), "
				strFrom = strFrom & " A_GL_DTL      			C (NOLOCK), "
				strFrom = strFrom & " A_GL_ITEM   			D (NOLOCK)	"
						
				strWhere =			  " D.GL_NO = " & FilterVar(UCase   (Trim(iStrGlNo)), "''", "S")   
				strWhere = strWhere & " AND D.ITEM_SEQ	= " & iStrOpenGlItemSeq & " "
				strWhere = strWhere & " AND D.GL_NO		=  C.GL_NO  "
				strWhere = strWhere & " AND D.ITEM_SEQ  =  C.ITEM_SEQ "
				strWhere = strWhere & "	AND D.ACCT_CD	*= B.ACCT_CD "
				strWhere = strWhere & " AND C.CTRL_CD	*= B.CTRL_CD "		
				strWhere = strWhere & " AND C.CTRL_CD	= A.CTRL_CD "
				strWhere = strWhere & " ORDER BY C.DTL_SEQ "	
			Else
				strSelect =				" C.DTL_SEQ, "		
				strSelect = strSelect & " A.CTRL_CD, "
				strSelect = strSelect & " A.CTRL_NM, "
				strSelect = strSelect & " C.CTRL_VAL, "
				strSelect = strSelect & " '', "		
				strSelect = strSelect & " Case     WHEN A.COLM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  Then " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  End, "
				strSelect = strSelect &	  iStrTempGlItemSeq  & ", "
				strSelect = strSelect & " LTrim(ISNULL(A.TBL_ID,'')), "		
				strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_ID,'')), "
				strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM,'')), "		
				strSelect = strSelect & " LTrim(ISNULL(A.COLM_DATA_TYPE,'')), "
				strSelect = strSelect & " LTrim(ISNULL(A.DATA_LEN,'')), "
				strSelect = strSelect & " Case   	WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  Then " & FilterVar("DC", "''", "S") & "  "
				strSelect = strSelect & "		WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  Then " & FilterVar("D", "''", "S") & "  "
				strSelect = strSelect & "		WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  Then " & FilterVar("C", "''", "S") & "  End, " & iStrTempGlItemSeq & ","		
				strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD,'')), CHAR(8) "
    	
				strFrom	=			" A_CTRL_ITEM        A (NOLOCK), "
				strFrom = strFrom & " A_ACCT_CTRL_ASSN   	B (NOLOCK), "
				strFrom = strFrom & " A_TEMP_GL_DTL   		C (NOLOCK), "
				strFrom = strFrom & " A_TEMP_GL_ITEM      	D (NOLOCK)	"
						
				strWhere =			  " D.TEMP_GL_NO = " & FilterVar(UCase(iStrTempGlNo), "''", "S")   
				strWhere = strWhere & " AND D.ITEM_SEQ		= " & iStrTempGlItemSeq & " "
				strWhere = strWhere & " AND D.TEMP_GL_NO	=  C.TEMP_GL_NO  "
				strWhere = strWhere & " AND D.ITEM_SEQ		=  C.ITEM_SEQ "
				strWhere = strWhere & "	AND D.ACCT_CD		*= B.ACCT_CD "
				strWhere = strWhere & " AND C.CTRL_CD		*= B.CTRL_CD "		
				strWhere = strWhere & " AND C.CTRL_CD		= A.CTRL_CD "
				strWhere = strWhere & " ORDER BY C.DTL_SEQ "
			End If
		End If	
		
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then   
	 		ggoSpread.Source = iObjSpread2
			ggoSpread.SSShowData lgF2By2

			Select Case pVspdData
				Case C_VSPDDATA1
					For lngRows = 1 To iObjSpread2.Maxrows
						iObjSpread2.row = lngRows	
						iObjSpread2.col = C_Tableid 
						If Trim(iObjSpread2.text) <> "" Then
							iObjSpread2.col = C_Tableid
							strTableid = iObjSpread2.text
							iObjSpread2.col = C_Colid
							strColid = iObjSpread2.text
							iObjSpread2.col = C_ColNm
							strColNm = iObjSpread2.text	
							iObjSpread2.col = C_MajorCd					
							strMajorCd = iObjSpread2.text	
							
							iObjSpread2.col = C_CtrlVal
							
							strNmwhere = strColid & " =  " & FilterVar(UCase(iObjSpread2.text), "''", "S")
							
							If Trim(strMajorCd) <> "" Then
								strNmwhere = strNmwhere & " AND MAJOR_CD =  " & FilterVar(strMajorCd , "''", "S") 
							End If				 
							
							If CommonQueryRs( strColNm , strTableid ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   
								iObjSpread2.col = C_CtrlValNm
								arrVal = Split(lgF0, Chr(11))  
								iObjSpread2.text = arrVal(0)
							End If
						End If								

						strVal = strVal & Chr(11) & iStrTempGlItemSeq

						iObjSpread2.Col = C_DtlSeq
						strVal = strVal & Chr(11) & iObjSpread2.Text
						
						iObjSpread2.Col = C_CtrlCd
						strVal = strVal & Chr(11) & iObjSpread2.Text
						
						iObjSpread2.Col = C_CtrlNm
						strVal = strVal & Chr(11) & iObjSpread2.Text
						    
						iObjSpread2.Col = C_CtrlVal
						strVal = strVal & Chr(11) & iObjSpread2.Text
						    
						iObjSpread2.Col = C_CtrlPB
						strVal = strVal & Chr(11) & iObjSpread2.Text
						    
						iObjSpread2.Col = C_CtrlValNm
						strVal = strVal & Chr(11) & iObjSpread2.Text
						    
						iObjSpread2.Col = C_Seq
						strVal = strVal & Chr(11) & iObjSpread2.Text
						    
						iObjSpread2.Col = C_Tableid
						strVal = strVal & Chr(11) & iObjSpread2.Text
						    
						iObjSpread2.Col = C_Colid
						strVal = strVal & Chr(11) & iObjSpread2.Text
						    
						iObjSpread2.Col = C_ColNm
						strVal = strVal & Chr(11) & iObjSpread2.Text
						    
						iObjSpread2.Col = C_Datatype
						strVal = strVal & Chr(11) & iObjSpread2.Text
						    
						iObjSpread2.Col = C_DataLen
						strVal = strVal & Chr(11) & iObjSpread2.Text
						    
						iObjSpread2.Col = C_DRFg
						strVal = strVal & Chr(11) & iObjSpread2.Text
						    
						iObjSpread2.Col = C_MajorCd
						strVal = strVal & Chr(11) & iObjSpread2.Text

						.vspdData2.Col = C_MajorCd + 1
						strVal = strVal & Chr(11) & lngRows
						strVal = strVal & Chr(11) & Chr(12)
					Next						
				Case C_vspdData4
					For lngRows = 1 To iObjSpread2.Maxrows
						iObjSpread2.row = lngRows	
						iObjSpread2.col = C_Tableid_2 
						If Trim(iObjSpread2.text) <> "" Then
							iObjSpread2.col = C_Tableid_2
							strTableid = iObjSpread2.text
							iObjSpread2.col = C_Colid_2
							strColid = iObjSpread2.text
							iObjSpread2.col = C_ColNm_2
							strColNm = iObjSpread2.text	
							iObjSpread2.col = C_MajorCd_2					
							strMajorCd = iObjSpread2.text	
								
							iObjSpread2.col = C_CtrlVal_2
								
							strNmwhere = strColid & " =  " & FilterVar(UCase(iObjSpread2.text), "''", "S")
								
							If Trim(strMajorCd) <> "" Then
								strNmwhere = strNmwhere & " AND MAJOR_CD =  " & FilterVar(strMajorCd , "''", "S") 
							End If				 
								
							If CommonQueryRs( strColNm , strTableid ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   
								iObjSpread2.col = C_CtrlValNm_2
								arrVal = Split(lgF0, Chr(11))  
								iObjSpread2.text = arrVal(0)
							End If
						End If												
					
						strVal = strVal & Chr(11) & iStrTempGlItemSeq

						iObjSpread2.Col = C_DtlSeq_2
						strVal = strVal & Chr(11) & iObjSpread2.Text
							
						iObjSpread2.Col = C_CtrlCd_2
						strVal = strVal & Chr(11) & iObjSpread2.Text
							
						iObjSpread2.Col = C_CtrlNm_2
						strVal = strVal & Chr(11) & iObjSpread2.Text
							    
						iObjSpread2.Col = C_CtrlVal_2
						strVal = strVal & Chr(11) & iObjSpread2.Text
							    
						iObjSpread2.Col = C_CtrlPB_2
						strVal = strVal & Chr(11) & iObjSpread2.Text
							    
						iObjSpread2.Col = C_CtrlValNm_2
						strVal = strVal & Chr(11) & iObjSpread2.Text
							    
						iObjSpread2.Col = C_Seq_2
						strVal = strVal & Chr(11) & iObjSpread2.Text
							    
						iObjSpread2.Col = C_Tableid_2
						strVal = strVal & Chr(11) & iObjSpread2.Text
							    
						iObjSpread2.Col = C_Colid_2
						strVal = strVal & Chr(11) & iObjSpread2.Text
							    
						iObjSpread2.Col = C_ColNm_2
						strVal = strVal & Chr(11) & iObjSpread2.Text
							    
						iObjSpread2.Col = C_Datatype_2
						strVal = strVal & Chr(11) & iObjSpread2.Text
							    
						iObjSpread2.Col = C_DataLen_2
						strVal = strVal & Chr(11) & iObjSpread2.Text
							    
						iObjSpread2.Col = C_DRFg_2
						strVal = strVal & Chr(11) & iObjSpread2.Text
							    
						iObjSpread2.Col = C_MajorCd_2
						strVal = strVal & Chr(11) & iObjSpread2.Text

						.vspdData2.Col = C_MajorCd_2 + 1
						strVal = strVal & Chr(11) & lngRows
						strVal = strVal & Chr(11) & Chr(12)
					Next						
				Case Else
			End Select					
			ggoSpread.Source = iObjSpread3			
			ggoSpread.SSShowData strVal	
		End If 		
		
		intItemCnt = iObjSpread1.MaxRows

		If pVspdData = C_VSPDDATA1 Then				

		ElseIf pVspdData = C_vspdData4 Then	
			If lgIntFlgMode = Parent.OPMD_CMODE Then
				Call SetSpread5Lock("Q","0","1","")
'				Call SetSpread4Color2()
			Else
				Call SetSpread5Lock("Q","0","1","")
'				Call SetSpread4Color2()
'				Call CtrlSpreadLock("","",1,1)
			End If				
		End If
	End With

	iObjSpread2.ReDraw = True
	Call LayerShowHide(0)
	
	DbQuery2 = True
	lgQueryOk = True
End Function

'========================================================================================================
' Function Name : InitData
' Function Desc : This function is data query and display
'========================================================================================================
Sub InitData(ByVal pVspdData)
	Dim intRow
	Dim intIndex
	
	If pVspdData = C_VSPDDATA1 Then
		With frm1.vspdData
			For intRow = 1 To .MaxRows
				.Row = intRow
				.col = C_DR_CR_FG1		:			intIndex = .value
				.col = C_DR_CR_NM1		:			.value = intindex
			Next
		End With
	ElseIf pVspdData = C_vspdData4 Then
		With frm1.vspdData4
			For intRow = 1 To .MaxRows
				.Row = intRow
				.col = C_DrCrFg			:			intIndex = .value
				.col = C_DrCrNm_2		:			.value = intindex
			Next
		End With
	End If
End Sub

'========================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    Dim pAP010M 
    Dim lngRows , itemRows
    Dim lGrpcnt
    DIM strVal 
    Dim tempItemSeq
	Dim	intRetCd
	Dim ii	
    Dim strNote
    Dim strItemDesc
    
    strNote = ""
    DbSave = False

    Call LayerShowHide(1)
	With frm1
		.txtFlgMode.value     = lgIntFlgMode
		.txtMode.value        = Parent.UID_M0002
	End With
	
    '-----------------------
    'Data manipulate area
    '-----------------------
    lGrpCnt = 1
    strVal = ""
    
    ggoSpread.Source = frm1.vspdData

    With frm1.vspdData
		For lngRows = 1 To .MaxRows
			.Row = lngRows
			.Col = 0
			
			If .Text <> ggoSpread.DeleteFlag Then
				'CRUD
				strVal = strVal & "C" & Parent.gColSep
				'CurrentRow 
				strVal = strVal & lngRows & Parent.gColSep
				
				.Col = C_ItemSeq1	
			    strVal = strVal & Trim(.Text) & Parent.gColSep				
'			    strVal = strVal & "" & Parent.gColSep
				.Col = C_ACCT_CD1		
			    strVal = strVal & Trim(.Text) & Parent.gColSep
				.Col = C_DR_CR_FG1		
			    strVal = strVal & Trim(.Text) & Parent.gColSep
			    
			    'OrgChangeId
			    .Col = C_OrgchangeId	
			     strVal = strVal & Trim(.Text) & Parent.gColSep
			     
			    .Col = C_DEPT_CD1	    
			    strVal = strVal & Trim(.Text) & Parent.gColSep
			    'DocCur
			    .Col = C_DOC_CUR1
			    strVal = strVal & Trim(.Text) & Parent.gColSep
			    .Col = C_XCH_RATE1	
			    strVal = strVal & UNICDbl(Trim(.Text)) & Parent.gColSep
			    
			    'VaTType
				strVal = strVal & "" & Parent.gColSep
			    .Col = C_OPEN_DOC_AMT1	
			    strVal = strVal & UNICDbl(Trim(.Text)) & Parent.gColSep
  				.Col = C_OPEN_AMT1	'6
				strVal = strVal & UNICDbl(Trim(.Text)) & Parent.gColSep
                'item_desc
                
				strVal = strVal & "" & Parent.gColSep	
				
				.Col = C_GL_DT1
			    strVal = strVal & UniConvDate(Trim(.Text)) & Parent.gColSep
			    .Col = C_GL_NO1
				strVal = strVal & Trim(.Text) & Parent.gColSep
				
			    .Col = C_OpenGlItemSeq 
			    strVal = strVal & Trim(.Text) & Parent.gColSep
			    
			    ' 미결관리여부 
			    strVal = strVal & "Y" & Parent.gColSep
		    
			    .Col = C_InternalCd 
			    strVal = strVal & Trim(.Text) & Parent.gColSep
			    .Col = C_Costcd 
			    strVal = strVal & Trim(.Text) & Parent.gRowSep	
	    		    
			End If		
		Next
    End With

    frm1.txtSpread.value  = strVal									'Spread Sheet 내용을 저장  
    
    strVal = ""    
    ggoSpread.Source = frm1.vspdData3

    With frm1.vspdData3      ' Dtl 저장 
		For itemRows = 1 To frm1.vspdData.MaxRows 
 		    frm1.vspdData.Row = itemRows
		    frm1.vspdData.Col = 0

		    If frm1.vspdData.Text <> ggoSpread.DeleteFlag Then	

				frm1.vspdData.Col = C_ItemSeq1
			    tempItemSEq = frm1.vspdData.Text  

			    For lngRows = 1 To .MaxRows
					.Row = lngRows
					.Col = C_ItemSeq1

					If .text = tempitemseq Then
						.Col = 0 					
						strVal = strVal & "C" & Parent.gColSep
						.Col = 1 		 			'ItemSEQ						        
						strVal = strVal & tempitemseq & Parent.gColSep

						.Col = 2 'C_DtlSeq + 1   				'Dtl SEQ					        
						strVal = strVal & Trim(.Text) & Parent.gColSep

						.Col = 3 'C_CtrlCd + 1		 		'관리항목코드							
						strVal = strVal & Trim(.Text) & Parent.gColSep

						.Col = 5 'C_CtrlVal + 1				'관리항목 Value 					        
						strVal = strVal & UCase   (Trim(.Text)) & Parent.gRowSep	
					End If			
		    	Next
		   End If
   		Next
    End With
    
    frm1.txtSpread6.value  = strVal						'Spread Sheet 내용을 저장 
    
    strval=""
 
    
    With frm1.vspdData4
		For lngRows = 1 To .MaxRows 
			.Row = lngRows
			.Col = 0
			If .Text <> ggoSpread.DeleteFlag Then
				'CRUD
				strVal = strVal & "C" & Parent.gColSep 
				'CurrentRow 
				strVal = strVal & lngRows & Parent.gColSep
				
			    .Col = C_ItemSeq	
			    strVal = strVal & Trim(.Text) & Parent.gColSep

				.Col = C_AcctCd		
			    strVal = strVal & Trim(.Text) & Parent.gColSep

				.Col = C_DrCrFg		
			    strVal = strVal & Trim(.Text) & Parent.gColSep

			    'OrgChangeId
			    strVal = strVal & frm1.hOrgChangeId.value & Parent.gColSep

				.Col = C_deptcd_2	    
			    strVal = strVal & Trim(.Text) & Parent.gColSep

			    'DocCur
			    .Col = C_DocCur_2
			    strVal = strVal & UCase(Trim(.Text)) & Parent.gColSep

			    .Col = C_ExchRate_2			    			    
			    strVal = strVal & UNICDbl(Trim(.Text)) & Parent.gColSep

			    'vat_type
				strVal = strVal & "" & Parent.gColSep

			    .Col = C_ItemAmt_2	
			    strVal = strVal & UNICDbl(Trim(.Text)) & Parent.gColSep

 				.Col = C_ItemLocAmt_2	'6
				strVal = strVal & UNICDbl(Trim(.Text)) & Parent.gColSep

			    .Col = C_ItemDesc_2	'7
				strVal = strVal & Trim(.Text) & Parent.gColSep
				
				'gl_dt	
				strVal = strVal & "" & Parent.gColSep	
				'gl_no
				.Col = C_OpenGlNo_2
				strVal = strVal & Trim(.Text) & Parent.gColSep
				'item_seq
			    .Col = C_OpenGlItemSeq_2
			    strVal = strVal & Trim(.Text) & Parent.gColSep
			    .Col = C_MgntFg_2     
			    strVal = strVal & Trim(.Text) & Parent.gColSep	 
			    'internal_cd 
			    strVal = strVal & "" & Parent.gColSep
				'cost_cd 
			    strVal = strVal & "" & Parent.gRowSep			    	 
			End If		
		Next
    End With

    frm1.txtSpread.value  = frm1.txtSpread.value & strVal										'Spread Sheet 내용을 저장    

    strVal = ""
    ggoSpread.Source = frm1.vspdData6

    With frm1.vspdData6      ' Dtl 저장 
		For itemRows = 1 To frm1.vspdData4.MaxRows 
 		    frm1.vspdData4.Row = itemRows
		    frm1.vspdData4.Col = 0

		    If frm1.vspdData4.Text <> ggoSpread.DeleteFlag Then	
				frm1.vspdData4.Col = C_ItemSeq
			    tempItemSEq = frm1.vspdData4.Text  
		        
			    For lngRows = 1 To .MaxRows
					.Row = lngRows
					.Col = C_ItemSeq
					
					If .text = tempitemseq Then
						.Col = 0 						
						strVal = strVal & "C" & Parent.gColSep
						
						.Col = 1 		 						'ItemSEQ							        
						strVal = strVal & tempitemseq & Parent.gColSep
						
						.Col = 2 'C_DtlSeq + 1   				'Dtl SEQ						        
						strVal = strVal & Trim(.Text) & Parent.gColSep
						
						.Col = 3 'C_CtrlCd + 1		 		'관리항목코드								
						strVal = strVal & Trim(.Text) & Parent.gColSep
						
						.Col = 5 'C_CtrlVal + 1				'관리항목 Value 						        
						strVal = strVal & UCase(Trim(.Text)) & Parent.gRowSep	
					End If			
		    	Next
			End If
   		Next
    End With

	With frm1    
		'권한관리추가 start
		.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
		.txthInternalCd.value =  lgInternalCd
		.txthSubInternalCd.value = lgSubInternalCd
		.txthAuthUsrID.value = lgAuthUsrID		
		'권한관리추가 end    

		.txtSpread6.value  = frm1.txtSpread6.value & strVal						'Spread Sheet 내용을 저장    
	End With

    Call ExecMyBizASP(frm1, BIZ_PGM_ID)											'저장 비지니스 ASP 를 가동 
    DbSave = True                                                           
End Function

'========================================================================================================
' Name : DbDelete
' Desc : This sub is called by FncDelete
'========================================================================================================
Function DbDelete()
	Dim strVal

    Err.Clear
    Call LayerShowHide(1)    

	DbDelete = False									'------ Developer Coding part (Start)  -------------------------------------------------------------- 

    Call LayerShowHide(1)
    
  	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003						'☜: 비지니스 처리 ASP의 상태 
	strVal = strVal 	& "&txtClsNo=" & Trim(frm1.txtClsNo.value)
	strVal = strVal		& "&hOrgChangeId=" & Trim(frm1.hOrgChangeId.value)
	strVal = strVal		& "&txtClsDt=" & Trim(frm1.txtClsDt.text)
	strVal = strVal		& "&txtGlNo=" & Trim(frm1.txtGlNo.value)

	Call RunMyBizASP(MyBizASP, strVal)											'☜: 비지니스 ASP 를 가동                                                      

    DbDelete = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Sub DbQueryOk	
	Dim ii
	Dim iIntRow
	Dim iIntIndex
	
	Call ClickTab1()
	
    lgBlnFlgChgValue = False
    
    lgIntFlgMode = Parent.OPMD_UMODE												'Indicates that current mode is Update mode        
    
	Call SetToolbar(C_MENU_UPD_TAB1)
	lgPreToolBarTab1 = C_MENU_UPD_TAB1
	lgPreToolBarTab2 = C_MENU_UPD_TAB2

    Call SetSpreadLock ("I", 0, 1, "" ) 

	SetSpreadColor 1, "0", 1, frm1.vspddata.maxrows
	
	Call SetSpread4Lock("Q", 1, 1, "")			
	Call SetSpread5Lock("Q", 1, 1, "")

	With frm1	
		For iIntRow = 1 To .vspdData.MaxRows					
			.vspdData.Row = iIntRow
			.vspdData.Col = C_DR_CR_FG1
			iIntIndex = .vspdData.value
			.vspdData.col = C_DR_CR_NM1
			.vspdData.value = iIntIndex					
		Next
			
		For iIntRow = 1 To .vspdData4.MaxRows					
			.vspdData4.Row = iIntRow
			.vspdData4.Col = C_DrCrFg
			iIntIndex = .vspdData4.value
			.vspdData4.col = C_DrCrNm_2
			.vspdData4.value = iIntIndex					
		Next
		
		If frm1.cboGlInputType.Value <> C_GLINPUTTYPE Then
			Call SetToolbar(C_MENU_NEW_TAB1)
			lgPreToolBarTab1 = C_MENU_NEW_TAB1
			lgPreToolBarTab2 = C_MENU_NEW_TAB1
		Else
			Call SetToolbar(C_MENU_UPD_TAB1)
			lgPreToolBarTab1 = C_MENU_UPD_TAB1
			lgPreToolBarTab2 = C_MENU_UPD_TAB2
		End If
		
		Call ggoOper.SetReqAttr(frm1.txtClsDt,	"Q")
		Call ggoOper.SetReqAttr(frm1.cboGlType,		"Q")
		Call ggoOper.SetReqAttr(frm1.txtTempGlNo,		"Q")
		Call ggoOper.SetReqAttr(frm1.txtGlNo,		"Q")
		Call ggoOper.SetReqAttr(frm1.txtdeptcd,		"Q")
		Call ggoOper.SetReqAttr(frm1.txtdesc,   "Q")			
		
		If .vspdData.MaxRows > 0 Then
			.vspdData.Row = 1
			.vspdData.Col = 1
			.hItemSeq.Value = .vspdData.Text
			Call DbQuery2(1, C_VSPDDATA1)
		End If
		
		If .vspdData4.MaxRows > 0 Then
			.vspdData4.Row = 1
			.vspdData4.Col = 1
			Call DbQuery2(1, C_vspdData4)
		End If
	End With

    Call txtDeptCd_OnChange()

    lgBlnFlgChgValue = False
    Call CancelResToreToolBar()
    
    Set gActiveElement = document.ActiveElement 
End Sub

'========================================================================================================
' Name : DbQueryOk2
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Sub DbQueryOk2	
	lgIntFlgMode     = parent.OPMD_UMODE
	lgBlnFlgChgValue = False
	SetSpreadColor 1, "01", 1, frm1.vspdData4.maxrows
	SetToolBar("1101000000001111")
	gSelframeFlg = TAB2
End Sub
		
'========================================================================================================
' Name : DbSaveOk
' Desc : Called by MB Area when save operation is successful
'========================================================================================================
Sub DbSaveOk(ByVal txtClsNo)
    lgBlnFlgChgValue = False
    
	Call ggoOper.ClearField(Document, "2")      '⊙: Condition field clear    
    Call InitVariables							'⊙: Initializes local global variables
	
	lgCurrentTabFg = TAB2
	DbQuery
End Sub

'=======================================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'=======================================================================================================
Function DbDeleteOk()													'삭제 성공후 실행 로직 
	Call FncNew()	
End Function


'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================

'==========================================================================================
'   Event Name : txtDeptCd_Change
'   Event Desc : 
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

	If Trim(frm1.txtClsDt.Text = "") Or Trim(frm1.txtDeptCd.value) = "" Then       
		Exit sub
    End If
    
    lgBlnFlgChgValue = True

	strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
	strFrom		=			 " b_acct_dept(NOLOCK) "		
	strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
	strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
	strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
	strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtClsDt.Text, Parent.gDateFormat,""), "''", "S") & "))"			
		
	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
		IntRetCD = DisplayMsgBox("124600","X","X","X")  
		frm1.txtDeptCd.value = ""
		frm1.txtDeptNm.value = ""
		frm1.hOrgChangeId.value = ""
	Else 
		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
		jj = Ubound(arrVal1,1)
					
		For ii = 0 To jj - 1
			arrVal2 = Split(arrVal1(ii), chr(11))			
			frm1.hOrgChangeId.value = Trim(arrVal2(2))
		Next	
	End If
End Sub

'==========================================================================================
'   Event Name : DeptCd_underChange(Byval strCode)
'   Event Desc : 
'==========================================================================================
Sub DeptCd_underChange(Byval strCode)
    Dim strSelect
	Dim strFrom
	Dim strWhere 	
    Dim IntRetCD 
    Dim DeptCd

    If Trim(frm1.txtClsDt.Text = "") Then    
		Exit sub
    End If
    
    lgBlnFlgChgValue = True

	strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
	strFrom		=			 " b_acct_dept(NOLOCK) "		
	strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(strCode)), "''", "S")
	strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
	strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
	strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtClsDt.Text, Parent.gDateFormat,""), "''", "S") & "))"			

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
		IntRetCD = DisplayMsgBox("124600","X","X","X")  

		frm1.vspdData4.Col = C_deptcd_2			
		frm1.vspdData4.Row = frm1.vspdData4.ActiveRow
		frm1.vspdData4.text = ""
		frm1.vspdData4.Col = C_deptnm_2		
		frm1.vspdData4.Row = frm1.vspdData4.ActiveRow	
		frm1.vspdData4.text = ""
	End If 
End Sub

'==========================================================================================
'   Event Name : cboGLType_OnChange
'   Event Desc : 
'==========================================================================================
Sub cboGLType_OnChange()
	dim	i		
	Dim IntRetCD	
	
	ggoSpread.Source = frm1.vspdData
	
	Select Case   UCase(Trim(frm1.cboGlType.value))
		Case   "01"						'입금전표로 바꾸면 차변이 입력되거나 현금계정이 입력되었는지 check한다.
			For i = 1 To  frm1.vspdData.maxrows
				frm1.vspddata.Row = i
				frm1.vspddata.col = C_Acctcd
				If  frm1.vspddata.text = lgCashAcct Then
					frm1.cboGlType.value = "03"
					Call CboGLType_ProtectGrid(frm1.cboGlType.value )		
					IntRetCD = DisplayMsgBox("113106", "X", "X", "X")					
					Exit sub
				End If
																			
				frm1.vspddata.col = C_DrCrFg
				If  Trim(frm1.vspddata.value) = "2" Then
					frm1.cboGlType.value = "03"
					Call CboGLType_ProtectGrid(frm1.cboGlType.value )		
					IntRetCD = DisplayMsgBox("113104", "X", "X", "X")					
					Exit Sub
				End If											
			Next				
			
			For i = 1 To  frm1.vspdData.maxrows
				frm1.vspddata.Row = i
				frm1.vspddata.col = C_DR_CR_FG1
				If Trim(frm1.vspddata.value) <> "1"  Then
					frm1.vspdData.value	= "1"
					frm1.vspddata.col = C_Dr_Cr_Nm1
					frm1.vspdData.value	= "1"
				End If
			Next
			
			Call CboGLType_ProtectGrid(frm1.cboGlType.value )		
		Case   "02"						'출금전표로 바꾸면 대변이 입력되거나 현금계정이 입력되었는지 check한다.	

			For i = 1 To  frm1.vspddata.maxrows
				frm1.vspddata.Row = i
				frm1.vspddata.col = C_Acctcd
				If  frm1.vspddata.text = lgCashAcct Then
					frm1.cboGlType.value = "03"
					Call CboGLType_ProtectGrid(frm1.cboGlType.value )
					IntRetCD = DisplayMsgBox("113106", "X", "X", "X")
					Exit sub
				End If
				
				frm1.vspddata.col = C_DR_CR_FG1
				If  Trim(frm1.vspddata.value) = "1" Then
					frm1.cboGlType.value = "03"
					Call CboGLType_ProtectGrid(frm1.cboGlType.value )		
					IntRetCD = DisplayMsgBox("113105", "X", "X", "X")					
					Exit sub				
				End If
			Next
				
			For i = 1 To  frm1.vspddata.maxrows
				frm1.vspddata.Row = i
				frm1.vspddata.col = C_DR_CR_FG1
				If Trim(frm1.vspddata.value) <> "2"  Then
					frm1.vspdData.value	= "2"
					frm1.vspddata.col = C_DR_CR_NM1
					frm1.vspdData.value	= "2"
				End If
			Next
			
			Call CboGLType_ProtectGrid(frm1.cboGlType.value )		
		Case   "03"						'대체로 바꾸면 Protect를 풀어준다.		
			Call CboGLType_ProtectGrid(frm1.cboGlType.value )		
	End Select	
	
	lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtClsDt_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtMBaseDt_Change()
	Dim strSelect
	Dim strFrom
	Dim strWhere
	Dim arrVal1
	Dim mm
	Dim dd
	Dim Temp
	Dim IntRetCD
	Dim strYear, strMonth, strDay, dtToday, EndDate, StartDate
	
	strSelect	=	  " ToP 1 card_mm, card_dd  "    		
	strFrom		=	  " a_open_acct_base "		
	strWhere	=	  " acct_base_no = 1"
			
	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then 	
		arrVal1 = Split(lgF2By2, Chr(11))			
		mm = UNICDbl(Trim(arrVal1(1)))
		dd = UNICDbl(Trim(arrVal1(2)))
	Else
	    IntRetCD = DisplayMsgBox("110900","X","X","X")  '☜ 바뀐부분 
	    Exit sub		
	End If 
	
    Call ExtractDateFrom(frm1.txtMBaseDt.Text,frm1.txtMBaseDt.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)

	Temp = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, dd)

    frm1.txtToBaseDt.text = UNIDateAdd("M", -mm, Temp, parent.gDateFormat)
    frm1.txtFromBaseDt.text = UNIDateAdd("D", +1,parent.UNIDateAdd("M", -(mm + 1), Temp, parent.gDateFormat),parent.gDateFormat)

    lgBlnFlgChgValue = True    
End Sub

'=======================================================================================================
'   Event Name : txtClsDt_DblClick(ButTon)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtClsDt_DblClick(ButTon)
    If ButTon = 1 Then
        frm1.txtClsDt.Action = 7                        
    End If
End Sub

'=======================================================================================================
'   Event Name : txtClsDt_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtClsDt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtMBaseDt_KeyDown(KeyCode, ShIft)
	If KeyCode = 13 Then Call FncQuery
End Sub

'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'==========================================================================================
'   Event Name :vspdData_KeyPress
'   Event Desc :
'==========================================================================================
Sub vspdData_KeyPress(index , KeyAscii )
    If KeyAscii = 13 Then
		Call fncquery()
    End If
End Sub
'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("0000111111")	
    gMouseClickStatus = "SPC"	'Split 상태코드

    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows <= 0 Then                                                    'If there is no data.
		Exit Sub
   	End If
   	
	If Row = 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort
			lgSortKey = 2
		Else
			ggoSpread.SSSort ,lgSortKey
			lgSortKey = 1
		End If    
        Exit Sub
	End If
End Sub

'========================================================================================================
'   Event Name : vspdData4_ColWidthChange
'   Event Desc : This function is called where spread sheet column width change
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================== 
' Event Name : vspdData_LeaveCell 
' Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'			   Item Row 변경시 관리항목 처리 
'			   hItemSeq에 Item Seq 입력 
'			   lgCurrRow에 Row Index 입력 
'==========================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
     If Row <> NewRow And NewRow > 0 Then
        With frm1        
            .vspdData.Row = NewRow
            .vspdData.Col = C_ItemSeq1
            .hItemSeq.value = .vspdData.Text
            .vspdData2.MaxRows = 0
        End With

        frm1.vspddata.Col = 0
        If frm1.vspddata.Text = ggoSpread.DeleteFlag Then
			Exit Sub
		End If
        		
		lgCurrRow = NewRow     		
		Call DbQuery2(lgCurrRow, C_VSPDDATA1)
    End If
End Sub

'========================================================================================== 
' Event Name : vspdData4_LeaveCell 
' Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'			   Item Row 변경시 관리항목 처리 
'			   hItemSeq에 Item Seq 입력 
'			   lgCurrRow에 Row Index 입력 
'==========================================================================================
Sub vspdData4_scriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	If Row <> NewRow And NewRow > 0 Then
        With frm1        
            .vspdData4.Row = NewRow
            .vspdData4.Col = C_ItemSeq
            .hItemSeq.value = .vspdData4.Text
             ggoSpread.Source = .vspdData5
             ggoSpread.ClearSpreadData   
        End With

        frm1.vspdData4.Col = 0
        If frm1.vspdData4.Text = ggoSpread.DeleteFlag Then
			Exit Sub
		End If

		lgCurrRow = NewRow
		Call DbQuery2(lgCurrRow, C_vspdData4)
    End If
End Sub

'==========================================================================================
' Event Name : vspdData4_ButTonClicked
' Event Desc : 버튼 컬럼을 클릭할 경우 
'==========================================================================================
Sub vspdData4_ButTonClicked(ByVal Col, ByVal Row, Byval ButTonDown)
	With frm1.vspdData4
		If Row > 0 And Col = C_AcctPopup  Then
			.Col = Col - 1
			.Row = Row
			
			Call OpenPopUp(.text, C_AcctPopup   , C_vspdData4)
		End If
		
		If Row > 0 And Col = C_DocCurPopup_2 Then
			.Col = Col - 1
			.Row = Row			
			Call OpenPopUp(.text, C_DocCurPopup_2, C_vspdData4)
		End If
		
		If Row > 0 And Col = C_deptPopup_2 Then
			.Col = Col - 1
			.Row = Row							
			Call OpenUnderDept(.Text, C_vspdData4)
    	End If
	End With
	
'	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData4,row,row,C_DocCur_2, C_ItemAmt_2,"A" ,"I","X","X")         
End Sub

'==========================================================================================
'   Event Name : vspdData4_ComboSelChange
'   Event Desc : Combo 변경 이벤트
'==========================================================================================
Sub vspdData4_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex	
	Dim tmpDrCrFg  
	With frm1
		.vspdData4.Row = Row
		
		Select Case   Col
			Case  C_DrCrNm_2
				.vspdData4.Col = Col
				intIndex = .vspdData4.Value
				.vspdData4.Col = C_DrCrFg
				.vspdData4.Value = intIndex				
				tmpDrCrFg = .vspdData4.text
				.vspddata4.Col = C_AcctCd

'				If AcctCheck2(frm1.vspdData4.text,frm1.cboGlType.value, tmpDrCrFg) = True Then
'					Call SetSpread4Color
'				End If

'				SetSpread4Color 	
			Case   C_DrCrFg
'				Call SetSpread4Color 						

				.vspdData4.Col = Col
				intIndex = .vspdData4.Value
				.vspdData4.Col = C_DrCrNm_2
				.vspdData4.Value = intIndex
		End Select		
	End With
End Sub

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc :
'========================================================================================================
Sub vspdData_MouseDown(ButTon , ShIft , x , y)
    If ButTon = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub  

'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData4_MouseDown(ButTon, ShIft, X, Y)
	If ButTon = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(Col , Row)
	lgBlnFlgChgValue = True
    ggoSpread.Source = frm1.vspddata
End Sub

'=======================================================================================================
'   Event Name : vspdData4_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData4_Change(ByVal pCol, ByVal pRow)	
	Dim tmpAcctCd
	Dim IntRetCD
	Dim iObjSpread
	Dim tmpDrCrFg
	Dim CurrencyCode
	Dim DeptCD   
	
	With frm1
		ggoSpread.Source = .vspdData4    
		ggoSpread.UpdateRow pRow    
		.vspdData4.Row = pRow   
    
		Select Case   pCol
		    Case    C_deptcd_2
				.vspdData4.Col	= C_deptcd_2
				 DeptCD			= .vspdData4.Text
				If DeptCd <> "" Then
					Call DeptCd_underChange(.vspdData4.text)
				End If
		    Case    C_AcctCd
			    .vspdData4.Col = 0
				If  .vspdData4.Text = ggoSpread.InsertFlag Then
					.vspdData4.Col = C_ItemSeq

					frm1.hItemSeq.value = .vspdData4.Text
					.vspdData4.Col = C_AcctCd								

					If Len(.vspdData4.Text) > 0 Then
						.vspdData4.Row = pRow					
						.vspdData4.Col = C_ItemSeq	 
						Call DeleteHsheet2(.vspdData4.Text)

						.vspdData4.Row = pRow	
						.vspdData4.Col = C_DrCrFg      		
						tmpDrCrFG = .vspdData4.text
						.vspdData4.Row = pRow
						.vspdData4.Col = C_AcctCd

						If AcctCheck2(.vspdData4.text, frm1.cboGlType.value, tmpDrCrFG) = True Then					
							Call Dbquery4(pRow)
							
							Call InputCtrlVal(pRow, C_vspdData4)
							Call SetSpread4Color()
						End If
					Else
						.vspdData4.Col = C_AcctNm            
						.vspdData4.Text = ""
					End If   
				End If
		  	Case 	C_DrCrFg      
				Call SetSumItem()
			Case  	C_DrCrNm_2   
				Call SetSumItem()
			Case	C_DocCur_2 	
				.vspdData4.Col = C_DocCur_2
				CurrencyCode = .vspdData4.Text
				If CurrencyCode <> "" Then
					IntRetCD = CommonQueryRs("Currency","B_CURRENCY"," Currency =  " & FilterVar(CurrencyCode , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
					If IntRetCD = False Then
						Call DisplayMsgBox("am0028","X","X","X")
						.vspdData4.Col	= C_DocCur_2
						.vspdData4.Text = ""
					Else
						.vspdData4.Col = C_ItemLocAmt_2
						.vspdData4.Text = ""
						.txtDrLocAmt2.text = ""
						.vspdData4.Col= C_DocCur_2 
						Call DocCur_OnChange(frm1.vspdData4.text, pRow )
						
						Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData4,pRow,pRow,C_DocCur_2, C_ExchRate_2,"D","I","X","X")
						Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData4,pRow,pRow,C_DocCur_2, C_ItemAmt_2 ,"A","I","X","X")
					End If
				End If
			Case  C_ExchRate_2    	
				Call FixDecimalPlaceByCurrency(frm1.vspdData4,Row,C_DocCur_2,C_ItemAmt_2,  "A" ,"X","X")
			Case  C_ItemAmt_2	
				Call SetSumItem()	
			Case  C_IsLAmtChange_2
				.vspdData4.Row = pRow
				.vspdData4.Col = C_IsLAmtChange_2  
				.vspdData4.Text = "Y"
				Call SetSumItem()	
		End Select
    End With
End Sub

'========================================================================================================
'   Event Name : vspdData4_Click
'   Event Desc : 컬럼을 클릭할 경우 발생
'=======================================================================================================
Sub vspdData4_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1110111111")	

    gMouseClickStatus = "SP2C"	'Split 상태코드
    Set gActiveSpdSheet = frm1.vspdData4
    
    If frm1.vspdData4.MaxRows <= 0 Then                                                    'If there is no data.
		Exit Sub
   	End If
   	
	If Row = 0 Then
		ggoSpread.Source = frm1.vspdData4
		If lgSortKey = 1 Then
			ggoSpread.SSSort
			lgSortKey = 2
		Else
			ggoSpread.SSSort ,lgSortKey
			lgSortKey = 1
		End If    
        Exit Sub
	End If
End Sub

'========================================================================================================
'   Event Name : vspdData4_MouseDown
'   Event Desc :
'========================================================================================================
Sub vspdData4_MouseDown(ButTon, ShIft, X, Y)
	If ButTon = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub 

'========================================================================================================
'   Event Name : vspdData4_ColWidthChange
'   Event Desc : This function is called where spread sheet column width change
'========================================================================================================
sub vspdData4_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.source = frm1.vspdData4
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData4_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData4_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData4
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("B")
End Sub

'========================================================================================================
'   Event Name : Radio2_onChange
'   Event Desc : 
'========================================================================================================
Function Radio2_onChange
    Dim lRow     
	Dim ii
	lgBlnFlgChgValue = True
	
	If frm1.Rb_ALL.checked = True Then
		frm1.vspddata.Col = 1 
		For ii = 1 To frm1.vspddata.MaxRows 
			frm1.vspddata.Row = ii
			If frm1.vspddata.text = 0 Then 
				frm1.vspddata.text = 1 
			End If 
		Next 
	Else 
		frm1.vspddata.Col = 1 
		For ii = 1 To frm1.vspddata.MaxRows 
			frm1.vspddata.Row = ii
			If frm1.vspddata.text = 1 Then 
				frm1.vspddata.text = 0 
			End If 
		Next 
	End If
End Function

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'===================================== PopResToreSpreadColumnInf()  ======================================
' Name : PopResToreSpreadColumnInf()
' Description : 컬럼의 순서정보를 복원함
'====================================================================================================
Sub  PopResToreSpreadColumnInf()
	Dim indx
	
	On Error Resume Next
	Err.Clear 		

	ggoSpread.Source = gActiveSpdSheet
	Select Case    Trim(UCase   (gActiveSpdSheet.Name))
		Case    "VSPDDATA" 
			Call ggoSpread.ResToreSpreadInf()
			Call InitSpreadSheet("A")
			Call InitComboBoxGrid("A")
			Call ggoSpread.ReOrderingSpreadData()
			Call InitData(C_VSPDDATA1)
			Call SetSpreadLock("Q", 1, 1, "")			
		Case   "VSPDDATA4" 
			Call PrevspdDataResTore2(gActiveSpdSheet)
			Call ggoSpread.ResToreSpreadInf()
			Call InitSpreadSheet("B")
			Call InitComboBoxGrid("B")
			Call ggoSpread.ReOrderingSpreadData()
			Call InitData(C_VSPDDATA4)
			Call SetSpread4Lock("Q", 1, 1, "")
		Case   "VSPDDATA5"
			Call PrevspdData2ResTore2(gActiveSpdSheet)   
			Call ggoSpread.ResToreSpreadInf()
			Call InitCtrlSpread2()			'관리항목 그리드 초기화
			Call ggoSpread.ReOrderingSpreadData()
			Call SetSpread4Color()  			
	End Select
	
	If frm1.vspdData5.MaxRows <= 0 Then
		Call DbQuery2(frm1.vspdData4.ActiveRow)
	End If			
End Sub

'========================================================================================================
' Name : fncResToreDbQuery2																				
' Desc : This function is data query and display												
'========================================================================================================
Function fncResToreDbQuery2(Row, CurrRow, Byval pInvalue1)
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

	fncResToreDbQuery2 = False

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
		
		strSelect =				" C.DTL_SEQ, "		
		strSelect = strSelect & " A.CTRL_CD, "
		strSelect = strSelect & " A.CTRL_NM, "
		strSelect = strSelect & " C.CTRL_VAL, "
		strSelect = strSelect & " '', "		
		strSelect = strSelect & " Case    WHEN A.COLM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  Then " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  End, "
		strSelect = strSelect &	  iStrTempGlItemSeq  & ", "
		strSelect = strSelect & " LTrim(ISNULL(A.TBL_ID,'')), "		
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_ID,'')), "
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM,'')), "		
		strSelect = strSelect & " LTrim(ISNULL(A.COLM_DATA_TYPE,'')), "
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_LEN,'')), "
		strSelect = strSelect & " Case  	WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  Then " & FilterVar("DC", "''", "S") & "  "
		strSelect = strSelect & "		WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  Then " & FilterVar("D", "''", "S") & "  "
		strSelect = strSelect & "		WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  Then " & FilterVar("C", "''", "S") & "  End, " & strItemSeq & ","		
		strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD,'')), CHAR(8) "
    	
		strFrom	=			" A_CTRL_ITEM  		A (NOLOCK), "
		strFrom = strFrom & " A_ACCT_CTRL_ASSN  	B (NOLOCK), "
		strFrom = strFrom & " A_GL_DTL  			C (NOLOCK), "
		strFrom = strFrom & " A_GL_ITEM  			D (NOLOCK)	"
					
		strWhere =			  " D.GL_NO = " & FilterVar(UCase  (Trim(iStrOpenGlNo)), "''", "S")   
		strWhere = strWhere & " AND D.ITEM_SEQ	= " & strItemSeq & " "
		strWhere = strWhere & " AND D.GL_NO		=  C.GL_NO  "
		strWhere = strWhere & " AND D.ITEM_SEQ  =  C.ITEM_SEQ "
		strWhere = strWhere & "	AND D.ACCT_CD	*= B.ACCT_CD "
		strWhere = strWhere & " AND C.CTRL_CD	*= B.CTRL_CD "		
		strWhere = strWhere & " AND C.CTRL_CD	= A.CTRL_CD "
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
			Call CopyFromData(strItemSeq)  '''
		End If

		Call LayerShowHide(0)
		Call ResToreToolBar()
	End With

	If Err.number = 0 Then
		fncResToreDbQuery2 = True
	End If
End Function

'===================================== PrevspdDataResTore2()  ========================================
' Name : PrevspdDataResTore2()
' Description : 그리드 복원시 관리항목 복원
'====================================================================================================
Sub PrevspdDataResTore2(pActiveSheetName)
	Dim indx, indx1

	For indx = 0 To frm1.vspdData4.MaxRows
        frm1.vspdData4.Row = indx
        frm1.vspdData4.Col = 0
		
		Select Case  frm1.vspdData4.Text			
			Case  ggoSpread.InsertFlag					
				frm1.vspdData4.Col = C_ItemSeq					
				Call DeleteHsheet2(frm1.vspdData4.Text)			'''		
			Case  ggoSpread.UpdateFlag		
				For indx1 = 0 To frm1.vspdData6.MaxRows					
					frm1.vspdData6.Row = indx1
					frm1.vspdData6.Col = 0
					Select Case  frm1.vspdData6.Text 
						Case  ggoSpread.UpdateFlag
							frm1.vspdData4.Col = C_ItemSeq
							frm1.vspdData6.Col = 1					
							If UCase  (Trim(frm1.vspdData4.Text)) = UCase  (Trim(frm1.vspdData6.Text)) Then
								Call DeleteHsheet2(frm1.vspdData4.Text)	 '''									
								Call FncResToreDbQuery22(indx, frm1.vspdData4.ActiveRow, frm1.txtClsNo.Value)
							End If
					End Select
				Next
			Case  ggoSpread.DeleteFlag
				Call fncResToreDbQuery22(indx, frm1.vspdData4.ActiveRow, frm1.txtClsNo.Value)
		End Select
	Next
	
	ggoSpread.Source = pActiveSheetName
End Sub

'========================================================================================================
' Name : fncResToreDbQuery22																			
' Desc : This function is data query and display												
'========================================================================================================
Function fncResToreDbQuery22(Row, CurrRow, Byval pInvalue1)
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

	fncResToreDbQuery22 = False

	Call DisableToolBar(parent.TBC_QUERY)
	Call LayerShowHide(1)
	
	With frm1
		.vspdData4.row = Row
	    .vspdData4.col = C_ItemSeq
		strItemSeq    = .vspdData4.Text

	    If Trim(strItemSeq) = "" Then
	        Exit Function
	    End If
    	
		Call LayerShowHide(1)
	
		DbQuery2 = False
		
		.vspdData4.Row = Row
		.vspdData4.Col = C_ItemSeq
		
		strSelect =				" C.DTL_SEQ, "		
		strSelect = strSelect & " A.CTRL_CD, "
		strSelect = strSelect & " A.CTRL_NM, "
		strSelect = strSelect & " C.CTRL_VAL, "
		strSelect = strSelect & " '', "		
		strSelect = strSelect & " Case    WHEN A.COLM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  Then " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  End, "
		strSelect = strSelect &	  iStrTempGlItemSeq  & ", "
		strSelect = strSelect & " LTrim(ISNULL(A.TBL_ID,'')), "		
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_ID,'')), "
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM,'')), "		
		strSelect = strSelect & " LTrim(ISNULL(A.COLM_DATA_TYPE,'')), "
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_LEN,'')), "
		strSelect = strSelect & " Case  	WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  Then " & FilterVar("DC", "''", "S") & "  "
		strSelect = strSelect & "		WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  Then " & FilterVar("D", "''", "S") & "  "
		strSelect = strSelect & "		WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  Then " & FilterVar("C", "''", "S") & "  End, " & strItemSeq & ","		
		strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD,'')), CHAR(8) "
    	
		strFrom	=			" A_CTRL_ITEM  		A (NOLOCK), "
		strFrom = strFrom & " A_ACCT_CTRL_ASSN  	B (NOLOCK), "
		strFrom = strFrom & " A_GL_DTL  			C (NOLOCK), "
		strFrom = strFrom & " A_GL_ITEM  			D (NOLOCK)	"
					
		strWhere =			  " D.GL_NO = " & FilterVar(UCase  (Trim(iStrOpenGlNo)), "''", "S")   
		strWhere = strWhere & " AND D.ITEM_SEQ	= " & strItemSeq & " "
		strWhere = strWhere & " AND D.GL_NO		=  C.GL_NO  "
		strWhere = strWhere & " AND D.ITEM_SEQ  =  C.ITEM_SEQ "
		strWhere = strWhere & "	AND D.ACCT_CD	*= B.ACCT_CD "
		strWhere = strWhere & " AND C.CTRL_CD	*= B.CTRL_CD "		
		strWhere = strWhere & " AND C.CTRL_CD	= A.CTRL_CD "
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
			ggoSpread.Source = .vspdData6
			ggoSpread.SSShowData strVal	
		End If 		

		If Row = CurrRow Then
			Call CopyFromData2(strItemSeq)  '''
		End If

		Call LayerShowHide(0)
		Call ResToreToolBar()
	End With

	If Err.number = 0 Then
		fncResToreDbQuery22 = True
	End If
End Function

'===================================== PrevspdDataResTore2()  ========================================
' Name : PrevspdData2ResTore2()
' Description : 그리드 복원시 관리항목 복원
'====================================================================================================
Sub PrevspdData2ResTore2(pActiveSheetName)
	Dim indx, indx1

	For indx = 0 To frm1.vspdData5.MaxRows
        frm1.vspdData5.Row    = indx
        frm1.vspdData5.Col    = 0

		If frm1.vspdData5.Text <> "" Then
			Select Case   frm1.vspdData5.Text
				Case   ggoSpread.InsertFlag
					frm1.vspdData5.Col = C_HItemSeq_2
					For indx1 = 0 To frm1.vspdData4.MaxRows
						frm1.vspdData4.Row = indx1
						frm1.vspdData4.Col = C_ItemSeq_2
						If frm1.vspdData4.Text = frm1.vspdData5.Text Then
							Call DeleteHsheet2(frm1.vspdData4.Text)
							ggoSpread.Source = frm1.vspdData4	
					        ggoSpread.EditUndo							
						End If
					Next
				Case  ggoSpread.UpdateFlag
					frm1.vspdData5.Col = C_HItemSeq_2
					For indx1 = 0 To frm1.vspdData4.MaxRows
						frm1.vspdData4.Row = indx1
						frm1.vspdData4.Col = C_ItemSeq_2
						If frm1.vspdData4.Text = frm1.vspdData5.Text Then
							Call DeleteHsheet2(frm1.vspdData4.Text)
							ggoSpread.Source = frm1.vspdData4
							ggoSpread.EditUndo
							Call fncResToreDbQuery22(indx1, frm1.vspdData4.ActiveRow, frm1.txtClsNo.Value) 
						End If
					Next
				Case  ggoSpread.DeleteFlag
			End Select
		Else
			If lgIntFlgMode = Parent.OPMD_CMODE Then
				frm1.vspddata5.Maxrows = 0
			End If					
		End If
	Next
	ggoSpread.Source = pActiveSheetName
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

'=======================================================================================================
'   Event Name : InputCtrlVal
'   Event Desc :
'=======================================================================================================  
Sub InputCtrlVal(ByVal pRow, ByVal pVspdData)
	Dim iObjSpread1
	Dim iObjSpread2
	Dim strAcctCd		
	Dim ii

	If pVspdData = C_VSPDDATA1 Then

	ElseIf pVspdData = C_vspdData4 Then
		Set iObjSpread1 = frm1.vspdData4
		Set iObjSpread2 = frm1.vspdData5
	Else
		Exit Sub
	End If
		
	lgBlnFlgChgValue = True

	If pVspdData = C_VSPDDATA1 Then

	ElseIf pVspdData = C_vspdData4 Then
		ggoSpread.Source = iObjSpread1
		iObjSpread1.Col = C_AcctCd
		iObjSpread1.Row = pRow	
		strAcctCd	= Trim(iObjSpread1.text)		
		
		iObjSpread1.Col = C_deptcd_2
		iObjSpread1.Row = pRow		
	End If
		
	If pVspdData = C_VSPDDATA1 Then

	ElseIf pVspdData = C_vspdData4 Then
		Call AuToInputDetail2(strAcctCd, Trim(iObjSpread1.text), frm1.txtClsDt.text, pRow)
		iObjSpread2.Col = C_CtrlVal_2
		For ii = 1 To iObjSpread2.MaxRows
			iObjSpread2.Row = ii					
			If Trim(iObjSpread2.text) <> "" Then
				Call CopyToHSheet4(iObjSpread1.ActiveRow,ii)			 			
			End If
		Next
	End If
End Sub	

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->
<BODY TABINDEX="-1" SCROLL="no">
<ForM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gIf"><img src="../../../CShared/image/table/seltab_up_left.gIf" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gIf" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gIf" align="right"><img src="../../../CShared/image/table/seltab_up_right.gIf" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gIf"><IMG height=23 src="../../../CShared/image/table/tab_up_left.gIf" width=9></td>
								<td background="../../../CShared/image/table/tab_up_bg.gIf" align="center" CLASS="CLSMTABP"><font color=white>미결정리내역</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gIf" align="right"><IMG height=23 src="../../../CShared/image/table/tab_up_right.gIf" width=10></td>
						    </TR>
						</TABLE>
					</TD>		
					<TD WIDTH=* align=right><A HREF="VBSCRIPT:OpenPopupOA()">미결연결</A>&nbsp;|&nbsp;
											<A href="vbscript:OpenPopuptempGL()">결의전표</A>&nbsp;|&nbsp;
											<A href="vbscript:OpenPopupGL()">회계전표</A>&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
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
					<TD HEIGHT=20 WIDTH="100%">
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>반제번호</TD>
									<TD CLASS="TD6" NOWRAP>
									<INPUT NAME="txtClsNo" ALT="반제번호" MAXLENGTH=18 STYLE="TEXT-ALIGN: left" tag ="12XXXU"><IMG align=Top name=btnCalType src="../../../CShared/image/btnPopup.gIf"  TYPE="BUTToN" onclick="vbscript: Call OpenClsPopup(frm1.txtClsNo.value,1)"></TD>								
									</TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD WIDTH="100%" HEIGHT=50 VALIGN=ToP >
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_20%>>					
								<TR>
									<TD CLASS=TD5 NOWRAP>전표일자</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtClsDt" CLASS=FPDTYYYYMMDD tag="22XXXU" Title="FPDATETIME" ALT="전표일자" id=fpDateTime2></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS=TD5 NOWRAP>전표형태</TD>								
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboGlType" tag="24" STYLE="WIDTH:82px:" ALT="전표형태"><OPTION VALUE="" selected></OPTION></SELECT></TD> 
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>부서</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE=TEXT NAME="txtDeptCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="22XXXU" ALT="부서"><IMG SRC="../../../CShared/image/btnPopup.gIf" NAME="btnDeptCd" ALIGN=Top TYPE="BUTToN" ONCLICK="vbscript:Call Opendept(frm1.txtDeptCd.Value)">
										<INPUT TYPE=TEXT NAME="txtDeptNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="24XXXU" ALT="부서명">
										<INPUT NAME="txtInternalCd" ALT="내부부서코드" TYPE="hidden" SiZE=7 MAXLENGTH=30  tag="14"  TABINDEX="-1">
									</TD>
									<TD CLASS=TD5 NOWRAP>전표입력경로</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboGlInputType" tag="24" STYLE="WIDTH:200px:" ALT="전표입력경로"><OPTION VALUE="" selected></OPTION></SELECT></TD>																
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>결의전표번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTempGlNo" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24XXXU" ALT="결의전표번호"> </TD>
									<TD CLASS="TD5" NOWRAP>전표번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtGlNo" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24XXXU" ALT="전표번호"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>비고</TD>
									<TD CLASS=TD656 NOWRAP COLSPAN=3><INPUT NAME="txtDesc" ALT="비고" MAXLENGTH="128" SIZE="70" tag="2X" ></TD>
								</TR>																	
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>		
					<TD WIDTH="100%" HEIGHT=* VALIGN=ToP>
						<DIV ID="TabDiv" SCROLL=no>		
							<TABLE <%=LR_SPACE_TYPE_60%> border=0>	
								<TR>
									<TD WIDTH=100% HEIGHT=100% valign=Top COLSPAN=4>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
									</TD>
								</TR>
								<TR>
									<TD COLSPAN=4>
										<TABLE <%=LR_SPACE_TYPE_20%>>
											<TR>
												<TD CLASS=TD656 WIDTH=* align=right COLSPAN=2><BUTToN NAME="btnCalc" CLASS="CLSSBTNCALC" ONCLICK="vbscript:FncBtnCalc()" Flag=1>자국금액계산</BUTToN>&nbsp;
												<TD CLASS=TD5 NOWRAP>차대합계(자국)</TD>
												<TD CLASS=TD6 NOWRAP>
													&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtDrLocAmt style="HEIGHT: 20px; LEFT: 0px; ToP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X2" ALT="차변합계(자국)" id=OBJECT3></OBJECT>');</SCRIPT>
													&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtCrLocAmt style="HEIGHT: 20px; LEFT: 0px; ToP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X2" ALT="대변합계(자국)" id=OBJECT4></OBJECT>');</SCRIPT>
												</TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
							</TABLE>
						</DIV>
						<!-- 두번째 탭 내용 -->  
						<DIV ID="TabDiv"  SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%> border=0>	
								<TR HEIGHT="66%">
									<TD WIDTH="100%" COLSPAN="4">
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData4 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=OBJECT5> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
									
								</TR>
								<TR>
									<TD COLSPAN=4>
										<TABLE <%=LR_SPACE_TYPE_20%>>
											<TR>
												<TD CLASS=TD656 WIDTH=* align=right COLSPAN=2><BUTToN NAME="btnCalc" CLASS="CLSSBTNCALC" ONCLICK="vbscript:FncBtnCalc()" Flag=1>자국금액계산</BUTToN>&nbsp;
												<TD CLASS=TD5 NOWRAP>차대합계(자국)</TD>
												<TD CLASS=TD6 NOWRAP>
													&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtDrLocAmt2 style="HEIGHT: 20px; LEFT: 0px; ToP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X2" ALT="차변합계(자국)" id=OBJECT3></OBJECT>');</SCRIPT>
													&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtCrLocAmt2 style="HEIGHT: 20px; LEFT: 0px; ToP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X2" ALT="대변합계(자국)" id=OBJECT4></OBJECT>');</SCRIPT>
												</TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
							    <TR HEIGHT="34%">
									<TD WIDTH="100%" COLSPAN="4">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" name=vspdData5 width="100%" tag="2" TITLE="SPREAD" id=OBJECT6> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IfRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX="-1"></IfRAME></TD> 
		<!--	 <TD WIDTH=100% HEIGHT=80%><IfRAME NAME="MyBizASP" WIDTH=100% HEIGHT=80% FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX="-1"></IfRAME></TD>  -->
	</TR>
</TABLE>
<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT TABINDEX="-1" TYPE=HIDDEN CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread3><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT TABINDEX="-1" TYPE=HIDDEN CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData3 WIDTH=0% HEIGHT=0% tag="23" TITLE="SPREAD" id=vaSpread3><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT> 
<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT TABINDEX="-1" TYPE=HIDDEN CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData6 WIDTH=0 HEIGHT=0 tag="23" TITLE="SPREAD" id=vaSpread6><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>


<TEXTAREA CLASS="HIDDEN" NAME="txtSpread"	tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA CLASS="HIDDEN" NAME="txtSpread2"	tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA CLASS="HIDDEN" NAME="txtSpread6"	tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtAuthorityFlag"	tag="24" TABINDEX="-1"><!--권한관리추가 -->
<INPUT TYPE=HIDDEN NAME="htxtDeptCd"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hOrgChangeId"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hItemSeq"			tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hmgnt_acct_cd"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtGlDt"			tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtDesc"			tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMode"			tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"		tag="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId"	tag="24" Tabindex="-1">
<TEXTAREA CLASS="HIDDEN" NAME="txtTempAPno" tag="X4" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=hidden NAME="txthAuthBizAreaCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthSubInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthUsrID"		tag="24" Tabindex="-1">
</Form>
<DIV ID="MousePT" NAME="MousePT">
<Iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></Iframe>
</DIV>
<ForM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname">
	<INPUT TYPE="HIDDEN" NAME="dbname">
	<INPUT TYPE="HIDDEN" NAME="filename">
	<INPUT TYPE="HIDDEN" NAME="condvar">
	<INPUT TYPE="HIDDEN" NAME="date">	
</Form>
</BODY>
</HTML>


