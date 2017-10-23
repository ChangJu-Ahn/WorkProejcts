<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID		    : A6105MA1
'*  4. Program Name         : 부가세신고디스켓CheckList조회 
'*  5. Program Desc         : 부가세신고디스켓CheckList조회 
'*  6. Component List       : +
'*  7. Modified date(First) : 2000/04/22
'*  8. Modified date(Last)  : 2002/07/31
'*  9. Modifier (First)     : Jong Hwan, Kim
'* 10. Modifier (Last)      : Hee Jung, Kim ; Nam Yo, Lee
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc"  -->								<!-- '⊙: 화면처리ASP에서 서버작업이 필요한 경우  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!-- '☆: 해당 위치에 따라 달라짐, 상대 경로  -->

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="javascript"	SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'☜: indicates that All variables must be declared in advance

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'⊙: 비지니스 로직 ASP명 
Const BIZ_PGM_ID = "a6105mb1.asp"			'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_ID2 = "a6105mb2.asp"			'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_ID3 = "a6105mb3.asp"			'☆: 비지니스 로직 ASP명 
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
'⊙: Grid Columns 
Dim C_BPRgstNO 
Dim C_PaperCnt 
Dim C_BlankCnt 
Dim C_NetAmt 
Dim C_VatAmt 
Dim C_Code 
Dim C_BPNM 
Dim C_IndTypeNM 
Dim C_IndClassNM 
 
Dim C_Title 
Dim C_BPCntSum 
Dim C_PaperCntSum 
Dim C_NetAmtSum 
Dim C_VatAmtSum 
 
'// ABOUT TAB2 ////////// 
'//spread2: B - Record 
Dim C_BRecord2 
Dim C_TaxOffice2 
Dim C_BPRgstNO2 
Dim C_BPNM2 
Dim C_BPPreNm2 
Dim C_ZipCode2 
Dim C_Addr 
Dim C_LoopCnt2 
 
'//spread3: C - Record 
 
Dim C_Title3 
Dim C_CRecord3 
Dim C_BPCntSum3 
Dim C_PaperCntSum3 
Dim C_NetAmtSum3 
Dim C_LoopCnt3 
 
 
'//spread4: D - Record 
Dim C_DRecord4 
Dim C_BPRgstNO4 
Dim C_BPNM4 
Dim C_PaperCnt4 
Dim C_NetAmt4 
Dim C_LoopCnt4 
 
 
'//spread5 : C - Record 
Dim C_CRecord5 
Dim C_Gigubun5 
Dim C_SingoGubun5 
Dim C_TaxOffice5 
Dim C_ReturnYear5 
Dim C_StartDt5 
Dim C_EndDt5 
Dim C_ReportDt5 
Dim C_BBPCntSum5 
Dim C_BPaperCntSum5 
Dim C_BNetAmtSum5 
Dim C_RBPCntSum5 
Dim C_RPaperCntSum5 
Dim C_RNetAmtSum5 
Dim C_HBPCntSum5 
Dim C_HPaperCntSum5 
Dim C_HNetAmtSum5 
Dim C_LoopCnt5 

'// ABOUT TAB3 //////////
'//Spread7 : C - Record
Dim  C_ExportNo7
Dim  C_FnDt7 
Dim  C_DocCur7 
Dim  C_XchRate7
Dim  C_DocAmt7
Dim  C_LocAmt7

'//Spread8 : B - Record
Dim  C_Title8 
Dim  C_CntSum8
Dim  C_DocSum8
Dim  C_LocSum8

Const TAB1 = 1																		'☜: Tab의 위치 
Const TAB2 = 2
Const TAB3 = 3


 '==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->	
'Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
'Dim lgIntGrpCount              ' Group View Size를 조사할 변수 
'Dim lgIntFlgMode               ' Variable is for Operation Status

'Dim lgStrPrevKey

'Dim lgLngCurRows

Dim lgBlnStartFlag				' 메세지 관련하여 프로그램 시작시점 Check Flag

'========================================================================================================= 
'Grid field vspdData1, vspdData3
'========================================================================================================= 
Dim lgRegNOPu                   '사업자등록번호발행분
Dim lgPreRgstNoPu               '주민등록번호발행분
Dim lgTotSum                    '합계
Dim lgExport					'수출하는재화
Dim lgEtcTax					'기타영세율
'========================================================================================================= 
'Grid2, Grid4 
'========================================================================================================= 

lgRegNOPu       = "사업자등록번호발행분"  '1
lgPreRgstNoPu   = "주민등록번호발행분"    '2
lgTotSum        = "합계"                '3
'========================================================================================================= 
'Grid8
'========================================================================================================= 
lgExport		= "수출하는재화"  '1
lgEtcTax		= "기타영세율"    '2


 '==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= 
 '----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 
Dim  IsOpenPop
'Dim  lgSortKey
Dim  gSelframeFlg
Dim lgFilePath
Dim lgFilePath2
Dim lgFilePath3

Dim strTmpGrid
Dim strTmpGrid1
Dim strTmpGrid2
Dim strTmpGrid3
Dim strTmpGrid4
Dim strTmpGrid7
Dim strTmpGrid8

 '==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey = 0                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count	
	lgSortKey = 1
	
End Sub
Sub initSpreadPosVariables()         '1.2 변수에 Constants 값을 할당 
    '⊙: Grid Columns
    '//spread: 
    C_BPRgstNO          = 1
    C_PaperCnt          = 2
    C_BlankCnt          = 3
    C_NetAmt            = 4
    C_VatAmt            = 5
    C_Code              = 6
    C_BPNM              = 7
    C_IndTypeNM         = 8
    C_IndClassNM        = 9
    '//spread1: 
    C_Title             = 0
    C_BPCntSum          = 1
    C_PaperCntSum       = 2
    C_NetAmtSum         = 3
    C_VatAmtSum         = 4

    '// ABOUT TAB2 //////////
    '//spread2: B - Record
    C_BRecord2          = 1
    C_TaxOffice2        = 2
    C_BPRgstNO2         = 3
    C_BPNM2             = 4
    C_BPPreNm2          = 5
    C_ZipCode2          = 6
    C_Addr              = 7
    C_LoopCnt2          = 8

    '//spread3: C - Record

    C_Title3            = 0
    C_CRecord3          = 1
    C_BPCntSum3         = 2
    C_PaperCntSum3      = 3
    C_NetAmtSum3        = 4
    C_LoopCnt3          = 5


    '//spread4: D - Record
    C_DRecord4          = 1
    C_BPRgstNO4         = 2
    C_BPNM4             = 3
    C_PaperCnt4         = 4
    C_NetAmt4           = 5
    C_LoopCnt4          = 6


    '//spread5 : C - Record = spread3 + a
    C_CRecord5          = 1
    C_Gigubun5          = 2
    C_SingoGubun5       = 3
    C_TaxOffice5        = 4
    C_ReturnYear5       = 5
    C_StartDt5          = 6
    C_EndDt5            = 7
    C_ReportDt5         = 8
    C_BBPCntSum5        = 9
    C_BPaperCntSum5     = 10
    C_BNetAmtSum5       = 11
    C_RBPCntSum5        = 12
    C_RPaperCntSum5     = 13
    C_RNetAmtSum5       = 14
    C_HBPCntSum5        = 15
    C_HPaperCntSum5     = 16
    C_HNetAmtSum5       = 17
    C_LoopCnt5          = 18


    '// ABOUT TAB3 //////////
    '//spread7:  - Record
	C_ExportNo7			= 1
	C_FnDt7				= 2
	C_DocCur7			= 3
	C_XchRate7			= 4
	C_DocAmt7			= 5
	C_LocAmt7			= 6

	'//Spread8 :  - Record
	C_Title8			= 0
	C_CntSum8			= 1
	C_DocSum8			= 2
	C_LocSum8			= 3
End Sub


'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 

Sub SetDefaultVal()

	'lgBlnStartFlag = False		' 메세지 관련하여 프로그램 시작시점 Check Flag
	
End Sub


'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("Q", "A","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("Q", "A","NOCOOKIE","MA") %>
End Sub


'=============================================== 2.2.3 InitSpreadSheet() ========================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()      '1.3 [initSpreadPosVariables] 호출 Logic 추가 
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021214",,Parent.gAllowDragDropSpread            '2.1 [Spreadinit]에 Source Version No. 와 DragDrop지원여부를 설정 

	With frm1.vspdData

		.MaxCols = C_IndClassNM + 1
		.MaxRows = 0


		.ReDraw = False

		Call GetSpreadColumnPos("A")
		ggoSpread.SSSetEdit  C_BPRgstNO, "사업자등록번호", 12, , , 20
		ggoSpread.SSSetFloat C_PaperCnt,"매수", 9, 0, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat C_BlankCnt,"공란", 9, parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat C_NetAmt,"공급가액", 18, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat C_VatAmt,"세액", 18, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit  C_Code, "주류코드", 10, , , 10
		ggoSpread.SSSetEdit  C_BPNM, "상호", 30, , , 30
		ggoSpread.SSSetEdit  C_IndTypeNM, "업태", 17, , , 20
		ggoSpread.SSSetEdit  C_IndClassNM, "종목", 25, , , 25
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)

		.ReDraw = True

	End With
	Call SetSpreadLock(0)
End Sub

Sub InitSpreadSheet1()
	Call initSpreadPosVariables()      '1.3 [initSpreadPosVariables] 호출 Logic 추가 
	ggoSpread.Source = frm1.vspdData1
	ggoSpread.Spreadinit "V20021214",,Parent.gAllowDragDropSpread            '2.1 [Spreadinit]에 Source Version No. 와 DragDrop지원여부를 설정 

	With frm1.vspdData1

		.ReDraw = False
		.MaxCols = C_VatAmtSum + 1

		.MaxRows = 0
		.MaxRows = 3
		Call GetSpreadColumnPos("B")
		ggoSpread.SSSetEdit  C_Title,      "",30, , , 25
		ggoSpread.SSSetFloat C_BPCntSum,   "거래처수", 20, 0, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat C_PaperCntSum,"매수",     20, 0, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat C_NetAmtSum,  "공급가액", 27, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat C_VatAmtSum,  "세액",     25, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

		.Row    = 1
		.Col    = 0 'C_Title
		.value  = lgRegNOPu      '"사업자등록번호발행분"
		.Col    = .MaxCols
		.text   = 1

		.Row    = 2
		.Col    = 0 'C_Title
		.value  = lgPreRgstNoPu    '"주민등록번호발행분"
		.Col    = .MaxCols
		.text   = 2

		.Row    = 3
		.Col    = 0 'C_Title
		.value  = lgTotSum       '"합계"
		.Col    = .MaxCols
		.text   = 3

		.ReDraw = True
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	End With
	Call SetSpreadLock(1)
End Sub


'=============================================== 2.2.3 InitSpreadSheet() ========================================
Sub InitSpreadSheet2()
	Call initSpreadPosVariables()      '1.3 [initSpreadPosVariables] 호출 Logic 추가 
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.Spreadinit "V20021214",,Parent.gAllowDragDropSpread            '2.1 [Spreadinit]에 Source Version No. 와 DragDrop지원여부를 설정 
	With frm1.vspdData2
		.MaxCols = C_LoopCnt2
		.MaxRows = 0
		.ReDraw = False
		Call GetSpreadColumnPos("C")
		ggoSpread.SSSetEdit C_BRecord2  , "레코드구분"    , 12, , , 12
		ggoSpread.SSSetEdit C_TaxOffice2, "세무서코드"    , 12, , , 12
		ggoSpread.SSSetEdit C_BPRgstNO2 , "사업자등록번호" , 18, , , 18
		ggoSpread.SSSetEdit C_BPNM2     , "법인명(상호)"  , 20, , , 20
		ggoSpread.SSSetEdit C_BPPreNm2  , "대표자(성명)"  , 15, , , 20
		ggoSpread.SSSetEdit C_ZipCode2  , "우편번호"      , 9, , , 20
		ggoSpread.SSSetEdit C_Addr      , "사업장소재지"   , 30, , , 50
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_LoopCnt2,C_LoopCnt2,True)
		.ReDraw = True
	End With
	Call SetSpreadLock(2)  
End Sub    

'=============================================== 2.2.3 InitSpreadSheet() ========================================
Sub InitSpreadSheet3()        
	ggoSpread.Source = frm1.vspdData3		
	ggoSpread.Spreadinit "V20021214",,Parent.gAllowDragDropSpread            '2.1 [Spreadinit]에 Source Version No. 와 DragDrop지원여부를 설정 
	With frm1.vspdData3	
		.MaxCols = C_LoopCnt3 + 1
		.MaxRows = 0
		.MaxRows = 3
		Call GetSpreadColumnPos("D")
		ggoSpread.SSSetEdit  C_Title3,      "",22, , , 25
		ggoSpread.SSSetEdit C_CRecord3      , "레코드구분", 20, , , 10
		ggoSpread.SSSetFloat C_BPCntSum3    , "매출처수"  , 20, 0, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat C_PaperCntSum3 , "계산서매수", 20, 0, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat C_NetAmtSum3   , "금액"     , 30, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

		.Row    = 1
		.Col    = 0 'C_Title3
		.value  = lgRegNOPu      '"사업자등록번호발행분"
		.Col    = .MaxCols
		.text   = 1

		.Row    = 2
		.Col    = 0 'C_Title3
		.value  = lgPreRgstNoPu    '"주민등록번호발행분"
		.Col    = .MaxCols
		.text   = 2

		.Row    = 3
		.Col    = 0 'C_Title3
		.value  = lgTotSum       '"합계"
		.Col    = .MaxCols
		.text   = 3

		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_LoopCnt3,C_LoopCnt3,True)
		.ReDraw = True
	End With
	frm1.vspdData5.MaxRows = 0
	frm1.vspdData5.MaxCols = C_LoopCnt5 + 1
	Call SetSpreadLock(3)  
End Sub   

'=============================================== 2.2.3 InitSpreadSheet() ========================================
Sub InitSpreadSheet4()
	'//spread4: D - Record
	ggoSpread.Source = frm1.vspdData4
	ggoSpread.Spreadinit "V20021214",,Parent.gAllowDragDropSpread            '2.1 [Spreadinit]에 Source Version No. 와 DragDrop지원여부를 설정 
	With frm1.vspdData4
		.MaxCols = C_LoopCnt4 + 1
		.MaxRows = 0
		.ReDraw = False

		Call GetSpreadColumnPos("E")
		ggoSpread.SSSetEdit C_DRecord4      , "레코드구분", 15, , , 15
		ggoSpread.SSSetEdit C_BPRgstNO4     , "사업자등록번호", 20, , , 20
		ggoSpread.SSSetEdit C_BPNM4         , "법인명(상호)", 20, , , 20		
		ggoSpread.SSSetFloat C_PaperCnt4    , "계산서매수", 20, 0, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat C_NetAmt4      , "금액", 30, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_LoopCnt4, C_LoopCnt4,True)
		.ReDraw = True
	End With
	frm1.vspdData6.MaxRows = 0
	frm1.vspdData6.MaxCols = C_LoopCnt4 + 1	
	Call SetSpreadLock(4)
End Sub   

'=============================================== 2.2.3 InitSpreadSheet() ========================================
Sub InitSpreadSheet5()        
	ggoSpread.Source = frm1.vspdData7
	ggoSpread.Spreadinit "V20021214",,Parent.gAllowDragDropSpread            '2.1 [Spreadinit]에 Source Version No. 와 DragDrop지원여부를 설정 
	With frm1.vspdData7	
		.MaxCols = C_LocAmt7 + 1
		.MaxRows = 0

		.ReDraw = False
		Call GetSpreadColumnPos("F")
		ggoSpread.SSSetEdit C_ExportNo7, "수출신고번호", 23, , , 20
		ggoSpread.SSSetEdit C_FnDt7, "선적(기)일자", 15, , , 20
		ggoSpread.SSSetEdit C_DocCur7, "거래통화", 10, , , 20
		'Call AppendNumberPlace("6","5","4")
		ggoSpread.SSSetFloat C_XchRate7,"환율",11, parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec													  
		'Call AppendNumberPlace("7","13","2")
		ggoSpread.SSSetFloat C_DocAmt7,"외화금액", 22, "A", ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		'Call AppendNumberPlace("8","15","0")
		ggoSpread.SSSetFloat C_LocAmt7,"원화금액", 22, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		.ReDraw = True
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	End With
	Call SetSpreadLock(5)
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
Sub InitSpreadSheet6()
	ggoSpread.Source = frm1.vspdData8
	ggoSpread.Spreadinit "V20021227",,Parent.gAllowDragDropSpread            '2.1 [Spreadinit]에 Source Version No. 와 DragDrop지원여부를 설정 
	With frm1.vspdData8	
		.ReDraw = False
		.MaxCols = C_LocSum8 + 1
		.MaxRows = 0
		.MaxRows = 3
		Call GetSpreadColumnPos("G")
		ggoSpread.SSSetEdit  C_Title8,      "",30, , , 25
		ggoSpread.SSSetFloat C_CntSum8,   "건수", 20, 0, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec	
		ggoSpread.SSSetFloat C_DocSum8,  "외화금액", 30, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat C_LocSum8,  "원화금액", 30, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

		.Row    = 1
		.Col    = 0 'C_Title8
		.value  = lgExport      '"수출하는재화"
		.Col    = .MaxCols
		.text   = 1

		.Row    = 2
		.Col    = 0 'C_Title8
		.value  = lgEtcTax    '"기타영세율"
		.Col    = .MaxCols
		.text   = 2

		.Row    = 3
		.Col    = 0 'C_Title8
		.value  = lgTotSum       '"합계"
		.Col    = .MaxCols
		.text   = 3
		.ReDraw = True
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	End With
	Call SetSpreadLock(6)  
End Sub


'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadLock(ByVal pvVal)

    Select Case pvVal
    Case 0
        With frm1.vspdData
            ggoSpread.Source = frm1.vspdData
            .ReDraw = False
			ggoSpread.SpreadLockWithOddEvenRowColor()
			ggoSpread.SSSetProtected	.MaxCols,-1,-1
            .ReDraw = True
        End With
    Case 1
        With frm1.vspdData1
            ggoSpread.Source = frm1.vspdData1
            .ReDraw = False
			ggoSpread.SpreadLockWithOddEvenRowColor()
			ggoSpread.SSSetProtected	.MaxCols,-1,-1
            .ReDraw = True
        End With

    Case 2
        With frm1.vspdData2
            ggoSpread.Source = frm1.vspdData2
            .ReDraw = False
			ggoSpread.SpreadLockWithOddEvenRowColor()
			ggoSpread.SSSetProtected	.MaxCols,-1,-1
           .ReDraw = True
        End With

    Case 3
        With frm1.vspdData3
            ggoSpread.Source = frm1.vspdData3
            .ReDraw = False
			ggoSpread.SpreadLockWithOddEvenRowColor()
			ggoSpread.SSSetProtected	.MaxCols,-1,-1
            .ReDraw = True
        End With

    Case 4
        With frm1.vspdData4
            ggoSpread.Source = frm1.vspdData4
            .ReDraw = False
			ggoSpread.SpreadLockWithOddEvenRowColor()
			ggoSpread.SSSetProtected	.MaxCols,-1,-1
            .ReDraw = True
        End With
    Case 5
        With frm1.vspdData7
            ggoSpread.Source = frm1.vspdData7
            .ReDraw = False
			ggoSpread.SpreadLockWithOddEvenRowColor()
			ggoSpread.SSSetProtected	.MaxCols,-1,-1
            .ReDraw = True
        End With
    Case 6
        With frm1.vspdData8
            ggoSpread.Source = frm1.vspdData8
            .ReDraw = False
			ggoSpread.SpreadLockWithOddEvenRowColor()
			ggoSpread.SSSetProtected	.MaxCols,-1,-1
            .ReDraw = True
        End With
    End Select
End Sub


'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadColor(ByVal lRow)
End Sub


 '==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Function InitComboBox()
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", " MAJOR_CD = " & FilterVar("A1003", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
	Call SetCombo2(frm1.cboIOFlag ,lgF0  ,lgF1  ,Chr(11))
	Call SetCombo2(frm1.cboIOFlag2 ,lgF0  ,lgF1  ,Chr(11))
End Function

Function OpenPopUp(Byval strCode, Byval iWhere)
End Function

Function SetPopUp(Byval arrRet, Byval iWhere)
End Function



'===============================================================================================
'   by Shin hyoung jae 
'	Name : ExtractFileName(strPath)
'	Description : ExtractFileName
'================================================================================================= 
Function ExtractFileName(byVal strPath)
	strPath = StrReverse(strPath)
	strPath = Left(strPath, InStr(strPath, "\") - 1)
	ExtractFileName = StrReverse(strPath)
End Function

'===============================================================================================
'   by Shin hyoung jae 
'	Name : GetOpenFilePath()
'	Description : GetTextFilePath	
'================================================================================================= 
Function GetOpenFilePath()
	Dim dlg
    Dim sPath

	On Error Resume Next
	Set dlg = CreateObject("uni2kCM.SaveFile")
	
	If Err.Number <> 0 Then
		Msgbox Err.Number & " : " & Err.Description
	End If
	
    sPath = dlg.GetOpenFilePath()

	If Err.Number <> 0 Then
		Msgbox Err.Number & " : " & Err.Description
	End If

	If gSelframeFlg = TAB1 Then 
		lgFilePath = sPath
		frm1.txtFileName.Value = ExtractFileName(sPath)
    ElseIf gSelframeFlg = TAB2 Then 
		lgFilePath2 = sPath
		frm1.txtFileName2.Value = ExtractFileName(sPath)
    ElseIf gSelframeFlg = TAB3 Then 
		lgFilePath3 = sPath
		frm1.txtFileName3.Value = ExtractFileName(sPath)
    End If
    Set dlg = Nothing
	frm1.hFileName.value = sPath		
End Function


Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            '⊙: Grid Columns
            C_BPRgstNO          = iCurColumnPos(1) 
            C_PaperCnt          = iCurColumnPos(2) 
            C_BlankCnt          = iCurColumnPos(3) 
            C_NetAmt            = iCurColumnPos(4) 
            C_VatAmt            = iCurColumnPos(5) 
            C_Code              = iCurColumnPos(6) 
            C_BPNM              = iCurColumnPos(7) 
            C_IndTypeNM         = iCurColumnPos(8) 
            C_IndClassNM        = iCurColumnPos(9) 

       Case "B"
            ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_Title             = iCurColumnPos(0)
            C_BPCntSum          = iCurColumnPos(1)
            C_PaperCntSum       = iCurColumnPos(2)
            C_NetAmtSum         = iCurColumnPos(3)
            C_VatAmtSum         = iCurColumnPos(4)

       Case "C"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            '// ABOUT TAB2 //////////
            '//spread2: B - Record
            C_BRecord2          = iCurColumnPos(1) 
            C_TaxOffice2        = iCurColumnPos(2) 
            C_BPRgstNO2         = iCurColumnPos(3) 
            C_BPNM2             = iCurColumnPos(4) 
            C_BPPreNm2          = iCurColumnPos(5) 
            C_ZipCode2          = iCurColumnPos(6) 
            C_Addr              = iCurColumnPos(7) 
            C_LoopCnt2          = iCurColumnPos(8) 

       Case "D"
            ggoSpread.Source = frm1.vspdData3
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)                                 
            '//spread3: C - Record                                  
            C_Title3            = iCurColumnPos(0)
            C_CRecord3          = iCurColumnPos(1)
            C_BPCntSum3         = iCurColumnPos(2)
            C_PaperCntSum3      = iCurColumnPos(3)
            C_NetAmtSum3        = iCurColumnPos(4)
            C_LoopCnt3          = iCurColumnPos(5)

       Case "E"
            ggoSpread.Source = frm1.vspdData4
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            '//spread4: D - RecordiCurColumnPos(7) 
            C_DRecord4          = iCurColumnPos(1)
            C_BPRgstNO4         = iCurColumnPos(2)
            C_BPNM4             = iCurColumnPos(3)
            C_PaperCnt4         = iCurColumnPos(4)
            C_NetAmt4           = iCurColumnPos(5)
            C_LoopCnt4          = iCurColumnPos(6)

       Case "F"
            ggoSpread.Source = frm1.vspdData7
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            '//spread7 :
			C_ExportNo7			= iCurColumnPos(1)
			C_FnDt7				= iCurColumnPos(2)
			C_DocCur7			= iCurColumnPos(3)  
			C_XchRate7			= iCurColumnPos(4) 
			C_DocAmt7			= iCurColumnPos(5)  
			C_LocAmt7			= iCurColumnPos(6)  

       Case "G"
            ggoSpread.Source = frm1.vspdData8
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            '//spread8 
			C_Title8			= iCurColumnPos(0)
			C_CntSum8			= iCurColumnPos(1)
			C_DocSum8			= iCurColumnPos(2)
			C_LocSum8			= iCurColumnPos(3)
	    End Select                    
End Sub                               
                                       

'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029                           '⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)

	Call ggoOper.LockField(Document, "N")

	Call InitSpreadSheet
	Call InitSpreadSheet1
	Call InitSpreadSheet2
	Call InitSpreadSheet3
	Call InitSpreadSheet4
	Call InitSpreadSheet5
	Call InitSpreadSheet6
	Call InitVariables
	'----------  Coding part  -------------------------------------------------------------
	Call InitComboBox
	Call SetDefaultVal
	Call SetToolbar("1100000000011111")										'⊙: 버튼 툴바 제어 
	frm1.txtFileName.disabled = True
	frm1.txtFileName2.disabled = True
	frm1.txtFileName3.disabled = True
	Call ClickTab1()
	frm1.cboIOFlag.focus 
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

 '==========================================  2.1.1 InitVariables()  ======================================
'	Name : subCompany()
'	Description : 조회내용의 헤더(세금계산서)
'========================================================================================================= 
Function subCompany(Byval pStrLine)
 
	' 자사정보 
	' 자료구분(1), 사업자등록번호(10), 상호(30), 성명(15), 사업장소재지(45), 업태(17), 종목(25), 
	' 거래기간(6), 거래기간(6), 작성일자(6), 공란(9)

	Dim Cnt, ColCnt
	Dim strChr

	strChr = ""
	ColCnt = 0
	For Cnt = 1 To Len(pStrLine)
		If Asc(Mid(pStrLine, Cnt, 1)) >= 0 Then
			ColCnt = ColCnt + 1		' 영숫자 
		Else
			ColCnt = ColCnt + 2		' 한글 
		End If
		
		strChr = strChr & Mid(pStrLine, Cnt, 1)

		Select Case ColCnt
			Case 1		' 자료구분(1)
				strChr = ""
			Case 11		' 사업자등록번호(10)
				frm1.txtRegNo.value = strChr
				strChr = ""
			Case 41		' 상호(30)
				frm1.txtBizAreaNm.value = strChr
				strChr = ""
			Case 56		' 성명(15)
				frm1.txtRepreNm.value = strChr
				strChr = ""
			Case 101	' 사업장소재지(45)
				frm1.txtAddr.value = strChr
				strChr = ""
			Case 118	' 업태(17)
				frm1.txtIndType.value = strChr
				strChr = ""
			Case 143	' 종목(25)
				frm1.txtIndClass.value = strChr
				strChr = ""
			Case 149	' 거래기간(6)
				frm1.txtStartDt.value = strChr
				strChr = ""
			Case 155	' 거래기간(6)
				frm1.txtEndDt.value = strChr
				strChr = ""
			Case 161	' 작성일자(6)
				frm1.txtReportDt.value = strChr
				strChr = ""
			Case 170	' 공란(9)
				strChr = ""
		End Select		
		If ColCnt >= 170 Then Exit For
	Next
End Function

'==========================================  2.1.1 subCompany2A()  ======================================
'	Name : subCompany2A()
'	Description : 조회내용의 헤더(계산서) A-Record
'========================================================================================================= 

Function subCompany2A(Byval pStrLine)
 
	Dim Cnt, ColCnt
	Dim strChr

	strChr = ""
	ColCnt = 0
	For Cnt = 1 To Len(pStrLine)
		If Asc(Mid(pStrLine, Cnt, 1)) >= 0 Then
			ColCnt = ColCnt + 1		' 영숫자 
		Else
			ColCnt = ColCnt + 2		' 한글 
		End If
		
		strChr = strChr & Mid(pStrLine, Cnt, 1)

		Select Case ColCnt
			Case 1		' 레코드구분(1)
				strChr = ""
			Case 4		' 세무서(3)
				strChr = ""
			Case 12		' 제출년월일(9)
				frm1.txtReportDt2.value = strChr
				strChr = ""
			Case 13		' 제출자구분(1)
				strChr = ""
			Case 19		' 세무대리인관리번호(6)
				strChr = ""
			Case 29		' 사업자등록번호(10)
				frm1.txtRegNo2.value = strChr
				strChr = ""
			Case 69		' 법인명(상호)(10)
				frm1.txtBizAreaNm2.value = strChr
				strChr = ""
			Case 82		' 주민등록번호(13)
				frm1.txtPreRgstNo2.value = strChr
				strChr = ""
			Case 112	' 대표자(성명)(30)
				frm1.txtRepreNm2.value = strChr
				strChr = ""
			Case 122	' 소재지우편번호(10)
				frm1.txtZipCode2.value = strChr
				strChr = ""
			Case 192	' 소재지주소(70)
				frm1.txtAddr2.value = strChr
				strChr = ""
			Case 207	' 전화번호(15)
				frm1.txtTelno2.value = strChr
				strChr = ""
			Case 212	' 제출건수계(5)
				strChr = ""
			Case 215	' 한글코드종류(3)
				strChr = ""
			Case 230	' 작성일자(6)
				strChr = ""
		End Select
		
		If ColCnt >= 230 Then Exit For
	Next

End Function

'==========================================  2.1.1 subCompany2B()  ======================================
'	Name : subCompany2B()
'	Description : 조회내용의 헤더(계산서) B-Record
'========================================================================================================= 

Function subCompany2B(Byval pStrLine, Byval BGubunCnt)
 
	Dim Cnt, ColCnt
	Dim strChr
	Dim LastAmt	
	Dim strBRecord2 , strTaxOffice2 , strBPRgstNO2 , strBPNM2 , strBPPreNm2 , strZipCode2 , strAddr , strLoopCnt2 

	strChr = ""
	ColCnt = 0
	
	For Cnt = 1 To Len(pStrLine)
		If Asc(Mid(pStrLine, Cnt, 1)) >= 0 Then
			ColCnt = ColCnt + 1		' 영숫자 
		Else
			ColCnt = ColCnt + 2		' 한글 
		End If
		
		strChr = strChr & Mid(pStrLine, Cnt, 1)

		Select Case ColCnt
			Case 1		' 레코드구분(1)
                strBRecord2 = strChr
				strChr = ""
			Case 4		' 세무서(3)
                strTaxOffice2 = strChr
				strChr = ""
			Case 10		' 일련번호(6)
				strChr = ""
			Case 20		' 사업자등록번호(10)
                strBPRgstNO2 = strChr
				strChr = ""
			Case 60		' 법인명(상호)(40)
                strBPNM2 = strChr
				strChr = ""
			Case 90		'대표자(성명)(30)
                strBPPreNm2 = strChr
				strChr = ""
			Case 100	'사업장우편번호(10)
                strZipCode2 = strChr
				strChr = ""
			Case 170	' 사업장소재지(70)
                strAddr = strChr
				strChr = ""
			Case 230	'공란(60)
				strChr = ""
		End Select	
		If ColCnt >= 230 Then Exit For
	Next
    strTmpGrid2 = strTmpGrid2 & chr(11) & strBRecord2 
    strTmpGrid2 = strTmpGrid2 & chr(11) & strTaxOffice2 
    strTmpGrid2 = strTmpGrid2 & chr(11) & strBPRgstNO2 
    strTmpGrid2 = strTmpGrid2 & chr(11) & strBPNM2 
    strTmpGrid2 = strTmpGrid2 & chr(11) & strBPPreNm2 
    strTmpGrid2 = strTmpGrid2 & chr(11) & strZipCode2 
    strTmpGrid2 = strTmpGrid2 & chr(11) & strAddr 
    strTmpGrid2 = strTmpGrid2 & chr(11) & BGubunCnt & chr(11) & chr(12)
End Function

 '==========================================  2.1.1 subPayment()  ======================================
'	Name : subPayment()
'	Description : 매출처정보 (세금계산서)
'========================================================================================================= 

Function subPayment(Byval pStrLine,Byval pStrLineNo)
	
	' 매출처정보 
	' 자료구분(1), 사업자등록번호(10), 일련번호(4), 사업자등록번호(10), 상호(30), 업태(17), 종목(25), 매수(7), 
	' 공란수(2), 공급가액(14), 세액(13), 주류도매(1), 주류소매(1), 권번호(4), 제출처(3), 공란(28)

	Dim Cnt, ColCnt
	Dim strChr
	dim LastAmt	
	Dim strBPRgstNO,strPaperCnt, strBlankCnt, strNetAmt, strVatAmt, strCode, strBPNM, strIndTypeNM, strIndClassNM

	strChr = ""
	ColCnt = 0
	
	For Cnt = 1 To Len(pStrLine)
		If Asc(Mid(pStrLine, Cnt, 1)) >= 0 Then
			ColCnt = ColCnt + 1		' 영숫자 
		Else
			ColCnt = ColCnt + 2		' 한글 
		End If
		
		strChr = strChr & Mid(pStrLine, Cnt, 1)

		Select Case ColCnt
			Case 1		' 자료구분(1)
				strChr = ""
			Case 11		' 사업자등록번호(10)
				strChr = ""
			Case 15		' 일련번호(4)
				strChr = ""
			Case 25		' 사업자등록번호(10)
                strBPRgstNO = strChr
				strChr = ""
			Case 55		' 상호(30)
                strBPNM = strChr
				strChr = ""
			Case 72		' 업태(17)
                strIndTypeNM  = strChr
				strChr = ""
			Case 97		' 종목(25)
                strIndClassNM  = strChr
				strChr = ""
			Case 104	' 매수(7)
                strPaperCnt  = strChr
				strChr = ""
			Case 106	' 공란수(2)
                strBlankCnt  = strChr
				strChr = ""
			Case 120	' 공급가액(14)
				If (Right(strChr,1) <= "9") Then
                    strNetAmt = strChr
				Else	
					If Right(strChr,1) = "}" Then
						LastAmt = "0"
					Else
						LastAmt = CHR(ASC(Right(strChr,1)) - 25)
					End if
                    strNetAmt = "-" & mid(strChr,1,len(strChr)-1) & LastAmt
				End If	
				strChr = ""
			Case 133	' 세액(13)
				If (Right(strChr,1) <= "9") Then
                    strVatAmt = strChr
				Else	
					If Right(strChr,1) = "}" Then
						LastAmt = "0"
					Else
						LastAmt = CHR(ASC(Right(strChr,1)) - 25)
					End if
                    strVatAmt = "-" & mid(strChr,1,len(strChr)-1) & LastAmt
				End If	
				strChr = ""
			Case 134	' 주류도매(1)
				strChr = strChr & " / "
			Case 135	' 주류소매(1)
                strCode = strChr
				strChr = ""
			Case 139	' 권번호(4)
				strChr = ""
			Case 142	' 제출처(3)
				strChr = ""
			Case 170	' 공란(28)
				strChr = ""
		End Select
		
		If ColCnt >= 170 Then Exit For
	Next

    strTmpGrid = strTmpGrid & chr(11) & strBPRgstNO 
    strTmpGrid = strTmpGrid & chr(11) & strPaperCnt 
    strTmpGrid = strTmpGrid & chr(11) & strBlankCnt 
    strTmpGrid = strTmpGrid & chr(11) & strNetAmt 
    strTmpGrid = strTmpGrid & chr(11) & strVatAmt 
    strTmpGrid = strTmpGrid & chr(11) & strCode 
    strTmpGrid = strTmpGrid & chr(11) & strBPNM 
    strTmpGrid = strTmpGrid & chr(11) & strIndTypeNM 
    strTmpGrid = strTmpGrid & chr(11) & strIndClassNM 
    strTmpGrid = strTmpGrid & chr(11) & pStrLineNo & chr(11) & chr(12)	
End Function



Function InitRowColVale(strSpread,iRegNOPuRowNo,iPreRgstNoPuRowNo,iTotSumRowNo)
    Dim ii , jj, strRowNo, strRowValue
    If strSpread = "A" then
        For ii = 1 to frm1.vspdData1.MaxRows
            For jj = 1 To frm1.vspdData1.MaxCols
                frm1.vspdData1.Row  = ii
                frm1.vspdData1.Col  = jj
                frm1.vspdData1.value = ""
                If jj = frm1.vspdData1.MaxCols Then
                    frm1.vspdData1.Row  = ii
                    frm1.vspdData1.Col  = jj
                    frm1.vspdData1.value = ii
                End if               
            Next
            frm1.vspdData1.Row  = ii
            frm1.vspdData1.Col  = 0
            strRowValue         = frm1.vspdData1.value
            Select Case Trim(strRowValue)
            Case Trim(lgRegNOPu)
                iRegNOPuRowNo      = ii
            Case Trim(lgPreRgstNoPu)
                iPreRgstNoPuRowNo  = ii
            Case Trim(lgTotSum)
                iTotSumRowNo       = ii
            End Select
        Next
    ElseIf strSpread = "B" then
        For ii = 1 to frm1.vspdData3.MaxRows
            For jj = 1 To frm1.vspdData3.MaxCols
                frm1.vspdData3.Row  = ii
                frm1.vspdData3.Col  = jj
                frm1.vspdData3.value = ""
                If jj = frm1.vspdData3.MaxCols Then
                    frm1.vspdData3.Row  = ii
                    frm1.vspdData3.Col  = jj
                    frm1.vspdData3.value = ii
                End if               
            Next
            frm1.vspdData3.Row  = ii
            frm1.vspdData3.Col  = 0
            strRowValue         = frm1.vspdData3.value
            Select Case Trim(strRowValue)
            Case Trim(lgRegNOPu)
                iRegNOPuRowNo      = ii
            Case Trim(lgPreRgstNoPu)
                iPreRgstNoPuRowNo  = ii
            Case Trim(lgTotSum)
                iTotSumRowNo       = ii
            End Select
        Next
    ElseIf strSpread = "C" then
        For ii = 1 to frm1.vspdData8.MaxRows
            For jj = 1 To frm1.vspdData8.MaxCols
                frm1.vspdData8.Row  = ii
                frm1.vspdData8.Col  = jj
                frm1.vspdData8.value = ""
                If jj = frm1.vspdData8.MaxCols Then
                    frm1.vspdData8.Row  = ii
                    frm1.vspdData8.Col  = jj
                    frm1.vspdData8.value = ii
                End if               
            Next
            frm1.vspdData8.Row  = ii
            frm1.vspdData8.Col  = 0
            strRowValue         = frm1.vspdData8.value
            Select Case Trim(strRowValue)
            Case Trim(lgExport)
                iRegNOPuRowNo      = ii
            Case Trim(lgEtcTax)
                iPreRgstNoPuRowNo  = ii
            Case Trim(lgTotSum)
                iTotSumRowNo       = ii
            End Select
        Next
    End If
End Function



Function subPaymentSum(Byval pStrLine)
 
	' 매출처합계정보 
	' 자료구분(1), 사업자등록번호(10), 전체매출처수(7), 세금계산서매수(7), 공급가액(15), 세액(14), 
	' 사업자별매출처수(7),  세금계산서매수(7), 공급가액(15), 세액(14), 
	' 주민별매출처수(7),  세금계산서매수(7), 공급가액(15), 세액(14), 공란(30) 

	Dim Cnt, ColCnt
	Dim strChr
	Dim LastAmt
	Dim iRegNOPuRowNo, iPreRgstNoPuRowNo, iTotSumRowNo

	Call InitRowColVale("A", iRegNOPuRowNo, iPreRgstNoPuRowNo, iTotSumRowNo)

    strChr = ""
	ColCnt = 0
	For Cnt = 1 To Len(pStrLine)
		If Asc(Mid(pStrLine, Cnt, 1)) >= 0 Then
			ColCnt = ColCnt + 1		' 영숫자 
		Else
			ColCnt = ColCnt + 2		' 한글 
		End If
		
		strChr = strChr & Mid(pStrLine, Cnt, 1)

		Select Case ColCnt
			Case 1		' 자료구분(1)
				strChr = ""
			Case 11		' 사업자등록번호(10)
				strChr = ""
			Case 18		' 전체매출처수(7)
				frm1.vspdData1.Row = iTotSumRowNo '3
				frm1.vspdData1.Col = C_BPCntSum '1 'jsk 20021216
				frm1.vspdData1.text = strChr
				strChr = ""
			Case 25		' 세금계산서매수(7)
				frm1.vspdData1.Row = iTotSumRowNo '3
				frm1.vspdData1.Col = C_PaperCntSum '2 'jsk 20021216
				frm1.vspdData1.text = strChr
				strChr = ""
			Case 40		' 공급가액(15)
				frm1.vspdData1.Row = iTotSumRowNo '3
				frm1.vspdData1.Col = C_NetAmtSum '3 'jsk 20021216
				'//frm1.vspdData1.text = strChr
				If (Right(strChr,1) <= "9") Then
					frm1.vspdData1.text = strChr	
				Else	
					If Right(strChr,1) = "}" Then
						LastAmt = "0"
					Else
						LastAmt = CHR(ASC(Right(strChr,1)) - 25)
					End if
					frm1.vspdData1.text = "-" & mid(strChr,1,len(strChr)-1) & LastAmt
				End If	
				strChr = ""
			Case 54		' 세액(14)
				frm1.vspdData1.Row = iTotSumRowNo '3
				frm1.vspdData1.Col = C_VatAmtSum '4 'jsk 20021216
				If (Right(strChr,1) <= "9") Then
					frm1.vspdData1.text = strChr	
				Else	
					If Right(strChr,1) = "}" Then
						LastAmt = "0"
					Else
						LastAmt = CHR(ASC(Right(strChr,1)) - 25)
					End if
					frm1.vspdData1.text = "-" & mid(strChr,1,len(strChr)-1) & LastAmt
				End If	
				strChr = ""
			Case 61		' 사업자별매출처수(7)
				frm1.vspdData1.Row = iRegNOPuRowNo '1
				frm1.vspdData1.Col =  C_BPCntSum '1 'jsk 20021216
				frm1.vspdData1.text = strChr
				strChr = ""
			Case 68		' 세금계산서매수(7)
				frm1.vspdData1.Row = iRegNOPuRowNo '1
				frm1.vspdData1.Col = C_PaperCntSum '2 'jsk 20021216
				frm1.vspdData1.text = strChr
				
				strChr = ""
			Case 83		' 공급가액(15)
				frm1.vspdData1.Row = iRegNOPuRowNo '1
				frm1.vspdData1.Col = C_NetAmtSum '3 'jsk 20021216
				If (Right(strChr,1) <= "9") Then
					frm1.vspdData1.text = strChr	
				Else	
					If Right(strChr,1) = "}" Then
						LastAmt = "0"
					Else
						LastAmt = CHR(ASC(Right(strChr,1)) - 25)
					End if
					frm1.vspdData1.text = "-" & mid(strChr,1,len(strChr)-1) & LastAmt
				End If	
				strChr = ""
			Case 97		' 세액(14)
				frm1.vspdData1.Row = iRegNOPuRowNo '1
				frm1.vspdData1.Col = C_VatAmtSum '4 'jsk 20021216
				If (Right(strChr,1) <= "9") Then
					frm1.vspdData1.text = strChr	
				Else	
					If Right(strChr,1) = "}" Then
						LastAmt = "0"
					Else
						LastAmt = CHR(ASC(Right(strChr,1)) - 25)
					End if
					frm1.vspdData1.text = "-" & mid(strChr,1,len(strChr)-1) & LastAmt
				End If	
				strChr = ""
			Case 104	' 주민별매출처수(7)
				frm1.vspdData1.Row = iPreRgstNoPuRowNo '2
				frm1.vspdData1.Col = C_BPCntSum '1 'jsk 20021216
				frm1.vspdData1.text = strChr
				strChr = ""
			Case 111	' 세금계산서매수(7)
				frm1.vspdData1.Row = iPreRgstNoPuRowNo '2
				frm1.vspdData1.Col = C_PaperCntSum '2 'jsk 20021216
				frm1.vspdData1.text = strChr
				strChr = ""
			Case 126	' 공급가액(15)
				frm1.vspdData1.Row = iPreRgstNoPuRowNo '2
				frm1.vspdData1.Col = C_NetAmtSum '3 'jsk 20021216
				'//frm1.vspdData1.text = strChr
				If (Right(strChr,1) <= "9") Then
					frm1.vspdData1.text = strChr	
				Else	
					If Right(strChr,1) = "}" Then
						LastAmt = "0"
					Else
						LastAmt = CHR(ASC(Right(strChr,1)) - 25)
					End if
					frm1.vspdData1.text = "-" & mid(strChr,1,len(strChr)-1) & LastAmt
				End If	
				strChr = ""
			Case 140	' 세액(14)
				frm1.vspdData1.Row = iPreRgstNoPuRowNo '2
				frm1.vspdData1.Col = C_VatAmtSum '4 'jsk 20021216
				If (Right(strChr,1) <= "9") Then
					frm1.vspdData1.text = strChr	
				Else	
					If Right(strChr,1) = "}" Then
						LastAmt = "0"
					Else
						LastAmt = CHR(ASC(Right(strChr,1)) - 25)
					End if
					frm1.vspdData1.text = "-" & mid(strChr,1,len(strChr)-1) & LastAmt
				End If	
				strChr = ""
			Case 170	' 공란(30)
				strChr = ""
		End Select
		
		If ColCnt >= 170 Then Exit For

	Next
    frm1.vspdData1.Col   = frm1.vspdData1.MaxCols
	frm1.vspdData1.value = pStrLineNo
End Function


Function subRceipt(Byval pStrLine,Byval pStrLineNo)
 
	' 매입처정보 
	' 자료구분(1), 사업자등록번호(10), 일련번호(4), 사업자등록번호(10), 상호(30), 업태(17), 종목(25), 매수(7), 
	' 공란수(2), 공급가액(14), 세액(13), 주류도매(1), 주류소매(1), 권번호(4), 제출처(3), 공란(28)

	Dim Cnt, ColCnt
	Dim strChr
	dim LastAmt	
	Dim str
	Dim strBPRgstNO,strPaperCnt, strBlankCnt, strNetAmt, strVatAmt, strCode, strBPNM, strIndTypeNM, strIndClassNM

	strChr = ""
	ColCnt = 0

	For Cnt = 1 To Len(pStrLine)
		If Asc(Mid(pStrLine, Cnt, 1)) >= 0 Then
			ColCnt = ColCnt + 1		' 영숫자 
		Else
			ColCnt = ColCnt + 2		' 한글 
		End If
		
		strChr = strChr & Mid(pStrLine, Cnt, 1)

		Select Case ColCnt
			Case 1		' 자료구분(1)
				strChr = ""
			Case 11		' 사업자등록번호(10)
				strChr = ""
			Case 15		' 일련번호(4)
				strChr = ""
			Case 25		' 사업자등록번호(10)
                strBPRgstNO = strChr
				strChr = ""
			Case 55		' 상호(30)
                strBPNM = strChr
				strChr = ""
			Case 72		' 업태(17)
                strIndTypeNM = strChr
				strChr = ""
			Case 97		' 종목(25)
                strIndClassNM = strChr
				strChr = ""
			Case 104	' 매수(7)
                strPaperCnt = strChr
				strChr = ""
			Case 106	' 공란수(2)
                strBlankCnt = strChr
				strChr = ""
			Case 120	' 공급가액(14)
				If (Right(strChr,1) <= "9") Then
                    strNetAmt = strChr
				Else	
					If Right(strChr,1) = "}" Then
						LastAmt = "0"
					Else
						LastAmt = CHR(ASC(Right(strChr,1)) - 25)
					End if
                    strNetAmt = "-" & mid(strChr,1,len(strChr)-1) & LastAmt

				End If	
				strChr = ""
			Case 133	' 세액(13)
				If (Right(strChr,1) <= "9") Then
                    strVatAmt = strChr
				Else	
					If Right(strChr,1) = "}" Then
						LastAmt = "0"
					Else
						LastAmt = CHR(ASC(Right(strChr,1)) - 25)
					End if
                    strVatAmt = "-" & mid(strChr,1,len(strChr)-1) & LastAmt
				End If	
				strChr = ""
			Case 134	' 주류도매(1)
				strChr = strChr & " / "
			Case 135	' 주류소매(1)
                strCode = strChr
				strChr = ""
			Case 139	' 권번호(4)
				strChr = ""
			Case 142	' 제출처(3)
				strChr = ""
			Case 170	' 공란(28)
				strChr = ""
		End Select
		
		If ColCnt >= 170 Then Exit For

	Next
    strTmpGrid = strTmpGrid & chr(11) & strBPRgstNO 
    strTmpGrid = strTmpGrid & chr(11) & strPaperCnt 
    strTmpGrid = strTmpGrid & chr(11) & strBlankCnt 
    strTmpGrid = strTmpGrid & chr(11) & strNetAmt 
    strTmpGrid = strTmpGrid & chr(11) & strVatAmt 
    strTmpGrid = strTmpGrid & chr(11) & strCode 
    strTmpGrid = strTmpGrid & chr(11) & strBPNM 
    strTmpGrid = strTmpGrid & chr(11) & strIndTypeNM 
    strTmpGrid = strTmpGrid & chr(11) & strIndClassNM 
    strTmpGrid = strTmpGrid & chr(11) & pStrLineNo & chr(11) & chr(12)
End Function
'#######################################################################################################
'  subRceiptSum(pStrLine): 매입처합계정보 
'  자료구분(1), 사업자등록번호(10), 매입처수(7), 계산서매수(7), 공급가액(15), 세액(14), 공란(116)
'####################################################################################################### 


Function subRceiptSum(pStrLine)
Dim Cnt, ColCnt
Dim strChr
Dim LastAmt
Dim iRegNOPuRowNo, iPreRgstNoPuRowNo, iTotSumRowNo
Call InitRowColVale("A",iRegNOPuRowNo, iPreRgstNoPuRowNo, iTotSumRowNo)

	strChr = ""
	ColCnt = 0

	For Cnt = 1 To Len(pStrLine)
		If Asc(Mid(pStrLine, Cnt, 1)) >= 0 Then
			ColCnt = ColCnt + 1		' 영숫자 
		Else
			ColCnt = ColCnt + 2		' 한글 
		End If
		
		strChr = strChr & Mid(pStrLine, Cnt, 1)
		Select Case ColCnt
			Case 1		' 자료구분(1)
				strChr = ""
			Case 11		' 사업자등록번호(10)
				strChr = ""
			Case 18		' 매입처수(7)
				frm1.vspdData1.Row = iTotSumRowNo '3
				frm1.vspdData1.Col = C_BPCntSum '1 'jsk 20021216
				frm1.vspdData1.text = strChr
				strChr = ""
			Case 25		' 계산서매수(7)
				frm1.vspdData1.Row = iTotSumRowNo '3
				frm1.vspdData1.Col = C_PaperCntSum '2 'jsk 20021216
				frm1.vspdData1.text = strChr
				strChr = ""
			Case 40		' 공급가액(15)
				frm1.vspdData1.Row =  iTotSumRowNo '3
				frm1.vspdData1.Col = C_NetAmtSum '3 'jsk 20021216
				If (Right(strChr,1) <= "9") Then
					frm1.vspdData1.text = strChr	
				Else	
					If Right(strChr,1) = "}" Then
						LastAmt = "0"
					Else
						LastAmt = CHR(ASC(Right(strChr,1)) - 25)
					End if
					frm1.vspdData1.text = "-" & mid(strChr,1,len(strChr)-1) & LastAmt
				End If	
				strChr = ""
			Case 54		' 세액(14)
				frm1.vspdData1.Row = iTotSumRowNo '3
				frm1.vspdData1.Col = C_VatAmtSum '4 'jsk 20021216
				If (Right(strChr,1) <= "9") Then
					frm1.vspdData1.text = strChr	
				Else	
					If Right(strChr,1) = "}" Then
						LastAmt = "0"
					Else
						LastAmt = CHR(ASC(Right(strChr,1)) - 25)
					End if
					frm1.vspdData1.text = "-" & mid(strChr,1,len(strChr)-1) & LastAmt
				End If	
				strChr = ""
			Case 170	' 공란(116)
				strChr = ""
		End Select
		
		If ColCnt >= 170 Then Exit For
	Next
End Function

'==========================================  2.1.1 subPayment2()  ======================================
'	Name : subPayment2()
'	Description : 매출처정보 (계산서)
'========================================================================================================= 

Function subPayment2(Byval pStrLine, Byval BGubunCnt)
Dim Cnt, ColCnt
Dim strChr
Dim signFlag
dim LastAmt	
	frm1.vspdData6.MaxRows = frm1.vspdData6.MaxRows + 1
	frm1.vspdData6.Row = frm1.vspdData6.MaxRows
	strChr = ""
	ColCnt = 0
	
	For Cnt = 1 To Len(pStrLine)
		If Asc(Mid(pStrLine, Cnt, 1)) >= 0 Then
			ColCnt = ColCnt + 1		' 영숫자 
		Else
			ColCnt = ColCnt + 2		' 한글 
		End If
		
		strChr = strChr & Mid(pStrLine, Cnt, 1)
		Select Case ColCnt
			Case 1		' 레코드구분(1)
				frm1.vspdData6.Col = 1 'C_DRecord4  
				frm1.vspdData6.text = strChr
				strChr = ""				
			Case 3		' 자료구분 
				strChr = ""
			Case 4		' 기구분(1)
				strChr = ""
			Case 5		' 신고구분(1)
				strChr = ""
			Case 8		' 세무서코드(3)
				strChr = ""
			Case 14		'일련번호(6)
				strChr = ""
			Case 24		' 제출의무자사업자등록번호(10)
				strChr = ""
			Case 34	' 사업자등록번호(10)
				frm1.vspdData6.Col = 2 'C_BPRgstNO4 
				frm1.vspdData6.text = strChr
				strChr = ""
			Case 74	' 법인명(상호)(40)
				frm1.vspdData6.Col = 3 'C_BPNM4 '
				frm1.vspdData6.text = strChr
				strChr = ""
			Case 79	' 계산서매수(5)
				frm1.vspdData6.Col = 4 'C_PaperCnt4 '
				frm1.vspdData6.text = strChr
				strChr = ""
			Case 80	' 음수표시(1)
				signFlag = strChr
				strChr = ""
			Case 94	' 금액(14)
				frm1.vspdData6.Col = 5 'C_NetAmt4 '
				If signFlag = "0" Then
					frm1.vspdData6.text = strChr	
				Else
					frm1.vspdData6.text = "-" & strChr	
				End If
				strChr = ""
				signFlag = ""
			Case 230	' 공란(136)
				strChr = ""
		End Select
		If ColCnt >= 230 Then Exit For
	Next
	frm1.vspdData6.Col = 6 'C_LoopCnt4 '
	frm1.vspdData6.text = BGubunCnt
	frm1.vspdData6.Col = 7 'C_LoopCnt4 '
	frm1.vspdData6.text = BGubunCnt
	
End Function


'==========================================  2.1.1 subPaymentSum2()  ======================================
'	Name : subPaymentSum2()
'	Description : 매출처합계정보 (계산서)
'========================================================================================================= 

Function subPaymentSum2(Byval pStrLine, BGubunCnt)
	Dim Cnt, ColCnt
	Dim strChr
	Dim LastAmt
	Dim signFlag
    '//Call InitSpreadSheet3()
	frm1.vspdData5.MaxRows = frm1.vspdData5.MaxRows + 1
	frm1.vspdData5.Row = frm1.vspdData5.MaxRows

	strChr = ""
	ColCnt = 0
	For Cnt = 1 To Len(pStrLine)
		If Asc(Mid(pStrLine, Cnt, 1)) >= 0 Then
			ColCnt = ColCnt + 1		' 영숫자 
		Else
			ColCnt = ColCnt + 2		' 한글 
		End If
		
		strChr = strChr & Mid(pStrLine, Cnt, 1)

		Select Case ColCnt
			Case 1		' 레코드구분(1)
				frm1.vspdData5.Col = 1 'C_CRecord5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 3		' 자료구분(2)
				strChr = ""
			Case 4		' 기구분(1)
				frm1.vspdData5.Col = 2 'C_Gigubun5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 5		' 신고구분(1)
				frm1.vspdData5.Col = 3 'C_SingoGubun5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 8		' 세무서(3)
				frm1.vspdData5.Col = 4 'C_TaxOffice5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 14		' 일련번호(6)
				strChr = ""
			Case 24		' 제출의무자 사업자등록번호(10)
				strChr = ""
			Case 28		'귀속년도(4)
				frm1.vspdData5.Col = 5 'C_ReturnYear5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 36		' 거래기간시작년월일(8)
				frm1.vspdData5.Col = 6 'C_StartDt5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 44		'거래기간종료년월일(8)
				frm1.vspdData5.Col = 7 'C_EndDt5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 52	' 작성일자(8)
				frm1.vspdData5.Col = 8 'C_ReportDt5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			
			Case 58	' 매출처수(6) - 합계 
				frm1.vspdData5.Col = 15 'C_HBPCntSum5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 64	'계산서매수(6) - 합계 
				frm1.vspdData5.Col = 16 'C_HPaperCntSum5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 65	'음수표시(1) - 합계 
				signFlag = strChr
				strChr = ""
			Case 79	'금액(14) - 합계 
				frm1.vspdData5.Col = 17 'C_HNetAmtSum5 '
				If signFlag = "0" Then
					frm1.vspdData5.text = strChr	
				Else
					frm1.vspdData5.text = "-" & strChr	
				End If
				strChr = ""
				signFlag = ""
			Case 85	' 매출처수(6)-사업자등록번호발행분 
				frm1.vspdData5.Col = 9 'C_BBPCntSum5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 91	'계산서매수(6) - 사업자등록번호발행분 
				frm1.vspdData5.Col = 10 'C_BPaperCntSum5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 92	' 음수표시 - 사업자등록번호발행분 
				signFlag = strChr
				strChr = ""
			Case 106	'금액 - 사업자등록번호발행분 
				frm1.vspdData5.Col = 11 'C_BNetAmtSum5 '
				If signFlag = "0" Then
					frm1.vspdData5.text = strChr	
				Else
					frm1.vspdData5.text = "-" & strChr	
				End If
				strChr = ""
				signFlag = ""
			Case 112	'매출처수 - 주민등록번호 발행분 
				frm1.vspdData5.Col = 12 'C_RBPCntSum5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 118	' 계산서매수 - 주민등록번호발행분 
				frm1.vspdData5.Col = 13 'C_RPaperCntSum5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 119	' 음수표시 - 주민등록번호발행분 
				signFlag = strChr
				strChr = ""
			Case 133	' 금액 - 주민등록번호발행분 
				frm1.vspdData5.Col = 14 'C_RNetAmtSum5 '
				If signFlag = "0" Then
					frm1.vspdData5.text = strChr	
				Else
					frm1.vspdData5.text = "-" & strChr	
				End If
				strChr = ""
				signFlag = ""
				
			Case 230	' 공란(97)
				strChr = ""
		End Select
		
		If ColCnt >= 230 Then Exit For
	Next
	frm1.vspdData5.Col = 18 'C_LoopCnt5 '
	frm1.vspdData5.text = BGubunCnt
	
End Function
'==========================================  2.1.1 subRceipt2()  ======================================
'	Name : subRceipt2()
'	Description : 매입처정보 (계산서)
'========================================================================================================= 

Function subRceipt2(Byval pStrLine, Byval BGubunCnt)
	Dim Cnt, ColCnt
	Dim strChr
	Dim signFlag
	Dim LastAmt	
	frm1.vspdData6.MaxRows = frm1.vspdData6.MaxRows + 1
	frm1.vspdData6.Row = frm1.vspdData6.MaxRows
	strChr = ""
	ColCnt = 0
	
	For Cnt = 1 To Len(pStrLine)
		If Asc(Mid(pStrLine, Cnt, 1)) >= 0 Then
			ColCnt = ColCnt + 1		' 영숫자 
		Else
			ColCnt = ColCnt + 2		' 한글 
		End If
		
		strChr = strChr & Mid(pStrLine, Cnt, 1)
		Select Case ColCnt
			Case 1		' 레코드구분(1)
				frm1.vspdData6.Col = 1 'C_DRecord4 '
				frm1.vspdData6.text = strChr
				strChr = ""
				
			Case 3		' 자료구분 
				strChr = ""
			Case 4		' 기구분(1)
				strChr = ""
			Case 5		' 신고구분(1)
				strChr = ""
			Case 8		' 세무서코드(3)
				strChr = ""
			Case 14		'일련번호(6)
				strChr = ""
			Case 24		' 제출의무자사업자등록번호(10)
				strChr = ""
			Case 34	' 사업자등록번호(10)
				frm1.vspdData6.Col = 2 'C_BPRgstNO4 '
				frm1.vspdData6.text = strChr
				strChr = ""
			Case 74	' 법인명(상호)(40)
				frm1.vspdData6.Col = 3 'C_BPNM4 '
				frm1.vspdData6.text = strChr
				strChr = ""
			Case 79	' 계산서매수(5)
				frm1.vspdData6.Col = 4 'C_PaperCnt4 '
				frm1.vspdData6.text = strChr
				strChr = ""
			Case 80	' 음수표시(1)
				signFlag = strChr
				strChr = ""
			Case 94	' 금액(14)
				frm1.vspdData6.Col = 5 'C_NetAmt4 '
				If signFlag = "0" Then
					frm1.vspdData6.text = strChr	
				Else
					frm1.vspdData6.text = "-" & strChr	
				End If
				strChr = ""
				signFlag = ""
			Case 230	' 공란(136)
				strChr = ""
		End Select
		If ColCnt >= 230 Then Exit For
	Next
	frm1.vspdData6.Col = 6 'C_LoopCnt4 '
	frm1.vspdData6.text = BGubunCnt
	frm1.vspdData6.Col = 7 'C_LoopCnt4 '
	frm1.vspdData6.text = BGubunCnt
	
End Function

'==========================================  2.1.1 subRceiptSum2()  ======================================
'	Name : subRceiptSum2()
'	Description : 매입처합계정보 (계산서)
'========================================================================================================= 
Function subRceiptSum2(pStrLine, BGubunCnt)
Dim Cnt, ColCnt
Dim strChr
Dim LastAmt
Dim signFlag
  '//  Call InitSpreadSheet3()
	frm1.vspdData5.MaxRows = frm1.vspdData5.MaxRows + 1
	frm1.vspdData5.Row = frm1.vspdData5.MaxRows
	

	strChr = ""
	ColCnt = 0
	For Cnt = 1 To Len(pStrLine)
		If Asc(Mid(pStrLine, Cnt, 1)) >= 0 Then
			ColCnt = ColCnt + 1		' 영숫자 
		Else
			ColCnt = ColCnt + 2		' 한글 
		End If
		
		strChr = strChr & Mid(pStrLine, Cnt, 1)

		Select Case ColCnt
			Case 1		' 레코드구분(1)
				frm1.vspdData5.Col = 1 'C_CRecord5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 3		' 자료구분(2)
				strChr = ""
			Case 4		' 기구분(1)
				frm1.vspdData5.Col = 2 'C_Gigubun5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 5		' 신고구분(1)
				frm1.vspdData5.Col = 3 'C_SingoGubun5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 8		' 세무서(3)
				frm1.vspdData5.Col = 4 'C_TaxOffice5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 14		' 일련번호(6)
				strChr = ""
			Case 24		' 제출의무자 사업자등록번호(10)
				strChr = ""
			Case 28		'귀속년도(4)
				frm1.vspdData5.Col = 5 'C_ReturnYear5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 36		' 거래기간시작년월일(8)
				frm1.vspdData5.Col = 6 'C_StartDt5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 44		'거래기간종료년월일(8)
				frm1.vspdData5.Col = 7 'C_EndDt5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 52	' 작성일자(8)
				frm1.vspdData5.Col = 8 'C_ReportDt5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			
			Case 58	' 매출처수(6) - 합계 
				frm1.vspdData5.Col = 15 'C_HBPCntSum5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 64	'계산서매수(6) - 합계 
				frm1.vspdData5.Col = 16 'C_HPaperCntSum5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 65	'음수표시(1) - 합계 
				signFlag = strChr
				strChr = ""
			Case 79	'금액(14) - 합계 
				frm1.vspdData5.Col = 17 'C_HNetAmtSum5 '
				If signFlag = "0" Then
					frm1.vspdData5.text = strChr	
				Else
					frm1.vspdData5.text = "-" & strChr	
				End If
				strChr = ""
				signFlag = ""
			Case 230	' 공란(97)
				strChr = ""
		End Select
		
		If ColCnt >= 230 Then Exit For
	Next
	frm1.vspdData5.Col = 18 'C_LoopCnt5 '
	frm1.vspdData5.text = BGubunCnt
End Function


Function subCompany3(Byval pStrLine)
 
	' 자사정보 
	' 자료구분(1), 귀속년월(6), 신고구분(1), 사업자등록번호(10), 상호(30), 성명(15), 사업장소재지(45), 
	' 업태(17), 종목(25), 거래기간(8), 거래기간(8), 작성일자(8), 공란(6)

	Dim Cnt, ColCnt
	Dim strChr

	strChr = ""
	ColCnt = 0
	For Cnt = 1 To Len(pStrLine)
		If Asc(Mid(pStrLine, Cnt, 1)) >= 0 Then
			ColCnt = ColCnt + 1		' 영숫자 
		Else
			ColCnt = ColCnt + 2		' 한글 
		End If
		
		strChr = strChr & Mid(pStrLine, Cnt, 1)
		Select Case ColCnt
			Case 1		' 자료구분(1)
				strChr = ""
			Case 7		' 귀속년월(6)
				frm1.txtYearMonth3.value = strChr
				strChr = ""
			Case 8		' 신고구분(1)
				frm1.txtSingo3.value = strChr
				strChr = ""	
				
			Case 18		' 사업자등록번호(10)
				frm1.txtRegNo3.value = strChr
				strChr = ""
			Case 48		' 상호(30)
				frm1.txtBizAreaNm3.value = strChr
				strChr = ""
			Case 63		' 성명(15)
				frm1.txtRepreNm3.value = strChr
				strChr = ""
			Case 108	' 사업장소재지(45)
				frm1.txtAddr3.value = strChr
				strChr = ""
			Case 125	' 업태(17)
				frm1.txtIndType3.value = strChr
				strChr = ""
			Case 150	' 종목(25)
				frm1.txtIndClass3.value = strChr
				strChr = ""
			Case 158	' 거래기간(8)
				frm1.txtStart3.value = strChr
				strChr = ""
			Case 166	' 거래기간(8)
				frm1.txtEnd3.value = strChr
				strChr = ""
			Case 174	' 작성일자(8)
				frm1.txtReport3.value = strChr
				strChr = ""
			Case 180	' 공란(6)
				strChr = ""
		End Select		
		If ColCnt >= 180 Then Exit For
	Next
End Function

 '==========================================  2.1.1 subExportList()  ======================================
'	Name : subExportList()
'	Description : 수출실적정보 
'========================================================================================================= 

Function subExportList(Byval pStrLine,Byval pStrLineNo)
	'자료구분(1), 귀속년월(6), 신고구분(1), 사업자등록번호(10), 일련번호(7), 수출신고번호(15), 선적(기)일자(8)
	'통화코드(3), 환율(5), 환율(4), 외화금액(13), 외화금액(2), 원화금액(15), 공란(90)
	Dim Cnt, ColCnt
	Dim strChr
	Dim LastAmt
	Dim strExportNo7, strFnDt7, strDocCur7, strXchRate7, strDocAmt7, strLocAmt7
	strChr = ""
	ColCnt = 0
	
	For Cnt = 1 To Len(pStrLine)
		If Asc(Mid(pStrLine, Cnt, 1)) >= 0 Then
			ColCnt = ColCnt + 1		' 영숫자 
		Else
			ColCnt = ColCnt + 2		' 한글 
		End If
		
		strChr = strChr & Mid(pStrLine, Cnt, 1)

		Select Case ColCnt
			Case 1		' 자료구분(1)
				strChr = ""
			Case 7		' 귀속년월(6)
				strChr = ""
			Case 8		' 신고구분(1)
				strChr = ""
			Case 18		' 사업자등록번호(10)
				strChr = ""
			Case 25		' 일련번호(7)
				strChr = ""
			Case 40		' 수출신고번호(15)
				strExportNo7 = strChr 
				strChr = ""
			Case 48		' 선적(기)일자(8)
				strFnDt7 = strChr 
				strChr = ""
			Case 51		' 통화코드(3)
				strDocCur7 = strChr 
				strChr = ""
			Case 56		' 환율(5)
				strChr = strChr & parent.gComNumDec '"."
			Case 60		' 환율(4)
				strXchRate7 = strChr 
				strChr = ""
			Case 73		'외화금액(13)
				If (Right(strChr,1) <= "9") Then
					strDocAmt7 = strChr 
				Else	
					If Right(strChr,1) = "}" Then
						LastAmt = "0"
					Else
						LastAmt = CHR(ASC(Right(strChr,1)) - 25)
					End if
					strFnDt7 = "-" & mid(strChr,1,len(strChr)-1) & LastAmt
				End If	
				strChr = strChr & parent.gComNumDec '"."
			Case 75		'외화금액(2)
				strDocAmt7 = strChr 
				strChr = ""	
			Case 90	' 원화금액(15)
				If (Right(strChr,1) <= "9") Then
					strLocAmt7 = strChr 
				Else	
					If Right(strChr,1) = "}" Then
						LastAmt = "0"
					Else
						LastAmt = CHR(ASC(Right(strChr,1)) - 25)
					End if
					strLocAmt7 = "-" & mid(strChr,1,len(strChr)-1) & LastAmt
				End If	
				strChr = ""
			Case 180	' 공란(90)
				strChr = ""
		End Select
		
		If ColCnt >= 180 Then Exit For
	Next	
    strTmpGrid7 = strTmpGrid7 & chr(11) & strExportNo7 
    strTmpGrid7 = strTmpGrid7 & chr(11) & strFnDt7 
    strTmpGrid7 = strTmpGrid7 & chr(11) & strDocCur7 
    strTmpGrid7 = strTmpGrid7 & chr(11) & strXchRate7 
    strTmpGrid7 = strTmpGrid7 & chr(11) & strDocAmt7 
    strTmpGrid7 = strTmpGrid7 & chr(11) & strLocAmt7 
    strTmpGrid7 = strTmpGrid7 & chr(11) & pStrLineNo & chr(11) & chr(12)
End Function



Function subExportSum(Byval pStrLine)

	'자료구분(1), 귀속년월(1), 신고구분(1), 사업자등록번호(10), 전체건수(7), 외화금액(전체)(13), 외화금액(전체)(2)
	'원화금액(전체)(15), 원화금액(전체)(15), 수출하는재화건수(7), 외화금액(수출)(13), 외화금액(수출)(2), 원화금액(수출)(15)
	'영세울건수(7), 외화금액(영세율)(13), 외화금액(영세율)(2), 원화금액(영세율)(15), 공란(51)
	Dim Cnt, ColCnt
	Dim strChr
	Dim LastAmt

	Dim iExport, ilgEtcTax, iTotSumRowNo

	Call InitRowColVale("C",iExport, ilgEtcTax, iTotSumRowNo)
	
	strChr = ""
	ColCnt = 0
	For Cnt = 1 To Len(pStrLine)
		If Asc(Mid(pStrLine, Cnt, 1)) >= 0 Then
			ColCnt = ColCnt + 1		' 영숫자 
		Else
			ColCnt = ColCnt + 2		' 한글 
		End If
		
		strChr = strChr & Mid(pStrLine, Cnt, 1)

		Select Case ColCnt
			Case 1		' 자료구분(1)
				strChr = ""
			Case 7		' 귀속년월(1)
				strChr = ""
			Case 8		' 신고구분(1)
				strChr = ""
			Case 18		' 사업자등록번호(10)
				strChr = ""
			Case 25		' 전체건수(7)
				frm1.vspdData8.Row = iTotSumRowNo '3
				frm1.vspdData8.Col = C_CntSum8
				frm1.vspdData8.text = strChr
				strChr = ""
			Case 38		' 외화금액(전체)(13)
				frm1.vspdData8.Row = iTotSumRowNo '3
				frm1.vspdData8.Col = C_DocSum8
				If (Right(strChr,1) <= "9") Then
					frm1.vspdData8.text = strChr	
				Else	
					If Right(strChr,1) = "}" Then
						LastAmt = "0"
					Else
						LastAmt = CHR(ASC(Right(strChr,1)) - 25)
					End if
					frm1.vspdData8.text = "-" & mid(strChr,1,len(strChr)-1) & LastAmt
				End If	
				strChr = strChr & parent.gComNumDec '"."
			Case 40		' 외화금액(전체)(2)
				frm1.vspdData8.Row = iTotSumRowNo '3
				frm1.vspdData8.Col = C_DocSum8
				frm1.vspdData8.text = strChr
				strChr = ""
			Case 55		' 원화금액(전체)(15)
				frm1.vspdData8.Row = iTotSumRowNo '3
				frm1.vspdData8.Col = C_LocSum8
				If (Right(strChr,1) <= "9") Then
					frm1.vspdData8.text = strChr	
				Else	
					If Right(strChr,1) = "}" Then
						LastAmt = "0"
					Else
						LastAmt = CHR(ASC(Right(strChr,1)) - 25)
					End if
					frm1.vspdData8.text = "-" & mid(strChr,1,len(strChr)-1) & LastAmt
				End If	
				strChr = ""
			Case 62		' 수출하는재화건수(7)
				frm1.vspdData8.Row = iExport '1
				frm1.vspdData8.Col = C_CntSum8
				frm1.vspdData8.text = strChr
				strChr = ""
			Case 75		' 외화금액(수출)(13)
				frm1.vspdData8.Row = iExport '1
				frm1.vspdData8.Col = C_DocSum8
				If (Right(strChr,1) <= "9") Then
					frm1.vspdData8.text = strChr	
				Else	
					If Right(strChr,1) = "}" Then
						LastAmt = "0"
					Else
						LastAmt = CHR(ASC(Right(strChr,1)) - 25)
					End if
					frm1.vspdData8.text = "-" & mid(strChr,1,len(strChr)-1) & LastAmt
				End If	
				strChr = strChr & parent.gComNumDec '"."
			Case 77		' 외화금액(수출)(2)
				frm1.vspdData8.Row = iExport '1
				frm1.vspdData8.Col = C_DocSum8
				frm1.vspdData8.text = strChr
				strChr = ""
			Case 92		' 원화금액(수출)(15)
				frm1.vspdData8.Row = iExport '1
				frm1.vspdData8.Col = C_LocSum8
				If (Right(strChr,1) <= "9") Then
					frm1.vspdData8.text = strChr	
				Else	
					If Right(strChr,1) = "}" Then
						LastAmt = "0"
					Else
						LastAmt = CHR(ASC(Right(strChr,1)) - 25)
					End if
					frm1.vspdData8.text = "-" & mid(strChr,1,len(strChr)-1) & LastAmt
				End If	
				strChr = ""
			Case 99	' 영세울건수(7)
				frm1.vspdData8.Row = ilgEtcTax '2
				frm1.vspdData8.Col = C_CntSum8
				frm1.vspdData8.text = strChr
				strChr = ""
			Case 112	' 외화금액(영세율)(13)
				frm1.vspdData8.Row = ilgEtcTax '2
				frm1.vspdData8.Col = C_DocSum8
				If (Right(strChr,1) <= "9") Then
					frm1.vspdData8.text = strChr	
				Else	
					If Right(strChr,1) = "}" Then
						LastAmt = "0"
					Else
						LastAmt = CHR(ASC(Right(strChr,1)) - 25)
					End if
					frm1.vspdData8.text = "-" & mid(strChr,1,len(strChr)-1) & LastAmt
				End If	
				strChr = strChr & parent.gComNumDec '"."
			Case 114	' 외화금액(영세율)(2)
				frm1.vspdData8.Row = ilgEtcTax '2
				frm1.vspdData8.Col = C_DocSum8
				frm1.vspdData8.text =  strChr
				strChr = ""
			Case 129	' 원화금액(영세율)(15)
				frm1.vspdData8.Row = ilgEtcTax '2
				frm1.vspdData8.Col = C_LocSum8
				If (Right(strChr,1) <= "9") Then
					frm1.vspdData8.text = strChr	
				Else	
					If Right(strChr,1) = "}" Then
						LastAmt = "0"
					Else
						LastAmt = CHR(ASC(Right(strChr,1)) - 25)
					End if
					frm1.vspdData8.text = "-" & mid(strChr,1,len(strChr)-1) & LastAmt
				End If	
				strChr = ""
			Case 180	' 공란(51)
				strChr = ""
		End Select
		
		If ColCnt >= 180 Then Exit For

	Next

End Function


 '******************************  3.2.1 Object Tag 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 


'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub
'========================================================================================================
'   Event Name : vspdData1_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub
'========================================================================================================
'   Event Name : vspdData2_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub
'========================================================================================================
'   Event Name : vspdData3_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData3_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData3
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub
'========================================================================================================
'   Event Name : vspdData4_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData4_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData4
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub
'========================================================================================================
'   Event Name : vspdData7_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData7_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData7
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub





'========================================================================================================
'   Event Name : vspdData8_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData8_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData8
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("0000111111")    
     gMouseClickStatus = "SPC"	'Split 상태코드 
    
    Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows <= 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
    
End Sub

'==========================================================================================
'   Event Name : vspdData1_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("0000111111")    
    gMouseClickStatus = "SP2C"	'Split 상태코드 
    
    Set gActiveSpdSheet = frm1.vspdData1

    If frm1.vspdData1.MaxRows <= 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
End Sub

Sub vspdData2_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("0000111111")    
     gMouseClickStatus = "SP3C"	'Split 상태코드 
    
    Set gActiveSpdSheet = frm1.vspdData2

    If frm1.vspdData2.MaxRows <= 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData2
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
    
End Sub

Sub vspdData3_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("0000111111")    
     gMouseClickStatus = "SP4C"	'Split 상태코드 
    
    Set gActiveSpdSheet = frm1.vspdData3

    If frm1.vspdData3.MaxRows <= 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
End Sub

Sub vspdData4_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("0000111111")    
     gMouseClickStatus = "SP5C"	'Split 상태코드 
    
    Set gActiveSpdSheet = frm1.vspdData4

    If frm1.vspdData4.MaxRows <= 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData4
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
    
End Sub
Sub vspdData7_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("0000111111")    
     gMouseClickStatus = "SP6C"	'Split 상태코드 
    
    Set gActiveSpdSheet = frm1.vspdData7

    If frm1.vspdData7.MaxRows <= 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData7
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
    
End Sub


Sub vspdData8_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("0000111111")    
    gMouseClickStatus = "SP7C"	'Split 상태코드 
    
    Set gActiveSpdSheet = frm1.vspdData8

    If frm1.vspdData8.MaxRows <= 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
End Sub

Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then
    End If	
End Sub
Sub vspdData1_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then
    End If	
End Sub
Sub vspdData2_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then
    End If	
End Sub
Sub vspdData3_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then
    End If	
End Sub
Sub vspdData4_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then
    End If	
End Sub
Sub vspdData7_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then
    End If	
End Sub
Sub vspdData8_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then
    End If	
End Sub



Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub
Sub vspdData1_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("B")
End Sub
Sub vspdData2_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("C")
End Sub
Sub vspdData3_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData3
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("D")
End Sub
Sub vspdData4_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData4
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("E")
End Sub
Sub vspdData7_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData7
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("F")
End Sub
Sub vspdData8_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData8
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("G")
End Sub

'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================

Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

Sub vspdData1_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub

Sub vspdData2_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP3C" Then
		gMouseClickStatus = "SP3CR"
	End If
End Sub

Sub vspdData3_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP4C" Then
		gMouseClickStatus = "SP4CR"
	End If
End Sub

Sub vspdData4_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP5C" Then
		gMouseClickStatus = "SP5CR"
	End If
End Sub


Sub vspdData7_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP6C" Then
		gMouseClickStatus = "SP6CR"
	End If
End Sub


Sub vspdData8_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP7C" Then
		gMouseClickStatus = "SP7CR"
	End If
End Sub

Sub vspdData2_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	Dim i
	Dim RowList
	Dim intRetCD
    If Row <> NewRow And NewRow > 0 Then
        ggoSpread.Source = frm1.vspdData2
		If CopyFromData(NewRow) = False Then Exit Sub 		
	End If

End Sub
'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 
    Dim RetFlag
    If gSelframeFlg = TAB1 Then 

        ggoSpread.Source = frm1.vspdData
        ggoSpread.ClearSpreadData

		If Trim(frm1.txtFileName.value) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtFileName.Alt, "X") 	
			Exit Function
		End If
		If Trim(frm1.cboIOFlag.value) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.cboIOFlag.Alt, "X") 	
			Exit Function
		End If

	ElseIf gSelframeFlg = TAB2 Then 
        ggoSpread.Source = frm1.vspdData2
        ggoSpread.ClearSpreadData

        ggoSpread.Source = frm1.vspdData4
        ggoSpread.ClearSpreadData

        ggoSpread.Source = frm1.vspdData5
        ggoSpread.ClearSpreadData

        ggoSpread.Source = frm1.vspdData6
        ggoSpread.ClearSpreadData

		If Trim(frm1.txtFileName2.value) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtFileName2.Alt, "X") 	
			Exit Function
		End If
		If Trim(frm1.cboIOFlag2.value) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.cboIOFlag2.Alt, "X") 	
			Exit Function
		End If
	ElseIf gSelframeFlg = TAB3 Then
        ggoSpread.Source = frm1.vspdData7
        ggoSpread.ClearSpreadData
		If Trim(frm1.txtFileName3.value) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtFileName3.Alt, "X") 	
			Exit Function
		End If
    End If
	
    Call DbQuery
End Function


'========================================================================================
Function FncNew() 
End Function


'========================================================================================
Function FncDelete() 
End Function


'========================================================================================
Function FncSave() 
End Function


'========================================================================================
Function FncCopy() 
End Function


'========================================================================================
Function FncCancel() 
End Function


'========================================================================================
Function FncInsertRow() 
End Function


'========================================================================================
Function FncDeleteRow() 
End Function


'========================================================================================
Function FncPrev() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function


'========================================================================================
Function FncNext() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	Call parent.FncPrint()
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)												'☜: 화면 유형 
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                         '☜:화면 유형, Tab 유무 
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

	Dim indx

	on Error Resume Next
	Err.Clear 

	ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Select Case Trim(UCase(gActiveSpdSheet.Name))
		Case "VSPDDATA"
            Call InitSpreadSheet()      
            Call ggoSpread.ReOrderingSpreadData()
			
		Case "VSPDDATA1"
            Call InitSpreadSheet1()      
            Call ggoSpread.ReOrderingSpreadData()
			
		Case "VSPDDATA2"
            Call InitSpreadSheet2()      
            Call ggoSpread.ReOrderingSpreadData()
			
		Case "VSPDDATA3"
            Call InitSpreadSheet3()      
            Call ggoSpread.ReOrderingSpreadData()
			
		Case "VSPDDATA4"
            Call InitSpreadSheet4()      
            Call ggoSpread.ReOrderingSpreadData()
			
		Case "VSPDDATA7"
            Call InitSpreadSheet5()      
            Call ggoSpread.ReOrderingSpreadData()
			Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData7,-1 , -1 ,C_DocCur7 ,C_DocAmt7 ,   "A" ,"Q","X","X")

			
		Case "VSPDDATA8"
            Call InitSpreadSheet6()      
            Call ggoSpread.ReOrderingSpreadData()
			
	End Select

End Sub


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
Dim IntRetCD
	
	FncExit = False
	
	If lgBlnStartFlag = True Then
		' 변경된 내용이 있는지 확인한다.
		If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")			'⊙: "Will you destory previous data"
	
			If IntRetCD = vbNo Then
				Exit Function
			End If
		End If
    End If
    
    FncExit = True
    
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Dim strVal,strPath        
	Call LayerShowHide(1)    
	Err.Clear                                                               '☜: Protect system from crashing      
	DbQuery = False                                                         '⊙: Processing is NG
    If gSelframeFlg = TAB1 Then 
		frm1.hFileName.value = lgFilePath
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'☜: 비지니스 처리 ASP의 상태 
		strVal = strVal & "&txtFileName=" & Trim(frm1.txtFileName.value)
		strVal = strVal & "&cboFlag="     & Trim(frm1.cboIOFlag.value)
	ElseIf gSelframeFlg = TAB2 Then
		frm1.hFileName.value = lgFilePath2
		strVal = BIZ_PGM_ID2 & "?txtMode=" & parent.UID_M0001							'☜: 비지니스 처리 ASP의 상태 
		strVal = strVal & "&txtFileName=" & Trim(frm1.txtFileName2.value)
		strVal = strVal & "&cboFlag="     & Trim(frm1.cboIOFlag2.value)
	ElseIf gSelframeFlg = TAB3 Then
		frm1.hFileName.value = lgFilePath3
		strVal = BIZ_PGM_ID3 & "?txtMode=" & parent.UID_M0001							'☜: 비지니스 처리 ASP의 상태 
		strVal = strVal & "&txtFileName=" & Trim(frm1.txtFileName3.value)
	End If	
	strVal = strVal & "&hFileName="   & Trim(frm1.hFileName.value)
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동	
	DbQuery = True                                                          '⊙: Processing is NG
End Function


'========================================================================================
' Function Name : CopyFromData
' Function Desc : This function is data query and display
'========================================================================================
Function CopyFromData(Row) 

	Dim BrecordRow
	Dim iRow, iCol
	Dim iRegNOPuRowNo, iPreRgstNoPuRowNo, iTotSumRowNo
	Call InitRowColVale("B", iRegNOPuRowNo, iPreRgstNoPuRowNo, iTotSumRowNo)
	Dim strDRecord4, strBPRgstNO4, strBPNM4, strPaperCnt4, strNetAmt4, strLoopCnt4
	Err.Clear                                                               '☜: Protect system from crashing
	CopyFromData = False                                                         '⊙: Processing is NG

	With frm1
		.vspdData4.MaxRows = 0
		.vspdData2.Row = Row
		.vspdData2.Col = C_LoopCnt2
		BrecordRow = .vspdData2.text

		'/// vspdData5 ====> vspdData3
		For iRow=1 To .vspdData5.maxRows
			.vspdData5.Row = iRow
			.vspdData5.Col = 18 'C_LoopCnt5 '
			If BrecordRow = .vspdData5.text Then
				'//vspdData3 setting (선별후 카피)
				.vspdData5.Col = 2 'C_Gigubun5 '
				.txtGiGubun3.value = .vspdData5.Text

				.vspdData5.Col = 3 'C_SingoGubun5 '
				.txtSingoGubun3.value = .vspdData5.Text

				.vspdData5.Col = 4 'C_TaxOffice5 '
				.txtTaxOffice3.value = .vspdData5.Text

				.vspdData5.Col = 5 'C_ReturnYear5 '
				.txtReturnYear3.value = .vspdData5.Text

				.vspdData5.Col = 6 'C_StartDt5 '
				.txtStartDt3.value = .vspdData5.Text


				.vspdData5.Col = 7 'C_EndDt5 '
				.txtEndDt3.value = .vspdData5.Text

				.vspdData5.Col = 8 'C_ReportDt5 '
				.txtReportDt3.value = .vspdData5.Text

				.vspdData5.Col = 1 'C_CRecord5 '
				.vspdData3.Row = iTotSumRowNo '3
				.vspdData3.Col = C_CRecord3
				.vspdData3.Text = .vspdData5.Text

				'/// 사업자등록번호 발행분 
				.vspdData5.Col = 9 'C_BBPCntSum5 '
				.vspdData3.Row = iRegNOPuRowNo '1
				.vspdData3.Col = C_BPCntSum3
				.vspdData3.Text = .vspdData5.Text

				.vspdData5.Col = 10 'C_BPaperCntSum5 '
				.vspdData3.Row = iRegNOPuRowNo '1
				.vspdData3.Col = C_PaperCntSum3
				.vspdData3.Text = .vspdData5.Text

				.vspdData5.Col = 11 'C_BNetAmtSum5 '
				.vspdData3.Row = iRegNOPuRowNo '1
				.vspdData3.Col = C_NetAmtSum3
				.vspdData3.Text = .vspdData5.Text

				'//주민등록번호 발행분 
				.vspdData5.Col = 12 'C_RBPCntSum5 '
				.vspdData3.Row = iPreRgstNoPuRowNo '2
				.vspdData3.Col = C_BPCntSum3
				.vspdData3.Text = .vspdData5.Text

				.vspdData5.Col = 13 'C_RPaperCntSum5 '
				.vspdData3.Row = iPreRgstNoPuRowNo '2
				.vspdData3.Col = C_PaperCntSum3
				.vspdData3.Text = .vspdData5.Text

				.vspdData5.Col = 14 'C_RNetAmtSum5 '
				.vspdData3.Row = iPreRgstNoPuRowNo '2
				.vspdData3.Col = C_NetAmtSum3
				.vspdData3.Text = .vspdData5.Text

				'//합계 
				.vspdData5.Col = 15 'C_HBPCntSum5 '
				.vspdData3.Row = iTotSumRowNo '3
				.vspdData3.Col = C_BPCntSum3
				.vspdData3.Text = .vspdData5.Text

				.vspdData5.Col = 16 'C_HPaperCntSum5 '
				.vspdData3.Row = iTotSumRowNo '3
				.vspdData3.Col = C_PaperCntSum3
				.vspdData3.Text = .vspdData5.Text

				.vspdData5.Col = 17 'C_HNetAmtSum5 '
				.vspdData3.Row = iTotSumRowNo '3
				.vspdData3.Col = C_NetAmtSum3
				.vspdData3.Text = .vspdData5.Text

				'//매출인 경우만 세팅해줌 (사업자등록번호발행분의 매출처수가 존재할경우만 세팅)////////////////////
				.vspdData3.Row = iRegNOPuRowNo '1
				.vspdData3.Col = C_BPCntSum3
				If .vspdData3.Text <> "" Then
					.vspdData5.Col = 1 'C_CRecord5 '
					.vspdData3.Row = iRegNOPuRowNo '1
					.vspdData3.Col = C_CRecord3
					.vspdData3.Text = .vspdData5.Text

					.vspdData5.Col = 1 'C_CRecord5 '
					.vspdData3.Row = iPreRgstNoPuRowNo '2
					.vspdData3.Col = C_CRecord3
					.vspdData3.Text = .vspdData5.Text
				Else
					.vspdData3.Row = iRegNOPuRowNo '1
					.vspdData3.Col = C_CRecord3
					.vspdData3.Text = ""

					.vspdData3.Row = iPreRgstNoPuRowNo '2
					.vspdData3.Col = C_CRecord3
					.vspdData3.Text = ""
				End If	
			End If 
		Next

		'/// vspdData6 ====> vspdData4
		For iRow= 1 To .vspdData6.maxRows
			.vspdData6.Row = iRow
			.vspdData6.Col = C_LoopCnt4
			If BrecordRow = .vspdData6.text Then
				'//vspdData4 setting(무조건 카피)

				.vspdData6.Col = 1

				.vspdData4.Col = C_DRecord4
				strDRecord4 = .vspdData6.Text

				.vspdData6.Col = 2
				strBPRgstNO4 = .vspdData6.Text

				.vspdData6.Col = 3
				strBPNM4 = .vspdData6.Text

				.vspdData6.Col = 4
				strPaperCnt4 = .vspdData6.Text

				.vspdData6.Col = 5
				strNetAmt4 = .vspdData6.Text

				.vspdData6.Col = 6
				strLoopCnt4 = .vspdData6.Text

				strTmpGrid4 = strTmpGrid4 & chr(11) & strDRecord4 
				strTmpGrid4 = strTmpGrid4 & chr(11) & strBPRgstNO4 
				strTmpGrid4 = strTmpGrid4 & chr(11) & strBPNM4 
				strTmpGrid4 = strTmpGrid4 & chr(11) & strPaperCnt4 
				strTmpGrid4 = strTmpGrid4 & chr(11) & strNetAmt4 
				strTmpGrid4 = strTmpGrid4 & chr(11) & strLoopCnt4 
				strTmpGrid4 = strTmpGrid4 & chr(11) & iRow & chr(11) & chr(12)
			End If 
		Next
	End With
	CopyFromData = True                                                          '⊙: Processing is NG
End Function

'========================================================================================
' Function Name : DbQueryOk_one
' Function Desc : 
'========================================================================================
Function DbQueryOk_one()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SSShowData strTmpGrid
	frm1.vspdData.focus
	Set gActiveElement = document.ActiveElement
End Function 

'========================================================================================
' Function Name : DbQueryOk_two
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================

Function DbQueryOk_two()														'☆: 조회 성공후 실행로직 
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.SSShowData strTmpGrid2
	If  CopyFromData(1) = False Then exit Function 
	ggoSpread.Source = frm1.vspdData4
	ggoSpread.SSShowData strTmpGrid4
	frm1.vspdData2.focus
	Set gActiveElement = document.ActiveElement
End Function

'========================================================================================
' Function Name : DbQueryOk_three
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================

Function DbQueryOk_three()														'☆: 조회 성공후 실행로직 
	ggoSpread.Source = frm1.vspdData7
	ggoSpread.SSShowData strTmpGrid7
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData7,-1 , -1 ,C_DocCur7 ,C_DocAmt7 ,   "A" ,"Q","X","X")
	frm1.vspdData7.focus
	Set gActiveElement = document.ActiveElement
End Function


'========================================================================================
Function DbSave() 
End Function


'========================================================================================
Function DbSaveOk()
End Function


'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete()
	On Error Resume Next
End Function


'========================================================================================
' Function Name : ClickTab1
' Function Desc : This function tab1 click
'========================================================================================
Function ClickTab1()
	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)	 '~~~ 첫번째 Tab 
	gSelframeFlg = TAB1
	Call SetDefaultVal()
	frm1.cboIOFlag.focus

End Function

'========================================================================================
' Function Name : ClickTab2
' Function Desc : This function tab2 click
'========================================================================================
Function ClickTab2()
	If gSelframeFlg = TAB2 Then Exit Function
	Call changeTabs(TAB2)	 '~~~ 두번째 Tab 
	gSelframeFlg = TAB2
	Call SetDefaultVal()
	frm1.cboIOFlag2.focus 

End Function

'========================================================================================
' Function Name : ClickTab3
' Function Desc : This function tab3 click
'========================================================================================
Function ClickTab3()
	If gSelframeFlg = TAB3 Then Exit Function
	Call changeTabs(TAB3)	 '~~~ 두번째 Tab 
	gSelframeFlg = TAB3
	Call SetDefaultVal()
	frm1.vspdData7.focus

End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>


<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>><!-- ' 상위 여백 --></TD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>디스켓CheckList(세금)</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">	
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>디스켓CheckList(계산서)</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">	
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab3()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>디스켓CheckList(수출실적)</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
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
		<!--첫번째 TAB  -->
		<DIV ID="TabDiv"  SCROLL="no">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5">화일명</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtFileName" NAME="txtFileName" SIZE=30 MAXLENGTH=100 STYLE="TEXT-ALIGN: left" ALT="화일명" tag="12X" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOpenFilePath" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:GetOpenFilePath()"></TD>
									<TD CLASS="TD5">매입매출구분</TD>
									<TD CLASS="TD6"><SELECT ID="cboIOFlag" NAME="cboIOFlag" ALT="매입매출구분" STYLE="WIDTH: 98px" tag="12X"></SELECT></TD>
								<TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH="100%" HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5">사업자등록번호</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtRegNo" NAME="txtRegNo" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="사업자등록번호" tag="24X" ></TD>
								<TD CLASS="TD5">상호(법인명)</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtBizAreaNm" NAME="txtBizAreaNm" SIZE=30 MAXLENGTH=30 STYLE="TEXT-ALIGN: left" ALT="상호(법인명)" tag="24X" ></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">성명(대표자)</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtRepreNm" NAME="txtRepreNm" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="성명(대표자)" tag="24X" ></TD>
								<TD CLASS="TD5">사업장소재지</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtAddr" NAME="txtAddr" SIZE=39 MAXLENGTH=50 STYLE="TEXT-ALIGN: left" ALT="사업장소재지" tag="24X" ></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">업태</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtIndType" NAME="txtIndType" SIZE=17 MAXLENGTH=25 STYLE="TEXT-ALIGN: left" ALT="업태" tag="24X" ></TD>
								<TD CLASS="TD5">종목</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtIndClass" NAME="txtIndClass" SIZE=25 MAXLENGTH=25 STYLE="TEXT-ALIGN: left" ALT="종목" tag="24X" ></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">거래기간</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtStartDt" NAME="txtStartDt" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="거래기간" tag="24X" >
												&nbsp; ~ &nbsp;
												<INPUT TYPE=TEXT ID="txtEndDt" NAME="txtEndDt" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="거래기간" tag="24X" ></TD>
								<TD CLASS="TD5">작성일자</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtReportDt" NAME="txtReportDt" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="작성일자" tag="24X" ></TD>
							</TR>
							<TR>
								<TD WIDTH="100%" HEIGHT = "70%" COLSPAN="4">
								<script language =javascript src='./js/a6105ma1_vaSpread1_vspdData.js'></script></TD>
							</TR>
							<TR>
								<TD WIDTH="100%" HEIGHT = "30%"COLSPAN="4">
								<script language =javascript src='./js/a6105ma1_vaSpread1_vspdData1.js'></script></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>	
		</div>
		<!--두번째 TAB  -->
		<DIV ID="TabDiv"  SCROLL="no">
		<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5">화일명</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtFileName2" NAME="txtFileName2" SIZE=30 MAXLENGTH=100 STYLE="TEXT-ALIGN: left" ALT="화일명" tag="12X" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOpenFilePath" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:GetOpenFilePath()"></TD>
									<TD CLASS="TD5">매입매출구분</TD>
									<TD CLASS="TD6"><SELECT ID="cboIOFlag2" NAME="cboIOFlag2" ALT="매입매출구분" STYLE="WIDTH: 98px" tag="12X"></SELECT></TD>
								<TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH="100%" HEIGHT=* VALIGN=TOP>
							<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5">사업자등록번호</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtRegNo2" NAME="txtRegNo2" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="사업자등록번호" tag="24X" ></TD>
								<TD CLASS="TD5">법인명(상호)</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtBizAreaNm2" NAME="txtBizAreaNm2" SIZE=30 MAXLENGTH=30 STYLE="TEXT-ALIGN: left" ALT="상호(법인명)" tag="24X" ></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">성명(대표자)</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtRepreNm2" NAME="txtRepreNm2" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="성명(대표자)" tag="24X" ></TD>
								<TD CLASS="TD5">주민등록번호</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtPreRgstNo2" NAME="txtPreRgstNo2" SIZE=15 MAXLENGTH=15 STYLE="TEXT-ALIGN: left" ALT="주민등록번호" tag="24X" ></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">우편번호</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtZipCode2" NAME="txtZipCode2" SIZE=15 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" ALT="우편번호" tag="24X" ></TD>
								<TD CLASS="TD5">사업장소재지</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtAddr2" NAME="txtAddr2" SIZE=39 MAXLENGTH=50 STYLE="TEXT-ALIGN: left" ALT="사업장소재지" tag="24X" ></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">전화번호</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtTelno2" NAME="txtTelno2" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="전화번호" tag="24X" >
								<TD CLASS="TD5">작성일자</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtReportDt2" NAME="txtReportDt2" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="작성일자" tag="24X" ></TD>
							</TR>
							<TR>
								<TD WIDTH="100%" HEIGHT = "15%" COLSPAN="4">
								<script language =javascript src='./js/a6105ma1_vaSpread1_vspdData2.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">기구분</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtGiGubun3" NAME="txtGiGubun3" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="성명(대표자)" tag="24X" ></TD>
								<TD CLASS="TD5">신고구분</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtSingoGubun3" NAME="txtSingoGubun3" SIZE=15 MAXLENGTH=15 STYLE="TEXT-ALIGN: left" ALT="주민등록번호" tag="24X" ></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">세무서코드</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtTaxOffice3" NAME="txtTaxOffice3" SIZE=15 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" ALT="우편번호" tag="24X" ></TD>
								<TD CLASS="TD5">귀속년도</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtReturnYear3" NAME="txtReturnYear3" SIZE=15 MAXLENGTH=4 STYLE="TEXT-ALIGN: left" ALT="사업장소재지" tag="24X" ></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">거래기간</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtStartDt3" NAME="txtStartDt3" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="거래기간" tag="24X" >
												&nbsp; ~ &nbsp;
												<INPUT TYPE=TEXT ID="txtEndDt3" NAME="txtEndDt3" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="거래기간" tag="24X" ></TD>
								<TD CLASS="TD5">작성일자</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtReportDt3" NAME="txtReportDt3" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="작성일자" tag="24X" ></TD>
							</TR>
							<TR>
								<TD WIDTH="100%" HEIGHT = "20%" COLSPAN="4">
								<script language =javascript src='./js/a6105ma1_vaSpread1_vspdData3.js'></script></TD>
							</TR>
							<TR>
								<TD WIDTH="100%" HEIGHT = "45%"COLSPAN="4">
								<script language =javascript src='./js/a6105ma1_vaSpread1_vspdData4.js'></script></TD>
							</TR>
						</TABLE>
			
					</TD>
				</TR>
				
			</TABLE>
		</div>
		<!--세번째 TAB  -->
		<DIV ID="TabDiv"  SCROLL="no">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5">화일명</TD>
									<TD CLASS="TD656"><INPUT TYPE=TEXT ID="txtFileName3" NAME="txtFileName3" SIZE=30 MAXLENGTH=100 STYLE="TEXT-ALIGN: left" ALT="화일명" tag="12X" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOpenFilePath" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:GetOpenFilePath()"></TD>
								<TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH="100%" HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5">사업자등록번호</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtRegNo3" NAME="txtRegNo3" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="사업자등록번호" tag="24X" ></TD>
								<TD CLASS="TD5">상호(법인명)</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtBizAreaNm3" NAME="txtBizAreaNm3" SIZE=30 MAXLENGTH=30 STYLE="TEXT-ALIGN: left" ALT="상호(법인명)" tag="24X" ></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">성명(대표자)</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtRepreNm3" NAME="txtRepreNm3" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="성명(대표자)" tag="24X" ></TD>
								<TD CLASS="TD5">사업장소재지</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtAddr3" NAME="txtAddr3" SIZE=39 MAXLENGTH=50 STYLE="TEXT-ALIGN: left" ALT="사업장소재지" tag="24X" ></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">업태</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtIndType3" NAME="txtIndType3" SIZE=17 MAXLENGTH=25 STYLE="TEXT-ALIGN: left" ALT="업태" tag="24X" ></TD>
								<TD CLASS="TD5">종목</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtIndClass3" NAME="txtIndClass3" SIZE=25 MAXLENGTH=25 STYLE="TEXT-ALIGN: left" ALT="종목" tag="24X" ></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">귀속년월</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtYearMonth3" NAME="txtYearMonth3" SIZE=17 MAXLENGTH=25 STYLE="TEXT-ALIGN: left" ALT="업태" tag="24X" ></TD>
								<TD CLASS="TD5">신고구분</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtSingo3" NAME="txtSingo3" SIZE=25 MAXLENGTH=25 STYLE="TEXT-ALIGN: left" ALT="종목" tag="24X" ></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">거래기간</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtStart3" NAME="txtStart3" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="거래기간" tag="24X" >
												&nbsp; ~ &nbsp;
												<INPUT TYPE=TEXT ID="txtEnd3" NAME="txtEnd3" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="거래기간" tag="24X" ></TD>
								<TD CLASS="TD5">작성일자</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtReport3" NAME="txtReport3" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="작성일자" tag="24X" ></TD>
							</TR>
							<TR>
								<TD WIDTH="100%" HEIGHT = "70%" COLSPAN="4">
								<script language =javascript src='./js/a6105ma1_vaSpread1_vspdData7.js'></script></TD>
							</TR>
							<TR>
								<TD WIDTH="100%" HEIGHT = "30%"COLSPAN="4">
								<script language =javascript src='./js/a6105ma1_vaSpread1_vspdData8.js'></script></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>	
		</div>

		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="14" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hFileName" tag="14" TABINDEX="-1">
<script language =javascript src='./js/a6105ma1_OBJECT2_vspdData5.js'></script>
<script language =javascript src='./js/a6105ma1_OBJECT2_vspdData6.js'></script>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>
