<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID		    : A6105MA1
'*  4. Program Name         : �ΰ����Ű����CheckList��ȸ 
'*  5. Program Desc         : �ΰ����Ű����CheckList��ȸ 
'*  6. Component List       : +
'*  7. Modified date(First) : 2000/04/22
'*  8. Modified date(Last)  : 2002/07/31
'*  9. Modifier (First)     : Jong Hwan, Kim
'* 10. Modifier (Last)      : Hee Jung, Kim ; Nam Yo, Lee
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc"  -->								<!-- '��: ȭ��ó��ASP���� �����۾��� �ʿ��� ���  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!-- '��: �ش� ��ġ�� ���� �޶���, ��� ���  -->

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="javascript"	SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'��: indicates that All variables must be declared in advance

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_ID = "a6105mb1.asp"			'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_ID2 = "a6105mb2.asp"			'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_ID3 = "a6105mb3.asp"			'��: �����Ͻ� ���� ASP�� 
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
'��: Grid Columns 
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

Const TAB1 = 1																		'��: Tab�� ��ġ 
Const TAB2 = 2
Const TAB3 = 3


 '==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->	
'Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
'Dim lgIntGrpCount              ' Group View Size�� ������ ���� 
'Dim lgIntFlgMode               ' Variable is for Operation Status

'Dim lgStrPrevKey

'Dim lgLngCurRows

Dim lgBlnStartFlag				' �޼��� �����Ͽ� ���α׷� ���۽��� Check Flag

'========================================================================================================= 
'Grid field vspdData1, vspdData3
'========================================================================================================= 
Dim lgRegNOPu                   '����ڵ�Ϲ�ȣ�����
Dim lgPreRgstNoPu               '�ֹε�Ϲ�ȣ�����
Dim lgTotSum                    '�հ�
Dim lgExport					'�����ϴ���ȭ
Dim lgEtcTax					'��Ÿ������
'========================================================================================================= 
'Grid2, Grid4 
'========================================================================================================= 

lgRegNOPu       = "����ڵ�Ϲ�ȣ�����"  '1
lgPreRgstNoPu   = "�ֹε�Ϲ�ȣ�����"    '2
lgTotSum        = "�հ�"                '3
'========================================================================================================= 
'Grid8
'========================================================================================================= 
lgExport		= "�����ϴ���ȭ"  '1
lgEtcTax		= "��Ÿ������"    '2


 '==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 
 '----------------  ���� Global ������ ����  ----------------------------------------------------------- 
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
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
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
Sub initSpreadPosVariables()         '1.2 ������ Constants ���� �Ҵ� 
    '��: Grid Columns
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
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 

Sub SetDefaultVal()

	'lgBlnStartFlag = False		' �޼��� �����Ͽ� ���α׷� ���۽��� Check Flag
	
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
	Call initSpreadPosVariables()      '1.3 [initSpreadPosVariables] ȣ�� Logic �߰� 
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021214",,Parent.gAllowDragDropSpread            '2.1 [Spreadinit]�� Source Version No. �� DragDrop�������θ� ���� 

	With frm1.vspdData

		.MaxCols = C_IndClassNM + 1
		.MaxRows = 0


		.ReDraw = False

		Call GetSpreadColumnPos("A")
		ggoSpread.SSSetEdit  C_BPRgstNO, "����ڵ�Ϲ�ȣ", 12, , , 20
		ggoSpread.SSSetFloat C_PaperCnt,"�ż�", 9, 0, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat C_BlankCnt,"����", 9, parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat C_NetAmt,"���ް���", 18, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat C_VatAmt,"����", 18, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit  C_Code, "�ַ��ڵ�", 10, , , 10
		ggoSpread.SSSetEdit  C_BPNM, "��ȣ", 30, , , 30
		ggoSpread.SSSetEdit  C_IndTypeNM, "����", 17, , , 20
		ggoSpread.SSSetEdit  C_IndClassNM, "����", 25, , , 25
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)

		.ReDraw = True

	End With
	Call SetSpreadLock(0)
End Sub

Sub InitSpreadSheet1()
	Call initSpreadPosVariables()      '1.3 [initSpreadPosVariables] ȣ�� Logic �߰� 
	ggoSpread.Source = frm1.vspdData1
	ggoSpread.Spreadinit "V20021214",,Parent.gAllowDragDropSpread            '2.1 [Spreadinit]�� Source Version No. �� DragDrop�������θ� ���� 

	With frm1.vspdData1

		.ReDraw = False
		.MaxCols = C_VatAmtSum + 1

		.MaxRows = 0
		.MaxRows = 3
		Call GetSpreadColumnPos("B")
		ggoSpread.SSSetEdit  C_Title,      "",30, , , 25
		ggoSpread.SSSetFloat C_BPCntSum,   "�ŷ�ó��", 20, 0, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat C_PaperCntSum,"�ż�",     20, 0, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat C_NetAmtSum,  "���ް���", 27, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat C_VatAmtSum,  "����",     25, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

		.Row    = 1
		.Col    = 0 'C_Title
		.value  = lgRegNOPu      '"����ڵ�Ϲ�ȣ�����"
		.Col    = .MaxCols
		.text   = 1

		.Row    = 2
		.Col    = 0 'C_Title
		.value  = lgPreRgstNoPu    '"�ֹε�Ϲ�ȣ�����"
		.Col    = .MaxCols
		.text   = 2

		.Row    = 3
		.Col    = 0 'C_Title
		.value  = lgTotSum       '"�հ�"
		.Col    = .MaxCols
		.text   = 3

		.ReDraw = True
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	End With
	Call SetSpreadLock(1)
End Sub


'=============================================== 2.2.3 InitSpreadSheet() ========================================
Sub InitSpreadSheet2()
	Call initSpreadPosVariables()      '1.3 [initSpreadPosVariables] ȣ�� Logic �߰� 
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.Spreadinit "V20021214",,Parent.gAllowDragDropSpread            '2.1 [Spreadinit]�� Source Version No. �� DragDrop�������θ� ���� 
	With frm1.vspdData2
		.MaxCols = C_LoopCnt2
		.MaxRows = 0
		.ReDraw = False
		Call GetSpreadColumnPos("C")
		ggoSpread.SSSetEdit C_BRecord2  , "���ڵ屸��"    , 12, , , 12
		ggoSpread.SSSetEdit C_TaxOffice2, "�������ڵ�"    , 12, , , 12
		ggoSpread.SSSetEdit C_BPRgstNO2 , "����ڵ�Ϲ�ȣ" , 18, , , 18
		ggoSpread.SSSetEdit C_BPNM2     , "���θ�(��ȣ)"  , 20, , , 20
		ggoSpread.SSSetEdit C_BPPreNm2  , "��ǥ��(����)"  , 15, , , 20
		ggoSpread.SSSetEdit C_ZipCode2  , "�����ȣ"      , 9, , , 20
		ggoSpread.SSSetEdit C_Addr      , "����������"   , 30, , , 50
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_LoopCnt2,C_LoopCnt2,True)
		.ReDraw = True
	End With
	Call SetSpreadLock(2)  
End Sub    

'=============================================== 2.2.3 InitSpreadSheet() ========================================
Sub InitSpreadSheet3()        
	ggoSpread.Source = frm1.vspdData3		
	ggoSpread.Spreadinit "V20021214",,Parent.gAllowDragDropSpread            '2.1 [Spreadinit]�� Source Version No. �� DragDrop�������θ� ���� 
	With frm1.vspdData3	
		.MaxCols = C_LoopCnt3 + 1
		.MaxRows = 0
		.MaxRows = 3
		Call GetSpreadColumnPos("D")
		ggoSpread.SSSetEdit  C_Title3,      "",22, , , 25
		ggoSpread.SSSetEdit C_CRecord3      , "���ڵ屸��", 20, , , 10
		ggoSpread.SSSetFloat C_BPCntSum3    , "����ó��"  , 20, 0, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat C_PaperCntSum3 , "��꼭�ż�", 20, 0, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat C_NetAmtSum3   , "�ݾ�"     , 30, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

		.Row    = 1
		.Col    = 0 'C_Title3
		.value  = lgRegNOPu      '"����ڵ�Ϲ�ȣ�����"
		.Col    = .MaxCols
		.text   = 1

		.Row    = 2
		.Col    = 0 'C_Title3
		.value  = lgPreRgstNoPu    '"�ֹε�Ϲ�ȣ�����"
		.Col    = .MaxCols
		.text   = 2

		.Row    = 3
		.Col    = 0 'C_Title3
		.value  = lgTotSum       '"�հ�"
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
	ggoSpread.Spreadinit "V20021214",,Parent.gAllowDragDropSpread            '2.1 [Spreadinit]�� Source Version No. �� DragDrop�������θ� ���� 
	With frm1.vspdData4
		.MaxCols = C_LoopCnt4 + 1
		.MaxRows = 0
		.ReDraw = False

		Call GetSpreadColumnPos("E")
		ggoSpread.SSSetEdit C_DRecord4      , "���ڵ屸��", 15, , , 15
		ggoSpread.SSSetEdit C_BPRgstNO4     , "����ڵ�Ϲ�ȣ", 20, , , 20
		ggoSpread.SSSetEdit C_BPNM4         , "���θ�(��ȣ)", 20, , , 20		
		ggoSpread.SSSetFloat C_PaperCnt4    , "��꼭�ż�", 20, 0, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat C_NetAmt4      , "�ݾ�", 30, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
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
	ggoSpread.Spreadinit "V20021214",,Parent.gAllowDragDropSpread            '2.1 [Spreadinit]�� Source Version No. �� DragDrop�������θ� ���� 
	With frm1.vspdData7	
		.MaxCols = C_LocAmt7 + 1
		.MaxRows = 0

		.ReDraw = False
		Call GetSpreadColumnPos("F")
		ggoSpread.SSSetEdit C_ExportNo7, "����Ű��ȣ", 23, , , 20
		ggoSpread.SSSetEdit C_FnDt7, "����(��)����", 15, , , 20
		ggoSpread.SSSetEdit C_DocCur7, "�ŷ���ȭ", 10, , , 20
		'Call AppendNumberPlace("6","5","4")
		ggoSpread.SSSetFloat C_XchRate7,"ȯ��",11, parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec													  
		'Call AppendNumberPlace("7","13","2")
		ggoSpread.SSSetFloat C_DocAmt7,"��ȭ�ݾ�", 22, "A", ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		'Call AppendNumberPlace("8","15","0")
		ggoSpread.SSSetFloat C_LocAmt7,"��ȭ�ݾ�", 22, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		.ReDraw = True
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	End With
	Call SetSpreadLock(5)
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
Sub InitSpreadSheet6()
	ggoSpread.Source = frm1.vspdData8
	ggoSpread.Spreadinit "V20021227",,Parent.gAllowDragDropSpread            '2.1 [Spreadinit]�� Source Version No. �� DragDrop�������θ� ���� 
	With frm1.vspdData8	
		.ReDraw = False
		.MaxCols = C_LocSum8 + 1
		.MaxRows = 0
		.MaxRows = 3
		Call GetSpreadColumnPos("G")
		ggoSpread.SSSetEdit  C_Title8,      "",30, , , 25
		ggoSpread.SSSetFloat C_CntSum8,   "�Ǽ�", 20, 0, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec	
		ggoSpread.SSSetFloat C_DocSum8,  "��ȭ�ݾ�", 30, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat C_LocSum8,  "��ȭ�ݾ�", 30, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

		.Row    = 1
		.Col    = 0 'C_Title8
		.value  = lgExport      '"�����ϴ���ȭ"
		.Col    = .MaxCols
		.text   = 1

		.Row    = 2
		.Col    = 0 'C_Title8
		.value  = lgEtcTax    '"��Ÿ������"
		.Col    = .MaxCols
		.text   = 2

		.Row    = 3
		.Col    = 0 'C_Title8
		.value  = lgTotSum       '"�հ�"
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
            '��: Grid Columns
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
	Call LoadInfTB19029                           '��: Load table , B_numeric_format
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
	Call SetToolbar("1100000000011111")										'��: ��ư ���� ���� 
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
'	Description : ��ȸ������ ���(���ݰ�꼭)
'========================================================================================================= 
Function subCompany(Byval pStrLine)
 
	' �ڻ����� 
	' �ڷᱸ��(1), ����ڵ�Ϲ�ȣ(10), ��ȣ(30), ����(15), ����������(45), ����(17), ����(25), 
	' �ŷ��Ⱓ(6), �ŷ��Ⱓ(6), �ۼ�����(6), ����(9)

	Dim Cnt, ColCnt
	Dim strChr

	strChr = ""
	ColCnt = 0
	For Cnt = 1 To Len(pStrLine)
		If Asc(Mid(pStrLine, Cnt, 1)) >= 0 Then
			ColCnt = ColCnt + 1		' ������ 
		Else
			ColCnt = ColCnt + 2		' �ѱ� 
		End If
		
		strChr = strChr & Mid(pStrLine, Cnt, 1)

		Select Case ColCnt
			Case 1		' �ڷᱸ��(1)
				strChr = ""
			Case 11		' ����ڵ�Ϲ�ȣ(10)
				frm1.txtRegNo.value = strChr
				strChr = ""
			Case 41		' ��ȣ(30)
				frm1.txtBizAreaNm.value = strChr
				strChr = ""
			Case 56		' ����(15)
				frm1.txtRepreNm.value = strChr
				strChr = ""
			Case 101	' ����������(45)
				frm1.txtAddr.value = strChr
				strChr = ""
			Case 118	' ����(17)
				frm1.txtIndType.value = strChr
				strChr = ""
			Case 143	' ����(25)
				frm1.txtIndClass.value = strChr
				strChr = ""
			Case 149	' �ŷ��Ⱓ(6)
				frm1.txtStartDt.value = strChr
				strChr = ""
			Case 155	' �ŷ��Ⱓ(6)
				frm1.txtEndDt.value = strChr
				strChr = ""
			Case 161	' �ۼ�����(6)
				frm1.txtReportDt.value = strChr
				strChr = ""
			Case 170	' ����(9)
				strChr = ""
		End Select		
		If ColCnt >= 170 Then Exit For
	Next
End Function

'==========================================  2.1.1 subCompany2A()  ======================================
'	Name : subCompany2A()
'	Description : ��ȸ������ ���(��꼭) A-Record
'========================================================================================================= 

Function subCompany2A(Byval pStrLine)
 
	Dim Cnt, ColCnt
	Dim strChr

	strChr = ""
	ColCnt = 0
	For Cnt = 1 To Len(pStrLine)
		If Asc(Mid(pStrLine, Cnt, 1)) >= 0 Then
			ColCnt = ColCnt + 1		' ������ 
		Else
			ColCnt = ColCnt + 2		' �ѱ� 
		End If
		
		strChr = strChr & Mid(pStrLine, Cnt, 1)

		Select Case ColCnt
			Case 1		' ���ڵ屸��(1)
				strChr = ""
			Case 4		' ������(3)
				strChr = ""
			Case 12		' ��������(9)
				frm1.txtReportDt2.value = strChr
				strChr = ""
			Case 13		' �����ڱ���(1)
				strChr = ""
			Case 19		' �����븮�ΰ�����ȣ(6)
				strChr = ""
			Case 29		' ����ڵ�Ϲ�ȣ(10)
				frm1.txtRegNo2.value = strChr
				strChr = ""
			Case 69		' ���θ�(��ȣ)(10)
				frm1.txtBizAreaNm2.value = strChr
				strChr = ""
			Case 82		' �ֹε�Ϲ�ȣ(13)
				frm1.txtPreRgstNo2.value = strChr
				strChr = ""
			Case 112	' ��ǥ��(����)(30)
				frm1.txtRepreNm2.value = strChr
				strChr = ""
			Case 122	' �����������ȣ(10)
				frm1.txtZipCode2.value = strChr
				strChr = ""
			Case 192	' �������ּ�(70)
				frm1.txtAddr2.value = strChr
				strChr = ""
			Case 207	' ��ȭ��ȣ(15)
				frm1.txtTelno2.value = strChr
				strChr = ""
			Case 212	' ����Ǽ���(5)
				strChr = ""
			Case 215	' �ѱ��ڵ�����(3)
				strChr = ""
			Case 230	' �ۼ�����(6)
				strChr = ""
		End Select
		
		If ColCnt >= 230 Then Exit For
	Next

End Function

'==========================================  2.1.1 subCompany2B()  ======================================
'	Name : subCompany2B()
'	Description : ��ȸ������ ���(��꼭) B-Record
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
			ColCnt = ColCnt + 1		' ������ 
		Else
			ColCnt = ColCnt + 2		' �ѱ� 
		End If
		
		strChr = strChr & Mid(pStrLine, Cnt, 1)

		Select Case ColCnt
			Case 1		' ���ڵ屸��(1)
                strBRecord2 = strChr
				strChr = ""
			Case 4		' ������(3)
                strTaxOffice2 = strChr
				strChr = ""
			Case 10		' �Ϸù�ȣ(6)
				strChr = ""
			Case 20		' ����ڵ�Ϲ�ȣ(10)
                strBPRgstNO2 = strChr
				strChr = ""
			Case 60		' ���θ�(��ȣ)(40)
                strBPNM2 = strChr
				strChr = ""
			Case 90		'��ǥ��(����)(30)
                strBPPreNm2 = strChr
				strChr = ""
			Case 100	'���������ȣ(10)
                strZipCode2 = strChr
				strChr = ""
			Case 170	' ����������(70)
                strAddr = strChr
				strChr = ""
			Case 230	'����(60)
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
'	Description : ����ó���� (���ݰ�꼭)
'========================================================================================================= 

Function subPayment(Byval pStrLine,Byval pStrLineNo)
	
	' ����ó���� 
	' �ڷᱸ��(1), ����ڵ�Ϲ�ȣ(10), �Ϸù�ȣ(4), ����ڵ�Ϲ�ȣ(10), ��ȣ(30), ����(17), ����(25), �ż�(7), 
	' ������(2), ���ް���(14), ����(13), �ַ�����(1), �ַ��Ҹ�(1), �ǹ�ȣ(4), ����ó(3), ����(28)

	Dim Cnt, ColCnt
	Dim strChr
	dim LastAmt	
	Dim strBPRgstNO,strPaperCnt, strBlankCnt, strNetAmt, strVatAmt, strCode, strBPNM, strIndTypeNM, strIndClassNM

	strChr = ""
	ColCnt = 0
	
	For Cnt = 1 To Len(pStrLine)
		If Asc(Mid(pStrLine, Cnt, 1)) >= 0 Then
			ColCnt = ColCnt + 1		' ������ 
		Else
			ColCnt = ColCnt + 2		' �ѱ� 
		End If
		
		strChr = strChr & Mid(pStrLine, Cnt, 1)

		Select Case ColCnt
			Case 1		' �ڷᱸ��(1)
				strChr = ""
			Case 11		' ����ڵ�Ϲ�ȣ(10)
				strChr = ""
			Case 15		' �Ϸù�ȣ(4)
				strChr = ""
			Case 25		' ����ڵ�Ϲ�ȣ(10)
                strBPRgstNO = strChr
				strChr = ""
			Case 55		' ��ȣ(30)
                strBPNM = strChr
				strChr = ""
			Case 72		' ����(17)
                strIndTypeNM  = strChr
				strChr = ""
			Case 97		' ����(25)
                strIndClassNM  = strChr
				strChr = ""
			Case 104	' �ż�(7)
                strPaperCnt  = strChr
				strChr = ""
			Case 106	' ������(2)
                strBlankCnt  = strChr
				strChr = ""
			Case 120	' ���ް���(14)
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
			Case 133	' ����(13)
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
			Case 134	' �ַ�����(1)
				strChr = strChr & " / "
			Case 135	' �ַ��Ҹ�(1)
                strCode = strChr
				strChr = ""
			Case 139	' �ǹ�ȣ(4)
				strChr = ""
			Case 142	' ����ó(3)
				strChr = ""
			Case 170	' ����(28)
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
 
	' ����ó�հ����� 
	' �ڷᱸ��(1), ����ڵ�Ϲ�ȣ(10), ��ü����ó��(7), ���ݰ�꼭�ż�(7), ���ް���(15), ����(14), 
	' ����ں�����ó��(7),  ���ݰ�꼭�ż�(7), ���ް���(15), ����(14), 
	' �ֹκ�����ó��(7),  ���ݰ�꼭�ż�(7), ���ް���(15), ����(14), ����(30) 

	Dim Cnt, ColCnt
	Dim strChr
	Dim LastAmt
	Dim iRegNOPuRowNo, iPreRgstNoPuRowNo, iTotSumRowNo

	Call InitRowColVale("A", iRegNOPuRowNo, iPreRgstNoPuRowNo, iTotSumRowNo)

    strChr = ""
	ColCnt = 0
	For Cnt = 1 To Len(pStrLine)
		If Asc(Mid(pStrLine, Cnt, 1)) >= 0 Then
			ColCnt = ColCnt + 1		' ������ 
		Else
			ColCnt = ColCnt + 2		' �ѱ� 
		End If
		
		strChr = strChr & Mid(pStrLine, Cnt, 1)

		Select Case ColCnt
			Case 1		' �ڷᱸ��(1)
				strChr = ""
			Case 11		' ����ڵ�Ϲ�ȣ(10)
				strChr = ""
			Case 18		' ��ü����ó��(7)
				frm1.vspdData1.Row = iTotSumRowNo '3
				frm1.vspdData1.Col = C_BPCntSum '1 'jsk 20021216
				frm1.vspdData1.text = strChr
				strChr = ""
			Case 25		' ���ݰ�꼭�ż�(7)
				frm1.vspdData1.Row = iTotSumRowNo '3
				frm1.vspdData1.Col = C_PaperCntSum '2 'jsk 20021216
				frm1.vspdData1.text = strChr
				strChr = ""
			Case 40		' ���ް���(15)
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
			Case 54		' ����(14)
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
			Case 61		' ����ں�����ó��(7)
				frm1.vspdData1.Row = iRegNOPuRowNo '1
				frm1.vspdData1.Col =  C_BPCntSum '1 'jsk 20021216
				frm1.vspdData1.text = strChr
				strChr = ""
			Case 68		' ���ݰ�꼭�ż�(7)
				frm1.vspdData1.Row = iRegNOPuRowNo '1
				frm1.vspdData1.Col = C_PaperCntSum '2 'jsk 20021216
				frm1.vspdData1.text = strChr
				
				strChr = ""
			Case 83		' ���ް���(15)
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
			Case 97		' ����(14)
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
			Case 104	' �ֹκ�����ó��(7)
				frm1.vspdData1.Row = iPreRgstNoPuRowNo '2
				frm1.vspdData1.Col = C_BPCntSum '1 'jsk 20021216
				frm1.vspdData1.text = strChr
				strChr = ""
			Case 111	' ���ݰ�꼭�ż�(7)
				frm1.vspdData1.Row = iPreRgstNoPuRowNo '2
				frm1.vspdData1.Col = C_PaperCntSum '2 'jsk 20021216
				frm1.vspdData1.text = strChr
				strChr = ""
			Case 126	' ���ް���(15)
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
			Case 140	' ����(14)
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
			Case 170	' ����(30)
				strChr = ""
		End Select
		
		If ColCnt >= 170 Then Exit For

	Next
    frm1.vspdData1.Col   = frm1.vspdData1.MaxCols
	frm1.vspdData1.value = pStrLineNo
End Function


Function subRceipt(Byval pStrLine,Byval pStrLineNo)
 
	' ����ó���� 
	' �ڷᱸ��(1), ����ڵ�Ϲ�ȣ(10), �Ϸù�ȣ(4), ����ڵ�Ϲ�ȣ(10), ��ȣ(30), ����(17), ����(25), �ż�(7), 
	' ������(2), ���ް���(14), ����(13), �ַ�����(1), �ַ��Ҹ�(1), �ǹ�ȣ(4), ����ó(3), ����(28)

	Dim Cnt, ColCnt
	Dim strChr
	dim LastAmt	
	Dim str
	Dim strBPRgstNO,strPaperCnt, strBlankCnt, strNetAmt, strVatAmt, strCode, strBPNM, strIndTypeNM, strIndClassNM

	strChr = ""
	ColCnt = 0

	For Cnt = 1 To Len(pStrLine)
		If Asc(Mid(pStrLine, Cnt, 1)) >= 0 Then
			ColCnt = ColCnt + 1		' ������ 
		Else
			ColCnt = ColCnt + 2		' �ѱ� 
		End If
		
		strChr = strChr & Mid(pStrLine, Cnt, 1)

		Select Case ColCnt
			Case 1		' �ڷᱸ��(1)
				strChr = ""
			Case 11		' ����ڵ�Ϲ�ȣ(10)
				strChr = ""
			Case 15		' �Ϸù�ȣ(4)
				strChr = ""
			Case 25		' ����ڵ�Ϲ�ȣ(10)
                strBPRgstNO = strChr
				strChr = ""
			Case 55		' ��ȣ(30)
                strBPNM = strChr
				strChr = ""
			Case 72		' ����(17)
                strIndTypeNM = strChr
				strChr = ""
			Case 97		' ����(25)
                strIndClassNM = strChr
				strChr = ""
			Case 104	' �ż�(7)
                strPaperCnt = strChr
				strChr = ""
			Case 106	' ������(2)
                strBlankCnt = strChr
				strChr = ""
			Case 120	' ���ް���(14)
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
			Case 133	' ����(13)
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
			Case 134	' �ַ�����(1)
				strChr = strChr & " / "
			Case 135	' �ַ��Ҹ�(1)
                strCode = strChr
				strChr = ""
			Case 139	' �ǹ�ȣ(4)
				strChr = ""
			Case 142	' ����ó(3)
				strChr = ""
			Case 170	' ����(28)
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
'  subRceiptSum(pStrLine): ����ó�հ����� 
'  �ڷᱸ��(1), ����ڵ�Ϲ�ȣ(10), ����ó��(7), ��꼭�ż�(7), ���ް���(15), ����(14), ����(116)
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
			ColCnt = ColCnt + 1		' ������ 
		Else
			ColCnt = ColCnt + 2		' �ѱ� 
		End If
		
		strChr = strChr & Mid(pStrLine, Cnt, 1)
		Select Case ColCnt
			Case 1		' �ڷᱸ��(1)
				strChr = ""
			Case 11		' ����ڵ�Ϲ�ȣ(10)
				strChr = ""
			Case 18		' ����ó��(7)
				frm1.vspdData1.Row = iTotSumRowNo '3
				frm1.vspdData1.Col = C_BPCntSum '1 'jsk 20021216
				frm1.vspdData1.text = strChr
				strChr = ""
			Case 25		' ��꼭�ż�(7)
				frm1.vspdData1.Row = iTotSumRowNo '3
				frm1.vspdData1.Col = C_PaperCntSum '2 'jsk 20021216
				frm1.vspdData1.text = strChr
				strChr = ""
			Case 40		' ���ް���(15)
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
			Case 54		' ����(14)
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
			Case 170	' ����(116)
				strChr = ""
		End Select
		
		If ColCnt >= 170 Then Exit For
	Next
End Function

'==========================================  2.1.1 subPayment2()  ======================================
'	Name : subPayment2()
'	Description : ����ó���� (��꼭)
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
			ColCnt = ColCnt + 1		' ������ 
		Else
			ColCnt = ColCnt + 2		' �ѱ� 
		End If
		
		strChr = strChr & Mid(pStrLine, Cnt, 1)
		Select Case ColCnt
			Case 1		' ���ڵ屸��(1)
				frm1.vspdData6.Col = 1 'C_DRecord4  
				frm1.vspdData6.text = strChr
				strChr = ""				
			Case 3		' �ڷᱸ�� 
				strChr = ""
			Case 4		' �ⱸ��(1)
				strChr = ""
			Case 5		' �Ű���(1)
				strChr = ""
			Case 8		' �������ڵ�(3)
				strChr = ""
			Case 14		'�Ϸù�ȣ(6)
				strChr = ""
			Case 24		' �����ǹ��ڻ���ڵ�Ϲ�ȣ(10)
				strChr = ""
			Case 34	' ����ڵ�Ϲ�ȣ(10)
				frm1.vspdData6.Col = 2 'C_BPRgstNO4 
				frm1.vspdData6.text = strChr
				strChr = ""
			Case 74	' ���θ�(��ȣ)(40)
				frm1.vspdData6.Col = 3 'C_BPNM4 '
				frm1.vspdData6.text = strChr
				strChr = ""
			Case 79	' ��꼭�ż�(5)
				frm1.vspdData6.Col = 4 'C_PaperCnt4 '
				frm1.vspdData6.text = strChr
				strChr = ""
			Case 80	' ����ǥ��(1)
				signFlag = strChr
				strChr = ""
			Case 94	' �ݾ�(14)
				frm1.vspdData6.Col = 5 'C_NetAmt4 '
				If signFlag = "0" Then
					frm1.vspdData6.text = strChr	
				Else
					frm1.vspdData6.text = "-" & strChr	
				End If
				strChr = ""
				signFlag = ""
			Case 230	' ����(136)
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
'	Description : ����ó�հ����� (��꼭)
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
			ColCnt = ColCnt + 1		' ������ 
		Else
			ColCnt = ColCnt + 2		' �ѱ� 
		End If
		
		strChr = strChr & Mid(pStrLine, Cnt, 1)

		Select Case ColCnt
			Case 1		' ���ڵ屸��(1)
				frm1.vspdData5.Col = 1 'C_CRecord5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 3		' �ڷᱸ��(2)
				strChr = ""
			Case 4		' �ⱸ��(1)
				frm1.vspdData5.Col = 2 'C_Gigubun5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 5		' �Ű���(1)
				frm1.vspdData5.Col = 3 'C_SingoGubun5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 8		' ������(3)
				frm1.vspdData5.Col = 4 'C_TaxOffice5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 14		' �Ϸù�ȣ(6)
				strChr = ""
			Case 24		' �����ǹ��� ����ڵ�Ϲ�ȣ(10)
				strChr = ""
			Case 28		'�ͼӳ⵵(4)
				frm1.vspdData5.Col = 5 'C_ReturnYear5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 36		' �ŷ��Ⱓ���۳����(8)
				frm1.vspdData5.Col = 6 'C_StartDt5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 44		'�ŷ��Ⱓ��������(8)
				frm1.vspdData5.Col = 7 'C_EndDt5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 52	' �ۼ�����(8)
				frm1.vspdData5.Col = 8 'C_ReportDt5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			
			Case 58	' ����ó��(6) - �հ� 
				frm1.vspdData5.Col = 15 'C_HBPCntSum5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 64	'��꼭�ż�(6) - �հ� 
				frm1.vspdData5.Col = 16 'C_HPaperCntSum5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 65	'����ǥ��(1) - �հ� 
				signFlag = strChr
				strChr = ""
			Case 79	'�ݾ�(14) - �հ� 
				frm1.vspdData5.Col = 17 'C_HNetAmtSum5 '
				If signFlag = "0" Then
					frm1.vspdData5.text = strChr	
				Else
					frm1.vspdData5.text = "-" & strChr	
				End If
				strChr = ""
				signFlag = ""
			Case 85	' ����ó��(6)-����ڵ�Ϲ�ȣ����� 
				frm1.vspdData5.Col = 9 'C_BBPCntSum5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 91	'��꼭�ż�(6) - ����ڵ�Ϲ�ȣ����� 
				frm1.vspdData5.Col = 10 'C_BPaperCntSum5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 92	' ����ǥ�� - ����ڵ�Ϲ�ȣ����� 
				signFlag = strChr
				strChr = ""
			Case 106	'�ݾ� - ����ڵ�Ϲ�ȣ����� 
				frm1.vspdData5.Col = 11 'C_BNetAmtSum5 '
				If signFlag = "0" Then
					frm1.vspdData5.text = strChr	
				Else
					frm1.vspdData5.text = "-" & strChr	
				End If
				strChr = ""
				signFlag = ""
			Case 112	'����ó�� - �ֹε�Ϲ�ȣ ����� 
				frm1.vspdData5.Col = 12 'C_RBPCntSum5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 118	' ��꼭�ż� - �ֹε�Ϲ�ȣ����� 
				frm1.vspdData5.Col = 13 'C_RPaperCntSum5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 119	' ����ǥ�� - �ֹε�Ϲ�ȣ����� 
				signFlag = strChr
				strChr = ""
			Case 133	' �ݾ� - �ֹε�Ϲ�ȣ����� 
				frm1.vspdData5.Col = 14 'C_RNetAmtSum5 '
				If signFlag = "0" Then
					frm1.vspdData5.text = strChr	
				Else
					frm1.vspdData5.text = "-" & strChr	
				End If
				strChr = ""
				signFlag = ""
				
			Case 230	' ����(97)
				strChr = ""
		End Select
		
		If ColCnt >= 230 Then Exit For
	Next
	frm1.vspdData5.Col = 18 'C_LoopCnt5 '
	frm1.vspdData5.text = BGubunCnt
	
End Function
'==========================================  2.1.1 subRceipt2()  ======================================
'	Name : subRceipt2()
'	Description : ����ó���� (��꼭)
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
			ColCnt = ColCnt + 1		' ������ 
		Else
			ColCnt = ColCnt + 2		' �ѱ� 
		End If
		
		strChr = strChr & Mid(pStrLine, Cnt, 1)
		Select Case ColCnt
			Case 1		' ���ڵ屸��(1)
				frm1.vspdData6.Col = 1 'C_DRecord4 '
				frm1.vspdData6.text = strChr
				strChr = ""
				
			Case 3		' �ڷᱸ�� 
				strChr = ""
			Case 4		' �ⱸ��(1)
				strChr = ""
			Case 5		' �Ű���(1)
				strChr = ""
			Case 8		' �������ڵ�(3)
				strChr = ""
			Case 14		'�Ϸù�ȣ(6)
				strChr = ""
			Case 24		' �����ǹ��ڻ���ڵ�Ϲ�ȣ(10)
				strChr = ""
			Case 34	' ����ڵ�Ϲ�ȣ(10)
				frm1.vspdData6.Col = 2 'C_BPRgstNO4 '
				frm1.vspdData6.text = strChr
				strChr = ""
			Case 74	' ���θ�(��ȣ)(40)
				frm1.vspdData6.Col = 3 'C_BPNM4 '
				frm1.vspdData6.text = strChr
				strChr = ""
			Case 79	' ��꼭�ż�(5)
				frm1.vspdData6.Col = 4 'C_PaperCnt4 '
				frm1.vspdData6.text = strChr
				strChr = ""
			Case 80	' ����ǥ��(1)
				signFlag = strChr
				strChr = ""
			Case 94	' �ݾ�(14)
				frm1.vspdData6.Col = 5 'C_NetAmt4 '
				If signFlag = "0" Then
					frm1.vspdData6.text = strChr	
				Else
					frm1.vspdData6.text = "-" & strChr	
				End If
				strChr = ""
				signFlag = ""
			Case 230	' ����(136)
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
'	Description : ����ó�հ����� (��꼭)
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
			ColCnt = ColCnt + 1		' ������ 
		Else
			ColCnt = ColCnt + 2		' �ѱ� 
		End If
		
		strChr = strChr & Mid(pStrLine, Cnt, 1)

		Select Case ColCnt
			Case 1		' ���ڵ屸��(1)
				frm1.vspdData5.Col = 1 'C_CRecord5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 3		' �ڷᱸ��(2)
				strChr = ""
			Case 4		' �ⱸ��(1)
				frm1.vspdData5.Col = 2 'C_Gigubun5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 5		' �Ű���(1)
				frm1.vspdData5.Col = 3 'C_SingoGubun5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 8		' ������(3)
				frm1.vspdData5.Col = 4 'C_TaxOffice5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 14		' �Ϸù�ȣ(6)
				strChr = ""
			Case 24		' �����ǹ��� ����ڵ�Ϲ�ȣ(10)
				strChr = ""
			Case 28		'�ͼӳ⵵(4)
				frm1.vspdData5.Col = 5 'C_ReturnYear5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 36		' �ŷ��Ⱓ���۳����(8)
				frm1.vspdData5.Col = 6 'C_StartDt5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 44		'�ŷ��Ⱓ��������(8)
				frm1.vspdData5.Col = 7 'C_EndDt5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 52	' �ۼ�����(8)
				frm1.vspdData5.Col = 8 'C_ReportDt5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			
			Case 58	' ����ó��(6) - �հ� 
				frm1.vspdData5.Col = 15 'C_HBPCntSum5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 64	'��꼭�ż�(6) - �հ� 
				frm1.vspdData5.Col = 16 'C_HPaperCntSum5 '
				frm1.vspdData5.text = strChr
				strChr = ""
			Case 65	'����ǥ��(1) - �հ� 
				signFlag = strChr
				strChr = ""
			Case 79	'�ݾ�(14) - �հ� 
				frm1.vspdData5.Col = 17 'C_HNetAmtSum5 '
				If signFlag = "0" Then
					frm1.vspdData5.text = strChr	
				Else
					frm1.vspdData5.text = "-" & strChr	
				End If
				strChr = ""
				signFlag = ""
			Case 230	' ����(97)
				strChr = ""
		End Select
		
		If ColCnt >= 230 Then Exit For
	Next
	frm1.vspdData5.Col = 18 'C_LoopCnt5 '
	frm1.vspdData5.text = BGubunCnt
End Function


Function subCompany3(Byval pStrLine)
 
	' �ڻ����� 
	' �ڷᱸ��(1), �ͼӳ��(6), �Ű���(1), ����ڵ�Ϲ�ȣ(10), ��ȣ(30), ����(15), ����������(45), 
	' ����(17), ����(25), �ŷ��Ⱓ(8), �ŷ��Ⱓ(8), �ۼ�����(8), ����(6)

	Dim Cnt, ColCnt
	Dim strChr

	strChr = ""
	ColCnt = 0
	For Cnt = 1 To Len(pStrLine)
		If Asc(Mid(pStrLine, Cnt, 1)) >= 0 Then
			ColCnt = ColCnt + 1		' ������ 
		Else
			ColCnt = ColCnt + 2		' �ѱ� 
		End If
		
		strChr = strChr & Mid(pStrLine, Cnt, 1)
		Select Case ColCnt
			Case 1		' �ڷᱸ��(1)
				strChr = ""
			Case 7		' �ͼӳ��(6)
				frm1.txtYearMonth3.value = strChr
				strChr = ""
			Case 8		' �Ű���(1)
				frm1.txtSingo3.value = strChr
				strChr = ""	
				
			Case 18		' ����ڵ�Ϲ�ȣ(10)
				frm1.txtRegNo3.value = strChr
				strChr = ""
			Case 48		' ��ȣ(30)
				frm1.txtBizAreaNm3.value = strChr
				strChr = ""
			Case 63		' ����(15)
				frm1.txtRepreNm3.value = strChr
				strChr = ""
			Case 108	' ����������(45)
				frm1.txtAddr3.value = strChr
				strChr = ""
			Case 125	' ����(17)
				frm1.txtIndType3.value = strChr
				strChr = ""
			Case 150	' ����(25)
				frm1.txtIndClass3.value = strChr
				strChr = ""
			Case 158	' �ŷ��Ⱓ(8)
				frm1.txtStart3.value = strChr
				strChr = ""
			Case 166	' �ŷ��Ⱓ(8)
				frm1.txtEnd3.value = strChr
				strChr = ""
			Case 174	' �ۼ�����(8)
				frm1.txtReport3.value = strChr
				strChr = ""
			Case 180	' ����(6)
				strChr = ""
		End Select		
		If ColCnt >= 180 Then Exit For
	Next
End Function

 '==========================================  2.1.1 subExportList()  ======================================
'	Name : subExportList()
'	Description : ����������� 
'========================================================================================================= 

Function subExportList(Byval pStrLine,Byval pStrLineNo)
	'�ڷᱸ��(1), �ͼӳ��(6), �Ű���(1), ����ڵ�Ϲ�ȣ(10), �Ϸù�ȣ(7), ����Ű��ȣ(15), ����(��)����(8)
	'��ȭ�ڵ�(3), ȯ��(5), ȯ��(4), ��ȭ�ݾ�(13), ��ȭ�ݾ�(2), ��ȭ�ݾ�(15), ����(90)
	Dim Cnt, ColCnt
	Dim strChr
	Dim LastAmt
	Dim strExportNo7, strFnDt7, strDocCur7, strXchRate7, strDocAmt7, strLocAmt7
	strChr = ""
	ColCnt = 0
	
	For Cnt = 1 To Len(pStrLine)
		If Asc(Mid(pStrLine, Cnt, 1)) >= 0 Then
			ColCnt = ColCnt + 1		' ������ 
		Else
			ColCnt = ColCnt + 2		' �ѱ� 
		End If
		
		strChr = strChr & Mid(pStrLine, Cnt, 1)

		Select Case ColCnt
			Case 1		' �ڷᱸ��(1)
				strChr = ""
			Case 7		' �ͼӳ��(6)
				strChr = ""
			Case 8		' �Ű���(1)
				strChr = ""
			Case 18		' ����ڵ�Ϲ�ȣ(10)
				strChr = ""
			Case 25		' �Ϸù�ȣ(7)
				strChr = ""
			Case 40		' ����Ű��ȣ(15)
				strExportNo7 = strChr 
				strChr = ""
			Case 48		' ����(��)����(8)
				strFnDt7 = strChr 
				strChr = ""
			Case 51		' ��ȭ�ڵ�(3)
				strDocCur7 = strChr 
				strChr = ""
			Case 56		' ȯ��(5)
				strChr = strChr & parent.gComNumDec '"."
			Case 60		' ȯ��(4)
				strXchRate7 = strChr 
				strChr = ""
			Case 73		'��ȭ�ݾ�(13)
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
			Case 75		'��ȭ�ݾ�(2)
				strDocAmt7 = strChr 
				strChr = ""	
			Case 90	' ��ȭ�ݾ�(15)
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
			Case 180	' ����(90)
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

	'�ڷᱸ��(1), �ͼӳ��(1), �Ű���(1), ����ڵ�Ϲ�ȣ(10), ��ü�Ǽ�(7), ��ȭ�ݾ�(��ü)(13), ��ȭ�ݾ�(��ü)(2)
	'��ȭ�ݾ�(��ü)(15), ��ȭ�ݾ�(��ü)(15), �����ϴ���ȭ�Ǽ�(7), ��ȭ�ݾ�(����)(13), ��ȭ�ݾ�(����)(2), ��ȭ�ݾ�(����)(15)
	'������Ǽ�(7), ��ȭ�ݾ�(������)(13), ��ȭ�ݾ�(������)(2), ��ȭ�ݾ�(������)(15), ����(51)
	Dim Cnt, ColCnt
	Dim strChr
	Dim LastAmt

	Dim iExport, ilgEtcTax, iTotSumRowNo

	Call InitRowColVale("C",iExport, ilgEtcTax, iTotSumRowNo)
	
	strChr = ""
	ColCnt = 0
	For Cnt = 1 To Len(pStrLine)
		If Asc(Mid(pStrLine, Cnt, 1)) >= 0 Then
			ColCnt = ColCnt + 1		' ������ 
		Else
			ColCnt = ColCnt + 2		' �ѱ� 
		End If
		
		strChr = strChr & Mid(pStrLine, Cnt, 1)

		Select Case ColCnt
			Case 1		' �ڷᱸ��(1)
				strChr = ""
			Case 7		' �ͼӳ��(1)
				strChr = ""
			Case 8		' �Ű���(1)
				strChr = ""
			Case 18		' ����ڵ�Ϲ�ȣ(10)
				strChr = ""
			Case 25		' ��ü�Ǽ�(7)
				frm1.vspdData8.Row = iTotSumRowNo '3
				frm1.vspdData8.Col = C_CntSum8
				frm1.vspdData8.text = strChr
				strChr = ""
			Case 38		' ��ȭ�ݾ�(��ü)(13)
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
			Case 40		' ��ȭ�ݾ�(��ü)(2)
				frm1.vspdData8.Row = iTotSumRowNo '3
				frm1.vspdData8.Col = C_DocSum8
				frm1.vspdData8.text = strChr
				strChr = ""
			Case 55		' ��ȭ�ݾ�(��ü)(15)
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
			Case 62		' �����ϴ���ȭ�Ǽ�(7)
				frm1.vspdData8.Row = iExport '1
				frm1.vspdData8.Col = C_CntSum8
				frm1.vspdData8.text = strChr
				strChr = ""
			Case 75		' ��ȭ�ݾ�(����)(13)
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
			Case 77		' ��ȭ�ݾ�(����)(2)
				frm1.vspdData8.Row = iExport '1
				frm1.vspdData8.Col = C_DocSum8
				frm1.vspdData8.text = strChr
				strChr = ""
			Case 92		' ��ȭ�ݾ�(����)(15)
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
			Case 99	' ������Ǽ�(7)
				frm1.vspdData8.Row = ilgEtcTax '2
				frm1.vspdData8.Col = C_CntSum8
				frm1.vspdData8.text = strChr
				strChr = ""
			Case 112	' ��ȭ�ݾ�(������)(13)
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
			Case 114	' ��ȭ�ݾ�(������)(2)
				frm1.vspdData8.Row = ilgEtcTax '2
				frm1.vspdData8.Col = C_DocSum8
				frm1.vspdData8.text =  strChr
				strChr = ""
			Case 129	' ��ȭ�ݾ�(������)(15)
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
			Case 180	' ����(51)
				strChr = ""
		End Select
		
		If ColCnt >= 180 Then Exit For

	Next

End Function


 '******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
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
     gMouseClickStatus = "SPC"	'Split �����ڵ� 
    
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
    gMouseClickStatus = "SP2C"	'Split �����ڵ� 
    
    Set gActiveSpdSheet = frm1.vspdData1

    If frm1.vspdData1.MaxRows <= 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
End Sub

Sub vspdData2_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("0000111111")    
     gMouseClickStatus = "SP3C"	'Split �����ڵ� 
    
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
     gMouseClickStatus = "SP4C"	'Split �����ڵ� 
    
    Set gActiveSpdSheet = frm1.vspdData3

    If frm1.vspdData3.MaxRows <= 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
End Sub

Sub vspdData4_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("0000111111")    
     gMouseClickStatus = "SP5C"	'Split �����ڵ� 
    
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
     gMouseClickStatus = "SP6C"	'Split �����ڵ� 
    
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
    gMouseClickStatus = "SP7C"	'Split �����ڵ� 
    
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
'   Event Desc : Spread Split �����ڵ� 
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
    On Error Resume Next                                                    '��: Protect system from crashing
End Function


'========================================================================================
Function FncNext() 
    On Error Resume Next                                                    '��: Protect system from crashing
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
    Call parent.FncExport(parent.C_MULTI)												'��: ȭ�� ���� 
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                         '��:ȭ�� ����, Tab ���� 
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
		' ����� ������ �ִ��� Ȯ���Ѵ�.
		If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")			'��: "Will you destory previous data"
	
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
	Err.Clear                                                               '��: Protect system from crashing      
	DbQuery = False                                                         '��: Processing is NG
    If gSelframeFlg = TAB1 Then 
		frm1.hFileName.value = lgFilePath
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'��: �����Ͻ� ó�� ASP�� ���� 
		strVal = strVal & "&txtFileName=" & Trim(frm1.txtFileName.value)
		strVal = strVal & "&cboFlag="     & Trim(frm1.cboIOFlag.value)
	ElseIf gSelframeFlg = TAB2 Then
		frm1.hFileName.value = lgFilePath2
		strVal = BIZ_PGM_ID2 & "?txtMode=" & parent.UID_M0001							'��: �����Ͻ� ó�� ASP�� ���� 
		strVal = strVal & "&txtFileName=" & Trim(frm1.txtFileName2.value)
		strVal = strVal & "&cboFlag="     & Trim(frm1.cboIOFlag2.value)
	ElseIf gSelframeFlg = TAB3 Then
		frm1.hFileName.value = lgFilePath3
		strVal = BIZ_PGM_ID3 & "?txtMode=" & parent.UID_M0001							'��: �����Ͻ� ó�� ASP�� ���� 
		strVal = strVal & "&txtFileName=" & Trim(frm1.txtFileName3.value)
	End If	
	strVal = strVal & "&hFileName="   & Trim(frm1.hFileName.value)
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ����	
	DbQuery = True                                                          '��: Processing is NG
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
	Err.Clear                                                               '��: Protect system from crashing
	CopyFromData = False                                                         '��: Processing is NG

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
				'//vspdData3 setting (������ ī��)
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

				'/// ����ڵ�Ϲ�ȣ ����� 
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

				'//�ֹε�Ϲ�ȣ ����� 
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

				'//�հ� 
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

				'//������ ��츸 �������� (����ڵ�Ϲ�ȣ������� ����ó���� �����Ұ�츸 ����)////////////////////
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
				'//vspdData4 setting(������ ī��)

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
	CopyFromData = True                                                          '��: Processing is NG
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
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk_two()														'��: ��ȸ ������ ������� 
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
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk_three()														'��: ��ȸ ������ ������� 
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
	Call changeTabs(TAB1)	 '~~~ ù��° Tab 
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
	Call changeTabs(TAB2)	 '~~~ �ι�° Tab 
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
	Call changeTabs(TAB3)	 '~~~ �ι�° Tab 
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
		<TD <%=HEIGHT_TYPE_00%>><!-- ' ���� ���� --></TD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>����CheckList(����)</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">	
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>����CheckList(��꼭)</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">	
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab3()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>����CheckList(�������)</font></td>
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
		<!--ù��° TAB  -->
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
									<TD CLASS="TD5">ȭ�ϸ�</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtFileName" NAME="txtFileName" SIZE=30 MAXLENGTH=100 STYLE="TEXT-ALIGN: left" ALT="ȭ�ϸ�" tag="12X" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOpenFilePath" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:GetOpenFilePath()"></TD>
									<TD CLASS="TD5">���Ը��ⱸ��</TD>
									<TD CLASS="TD6"><SELECT ID="cboIOFlag" NAME="cboIOFlag" ALT="���Ը��ⱸ��" STYLE="WIDTH: 98px" tag="12X"></SELECT></TD>
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
								<TD CLASS="TD5">����ڵ�Ϲ�ȣ</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtRegNo" NAME="txtRegNo" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="����ڵ�Ϲ�ȣ" tag="24X" ></TD>
								<TD CLASS="TD5">��ȣ(���θ�)</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtBizAreaNm" NAME="txtBizAreaNm" SIZE=30 MAXLENGTH=30 STYLE="TEXT-ALIGN: left" ALT="��ȣ(���θ�)" tag="24X" ></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">����(��ǥ��)</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtRepreNm" NAME="txtRepreNm" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="����(��ǥ��)" tag="24X" ></TD>
								<TD CLASS="TD5">����������</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtAddr" NAME="txtAddr" SIZE=39 MAXLENGTH=50 STYLE="TEXT-ALIGN: left" ALT="����������" tag="24X" ></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">����</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtIndType" NAME="txtIndType" SIZE=17 MAXLENGTH=25 STYLE="TEXT-ALIGN: left" ALT="����" tag="24X" ></TD>
								<TD CLASS="TD5">����</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtIndClass" NAME="txtIndClass" SIZE=25 MAXLENGTH=25 STYLE="TEXT-ALIGN: left" ALT="����" tag="24X" ></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">�ŷ��Ⱓ</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtStartDt" NAME="txtStartDt" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="�ŷ��Ⱓ" tag="24X" >
												&nbsp; ~ &nbsp;
												<INPUT TYPE=TEXT ID="txtEndDt" NAME="txtEndDt" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="�ŷ��Ⱓ" tag="24X" ></TD>
								<TD CLASS="TD5">�ۼ�����</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtReportDt" NAME="txtReportDt" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="�ۼ�����" tag="24X" ></TD>
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
		<!--�ι�° TAB  -->
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
									<TD CLASS="TD5">ȭ�ϸ�</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtFileName2" NAME="txtFileName2" SIZE=30 MAXLENGTH=100 STYLE="TEXT-ALIGN: left" ALT="ȭ�ϸ�" tag="12X" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOpenFilePath" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:GetOpenFilePath()"></TD>
									<TD CLASS="TD5">���Ը��ⱸ��</TD>
									<TD CLASS="TD6"><SELECT ID="cboIOFlag2" NAME="cboIOFlag2" ALT="���Ը��ⱸ��" STYLE="WIDTH: 98px" tag="12X"></SELECT></TD>
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
								<TD CLASS="TD5">����ڵ�Ϲ�ȣ</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtRegNo2" NAME="txtRegNo2" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="����ڵ�Ϲ�ȣ" tag="24X" ></TD>
								<TD CLASS="TD5">���θ�(��ȣ)</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtBizAreaNm2" NAME="txtBizAreaNm2" SIZE=30 MAXLENGTH=30 STYLE="TEXT-ALIGN: left" ALT="��ȣ(���θ�)" tag="24X" ></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">����(��ǥ��)</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtRepreNm2" NAME="txtRepreNm2" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="����(��ǥ��)" tag="24X" ></TD>
								<TD CLASS="TD5">�ֹε�Ϲ�ȣ</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtPreRgstNo2" NAME="txtPreRgstNo2" SIZE=15 MAXLENGTH=15 STYLE="TEXT-ALIGN: left" ALT="�ֹε�Ϲ�ȣ" tag="24X" ></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">�����ȣ</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtZipCode2" NAME="txtZipCode2" SIZE=15 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" ALT="�����ȣ" tag="24X" ></TD>
								<TD CLASS="TD5">����������</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtAddr2" NAME="txtAddr2" SIZE=39 MAXLENGTH=50 STYLE="TEXT-ALIGN: left" ALT="����������" tag="24X" ></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">��ȭ��ȣ</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtTelno2" NAME="txtTelno2" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="��ȭ��ȣ" tag="24X" >
								<TD CLASS="TD5">�ۼ�����</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtReportDt2" NAME="txtReportDt2" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="�ۼ�����" tag="24X" ></TD>
							</TR>
							<TR>
								<TD WIDTH="100%" HEIGHT = "15%" COLSPAN="4">
								<script language =javascript src='./js/a6105ma1_vaSpread1_vspdData2.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">�ⱸ��</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtGiGubun3" NAME="txtGiGubun3" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="����(��ǥ��)" tag="24X" ></TD>
								<TD CLASS="TD5">�Ű���</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtSingoGubun3" NAME="txtSingoGubun3" SIZE=15 MAXLENGTH=15 STYLE="TEXT-ALIGN: left" ALT="�ֹε�Ϲ�ȣ" tag="24X" ></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">�������ڵ�</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtTaxOffice3" NAME="txtTaxOffice3" SIZE=15 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" ALT="�����ȣ" tag="24X" ></TD>
								<TD CLASS="TD5">�ͼӳ⵵</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtReturnYear3" NAME="txtReturnYear3" SIZE=15 MAXLENGTH=4 STYLE="TEXT-ALIGN: left" ALT="����������" tag="24X" ></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">�ŷ��Ⱓ</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtStartDt3" NAME="txtStartDt3" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="�ŷ��Ⱓ" tag="24X" >
												&nbsp; ~ &nbsp;
												<INPUT TYPE=TEXT ID="txtEndDt3" NAME="txtEndDt3" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="�ŷ��Ⱓ" tag="24X" ></TD>
								<TD CLASS="TD5">�ۼ�����</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtReportDt3" NAME="txtReportDt3" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="�ۼ�����" tag="24X" ></TD>
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
		<!--����° TAB  -->
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
									<TD CLASS="TD5">ȭ�ϸ�</TD>
									<TD CLASS="TD656"><INPUT TYPE=TEXT ID="txtFileName3" NAME="txtFileName3" SIZE=30 MAXLENGTH=100 STYLE="TEXT-ALIGN: left" ALT="ȭ�ϸ�" tag="12X" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOpenFilePath" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:GetOpenFilePath()"></TD>
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
								<TD CLASS="TD5">����ڵ�Ϲ�ȣ</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtRegNo3" NAME="txtRegNo3" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="����ڵ�Ϲ�ȣ" tag="24X" ></TD>
								<TD CLASS="TD5">��ȣ(���θ�)</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtBizAreaNm3" NAME="txtBizAreaNm3" SIZE=30 MAXLENGTH=30 STYLE="TEXT-ALIGN: left" ALT="��ȣ(���θ�)" tag="24X" ></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">����(��ǥ��)</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtRepreNm3" NAME="txtRepreNm3" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="����(��ǥ��)" tag="24X" ></TD>
								<TD CLASS="TD5">����������</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtAddr3" NAME="txtAddr3" SIZE=39 MAXLENGTH=50 STYLE="TEXT-ALIGN: left" ALT="����������" tag="24X" ></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">����</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtIndType3" NAME="txtIndType3" SIZE=17 MAXLENGTH=25 STYLE="TEXT-ALIGN: left" ALT="����" tag="24X" ></TD>
								<TD CLASS="TD5">����</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtIndClass3" NAME="txtIndClass3" SIZE=25 MAXLENGTH=25 STYLE="TEXT-ALIGN: left" ALT="����" tag="24X" ></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">�ͼӳ��</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtYearMonth3" NAME="txtYearMonth3" SIZE=17 MAXLENGTH=25 STYLE="TEXT-ALIGN: left" ALT="����" tag="24X" ></TD>
								<TD CLASS="TD5">�Ű���</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtSingo3" NAME="txtSingo3" SIZE=25 MAXLENGTH=25 STYLE="TEXT-ALIGN: left" ALT="����" tag="24X" ></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">�ŷ��Ⱓ</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtStart3" NAME="txtStart3" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="�ŷ��Ⱓ" tag="24X" >
												&nbsp; ~ &nbsp;
												<INPUT TYPE=TEXT ID="txtEnd3" NAME="txtEnd3" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="�ŷ��Ⱓ" tag="24X" ></TD>
								<TD CLASS="TD5">�ۼ�����</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtReport3" NAME="txtReport3" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="�ۼ�����" tag="24X" ></TD>
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
