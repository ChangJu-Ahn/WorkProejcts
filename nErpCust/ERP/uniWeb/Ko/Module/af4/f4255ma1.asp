<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Finance
'*  2. Function Name        : ���Աݰ��� 
'*  3. Program ID           : a4225ma1
'*  4. Program Name         : ���Աݸ�Ƽ��ȯ 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/04/28
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Jeong Yong Kyun
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. History              :
'**********************************************************************************************
 -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################
'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->

<!--'==========================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="javascript"   SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../ag/AcctCtrl.vbs">				</SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit                                                             '��: indicates that All variables must be declared in advance

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Const BIZ_PGM_QUERY_ID = "f4255mb1.ASP"
Const BIZ_PGM_SAVE_ID  = "f4255mb2.ASP"
Const BIZ_PGM_DEL_ID   = "f4255mb3.ASP"
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================

Const TAB1 = 1																		'��: Tab�� ��ġ 
Const TAB2 = 2
Const TAB3 = 3

' ��ݵ�� vspdData4
Dim C_REPAY_MEAN_SEQ
Dim C_REPAY_TYPE   
Dim C_REPAY_POP   
Dim C_REPAY_TYPE_NM
Dim C_BANK_ACCT_NO       
Dim C_BANK_ACCT_NO_POP  
Dim C_BANK_CD
DIm C_BANK_NM
Dim C_REPAY_ACCT_CD
Dim C_REPAY_ACCT_CD_POP 
DIm C_REPAY_ACCT_NM
DIm C_DOCCUR
DIm C_DOCCUR_POP
DIM C_XCH_RATE
Dim C_REPAY_AMT
Dim C_REPAY_LOC_AMT
Dim C_DESC
 
'���Ա� vspdData1
Dim C_LOAN_NO
Dim C_LOAN_DT
Dim C_LOAN_DUE_DT
Dim C_LOAN_PLAN_DT
DIm C_LOAN_DOCCUR
DIm C_LOAN_XCH_RATE
Dim C_REPAY_PLAN_AMT
Dim C_REPAY_PLAN_LOC_AMT
Dim C_REPAY_INT_DFR_AMT
Dim C_REPAY_INT_DFR_LOC_AMT
Dim C_INT_XCH_RATE
Dim C_REPAY_PLAN_INT_AMT
Dim C_REPAY_PLAN_INT_LOC_AMT
Dim C_REPAY_INT_ACCT_CD
Dim C_REPAY_INT_ACCT_CD_POP
Dim C_REPAY_INT_ACCT_NM
Dim C_LOAN_BAL_AMT
Dim C_LOAN_BAL_LOC_AMT
Dim C_LOAN_RDP_TOT_AMT
Dim C_LOAN_RDP_TOT_LOC_AMT
Dim C_LOAN_INT_TOT_AMT
Dim C_LOAN_INT_TOT_LOC_AMT
Dim C_REPAY_ITEM_DESC  
Dim C_REPAY_PAY_OBJ

'�ܾ�ó�� vspdData
Dim C_ITEM_SEQ
Dim C_ACCT_CD
Dim C_ACCT_CD_POP
Dim C_ACCT_CD_NM
Dim C_ETC_DOCCUR
Dim C_ITEMAMT
Dim C_ITEMLOCAMT
Dim C_ITEMDESC

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim  strMode
Dim  intItemCnt

 '==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 
 '----------------  ���� Global ������ ����  ----------------------------------------------------------- 
Dim  IsOpenPop          
Dim  lgRetFlag
Dim  lgCheckIntAmt
Dim  gSelframeFlg
Dim  lgCurrRow
Dim  lgDBSaveOK
<% Dim  dtToday

dtToday = GetSvrDate
%>

 '++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 

 '#########################################################################################################
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'######################################################################################################### 
'======================================================================================================
' Name : initSpreadPosVariables()
' Description : �׸���(��������) �÷� ���� ���� �ʱ�ȭ 
'=======================================================================================================
Sub initSpreadPosVariables(ByVal pvSpdNo)
	Select Case UCase(Trim(pvSpdNo))
		Case "A"
			C_REPAY_MEAN_SEQ		 = 1
			C_REPAY_TYPE			 = 2
			C_REPAY_POP				 = 3
			C_REPAY_TYPE_NM			 = 4
			C_BANK_ACCT_NO			 = 5
			C_BANK_ACCT_NO_POP		 = 6
			C_BANK_CD				 = 7
			C_BANK_NM				 = 8
			C_REPAY_ACCT_CD			 = 9 
			C_REPAY_ACCT_CD_POP		 = 10
			C_REPAY_ACCT_NM		 = 11
			C_DOCCUR				 = 12 
			C_DOCCUR_POP			 = 13
			C_XCH_RATE				 = 14
			C_REPAY_AMT				 = 15
			C_REPAY_LOC_AMT			 = 16
			C_DESC					 = 17
 		Case "B"
			C_LOAN_NO                = 1   
			C_LOAN_DT                = 2
			C_LOAN_DUE_DT			 = 3
			C_LOAN_PLAN_DT			 = 4
			C_LOAN_DOCCUR			 = 5
			C_LOAN_XCH_RATE			 = 6
			C_REPAY_PLAN_AMT		 = 7
			C_REPAY_PLAN_LOC_AMT	 = 8
			C_REPAY_INT_DFR_AMT      = 9
			C_REPAY_INT_DFR_LOC_AMT  = 10
			C_INT_XCH_RATE			 = 11
			C_REPAY_PLAN_INT_AMT	 = 12
			C_REPAY_PLAN_INT_LOC_AMT = 13  
			C_REPAY_INT_ACCT_CD		 = 14
			C_REPAY_INT_ACCT_CD_POP  = 15
			C_REPAY_INT_ACCT_NM		 = 16
			C_LOAN_BAL_AMT			 = 17
			C_LOAN_BAL_LOC_AMT		 = 18
			C_LOAN_RDP_TOT_AMT		 = 19
			C_LOAN_RDP_TOT_LOC_AMT	 = 20
			C_LOAN_INT_TOT_AMT		 = 21
			C_LOAN_INT_TOT_LOC_AMT   = 22
			C_REPAY_ITEM_DESC        = 23
			C_REPAY_PAY_OBJ			 = 24
		Case "C"
			C_ITEM_SEQ               = 1
			C_ACCT_CD                = 2
			C_ACCT_CD_POP			 = 3
			C_ACCT_CD_NM			 = 4
			C_ETC_DOCCUR             = 5
			C_ITEMAMT				 = 6
			C_ITEMLOCAMT			 = 7
			C_ITEMDESC				 = 8
			
	End Select	
End Sub 
'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub  InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE            'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgLngCurRows = 0  
    gSelframeFlg = TAB1
    lgDBSaveOK = 0    
    lgSortKey = 1
    lgCheckIntAmt = False
End Sub


'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'********************************************************************************************************* 
Sub  SetDefaultVal()
	frm1.txtRePayDT.Text = UniConvDateAToB("<%=dtToday%>", parent.gServerDateFormat,parent.gDateFormat)

	frm1.hOrgChangeId.value = parent.gChangeOrgId	
	frm1.txtRepayIntLocAmt.text = "0"
	frm1.txtRepayTotLocAmt.text = "0"
	frm1.txtEtcLocAmt.text			= "0"
	frm1.txtPaymTotLocAmt.text		= "0"
		
   lgBlnFlgChgValue = False
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub  LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)
	Call initSpreadPosVariables(pvSpdNo)
	Select Case UCase(Trim(pvSpdNo))
		Case "A"
			With frm1.vspdData4

				.MaxCols = C_DESC + 1
				.Col =.MaxCols
				.ColHidden = true
				ggoSpread.Source = frm1.vspdData4
				.Redraw = False	
				.MaxRows = 0
				ggoSpread.SpreadInit "V200303220",,parent.gAllowDragDropSpread 

				Call GetSpreadColumnPos(pvSpdNo)
	
				ggoSpread.SSSetEdit		C_REPAY_MEAN_SEQ,   "����",			  3, 3
				ggoSpread.SSSetEdit		C_REPAY_TYPE,       "�������",      10, 3, , , 2
				ggoSpread.SSSetButton	C_REPAY_POP
				ggoSpread.SSSetEdit		C_REPAY_TYPE_NM,    "���������",    15, , , 20, 2 
				ggoSpread.SSSetEdit		C_BANK_ACCT_NO ,    "���¹�ȣ",      15, 3
				ggoSpread.SSSetButton	C_BANK_ACCT_NO_POP		       
				ggoSpread.SSSetEdit		C_BANK_CD,		    "�����ڵ�",		 15, , ,	30
				ggoSpread.SSSetEdit		C_BANK_NM,		    "�����",	 	 15, , ,	30
				ggoSpread.SSSetEdit     C_REPAY_ACCT_CD,	"��ݰ����ڵ�"  ,12,,,20,2
				ggoSpread.SSSetButton   C_REPAY_ACCT_CD_POP 
				ggoSpread.SSSetEdit		C_REPAY_ACCT_NM,	"��ݰ����ڵ��",20, , , 30
				ggoSpread.SSSetEdit		C_DOCCUR,		    "�ŷ���ȭ",       9, , , 10, 2
				ggoSpread.SSSetButton	C_DOCCUR_POP
				ggoSpread.SSSetFloat	C_XCH_RATE,	        "ȯ��",           8, Parent.ggExchRateNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
				ggoSpread.SSSetFloat	C_REPAY_AMT,        "��ݱݾ�",      15, "A"  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec																 
				ggoSpread.SSSetFloat	C_REPAY_LOC_AMT,    "��ݱݾ�(�ڱ�)",15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec        
				ggoSpread.SSSetEdit		C_DESC,             "���",			 20, 3
				
				Call ggoSpread.MakePairsColumn(C_DOCCUR,C_DOCCUR_POP)
				Call ggoSpread.MakePairsColumn(C_REPAY_TYPE,C_REPAY_POP)
				Call ggoSpread.MakePairsColumn(C_REPAY_ACCT_CD,C_REPAY_ACCT_CD_POP)
				Call ggoSpread.SSSetColHidden(C_REPAY_MEAN_SEQ,C_REPAY_MEAN_SEQ,True)				
'				Call ggoSpread.SSSetColHidden(C_BANK_CD,C_BANK_CD,True)	
				
				.Redraw = True
			End With
		Case "B"
			With frm1.vspdData1
    
				.MaxCols = C_REPAY_PAY_OBJ + 1    
				.Col =.MaxCols
				.ColHidden = true
				ggoSpread.Source = frm1.vspdData1
				.Redraw = False	
			
				.MaxRows = 0
				ggoSpread.SpreadInit "V20030323",,parent.gAllowDragDropSpread 

				Call GetSpreadColumnPos(pvSpdNo)
	
				ggoSpread.SSSetEdit   C_LOAN_NO,				"���Աݹ�ȣ"			,18, 2
				ggoSpread.SSSetDate   C_LOAN_DT,				"������"				,10, 2, parent.gDateFormat
				ggoSpread.SSSetDate   C_LOAN_DUE_DT,			"��������"				,10, 2, parent.gDateFormat  		        
				ggoSpread.SSSetDate   C_LOAN_PLAN_DT,			"��ȯ������"			,10, 2, parent.gDateFormat  		        
				ggoSpread.SSSetEdit   C_LOAN_DOCCUR,			"�ŷ���ȭ"				, 9, , , 10, 2
				ggoSpread.SSSetFloat  C_LOAN_XCH_RATE,			"����ȯ��"				,10, Parent.ggExchRateNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
				ggoSpread.SSSetFloat  C_REPAY_PLAN_AMT,			"���ݻ�ȯ��"			,15, "A" ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat  C_REPAY_PLAN_LOC_AMT,		"���ݻ�ȯ��(�ڱ�)"		,15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat  C_REPAY_INT_DFR_AMT,		"���������ھ�"		    ,15, "A" ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat  C_REPAY_INT_DFR_LOC_AMT,	"���������ھ�(�ڱ�)"	,15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat  C_INT_XCH_RATE,			"ȯ��"					,10, Parent.ggExchRateNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
				ggoSpread.SSSetFloat  C_REPAY_PLAN_INT_AMT,		"�������޾�"			,15, "A" ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat  C_REPAY_PLAN_INT_LOC_AMT,	"�������޾�(�ڱ�)"		,15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetEdit   C_REPAY_INT_ACCT_CD,      "���ں�����"			,12,,,20,2
				ggoSpread.SSSetButton C_REPAY_INT_ACCT_CD_POP
				ggoSpread.SSSetEdit   C_REPAY_INT_ACCT_NM   ,   "���ں�������"		,12,,,20,2
				ggoSpread.SSSetFloat  C_LOAN_BAL_AMT,			"�����ܾ�"				,15, "A" ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat  C_LOAN_BAL_LOC_AMT,		"�����ܾ�(�ڱ�)"		,15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat  C_LOAN_RDP_TOT_AMT,		"���ݻ�ȯ�Ѿ�"			,15, "A" ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat  C_LOAN_RDP_TOT_LOC_AMT,	"���ݻ�ȯ�Ѿ�(�ڱ�)"	,15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat  C_LOAN_INT_TOT_AMT,		"���������Ѿ�"			,15, "A" ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat  C_LOAN_INT_TOT_LOC_AMT,	"���������Ѿ�(�ڱ�)"	,15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetEdit   C_REPAY_ITEM_DESC,        "���"					,20, 0
				ggoSpread.SSSetEdit   C_REPAY_PAY_OBJ,	        "������������"			,20, 0

				Call ggoSpread.MakePairsColumn(C_REPAY_INT_ACCT_CD,C_REPAY_INT_ACCT_CD_POP)
				Call ggoSpread.SSSetColHidden(C_REPAY_PAY_OBJ,C_REPAY_PAY_OBJ,True)


				.Redraw = True
			End With
		Case "C"
			With frm1.vspdData
				ggoSpread.Source = frm1.vspdData
				ggoSpread.Spreadinit "V20030218",,parent.gAllowDragDropSpread    
				
				.MaxCols = C_ITEMDESC + 1
				.Col = .MaxCols				'��: ������Ʈ�� ��� Hidden Column
				.ColHidden = True
				.MaxRows = 0
				.ReDraw = False

			    Call GetSpreadColumnPos(pvSpdNo)
			   
			    ggoSpread.SSSetEdit   C_ITEM_SEQ,    "����"      ,5 , 2, -1, 5
				ggoSpread.SSSetEdit   C_ACCT_CD,     "�����ڵ�"  ,15, , , 18
				ggoSpread.SSSetButton C_ACCT_CD_POP
				ggoSpread.SSSetEdit   C_ACCT_CD_NM,  "�����ڵ��",25, , , 30
				ggoSpread.SSSetEdit   C_ETC_DOCCUR,  "��ȭ"		 ,15, , , 18
				ggoSpread.SSSetFloat  C_ITEMAMT,     "�ݾ�"      ,20, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat  C_ITEMLOCAMT,  "�ݾ�(�ڱ�)",25, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetEdit   C_ITEMDESC,    "���"      ,40, , , 128

				Call ggoSpread.MakePairsColumn(C_ACCT_CD,C_ACCT_CD_POP)
				Call ggoSpread.SSSetColHidden(C_ITEM_SEQ,C_ITEM_SEQ,True)												'������Ʈ�� ��� Hidden Column
				Call ggoSpread.SSSetColHidden(C_ETC_DOCCUR,C_ETC_DOCCUR,True)												'������Ʈ�� ��� Hidden Column				
				
				.ReDraw = True
			End with
	End Select	
	
    intItemCnt = 0
    Call SetSpreadLock(pvSpdNo)
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)
    With frm1
    	Select Case UCase(Trim(pvSpdNo))
			Case "A"
				ggoSpread.source = .vspdData4      
				.vspdData4.ReDraw = False         
				ggoSpread.SpreadLock	C_REPAY_TYPE			,-1, C_REPAY_TYPE      , -1
				ggoSpread.SpreadLock	C_REPAY_TYPE_NM			,-1, C_REPAY_TYPE_NM   , -1
				ggoSpread.SpreadLock	C_BANK_ACCT_NO			,-1, C_BANK_ACCT_NO    , -1
				ggoSpread.SpreadLock	C_BANK_ACCT_NO_POP		,-1, C_BANK_ACCT_NO_POP, -1
				ggoSpread.SpreadLock	C_BANK_CD				,-1, C_BANK_CD         , -1
				ggoSpread.SpreadLock	C_BANK_NM				,-1, C_BANK_NM         , -1
				ggoSpread.SpreadLock	C_REPAY_ACCT_CD			,-1, C_REPAY_ACCT_CD   , -1
				ggoSpread.SpreadLock	C_REPAY_ACCT_NM		,-1, C_REPAY_ACCT_NM, -1
				ggoSpread.SpreadLock	C_DOCCUR				,-1, C_DOCCUR          , -1
				ggoSpread.SpreadLock	C_XCH_RATE				,-1, C_XCH_RATE        , -1																																
				ggoSpread.SpreadUnLock	C_REPAY_AMT				,-1, C_REPAY_AMT       , -1 
				ggoSpread.SpreadLock	C_REPAY_LOC_AMT			,-1, C_REPAY_LOC_AMT   , -1
				ggoSpread.SpreadUnLock	C_DESC					,-1, C_DESC            , -1 
'				ggoSpread.SSSetRequired C_REPAY_AMT				,-1, -1
				.vspdData4.ReDraw = True   
			Case "B"     
				ggoSpread.source = .vspdData1      
				ggoSpread.SpreadLock	C_LOAN_NO					,-1, C_LOAN_NO				 , -1
				ggoSpread.SpreadLock	C_LOAN_DT					,-1, C_LOAN_DT				 , -1
				ggoSpread.SpreadLock	C_LOAN_DUE_DT				,-1, C_LOAN_DUE_DT			 , -1
				ggoSpread.SpreadLock	C_LOAN_PLAN_DT				,-1, C_LOAN_PLAN_DT			 , -1
				ggoSpread.SpreadLock	C_LOAN_DOCCUR				,-1, C_LOAN_DOCCUR			 , -1
				ggoSpread.SpreadLock	C_LOAN_XCH_RATE				,-1, C_LOAN_XCH_RATE	     , -1
				ggoSpread.SpreadUnLock	C_REPAY_PLAN_AMT			,-1, C_REPAY_PLAN_AMT	     , -1
				ggoSpread.SpreadLock	C_REPAY_PLAN_LOC_AMT		,-1, C_REPAY_PLAN_LOC_AMT	 , -1
				ggoSpread.SpreadLock	C_REPAY_INT_DFR_AMT			,-1, C_REPAY_INT_DFR_AMT     , -1  
				ggoSpread.SpreadLock	C_REPAY_INT_DFR_LOC_AMT		,-1, C_REPAY_INT_DFR_LOC_AMT , -1
				ggoSpread.SpreadLock	C_INT_XCH_RATE				,-1, C_INT_XCH_RATE			 , -1
				ggoSpread.SpreadUnLock	C_REPAY_PLAN_INT_AMT		,-1, C_REPAY_PLAN_INT_AMT    , -1  
				ggoSpread.SpreadLock	C_REPAY_PLAN_INT_LOC_AMT	,-1, C_REPAY_PLAN_INT_LOC_AMT, -1  
				ggoSpread.SpreadUnLock	C_REPAY_INT_ACCT_CD			,-1, C_REPAY_INT_ACCT_CD	 , -1  
				ggoSpread.SpreadUnLock	C_REPAY_INT_ACCT_CD_POP		,-1, C_REPAY_INT_ACCT_CD_POP , -1  
				ggoSpread.SpreadLock	C_REPAY_INT_ACCT_NM 		,-1, C_REPAY_INT_ACCT_NM  , -1
				ggoSpread.SpreadLock	C_LOAN_BAL_AMT				,-1, C_LOAN_BAL_AMT			 , -1  								
				ggoSpread.SpreadLock	C_LOAN_BAL_LOC_AMT			,-1, C_LOAN_BAL_LOC_AMT		 , -1  																  								
				ggoSpread.SpreadLock	C_LOAN_RDP_TOT_AMT			,-1, C_LOAN_RDP_TOT_AMT		 , -1  								
				ggoSpread.SpreadLock	C_LOAN_RDP_TOT_LOC_AMT		,-1, C_LOAN_RDP_TOT_LOC_AMT  , -1  												
				ggoSpread.SpreadLock	C_LOAN_INT_TOT_AMT			,-1, C_LOAN_INT_TOT_AMT		 , -1  								
				ggoSpread.SpreadLock	C_LOAN_INT_TOT_LOC_AMT		,-1, C_LOAN_INT_TOT_LOC_AMT  , -1  												
				ggoSpread.SpreadunLock	C_REPAY_ITEM_DESC			,-1, C_REPAY_ITEM_DESC       , -1
				ggoSpread.SpreadunLock	C_REPAY_PAY_OBJ				,-1, C_REPAY_PAY_OBJ       , -1      

				.vspdData1.ReDraw = True   
			Case "C"
				ggoSpread.source = .vspdData       
				.vspdData.ReDraw = False         
				ggoSpread.SpreadLock	C_ITEM_SEQ              ,-1, C_ITEM_SEQ        , -1
				ggoSpread.SpreadunLock	C_ACCT_CD				,-1, C_ACCT_CD		   , -1
				ggoSpread.SpreadunLock	C_ACCT_CD_POP			,-1, C_ACCT_CD_POP     , -1
				ggoSpread.SpreadLock	C_ACCT_CD_NM			,-1, C_ACCT_CD_NM      , -1
				ggoSpread.SpreadUnLock	C_ITEMAMT				,-1, C_ITEMAMT         , -1
				ggoSpread.SpreadLock	C_ITEMLOCAMT			,-1, C_ITEMLOCAMT      , -1
				ggoSpread.SpreadUnLock	C_ITEMDESC				,-1, C_ITEMDESC        , -1

				.vspdData.ReDraw = True  
         End Select
    End With 
    Call SetSpreadColor(-1, -1,pvSpdNo)            
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal lRow, ByVal lRow2,ByVal pvSpd)
	Dim iSelFlg

	If pvSpd = "" Then
		iSelFlg = gSelframeFlg
	Else
		If pvSpd = "A" Then	
			iSelFlg = 1
		Elseif	pvSpd = "B" Then
			iSelFlg = 2
		ElseIf pvSpd = "C" Then
			iSelFlg = 3
		ElseIf pvSpd = "D" Then			
			iSelFlg = gSelframeFlg
		End If
	End If		
	
	With frm1
		Select Case iSelFlg
			Case TAB1
				ggoSpread.Source = .vspdData4	     
				.vspdData4.ReDraw = False 	    
				ggoSpread.SSSetRequired  C_REPAY_TYPE		,lRow, lRow2    	      
				ggoSpread.SSSetProtected C_REPAY_TYPE_NM	,lRow, lRow2
				ggoSpread.SSSetProtected C_BANK_ACCT_NO   	,lRow, lRow2
				ggoSpread.SSSetProtected C_BANK_CD      	,lRow, lRow2
				ggoSpread.SSSetProtected C_BANK_NM      	,lRow, lRow2
				ggoSpread.SSSetRequired  C_REPAY_ACCT_CD	,lRow, lRow2   
				ggoSpread.SSSetProtected C_REPAY_ACCT_NM	,lRow, lRow2  
				ggoSpread.SSSetProtected C_BANK_ACCT_NO		,lRow, lRow2 
				ggoSpread.SSSetRequired  C_DOCCUR			,lRow, lRow2
				ggoSpread.SSSetProtected C_XCH_RATE			,lRow, lRow2
				ggoSpread.SSSetRequired  C_REPAY_AMT		,lRow, lRow2
				ggoSpread.SSSetProtected C_REPAY_LOC_AMT	,lRow, lRow2
'				ggoSpread.SpreadUnLock   C_DESC				,-1, C_DESC  , -1

				.vspdData4.ReDraw = True		
			
			Case TAB2
				ggoSpread.Source = .vspdData1	     
				.vspdData1.ReDraw = False 	  
				ggoSpread.SSSetRequired  C_REPAY_PLAN_AMT		,lRow, lRow2
				ggoSpread.SSSetRequired  C_REPAY_PLAN_INT_AMT	,lRow, lRow2
				If lgCheckIntAmt = True Then
					ggoSpread.SSSetRequired  C_REPAY_INT_ACCT_CD	,lRow, lRow2												
					ggoSpread.SpreadUnLock	 C_REPAY_INT_ACCT_CD_POP,-1, C_REPAY_INT_ACCT_CD_POP , -1      					
				Else
					ggoSpread.SpreadLock	 C_REPAY_INT_ACCT_CD	,-1, C_REPAY_INT_ACCT_CD     , -1   				
					ggoSpread.SpreadLock	 C_REPAY_INT_ACCT_CD_POP,-1, C_REPAY_INT_ACCT_CD_POP , -1      										
				End If
				
				.vspdData1.ReDraw = True	
			Case TAB3
				ggoSpread.Source = .vspdData	     
				.vspdData.ReDraw = False 	  
				ggoSpread.SSSetProtected C_ITEM_SEQ		, lRow, lRow2   ' 
				ggoSpread.SSSetRequired  C_ACCT_CD		, lRow, lRow2	' �����ڵ� 
				ggoSpread.SSSetProtected C_ACCT_CD_NM	, lRow, lRow2   ' �����ڵ��		
				ggoSpread.SSSetRequired  C_ITEMAMT		, lRow, lRow2	' �ݾ� 
				ggoSpread.SSSetProtected C_ITEMLOCAMT	, lRow, lRow2   ' �ݾ�(�ڱ�)

				.vspdData.ReDraw = True	          
		End Select				 
	End With	
End Sub

'======================================================================================================
' Function Name : GetSpreadColumnPos
' Function Desc : This method call saved columnorder
'=======================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData4
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)		

			C_REPAY_MEAN_SEQ		 = iCurColumnPos(1)
			C_REPAY_TYPE			 = iCurColumnPos(2)
			C_REPAY_POP				 = iCurColumnPos(3)
			C_REPAY_TYPE_NM			 = iCurColumnPos(4)
			C_BANK_ACCT_NO			 = iCurColumnPos(5)
			C_BANK_ACCT_NO_POP		 = iCurColumnPos(6)
			C_BANK_CD				 = iCurColumnPos(7)
			C_BANK_NM				 = iCurColumnPos(8)
			C_REPAY_ACCT_CD			 = iCurColumnPos(9) 
			C_REPAY_ACCT_CD_POP		 = iCurColumnPos(10)
			C_REPAY_ACCT_NM		 = iCurColumnPos(11)
			C_DOCCUR				 = iCurColumnPos(12) 
			C_DOCCUR_POP			 = iCurColumnPos(13)
			C_XCH_RATE				 = iCurColumnPos(14)
			C_REPAY_AMT				 = iCurColumnPos(15)
			C_REPAY_LOC_AMT			 = iCurColumnPos(16)
			C_DESC					 = iCurColumnPos(17)
		Case "B"
			ggoSpread.Source = frm1.vspdData1
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)		
						
			C_LOAN_NO                = iCurColumnPos(1)   
			C_LOAN_DT                = iCurColumnPos(2)
			C_LOAN_DUE_DT			 = iCurColumnPos(3)
			C_LOAN_PLAN_DT			 = iCurColumnPos(4)
			C_LOAN_DOCCUR			 = iCurColumnPos(5)
			C_LOAN_XCH_RATE			 = iCurColumnPos(6)
			C_REPAY_PLAN_AMT		 = iCurColumnPos(7)
			C_REPAY_PLAN_LOC_AMT	 = iCurColumnPos(8)
			C_REPAY_INT_DFR_AMT      = iCurColumnPos(9)
			C_REPAY_INT_DFR_LOC_AMT  = iCurColumnPos(10)
			C_INT_XCH_RATE			 = iCurColumnPos(11)
			C_REPAY_PLAN_INT_AMT	 = iCurColumnPos(12)
			C_REPAY_PLAN_INT_LOC_AMT = iCurColumnPos(13)  
			C_REPAY_INT_ACCT_CD		 = iCurColumnPos(14)
			C_REPAY_INT_ACCT_CD_POP  = iCurColumnPos(15)
			C_REPAY_INT_ACCT_NM      = iCurColumnPos(16)
			C_LOAN_BAL_AMT			 = iCurColumnPos(17)
			C_LOAN_BAL_LOC_AMT		 = iCurColumnPos(18)
			C_LOAN_RDP_TOT_AMT		 = iCurColumnPos(19)
			C_LOAN_RDP_TOT_LOC_AMT	 = iCurColumnPos(20)
			C_LOAN_INT_TOT_AMT		 = iCurColumnPos(21)
			C_LOAN_INT_TOT_LOC_AMT   = iCurColumnPos(22)
			C_REPAY_ITEM_DESC        = iCurColumnPos(23)
			C_REPAY_PAY_OBJ			 = iCurColumnPos(24)
		Case "C"
			ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            
			C_ITEM_SEQ               = iCurColumnPos(1)
			C_ACCT_CD                = iCurColumnPos(2)
			C_ACCT_CD_POP			 = iCurColumnPos(3)
			C_ACCT_CD_NM			 = iCurColumnPos(4)
			C_ETC_DOCCUR             = iCurColumnPos(5)     
			C_ITEMAMT				 = iCurColumnPos(6)
			C_ITEMLOCAMT			 = iCurColumnPos(7)
			C_ITEMDESC				 = iCurColumnPos(8)

	End select
End Sub

'======================================================================================================
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'=======================================================================================================
'======================================================================================================
' Function Name : OpenPopupGL
' Function Desc : This method Open The Popup window for GL
'=======================================================================================================
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
	
	arrParam(0) = Trim(frm1.txthGlNo.value)							'ȸ����ǥ��ȣ 
	arrParam(1) = ""												'Reference��ȣ 

	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	frm1.txtRePayNO.focus
End Function

'======================================================================================================
' Function Name : OpenPopupTempGL
' Function Desc : This method Open The Popup window for TempGL
'=======================================================================================================
Function OpenPopuptempGL()
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
	
	arrParam(0) = Trim(frm1.txthTempGlNo.value)						'ȸ����ǥ��ȣ 
	arrParam(1) = ""												'Reference��ȣ 

	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	frm1.txtRePayNO.focus
End Function

'======================================================================================================
'   Function Name : OpenPopupLoan
'	Function Desc : ���Ա����� �˾� 
'======================================================================================================
Function OpenPopupLoan()
	Dim arrRet
	Dim arrParam(3)
	Dim iCalledAspName
	Dim i
	
	iCalledAspName = AskPRAspName("F4250RA1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "F4250RA1", "X")
		IsOpenPop = False
		Exit Function
	End If	

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	arrParam(0) = "" 'loan_no
	arrParam(1) = "" 'pay_plan_dt
	arrParam(2) = "" 'pay_obj

	If frm1.vspdData1.Maxrows > 0 then
		For i = 1 to frm1.vspdData1.maxRows
			frm1.vspdData1.Row = i
			frm1.vspdData1.Col = C_LOAN_NO
			arrParam(0) = arrParam(0) & frm1.vspdData1.text & chr(11) 
			frm1.vspdData1.Col = C_LOAN_PLAN_DT
			arrParam(1) = arrParam(1) & UniConvDate(frm1.vspdData1.text) & chr(11) 
		Next
	End if
   
	arrRet = window.showModalDialog(iCalledAspName & "?PGM=" & gStrRequestMenuID, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: Yes; status: No;")

	IsOpenPop = False
	
	If arrRet(0,0) = ""  Then
		frm1.txtRePayNO.focus
		Exit Function
	Else
		Call SetRefOpenLoan(arrRet)
	End If

	frm1.vspddata1.focus
'	Call CurFormatNumericOCX()
	Call SetToolbar("1111100100001111")													'��: Developer must customize

'	lgIntFlgMode      = Parent.OPMD_CMODE												'��: Indicates that current mode is Create mode
	lgBlnFlgChgValue = True
End Function

'======================================================================================================
'   Function Name : SetRefOpenLoan
'   Function Desc : ���Ա����� Popup ��� Set
'=======================================================================================================
Function SetRefOpenLoan(Byval arrRet)
	Dim intRtnCnt, strData
	Dim TempRow, I
	DIM X
	Dim strIntClsPlanAmt
	Dim strIntClsPlanLocAmt	

	With frm1
		.vspdData1.focus		
		ggoSpread.Source = .vspdData1
		.vspdData1.ReDraw = False	
	
		TempRow = .vspdData1.MaxRows												'��: ��������� MaxRows

		For I = TempRow To TempRow + Ubound(arrRet, 1) 

			.vspdData1.MaxRows = .vspdData1.MaxRows + 1
			.vspdData1.Row = I + 1				
			.vspdData1.Col = 0
			.vspdData1.Text = ggoSpread.InsertFlag
				
			.vspdData1.Col = C_LOAN_NO        												
			.vspdData1.text = arrRet(I - TempRow, 0)
			.vspdData1.Col = C_LOAN_DT        												
			.vspdData1.text = arrRet(I - TempRow, 16)
			.vspdData1.Col = C_LOAN_DUE_DT         											
			.vspdData1.text = arrRet(I - TempRow, 17)
			.vspdData1.Col = C_LOAN_PLAN_DT
			.vspdData1.text = arrRet(I - TempRow, 2)				
			.vspdData1.Col = C_LOAN_DOCCUR        											
			.vspdData1.text = arrRet(I - TempRow, 7)
			.vspdData1.Col = C_LOAN_XCH_RATE         										
			.vspdData1.text = arrRet(I - TempRow, 24)
			.vspdData1.Col = C_REPAY_PLAN_AMT
			.vspdData1.text = arrRet(I - TempRow, 8)
			.vspdData1.Col = C_REPAY_PLAN_LOC_AMT
			.vspdData1.text = arrRet(I - TempRow, 9)
			.vspdData1.Col = C_REPAY_PLAN_INT_AMT
			.vspdData1.text = arrRet(I - TempRow, 10)
			.vspdData1.Col = C_REPAY_PLAN_INT_LOC_AMT
			.vspdData1.text = arrRet(I - TempRow, 11)
'			.vspdData1.Col = C_REPAY_INT_ACCT_CD
'			.vspdData1.text = arrRet(I - TempRow, 10)
'			.vspdData1.Col = C_REPAY_INT_ACCT_NM
'			.vspdData1.text = arrRet(I - TempRow, 11)
			.vspdData1.Col = C_LOAN_BAL_AMT
			.vspdData1.text = arrRet(I - TempRow, 12)
			.vspdData1.Col = C_LOAN_BAL_LOC_AMT
			.vspdData1.text = arrRet(I - TempRow, 13)
			.vspdData1.Col = C_LOAN_RDP_TOT_AMT
			.vspdData1.text = arrRet(I - TempRow, 20)
			.vspdData1.Col = C_LOAN_RDP_TOT_LOC_AMT
			.vspdData1.text = arrRet(I - TempRow, 21)
			.vspdData1.Col = C_LOAN_INT_TOT_AMT
			.vspdData1.text = arrRet(I - TempRow, 22)
			.vspdData1.Col = C_LOAN_INT_TOT_LOC_AMT
			.vspdData1.text = arrRet(I - TempRow, 23)
			.vspdData1.Col = C_REPAY_PAY_OBJ
			.vspdData1.text = arrRet(I - TempRow, 26)

			If arrRet(I - TempRow,26) = "DI" Then			'���������� 
				Call CommonQueryRs("sum(int_cls_amt),sum(int_cls_loc_amt), xch_rate","f_ln_mon_dfr_int"," loan_no = " & FilterVar(arrRet(0,0), "''", "S") & _
					" and int_pay_plan_dt =  " & FilterVar(UNIConvDate(arrRet(0,2)), "''", "S") & " and CLS_FG = " & FilterVar("Y", "''", "S") & "  group by xch_rate"  ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				strIntClsPlanAmt = Replace(lgf0,chr(11),"")
				strIntClsPlanLocAmt = Replace(lgf1,chr(11),"")
				.vspdData1.Col  = C_REPAY_INT_DFR_AMT
				.vspdData1.Text = UNIFormatNumber(strIntClsPlanAmt,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
				.vspdData1.Col  = C_REPAY_INT_DFR_LOC_AMT
				.vspdData1.Text = UNIFormatNumber(strIntClsPlanLocAmt,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
			Else
				.vspdData1.Col  = C_REPAY_INT_DFR_AMT
				.vspdData1.Text = UNIFormatNumber(0,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
				.vspdData1.Col  = C_REPAY_INT_DFR_LOC_AMT
				.vspdData1.Text = UNIFormatNumber(0,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
			End If				
		Next	
		
		Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData1,-1,-1,C_LOAN_DOCCUR,C_LOAN_XCH_RATE		, "A" ,"I","X","X")
		Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData1,-1,-1,C_LOAN_DOCCUR,C_REPAY_PLAN_AMT		, "A" ,"I","X","X")
		Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData1,-1,-1,C_LOAN_DOCCUR,C_REPAY_PLAN_INT_AMT , "A" ,"I","X","X")
		Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData1,-1,-1,C_LOAN_DOCCUR,C_REPAY_INT_DFR_AMT  , "A" ,"I","X","X")
		Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData1,-1,-1,C_LOAN_DOCCUR,C_LOAN_BAL_AMT		, "A" ,"I","X","X")
		Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData1,-1,-1,C_LOAN_DOCCUR,C_LOAN_RDP_TOT_AMT	, "A" ,"I","X","X")		
		Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData1,-1,-1,C_LOAN_DOCCUR,C_LOAN_INT_TOT_AMT	, "A" ,"I","X","X")		

	End With
	
	Call DoSum()
	Call SetSpreadLock("B")
	Call SetSpreadColor(-1,-1,"D")
	Call CheckIntAmt()
	frm1.txtRePayNO.focus
End Function

'******************************************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'********************************************************************************************************* 
'======================================================================================================
'   Function Name : OpenPopUp()
'   Function Desc : 
'=======================================================================================================
Function  OpenPopUp(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
		
	Select Case iWhere
		Case 1			'�������� �����ڵ��˾� 
		    frm1.vspdData1.Col = C_REPAY_INT_ACCT_CD
		    frm1.vspdData1.Row = frm1.vspdData1.ActiveRow

   			arrParam(0) = "�����ڵ��˾�"							' �˾� ��Ī 
			arrParam(1) = "A_Acct A, A_ACCT_GP B, A_JNL_ACCT_ASSN C"    ' TABLE ��Ī 
			arrParam(2) = ""											' Code Condition
			arrParam(3) = ""											' Name Condition
			arrParam(4) = "A.GP_CD=B.GP_CD AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND A.Acct_cd=C.Acct_CD and C.jnl_cd = " & FilterVar("PI", "''", "S") & "  "          ' Where Condition
			arrParam(4) = arrParam(4) & " and C.trans_type = " & FilterVar("FI002", "''", "S") & " "
			arrParam(5) = "�����ڵ�"								' �����ʵ��� �� ��Ī 

			arrField(0) = "A.Acct_CD"									' Field��(0)
			arrField(1) = "A.Acct_NM"									' Field��(1)
			arrField(2) = "B.GP_CD"										' Field��(2)
			arrField(3) = "B.GP_NM"										' Field��(3)
		 
			arrHeader(0) = "�����ڵ�"								' Header��(0)
			arrHeader(1) = "�����ڵ��"								' Header��(1)
			arrHeader(2) = "�׷��ڵ�"								' Header��(2)
			arrHeader(3) = "�׷��"				
		Case 2			'��ݰ����ڵ��˾� 
		    frm1.vspdData4.Col = C_REPAY_TYPE
		    frm1.vspdData4.Row = frm1.vspdData4.ActiveRow

   			arrParam(0) = "�����ڵ��˾�"							' �˾� ��Ī 
			arrParam(1) = "A_Acct A, A_ACCT_GP B, A_JNL_ACCT_ASSN C"    ' TABLE ��Ī 
			arrParam(2) = ""											' Code Condition
			arrParam(3) = ""											' Name Condition
			arrParam(4) = "A.GP_CD=B.GP_CD AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND A.Acct_cd=C.Acct_CD and C.jnl_cd = " & FilterVar(frm1.vspdData4.Text, "''", "S")         ' Where Condition
			arrParam(4) = arrParam(4) & " and C.trans_type = " & FilterVar("FI002", "''", "S") & " "
			arrParam(5) = "�����ڵ�"								' �����ʵ��� �� ��Ī 

			arrField(0) = "A.Acct_CD"									' Field��(0)
			arrField(1) = "A.Acct_NM"									' Field��(1)
			arrField(2) = "B.GP_CD"										' Field��(2)
			arrField(3) = "B.GP_NM"										' Field��(3)
		 
			arrHeader(0) = "�����ڵ�"								' Header��(0)
			arrHeader(1) = "�����ڵ��"								' Header��(1)
			arrHeader(2) = "�׷��ڵ�"								' Header��(2)
			arrHeader(3) = "�׷��"			
		Case 3			'��ݵ���� �ŷ���ȭ �˾� 
			frm1.vspdData4.Col = C_DOCCUR
		    frm1.vspdData4.Row = frm1.vspdData4.ActiveRow
		
			arrParam(0)  = "�ŷ���ȭ�˾�"
			arrParam(1)  = "B_CURRENCY"
			arrParam(2)  = Trim(frm1.vspdData4.text)
			arrParam(3)  = ""
			arrParam(4)  = ""
			arrParam(5)  = "�ŷ���ȭ"	

			arrField(0)  = "CURRENCY"
			arrField(1)  = "CURRENCY_DESC"    

			arrHeader(0) = "�ŷ���ȭ"
			arrHeader(1) = "�ŷ���ȭ��"
		Case 4			'�δ����� �����ڵ� 
		    frm1.vspdData.Col = C_ACCT_CD
		    frm1.vspdData.Row = frm1.vspdData.ActiveRow
		    		
			arrParam(0) = "�����ڵ��˾�"													' �˾� ��Ī 
			arrParam(1) = "A_Acct A , A_ACCT_GP B , A_JNL_ACCT_ASSN C "							' TABLE ��Ī 
			arrParam(2) = ""  																	' Code Condition
			arrParam(3) = ""																	' Name Cindition
			arrParam(4) = "A.GP_CD=B.GP_CD AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND "							' Where Condition
			arrParam(4) = arrParam(4) & "  A.Acct_cd=C.Acct_CD and C.jnl_cd in (" & FilterVar("BP", "''", "S") & " ," & FilterVar("BC", "''", "S") & " )  AND C.trans_type = " & FilterVar("FI002", "''", "S") & "  " 
			arrParam(5) = "�����ڵ�"														' �����ʵ��� �� ��Ī 

			arrField(0) = "A.Acct_CD"															' Field��(0)
			arrField(1) = "A.Acct_NM"															' Field��(1)
    		arrField(2) = "B.GP_CD"																' Field��(2)
			arrField(3) = "B.GP_NM"																' Field��(3)
			
			arrHeader(0) = "�����ڵ�"														' Header��(0)
			arrHeader(1) = "�����ڵ��"														' Header��(1)
			arrHeader(2) = "�׷��ڵ�"														' Header��(2)
			arrHeader(3) = "�׷��"															' Header��(3)
		Case 6			'��ݵ���� ������� 
		    frm1.vspdData4.Col = C_REPAY_TYPE
		    frm1.vspdData4.Row = frm1.vspdData4.ActiveRow
		    		     
			arrParam(0) = "�������"														' �˾� ��Ī						
			arrParam(1) = " B_MINOR A , B_CONFIGURATION B "
			arrParam(2) = Trim(frm1.vspdData4.text)												' Code Condition
			arrParam(3) = ""																	' Name Cindition
			arrParam(4) = "A.MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND A.MAJOR_CD = B.MAJOR_CD AND A.MINOR_CD = B.MINOR_CD "
			arrParam(4) = arrParam(4) & " AND B.SEQ_NO = 4 AND B.REFERENCE IN (" & FilterVar("FO", "''", "S") & " ," & FilterVar("DP", "''", "S") & " ," & FilterVar("CS", "''", "S") & " ," & FilterVar("CK", "''", "S") & " ) "
			arrParam(5) = "�������"														' TextBox ��Ī 
		
			arrField(0) = "A.MINOR_CD"															' Field��(0)
			arrField(1) = "A.MINOR_NM"															' Field��(1)
				    
			arrHeader(0) = "�������"														' Header��(0)
			arrHeader(1) = "���������"														' Header��(1)		
		Case 9						'��ݵ���� �����˾� 
			frm1.vspdData4.Col = C_BANK_ACCT_NO
		    frm1.vspdData4.Row = frm1.vspdData4.ActiveRow
		    
			arrParam(0) = "���¹�ȣ�˾�"
			arrParam(1) = "F_DPST, B_BANK"				
			arrParam(2) = Trim(frm1.vspdData4.text)
			arrParam(3) = ""
			arrParam(4) = "F_DPST.BANK_CD = B_BANK.BANK_CD "
			arrParam(5) = "���¹�ȣ"			
			
		    arrField(0) = "F_DPST.BANK_ACCT_NO"	
			arrField(1) = "B_BANK.BANK_CD"
			arrField(2) = "B_BANK.BANK_NM"	
		    arrField(3) = "F_DPST.DOC_CUR"	

		    arrHeader(0) = "���¹�ȣ"		
		    arrHeader(1) = "�����ڵ�"	
		    arrHeader(2) = "�����"	
		    arrHeader(3) = "�ŷ���ȭ"	
	End Select
		
	IsOpenPop = True	
	   
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")			
		     		
	IsOpenPop = False
		
	If arrRet(0) = "" Then	    
		Exit Function
	Else
		Call SetPopup(arrRet, iWhere)
	End If
End Function

'================================================================
'	��ȯ��ȣ ���� �˾� 
'================================================================
Function OpenPopupPay()
	Dim arrRet
	Dim arrParam(3)
	Dim iCalledAspName

	iCalledAspName = AskPRAspName("F4255RA2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "F4255RA2", "X")
		IsOpenPop = False
		Exit Function
	End If	

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0,1) = ""  Then
		Exit Function
	Else
	    Call ggoOper.ClearField(Document, "1")										'��: Clear Contents  Field
		Call ggoOper.LockField(Document, "N")										'��: Lock Field

		frm1.txtRePayNO.value = arrRet(0,1)
	End If
	frm1.txtRePayNO.focus

End Function


'======================================================================================================
'   Function Name : SetPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function SetPopup(Byval arrRet,Byval iWhere)
	With frm1
		Select Case iWhere			
			Case 0
				.txtRePayNO.value = arrRet(0)	
			Case 1
				.vspdData1.Col  = C_REPAY_INT_ACCT_CD
				.vspdData1.Text = arrRet(0)
				.vspdData1.Col  = C_REPAY_INT_ACCT_NM
				.vspdData1.Text = arrRet(1)
			Case 2
				.vspdData4.Col  = C_REPAY_ACCT_CD
				.vspdData4.Text = arrRet(0)
				.vspdData4.Col  = C_REPAY_ACCT_NM
				.vspdData4.Text = arrRet(1)
			Case 3
				.vspdData4.Col  = C_DOCCUR
				.vspdData4.Text = arrRet(0)

				Call DocCur_OnChange(.vspdData4.Text,.vspdData4.ActiveRow)	
			Case 4		'�δ����� �����ڵ� 
				.vspdData.Col  = C_ACCT_CD
				.vspdData.Text = arrRet(0)
				.vspdData.Col  = C_ACCT_CD_NM
				.vspdData.Text = arrRet(1)			

				Call vspdData_Change(C_ACCT_CD, frm1.vspdData.ActiveRow)
			Case 6		'��ݵ���� �������	    			  
			    .vspdData4.Col  = C_REPAY_TYPE
			    .vspdData4.Text = arrRet(0)
				.vspdData4.Col  = C_REPAY_TYPE_NM
				.vspdData4.Text = arrRet(1)						

				Call vspdData4_Change(C_REPAY_TYPE, frm1.vspdData4.ActiveRow)
			Case 9
				.vspdData4.Col  = C_BANK_ACCT_NO
				.vspdData4.Text = arrRet(0)
				.vspdData4.Col  = C_BANK_CD		
				.vspdData4.Text  = arrRet(1)				
				.vspdData4.Col  = C_BANK_NM		
				.vspdData4.Text  = arrRet(2)				
				.vspdData4.Col  = C_DOCCUR		
				.vspdData4.Text  = arrRet(3)

				Call DocCur_OnChange(.vspdData4.Text,.vspdData4.ActiveRow)			
		End Select		
	End With
End Function

'======================================================================================================
'	Name : OpenDept
'	Description : 
'=======================================================================================================%>
Function OpenDept(Byval strCode)
	Dim arrRet
	Dim arrParam(3)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = frm1.txtRePayDT.Text
	arrParam(2) = lgUsrIntCd                            ' �ڷ���� Condition  
	arrParam(3) = "F"									' �������� ���� Condition  
	
	arrRet = window.showModalDialog("../../comasp/DeptPopupDtA2.asp", Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		.txtDeptCd.focus
		Exit Function
	Else
		Call SetDept(arrRet)
	End If	
End Function
'------------------------------------------  SetDept()  ------------------------------------------------
'	Name : SetDept()
'	Description : Dept Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetDept(Byval arrRet)
	With frm1
		.txtDeptCd.value    = arrRet(0)		
		.txtDeptNm.value    = arrRet(1)
		.horgChangeId.value = arrRet(2)
		.txtRePayDT.text	= arrRet(3)
		.txtDeptCd.focus
	End With
	
	Call txtDeptCD_OnChange()
End Function 

'======================================================================================================
'	���: Tab Click
'	����: Tab Click�� �ʿ��� ����� �����Ѵ�.
'=======================================================================================================
Function ClickTab1()
	Call changeTabs(TAB1)
	gSelframeFlg = TAB1

	Call MoveJmpClick() 
	Call SetToolbar("1111111100101111") 				 
End Function

Function ClickTab2()
	Call changeTabs(TAB2)	 
	gSelframeFlg = TAB2
	
	Call MoveJmpClick()
	Call SetToolbar("1111100100001111") 
End Function

Function ClickTab3()
	Call changeTabs(TAB3)	 
	gSelframeFlg = TAB3

	Call MoveJmpClick()
	Call SetToolbar("1111111100101111")
End Function

'======================================================================================================
'	���: 
'	����: Tab Click�� �ʿ��� ����� �����Ѵ�.
'=======================================================================================================
Function MoveJmpClick()
	Select Case gSelframeFlg
		Case TAB1, TAB3
			spnArInfo.innerHTML =  "&nbsp;&nbsp;"			
		Case TAB2
			spnArInfo.innerHTML =  "<a href='vbscript:OpenPopupLoan()'>���Ա�����</A>&nbsp;|&nbsp;"
	End Select    
End Function




'#########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'##########################################################################################################
 '******************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 
'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub  Form_Load()
    Call LoadInfTB19029()  
        
	Call ggoOper.FormatField(Document, "A", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)    
        
    Call ggoOper.LockField(Document, "N")
    Call InitSpreadSheet("A")	    															'Setup the Spread sheet
    Call InitSpreadSheet("B")	
    Call InitSpreadSheet("C")	

    Call InitVariables()
    Call SetDefaultVal()
    Call ClickTab1()
	
	gIsTab     = "Y" 
	gTabMaxCnt = 3
	
	frm1.txtRePayNO.focus
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub  Form_QueryUnload(Cancel , UnloadMode )
    
End Sub

'==========================================================================================
'   Event Name : vspdData1_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)
	Call SetPopUpMenuItemInf("0000111111")
    gMouseClickStatus = "SP1C"	'Split �����ڵ� 
    Set gActiveSpdSheet = frm1.vspdData1

	If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData1
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If 
		Exit sub   
    End If
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Dim i  
	
	Call SetPopUpMenuItemInf("1101111111")
	gMouseClickStatus = "SPC" 'Split �����ڵ� 
    Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows = 0 then
	    Exit Sub
	End if
  	
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
    
'    Call SetSpread2Color()
     
'    If Col <> C_ACCT_CD Then  Exit Sub
End Sub

'==========================================================================================
'   Event Name : vspdData4_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData4_Click(ByVal Col, ByVal Row)
	Call SetPopUpMenuItemInf("1101111111")
    gMouseClickStatus = "SP5C"	'Split �����ڵ� 
    Set gActiveSpdSheet = frm1.vspdData4

    If frm1.vspdData4.MaxRows = 0 then
	    Exit Sub
	End if
	
    Dim i    
    	
	If Row <= 0 Then
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
' Event Name :vspdData_MouseDown
' Event Desc :Spread Split �����ڵ� 
'========================================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 and gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'========================================================================================================
' Event Name :vspdData1_MouseDown
' Event Desc :Spread Split �����ڵ� 
'========================================================================================================
Sub vspdData1_MouseDown(Button, Shift, X, Y)
	If Button = 2 and gMouseClickStatus = "SP1C" Then
		gMouseClickStatus = "SP1CR"
	End If
End Sub

'========================================================================================================
' Event Name :vspdData4_MouseDown
' Event Desc :Spread Split �����ڵ� 
'========================================================================================================
Sub vspdData4_MouseDown(Button, Shift, X, Y)
	If Button = 2 and gMouseClickStatus = "SP5C" Then
		gMouseClickStatus = "SP5CR"
	End If
End Sub

'======================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : �󼼳��� �׸����� (��Ƽ)�÷��� �ʺ� �����ϴ� ��� 
'=======================================================================================================
Sub  vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.Source = frm1.vspdData
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'======================================================================================================
'   Event Name : vspdData1_ColWidthChange
'   Event Desc : �󼼳��� �׸����� (��Ƽ)�÷��� �ʺ� �����ϴ� ��� 
'=======================================================================================================
Sub  vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.Source = frm1.vspdData1
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'======================================================================================================
'   Event Name : vspdData1_ColWidthChange
'   Event Desc : �󼼳��� �׸����� (��Ƽ)�÷��� �ʺ� �����ϴ� ��� 
'=======================================================================================================
Sub  vspdData4_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.Source = frm1.vspdData4
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'======================================================================================================
'   Event Name :vspddata_ScriptDragDropBlock
'   Event Desc :
'=======================================================================================================
Sub  vspddata_ScriptDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
	ggoSpread.Source = frm1.vspdData 
	Call ggoSpread.SpreadDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
	Call GetSpreadColumnPos("C")
End Sub

'======================================================================================================
'   Event Name :vspddata_ScriptDragDropBlock
'   Event Desc :
'=======================================================================================================
Sub  vspddata1_ScriptDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
	ggoSpread.Source = frm1.vspdData1 
	Call ggoSpread.SpreadDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
	Call GetSpreadColumnPos("B")
End Sub

'======================================================================================================
'   Event Name :vspddata_ScriptDragDropBlock
'   Event Desc :
'=======================================================================================================
Sub  vspddata4_ScriptDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
	ggoSpread.Source = frm1.vspdData4 
	Call ggoSpread.SpreadDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
	Call GetSpreadColumnPos("A")
End Sub

 '**************************  3.2 HTML Form Element & Object Eventó��  **********************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'********************************************************************************************************* 

 '******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 

'=======================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc :
'=======================================================================================================
Sub  vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If Row <> NewRow And NewRow > 0 Then
        With frm1       
            .vspdData.Row    = NewRow
            .vspdData1.Col   = 1            
            .vspdData.Col    = C_ITEM_SEQ
            .hItemSeq.value  = .vspdData.Text
        End With
        
        frm1.vspdData.Col = 0
        If frm1.vspdData.Text = ggoSpread.DeleteFlag Then
			Exit Sub
        End If
        
        lgCurrRow = NewRow
    End If
End Sub

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : ��ư �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub  vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
    If Row < 0 Then Exit Sub
    
	Select Case Col
        Case C_ACCT_CD_POP
			Call OpenPopUp(4)              
    End Select

    Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,row,row,C_ETC_DOCCUR, C_ITEMAMT,"A" ,"I","X","X")   
End Sub

'======================================================================================================
'   Event Name : vspdData1_ButtonClicked
'   Event Desc : ��ư �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub  vspdData1_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
    If Row < 0 Then Exit Sub
    
	Select Case Col
		Case C_REPAY_INT_ACCT_CD_POP
			Call OpenPopUp(1)              
    End Select
End Sub

'======================================================================================================
'   Event Name : vspdData4_ButtonClicked
'   Event Desc : ��ư �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub  vspdData4_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim IsWhere, lsSelect, IsFrom
	If Row < 0 Then Exit Sub

	Select Case Col
		Case C_REPAY_POP
			Call OpenPopup(6)
		Case C_BANK_ACCT_NO_POP
			frm1.vspdData4.Col = C_REPAY_TYPE
			frm1.vspdData4.Row = Row

			IsFrom   = " B_MINOR  A , B_CONFIGURATION B "
			IsWhere  = " A.MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND A.MAJOR_CD = B.MAJOR_CD AND A.MINOR_CD = B.MINOR_CD "
			IsWhere  = IsWhere & " and A.MINOR_CD = " & FilterVar(frm1.vspdData4.text, "''", "S")  & " And B.seq_no=4 "
			lsSelect = " A.MINOR_CD,A.MINOR_NM, B.REFERENCE "
				   
			Call CommonQueryRs( lsSelect, IsFrom , IsWhere, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 	   	   
			
			If (Trim(lgF0) = "X") Or (Trim(lgF0) = "") Then Exit Sub

			Select Case UCase(Trim(Left(lgF2, Len(lgF2)-1)))
				Case "DP"	' ������ 
					Call OpenPopup(9)
			End Select
		Case C_REPAY_ACCT_CD_POP                
			Call OpenPopup(2)
	    Case C_DOCCUR_POP
            Call OpenPopUp(3)  
    End Select
    
    Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData4,Row,Row,C_DOCCUR, C_REPAY_AMT,"A" ,"I","X","X")         
End Sub

'======================================================================================================
'   Event Name :vspdData1_EditChange
'   Event Desc :
'=======================================================================================================
Sub  vspdData1_EditChange(ByVal Col , ByVal Row )
	With frm1.vspdData1 
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.UpdateRow Row
    End With
End Sub

'======================================================================================================
'   Event Name :vspdData4_EditChange
'   Event Desc :
'=======================================================================================================
Sub  vspdData4_EditChange(ByVal Col , ByVal Row )
	With frm1.vspdData4 
		ggoSpread.Source = frm1.vspdData4
		ggoSpread.UpdateRow Row
    End With
End Sub

'======================================================================================================
'   Event Name :vspdData_EditChange
'   Event Desc :
'=======================================================================================================
Sub  vspdData_EditChange(ByVal Col , ByVal Row )
	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		ggoSpread.UpdateRow Row
    End With
End Sub

'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'=======================================================================================================
Sub  vspdData_Change(ByVal Col, ByVal Row )
    Dim IsWhere, lsSelect, IsFrom
    Dim RetFlag

	lgBlnFlgChgValue = True
	Call CheckMinNumSpread(frm1.vspdData,Col,Row)

	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
    
    With frm1.vspdData    
		.Row = Row
    
		Select Case Col
			Case C_ITEMAMT 
				Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_ETC_DOCCUR,C_ITEMAMT,  "A" ,"X","X")
				Call DoMulti(Row)	
			    Call DoSum()
			Case C_ACCT_CD
				.Col = C_ACCT_CD
				.Row = Row
					
				IsFrom   = " A_Acct A, A_ACCT_GP B, A_JNL_ACCT_ASSN C  "
				IsWhere  = " A.GP_CD=B.GP_CD AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND A.Acct_cd=C.Acct_CD and C.jnl_cd in (" & FilterVar("BC", "''", "S") & " ," & FilterVar("BP", "''", "S") & " ) "
				IsWhere  = IsWhere & " and C.trans_type = " & FilterVar("FI002", "''", "S") & "  and a.acct_cd = " & FilterVar(.text, "''", "S") 
				lsSelect = " distinct a.acct_cd ,a.acct_nm "

				Call CommonQueryRs( lsSelect, IsFrom , IsWhere, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 	   	   

				If (Trim(lgF0) = "X") Or (Trim(lgF0) = "") Then
					If .Text = "" Then Exit Sub
					RetFlag = DisplayMsgBox("110103","X" , "X", "X") 	
					'// MsgBox "%1 �����ڵ� : Permitted Value�� �̻� �ֽ��ϴ�."
					.Text = ""
					.Col = C_ACCT_CD_NM		
					.Text = ""
				    Exit Sub
				Elseif Trim(frm1.vspdData1.Text) <> Trim(Left(lgF1, Len(lgF1)-1)) Then
					.Col = C_ACCT_CD_NM						
				 	.Text =Trim(Left(lgF1, Len(lgF1)-1))
				End If						    
		End Select
	End With		
End Sub

'======================================================================================================
'   Event Name : vspdData1_Change
'   Event Desc :0
'=======================================================================================================
Sub  vspdData1_Change(ByVal Col, ByVal Row )
	DIm TmpPlanAmt,TmpPlanIntAmt
    Dim IsWhere, lsSelect, IsFrom	
    Dim RetFlag
	
	TmpPlanAmt= 0
	TmpPlanIntAmt=0
	
    lgBlnFlgChgValue = True
    Call CheckMinNumSpread(frm1.vspdData1,Col,Row)
    
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.UpdateRow Row
 	
 	WIth frm1.vspdData1   
		.Row = Row
		.Col = 0             
		.Col = C_REPAY_PLAN_AMT
		TmpPlanAmt = UNICDbl(.Text)
		.Col = C_REPAY_PLAN_INT_AMT
		TmpPlanIntAmt = UNICDbl(.Text)

		Select Case Col
			Case C_REPAY_PLAN_AMT
				If (TmpPlanAmt > 0 And TmpPlanIntAmt < 0) Or (TmpPlanAmt < 0 And TmpPlanIntAmt > 0) then
					.Col  = C_LOAN_DOCCUR
					TmpPlanIntAmt= UNIConvNumPCToCompanyByCurrency(TmpPlanIntAmt * (-1),.text,parent.ggAmtOfMoneyNo, "X", "X")
					.Col  = C_REPAY_PLAN_AMT
					.Text = TmpPlanIntAmt
				End If
					
				Call FixDecimalPlaceByCurrency(frm1.vspdData1,Row,C_LOAN_DOCCUR,C_REPAY_PLAN_AMT,  "A" ,"X","X")
				Call DoMulti(Row)	' ApClsLocAmt 
				Call DoSum()
			Case C_REPAY_PLAN_INT_AMT
				If (TmpPlanAmt > 0 And TmpPlanIntAmt < 0) Or (TmpPlanAmt < 0 And TmpPlanIntAmt > 0) Then
					.Col  = C_LOAN_DOCCUR
					TmpPlanIntAmt= UNIConvNumPCToCompanyByCurrency(TmpPlanIntAmt * (-1),.text,parent.ggAmtOfMoneyNo, "X", "X")
					.Col  = C_REPAY_PLAN_INT_AMT
					.Text = TmpPlanIntAmt
				End If					
					
				Call FixDecimalPlaceByCurrency(frm1.vspdData1,Row,C_LOAN_DOCCUR,C_REPAY_PLAN_INT_AMT,  "A" ,"X","X")
				Call DoMulti(Row)	' ApClsLocAmt 
				Call DoSum()
				
				.Col = C_REPAY_PLAN_INT_AMT
				.Row = Row
				
				If ABS(UNICDbl(.Text)) <> 0  Then
					ggoSpread.SpreadUnLock  C_REPAY_INT_ACCT_CD_POP, Row, C_REPAY_INT_ACCT_CD_POP, Row
					ggoSpread.SpreadUnLock  C_REPAY_INT_ACCT_CD    , Row, C_REPAY_INT_ACCT_CD    , Row
					ggoSpread.SSSetRequired C_REPAY_INT_ACCT_CD    , Row, Row	
				Elseif ABS(UNICDbl(.Text)) = 0 Then
					.Col  = C_REPAY_INT_ACCT_CD
					.Text = ""
					ggoSpread.SpreadLock     C_REPAY_INT_ACCT_CD_POP, Row, C_REPAY_INT_ACCT_CD_POP, Row
					ggoSpread.SpreadLock     C_REPAY_INT_ACCT_CD    , Row, C_REPAY_INT_ACCT_NM , Row
					ggoSpread.SSSetProtected C_REPAY_INT_ACCT_CD    , Row, Row
				End If
			Case C_REPAY_INT_ACCT_CD
				.Col = C_REPAY_INT_ACCT_CD
				.Row = Row
				
				IsFrom   = " A_Acct A, A_ACCT_GP B, A_JNL_ACCT_ASSN C  "
				IsWhere  = " A.GP_CD=B.GP_CD AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND A.Acct_cd=C.Acct_CD and C.jnl_cd = " & FilterVar("PI", "''", "S") & "  "
				IsWhere  = IsWhere & " and C.trans_type = " & FilterVar("FI002", "''", "S") & "  and a.acct_cd = " & FilterVar(.text, "''", "S") 
				lsSelect = " a.acct_cd ,a.acct_nm "
				   
				Call CommonQueryRs( lsSelect, IsFrom , IsWhere, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 	   	   
				   
				If (Trim(lgF0) = "X") Or (Trim(lgF0) = "") Then
					If .Text = "" Then Exit Sub
					RetFlag = DisplayMsgBox("110103","X" , "X", "X") 	
					'// MsgBox "%1 �����ڵ� : Permitted Value�� �̻� �ֽ��ϴ�."
					.text = ""
					.Col = C_REPAY_INT_ACCT_NM		
					.text = ""
				    Exit Sub
				Elseif Trim(frm1.vspdData1.Text) <> Trim(Left(lgF1, Len(lgF1)-1)) Then
					.Col = C_REPAY_INT_ACCT_NM				
				 	.Text =Trim(Left(lgF1, Len(lgF1)-1))
				End If
		End Select
	End With
End Sub

'======================================================================================================
'   Event Name : vspdData4_Change
'   Event Desc :
'=======================================================================================================
Sub  vspdData4_Change(ByVal Col, ByVal Row )
    Dim IsWhere, lsSelect, IsFrom
    Dim RetFlag,tmpJnlCd,tmpAcctcd
    
    lgBlnFlgChgValue = True
    
    Call CheckMinNumSpread(frm1.vspdData4,Col,Row)
    
    ggoSpread.Source = frm1.vspdData4
    ggoSpread.UpdateRow Row
    
    With frm1.vspdData4
		.Row = Row
    
		Select Case Col
			Case C_REPAY_AMT
				Call FixDecimalPlaceByCurrency(frm1.vspdData4,Row,C_DOCCUR,C_REPAY_AMT,  "A" ,"X","X")
				Call DoMulti(Row)
				Call DoSum()  
			Case C_DOCCUR
				.Col  = C_DOCCUR
				Call DocCur_OnChange(.Text,Row)	
			Case C_REPAY_TYPE		
				.Col = C_REPAY_TYPE
				.Row = Row
				
				IsFrom   = " B_MINOR  A , B_CONFIGURATION B "
				IsWhere  = " A.MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND A.MAJOR_CD = B.MAJOR_CD AND A.MINOR_CD = B.MINOR_CD "
				IsWhere  = IsWhere & " and A.MINOR_CD = " & FilterVar(.Text, "''", "S")  & " And B.seq_no=4 "
				lsSelect = " A.MINOR_CD,A.MINOR_NM, B.REFERENCE "
				   
				Call CommonQueryRs( lsSelect, IsFrom , IsWhere, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 	   	   
				   
   				 .Col = C_REPAY_TYPE_NM

				If (Trim(lgF0) = "X") Or (Trim(lgF0) = "") Then
					.Col = C_REPAY_TYPE
					If .text = "" Then Exit Sub
					RetFlag = DisplayMsgBox("141140","X" , "X", "X") 	
					'// MsgBox "%1 �ش� ��������� �����ϴ�."
					.Text = ""
					.Col = C_REPAY_TYPE_NM
					.Text = ""
				    Exit Sub
				Elseif Trim(.Text) <> Trim(Left(lgF1, Len(lgF1)-1)) Then
				 	.Text =Trim(Left(lgF1, Len(lgF1)-1))
				End If
					
				.ReDraw = False
				ggoSpread.source = frm1.vspdData4    
				.Col = C_REPAY_TYPE   

				Select Case UCase(Trim(Left(lgF2, Len(lgF2)-1)))
				 	Case "CS" , "CK"  '���� 
				 		ggoSpread.SSSetProtected C_BANK_ACCT_NO     , Row, Row	
				 		ggoSpread.SSSetProtected C_BANK_ACCT_NO_POP , Row, Row
				 		ggoSpread.SSSetRequired	 C_REPAY_ACCT_CD    , Row, Row	
				 		ggoSpread.SpreadUnLock	 C_REPAY_ACCT_CD_POP, Row, C_REPAY_ACCT_CD_POP, Row
				 		ggoSpread.SpreadUnLock	 C_DOCCUR           , Row, C_DOCCUR           , Row	
				 		ggoSpread.SpreadUnLock	 C_DOCCUR_POP       , Row, C_DOCCUR_POP       , Row	
				 		ggoSpread.SSSetRequired	 C_DOCCUR           , Row, Row	
				 		ggoSpread.SpreadLock	 C_XCH_RATE         , Row, C_XCH_RATE         ,	Row  
				 	Case "DP" '������ 
				 		ggoSpread.SSSetRequired	 C_BANK_ACCT_NO     , Row, Row	
				 		ggoSpread.SpreadUnLock	 C_BANK_ACCT_NO_POP , Row, C_BANK_ACCT_NO_POP ,	Row	
				 		ggoSpread.SSSetProtected C_DOCCUR           , Row, Row
				 		ggoSpread.SpreadLock	 C_DOCCUR_POP       , Row, C_DOCCUR_POP       , Row	
				 		ggoSpread.SSSetProtected C_XCH_RATE         , Row, Row   
				 		ggoSpread.SSSetRequired	 C_REPAY_ACCT_CD    , Row, Row	
				 		ggoSpread.SpreadUnLock	 C_REPAY_ACCT_CD_POP, Row, C_REPAY_ACCT_CD_POP, Row
				 End Select	
	
				 .ReDraw = True
				 Call ClearRow(Row)
				 Call DoSum()
			Case C_REPAY_ACCT_CD
				.Row = Row
				.Col = C_REPAY_TYPE
				tmpJnlCd = .Text
				.Col = C_REPAY_ACCT_CD
				tmpAcctcd = .Text
				
				IsFrom   = " A_Acct A, A_ACCT_GP B, A_JNL_ACCT_ASSN C  "
				IsWhere  = " A.GP_CD=B.GP_CD AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND A.Acct_cd=C.Acct_CD and C.jnl_cd = " & FilterVar(tmpJnlCd, "''", "S")
				IsWhere  = IsWhere & " and C.trans_type = " & FilterVar("FI002", "''", "S") & "  and a.acct_cd = " & FilterVar(tmpAcctcd, "''", "S") 
				lsSelect = " a.acct_cd ,a.acct_nm "
				   
				Call CommonQueryRs( lsSelect, IsFrom , IsWhere, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 	   	   
				   
				If (Trim(lgF0) = "X") Or (Trim(lgF0) = "") Then
					If .Text = "" Then Exit Sub
					RetFlag = DisplayMsgBox("110103","X" , "X", "X") 	
					'// MsgBox "%1 �����ڵ� : Permitted Value�� �̻� �ֽ��ϴ�."
					.Text = ""
					.Col = C_REPAY_ACCT_NM		
					.Text = ""
				    Exit Sub
				Elseif Trim(frm1.vspdData1.Text) <> Trim(Left(lgF1, Len(lgF1)-1)) Then
					.Col = C_REPAY_ACCT_NM				
				 	.Text =Trim(Left(lgF1, Len(lgF1)-1))
				End If								 
		End Select
	End With		
End Sub

'======================================================================================================
'   Event Name :vspdData_EditMode
'   Event Desc :
'=======================================================================================================
Sub vspdData_EditMode(ByVal Col, ByVal Row, ByVal Mode, ByVal ChangeMade)
    Select Case Col
        Case C_ITEMAMT
            Call EditModeCheck(frm1.vspdData, Row,C_ETC_DOCCUR,C_ITEMAMT, "A" ,"I", Mode, "X", "X")
    End Select
End Sub

'======================================================================================================
'   Event Name :vspdData1_EditMode
'   Event Desc :
'=======================================================================================================
Sub vspdData1_EditMode(ByVal Col, ByVal Row, ByVal Mode, ByVal ChangeMade)
    Select Case Col
        Case C_REPAY_PLAN_AMT
            Call EditModeCheck(frm1.vspdData1, Row,C_LOAN_DOCCUR,C_REPAY_PLAN_AMT, "A" ,"I", Mode, "X", "X")
    End Select
End Sub

'======================================================================================================
'   Event Name :vspdData4_EditMode
'   Event Desc :
'=======================================================================================================
Sub vspdData4_EditMode(ByVal Col, ByVal Row, ByVal Mode, ByVal ChangeMade)
    Select Case Col
        Case C_REPAY_AMT
            Call EditModeCheck(frm1.vspdData4, Row,C_DOCCUR,C_REPAY_AMT, "A" ,"I", Mode, "X", "X")
    End Select
End Sub
'======================================================================================================
'   Event Name :vspdData1_DblClick
'   Event Desc :
'=======================================================================================================
Sub  vspdData_DblClick( ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    
End Sub

'======================================================================================================
'   Event Name :vspdData1_KeyPress
'   Event Desc :
'=======================================================================================================
Sub  vspdData1_KeyPress(KeyAscii)
     
End Sub

'======================================================================================================
'   Event Name : vspdData1_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
Sub  vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
End Sub

'======================================================================================================
'   Event Name : txtDeptCd_Change()
'   Event Desc : 
'=======================================================================================================
Sub txtDeptCD_OnChange()
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii, jj

    lgBlnFlgChgValue = True

	If TRim(frm1.txtDeptCd.value) <>"" Then
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtRePayDT.Text, gDateFormat,""), "''", "S") & "))"			
		
	
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
	
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
			frm1.txtDeptCd.focus
		Else 
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)
					
			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))			
				frm1.hOrgChangeId.value = Trim(arrVal2(2))
			Next	
		End If
	End If		
End Sub



'#########################################################################################################
'												4. Common Function�� 
'	���: Common Function
'	����: ȯ��ó���Լ�, VAT ó�� �Լ� 
'######################################################################################################### 


 '#########################################################################################################
'												5. Interface�� 
'	���: Interface
'	����: ������ Toolbar�� ���� ó���� ���Ѵ�. 
'	      Toolbar�� ��ġ������� ����ϴ� ������ �Ѵ�. 
'	<< ���뺯�� ���� �κ� >>
' 	���뺯�� : Global Variables�� �ƴ����� ������ Sub�� Function���� ���� ����ϴ� ������ �������� 
'				�����ϵ��� �Ѵ�.
' 	1. ������Ʈ���� Call�ϴ� ���� 
'    	   ADF (ADS, ADC, ADF�� �״�� ���)
'    	   - ADF�� Set�ϰ� ����� �� �ٷ� Nothing �ϵ��� �Ѵ�.
' 	2. ������Ʈ�ѿ��� Return�� ���� �޴� ���� 
'    		strRetMsg
'######################################################################################################### 
 '*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'********************************************************************************************************* 

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function  FncQuery() 
    Dim IntRetCD 
    Dim Spr1, Spr2, Spr3, Spr4
    
    FncQuery = False
    Err.Clear
    
    ggoSpread.Source = frm1.vspdData1	:    Spr1 = ggoSpread.SSCheckChange
	ggoSpread.Source = frm1.vspdData	:    Spr2 = ggoSpread.SSCheckChange    
'    ggoSpread.Source = frm1.vspdData3	:    Spr3 = ggoSpread.SSCheckChange    
    
    If lgBlnFlgChgValue = True Or Spr1 = True Or Spr2 = True Or Spr3 = True Then
		If DisplayMsgBox("900013", parent.VB_YES_NO,"x","x") = vbNO Then	   
			Exit Function
		End If
    End If        
    															
    If Not chkField(Document, "1") Then
		Exit Function
    End If

	ggoSpread.Source = frm1.vspdData	:	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData1	:	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData4	:	ggoSpread.ClearSpreadData

    Call InitVariables()
    Call ggoOper.ClearField(Document, "2")
    
    If DbQuery = False Then
        Exit Function
    End If
              
    FncQuery = True              
 
	Set gActiveElement = document.ActiveElement	   
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function  FncNew()              
    Dim Spr1, Spr2, Spr3, Spr4
    
    FncNew = False
    
    ggoSpread.Source = frm1.vspdData1	:    Spr1 = ggoSpread.SSCheckChange
	ggoSpread.Source = frm1.vspdData	:    Spr2 = ggoSpread.SSCheckChange    
    
    If lgBlnFlgChgValue = True Or Spr1 = True Or Spr2 = True Or Spr3 = True then
		If DisplayMsgBox("900015", parent.VB_YES_NO,"X","X") = vbNO Then	   
			Exit Function
		End If
    End If    
    
    Call SetToolbar("1111111100001111") 
    Call ggoOper.ClearField(Document, "A")  
    Call ggoOper.LockField(Document, "N")
        
    ggoSpread.Source = frm1.vspdData	:	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData1	:	ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData4	:	ggoSpread.ClearSpreadData
    
    Call InitVariables()    
    Call ClickTab1()   
    Call SetDefaultVal()
        
    FncNew = True
  
	Set gActiveElement = document.ActiveElement	 
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function  FncDelete()
    FncDelete = False
    
    Dim liRet
        
    Err.Clear
    On Error Resume Next
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then
		liRet = DisplayMsgBox("900002","X","X","X")                                       
		Exit Function
    End If
    
	If DisplayMsgBox("900003", parent.VB_YES_NO,"X","X") = vbNo Then
		Exit Function
	End If
    
    If DbDelete = False Then
       Exit Function
    End If       
        
    FncDelete = True

	Set gActiveElement = document.ActiveElement	    
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function  FncSave() 	                
    Dim Spr1, Spr2, Spr3, Spr4
	Dim RetFlag
	
	FncSave = False                                                         
    
    Err.Clear 
    
    ggoSpread.Source = frm1.vspdData1	:    Spr1 = ggoSpread.SSCheckChange
	ggoSpread.Source = frm1.vspdData	:    Spr2 = ggoSpread.SSCheckChange    
    ggoSpread.Source = frm1.vspdData4	:    Spr4 = ggoSpread.SSCheckChange   
    
	If Not chkField(Document, "2") Then
		Exit Function
    End If    

    If lgBlnFlgChgValue = False And Spr1 = False And Spr2 = False And Spr4 = False Then
		DisplayMsgBox "900001","X","X","X"   
		Exit Function
    End If                                                                      
    
	ggoSpread.Source = frm1.vspdData1	:	 Spr1 = ggoSpread.SSDefaultCheck
    ggoSpread.Source = frm1.vspdData	:    Spr2 = ggoSpread.SSDefaultCheck
    ggoSpread.Source = frm1.vspdData4	:    Spr4 = ggoSpread.SSDefaultCheck    
   
    If Spr1 = False Or Spr2 = False Or Spr4 = False Then       
		Exit Function
    End If          
        
    Call DbSave()				                                             '��: Save db data
       
    FncSave = True                                                                '��:                                        
	
	Set gActiveElement = document.ActiveElement	
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function  FncCopy() 
	Select Case gSelframeFlg
	    Case TAB1       
	        If frm1.vspdData4.MaxRows < 1 Then Exit Function
	        frm1.vspdData4.ReDraw = False
	        ggoSpread.Source = frm1.vspdData4	
	        ggoSpread.CopyRow        
		    frm1.vspdData4.ReDraw = True

			Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData4,frm1.vspdData4.ActiveRow,frm1.vspdData4.ActiveRow,C_DOCCUR,C_REPAY_AMT,"A" ,"I","X","X")
			Call SetSpreadColor(frm1.vspdData4.ActiveRow,  frm1.vspdData4.ActiveRow,"D")
			Call MaxSpreadVal(frm1.vspdData4, C_REPAY_MEAN_SEQ, frm1.vspdData4.ActiveRow)
			Call vspdData4_Change(C_REPAY_ACCT_CD, frm1.vspddata4.activerow)

	    Case TAB3
	        If frm1.vspdData.MaxRows < 1 Then Exit Function
	         
	        frm1.vspdData.ReDraw = False
	        ggoSpread.Source = frm1.vspdData	
	        ggoSpread.CopyRow        

			Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData,frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow,C_ETC_DOCCUR,C_ITEMAMT,"A" ,"I","X","X")
			Call SetSpreadColor(frm1.vspdData.ActiveRow,  frm1.vspdData.ActiveRow,"D")
			Call MaxSpreadVal(frm1.vspdData, C_Item_Seq, frm1.vspdData.ActiveRow)
			Call vspdData_Change(C_ACCT_CD, frm1.vspddata.activerow)
		         
		    frm1.vspdData.ReDraw = True
	End Select

	Call Dosum()
	Set gActiveElement = document.ActiveElement	
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function  FncCancel() 
	Select Case gSelframeFlg
	    Case TAB1
	    	With frm1.vspdData4 
				If .Maxrows < 1 Then Exit Function
				.Row = .ActiveRow
				.Col = 0            
				ggoSpread.Source = frm1.vspdData4
				ggoSpread.EditUndo
				
				Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData4,.ActiveRow,.ActiveRow,C_DOCCUR,C_REPAY_AMT,"A" ,"I","X","X")
				Call DoSum()

				If .Maxrows < 1 Then Exit Function			
				.Row = .ActiveRow
				.Col = 1                        
			End With
	    Case TAB2
			With frm1.vspdData1 
				If .Maxrows < 1 Then Exit Function
				.Row = .ActiveRow
				.Col = 0   

				ggoSpread.Source = frm1.vspdData1
				ggoSpread.EditUndo
'			C_REPAY_INT_DFR_AMT      = 9
'			C_LOAN_RDP_TOT_AMT		 = 19
'			C_LOAN_INT_TOT_AMT		 = 21

				Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData1,.Row,.Row,C_LOAN_DOCCUR,C_LOAN_BAL_AMT      ,"A" ,"I","X","X")
				Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData1,.Row,.Row,C_LOAN_DOCCUR,C_REPAY_PLAN_AMT    ,"A" ,"I","X","X")
				Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData1,.Row,.Row,C_LOAN_DOCCUR,C_REPAY_PLAN_INT_AMT,"A" ,"I","X","X")

				Call DoSum()

				If .Maxrows < 1 Then Exit Function			
				.Row = .ActiveRow
				.Col = 1               
			End With
	    Case TAB3
			With frm1.vspdData
				if .Maxrows < 1 Then Exit Function
				.Row = .ActiveRow
				.Col = 0      
				If  .Text = ggoSpread.InsertFlag Then
					.Col = C_ACCT_CD
					If Len(Trim( .Text)) > 0 Then  
						.Col = C_ITEM_SEQ
						Call DeleteHSheet( .Text)
					End If
				End If      
				ggoSpread.Source = frm1.vspdData
				ggoSpread.EditUndo
				
				Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,.ActiveRow,.ActiveRow,C_ETC_DOCCUR,C_ITEMAMT,"A" ,"I","X","X")    
				Call DoSum()
			End With
	End Select

	Set gActiveElement = document.ActiveElement	
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function  FncInsertRow(ByVal pvRowCnt) 
	Dim liTemp
	Dim imRow
	Dim ii
	Dim iCurRowPos
	
	FncInsertRow = False	

	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = Cint(pvRowCnt)
	Else
	    imRow = AskSpdSheetAddRowCount()
    
		If imRow = "" Then
		    Exit Function
		End If
	End If		 

	Select Case gSelframeFlg
		Case TAB1  
			With frm1.vspdData4   
				iCurRowPos = .ActiveRow	
				.ReDraw = False		    
				ggoSpread.Source = frm1.vspdData4
				ggoSpread.InsertRow ,imRow

				For ii = .ActiveRow To  .ActiveRow + imRow - 1
					Call MaxSpreadVal(frm1.vspdData4, C_REPAY_MEAN_SEQ , ii)
				Next        

				.Col = 1																	' �÷��� ���� ��ġ�� �̵� 
				.Row = ii - 1
				.Action = 0		
				.ReDraw = True

				Call SetSpreadColor(iCurRowPos + 1, iCurRowPos + imRow,"D")   
			    Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData4,iCurRowPos + 1,iCurRowPos + imRow,"X",C_REPAY_AMT,"A" ,"I","X","X")
			End With   
		Case TAB3     
			With frm1.vspdData         
				iCurRowPos = .ActiveRow	
				.ReDraw = False		    
				ggoSpread.Source = frm1.vspdData
				ggoSpread.InsertRow ,imRow

				For ii = .ActiveRow To  .ActiveRow + imRow - 1
					Call MaxSpreadVal(frm1.vspdData, C_ITEM_SEQ , ii)
				Next        

				.Col = 1																	' �÷��� ���� ��ġ�� �̵� 
				.Row = ii - 1
				.Action = 0		
				.ReDraw = True

				Call SetSpreadColor(iCurRowPos + 1, iCurRowPos + imRow,"D")        
				Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData,iCurRowPos + 1,iCurRowPos + imRow,"X",C_ItemAmt,"A" ,"I","X","X")						  

				If .MaxRows = 1 then	         
					.Row = .ActiveRow
					.Col = 1
					.Text = .MaxRows
				Else
					.Row = .ActiveRow - 1
					.Col = 1
					liTemp = CInt(.Text)+1
					.Row = .ActiveRow
					.Col = 1
					.Text = liTemp
				End if   
			End With 						
	End Select

    If Err.number = 0 Then
		FncInsertRow = True																	'��: Processing is OK
    End If  		

	Set gActiveElement = document.ActiveElement	 		
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function  FncDeleteRow() 
	Select Case gSelframeFlg
		Case TAB1               
		    frm1.vspdData4.ReDraw = False
		    ggoSpread.Source = frm1.vspdData4	
		    ggoSpread.DeleteRow    		
		    frm1.vspdData4.ReDraw = True	        
		Case TAB2            
		    frm1.vspdData1.ReDraw = False
		    ggoSpread.Source = frm1.vspdData1	
		    ggoSpread.DeleteRow   		
		    frm1.vspdData1.ReDraw = True
		Case TAB3            
		    frm1.vspdData.ReDraw = False
		    ggoSpread.Source = frm1.vspdData
			ggoSpread.DeleteRow    		
		    frm1.vspdData.ReDraw = True
  End Select
  
  Call DoSum()
  
  Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function  FncPrint() 
    On Error Resume Next                                                                 '��: If process fails
    Err.Clear                                                                            '��: Clear error status

    FncPrint = False                                                                     '��: Processing is NG

	Call Parent.FncPrint()                                                               '��: Protect system from crashing

    If Err.number = 0 Then
		FncPrint = True                                                                   '��: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function  FncPrev() 
    On Error Resume Next                                                                 '��: If process fails
    Err.Clear                                                                            '��: Clear error status

    FncPrev = False                                                                      '��: Processing is NG

    If Err.number = 0 Then
		FncPrev = True                                                                    '��: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement                                              
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function  FncNext() 
    On Error Resume Next                                                                 '��: If process fails
    Err.Clear                                                                            '��: Clear error status

    FncNext = False                                                                      '��: Processing is NG

    If Err.number = 0 Then
		FncNext = True                                                                    '��: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement                                              
End Function

'=======================================================================================================
' Function Name : FncFind
' Function Desc : ȭ�� �Ӽ�, Tab���� 
'========================================================================================================
Function  FncFind() 
    On Error Resume Next                                                                 '��: If process fails
    Err.Clear                                                                            '��: Clear error status

    FncFind = False                                                                      '��: Processing is NG

	Call Parent.FncFind(Parent.C_MULTI, True)

    If Err.number = 0 Then
		FncFind = True                                                                    '��: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement                        
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function  FncExcel() 
    On Error Resume Next                                                                 '��: If process fails
    Err.Clear                                                                            '��: Clear error status

    FncExcel = False                                                                     '��: Processing is NG

	Call Parent.FncExport(Parent.C_MULTI)

    If Err.number = 0 Then
		FncExcel = True                                                                   '��: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)    
End Function

'=======================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function  FncExit()
	Dim Spr1, Spr2, Spr3, Spr4
	
	FncExit = False
   
	ggoSpread.Source = frm1.vspdData1	:   Spr1 = ggoSpread.SSCheckChange
	ggoSpread.Source = frm1.vspdData		:   Spr2 = ggoSpread.SSCheckChange
    
	If lgBlnFlgChgValue = True or Spr1 = True or Spr2 = True or Spr3 = True Then
		If DisplayMsgBox("900016", parent.VB_YES_NO,"X","X") = vbNo Then		
			Exit Function
		End If
	End If
        
	FncExit = True   
	
   	Set gActiveElement = document.ActiveElement   
End Function

'*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'********************************************************************************************************* 

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery() 
	DbQuery = False
	
	On Error Resume Next
	Err.Clear

	Call LayerShowHide(1)
	
	Dim strVal
		
	With frm1
		strVal = BIZ_PGM_QUERY_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&txtRePayNO=" & Trim(.txtRePayNO.value)				'��ȸ ���� ����Ÿ 
	End With
	
	Call RunMyBizASP(MyBizASP, strVal)
	  
	DbQuery = True
	Set gActiveElement = document.ActiveElement  
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()
	 Dim Row
	 Dim RepayPayObj

     With frm1
     	Call SetSpreadLock("A")
		
		For Row = 0 to .vspdData4.MaxRows
			ggoSpread.source = frm1.vspdData4    
			.vspdData4.row= Row
			.vspdData4.Col =C_REPAY_TYPE
			ggoSpread.SpreadLock	C_REPAY_ACCT_CD		, Row , C_REPAY_ACCT_CD		, Row  
			ggoSpread.SpreadLock	C_REPAY_ACCT_CD_POP , Row , C_REPAY_ACCT_CD_POP , Row  
		Next 

		Call SetSpreadLock("B")
		
		For Row = 1 To .vspdData1.MaxRows
			ggoSpread.source = frm1.vspdData1
			.vspdData1.Row = Row
			
			.vspdData1.Col = C_REPAY_PLAN_AMT
			If Trim(.vspdData1.Text) = "" Or Trim(.vspdData1.Text) = "0" Or UNICdbl(.vspdData1.Text) = 0 Then			
				.vspddata1.Text = "0"
				ggoSpread.SpreadLock     C_REPAY_PLAN_AMT    , Row, C_REPAY_PLAN_AMT , Row
				ggoSpread.SSSetProtected C_REPAY_PLAN_AMT    , Row, Row					
			End If

			.vspdData1.Col = C_REPAY_PAY_OBJ
			RepayPayObj = Trim(.vspdData1.Text)
			
			.vspdData1.Col = C_REPAY_PLAN_INT_AMT
				
			If RepayPayObj = "DI" AND UniCdbl(.vspdData1.Text) > 0 Then
				ggoSpread.SpreadUnLock  C_REPAY_INT_ACCT_CD_POP, Row, C_REPAY_INT_ACCT_CD_POP, Row
				ggoSpread.SpreadUnLock  C_REPAY_INT_ACCT_CD    , Row, C_REPAY_INT_ACCT_CD    , Row
				ggoSpread.SSSetRequired C_REPAY_INT_ACCT_CD    , Row, Row	
			Else
				.vspddata1.Col  = C_REPAY_INT_ACCT_CD
				.vspddata1.Text = ""
				ggoSpread.SpreadLock     C_REPAY_INT_ACCT_CD_POP, Row, C_REPAY_INT_ACCT_CD_POP, Row
				ggoSpread.SpreadLock     C_REPAY_INT_ACCT_CD    , Row, C_REPAY_INT_ACCT_NM , Row
				ggoSpread.SpreadLock     C_REPAY_PLAN_INT_AMT   , Row, C_REPAY_INT_ACCT_NM , Row				
				ggoSpread.SSSetProtected C_REPAY_PLAN_INT_AMT   , Row, Row				
				ggoSpread.SSSetProtected C_REPAY_INT_ACCT_CD    , Row, Row
			End If
			
		Next			

'		Call SetSpreadColor()

        '-----------------------
        'Reset variables area
        '-----------------------
        If .vspdData.MaxRows > 0 Then
            Call SetSpreadColor(-1, -1,"D")
            
            ggoSpread.source = .vspdData       
			.vspdData.ReDraw = False         

			ggoSpread.SpreadLock C_ACCT_CD	   , -1 , C_ACCT_CD
		    ggoSpread.SpreadLock C_ACCT_CD_POP , -1 , C_ACCT_CD_POP
		    
			.vspdData.ReDraw = True
        End If
    End With    

'	frm1.txtRePayNO.focus 
'	frm1.vspdData4.focus 

    lgIntFlgMode = parent.OPMD_UMODE												'Indicates that current mode is Update mode
    
    
	Call ClickTab1()
     
    lgBlnFlgChgValue = False
End Function

'=======================================================================================================
'   Function Name : FindExchRate
'   Function Desc : 1.��¥, Row�� �Է¹޾� ��¥�� �ش��ϴ� ȯ�������� �о�´�.
'=======================================================================================================
Function FindExchRate(Byval strDate, Byval FromCurrency,Byval Row )
	Dim strSelect, strFrom, strWhere
	Dim arrTemp
	Dim strExchFg
	Dim strExchRate
	Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6	

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
		strWhere	= strWhere & " And to_currency   =  " & FilterVar(parent.gCurrency , "''", "S") & ""
		strWhere	= strWhere & " And apprl_yrmnth  =  " & FilterVar(strDate , "''", "S") & ""

		If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then				
			arrTemp = Split(lgF0, chr(11))
			strExchRate =  arrTemp(0)
			frm1.vspdData4.row  = Row
			frm1.vspdData4.Col  = C_XCH_RATE
			frm1.vspdData4.Text = UNIConvNumPCToCompanyByCurrency(strExchRate, parent.gCurrency, parent.ggExchRateNo, parent.gLocRndPolicyNo, "X")
		Else
			IntRetCD = DisplayMsgBox("121600", "X", "X", "X")
		End If
	Else					' Floating Exchange Rate
		strSelect	= "top 1 std_rate"
		strFrom		= "b_daily_exchange_rate (noLock) "
		strWhere	= "from_currency =  " & FilterVar(FromCurrency , "''", "S") & ""
		strWhere	= strWhere & " And to_currency   =  " & FilterVar(parent.gCurrency , "''", "S") & ""
		strWhere	= strWhere & " And apprl_dt  <= convert(char(21), " & FilterVar(strDate, "''", "S") & ", 20) order by apprl_dt desc"

		If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then				
			arrTemp = Split(lgF0, chr(11))
			strExchRate =  arrTemp(0)
			frm1.vspdData4.row  = Row
			frm1.vspdData4.Col  = C_XCH_RATE
			frm1.vspdData4.Text = UNIConvNumPCToCompanyByCurrency(strExchRate, parent.gCurrency, parent.ggExchRateNo, parent.gLocRndPolicyNo, "X")
		Else
			IntRetCD = DisplayMsgBox("121500", "X", "X", "X")
		End If
	End If
End Function    

'=======================================================================================================
' Function Name : Spread�� ���� Return
' Function Desc : 
'=======================================================================================================
Sub GetSpread1()
	Dim IGrpCnt
    Dim strVal, iColSep, iRowSep
    Dim IRow
    
    '���Ա� ���� 
    
    IGrpCnt = 1
	strVal	= ""  
	iColSep = Parent.gColSep
	iRowSep = Parent.gRowSep
	
	With frm1.vspdData1
		For IRow = 1 To .MaxRows
			.Row = IRow
			.Col = 0
			  		  
			Select Case .Text
				Case ggoSpread.DeleteFlag                                                                '��: Update�߰� 

				Case Else
					strVal = strVal & "C" & iColSep & IRow & iColSep   
					.Col = C_LOAN_NO  
					strVal = strVal & Trim(.text) & iColSep
					.Col = C_LOAN_PLAN_DT  
					strval = strval & UNIConvDate(Trim(.text)) & iColSep 
					.Col = C_REPAY_PLAN_AMT   
					strVal = strVal & Trim(UNIConvNum(.Text,0))  & iColSep 
					.Col = C_REPAY_PLAN_LOC_AMT  
					strVal = strVal & Trim(UNIConvNum(.Text,0))  & iColSep 
					.Col = C_REPAY_PLAN_INT_AMT   
					strVal = strVal & Trim(UNIConvNum(.Text,0))  & iColSep 
					.Col = C_REPAY_PLAN_INT_LOC_AMT  
					strVal = strVal & Trim(UNIConvNum(.Text,0))  & iColSep 					
					.Col = C_REPAY_INT_ACCT_CD
					strVal = strVal & Trim(.text) & iColSep
					.Col = C_REPAY_ITEM_DESC 
					strVal = strVal & Trim(.text) & iRowSep 
					IGrpCnt = IGrpCnt + 1 
					
			End Select
		 Next		
	End With

	frm1.txtMaxRows1.value = IGrpCnt-1										'Spread Sheet�� ����� �ִ밹�� 
	frm1.txtSpread1.value  = strVal
End Sub

'=======================================================================================================
' Function Name : Spread�� ���� Return
' Function Desc : 
'=======================================================================================================
Sub GetSpread4()
    Dim IGrpCnt
    Dim strVal, iColSep, iRowSep
    Dim IRow
        
    '��ݵ�� 
    IGrpCnt = 1
	strVal  = ""
	iColSep = Parent.gColSep
	iRowSep = Parent.gRowSep

	With frm1.vspdData4
		For IRow = 1 To .MaxRows
			.Row = IRow
			.Col = 0

			Select Case .Text
         		Case ggoSpread.DeleteFlag                                                                '��: Update�߰� 
					
         		Case Else 
					strVal = strVal & "C" & iColSep & IRow & iColSep 
					.Col = C_REPAY_TYPE  
					strVal = strVal & UCase(Trim(.text))  & iColSep 
					.Col = C_BANK_ACCT_NO 
					strVal = strVal & Trim(.text) & iColSep
					.Col = C_BANK_CD 
					strVal = strVal & Trim(.text) & iColSep				
					.Col = C_REPAY_ACCT_CD  
					strVal = strVal & Trim(.text) & iColSep
					.Col = C_DOCCUR
					strVal = strVal & Trim(.text) & iColSep
					.Col = C_XCH_RATE
					strVal = strVal & Trim(UNIConvNum(.text,0))  & iColSep
					.Col = C_REPAY_AMT 
					strVal = strVal & Trim(UNIConvNum(.Text,0)) & iColSep
					.Col = C_REPAY_LOC_AMT
					strVal = strVal & Trim(UNIConvNum(.Text,0)) & iColSep
					.Col = C_DESC 
					strVal = strVal & Trim(.text) & iRowSep
					
                    IGrpCnt = IGrpCnt + 1
			End Select            
		Next		
	End With
	
	frm1.txtMaxRows4.value = IGrpCnt - 1
	frm1.txtSpread4.value  = strVal   
End Sub

'=======================================================================================================
' Function Name : Spread�� ���� Return
' Function Desc : 
'=======================================================================================================
Sub GetSpread()
    Dim IGrpCnt
    Dim strVal, iColSep, iRowSep
    Dim IRow
    
    IGrpCnt = 1
	strVal = ""  
	iColSep = Parent.gColSep
	iRowSep = Parent.gRowSep
	 
	With frm1.vspdData
		For IRow = 1 To .MaxRows
			.Row = IRow
			.Col = 0
					  
			Select Case .Text
				Case ggoSpread.DeleteFlag                                                                '��: Update�߰� 
					
         		Case Else 
         			strVal = strVal & "C" & iColSep 
					.Col = C_ITEM_SEQ  
					strVal = strVal & UCase(Trim(.Text)) & iColSep 
					.Col = C_ACCT_CD  
					strVal = strVal & Trim(.text) & iColSep 
					.Col = C_ITEMAMT
					strVal = strVal & Trim(UNIConvNum(.Text,0))   & iColSep   
					.Col = C_ITEMLOCAMT
					strVal = strVal & Trim(UNIConvNum(.Text,0))   & iColSep                
					.Col = C_ITEMDESC 
					strVal = strVal & Trim(.text)  & iRowSep		              
                    IGrpCnt = IGrpCnt + 1
			End Select            
		Next		
	End With

	frm1.txtMaxRows.value = IGrpCnt-1
	frm1.txtSpread.value  = strVal     
End Sub

'==========================================================================================
'   Event Name : DocCur_OnChange
'   Event Desc : 
'==========================================================================================
Sub DocCur_OnChange(byVal strDocCur, byVal Row)
    lgBlnFlgChgValue = True

	If UCase(Trim(strDocCur)) = UCase(parent.gCurrency) Then
		frm1.vspdData4.Col  = C_XCH_RATE
		frm1.vspdData4.Text = "1"
		ggoSpread.Source    = frm1.vspdData4	
		ggoSpread.SpreadLock     C_XCH_RATE , Row , C_XCH_RATE , Row
		ggoSpread.SSSetProtected C_XCH_RATE	, Row , Row
	Else
		Call FindExchRate(UniConvDateToYYYYMMDD(frm1.txtRePayDT.text,parent.gDateFormat,""), UCase(Trim(strDocCur)),Row)
		ggoSpread.Source = frm1.vspdData4	
		ggoSpread.SpreadUnLock  C_XCH_RATE , Row , C_XCH_RATE , Row
		ggoSpread.SSSetRequired C_XCH_RATE , Row , Row
	End If
	
	Call DoMulti(Row)
	Call Dosum()
End Sub

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave()     
    On Error Resume Next                                                   
	Err.Clear 
	 
	Call LayerShowHide(1)
	
    DbSave    = False                                                          '��: Processing is NG
	lgRetFlag = False
	
	With frm1
	    .txtMode.value = lgIntFlgMode	
		 strMode	   = .txtMode.value
		 		 
		 Call GetSpread1()	'���Աݳ��� 
		 Call GetSpread4()	'��ݵ�� 
		 Call GetSpread()	'�δ��� 
		 Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)
	End With		
	
	DbSave = True                                 
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk(ByVal PAYMNo)		
    Call LayerShowHide(0)
    
    If lgIntFlgMode = parent.OPMD_CMODE Then
		frm1.txtRePayNO.value = PAYMNo
	End If	      

    Call ggoOper.ClearField(Document, "2")
    
    ggoSpread.Source = frm1.vspddata1	:	Call ggoSpread.ClearSpreadData()
	ggoSpread.Source = frm1.vspddata	:	Call ggoSpread.ClearSpreadData()
 	ggoSpread.Source = frm1.vspdData4	:	Call ggoSpread.ClearSpreadData()
       
    Call InitVariables()

	If DbQuery = False Then
		Exit Function
	End If

	lgDBSaveOK = 1
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function  DbDelete() 
	Dim strVal, lsALLCNO 

	Call LayerShowHide(1)
	On Error Resume Next
	Err.Clear
   
	DbDelete = False    
   
	frm1.txtMode.value = parent.UID_M0003	
    strMode = frm1.txtMode.value
 
	strVal = BIZ_PGM_DEL_ID & "?txtRePayNO=" & Trim(frm1.txtRePayNO.value)
	strVal = strVal & "&txtMode=" & strMode
   
	Call RunMyBizASP(MyBizASP, strVal)
    
	DbDelete = True                            
    Set gActiveElement = document.ActiveElement     
End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete�� �������϶� ���� 
'========================================================================================
Function DbDeleteOk()		
	Call LayerShowHide(0)
	 
	Call ggoOper.ClearField(Document, "A")
    Call ggoOper.LockField (Document, "N")
    
    ggoSpread.Source = frm1.vspdData1	:	Call ggoSpread.ClearSpreadData()
    ggoSpread.Source = frm1.vspdData	:	Call ggoSpread.ClearSpreadData()
    ggoSpread.Source = frm1.vspdData4	:	Call ggoSpread.ClearSpreadData()

    Call InitVariables()
    Call ClickTab1()
    Call SetDefaultVal()
End Function

'========================================================================================
' Function Name : ClearRow() 
' Function Desc : ClearRow
'========================================================================================
Sub ClearRow(ByVal Row )
	With frm1.vspdData4
		.Row    = Row
		.ReDraw = False
		.Col	= C_BANK_ACCT_NO	 :	.text = ""
		.Col	= C_BANK_CD			 :	.text = ""
		.Col	= C_BANK_NM			 :	.text = ""
		.Col	= C_DOCCUR			 :	.text = parent.gCurrency
		.Col	= C_XCH_RATE		 :	.text = "1"
		.Col	= C_REPAY_AMT		 :	.text = ""
		.Col	= C_REPAY_LOC_AMT	 :	.text = ""
		.Col	= C_REPAY_ACCT_CD	 :	.text = ""
		.Col	= C_REPAY_ACCT_NM :	.text = ""
		.ReDraw = True
	End With
End Sub

'========================================================================================
' Function Name : DoSum() 
' Function Desc : ���������� ���� ���� Display�Ѵ�.
'========================================================================================
Sub DoSum()
	Dim tmpRePayLocSum, tmpPlanLocSum, tmpPlanIntLocSum, tmpEtcLocSum,tmpDrLocSum,tmpCrLocSum
	Dim tmpRePaySum, tmpPlanSum, tmpPlanIntSum, tmpEtcSum,tmpDrSum,tmpCrSum
	
	DIm Row
	Dim liTemp,liTemp2
	
	tmpRePayLocSum	 = 0
	tmpPlanLocSum	 = 0
	tmpPlanIntLocSum = 0
	tmpEtcLocSum	 = 0
	tmpDrLocSum		 = 0
	tmpCrLocSum		 = 0
	tmpRePaySum      = 0
	tmpPlanSum       = 0
	

	With frm1
		Select Case gSelframeFlg
			Case TAB1				'��ݵ�� 
				tmpRePayLocSum      = FncSumSheet1(.vspdData4, C_REPAY_LOC_AMT ,   1, .vspdData4.MaxRows, False, -1, -1, "V")
				.txtPaymTotLocAmt.text = UNIConvNumPCToCompanyByCurrency(tmpRePayLocSum,parent.gCurrency,parent.ggAmtOfMoneyNo, "X", "X")			
			Case TAB2				'���Աݹ��� 
				tmpPlanLocSum    = FncSumSheet1(.vspdData1, C_REPAY_PLAN_LOC_AMT    ,1, .vspdData1.MaxRows, False, -1, -1, "V")
				tmpPlanIntLocSum = FncSumSheet1(.vspdData1, C_REPAY_PLAN_INT_LOC_AMT,1, .vspdData1.MaxRows, False, -1, -1, "V")
				tmpRePayLocSum   = FncSumSheet1(.vspdData4, C_REPAY_LOC_AMT         ,1, .vspdData4.MaxRows, False, -1, -1, "V")
				
				.txtRePayTotLocAmt.text = UNIConvNumPCToCompanyByCurrency(tmpPlanLocSum,parent.gCurrency,parent.ggAmtOfMoneyNo, "X", "X")			
				.txtRePayIntLocAmt.text = UNIConvNumPCToCompanyByCurrency(tmpPlanIntLocSum,parent.gCurrency,parent.ggAmtOfMoneyNo, "X", "X")									
				.txtPaymTotLocAmt.text = UNIConvNumPCToCompanyByCurrency(tmpRePayLocSum,parent.gCurrency,parent.ggAmtOfMoneyNo, "X", "X")			
			Case TAB3				'�δ��� 
				For row = 1 To .vspdData.maxRows
					.vspdData.Col = 0
					.vspdData.Row = row
				
					If .vspdData.Text <> ggoSpread.DeleteFlag Then
						.vspdData.Col = C_ITEMLOCAMT
						tmpDrLocSum = CDbl(tmpDrLocSum) + unicdbl(.vspdData.text) 
					End If	
				Next
				
				tmpEtcLocSum = tmpDrLocSum - tmpCrLocSum
				.txtEtcLocAmt.text     = UNIConvNumPCToCompanyByCurrency(tmpEtcLocSum,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")		
		End Select	
	End With		
End Sub    
 
'========================================================================================
' Function Name : DoMulti()
' Function Desc : �ݾװ� ȯ���� ���ؼ� �ڱ��� ���Ѵ�.
'========================================================================================
Sub DoMulti(Byval Row)
	Dim TmpXchRate, TmpLocAmt
	DIm TmpLoanBalAmt, TmpLoanIntAmt,tmpItemAmt
	
	TmpXchRate=0
	TmpLocAmt=0
	TmpLoanBalAmt= 0
	TmpLoanIntAmt =0
	Select Case gSelframeFlg
		Case TAB1
			With frm1.vspdData4
				.Row = Row
				.Col = C_XCH_RATE
				TmpXchRate= UniCDbl(.text)
				.Col = C_REPAY_AMT
				TmpLocAmt= UniCDbl(.text) * TmpXchRate
				TmpLocAmt= UNIConvNumPCToCompanyByCurrency(TmpLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, "X", "X")			
				.Col = C_REPAY_LOC_AMT 
				.text= TmpLocAmt
			End With
		Case TAB2
			With frm1.vspdData1
				.Row = Row
				.Col = C_LOAN_XCH_RATE
				TmpXchRate= UniCDbl(.text)				
				.Col = C_LOAN_BAL_AMT
				TmpLoanBalAmt= UniCDbl(.text)
				.Col = C_REPAY_PLAN_AMT
				
				If UniCDbl(.text) <> TmpLoanBalAmt Then	' �����ݾװ� �ܾװ� ���Ͽ� ������ �ڱ��ݾ� ���� 
					TmpLocAmt = UniCDbl(.text) * TmpXchRate
					TmpLocAmt = UNIConvNumPCToCompanyByCurrency(TmpLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, "X", "X")			
				Else
					.Col = C_LOAN_BAL_LOC_AMT 
					TmpLocAmt= .text
				End If
				.Col  = C_REPAY_PLAN_LOC_AMT 
				.Text = TmpLocAmt				
				
				.Col  =C_REPAY_PLAN_INT_AMT			'DCLOCAMT
				TmpLoanIntAmt = UniCDbl(.text)
				.Col  = C_REPAY_PLAN_INT_LOC_AMT
				.Text = UNIConvNumPCToCompanyByCurrency(TmpLoanIntAmt* TmpXchRate,parent.gCurrency,parent.ggAmtOfMoneyNo, "X", "X")			
			End With
		Case TAB3
			With frm1.vspdData		
				.Row = Row
				.Col = C_ITEMAMT
				tmpItemAmt = .Text
				.Col = C_ITEMLOCAMT
				.text = tmpItemAmt
			End With				
	End Select 
End Sub  

'========================================================================================
' Function Name : CheckIntAmt()
' Function Desc : �������޾��� �ݾ��� Ȯ���Ͽ� ���ڰ����� �Է��ʼ� ���� ���� 
'========================================================================================
Sub CheckIntAmt()
	Dim ii
	Dim RepayPayObj
	
	With frm1.vspdData1	
		For ii = 1 To .MaxRows
		
			.Col = C_REPAY_PAY_OBJ
			.Row = ii
			RepayPayObj = Trim(.Text)
			
			.Col = C_REPAY_PLAN_INT_AMT
				
			If RepayPayObj = "DI" AND UniCdbl(.Text) > 0 Then
				ggoSpread.SpreadUnLock  C_REPAY_INT_ACCT_CD_POP, ii, C_REPAY_INT_ACCT_CD_POP, ii
				ggoSpread.SpreadUnLock  C_REPAY_INT_ACCT_CD    , ii, C_REPAY_INT_ACCT_CD    , ii
				ggoSpread.SSSetRequired C_REPAY_INT_ACCT_CD    , ii, ii	
			Else
				.Col  = C_REPAY_INT_ACCT_CD
				.Text = ""
				ggoSpread.SpreadLock     C_REPAY_INT_ACCT_CD_POP, ii, C_REPAY_INT_ACCT_CD_POP, ii
				ggoSpread.SpreadLock     C_REPAY_INT_ACCT_CD    , ii, C_REPAY_INT_ACCT_NM , ii
				ggoSpread.SSSetProtected C_REPAY_INT_ACCT_CD    , ii, ii
			End If
		Next
	End With					
End Sub

'=======================================================================================================
'   Event Name : txtRePayDT_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtRePayDT_DblClick(Button)
    If Button = 1 Then
        frm1.txtRePayDT.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtRePayDT.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtRePayDT_onblur()
'   Event Desc : 
'=======================================================================================================
Function txtRePayDT_onblur()
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii, jj
	Dim arrVal1, arrVal2
 
	lgBlnFlgChgValue = True
	
	With frm1
		If LTrim(RTrim(.txtDeptCd.value)) <> "" and Trim(.txtRePayDT.Text <> "") Then
			strSelect	=			 " Distinct org_change_id "    		
			strFrom		=			 " b_acct_dept(NOLOCK) "		
			strWhere	=			 " dept_Cd =  " & FilterVar(LTrim(RTrim(.txtDeptCd.value)), "''", "S") 
			strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
			strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
			strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(.txtRePayDT.Text, gDateFormat,""), "''", "S") & "))"			
	
			IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
					
			If IntRetCD = False  OR Trim(Replace(lgF0,Chr(11),"")) <> Trim(.hOrgChangeId.value) Then
				.txtDeptCd.value = ""
				.txtDeptNm.value = ""
				.hOrgChangeId.value = ""
				.txtDeptCd.focus
			End If
		End If
	End With
End Function

'=======================================================================================================
'   Event Name : deptCheck()
'   Event Desc : 
'=======================================================================================================
Function deptCheck()
	Dim strSelect
	Dim strFrom
	Dim strWhere 	
    Dim IntRetCD 
	
	With Frm1
		If  .txtDeptCd.value = "" Then
			.txtDeptCd.value = ""
			.txtDeptNm.value=""
			.hCostCd.value=""
			.hInternalCD.value=""
			.hbizcd.value = ""
			.horgChangeId.value = ""
			lgBlnFlgChgValue = True 
			Exit Function
		End If
    
		If Trim(.txtRePayDT.Text = "") Then    
			Exit Function
		End If

		strSelect	=			 " distinct org_change_id "
		strFrom		=			 " b_acct_dept "
		strWhere	=			 " org_change_id = (select distinct org_change_id "
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(.txtRePayDT.Text, parent.gDateFormat,""), "''", "S") & "))"			
			
		IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		If IntRetCD = False or Trim(Replace(lgF0,Chr(11),"")) <> .horgChangeId.value  Then
			.txtDeptCd.value = ""
			.txtDeptNm.value=""
			.hCostCd.value=""
			.hInternalCD.value=""
			.hbizcd.value = ""
			.horgChangeId.value = ""
			lgBlnFlgChgValue = True
			Exit Function
		End If	
	End With
End Function

'=======================================================================================================
' Area Name   : 4.5.5 Spread Popup method 
' Description : This part declares spread popup method
'=======================================================================================================
'*******************************************************************************************************
'===================================== PopSaveSpreadColumnInf()  ======================================
' Name : PopSaveSpreadColumnInf()
' Description : �̵��� �÷��� ������ ���� 
'====================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub
'===================================== PopRestoreSpreadColumnInf()  ======================================
' Name : PopRestoreSpreadColumnInf()
' Description : �÷��� ���������� ������ 
'====================================================================================================
Sub PopRestoreSpreadColumnInf()
	Dim Row
    ggoSpread.Source = gActiveSpdSheet
    
	On Error Resume Next
	Err.Clear 	
	
	Select Case UCase(Trim(gActiveSpdSheet.Name))
		Case "VSPDDATA"					'�δ��� 
			Call ggoSpread.RestoreSpreadInf()
			Call InitSpreadSheet("C")
			Call ggoSpread.ReOrderingSpreadData()
			Call SetSpreadColor(-1, -1,"D")
			Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,-1,-1,parent.C_ETC_DOCCUR,C_ITEMAMT,"A" ,"I","X","X")    
			
			If lgIntFlgMode = parent.OPMD_UMODE Then
				ggoSpread.source = frm1.vspdData
				frm1.vspdData.ReDraw = False
				ggoSpread.SpreadLock	C_ACCT_CD       , -1    , C_ACCT_CD
				frm1.vspdData.ReDraw = True
			End If
		Case "VSPDDATA1"
			Call ggoSpread.RestoreSpreadInf()
			Call InitSpreadSheet("B")
			Call ggoSpread.ReOrderingSpreadData()
			Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData1,-1,-1,C_LOAN_DOCCUR,C_REPAY_PLAN_AMT    ,"A" ,"I","X","X")
			Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData1,-1,-1,C_LOAN_DOCCUR,C_REPAY_PLAN_INT_AMT,"A" ,"I","X","X")
			Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData1,-1,-1,C_LOAN_DOCCUR,C_LOAN_BAL_AMT      ,"A" ,"I","X","X")
			Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData1,-1,-1,C_LOAN_DOCCUR,C_LOAN_RDP_TOT_AMT  ,"A" ,"I","X","X")
			Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData1,-1,-1,C_LOAN_DOCCUR,C_LOAN_INT_TOT_AMT  ,"A" ,"I","X","X")
		Case "VSPDDATA4"
			Call ggoSpread.RestoreSpreadInf()
			Call InitSpreadSheet("A")
			Call ggoSpread.ReOrderingSpreadData()
			Call SetSpreadColor(-1, -1,"D")
			Call SetSpreadLock("A")
			Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData4,-1,-1,C_DOCCUR,C_REPAY_AMT,"A" ,"I","X","X")
			
			If lgIntFlgMode = parent.OPMD_UMODE Then
				For Row= 0 to frm1.vspdData4.MaxRows
					ggoSpread.source = frm1.vspdData4    
					frm1.vspdData4.Row = Row
					frm1.vspdData4.Col = C_REPAY_TYPE
					ggoSpread.SSSetProtected C_REPAY_ACCT_CD    ,Row, Row
					ggoSpread.SpreadUnLock	 C_REPAY_ACCT_CD_POP,Row, Row
				Next 
			End If
	End Select
	
	Call DoSum()
End Sub




'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc"  -->	
</HEAD>

<!--'======================================================================================================
'       					6. Tag�� 
'	���: Tag�κ� ���� 
'======================================================================================================= -->
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../image/table/seltab_up_bg.gif"><IMG height=23 src="../../image/table/seltab_up_left.gif" width=9></td>
								<td background="../../image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>��ݵ��</font></td>
								<td background="../../image/table/seltab_up_bg.gif" align="right"><IMG height=23 src="../../image/table/seltab_up_right.gif" width=10></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../image/table/tab_up_bg.gif"><IMG height=23 src="../../image/table/tab_up_left.gif" width=9></td>
								<td background="../../image/table/tab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>���Ա�����</font></td>
								<td background="../../image/table/tab_up_bg.gif" align="right"><IMG height=23 src="../../image/table/tab_up_right.gif" width=10></td>
						    </TR>
						</TABLE>
					</TD>	
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab3()">
							<TR>
								<td background="../../image/table/tab_up_bg.gif"><IMG height=23 src="../../image/table/tab_up_left.gif" width=9></td>
								<td background="../../image/table/tab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>�δ���</font></td>
								<td background="../../image/table/tab_up_bg.gif" align="right"><IMG height=23 src="../../image/table/tab_up_right.gif" width=10></td>
						    </TR>
						</TABLE>
					</TD>	
					
					<TD WIDTH=* align=right><span id="spnArInfo"><a href="vbscript:OpenPopupLoan()">���Ա�����</A>&nbsp;|&nbsp;</span><A HREF="VBSCRIPT:OpenPopupGL()">ȸ����ǥ</A>&nbsp;|&nbsp;<A HREF="VBSCRIPT:OpenPopupTempGL()">������ǥ</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
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
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR HEIGHT=20>
									<TD CLASS="TD5" NOWRAP>��ȯ��ȣ</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtRePayNO" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="13NXXU" ALT="��ȯ��ȣ"><IMG SRC="../../image/btnPopup.gif" NAME="btnRcptCD" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopupPay()"></TD>								
									<TD CLASS="TD6" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				
				<TR HEIGHT=120>
				   <TD WIDTH="100%">
				       <TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>��ȯ��</TD>
								<TD CLASS=TD6 NOWRAP>
								          <script language =javascript src='./js/f4255ma1_fpDateTime1_txtRePayDT.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>�μ�</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDeptCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag=22NXXU" ALT="�μ�"><IMG SRC="../../image/btnPopup.gif" NAME="btnDeptCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(frm1.txtDeptCd.value)"> <INPUT TYPE=TEXT NAME="txtDeptNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="24" ALT="�μ���"></TD>								
							<TR>
								<TD CLASS=TD5 NOWRAP>���ݻ�ȯ�ݾ��հ�(�ڱ�)</TD>
								<TD CLASS=TD6 NOWRAP>
								          <script language =javascript src='./js/f4255ma1_OBJECT3_txtRePayTotLocAmt.js'></script></TD>								
								<TD CLASS=TD5 NOWRAP>�������ޱݾ��հ�(�ڱ�)</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/f4255ma1_OBJECT2_txtRePayIntLocAmt.js'></script></TD>							          								          
							</TR>
							<TR>
								<TD class=TD5 NOWRAP>�δ���(�ڱ�)</TD>
								<TD class=TD6 NOWRAP><script language =javascript src='./js/f4255ma1_OBJECT22_txtEtcLocAmt.js'></script></TD></TD>
								<TD class=TD5 NOWRAP>��ݱݾ��հ�(�ڱ�)</TD>
								<TD class=TD6 NOWRAP><script language =javascript src='./js/f4255ma1_OBJECT22_txtPaymTotLocAmt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>������ʵ�1</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtUserFld1" SIZE=20 MAXLENGTH=18 tag="21" ALT="������ʵ�1"></TD>
								<TD CLASS="TD5" NOWFRAP>������ʵ�2</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtUserFld2" SIZE=20 MAXLENGTH=18 tag="21" ALT="������ʵ�2"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>���</TD>
								<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT TYPE=TEXT NAME="txtRePayDesc" SIZE=90 MAXLENGTH=128 tag="21" ALT="���"></TD>
							</TR>
					</TABLE>
					</TD>
				</TR>

				<TR>		
					<TD WIDTH="100%">
					<DIV ID="TabDiv" SCROLL=no>
						<TABLE <%=LR_SPACE_TYPE_20%>>
						    <TR HEIGHT="100%">
								<TD WIDTH="100%" COLSPAN="4">
									<script language =javascript src='./js/f4255ma1_OBJECT4_vspdData4.js'></script>
								</TD>
							</TR>
						</TABLE>
					</DIV>
					<DIV ID="TabDiv"  SCROLL="no">
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR HEIGHT="100%">
								<TD WIDTH="100%" COLSPAN="4">
									<script language =javascript src='./js/f4255ma1_OBJECT5_vspdData1.js'></script>
								</TD>
							</TR>
						</TABLE>
					</DIV>
					<DIV ID="TabDiv" SCROLL=no>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR HEIGHT="100%">
								<TD WIDTH="100%" COLSPAN="4">
									<script language =javascript src='./js/f4255ma1_OBJECT6_vspdData.js'></script>
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
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 tabindex=-1></IFRAME>
		</TD>		
	</TR>
</TABLE>

<TEXTAREA Class=hidden name="txtSpread"		tag="24" tabindex=-1></TEXTAREA>
<TEXTAREA Class=hidden name="txtSpread1"	tag="24" tabindex=-1></TEXTAREA>
<TEXTAREA Class=hidden name="txtSpread4"	tag="24" tabindex=-1></TEXTAREA>

<INPUT TYPE=hidden NAME="txtMaxRows"		tag="24" tabindex=-1>
<INPUT TYPE=hidden NAME="txtMaxRows1"		tag="24" tabindex=-1>
<INPUT TYPE=hidden NAME="txtMaxRows3"		tag="24" tabindex=-1>
<INPUT TYPE=hidden NAME="txtMaxRows4"		tag="24" tabindex=-1>

<INPUT TYPE=hidden NAME="txtMode"			tag="24" tabindex=-1>
<INPUT TYPE=hidden NAME="hItemSeq"			tag="24" tabindex=-1>
<INPUT TYPE=hidden NAME="txtFlgMode"		tag="24" tabindex=-1>
<INPUT TYPE=hidden NAME="txthTempglno"		tag="24" tabindex=-1>
<INPUT TYPE=hidden NAME="txthGlno"   		tag="24" tabindex=-1>

<!-- �μ��ڵ��Է½� �⺻����  -->

<INPUT TYPE=hidden NAME="horgChangeId" tag="24">


<script language =javascript src='./js/f4255ma1_OBJECT8_vspdData3.js'></script>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
