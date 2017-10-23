<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : ���ְ��� 
'*  3. Program ID           : S3171MA1
'*  4. Program Name         : STO�������� 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2002/11/11
'*  9. Modifier (First)     : cho in kuk
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit															'��: Turn on the Option Explicit option.

'========================================================================================================
Const BIZ_PGM_ID				= "s3171mb1.asp"						'Biz Logic ASP								
Const BIZ_PGM_JUMP_SOHDR_ID		= "s3111ma1"
Const BIZ_PGM_JUMP_SOSCHE_ID	= "s3114ma8"

Dim C_ItemCd			 'ǰ��  
Dim C_ItemPopup			 'ǰ���˾� 
Dim C_ItemName			 'ǰ��� 
Dim C_ItemSpec			 
Dim C_SoUnit			 '���� 
Dim C_SoUnitPopup		 '�����˾� 
Dim C_TrackingNo		 'Tracking No
Dim C_SoSupplyQty		 '�����ܷ�(ReqQty - SoQty)
Dim C_SoQty				 '���� 
Dim C_SoPrice			 '�ܰ� 
Dim C_SoPriceAutoChk	 '�ܰ��ڵ����üũ 
Dim C_SoPriceFlag		 '���ܰ�/���ܰ� 
Dim C_TotalAmt			 'ȭ�� ���ֱݾ�(�ŷ�ȭ��)
Dim C_NetAmt			 'Hidden ���ֱݾ� 
Dim C_VATAmt				
Dim C_PlantCd			 '�����ڵ� 
Dim C_PlantCdPopup		 '�����˾� 
Dim C_PlantNm			 '����� 
Dim C_DlvyDt			 '������       
Dim C_ShipToParty		 '��ǰó 
Dim C_ShipToPartyPopup	 '��ǰó�˾� 
Dim C_ShipToPartyNm		 '��ǰó�� 
Dim C_HsNo				 'HS��ȣ 
Dim C_HsNoPopup			 'HS��ȣ Popup
Dim C_TolMoreRate		 '�����������(+)
Dim C_TolLessRate		 '�����������(-)
Dim C_VatType				
Dim C_VatTypePopup		 
Dim C_VatTypeNm			
Dim C_VatRate				
Dim C_VatIncFlag			
Dim C_VatIncFlagNm		
Dim C_RetType				
Dim C_RetTypePopup		 
Dim C_RetTypeNm			 
Dim C_LotNo				 
Dim C_LotSeq				 
Dim C_PreDnNo			 '���Ϲ�ȣ for ���ֳ������� 
Dim C_PreDnSeq			 '���ϼ��� for ���ֳ������� 
Dim C_DnReqDt			 '���Ͽ�û���� 
Dim C_BonusQty			 '��������(��)        
Dim C_SlCd				 'â���ڵ� 
Dim C_SlCdPopup			 'â���˾� 
Dim C_SlNm				 'â��� 
Dim C_Remark			 '��� 
Dim C_SoSts				 '����������� 
Dim C_BillQty			 '������� 
Dim C_BaseQty			 '������ 
Dim C_BonusBaseQty		 '�������� 
Dim C_MaintSeq			 '�������� 
Dim C_OrderSeq			 '�ֹ������� 
Dim C_APSHost			 'APS Host
Dim C_APSPort			 'APS Port
Dim C_CTPTimes			 'CTP Check Ƚ�� 
Dim C_CTPCheckFlag		 'CTP Check Flag
Dim C_SoSeq				 '���ּ��� 
Dim C_PreSoNo			 '���ֹ�ȣ for ���ֳ������� 
Dim C_PreSoSeq			 '���ּ��� for ���ֳ������� 
Dim C_CustPoSeq			 '���Ź��ֹ�ȣ 

Dim ext1_qty 
Dim ext2_qty 
Dim ext3_qty 
Dim ext1_amt 
Dim ext2_amt 
Dim ext3_amt 
Dim ext1_cd  
Dim ext2_cd 
Dim ext3_cd 

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim IsOpenPop
Dim iDBSYSDate
Dim EndDate, StartDate
iDBSYSDate = "<%=GetSvrDate%>"
'------ ��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ ------
EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
'------ ��: �ʱ�ȭ�鿡 �ѷ����� ���� ��¥ ------
StartDate = UNIDateAdd("m", -1, EndDate, parent.gDateFormat)

Const lsConfirm		= "CONFIRM"
Const lsPricePad	= "PRICE"

Dim lsItemCode
Dim lsSoUnit
Dim lsSoQty
Dim lsPriceQty
Dim lsAPSHost
Dim lsAPSPort
Dim lsCTPTimes
Dim lsCTPCheckFlag
Dim arrCollectVatType

'========================================================================================================
Sub initSpreadPosVariables()  
	C_ItemCd			 = 1   'ǰ��  
	C_ItemPopup			 = 2   'ǰ���˾� 
	C_ItemName			 = 3   'ǰ��� 
	C_ItemSpec			 = 4
	C_SoUnit			 = 5   '���� 
	C_SoUnitPopup		 = 6   '�����˾� 
	C_TrackingNo		 = 7   'Tracking No
	C_SoSupplyQty		 = 8   '�����ܷ�(ReqQty - SoQty)
	C_SoQty				 = 9   '���� 
	C_SoPrice			 = 10  '�ܰ� 
	C_SoPriceAutoChk	 = 11  '�ܰ��ڵ����üũ 
	C_SoPriceFlag		 = 12  '���ܰ�/���ܰ� 
	C_TotalAmt			 = 13  'ȭ�� ���ֱݾ�(�ŷ�ȭ��)
	C_NetAmt			 = 14  'Hidden ���ֱݾ� 
	C_VATAmt			 = 15
	C_PlantCd			 = 16  '�����ڵ� 
	C_PlantCdPopup		 = 17  '�����˾� 
	C_PlantNm			 = 18  '����� 
	C_DlvyDt			 = 19  '������       
	C_ShipToParty		 = 20  '��ǰó 
	C_ShipToPartyPopup	 = 21  '��ǰó�˾� 
	C_ShipToPartyNm		 = 22  '��ǰó�� 
	C_HsNo				 = 23  'HS��ȣ 
	C_HsNoPopup			 = 24  'HS��ȣ Popup
	C_TolMoreRate		 = 25  '�����������(+)
	C_TolLessRate		 = 26  '�����������(-)
	C_VatType			 = 27
	C_VatTypePopup		 = 28
	C_VatTypeNm			 = 29
	C_VatRate			 = 30
	C_VatIncFlag		 = 31
	C_VatIncFlagNm		 = 32
	C_RetType			 = 33
	C_RetTypePopup		 = 34
	C_RetTypeNm			 = 35
	C_LotNo				 = 36
	C_LotSeq			 = 37
	C_PreDnNo			 = 38  '���Ϲ�ȣ for ���ֳ������� 
	C_PreDnSeq			 = 39  '���ϼ��� for ���ֳ������� 
	C_DnReqDt			 = 40  '���Ͽ�û���� 
	C_BonusQty			 = 41  '��������(��)        
	C_SlCd				 = 42  'â���ڵ� 
	C_SlCdPopup			 = 43  'â���˾� 
	C_SlNm				 = 44  'â��� 
	C_Remark			 = 45  '��� 
	C_SoSts				 = 46  '����������� 
	C_BillQty			 = 47  '������� 
	C_BaseQty			 = 48  '������ 
	C_BonusBaseQty		 = 49  '�������� 
	C_MaintSeq			 = 50  '�������� 
	C_OrderSeq			 = 51  '�ֹ������� 
	C_APSHost			 = 52  'APS Host
	C_APSPort			 = 53  'APS Port
	C_CTPTimes			 = 54  'CTP Check Ƚ�� 
	C_CTPCheckFlag		 = 55  'CTP Check Flag
	C_SoSeq				 = 56  '���ּ��� 
	C_PreSoNo			 = 57  '���ֹ�ȣ for ���ֳ������� 
	C_PreSoSeq			 = 58  '���ּ��� for ���ֳ������� 
	C_CustPoSeq			 = 59  '���Ź��ֹ�ȣ 
	
	ext1_qty =	0 
	ext2_qty =	0
	ext3_qty =	0
	ext1_amt =	0
	ext2_amt =	0
	ext3_amt =	0
	ext1_cd  =	""
	ext2_cd	 =	""
	ext3_cd  =	""
	
End Sub

'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = Parent.OPMD_CMODE						'��: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '��: Indicates that no value changed    
    lgStrPrevKey      = ""                                      '��: initializes Previous Key
    lgSortKey         = 1                                       '��: initializes sort direction
End Sub

'========================================================================================================	
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	frm1.txtConSoNo.focus
	
	frm1.txtSoDt.text = EndDate
	
	frm1.btnConfirm.disabled	= True
	frm1.btnConfirm.value		= "Ȯ��ó��"
	frm1.btnDNCheck.disabled	= True
	frm1.btnATPCheck.disabled	= True
	frm1.btnCTPCheck.disabled	= True
	frm1.btnAvlStkRef.disabled	= True
	lgBlnFlgChgValue			= False
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'==========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"-->
	<% Call LoadInfTB19029A("I","*","NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()
	
	With frm1.vspdData
 
	  ggoSpread.Source = frm1.vspdData
	  'patch version
	  ggoSpread.Spreadinit "V20021214",,parent.gAllowDragDropSpread
	  .ReDraw = false
	  
	  .MaxCols	=	C_CustPoSeq	  +  1							' ��: Add 1 to Maxcols	  	  
	  .MaxRows = 0												' ��: Clear spreadsheet data 
  
	  Call GetSpreadColumnPos("A")	 
	  
							'ColumnPosition		Header              Width	Align(0:L,1:R,2:C)  Row   Length  CharCase(0:L,1:N,2:U)
	  ggoSpread.SSSetEdit	C_SoSeq,			"���ּ���",		3,		1
	  ggoSpread.SSSetEdit	C_CustPoSeq,		"���Ź��ּ���",	3,		1
	  ggoSpread.SSSetEdit	C_ItemCd,			"ǰ��",			18,		,					,	  18,	  2
	  ggoSpread.SSSetEdit	C_ItemSpec,			"�԰�",			20
							'ColumnPosition		Row
	  ggoSpread.SSSetButton	C_ItemPopup			
	  ggoSpread.SSSetEdit	C_ItemName,			"ǰ���",		25,		,					,	  40
	  ggoSpread.SSSetEdit	C_TrackingNo,		"Tracking No",	25,		,					,	  25,	  2
								   'ColumnPosition      Header              Width	Grp					  IntegeralPart					DeciPointpart               Align				 Sep				PZ  Min Max 
	  ggoSpread.SSSetFloat	C_SoSupplyQty,		"�����ܷ�",		15,		parent.ggQtyNo,	      ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
	  ggoSpread.SSSetFloat	C_SoQty,			"����",			15,		parent.ggQtyNo,		  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
	  ggoSpread.SSSetEdit	C_SoUnit,			"����",			8,		,					,	  3,	  2
	  ggoSpread.SSSetButton	C_SoUnitPopup		
								   'ColumnPosition      Header				Width	Align(0:L,1:R,2:C)  Format         Row
	  ggoSpread.SSSetDate	C_DlvyDt,			"������",		10,		2,					parent.gDateFormat
	  ggoSpread.SSSetDate	C_DnReqDt,			"���Ͽ�������",	15,		2,					parent.gDateFormat
	  ggoSpread.SSSetEdit	C_ShipToParty,		"��ǰó",		10,		,					,	  10,	  2
	  ggoSpread.SSSetButton	C_ShipToPartyPopup	
	  ggoSpread.SSSetEdit	C_ShipToPartyNm,	"��ǰó��",		10
	  ggoSpread.SSSetFloat	C_SoPrice,			"�ܰ�",			15,		parent.ggUnitCostNo,  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
	  ggoSpread.SSSetCheck	C_SoPriceAutoChk,	"",		2,	,	,	True
								   'ColumnPosition      Header				Width	Align(0:L,1:R,2:C)  ComboEditable  Row
	  ggoSpread.SSSetCombo	C_SoPriceFlag,		"�ܰ�����",		10,		2
	  ggoSpread.SetCombo		"���ܰ�" & vbTab & "���ܰ�",C_SoPriceFlag
								   'ColumnPosition      Header              Width	Grp						IntegeralPart				DeciPointpart				Align				Sep					PZ  Min Max 	  
	  ggoSpread.SSSetFloat	C_TotalAmt,			"�ݾ�",			15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
	  ggoSpread.SSSetFloat	C_NetAmt,			"���ݾ�",		15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
	  ggoSpread.SSSetFloat	C_VATAmt,			"VAT�ݾ�",		15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
	  ggoSpread.SSSetFloat	C_BonusQty,			"������" ,		15,		parent.ggQtyNo,			ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
								   'ColumnPosition		Header              Width	Align(0:L,1:R,2:C)  Row   Length  CharCase(0:L,1:N,2:U)
	  ggoSpread.SSSetEdit	C_HsNo,				"HS��ȣ",		15,		,					,	  20,	  2
	  ggoSpread.SSSetButton	C_HsNoPopup			
	  ggoSpread.SSSetEdit	C_VatType,			"VAT����",		10,		,					,	  5,	  2
	  ggoSpread.SSSetButton	C_VatTypePopup		
	  ggoSpread.SSSetEdit	C_VatTypeNm,		"VAT������",	20
	  ggoSpread.SSSetFloat	C_VatRate,			"VAT��",		10,		parent.ggExchRateNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
	
	  ggoSpread.SSSetCombo	C_VatIncFlagNm,		"VAT���Ա��и�",15,		2
	  ggoSpread.SetCombo		"����" & vbTab & "����",C_VatIncFlagNm
	  ggoSpread.SSSetEdit	C_VatIncFlag,		"VAT���Ա���",	5,		2
	  ggoSpread.SetCombo		"1"		   & vbTab & "2",		C_VatIncFlag
	  ggoSpread.SSSetEdit	C_RetType,			"��ǰ����",		10,		,					,	  5,	  2
	  ggoSpread.SSSetButton	C_RetTypePopup
	  ggoSpread.SSSetEdit	C_RetTypeNm,		"��ǰ������",	20
	  ggoSpread.SSSetEdit	C_LotNo,			"LOT NO",		12,		,					,	  18,	  2
	  Call AppendNumberPlace("7","3","0")
								   'ColumnPosition      Header					Width	Grp		IntegeralPart				DeciPointpart				Align				Sep					PZ  Min Max 	  
	  ggoSpread.SSSetFloat	C_LotSeq,			"LOT NO ����" ,		15,		"7",	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"  
	  Call AppendNumberPlace("6","9","6")
	  ggoSpread.SSSetFloat	C_TolMoreRate,		"�����������(+)" ,	15,		"6",	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"  
	  ggoSpread.SSSetFloat	C_TolLessRate,		"�����������(-)" ,	15,		"6",	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
								   'ColumnPosition		Header              Width	Align(0:L,1:R,2:C)  Row   Length  CharCase(0:L,1:N,2:U)
	  ggoSpread.SSSetEdit	C_PlantCd,			"����",			8,		,					,	  4,	  2
	  ggoSpread.SSSetButton	C_PlantCdPopup		
	  ggoSpread.SSSetEdit	C_PlantNm,			"�����",		8
	  ggoSpread.SSSetEdit	C_SlCd,				"â��",			8,		,					,	  7,	  2
	  ggoSpread.SSSetButton	C_SlCdPopup
	  ggoSpread.SSSetEdit	C_SlNm,				"â���",		8
	  ggoSpread.SSSetEdit	C_Remark,			"���",			60,		,					,	  60
								   'ColumnPosition      Header				Width	Grp				IntegeralPart				DeciPointpart				Align				Sep					PZ  Min Max 	  	
	  ggoSpread.SSSetFloat	C_BillQty,			"�������",		15,		parent.ggQtyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
								   'ColumnPosition		Header              Width	Align(0:L,1:R,2:C)  Row   Length  CharCase(0:L,1:N,2:U)
	  ggoSpread.SSSetEdit	C_MaintSeq,			"����SEQ",		10
	  ggoSpread.SSSetEdit	C_OrderSeq,			"�ֹ�������",	10
	  ggoSpread.SSSetEdit	C_APSHost,			"APSHost",			5,		,					,	  20
	  ggoSpread.SSSetEdit	C_APSPort,			"APSPort",			5,		,					,	  20
	  ggoSpread.SSSetEdit	C_CTPTimes,			"CTPTimes",			5,		,					,	  3
	  ggoSpread.SSSetEdit	C_CTPCheckFlag,		"CTPCheckFlag",		5,		,					,	  2
	  ggoSpread.SSSetEdit	C_PreDnNo,			"���Ϲ�ȣ",		18,		,					,	  18,	  2
	  ggoSpread.SSSetEdit	C_PreDnSeq,			"���ϼ���",		10,		,					,	  3,	  1
	  

      Call ggoSpread.MakePairsColumn(C_ItemCd,C_ItemPopup)
      Call ggoSpread.MakePairsColumn(C_SoUnit,C_SoUnitPopup)
      Call ggoSpread.MakePairsColumn(C_PlantCd,C_PlantCdPopup)      
      Call ggoSpread.MakePairsColumn(C_ShipToParty,C_ShipToPartyPopup)
      Call ggoSpread.MakePairsColumn(C_HsNo,C_HsNoPopup)      
      Call ggoSpread.MakePairsColumn(C_VatType,C_VatTypePopup)      
      Call ggoSpread.MakePairsColumn(C_RetType,C_RetTypePopup)
      Call ggoSpread.MakePairsColumn(C_SlCd,C_SlCdPopup)

      Call ggoSpread.SSSetColHidden(C_PreDnNo,C_PreDnNo,True)
      Call ggoSpread.SSSetColHidden(C_PreDnSeq,C_PreDnSeq,True)
      Call ggoSpread.SSSetColHidden(C_PreSoNo,C_PreSoNo,True)
      Call ggoSpread.SSSetColHidden(C_PreSoSeq,C_PreSoSeq,True)
      Call ggoSpread.SSSetColHidden(C_SoSeq,C_SoSeq,True)      
      Call ggoSpread.SSSetColHidden(C_CustPoSeq,C_CustPoSeq,True)      
      Call ggoSpread.SSSetColHidden(C_SoSts,C_SoSts,True)
      Call ggoSpread.SSSetColHidden(C_BillQty,C_BillQty,True)
      Call ggoSpread.SSSetColHidden(C_ShipToPartyNm,C_ShipToPartyNm,True)
      Call ggoSpread.SSSetColHidden(C_PlantNm,C_PlantNm,True)
      Call ggoSpread.SSSetColHidden(C_SlNm,C_SlNm,True)
      Call ggoSpread.SSSetColHidden(C_BaseQty,C_BaseQty,True)
      Call ggoSpread.SSSetColHidden(C_BonusBaseQty,C_BonusBaseQty,True)
      Call ggoSpread.SSSetColHidden(C_MaintSeq,C_MaintSeq,True)
      Call ggoSpread.SSSetColHidden(C_OrderSeq,C_OrderSeq,True)      
      Call ggoSpread.SSSetColHidden(C_APSHost,C_APSHost,True)
      Call ggoSpread.SSSetColHidden(C_APSPort,C_APSPort,True)
      Call ggoSpread.SSSetColHidden(C_CTPTimes,C_CTPTimes,True)
      Call ggoSpread.SSSetColHidden(C_CTPCheckFlag,C_CTPCheckFlag,True)
      Call ggoSpread.SSSetColHidden(C_VatIncFlag,C_VatIncFlag,True)
      Call ggoSpread.SSSetColHidden(C_NetAmt,C_NetAmt,True)     
      Call ggoSpread.SSSetColHidden(C_SoPriceAutoChk,C_SoPriceAutoChk,True)
      Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)				'��: ������Ʈ�� ��� Hidden Column
           
	  .ReDraw = true
	  
   End With
    
End Sub

'======================================================================================================
Sub SetSpreadLock()
End Sub


'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)

    With frm1
		.vspdData.ReDraw = False
									   'Col				Row         Row2
		ggoSpread.SSSetRequired	C_ItemCd,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ItemName,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ItemSpec,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_TrackingNo,	pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_SoSupplyQty,	pvStartRow, pvEndRow
		ggoSpread.SSSetRequired	C_SoPriceFlag,	pvStartRow, pvEndRow
		ggoSpread.SSSetRequired	C_SoQty,		pvStartRow, pvEndRow
		ggoSpread.SSSetRequired	C_SoUnit,		pvStartRow, pvEndRow
		ggoSpread.SSSetRequired	C_DlvyDt,		pvStartRow, pvEndRow
		ggoSpread.SSSetRequired	C_ShipToParty,	pvStartRow, pvEndRow
		ggoSpread.SSSetRequired	C_PlantCd,		pvStartRow, pvEndRow
		ggoSpread.SSSetRequired	C_TotalAmt,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_NetAmt,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_DnReqDt,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_VATAmt,		pvStartRow, pvEndRow
		ggoSpread.SSSetRequired	C_VatType,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_VatTypeNm,	pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_VatRate,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_VatIncFlag,	pvStartRow, pvEndRow
		ggoSpread.SSSetRequired	C_VatIncFlagNm, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_RetTypeNm,	pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_LotNo,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_LotSeq,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PreDNNo,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PreDNSeq,		pvStartRow, pvEndRow

		If frm1.HRetItemFlag.value = "Y" Then
			ggoSpread.SSSetRequired	C_RetType,	pvStartRow, pvEndRow
		Else
			ggoSpread.SSSetProtected C_RetType,	pvStartRow, pvEndRow
		End If 
 
		' �������/�������� ���ο� ���� ������,HS��ȣ ���� 
		If Trim(.HExportFlag.value) = "Y" Then 
			ggoSpread.SSSetProtected C_BonusQty, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired	C_HsNo,		pvStartRow, pvEndRow
		Else  
		    ggoSpread.SSSetProtected C_HsNo,		pvStartRow, pvEndRow
			ggoSpread.SSSetProtected C_HsNoPopup,pvStartRow, pvEndRow
		End If
    
		.vspdData.Col = C_ItemCd 
		.vspdData.Row = .vspdData.ActiveRow
		.vspdData.Action = 0
		.vspdData.EditMode = True
		
		.vspdData.ReDraw = True
    End With    
End Sub


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
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_ItemCd			 = iCurColumnPos(1)   'ǰ��  
			C_ItemPopup			 = iCurColumnPos(2)   'ǰ���˾� 
			C_ItemName			 = iCurColumnPos(3)   'ǰ��� 
			C_ItemSpec			 = iCurColumnPos(4)
			C_SoUnit			 = iCurColumnPos(5)   '���� 
			C_SoUnitPopup		 = iCurColumnPos(6)   '�����˾� 
			C_TrackingNo		 = iCurColumnPos(7)   'Tracking No
			C_SoSupplyQty		 = iCurColumnPos(8)   '�����ܷ�(ReqQty - SoQty)
			C_SoQty				 = iCurColumnPos(9)   '���� 
			C_SoPrice			 = iCurColumnPos(10)  '�ܰ� 
			C_SoPriceAutoChk	 = iCurColumnPos(11)  '�ܰ��ڵ����üũ 
			C_SoPriceFlag		 = iCurColumnPos(12)  '���ܰ�/���ܰ� 
			C_TotalAmt			 = iCurColumnPos(13)  'ȭ�� ���ֱݾ�(�ŷ�ȭ��)
			C_NetAmt			 = iCurColumnPos(14)  'Hidden ���ֱݾ� 
			C_VATAmt			 = iCurColumnPos(15)
			C_PlantCd			 = iCurColumnPos(16)  '�����ڵ� 
			C_PlantCdPopup		 = iCurColumnPos(17)  '�����˾� 
			C_PlantNm			 = iCurColumnPos(18)  '����� 
			C_DlvyDt			 = iCurColumnPos(19)  '������       
			C_ShipToParty		 = iCurColumnPos(20)  '��ǰó 
			C_ShipToPartyPopup	 = iCurColumnPos(21)  '��ǰó�˾� 
			C_ShipToPartyNm		 = iCurColumnPos(22)  '��ǰó�� 
			C_HsNo				 = iCurColumnPos(23)  'HS��ȣ 
			C_HsNoPopup			 = iCurColumnPos(24)  'HS��ȣ Popup
			C_TolMoreRate		 = iCurColumnPos(25)  '�����������(+)
			C_TolLessRate		 = iCurColumnPos(26)  '�����������(-)
			C_VatType			 = iCurColumnPos(27)
			C_VatTypePopup		 = iCurColumnPos(28)
			C_VatTypeNm			 = iCurColumnPos(29)
			C_VatRate			 = iCurColumnPos(30)
			C_VatIncFlag		 = iCurColumnPos(31)
			C_VatIncFlagNm		 = iCurColumnPos(32)
			C_RetType			 = iCurColumnPos(33)
			C_RetTypePopup		 = iCurColumnPos(34)
			C_RetTypeNm			 = iCurColumnPos(35)
			C_LotNo				 = iCurColumnPos(36)
			C_LotSeq			 = iCurColumnPos(37)
			C_PreDnNo			 = iCurColumnPos(38)  '���Ϲ�ȣ for ���ֳ������� 
			C_PreDnSeq			 = iCurColumnPos(39)  '���ϼ��� for ���ֳ������� 
			C_DnReqDt			 = iCurColumnPos(40)  '���Ͽ�û���� 
			C_BonusQty			 = iCurColumnPos(41)  '��������(��)        
			C_SlCd				 = iCurColumnPos(42)  'â���ڵ� 
			C_SlCdPopup			 = iCurColumnPos(43)  'â���˾� 
			C_SlNm				 = iCurColumnPos(44)  'â��� 
			C_Remark			 = iCurColumnPos(45)  '��� 
			C_SoSts				 = iCurColumnPos(46)  '����������� 
			C_BillQty			 = iCurColumnPos(47)  '������� 
			C_BaseQty			 = iCurColumnPos(48)  '������ 
			C_BonusBaseQty		 = iCurColumnPos(49)  '�������� 
			C_MaintSeq			 = iCurColumnPos(50)  '�������� 
			C_OrderSeq			 = iCurColumnPos(51)  '�ֹ������� 
			C_APSHost			 = iCurColumnPos(52)  'APS Host
			C_APSPort			 = iCurColumnPos(53)  'APS Port
			C_CTPTimes			 = iCurColumnPos(54)  'CTP Check Ƚ�� 
			C_CTPCheckFlag		 = iCurColumnPos(55)  'CTP Check Flag
			C_SoSeq				 = iCurColumnPos(56)  '���ּ��� 
			C_PreSoNo			 = iCurColumnPos(57)  '���ֹ�ȣ for ���ֳ������� 
			C_PreSoSeq			 = iCurColumnPos(58)  '���ּ��� for ���ֳ������� 
			C_CustPoSeq			 = iCurColumnPos(59)  '���Ź��ּ��� 	
    End Select    
End Sub


'==================================================================================================== 
Sub SetQuerySpreadColor(ByVal lRow)
 
 Dim SoSts, BillQty
    
    With frm1

		.btnConfirm.disabled = False
		.vspdData.ReDraw = False
    
		If .RdoConfirm.value = "Y" Then       
  
			ggoSpread.SSSetProtected C_ItemCd, -1, -1
			ggoSpread.SSSetProtected C_ItemPopup, -1, -1
			ggoSpread.SSSetProtected C_ItemName, -1, -1
			ggoSpread.SSSetProtected C_ItemSpec, -1, -1
			ggoSpread.SSSetProtected C_TrackingNo, -1, -1
			ggoSpread.SSSetProtected C_SoSupplyQty, -1, -1 
			ggoSpread.SSSetProtected C_VatRate, -1, -1
			ggoSpread.SSSetProtected C_RetTypeNm, -1, -1
			ggoSpread.SSSetProtected C_LotNo, -1, -1
			ggoSpread.SSSetProtected C_LotSeq, -1, -1
			ggoSpread.SSSetProtected C_PreDNNo, -1, -1
			ggoSpread.SSSetProtected C_PreDNSeq, -1, -1
			ggoSpread.SSSetProtected C_DnReqDt, -1, -1
			ggoSpread.SSSetProtected C_NetAmt, -1, -1
			ggoSpread.SSSetProtected C_VATAmt, -1, -1
			ggoSpread.SSSetProtected C_VatTypeNm, -1, -1
			ggoSpread.SSSetProtected C_VatIncFlag, -1, -1        
			ggoSpread.SpreadUnLock	C_SoUnit, -1, -1    
			ggoSpread.SpreadUnLock	C_SoUnitPopup, -1, -1
			ggoSpread.SpreadUnLock	C_SoQty, -1, -1
			ggoSpread.SpreadUnLock	C_SoPrice, -1, -1  
			ggoSpread.SpreadUnLock	C_SoPriceFlag, -1, -1
			ggoSpread.SpreadUnLock	C_SoPriceAutoChk, -1, -1
			ggoSpread.SpreadUnLock	C_NetAmt, -1, -1  
			ggoSpread.SpreadUnLock	C_TotalAmt, -1, -1 
			ggoSpread.SpreadUnLock	C_PlantCd, -1, -1
			ggoSpread.SpreadUnLock	C_PlantCdPopup, -1, -1  
			ggoSpread.SpreadUnLock	C_DlvyDt, -1, -1
			ggoSpread.SpreadUnLock	C_ShipToParty, -1, -1
			ggoSpread.SpreadUnLock	C_ShipToPartyPopup, -1, -1     
			ggoSpread.SpreadUnLock	C_VatType, -1, -1 
			ggoSpread.SpreadUnLock	C_VatTypePopup, -1, -1   
			ggoSpread.SpreadUnLock	C_VatIncFlagNm, -1, -1 
			ggoSpread.SpreadUnLock	C_TolMoreRate, -1, -1
			ggoSpread.SpreadUnLock	C_TolLessRate, -1, -1 
			ggoSpread.SpreadUnLock	C_BonusQty, -1, -1 
			ggoSpread.SpreadUnLock	C_SlCd, -1, -1
			ggoSpread.SpreadUnLock	C_SlCdPopup, -1, -1 
			ggoSpread.SpreadUnLock	C_Remark, -1, -1     
			ggoSpread.SSSetRequired  C_SoUnit, -1, -1 
			ggoSpread.SSSetRequired  C_SoQty, -1, -1  
			ggoSpread.SSSetRequired  C_SoPriceFlag, -1, -1
			ggoSpread.SSSetRequired  C_NetAmt, -1, -1  
			ggoSpread.SSSetRequired  C_TotalAmt, -1, -1 
			ggoSpread.SSSetRequired  C_PlantCd, -1, -1  
			ggoSpread.SSSetRequired  C_DlvyDt, -1, -1
			ggoSpread.SSSetRequired  C_ShipToParty, -1, -1       
			ggoSpread.SSSetRequired  C_VatType, -1, -1  
			ggoSpread.SSSetRequired  C_VatIncFlagNm, -1, -1  
			    
			' ��ǰ���� 
			If frm1.HRetItemFlag.value = "Y" Then
				ggoSpread.SpreadUnLock  C_RetType, -1, -1
				ggoSpread.SSSetRequired C_RetType, -1, -1
			Else
				ggoSpread.SSSetProtected C_RetType, -1, -1
			End If

			' �������/�������� ���ο� ���� ������,HS��ȣ ���� 
			If Trim(.HExportFlag.value) = "Y" Then 
				ggoSpread.SSSetProtected C_BonusQty, -1, -1
				ggoSpread.SpreadUnLock	C_HsNo, -1, -1
				ggoSpread.SSSetRequired  C_HsNo, -1, -1
			Else
				ggoSpread.SSSetProtected C_HsNo, -1, -1
				ggoSpread.SSSetProtected C_HsNoPopup, -1, -1
			End If
			    
			' ����û���ڰ� �������ں��� ������� �˸�ǥ�� 
			For lRow = 1 To .vspdData.MaxRows          
				.vspdData.Row = lRow : .vspdData.Col = C_DnReqDt
				If UniConvDateToYYYYMMDD(.vspdData.Text,parent.gDateFormat,"") < UniConvDateToYYYYMMDD(EndDate,parent.gDateFormat,"") Then Call sprRedComColor(C_DnReqDt,-1,-1)
			Next         
  
		Else  ' Ȯ��ó���� ���   
  
			ggoSpread.SSSetProtected C_ItemCd, -1, -1
			ggoSpread.SSSetProtected C_ItemPopup, -1, -1
			ggoSpread.SSSetProtected C_ItemName, -1, -1
			ggoSpread.SSSetProtected C_ItemSpec, -1, -1
			ggoSpread.SSSetProtected C_TrackingNo, -1, -1
			ggoSpread.SSSetProtected C_SoSupplyQty, -1, -1
			ggoSpread.SSSetProtected C_SoUnit, -1, -1  
			ggoSpread.SSSetProtected C_SoPrice, -1, -1  
			ggoSpread.SSSetProtected C_SoPriceFlag, -1, -1
			ggoSpread.SSSetProtected C_SoPriceAutoChk, -1, -1
			ggoSpread.SSSetProtected C_NetAmt, lRow, lRow
			ggoSpread.SSSetProtected C_SoQty, -1, -1
			ggoSpread.SSSetProtected C_DlvyDt, -1, -1
			ggoSpread.SSSetProtected C_ShipToParty, -1, -1
			ggoSpread.SSSetProtected C_ShipToPartyPopup, -1, -1
			ggoSpread.SSSetProtected C_PlantCd, -1, -1
			ggoSpread.SSSetProtected C_PlantCdPopup, -1, -1
			ggoSpread.SSSetProtected C_SlCd, -1, -1
			ggoSpread.SSSetProtected C_SlCdPopup, -1, -1
			ggoSpread.SSSetProtected C_TolMoreRate, -1, -1
			ggoSpread.SSSetProtected C_TolLessRate, -1, -1
			ggoSpread.SSSetProtected C_TotalAmt, -1, -1
			ggoSpread.SSSetProtected C_NetAmt, -1, -1
			ggoSpread.SSSetProtected C_VatAmt, -1, -1
			ggoSpread.SSSetProtected C_VatType, -1, -1
			ggoSpread.SSSetProtected C_VatTypeNm, -1, -1
			ggoSpread.SSSetProtected C_VatIncFlag, -1, -1
			ggoSpread.SSSetProtected C_VatIncFlagNm, -1, -1
			ggoSpread.SSSetProtected C_VatRate, -1, -1
			ggoSpread.SSSetProtected C_RetType, -1, -1
			ggoSpread.SSSetProtected C_RetTypeNm, -1, -1
			ggoSpread.SSSetProtected C_LotNo, -1, -1
			ggoSpread.SSSetProtected C_LotSeq, -1, -1
			ggoSpread.SSSetProtected C_PreDNNo, -1, -1
			ggoSpread.SSSetProtected C_PreDNSeq, -1, -1
			ggoSpread.SSSetProtected C_DnReqDt, -1, -1
			ggoSpread.SSSetProtected C_BonusQty, -1, -1
			ggoSpread.SSSetProtected C_HsNo, -1, -1
			ggoSpread.SSSetProtected C_HsNoPopup, -1, -1
			ggoSpread.SSSetProtected C_Remark, -1, -1
			ggoSpread.SSSetProtected C_SoUnitPopup, -1, -1
			ggoSpread.SSSetProtected C_VatTypePopup, -1, -1
			ggoSpread.SSSetProtected C_RetTypePopup, -1, -1        
			   
			' ����û���ڰ� �������ں��� ������� �˸�ǥ�� 
			For lRow = 1 To .vspdData.MaxRows          
				  .vspdData.Row = lRow : .vspdData.Col = C_DnReqDt
				  If UniConvDateToYYYYMMDD(.vspdData.Text,parent.gDateFormat,"") < UniConvDateToYYYYMMDD(EndDate,parent.gDateFormat,"") Then Call sprRedComColor(C_DnReqDt,-1,-1)
			Next       

		End If

		.vspdData.ReDraw = True

		If Trim(.RdoConfirm.value) = "N" Then Call SetToolbar("11000000000111")

    End With

End Sub

'========================================================================================================
Sub Form_Load()

	Err.Clear                                                                '��: Clear err status
	Call LoadInfTB19029                                                      '��: Load table , B_numeric_format
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 	
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                             '��: Lock  Suitable  Field

	Call InitSpreadSheet
	Call InitVariables
	Call SetDefaultVal
	
	Call SetToolbar("11000000000011")								 '��: ��ư ���� ���� 	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call CookiePage(0)
End Sub


'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
End Sub


'========================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    On Error Resume Next                                                          '��: If process fails
    Err.Clear																	  '��: Clear error status
    
    FncQuery = False															  '��: Processing is NG    

	ggoSpread.Source = frm1.vspdData 
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")	  '��: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "2")								  '��: Clear Contents  Field
    
    If Not chkField(Document, "1") Then								 '��: This function check required field
       Exit Function
    End If
    
    '------ Developer Coding part (Start ) --------------------------------------------------------------     
    Call InitVariables												 '��: Initializes local global variables

    If DbQuery = False Then                                    '��: Query db data
       Exit Function
    End If
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncQuery = True                                               '��: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement
        
End Function

'========================================================================================================
Function FncSave() 
	Dim IntRetCD 
    
    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

    FncSave = False																  '��: Processing is NG
           
	ggoSpread.Source = frm1.vspdData 
	If ggoSpread.SSCheckChange = False and lgBlnFlgChgValue = False Then								  '��:match pointer
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")				  '��:There is no changed data.  
        Exit Function
    End If  
    
    ggoSpread.Source = frm1.vspdData      
    IF ggoSpread.SSDefaultCheck = False Then								  '��: Check contents area
		Exit Function
    End If
    
    
    '------ Developer Coding part (Start ) -------------------------------------------------------------- 
	If Not chkField(Document, "2")  Then                          
       Exit Function
    End If

    '�������������� ���� ���� ��ǰ������ ��� Lot ��ȣ�� Assign
	If frm1.HRetItemFlag.value = "Y" Then
		Dim iRow

		For iRow = 1 to frm1.vspdData.MaxRows
			frm1.vspdData.Col = 0
			frm1.vspdData.Row = iRow
			If frm1.vspdData.text = ggoSpread.InsertFlag Or frm1.vspdData.text = ggoSpread.UpdateFlag Then
				frm1.vspdData.Col = C_LotNo
				If frm1.vspdData.text = "" Then
					frm1.vspdData.text = "*"
					frm1.vspdData.Col = C_LotSeq
					frm1.vspdData.text = 0
				End If
			End If       
		Next   
	End If  

	Call CheckCreditlimitSvr									'��: �����ϱ��� �����ѵ� üũ 
	    
'------ Developer Coding part (End )   -------------------------------------------------------------- 

    If DbSave = False Then                                                        '��: Query db data
       Exit Function
    End If

    If Err.number = 0 Then	
       FncSave = True                                                             '��: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement
    
End Function


'========================================================================================================
Function FncCancel() 
	Dim iDx
	
    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

    FncCancel = False                                                             '��: Processing is NG
	
	If frm1.vspdData.MaxRows < 1 Then Exit Function
    ggoSpread.Source = Frm1.vspdData	
    ggoSpread.EditUndo  
	'------ Developer Coding part (Start ) --------------------------------------------------------------
    Call TotalSum														'��: Protect system from crashing    
    '------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncCancel = True                                                            '��: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement 
    
End Function


'========================================================================================================
Function FncPrint() 
	On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

    FncPrint = False                                                              '��: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()                                                        '��: Protect system from crashing

    If Err.number = 0 Then	 
       FncPrint = True                                                            '��: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement
End Function


'========================================================================================================
Function FncExcel() 
	On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

    FncExcel = False                                                              '��: Processing is NG

	Call Parent.FncExport(Parent.C_SINGLEMULTI)

    If Err.number = 0 Then	 
       FncExcel = True                                                            '��: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement 
End Function

'========================================================================================================
Function FncFind() 
	On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

    FncFind = False                                                               '��: Processing is NG

	'Call Parent.FncFind(Parent.C_MULTI, True)
	Call parent.FncFind(parent.C_SINGLEMULTI, False)
    If Err.number = 0 Then	 
       FncFind = True                                                             '��: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement  
End Function


'========================================================================================
Function FncSplitColumn()
    
    Dim ACol
    Dim ARow
    Dim iRet
    Dim iColumnLimit

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    iColumnLimit  = C_SoSupplyQty
    '------ Developer Coding part (End )   -------------------------------------------------------------- 
    ACol = Frm1.vspdData.ActiveCol
    ARow = Frm1.vspdData.ActiveRow

    If ACol > iColumnLimit Then
		'Frm1.vspdData.Col = iColumnLimit : Frm1.vspdData.Row = 0  ����??
		iRet = DisplayMsgBox("900030", "X", Trim(frm1.vspdData.Text), "X")
		Exit Function
    End If   
    
    Frm1.vspdData.ScrollBars = parent.SS_SCROLLBAR_NONE    
    
    ggoSpread.Source = Frm1.vspdData    
    ggoSpread.SSSetSplit(ACol)      
    Frm1.vspdData.Col = ACol
    Frm1.vspdData.Row = ARow
   
    Frm1.vspdData.Action = Parent.SS_ACTION_ACTIVE_CELL   
    
    Frm1.vspdData.ScrollBars = parent.SS_SCROLLBAR_BOTH
    
End Function

'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
	
	Call SetQuerySpreadColor(1)    

End Sub

'========================================================================================================
Function FncExit()
	Dim IntRetCD

    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

    FncExit = False                                                               '��: Processing is NG
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")		  '��: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    If Err.number = 0 Then	 
       FncExit = True                                                             '��: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function


'========================================================================================================
Function DbQuery() 

    Dim strVal
	
    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status
    
    DbQuery = False                                                               '��: Processing is NG
	
    Call DisableToolBar(Parent.TBC_QUERY)                                                '��: Disable Query Button Of ToolBar
    Call LayerShowHide(1)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    If lgIntFlgMode = parent.OPMD_UMODE Then    
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001				'��: �����Ͻ� ó�� ASP�� ���� 
		strVal = strVal & "&txtConSoNo=" & Trim(frm1.txtHSoNo.value)        '��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	Else
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001        
		strVal = strVal & "&txtConSoNo=" & Trim(frm1.txtConSoNo.value)     
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	End If 
	'--------- Developer Coding Part (End) ------------------------------------------------------------
    Call RunMyBizASP(MyBizASP, strVal)                                           '��:  Run biz logic

    If Err.number = 0 Then	 
       DbQuery = True                                                             '��: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement
End Function


'========================================================================================================
Function DbSave() 

    Dim lRow        
    Dim lGrpCnt     
	Dim strVal

    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

    DbSave = False                                                                '��: Processing is NG

    Call DisableToolBar(Parent.TBC_SAVE)                                   '��: Disable Save Button Of ToolBar
    Call LayerShowHide(1)                                                         '��: Show Processing Message

	Frm1.txtMode.value        = Parent.UID_M0002                                  '��: Delete		
	
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	ggoSpread.Source = frm1.vspdData
	
	lGrpCnt = 0    
	strVal = ""
	
	With frm1	

		For lRow = 1 To .vspdData.MaxRows
			.vspdData.Row = lRow
			.vspdData.Col = 0
	
			Select Case .vspdData.Text
				Case ggoSpread.UpdateFlag       '��: ���� 
					strVal = strVal & "U" & parent.gColSep & lRow & parent.gColSep'��: U=Update
			End Select

			Select Case .vspdData.Text
			
				Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag,ggoSpread.DeleteFlag
					.vspdData.Col = C_SoSeq        '--- ���ּ���              
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep                    
              
					.vspdData.Col = C_ItemCd       '--- ǰ�� 
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep             
              
					.vspdData.Col = C_SoUnit       '--- ���ִ���  
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep              
              
					.vspdData.Col = C_TrackingNo   '--- Tracking No  
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep              
              
					.vspdData.Col = C_SoQty        '--- ���ּ��� 
					IF UNIConvNum(Trim(.vspdData.Text), 0) <= 0 Then
					   Call DisplayMsgBox("203233", "X", "X", "X")              
					   Exit Function
					Else
					   strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & parent.gColSep 
					END IF                              
              
					.vspdData.Col = C_SoPrice     '--- ���ִܰ� 
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & parent.gColSep              
              
					.vspdData.Col = C_SoPriceFlag  '--- ���� ���ܰ�/���ܰ� 
					
					Select Case Trim(.vspdData.TypeComboBoxCurSel)
					
						Case 1  '���ܰ� 
								strVal = strVal & "Y" & parent.gColSep    
						Case 0  '���ܰ� 
								strVal = strVal & "N" & parent.gColSep    
						Case Else
								MsgBox "���ܰ�/���ܰ� ���� �����ϴ�.", vbExclamation, parent.gLogoName 
								frm1.vspdData.Row = lRow
								frm1.vspdData.Action = 0
								'--�۾��� ǥ��ȭ�� �� ���콺 ����Ʈ ����      
								Call BtnDisabled(False)
								Call LayerShowHide(0)
								Exit Function
					End Select
     
					.vspdData.Col = C_NetAmt     '--- ���ֱݾ� 
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & parent.gColSep                            
              
					.vspdData.Col = C_VatAmt     '--- VAT �ݾ� 
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & parent.gColSep              
              
					.vspdData.Col = C_PlantCd     '--- �����ڵ� 
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep     
              
					.vspdData.Col = C_DlvyDt     '--- ������ 
					strVal = strVal & UNIConvDate(Trim(.vspdData.Text)) & parent.gColSep              
              
					.vspdData.Col = C_ShipToParty  '--- ��ǰó 
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep              
              
					.vspdData.Col = C_HsNo      '--- HS��ȣ 
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep              
              
					.vspdData.Col = C_TolMoreRate  '--- �����������(+)
					strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & parent.gColSep     
              
					.vspdData.Col = C_TolLessRate  '--- �����������(-)
					strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & parent.gColSep              
              
					.vspdData.Col = C_VatType     '--- VAT���� 
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep              
              
					.vspdData.Col = C_VatRate     '--- VAT�� 
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & parent.gColSep     
              
					.vspdData.Col = C_VatIncFlag   '--- VAT ���Ա��� 
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep     
              
					.vspdData.Col = C_RetType     '--- ��ǰ���� 
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep     
              
					.vspdData.Col = C_LotNo     '--- Lot ��ȣ 
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep    
              
					.vspdData.Col = C_LotSeq     '--- Lot ���� 
					strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & parent.gColSep     
              
					.vspdData.Col = C_PreDnNo     '--- ��ǰ���Ϲ�ȣ 
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep     
              
					.vspdData.Col = C_PreDnSeq     '--- ��ǰ���ϼ��� 
					strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & parent.gColSep     
              
					.vspdData.Col = C_BonusQty     '--- ������ 
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & parent.gColSep     
              
					.vspdData.Col = C_SlCd      '--- â���ڵ� 
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep     
              
					.vspdData.Col = C_Remark     '--- ��� 
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep    
					      
					.vspdData.Col = C_PreSoNo     '--- ��ǰ���ֹ�ȣ 
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep     
                    
				    .vspdData.Col = C_PreSoSeq     '--- ��ǰ���ּ��� 
				    strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & parent.gColSep  

					.vspdData.Col = C_CustPoSeq        '--- ���ּ���            
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep '30
					
					strVal = strVal & ext1_qty & parent.gColSep
				    
				    strVal = strVal & ext2_qty & parent.gColSep
				    
				    strVal = strVal & ext3_qty & parent.gColSep
				    
				    strVal = strVal & ext1_amt & parent.gColSep
				    
				    strVal = strVal & ext2_amt & parent.gColSep
				    
				    strVal = strVal & ext3_amt & parent.gColSep
				    
				    strVal = strVal & ext1_cd  & parent.gColSep
				    
				    strVal = strVal & ext2_cd  & parent.gColSep
				    
				    strVal = strVal & ext3_cd  & parent.gRowSep
					
				    lGrpCnt = lGrpCnt + 1 
			End Select       
        Next

		.txtMaxRows.value = lGrpCnt
		.txtSpread.value = strVal	
   
	End With
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)          '��: �����Ͻ� ASP �� ���� 
		
    '------ Developer Coding part (End )   -------------------------------------------------------------- 

    If Err.number = 0 Then	 
       DbSave = True                                                             '��: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement 
      
End Function

'========================================================================================================
Function DbDelete() 
    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

    DbDelete = False                                                              '��: Processing is NG
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	 
       DbDelete = True                                                             '��: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement
End Function


'========================================================================================================
Sub DbQueryOk()         
	
	On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

    lgIntFlgMode = Parent.OPMD_UMODE    
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
  
	frm1.btnAvlStkRef.disabled = False

	If Trim(frm1.txtHPreSONo.value) <> "" And UCase(Trim(frm1.HRetItemFlag.value)) = "Y" Then
		Call SetToolbar("11001001000111")
	ElseIf Trim(frm1.txtHPreSONo.value) = "" And UCase(Trim(frm1.HRetItemFlag.value)) = "Y" Then
		Call SetToolbar("11001001000111")
	ElseIf UCase(Trim(frm1.HRetItemFlag.value)) <> "Y" Then
		Call SetToolbar("11001001000111")
	Else
		Call SetToolbar("11001001000111")
	End If
	
	frm1.vspdData.Focus
	
	Call SetQuerySpreadColor(1)    
	Call TotalSum

	If frm1.RdoConfirm.value = "N" Then
		Call ggoOper.SetReqAttr(frm1.txtSoDt, "Q") 
		Call ggoOper.SetReqAttr(frm1.txtDealType, "Q") 
		Call ggoOper.SetReqAttr(frm1.txtSalesGrp, "Q") 
		Call ggoOper.SetReqAttr(frm1.txtPaymeth, "Q") 
		Call ggoOper.SetReqAttr(frm1.txtRemark, "Q") 
		Call ggoOper.SetReqAttr(frm1.txtXchgRate, "Q") 
	Else
		Call ggoOper.SetReqAttr(frm1.txtSoDt, "N") 
		Call ggoOper.SetReqAttr(frm1.txtDealType, "D") 
		Call ggoOper.SetReqAttr(frm1.txtSalesGrp, "N") 
		Call ggoOper.SetReqAttr(frm1.txtPaymeth, "D") 
		Call ggoOper.SetReqAttr(frm1.txtRemark, "D") 
		Call ggoOper.SetReqAttr(frm1.txtXchgRate, "D") 
	End If
	
	If Trim(frm1.txtCurrency.value) = parent.gCurrency Then
		Call ggoOper.SetReqAttr(frm1.txtXchgRate, "Q") 
	End If
	
	lgBlnFlgChgValue = False
      
	Call ButtonVisible(1)
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    Set gActiveElement = document.ActiveElement  
    
End Sub

'========================================================================================================
Sub DbSaveOk()               
	On Error Resume Next                                                   '��: If process fails
    Err.Clear                                                              '��: Clear error status

    Call InitVariables													   '��: Initializes local global variables
    
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
    frm1.txtConSoNo.value = frm1.txtHSoNo.value
	frm1.vspdData.MaxRows = 0
    Call MainQuery()
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    
    Set gActiveElement = document.ActiveElement
    
End Sub


'========================================================================================================
Function DbDeleteOk()            
    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    Set gActiveElement = document.ActiveElement 
End Function

'========================================================================================================
Function OpenAvalStockRef()
	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(4)
 
	If UCase(Trim(frm1.txtConSoNo.value)) = "" Then
		Call DisplayMsgBox("900002", "x", "x", "x") 
		Exit Function
	End If

	ggoSpread.Source = frm1.vspdData
  
	If frm1.vspdData.Maxrows < 1 Then
		Call DisplayMsgBox("209001", "x", "x", "x")
		Exit Function
	End If 
  
	frm1.vspdData.Col = C_ItemCd
	frm1.vspdData.Row = frm1.vspdData.ActiveRow

	arrParam(0) = frm1.vspdData.Text

	If arrParam(0) = "" Then
		Call DisplayMsgBox("202250", "x", "x", "x")				'��: "Will you destory previous data" 
		Exit Function
	End If 
  
	frm1.vspdData.Col = C_ItemName   
 
	arrParam(1) = frm1.vspdData.Text

	frm1.vspdData.Col = C_PlantCd
 
	arrParam(2) = frm1.vspdData.Text

	frm1.vspdData.Col = C_PlantNm
  
	arrParam(3) = frm1.vspdData.Text

	iCalledAspName = AskPRAspName("s1912ra1")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s1912ra1", "x")
		IsOpenPop = False
		exit Function
	end if

	IsOpenPop = True
 
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
	  "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
  
	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetSODtlRef(arrRet)
	End If 
	
 End Function


'========================================================================================================
Function OpenStockDtlRef()
	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(5)
 
	If UCase(Trim(frm1.txtConSoNo.value)) = "" Then
		Call DisplayMsgBox("900002", "x", "x", "x")		 '��: "Will you destory previous data"
		Exit Function
	End If

	ggoSpread.Source = frm1.vspdData
  
	If frm1.vspdData.Maxrows < 1 Then
		Call DisplayMsgBox("209001", "x", "x", "x")		 '��: "Will you destory previous data"
		Exit Function
	End If 
  
	frm1.vspdData.Col = C_ItemCd
	frm1.vspdData.Row = frm1.vspdData.ActiveRow

	arrParam(0) = frm1.vspdData.Text

	If arrParam(0) = "" Then
		Call DisplayMsgBox("202250", "x", "x", "x") 
		Exit Function
	End If 
  
	frm1.vspdData.Col = C_ItemName 
	arrParam(1) = frm1.vspdData.Text

	frm1.vspdData.Col = C_PlantCd 
	arrParam(2) = frm1.vspdData.Text

	frm1.vspdData.Col = C_PlantNm  
	arrParam(3) = frm1.vspdData.Text

	frm1.vspdData.Col = C_SlCd 
	arrParam(4) = frm1.vspdData.Text

	frm1.vspdData.Col = C_SlNm  
	arrParam(5) = frm1.vspdData.Text

	'20021227 kangjungu dynamic popup
	iCalledAspName = AskPRAspName("s1912ra2")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s1912ra2", "x")
		IsOpenPop = False
		exit Function
	end if

	IsOpenPop = True
 
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
	  "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
  
	IsOpenPop = False

	If arrRet(0, 0) = "" Then
		Exit Function
	End If 
  
End Function


'========================================================================================================
Function OpenHdr(ByVal iOption)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iOption	
	Case 0												
		If frm1.txtDealType.readOnly = True Then
			IsOpenPop = False
			Exit Function
		End If

		arrParam(0) = "�Ǹ�����"					
		arrParam(1) = "B_MINOR"							
		arrParam(2) = Trim(frm1.txtDealType.value)		
		arrParam(3) = Trim(frm1.txtDealTypeNm.value)	
		arrParam(4) = "MAJOR_CD=" & FilterVar("S0001", "''", "S") & ""				
		arrParam(5) = "�Ǹ�����"					
		
	    arrField(0) = "MINOR_CD"						
	    arrField(1) = "MINOR_NM"					
	    
	    arrHeader(0) = "�Ǹ�����"				
	    arrHeader(1) = "�Ǹ�������"				
		
		frm1.txtDealType.focus 
		
	Case 1
		If frm1.txtSalesGrp.readOnly = True Then
			IsOpenPop = False
			Exit Function
		End If
		arrParam(0) = "�����׷�"					
		arrParam(1) = "B_SALES_GRP"						
		arrParam(2) = Trim(frm1.txtSalesGrp.value)		
		arrParam(3) = Trim(frm1.txtSalesGrpNm.value)
		arrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "					
		arrParam(5) = "�����׷�"				
		
	    arrField(0) = "SALES_GRP"						
	    arrField(1) = "SALES_GRP_NM"					
	    
	    arrHeader(0) = "�����׷�"				
	    arrHeader(1) = "�����׷��"					
		
		frm1.txtSalesGrp.focus 
		
	Case 2												
		If frm1.txtPaymeth.readOnly = True Then
			IsOpenPop = False
			Exit Function
		End If

		arrParam(0) = "�������"					
		arrParam(1) = "B_MINOR"							
		arrParam(2) = Trim(frm1.txtPaymeth.value)		
		arrParam(3) = Trim(frm1.txtPaymethNm.value)
		arrParam(4) = "MAJOR_CD=" & FilterVar("B9004", "''", "S") & ""				
		arrParam(5) = "�������"				
		
	    arrField(0) = "MINOR_CD"						
	    arrField(1) = "MINOR_NM"					
	    
	    arrHeader(0) = "�������"					
	    arrHeader(1) = "���������"					
		
		frm1.txtPaymeth.focus 			
	
	End Select

	arrParam(3) = ""			'��: [Condition Name Delete]

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetHdr(arrRet,iOption)
	End If	
End Function


'===========================================================================
Function OpenItem(ByVal strCode)
	Dim iCalledAspName
	Dim arrParam(1)
	Dim strRet

	arrParam(0) = strCode
	frm1.vspdData.Col = C_PlantCd
	arrParam(1) = frm1.vspdData.text 

	If IsOpenPop = True Then Exit Function
	  
	'20021227 kangjungu dynamic popup
	iCalledAspName = AskPRAspName("s3112pa2")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3112pa2", "x")
		IsOpenPop = False
		exit Function
	end if

	IsOpenPop = True
	 
	strRet = window.showModalDialog(iCalledAspName, Array(arrParam), _
	 "dialogWidth=820px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If strRet(0) = "" Then
		Exit Function
	Else
		frm1.vspdData.Col = C_ItemCd
		frm1.vspdData.Text = strRet(0)
		frm1.vspdData.Col = C_ItemName
		frm1.vspdData.Text = strRet(1)
		frm1.vspdData.Col = C_PlantCd
		frm1.vspdData.Text = strRet(2)
		Call vspdData_Change(C_ItemCd, frm1.vspdData.Row)  ' ������ �о�ٰ� �˷��� 
	End If 

End Function 

'===========================================================================
Function OpenConSoDtl()
	Dim iCalledAspName
	Dim strRet

	If IsOpenPop = True Then Exit Function
	  
	'20021227 kangjungu dynamic popup
	iCalledAspName = AskPRAspName("s3111pa1")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3111pa1", "x")
		IsOpenPop = False
		exit Function
	end if

	IsOpenPop = True
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent, "STO"), _
	 "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If strRet = "" Then
		Exit Function
	Else
		frm1.txtConSoNo.value = strRet
		frm1.txtConSoNo.focus
	End If 

End Function 

'===========================================================================
Function OpenSoDtl(Byval strCode, Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim OriginCol,TempCd

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	 Select Case iWhere
	 Case 0 'ǰ�� 
		arrParam(1) = "b_item item, b_plant plant, b_item_by_plant item_plant"   
		arrParam(2) = strCode               
		arrParam(4) = "item.item_cd=item_plant.item_cd and plant.plant_cd=item_plant.plant_cd"
		arrParam(5) = "ǰ��"     
	 
		arrField(0) = "item.item_cd"    
		arrField(1) = "item.item_nm"    
		arrField(2) = "plant.plant_cd"    
		arrField(3) = "plant.plant_nm" 
    
		arrHeader(0) = "ǰ��"      
		arrHeader(1) = "ǰ���"      
		arrHeader(2) = "����"      
		arrHeader(3) = "�����"      

	Case 1 '���� 
		arrParam(1) = "b_unit_of_measure"    
		arrParam(2) = strCode       
		arrParam(4) = ""      		 
		arrParam(5) = "����"   
		    
		arrField(0) = "unit"       
		arrField(1) = "unit_nm"   	
		  
		arrHeader(0) = "����"      
		arrHeader(1) = "������"      
	
	Case 3 '��ǰó 
		arrParam(1) = "b_biz_partner bp, b_biz_partner_ftn bp_ftn"   
		arrParam(2) = strCode            
		arrParam(4) = "bp.bp_cd=bp_ftn.partner_bp_cd and bp_ftn.bp_cd= " + FilterVar(frm1.txtSoldToParty.value, "''", "S") + " and bp_ftn.partner_ftn = " & FilterVar("SSH", "''", "S") & " and bp_ftn.usage_flag = " & FilterVar("Y", "''", "S") & " "     <%' Where Condition%>
		arrParam(5) = "��ǰó"      
 
		arrField(0) = "bp_ftn.partner_bp_cd"   
		arrField(1) = "bp.bp_nm"      
		  
		arrHeader(0) = "��ǰó"      
		arrHeader(1) = "��ǰó��"    
		 
	Case 4 'HS��ȣ 
		arrParam(1) = "b_hs_code"      
		arrParam(2) = strCode       
		arrParam(4) = ""        
		arrParam(5) = "HS��ȣ"      
 
		arrField(0) = "hs_cd"       
		arrField(1) = "hs_nm"       
		  
		arrHeader(0) = "HS��ȣ"      
		arrHeader(1) = "HS��ȣ��"     
	
	Case 5 '���� 
		With frm1
			OriginCol = .vspdData.Col
			.vspdData.Col = C_ItemCd
			TempCd = .vspdData.Text
			.vspdData.Col = OriginCol
		End With
		arrParam(1) = "b_plant plant, b_item_by_plant item_plant"    
		arrParam(2) = strCode             
		arrParam(4) = "plant.plant_cd=item_plant.plant_cd and item_plant.item_cd =  " + FilterVar(TempCd, "''", "S") 
		arrParam(5) = "����"      
 
		arrField(0) = "plant.plant_cd"     
		arrField(1) = "plant.plant_nm"     
		    
		arrHeader(0) = "����"      
		arrHeader(1) = "�����"      
  
	Case 6 'â�� 
		With frm1
			OriginCol = .vspdData.Col
			.vspdData.Col = C_PlantCd
			TempCd = .vspdData.Text
			.vspdData.Col = OriginCol
		End With
		arrParam(1) = "b_storage_location"    
		arrParam(2) = strCode       
		arrParam(4) = "plant_cd = " + FilterVar(TempCd, "''", "S") 
		arrParam(5) = "â��"      
 
		arrField(0) = "sl_cd"       
		arrField(1) = "sl_nm"       
		  
		arrHeader(0) = "â��"      
		arrHeader(1) = "â���"      

	Case 7 'VAT���� 
		arrParam(1) = "B_MINOR Minor,B_CONFIGURATION Config"
		arrParam(2) = strCode        
		arrParam(4) = "Minor.MAJOR_CD=" & FilterVar("B9001", "''", "S") & " And Config.MAJOR_CD = Minor.MAJOR_CD" _
		    & " And Config.MINOR_CD = Minor.MINOR_CD" _
		    & " And Config.SEQ_NO = 1"   
		arrParam(5) = "VAT����"      

		arrField(0) = "Minor.MINOR_CD"      
		arrField(1) = "Minor.MINOR_NM"      
		arrField(2) = "Config.REFERENCE"     
			         
		arrHeader(0) = "VAT����"      
		arrHeader(1) = "VAT������"      
		arrHeader(2) = "VAT��" 
		     
	Case 8 
		arrParam(1) = "B_MINOR"      
		arrParam(2) = strCode       
		arrParam(4) = "MAJOR_CD=" & FilterVar("B9017", "''", "S") & ""        
		arrParam(5) = "��ǰ����"      
 
		arrField(0) = "Minor_cd"       
		arrField(1) = "Minor_nm"       
			   
		arrHeader(0) = "��ǰ����"      
		arrHeader(1) = "��ǰ������" 
		    
	Case 9
		arrParam(0) = "VAT���Ա���"    
		arrParam(1) = "B_MINOR"      
		arrParam(2) = strCode
		arrParam(4) = "MAJOR_CD=" & FilterVar("S4035", "''", "S") & ""    
		arrParam(5) = "VAT���Ա���"    
	 
		arrField(0) = "MINOR_CD"     
		arrField(1) = "MINOR_NM"     
		        
		arrHeader(0) = "VAT���Ա���"   
		arrHeader(1) = "VAT���Ա��и�"   
	End Select

	arrParam(0) = arrParam(5)       

	Select Case iWhere
	Case 0 <% 'ǰ�� %>
	 arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	  "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Case Else
	 arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End Select

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetSoDtl(arrRet, iWhere)
	End If 
 
End Function


'========================================================================================================
Function SetHdr(Byval arrRet,ByVal iOption)

If arrRet(0) <> "" Then 
	Select Case iOption
	Case 0												' �ŷ����� 
		frm1.txtDealType.value = arrRet(0)
		frm1.txtDealTypeNm.value = arrRet(1)
	Case 1												' �����׷� 
		frm1.txtSalesGrp.value = arrRet(0)
		frm1.txtSalesGrpNm.value = arrRet(1)
	Case 2												' ������� 
		frm1.txtPaymeth.value = arrRet(0)
		frm1.txtPaymethNm.value = arrRet(1)	
	End Select

	lgBlnFlgChgValue = True

End If

End Function


'========================================================================================================
Function SetSODtlRef(arrRet)
	Dim intRtnCnt, strData
	Dim TempRow, I, j
	Dim blnEqualFlg
	Dim intLoopCnt
	Dim intCnt
	Dim strSoNo,strSoSeqNo
	Dim strSOJungBokMsg

	With frm1
		.vspdData.focus
		ggoSpread.Source = .vspdData
		.vspdData.ReDraw = False 

		TempRow = .vspdData.MaxRows           
		intLoopCnt = Ubound(arrRet, 1)         

		strSOJungBokMsg = ""
		 
		For intCnt = 1 to intLoopCnt + 1
			blnEqualFlg = False

			If TempRow <> 0 Then

				strSoNo=""
				strSoSeqNo=""

					For j = 1 To TempRow

						<% '���ֹ�ȣ %>
						.vspdData.Row = j
						.vspdData.Col = C_PreSoNo
						strSoNo = .vspdData.text

						If strSoNo = arrRet(intCnt - 1, 0) Then

							<% '���ּ��� %>
							.vspdData.Row = j
							.vspdData.Col = C_PreSoSeq
							strSoSeqNo = .vspdData.text

							If strSoSeqNo = arrRet(intCnt - 1, 1) Then
								blnEqualFlg = True
								strSOJungBokMsg = strSOJungBokMsg & Chr(13) & strSoNo & "-" & strSoSeqNo
								Exit For
							End If

						End If
					Next
			End If

			If blnEqualFlg = False Then
				.vspdData.MaxRows = .vspdData.MaxRows + 1
				.vspdData.Row = CLng(TempRow) + CLng(intCnt)

				.vspdData.Col = 0
				.vspdData.Text = ggoSpread.InsertFlag

				.vspdData.Col = C_PreSoNo
				.vspdData.text = arrRet(intCnt - 1, 0)
				.vspdData.Col = C_PreSoSeq
				.vspdData.text = arrRet(intCnt - 1, 1)

				.vspdData.Col = C_ItemCd
				.vspdData.text = arrRet(intCnt - 1, 6)
				.vspdData.Col = C_ItemName
				.vspdData.text = arrRet(intCnt - 1, 7)
				.vspdData.Col = C_SoUnit
				.vspdData.text = arrRet(intCnt - 1, 8)
				.vspdData.Col = C_SoPrice
				.vspdData.text = arrRet(intCnt - 1, 9)        
				.vspdData.Col = C_SoQty
				.vspdData.text = arrRet(intCnt - 1, 14)
				.vspdData.Col = C_NetAmt
				.vspdData.text = arrRet(intCnt - 1, 12)
				.vspdData.Col = C_VATType
				.vspdData.text = arrRet(intCnt - 1, 13)
				.vspdData.Col = C_TrackingNo
				.vspdData.text = arrRet(intCnt - 1, 15)
    
				'###### 2001_11_28 ��ǰ�϶��� �������ڷ�. #######
				.vspdData.Col = C_DlvyDt
				'.vspdData.text = arrRet(intCnt - 1, 16)
				.vspdData.text = Trim(frm1.txtSoDt.value)
				'################################################

				.vspdData.Col = C_ShipToParty
				.vspdData.text = arrRet(intCnt - 1, 17)
				.vspdData.Col = C_SoPriceFlag

				If Trim(arrRet(intCnt - 1, 18)) = "Y" Then
					.vspdData.text = "���ܰ�"
				ElseIf  Trim(arrRet(intCnt - 1, 18)) = "N" Then
					.vspdData.text = "���ܰ�"
				End If      

				.vspdData.Col = C_HsNo
				.vspdData.text = arrRet(intCnt - 1, 19)
				.vspdData.Col = C_TolMoreRate
				.vspdData.text = arrRet(intCnt - 1, 20)
				.vspdData.Col = C_TolLessRate
				.vspdData.text = arrRet(intCnt - 1, 21)
				.vspdData.Col = C_PlantCd
				.vspdData.text = arrRet(intCnt - 1, 22)
				.vspdData.Col = C_SlCd
				.vspdData.text = arrRet(intCnt - 1, 23)
				.vspdData.Col = C_Remark
				.vspdData.text = arrRet(intCnt - 1, 24)
				.vspdData.Col = C_PreDnNo
				.vspdData.text = arrRet(intCnt - 1, 25)
				.vspdData.Col = C_PreDnSeq
				.vspdData.text = arrRet(intCnt - 1, 26)
				.vspdData.Col = C_LotNo
				.vspdData.text = arrRet(intCnt - 1, 27)
				.vspdData.Col = C_LotSeq
				.vspdData.text = arrRet(intCnt - 1, 28)
				.vspdData.Col = C_VatIncFlag
				.vspdData.text = arrRet(intCnt - 1, 29)
				.vspdData.Col = C_VatIncFlagNm

				If Trim(arrRet(intCnt - 1, 29)) = "1" Then
					 .vspdData.text = "����"
				ElseIf  Trim(arrRet(intCnt - 1, 29)) = "2" Then
					 .vspdData.text = "����"
				End If      

				SetSpreadColor CLng(TempRow) + CLng(intCnt),CLng(TempRow) + CLng(intCnt)
				SetVatType (intCnt)
				lgBlnFlgChgValue = True
			End If
		Next
		
		.vspdData.ReDraw = True
		
		Call JungBokMsg(strSOJungBokMsg,"���ֹ�ȣ" & "-" & "���ּ���")
		Call QtyPriceChange()
		
	End With
	
End Function


'========================================================================================================
Function SetSoDtl(Byval arrRet,ByVal iWhere)

	With frm1

		Select Case iWhere
		Case 0 'ǰ�� 
		 .vspdData.Col = C_ItemCd
		 .vspdData.Text = arrRet(0)
		 .vspdData.Col = C_ItemName
		 .vspdData.Text = arrRet(1)
		 .vspdData.Col = C_PlantCd
		 .vspdData.Text = arrRet(2)
		 Call vspdData_Change(C_ItemCd, .vspdData.Row)  
		 
		Case 1 '���� 
		 .vspdData.Col = C_SoUnit
		 .vspdData.Text = arrRet(0)
		 Call vspdData_Change(C_SoUnit, .vspdData.Row)  
		 
		Case 2 '������ 
	 
		Case 3 '��ǰó 
		 .vspdData.Col = C_ShipToParty
		 .vspdData.Text = arrRet(0)
		 Call vspdData_Change(C_ShipToParty, .vspdData.Row) 
		 
		Case 4 'HS��ȣ 
		 .vspdData.Col = C_HsNo
		 .vspdData.Text = arrRet(0)
		 Call vspdData_Change(C_HsNo, .vspdData.Row)   
		 
		Case 5 '���� 
		 .vspdData.Col = C_PlantCd
		 .vspdData.Text = arrRet(0)
		 Call vspdData_Change(C_PlantCd, .vspdData.Row)  
		 
		Case 6 'â�� 
		 .vspdData.Col = C_SlCd
		 .vspdData.Text = arrRet(0)
		 Call vspdData_Change(C_SlCd, .vspdData.Row)   
	 
		Case 7 'VAT���� 
		 .vspdData.Col = C_VatType
		 .vspdData.Text = arrRet(0)
		 .vspdData.Col = C_VatTypeNm
		 .vspdData.Text = arrRet(1)
		 .vspdData.Col = C_VatRate
		 .vspdData.text = UNIFormatNumber(UNICDbl(arrRet(2)), ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
		 Call vspdData_Change(C_VatType, .vspdData.Row)   
	 
		Case 8 '��ǰ���� 
		 .vspdData.Col = C_RetType
		 .vspdData.Text = arrRet(0)
		 .vspdData.Col = C_RetTypeNm
		 .vspdData.Text = arrRet(1)
		 Call vspdData_Change(C_RetType, .vspdData.Row)   
	 
		Case 9 'VAT ���Ա��� 
		 .vspdData.Col = C_VatIncFlag
		 .vspdData.Text = arrRet(0)
		 .vspdData.Col = C_VatIncFlagNm
		 .vspdData.Text = arrRet(1)
		 Call vspdData_Change(C_VatIncFlag, .vspdData.Row)   
	 
		Case Else
		 Exit Function
		End Select
	 
	End With

	lgBlnFlgChgValue = True
 
End Function

'========================================================================================================
Sub txtSoDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtSoDt.Action = 7
		Call SetFocusToDocument("M")
        frm1.txtSoDt.Focus
	End If
End Sub


'========================================================================================================
Function JungBokMsg(strJungBok,strID)

	Dim strJugBokMsg

	If Len(Trim(strJungBok)) Then strJungBok = strID & Chr(13) & String(30,"=") & strJungBok
	If Len(Trim(strJungBok)) Then strJugBokMsg = strJungBok & Chr(13) & Chr(13)
	If Len(Trim(strJugBokMsg)) Then
		strJugBokMsg = strJugBokMsg & "�̹� ������ ��ȣ�� ������ �����մϴ�"
		MsgBox strJugBokMsg, vbInformation, parent.gLogoName
	End If

End Function


'========================================================================================================
Sub HideLotRetField()
	If frm1.HRetItemFlag.value = "N" Then
		frm1.vspdData.Col = C_PreDnNo  :    frm1.vspdData.ColHidden = True
		frm1.vspdData.Col = C_PreDnSeq  :    frm1.vspdData.ColHidden = True
		frm1.vspdData.Col = C_LotNo   :    frm1.vspdData.ColHidden = True
		frm1.vspdData.Col = C_LotSeq  :    frm1.vspdData.ColHidden = True
		frm1.vspdData.Col = C_RetType  :    frm1.vspdData.ColHidden = True
		frm1.vspdData.Col = C_RetTypePopup :    frm1.vspdData.ColHidden = True
		frm1.vspdData.Col = C_RetTypeNm  :    frm1.vspdData.ColHidden = True
	Else
		frm1.vspdData.Col = C_PreDnNo  :    frm1.vspdData.ColHidden = False
		frm1.vspdData.Col = C_PreDnSeq  :    frm1.vspdData.ColHidden = False
		frm1.vspdData.Col = C_LotNo   :    frm1.vspdData.ColHidden = False
		frm1.vspdData.Col = C_LotSeq  :    frm1.vspdData.ColHidden = False
		frm1.vspdData.Col = C_RetType  :    frm1.vspdData.ColHidden = False
		frm1.vspdData.Col = C_RetTypePopup :    frm1.vspdData.ColHidden = False
		frm1.vspdData.Col = C_RetTypeNm  :    frm1.vspdData.ColHidden = False
    End If  
End Sub

'========================================================================================================
Sub InitCollectType()
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	Dim iCodeArr, iNameArr, iRateArr

	Err.Clear

	Call CommonQueryRs(" Minor.MINOR_CD,  Minor.MINOR_NM, Config.REFERENCE ", " B_MINOR Minor,B_CONFIGURATION Config ", " Minor.MAJOR_CD=" & FilterVar("B9001", "''", "S") & " And Config.MAJOR_CD = Minor.MAJOR_CD And Config.MINOR_CD = Minor.MINOR_CD And Config.SEQ_NO = 1 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	iCodeArr = Split(lgF0, Chr(11))
	iNameArr = Split(lgF1, Chr(11))
	iRateArr = Split(lgF2, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.description 
		Err.Clear 
		Exit Sub
	End If

	Redim arrCollectVatType(UBound(iCodeArr) - 1, 2)

	For i = 0 to UBound(iCodeArr) - 1
		arrCollectVatType(i, 0) = iCodeArr(i)
		arrCollectVatType(i, 1) = iNameArr(i)
		arrCollectVatType(i, 2) = iRateArr(i)
	Next
End Sub

'========================================================================================================
Sub GetCollectTypeRef(ByVal CollectType, ByRef VatTypeNm, ByRef VatRate)
	Dim iCnt

	For iCnt = 0 To Ubound(arrCollectVatType)
		If arrCollectVatType(iCnt, 0) = CollectType Then
			VatTypeNm = arrCollectVatType(iCnt, 1)
			VatRate   = UNICDbl(arrCollectVatType(iCnt, 2))
			Exit Sub
		End If
	Next

	VatTypeNm = ""
	VatRate = ""
End Sub

'========================================================================================================
Sub SetVatType(Row)
	Dim VatType, VatTypeNm, VatRate

	frm1.vspdData.Col = C_VatType : 
	VatType = frm1.vspdData.text
 
	Call InitCollectType
	Call GetCollectTypeRef(VatType, VatTypeNm, VatRate)
	frm1.vspdData.Col = C_VatTypeNm :  frm1.vspdData.Text = VatTypeNm
	frm1.vspdData.Col = C_VatRate :  frm1.vspdData.Text = UNIFormatNumber(VatRate, ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
 
	Call QtyPriceChange()
End Sub

'========================================================================================================
Function QtyPriceChange()

	Dim SoQty, SoPrice, NetAmt, VatAmt, CalSOAmt, VatRate, TotalAmt

	frm1.vspdData.Col = C_SoQty
	If Trim(frm1.vspdData.Text) = "" Or isnull(frm1.vspdData.Text) Then
	 SoQty = 0
	Else
	 SoQty = UNICDbl(frm1.vspdData.Text)
	End If

	frm1.vspdData.Col = C_SoPrice
	If Trim(frm1.vspdData.Text) = "" Or isnull(frm1.vspdData.Text) Then
	 SoPrice = 0
	Else
	 SoPrice = UNICDbl(frm1.vspdData.Text)
	End If

	frm1.vspdData.Col = C_VatRate
	
	If Trim(frm1.vspdData.Text) = "" Or isnull(frm1.vspdData.Text) Then
		VatRate = 0
	Else
		VatRate = UNIFormatNumber(UNICDbl(frm1.vspdData.text), ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
	End If

	CalSOAmt = UNIFormatNumberByCurrecny(SoQty * SoPrice,frm1.txtCurrency.value,parent.ggAmtOfMoneyNo)

	frm1.vspdData.Col = C_VatIncFlag
	If Trim(frm1.vspdData.Text) = "2" Then
		VatAmt = uniFormatNumberByTax(UNICDbl(CalSOAmt) * (UNICDbl(VatRate) / (100 + UNICDbl(VatRate))), frm1.txtCurrency.value, parent.ggAmtOfMoneyNo)
		NetAmt = UNIFormatNumberByCurrecny(UNICDbl(CalSOAmt) - UNICDbl(VatAmt), frm1.txtCurrency.value,parent.ggAmtOfMoneyNo)
		TotalAmt = UNIFormatNumberByCurrecny(UNICDbl(NetAmt) + UNICDbl(VatAmt), frm1.txtCurrency.value,parent.ggAmtOfMoneyNo)
	Else
		VatAmt = uniFormatNumberByTax(UNICDbl(CalSOAmt) * UNICDbl(VatRate)/100, frm1.txtCurrency.value, parent.ggAmtOfMoneyNo)
		NetAmt = CalSOAmt
		TotalAmt = NetAmt
	End If 
 
	frm1.vspdData.Col = C_NetAmt
	frm1.vspdData.Text = NetAmt
	frm1.vspdData.Col = C_VatAmt
	frm1.vspdData.Text = VatAmt
	frm1.vspdData.Col = C_TotalAmt
	frm1.vspdData.Text = TotalAmt

	Call TotalSum

End Function

'========================================================================================================
Function TotalAmtChange()

	Dim NetAmt, VatAmt, CalSOAmt, VatRate, TotalAmt

	frm1.vspdData.Col = C_VatRate
	If Trim(frm1.vspdData.Text) = "" Or isnull(frm1.vspdData.Text) Then
		VatRate = 0
	Else
		VatRate = UNIFormatNumber(UNICDbl(frm1.vspdData.text), ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
	End If

	frm1.vspdData.Col = C_TotalAmt
	If Trim(frm1.vspdData.Text) = "" Or isnull(frm1.vspdData.Text) Then
		CalSOAmt = 0
	Else
		CalSOAmt = UNIFormatNumber(UNICDbl(frm1.vspdData.text), ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
	End If

	frm1.vspdData.Col = C_VatIncFlag
	If Trim(frm1.vspdData.Text) = "2" Then
		VatAmt = uniFormatNumberByTax(UNICDbl(CalSOAmt) * (UNICDbl(VatRate) / (100 + UNICDbl(VatRate))), frm1.txtCurrency.value, parent.ggAmtOfMoneyNo)
		NetAmt = UNIFormatNumberByCurrecny(UNICDbl(CalSOAmt) - UNICDbl(VatAmt), frm1.txtCurrency.value,parent.ggAmtOfMoneyNo)
		TotalAmt = UNIFormatNumberByCurrecny(UNICDbl(NetAmt) + UNICDbl(VatAmt), frm1.txtCurrency.value,parent.ggAmtOfMoneyNo)
	Else
		VatAmt = uniFormatNumberByTax(UNICDbl(CalSOAmt) * UNICDbl(VatRate)/100, frm1.txtCurrency.value, parent.ggAmtOfMoneyNo)
		NetAmt = CalSOAmt
		TotalAmt = NetAmt
	End If 
 
	frm1.vspdData.Col = C_NetAmt
	frm1.vspdData.Text = NetAmt
	frm1.vspdData.Col = C_VatAmt
	frm1.vspdData.Text = VatAmt
	frm1.vspdData.Col = C_TotalAmt
	frm1.vspdData.Text = TotalAmt

	Call TotalSum

End Function

'========================================================================================================
Sub TotalSum()

	Dim SumTotal, iMonth, lRow
 
	SumTotal = 0
	ggoSpread.source = frm1.vspdData
 
	For lRow = 1 To frm1.vspdData.MaxRows   
		frm1.vspdData.Row = lRow
		frm1.vspdData.Col = 0
		If frm1.vspdData.Text <> ggoSpread.DeleteFlag then
		 frm1.vspdData.Col = C_NetAmt
		 SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
		End If
	Next

	frm1.txtNetAmt.Text = UNIFormatNumberByCurrecny(SumTotal,frm1.txtCurrency.value,parent.ggAmtOfMoneyNo)

End Sub

'========================================================================================================
Function PricePadChange(PRow)

	If PricePadCheckMsg(PRow) = False Then Exit Function       

	Dim strval

	If LayerShowHide(1) = False Then
		Exit Function
	End If

	strVal = ""    
	strVal = BIZ_PGM_ID & "?txtMode=" & lsPricePad			 
	strVal = strVal & "&lsItemCode=" & lsItemCode			  
	strVal = strVal & "&lsSoUnit=" & lsSoUnit
	strVal = strVal & "&lsSoQty=" & lsSoQty
	strVal = strVal & "&lsPriceQty=" & lsPriceQty
	strVal = strVal & "&PRow=" & PRow
	strVal = strVal & "&txtHSoNo=" & Trim(frm1.txtHSoNo.value)
	strVal = strVal & "&txtCurrency=" & Trim(frm1.txtCurrency.value)

	Call RunMyBizASP(MyBizASP, strVal)							

End Function

'========================================================================================================
Function PricePadCheckMsg(CRow)

	PricePadCheckMsg = False

	frm1.vspdData.Row = CRow
	frm1.vspdData.Col = C_ItemCd
	If Len(Trim(frm1.vspdData.Text)) = 0 Then	 
		Exit Function
	End If

	frm1.vspdData.Row = CRow
	frm1.vspdData.Col = C_SoQty
	If Len(Trim(frm1.vspdData.Text)) = 0 Then  
		Exit Function
	End If

	frm1.vspdData.Row = CRow
	frm1.vspdData.Col = C_SoUnit
	If Len(Trim(frm1.vspdData.Text)) = 0 Then
		Exit Function
	End If

	frm1.vspdData.Col = C_ItemCd
	frm1.vspdData.Row = CRow
	lsItemCode = Trim(frm1.vspdData.Text)

	frm1.vspdData.Col = C_SoUnit
	frm1.vspdData.Row = CRow
	lsSoUnit = Trim(frm1.vspdData.Text)

	frm1.vspdData.Col = C_SoQty
	frm1.vspdData.Row = CRow
	lsSoQty = frm1.vspdData.Text

	PricePadCheckMsg = True

End Function


'========================================================================================================
Function ButtonVisible(ByVal BRow)

	ButtonVisible = False

	If frm1.txtHConfirmFlg.value = "N" AND frm1.vspdData.Maxrows > 1 Then
		frm1.vspdData.Row = BRow
		frm1.vspdData.Col = C_APSHost  : lsAPSHost = frm1.vspdData.Text
		frm1.vspdData.Col = C_APSPort  : lsAPSPort = frm1.vspdData.Text
		frm1.vspdData.Col = C_CTPTimes  : lsCTPTimes = UNICDbl(frm1.vspdData.Text)
		frm1.vspdData.Col = C_CTPCheckFlag : lsCTPCheckFlag = frm1.vspdData.Text

		If lsCTPCheckFlag = "Y" Then
			frm1.btnCTPCheck.disabled = False
		Else
			frm1.btnCTPCheck.disabled = True
		End If

	Else
		frm1.btnCTPCheck.disabled = True
	End If

	If frm1.txtHConfirmFlg.value = "N" And lgIntFlgMode = parent.OPMD_UMODE Then
		frm1.btnATPCheck.disabled = False
		If UCase(frm1.HRetItemFlag.value) = "Y" Then frm1.btnATPCheck.disabled = True
	End If     

	ButtonVisible = True

End Function


'========================================================================================================
Function BtnSpreadCheck()

	BtnSpreadCheck = False

	Dim IntRetCD

	ggoSpread.Source = frm1.vspdData 

	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")                
		If IntRetCD = vbNo Then Exit Function
	End If

	If ggoSpread.SSCheckChange = False Then
		IntRetCD = DisplayMsgBox("900018", parent.VB_YES_NO, "X", "X")                
		If IntRetCD = vbNo Then Exit Function
	End If

	BtnSpreadCheck = True

End Function


'========================================================================================================
Function CookiePage(Byval Kubun)
	Const CookieSplit = 4877						
	Dim strTemp, arrVal

	If Kubun = 1 Then
		WriteCookie CookieSplit , frm1.txtConSoNo.value 
	ElseIf Kubun = 0 Then

		strTemp = ReadCookie(CookieSplit)

		If strTemp = "" then Exit Function

		arrVal = Split(strTemp, parent.gRowSep)

		If arrVal(0) = "" Then Exit Function
  
		 frm1.txtConSoNo.value =  arrVal(0) 

		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If

		Call MainQuery()
     
		WriteCookie CookieSplit , ""

	End IF

End Function


'========================================================================================================
Function JumpChgCheck(strJump)

	Dim IntRetCD
	ggoSpread.Source = frm1.vspdData 
	
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")              
		If IntRetCD = vbNo Then	
			Exit Function
		End If
	End If

	Call CookiePage(1)
 
	Select Case Trim(strJump)
	Case BIZ_PGM_JUMP_SOHDR_ID
		Call PgmJump(BIZ_PGM_JUMP_SOHDR_ID)
	Case BIZ_PGM_JUMP_SOSCHE_ID
		Call PgmJump(BIZ_PGM_JUMP_SOSCHE_ID)
	End Select

End Function


'========================================================================================================
Function ItemByHScodeChange(CRow)

	Dim strVal

	frm1.vspdData.Row = CRow

	<% 'ǰ�� %>
	frm1.vspdData.Col = C_ItemCd
	If Trim(frm1.vspdData.Text) = "" Then Exit Function

	If LayerShowHide(1) = False Then
		Exit Function
	End If

	strVal = ""    
	strVal = BIZ_PGM_ID & "?txtMode=" & "ItemByHsCode"        
	<% 'ǰ�� %>
	frm1.vspdData.Col = C_ItemCd
	strVal = strVal & "&ItemCd=" & Trim(frm1.vspdData.Text)
	<% '���� ROW ��ġ %>
	strVal = strVal & "&CRow=" & CRow

	Call RunMyBizASP(MyBizASP, strVal)           

End Function

'========================================================================================================
Function CheckCreditlimitSvr()

	Dim strVal
	Dim TotalVatAmt, AmtDifference, i

	TotalVatAmt = 0

	ggoSpread.Source = frm1.vspdData 
 
	For i = 1 to frm1.vspdData.Maxrows
		frm1.vspdData.Col = C_VatAmt
		frm1.vspdData.Row = i
	 
		TotalVatAmt = TotalVatAmt + UNICDbl(frm1.vspdData.text)
	Next 
 
	AmtDifference = UNICDbl(frm1.txtNetAmt.Text) - UNICDbl(frm1.txtHNetAmt.value) - UNICDbl(frm1.txtHVATAmt.value) + TotalVatAmt 
 
	If LayerShowHide(1) = False Then
		Exit Function
	End If

	strVal = ""    
	strVal = BIZ_PGM_ID & "?txtMode=" & "CheckCreditlimit"       
	strVal = strVal & "&txtHSoNo=" & Trim(frm1.txtHSoNo.value)
	strVal = strVal & "&txtNetAmt=" & AmtDifference
	strVal = strVal & "&txtCurrency=" & Trim(frm1.txtCurrency.value)

	Call RunMyBizASP(MyBizASP, strVal)							

End Function


'========================================================================================================
Function RunAutoDN()

	If LayerShowHide(1) = False Then
		Exit Function
	End If	
	    
	Dim strVal

	If lgIntFlgMode = parent.OPMD_UMODE Then    
		strVal = BIZ_PGM_ID & "?txtMode=" & "DNCheck"        
		strVal = strVal & "&txtSoNo=" & Trim(frm1.txtHSoNo.value)   
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtInsrtUserId=" & Trim(parent.gUsrID)
	Else
		strVal = BIZ_PGM_ID & "?txtMode=" & "DNCheck"       
		strVal = strVal & "&txtSoNo=" & Trim(frm1.txtConSoNo.value)
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtInsrtUserId=" & Trim(parent.gUsrID)
	End If 

	Call RunMyBizASP(MyBizASP, strVal)           

End Function

'========================================================================================================
Sub sprRedComColor(ByVal Col, ByVal Row, ByVal Row2)

    With frm1

		.vspdData.Col = Col
		.vspdData.Col2 = Col
		.vspdData.Row = Row
		.vspdData.Row2 = Row2
		.vspdData.ForeColor = vbRed

    End With
    
End Sub

'========================================================================================================
Function BizProcessCheck()

	BizProcessCheck = False

	If window.document.all("MousePT").style.visibility = "visible" Then Exit Function

	BizProcessCheck = True

End Function

'========================================================================================================
Sub CurFormatNumericOCX()
	With frm1
	 '���ּ��ݾ� 
		ggoOper.FormatFieldByObjectOfCur .txtNetAmt, .txtCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
	 
	End With
End Sub

'========================================================================================================
Sub CurFormatNumSprSheet()
	With frm1

		ggoSpread.Source = frm1.vspdData
		'�ܰ� 
		ggoSpread.SSSetFloatByCellOfCur C_SoPrice,-1, .txtCurrency.value, parent.ggUnitCostNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec,,,"Z"
		'�ݾ� 
		ggoSpread.SSSetFloatByCellOfCur C_NetAmt,-1, .txtCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloatByCellOfCur C_TotalAmt,-1, .txtCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec,,,"Z"
		
	End With
End Sub


'========================================================================================================
Function GetItemPrice(IRow)
	Dim strSoldToParty, strItemCd, strSOUnit, strPayMeth, strDealType, strCurrency, strValidDt
	Dim strSelectList, strFromList, strWhereList
	Dim strRs, strItemInfo

	With frm1
		.vspdData.Row = IRow
		.vspdData.col = C_ItemCd      'ǰ���ڵ� 
		strItemCd = .vspdData.text
		.vspdData.Col = C_SoUnit      '���� 
		strSOUnit = .vspdData.text

		strSoldToParty = .txtSoldToParty.value      '�ֹ�ó 
		strPayMeth = .txtHPayTermsCd.value    '������� 
		strCurrency = .txtCurrency.value    'ȭ����� 
		strValidDt = UniConvDateToYYYYMMDD(.txtSoDt.value, parent.gDateFormat,"")
		strDealType = .txtHDealType.value
 	End With

	If Len(Trim(strItemCd)) = 0 Or Len(Trim(strSOUnit)) = 0 Then Exit Function
	
	strSelectList = " dbo.ufn_s_GetItemSalesPrice( " & FilterVar(strSoldToParty, "''", "S") & ",  " & FilterVar(strItemCd, "''", "S") & ",  " & FilterVar(strDealType, "''", "S") & ",  " & FilterVar(strPayMeth, "''", "S") & "," & _
	    " " & FilterVar(strSOUnit, "''", "S") & ",  " & FilterVar(strCurrency, "''", "S") & ",  " & FilterVar(strValidDt, "''", "S") & ")"
	strFromList  = ""
	strWhereList = ""

    Err.Clear
    
	'ǰ������ �ܰ� Fetch
	If CommonQueryRs2by2(strSelectList,strFromList,strWhereList,strRs) Then
		strItemInfo = Split(strRs, Chr(11))

		frm1.vspdData.Col = C_SoPrice
		frm1.vspdData.text = UNIFormatNumber(strItemInfo(1), ggUnitCost.DecPoint, -2, 0, ggUnitCost.RndPolicy, ggUnitCost.RndUnit)

		Call QtyPriceChange
		Exit Function
	Else
		If Err.number <> 0 Then
			MsgBox Err.description 
			Err.Clear 
			Exit Function
		End If
	End if 

End Function

'========================================================================================================
Sub btnConfirm_OnClick()

	If BtnSpreadCheck = False Then Exit Sub

	Err.Clear                                                               
	   
	If LayerShowHide(1) = False Then
		Exit Sub
	End If
     
    Dim strConfirmVal

	strConfirmVal = BIZ_PGM_ID & "?txtMode=" & "btnCONFIRM"       
	strConfirmVal = strConfirmVal & "&txtHSoNo=" & Trim(frm1.txtHSoNo.value)  
	strConfirmVal = strConfirmVal & "&RdoConfirm=" & Trim(frm1.RdoConfirm.value)
	strConfirmVal = strConfirmVal & "&txtInsrtUserId=" & Trim(parent.gUsrID)
	strConfirmVal = strConfirmVal & "&lgStrPrevKey=" & lgStrPrevKey

	Call RunMyBizASP(MyBizASP, strConfirmVal)          
	Exit Sub


	frm1.txtConfirm.value = lsConfirm

    Err.Clear											 
 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal
 
	If LayerShowHide(1) = False Then
		Exit Sub
	End If

	With frm1
		.txtMode.value = "btnCONFIRM"
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID
		  
		lGrpCnt = 0    
		strVal = ""
		  
		ggoSpread.Source = .vspdData

		For lRow = 1 To .vspdData.MaxRows
    
			.vspdData.Row = lRow
			.vspdData.Col = 0

			Select Case .txtConfirm.value
			Case lsConfirm
				strVal = strVal & .txtConfirm.value & parent.gColSep & lRow & parent.gColSep
				'--- ���ּ��� 
		        .vspdData.Col = C_SoSeq               
		        strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		         
		        '--- ǰ�� 
		        .vspdData.Col = C_ItemCd 
		        strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		         
				'--- ���ּ��� 
		        .vspdData.Col = C_SoQty   
		        strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		        
				'--- ���ִ��� 
		        .vspdData.Col = C_SoUnit   
		        strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		          
				'--- ������ 
		        .vspdData.Col = C_DlvyDt   
		        strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		          
				'--- ��ǰó 
		        .vspdData.Col = C_ShipToParty   
		        strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
              
				'--- ���ִܰ� 
				.vspdData.Col = C_SoPrice   
				strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

				'--- ���� ���ܰ�/���ܰ� 
                .vspdData.Col = C_SoPriceFlag
              
				Select Case Trim(.vspdData.TypeComboBoxCurSel)
				Case 1  '���ܰ� 
					strVal = strVal & "Y" & parent.gColSep    
				Case 0  '���ܰ� 
					strVal = strVal & "N" & parent.gColSep    
				Case Else
					MsgBox "���ܰ�/���ܰ� ���� �����ϴ�.", vbExclamation, parent.gLogoName
					frm1.vspdData.Row = lRow
					frm1.vspdData.Action = 0
					'--�۾��� ǥ��ȭ�� �� ���콺 ����Ʈ ����      
					Call BtnDisabled(False)
					Call LayerShowHide(0)
					Exit Sub
				End Select

				'--- ������ 
		         .vspdData.Col = C_BonusQty   
		         strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		         
				'--- HS��ȣ 
		         .vspdData.Col = C_HsNo   
		         strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		         
				'--- �����������(+)
		         .vspdData.Col = C_TolMoreRate   
		         strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		         
				'--- �����������(-)
		         .vspdData.Col = C_TolLessRate   
		         strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		         
				'--- �����ڵ� 
		         .vspdData.Col = C_PlantCd   
		         strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		         
		         '--- â���ڵ� 
		         .vspdData.Col = C_SlCd   
		         strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		         
				'--- ��� 
		         .vspdData.Col = C_Remark   
		         strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		         
				'--- �������� 
		         .vspdData.Col = C_MaintSeq   
		         strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		         
				'--- �ֹ�����              
		         .vspdData.Col = C_OrderSeq   
		         strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep         

				 lGrpCnt = lGrpCnt + 1 
			End Select       

		Next
 
		.txtMaxRows.value = lGrpCnt
		.txtSpread.value = strVal

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)          '��: �����Ͻ� ASP �� ���� 
 
	End With

End Sub

'========================================================================================================
Sub btnATPCheck_OnClick()

	If BtnSpreadCheck = False Then Exit Sub
	Dim iCalledAspName
	Dim arrParam(1)
	Dim strRet
	 
	ggoSpread.Source = frm1.vspdData
	frm1.vspdData.Col = C_SoSeq
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
 
	arrParam(0) = Trim(frm1.txtHSoNo.value)     
	arrParam(1) = frm1.vspdData.Text
 
	'20021227 kangjungu dynamic popup
	iCalledAspName = AskPRAspName("s3112ra3")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3112ra3", "x")
		exit Sub
	end if

	strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
	  "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

End Sub


'========================================================================================================
Sub btnDNCheck_OnClick()

	If BtnSpreadCheck = False Then Exit Sub

    Call RunAutoDN

End Sub

'========================================================================================================
Sub btnCTPCheck_OnClick()

	Dim Answer
	Answer = DisplayMsgBox("900018", parent.VB_YES_NO, "X", "X") 
	If Answer = VBNO Then Exit Sub

	If frm1.vspdData.ActiveRow  < 1 Or frm1.vspdData.ActiveCol < C_ItemCd Then      
		MsgBox "CTP Check �� ǰ���� �����ϼ���", vbExclamation, parent.gLogoName
	    Exit Sub
	End If

	ggoSpread.Source = frm1.vspdData
	frm1.vspdData.Row = frm1.vspdData.ActiveRow 
	frm1.vspdData.Col = 0
	
	Select Case frm1.vspdData.Text
	Case ggoSpread.InsertFlag
		MsgBox "�ű��Է»����� �����Ŀ� CTP CHECK�� �Ҽ� �ֽ��ϴ�.", vbExclamation, parent.gLogoName
		Exit Sub
	Case ggoSpread.UpdateFlag
		MsgBox "���������� �����Ŀ� CTP CHECK�� �Ҽ� �ֽ��ϴ�.", vbExclamation, parent.gLogoName
		Exit Sub
	End Select
  
	'--- ����û�� ����Ǿ������� Ȯ���Ѵ�.
	Dim SoSts
	frm1.vspdData.Row = frm1.vspdData.ActiveRow 
	frm1.vspdData.Col = C_SoSts
	SoSts = CInt(Trim(frm1.vspdData.Text))

	'---SoStatus ���ֵ��:0,����û:1,���Ϸ�:2,������:3

	If SoSts > 0 Then
		MsgBox "�̹� CTP CHECK�� �Ϸ�� ��ǰ�Դϴ�.", vbExclamation, parent.gLogoName
		Exit Sub
	End If

	Dim iCalledAspName
	Dim arrRet
	Dim arrParam, arrSoNo(0), arrGridCount(0)
	Dim iRow
	Dim strBaseQty, strBonusBaseQty

	If IsOpenPop = True Then Exit Sub

	IsOpenPop = True

	arrSoNo(0) = frm1.txtHSoNo.value
	arrGridCount(0) = frm1.vspdData.MaxRows - 1

	ReDim arrParam(0,9)

	iRow = frm1.vspdData.ActiveRow 

	frm1.vspdData.Col = C_ItemCd
	frm1.vspdData.Row = iRow
	arrParam(0,0) = frm1.vspdData.Text
 
	frm1.vspdData.Col = C_ItemName
	frm1.vspdData.Row = iRow
	arrParam(0,1) = frm1.vspdData.Text
	 
	frm1.vspdData.Col = C_SoSeq
	frm1.vspdData.Row = iRow
	arrParam(0,2) = frm1.vspdData.Text

	frm1.vspdData.Col = C_DnReqDt '���Ͽ�û���� 
	frm1.vspdData.Row = iRow
	arrParam(0,3) = frm1.vspdData.Text

	frm1.vspdData.Col = C_PlantCd
	frm1.vspdData.Row = iRow
	arrParam(0,4) = frm1.vspdData.Text

	frm1.vspdData.Col = C_BaseQty
	frm1.vspdData.Row = iRow
	strBaseQty = frm1.vspdData.Text

	frm1.vspdData.Col = C_BonusBaseQty
	frm1.vspdData.Row = iRow
	strBonusBaseQty = frm1.vspdData.Text

	If Trim(strBaseQty) <> "" Then
	 strBaseQty = UNICDbl(strBaseQty)
	Else
	 strBaseQty = 0
	End If

	If Trim(strBonusBaseQty) <> "" Then
	 strBonusBaseQty = UNICDbl(strBonusBaseQty)
	Else
	 strBonusBaseQty = 0
	End If

	arrParam(0,5) = strBaseQty + strBonusBaseQty

	frm1.vspdData.Col = C_TrackingNo
	frm1.vspdData.Row = iRow
	arrParam(0,6) = frm1.vspdData.Text

	arrParam(0,7) = lsAPSHost
	arrParam(0,8) = lsAPSPort
	arrParam(0,9) = lsCTPTimes
	'20021227 kangjungu dynamic popup
	iCalledAspName = AskPRAspName("s3113pa1")	

	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3113pa1", "x")
		IsOpenPop = False
		exit sub
	end if

	IsOpenPop = True

	arrRet = window.showModalDialog(iCalledAspName,Array(window.parent,arrParam,arrSoNo,arrGridCount),_
	 "dialogWidth=480px; dialogHeight=400px; center: Yes; help: No; resizable: No; status: No;")
  
	IsOpenPop = False

	If arrRet = "" Then
		Exit Sub
	ElseIf arrRet = "CTPAccept" Then
		Exit Sub	 
	ElseIf arrRet = "CTPModify" Then
		Call ggoOper.ClearField(Document, "2")
		Call DbQuery
	ElseIf arrRet = "Cancel" Then
		Exit Sub
	ElseIf arrRet = "Save" Then
		Call ggoOper.ClearField(Document, "2")
		Call DbQuery
	End If

End Sub


'========================================================================================================
Sub btnAvlStkRef_OnClick()
	Call OpenAvalStockRef()
End Sub 

'========================================================================================================
Sub vspdData_ButtonClicked(Col, Row, ButtonDown)

	With frm1.vspdData 

		ggoSpread.Source = frm1.vspdData
	  
		If Row > 0 And Col = C_ItemPopup Then
		    .Col = Col - 1
		    .Row = Row
			Call OpenItem(.Text)
		ElseIf Row > 0 And Col = C_SoUnitPopup Then
		    .Col = Col - 1
		    .Row = Row
		    Call OpenSoDtl(.Text, 1)
		ElseIf Row > 0 And Col = C_ShipToPartyPopup Then
		    .Col = Col - 1
		    .Row = Row
		    Call OpenSoDtl(.Text, 3)
		ElseIf Row > 0 And Col = C_HsNoPopup Then
		    .Col = Col - 1
		    .Row = Row
		    Call OpenSoDtl(.Text, 4)
		ElseIf Row > 0 And Col = C_PlantCdPopup Then
		    .Col = Col - 1
		    .Row = Row
		    Call OpenSoDtl(.Text, 5)
		ElseIf Row > 0 And Col = C_SlCdPopup Then
		    .Col = Col - 1
		    .Row = Row
		    Call OpenSoDtl(.Text, 6)
		ElseIf Row > 0 And Col = C_VatTypePopup Then
		    .Col = Col - 1
		    .Row = Row
		    Call OpenSoDtl(.Text, 7)
		ElseIf Row > 0 And Col = C_RetTypePopup Then
		    .Col = Col - 1
		    .Row = Row
		    Call OpenSoDtl(.Text, 8)
		End If
    
 End With
	Call SetActiveCell(frm1.vspdData,Col - 1,Row,"M","X","X")
End Sub


'==========================================================================================
Sub vspdData_Change(Col , Row)
	
	Dim iDx
       
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	 Select Case Col
	 Case C_SoQty
		lsPriceQty = "Q"
		Call PricePadChange(Row)
		Call QtyPriceChange()

	 Case C_SoPrice
		Call QtyPriceChange() 
	 
	 Case C_ItemCd
		Call ItemByHScodeChange(Row)

	 Case C_SoUnit
		Call GetItemPrice(Row)
		lsPriceQty = "Q"
		Call PricePadChange(Row)
 	 
 	 Case C_VatType
		Call SetVatType(Row)

	 Case C_VatIncFlag
		Call QtyPriceChange()

	 Case C_TotalAmt
		Call TotalAmtChange()
		
	 End Select
	 '------ Developer Coding part (End   ) -------------------------------------------------------------- 
             
	Call CheckMinNumSpread(frm1.vspdData, Col, Row)		
    
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

	lgBlnFlgChgValue = True

End Sub


'========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )

    gMouseClickStatus = "SPC"    
    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
   	If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData
       
       If lgSortKey = 1 Then
		  ggoSpread.SSSort Col				'Sort in Ascending
		  lgSortkey = 2
       Else
		  ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
		  lgSortkey = 1
	   End If
	   Call SetPopupMenuItemInf("0000000000")
       frm1.btnATPCheck.disabled = True
	   frm1.btnCTPCheck.disabled = True
       Exit Sub
    End If
    
    '------ Developer Coding part (Start ) -------------------------------------------------------------- 
	If ButtonVisible(Row) = False Then Exit Sub    <% 'CTP ��󿩺ο� ���� ��ư üũ %>
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
	If frm1.RdoConfirm.Value = "Y" Then   
		Call SetPopupMenuItemInf("0101111111")   
	Else
		Call SetPopupMenuItemInf("0000111111")   
	End IF	
End Sub


'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
    Dim intIndex

	If Col = C_VatIncFlagNm Then
 
		With frm1.vspdData
		  .Row = Row
		  .Col = Col
		  intIndex = .Value
		  
		  .Col = C_VatIncFlag
		  .Value = intIndex+1
		End With
		
		Call vspdData_Change(C_VatIncFlag , Row)
		
	End If
End Sub


'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub


'==========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) --------------------------------------------------------------
End Sub


'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

	If Col < 0 Or Row < 0 Or NewCol < 0 Or NewRow < 0 Then Exit Sub
	 
	If Row < NewRow Then
		Call ButtonVisible(Row+1)
	ElseIf Row > NewRow Then
		Call ButtonVisible(Row-1)
	Else
		Call ButtonVisible(Row)
	End If

End Sub

'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )    
    If OldLeft <> NewLeft Then Exit Sub
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    
    	If lgStrPrevKey <> "" Then								
           Call DisableToolBar(parent.TBC_QUERY)
           Call DBQuery
    	End If
    End If    
End Sub



</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" --> 
</HEAD>

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
     <TD CLASS="CLSMTABP">
      <TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
       <TR>
        <td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>STO��������</font></td>
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
          </TR>
      </TABLE>
     </TD>
     <TD WIDTH=* align=right><A href="vbscript:OpenStockDtlRef">�����Ȳ����</TD>
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
     <TD HEIGHT=20 WIDTH=100%>
      <FIELDSET CLASS="CLSFLD">
       <TABLE <%=LR_SPACE_TYPE_40%>>
        <TR>
         <TD CLASS="TD5" NOWRAP>���ֹ�ȣ</TD>
         <TD CLASS="TD6"><INPUT NAME="txtConSoNo" ALT="���ֹ�ȣ" TYPE="Text" MAXLENGTH=18 SiZE=34 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSSoDtl" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSoDtl()"></TD>
         <TD CLASS="TDT"></TD>
         <TD CLASS="TD6"></TD>
        </TR>
       </TABLE>
      </FIELDSET>
     </TD>
    </TR>
    <TR>
     <TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
    </TR>
    <TR>
     <TD WIDTH=100% HEIGHT=* VALIGN=TOP>
      <TABLE <%=LR_SPACE_TYPE_60%>>
		<TR>
			<TD CLASS="TD5" NOWRAP>���ֹ�ȣ</TD>
			<TD CLASS="TD6"><INPUT NAME="txtSoNo" ALT="���ֹ�ȣ" TYPE="Text" MAXLENGTH=18 SiZE=34 tag="24XXXU"></TD>
			<TD CLASS=TD5 NOWRAP>����Ȯ��</TD>
			<TD CLASS=TD6 NOWRAP>
				<input type=radio CLASS="RADIO" name="rdoCfm_flag" id="rdoCfm_flag1" value="Y" tag = "24">
					<label for="rdoCfm_flag1">Ȯ��</label>&nbsp;&nbsp;&nbsp;&nbsp;
				<input type=radio CLASS = "RADIO" name="rdoCfm_flag" id="rdoCfm_flag2" value="N" tag = "24" checked>
					<label for="rdoCfm_flag2">��Ȯ��</label></TD>
		</TR>
		<TR>
			<TD CLASS=TD5 NOWRAP>��������</TD>
			<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSoType" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="24XXXU" ALT="��������">&nbsp;<INPUT NAME="txtSoTypeNm" TYPE="Text" SIZE=25 tag="24"></TD>
			<TD CLASS=TD5 NOWRAP>���ֹ���ȣ</TD>
			<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCustpono" TYPE="Text" MAXLENGTH="20" SIZE=20 ALT="���ֹ���ȣ" tag="24XXXU"></TD>
		</TR>
		<TR>
			<TD CLASS=TD5 NOWRAP>������</TD>
			<TD CLASS=TD6 NOWRAP>
				<script language =javascript src='./js/s3171ma1_fpDateTime1_txtSoDt.js'></script>
			</TD>
			<TD CLASS=TD5 NOWRAP>�ֹ�ó</TD>
			<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSoldtoparty" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="24XXXU" ALT="�ֹ�ó">&nbsp;<INPUT NAME="txtSoldtopartyNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24"></TD>
		</TR>
		<TR>
			<TD CLASS=TD5 NOWRAP>�Ǹ�����</TD>
			<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDealType" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="21XXXU" ALT="�Ǹ�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenHdr 0">&nbsp;<INPUT NAME="txtDealTypeNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24"></TD>
			<TD CLASS=TD5 NOWRAP>�����׷�</TD>
			<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSalesGrp" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="22XXXU" ALT="�����׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenHdr 1">&nbsp;<INPUT NAME="txtSalesGrpNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24"></TD>
		</TR>
		<TR>
			<TD CLASS=TD5 NOWRAP>�������</TD>
			<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPaymeth" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="21XXXU" ALT="�������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenHdr 2">&nbsp;<INPUT NAME="txtPaymethNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24"></TD>
			<TD CLASS=TD5 NOWRAP>���ֱݾ�</TD>
			<TD CLASS=TD6 NOWRAP>
				<TABLE CELLSPACING=0 CELLPADDING=0>
					<TR>
						<TD>
							<script language =javascript src='./js/s3171ma1_fpDoubleSingle2_txtNetAmt.js'></script>&nbsp;
						</TD>
						<TD>
							<INPUT NAME="txtCurrency" TYPE="Text" MAXLENGTH="3" SIZE=10 tag="24XXXU" ALT="ȭ��">
						</TD>						
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD CLASS=TD5 NOWRAP>���</TD>
			<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRemark" TYPE="Text" MAXLENGTH="70" SIZE=30 ALT="���" tag="21"></TD>			
			<TD CLASS=TD5 NOWRAP>ȯ��</TD>
			<TD CLASS=TD6 NOWRAP>
				<script language =javascript src='./js/s3171ma1_fpDoubleSingle3_txtXchgRate.js'></script>
			</TD>
		</TR>
		<TR>
        <TD HEIGHT="100%" WIDTH="100%" COLSPAN=4>
         <script language =javascript src='./js/s3171ma1_OBJECT1_vspdData.js'></script>
        </TD>
       </TR>
       
      </TABLE>
     </TD>
    </TR>
   </TABLE>
  </TD>
 </TR>
 <TR >
  <TD <%=HEIGHT_TYPE_01%>></TD>
 </TR>
 <TR HEIGHT=20>
  <TD WIDTH=100%>
   <TABLE <%=LR_SPACE_TYPE_30%>>
       <TR>
     <TD WIDTH=10>&nbsp;</TD>
     <TD>
      <BUTTON NAME="btnConfirm" CLASS="CLSSBTN">Ȯ��ó��</BUTTON>&nbsp;
      <BUTTON NAME="btnATPCheck" CLASS="CLSSBTN">ATP Check</BUTTON>&nbsp;
      <BUTTON NAME="btnCTPCheck" CLASS="CLSSBTN">CTP Check</BUTTON>&nbsp;
      <BUTTON NAME="btnDNCheck" CLASS="CLSSBTN">���Ͽ�ûó��</BUTTON>&nbsp;
      <BUTTON NAME="btnAvlStkRef" CLASS="CLSSBTN">���������Ȳ</BUTTON>
     </TD>     
    </TR>
   </TABLE>
  </TD>
 </TR>
 <TR >
  <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
  </TD>
 </TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtConfirm" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="RdoConfirm" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConfirmFlg" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtShipToParty" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHNetAmt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHVATAmt" tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="HReqDlvyDt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="HPriceFlag" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="HExportFlag" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="HRetItemFlag" tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtHPreSONo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHSoNo" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHVatRate" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHVATType" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHVATIncFlag" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHVATIncFlagNm" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHMaintNo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHSoldToParty" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHPayTermsCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHDealType" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtCtpCDFlag" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtRadioFlag" tag="14" TABINDEX="-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
 <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>
