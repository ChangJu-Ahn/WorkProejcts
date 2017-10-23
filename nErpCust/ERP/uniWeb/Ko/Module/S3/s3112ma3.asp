

<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 수주관리 
'*  3. Program ID           : S3112MA3
'*  4. Program Name         : 수주내역등록 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/12/04
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Son bum Yeol
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/20 : 3rd 화면 layout & ASP Coding
'*                            -2000/08/11 : 4th 화면 layout
'*                            -2001/12/18 : Date 표준적용 
'**********************************************************************************************
-->

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
Option Explicit															'☜: Turn on the Option Explicit option.


'========================================================================================================
'=								2.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID				= "s3112mb3.asp"						'Biz Logic ASP								
Const BIZ_PGM_JUMP_SOHDR_ID		= "s3111ma1"
Const BIZ_PGM_JUMP_SOSCHE_ID	= "s3114ma8"

'========================================================================================================
'=								2.2 Constant variables For spreadsheet
'========================================================================================================
Dim C_ItemCd			 '품목  
Dim C_ItemPopup			 '품목팝업 
Dim C_ItemName			 '품목명 
Dim C_ItemSpec			 
Dim C_SoUnit			 '단위 
Dim C_SoUnitPopup		 '단위팝업 
Dim C_TrackingNo		 'Tracking No
Dim C_SoSupplyQty		 '수주잔량(ReqQty - SoQty)
Dim C_SoQty				 '수량 
Dim C_SoPrice			 '단가 
Dim C_SoPriceAutoChk	 '단가자동계산체크 
Dim C_SoPriceFlag		 '가단가/진단가 
Dim C_TotalAmt			 '화면 수주금액(거래화폐)
Dim C_NetAmt			 'Hidden 수주금액 
Dim C_VATAmt				
Dim C_LocAmt			 '수주자국금액 
Dim C_VatLocAmt			 'VAT자국금액 
Dim C_PlantCd			 '공장코드 
Dim C_PlantCdPopup		 '공장팝업 
Dim C_PlantNm			 '공장명 
Dim C_DlvyDt			 '납기일       
Dim C_ShipToParty		 '납품처 
Dim C_ShipToPartyPopup	 '납품처팝업 
Dim C_ShipToPartyNm		 '납품처명 
Dim C_HsNo				 'HS번호 
Dim C_HsNoPopup			 'HS번호 Popup
Dim C_TolMoreRate		 '과부족허용율(+)
Dim C_TolLessRate		 '과부족허용율(-)
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
Dim C_PreDnNo			 '출하번호 for 수주내역참조 
Dim C_PreDnSeq			 '출하순번 for 수주내역참조 
Dim C_DnReqDt			 '출하요청일자 
Dim C_BonusQty			 '할증수량(덤)        
Dim C_SlCd				 '창고코드 
Dim C_SlCdPopup			 '창고팝업 
Dim C_SlNm				 '창고명 
Dim C_Remark			 '비고 
Dim C_SoSts				 '수주진행상태 
Dim C_BillQty			 '매출수량 
Dim C_BaseQty			 '재고수량 
Dim C_BonusBaseQty		 '덤재고수량 
Dim C_MaintSeq			 '관리순번 
Dim C_OrderSeq			 '주문서순번 
Dim C_APSHost			 'APS Host
Dim C_APSPort			 'APS Port
Dim C_CTPTimes			 'CTP Check 횟수 
Dim C_CTPCheckFlag		 'CTP Check Flag
Dim C_SoSeq				 '수주순번 
Dim C_PreSoNo			 '수주번호 for 수주내역참조 
Dim C_PreSoSeq			 '수주순번 for 수주내역참조 

Dim ext1_qty 
Dim ext2_qty 
Dim ext3_qty 
Dim ext1_amt 
Dim ext2_amt 
Dim ext3_amt 
Dim ext1_cd  
Dim ext2_cd 
Dim ext3_cd  



'========================================================================================================
'=							2.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->

'========================================================================================================
'=							2.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop
<% 
   BaseDate = GetSvrDate                                                          'Get DB Server Date
%> 
Dim EndDate
Dim StartDate
'초기화면에 뿌려지는 마지막 날짜 (Convert DB date type to Company)
EndDate		= UniConvDateAToB("<%=BaseDate%>", parent.gServerDateFormat, parent.gDateFormat)
'초기화면에 뿌려지는 시작 날짜 
StartDate	= UNIDateAdd("m", -1, EndDate, parent.gDateFormat)

Const C_SHEETMAXROWS = 30		'Sheet Max Rows
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

'########################################################################################################
'#								3.	Method Declaration Part
'########################################################################################################
'========================================================================================================
'								3.1 Common Group-1
'========================================================================================================
'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
	C_ItemCd			 = 1   '품목  
	C_ItemPopup			 = 2   '품목팝업 
	C_ItemName			 = 3   '품목명 
	C_ItemSpec			 = 4
	C_SoUnit			 = 5   '단위 
	C_SoUnitPopup		 = 6   '단위팝업 
	C_TrackingNo		 = 7   'Tracking No
	C_SoSupplyQty		 = 8   '수주잔량(ReqQty - SoQty)
	C_SoQty				 = 9   '수량 
	C_SoPrice			 = 10  '단가 
	C_SoPriceAutoChk	 = 11  '단가자동계산체크 
	C_SoPriceFlag		 = 12  '가단가/진단가 
	C_TotalAmt			 = 13  '화면 수주금액(거래화폐)
	C_NetAmt			 = 14  'Hidden 수주금액 
	C_VATAmt			 = 15
	C_LocAmt			 = 16  '수주자국금액 
	C_VatLocAmt			 = 17  'VAT자국금액 
	C_PlantCd			 = 18  '공장코드 
	C_PlantCdPopup		 = 19  '공장팝업 
	C_PlantNm			 = 20  '공장명 
	C_DlvyDt			 = 21  '납기일       
	C_ShipToParty		 = 22  '납품처 
	C_ShipToPartyPopup	 = 23  '납품처팝업 
	C_ShipToPartyNm		 = 24  '납품처명 
	C_HsNo				 = 25  'HS번호 
	C_HsNoPopup			 = 26  'HS번호 Popup
	C_TolMoreRate		 = 27  '과부족허용율(+)
	C_TolLessRate		 = 28  '과부족허용율(-)
	C_VatType			 = 29
	C_VatTypePopup		 = 30
	C_VatTypeNm			 = 31
	C_VatRate			 = 32
	C_VatIncFlag		 = 33
	C_VatIncFlagNm		 = 34
	C_RetType			 = 35
	C_RetTypePopup		 = 36
	C_RetTypeNm			 = 37
	C_LotNo				 = 38
	C_LotSeq			 = 39
	C_PreDnNo			 = 40  '출하번호 for 수주내역참조 
	C_PreDnSeq			 = 41  '출하순번 for 수주내역참조 
	C_DnReqDt			 = 42  '출하요청일자 
	C_BonusQty			 = 43  '할증수량(덤)        
	C_SlCd				 = 44  '창고코드 
	C_SlCdPopup			 = 45  '창고팝업 
	C_SlNm				 = 46  '창고명 
	C_Remark			 = 47  '비고 
	C_SoSts				 = 48  '수주진행상태 
	C_BillQty			 = 49  '매출수량 
	C_BaseQty			 = 50  '재고수량 
	C_BonusBaseQty		 = 51  '덤재고수량 
	C_MaintSeq			 = 52  '관리순번 
	C_OrderSeq			 = 53  '주문서순번 
	C_APSHost			 = 54  'APS Host
	C_APSPort			 = 55  'APS Port
	C_CTPTimes			 = 56  'CTP Check 횟수 
	C_CTPCheckFlag		 = 57  'CTP Check Flag
	C_SoSeq				 = 58  '수주순번 
	C_PreSoNo			 = 59  '수주번호 for 수주내역참조 
	C_PreSoSeq			 = 60  '수주순번 for 수주내역참조 
	
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
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()

	lgIntFlgMode      = Parent.OPMD_CMODE						'⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed    
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	frm1.txtConSoNo.focus
'	frm1.btnConfirm.disabled	= True
'	frm1.btnConfirm.value		= "확정처리"
'	frm1.btnDNCheck.disabled	= True
'	frm1.btnATPCheck.disabled	= True
'	frm1.btnCTPCheck.disabled	= True
	frm1.btnAvlStkRef.disabled	= True
	frm1.txtPlant.value			= parent.gPlant
	frm1.txtPlantNm.value		= parent.gPlantNm  
	lgBlnFlgChgValue			= False
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
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()
	
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		'patch version
		ggoSpread.Spreadinit "V20021214",,parent.gAllowDragDropSpread
		.ReDraw = false

		
	  .MaxCols	=	C_PreSoSeq	+  1							' ☜: Add 1 to Maxcols  	  
	  .MaxRows = 0												' ☜: Clear spreadsheet data 
	
	Call GetSpreadColumnPos("A")	 
								   'ColumnPosition		Header              Width	Align(0:L,1:R,2:C)  Row   Length  CharCase(0:L,1:N,2:U)
	  ggoSpread.SSSetEdit	C_SoSeq,			"수주순번",		3,		1
	  ggoSpread.SSSetEdit	C_ItemCd,			"품목",			18,		,					,	  18,	  2
	  ggoSpread.SSSetEdit	C_ItemSpec,			"규격",			20
								   'ColumnPosition		Row
	  ggoSpread.SSSetButton	C_ItemPopup			
	  ggoSpread.SSSetEdit	C_ItemName,			"품목명",		25,		,					,	  40
	  ggoSpread.SSSetEdit	C_TrackingNo,		"Tracking No",	15,		,					,	  25,	  2
								   'ColumnPosition      Header              Width	Grp					  IntegeralPart					DeciPointpart               Align				 Sep				PZ  Min Max 
	  ggoSpread.SSSetFloat	C_SoSupplyQty,		"수주잔량",		15,		parent.ggQtyNo,	      ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
	  ggoSpread.SSSetFloat	C_SoQty,			"수량",			15,		parent.ggQtyNo,		  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
	  ggoSpread.SSSetEdit	C_SoUnit,			"단위",			8,		,					,	  3,	  2
	  ggoSpread.SSSetButton	C_SoUnitPopup		
								   'ColumnPosition      Header				Width	Align(0:L,1:R,2:C)  Format         Row
	  ggoSpread.SSSetDate	C_DlvyDt,			"납기일",		10,		2,					parent.gDateFormat
	  ggoSpread.SSSetDate	C_DnReqDt,			"출하예정일자",	15,		2,					parent.gDateFormat
	  ggoSpread.SSSetEdit	C_ShipToParty,		"납품처",		10,		,					,	  10,	  2
	  ggoSpread.SSSetButton	C_ShipToPartyPopup	
	  ggoSpread.SSSetEdit	C_ShipToPartyNm,	"납품처명",		10
	  ggoSpread.SSSetFloat	C_SoPrice,			"단가",			15,		parent.ggUnitCostNo,  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
	  ggoSpread.SSSetCheck	C_SoPriceAutoChk,	"",		2,	,	,	True
								   'ColumnPosition      Header				Width	Align(0:L,1:R,2:C)  ComboEditable  Row
	  ggoSpread.SSSetCombo	C_SoPriceFlag,		"단가구분",		10,		2
	  ggoSpread.SetCombo		"가단가" & vbTab & "진단가",C_SoPriceFlag
								   'ColumnPosition      Header              Width	Grp						IntegeralPart				DeciPointpart				Align				Sep					PZ  Min Max 	  
	  ggoSpread.SSSetFloat	C_TotalAmt,			"금액",			15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
	  ggoSpread.SSSetFloat	C_NetAmt,			"순금액",		15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
	  ggoSpread.SSSetFloat	C_VATAmt,			"VAT금액",		15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
	  ggoSpread.SSSetFloat	C_LocAmt,			"수주자국금액",		15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
	  ggoSpread.SSSetFloat	C_VatLocAmt,		"VAT자국금액",		15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
	  ggoSpread.SSSetFloat	C_BonusQty,			"덤수량" ,		15,		parent.ggQtyNo,			ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
								   'ColumnPosition		Header              Width	Align(0:L,1:R,2:C)  Row   Length  CharCase(0:L,1:N,2:U)
	  ggoSpread.SSSetEdit	C_HsNo,				"HS부호",		15,		,					,	  20,	  2
	  ggoSpread.SSSetButton	C_HsNoPopup			
	  ggoSpread.SSSetEdit	C_VatType,			"VAT유형",		10,		,					,	  5,	  2
	  ggoSpread.SSSetButton	C_VatTypePopup		
	  ggoSpread.SSSetEdit	C_VatTypeNm,		"VAT유형명",	20
	  ggoSpread.SSSetFloat	C_VatRate,			"VAT율",		10,		parent.ggExchRateNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
	' ggoSpread.SSSetEdit	C_VatIncFlag,		"VAT포함구분",	12,		,					,	  5,	  2
	' ggoSpread.SSSetButton	C_VatIncPopup
	' ggoSpread.SSSetEdit	C_VatIncFlagNm,		"VAT포함구분명",20
	  ggoSpread.SSSetCombo	C_VatIncFlagNm,		"VAT포함구분명",15,		2
	  ggoSpread.SetCombo		"별도" & vbTab & "포함",C_VatIncFlagNm
	  ggoSpread.SSSetEdit	C_VatIncFlag,		"VAT포함구분",	5,		2
	  ggoSpread.SetCombo		"1"		   & vbTab & "2",		C_VatIncFlag
	  ggoSpread.SSSetEdit	C_RetType,			"반품유형",		10,		,					,	  5,	  2
	  ggoSpread.SSSetButton	C_RetTypePopup
	  ggoSpread.SSSetEdit	C_RetTypeNm,		"반품유형명",	20
	  ggoSpread.SSSetEdit	C_LotNo,			"LOT NO",		12,		,					,	  18,	  2
	  Call AppendNumberPlace("7","3","0")
								   'ColumnPosition      Header					Width	Grp		IntegeralPart				DeciPointpart				Align				Sep					PZ  Min Max 	  
	  ggoSpread.SSSetFloat	C_LotSeq,			"LOT NO 순번" ,		15,		"7",	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"  
	  Call AppendNumberPlace("6","9","6")
	  ggoSpread.SSSetFloat	C_TolMoreRate,		"과부족허용율(+)" ,	15,		"6",	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"  
	  ggoSpread.SSSetFloat	C_TolLessRate,		"과부족허용율(-)" ,	15,		"6",	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
								   'ColumnPosition		Header              Width	Align(0:L,1:R,2:C)  Row   Length  CharCase(0:L,1:N,2:U)
	  ggoSpread.SSSetEdit	C_PlantCd,			"공장",			8,		,					,	  4,	  2
	  ggoSpread.SSSetButton	C_PlantCdPopup		
	  ggoSpread.SSSetEdit	C_PlantNm,			"공장명",		8
	  ggoSpread.SSSetEdit	C_SlCd,				"창고",			8,		,					,	  7,	  2
	  ggoSpread.SSSetButton	C_SlCdPopup
	  ggoSpread.SSSetEdit	C_SlNm,				"창고명",		8
	  ggoSpread.SSSetEdit	C_Remark,			"비고",			60,		,					,	  60
								   'ColumnPosition      Header				Width	Grp				IntegeralPart				DeciPointpart				Align				Sep					PZ  Min Max 	  
	' ggoSpread.SSSetFloat	C_SoSts,			"수주진행상태",	15,		parent.ggQtyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
	  ggoSpread.SSSetFloat	C_BillQty,			"매출수량",		15,		parent.ggQtyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
								   'ColumnPosition		Header              Width	Align(0:L,1:R,2:C)  Row   Length  CharCase(0:L,1:N,2:U)
	  ggoSpread.SSSetEdit	C_MaintSeq,			"관리SEQ",		10
	  ggoSpread.SSSetEdit	C_OrderSeq,			"주문서순번",	10
	  ggoSpread.SSSetEdit	C_APSHost,			"APSHost",			5,		,					,	  20
	  ggoSpread.SSSetEdit	C_APSPort,			"APSPort",			5,		,					,	  20
	  ggoSpread.SSSetEdit	C_CTPTimes,			"CTPTimes",			5,		,					,	  3
	  ggoSpread.SSSetEdit	C_CTPCheckFlag,		"CTPCheckFlag",		5,		,					,	  2
	  ggoSpread.SSSetEdit	C_PreDnNo,			"출하번호",		18,		,					,	  18,	  2
	  ggoSpread.SSSetEdit	C_PreDnSeq,			"출하순번",		10,		,					,	  3,	  1

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
      Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)				'☜: 공통콘트롤 사용 Hidden Column
           
	  .ReDraw = true
   
	 ' Call SetSpreadLock
	  
   End With
    
End Sub
'======================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect cell in spread sheet
'======================================================================================================
Sub SetSpreadLock()
    'With frm1
    
    '.vspdData.ReDraw = False 
                                 'Col-1          Row-1       Col-2           Row-2   
    '  ggoSpread.SpreadLock       C_SID        , -1         , C_SID        , -1 
    '  ggoSpread.SpreadLock       C_AddressNm  , -1         , C_AddressNm  , -1 
                                 'Col            Row         Row2
    '  ggoSpread.SSSetRequired    C_SNm        , -1         ,-1
    '.vspdData.ReDraw = True

    'End With
End Sub

'======================================================================================================
' Name : SetSpreadColor
' Desc : This method set color and protect cell by cell in spread sheet
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
		ggoSpread.SSSetProtected C_LocAmt,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_DnReqDt,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_VATAmt,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_VATLocAmt,	pvStartRow, pvEndRow
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
 
		' 수출수주/국내수주 여부에 따라 덤수량,HS부호 관리 
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
			C_ItemCd			 = iCurColumnPos(1)   '품목  
			C_ItemPopup			 = iCurColumnPos(2)   '품목팝업 
			C_ItemName			 = iCurColumnPos(3)   '품목명 
			C_ItemSpec			 = iCurColumnPos(4)
			C_SoUnit			 = iCurColumnPos(5)   '단위 
			C_SoUnitPopup		 = iCurColumnPos(6)   '단위팝업 
			C_TrackingNo		 = iCurColumnPos(7)   'Tracking No
			C_SoSupplyQty		 = iCurColumnPos(8)   '수주잔량(ReqQty - SoQty)
			C_SoQty				 = iCurColumnPos(9)   '수량 
			C_SoPrice			 = iCurColumnPos(10)  '단가 
			C_SoPriceAutoChk	 = iCurColumnPos(11)  '단가자동계산체크 
			C_SoPriceFlag		 = iCurColumnPos(12)  '가단가/진단가 
			C_TotalAmt			 = iCurColumnPos(13)  '화면 수주금액(거래화폐)
			C_NetAmt			 = iCurColumnPos(14)  'Hidden 수주금액 
			C_VATAmt			 = iCurColumnPos(15)
			C_LocAmt			 = iCurColumnPos(16)  '수주자국금액 
			C_VATLocAmt			 = iCurColumnPos(17)  'vat자국금액	
			C_PlantCd			 = iCurColumnPos(18)  '공장코드 
			C_PlantCdPopup		 = iCurColumnPos(19)  '공장팝업 
			C_PlantNm			 = iCurColumnPos(20)  '공장명 
			C_DlvyDt			 = iCurColumnPos(21)  '납기일       
			C_ShipToParty		 = iCurColumnPos(22)  '납품처 
			C_ShipToPartyPopup	 = iCurColumnPos(23)  '납품처팝업 
			C_ShipToPartyNm		 = iCurColumnPos(24)  '납품처명 
			C_HsNo				 = iCurColumnPos(25)  'HS번호 
			C_HsNoPopup			 = iCurColumnPos(26)  'HS번호 Popup
			C_TolMoreRate		 = iCurColumnPos(27)  '과부족허용율(+)
			C_TolLessRate		 = iCurColumnPos(28)  '과부족허용율(-)
			C_VatType			 = iCurColumnPos(29)
			C_VatTypePopup		 = iCurColumnPos(30)
			C_VatTypeNm			 = iCurColumnPos(31)
			C_VatRate			 = iCurColumnPos(32)
			C_VatIncFlag		 = iCurColumnPos(33)
			C_VatIncFlagNm		 = iCurColumnPos(34)
			C_RetType			 = iCurColumnPos(35)
			C_RetTypePopup		 = iCurColumnPos(36)
			C_RetTypeNm			 = iCurColumnPos(37)
			C_LotNo				 = iCurColumnPos(38)
			C_LotSeq			 = iCurColumnPos(39)
			C_PreDnNo			 = iCurColumnPos(40)  '출하번호 for 수주내역참조 
			C_PreDnSeq			 = iCurColumnPos(41)  '출하순번 for 수주내역참조 
			C_DnReqDt			 = iCurColumnPos(42)  '출하요청일자 
			C_BonusQty			 = iCurColumnPos(43)  '할증수량(덤)        
			C_SlCd				 = iCurColumnPos(44)  '창고코드 
			C_SlCdPopup			 = iCurColumnPos(45)  '창고팝업 
			C_SlNm				 = iCurColumnPos(46)  '창고명 
			C_Remark			 = iCurColumnPos(47)  '비고 
			C_SoSts				 = iCurColumnPos(48)  '수주진행상태 
			C_BillQty			 = iCurColumnPos(49)  '매출수량 
			C_BaseQty			 = iCurColumnPos(50)  '재고수량 
			C_BonusBaseQty		 = iCurColumnPos(51)  '덤재고수량 
			C_MaintSeq			 = iCurColumnPos(52)  '관리순번 
			C_OrderSeq			 = iCurColumnPos(53)  '주문서순번 
			C_APSHost			 = iCurColumnPos(54)  'APS Host
			C_APSPort			 = iCurColumnPos(55)  'APS Port
			C_CTPTimes			 = iCurColumnPos(56)  'CTP Check 횟수 
			C_CTPCheckFlag		 = iCurColumnPos(57)  'CTP Check Flag
			C_SoSeq				 = iCurColumnPos(58)  '수주순번 
			C_PreSoNo			 = iCurColumnPos(59)  '수주번호 for 수주내역참조 
			C_PreSoSeq			 = iCurColumnPos(60)  '수주순번 for 수주내역참조 
    End Select    
End Sub




'===================================   SetQuerySpreadColor()  ======================================
' Name : SetQuerySpreadColor()
' Description : 조회후 그리드 Color
'==================================================================================================== 
Sub SetQuerySpreadColor(ByVal lRow)
	
	Dim SoSts, BillQty
    With frm1

	
    .vspdData.ReDraw = False
   		

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
		
		ggoSpread.SSSetProtected C_LocAmt, -1, -1
   	    ggoSpread.SSSetProtected C_VatLocAmt, -1, -1


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
        
        .vspdData.Row = lRow
		.vspdData.Col = C_SoSts
		SoSts = CInt(Trim(.vspdData.Text))
		
		If Trim(frm1.txtHConfirmFlg.value) = "N" Then
			ggoSpread.SSSetProtected  C_SoQty, -1, -1
			ggoSpread.SSSetProtected  C_DlvyDt, -1, -1
		else
			ggoSpread.SpreadUnLock  C_SoQty, -1, -1
			ggoSpread.SpreadUnLock  C_DlvyDt, -1, -1
			
			ggoSpread.SSSetRequired  C_SoQty, -1, -1
			ggoSpread.SSSetRequired  C_DlvyDt, -1, -1
		End If


' 출고요청일자가 현재일자보다 적은경우 알림표시 
		For lRow = 1 To .vspdData.MaxRows          
			  .vspdData.Row = lRow : .vspdData.Col = C_DnReqDt
			  If UniConvDateToYYYYMMDD(.vspdData.Text,parent.gDateFormat,"") < UniConvDateToYYYYMMDD(EndDate,parent.gDateFormat,"") Then Call sprRedComColor(C_DnReqDt,-1,-1)
		Next       

    .vspdData.ReDraw = True

	'If Trim(.RdoConfirm.value) = "N" Then 

    End With

End Sub
'========================================================================================================
'								3.2 Common Group-2
'========================================================================================================
'========================================================================================================
' Name : Form_Load
' Desc : This sub is called from window_OnLoad in Common.vbs automatically
'========================================================================================================
Sub Form_Load()

	Err.Clear                                                                '☜: Clear err status
	Call LoadInfTB19029                                                      '☜: Load table , B_numeric_format
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 	
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                             '⊙: Lock  Suitable  Field

	Call InitSpreadSheet
	Call InitVariables
	Call SetDefaultVal
	
	Call SetToolbar("11000000000011")								 '⊙: 버튼 툴바 제어 	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call CookiePage(0)
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
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear																	  '☜: Clear error status
    
    FncQuery = False															  '☜: Processing is NG    

	ggoSpread.Source = frm1.vspdData 
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")	  '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "2")								  '☜: Clear Contents  Field
    
    If Not chkField(Document, "1") Then								 '☜: This function check required field
       Exit Function
    End If
    
    '------ Developer Coding part (Start ) --------------------------------------------------------------     
    Call InitVariables												 '⊙: Initializes local global variables

    If DbQuery = False Then                                    '☜: Query db data
       Exit Function
    End If
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncQuery = True                                               '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement
        
End Function

<%
'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
%>
Function FncNew() 
    Dim IntRetCD 
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncNew = False
    '------ Developer Coding part (Start ) -------------------------------------------------------------- 
    '필요없는지???
    If lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X") 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "1")                                 '⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                 '⊙: Clear Contents  Field
    Call ggoOper.LockField(Document, "N")                                  '⊙: Lock  Suitable  Field
    Call SetDefaultVal
    Call InitVariables
    Call SetToolbar("11000000000011")									  '⊙: 버튼 툴바 제어    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncNew = True                                                              '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement

End Function


'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete() 
    Dim intRetCD
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncDelete = False                                                             '☜: Processing is NG
    
    '------ Developer Coding part (Start ) -------------------------------------------------------------- 
    '필요없는지???
    If lgIntFlgMode <> parent.OPMD_UMODE Then      
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If
    
    If DbDelete = False Then                                                '☜: Delete db data
       Exit Function                                                        '☜:
    End If
    
    Call ggoOper.ClearField(Document, "1")                              '⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                              '⊙: Clear Contents  Field
    
    '------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncDelete = True                                                           '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement  
    
End Function

'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave() 
	Dim IntRetCD 
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncSave = False																  '☜: Processing is NG
           
	ggoSpread.Source = frm1.vspdData 
	If ggoSpread.SSCheckChange = False Then								  '☜:match pointer
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")				  '☜:There is no changed data.  
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData      
    IF ggoSpread.SSDefaultCheck = False Then								  '☜: Check contents area
		Exit Function
    End If
    
    
    '------ Developer Coding part (Start ) -------------------------------------------------------------- 
	If Not chkField(Document, "2")  Then                          
       Exit Function
    End If

    '이전수주참조를 하지 않은 반품수주일 경우 Lot 번호를 Assign
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

	Call CheckCreditlimitSvr									'☜: 저장하기전 여신한도 체크 
	    
'------ Developer Coding part (End )   -------------------------------------------------------------- 

    If DbSave = False Then                                                        '☜: Query db data
       Exit Function
    End If

    If Err.number = 0 Then	
       FncSave = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement
    
End Function


'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy() 
	Dim IntRetCD

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCopy = False                                                               '☜: Processing is NG

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

	With frm1 
	  .vspdData.Col = C_TrackingNo :  .vspdData.Text = ""
	  .vspdData.Col = C_SoSupplyQty :  .vspdData.Text = ""
	  .vspdData.Col = C_SoQty   :  .vspdData.Text = ""
	  .vspdData.Col = C_SoPrice  :  .vspdData.Text = 0
	  .vspdData.Col = C_SoPriceAutoChk:  .vspdData.Text = "0"
	  .vspdData.Col = C_SoPriceFlag :  .vspdData.Text = "진단가"
	  .vspdData.Col = C_NetAmt  :  .vspdData.Text = 0
	  .vspdData.Col = C_TotalAmt  :  .vspdData.Text = 0
	  .vspdData.Col = C_LocAmt  :  .vspdData.Text = 0
	  .vspdData.Col = C_VatLocAmt  :  .vspdData.Text = 0

	  .vspdData.Col = C_VATAmt  :  .vspdData.Text = 0
	  .vspdData.Col = C_BonusQty  :  .vspdData.Text = 0
	  .vspdData.Col = C_TolMoreRate :  .vspdData.Text = 0
	  .vspdData.Col = C_TolLessRate :  .vspdData.Text = 0
	  .vspdData.Col = C_Remark  :  .vspdData.Text = ""
	  .vspdData.Col = C_SoSts   :  .vspdData.Text = ""
	  .vspdData.Col = C_BillQty  :  .vspdData.Text = 0
	  .vspdData.Col = C_BaseQty  :  .vspdData.Text = ""
	  .vspdData.Col = C_BonusBaseQty :  .vspdData.Text = ""
	  .vspdData.Col = C_MaintSeq  :  .vspdData.Text = ""
	  .vspdData.Col = C_OrderSeq  :  .vspdData.Text = ""
	  .vspdData.Col = C_APSHost  :  .vspdData.Text = ""
	  .vspdData.Col = C_APSPort  :  .vspdData.Text = ""
	  .vspdData.Col = C_CTPTimes  :  .vspdData.Text = ""
	  .vspdData.Col = C_CTPCheckFlag :  .vspdData.Text = ""
	  '.vspdData.ReDraw = True  삭제가능?
	  
 End With
 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncCopy = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
    
End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
	Dim iDx
	
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCancel = False                                                             '☜: Processing is NG
	
	If frm1.vspdData.MaxRows < 1 Then Exit Function
    ggoSpread.Source = Frm1.vspdData	
    ggoSpread.EditUndo  
	'------ Developer Coding part (Start ) --------------------------------------------------------------
    Call TotalSum														'☜: Protect system from crashing    
    '------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncCancel = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement 
    
End Function


'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
	Dim IntRetCD
    Dim imRow
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
        ggoSpread.InsertRow ,imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        .vspdData.ReDraw = True
        lgBlnFlgChgValue = True  
    End With
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 	
	With frm1   
  
		'###### 2001_11_28 반품일때는 수주일자로. ##########
		.vspdData.Col	= C_DlvyDt
		.vspdData.Text	= .HReqDlvyDt.value
		If Trim(frm1.txtHPreSONo.value) = "" And UCase(Trim(frm1.HRetItemFlag.value)) = "Y" Then
		 .vspdData.Text = .txtSoDt.value
		End If
		'################################################

		.vspdData.Col	= C_ShipToParty
		.vspdData.Text	= .txtShipToParty.value
		.vspdData.Col	= C_TolMoreRate
		.vspdData.Text	= 0
		.vspdData.Col	= C_TolLessRate
		.vspdData.Text	= 0
		.vspdData.Col	= C_SoPrice
		.vspdData.Text	= 0
		.vspdData.Col	= C_NetAmt
		.vspdData.Text	= 0
		.vspdData.Col	= C_BonusQty
		.vspdData.Text	= 0
		.vspdData.Col	= C_MaintSeq
		.vspdData.Text	= 0
		.vspdData.Col	= C_OrderSeq
		.vspdData.Text	= 0
		.vspdData.Col	= C_PlantCd
		.vspdData.Text	= .txtPlant.value 
		.vspdData.Col	= C_SoPriceFlag

		Select Case .HPriceFlag.value
		Case "Y"
		 .vspdData.Text = "진단가"
		Case "N"
		 .vspdData.Text = "가단가"
		End Select

		.vspdData.Col= C_VatType

		If Len(.txtHVATType.value) Then
			.vspdData.text = frm1.txtHVATType.value
			Call SetVatType(.vspdData.ActiveRow)
		End If

		If Len(.txtHVATIncFlag.value) Then
			.vspdData.Col = C_VatIncFlag
			.vspdData.Row = .vspdData.ActiveRow
			.vspdData.text = .txtHVATIncFlag.value

			.vspdData.Col = C_VatIncFlagNm
			.vspdData.Row = .vspdData.ActiveRow

			Select Case .txtHVATIncFlag.value
			Case "1"
			 .vspdData.Text = "별도"
			Case "2"
			 .vspdData.Text = "포함"
			End Select
		End If   

		.vspdData.ReDraw = True   

    End With
  '------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow() 
	Dim lDelRows

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncDeleteRow = False                                                          '☜: Processing is NG

    If Frm1.vspdData.MaxRows < 1 then
       Exit function
	End if	
	
    With Frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

    lgBlnFlgChgValue = True 
    Call TotalSum
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncDeleteRow = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
    
End Function


'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
'========================================================================================================
Function FncPrint() 
	On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncPrint = False                                                              '☜: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()                                                        '☜: Protect system from crashing

    If Err.number = 0 Then	 
       FncPrint = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement
End Function


'========================================================================================================
' Name : FncPrev
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev()
	On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncPrev = False                                                               '☜: Processing is NG
    '--------- Developer Coding Part (Start) ---------------------------------------------------------- 
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      
        Call DisplayMsgBox("900002", "X", "X", "X")  '☜ 바뀐부분 
        Exit Function
    ElseIf lgPrevNo = "" Then
		Call DisplayMsgBox("900011", "X", "X", "X")  '☜ 바뀐부분 
		Exit Function
    End If
    '--------- Developer Coding Part (End) ------------------------------------------------------------

    If Err.number = 0 Then	 
       FncPrev = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement

End Function


'========================================================================================================
' Name : FncNext
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncNext() 
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncNext = False                                                               '☜: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002", "X", "X", "X")  '☜ 바뀐부분 
        Exit Function
    ElseIf lgNextNo = "" Then
		Call DisplayMsgBox("900012", "X", "X", "X")  '☜ 바뀐부분 
		Exit Function
    End If
	'--------- Developer Coding Part (End) ------------------------------------------------------------
    If Err.number = 0 Then	 
       FncNext = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
	On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExcel = False                                                              '☜: Processing is NG

	Call Parent.FncExport(Parent.C_SINGLEMULTI)

    If Err.number = 0 Then	 
       FncExcel = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement 
End Function


'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncFind = False                                                               '☜: Processing is NG

	'Call Parent.FncFind(Parent.C_MULTI, True)
	Call parent.FncFind(parent.C_SINGLEMULTI, False)
    If Err.number = 0 Then	 
       FncFind = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement  
End Function


'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
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
		'Frm1.vspdData.Col = iColumnLimit : Frm1.vspdData.Row = 0  삭제??
		iRet = DisplayMsgBox("900030", "X", Trim(frm1.vspdData.Text), "X")
		Exit Function
    End If   
    
    Frm1.vspdData.ScrollBars = parent.SS_SCROLLBAR_NONE    
    
    ggoSpread.Source = Frm1.vspdData    
    ggoSpread.SSSetSplit(ACol)      
    Frm1.vspdData.Col = ACol
    Frm1.vspdData.Row = ARow
    
    'Frm1.vspdData.Action = 0    
    Frm1.vspdData.Action = Parent.SS_ACTION_ACTIVE_CELL   
    
    Frm1.vspdData.ScrollBars = parent.SS_SCROLLBAR_BOTH
    
End Function

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
	Call ggoSpread.ReOrderingSpreadData()
	
	Call SetQuerySpreadColor(1)    

End Sub

'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExit = False                                                               '☜: Processing is NG
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")		  '⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    If Err.number = 0 Then	 
       FncExit = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function


'========================================================================================================
'								3.3 Common Group-3
'========================================================================================================
'========================================================================================================
' Name : DbQuery
' Desc : This sub is called by FncQuery
'========================================================================================================
Function DbQuery() 

    Dim strVal
	
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    DbQuery = False                                                               '☜: Processing is NG
	
    Call DisableToolBar(Parent.TBC_QUERY)                                                '☜: Disable Query Button Of ToolBar
    Call LayerShowHide(1)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
	'Call MakeKeyStream(pDirect)
    If lgIntFlgMode = parent.OPMD_UMODE Then    
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001				'☜: 비지니스 처리 ASP의 상태 
		strVal = strVal & "&txtConSoNo=" & Trim(frm1.txtHSoNo.value)        '☜: 조회 조건 데이타 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	Else
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001         
		strVal = strVal & "&txtConSoNo=" & Trim(frm1.txtConSoNo.value)     
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	End If 
	'--------- Developer Coding Part (End) ------------------------------------------------------------
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    If Err.number = 0 Then	 
       DbQuery = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement
End Function


'========================================================================================================
' Name : DbSave
' Desc : This sub is called by FncSave
'========================================================================================================
Function DbSave() 

    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    DbSave = False                                                                '☜: Processing is NG

    Call DisableToolBar(Parent.TBC_SAVE)                                   '☜: Disable Save Button Of ToolBar
    Call LayerShowHide(1)                                                         '☜: Show Processing Message

	Frm1.txtMode.value        = Parent.UID_M0002                                  '☜: Delete		
	
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	ggoSpread.Source = frm1.vspdData
	
	lGrpCnt = 0    
	strVal = ""
	strDel = ""
	
	With frm1	
		'.txtUpdtUserId.value = parent.gUsrID
		'.txtInsrtUserId.value = parent.gUsrID	

		For lRow = 1 To .vspdData.MaxRows
			.vspdData.Row = lRow
			.vspdData.Col = 0
	
			Select Case .vspdData.Text
			
				Case ggoSpread.InsertFlag       '☜: 신규 
					strVal = strVal & "C" & parent.gColSep & lRow & parent.gColSep'☜: C=Create
				
				Case ggoSpread.UpdateFlag       '☜: 수정 
					strVal = strVal & "U" & parent.gColSep & lRow & parent.gColSep'☜: U=Update
				
				Case ggoSpread.DeleteFlag       '☜: 삭제 
					strVal = strVal & "D" & parent.gColSep & lRow & parent.gColSep'☜: D=Delete

			End Select

			Select Case .vspdData.Text
			
				Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag,ggoSpread.DeleteFlag
					.vspdData.Col = C_SoSeq        '--- 수주순번	2              
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep                    
              
					.vspdData.Col = C_ItemCd       '--- 품목	3
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep             
              
					.vspdData.Col = C_SoUnit       '--- 수주단위	4  
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep              
              
					.vspdData.Col = C_TrackingNo   '--- Tracking No	5  
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep              
              
					.vspdData.Col = C_SoQty        '--- 수주수량	6
					IF UNIConvNum(Trim(.vspdData.Text), 0) <= 0 Then
					   Call DisplayMsgBox("203233", "X", "X", "X")              
					   Exit Function
					Else
					   strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & parent.gColSep 
					END IF                              
              
					.vspdData.Col = C_SoPrice     '--- 수주단가 7
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & parent.gColSep              
              
					.vspdData.Col = C_SoPriceFlag  '--- 수주 가단가/진단가	8
					
					Select Case Trim(.vspdData.TypeComboBoxCurSel)
					
						Case 1  '진단가 
								strVal = strVal & "Y" & parent.gColSep    
						Case 0  '가단가 
								strVal = strVal & "N" & parent.gColSep    
						Case Else
								MsgBox "가단가/진단가 값이 없습니다.", vbExclamation, parent.gLogoName 
								frm1.vspdData.Row = lRow
								frm1.vspdData.Action = 0
								'--작업중 표시화면 및 마우스 포인트 원복      
								Call BtnDisabled(False)
								Call LayerShowHide(0)
								Exit Function
					End Select
     
					.vspdData.Col = C_NetAmt     '--- 수주금액	9
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & parent.gColSep                            
              
					.vspdData.Col = C_VatAmt     '--- VAT 금액	10
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & parent.gColSep              
              
					.vspdData.Col = C_LocAmt     '--- 수주자국금액 
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & parent.gColSep                            
              
					.vspdData.Col = C_VatLocAmt     '--- VAT자국금액 
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & parent.gColSep              
              
					.vspdData.Col = C_PlantCd     '--- 공장코드	11
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep     
              
					.vspdData.Col = C_DlvyDt     '--- 납기일 12
					strVal = strVal & UNIConvDate(Trim(.vspdData.Text)) & parent.gColSep              
              
					.vspdData.Col = C_ShipToParty  '--- 납품처	13
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep              
              
					.vspdData.Col = C_HsNo      '--- HS번호	14
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep              
              
					.vspdData.Col = C_TolMoreRate  '--- 과부족허용율(+)	15
					strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & parent.gColSep     
              
					.vspdData.Col = C_TolLessRate  '--- 과부족허용율(-)	16
					strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & parent.gColSep              
              
					.vspdData.Col = C_VatType     '--- VAT유형	17
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep              
              
					.vspdData.Col = C_VatRate     '--- VAT율	18
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & parent.gColSep     
              
					.vspdData.Col = C_VatIncFlag   '--- VAT 포함구분	19
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep     
              
					.vspdData.Col = C_RetType     '--- 반품유형	20
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep     
              
					.vspdData.Col = C_LotNo     '--- Lot 번호 21
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep    
              
					.vspdData.Col = C_LotSeq     '--- Lot 순번	22
					strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & parent.gColSep     
              
					.vspdData.Col = C_PreDnNo     '--- 반품출하번호	23
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep     
              
					.vspdData.Col = C_PreDnSeq     '--- 반품출하순번	24
					strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & parent.gColSep     
              
					.vspdData.Col = C_BonusQty     '--- 덤수량	25
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & parent.gColSep     
              
					.vspdData.Col = C_SlCd      '--- 창고코드	26
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep     
              
					.vspdData.Col = C_Remark     '--- 비고	27
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep    
					      
					.vspdData.Col = C_PreSoNo     '--- 반품수주번호	28
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep     
                    
				    .vspdData.Col = C_PreSoSeq     '--- 반품수주순번	29
				    strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & parent.gColSep  					
													'--- 구매발주순번	30
					strVal = strVal & 0 & parent.gColSep 
					
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
		.txtSpread.value = strDel & strVal	
		

   
	End With
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)          '☜: 비지니스 ASP 를 가동 
		
    '------ Developer Coding part (End )   -------------------------------------------------------------- 

    If Err.number = 0 Then	 
       DbSave = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement 
      
End Function


'========================================================================================================
' Name : DbDelete
' Desc : This sub is called by FncDelete
'========================================================================================================
Function DbDelete() 
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    DbDelete = False                                                              '☜: Processing is NG
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	'In Multi, You need not to implement this area

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	 
       DbDelete = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement
End Function


'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Sub DbQueryOk()         
	
	On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    lgIntFlgMode = Parent.OPMD_UMODE    
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
  
	frm1.btnAvlStkRef.disabled = False

'	If Trim(frm1.txtHPreSONo.value) <> "" And UCASE(Trim(frm1.HRetItemFlag.value)) = "Y" Then
'		Call SetToolbar("11101011000111")
'	ElseIf Trim(frm1.txtHPreSONo.value) = "" And UCASE(Trim(frm1.HRetItemFlag.value)) = "Y" Then
'		Call SetToolbar("11101111001111")
'	ElseIf UCASE(Trim(frm1.HRetItemFlag.value)) <> "Y" Then
'		Call SetToolbar("11101111001111")
'	Else
'		Call SetToolbar("11101111001111")
'	End If

	If Trim(frm1.txtHConfirmFlg.value) = "N" Then
		Call SetToolbar("111000000001111")
	 Else
		Call SetToolbar("1110110100111111")	 
	End if

	
	frm1.vspdData.Focus
	
	Call SetQuerySpreadColor(1)    
	Call TotalSum

	lgBlnFlgChgValue = False
    
    Call ChangePlantColor()    
    
	Call ButtonVisible(1)
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call ggoOper.LockField(Document, "Q")
	
    Set gActiveElement = document.ActiveElement  
    
End Sub

'========================================================================================================
' Name : DbSaveOk
' Desc : Called by MB Area when save operation is successful
'========================================================================================================
Sub DbSaveOk()               
	On Error Resume Next                                                   '☜: If process fails
    Err.Clear                                                              '☜: Clear error status

    Call InitVariables													   '⊙: Initializes local global variables
    
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
    frm1.txtConSoNo.value = frm1.txtHSoNo.value
	frm1.vspdData.MaxRows = 0
    Call MainQuery()
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    
    Set gActiveElement = document.ActiveElement
    
End Sub


'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()            
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    Set gActiveElement = document.ActiveElement 
End Function

'========================================================================================================
'								3.4		User-defined Method
'========================================================================================================

'========================================================================================================
' Name : OpenSODtlRef()
' Desc : S/O Reference Window Call 
'========================================================================================================
Function OpenSODtlRef()
	Dim iCalledAspName
	Dim arrRet
	Dim strSONo
  
	If Trim(frm1.txtConSoNo.value) = "" Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If

	If UCase(frm1.HRetItemFlag.value) <> "Y" Then
		Call DisplayMsgBox("203155", "x", "x", "x")
		Exit Function
	End If   

	If UCase(frm1.txtHPreSONo.value) = "" Then
		Call DisplayMsgBox("203156", "x", "x", "x")
		Exit Function
	End If   

	If Trim(frm1.RdoConfirm.value) = "N" Then	
		Msgbox "확정처리된 품목은 수주내역을 참조 할 수 없습니다",vbInformation, parent.gLogoName
		Exit Function
	End If

	strSONo = frm1.txtHPreSONo.value 
	strSONo = strSONo & parent.gRowSep & frm1.txtCurrency.value

	'20021227 kangjungu dynamic popup
	iCalledAspName = AskPRAspName("s3112ra4")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3112ra4", "x")
		IsOpenPop = False
		exit Function
	end if

	IsOpenPop = True
  
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, strSONo), _
	  "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	  

  
	IsOpenPop = False

	If arrRet(0, 0) = "" Then
		Exit Function
	Else
		Call SetSODtlRef(arrRet)
	End If 

End Function


'========================================================================================================
' Name : OpenAvalStockRef()
' Desc : S/O Reference Window Call
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
		Call DisplayMsgBox("202250", "x", "x", "x")				'⊙: "Will you destory previous data" 
		Exit Function
	End If 
  
	frm1.vspdData.Col = C_ItemName   
 
	arrParam(1) = frm1.vspdData.Text

	frm1.vspdData.Col = C_PlantCd
 
	arrParam(2) = frm1.vspdData.Text

	frm1.vspdData.Col = C_PlantNm
  
	arrParam(3) = frm1.vspdData.Text

	'20021227 kangjungu dynamic popup
	iCalledAspName = AskPRAspName("s1912ra1")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s1912ra1", "x")
		IsOpenPop = False
		exit Function
	end if

	IsOpenPop = True
 
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
	  "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
  
	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetSODtlRef(arrRet)
	End If 
	
 End Function


'========================================================================================================
' Name : OpenStockDtlRef()
' Desc : Reference Window Call
'========================================================================================================
Function OpenStockDtlRef()
	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(5)
 
	If UCase(Trim(frm1.txtConSoNo.value)) = "" Then
		Call DisplayMsgBox("900002", "x", "x", "x")		 '⊙: "Will you destory previous data"
		Exit Function
	End If

	ggoSpread.Source = frm1.vspdData
  
	If frm1.vspdData.Maxrows < 1 Then
		Call DisplayMsgBox("209001", "x", "x", "x")		 '⊙: "Will you destory previous data"
		Exit Function
	End If 
  
	frm1.vspdData.Col = C_ItemCd
	frm1.vspdData.Row = frm1.vspdData.ActiveRow

	arrParam(0) = frm1.vspdData.Text

	If arrParam(0) = "" Then
		Call DisplayMsgBox("202250", "x", "x", "x") <% '⊙: "Will you destory previous data" %>
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
 
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam), _
	  "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
  
	IsOpenPop = False

	If arrRet(0, 0) = "" Then
		Exit Function
	End If 
  
End Function

'===========================================================================
' Function Name : OpenItem
' Function Desc : OpenItem Reference Popup
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
	 
	strRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam), _
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
		Call vspdData_Change(C_ItemCd, frm1.vspdData.Row)  ' 변경이 읽어났다고 알려줌 
	End If 

End Function 

'===========================================================================
' Function Name : OpenConSoDtl
' Function Desc : OpenConSoDtl Reference Popup
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
	 
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent, ""), _
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
' Function Name : OpenSoDtl
' Function Desc : OpenSoDtl Reference Popup
'===========================================================================
Function OpenSoDtl(Byval strCode, Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim OriginCol,TempCd

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	 Select Case iWhere
	 Case 0 '품목 
		arrParam(1) = "b_item item, b_plant plant, b_item_by_plant item_plant"   
		arrParam(2) = strCode               
		arrParam(4) = "item.item_cd=item_plant.item_cd and plant.plant_cd=item_plant.plant_cd"
		arrParam(5) = "품목"     
	 
		arrField(0) = "item.item_cd"    
		arrField(1) = "item.item_nm"    
		arrField(2) = "plant.plant_cd"    
		arrField(3) = "plant.plant_nm" 
    
		arrHeader(0) = "품목"      
		arrHeader(1) = "품목명"      
		arrHeader(2) = "공장"      
		arrHeader(3) = "공장명"      

	Case 1 '단위 
		arrParam(1) = "b_unit_of_measure"    
		arrParam(2) = strCode       
		arrParam(4) = ""      		 
		arrParam(5) = "단위"   
		    
		arrField(0) = "unit"       
		arrField(1) = "unit_nm"   	
		  
		arrHeader(0) = "단위"      
		arrHeader(1) = "단위명"      
	
	Case 3 '납품처 
		arrParam(1) = "b_biz_partner bp, b_biz_partner_ftn bp_ftn"   
		arrParam(2) = strCode            
		arrParam(4) = "bp.bp_cd=bp_ftn.partner_bp_cd and bp_ftn.bp_cd= " + FilterVar(frm1.txtSoldToParty.value, "''", "S") + " and bp_ftn.partner_ftn = " & FilterVar("SSH", "''", "S") & " and bp_ftn.usage_flag = " & FilterVar("Y", "''", "S") & " "     
		arrParam(5) = "납품처"      
 
		arrField(0) = "bp_ftn.partner_bp_cd"   
		arrField(1) = "bp.bp_nm"      
		  
		arrHeader(0) = "납품처"      
		arrHeader(1) = "납품처명"    
		 
	Case 4 'HS번호 
		arrParam(1) = "b_hs_code"      
		arrParam(2) = strCode       
		arrParam(4) = ""        
		arrParam(5) = "HS부호"      
 
		arrField(0) = "hs_cd"       
		arrField(1) = "hs_nm"       
		  
		arrHeader(0) = "HS부호"      
		arrHeader(1) = "HS부호명"     
	
	Case 5 '공장 
		With frm1
			OriginCol = .vspdData.Col
			.vspdData.Col = C_ItemCd
			TempCd = .vspdData.Text
			.vspdData.Col = OriginCol
		End With
		arrParam(1) = "b_plant plant, b_item_by_plant item_plant"    
		arrParam(2) = strCode             
		arrParam(4) = "plant.plant_cd=item_plant.plant_cd and item_plant.item_cd = " + FilterVar(TempCd, "''", "S") 
		arrParam(5) = "공장"      
 
		arrField(0) = "plant.plant_cd"     
		arrField(1) = "plant.plant_nm"     
		    
		arrHeader(0) = "공장"      
		arrHeader(1) = "공장명"      
  
	Case 6 '창고 
		With frm1
			OriginCol = .vspdData.Col
			.vspdData.Col = C_PlantCd
			TempCd = .vspdData.Text
			.vspdData.Col = OriginCol
		End With
		arrParam(1) = "b_storage_location"    
		arrParam(2) = strCode       
		arrParam(4) = "plant_cd = " + FilterVar(TempCd, "''", "S")
		arrParam(5) = "창고"      
 
		arrField(0) = "sl_cd"       
		arrField(1) = "sl_nm"       
		  
		arrHeader(0) = "창고"      
		arrHeader(1) = "창고명"      

	Case 7 'VAT유형 
		arrParam(1) = "B_MINOR Minor,B_CONFIGURATION Config"
		arrParam(2) = strCode        
		arrParam(4) = "Minor.MAJOR_CD=" & FilterVar("B9001", "''", "S") & " And Config.MAJOR_CD = Minor.MAJOR_CD" _
		    & " And Config.MINOR_CD = Minor.MINOR_CD" _
		    & " And Config.SEQ_NO = 1"   
		arrParam(5) = "VAT유형"      

		arrField(0) = "Minor.MINOR_CD"      
		arrField(1) = "Minor.MINOR_NM"      
		arrField(2) = "Config.REFERENCE"     
			         
		arrHeader(0) = "VAT유형"      
		arrHeader(1) = "VAT유형명"      
		arrHeader(2) = "VAT율" 
		     
	Case 8 
		arrParam(1) = "B_MINOR"      
		arrParam(2) = strCode       
		arrParam(4) = "MAJOR_CD=" & FilterVar("B9017", "''", "S") & ""        
		arrParam(5) = "반품유형"      
 
		arrField(0) = "Minor_cd"       
		arrField(1) = "Minor_nm"       
			   
		arrHeader(0) = "반품유형"      
		arrHeader(1) = "반품유형명" 
		    
	Case 9
		arrParam(0) = "VAT포함구분"    
		arrParam(1) = "B_MINOR"      
		arrParam(2) = strCode
		arrParam(4) = "MAJOR_CD=" & FilterVar("S4035", "''", "S") & ""    
		arrParam(5) = "VAT포함구분"    
	 
		arrField(0) = "MINOR_CD"     
		arrField(1) = "MINOR_NM"     
		        
		arrHeader(0) = "VAT포함구분"   
		arrHeader(1) = "VAT포함구분명"   
	End Select

	arrParam(0) = arrParam(5)       

	Select Case iWhere
	Case 0 <% '품목 %>
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

'===========================================================================
' Function Name : OpenPlant
' Function Desc : OpenPlant Reference Popup
'===========================================================================
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	If frm1.txtPlant.readOnly = True Then Exit Function

	IsOpenPop = True

	arrParam(1) = "B_PLANT"       
	arrParam(2) = Trim(frm1.txtPlant.value)  
	arrParam(4) = ""       
	arrParam(5) = "공장"    
 
	arrField(0) = "PLANT_CD"    
	arrField(1) = "PLANT_NM"    
	   
	arrHeader(0) = "공장"     
	arrHeader(1) = "공장명"    

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPlant(arrRet)
	End If 
	
End Function

'========================================================================================================
' Name : SetSODtlRef()
' Desc : Set Return array from S/O Reference Window 
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

		TempRow = .vspdData.MaxRows           <% '☜: 현재까지의 MaxRows %>
		intLoopCnt = Ubound(arrRet, 1)          <% '☜: Reference Popup에서 선택되어진 Row만큼 추가 %>

		strSOJungBokMsg = ""
		 
		For intCnt = 1 to intLoopCnt + 1
			blnEqualFlg = False

			If TempRow <> 0 Then

				strSoNo=""
				strSoSeqNo=""

					For j = 1 To TempRow

						<% '수주번호 %>
						.vspdData.Row = j
						.vspdData.Col = C_PreSoNo
						strSoNo = .vspdData.text

						If strSoNo = arrRet(intCnt - 1, 0) Then

							<% '수주순번 %>
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
	<%'자국금액 %>
				.vspdData.Col = C_LocAmt										
				.vspdData.text = arrRet(intCnt - 1, 10)
	<%'소비세자국금액 %>
				.vspdData.Col = C_VatLocAmt										
				.vspdData.text = arrRet(intCnt - 1, 11)
				
				.vspdData.Col = C_VATType
				.vspdData.text = arrRet(intCnt - 1, 13)
				.vspdData.Col = C_TrackingNo
				.vspdData.text = arrRet(intCnt - 1, 15)
    
				'###### 2001_11_28 반품일때는 수주일자로. #######
				.vspdData.Col = C_DlvyDt
				'.vspdData.text = arrRet(intCnt - 1, 16)
				.vspdData.text = Trim(frm1.txtSoDt.value)
				'################################################

				.vspdData.Col = C_ShipToParty
				.vspdData.text = arrRet(intCnt - 1, 17)
				.vspdData.Col = C_SoPriceFlag

				If Trim(arrRet(intCnt - 1, 18)) = "Y" Then
					.vspdData.text = "진단가"
				ElseIf  Trim(arrRet(intCnt - 1, 18)) = "N" Then
					.vspdData.text = "가단가"
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
					 .vspdData.text = "별도"
				ElseIf  Trim(arrRet(intCnt - 1, 29)) = "2" Then
					 .vspdData.text = "포함"
				End If      

				SetSpreadColor CLng(TempRow) + CLng(intCnt),CLng(TempRow) + CLng(intCnt)
				SetVatType (intCnt)
				lgBlnFlgChgValue = True
			End If
		Next
		
		.vspdData.ReDraw = True
		
		Call JungBokMsg(strSOJungBokMsg,"수주번호" & "-" & "수주순번")
		Call QtyPriceChange()
		
	End With
	
End Function


'========================================================================================================
' Name : SetSoDtl()
' Desc : SetSoDtl Popup에서 Return되는 값 setting 
'========================================================================================================
Function SetSoDtl(Byval arrRet,ByVal iWhere)

	With frm1

		Select Case iWhere
		Case 0 '품목 
		 .vspdData.Col = C_ItemCd
		 .vspdData.Text = arrRet(0)
		 .vspdData.Col = C_ItemName
		 .vspdData.Text = arrRet(1)
		 .vspdData.Col = C_PlantCd
		 .vspdData.Text = arrRet(2)
		 Call vspdData_Change(C_ItemCd, .vspdData.Row)  <% ' 변경이 읽어났다고 알려줌 %>
		 
		Case 1 '단위 
		 .vspdData.Col = C_SoUnit
		 .vspdData.Text = arrRet(0)
		 Call vspdData_Change(C_SoUnit, .vspdData.Row)  <% ' 변경이 읽어났다고 알려줌 %>
		 
		Case 2 '납기일 
	 
		Case 3 '납품처 
		 .vspdData.Col = C_ShipToParty
		 .vspdData.Text = arrRet(0)
		 Call vspdData_Change(C_ShipToParty, .vspdData.Row) <% ' 변경이 읽어났다고 알려줌 %>
		 
		Case 4 'HS번호 
		 .vspdData.Col = C_HsNo
		 .vspdData.Text = arrRet(0)
		 Call vspdData_Change(C_HsNo, .vspdData.Row)   <% ' 변경이 읽어났다고 알려줌 %>
		 
		Case 5 '공장 
		 .vspdData.Col = C_PlantCd
		 .vspdData.Text = arrRet(0)
		 Call vspdData_Change(C_PlantCd, .vspdData.Row)  <% ' 변경이 읽어났다고 알려줌 %>
		 
		Case 6 '창고 
		 .vspdData.Col = C_SlCd
		 .vspdData.Text = arrRet(0)
		 Call vspdData_Change(C_SlCd, .vspdData.Row)   <% ' 변경이 읽어났다고 알려줌 %>
	 
		Case 7 'VAT유형 
		 .vspdData.Col = C_VatType
		 .vspdData.Text = arrRet(0)
		 .vspdData.Col = C_VatTypeNm
		 .vspdData.Text = arrRet(1)
		 .vspdData.Col = C_VatRate
		 .vspdData.text = UNIFormatNumber(UNICDbl(arrRet(2)), ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
		 Call vspdData_Change(C_VatType, .vspdData.Row)   <% ' 변경이 읽어났다고 알려줌 %>
	 
		Case 8 '반품유형 
		 .vspdData.Col = C_RetType
		 .vspdData.Text = arrRet(0)
		 .vspdData.Col = C_RetTypeNm
		 .vspdData.Text = arrRet(1)
		 Call vspdData_Change(C_RetType, .vspdData.Row)   <% ' 변경이 읽어났다고 알려줌 %>
	 
		Case 9 'VAT 포함구분 
		 .vspdData.Col = C_VatIncFlag
		 .vspdData.Text = arrRet(0)
		 .vspdData.Col = C_VatIncFlagNm
		 .vspdData.Text = arrRet(1)
		 Call vspdData_Change(C_VatIncFlag, .vspdData.Row)   <% ' 변경이 읽어났다고 알려줌 %>
	 
		Case Else
		 Exit Function
		End Select
	 
	End With

	lgBlnFlgChgValue = True
 
End Function


'========================================================================================================
' Name : SetPlant()
' Desc : SetPlant Popup에서 Return되는 값 setting
'========================================================================================================
Function SetPlant(Byval arrRet)
	With frm1
		.txtPlant.value = arrRet(0) 
		.txtPlantNm.value = arrRet(1)   
	End With
End Function


'========================================================================================================
' Name : SetDefaultPlant()
' Desc : Default Plant Cd setting
'========================================================================================================
Function SetDefaultPlant()
	With frm1
		.txtPlant.value = parent.gPlant 
		.txtPlantNm.value = parent.gPlantNm   
	End With
End Function


'========================================================================================================
' Name : JungBokMsg()
' Desc : 
'========================================================================================================
Function JungBokMsg(strJungBok,strID)

	Dim strJugBokMsg

	If Len(Trim(strJungBok)) Then strJungBok = strID & Chr(13) & String(30,"=") & strJungBok
	If Len(Trim(strJungBok)) Then strJugBokMsg = strJungBok & Chr(13) & Chr(13)
	If Len(Trim(strJugBokMsg)) Then
		strJugBokMsg = strJugBokMsg & "이미 동일한 번호와 순번이 존재합니다"
		MsgBox strJugBokMsg, vbInformation, parent.gLogoName
	End If

End Function


'========================================================================================================
' Name : HideLotRetField()
' Desc : Combo Display
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
' Name : InitCollectType()
' Desc : 소비세유형코드/명/율 저장하기 
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
' Name : GetCollectTypeRef()
' Desc : 
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
' Name : SetVatType()
' Desc : 
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
' Name : QtyPriceChange()
' Desc : 수량 * 단가 = 금액 
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
' Name : TotalAmtChange()
' Desc : 금액 
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
' Name : TotalSum()
' Desc : 수주순금액 자동합계 
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
' Name : PricePadChange()
' Desc : 품목/단위수량별 단가 자동 Pad
'========================================================================================================
Function PricePadChange(PRow)

	If PricePadCheckMsg(PRow) = False Then Exit Function       <% '단가Pad 호출시 필요사항 체크 %>

	Dim strval

	If LayerShowHide(1) = False Then
		Exit Function
	End If

	strVal = ""    
	strVal = BIZ_PGM_ID & "?txtMode=" & lsPricePad			  
	strVal = strVal & "&lsItemCode=" & lsItemCode			  <%'☜: Batch 조건 데이타 %>
	strVal = strVal & "&lsSoUnit=" & lsSoUnit
	strVal = strVal & "&lsSoQty=" & lsSoQty
	strVal = strVal & "&lsPriceQty=" & lsPriceQty
	strVal = strVal & "&PRow=" & PRow
	strVal = strVal & "&txtHSoNo=" & Trim(frm1.txtHSoNo.value)
	strVal = strVal & "&txtCurrency=" & Trim(frm1.txtCurrency.value)

	Call RunMyBizASP(MyBizASP, strVal)							

End Function

'========================================================================================================
' Name : PricePadCheckMsg()
' Desc : 품목/단위수량별 단가 자동 Pad 호출하기전 체크사항 
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
' Name : ButtonVisible()
' Desc : Button Enable/Disable Flag Function
'========================================================================================================
Function ButtonVisible(ByVal BRow)

	ButtonVisible = False

	If frm1.txtHConfirmFlg.value = "N" AND frm1.vspdData.Maxrows > 1 Then
		frm1.vspdData.Row = BRow
		frm1.vspdData.Col = C_APSHost  : lsAPSHost = frm1.vspdData.Text
		frm1.vspdData.Col = C_APSPort  : lsAPSPort = frm1.vspdData.Text
		frm1.vspdData.Col = C_CTPTimes  : lsCTPTimes = UNICDbl(frm1.vspdData.Text)
		frm1.vspdData.Col = C_CTPCheckFlag : lsCTPCheckFlag = frm1.vspdData.Text

'		If lsCTPCheckFlag = "Y" Then
'			frm1.btnCTPCheck.disabled = False
'		Else
'			frm1.btnCTPCheck.disabled = True
'		End If

	Else
'		frm1.btnCTPCheck.disabled = True
	End If
 
'	If frm1.txtHConfirmFlg.value = "N" And lgIntFlgMode = parent.OPMD_UMODE Then
'		frm1.btnATPCheck.disabled = False
'		If UCASE(frm1.HRetItemFlag.value) = "Y" Then frm1.btnATPCheck.disabled = True
'	End If     

	ButtonVisible = True

End Function


'========================================================================================================
' Name : BtnSpreadCheck()
' Desc : Before Button Click, Spread Check Function
'========================================================================================================
Function BtnSpreadCheck()

	BtnSpreadCheck = False

	Dim IntRetCD

	ggoSpread.Source = frm1.vspdData 

	<% '-- 멀티일때 -- %>
	<% '변경이 있을떄 저장 여부 먼저 체크후, YES이면 작업진행여부 체크 안한다 %>
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")                
		If IntRetCD = vbNo Then Exit Function
	End If

	<% '변경이 없을때 작업진행여부 체크 %>
	If ggoSpread.SSCheckChange = False Then
		IntRetCD = DisplayMsgBox("900018", parent.VB_YES_NO, "X", "X")                
		If IntRetCD = vbNo Then Exit Function
	End If

	BtnSpreadCheck = True

End Function


'========================================================================================================
' Name : CookiePage()
' Desc : JUMP시 Load화면으로 조건부로 Value
'========================================================================================================
Function CookiePage(Byval Kubun)
	Const CookieSplit = 4877						<%'Cookie Split String : CookiePage Function Use%>
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
' Name : JumpChgCheck()
' Desc : Jump시 데이타 변경여부 체크 
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
' Name : ItemByHScodeChange()
' Desc : 품목별에 따른 HS Code 자동 Pad
'========================================================================================================
Function ItemByHScodeChange(CRow)

	Dim strVal

	frm1.vspdData.Row = CRow

	<% '품목 %>
	frm1.vspdData.Col = C_ItemCd
	If Trim(frm1.vspdData.Text) = "" Then Exit Function

	If LayerShowHide(1) = False Then
		Exit Function
	End If

	strVal = ""    
	strVal = BIZ_PGM_ID & "?txtMode=" & "ItemByHsCode"        
	<% '품목 %>
	frm1.vspdData.Col = C_ItemCd
	strVal = strVal & "&ItemCd=" & Trim(frm1.vspdData.Text)
	<% '현재 ROW 위치 %>
	strVal = strVal & "&CRow=" & CRow

	Call RunMyBizASP(MyBizASP, strVal)            

End Function



'========================================================================================================
' Name : CheckCreditlimitSvr()
' Desc : 여신한도 초과 여부 체크 
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
' Name : RunAutoDN()
' Desc : 자동출하생성 
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
' Name : sprRedComColor()
' Desc : 출하요청일자가 현재 일자보다 적을떄 적색 신호...
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
' Name : BizProcessCheck()
' Desc : Biz ASP Call Check Logic
'========================================================================================================
Function BizProcessCheck()

	BizProcessCheck = False

	If window.document.all("MousePT").style.visibility = "visible" Then Exit Function

	BizProcessCheck = True

End Function


'========================================================================================================
' Name : CurFormatNumericOCX()
' Desc : 화면에서 일괄적으로 Rounding되는 Numeric OCX
'========================================================================================================
Sub CurFormatNumericOCX()
	With frm1
	 '수주순금액 
		ggoOper.FormatFieldByObjectOfCur .txtNetAmt, .txtCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
	 
	End With
End Sub


'========================================================================================================
' Name : CurFormatNumSprSheet()
' Desc : 화면에서 일괄적으로 Rounding되는 Numeric Spread Sheet
'========================================================================================================
Sub CurFormatNumSprSheet()
	With frm1

		ggoSpread.Source = frm1.vspdData
		'단가 
		ggoSpread.SSSetFloatByCellOfCur C_SoPrice,-1, .txtCurrency.value, parent.ggUnitCostNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec,,,"Z"
		'금액 
		ggoSpread.SSSetFloatByCellOfCur C_NetAmt,-1, .txtCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloatByCellOfCur C_TotalAmt,-1, .txtCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec,,,"Z"
		
	End With
End Sub



'========================================================================================================
' Name : ChangePlantColor()
' Desc : 수주확정여부가 확정일때 Protected
'========================================================================================================
Sub ChangePlantColor()
	If frm1.txtHConfirmFlg.value = "Y" Then
		Call ggoOper.SetReqAttr(frm1.txtPlant, "Q")
	Else
		Call ggoOper.SetReqAttr(frm1.txtPlant, "D")
	End If 
End Sub


'========================================================================================================
' Name : GetItemPrice()
' Desc : 품목/단위수량별 단가 자동 Pad
'========================================================================================================
Function GetItemPrice(IRow)
	Dim strSoldToParty, strItemCd, strSOUnit, strPayMeth, strDealType, strCurrency, strValidDt
	Dim strSelectList, strFromList, strWhereList
	Dim strRs, strItemInfo

	With frm1
		.vspdData.Row = IRow
		.vspdData.col = C_ItemCd      '품목코드 
		strItemCd = .vspdData.text
		.vspdData.Col = C_SoUnit      '단위 
		strSOUnit = .vspdData.text

		strSoldToParty = .txtSoldToParty.value      '주문처 
		strPayMeth = .txtHPayTermsCd.value    '결제방법 
		strCurrency = .txtCurrency.value    '화폐단위 
		strValidDt = UniConvDateToYYYYMMDD(.txtSoDt.value, parent.gDateFormat,"")
		strDealType = .txtHDealType.value
 	End With

	If Len(Trim(strItemCd)) = 0 Or Len(Trim(strSOUnit)) = 0 Then Exit Function
	
	strSelectList = " dbo.ufn_s_GetItemSalesPrice( " & FilterVar(strSoldToParty, "''", "S") & ",  " & FilterVar(strItemCd, "''", "S") & ", " & FilterVar(strDealType, "''", "S") & ",  " & FilterVar(strPayMeth, "''", "S") & "," & _
	    " " & FilterVar(strSOUnit, "''", "S") & ",  " & FilterVar(strCurrency, "''", "S") & ",  " & FilterVar(strValidDt, "''", "S") & ")"
	strFromList  = ""
	strWhereList = ""

    Err.Clear
    
	'품목정보 단가 Fetch
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
'								3.5		Tag Event
'========================================================================================================



'========================================================================================================
'   Event Name : btnAvlStkRef_OnClick
'   Event Desc : 가용재고현황 버튼을 클릭할 경우 발생 
'========================================================================================================
Sub btnAvlStkRef_OnClick()
	Call OpenAvalStockRef()
End Sub 



'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
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
'   Event Name : vspdData_Change
'   Event Desc :
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
'   Event Name : vspdData_Click
'   Event Desc : 
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
	   
'       frm1.btnATPCheck.disabled = True
'	   frm1.btnCTPCheck.disabled = True
       Exit Sub
    End If
    
    '------ Developer Coding part (Start ) -------------------------------------------------------------- 
	If ButtonVisible(Row) = False Then Exit Sub    <% 'CTP 대상여부에 따른 버튼 체크 %>
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
'	If frm1.RdoConfirm.Value = "Y" Then   
'		Call SetPopupMenuItemInf("0101111111")   
'	Else
'		Call SetPopupMenuItemInf("0000111111")   
'	End IF
'son	
End Sub

'==========================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc :
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
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
       frm1.vspdData.Row=Row
       frm1.vspdData.Col=Col
       iColumnName = frm1.vspdData.Text

       iColumnName = AskSpdSheetColumnName(iColumnName)
        
       If iColumnName <> "" Then
          ggoSpread.Source = frm1.vspdData
          Call ggoSpread.SSSetReNameHeader(Col,iColumnName)
       End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    End If
	
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


'==========================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) --------------------------------------------------------------
End Sub


'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    'If Col <= C_SNm Or NewCol <= C_SNm Then
     '   Cancel = True
      '  Exit Sub
   ' End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub



'필요한지??
'========================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : 
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
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	
	If OldLeft <> NewLeft Then
		Exit Sub
	End If

	If frm1.vspdData.MaxRows < NewTop + C_SHEETMAXROWS And lgStrPrevKey <> "" Then 
	
		If CheckRunningBizProcess = True Then
			Exit Sub
		End If 
	  
		Call DisableToolBar(parent.TBC_QUERY)
		
		If DBQuery = False Then
			Call RestoreToolBar()
			Exit Sub
		End If
		
	End if    
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
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>수주내역Amend</font></td>
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
          </TR>
      </TABLE>
     </TD>
     <TD WIDTH=* align=right><A href="vbscript:OpenStockDtlRef">재고현황참조</A>&nbsp;|&nbsp;<A href="vbscript:OpenSoDtlRef">수주내역참조</A></TD>
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
         <TD CLASS="TD5" NOWRAP>수주번호</TD>
         <TD CLASS="TD6"><INPUT NAME="txtConSoNo" ALT="수주번호" TYPE="Text" MAXLENGTH=18 SiZE=34 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSSoDtl" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSoDtl()"></TD>
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
        <TD CLASS="TD5" NOWRAP>주문처</TD>
        <TD CLASS="TD6"><INPUT NAME="txtSoldToParty" ALT="주문처" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="24XXXU">&nbsp;<INPUT NAME="txtSoldToPartyNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24"></TD>
        <TD CLASS="TD5" NOWRAP>고객주문번호</TD>
        <TD CLASS="TD6"><INPUT NAME="txtCustPoNo" ALT="고객주문번호" TYPE="Text" MAXLENGTH=20 SiZE=20 tag="24XXXU"></TD>
       </TR>
       <TR>
        <TD CLASS="TD5" NOWRAP>수주순금액</TD>
        <TD CLASS="TD6">
         <TABLE CELLSPACING=0 CELLPADDING=0>
          <TR>
           <TD>
            <script language =javascript src='./js/s3112ma3_fpDoubleSingle1_txtNetAmt.js'></script>
           </TD>
           <TD>
            &nbsp;<INPUT NAME="txtCurrency" ALT="" TYPE="Text" MAXLENGTH=3 SiZE=4 tag="24XXXU">
           </TD>
          </TR>
         </TABLE>
        </TD>
        <TD CLASS="TD5" NOWRAP>부가세구분</TD>
        <TD CLASS="TD6"><INPUT NAME="txtVatIncFlag" ALT="부가세구분" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="24XXXU">&nbsp;<INPUT NAME="txtVatIncFlagNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24"></TD>
       </TR>
       <TR>
        <TD CLASS="TD5" NOWRAP>공장</TD>
        <TD CLASS="TD6" NOWRAP><INPUT NAME="txtPlant" ALT="공장" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLCType" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenPlant()">&nbsp;<INPUT NAME="txtPlantNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24"></TD>
        <TD CLASS="TD5" NOWRAP></TD>
        <TD CLASS="TD6" NOWRAP></TD>
       </TR>
       <TR>
        <TD HEIGHT="100%" WIDTH="100%" COLSPAN=4>
         <script language =javascript src='./js/s3112ma3_OBJECT1_vspdData.js'></script>
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
     <!--
      <BUTTON NAME="btnConfirm" CLASS="CLSSBTN">확정처리</BUTTON>&nbsp;
      <BUTTON NAME="btnATPCheck" CLASS="CLSSBTN">ATP Check</BUTTON>&nbsp;
      <BUTTON NAME="btnCTPCheck" CLASS="CLSSBTN">CTP Check</BUTTON>&nbsp;
      <BUTTON NAME="btnDNCheck" CLASS="CLSSBTN">출하요청처리</BUTTON>&nbsp;
      -->
      <BUTTON NAME="btnAvlStkRef" CLASS="CLSSBTN">가용재고현황</BUTTON>
     </TD>
     <TD WIDTH=* Align=right><A HREF = "VBSCRIPT:JumpChgCheck(BIZ_PGM_JUMP_SOHDR_ID)">수주등록</A>&nbsp;|&nbsp;<A HREF = "VBSCRIPT:JumpChgCheck(BIZ_PGM_JUMP_SOSCHE_ID)">회답납기조회</A></TD>
     <TD WIDTH=10>&nbsp;</TD>
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
<INPUT TYPE=HIDDEN NAME="txtSoDt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtSoType" tag="24" TABINDEX="-1">  
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

</FORM>
<DIV ID="MousePT" NAME="MousePT">
 <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>
