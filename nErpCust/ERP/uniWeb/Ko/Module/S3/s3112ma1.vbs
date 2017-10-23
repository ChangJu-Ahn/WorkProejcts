'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Const BIZ_PGM_ID				= "s3112mb1.asp"									
Const BIZ_PGM_JUMP_SOHDR_ID		= "s3111ma1"
Const BIZ_PGM_JUMP_SOSCHE_ID	= "s3114ma8"

Const C_MAX_FORM_ARRAY_DATA		= 100
Const C_MAX_FORM_LIMIT_DATA		= 102399
 
Dim C_ItemCd			 '품목  
Dim C_ItemPopup			 '품목팝업 
Dim C_ItemName			 '품목명 
Dim C_ItemSpec			 
Dim C_SoUnit			 '단위 
Dim C_SoUnitPopup		 '단위팝업 
Dim C_TrackingNo		 'Tracking No
Dim C_TrackingNoPopup
Dim C_SoSupplyQty		 '수주잔량(ReqQty - SoQty)
Dim C_SoQty				 '수량 
Dim C_SoPrice			 '단가 
Dim C_SoPriceAutoChk	 '단가자동계산체크 
Dim C_SoPriceFlag		 '가단가/진단가 
Dim C_TotalAmt			 '화면 수주금액(거래화폐)
Dim C_NetAmt			 'Hidden 수주금액 
Dim C_VATAmt				
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
'--- V3.0 : Tracking No 수동채번 
Dim C_TrackingFlg

Dim ext1_qty 
Dim ext2_qty 
Dim ext3_qty 
Dim ext1_amt 
Dim ext2_amt 
Dim ext3_amt 
Dim ext1_cd  
Dim ext2_cd 
Dim ext3_cd  

Dim C_OldNetAmt			'Hidden 수주금액(속도개선)
Dim C_OriginalNetAmt	'Hidden 수주금액(속도개선)

Dim IsOpenPop 
Dim BaseDate

BaseDate     = "<%=GetSvrDate%>"  

Dim EndDate
Dim StartDate
'초기화면에 뿌려지는 마지막 날짜 (Convert DB date type to Company)
EndDate		= UniConvDateAToB(BaseDate, parent.gServerDateFormat, parent.gDateFormat)

'초기화면에 뿌려지는 시작 날짜 
StartDate	= UNIDateAdd("m", -1, EndDate, parent.gDateFormat)

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
	C_ItemCd			 = 1   '품목  
	C_ItemPopup			 = 2   '품목팝업 
	C_ItemName			 = 3   '품목명 
	C_ItemSpec			 = 4
	C_SoUnit			 = 5   '단위 
	C_SoUnitPopup		 = 6   '단위팝업 
	C_TrackingNo		 = 7   'Tracking No
	C_TrackingNoPopup	 = 8   'Tracking No
	C_SoSupplyQty		 = 9   '수주잔량(ReqQty - SoQty)
	C_SoQty				 = 10   '수량 
	C_SoPrice			 = 11  '단가 
	C_SoPriceAutoChk	 = 12  '단가자동계산체크 
	C_SoPriceFlag		 = 13  '가단가/진단가 
	C_TotalAmt			 = 14  '화면 수주금액(거래화폐)
	C_NetAmt			 = 15  'Hidden 수주금액 
	C_VATAmt			 = 16
	C_PlantCd			 = 17  '공장코드 
	C_PlantCdPopup		 = 18  '공장팝업 
	C_PlantNm			 = 19  '공장명 
	C_DlvyDt			 = 20  '납기일       
	C_ShipToParty		 = 21  '납품처 
	C_ShipToPartyPopup	 = 22  '납품처팝업 
	C_ShipToPartyNm		 = 23  '납품처명 
	C_HsNo				 = 24  'HS번호 
	C_HsNoPopup			 = 25  'HS번호 Popup
	C_TolMoreRate		 = 26  '과부족허용율(+)
	C_TolLessRate		 = 27  '과부족허용율(-)
	C_VatType			 = 28
	C_VatTypePopup		 = 29
	C_VatTypeNm			 = 30
	C_VatRate			 = 31
	C_VatIncFlag		 = 32
	C_VatIncFlagNm		 = 33
	C_RetType			 = 34
	C_RetTypePopup		 = 35
	C_RetTypeNm			 = 36
	C_LotNo				 = 37
	C_LotSeq			 = 38
	C_PreDnNo			 = 39  '출하번호 for 수주내역참조 
	C_PreDnSeq			 = 40  '출하순번 for 수주내역참조 
	C_DnReqDt			 = 41  '출하요청일자 
	C_BonusQty			 = 42  '할증수량(덤)        
	C_SlCd				 = 43  '창고코드 
	C_SlCdPopup			 = 44  '창고팝업 
	C_SlNm				 = 45  '창고명 
	C_Remark			 = 46  '비고 
	C_SoSts				 = 47  '수주진행상태 
	C_BillQty			 = 48  '매출수량 
	C_BaseQty			 = 49  '재고수량 
	C_BonusBaseQty		 = 50  '덤재고수량 
	C_MaintSeq			 = 51  '관리순번 
	C_OrderSeq			 = 52  '주문서순번 
	C_APSHost			 = 53  'APS Host
	C_APSPort			 = 54  'APS Port
	C_CTPTimes			 = 55  'CTP Check 횟수 
	C_CTPCheckFlag		 = 56  'CTP Check Flag
	C_SoSeq				 = 57  '수주순번 
	C_PreSoNo			 = 58  '수주번호 for 수주내역참조 
	C_PreSoSeq			 = 59  '수주순번 for 수주내역참조 		
	'--- V3.0 : Tracking No 수동채번 
	C_TrackingFlg			= 60
	
	C_OldNetAmt			 = 61	'Hidden 수주금액(속도개선)
	C_OriginalNetAmt	 = 62	'Hidden 수주금액(속도개선)
	
		
	ext1_qty =	0 
	ext2_qty =	0
	ext3_qty =	0	'발주순번 BOM참조 
	ext1_amt =	0
	ext2_amt =	0
	ext3_amt =	0
	ext1_cd  =	""
	ext2_cd	 =	""
	ext3_cd  =	""	'발주번호 BOM참조 

End Sub

'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = Parent.OPMD_CMODE						'⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed    
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction
End Sub


'========================================================================================================
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	frm1.txtConSoNo.focus
	frm1.btnConfirm.disabled	= True
	frm1.btnConfirm.value		= "확정처리"
	frm1.btnDNCheck.disabled	= True
	frm1.btnATPCheck.disabled	= True
	frm1.btnCTPCheck.disabled	= True
	frm1.btnAvlStkRef.disabled	= True
	frm1.txtPlant.value			= parent.gPlant
	frm1.txtPlantNm.value		= parent.gPlantNm  
	'-------------------------------------
	' V3.0 : Tracking No 수동채번 
	'-------------------------------------
	Call GetNumberingRuleforTracking()
		
	lgBlnFlgChgValue			= False
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()
	
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		'patch version
		ggoSpread.Spreadinit "V20031214",,parent.gAllowDragDropSpread
		.ReDraw = false	
		
	  '.MaxCols	=	C_PreSoSeq	+  1							' ☜: Add 1 to Maxcols  	  
	  '--- V3.0 : Tracking No 수동채번 
	  .MaxCols = C_OriginalNetAmt + 1
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
	  ggoSpread.SSSetButton	C_TrackingNoPopup
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
	  ggoSpread.SSSetEdit	C_LotNo,			"LOT NO",		12,		,					,	  25,	  2
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
	  ggoSpread.SSSetEdit	C_Remark,			"비고",			60,		,					,	  120
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
	  '--- V3.0 : Tracking No 수동채번 
	  ggoSpread.SSSetEdit C_TrackingFlg, "Tracking 관리여부", 5,,,1,2
	  
	  ggoSpread.SSSetFloat	C_OldNetAmt,		"순금액",		15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
	  ggoSpread.SSSetFloat	C_OriginalNetAmt,	"순금액",		15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
	  
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
      '--- V3.0 : Tracking No 수동채번 
      Call ggoSpread.SSSetColHidden(C_TrackingFlg,C_TrackingFlg,True)      
      Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)				'☜: 공통콘트롤 사용 Hidden Column
	  
	  Call ggoSpread.SSSetColHidden(C_OldNetAmt,C_OldNetAmt,True)
	  Call ggoSpread.SSSetColHidden(C_OriginalNetAmt,C_OriginalNetAmt,True)
	              
	  .ReDraw = true
   
   End With
    
End Sub


'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)

    With frm1
		.vspdData.ReDraw = False
								   'Col				Row         Row2			
		ggoSpread.SSSetRequired		C_ItemCd,		pvStartRow, pvEndRow	
		ggoSpread.SSSetProtected	C_ItemName,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_ItemSpec,		pvStartRow, pvEndRow	
		ggoSpread.SSSetRequired		C_SoUnit,		pvStartRow, pvEndRow		
		ggoSpread.SSSetProtected	C_TrackingNo,	pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_TrackingNoPopup,	pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_SoSupplyQty,	pvStartRow, pvEndRow		
		ggoSpread.SSSetRequired		C_SoPriceFlag,	pvStartRow, pvEndRow		
		ggoSpread.SSSetRequired		C_SoQty,		pvStartRow, pvEndRow		
		ggoSpread.SSSetRequired		C_DlvyDt,		pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C_ShipToParty,	pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C_PlantCd,		pvStartRow, pvEndRow		
		ggoSpread.SSSetRequired		C_TotalAmt,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_NetAmt,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_DnReqDt,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_VATAmt,		pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C_VatType,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_VatTypeNm,	pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_VatRate,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_VatIncFlag,	pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C_VatIncFlagNm, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_RetTypeNm,	pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_LotNo,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_LotSeq,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_PreDNNo,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_PreDNSeq,		pvStartRow, pvEndRow

		If frm1.HRetItemFlag.value = "Y" Then
			ggoSpread.SSSetRequired	C_RetType,	pvStartRow, pvEndRow
		Else
			ggoSpread.SSSetProtected C_RetType,	pvStartRow, pvEndRow
		End If 
 
		' 수출수주/국내수주 여부에 따라 덤수량,HS부호 관리 

		If Trim(.HExportFlag.value) = "Y" Or Trim(.HCiFlag.value) = "Y" Then 
			ggoSpread.SSSetProtected	C_BonusQty, pvStartRow, pvEndRow		
			ggoSpread.SSSetRequired		C_HsNo,		pvStartRow, pvEndRow
		Else  
		    ggoSpread.SSSetProtected	C_HsNo,		pvStartRow, pvEndRow
			ggoSpread.SSSetProtected	C_HsNoPopup,pvStartRow, pvEndRow
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
              Call SetActiveCell(frm1.vspdData, iDx, iRow,"M","X","X")			
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
			C_ItemCd			 = iCurColumnPos(1)   '품목  
			C_ItemPopup			 = iCurColumnPos(2)   '품목팝업 
			C_ItemName			 = iCurColumnPos(3)   '품목명 
			C_ItemSpec			 = iCurColumnPos(4)
			C_SoUnit			 = iCurColumnPos(5)   '단위 
			C_SoUnitPopup		 = iCurColumnPos(6)   '단위팝업 
			C_TrackingNo		 = iCurColumnPos(7)   'Tracking No
			C_TrackingNoPopup	 = iCurColumnPos(8)   'Tracking No
			C_SoSupplyQty		 = iCurColumnPos(9)   '수주잔량(ReqQty - SoQty)
			C_SoQty				 = iCurColumnPos(10)   '수량 
			C_SoPrice			 = iCurColumnPos(11)  '단가 
			C_SoPriceAutoChk	 = iCurColumnPos(12)  '단가자동계산체크 
			C_SoPriceFlag		 = iCurColumnPos(13)  '가단가/진단가 
			C_TotalAmt			 = iCurColumnPos(14)  '화면 수주금액(거래화폐)
			C_NetAmt			 = iCurColumnPos(15)  'Hidden 수주금액 
			C_VATAmt			 = iCurColumnPos(16)
			C_PlantCd			 = iCurColumnPos(17)  '공장코드 
			C_PlantCdPopup		 = iCurColumnPos(18)  '공장팝업 
			C_PlantNm			 = iCurColumnPos(19)  '공장명 
			C_DlvyDt			 = iCurColumnPos(20)  '납기일       
			C_ShipToParty		 = iCurColumnPos(21)  '납품처 
			C_ShipToPartyPopup	 = iCurColumnPos(22)  '납품처팝업 
			C_ShipToPartyNm		 = iCurColumnPos(23)  '납품처명 
			C_HsNo				 = iCurColumnPos(24)  'HS번호 
			C_HsNoPopup			 = iCurColumnPos(25)  'HS번호 Popup
			C_TolMoreRate		 = iCurColumnPos(26)  '과부족허용율(+)
			C_TolLessRate		 = iCurColumnPos(27)  '과부족허용율(-)
			C_VatType			 = iCurColumnPos(28)
			C_VatTypePopup		 = iCurColumnPos(29)
			C_VatTypeNm			 = iCurColumnPos(30)
			C_VatRate			 = iCurColumnPos(31)
			C_VatIncFlag		 = iCurColumnPos(32)
			C_VatIncFlagNm		 = iCurColumnPos(33)
			C_RetType			 = iCurColumnPos(34)
			C_RetTypePopup		 = iCurColumnPos(35)
			C_RetTypeNm			 = iCurColumnPos(36)
			C_LotNo				 = iCurColumnPos(37)
			C_LotSeq			 = iCurColumnPos(38)
			C_PreDnNo			 = iCurColumnPos(39)  '출하번호 for 수주내역참조 
			C_PreDnSeq			 = iCurColumnPos(40)  '출하순번 for 수주내역참조 
			C_DnReqDt			 = iCurColumnPos(41)  '출하요청일자 
			C_BonusQty			 = iCurColumnPos(42)  '할증수량(덤)        
			C_SlCd				 = iCurColumnPos(43)  '창고코드 
			C_SlCdPopup			 = iCurColumnPos(44)  '창고팝업 
			C_SlNm				 = iCurColumnPos(45)  '창고명 
			C_Remark			 = iCurColumnPos(46)  '비고 
			C_SoSts				 = iCurColumnPos(47)  '수주진행상태 
			C_BillQty			 = iCurColumnPos(48)  '매출수량 
			C_BaseQty			 = iCurColumnPos(49)  '재고수량 
			C_BonusBaseQty		 = iCurColumnPos(50)  '덤재고수량 
			C_MaintSeq			 = iCurColumnPos(51)  '관리순번 
			C_OrderSeq			 = iCurColumnPos(52)  '주문서순번 
			C_APSHost			 = iCurColumnPos(53)  'APS Host
			C_APSPort			 = iCurColumnPos(54)  'APS Port
			C_CTPTimes			 = iCurColumnPos(55)  'CTP Check 횟수 
			C_CTPCheckFlag		 = iCurColumnPos(56)  'CTP Check Flag
			C_SoSeq				 = iCurColumnPos(57)  '수주순번 
			C_PreSoNo			 = iCurColumnPos(58)  '수주번호 for 수주내역참조 
			C_PreSoSeq			 = iCurColumnPos(59)  '수주순번 for 수주내역참조 
			C_TrackingFlg		 = iCurColumnPos(60)  
			C_OldNetAmt			 = iCurColumnPos(61)  'Hidden 수주금액(속도개선)
			C_OriginalNetAmt	 = iCurColumnPos(62)  'Hidden 수주금액(속도개선)
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
			ggoSpread.SSSetProtected C_TrackingNoPopup, -1, -1
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
			    
			' 반품여부 
			If frm1.HRetItemFlag.value = "Y" Then
				ggoSpread.SpreadUnLock  C_RetType, -1, -1
				ggoSpread.SSSetRequired C_RetType, -1, -1
				ggoSpread.SpreadUnLock  C_RetTypePopup, -1, -1
			Else
				ggoSpread.SSSetProtected C_RetType, -1, -1
			End If

			' 수출수주/국내수주 여부에 따라 덤수량,HS부호 관리 
			If Trim(.HExportFlag.value) = "Y" Or Trim(.HCiFlag.value) = "Y" Then 
				ggoSpread.SSSetProtected C_BonusQty, -1, -1
				ggoSpread.SpreadUnLock	C_HsNo, -1, -1
				ggoSpread.SSSetRequired  C_HsNo, -1, -1
			Else
				ggoSpread.SSSetProtected C_HsNo, -1, -1
				ggoSpread.SSSetProtected C_HsNoPopup, -1, -1
			End If
			    
			' 출고요청일자가 현재일자보다 적은경우 알림표시 
			For lRow = 1 To .vspdData.MaxRows          
				.vspdData.Row = lRow : .vspdData.Col = C_DnReqDt
				If UniConvDateToYYYYMMDD(.vspdData.Text,parent.gDateFormat,"") < UniConvDateToYYYYMMDD(EndDate,parent.gDateFormat,"") Then 
					Call sprRedComColor(C_DnReqDt,lRow,lRow)
				End If
			Next         
			
			ggoSpread.SpreadUnLock	C_OldNetAmt, -1, -1
			ggoSpread.SpreadUnLock	C_OriginalNetAmt, -1, -1  
			
		Else  ' 확정처리된 경우   
  
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
			   
			' 출고요청일자가 현재일자보다 적은경우 알림표시 
			For lRow = 1 To .vspdData.MaxRows          
				.vspdData.Row = lRow : .vspdData.Col = C_DnReqDt
				If UniConvDateToYYYYMMDD(.vspdData.Text,parent.gDateFormat,"") < UniConvDateToYYYYMMDD(EndDate,parent.gDateFormat,"") Then 
					Call sprRedComColor(C_DnReqDt,lRow,lRow)
				End If				
			Next       
			
			ggoSpread.SSSetProtected C_OldNetAmt, lRow, lRow
			ggoSpread.SSSetProtected C_OriginalNetAmt, lRow, lRow
			
		End If

		.vspdData.ReDraw = True

		If Trim(.RdoConfirm.value) = "N" Then Call SetToolbar("11100000000111")

    End With

End Sub

'========================================================================================================
Sub Form_Load()

	Err.Clear                                                                '☜: Clear err status
	Call LoadInfTB19029                                                      '☜: Load table , B_numeric_format
	
	If GetSetupMod(parent.gSetupMod,"y") = "Y" Then
		txtOpenPrjRef.style.display = ""
    End If                        	 

    Call FormatDoubleSingleField(frm1.txtNetAmt)
    Call LockObjectField(frm1.txtNetAmt,"P")

	Call InitSpreadSheet

	Call InitVariables

	Call SetDefaultVal
	
	Call SetToolbar("11000000000011")								 '⊙: 버튼 툴바 제어 	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call CookiePage(0)
End Sub
'========================================================================================================
'20050218 hjo
Function ChkRulePrice()
	DIM lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	
	  ChkRulePrice=True
		Call CommonQueryRs(" MINOR_CD "," B_CONFIGURATION "," MAJOR_CD = " & FilterVar("S1000", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    If lgF0="" then
    	Call DisplayMsgBox("171214","X","X","X")
    	ChkRulePrice=False
    End if

End Function 	
'========================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear																	  '☜: Clear error status
    
    FncQuery = False															  '☜: Processing is NG    
	
	If Not chkFieldByCell(frm1.txtConSoNo,"A",gPageNo) Then Exit Function
		    '단가 규칙체크 
	if  Not(ChkRulePrice) then
		Exit function
	End If

	ggoSpread.Source = frm1.vspdData 
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")	  '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
        
    'Call ggoSpread.ClearSpreadData
	Call ggoOper.ClearField(Document, "2")								  '☜: Clear Contents  Field

    '------ Developer Coding part (Start ) --------------------------------------------------------------     
    Call InitVariables												 

    If DbQuery = False Then                                    
       Exit Function
    End If
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncQuery = True                                               '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement
        
End Function


'========================================================================================
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
        
    Call ggoOper.ClearField(Document, "A")                                 
    Call ggoOper.LockField(Document, "N")                                        
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
    
    Call ggoOper.ClearField(Document, "A")                              
    
    '------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncDelete = True                                                           '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement  
    
End Function


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
 
    IF ggoSpread.SSDefaultCheck = False Then								  '☜: Check contents area
		Exit Function
    End If
        
    '------ Developer Coding part (Start ) -------------------------------------------------------------- 

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
	
	Call DbSave()  
'------ Developer Coding part (End )   -------------------------------------------------------------- 
	
    If Err.number = 0 Then	
       FncSave = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement
    
End Function


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
	.vspdData.Col = C_SoSeq
		
	  .vspdData.Col = C_TrackingNo :  .vspdData.Text = ""
	  .vspdData.Col = C_SoSupplyQty :  .vspdData.Text = ""
	  .vspdData.Col = C_SoQty   :  .vspdData.Text = ""	  
	  .vspdData.Col = C_SoPriceAutoChk:  .vspdData.Text = "0"
	  .vspdData.Col = C_SoPriceFlag :  .vspdData.Text = "진단가"
	  .vspdData.Col = C_NetAmt  :  .vspdData.Text = 0
	  .vspdData.Col = C_TotalAmt  :  .vspdData.Text = 0
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
	  .vspdData.Col = C_OldNetAmt  :  .vspdData.Text = ""
	  .vspdData.Col = C_OriginalNetAmt  :  .vspdData.Text = ""
	End With
		
	Call SetActiveCell(frm1.vspdData,C_SoQty,frm1.vspdData.ActiveRow,"M","X","X")
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncCopy = True                                                            '☜: Processing is OK
    End If
    
	'-------------------------------
	' V3.0 : Tracking No 수동채번 
	'-------------------------------
	Call ChangeTrackingRetField(frm1.vspdData.ActiveRow)  
	
    Set gActiveElement = document.ActiveElement   
    
End Function


'========================================================================================================
Function FncCancel() 
	Dim iDx
	
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCancel = False                                                             '☜: Processing is NG
	
	If frm1.vspdData.MaxRows < 1 Then Exit Function
    ggoSpread.Source = Frm1.vspdData	
    
    Call CancelSum()
    
    ggoSpread.EditUndo  
	'------ Developer Coding part (Start ) --------------------------------------------------------------
													    
    '------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncCancel = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement 
    
End Function



'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
	Dim IntRetCD
    Dim imRow, i
    Dim iIntIndex
    
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
	
		For i = .vspdData.ActiveRow To .vspdData.ActiveRow + imRow - 1
		
			.vspdData.Row = i
			
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
			
			If Len(.txtHVATIncFlag.value) Then
				.vspdData.Col = C_VatIncFlag
				.vspdData.text = .txtHVATIncFlag.value
				
				iIntIndex = .vspdData.value
								
				.vspdData.Col = C_VatIncFlagNm
				.vspdData.value = iIntIndex - 1		

			End If  
			
			.vspdData.Col= C_VatType

			If Len(.txtHVATType.value) Then
				.vspdData.text = frm1.txtHVATType.value
				Call SetVatType(.vspdData.Row)
			End If
		Next
		
		.vspdData.ReDraw = True   

    End With
    
    
  '------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
Function FncDeleteRow() 
	Dim lDelRows, lDelRow

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
    	
    	For lDelRow = .SelBlockRow to .SelBlockRow2
    		Call deleteSum(lDelRow)
    	Next
    	
    End With
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

    lgBlnFlgChgValue = True 
    'Call TotalSum
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncDeleteRow = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
    
End Function


'========================================================================================================
Function FncPrint() 
	On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncPrint = False                                                              '☜: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()                                                        

    If Err.number = 0 Then	 
       FncPrint = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement
End Function


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
Sub FncSplitColumn()    

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    
End Sub


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
Function DbQuery() 

    Dim strVal
	
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    DbQuery = False                                                               '☜: Processing is NG
	
    Call LayerShowHide(1)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    If lgIntFlgMode = parent.OPMD_UMODE Then    
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001				'☜: 비지니스 처리 ASP의 상태 
		strVal = strVal & "&txtConSoNo=" & Trim(frm1.txtHSoNo.value)        '☜: 조회 조건 데이타 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	Else
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001         '☜: 비지니스 처리 ASP의 상태 %>
		strVal = strVal & "&txtConSoNo=" & Trim(frm1.txtConSoNo.value)     '☜: 조회 조건 데이타 %>
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
Function DbSave() 
    Dim lRow        
	Dim strVal, strDel
	Dim igColSep
	Dim igRowSep
	Dim ii
	
	Dim strCUTotalvalLen
	Dim strDTotalvalLen

	Dim objTEXTAREA

	Dim iTmpCUBuffer
	Dim iTmpCUBufferCount
	Dim iTmpCUBufferMaxCount

	Dim iTmpDBuffer
	Dim iTmpDBufferCount
	Dim iTmpDBufferMaxCount
	

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    DbSave = False                                                               '☜: Processing is NG

    Call DisableToolBar(Parent.TBC_SAVE)                                          '☜: Disable Save Button Of ToolBar
    Call LayerShowHide(1)                                                         '☜: Show Processing Message
		
    frm1.txtMode.value        = Parent.UID_M0002                                  '☜: Delete
'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	
    ggoSpread.Source = frm1.vspdData

    strVal = ""

    igColSep = Parent.gColSep
    igRowSep = Parent.gRowSep
    
    iTmpCUBufferMaxCount = C_MAX_FORM_ARRAY_DATA
    iTmpDBufferMaxCount  = C_MAX_FORM_ARRAY_DATA

    ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
    ReDim iTmpDBuffer (iTmpDBufferMaxCount)

    iTmpCUBufferMaxCount = -1 
    iTmpDBufferMaxCount = -1 

    strCUTotalvalLen = 0
    strDTotalvalLen  = 0

	strVal = ""

	With Frm1
       For lRow = 1 To .vspdData.MaxRows

           .vspdData.Row = lRow
           .vspdData.Col = 0
        
 			Select Case .vspdData.Text
				Case ggoSpread.InsertFlag     
					strVal = "C" & igColSep & lRow & igColSep
				Case ggoSpread.UpdateFlag     
					strVal = "U" & igColSep & lRow & igColSep
				Case ggoSpread.DeleteFlag     
					strVal = "D" & igColSep & lRow & igColSep
			End Select
			
			
           .vspdData.Col = 0
 			Select Case .vspdData.Text
			Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag, ggoSpread.DeleteFlag
						
				.vspdData.Col = C_SoSeq			:	strVal = strVal & Trim(.vspdData.Text)      & igColSep
				.vspdData.Col = C_ItemCd		:	strVal = strVal & Trim(.vspdData.Text)      & igColSep
				.vspdData.Col = C_SoUnit		:	strVal = strVal & Trim(.vspdData.Text)      & igColSep   
				.vspdData.Col = C_TrackingNo    :	strVal = strVal & Trim(.vspdData.Text)      & igColSep   
				        
				.vspdData.Col = C_SoQty      
				If UNIConvNum(Trim(.vspdData.Text), 0) <= 0 Then
							
				   Call DisplayMsgBox("203233", "X", "X", "X")   
				   Call LayerShowHide(0)
				   		   
				   .vspdData.Col = C_SoQty  
				   .vspdData.Action = Parent.SS_ACTION_ACTIVE_CELL
						        					
				   Exit Function
				Else
				    strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0)      & igColSep   
				End If                              
				        
				.vspdData.Col = C_SoPrice		:	strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & igColSep   
				.vspdData.Col = C_SoPriceFlag	
				        
				Select Case Trim(.vspdData.TypeComboBoxCurSel)
					Case 1  '진단가 
							strVal = strVal & "Y" & igColSep    
					Case 0  '가단가 
							strVal = strVal & "N" & igColSep   
					Case Else
							MsgBox "가단가/진단가 값이 없습니다.", vbExclamation, parent.gLogoName 
							frm1.vspdData.Row = lRow
							frm1.vspdData.Action = 0
							'--작업중 표시화면 및 마우스 포인트 원복      
							Call BtnDisabled(False)
							Call LayerShowHide(0)
							Exit Function
				End Select

				.vspdData.Col = C_NetAmt		:	strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0)	& igColSep
				'.vspdData.Col = C_OldNetAmt	:	strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0)	& igColSep   
				'.vspdData.Col = C_OriginalNetAmt:	strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0)	& igColSep      
				.vspdData.Col = C_VatAmt		:	strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0)	& igColSep   
				.vspdData.Col = C_PlantCd		:	strVal = strVal & Trim(.vspdData.Text)					& igColSep   
				.vspdData.Col = C_DlvyDt		:	strVal = strVal & UNIConvDate(Trim(.vspdData.Text))		& igColSep   
				.vspdData.Col = C_ShipToParty	:	strVal = strVal & Trim(.vspdData.Text)					& igColSep   
				.vspdData.Col = C_HsNo			:	strVal = strVal & Trim(.vspdData.Text)					& igColSep
				.vspdData.Col = C_TolMoreRate	:	strVal = strVal & UNICDbl(Trim(.vspdData.Text))			& igColSep   
				.vspdData.Col = C_TolLessRate	:	strVal = strVal & UNICDbl(Trim(.vspdData.Text))			& igColSep   
				.vspdData.Col = C_VatType		:	strVal = strVal & Trim(.vspdData.Text)					& igColSep   
				.vspdData.Col = C_VatRate		:	strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0)   & igColSep   
				.vspdData.Col = C_VatIncFlag	:	strVal = strVal & Trim(.vspdData.Text)					& igColSep   
				.vspdData.Col = C_RetType		:	strVal = strVal & Trim(.vspdData.Text)					& igColSep   
				.vspdData.Col = C_LotNo			:	strVal = strVal & Trim(.vspdData.Text)					& igColSep   
				.vspdData.Col = C_LotSeq		:	strVal = strVal & UNICDbl(Trim(.vspdData.Text))			& igColSep   
				.vspdData.Col = C_PreDnNo		:	strVal = strVal & Trim(.vspdData.Text)					& igColSep   
				.vspdData.Col = C_PreDnSeq		:	strVal = strVal & UNICDbl(Trim(.vspdData.Text))			& igColSep   
				.vspdData.Col = C_BonusQty		:	strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0)   & igColSep   
				.vspdData.Col = C_SlCd			:	strVal = strVal & Trim(.vspdData.Text)					& igColSep   
				.vspdData.Col = C_Remark		:	strVal = strVal & Trim(.vspdData.Text)					& igColSep   
				.vspdData.Col = C_PreSoNo		:	strVal = strVal & Trim(.vspdData.Text)					& igColSep   
				.vspdData.Col = C_PreSoSeq		:	strVal = strVal & UNICDbl(Trim(.vspdData.Text))			& igColSep   
 
				strVal = strVal & 0 & igColSep
				strVal = strVal & ext1_qty & igColSep
				strVal = strVal & ext2_qty & igColSep
				strVal = strVal & ext3_qty & igColSep
				strVal = strVal & ext1_amt & igColSep
				strVal = strVal & ext2_amt & igColSep
				strVal = strVal & ext3_amt & igColSep
				strVal = strVal & ext1_cd  & igColSep
				strVal = strVal & ext2_cd  & igColSep
				strVal = strVal & ext3_cd  & igRowSep
 
				If strCUTotalvalLen + Len(strVal) >  C_MAX_FORM_LIMIT_DATA Then  '한개의 form element에 넣을 한개치가 넘으면 
					Set objTEXTAREA = document.createElement("TEXTAREA")
					objTEXTAREA.name = "txtCUSpread"
					objTEXTAREA.value = Join(iTmpCUBuffer,"")
					divTextArea.appendChild(objTEXTAREA)     

					iTmpCUBufferMaxCount = C_MAX_FORM_ARRAY_DATA
					ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
					iTmpCUBufferCount = -1
					strCUTotalvalLen  = 0

				End If
		        
		        iTmpCUBufferCount = iTmpCUBufferCount + 1
				If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                     '버퍼의 조정 증가치를 넘으면 
					iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + C_MAX_FORM_ARRAY_DATA
					ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
				End If   

				iTmpCUBuffer(iTmpCUBufferCount) =  strVal         
				strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
			End Select 
       Next

	End With

	If iTmpCUBufferCount > -1 Then   ' 나머지 데이터 처리 
       Set objTEXTAREA = document.createElement("TEXTAREA")
       objTEXTAREA.name   = "txtCUSpread"
       objTEXTAREA.value = Join(iTmpCUBuffer,"")
       divTextArea.appendChild(objTEXTAREA)     
    End If   

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    If Err.number = 0 Then	 
       DbSave = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function


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
Sub DbQueryOk()         
	
	On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    lgIntFlgMode = Parent.OPMD_UMODE    
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
  
	frm1.btnAvlStkRef.disabled = False

	If Trim(frm1.txtHPreSONo.value) <> "" And UCase(Trim(frm1.HRetItemFlag.value)) = "Y" Then
		Call SetToolbar("11101011000111")
	ElseIf Trim(frm1.txtHPreSONo.value) = "" And UCase(Trim(frm1.HRetItemFlag.value)) = "Y" Then
		Call SetToolbar("11101111001111")
	ElseIf UCase(Trim(frm1.HRetItemFlag.value)) <> "Y" Then
		Call SetToolbar("11101111001111")
	Else
		Call SetToolbar("11101111001111")
	End If
	
	frm1.vspdData.Focus
	
	Call SetQuerySpreadColor(1)    
	Call TotalSum(frm1.vspdData.ActiveRow)

	lgBlnFlgChgValue = False
    
    Call ChangePlantColor()    
    
	Call ButtonVisible(1)
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call LockHTMLField(frm1.txtPlant, "O")
	
    Set gActiveElement = document.ActiveElement  
    
End Sub


'========================================================================================================
Sub DbSaveOk()              
	On Error Resume Next                                                   '☜: If process fails
    Err.Clear                                                              '☜: Clear error status
	Dim ii
	
    Call InitVariables													   
    
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
    frm1.txtConSoNo.value = frm1.txtHSoNo.value
	frm1.vspdData.MaxRows = 0
	
	For ii = 1 To divTextArea.children.length
        divTextArea.removeChild(divTextArea.children(0))
    Next    
        
    Call MainQuery()
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    
    Set gActiveElement = document.ActiveElement
    
End Sub


'========================================================================================================
Function DbDeleteOk()            
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    Set gActiveElement = document.ActiveElement 
End Function

'========================================================================================================
Function OpenSODtlRef()
 
	Dim arrRet
	Dim iStrSONo
	Dim iCalledAspName
	Dim arrRet2
	
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	
	If Trim(frm1.txtHSoNo.value) = "" Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If
	
	If Trim(frm1.RdoConfirm.value) = "N" Then					
		' 확정처리된 수주는 수정, 삭제, 추가를 할 수 없습니다.
		Call DisplayMsgBox("203022", "x", "x", "x")		
		Exit Function
	End If
	
	Call CommonQueryRs(" A.PROJECT_CODE , A.PROJECT_NM ", " PMS_PROJECT A ", " SO_NO = " & FilterVar(Trim(frm1.txtHSoNo.value), "''", "S") & " ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	

	If Trim(lgF0) <> "" then
		'msgbox "프로젝트 수주참조된건은 참조할수 없습니다.message code 추가"
		Call DisplayMsgBox("YM0049", "x", "x", "x")	
		Exit Function
    End if
    
	If UCase(frm1.HRetItemFlag.value) = "Y" Then
	
		If UCase(frm1.txtHPreSONo.value) = "" Then
			'직접 입력한 반품수주는 수주 내역을 참조할 수 없습니다.
			Call DisplayMsgBox("203156", "x", "x", "x")			
			Exit Function
		End If   	

		' 박정순 수정 (2007-03-16) 거래처 , 화폐단위 Default 처리.
		iStrSONo = frm1.txtHPreSONo.value 
		iStrSONo = iStrSONo & parent.gRowSep & frm1.txtSoldToParty.value
		iStrSONo = iStrSONo & parent.gRowSep & frm1.txtSoldToPartyNm.value
		iStrSONo = iStrSONo & parent.gRowSep & frm1.txtSoType.value
		iStrSONo = iStrSONo & parent.gRowSep & frm1.txtCurrency.value	

		IsOpenPop = True
	
		iCalledAspName = AskPRAspName("s3112ra4")	
		If Trim(iCalledAspName) = "" then
			Call DisplayMsgBox("900040",parent.VB_INFORMATION, "s3112ra4", "x")
			IsOpenPop = False
			Exit Function
		End If
	
		arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, iStrSONo), _
		  "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		IsOpenPop = False

		If arrRet(0, 0) = "" Then
			Exit Function
		Else
			Call SetSODtlRef_RetItem(arrRet)
		End If 
	
	Else
	
		iStrSONo = frm1.txtHPreSONo.value 
		iStrSONo = iStrSONo & parent.gRowSep & frm1.txtSoldToParty.value
		iStrSONo = iStrSONo & parent.gRowSep & frm1.txtSoldToPartyNm.value
		iStrSONo = iStrSONo & parent.gRowSep & frm1.txtSoType.value
		iStrSONo = iStrSONo & parent.gRowSep & frm1.txtCurrency.value		
		
		IsOpenPop = True
	
		iCalledAspName = AskPRAspName("s3112ra41")	
		
		If Trim(iCalledAspName) = "" then
			Call DisplayMsgBox("900040",parent.VB_INFORMATION, "s3112ra41", "x")
			IsOpenPop = False
			Exit Function
		End If		
	
		arrRet2 = window.showModalDialog(iCalledAspName, Array(window.parent, iStrSONo), _
		  "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		IsOpenPop = False				
	
		arrRet = arrRet2(0)	
		
		If arrRet(0,0) = "" Then			
			Exit Function
		Else		
			Call SetSODtlRef(arrRet2)
		End If
		
	End If

End Function

'========================================================================================================
Function OpenPrjRef()
 
	Dim arrRet
	Dim iStrSONo
	Dim iCalledAspName
	Dim arrParam(2)
	
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	
	If Trim(frm1.txtHSoNo.value) = "" Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If
	
	If Trim(frm1.RdoConfirm.value) = "N" Then					
		' 확정처리된 수주는 수정, 삭제, 추가를 할 수 없습니다.
		Call DisplayMsgBox("203022", "x", "x", "x")		
		Exit Function
	End If
		
	Call CommonQueryRs(" A.PROJECT_CODE , A.PROJECT_NM ", " PMS_PROJECT A ", " SO_NO = " & FilterVar(Trim(frm1.txtHSoNo.value), "''", "S") & " ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	

	If trim(lgF0) = "" then
		'msgbox "프로젝트 수주참조 아니한건은 참조할수 없습니다.message code 추가"
		Call DisplayMsgBox("YM0048", "x", "x", "x")	
		Exit Function
    End if
		
		arrParam(0) = frm1.txtHSoNo.value 
		arrParam(1) = Replace(trim(lgF0) ,Chr(11) ,"")
		arrParam(2) = Replace(trim(lgF1) ,Chr(11) ,"")
		
		IsOpenPop = True
	
		iCalledAspName = AskPRAspName("s3112ra51")	
		
		If Trim(iCalledAspName) = "" then
			Call DisplayMsgBox("900040",parent.VB_INFORMATION, "s3112ra51", "x")
			IsOpenPop = False
			Exit Function
		End If		
	
		arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		  "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		IsOpenPop = False				
		
		If arrRet(0,0) = "" Then			
			Exit Function
		Else		
			Call SetPrjRef(arrRet)
		End If
	
End Function


'========================================================================================================
Function OpenAvalStockRef()

	Dim arrRet
	Dim arrParam(4)
	Dim iCalledAspName
	
	If UCase(Trim(frm1.txtConSoNo.value)) = "" Then
		Call DisplayMsgBox("900002", "x", "x", "x") 
		Exit Function
	End If

	ggoSpread.Source = frm1.vspdData
  
	If frm1.vspdData.Maxrows < 1 Then
		Call DisplayMsgBox("209001", "x", "x", "x")
		'행을 먼저 선택하십시오.
		Exit Function
	End If 
  
	frm1.vspdData.Col = C_ItemCd
	frm1.vspdData.Row = frm1.vspdData.ActiveRow

	arrParam(0) = frm1.vspdData.Text

	If arrParam(0) = "" Then
		Call DisplayMsgBox("202250", "x", "x", "x")				'⊙: "Will you destory previous data" 
		'%1 품목을 먼저 선택하세요.
		Exit Function
	End If 
  
	frm1.vspdData.Col = C_ItemName   
 
	arrParam(1) = frm1.vspdData.Text

	frm1.vspdData.Col = C_PlantCd
 
	arrParam(2) = frm1.vspdData.Text

	frm1.vspdData.Col = C_PlantNm
  
	arrParam(3) = frm1.vspdData.Text

	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("s1912ra1")	
	If Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s1912ra1", "x")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
	  "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
  
	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetSODtlRef_RetItem(arrRet)
	End If 
	
 End Function

'===========add by HJO , 20050118=============================================================================================
Function OpenBOMRef()

	Dim arrRet, arrRet2
	Dim arrParam
	Dim iCalledAspName
	
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	
	If UCase(Trim(frm1.txtSoldToParty.value)) = "" Then
		Call DisplayMsgBox("900002", "x", "x", "x")		 '⊙: "Will you destory previous data"
		frm1.txtConSoNo.Focus
		Exit Function
	End If
	If Trim(frm1.RdoConfirm.value) = "N" Then					
		' 확정처리된 수주는 수정, 삭제, 추가를 할 수 없습니다.
		Call DisplayMsgBox("203022", "x", "x", "x")		
		Exit Function
	End If
	
	Call CommonQueryRs(" A.PROJECT_CODE , A.PROJECT_NM ", " PMS_PROJECT A ", " SO_NO = " & FilterVar(Trim(frm1.txtHSoNo.value), "''", "S") & " ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	

	If Trim(lgF0) <> "" then
		'msgbox "프로젝트 수주참조된건은 참조할수 없습니다.message code 추가"
		Call DisplayMsgBox("YM0049", "x", "x", "x")	
		Exit Function
    End if
    
              arrParam = arrParam & trim(frm1.txtPlant.value)
              arrParam = arrParam & parent.gRowSep & trim(frm1.txtPlantNm.value)
	arrParam = arrParam & parent.gRowSep & trim(frm1.txtSoldToParty.value)
	arrParam = arrParam & parent.gRowSep &Trim(frm1.txtCurrency.value)

	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("s3112ra20")	
	If Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3112ra20", "x")
		IsOpenPop = False
		Exit Function
	End If

	arrRet2 = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam),  _
		 "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	arrRet=arrRet2(0)

	If arrRet (0,0)= "" Then
		Exit Function
	else 
		call SetBomRef(arrRet2)
	End If  
		
End Function
'========================================================================================================
Function SetBomRef(arrRet2)


	Dim intRtnCnt, strData

	Dim TempRow, I, j
	Dim blnEqualFlg
	Dim intLoopCnt
	Dim intCnt
	Dim strSoNo,strSoSeqNo
	Dim arrRet

	Dim strPlant, strPlantNm, strPoNo, strPoNoSeq
	Dim strSoPrice, strVatRate
	
	arrRet = arrRet2(0)
	strPlant = arrRet2(1)
	strPlantNm = arrRet2(2)
	strPoNo = arrRet2(3)
	strPoNoSeq = arrRet2(4)

	If  Trim(frm1.txtPlant.value) =""  AND Trim(frm1.txtPlantNm.value) =""  Then
		frm1.txtPlant.value = strPlant
		frm1.txtPlantNm.value = strPlantNm
	ElseIf Trim(frm1.txtPlantNm.value) ="" then
		frm1.txtPlantNm.value = strPlantNm
	End If
	
	With frm1
		.vspdData.focus
		ggoSpread.Source = .vspdData
		.vspdData.ReDraw = False 

		TempRow = .vspdData.MaxRows         
		intLoopCnt = Ubound(arrRet, 1)
		
		For intCnt = 1 to intLoopCnt 
			blnEqualFlg = False			

			If blnEqualFlg = False Then
				.vspdData.MaxRows = .vspdData.MaxRows + 1
				.vspdData.Row = CLng(TempRow) + CLng(intCnt)

				.vspdData.Col = 0
				.vspdData.Text = ggoSpread.InsertFlag

				.vspdData.Col = C_ItemCd
				.vspdData.text = arrRet(intCnt - 1, 0)		
				 

				.vspdData.Col = C_ItemName
				.vspdData.text = arrRet(intCnt - 1,1)				
				.vspdData.Col = C_ItemSpec
				.vspdData.text = arrRet(intCnt - 1,2)
				.vspdData.Col = C_SoQty
				.vspdData.text = arrRet(intCnt - 1, 3)
								
				.vspdData.Col = C_SoUnit
				.vspdData.text = arrRet(intCnt - 1, 4)	

				.vspdData.Col = C_VATType
				'VAT의 유형이 없을 경우 수주정보의 VAT유형을 가지고온다.
				If arrRet(intCnt-1,9) <>"" then		
					.vspdData.text = arrRet(intCnt - 1, 9)
				Else
					.vspdData.text = frm1.txtHVATType.value					
				End If
				Call SetVatType(.vspdData.Row)
				.vspdData.Col = C_VatRate
				strVatRate = .vspdData.Text	
	
				.vspdData.Col = C_DlvyDt
				.vspdData.text = Trim(frm1.txtSoDt.value)				
					
				.vspdData.Col = C_ShipToParty
				.vspdData.text = .txtShipToParty.value
				
				.vspdData.Col = C_SoPriceFlag
				Select Case .HPriceFlag.value
					Case "Y"
					 .vspdData.Text = "진단가"
					Case "N"
					 .vspdData.Text = "가단가"
				End Select
				
				.vspdData.Col = C_HsNo
				'수출수주여부확인	
				If frm1.HExportFlag.value ="Y" Then				
					.vspdData.text = arrRet(intCnt - 1, 7)		
				Else
					.vspdData.text = ""
				End If

				'단가				
				Call GetItemPrice(CLng(TempRow) + CLng(intCnt))
				.vspdData.Row = CLng(TempRow) + CLng(intCnt)
				.vspdData.Col = C_SoPrice : strSoPrice = .vspdData.Text 
				
				.vspdData.Col = C_TolMoreRate
				.vspdData.text = 0
				.vspdData.Col = C_TolLessRate
				.vspdData.text = 0
				.vspdData.Col = C_PlantCd
				.vspdData.text = strPlant

				.vspdData.Col = C_VatIncFlag
				.vspdData.text = Trim(.txtVatIncFlag.value)
				.vspdData.Col = C_VatIncFlagNm				
				If Trim(.txtVatIncFlag.value) = "1" Then
					 .vspdData.text = "별도"
				ElseIf  Trim(.txtVatIncFlag.value)  = "2" Then
					 .vspdData.text = "포함"
				End If      
				'발주번호 
				ext3_cd = arrRet2(3)
				'발주순번 
				if  arrRet2(4) <>"" then
					ext3_qty = arrRet2(4)
				'else
				'	ext3_qty =0	
				end if

				SetSpreadColor CLng(TempRow) + CLng(intCnt), CLng(TempRow) + CLng(intCnt)
				
'				SetVatType (CLng(TempRow) + CLng(intCnt))

				.vspdData.Row = CLng(TempRow) + CLng(intCnt)

				'금액 : 수량 *단가	
				.vspdData.Col=C_TotalAmt			
				.vspdData.text = arrRet(intCnt - 1, 3) * strSoPrice
				'VAT금액 : 금액 * VAT Rate
				.vspdData.Col = C_VATAmt
				.vspdData.Text = arrRet(intCnt - 1, 3) * strSoPrice * (strVatRate/100)
								
				'참조한 품목에대한 TrackingNo 수동채번 결정	
				Call SetTrackingNoByItem(CLng(TempRow) + CLng(intCnt))
				
				lgBlnFlgChgValue = True
			End If
			Call SetQuerySpreadColor( CLng(TempRow) + CLng(intCnt))

		Next		
		

		.vspdData.ReDraw = True
		
		
	End With
end function

'========================================================================================================
Function OpenStockDtlRef()
	Dim arrRet
	Dim arrParam(5)
	Dim iCalledAspName
	
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
		Call DisplayMsgBox("202250", "x", "x", "x")  '⊙: "Will you destory previous data" %>
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

	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("s1912ra2")	
	If Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s1912ra2", "x")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam), _
	  "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
  
	IsOpenPop = False

	If arrRet(0, 0) = "" Then
		Exit Function
	End If 
  
End Function


'===========================================================================
Function OpenItem(ByVal strCode)

	Dim arrParam(1)
	Dim strRet
	Dim iCalledAspName
	
	arrParam(0) = strCode
	frm1.vspdData.Col = C_PlantCd
	arrParam(1) = frm1.vspdData.text 

	If IsOpenPop = True Then Exit Function
	  
	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("s3112pa2")
	
	If Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3112pa2", "x")
		IsOpenPop = False
		exit Function
	End if
	 
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
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
Function OpenConSoDtl()

	Dim iCalledAspName
	Dim strRet

	If IsOpenPop = True Then Exit Function
	  
	IsOpenPop = True
	 
	iCalledAspName = AskPRAspName("s3111pa1")
			
	If Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3111pa1", "x")
		IsOpenPop = False
		exit Function
	End if

	strRet = window.showModalDialog(iCalledAspName, Array(window.parent, "SO_REG"), _
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
	
	'2005.10.27 smj 수동채번인 경우 팝업추가 
	Case 2 'Tracking no 팝업 
		arrParam(1) = "s_so_tracking a, b_item_by_plant b, b_item c"    
		arrParam(2) = strCode       
		arrParam(4) = "a.item_cd = b.item_cd and a.sl_cd = b.plant_cd and b.item_cd = c.item_cd"      		 
		arrParam(5) = "Tracking No"   
		    
		arrField(0) = "a.tracking_no"       
		arrField(1) = "a.item_cd"   
		arrField(2) = "c.item_nm"   
		arrField(3) = "c.spec"   	
		  
		arrHeader(0) = "Tracking No"		
		arrHeader(1) = "품목"
		arrHeader(2) = "품목명"
		arrHeader(3) = "Spec"
		
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
	Case 0, 2 '품목 %>
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
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	If frm1.txtPlant.readOnly = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장"   
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
Function SetSODtlRef_RetItem(arrRet)
	Dim intRtnCnt, strData
	Dim TempRow, I, j
	Dim blnEqualFlg
	Dim intLoopCnt
	Dim intCnt
	Dim iStrDnNo,iStrDnSeqNo
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

				iStrDnNo=""
				iStrDnSeqNo=""

					For j = 1 To TempRow

						 '출하번호 %>
						.vspdData.Row = j
						.vspdData.Col = C_PreDnNo
						iStrDnNo = .vspdData.text

						If iStrDnNo = arrRet(intCnt - 1, 26) Then

							 '출하순번 %>
							.vspdData.Row = j
							.vspdData.Col = C_PreDnSeq
							iStrDnSeqNo = .vspdData.text

							If iStrDnSeqNo = arrRet(intCnt - 1, 27) Then
								blnEqualFlg = True
								strSOJungBokMsg = strSOJungBokMsg & Chr(13) & iStrDnNo & "-" & iStrDnSeqNo
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
				.vspdData.text = arrRet(intCnt - 1, 10)
				.vspdData.Col = C_NetAmt
				.vspdData.text = arrRet(intCnt - 1, 12)
				.vspdData.Col = C_VATType
				.vspdData.text = arrRet(intCnt - 1, 14)
				.vspdData.Col = C_TrackingNo
				.vspdData.text = arrRet(intCnt - 1, 16)    
				.vspdData.Col = C_DlvyDt
				.vspdData.text = Trim(frm1.txtSoDt.value)
				.vspdData.Col = C_ShipToParty
				.vspdData.text = arrRet(intCnt - 1, 18)
				.vspdData.Col = C_SoPriceFlag

				If Trim(arrRet(intCnt - 1, 19)) = "Y" Then
					.vspdData.text = "진단가"
				ElseIf  Trim(arrRet(intCnt - 1, 19)) = "N" Then
					.vspdData.text = "가단가"
				End If      

				.vspdData.Col = C_HsNo
				.vspdData.text = arrRet(intCnt - 1, 20)
				.vspdData.Col = C_TolMoreRate
				.vspdData.text = arrRet(intCnt - 1, 21)
				.vspdData.Col = C_TolLessRate
				.vspdData.text = arrRet(intCnt - 1, 22)
				.vspdData.Col = C_PlantCd
				.vspdData.text = arrRet(intCnt - 1, 23)
				.vspdData.Col = C_SlCd
				.vspdData.text = arrRet(intCnt - 1, 24)
				.vspdData.Col = C_Remark
				.vspdData.text = arrRet(intCnt - 1, 25)
				.vspdData.Col = C_PreDnNo
				.vspdData.text = arrRet(intCnt - 1, 26)
				.vspdData.Col = C_PreDnSeq
				.vspdData.text = arrRet(intCnt - 1, 27)
				.vspdData.Col = C_LotNo
				.vspdData.text = arrRet(intCnt - 1, 28)
				.vspdData.Col = C_LotSeq
				.vspdData.text = arrRet(intCnt - 1, 29)
				.vspdData.Col = C_VatIncFlag
				.vspdData.text = arrRet(intCnt - 1, 15)
				.vspdData.Col = C_VatIncFlagNm

				If Trim(arrRet(intCnt - 1, 15)) = "1" Then
					 .vspdData.text = "별도"
				ElseIf  Trim(arrRet(intCnt - 1, 15)) = "2" Then
					 .vspdData.text = "포함"
				End If      

				SetSpreadColor CLng(TempRow) + CLng(intCnt),CLng(TempRow) + CLng(intCnt)
			
				SetVatType (TempRow + intCnt)				
				
				Call QtyPriceChange(TempRow + intCnt)
				lgBlnFlgChgValue = True
			End If
			
			Call JungBokMsg(strSOJungBokMsg,"출하번호" & "-" & "출하순번")
			
		Next
			
		.vspdData.ReDraw = True	
		
	End With
	
End Function



'========================================================================================================
Function SetSODtlRef(arrRet2)
	Dim intRtnCnt, strData
	Dim TempRow, I, j
	Dim blnEqualFlg
	Dim intLoopCnt
	Dim intCnt
	Dim strSoNo,strSoSeqNo
	Dim arrRet
	Dim iSoldToParty
	
	arrRet = arrRet2(0)
	iSoldToParty = arrRet2(1)
	
	With frm1
		.vspdData.focus
		ggoSpread.Source = .vspdData
		.vspdData.ReDraw = False 

		TempRow = .vspdData.MaxRows         
		intLoopCnt = Ubound(arrRet, 1)
		
		For intCnt = 1 to intLoopCnt + 1
			blnEqualFlg = False			

			If blnEqualFlg = False Then
				.vspdData.MaxRows = .vspdData.MaxRows + 1
				.vspdData.Row = CLng(TempRow) + CLng(intCnt)
				.vspdData.Col = 0
				.vspdData.Text = ggoSpread.InsertFlag

				.vspdData.Col = C_ItemCd
				.vspdData.text = arrRet(intCnt - 1, 2)
				.vspdData.Col = C_ItemName
				.vspdData.text = arrRet(intCnt - 1, 3)				
				.vspdData.Col = C_ItemSpec
				.vspdData.text = arrRet(intCnt - 1, 22)
				.vspdData.Col = C_SoUnit
				.vspdData.text = arrRet(intCnt - 1, 4)										
				.vspdData.Col = C_VATType
				.vspdData.text = arrRet(intCnt - 1, 10)
				.vspdData.Col = C_VatTypeNm
				.vspdData.text = arrRet(intCnt - 1, 11)
				.vspdData.Col = C_VatRate
				.vspdData.text = arrRet(intCnt - 1, 12)
				.vspdData.Col = C_DlvyDt
				.vspdData.text = Trim(frm1.txtSoDt.value)				
					
				.vspdData.Col = C_ShipToParty
				
				If Trim(iSoldToParty) <> Trim(frm1.txtSoldToParty.value) Then
					.vspdData.text = .txtShipToParty.value
				Else
					.vspdData.text = arrRet(intCnt - 1, 14)
				End If
				
				.vspdData.Col = C_SoPriceFlag
				.vspdData.text = "진단가"	
				
				.vspdData.Col = C_HsNo
				.vspdData.text = arrRet(intCnt - 1, 23)						
				.vspdData.Col = C_TolMoreRate
				.vspdData.text = 0
				.vspdData.Col = C_TolLessRate
				.vspdData.text = 0
				.vspdData.Col = C_PlantCd
				.vspdData.text = arrRet(intCnt - 1, 16)
				.vspdData.Col = C_SlCd
				.vspdData.text = arrRet(intCnt - 1, 18)
				.vspdData.Col = C_VatIncFlag
				.vspdData.text = arrRet(intCnt - 1, 7)
				.vspdData.Col = C_VatIncFlagNm
				
				If Trim(arrRet(intCnt - 1, 7)) = "1" Then
					 .vspdData.text = "별도"
				ElseIf  Trim(arrRet(intCnt - 1, 7)) = "2" Then
					 .vspdData.text = "포함"
				End If      

				SetSpreadColor CLng(TempRow) + CLng(intCnt), CLng(TempRow) + CLng(intCnt)
				
				SetVatType (CLng(TempRow) + CLng(intCnt))
				
				If Not GetItemPrice(CLng(TempRow) + CLng(intCnt)) Then									
					.vspdData.Col = C_SoPrice
					.vspdData.text = arrRet(intCnt - 1, 8)
				End If				
								
				'참조한 품목에대한 TrackingNo 수동채번 결정	
				Call SetTrackingNoByItem(CLng(TempRow) + CLng(intCnt))
				
				lgBlnFlgChgValue = True
			End If
		Next		
		
		.vspdData.ReDraw = True
		
		
	End With
	
End Function


'========================================================================================================
Function SetPrjRef(arrRet)
	Dim intRtnCnt, strData
	Dim TempRow, I, j
	Dim blnEqualFlg
	Dim intLoopCnt
	Dim intCnt
	Dim strSoNo,strSoSeqNo
	Dim iSoldToParty
	

	
	With frm1
		.vspdData.focus
		ggoSpread.Source = .vspdData
		.vspdData.ReDraw = False 

		TempRow = .vspdData.MaxRows         
		intLoopCnt = Ubound(arrRet, 1)
		
		For intCnt = 1 to intLoopCnt + 1
			blnEqualFlg = False			

			If blnEqualFlg = False Then
				.vspdData.MaxRows = .vspdData.MaxRows + 1
				.vspdData.Row = CLng(TempRow) + CLng(intCnt)
				.vspdData.Col = 0
				.vspdData.Text = ggoSpread.InsertFlag

				.vspdData.Col = C_ItemCd
				.vspdData.text = arrRet(intCnt - 1, 0)
				.vspdData.Col = C_ItemName
				.vspdData.text = arrRet(intCnt - 1, 1)				
				.vspdData.Col = C_ItemSpec
				.vspdData.text = arrRet(intCnt - 1, 2)
				.vspdData.Col = C_SoUnit
				.vspdData.text = arrRet(intCnt - 1, 3)
				.vspdData.Col = C_TrackingNo
				.vspdData.text = arrRet(intCnt - 1, 4)		
				.vspdData.Col = C_SoQty
				.vspdData.text = arrRet(intCnt - 1, 5)	
				.vspdData.Col = C_VATType
				.vspdData.text = arrRet(intCnt - 1, 16)
				.vspdData.Col = C_VatTypeNm
				.vspdData.text = arrRet(intCnt - 1, 17)
				.vspdData.Col = C_VatRate
				.vspdData.text = arrRet(intCnt - 1, 18)
				.vspdData.Col = C_DlvyDt
				.vspdData.text = arrRet(intCnt - 1, 13)
				'.vspdData.text = Trim(frm1.txtSoDt.value)				
				
				.vspdData.Col = C_SoPrice
				.vspdData.text = arrRet(intCnt - 1, 6)
				
				.vspdData.Col = C_TotalAmt
				.vspdData.text = arrRet(intCnt - 1, 9)
				.vspdData.Col = C_NetAmt
				.vspdData.text = arrRet(intCnt - 1, 9)
				.vspdData.Col = C_VatAmt
				.vspdData.text = arrRet(intCnt - 1, 10)
				
					
				.vspdData.Col = C_ShipToParty
				
				If Trim(iSoldToParty) <> Trim(frm1.txtSoldToParty.value) Then
					.vspdData.text = .txtShipToParty.value
				Else
					.vspdData.text = arrRet(intCnt - 1, 13)
				End If
				
				.vspdData.Col = C_SoPriceFlag
				.vspdData.text = "진단가"	
				
				.vspdData.Col = C_TolMoreRate
				.vspdData.text = 0
				.vspdData.Col = C_TolLessRate
				.vspdData.text = 0
				.vspdData.Col = C_PlantCd
				.vspdData.text = arrRet(intCnt - 1, 11)
				
				.vspdData.Col = C_VatIncFlag
				.vspdData.text = arrRet(intCnt - 1, 19)
				.vspdData.Col = C_VatIncFlagNm
				
				If Trim(arrRet(intCnt - 1, 19)) = "1" Then
					 .vspdData.text = "별도"
				ElseIf  Trim(arrRet(intCnt - 1, 19)) = "2" Then
					 .vspdData.text = "포함"
				End If      

				SetSpreadColor CLng(TempRow) + CLng(intCnt), CLng(TempRow) + CLng(intCnt)
				
				
				'SetVatType (CLng(TempRow) + CLng(intCnt))
				
				'If Not GetItemPrice(CLng(TempRow) + CLng(intCnt)) Then									
				'	.vspdData.Col = C_SoPrice
				'	.vspdData.text = arrRet(intCnt - 1, 8)
				'End If				
								
				'참조한 품목에대한 TrackingNo 수동채번 결정	
				'Call SetTrackingNoByItem(CLng(TempRow) + CLng(intCnt))
				
				lgBlnFlgChgValue = True
			End If
		Next		
		
		.vspdData.ReDraw = True
		
	End With
	
End Function


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
		 Call vspdData_Change(C_ItemCd, .vspdData.Row)  
		 
		Case 1 '단위 
		 .vspdData.Col = C_SoUnit
		 .vspdData.Text = arrRet(0)
		 Call vspdData_Change(C_SoUnit, .vspdData.Row)  
		 
		Case 2 'tracking no
		 .vspdData.Col = C_TrackingNo
		 .vspdData.Text = arrRet(0)
		  Call vspdData_Change(C_TrackingNo, .vspdData.Row)  		
	 
		Case 3 '납품처 
		 .vspdData.Col = C_ShipToParty
		 .vspdData.Text = arrRet(0)
		 Call vspdData_Change(C_ShipToParty, .vspdData.Row) 
		 
		Case 4 'HS번호 
		 .vspdData.Col = C_HsNo
		 .vspdData.Text = arrRet(0)
		 Call vspdData_Change(C_HsNo, .vspdData.Row)   
		 
		Case 5 '공장 
		 .vspdData.Col = C_PlantCd
		 .vspdData.Text = arrRet(0)
		 Call vspdData_Change(C_PlantCd, .vspdData.Row)  
		 
		Case 6 '창고 
		 .vspdData.Col = C_SlCd
		 .vspdData.Text = arrRet(0)
		 Call vspdData_Change(C_SlCd, .vspdData.Row)   
	 
		Case 7 'VAT유형 
		 .vspdData.Col = C_VatType
		 .vspdData.Text = arrRet(0)
		 .vspdData.Col = C_VatTypeNm
		 .vspdData.Text = arrRet(1)
		 .vspdData.Col = C_VatRate
		 .vspdData.text = UNIFormatNumber(UNICDbl(arrRet(2)), ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
		 Call vspdData_Change(C_VatType, .vspdData.Row)   
	 
		Case 8 '반품유형 
		 .vspdData.Col = C_RetType
		 .vspdData.Text = arrRet(0)
		 .vspdData.Col = C_RetTypeNm
		 .vspdData.Text = arrRet(1)
		 Call vspdData_Change(C_RetType, .vspdData.Row)   
	 
		Case 9 'VAT 포함구분 
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
Function SetPlant(Byval arrRet)
	With frm1
		.txtPlant.value = arrRet(0) 
		.txtPlantNm.value = arrRet(1)   
	End With
End Function


'========================================================================================================
Function SetDefaultPlant()
	With frm1
		.txtPlant.value = parent.gPlant 
		.txtPlantNm.value = parent.gPlantNm   
	End With
End Function


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
Sub SetVatType(pvRow)

	Dim VatType, VatTypeNm, VatRate
	frm1.vspdData.Row = pvRow
	frm1.vspdData.Col = C_VatType : 
	VatType = frm1.vspdData.text
 
	Call InitCollectType

	Call GetCollectTypeRef(VatType, VatTypeNm, VatRate)
	
	frm1.vspdData.Col = C_VatTypeNm :  frm1.vspdData.Text = VatTypeNm
	frm1.vspdData.Col = C_VatRate :  frm1.vspdData.Text = UNIFormatNumber(VatRate, ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
 
	Call QtyPriceChange(pvRow)
End Sub


'========================================================================================================
Function QtyPriceChange(iRow)

	Dim SoQty, SoPrice, NetAmt, VatAmt, CalSOAmt, VatRate, TotalAmt, OldNetAmt, OriginalNetAmt

	frm1.vspdData.Row = iRow
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

	frm1.vspdData.Col = C_OldNetAmt
	If Trim(frm1.vspdData.Text) = "" Or isnull(frm1.vspdData.Text) Then
		OldNetAmt = NetAmt
	Else
		OldNetAmt = UNIFormatNumberByCurrecny(UNICDbl(frm1.vspdData.Text), frm1.txtCurrency.value,parent.ggAmtOfMoneyNo)
	End If

	frm1.vspdData.Col = C_OriginalNetAmt
	If Trim(frm1.vspdData.Text) = "0" Or isnull(frm1.vspdData.Text) Then
		OriginalNetAmt = NetAmt
	Else
		OriginalNetAmt = UNIFormatNumberByCurrecny(UNICDbl(frm1.vspdData.Text), frm1.txtCurrency.value,parent.ggAmtOfMoneyNo)
	End If

	With frm1
	.vspdData.Col = C_NetAmt
	.vspdData.Text = NetAmt
	.vspdData.Col = C_VatAmt
	.vspdData.Text = VatAmt
	.vspdData.Col = C_TotalAmt
	.vspdData.Text = TotalAmt
	
	.vspdData.Col = C_OldNetAmt
	.vspdData.Text = OldNetAmt
	.vspdData.Col = C_OriginalNetAmt
	.vspdData.Text = OriginalNetAmt
	End with

	frm1.vspdData.Col = 0
 	Select Case frm1.vspdData.Text
		Case ggoSpread.DeleteFlag     
			Exit Function			
		Case Else
			Call TotalSum(iRow)
	End Select
		

End Function

'========================================================================================================
Function TotalAmtChange(iRow)

	Dim NetAmt, VatAmt, CalSOAmt, VatRate, TotalAmt

	frm1.vspdData.Row = iRow	
	frm1.vspdData.Col = C_VatRate
	If Trim(frm1.vspdData.Text) = "" Or isnull(frm1.vspdData.Text) Then
		VatRate = 0
	Else
		VatRate = UNIFormatNumber(UNICDbl(frm1.vspdData.text), ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
	End If
	
	frm1.vspdData.Row = iRow
	frm1.vspdData.Col = C_TotalAmt
	If Trim(frm1.vspdData.Text) = "" Or isnull(frm1.vspdData.Text) Then
		CalSOAmt = 0
	Else
		CalSOAmt = UNIFormatNumber(UNICDbl(frm1.vspdData.text), ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
	End If
	
	frm1.vspdData.Row = iRow
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

	frm1.vspdData.Col = 0
 	Select Case frm1.vspdData.Text
		Case ggoSpread.DeleteFlag     
			Exit Function			
		Case Else
			Call TotalSum(iRow)
	End Select
	
End Function


'========================================================================================================
Function TotalSum(iRow)
	Dim SumTotal, NetAmt, OldNetAmt, OriginalNetAmt

	With frm1
		.vspdData.Row = iRow
		.vspdData.col = C_NetAmt
		NetAmt = UNIFormatNumberByCurrecny(UNICDbl(frm1.vspdData.Text), frm1.txtCurrency.value,parent.ggAmtOfMoneyNo)
		
		.vspdData.Col = C_OldNetAmt
		OldNetAmt = UNIFormatNumberByCurrecny(UNICDbl(frm1.vspdData.Text), frm1.txtCurrency.value,parent.ggAmtOfMoneyNo)
		
		SumTotal = frm1.txtNetAmt.text
		SumTotal = UNIFormatNumberByCurrecny(UNICDbl(SumTotal) - UNICDbl(OldNetAmt) + UNICDbl(NetAmt),frm1.txtCurrency.value,parent.ggAmtOfMoneyNo)
		
		frm1.txtNetAmt.text = sumTotal
	
		.vspdData.col = C_OldNetAmt
		.vspdData.text = NetAmt
 	End With

End Function


'========================================================================================================
Function CancelSum()

	Dim SumTotal, NetAmt, OriginalNetAmt
	With frm1
		.vspdData.Row = .vspdData.ActiveRow
		.vspdData.col = C_NetAmt
		NetAmt = UNIFormatNumberByCurrecny(UNICDbl(frm1.vspdData.Text), frm1.txtCurrency.value,parent.ggAmtOfMoneyNo)
		
		.vspdData.Col = C_OriginalNetAmt
		OriginalNetAmt = UNIFormatNumberByCurrecny(UNICDbl(frm1.vspdData.Text), frm1.txtCurrency.value,parent.ggAmtOfMoneyNo)
		
		SumTotal = frm1.txtNetAmt.text
		
        .vspdData.Col = 0
 		Select Case .vspdData.Text
			Case ggoSpread.InsertFlag 
				SumTotal = UNIFormatNumberByCurrecny(UNICDbl(SumTotal) - UNICDbl(NetAmt),frm1.txtCurrency.value,parent.ggAmtOfMoneyNo)
			Case ggoSpread.UpdateFlag     
				SumTotal = UNIFormatNumberByCurrecny(UNICDbl(SumTotal) - UNICDbl(NetAmt) + UNICDbl(OriginalNetAmt),frm1.txtCurrency.value,parent.ggAmtOfMoneyNo)
			Case ggoSpread.DeleteFlag     
				SumTotal = UNIFormatNumberByCurrecny(UNICDbl(SumTotal) + UNICDbl(OriginalNetAmt),frm1.txtCurrency.value,parent.ggAmtOfMoneyNo)
		End Select
		
		frm1.txtNetAmt.text = sumTotal
	End With

End Function


'========================================================================================================
Function DeleteSum(iRow)

	Dim SumTotal, NetAmt, OriginalNetAmt
	With frm1
		.vspdData.Row = iRow 
		.vspdData.col = C_NetAmt
		NetAmt = UNIFormatNumberByCurrecny(UNICDbl(frm1.vspdData.Text), frm1.txtCurrency.value,parent.ggAmtOfMoneyNo)
		
		SumTotal = frm1.txtNetAmt.text		
 
		SumTotal = UNIFormatNumberByCurrecny(UNICDbl(SumTotal) - UNICDbl(NetAmt),frm1.txtCurrency.value,parent.ggAmtOfMoneyNo)
		
		frm1.txtNetAmt.text = sumTotal
	End With

End Function


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


'===================================================================================================
Function ChangeTrackingRetField(ByVal IRow)
	Dim strTrackingFlag	
		
	ggoSpread.Source = frm1.vspdData
	frm1.vspdData.Col = C_TrackingFlg
	frm1.vspdData.Row = IRow
	strTrackingFlag = frm1.vspdData.text
	
	
	If frm1.txtHTrackingNORule.value = "M" and strTrackingFlag = "Y" Then
 		ggoSpread.SpreadUnLock C_TrackingNo, IRow, C_TrackingNo, IRow
	    ggoSpread.SSSetRequired C_TrackingNo, IRow, IRow
	    ggoSpread.SpreadUnLock C_TrackingNoPopup, IRow, C_TrackingNoPopup, IRow	    
	Else	
		ggoSpread.SSSetProtected C_TrackingNo, IRow, IRow
		ggoSpread.SSSetProtected C_TrackingNoPopup, IRow, IRow
	End If

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

	 '-- 멀티일때 -- %>
	 '변경이 있을떄 저장 여부 먼저 체크후, YES이면 작업진행여부 체크 안한다 %>
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")                
		If IntRetCD = vbNo Then Exit Function
	End If

	 '변경이 없을때 작업진행여부 체크 %>
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

'====================================================================================================
Function PlantChange(ByVal CRow)
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	Dim iNameArr, iSpecArr, iHscdArr, iUnitArr, iVatTypeArr, iTrackingFlg 
	Dim strSelectList
	Dim iItem, iPlant
	
	frm1.vspdData.Col = C_ItemCd
	frm1.vspdData.Row = CRow
	iItem = frm1.vspdData.text
	
	frm1.vspdData.Col = C_PlantCd
	iPlant = frm1.vspdData.text
	
    Err.Clear

	Call CommonQueryRs(" item_nm, spec, hs_cd, basic_unit, vat_type, tracking_flg  ", " Dbo.ufn_s_ListItemInfo( " & FilterVar(iItem, "''", "S") & ",  " & FilterVar(iPlant, "''", "S") & ")", " ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If lgF0 = "" Then Exit Function

    iTrackingFlg	= Split(lgF5, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.description 
		Err.Clear 
		Exit Function
	End If

	With frm1.vspdData
	
		.Col = C_TrackingFlg
		.text = iTrackingFlg(0)
	
	End With

	Call ChangeTrackingRetField(CRow)
End Function

'========================================================================================================
Function ItemByHScodeChange(CRow)
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	Dim iNameArr, iSpecArr, iHscdArr, iUnitArr, iVatTypeArr, iTrackingFlg 
	Dim strSelectList
	Dim iItem, iPlant

	frm1.vspdData.Col = C_ItemCd
	frm1.vspdData.Row = CRow
	iItem = frm1.vspdData.text
	
	frm1.vspdData.Col = C_PlantCd
	iPlant = frm1.vspdData.text
	
    Err.Clear

	If Len(iPlant) Then
		Call CommonQueryRs(" item_nm, spec, hs_cd, basic_unit, vat_type, tracking_flg  ", " Dbo.ufn_s_ListItemInfo( " & FilterVar(iItem, "''", "S") & ",  " & FilterVar(iPlant, "''", "S") & ")", " ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Else
		Call CommonQueryRs(" item_nm, spec, hs_cd, basic_unit, vat_type, tracking_flg  ", " Dbo.ufn_s_ListItemInfo( " & FilterVar(iItem, "''", "S") & ", default)", " ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	End If

	If lgF0 = "" Then Exit Function

    iNameArr		= Split(lgF0, Chr(11))
    iSpecArr		= Split(lgF1, Chr(11))
    iHscdArr		= Split(lgF2, Chr(11))
    iUnitArr		= Split(lgF3, Chr(11))
    iVatTypeArr		= Split(lgF4, Chr(11))
    iTrackingFlg	= Split(lgF5, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.description 
		Err.Clear 
		Exit Function
	End If

	With frm1.vspdData
	
		.Row = CRow
		.Col = C_ItemName
		.text = iNameArr(0)
	
		.Col = C_ItemSpec
		.text = iSpecArr(0)

		.Col = C_HsNo
		.text = iHscdArr(0)
	
		.Col = C_SoUnit
		.text = iUnitArr(0)
	
		.Col	= C_VatType
		If .text = "" Then
			If Len(frm1.txtHVATType.value) Then
				.text	= frm1.txtHVATType.value
			Else
				.text	= iVatTypeArr(0)
			End If
		End If

		.Col	= C_VatIncFlag
		If .text = "" Then 
			If Len(frm1.txtHVATIncFlag.value) Then
				.Col	= C_VatIncFlag
				.text	= frm1.txtHVATIncFlag.value

				.Col	= C_VatIncFlagNm
				Select Case frm1.txtHVATIncFlag.value
				Case "1"
					frm1.vspdData.Text = "별도"
				Case "2"
					frm1.vspdData.Text = "포함"
				End Select
			End If			
		End If

		.Col = C_TrackingFlg
		.text = iTrackingFlg(0)
	
	End With

	Call SetVatType(CRow)
				
	lsPriceQty = "Q"

	Call GetItemPrice(CRow)

	Call PricePadChange(CRow)
	Call ChangeTrackingRetField(CRow)

End Function


'========================================================================================================
Function SetTrackingNoByItem(CRow)
	
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	Dim iNameArr, iSpecArr, iHscdArr, iUnitArr, iVatTypeArr, iTrackingFlg 
	Dim strSelectList
	Dim iItem, iPlant

	frm1.vspdData.Col = C_ItemCd
	frm1.vspdData.Row = CRow
	iItem = frm1.vspdData.text
	
	frm1.vspdData.Col = C_PlantCd
	iPlant = frm1.vspdData.text
	
    Err.Clear

	If Len(iPlant) Then
		Call CommonQueryRs(" item_nm, spec, hs_cd, basic_unit, vat_type, tracking_flg  ", " Dbo.ufn_s_ListItemInfo( " & FilterVar(iItem, "''", "S") & ",  " & FilterVar(iPlant, "''", "S") & ")", " ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Else
		Call CommonQueryRs(" item_nm, spec, hs_cd, basic_unit, vat_type, tracking_flg  ", " Dbo.ufn_s_ListItemInfo( " & FilterVar(iItem, "''", "S") & ", default)", " ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	End If

	If lgF0 = "" Then Exit Function
    iTrackingFlg	= Split(lgF5, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.description 
		Err.Clear 
		Exit Function
	End If

	With frm1.vspdData
	
		.Row = CRow
		.Col = C_TrackingFlg
		.text = iTrackingFlg(0)
	
	End With	

	Call ChangeTrackingRetField(CRow)

End Function


'========================================================================================================
Function CheckCreditlimitSvr()

	Dim iStrVal
	
	If LayerShowHide(1) = False Then
		Exit Function
	End If
	    
	iStrVal = BIZ_PGM_ID & "?txtMode=" & "CheckCreditlimit"      
	iStrVal = iStrVal & "&txtCaller=" & "SC"								'확정처리시 
	iStrVal = iStrVal & "&txtHSoNo=" & Trim(frm1.txtHSoNo.value)
	iStrVal = iStrVal & "&txtTotalAmt=" & "0"
	
	Call RunMyBizASP(MyBizASP, iStrVal)	

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
		strVal = strVal & "&RdoDnReq=" & Trim(frm1.RdoDnReq.value)   
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtInsrtUserId=" & Trim(parent.gUsrID)
	Else
		strVal = BIZ_PGM_ID & "?txtMode=" & "DNCheck"       
		strVal = strVal & "&txtSoNo=" & Trim(frm1.txtConSoNo.value)
		strVal = strVal & "&RdoDnReq=" & Trim(frm1.RdoDnReq.value) 
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
	 '수주순금액 
		ggoOper.FormatFieldByObjectOfCur .txtNetAmt, .txtCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
	 
	End With
End Sub

'========================================================================================================
Sub CurFormatNumSprSheet()
	With frm1

		ggoSpread.Source = frm1.vspdData
		'단가 
		ggoSpread.SSSetFloatByCellOfCur C_SoPrice,-1, .txtCurrency.value, parent.ggUnitCostNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec,,,"Z"
		'금액 
		ggoSpread.SSSetFloatByCellOfCur C_NetAmt,-1, .txtCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloatByCellOfCur C_TotalAmt,-1, .txtCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec,,,"Z"
		
		ggoSpread.SSSetFloatByCellOfCur C_OldNetAmt,-1, .txtCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloatByCellOfCur C_OriginalNetAmt,-1, .txtCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec,,,"Z"
		
		'VAT금액		
		ggoSpread.SSSetFloatByCellOfCur C_VATAmt,-1, .txtCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec,,,"Z"

	End With
End Sub


'========================================================================================================
Sub ChangePlantColor()
	If frm1.txtHConfirmFlg.value = "Y" Then
		Call ggoOper.SetReqAttr(frm1.txtPlant, "Q")
	Else
		Call ggoOper.SetReqAttr(frm1.txtPlant, "D")
	End If 
End Sub


'========================================================================================================
Function GetItemPrice(IRow)

	GetItemPrice = False
	
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
		strPayMeth = .txtHPayTermsCd.value			'결제방법 
		strCurrency = .txtCurrency.value			'화폐단위 
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
		frm1.vspdData.Row=iRow
		frm1.vspdData.Col = C_SoPrice
		
		If strItemInfo(1) <> 0 Then
			GetItemPrice = True			
'			frm1.vspdData.text = UNIFormatNumber(strItemInfo(1), ggUnitCost.DecPoint, -2, 0, ggUnitCost.RndPolicy, ggUnitCost.RndUnit)
			frm1.vspdData.Text = UNIFormatNumberByCurrecny(UNICDbl(strItemInfo(1)), strCurrency, parent.ggUnitCostNo) '2006-05-04 박정순 수정 
			Call QtyPriceChange(IRow)			
		Else
'			frm1.vspdData.text = UNIFormatNumber(strItemInfo(1), ggUnitCost.DecPoint, -2, 0, ggUnitCost.RndPolicy, ggUnitCost.RndUnit)
			frm1.vspdData.Text = UNIFormatNumberByCurrecny(UNICDbl(strItemInfo(1)), strCurrency, parent.ggUnitCostNo) '2006-05-04 박정순 수정 
			GetItemPrice = False
		End If	
	Else		
		If Err.number <> 0 Then
			MsgBox Err.description 
			Err.Clear 
			Exit Function
		ELSE
			Call DisplayMsgBox("171214", "X", "X", "X")			
		End If
	End if 
'품목단가에 대한 단가구분처리 add 20050526 byHJO
	Call setPriceFlag(iRow,strItemCd, strSoldToParty, strSOUnit, strCurrency, strValidDt,strDealType, strPayMeth)
End Function
'========================================================================================================
Function setPriceFlag(iRow,strItemCd, strSoldToParty, strSOUnit, strCurrency, strValidDt, strDealType, strPayMeth)

	Dim strSelectList, strFromList, strWhereList
	Dim strRs, strItemFlag
	
	setPriceFlag = False
	
	strSelectList = " dbo.ufn_s_GetItemPriceFlag( " & FilterVar(strSoldToParty, "''", "S") & ",  " & FilterVar(strItemCd, "''", "S") & ", " & FilterVar(strDealType, "''", "S") & ",  " & FilterVar(strPayMeth, "''", "S") & "," & _
					" " & FilterVar(strSOUnit, "''", "S") & ",  " & FilterVar(strCurrency, "''", "S") & ",  " & FilterVar(strValidDt, "''", "S") & ")"
	strFromList  = ""
	strWhereList = ""
    Err.Clear
	'단가구분처리 Fetch
	If CommonQueryRs2by2(strSelectList,strFromList,strWhereList,strRs) Then

		strItemFlag = Split(strRs, Chr(11))

		frm1.vspdData.Row=iRow
		frm1.vspdData.Col = C_SoPriceFlag
		If strItemFlag(1)="T" then
			frm1.vspdData.text = "진단가"
			frm1.HPriceFlag.value="Y"
		Else
			frm1.vspdData.text ="가단가"
			frm1.HPriceFlag.value="N"
		End If 
				
	Else		
		If Err.number <> 0 Then
			MsgBox Err.description 
			Err.Clear 
			Exit Function
		ELSE
			Call DisplayMsgBox("171214", "X", "X", "X")			
		End If
	End if 


End Function

'========================================================================================================
Function GetNumberingRuleforTracking()
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	Dim iCodeArr

    Err.Clear

	Call CommonQueryRs(" MINOR_CD", " B_CONFIGURATION ", "  major_cd = " & FilterVar("S0024", "''", "S") & " and seq_no = 1 and Reference = " & FilterVar("Y", "''", "S") & " ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If Len(lgF0) Then 
	    iCodeArr = Split(lgF0, Chr(11))

		frm1.txtHTrackingNORule.value = iCodeArr(0)	
	Else
		frm1.txtHTrackingNORule.value = "A"
	End If	

	If Err.number <> 0 Then
		MsgBox Err.description 
		Err.Clear 
		Exit Function
	End If
End Function


'========================================================================================================
Sub ConfirmSO()
	
	Dim iStrVal

	iStrVal = BIZ_PGM_ID & "?txtMode=" & "btnCONFIRM"					
	iStrVal = iStrVal & "&txtHSoNo=" & Trim(frm1.txtHSoNo.value)		
	iStrVal = iStrVal & "&RdoConfirm=" & Trim(frm1.RdoConfirm.value)	

	Call RunMyBizASP(MyBizASP, iStrVal)          

End Sub


'========================================================================================================
Sub btnConfirm_OnClick()

	If BtnSpreadCheck = False Then Exit Sub

	Err.Clear                                                        

	If Trim(frm1.RdoConfirm.value) = "Y" Then						'확정처리시 여신한도 체크 
		Call CheckCreditlimitSvr
	Else
	    Call ConfirmSO()	
	End If

End Sub



'========================================================================================================
Sub btnATPCheck_OnClick()

	Dim arrParam(1)
	Dim strRet
	Dim iCalledAspName

	ggoSpread.Source = frm1.vspdData
 		 
    frm1.vspdData.Col = 0 'ColHeader
    frm1.vspdData.Row = frm1.vspdData.ActiveRow
    
    If frm1.vspdData.text <> "" Then
		Call DisplayMsgBox("203243", "X", "X", "X")			
		Exit sub
    End If
    
	frm1.vspdData.Col = C_SoSeq
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
 
	arrParam(0) = Trim(frm1.txtHSoNo.value)     
	arrParam(1) = frm1.vspdData.Text
	
	iCalledAspName = AskPRAspName("s3112ra3")	
	If Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3112ra3", "x")
		IsOpenPop = False
		Exit Sub
	End If
	
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
		MsgBox "CTP Check 할 품목을 선택하세요", vbExclamation, parent.gLogoName
	    Exit Sub
	End If

	ggoSpread.Source = frm1.vspdData
	frm1.vspdData.Row = frm1.vspdData.ActiveRow 
	frm1.vspdData.Col = 0
	
	Select Case frm1.vspdData.Text
	Case ggoSpread.InsertFlag
		MsgBox "신규입력사항을 저장후에 CTP CHECK를 할수 있습니다.", vbExclamation, parent.gLogoName
		Exit Sub
	Case ggoSpread.UpdateFlag
		MsgBox "수정사항을 저장후에 CTP CHECK를 할수 있습니다.", vbExclamation, parent.gLogoName
		Exit Sub
	End Select
  
	'--- 출고요청이 진행되었는지를 확인한다.
	Dim SoSts
	frm1.vspdData.Row = frm1.vspdData.ActiveRow 
	frm1.vspdData.Col = C_SoSts
	SoSts = CInt(Trim(frm1.vspdData.Text))

	'---SoStatus 수주등록:0,출고요청:1,출고완료:2,매출등록:3

	If SoSts > 0 Then
		MsgBox "이미 CTP CHECK가 완료된 제품입니다.", vbExclamation, parent.gLogoName
		Exit Sub
	End If

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

	frm1.vspdData.Col = C_DnReqDt '출하요청일자 
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
	
	Dim iCalledAspName
	iCalledAspName = AskPRAspName("s3113pa1")	
	If Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3113pa1", "x")
		IsOpenPop = False
		Exit Sub
	End If
	
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
Function OpenStyleRef()
 
	Dim iIntRetVal
	Dim iCalledAspName
	
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	
	If IsOpenPop Then Exit Function
	
	If Trim(frm1.txtHSoNo.value) = "" Then
		' 조회를 선행 하십시오.
		Call DisplayMsgBox("900002", "x", "x", "x")
		frm1.txtConSoNo.focus
		Exit Function
	End If

	If Trim(frm1.RdoConfirm.value) = "N" Then	
		Msgbox "확정처리된 품목은 수주내역을 참조 할 수 없습니다",vbInformation, parent.gLogoName
		Exit Function
	End If
	
	Call CommonQueryRs(" A.PROJECT_CODE , A.PROJECT_NM ", " PMS_PROJECT A ", " SO_NO = " & FilterVar(Trim(frm1.txtHSoNo.value), "''", "S") & " ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	

	If Trim(lgF0) <> "" then
		'msgbox "프로젝트 수주참조된건은 참조할수 없습니다.message code 추가"
		Call DisplayMsgBox("YM0049", "x", "x", "x")	
		Exit Function
    End if
    
	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("s3112ra9")	
	If Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3112ra9", "x")
		IsOpenPop = False
		Exit Function
	End If
	
	iIntRetVal = window.showModalDialog(iCalledAspName, self, _
	  "dialogWidth=850px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False

	If iIntRetVal = vbYes Then
		Call fncSave()
	End If
	
End Function

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
		ElseIf Row > 0 And Col = C_TrackingNoPopup Then
			.Col = C_TrackingNo
			.Row = Row
			Call OpenSoDtl(.Text, 2)		
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
		Call PricePadChange(Row)	 'PS3G112.cSGetSoPriceSvr(덤수량 가져옴)
		Call QtyPriceChange(Row)		 'N

	 Case C_SoPrice
		Call QtyPriceChange(Row)		 'N
	 
	 Case C_ItemCd

		Call ItemByHScodeChange(Row) 

	 Case C_SoUnit
		Call GetItemPrice(Row)		 'CommonQueryRs2by2()
		lsPriceQty = "Q"
		Call PricePadChange(Row)	 'PS3G112.cSGetSoPriceSvr
 	 
 	 Case C_VatType
		Call SetVatType(Row)

	 Case C_VatIncFlag
		Call QtyPriceChange(Row)		 'N

	 Case C_TotalAmt
		Call TotalAmtChange(Row)
	'----------------------------
	' V3.0 : Tracking No 수동채번 
	'----------------------------
	Case C_PlantCd
		Call PlantChange(Row)		
		
	 End Select
	 '------ Developer Coding part (End   ) -------------------------------------------------------------- 
             
	Call CheckMinNumSpread(frm1.vspdData, Col, Row)		
    
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

	lgBlnFlgChgValue = True

End Sub


'========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	
	Call SetPopupMenuItemInf(Mid(gToolBarBit, 6, 2) + "0" + Mid(gToolBarBit, 8, 1) & "111111")
	
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
	   
       frm1.btnATPCheck.disabled = True
	   frm1.btnCTPCheck.disabled = True
       Exit Sub
    End If
    
    '------ Developer Coding part (Start ) -------------------------------------------------------------- 
	If ButtonVisible(Row) = False Then Exit Sub    
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
	If frm1.RdoConfirm.Value = "Y" Then   
		Call SetPopupMenuItemInf("1111111111")   
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
End Sub



'==========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub

'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub


'========================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
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
