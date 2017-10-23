<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : 영업																		*
'*  2. Function Name        : 출하관리																	*
'*  3. Program ID           : S3112AA1																	*
'*  4. Program Name         : 수주내역참조																*
'*  5. Program Desc         : 출하내역등록을 위한 수주내역참조 (Business Logic Asp)						*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2000/03/21																*
'*  8. Modified date(Last)  : 2003/06/11																*
'*  9. Modifier (First)     : Cho Song Hyon																*
'* 10. Modifier (Last)      : Hwang Seongbae															*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/21 : 화면 design												*
'*				            : 2. 2000/09/21 : 4th Coding												*
'*				            : 3. 2001/12/19 : Date 표준적용												*
'*                          : 4. 2002/12/23 : include 성능향상만 반영									*
'********************************************************************************************************
%>
<HTML>
<HEAD>
<TITLE>수주내역참조</TITLE>
<!-- #Include file="../../inc/IncSvrCcm.inc" --> 
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<Script Language="VBScript">

Option Explicit

'========================================
Dim arrParent, lgDn_Type, lgFLAG
ArrParent = window.dialogArguments
Set PopupParent  = ArrParent(0)
top.document.title = PopupParent.gActivePRAspName

Const BIZ_PGM_QRY_ID    = "s3112ab1_KO441.asp"	    		
Const BIZ_PGM_QRY_ID2   = "s3112ab2_KO441.asp"		    '20080220::hanc	

Dim lgstrprevkey2
' Popup Index
Const C_PopShipToParty	= 1			' 납품처 
Const C_PopSlCd			= 2			' 창고 
Const C_PopItemCd		= 3			' 품목코드 
Const C_PopSoNo			= 4			' S/O 번호 
Const C_PopTrackingNo	= 5			' Tracking 번호 

Dim lgOldRow


'☆: Spread Sheet 의 Columns 인덱스 
Dim C_ShipToParty
Dim C_ShipToPartyNm
Dim C_SoNo
Dim C_SoSeq
Dim C_SoSchdNo
Dim C_ItemCd
Dim C_ItemNm
Dim C_TrackingNo
Dim C_Unit
Dim C_Qty
Dim C_OnStkQty
Dim C_BasicUnit
Dim C_BonusQty
Dim C_PlannedGIDt
Dim C_PlantCd
Dim C_PlantNm
Dim C_SlCd
Dim C_SlNm
Dim C_TolMoreQty
Dim C_TolLessQty
Dim C_LcNo
Dim C_LcSeq
Dim C_Remark
Dim C_RetItemFlag
Dim C_LotFlag
Dim C_LotNo
Dim C_LotSeq
Dim C_RetType
Dim C_RetTypeNm
Dim C_Spec

'20080220::HANC Grid 2(vspdData2) - Operation
dim C_OUT_NO             
dim C_BP_CD              
dim C_BP_NM              
dim C_ITEM_CD            
dim C_PLANT_CD           
dim C_SL_CD              
dim C_OUT_TYPE           
dim C_UD_MINOR_NM        
dim C_GOOD_ON_HAND_QTY   
dim C_GI_QTY             
dim C_GI_UNIT            
dim C_LOT_NO             
dim C_LOT_SUB_NO         
dim C_ACTUAL_GI_DT       
dim C_ITEM_NM            
dim C_SPEC1              
dim C_rcpt_lot_no        
dim C_CUST_LOT_NO        
'2008-06-16 6:13오후 :: hanc
dim C_pgm_name      
dim C_pgm_price     

dim C_TRANS_TIME         
dim C_CREATE_TYPE        

'Dim C_OUT_NO
'Dim C_SHIP_TO_PARTY
'Dim C_SHIP_TO_PARTY_NM
'Dim C_ITEM_CD
'Dim C_PLANT_CD
'Dim C_OUT_TYPE
'Dim C_OUT_TYPE_NM
'Dim C_GI_QTY
'Dim C_GI_UNIT
'Dim C_LOT_NO
'Dim C_LOT_SEQ
'Dim C_ACTUAL_GI_DT
'Dim C_ITEM_NM
'Dim C_SPEC1


Dim C2_ShipToParty
Dim C2_ShipToPartyNm
Dim C2_SoNo
Dim C2_SoSeq
Dim C2_SoSchdNo
Dim C2_ItemCd
Dim C2_ItemNm
Dim C2_TrackingNo
Dim C2_Unit
Dim C2_Qty
Dim C2_OnStkQty
Dim C2_BasicUnit
Dim C2_BonusQty
Dim C2_PlannedGIDt
Dim C2_PlantCd
Dim C2_PlantNm
Dim C2_SlCd
Dim C2_SlNm
Dim C2_TolMoreQty
Dim C2_TolLessQty
Dim C2_LcNo
Dim C2_LcSeq
Dim C2_Remark
Dim C2_RetItemFlag
Dim C2_LotFlag
Dim C2_LotNo
Dim C2_LotSeq
Dim C2_RetType
Dim C2_RetTypeNm
Dim C2_Spec

'========================================
<!-- #Include file="../../inc/lgvariables.inc" -->

Dim arrReturn
Dim gblnWinEvent
Dim lgStrAllocInvFlag		' 재고할당 사용여부 

'========================================
Function InitVariables()
	lgIntGrpCount = 0										
	lgStrPrevKey = ""
	lgstrprevkey2 = "" '2008-06-10 4:59오후 :: hanc
	lgSortKey = 1										
	lgIntFlgMode = PopupParent.OPMD_CMODE										
		
	gblnWinEvent = False
	ReDim arrReturn(0,0)
	Self.Returnvalue = arrReturn
End Function

'========================================
Sub SetDefaultVal()

	Dim arrTemp
		
	arrTemp = Split(ArrParent(1), PopupParent.gColSep)

	txtSoNo.value = arrTemp(0)
	txtToDt.Text = arrTemp(1)
	txtHMovType.value = arrTemp(2)
	txtShipToParty.value = arrTemp(3)
	txtShipToPartyNm.value = arrTemp(4)
	txtHSoType.value = arrTemp(5)
	txtHRetFlag.value = arrTemp(6)
	txtHPlantCd.value = arrTemp(7)
	lgDn_Type	    = arrTemp(8)      '출하형태

	txtFromDt.text = UnIDateAdd("D", -7, arrTemp(1), PopupParent.gDateFormat)

	' 수주번호 지정인 경우 
	If Len(Trim(txtSoNo.value)) Then Call ggoOper.SetReqAttr(txtSoNo, "Q")

End Sub

'=========================================
'20080220::hanc Sub initSpreadPosVariables()
Sub InitSpreadPosVariables(ByVal pvSpdNo)
	Select Case UCase(pvSpdNo)
		Case "A"
        	C_PlannedGIDt	= 1
        	C_ItemCd		= 2
        	C_ItemNm		= 3
        	C_Qty			= 4
        	C_BonusQty		= 5
        	C_Unit			= 6
        	C_OnStkQty		= 7
        	C_BasicUnit		= 8
        	C_SoNo			= 9
        	C_SoSeq			= 10
        	C_SoSchdNo		= 11
        	C_TrackingNo	= 12
        	C_ShipToParty	= 13
        	C_ShipToPartyNm = 14
        	C_PlantCd		= 15
        	C_PlantNm		= 16
        	C_SlCd			= 17
        	C_SlNm			= 18
        	C_TolMoreQty	= 19
        	C_TolLessQty	= 20
        	C_LcNo			= 21
        	C_LcSeq			= 22
        	C_LotFlag		= 23
        	C_LotNo			= 24
        	C_LotSeq		= 25
        	C_RetItemFlag	= 26
        	C_RetType		= 27
        	C_RetTypeNm		= 28
        	C_Spec			= 29
        	C_Remark		= 30
		Case "B"
            C_OUT_NO                  = 1
            C_BP_CD                   = 2
            C_BP_NM                   = 3
            C_ITEM_CD                 = 4
            C_PLANT_CD                = 5
            C_SL_CD                   = 6
            C_OUT_TYPE                = 7
            C_UD_MINOR_NM             = 8
            C_GOOD_ON_HAND_QTY        = 9
            C_GI_QTY                  = 10
            C_GI_UNIT                 = 11
            C_LOT_NO                  = 12
            C_LOT_SUB_NO              = 13
            C_ACTUAL_GI_DT            = 14
            C_ITEM_NM                 = 15
            C_SPEC1                   = 16
            C_rcpt_lot_no             = 17
            C_CUST_LOT_NO             = 18
            C_pgm_name                = 19      '2008-06-16 6:17오후 :: hanc
            C_pgm_price               = 20      '2008-06-16 6:17오후 :: hanc
            C_TRANS_TIME              = 21
            C_CREATE_TYPE             = 22

'        	C_OUT_NO                 = 1
'        	C_SHIP_TO_PARTY          = 2
'        	C_SHIP_TO_PARTY_NM       = 3
'        	C_ITEM_CD                = 4
'        	C_PLANT_CD               = 5
'        	C_OUT_TYPE               = 6
'        	C_OUT_TYPE_NM            = 7
'        	C_GI_QTY                 = 8
'        	C_GI_UNIT                = 9
'        	C_LOT_NO                 = 10
'        	C_LOT_SEQ                = 11
'            C_ACTUAL_GI_DT           = 12
'            C_ITEM_NM           = 13
'            C_SPEC1           = 14

'        	C2_PlannedGIDt	= 1
'        	C2_ItemCd		= 2
'        	C2_ItemNm		= 3
'        	C2_Qty			= 4
'        	C2_BonusQty		= 5
'        	C2_Unit			= 6
'        	C2_OnStkQty		= 7
'        	C2_BasicUnit		= 8
'        	C2_SoNo			= 9
'        	C2_SoSeq			= 10
'        	C2_SoSchdNo		= 11
'        	C2_TrackingNo	= 12
'        	C2_ShipToParty	= 13
'        	C2_ShipToPartyNm = 14
'        	C2_PlantCd		= 15
'        	C2_PlantNm		= 16
'        	C2_SlCd			= 17
'        	C2_SlNm			= 18
'        	C2_TolMoreQty	= 19
'        	C2_TolLessQty	= 20
'        	C2_LcNo			= 21
'        	C2_LcSeq			= 22
'        	C2_LotFlag		= 23
'        	C2_LotNo			= 24
'        	C2_LotSeq		= 25
'        	C2_RetItemFlag	= 26
'        	C2_RetType		= 27
'        	C2_RetTypeNm		= 28
'        	C2_Spec			= 29
'        	C2_Remark		= 30
	End Select			
End Sub

'=====================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "RA") %>
End Sub

'========================================
'20080220::hanc Sub InitSpreadSheet()
Sub InitSpreadSheet(ByVal pvSpdNo)
	Select Case UCase(pvSpdNo)
		Case "A"
        	Call initSpreadPosVariables(pvSpdNo)    
        	
        	With ggoSpread
        
        		.Source = vspdData1
        		.Spreadinit "V20030901",,PopupParent.gAllowDragDropSpread    
        
        		vspdData1.MaxCols = C_Remark + 1
        		vspdData1.MaxRows = 0
        
        		vspdData1.ReDraw = False
        
        		Call GetSpreadColumnPos("A")
        
        		Call AppendNumberPlace("7","5","0")
        
        		.SSSetFloat C_LotSeq,"LOT NO 순번" ,15,"7", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec,,,"Z"  
        
        		.SSSetDate		C_PlannedGIDt, "출하예정일",15,2,PopupParent.gDateFormat
        		.SSSetEdit		C_ItemCd, "품목", 20, 0
        		.SSSetEdit		C_ItemNm, "품목명", 40, 0
        		.SSSetFloat		C_Qty,"미출고수량" ,15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
        		.SSSetFloat		C_BonusQty,"미출고덤수량" ,15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
        		.SSSetEdit		C_Unit, "단위", 10, 2
        		.SSSetFloat		C_OnStkQty,"재고량" ,15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
        		.SSSetEdit		C_BasicUnit, "재고단위", 10, 2
        		.SSSetEdit		C_SoNo, "수주번호", 18, 0
        		.SSSetFloat		C_SoSeq,"수주순번" ,15,"7", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec,,,"Z"  
        		.SSSetFloat		C_SoSchdNo,"납품순번" ,15,"7", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec,,,"Z"  
        		.SSSetEdit		C_TrackingNo, "Tracking No", 18, 0
        		.SSSetEdit		C_ShipToParty, "납품처", 15, 0
        		.SSSetEdit		C_ShipToPartyNm, "납품처명", 25, 0
        		.SSSetEdit		C_PlantCd, "공장", 15, 0
        		.SSSetEdit		C_PlantNm, "공장명", 20, 0
        		.SSSetEdit		C_SlCd, "창고", 15, 0
        		.SSSetEdit		C_SlNm, "창고명", 20, 0
        		.SSSetFloat		C_TolMoreQty,"과부족허용량(+)" ,20,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
        		.SSSetFloat		C_TolLessQty,"과부족허용량(-)" ,20,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
        		.SSSetEdit		C_LcNo, "L/C번호", 18, 0
        		.SSSetFloat		C_LcSeq,"L/C순번" ,15,"7", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec,,,"Z"  
        							   'ColumnPosition		Header              Width	Align(0:L,1:R,2:C)  Row   Length  CharCase(0:L,1:N,2:U)
        		.SSSetEdit		C_LotFlag, "LOT관리여부", 15, 2
        		.SSSetEdit		C_LotNo, "LOT NO", 18, 0
        		.SSSetFloat		C_LotSeq,"LOT NO 순번" ,15,"7", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec,,,"Z"  
        		.SSSetEdit		C_RetItemFlag, "반품여부", 10, 2
        		.SSSetEdit		C_RetType, "반품유형", 15, 0
        		.SSSetEdit		C_RetTypeNm, "반품유형명", 20, 0
        		.SSSetEdit		C_Spec, "규격", 30, 0
        		.SSSetEdit		C_Remark,			"비고",			30,	0,					,	  120
        
        		Call .SSSetColHidden(C_PlantCd,C_PlantNm,True)
        		Call .SSSetColHidden(vspdData1.MaxCols,vspdData1.MaxCols,True)
        			
        		vspdData1.ReDraw = True
        	End With
		Case "B"
        	Call initSpreadPosVariables(pvSpdNo)    
        	
        	With ggoSpread
        
        		.Source = vspdData2
        		.Spreadinit "V20030901",,PopupParent.gAllowDragDropSpread    
        
        		vspdData2.MaxCols = C_CREATE_TYPE + 1
        		vspdData2.MaxRows = 0
        
        		vspdData2.ReDraw = False
        
        		Call GetSpreadColumnPos("B")
        
        		Call AppendNumberPlace("7","5","0")
        
        		.SSSetEdit		C_OUT_NO                 , "출하(P/L NO)", 12, 0
        		.SSSetEdit		C_BP_CD                  , "납품처코드", 12, 0
        		.SSSetEdit		C_BP_NM                  , "납품처명", 12, 0
        		.SSSetEdit		C_ITEM_CD                , "품목코드", 12, 0
        		.SSSetEdit		C_PLANT_CD               , "공장코드", 8, 0
        		.SSSetEdit		C_SL_CD                  , "창고코드", 8, 0
        		.SSSetEdit		C_OUT_TYPE               , "MES출하TYPE", 8, 0
        		.SSSetEdit		C_UD_MINOR_NM            , "MES출하TYPE명", 20, 0
        		.SSSetFloat		C_GOOD_ON_HAND_QTY       , "재고량" ,13,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
        		.SSSetFloat		C_GI_QTY                 , "출고량" ,13,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
        		.SSSetEdit		C_GI_UNIT                , "출고단위", 10, 0
        		.SSSetEdit		C_LOT_NO                 , "LOT번호", 13, 0
        		.SSSetEdit		C_LOT_SUB_NO             , "LOT순번", 10, 0
        		.SSSetDate		C_ACTUAL_GI_DT           , "출고일",12,2,PopupParent.gDateFormat
        		.SSSetEdit		C_ITEM_NM                 , "품목명", 10, 0
        		.SSSetEdit		C_SPEC1                  , "SPEC", 20, 0
        		.SSSetEdit		C_rcpt_lot_no            , "입고 LOT번호", 13, 0
        		.SSSetEdit		C_CUST_LOT_NO            , "고객 LOT번호", 13, 0
        		.SSSetEdit		C_pgm_name               , "PGM NAME", 50, 0
        		.SSSetFloat		C_pgm_price              , "PGM 적용단가" ,13,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
        		.SSSetEdit		C_TRANS_TIME             , "TRANS_TIME", 20, 0
        		.SSSetEdit		C_CREATE_TYPE            , "CREATE_TYPE", 20, 0
            
'        		.SSSetEdit		C_OUT_NO                 , "출하(P/L NO)", 20, 0
'        		.SSSetEdit		C_SHIP_TO_PARTY          , "납품처코드", 20, 0
'        		.SSSetEdit		C_SHIP_TO_PARTY_NM       , "납품처명", 20, 0
'        		.SSSetEdit		C_ITEM_CD                , "품목코드", 20, 0
'        		.SSSetEdit		C_PLANT_CD               , "공장코드", 20, 0
'        		.SSSetEdit		C_OUT_TYPE               , "MES출하TYPE", 20, 0
'        		.SSSetEdit		C_OUT_TYPE_NM            , "MES출하TYPE명", 20, 0
'        		.SSSetFloat		C_GI_QTY                 , "출고량" ,15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
'        		.SSSetEdit		C_GI_UNIT                , "출고단위", 20, 0
'        		.SSSetEdit		C_LOT_NO                 , "LOT번호", 20, 0
'        		.SSSetEdit		C_LOT_SEQ                , "LOT순번", 20, 0
'        		.SSSetDate		C_ACTUAL_GI_DT           , "출고일",15,2,PopupParent.gDateFormat
'        		.SSSetEdit		C_ITEM_NM                 , "품목명", 20, 0
'        		.SSSetEdit		C_SPEC1                , "SPEC", 20, 0

        
        		Call .SSSetColHidden(C_ITEM_NM,C_SPEC1,True)
        		Call .SSSetColHidden(C_TRANS_TIME,C_TRANS_TIME,True)
        		Call .SSSetColHidden(C_CREATE_TYPE,C_CREATE_TYPE,True)
        		Call .SSSetColHidden(vspdData2.MaxCols,vspdData2.MaxCols,True)
        			
        		vspdData2.ReDraw = True
        	End With
	End Select


	Call SetSpreadLock

End Sub

'========================================
Sub SetSpreadLock()
    '--------------------------------
    '20080220::hanc::Grid 1
    '--------------------------------
    ggoSpread.Source = vspdData1
	ggoSpread.SpreadLockWithOddEvenRowColor()
    
    '--------------------------------
    '20080220::hanc::Grid 2
    '--------------------------------
    ggoSpread.Source = vspdData2
	ggoSpread.SpreadLockWithOddEvenRowColor()

	vspdData1.OperationMode = 5	'Multi Select Mode
	vspdData2.OperationMode = 5	'Multi Select Mode      '20080225::hanc
End Sub

''========================================
'Function OKClick()
'	on error resume next
'	
'	Dim intColCnt, intRowCnt, intInsRow, iLngSelectedRows
'
'	With vspdData1
'		iLngSelectedRows = .SelModeSelCount
'		' 전체선택시 
'		If iLngSelectedRows = -1 Then
'			iLngSelectedRows = .MaxRows
'		End If
'
'		If iLngSelectedRows > 0 Then 
'			intInsRow = 0
'
'			Redim arrReturn(iLngSelectedRows, .MaxCols)
'
'			For intRowCnt = 1 To .MaxRows
'
'				.Row = intRowCnt
'
'				If .SelModeSelected Then
'					.Col = C_PlannedGIDt	: arrReturn(intInsRow, 0) = .Text
'					.Col = C_ItemCd			: arrReturn(intInsRow, 1) = .Text
'					.Col = C_ItemNm			: arrReturn(intInsRow, 2) = .Text
'					.Col = C_Qty			: arrReturn(intInsRow, 3) = .Text
'					.Col = C_BonusQty		: arrReturn(intInsRow, 4) = .Text
'					.Col = C_Unit			: arrReturn(intInsRow, 5) = .Text
'					.Col = C_OnStkQty		: arrReturn(intInsRow, 6) = .Text
'					.Col = C_BasicUnit		: arrReturn(intInsRow, 7) = .Text
'					.Col = C_SoNo			: arrReturn(intInsRow, 8) = .Text
'					.Col = C_SoSeq			: arrReturn(intInsRow, 9) = .Text
'					.Col = C_SoSchdNo		: arrReturn(intInsRow, 10) = .Text
'					.Col = C_TrackingNo		: arrReturn(intInsRow, 11) = .Text
'					.Col = C_ShipToParty	: arrReturn(intInsRow, 12) = .Text
'					.Col = C_ShipToPartyNm	: arrReturn(intInsRow, 13) = .Text
'					.Col = C_PlantCd		: arrReturn(intInsRow, 14) = .Text
'					.Col = C_PlantNm		: arrReturn(intInsRow, 15) = .Text
'					.Col = C_SlCd			: arrReturn(intInsRow, 16) = .Text
'					.Col = C_SlNm			: arrReturn(intInsRow, 17) = .Text
'					.Col = C_TolMoreQty		: arrReturn(intInsRow, 18) = .Text
'					.Col = C_TolLessQty		: arrReturn(intInsRow, 19) = .Text
'					.Col = C_LcNo			: arrReturn(intInsRow, 20) = .Text
'					.Col = C_LcSeq			: arrReturn(intInsRow, 21) = .Text
'					.Col = C_LotFlag		: arrReturn(intInsRow, 22) = .Text
'					.Col = C_LotNo			: arrReturn(intInsRow, 23) = .Text
'					.Col = C_LotSeq			: arrReturn(intInsRow, 24) = .Text
'					.Col = C_RetItemFlag	: arrReturn(intInsRow, 25) = .Text
'					.Col = C_RetType		: arrReturn(intInsRow, 26) = .Text
'					.Col = C_RetTypeNm		: arrReturn(intInsRow, 27) = .Text
'					.Col = C_Spec			: arrReturn(intInsRow, 28) = .Text
'					.Col = C_Remark			: arrReturn(intInsRow, 29) = .Text
'
'					intInsRow = intInsRow + 1
'
'				End IF
'			Next
'		End if			
'	End With
'
'	Self.Returnvalue = arrReturn
'	Self.Close()
'End Function

'20080220::HANC========================================
Function OKClick()
	on error resume next

	Dim intColCnt, intRowCnt, intInsRow, iLngSelectedRows

'	With vspdData1
		iLngSelectedRows = vspdData2.SelModeSelCount
		' 전체선택시 
		If iLngSelectedRows = -1 Then
			iLngSelectedRows = vspdData2.MaxRows
		End If


		If iLngSelectedRows > 0 Then 
			intInsRow = 0

'			Redim arrReturn(iLngSelectedRows, vspdData1.MaxCols)
			Redim arrReturn(iLngSelectedRows, 37)

			For intRowCnt = 1 To vspdData2.MaxRows

				vspdData2.Row = intRowCnt
				vspdData1.Row = vspdData1.ActiveRow

				If vspdData2.SelModeSelected Then


					vspdData2.Col = C_ACTUAL_GI_DT	: arrReturn(intInsRow, 0) = vspdData2.Text
					vspdData2.Col = C_ITEM_CD		: arrReturn(intInsRow, 1) = vspdData2.Text
					vspdData2.Col = C_ITEM_NM		: arrReturn(intInsRow, 2) = vspdData2.Text
					vspdData2.Col = C_GI_QTY		: arrReturn(intInsRow, 3) = vspdData2.Text      '출고량
					vspdData1.Col = C_BonusQty		: arrReturn(intInsRow, 4) = 0                   '덤수량
					vspdData2.Col = C_GI_UNIT		: arrReturn(intInsRow, 5) = vspdData2.Text      '단위
					vspdData1.Col = C_OnStkQty		: arrReturn(intInsRow, 6) = vspdData1.Text      '재고량
					vspdData1.Col = C_BasicUnit		: arrReturn(intInsRow, 7) = vspdData1.Text      '재고단위
					vspdData1.Col = C_SoNo			: arrReturn(intInsRow, 8) = vspdData1.Text      '수주번호
					vspdData1.Col = C_SoSeq			: arrReturn(intInsRow, 9) = vspdData1.Text      '수주순번
					vspdData1.Col = C_SoSchdNo		: arrReturn(intInsRow, 10) = vspdData1.Text     '납품순번
					vspdData1.Col = C_TrackingNo	: arrReturn(intInsRow, 11) = vspdData1.Text
					vspdData2.Col = C_SHIP_TO_PARTY	: arrReturn(intInsRow, 12) = vspdData2.Text     '납품처
					vspdData2.Col = C_SHIP_TO_PARTY_NM	: arrReturn(intInsRow, 13) = vspdData2.Text '납품처명
					vspdData2.Col = C_PLANT_CD		: arrReturn(intInsRow, 14) = vspdData2.Text     '공장
					vspdData1.Col = C_PlantNm		: arrReturn(intInsRow, 15) = vspdData1.Text
					vspdData1.Col = C_SlCd			: arrReturn(intInsRow, 16) = vspdData1.Text
					vspdData1.Col = C_SlNm			: arrReturn(intInsRow, 17) = vspdData1.Text
					vspdData1.Col = C_TolMoreQty		: arrReturn(intInsRow, 18) = vspdData1.Text
					vspdData1.Col = C_TolLessQty		: arrReturn(intInsRow, 19) = vspdData1.Text
					vspdData1.Col = C_LcNo			: arrReturn(intInsRow, 20) = vspdData1.Text
					vspdData1.Col = C_LcSeq			: arrReturn(intInsRow, 21) = vspdData1.Text
					vspdData1.Col = C_LotFlag		: arrReturn(intInsRow, 22) = vspdData1.Text
					vspdData2.Col = C_LOT_NO			: arrReturn(intInsRow, 23) = vspdData2.Text
					vspdData2.Col = C_LOT_SEQ			: arrReturn(intInsRow, 24) = vspdData2.Text
					vspdData1.Col = C_RetItemFlag	: arrReturn(intInsRow, 25) = vspdData1.Text
					vspdData1.Col = C_RetType		: arrReturn(intInsRow, 26) = vspdData1.Text
					vspdData1.Col = C_RetTypeNm		: arrReturn(intInsRow, 27) = vspdData1.Text
					vspdData2.Col = C_SPEC			: arrReturn(intInsRow, 28) = vspdData2.Text
					vspdData1.Col = C_remark		: arrReturn(intInsRow, 29) = vspdData1.Text
					vspdData2.Col = C_OUT_NO			: arrReturn(intInsRow, 30) = vspdData2.Text
					vspdData2.Col = C_TRANS_TIME		: arrReturn(intInsRow, 31) = vspdData2.Text
					vspdData2.Col = C_CREATE_TYPE		: arrReturn(intInsRow, 32) = vspdData2.Text
					vspdData2.Col = C_CUST_LOT_NO		: arrReturn(intInsRow, 33) = vspdData2.Text
					vspdData2.Col = C_RCPT_LOT_NO		: arrReturn(intInsRow, 34) = vspdData2.Text
					vspdData2.Col = C_pgm_name		: arrReturn(intInsRow, 35) = vspdData2.Text '2008-06-16 7:28오후 :: hanc
					vspdData2.Col = C_pgm_price		: arrReturn(intInsRow, 36) = vspdData2.Text

					intInsRow = intInsRow + 1

				End IF
			Next
		End if			
'	End With

	Self.Returnvalue = arrReturn
	Self.Close()
End Function


'========================================
Function CancelClick()
	Redim arrReturn(1,1)
	arrReturn(0,0) = ""
	Self.Returnvalue = arrReturn
	Self.Close()
End Function

'========================================
Function OpenConPopup(ByVal pvIntWhere)

	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	OpenConPopup = False
	
	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	Select Case pvIntWhere
		'납품처 
		Case C_PopShipToParty
			iArrParam(1) = "dbo.B_BIZ_PARTNER BP INNER JOIN dbo.B_COUNTRY CT ON (CT.COUNTRY_CD = BP.CONTRY_CD)"								
			iArrParam(2) = Trim(txtShipToParty.value)			
			iArrParam(3) = ""											
			iArrParam(4) = "BP.BP_TYPE IN (" & FilterVar("C", "''", "S") & " , " & FilterVar("CS", "''", "S") & ") AND EXISTS (SELECT * FROM B_BIZ_PARTNER_FTN BF WHERE BF.PARTNER_BP_CD = BP.BP_CD AND BF.PARTNER_FTN = " & FilterVar("SSH", "''", "S") & ")"						
	
			iArrField(0) = "ED15" & PopupParent.gColSep & "BP.BP_CD"
			iArrField(1) = "ED30" & PopupParent.gColSep & "BP.BP_NM"
			iArrField(2) = "ED10" & PopupParent.gColSep & "BP.CONTRY_CD"
			iArrField(3) = "ED20" & PopupParent.gColSep & "CT.COUNTRY_NM"
    
			iArrHeader(0) = txtShipToParty.alt					
			iArrHeader(1) = txtShipToPartyNm.alt					
			iArrHeader(2) = "국가"
			iArrHeader(3) = "국가명"

			txtShipToParty.focus

		'창고 
		Case C_PopSlCd
			iArrParam(1) = "dbo.B_STORAGE_LOCATION"								
			iArrParam(2) = Trim(txtSlCd.value)			
			iArrParam(3) = ""											
			iArrParam(4) = "PLANT_CD = " & FilterVar(txtHPlantCd.value, "''", "S") & ""
	
			iArrField(0) = "ED15" & PopupParent.gColSep & "SL_CD"
			iArrField(1) = "ED30" & PopupParent.gColSep & "SL_NM"
    
			iArrHeader(0) = txtSlCd.alt					
			iArrHeader(1) = txtSlNm.alt					

			txtSlCd.focus

		' 품목 
		Case C_PopItemCd
			OpenConPopup = OpenConItemPopup(C_PopItemCd, txtItemCd.value)
			txtItemCd.focus
			Exit Function

		' S/O 번호 
		Case C_PopSoNo
			iArrParam(1) = "S_SO_HDR SH, B_BIZ_PARTNER SP, B_SALES_GRP SG"
			iArrParam(2) = Trim(txtSoNo.value)
			iArrParam(3) = ""
	
			' 재고할당을 사용여부 
			If lgStrAllocInvFlag = "N" Then
				iArrParam(4) = "SH.SOLD_TO_PARTY = SP.BP_CD AND SH.SALES_GRP = SG.SALES_GRP AND SH.CFM_FLAG = " & FilterVar("Y", "''", "S") & "  AND SH.REL_DN_FLAG = " & FilterVar("Y", "''", "S") & "  AND EXISTS (SELECT * FROM S_SO_DTL SD WHERE SD.SO_NO = SH.SO_NO AND SD.SO_QTY + SD.BONUS_QTY > SD.REQ_QTY + SD.REQ_BONUS_QTY "
			Else
				iArrParam(4) = "SH.SOLD_TO_PARTY = SP.BP_CD AND SH.SALES_GRP = SG.SALES_GRP AND SH.CFM_FLAG = " & FilterVar("Y", "''", "S") & "  AND SH.REL_DN_FLAG = " & FilterVar("Y", "''", "S") & "  AND EXISTS (SELECT * FROM S_SO_SCHD SC INNER JOIN dbo.S_SO_DTL SD ON (SD.SO_NO = SC.SO_NO AND SD.SO_SEQ = SC.SO_SEQ) WHERE SC.SO_NO = SH.SO_NO AND SC.ALLC_QTY + SC.ALLC_BONUS_QTY > SC.REQ_QTY + SC.REQ_BONUS_QTY "
			End If
			' 공장 
			iArrParam(4) = iArrParam(4) & " AND SD.PLANT_CD =  " & FilterVar(txtHPlantCd.value, "''", "S") & ""
			' 납품처 
			If Trim(txtShipToParty.value) = "" Then
				iArrParam(4) = iArrParam(4) & ")"
			Else
				iArrParam(4) = iArrParam(4) & " AND SD.SHIP_TO_PARTY =  " & FilterVar(txtShipToParty.value, "''", "S") & ")"
			End If
	
			iArrParam(5) = "수주번호"

			iArrField(0) = "ED12" & PopupParent.gColSep & "SH.SO_NO"
			iArrField(1) = "ED10" & PopupParent.gColSep & "SH.SOLD_TO_PARTY"
			iArrField(2) = "ED20" & PopupParent.gColSep & "SP.BP_NM"
			iArrField(3) = "DD10" & PopupParent.gColSep & "SH.SO_DT"
			iArrField(4) = "ED15" & PopupParent.gColSep & "SG.SALES_GRP_NM"
			iArrField(5) = "ED10" & PopupParent.gColSep & "SH.PAY_METH"
						
			iArrHeader(0) = "수주번호"
			iArrHeader(1) = "주문처"
			iArrHeader(2) = "주문처명"
			iArrHeader(3) = "수주일"
			iArrHeader(4) = "영업그룹명"
			iArrHeader(5) = "결제방법"
			
			txtSoNo.focus
	End Select
 
	iArrParam(0) = iArrHeader(0)							' 팝업 Title
	iArrParam(5) = iArrHeader(0)							' 조회조건 명칭 

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If iArrRet(0) <> "" Then
		OpenConPopup = SetConPopup(iArrRet,pvIntWhere)
	End If	
	
End Function

'========================================
Function SetConPopup(Byval pvArrRet,ByVal pvIntWhere)

	SetConPopup = False

	Select Case pvIntWhere
		' 납품처 
		Case C_PopShipToParty
			txtShipToParty.value = pvArrRet(0)
			txtShipToPartyNm.value = pvArrRet(1) 

		' 창고 
		Case C_PopSlCd
			txtSlCd.value = pvArrRet(0) 
			txtSlNm.value = pvArrRet(1)   

		' 품목 
		Case C_PopItemCd
			txtItemCd.value = pvArrRet(0) 
			txtItemNm.value = pvArrRet(1)   

		' S/O 번호 
		Case C_PopSoNo
			txtSoNo.value = pvArrRet(0)

	End Select
	
	SetConPopup = True

End Function

'===========================================================================
' Function Name : OpenTrackingNo
' Function Desc : OpenTrackingNo Reference Popup
'===========================================================================
Function OpenTrackingNo()
	Dim iCalledAspName
	Dim strRet
	
	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	Dim arrTNParam(5), i

	For i = 0 to UBound(arrTNParam) - 1
		arrTNParam(i) = ""
	Next	    

	arrTNParam(5) = "DN"
	
	'20021227 kangjungu dynamic popup
	iCalledAspName = AskPRAspName("s3135pa3")	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3135pa3", "x")
		gblnWinEvent = False
		Exit Function
	End if
	gblnWinEvent = True

	strRet = window.showModalDialog(iCalledAspName, Array(PopupParent, arrTNParam), _
		"dialogWidth=655px; dialogHeight=400px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If strRet <> "" Then
		txtTrackingNo.value = strRet 
	End If		
		
	txtTrackingNo.focus
End Function

' Item Popup
'========================================
Function OpenConItemPopup(ByVal pvIntWhere, ByVal pvStrData)
	Dim iArrRet
	Dim iArrParam(3)
	Dim iCalledAspName

	OpenConItemPopup = False

	iCalledAspName = AskPRAspName("s2210pa1")
	
	If Trim(iCalledAspName) = "" Then
		Call DisplayMsgBox("900040", parent.VB_INFORMATION, "s2210pa1", "X")
		gblnWinEvent = False
		Exit Function
	End If

	iArrParam(0) = pvStrData
	iArrParam(3) = txtHPlantCd.value
	
	iArrRet = window.showModalDialog(iCalledAspName, Array(PopupParent,iArrParam), _
	 "dialogWidth=850px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If iArrRet(0) <> "" Then
		OpenConItemPopup = SetConPopup(iArrRet,pvIntWhere)
	End If	
End Function

'=====================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_PlannedGIDt	= iCurColumnPos(1)
			C_ItemCd		= iCurColumnPos(2)
			C_ItemNm		= iCurColumnPos(3)
			C_Qty			= iCurColumnPos(4)
			C_BonusQty		= iCurColumnPos(5)
			C_Unit			= iCurColumnPos(6)
			C_OnStkQty		= iCurColumnPos(7)
			C_BasicUnit		= iCurColumnPos(8)
			C_SoNo			= iCurColumnPos(9)
			C_SoSeq			= iCurColumnPos(10)
			C_SoSchdNo		= iCurColumnPos(11)
			C_TrackingNo	= iCurColumnPos(12)
			C_ShipToParty	= iCurColumnPos(13)
			C_ShipToPartyNm = iCurColumnPos(14)
			C_PlantCd		= iCurColumnPos(15)
			C_PlantNm		= iCurColumnPos(16)
			C_SlCd			= iCurColumnPos(17)
			C_SlNm			= iCurColumnPos(18)
			C_TolMoreQty	= iCurColumnPos(19)
			C_TolLessQty	= iCurColumnPos(20)
			C_LcNo			= iCurColumnPos(21)
			C_LcSeq			= iCurColumnPos(22)
			C_LotFlag		= iCurColumnPos(23)
			C_LotNo			= iCurColumnPos(24)
			C_LotSeq		= iCurColumnPos(25)
			C_RetItemFlag	= iCurColumnPos(26)
			C_RetType		= iCurColumnPos(27)
			C_RetTypeNm		= iCurColumnPos(28)
			C_Spec			= iCurColumnPos(29)
			C_Remark		= iCurColumnPos(30)
       Case "B"
            ggoSpread.Source = vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

    
            C_OUT_NO                 = iCurColumnPos(1)        
            C_BP_CD                  = iCurColumnPos(2)        
            C_BP_NM                  = iCurColumnPos(3)        
            C_ITEM_CD                = iCurColumnPos(4)        
            C_PLANT_CD               = iCurColumnPos(5)        
            C_SL_CD                  = iCurColumnPos(6)        
            C_OUT_TYPE               = iCurColumnPos(7)        
            C_UD_MINOR_NM            = iCurColumnPos(8)        
            C_GOOD_ON_HAND_QTY       = iCurColumnPos(9)        
            C_GI_QTY                 = iCurColumnPos(10)        
            C_GI_UNIT                = iCurColumnPos(11)        
            C_LOT_NO                 = iCurColumnPos(12)        
            C_LOT_SUB_NO             = iCurColumnPos(13)        
            C_ACTUAL_GI_DT           = iCurColumnPos(14)        
            C_ITEM_NM                = iCurColumnPos(15)        
            C_SPEC1                  = iCurColumnPos(16)        
            C_rcpt_lot_no            = iCurColumnPos(17)        
            C_CUST_LOT_NO            = iCurColumnPos(18)        
            C_pgm_name            = iCurColumnPos(19)        
            C_pgm_price            = iCurColumnPos(20)        
            C_TRANS_TIME             = iCurColumnPos(21)        
            C_CREATE_TYPE            = iCurColumnPos(22)  
            
'            C_OUT_NO                 = iCurColumnPos(1)
'            C_SHIP_TO_PARTY          = iCurColumnPos(2)
'            C_SHIP_TO_PARTY_NM       = iCurColumnPos(3)
'            C_ITEM_CD                = iCurColumnPos(4)
'            C_PLANT_CD               = iCurColumnPos(5)
'            C_OUT_TYPE               = iCurColumnPos(6)
'            C_OUT_TYPE_NM            = iCurColumnPos(7)
'            C_GI_QTY                 = iCurColumnPos(8)
'            C_GI_UNIT                = iCurColumnPos(9)
'            C_LOT_NO                 = iCurColumnPos(10)
'            C_LOT_SEQ                = iCurColumnPos(11)
'            C_ACTUAL_GI_DT           = iCurColumnPos(12)
'            C_ITEM_NM                = iCurColumnPos(13)
'            C_SPEC1                  = iCurColumnPos(14)

'			C2_PlannedGIDt	= iCurColumnPos(1)
'			C2_ItemCd		= iCurColumnPos(2)
'			C2_ItemNm		= iCurColumnPos(3)
'			C2_Qty			= iCurColumnPos(4)
'			C2_BonusQty		= iCurColumnPos(5)
'			C2_Unit			= iCurColumnPos(6)
'			C2_OnStkQty		= iCurColumnPos(7)
'			C2_BasicUnit		= iCurColumnPos(8)
'			C2_SoNo			= iCurColumnPos(9)
'			C2_SoSeq			= iCurColumnPos(10)
'			C2_SoSchdNo		= iCurColumnPos(11)
'			C2_TrackingNo	= iCurColumnPos(12)
'			C2_ShipToParty	= iCurColumnPos(13)
'			C2_ShipToPartyNm = iCurColumnPos(14)
'			C2_PlantCd		= iCurColumnPos(15)
'			C2_PlantNm		= iCurColumnPos(16)
'			C2_SlCd			= iCurColumnPos(17)
'			C2_SlNm			= iCurColumnPos(18)
'			C2_TolMoreQty	= iCurColumnPos(19)
'			C2_TolLessQty	= iCurColumnPos(20)
'			C2_LcNo			= iCurColumnPos(21)
'			C2_LcSeq			= iCurColumnPos(22)
'			C2_LotFlag		= iCurColumnPos(23)
'			C2_LotNo			= iCurColumnPos(24)
'			C2_LotSeq		= iCurColumnPos(25)
'			C2_RetItemFlag	= iCurColumnPos(26)
'			C2_RetType		= iCurColumnPos(27)
'			C2_RetTypeNm		= iCurColumnPos(28)
'			C2_Spec			= iCurColumnPos(29)
'			C2_Remark		= iCurColumnPos(30)
    End Select    
End Sub

'========================================
Sub Form_Load()
	Call LoadInfTB19029
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
	Call ggoOper.LockField(Document, "N")						<% '⊙: Lock  Suitable  Field %>
	Call InitSpreadSheet("A")'20080220::hanc
	Call InitSpreadSheet("B")'20080220::hanc
	Call SetDefaultVal
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")

	txtTrackingNo.style.display = "none"

	Call InitVariables
	Call DbQuery()
	Call GetAllocInvFlag()
End Sub

'========================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'========================================
Function vspdData1_DblClick(ByVal Col, ByVal Row)
	If vspdData1.ActiveRow > 0 Then	Call OKClick
End Function

'==========================================
Sub vspdData1_Click(ByVal Col , ByVal Row)
	
	gMouseClickStatus = "SPC"					'SpreadSheet 대상명이 vspdData1일경우 
	Set gActiveSpdSheet = vspdData1
    Call SetPopupMenuItemInf("0000111111")

    If vspdData1.MaxRows <= 0 Then Exit Sub
   	    
    If Row = 0 Then
		vspdData1.OperationMode = 0
        ggoSpread.Source = vspdData1
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
	Else
		vspdData1.OperationMode = 5
        Call DbPreDtlQuery()    '2008-06-10 5:06오후 :: hanc
    End If
End Sub


'========================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub


'========================================
Sub vspdData1_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub


'========================================
Sub vspdData1_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'========================================
Function vspdData1_KeyPress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 13 And vspdData1.ActiveRow > 0 Then    'Frm1없으면 frm1삭제 
        Call OKClick()
     ElseIf KeyAscii = 27 Then
        Call CancelClick()
     End If
End Function

'========================================
Sub vspdData1_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	If OldLeft <> NewLeft Then Exit Sub

	If vspdData1.MaxRows < NewTop + VisibleRowCnt(vspdData1,NewTop) Then	    '☜: 재쿼리 체크	
		If CheckRunningBizProcess Then Exit Sub
		If lgStrPrevKey <> "" Then Call DbQuery
	End If		 

End Sub

'=======================================================
Sub txtFromDt_DblClick(Button)
	If Button = 1 Then
		txtFromDt.Action = 7
		Call SetFocusToDocument("P")
		txtFromDt.Focus
	End If
End Sub

'=======================================================
Sub txtToDt_DblClick(Button)
	If Button = 1 Then
		txtToDt.Action = 7
		Call SetFocusToDocument("P")
		txtToDt.Focus
	End If
End Sub

'=======================================================
Sub txtFromDt_Keypress(KeyAscii)
    On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call FncQuery()
    End if
End Sub

'=======================================================
Sub txtToDt_Keypress(KeyAscii)
    On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call FncQuery()
    End if
End Sub

'=====================================================
'20080220::hanc Function FncQuery() 
'20080220::hanc     
'20080220::hanc     FncQuery = False                                                        
'20080220::hanc     
'20080220::hanc     Err.Clear                                                               
'20080220::hanc 
'20080220::hanc 	If Not chkField(Document, "1") Then Exit Function
'20080220::hanc 
'20080220::hanc 	If ValidDateCheck(txtFromDt, txtToDt) = False Then Exit Function
'20080220::hanc 
'20080220::hanc 	Call ggoOper.ClearField(Document, "2")
'20080220::hanc 	Call InitVariables
'20080220::hanc 
'20080220::hanc     Call DbQuery
'20080220::hanc 
'20080220::hanc     FncQuery = True																
'20080220::hanc         
'20080220::hanc End Function

'20080220::hanc
Function FncQuery()
	FncQuery = False
		If vspddata1.MaxRows = 0 Then
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Function
			End If
		Else
			Call SetActiveCell(vspdData1,1,1,"P","X","X")
			Set gActiveElement = document.activeElement
			Call DbPreDtlQuery()
		End If
	FncQuery = False
End Function
'20080220::hanc========================================================================================
Function DbPreDtlQuery()											'☆: 조회 성공후 실행로직 
	lgStrPrevKey2 = ""   '2008-06-10 5:00오후 :: hanc
'	lgStrPrevKey4 = ""
'  	lgStrPrevKey5 = ""
	vspdData2.MaxRows = 0
	
	If DbDtlQuery = False Then	
		Call RestoreToolBar()
		Exit Function
	End If	
End Function

'20080220::hanc========================================================================================
Function DbDtlQuery() 

    Dim strVal
    Dim lngRows
	DbDtlQuery = False   
	vspdData1.Row = vspdData1.ActiveRow

	If LayerShowHide(1) = False Then
		Exit Function
	End If

    IF lgDn_Type = "I33" THEN
        lgFLAG = "B"
    ELSEIF lgDn_Type = "I35" THEN
        lgFLAG = "B"
    ELSE
        lgFLAG = "A"
    END IF

	'2009.09.04  거래처창고와 제품창고를 사용자가 선택 가능하게 변경
	If rdoSLFlg1.checked = True Then
		txtSLRadio.value = "1"
	ElseIf rdoSLFlg2.checked = True Then
		txtSLRadio.value = "0"
	End If		

 
 	If lgIntFlgMode = PopupParent.OPMD_UMODE Then
 		strVal = BIZ_PGM_QRY_ID2 & "?txtMode=" & PopupParent.UID_M0001					
 		strVal = strVal & "&txtFromDt=" & Trim(HFromDt.value)				
 		strVal = strVal & "&txtToDt=" & Trim(HToDt.value)
 		strVal = strVal & "&txtSlCd=" & Trim(HSlCd.value)		
 '		strVal = strVal & "&txtSoNo=" & Trim(HSoNo.value)
 		strVal = strVal & "&txtTrackingNo=" & Trim(HTrackingNo.value)
 	Else
 		strVal = BIZ_PGM_QRY_ID2 & "?txtMode=" & PopupParent.UID_M0001					
 		strVal = strVal & "&txtFromDt=" & Trim(txtFromDt.Text)				
 		strVal = strVal & "&txtToDt=" & Trim(txtToDt.Text)
 		strVal = strVal & "&txtSlCd=" & Trim(txtSlCd.value)		
 		strVal = strVal & "&txtTrackingNo=" & Trim(txtTrackingNo.value)
' 		strVal = strVal & "&txtSoNo=" & Trim(txtSoNo.value)
 	End If

 	
    vspdData1.Col = C_SoNo
	strVal = strVal & "&txtSoNo=" & Trim(vspdData1.Text)

 	strVal = strVal & "&txtPlantCd=" & Trim(txtHPlantCd.value)
 	strVal = strVal & "&txtMovType=" & Trim(txtHMovType.value)
 	strVal = strVal & "&txtSoType=" & Trim(txtHSoType.value)		
    vspdData1.Col = C_ItemCd
    strVal = strVal & "&txtItemCd=" & Trim(vspdData1.Text)
    vspdData1.Col = C_ShipToParty
	strVal = strVal & "&txtShipToParty=" & Trim(vspdData1.Text)
	strVal = strVal & "&txtDnType=" & lgDn_Type                 ' 20080225::hanc::출하형태
	strVal = strVal & "&txtFlag=" & lgFLAG                 ' 20080225::hanc::lgFLAG
 	strVal = strVal & "&txtHRetFlag=" & Trim(txtHRetFlag.value)

	'2009.09.04  거래처창고와 제품창고를 사용자가 선택 가능하게 변경
 	strVal = strVal & "&txtSLRadio=" & Trim(txtSLRadio.value) 

 	strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
 	strVal = strVal & "&txtLastRow=" & vspdData2.MaxRows


 	Call RunMyBizASP(MyBizASP, strVal)									

    DbDtlQuery = True

End Function

Function DbDtlQueryOk()												'☆: 조회 성공후 실행로직 

End Function

'=====================================================
Sub FncSplitColumn()
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    
End Sub

'=====================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'=====================================================
'20080220::hanc Sub PopRestoreSpreadColumnInf()
'20080220::hanc     ggoSpread.Source = gActiveSpdSheet
'20080220::hanc     Call ggoSpread.RestoreSpreadInf()
'20080220::hanc     Call InitSpreadSheet()      
'20080220::hanc     
'20080220::hanc 	Call ggoSpread.ReOrderingSpreadData()
'20080220::hanc End Sub

Sub PopRestoreSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.RestoreSpreadInf()
	Call InitSpreadSheet(gActiveSpdSheet.Id)

    If gActiveSpdSheet.Id = "A" Then
		ggoSpread.Source = vspdData1
		Call InitComboBox()
		Call ggoSpread.ReOrderingSpreadData()
		Call InitData(1,1)
	Else
		ggoSpread.Source = vspdData2
		Call ggoSpread.ReOrderingSpreadData()
	End If
End Sub

'=====================================================
Function DbQuery()

	Err.Clear															

	DbQuery = False														

	If LayerShowHide(1) = False Then
		Exit Function
	End If

	Dim strVal



	If lgIntFlgMode = PopupParent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001					
		strVal = strVal & "&txtFromDt=" & Trim(HFromDt.value)				
		strVal = strVal & "&txtToDt=" & Trim(HToDt.value)
		strVal = strVal & "&txtShipToParty=" & Trim(HShipToParty.value)
		strVal = strVal & "&txtSlCd=" & Trim(HSlCd.value)		
		strVal = strVal & "&txtSoNo=" & Trim(HSoNo.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(HTrackingNo.value)
	Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001					
		strVal = strVal & "&txtFromDt=" & Trim(txtFromDt.Text)				
		strVal = strVal & "&txtToDt=" & Trim(txtToDt.Text)
		strVal = strVal & "&txtShipToParty=" & Trim(txtShipToParty.value)
		strVal = strVal & "&txtSlCd=" & Trim(txtSlCd.value)		
		strVal = strVal & "&txtItemCd=" & Trim(txtItemCd.value)
		strVal = strVal & "&txtSoNo=" & Trim(txtSoNo.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(txtTrackingNo.value)
	End If
	strVal = strVal & "&txtPlantCd=" & Trim(txtHPlantCd.value)
	strVal = strVal & "&txtMovType=" & Trim(txtHMovType.value)
	strVal = strVal & "&txtSoType=" & Trim(txtHSoType.value)		
	strVal = strVal & "&txtHRetFlag=" & Trim(txtHRetFlag.value)

	strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	strVal = strVal & "&txtLastRow=" & vspdData1.MaxRows

	Call RunMyBizASP(MyBizASP, strVal)									
	DbQuery = True														
End Function

'=====================================================
Function DbQueryOk()
	If vspdData1.MaxRows > 0 Then
		If lgIntFlgMode <> PopupParent.OPMD_UMODE Then
			lgIntFlgMode = PopupParent.OPMD_UMODE
			vspdData1.Row = 1
			vspdData1.SelModeSelected = True		
		End If
		vspdData1.Focus
	Else
		Call SetFocusToDocument("P")
		txtFromDt.Focus
	End If

    '20080220::hanc::begin
    Call SetActiveCell(vspdData1,1,1,"P","X","X")
	Set gActiveElement = document.activeElement

	vspdData2.MaxRows = 0
	If DbDtlQuery = False Then	
		Call RestoreToolBar()
		Exit Function
	End If	
	lgOldRow = 1
	vspdData1.Focus
    '20080220::hanc::end
	
End Function

' 재고할당 여부를 Fetch한다.
'=========================================
Sub GetAllocInvFlag()
	Dim iStrSelectList, iStrFromList, iStrWhereList, iStrRs
	Dim iArrRs 

	iStrSelectList = "REFERENCE"
	iStrFromList = "dbo.B_CONFIGURATION"
	iStrWhereList = "MAJOR_CD = " & FilterVar("S0017", "''", "S") & " AND MINOR_CD = " & FilterVar("A", "''", "S") & "  AND SEQ_NO = 1 "

	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList, iStrRs) Then
		iArrRs = Split(iStrRs, Chr(11))
		lgStrAllocInvFlag = iArrRs(1)
	Else
		err.Clear
		lgStrAllocInvFlag = "N"
	End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY SCROLL=NO TABINDEX="-1">

<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS=TD5>출하예정일</TD>
						<TD CLASS=TD6>
							<TABLE CELLSPACING=0 CELLPADDING=0>
								<TR>
									<TD>
										<script language =javascript src='./js/s3112aa1_ko441_fpDateTime1_txtFromDt.js'></script>
									</TD>
									<TD>
										&nbsp;~&nbsp;
									</TD>
									<TD>
										<script language =javascript src='./js/s3112aa1_ko441_fpDateTime2_txtToDt.js'></script>
									</TD>
								</TR>
							</TABLE>
						</TD>								
						<TD CLASS="TD5" NOWRAP>납품처</TD>						
						<TD CLASS="TD6"><INPUT NAME="txtShipToParty" ALT="납품처" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDnRef" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup C_PopShipToParty">&nbsp;<INPUT TYPE=TEXT NAME="txtShipToPartyNm" SIZE=25 TAG="14" ALT="납품처명"></TD> 
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>창고</TD>
						<TD CLASS="TD6"><INPUT NAME="txtSlCd" ALT="창고" TYPE="Text" MAXLENGTH=18 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDnRef" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup C_PopSlCd">&nbsp;<INPUT NAME="txtSlNm" TYPE="Text" SIZE=25 tag="14" Alt="창고명"></TD>
						<TD CLASS="TD5" NOWRAP>품목</TD>
						<TD CLASS="TD6"><INPUT NAME="txtItemCd" ALT="품목" TYPE="Text" MAXLENGTH=18 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopItemCd">&nbsp;<INPUT NAME="txtItemNm" TYPE="Text" SIZE=25 tag="14" Alt="품목명"></TD>
					</TR>

				<!--
					<TR>
						<TD CLASS=TD5>수주번호</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtSoNo" ALT="수주번호" SIZE=25 MAXLENGTH=18 TAG="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDnRef" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopSoNo"></TD>
						<TD CLASS=TD5 NOWRAP style="display:none">Tracking No.</TD>
						<TD CLASS=TD6 NOWRAP style="display:none"><INPUT NAME="txtTrackingNo" ALT="Tracking 번호" TYPE=TEXT MAXLENGTH=25 SIZE=25 TAG="11XXXU" TABINDEX=-1><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenTrackingNo()"></TD>
					</TR>
				-->

					<TR>
						<TD CLASS=TD5>수주번호</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtSoNo" ALT="수주번호" SIZE=25 MAXLENGTH=18 TAG="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDnRef" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopSoNo"></TD>
						<TD CLASS=TD5 NOWRAP>창고유형</TD>
						<TD CLASS=TD6 >
							<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoDNFlg" TAG="11X" VALUE="1" ID="rdoSLFlg1" CHECKED ><LABEL FOR="rdoDNFlg1">거래처창고</LABEL>&nbsp;&nbsp;&nbsp;
							<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoDNFlg" TAG="11X" VALUE="0" ID="rdoSlFlg2"><LABEL FOR="rdoDNFlg3">제품창고</LABEL>			

							<INPUT NAME="txtTrackingNo" ALT="Tracking 번호" TYPE=TEXT MAXLENGTH=25 SIZE=25 TAG="11XXXU" TABINDEX=-1>
						</TD>
					</TR>

				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=* valign=top>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT="100%">
						<script language =javascript src='./js/s3112aa1_ko441_vaSpread_vspdData1.js'></script>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=* valign=top>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT="100%">
						<script language =javascript src='./js/s3112aa1_ko441_vaSpread_vspdData2.js'></script>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
    <TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
							                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=bizsize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=bizsize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX ="-1"></IFRAME></TD>
		<!--<TD WIDTH=100% HEIGHT=1000><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=1000 FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX ="-1"></IFRAME></TD>-->
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtHRetFlag" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHMovType" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHSoType" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHPlantCd" tag="14">

<INPUT TYPE=HIDDEN NAME="txtSLRadio" tag="14">

<INPUT TYPE=HIDDEN NAME="HFromDt" tag="24">
<INPUT TYPE=HIDDEN NAME="HToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="HShipToParty" tag="24">
<INPUT TYPE=HIDDEN NAME="HSlCd" tag="24">
<INPUT TYPE=HIDDEN NAME="HItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="HTrackingNo" tag="24">
<INPUT TYPE=HIDDEN NAME="HSoNo" tag="24">

<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
