<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : ����																		*
'*  2. Function Name        : ���ϰ���																	*
'*  3. Program ID           : s3112xa1_ko441     														*
'*  4. Program Name         : MES�������																*
'*  5. Program Desc         : �������_��ǰ����� ���� MES������� 	                                    *
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2008/02/21																*
'*  8. Modified date(Last)  : 2008/02/21																*
'*  9. Modifier (First)     : HAN cheol 																*
'* 10. Modifier (Last)      : HAN cheol     															*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"									*
'*                            this mark(��) Means that "may  change"									*
'*                            this mark(��) Means that "must change"									*
'* 13. History              :                                       									*
'********************************************************************************************************
%>
<HTML>
<HEAD>
<TITLE>MES�������</TITLE>
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
Dim arrParent, lgPlantCD, lgPlantNM, lgDn_Type, lgFLAG
ArrParent = window.dialogArguments
Set PopupParent  = arrParent(0)
    lgPlantCD	= arrParent(1)
    lgPlantNM	= arrParent(2)
    lgDn_Type	= arrParent(3)      '��������

top.document.title = PopupParent.gActivePRAspName

Const BIZ_PGM_QRY_ID = "s3112xb1_ko441.asp"			

' Popup Index
Const C_PopShipToParty	= 1			' ��ǰó 
Const C_PopSlCd			= 2			' â�� 
Const C_PopItemCd		= 3			' ǰ���ڵ� 
Const C_PopSoNo			= 4			' S/O ��ȣ 
Const C_PopTrackingNo	= 5			' Tracking ��ȣ 


Dim iDBSYSDate
Dim EndDate, StartDate
Dim IsOpenPop      ' Popup

iDBSYSDate = "<%=GetSvrDate%>"
'------ ��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ ------
EndDate = UniConvDateAToB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)
'------ ��: �ʱ�ȭ�鿡 �ѷ����� ���� ��¥ ------
StartDate = UNIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)

'��: Spread Sheet �� Columns �ε��� 
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
'Dim C_Spec

'20080223::hanc
Dim C_OUT_NO
Dim C_SHIP_TO_PARTY
Dim C_SHIP_TO_PARTY_NM
Dim C_ITEM_CD
Dim C_PLANT_CD
Dim C_OUT_TYPE
Dim C_OUT_TYPE_SUB
Dim C_OUT_TYPE_NM
Dim C_GI_QTY
Dim C_GI_UNIT
Dim C_LOT_NO
Dim C_LOT_SEQ
Dim C_ACTUAL_GI_DT
Dim C_SL_CD
Dim C_ITEM_NM
Dim C_SPEC
'2008-04-16 8:38���� :: hanc
Dim C_VAT_TYPE
Dim C_VAT_TYPE_NM
Dim C_VAT_INC_FLAG
Dim C_VAT_INC_FLAG_NM
Dim C_TRANS_TIME
Dim C_CREATE_TYPE
Dim C_PRICE
Dim C_CUST_LOT_NO


'========================================
<!-- #Include file="../../inc/lgvariables.inc" -->

Dim arrReturn
Dim gblnWinEvent
Dim lgStrAllocInvFlag		' ����Ҵ� ��뿩�� 

'========================================
Function InitVariables()
	lgIntGrpCount = 0										
	lgStrPrevKey = ""
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

'	txtSoNo.value = arrTemp(0)
'	txtToDt.Text = arrTemp(1)
'	txtHMovType.value = arrTemp(2)
'	txtShipToParty.value = arrTemp(3)
'	txtShipToPartyNm.value = arrTemp(4)
'	txtHSoType.value = arrTemp(5)
'	txtHRetFlag.value = arrTemp(6)
'	txtHPlantCd.value = arrTemp(7)

	txtFromDt.text  = StartDate
	txtToDt.text    = EndDate
	txtPlant.value  =   lgPlantCD '20080225::hanc
	txtPlantNM.value  =   lgPlantNM '20080225::hanc


	' ���ֹ�ȣ ������ ��� 
'	If Len(Trim(txtSoNo.value)) Then Call ggoOper.SetReqAttr(txtSoNo, "Q")

End Sub

'=========================================
Sub initSpreadPosVariables()
'	C_PlannedGIDt	= 1
'	C_ItemCd		= 2
'	C_ItemNm		= 3
'	C_Qty			= 4
'	C_BonusQty		= 5
'	C_Unit			= 6
'	C_OnStkQty		= 7
'	C_BasicUnit		= 8
'	C_SoNo			= 9
'	C_SoSeq			= 10
'	C_SoSchdNo		= 11
'	C_TrackingNo	= 12
'	C_ShipToParty	= 13
'	C_ShipToPartyNm = 14
'	C_PlantCd		= 15
'	C_PlantNm		= 16
'	C_SlCd			= 17
'	C_SlNm			= 18
'	C_TolMoreQty	= 19
'	C_TolLessQty	= 20
'	C_LcNo			= 21
'	C_LcSeq			= 22
'	C_LotFlag		= 23
'	C_LotNo			= 24
'	C_LotSeq		= 25
'	C_RetItemFlag	= 26
'	C_RetType		= 27
'	C_RetTypeNm		= 28
'	C_Spec			= 29
'	C_Remark		= 30
        	C_OUT_NO                 = 1
        	C_SHIP_TO_PARTY          = 2
        	C_SHIP_TO_PARTY_NM       = 3
        	C_ITEM_CD                = 4
        	C_PLANT_CD               = 5
        	C_OUT_TYPE               = 6
            C_OUT_TYPE_SUB           = 7
        	C_OUT_TYPE_NM            = 8
        	C_GI_QTY                 = 9
        	C_GI_UNIT                = 10
        	C_LOT_NO                 = 11
        	C_LOT_SEQ                = 12
            C_ACTUAL_GI_DT           = 13
            C_SL_CD           = 14
            C_ITEM_NM           = 15
            C_SPEC           = 16
            C_VAT_TYPE          = 17
            C_VAT_TYPE_NM       = 18
            C_VAT_INC_FLAG      = 19
            C_VAT_INC_FLAG_NM   = 10
            C_TRANS_TIME        = 21
            C_CREATE_TYPE       = 22

            C_PRICE          = 23
            C_CUST_LOT_NO		= 24

End Sub

'=====================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "RA") %>
End Sub

'========================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()    
	
	With ggoSpread

		.Source = vspdData
		.Spreadinit "V20030901",,PopupParent.gAllowDragDropSpread    

		vspdData.MaxCols = C_CUST_LOT_NO + 1
		vspdData.MaxRows = 0

		vspdData.ReDraw = False

		Call GetSpreadColumnPos("A")

		Call AppendNumberPlace("7","5","0")

    		.SSSetEdit		C_OUT_NO                 , "����(P/L NO)", 20, 0
    		.SSSetEdit		C_SHIP_TO_PARTY          , "��ǰó�ڵ�", 20, 0
    		.SSSetEdit		C_SHIP_TO_PARTY_NM       , "��ǰó��", 20, 0
    		.SSSetEdit		C_ITEM_CD                , "ǰ���ڵ�", 20, 0
    		.SSSetEdit		C_PLANT_CD               , "�����ڵ�", 20, 0
    		.SSSetEdit		C_OUT_TYPE               , "MES����TYPE", 20, 0
    		.SSSetEdit		C_OUT_TYPE_SUB               , "MES����TYPE_SUB", 20, 0
    		.SSSetEdit		C_OUT_TYPE_NM            , "MES����TYPE��", 20, 0
    		.SSSetFloat		C_GI_QTY                 , "���" ,15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
    		.SSSetEdit		C_GI_UNIT                , "������", 20, 0
    		.SSSetEdit		C_LOT_NO                 , "LOT��ȣ", 20, 0
    		.SSSetEdit		C_LOT_SEQ                , "LOT����", 20, 0
    		.SSSetDate		C_ACTUAL_GI_DT           , "�����",15,2,PopupParent.gDateFormat
    		.SSSetEdit		C_SL_CD                  , "SL_CD", 20, 0
    		.SSSetEdit		C_ITEM_NM                , "ǰ���", 20, 0
    		.SSSetEdit		C_SPEC                   , "�԰�", 20, 0
    		.SSSetEdit		C_VAT_TYPE_NM       , "VAT_TYPENM", 20, 0
    		.SSSetEdit		C_VAT_INC_FLAG      , "VAT_INC_FLAG", 20, 0
    		.SSSetEdit		C_VAT_INC_FLAG_NM   , "VAT_INC_FLAG_NM", 20, 0
    		.SSSetEdit		C_TRANS_TIME        , "C_TRANS_TIME", 20, 0
    		.SSSetEdit		C_CREATE_TYPE       , "C_CREATE_TYPE", 20, 0
    		.SSSetFloat		C_PRICE                 , "�ܰ�" ,15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
    		.SSSetEdit		C_CUST_LOT_NO       , "C_CUST_LOT_NO", 20, 0

    		'Call .SSSetColHidden(C_SL_CD,C_PRICE,True)
    		'Call .SSSetColHidden(C_VAT_TYPE          , C_VAT_TYPE         ,True)   
    		'Call .SSSetColHidden(C_VAT_TYPE_NM       , C_VAT_TYPE_NM      ,True)
    		'Call .SSSetColHidden(C_VAT_INC_FLAG      , C_VAT_INC_FLAG     ,True)
    		'Call .SSSetColHidden(C_VAT_INC_FLAG_NM   , C_VAT_INC_FLAG_NM  ,True)
    		'Call .SSSetColHidden(C_TRANS_TIME   , C_TRANS_TIME  ,True)
    		'Call .SSSetColHidden(C_CREATE_TYPE   , C_CREATE_TYPE  ,True)
    		'Call .SSSetColHidden(C_CUST_LOT_NO   , C_CUST_LOT_NO  ,True)


'		.SSSetFloat C_LotSeq,"LOT NO ����" ,15,"7", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec,,,"Z"  
'
'		.SSSetDate		C_PlannedGIDt, "���Ͽ�����",15,2,PopupParent.gDateFormat
'		.SSSetEdit		C_ItemCd, "ǰ��", 20, 0
'		.SSSetEdit		C_ItemNm, "ǰ���", 40, 0
'		.SSSetFloat		C_Qty,"��������" ,15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
'		.SSSetFloat		C_BonusQty,"����������" ,15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
'		.SSSetEdit		C_Unit, "����", 10, 2
'		.SSSetFloat		C_OnStkQty,"���" ,15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
'		.SSSetEdit		C_BasicUnit, "������", 10, 2
'		.SSSetEdit		C_SoNo, "���ֹ�ȣ", 18, 0
'		.SSSetFloat		C_SoSeq,"���ּ���" ,15,"7", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec,,,"Z"  
'		.SSSetFloat		C_SoSchdNo,"��ǰ����" ,15,"7", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec,,,"Z"  
'		.SSSetEdit		C_TrackingNo, "Tracking No", 18, 0
'		.SSSetEdit		C_ShipToParty, "��ǰó", 15, 0
'		.SSSetEdit		C_ShipToPartyNm, "��ǰó��", 25, 0
'		.SSSetEdit		C_PlantCd, "����", 15, 0
'		.SSSetEdit		C_PlantNm, "�����", 20, 0
'		.SSSetEdit		C_SlCd, "â��", 15, 0
'		.SSSetEdit		C_SlNm, "â���", 20, 0
'		.SSSetFloat		C_TolMoreQty,"��������뷮(+)" ,20,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
'		.SSSetFloat		C_TolLessQty,"��������뷮(-)" ,20,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
'		.SSSetEdit		C_LcNo, "L/C��ȣ", 18, 0
'		.SSSetFloat		C_LcSeq,"L/C����" ,15,"7", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec,,,"Z"  
'							   'ColumnPosition		Header              Width	Align(0:L,1:R,2:C)  Row   Length  CharCase(0:L,1:N,2:U)
'		.SSSetEdit		C_LotFlag, "LOT��������", 15, 2
'		.SSSetEdit		C_LotNo, "LOT NO", 18, 0
'		.SSSetFloat		C_LotSeq,"LOT NO ����" ,15,"7", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec,,,"Z"  
'		.SSSetEdit		C_RetItemFlag, "��ǰ����", 10, 2
'		.SSSetEdit		C_RetType, "��ǰ����", 15, 0
'		.SSSetEdit		C_RetTypeNm, "��ǰ������", 20, 0
'		.SSSetEdit		C_Spec, "�԰�", 30, 0
'		.SSSetEdit		C_Remark,			"���",			30,	0,					,	  120
'
'		Call .SSSetColHidden(C_PlantCd,C_PlantNm,True)
'		Call .SSSetColHidden(vspdData.MaxCols,vspdData.MaxCols,True)
			
		vspdData.ReDraw = True
	End With

	Call SetSpreadLock

End Sub

'========================================
Sub SetSpreadLock()
	ggoSpread.SpreadLockWithOddEvenRowColor()
	vspdData.OperationMode = 5	'Multi Select Mode
End Sub

'========================================
Function OKClick()
	on error resume next
	
	Dim intColCnt, intRowCnt, intInsRow, iLngSelectedRows

	With vspdData
		iLngSelectedRows = .SelModeSelCount
		' ��ü���ý� 
		If iLngSelectedRows = -1 Then
			iLngSelectedRows = .MaxRows
		End If

		If iLngSelectedRows > 0 Then 
			intInsRow = 0

			Redim arrReturn(iLngSelectedRows, .MaxCols)

			For intRowCnt = 1 To .MaxRows

				.Row = intRowCnt

				If .SelModeSelected Then
					.Col = C_SL_CD	        : arrReturn(intInsRow, 0) = .Text
					.Col = C_ITEM_CD        : arrReturn(intInsRow, 1) = .Text
					.Col = C_ITEM_NM        : arrReturn(intInsRow, 2) = .Text
					.Col = C_SPEC           : arrReturn(intInsRow, 3) = .Text
					.Col = C_GI_UNIT        : arrReturn(intInsRow, 4) = .Text
					.Col = C_GI_QTY         : arrReturn(intInsRow, 5) = .Text
					.Col = C_PRICE          : arrReturn(intInsRow, 6) = .Text
					.Col = C_LOT_NO         : arrReturn(intInsRow, 7) = .Text
					.Col = C_LOT_SEQ        : arrReturn(intInsRow, 8) = .Text
					.Col = C_OUT_NO         : arrReturn(intInsRow, 9) = .Text       '20080226::hanc
					.Col = C_SHIP_TO_PARTY         : arrReturn(intInsRow, 10) = .Text       '20080416::hanc
					.Col = C_SHIP_TO_PARTY_NM      : arrReturn(intInsRow, 11) = .Text       '20080416::hanc
					.Col = C_VAT_TYPE              : arrReturn(intInsRow, 12) = .Text       '20080416::hanc
					.Col = C_VAT_TYPE_NM           : arrReturn(intInsRow, 13) = .Text       '20080416::hanc
					.Col = C_VAT_INC_FLAG          : arrReturn(intInsRow, 14) = .Text       '20080416::hanc
					.Col = C_VAT_INC_FLAG_NM       : arrReturn(intInsRow, 15) = .Text       '20080416::hanc
					.Col = C_TRANS_TIME            : arrReturn(intInsRow, 16) = .Text       '20080416::hanc
					.Col = C_CREATE_TYPE           : arrReturn(intInsRow, 17) = .Text       '20080416::hanc
					.Col = C_OUT_TYPE              : arrReturn(intInsRow, 18) = .Text       '20080417::hanc
					.Col = C_OUT_TYPE_SUB              : arrReturn(intInsRow, 19) = .Text       '20080417::hanc
					.Col = C_CUST_LOT_NO              : arrReturn(intInsRow, 20) = .Text       '20080417::hanc

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

					intInsRow = intInsRow + 1

				End IF
			Next
		End if			
	End With

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
		'��ǰó 
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
			iArrHeader(2) = "����"
			iArrHeader(3) = "������"

			txtShipToParty.focus

		'â�� 
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

		' ǰ�� 
		Case C_PopItemCd
			OpenConPopup = OpenConItemPopup(C_PopItemCd, txtItemCd.value)
			txtItemCd.focus
			Exit Function

		' S/O ��ȣ 
		Case C_PopSoNo
			iArrParam(1) = "S_SO_HDR SH, B_BIZ_PARTNER SP, B_SALES_GRP SG"
			iArrParam(2) = Trim(txtSoNo.value)
			iArrParam(3) = ""
	
			' ����Ҵ��� ��뿩�� 
			If lgStrAllocInvFlag = "N" Then
				iArrParam(4) = "SH.SOLD_TO_PARTY = SP.BP_CD AND SH.SALES_GRP = SG.SALES_GRP AND SH.CFM_FLAG = " & FilterVar("Y", "''", "S") & "  AND SH.REL_DN_FLAG = " & FilterVar("Y", "''", "S") & "  AND EXISTS (SELECT * FROM S_SO_DTL SD WHERE SD.SO_NO = SH.SO_NO AND SD.SO_QTY + SD.BONUS_QTY > SD.REQ_QTY + SD.REQ_BONUS_QTY "
			Else
				iArrParam(4) = "SH.SOLD_TO_PARTY = SP.BP_CD AND SH.SALES_GRP = SG.SALES_GRP AND SH.CFM_FLAG = " & FilterVar("Y", "''", "S") & "  AND SH.REL_DN_FLAG = " & FilterVar("Y", "''", "S") & "  AND EXISTS (SELECT * FROM S_SO_SCHD SC INNER JOIN dbo.S_SO_DTL SD ON (SD.SO_NO = SC.SO_NO AND SD.SO_SEQ = SC.SO_SEQ) WHERE SC.SO_NO = SH.SO_NO AND SC.ALLC_QTY + SC.ALLC_BONUS_QTY > SC.REQ_QTY + SC.REQ_BONUS_QTY "
			End If
			' ���� 
			iArrParam(4) = iArrParam(4) & " AND SD.PLANT_CD =  " & FilterVar(txtHPlantCd.value, "''", "S") & ""
			' ��ǰó 
			If Trim(txtShipToParty.value) = "" Then
				iArrParam(4) = iArrParam(4) & ")"
			Else
				iArrParam(4) = iArrParam(4) & " AND SD.SHIP_TO_PARTY =  " & FilterVar(txtShipToParty.value, "''", "S") & ")"
			End If
	
			iArrParam(5) = "���ֹ�ȣ"

			iArrField(0) = "ED12" & PopupParent.gColSep & "SH.SO_NO"
			iArrField(1) = "ED10" & PopupParent.gColSep & "SH.SOLD_TO_PARTY"
			iArrField(2) = "ED20" & PopupParent.gColSep & "SP.BP_NM"
			iArrField(3) = "DD10" & PopupParent.gColSep & "SH.SO_DT"
			iArrField(4) = "ED15" & PopupParent.gColSep & "SG.SALES_GRP_NM"
			iArrField(5) = "ED10" & PopupParent.gColSep & "SH.PAY_METH"
						
			iArrHeader(0) = "���ֹ�ȣ"
			iArrHeader(1) = "�ֹ�ó"
			iArrHeader(2) = "�ֹ�ó��"
			iArrHeader(3) = "������"
			iArrHeader(4) = "�����׷��"
			iArrHeader(5) = "�������"
			
			txtSoNo.focus
	End Select
 
	iArrParam(0) = iArrHeader(0)							' �˾� Title
	iArrParam(5) = iArrHeader(0)							' ��ȸ���� ��Ī 

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
		' ��ǰó 
		Case C_PopShipToParty
			txtShipToParty.value = pvArrRet(0)
			txtShipToPartyNm.value = pvArrRet(1) 

		' â�� 
		Case C_PopSlCd
			txtSlCd.value = pvArrRet(0) 
			txtSlNm.value = pvArrRet(1)   

		' ǰ�� 
		Case C_PopItemCd
			txtItemCd.value = pvArrRet(0) 
			txtItemNm.value = pvArrRet(1)   

		' S/O ��ȣ 
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
            ggoSpread.Source = vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)


            C_OUT_NO                 = iCurColumnPos(1)
            C_SHIP_TO_PARTY          = iCurColumnPos(2)
            C_SHIP_TO_PARTY_NM       = iCurColumnPos(3)
            C_ITEM_CD                = iCurColumnPos(4)
            C_PLANT_CD               = iCurColumnPos(5)
            C_OUT_TYPE               = iCurColumnPos(6)
            C_OUT_TYPE_SUB           = iCurColumnPos(7) 
            C_OUT_TYPE_NM            = iCurColumnPos(8) 
            C_GI_QTY                 = iCurColumnPos(9) 
            C_GI_UNIT                = iCurColumnPos(10)
            C_LOT_NO                 = iCurColumnPos(11)
            C_LOT_SEQ                = iCurColumnPos(12)
            C_ACTUAL_GI_DT           = iCurColumnPos(13)
            C_SL_CD                  = iCurColumnPos(14)
            C_ITEM_NM                = iCurColumnPos(15)
            C_SPEC                   = iCurColumnPos(16)
            C_VAT_TYPE               = iCurColumnPos(17)
            C_VAT_TYPE_NM            = iCurColumnPos(18)
            C_VAT_INC_FLAG           = iCurColumnPos(19)
            C_VAT_INC_FLAG_NM        = iCurColumnPos(20)
            C_TRANS_TIME             = iCurColumnPos(21)
            C_CREATE_TYPE            = iCurColumnPos(22)
            C_PRICE                  = iCurColumnPos(23)
						C_CUST_LOT_NO						 = iCurColumnPos(24)
'			C_PlannedGIDt	= iCurColumnPos(1)
'			C_ItemCd		= iCurColumnPos(2)
'			C_ItemNm		= iCurColumnPos(3)
'			C_Qty			= iCurColumnPos(4)
'			C_BonusQty		= iCurColumnPos(5)
'			C_Unit			= iCurColumnPos(6)
'			C_OnStkQty		= iCurColumnPos(7)
'			C_BasicUnit		= iCurColumnPos(8)
'			C_SoNo			= iCurColumnPos(9)
'			C_SoSeq			= iCurColumnPos(10)
'			C_SoSchdNo		= iCurColumnPos(11)
'			C_TrackingNo	= iCurColumnPos(12)
'			C_ShipToParty	= iCurColumnPos(13)
'			C_ShipToPartyNm = iCurColumnPos(14)
'			C_PlantCd		= iCurColumnPos(15)
'			C_PlantNm		= iCurColumnPos(16)
'			C_SlCd			= iCurColumnPos(17)
'			C_SlNm			= iCurColumnPos(18)
'			C_TolMoreQty	= iCurColumnPos(19)
'			C_TolLessQty	= iCurColumnPos(20)
'			C_LcNo			= iCurColumnPos(21)
'			C_LcSeq			= iCurColumnPos(22)
'			C_LotFlag		= iCurColumnPos(23)
'			C_LotNo			= iCurColumnPos(24)
'			C_LotSeq		= iCurColumnPos(25)
'			C_RetItemFlag	= iCurColumnPos(26)
'			C_RetType		= iCurColumnPos(27)
'			C_RetTypeNm		= iCurColumnPos(28)
'			C_Spec			= iCurColumnPos(29)
'			C_Remark		= iCurColumnPos(30)
    End Select    
End Sub

'========================================
Sub Form_Load()
	Call LoadInfTB19029
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
	Call ggoOper.LockField(Document, "N")						<% '��: Lock  Suitable  Field %>
	Call InitSpreadSheet()
	Call SetDefaultVal
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call InitVariables
	Call DbQuery()
	Call GetAllocInvFlag()
End Sub

'========================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'========================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If vspdData.ActiveRow > 0 Then	Call OKClick
End Function

'==========================================
Sub vspdData_Click(ByVal Col , ByVal Row)
	
	gMouseClickStatus = "SPC"					'SpreadSheet ������ vspdData�ϰ�� 
	Set gActiveSpdSheet = vspdData
    Call SetPopupMenuItemInf("0000111111")

    If vspdData.MaxRows <= 0 Then Exit Sub
   	    
    If Row = 0 Then
		vspdData.OperationMode = 0
        ggoSpread.Source = vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
	Else
		vspdData.OperationMode = 5
    End If
End Sub

'========================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'========================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'========================================
Function vspdData_KeyPress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 13 And vspdData.ActiveRow > 0 Then    'Frm1������ frm1���� 
        Call OKClick()
     ElseIf KeyAscii = 27 Then
        Call CancelClick()
     End If
End Function

'========================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	If OldLeft <> NewLeft Then Exit Sub

	If vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData,NewTop) Then	    '��: ������ üũ	
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
Function FncQuery() 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               

    '20080218::hanc
    If Trim(txtPlant.Value) = "" then
       MSGBOX "������ �ʼ� �׸� �Դϴ�."
       Exit Function
    End If

	If Not chkField(Document, "1") Then Exit Function

	If ValidDateCheck(txtFromDt, txtToDt) = False Then Exit Function
	Call ggoOper.ClearField(Document, "2")
	Call InitVariables

    Call DbQuery

    FncQuery = True																
        
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
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'=====================================================
Function DbQuery()
	Err.Clear															

	DbQuery = False														

	If LayerShowHide(1) = False Then
		Exit Function
	End If

	Dim strVal

    IF lgDn_Type = "I33" THEN
        lgFLAG = "B"
    ELSEIF lgDn_Type = "I35" THEN
        lgFLAG = "B"
    ELSE
        lgFLAG = "A"
    END IF
    
	If lgIntFlgMode = PopupParent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001					
'		strVal = strVal & "&txtFromDt=" & Trim(HFromDt.value)				
'		strVal = strVal & "&txtToDt=" & Trim(HToDt.value)
		strVal = strVal & "&txtShipToParty=" & Trim(HShipToParty.value)
		strVal = strVal & "&txtSlCd=" & Trim(HSlCd.value)		
		strVal = strVal & "&txtSoNo=" & Trim(HSoNo.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(HTrackingNo.value)
	Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001					
		strVal = strVal & "&txtShipToParty=" & Trim(txtShipToParty.value)
		strVal = strVal & "&txtSlCd=" & Trim(txtSlCd.value)		
		strVal = strVal & "&txtItemCd=" & Trim(txtItemCd.value)
		strVal = strVal & "&txtSoNo=" & Trim(txtSoNo.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(txtTrackingNo.value)
	End If
	strVal = strVal & "&txtFromDt=" & Trim(txtFromDt.Text)				
	strVal = strVal & "&txtToDt=" & Trim(txtToDt.Text)
	strVal = strVal & "&txtPlantCd=" & Trim(txtPlant.value)
	strVal = strVal & "&txtDnType=" & lgDn_Type                 ' 20080225::hanc::��������
	strVal = strVal & "&txtFlag=" & lgFLAG                 ' 20080225::hanc::lgFLAG
	strVal = strVal & "&txtMovType=" & Trim(txtHMovType.value)
	strVal = strVal & "&txtSoType=" & Trim(txtHSoType.value)		
	strVal = strVal & "&txtHRetFlag=" & Trim(txtHRetFlag.value)

	strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	strVal = strVal & "&txtLastRow=" & vspdData.MaxRows

	Call RunMyBizASP(MyBizASP, strVal)									

	DbQuery = True														
End Function

'=====================================================
Function DbQueryOk()
	If vspdData.MaxRows > 0 Then
		If lgIntFlgMode <> PopupParent.OPMD_UMODE Then
			lgIntFlgMode = PopupParent.OPMD_UMODE
			vspdData.Row = 1
			vspdData.SelModeSelected = True		
		End If
		vspdData.Focus
	Else
		Call SetFocusToDocument("P")
		txtFromDt.Focus
	End If

End Function

' ����Ҵ� ���θ� Fetch�Ѵ�.
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

'20080225::HANC========================================
Function OpenPlant()
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    If txtPlant.ReadOnly = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "����"
    arrParam(1) = "B_PLANT"
    arrParam(2) = Trim(txtPlant.Value)
    arrParam(4) = ""
    arrParam(5) = "����"
     
    arrField(0) = "PLANT_CD"
    arrField(1) = "PLANT_NM"
        
    arrHeader(0) = "����"
    arrHeader(1) = "�����"

    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
                                    "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False

    txtPlant.focus
    
    If arrRet(0) <> "" Then
        Call SetPlant(arrRet)
    End If
End Function
'20080225::HANC========================================
Function SetPlant(arrRet)
    txtPlant.Value = arrRet(0)
    txtPlantNm.Value = arrRet(1)
    lgBlnFlgChgValue = True

End Function


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
						<TD CLASS=TD5>MES�����</TD>
						<TD CLASS=TD6>
							<TABLE CELLSPACING=0 CELLPADDING=0>
								<TR>
									<TD>
										<script language =javascript src='./js/s3112xa1_ko441_fpDateTime1_txtFromDt.js'></script>
									</TD>
									<TD>
										&nbsp;~&nbsp;
									</TD>
									<TD>
										<script language =javascript src='./js/s3112xa1_ko441_fpDateTime2_txtToDt.js'></script>
									</TD>
								</TR>
							</TABLE>
						</TD>								
						<TD CLASS="TD5" NOWRAP>��ǰó</TD>						
						<TD CLASS="TD6"><INPUT NAME="txtShipToParty" ALT="��ǰó" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDnRef" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup C_PopShipToParty">&nbsp;<INPUT TYPE=TEXT NAME="txtShipToPartyNm" SIZE=25 TAG="14" ALT="��ǰó��"></TD> 
					</TR>
                    <TR>
						<TD CLASS="TD5" NOWRAP></TD>
						<TD CLASS="TD6" NOWRAP></TD>
						<TD CLASS="TD5" NOWRAP>����</TD>
						<TD CLASS="TD6" NOWRAP><INPUT NAME="txtPlant" ALT="����" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="12XXXU" CLASS="required" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLCType" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenPlant()">&nbsp;<INPUT NAME="txtPlantNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
                    </TR>
					<TR STYLE="DISPLAY:NONE">
						<TD CLASS="TD5" NOWRAP>â��</TD>
						<TD CLASS="TD6"><INPUT NAME="txtSlCd" ALT="â��" TYPE="Text" MAXLENGTH=18 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDnRef" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup C_PopSlCd">&nbsp;<INPUT NAME="txtSlNm" TYPE="Text" SIZE=25 tag="14" Alt="â���"></TD>
						<TD CLASS="TD5" NOWRAP>ǰ��</TD>
						<TD CLASS="TD6"><INPUT NAME="txtItemCd" ALT="ǰ��" TYPE="Text" MAXLENGTH=18 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopItemCd">&nbsp;<INPUT NAME="txtItemNm" TYPE="Text" SIZE=25 tag="14" Alt="ǰ���"></TD>
					</TR>
					<TR STYLE="DISPLAY:NONE">
						<TD CLASS=TD5>���ֹ�ȣ</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtSoNo" ALT="���ֹ�ȣ" SIZE=25 MAXLENGTH=18 TAG="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDnRef" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopSoNo"></TD>
						<TD CLASS=TD5 NOWRAP>Tracking No.</TD>
						<TD CLASS=TD6><INPUT NAME="txtTrackingNo" ALT="Tracking ��ȣ" TYPE=TEXT MAXLENGTH=25 SIZE=25 TAG="11XXXU" TABINDEX=-1><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenTrackingNo()"></TD>
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
						<script language =javascript src='./js/s3112xa1_ko441_vaSpread_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX ="-1"></IFRAME></TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtHRetFlag" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHMovType" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHSoType" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHPlantCd" tag="14">

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
