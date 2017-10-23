<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : 영업																		*
'*  2. Function Name        : 출하관리																	*
'*  3. Program ID           : I1311xa1_ko441     														*
'*  4. Program Name         : MES출고참조																*
'*  5. Program Desc         : 국내출고_반품등록을 위한 MES출고참조 	                                    *
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2008/02/21																*
'*  8. Modified date(Last)  : 2008/02/21																*
'*  9. Modifier (First)     : HAN cheol 																*
'* 10. Modifier (Last)      : HAN cheol     															*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              :                                       									*
'********************************************************************************************************
%>
<HTML>
<HEAD>
<TITLE>MES출고참조</TITLE>
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
Dim arrParent, lgPlantCD, lgPlantNM, lgDn_Type, lgFLAG, lgMovType
ArrParent = window.dialogArguments
Set PopupParent  = arrParent(0)
    lgPlantCD	= arrParent(1)
    lgPlantNM	= arrParent(2)
    lgDn_Type	= arrParent(3)      '
    lgMovType	= arrParent(4)      'trns_type

top.document.title = PopupParent.gActivePRAspName

Const BIZ_PGM_QRY_ID = "p2250xb1_ko441.asp"			

' Popup Index
Const C_PopShipToParty	= 1			' 납품처 
Const C_PopSlCd			= 2			' 창고 
Const C_PopItemCd		= 3			' 품목코드 
Const C_PopSoNo			= 4			' S/O 번호 
Const C_PopTrackingNo	= 5			' Tracking 번호 


Dim iDBSYSDate
Dim EndDate, StartDate
Dim IsOpenPop      ' Popup

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UniConvDateAToB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UNIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)

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
'Dim C_Spec

'20080223::hanc
Dim C_OUT_NO
Dim C_SHIP_TO_PARTY
Dim C_SHIP_TO_PARTY_NM
'Dim C_ITEM_CD
Dim C_PLANT_CD
Dim C_OUT_TYPE
Dim C_OUT_TYPE_NM
Dim C_GI_QTY
Dim C_GI_UNIT
Dim C_LOT_NO
Dim C_LOT_SEQ
Dim C_ACTUAL_GI_DT
Dim C_SL_CD
'Dim C_ITEM_NM
'Dim C_SPEC
Dim C_PRICE
Dim c_version

'20080226::hanc
Dim C_ITEM_CD
Dim C_ITEM_NM
Dim C_SPEC
Dim C_PRODT_ORDER_NO
Dim C_BASIC_UNIT
Dim C_REQ_QTY
Dim C_ISSUE_QTY
Dim C_REMIND_QTY
Dim C_ITEM_SEQ
Dim C_ISSUE_REQ_NO
'20080305
Dim C_req_dt
Dim C_dept_cd
Dim C_dept_nm
Dim C_emp_no
Dim C_emp_nm
Dim C_mov_type '20080312::hanc


'========================================
<!-- #Include file="../../inc/lgvariables.inc" -->

Dim arrReturn
Dim gblnWinEvent
Dim lgStrAllocInvFlag		' 재고할당 사용여부 

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

	txtFromDt.text  = EndDate
	txtToDt.text    = EndDate
	txtPlant.value  =   lgPlantCD '20080225::hanc
	txtPlantNM.value  =   lgPlantNM '20080225::hanc


	' 수주번호 지정인 경우 
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
'            C_SL_CD           = 13
'            C_ITEM_NM           = 14
'            C_SPEC           = 15
'            C_PRICE          = 16

        	C_version               = 1       
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

		vspdData.MaxCols = c_version + 1
		vspdData.MaxRows = 0

		vspdData.ReDraw = False

		Call GetSpreadColumnPos("A")

		Call AppendNumberPlace("7","5","0")

    		.SSSetEdit		c_version               , "Ver.", 15, 0
'    		.SSSetEdit		C_ITEM_NM               , "품명", 18, 0
'    		.SSSetEdit		C_SPEC                  , "규격", 10, 0
'    		.SSSetEdit		C_PRODT_ORDER_NO        , "제조오더번호", 10, 0
'    		.SSSetEdit		C_BASIC_UNIT            , "단위", 10, 0
'    		.SSSetFloat		C_REQ_QTY               , "요청량" ,15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
'    		.SSSetFloat		C_ISSUE_QTY             , "기출고량" ,15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
'    		.SSSetFloat		C_REMIND_QTY            , "요청잔량" ,15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
'    		.SSSetEdit		C_ITEM_SEQ              , "순번", 10, 0
'    		.SSSetEdit		C_ISSUE_REQ_NO          , "불출의뢰No", 20, 0
'    		.SSSetEdit		C_req_dt                , "의뢰일", 10, 0  '20080305
'    		.SSSetEdit		C_dept_cd               , "의뢰부서", 10, 0
'    		.SSSetEdit		C_dept_nm               , "의뢰부서명", 15, 0
'    		.SSSetEdit		C_emp_no                , "의뢰담당자", 10, 0
'    		.SSSetEdit		C_mov_type              , "불출의뢰유형", 10, 0     '20080312::HANC
'    		.SSSetEdit		C_emp_nm                , "의뢰담당자명", 10, 0

'    		.SSSetEdit		C_OUT_NO                 , "출하(P/L NO)", 20, 0
'    		.SSSetEdit		C_SHIP_TO_PARTY          , "납품처코드", 20, 0
'    		.SSSetEdit		C_SHIP_TO_PARTY_NM       , "납품처명", 20, 0
'    		.SSSetEdit		C_ITEM_CD                , "품목코드", 20, 0
'    		.SSSetEdit		C_PLANT_CD               , "공장코드", 20, 0
'    		.SSSetEdit		C_OUT_TYPE               , "MES출하TYPE", 20, 0
'    		.SSSetEdit		C_OUT_TYPE_NM            , "MES출하TYPE명", 20, 0
'    		.SSSetFloat		C_GI_QTY                 , "출고량" ,15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
'    		.SSSetEdit		C_GI_UNIT                , "출고단위", 20, 0
'    		.SSSetEdit		C_LOT_NO                 , "LOT번호", 20, 0
'    		.SSSetEdit		C_LOT_SEQ                , "LOT순번", 20, 0
'    		.SSSetDate		C_ACTUAL_GI_DT           , "출고일",15,2,PopupParent.gDateFormat
'    		.SSSetEdit		C_SL_CD                  , "SL_CD", 20, 0
'    		.SSSetEdit		C_ITEM_NM                , "품목명", 20, 0
'    		.SSSetEdit		C_SPEC                   , "규격", 20, 0
'    		.SSSetFloat		C_PRICE                 , "단가" ,15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"

    		Call .SSSetColHidden(vspdData.MaxCols,vspdData.MaxCols,True)


'		.SSSetFloat C_LotSeq,"LOT NO 순번" ,15,"7", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec,,,"Z"  
'
'		.SSSetDate		C_PlannedGIDt, "출하예정일",15,2,PopupParent.gDateFormat
'		.SSSetEdit		C_ItemCd, "품목", 20, 0
'		.SSSetEdit		C_ItemNm, "품목명", 40, 0
'		.SSSetFloat		C_Qty,"미출고수량" ,15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
'		.SSSetFloat		C_BonusQty,"미출고덤수량" ,15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
'		.SSSetEdit		C_Unit, "단위", 10, 2
'		.SSSetFloat		C_OnStkQty,"재고량" ,15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
'		.SSSetEdit		C_BasicUnit, "재고단위", 10, 2
'		.SSSetEdit		C_SoNo, "수주번호", 18, 0
'		.SSSetFloat		C_SoSeq,"수주순번" ,15,"7", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec,,,"Z"  
'		.SSSetFloat		C_SoSchdNo,"납품순번" ,15,"7", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec,,,"Z"  
'		.SSSetEdit		C_TrackingNo, "Tracking No", 18, 0
'		.SSSetEdit		C_ShipToParty, "납품처", 15, 0
'		.SSSetEdit		C_ShipToPartyNm, "납품처명", 25, 0
'		.SSSetEdit		C_PlantCd, "공장", 15, 0
'		.SSSetEdit		C_PlantNm, "공장명", 20, 0
'		.SSSetEdit		C_SlCd, "창고", 15, 0
'		.SSSetEdit		C_SlNm, "창고명", 20, 0
'		.SSSetFloat		C_TolMoreQty,"과부족허용량(+)" ,20,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
'		.SSSetFloat		C_TolLessQty,"과부족허용량(-)" ,20,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
'		.SSSetEdit		C_LcNo, "L/C번호", 18, 0
'		.SSSetFloat		C_LcSeq,"L/C순번" ,15,"7", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec,,,"Z"  
'							   'ColumnPosition		Header              Width	Align(0:L,1:R,2:C)  Row   Length  CharCase(0:L,1:N,2:U)
'		.SSSetEdit		C_LotFlag, "LOT관리여부", 15, 2
'		.SSSetEdit		C_LotNo, "LOT NO", 18, 0
'		.SSSetFloat		C_LotSeq,"LOT NO 순번" ,15,"7", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec,,,"Z"  
'		.SSSetEdit		C_RetItemFlag, "반품여부", 10, 2
'		.SSSetEdit		C_RetType, "반품유형", 15, 0
'		.SSSetEdit		C_RetTypeNm, "반품유형명", 20, 0
'		.SSSetEdit		C_Spec, "규격", 30, 0
'		.SSSetEdit		C_Remark,			"비고",			30,	0,					,	  120
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
	Dim before_ISSUE_REQ_NO, curr_ISSUE_REQ_NO


	before_ISSUE_REQ_NO =   "" 
	curr_ISSUE_REQ_NO   =   "" 


	With vspdData
		iLngSelectedRows = .SelModeSelCount
		' 전체선택시 
		If iLngSelectedRows = -1 Then
			iLngSelectedRows = .MaxRows
		End If

		If iLngSelectedRows > 0 Then 
			intInsRow = 0

			Redim arrReturn(iLngSelectedRows, .MaxCols)

			For intRowCnt = 1 To .MaxRows

				.Row = intRowCnt

				If .SelModeSelected Then
					.Col = C_version        : arrReturn(intInsRow, 0) = .Text
					

            				
                    if intInsRow <> 0 then
                        if curr_ISSUE_REQ_NO <> before_ISSUE_REQ_NO then
                            call DisplayMsgBox("ZZ0006", PopupParent.VB_INFORMATION, "X", "X")
                            Exit Function
                        end if
                       
                    end if
				

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
            ggoSpread.Source = vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

            C_version               = iCurColumnPos(1)   

    End Select    
End Sub

'========================================
Sub Form_Load()
	Call LoadInfTB19029
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
	Call ggoOper.LockField(Document, "N")						<% '⊙: Lock  Suitable  Field %>
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
	
	gMouseClickStatus = "SPC"					'SpreadSheet 대상명이 vspdData일경우 
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
     If KeyAscii = 13 And vspdData.ActiveRow > 0 Then    'Frm1없으면 frm1삭제 
        Call OKClick()
     ElseIf KeyAscii = 27 Then
        Call CancelClick()
     End If
End Function

'========================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	If OldLeft <> NewLeft Then Exit Sub

	If vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData,NewTop) Then	    '☜: 재쿼리 체크	
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
       MSGBOX "공장은 필수 항목 입니다."
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


    
	If lgIntFlgMode = PopupParent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001					
		strVal = strVal & "&txtToDt=" & Trim(HToDt.value)
		strVal = strVal & "&txtShipToParty=" & Trim(HShipToParty.value)
		strVal = strVal & "&txtSlCd=" & Trim(HSlCd.value)		
		strVal = strVal & "&txtSoNo=" & Trim(HSoNo.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(HTrackingNo.value)
	Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001					
		strVal = strVal & "&txtToDt=" & Trim(txtToDt.Text)
		strVal = strVal & "&txtShipToParty=" & Trim(txtShipToParty.value)
		strVal = strVal & "&txtSlCd=" & Trim(txtSlCd.value)		
		strVal = strVal & "&txtItemCd=" & Trim(txtItemCd.value)
		strVal = strVal & "&txtSoNo=" & Trim(txtSoNo.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(txtTrackingNo.value)
	End If
	strVal = strVal & "&txtFromDt=" & Trim(txtFromDt.Text)				
	strVal = strVal & "&txtPlantCd=" & Trim(txtPlant.value)
	strVal = strVal & "&txtDnType=" & lgDn_Type                 ' 20080225::hanc
	strVal = strVal & "&txtISSUE_REQ_NO=" & Trim(txtPoNo.value)                 ' 20080225::hanc::lgFLAG
	strVal = strVal & "&txtMovType=" & lgMovType                 ' 20080226::hanc	
'	strVal = strVal & "&txtMovType=" & Trim(txtHMovType.value)
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

'20080225::HANC========================================
Function OpenPlant()
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    If txtPlant.ReadOnly = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "공장"
    arrParam(1) = "B_PLANT"
    arrParam(2) = Trim(txtPlant.Value)
    arrParam(4) = ""
    arrParam(5) = "공장"
     
    arrField(0) = "PLANT_CD"
    arrField(1) = "PLANT_NM"
        
    arrHeader(0) = "공장"
    arrHeader(1) = "공장명"

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
'20080211::hanc
Function OpenIssueReq()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Or UCase(txtPoNo.className) = UCase(popupparent.UCN_PROTECTED) Then Exit Function
	
	
	IsOpenPop = True


    arrParam(0) = "불출의뢰번호 팝업"                   ' 팝업 명칭
    arrParam(1) = "M_ISSUE_REQ_HDR_KO441"               ' TABLE 명칭
    arrParam(2) = Trim(txtPoNo.Value)
    arrParam(3) = ""                                    ' Name Cindition
    arrParam(4) = ""
    arrParam(5) = "불출의뢰번호"
    
    arrField(0) = "ISSUE_REQ_NO"                    ' Field명(0)
    arrField(1) = "REQ_DT"                     	    ' Field명(1)
    arrField(2) = "DEPT_CD"                     	' Field명(2)

    arrHeader(0) = "불출의뢰번호"                     	' Header명(0)
    arrHeader(1) = "블츨의뢰일"                  	 	' Header명(1)
    arrHeader(2) = "불출의뢰부서"                  	 	' Header명(2)


    	
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetIssueReq(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	txtPoNo.focus
	
End Function

Function SetIssueReq(byval arrRet)
	txtPoNo.value          = arrRet(0)
	
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY SCROLL=NO TABINDEX="-1">

<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR >
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR STYLE="DISPLAY:NONE">
						<TD CLASS="TD5" NOWRAP>공장</TD>
						<TD CLASS="TD6" NOWRAP><INPUT NAME="txtPlant" ALT="공장" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="12XXXU" CLASS="required" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLCType" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenPlant()">&nbsp;<INPUT NAME="txtPlantNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
						<TD CLASS="TD5" NOWRAP>불출의뢰번호</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPoNo" SIZE=29 MAXLENGTH=18 ALT="불출의뢰번호" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenIssueReq()"></TD>
					</TR>
					<TR >
						<TD CLASS=TD5>Ver.DT</TD>
						<TD CLASS=TD6>
							<TABLE CELLSPACING=0 CELLPADDING=0>
								<TR>
									<TD>
										<script language =javascript src='./js/p2250xa1_ko441_fpDateTime1_txtFromDt.js'></script>
									</TD>
									<TD STYLE="DISPLAY:NONE">
										&nbsp;~&nbsp;
									</TD>
									<TD STYLE="DISPLAY:NONE">
										<script language =javascript src='./js/p2250xa1_ko441_fpDateTime2_txtToDt.js'></script>
									</TD>
								</TR>
							</TABLE>
						</TD>								
						<TD CLASS="TD5" NOWRAP STYLE="DISPLAY:NONE">납품처</TD>						
						<TD CLASS="TD6" STYLE="DISPLAY:NONE"><INPUT NAME="txtShipToParty" ALT="납품처" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDnRef" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup C_PopShipToParty">&nbsp;<INPUT TYPE=TEXT NAME="txtShipToPartyNm" SIZE=25 TAG="14" ALT="납품처명"></TD> 
					</TR>
                    <TR  STYLE="DISPLAY:NONE">
						<TD CLASS="TD5" NOWRAP>품목</TD>
						<TD CLASS="TD6"><INPUT NAME="txtItemCd" ALT="품목" TYPE="Text" MAXLENGTH=18 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopItemCd">&nbsp;<INPUT NAME="txtItemNm" TYPE="Text" SIZE=25 tag="14" Alt="품목명"></TD>
						<TD CLASS="TD5" NOWRAP>창고</TD>
						<TD CLASS="TD6"><INPUT NAME="txtSlCd" ALT="창고" TYPE="Text" MAXLENGTH=18 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDnRef" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup C_PopSlCd">&nbsp;<INPUT NAME="txtSlNm" TYPE="Text" SIZE=25 tag="14" Alt="창고명"></TD>
                    </TR>
					<TR STYLE="DISPLAY:NONE">
						<TD CLASS=TD5>수주번호</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtSoNo" ALT="수주번호" SIZE=25 MAXLENGTH=18 TAG="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDnRef" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopSoNo"></TD>
						<TD CLASS=TD5 NOWRAP>Tracking No.</TD>
						<TD CLASS=TD6><INPUT NAME="txtTrackingNo" ALT="Tracking 번호" TYPE=TEXT MAXLENGTH=25 SIZE=25 TAG="11XXXU" TABINDEX=-1><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenTrackingNo()"></TD>
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
						<script language =javascript src='./js/p2250xa1_ko441_vaSpread_vspdData.js'></script>
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
					<TD WIDTH=70% NOWRAP >     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG></TD>
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
