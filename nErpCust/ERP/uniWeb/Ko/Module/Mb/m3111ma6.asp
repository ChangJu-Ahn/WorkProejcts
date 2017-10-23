<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name		  : Procurement
'*  2. Function Name		:
'*  3. Program ID		   : m3111ma6
'*  4. Program Name		 : 사급소요량조정등록 
'*  5. Program Desc		 :
'*  6. Comproxy List		:
'*  7. Modified date(First) : 2001/11/13
'*  8. Modified date(Last)  : 2003/04/07
'*  9. Modifier (First)	 : Jin-hyun Shin
'* 10. Modifier (Last)	  : Kim, Jinha
'* 11. Comment			  :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*							this mark(⊙) Means that "may  change"
'*							this mark(☆) Means that "must change"
'* 13. History			  : 2001/11/13
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '******************************************  1.1 Inc 선언   ***************************************** !-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  =======================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<!--'==========================================  1.1.2 공통 Include   =====================================!-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit

<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_PGM_ID 				= "m3111mb601.asp"
Const BIZ_PGM_ID_01				= "m3111mb602.asp"
Const BIZ_PGM_JUMP_ID_FOR_PO 	= "m3111ma1"
Const C_SHEETMAXROWS_D  = 100						'☜ : MB단과 꼭 일치시킬것.

'=== 첫번째 spread 상수 ===
Dim C_SpplCd
Dim C_SpplNm
Dim C_PlantCd
Dim C_PlantNm
Dim C_ItemCd
Dim C_ItemNm
Dim C_ItemSpec
Dim C_PoNo
Dim C_PoSeq
Dim C_PoDt
Dim C_PoQty
Dim C_PoUnit
Dim C_RcptQty
Dim C_SlCd
Dim C_SlNm
Dim C_TrackingNo
Dim C_GrpNm
Dim C_PrNo

'=== 두번째 spread 상수 ===
Dim C_ChildItemCd
Dim C_ChildItemPopup
Dim C_ChildItemNm
Dim C_ChildItemSpec
Dim C_SpplTypeNm
Dim C_IssueSlCd
Dim C_IssueSlPopup
Dim C_IssueSlNm
Dim C_ReservDt
Dim C_ReservQty
Dim C_LotPopup
Dim C_BkQty
Dim C_IssueQty
Dim C_IssueUnit
'hidden
Dim C_ResvdSeqNo
Dim C_PrStateCd
Dim C_HisSubSeqNo
Dim C_ReqmtNo
Dim C_pPrNo
Dim C_pPoNo
Dim C_pPoSeq
Dim C_SpplTypeCd
Dim C_pPoQty
Dim C_pPoUnit
Dim C_pPoDt
Dim C_pRcptQty
Dim C_pTracking_no
Dim C_pPlantCd
Dim C_pSpplCd
Dim C_OrgChildItemCd
Dim C_OrgSpplTypeCd
Dim C_OrgSlCd
Dim C_OrgReservQty
Dim C_OrgReservDt
Dim C_ParentRowNo
Dim C_ChildRowNo

Dim StartDate, EndDate

EndDate = "<%=GetSvrDate%>"
StartDate = UNIDateAdd("m", -1, EndDate, Parent.gServerDateFormat)
EndDate   = UniConvDateAToB(EndDate, Parent.gServerDateFormat, Parent.gDateFormat)
StartDate = UniConvDateAToB(StartDate, Parent.gServerDateFormat, Parent.gDateFormat)

Dim lgTopLeft
dim isQuery
Dim lgIntFlgModeM									'Variable is for Operation Status
Dim lglngHiddenRows()								'Multi에서 재쿼리를 위한 변수	'ex) 첫번째 그리드의 특정Row에 해당하는 두번째 그리드의 Row 갯수를 저장하는 배열.
Dim lgStrPrevKeyM()
Dim lgSortKey1
Dim lgSortKey2
Dim lgPageNo1
Dim lgSpdHdrClicked
Dim lgStrResvdSeqNo

'===================================================================================================================================
Dim IsOpenPop
Dim IgPrevRow
Dim lgCurrRow
Dim lgOldRow
Dim lgPopupMenuFlg

'===================================================================================================================================
Sub setCookieForPo()

	Dim strTemp, arrVal

	if frm1.vspdData.ActiveRow > 0 then
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = C_PoNo
		Call WriteCookie("PoNo" , frm1.vspdData.Text)
	end if

End Sub

'===================================================================================================================================
Sub InitVariables()

	lgIntFlgMode = Parent.OPMD_CMODE				   'Indicates that current mode is Create mode
	lgIntFlgModeM = Parent.OPMD_CMODE		'Indicates that current mode is Create mode
	lgBlnFlgChgValue = False					'Indicates that no value changed
	lgIntGrpCount = 0						   'initializes Group View Size

	lgLngCurRows = 0							'initializes Deleted Rows Count
	lgOldRow = 0
	lgSortKey1 = 2
	lgSortKey2 = 2
	lgPageNo = 0
	lgPageNo1 = 0
	lgStrResvdSeqNo = ""
	frm1.hdnQueryRow.Value = ""
	frm1.vspdData.MaxRows = 0
	frm1.vspdData1.MaxRows = 0

End Sub
'===================================================================================================================================
Sub SetDefaultVal()
	frm1.txtSupplierCd.focus
	frm1.txtPrFrDt.text = startDate
	frm1.txtPrToDt.text = endDate
	Set gActiveElement = document.activeElement
End Sub
'===================================================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
	<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
End Sub

'===================================================================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)

	If pvSpdNo = "A" Then
		C_SpplCd 			= 1
		C_SpplNm			= 2
		C_PlantCd			= 3
		C_PlantNm			= 4
		C_ItemCd			= 5
		C_ItemNm 			= 6
		C_ItemSpec			= 7
		C_PoNo				= 8
		C_PoSeq				= 9
		C_PoDt				= 10
		C_PoQty				= 11
		C_PoUnit			= 12
		C_RcptQty			= 13
		C_SlCd	 			= 14
		C_SlNm				= 15
		C_TrackingNo		= 16
		C_GrpNm				= 17
		C_PrNo				= 18
	Else
		C_ChildItemCd		= 1
		C_ChildItemPopup	= 2
		C_ChildItemNm		= 3
		C_ChildItemSpec		= 4
		C_SpplTypeNm		= 5
		C_IssueSlCd			= 6
		C_IssueSlPopup 		= 7
		C_IssueSlNm 		= 8
		C_ReservDt 			= 9
		C_ReservQty 		= 10
		C_LotPopup			= 11
		C_IssueQty 			= 12
		C_BkQty 			= 13
		C_IssueUnit			= 14
		C_ResvdSeqNo		= 15
		C_PrStateCd			= 16
		C_HisSubSeqNo		= 17
		C_ReqmtNo			= 18
		C_pPrNo				= 19
		C_pPoNo				= 20
		C_pPoSeq			= 21
		C_SpplTypeCd		= 22
		C_pPoQty			= 23
		C_pPoUnit			= 24
		C_pPoDt				= 25
		C_pRcptQty			= 26
		C_pTracking_no		= 27
		C_pPlantCd			= 28
		C_pSpplCd			= 29
		C_OrgChildItemCd	= 30
		C_OrgSpplTypeCd		= 31
		C_OrgSlCd			= 32
		C_OrgReservQty		= 33
		C_OrgReservDt		= 34
		C_ParentRowNo		= 35
		C_ChildRowNo		= 36
	End If

End Sub
'===================================================================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)

	Call InitSpreadPosVariables(pvSpdNo)

	If pvSpdNo = "A" Then

		With frm1.vspdData

			ggoSpread.Source = frm1.vspdData
			ggoSpread.Spreadinit "V20060321",,Parent.gAllowDragDropSpread

			.ReDraw  = false

			.MaxCols = C_PrNo + 1
			.Col = .MaxCols:		.ColHidden = True
			.MaxRows = 0
			Call AppendNumberPlace("6","5","0")
			Call GetSpreadColumnPos("A")

			ggoSpread.SSSetEdit	C_SpplCd		, "공급처", 10
			ggoSpread.SSSetEdit	C_SpplNm		, "공급처명", 20
			ggoSpread.SSSetEdit	C_PlantCd		, "공장",10
			ggoSpread.SSSetEdit	C_PlantNm		, "공장명", 20
			ggoSpread.SSSetEdit	C_ItemCd		, "모품목",10
			ggoSpread.SSSetEdit	C_ItemNm		, "모품목명",20
			ggoSpread.SSSetEdit	C_ItemSpec		, "규격",20
			ggoSpread.SSSetEdit	C_PoNo			, "발주번호",20
			ggoSpread.SSSetFloat C_PoSeq		, "발주순번" ,10,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,0
			ggoSpread.SSSetDate	C_PoDt			, "발주일",15, 2, Parent.gDateFormat
			SetSpreadFloatLocal	C_PoQty			, "발주수량",15,1,3
			ggoSpread.SSSetEdit	C_PoUnit		, "발주단위",10
			SetSpreadFloatLocal	C_RcptQty		, "입고수량",15,1,3'신규추가 
			ggoSpread.SSSetEdit	C_SlCd			, "창고", 10
			ggoSpread.SSSetEdit	C_SlNm			, "창고명",20
			ggoSpread.SSSetEdit	C_TrackingNo	, "Tracking No.", 20
			ggoSpread.SSSetEdit	C_GrpNm			, "구매그룹명",20
			ggoSpread.SSSetEdit	C_PrNo			, "요청번호",20

			Call ggoSpread.MakePairsColumn(C_SpplCd,C_SpplNm)
			Call ggoSpread.MakePairsColumn(C_PlantCd,C_PlantNm)
			Call ggoSpread.MakePairsColumn(C_ItemCd,C_ItemNm)
			Call ggoSpread.MakePairsColumn(C_SlCd,C_SlNm)

			.ReDraw  = True

		End With
	Else

		'------------------------------------------
		' Grid 2 - Component Spread Setting
		'------------------------------------------
		With frm1.vspdData1

			ggoSpread.Source = frm1.vspdData1
			ggoSpread.Spreadinit "V20060321",,Parent.gAllowDragDropSpread

			.ReDraw = false

			.MaxCols = C_ChildRowNo + 1
			.Col = .MaxCols:		.ColHidden = True
			.MaxRows = 0

			Call AppendNumberPlace("6","5","0")
			Call GetSpreadColumnPos("B")

			ggoSpread.SSSetEdit		C_ChildItemCd		, "자품목", 10,,,18,2
			ggoSpread.SSSetButton	C_ChildItemPopup
			ggoSpread.SSSetEdit		C_ChildItemNm		, "자품목명",20
			ggoSpread.SSSetEdit		C_ChildItemSpec		, "규격",20
			ggoSpread.SSSetCombo 	C_SpplTypeNm		, "지급구분",10,0,False
			ggoSpread.SSSetEdit		C_IssueSlCd			, "출고창고", 10
			ggoSpread.SSSetButton	C_IssueSlPopup
			ggoSpread.SSSetEdit		C_IssueSlNm			, "출고창고명",20
			ggoSpread.SSSetDate		C_ReservDt			, "출고예정일", 10, 2, Parent.gDateFormat
			SetSpreadFloatLocal		C_ReservQty			, "출고예정량",15,1,3
			ggoSpread.SSSetButton	C_LotPopup
			SetSpreadFloatLocal		C_IssueQty			, "출고수량",15,1,3
			SetSpreadFloatLocal		C_BkQty				, "소비수량",15,1,3
			ggoSpread.SSSetEdit		C_IssueUnit			, "단위",10

			ggoSpread.SSSetEdit		C_ResvdSeqNo		, "C_ResvdSeqNo",10
			ggoSpread.SSSetEdit		C_PrStateCd			, "C_PrStateCd",10
			ggoSpread.SSSetEdit		C_HisSubSeqNo		, "C_HisSubSeqNo",10
			ggoSpread.SSSetEdit		C_ReqmtNo			, "C_ReqmtNo",10
			ggoSpread.SSSetEdit		C_pPrNo				, "C_pPrNo", 10
			ggoSpread.SSSetEdit		C_pPoNo				, "C_pPoNo", 10
			ggoSpread.SSSetFloat	C_pPoSeq			, "C_pPoSeq" ,10,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,0
			ggoSpread.SSSetCombo 	C_SpplTypeCd		, "C_SpplTypeCd",10,0,False
			SetSpreadFloatLocal		C_pPoQty			, "C_pPoQty",15,1,3
			ggoSpread.SSSetEdit		C_pPoUnit			, "C_pPoUnit",10
			ggoSpread.SSSetDate		C_pPoDt				, "C_pPoDt", 10, 2, Parent.gDateFormat
			SetSpreadFloatLocal		C_pRcptQty			, "C_pRcptQty",15,1,3
			ggoSpread.SSSetEdit		C_pTracking_no		, "C_pTracking_no",20
			ggoSpread.SSSetEdit		C_pPlantCd			, "C_pPlantCd",10
			ggoSpread.SSSetEdit		C_pSpplCd			, "C_pSpplCd",10
			ggoSpread.SSSetEdit		C_OrgChildItemCd	, "C_OrgChildItemCd", 10,,,18,2
			ggoSpread.SSSetEdit		C_OrgSpplTypeCd		, "C_OrgSpplTypeCd", 10,,,18,2
			ggoSpread.SSSetEdit		C_OrgSlCd			, "C_OrgSlCd", 10,,,7,2
			SetSpreadFloatLocal		C_OrgReservQty		, "C_OrgReservQty",15,1,3
			ggoSpread.SSSetDate		C_OrgReservDt		, "C_OrgReservDt", 10, 2, Parent.gDateFormat

			ggoSpread.SSSetEdit		C_ParentRowNo		, "C_ParentRowNo", 10
			ggoSpread.SSSetEdit		C_ChildRowNo		, "C_ChildRowNo", 10

			Call ggoSpread.MakePairsColumn(C_ChildItemCd,C_ChildItemPopup)
			Call ggoSpread.MakePairsColumn(C_IssueSlCd,C_IssueSlPopup)
			Call ggoSpread.SSSetColHidden(C_ResvdSeqNo,	C_OrgReservDt,	True)
			Call ggoSpread.SSSetColHidden(C_ParentRowNo,	C_ChildRowNo,	True)

			ggoSpread.SpreadLock		-1,-1

		End With
	End If

End Sub

'===================================================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	frm1.vspdData1.ReDraw = False
	ggoSpread.SSSetProtected	frm1.vspddata.maxcols, pvStartRow, pvEndRow
	ggoSpread.SpreadUnlock		C_ChildItemCd	, pvStartRow, C_ChildItemCd, pvEndRow
	ggoSpread.SSSetRequired		C_ChildItemCd	, pvStartRow, pvEndRow
	ggoSpread.SpreadUnlock		C_ChildItemPopup, pvStartRow, C_ChildItemPopup, pvEndRow
	ggoSpread.SpreadLock		C_ChildItemNm	, pvStartRow, C_ChildItemNm, pvEndRow
	ggoSpread.SSSetProtected	C_ChildItemNm	, pvStartRow, pvEndRow
	ggoSpread.SpreadLock		C_ChildItemSpec	, pvStartRow, C_ChildItemSpec, pvEndRow
	ggoSpread.SpreadUnlock		C_SpplTypeNm	, pvStartRow, C_SpplTypeNm, pvEndRow
	ggoSpread.SSSetRequired		C_SpplTypeNm	, pvStartRow, pvEndRow
	ggoSpread.SpreadUnlock		C_IssueSlCd		, pvStartRow, C_IssueSlCd, pvEndRow
	ggoSpread.SSSetRequired		C_IssueSlCd		, pvStartRow, pvEndRow
	ggoSpread.SpreadUnlock		C_IssueSlPopup	, pvStartRow, C_IssueSlPopup, pvEndRow
	ggoSpread.SpreadLock		C_IssueSlNm		, pvStartRow, C_IssueSlNm, pvEndRow
	ggoSpread.SSSetProtected	C_IssueSlNm		, pvStartRow, pvEndRow
	ggoSpread.SpreadUnlock		C_ReservDt		, pvStartRow, C_ReservDt, pvEndRow
	ggoSpread.SSSetRequired		C_ReservDt		, pvStartRow, pvEndRow
	ggoSpread.SpreadUnlock		C_ReservQty		, pvStartRow, C_ReservQty, pvEndRow
	ggoSpread.SSSetRequired		C_ReservQty		, pvStartRow, pvEndRow
	ggoSpread.SpreadUnlock		C_LotPopup		, pvStartRow, C_LotPopup, pvEndRow
	ggoSpread.SpreadLock		C_IssueQty		, pvStartRow, C_IssueQty, pvEndRow
	ggoSpread.SSSetProtected	C_IssueQty		, pvStartRow, pvEndRow
	ggoSpread.SpreadLock		C_IssueUnit		, pvStartRow, C_IssueUnit, pvEndRow
	ggoSpread.SSSetProtected	C_IssueUnit		, pvStartRow, pvEndRow
	ggoSpread.SpreadLock		C_BkQty			, pvStartRow, C_BkQty, pvEndRow
	ggoSpread.SSSetProtected	C_BkQty			, pvStartRow, pvEndRow

	frm1.vspdData1.ReDraw = True
End Sub

'===================================================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos

	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_SpplCd 	= iCurColumnPos(1)
			C_SpplNm	= iCurColumnPos(2)
			C_PlantCd	= iCurColumnPos(3)
			C_PlantNm	= iCurColumnPos(4)
			C_ItemCd	= iCurColumnPos(5)
			C_ItemNm 	= iCurColumnPos(6)
			C_ItemSpec	= iCurColumnPos(7)
			C_PoNo		= iCurColumnPos(8)
			C_PoSeq		= iCurColumnPos(9)
			C_PoDt		= iCurColumnPos(10)
			C_PoQty		= iCurColumnPos(11)
			C_PoUnit	= iCurColumnPos(12)
			C_RcptQty	= iCurColumnPos(13)
			C_SlCd	 	= iCurColumnPos(14)
			C_SlNm		= iCurColumnPos(15)
			C_TrackingNo= iCurColumnPos(16)
			C_GrpNm		= iCurColumnPos(17)
			C_PrNo		= iCurColumnPos(18)

		Case "B"
			ggoSpread.Source = frm1.vspdData1
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_ChildItemCd		= iCurColumnPos(1)
			C_ChildItemPopup	= iCurColumnPos(2)
			C_ChildItemNm		= iCurColumnPos(3)
			C_ChildItemSpec		= iCurColumnPos(4)
			C_SpplTypeNm		= iCurColumnPos(5)
			C_IssueSlCd			= iCurColumnPos(6)
			C_IssueSlPopup 		= iCurColumnPos(7)
			C_IssueSlNm 		= iCurColumnPos(8)
			C_ReservDt 			= iCurColumnPos(9)
			C_ReservQty 		= iCurColumnPos(10)
			C_LotPopup			= iCurColumnPos(11)
			C_IssueQty 			= iCurColumnPos(12)
			C_BkQty 			= iCurColumnPos(13)
			C_IssueUnit			= iCurColumnPos(14)
			C_ResvdSeqNo		= iCurColumnPos(15)
			C_PrStateCd			= iCurColumnPos(16)
			C_HisSubSeqNo		= iCurColumnPos(17)
			C_ReqmtNo			= iCurColumnPos(18)
			C_pPrNo				= iCurColumnPos(19)
			C_pPoNo				= iCurColumnPos(20)
			C_pPoSeq			= iCurColumnPos(21)
			C_SpplTypeCd		= iCurColumnPos(22)
			C_pPoQty			= iCurColumnPos(23)
			C_pPoUnit			= iCurColumnPos(24)
			C_pPoDt				= iCurColumnPos(25)
			C_pRcptQty			= iCurColumnPos(26)
			C_pTracking_no		= iCurColumnPos(27)
			C_pPlantCd			= iCurColumnPos(28)
			C_pSpplCd			= iCurColumnPos(29)
			C_OrgChildItemCd	= iCurColumnPos(30)
			C_OrgSpplTypeCd		= iCurColumnPos(31)
			C_OrgSlCd			= iCurColumnPos(32)
			C_OrgReservQty		= iCurColumnPos(33)
			C_OrgReservDt		= iCurColumnPos(34)
			C_ParentRowNo		= iCurColumnPos(35)
			C_ChildRowNo		= iCurColumnPos(36)
	End Select
End Sub
'===================================================================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)

	With frm1
		If pvSpdNo = "A" Then
			ggoSpread.Source = frm1.vspdData
			.vspdData.ReDraw = False
			ggoSpread.SpreadLock -1, -1
			.vspdData.ReDraw = True
		Else
			ggoSpread.Source = frm1.vspdData1
			.vspdData1.ReDraw = False
			ggoSpread.SpreadLock -1, -1
			.vspdData1.ReDraw = True
		End IF
	End With
End Sub
'===================================================================================================================================
sub setSpreadColorQueryOk(ByVal pvStartRow, ByVal pvEndRow)

	Dim iRcptQty, iIssueQty, iBkQty ,ipoQty
	Dim index

	If UniCdbl(GetSpreadText(frm1.vspdData,C_RcptQty,frm1.vspdData.ActiveRow,"X","X")) = 0 Then
		ggoSpread.SpreadUnlock		C_SpplTypeNm, pvStartRow, C_SpplTypeNm, pvEndRow
		ggoSpread.SSSetRequired		C_SpplTypeNm, pvStartRow, pvEndRow
	Else
		ggoSpread.Spreadlock		C_SpplTypeNm, pvStartRow, C_SpplTypeNm, pvEndRow
		ggoSpread.SSSetProtected	C_SpplTypeNm, pvStartRow, pvEndRow
	end if
	
	
	
	
				

' ==== 2005.07.12 수정 ====================================================================
	for index = pvStartRow to pvEndRow
	    With frm1.vspdData1
		.Row = index
		.Col = C_pRcptQty
		iRcptQty = UniCdbl(.text)
		.Row = index
		.Col = C_pPoQty
		ipoQty = UniCDbl(.text)
	
	
    	End With
	
	    if iRcptQty < ipoQty then   '발주수량보다 작으면 

				If UniCdbl(GetSpreadText(frm1.vspdData1,C_IssueQty,index,"X","X")) <= 0 Then   '출고가 나가지 않으면 
				   If iRcptQty = 0 then
					  ggoSpread.SpreadUnlock		C_SpplTypeNm, index, C_SpplTypeNm, index
					  ggoSpread.SSSetRequired		C_SpplTypeNm, index, index
					End If  

					ggoSpread.SpreadUnlock		C_IssueSlCd, index, C_SpplTypeNm, index
					ggoSpread.SSSetRequired		C_IssueSlCd, index, index
					ggoSpread.SpreadUnlock		C_IssueSlPopup, index, C_IssueSlPopup, index
					ggoSpread.SpreadUnlock		C_ReservDt, index, C_SpplTypeNm, index
					ggoSpread.SSSetRequired		C_ReservDt, index, index
				    

					
				  
				Else
					ggoSpread.Spreadlock		C_SpplTypeNm, index, C_SpplTypeNm, index
					ggoSpread.SSSetProtected	C_SpplTypeNm, index, index

					ggoSpread.Spreadlock		C_IssueSlCd, index, C_SpplTypeNm, index
					ggoSpread.SSSetProtected	C_IssueSlCd, index, index
					ggoSpread.Spreadlock		C_IssueSlPopup, index, C_IssueSlPopup, index

					ggoSpread.Spreadlock		C_ReservDt, index, C_SpplTypeNm, index
					ggoSpread.SSSetProtected	C_ReservDt, index, index
				End if
				
				    ggoSpread.SpreadUnlock		C_ReservQty, index, C_SpplTypeNm, index
			        ggoSpread.SSSetRequired		C_ReservQty, index, index

			
		Else
		    ggoSpread.Spreadlock		C_ReservQty, index, C_SpplTypeNm, index
			ggoSpread.SSSetProtected	C_ReservQty, index, index	
		    		
		End If		
		
	
	Next
' ==== 2005.07.12 수정 ====================================================================

	ggoSpread.SpreadUnlock		C_LotPopup		, pvStartRow, C_LotPopup, pvEndRow
	ggoSpread.SpreadUnlock		C_SpplTypeCd	, pvStartRow, C_SpplTypeCd, pvEndRow
End sub
'===================================================================================================================================
Sub InitComboBox()
	Dim strDataCd, strDataNm
	Dim strCboDataCd
	Dim strCboDataNm
	Dim i

	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("M2201", "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	ggoSpread.Source = frm1.vspdData1

	ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_SpplTypeCd
	ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_SpplTypeNm

End Sub
'===================================================================================================================================
Sub InitData(ByVal lngStartRow)
	Dim intRow
	Dim intIndex

	With frm1.vspdData1
		For intRow = lngStartRow To .MaxRows
			.Row = intRow
			.col = C_SpplTypeCd
			intIndex = .value
			.Col = C_SpplTypeNm
			.value = intindex
		Next
	End With
End Sub
'===================================================================================================================================
Function OpenSupplier()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공급처"
	arrParam(1) = "B_BIZ_PARTNER"
	arrParam(2) = Trim(frm1.txtSupplierCd.Value)
	arrParam(3) = ""
	arrParam(4) = "BP_TYPE <> " & FilterVar("C", "''", "S") & "  And usage_flag=" & FilterVar("Y", "''", "S") & " "
	arrParam(5) = "공급처"

	arrField(0) = "BP_Cd"
	arrField(1) = "BP_NM"

	arrHeader(0) = "공급처"
	arrHeader(1) = "공급처명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtSupplierCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtSupplierCd.value  = arrRet(0)
		frm1.txtSupplierNm.value  = arrRet(1)
		frm1.txtSupplierCd.focus
		Set gActiveElement = document.activeElement
	End If
End Function

'===================================================================================================================================
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장"
	arrParam(1) = "B_PLANT"
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
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
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtPlantCd.Value= arrRet(0)
		frm1.txtPlantNm.value= arrRet(1)
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
	End If

End Function

'===================================================================================================================================
Function OpenPoNo()

	Dim strRet
	Dim iCalledAspName
	Dim IntRetCD

	If IsOpenPop = True Or UCase(frm1.txtPoNo.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	iCalledAspName = AskPRAspName("M3111PA6")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3111PA6", "X")
		IsOpenPop = False
		Exit Function
	End If

	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,""), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If strRet(0) = "" Then
		frm1.txtPoNo.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtPoNo.value = strRet(0)
		frm1.txtPoNo.focus
		Set gActiveElement = document.activeElement
	End If

End Function

'===================================================================================================================================
Function OpenItem()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName
	Dim IsOpenPop

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "자품목"
	arrParam(1) = "B_Item_By_Plant,B_Plant,B_Item"
	arrParam(2) = UCase(Trim(GetSpreadText(frm1.vspdData1,C_ChildItemCd,frm1.vspdData1.ActiveRow,"X","X")))
	arrParam(4) = "B_Item_By_Plant.Plant_Cd = B_Plant.Plant_Cd And "
	arrParam(4) = arrParam(4) & "B_Item_By_Plant.Item_Cd = B_Item.Item_Cd and B_Item.Phantom_flg=" & FilterVar("N", "''", "S") & "   "
	arrParam(4) = arrParam(4) & "And B_Plant.Plant_Cd= " & FilterVar(UCase(GetSpreadText(frm1.vspdData,C_PlantCd,frm1.vspdData.ActiveRow,"X","X")), "''", "S") & " "
	arrParam(4) = arrParam(4) & " AND B_Item_By_Plant.VALID_FROM_DT <=  " & FilterVar(UNIConvDate(Trim(GetSpreadText(frm1.vspdData,C_PoDt,frm1.vspdData.ActiveRow,"X","X"))), "''", "S") & " "
	arrParam(4) = arrParam(4) & " AND B_Item_By_Plant.VALID_TO_DT   >=  " & FilterVar(UNIConvDate(Trim(GetSpreadText(frm1.vspdData,C_PoDt,frm1.vspdData.ActiveRow,"X","X"))), "''", "S") & " "
	arrParam(5) = "자품목"

	arrField(0) = "B_Item.Item_Cd"
	arrField(1) = "B_Item.Item_NM"
	arrField(2) = "B_Item.Basic_unit"
	arrField(3) = "B_Plant.Plant_Cd"
	arrField(4) = "B_Plant.Plant_NM"

	arrHeader(0) = "자품목"
	arrHeader(1) = "자품목명"
	arrHeader(2) = "단위"
	arrHeader(3) = "공장"
	arrHeader(4) = "공장명"

	iCalledAspName = AskPRAspName("m1111pa1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "m1111pa1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam,arrField, arrHeader), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.vspdData1.Col = C_ChildItemCd
		frm1.vspdData1.Action = 0
		Set gActiveSpdSheet = document.activeElement
		Exit Function
	Else
		Call frm1.vspdData1.SetText(C_ChildItemCd,	frm1.vspdData1.ActiveRow,	arrRet(0))
		Call frm1.vspdData1.SetText(C_ChildItemNm,	frm1.vspdData1.ActiveRow,	arrRet(1))
		Call frm1.vspdData1.SetText(C_IssueUnit,	frm1.vspdData1.ActiveRow,	arrRet(3))

		Call vspdData1_Change(C_ChildItemCd, frm1.vspdData1.Row)
	End If
End Function
'===================================================================================================================================
Function OpenSl()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "창고"
	arrParam(1) = "B_Plant A ,B_Storage_Location B "
	arrParam(2) = UCase(GetSpreadText(frm1.vspdData1,C_IssueSlCd,frm1.vspdData1.ActiveRow,"X","X"))
	arrParam(3) = ""
	arrParam(4) = "A.Plant_CD = B.Plant_CD "
	arrParam(4) = arrParam(4) & "And B.Plant_Cd = " & FilterVar(UCase(GetSpreadText(frm1.vspdData1,C_pPlantCd,frm1.vspdData1.ActiveRow,"X","X")), "''", "S") & " "
' ==== 2005.07.12 출고창고 수정 ====================================================================
'	arrParam(4) = arrParam(4) & " And B.SL_TYPE <> " & FilterVar("E", "''", "S") & " "
' ==== 2005.07.12 출고창고 수정 ====================================================================
	arrParam(5) = "창고"

	arrField(0) = "B.Sl_Cd"
	arrField(1) = "B.Sl_Nm"

	arrHeader(0) = "창고"
	arrHeader(1) = "창고명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call frm1.vspdData1.SetText(C_IssueSlCd,	frm1.vspdData1.ActiveRow,	arrRet(0))
		Call frm1.vspdData1.SetText(C_IssueSlNm,	frm1.vspdData1.ActiveRow,	arrRet(1))

		Call vspdData1_Change(C_IssueSlCd, frm1.vspdData1.Row)
	End If
End Function

'===================================================================================================================================
Function OpenLotNo()

	Dim strRet
	Dim arrParam(8)
	Dim iCalledAspName
	Dim IntRetCD

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = UCase(GetSpreadText(frm1.vspdData1,C_IssueSlCd,frm1.vspdData1.ActiveRow,"X","X"))
	arrParam(1) = UCase(GetSpreadText(frm1.vspdData1,C_ChildItemCd,frm1.vspdData1.ActiveRow,"X","X"))
	arrParam(2) = GetSpreadText(frm1.vspdData1,C_pTracking_no,frm1.vspdData1.ActiveRow,"X","X")						'tracking No
	arrParam(3) = UCase(GetSpreadText(frm1.vspdData1,C_pPlantCd,frm1.vspdData1.ActiveRow,"X","X"))
	arrParam(4) = "J"						'Userflag
	arrParam(5) = ""
	arrParam(6) = ""
	arrParam(7) = ""
	arrParam(8) = UCase(GetSpreadText(frm1.vspdData1,C_IssueUnit,frm1.vspdData1.ActiveRow,"X","X"))

	iCalledAspName = AskPRAspName("I2212RA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "I2212RA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	strRet = window.showModalDialog(iCalledAspName, _
		Array(window.parent,arrParam(0),arrParam(1),arrParam(2),arrParam(3),arrParam(4),arrParam(5),arrParam(6),arrParam(7),arrParam(8)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If isEmpty(strRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.
	If strRet(0) = "" Then
		Exit Function
	End If

End Function

Function OpenTrackingNo()
	Dim arrRet
	Dim arrParam(5)
	Dim IntRetCD
	Dim iCalledAspName

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = ""	'주문처 
	arrParam(1) = ""	'영업그룹 
	arrParam(2) = Trim(frm1.txtPlantCd.value)	'공장 
	arrParam(3) = ""	'모품목 
	arrParam(4) = ""	'수주번호 
	arrParam(5) = ""	'추가 Where절 

'	arrRet = window.showModalDialog("../s3/s3135pa1.asp", Array(arrParam), _
'			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
 	iCalledAspName = AskPRAspName("S3135PA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "S3135PA1", "X")
		lgIsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet = "" Then
		frm1.txtTrackNo.focus
		Exit Function
	Else
		frm1.txtTrackNo.Value = Trim(arrRet)
		frm1.txtTrackNo.focus
	End If
End Function

'===================================================================================================================================
Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , _
					ByVal dColWidth , ByVal HAlign , _
					ByVal iFlag )

   Select Case iFlag
		Case 2															  '금액 
			ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggAmtOfMoneyNo	,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
		Case 3															  '수량 
			ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggQtyNo			,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
		Case 4															  '단가 
			ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggUnitCostNo	,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
		Case 5															  '환율 
			ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggExchRateNo	,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
		Case 6															  '환율 
			ggoSpread.SSSetFloat iCol, Header, dColWidth, "6" ,ggStrIntegeralPart,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,"0","999"
	End Select

End Sub

'===================================================================================================================================
Function ValidDateCheckLocal(pParentRow, pChildRow, pTmpPrNo, pObjFromDt, pObjToDt)
	Dim TmpMsg

	ValidDateCheckLocal = true

	If Len(Trim(pObjToDt)) And Len(Trim(pObjFromDt)) Then
		If UniConvDateToYYYYMMDD(pObjFromDt,Parent.gDateFormat,"") > UniConvDateToYYYYMMDD(pObjToDt,Parent.gDateFormat,"") Then
			Call DisplayMsgBox("970023","X", pParentRow & ": " & "행 " & "요청번호" & "[" & pTmpPrNo & "]" & " : " & pChildRow & "행" & ", " & "출고예정일", "발주일")
			Exit Function
		End If
	End If

	ValidDateCheckLocal = false

End Function

'===================================================================================================================================
Function changeItem()
	Dim iTmpChildCd, iTmpPlantCd
	Dim iTmpWhere

	Err.Clear

	If CheckRunningBizProcess = True Then
		Exit Function
	End If

	changeItem = False

	With frm1

		iTmpChildCd = 	FilterVar(GetSpreadText(frm1.vspdData1,C_ChildItemCd,frm1.vspdData1.ActiveRow,"X","X"), "''", "S")
		iTmpPlantCd = FilterVar(GetSpreadText(frm1.vspdData1,C_pPlantCd,frm1.vspdData1.ActiveRow,"X","X"), "''", "S")

		iTmpWhere = " A.PLANT_CD = " & iTmpPlantCd & " AND A.ITEM_CD = b.ITEM_CD AND B.ITEM_CD = " & iTmpChildCd & " AND A.ISSUED_SL_CD = C.SL_CD "

		If 	CommonQueryRs(" B.ITEM_NM, B.SPEC, B.BASIC_UNIT, C.SL_CD, C.SL_NM ", " B_ITEM_BY_PLANT A, B_ITEM B, B_STORAGE_LOCATION C ", iTmpWhere , _
					lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then

			Call DisplayMsgBox("122700","X","X","X")
			Call ClearCellByChlidItem()
			Call SheetFocus(.vspdData1.activeRow,C_ChildItemCd)
			Exit function
		End If

		lgF0 = Split(lgF0, Chr(11))
		lgF1 = Split(lgF1, Chr(11))
		lgF2 = Split(lgF2, Chr(11))
		lgF3 = Split(lgF3, Chr(11))
		lgF4 = Split(lgF4, Chr(11))

		Call .vspdData1.SetText(C_ChildItemNm,	.vspdData1.ActiveRow,	lgF0(0))
		Call .vspdData1.SetText(C_ChildItemSpec,.vspdData1.ActiveRow,	lgF1(0))
		Call .vspdData1.SetText(C_IssueUnit,	.vspdData1.ActiveRow,	lgF2(0))
		Call .vspdData1.SetText(C_IssueSlCd,	.vspdData1.ActiveRow,	lgF3(0))			'2005-05-16 수정 
		Call .vspdData1.SetText(C_IssueSlNm,	.vspdData1.ActiveRow,	lgF4(0))
	End With

	changeItem = True
End Function
'===================================================================================================================================
Function ClearCellByChlidItem()
	ClearCellByChlidItem = false
	With frm1.vspdData1
		Call .SetText(C_ChildItemNm,	.ActiveRow,		"")
		Call .SetText(C_ChildItemSpec,	.ActiveRow,		"")
		Call .SetText(C_IssueUnit,		.ActiveRow,		"")
		Call .SetText(C_IssueSlCd,		.ActiveRow,		"")
		Call .SetText(C_IssueSlNm,		.ActiveRow,		"")
	End With
	ClearCellByChlidItem = true
End Function
'===================================================================================================================================
Sub setFocusRow(ByVal pRow, ByVal cRow, ByVal cCol)
	Call SetActiveCell(frm1.vspdData,1,pRow,"M","X","X")
	Call DbQuery2(pRow,False)
	Call SetActiveCell(frm1.vspdData1,cCol,cRow,"M","X","X")
End Sub
'===================================================================================================================================
Sub fncDbDtlQuery()
	frm1.vspdData1.MaxRows = 0
	Call DbDtlQuery()
End Sub
'===================================================================================================================================
Function DbQuery2(ByVal Row, Byval NextQueryFlag)
	DbQuery2 = False

	Dim strVal
	Dim lngRet
	Dim lngRangeFrom
	Dim lngRangeTo
	Dim iStrPrNo, iStrPoNo , iStrPoSeq, iStrResvdSeqNo
	Dim pRow, lRow

	Call LayerShowHide(1)

	With frm1


		iStrPrNo	= UCase(Trim(GetSpreadText(frm1.vspdData,C_PrNo,Row,"X","X")))
		iStrPoNo	= UCase(Trim(GetSpreadText(frm1.vspdData,C_PoNo,Row,"X","X")))
		iStrPoSeq	= Trim(GetSpreadText(frm1.vspdData,C_PoSeq,Row,"X","X"))
		pRow		= Cint(GetSpreadText(frm1.vspdData,.vspdData.MaxCols,Row,"X","X"))

		If lglngHiddenRows(pRow - 1) <> 0 And NextQueryFlag = False Then
			.vspdData1.ReDraw = False

			lngRet = ShowFromData(pRow, lglngHiddenRows(pRow - 1))	'ex) 첫번째 그리드의 특정 Row에 해당하는 두번째 그리드의 Row수가 10개일때 보여줄 데이터가 3번째 부터 6번째까지 4개이면 3을 리턴하는 기능을 수행하는 함수다.
			Call ResetToolBar(Row,ShowDataFirstRow2())

			Call LayerShowHide(0)
			.vspdData1.ReDraw = True
			DbQuery2 = True
			Exit Function

		Else
			if (Trim(lgPageNo1)="") OR  (NextQueryFlag = False) then lgPageNo1=0

			strVal = BIZ_PGM_ID_01 & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtMaxRows=" & .vspdData1.MaxRows
			strVal = strVal & "&txtPrNo=" & Trim(iStrPrNo)
			strVal = strVal & "&txtPoNo=" & Trim(iStrPoNo)
			strVal = strVal & "&txtPoSeq=" & Trim(iStrPoSeq)
			strVal = strVal & "&txtNextFlag=" & NextQueryFlag
			strVal = strVal & "&lgPageNo1="	& lgPageNo1						'☜: Next key tag
			strVal = strVal & "&lgStrResvdSeqNo=" & lgStrPrevKeyM(Row-1)
			strVal = strVal & "&lglngHiddenRows=" & lglngHiddenRows(Row - 1)
			strVal = strVal & "&lRow=" & CStr(pRow)
		End IF
	End With

	Call RunMyBizASP(MyBizASP, strVal)
	DbQuery2 = True
End Function
'===================================================================================================================================
Function DbQueryOk2(Byval DataCount)
	DbQueryOk2 = false
	Dim lngRangeFrom
	Dim lngRangeTo
	Dim Index

	Call InitData(1)
	frm1.vspdData1.ReDraw = False

	With frm1.vspdData1
		lngRangeFrom = ShowDataFirstRow2()
		lngRangeTo = ShowDataLastRow2()

		.BlockMode = True
		.Row = lngRangeFrom
		.Row2 = lngRangeTo
		.Col = C_ChildRowNo

		.Col2 = C_ChildRowNo
		.DestCol = 0
		.DestRow = lngRangeFrom
		.Action = 19	'SS_ACTION_COPY_RANGE
		.BlockMode = False
	End With

' ==== 2005.07.12 수정 ====================================================================
'	For index = lngRangeFrom to lngRangeTo
		Call setSpreadColorQueryOk(lngRangeFrom,lngRangeTo)
'	Next
' ==== 2005.07.12 수정 ====================================================================

	frm1.vspdData1.ReDraw = True

	Call ResetToolBar(frm1.vspddata.activerow,lngRangeFrom)

	frm1.vspdData.focus
	Set gActiveElement = document.activeElement

	DbQueryOk2 = true
End Function
'===================================================================================================================================
Sub ResetToolBar(ByVal pRow,ByVal cRow)
	Dim prSts, releaseFlg
	Dim iIssueQty, iRcptQty, ipoQty
	Dim iHisSubSeqNo
	Dim lngRangeFrom, lngRangeTo

	lngRangeFrom = ShowDataFirstRow2()
	lngRangeTo = ShowDataLastRow2()

	If pRow <= 0 Then Exit Sub
	If cRow <= 0 Then Exit Sub

	With frm1.vspdData
		.Row = pRow
		.Col = C_RcptQty
		iRcptQty = UniCdbl(.text)
		
		.Col = C_PoQty
		ipoQty = UniCDbl(.text)
	
	
	End With

	If lngRangeTo - lngRangeFrom >= 0 and lngRangeTo>0 and lngRangeFrom>0 then
		With frm1.vspdData1
			.Row = cRow
			.Col = C_IssueQty
			iIssueQty = UniCDbl(.text)
		End With
		If iRcptQty = 0 OR iRcptQty = "" Then
			If iIssueQty > 0 Then
				Call SetToolbar("1110110100111111")					'Case 1 : 추가 
				lgPopupMenuFlg=1
			Else
				Call SetToolbar("1110111100111111")					'Case 2 : 추가/삭제 
				lgPopupMenuFlg=2
			End If
		Else
			If iIssueQty > 0   Then
				Call SetToolbar("1110100100011111")					'Case 3 : 추가(X)/삭제(X) =>변경 200612 hong
				lgPopupMenuFlg=3
			Else
				Call SetToolbar("1110111100111111")					'Case 4 : 추가/삭제 
				lgPopupMenuFlg=4
			End If
		End If
	Else
		If iRcptQty <= 0 Then
			Call SetToolbar("1110110100111111")					'Case 1 : 추가 
			lgPopupMenuFlg=1
		Else
		    if  ipoQty <=  iRcptQty  then
			    Call SetToolbar("1110100100011111")					'Case 3 : 추가(X)/삭제(X)
			    lgPopupMenuFlg=3
			End If   
		End If
	End IF

End Sub
'===================================================================================================================================
Function ShowFromData(Byval Row, Byval lngShowingRows)	'###그리드 컨버전 주의부분###
'ex) 첫번째 그리드의 특정 Row에 해당하는 두번째 그리드의 Row수가 10개일때 보여줄 데이터가 3번째 부터 6번째까지 4개이면 3을 리턴하는 기능을 수행하는 함수다.
	ShowFromData = 0

	Dim lngRow
	Dim lngStartRow

	With frm1.vspdData1

		Call SortSheet()

		lngStartRow = 0

		If .MaxRows < 1 Then Exit Function

		For lngRow = 1 To .MaxRows
			.Row = lngRow
			.Col = C_ParentRowNo
			If Row = CInt(.Text) Then
				lngStartRow = lngRow
				ShowFromData = lngRow
				Exit For
			End If
		Next

		If lngStartRow > 0 Then
			.BlockMode = True
			.Row = 1
			.Row2 = .MaxRows
			.Col = C_ChildRowNo
			.Col2 = C_ChildRowNo
			.DestCol = 0
			.DestRow = 1
			.Action = 19	'SS_ACTION_COPY_RANGE
			.RowHidden = False

			.BlockMode = False

			'ex) 첫번째 그리드의 특정 Row에 해당하는 두번째 그리드의 Row수가 10개일때 보여줄 데이터가 3번째 부터 6번째까지 4개이면 첫번째 부터 2번째 까지의 Row를 숨긴다.
			If lngStartRow > 1 Then
				.BlockMode = True
				.Row = 1
				.Row2 = lngStartRow - 1
				.RowHidden = True
				.BlockMode = False
			End If

			'ex) 첫번째 그리드의 특정 Row에 해당하는 두번째 그리드의 Row수가 10개일때 보여줄 데이터가 3번째 부터 6번째까지 4개이면 7번째 부터 마지막 까지의 Row를 숨긴다.
			If lngStartRow < .MaxRows Then
				If lngStartRow + lngShowingRows <= .MaxRows Then
					.BlockMode = True
					.Row = lngStartRow + lngShowingRows
					.Row2 = .MaxRows
					.RowHidden = True
					.BlockMode = False
				End If
			End If

			.BlockMode = False

			.Row = lngStartRow
			.Col = 0
			.Action = 0
		End If
	End With
End Function
'===================================================================================================================================
Function DeleteDataForInsertSampleRows(ByVal Row, Byval lngShowingRows)
	DeleteDataForInsertSampleRows = False

	Dim lngRow
	Dim lngStartRow

	With frm1.vspdData1

		Call SortSheet()

		lngStartRow = 0
		For lngRow = 1 To .MaxRows
			.Row = lngRow
			.Col = C_ParentRowNo
			If Row = Clng(.Text) Then
				lngStartRow = lngRow
				DeleteDataForInsertSampleRows = True
				Exit For
			End If
		Next

		If lngStartRow > 0 Then
			.BlockMode = True
			.Row = lngStartRow
			.Row2 = lngStartRow + lngShowingRows - 1
			.Action = 5		'5 - Delete Row 	SS_ACTION_DELETE_ROW
			.MaxRows = .MaxRows - lngShowingRows
			.BlockMode = False
		End If
	End With
End Function
'===================================================================================================================================
Function SortSheet()
	SortSheet = false

	With frm1.vspdData1
		.BlockMode = True
		.Col = 0
		.Col2 = .MaxCols
		.Row = 1
		.Row2 = .MaxRows
		.SortBy = 0 'SS_SORT_BY_ROW

		.SortKey(1) = C_ParentRowNo
		.SortKey(2) = C_ChildRowNo

		.SortKeyOrder(1) = 0 'SS_SORT_ORDER_ASCENDING
		.SortKeyOrder(2) = 0 'SS_SORT_ORDER_ASCENDING

		.Col = 1	'C_SupplierCd	'###그리드 컨버전 주의부분###
		.Col2 = .MaxCols
		.Row = 1
		.Row2 = .MaxRows
		.Action = 25 'SS_ACTION_SORT

		.BlockMode = False
	End With
	SortSheet = true
End Function
'===================================================================================================================================
Function DefaultCheck()
	DefaultCheck = False
	Dim i
	Dim j
	Dim RequiredColor

	ggoSpread.Source = frm1.vspdData1
	RequiredColor = ggoSpread.RequiredColor
	With frm1.vspdData1
		For i = 1 To .MaxRows
			.Row = i
			If .RowHidden = False Then
				.Col = 0
				If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Then
					For j = 1 To .MaxCols
						.Col = j
						If .BackColor = RequiredColor Then
							If Len(Trim(.Text)) < 1 Then
								.Row = 0
								Call DisplayMsgBox("970021","X",.Text,"")
								.Row = i
								.Action = 0
								Exit Function
							End If
						End If
					Next
				End If
			End If
		Next
	End With
	DefaultCheck = True
End Function
'===================================================================================================================================
Function ChangeCheck()
	ChangeCheck = False

	Dim i
	Dim strInsertMark
	Dim strDeleteMark
	Dim strUpdateMark

	ggoSpread.Source = frm1.vspdData1
	strInsertMark = ggoSpread.InsertFlag
	strDeleteMark = ggoSpread.UpdateFlag
	strUpdateMark = ggoSpread.DeleteFlag

	With frm1.vspdData1
		For i = 1 To .MaxRows
			.Row = i
			.Col = 0
			If .Text = strInsertMark Or .Text = strDeleteMark Or .Text = strUpdateMark Then
				ChangeCheck = True
			End If
		Next
	End With
End Function
'===================================================================================================================================
Function CheckDataExist()
	CheckDataExist = False
	Dim i

	If frm1.vspdData1.MaxRows = 0 Then
		With frm1.vspdData
			.Row = .ActiveRow
			.Col = C_CfmFlg
			If .value = 1 AND frm1.vspdData1.RowHidden = False Then
				CheckDataExist = True
				Exit Function
			End If
		End With
	Else
		With frm1.vspdData1
			For i = 1 To .MaxRows
				.Row = i
				If .RowHidden = False Then
					CheckDataExist = True
					Exit Function
				End IF
			Next
		End With
	End IF
End Function
'===================================================================================================================================
Function ShowDataFirstRow()
	ShowDataFirstRow = 0
	Dim i

	With frm1.vspdData
		For i = 1 To .MaxRows
			.Row = i
			If .RowHidden = False Then
				ShowDataFirstRow = i
				Exit Function
			End If
		Next
	End With
End Function
'===================================================================================================================================
Function ShowDataFirstRow2()
	ShowDataFirstRow2 = 0
	Dim i

	With frm1.vspdData1
		For i = 1 To .MaxRows
			.Row = i
			If .RowHidden = False Then
				ShowDataFirstRow2 = i
				Exit Function
			End If
		Next
	End With
End Function
'===================================================================================================================================
Function ShowDataLastRow()
	ShowDataLastRow = 0
	Dim i

	With frm1.vspdData
		For i = .MaxRows To 1 Step -1
			.Row = i
			If .RowHidden = False Then
				ShowDataLastRow = i
				Exit Function
			End If
		Next
	End With
End Function
'===================================================================================================================================
Function ShowDataLastRow2()
	ShowDataLastRow2 = 0
	Dim i

	With frm1.vspdData1
		For i = .MaxRows To 1 Step -1
			.Row = i
			If .RowHidden = False Then
				ShowDataLastRow2 = i
				Exit Function
			End If
		Next
	End With
End Function
'===================================================================================================================================
Function DataFirstRow(ByVal Row)
	DataFirstRow = 0
	Dim i
	With frm1.vspdData1
		For i = 1 To .MaxRows
			.Row = i
			.Col = C_ParentRowNo
			If Clng(.text) = Clng(Row) Then
				DataFirstRow = i
				Exit Function
			End If
		Next
	End With
End Function
'===================================================================================================================================
Function DataLastRow(ByVal Row)
	DataLastRow = 0
	Dim i

	With frm1.vspdData1
		For i = .MaxRows To 1 Step -1
			.Row = i
			.Col = C_ParentRowNo
			If Clng(.text) = Clng(Row) Then
				DataLastRow = i
				Exit Function
			End If
		Next
	End With
End Function
'===================================================================================================================================
Sub InsertSampleRows()
	Dim i
	Dim j
	Dim lngMaxRows
	Dim strInspItemCd
	Dim strInspSeries
	Dim lngOldMaxRows
	Dim strMark
	Dim lRow

	With frm1
		If .vspdData.Row < 1 Then
			Exit Sub
		End If

   		Call LayerShowHide(1)

		lRow = .vspdData.ActiveRow
		' 해당 검사항목/차수를 가지고 있는 측정치들 삭제 
		Call DeleteDataForInsertSampleRows(lRow, lglngHiddenRows(lRow - 1))

		' 행 추가 
		lngOldMaxRows = .vspdData1.MaxRows

		 .vspdData.Row = lRow
		.vspdData.Col = C_ApportionQty
		lngMaxRows = UNICDbl(.vspdData.Text)
		.vspdData1.MaxRows = lngOldMaxRows + lngMaxRows

	End With

	ggoSpread.Source = frm1.vspdData1
	strMark = ggoSpread.InsertFlag

	With frm1.vspdData1
		.BlockMode = True
		.Row = lngOldMaxRows + 1
		.Row2 = .MaxRows
		.Col = C_ParentRowNo
		.Col2 = C_ParentRowNo
		.Text = lRow
		.BlockMode = False

		j = 0
		For i = lngOldMaxRows + 1 To .MaxRows
			j = j + 1
			.Row = i
			.Col = 0
			.Text = strMark
			.Col = C_SupplierCd
			.Text = j
		Next
	End With

	frm1.vspdData.Col = C_InspUnitIndctnCd

	Call SetSpreadColor2byInspUnitIndctn(lngOldMaxRows + 1, frm1.vspdData1.MaxRows, frm1.vspdData.Text, "I")

	frm1.vspdData1.Row = lngOldMaxRows + 1
	frm1.vspdData1.Col = C_SpplCd
	frm1.vspdData1.Action = 0
	lglngHiddenRows(lRow - 1) = lngMaxRows

	Call LayerShowHide(0)
End Sub
'===================================================================================================================================
Sub Form_Load()

	Call LoadInfTB19029
	Call ggoOper.LockField(Document, "N")								   '⊙: Lock  Suitable  Field
	Call InitSpreadSheet("A")
	Call InitSpreadSheet("B")												 '⊙: Setup the Spread sheet
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call InitVariables
	Call SetDefaultVal
	Call SetToolbar("1100000000001111")										'⊙: 버튼 툴바 제어 
	Call InitComboBox()
	lgPopupMenuFlg=0
End Sub
'===================================================================================================================================
Sub txtPrFrDt_DblClick(Button)
	if Button = 1 then
		frm1.txtPrFrDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtPrFrDt.Focus
	End if
End Sub
'===================================================================================================================================
Sub txtPrToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtPrToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtPrToDt.Focus
	End if
End Sub
'===================================================================================================================================
Sub txtPrFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'===================================================================================================================================
Sub txtPrToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'===================================================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)	'###그리드 컨버전 주의부분###
 	gMouseClickStatus = "SPC"

 	Set gActiveSpdSheet = frm1.vspdData

 	If Row <= 0 Then
 		Call SetPopupMenuItemInf("0000111111")		 '화면별 설정 
	Else
		Call SetPopupMenuItemInf("0001111111")		 '화면별 설정 
	End IF

	If frm1.vspdData.MaxRows = 0 Then
 		Exit Sub
 	End If

 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData
 		If lgSortKey1 = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey1 = 2
 		ElseIf lgSortKey1 = 2 Then

 			ggoSpread.SSSort Col, lgSortKey1		'Sort in Descending
 			lgSortKey1 = 1
 		End If

 		lgIntFlgModeM = Parent.OPMD_CMODE

 	Else
 		'------ Developer Coding part (Start)
 		lgSpdHdrClicked = 0
 		Call Sub_vspdData_ScriptLeaveCell(0, 0, Col, frm1.vspdData.ActiveRow, False)
	 	'------ Developer Coding part (End)
 	End If

End Sub
'===================================================================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)

	Dim iActiveRow
	Dim lngStartRow
	Dim iStrChildRow
	Dim i, K
	gMouseClickStatus = "SP2C"
	Set gActiveSpdSheet = frm1.vspdData1
	ggoSpread.Source = frm1.vspdData1

	With frm1

		If .vspdData1.MaxRows = 0 Then Exit Sub

		If Row <= 0 AND Col <> 0 Then
 			ggoSpread.Source = .vspdData1

 			.vspdData.Row = .vspdData.ActiveRow
 			.vspdData.Col = .vspdData.MaxCols
			iActiveRow = Cint(.vspdData.Text)

 			.vspdData1.Redraw = False
			lngStartRow = CInt(ShowFromData(iActiveRow, CInt(lglngHiddenRows(iActiveRow - 1))))
			.vspdData1.Redraw = True

			If lgSortKey2 = 1 Then
 				ggoSpread.SSSort Col, lgSortKey2, lngStartRow, lngStartRow + CInt(lglngHiddenRows(iActiveRow - 1)) - 1	'Sort in Ascending
 				lgSortKey2 = 2
 			ElseIf lgSortKey2 = 2 Then
 				ggoSpread.SSSort Col, lgSortKey2, lngStartRow, lngStartRow + CInt(lglngHiddenRows(iActiveRow - 1)) - 1	'Sort in Descending
 				lgSortKey2 = 1
			End If

		Else
 		End If
 	End With

 	Call ResetToolBar(frm1.vspdData.ActiveRow,Row)

	If lgIntFlgMode <> Parent.OPMD_CMODE Then
		Call SetPopupMenuItemInf("0000111111")
	Else
		If lgPopupMenuFlg=1 Then
			Call SetPopupMenuItemInf("1001111111")
		ElseIF lgPopupMenuFlg=2 Then
			Call SetPopupMenuItemInf("1101111111")
		ElseIF lgPopupMenuFlg=3 Then
			Call SetPopupMenuItemInf("0000111111")
		ElseIF lgPopupMenuFlg=4 Then
			Call SetPopupMenuItemInf("0101111111")
		Else
			Call SetPopupMenuItemInf("0001111111")
		End If
	End If

 	With frm1.vspdData1
 		For i = 1 to .MaxRows
 			.Row = i
 			.col = 0
 			If .Rowhidden = False Then
 				k = K + 1
 				if .text <> ggoSpread.InsertFlag  AND .text <> ggoSpread.UpdateFlag AND .text <> ggoSpread.deleteFlag then
 					.text = k
 				end if
 			End If
 		Next
 	End With

End Sub
'===================================================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	If lgSpdHdrClicked = 1 Then
		Exit Sub
	End If
	if frm1.vspddata.row = 0 then exit sub
	Call Sub_vspdData_ScriptLeaveCell(Col, Row, NewCol, NewRow, Cancel)
End Sub
'===================================================================================================================================
Sub Sub_vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	Dim lRow
	if Row = 0 then exit sub
	If Row <> NewRow And NewRow > 0 Then
		With frm1
			If CheckRunningBizProcess = True Then
				.vspdData.Row = Row
				.vspdData.Col = 1
				.vspdData.Action = 0
				Exit Sub
			End If
			'/* 다른 작업이 이루어지는 상황에서 다른 행 이동 시 조회가 이루어 지지 않도록 한다. - END */
			lgCurrRow = NewRow
		End With

		lgIntFlgModeM = Parent.OPMD_CMODE

		With frm1.vspdData1
			.ReDraw = False
			.BlockMode = True
			.Row = 1
			.Row2 = .MaxRows
			.RowHidden = True
			.BlockMode = False
			.ReDraw = True
		End With
		If DbQuery2(NewRow, False) = False Then	Exit Sub

	End If
End Sub
'===================================================================================================================================
Sub vspdData1_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

	if frm1.vspddata1.row = 0 then exit sub

End Sub
'===================================================================================================================================
Sub vspdData1_Change(ByVal Col , ByVal Row)
	Dim LngFindRow
	Dim strPrNo, strPrSeq
	Dim iparentrow
	Dim strMark

	With frm1

		ggoSpread.Source = frm1.vspdData1
		ggoSpread.UpdateRow Row

		iparentrow = GetSpreadText(frm1.vspdData1,C_ParentRowNo,Row,"X","X")

		Call .vspdData1.SetText(C_ChildRowNo,	Row,	Trim(GetSpreadText(frm1.vspdData1,0,Row,"X","X")))

		Select Case Col
			Case C_ChildItemCd
				If Trim(GetSpreadText(frm1.vspdData1,C_ChildItemCd,Row,"X","X")) = "" Then
					Exit Sub
				End If

				Call changeItem()
		End Select

'		Call .vspdData.SetText(0,	CLng(iparentrow),	ggoSpread.UpdateFlag)

	End with

	' === 두번째 스프레드 변경시 상위 스프레드 플래그 부분 수정 by MJG 2005.07.12 ========================

	Dim IRow

	With frm1

		IRow = .vspdData.ActiveRow

		If Trim(GetSpreadText(.vspdData,0,IRow ,"X","X")) <> ggoSpread.InsertFlag then
'			.vspdData1.Row = .vspdData1.ActiveRow
			.vspdData.Col = 0
			.vspdData.text = ggoSpread.UpdateFlag

		End If
	End With

	' === 두번째 스프레드 변경시 상위 스프레드 플래그 부분 수정 by MJG 2005.07.12 ========================

End Sub


'===================================================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

	If y<20 Then
		lgSpdHdrClicked = 1
	End If

	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub
'===================================================================================================================================
Sub vspdData1_MouseDown(Button , Shift , x , y)
	If y<20 Then
		lgSpdHdrClicked = 1
	End If

	If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub
'===================================================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.Source = frm1.vspdData
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub
'===================================================================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.Source = frm1.vspdData1
	Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub
'===================================================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
	ggoSpread.Source = frm1.vspdData
	Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )
	Call GetSpreadColumnPos("A")
End Sub
'===================================================================================================================================
Sub vspdData1_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

	ggoSpread.Source = frm1.vspdData1
	Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
	Call GetSpreadColumnPos("B")
End Sub
'===================================================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
	If Row <= 0 Then
		Exit Sub
	End If
	If frm1.vspddata.MaxRows=0 Then
		Exit Sub
	End If

End Sub
'===================================================================================================================================
Sub vspdData1_DblClick(ByVal Col, ByVal Row)
 	Dim iColumnName

 	If Row <= 0 Then
		Exit Sub
 	End If

  	If frm1.vspdData1.MaxRows = 0 Then
		Exit Sub
 	End If
End Sub
'===================================================================================================================================
Sub vspdData1_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	With frm1.vspdData1
		.Row = Row
		.Col = Col
		intIndex = .Value
		.Col = C_SpplTypeCd
		.Value = intIndex
	End With
End Sub
'===================================================================================================================================
Sub FncSplitColumn()

	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
	   Exit Sub
	End If

	ggoSpread.Source = gActiveSpdSheet
	ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)

End Sub
'===================================================================================================================================
Sub PopSaveSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.SaveSpreadColumnInf()

End Sub
'===================================================================================================================================
Sub PopRestoreSpreadColumnInf()	'###그리드 컨버전 주의부분###
	Dim iActiveRow
	Dim iConvActiveRow
	Dim lngRangeFrom
	Dim lngRangeTo
	Dim lRow
	Dim index
	Dim strFlag
	Dim strParentRowNo
	Dim i
	ggoSpread.Source = gActiveSpdSheet

	If gActiveSpdSheet.ID = "A" Then
		Call ggoSpread.RestoreSpreadInf()
		Call InitSpreadSheet("A")
		Call ggoSpread.ReOrderingSpreadData
		ggoSpread.SpreadLock -1, -1

	ElseIf gActiveSpdSheet.ID = "B" Then
		'이해안가는 코드 
		'For i = 1 To frm1.vspdData1.MaxRows
		'	strFlag = GetSpreadText(frm1.vspdData1,0,i,"X","X")
		'Next

		Call ggoSpread.RestoreSpreadInf()
		Call InitSpreadSheet("B")
		Call InitComboBox()
		frm1.vspdData1.Redraw = False

		Call ggoSpread.ReOrderingSpreadData("F")

		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = frm1.vspdData.MaxCols

		Call DbQuery2(frm1.vspdData.Text,False)
		Call InitData(1)

		lngRangeFrom = Clng(ShowDataFirstRow2)
		lngRangeTo = Clng(ShowDataLastRow2)

		frm1.vspdData1.ReDraw = True
' ==== 2005.07.12 수정 ====================================================================
'		For index = lngRangeFrom to lngRangeTo
			Call setSpreadColorQueryOk(lngRangeFrom,lngRangeTo)
'		Next
' ==== 2005.07.12 수정 ====================================================================
		frm1.vspdData1.ReDraw = False

		lRow = frm1.vspdData.ActiveRow	'###그리드 컨버전 주의부분###
		frm1.vspdData1.Redraw = True
		ggoSpread.Source = frm1.vspdData
		ggoSpread.EditUndo lRow
	End If
End Sub
'===================================================================================================================================
Sub vspdData_DragDropBlock(ByVal Col , ByVal Row , ByVal Col2 , ByVal Row2 , ByVal NewCol , ByVal NewRow , ByVal NewCol2 , ByVal NewRow2 , ByVal Overwrite , Action , DataOnly , Cancel )
	Row = 0: Row2 = -1: NewRow = 0
	ggoSpread.SwapRange Col, Row, Col2, Row2, NewCol, NewRow, Cancel
End Sub
'===================================================================================================================================
Sub vspdData_GotFocus()
	ggoSpread.Source = frm1.vspdData
End Sub
'===================================================================================================================================
Sub vspdData1_GotFocus()
	ggoSpread.Source = frm1.vspdData1
End Sub
'===================================================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	Dim intListGrvCnt
	Dim LngLastRow
	Dim LngMaxRow

	If OldLeft <> NewLeft Then
		Exit Sub
	End If

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then			'☜: 재쿼리 체크 
  		if Trim(lgPageNo)="" Then exit sub
		If lgPageNo > 0 Then			'다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If

			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Exit Sub
			End If
		End If
	End If
End Sub
'===================================================================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	Dim intListGrvCnt
	Dim LngLastRow
	Dim LngMaxRow
	Dim lRow
	Dim lConvRow

	If OldLeft <> NewLeft Then Exit Sub

	With frm1
		.vspdData1.Row = .vspdData1.ActiveRow
		.vspdData1.Col = C_ParentRowNo
		lRow = .vspdData1.text
		if Trim(lgPageNo1)="" Then exit sub
		if Trim(lRow)="" Then exit sub

		If ShowDataLastRow2 < NewTop + VisibleRowCnt(.vspdData1, NewTop) Then			'☜: 재쿼리 체크 
			If lgStrPrevKeyM(lRow-1) <> "" then
				lgPageNo1 = lglngHiddenRows(lRow - 1) \ C_SHEETMAXROWS_D
				If CheckRunningBizProcess = True Then
					Exit Sub
				End If

				Call DisableToolBar(Parent.TBC_QUERY)
				If DbQuery2(lRow, True) = False Then
					Exit Sub
				End If
			End If
		End If
	End With
End Sub
'===================================================================================================================================
Sub vspdData1_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Select Case Col
		Case C_ChildItemPopup
			Call OpenItem()
		Case C_IssueSlPopup
			Call OpenSl()
		Case C_LotPopup
			Call OpenLotNo()
	End Select
End Sub
'===================================================================================================================================
Function FncSave()
	FncSave = False

	Dim IntRetCD

	Err.Clear

	If CheckRunningBizProcess = True Then
		Exit Function
	End If

	ggoSpread.Source = frm1.vspdData1

	If ChangeCheck = False Then
		IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
		Exit Function
	End If

	If DefaultCheck = False Then
		Exit Function
	End If

	If Not chkField(Document, "1") Then
	   		Exit Function
	End If

	If Not chkField(Document, "2") Then
	   		Exit Function
	End If

	If DbSave = False then
		Exit Function
	End If

	FncSave = True
End Function
'===================================================================================================================================
 Function FncQuery()
	Dim IntRetCD

	FncQuery = False														'⊙: Processing is NG

	On Error Resume Next
	Err.Clear															   '☜: Protect system from crashing

	ggoSpread.Source = frm1.vspdData1

	If ChangeCheck = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
	Call InitVariables

	If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
	   Exit Function
	End If

	If (UniConvDateToYYYYMMDD(frm1.txtPrFrDt.text,Parent.gDateFormat,"") > Parent.UniConvDateToYYYYMMDD(frm1.txtPrToDt.text,Parent.gDateFormat,"")) and Trim(frm1.txtPrFrDt.text)<>"" and Trim(frm1.txtPrToDt.text)<>"" then
		Call DisplayMsgBox("17a003", "X","요청일", "X")
		Exit Function
	End if

	If Dbquery = False then Exit Function
	FncQuery = True																'⊙: Processing is OK

End Function
'===================================================================================================================================
Function FncNew()
	Dim IntRetCD

	FncNew = False														  '⊙: Processing is NG

	Err.Clear

	ggoSpread.Source = frm1.vspdData1

	If ChangeCheck = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	Call ggoOper.ClearField(Document, "1")										 '⊙: Clear Condition Field
	Call ggoOper.ClearField(Document, "Q")

	Call InitVariables

	ggoSpread.Source = frm1.vspdData	'###그리드 컨버전 주의부분###
	ggoSpread.ClearSpreadData

	ggoSpread.Source = frm1.vspdData1
	ggoSpread.ClearSpreadData

	Call SetDefaultVal
	Call SetToolbar("11100000000000")

	FncNew = True														   '⊙: Processing is OK

End Function
'===================================================================================================================================
Function FncDeleteRow()		'###그리드 컨버전 주의부분###
	FncDeleteRow = false

	Dim lDelRows
	Dim iDelRowCnt, i
	Dim lngRangeFrom
	Dim lngRangeTo
	Dim iparentrow

	If frm1.vspdData.MaxRows < 1 then
		Exit function
	End if

	'Check Spread2 Data Exists for the keys
	If CheckDataExist = False Then
		Exit function
	End If

	With frm1.vspdData1
		.Redraw = False

		.Focus

		'범위가 보이지 않는 곳까지 넘어갔을 경우에 대한 처리 - START
		lngRangeFrom = .SelBlockRow
		.Row = lngRangeFrom
		If .RowHidden = True Then
			lngRangeFrom = ShowDataFirstRow2()
		End If

		lngRangeTo = .SelBlockRow2
		.Row = lngRangeTo
		If .RowHidden = True Then
			lngRangeTo = ShowDataLastRow2()
		End If

		.BlockMode = True
		.Row = lngRangeFrom
		.Row2 = lngRangeTo
		.Action = 2			'Select Block	SS_ACTION_SELECT_BLOCK
		.BlockMode = False
		'범위가 보이지 않는 곳까지 넘어갔을 경우에 대한 처리 - END

		ggoSpread.Source = frm1.vspdData1
		 '----------  Coding part  -------------------------------------------------------------
		lDelRows = ggoSpread.DeleteRow
		.Row = lngRangeFrom
		.Col = C_ParentRowNo
		iparentrow = .text

		.BlockMode = True
		.Row = lngRangeFrom
		.Row2 = lngRangeTo
		.Col = 0
		.Col2 = 0
		.DestCol = C_ChildRowNo
		.DestRow = .SelBlockRow
		.Action = 19	'SS_ACTION_COPY_RANGE
		.BlockMode = False

		.Redraw = True
	End With

	With frm1.vspdData
' update 20060315 by kjt (scr : 20060310-34614)
'		.Row = iparentrow
		.Row = .ActiveRow
		.Col = 0
		.text = ggoSpread.UpdateFlag
	End With
	FncDeleteRow = true
End Function
'===================================================================================================================================
Function FncDelete()
	Dim lDelRows
	Dim iDelRowCnt, i
	if frm1.vspdData.Maxrows < 1 then exit function

	With frm1.vspdData
		.focus
		ggoSpread.Source = frm1.vspdData
		lDelRows = ggoSpread.DeleteRow

	End With
End Function
'===================================================================================================================================
Function FncCopy()
	FncCopy = false

	Dim lRow, lRow2
	Dim lngRangeFrom, lngRangeTo
	Dim strFlag
	Dim i, k

	With frm1

		If CheckDataExist = False Then
			Exit function
		End If

		.vspdData1.ReDraw = False

		ggoSpread.Source = frm1.vspdData1
		ggoSpread.CopyRow

		Call frm1.vspdData1.SetText(C_ChildRowNo,		.vspdData1.ActiveRow,	ggoSpread.InsertFlag)
		Call SetSpreadColor(.vspdData1.ActiveRow, .vspdData1.ActiveRow)

		'재쿼리를 위해 해당 키에 대한 Client의 Data Row수를 가져감 
		Call frm1.vspdData1.SetText(C_ParentRowNo,		.vspdData1.ActiveRow,	.vspdData.ActiveRow)

		lglngHiddenRows(Cint(.vspdData.ActiveRow) - 1) = lglngHiddenRows(Cint(.vspdData.ActiveRow) - 1) + 1

		lngRangeFrom = ShowDataFirstRow2()
		lngRangeTo = ShowDataLastRow2()

		Call frm1.vspdData.SetText(0,		.vspdData.ActiveRow,	ggoSpread.UpdateFlag)

		.vspdData1.ReDraw = True
		.vspdData1.focus
	End With
	FncCopy = true
	Set gActiveSpdSheet = frm1.vspdData1

End Function
'===================================================================================================================================
Function FncInsertRow(ByVal pvRowCnt)	'###그리드 컨버전 주의부분###
	FncInsertRow = false

	On Error Resume Next

	Dim lRow
	Dim lRow2
	Dim lconvRow
	Dim strMark
	Dim iInsertRow
	Dim IntRetCD
	Dim imRow
	Dim strInspUnitIndctnCd
	Dim iParentRowNo,iparentrow, iStrPrNo, iStrPoNo, iStrPoSeq
	Dim iStrPoQty, iStrPoUnit, iStrPoDt, iStrRcptQty, iStrTracking_no
	Dim iStrPlantCd, iStrSpplCd

	With frm1
		If .vspdData.MaxRows <= 0 Then
			Exit Function
		End If

		.vspdData1.ReDraw = False

		If IsNumeric(Trim(pvRowCnt)) Then
			imRow = CInt(pvRowCnt)
		Else
			imRow = AskSpdSheetAddRowCount()
			If imRow = "" Then
				Exit Function
			End If
		End If

		.vspdData1.focus
		ggoSpread.Source = .vspdData1
		ggoSpread.InsertRow .vspdData1.ActiveRow, imRow


		lRow = .vspdData.ActiveRow
		.vspdData.Row = lRow
		.vspdData.Col = .vspdData.MaxCols:		lconvRow = CInt(.vspdData.Text)
		.vspdData.Col = C_PrNo:					iStrPrNo = .vspdData.Text
		.vspdData.Col = C_PoNo:					iStrPoNo = .vspdData.Text
		.vspdData.Col = C_PoSeq:				iStrPoSeq = .vspdData.Text
		.vspdData.Col = C_PoQty:				iStrPoQty = Unicdbl(.vspdData.Text)
		.vspdData.Col = C_PoUnit:				iStrPoUnit = .vspdData.Text
		.vspdData.Col = C_PoDt:					iStrPoDt = .vspdData.Text
		.vspdData.Col = C_RcptQty:				iStrRcptQty = UniCdbl(.vspdData.Text)
		.vspdData.Col = C_TrackingNo:			iStrTracking_no = .vspdData.Text
		.vspdData.Col = C_PlantCd:				iStrPlantCd = .vspdData.Text
		.vspdData.Col = C_SpplCd:				iStrSpplCd = .vspdData.Text
		.vspdData.Col = C_ParentRowNo:			iParentRowNo = .vspdData.Text

		For iInsertRow = 0 To imRow - 1
			lRow2 = .vspdData1.ActiveRow + iInsertRow

			.vspdData1.Row = lRow2
			.vspdData1.Col = 0:						strMark = .vspdData1.Text
			.vspdData1.Col = C_ReservDt:			.vspdData1.value = iStrPoDt
			.vspdData1.Col = C_ReservQty:			.vspdData1.Text = 0
			.vspdData1.Col = C_BkQty:				.vspdData1.Text = 0
			.vspdData1.Col = C_IssueQty:			.vspdData1.Text = 0
			.vspdData1.Col = C_pPrNo:				.vspdData1.Text = iStrPrNo
			.vspdData1.Col = C_pPoNo:				.vspdData1.Text = iStrPoNo
			.vspdData1.Col = C_pPoSeq:				.vspdData1.Text = iStrPoSeq
			.vspdData1.Col = C_pPoQty:				.vspdData1.Text = iStrPoQty
			.vspdData1.Col = C_pPoUnit:				.vspdData1.Text = iStrPoUnit
			.vspdData1.Col = C_pPoDt:				.vspdData1.Text = iStrPoDt
			.vspdData1.Col = C_pRcptQty:			.vspdData1.Text = iStrRcptQty
			.vspdData1.Col = C_pTracking_no:		.vspdData1.Text = iStrTracking_no
			.vspdData1.Col = C_pPlantCd:			.vspdData1.Text = iStrPlantCd
			.vspdData1.Col = C_pSpplCd:				.vspdData1.Text = iStrSpplCd
			.vspdData1.Col = C_ParentRowNo:			.vspdData1.Text = lconvRow
			.vspdData1.Col = C_ChildRowNo:			.vspdData1.Text = strMark

			'2005-05-16 유/무상 무상으로 Default Setting
			.vspdData1.Col = C_SpplTypeCd:			.vspdData1.Text = "F"
			.vspdData1.Col = C_SpplTypeNm:			.vspdData1.Text = "무상"

			'재쿼리를 위해 해당 키에 대한 Client의 Data Row수를 가져감 
			lglngHiddenRows(lconvRow - 1) = CInt(lglngHiddenRows(lconvRow - 1)) + 1

			Call SetSpreadColor(lRow2, lRow2)
		Next

		'/* 수정 : 행헤더 재 넘버링 로직 추가 START */
		Dim i
		Dim lngRangeFrom
		Dim lngRangeTo
		Dim strFlag
		Dim k

		ggoSpread.Source = .vspdData1

		lngRangeFrom = ShowDataFirstRow2()
		lngRangeTo = ShowDataLastRow2()
		k = 0

		for i = lngRangeFrom To lngRangeTo
			k = k + 1
			strFlag = GetSpreadText(frm1.vspdData1,0,i,"X","X")

			If strFlag <> ggoSpread.InsertFlag and strFlag <> ggoSpread.UpdateFlag and strFlag <> ggoSpread.DeleteFlag then
				Call .vspdData1.SetText(0,	i,	CStr(k))
			End If
		Next
	End With

	Call frm1.vspdData.SetText(0,	.ActiveRow,		ggoSpread.UpdateFlag)

	.vspdData1.ReDraw = True

	FncInsertRow = true

	Set gActiveSpdSheet = document.activeElement
End Function
'===================================================================================================================================
Function FncCancel()
	FncCancel = false
	Dim lRow
	Dim i,k,iCnt
	Dim lngRangeFrom
	Dim lngRangeTo
	Dim iActiveRow
	Dim iConvActiveRow
	Dim strFlag

	iActiveRow = frm1.vspdData.ActiveRow
	frm1.vspdData.Row = iActiveRow
	frm1.vspdData.Col = frm1.vspdData.MaxCols
	iConvActiveRow = frm1.vspdData.Text

	If frm1.vspdData.MaxRows < 1 then
		Exit function
	End If

	if isEmpty(gActiveSpdSheet) then
		Set gActiveSpdSheet = frm1.vspdData1
	end if

	If gActiveSpdSheet.ID = "B" Then

		'Check Spread2 Data Exists for the keys
		If CheckDataExist = False Then
			Exit function
		End If

		ggoSpread.Source = frm1.vspdData1
		With frm1.vspdData1
			.Redraw = False
			lngRangeFrom = ShowDataFirstRow2()
			lngRangeTo = ShowDataLastRow2()
			ggoSpread.EditUndo

			lngRangeFrom = ShowDataFirstRow2()
			lngRangeTo = ShowDataLastRow2()

			If lngRangeFrom > 0 Then
				iCnt=1
				For k=lngRangeFrom To lngRangeTo
					.Row=k
					.col=0
					if Isnumeric(.text) or Trim(.text)="" Then .text=iCnt
					iCnt = iCnt + 1
				Next
			End If
			Call InitData(1)
			.Redraw = True
		End With
	Else
		ggoSpread.Source = frm1.vspdData
		ggoSpread.EditUndo												  '☜: Protect system from crashing

		if frm1.vspdData1.maxrowS > 0 Then
			ggoSpread.Source = frm1.vspdData1
			With frm1.vspdData1
				.Redraw = False

				lngRangeFrom = ShowDataFirstRow2()
				lngRangeTo = ShowDataLastRow2()
				'msgbox lngRangeFrom &  ", " & lngRangeTo
				If lngRangeFrom > 0 Then
					For k=lngRangeTo to lngRangeFrom step -1
						.Row=k
						.Col= 0
						If .Text = ggoSpread.InsertFlag or .Text = ggoSpread.UpdateFlag or .Text = ggoSpread.DeleteFlag then
							ggoSpread.EditUndo k												 '☜: Protect system from crashing
						End If
						Call InitData(1)
					Next
					iCnt=1
					For k=lngRangeFrom To lngRangeTo
						.Row=k
						.col=0
						if Isnumeric(.text) or Trim(.text)="" Then .text=iCnt
						iCnt = iCnt + 1
					Next
				End If
				.Redraw = True
			End WIth
		End If
	End If

	lRow = iActiveRow
	lngRangeFrom = ShowDataFirstRow2()
	lngRangeTo = ShowDataLastRow2()
	If lngRangeTo = 0 Then
		lglngHiddenRows(lRow - 1) = 0
	Else
		lglngHiddenRows(lRow - 1) = CInt(lngRangeTo) - CInt(lngRangeFrom) + 1
	End If

	k = 0
	If lngRangeFrom > 0 Then
		for i = lngRangeFrom to lngRangeTo
			frm1.vspdData1.Row = i
			frm1.vspdData1.Col = 0
			strFlag = Trim(frm1.vspdData1.Text)
			If strFlag = ggoSpread.InsertFlag or strFlag = ggoSpread.UpdateFlag or strFlag = ggoSpread.DeleteFlag then
				k = 1
				Exit for
			End If
		next
	End If

	if k = 0 then
		frm1.vspdData.Row = lRow
		frm1.vspdData.Col = 0
		frm1.vspdData.value = ""
	End If

	FncCancel = true

	Set gActiveSpdSheet = frm1.vspdData1
End Function
'===================================================================================================================================
Function FncPrint()
	FncPrint = False
	Call Parent.FncPrint()
	FncPrint = True
End Function
'===================================================================================================================================
Function FncExcel()
	FncExcel = False
 	Call parent.FncExport(Parent.C_MULTI)
 	FncExcel = True
End Function
'===================================================================================================================================
Function FncFind()
	FncFind = False
	Call parent.FncFind(Parent.C_MULTI, False)										 '☜:화면 유형, Tab 유무 
	FncFind = True
End Function
'===================================================================================================================================
Function FncExit()
	FncExit = False

	Dim IntRetCD

	If ChangeCheck = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")			'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	FncExit = True
End Function
'===================================================================================================================================
Function DbQuery()
	Dim LngLastRow
	Dim LngMaxRow
	Dim LngRow
	Dim strTemp
	Dim StrNextKey
	Dim pP21018		 'As New P21018ListIndReqSvr

	DbQuery = False

	If LayerShowHide(1) = False Then Exit Function

	Err.Clear															   '☜: Protect system from crashing

	Dim strVal

	With frm1

		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtSupplierCd=" & .hdnSupplier.value
			strVal = strVal & "&txtPlantCd=" & .hdnPlant.value
			strVal = strVal & "&txtPoNo=" & .hdnPoNo.value
			strVal = strVal & "&txtPrFrDt=" & .hdnPrFrDt.Value
			strVal = strVal & "&txtPrToDt=" & .hdnPrToDt.Value
			strVal = strVal & "&txtTrackNo=" & .hdnTrackNo.Value
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtSupplierCd=" & Trim(.txtSupplierCd.value)
			strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantcd.value)
			strVal = strVal & "&txtPoNo=" & Trim(.txtPoNo.value)
			strVal = strVal & "&txtPrFrDt=" & Trim(.txtPrFrDt.text)
			strVal = strVal & "&txtPrToDt=" & Trim(.txtPrToDt.text)
			strVal = strVal & "&txtTrackNo=" & Trim(.txtTrackNo.Value)
		End If

		strVal = strVal & "&lgPageNo="   & lgPageNo					  '☜: Next key tag
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows

		Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

	End With

	DbQuery = True

End Function
'===================================================================================================================================
Function DbQueryOk(byVal intARow,byVal intTRow)														'☆: 조회 성공후 실행로직 
	Dim i, lRow
	Dim TmpArrPrevKey
	Dim TmpArrHiddenRows
	'-----------------------
	'Reset variables area
	'-----------------------

	Call ggoOper.LockField(Document, "N")									'⊙: This function lock the suitable field
	Call SetSpreadLock("A")

	With frm1
		lRow = .vspdData.MaxRows

		If lRow > 0 And intARow > 0 Then
			ReDim lgStrPrevKeyM(lRow - 1)

			If intTRow<=0 Then
				ReDim lglngHiddenRows(intARow - 1)			'lRow = .vspdData.MaxRows	'ex) 첫번째 그리드의 특정Row에 해당하는 두번째 그리드의 Row 갯수를 저장하는 배열.
			Else
				TmpArrHiddenRows=lglngHiddenRows

				ReDim lglngHiddenRows(intTRow+intARow - 1)			'lRow = .vspdData.MaxRows	'ex) 첫번째 그리드의 특정Row에 해당하는 두번째 그리드의 Row 갯수를 저장하는 배열.
				For i = 0 To intTRow-1
					lglngHiddenRows(i) = TmpArrHiddenRows(i)
				Next
			End If

			For i = intTRow To intTRow+intARow-1
				lglngHiddenRows(i) = 0
			Next

			if lgIntFlgModeM = Parent.OPMD_CMODE then
				If DbQuery2(1, False) = False Then	Exit Function
			end if
			lgIntFlgModeM = Parent.OPMD_UMODE
		End If
	End With
	DbQueryOk = true

End Function

Sub RemovedivTextArea()
	Dim ii

	For ii = 1 To divTextArea.children.length
		divTextArea.removeChild(divTextArea.children(0))
	Next
End Sub

'===================================================================================================================================
Function DbSave()

	Dim lRow , parentRow , pRow
	Dim lGrpCnt
	Dim strVal
	Dim strDel
	Dim strVal1
	Dim strVal2
	Dim intIndex
	dim poDt, reservDt
	dim prSts
	Dim lgTransSep
	Dim pTmpPrNo
	Dim pTmpChildRowNo

	Dim lngRangeFrom
	Dim lngRangeTo
	Dim PvArr
	Dim iColSep, iRowSep

	Dim strCUTotalvalLen '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규]
	Dim strDTotalvalLen  '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[삭제]

	Dim objTEXTAREA '동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 

	Dim iTmpCUBuffer		 '현재의 버퍼 [수정,신규]
	Dim iTmpCUBufferCount	'현재의 버퍼 Position
	Dim iTmpCUBufferMaxCount '현재의 버퍼 Chunk Size

	Dim iTmpDBuffer		  '현재의 버퍼 [삭제]
	Dim iTmpDBufferCount	 '현재의 버퍼 Position
	Dim iTmpDBufferMaxCount  '현재의 버퍼 Chunk Size

	DbSave = False														  '⊙: Processing is NG

	With frm1

		.txtMode.value = Parent.UID_M0002

		lGrpCnt = 0
		strVal = ""
		strDel = ""
		strVal2 = ""
		lgTransSep = "Ð"
		pTmpChildRowNo = 0
		iColSep = Parent.gColSep
		iRowSep	= Parent.gRowSep

		iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[수정,신규]
		iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[삭제]

		ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '최기 버퍼의 설정[수정,신규]
		ReDim iTmpDBuffer(iTmpDBufferMaxCount)  '최기 버퍼의 설정[수정,신규]

		iTmpCUBufferCount = -1
		iTmpDBufferCount = -1

		strCUTotalvalLen = 0
		strDTotalvalLen  = 0

		If LayerShowHide(1) = False Then Exit Function

		For parentRow = 1 To .vspdData.MaxRows

			.vspdData.Row = parentRow
			.vspdData.Col = 0

			if Trim(.vspdData.text) = ggoSpread.UpdateFlag then

				pRow = GetSpreadText(frm1.vspdData,.vspdData.MaxCols,parentRow,"X","X")
				pTmpPrNo = UCase(Trim(GetSpreadText(frm1.vspdData,C_PrNo,parentRow,"X","X")))

				lngRangeFrom = DataFirstRow(pRow)
				lngRangeTo   = DataLastRow(pRow)

				For lRow =  lngRangeFrom To lngRangeTo

					pTmpChildRowNo = pTmpChildRowNo + 1

					.vspdData1.Row = lRow
					Select Case Trim(GetSpreadText(frm1.vspdData1,0,lRow,"X","X"))
						Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag

							If UniCdbl(GetSpreadText(frm1.vspdData1,C_ReservQty,lRow,"X","X")) = "" Or UniCdbl(GetSpreadText(frm1.vspdData1,C_ReservQty,lRow,"X","X")) = 0 then
								Call DisplayMsgBox("970021", "X",parentRow & ": " & "행 " & "요청번호" & "[" & pTmpPrNo & "]" & " : " & pTmpChildRowNo & "행" & ", " & "출고예정량", "X")
								Call LayerShowHide(0)
								Call setFocusRow(parentRow,lRow,C_ReservQty)
								Call RemovedivTextArea
								Exit Function
							End if

							poDt = GetSpreadText(frm1.vspdData,C_PoDt,C_ParentRowNo,"X","X")
							reservDt = GetSpreadText(frm1.vspdData1,C_ReservDt,lRow,"X","X")

							If ValidDateCheckLocal(parentRow, pTmpChildRowNo, pTmpPrNo, poDt, reservDt) Then	'출고예정일 >= 발주일 check
								Call LayerShowHide(0)
								Call setFocusRow(parentRow,IRow,C_ReservDt)
								Call RemovedivTextArea
								Exit Function
							End If

							'prSts = UCase(Trim(GetSpreadText(frm1.vspdData,C_PrStateCd,parentRow,"X","X")))
							If UniCdbl(GetSpreadText(frm1.vspdData1,C_ReservQty,lRow,"X","X")) < UniCdbl(GetSpreadText(frm1.vspdData1,C_IssueQty,lRow,"X","X")) then
								Call DisplayMsgBox("970023","X", parentRow & ": " & "행 " & "요청번호" & "[" & pTmpPrNo & "]" & " : " & pTmpChildRowNo & "행" & ", " & "출고예정량", "출고수량")
								Call LayerShowHide(0)
								Call setFocusRow(parentRow,IRow,C_ReservQty)
								Call RemovedivTextArea
								Exit Function
							End if

							'2005-05-16 단위정보가 입력되어 있는지 Check
							If Trim(GetSpreadText(frm1.vspdData1,C_IssueUnit,lRow,"X","X")) = "" Then
'								Call DisplayMsgBox("970023","X", parentRow & ": " & "행 " & "요청번호" & "[" & pTmpPrNo & "]" & " : " & pTmpChildRowNo & "행" & ", " & "출고예정량", "출고수량")
								Call DisplayMsgBox("124010", "X", parentRow & ": " & "행 " & "요청번호" & "[" & pTmpPrNo & "]" & " : " & pTmpChildRowNo & "행" & ", " & "자품목", "X")
								Call LayerShowHide(0)
								Call setFocusRow(parentRow,IRow,C_ReservQty)
								Call RemovedivTextArea
								Exit Function
							End if

							If Trim(GetSpreadText(frm1.vspdData1,0,lRow,"X","X")) = ggoSpread.InsertFlag Then
								strVal = strVal & "C" & iColSep
							ElseIf Trim(GetSpreadText(frm1.vspdData1,0,lRow,"X","X")) = ggoSpread.UpdateFlag Then
								strVal = strVal & "U" & iColSep
							End If

							.vspdData1.Row = lRow
							.vspdData1.Col = C_ChildItemCd:		strVal = strVal & UCase(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_ChildItemSpec:	strVal = strVal & UCase(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_SpplTypeCd:		strVal = strVal & UCase(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_IssueSlCd:		strVal = strVal & UCase(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_ReservDt:		strVal = strVal & UNIConvDate(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_ReservQty:		strVal = strVal & UNIConvNum(Trim(.vspdData1.Text),0) & iColSep
							.vspdData1.Col = C_BkQty:			strVal = strVal & UNIConvNum(Trim(.vspdData1.Text),0) & iColSep
							.vspdData1.Col = C_IssueQty:		strVal = strVal & UNIConvNum(Trim(.vspdData1.Text),0) & iColSep
							.vspdData1.Col = C_IssueUnit:		strVal = strVal & UCase(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_PrStateCd:		strVal = strVal & UCase(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_HisSubSeqNo:		strVal = strVal & Trim(.vspdData1.Text) & iColSep
							.vspdData1.Col = C_ReqmtNo:			strVal = strVal & UCase(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_pPrNo:			strVal = strVal & UCase(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_ResvdSeqNo:		strVal = strVal & UCase(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_pPoNo:			strVal = strVal & UCase(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_pPoSeq:			strVal = strVal & UCase(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_pPoQty:			strVal = strVal & UNIConvNum(Trim(.vspdData1.Text),0) & iColSep
							.vspdData1.Col = C_pPoUnit:			strVal = strVal & UCase(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_pPoDt:			strVal = strVal & UNIConvDate(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_pRcptQty:		strVal = strVal & UNIConvNum(Trim(.vspdData1.Text),0) & iColSep
							.vspdData1.Col = C_pTracking_no:	strVal = strVal & UCase(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_pPlantCd:		strVal = strVal & UCase(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_pSpplCd:			strVal = strVal & UCase(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_OrgChildItemCd:	strVal = strVal & UCase(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_OrgSpplTypeCd:	strVal = strVal & UCase(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_OrgSlCd:			strVal = strVal & UCase(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_OrgReservQty:	strVal = strVal & UNIConvNum(Trim(.vspdData1.Text),0) & iColSep
							.vspdData1.Col = C_OrgReservDt:		strVal = strVal & UNIConvDate(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_ParentRowNo:		strVal = strVal & Trim(.vspdData1.Text) & iColSep
							.vspdData1.Col = C_ChildRowNo:		strVal = strVal & Trim(.vspdData1.Text) & iColSep

							strVal = strVal & lRow & iRowSep
					 		lGrpCnt = lGrpCnt + 1

					 	Case ggoSpread.DeleteFlag
							strDel = strDel & "D" & iColSep
							.vspdData1.Col = C_ChildItemCd:		strDel = strDel & UCase(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_ChildItemSpec:	strDel = strDel & UCase(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_SpplTypeCd:		strDel = strDel & UCase(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_IssueSlCd:		strDel = strDel & UCase(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_ReservDt:		strDel = strDel & UNIConvDate(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_ReservQty:		strDel = strDel & UNIConvNum(Trim(.vspdData1.Text),0) & iColSep
							.vspdData1.Col = C_BkQty:			strDel = strDel & UNIConvNum(Trim(.vspdData1.Text),0) & iColSep
							.vspdData1.Col = C_IssueQty:		strDel = strDel & UNIConvNum(Trim(.vspdData1.Text),0) & iColSep
							.vspdData1.Col = C_IssueUnit:		strDel = strDel & UCase(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_PrStateCd:		strDel = strDel & UCase(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_HisSubSeqNo:		strDel = strDel & Trim(.vspdData1.Text) & iColSep
							.vspdData1.Col = C_ReqmtNo:			strDel = strDel & UCase(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_pPrNo:			strDel = strDel & UCase(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_ResvdSeqNo:		strDel = strDel & UCase(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_pPoNo:			strDel = strDel & UCase(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_pPoSeq:			strDel = strDel & UCase(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_pPoQty:			strDel = strDel & UNIConvNum(Trim(.vspdData1.Text),0) & iColSep
							.vspdData1.Col = C_pPoUnit:			strDel = strDel & UCase(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_pPoDt:			strDel = strDel & UNIConvDate(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_pRcptQty:		strDel = strDel & UNIConvNum(Trim(.vspdData1.Text),0) & iColSep
							.vspdData1.Col = C_pTracking_no:	strDel = strDel & UCase(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_pPlantCd:		strDel = strDel & UCase(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_pSpplCd:			strDel = strDel & UCase(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_OrgChildItemCd:	strDel = strDel & UCase(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_OrgSpplTypeCd:	strDel = strDel & UCase(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_OrgSlCd:			strDel = strDel & UCase(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_OrgReservQty:	strDel = strDel & UNIConvNum(Trim(.vspdData1.Text),0) & iColSep
							.vspdData1.Col = C_OrgReservDt:		strDel = strDel & UNIConvDate(Trim(.vspdData1.Text)) & iColSep
							.vspdData1.Col = C_ParentRowNo:		strDel = strDel & Trim(.vspdData1.Text) & iColSep
							.vspdData1.Col = C_ChildRowNo:		strDel = strDel & Trim(.vspdData1.Text) & iColSep
							strDel = strDel & lRow & iRowSep
					 		lGrpCnt = lGrpCnt + 1
					End Select
				Next

				strVal =  strDel & strVal & lgTransSep
				If strCUTotalvalLen + Len(strVal) >  parent.C_FORM_LIMIT_BYTE Then  '한개의 form element에 넣을 Data 한개치가 넘으면 

				   Set objTEXTAREA = document.createElement("TEXTAREA")				 '동적으로 한개의 form element를 동저으로 생성후 그곳에 데이타 넣음 
				   objTEXTAREA.name = "txtCUSpread"
				   objTEXTAREA.value = Join(iTmpCUBuffer,"")
				   divTextArea.appendChild(objTEXTAREA)

				   iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT				  ' 임시 영역 새로 초기화 
				   ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
				   iTmpCUBufferCount = -1
				   strCUTotalvalLen  = 0
				End If

				iTmpCUBufferCount = iTmpCUBufferCount + 1

				If iTmpCUBufferCount > iTmpCUBufferMaxCount Then							  '버퍼의 조정 증가치를 넘으면 
				   iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT '버퍼 크기 증성 
				   ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
				End If
				iTmpCUBuffer(iTmpCUBufferCount) =  strVal
				strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
			End If
			strVal  = ""
			strDel  = ""
		Next
		.txtMaxRows.value = lGrpCnt-1
		If iTmpCUBufferCount > -1 Then   ' 나머지 데이터 처리 
		   Set objTEXTAREA = document.createElement("TEXTAREA")
		   objTEXTAREA.name   = "txtCUSpread"
		   objTEXTAREA.value = Join(iTmpCUBuffer,"")
		   divTextArea.appendChild(objTEXTAREA)
		End If

		Call ExecMyBizASP(frm1, BIZ_PGM_ID_01)

	End With

	DbSave = True

End Function
'===================================================================================================================================
Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
	Call InitVariables
	Call MainQuery()
End Function
'===================================================================================================================================
Function SheetFocus(lRow, lCol)
	frm1.vspdData1.focus
	frm1.vspdData1.Row = lRow
	frm1.vspdData1.Col = lCol
	frm1.vspdData1.Action = 0
	frm1.vspdData1.SelStart = 0
	frm1.vspdData1.SelLength = len(frm1.vspdData1.Text)
End Function
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>사급소요량조정</font></td>
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
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET>
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>공급처</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공급처" NAME="txtSupplierCd" SIZE=10 MAXLENGTH=10 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSupplier()">
														   <INPUT TYPE=TEXT ALT="공급처" NAME="txtSupplierNm" SIZE=20 tag="14X"></TD>
									<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공장" NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant() ">
														   <INPUT TYPE=TEXT ALT="공장" NAME="txtPlantNm" SIZE=20 tag="14X"></TD>
								</TR>
								<tr>
									<TD CLASS="TD5" NOWRAP>발주번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPoNo" SIZE=32  MAXLENGTH=18 ALT="발주번호" tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPoNo()"></TD>
									<TD CLASS="TD5" NOWRAP>발주일</TD>
									<TD CLASS="TD6" NOWRAP>
										<table cellspacing=0 cellpadding=0>
											<tr>
												<td>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=요청일 NAME="txtPrFrDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 CLASS=FPDTYYYYMMDD tag="11N1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</td>
												<td>~</td>
												<td>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=요청일 NAME="txtPrToDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 CLASS=FPDTYYYYMMDD tag="11N1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</td>
											<tr>
										</table>
									</TD>
								</tr>
								<tr>
									<TD CLASS="TD5" NOWRAP>Tracking No.</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="Tracking No." NAME="txtTrackNo" SIZE=34 MAXLENGTH=25  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenTrackingNo()"></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
								</tr>
							</TABLE>
						</FIELDSET>
					</TD>
			</TR>
			<TR>
				<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
			</TR>
			<TR>
				<TD WIDTH=100% valign=top>
					<TABLE <%=LR_SPACE_TYPE_20%>>
						<TR>
							<TD HEIGHT="100%">
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id="A"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
							</TD>
						</TR>
					</TABLE>
				</TD>
			</TR>
			<TR>
				<TD WIDTH=100% valign=top>
					<TABLE <%=LR_SPACE_TYPE_20%>>
						<TR>
							<TD HEIGHT="100%">
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData1 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id="B"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
							</TD>
						</TR>
					</TABLE>
				</TD>
			</TR>
		</TABLE></TD>
	</TR>
	<tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</tr>
	<tr HEIGHT="20">
		<td WIDTH="100%">
			<table <%=LR_SPACE_TYPE_30%>>
				<tr>
					<td WIDTH="*" ALIGN="RIGHT"><a href="VBSCRIPT:PgmJump(BIZ_PGM_JUMP_ID_FOR_PO)" ONCLICK="VBSCRIPT:setCookieForPo()">발주등록</a></td>
					<td WIDTH="20"></td>
				</tr>
			</table>
		</td>
	</tr>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<P ID="divTextArea"></P>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">

<INPUT TYPE=HIDDEN NAME="hdnSupplier" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPlant" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPoNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPrFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPrToDt" tag="24">

<INPUT TYPE=HIDDEN NAME="hdnQueryRow" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnLastRow" tag="14">

<INPUT TYPE=HIDDEN NAME="txtHPrNo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHPoNo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHPoSeq" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnTrackNo" tag="24">

</FORM>

	<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
	</DIV>

</BODY>
</HTML>
