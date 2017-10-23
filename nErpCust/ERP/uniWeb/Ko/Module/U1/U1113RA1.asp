<%@ LANGUAGE="VBSCRIPT" %>
<!--
'********************************************************************************************************
'*  1. Module Name          : Procurement																*
'*  2. Function Name        :																			*
'*  3. Program ID           : U1113RA1.asp      														*
'*  4. Program Name         : 반품예정참조(반품등록)													*
'*  5. Program Desc         : 구매반품에서 반품예정참조													*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2004/08/11																*
'*  8. Modified date(Last)  : 2004/08/11																*
'*  9. Modifier (First)     :	NHG 																	*
'* 10. Modifier (Last)      :	NHG																		*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'********************************************************************************************************
-->
<HTML>
<HEAD>
<TITLE></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<Script Language="VBScript">

Option Explicit		

'================================================================================================================================
Const BIZ_PGM_ID 		= "U1113RB1.asp"
'================================================================================================================================
Dim C_PoNo
Dim C_PoSeqNo
Dim C_PlantCd
Dim C_SLCd
Dim C_ItemCd
Dim C_ItemNm
Dim C_Spec
Dim C_TrackingNo
Dim C_POQty
Dim C_Unit
Dim C_POPrc
Dim C_POAmt
Dim C_POCur
Dim C_PODlvyDt
Dim C_GRQty
Dim C_LCQty
Dim C_PreIvQty
Dim C_InspectQty
Dim C_IvQty
Dim C_InspFlg
Dim C_InspMeth
Dim C_InspMethCd
Dim C_PlantNm
Dim C_SLNm
Dim C_Pur_Grp
Dim C_LCRCPTQTY
Dim C_Lot_flg
Dim C_Lot_gen_mtd
Dim C_MakerLotNo
Dim	C_MakerLotSeqNo
Dim	C_PlanDvryDt
Dim C_PlanDvryQty
Dim C_SplitSeqNo

'================================================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->
'================================================================================================================================

'================================================================================================================================
Const C_MaxKey          = 28                                           '☆: key count of SpreadSheet
Dim gblnWinEvent
Dim arrReturn					<% '--- Return Parameter Group %>
Dim arrParent
Dim arrParam
Dim EndDate, StartDate
'================================================================================================================================    
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
top.document.title = PopupParent.gActivePRAspName
arrParam= arrParent(1)

EndDate = UNIConvDateAtoB("<%=GetSvrDate%>", PopupParent.gServerDateFormat, PopupParent.gDateFormat)
StartDate = UnIDateAdd("d", -7, EndDate, PopupParent.gDateFormat)
'================================================================================================================================
Function InitVariables()
    lgPageNo         = ""
	lgBlnFlgChgValue = False
    lgIntFlgMode     = PopupParent.OPMD_CMODE                        'Indicates that current mode is Create mode
    lgSortKey        = 1
						
	frm1.vspdData.MaxRows = 0	
	
	gblnWinEvent = False
	ReDim arrReturn(0,0)
	Self.Returnvalue = arrReturn
End Function
'================================================================================================================================
Sub SetDefaultVal()
	
	Dim iCodeArr
		
	Err.Clear
	
	With frm1
		.txtFrPoDt.text = StartDate
		.txtToPoDt.text = EndDate
		.txtBpCd.value = arrParam(0)
		Call CommonQueryRs(" BP_NM", " B_BIZ_PARTNER", " BP_CD = '" & FilterVar(Trim(.txtBpCd.value),"","SNM") & "'", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		.txtBpNm.value = Replace(lgF0, Chr(11),"")
		.hdnSupplierCd.value 	= arrParam(0)
		.hdnGroupCd.value 		= arrParam(2)
		.hdnGroupNm.value 		= arrParam(3)
		.hdnRefType.value 		= arrParam(8)
		.hdnRcptType.value 		= arrParam(9)
		
		.txtFrPoDt.Text			= arrParam(10)
		.txtToPoDt.Text			= arrParam(10)
		
		'.HDNPlantCd.value		= PopupParent.gPlant
		'.HDNPlantNm.value		= PopupParent.gPlantNm
		
		.hdnDistinctNo.value	= arrParam(13)
		
	End With
	
	Call CommonQueryRs(" RCPT_FLG", " M_MVMT_TYPE", " IO_TYPE_CD = '" & FilterVar(Trim(frm1.hdnRcptType.value),"","SNM") & "'", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    IF Len(lgF0) Then
		iCodeArr = Split(lgF0, Chr(11))
		    
		If Err.number <> 0 Then
			MsgBox Err.description,vbInformation,PopupParent.gLogoName 
			Err.Clear 
			Exit Sub
		End If
		frm1.hdnRcptFlg.value 	= iCodeArr(0)
	End if	
	
End Sub
'================================================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "RA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "RA") %>
End Sub
'================================================================================================================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()

	'------------------------------------------
	' Grid 1 - Operation Spread Setting
	'------------------------------------------
	With frm1.vspdData 
			
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit
		frm1.vspdData.OperationMode = 5
		
		.ReDraw = false
					
		.MaxCols = C_SplitSeqNo + 1    
		.MaxRows = 0    
			
		Call GetSpreadColumnPos()

		ggoSpread.SSSetEdit 		C_PoNo,			"발주번호", 20
		ggoSpread.SSSetFloat 		C_PoSeqNo,		"발주순번",10,"6",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,0
		ggoSpread.SSSetEdit 		C_PlantCd,		"공장", 10
		ggoSpread.SSSetEdit			C_SlCd,			"창고", 10				
		ggoSpread.SSSetEdit 		C_ItemCd,		"품목",18
		ggoSpread.SSSetEdit 		C_ItemNm,		"품목명",20
		ggoSpread.SSSetEdit 		C_Spec,			"규격",20
		ggoSpread.SSSetEdit 		C_TrackingNo,	"Tracking No.", 15 	
		ggoSpread.SSSetFloat 		C_POQty,		"발주수량",15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit 		C_Unit,			"단위", 10
		ggoSpread.SSSetFloat		C_PoPrc,		"단가",			15,		Popupparent.ggUnitCostNo,  ggStrIntegeralPart,	ggStrDeciPointPart,	Popupparent.gComNum1000,Popupparent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat		C_PoAmt,		"발주금액",			15,		Popupparent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	Popupparent.gComNum1000,	Popupparent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetEdit 		C_POCur,	    "화폐", 10
		ggoSpread.SSSetDate 		C_PODlvyDt,		"납기일"	,10, 2, PopupParent.gDateFormat		 
		ggoSpread.SSSetFloat 		C_GRQty,		"입고량",15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat 		C_LCQty,		"L/C수량",15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat 		C_PreIvQty,		"선매입수량",15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat 		C_InspectQty,	"검사중수량",15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat 		C_IvQty,		"매입수량",15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
		ggoSpread.SSSetCheck 		C_InspFlg,		"검사품유무",10,,,true
		ggoSpread.SSSetEdit 		C_InspMeth,		"검사방법명",10
		ggoSpread.SSSetEdit 		C_InspMethCd,	"검사방법", 10
		ggoSpread.SSSetEdit 		C_PlantNm,		"공장명", 20
		ggoSpread.SSSetEdit 		C_SlNm,			"창고명", 20	    		
		ggoSpread.SSSetEdit 		C_Pur_Grp,		"구매그룹", 20	    		
		ggoSpread.SSSetFloat 		C_LCRCPTQTY,	"LC분에 대한 입고수량",15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit 		C_Lot_flg,		"LOT관리여부", 20	    		
		ggoSpread.SSSetEdit 		C_Lot_gen_mtd,	"LOT생성방법", 20	    		
		ggoSpread.SSSetEdit 		C_MakerLotNo,	"MAKER LOT NO.", 20,,,12,2    
		ggoSpread.SSSetFloat 		C_MakerLotSeqNo,"Maker 순번", 10,"6",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,0
		ggoSpread.SSSetFloat 		C_PlanDvryQty,	"납품수량",15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
		ggoSpread.SSSetDate 		C_PlanDvryDt,	"납품일"	,10, 2, PopupParent.gDateFormat
		ggoSpread.SSSetEdit 		C_SplitSeqNo,	"분할번호", 4	    		
		
		Call ggoSpread.SSSetColHidden( C_SplitSeqNo, C_SplitSeqNo, True)	
		Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols, True)
			
		ggoSpread.SSSetSplit2(2)
						
		Call SetSpreadLock()
						
		.ReDraw = true    
    
	End With
	   
End Sub
'================================================================================================================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'================================================================================================================================
Sub InitSpreadPosVariables()
	C_PoNo			= 1
	C_PoSeqNo		= 2
	C_PlantCd		= 3
	C_SLCd			= 4
	C_ItemCd		= 5
	C_ItemNm		= 6
	C_Spec			= 7
	C_TrackingNo	= 8
	C_POQty			= 9
	C_Unit			= 10
	C_POPrc			= 11
	C_POAmt			= 12
	C_POCur			= 13
	C_PODlvyDt		= 14
	C_GRQty			= 15
	C_LCQty			= 16
	C_PreIvQty		= 17
	C_InspectQty	= 18
	C_IvQty			= 19
	C_InspFlg		= 20
	C_InspMeth		= 21
	C_InspMethCd	= 22
	C_PlantNm		= 23
	C_SLNm			= 24
	C_Pur_Grp		= 25
	C_LCRCPTQTY		= 26
	C_Lot_flg		= 27
	C_Lot_gen_mtd	= 28
	C_MakerLotNo	= 29
	C_MakerLotSeqNo	= 30
	C_PlanDvryDt	= 31
	C_PlanDvryQty	= 32
	C_SplitSeqNo	= 33
	
End Sub
'================================================================================================================================
Sub GetSpreadColumnPos()
      
    Dim iCurColumnPos
    
 	ggoSpread.Source = frm1.vspdData
		
	Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

	C_PoNo			= iCurColumnPos(1)
	C_PoSeqNo		= iCurColumnPos(2)
	C_PlantCd		= iCurColumnPos(3)
	C_SLCd			= iCurColumnPos(4)
	C_ItemCd		= iCurColumnPos(5)
	C_ItemNm		= iCurColumnPos(6)
	C_Spec			= iCurColumnPos(7)
	C_TrackingNo	= iCurColumnPos(8)
	C_POQty			= iCurColumnPos(9)
	C_Unit			= iCurColumnPos(10)
	C_POPrc			= iCurColumnPos(11)
	C_POAmt			= iCurColumnPos(12)
	C_POCur			= iCurColumnPos(13)
	C_PODlvyDt		= iCurColumnPos(14)
	C_GRQty			= iCurColumnPos(15)
	C_LCQty			= iCurColumnPos(16)
	C_PreIvQty		= iCurColumnPos(17)
	C_InspectQty	= iCurColumnPos(18)
	C_IvQty			= iCurColumnPos(19)
	C_InspFlg		= iCurColumnPos(20)
	C_InspMeth		= iCurColumnPos(21)
	C_InspMethCd	= iCurColumnPos(22)
	C_PlantNm		= iCurColumnPos(23)
	C_SLNm			= iCurColumnPos(24)
	C_Pur_Grp		= iCurColumnPos(25)
	C_LCRCPTQTY		= iCurColumnPos(26)
	C_Lot_flg		= iCurColumnPos(27)
	C_Lot_gen_mtd	= iCurColumnPos(28)
	C_MakerLotNo	= iCurColumnPos(29)
	C_MakerLotSeqNo	= iCurColumnPos(30)
	C_PlanDvryDt	= iCurColumnPos(31)
	C_PlanDvryQty	= iCurColumnPos(32)
	C_SplitSeqNo	= iCurColumnPos(33)

End Sub    
'================================================================================================================================
Function OKClick()
	
	Dim intColCnt, intRowCnt, intInsRow

		If frm1.vspdData.SelModeSelCount > 0 Then 

			intInsRow = 0

			Redim arrReturn(frm1.vspdData.SelModeSelCount-1, frm1.vspdData.MaxCols - 2)

			For intRowCnt = 1 To frm1.vspdData.MaxRows

				frm1.vspdData.Row = intRowCnt

				If frm1.vspdData.SelModeSelected Then
								
					For intColCnt = 0 To frm1.vspdData.MaxCols - 2
						frm1.vspdData.Col = intColCnt+1 ' GetKeyPos("A",intColCnt+1)
						arrReturn(intInsRow, intColCnt) = frm1.vspdData.Text
					Next
					intInsRow = intInsRow + 1
				End IF								
			Next
			
		End if			
		Self.Returnvalue = arrReturn
		Self.Close()
End Function	
'================================================================================================================================
Function CancelClick()
	Redim arrReturn(1,1)
	arrReturn(0,0) = ""
	Self.Returnvalue = arrReturn
	Self.Close()
End Function

'================================================================================================================================
Function OpenSlcd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Or UCase(frm1.txtslCd.className) = Ucase(PopupParent.UCN_PROTECTED) Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "입고창고"	
	arrParam(1) = "(SELECT DISTINCT B.D_BP_CD,C.SL_NM          "
	arrParam(1) = arrParam(1) & " FROM M_PUR_ORD_HDR 	    A, "
	arrParam(1) = arrParam(1) & "       m_scm_firm_pur_rcpt B, "
	arrParam(1) = arrParam(1) & "       b_storage_location  C  "
	arrParam(1) = arrParam(1) & " WHERE A.PO_NO = B.PO_NO      " 
	arrParam(1) = arrParam(1) & "   AND B.D_BP_CD = C.SL_CD    "		
	arrParam(1) = arrParam(1) & "   AND A.BP_CD = " & FilterVar(Trim(frm1.txtBpCd.value),"''","S") & ") A "			
	
	arrParam(2) = Trim(frm1.txtSLCd.Value)
	arrParam(3) = Trim(frm1.txtSLNM.Value)	
	
	arrParam(4) = ""			
	arrParam(5) = "입고창고"
	
    arrField(0) = "D_BP_CD"	
    arrField(1) = "SL_NM"
    
    arrHeader(0) = "입고창고"		
    arrHeader(1) = "입고창고명"	
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	gblnWinEvent = False
	
	If arrRet(0) = "" Then
		frm1.txtslCd.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtslCd.Value= arrRet(0)		
		frm1.txtslNm.Value= arrRet(1)	
		frm1.txtslCd.focus	
		Set gActiveElement = document.activeElement
	End If	

End Function 

'================================================================================================================================
Function OpenSortPopup()

	
	On Error Resume Next
	
End Function
'================================================================================================================================
Function OpentxtDlvyNo()
	Dim arrRet,lgIsOpenPop
	Dim arrParam(5), arrField(6), arrHeader(6)

	Call CommonQueryRs(" RCPT_FLG ", " M_MVMT_TYPE ", " IO_TYPE_CD = '" & FilterVar(Trim(frm1.hdnRcptType.value),"","SNM") & "'", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "발행번호"	
	arrParam(1) = "(SELECT DISTINCT A.DLVY_NO , A.BP_CD , B.BP_NM   "
    arrParam(1) = arrParam(1) & " FROM M_SCM_DLVY_PUR_RCPT A ,B_BIZ_PARTNER B , M_SCM_FIRM_PUR_RCPT C , M_PUR_ORD_HDR D "
    arrParam(1) = arrParam(1) & " WHERE A.BP_CD = B.BP_CD AND D.RET_FLG  = 'Y' AND A.BP_CD = '" & Frm1.txtBpCd.value & "'" 
    arrParam(1) = arrParam(1) & "  AND C.RCPT_QTY = 0 "
    If Replace(lgF0, Chr(11),"") = "N" Then
		arrParam(1) = arrParam(1) & " AND D.ISSUE_TYPE = '" & frm1.hdnRcptType.value & "' "
	Else
		arrParam(1) = arrParam(1) & " AND D.RCPT_TYPE = '" & frm1.hdnRcptType.value & "' "
    End If
    
    arrParam(1) = arrParam(1) & " AND A.DLVY_NO = C.DLVY_NO AND C.PO_NO = D.PO_NO )a "
    
	arrParam(2) = Trim(frm1.txtDlvyNo.value)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "발행번호"			
	
    arrField(0) = "ED15" & PopupParent.gColSep & "Dlvy_No"	
    arrField(1) = "ED06" & PopupParent.gColSep & "BP_CD"
    arrField(2) = "ED20" & PopupParent.gColSep & "BP_NM"	
    
    arrHeader(0) = "발행번호"		
    arrHeader(1) = "공급처"
    arrHeader(2) = "공급처명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		
		Set gActiveElement = document.activeElement
		Exit Function
	Else	
		frm1.txtDlvyNo.value = arrRet(0)
		frm1.txtDlvyNo.focus	
		Set gActiveElement = document.activeElement
	End If	
		
End Function	
'================================================================================================================================
Sub Form_Load()
	
	Call LoadInfTB19029															'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)	                                           
	Call ggoOper.LockField(Document, "N")										'⊙: Lock  Suitable  Field 
	Call InitVariables														    '⊙: Initializes local global variables
	Call SetDefaultVal	
		
	Call InitSpreadSheet()
		
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")

	Call FncQuery()
End Sub
'================================================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	
	gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
	
	If frm1.vspdData.MaxRows = 0 Then Exit Sub
	   
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
			lgSortkey = 1
		End If
		
		Exit Sub
	End If    
End Sub
'================================================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Or Frm1.vspdData.MaxRows = 0 Then 
	     Exit Sub
	End If

	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Sub
'================================================================================================================================
Function vspdData_KeyPress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then    'Frm1없으면 frm1삭제 
        Call OKClick()
     ElseIf KeyAscii = 27 Then
        Call CancelClick()
     End If
End Function
'================================================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
	If OldLeft <> NewLeft Then
	    Exit Sub
	End If
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
		If lgPageNo <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then
				Exit Sub
			End if
		End If
	End If
End Sub
'================================================================================================================================
Sub txtFrPoDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub
'================================================================================================================================
Sub txtToPoDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub
'================================================================================================================================
Sub txtFrPoDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrPoDt.Action = 7
		Call SetFocusToDocument("P")	
        frm1.txtFrPoDt.Focus
	End if
End Sub
'================================================================================================================================
Sub txtToPoDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToPoDt.Action = 7
		Call SetFocusToDocument("P")	
        frm1.txtToPoDt.Focus
	End if
End Sub
'================================================================================================================================
Function FncQuery() 
    
    FncQuery = False                                                 
    
    Err.Clear                                                        
	
	With frm1
		if (UniConvDateToYYYYMMDD(.txtFrPoDt.text,PopupParent.gDateFormat,"") > UniConvDateToYYYYMMDD(.txtToPoDt.text,PopupParent.gDateFormat,"")) And trim(.txtFrPoDt.text) <> "" And trim(.txtToPoDt.text) <> "" then	
			Call DisplayMsgBox("17a003", "X","발주일", "X")	
			.txtToPoDt.Focus()
			Exit Function
		End if   
	End with
	
	ggoSpread.Source = frm1.vspdData	
    ggoSpread.ClearSpreadData
        
	Call InitVariables												
	
	If CheckRunningBizProcess = True Then Exit Function
    If DbQuery = False Then Exit Function

    FncQuery = True									
        
End Function
'================================================================================================================================
Function DbQuery()
	
	Dim strVal
	
	Err.Clear															'☜: Protect system from crashing

	DbQuery = False														'⊙: Processing is NG

    If LayerShowHide(1) = False Then Exit Function
    
    Call MakeKeyStream()
    
	strVal = BIZ_PGM_ID & "?txtMode="	& PopupParent.UID_M0001
	strVal = strVal & "&txtKeyStream="  & lgKeyStream
	strVal = strVal & "&lgStrPrevKey="  & lgPageNo

	Call RunMyBizASP(MyBizASP, strVal)									'☜: 비지니스 ASP 를 가동 

	DbQuery = True														'⊙: Processing is NG
End Function
'================================================================================================================================
Function DbQueryOk()														<%'☆: 조회 성공후 실행로직 %>

	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	:	frm1.vspdData.SelModeSelected = True		
	Else
		frm1.txtPoNo.focus
	End If

End Function
'========================================================================================================
' Name : MakeKeyStream
' Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream()
	Dim strPONO
	Dim ArrRowVal , ArrColVal
	Dim IDx
	
	With frm1
		lgKeyStream = UCase(Trim(.hdnRcptflg.value))  & PopupParent.gColSep
	
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then
			'lgKeyStream = lgKeyStream & UCase(Trim(.hdnPoNo.value))  & PopupParent.gColSep
			lgKeyStream = lgKeyStream & Trim(.hdnFrPoDt.value)  & PopupParent.gColSep
			lgKeyStream = lgKeyStream & Trim(.hdnToPoDt.value)  & PopupParent.gColSep
			lgKeyStream = lgKeyStream & UCase(Trim(.hdnPlantCd.value))  & PopupParent.gColSep
			
		Else
			
			lgKeyStream = lgKeyStream & Trim(.txtFrPoDt.Text)  & PopupParent.gColSep
			lgKeyStream = lgKeyStream & Trim(.txtToPoDt.Text)  & PopupParent.gColSep
			lgKeyStream = lgKeyStream & UCase(Trim(.hdnPlantCd.value))  & PopupParent.gColSep
			
			.hdnFrPoDt.value		= .txtFrPoDt.Text
			.hdnToPoDt.value		= .txtToPoDt.Text						
			'.hdnPlantCd.value		= .txtPlantCd.value
		End If
		lgKeyStream = lgKeyStream & UCase(Trim(.txtDlvyNo.value))  & PopupParent.gColSep
		lgKeyStream = lgKeyStream & UCase(Trim(.txtSlCd.value))  & PopupParent.gColSep
		
		IF frm1.rdoSelFlg0.checked = True THEN 
			lgKeyStream = lgKeyStream & ""	& PopupParent.gColSep
		ElseIf	frm1.rdoSelFlg1.checked = True THEN 
			lgKeyStream = lgKeyStream & "Y"	& PopupParent.gColSep
		ElseIf	frm1.rdoSelFlg2.checked = True THEN 
			lgKeyStream = lgKeyStream & "N"	& PopupParent.gColSep
		End If
    
				
		lgKeyStream = lgKeyStream & UCase(Trim(.hdnSupplierCd.value))  & PopupParent.gColSep
		lgKeyStream = lgKeyStream & UCase(Trim(.hdnClsflg.value))  & PopupParent.gColSep
		lgKeyStream = lgKeyStream & UCase(Trim(.hdnReleaseflg.value))  & PopupParent.gColSep
		lgKeyStream = lgKeyStream & UCase(Trim(.hdnRetflg.value))  & PopupParent.gColSep
		lgKeyStream = lgKeyStream & UCase(Trim(.hdnRefType.value))  & PopupParent.gColSep
		lgKeyStream = lgKeyStream & UCase(Trim(.hdnRcptType.value))  & PopupParent.gColSep
		lgKeyStream = lgKeyStream & UCase(Trim(.hdnIvflg.value))  & PopupParent.gColSep
		lgKeyStream = lgKeyStream & UCase(Trim(.hdnIvType.value))  & PopupParent.gColSep
		lgKeyStream = lgKeyStream & UCase(Trim(.hdnPoType.value))  & PopupParent.gColSep
		
		
		ArrRowVal = Split(.hdnDistinctNo.value , PopupParent.gRowSep)
		
		strPONO = ""		
		For iDx = 1 To  UBOUND(ArrRowVal)
			ArrColVal = Split(ArrRowVal(iDx -1) , PopupParent.gColSep)
			strPONO = strPONO & " AND (X.PO_NO <> '" & ArrColVal(0) & "' OR X.PO_SEQ_NO <> '" & ArrColVal(1) & "' OR X.SPLIT_SEQ_NO <> '" & ArrColVal(2) & "') "
		Next
		
		lgKeyStream = lgKeyStream & strPONO  & PopupParent.gColSep
		
	End With
			 
End Sub    

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS="TD5" NOWRAP>공급처</TD>
						<TD CLASS="TD6" NOWRAP>
						<INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=4 ALT="공급처" tag="14NXXU">
						<INPUT TYPE=TEXT AlT="공급처" ID="txtBpNm" tag="14X">
						<TD CLASS="TD5" NOWRAP>납품예정일</TD>
						<TD CLASS="TD6" NOWRAP>
							<table cellspacing=0 cellpadding=0>
								<tr>
									<td NOWRAP>
										<script language =javascript src='./js/u1113ra1_fpDateTime1_txtFrPoDt.js'></script>
									</td>
									<td NOWRAP>~</td>
									<td NOWRAP>
										<script language =javascript src='./js/u1113ra1_fpDateTime1_txtToPoDt.js'></script>
									</td>
								<tr>
							</table>
						</TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>발행번호</TD> 
						<TD CLASS=TD6 NOWRAP>
						<INPUT TYPE=TEXT AlT="발행번호" NAME="txtDlvyNo" SIZE=18 MAXLENGTH=18 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDlvyNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpentxtDlvyNo()"></TD>
						</TD>
						<TD CLASS="TD5" NOWRAP>입고창고</TD>
						<TD CLASS="TD6" NOWRAP>
						<INPUT TYPE=TEXT NAME="txtSlCd" SIZE=10 MAXLENGTH=4 ALT="입고창고" tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSlcd()">
						<INPUT TYPE=TEXT AlT="입고창고명" ID="txtSlNm" tag="14X">
						</TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>확정여부</TD>
						<TD CLASS="TD656" NOWRAP colspan =3><INPUT TYPE=radio AlT="전체" NAME="rdoSelFlg" ID="rdoSelFlg0" CLASS="RADIO" value = "A" tag="11" ><label for="rdoSelFlg0">&nbsp;전체&nbsp;&nbsp;</label>
															<INPUT TYPE=radio AlT="확정" NAME="rdoSelFlg" ID="rdoSelFlg1" CLASS="RADIO" value = "Y" tag="11" checked ><label for="rdoSelFlg1">&nbsp;확정&nbsp;&nbsp;</label>
															<INPUT TYPE=radio AlT="미확정" NAME="rdoSelFlg" ID="rdoSelFlg2" CLASS="RADIO" value = "N" tag="11"><label for="rdoSelFlg2">&nbsp;미확정&nbsp;&nbsp;</label></TD>
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
					<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
						<script language =javascript src='./js/u1113ra1_vspdData_vspdData.js'></script>
					</TD>
				</TR>			
			</TABLE>
		</TD>
	</TR>
    <tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
	<TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD >&nbsp;&nbsp;<IMG SRC="../../../CShared/image/query_d.gif"    Style="CURSOR: hand" ALT="Search" NAME="Search" OnClick="FncQuery()"        onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)" ></IMG></TD>
					<TD ALIGN=RIGHT> <IMG SRC="../../../CShared/image/ok_d.gif"       Style="CURSOR: hand" ALT="OK"     NAME="Ok"     OnClick="OkClick()"         onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"    ></IMG>&nbsp;
                                         <IMG SRC="../../../CShared/image/cancel_d.gif"   Style="CURSOR: hand" ALT="CANCEL" NAME="Cancel" OnClick="CancelClick()"     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
                    <TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>



<INPUT TYPE=HIDDEN NAME="hdnPoNo" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnFrPoDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnToPoDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSupplierCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnGroupCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnGroupNm" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnClsflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnReleaseflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnRcptflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnRetflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnRefType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnRcptType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnIvType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnIvFlg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnPoType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnPlantCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnTrackingNo" tag="14">

<!-- 데이터 중복등록을 피하기 위한 필드 추가 -->
<INPUT TYPE=HIDDEN NAME="hdnDistinctNo" tag="14">


</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>