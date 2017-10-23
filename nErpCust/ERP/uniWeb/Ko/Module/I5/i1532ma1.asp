<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : Vendor Managed Inventory
'*  3. Program ID           : i1532ma1
'*  4. Program Name         : Subcontract Parts Goods Receipt
'*  5. Program Desc         : VMI Goods Receipt for Subcontract Parts
'*  6. Comproxy List        : PI5G210, PI5S220, PI5S230
'*  7. Modified date(First) : 2003/02/03
'*  8. Modified date(Last)  : 2003/04/28
'*  9. Modifier (First)     : Ahn, JungJe
'* 10. Modifier (Last)      : Ahn, JungJe
'* 11. Comment              :
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit															


Const BIZ_PGM_QRY1_ID	= "i1532mb1.asp"							
Const BIZ_PGM_QRY2_ID	= "i1532mb2.asp"						
Const BIZ_PGM_SAVE_ID	= "i1532mb3.asp"

Dim C_ItemCd			
Dim C_ItemNm				
Dim C_TrackingNo
Dim C_ResvrQty			
Dim C_BasicUnit				
Dim C_IssuedQty
Dim C_RemainQty			
Dim C_OnhandQty			
Dim C_NeedQty			
Dim C_SchdRcptQty				
Dim C_SLCd				
Dim C_SLNm
Dim C_Spec

Dim C_BPCd			
Dim C_BPNm				
Dim C_PONo
Dim C_POSeqNo
Dim C_RcptType
Dim C_POUnit
Dim C_PORemainQty
Dim C_TempGRQty			
Dim C_VMISLCd
Dim C_VMISLPopup
Dim C_VMISLNm
Dim C_GoodOnhandQty
Dim C_GRQty			
Dim C_TempGoodOnhandQty			
Dim C_MakerLotNo
Dim C_MakerLotSubNo
Dim C_MakerLotNoPopup
Dim C_LotNo
Dim C_LotSubNo
Dim C_INSNo		

Dim C_HItemCd			
Dim C_HTrackingNo
Dim C_HBasicUnit
Dim C_HSLCd				
Dim C_HBPCd
Dim C_HPOUnit
Dim C_HPONo
Dim C_HPOSeqNo
Dim C_HVMISLCd
Dim C_HLotNo
Dim C_HLotSubNo
Dim C_HMakerLotNo
Dim C_HMakerLotSubNo
Dim C_HGRQty			
Dim C_HOnhandQty
Dim C_HVMISLNm
Dim C_HIndex1
Dim C_HIndex2
Dim C_HINSNo

Dim iDBSYSDate
Dim LocSvrDate

iDBSYSDate = "<%=GetSvrDate%>"		
LocSvrDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)

<!-- #Include file="../../inc/lgvariables.inc" -->
Dim lgIntPrevKey
Dim lgCurrRow
Dim lgShift, lgShiftCnt

Dim IsOpenPop 
Dim lgOldRow1, lgOldRow2
Dim lgSortKey1, lgSortKey2
Dim lgIntFlgMode1, lgIntFlgMode2
Dim lgStrPrevKey11, lgStrPrevKey12, lgStrPrevKey13, lgStrPrevKey21, lgStrPrevKey22, lgStrPrevKey23
Dim SaveCheck

Dim strItemCd
Dim strSLCd   
Dim strTrackingNo  
Dim strBPCd
Dim strVMISLCd
Dim DtlFromNo
Dim DtlToNo
Dim CurrentRow1

'==========================================  2.1.1 InitVariables()  ======================================
Sub InitVariables()
    lgIntFlgMode1 = parent.OPMD_CMODE                   
    lgIntFlgMode2 = parent.OPMD_CMODE                   
    lgIntPrevKey = 0
	lgStrPrevKey11 = ""						
	lgStrPrevKey12 = ""						
	lgStrPrevKey13 = ""						
	lgStrPrevKey21 = ""						
	lgStrPrevKey22 = ""						
	lgStrPrevKey23 = ""						
    lgLngCurRows = 0                        
    lgOldRow1 = 0
    lgOldRow2 = 0
	lgSortKey1    = 1
	lgSortKey2    = 1
	strItemCd		= ""
	strSLCd   		= ""
	strTrackingNo  	= ""
	strBPCd			= ""
	strVMISLCd		= ""
	DtlFromNo		= ""
	DtlToNo			= ""
	CurrentRow1		= ""
End Sub

'==========================================  2.2.1 SetDefaultVal()  ======================================
Sub SetDefaultVal()
    frm1.txtGRDt.text		= LocSvrDate
    frm1.cboMvmtType.Value  = "DGR"
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
'========================================================================================
Sub LoadInfTB19029()     
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I","*","NOCOOKIE","MA") %>
End Sub

'============================= 2.2.3 InitSpreadSheet() ==================================
Sub InitSpreadSheet(ByVal pvSpdNo)

    Call InitSpreadPosVariables(pvSpdNo)
	Call AppendNumberPlace("6", "3", "0")
	
	If pvSpdNo = "A" Or pvSpdNo = "*" Then

		With frm1.vspdData1 
			
			ggoSpread.Source = frm1.vspdData1
			ggoSpread.Spreadinit "V20021106", ,Parent.gAllowDragDropSpread
			
			.ReDraw = false
			
			.MaxCols = C_Spec + 1										
			.MaxRows = 0
			
			Call GetSpreadColumnPos("A")
			ggoSpread.SSSetEdit		C_ItemCd,			"품목", 15,,,,2
			ggoSpread.SSSetEdit		C_ItemNm,			"품목명", 20
			ggoSpread.SSSetEdit		C_TrackingNo,		"Tracking No.", 20
			ggoSpread.SSSetFloat	C_ResvrQty,			"필요수량", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_BasicUnit,		"단위", 8,2,,,2
			ggoSpread.SSSetFloat	C_IssuedQty,		"기출고수량",15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_RemainQty,		"출고잔량", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_OnhandQty,		"재고수량",15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_NeedQty,			"입고필요량",15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_SchdRcptQty,		"입고예정량",15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"	
			ggoSpread.SSSetEdit		C_SLCd	,			"사급품창고", 10,,,,2
			ggoSpread.SSSetEdit		C_SLNm,				"사급품창고명", 12
			ggoSpread.SSSetEdit		C_Spec,				"규격", 20

			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
			    
			ggoSpread.SSSetSplit2(2)
			ggoSpread.SpreadLock -1, -1

			.ReDraw = true
		End With
	End If
	
	If pvSpdNo = "B" Or pvSpdNo = "*" Then
		'------------------------------------------
		' Grid 2 - Component Spread Setting
		'------------------------------------------
		With frm1.vspdData2
			
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.Spreadinit "V20021106", ,Parent.gAllowDragDropSpread

			.ReDraw = false
			
			.MaxCols = C_INSNo + 1									
			.MaxRows = 0
			
			Call GetSpreadColumnPos("B")
			
			ggoSpread.SSSetEdit 	C_BPCd,				"공급처", 10,,,,2
			ggoSpread.SSSetEdit		C_BPNm,				"공급처명", 12
			ggoSpread.SSSetEdit		C_PONo,				"발주번호", 15
			ggoSpread.SSSetFloat	C_POSeqNo,			"발주순번", 4, "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, , ,"Z"
			ggoSpread.SSSetEdit		C_RcptType,			"입고형태", 8,2,,,2
			ggoSpread.SSSetEdit		C_POUnit,			"발주단위", 8,2,,,2
			ggoSpread.SSSetFloat	C_PORemainQty,		"발주잔량",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_TempGRQty,		"입고예정량",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_VMISLCd,			"VMI창고", 10,,,7,2
			ggoSpread.SSSetButton	C_VMISLPopup
			ggoSpread.SSSetEdit		C_VMISLNm,			"VMI창고명", 12
			ggoSpread.SSSetFloat	C_GoodOnhandQty,	"재고수량",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_GRQty	,			"입고수량",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_TempGoodOnhandQty,"예상재고수량",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_MakerLotNo,		"MAKER LOT NO.", 12
			ggoSpread.SSSetFloat	C_MakerLotSubNo,	"Maker Lot 순번", 4, "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, , ,"Z"
			ggoSpread.SSSetButton	C_MakerLotNoPopup
			ggoSpread.SSSetEdit		C_LotNo,			"Lot No.", 12
			ggoSpread.SSSetFloat	C_LotSubNo,			"Lot 순번", 4, "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, , ,"Z"
			ggoSpread.SSSetFloat	C_INSNo,			"INSNo", 4, "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, , ,"Z"

			Call ggoSpread.MakePairsColumn(C_VMISLCd, C_VMISLPopup)
			Call ggoSpread.MakePairsColumn(C_MakerLotNo, C_MakerLotNoPopup)
			Call ggoSpread.SSSetColHidden(C_POUnit, C_POUnit, True)
			Call ggoSpread.SSSetColHidden(C_INSNo, C_INSNo, True)
			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)

			ggoSpread.SSSetSplit2(1)
			ggoSpread.SpreadLock -1, -1
			
			.ReDraw = true
    
		End With
	End If

	If pvSpdNo = "C" Or pvSpdNo = "*" Then
		'------------------------------------------
		' Grid 2 - Component Spread Setting
		'------------------------------------------
		With frm1.vspdData3
			
			ggoSpread.Source = frm1.vspdData3
			ggoSpread.Spreadinit "V20021106", ,Parent.gAllowDragDropSpread
			
			.ReDraw = false
			
			.MaxCols = C_HINSNo + 1									
			.MaxRows = 0

			Call GetSpreadColumnPos("C")
			ggoSpread.SSSetEdit		C_HItemCd,			"품목", 15,,,,2
			ggoSpread.SSSetEdit		C_HTrackingNo,		"Tracking No.", 20
			ggoSpread.SSSetEdit		C_HBasicUnit,		"단위", 8,2,,,2
			ggoSpread.SSSetEdit		C_HSLCd	,			"현장창고", 10,,,,2
			ggoSpread.SSSetEdit 	C_HBPCd,			"공급처", 10,,,,2
			ggoSpread.SSSetEdit		C_HPOUnit,			"발주단위", 8,2,,,2
			ggoSpread.SSSetEdit		C_HPONo,			"발주번호", 15
			ggoSpread.SSSetFloat	C_HPOSeqNo,			"발주순번", 4, "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, , ,"Z"
			ggoSpread.SSSetEdit		C_HVMISLCd,			"VMI창고", 10,,,7,2
			ggoSpread.SSSetEdit		C_HLotNo,			"Lot No.", 12
			ggoSpread.SSSetFloat	C_HLotSubNo,		"Lot 순번", 4, "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, , ,"Z"
			ggoSpread.SSSetEdit		C_HMakerLotNo,		"MAKER LOT NO.", 12
			ggoSpread.SSSetFloat	C_HMakerLotSubNo,	"Maker Lot 순번", 4, "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, , ,"Z"
			ggoSpread.SSSetFloat	C_HGRQty,			"입고수량",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_HOnhandQty,		"재고수량",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_HVMISLNm,			"VMI창고명", 12
			ggoSpread.SSSetFloat	C_HIndex1,		"No1", 4, "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, , ,"Z"
			ggoSpread.SSSetFloat	C_HIndex2,		"No2", 4, "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, , ,"Z"
			ggoSpread.SSSetFloat	C_HINSNo,		"INSNo", 4, "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, , ,"Z"

			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
			    
			.ReDraw = true

		End With
	End If
End Sub

'============================ 2.2.4 SetSpreadLock() =====================================
Sub SetSpreadLock(ByVal Row)

    With frm1.vspdData2
			'--------------------------------
			'Grid 2
			'--------------------------------
			ggoSpread.Source = frm1.vspdData2

			ggoSpread.SpreadUnLock C_VMISLCd,	Row, C_VMISLPopup,	Row		
			ggoSpread.SpreadUnLock C_GRQty,		Row, C_GRQty,		Row		
			ggoSpread.SpreadUnLock C_MakerLotNo,Row, C_MakerLotNoPopup,	Row
			ggoSpread.SpreadUnLock C_LotNo,		Row, C_LotNo,		Row
			ggoSpread.SpreadUnLock C_LotSubNo,	Row, C_LotSubNo,	Row
			
			ggoSpread.SSSetRequired	C_GRQty,	Row, Row
			
			.Row = Row
			.Col = C_VMISLCd
			.lock = True
			.Col = C_MakerLotNo
			.lock = True
			.Col = C_MakerLotSubNo
			.lock = True
	End With
End Sub

'==========================================  2.2.7 InitSpreadPosVariables() =================
Sub InitSpreadPosVariables(ByVal pvSpdNo)
	
	If pvSpdNo = "A" Or pvSpdNo = "*" Then
		 C_ItemCd		= 1
		 C_ItemNm		= 2
		 C_TrackingNo	= 3
		 C_ResvrQty		= 4
		 C_BasicUnit	= 5
		 C_IssuedQty	= 6
		 C_RemainQty	= 7
		 C_OnhandQty	= 8
		 C_NeedQty		= 9
		 C_SchdRcptQty	= 10
		 C_SLCd			= 11
		 C_SLNm			= 12
		 C_Spec			= 13
	End If	 

	If pvSpdNo = "B" Or pvSpdNo = "*" Then
		 C_BPCd			= 1
		 C_BPNm			= 2
		 C_PONo			= 3
		 C_POSeqNo		= 4
		 C_RcptType		= 5
		 C_POUnit		= 6		
		 C_PORemainQty	= 7 
		 C_TempGRQty	= 8
		 C_VMISLCd			= 9
		 C_VMISLPopup		= 10
		 C_VMISLNm			= 11
		 C_GoodOnhandQty	= 12
		 C_GRQty			= 13
		 C_TempGoodOnhandQty= 14
		 C_MakerLotNo		= 15
		 C_MakerLotSubNo	= 16
		 C_MakerLotNoPopup	= 17
		 C_LotNo			= 18
		 C_LotSubNo			= 19
		 C_INSNo			= 20
	End If

	If pvSpdNo = "C" Or pvSpdNo = "*" Then
		C_HItemCd		= 1
		C_HTrackingNo	= 2
		C_HBasicUnit	= 3
		C_HSLCd			= 4 
		C_HBPCd			= 5
		C_HPOUnit		= 6
		C_HPONo			= 7
		C_HPOSeqNo		= 8
		C_HVMISLCd		= 9	
		C_HLotNo		= 10
		C_HLotSubNo		= 11
		C_HMakerLotNo	= 12
		C_HMakerLotSubNo= 13
		C_HGRQty		= 14	
		C_HOnhandQty	= 15
		C_HVMISLNm		= 16
		C_HIndex1		= 17
		C_HIndex2		= 18
		C_HINSNo		= 19
	End If	
End Sub

'==========================================  2.2.8 GetSpreadColumnPos()  ==========
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
		Case "A"
 			ggoSpread.Source = frm1.vspdData1
			
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_ItemCd		= iCurColumnPos(1)
			C_ItemNm		= iCurColumnPos(2)
			C_TrackingNo	= iCurColumnPos(3)
			C_ResvrQty		= iCurColumnPos(4)
			C_BasicUnit		= iCurColumnPos(5)
			C_IssuedQty		= iCurColumnPos(6)
			C_RemainQty		= iCurColumnPos(7)
			C_OnhandQty		= iCurColumnPos(8)
			C_NeedQty		= iCurColumnPos(9)
			C_SchdRcptQty	= iCurColumnPos(10)
			C_SLCd			= iCurColumnPos(11)
			C_SLNm			= iCurColumnPos(12)
			C_Spec			= iCurColumnPos(13)
			
		Case "B"
			ggoSpread.Source = frm1.vspdData2
			
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_BPCd			= iCurColumnPos(1)
			C_BPNm			= iCurColumnPos(2)
			C_PONo			= iCurColumnPos(3)
			C_POSeqNo		= iCurColumnPos(4)
			C_RcptType		= iCurColumnPos(5)
			C_POUnit		= iCurColumnPos(6)
			C_PORemainQty	= iCurColumnPos(7)
			C_TempGRQty		= iCurColumnPos(8)
			C_VMISLCd			= iCurColumnPos(9)
			C_VMISLPopup		= iCurColumnPos(10)
			C_VMISLNm			= iCurColumnPos(11)
			C_GoodOnhandQty		= iCurColumnPos(12)
			C_GRQty				= iCurColumnPos(13)
			C_TempGoodOnhandQty = iCurColumnPos(14)
			C_MakerLotNo		= iCurColumnPos(15)
			C_MakerLotSubNo		= iCurColumnPos(16)
			C_MakerLotNoPopup	= iCurColumnPos(17)
			C_LotNo				= iCurColumnPos(18)
			C_LotSubNo			= iCurColumnPos(19)
			C_INSNo				= iCurColumnPos(20)
  
   		Case "C"
 			ggoSpread.Source = frm1.vspdData3
			
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
				C_HItemCd		= iCurColumnPos(1)
				C_HTrackingNo	= iCurColumnPos(2)
				C_HBasicUnit	= iCurColumnPos(3)
				C_HSLCd			= iCurColumnPos(4) 
				C_HBPCd			= iCurColumnPos(5)
				C_HPOUnit		= iCurColumnPos(6)
				C_HPONo			= iCurColumnPos(7)
				C_HPOSeqNo		= iCurColumnPos(8)
				C_HVMISLCd		= iCurColumnPos(9)	
				C_HLotNo		= iCurColumnPos(10)
				C_HLotSubNo		= iCurColumnPos(11)
				C_HMakerLotNo	= iCurColumnPos(12)
				C_HMakerLotSubNo= iCurColumnPos(13)
				C_HGRQty		= iCurColumnPos(14)	
				C_HOnhandQty	= iCurColumnPos(15)	  
				C_HVMISLNm		= iCurColumnPos(16)
				C_HIndex1		= iCurColumnPos(17)
				C_HIndex2		= iCurColumnPos(18)
				C_HINSNo		= iCurColumnPos(19)
    End Select    
End Sub    

'==========================================  2.2.2 InitComboBox()  ========================================
Sub InitComboBox()
	Call CommonQueryRs(" IO_TYPE_CD,IO_TYPE_NM "," M_MVMT_TYPE ", " ((RCPT_FLG = " & FilterVar("Y", "''", "S") & "  AND RET_FLG = " & FilterVar("N", "''", "S") & " ) OR (RET_FLG = " & FilterVar("N", "''", "S") & "  AND SUBCONTRA_FLG = " & FilterVar("N", "''", "S") & " )) AND USAGE_FLG = " & FilterVar("Y", "''", "S") & " ", _
						lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboMvmtType ,lgF0  ,lgF1  ,Chr(11))
End Sub

'------------------------------------------  OpenPlant()  ------------------------------------------------
Function OpenPlant()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"			
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
		Exit Function
	Else
		Call SetPlant(arrRet)
	End If
End Function

'------------------------------------------  OpenSBP()  ------------------------------------------------
Function OpenSBP()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "사급업체팝업"				
	arrParam(1) = "B_BIZ_PARTNER "					
	arrParam(2) = Trim(frm1.txtSBPCd.Value)		
	arrParam(3) = ""
	arrParam(4) = "BP_CD in (select distinct BP_CD from B_STORAGE_LOCATION where SL_TYPE = " & FilterVar("E", "''", "S") & " )"							' Where Condition
	arrParam(5) = "사급업체"					
	
    arrField(0) = "BP_CD"					
    arrField(1) = "BP_NM"					
    
    arrHeader(0) = "사급업체"			
    arrHeader(1) = "사급업체명"			
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtSBPCd.focus
		Exit Function
	Else
		Call SetSBP(arrRet)
	End If
End Function

'------------------------------------------  OpenItemCd()  ---------------------------------------------
Function OpenItemCd()

	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("169901","X","X","X")    
		frm1.txtPlantCd.focus 
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("i1522pa1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "i1522pa1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	
	arrParam(1) = Trim(frm1.txtItemCd.value)	
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1)), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		Call SetItemCd(arrRet)
	End If	
End Function

'------------------------------------------  OpenPurchaseGroup()  ------------------------------------------------
Function OpenPurchaseGroup()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "구매그룹"				
	arrParam(1) = "B_PUR_GRP"					
	arrParam(2) = Trim(frm1.txtGroupCd.Value)	
	arrParam(3) = ""
	arrParam(4) = "USAGE_FLG = " & FilterVar("Y", "''", "S") & "  "			
	arrParam(5) = "구매그룹"				
	
    arrField(0) = "PUR_GRP"					
    arrField(1) = "PUR_GRP_NM"				
    
    arrHeader(0) = "구매그룹"			
    arrHeader(1) = "구매그룹명"			
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtGroupCd.focus
		Exit Function
	Else
		Call SetPurchaseGroup(arrRet)
	End If
End Function

'------------------------------------------  OpenVMISL()  ------------------------------------------------
Function OpenVMISL(Byval strCode)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "VMI창고팝업"				
	arrParam(1) = "I_VMI_STORAGE_LOCATION"						
	arrParam(2) = strCode	
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txthPlantCd.Value, "''", "S")	
	arrParam(5) = "VMI창고"				
	
    arrField(0) = "SL_CD"					
    arrField(1) = "SL_NM"					
    
    arrHeader(0) = "VMI창고"				
    arrHeader(1) = "VMI창고명"				
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Call SetActiveCell(frm1.vspdData2,C_VMISLCd,1,"M","X","X")
		Exit Function
	Else
		Call SetVMISL(arrRet)
	End If
End Function

'------------------------------------------  OpenItem()  -------------------------------------------------
Function OpenMakerLotNo()
	Dim iCalledAspName
	Dim IntRetCD

	Dim arrRet
	Dim arrParam1, arrParam2, arrParam3, arrParam4, arrParam5, arrParam6, arrParam7, arrParam8
	
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("169901","X", "X", "X")   
		frm1.txtPlantCd.focus
		Exit Function
	End If

	If UCase(Trim(frm1.txtPlantCd.value)) <> UCase(Trim(frm1.txthPlantCd.value)) Then
		Call DisplayMsgBox("900002","X","X","X")
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function	

	IsOpenPop = True
	
	With frm1.vspdData2
	
		arrParam1 = Trim(frm1.txtPlantCd.Value)   
		arrParam2 = Trim(frm1.txtPlantNm.Value)   
		.Row = .ActiveRow
		.Col = C_VMISLCd
		arrParam3 = Trim(.Text)      
		.Col = C_VMISLNm
		arrParam4 = Trim(.Text)      
		.Col = C_BPCd	
		arrParam5 = Trim(.Text)      
		.Col = C_BPNm
		arrParam6 = Trim(.Text)      
	End With
	With frm1.vspdData1
		.Row = .ActiveRow
		.Col = C_ItemCd
		arrParam7 = Trim(.Text)              
		.Col = C_ItemNm
		arrParam8 = Trim(.Text)              
	End With
	
	iCalledAspName = AskPRAspName("I1523PA1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"I1523PA1","x")
		IsOpenPop = False
		Call SetFocusToDocument("M") 
		.focus
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam1, arrParam2, arrParam3, arrParam4, arrParam5, arrParam6, arrParam7, arrParam8), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Call SetActiveCell(frm1.vspdData2,C_MakerLotNo,1,"M","X","X")
		Exit Function
	Else
		Call SetMakerLotNo(arrRet)
	End If
End Function

'------------------------------------------  OpenRcptRef()  ----------------------------------------------
Function OpenGRRef()

	Dim arrRet
	Dim arrParam(2)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	If frm1.txtMvmtNo.value = "" Then
		Call DisplayMsgBox("900027","X","X","X")
		frm1.txtPlantCd.focus    
		Exit Function
	End If

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("169901","X", "X", "X")   
		frm1.txtPlantCd.focus 
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("I1524RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "I1524RA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = Trim(frm1.txtPlantNm.value)
	arrParam(2) = Trim(frm1.txtMvmtNo.value)			

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1), arrParam(2)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function

'------------------------------------------  SetPlant()  -------------------------------------------------
Function SetPlant(byRef arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
	frm1.txtPlantCd.focus
End Function

'------------------------------------------  SetPlant()  -------------------------------------------------
Function SetSBP(byRef arrRet)
	frm1.txtSBPCd.Value    = arrRet(0)		
	frm1.txtSBPNm.Value    = arrRet(1)
	frm1.txtSBPCd.focus
End Function

'------------------------------------------  SetItemCd()  ----------------------------------------------
Function SetItemCd(byRef arrRet)
	frm1.txtItemCd.value = arrRet(0)
	frm1.txtItemNm.value = arrRet(1)
	frm1.txtItemCd.focus
End Function

'------------------------------------------  SetPurchaseGroup()  ----------------------------------------------
Function SetPurchaseGroup(byRef arrRet)
	frm1.txtGroupCd.value = arrRet(0)
	frm1.txtGroupNm.value = arrRet(1)
	frm1.txtGroupCd.focus
End Function

'------------------------------------------  SetPlant()  -------------------------------------------------
Function SetVMISL(byRef arrRet)
	Dim SEQ1
	Dim Row2
	Dim INSRow

	frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
	frm1.vspdData1.Col = frm1.vspdData1.MaxCols
	SEQ1 = UNICDbl(frm1.vspdData1.Text)

	With frm1.vspdData2
		.Row = .ActiveRow
		.Col = .MaxCols
		Row2 = UNICDbl(.Text)
		.Col = C_INSNo
		INSRow = UNICDbl(.Text)
		.Col = C_VMISLCd
		If .Text <> arrRet(0) Then
			Call .SetText(C_VMISLCd,		.ActiveRow, arrRet(0))
			Call .SetText(C_GoodOnhandQty,	.ActiveRow, "")
			Call .SetText(C_GRQty,			.ActiveRow, "")
			Call .SetText(C_TempGoodOnhandQty, .ActiveRow, "")
			Call .SetText(C_MakerLotNo,		.ActiveRow, "")
			Call .SetText(C_MakerLotSubNo,	.ActiveRow, "")
			Call .SetText(C_VMISLNm,		.ActiveRow, arrRet(1))
		Else
			Call DisplayMsgBox("169906","X","X","X")
			Call SetFocusToDocument("M") 
			.focus
			Exit Function
		End If
		
		Call CopyToHSheet(SEQ1, Row2, INSRow, 0)
		Call SetTempGOHQtyToGrid2(SEQ1, UNICDbl(.ActiveRow))
		Call SetTempPORemainQtyToGrid2(SEQ1, INSRow, 0)
		Call TOTBPInfo(.ActiveRow, StrBPCd)
		Call TOTVMISLInfo(arrRet(0))
	End With
End Function

'------------------------------------------  SetTrackingNo()  --------------------------------------------------
Function SetMakerLotNo(byRef arrRet)
	Dim PopupItemCd
	Dim PopupTrackingNo
	Dim SEQ1
	Dim Row2
	Dim INSRow

	With frm1.vspdData1
		.Row = .ActiveRow
		.Col = .MaxCols
		SEQ1 = UNICDbl(.Text)
		.Col = C_ItemCd
		PopupItemCd = .Text
		.Col = C_TrackingNo
		PopupTrackingNo = .Text
		If Trim(arrRet(0)) <> Trim(PopupItemCd) or Trim(arrRet(4)) <> Trim(PopupTrackingNo) Then 
			Call DisplayMsgBox("162052","X","X","X")
			Call SetFocusToDocument("M") 
			frm1.vspdData2.focus
			Exit Function
		End If
	End With
	
	With frm1.vspdData2
		.Row = .ActiveRow
		.Col = .MaxCols
		Row2 = UNICDbl(.Text)
		.Col = C_INSNo
		INSRow = UNICDbl(.Text)
		Call .SetText(C_GoodOnhandQty,		.ActiveRow, UNIFormatNumber(UNICDbl(arrRet(2)),ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))
		Call .SetText(C_TempGoodOnhandQty,	.ActiveRow, UNIFormatNumber(UNICDbl(arrRet(2)),ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))
		Call .SetText(C_MakerLotNo,			.ActiveRow, arrRet(5))
		Call .SetText(C_MakerLotSubNo,		.ActiveRow, arrRet(6))
		Call .SetText(C_GRQty,				.ActiveRow, 0)

		.Col = 0
		If .Text = ggoSpread.UpdateFlag Then .Text = ""

		Call CopyToHSheet(SEQ1, Row2, INSRow, 0)
		Call SetTempGOHQtyToGrid2(SEQ1, UNICDbl(.ActiveRow))
		Call SetTempPORemainQtyToGrid2(SEQ1, INSRow, 0)
		Call TOTBPInfo(.ActiveRow, StrBPCd)
	End With
End Function

'------------------------------------------  txtGRDt_KeyDown ------------------------------------------
Sub txtGRDt_KeyDown(keycode, shift)
	If Keycode = 13 Then
		Call MainQuery()
	End If
End Sub	

'=======================================================================================================
'   Event Name : txtReportDT_DblClick(Button)
'=======================================================================================================
Sub txtGRDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtGRDt.Action = 7
    End If
End Sub

'==========================================  3.1.1 Form_Load()  ===========================================
Sub Form_Load()

    Call LoadInfTB19029                                               
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)        
    Call ggoOper.LockField(Document, "N")                             
    
    Call InitSpreadSheet("*")                                         
    Call InitVariables
   	Call InitComboBox
    Call SetDefaultVal
    Call SetToolBar("11000000000011")								
    
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtReqStartDt.focus 
	Else
		frm1.txtPlantCd.focus 
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData2_Keydown
'==========================================================================================
Sub vspdData2_Keydown(keycode, shift)
	With frm1.vspdData2
		If .MaxRows = 0 Then Exit Sub
		If keycode = 9 or (keycode => 37 and keycode =< 40) Then Exit Sub
		
		.Row = .ActiveRow
		.Col = C_VMISLCd
		If UNICDbl(.ActiveCol) = C_VMISLCd Then
			If .Text <> "" Then
				Call DisplayMsgBox("162053","X","X","X")
				Call SetFocusToDocument("M") 
				.focus
			End If 
		ElseIf UNICDbl(.ActiveCol) = C_MakerLotNo or UNICDbl(.ActiveCol) = C_MakerLotSubNo Then
			.Col = C_MakerLotNoPopup
			If .Lock = False Then
				Call DisplayMsgBox("162053","X","X","X")
				Call SetFocusToDocument("M") 
				.focus
			End If
		End If	
	End With
End Sub	

'==========================================================================================
'   Event Name : vspdData1_Click
'==========================================================================================
Sub vspdData1_Click(ByVal Col , ByVal Row )
	
	If lgIntFlgMode1 = parent.OPMD_CMODE Then
		Call SetPopupMenuItemInf("0000111111")     
	Else
		Call SetPopupMenuItemInf("0001111111")     
	End If
	
	gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData1
	
	If frm1.vspdData1.MaxRows = 0 Then Exit Sub
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData1 
 		If lgSortKey1 = 1 Then
 			ggoSpread.SSSort Col				
 			lgSortKey1 = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey1	
 			lgSortKey1 = 1
 		End If
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.ClearSpreadData
		lgIntFlgMode2 = parent.OPMD_CMODE
 		Call DbDtlQuery1(1)
 	End If
End Sub

'==========================================================================================
'   Event Name : vspdData2_Click
'==========================================================================================
Sub vspdData2_Click(ByVal Col , ByVal Row )
	
	If lgIntFlgMode2 = parent.OPMD_CMODE Then
		Call SetPopupMenuItemInf("0000111111")       
	Else
		Call SetPopupMenuItemInf("1001111111")       
	End If
	
	gMouseClickStatus = "SP2C"
	Set gActiveSpdSheet = frm1.vspdData2
	
 	If frm1.vspdData2.MaxRows = 0 Then Exit Sub
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData2 
 		If lgSortKey2 = 1 Then
 			ggoSpread.SSSort Col				
 			lgSortKey2 = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey2	
 			lgSortKey2 = 1
 		End If
	End If 	
End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'==========================================================================================
Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	With frm1.vspdData2 
		ggoSpread.Source = frm1.vspdData2
		
		If Row > 0 And Col = C_VMISLPopUp Then
			.Row = Row
			.Col = C_VMISLCd
			Call OpenVMISL(.Text)
		Elseif Row > 0 And Col = C_MakerLotNoPopup Then
			Call OpenMakerLotNo()
		End If
	End With
End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'==========================================================================================
Sub vspdData1_MouseDown(Button,Shift,x,y)
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

Sub vspdData2_MouseDown(Button,Shift,x,y)
	If Button <> "1" And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub

'========================================================================================
' Function Name : vspdData_ColWidthChange
'========================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 

Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 

'=======================================================================================================
'   Event Name : vspdData2_Change
'=======================================================================================================
Sub vspdData2_Change(ByVal Col, ByVal Row)
	Dim SEQ_NO1
	Dim OldSEQ_NO
	Dim NewSEQ_NO
	Dim Onhand_Qty	
	Dim GR_Qty
	Dim Row2
	Dim INSRow
	Dim Temp_Onhand_Qty	
	Dim Temp_PORemain_Qty
	Dim chgVMISLCd
	
	frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
	frm1.vspdData1.Col = frm1.vspdData1.MaxCols
	SEQ_NO1 = UNICDbl(frm1.vspdData1.Text)

	With frm1.vspdData2
		
		.Row = Row
	    .Col = C_PORemainQty
		Temp_PORemain_Qty = UNICDbl(.Text)
	    .Col = C_GRQty
		GR_Qty = UNICDbl(.text)
  		.Col = C_VMISLCd
		chgVMISLCd = .text
	    .Col = C_GoodOnhandQty
		Onhand_Qty = UNICDbl(.text)
	    .Col = C_INSNo
		INSRow = UNICDbl(.text)
		.Col = .MaxCols
		Row2 = UNICDbl(.text)	
		
		.Col = C_MakerLotNo
		If .Text = "" Then
			Call DisplayMsgBox("169925","X","X","X")
			Call .SetText(C_GRQty, Row, 0)
			Call SetFocusToDocument("M") 
			.focus
			Exit Sub
		End If
		
		If GR_Qty <> 0 Then
			If CheckPOItem(INSRow) = False Then
				Call .SetText(C_GRQty, Row, 0)
				GR_Qty = 0
			End If
		End If

		If Onhand_Qty < GR_Qty Then
			Call DisplayMsgBox("162054","X","X","X") 
			.Col = 0
			If .Text = ggoSpread.UpdateFlag Then .Text = ""
			Call .SetText(C_GRQty, Row, 0)
			Call SetFocusToDocument("M") 
			.focus
			Exit Sub
		End If
		
		If GR_Qty = 0 Then
			If INSRow = Row2 Then
				.Col = 0
				If .Text = ggoSpread.UpdateFlag Then .text = ""
			End If
		Else
			If GR_Qty > Temp_PORemain_Qty Then
				Call DisplayMsgBox("162055","X","X","X")  
				GR_Qty = 0
				Call .SetText(C_GRQty, Row, 0)
				Call .SetText(C_TempGoodOnhandQty, Row, UNIFormatNumber(Onhand_Qty,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))
				.Col = 0
				If .Text = ggoSpread.UpdateFlag Then .Text = ""
			Else
				If INSRow = Row2 Then
					Call .SetText(0, Row, ggoSpread.UpdateFlag)
				End If
			End If
		End If

		Call CopyToHSheet(SEQ_NO1, Row2, INSRow, GR_Qty)
		If SetTempGOHQtyToGrid2(SEQ_NO1, UNICDbl(.ActiveRow)) = True Then Call SetTempPORemainQtyToGrid2(SEQ_NO1, INSRow, GR_Qty)
	
		Call TOTVMISLInfo(chgVMISLCd)
		Call TOTBPInfo(Row, strBPCd)
	End With
End Sub

'==========================================================================================
'   Event Name : vspdData_GotFocus
'==========================================================================================
Sub vspdData1_GotFocus()
    ggoSpread.Source = frm1.vspdData1
End Sub

Sub vspdData2_GotFocus()
    ggoSpread.Source = frm1.vspdData2
End Sub

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'==========================================================================================
Sub vspdData1_ScriptLeaveCell(ByVal Col, ByVal Row, Byval NewCol, Byval NewRow, Byval Cancel)
	
	If NewRow <= 0 Or Row = NewRow Then Exit Sub
	
	lgOldRow1 = NewRow
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData
	lgIntFlgMode2 = parent.OPMD_CMODE

	frm1.txtTOTBP.value				= ""
	frm1.txtTOTPORemainQty.text		= 0
	frm1.txtTOTTempGRQty.text		= 0
	frm1.txtTOTTempPORemainQty.text = 0
	frm1.txtTOTSL.value		= ""
	frm1.txtTOTGRQty.text	= 0
	frm1.txtTOTGOHQty.text	= 0
	frm1.txtTOTTempGOHQty.text = 0

	If DbDtlQuery1(NewRow) = False Then	Exit Sub
End Sub

Sub vspdData2_ScriptLeaveCell(ByVal Col, ByVal Row, Byval NewCol, Byval NewRow, Byval Cancel)
	
	If NewRow <= 0 or Row = NewRow Then Exit Sub

	With frm1.vspdData2
		.Row = NewRow
		.Col = C_BPCd

		If .Text <> strBPCd Then 
			strBPCd = .Text
			Call TOTBPInfo(NewRow, strBPCd)
		End If
				
		.Row = Row
		.Col = C_VMISLCd
		strVMISLCd = .Text	
		.Row = NewRow
		.Col = C_VMISLCd
		If .text = strVMISLCd Then
			Exit Sub
		ElseIf .Text = "" Then
			frm1.txtTOTSL.value		= ""
			frm1.txtTOTGRQty.text	= 0
			frm1.txtTOTGOHQty.text	= 0
			frm1.txtTOTTempGOHQty.text = 0
			Exit Sub
		End If
		strVMISLCd = .Text
		Call TOTVMISLInfo(strVMISLCd)

		lgOldRow2 = NewRow

	End With
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'==========================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then Exit Sub
    If CheckRunningBizProcess = True Then Exit Sub		
   
    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) and lgStrPrevKey11 <> "" Then					
		If DbQuery = False Then	
			Call RestoreToolBar()
			Exit Sub
		End If
    End if
End Sub

Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

    If OldLeft <> NewLeft Then Exit Sub
    If CheckRunningBizProcess = True Then Exit Sub		

    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) and lgStrPrevKey21 <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
		Call LayerShowHide(1)
		If DbDtlQuery1(frm1.vspdData1.ActiveRow) = False Then	
			Call RestoreToolBar()
			Exit Sub
		End If
    End if
End Sub


'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
'========================================================================================
Sub vspdData1_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub 

Sub vspdData2_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("B")
End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub 
 
'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
'========================================================================================
Sub PopRestoreSpreadColumnInf()
	Dim ActiveRow2
	Dim i
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.RestoreSpreadInf()
	Call InitSpreadSheet(gActiveSpdSheet.Id)
    Call ggoSpread.ReOrderingSpreadData
	
	ActiveRow2 = UNICDbl(frm1.vspdData2.MaxRows)
	
	Select Case gActiveSpdSheet.id 
		Case "A"
			frm1.vspdData1.ReDraw = False
			For i = 1 To frm1.vspdData1.MaxRows
				frm1.vspdData1.Row = i
				frm1.vspdData1.Action = 0
				Call SetSchRcptQtyToGrid1
			Next
			frm1.vspdData1.ReDraw = True

		Case "B"
			With frm1.vspdData1
				.Row = .ActiveRow
				.Col = .MaxCols
				CurrentRow1 = UNICDbl(.Text)

				Call SetVMIItemDetail(1, ActiveRow2)
	
				.Row = .ActiveRow
				.Col = 0
				If 	.Text <> "" Then
					If CheckItemMaxRow(CurrentRow1, 1, ActiveRow2) = True Then Call CopyHiddenToGrid2(CurrentRow1, 1)
					Exit Sub
				End If
			End With
	
			With frm1.vspdData2
				.Row = 1
				.Col = C_BPCd
				strBPCd = .Text
				Call TOTBPInfo(1, strBPCd)
				
				.Row = 1
				.Col = C_VMISLCd
				If .Text <> "" Then Call TOTVMISLInfo(.Text)
			End With
		End Select
	End Sub 

'========================================================================================
' Function Name : FncQuery
'========================================================================================
Function FncQuery() 
    
    Dim IntRetCD 
    
    FncQuery = False								
    
    Err.Clear										
 	
 	If Not chkField(Document, "1") Then Exit Function
 	
 	If Trim(frm1.txtPlantCd.Value) = "" Then
		Call DisplayMsgBox("189220","X","X","X")
		frm1.txtPlantNm.Value = ""
		frm1.txtPlantCd.focus
		Exit function
	End If
    
    ggoSpread.Source = frm1.vspdData3							
    If ggoSpread.SSCheckChange = True Then						
        IntRetCD = displaymsgbox("900013", parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then Exit Function
    End If
   
    If Plant_Item_POInfo_Check(1) = False Then Exit Function
    
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")					
    frm1.txtGRDt.text = LocSvrDate

    If SaveCheck = False Then
		frm1.txtMvmtNo.value = ""
	End If
    
    Call InitVariables
  
	If DbQuery = False Then	Exit Function
	Set gActiveElement = document.activeElement
    FncQuery = True										
   
End Function

'========================================================================================
' Function Name : FncSave
'========================================================================================
Function FncSave() 

    Dim IntRetCD 
    Dim	LngRows
    
    FncSave = False											
    
    Err.Clear                                               

    ggoSpread.Source = frm1.vspdData3						
    If ggoSpread.SSCheckChange = False Then					
        Call displaymsgbox("900001", "x", "x", "x")		
        Exit Function
    End If
 
   	If UCase(Trim(frm1.txtPlantCd.value)) <> UCase(Trim(frm1.txthPlantCd.value)) Then
		IntRetCD = DisplayMsgBox("162060", parent.VB_YES_NO, frm1.txtPlantCd.Alt,"X")
		If IntRetCD = vbNo Then
			frm1.txtPlantCd.focus 
			Exit Function
		End If
		frm1.txtPlantCd.Value = frm1.txthPlantCd.value
	End If
   	
   	If UCase(Trim(frm1.cbohMvmtType.value)) <> UCase(Trim(frm1.cboMvmtType.value)) Then
		IntRetCD = DisplayMsgBox("162060",parent.VB_YES_NO, frm1.cboMvmtType.Alt,"X")
		If IntRetCD = vbNo Then
			frm1.cboMvmtType.focus 
			Exit Function
		End If
		frm1.cboMvmtType.Value = frm1.cbohMvmtType.value
	End If

    ggoSpread.Source = frm1.vspdData2						
    If Not ggoSpread.SSDefaultCheck Then Exit Function		
    
    If Not chkfield(Document, "2") Then	Exit Function		
       
    If Plant_Item_POInfo_Check(2) = False Then Exit Function
   
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then 
		Call SetToolBar("11001001001111")
		Exit Function
	End If	
	Set gActiveElement = document.activeElement    
    FncSave = True											
    
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
	Dim aRow
	
    With frm1.vspdData2
    
		If .Maxrows < 1 then exit function
		aRow = .ActiveRow
 		.Row = aRow
		.Col = C_VMISLCd
 		If .Text = "" then exit function

		.ReDraw = False

		ggoSpread.Source = frm1.vspdData2	
		ggoSpread.CopyRow
		ggoSpread.SpreadLock -1, aRow ,, aRow

		Call SetSpreadLock(aRow)

		Call .SetText(C_VMISLCd,		aRow, "")
		Call .SetText(C_VMISLNm,		aRow, "")
		Call .SetText(C_GoodOnhandQty,	aRow, 0)
		Call .SetText(C_GRQty,			aRow, 0)
		Call .SetText(C_TempGoodOnhandQty, aRow, 0)
		Call .SetText(C_MakerLotNo,		aRow, "")
		Call .SetText(C_MakerLotSubNo,	aRow, 0)
		Call .SetText(.MaxCols,			aRow, UNICDbl(.MaxRows))
    
		.ReDraw = True
	End With
 	
 	frm1.txtTOTSL.value		= ""
	frm1.txtTOTGRQty.text	= 0
	frm1.txtTOTGOHQty.text	= 0
	frm1.txtTOTTempGOHQty.text = 0
End Function

'================================================================
' Function Name : FncCancel
'========================================================================================
Function FncCancel() 
	Dim SEQ1
	Dim CnlVMISLCd   
	Dim CnlBPCd
	Dim CancelGoodOnhandQty 
	Dim i

	ggoSpread.Source = gActiveSpdSheet

	frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
	frm1.vspdData1.Col = frm1.vspdData1.MaxCols
	SEQ1 = UNICDbl(frm1.vspdData1.Text)

	Select Case gActiveSpdSheet.id 
		Case "A"
			With frm1.vspdData1
				If .MaxRows < 1 Then Exit Function	
				Call .SetText(C_SchdRcptQty, .ActiveRow , 0)
				Call .SetText(0, .ActiveRow, "")
			End With
			
			Call CopyToHSheet(SEQ1, 0, 0, 0)
			
			With frm1.vspdData2

				For i = 1 To .MaxRows
					.Row = i
					.Col = C_GoodOnhandQty
					CancelGoodOnhandQty = UNICDbl(.Text)
					Call .SetText(C_TempGRQty, i, 0)
					Call .SetText(C_TempGoodOnhandQty, i, UNIFormatNumber(CancelGoodOnhandQty,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))
					.Col = 0
					If .Text = ggoSpread.InsertFlag Then
						.Action = 0
						ggoSpread.Source = frm1.vspdData2
						ggoSpread.EditUndo
						i = i - 1
					Else
						.Text = ""
						Call .SetText(C_GRQty, i, 0)
					End If
				Next

				.Row = .ActiveRow
				.Col = C_VMISLCd
				CnlVMISLCd = .text
				.Col = C_BPCd
				CnlBPCd = .Text

				Call TOTBPInfo(.ActiveRow, CnlBPCd)
				Call TOTVMISLInfo(CnlVMISLCd)
	
			End With
			
		Case "B"
			Dim Row2
		    Dim INSRow

			With frm1.vspdData2
				If .MaxRows < 1 Then Exit Function	
				.Row = .ActiveRow
				.Col = C_INSNo
				INSRow = UNICDbl(.Text)
				.Col = .MaxCols
				Row2 = UNICDbl(.Text)
				.Col = C_GoodOnhandQty
				CancelGoodOnhandQty = UNICDbl(.text) 

				Call .SetText(C_TempGRQty, .ActiveRow , 0)
				Call .SetText(C_TempGoodOnhandQty, .ActiveRow , UNIFormatNumber(CancelGoodOnhandQty,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))

				ggoSpread.Source = frm1.vspdData2
				ggoSpread.EditUndo

				.Row = .ActiveRow 
				.Col = C_VMISLCd
				CnlVMISLCd = .text
				.Col = C_BPCd
				CnlBPCd = .Text

				Call CopyToHSheet(SEQ1, Row2, INSRow, 0) 
				Call SetTempGOHQtyToGrid2(SEQ1, UNICDbl(.ActiveRow))
				Call SetTempPORemainQtyToGrid2(SEQ1, INSRow, 0)
				Call TOTBPInfo(.ActiveRow, CnlBPCd)
				Call TOTVMISLInfo(CnlVMISLCd)
			End With
		End Select

		Call RestoreToolBar()
End Function

'========================================================================================
' Function Name : FncCopy
'========================================================================================
Function FncCopy() 

    With frm1.vspdData2
    
		If .Maxrows < 1 then exit function
		
		.Row = .ActiveRow
		.Col = C_VMISLCd
 		If .Text = "" then exit function

		.ReDraw = False

		ggoSpread.Source = frm1.vspdData2	
		ggoSpread.CopyRow
		ggoSpread.SpreadLock -1, .ActiveRow ,, .ActiveRow
		
		Call SetSpreadLock(.ActiveRow)
		
		Call .SetText(C_VMISLCd,		.ActiveRow , "")
		Call .SetText(C_VMISLNm,		.ActiveRow , "")
		Call .SetText(C_GoodOnhandQty,	.ActiveRow , 0)
		Call .SetText(C_GRQty,			.ActiveRow , 0)
		Call .SetText(C_TempGoodOnhandQty, .ActiveRow , 0)
		Call .SetText(C_MakerLotNo,		.ActiveRow , "")
		Call .SetText(C_MakerLotSubNo,	.ActiveRow , 0)
		Call .SetText(.MaxCols,			.ActiveRow , UNICDbl(.MaxRows))
   
		.ReDraw = True
		Call RestoreToolBar()
	End With
	
	frm1.txtTOTSL.value		= ""
	frm1.txtTOTGRQty.text	= 0
	frm1.txtTOTGOHQty.text	= 0
	frm1.txtTOTTempGOHQty.text = 0

End Function

Function FncCopy1(ByVal Row) 

    With frm1.vspdData2
    
		If .Maxrows < 1 then exit function
 
		.ReDraw = False

		ggoSpread.Source = frm1.vspdData2	
		ggoSpread.CopyRow Row
		ggoSpread.SpreadLock -1, .ActiveRow ,, .ActiveRow
		
		Call SetSpreadLock(.ActiveRow)
		
		Call .SetText(C_GoodOnhandQty,	.ActiveRow , 0)
		Call .SetText(C_GRQty,			.ActiveRow , 0)
		Call .SetText(C_TempGoodOnhandQty, .ActiveRow , 0)
		Call .SetText(C_MakerLotNo,		.ActiveRow , "")
		Call .SetText(C_MakerLotSubNo,	.ActiveRow , 0)
		Call .SetText(.MaxCols,			.ActiveRow , UNICDbl(.MaxRows))
    
		.ReDraw = True
		Call RestoreToolBar()
	End With
End Function

'========================================================================================
' Function Name : FncPrint
'========================================================================================
Function FncPrint() 
    Call parent.fncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLEMULTI)
End Function

'========================================================================================
' Function Name : FncFind
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)
End Function

'========================================================================================
' Function Name : FncSplitColumn
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then Exit Sub
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)  
End Sub

'========================================================================================
' Function Name : FncExit
'========================================================================================
Function FncExit()

	Dim IntRetCD
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData1						
    If ggoSpread.SSCheckChange = True Then					
		IntRetCD = displaymsgbox("900016", parent.VB_YES_NO, "x", "x")	
		If IntRetCD = vbNo Then	Exit Function
    End If
    
    FncExit = True
End Function

'========================================================================================
' Function Name : DbQuery
'========================================================================================
Function DbQuery() 

    Dim strVal
    
    DbQuery = False
    Call LayerShowHide(1)
    Err.Clear

    With frm1

		If lgIntFlgMode1 = parent.OPMD_UMODE Then
			strVal =  BIZ_PGM_QRY1_ID &	"?txtMode="			& parent.UID_M0001			& _
										"&lgStrPrevKey11="	& lgStrPrevKey11			& _
										"&lgStrPrevKey12="	& lgStrPrevKey12			& _
										"&lgStrPrevKey13="	& lgStrPrevKey13			& _
										"&txtPlantCd="		& Trim(.txthPlantCd.Value)	& _
										"&txtSBPCd="		& Trim(.txthSBPCd.Value)	& _
										"&txtMaxRows="		& .vspdData1.MaxRows
		Else
			strVal =  BIZ_PGM_QRY1_ID & "?txtMode="			& parent.UID_M0001			& _
										"&lgStrPrevKey11=" & lgStrPrevKey11			& _
										"&lgStrPrevKey12=" & lgStrPrevKey12			& _
										"&lgStrPrevKey13=" & lgStrPrevKey13			& _
										"&txtPlantCd="      & Trim(.txtPlantCd.Value)	& _
										"&txtSBPCd="		& Trim(.txtSBPCd.Value)		& _
										"&txtItemCd="       & Trim(.txtItemCd.Value)	& _		
										"&txtMaxRows="      & .vspdData1.MaxRows
		End IF	

		Call RunMyBizASP(MyBizASP, strVal)
		.vspdData1.focus
		
    End With
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()

	lgOldRow1 = 1

	If lgIntFlgMode1 <> parent.OPMD_UMODE Then
		frm1.cbohMvmtType.Value = frm1.cboMvmtType.value
		If DbDtlQuery1(lgOldRow1) = False Then	Exit Function
	End If
    lgIntFlgMode1 = parent.OPMD_UMODE
	SaveCheck = False

End Function

'========================================================================================
' Function Name : DbDtlQuery1
'========================================================================================
Function DbDtlQuery1(ByVal Row) 
	
	Dim IntRetCD
    Dim strVal

	DbDtlQuery1 = False   
    
    Err.Clear

	If frm1.txthPlantCd.value <> "" Then
   		If UCase(Trim(frm1.txtPlantCd.value)) <> UCase(Trim(frm1.txthPlantCd.value)) Then
			IntRetCD = DisplayMsgBox("162059", parent.VB_YES_NO, frm1.txtPlantCd.Alt,"X")
			If IntRetCD = vbNo Then
				frm1.txtPlantCd.focus 
				Exit Function
			End If
			frm1.txtPlantCd.Value = frm1.txthPlantCd.value
		End If
   	End If
   	
   	If frm1.cbohMvmtType.value <> "" Then
   		If UCase(Trim(frm1.cbohMvmtType.value)) <> UCase(Trim(frm1.cboMvmtType.value)) Then
			IntRetCD = DisplayMsgBox("162059",parent.VB_YES_NO, frm1.cboMvmtType.Alt,"X")
			If IntRetCD = vbNo Then
				frm1.cboMvmtType.focus 
				Exit Function
			End If
			frm1.cboMvmtType.Value = frm1.cbohMvmtType.value
		End If
	End If	
 
    Call LayerShowHide(1)
	
    With frm1

	    .vspdData1.Row = Row

		.vspdData1.Col = C_ItemCd
		strItemCd = .vspdData1.Text
		
		.vspdData1.Col = C_SLCd
		strSLCd  = .vspdData1.Text  
		
		.vspdData1.Col = C_TrackingNo
		strTrackingNo  = .vspdData1.Text

		strVal =  BIZ_PGM_QRY2_ID & "?txtMode="			& parent.UID_M0001			& _
									"&txtPlantCd="		& Trim(.txthPlantCd.value)	& _				
									"&txtItemCd="		& Trim(strItemCd)			& _
									"&txtSLCd="			& Trim(strSLCd)				& _
									"&txtTrackingNo="	& Trim(strTrackingNo)		& _
									"&lgStrPrevKey21="	& lgStrPrevKey21			& _
									"&lgStrPrevKey22="	& lgStrPrevKey22			& _
									"&lgStrPrevKey23="	& lgStrPrevKey23			& _
									"&txtMaxRows="		& .vspdData2.MaxRows

		If lgIntFlgMode2 = parent.OPMD_UMODE Then
			strVal = strVal & "&cboMvmtType="	& Trim(.cbohMvmtType.value)			
		Else
			strVal = strVal & "&cboMvmtType="	& Trim(.cboMvmtType.value)			
		End If

		Call RunMyBizASP(MyBizASP, strVal)
	End With
	Set gActiveElement = document.activeElement
    DbDtlQuery1 = True
End Function

Function DbDtlQuery1Ok(ByVal FromNo, ByVal ToNo)

	lgOldRow2 = 1
	DtlFromNo = UNICDbl(FromNo)
	DtlToNo = UNICDbl(FromNo) + UNICDbl(ToNo)
    lgIntFlgMode2 = parent.OPMD_UMODE
    Call SetActiveCell(frm1.vspdData2,C_GRQty,1,"M","X","X")

	With frm1.vspdData1
		.Row = .ActiveRow
		.Col = .MaxCols
		CurrentRow1 = UNICDbl(.Text)

		Call SetVMIItemDetail(DtlFromNo, DtlToNo)
	
		.Row = .ActiveRow
		.Col = 0
		If 	.Text <> "" Then
			If CheckItemMaxRow(CurrentRow1, DtlFromNo, DtlToNo) = True Then Call CopyHiddenToGrid2(CurrentRow1, DtlFromNo)
			Exit Function
		End If
	End With
	
	With frm1.vspdData2
		.Row = 1
		.Col = C_BPCd
		strBPCd = .Text
		Call TOTBPInfo(1, strBPCd)
		
		.Row = 1
		.Col = C_VMISLCd
		If .Text <> "" Then Call TOTVMISLInfo(.Text)
	End With
End Function

'========================================================================================
' Function Name : DbSave
'========================================================================================
Function DbSave() 

    Dim lGrpcnt 
    Dim strVal
	Dim lRow
	Dim ICol
    Dim PvArr     

    DbSave = False
    Call SetToolBar("10000000000011")								

    Call LayerShowHide(1)

    With frm1
		.txtMode.value = parent.UID_M0002
		.txtFlgMode.value = lgIntFlgMode2
	End With

    '-----------------------
    'Data manipulate area
    '-----------------------
    lGrpCnt = 1
    strVal = ""
	
	With frm1.vspdData3

		ReDim PvArr(.MaxRows)

  		For lRow = 1 To .MaxRows
		    .Row = lRow
			strVal = lRow & Parent.gColSep 	
			
			For ICol = 1 To C_HMakerLotSubNo
				.Col = ICol
				strVal = strVal & Trim(.Text) & Parent.gColSep
			Next
			
			.Col = C_HGRQty
			strVal = strVal & UNIConvNum(Trim(.Text), 0) & Parent.gRowSep	

   	   		PvArr(lGrpCnt - 1) = strVal                
            lGrpCnt = lGrpCnt + 1
		Next
    End With
    
	frm1.txtMaxRows.value = lGrpCnt - 1
	frm1.txtSpread.value = Join(PvArr,"")
	
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)
    Call SetToolBar("11000000000011")								

    DbSave = True
    
End Function

'========================================================================================
' Function Name : DbSaveOk
'========================================================================================
Function DbSaveOk()
    Dim DocNo
	DocNo = frm1.txtMvmtNo.Value
    Call DisplayMsgBox("169910","X" ,DocNo, "X") 	
  	Call InitVariables
    Call ggoOper.ClearField(Document, "2")					
	frm1.txtMvmtNo.Value = DocNo
	SaveCheck = True
	Call MainQuery
End Function

'==============================================================================
' Function : SheetFocus
'==============================================================================
Function SheetFocus(lRow)
	Dim Focus1
	Dim Focus2
	DIm i

	With frm1.vspdData3
		.Row = lRow
		.Col = C_HIndex1
		Focus1 = UNICDbl(.Text)
		.Col = C_HIndex2
		Focus2 = UNICDbl(.Text)
	End With
	
	With frm1.vspdData1
		For i = 1 To .MaxRows
			.Row = i  
			.Col = .MaxCols
			If UNICDbl(.Text) = Focus1 Then 
				If UNICDbl(.ActiveRow) <> i Then
					.Col = 1
					.Action = 0
					ggoSpread.Source = frm1.vspdData2
					ggoSpread.ClearSpreadData
					Call DbDtlQuery1(i)
					Exit For
				End If
			End If
		Next
	End With
	
	With frm1.vspdData2
		For i = 1 To .MaxRows
			.Row = i  
			.Col = .MaxCols
			If UNICDbl(.Text) = Focus2 Then 
				If UNICDbl(.ActiveRow) <> i Then
					.Col = C_GRQty	
					.Action = 0
					Exit For
				End If
			End If
		Next
		Call SetFocusToDocument("M") 
		.focus
	End With
End Function

'==============================================================================
' Function : SetTempPORemainQtyToGrid2
' Description : 입고예정량 업데이트 
'==============================================================================
Function SetTempPORemainQtyToGrid2(ByVal SEQ_NO1, ByVal INSRow, ByVal GR_Qty)

	Dim LngGRQty
	Dim LngPORemainQty
	Dim LngGoodOnHandQty
	Dim LngTempGRQty
	Dim LngHIndex1
	Dim LngHINSNo
	Dim i, j
	
	SetTempPORemainQtyToGrid2 = True	
	
	j = 0
	
	frm1.vspdData2.Row = frm1.vspdData2.ActiveRow
	frm1.vspdData2.Col = C_PORemainQty
	LngPORemainQty = UNICDbl(frm1.vspdData2.Text)
	frm1.vspdData2.Col = C_GoodOnhandQty
	LngGoodOnHandQty = UNICDbl(frm1.vspdData2.Text)
	
	With frm1.vspdData3
		For i = 1 To .MaxRows
			.Row = i
			.Col = C_HIndex1
			LngHIndex1 = UNICDbl(.Text)
			.Col = C_HINSNo
			LngHINSNo = UNICDbl(.Text)
			If SEQ_NO1 = LngHIndex1 and INSRow = LngHINSNo Then
				.Col = C_HGRQty
				LngTempGRQty = LngTempGRQty + UNICDbl(.Text)
			End If
		Next
	End With

	With frm1.vspdData2
		If LngTempGRQty > LngPORemainQty Then 
			Call DisplayMsgBox("162056","X","X","X")  
			SetTempPORemainQtyToGrid2 = False	
			Call .SetText(C_GRQty, .ActiveRow , 0)
			Call .SetText(C_TempGoodOnhandQty, .ActiveRow , UNIFormatNumber(LngGoodOnHandQty,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))
			.Row = .ActiveRow
			.Col = 0
			If .Text = ggoSpread.UpdateFlag Then .text = ""

			LngTempGRQty = LngTempGRQty - GR_Qty
			
			frm1.vspdData3.Row = frm1.vspdData3.MaxRows
			frm1.vspdData3.Action = 0
			ggoSpread.Source = frm1.vspdData3
			ggoSpread.EditUndo
		End If
		
		For i = 1 To .MaxRows
			.Row = i
			.Col = C_INSNo
			If INSRow = UNICDbl(.Text) Then
				Call .SetText(C_TempGRQty, i, UNIFormatNumber(LngTempGRQty,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))
				.Col = 0
				If LngTempGRQty = 0 Then
					If .Text = ggoSpread.UpdateFlag Then .Text = ""
				Else
					If .Text <> ggoSpread.InsertFlag Then .Text = ggoSpread.UpdateFlag
				End If 
				j = 1
			Else
				If j = 1 Then Exit For
			End If
		Next
		
		Call SetSchRcptQtyToGrid1
	End With
End Function

'==============================================================================
' Function : SetSchRcptQtyToGrid1
' Description : 그리드 1의 입고예정량 업데이트 
'==============================================================================
Function SetSchRcptQtyToGrid1()

	Dim LngSchRcptQty
	Dim Grid1Row
	Dim i
	
	SetSchRcptQtyToGrid1 = False	

	With frm1.vspdData1

		.Row = .ActiveRow
		.Col = .MaxCols
		Grid1Row = UNICDbl(.Text)
	
		For i = 1 To frm1.vspdData3.MaxRows
			frm1.vspdData3.Row = i
			frm1.vspdData3.Col = C_HIndex1
			If UNICDbl(frm1.vspdData3.Text) = Grid1Row Then 
				frm1.vspdData3.Col = C_HGRQty
				LngSchRcptQty = LngSchRcptQty + UNICDbl(frm1.vspdData3.Text)
			End If
		Next

		Call .SetText(C_SchdRcptQty, .ActiveRow, UNIFormatNumber(LngSchRcptQty,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))
		.Row = .ActiveRow
		.Col = 0
		If LngSchRcptQty = 0 Then
			.Text = ""
		Else
			.Text = ggoSpread.UpdateFlag
		End If 
	End With
	SetSchRcptQtyToGrid1 = True	
End Function

'==============================================================================
' Function : SetTempGOHQtyToGrid2
' Description : 예상재고수량 업데이트 
'==============================================================================
Function SetTempGOHQtyToGrid2(ByVal SEQ_NO1, ByVal Row2)

	Dim LngGRQty
	Dim LngGOHQty
	Dim LngTempGRQty
	Dim LngHMakerLotNo
	Dim LngHMakerLotSubNo
	Dim LngH1
	Dim LngBP
	Dim LngVMISL
	Dim LngMakerLotNo
	Dim LngMakerLotSubNo
	Dim LngHBP
	Dim LngHVMISL
	Dim i, j
	
	SetTempGOHQtyToGrid2 = True	
	j = 0
	
	With frm1.vspdData2
		.Row = Row2
		.Col = C_BPCd
		LngBP = .Text
		.Col = C_VMISLCd
		LngVMISL = .Text
		.Col = C_GoodOnhandQty
		LngGOHQty =  UNICDbl(.Text)
		.Col = C_MakerLotNo
		LngMakerLotNo = .Text
		.Col = C_MakerLotSubNo
		LngMakerLotSubNo = UNICDbl(.Text)
	End With

	With frm1.vspdData3
		For i = 1 To .MaxRows
			.Row = i
			.Col = C_HIndex1
			LngH1 = UNICDbl(.Text)
			.Col = C_HBPCd
			LngHBP = .Text
			.Col = C_HVMISLCd
			LngHVMISL = .Text
			.Col = C_HMakerLotNo
			LngHMakerLotNo = .Text
			.Col = C_HMakerLotSubNo
			LngHMakerLotSubNo = UNICDbl(.Text)
			
			If SEQ_NO1 = LngH1 and LngBP = LngHBP and LngVMISL = LngHVMISL and LngMakerLotNo = LngHMakerLotNo and LngMakerLotSubNo = LngHMakerLotSubNo Then
				.Row = i
				.Col = C_HGRQty
				LngGRQty = LngGRQty + UNICDbl(.Text)
			End If
		Next
	End With

	With frm1.vspdData2
		If LngGRQty > LngGOHQty Then 
			Call Displaymsgbox("162057", "x", "x", "x") 
			SetTempGOHQtyToGrid2 = False	
			Call .SetText(C_GRQty, Row2, 0)
			.Row = Row2
			.Col = 0
			If .Text = ggoSpread.UpdateFlag Then frm1.vspdData3.Text = ""
				
			frm1.vspdData3.Row = frm1.vspdData3.MaxRows
			frm1.vspdData3.Action = 0
			ggoSpread.Source = frm1.vspdData3
			ggoSpread.EditUndo
		Else
			For i = 1 To .MaxRows
				.Row = i
				.Col = C_BPCd
				LngHBP = .Text
				.Col = C_VMISLCd
				LngHVMISL = .Text
				.Col = C_MakerLotNo
				LngHMakerLotNo = .Text
				.Col = C_MakerLotSubNo
				LngHMakerLotSubNo = UNICDbl(.Text)
				If LngBP = LngHBP and LngVMISL = LngHVMISL and LngMakerLotNo = LngHMakerLotNo and LngMakerLotSubNo = LngHMakerLotSubNo Then
					Call .SetText(C_TempGoodOnhandQty, i, UNIFormatNumber((LngGOHQty - LngGRQty),ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))
					j = 1
				Else 
					If j = 1 Then Exit For
				End If
			Next
		End If
	End With		
End Function

'==============================================================================
' Function : TOTVMISLInfo
' Description : vspdData2 출고수량 변경 시 에러출력과 vspdData1 데이터 수정 
'==============================================================================
Function TOTVMISLInfo(ByVal VMISLCd)

	Dim LngTOTSLGRQty
	Dim LngTOTGOHQty
	Dim i, j

	TOTVMISLInfo = False	

	If VMISLCd = "" Then Exit Function
	
	LngTOTGOHQty = UNICDbl(frm1.txtTOTGOHQty.value)
	
	If frm1.txtTOTSL.value <> VMISLCd Then
		If 	CommonQueryRs(" SUM(GOOD_ONHAND_QTY) "," I_VMI_ONHAND_STOCK ", _
						" PLANT_CD = " & FilterVar(frm1.txthPlantCd.Value, "''", "S") & _
						" AND BP_CD = " & FilterVar(strBPCd, "''", "S") & _
						" AND ITEM_CD = " & FilterVar(strItemCd, "''", "S") & _
						" AND TRACKING_NO = " & FilterVar(strTrackingNo, "''", "S") & _
						" AND SL_CD = " & FilterVar(VMISLCd, "''", "S") & _
						" GROUP BY SL_CD " , _
						lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
			lgF0 = Split(lgF0, Chr(11))
			LngTOTGOHQty = UNICDbl(lgF0(0))
		Else
			frm1.txtTOTSL.Value = ""
		End If
	End If
	
	With frm1.vspdData2

		For i = 1 To .MaxRows
			.Row = i
			.Col = C_VMISLCd
			If .Text = VMISLCd Then
				.Col = C_GRQty
				LngTOTSLGRQty = LngTOTSLGRQty + UNICDbl(.Text)
			End If
		Next
		
		frm1.txtTOTSL.value		= VMISLCd
		frm1.txtTOTGRQty.text	= UNIFormatNumber(LngTOTSLGRQty,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
		frm1.txtTOTGOHQty.text	= UNIFormatNumber(LngTOTGOHQty,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
		frm1.txtTOTTempGOHQty.text = UNIFormatNumber((LngTOTGOHQty - LngTOTSLGRQty),ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
	End With
	TOTVMISLInfo = True	
End Function

'==============================================================================
' Function : TOTBPInfo
' Description : vspdData2 출고수량 변경 시 에러출력과 vspdData1 데이터 수정 
'==============================================================================
Function TOTBPInfo(ByVal Row2, ByVal StrBPCd)

	Dim LngTOTTempGRQty2
	Dim LngTOTPORemainQty2
	Dim LngTOTTempPORemainQty2
	DIm INSNo
	Dim RINSNo
	Dim i, j

	TOTBPInfo = False	
	j = 0 	

	With frm1.vspdData2
		.Row = Row2
		.Col = C_INSNo
		INSNo = UNICDbl(.Text)
		.Col = C_PORemainQty
		LngTOTPORemainQty2 = UNICDbl(.Text)
		.Col = C_TempGRQty
		LngTOTTempGRQty2 = UNICDbl(.Text)

		For i = 1 To .MaxRows
			.Row = i
			.Col = C_BPCd
			If .Text = StrBPCd Then
				.Col = C_INSNo
				If UNICDbl(.Text) <> INSNo Then 
					If RINSNo <> UNICDbl(.Text) Then
						RINSNo = UNICDbl(.Text)
						.Col = C_PORemainQty
						LngTOTPORemainQty2 = LngTOTPORemainQty2 + UNICDbl(.Text)
						.Col = C_TempGRQty
						LngTOTTempGRQty2 = LngTOTTempGRQty2 + UNICDbl(.Text)
					End If
				End If
				j = 1
			Else
				If j = 1 Then Exit For
			End If
		Next
		
		LngTOTTempPORemainQty2 = LngTOTPORemainQty2 - LngTOTTempGRQty2

		frm1.txtTOTBP.value				= StrBPCd
		frm1.txtTOTPORemainQty.text		= UNIFormatNumber(LngTOTPORemainQty2,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
		frm1.txtTOTTempGRQty.text		= UNIFormatNumber(LngTOTTempGRQty2,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
		frm1.txtTOTTempPORemainQty.text = UNIFormatNumber(LngTOTTempPORemainQty2,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
	End With
	TOTBPInfo = True	
End Function

'==============================================================================
' Function : CheckPOItem
' Description : 하나의 발주내역에 동일 품목이 중복으로 입고예정이 되어있는지 체크 
'==============================================================================
Function CheckPOItem(ByVal INSRow)

	Dim ChkVMISLCd
	Dim ChkMakerLotNo
	Dim ChkMakerLotSubNo
	Dim i
	
	CheckPOItem = False	

	With frm1.vspdData2
		.Row = .ActiveRow
		.Col = C_MakerLotNo
		ChkMakerLotNo = .Text
		.Col = C_MakerLotSubNo
		ChkMakerLotSubNo = UNICDbl(.Text)
		.Col = C_VMISLCd
		ChkVMISLCd = Trim(.Text)

	End With
	CheckPOItem = True	
End Function

'==============================================================================
' Function : CheckItemMaxRow
' Description : 조회된 품목중에 수정사항 있는지 체크 
'==============================================================================
Function CheckItemMaxRow(ByVal CurrentRow1, ByVal FromNo, ByVal ToNo)
	
	Dim chkBP
	Dim chkIndex3
	Dim i, j

	CheckItemMaxRow = False	

	With frm1.vspdData3
		For i = 1 To .MaxRows
			.Row = i
			.Col = C_HIndex1
			If UNICDbl(.Text) = CurrentRow1 Then
				.Col = C_HINSNo
				If UNICDbl(.Text) >= FromNo and UNICDbl(.Text) < ToNo Then
					CheckItemMaxRow = True
					Exit Function
				End If
			End If
		Next
	End With
End Function

'========================================================================================
' Function Name : CopyHiddenToGrid2
'========================================================================================
Function CopyHiddenToGrid2(ByVal CurrentRow1, ByVal DtlFromNo)
	
	DIm TempIndex2
	DIm HiddenIndex2
	Dim FHGRQty
	Dim FHVMISLCd
	Dim FHVMISLNm
	Dim FHOnhandQty
	Dim FHMakerLotNo
	Dim FHMakerLotSubNo
	Dim FHLotNo
	Dim FHLotSubNo
	DIm HIns
	Dim SEQ_NO1
	Dim i, j
	
	CopyHiddenToGrid2 = False
	
	With frm1
		
		.vspdData2.ReDraw = False

		For i = DtlFromNo To .vspdData2.MaxRows
			.vspdData2.Row = i
			.vspdData2.Col = .vspdData2.MaxCols
			TempIndex2 = UNICDbl(.vspdData2.Text)
			FHGRQty = 0
			For j = 1 To .vspdData3.MaxRows
				.vspdData3.Row = j
				.vspdData3.Col = C_HIndex1
				If UNICDbl(.vspdData3.Text) = CurrentRow1 Then
					.vspdData3.Col = C_HINSNo
					If  UNICDbl(.vspdData3.Text) = TempIndex2 Then
						.vspdData3.Col = C_HVMISLCd
						FHVMISLCd = .vspdData3.Text
						.vspdData3.Col = C_HVMISLNm
						FHVMISLNm = .vspdData3.Text
						.vspdData3.Col = C_HMakerLotNo
						FHMakerLotNo = .vspdData3.Text
						.vspdData3.Col = C_HMakerLotSubNo
						FHMakerLotSubNo = UNICDbl(.vspdData3.Text) 
						.vspdData3.Col = C_HLotNo
						FHLotNo = .vspdData3.Text
						.vspdData3.Col = C_HLotSubNo
						FHLotSubNo = UNICDbl(.vspdData3.Text) 
						.vspdData3.Col = C_HOnhandQty
						FHOnhandQty = UNICDbl(.vspdData3.Text) 
						.vspdData3.Col = C_HIndex2
						HiddenIndex2 = UNICDbl(.vspdData3.Text) 
						.vspdData3.Col = C_HGRQty
						FHGRQty = UNICDbl(.vspdData3.Text)

						If  HiddenIndex2 = TempIndex2 Then
							Call .vspdData2.SetText(0, i, ggoSpread.UpdateFlag)
						Else
							Call FncCopy1(i)
							i = i + 1
							Call .vspdData2.SetText(0, i, ggoSpread.InsertFlag)
						End If
						
						Call .vspdData2.SetText(.vspdData2.MaxCols, i, HiddenIndex2)
						Call .vspdData2.SetText(C_GRQty,			i, UNIFormatNumber(FHGRQty,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))
						Call .vspdData2.SetText(C_VMISLCd,			i, FHVMISLCd)
						Call .vspdData2.SetText(C_VMISLNm,			i, FHVMISLNm)
						Call .vspdData2.SetText(C_GoodOnhandQty,	i, UNIFormatNumber(FHOnhandQty,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))
						Call .vspdData2.SetText(C_MakerLotNo,		i, FHMakerLotNo)
						Call .vspdData2.SetText(C_MakerLotSubNo,	i, UNIFormatNumber(FHMakerLotSubNo,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))
						Call .vspdData2.SetText(C_LotNo,			i, FHLotNo)
						Call .vspdData2.SetText(C_LotSubNo,			i, UNIFormatNumber(FHLotSubNo,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))

						.vspdData1.Row = .vspdData1.ActiveRow
						.vspdData1.Col = .vspdData1.MaxCols
						SEQ_NO1 = UNICDbl(.vspdData1.Text)		

						.vspdData2.Row = i
						.vspdData2.Col = C_INSNo
						HIns = UNICDbl(.vspdData2.Text)
						
						Call SetTempGOHQtyToGrid2(SEQ_NO1, i)
						Call SetTempPORemainQtyToGrid2(SEQ_NO1, HIns, FHGRQty)						
						Call TOTBPInfo(UNICDbl(.vspdData2.ActiveRow), strBPCd)
						Call TOTVMISLInfo(FHVMISLCd)
					End If 
				End If
			Next	
		Next
		.vspdData2.ReDraw = True
	End With
	CopyHiddenToGrid2 = True
End Function

'========================================================================================
' Function Name : CopyToHSheet
'========================================================================================
Function CopyToHSheet(ByVal SeqNo, ByVal Row2, ByVal INSRow, ByVal GR_Qty) 
	
	Dim HItemCd
	Dim HTrackingNo
	Dim HSLCd	
	Dim HBPCd	
	Dim HBasicUnit
	Dim HPOUnit
	Dim HPONo	
	Dim HPOSubNo	
	Dim HVMISLCd
	Dim HVMISLNm
	Dim HOnhandQty	
	Dim HMakerLotNo
	Dim HMakerLotSubNo
	Dim HLotNo
	Dim HLotSubNo
	Dim HIndex1
	Dim HIndex2
	Dim HINSNO
	Dim i 
	
	With frm1.vspdData1
		.Row = .ActiveRow
		.Col = C_ItemCd
		HItemCd = .Text
		.Col = C_BasicUnit
		HBasicUnit = .Text
		.Col = C_TrackingNo
		HTrackingNo = .Text
		.Col = C_SLCd
		HSLCd = .Text
	End With
	
	With frm1.vspdData2
		.Row = .ActiveRow
		.Col = C_BPCd
		HBPCd = .Text
		.Col = C_PONo
		HPONo = .Text
		.Col = C_POSeqNo
		HPOSubNo = UNICDbl(.Text)
		.Col = C_POUnit
		HPOUnit = .Text
		.Col = C_VMISLCd
		HVMISLCd = .Text
		.Col = C_VMISLNm
		HVMISLNm = .Text
		.Col = C_MakerLotNo
		HMakerLotNo = .Text
		.Col = C_MakerLotSubNo
		HMakerLotSubNo = UNICDbl(.Text)
		.Col = C_LotNo
		HLotNo = .Text
		.Col = C_LotSubNo
		HLotSubNo = UNICDbl(.Text)
		.Col = C_GoodOnhandQty
		HOnhandQty = UNICDbl(.Text)
	End With
	
	With frm1.vspdData3
		If GR_Qty = 0 Then
			If .MaxRows = 0 Then Exit Function

			For i = 1 To .MaxRows
				.Row = i
				.Col = C_HIndex1
				HIndex1 = UNICDbl(.Text)
				.Col = C_HIndex2
				HIndex2 = UNICDbl(.Text)
				.Col = C_HINSNo
				HINSNO = UNICDbl(.Text)

				If SeqNo = HIndex1 and Row2 = HIndex2 and INSRow = HINSNO Then
					.Row = i
					.Action = 0
					ggoSpread.Source = frm1.vspdData3
					ggoSpread.EditUndo
					Exit Function
				End If
				
				If SeqNo = HIndex1 and Row2 = 0 and INSRow = 0 Then
					.Row = i
					.Action = 0
					ggoSpread.Source = frm1.vspdData3
					ggoSpread.EditUndo
					i = i - 1
				End If 
			Next
			Exit Function
		Else
			If .MaxRows <> 0 Then
				For i = 1 To .MaxRows
					.Row = i
					.Col = C_HIndex1
					HIndex1 = UNICDbl(.Text)
					.Col = C_HIndex2
					HIndex2 = UNICDbl(.Text)
					.Col = C_HINSNo
					HINSNO = UNICDbl(.Text)
					If SeqNo = HIndex1 and Row2 = HIndex2 and INSRow = HINSNO Then
						Call .SetText(C_HVMISLCd,	i, HVMISLCd)
						Call .SetText(C_HVMISLNm,	i, HVMISLNm)
						Call .SetText(C_HGRQty,		i, UNIFormatNumber(GR_Qty,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))
						Call .SetText(C_HOnhandQty, i, UNIFormatNumber(HOnhandQty,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))
						Call .SetText(C_HMakerLotNo,i, HMakerLotNo)
						Call .SetText(C_HMakerLotSubNo, i, HMakerLotSubNo)
						Call .SetText(C_HLotNo,		i, HLotNo)
						Call .SetText(C_HLotSubNo,	i, HLotSubNo)
						Exit Function
					End If
				Next
			End If
		End If
	End With
					
	With frm1.vspdData3	
		.ReDraw = False
		
		ggoSpread.Source = frm1.vspdData3
	    ggoSpread.InsertRow .MaxRows, 1
		Call .SetText(C_HItemCd,	.MaxRows, HItemCd)
		Call .SetText(C_HTrackingNo,.MaxRows, HTrackingNo)
		Call .SetText(C_HBasicUnit, .MaxRows, HBasicUnit)
		Call .SetText(C_HSLCd,		.MaxRows, HSLCd)
		Call .SetText(C_HBPCd,		.MaxRows, HBPCd)
		Call .SetText(C_HPOUnit,	.MaxRows, HPOUnit)
		Call .SetText(C_HPONo,		.MaxRows, HPONo)
		Call .SetText(C_HPOSeqNo,	.MaxRows, HPOSubNo)
		Call .SetText(C_HVMISLCd,	.MaxRows, HVMISLCd)
		Call .SetText(C_HVMISLNm,	.MaxRows, HVMISLNm)
		Call .SetText(C_HGRQty,		.MaxRows, UNIFormatNumber(GR_Qty,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))
		Call .SetText(C_HOnhandQty, .MaxRows, UNIFormatNumber(HOnhandQty,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))
		Call .SetText(C_HMakerLotNo,.MaxRows, HMakerLotNo)
		Call .SetText(C_HMakerLotSubNo, .MaxRows, HMakerLotSubNo)
		Call .SetText(C_HLotNo,		.MaxRows, HLotNo)
		Call .SetText(C_HLotSubNo,	.MaxRows, HLotSubNo)
		Call .SetText(C_HIndex1,	.MaxRows, SeqNo)
		Call .SetText(C_HIndex2,	.MaxRows, Row2)
		Call .SetText(C_HINSNo,		.MaxRows, INSRow)

		.ReDraw = True
	End With
	frm1.txtMvmtNo.value = ""
End Function

'========================================================================================
' Function Name : SetVMIItemDetail
' Function Desc : vspdData2의 Detail을 가져온다(I_VMI_ONHAND_STOCK)
'========================================================================================
Function SetVMIItemDetail(ByVal DtlFromNo, ByVal DtlToNo)
	
	Dim i

	SetVMIItemDetail = False
		
	With frm1.vspdData2
		.ReDraw = False
		
		For i = DtlFromNo To .MaxRows
			.Row = i
			.Col = C_VMISLCd
			If .Text <> "" Then
		   		Call SetSpreadLock(i)
				SetVMIItemDetail = True
			End If
		Next
		.ReDraw = True
	End With
	
	If 	SetVMIItemDetail = True Then Call SetToolBar("11001001001111")				
		
End Function

'========================================================================================
' Function Name : Plant_Item_POInfo_Check
'========================================================================================
Function Plant_Item_POInfo_Check(ByVal ChkIndex)

	Plant_Item_POInfo_Check = False

	Select Case ChkIndex

		Case 1
			'-----------------------
			'Check Plant CODE		
			'-----------------------
			If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
				lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
					
				Call DisplayMsgBox("125000","X","X","X")
				frm1.txtPlantNm.Value = ""
				frm1.txtPlantCd.focus 
				Exit function
			End If
			lgF0 = Split(lgF0, Chr(11))
			frm1.txtPlantNm.Value = lgF0(0)
			'-----------------------
			'Check BP CODE			
			'-----------------------
			If 	CommonQueryRs(" A.BP_NM "," B_BIZ_PARTNER A, B_STORAGE_LOCATION B ", " A.BP_CD = B.BP_CD AND B.SL_TYPE = " & FilterVar("E", "''", "S") & "  AND B.BP_CD = " & FilterVar(frm1.txtSBPCd.Value, "''", "S"), _
				lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then

				If 	CommonQueryRs(" BP_NM "," B_BIZ_PARTNER ", " BP_CD = " & FilterVar(frm1.txtSBPCd.Value, "''", "S"), _
					lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
					
					Call DisplayMsgBox("229927","X","X","X")
					frm1.txtSBPNm.Value = ""
					frm1.txtSBPCd.focus 
					Exit function
				End If
				lgF0 = Split(lgF0, Chr(11))
				frm1.txtSBPNm.Value = lgF0(0)

				Call DisplayMsgBox("162064","X","X","X")
				frm1.txtSBPCd.focus 
				Exit function
			End If
			lgF0 = Split(lgF0, Chr(11))
			frm1.txtSBPNm.Value = lgF0(0)
			
			If frm1.txtItemCd.Value <> "" Then
				If 	CommonQueryRs(" ITEM_NM "," B_ITEM ", " ITEM_CD = " & FilterVar(frm1.txtItemCd.Value, "''", "S"), _
					lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
					lgF0 = Split(lgF0, Chr(11))
					frm1.txtItemNm.Value = lgF0(0)
				Else
					frm1.txtItemNm.Value = ""
				End If
			End If
		
		Case 2
		
			If 	CommonQueryRs(" PUR_GRP_NM, USAGE_FLG "," B_PUR_GRP ", " PUR_GRP = " & FilterVar(frm1.txtGroupCd.Value, "''", "S"), _
				lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
					
				Call DisplayMsgBox("125100","X","X","X") 
				frm1.txtGroupNm.Value = ""
				frm1.txtGroupCd.focus 
				Exit function
			End If
			lgF0 = Split(lgF0, Chr(11))
			lgF1 = Split(lgF1, Chr(11))
			frm1.txtGroupNm.Value = lgF0(0)
			
			If Trim(lgF1(0)) <> "Y" Then
				Call DisplayMsgBox("125114","X","X","X")
				frm1.txtGroupCd.focus
				Exit function
			End If
	End Select
	Set gActiveElement = document.activeElement		  
	Plant_Item_POInfo_Check = True
End Function
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>사급품 입고등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenGRRef()">입고현황</A> </TD>					
					<TD WIDTH=10></TD>
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
									<TD CLASS=TD5 NOWRAP>공장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12xxxU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>입고형태</TD>
									<TD CLASS=TD6><SELECT Name="cboMvmtType" ALT="입고형태"  STYLE="WIDTH: 150px" tag="12"></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>사급업체</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSBPCd" SIZE=10 MAXLENGTH=7 tag="12xxxU" ALT="사급업체"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSBP()">&nbsp;<INPUT TYPE=TEXT NAME="txtSBPNm" SIZE=20 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>품목</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=18 tag="11xxxU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemCd">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 tag="14"></TD>								
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=10 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
								<TR>
									<TD CLASS=TD5 NOWRAP>입고일</TD>
									<TD CLASS="TD6">
										<script language =javascript src='./js/i1532ma1_OBJECT4_txtGRDt.js'></script>
									</TD>
									<TD CLASS=TD5 NOWRAP>구매그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtGroupCd" SIZE=10 MAXLENGTH=4 tag="22xxxU" ALT="구매그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPurchaseGroup">&nbsp;<INPUT TYPE=TEXT NAME="txtGroupNm" SIZE=20 tag="24"></TD>								
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>구매입고번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtMvmtNo" SIZE=20 MAXLENGTH=18 tag="34xxxU" ALT="구매입고번호"></TD>								
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>								
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
							<TR HEIGHT="50%">
								<TD WIDTH="100%" colspan=4>
									<script language =javascript src='./js/i1532ma1_A_vspdData1.js'></script>
								</TD>
							</TR>
							<TR HEIGHT="50%">
								<TD WIDTH="100%" colspan=4>
									<script language =javascript src='./js/i1532ma1_B_vspdData2.js'></script>
								</TD>
							</TR>
							<TR>
								<TD HEIGHT=5 WIDTH=100% colspan=4></TD>
							</TR>
							<TR>
								<TD WIDTH=50% colspan=2>
									<FIELDSET valign=top>
									<LEGEND>공급처 합계</LEGEND>
										<TABLE CLASS="TB2" CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD CLASS=TD5 NOWRAP>공급처</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTOTBP" SIZE=10 MAXLENGTH=18 tag="24xxxU" ALT="공급처"></TD>																
												<TD CLASS=TD5 NOWRAP>발주잔량</TD>
												<TD CLASS=TD6 NOWRAP>
													<script language =javascript src='./js/i1532ma1_OBJECT1_txtTOTPORemainQty.js'></script>
												</TD>																
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>예상잔량</TD>
												<TD CLASS=TD6 NOWRAP>
													<script language =javascript src='./js/i1532ma1_OBJECT1_txtTOTTempPORemainQty.js'></script>
												</TD>																
												<TD CLASS=TD5 NOWRAP>입고예정량</TD>
												<TD CLASS=TD6 NOWRAP>
													<script language =javascript src='./js/i1532ma1_OBJECT1_txtTOTTempGRQty.js'></script>
												</TD>																
											</TR>
										</TABLE>
									</FIELDSET>
								</TD>
								<TD WIDTH=50% colspan=2>
									<FIELDSET valign=top>
									<LEGEND>VMI창고 합계</LEGEND>
										<TABLE CLASS="TB2" CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD CLASS=TD5 NOWRAP>VMI창고</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTOTSL" SIZE=10 MAXLENGTH=18 tag="24xxxU" ALT="VMI창고"></TD>																
												<TD CLASS=TD5 NOWRAP>재고수량</TD>
												<TD CLASS=TD6 NOWRAP>
													<script language =javascript src='./js/i1532ma1_OBJECT1_txtTOTGOHQty.js'></script>
												</TD>																
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>예상재고량</TD>
												<TD CLASS=TD6 NOWRAP>
													<script language =javascript src='./js/i1532ma1_OBJECT1_txtTOTTempGOHQty.js'></script>
												</TD>																
												<TD CLASS=TD5 NOWRAP>입고수량</TD>
												<TD CLASS=TD6 NOWRAP>
													<script language =javascript src='./js/i1532ma1_OBJECT1_txtTOTGRQty.js'></script>
												</TD>																
											</TR>
										</TABLE>
									</FIELDSET>
								</TD>
							</TR>
						</Table>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<script language =javascript src='./js/i1532ma1_C_vspdData3.js'></script>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txthPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="txthSBPCd" tag="24">
<INPUT TYPE=HIDDEN NAME="cbohMvmtType" tag="24">
<INPUT TYPE=HIDDEN NAME="txthItemCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
