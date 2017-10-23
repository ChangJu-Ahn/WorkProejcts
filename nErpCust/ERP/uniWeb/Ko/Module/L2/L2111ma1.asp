<%@ LANGUAGE="VBSCRIPT" %>
<%
'======================================================================================================
'*  1. Module Name          : 영업 
'*  2. Function Name        : 
'*  3. Program ID           : L2111ma1.asp
'*  4. Program Name         : 수주정보등록(IF)
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/03/29
'*  8. Modified date(Last)  :
'*  9. Modifier (First)     : Seongbae Hwang
'* 10. Modifier (Last)      :
'* 11. Comment              :
'=======================================================================================================
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.Inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT> 
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                             

Const BIZ_PGM_ID = "L2111mb1.asp"            
 
Const C_PopSoldToParty	= 1
Const C_PopSalesGrp		= 1
Const C_PopPlantCd		= 2
Const C_PopVatType		= 3
Const C_TitleForVatIncflag = "VAT포함구분"
<!-- #Include file="../../inc/lgvariables.inc" --> 
Dim	C_Selector
Dim C_RcptDt
Dim C_DocIssueDt
Dim C_SoldToParty 
Dim C_SoldToPartyNm 
Dim C_ShipToParty 
Dim C_ShipToPartyNm 
Dim C_ItemCd
Dim C_ItemNm
Dim C_Qty
Dim C_Unit  
Dim C_SoQty 
Dim C_SoBonusQty  
Dim C_Cur 
Dim C_Price 
Dim C_SoPrice 
Dim C_PriceFlag   
Dim C_PriceFlagNm  
Dim C_VatIncFlag  
Dim C_VatIncFlagNm  
Dim C_ReqDlvyDt   
Dim C_DlvyDt  
Dim C_VatType  
Dim C_VatTypePopup  
Dim C_VatTypeNm  
Dim C_VatRate  
Dim C_PlantCd 
Dim C_PlantCdPopup
Dim C_PlantNm 
Dim C_SlCd  
Dim C_SlCdPopup 
Dim C_SlNm
Dim C_TrackingNo
Dim C_DocNo  
Dim C_DocSeq
Dim C_Remark
Dim C_TrackingFlg 
Dim C_DealType 
Dim C_PayMeth
Dim C_InfNo
Dim C_InfSeq

Dim lgBlnOpenPop
Dim lgStrWhere					' Scrollbar를 조회조건 
Dim lgBlnDisplayMsg				' 수주일이 없는 경우 경고 메세지 Display 여부 
Dim lgBlnDisplayMsgForVatIncflag' VAT포함구분이 없는 경고 메세지 Display 여부 
Dim lgLngStartRow				' Start row to be queryed
Dim lgArrVATTypeInfo			' VAT Type Info.
Dim lgStrBaseDt
Dim lgStrFirstDt

lgStrBaseDt = "<%=GetSvrDate%>"
lgStrFirstDt = UNIConvDateAToB(UNIGetFirstDay(lgStrBaseDt,parent.gServerDateFormat), parent.gServerDateFormat,parent.gDateFormat)
lgStrBaseDt	 = UNIConvDateAToB(lgStrBaseDt, parent.gServerDateFormat,parent.gDateFormat)

'========================================================================================================
Sub initSpreadPosVariables()  
	C_Selector		= 1
	C_RcptDt		= 2
	C_DocIssueDt	= 3
	C_SoldToParty	= 4
	C_SoldToPartyNm	= 5
	C_ShipToParty	= 6
	C_ShipToPartyNm	= 7
	C_ItemCd		= 8
	C_ItemNm		= 9
	C_Qty			= 10
	C_Unit			= 11
	C_SoQty			= 12
	C_SoBonusQty	= 13 
	C_Cur			= 14
	C_Price			= 15
	C_SoPrice		= 16
	C_PriceFlag		= 17 
	C_PriceFlagNm	= 18
	C_VatIncFlag	= 19
	C_VatIncFlagNm	= 20
	C_ReqDlvyDt		= 21
	C_DlvyDt		= 22
	C_VatType		= 23
	C_VatTypePopup  = 24
	C_VatTypeNm		= 25
	C_VatRate		= 26
	C_PlantCd		= 17
	C_PlantCdPopup	= 28
	C_PlantNm		= 29
	C_SlCd			= 30
	C_SlCdPopup		= 31
	C_SlNm			= 32
	C_TrackingNo	= 33
	C_DocNo			= 34
	C_DocSeq		= 35
	C_Remark		= 36
	C_TrackingFlg	= 37
	C_DealType		= 38
	C_PayMeth		= 39
	C_InfNo			= 40
	C_InfSeq		= 41
End Sub

'========================================================================================================
Sub InitVariables()
	on error resume next
    lgIntFlgMode = Parent.OPMD_CMODE            'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey = ""
    lgBlnDisplayMsg = True
    lgBlnDisplayMsgForVatIncflag = True
End Sub

'========================================================================================
Sub InitVATTypeInfo()
	On Error Resume Next

	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iArrVATType, iArrVATTypeNm, iArrVATRate
	Dim iIntIndex
	
	Err.Clear
	
	iStrSelectList	= " Minor.MINOR_CD,  Minor.MINOR_NM, Config.REFERENCE "
	iStrFromList	= " B_MINOR Minor,B_CONFIGURATION Config "
	iStrWhereList	= " Minor.MAJOR_CD=" & FilterVar("B9001", "''", "S") & " And Config.MAJOR_CD = Minor.MAJOR_CD And Config.MINOR_CD = Minor.MINOR_CD And Config.SEQ_NO = 1 " 
	
	If CommonQueryRs(iStrSelectList, iStrFromList, iStrWhereList, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
		iArrVATType		= Split(lgF0, parent.gColSep)
		iArrVATTypeNm	= Split(lgF1, parent.gColSep)
		iArrVATRate		= Split(lgF2, parent.gColSep)
		
		Redim lgArrVATTypeInfo(UBound(iArrVATType) - 1, 2)

		For iIntIndex = 0 to UBound(iArrVATType) - 1
			lgArrVATTypeInfo(iIntIndex, 0) = iArrVATType(iIntIndex)
			lgArrVATTypeInfo(iIntIndex, 1) = iArrVATTypeNm(iIntIndex)
			lgArrVATTypeInfo(iIntIndex, 2) = iArrVATRate(iIntIndex)
		Next
	Else
		If Err.number <> 0 Then
			MsgBox Err.description 
			Err.Clear 
		End If
		Exit Sub
	End If

End Sub

'=========================================================================================================
Sub SetDefaultVal()
    With frm1
		.txtConRcptFromDt.Text = lgStrFirstDt
		.txtConRcptToDt.Text = lgStrBaseDt
		.txtConRcptFromDt.focus
		.txtSoDt.Text = lgStrBaseDt
		.txtSalesGrp.value = parent.gSalesGrp
		.txtPlantCd.value = parent.gPlant
		.chkApplyCurrPrice.checked = True
    End With
End Sub

'=========================================================================================================
Sub SetRowDefaultVal(ByVal pvLngRow)
	Dim iStrData
	Dim iIntValue
	
	With frm1.vspdData
		.Row = pvLngRow
		.Col = C_PlantCd
		If Trim(.Text) = "" And Trim(frm1.txtPlantCd.value) <> "" Then
			If GetItemByPlantInfo(pvLngRow, frm1.txtPlantCd.value, True) Then
				Call SetSpreadUnLockByTrackingFlg(pvLngRow)
			End If
		End If

		' 수량 
		.Col = C_Qty	: iStrData = .Text
		.Col = C_SoQty	: .Text = iStrData
		
		' 단가 Fetch
		If frm1.chkApplyCurrPrice.checked Then
			Call GetItemPrice(pvLngRow)
		Else
			' 단가 
			.Col = C_Price	: iStrData = .Text
			.Col = C_SoPrice		: .Text = iStrData
		End If
		
		' 납기일 
		.Col = C_ReqDlvyDt	: iStrData = .Text
		.Col = C_DlvyDt		: .Text = iStrData
		
		' 부가세유형 
		iStrData = Trim(frm1.txtVatType.value)
		.Col = C_VatType
		If Trim(.Text) = "" And iStrData <> "" Then
			Call SetVATTypeInSpread(pvLngRow, iStrData)
			
			If Trim(frm1.cboVatIncFlag.value) <> "" Then
				.Col = C_VatIncFlag		:	.Text = frm1.cboVatIncFlag.value	:	iIntValue = .Value
				.Col = C_VatIncFlagNm	:	.Value = iIntValue
			Else
				.Col = C_VatIncFlag		:	.Text = "1"			:	iIntValue = .Value
				.Col = C_VatIncFlagNm	:	.Value = iIntValue

				If lgBlnDisplayMsgForVatIncflag Then
					' VAT포함구분이 등록되지 않았습니다. 기본값은 '별도'로 설정합니다. 이 메세지는 더이상 Display되지 않습니다.
					Call DisplayMsgBox("203162", "X", C_TitleForVatIncflag, frm1.cboVatIncFlag(1).innerText)
					lgBlnDisplayMsgForVatIncflag = False
				End If
			End If
		End If
	End With		
	
End Sub

' Copy row
Sub SetRowCopyDefaultVal(ByVal pvLngRow)
End Sub

'==========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub

'==========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables() 
	
	With frm1.vspdData		
			
	   	'☜:--------Spreadsheet #1-----------------------------------------------------------------------------   
		ggoSpread.Source = frm1.vspdData
		'patch version
	    ggoSpread.Spreadinit "V20021214",,parent.gAllowDragDropSpread    		

		.ReDraw = false
			
		.MaxRows = 0 : .MaxCols = 0
		.MaxCols = C_InfSeq + 1            '☜: 최대 Columns의 항상 1개 증가시킴 
		    
	    Call GetSpreadColumnPos("A")

		Call AppendNumberPlace("6","4","0")
		ggoSpread.SSSetCheck	C_Selector, "선택", 6,,,true
	    ggoSpread.SSSetDate		C_RcptDt, "접수일",12,2,Parent.gDateFormat    
	    ggoSpread.SSSetDate		C_DocIssueDt, "주문일",12,2,Parent.gDateFormat    
		ggoSpread.SSSetEdit		C_SoldToParty, "주문처", 18,,,10,2
		ggoSpread.SSSetEdit		C_SoldToPartyNm, "주문처명", 18
		ggoSpread.SSSetEdit		C_ShipToParty, "납품처", 18,,,10,2
		ggoSpread.SSSetEdit		C_ShipToPartyNm, "납품처명", 18
		ggoSpread.SSSetEdit		C_ItemCd,		"품목", 18,,,18,2
		ggoSpread.SSSetEdit		C_ItemNm,		"품목명", 18
		ggoSpread.SSSetFloat	C_Qty,			"주문수량" ,15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit		C_Unit,			"단위", 8,2,,3,2
		ggoSpread.SSSetFloat	C_SoQty,		"수량" ,15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_SoBonusQty,	"덤수량" ,15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit		C_Cur,			"화폐", 8,2,,3,2
		ggoSpread.SSSetFloat	C_Price,		"주문단가",15,"C",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_SoPrice,		"단가",15,"C",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetCombo	C_PriceFlag,	"단가구분", 1
		ggoSpread.SSSetCombo	C_PriceFlagNm,	"단가구분", 10
		ggoSpread.SSSetCombo	C_VatIncFlag,	"VAT포함구분", 1
		ggoSpread.SSSetCombo	C_VatIncFlagNm,	"VAT포함구분", 10
	    ggoSpread.SSSetDate		C_ReqDlvyDt,	"요청납기일",12,2,Parent.gDateFormat    
	    ggoSpread.SSSetDate		C_DlvyDt,		"납기일",12,2,Parent.gDateFormat    
		ggoSpread.SSSetEdit		C_VatType,		"VAT유형", 10,,,5,2
		ggoSpread.SSSetButton	C_VatTypePopup
		ggoSpread.SSSetEdit		C_VatTypeNm,	"VAT유형명", 18
		ggoSpread.SSSetFloat	C_VatRate,		"VAT율",15,parent.ggExchRateNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit		C_PlantCd,		"공장", 10,,,4,2
		ggoSpread.SSSetButton	C_PlantCdPopUp
		ggoSpread.SSSetEdit		C_PlantNm,		"공장명", 18
		ggoSpread.SSSetEdit		C_SlCd,			"창고", 10,,,7,2
		ggoSpread.SSSetButton	C_SlCdPopup
		ggoSpread.SSSetEdit		C_SlNm,			"창고명", 18
		ggoSpread.SSSetEdit		C_TrackingNo,	"Tracking No", 25,,,25
		ggoSpread.SSSetEdit		C_DocNo,		"고객주문번호", 18
		ggoSpread.SSSetFloat	C_DocSeq,		"고객주문순번" ,15,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"  
		ggoSpread.SSSetEdit		C_Remark,		"비고", 1
		ggoSpread.SSSetEdit		C_TrackingFlg,	"Tracking여부", 1
		ggoSpread.SSSetEdit		C_DealType,		"판매유형", 1
		ggoSpread.SSSetEdit		C_PayMeth,		"결제방법", 1
		ggoSpread.SSSetEdit		C_InfNo,		"Interface 번호", 1
		ggoSpread.SSSetEdit		C_InfSeq,		"Interface 순번", 1

		Call ggoSpread.MakePairsColumn(C_PlantCd,C_PlantCdPopUp)
		Call ggoSpread.MakePairsColumn(C_VatType,C_VatTypePopup)
		Call ggoSpread.MakePairsColumn(C_SlCd,C_SlCdPopup)

	    Call ggoSpread.SSSetColHidden(C_PriceFlag, C_PriceFlag,True)
	    Call ggoSpread.SSSetColHidden(C_VatIncFlag, C_VatIncFlag,True)
	    Call ggoSpread.SSSetColHidden(C_Remark, .MaxCols,True)

		Call SetSpreadLock(-1)		    
		.ReDraw = True
	End With
End Sub

'==========================================================================================================
Sub SetSpreadLock(ByVal pvLngRow)
	ggoSpread.SpreadLock 2, pvLngRow, , pvLngRow
End Sub

'==========================================================================================================
Sub SetSpreadUnLock(ByVal pvLngRow)
	ggoSpread.SpreadUnLock C_PlantCd,		pvLngRow, C_PlantCdPopup,	pvLngRow
	ggoSpread.SpreadUnLock C_SoQty,			pvLngRow, C_SoQty,			pvLngRow
	ggoSpread.SpreadUnLock C_SoBonusQty,	pvLngRow, C_SoBonusQty,		pvLngRow
	ggoSpread.SpreadUnLock C_SoPrice,		pvLngRow, C_SoPrice,		pvLngRow
	ggoSpread.SpreadUnLock C_PriceFlagNm,	pvLngRow, C_PriceFlagNm,	pvLngRow
	ggoSpread.SpreadUnLock C_VatIncFlagNm,	pvLngRow, C_VatIncFlagNm,	pvLngRow
	ggoSpread.SpreadUnLock C_DlvyDt,		pvLngRow, C_DlvyDt,			pvLngRow
	ggoSpread.SpreadUnLock C_VatType,		pvLngRow, C_VatTypePopup,	pvLngRow
	ggoSpread.SpreadUnLock C_SlCd,			pvLngRow, C_SlCdPopup,		pvLngRow
	
End Sub

Sub SetSpreadUnLockByTrackingFlg(ByVal pvLngRow)	
	If UCase(frm1.txtHTrackingNORule.value) = "M" Then
		With frm1.vspdData
			.Row = pvLngRow		:	.Col = C_TrackingFlg
			If UCase(.Text) = "Y" Then
				ggoSpread.SpreadUnLock	C_TrackingNo, pvLngRow, C_TrackingNo, pvLngRow
				ggoSpread.SSSetRequired C_TrackingNo, pvLngRow, pvLngRow
			Else
				ggoSpread.SpreadLock		C_TrackingNo, pvLngRow, C_TrackingNo, pvLngRow
				ggoSpread.SSSetProtected	C_TrackingNo, pvLngRow, pvLngRow
			End If
		End With
	End If
End Sub

'==========================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
	' 새로이 등록한 경우 
	ggoSpread.SSSetRequired  C_PlantCd		, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired  C_SoQty		, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired  C_SoPrice		, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired  C_PriceFlagNm	, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired  C_VatIncFlagNm	, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired  C_DlvyDt		, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired  C_VatType		, pvStartRow, pvEndRow
End Sub

' Afetr query
Sub SetQuerySpreadColor(ByVal pvStartRow)
End Sub

'==========================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,Parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Not Frm1.vspdData.ColHidden Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub

'==========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_Selector		= iCurColumnPos(1)
			C_RcptDt		= iCurColumnPos(2)
			C_DocIssueDt	= iCurColumnPos(3)
			C_SoldToParty	= iCurColumnPos(4)
			C_SoldToPartyNm	= iCurColumnPos(5)
			C_ShipToParty	= iCurColumnPos(6)
			C_ShipToPartyNm	= iCurColumnPos(7)
			C_ItemCd		= iCurColumnPos(8)
			C_ItemNm		= iCurColumnPos(9)
			C_Qty			= iCurColumnPos(10)
			C_Unit			= iCurColumnPos(11)
			C_SoQty			= iCurColumnPos(12)
			C_SoBonusQty	= iCurColumnPos(13) 
			C_Cur			= iCurColumnPos(14)
			C_Price			= iCurColumnPos(15)
			C_SoPrice		= iCurColumnPos(16)
			C_PriceFlag		= iCurColumnPos(17) 
			C_PriceFlagNm	= iCurColumnPos(18)
			C_VatIncFlag	= iCurColumnPos(19)
			C_VatIncFlagNm	= iCurColumnPos(20)
			C_ReqDlvyDt		= iCurColumnPos(21)
			C_DlvyDt		= iCurColumnPos(22)
			C_VatType		= iCurColumnPos(23)
			C_VatTypePopup  = iCurColumnPos(24)
			C_VatTypeNm		= iCurColumnPos(25)
			C_VatRate		= iCurColumnPos(26)
			C_PlantCd		= iCurColumnPos(27)
			C_PlantCdPopup	= iCurColumnPos(28)
			C_PlantNm		= iCurColumnPos(29)
			C_SlCd			= iCurColumnPos(30)
			C_SlCdPopup		= iCurColumnPos(31)
			C_SlNm			= iCurColumnPos(32)
			C_TrackingNo	= iCurColumnPos(33)
			C_DocNo			= iCurColumnPos(34)
			C_DocSeq		= iCurColumnPos(35)
			C_Remark		= iCurColumnPos(36)
			C_TrackingFlg	= iCurColumnPos(37)
			C_DealType		= iCurColumnPos(38)
			C_PayMeth		= iCurColumnPos(39)
			C_InfNo			= iCurColumnPos(40)
			C_InfSeq		= iCurColumnPos(41)
    End Select    
End Sub


'==========================================================================================================
Sub InitComboBox()
	' VAT 포함구분 
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD =" & FilterVar("S4035", "''", "S") & " ORDER BY MINOR_NM ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboVatIncFlag,lgF0,lgF1,parent.gColSep)
End Sub

'=========================================================================================================
Sub InitSpreadComboBox()
	Dim iStrCboData    'lgF0
	Dim iStrCboDesc    'lgF1

	ggoSpread.Source = frm1.vspdData
	' VAT포함구분 
	If CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("S4035", "''", "S") & " ORDER BY MINOR_CD ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
		iStrCboData = Replace(lgF0,parent.gColSep,vbTab)
		iStrCboDesc = Replace(lgF1,parent.gColSep,vbTab)
		iStrCboData = Left(iStrCboData,Len(iStrCboData) - 1)
		iStrCboDesc = Left(iStrCboDesc,Len(iStrCboDesc) - 1)
    
		ggoSpread.SetCombo iStrCboData, C_VatIncFlag
		ggoSpread.SetCombo iStrCboDesc, C_VatIncFlagNm
	End If

	' 단가구분 
	ggoSpread.SetCombo		"Y"		& vbTab &		"N",	 C_PriceFlag
	ggoSpread.SetCombo "진단가" & vbTab & "가단가" , C_PriceFlagNm
End Sub

'==========================================================================================================
Sub InitData()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Dim iLngRow
	Dim iIntIndex 
	Dim iIntIndex2 
	
	With frm1.vspdData
		For iLngRow = lgLngStartRow To .MaxRows
			
			.Row = iLngRow
			' 가단가 구분 
			.col = C_PriceFlag		:	iIntIndex = .value
			.col = C_PriceFlagNm	:	.value = iIntIndex
			' 부가세포함여부 
			.col = C_VatIncFlag			:	iIntIndex = .value
			.col = C_VatIncFlagNm		:	.value = iIntIndex
		Next	
	End With
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


'==========================================================================================================
Sub InitDataOnUndo(ByVal pvLngRow)
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Dim iLngRow
	Dim iIntIndex 
	
	With frm1.vspdData
		.Row = pvLngRow
		' 가단가 구분 
		.col = C_PriceFlag		:	iIntIndex = .value
		.col = C_PriceFlagNm	:	.value = iIntIndex
		' 부가세포함여부 
		.col = C_VatIncFlag			:	iIntIndex = .value
		.col = C_VatIncFlagNm		:	.value = iIntIndex
	End With
End Sub

'==========================================================================================================
Sub Form_Load()
	Call LoadInfTB19029             '⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	'----------  Coding part  -------------------------------------------------------------
	Call InitComboBox()
	Call InitSpreadSheet
	call InitSpreadComboBox()
	Call InitVariables              '⊙: Initializes local global variables
	Call InitVATTypeInfo()			' 부가세 유형 정보의 배열을 정의한다.
	Call SetDefaultVal
	' Tracking No. 관리 방법 Fetch
	Call GetNumberingRuleforTracking

	Call SetToolbar("11101001000011")          '⊙: 버튼 툴바 제어 
End Sub

'==========================================================================================
Function txtVatType_OnChange()
	Dim iStrVatType
	
	iStrVatType = Trim(frm1.txtVatType.value)
	If iStrVatType <> "" Then
		If Not SetVATType(iStrVatType) Then
			If Not OpenPopup(C_PopVatType) Then
				txtVatType_OnChange = False
				frm1.txtVatType.value = ""
				frm1.txtVatTypeNm.value = ""
				frm1.txtVatRate.text = ""
			End If
		End If
	End If
End Function


'==========================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    FncQuery = False                                  
    
    Err.Clear             

	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")		
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

    If Not chkField(Document, "1") Then         
       Exit Function
    End If

	If ValidDateCheck(frm1.txtConRcptFromDt, frm1.txtConRcptToDt) = False Then Exit Function

    Call ggoSpread.ClearSpreadData()
    Call InitVariables               

    Call DbQuery                

    FncQuery = True                
        
End Function

'========================================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          
    
    If lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X") 		
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "A")                                                                               
    Call ggoOper.LockField(Document, "N")                                      
    Call InitVariables
	Call SetDefaultVal
    FncNew = True

End Function


'========================================================================================================
Function FncDelete()
	Dim IntRetCD
	FncDelete = False									

	If lgIntFlgMode <> Parent.OPMD_UMODE Then				                             	'Check if there is retrived data
		Call DisplayMsgBox("900002", "X", "X", "X")                                		
		Exit Function
	End If
	
	IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO, "X", "X") 
	If IntRetCD = vbNo Then
		Exit Function
	End If
	'-----------------------
	'Delete function call area
	'-----------------------
	If DBDelete = False Then									
		Exit Function
	End if									'☜: Delete db data
	
	FncDelete = True                                                        					'⊙: Processing is OK
End Function

'========================================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         
    
    Err.Clear                                                               

	ggoSpread.Source = frm1.vspdData 
	If ggoSpread.SSCheckChange = False Then
		IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
		Exit Function
	End If
    
    If Not chkField(Document, "2") Then     
       Exit Function
    End If
    If ggoSpread.SSDefaultCheck = False Then     
       Exit Function
    End If

    CAll  DbSave                                                   
    
    FncSave = True                                                          
    
End Function

'========================================================================================================
Function FncCopy() 

	If frm1.vspdData.MaxRows < 1 Then Exit Function

	FncCopy = False
	
	With frm1.vspdData
		.ReDraw = False
		.focus
			 
		ggoSpread.Source = frm1.vspdData 
		ggoSpread.CopyRow
		
		SetSpreadColor .ActiveRow, .ActiveRow

		Call SetRowCopyDefaultVal(.ActiveRow)
		.ReDraw = True
	End With

	lgBlnFlgChgValue = True

	If Err.number = 0 Then  FncCopy = True				                                
	
    Set gActiveElement = document.ActiveElement   
    
End Function

'========================================================================================================
Function FncCancel() 
	On Error Resume Next
	
	If frm1.vspdData.MaxRows < 1 Then Exit Function
	
	ggoSpread.Source = frm1.vspdData 

	With frm1.vspdData
		.ReDraw = False
		.Row = .ActiveRow
				
		ggoSpread.EditUndo

		Call InitDataOnUndo(.Row)
		Call FormatSpreadCellByCurrency(.Row, .Row, "Q")
		Call SetSpreadLock(.Row)
	
		.ReDraw = True
	End With

End Function

'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
	Dim iLngRowsToInsert,iLngRow
	On Error Resume Next                                                          '☜: If process fails
    Err.Clear
    
    FncInsertRow = False                                                         '☜: Processing is NG

    If IsNumeric(Trim(pvRowCnt)) Then
        iLngRowsToInsert = CInt(pvRowCnt)
    Else
		' 추가할 Row를 User에게 물어봄.
        iLngRowsToInsert = AskSpdSheetAddRowCount()
        If iLngRowsToInsert = "" Then
            Exit Function
        End If
    End If
   
	ggoSpread.Source = frm1.vspdData
	With frm1.vspdData
		.focus
		.ReDraw = False
		ggoSpread.InsertRow, iLngRowsToInsert
				
		SetSpreadColor .ActiveRow, .ActiveRow
		'------ Developer Coding part (Start ) -------------------------------------------------------------- 
		SetRowDefaultVal .ActiveRow
		'------ Developer Coding part (End )   --------------------------------------------------------------	
		.ReDraw = True
						 
		lgBlnFlgChgValue = True
	End With

	If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If    
	
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
Function FncDeleteRow() 
	If frm1.vspdData.MaxRows < 1 Then Exit Function

	Dim iLngRow, iLngFirstRow, iLngLastRow
	Dim iDblQty, iDblAmt
	
	frm1.vspdData.focus
	ggoSpread.Source = frm1.vspdData 
	Call ggoSpread.DeleteRow
		
	lgBlnFlgChgValue = True
End Function

'========================================================================================================
Function FncPrint() 
 Call parent.FncPrint()
End Function

'========================================================================================================
Function FncPrev() 
    On Error Resume Next                                            
End Function

'========================================================================================================
Function FncNext() 
    On Error Resume Next                                            
End Function

'========================================================================================================
Function FncExcel() 
 On Error Resume Next                                                             '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExcel = False                                                              '☜: Processing is NG

	Call parent.FncExport(Parent.C_SINGLEMULTI)	                     			  '☜: 화면 유형 

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
     
    Call parent.FncFind(Parent.C_SINGLEMULTI, False)                              '☜:화면 유형, Tab 유무 
    
    If Err.number = 0 Then	 
       FncFind = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
Sub FncSplitColumn()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
	   Exit Sub
	End If

	ggoSpread.Source = gActiveSpdSheet
	ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    
End Sub

'========================================================================================================
Function FncExit()
 Dim IntRetCD
 FncExit = False

	ggoSpread.Source = frm1.vspdData 
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")		
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	FncExit = True
End Function

'========================================================================================================
Function DbQuery() 
	Err.Clear                                                               
	    
	DbQuery = False                                                         
	If  LayerShowHide(1) = False Then
		Exit Function 
	End If
	    
	Dim iStrVal
	With Frm1
	    If lgIntFlgMode <> Parent.OPMD_UMODE Then    
			' Initial query
			lgStrWhere = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001 & _
						 "&txtWhere=" & "E" & parent.gColSep & _
										Trim(.txtConRcptFromDt.Text) & parent.gColSep & _
										Trim(.txtConRcptToDt.Text) & parent.gColSep & _
										Trim(.txtConSoldToParty.value) & parent.gColSep
		End If
		iStrVal = lgStrWhere & "&lgStrPrevKey=" & lgStrPrevKey & _							   
							   "&txtLastRow=" & frm1.vspdData.MaxRows

		lgLngStartRow = frm1.vspdData.MaxRows + 1
	End With
	Call RunMyBizASP(MyBizASP, iStrVal)            

	DbQuery = True   
                          
End Function

'========================================================================================================
Function DbQueryOk()              
    If lgIntFlgMode <> Parent.OPMD_UMODE Then
		lgIntFlgMode = Parent.OPMD_UMODE            
    End If
    
	If Trim(lgStrPrevKey) = "" Then
		lgStrWhere = ""
    End If

	Call InitData()
	Call FormatSpreadCellByCurrency(lgLngStartRow, frm1.vspdData.MaxRows, "Q")
	frm1.vspdData.focus
End Function


'========================================================================================================
Function DbSave() 

	  Err.Clear                
		 
	  Dim iLngRow
	  Dim iArrData
		
	  DbSave = False                                                          '⊙: Processing is NG
			    
	  On Error Resume Next                                                   '☜: Protect system from crashing
			   
	  If LayerShowHide(1) = False Then
	  	Exit Function 
	  End If
	  
	Redim iArrData(24)

	frm1.txtSpread.value = ""
	
	With frm1.vspdData
		For iLngRow = 1 To .MaxRows
    
			.Row = iLngRow
			.Col = 0

			if .Text = ggoSpread.UpdateFlag Then
				iArrData(0) = CStr(iLngRow)		' Row No.
				.Col = C_InfNo			:		 iArrData(1) = .Text						' Interface 번호 
				.Col = C_InfSeq			:		 iArrData(2) = .Text						' Interface 순번 
				.Col = C_DocSeq			:		 iArrData(3) = .Text						' 발주순번 
				.Col = C_ItemCd			:		 iArrData(4) = .Text						' 품목코드 
				.Col = C_PlantCd		:		 iArrData(5) = .Text						' 공장 
				.Col = C_SlCd			:		 iArrData(6) = .Text						' 창고 
				.Col = C_ShipToParty	:		 iArrData(7) = .Text						' 납품처 
				.Col = C_DlvyDt			:		 iArrData(8) = uniConvDate(.Text)			' 납기일 
				.Col = C_PriceFlag		:		 iArrData(9) = .Text						' 단가구분 
				.Col = C_SoPrice			:		 iArrData(10) = uniConvNum(.Text, 0)		' 단가 
				.Col = C_VatIncFlag		:		 iArrData(11) = .Text						' 부가세 포함여부 
				iArrData(12) = "0"		' Amt.(Doc)
				iArrData(13) = "0"		' Amt.(Loc)
				.Col = C_VatType		:		 iArrData(14) = .Text						' 부가세유형 
				.Col = C_VatRate		:		 iArrData(15) = uniConvNum(.Text, 0)		' 부가세율 
				iArrData(16) = "0"		' VAT Amt.(Doc)
				iArrData(17) = "0"		' VAT Amt.(Loc)
				.Col = C_SoQty			:		 iArrData(18) = uniConvNum(.Text, 0)		' 수주수량 
				.Col = C_SoBonusQty		:		 iArrData(19) = uniConvNum(.Text, 0)		' 수주덤수량 
				
				.Col = C_Unit			:		 iArrData(20) = .Text						' 단위 
				iArrData(21) = "0"		' 과부족허용율(+)
				iArrData(22) = "0"		' 과부족허용율(-)
				.Col = C_Remark			:		 iArrData(23) = .Text						' 비고 
				.Col = C_TrackingNo		:		 iArrData(24) = .Text						' 트랙킹번호 
				frm1.txtSpread.value = frm1.txtSpread.value & Join(iArrData, parent.gColSep) & Parent.gRowSep
			End If
		Next
	End With
 
	With frm1
		.txtMode.value = Parent.UID_M0002
		.txtHeader.value = uniConvDate(.txtSoDt.Text)
		.txtHeader.value = .txtHeader.value & parent.gColSep & Trim(.txtSalesGrp.value)
		.txtHeader.value = .txtHeader.value & parent.gColSep & "E"
		If .chkCfmSo.checked Then
			.txtHeader.value = .txtHeader.value & parent.gColSep & "Y"
		Else
			.txtHeader.value = .txtHeader.value & parent.gColSep & "N"
		End If
	End With

 	Call ExecMyBizASP(frm1, BIZ_PGM_ID)         '☜: 비지니스 ASP 를 가동 
 
    DbSave = True                                                           '⊙: Processing is NG
    
End Function

'========================================================================================================
Function DbSaveOk()               
    Call ggoSpread.ClearSpreadData()
End Function

'========================================================================================================
Function DbDelete() 
    On Error Resume Next                                            
	Err.Clear                                                               					

   	Call LayerShowHide(1)

	DbDelete = False									'⊙: Processing is NG
	
	Dim iStrDel
	
	With frm1
		iStrDel = iStrDel & "0" & parent.gColSep & _
							.txtProgId.value & parent.gColSep & _
							.cboLanguage.value & parent.gColSep & _
							.cboTypeCd.value & parent.gColSep & _
							.cboSpdNo.value & Parent.gRowSep

		.txtMode.value		= Parent.UID_M0003							'☜: 비지니스 처리 ASP 의 상태 
		.txtSpreadDel.value = iStrDel
	End With
	
 	Call ExecMyBizASP(frm1, BIZ_PGM_ID)         '☜: 비지니스 ASP 를 가동 
	
	DbDelete = True			                                                   			'⊙: Processing is NG
End Function

'========================================================================================================
Function DbDeleteOk()              
    On Error Resume Next                                            
	Call MainNew()
End Function


'========================================================================================================
Function OpenConPopup(ByVal pvIntWhere)
	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	OpenConPopup = False

	If lgBlnOpenPop Then Exit Function

	lgBlnOpenPop = True
	
	With frm1
		Select Case pvIntWhere
			' 주문처 
			Case C_PopSoldToParty												
				iArrParam(1) = "B_BIZ_PARTNER BP"								<%' TABLE 명칭 %>
				iArrParam(2) = Trim(.txtConSoldToParty.value)					<%' Code Condition%>
				iArrParam(3) = ""												<%' Name Cindition%>
				iArrParam(4) = "BP.USAGE_FLAG=" & FilterVar("Y", "''", "S") & "  AND BP.BP_TYPE LIKE " & FilterVar("C%", "''", "S") & " "	<%' Where Condition%>
				iArrParam(5) = .txtConSoldToParty.alt							<%' TextBox 명칭 %>
					
				iArrField(0) = "ED15" & Parent.gColSep & "BP.BP_CD"
				iArrField(1) = "ED30" & Parent.gColSep & "BP.BP_NM"
				    
				iArrHeader(0) = .txtConSoldToParty.alt
				iArrHeader(1) = .txtConSoldToPartyNm.alt

				.txtConSoldToParty.focus

		End Select
	End With
 
	iArrParam(0) = iArrParam(5)							<%' 팝업 명칭 %> 

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPop = False
	
	If iArrRet(0) = "" Then
		Exit Function
	Else
		OpenConPopup = SetConPopup(iArrRet,pvIntWhere)
	End If	

End Function

' 등록 Sheet Popup
Function OpenPopup(ByVal pvIntWhere)
	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	OpenPopup = False
	If lgBlnOpenPop Then Exit Function

	lgBlnOpenPop = True
	
	With frm1
		Select Case pvIntWhere
			Case C_PopSalesGrp												
				iArrParam(1) = "B_SALES_GRP"					<%' TABLE 명칭 %>
				iArrParam(2) = Trim(.txtSalesGrp.value)			<%' Code Condition%>
				iArrParam(3) = ""								<%' Name Cindition%>
				iArrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "					<%' Where Condition%>
				iArrParam(5) = .txtSalesGrp.alt					<%' TextBox 명칭 %>
				
				iArrField(0) = "ED15" & Parent.gColSep & "SALES_GRP"						<%' Field명(0)%>
				iArrField(1) = "ED30" & Parent.gColSep & "SALES_GRP_NM"					<%' Field명(1)%>
    
			    iArrHeader(0) = .txtSalesGrp.alt					<%' Header명(0)%>
			    iArrHeader(1) = .txtSalesGrpNm.alt					<%' Header명(1)%>

				.txtSalesGrp.focus 

			Case C_PopPlantCd
				iArrParam(1) = "B_PLANT"
				iArrParam(2) = Trim(.txtPlantCd.value)
				iArrParam(3) = ""
				iArrParam(4) = ""
				iArrParam(5) = .txtPlantCd.alt
				
				iArrField(0) = "ED15" & Parent.gColSep & "PLANT_CD"
				iArrField(1) = "ED30" & Parent.gColSep & "PLANT_NM"
    
			    iArrHeader(0) = .txtPlantCd.alt
			    iArrHeader(1) = .txtPlantNm.alt

				.txtPlantCd.focus 

			' VAT유형 
			Case C_PopVatType
				iArrParam(1) = "B_MINOR MI INNER JOIN B_CONFIGURATION CF ON (CF.MAJOR_CD = MI.MAJOR_CD AND CF.MINOR_CD = MI.MINOR_CD) "
				iArrParam(2) = Trim(.txtVatType.value)
				iArrParam(3) = ""
				iArrParam(4) = "MI.MAJOR_CD = " & FilterVar("B9001", "''", "S") & " AND CF.SEQ_NO = 1 "
				iArrParam(5) = .txtVatType.alt
					
				iArrField(0) = "ED10" & Parent.gColSep & "MI.MINOR_CD"
				iArrField(1) = "ED25" & Parent.gColSep & "MI.MINOR_NM"
				iArrField(2) = "ED10" & Parent.gColSep & "CF.REFERENCE"
				    
				iArrHeader(0) = .txtVatType.alt
				iArrHeader(1) = .txtVatTypeNm.alt
				iArrHeader(2) = .txtVatRate.alt

				.txtVatType.focus
		End Select
 
	End With

	iArrParam(0) = iArrParam(5)							<%' 팝업 명칭 %> 

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPop = False
	
	If iArrRet(0) = "" Then
		Exit Function
	Else
		OpenPopup = SetPopup(iArrRet,pvIntWhere)
	End If	

End Function

'===============================================================================================================
Function OpenSpreadPopup(ByVal pvLngCol, ByVal pvLngRow)
	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	OpenSpreadPopup = False
	
	If lgBlnOpenPop Then Exit Function

	lgBlnOpenPop = True
	
	With frm1.vspdData
		.Row = pvLngRow		:	.Col = pvLngCol
		
		Select Case pvLngCol
			' 공장 
			Case C_PlantCd
				iArrParam(1) = "B_PLANT PT INNER JOIN B_ITEM_BY_PLANT IP ON (IP.PLANT_CD = PT.PLANT_CD) " & _
								" INNER JOIN B_STORAGE_LOCATION SL ON (SL.PLANT_CD = IP.PLANT_CD AND SL.SL_CD = IP.ISSUED_SL_CD) "							' FROM Clause
				iArrParam(2) = .Text													' Code Condition
				iArrParam(3) = ""														' Name Cindition
				
				.Col = C_ItemCd
				iArrParam(4) = "IP.ITEM_CD = '"	& Replace(.Text, "'", "''") & "'"		' Where Condition
				
				iArrField(0) = "ED15" & Parent.gColSep & "PT.PLANT_CD"		' 공장 
				iArrField(1) = "ED30" & Parent.gColSep & "PT.PLANT_NM"		' 공장명 
				iArrField(2) = "ED30" & Parent.gColSep & "IP.ISSUED_SL_CD"	' 창고 
				iArrField(3) = "ED30" & Parent.gColSep & "SL.SL_NM"			' 창고명 
				iArrField(4) = "ED30" & Parent.gColSep & "IP.TRACKING_FLG"	' 트랙킹 관련 여부 

				.Row = 0
				iArrHeader(0) = .Text								' Header명(0)
				.Col = C_PlantNm	:	iArrHeader(1) = .Text		' Header명(1)

			' 부가세유형 
			Case C_VatType
				iArrParam(1) = "B_MINOR MI INNER JOIN B_CONFIGURATION CF ON (CF.MAJOR_CD = MI.MAJOR_CD AND CF.MINOR_CD = MI.MINOR_CD) "
				iArrParam(2) = .Text
				iArrParam(3) = ""
				iArrParam(4) = "MI.MAJOR_CD = " & FilterVar("B9001", "''", "S") & " AND CF.SEQ_NO = 1 "
					
				iArrField(0) = "ED15" & Parent.gColSep & "MI.MINOR_CD"
				iArrField(1) = "ED30" & Parent.gColSep & "MI.MINOR_NM"
				iArrField(2) = "ED8" & Parent.gColSep & "CF.REFERENCE"
				    
				.Row = 0
				iArrHeader(0) = .Text
				.Col = C_VatTypeNm	:	iArrHeader(1) = .Text
				.Col = C_VatRate	:	iArrHeader(2) = .Text

			' 창고 
			Case C_SlCd
				iArrParam(1) = "B_STORAGE_LOCATION "
				iArrParam(2) = .Text
				iArrParam(3) = ""		
				.Col = C_PlantCd
				iArrParam(4) = " PLANT_CD =  " & FilterVar(.Text, "''", "S") & ""
				
				iArrField(0) = "ED15" & Parent.gColSep & "SL_CD"
				iArrField(1) = "ED30" & Parent.gColSep & "SL_NM"
				    
				.Row = 0	: .Col = C_SlCd
				iArrHeader(0) = .Text
				.Col = C_SlNm	:	iArrHeader(1) = .Text
		End Select
	End With
 
	iArrParam(0) = iArrHeader(0)							' 팝업 명칭 
	iArrParam(5) = iArrHeader(0)							' 조회조건 TextBox 명칭 

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPop = False
	If iArrRet(0) = "" Then
		Exit Function
	Else
		OpenSpreadPopup = SetSpreadPopup(iArrRet,pvLngCol, pvLngRow)
	End If	

End Function

'===========================================================================================================
Function SetConPopup(Byval pvArrRet,ByVal pvIntWhere)
	SetConPopup = False

	With frm1
		Select Case pvIntWhere
		Case C_PopSoldToParty
			.txtConSoldToParty.value = pvArrRet(0) 
			.txtConSoldToPartyNm.value = pvArrRet(1)   
			
		End Select
	End With

	SetConPopup = True
End Function

'===========================================================================================================
Function SetPopup(Byval pvArrRet,ByVal pvIntWhere)
	SetPopup = False

	With frm1
		Select Case pvIntWhere
		Case C_PopSalesGrp
			.txtSalesGrp.value = pvArrRet(0) 
			.txtSalesGrpNm.value = pvArrRet(1)   
			
		Case C_PopPlantCd
			.txtPlantCd.value = pvArrRet(0) 
			.txtPlantNm.value = pvArrRet(1)   

		Case C_PopVatType
			.txtVatType.value = pvArrRet(0) 
			.txtVatTypeNm.value = pvArrRet(1)   
			.txtVatRate.value = pvArrRet(2)   

		End Select
		
	End With

	SetPopup = True
End Function
'===========================================================================================================
Function SetSpreadPopup(Byval pvArrRet,ByVal pvLngCol, ByVal pvLngRow)
	SetSpreadPopup = False

	With frm1.vspdData
		.Row = pvLngRow		:	.Col = pvLngCol
		Select Case pvLngCol
			Case C_PlantCd
				.Text = pvArrRet(0)
				.Col = C_PlantNm		: .Text = pvArrRet(1)
				.Col = C_SlCd			: .Text = pvArrRet(2)
				.Col = C_SlNm			: .Text = pvArrRet(3)
				.Col = C_TrackingFlg	: .Text = pvArrRet(4)

			Case C_VatType
				.Text = pvArrRet(0)
				.Col = C_VatTypeNm		: .Text = pvArrRet(1)
				.Col = C_VatRate		: .Text = pvArrRet(2)

			Case C_SlCd
				.Text = pvArrRet(0)
				.Col = C_SlNm			: .Text = pvArrRet(1)
		End Select
	End With

	SetSpreadPopup = True
End Function

'========================================================================================
Function SetVATType(ByVal pvStrData)
	Dim iIntIndex
	
	SetVATType = False
	For iIntIndex = 0 To Ubound(lgArrVATTypeInfo, 1)
		If UCase(lgArrVATTypeInfo(iIntIndex, 0)) = UCase(pvStrData) Then
			With frm1
				.txtVatType.Value = UCase(pvStrData)
				.txtVatTypeNm.Value = lgArrVATTypeInfo(iIntIndex, 1)
				.txtVatRate.Text = lgArrVATTypeInfo(iIntIndex, 2)
			End With
			SetVATType = True
			Exit Function
		End If
	Next
End Function


'========================================================================================
Function SetVATTypeInSpread(ByVal pvLngRow, ByVal pvStrData)
	Dim iIntIndex

	SetVATTypeInSpread = False
	For iIntIndex = 0 To Ubound(lgArrVATTypeInfo, 1)
		If UCase(lgArrVATTypeInfo(iIntIndex, 0)) = UCase(pvStrData) Then
			With frm1.vspdData
				.Row = pvLngRow
				.Col = C_VatType	:	.Text = lgArrVATTypeInfo(iIntIndex, 0)
				.Col = C_VatTypeNm	:	.Text = lgArrVATTypeInfo(iIntIndex, 1)
				.Col = C_VatRate	:	.Text = lgArrVATTypeInfo(iIntIndex, 2)
			End With
			SetVATTypeInSpread = True
			Exit Function
		End If
	Next
End Function

'===========================================================================================================
Sub SetRowStatus(intRow)
    ggoSpread.UpdateRow intRow
	lgBlnFlgChgValue = True
End Sub


'========================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call InitSpreadComboBox()
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
	 '------ Developer Coding part (Start ) --------------------------------------------------------------
	If gMouseClickStatus = "SPCRP" Then Call FormatSpreadCellByCurrency(-1, -1, "Q")	
    '------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

 	With frm1.vspdData 
		If Row > 0 Then
			Select Case Col
				Case C_Selector
					If ButtonDown = 0 then	'---선택이 취소된 경우 
						Call FncCancel()				
					Else	'--- 선택된 경우 
						ggoSpread.UpdateRow Row
						Call SetRowDefaultVal(Row)
						Call SetSpreadUnLock(Row)
						Call SetSpreadColor(Row, Row)
					End if
				
				CASE C_PlantCdPopup
					If OpenSpreadPopup(C_PlantCd, Row) Then
						Call SetSpreadUnLockByTrackingFlg(Row)
					End If

				CASE C_VatTypePopup
					Call OpenSpreadPopup(C_VatType, Row)

				CASE C_SlCdPopup
					Call OpenSpreadPopup(C_SlCd, Row)
			End Select
		End If
	End With
	Call SetActiveCell(frm1.vspdData,Col - 1,Row,"M","X","X")
 
End Sub

'========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	Call SetPopupMenuItemInf("0000111111") 
    gMouseClickStatus = "SPC"
    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then 
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
		
		Exit Sub
	End If    	
End Sub


'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
	Dim iStrData, iStrQty
	Dim iDblQty, iDblSoQty

	With frm1.vspdData
		.Row = Row
		.Col = Col	: iStrData = .Text
		
		If iStrData = "" Then Exit Sub
		
		Select Case Col
			Case C_SoQty
				iDblSoQty = UNICdbl(iStrData)
				.Col = C_Qty	:	iStrQty = .Text	:	iDblQty = UNICDbl(iStrQty)
				If iDblSoQty > iDblQty Then
					' 수주수량이 주문수량을 초과하였습니다.
					Call DisplayMsgBox("203160", "X", "X", "")
					.Col = C_SoQty	:	.Text = iStrQty
				End If

				If iDblSoQty = 0 Then
					' 수주량은 0보다 작을 수 없습니다.
					Call DisplayMsgBox("203160", "X", "X", "")
					.Col = C_SoQty	:	.Text = iStrQty
				End If
				
			Case C_PlantCd
				If GetItemByPlantInfo(Row, iStrData, False) Then
					Call SetSpreadUnLockByTrackingFlg(Row)
				Else
					.Text = ""
				End If
				
			Case C_VatType
				If Not SetVATTypeInSpread(Row, iStrData) Then
					If Not OpenSpreadPopup(C_VatType, Row) Then
						.Row = Row
						.Col = C_VatType	:	.Text = ""
					End If
				End If
		End Select
	End With

End Sub


'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim iIntIndex

	 '---------- Coding part -------------------------------------------------------------
	With frm1.vspddata
		.Row = Row
		
		Select Case Col
			Case C_PriceFlagNm 
				.Col = Col			:	iIntIndex = .Value
				.Col = C_PriceFlag	:	.Value = iIntIndex
			Case C_VatIncFlagNm
				.Col = Col			:	iIntIndex = .Value
				.Col = C_VatIncFlag		:	.Value = iIntIndex
		End Select
		
	End With
End Sub


'========================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
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

'========================================================================================================
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
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )    
    If OldLeft <> NewLeft Then Exit Sub
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    
		If CheckRunningBizProcess = True Then Exit Sub	
    	If lgStrPrevKey <> "" Then								
           Call DisableToolBar(parent.TBC_QUERY)
           Call DBQuery
    	End If
    End If    
End Sub


'========================================================================================================
Sub txtConRcptFromDt_DblClick(Button)
	If Button = 1 Then
       Frm1.txtConRcptFromDt.Action = 7
       Call SetFocusToDocument("M")	
       Frm1.txtConRcptFromDt.Focus
	End If
End Sub


'========================================================================================================
Sub txtConRcptToDt_DblClick(Button)
	If Button = 1 Then
       Frm1.txtConRcptToDt.Action = 7
       Call SetFocusToDocument("M")	
       Frm1.txtConRcptToDt.Focus
	End If
End Sub

'========================================================================================================
Sub txtSoDt_DblClick(Button)
	If Button = 1 Then
       Frm1.txtSoDt.Action = 7
       Call SetFocusToDocument("M")	
       Frm1.txtSoDt.Focus
	End If
End Sub

'========================================================================================================
Sub txtConRcptFromDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub

'========================================================================================================
Sub txtConRcptToDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub

'===========================================================================================================
Function GetItemByPlantInfo(ByVal pvLngRow, ByVal pvStrData, ByVal pvBlnChkFlag)
	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs, iStrPlantCd
	Dim iArrItemInfo
	
	Err.Clear
	    
	GetItemByPlantInfo = False
	
	If pvStrData = "" Then Exit Function
	
	iStrSelectList = " PT.PLANT_CD, PT.PLANT_NM, IP.ISSUED_SL_CD, SL.SL_NM, IP.TRACKING_FLG "
	iStrFromList   = " B_PLANT PT " & _
					 " INNER JOIN B_ITEM_BY_PLANT IP ON (IP.PLANT_CD = PT.PLANT_CD) " & _
					 " INNER JOIN B_STORAGE_LOCATION SL ON (SL.PLANT_CD = IP.PLANT_CD AND SL.SL_CD = IP.ISSUED_SL_CD) "

	With frm1.vspdData
		.Row = pvLngRow	:	.Col = C_ItemCd
		iStrWhereList  = " IP.PLANT_CD =  " & FilterVar(pvStrData, "''", "S") & "" & _
						 " AND IP.ITEM_CD =  " & FilterVar(.Text, "''", "S") & " "

		'품목정보 Fetch
		If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
			iArrItemInfo = Split(iStrRs, parent.gColSep)
			.Col = C_PlantCd	: .text = Trim(iArrItemInfo(1))
			.Col = C_PlantNm	: .text = Trim(iArrItemInfo(2))
			.Col = C_SlCd		: .text = Trim(iArrItemInfo(3))
			.Col = C_SlNm		: .text = Trim(iArrItemInfo(4))
			.Col = C_TrackingFlg: .text = UCase(Trim(iArrItemInfo(5)))

			GetItemByPlantInfo = True
			Exit Function
		Else
			If Err.number = 0 Then
				If Not pvBlnChkFlag Then GetItemByPlantInfo = OpenSpreadPopup(C_PlantCd, pvLngRow)
			Else
				MsgBox Err.description 
				Err.Clear
				Exit Function
			End If
		End If
	End With
		
End Function

'========================================================================================================
Function GetItemPrice(ByVal pvLngRow)
	Dim iStrSoldToParty, iStrItemCd, iStrUnit, iStrPayMeth, iStrDealType, iStrCurrency, iStrSoDt
	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrPrice

	With frm1.vspdData
		.Row = pvLngRow
		.col = C_ItemCd		:	iStrItemCd = Replace(.text,"'","''")				'품목코드 
		.Col = C_Unit		:	iStrUnit = Replace(.text,"'","''")					'단위 
		.Col = C_SoldToParty:	iStrSoldToParty = Replace(.text,"'","''")			'주문처 
		.Col = C_Cur		:	iStrCurrency = Replace(.text,"'","''")				'화폐단위 
		.Col = C_PayMeth	:	iStrPayMeth = Replace(.text,"'","''")				'결제방법 
		.Col = C_DealType	:	iStrDealType = Replace(.text,"'","''")				'판매유형 
		iStrSoDt = UniConvDateToYYYYMMDD(frm1.txtSoDt.Text, parent.gDateFormat,"")		' 수주일 
		
		If Trim(iStrSoDt) = "" Then
			' 수주일이 등록되지 않았습니다. 단가는 현재일 기준을 Patch합니다.
			If lgBlnDisplayMsg Then
				Call DisplayMsgBox("203157", "X", lgStrBaseDt, "")
				lgBlnDisplayMsg = False
			End If

			iStrSoDt = UniConvDateToYYYYMMDD(lgStrBaseDt, parent.gDateFormat,"")		' 수주일 
		End If
		
		iStrSelectList = " dbo.ufn_s_GetItemSalesPrice( " & FilterVar(iStrSoldToParty, "''", "S") & ",  " & FilterVar(iStrItemCd, "''", "S") & ", '" &iStrDealType& "',  " & FilterVar(iStrPayMeth, "''", "S") & "," & _
		    " " & FilterVar(iStrUnit, "''", "S") & ",  " & FilterVar(iStrCurrency, "''", "S") & ",  " & FilterVar(iStrSoDt, "''", "S") & ")"
		iStrFromList  = ""
		iStrWhereList = ""

		Err.Clear
    
		'품목정보 단가 Fetch
		If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
			iArrPrice = Split(iStrRs, parent.gColSep)

			.Col = C_SoPrice
			.text = UNIConvNumPCToCompanyByCurrency(iArrPrice(1), iStrCurrency, Parent.ggUnitCostNo, "X" , "X")
		Else
			If Err.number <> 0 Then
				MsgBox Err.description 
				Err.Clear 
			End If
		End if 
 	End With

End Function

'=================================================================================================
Function GetNumberingRuleforTracking()
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	Dim iArrCode

    Err.Clear

	Call CommonQueryRs(" MINOR_CD", " B_CONFIGURATION ", "  major_cd = " & FilterVar("S0024", "''", "S") & " and seq_no = 1 and Reference = " & FilterVar("Y", "''", "S") & " ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If Len(lgF0) Then 
	    iArrCode = Split(lgF0, parent.gColSep)

		frm1.txtHTrackingNoRule.value = iArrCode(0)	
	Else
		If Err.number = 0 Then
			' 설정되어 있지 않은 경우 자동으로 처리 
			frm1.txtHTrackingNoRule.value = "A"
		Else
			MsgBox Err.description 
			Err.Clear 
			Exit Function
		End If
	End If	

End Function

' 화폐별로 Cell Formating을 재설정한다.
Sub FormatSpreadCellByCurrency(ByVal pvLngStartRow, ByVal pvLngEndRow, ByVal pvStrEditMode)
	Dim iLngPointer
	Dim iStrCur
	
	' 입력인 경우 
	If pvStrEditMode = "I" Then
		Call FixDecimalPlaceByCurrency(frm1.vspdData,pvLngStartRow,C_Cur,C_SoPrice,"C" ,"X","X")				
	End If
	
	Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData,pvLngStartRow,pvLngEndRow,C_Cur,C_Price,"C" ,"I","X","X")         
	Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData,pvLngStartRow,pvLngEndRow,C_Cur,C_SoPrice,"C" ,"I","X","X")         
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
</HEAD>

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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>수주정보등록(IF)</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
							</TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
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
								<TR>
									<TD CLASS="TD5" NOWRAP>접수일</TD>									
									<TD CLASS="TD6" NOWRAP>							        
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
											<script language =javascript src='./js/l2111ma1_OBJECT1_txtConRcptFromDt.js'></script>
											</TD>
											<TD>
											&nbsp;~&nbsp;
											</TD>
											<TD>
											<script language =javascript src='./js/l2111ma1_OBJECT2_txtConRcptToDt.js'></script>
											</TD>
										</TR>
									</TABLE>							        
							        </TD>
									<TD CLASS="TD5" NOWRAP>주문처</TD>
									<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConSoldToParty" SIZE=10 MAXLENGTH=5 tag="11NXXU" ALT="주문처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConSoldToParty" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopSoldToParty) ">
															<INPUT TYPE=TEXT NAME="txtConSoldToPartyNm" SIZE=20 tag="14" ALT="주문처명"></TD>
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
								<TD CLASS=TD5 NOWRAP>수주일</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/l2111ma1_fpDateTime1_txtSoDt.js'></script>
								</TD>
								<TD CLASS="TD5" NOWRAP>영업그룹</TD>
								<TD CLASS="TD6"><INPUT NAME="txtSalesGrp" ALT="영업그룹" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGrp" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(C_PopSalesGrp)">&nbsp;<INPUT NAME="txtSalesGrpNm" TYPE="Text" alt="영업그룹명" MAXLENGTH="50" SIZE=25 tag="24"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>공장</TD>
								<TD CLASS="TD6"><INPUT NAME="txtPlantCd" ALT="공장" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(C_PopPlantCd)">&nbsp;<INPUT NAME="txtPlantNm" TYPE="Text" alt="공장명" MAXLENGTH="50" SIZE=25 tag="24"></TD>
								<TD CLASS=TD5 NOWRAP>VAT포함구분</TD>
								<TD CLASS="TD6"><SELECT NAME="cboVatIncFlag" tag="21X"><Option></Option></SELECT></TD>									
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>VAT유형</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtVatType" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="21XXXU" ALT="VAT유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopUp(C_PopVatType)">&nbsp;<INPUT NAME="txtVatTypeNm" TYPE="Text" alt="VAT유형명" MAXLENGTH="20" SIZE=25 tag="24"></TD>
								<TD CLASS="TD5" NOWRAP>VAT율</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/l2111ma1_fpDoubleSingle6_txtVatRate.js'></script>&nbsp;<LABEL><b>%</b></LABEL></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>최신단가적용</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=CHECKBOX NAME="chkApplyCurrPrice" tag="21" Class="Check"></TD>
								<TD CLASS=TD5 NOWRAP>수주확정</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=CHECKBOX NAME="chkCfmSo" tag="21" Class="Check"></TD>
							</TR>
							<TR>
								<TD HEIGHT="100%" WIDTH="100%" COLSPAN=4>
									<script language =javascript src='./js/l2111ma1_OBJECT3_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
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
			<IFRAME NAME="MyBizASP" WIDTH="100%" src="../../blank.htm"  HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<TEXTAREA class=hidden name=txtHeader tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA class=hidden name=txtSpread tag="24" TABINDEX="-1"></TEXTAREA>

<INPUT TYPE=hidden NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHTrackingNoRule" tag="14" TABINDEX="-1">

<INPUT TYPE=hidden NAME="txtMaxRows" tag="24" TABINDEX="-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


