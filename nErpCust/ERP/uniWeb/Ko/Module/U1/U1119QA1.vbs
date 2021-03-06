'**********************************************************************************************
'*  1. Module Name          : SCM
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         : 미입고집계조회 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2004/07/27
'*  8. Modified date(Last)  : 2004/07/28
'*  9. Modifier (First)     : Park, BumSoo
'* 10. Modifier (Last)      : Park, BumSoo
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            
'**********************************************************************************************
Option Explicit				

Dim IsOpenPop                                   
Dim lgSaveRow      
Dim IsCookieSplit
Dim lgStrPrevKey1
Dim lgSortKey1
Dim lgSortKey2
Dim lgOldRow
Dim lgStrColorFlag
Dim lgBPCD

Const BIZ_PGM_ID1		= "U1119QB1.asp"
Const BIZ_PGM_ID2		= "U1119QB2.asp"

'================================================================================================================================
' Grid 1(vspdData1) - Result
Dim C_ItemCd
Dim C_ItemNm
Dim C_Spec
Dim C_Trackingno
Dim C_PlantCd
Dim C_PlantNm
Dim	C_DvryDT
Dim	C_POUnit
Dim	C_POQty
Dim	C_RcptQty
Dim	C_UnRcptQty
Dim	C_UnRcptWaitQty
Dim	C_UnRcptFirmQty
Dim C_GrpFlag

' Grid 2(vspdData2) - Result
Dim C_PONo
Dim C_POSeq
Dim C_PODt
Dim C_ItemCd2
Dim C_ItemNm2
Dim C_Spec2
Dim C_Trackingno2
Dim C_SaleQty
Dim C_IssueQty
Dim C_RemainQty
Dim	C_RemainWaitQty
Dim	C_RemainFirmQty
Dim C_SlCd
Dim C_SlNm

'================================================================================================================================
Sub InitVariables()
    lgPageNo         = ""
	lgBlnFlgChgValue = False
    lgSortKey1       = 1
    lgSortKey2       = 1
    lgSaveRow        = 0
    lgOldRow		 = 0
    lgStrPrevKey	 = ""
    lgStrPrevKey1	 = ""
    lgIntFlgMode = Parent.OPMD_CMODE 
End Sub

'================================================================================================================================
Sub SetDefaultVal()
	frm1.txtDvFrDt.Text	= StartDate
	frm1.txtDvToDt.Text	= EndDate
	'Call SetBPCD()
	Call SetToolBar("11000000000011")
End Sub

'================================================================================================================================
Sub SetBPCD()

	If 	CommonQueryRs2by2(" BP_NM ", " B_BIZ_PARTNER ", " BP_CD = " & FilterVar(parent.gUsrId, "", "S"), lgF0) = False Then
		Call ggoOper.SetReqAttr(frm1.txtPlantCd,"Q")
		Call ggoOper.SetReqAttr(frm1.txtItemCd,"Q")
		Call ggoOper.SetReqAttr(frm1.txtDvFrDt,"Q")
		Call ggoOper.SetReqAttr(frm1.txtDvToDt,"Q")
		Call DisplayMsgBox("210033","X","X","X")
		Call SetToolBar("10000000000011")
		Exit Sub
	Else
		Call SetToolBar("11000000000011")
	End If

	lgF0 = Split(lgF0, Chr(11))
	frm1.txtBpCd.value = parent.gUsrId
	frm1.txtBpNm.value = lgF0(1)

End Sub

'================================================================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)

	Call InitSpreadPosVariables(pvSpdNo)

	If pvSpdNo = "A" Or pvSpdNo = "*" Then
		'------------------------------------------
		' Grid 1 - Operation Spread Setting
		'------------------------------------------
		With frm1.vspdData1 
			
			ggoSpread.Source = frm1.vspdData1
			ggoSpread.Spreadinit "V20021224", ,Parent.gAllowDragDropSpread
					
			.ReDraw = false
					
			.MaxCols = C_GrpFlag + 1    
			.MaxRows = 0    
			
			Call GetSpreadColumnPos("A")

			ggoSpread.SSSetEdit		C_ItemCd,		"품목",			18
			ggoSpread.SSSetEdit		C_ItemNm,		"품목명",		18
			ggoSpread.SSSetEdit		C_Spec,			"규격",			15
			ggoSpread.SSSetEdit		C_Trackingno,	"Tracking No.",	15
			ggoSpread.SSSetEdit		C_PlantCd,		"공장",			8
			ggoSpread.SSSetEdit		C_PlantNm,		"공장명",		12
			ggoSpread.SSSetDate 	C_DvryDT,		"납기일",		10, 2, parent.gDateFormat		 
			ggoSpread.SSSetEdit		C_POUnit,		"입고단위",		8
			ggoSpread.SSSetFloat	C_POQty,		"발주수량",		12,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_RcptQty,		"입고수량",		12,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_UnRcptQty,	"미입고량",		12,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_UnRcptWaitQty,"입고대기수량",	12,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_UnRcptFirmQty,"미입고잔량",	12,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit 	C_GrpFlag,      ""			,1
			
			Call ggoSpread.SSSetColHidden( C_GrpFlag, C_GrpFlag, True)
			Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols, True)
			
			ggoSpread.SSSetSplit2(2)
			
			Call SetSpreadLock("A")
			
			.Col = 1 : .ColMerge = 2
			.Col = 2 : .ColMerge = 2
			
			.ReDraw = true    
    
		End With
	
    End If
    
    If pvSpdNo = "B" Or pvSpdNo = "*" Then
		'------------------------------------------
		' Grid 1 - Operation Spread Setting
		'------------------------------------------
		With frm1.vspdData2 
			
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.Spreadinit "V20021225", ,Parent.gAllowDragDropSpread
					
			.ReDraw = false
					
			.MaxCols = C_SlNm + 1    
			.MaxRows = 0    
			
			Call GetSpreadColumnPos("B")

			ggoSpread.SSSetEdit 	C_PONo,			"발주번호",		18
			ggoSpread.SSSetEdit		C_POSeq,		"행번",			6
			ggoSpread.SSSetDate 	C_PODT,			"발주일",		10, 2, parent.gDateFormat		 
			ggoSpread.SSSetEdit		C_ItemCd2,		"품목",			12
			ggoSpread.SSSetEdit		C_ItemNm2,		"품목명",		18
			ggoSpread.SSSetEdit		C_Spec2,		"규격",			15
			ggoSpread.SSSetEdit		C_TrackingNo2,	"Tracking No.",	15
			ggoSpread.SSSetFloat	C_SaleQty,		"발주수량",		15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_IssueQty,		"입고수량",		15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_RemainQty,	"미입고량",		15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_RemainWaitQty,"입고대기수량",	12,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_RemainFirmQty,"미입고잔량",	12,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_SlCd,			"입고창고"  ,	10
			ggoSpread.SSSetEdit		C_SlNm,			"입고창고명",	18
					
			Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols, True)
			
			ggoSpread.SSSetSplit2(2)
			
			Call SetSpreadLock("B")
			
			.ReDraw = true    
    
		End With
    End If
    
End Sub

'================================================================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)
	If pvSpdNo = "A" Then
		'--------------------------------
		'Grid 1
		'--------------------------------
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.SpreadLockWithOddEvenRowColor()
	End If
		
	If pvSpdNo = "B" Then 
		'--------------------------------
		'Grid 2
		'--------------------------------
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.SpreadLockWithOddEvenRowColor()
	End If	
End Sub

'================================================================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)
	If pvSpdNo = "A" Or pvSpdNo = "*" Then
		' Grid 1(vspdData1)
		C_ItemCd		= 1
		C_ItemNm		= 2
		C_Spec			= 3
		C_TRACKINGNO	= 4
		C_PlantCd		= 5
		C_PlantNm		= 6
		C_DvryDT		= 7
		C_POUnit		= 8
		C_POQty			= 9
		C_RcptQty		= 10
		C_UnRcptQty		= 11
		C_UnRcptWaitQty	= 12
		C_UnRcptFirmQty	= 13
		C_GrpFlag		= 14
	End If	
	
	If pvSpdNo = "B" Or pvSpdNo = "*" Then
		' Grid 2(vspdData2)
		C_PONo				= 1
		C_POSeq				= 2
		C_PODt				= 3
		C_ItemCd2			= 4
		C_ItemNm2			= 5
		C_Spec2				= 6
		C_TRACKINGNO2		= 7
		C_SaleQty			= 8
		C_IssueQty			= 9
		C_RemainQty			= 10
		C_RemainWaitQty		= 11
		C_RemainFirmQty		= 12
		C_SLCD				= 13
		C_SLNM				= 14
		
	End If	

End Sub

'================================================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos

 	Select Case Ucase(pvSpdNo)
 	Case "A"
 		ggoSpread.Source = frm1.vspdData1 
 		
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

		C_ItemCd		= iCurColumnPos(1)
		C_ItemNm		= iCurColumnPos(2)
		C_Spec			= iCurColumnPos(3)
		C_TRACKINGNO2	= iCurColumnPos(4)
		C_PlantCd		= iCurColumnPos(5)
		C_PlantNm		= iCurColumnPos(6)
		C_DvryDT		= iCurColumnPos(7)
		C_POUnit		= iCurColumnPos(8)
		C_POQty			= iCurColumnPos(9)
		C_RcptQty		= iCurColumnPos(10)
		C_UnRcptQty		= iCurColumnPos(11)
		C_UnRcptWaitQty	= iCurColumnPos(12)
		C_UnRcptFirmQty	= iCurColumnPos(13)
		C_GrpFlag		= iCurColumnPos(14)
		
 	Case "B"
 		ggoSpread.Source = frm1.vspdData2 
 		
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

		C_PONo			= iCurColumnPos(1)
		C_POSeq			= iCurColumnPos(2)
		C_PODt			= iCurColumnPos(3)
		C_ItemCd2		= iCurColumnPos(4)
		C_ItemNm2		= iCurColumnPos(5)
		C_Spec2			= iCurColumnPos(6)
		C_TRACKINGNO2	= iCurColumnPos(7)
		C_SaleQty		= iCurColumnPos(8)
		C_IssueQty		= iCurColumnPos(9)
		C_RemainQty		= iCurColumnPos(10)
		C_RemainWaitQty	= iCurColumnPos(11)
		C_RemainFirmQty	= iCurColumnPos(12)
		C_SLCD			= iCurColumnPos(13)
		C_SLNM			= iCurColumnPos(14)
		
 	End Select
  
End Sub

'================================================================================================================================
Function OpenBpCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtBpCd.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "업체팝업"
	arrParam(1) = " b_biz_partner "
	arrParam(2) = Trim(frm1.txtBpCd.Value)
	arrParam(3) = ""
	arrParam(4) = "BP_TYPE IN ('S','CS')"			
	arrParam(5) = "업체"			
	
    arrField(0) = "BP_CD"	
    arrField(1) = "BP_NM"	
    
    arrHeader(0) = "업체"		
    arrHeader(1) = "업체명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtBpCd.focus
		Set gActiveElement = document.activeElement	
		Exit Function
	Else
		frm1.txtBpCd.Value = arrRet(0)
		frm1.txtBpNm.Value = arrRet(1)
		frm1.txtBpCd.focus
		Set gActiveElement = document.activeElement	
	End If	
	
End Function

'================================================================================================================================
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장"
	arrParam(1) = "(			SELECT	DISTINCT B.PLANT_CD FROM M_SCM_PLAN_PUR_RCPT A, M_PUR_ORD_DTL B, M_PUR_ORD_HDR C "
	arrParam(1) = arrParam(1) & "WHERE	A.PO_NO = B.PO_NO AND A.PO_SEQ_NO = B.PO_SEQ_NO AND A.SPLIT_SEQ_NO = 0 "
	arrParam(1) = arrParam(1) & "AND	A.PO_NO = C.PO_NO AND C.BP_CD = '" & frm1.txtBpCd.value & "') A, B_PLANT B"
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = "A.PLANT_CD = B.PLANT_CD"			
	arrParam(5) = "공장"			
	
    arrField(0) = "A.PLANT_CD"	
    arrField(1) = "B.PLANT_NM"	
    
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
		frm1.txtPlantCd.Value = arrRet(0)
		frm1.txtPlantNm.Value = arrRet(1)
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement	
	End If	
	
End Function
'================================================================================================================================
Function OpenItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "품목팝업"
	arrParam(1) = "(			SELECT	DISTINCT ITEM_CD FROM M_SCM_PLAN_PUR_RCPT A, M_PUR_ORD_HDR B "
	arrParam(1) = arrParam(1) & "WHERE	A.PO_NO = B.PO_NO AND A.SPLIT_SEQ_NO = 0 AND B.BP_CD = '" & frm1.txtBpCd.value & "') A, B_ITEM B"
	arrParam(2) = Trim(frm1.txtItemCd.Value)															' Code Condition Value
	arrParam(3) = ""																					' Name Cindition Value
	arrParam(4) = "A.ITEM_CD = B.ITEM_CD "
	arrParam(5) = "품목"
	 
    arrField(0) = "A.ITEM_CD"																			' Field명(0)
    arrField(1) = "B.ITEM_NM"																			' Field명(1)
    
    arrHeader(0) = "품목"																				' Header명(0)
    arrHeader(1) = "품목명"																				' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement	
		Exit Function
	Else
		frm1.txtItemCd.value = arrRet(0)
		frm1.txtItemNm.value = arrRet(1)
		frm1.txtItemCd.focus()	
		Call SetFocusToDocument("M")
	End If	
	
End Function



'================================================================================================================================
Function OpenSlCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtBpCd.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "창고팝업"
	arrParam(1) = " B_STORAGE_LOCATION "
	arrParam(2) = Trim(frm1.txtSLCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "창고"			
	
    arrField(0) = "SL_CD"	
    arrField(1) = "SL_NM"	
    
    arrHeader(0) = "창고"		
    arrHeader(1) = "창고명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtSlCd.focus
		Set gActiveElement = document.activeElement	
		Exit Function
	Else
		frm1.txtSlCd.Value = arrRet(0)
		frm1.txtSlNm.Value = arrRet(1)
		frm1.txtSlCd.focus
		Set gActiveElement = document.activeElement	
	End If	
	
End Function


'=================================================================================================================================
Function SetItemInfo(Byval arrRet)

End Function

'--------------------------------------  OpenTrackingInfo()  ------------------------------------------
'	Name : OpenTrackingInfo()
'	Description : OpenTracking Info PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenTrackingInfo()
	
	Dim arrRet
	Dim arrParam(4)
	Dim iCalledAspName
	
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	If IsOpenPop = True Or UCase(frm1.txtTrackingNo.className) = "PROTECTED" Then Exit Function
	
	iCalledAspName = AskPRAspName("P4600PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4600PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = Trim(frm1.txtTrackingNo.value)
	arrParam(2) = Trim(frm1.txtItemCd.value)
	arrParam(3) = frm1.txtDvFrDt.Text
	arrParam(4) = frm1.txtDvToDt.Text
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetTrackingNo(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtTrackingNo.focus
	
End Function



'------------------------------------------  SetTrackingNo()  -----------------------------------------
'	Name : SetTrackingNo()
'	Description : Tracking No. Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetTrackingNo(Byval arrRet)
    frm1.txtTrackingNo.Value = arrRet(0)
End Function



'=================================================================================================================================
Function CookiePage(ByVal Kubun)

End Function

'================================================================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub
'================================================================================================================================
Sub txtDvFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtDvFrDt.Action = 7
		Call SetFocusToDocument("M") 
		frm1.txtDvFrDt.focus
	End If
End Sub
'================================================================================================================================
Sub txtDvToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtDvToDt.Action = 7
		Call SetFocusToDocument("M") 
		frm1.txtDvToDt.focus
	End If
End Sub
'================================================================================================================================
Sub txtDvFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call FncQuery()
End Sub
'================================================================================================================================
Sub txtDvToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call FncQuery()
End Sub
'================================================================================================================================
Sub vspdData1_GotFocus()
    ggoSpread.Source = frm1.vspdData1
End Sub
'================================================================================================================================
Sub vspdData2_GotFocus()
    ggoSpread.Source = frm1.vspdData2
End Sub
'================================================================================================================================
Sub vspdData1_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Or Frm1.vspdData1.MaxRows = 0 Then 
	     Exit Sub
	End If
End Sub
'================================================================================================================================
Sub vspdData2_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Or Frm1.vspdData2.MaxRows = 0 Then 
	     Exit Sub
	End If
End Sub
'================================================================================================================================
Sub vspdData1_Click(ByVal Col , ByVal Row )
	
	Call SetPopupMenuItemInf("0000111111")
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData1

    If frm1.vspdData1.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
   	If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData1
        If lgSortKey1 = 1 Then
            ggoSpread.SSSort Col
            lgSortKey1 = 2
        Else
            ggoSpread.SSSort Col, lgSortKey1
            lgSortKey1 = 1
        End If
   
    End If
    
    If lgOldRow <> Row Then
				
		frm1.vspdData2.MaxRows = 0 
		lgStrPrevKey1 = ""
		If DbDtlQuery = False Then	
			Call RestoreToolBar()
			Exit Sub
		End If
		
		lgOldRow = frm1.vspdData1.ActiveRow
			
	End If
    
End Sub
'================================================================================================================================
Sub vspdData2_Click(ByVal Col , ByVal Row )
	
	Call SetPopupMenuItemInf("0000111111")
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SP2C"
	
	Set gActiveSpdSheet = frm1.vspdData2

    If frm1.vspdData2.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
   	If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData2
        If lgSortKey2 = 1 Then
            ggoSpread.SSSort Col
            lgSortKey2 = 2
        Else
            ggoSpread.SSSort Col, lgSortKey2
            lgSortKey2 = 1
        End If
    Else
        
    End If
    
End Sub
'================================================================================================================================
Sub vspdData1_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
	
End Sub
'================================================================================================================================
Sub vspdData2_MouseDown(Button,Shift,x,y)
		
	If Button = 2 And gMouseClickStatus = "SP2C" Then
       gMouseClickStatus = "SP2CR"
    End If

End Sub
'================================================================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
    If CheckRunningBizProcess = True Then
        Exit Sub
	End If
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then
		If lgStrPrevKey <> "" Then
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If     
    End if
    
End Sub
'================================================================================================================================
Sub vspdData2_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
    If CheckRunningBizProcess = True Then
        Exit Sub
	End If
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then
		If lgStrPrevKey1 <> "" Then
			If DbDtlQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If     
    End if
    
End Sub
'================================================================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 
 
'================================================================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub
'================================================================================================================================
Sub vspdData1_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub
'================================================================================================================================
Sub vspdData2_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")
End Sub
'================================================================================================================================
Sub vspdData1_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    With frm1.vspdData1 
		If Row >= NewRow Then
		    Exit Sub
		End If
    End With
End Sub
'================================================================================================================================
Sub vspdData2_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    With frm1.vspdData2 
		If Row >= NewRow Then
		    Exit Sub
		End If
    End With
End Sub

'================================================================================================================================
' Function Name : PopSaveSpreadColumnInf
' Function Desc : 그리드 현상태를 저장한다.
'================================================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub 
 
'================================================================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : 그리드를 예전 상태로 복원한다.
'================================================================================================================================
Sub PopRestoreSpreadColumnInf()
     ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet("*")
    Call ggoSpread.ReOrderingSpreadData
End Sub 

'================================================================================================================================
Function FncQuery() 

    FncQuery = False                                        
    
    Err.Clear                                               
    
    If ValidDateCheck(frm1.txtDvFrDt, frm1.txtDvToDt) = False Then Exit Function
	
	'-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    Call InitVariables														'⊙: Initializes local global variables

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then										'⊙: This function check indispensable field
       Exit Function
    End If
    
    If DbQuery = False Then Exit Function

    FncQuery = True											
	Set gActiveElement = document.activeElement
	
End Function

'================================================================================================================================
Function FncPrint() 
    Call parent.FncPrint()
    Set gActiveElement = document.activeElement
End Function
'================================================================================================================================
Function FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
	Set gActiveElement = document.activeElement
End Function
'================================================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI , False)    
    Set gActiveElement = document.activeElement                
End Function
'================================================================================================================================
Function FncExit()
    FncExit = True
    Set gActiveElement = document.activeElement
End Function
'================================================================================================================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear           
                                        
	If CheckRunningBizProcess = True Then
       Exit Function
    End If                                              
    
    Call LayerShowHide(1)
    
    Call MakeKeyStream("M")
    
	strVal = BIZ_PGM_ID1 & "?txtMode="	& parent.UID_M0001
	strVal = strVal & "&txtKeyStream="  & lgKeyStream
	strVal = strVal & "&lgStrPrevKey="  & lgStrPrevKey
	strVal = strVal & "&txtMaxRows="	& frm1.vspddata1.MaxRows
	
    Call RunMyBizASP(MyBizASP, strVal)
    
    DbQuery = True

End Function

'================================================================================================================================
Function DbQueryOk()													

    lgBlnFlgChgValue = False
    lgSaveRow        = 1
    
    If frm1.vspdData1.MaxRows > 0 Then
    	Call SetQuerySpreadColor

		Call SetToolbar("1100000000011111")	
    
		If lgIntFlgMode <> parent.OPMD_UMODE Then
    		If DbDtlQuery = False Then	
				Call RestoreToolBar()
			End If
		End If
		frm1.vspdData1.Focus
	Else
		frm1.txtPlantCd.focus
	End If
	
    lgIntFlgMode = parent.OPMD_UMODE
	
	Set gActiveElement = document.activeElement	
End Function

'================================================================================================================================
Sub SetQuerySpreadColor()

	Dim iArrColor1, iArrColor2
	Dim iLoopCnt
	
	iArrColor1 = Split(lgStrColorFlag,Parent.gRowSep)

	For iLoopCnt=0 to ubound(iArrColor1,1) - 1
		iArrColor2 = Split(iArrColor1(iLoopCnt),Parent.gColSep)
		
		With frm1.vspdData1	
		.Col = -1
		.Row =  iArrColor2(0)
		
		Select Case iArrColor2(1)
			Case "1"
				'.BackColor = RGB(204,255,153) '연두 
			Case "2"
				.BackColor = RGB(176,234,244) '하늘색 
				.ForeColor = vbBlue
			Case "3"
				.BackColor = RGB(224,206,244) '연보라 
			Case "4"  
				.BackColor = RGB(251,226,153) '연주황 
			Case "5" 
				.BackColor = RGB(255,255,153) '연노랑 
				.ForeColor = vbBlue
		End Select
		End With
	Next



End Sub

'================================================================================================================================
Function DbDtlQuery() 
    
    Dim strVal
	
    DbDtlQuery = False

	Call LayerShowHide(1)
    
    Call MakeKeyStream("S")
    
	strVal = BIZ_PGM_ID2 & "?txtMode="	& parent.UID_M0001
	strVal = strVal & "&txtKeyStream="  & lgKeyStream
	strVal = strVal & "&lgStrPrevKey="  & lgStrPrevKey1
			   
    Call RunMyBizASP(MyBizASP, strVal)
    
    DbDtlQuery = True
    
End Function

'========================================================================================
Function DbDtlQueryOk()

End Function                          

'========================================================================================
Sub MakeKeyStream(pOpt)

	Dim strPlantcd
	Dim strItemCd
	Dim dtDvryDt

   Select Case pOpt
       Case "M"
           
				With frm1
					If lgIntFlgMode = parent.OPMD_UMODE Then
						lgKeyStream = UCase(Trim(.hPlantCd.value))  & Parent.gColSep
						lgKeyStream = lgKeyStream & UCase(Trim(.hItemCd.value))  & Parent.gColSep
						lgKeyStream = lgKeyStream & UCase(Trim(.hBPCd.value))  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(.hDvFrDt.value)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(.hDvToDt.value)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(.hSLCD.value)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(.hTRACKINGNO.value)  & Parent.gColSep
						
					Else
						lgKeyStream = UCase(Trim(.txtPlantCd.value))  & Parent.gColSep
						lgKeyStream = lgKeyStream & UCase(Trim(.txtItemCd.value))  & Parent.gColSep
						lgKeyStream = lgKeyStream & UCase(Trim(.txtBPCd.value))  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(.txtDvFrDt.Text)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(.txtDvToDt.Text)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(.txtSLCD.value)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(.txtTRACKINGNO.value)  & Parent.gColSep
						
						.hPlantCd.value		= .txtPlantCd.value
						.hItemCd.value		= .txtItemCd.value
						.hBPCd.value		= .txtBPCd.value
						.hDvFrDt.value		= .txtDvFrDt.Text
						.hDvToDt.value		= .txtDvToDt.Text
						.hSlCd.value		= .txtSlCd.value
						.hTRACKINGNO.value	= .txtTRACKINGNO.value
					End If
				End With
			
       Case "S"
				With frm1
					.vspdData1.Row = .vspdData1.ActiveRow
					.vspdData1.Col = C_PlantCd
					strPlantcd = .vspdData1.text
					.vspdData1.Col = C_ItemCd
					strItemCd = .vspdData1.text
					.vspdData1.Col = C_DvryDt
					dtDvryDt = .vspdData1.text
					
					lgKeyStream = UCase(Trim(strPlantcd))  & Parent.gColSep
					lgKeyStream = lgKeyStream & UCase(Trim(strItemCd))    & Parent.gColSep
					lgKeyStream = lgKeyStream & UCase(Trim(.hBPCd.value)) & Parent.gColSep
					lgKeyStream = lgKeyStream & UCase(Trim(dtDvryDt))     & Parent.gColSep
					lgKeyStream = lgKeyStream & Trim(.hDvFrDt.value)      & Parent.gColSep
					lgKeyStream = lgKeyStream & Trim(.hDvToDt.value)      & Parent.gColSep
					lgKeyStream = lgKeyStream & Trim(.hSLCD.value)		  & Parent.gColSep
					lgKeyStream = lgKeyStream & Trim(.hTRACKINGNO.value)  & Parent.gColSep
					
				End With

	End Select
   
End Sub                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              