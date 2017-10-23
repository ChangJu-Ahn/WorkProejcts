'**********************************************************************************************
'*  1. Module Name          : PROCUREMENT
'*  2. Function Name        : 
'*  3. Program ID           : u2123ma1.vbs
'*  4. Program Name         :
'*  5. Program Desc         : 납품예정확정등록 (Manage Planned Delivery Date)
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2004/07/27
'*  8. Modified date(Last)  : 2004/07/28
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : LEE SEUNG WOOK
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            
'**********************************************************************************************

Const BIZ_PGM_ID			= "u2123mb1.asp"
Const BIZ_PGM_ID2			= "u2123mb2.asp"

'Spreadsheet1
Dim C_OrderDt
Dim C_ItemCode
Dim C_ItemName
Dim C_Spec
Dim C_PlantCd
Dim C_PlantNm
Dim C_BpCd
Dim C_BpNm
Dim C_OrderUnit
Dim C_OrderNo
Dim C_OrderSeq
Dim C_OrderQty
Dim C_DvryDt
Dim	C_RcptQty
Dim C_InspectQty
Dim	C_UnRcptQty
Dim	C_FirmDvryQty
Dim	C_RemainQty
Dim C_DvryQty

'Spreadsheet2
Dim C_Title
Dim C_DlvyPlanDt
Dim C_DlvyPlanQty
Dim C_confirmYN
Dim C_ConfirmQty
Dim C_MType
Dim C_InspQty
Dim C_RcptQty2
Dim C_SplitSeqNo

Dim IsOpenPop

'================================================================================================================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE
    lgIntGrpCount = 0
    lgStrPrevKey = ""
    lgStrPrevKey1 = ""
    lgLngCurRows = 0
    lgSortKey = 1
    lgSortKey2 = 1
End Sub

'================================================================================================================================
Sub InitData(ByVal lngStartRow)

	Dim intRow
    Dim intIndex
    Dim strInspQty,strRcptQty
    
	With frm1.vspdData2
		.ReDraw = False
		For intRow = 1 To .MaxRows
			
			.Row = intRow
			.Col = C_InspQty
			strInspQty = CDbl(.Text)
			
			.Col = C_RcptQty2
			strRcptQty = CDbl(.Text)
			
			.col = C_ConfirmYN
			.Row = intRow
			If .value = "1" or .value = "Y" then
				If strInspQty = 0 Or strRcptQty = 0 Then
					ggoSpread.SpreadUnLock		C_ConfirmQty		, intRow	,intRow
					ggoSpread.sssetrequired     C_ConfirmQty		, intRow	,intRow
				Else
					ggoSpread.SSSetProtected	C_ConfirmYN			,intRow		,intRow
					ggoSpread.SSSetProtected	C_ConfirmQty		,intRow		,intRow
				End If
			Else
				ggoSpread.SSSetProtected	C_ConfirmQty		,intRow		,intRow
			End If
		
		Next
		.ReDraw = True	
	End With

End Sub

'================================================================================================================================
Sub SetDefaultVal()
	frm1.txtDvFrDt.text = UniConvDateAToB(UNIDateAdd ("D", -7, LocSvrDate, parent.gServerDateFormat), parent.gServerDateFormat, parent.gDateFormat)
	frm1.txtDvToDt.text = UniConvDateAToB(UNIDateAdd ("D", 7, LocSvrDate, parent.gServerDateFormat), parent.gServerDateFormat, parent.gDateFormat)
	Call SetBPCD()
End Sub

'================================================================================================================================
Sub SetBPCD()

	If 	CommonQueryRs2by2(" BP_NM ", " B_BIZ_PARTNER ", " BP_CD = " & FilterVar(parent.gUsrId, "", "S"), lgF0) = False Then
		Call ggoOper.SetReqAttr(frm1.txtItemCd,"Q")
		Call ggoOper.SetReqAttr(frm1.txtDvFrDt,"Q")
		Call ggoOper.SetReqAttr(frm1.txtDvToDt,"Q")
		Call ggoOper.SetReqAttr(frm1.txtdlvyno,"Q")
		Call DisplayMsgBox("210033","X","X","X")
		Call SetToolBar("10000000000011")
		Exit Sub
	Else
		Call SetToolBar("11000001000011")								'⊙: 버튼 툴바 제어 
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
		With frm1
			ggoSpread.Source = .vspdData1
			ggoSpread.Spreadinit "V20060713", , Parent.gAllowDragDropSpread
			.vspdData1.ReDraw = False
	
			.vspdData1.MaxCols = C_DvryQty + 1
			.vspdData1.MaxRows = 0
			
			Call GetSpreadColumnPos("A")

			ggoSpread.SSSetDate 	C_OrderDt,		"수주일자", 10, 2, parent.gDateFormat		 
			ggoSpread.SSSetEdit		C_ItemCode,		"품목"    , 18,,,18,2
			ggoSpread.SSSetEdit		C_ItemName,		"품목명"  , 18
			ggoSpread.SSSetEdit		C_Spec,			"규격"    , 15
			ggoSpread.SSSetEdit		C_PlantCd,		"납품처"  ,  8
			ggoSpread.SSSetEdit		C_PlantNm,		"납품처명",	12
			ggoSpread.SSSetEdit		C_BpCd,			"발주처"  ,  8
			ggoSpread.SSSetEdit		C_BpNm,			"발주처명",	12
			ggoSpread.SSSetEdit		C_OrderUnit,	"단위"    ,  4,,,3,2
			ggoSpread.SSSetEdit		C_OrderNo,		"수주번호", 15
			ggoSpread.SSSetEdit		C_OrderSeq,		"행번"    ,  4
			ggoSpread.SSSetFloat	C_OrderQty,		"수주량"  , 10,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetDate 	C_DvryDt,		"납기일",   10, 2, parent.gDateFormat
			ggoSpread.SSSetFloat	C_RcptQty,		"납품수량",		12,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_InspectQty,	"검사중수량", 12,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_UnRcptQty,	"미납수량",		12,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_FirmDvryQty,	"납품대기량",	10,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_RemainQty,	"납품잔량",		10,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			'ggoSpread.SSSetDate 	C_DvryPlanDt,	"납품예정일자", 10, 2, parent.gDateFormat
			ggoSpread.SSSetFloat	C_DvryQty,		"납품예정수량",	10,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			
			Call ggoSpread.SSSetColHidden( .vspdData1.MaxCols, .vspdData1.MaxCols , True)
			
			.vspdData1.ReDraw = True
			'ggoSpread.SSSetSplit2(3)
			
			Call SetSpreadLock("A")
		End With
	End If
	
	If pvSpdNo = "B" Or pvSpdNo = "*" Then
	
	'------------------------------------------
	' Grid 2 - Operation Spread Setting
	'------------------------------------------
		With frm1
			ggoSpread.Source = .vspdData2
			ggoSpread.Spreadinit "V20060715", , Parent.gAllowDragDropSpread
			.vspdData2.ReDraw = False
			
			.vspdData2.MaxCols = C_SplitSeqNo + 1
			.vspdData2.MaxRows = 0
			
			Call GetSpreadColumnPos("B")
			
			ggoSpread.SSSetEdit		C_Title,		"Title"       , 10,2,,18,2
			ggoSpread.SSSetDate 	C_DlvyPlanDt,	"납품예정일자", 12, 2, parent.gDateFormat
			ggoSpread.SSSetFloat	C_DlvyPlanQty,	"납품예정수량",	10,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		    ggoSpread.SSSetCheck    C_confirmYN,    "확정여부",		10     ,2				  ,""   , True
		    ggoSpread.SSSetFloat	C_ConfirmQty,	"납품확정수량",	10,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		    ggoSpread.SSSetEdit		C_MType,		"직납여부",		10     ,2
		    ggoSpread.SSSetFloat	C_InspQty,		"검사중수량",	10,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		    ggoSpread.SSSetFloat	C_RcptQty2,		"납품수량",		10,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		    ggoSpread.SSSetEdit		C_SplitSeqNo,	"", 4
		    
			Call ggoSpread.SSSetColHidden( C_SplitSeqNo, C_SplitSeqNo, True)
			Call ggoSpread.SSSetColHidden( C_MType, C_MType, True)
			Call ggoSpread.SSSetColHidden( C_InspQty, C_InspQty, True)
			Call ggoSpread.SSSetColHidden( C_RcptQty2, C_RcptQty2, True)
			Call ggoSpread.SSSetColHidden( .vspdData2.MaxCols, .vspdData2.MaxCols , True)
			
			'ggoSpread.SSSetSplit2(2)
			
			Call SetSpreadLock("B")
			
			.vspdData2.ReDraw = True
   
		End With
	End If
    
End Sub

'================================================================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)
	If pvSpdNo = "A" Then
		'--------------------------------
		'Grid 1
		'--------------------------------
		With frm1
			ggoSpread.Source = .vspdData1
			
			.vspdData1.ReDraw = False
			
			'ggoSpread.SpreadLock	 C_orderdt, -1, C_DvryQty
			ggoSpread.SpreadLockWithOddEvenRowColor()
			
			.vspdData1.ReDraw = True
		End With
	End If
	
	If pvSpdNo = "B" Then
		'--------------------------------
		'Grid 2
		'--------------------------------
		With frm1
			ggoSpread.Source = .vspdData2	
			.vspdData2.ReDraw = False
			
			ggoSpread.SpreadLock	 C_Title, -1, C_DlvyPlanQty
	
			.vspdData2.ReDraw = True
		End With
	End If

End Sub

'================================================================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)
	If pvSpdNo = "A" Or pvSpdNo = "*" Then
		' Grid 1(vspdData1)
		C_OrderDt		= 1
		C_ItemCode		= 2
		C_ItemName		= 3
		C_Spec			= 4
		C_PlantCd		= 5
		C_PlantNm		= 6
		C_BpCd			= 7
		C_BpNm			= 8
		C_OrderUnit		= 9
		C_OrderNo		= 10
		C_OrderSeq		= 11
		C_OrderQty		= 12
		C_DvryDt		= 13
		C_RcptQty		= 14
		C_InspectQty	= 15
		C_UnRcptQty		= 16
		C_FirmDvryQty	= 17
		C_RemainQty		= 18
		C_DvryQty		= 19
	End If
	
	If pvSpdNo = "B" Or pvSpdNo = "*" Then
		' Grid 2(vspdData2)
		
		C_Title			= 1
		C_DlvyPlanDt	= 2
		C_DlvyPlanQty	= 3
		C_confirmYN		= 4
		C_ConfirmQty	= 5
		C_MType			= 6
		C_InspQty		= 7
		C_RcptQty2		= 8
		C_SplitSeqNo	= 9
	End If

End Sub
 
'================================================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case Ucase(pvSpdNo)
 		Case "A"
 			ggoSpread.Source = frm1.vspdData1 
 			
 			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_OrderDt		= iCurColumnPos(1)
			C_ItemCode		= iCurColumnPos(2)
			C_ItemName		= iCurColumnPos(3)
			C_Spec			= iCurColumnPos(4)
			C_PlantCd		= iCurColumnPos(5)
			C_PlantNm		= iCurColumnPos(6)
			C_BpCd			= iCurColumnPos(7)
			C_BpNm			= iCurColumnPos(8)
			C_OrderUnit		= iCurColumnPos(9)
			C_OrderNo		= iCurColumnPos(10)
			C_OrderSeq		= iCurColumnPos(11)
			C_OrderQty		= iCurColumnPos(12)
			C_DvryDt		= iCurColumnPos(13)
			C_RcptQty		= iCurColumnPos(14)
			C_InspectQty	= iCurColumnPos(15)
			C_UnRcptQty		= iCurColumnPos(16)
			C_FirmDvryQty	= iCurColumnPos(17)
			C_RemainQty		= iCurColumnPos(18)
			C_DvryQty		= iCurColumnPos(19)
	
		Case "B"
			ggoSpread.Source = frm1.vspdData2 
 			
 			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_Title			= iCurColumnPos(1)
			C_DlvyPlanDt	= iCurColumnPos(2)
			C_DlvyPlanQty	= iCurColumnPos(3)
			C_confirmYN		= iCurColumnPos(4)
			C_ConfirmQty	= iCurColumnPos(5)
			C_MType			= iCurColumnPos(6)
			C_InspQty		= iCurColumnPos(7)
			C_RcptQty2		= iCurColumnPos(8)
			C_SplitSeqNo	= iCurColumnPos(9)
			
 	End Select
 
End Sub


'================================================================================================================================
Function OpenItemInfo(Byval strCode)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtItemCd.className) = "PROTECTED" Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "품목팝업"
	
	arrParam(1) = "(			SELECT distinct A.ITEM_CD  FROM M_PUR_ORD_DTL A , B_STORAGE_LOCATION B , M_SCM_FIRM_PUR_RCPT C "
	arrParam(1) = arrParam(1) & " WHERE C.D_BP_CD = B.SL_CD AND A.PO_NO = C.PO_NO AND A.PO_SEQ_NO = C.PO_SEQ_NO AND C.RCPT_QTY = 0  AND B.BP_CD = '" & frm1.txtBpCd.value & "') A, B_ITEM B"
	arrParam(2) = Trim(frm1.txtItemCd.Value)								' Code Condition
	arrParam(3) = ""														' Name Cindition
	arrParam(4) = "A.ITEM_CD = B.ITEM_CD "
	arrParam(5) = "품목"
	 
    arrField(0) = "A.ITEM_CD"												' Field명(0)
    arrField(1) = "B.ITEM_NM"												' Field명(1)
    
    arrHeader(0) = "품목"													' Header명(0)
    arrHeader(1) = "품목명"													' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetItemInfo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus

End Function

'------------------------------------------  OpenDlvyNo()  -------------------------------------------------
'	Name : OpenDlvyNo()
'	Description : Purchase_Order PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenDlvyNo()
	
	Dim strRet
	Dim arrParam(2)
	Dim iCalledAspName
	Dim IntRetCD
			
	If IsOpenPop = True Or UCase(frm1.txtDlvyNo.className) = Ucase(Parent.UCN_PROTECTED) Then Exit Function
		
	IsOpenPop = True
		
	arrParam(0) = parent.gUsrId

	iCalledAspName = AskPRAspName("U2124PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "U2124PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If strRet(0) = "" Then
		Exit Function
	Else
		frm1.txtDlvyNo.value = strRet(0)
		frm1.txtDlvyNo.focus	
		Set gActiveElement = document.activeElement
	End If	
		
End Function


'================================================================================================================================
Function SetItemInfo(Byval arrRet)
    With frm1
		.txtItemCd.value = arrRet(0)
		.txtItemNm.value = arrRet(1)
    End With
End Function


'--------------------------------------  OpenTrackingInfo()  ------------------------------------------
'	Name : OpenTrackingInfo()
'	Description : OpenTracking Info PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenTrackingInfo()
	
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtTrackingNo.className) = "PROTECTED" Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "Tracking No."
	arrParam(1) = "S_SO_TRACKING "
	arrParam(2) = Trim(frm1.txtTrackingNo.Value)								' Code Condition
	arrParam(3) = ""														' Name Cindition
	arrParam(4) = ""
	arrParam(5) = "Tracking No."
	 
    arrField(0) = "TRACKING_NO"												' Field명(0)
    
    arrHeader(0) = "Tracking No"													' Header명(0)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		frm1.txtTrackingNo.value = arrRet(0)
'		Call SetTrackingNo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtTrackingNo.focus
	
End Function


'================================================================================================================================
Sub vspdData1_Change(ByVal Col , ByVal Row)
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.UpdateRow Row
End Sub

'================================================================================================================================
Sub vspdData2_Change(ByVal Col , ByVal Row)
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.UpdateRow Row
End Sub

'================================================================================================================================
Sub vspdData1_Click(ByVal Col , ByVal Row )

	If lgIntFlgMode = Parent.OPMD_UMODE Then
  		Call SetPopupMenuItemInf("1101111111")
  	Else
  		Call SetPopupMenuItemInf("1001111111")
  	End If

	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData1

	If frm1.vspdData1.MaxRows = 0 Then
		Exit Sub
	End If
   	
   	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData1
	        If lgSortKey = 1 Then
        	    ggoSpread.SSSort Col
	            lgSortKey = 2
	        Else
        	    ggoSpread.SSSort Col, lgSortKey
	            lgSortKey = 1
        	End If
	End If
	

'''' jsa 2006-9-11 없앰	
'''' 	If lgOldRow <> Row Then
'''' 				
'''' 		frm1.vspdData2.MaxRows = 0 
'''' 		lgStrPrevKey1 = ""
'''' 		
'''' 		If DbDtlQuery = False Then	
'''' 			Call RestoreToolBar()
'''' 			Exit Sub
'''' 		End If
'''' 		
'''' 		lgOldRow = frm1.vspdData2.ActiveRow
'''' 			
'''' 	End If

End Sub

'================================================================================================================================
Sub vspdData2_Click(ByVal Col , ByVal Row )

	If lgIntFlgMode = Parent.OPMD_UMODE Then
  		Call SetPopupMenuItemInf("1101111111")
  	Else
  		Call SetPopupMenuItemInf("1001111111")
  	End If

	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData2

	If frm1.vspdData2.MaxRows = 0 Then
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
		
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub

'================================================================================================================================
Sub vspdData1_DblClick(ByVal Col, ByVal Row)
   
 	If Row <= 0 Then
		Exit Sub
 	End If
    
  	If frm1.vspdData1.MaxRows = 0 Then
		Exit Sub
 	End If

End Sub

'================================================================================================================================
Sub vspdData2_DblClick(ByVal Col, ByVal Row)
   
 	If Row <= 0 Then
		Exit Sub
 	End If
    
  	If frm1.vspdData2.MaxRows = 0 Then
		Exit Sub
 	End If

End Sub
 
'================================================================================================================================
Sub vspddata1_KeyPress(index , KeyAscii )
    On Error Resume Next
End Sub

'================================================================================================================================
Sub vspddata2_KeyPress(index , KeyAscii )
    On Error Resume Next
End Sub


'================================================================================================================================
'   Event Name : vspdData1_ScriptLeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'================================================================================================================================
Sub vspdData1_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If NewRow <= 0 Or Row = NewRow Then
		Exit Sub
	End If
	
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData
	
	lgStrPrevKey1 = ""

	Call SetActiveCell(frm1.vspdData1,NewCol,NewRow,"M","X","X")

	If DbDtlQuery() = False Then	
		Exit Sub
	End If
	
End Sub

'================================================================================================================================
'   Event Name : vspdData2_LeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'================================================================================================================================
Sub vspdData2_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    On Error Resume Next
End Sub

'================================================================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	
	If CheckRunningBizProcess = True Then Exit Sub
	
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
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	
	If CheckRunningBizProcess = True Then Exit Sub
	
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
Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	With frm1.vspdData2
		 ggoSpread.Source = frm1.vspdData2
		 .Row = Row
         .Col = Col
		If Row > 0 Then
			Select Case Col

				Case C_ConfirmYN
				    
				    .Row = .activeRow
				    If .value = "1" or .value = "Y" Then
						ggoSpread.SpreadUnLock		C_ConfirmQty		, .activeRow,	C_ConfirmQty,		.activeRow
						ggoSpread.sssetrequired     C_ConfirmQty		, .activeRow,	.activeRow

''''  JSA 2006-9-11 없앰						
''''						.Col = 0
''''						.Text = ggoSpread.UpdateFlag
					Else
						ggoSpread.SSSetProtected	C_ConfirmQty		,.activeRow,	.activeRow
						
''''  JSA 2006-9-11 없앰						
''''						.Col = 0
''''						.Text = ggoSpread.UpdateFlag
						
					End If
							
			End Select
		End If
    
	End With	
    
End Sub

'================================================================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'================================================================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub  
 
'================================================================================================================================
Sub vspdData1_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub

'================================================================================================================================
Sub vspdData2_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("B")
End Sub 
 
'================================================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub 
 
'================================================================================================================================
Sub PopRestoreSpreadColumnInf()
	Dim LngRow

    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet("*")
	Call ggoSpread.ReOrderingSpreadData()

End Sub 

'================================================================================================================================
Sub txtDvFrDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDvFrDt.Action = 7
        SetFocusToDocument("M")
		Frm1.txtDvFrDt.Focus
    End If
End Sub

'================================================================================================================================
Sub txtDvToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDvToDt.Action = 7
        SetFocusToDocument("M")
		Frm1.txtDvToDt.Focus
    End If
End Sub

'================================================================================================================================
Sub txtDvFrDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'================================================================================================================================
Sub txtDvToDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
Function FncQuery()
 
    Dim IntRetCD 
    
    FncQuery = False
    
    Err.Clear

    ggoSpread.Source = frm1.vspdData1
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    If ValidDateCheck(frm1.txtDvFrDt, frm1.txtDvToDt) = False Then Exit Function
	
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.ClearSpreadData

    Call InitVariables

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function
       
    FncQuery = True
   
End Function

'================================================================================================================================
Function FncNew() 
	On Error Resume Next	
End Function

'================================================================================================================================
Function FncDelete() 
	On Error Resume Next   
End Function

'================================================================================================================================
Function FncSave()
    Dim IntRetCD 
    Dim intVar ,intVar2 ,intVar3
         
    FncSave = False 
    Err.Clear
    
    ggoSpread.Source = frm1.vspdData2
    If  ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData2
    If Not ggoSpread.SSDefaultCheck Then
       Exit Function
    End If
    
    If Not chkField(Document, "2") Then
       Exit Function
    End If
    
	With Frm1
    
		For lRow = 1 To .vspdData2.MaxRows
    
			.vspdData2.Row = lRow
			.vspdData2.Col = C_ConfirmQty
			intVar = Cdbl(.vspdData2.Text)
			
			If Cdbl(.vspdData2.Text) = 0 Then 
				Call DisplayMsgBox("169918","x","확정수량","x")
				Exit Function
			End If
	   
'			.vspdData1.Col = C_OrderQty
'			intVar2 = Cdbl(.vspdData1.Text)
			
			.vspdData2.Row = lRow
			.vspdData2.Col = C_DlvyPlanQty
			intVar3 = Cdbl(.vspdData2.Text)
			If intVar > intVar3 Then 
				Call DisplayMsgBox("127928","x","수주량","x")
				Exit Function
			End If
			
		Next	
	End With
	
	Call DisableToolBar( parent.TBC_SAVE)
	If DbSave = False Then
		Call  RestoreToolBar()
		Exit Function
	End If
    
    FncSave = True
    
End Function

'================================================================================================================================
Function FncCopy() 

End Function

'================================================================================================================================
Function FncPaste() 
     ggoSpread.SpreadPaste
End Function

'================================================================================================================================
Function FncCancel() 

	If frm1.vspdData2.MaxRows < 1 Then Exit Function	
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.EditUndo
	Call initData(frm1.vspdData2.ActiveRow)
End Function

'================================================================================================================================
Function FncInsertRow(ByVal pvRowCnt) 

End Function

'================================================================================================================================
Function FncDeleteRow() 

    Dim lDelRows
    
    If frm1.vspdData1.MaxRows < 1 Then Exit Function

    lDelRows = ggoSpread.DeleteRow
    lgLngCurRows = lDelRows + lgLngCurRows
    
End Function

'================================================================================================================================
Function FncPrint() 
	Call parent.FncPrint()
End Function

'================================================================================================================================
Function FncPrev() 
    On Error Resume Next
End Function

'================================================================================================================================
Function FncNext() 
    On Error Resume Next
End Function

'================================================================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)												
End Function

'================================================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                         
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
Function FncExit()

    Dim IntRetCD
    FncExit = False
    
    ggoSpread.Source = frm1.vspdData2							'⊙: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then						'⊙: Check If data is chaged
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "x", "x")	'⊙: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    FncExit = True
    
End Function

'================================================================================================================================
Function FncScreenSave() 
    Call ggoSpread.SaveLayout
End Function

'================================================================================================================================
Function FncScreenRestore() 
    If ggoSpread.AllClear = True Then
       ggoSpread.LoadLayout
    End If
End Function

'================================================================================================================================
Sub MakeKeyStream(pOpt)
	Dim strPoNo
	Dim strPoSeqNo
	
	Select Case pOpt
		Case "M"

			With frm1
				If lgIntFlgMode = parent.OPMD_UMODE Then
					lgKeyStream =               UCase(Trim(.hItemCd.value))  & Parent.gColSep
					lgKeyStream = lgKeyStream & UCase(Trim(.hBPCd.value))  & Parent.gColSep
					lgKeyStream = lgKeyStream & Trim(.hDvFrDt.text)   & Parent.gColSep
					lgKeyStream = lgKeyStream & Trim(.hDvToDt.text)   & Parent.gColSep
					lgKeyStream = lgKeyStream & Trim(.txtDLVYNO.value) & Parent.gColSep
					lgKeyStream = lgKeyStream & Trim(.htrackingno.value) & Parent.gColSep
				Else
					lgKeyStream =               UCase(Trim(.txtItemCd.value))  & Parent.gColSep
					lgKeyStream = lgKeyStream & UCase(Trim(.txtBPCd.value))  & Parent.gColSep
					lgKeyStream = lgKeyStream & Trim(.txtDvFrDt.text)  & Parent.gColSep
					lgKeyStream = lgKeyStream & Trim(.txtDvToDt.text)  & Parent.gColSep
					lgKeyStream = lgKeyStream & Trim(.txtDLVYNO.value) & Parent.gColSep
					lgKeyStream = lgKeyStream & Trim(.txttrackingno.value) & Parent.gColSep
					
					.hItemCd.value		= .txtItemCd.value
					.hBPCd.value		= .txtBPCd.value
					.hDvFrDt.value		= .txtDvFrDt.Text
					.hDvToDt.value		= .txtDvToDt.Text
					.hTrackingno.value	= .txttrackingno.value
					
				End If
			End With
		Case "S"
			With frm1
				.vspdData1.Row = .vspdData1.ActiveRow
				.vspdData1.Col = C_OrderNo
				
				strPoNo = .vspdData1.value
				
				.vspdData1.Col = C_OrderSeq
				
				strPoSeqNo = .vspdData1.value
				
				lgKeyStream = lgKeyStream & UCase(Trim(strPoNo))  & Parent.gColSep
				lgKeyStream = lgKeyStream & UCase(Trim(strPoSeqNo))  & Parent.gColSep
			
			End With
	End Select
End Sub

'================================================================================================================================
Function DbQuery() 
    
    Dim strVal
    
    Err.Clear

    DbQuery = False
    
    Call LayerShowHide(1)
 
    Call MakeKeyStream("M")
    
	strVal = BIZ_PGM_ID & "?txtMode="	& parent.UID_M0001
	strVal = strVal & "&txtKeyStream="  & lgKeyStream
	strVal = strVal & "&lgStrPrevKey="  & lgStrPrevKey
	
	
    Call RunMyBizASP(MyBizASP, strVal)

    DbQuery = True
    
End Function

'================================================================================================================================
Function DbQueryOk(ByVal LngMaxRow)

 	Dim lRow
 	Dim LngRow    
	
    Call ggoOper.LockField(Document, "Q")
    Call SetToolBar("11001001000111")
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then
		If DbDtlQuery = False Then	
			Call RestoreToolBar()
			Exit Function
		End If
	End If

    '-----------------------
    'Reset variables area
    '-----------------------
    Frm1.vspddata1.Focus
    lgIntFlgMode = parent.OPMD_UMODE
   
End Function

'================================================================================================================================
Function DbQueryNotOk()	

	Call SetToolBar("11001101001111")
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_CMODE

End Function

'================================================================================================================================
Function DbDtlQuery() 
    Dim strVal
	
    DbDtlQuery = False

	Call LayerShowHide(1)

	lgKeyStream = ""
    Call MakeKeyStream("S")
    
	strVal = BIZ_PGM_ID2 & "?txtMode="	& parent.UID_M0001
	strVal = strVal & "&txtKeyStream="  & lgKeyStream
	strVal = strVal & "&lgStrPrevKey="  & lgStrPrevKey1

    Call RunMyBizASP(MyBizASP, strVal)													'☜: 비지니스 ASP 를 가동 
    
    DbDtlQuery = True
    
End Function

'================================================================================================================================
Function DbDtlQueryOk()
	Call InitData(1)
	Call SetQuerySpreadColor
End Function

Sub SetQuerySpreadColor()
	
	Dim iArrColor1, iArrColor2
	Dim iLoopCnt
	
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.SSSetProtected C_ConfirmYN, frm1.vspdData2.MaxRows, frm1.vspdData2.MaxRows
	ggoSpread.SSSetProtected C_ConfirmQty, frm1.vspdData2.MaxRows, frm1.vspdData2.MaxRows
	
	iArrColor1 = Split(lgStrColorFlag,Parent.gRowSep)

	For iLoopCnt=0 to ubound(iArrColor1,1) - 1
		iArrColor2 = Split(iArrColor1(iLoopCnt),Parent.gColSep)
		
		With frm1.vspdData2	
			.Col = -1
			.Row =  iArrColor2(0)
		
			Select Case iArrColor2(1)
				Case "1"
					.BackColor = RGB(176,234,244) '하늘색 
					.ForeColor = vbBlue
			End Select
		End With
	Next

End Sub

'================================================================================================================================
Function DbSave() 

    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
	Dim strVal, strDel
	Dim strDlvyQty, strConfirmQty, strRemainQty
	Dim strPoNo,strPoSeqNo,strSplitNo
	Dim strOverTolQty, strOverTol
	
    DbSave = False                                                          
    
	If LayerShowHide(1) = false then
		Exit Function
	End if

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	With Frm1
    
       For lRow = 1 To .vspdData2.MaxRows
           .vspdData2.Row = lRow
           .vspdData2.Col = 0
        
           Select Case .vspdData2.Text

               Case  ggoSpread.UpdateFlag
					
					.vspdData1.Col = C_OrderNo
					strPoNo = Trim(.vspdData1.Value)
					
					.vspdData1.Col = C_OrderSeq
					strPoSeqNo = Trim(.vspdData1.Text)
					
					.vspdData2.Col = C_SplitSeqNo
					strSplitNo = Trim(.vspdData2.Text)
					
					Call CommonQueryRs(" ISNULL(SUM(CONFIRM_QTY),0) ", " M_SCM_FIRM_PUR_RCPT ", " CONFIRM_YN = " & FilterVar("Y", "''", "S") _
						 & " AND PO_NO = " & FilterVar(strPoNo, "''", "S") & " AND PO_SEQ_NO = " & FilterVar(strPoSeqNo, "''", "S") _
						 & " AND SPLIT_SEQ_NO <> " & FilterVar(strSplitNo, "''", "S"), lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
						 
					strDlvyQty = Split(lgF0, Chr(11))(0)

					.vspdData2.Col = C_ConfirmQTY
					strConfirmQty = Trim(.vspdData2.Text)
					
					.vspdData1.Col = C_OrderQty
					strOrderQty = Trim(.vspdData1.Text)
					
					Call CommonQueryRs(" OVER_TOL ", " M_PUR_ORD_DTL ", "PO_NO = " & FilterVar(strPoNo, "''", "S") & _
					        " AND PO_SEQ_NO = " & FilterVar(strPoSeqNo, "''", "S") , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
						 
					strOverTol = Split(lgF0, Chr(11))(0)
					strOverTolQty = Unicdbl(strOrderQty) * (Unicdbl(strOverTol) / 100)

					strRemainQty = Unicdbl(strOrderQty) + Unicdbl(strOverTolQty) - Unicdbl(strDlvyQty)
					
					If Unicdbl(strRemainQty) < Unicdbl(strConfirmQty) Then
						Call DisplayMsgBox("U20001","X","X","X")
						Call LayerShowHide(0)
						Exit Function	
					End If
               
													  strVal = strVal & "U"  &  parent.gColSep					
													  strVal = strVal & lRow &  parent.gColSep
					.vspdData1.Col = C_OrderNo	    : strVal = strVal & Trim(.vspdData1.Text)  & parent.gColSep
					.vspdData1.Col = C_OrderSeq     : strVal = strVal & Trim(.vspdData1.Value) & parent.gColSep
					.vspdData2.Col = C_SplitSeqNo   : strVal = strVal & Trim(.vspdData2.Value) & parent.gColSep

					.vspdData2.Col = C_ConfirmYN
					if .vspdData2.Value = "Y" or .vspdData2.Value = "1" then
													: strVal = strVal & "Y" & parent.gColSep	'5
					Else
													: strVal = strVal & "N" & parent.gColSep	'5
					End if
					
					.vspdData2.Col = C_ConfirmQTY	: strVal = strVal & strConfirmQty & parent.gColSep
					
					.vspdData2.Col = C_MType
					if .vspdData2.Value = "Y" or .vspdData2.Value = "1" then
													: strVal = strVal & "Y" & parent.gRowSep	'7
					Else
													: strVal = strVal & "N" & parent.gRowSep	'7
					End if
					
					lGrpCnt = lGrpCnt + 1

           End Select
       Next
	
       .txtMode.value        =  parent.UID_M0002
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strDel & strVal

	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID2)	
    DbSave = True
    
End Function

'================================================================================================================================
Function DbSaveOk()

	Call InitVariables
	ggoSpread.source = frm1.vspddata1
    frm1.vspdData1.MaxRows = 0
    Call RemovedivTextArea
    Call MainQuery()

End Function

'================================================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr, parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData2.MaxCols - 1
           Frm1.vspdData2.Col = iDx
           Frm1.vspdData2.Row = iRow
           If Frm1.vspdData2.ColHidden <> True And Frm1.vspdData2.BackColor <>  parent.UC_PROTECTED Then
              Frm1.vspdData2.Col = iDx
              Frm1.vspdData2.Row = iRow
              Frm1.vspdData2.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub

'================================================================================================================================
Function RemovedivTextArea()

	Dim ii
		
	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next

End Function

'================================================================================================================================
Function SheetFocus(lRow, lCol)
	frm1.vspdData1.focus
	frm1.vspdData1.Row = lRow
	frm1.vspdData1.Col = lCol
	frm1.vspdData1.Action = 0
	frm1.vspdData1.SelStart = 0
	frm1.vspdData1.SelLength = len(frm1.vspdData1.Text)
End Function
