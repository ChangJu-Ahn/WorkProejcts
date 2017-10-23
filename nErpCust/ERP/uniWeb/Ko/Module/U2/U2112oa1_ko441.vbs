'**********************************************************************************************
'*  1. Module Name          : SCM
'*  2. Function Name        : 
'*  3. Program ID           : U2112oa1_KO441.ASP
'*  4. Program Name         :
'*  5. Program Desc         : 주문서 발행 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2004/07/27
'*  8. Modified date(Last)  : 2007/09/11
'*  9. Modifier (First)     : Park, BumSoo
'* 10. Modifier (Last)      : WYSO
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            
'**********************************************************************************************

Const BIZ_PGM_ID			= "U2112ob1_KO441.asp"			'☆: List SCM Orders

Dim C_OrderDt
Dim C_ItemCode
Dim C_ItemName
Dim C_Spec
Dim C_OrderUnit
Dim C_OrderNo
Dim C_OrderSeq
Dim C_po_prc
Dim C_po_loc_amt
Dim C_OrderQty
Dim C_DvryDt
Dim	C_RcptQty
Dim	C_UnRcptQty
Dim	C_FirmDvryQty
Dim C_RemainQty
Dim C_PlantCd
Dim C_PlantNm
Dim C_SLCD
Dim C_SLNM

'================================================================================================================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE
    lgIntGrpCount = 0
    lgStrPrevKey = ""
    lgLngCurRows = 0
    lgSortKey = 1
End Sub

'================================================================================================================================
Sub InitSpreadComboBox()

End Sub

'================================================================================================================================
Sub InitData(ByVal lngStartRow)

End Sub

'================================================================================================================================
Sub SetDefaultVal()
	
	frm1.txtPoFrDt.text = UniConvDateAToB(UNIDateAdd ("M", -1, LocSvrDate, parent.gServerDateFormat), parent.gServerDateFormat, parent.gDateFormat)
	frm1.txtPoToDt.text = UniConvDateAToB(UNIDateAdd ("M", 1, LocSvrDate, parent.gServerDateFormat), parent.gServerDateFormat, parent.gDateFormat)
	
	If frm1.rdoDocflg(0).checked = True Then
		LocDocFlag = "1"
	ElseIf frm1.rdoDocflg(1).checked = True Then
		LocDocFlag = "2"	
	End If

	Call SetBPCD()
	
End Sub

'================================================================================================================================
Sub SetBPCD()

	If 	CommonQueryRs2by2(" BP_NM ", " B_BIZ_PARTNER ", " BP_CD = " & FilterVar(parent.gUsrId, "", "S"), lgF0) = False Then
		Call ggoOper.SetReqAttr(frm1.txtPlantCd,"Q")
		Call ggoOper.SetReqAttr(frm1.txtItemCd,"Q")
		Call ggoOper.SetReqAttr(frm1.txtDvFrDt,"Q")
		Call ggoOper.SetReqAttr(frm1.txtDvToDt,"Q")
		Call ggoOper.SetReqAttr(frm1.rdoAppflg,"Q")
		Call ggoOper.SetReqAttr(frm1.txtPoFrDt,"Q")
		Call ggoOper.SetReqAttr(frm1.txtPoToDt,"Q")
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
Sub InitSpreadSheet()
    Call InitSpreadPosVariables()    
    With frm1
		ggoSpread.Source = .vspdData
		ggoSpread.Spreadinit "V20050420", , Parent.gAllowDragDropSpread
		.vspdData.ReDraw = False
	
		.vspdData.MaxCols = C_SlNm + 1
		.vspdData.MaxRows = 0
		
		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetDate 	C_OrderDt,		"수주일자"  ,10, 2, parent.gDateFormat		 
		ggoSpread.SSSetEdit		C_ItemCode,		"품목"      ,18,,,18,2
		ggoSpread.SSSetEdit		C_ItemName,		"품목명"    ,18
		ggoSpread.SSSetEdit		C_Spec,			"규격"      ,15
		ggoSpread.SSSetEdit		C_OrderUnit,	"단위"      , 4,,,3,2
		ggoSpread.SSSetEdit		C_OrderNo,		"수주번호"  ,15
		ggoSpread.SSSetEdit		C_OrderSeq,		"행번"      , 4
		ggoSpread.SSSetFloat	C_po_prc,		"수주단가"  ,12,parent.ggUnitCostNo, ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat 	C_po_loc_amt,	"수주금액"  ,15,parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_OrderQty,		"수주량"    ,10,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetDate 	C_DvryDt,		"납기일"    ,10, 2, parent.gDateFormat
		ggoSpread.SSSetFloat	C_RcptQty,		"납품량"    ,10,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_UnRcptQty,	"미납품량"  ,10,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_FirmDvryQty,	"납품대기량",10,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_RemainQty,	"납품잔량"  ,10,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit		C_PlantCd,		"납품공장"    , 8
		ggoSpread.SSSetEdit		C_PlantNm,		"납품공장명"  ,12
		ggoSpread.SSSetEdit		C_SlCd,			"납품창고"  ,10
		ggoSpread.SSSetEdit		C_SlNm,			"납품창고명",18
		
		Call ggoSpread.SSSetColHidden( .vspdData.MaxCols, .vspdData.MaxCols , True)
		
		.vspdData.ReDraw = True
   
		ggoSpread.SSSetSplit2(3)
    
    End With
    
    Call SetSpreadLock()
    
End Sub

'================================================================================================================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'================================================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

End Sub

'================================================================================================================================
Sub InitSpreadPosVariables()	
	C_OrderDt		= 1
	C_ItemCode		= 2
	C_ItemName		= 3
	C_Spec			= 4
	C_OrderUnit		= 5
	C_OrderNo		= 6
	C_OrderSeq		= 7
	C_po_prc		= 8
	C_po_loc_amt	= 9
	C_OrderQty		= 10
	C_DvryDt		= 11
	C_RcptQty		= 12
	C_UnRcptQty		= 13
	C_FirmDvryQty	= 14
	C_RemainQty		= 15
	C_PlantCd		= 16
	C_PlantNm		= 17
	C_SLCD			= 18
	C_SLNM			= 19
	
End Sub
 
'================================================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case Ucase(pvSpdNo)
 	Case "A"
 		ggoSpread.Source = frm1.vspdData 
 		
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

		C_OrderDt		= iCurColumnPos(1)
		C_ItemCode		= iCurColumnPos(2)
		C_ItemName		= iCurColumnPos(3)
		C_Spec			= iCurColumnPos(4)
		C_OrderUnit		= iCurColumnPos(5)
		C_OrderNo		= iCurColumnPos(6)
		C_OrderSeq		= iCurColumnPos(7)
		C_po_prc		= iCurColumnPos(8)
		C_po_loc_amt	= iCurColumnPos(9)
		C_OrderQty		= iCurColumnPos(10)
		C_DvryDt		= iCurColumnPos(11)
		C_RcptQty		= iCurColumnPos(12)
		C_UnRcptQty		= iCurColumnPos(13)
		C_FirmDvryQty	= iCurColumnPos(14)
		C_RemainQty		= iCurColumnPos(15)
		C_PlantCd		= iCurColumnPos(16)
		C_PlantNm		= iCurColumnPos(17)
		C_SlCd			= iCurColumnPos(18)
		C_SlNm			= iCurColumnPos(19)

 	End Select
 
End Sub

'================================================================================================================================
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "납품공장"
	arrParam(1) = "(			SELECT	DISTINCT B.PLANT_CD FROM M_SCM_PLAN_PUR_RCPT A, M_PUR_ORD_DTL B, M_PUR_ORD_HDR C "
	arrParam(1) = arrParam(1) & "WHERE	A.PO_NO = B.PO_NO AND A.PO_SEQ_NO = B.PO_SEQ_NO AND A.SPLIT_SEQ_NO = 0 "
	arrParam(1) = arrParam(1) & "AND	A.PO_NO = C.PO_NO AND C.BP_CD = '" & frm1.txtBpCd.value & "') A, B_PLANT B"
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = "A.PLANT_CD = B.PLANT_CD"			
	arrParam(5) = "납품공장"			
	
    arrField(0) = "A.PLANT_CD"	
    arrField(1) = "B.PLANT_NM"	
    
    arrHeader(0) = "납품공장"		
    arrHeader(1) = "납품공장명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'================================================================================================================================
Function OpenItemInfo(Byval strCode)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtItemCd.className) = "PROTECTED" Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "품목팝업"

	arrParam(1) = "(			SELECT	DISTINCT ITEM_CD FROM M_SCM_PLAN_PUR_RCPT A, M_PUR_ORD_HDR B "
	arrParam(1) = arrParam(1) & "WHERE	A.PO_NO = B.PO_NO AND A.SPLIT_SEQ_NO = 0 AND B.BP_CD = '" & frm1.txtBpCd.value & "') A, B_ITEM B"
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

'================================================================================================================================
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)
	frm1.txtPlantCd.focus()		
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
	arrParam(3) = frm1.txtPoFrDt.Text
	arrParam(4) = frm1.txtPoToDt.Text
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetTrackingNo(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtTrackingNo.focus
	
End Function
'------------------------------------------  OpenPoNo()  -------------------------------------------------
'	Name : OpenPoNo()
'	Description : OpenPoNo PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPoNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim PoFlg
	
	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "수주번호"						
	arrParam(1) = "M_PUR_ORD_HDR,B_Biz_Partner,B_PUR_GRP"					
	arrParam(2) = Trim(frm1.txtPoNo.Value)			
	'arrParam(3) = Trim(frm1.txtBpNm.Value)	
	
'	If 	frm1.rdoPoflg1.checked = true then 
		PoFlg	= "Y"			'단가표시
'	Else
'		PoFlg	= "N"	
'	End if	
	
'	arrParam(4) = "M_PUR_ORD_HDR.IMPORT_FLG = '" & PoFlg & "' AND M_PUR_ORD_HDR.BP_CD = B_Biz_Partner.BP_CD AND M_PUR_ORD_HDR.PUR_GRP = B_PUR_GRP.PUR_GRP"
	arrParam(4) = "M_PUR_ORD_HDR.release_flg	= 'Y' AND M_PUR_ORD_HDR.BP_CD = B_Biz_Partner.BP_CD AND M_PUR_ORD_HDR.PUR_GRP = B_PUR_GRP.PUR_GRP"
	arrParam(5) = "수주번호"						
	
    arrField(0) = "ED15" & Parent.gColSep &"M_PUR_ORD_HDR.PO_NO"							
    arrField(1) = "ED7" & Parent.gColSep &"M_PUR_ORD_HDR.BP_CD"							
    arrField(2) = "ED15" & Parent.gColSep &"B_Biz_Partner.BP_NM"
    arrField(3) = "DD10" & Parent.gColSep & " M_PUR_ORD_HDR.PO_DT "
    arrField(4) = "F212" & Parent.gColSep & " M_PUR_ORD_HDR.TOT_PO_DOC_AMT "
    
    if Trim(frm1.txtBpCd.Value)<>"" then
		arrParam(4) = arrParam(4) & " And M_PUR_ORD_HDR.BP_CD='" & Trim(frm1.txtBpCd.Value) & "'"    
	End if
	
'	if Trim(frm1.txtPurGrpCd.Value)<>"" then
'		arrParam(4) = arrParam(4) & " And M_PUR_ORD_HDR.PUR_GRP='" & Trim(frm1.txtPurGrpCd.Value) & "'"    
'	End if
	
	if Trim(frm1.txtPoFrDt.Text)<>"" then
		arrParam(4) = arrParam(4) & " And M_PUR_ORD_HDR.PO_DT >= '" &UNIConvDate(Trim(frm1.txtPoFrDt.Text)) & "'"  
	Else
		arrParam(4) = arrParam(4) & " And M_PUR_ORD_HDR.PO_DT >='1900-01-01'"    
	End if
	
	if Trim(frm1.txtPoToDt.Text)<>"" then
		arrParam(4) = arrParam(4) & " And M_PUR_ORD_HDR.PO_DT <= '" & UNIConvDate(Trim(frm1.txtPoToDt.Text)) & "'" 
	Else
		arrParam(4) = arrParam(4) & " And M_PUR_ORD_HDR.PO_DT <='2999-12-31'" 
	End if
	'--arrParam(4) = arrParam(4) & " order by po_no " 
    arrField(5) = "ED6" & Parent.gColSep & "M_PUR_ORD_HDR.PO_CUR"    
    arrField(6) = "ED10" & Parent.gColSep & "B_PUR_GRP.PUR_GRP_NM"
    
    
    
    arrHeader(0) = "수주번호"						
    arrHeader(1) = "공급처"					
    arrHeader(2) = "공급처명"					
    arrHeader(3) = "발주일"					
    arrHeader(4) = "발주금액"					
    arrHeader(5) = "화폐"					
    arrHeader(6) = "구매그룹"					
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=600px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPoNo.focus
		Exit Function
	Else
		frm1.txtPoNo.Value = arrRet(0)		
		frm1.txtPoNo.focus
	End If	
End Function
'------------------------------------------  SetTrackingNo()  -----------------------------------------
'	Name : SetTrackingNo()
'	Description : Tracking No. Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetTrackingNo(Byval arrRet)
    frm1.txtTrackingNo.Value = arrRet(0)
End Function

'================================================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub
 
'================================================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row)
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

'================================================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )

	If lgIntFlgMode = Parent.OPMD_UMODE Then
  		Call SetPopupMenuItemInf("1101111111")
  	Else
  		Call SetPopupMenuItemInf("1001111111")
  	End If

	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData

	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If
   	
   	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
	        If lgSortKey = 1 Then
        	    ggoSpread.SSSort Col
	            lgSortKey = 2
	        Else
        	    ggoSpread.SSSort Col, lgSortKey
	            lgSortKey = 1
        	End If
	End If

End Sub

'================================================================================================================================
Sub vspdData_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub

'================================================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
   
 	If Row <= 0 Then
		Exit Sub
 	End If
    
  	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
 	End If

End Sub
 
'================================================================================================================================
Sub vspddata_KeyPress(index , KeyAscii )
    On Error Resume Next                                                    '☜: Protect system from crashing
End Sub


'================================================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'================================================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
'    On Error Resume Next                                                    '☜: Protect system from crashing
	if NewRow = -1 then
		exit sub
	end if
	Call GetSpreadColumnPos("A")
	frm1.vspdData.col = C_RcptQty
	frm1.vspdData.Row = NewRow
	
	If frm1.rdoDocflg(0).checked = True Then
		LocDocFlag = "1"
	ElseIf frm1.rdoDocflg(1).checked = True Then
		LocDocFlag = "2"	
	End If
	
	If LocDocFlag = "1" Then
		frm1.btnRun(0).disabled = false
		frm1.btnRun(1).disabled = false
	Else
		If Cdbl("0" & frm1.vspdData.Text) = 0 then
			frm1.btnRun(0).disabled = True
			frm1.btnRun(1).disabled = True
		Else
			frm1.btnRun(0).disabled = false
			frm1.btnRun(1).disabled = false
		End if
	End If

End Sub

'================================================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
    On Error Resume Next                                                    '☜: Protect system from crashing
End Sub

'================================================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	
	If CheckRunningBizProcess = True Then Exit Sub
	
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If	
		End If
    End if
    
End Sub

'================================================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		If Row < 1 Then Exit Sub

    End With
    
End Sub

'================================================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 
 
'================================================================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
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
    Call InitSpreadSheet
    Call InitSpreadComboBox
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
Sub txtPoFrDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtPoFrDt.Action = 7
        SetFocusToDocument("M")
		Frm1.txtPoFrDt.Focus
    End If
End Sub

'================================================================================================================================
Sub txtPoToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtPoToDt.Action = 7
        SetFocusToDocument("M")
		Frm1.txtPoToDt.Focus
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

'================================================================================================================================
Sub txtPoFrDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'================================================================================================================================
Sub txtPoToDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'================================================================================================================================
Function BtnPreview()
    
    Dim strEbrFile
    Dim objName
	
	Dim var1
	Dim var2
	Dim var3
	Dim var4
	Dim var5
	dim var6
	dim var7
	
	dim strUrl
	dim arrParam, arrField, arrHeader

	If lgIntFlgMode = Parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If

	Call BtnDisabled(1)
	
	Call GetSpreadColumnPos("A")
	
	If frm1.rdoDocflg(0).checked = True Then
		LocDocFlag = "1"
	ElseIf frm1.rdoDocflg(1).checked = True Then
		LocDocFlag = "2"	
	End If

	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_OrderNo
	
	var1 = Trim(frm1.vspdData.text)
	
	var2 = "%"
	
	If frm1.hBPCd.value = "" Then
		var3 = "%"
	Else
		var3 = Trim(frm1.hBPCd.value)
	End If
	
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_OrderSeq
	var4 = Trim(frm1.vspdData.text)
	
	If frm1.hPoFrDt.value = "" Then
		var5 = UniConvDateAtoB(UniConvYYYYMMDDToDate(parent.gDateFormat, "1900", "01", "01"),parent.gDateFormat,parent.gServerDateFormat)
	Else
		var5 = UniConvDateAtoB(frm1.hPoFrDt.value,parent.gDateFormat,parent.gServerDateFormat)
	End If
	
	If frm1.hPoToDt.value = "" Then
		var6 = UniConvDateAtoB(UniConvYYYYMMDDToDate(parent.gDateFormat, "2999", "12", "31"),parent.gDateFormat,parent.gServerDateFormat)
	Else
		var6 = UniConvDateAtoB(frm1.hPoToDt.value,parent.gDateFormat,parent.gServerDateFormat)
	End If
	
	var7 = "Y"	'단가표시
	
	If LocDocFlag <> "1" Then
		strUrl = "PO_NO|"					& var1 
		strUrl = strUrl		& "|PO_SEQ_NO|" & var4
		strEbrFile = "U2112oa1" 
		
	Else
		strUrl =  "po_no|"					& var1
		strUrl = strUrl & "|po_no1|"		& var1
		strUrl = strUrl & "|bp_cd|"			& var3
		strUrl = strUrl & "|pur_grp|"		& var2
		strUrl = strUrl & "|fr_dt|"			& var5
		strUrl = strUrl & "|to_dt|"			& var6
		strUrl = strUrl & "|Gb_fg|"			& var7
		
		strEbrFile = "M3111OA2_KO441"
	End If
	
	objName = AskEBDocumentName(strEbrFile,"ebr")
	call FncEBRPreview(objName, strUrl)
	
	Call BtnDisabled(0)
	
	frm1.btnRun(0).focus
	Set gActiveElement = document.activeElement
	
End Function

'================================================================================================================================
Function BtnPrint()
	
	Dim strEbrFile
    Dim objName
	
	Dim var1
	Dim var2
	Dim var3
	Dim var4
	Dim var5
	dim var6
	dim var7
	
	dim strUrl
	dim arrParam, arrField, arrHeader

	If lgIntFlgMode = Parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If

	Call BtnDisabled(1)	

	Call GetSpreadColumnPos("A")
	
	If frm1.rdoDocflg(0).checked = True Then
		LocDocFlag = "1"
	ElseIf frm1.rdoDocflg(1).checked = True Then
		LocDocFlag = "2"	
	End If

	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_OrderNo
	
	var1 = Trim(frm1.vspdData.text)
	
	var2 = "%"
	
	If frm1.hBPCd.value = "" Then
		var3 = "%"
	Else
		var3 = Trim(frm1.hBPCd.value)
	End If
	
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_OrderSeq
	var4 = Trim(frm1.vspdData.text)
	
	If frm1.hPoFrDt.value = "" Then
		var5 = UniConvDateAtoB(UniConvYYYYMMDDToDate(parent.gDateFormat, "1900", "01", "01"),parent.gDateFormat,parent.gServerDateFormat)
	Else
		var5 = UniConvDateAtoB(frm1.hPoFrDt.value,parent.gDateFormat,parent.gServerDateFormat)
	End If
	
	If frm1.hPoToDt.value = "" Then
		var6 = UniConvDateAtoB(UniConvYYYYMMDDToDate(parent.gDateFormat, "2999", "12", "31"),parent.gDateFormat,parent.gServerDateFormat)
	Else
		var6 = UniConvDateAtoB(frm1.hPoToDt.value,parent.gDateFormat,parent.gServerDateFormat)
	End If

	var7 = "Y"	'단가표시
	
	If LocDocFlag <> "1" Then
		strUrl = "PO_NO|"					& var1 
		strUrl = strUrl		& "|PO_SEQ_NO|" & var4
		
		strEbrFile = "U2112oa1"
	Else
		strUrl =  "po_no|"					& var1
		strUrl = strUrl & "|po_no1|"		& var1
		strUrl = strUrl & "|bp_cd|"			& var3
		strUrl = strUrl & "|pur_grp|"		& var2
		strUrl = strUrl & "|fr_dt|"			& var5
		strUrl = strUrl & "|to_dt|"			& var6
		strUrl = strUrl & "|Gb_fg|"			& var7
		
		strEbrFile = "M3111OA2_KO441"
	End If
	
	objName = AskEBDocumentName(strEbrFile,"ebr")
	
	call FncEBRprint(EBAction, objName, strUrl)
	
	Call BtnDisabled(0)	
	
	frm1.btnRun(1).focus
	Set gActiveElement = document.activeElement

End Function

'=======================================================================================================
Function FncQuery()
 
    Dim IntRetCD 
    
    ncQuery = False
    
    Err.Clear

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x")
	If IntRetCD = vbNo Then
	    Exit Function
	End If
    End If

    If ValidDateCheck(frm1.txtDvFrDt, frm1.txtDvToDt) = False Then Exit Function
    If ValidDateCheck(frm1.txtPoFrDt, frm1.txtPoToDt) = False Then Exit Function
	
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")
    ggoSpread.Source = frm1.vspdData
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
	On Error Resume Next
End Function

'================================================================================================================================
Function FncCopy() 
	On Error Resume Next
End Function

'================================================================================================================================
Function FncPaste() 
     ggoSpread.SpreadPaste
End Function

'================================================================================================================================
Function FncCancel() 
	On Error Resume Next
End Function

'================================================================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
	On Error Resume Next
End Function

'================================================================================================================================
Function FncDeleteRow() 
	On Error Resume Next
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
    
    ggoSpread.Source = frm1.vspdData							'⊙: Preset spreadsheet pointer 
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
Sub MakeKeyStream()

	With frm1
		If lgIntFlgMode = parent.OPMD_UMODE Then
			lgKeyStream = UCase(Trim(.hPlantCd.value))  & Parent.gColSep
			lgKeyStream = lgKeyStream & UCase(Trim(.hItemCd.value))  & Parent.gColSep
			lgKeyStream = lgKeyStream & UCase(Trim(.hBPCd.value))  & Parent.gColSep
			lgKeyStream = lgKeyStream & Trim(.hDvFrDt.value)  & Parent.gColSep
			lgKeyStream = lgKeyStream & Trim(.hDvToDt.value)  & Parent.gColSep
			lgKeyStream = lgKeyStream & Trim(.hrdoAppflg.value)  & Parent.gColSep
			lgKeyStream = lgKeyStream & Trim(.hPoFrDt.value)  & Parent.gColSep
			lgKeyStream = lgKeyStream & Trim(.hPoToDt.value)  & Parent.gColSep
			lgKeyStream = lgKeyStream & UCase(Trim(.hTRACKINGNO.value))  & Parent.gColSep
			lgKeyStream = lgKeyStream & UCase(Trim(.hPoNo.value))  & Parent.gColSep
		Else
			lgKeyStream = UCase(Trim(.txtPlantCd.value))  & Parent.gColSep
			lgKeyStream = lgKeyStream & UCase(Trim(.txtItemCd.value))  & Parent.gColSep
			lgKeyStream = lgKeyStream & UCase(Trim(.txtBPCd.value))  & Parent.gColSep
			lgKeyStream = lgKeyStream & Trim(.txtDvFrDt.Text)  & Parent.gColSep
			lgKeyStream = lgKeyStream & Trim(.txtDvToDt.Text)  & Parent.gColSep
			If .rdoAppflg(0).checked = true Then
				lgKeyStream = lgKeyStream & "A" & Parent.gColSep
			ElseIf .rdoAppflg(1).checked = true Then
				lgKeyStream = lgKeyStream & "N" & Parent.gColSep
			Else
				lgKeyStream = lgKeyStream & "Y" & Parent.gColSep
			End If
			lgKeyStream = lgKeyStream & Trim(.txtPoFrDt.Text)  & Parent.gColSep
			lgKeyStream = lgKeyStream & Trim(.txtPoToDt.Text)  & Parent.gColSep
			lgKeyStream = lgKeyStream & UCase(Trim(.txtTRACKINGNO.value))  & Parent.gColSep
			lgKeyStream = lgKeyStream & UCase(Trim(.txtPoNo.value))  & Parent.gColSep
			
			.hPlantCd.value		= .txtPlantCd.value
			.hItemCd.value		= .txtItemCd.value
			.hBPCd.value		= .txtBPCd.value
			.hDvFrDt.value		= .txtDvFrDt.Text
			.hDvToDt.value		= .txtDvToDt.Text
			If .rdoAppflg(0).checked = true Then
				.hrdoAppflg.value = "A"
			ElseIf .rdoAppflg(1).checked = true Then
				.hrdoAppflg.value = "N"
			Else
				.hrdoAppflg.value = "Y"
			End If
			.hPoFrDt.value		= .txtPoFrDt.Text
			.hPoToDt.value		= .txtPoToDt.Text	
			.htrackingno.value	= .txttrackingno.value
			.hPoNo.value		= .txtPoNo.value
			
		End If
	End With 
End Sub

'================================================================================================================================
Function DbQuery() 
    
    Dim strVal
    
    Err.Clear

    DbQuery = False
    
    Call LayerShowHide(1)
 
    Call MakeKeyStream()
    
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
    Call SetToolBar("11000000000111")

    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE
	
	Call GetSpreadColumnPos("A")
	frm1.vspdData.col = C_RcptQty
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	
	If frm1.rdoDocflg(0).checked = True Then
		LocDocFlag = "1"
	ElseIf frm1.rdoDocflg(1).checked = True Then
		LocDocFlag = "2"	
	End If
	
	If LocDocFlag = "1" Then
		frm1.btnRun(0).disabled = false
		frm1.btnRun(1).disabled = false
	Else
		If Cdbl("0" & frm1.vspdData.Text) = 0 then
			frm1.btnRun(0).disabled = True
			frm1.btnRun(1).disabled = True
		Else
			frm1.btnRun(0).disabled = false
			frm1.btnRun(1).disabled = false
		End if
	End If

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
Function DbSave() 
	On Error Resume Next    
End Function

'================================================================================================================================
Function DbSaveOk()
	On Error Resume Next
End Function

'================================================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr, parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <>  parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
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
	frm1.vspdData.focus
	frm1.vspdData.Row = lRow
	frm1.vspdData.Col = lCol
	frm1.vspdData.Action = 0
	frm1.vspdData.SelStart = 0
	frm1.vspdData.SelLength = len(frm1.vspdData.Text)
End Function
                                                                                            