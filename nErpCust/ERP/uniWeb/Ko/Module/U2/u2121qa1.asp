<%@ LANGUAGE="VBSCRIPT" %>
<!--**********************************************************************************************
'*  1. Module Name          : SCM
'*  2. Function Name        : 
'*  3. Program ID           : u2121qa1.asp
'*  4. Program Name         : 수주예정현황 - Query Return Summary
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2004-07-25
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Park, BumSoo
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. History              : 
'********************************************************************************************** -->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'########################################################################################################## -->
<!-- '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'================================================================================================================================
Const BIZ_PGM_QRY_ID	= "u2121qb1.asp"							'☆: 비지니스 로직 ASP명 

'================================================================================================================================
' Grid(vspdData)
Dim C_ItemCd
Dim C_ItemNm
Dim C_Spec
Dim C_BpCd
Dim C_BpNm
Dim C_Dt
Dim C_Unit
Dim C_Qty
Dim C_DT2

'================================================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'================================================================================================================================
Dim IsOpenPop 
Dim lgAfterQryFlg
Dim lgLngCnt
Dim lgOldRow

Dim strDate
Dim iDBSYSDate

'================================================================================================================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0

    lgStrPrevKey = ""
    lgLngCurRows = 0
    lgAfterQryFlg = False
    lgLngCnt = 0
    lgOldRow = 0
    lgSortKey = 1

End Sub

'================================================================================================================================
Sub SetDefaultVal()
	Dim strDate
	Dim BaseDate
	Dim strYear
	Dim strMonth
	Dim strDay

	BaseDate = "<%=GetSvrDate%>"

	Call ExtractDateFrom(BaseDate, parent.gServerDateFormat, parent.gServerDateType, strYear, StrMonth, StrDay)
	strDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)
	
	frm1.txtDvFrDt.text = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, "01")
	frm1.txtDvToDt.text   = strDate
	
	frm1.txtFrDt.text = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, "01")
	frm1.txtToDt.text   = strDate
	
	Call SetBPCD()
	
End Sub

Sub SetBPCD()

	If 	CommonQueryRs2by2(" BP_NM ", " B_BIZ_PARTNER ", " BP_CD = " & FilterVar(parent.gUsrId, "", "S"), lgF0) = False Then
		Call ggoOper.SetReqAttr(frm1.txtPlantCd,"Q")
		Call ggoOper.SetReqAttr(frm1.txtItemCd,"Q")
		Call ggoOper.SetReqAttr(frm1.txtDvFrDt,"Q")
		Call ggoOper.SetReqAttr(frm1.txtDvToDt,"Q")
		Call ggoOper.SetReqAttr(frm1.txtFrDt,"Q")
		Call ggoOper.SetReqAttr(frm1.txtToDt,"Q")
		Call ggoOper.SetReqAttr(frm1.txtTrackingNo,"Q")
		Call DisplayMsgBox("210033","X","X","X")
		Call SetToolBar("10000000000011")
		Exit Sub
	Else
	    Call SetToolBar("11000000000011")										'⊙: 버튼 툴바 제어 
	End If

	lgF0 = Split(lgF0, Chr(11))
	frm1.txtBpCd.value = parent.gUsrId
	frm1.txtBpNm.value = lgF0(1)

End Sub

'================================================================================================================================
Sub LoadInfTB19029()     
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q","P","NOCOOKIE","MA") %>
End Sub

'================================================================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)

	Call InitSpreadPosVariables(pvSpdNo)

	If pvSpdNo = "A" Or pvSpdNo = "*" Then
		'------------------------------------------
		' Grid 1 - Operation Spread Setting
		'------------------------------------------
		With frm1.vspdData 
			
			ggoSpread.Source = frm1.vspdData
			ggoSpread.Spreadinit "V20050420", ,Parent.gAllowDragDropSpread
					
			.ReDraw = false
					
			.MaxCols = C_Dt2 + 1    
			.MaxRows = 0    
			
			Call GetSpreadColumnPos("A")

			ggoSpread.SSSetEdit 	C_ItemCd,       "품목"			,18
			ggoSpread.SSSetEdit 	C_ItemNm,       "품목명"		,20
			ggoSpread.SSSetEdit 	C_Spec,			"규격"			,18
			ggoSpread.SSSetEdit 	C_BpCd,         "납품처"		,15
			ggoSpread.SSSetEdit 	C_BPNm,         "납품처명"		,20
			ggoSpread.SSSetDate 	C_Dt,  		    "수주예정일"	,12, 2, parent.gDateFormat		 
			ggoSpread.SSSetEdit 	C_Unit,		    "단위"			, 8
			ggoSpread.SSSetFloat	C_Qty,		    "수주예정수량"	,13,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetDate 	C_Dt2,		    "납기예정일"	,11, 2, parent.gDateFormat		 
			
			Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols, True)
			
			'ggoSpread.SSSetSplit2(2)
			
			Call SetSpreadLock("A")
			
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
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadLockWithOddEvenRowColor()
	End If
		
End Sub

'================================================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

End Sub

'================================================================================================================================
Sub InitComboBox()

End Sub

'================================================================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)
	If pvSpdNo = "A" Or pvSpdNo = "*" Then
		' Grid 1(vspdData)

		C_ItemCd		= 1
		C_ItemNm		= 2
		C_Spec			= 3
		C_BpCd			= 4
		C_BpNm			= 5
		C_Dt			= 6
		C_Unit			= 7
		C_Qty			= 8
		C_Dt2			= 9
		
	End If	
	
End Sub

'================================================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
      
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
		Case "A"
		
 			ggoSpread.Source = frm1.vspdData
		
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_ItemCd		= iCurColumnPos(1)
			C_ItemNm		= iCurColumnPos(2)
			C_Spec			= iCurColumnPos(3)
			C_BpCd			= iCurColumnPos(4)
			C_BpNm			= iCurColumnPos(5)
			C_Dt			= iCurColumnPos(6)
			C_Unit			= iCurColumnPos(7)
			C_Qty			= iCurColumnPos(8)
			C_Dt2			= iCurColumnPos(9)
							
    End Select

End Sub    

'================================================================================================================================
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "납품처팝업"
	arrParam(1) = "(			SELECT	DISTINCT B.PLANT_CD FROM M_SCM_PLAN_PUR_RCPT A, M_PUR_ORD_DTL B, M_PUR_ORD_HDR C "
	arrParam(1) = arrParam(1) & "WHERE	A.PO_NO = B.PO_NO AND A.PO_SEQ_NO = B.PO_SEQ_NO AND A.SPLIT_SEQ_NO = 0 "
	arrParam(1) = arrParam(1) & "AND	A.PO_NO = C.PO_NO AND C.BP_CD = '" & frm1.txtBpCd.value & "') A, B_PLANT B"
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = "A.PLANT_CD = B.PLANT_CD"			
	arrParam(5) = "납품처"			
	
    arrField(0) = "A.PLANT_CD"	
    arrField(1) = "B.PLANT_NM"	
    
    arrHeader(0) = "납품처"		
    arrHeader(1) = "납품처명"		
    
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
Function OpenItemInfo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "품목팝업"
	'arrParam(1) = "(			SELECT	DISTINCT ITEM_CD FROM M_SCM_PLAN_PUR_RCPT A, M_PUR_ORD_HDR B "
	'arrParam(1) = arrParam(1) & "WHERE	A.PO_NO = B.PO_NO AND A.SPLIT_SEQ_NO = 0 AND B.BP_CD = '" & frm1.txtBpCd.value & "') A, B_ITEM B"
	'arrParam(2) = Trim(frm1.txtItemCd.Value)								' Code Condition
	
	arrParam(1) = " B_ITEM "
	arrParam(3) = ""														' Name Cindition
	'arrParam(4) = "A.ITEM_CD = B.ITEM_CD "
	arrParam(4) = " "
	arrParam(5) = "품목"
	 
    arrField(0) = "ITEM_CD"												' Field명(0)
    arrField(1) = "ITEM_NM"												' Field명(1)
    
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
	frm1.txtItemCd.value = arrRet(0)
	frm1.txtItemNm.value = arrRet(1)
	frm1.txtItemCd.focus()
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


'================================================================================================================================
Sub Form_Load()
    Call LoadInfTB19029

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")
    Call InitSpreadSheet("*")
   
    Call SetDefaultVal
    Call InitVariables
    Call InitComboBox
 
	frm1.txtPlantCd.focus
	Set gActiveElement = document.activeElement

End Sub

'================================================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'================================================================================================================================
Sub txtDvFrDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtDvFrDt.Action = 7 
		Call SetFocusToDocument("M")
		frm1.txtDvFrDt.Focus
	End If 
End Sub

'================================================================================================================================
Sub txtDvToDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtDvToDt.Action = 7 
		Call SetFocusToDocument("M")
		frm1.txtDvToDt.Focus
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
Sub txtFrDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtFrDt.Action = 7 
		Call SetFocusToDocument("M")
		frm1.txtFrDt.Focus
	End If 
End Sub

'================================================================================================================================
Sub txtToDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtToDt.Action = 7 
		Call SetFocusToDocument("M")
		frm1.txtToDt.Focus
	End If 
End Sub

'================================================================================================================================
Sub txtFrDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'================================================================================================================================
Sub txtToDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'================================================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'================================================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then
        Exit Sub
	End If
        
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
Sub vspdData_Click(ByVal Col , ByVal Row )
	
	Call SetPopupMenuItemInf("0000111111")
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
		
	If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

'================================================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'================================================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'================================================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

End Sub

'================================================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    With frm1.vspdData 
		If Row >= NewRow Then
		    Exit Sub
		End If
    End With
End Sub

'================================================================================================================================
Function FncQuery() 
    Dim IntRetCD 

    FncQuery = False
    Err.Clear

	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = "" 
	End If	
	
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = "" 
	End If

	If ValidDateCheck(frm1.txtDvFrDt, frm1.txtDvToDt) = False Then Exit Function
		
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
    If DbQuery = False Then
		Call RestoreToolBar()
		Exit Function														'☜: Query db data
	End If
	
    FncQuery = True															'⊙: Processing is OK
End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
	On Error Resume Next
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
	On Error Resume Next    
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
	On Error Resume Next    
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	On Error Resume Next	
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	On Error Resume Next	
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow() 
	On Error Resume Next	
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
	On Error Resume Next	
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
    On Error Resume Next													'☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    On Error Resume Next													'☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)									'☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)								'☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
End Sub

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	FncExit = True
End Function

'******************  5.2 Fnc함수명에서 호출되는 개발 Function  **************************
'	설명 : 
'**************************************************************************************** 

'========================================================================================================
' Name : MakeKeyStream
' Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)

   Select Case pOpt
       Case "M"
       
				With frm1
					If lgIntFlgMode = parent.OPMD_UMODE Then
						lgKeyStream = UCase(Trim(.hPlantCd.value))  & Parent.gColSep
						lgKeyStream = lgKeyStream & UCase(Trim(.hItemCd.value))  & Parent.gColSep
						lgKeyStream = lgKeyStream & UCase(Trim(.hBPCd.value))  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(.hDvryFromDt.value)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(.hDvryToDt.value)  & Parent.gColSep
						lgKeyStream = lgKeyStream & UCase(Trim(.hTRACKINGNO.value))  & Parent.gColSep
					Else
						lgKeyStream =               UCase(Trim(.txtBPCd.value))  & Parent.gColSep
						lgKeyStream = lgKeyStream & UCase(Trim(.txtPlantCd.value))  & Parent.gColSep
						lgKeyStream = lgKeyStream & UCase(Trim(.txtItemCd.value))  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(.txtFrDt.Text)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(.txtToDt.Text)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(.txtDvFrDt.Text)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(.txtDvToDt.Text)  & Parent.gColSep
						lgKeyStream = lgKeyStream & UCase(Trim(.txtTrackingNo.value))  & Parent.gColSep
						
						.hPlantCd.value		= .txtPlantCd.value
						.hItemCd.value		= .txtItemCd.value
						.hBPCd.value		= .txtBPCd.value
						.hDvryFromDt.value	= .txtDvFrDt.Text
						.hDvryToDt.value	= .txtDvToDt.Text
						.hTRACKINGNO.value	= .txtTrackingNo.value
					End If
				End With
			
	End Select
   
End Sub    

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 

	Dim strVal

    DbQuery = False

	Call LayerShowHide(1)
    
    Call MakeKeyStream("M")
    
	strVal = BIZ_PGM_QRY_ID & "?txtMode="	& parent.UID_M0001
	strVal = strVal & "&txtKeyStream="     & lgKeyStream
    
    Call RunMyBizASP(MyBizASP, strVal)
    
    DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()

	Call SetToolBar("11000000000111")														'⊙: 버튼 툴바 제어 
	
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
	End If

	lgIntFlgMode = parent.OPMD_UMODE														'⊙: Indicates that current mode is Update mode
    lgBlnFlgChgValue = False
	lgAfterQryFlg = True
	lgOldRow = 1
		
End Function

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Function Desc : 그리드 현상태를 저장한다.
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub 

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : 그리드를 예전 상태로 복원한다.
'========================================================================================
Sub PopRestoreSpreadColumnInf()
	Dim LngRow

    ggoSpread.Source = gActiveSpdSheet
    
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet(gActiveSpdSheet.Id)
	Call ggoSpread.ReOrderingSpreadData()
	
End Sub 

'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------


'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>수주예정현황조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=500>&nbsp;</TD>
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
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
			 					<TR>
			 						<TD CLASS=TD5 NOWRAP>업체</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 tag="14xxxU" ALT="업체">&nbsp;<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=25 tag="14" ALT="업체명"></TD>
			 						<TD CLASS=TD5 NOWRAP>납품공장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="11xxxU" ALT="납품공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14" ALT="공장명"></TD>
								</TR>
			 					<TR>
									<TD CLASS=TD5 NOWRAP>품목</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>필요납기일</TD> 
									<TD CLASS=TD6>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtFrDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="11" ALT="시작일" id=OBJECT1></OBJECT>');</SCRIPT>
										&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtToDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="11" ALT="종료일" id=OBJECT2></OBJECT>');</SCRIPT>
									</TD>
								</TR>
								<TR>	
									<TD CLASS=TD5 NOWRAP>수주예정일</TD> 
									<TD CLASS=TD6>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtDvFrDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="수주예정시작일"></OBJECT>');</SCRIPT>
										&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtDvToDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="수주예정종료일"></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS=TD5 NOWRAP>Tracking No.</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTrackingNo" SIZE=25 MAXLENGTH=25 tag="11xxxU" ALT="Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNoBtn" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenTrackingInfo()"></TD>
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
							<TR HEIGHT="100%">
								<TD WIDTH="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> name=vspdData Id = "A" HEIGHT="100%" width="100%" tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hDvryFromDt" tag="24"><INPUT TYPE=HIDDEN NAME="hDvryToDt" tag="24"><INPUT TYPE=HIDDEN NAME="hItemCd" tag="24"><INPUT TYPE=HIDDEN NAME="hBpCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hTRACKINGNO" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>