<%@ LANGUAGE="VBSCRIPT" %>
<!--**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : U1116QA1.asp
'*  4. Program Name         : �԰����Ȳ - Query Receipt Details 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2004-08-02
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : NHG
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. History              : 
'********************************************************************************************** -->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. �� �� �� 
'########################################################################################################## -->
<!-- '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'================================================================================================================================
Const BIZ_PGM_QRY_ID	= "U1116QB1.asp"							'��: �����Ͻ� ���� ASP�� 

'================================================================================================================================
' Grid (vspdData)
Dim C_PONo
Dim C_POSeq
Dim C_ItemCd
Dim C_ItemNm
Dim C_Spec
Dim C_MvmtDt
Dim C_MvmtUnit
Dim C_DvryQty
Dim C_DvryPrice
Dim C_DvryAmt
Dim C_LotNo
Dim C_MakerLotNo
Dim C_BPCd
Dim C_BPNm
Dim C_SlCd
Dim C_SlNm

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
	
	frm1.txtDvryFromDt.text = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, "01")
	frm1.txtDvryToDt.text   = strDate
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
			ggoSpread.Spreadinit "V20021224", ,Parent.gAllowDragDropSpread
					
			.ReDraw = false
					
			.MaxCols = C_SlNm + 1    
			.MaxRows = 0    
			
			Call GetSpreadColumnPos("A")

			ggoSpread.SSSetEdit 	C_PONo,			"���ֹ�ȣ"		,15
			ggoSpread.SSSetEdit		C_POSeq,		"����"			,4
			ggoSpread.SSSetEdit 	C_ItemCd,       "ǰ��"			,18
			ggoSpread.SSSetEdit 	C_ItemNm,       "ǰ���"		,20
			ggoSpread.SSSetEdit 	C_Spec,			"�԰�"			,20
			ggoSpread.SSSetDate 	C_MvmtDt,		"�԰���"		,10, 2, parent.gDateFormat		 
			ggoSpread.SSSetEdit 	C_MvmtUnit,		"����"			,4
			ggoSpread.SSSetFloat	C_DvryQty,		"�԰����"		,12,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_DvryPrice,	"�԰�ܰ�"		,12,parent.ggUnitCostNo ,ggStrIntegeralPart ,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
			ggoSpread.SSSetFloat	C_DvryAmt,		"�԰�ݾ�"		,15,parent.ggAmtOfMoneyNo ,ggStrIntegeralPart ,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
			ggoSpread.SSSetEdit 	C_LotNo,		"Lot No."		,18
			ggoSpread.SSSetEdit 	C_MakerLotNo,   "MAKER LOT NO."	,18
			ggoSpread.SSSetEdit 	C_BpCd,			"����ó"		,6
			ggoSpread.SSSetEdit 	C_BpNm,			"����ó��"		,15
			ggoSpread.SSSetEdit 	C_SLCd,			"����â��"		,6
			ggoSpread.SSSetEdit 	C_SLNm,			"����â���"	,15			
			
'			Call ggoSpread.SSSetColHidden( C_LotNo, C_MakerLotNo, True)
			Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols, True)
			
			ggoSpread.SSSetSplit2(2)
			
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
		C_PONo			= 1
		C_POSeq			= 2
		C_ItemCd		= 3
		C_ItemNm		= 4
		C_Spec			= 5
		C_MvmtDt		= 6
		C_MvmtUnit		= 7
		C_DvryQty		= 8
		C_DvryPrice		= 9
		C_DvryAmt		= 10
		C_LotNo			= 11
		C_MakerLotNo	= 12
		C_BPCd			= 13
		C_BPNm			= 14
		C_SLCd			= 15
		C_SLNm			= 16
	End If	
End Sub

'================================================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
      
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
		Case "A"
		
 			ggoSpread.Source = frm1.vspdData
		
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		
			C_PONo			= iCurColumnPos(1)
			C_POSeq			= iCurColumnPos(2)
			C_ItemCd		= iCurColumnPos(3)
			C_ItemNm		= iCurColumnPos(4)
			C_Spec			= iCurColumnPos(5)
			C_MvmtDt		= iCurColumnPos(6)
			C_MvmtUnit		= iCurColumnPos(7)
			C_DvryQty		= iCurColumnPos(8)
			C_DvryPrice		= iCurColumnPos(9)
			C_DvryAmt		= iCurColumnPos(10)
			C_LotNo			= iCurColumnPos(11)
			C_MakerLotNo	= iCurColumnPos(12)
			C_BPCd			= iCurColumnPos(13)
			C_BPNm			= iCurColumnPos(14)
			C_SLCd			= iCurColumnPos(15)
			C_SLNm			= iCurColumnPos(16)
		
    End Select

End Sub    

'================================================================================================================================
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "����"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "����"		
    arrHeader(1) = "�����"		
    
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

	If IsOpenPop = True Then Exit Function
	
	if  Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "����", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End if
	
	IsOpenPop = True

	arrParam(0) = "ǰ��"							' �˾� ��Ī 
	arrParam(1) = "B_Item_By_Plant,B_Plant,B_Item"	' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtItemCd.Value)		' Code Condition
'	arrParam(3) = Trim(frm1.txtItemNm.Value)		' Name Condition
	arrParam(4) = "B_Item_By_Plant.Plant_Cd = B_Plant.Plant_Cd And "
	arrParam(4) = arrParam(4) & "B_Item_By_Plant.Item_Cd = B_Item.Item_Cd and B_Item.phantom_flg = 'N' "
	if Trim(frm1.txtPlantCd.Value)<>"" then
		arrParam(4) = arrParam(4) & "And B_Plant.Plant_Cd='" & FilterVar(Trim(frm1.txtPlantCd.Value), "", "SNM") & "'"    
	End if
	arrParam(5) = "ǰ��"							' TextBox ��Ī 

    arrField(0) = "B_Item.Item_Cd"					' Field��(0)
    arrField(1) = "B_Item.Item_NM"					' Field��(1)
    arrField(2) = "B_Plant.Plant_Cd"				' Field��(2)
    arrField(3) = "B_Plant.Plant_NM"				' Field��(3)
    
    arrHeader(0) = "ǰ��"							' Header��(0)
    arrHeader(1) = "ǰ���"							' Header��(1)
    arrHeader(2) = "����"							' Header��(2)
    arrHeader(3) = "�����"							' Header��(3)
    
	arrRet = window.showModalDialog("../m1/m1111pa1.asp", Array(window.parent,arrParam, arrField, arrHeader), _
		"dialogWidth=695px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetItemInfo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus

End Function

'================================================================================================================================
Function OpenBpCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtBpCd.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "����ó"										' �˾� ��Ī 
	arrParam(1) = "B_Biz_Partner"								' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtBpCd.Value)						' Code Condition
	arrParam(3) = ""
	arrParam(4) = "BP_TYPE In ('S','CS') And usage_flag='Y'"	' Where Condition
	arrParam(5) = "����ó"										' TextBox ��Ī 
	
    arrField(0) = "BP_CD"										' Field��(0)
    arrField(1) = "BP_NM"										' Field��(1)
    
    arrHeader(0) = "����ó"										' Header��(0)
    arrHeader(1) = "����ó��"									' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBpCd.focus
		Exit Function
	Else
		frm1.txtBpCd.Value = arrRet(0)
		frm1.txtBpNm.Value = arrRet(1)
		frm1.txtBpCd.focus
	End If	
End Function

'------------------------------------------  OpenPoNo()  -------------------------------------------------
'	Name : OpenPoNo()
'	Description : Purchase_Order PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPoNo()
	
	Dim strRet
	Dim arrParam(2)
	Dim iCalledAspName
	Dim IntRetCD
			
	If IsOpenPop = True Or UCase(frm1.txtPoNo.className) = Ucase(Parent.UCN_PROTECTED) Then Exit Function
		
	IsOpenPop = True
		
	arrParam(0) = frm1.txtPlantCd.value
	arrParam(1) = frm1.txtBpCd.value
	arrParam(2) = frm1.txtBpNm.value

	iCalledAspName = AskPRAspName("U1111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "U1111PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If strRet(0) = "" Then
		Exit Function
	Else
		frm1.txtPoNo.value = strRet(0)
		frm1.txtPoNo.focus	
		Set gActiveElement = document.activeElement
	End If	
		
End Function

'================================================================================================================================
Function OpenPurGrp()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtBpCd.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���ű׷�"										' �˾� ��Ī 
	arrParam(1) = "B_Pur_Grp"								' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtpURgRP.Value)						' Code Condition
	arrParam(3) = ""
	arrParam(4) = "B_Pur_Grp.USAGE_FLG='Y'"	' Where Condition
	arrParam(5) = "���ű׷�"										' TextBox ��Ī 
	
    arrField(0) = "PUR_GRP"										' Field��(0)
    arrField(1) = "PUR_GRP_NM"										' Field��(1)
    
    arrHeader(0) = "���ű׷�"										' Header��(0)
    arrHeader(1) = "���ű׷��"									' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtpURgRP.focus
		Exit Function
	Else
		frm1.txtpURgRP.Value = arrRet(0)
		frm1.txtpURgRPNM.Value = arrRet(1)
		frm1.txtpURgRP.focus
	End If	
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
		Call DisplayMsgBox("971012","X", "����","X")
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
	arrParam(3) = frm1.txtDvryFromDt.Text
	arrParam(4) = frm1.txtDvryToDt.Text
	
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
'	Description : Tracking No. Popup���� Return�Ǵ� �� setting
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
 
    Call SetToolBar("11000000000011") 
    
	If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = parent.gPlant
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtBpCd.focus
	Else
		frm1.txtPlantCd.focus 
	End If
	
	Set gActiveElement = document.activeElement

End Sub

'================================================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'================================================================================================================================
Sub txtDvryFromDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtDvryFromDt.Action = 7 
		Call SetFocusToDocument("M")
		frm1.txtDvryFromDt.Focus
	End If 
End Sub

'================================================================================================================================
Sub txtDvryToDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtDvryToDt.Action = 7 
		Call SetFocusToDocument("M")
		frm1.txtDvryToDt.Focus
	End If 
End Sub

'================================================================================================================================
Sub txtDvryFromDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'================================================================================================================================
Sub txtDvryToDt_KeyDown(keycode, shift)
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

	If ValidDateCheck(frm1.txtDvryFromDt, frm1.txtDvryToDt) = False Then Exit Function
		
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
		Exit Function														'��: Query db data
	End If
	
    FncQuery = True															'��: Processing is OK
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
    On Error Resume Next													'��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    On Error Resume Next													'��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)									'��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)								'��: Protect system from crashing
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

'******************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  **************************
'	���� : 
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
						lgKeyStream = lgKeyStream & Trim(.hrdoAppflg.value)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(.hPoNo.value)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(.hPurGrp.value)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(.htrackingno.value)  & Parent.gColSep
					Else
						lgKeyStream = UCase(Trim(.txtPlantCd.value))  & Parent.gColSep
						lgKeyStream = lgKeyStream & UCase(Trim(.txtItemCd.value))  & Parent.gColSep
						lgKeyStream = lgKeyStream & UCase(Trim(.txtBPCd.value))  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(.txtDvryFromDt.Text)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(.txtDvryToDt.Text)  & Parent.gColSep
						
						If .rdoAppflg(0).checked = true Then
							lgKeyStream = lgKeyStream & "A" & Parent.gColSep
						ElseIf .rdoAppflg(1).checked = true Then
							lgKeyStream = lgKeyStream & "N" & Parent.gColSep
						Else
							lgKeyStream = lgKeyStream & "Y" & Parent.gColSep
						End If
						
						lgKeyStream = lgKeyStream & Trim(.txtPoNo.value)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(.txtPurGrp.value)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(.txtTrackingNo.value)  & Parent.gColSep
						
						.hPlantCd.value		= .txtPlantCd.value
						.hItemCd.value		= .txtItemCd.value
						.hBPCd.value		= .txtBPCd.value
						.hDvryFromDt.value	= .txtDvryFromDt.Text
						.hDvryToDt.value	= .txtDvryToDt.Text
						If .rdoAppflg(0).checked = true Then
							.hrdoAppflg.value = "A"
						ElseIf .rdoAppflg(1).checked = true Then
							.hrdoAppflg.value = "N"
						Else
							.hrdoAppflg.value = "Y"
						End If
						.hPoNo.value	= .txtPoNo.value
						.hPurGrp.value	= .txtPurGrp.value
						.htrackingno.value	= .txtTrackingNo.value
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
    strVal = strVal & "&lgStrPrevKey="  & lgStrPrevKey
    
    Call RunMyBizASP(MyBizASP, strVal)
    
    DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()

	Call SetToolBar("11000000000111")														'��: ��ư ���� ���� 
	
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
	End If

	lgIntFlgMode = parent.OPMD_UMODE														'��: Indicates that current mode is Update mode
    lgBlnFlgChgValue = False
	lgAfterQryFlg = True
	lgOldRow = 1
		
End Function

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Function Desc : �׸��� �����¸� �����Ѵ�.
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub 

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : �׸��带 ���� ���·� �����Ѵ�.
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


'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�԰����Ȳ</font></td>
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
			 						<TD CLASS=TD5 NOWRAP>����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12xxxU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14" ALT="�����"></TD>
			 						<TD CLASS=TD5 NOWRAP>����ó</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 tag="11xxxU" ALT="����ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBpCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=25 tag="14" ALT="����ó��"></TD>
								</TR>
			 					<TR>
									<TD CLASS=TD5 NOWRAP>ǰ��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>�԰���</TD> 
									<TD CLASS=TD6>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtDvryFromDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="������"></OBJECT>');</SCRIPT>
										&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtDvryToDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="������"></OBJECT>');</SCRIPT>
									</TD>
								</TR>
			 					<TR>								
									<TD CLASS="TD5" NOWRAP>���ֹ�ȣ</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPoNo" SIZE=18 MAXLENGTH=18 ALT="���ֹ�ȣ" tag="11xxxU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPoNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPoNo()"></TD>
									<TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio Class="Radio" ALT="����" NAME="rdoAppflg" id = "rdoAppflg1" Value="A" checked tag="11"><label for="rdoAppflg1">&nbsp;��ü&nbsp;</label>
														   <INPUT TYPE=radio Class="Radio" ALT="����" NAME="rdoAppflg" id = "rdoAppflg2" Value="N" tag="11"><label for="rdoAppflg2">&nbsp;����&nbsp;</label>
														   <INPUT TYPE=radio Class="Radio" ALT="����" NAME="rdoAppflg" id = "rdoAppflg3" Value="Y" tag="11"><label for="rdoAppflg3">&nbsp;��ǰ&nbsp;</label></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>���ű׷�</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPurGrp" SIZE=10 MAXLENGTH=10 tag="11xxxU" ALT="���ű׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPurGrp()">&nbsp;<INPUT TYPE=TEXT NAME="txtPurGrpNm" SIZE=25 tag="14" ALT="���ű׷��"></TD>
									<TD CLASS="TD5" NOWRAP>Tracking No.</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtTrackingNo" SIZE=25 MAXLENGTH=25 tag="11xxxU" ALT="Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNoBtn" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenTrackingInfo()"></TD>
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
<INPUT TYPE=HIDDEN NAME="hrdoAppflg" tag="24"><INPUT TYPE=HIDDEN NAME="hPoNo" tag="24"><INPUT TYPE=HIDDEN NAME="hPurGrp" tag="24">
<INPUT TYPE=HIDDEN NAME="hTrackingno" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>