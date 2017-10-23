Const BIZ_PGM_QRY_ID	= "p1401mb17_ko441.asp"											'��: �����Ͻ� ���� ASP�� 

Dim C_Level
Dim C_Seq
Dim C_ChildItemCd
Dim C_ChildItemNm
Dim C_Spec
Dim C_ChildItemUnit
Dim C_ItemAcctNm
Dim C_ProcTypeNm
Dim C_BomType
Dim C_ChildItemBaseQty
Dim C_ChildBasicUnit
Dim C_PrntItemBaseQty
Dim C_PrntBasicUnit
Dim C_SafetyLT
Dim C_LossRate
Dim C_SupplyFlgNm
Dim C_ValidFromDt
Dim C_ValidToDt
Dim	C_ECNNo
Dim	C_ECNDescription
Dim	C_ECNReasonCd
Dim	C_DrawingPath
Dim C_Remark
Dim C_HdrItemCd
Dim C_HdrBomNo
Dim C_Row	

Dim IsOpenPop

' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	C_Level					= 1
	C_Seq					= 2
	C_ChildItemCd			= 3
	C_ChildItemNm			= 4
	C_Spec					= 5
	C_ChildItemUnit			= 6
	C_ItemAcctNm			= 7
	C_ProcTypeNm			= 8
	C_BomType				= 9
	C_ChildItemBaseQty		= 10
	C_ChildBasicUnit		= 11
	C_PrntItemBaseQty		= 12
	C_PrntBasicUnit			= 13
	C_SafetyLT				= 14
	C_LossRate				= 15
	C_SupplyFlgNm			= 16
	C_ValidFromDt			= 17
	C_ValidToDt				= 18
	C_ECNNo					= 19	
	C_ECNDescription		= 20
	C_ECNReasonCd			= 21
	C_DrawingPath			= 22
	C_Remark				= 23
	C_HdrItemCd				= 24
	C_HdrBomNo				= 25
	C_Row					= 26
End Sub

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKeyIndex = ""
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1                                       '��: initializes sort direction
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtBaseDt.Text = StartDate
	frm1.txtBomNo.value = "1"
	frm1.cboItemAcct.value = ""
	If lgPLCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPlantCd, "Q") 
		frm1.txtPlantCd.Tag = left(frm1.txtPlantCd.Tag,1) & "4" & mid(frm1.txtPlantCd.Tag,3,len(frm1.txtPlantCd.Tag))
        	frm1.txtPlantCd.value = lgPLCd
	End If
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
	
	Call initSpreadPosVariables()

	With frm1.vspdData
	
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20030109",, parent.gAllowDragDropSpread

		.ReDraw = False

		.MaxCols = C_Row												'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.MaxRows = 0

		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit 	C_Level, 				"����"		,	8
		ggoSpread.SSSetFloat	C_Seq,					"����"		,	6,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,1,FALSE,"Z" 
		ggoSpread.SSSetEdit		C_ChildItemCd,			"��ǰ��"	,	20,,,18,2
		ggoSpread.SSSetEdit 	C_ChildItemNm, 			"��ǰ���"	,	30
		ggoSpread.SSSetEdit 	C_Spec,	 				"�԰�"		,	30
		ggoSpread.SSSetEdit		C_ChildItemUnit,		"����"		,	6,,,3,2
		ggoSpread.SSSetEdit		C_ItemAcctNm,			"ǰ�����"	,	10
		ggoSpread.SSSetEdit 	C_ProcTypeNm, 			"���ޱ���"	,	12
		ggoSpread.SSSetEdit		C_BomType,				"BOM Type"	,	10,,,1,2
		ggoSpread.SSSetFloat	C_ChildItemBaseQty,		"��ǰ����ؼ�",	15, "8",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit 	C_ChildBasicUnit,		"����"		,	6,,,3,2
		ggoSpread.SSSetFloat	C_PrntItemBaseQty,		"��ǰ����ؼ�", 15, "8",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit		C_PrntBasicUnit,		"����"		,	6,,,3,2
		ggoSpread.SSSetFloat	C_SafetyLT, 			"����L/T"	,	10,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,1,FALSE,"Z" 
		ggoSpread.SSSetFloat	C_LossRate,				"Loss��"	,	10,"7",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,1,FALSE,"Z" 
		ggoSpread.SSSetEdit		C_SupplyFlgNm,			"�����󱸺�",	8
		ggoSpread.SSSetDate		C_ValidFromDt,			"������"	,	11, 2, parent.gDateFormat
		ggoSpread.SSSetDate 	C_ValidToDt,			"������"	,	11, 2, parent.gDateFormat
		ggoSpread.SSSetEdit		C_ECNNo,				"���躯���ȣ", 18
		ggoSpread.SSSetEdit		C_ECNDescription,		"���躯�泻��", 30
		ggoSpread.SSSetEdit		C_ECNReasonCd,			"���躯��ٰ�", 10
		ggoSpread.SSSetEdit		C_DrawingPath,			"������"	,	30
		ggoSpread.SSSetEdit 	C_Remark,	 			"���"		,	30,,, 1000
		ggoSpread.SSSetEdit		C_HdrItemCd,			"Headerǰ��",	5
		ggoSpread.SSSetEdit		C_HdrBomNo,				"header BOM No.", 5
		ggoSpread.SSSetEdit		C_Row,					"����", 5

		ggoSpread.SSSetSplit2(3)											'frozen ��� �߰� 

		Call ggoSpread.MakePairsColumn(C_Level, C_ChildItemCd)
		Call ggoSpread.MakePairsColumn(C_ChildItemBaseQty, C_ChildBasicUnit)
		Call ggoSpread.MakePairsColumn(C_PrntItemBaseQty, C_PrntBasicUnit)

		Call ggoSpread.SSSetColHidden(C_ChildItemUnit, C_ChildItemUnit, True)
		Call ggoSpread.SSSetColHidden(C_HdrItemCd, C_HdrBomNo, True)
		Call ggoSpread.SSSetColHidden(C_Row, C_Row, True)
    
		.ReDraw = True

		Call SetSpreadLock 

    End With
    
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_Level					= iCurColumnPos(1)
			C_Seq					= iCurColumnPos(2)
			C_ChildItemCd			= iCurColumnPos(3)
			C_ChildItemNm			= iCurColumnPos(4)
			C_Spec					= iCurColumnPos(5)
			C_ChildItemUnit			= iCurColumnPos(6)
			C_ItemAcctNm			= iCurColumnPos(7)
			C_ProcTypeNm			= iCurColumnPos(8)
			C_BomType				= iCurColumnPos(9)
			C_ChildItemBaseQty		= iCurColumnPos(10)
			C_ChildBasicUnit		= iCurColumnPos(11)
			C_PrntItemBaseQty		= iCurColumnPos(12)
			C_PrntBasicUnit			= iCurColumnPos(13)
			C_SafetyLT				= iCurColumnPos(14)
			C_LossRate				= iCurColumnPos(15)
			C_SupplyFlgNm			= iCurColumnPos(16)
			C_ValidFromDt			= iCurColumnPos(17)
			C_ValidToDt				= iCurColumnPos(18)
			C_ECNNo					= iCurColumnPos(19)
			C_ECNDescription		= iCurColumnPos(20)
			C_ECNReasonCd			= iCurColumnPos(21)
			C_DrawingPath			= iCurColumnPos(22)
			C_Remark				= iCurColumnPos(23)
			C_HdrItemCd				= iCurColumnPos(24)
			C_HdrBomNo				= iCurColumnPos(25)
			C_Row					= iCurColumnPos(26)
    End Select    
End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'=======================================================================================================
'   Event Name : txtBaseDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtBaseDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtBaseDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtBaseDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtBaseDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtBaseDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'------------------------------------------  OpenCondPlant()  -------------------------------------------------
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
    If frm1.txtPlantCd.className = "protected" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"				' �˾� ��Ī 
	arrParam(1) = "B_PLANT"						' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "����"					' TextBox ��Ī 
	
    arrField(0) = "PLANT_CD"					' Field��(0)
    arrField(1) = "PLANT_NM"					' Field��(1)
    
    arrHeader(0) = "����"					' Header��(0)
    arrHeader(1) = "�����"					' Header��(1)
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenIremCd()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemCd(ByVal str, ByVal iPos)
	Dim arrRet
	Dim arrParam(5), arrField(11)
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "����", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)   ' Plant Code
	arrParam(1) = Trim(str)	' Item Code
	
	arrParam(2) = ""												' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrField(0) = 1							'ITEM_CD
    arrField(1) = 2 						'ITEM_NM											
    
	iCalledAspName = AskPRAspName("B1B11PA4")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA4", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemCd(arrRet, iPos)
	End If	
	  
	If iPos = 0 Then
		Call SetFocusToDocument("M")
		frm1.txtItemCd.focus			
	Else
		Call SetActiveCell(frm1.vspdData,C_ChildItemCd,frm1.vspdData.ActiveRow,"M","X","X")
		Set gActiveElement = document.activeElement
	End IF
	
End Function

'------------------------------------------  OpenBomNo()  -------------------------------------------------
'	Name : OpenBomNo()
'	Description : Condition BomNo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBomNo(ByVal strItem, ByVal strBom)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then
		IsOpenPop = False
		Exit Function
	End If
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "����", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	If strItem = "" Then
		Call DisplayMsgBox("971012", "X", "ǰ��", "X")
		
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	'---------------------------------------------
	 ' Parameter Setting
	 '--------------------------------------------- 

	IsOpenPop = True

	arrParam(0) = "BOM�˾�"						' �˾� ��Ī 
	arrParam(1) = "B_MINOR"							' TABLE ��Ī 
	
	arrParam(2) = Trim(frm1.txtBomNo.value)		' Code Condition
	
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = "MAJOR_CD = " & FilterVar("P1401", "''", "S") & " "
	
	arrParam(5) = "BOM Type"						' TextBox ��Ī 
	
    arrField(0) = "MINOR_CD"						' Field��(0)
    arrField(1) = "MINOR_NM"						' Field��(1)
        
    arrHeader(0) = "BOM Type"					' Header��(0)
    arrHeader(1) = "BOM Ư��"					' Header��(1)
    
	arrRet = window.showModalDialog("../../ComAsp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetBomNo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtBomNo.focus
	
End Function

'------------------------------------------  SetConPlant()  --------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetConPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)	
End Function

'------------------------------------------  SetItemCd()  --------------------------------------------------
'	Name : SetItemCd()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetItemCd(byval arrRet, ByVal iPos)
	
	If iPos = 0 Then
		frm1.txtItemCd.Value    = arrRet(0)		
		frm1.txtItemNm.Value    = arrRet(1)
		
	Else 
		With frm1.vspdData
			.Col = C_ChildItemCd
			.Row = .ActiveRow
			.Text = arrRet(0)		
			
			Call LookUpItemByPlant(arrRet(0),.Row)

		End With
		
	End IF
	
End Function

'------------------------------------------  SetBomNo()  --------------------------------------------------
'	Name : SetBomNo()
'	Description : Bom No Popup���� return�� �� 
'--------------------------------------------------------------------------------------------------------- 
Function SetBomNo(byval arrRet)
	frm1.txtBomNo.Value    = arrRet(0)		
	frm1.txtBomNm.Value    = arrRet(1)		
End Function

'==========================================================================================
'   Function Name :SetFieldProp
'   Function Desc :���� Case�� ���� Field���� �Ӽ��� �����Ѵ�.
'==========================================================================================
Function SetFieldProp(ByVal lRow, ByVal Level, ByVal ProcType)
	
End Function

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )
    
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
    gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
    Call SetPopupMenuItemInf("0000110111")

	If Row <= 0 Or Col < 0 Then
		Exit Sub
	End If
	
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell�� ����� �����ǹ߻� �̺�Ʈ 
'==========================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
    With frm1.vspdData 

    If Row >= NewRow Then
        Exit Sub
    End If

	'----------  Coding part  -------------------------------------------------------------   

    End With

End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    '----------  Coding part  -------------------------------------------------------------   
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then
    
		If lgStrPrevKeyIndex <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			Call DisableToolBar(parent.TBC_QUERY)  
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
    
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    
    '-----------------------
    'Erase contents area
    '-----------------------
    If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End If
		
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = ""
	End If
		
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    Call InitVariables															'��: Initializes local global variables
    																			
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then   
		Exit Function           
    End If     																		'��: Query db data

    FncQuery = True																'��: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	If frm1.vspdData.maxrows < 1 Then Exit Function
	frm1.vspdData.Redraw = False
    ggoSpread.EditUndo                                                  '��: Protect system from crashing
	frm1.vspdData.Redraw = True
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
   Call parent.fncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLEMULTI)							'��: ȭ�� ���� 
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI, False)	                   '��:ȭ�� ����, Tab ���� 
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function FncSplitColumn()
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
    
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    FncExit = True
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 

    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim StrNextKey      

    DbQuery = False
    
    LayerShowHide(1)
		
    Err.Clear                                                               '��: Protect system from crashing

	Dim strVal

    With frm1
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001						'��: 
		
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtBomNo=" & Trim(.txtBomNo.value) 
		strVal = strVal & "&txtBaseDt=" & Trim(.txtBaseDt.Text)
		
		If frm1.rdoSrchType1.checked = True Then
			strVal = strval & "&rdoSrchType=" & frm1.rdoSrchType1.value 
		ElseIf frm1.rdoSrchType2.checked = True Then
			strVal = strval & "&rdoSrchType=" & frm1.rdoSrchType2.value 
		End If       
		
		strVal = strVal & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex         '��: Next key tag

		Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 

    End With
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk(LngMaxRow)										'��: ��ȸ ������ ������� 
	Call SetToolbar("11000000000111")								'��: ��ư ���� ���� 
    If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
    End If
    lgIntFlgMode = parent.OPMD_UMODE												'��: Indicates that current mode is Update mode
    
End Function


'=====================================================
'NOTE. 20080407 ��/������ �ɼǿ� ���� Label���� >>air
'=====================================================
Sub OnrdoSrchType(optValue)
	With frm1.vspdData
		.Row = 0
		.Col = C_ChildItemCd
	
		If optValue = 1 Then
			.text = "��ǰ��"
		Else
			.text = "��ǰ��"
		End If

		.Col = C_ChildItemNm
	
		If optValue = 1 Then
			.text = "��ǰ���"
		Else
			.text = "��ǰ���"
		End If		
		 
	End With
End Sub