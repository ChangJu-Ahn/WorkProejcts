<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1205ma2.asp
'*  4. Program Name         : �ڿ�����������ȸ 
'*  5. Program Desc         :
'*  6. Component List       : 
'*  7. Modified date(First) : 2003/04/09
'*  8. Modified date(Last)  :  
'*  9. Modifier (First)     : Ryu Sung Won
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

Const BIZ_PGM_QRY_ID = "p1205mb9.asp"

Const C_SHEETMAXROWS = 100

Dim C_ItemCd
Dim C_ItemNm
Dim C_Spec
Dim C_RoutNo
Dim C_RoutNm
Dim C_OprNo
Dim C_Rank
Dim C_ResourceCd
Dim C_ResourceNm
Dim C_ResourceType
Dim C_Efficiency


<!-- #Include file="../../inc/lgVariables.inc" -->	

Dim lgBlnFlgConChg				'��: Condition ���� Flag
Dim lgOldRow
Dim iDBSYSDate

Dim IsOpenPop
'Dim lgStrPrevKey
Dim lgStrNextKey1	'item_cd
Dim lgStrNextKey2	'rout_no
Dim lgStrNextKey3	'opr_no
Dim lgStrNextKey4	'resourcd_cd
Dim lgStrNextKey5	'rank

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	C_ItemCd		= 1
	C_ItemNm		= 2	
	C_Spec			= 3
	C_RoutNo		= 4
	C_RoutNm		= 5
	C_OprNo			= 6
	C_Rank			= 7
	C_ResourceCd	= 8
	C_ResourceNm	= 9
	C_ResourceType	= 10
	C_Efficiency	= 11
End Sub

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'==================================================================================================== 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE			
    lgBlnFlgChgValue = False			
    lgIntGrpCount = 0					

    IsOpenPop = False												
	'lgStrPrevKey = ""
	lgStrNextKey1 = ""
	lgStrNextKey2 = ""
	lgStrNextKey3 = ""
	lgStrNextKey4 = ""
	lgStrNextKey5 = ""
	
	lgSortKey = 1
	lgOldRow = 0
End Sub

'=======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<%Call loadInfTB19029A("Q", "P", "NOCOOKIE", "MA")%>
End Sub

'========================= 2.2.3 InitSpreadSheet() ======================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()    
	Call AppendNumberPlace("6","6","0")
	Call AppendNumberPlace("7","3","2")
		
	With frm1.vspdData

		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021126",, parent.gAllowDragDropSpread

		.ReDraw = False

		.MaxCols = C_Efficiency + 1
		.MaxRows = 0

		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit		C_ItemCd,		"ǰ��", 15,,, 18, 2
		ggoSpread.SSSetEdit		C_ItemNm,		"ǰ���", 25,,, 40
		ggoSpread.SSSetEdit		C_Spec,			"�԰�", 25,,,40
		ggoSpread.SSSetEdit		C_RoutNo,		"�����", 10
		ggoSpread.SSSetEdit		C_RoutNm,		"����ø�", 14
		ggoSpread.SSSetEdit		C_OprNo,		"����", 10
		ggoSpread.SSSetFloat	C_Rank,			"����", 10, "6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit		C_ResourceCd,	"�ڿ�", 15
		ggoSpread.SSSetEdit		C_ResourceNm,	"�ڿ���", 25
		ggoSpread.SSSetEdit		C_ResourceType,	"�ڿ�����", 15
		ggoSpread.SSSetFloat	C_Efficiency,	"ȿ��", 10, "7",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"

		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_Rank, C_Rank, True)
		
		ggoSpread.SSSetSplit2(1)										'frozen ����߰� 

		.ReDraw = True

		Call SetSpreadLock 

    End With
    
End Sub

'============================== 2.2.4 SetSpreadLock() ===================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    With frm1
		ggoSpread.Source = .vspdData
		.vspdData.ReDraw = False
		ggoSpread.SpreadLock -1, -1
		.vspdData.ReDraw = True
	End With
End Sub

'============================ 2.2.5 SetSpreadColor() ====================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
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
			C_ItemCd		= iCurColumnPos(1)
			C_ItemNm		= iCurColumnPos(2)	
			C_Spec			= iCurColumnPos(3)
			C_RoutNo		= iCurColumnPos(4)
			C_RoutNm		= iCurColumnPos(5)
			C_OprNo			= iCurColumnPos(6)
			C_Rank			= iCurColumnPos(7)
			C_ResourceCd	= iCurColumnPos(8)
			C_ResourceNm	= iCurColumnPos(9)
			C_ResourceType	= iCurColumnPos(10)
			C_Efficiency	= iCurColumnPos(11)
			
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

'========================================  2.2.1 SetDefaultVal()  ==================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'===================================================================================================
Sub SetDefaultVal()
	'iDBSYSDate = "<%=GetSvrDate%>"
	'frm1.txtValidDt.text = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
End Sub

Sub InitComboBox()

End Sub

'------------------------------------------  OpenConItemInfo()  -------------------------------------------------
'	Name : OpenConItemInfo()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConItemInfo()
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Or UCase(frm1.txtItemCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X" , "����", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)	' Item Code
	arrParam(2) = "12!MO"							' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field��(0) :"ITEM_CD"
    arrField(1) = 2 							' Field��(1) :"ITEM_NM"
    
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemInfo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus

End Function

'------------------------------------------  OpenConPlant()  -----------------------------------------
'	Name : OpenConPlant()
'	Description : Plant PopUp
'-------------------------------------------------------------------------------------------------------- 
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

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
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

Function OpenConRouting()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Or UCase(frm1.txtRoutNo.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "����", "X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	If frm1.txtItemCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "ǰ��", "X")
		frm1.txtItemCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = "������˾�"	
	arrParam(1) = "P_ROUTING_HEADER"				
	arrParam(2) = Trim(frm1.txtRoutNo.Value)
	arrParam(3) = ""
	arrParam(4) =  "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " And ITEM_CD = " & FilterVar(frm1.txtItemCd.value, "''", "S")
	arrParam(5) = "�����"			

    arrField(0) = "ROUT_NO"	
    arrField(1) = "DESCRIPTION"	
    arrField(2) = "BOM_NO"
    arrField(3) = "MAJOR_FLG"

    arrHeader(0) = "�����"		
    arrHeader(1) = "����ø�"		
    arrHeader(2) = "BOM Type"
    arrHeader(3) = "�ֶ����"
    
    	
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetRouting(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtRoutNo.focus
	
End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'=========================================================================================================
'------------------------------------------  SetItemInfo()  --------------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemInfo(Byval arrRet)
	frm1.txtItemCd.value = arrRet(0)
	frm1.txtItemNm.value = arrRet(1)
End Function

'------------------------------------------  SetPlant()  --------------------------------------------------
'	Name : SetPlant()
'	Description : Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)
	frm1.txtPlantNm.Value    = arrRet(1)
End Function
'------------------------------------------  SetRouting()  --------------------------------------------------
'	Name : SetPlant()
'	Description : Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetRouting(byval arrRet)
	frm1.txtRoutNo.Value    = arrRet(0)
	frm1.txtRoutNm.Value    = arrRet(1)
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)    

    Call ggoOper.LockField(Document, "N")											'��: Lock  Suitable  Field

	Call SetDefaultVal
   	Call InitComboBox
    Call InitVariables		
    Call InitSpreadSheet	
	Call SetToolbar("11000000000011")

	If parent.gPlant <> "" And frm1.txtPlantCd.value = "" Then
		frm1.txtPlantCd.value = parent.gPlant
		frm1.txtPlantNm.value = parent.gPlantNm
		
		frm1.txtItemCd.focus 
	Else
		frm1.txtPlantCd.focus 
	End If
	
	Set gActiveElement = document.activeElement 

End Sub

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
	Dim IntRetCD
	
	gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData
	Call SetPopupMenuItemInf("0000111111")    
	
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
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
    
	If Row <= 0 Or Col < 0 Then
		ggoSpread.Source = frm1.vspdData
		Exit Sub
	End If
	
	frm1.vspdData.Row = Row
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'=======================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 1.����ð�(runtime)�� �˾��޴��� ���ؼ� �������� �ٲ���.
'				 2.Mouse�� Ư��Cell�� ����("SPC")�ϰ� ������ ��ư("SPCR")�� ������ �˾��� ���δ�.
'				   �˾����� Ư�� �޴� item�� ����("SPCRP") ���� Į���� freeze�Ѵ�.
'=======================================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
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
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then
		If lgStrNextKey1 <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			Call DisableToolBar(parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If     
    End if
    
End Sub

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell�� ����� �����ǹ߻� �̺�Ʈ 
'==========================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
        
    FncQuery = False                                                       
    
    Err.Clear                                                              

    Call ggoOper.ClearField(Document, "2")		
    Call InitVariables							
        
    If Not chkField(Document, "1") Then								
       Exit Function
    End If
    
    If DbQuery = False Then
		Exit Function
    End If													
       
    FncQuery = True													
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
    On Error Resume Next                                                    '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)                                                   <%'��: Protect system from crashing%>
    'Call parent.FncExport(parent.C_SINGLE)											
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)     
    'Call parent.FncFind(parent.C_SINGLE, False)                                         '��:ȭ�� ����, Tab ���� 
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
'Function FncSplitColumn()
'    
'    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
'       Exit Function
'    End If
'
'    ggoSpread.Source = gActiveSpdSheet
'    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
'    
'End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
    FncExit = True
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	
	Dim strcboStatus, strcboEBomFlg, strcboMBomFlg
	
	Err.Clear															

	DbQuery = False														

	LayerShowHide(1)
		
	Dim strVal
	
	If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode="	& parent.UID_M0001
		strVal = strVal & "&txtPlantCd="		& UCase(Trim(frm1.hPlantCd.value))
		strVal = strVal & "&txtItemCd="			& UCase(Trim(frm1.hItemCd.value))
		strVal = strVal & "&txtRoutNo="			& UCase(Trim(frm1.hRoutNo.value))
		strVal = strVal & "&lgIntFlgMode="		& lgIntFlgMode
		strVal = strVal & "&lgStrNextKey1="		& lgStrNextKey1		'item_cd
		strVal = strVal & "&lgStrNextKey2="		& lgStrNextKey2		'rout_no
		strVal = strVal & "&lgStrNextKey3="		& lgStrNextKey3		'opr_no
		strVal = strVal & "&lgStrNextKey4="		& lgStrNextKey4		'resourcd_cd
		strVal = strVal & "&lgStrNextKey5="		& lgStrNextKey5		'rank
		strVal = strVal & "&txtMaxRows="		& frm1.vspdData.MaxRows
	Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode="	& parent.UID_M0001
		strVal = strVal & "&txtPlantCd="		& UCase(Trim(frm1.txtPlantCd.value))
		strVal = strVal & "&txtItemCd="			& UCase(Trim(frm1.txtItemCd.value))
		strVal = strVal & "&txtRoutNo="			& UCase(Trim(frm1.txtRoutNo.value))
		strVal = strVal & "&lgIntFlgMode="		& lgIntFlgMode
		strVal = strVal & "&lgStrNextKey1="		& ""				'item_cd
		strVal = strVal & "&lgStrNextKey2="		& ""				'rout_no
		strVal = strVal & "&lgStrNextKey3="		& ""				'opr_no
		strVal = strVal & "&lgStrNextKey4="		& ""				'resourcd_cd
		strVal = strVal & "&lgStrNextKey5="		& ""				'rank
		strVal = strVal & "&txtMaxRows="		& frm1.vspdData.MaxRows
	
	End If  

	Call RunMyBizASP(MyBizASP, strVal)									

	DbQuery = True																					
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()
	
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
	End If													
	lgIntFlgMode = parent.OPMD_UMODE											

	Call ggoOper.LockField(Document, "Q")								
	Call SetToolbar("11000000000111")
	
	frm1.vspddata.Focus
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�ڿ�����������ȸ</font></td>
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
						<FIELDSET CLASS="CLSFLD">
							
							<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
			 					<TR>
			 						<TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12XXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=30 tag="14"></TD>
			 						<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD6" NOWRAP>&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>ǰ��</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=20 MAXLENGTH=18 tag="11XXXU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConItemInfo()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>�����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtRoutNo" SIZE=15 MAXLENGTH=7 tag="11XXXU" ALT="�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRoutNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConRouting()">&nbsp;<INPUT TYPE=TEXT NAME="txtRoutNm" SIZE=20 tag="14"></TD>
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
								<TD HEIGHT=* WIDTH=100%>
									<script language =javascript src='./js/p1205ma2_vspdData_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hItemCd" tag="24"><INPUT TYPE=HIDDEN NAME="hRoutNo" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
