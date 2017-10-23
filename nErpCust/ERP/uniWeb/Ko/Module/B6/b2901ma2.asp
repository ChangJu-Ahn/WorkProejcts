<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : 
'*  3. Program ID           : B2901ma2
'*  4. Program Name         : ȸ������Match
'*  5. Program Desc         : ��ϵ� ȸ��μ��� ���� �����,�����,Cost Center�� Matching�Ѵ�.
'*  6. Component List       : PB6SA20
'*  7. Modified date(First) : 2000/09/04
'*  8. Modified date(Last)  : 2000/09/04
'*  9. Modifier (First)     : Kwon Yong Gyoun
'* 10. Modifier (Last)      : Kwon Yong Gyoun / Cho Ig Sung
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*********************************************************************************************** -->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit

<!-- #Include file="../../inc/lgvariables.inc" -->	
'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_ID = "B2901mb2.asp"			 '��: �����Ͻ� ���� ASP�� 


'==========================================  1.2.1 Global ��� ����  ======================================
'��: Grid Columns
Dim C_DeptCd
Dim C_DeptNm
Dim C_OrgChgDt
Dim C_DeptFullNm
Dim C_DeptEngNm
Dim C_ParDeptCd
Dim C_ParDeptNm
Dim C_CostCd
Dim C_CostPopUp
Dim C_CostNm
Dim C_BizUnitCd
Dim C_BizUnitPopUp
Dim C_BizUnitNm
Dim C_Level
Dim C_Seq
Dim C_EndDeptFg
Dim C_InternalCd
Dim C_OrgChgId
Dim C_OldInternalCd
Dim C_ENTRY_FG

'========================================================================================================= 

'----------------  ���� Global ������ ����  -------------------------------------------------------------- 
Dim IsOpenPop
'Dim  lgSortKey


'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
    lgPageNo  = 0

End Sub

Sub initSpreadPosVariables()
     C_DeptCd			= 1
	 C_DeptNm			= 2
	 C_OrgChgDt			= 3
	 C_DeptFullNm		= 4
	 C_DeptEngNm		= 5
	 C_ParDeptCd		= 6
	 C_ParDeptNm		= 7
	 C_CostCd			= 8
	 C_CostPopUp		= 9
	 C_CostNm			= 10
	 C_BizUnitCd		= 11
	 C_BizUnitPopUp		= 12
	 C_BizUnitNm		= 13
	 C_Level			= 14
	 C_Seq				= 15
	 C_EndDeptFg		= 16
	 C_InternalCd		= 17
	 C_OrgChgId			= 18
	 C_OldInternalCd	= 19
	 C_ENTRY_FG         = 20
End Sub

'========================================================================================================= 
Sub SetDefaultVal()

End Sub


'========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE", "MA") %>
End Sub

'========================================================================================
Sub InitSpreadSheet()

    Call initSpreadPosVariables()
	ggoSpread.Source= frm1.vspdData
	ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    

	With frm1.vspdData

		.MaxCols = C_ENTRY_FG + 1
		.MaxRows = 0

		.ReDraw = False

		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit   C_DeptCd,			"�μ��ڵ�",				10,  , , 10, 2
		ggoSpread.SSSetEdit   C_DeptNm,			"�μ���",				20,  , , 40
		ggoSpread.SSSetDate   C_OrgChgDt,	    "����������",			10, 2, Parent.gDateFormat
		ggoSpread.SSSetEdit   C_DeptFullNm,		"�μ��幮��",			20,  , , 200
		ggoSpread.SSSetEdit   C_DeptEngNm,		"�μ�������",			20,  , , 100
		ggoSpread.SSSetEdit   C_ParDeptCd,		"�����μ��ڵ�",			10,  , , 10, 2
		ggoSpread.SSSetEdit   C_ParDeptNm,		"�����μ���",			20,  , , 40
		ggoSpread.SSSetEdit   C_CostCd,			"�ڽ�Ʈ��Ÿ�ڵ�",		10,  , , 10, 2
		ggoSpread.SSSetButton C_CostPopUp
		ggoSpread.SSSetEdit   C_CostNm,			"�ڽ�Ʈ��Ÿ��",			20,  , , 20
		ggoSpread.SSSetEdit   C_BizUnitCd,		"������ڵ�",			10,  , , 10, 2
		ggoSpread.SSSetButton C_BizUnitPopUp
		ggoSpread.SSSetEdit   C_BizUnitNm,		"����θ�",				20,  , , 30
		Call AppendNumberPlace("6","3","0")
		ggoSpread.SSSetFloat  C_Level,			"LEVEL",				8,"6"  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"P"
		ggoSpread.SSSetFloat  C_Seq,			"����",					8,"6"  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"P"
		ggoSpread.SSSetEdit   C_EndDeptFg,		"����������",			8, 2, , 1
		ggoSpread.SSSetEdit   C_InternalCd,		"���κμ��ڵ�",			12,  , , 30, 2
		ggoSpread.SSSetEdit   C_OrgChgId,		    "",			12,  , , 30, 2
		ggoSpread.SSSetEdit   C_OldInternalCd,	"�������κμ��ڵ�",		12,  , , 30, 2
		ggoSpread.SSSetEdit   C_ENTRY_FG  , "", 4

		.ReDraw = True

		Call ggoSpread.MakePairsColumn(C_DeptCd,C_DeptNm)
		Call ggoSpread.MakePairsColumn(C_ParDeptCd,C_ParDeptNm)
		Call ggoSpread.MakePairsColumn(C_CostCd,C_CostNm)
		Call ggoSpread.MakePairsColumn(C_BizUnitCd,C_BizUnitNm)
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_OrgChgId,C_OrgChgId,True)
		Call ggoSpread.SSSetColHidden(C_ENTRY_FG,C_ENTRY_FG,True)		
		
		Call SetSpreadLock
    End With
End Sub


'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'===========================================================================================================
Sub SetSpreadLock()
	Dim ii
    
    With frm1
		.vspdData.ReDraw = False
        
        ggoSpread.Source = .vspdData
        
		ggoSpread.SpreadLock C_DeptCd			, -1, C_DeptCd			, -1
		ggoSpread.SpreadLock C_DeptNM			, -1, C_DeptNM			, -1
		ggoSpread.SpreadLock C_OrgChgDt			, -1, C_OrgChgDt		, -1
		ggoSpread.SpreadLock C_DeptFullNm		, -1, C_DeptFullNm		, -1
		ggoSpread.SpreadLock C_DeptEngNm		, -1, C_DeptEngNm		, -1
		ggoSpread.SpreadLock C_ParDeptCd		, -1, C_ParDeptCd		, -1
		ggoSpread.SpreadLock C_ParDeptNm		, -1, C_ParDeptNm		, -1
		ggoSpread.SSSetRequired C_CostCd		, -1, -1
		ggoSpread.SpreadLock C_CostNm			, -1, C_CostNm			, -1
		ggoSpread.SSSetRequired C_BizUnitCd		, -1, -1
		ggoSpread.SpreadLock C_BizUnitNm		, -1, C_BizUnitNm		, -1
		ggoSpread.SpreadLock C_Level			, -1, C_Level			, -1
		ggoSpread.SpreadLock C_Seq				, -1, C_Seq				, -1
		ggoSpread.SpreadLock C_EndDeptFg		, -1, C_EndDeptFg		, -1
		ggoSpread.SpreadLock C_InternalCd		, -1, C_InternalCd		, -1
		ggoSpread.SpreadLock C_OldInternalCd	, -1, C_OldInternalCd	, -1
		.vspdData.ReDraw = True
		
		For ii = 1 To .vspdData.MaxRows
			.vspddata.col = C_ENTRY_FG
			.vspddata.row = ii
			
			If Trim(.vspddata.text) = "E" Then			
				ggoSpread.SpreadLock C_DeptCd, ii, C_ENTRY_FG ,ii
			End If
		Next

    End With

End Sub


'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'============================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)

End Sub

'------------------------------------------  OpenPopUp()  -------------------------------------------------
'	Name : OpenPopUp()
'	Description : PopUp
'--------------------------------------------------------------------------------------------------------- 
Function HorgAbsPopUp(Byval strCode)
	Dim arrRet
	Dim	arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "��������ID �˾�"			' �˾� ��Ī 
	arrParam(1) = "HORG_ABS A"	 				' TABLE ��Ī 
	arrParam(2) = strCode							' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""
	arrParam(5) = "��������ID"								' �����ʵ��� �� ��Ī 

    arrField(0) = "A.ORGID"						' Field��(0)
    arrField(1) = "A.ORGNM"						' Field��(1)
    arrField(2) = "A.ORGDT"
    arrField(3) = "A.CURRENTYN"

    arrHeader(0) = "��������ID"				' Header��(0)
    arrHeader(1) = "���������"				' Header��(1)
    arrHeader(2) = "����������"
    arrHeader(3) = "�������"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtOrgChgID.focus
		Exit Function
	Else
		Call SetHorgAbsPopUp(arrRet)
	End If	

End Function

'========================================================================================================= 
Function OpenAcctDeptPopUp(Byval strCode)
	Dim arrRet
	Dim	arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�μ� �˾�"				' �˾� ��Ī 
	arrParam(1) = "B_ACCT_DEPT A" 				' TABLE ��Ī 
	arrParam(2) = strCode						' Code Condition
	arrParam(3) = ""							' Name Cindition

	if Trim(frm1.txtOrgChgID.value) = "" Then
'	    Call ggoOper.ClearField(Document, "1")
	Else
		arrParam(4) = "A.ORG_CHANGE_ID =  " & FilterVar(frm1.txtOrgChgID.value, "''", "S") & " "	'Where Condition
	End If

	arrParam(5) = "�μ�"								' �����ʵ��� �� ��Ī 

    arrField(0) = "A.DEPT_CD"					' Field��(0)
    arrField(1) = "A.DEPT_NM"					' Field��(1)
    arrField(2) = "A.ORG_CHANGE_ID"
    
    arrHeader(0) = "�μ��ڵ�"				' Header��(0)
    arrHeader(1) = "�μ���"					' Header��(1)
    arrHeader(2) = "��������ID"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtDeptCd.focus
		Exit Function
	Else
		Call SetAcctDeptPopUp(arrRet)
	End If

End Function

'========================================================================================================= 
Function OpenBizUnitPopUp(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "����� �˾�"						' �˾� ��Ī 
	arrParam(1) = "B_Biz_Unit"							' TABLE ��Ī 
	arrParam(2) = strCode    							' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = ""
	arrParam(5) = "�����"

    arrField(0) = "BIZ_UNIT_CD"							' Field��(0)
    arrField(1) = "BIZ_UNIT_NM"							' Field��(1)

    arrHeader(0) = "������ڵ�"						' Header��(0)
    arrHeader(1) = "����θ�"					    ' Header��(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetBizUnitPopUp(arrRet)
	End If	

End Function

'========================================================================================================= 
Function OpenCostPopUp(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�ڽ�Ʈ��Ÿ �˾�"					' �˾� ��Ī 
	arrParam(1) = "B_COST_CENTER"						' TABLE ��Ī 
	arrParam(2) = strCode    							' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = ""
	arrParam(5) = "�ڽ�Ʈ��Ÿ"

    arrField(0) = "COST_CD"								' Field��(0)
    arrField(1) = "COST_NM"								' Field��(1)

    arrHeader(0) = "�ڽ�Ʈ��Ÿ�ڵ�"					' Header��(0)
    arrHeader(1) = "�ڽ�Ʈ��Ÿ��"				    ' Header��(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCostCenterPopUp(arrRet)
	End If	

End Function


'---------------------------------------------------------------------------------------------------------- 
Function SetHorgAbsPopUp(Byval arrRet)
	With frm1
		.txtOrgChgID.focus
		.txtOrgChgID.value = Trim(arrRet(0))
		.txtOrgChgNm.value = arrRet(1)
	End With
End Function

Function SetAcctDeptPopUp(Byval arrRet)
	With frm1
		.txtDeptCd.focus
		.txtDeptCd.value = Trim(arrRet(0))
		.txtDeptNm.value = arrRet(1)
	End With
End Function

Function SetBizUnitPopUp(Byval arrRet)
	With frm1
		.vspdData.Col = C_BizUnitCd
		.vspdData.Text = arrRet(0)
		.vspdData.Col = C_BizUnitNm
		.vspdData.Text = arrRet(1)

	    Call vspdData_Change(.vspdData.Col, .vspdData.Row)				 ' ������ �о�ٰ� �˷��� 
	End With
End Function

Function SetCostCenterPopUp(Byval arrRet)
	With frm1
		.vspdData.Col = C_CostCd
		.vspdData.Text = arrRet(0)
		.vspdData.Col = C_CostNm
		.vspdData.Text = arrRet(1)

	    Call vspdData_Change(.vspdData.Col, .vspdData.Row)				 ' ������ �о�ٰ� �˷��� 
	End With
End Function


'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N")
    Call InitSpreadSheet
    Call InitVariables

	Call SetDefaultVal
	Call SetToolbar("1100000000001111")
    frm1.txtOrgChgID.focus
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
 			C_DeptCd		= iCurColumnPos(1)
			C_DeptNm		= iCurColumnPos(2)
			C_OrgChgDt		= iCurColumnPos(3)
	 		C_DeptFullNm	= iCurColumnPos(4)
	 		C_DeptEngNm		= iCurColumnPos(5)
	 		C_ParDeptCd		= iCurColumnPos(6)
	 		C_ParDeptNm		= iCurColumnPos(7)
	 		C_CostCd		= iCurColumnPos(8)
	 		C_CostPopUp		= iCurColumnPos(9)
	 		C_CostNm		= iCurColumnPos(10)
	 		C_BizUnitCd		= iCurColumnPos(11)
	 		C_BizUnitPopUp	= iCurColumnPos(12)
	 		C_BizUnitNm		= iCurColumnPos(13)
	 		C_Level			= iCurColumnPos(14)
	 		C_Seq			= iCurColumnPos(15)
	 		C_EndDeptFg		= iCurColumnPos(16)
	 		C_InternalCd	= iCurColumnPos(17)
	 		C_OrgChgId		= iCurColumnPos(18)
	 		C_OldInternalCd	= iCurColumnPos(19)
    End Select
End Sub


'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub



'+++++******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 
Sub vspdData_Click(ByVal Col, ByVal Row)

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call SetPopupMenuItemInf("0000111111")
	Else
		Call SetPopupMenuItemInf("0001111111")
	End If
	gMouseClickStatus = "SPC"

	Set gActiveSpdSheet = frm1.vspdData

	If Row = 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort
			lgSortKey = 2
		Else
			ggoSpread.SSSort ,lgSortKey
			lgSortKey = 1
		End If
	End If
End Sub



'========================================================================================================= 
Sub vspdData_DblClick(ByVal Col, ByVal Row)
    Dim iColumnName

	If Row<=0 then
		Exit Sub
	End If
	If Frm1.vspdData.MaxRows =0 then
		Exit Sub
	End if
End Sub



'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub


'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub



'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================= 
Sub vspdData_MouseDown(Button , Shift , x , y)

	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub


'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This event is spread sheet data changed
'==========================================================================================

Sub vspdData_Change(ByVal Col , ByVal Row )
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

    If Frm1.vspdData.CellType = Parent.SS_CELL_TYPE_FLOAT Then
      If CDbl(Frm1.vspdData.text) < CDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
    End If

	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
End Sub

'========================================================================================================= 
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	With frm1.vspdData 
		If Row >= NewRow Then
			Exit Sub
		End If
	End With
End Sub


'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 
'==========================================================================================

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)
	Dim strTemp
	Dim intPos1

	With frm1.vspdData 

		ggoSpread.Source = frm1.vspdData

		If Row > 0 And Col = C_BizUnitPOPUP Then
		    .Col = Col
		    .Row = Row
		    
			.Col = C_BizUnitCd
		    Call OpenBizUnitPopUP(.Text)
		End If

		If Row > 0 And Col = C_CostPOPUP Then
		    .Col = Col
		    .Row = Row

			.Col = C_CostCd
		    Call OpenCostPopUP(.Text)
		End If
		Call SetActiveCell(frm1.vspdData,Col-1,frm1.vspdData.ActiveRow ,"M","X","X")

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

     If CheckRunningBizProcess = True Then
	   Exit Sub
	End If

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
    	If lgPageNo <> "" Then
           Call DisableToolBar(Parent.TBC_QUERY)
           If DbQuery = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
    	End If
    End If
End Sub


'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery()
	Dim IntRetCD 

    FncQuery = False          '��: Processing is NG
    Err.Clear                 '��: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    ' ����� ������ �ִ��� Ȯ���Ѵ�.
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")			    '����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?
    	If IntRetCD = vbNo Then
	      	Exit Function
    	End If
    End If

    '-----------------------
    'Erase contents area
    '-----------------------
	' ���� Page�� Form Element���� Clear�Ѵ�. 
	' ClearField(pDoc, Optional ByVal pStrGrp)
	ggoSpread.Source= frm1.vspdData
	ggoSpread.ClearSpreadData
    Call InitVariables

    '-----------------------
    'Check condition area
    '-----------------------
	' Required�� ǥ�õ� Element���� �Է� [��/��]�� Check �Ѵ�.
    If frm1.txtDeptCD.value = "" Then
		frm1.txtDeptNM.value = ""
    End If

    If Not chkField(Document, "1") Then
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery
       
    FncQuery = True
    
End Function


'========================================================================================
Function FncNew() 
	On Error Resume Next
End Function


'========================================================================================
Function FncDelete() 
	On Error Resume Next
End Function


'========================================================================================
Function FncSave() 
	Dim IntRetCD 

    FncSave = False
    Err.Clear
    On Error Resume Next

    '-----------------------
    'Precheck area
    '-----------------------
    ' ����� ������ �ִ��� Ȯ���Ѵ�.
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False  Then  '��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","X","X","X")            '��: Display Message(There is no changed data.)
        Exit Function
    End If

    If Not chkField(Document, "1") Then               '��: Check required field(Single area)
       Exit Function
    End If

  '-----------------------
    'Check content area
    '----------------------- 
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                  '��: Check contents area
       Exit Function
    End If
    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave	
	FncSave = True 
End Function


'========================================================================================
Function FncCopy()
	On Error Resume Next            '��: Protect system from crashing
End Function


'========================================================================================
Function FncCancel() 
	If frm1.vspdData.MaxRows < 1 Then Exit Function
    ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo
End Function


'========================================================================================
Function FncInsertRow() 
	On Error Resume Next
End Function


'========================================================================================
Function FncDeleteRow()
	On Error Resume Next
End Function


'========================================================================================
Function FncPrint()
    parent.FncPrint()
End Function


'========================================================================================
Function FncPrev() 
    On Error Resume Next
End Function


'========================================================================================
Function FncNext() 
    On Error Resume Next
End Function


'========================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)
End Function


'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)
End Function

'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
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


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False

    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")                '����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    FncExit = True
End Function



'========================================================================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    Err.Clear                '��: Protect system from crashing

	Call LayerShowHide(1)

    With frm1
	    If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtDeptCd=" & Trim(.hDeptCd.value)	'��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtOrgChgID=" & Trim(.txtOrgChgId.value)	'��ȸ ���� ����Ÿ 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtDeptCd=" & Trim(.txtDeptCd.value)	'��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtOrgChgId=" & Trim(.txtOrgChgId.value)	'��ȸ ���� ����Ÿ 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If

		strVal = strVal & "&lgPageNo=" & lgPageNo

		Call RunMyBizASP(MyBizASP, strVal)		'��: �����Ͻ� ASP �� ���� 

    End With

    DbQuery = True

End Function

'========================================================================================
Function DbQueryOk()
    lgIntFlgMode = Parent.OPMD_UMODE
    Call ggoOper.LockField(Document, "Q")
	Call SetToolbar("1100100100011111")
	frm1.vspdData.focus
	Call SetSpreadLock()
    Set gActiveElement = document.ActiveElement
End Function


'========================================================================================
Function DbSave() 
	Dim lRow
	Dim lGrpCnt
	Dim strVal, strDel

    DbSave = False
    On Error Resume Next

	Call LayerShowHide(1)

	With frm1
		.txtMode.value = Parent.UID_M0002

		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
		strVal = ""
		strDel = ""

		'-----------------------
		'Data manipulate area
		'-----------------------
		For lRow = 1 To .vspdData.MaxRows

		    .vspdData.Row = lRow
		    .vspdData.Col = 0

		    Select Case .vspdData.Text
				Case ggoSpread.UpdateFlag							'��: ���� 
					strVal = strVal & "U" & Parent.gColSep & lRow & Parent.gColSep					'��: U=Update
				    .vspdData.Col = C_OrgChgID
				    strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
				    .vspdData.Col = C_DeptCd
				    strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
				    .vspdData.Col = C_BizUnitCd
				    strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
				    .vspdData.Col = C_CostCd
				    strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
				    lGrpCnt = lGrpCnt + 1
		    End Select

		Next

		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strDel & strVal

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)		'��: �����Ͻ� ASP �� ���� 

	End With
	DbSave = True 
End Function


'========================================================================================
Function DbSaveOk()
	ggoSpread.Source= frm1.vspdData
	ggoSpread.ClearSpreadData
	Call Dbquery
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
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
									<TD CLASS="TD5">��������ID</TD>
									<TD CLASS="TD6"><INPUT NAME="txtOrgChgID" MAXLENGTH="5"  SIZE=10 ALT ="��������ID" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOrgChangeID" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call HorgAbsPopUp(frm1.txtOrgChgID.Value)">
												    <INPUT NAME="txtOrgChgNm" MAXLENGTH="30" SIZE=30 ALT ="���������" tag="14X"></TD>
									<TD CLASS="TD5">�μ�</TD>
									<TD CLASS="TD6"><INPUT NAME="txtDeptCD" MAXLENGTH="10" SIZE=10 ALT ="�μ��ڵ�" tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCountryCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenAcctDeptPopUp(frm1.txtDeptCD.Value)">
													<INPUT NAME="txtDeptNM" MAXLENGTH="30" SIZE=30 ALT ="�μ���" tag="14X"></TD>
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
								<TD HEIGHT="100%" NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA class=hidden name=txtSpread tag="24"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode" tag="24">
<INPUT TYPE=hidden NAME="txtMaxRows" tag="24">
<INPUT TYPE=hidden NAME="hOrgChgId" tag="24">
<INPUT TYPE=hidden NAME="hDeptCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

