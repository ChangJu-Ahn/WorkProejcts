<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID		    : A6101MA1
'*  4. Program Name         : �ΰ����հ�ǥ��ȸ 
'*  5. Program Desc         : �ΰ����հ�ǥ��ȸ 
'*  6. Component List       : +
'*  7. Modified date(First) : 2000/04/22
'*  8. Modified date(Last)  : 2001/03/17
'*  9. Modifier (First)     : Jong Hwan, Kim
'* 10. Modifier (Last)      : ������ 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc"  -->								<!-- '��: ȭ��ó��ASP���� �����۾��� �ʿ��� ���  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!-- '��: �ش� ��ġ�� ���� �޶���, ��� ���  -->

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">	</SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'��: indicates that All variables must be declared in advance

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_ID = "a6101mb1.asp"			'��: �����Ͻ� ���� ASP�� 

'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
'��: Grid Columns
'��ȸ������ ����ڹ�ȣ, �ŷ�ó��, ����, ���¼��� 
Dim  C_BPRgstNO      
Dim  C_BPNM          
Dim  C_IndClassNM    
Dim  C_IndTypeNM     
Dim  C_Cnt           
Dim  C_NetAmt        
Dim  C_VatAmt        
                     

 '==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->	

'Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
'Dim lgIntGrpCount              ' Group View Size�� ������ ���� 
'Dim lgIntFlgMode               ' Variable is for Operation Status

'Dim lgStrPrevKey
Dim lgStrPrevKeyBPNM

'Dim lgLngCurRows

Dim lgBlnStartFlag				' �޼��� �����Ͽ� ���α׷� ���۽��� Check Flag

 '==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 
 '----------------  ���� Global ������ ����  ----------------------------------------------------------- 
Dim  IsOpenPop
'Dim  lgSortKey

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
    C_BPRgstNO      = 1       
    C_BPNM          = 2           
    C_IndClassNM    = 3     
    C_IndTypeNM     = 4      
    C_Cnt           = 5            
    C_NetAmt        = 6         
    C_VatAmt        = 7         
End Sub
'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = 0                           'initializes Previous Key

    lgLngCurRows = 0                            'initializes Deleted Rows Count
	
	lgSortKey = 1
	
End Sub

'========================================================================================================= 

Sub SetDefaultVal()
	lgBlnStartFlag = False
	
	Dim strSvrDate
	DIm strYear, strMonth, strDay
	Dim frDt, toDt
	
	strSvrDate = "<%=GetSvrDate%>"
	Call ExtractDateFrom(strSvrDate, parent.gServerDateFormat, parent.gServerDateType, strYear,strMonth,strDay)
		
	frDt = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, "01")
	toDt = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)
	
	frm1.txtIssueDT1.Text = frDt
	frm1.txtIssueDT2.Text = toDt
	
	'frm1.txtBizAreaCD.value	= parent.gBizArea
	'frm1.txtIssueDT1.focus
	
End Sub
'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 


Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("Q", "A","NOCOOKIE","QA") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================

Sub InitSpreadSheet()
    Call initSpreadPosVariables()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021205",, parent.gAllowDragDropSpread
    With frm1.vspdData

        .MaxCols = C_VatAmt + 1
        .MaxRows = 0

        Call GetSpreadColumnPos("A")
        ggoSpread.SSSetEdit C_BPRgstNO, "����ڵ�Ϲ�ȣ", 14, , , 20
        ggoSpread.SSSetEdit C_BPNM,     "��ȣ", 20, , , 40

        ggoSpread.SSSetEdit C_IndClassNM, "����", 15, , , 20
        ggoSpread.SSSetEdit C_IndTypeNM, "����", 15, , , 20
        Call AppendNumberPlace("6","4","0")
        ggoSpread.SSSetFloat C_Cnt,    "�ż�",      7, "6",            ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec',2,,,"0","999"
        ggoSpread.SSSetFloat C_NetAmt, "���ް���", 17, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
        ggoSpread.SSSetFloat C_VatAmt, "����",     17, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec

 		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)

       .ReDraw = True

        Call SetSpreadLock                                              '�ٲ�κ� 

    End With
    
End Sub


'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadLock()
    With frm1.vspdData
        .ReDraw = False
		ggoSpread.SpreadLockWithOddEvenRowColor()
		ggoSpread.SSSetProtected	.MaxCols,-1,-1
        .ReDraw = True
    End With
End Sub


'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadColor(ByVal lRow)
End Sub


 '==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Function InitComboBox()


	 Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", _
                         " MAJOR_CD = " & FilterVar("A1003", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
     Call SetCombo2(frm1.cboIOFlag ,lgF0  ,lgF1  ,Chr(11))

End Function


'------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function
		
	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = ""							' ä�ǰ� ����(�ŷ�ó ����)
	arrParam(2) = ""								'FrDt
	arrParam(3) = ""								'ToDt
	arrParam(4) = "T"							'B :���� S: ���� T: ��ü 
	arrParam(5) = ""									'SUP :����ó PAYTO: ����ó SOL:�ֹ�ó PAYER :����ó INV:���ݰ�� 	
	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
			frm1.txtBPCd.focus
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
			lgBlnFlgChgValue = True
	End If
	

End Function
'------------------------------------------  OpenPopUp()  -------------------------------------------------
'	Name : OpenPopUp()
'	Description : PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopUp(Byval strCode, Byval iWhere)
Dim arrRet
Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
		Case 0
			
				arrParam(0) = "���ݽŰ����� �˾�"				' �˾� ��Ī 
			arrParam(1) = "B_TAX_BIZ_AREA"	 				' TABLE ��Ī 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "���ݽŰ������ڵ�"					' �����ʵ��� �� ��Ī 

			arrField(0) = "TAX_BIZ_AREA_CD"					' Field��(0)
			arrField(1) = "TAX_BIZ_AREA_NM"					' Field��(0)
    
			arrHeader(0) = "���ݽŰ������ڵ�"					' Header��(0)
			arrHeader(1) = "���ݽŰ������"
		Case 1
			arrParam(0) = "�ŷ�ó �˾�"					' �˾� ��Ī 
			arrParam(1) = "B_BIZ_PARTNER" 				' TABLE ��Ī 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "�ŷ�ó"						' �����ʵ��� �� ��Ī 

			arrField(0) = "BP_CD"						' Field��(0)
			arrField(1) = "BP_NM"						' Field��(1)
    
			arrHeader(0) = "�ŷ�ó�ڵ�"					' Header��(0)
			arrHeader(1) = "�ŷ�ó��"					' Header��(1)
	End Select
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
			Case 0		' ����� 
				frm1.txtBizAreaCD.focus
			Case 1		' ����� 
				frm1.txtBPCd.focus
		End Select
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function

'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : CtrlItem Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 

Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0		' ����� 
				.txtBizAreaCD.focus
				.txtBizAreaCD.value = UCase(Trim(arrRet(0)))
				.txtBizAreaNM.value = arrRet(1)				
			Case 1		' �ŷ�ó 
				.txtBPCd.focus
				.txtBPCd.value = UCase(Trim(arrRet(0)))
				.txtBPNM.value = arrRet(1)				
		End Select
	End With
End Function

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
            C_BPRgstNO      = iCurColumnPos(1)
            C_BPNM          = iCurColumnPos(2)
            C_IndClassNM    = iCurColumnPos(3)
            C_IndTypeNM     = iCurColumnPos(4)
            C_Cnt           = iCurColumnPos(5)
            C_NetAmt        = iCurColumnPos(6)
            C_VatAmt        = iCurColumnPos(7)
    End Select    
End Sub

'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N")
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)

    Call InitSpreadSheet
    Call InitVariables
    '----------  Coding part  -------------------------------------------------------------
	Call InitComboBox
	Call SetDefaultVal

    Call SetToolbar("1100000000001111")

     frm1.txtIssueDt1.focus
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

 '******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 

Sub InitData()
	Dim intRow
	Dim intIndex 
	
	For intRow = 1 To frm1.vspdData.MaxRows
		frm1.vspdData.Row = intRow
		' ���� 
		frm1.vspdData.Col = C_IndTypeCD
		intIndex = frm1.vspdData.value
		frm1.vspdData.col = C_IndTypeNM
		frm1.vspdData.value = intindex
		' ���� 
		frm1.vspdData.Col = C_IndClassCD
		intIndex = frm1.vspdData.value
		frm1.vspdData.col = C_IndClassNM
		frm1.vspdData.value = intindex
	Next
End Sub

Sub txtIssueDt1_Keypress(Key)
    If Key = 13 Then
		frm1.txtIssueDt2.focus
        FncQuery()
    End If
End Sub

Sub txtIssueDt2_Keypress(Key)
    If Key = 13 Then
		frm1.txtIssueDt1.focus
        FncQuery()
    End If
End Sub

'=======================================================================================================
'   Event Name : txtIssueDt1_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtIssueDt1_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssueDt1.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtIssueDt1.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtStartDt1_Change()
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtIssueDt1_Change()
    'lgBlnFlgChgValue = True
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
'   Event Name : txtIssueDt2_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtIssueDt2_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssueDt2.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtIssueDt2.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtStartDt2_Change()
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtIssueDt2_Change()
    'lgBlnFlgChgValue = True
End Sub



'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This event is spread sheet data changed
'==========================================================================================

Sub vspdData_Click(ByVal Col , ByVal Row)

    Call SetPopupMenuItemInf("0000111111")    
    gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData
    If frm1.vspdData.MaxRows <= 0 Then                                                    'If there is no data.
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
        Exit Sub
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub


'==========================================================================================
'   Event Desc : Spread Split �����ڵ� 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
    With frm1.vspdData 

    If Row >= NewRow Then
        Exit Sub
    End If

	 '----------  Coding part  -------------------------------------------------------------   

    End With

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


Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
     '----------  Coding part  -------------------------------------------------------------   
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKey <> 0 Then                         
			Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End if
    	End If
    End if
    
End Sub



'========================================================================================

Function FncQuery() 
	Dim IntRetCD 
    
    FncQuery = False          '��: Processing is NG
    Err.Clear                 '��: Protect system from crashing

    '-----------------------
    'Erase contents area
    '-----------------------
	' ���� Page�� Form Element���� Clear�Ѵ�. 
	' ClearField(pDoc, Optional ByVal pStrGrp)
    Call ggoOper.ClearField(Document, "2")      '��: Condition field clear    
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData

    'Call InitSpreadSheet                          '��: Setup the Spread Sheet
    Call InitVariables							'��: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
	' Required�� ǥ�õ� Element���� �Է� [��/��]�� Check �Ѵ�.
	' ChkField(pDoc, pStrGrp) As Boolean
    If Not chkField(Document, "1") Then	'��: This function check indispensable field
       Exit Function
    End If
    
	if frm1.txtBPCd.value = "" then
		frm1.txtBPNm.value = ""
	end if
	

    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery																'��: Query db data
       
    FncQuery = True																'��: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================

Function FncNew() 
End Function


'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncDelete() 
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave() 
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow() 
End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow() 
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
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
	Call Parent.FncPrint()
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
    Call Parent.FncExport(parent.C_MULTI)												'��: ȭ�� ���� 
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call FncFind(parent.C_MULTI, False)                                         '��:ȭ�� ����, Tab ���� 
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
    'Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	'Call InitData()
End Sub

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
Dim IntRetCD
	
	FncExit = False
	
	If lgBlnStartFlag = True Then
		' ����� ������ �ִ��� Ȯ���Ѵ�.
		If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")			'��: "Will you destory previous data"
	
			If IntRetCD = vbNo Then
				Exit Function
			End If
		End If
    End If
    
    FncExit = True
    
End Function

'========================================================================================
Function DbQuery() 
Dim strVal
Dim RetFlag

    DbQuery = False
    Err.Clear                '��: Protect system from crashing
    
    With frm1
		If Trim(.txtIssueDT1.text) = "" Then
	        RetFlag = DisplayMsgBox("A00001","x","x","x")   '�� �ٲ�κ� 
			'RetFlag = Msgbox("�������� ����ֽ��ϴ�!", vbOKOnly + vbInformation, "����")
			Exit Function
		End IF
		If Trim(.txtIssueDT2.text) = "" Then
	        RetFlag = DisplayMsgBox("A00001","x","x","x")   '�� �ٲ�κ� 
			'RetFlag = Msgbox("�������� ����ֽ��ϴ�!", vbOKOnly + vbInformation, "����")
			Exit Function
		End IF
    
		If UniConvDateToYYYYMMDD(.txtIssueDT1.text, parent.gDateFormat, "")  > UniConvDateToYYYYMMDD(.txtIssueDT2.text, parent.gDateFormat, "") Then
			RetFlag = DisplayMsgBox("113118", "X", "X", "X")			'��: "Will you destory previous data"
			Call SetToolbar("1100000000001111")										'��: ��ư ���� ���� 
			Exit Function
		End If
		Call LayerShowHide(1)
	
	    If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
			strVal = strVal & "&txtIssueDT1=" & Trim(.txtIssueDT1.Text)
			strVal = strVal & "&txtIssueDT2=" & Trim(.txtIssueDT2.Text)
			strVal = strVal & "&cboIOFlag=" & Trim(.hIOFlag.value)
			strVal = strVal & "&txtBizAreaCD=" & UCase(Trim(.hBizAreaCD.value))
			strVal = strVal & "&txtBPCd=" & UCase(Trim(.hBPCd.value))
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
			strVal = strVal & "&txtIssueDT1=" & Trim(.txtIssueDT1.Text)
			strVal = strVal & "&txtIssueDT2=" & Trim(.txtIssueDT2.Text)
			strVal = strVal & "&cboIOFlag=" & Trim(.cboIOFlag.value)
			strVal = strVal & "&txtBizAreaCD=" & UCase(Trim(.txtBizAreaCD.value))
			strVal = strVal & "&txtBPCd=" & UCase(Trim(.txtBPCd.value))
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If

		Call RunMyBizASP(MyBizASP, strVal)		'��: �����Ͻ� ASP �� ���� 
		    
    End With
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk()														'��: ��ȸ ������ ������� 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE	'��: Indicates that current mode is Update mode

	lgBlnFlgChgValue = False
	
	lgBlnStartFlag = True		' �޼��� �����Ͽ� ���α׷� ���۽��� Check Flag
	
    Call ggoOper.LockField(Document, "Q")	'��: This function lock the suitable field

    Call SetToolbar("1100000000011111")										'��: ��ư ���� ���� 
	frm1.vspdData.focus
	Set gActiveElement = document.ActiveElement

End Function


'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================

Function DbSave() 
End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================

Function DbSaveOk()													'��: ���� ������ ���� ���� 
End Function


'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================

Function DbDelete()
	On Error Resume Next
End Function

Sub txtBizAreaCD_onBlur()
	if frm1.txtBizAreaCD.value = "" then
		frm1.txtBizAreaNM.value = ""
	end if
End SUb
Sub txtBpCD_onBlur()
	if frm1.txtBPCd.value = "" then
		frm1.txtBPNM.value = ""
	end if
End SUb

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>><!-- ' ���� ���� --></TD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�ΰ����հ�ǥ��ȸ</font></td>
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
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5">������</TD>
									<TD CLASS="TD6"><script language =javascript src='./js/a6101ma1_fpDateTime2_txtIssueDt1.js'></script>&nbsp;~&nbsp;
													<script language =javascript src='./js/a6101ma1_fpDateTime2_txtIssueDt2.js'></script></TD>
									<TD CLASS="TD5">���ⱸ��</TD>
									<TD CLASS="TD6"><SELECT ID="cboIOFlag" NAME="cboIOFlag" ALT="���ⱸ��" STYLE="WIDTH: 98px" tag="12X"></SELECT>
									</TD>
								</TR>
								<TR>
									
									<TD CLASS="TD5">���ݽŰ�����</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtBizAreaCD" NAME="txtBizAreaCD" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" ALT="���ݽŰ�����" tag="11XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBizAreaCD.Value, 0)">&nbsp;
													<INPUT TYPE=TEXT ID="txtBizAreaNM" NAME="txtBizAreaNM" SIZE=20 MAXLENGTH=50 STYLE="TEXT-ALIGN: left" ALT="���ݽŰ�����" tag="14X" ></TD>
									<TD CLASS="TD5">�ŷ�ó</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtBPCd" NAME="txtBPCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" ALT="�ŷ�ó" tag="1XNXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBp(frm1.txtBPCd.Value, 1)">&nbsp;
													<INPUT TYPE=TEXT ID="txtBPNm" NAME="txtBPNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="�ŷ�ó" tag="14X" ></TD>									
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
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD HEIGHT="100%" COLSPAN=7>
								<script language =javascript src='./js/a6101ma1_vaSpread1_vspdData.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS="TD18"><FONT COLOR=Blue>�ֹε�Ϻ�  :</TD>
								<TD CLASS="TD18">�ż��հ�</TD>
								<TD WIDTH=10%><script language =javascript src='./js/a6101ma1_fpDoubleSingle1_txtCntPer.js'></script></TD>
								<TD CLASS="TD18">���ް����հ�</TD>
								<TD WIDTH=10%><script language =javascript src='./js/a6101ma1_fpDoubleSingle1_txtAmtPer.js'></script></TD>
								<TD CLASS="TD18">�����հ�</TD>
								<TD WIDTH=10%><script language =javascript src='./js/a6101ma1_fpDoubleSingle1_txtTaxPer.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS="TD18"><FONT COLOR=Blue>��ü�հ�  :</TD>
								<TD CLASS="TD18">�ż��հ�</TD>
								<TD WIDTH=10%><script language =javascript src='./js/a6101ma1_fpDoubleSingle1_txtCntSum.js'></script></TD>
								<TD CLASS="TD18">���ް����հ�</TD>
								<TD WIDTH=10%><script language =javascript src='./js/a6101ma1_fpDoubleSingle1_txtAmtSum.js'></script></TD>
								<TD CLASS="TD18">�����հ�</TD>
								<TD WIDTH=10%><script language =javascript src='./js/a6101ma1_fpDoubleSingle1_txtTaxSum.js'></script></TD>
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
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" tabindex="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" tabindex="-1">
<INPUT TYPE=HIDDEN NAME="hIssueDT1" tag="24" tabindex="-1">
<INPUT TYPE=HIDDEN NAME="hIssueDT2" tag="24" tabindex="-1">
<INPUT TYPE=HIDDEN NAME="hIOFlag" tag="24" tabindex="-1">
<INPUT TYPE=HIDDEN NAME="hBizAreaCD" tag="24" tabindex="-1">
<INPUT TYPE=HIDDEN NAME="hBPCd" tag="24" tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

