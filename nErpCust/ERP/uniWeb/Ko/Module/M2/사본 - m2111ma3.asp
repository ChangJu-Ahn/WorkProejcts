<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : M2111MA3
'*  4. Program Name         : ���ſ�ûȮ����� 
'*  5. Program Desc         : ���ſ�ûȮ����� 
'*  6. Component List       : 
'*  7. Modified date(First) : 1999/09/10
'*  8. Modified date(Last)  : 2003/05/21
'*  9. Modifier (First)     : Shin Jin Hyun
'* 10. Modifier (Last)      : KANG SU HWAN
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* !-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================!-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit		
Const BIZ_PGM_ID = "m2111mb3.asp"			
'==========================================  1.2.1 Global ��� ����  ======================================
Dim C_Check 
Dim C_Conflg
Dim C_Conflgstr
Dim C_ReqNo 
Dim C_PlantCd 
Dim C_PlantNm
Dim C_ItemCd 
Dim C_ItemNm 
Dim C_ItemSpec
Dim C_ReqQty 
Dim C_Unit 	
Dim C_DlvyDt 
Dim C_ReqDt 
Dim C_PrType
Dim C_ReqDeptCd
Dim C_ReqDeptNm
Dim C_ReqPrsn

'==========================================  1.2.2 Global ���� ����  =====================================
<!-- #Include file="../../inc/lgvariables.inc" -->

Dim bUpDataRowflg
Dim StartDate,EndDate

EndDate = "<%=GetSvrDate%>"
StartDate = UNIDateAdd("m", -1, EndDate, Parent.gServerDateFormat)
EndDate   = UniConvDateAToB(EndDate, Parent.gServerDateFormat, Parent.gDateFormat)
StartDate = UniConvDateAToB(StartDate, Parent.gServerDateFormat, Parent.gDateFormat)  

Dim IsOpenPop          

'==========================================   Selection()  ======================================
'	Name : Selection()
'	Description : �ϰ����ù�ư�� Event �ռ� 
'=========================================================================================================
Sub Selection(ByVal pFlag)
	Dim index,Count
	Dim strColValue
	
	frm1.vspdData.ReDraw = false
	
	Count = frm1.vspdData.MaxRows 
	If Trim(pFlag) = "ON" Then '�ϰ����� ��ư Ŭ���� 
		If frm1.rdoCfmflg(0).checked = true Then	'Ȯ���� ��ȸ�� ��� 
			For index = 1 to Count
				Call frm1.vspdData.SetText(C_Check,index,"1")
				Call frm1.vspdData.SetText(0,index,"")
			Next
		Else										'��Ȯ���� ��ȸ�� ��� 
			For index = 1 to Count
				Call frm1.vspdData.SetText(C_Check,index,"1")
				ggoSpread.UpdateRow Index
			Next
		End If 
	Else					'�ϰ�������� ��ư Ŭ���� 
		If frm1.rdoCfmflg(0).checked = true Then	'Ȯ���� ��ȸ�� ��� 
			For index = 1 to Count
				Call frm1.vspdData.SetText(C_Check,index,"0")
				ggoSpread.UpdateRow Index
			Next
		Else										'��Ȯ���� ��ȸ�� ��� 
			For index = 1 to Count
				Call frm1.vspdData.SetText(C_Check,index,"0")
				Call frm1.vspdData.SetText(0,index,"")
			Next
		End If 
	End If
	
	frm1.vspdData.ReDraw = true
	lgBlnFlgChgValue = true
End Sub

'==========================================================================================
'   Event Name : btnSelect_OnClick()
'   Event Desc : 
'==========================================================================================
Sub btnSelect_OnClick()
	If frm1.vspdData.Maxrows > 0 then
	    Call Selection("ON")
	End If
End Sub

'==========================================================================================
'   Event Name : btnDisSelect_OnClick()
'   Event Desc : 
'==========================================================================================
Sub btnDisSelect_OnClick()
	If frm1.vspdData.Maxrows > 0 then
	    Call Selection("OFF")
	End If
End Sub

'==========================================  2.1.1 InitVariables()  ======================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE 
    lgBlnFlgChgValue = False  
    lgIntGrpCount = 0         
    lgStrPrevKey = ""         
    lgLngCurRows = 0          
    frm1.vspdData.MaxRows = 0
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()
	frm1.rdoCfmflg(1).checked = true
	frm1.txtORGCd.Value = Parent.gPurOrg
	frm1.txtPlantCd.Value = Parent.gPlant
	Call SetToolbar("1110000000001111")
	 'frm1.btnAutoSel.disabled = True    
    frm1.txtOrgCd.focus 
	Set gActiveElement = document.activeElement
	
	frm1.txtFrReqDt.Text  = StartDate
	frm1.txtToReqDt.Text  = EndDate
	frm1.txtFrDlvyDt.Text = StartDate
	frm1.txtToDlvyDt.Text = EndDate
	
	frm1.btnSelect.disabled = True
	frm1.btnDisSelect.disabled = True
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
'========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub

'=============================================== 2.2.3 InitSpreadPosVariables() ========================================
Sub InitSpreadPosVariables()
	C_Check		= 1      
	C_Conflgstr = 2
	C_ReqNo		= 3
	C_PlantCd	= 4
	C_PlantNm	= 5
	C_ItemCd	= 6
	C_ItemNm	= 7
	C_ItemSpec	= 8
	C_ReqQty	= 9
	C_Unit		= 10
	C_DlvyDt	= 11
	C_ReqDt		= 12
	C_PrType	= 13
	C_ReqDeptCd = 14
	C_ReqDeptNm = 15
	C_ReqPrsn   = 16
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()
	
	With frm1.vspdData
	
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20030513",,Parent.gAllowDragDropSpread  

		.ReDraw = false

		.MaxCols = C_ReqPrsn+1												'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols:    .ColHidden = True
		.MaxRows = 0

		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetCheck C_Check		, "Ȯ������",10,,,true
		ggoSpread.SSSetCombo C_Conflgstr	, "Ȯ������", 10,0,False
		ggoSpread.SSSetEdit  C_ReqNo		, "��û��ȣ", 20
		ggoSpread.SSSetEdit  C_PlantCd		, "����",10,,,,2
		ggoSpread.SSSetEdit  C_PlantNm		, "�����",20
		ggoSpread.SSSetEdit  C_ItemCd		, "ǰ��", 18,,,,2
		ggoSpread.SSSetEdit  C_ItemNm		, "ǰ���", 20
		ggoSpread.SSSetEdit  C_ItemSpec		, "ǰ��԰�", 20
		SetSpreadFloatLocal	 C_ReqQty		, "��û��", 15, 1,3
		ggoSpread.SSSetEdit  C_Unit			, "����", 10,,,,2
		ggoSpread.SSSetDate  C_DlvyDt		, "�ʿ���", 10, 2, Parent.gDateFormat
		ggoSpread.SSSetDate  C_ReqDt		, "��û��", 10, 2, Parent.gDateFormat
		ggoSpread.SSSetEdit  C_PrType		, "PRTYPE", 20
		ggoSpread.SSSetEdit  C_ReqDeptCd	, "��û�μ�", 10
		ggoSpread.SSSetEdit  C_ReqDeptNm	, "��û�μ���", 20
		ggoSpread.SSSetEdit  C_ReqPrsn		, "��û��", 20
		
		ggoSpread.SetCombo "Y" & vbtab & "N",C_Conflgstr
    
		Call ggoSpread.SSSetColHidden(C_PrType,C_PrType,True)	
		Call ggoSpread.SSSetColHidden(C_Conflgstr,C_Conflgstr,True)	

		Call SetSpreadLock 
    
		.ReDraw = true
    End With
    
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
Sub SetSpreadLock()
    With frm1
	ggoSpread.spreadlock -1, -1
    ggoSpread.spreadUnlock C_Check, -1,C_Check, -1
    End With
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData
			
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_Check		= iCurColumnPos(1)     
			C_Conflgstr = iCurColumnPos(2)
			C_ReqNo		= iCurColumnPos(3)
			C_PlantCd	= iCurColumnPos(4)
			C_PlantNm	= iCurColumnPos(5)
			C_ItemCd	= iCurColumnPos(6)
			C_ItemNm	= iCurColumnPos(7)
			C_ItemSpec	= iCurColumnPos(8)
			C_ReqQty	= iCurColumnPos(9)
			C_Unit		= iCurColumnPos(10)
			C_DlvyDt	= iCurColumnPos(11)
			C_ReqDt		= iCurColumnPos(12)
			C_PrType	= iCurColumnPos(13)
			C_ReqDeptCd = iCurColumnPos(14)
			C_ReqDeptNm = iCurColumnPos(15)			
			C_ReqPrsn   = iCurColumnPos(16)			
	End Select

End Sub	

'------------------------------------------  OpenORG()  -------------------------------------------------
Function OpenORG()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "��������"					<%' �˾� ��Ī %>
	arrParam(1) = "B_Pur_Org"						<%' TABLE ��Ī %>
	
	arrParam(2) = Trim(frm1.txtORGCd.Value)     	<%' Code Condition%>
	
	arrParam(4) = ""							<%' Where Condition%>
	arrParam(5) = "��������"							<%' TextBox ��Ī %>
	
    arrField(0) = "PUR_ORG"					<%' Field��(0)%>
    arrField(1) = "PUR_ORG_NM"					<%' Field��(1)%>
    
    arrHeader(0) = "��������"						<%' Header��(0)%>
    arrHeader(1) = "����������"						<%' Header��(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtORGCd.focus
		Exit Function
	Else
		frm1.txtOrgCd.Value    = arrRet(0)		
		frm1.txtOrgNm.Value    = arrRet(1)		
		frm1.txtORGCd.focus
	End If	
End Function

'------------------------------------------  OpenItem()  -------------------------------------------------
Function OpenItem()
	Dim arrRet
	Dim arrParam(5), arrField(1)
	Dim iCalledAspName
	Dim IntRetCD

	If IsOpenPop = True Then Exit Function
		
	IsOpenPop = True

	if  Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "����", "X")
		IsOpenPop = False
		frm1.txtPlantCd.Focus
		Exit Function
	End if

	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)		' Item Code
	arrParam(2) = "!"	' ��12!MO"�� ���� -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- ��¥ 
	arrParam(5) = ""		'-- ����(b_item_by_plant a, b_item b: and ���� ����)
	
	arrField(0) = 1 ' -- ǰ���ڵ� 
	arrField(1) = 2 ' -- ǰ���				
    			
    iCalledAspName = AskPRAspName("B1B11PA3")
    
    If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		frm1.txtItemCd.Value    = arrRet(0)		
		frm1.txtItemNm.Value    = arrRet(1)		
		frm1.txtItemCd.focus
	End If	
End Function

'------------------------------------------  OpenPlant()  ------------------------------------------------
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "����"	
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
	
	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		frm1.txtPlantCd.Value= arrRet(0)		
		frm1.txtPlantNm.value= arrret(1)
	End If	
End Function

'==========================================================================================
'   Event Name : SetSpreadFloatLocal
'==========================================================================================
Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , ByVal dColWidth , ByVal HAlign , ByVal iFlag )
   Select Case iFlag
        Case 2                                                              '�ݾ� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign
        Case 3                                                              '���� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggQtyNo       ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"P"
        Case 4                                                              '�ܰ� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"P"
        Case 5                                                              'ȯ�� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"P"
    End Select
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()
    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call InitSpreadSheet                                                    '��: Setup the Spread sheet
    Call SetDefaultVal
    Call InitVariables                                                      '��: Initializes local global variables
End Sub

'==========================================================================================
'   Event Name : txtFrReqDt  	 
'==========================================================================================
Sub txtFrReqDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrReqDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtFrReqDt.Focus
	End if
End Sub
'==========================================================================================
'   Event Name : txtToReqDt  	 
'==========================================================================================
Sub txtToReqDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToReqDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtToReqDt.Focus
	End if
End Sub
'==========================================================================================
'   Event Name : txtFrDlvyDt
'==========================================================================================
Sub txtFrDlvyDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrDlvyDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtFrDlvyDt.Focus
	End if
End Sub
'==========================================================================================
'   Event Name : txtToDlvyDt
'==========================================================================================
Sub txtToDlvyDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToDlvyDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtToDlvyDt.Focus
	End if
End Sub
'==========================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc : 
'==========================================================================================
Sub txtFrReqDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtToReqDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtFrDlvyDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtToDlvyDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'==========================================================================================
'   Event Name : vspdData_Change
'==========================================================================================
'-- Modify for 9001 issue by Byun Jee Hyun 2004-11-30
'Sub vspdData_Change(ByVal Col , ByVal Row )
'    ggoSpread.Source = frm1.vspdData
'    
'	frm1.vspdData.Row = Row
'	frm1.vspdData.Col = 0
'	
'	If Col = C_Check And ggoSpread.UpdateFlag = frm1.vspdData.Text Then
'		ggoSpread.EditUndo
'	ElseIf Col = C_Check And ggoSpread.UpdateFlag <> frm1.vspdData.Text Then
'		ggoSpread.UpdateRow Row
'	ElseIf Col <> C_Check Then
'		ggoSpread.UpdateRow Row
'	End If
'		
'	Frm1.vspdData.Row = Row
'	Frm1.vspdData.Col = Col
'	
'    Call CheckMinNumSpread(frm1.vspdData, Col, Row)       
'End Sub

'========================================================================================
' Function Name : vspdData_Click
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
	   
	If frm1.vspdData.MaxRows > 0 Then
		Call SetPopupMenuItemInf("0001111111")
	Else
		Call SetPopupMenuItemInf("0000111111")
	End If   
	
	If frm1.vspdData.MaxRows = 0 Then Exit Sub
	   
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
	
	frm1.vspdData.Row = Row
	frm1.vspdData.Col = 0
	
	If Col = C_Check And ggoSpread.UpdateFlag = frm1.vspdData.Text Then
		ggoSpread.EditUndo
	ElseIf Col = C_Check And ggoSpread.UpdateFlag <> frm1.vspdData.Text Then
		ggoSpread.UpdateRow Row
	ElseIf Col <> C_Check Then
		ggoSpread.UpdateRow Row
	End If
		
	Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = Col
	
    Call CheckMinNumSpread(frm1.vspdData, Col, Row)       		
End Sub
'-- End of 9001 issue by Byun Jee Hyun 2004-11-30

'========================================================================================
' Function Name : vspdData_MouseDown
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    
'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then
		Exit Sub
	End If
	If frm1.vspddata.MaxRows=0 Then
		Exit Sub
	End If
End Sub
'========================================================================================
' Function Name : FncSplitColumn
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
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call ggoSpread.ReOrderingSpreadData()
End Sub


Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	
	With frm1.vspdData
	
		.Row = Row
    
		.Col = Col
		intIndex = .Value
		.Col = C_Conflgstr+1
		.Value = intIndex
    
		.Row = frm1.vspdData.ActiveRow
		.Col = C_Check
		.Text = "1"
		
    End With
    
    lgBlnFlgChgValue = True
    
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	'��: ������ üũ 
		If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
    
End Sub

'========================================================================================
' Function Name : FncQuery
'========================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing

	ggoSpread.Source = frm1.vspdData
	
    If ggoSpread.SSCheckChange = true Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    Call InitVariables
    															'��: Initializes local global variables
    
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If
        
	with frm1
		if (UniConvDateToYYYYMMDD(.txtFrDlvyDt.text,Parent.gDateFormat,"") > Parent.UniConvDateToYYYYMMDD(.txtToDlvyDt.text,Parent.gDateFormat,"")) and Trim(.txtFrDlvyDt.text)<>"" and Trim(.txtToDlvyDt.text)<>"" then	
			Call DisplayMsgBox("17a003", "X","�ʿ���", "X")			
			Exit Function
		End if   
        
		if (UniConvDateToYYYYMMDD(.txtFrReqDt.text,Parent.gDateFormat,"") > Parent.UniConvDateToYYYYMMDD(.txtToReqDt.text,Parent.gDateFormat,"")) and Trim(.txtFrReqDt.text)<>"" and Trim(.txtToReqDt.text)<>"" then	
			Call DisplayMsgBox("17a003", "X","���ſ�û��", "X")			
			Exit Function
		End if   

	End with
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function
       
	Set gActiveElement = document.activeElement
    FncQuery = True																'��: Processing is OK
End Function

'========================================================================================
' Function Name : FncNew
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          '��: Processing is NG
    Err.Clear                                                               '��: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData
    
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "A")                                         '��: Clear Condition Field
    Call ggoOper.LockField(Document, "N")                                          '��: Lock  Suitable  Field
    Call InitVariables                                                      '��: Initializes local global variables
    Call SetDefaultVal
           
    frm1.btnSelect.disabled = True
	frm1.btnDisSelect.disabled = True
	
	Set gActiveElement = document.activeElement
    FncNew = True                                                           '��: Processing is OK
End Function

'========================================================================================
' Function Name : FncSave
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         '��: Processing is NG
    Err.Clear                                                               '��: Protect system from crashing
    
    If CheckRunningBizProcess = True Then
		Exit Function
	End If
    
    ggoSpread.Source = frm1.vspdData
    
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData                          '��: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck         Then              '��: Check required field(Multi area)
       Exit Function
    End If
    
	If DbSave = False Then Exit Function
    
	Set gActiveElement = document.activeElement
    FncSave = True                                                          '��: Processing is OK
End Function

'========================================================================================
' Function Name : FncCancel
'========================================================================================
Function FncCancel() 
	if frm1.vspdData.Maxrows < 1	then exit function
	ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo   	        
	Set gActiveElement = document.activeElement
End Function
'========================================================================================
' Function Name : FncPrint
'========================================================================================
Function FncPrint() 
	ggoSpread.Source = frm1.vspdData
    Call parent.FncPrint()
	Set gActiveElement = document.activeElement
End Function

Function FncExcel()
	ggoSpread.Source = frm1.vspdData 
    Call parent.FncExport(Parent.C_SINGLE)												<%'��: ȭ�� ���� %>
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncFind
'========================================================================================
Function FncFind() 
	ggoSpread.Source = frm1.vspdData
    Call parent.FncFind(Parent.C_SINGLE , False)                                     <%'��:ȭ�� ����, Tab ���� %>
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncExit
'========================================================================================
Function FncExit()
	Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")                <%'����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
	Set gActiveElement = document.activeElement
    FncExit = True
End Function

'========================================================================================
' Function Name : DbQuery
'========================================================================================
Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim StrNextKey      
    
    DbQuery = False
    
    If LayerShowHide(1) = False Then Exit Function
    
    Err.Clear                                                               '��: Protect system from crashing

	Dim strVal
    
    With frm1
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then
	    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
	    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	    strVal = strVal & "&txtOrgCd=" & .hdnOrg.Value
	    strVal = strVal & "&txtPlantCd=" & .hdnPlant.value
	    strVal = strVal & "&txtitemCd=" & .hdnItem.value
		strVal = strVal & "&txtFrDlvyDt=" & .hdnFrDDt.Value
		strVal = strVal & "&txtToDlvyDt=" & .hdnToDDt.Value
		strVal = strVal & "&txtFrReqDt=" & .hdnFrRDt.Value
		strVal = strVal & "&txtToReqDt=" & .hdnToRDt.Value
		strVal = strVal & "&txtflg=" & .hdnflg.value
	    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	else
	    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
	    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	    strVal = strVal & "&txtOrgCd=" & Trim(.txtOrgCd.Value)
	    strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)
	    strVal = strVal & "&txtitemCd=" & Trim(.txtItemCd.value)
		strVal = strVal & "&txtFrDlvyDt=" & Trim(.txtFrDlvyDt.text)
		strVal = strVal & "&txtToDlvyDt=" & Trim(.txtToDlvyDt.text)
		strVal = strVal & "&txtFrReqDt=" & Trim(.txtFrReqDt.text)
		strVal = strVal & "&txtToReqDt=" & Trim(.txtToReqDt.text)
		if .rdoCfmflg(0).checked = true then
			strVal = strVal & "&txtFlg=" & "Y"
		elseif .rdoCfmflg(1).checked = true then
			strVal = strVal & "&txtFlg=" & "N"
		end if
	    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	end if 
    
	Call RunMyBizASP(MyBizASP, strVal)
        
    End With
    
    DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()														'��: ��ȸ ������ ������� 
    lgIntFlgMode = Parent.OPMD_UMODE												'��: Indicates that current mode is Update mode
    
    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
    
    Call SetToolbar("11101001000111")
	
	frm1.btnSelect.disabled = False
	frm1.btnDisSelect.disabled = False
	
	frm1.vspddata.focus
	Set gActiveElement = document.activeElement

End Function

Sub RemovedivTextArea()
	Dim ii

	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next
End Sub

'========================================================================================
' Function Name : DbSave
'========================================================================================
Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal
	Dim ColSep, RowSep
	
	Dim strCUTotalvalLen '���ۿ� ä������ ���� 102399byte�� �Ѿ� ���°��� üũ�ϱ����� ���� ����Ÿ ũ�� ����[����,�ű�] 
	Dim objTEXTAREA '������ HTML��ü(TEXTAREA)�� ��������� �ӽ� ���� 
	Dim iTmpCUBuffer         '������ ���� [����,�ű�] 
	Dim iTmpCUBufferCount    '������ ���� Position
	Dim iTmpCUBufferMaxCount '������ ���� Chunk Size
	Dim ii
	
    DbSave = False                                                          '��: Processing is NG
    
    If LayerShowHide(1) = False Then Exit Function
    
	ColSep = Parent.gColSep															'��: Column �и� �Ķ��Ÿ 
	RowSep = Parent.gRowSep															'��: Row �и� �Ķ��Ÿ 

	With frm1
		.txtMode.value = Parent.UID_M0002
    
    '-----------------------
    'Data manipulate area
    '-----------------------
    lGrpCnt = 1
	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '�ѹ��� ������ ������ ũ�� ����[����,�ű�]
	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '�ֱ� ������ ����[����,�ű�]
	
	iTmpCUBufferCount = -1
	strCUTotalvalLen = 0
    strVal = ""
    '-----------------------
    'Data manipulate area
    '-----------------------
    For lRow = 1 To .vspdData.MaxRows
    
        If Trim(GetSpreadText(.vspdData,0,lRow,"X","X")) = ggoSpread.UpdateFlag Then						'��: ���� 
			
			strVal = "U" & ColSep
			
            If Trim(GetSpreadText(.vspdData,C_Check,lRow,"X","X")) = "1" Then
            	strVal = strVal & "Y" & ColSep
            Else
            	strVal = strVal & "N" & ColSep
            End If

            strVal = strVal & Trim(GetSpreadText(.vspdData,C_ReqNo,lRow,"X","X")) & ColSep
			strVal = strVal & lRow & ColSep & Parent.gRowSep
			
            lGrpCnt = lGrpCnt + 1
        End if 

		Select Case Trim(GetSpreadText(.vspdData,0,lRow,"X","X"))
		    Case ggoSpread.UpdateFlag
		         If strCUTotalvalLen + Len(strVal) >  parent.C_FORM_LIMIT_BYTE Then  '�Ѱ��� form element�� ���� Data �Ѱ�ġ�� ������ 
			                            
		            Set objTEXTAREA = document.createElement("TEXTAREA")                 '�������� �Ѱ��� form element�� �������� ������ �װ��� ����Ÿ ���� 
		            objTEXTAREA.name = "txtCUSpread"
		            objTEXTAREA.value = Join(iTmpCUBuffer,"")
		            divTextArea.appendChild(objTEXTAREA)     
			 
		            iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT                  ' �ӽ� ���� ���� �ʱ�ȭ 
		            ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
		            iTmpCUBufferCount = -1
		            strCUTotalvalLen  = 0
		         End If
			       
		         iTmpCUBufferCount = iTmpCUBufferCount + 1
			      
		         If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                              '������ ���� ����ġ�� ������ 
		            iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT '���� ũ�� ���� 
		            ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
		         End If   
		         iTmpCUBuffer(iTmpCUBufferCount) =  strVal         
		         strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
		End Select   
                
    Next
		
	If iTmpCUBufferCount > -1 Then   ' ������ ������ ó�� 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If   
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										'��: �����Ͻ� ASP �� ���� 
	
	End With
	
    DbSave = True                                                           '��: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk()													'��: ���� ������ ���� ���� 
   
	Call InitVariables
	
	frm1.txtPlantCd.Value = frm1.hdnPlant.Value
	frm1.txtItemCd.Value = frm1.hdnItem.Value
	frm1.txtFrDlvyDt.text = frm1.hdnFrDDt.Value
	frm1.txtToDlvyDt.text = frm1.hdnToDDt.Value
	frm1.txtFrReqDt.text = frm1.hdnFrRDt.Value
	frm1.txtToReqDt.text = frm1.hdnToRDt.Value
	
	lgBlnFlgChgValue = False
	
	Call MainQuery()
	
End Function
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
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
					<TD CLASS="CLSLTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 border="0">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTAB"><font color=white>���ſ�ûȮ��/Ȯ�����</font></td>
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
								<TD CLASS="TD5" NOWRAP>��������</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtORGCd" ALT="��������" SIZE=10 MAXLENGTH=4  tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenORG()">
													   <INPUT TYPE=TEXT ID="txtORGNm" ALT="��������" NAME="arrCond" tag="14X"></TD>
								<TD CLASS="TD5" NOWRAP>Ȯ������</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio Class="Radio" ALT="Ȯ������" NAME="rdoCfmflg" id = "rdoCfmflg1" Value="Y" tag="1X"><label for="rdoCfmflg1">&nbsp;Ȯ��&nbsp;</label>
												 	   <INPUT TYPE=radio Class="Radio" ALT="Ȯ������" NAME="rdoCfmflg" id = "rdoCfmflg2" Value="N" checked tag="1X"><label for="rdoCfmflg2">&nbsp;��Ȯ��&nbsp;</label></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>����</TD>
								<TD CLASS="TD6" NOWRAP><INPUT ALT="����" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="1NNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">
													   <INPUT ALT="����" TYPE=TEXT NAME="txtPlantNm" SIZE=20 MAXLENGTH=20 tag="14X"></TD>
								<TD CLASS="TD5" NOWRAP>ǰ��</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="ǰ��" NAME="txtitemcd" SIZE=10 MAXLENGTH=18 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItem()">
													   <INPUT TYPE=TEXT ALT="ǰ��" NAME="txtitemNm" SIZE=20 tag="14X"></TD>
							</TR>
							
							<TR>
								<TD CLASS="TD5" NOWRAP>���ſ�û��</TD>
								<TD CLASS="TD6" NOWRAP>
									<table cellspacing=0 cellpadding=0>
										<tr>
											<td>
												<script language =javascript src='./js/m2111ma3_fpDateTime1_txtFrReqDt.js'></script>
											</td>
											<td>~</td>
											<td>
												<script language =javascript src='./js/m2111ma3_fpDateTime2_txtToReqDt.js'></script>
											</td>
										<tr>
									</table>
								</TD>
								<TD CLASS="TD5" NOWRAP>�ʿ���</TD>
								<TD CLASS="TD6" NOWRAP>
									<table cellspacing=0 cellpadding=0>
										<tr>
											<td>
												<script language =javascript src='./js/m2111ma3_fpDateTime3_txtFrDlvyDt.js'></script>
											</td>
											<td>~</td>
											<td>
												<script language =javascript src='./js/m2111ma3_fpDateTime2_txtToDlvyDt.js'></script>
											</td>
										<tr>
									</table>
								</TD>	   
							</TR>
						</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
			<TR>
				<TD WIDTH=100% valign=top>
					<TABLE <%=LR_SPACE_TYPE_20%>>
						<TR>
							<TD HEIGHT="100%">
								<script language =javascript src='./js/m2111ma3_I460027068_vspdData.js'></script>
							</TD>
						</TR>
					</TABLE>
				</TD>
			</TR>
		</TABLE></TD>
	</TR>
    <tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
    <tr HEIGHT="20">
      <td WIDTH="100%">
		<table <%=LR_SPACE_TYPE_30%>>
			<tr> 
				<TD WIDTH=10>&nbsp;</TD>
				<td WIDTH="*" align="left">
				<button name="btnSelect" class="clsmbtn" >�ϰ�����</button>&nbsp;
				<BUTTON NAME="btnDisSelect" CLASS="CLSMBTN">�ϰ��������</BUTTON>
				</td>
				<TD WIDTH=10>&nbsp;</TD>
			</tr>
		</table>
      </td>
    </tr>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<P ID="divTextArea"></P>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnOrg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPlant" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnItem" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnFrDDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnToDDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnFrRDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnToRDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnflg" tag="24">
<INPUT TYPE=HIDDEN NAME="txtAction" tag="24">
</FORM>

    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>

</BODY>
</HTML>
