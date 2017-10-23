
Option Explicit	

Const BIZ_PGM_ID 		= "m2111qb1.asp"												'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_ID2 		= "m2111mb1_1.asp"												'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_JUMP_ID 	= "m2111ma1"
Const C_ReqNo 		= 1															'��: Spread Sheet�� Column�� ��� 
Const C_MaxKey       = 23            

Dim lgIsOpenPop          
Dim lgSaveRow     
Dim DBQueryCheck

Dim C_SpplCd
Dim C_SpplNm
Dim C_QuotaRate
Dim C_ApportionQty
Dim C_PlanDt
Dim C_GrpCd
Dim C_GrpNm
Dim lgPageNo2
Dim lgSpdHdrClicked

'========================================================================================
' Function Name : CookiePage
'========================================================================================
Sub WriteCookiePage()
	Dim strTemp, arrVal
	
	if frm1.vspdData.ActiveRow > 0 then
		frm1.vspdData.Row = frm1.vspdData.ActiveRow 
		frm1.vspdData.Col = GetKeyPos("A",C_ReqNo)
		Call WriteCookie("ReqNo" , frm1.vspdData.Text)
	end if 
End Sub

Sub ReadCookiePage()
	If Trim(ReadCookie("m2111ma1_plantcd")) = "" then Exit sub
	
	frm1.txtPlantCd.Value = ReadCookie("m2111ma1_plantcd")
	frm1.txtItemCd.Value = ReadCookie("m2111ma1_itemcd")
	
	Call WriteCookie("m2111ma1_plantcd","")
	Call WriteCookie("m2111ma1_itemcd","")
	
	Call MainQuery()
End Sub

'==========================================  2.1.1 InitVariables()  ======================================
Sub InitVariables()
	lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgSaveRow    = 0                      'initializes Group View Size
    lgStrPrevKey = ""                           'initializes Previous Key
    lgSortKey    = 1
    lgPageNo	 = ""
    lgPageNo2    = ""
    DBQueryCheck = True
End Sub
'=============================================== 2.2.3 InitSpreadSheet() ========================================
Sub InitSpreadSheet()
    Call SetZAdoSpreadSheet("M2111QA1","S","A","V20030510", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	Call SetSpreadLock("A") 
	
	Call InitSpreadSheet2
End Sub

Sub InitSpreadSheet2()
	Call InitSpreadPosVariables()

	With frm1
		.vspdData2.ReDraw = false

		ggoSpread.Source = frm1.vspdData2
        ggoSpread.Spreadinit "V20021103",, parent.gAllowDragDropSpread

	   .vspdData2.MaxCols = C_GrpNm+1
	   .vspdData2.MaxRows = 0

		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit   C_SpplCd, "����ó", 15,,,15,2
		ggoSpread.SSSetEdit	  C_SpplNm, "����ó��", 20
		SetSpreadFloatLocal	  C_QuotaRate, "��к���(%)",15,1,5
		SetSpreadFloatLocal   C_ApportionQty, "��η�", 15, 1,3
		ggoSpread.SSSetDate	  C_PlanDt, "���ֿ�����", 15,2,gDateFormat		
		ggoSpread.SSSetEdit	  C_GrpCd, "���ű׷�", 10,,,10,2
		ggoSpread.SSSetEdit   C_GrpNm, "���ű׷��", 20
				
		Call ggoSpread.MakePairsColumn(C_SpplCd,C_SpplNm)
		Call ggoSpread.MakePairsColumn(C_GrpCd,C_GrpNm)

		Call ggoSpread.SSSetColHidden(.vspdData2.MaxCols,	.vspdData2.MaxCols,	True)	
		
		.vspdData2.ReDraw = True
    End With

    Call SetSpreadLock("B")
End Sub

'================================== 2.2.4 SetSpreadLock() ===============================
Sub SetSpreadLock(ByVal pOpt)
    If pOpt = "A" Then
      ggoSpread.Source = frm1.vspdData
      ggoSpread.SpreadLockWithOddEvenRowColor()
	ElseIF pOpt = "B" Then
      ggoSpread.Source = frm1.vspdData2
      ggoSpread.SpreadLockWithOddEvenRowColor()
	End If
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
Sub InitSpreadPosVariables()
	C_SpplCd			=	1
	C_SpplNm			=	2
	C_QuotaRate			=	3
	C_ApportionQty		=	4
	C_PlanDt			=	5
	C_GrpCd				=	6
	C_GrpNm				=	7
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_SpplCd			=	iCurColumnPos(1)
			C_SpplNm			=	iCurColumnPos(2)
			C_QuotaRate			=	iCurColumnPos(3)
			C_ApportionQty		=	iCurColumnPos(4)	
			C_PlanDt			=	iCurColumnPos(5)
			C_GrpCd				=	iCurColumnPos(6)
			C_GrpNm				=	iCurColumnPos(7)
	End Select    
End Sub

Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , ByVal dColWidth , ByVal HAlign , ByVal iFlag )
   Select Case iFlag
        Case 2                                                              '�ݾ� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign
        Case 3                                                              '���� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggQtyNo       ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"P"
        Case 4                                                              '�ܰ� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"Z"
        Case 5                                                              'ȯ�� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"Z"
    End Select
End Sub

'------------------------------------  PopZAdoConfigGrid()  ----------------------------------------------
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If
	
	Call OpenGroupPopup("A")
End Sub

'========================================================================================================
' Function Name : PopZAdoConfigGrid
'========================================================================================================
Function OpenGroupPopup(ByVal pSpdNo)
	Dim arrRet
	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData(pSpdNo,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function

'------------------------------------------  OpenItem()  -------------------------------------------------
'Function OpenItem()
'	Dim arrRet
'	Dim arrParam(5), arrField(6), arrHeader(6)
'	Dim iCalledAspName
'	Dim IntRetCD
'	
'	If lgIsOpenPop = True Then Exit Function
'	
'	if  Trim(frm1.txtPlantCd.Value) = "" then
'		Call DisplayMsgBox("17A002", "X", "����", "X")
'		frm1.txtPlantCd.focus
'		Exit Function
'	End if
'
'	lgIsOpenPop = True
'
'	arrParam(0) = "ǰ��"
'	arrParam(1) = "B_Item_By_Plant,B_Plant,B_Item"	
'	
'	arrParam(2) = Trim(frm1.txtItemCd.Value)
'	
'	arrParam(4) = "B_Item_By_Plant.Plant_Cd = B_Plant.Plant_Cd And "
'	arrParam(4) = arrParam(4) & "B_Item_By_Plant.Item_Cd = B_Item.Item_Cd and B_Item.phantom_flg = " & FilterVar("N", "''", "S") & "  "
'	    
'	if Trim(frm1.txtPlantCd.Value)<>"" then
'		arrParam(4) = arrParam(4) & "And B_Plant.Plant_Cd= " & FilterVar(UCase(frm1.txtPlantCd.Value), "''", "S") & " " 
'	End if 
'	
'	arrParam(5) = "ǰ��"			
'   arrField(0) = "B_Item.Item_Cd"		
'    arrField(1) = "B_Item.Item_NM"	
'   arrField(2) = "B_Plant.Plant_Cd"
'  arrField(3) = "B_Plant.Plant_NM"
'    
'    arrHeader(0) = "ǰ��"		
'    arrHeader(1) = "ǰ���"		
'    arrHeader(2) = "����"		
'    arrHeader(3) = "�����"		
'    
'	iCalledAspName = AskPRAspName("M1111PA1")
'	
'	If Trim(iCalledAspName) = "" Then
'		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M1111PA1", "X")
'		lgIsOpenPop = False
'		Exit Function
'	End If
'	
'	arrRet = window.showModalDialog(iCalledAspName, Array(parent.window, arrParam, arrField, arrHeader), _
'		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
'		
'	lgIsOpenPop = False
'
'	If arrRet(0) = "" Then
'		frm1.txtItemCd.focus
'		Exit Function
'	Else
'		frm1.txtItemCd.Value    = arrRet(0)		
'		frm1.txtItemNm.Value    = arrRet(1)		
'		frm1.txtItemCd.focus
'	End If	
'End Function

'------------------------------------------  OpenItem()  -------------------------------------------------
Function OpenItem()
	Dim arrRet
	Dim arrParam(5), arrField(1)
	Dim iCalledAspName
	Dim IntRetCD
	
	If lgIsOpenPop = True Then Exit Function
	
	if  Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "����", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End if

	lgIsOpenPop = True
	
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
		lgIsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(parent.window, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		frm1.txtItemCd.Value    = arrRet(0)		
		frm1.txtItemNm.Value    = arrRet(1)		
		frm1.txtItemCd.focus
	End If	
End Function

'------------------------------------------  OpenPlant()  -------------------------------------------------
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

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
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		frm1.txtPlantCd.Value= arrRet(0)		
		frm1.txtPlantNm.value= arrRet(1)		
		frm1.txtPlantCd.focus
	End If	
	frm1.txtItemCd.value=""
	frm1.txtItemNm.value=""
End Function


Function OpenState()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "��û�������"
	arrParam(1) = "B_MINOR"			
	
	arrParam(2) = Trim(frm1.txtStateCd.Value)
	
	
	arrParam(4) = "Major_Cd=" & FilterVar("m2101", "''", "S") & ""	
	arrParam(5) = "��û�������"	
	
    arrField(0) = "Minor_cd"			
    arrField(1) = "Minor_Nm"
    
    arrHeader(0) = "��û�������"	
    arrHeader(1) = "��û������¸�"	
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtStateCd.focus
		Exit Function
	Else
		frm1.txtStateCd.Value = arrRet(0)
		frm1.txtStateNm.Value = arrRet(1)
		frm1.txtStateCd.focus
	End If	
End Function

'------------------------------------------  OpenDept()  -------------------------------------------------
Function OpenDept()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "��û�μ�"				
	arrParam(1) = "B_ACCT_DEPT"					
	
	arrParam(2) = Trim(frm1.txtDeptCd.Value)	
		
	arrParam(4) = "ORG_CHANGE_ID= " & FilterVar(Parent.gChangeOrgId, "''", "S") & " "
	arrParam(5) = "��û�μ�"							
	
    arrField(0) = "DEPT_CD"	
    arrField(1) = "DEPT_NM"
    
    arrHeader(0) = "��û�μ�"			
    arrHeader(1) = "��û�μ���"			
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtDeptCd.focus
		Exit Function
	Else
		frm1.txtDeptCd.Value = arrRet(0)
		frm1.txtDeptNm.Value = arrRet(1)
		frm1.txtDeptCd.focus
	End If	
End Function

'==========================================================================================
'   Event Name : OCX_EVENT
'==========================================================================================
 Sub txtDlvyFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtDlvyFrDt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtDlvyFrDt.Focus
	End if
End Sub

 Sub txtDlvyToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtDlvyToDt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtDlvyToDt.Focus
	End if
End Sub

Sub txtReqFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtReqFrDt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtReqFrDt.Focus
	End if
End Sub

Sub txtReqToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtReqToDt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtReqToDt.Focus
	End if
End Sub

'==========================================================================================
'   Event Name : OCX_KeyDown()
'==========================================================================================
Sub txtDlvyFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtDlvyToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtReqFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtReqToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
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
    Call InitSpreadSheet2()      
    Call ggoSpread.ReOrderingSpreadData()
End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
'========================================================================================
Sub vspdData2_MouseDown(Button , Shift , x , y)
   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'========================================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

Sub vspdData2_GotFocus()
    ggoSpread.Source = frm1.vspdData2
End Sub

'========================================================================================================
'   Event Name : vspdData2_ScriptDragDropBlock
'========================================================================================================
Sub vspdData2_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

Sub vspdData2_Click(ByVal Col, ByVal Row)	'###�׸��� ������ ���Ǻκ�###
 	gMouseClickStatus = "SPC"   
	 	 	
 	Set gActiveSpdSheet = frm1.vspdData2
 	
	Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
	
	If frm1.vspdData2.MaxRows = 0 Then
 		Exit Sub
 	End If
End Sub

'========================================================================================
' Function Name : vspdData_Click
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("00000000001")		
	gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
	   
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
	Else
 		lgSpdHdrClicked = 0		'2003-03-01 Release �߰� 
 		Call Sub_vspdData_ScriptLeaveCell(0, 0, Col, frm1.vspdData.ActiveRow, False)
	End If    

    Call SetSpreadColumnValue("A",Frm1.vspdData, Col, Row)  
End Sub

'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'=======================================================================================================
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
	If frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspddata2,NewTop) Then
    	If lgPageNo2 <> "" Then							
 			If CheckRunningBizProcess = True Then
				Exit Sub
			End If			
			Call DisableToolBar(parent.TBC_QUERY)
			If DBQuery2(0,False) = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
    	End If
    End If
End Sub

'=======================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'=======================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	If lgSpdHdrClicked = 1 Then	'2003-03-01 Release �߰� 
		Exit Sub
	End If
	
	Call Sub_vspdData_ScriptLeaveCell(Col, Row, NewCol, NewRow, Cancel)	
End Sub

'=======================================================================================================
'   Event Name : Sub_vspdData_ScriptLeaveCell
'=======================================================================================================
Sub Sub_vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)	
	Dim lRow
	if Row = 0 then exit sub
	If Row <> NewRow And NewRow > 0 Then
		frm1.vspdData2.MaxRows = 0
		Call Dbquery2(NewRow,True)
	End If
End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
	If y<20 Then			'2003-03-01 Release �߰� 
	    lgSpdHdrClicked = 1 
	End If

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
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

'==========================================================================================
'   Event Name : vspdData_DragDropBlock
'==========================================================================================
Sub vspdData_DragDropBlock(ByVal Col , ByVal Row , ByVal Col2 , ByVal Row2 , ByVal NewCol , ByVal NewRow , ByVal NewCol2 , ByVal NewRow2 , ByVal Overwrite , Action , DataOnly , Cancel )
    Row = 0: Row2 = -1: NewRow = 0
    ggoSpread.SwapRange Col, Row, Col2, Row2, NewCol, NewRow, Cancel
End Sub

'==========================================================================================
'   Event Name : vspdData_GotFocus
'==========================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspddata,NewTop) Then	'��: ������ üũ 
		If lgPageNo <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
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
	
	With frm1
	     If CompareDateByFormat(.txtDlvyFrDt.text,.txtDlvyToDt.text,.txtDlvyFrDt.Alt,.txtDlvyToDt.Alt, _
                   "970025",.txtDlvyFrDt.UserDefinedFormat,parent.gComDateType,False) = False And Trim(.txtDlvyFrDt.text) <> "" And Trim(.txtDlvyToDt.text) <> "" Then
			Call DisplayMsgBox("17a003", "X","�ʿ���", "X")
			Exit Function
		End if 
	
	    If CompareDateByFormat(.txtReqFrDt.text,.txtReqToDt.text,.txtReqFrDt.Alt,.txtReqToDt.Alt, _
                   "970025",.txtReqFrDt.UserDefinedFormat,parent.gComDateType,False) = False And Trim(.txtReqFrDt.text) <> "" And Trim(.txtReqToDt.text) <> "" Then
			Call DisplayMsgBox("17a003", "X","��û��", "X")			
			Exit Function
		End If     
	End With
	

	ggoSpread.Source = frm1.vspdData	'###�׸��� ������ ���Ǻκ�###
    ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2	'###�׸��� ������ ���Ǻκ�###
    ggoSpread.ClearSpreadData

    Call InitVariables    														'��: Initializes local global variables
    
    DBQueryCheck = True

    If Dbquery = False then Exit Function

    FncQuery = True																'��: Processing is OK
End Function

'========================================================================================
' Function Name : FncCancel
'========================================================================================
Function FncCancel() 
    ggoSpread.EditUndo                                                  '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncPrint
'========================================================================================
Function FncPrint() 
	ggoSpread.Source = frm1.vspdData
    Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_Multi)												
End Function

'========================================================================================
' Function Name : FncFind
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_Multi , False)                                     
End Function

'========================================================================================
' Function Name : FncExit
'========================================================================================
Function FncExit()
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
	Dim strVal
	
    DbQuery = False
    Err.Clear                                                               '��: Protect system from crashing
    
    If CheckRunningBizProcess = True Then
       Exit Function
    End If                                              
    
    Call LayerShowHide(1)

    With frm1
    
		If lgIntFlgMode = Parent.OPMD_UMODE Then
		    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		    strVal = strVal & "&txtPlantCd=" & .hdnPlant.value
		    strVal = strVal & "&txtItemCd=" & .hdnItem.value
		    strVal = strVal & "&txtStateCd=" & .hdnState.value
			strVal = strVal & "&txtDlvyFrDt=" & .hdnDFrDt.Value
			strVal = strVal & "&txtDlvyToDt=" & .hdnDToDt.Value
			strVal = strVal & "&txtReqFrDt=" & .hdnRFrDt.Value
			strVal = strVal & "&txtReqToDt=" & .hdnRToDt.Value
			strVal = strVal & "&txtDeptCd=" & .hdnDept.Value
		    strVal = strVal & "&txtchangorgid=" & parent.gchangeorgid
		Else
		    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		    strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantcd.value)
		    strVal = strVal & "&txtItemCd=" & Trim(.txtItemcd.value)
		    strVal = strVal & "&txtStateCd=" & Trim(.txtStateCd.value)
			strVal = strVal & "&txtDlvyFrDt=" & Trim(.txtDlvyFrDt.text)
			strVal = strVal & "&txtDlvyToDt=" & Trim(.txtDlvyToDt.text)
			strVal = strVal & "&txtReqFrDt=" & Trim(.txtReqFrDt.text)
			strVal = strVal & "&txtReqToDt=" & Trim(.txtReqToDt.text)
			strVal = strVal & "&txtDeptCd=" & Trim(.txtDeptCd.Value)
		    strVal = strVal & "&txtchangorgid=" & parent.gchangeorgid
	
		End If		
	        strVal = strVal & "&lgPageNo="   & lgPageNo         
		    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
			strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		    strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
			strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		
			Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
    
    End With
    
    DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()														'��: ��ȸ ������ ������� 

    lgIntFlgMode = Parent.OPMD_UMODE										'��: Indicates that current mode is Update mode
    'Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
    Call SetToolbar("1100000000011111")										'��: ��ư ���� ����	
    If DBQueryCheck = True Then
		Call DbQuery2(0,False)
	End If
    If frm1.vspdData.MaxRows > 0 Then
		frm1.vspddata.focus
	Else
		frm1.txtPlantCd.focus
	End If
	Set gActiveElement = document.activeElement

	DBQueryCheck = False
End Function

'========================================================================================
' Function Name : DbQuery
'========================================================================================
 Function DbQuery2(ByVal pRow,ByVal pFlag) 
	Dim strVal

    DbQuery2 = False
    
    Err.Clear                                               
	If LayerShowHide(1) = False Then Exit Function

    With frm1
    	If pFlag=True Then
			.vspdData.Row = pRow
		Else
			.vspdData.Row = .vspdData.ActiveRow
		End If
		.vspdData.Col = GetKeyPos("A",C_ReqNo)
		strVal = BIZ_PGM_ID2 & "?txtPrno=" & Trim(.vspdData.text)
	    strVal = strVal & "&lgPageNo="   & lgPageNo2         
	    strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
        
		Call RunMyBizASP(MyBizASP, strVal)								
    End With
    
    DbQuery2 = True
    Call SetToolbar("1100000000011111")									
End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk2()	

	lgBlnFlgChgValue = False
    lgSaveRow        = 1
    lgIntFlgMode = parent.OPMD_UMODE
End Function

